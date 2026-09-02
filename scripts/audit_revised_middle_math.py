#!/usr/bin/env python3
"""Strict release audit for the 371 revised middle-Math detail pages.

The audit is intentionally read-only.  It compares the working tree with a
caller-supplied pre-revision Git commit, validates the rendered pages against
the supplied center CSV and source workbook, and performs exhaustive
corpus-wide similarity checks that are practical for a single 371-page
category.

Reports should be written outside the repository so that running the audit
cannot make the release worktree dirty.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import math
import mimetypes
import random
import re
import statistics
import subprocess
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from functools import lru_cache
from itertools import combinations
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, unquote, urlsplit
from xml.etree import ElementTree as ET

from add_national_anchor_tocs import validate_page as validate_anchor_page
from audit_subject_pages_from_xlsx import (
    AUTHORING_RE,
    EXPECTED_GRAPH_SIGNATURES,
    JSONLD_RE,
    MIN_SOURCE_SENTENCE_CHARS,
    OVERCLAIM_PATTERNS,
    SOURCE_SHINGLE_WORDS,
    authored_parts,
    canonical_values,
    expected_image_url,
    first_fragment,
    graph_nodes,
    mask_document,
    meta_values,
    node_for,
    node_types,
    normalize,
    normalize_key,
    parse_attrs,
    recursive_schema_types,
    resolve_local,
    sentence_values,
    shingle_hashes,
    shingle_hits,
    single_tag_text,
    strip_visible,
    tag_attrs,
    word_values,
)
from generate_subject_pages_from_xlsx import (
    CONFIG_BY_SLUG,
    EXPECTED_ROWS,
    ORIGIN,
    PHONE,
    ROOT,
    SITE_NAME,
    TARGET_ROOT,
    TITLE_SUFFIX,
    absolute_route,
    display_region_label,
    load_centers,
    load_source_rows,
    relevant_grades,
    relevant_schools,
    source_text,
)


CATEGORY_SLUG = "중등수학학원"
EXPECTED_DETAILS = 371
EXPECTED_MANUSCRIPT_SECTIONS = 6
EXPECTED_FAQ_ITEMS = 5
EXPECTED_SCENARIOS = 2
EXPECTED_ANCHOR_TARGETS = 9

MASKED_SHINGLE_WORDS = 5
MASKED_MAX_LIMIT = 0.35
MASKED_P99_LIMIT = 0.25
MASKED_P95_LIMIT = 0.20

TEMPLATE_SHINGLE_WORDS = 8
TEMPLATE_COMMON_PAGE_RATIO = 0.10
TEMPLATE_MEDIAN_LIMIT = 0.40
TEMPLATE_P95_LIMIT = 0.50
TEMPLATE_MAX_LIMIT = 0.60

EXACT_PARAGRAPH_DF_LIMIT = 30
MASKED_PARAGRAPH_DF_LIMIT = 30
MASKED_SENTENCE_DF_LIMIT = 75
MASKED_META_DF_LIMIT = 18
MASKED_SENTENCE_MIN_CHARS = 45
HEADING_POSITION_UNIQUE_MIN = 75
HEADING_POSITION_DF_LIMIT = 37

AUTHORED_MIN_CHARS = 3300
AUTHORED_MAX_CHARS = 4800
SUMMARY_MIN_CHARS = 100
SUMMARY_MAX_CHARS = 700
SECTION_MIN_CHARS = 200
PARAGRAPH_MIN_CHARS = 35
PARAGRAPH_MAX_CHARS = 750
SENTENCE_MAX_CHARS = 180
META_MIN_CHARS = 70
META_MAX_CHARS = 160
FAQ_ANSWER_MIN_CHARS = 55
FAQ_ANSWER_MAX_CHARS = 300
HEADING_MAX_CHARS = 50
HEADING_MIDDLE_DOT_MAX = 3
KEYPHRASE_MAX_COUNT = 15

ALLOWED_NON_DETAIL_PATHS = {
    "sitemap.xml",
    "scripts/audit_revised_high_english.py",
    "scripts/audit_revised_high_math.py",
    "scripts/audit_revised_middle_english.py",
    "scripts/audit_revised_middle_math.py",
    "scripts/generate_subject_pages_from_xlsx.py",
}
PRIOR_DETAIL_PREFIXES = (
    "과목별학원/고등영어학원/",
    "과목별학원/고등수학학원/",
    "과목별학원/중등영어학원/",
)

# These workbook rows have a clearly wrong subject, level or assessment frame
# in the source H1.  The safe focus contract is intentionally independent from
# the generator so source drift cannot silently restore those fragments.
SOURCE_H1_CONTAMINATION_CONTRACT: dict[str, tuple[str, str]] = {
    "시흥동": (
        "내신과 모의고사를 함께 준비하는 방법",
        "학교 시험과 누적 유형 문제의 오답 원인을 나누는 방법",
    ),
    "목동": (
        "흔들리는 모의고사 점수부터 점검",
        "학교 시험과 누적 유형 문제의 풀이 시간을 점검하는 방법",
    ),
    "갈현동": (
        "내신과 모의고사를 함께 준비하는 방법",
        "학교 시험과 누적 유형 문제를 위한 수학 학습 점검",
    ),
    "송도": (
        "고등 수능까지 보는 학습전략",
        "중3 수학의 기초를 고등 과정 준비로 연결하는 방법",
    ),
    "죽전동": (
        "모의고사 등급을 점검하는 학습 안내",
        "학교 시험과 누적 유형 문제에서 조건 해석을 비교하는 방법",
    ),
    "옥산동": (
        "본문 변형문제까지 준비하는 학습관리",
        "교과서 유형 변형 문제의 조건을 해석하는 방법",
    ),
}

EXPECTED_BLANK_GRADE_LOCALITIES = {
    "진관동", "구파발", "갈현동", "다산동", "다산신도시", "부천 중동",
    "약대동", "고잔동", "초지동", "둔산동", "탄방동", "석사동", "퇴계동",
}
EXPECTED_BLANK_SCHOOL_LOCALITIES = {
    "염창동", "등촌동", "남가좌동", "북가좌동", "덕이동", "덕이지구",
    "주엽동", "대화동", "일산동", "후곡마을", "철산동", "하안동",
    "탄벌동", "경안동", "배곧", "배곧동", "정왕동", "고잔동", "초지동",
    "옥정동", "옥정신도시", "산내마을", "목동동", "동패동", "운정",
    "운정신도시", "야당동", "운정호수", "풍산동", "미사", "미사신도시",
    "반송동", "석우동", "봉담2지구", "봉담읍", "동춘동", "연수동",
    "관저동", "원내동", "둔산동", "탄방동", "불당동", "천안 백석동",
    "신불당", "울산 삼산동", "달동", "복산동", "약사동", "반구동",
    "서신동", "중화산동", "효자동", "송천동",
}

ENCODING_ARTIFACT_RE = re.compile(
    r"\ufffd|(?:Ã.|Â.|â(?:€|€™|€œ|€\x9d))|(?:\?{3,})",
    re.IGNORECASE,
)
REPEATED_WORD_RE = re.compile(r"(?<![0-9A-Za-z가-힣])([가-힣]{2,12})(?:\s+\1){1,}(?![0-9A-Za-z가-힣])")
REPEATED_WORD_WITH_PARTICLE_RE = re.compile(
    r"(?<![0-9A-Za-z가-힣])([가-힣]{2,12}?)(?:에는|에서|으로|에게|부터|까지|은|는|이|가|을|를|에|로|와|과|도|만)\s+\1(?![0-9A-Za-z가-힣])"
)
MIDDLE_DOT_LABEL_CHAIN_RE = re.compile(
    r"(?<![0-9A-Za-z가-힣])"
    r"(?:[0-9A-Za-z가-힣]{1,12}(?:\s+[0-9A-Za-z가-힣]{1,12})?)"
    r"(?:·[0-9A-Za-z가-힣]{1,12}(?:\s+[0-9A-Za-z가-힣]{1,12})?){2,}"
)
MIDDLE_MATH_LABELS = (
    "개념 연결",
    "계산 정확도",
    "조건 해석·식 세우기",
    "오답 원인·재풀이",
    "서술형 풀이",
    "함수·그래프·도형 해석",
    "선수 개념·단원 연결",
    "시험 시간 배분",
    "내신 범위·학교별 출제 유형",
    "고난도 다단계 문제",
    "수학 수행평가·탐구",
    "수학 학습 루틴",
    "시험 긴장·수학 자신감",
    "중1 첫 시험 적응",
    "고등 과정 전환",
)
EVIDENCE_CONTEXT_MISMATCH_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    (
        "score_sheet_before_marked_exam",
        re.compile(r"점수표보다[^.!?]{0,120}?(?:이|가)\s+남은\s+시험지를"),
    ),
    (
        "exam_sheet_followed_by_planning_evidence",
        re.compile(r"최근\s+시험지\s+한\s+장에서[^.!?]{0,140}?(?:주간\s+계획표|계획표|학습\s+기록|진단표|달력)"),
    ),
)
AWKWARD_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("duplicate_particle", re.compile(r"(?:문제집를|학원를|점검를|점검로|점검가|일정를|방식를|내용를|과정를|습관를|답안[’']?를|질문[’']?를|구간[’']?를|부분[’']?를|기[’']을)")),
    ("duplicate_noun", re.compile(r"(?:학생\s+학생|자료\s+자료|상담\s+상담|시기\s+시기|기준\s+기준|운영\s+운영|문제\s+문제|풀이\s+풀이|개념\s+개념|조건\s+조건)")),
    ("duplicate_compound", re.compile(r"(?<![가-힣])([가-힣]{2,12})·\1(?![가-힣])")),
    ("duplicate_suffix", re.compile(r"(?:서술·서술형|현재\s*수준에\s*맞춘.{0,20}현재\s*수준에\s*맞춘)")),
    ("broken_phrase", re.compile(r"(?:수\s*있은|묻은\s*질문|찾은\s*자리|맞은\s*순서|함께\s*읽은\s*것입니다|개념의\s*원인|공식의\s*원인|학습\s*루틴의\s*원인|계산\s*실수의\s*근거|평소에는\s+평소|준비\s+준비(?:까지|의)|점검할\s*학습\s+기록\s+점검|준비피드백|문장제\s+고민|진단과\s+학교\s+시험\s+준비|집중\s+점검학습|과제점검학습|수학피드백|오후학습|학습연습|학원학습|학습반복|학습클리닉|학원프로그램|준비종합점검|학습취약점|학습자립도|내신보충학습|학습진단|학습목표|학습기록|학습이력|대면학습|학습자율성|학습플래너|학습실전|화상학습|학습성취도|내신학습|선택\s+전\s+학원\s+변경\s+전|중등\s+수학학원의\s+중등\s+수학|최소\s+과제를\s+현재\s+시간표와\s+대조|참고\s+학교와\s+학습\s+우선순위\s+개설\s+범위|학습\s+우선순위와[^.!?]{0,24}의\s+우선순위를|학습\s+우선순위에\s+설명이\s+필요한지|시험\s+대비와\s+학습\s+습관을\s+함께(?=\s*[,.:]|</|$)|다음\s+학년\s+준비\s+점검과\s+학습\s+점검을\s+함께|시험\s+준비\s+점검까지\s+살피는\s+학습\s+점검|학습\s+기록\s+점검으로\s+공부\s+흐름\s+점검하기|(?:봄방학\s+수학\s+계획|학기\s+초\s+공부\s+습관|겨울방학\s+기초)를?\s+(?:세우는|다지는)(?=\s*[,.:]|</|$)|답은\s+성적표가\s+아니라\s+최근\s+시험지|다음\s+점검에서도\s+같은\s+시험지를\s+사용하되|시험지의\s+표시와\s+재풀이\s+결과를\s+각각\s+이어|센터\s+정보와[^.!?]{0,50}와[^.!?]{0,30}질문)")),
    ("routine_wrong_evidence", re.compile(r"복습·과제\s+루틴[^.!?]{0,260}새\s+문제\s+기록과\s+비교")),
    ("checklist_material_mismatch", re.compile(r"최근\s+시험지와[^.!?<>]{1,60}표시\s+위치")),
    ("checklist_deadline_mismatch", re.compile(r"학교\s+범위표와[^.!?<>]{1,60}마감일")),
    ("nested_exam_material", re.compile(r"최근\s+시험지와\s+교재에서\s+[‘\"][^’\"]*최근\s+시험지")),
    ("abstract_problem_number", re.compile(r"한\s+표\s+안에서도[^.!?]{1,80}의\s+문제\s+번호와\s+다음\s+행동을\s+다른\s+열에")),
    ("abstract_intent_operation", re.compile(r"오답\s+원인의\s+(?:실행\s+기록|실제\s+개설\s+여부|최소\s+행동)")),
    ("intent_as_course_name", re.compile(r"(?:학습\s+우선순위|오답\s+원인|복습·과제\s+루틴)\s+과정의\s+운영\s+여부")),
    ("wrong_method_particle", re.compile(r"(?:적용\s+방식와|운영\s+방식와|확인\s+방식와)")),
    ("duplicate_disclaimer", re.compile(r"이\s+표기만으로\s+이\s+표기만으로")),
    ("duplicate_learning_join", re.compile(r"학습\s*,\s*학습\s+우선순위")),
    ("duplicate_question_location_particle", re.compile(r"질문에서는\s+[가-힣 ]+에서는")),
    ("duplicate_check_verb", re.compile(r"[가-힣 ]*점검[’']?을\s+실제로\s+점검할\s+수\s+있는\s+순서")),
    ("priority_practice_deadline", re.compile(r"학습\s+우선순위\s+연습\s+완료일")),
    ("abstract_area_problem_number", re.compile(r"두\s+영역에\s+해당하는\s+문제\s+번호와\s+다음\s+행동")),
    ("answering_status_vs_record", re.compile(r"학생이\s+직접\s+답하는지를\s+새\s+문제\s+기록과\s+비교")),
    ("remaining_focus_phrase", re.compile(r"(?:공식만\s+암기하는|풀이\s+과정\s+과정|계산\s+실수를\s+실력으로|2등급을\s+위한\s+시험\s+대비|학습\s+기록\s+점검으로\s+완성하는\s+실력|개념의\s+방향을\s+바르게)")),
    ("unsupported_delivery_focus", re.compile(r"(?:화상\s*학습|대면\s*학습|매주\s+단원\s+테스트|개별지도\s+학습\s+안내)")),
    ("unspaced_learning_compound", re.compile(r"(?:학교시험|시험복습|시험분석|시험오답|시험전략|조건해석|식세우기|계산정확도|풀이과정|단원연결|목표점검|일정점검|내신진도점검|내신상담|문제풀이|오답노트|자기주도|내신시험)")),
    ("unsupported_outcome_focus", re.compile(r"(?:흔들림\s+없는\s+성적|꾸준한\s+성적을\s+만드는|성적\s+기복을\s+줄이는)")),
)
UNSUPPORTED_OPERATION_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("one_to_one", re.compile(r"(?:1\s*:\s*1|일대일|개인\s*과외)")),
    (
        "delivery_or_class_format",
        re.compile(
            r"(?:소수\s*정예|개별\s*지도|화상\s*수업|녹화\s*수업|온라인\s*수업|"
            r"대면\s*수업|오전\s*수업|주말\s*수업|그룹\s*수업|집중\s*수업|"
            r"정규반|집중반|시험\s*대비반|내신\s*특강|방학\s*특강|보강\s*수업)"
        ),
    ),
    ("parent_reporting", re.compile(r"학부모(?:와|에게).{0,24}(?:공유|보고|전달)")),
)
UNSUPPORTED_OUTCOME_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("score_improvement", re.compile(r"(?:성적|내신|등급)\s*(?:향상|상승)|성적을\s*(?:올리|바꾸)")),
    ("admission_result", re.compile(r"(?:합격|진학)\s*(?:보장|완성)|합격\s*전략")),
    ("result_guarantee", re.compile(r"(?:결과|성과|성적|실력).{0,16}보장")),
    ("rank_promise", re.compile(r"(?:최상위|상위권)\s*(?:보장|완성)")),
)
ALWAYS_FORBIDDEN_CONTAMINATION_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("english_subject", re.compile(r"(?:영어|독해|문법|어휘|구문\s*독해|영작|영어\s*듣기|영어\s*말하기)")),
    ("elementary_grade", re.compile(r"(?<![0-9A-Za-z가-힣])초[1-6](?![0-9A-Za-z가-힣])")),
    ("unsupported_high_grade", re.compile(r"(?<![0-9A-Za-z가-힣])고[2-3](?![0-9A-Za-z가-힣])")),
    ("csat_or_mock_exam", re.compile(r"(?:수능|모의고사|전국연합|학력평가|(?<![가-힣])모고(?![가-힣]))")),
)
READER_META_TERMS = (
    "검색 노출",
    "SEO",
    "AEO",
    "GEO",
    "키워드 밀도",
    "상위 노출",
    "이 원고",
    "이 글의 작성",
    "프롬프트",
    "생성형 AI",
)
UNSAFE_REVIEW_TERMS = (
    "실제 수강 후기",
    "실제 학부모 후기",
    "수강생 후기",
    "별점",
    "만족도 100",
)
ARTIFACT_TERMS = (
    "시험지",
    "시험 범위표",
    "범위표",
    "교재",
    "개념서",
    "유형서",
    "풀이 과정",
    "조건",
    "식",
    "오답",
    "모의고사",
    "학습 기록",
    "주간 계획표",
    "계획표",
    "진단표",
    "달력",
    "문제 번호",
    "답안",
    "단원표",
    "공식",
    "평가 조건표",
    "수행평가",
    "서술형",
)
ACTION_TERMS = (
    "표시",
    "비교",
    "분류",
    "설명",
    "기록",
    "다시 풀",
    "재확인",
    "밑줄",
    "질문",
    "구분",
    "검산",
    "유도",
    "대입",
)
FOCUS_EVIDENCE_TERMS = (
    "시험지",
    "교재",
    "범위표",
    "오답",
    "답안",
    "근거",
    "표시",
    "풀이",
    "기록",
    "진단표",
)
FOCUS_ACTION_TERMS = (
    "표시",
    "비교",
    "분류",
    "구분",
    "설명",
    "기록",
    "다시 풀",
    "밑줄",
    "질문",
    "적으",
    "배치",
)
FOCUS_CHECKPOINT_TERMS = (
    "재확인",
    "다음 점검",
    "확인일",
    "완료일",
    "일주일 뒤",
    "며칠 뒤",
    "다음 주",
    "변화",
    "다시 설명",
    "판단",
    "새 문제",
    "새 문장",
)
FOCUS_REPEAT_STOPWORDS = {
    "중등",
    "수학",
    "학원",
    "학생",
    "학교",
    "지역",
    "센터",
    "학년",
    "시험",
    "자료",
    "확인",
    "기준",
}


@dataclass
class Finding:
    severity: str
    code: str
    path: str
    message: str
    details: object | None = None

    def as_dict(self) -> dict[str, object]:
        result: dict[str, object] = {
            "severity": self.severity,
            "code": self.code,
            "path": self.path,
            "message": self.message,
        }
        if self.details is not None:
            result["details"] = self.details
        return result


@dataclass
class AuditState:
    findings: list[Finding] = field(default_factory=list)
    checks: Counter[str] = field(default_factory=Counter)

    def add(
        self,
        severity: str,
        code: str,
        path: str | Path,
        message: str,
        details: object | None = None,
    ) -> None:
        self.findings.append(Finding(severity, code, display_path(path), message, details))

    def error(self, code: str, path: str | Path, message: str, details: object | None = None) -> None:
        self.add("error", code, path, message, details)

    def warn(self, code: str, path: str | Path, message: str, details: object | None = None) -> None:
        self.add("warning", code, path, message, details)

    def check(
        self,
        code: str,
        condition: bool,
        path: str | Path,
        message: str,
        details: object | None = None,
        *,
        severity: str = "error",
    ) -> bool:
        self.checks[code] += 1
        if not condition:
            self.add(severity, code, path, message, details)
        return condition


@dataclass
class PageRecord:
    path: Path
    relative: str
    locality: str
    center: dict[str, object]
    title: str
    canonical: str
    authored_text: str
    masked_text: str
    paragraphs: list[str]
    masked_paragraphs: list[str]
    sections: list[str]
    headings: list[str]
    masked_headings: list[str]
    shingle5: set[int]
    shingle8: set[int]
    chars: int
    words: int
    keyphrase_count: int
    blank_grades: bool
    blank_schools: bool


def display_path(value: str | Path) -> str:
    if isinstance(value, Path):
        try:
            return value.resolve().relative_to(ROOT.resolve()).as_posix()
        except ValueError:
            return str(value)
    return str(value)


def percentile(values: Iterable[float | int], fraction: float) -> float:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        return 0.0
    position = min(len(ordered) - 1, max(0, round((len(ordered) - 1) * fraction)))
    return ordered[position]


def distribution(values: Iterable[float | int]) -> dict[str, float | int]:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        return {"count": 0, "min": 0, "median": 0, "p95": 0, "p99": 0, "max": 0}
    return {
        "count": len(ordered),
        "min": round(ordered[0], 6),
        "median": round(statistics.median(ordered), 6),
        "p95": round(percentile(ordered, 0.95), 6),
        "p99": round(percentile(ordered, 0.99), 6),
        "max": round(ordered[-1], 6),
    }


def run_git(*args: str, check: bool = True) -> bytes:
    result = subprocess.run(
        ["git", *args],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    if check and result.returncode:
        raise RuntimeError(result.stderr.decode("utf-8", errors="replace").strip())
    return result.stdout


def resolve_baseline(value: str) -> str:
    return run_git("rev-parse", "--verify", f"{value}^{{commit}}").decode("ascii").strip()


def baseline_blob(ref: str, path: Path) -> bytes:
    relative = path.relative_to(ROOT).as_posix()
    result = subprocess.run(
        ["git", "show", f"{ref}:{relative}"],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    if result.returncode:
        raise FileNotFoundError(f"{ref}:{relative}: {result.stderr.decode('utf-8', errors='replace').strip()}")
    return result.stdout


def changed_paths(ref: str) -> set[str]:
    tracked = run_git("diff", "--name-only", "-z", ref, "--").decode("utf-8").split("\0")
    untracked = run_git("ls-files", "--others", "--exclude-standard", "-z").decode("utf-8").split("\0")
    return {value for value in (*tracked, *untracked) if value}


def jsonld_nodes(raw: str, path: Path, state: AuditState) -> list[dict[str, object]]:
    scripts = JSONLD_RE.findall(raw)
    if not state.check("technical.jsonld_script_count", len(scripts) == 1, path, "expected exactly one JSON-LD script", len(scripts)):
        return []
    try:
        payload = json.loads(scripts[0])
    except json.JSONDecodeError as exc:
        state.error("technical.jsonld_parse", path, "invalid JSON-LD", str(exc))
        return []
    if not isinstance(payload, dict) or payload.get("@context") != "https://schema.org":
        state.error("technical.jsonld_root", path, "JSON-LD root/context is invalid")
        return []
    return graph_nodes(payload)


def visible_faq(raw: str) -> list[tuple[str, str]]:
    fragment = first_fragment(raw, r'<div class="faq-list"[^>]*>(.*?)</div>\s*</div>\s*</section>')
    return [
        (normalize(strip_visible(question)), normalize(strip_visible(answer)))
        for question, answer in re.findall(
            r"<details(?:\s+open)?[^>]*>\s*<summary>(.*?)</summary>\s*<p>(.*?)</p>\s*</details>",
            fragment,
            re.DOTALL | re.IGNORECASE,
        )
    ]


def schema_faq(nodes: list[dict[str, object]]) -> list[tuple[str, str]]:
    node = node_for(nodes, "FAQPage")
    if not node:
        return []
    result: list[tuple[str, str]] = []
    for item in node.get("mainEntity", []):
        if not isinstance(item, dict):
            continue
        answer = item.get("acceptedAnswer", {})
        result.append(
            (
                normalize(item.get("name", "")),
                normalize(answer.get("text", "") if isinstance(answer, dict) else ""),
            )
        )
    return result


def manuscript_sections(raw: str) -> list[tuple[str, list[str], str]]:
    manuscript = first_fragment(raw, r'<article class="site-shell manuscript-article"[^>]*>(.*?)</article>\s*</section>')
    result: list[tuple[str, list[str], str]] = []
    for fragment in re.findall(
        r'<section class="manuscript-section"[^>]*>(.*?)</section>',
        manuscript,
        re.DOTALL | re.IGNORECASE,
    ):
        heading = first_fragment(fragment, r"<h2\b[^>]*>(.*?)</h2>")
        paragraphs = [
            normalize(strip_visible(value))
            for value in re.findall(r"<p\b[^>]*>(.*?)</p>", fragment, re.DOTALL | re.IGNORECASE)
            if normalize(strip_visible(value))
        ]
        result.append((normalize(strip_visible(heading)), paragraphs, normalize(strip_visible(fragment))))
    return result


def punctuation_issues(value: str) -> list[str]:
    pairs = (("(", ")"), ("[", "]"), ("{", "}"), ("“", "”"), ("‘", "’"))
    issues = [f"{left}{right}:{value.count(left)}/{value.count(right)}" for left, right in pairs if value.count(left) != value.count(right)]
    if value.count('"') % 2:
        issues.append(f'double_quote:{value.count(chr(34))}')
    return issues


def near_repeated_focus_words(value: str, window: int = 4) -> list[str]:
    """Find a content word repeated within a short editorial-focus phrase."""

    words = re.findall(r"[가-힣]{2,12}", value)
    repeats: set[str] = set()
    for index, word in enumerate(words):
        if word in FOCUS_REPEAT_STOPWORDS:
            continue
        if word in words[index + 1:index + 1 + window]:
            repeats.add(word)
    return sorted(repeats)


def middle_dot_label_chains(value: str) -> list[str]:
    """Return two intent labels fused with a middle dot, not valid short lists."""

    issues: set[str] = set()
    for left in MIDDLE_MATH_LABELS:
        for right in MIDDLE_MATH_LABELS:
            if left == right:
                continue
            fused = f"{left}·{right}"
            if fused in value:
                issues.add(fused)
    for old_chain in ("중3·고등 전환·상담 진단·우선순위", "고등 과정 전환·중1 첫 시험 적응"):
        if old_chain in value:
            issues.add(old_chain)
    return sorted(issues)


def focus_workflow_sections(
    focus: str,
    sections: list[tuple[str, list[str], str]],
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    """Describe where the unique focus is tied to evidence, action and recheck."""

    linked: list[dict[str, object]] = []
    incomplete: list[dict[str, object]] = []
    for index, (heading, paragraphs, _section_text) in enumerate(sections, 1):
        paragraph_text = normalize(" ".join(paragraphs))
        if focus not in paragraph_text:
            continue
        signals = {
            "evidence": sorted(term for term in FOCUS_EVIDENCE_TERMS if term in paragraph_text),
            "action": sorted(term for term in FOCUS_ACTION_TERMS if term in paragraph_text),
            "checkpoint": sorted(term for term in FOCUS_CHECKPOINT_TERMS if term in paragraph_text),
        }
        item: dict[str, object] = {
            "section": index,
            "heading": heading,
            "signals": signals,
        }
        if all(signals.values()):
            linked.append(item)
        else:
            item["missing"] = sorted(name for name, values in signals.items() if not values)
            incomplete.append(item)
    return linked, incomplete


def mentioned_token(text: str, token: str) -> bool:
    particles = r"(?:에서|으로|부터|까지|처럼|보다|은|는|이|가|을|를|의|에|와|과|도|만)?"
    return bool(
        re.search(
            rf"(?<![가-힣A-Za-z0-9]){re.escape(token)}{particles}(?![가-힣A-Za-z0-9])",
            text,
        )
    )


def reader_meta_terms(text: str) -> list[str]:
    """Find reader-facing authoring terms without substring false positives."""

    return [
        term
        for term in READER_META_TERMS
        if re.search(
            rf"(?<![가-힣A-Za-z0-9]){re.escape(term)}(?![가-힣A-Za-z0-9])",
            text,
            re.IGNORECASE,
        )
    ]


def authoring_term_hits(text: str) -> list[str]:
    """Return production-language leaks, allowing genuine presentation scripts."""

    hits: set[str] = set()
    for match in AUTHORING_RE.finditer(text):
        value = match.group(0)
        if value.startswith("원고"):
            before = text[max(0, match.start() - 16):match.start()]
            after = text[match.end():min(len(text), match.end() + 24)]
            if re.search(r"(?:발표|말하기|수행평가|스피치)\s*$", before) or re.match(
                r"\s*(?:를\s*)?그대로\s*읽",
                after,
            ):
                continue
        hits.add(value)
    return sorted(hits)


def foreign_tokens(current: str, text: str, candidates: Iterable[str]) -> list[str]:
    current_compact = re.sub(r"\s+", "", current)
    result: list[str] = []
    for candidate in sorted(set(candidates) - {current}, key=len, reverse=True):
        candidate_compact = re.sub(r"\s+", "", candidate)
        if candidate_compact in current_compact or current_compact in candidate_compact:
            continue
        if mentioned_token(text, candidate):
            result.append(candidate)
    return result


def local_info_contract(raw: str) -> dict[str, object]:
    fragment = first_fragment(raw, r'<aside class="local-info-card"[^>]*>(.*?)</aside>')
    fields: dict[str, str] = {}
    field_fragments: dict[str, str] = {}
    for label, value in re.findall(r"<dt>(.*?)</dt>\s*<dd>(.*?)</dd>", fragment, re.DOTALL | re.IGNORECASE):
        key = normalize(strip_visible(label))
        fields[key] = normalize(strip_visible(value))
        field_fragments[key] = value
    grade_fragment = first_fragment(fragment, r'<ul class="grade-list"[^>]*>(.*?)</ul>')
    grades = [
        (normalize(strip_visible(subject)), normalize(strip_visible(values)))
        for subject, values in re.findall(
            r"<li>\s*<strong>(.*?)</strong>\s*<span>(.*?)</span>\s*</li>",
            grade_fragment,
            re.DOTALL | re.IGNORECASE,
        )
    ]
    schools: list[str] = []
    for label, value in field_fragments.items():
        if "학교 참고" in label:
            schools.extend(normalize(strip_visible(item)) for item in re.findall(r"<span>(.*?)</span>", value, re.DOTALL | re.IGNORECASE))
    return {
        "visible": normalize(strip_visible(fragment)),
        "fields": fields,
        "grades": grades,
        "schools": schools,
    }


def image_contract(raw: str) -> list[dict[str, str]]:
    return [
        {
            key: attrs.get(key, "")
            for key in ("src", "alt", "style", "loading", "decoding", "width", "height")
        }
        for attrs in tag_attrs(raw, "img")
    ]


SCHEMA_REGION_ALIASES = {
    "서울특별시": "서울", "서울": "서울",
    "경기도": "경기", "경기": "경기",
    "인천광역시": "인천", "인천": "인천",
    "부산광역시": "부산", "부산": "부산",
    "대구광역시": "대구", "대구": "대구",
    "대전광역시": "대전", "대전": "대전",
    "광주광역시": "광주", "광주": "광주",
    "울산광역시": "울산", "울산": "울산",
    "세종특별자치시": "세종", "세종": "세종",
    "강원특별자치도": "강원", "강원도": "강원", "강원": "강원",
    "충청북도": "충청", "충북": "충청", "충청남도": "충청", "충남": "충청", "충청": "충청",
    "전북특별자치도": "전라", "전라북도": "전라", "전북": "전라",
    "전라남도": "전라", "전남": "전라", "전라": "전라",
    "경상북도": "경상", "경북": "경상", "경상남도": "경상", "경남": "경상", "경상": "경상",
    "제주특별자치도": "제주", "제주": "제주",
}


def canonical_schema_region(value: object) -> str:
    text = normalize(str(value or ""))
    return SCHEMA_REGION_ALIASES.get(text, text)


def physical_schema_locality(address: object) -> str:
    parts = normalize(str(address or "")).split()
    if len(parts) < 2:
        return ""
    if parts[0].startswith("세종"):
        return "세종시"
    if parts[1].endswith("시"):
        return parts[1]
    return parts[1]


def normalize_contradictory_schema_address(node: dict[str, object]) -> None:
    types = set(node_types(node))
    if not {"EducationalOrganization", "LocalBusiness"} <= types:
        return
    address = node.get("address")
    if not isinstance(address, dict):
        return
    street = normalize(str(address.get("streetAddress", "")))
    parts = street.split()
    if not parts:
        return
    physical_region = parts[0]
    recorded_region = address.get("addressRegion", "")
    if canonical_schema_region(physical_region) == canonical_schema_region(recorded_region):
        return
    corrected = dict(address)
    corrected["addressRegion"] = physical_region
    locality = physical_schema_locality(street)
    if locality:
        corrected["addressLocality"] = locality
    node["address"] = corrected


def immutable_jsonld_contract(raw: str) -> list[dict[str, object]] | None:
    """Return schema fields that a prose-only revision may not change."""

    scripts = JSONLD_RE.findall(raw)
    if len(scripts) != 1:
        return None
    try:
        payload = json.loads(scripts[0])
    except json.JSONDecodeError:
        return None
    nodes = graph_nodes(payload) if isinstance(payload, dict) else []
    retained: list[dict[str, object]] = []
    for node in nodes:
        types = node_types(node)
        if "FAQPage" in types:
            continue
        current = json.loads(json.dumps(node, ensure_ascii=False))
        normalize_contradictory_schema_address(current)
        if "Article" in types:
            for key in (
                "description",
                "abstract",
                "about",
                "mentions",
                "hasPart",
                "dateModified",
                "keywords",
            ):
                current.pop(key, None)
        elif "WebPage" in types:
            for key in ("description", "about", "mentions", "hasPart"):
                current.pop(key, None)
        elif "Service" in types:
            current.pop("description", None)
        retained.append(current)
    return sorted(
        retained,
        key=lambda item: (
            "+".join(sorted(node_types(item))),
            str(item.get("@id", "")),
        ),
    )


def immutable_contract(raw: str) -> dict[str, object]:
    header = first_fragment(raw, r"(<header\b.*?</header>)")
    footer = first_fragment(raw, r"(<footer\b.*?</footer>)")
    breadcrumbs = first_fragment(raw, r'(<nav class="breadcrumbs".*?</nav>)')
    related = first_fragment(raw, r'(<section class="section local-links-section".*?</section>)')
    hrefs = [attrs.get("href", "") for attrs in tag_attrs(raw, "a")]
    stylesheets = [attrs.get("href", "") for attrs in tag_attrs(raw, "link") if "stylesheet" in attrs.get("rel", "").split()]
    scripts = [attrs.get("src", "") for attrs in tag_attrs(raw, "script") if attrs.get("src")]
    local_info = local_info_contract(raw)
    return {
        "header": normalize_key(header),
        "footer": normalize_key(footer),
        "breadcrumbs": normalize_key(breadcrumbs),
        # The explanatory verification note is authored prose and may change
        # during a content revision.  Preserve only the supplied facts and
        # their rendered grouping as the baseline contract.
        "local_info": {
            "fields": local_info["fields"],
            "grades": local_info["grades"],
            "schools": local_info["schools"],
        },
        "related": normalize_key(related),
        "hrefs": hrefs,
        "images": image_contract(raw),
        "stylesheets": stylesheets,
        "scripts": scripts,
        "jsonld_immutable": immutable_jsonld_contract(raw),
    }


@lru_cache(maxsize=None)
def image_is_decodable(path: Path) -> bool:
    try:
        from PIL import Image  # type: ignore[import-not-found]

        with Image.open(path) as image:
            image.verify()
        return True
    except ImportError:
        data = path.read_bytes()[:16]
        return (
            data.startswith(b"\xff\xd8\xff")
            or data.startswith(b"GIF8")
            or data.startswith(b"\x89PNG\r\n\x1a\n")
            or (data.startswith(b"RIFF") and data[8:12] == b"WEBP")
        )
    except Exception:
        return False


def validate_scope(
    baseline: str,
    expected_paths: set[Path],
    state: AuditState,
) -> dict[str, object]:
    changed = changed_paths(baseline)
    expected_relatives = {path.relative_to(ROOT).as_posix() for path in expected_paths}
    changed_details = changed & expected_relatives
    prior_detail_changes = {
        relative
        for relative in changed
        if relative.startswith(PRIOR_DETAIL_PREFIXES)
        and relative.endswith("/index.html")
        and relative.count("/") == 3
    }
    unexpected = sorted(changed - expected_relatives - prior_detail_changes - ALLOWED_NON_DETAIL_PATHS)
    state.check(
        "scope.changed_details",
        changed_details == expected_relatives,
        CATEGORY_SLUG,
        "all and only the 371 middle-Math detail pages must be changed",
        {
            "changed": len(changed_details),
            "missing": sorted(expected_relatives - changed_details)[:30],
        },
    )
    state.check("scope.unexpected_paths", not unexpected, ROOT, "files outside the explicit release allowlist changed", unexpected[:100])
    for path in expected_paths:
        try:
            baseline_blob(baseline, path)
        except FileNotFoundError as exc:
            state.error("scope.baseline_detail_missing", path, "detail page is absent from baseline", str(exc))
    return {
        "changed_paths": len(changed),
        "changed_details": len(changed_details),
        "unexpected_paths": unexpected,
        "preserved_prior_subject_details": len(prior_detail_changes),
        "allowed_non_detail_changes": sorted(changed & ALLOWED_NON_DETAIL_PATHS),
    }


def validate_sitemap(
    baseline: str,
    expected_urls: set[str],
    state: AuditState,
) -> dict[str, object]:
    path = ROOT / "sitemap.xml"

    def pairs(data: bytes) -> list[tuple[str, str]]:
        root = ET.fromstring(data.decode("utf-8"))
        result: list[tuple[str, str]] = []
        for node in root.findall("{*}url"):
            loc = node.find("{*}loc")
            lastmod = node.find("{*}lastmod")
            result.append(((loc.text or "").strip() if loc is not None else "", (lastmod.text or "").strip() if lastmod is not None else ""))
        return result

    try:
        before = pairs(baseline_blob(baseline, path))
        after = pairs(path.read_bytes())
    except (FileNotFoundError, ET.ParseError, UnicodeDecodeError) as exc:
        state.error("technical.sitemap_parse", path, "could not parse baseline/current sitemap", str(exc))
        return {}
    before_urls = [url for url, _ in before]
    after_urls = [url for url, _ in after]
    state.check("technical.sitemap_url_order", before_urls == after_urls, path, "sitemap URL scope/order changed")
    state.check("technical.sitemap_unique", len(after_urls) == len(set(after_urls)), path, "sitemap contains duplicate URLs")
    state.check("technical.sitemap_target_scope", expected_urls <= set(after_urls), path, "one or more middle-Math URLs are absent from sitemap", sorted(expected_urls - set(after_urls)))
    before_map = dict(before)
    after_map = dict(after)
    changed_lastmod = {url for url in set(before_map) & set(after_map) if before_map[url] != after_map[url]}
    prior_subject_urls = {
        url
        for url in set(after_map)
        if (
            len(parts := unquote(urlsplit(html.unescape(url)).path).strip("/").split("/")) == 3
            and parts[0] == "과목별학원"
            and parts[1] in {"고등영어학원", "고등수학학원", "중등영어학원"}
        )
    }
    allowed_changed_lastmod = expected_urls | prior_subject_urls
    state.check(
        "scope.sitemap_only_target_lastmod",
        changed_lastmod <= allowed_changed_lastmod,
        path,
        "a sitemap lastmod outside middle-Math and the three preserved subject categories changed",
        sorted(changed_lastmod - allowed_changed_lastmod)[:50],
    )
    state.check(
        "scope.sitemap_all_target_lastmod",
        expected_urls <= changed_lastmod,
        path,
        "all 371 revised middle-Math URLs must receive a fresh lastmod",
        {
            "changed": len(changed_lastmod),
            "missing": sorted(expected_urls - changed_lastmod)[:50],
        },
    )
    return {
        "urls": len(after_urls),
        "target_urls": len(expected_urls & set(after_urls)),
        "changed_lastmod": len(changed_lastmod),
        "preserved_prior_subject_lastmods": len(changed_lastmod & prior_subject_urls),
        "unexpected_lastmod_changes": len(changed_lastmod - allowed_changed_lastmod),
    }


def validate_baseline_contract(
    raw: str,
    baseline_raw: str,
    path: Path,
    state: AuditState,
) -> None:
    before = immutable_contract(baseline_raw)
    after = immutable_contract(raw)
    for field_name in before:
        state.check(
            f"facts.baseline_{field_name}",
            before[field_name] == after[field_name],
            path,
            f"immutable visible/design contract changed: {field_name}",
        )


def validate_meta_and_structure(
    raw: str,
    title: str,
    canonical: str,
    path: Path,
    state: AuditState,
) -> str:
    full_title = f"{title} | {TITLE_SUFFIX}"
    titles = single_tag_text(raw, "title")
    h1s = single_tag_text(raw, "h1")
    descriptions = meta_values(raw, name="description")
    state.check("seo.title", titles == [full_title], path, "title is not exact", titles)
    state.check("seo.h1", h1s == [title], path, "H1 is not exact", h1s)
    state.check("seo.canonical", canonical_values(raw) == [canonical], path, "canonical is not exact", canonical_values(raw))
    state.check("seo.og_url", meta_values(raw, prop="og:url") == [canonical], path, "og:url is not exact")
    state.check("seo.og_title", meta_values(raw, prop="og:title") == [full_title], path, "og:title is not exact")
    state.check("seo.description_count", len(descriptions) == 1, path, "meta description count is not one", descriptions)
    description = descriptions[0] if len(descriptions) == 1 else ""
    state.check(
        "seo.description_length",
        META_MIN_CHARS <= len(description) <= META_MAX_CHARS,
        path,
        "meta description length is outside the release range",
        len(description),
    )
    state.check("seo.description_og_parity", meta_values(raw, prop="og:description") == [description], path, "description and og:description differ")
    locality = title.removesuffix(" 중등 수학학원")
    state.check("seo.description_identity", locality in description and "중등 수학학원" in description, path, "description lacks locality/category identity")
    state.check("seo.lang", bool(re.search(r'<html\b[^>]*\blang=["\']ko["\']', raw, re.IGNORECASE)), path, "html lang is not ko")
    state.check("seo.viewport", bool(meta_values(raw, name="viewport")), path, "viewport metadata is missing")
    state.check("seo.indexable", not re.search(r'<meta\b[^>]*name=["\']robots["\'][^>]*content=["\'][^"\']*noindex', raw, re.IGNORECASE), path, "page contains noindex")
    return description


def validate_facts_and_schema(
    raw: str,
    nodes: list[dict[str, object]],
    center: dict[str, object],
    title: str,
    canonical: str,
    authored_text: str,
    all_localities: set[str],
    all_schools: set[str],
    all_centers: set[str],
    path: Path,
    state: AuditState,
) -> None:
    config = CONFIG_BY_SLUG[CATEGORY_SLUG]
    locality = str(center["locality"])
    expected_grades = relevant_grades(config, center, "수학")
    expected_schools = relevant_schools(config, center)
    info = local_info_contract(raw)
    fields = info["fields"]
    expected_region = normalize(display_region_label(center))
    state.check("facts.visible_region", fields.get("지역") == expected_region, path, "visible region differs from center CSV", {"actual": fields.get("지역"), "expected": expected_region})
    state.check("facts.visible_center", fields.get("센터 기준") == center["center_name"], path, "visible center differs from center CSV")
    state.check("facts.visible_address", fields.get("제공 주소") == center["address"], path, "visible address differs from center CSV")
    registration = normalize(str(center.get("registration", "")))
    office_name = normalize(str(center.get("office_name", "")))
    registration_text = normalize(fields.get("등록 정보", ""))
    state.check(
        "facts.visible_registration",
        (not registration and (not registration_text or "확인" in registration_text))
        or registration_text == registration,
        path,
        "visible registration differs from center CSV",
        {"actual": registration_text, "expected": registration},
    )
    visible_grade_map = dict(info["grades"])
    visible_grade_values = [value for value in re.split(r"[·,/\s]+", visible_grade_map.get("수학", "")) if re.fullmatch(r"중[1-3]", value)]
    state.check("facts.visible_grades", visible_grade_values == expected_grades, path, "visible Math grades differ from center CSV", {"actual": visible_grade_values, "expected": expected_grades})
    if not expected_grades:
        state.check("facts.blank_grade_label", "상담 시 확인" in normalize(str(info["visible"])), path, "blank grades are not labelled for consultation")
    state.check("facts.visible_schools", info["schools"] == expected_schools[:8], path, "visible schools differ from center CSV", {"actual": info["schools"], "expected": expected_schools[:8]})
    if not expected_schools:
        state.check("facts.blank_school_label", "상담 시 실제 학교 자료 확인" in normalize(str(info["visible"])), path, "blank schools lack the consultation label")
        invented_school_names = sorted(
            {
                match.group(0)
                for match in re.finditer(
                    r"(?<![0-9A-Za-z가-힣])(?:[가-힣]{2,20}중학교|[가-힣]{2,20}중)(?![0-9A-Za-z가-힣])",
                    authored_text,
                )
                if match.group(0) != "중학교" and match.group(0) not in all_schools
            }
        )
        state.check(
            "facts.blank_school_no_invented_names",
            not invented_school_names,
            path,
            "a blank-school CSV row contains an unsupported school name",
            invented_school_names,
        )

    schema_signatures = Counter(tuple(sorted(node_types(node))) for node in nodes)
    state.check("technical.jsonld_nodes", schema_signatures == EXPECTED_GRAPH_SIGNATURES, path, "JSON-LD is not the exact eight-node contract", {"actual": {"+".join(key): value for key, value in schema_signatures.items()}})
    schema_types = recursive_schema_types(nodes)
    state.check("claims.no_review_schema", not ({"Review", "AggregateRating", "Product"} & schema_types), path, "forbidden review/rating/product schema found", sorted({"Review", "AggregateRating", "Product"} & schema_types))
    organization = node_for(nodes, "EducationalOrganization")
    article = node_for(nodes, "Article")
    webpage = node_for(nodes, "WebPage")
    service = node_for(nodes, "Service")
    state.check("facts.organization_identity", bool(organization) and organization.get("name") == center["center_name"] and organization.get("telephone") == PHONE, path, "organization identity/phone differs from supplied facts")
    address = organization.get("address", {}) if organization else {}
    state.check("facts.organization_address", isinstance(address, dict) and address.get("streetAddress") == center["address"], path, "organization address differs from supplied facts")
    address_parts = normalize(str(center["address"])).split()
    expected_physical_region = address_parts[0] if address_parts else ""
    actual_address_region = address.get("addressRegion", "") if isinstance(address, dict) else ""
    actual_address_locality = normalize(str(address.get("addressLocality", ""))) if isinstance(address, dict) else ""
    state.check(
        "facts.organization_address_region",
        bool(expected_physical_region)
        and canonical_schema_region(actual_address_region) == canonical_schema_region(expected_physical_region),
        path,
        "organization addressRegion contradicts streetAddress",
        {"actual": actual_address_region, "streetAddress": center["address"]},
    )
    locality_tokens = actual_address_locality.split()
    normalized_street_address = normalize(str(center["address"]))
    locality_matches_street = (
        actual_address_locality == "세종시" and normalized_street_address.startswith("세종")
    ) or (
        bool(locality_tokens) and all(token in normalized_street_address for token in locality_tokens)
    )
    state.check(
        "facts.organization_address_locality",
        locality_matches_street,
        path,
        "organization addressLocality contradicts streetAddress",
        {"actual": actual_address_locality, "streetAddress": center["address"]},
    )
    identifier = organization.get("identifier", {}) if organization else {}
    identifier_name = normalize(identifier.get("name", "") if isinstance(identifier, dict) else "")
    identifier_value = normalize(identifier.get("value", "") if isinstance(identifier, dict) else "")
    state.check(
        "facts.organization_registration",
        (not registration)
        or (
            identifier_value == registration
            and (not office_name or identifier_name == office_name)
        ),
        path,
        "organization registration differs from supplied facts",
        {
            "actual": {"name": identifier_name, "value": identifier_value},
            "expected": {"name": office_name, "value": registration},
        },
    )
    state.check("technical.article_identity", bool(article) and article.get("headline") == title and article.get("url") == canonical, path, "Article identity differs from visible page")
    state.check("technical.webpage_identity", bool(webpage) and webpage.get("name") == title and webpage.get("url") == canonical, path, "WebPage identity differs from visible page")
    state.check("facts.article_grades", bool(article) and set(article.get("educationalLevel", [])) == set(expected_grades), path, "Article educationalLevel differs from center CSV", article.get("educationalLevel") if article else None)
    service_area = service.get("areaServed", {}) if service else {}
    state.check("facts.service_area", isinstance(service_area, dict) and service_area.get("name") == locality, path, "Service area differs from locality")
    schema_school_mentions = {
        normalize(item.get("name", ""))
        for item in (article.get("mentions", []) if article else [])
        if isinstance(item, dict) and item.get("@type") == "EducationalOrganization"
    }
    state.check("facts.schema_schools", schema_school_mentions == set(expected_schools), path, "Article school mentions differ from center CSV", {"actual": sorted(schema_school_mentions), "expected": sorted(expected_schools)})

    grade_mentions = set(re.findall(r"(?<![0-9A-Za-z가-힣])중[1-3](?![0-9A-Za-z가-힣])", authored_text))
    state.check("facts.authored_grades", grade_mentions <= set(expected_grades), path, "authored text contains an unsupported grade", sorted(grade_mentions - set(expected_grades)))
    supplied_center_context = f"{center['center_name']} {center['address']}"
    factual_address_localities = {
        candidate
        for candidate in all_localities
        if mentioned_token(supplied_center_context, candidate)
    }
    foreign_locality_hits = foreign_tokens(
        locality,
        authored_text,
        all_localities - factual_address_localities,
    )
    state.check("facts.foreign_locality", not foreign_locality_hits, path, "authored text contains another locality", foreign_locality_hits)
    foreign_center_hits = foreign_tokens(str(center["center_name"]), authored_text, all_centers)
    state.check("facts.foreign_center", not foreign_center_hits, path, "authored text contains another center", foreign_center_hits)
    foreign_school_hits = [
        school
        for school in sorted(all_schools - set(expected_schools), key=len, reverse=True)
        if not any(school in expected or expected in school for expected in expected_schools)
        and mentioned_token(authored_text, school)
    ]
    state.check("facts.foreign_school", not foreign_school_hits, path, "authored text contains a school absent from this CSV row", foreign_school_hits[:30])


def validate_faq_aeo(
    raw: str,
    nodes: list[dict[str, object]],
    locality: str,
    path: Path,
    state: AuditState,
) -> None:
    visible = visible_faq(raw)
    schema = schema_faq(nodes)
    state.check("aeo.faq_count", len(visible) == EXPECTED_FAQ_ITEMS, path, "visible FAQ count is not five", len(visible))
    state.check("technical.faq_schema_parity", visible == schema, path, "visible FAQ and FAQPage differ")
    questions = [question for question, _ in visible]
    state.check("aeo.faq_questions_unique", len(questions) == len(set(questions)), path, "FAQ questions are duplicated within the page")
    locality_questions = sum(locality in question for question in questions)
    state.check(
        "aeo.faq_locality_balance",
        locality_questions >= 2,
        path,
        "fewer than two FAQ questions identify the page locality",
        locality_questions,
    )
    for index, (question, answer) in enumerate(visible, 1):
        state.check(
            "aeo.faq_answer_length",
            FAQ_ANSWER_MIN_CHARS <= len(answer) <= FAQ_ANSWER_MAX_CHARS,
            path,
            f"FAQ {index} answer length is outside the release range",
            len(answer),
        )


def validate_schema_content_contract(
    nodes: list[dict[str, object]],
    title: str,
    canonical: str,
    description: str,
    headings: list[str],
    path: Path,
    state: AuditState,
) -> None:
    article = node_for(nodes, "Article")
    webpage = node_for(nodes, "WebPage")
    service = node_for(nodes, "Service")
    breadcrumb = node_for(nodes, "BreadcrumbList")
    state.check(
        "technical.description_schema_parity",
        bool(article)
        and bool(webpage)
        and bool(service)
        and article.get("description") == description
        and article.get("abstract") == description
        and webpage.get("description") == description
        and service.get("description") == description,
        path,
        "meta, Article, WebPage and Service descriptions differ",
    )
    for label, node in (("article", article), ("webpage", webpage)):
        parts = node.get("hasPart", []) if node else []
        part_names = [normalize(item.get("name", "")) for item in parts if isinstance(item, dict)]
        state.check(
            f"technical.{label}_haspart",
            part_names == headings,
            path,
            f"{label} hasPart does not exactly match the six visible manuscript H2 headings",
            {"actual": part_names, "expected": headings},
        )
        state.check(
            f"geo.{label}_about",
            bool(node) and isinstance(node.get("about"), list) and bool(node.get("about")),
            path,
            f"{label} about is missing",
        )
        state.check(
            f"geo.{label}_mentions",
            bool(node) and isinstance(node.get("mentions"), list) and bool(node.get("mentions")),
            path,
            f"{label} mentions is missing",
        )
    expected_breadcrumb = [
        (1, "홈", ORIGIN + "/"),
        (2, "과목별학원", ORIGIN + quote("/과목별학원/", safe="/%:@")),
        (3, "중등 수학학원", absolute_route(CATEGORY_SLUG)),
        (4, title, canonical),
    ]
    actual_breadcrumb: list[tuple[int, str, str]] = []
    if breadcrumb:
        for item in breadcrumb.get("itemListElement", []):
            if isinstance(item, dict):
                actual_breadcrumb.append(
                    (
                        int(item.get("position", 0) or 0),
                        normalize(item.get("name", "")),
                        str(item.get("item", "")),
                    )
                )
    state.check(
        "technical.schema_breadcrumb",
        actual_breadcrumb == expected_breadcrumb,
        path,
        "BreadcrumbList does not match the visible hierarchy",
        {"actual": actual_breadcrumb, "expected": expected_breadcrumb},
    )
    state.check(
        "technical.article_section",
        bool(article)
        and isinstance(article.get("articleSection"), list)
        and len(article.get("articleSection", [])) >= 4,
        path,
        "Article articleSection is incomplete",
    )
    service_offers = service.get("offers", []) if service else []
    state.check(
        "geo.service_offers",
        isinstance(service_offers, list) and bool(service_offers),
        path,
        "Service offers are missing",
    )


def validate_images(
    raw: str,
    nodes: list[dict[str, object]],
    center: dict[str, object],
    title: str,
    path: Path,
    representative_digests: list[str],
    state: AuditState,
) -> None:
    images = tag_attrs(raw, "img")
    representative = [attrs for attrs in images if "/assets/representative/" in attrs.get("src", "")]
    body = [attrs for attrs in images if "/assets/centers/common/" in attrs.get("src", "")]
    maps = [attrs for attrs in images if "/assets/maps/" in attrs.get("src", "")]
    state.check("technical.image_counts", len(representative) == len(body) == len(maps) == 1 and len(images) == 3, path, "page must contain exactly one representative/body/map image", {"all": len(images), "representative": len(representative), "body": len(body), "map": len(maps)})
    if len(representative) != 1 or len(body) != 1 or len(maps) != 1:
        return
    rep = representative[0]
    body_image = body[0]
    map_image = maps[0]
    state.check("technical.representative_hidden", "display:none" in rep.get("style", "").replace(" ", ""), path, "representative image is not hidden")
    state.check("technical.representative_alt", title in rep.get("alt", ""), path, "representative alt lacks title")
    state.check("technical.body_src", body_image.get("src") == f"../../../assets/centers/common/{center['body_image']}", path, "body image source differs from center mapping")
    state.check("technical.body_attributes", body_image.get("loading") == "lazy" and body_image.get("decoding") == "async" and title in body_image.get("alt", ""), path, "body image attributes/alt are incomplete")
    state.check("technical.map_src", map_image.get("src") == f"../../../assets/maps/{center['map_name']}", path, "map image source differs from center mapping")
    state.check("technical.map_attributes", map_image.get("loading") == "lazy" and map_image.get("decoding") == "async" and title in map_image.get("alt", "") and str(center["center_name"]) in map_image.get("alt", ""), path, "map image attributes/alt are incomplete")
    resolved: list[Path] = []
    representative_target: Path | None = None
    for label, attrs in (("representative", rep), ("body", body_image), ("map", map_image)):
        target = resolve_local(path, attrs.get("src", ""))
        state.check(f"technical.{label}_file", target is not None and target.is_file(), path, f"{label} image file is missing", str(target))
        if target and target.is_file():
            resolved.append(target)
            state.check(f"technical.{label}_decode", image_is_decodable(target), path, f"{label} image cannot be decoded")
            if label == "representative":
                representative_target = target
    if representative_target:
        representative_digests.append(hashlib.sha256(representative_target.read_bytes()).hexdigest())

    image_object = node_for(nodes, "ImageObject")
    article = node_for(nodes, "Article")
    rep_url = expected_image_url("/assets/representative/" + Path(rep["src"]).name)
    body_url = expected_image_url("/assets/centers/common/" + Path(body_image["src"]).name)
    map_url = expected_image_url("/assets/maps/" + Path(map_image["src"]).name)
    expected_mime = mimetypes.guess_type(str(representative_target or rep["src"]))[0] or ""
    expected_width: int | None = None
    expected_height: int | None = None
    if representative_target:
        try:
            from PIL import Image  # type: ignore[import-not-found]

            with Image.open(representative_target) as image:
                expected_mime = Image.MIME.get(image.format, expected_mime)
                expected_width, expected_height = image.width, image.height
        except (ImportError, OSError):
            pass
    expected_card = (
        "summary_large_image"
        if expected_width and expected_height and expected_width / expected_height >= 1.5
        else "summary"
    )
    state.check("technical.imageobject", bool(image_object) and image_object.get("url") == rep_url and image_object.get("contentUrl") == rep_url, path, "ImageObject differs from representative image")
    state.check("technical.og_image", meta_values(raw, prop="og:image") == [rep_url], path, "og:image differs from representative image")
    state.check("technical.og_image_secure", meta_values(raw, prop="og:image:secure_url") == [rep_url], path, "og:image:secure_url differs from representative image")
    state.check("technical.og_image_type", meta_values(raw, prop="og:image:type") == [expected_mime], path, "og:image:type differs from the representative file")
    state.check(
        "technical.og_image_dimensions",
        meta_values(raw, prop="og:image:width") == ([str(expected_width)] if expected_width else [])
        and meta_values(raw, prop="og:image:height") == ([str(expected_height)] if expected_height else []),
        path,
        "Open Graph image dimensions differ from the representative file",
    )
    state.check("technical.twitter_card", meta_values(raw, name="twitter:card") == [expected_card], path, "Twitter card type differs from the representative image ratio")
    state.check("technical.twitter_image", meta_values(raw, name="twitter:image") == [rep_url], path, "Twitter image differs from representative image")
    state.check("technical.article_images", bool(article) and article.get("image") == [rep_url, body_url, map_url], path, "Article image list differs from visible images")


def validate_links(raw: str, path: Path, state: AuditState) -> None:
    broken: list[str] = []
    for attrs in tag_attrs(raw, "a"):
        href = attrs.get("href", "")
        target = resolve_local(path, href)
        if target is not None and not target.is_file():
            broken.append(href)
    state.check("technical.internal_links", not broken, path, "one or more local links are broken", sorted(set(broken)))


def validate_natural_reader_value(
    raw: str,
    authored_text: str,
    paragraphs: list[str],
    sections: list[tuple[str, list[str], str]],
    locality: str,
    expected_grades: list[str],
    path: Path,
    state: AuditState,
) -> None:
    state.check("reader.authored_length", AUTHORED_MIN_CHARS <= len(authored_text) <= AUTHORED_MAX_CHARS, path, "authored visible text length is outside the release range", len(authored_text))
    summary_fragment = first_fragment(raw, r'<div class="local-summary"[^>]*>(.*?)</div>\s*<aside\b')
    summary_text = normalize(strip_visible(summary_fragment))
    state.check("aeo.summary_length", SUMMARY_MIN_CHARS <= len(summary_text) <= SUMMARY_MAX_CHARS, path, "answer-first summary length is outside the release range", len(summary_text))
    state.check("aeo.summary_identity", locality in summary_text and "중등 수학" in summary_text, path, "answer-first summary lacks locality/category identity")
    state.check("reader.section_count", len(sections) == EXPECTED_MANUSCRIPT_SECTIONS, path, "manuscript section count is not six", len(sections))
    for index, (heading, section_paragraphs, section_text) in enumerate(sections, 1):
        state.check("reader.section_heading", bool(heading), path, f"section {index} heading is empty")
        state.check(
            "reader.section_heading_length",
            len(heading) <= HEADING_MAX_CHARS,
            path,
            f"section {index} heading is too long",
            len(heading),
        )
        state.check(
            "reader.section_heading_dots",
            heading.count("·") <= HEADING_MIDDLE_DOT_MAX,
            path,
            f"section {index} heading contains too many middle-dot keyword joins",
            heading.count("·"),
        )
        state.check("reader.section_paragraphs", len(section_paragraphs) >= 2, path, f"section {index} has fewer than two paragraphs", len(section_paragraphs))
        state.check("reader.section_density", len(section_text) >= SECTION_MIN_CHARS, path, f"section {index} is too thin", len(section_text))

    intro_text = normalize(
        strip_visible(first_fragment(raw, r'<div class="manuscript-intro"[^>]*>(.*?)</div>'))
    )
    focus_match = re.search(r"핵심\s+주제는\s*[‘'\"](.+?)[’'\"]입니다", intro_text)
    focus = normalize(focus_match.group(1)) if focus_match else ""
    state.check(
        "reader.focus_extract",
        bool(focus),
        path,
        "the page-specific editorial focus could not be extracted from the manuscript introduction",
        intro_text[:240],
    )
    if focus:
        focus_repeats = near_repeated_focus_words(focus)
        state.check(
            "natural.focus_repeated_word",
            not focus_repeats,
            path,
            "the page-specific editorial focus repeats a nearby content word",
            {"focus": focus, "repeated": focus_repeats},
        )
        linked_focus_sections, incomplete_focus_sections = focus_workflow_sections(focus, sections)
        state.check(
            "reader.focus_workflow_sections",
            len(linked_focus_sections) >= 2,
            path,
            "the page-specific focus must appear in paragraph text, not only headings, and connect to evidence, action and checkpoint signals in at least two different manuscript sections",
            {
                "focus": focus,
                "linked_sections": linked_focus_sections,
                "incomplete_sections": incomplete_focus_sections,
                "required": 2,
            },
        )

    natural_issues: list[dict[str, object]] = []
    if ENCODING_ARTIFACT_RE.search(authored_text):
        natural_issues.append({"kind": "encoding_artifact"})
    authoring_terms = authoring_term_hits(authored_text)
    if authoring_terms:
        natural_issues.append({"kind": "authoring_terms", "values": authoring_terms})
    reader_terms = reader_meta_terms(authored_text)
    if reader_terms:
        natural_issues.append({"kind": "reader_meta_terms", "values": reader_terms})
    review_terms = [term for term in UNSAFE_REVIEW_TERMS if term in authored_text]
    if review_terms:
        natural_issues.append({"kind": "unsafe_review_terms", "values": review_terms})
    for label, pattern in AWKWARD_PATTERNS:
        hits = sorted(set(match.group(0) for match in pattern.finditer(authored_text)))
        if hits:
            natural_issues.append({"kind": label, "values": hits[:10]})
    particle_repeats = sorted(set(match.group(0) for match in REPEATED_WORD_WITH_PARTICLE_RE.finditer(authored_text)))
    if particle_repeats:
        natural_issues.append({"kind": "repeated_word_with_particle", "values": particle_repeats[:10]})
    label_chains = middle_dot_label_chains(authored_text)
    if label_chains:
        natural_issues.append({"kind": "middle_dot_label_chain", "values": label_chains[:20]})
    for label, pattern in EVIDENCE_CONTEXT_MISMATCH_PATTERNS:
        hits = sorted(set(match.group(0) for match in pattern.finditer(authored_text)))
        if hits:
            natural_issues.append({"kind": label, "values": hits[:10]})
    if len(sections) >= 4:
        facts_heading = sections[3][0]
        if "확인된 사실과 상담에서 물을 항목" in facts_heading:
            natural_issues.append({
                "kind": "fact_question_boundary_heading",
                "heading": facts_heading,
            })
        if re.search(r"적용(?:과|와).{0,30}적용\s*전", facts_heading):
            natural_issues.append({
                "kind": "repeated_application_heading",
                "heading": facts_heading,
            })
        if "학습 우선순위 학습" in facts_heading:
            natural_issues.append({
                "kind": "repeated_learning_heading",
                "heading": facts_heading,
            })
        if "중등 수학 전환 질문" in facts_heading:
            natural_issues.append({
                "kind": "reversed_transition_heading",
                "heading": facts_heading,
            })
        if "중학교 입학 준비 질문" in facts_heading:
            natural_issues.append({
                "kind": "diluted_entry_heading",
                "heading": facts_heading,
            })
    focus_compact = re.sub(r"\s+", "", focus)
    entry_context = "중1" in expected_grades and any(
        marker in focus_compact
        for marker in ("예비중", "중1", "중학교첫", "첫시험", "초등수학")
    )
    transition_context = "중3" in expected_grades and any(
        marker in focus_compact
        for marker in (
            "예비고1",
            "예비고등학생",
            "고등첫내신",
            "고등수학",
            "고등과정전환",
            "고등과정준비",
            "고등준비",
        )
    )
    if transition_context:
        if "중1 첫 시험 적응" in authored_text:
            natural_issues.append({
                "kind": "school_transition_intent_conflict",
                "focus": focus,
                "conflicting_label": "중1 첫 시험 적응",
            })
    if entry_context:
        if "고등 과정 전환" in authored_text:
            natural_issues.append({
                "kind": "school_transition_intent_conflict",
                "focus": focus,
                "conflicting_label": "고등 과정 전환",
            })

    contamination_hits: list[dict[str, object]] = []
    for label, pattern in ALWAYS_FORBIDDEN_CONTAMINATION_PATTERNS:
        hits = sorted(set(match.group(0) for match in pattern.finditer(authored_text)))
        if hits:
            contamination_hits.append({"kind": label, "values": hits[:20]})
    elementary_hits = sorted(
        set(
            match.group(0)
            for match in re.finditer(
                r"(?<![0-9A-Za-z가-힣])초등(?:학교|\s*수학|\s*과정|\s*단계)?(?![0-9A-Za-z가-힣])",
                authored_text,
            )
        )
    )
    if elementary_hits and not entry_context:
        contamination_hits.append({"kind": "elementary_context_outside_entry_intent", "values": elementary_hits[:20]})
    high_hits = sorted(
        set(
            match.group(0)
            for match in re.finditer(
                r"(?<![0-9A-Za-z가-힣])(?:예비고1|고1|고등(?:학교|\s*수학|\s*과정)?)(?![0-9A-Za-z가-힣])",
                authored_text,
            )
        )
    )
    if high_hits and not transition_context:
        contamination_hits.append({"kind": "high_school_context_outside_transition_intent", "values": high_hits[:20]})
    state.check(
        "natural.no_foreign_subject_level_assessment",
        not contamination_hits,
        path,
        "English, elementary, CSAT/mock-exam or unsupported high-school contamination was found",
        contamination_hits,
    )

    paragraph_seen: set[str] = set()
    sentence_seen: set[str] = set()
    for index, paragraph in enumerate(paragraphs, 1):
        key = normalize_key(paragraph)
        if key in paragraph_seen:
            natural_issues.append({"kind": "within_page_paragraph_duplicate", "paragraph": index})
        paragraph_seen.add(key)
        if not PARAGRAPH_MIN_CHARS <= len(paragraph) <= PARAGRAPH_MAX_CHARS:
            natural_issues.append({"kind": "paragraph_length", "paragraph": index, "length": len(paragraph)})
        if paragraph and paragraph[-1] not in ".?!다요죠”’":
            natural_issues.append({"kind": "paragraph_ending", "paragraph": index, "ending": paragraph[-12:]})
        punctuation = punctuation_issues(paragraph)
        if punctuation:
            natural_issues.append({"kind": "punctuation", "paragraph": index, "values": punctuation})
        repeated = REPEATED_WORD_RE.search(paragraph)
        if repeated:
            natural_issues.append({"kind": "repeated_word", "paragraph": index, "value": repeated.group(0)})
        for sentence in sentence_values(paragraph):
            sentence_key = normalize_key(sentence)
            if len(sentence_key) >= 35 and sentence_key in sentence_seen:
                natural_issues.append(
                    {
                        "kind": "within_page_sentence_duplicate",
                        "paragraph": index,
                        "sentence": sentence[:120],
                    }
                )
            sentence_seen.add(sentence_key)
            if len(sentence) > SENTENCE_MAX_CHARS:
                natural_issues.append({"kind": "sentence_too_long", "paragraph": index, "length": len(sentence)})
    state.check("natural.language_risks", not natural_issues, path, "natural-language risk detected", natural_issues[:50])

    claims: list[dict[str, str]] = []
    for label, pattern in OVERCLAIM_PATTERNS:
        for match in pattern.finditer(authored_text):
            context = authored_text[max(0, match.start() - 24):min(len(authored_text), match.end() + 32)]
            if re.search(r"보장(?:하지|되지|할\s*수\s*없)|보장하는.{0,12}(?:아니|아닌)", context):
                continue
            claims.append({"kind": label, "text": match.group(0)})
    state.check("claims.no_overclaim", not claims, path, "unsupported ranking/result claim found", claims)

    unsupported_operations: list[dict[str, object]] = []
    for label, pattern in UNSUPPORTED_OPERATION_PATTERNS:
        hits = sorted(set(match.group(0) for match in pattern.finditer(authored_text)))
        if hits:
            unsupported_operations.append({"kind": label, "values": hits[:20]})
    state.check(
        "claims.no_unsupported_operation",
        not unsupported_operations,
        path,
        "unsupported class format, delivery method or parent-reporting operation was claimed",
        unsupported_operations,
    )
    unsupported_outcomes: list[dict[str, object]] = []
    for label, pattern in UNSUPPORTED_OUTCOME_PATTERNS:
        hits = sorted(set(match.group(0) for match in pattern.finditer(authored_text)))
        if hits:
            unsupported_outcomes.append({"kind": label, "values": hits[:20]})
    state.check(
        "claims.no_unsupported_outcome",
        not unsupported_outcomes,
        path,
        "unsupported score, rank, admission or result outcome was claimed",
        unsupported_outcomes,
    )

    artifacts = sorted(term for term in ARTIFACT_TERMS if term in authored_text)
    actions = sorted(term for term in ACTION_TERMS if term in authored_text)
    state.check("reader.artifact_density", len(artifacts) >= 3, path, "fewer than three concrete learning artifacts are present", artifacts, severity="warning")
    state.check("reader.action_density", len(actions) >= 4, path, "fewer than four concrete learning actions are present", actions, severity="warning")
    scenarios = re.findall(r'<article class="scenario-card"[^>]*>.*?<p>(.*?)</p>\s*</article>', raw, re.DOTALL | re.IGNORECASE)
    state.check("reader.scenario_count", len(scenarios) == EXPECTED_SCENARIOS, path, "scenario count is not two", len(scenarios))
    scenario_context = normalize(strip_visible(first_fragment(raw, r'(<section class="section blue-wash".*?</section>)')))
    state.check("claims.scenario_disclosure", "가상 예시" in scenario_context and "실제 이용 후기" in scenario_context and "아닌" in scenario_context, path, "fictional consultation disclosure is incomplete")


def validate_source_reuse(
    record: PageRecord,
    source_sentences: set[str],
    source_shingles: set[int],
    state: AuditState,
) -> None:
    copied_sentences = sorted(
        {
            sentence
            for sentence in sentence_values(record.authored_text, MIN_SOURCE_SENTENCE_CHARS)
            if sentence in source_sentences
        }
    )
    copied_shingles = shingle_hits(record.authored_text, source_shingles, SOURCE_SHINGLE_WORDS, limit=8)
    state.check("source.exact_sentence", not copied_sentences, record.path, "a source sentence of at least 42 characters was reused", copied_sentences[:10])
    state.check("source.twelve_word_shingle", not copied_shingles, record.path, "a source twelve-word shingle was reused", copied_shingles)


def build_page_record(
    path: Path,
    center: dict[str, object],
    baseline: str,
    all_localities: set[str],
    all_schools: set[str],
    all_centers: set[str],
    source_sentences: set[str],
    source_shingles: set[int],
    representative_digests: list[str],
    descriptions: list[tuple[str, str]],
    state: AuditState,
) -> PageRecord | None:
    config = CONFIG_BY_SLUG[CATEGORY_SLUG]
    relative = path.relative_to(ROOT).as_posix()
    locality = str(center["locality"])
    title = f"{locality} {config.label}"
    canonical = absolute_route(config.slug, str(center["slug"]))
    try:
        raw_bytes = path.read_bytes()
        if raw_bytes.startswith(b"\xef\xbb\xbf"):
            state.error("technical.utf8_bom", path, "page contains a UTF-8 BOM")
        raw = raw_bytes.decode("utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        state.error("technical.page_read", path, "could not read UTF-8 page", str(exc))
        return None
    try:
        baseline_raw = baseline_blob(baseline, path).decode("utf-8")
    except (FileNotFoundError, UnicodeDecodeError) as exc:
        state.error("scope.baseline_read", path, "could not read baseline page", str(exc))
        return None

    validate_baseline_contract(raw, baseline_raw, path, state)
    description = validate_meta_and_structure(raw, title, canonical, path, state)
    descriptions.append((relative, description))
    nodes = jsonld_nodes(raw, path, state)
    authored_text, paragraphs, sections = authored_parts(raw)
    parsed_sections = manuscript_sections(raw)
    headings = [heading for heading, _, _ in parsed_sections]
    if locality in SOURCE_H1_CONTAMINATION_CONTRACT:
        contaminated, safe_focus = SOURCE_H1_CONTAMINATION_CONTRACT[locality]
        rendered_visible = normalize(strip_visible(raw))
        state.check(
            "source.polluted_h1_absent",
            contaminated not in rendered_visible,
            path,
            "a known wrong-subject source H1 fragment leaked into the rendered page",
            contaminated,
        )
        state.check(
            "source.polluted_h1_safe_focus",
            safe_focus in authored_text,
            path,
            "the audited middle-Math replacement focus for a contaminated source H1 is absent",
            safe_focus,
        )
    validate_facts_and_schema(raw, nodes, center, title, canonical, authored_text, all_localities, all_schools, all_centers, path, state)
    validate_schema_content_contract(nodes, title, canonical, description, headings, path, state)
    validate_faq_aeo(raw, nodes, locality, path, state)
    validate_images(raw, nodes, center, title, path, representative_digests, state)
    validate_links(raw, path, state)
    expected_grades = relevant_grades(config, center, "수학")
    validate_natural_reader_value(
        raw,
        authored_text,
        paragraphs,
        parsed_sections,
        locality,
        expected_grades,
        path,
        state,
    )
    anchor_errors = validate_anchor_page(raw)
    state.check("technical.anchor_toc", not anchor_errors, path, "anchor TOC validation failed", anchor_errors)
    state.check("technical.anchor_target_count", len(re.findall(r'class=["\'][^"\']*\bnational-page-anchor-target\b', raw)) == EXPECTED_ANCHOR_TARGETS, path, "anchor target count is not nine")
    keyphrase = f"{locality} 중등 수학학원"
    keyphrase_count = authored_text.count(keyphrase)
    state.check(
        "seo.keyphrase_count",
        keyphrase_count <= KEYPHRASE_MAX_COUNT,
        path,
        "exact local keyphrase is excessively repeated",
        keyphrase_count,
        severity="warning",
    )
    masked_paragraphs = [mask_document(value, config, center) for value in paragraphs]
    record = PageRecord(
        path=path,
        relative=relative,
        locality=locality,
        center=center,
        title=title,
        canonical=canonical,
        authored_text=authored_text,
        masked_text=mask_document(authored_text, config, center),
        paragraphs=paragraphs,
        masked_paragraphs=masked_paragraphs,
        sections=sections,
        headings=headings,
        masked_headings=[mask_document(value, config, center) for value in headings],
        shingle5=shingle_hashes(mask_document(authored_text, config, center), MASKED_SHINGLE_WORDS),
        shingle8=shingle_hashes(mask_document(authored_text, config, center), TEMPLATE_SHINGLE_WORDS),
        chars=len(authored_text),
        words=len(word_values(authored_text)),
        keyphrase_count=keyphrase_count,
        blank_grades=not bool(expected_grades),
        blank_schools=not bool(relevant_schools(config, center)),
    )
    validate_source_reuse(record, source_sentences, source_shingles, state)
    return record


def validate_corpus(records: list[PageRecord], state: AuditState) -> dict[str, object]:
    paragraph_pages: dict[str, set[str]] = defaultdict(set)
    section_pages: dict[str, set[str]] = defaultdict(set)
    document_pages: dict[str, set[str]] = defaultdict(set)
    masked_paragraph_pages: dict[str, set[str]] = defaultdict(set)
    masked_sentence_pages: dict[str, set[str]] = defaultdict(set)
    heading_sequence_pages: dict[str, set[str]] = defaultdict(set)
    heading_position_pages: list[dict[str, set[str]]] = [defaultdict(set) for _ in range(EXPECTED_MANUSCRIPT_SECTIONS)]
    for record in records:
        for paragraph in set(record.paragraphs):
            paragraph_pages[normalize_key(paragraph)].add(record.relative)
        for section in set(record.sections):
            section_pages[normalize_key(section)].add(record.relative)
        document_pages[normalize_key("\n".join(record.sections))].add(record.relative)
        for paragraph in set(record.masked_paragraphs):
            masked_paragraph_pages[paragraph].add(record.relative)
            for sentence in sentence_values(paragraph, MASKED_SENTENCE_MIN_CHARS):
                masked_sentence_pages[normalize_key(sentence)].add(record.relative)
        heading_sequence_pages[normalize_key("\n".join(record.masked_headings))].add(record.relative)
        for index, heading in enumerate(record.masked_headings[:EXPECTED_MANUSCRIPT_SECTIONS]):
            heading_position_pages[index][heading].add(record.relative)

    def duplicates(index: dict[str, set[str]]) -> list[dict[str, object]]:
        return [
            {"text": text[:300], "pages": sorted(pages)}
            for text, pages in index.items()
            if text and len(pages) > 1
        ]

    exact_paragraphs = duplicates(paragraph_pages)
    exact_sections = duplicates(section_pages)
    exact_documents = duplicates(document_pages)
    exact_paragraph_top = sorted(
        ((len(pages), text, sorted(pages)[:8]) for text, pages in paragraph_pages.items()),
        reverse=True,
    )
    exact_paragraph_max_df = exact_paragraph_top[0][0] if exact_paragraph_top else 0
    state.check(
        "duplicates.exact_paragraph",
        exact_paragraph_max_df <= EXACT_PARAGRAPH_DF_LIMIT,
        CATEGORY_SLUG,
        "exact authored paragraph document frequency is too high",
        {"max_df": exact_paragraph_max_df, "limit": EXACT_PARAGRAPH_DF_LIMIT, "top": exact_paragraph_top[:20]},
    )
    state.check("duplicates.exact_section", not exact_sections, CATEGORY_SLUG, "an exact authored section is reused across pages", exact_sections[:30])
    state.check("duplicates.exact_document", not exact_documents, CATEGORY_SLUG, "an exact authored document is reused across pages", exact_documents[:20])

    masked_paragraph_top = sorted(
        ((len(pages), text, sorted(pages)[:8]) for text, pages in masked_paragraph_pages.items()),
        reverse=True,
    )
    masked_sentence_top = sorted(
        ((len(pages), text, sorted(pages)[:8]) for text, pages in masked_sentence_pages.items()),
        reverse=True,
    )
    paragraph_max_df = masked_paragraph_top[0][0] if masked_paragraph_top else 0
    sentence_max_df = masked_sentence_top[0][0] if masked_sentence_top else 0
    state.check("duplicates.masked_paragraph_df", paragraph_max_df <= MASKED_PARAGRAPH_DF_LIMIT, CATEGORY_SLUG, "entity-masked paragraph document frequency is too high", {"max_df": paragraph_max_df, "limit": MASKED_PARAGRAPH_DF_LIMIT, "top": masked_paragraph_top[:20]})
    state.check("duplicates.masked_sentence_df", sentence_max_df <= MASKED_SENTENCE_DF_LIMIT, CATEGORY_SLUG, "entity-masked sentence document frequency is too high", {"max_df": sentence_max_df, "limit": MASKED_SENTENCE_DF_LIMIT, "top": masked_sentence_top[:20]})
    heading_sequence_duplicates = duplicates(heading_sequence_pages)
    state.check("duplicates.heading_sequence", not heading_sequence_duplicates, CATEGORY_SLUG, "the complete masked H2 sequence is reused", heading_sequence_duplicates[:20])
    heading_position_metrics: list[dict[str, object]] = []
    for index, values in enumerate(heading_position_pages, 1):
        maximum = max((len(pages) for pages in values.values()), default=0)
        heading_position_metrics.append({"position": index, "unique": len(values), "max_df": maximum})
        state.check("duplicates.heading_position_unique", len(values) >= HEADING_POSITION_UNIQUE_MIN, CATEGORY_SLUG, f"H2 position {index} has too few masked variants", {"unique": len(values), "minimum": HEADING_POSITION_UNIQUE_MIN})
        state.check("duplicates.heading_position_df", maximum <= HEADING_POSITION_DF_LIMIT, CATEGORY_SLUG, f"H2 position {index} repeats too often", {"max_df": maximum, "limit": HEADING_POSITION_DF_LIMIT})

    pair_scores: list[tuple[float, str, str]] = []
    for left, right in combinations(records, 2):
        union = left.shingle5 | right.shingle5
        score = len(left.shingle5 & right.shingle5) / len(union) if union else 1.0
        pair_scores.append((score, left.relative, right.relative))
    pair_scores.sort(reverse=True)
    pair_values = [score for score, _, _ in pair_scores]
    maximum = pair_values[0] if pair_values else 0.0
    p99 = percentile(pair_values, 0.99)
    p95 = percentile(pair_values, 0.95)
    state.check("duplicates.masked_jaccard_max", maximum < MASKED_MAX_LIMIT, CATEGORY_SLUG, "exhaustive masked five-word Jaccard maximum is too high", {"actual": round(maximum, 6), "limit_exclusive": MASKED_MAX_LIMIT, "pair": pair_scores[0][1:] if pair_scores else []})
    state.check("duplicates.masked_jaccard_p99", p99 < MASKED_P99_LIMIT, CATEGORY_SLUG, "masked Jaccard p99 is too high", {"actual": round(p99, 6), "limit_exclusive": MASKED_P99_LIMIT})
    state.check("duplicates.masked_jaccard_p95", p95 < MASKED_P95_LIMIT, CATEGORY_SLUG, "masked Jaccard p95 is too high", {"actual": round(p95, 6), "limit_exclusive": MASKED_P95_LIMIT})

    shingle_df = Counter(value for record in records for value in record.shingle8)
    common_threshold = math.ceil(len(records) * TEMPLATE_COMMON_PAGE_RATIO)
    common_shingles = {value for value, count in shingle_df.items() if count >= common_threshold}
    burdens = {
        record.relative: (len(record.shingle8 & common_shingles) / len(record.shingle8) if record.shingle8 else 1.0)
        for record in records
    }
    burden_values = list(burdens.values())
    burden_median = statistics.median(burden_values) if burden_values else 0.0
    burden_p95 = percentile(burden_values, 0.95)
    burden_max = max(burden_values, default=0.0)
    state.check("duplicates.template_burden_median", burden_median <= TEMPLATE_MEDIAN_LIMIT, CATEGORY_SLUG, "median common eight-word template burden is too high", {"actual": round(burden_median, 6), "limit": TEMPLATE_MEDIAN_LIMIT})
    state.check("duplicates.template_burden_p95", burden_p95 <= TEMPLATE_P95_LIMIT, CATEGORY_SLUG, "p95 common eight-word template burden is too high", {"actual": round(burden_p95, 6), "limit": TEMPLATE_P95_LIMIT})
    state.check("duplicates.template_burden_max", burden_max <= TEMPLATE_MAX_LIMIT, CATEGORY_SLUG, "maximum common eight-word template burden is too high", {"actual": round(burden_max, 6), "limit": TEMPLATE_MAX_LIMIT})

    return {
        "exact": {
            "paragraph_duplicates": len(exact_paragraphs),
            "paragraph_max_df": exact_paragraph_max_df,
            "paragraph_df_limit": EXACT_PARAGRAPH_DF_LIMIT,
            "section_duplicates": len(exact_sections),
            "document_duplicates": len(exact_documents),
        },
        "masked_repetition": {
            "paragraph_max_df": paragraph_max_df,
            "sentence_max_df": sentence_max_df,
            "top_paragraphs": [
                {"df": count, "text": text[:300], "sample": pages}
                for count, text, pages in masked_paragraph_top[:20]
            ],
            "top_sentences": [
                {"df": count, "text": text[:300], "sample": pages}
                for count, text, pages in masked_sentence_top[:20]
            ],
        },
        "headings": {
            "sequence_duplicates": len(heading_sequence_duplicates),
            "positions": heading_position_metrics,
        },
        "masked_jaccard": {
            "method": "exhaustive all 68,635 pairs, entity-masked five-word shingle Jaccard",
            "pairs": len(pair_scores),
            "distribution": distribution(pair_values),
            "top_pairs": [
                {"jaccard": round(score, 6), "left": left, "right": right}
                for score, left, right in pair_scores[:30]
            ],
        },
        "template_burden": {
            "shingle_words": TEMPLATE_SHINGLE_WORDS,
            "common_document_frequency": common_threshold,
            "common_shingles": len(common_shingles),
            "distribution": distribution(burden_values),
            "worst_pages": [
                {"burden": round(value, 6), "path": path}
                for path, value in sorted(burdens.items(), key=lambda item: (-item[1], item[0]))[:30]
            ],
        },
    }


def validate_meta_uniqueness(
    descriptions: list[tuple[str, str]],
    records: list[PageRecord],
    state: AuditState,
) -> dict[str, object]:
    exact: dict[str, list[str]] = defaultdict(list)
    masked: dict[str, list[str]] = defaultdict(list)
    config = CONFIG_BY_SLUG[CATEGORY_SLUG]
    center_by_path = {record.relative: record.center for record in records}
    for path, value in descriptions:
        exact[normalize_key(value)].append(path)
        center = center_by_path.get(path)
        if center:
            masked[mask_document(value, config, center)].append(path)
    exact_duplicates = {text: paths for text, paths in exact.items() if text and len(paths) > 1}
    max_masked_df = max((len(paths) for paths in masked.values()), default=0)
    state.check("seo.description_unique", not exact_duplicates, CATEGORY_SLUG, "meta descriptions are not unique", list(exact_duplicates.values())[:20])
    state.check("seo.description_masked_df", max_masked_df <= MASKED_META_DF_LIMIT, CATEGORY_SLUG, "entity-masked meta description pattern repeats too often", {"max_df": max_masked_df, "limit": MASKED_META_DF_LIMIT})
    return {"exact_duplicates": len(exact_duplicates), "masked_max_df": max_masked_df}


def recommended_samples(records: list[PageRecord], corpus: dict[str, object]) -> list[dict[str, object]]:
    reasons: dict[str, set[str]] = defaultdict(set)
    jaccard = corpus.get("masked_jaccard", {})
    for pair in jaccard.get("top_pairs", [])[:10] if isinstance(jaccard, dict) else []:
        reasons[str(pair["left"])].add("top masked-similarity pair")
        reasons[str(pair["right"])].add("top masked-similarity pair")
    burden = corpus.get("template_burden", {})
    for item in burden.get("worst_pages", [])[:10] if isinstance(burden, dict) else []:
        reasons[str(item["path"])].add("high template burden")
    for record in sorted(records, key=lambda item: (item.chars, item.relative))[:5]:
        reasons[record.relative].add("shortest authored text")
    for record in sorted(records, key=lambda item: (-item.chars, item.relative))[:5]:
        reasons[record.relative].add("longest authored text")
    for record in records:
        if record.blank_grades:
            reasons[record.relative].add("blank middle-Math grades")
    for record in [item for item in records if item.blank_schools][:5]:
        reasons[record.relative].add("blank school facts")
    center_groups: dict[tuple[str, str], list[PageRecord]] = defaultdict(list)
    for record in records:
        center_groups[(str(record.center["center_name"]), str(record.center["address"]))].append(record)
    largest = max(center_groups.values(), key=len, default=[])
    for record in largest:
        reasons[record.relative].add("largest shared-center cluster")
    seeded = random.Random(20260901)
    for record in seeded.sample(records, min(5, len(records))):
        reasons[record.relative].add("deterministic random coverage")
    ordered = sorted(reasons.items(), key=lambda item: (-len(item[1]), item[0]))[:50]
    return [{"path": path, "reasons": sorted(values)} for path, values in ordered]


def automated_score(state: AuditState) -> dict[str, object]:
    errors = [finding for finding in state.findings if finding.severity == "error"]
    warnings = [finding for finding in state.findings if finding.severity == "warning"]

    def clear(prefixes: tuple[str, ...]) -> bool:
        return not any(finding.code.startswith(prefixes) for finding in errors)

    criteria = [
        ("individuality_exact", 6, ("duplicates.exact",)),
        ("individuality_similarity", 10, ("duplicates.masked_jaccard",)),
        ("individuality_template", 8, ("duplicates.template", "duplicates.masked_paragraph", "duplicates.masked_sentence")),
        ("individuality_headings", 6, ("duplicates.heading",)),
        ("facts_csv_baseline", 15, ("facts.",)),
        ("scope_isolation", 5, ("scope.",)),
        ("source_independence", 5, ("source.",)),
        ("natural_korean", 10, ("natural.", "claims.no_overclaim", "claims.no_review")),
        ("reader_value", 10, ("reader.",)),
        ("seo", 6, ("seo.",)),
        ("aeo", 5, ("aeo.",)),
        ("geo_claim_grounding", 4, ("claims.",)),
        ("technical", 10, ("technical.",)),
    ]
    awarded: dict[str, int] = {}
    for name, points, prefixes in criteria:
        awarded[name] = points if clear(prefixes) else 0
    score = sum(awarded.values())
    warning_penalty = min(5, math.ceil(len(warnings) / max(1, EXPECTED_DETAILS // 5)))
    score = max(0, score - warning_penalty)
    release = not errors and score == 100
    return {
        "automated_score": score,
        "maximum": 100,
        "warning_penalty": warning_penalty,
        "criteria": awarded,
        "hard_error_count": len(errors),
        "warning_count": len(warnings),
        "release_gate": "pass" if release else "fail",
        "manual_review_required": True,
        "grading": "100 only: release candidate; any deduction or hard finding requires repair and rerun",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-ref", required=True, help="Git commit immediately before the middle-Math rewrite")
    parser.add_argument("--report", type=Path, help="Write the full JSON report outside the repository")
    args = parser.parse_args()

    if args.report:
        report_path = args.report.expanduser().resolve()
        try:
            report_path.relative_to(ROOT.resolve())
        except ValueError:
            pass
        else:
            parser.error("--report must be outside the repository")
    else:
        report_path = None

    state = AuditState()
    try:
        baseline = resolve_baseline(args.baseline_ref)
    except RuntimeError as exc:
        parser.error(f"invalid --baseline-ref: {exc}")

    config = CONFIG_BY_SLUG[CATEGORY_SLUG]
    centers = load_centers()
    state.check("scope.center_count", len(centers) == EXPECTED_ROWS == EXPECTED_DETAILS, "center CSV", "center CSV must contain exactly 371 rows", len(centers))
    blank_grade_localities = {
        str(center["locality"])
        for center in centers
        if not relevant_grades(config, center, "수학")
    }
    blank_school_localities = {
        str(center["locality"])
        for center in centers
        if not relevant_schools(config, center)
    }
    state.check(
        "facts.blank_grade_source_contract",
        blank_grade_localities == EXPECTED_BLANK_GRADE_LOCALITIES,
        "center CSV",
        "middle-Math blank-grade rows differ from the audited 13-row source contract",
        {
            "actual": sorted(blank_grade_localities),
            "expected": sorted(EXPECTED_BLANK_GRADE_LOCALITIES),
        },
    )
    state.check(
        "facts.blank_school_source_contract",
        blank_school_localities == EXPECTED_BLANK_SCHOOL_LOCALITIES,
        "center CSV",
        "middle-Math blank-school rows differ from the audited 53-row source contract",
        {
            "actual": sorted(blank_school_localities),
            "expected": sorted(EXPECTED_BLANK_SCHOOL_LOCALITIES),
        },
    )
    expected_paths = {TARGET_ROOT / config.slug / str(center["slug"]) / "index.html" for center in centers}
    actual_paths = set((TARGET_ROOT / config.slug).glob("*/index.html"))
    state.check("scope.filesystem", actual_paths == expected_paths, CATEGORY_SLUG, "middle-Math detail filesystem scope differs from the 371 CSV rows", {"actual": len(actual_paths), "expected": len(expected_paths), "missing": [display_path(path) for path in sorted(expected_paths - actual_paths)], "extra": [display_path(path) for path in sorted(actual_paths - expected_paths)]})
    scope_metrics = validate_scope(baseline, expected_paths, state)

    source_rows = load_source_rows(config)
    source_sentences: set[str] = set()
    source_shingles: set[int] = set()
    contamination_rows: list[str] = []
    for center, value in zip(centers, source_rows):
        visible = source_text(value)
        source_sentences.update(sentence_values(visible, MIN_SOURCE_SENTENCE_CHARS))
        source_shingles.update(shingle_hashes(visible, SOURCE_SHINGLE_WORDS))
        locality = str(center["locality"])
        contract = SOURCE_H1_CONTAMINATION_CONTRACT.get(locality)
        if contract:
            contaminated, _safe_focus = contract
            source_h1 = (single_tag_text(value, "h1") or [""])[0]
            state.check(
                "source.polluted_h1_contract",
                contaminated in source_h1,
                config.source_path,
                f"the audited source H1 contamination contract drifted for {locality}",
                {"expected_fragment": contaminated, "actual_h1": source_h1},
            )
            if contaminated in source_h1:
                contamination_rows.append(locality)
    state.check(
        "source.polluted_h1_count",
        set(contamination_rows) == set(SOURCE_H1_CONTAMINATION_CONTRACT),
        config.source_path,
        "the six known wrong-subject, wrong-level or wrong-assessment source H1 rows were not identified exactly",
        {"actual": sorted(contamination_rows), "expected": sorted(SOURCE_H1_CONTAMINATION_CONTRACT)},
    )

    all_localities = {str(center["locality"]) for center in centers}
    all_schools = {school for center in centers for school in relevant_schools(config, center)}
    all_centers = {str(center["center_name"]) for center in centers}
    representative_digests: list[str] = []
    descriptions: list[tuple[str, str]] = []
    records: list[PageRecord] = []
    for center in centers:
        path = TARGET_ROOT / config.slug / str(center["slug"]) / "index.html"
        if not path.is_file():
            continue
        record = build_page_record(
            path,
            center,
            baseline,
            all_localities,
            all_schools,
            all_centers,
            source_sentences,
            source_shingles,
            representative_digests,
            descriptions,
            state,
        )
        if record:
            records.append(record)

    state.check("scope.records", len(records) == EXPECTED_DETAILS, CATEGORY_SLUG, "could not audit all 371 detail pages", len(records))
    state.check(
        "facts.blank_grade_rendered_count",
        sum(record.blank_grades for record in records) == 13,
        CATEGORY_SLUG,
        "rendered blank-grade page count is not 13",
        sum(record.blank_grades for record in records),
    )
    state.check(
        "facts.blank_school_rendered_count",
        sum(record.blank_schools for record in records) == 53,
        CATEGORY_SLUG,
        "rendered blank-school page count is not 53",
        sum(record.blank_schools for record in records),
    )
    state.check("technical.representative_unique", len(representative_digests) == EXPECTED_DETAILS and len(set(representative_digests)) == EXPECTED_DETAILS, CATEGORY_SLUG, "representative image contents are not unique within middle Math", {"files": len(representative_digests), "unique": len(set(representative_digests))})
    corpus_metrics = validate_corpus(records, state) if len(records) == EXPECTED_DETAILS else {}
    meta_metrics = validate_meta_uniqueness(descriptions, records, state)
    expected_urls = {record.canonical for record in records}
    sitemap_metrics = validate_sitemap(baseline, expected_urls, state)
    score = automated_score(state)

    severity_counts = Counter(finding.severity for finding in state.findings)
    code_counts = Counter(finding.code for finding in state.findings)
    report: dict[str, object] = {
        "schema_version": 1,
        "status": "pass" if score["release_gate"] == "pass" else "fail",
        "category": CATEGORY_SLUG,
        "baseline_commit": baseline,
        "contract": {
            "details": EXPECTED_DETAILS,
            "source_sentence_min_chars": MIN_SOURCE_SENTENCE_CHARS,
            "source_shingle_words": SOURCE_SHINGLE_WORDS,
            "masked_shingle_words": MASKED_SHINGLE_WORDS,
            "masked_jaccard_limits": {"max_exclusive": MASKED_MAX_LIMIT, "p99_exclusive": MASKED_P99_LIMIT, "p95_exclusive": MASKED_P95_LIMIT},
            "template_burden_limits": {"median": TEMPLATE_MEDIAN_LIMIT, "p95": TEMPLATE_P95_LIMIT, "max": TEMPLATE_MAX_LIMIT},
            "exact_repetition_limits": {"paragraph_df": EXACT_PARAGRAPH_DF_LIMIT, "section_df": 1, "document_df": 1},
            "masked_repetition_limits": {"paragraph_df": MASKED_PARAGRAPH_DF_LIMIT, "sentence_df": MASKED_SENTENCE_DF_LIMIT},
        },
        "scope": scope_metrics,
        "counts": {
            "records": len(records),
            "authored_characters": distribution(record.chars for record in records),
            "authored_words": distribution(record.words for record in records),
            "keyphrase_count": distribution(record.keyphrase_count for record in records),
            "blank_grade_pages": sum(record.blank_grades for record in records),
            "blank_school_pages": sum(record.blank_schools for record in records),
        },
        "source": {
            "rows": len(source_rows),
            "unique_sentences_min_42_chars": len(source_sentences),
            "unique_twelve_word_shingles": len(source_shingles),
        },
        "corpus": corpus_metrics,
        "meta": meta_metrics,
        "sitemap": sitemap_metrics,
        "recommended_manual_samples": recommended_samples(records, corpus_metrics) if records else [],
        "score": score,
        "check_counts": dict(sorted(state.checks.items())),
        "finding_counts": {
            "by_severity": dict(sorted(severity_counts.items())),
            "by_code": dict(sorted(code_counts.items())),
        },
        "findings": [finding.as_dict() for finding in state.findings],
    }
    output = json.dumps(report, ensure_ascii=False, indent=2, default=str)
    print(output)
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(output + "\n", encoding="utf-8")
    return 0 if score["release_gate"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
