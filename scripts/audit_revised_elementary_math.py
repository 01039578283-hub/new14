#!/usr/bin/env python3
"""Strict release audit for the 371 revised elementary-math detail pages.

This read-only audit deliberately reuses the latest elementary-English audit
engine so both elementary subjects retain the same hard gates. Only the
subject, source-workbook contracts, factual exceptions, and mathematics
language signals are adapted here. Reports must be written outside the
repository.
"""

from __future__ import annotations

import inspect
import re
import sys
import textwrap
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_revised_elementary_english as core
from generate_subject_pages_from_xlsx import (
    ELEMENTARY_MATH_SAFE_FOCUS_OVERRIDES,
    ELEMENTARY_MATH_TRANSITION_FOCUS_OVERRIDES,
)


CATEGORY_SLUG = "초등수학학원"

SOURCE_H1_UNSAFE_FRAGMENTS: dict[str, str] = {
    "구의동": "내신과 모의고사 준비의 시작",
    "북가좌동": "모의고사 시간 배분까지 준비하는 수업",
    "권선동": "예비고1 수학 개념 점검 가이드",
    "금오동": "단원평가부터 내신까지 연결하는 공부법",
    "이충동": "예비고1 입학 전 학습 기반 준비",
    "수곡동": "내신 대비를 시작하는 학습 흐름",
    "도남지구": "내신평가와 학습 습관 관리",
    "복현동": "내신성적을 위한 맞춤 학습관리",
    "침산동": "기초부터 내신 관리까지",
    "칠성동": "방학에도 이어지는 내신진도관리",
    "수성동": "기초부터 내신과제관리까지",
    "만촌동": "개념부터 내신까지 이어지는 학습관리",
    "신천동": "내신특강으로 공부 자신감 키우기",
    "달동": "예비고1 영어를 준비하는 겨울방학",
}

# Import the exact safe-focus dictionary used by the generator. This prevents
# the audit and generator from silently drifting to different promises.
if set(SOURCE_H1_UNSAFE_FRAGMENTS) != set(ELEMENTARY_MATH_SAFE_FOCUS_OVERRIDES):
    raise RuntimeError("elementary-math safe-focus source contract drifted")
SOURCE_H1_CONTAMINATION_CONTRACT: dict[str, tuple[str, str]] = {
    locality: (
        SOURCE_H1_UNSAFE_FRAGMENTS[locality],
        ELEMENTARY_MATH_SAFE_FOCUS_OVERRIDES[locality],
    )
    for locality in SOURCE_H1_UNSAFE_FRAGMENTS
}

# These are valid only when the supplied math grade range includes 초6. The
# revised copy must keep them explicitly elementary-to-middle, never present a
# middle-school program as a supplied fact.
SOURCE_MIDDLE_TRANSITION_FRAGMENTS: dict[str, str] = {
    "삼각산동": "중1 첫 시험 준비",
    "보라동": "중등 준비를 시작하는 학습기록 관리",
    "풍덕천동": "중학교 수학을 준비하는 학습노트",
    "영덕동": "중등 준비까지 연결하는 학습코칭",
    "부발읍": "예비 중학생을 위한",
    "상남동": "중1 첫 시험까지 준비",
}
if set(SOURCE_MIDDLE_TRANSITION_FRAGMENTS) != set(ELEMENTARY_MATH_TRANSITION_FOCUS_OVERRIDES):
    raise RuntimeError("elementary-math transition source contract drifted")
SOURCE_MIDDLE_TRANSITION_CONTRACT: dict[str, tuple[str, str]] = {
    locality: (
        SOURCE_MIDDLE_TRANSITION_FRAGMENTS[locality],
        ELEMENTARY_MATH_TRANSITION_FOCUS_OVERRIDES[locality],
    )
    for locality in SOURCE_MIDDLE_TRANSITION_FRAGMENTS
}

EXPECTED_BLANK_GRADE_LOCALITIES = {
    "진관동", "구파발", "갈현동", "다산동", "다산신도시", "부천 중동",
    "약대동", "고잔동", "초지동", "둔산동", "탄방동", "석사동", "퇴계동",
}

EXPECTED_BLANK_SCHOOL_LOCALITIES = {
    "명일동", "천호동", "광장동", "구의동", "하계동", "중계동", "종암동", "길음동",
    "구산동", "역촌동", "화정동", "토당동", "일산동", "후곡마을", "광명동", "철산동",
    "하안동", "탄벌동", "경안동", "배곧", "배곧동", "정왕동", "고잔동", "초지동",
    "옥정동", "옥정신도시", "세교", "금암동", "신곡동", "금오동", "산내마을", "목동동",
    "야당동", "운정호수", "이충동", "서정동", "송탄", "풍산동", "미사", "미사신도시",
    "반송동", "석우동", "봉담2지구", "봉담읍", "청라", "동춘동", "연수동", "관저동",
    "원내동", "둔산동", "탄방동", "불당동", "천안 백석동", "신불당", "만촌동", "범어동",
    "신천동", "울산 삼산동", "달동", "복산동", "약사동", "반구동", "석동", "자은동",
    "경화동", "두호동", "장성동", "수완동", "수완지구", "신가동", "전주 장동",
    "전주혁신도시", "송천동", "후평동",
}

SOURCE_H1_LOCALITY_ALIASES = {
    # The center CSV uses a disambiguated locality while the supplied workbook
    # H1 uses the shorter local name.
    "부천 상동": ("부천 상동", "상동"),
}

ELEMENTARY_MATH_LABELS = (
    "수 감각과 연산 원리", "문장제 조건", "분수·소수 이해", "도형과 측정",
    "수학 개념 설명", "수학 오답 재학습", "서술형 풀이", "그래프와 자료 해석",
    "단원 간 연결", "수학 시간 배분", "수학 학습 루틴", "학교 수학·단원평가",
    "학습 우선순위", "질문·혼자 설명", "초6·중1 전환 준비",
)

ELEMENTARY_MATH_SIGNAL_TERMS = (
    "연산", "계산", "수 감각", "자릿값", "받아올림", "받아내림",
    "문장제", "조건", "식", "개념", "원리", "분수", "소수", "도형",
    "측정", "각도", "넓이", "수직선", "풀이", "오답", "재풀이",
    "서술형", "그래프", "표", "검산",
)

ARTIFACT_TERMS = (
    "시험지", "시험 범위표", "범위표", "학교 진도표", "교재", "학습지",
    "개념서", "유형서", "연산 기록", "풀이 과정", "조건", "식", "그림",
    "수직선", "도형", "표", "그래프", "오답", "재풀이", "학습 기록",
    "주간 계획표", "계획표", "진단표", "달력", "문항 번호", "문제 번호",
    "답안", "단원 평가", "단원표", "공식", "평가 조건표", "수행평가", "서술형",
)

ACTION_TERMS = (
    "표시", "비교", "분류", "설명", "기록", "다시 풀", "재확인", "밑줄",
    "질문", "구분", "검산", "식 세우", "그림으로 나타내", "수직선에 나타내",
    "소리 내어 설명", "적으",
)

FOCUS_EVIDENCE_TERMS = (
    "시험지", "교재", "학습지", "범위표", "학교 진도표", "오답", "답안",
    "근거", "표시", "풀이", "기록", "진단표", "식", "그림", "수직선",
    "도형", "표", "그래프",
)

FOCUS_ACTION_TERMS = (
    "표시", "비교", "분류", "구분", "설명", "기록", "다시 풀", "밑줄",
    "질문", "적으", "배치", "검산", "식 세우", "그림으로", "수직선에",
)

FOCUS_CHECKPOINT_TERMS = (
    "재확인", "다음 점검", "확인일", "완료일", "일주일 뒤", "며칠 뒤",
    "다음 주", "변화", "다시 설명", "판단", "새 문제", "다른 문제", "재풀이",
)

FOCUS_REPEAT_STOPWORDS = {
    "초등", "수학", "학원", "학생", "학교", "지역", "센터", "학년", "시험",
    "자료", "확인", "기준", "학습", "문제", "풀이",
}

ALWAYS_FORBIDDEN_CONTAMINATION_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("english_subject", re.compile(r"(?:영어|파닉스|알파벳|영단어|영문법|영작|구문\s*독해|영어\s*(?:듣기|말하기))")),
    ("high_school_level", re.compile(r"(?:예비\s*고1|고등\s*(?:수학|학교|학생|과정)|(?<![0-9A-Za-z가-힣])고[1-3](?![0-9A-Za-z가-힣]))")),
    ("csat_or_mock_exam", re.compile(r"(?:수능|모의고사|전국연합|학력평가|(?<![가-힣])모고(?![가-힣]))")),
    ("admissions_frame", re.compile(r"(?:입시|대입|수시|정시)")),
    ("adult_exam", re.compile(r"(?:토익|토플|아이엘츠|성인\s*영어)")),
    ("other_subject", re.compile(r"(?:국어|과학)\s*(?:학습|수업|문제|과목|시험|교재)?")),
)

# Source-H1 classification is intentionally narrower than the rendered-copy
# prohibition above. Admissions language is sanitized generically; these four
# patterns define the 14 rows with explicit safe-focus overrides.
SOURCE_H1_CONTAMINATION_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("english_subject", re.compile(r"(?:영어|파닉스|알파벳|영단어|영작)")),
    ("high_school_level", re.compile(r"(?:예비\s*고1|고등|(?<![0-9A-Za-z가-힣])고[1-3](?![0-9A-Za-z가-힣]))")),
    ("mock_exam", re.compile(r"(?:수능|모의고사|전국연합|학력평가|(?<![가-힣])모고(?![가-힣]))")),
    ("middle_high_assessment", re.compile(r"(?<!별)내신")),
)

MATH_AWKWARD_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("duplicate_math_explanation", re.compile(r"수학\s+개념\s+설명(?:을|을\s+다시)?\s+설명")),
    ("duplicate_word_problem_condition", re.compile(r"문장제\s+조건(?:을|의)?\s+조건")),
    ("duplicate_equation_action", re.compile(r"식\s+세우기(?:를|의)?\s+식을?\s+세우")),
    ("awkward_number_word_problem_join", re.compile(r"수\s+감각과\s+연산\s+원리와\s+문장제\s+조건")),
    ("awkward_fraction_geometry_join", re.compile(r"분수·소수\s+이해와\s+도형과\s+측정")),
    ("unspaced_math_compound", re.compile(r"(?:조건해석|식세우기|계산정확도|풀이과정|단원연결|문제풀이|학교시험|시험복습|시험분석|시험오답|시험전략)")),
)

_CONFIGURED = False


def source_h1_identity_matches(locality: str, source_h1: str) -> bool:
    aliases = SOURCE_H1_LOCALITY_ALIASES.get(locality, (locality,))
    return any(alias in source_h1 for alias in aliases)


def source_h1_is_allowed_transition(locality: str, source_h1: str) -> bool:
    return locality in SOURCE_MIDDLE_TRANSITION_CONTRACT and bool(
        re.search(
            r"(?:예비\s*중학생|중학생|중학교|중등|(?<![0-9A-Za-z가-힣])중1(?![0-9A-Za-z가-힣]))",
            source_h1,
        )
    )


def adapt_function(function: object) -> None:
    """Recompile one core function with subject-specific literal changes."""

    source = textwrap.dedent(inspect.getsource(function))
    source = source.replace("elementary-English", "elementary-Math")
    source = source.replace("elementary English", "elementary Math")
    source = source.replace("초등 영어", "초등 수학")
    source = source.replace('"영어"', '"수학"')
    source = source.replace('"중학교영어"', '"중학교수학"')
    source = source.replace('"중등영어"', '"중등수학"')
    source = source.replace("visible English grades", "visible Math grades")
    source = source.replace('"reader.elementary_english_signal_density"', '"reader.elementary_math_signal_density"')
    source = source.replace(
        "fewer than three concrete elementary-Math learning signals",
        "fewer than three concrete elementary-math learning signals",
    )
    source = source.replace(
        "Math, high-school/adult-exam, CSAT/mock-exam or invalid middle-school transition contamination",
        "English, admissions, high-school/adult-exam, CSAT/mock-exam or invalid middle-school transition contamination",
    )
    source = source.replace(
        "if locality not in source_h1:\n            source_h1_locality_mismatches.append(locality)",
        "if not source_h1_identity_matches(locality, source_h1):\n            source_h1_locality_mismatches.append(locality)",
    )
    source = source.replace(
        'sum(record.blank_grades for record in records) == 8, CATEGORY_SLUG, "rendered blank-grade page count is not 8"',
        'sum(record.blank_grades for record in records) == len(EXPECTED_BLANK_GRADE_LOCALITIES), CATEGORY_SLUG, "rendered blank-grade page count differs from the source contract"',
    )
    if getattr(function, "__name__", "") == "main":
        source = source.replace(
            "ALWAYS_FORBIDDEN_CONTAMINATION_PATTERNS",
            "SOURCE_H1_CONTAMINATION_PATTERNS",
        )
        source = source.replace(
            'if re.search(r"(?:예비\\s*중학생|중학생|중학교|중등|(?<![0-9A-Za-z가-힣])중1(?![0-9A-Za-z가-힣]))", source_h1):\n            detected_middle_transition_rows.append(locality)',
            "if source_h1_is_allowed_transition(locality, source_h1):\n            detected_middle_transition_rows.append(locality)",
        )
        source = source.replace("the nine known", "the 14 audited")
        source = source.replace("the nine-row contract", "the 14-row contract")
    if getattr(function, "__name__", "") == "validate_natural_reader_value":
        source = source.replace(
            "    # Locality is source-backed entity data rather than editorial subject\n",
            "    transition_context = transition_context or (\n"
            "        locality in SOURCE_MIDDLE_TRANSITION_CONTRACT\n"
            "        and \"초6\" in expected_grades\n"
            "    )\n"
            "    # Locality is source-backed entity data rather than editorial subject\n",
        )
    namespace = core.__dict__
    exec(compile(source, str(Path(core.__file__).resolve()), "exec"), namespace)


def configure_core() -> None:
    global _CONFIGURED
    if _CONFIGURED:
        return

    functions = [
        value
        for value in tuple(core.__dict__.values())
        if inspect.isfunction(value) and value.__module__ == core.__name__
    ]
    for function in functions:
        adapt_function(function)

    core.__doc__ = __doc__
    core.CATEGORY_SLUG = CATEGORY_SLUG
    core.ALLOWED_NON_DETAIL_PATHS = {
        "sitemap.xml",
        "scripts/audit_revised_high_english.py",
        "scripts/audit_revised_high_math.py",
        "scripts/audit_revised_middle_english.py",
        "scripts/audit_revised_middle_math.py",
        "scripts/audit_revised_elementary_english.py",
        "scripts/audit_revised_elementary_math.py",
        "scripts/generate_subject_pages_from_xlsx.py",
    }
    core.PRIOR_DETAIL_PREFIXES = (
        "과목별학원/고등영어학원/",
        "과목별학원/고등수학학원/",
        "과목별학원/중등영어학원/",
        "과목별학원/중등수학학원/",
        "과목별학원/초등영어학원/",
    )
    core.SITEMAP_PRESERVED_SUBJECT_SLUGS = {
        "고등영어학원",
        "고등수학학원",
        "중등영어학원",
        "중등수학학원",
        "초등영어학원",
    }
    core.SOURCE_H1_CONTAMINATION_CONTRACT = SOURCE_H1_CONTAMINATION_CONTRACT
    core.SOURCE_MIDDLE_TRANSITION_CONTRACT = SOURCE_MIDDLE_TRANSITION_CONTRACT
    core.EXPECTED_BLANK_GRADE_LOCALITIES = EXPECTED_BLANK_GRADE_LOCALITIES
    core.EXPECTED_BLANK_SCHOOL_LOCALITIES = EXPECTED_BLANK_SCHOOL_LOCALITIES
    core.SOURCE_H1_LOCALITY_ALIASES = SOURCE_H1_LOCALITY_ALIASES
    core.source_h1_identity_matches = source_h1_identity_matches
    core.SOURCE_H1_CONTAMINATION_PATTERNS = SOURCE_H1_CONTAMINATION_PATTERNS
    core.source_h1_is_allowed_transition = source_h1_is_allowed_transition
    core.ELEMENTARY_ENGLISH_LABELS = ELEMENTARY_MATH_LABELS
    core.ELEMENTARY_ENGLISH_SIGNAL_TERMS = ELEMENTARY_MATH_SIGNAL_TERMS
    core.ALWAYS_FORBIDDEN_CONTAMINATION_PATTERNS = ALWAYS_FORBIDDEN_CONTAMINATION_PATTERNS
    core.ARTIFACT_TERMS = ARTIFACT_TERMS
    core.ACTION_TERMS = ACTION_TERMS
    core.FOCUS_EVIDENCE_TERMS = FOCUS_EVIDENCE_TERMS
    core.FOCUS_ACTION_TERMS = FOCUS_ACTION_TERMS
    core.FOCUS_CHECKPOINT_TERMS = FOCUS_CHECKPOINT_TERMS
    core.FOCUS_REPEAT_STOPWORDS = FOCUS_REPEAT_STOPWORDS
    core.AWKWARD_PATTERNS = tuple(core.AWKWARD_PATTERNS) + MATH_AWKWARD_PATTERNS
    _CONFIGURED = True


def main() -> int:
    configure_core()
    return core.main()


if __name__ == "__main__":
    raise SystemExit(main())
