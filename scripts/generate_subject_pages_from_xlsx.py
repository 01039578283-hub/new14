from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import random
import re
import shutil
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import quote

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
COMMON = ROOT.parent / "참고자료" / "공통자료"
SOURCE_DIR = Path.home() / "Desktop" / "구글시트로 뽑은거"
CENTER_CSV = COMMON / "센터정보 정리.csv"
IMAGE_CSV = COMMON / "이미지링크.csv"
REP_SOURCE = COMMON / "대표이미지"
REP_TARGET = ROOT / "assets" / "representative"
MAP_DIR = ROOT / "assets" / "maps"
TARGET_ROOT = ROOT / "과목별학원"
ORIGIN = "https://xn--3e0bz50b1zcyxat54c.com"
SITE_NAME = "와와학습코칭센터"
PHONE = "010-6839-8283"
CONTENT_DATE = "2026-08-15"
EXPECTED_ROWS = 371
REGION_ORDER = ("서울", "경기", "인천", "충청", "대전", "대구", "울산", "부산", "경상", "광주", "전라", "강원", "제주")


@dataclass(frozen=True)
class CategoryConfig:
    slug: str
    label: str
    workbook: str
    level: str
    grade_prefix: str
    school_name: str
    school_columns: tuple[str, ...]
    subjects: tuple[str, ...]
    english: str
    focus: str
    process: str
    rep_prefix: str
    source_mode: str = "single"

    @property
    def source_path(self) -> Path:
        return SOURCE_DIR / self.workbook

    @property
    def is_elementary(self) -> bool:
        return self.grade_prefix == "초"


CONFIGS: tuple[CategoryConfig, ...] = (
    CategoryConfig("영수학원", "영수학원", "영수학원 원고.xlsx", "초·중·고", "", "학교", ("타깃학교\n(초)", "타깃학교\n(중)", "타깃학교\n(고)"), ("영어", "수학"), "ENGLISH · MATH", "과목별 병목·주간 균형", "두 과목 실행 점검", "ys", "combined"),
    CategoryConfig("초등영어학원", "초등 영어학원", "초등 영어학원.xlsx", "초등", "초", "초등학교", ("타깃학교\n(초)",), ("영어",), "ELEMENTARY ENGLISH", "어휘·문장 읽기·기초 표현", "교과 이해와 습관", "es-eng"),
    CategoryConfig("초등수학학원", "초등 수학학원", "초등 수학학원.xlsx", "초등", "초", "초등학교", ("타깃학교\n(초)",), ("수학",), "ELEMENTARY MATH", "연산 원리·개념 설명·문장제", "풀이 과정 점검", "es-math"),
    CategoryConfig("중등영어학원", "중등 영어학원", "중등 영어학원.xlsx", "중등", "중", "중학교", ("타깃학교\n(중)",), ("영어",), "MIDDLE ENGLISH", "어휘·문법 적용·독해 근거", "내신 자료 연결", "ms-eng"),
    CategoryConfig("중등수학학원", "중등 수학학원", "중등 수학학원.xlsx", "중등", "중", "중학교", ("타깃학교\n(중)",), ("수학",), "MIDDLE MATH", "개념 연결·조건 해석·오답", "내신 풀이 점검", "ms-math"),
    CategoryConfig("고등영어학원", "고등 영어학원", "고등 영어학원.xlsx", "고등", "고", "고등학교", ("타깃학교\n(고)",), ("영어",), "HIGH SCHOOL ENGLISH", "어휘 누적·구문·독해 근거", "내신·모의고사 연결", "hs-eng"),
    CategoryConfig("고등수학학원", "고등 수학학원", "고등 수학학원.xlsx", "고등", "고", "고등학교", ("타깃학교\n(고)",), ("수학",), "HIGH SCHOOL MATH", "개념 통합·조건 해석·서술형", "내신·모의고사 오답", "hs-math"),
)

CONFIG_BY_SLUG = {config.slug: config for config in CONFIGS}
ALL_CATEGORY_SLUGS = (
    "고등학생국영수학원", "중학생국영수학원", "초등학생국영수학원",
    *(config.slug for config in CONFIGS),
)
ALL_CATEGORY_LABELS = {
    "고등학생국영수학원": "고등학생 국영수학원",
    "중학생국영수학원": "중학생 국영수학원",
    "초등학생국영수학원": "초등학생 국영수학원",
    **{config.slug: config.label for config in CONFIGS},
}


@dataclass(frozen=True)
class TopicSignal:
    code: str
    label: str
    keywords: tuple[str, ...]
    check: str
    evidence: str
    practice: str
    home_action: str


ENGLISH_SIGNALS: tuple[TopicSignal, ...] = (
    TopicSignal("vocabulary", "어휘 회상", ("어휘", "단어", "암기"), "뜻을 외운 단어를 문장 안에서 다시 알아보는지", "단어 시험과 독해 지문에서 같은 어휘를 놓친 기록", "뜻·품사·예문을 한 묶음으로 확인하기", "하루 뒤 예문 속 단어 뜻을 다시 말해 보기"),
    TopicSignal("sentence", "문장 구조", ("구문", "문장", "해석", "구조"), "주어와 동사, 수식 범위를 구분해 문장을 읽는지", "긴 문장에서 해석이 끊긴 위치와 표시한 문장 성분", "문장 뼈대를 먼저 적고 수식어를 붙여 읽기", "한 문장을 짧게 끊어 소리 내어 설명하기"),
    TopicSignal("grammar", "문법 적용", ("문법", "어법", "시제", "관계사"), "규칙을 아는 문제와 실제 문장에 적용하지 못한 문제를 나눴는지", "고른 답의 근거를 문법 규칙으로 설명한 흔적", "규칙 한 줄 뒤에 맞는 예문과 틀린 예문을 함께 적기", "틀린 문장을 고치고 이유를 한 문장으로 남기기"),
    TopicSignal("reading", "독해 근거", ("독해", "지문", "주제", "근거"), "답을 고른 문장과 선택지의 표현을 연결하는지", "정답 근거에 밑줄을 긋고 선택지를 지운 이유를 적은 기록", "문단별 핵심 문장과 연결어를 표시하며 읽기", "짧은 지문 한 편에서 근거 문장만 다시 찾기"),
    TopicSignal("listening", "듣기 확인", ("듣기", "음원", "발음"), "놓친 소리와 뜻을 모르는 표현을 구분했는지", "받아쓰기와 다시 들은 뒤 수정한 부분", "짧은 구간을 듣고 핵심 표현을 따라 말하기", "한 문장을 듣고 의미 단위로 끊어 적기"),
    TopicSignal("writing", "서술형 쓰기", ("서술형", "쓰기", "영작", "작문"), "조건에 맞는 문장을 스스로 구성하고 검토하는지", "서술형 답안에서 빠진 조건과 고친 문장", "핵심 표현을 넣어 짧은 답안을 완성하기", "답안을 다시 써 보고 빠진 조건을 표시하기"),
    TopicSignal("error", "영어 오답 유형", ("오답", "틀린", "실수", "재시험"), "어휘·문법·독해 중 어디에서 같은 실수가 이어지는지", "첫 풀이와 다시 푼 답 사이에 달라진 근거", "오답을 유형별로 나누고 다시 볼 날짜 정하기", "같은 유형 한 문제를 며칠 뒤 다시 풀기"),
    TopicSignal("pace", "시험 시간 배분", ("시간", "시험", "속도", "시간배분"), "어느 지문과 문항에서 시간이 길어지는지", "문항별 소요 시간과 끝까지 풀지 못한 구간", "읽기와 답 확인 시간을 나누어 연습하기", "짧은 세트를 정해진 시간 안에 풀고 기록하기"),
    TopicSignal("routine", "영어 학습 루틴", ("습관", "플래너", "기록", "과제"), "계획한 분량과 실제 완료한 분량이 구분되는지", "학습 시작 시각과 완료 여부를 적은 주간 기록", "매일 할 최소 분량과 다시 볼 항목을 나누기", "완료한 분량과 남은 질문을 짧게 적기"),
)

ELEMENTARY_ENGLISH_SIGNALS: tuple[TopicSignal, ...] = (
    TopicSignal("phonics", "소리와 철자 연결", ("파닉스", "소리", "철자", "발음"), "알파벳 소리와 낱말 철자를 연결해 읽는지", "처음 본 낱말을 소리 내어 읽고 고친 흔적", "소리 단위로 끊어 읽고 철자를 함께 적기", "짧은 낱말을 보고 소리와 뜻을 말해 보기"),
    TopicSignal("basic_words", "기초 어휘", ("어휘", "단어", "뜻"), "배운 낱말을 그림과 짧은 문장에서 알아보는지", "교재에서 반복해 헷갈린 낱말과 다시 읽은 기록", "낱말·그림·짧은 예문을 한 묶음으로 익히기", "낱말 세 개로 짧은 문장을 만들어 보기"),
    TopicSignal("sentence_reading", "문장 읽기", ("문장", "읽기", "해석"), "짧은 문장의 주어와 행동을 구분해 의미를 말하는지", "멈춰 읽은 문장과 뜻을 다시 설명한 기록", "짧은 문장을 의미 단위로 끊어 읽기", "한 문장을 소리 내어 읽고 우리말로 설명하기"),
    TopicSignal("basic_grammar", "기초 문장 규칙", ("문법", "시제", "동사", "복수"), "배운 문장 규칙을 예문에 맞게 적용하는지", "규칙은 알지만 문장에서 틀린 부분을 고친 기록", "맞는 문장과 틀린 문장을 비교하며 규칙 찾기", "짧은 문장 하나를 바르게 고쳐 쓰기"),
    TopicSignal("reading_flow", "짧은 글 이해", ("독해", "지문", "내용", "이야기"), "짧은 글의 인물·장소·행동을 순서대로 말하는지", "질문에 답한 근거가 있는 문장을 찾은 흔적", "문단마다 중요한 낱말과 문장을 표시하기", "읽은 내용을 두 문장으로 설명해 보기"),
    *ENGLISH_SIGNALS[4:],
)

MATH_SIGNALS: tuple[TopicSignal, ...] = (
    TopicSignal("concept", "개념 연결", ("개념", "정의", "원리", "이해"), "정의와 공식을 말로 설명한 뒤 문제에 적용하는지", "기본 문제와 조건이 달라진 문제의 풀이 근거", "개념 한 줄과 대표 문제를 한 묶음으로 정리하기", "해설 없이 개념을 설명하고 한 문제를 다시 풀기"),
    TopicSignal("calculation", "계산 정확도", ("계산", "연산", "부호", "실수"), "부호·괄호·연산 순서에서 같은 실수가 반복되는지", "중간 계산과 답을 고친 뒤의 재풀이 기록", "계산 단계를 한 줄씩 나누고 틀린 위치 표시하기", "짧은 계산 세트를 정확도 기준으로 다시 풀기"),
    TopicSignal("condition", "문제 조건 해석", ("조건", "문장제", "응용", "문제해결"), "문장에서 필요한 조건을 골라 식이나 그림으로 바꾸는지", "조건을 빠뜨린 문제와 식을 잘못 세운 문제의 구분", "조건에 밑줄을 긋고 구하려는 값을 먼저 적기", "비슷한 문제에서 식을 세우는 과정만 다시 연습하기"),
    TopicSignal("error", "수학 오답 재학습", ("오답", "틀린", "재풀이", "복습"), "틀린 이유를 개념·조건·계산으로 나누어 기록하는지", "첫 풀이와 다시 푼 풀이에서 달라진 단계", "오답 원인을 표시하고 다시 볼 날짜 정하기", "같은 유형 한 문제를 며칠 뒤 다시 풀기"),
    TopicSignal("written", "서술형 풀이", ("서술형", "풀이과정", "과정", "설명"), "답뿐 아니라 식을 세운 이유와 풀이 순서를 적는지", "서술형 답안에서 생략된 조건과 고친 표현", "풀이 단계마다 근거를 짧은 문장으로 남기기", "완성한 풀이를 소리 내어 설명해 보기"),
    TopicSignal("graph", "그래프와 자료 해석", ("그래프", "도형", "표", "좌표"), "그림·표·그래프의 정보를 식과 연결하는지", "자료를 잘못 읽은 부분과 다시 표시한 값", "주어진 정보를 그림에 옮기고 관계를 설명하기", "표나 그림을 보고 조건을 한 문장으로 적기"),
    TopicSignal("units", "단원 간 연결", ("단원", "누적", "선행", "연결"), "앞 단원의 개념이 현재 문제에서 어떻게 쓰이는지 설명하는지", "최근 문제에서 다시 필요해진 이전 단원 기록", "현재 단원과 연결된 앞 개념을 짧게 복습하기", "오늘 문제에 쓰인 이전 개념을 한 줄로 적기"),
    TopicSignal("pace", "수학 시간 배분", ("시간", "시험", "속도", "시간배분"), "어느 유형과 계산 단계에서 시간이 길어지는지", "문항별 소요 시간과 끝까지 풀지 못한 구간", "풀이와 검산 시간을 나누어 연습하기", "짧은 세트를 풀고 오래 걸린 문제를 표시하기"),
    TopicSignal("routine", "수학 학습 루틴", ("습관", "플래너", "기록", "과제"), "계획한 문제와 실제 완료한 문제를 구분하는지", "학습 시작 시각과 재풀이 완료 여부를 적은 기록", "매일 풀 분량과 다시 볼 문제를 나누기", "완료한 문제와 남은 질문을 짧게 적기"),
)

ELEMENTARY_MATH_SIGNALS: tuple[TopicSignal, ...] = (
    TopicSignal("number", "수 감각과 연산 원리", ("연산", "계산", "수", "받아올림", "받아내림"), "계산 절차를 외우기보다 수가 바뀌는 이유를 설명하는지", "계산 과정에서 자릿값과 부호를 표시한 흔적", "수 모형과 식을 연결한 뒤 같은 계산을 다시 풀기", "짧은 연산을 풀고 틀린 단계만 말로 설명하기"),
    TopicSignal("word_problem", "문장제 조건", ("문장제", "문제", "조건", "응용"), "문장에서 필요한 수와 질문을 구분해 표시하는지", "식을 세우기 전 그린 그림과 밑줄 친 조건", "질문을 한 문장으로 바꾸고 필요한 수를 고르기", "생활 장면 문제 하나를 그림과 식으로 나타내기"),
    TopicSignal("fraction", "분수·소수 이해", ("분수", "소수", "비율"), "전체와 부분의 관계를 그림과 수로 함께 설명하는지", "분수나 소수를 수직선과 그림에 표시한 기록", "그림·말·식 세 가지로 같은 값을 나타내기", "생활 속 양을 분수나 소수로 표현해 보기"),
    TopicSignal("geometry", "도형과 측정", ("도형", "각도", "넓이", "길이", "측정"), "도형의 성질과 측정 단위를 구분해 사용하는지", "도형에 표시한 길이·각도와 풀이 식", "그림에 조건을 옮기고 사용한 성질을 적기", "주변 물건의 길이와 모양을 직접 설명해 보기"),
    TopicSignal("explain", "수학 개념 설명", ("개념", "설명", "이해", "원리"), "답을 구한 이유를 자기 말로 설명할 수 있는지", "풀이 뒤에 이유를 한 문장으로 적은 흔적", "대표 문제를 풀고 사용한 개념을 말로 정리하기", "가족에게 풀이 순서를 짧게 설명해 보기"),
    *MATH_SIGNALS[3:],
)


def esc(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", html.unescape(str(value or ""))).strip()


def final_consonant_index(value: str) -> int | None:
    """Return the Korean jongseong index for the final pronounced character."""
    digit_jongseong = {
        "0": 21,  # 영
        "1": 8,   # 일
        "2": 0,   # 이
        "3": 16,  # 삼
        "4": 0,   # 사
        "5": 0,   # 오
        "6": 1,   # 육
        "7": 8,   # 칠
        "8": 8,   # 팔
        "9": 0,   # 구
    }
    for character in reversed(clean(value)):
        codepoint = ord(character)
        if 0xAC00 <= codepoint <= 0xD7A3:
            return (codepoint - 0xAC00) % 28
        if character in digit_jongseong:
            return digit_jongseong[character]
        if character.isalnum():
            return None
    return None


def particle_for(value: str, consonant_form: str, vowel_form: str) -> str:
    jongseong = final_consonant_index(value)
    if consonant_form == "으로" and jongseong == 8:
        return vowel_form
    return consonant_form if jongseong else vowel_form


def normalize_particle_joins(text: str, tokens: Iterable[str]) -> str:
    """Correct particles attached to page-specific generated nouns/phrases."""
    pairs = (("으로", "로"), ("은", "는"), ("이", "가"), ("을", "를"), ("과", "와"))
    values = sorted({clean(token) for token in tokens if clean(token)}, key=len, reverse=True)
    for value in values:
        escaped = re.escape(value)
        for consonant_form, vowel_form in pairs:
            expected = particle_for(value, consonant_form, vowel_form)
            text = re.sub(
                rf"{escaped}(?:{re.escape(consonant_form)}|{re.escape(vowel_form)})(?![가-힣])",
                lambda _match, replacement=value + expected: replacement,
                text,
            )
    return text


def page_particle_tokens(
    config: CategoryConfig,
    center: dict[str, object],
    signals: tuple[TopicSignal, ...],
    student: str = "",
) -> tuple[str, ...]:
    schools = relevant_schools(config, center)
    school_names = "·".join(schools[:3]) if schools else "실제 재학 학교"
    material = material_label(config, schools)
    grade_values = [
        f"{subject} {'·'.join(relevant_grades(config, center, subject)) or '상담 시 확인'}"
        for subject in config.subjects
    ]
    values: list[str] = [
        str(center["address"]),
        str(center["center_name"]),
        config.label,
        config.level,
        config.school_name,
        f"{center['locality']} {config.label}",
        "·".join(config.subjects),
        material,
        school_names,
        f"{school_names} 자료",
        ", ".join(grade_values),
        " / ".join(grade_values),
        *grade_values,
    ]
    if student:
        values.append(student)
    for signal in signals:
        values.extend((signal.label, signal.check, signal.evidence, signal.practice, signal.home_action))
    return tuple(dict.fromkeys(value for value in values if value))


def normalize_generated_value(value: object, tokens: tuple[str, ...]) -> object:
    if isinstance(value, str):
        return normalize_particle_joins(value, tokens)
    if isinstance(value, list):
        return [normalize_generated_value(item, tokens) for item in value]
    if isinstance(value, tuple):
        return tuple(normalize_generated_value(item, tokens) for item in value)
    if isinstance(value, dict):
        return {key: normalize_generated_value(item, tokens) for key, item in value.items()}
    return value


def stable_int(seed: str, label: str) -> int:
    return int(hashlib.sha256(f"{seed}|{label}".encode("utf-8")).hexdigest()[:16], 16)


def stable_pick(seed: str, label: str, choices: tuple[str, ...] | list[str]) -> str:
    return choices[stable_int(seed, label) % len(choices)]


def list_values(value: object) -> list[str]:
    return list(dict.fromkeys(clean(part) for part in re.split(r"[,/\n]+", str(value or "")) if clean(part)))


def normalize_school_values(value: object) -> list[str]:
    text = clean(value)
    if "지역내 모든 고등학교 가능" in text:
        return []
    text = re.sub(r"[.·]+", ",", text)
    known_joined = {"오현초호매실중": ("오현초", "호매실중")}
    values: list[str] = []
    for chunk in re.split(r"[,/\s]+", text):
        chunk = clean(chunk)
        if not chunk:
            continue
        if chunk in known_joined:
            values.extend(known_joined[chunk])
            continue
        values.append(chunk)
    return list(dict.fromkeys(value for value in values if len(value) >= 2))


def folder_slug(locality: str) -> str:
    return re.sub(r"\s+", "", locality)


def display_region_label(center: dict[str, object]) -> str:
    """Avoid repeating a city prefix already present in region/district."""
    region = str(center["region"])
    district = str(center["district"])
    locality = str(center["locality"])
    district_base = re.sub(r"(?:특별자치시|특별시|광역시|시|군|구)$", "", district)
    display_locality = locality
    for prefix in (region, district_base):
        if prefix and display_locality.startswith(prefix + " "):
            display_locality = display_locality[len(prefix):].strip()
            break
    return " ".join(part for part in (region, district, display_locality) if part)


def route_for(category_slug: str, locality_slug: str | None = None) -> str:
    path = f"/과목별학원/{category_slug}/"
    return path + (f"{locality_slug}/" if locality_slug else "")


def absolute_route(category_slug: str, locality_slug: str | None = None) -> str:
    return ORIGIN + quote(route_for(category_slug, locality_slug), safe="/%:@")


def center_entity_id(center: dict[str, object]) -> str:
    key = "|".join(str(center.get(name, "")) for name in ("center_name", "address", "registration"))
    return ORIGIN + "/#center-" + hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]


def source_text(value: str) -> str:
    value = re.sub(r"(?is)<(?:script|style).*?>.*?</(?:script|style)>", " ", value)
    return clean(re.sub(r"<[^>]+>", " ", value))


def physical_region(address: str, fallback: str) -> str:
    first = clean(address).split(" ", 1)[0]
    return first or fallback


def safe_district(region: str, district: str, address: str) -> str:
    if region == "세종" or address.startswith("세종"):
        return "세종시"
    if district.endswith(("로", "길")):
        parts = address.split()
        return parts[1] if len(parts) > 1 and parts[1].endswith(("시", "군", "구")) else region
    return district


def load_source_rows(config: CategoryConfig) -> list[str]:
    if not config.source_path.is_file():
        raise FileNotFoundError(config.source_path)
    workbook = load_workbook(config.source_path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    # Google Sheets exports can retain formatting far below the 371 content
    # rows. Read the contractual range directly so styled empty rows do not
    # turn a small workbook into a million-row scan.
    values = [
        str(row[0]).strip()
        for row in sheet.iter_rows(
            min_row=1,
            max_row=EXPECTED_ROWS,
            min_col=1,
            max_col=1,
            values_only=True,
        )
        if row and isinstance(row[0], str) and row[0].strip()
    ]
    workbook.close()
    if len(values) != EXPECTED_ROWS:
        raise ValueError(f"{config.workbook}: expected {EXPECTED_ROWS}, found {len(values)}")
    return values


def load_centers() -> list[dict[str, object]]:
    with IMAGE_CSV.open(encoding="utf-8-sig", newline="") as handle:
        image_rows = list(csv.DictReader(handle))
    with CENTER_CSV.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    if len(rows) != EXPECTED_ROWS or len(image_rows) != EXPECTED_ROWS:
        raise ValueError(f"common rows center={len(rows)} image={len(image_rows)}")
    image_by_locality = {clean(row["제목"]): row for row in image_rows}
    centers: list[dict[str, object]] = []
    seen_localities: set[str] = set()
    seen_slugs: set[str] = set()
    for row in rows:
        locality = clean(row["근처 수업가능 동네"])
        slug = folder_slug(locality)
        if locality in seen_localities or slug in seen_slugs:
            raise ValueError(f"duplicate locality/slug: {locality}/{slug}")
        seen_localities.add(locality)
        seen_slugs.add(slug)
        english_slug = clean(row["동 영어"]).replace(" ", "-")
        maps = sorted(MAP_DIR.glob(english_slug + ".*"))
        if len(maps) != 1:
            raise FileNotFoundError(f"map {locality}: {english_slug} -> {len(maps)}")
        image = image_by_locality.get(locality)
        if not image:
            raise ValueError(f"image mapping missing: {locality}")
        address = clean(row["센터 주소"])
        region = clean(row["지역"])
        district = safe_district(region, clean(row["시or구"]), address)
        centers.append({
            "locality": locality,
            "slug": slug,
            "english_slug": english_slug,
            "region": region,
            "district": district,
            "address_region": physical_region(address, region),
            "center_name": clean(row["센터명"]),
            "tuition_url": clean(row["센터 교습비"]),
            "office_name": clean(row["교육지원청명칭"]),
            "registration": clean(row["교육지원청 등록번호"]),
            "address": address,
            "schools": {
                "타깃학교\n(초)": normalize_school_values(row["타깃학교\n(초)"]),
                "타깃학교\n(중)": normalize_school_values(row["타깃학교\n(중)"]),
                "타깃학교\n(고)": normalize_school_values(row["타깃학교\n(고)"]),
            },
            "grades": {
                "국어": list_values(row["가능학년\n(국어)"]),
                "영어": list_values(row["가능학년\n(영어)"]),
                "수학": list_values(row["가능학년\n(수학)"]),
            },
            "map_name": maps[0].name,
            "body_image": "seoul6839.webp" if clean(image["본문"]).lower() == "seoul.jpg" else "local6839.webp",
        })
    return centers


def relevant_grades(config: CategoryConfig, center: dict[str, object], subject: str) -> list[str]:
    values = list(center["grades"][subject])  # type: ignore[index]
    if not config.grade_prefix:
        return values
    return [value for value in values if value.startswith(config.grade_prefix)]


def relevant_schools(config: CategoryConfig, center: dict[str, object]) -> list[str]:
    result: list[str] = []
    for column in config.school_columns:
        result.extend(center["schools"][column])  # type: ignore[index]
    return list(dict.fromkeys(result))


def signal_bank(config: CategoryConfig, subject: str) -> tuple[TopicSignal, ...]:
    if subject == "영어":
        return ELEMENTARY_ENGLISH_SIGNALS if config.is_elementary else ENGLISH_SIGNALS
    return ELEMENTARY_MATH_SIGNALS if config.is_elementary else MATH_SIGNALS


def rank_signals(config: CategoryConfig, raw: str, seed: str) -> tuple[TopicSignal, ...]:
    text = source_text(raw)
    ranked_by_subject: dict[str, list[TopicSignal]] = {}
    for subject in config.subjects:
        bank = signal_bank(config, subject)
        scored = sorted(
            bank,
            key=lambda signal: (
                -sum(text.count(keyword) for keyword in signal.keywords),
                stable_int(seed, signal.code),
            ),
        )
        ranked_by_subject[subject] = scored
    if len(config.subjects) == 2:
        english = ranked_by_subject["영어"]
        math = ranked_by_subject["수학"]
        # English and math banks intentionally share a few generic actions
        # (for example, retrying the same question or timing a short set).
        # Avoid selecting an identical action/evidence pair for a combined
        # page because it would ask families to compare the same thing twice.
        english_primary = english[0]
        critical_fields = ("evidence", "practice", "home_action")
        math_primary = next(
            (
                candidate
                for candidate in math
                if all(
                    getattr(candidate, field) != getattr(english_primary, field)
                    for field in critical_fields
                )
            ),
            math[0],
        )
        math_secondary = next(candidate for candidate in math if candidate != math_primary)
        return (english_primary, math_primary, english[1], math_secondary)
    return tuple(ranked_by_subject[config.subjects[0]][:3])


def representative_pool() -> list[Path]:
    unique: list[Path] = []
    seen: set[str] = set()
    for path in sorted(REP_SOURCE.iterdir(), key=lambda item: item.name):
        if not path.is_file() or path.suffix.lower() not in {".gif", ".jpg", ".jpeg", ".png", ".webp"}:
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest in seen:
            continue
        seen.add(digest)
        unique.append(path)
    if len(unique) < EXPECTED_ROWS:
        raise ValueError(f"representative pool unique={len(unique)}")
    return unique


def existing_row_digests() -> list[set[str]]:
    rows = [set() for _ in range(EXPECTED_ROWS)]
    for prefix in ("hs-kem", "ms-kem", "es-kem"):
        for index in range(EXPECTED_ROWS):
            matches = sorted(REP_TARGET.glob(f"{prefix}-{index + 1:03d}.*"))
            if len(matches) == 1:
                rows[index].add(hashlib.sha256(matches[0].read_bytes()).hexdigest())
    return rows


def assign_representatives(configs: Iterable[CategoryConfig]) -> dict[str, list[str]]:
    REP_TARGET.mkdir(parents=True, exist_ok=True)
    pool = representative_pool()
    pool_digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in pool]
    row_digests = existing_row_digests()
    assignments: dict[str, list[str]] = {}
    for config in configs:
        existing: list[Path] = []
        for index in range(EXPECTED_ROWS):
            matches = sorted(REP_TARGET.glob(f"{config.rep_prefix}-{index + 1:03d}.*"))
            if len(matches) != 1:
                existing = []
                break
            existing.append(matches[0])
        if existing:
            names = [path.name for path in existing]
            digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in existing]
            if len(set(digests)) != EXPECTED_ROWS:
                raise ValueError(f"{config.slug}: existing representative duplicates")
            for index, digest in enumerate(digests):
                if digest in row_digests[index]:
                    raise ValueError(f"{config.slug}: same-locality representative collision at {index + 1}")
                row_digests[index].add(digest)
            assignments[config.slug] = names
            continue
        used: set[str] = set()
        names: list[str] = []
        for index in range(EXPECTED_ROWS):
            start = stable_int(config.slug, f"rep-{index}") % len(pool)
            selected = None
            for step in range(len(pool)):
                position = (start + step) % len(pool)
                digest = pool_digests[position]
                if digest not in used and digest not in row_digests[index]:
                    selected = pool[position]
                    used.add(digest)
                    row_digests[index].add(digest)
                    break
            if selected is None:
                raise ValueError(f"{config.slug}: representative allocation failed at {index + 1}")
            name = f"{config.rep_prefix}-{index + 1:03d}{selected.suffix.lower()}"
            shutil.copy2(selected, REP_TARGET / name)
            names.append(name)
        assignments[config.slug] = names
    return assignments


def material_label(config: CategoryConfig, schools: list[str]) -> str:
    if config.is_elementary:
        return "현재 교재·학교 진도·단원평가 자료" if schools else "현재 교재·학습지·단원 기록"
    if config.grade_prefix in {"중", "고"}:
        return "최근 시험지·시험 범위표·현재 교재" if schools else "최근 시험지·현재 교재·오답 기록"
    return "최근 영어·수학 교재·학교 자료·오답 기록"


def grade_summary(config: CategoryConfig, center: dict[str, object]) -> list[tuple[str, str]]:
    return [
        (subject, "·".join(relevant_grades(config, center, subject)) or "상담 시 확인")
        for subject in config.subjects
    ]


def student_level_label(config: CategoryConfig) -> str:
    """Use standard compound labels while keeping the multi-level label readable."""
    return f"{config.level}학생" if config.grade_prefix else f"{config.level} 학생"


def student_type(config: CategoryConfig, signals: tuple[TopicSignal, ...], seed: str, locality: str) -> str:
    primary, secondary = signals[:2]
    subject_text = "·".join(config.subjects)
    student_level = student_level_label(config)
    situation = stable_pick(seed, "student-type", [
        f"{primary.label}에서 겪는 어려움과 {secondary.label}에서 막히는 원인을 구분해 {subject_text} 학습 순서를 정해야 하는 {student_level}",
        f"현재 교재의 진도는 나가지만 {primary.label} 상태를 스스로 점검하거나 {secondary.label} 관련 기록을 남기기 어려운 {student_level}",
        f"과제량보다 {primary.label}과 {secondary.label}을 실제 자료로 확인한 뒤 주간 계획을 세워야 하는 {student_level}",
        f"학습 결과만 보지 않고 {primary.label}이 드러난 자료와 {secondary.label}을 다시 살필 시점을 함께 점검해야 하는 {student_level}",
        f"혼자 공부할 때 {primary.label}에서 멈추고 {secondary.label} 항목을 미루는 경향을 함께 점검할 필요가 있는 {student_level}",
        f"학교 진도와 현재 수준 사이에서 {primary.label} 및 {secondary.label}의 우선순위를 정하기 어려운 {student_level}",
    ])
    return f"{locality}에서 {config.level} {subject_text} 상담을 준비하며 {situation}"


def meta_description(config: CategoryConfig, center: dict[str, object], signals: tuple[TopicSignal, ...]) -> str:
    locality = str(center["locality"])
    title = f"{locality} {config.label}"
    candidates = (
        clean(
            f"{title} 선택 전 {signals[0].label}·{signals[1].label} 진단, {config.school_name} 자료, "
            "가능 학년과 센터 위치를 확인하세요. 실제 교재로 상담할 질문도 정리했습니다."
        ),
        clean(
            f"{title} 선택 전 {signals[0].label}·{signals[1].label} 진단과 {config.school_name} 자료, "
            "가능 학년, 센터 위치 및 상담 질문을 확인하세요."
        ),
        clean(
            f"{title}의 {signals[0].label}·{signals[1].label} 진단, 가능 학년, 센터 위치와 "
            "상담 전 질문을 실제 학생 자료 기준으로 정리했습니다."
        ),
    )
    for value in candidates:
        if 70 <= len(value) <= 100:
            return value
    eligible = [value for value in candidates if len(value) <= 100]
    if not eligible:
        raise ValueError(f"meta description too long: {title}")
    value = max(eligible, key=len)
    addition = " 실제 학생 자료를 기준으로 확인합니다."
    if len(value) + len(addition) <= 100:
        value += addition
    if len(value) < 70:
        raise ValueError(f"meta description too short: {title} ({len(value)})")
    return value


def quick_answer(config: CategoryConfig, center: dict[str, object], signals: tuple[TopicSignal, ...], seed: str) -> str:
    locality = str(center["locality"])
    schools = relevant_schools(config, center)
    material = material_label(config, schools)
    primary, secondary = signals[:2]
    first = stable_pick(seed, "quick-first", [
        f"{locality} {config.label}을 비교할 때는 진도보다 {primary.label}과 {secondary.label}이 실제 학습에서 어떻게 드러나는지 먼저 확인해야 합니다.",
        f"{locality}에서 {config.label}을 찾는다면 현재 점수만 보지 말고 {primary.label}이 드러난 자료와 {secondary.label} 재확인 과정을 나누어 살펴보세요.",
        f"{locality} {config.level} 학습의 출발점은 분량을 늘리는 일이 아니라 {primary.label}과 {secondary.label}의 상태를 자료로 구분하는 것입니다.",
        f"좋은 {locality} {config.label} 상담은 {primary.label}에서 막히는 원인과 {secondary.label}에 필요한 연습을 같은 계획 안에서 구분하는 데서 시작합니다.",
        f"{locality} {config.label} 선택 전에는 학생이 {primary.label}에서 멈추는 장면과 {secondary.label}을 다시 확인하는 방식을 먼저 기록해야 합니다.",
        f"{locality}의 {config.label} 안내를 볼 때는 광고 문구보다 {primary.label}·{secondary.label}을 어떤 자료로 판단하는지 확인하는 편이 정확합니다.",
    ])
    school_note = f"{schools[0]} 등 공개된 참고 학교명은" if schools else f"{config.school_name} 정보가 비어 있는 이 지역은"
    second = stable_pick(seed, "quick-second", [
        f"{material}를 준비하면 {center['center_name']} 상담에서 현재 수준과 다음 점검일을 구체적으로 비교할 수 있습니다.",
        f"{school_note} 수업 가능 여부를 보장하지 않으므로 실제 학생 자료와 센터의 현재 개설 범위를 함께 확인해야 합니다.",
        f"제공 주소인 {center['address']}와 과목별 가능 학년, 통학 시간은 등록 전에 다시 확인하고 학습 계획과 분리해 판단하세요.",
        f"상담에서는 {material}에 표시된 어려운 부분을 바탕으로 첫 학습 행동과 재확인 기준을 정하는 것이 좋습니다.",
        f"학생이 사용 중인 자료를 가져가면 {primary.evidence}와 {secondary.evidence}, 두 자료를 바탕으로 우선순위를 질문할 수 있습니다.",
        f"센터 정보와 학습 자료를 함께 보면 확인된 사실, 상담에서 물을 조건과 학생의 학습 과제를 구분할 수 있습니다.",
    ])
    return first + " " + second


def answer_cards(config: CategoryConfig, center: dict[str, object], signals: tuple[TopicSignal, ...], student: str) -> list[tuple[str, str]]:
    locality = str(center["locality"])
    grade_items = grade_summary(config, center)
    listed_items = [(subject, value) for subject, value in grade_items if value != "상담 시 확인"]
    missing_subjects = [subject for subject, value in grade_items if value == "상담 시 확인"]
    if missing_subjects:
        if listed_items:
            listed_text = " / ".join(f"{subject} {value}" for subject, value in listed_items)
            grade_answer = (
                f"공개 자료에는 {listed_text} 범위가 기재되어 있고, {'·'.join(missing_subjects)} 가능 학년은 상담에서 확인합니다. "
                "실제 시간표도 함께 확인합니다."
            )
        else:
            grade_answer = (
                f"공개 자료에는 {'·'.join(missing_subjects)} 가능 학년이 기재되지 않아 상담에서 확인해야 합니다. "
                "실제 시간표도 함께 확인합니다."
            )
    else:
        grade_text = " / ".join(f"{subject} {value}" for subject, value in grade_items)
        grade_answer = f"공개된 가능 학년은 {grade_text}이며 실제 시간표는 상담에서 확인합니다."
    return [
        ("01 / 지역", f"{display_region_label(center)}의 제공 센터·주소 자료를 확인합니다."),
        ("02 / 학년", grade_answer),
        ("03 / 추천 학생", student),
    ]


def build_sections(config: CategoryConfig, center: dict[str, object], signals: tuple[TopicSignal, ...], seed: str) -> list[dict[str, object]]:
    locality = str(center["locality"])
    title = f"{locality} {config.label}"
    primary, secondary, support = signals[:3]
    extra = signals[3] if len(signals) > 3 else primary
    schools = relevant_schools(config, center)
    material = material_label(config, schools)
    school_text = "·".join(schools[:3]) if schools else "실제 재학 학교"
    school_materials = f"{school_text} 자료"
    school_reference = (
        f"{title} 페이지에 표시된 {school_text}은 상담 준비를 위한 참고 정보입니다. 학교명만으로 수업 가능 여부를 판단하지 말고 학생의 {material}를 센터의 현재 개설 범위와 함께 확인해야 합니다."
        if schools else
        f"{title} 페이지에는 참고 학교 목록이 제공되지 않았습니다. 실제 재학 학교와 {material}를 상담에서 알려 주고 센터의 현재 개설 범위와 함께 확인해야 합니다."
    )
    subject_grade_pairs = [
        (subject, relevant_grades(config, center, subject))
        for subject in config.subjects
    ]
    grades = ", ".join(
        f"{subject} {'·'.join(subject_grades) or '상담 시 확인'}"
        for subject, subject_grades in subject_grade_pairs
    )
    has_listed_grades = any(subject_grades for _, subject_grades in subject_grade_pairs)
    all_grades_listed = all(subject_grades for _, subject_grades in subject_grade_pairs)
    subject_text = "·".join(config.subjects)
    student_level = student_level_label(config)
    student = student_type(config, signals, seed, locality)

    diagnosis = {
        "key": "diagnosis",
        "heading": stable_pick(seed, "h2-diagnosis", [
            f"{title} 진단은 어디서 시작할까",
            f"{locality} {config.level} {subject_text} 학습의 출발점을 찾는 방법",
            f"{title} 상담 전에 확인할 첫 근거",
            f"{locality}에서 {config.label}을 비교하는 진단 기준",
        ]),
        "paragraphs": [
            stable_pick(seed, "diagnosis-p1", [
                f"{title}의 첫 진단에서는 {primary.label}과 {secondary.label}을 한 가지 문제로 묶지 않습니다. {primary.check}와 {secondary.check}를 각각 확인해야 보완 순서가 선명해집니다.",
                f"{locality} {student_level}의 현재 상태는 점수 한 줄보다 풀이와 읽기 과정에서 더 잘 드러납니다. {primary.evidence}와 {secondary.evidence}를 나란히 보면 반복되는 어려움을 구분할 수 있습니다.",
                f"{title} 상담은 학생이 막힌 장면을 구체적으로 설명하는 데서 시작합니다. {primary.label}은 현재 자료에서 확인하고 {secondary.label}은 다시 수행한 기록으로 비교하는 편이 좋습니다.",
                f"{locality}에서 {config.label}을 알아볼 때 진도표만 보면 시작점을 놓칠 수 있습니다. 먼저 {primary.label}이 드러난 자료와 {secondary.label} 실행 여부를 별도의 항목으로 표시하세요.",
                f"{title}의 진단 자료는 많을 필요가 없습니다. 최근 기록에서 {primary.label}과 {secondary.label}이 드러난 부분을 골라 오면 상담 질문을 구체화할 수 있습니다.",
                f"{locality} {config.level} 과정의 우선순위는 잘한 단원보다 멈춘 과정에서 찾습니다. {primary.check}를 살핀 뒤 {secondary.check}를 이어 확인하는 순서가 알맞습니다.",
            ]),
            stable_pick(seed, "diagnosis-p2", [
                f"첫 자료에서는 ‘{primary.evidence}’부터 살펴보고, 별도 기록에서는 ‘{support.evidence}’도 확인하세요. 두 자료의 차이가 {title} 상담에서 첫 학습 행동을 정하는 근거가 됩니다.",
                f"학생에게 답을 다시 외우게 하기보다 ‘{primary.practice}’와 ‘{secondary.practice}’ 두 활동을 짧게 수행하게 해 보세요. 결과가 아니라 과정의 변화를 기록해야 다음 점검 기준을 정할 수 있습니다.",
                f"{material}에서 어려운 부분을 한두 곳만 골라도 충분합니다. 상담에서는 {primary.label} 관련 안내와 {support.label}을 점검할 연습량을 구분해 질문하세요.",
                f"같은 오답이라도 {primary.label}의 문제인지 {secondary.label}의 문제인지에 따라 다음 과제가 달라집니다. {locality} 페이지의 센터 정보와 학생 기록을 함께 놓고 판단하세요.",
                f"진단 뒤에는 ‘{support.home_action}’를 첫 점검 행동으로 정할 수 있습니다. 이 행동이 현재 통학과 과제 일정 안에서 가능한지도 {center['center_name']}에 확인하세요.",
                f"학부모는 학생 대신 원인을 단정하지 말고 {primary.check}를 질문으로 바꾸어 적어 두는 편이 좋습니다. 학생의 설명과 실제 자료가 일치하는지를 상담에서 비교할 수 있습니다.",
            ]),
        ],
    }

    method = {
        "key": "method",
        "heading": stable_pick(seed, "h2-method", [
            f"{primary.label}과 {secondary.label}을 나누는 {locality} {config.level} {subject_text} 기준",
            f"{title}에서 과목별 병목을 구분하는 방법",
            f"{locality} {student_level}의 {subject_text} 학습 과정을 확인하는 순서",
            f"{title} 학습 기록을 다음 수업으로 연결하는 법",
        ]),
        "paragraphs": [
            stable_pick(seed, "method-p1", [
                f"{title} 학습에서는 {primary.label}과 {secondary.label}을 같은 분량으로 다루기보다 막힌 원인에 맞춰 시간을 나누어야 합니다. 먼저 ‘{primary.practice}’를 실행하고 기록이 남으면 ‘{secondary.practice}’로 이어 갑니다.",
                f"{locality} {student_level}에게 필요한 연습은 문제 수보다 확인 가능한 과정입니다. {primary.evidence}를 살핀 뒤 {secondary.evidence}를 비교하면 설명과 반복 연습의 비중을 정할 수 있습니다.",
                f"{title}의 수업 흐름은 진단, 짧은 연습, 재확인의 순서가 분명해야 합니다. {primary.label}을 점검할 때는 ‘{primary.practice}’를, {secondary.label}을 점검할 때는 ‘{secondary.practice}’를 활용합니다.",
                f"{locality}에서 학습 계획을 세울 때는 학생이 이미 아는 내용과 혼자 적용하지 못하는 내용을 나눕니다. {primary.check}를 확인한 다음 {secondary.check}를 별도로 살펴보세요.",
                f"{title} 상담에서 한 번에 많은 과제를 약속할 필요는 없습니다. {primary.label}과 {secondary.label} 중 먼저 바꿀 행동 하나를 고르고 다음 점검일에 실제 기록을 확인하면 됩니다.",
                f"과목 학습을 함께 관리하더라도 {locality} 학생의 병목은 서로 다를 수 있습니다. {primary.evidence}와 {secondary.evidence}를 바탕으로 설명 시간과 독립 연습 시간을 구분하세요.",
            ]),
            stable_pick(seed, "method-p2", [
                f"가정에서는 {primary.home_action}와 {secondary.home_action} 가운데 하나만 정해 실행해도 좋습니다. {title} 상담에서 완료 여부와 어려웠던 부분을 공유하면 다음 과제량을 조정하기 쉽습니다.",
                f"피드백은 정답 수보다 학생이 이유를 설명하고 다시 수행할 수 있는지에 초점을 맞춥니다. {support.label} 기록이 남아야 수업 뒤 복습이 이어졌는지 확인할 수 있습니다.",
                f"첫 주에는 {primary.label}, 다음 점검에서는 {secondary.label}을 살피는 식으로 순서를 정할 수 있습니다. 다만 이 흐름은 학생 자료와 실제 가능 일정에 따라 달라지는 조건부 예시입니다.",
                f"교재를 바꾸기 전에는 현재 자료에서 {support.check}를 먼저 확인하세요. 자료가 어려운 것인지 학습 절차가 빠진 것인지 구분한 뒤 교재와 분량을 결정해야 합니다.",
                f"{center['center_name']}에 문의할 때는 설명 방식, 독립 연습과 피드백 주기를 함께 물어보세요. 실제 수업 시간과 반 편성은 센터의 현재 운영 범위에 따라 달라질 수 있습니다.",
                f"학생이 혼자 다시 해 본 흔적을 남기면 같은 설명을 반복하는 시간을 줄일 수 있습니다. ‘{support.practice}’ 활동을 다음 수업 전에 해낼 수 있는지 현실적으로 확인하세요.",
            ]),
        ],
    }

    school = {
        "key": "school",
        "heading": stable_pick(seed, "h2-school", [
            f"{school_materials}로 살펴보는 {title} 학습 범위",
            f"{locality} {config.school_name} {subject_text} 자료와 현재 교재를 연결하는 법",
            f"{title}에서 학교 자료를 상담 근거로 쓰는 방법",
            f"{locality} 학생의 실제 진도와 {config.label} 계획 비교",
        ]),
        "paragraphs": [
            stable_pick(seed, "school-p1", [
                school_reference,
                f"{locality} 학생의 학교 진도는 같은 학년 안에서도 다를 수 있습니다. {material}에서 {primary.label}과 {secondary.label}이 드러난 부분을 표시해 실제 수업 계획과 대조하세요.",
                f"{school_materials}를 준비할 때는 이름을 나열하기보다 현재 단원과 평가 일정을 확인하는 편이 좋습니다. {title} 상담에서는 그 범위가 주간 과제에 어떻게 반영되는지를 물어보세요.",
                f"학교 자료가 없는 경우에도 {locality} 학생의 현재 교재와 학습 기록으로 시작할 수 있습니다. 실제 학교명과 진도는 상담에서 확인하고 임의로 추정하지 않습니다.",
                f"{title}의 학교 정보는 광고 문구가 아니라 상담 자료를 고르는 기준으로 사용합니다. {school_text}의 학습 자료와 학생이 푼 교재를 함께 보면 보완할 범위를 구체화할 수 있습니다.",
                f"{locality} 학부모는 {material} 가운데 학생이 설명하지 못한 부분을 표시해 가져가면 좋습니다. 학교 자료와 센터 운영 범위가 일치하는지는 등록 전에 별도로 확인해야 합니다.",
            ]),
            stable_pick(seed, "school-p2", [
                (
                    f"공개 자료에 적힌 가능 학년은 {grades}입니다. 실제 시간표가 필요한 경우에는 {center['center_name']}에서 과목·학년·요일을 다시 확인하세요."
                    if all_grades_listed else
                    (
                        f"공개 자료에는 일부 과목의 가능 학년이 기재되지 않았습니다. {center['center_name']}에서 희망 과목·학년·요일과 실제 개설 범위를 확인하세요."
                        if len(config.subjects) > 1 and has_listed_grades else
                        f"제공된 센터 자료만으로는 {subject_text} 가능 학년을 확인할 수 없습니다. {center['center_name']}에서 희망 학년·요일과 실제 개설 범위를 확인하세요."
                    )
                ),
                f"제공된 센터 기준 주소는 {center['address']}입니다. {locality}에서의 실제 통학 시간과 수업 시작 시각은 학생 일정에 맞춰 직접 확인해야 합니다.",
                f"{center['center_name']}의 제공 정보와 학생 학교 자료는 서로 다른 사실입니다. 센터 위치·교습비 경로와 학습 범위·과제 기준을 구분해서 상담 메모에 적어 두세요.",
                (
                    f"학년 표기가 있더라도 반 편성이나 시간표까지 확정된 뜻은 아닙니다. {title} 등록 전에는 표시된 학년 범위와 실제 개설 여부, 수업 방식을 함께 확인하세요."
                    if has_listed_grades else
                    f"공개 자료의 가능 학년 정보가 비어 있으므로 {title} 등록 전에는 희망 학년과 실제 개설 여부, 수업 방식을 함께 확인하세요."
                ),
                (
                    f"표시된 {school_text} 목록은 상담 준비용 참고 자료입니다. 학생의 실제 재학 학교와 현재 진도를 알려 주고 자료 반영 방식을 질문하세요."
                    if schools else
                    "학교명 목록이 비어 있으면 실제 재학 학교를 상담에서 알려 주고 자료 반영 방식을 질문하세요. 공개되지 않은 학교 정보나 시험 범위를 페이지가 임의로 보충하지 않습니다."
                ),
                f"교습비 링크가 제공된 경우에도 금액과 과정은 변경될 수 있습니다. {locality} 상담에서 희망 과목, 학년과 주당 일정을 먼저 말한 뒤 최신 자료를 확인하세요.",
            ]),
        ],
    }

    recommended = {
        "key": "recommended",
        "heading": stable_pick(seed, "h2-recommended", [
            f"{title}이 필요한 학생과 확인할 수업 조건",
            f"{locality}에서 {config.label} 상담이 특히 유용한 경우",
            f"{locality} {student_level}에게 맞는 {'·'.join(config.subjects)} 점검 기준",
            f"{title} 추천 학생을 판단하는 실제 기준",
        ]),
        "paragraphs": [
            stable_pick(seed, "recommended-p1", [
                f"{title}을 우선 살펴볼 학생은 {student}입니다. 이 표현은 등록을 권하는 기준이 아니라 현재 자료와 상담 질문을 정리하기 위한 학습 상황입니다.",
                f"{locality}에서 {config.label} 상담이 유용한 경우는 {student}일 때입니다. 학생이 사용 중인 교재와 실행 기록을 확인해 실제 원인과 일치하는지 먼저 살펴야 합니다.",
                f"{student}이라면 {title} 상담에서 도움을 받을 부분과 혼자 연습할 부분을 구분해 질문해 보세요. 필요한 수업 조건은 학생마다 달라질 수 있습니다.",
                f"{title}의 추천 학생 기준은 성적 구간이 아니라 {primary.label}과 {secondary.label}의 실행 상태입니다. 최근 자료에서 두 항목이 어떻게 나타나는지 확인한 뒤 판단하세요.",
                f"{locality} {student_level}이 과제를 마쳐도 같은 어려움을 반복한다면 원인 분류가 필요합니다. {primary.label}과 {support.label}을 나눠 보면 필요한 피드백을 구체적으로 질문할 수 있습니다.",
                f"{title}을 알아보는 학부모는 학생에게 많은 문제를 먼저 제시하지 말고 {primary.check}를 확인해 보세요. 실제 자료와 학생 설명이 다를 때 상담에서 그 차이를 다루면 됩니다.",
            ]),
            stable_pick(seed, "recommended-p2", [
                f"학부모는 정답을 대신 알려 주기보다 학생이 ‘{support.home_action}’ 활동을 했는지 기록할 수 있습니다. {locality} 상담에서 이 기록을 공유하면 가정과 수업의 역할을 구분하기 쉽습니다.",
                f"학생이 도움을 요청할 시점과 혼자 다시 해 볼 범위를 정해 두는 편이 좋습니다. {center['center_name']}의 피드백 방식이 이 기준과 맞는지 상담에서 확인하세요.",
                f"추천 여부는 한 번의 점수나 설명만으로 결정하지 않습니다. {primary.evidence}와 {secondary.evidence}를 일정 간격으로 비교해 실제 학습 흐름을 살펴보세요.",
                f"과제량이 많은 수업보다 학생이 ‘{support.practice}’ 활동을 해 보고 확인받을 수 있는지가 중요합니다. 실제 운영 주기와 보완 방식은 센터에 직접 질문해야 합니다.",
                f"{locality} 학생의 귀가 시간과 학교 과제를 함께 놓고 실행 가능한 분량을 정하세요. 계획을 지키지 못한 경우에는 의지로 단정하지 말고 시간·난도·방법을 나누어 확인합니다.",
                f"첫 상담에서는 학생이 바꾸고 싶은 행동도 직접 말하게 해 보세요. 학부모 질문, 학생 목표와 실제 자료가 한 방향인지 확인해야 계획이 현실적입니다.",
            ]),
        ],
    }

    plan = {
        "key": "plan",
        "heading": stable_pick(seed, "h2-plan", [
            f"{title} 상담에서 정하는 조건부 4주 계획",
            f"{locality} {config.level} {subject_text} 학습의 첫 점검 주기 설계",
            f"{title} 학습을 진단·연습·재확인으로 나누기",
            f"{locality} {student_level}에게 맞는 {subject_text} 첫 달 실행 순서",
        ]),
        "paragraphs": [
            stable_pick(seed, "plan-p1", [
                f"{title}의 4주 계획은 성과를 약속하는 프로그램이 아니라 상담 내용을 정리하는 조건부 예시입니다. 첫 주에는 {primary.label}, 둘째 주에는 {secondary.label}, 이후에는 {support.label} 기록을 살피고 마지막 점검에서 다음 범위를 조정할 수 있습니다.",
                f"{locality} 학생의 첫 학습 주기는 진단, 짧은 연습과 재확인으로 나눌 수 있습니다. 먼저 ‘{primary.practice}’를 실행하고 이어 ‘{secondary.practice}’도 수행할 수 있는지 확인해 분량을 조정하세요.",
                f"{title} 상담 뒤에는 한꺼번에 여러 행동을 바꾸지 않습니다. 먼저 ‘{primary.home_action}’ 활동을 하고 다음 점검에서 기록을 확인한 뒤 {support.label} 과제를 더할 수 있습니다.",
                f"{locality} {config.level} 학습의 첫 달은 현재 자료를 읽는 기간으로 사용할 수 있습니다. {primary.evidence}, {secondary.evidence}와 실제 완료 분량을 비교해 다음 계획을 정합니다.",
                f"{title}에서 제안받은 계획은 학생 일정에 맞게 줄이거나 순서를 바꿀 수 있어야 합니다. {primary.label}과 {secondary.label} 중 먼저 확인할 항목을 정하고 점검일을 남기세요.",
                f"첫 계획에서는 ‘{support.practice}’ 활동을 수행할 시간을 확보하는 것이 중요합니다. {locality} 학생의 학교 일정과 통학 시간을 반영해 무리하지 않는 주간 분량을 정해야 합니다.",
            ]),
            stable_pick(seed, "plan-p2", [
                f"계획을 확인할 때는 문제 수보다 학생이 이유를 설명하고 다시 수행한 흔적을 봅니다. {title} 상담에서 첫 점검일과 계획을 바꿀 조건을 미리 질문하세요.",
                f"넷째 주에는 점수만 비교하지 않고 {primary.label}과 {support.label}에서 달라진 점이 있는지 살펴봅니다. 학습 결과는 출발점과 실천 정도에 따라 달라질 수 있습니다.",
                f"실행 기록이 남지 않았다면 새 진도를 늘리기 전에 분량과 방법을 다시 조정합니다. {center['center_name']}의 실제 피드백 주기가 학생 일정과 맞는지도 확인하세요.",
                f"한 주의 계획은 학교 일정이 바뀌면 함께 조정되어야 합니다. {material}를 기준으로 집중할 범위와 이어 갈 복습 내용을 나누어 두세요.",
                f"다음 단계는 첫 계획을 모두 마쳤을 때만 정하지 않습니다. 학생이 어려움을 설명한 시점과 다시 수행한 기록을 보고 보완 순서를 바꿀 수 있습니다.",
                f"이 예시는 특정 학생의 성과나 실제 수강 결과가 아닙니다. {locality} 학생의 자료, 가능한 일정과 센터의 현재 운영 조건을 확인한 뒤 개별 계획을 정해야 합니다.",
            ]),
        ],
    }

    checklist_items = [
        ("현재 자료", f"{title} 확인용으로 {material}에서 {primary.label}과 {secondary.label}이 드러난 부분을 표시합니다."),
        ("오답 근거", f"{locality} {config.level} {subject_text} 상담에는 {primary.evidence}와 {support.evidence}를 서로 나누어 가져갑니다."),
        ("가능 일정", f"{locality} {config.level} {subject_text} 학습에 필요한 통학 시간, 학교 일정과 가정 복습 시간을 적습니다."),
        ("센터 확인", f"{title} 등록 전에는 {center['center_name']}의 과목·학년·시간표·교습비·보완 방식을 확인합니다."),
    ]
    checklist = {
        "key": "checklist",
        "heading": stable_pick(seed, "h2-checklist", [
            f"{title} 상담 전 체크리스트",
            f"{locality} {config.label} 방문 전에 준비할 네 가지",
            f"{title} 등록 판단 전에 확인할 질문",
            f"{locality} 학부모가 {config.level} {subject_text} 상담 메모에 남길 항목",
        ]),
        "paragraphs": [
            stable_pick(seed, "checklist-p1", [
                f"{title} 상담 전에는 학생의 실제 자료, 어려웠던 과정과 가능한 일정을 한 장에 정리하세요. 아래 네 항목을 준비하면 학습 문제와 센터 운영 조건을 섞지 않고 질문할 수 있습니다.",
                f"{locality} 학부모는 많은 자료보다 최근 상태를 보여 주는 자료를 고르는 편이 좋습니다. ‘{primary.label}·{secondary.label}’ 관련 질문과 실제 학습 시간 확인 항목으로 나누어 보세요.",
                f"{title} 등록 판단은 수업 설명만 듣고 끝내지 않습니다. 현재 교재와 학생 기록, 가능 학년과 통학 조건을 차례로 확인해야 합니다.",
                f"{locality} 상담 메모에는 학습 질문과 운영 질문을 구분해 적으세요. 학생 자료는 수업 방향을, 주소·시간표·교습비는 실제 이용 조건을 확인하는 데 사용합니다.",
                f"{title}에 문의할 때는 최근 자료와 함께 학생이 혼자 공부할 수 있는 시간도 알려 주세요. 계획의 분량과 피드백 주기를 현실적으로 비교할 수 있습니다.",
                f"상담 준비는 학생의 약점을 많이 나열하는 일이 아닙니다. {locality} 학생이 지금 바꿀 행동, 확인할 자료와 센터에 물을 조건을 네 묶음으로 정리하면 충분합니다.",
            ]),
            stable_pick(seed, "checklist-p2", [
                (
                    f"표시된 가능 학년은 {grades}이며 실제 개설 여부는 상담 시점에 달라질 수 있습니다. 확인되지 않은 시간표, 차량·주차와 보강 운영은 페이지에서 단정하지 않습니다."
                    if all_grades_listed else
                    f"과목별 가능 학년은 제공 자료에서 모두 확인되지 않습니다. 희망 과목·학년과 실제 개설 여부는 {center['center_name']}에서 확인하세요. 확인되지 않은 시간표, 차량·주차와 보강 운영은 페이지에서 단정하지 않습니다."
                ),
                f"제공 주소 {center['address']}를 기준으로 실제 이동 시간을 계산해 보세요. 온라인 지도와 현장 동선이 다를 수 있으므로 방문 전에 위치도 다시 확인합니다.",
                f"교습비 링크가 있으면 최신 과정과 금액을 확인하고, 링크가 없으면 센터에 직접 문의하세요. 상담 질문과 실제 계약 조건은 구분해서 기록하는 편이 좋습니다.",
                f"체크리스트의 목적은 한 번에 등록을 결정하는 것이 아니라 비교 기준을 남기는 데 있습니다. {title} 상담 뒤에는 확인된 사실과 추가 질문을 나누어 적으세요.",
                f"학교명과 가능 학년 표기는 참고 자료입니다. 학생의 실제 학교·학년과 희망 과목이 현재 운영 범위에 포함되는지 {center['center_name']}에서 재확인하세요.",
                f"상담이 끝나면 첫 점검일, 가정에서 볼 기록과 계획을 조정할 조건을 한 줄씩 남기세요. 설명이 실제 학습 과정으로 이어지는지 확인하는 기준이 됩니다.",
            ]),
        ],
        "checklist": checklist_items,
    }

    middle = [method, school, recommended, plan]
    rotation = stable_int(seed, "section-order") % len(middle)
    middle = middle[rotation:] + middle[:rotation]
    if stable_int(seed, "section-reverse") % 2:
        middle[1:3] = reversed(middle[1:3])
    result = [diagnosis, *middle, checklist]

    return result


def build_faq(config: CategoryConfig, center: dict[str, object], signals: tuple[TopicSignal, ...], seed: str) -> list[tuple[str, str]]:
    locality = str(center["locality"])
    title = f"{locality} {config.label}"
    primary, secondary, support = signals[:3]
    schools = relevant_schools(config, center)
    material = material_label(config, schools)
    grade_items = grade_summary(config, center)
    listed_items = [(subject, value) for subject, value in grade_items if value != "상담 시 확인"]
    missing_subjects = [subject for subject, value in grade_items if value == "상담 시 확인"]
    if missing_subjects:
        if listed_items:
            listed_text = ", ".join(f"{subject} {value}" for subject, value in listed_items)
            grade_answer = (
                f"공개 자료에는 {listed_text} 범위가 기재되어 있습니다. {'·'.join(missing_subjects)} 가능 학년은 "
                f"{center['center_name']}에 희망 학년을 알려 확인하세요. 실제 반 편성·요일도 함께 확인해야 합니다."
            )
        else:
            grade_answer = (
                f"공개 자료에는 {'·'.join(missing_subjects)} 가능 학년이 기재되지 않았습니다. "
                f"{center['center_name']}에 희망 학년을 알려 현재 개설 범위와 실제 반 편성·요일을 확인하세요."
            )
    else:
        grades = ", ".join(f"{subject} {value}" for subject, value in grade_items)
        grade_answer = (
            f"공개된 표기는 {grades}입니다. 실제 반 편성·요일이 필요한 경우에는 "
            f"{center['center_name']}에 희망 과목과 학년을 알려 현재 개설 범위를 확인하세요."
        )
    subject_text = "·".join(config.subjects)
    faq_context = f"{locality} {config.level} {subject_text}"
    faq_topic = f"{config.level} {subject_text}"
    school_text = "·".join(schools[:3]) if schools else "실제 재학 학교"

    intents: dict[str, tuple[str, str]] = {
        "diagnosis": (
            f"{faq_context} 상담에서는 무엇을 먼저 진단하나요?",
            f"최근 점수만 보지 말고 {primary.check}와 {secondary.check}를 먼저 살핍니다. {locality} 학생의 {material}를 가져오면 두 항목이 실제 학습에서 어떻게 나타나는지 구분할 수 있습니다.",
        ),
        "materials": (
            f"{faq_context} 상담 때 학생은 어떤 자료를 준비하면 좋나요?",
            f"{material}에서 어려웠던 부분을 표시해 가져오면 됩니다. 자료가 많지 않아도 {primary.evidence}와 {support.evidence}를 확인할 수 있으면 학습 순서를 질문하기에 충분합니다.",
        ),
        "grades": (
            f"{locality}의 {config.level} {subject_text} 가능 학년과 실제 수업 일정은 어떻게 확인하나요?",
            grade_answer,
        ),
        "school": (
            f"{school_text} 자료는 {faq_context} 계획에 어떻게 반영하나요?",
            f"학교명 자체보다 학생이 가져온 범위와 현재 교재의 진행 위치를 확인합니다. {primary.label}과 {secondary.label}이 드러난 부분을 표시해 주간 과제와 복습 순서에 반영되는지 질문하세요.",
        ),
        "tuition": (
            f"{faq_context} 상담의 교습비와 센터 주소는 어디에서 확인하나요?",
            (f"제공 주소는 {center['address']}이며 교습비 자료는 페이지의 확인 링크로 연결됩니다. " if center["tuition_url"] else f"제공 주소는 {center['address']}이며 교습비는 희망 센터에 직접 문의해야 합니다. ")
            + "과정과 금액, 실제 방문 동선은 변경될 수 있으므로 등록 전에 최신 정보를 다시 확인하세요.",
        ),
        "schedule": (
            f"{faq_context} 수업 시간과 결석 시 보완 방식은 어떻게 확인하나요?",
            f"시간표·반 편성·보완 운영은 센터별로 다를 수 있어 페이지에서 확정하지 않습니다. {locality} 학생의 학교 일정과 통학 시간을 알려 주고 가능한 요일, 피드백 주기와 결석 시 절차를 함께 질문하세요.",
        ),
        "homework": (
            f"{locality}에서 {config.level} {subject_text} 학습을 하는 학생의 과제량은 어떤 기준으로 정하나요?",
            f"많은 분량보다 ‘{primary.practice}’와 ‘{secondary.practice}’ 두 활동을 실제로 마칠 수 있는지를 먼저 봅니다. 완료 기록을 다음 점검에서 확인하고 학교 일정과 가정 학습 시간에 맞춰 분량을 조정하세요.",
        ),
        "feedback": (
            f"{faq_context} 학습에서는 피드백 주기를 어떻게 비교해야 하나요?",
            f"정답을 알려 주는 횟수보다 학생이 이유를 설명하고 다시 수행한 기록을 확인하는지가 중요합니다. ‘{support.home_action}’ 활동을 해 본 뒤의 결과를 언제 확인하는지와 계획을 바꾸는 기준을 상담에서 물어보세요.",
        ),
        "balance": (
            f"{faq_context} 학습 순서는 어떻게 정하나요?",
            f"모든 단원이나 과목을 같은 비중으로 다루지 않고 {primary.label}과 {secondary.label}의 어려움을 구분합니다. 학교 일정과 현재 자료를 바탕으로 집중할 범위와 이어 갈 복습 내용을 나누는 편이 좋습니다.",
        ),
        "four_weeks": (
            f"{faq_context} 4주 계획은 어떤 흐름으로 확인하나요?",
            f"4주 계획은 성과를 보장하는 과정이 아니라 진단·연습·재확인을 정리한 조건부 예시입니다. {primary.label}의 첫 기록과 {support.label}의 재확인 결과를 비교해 다음 범위를 조정합니다.",
        ),
        "consultation": (
            f"{faq_context} 첫 상담 전에 학부모가 정리할 질문은 무엇인가요?",
            f"학생이 막힌 과정, 실제 가능한 학습 시간과 센터 운영 조건을 구분해 적으세요. 이 상담에서는 첫 점검일, 가정에서 확인할 기록과 시간표·교습비 확인 방법을 차례로 질문하면 됩니다.",
        ),
        "results": (
            f"{faq_context} 상담 예시를 실제 후기나 성과로 보아도 되나요?",
            f"페이지의 상담 상황은 실제 이용 후기나 특정 학생의 성과가 아닌 준비용 가상 예시입니다. 학습 결과는 학생의 출발점과 실천 정도에 따라 달라질 수 있으므로 실제 자료와 센터 조건을 확인해 판단하세요.",
        ),
    }
    answer_contexts = {
        "diagnosis": f"{faq_topic} 상담에서 사용할 첫 진단 기준입니다.",
        "materials": f"{faq_topic} 상담 자료는 현재 학습 상태를 보여 주는 범위에서 고릅니다.",
        "grades": f"{faq_topic} 학년·시간표 정보는 공개 자료와 상담 확인 내용을 구분합니다.",
        "school": f"{faq_topic} 학교 자료 반영 방식은 실제 학생 자료를 기준으로 확인합니다.",
        "tuition": f"{faq_topic} 주소·교습비는 제공된 센터 자료를 기준으로 안내합니다.",
        "schedule": f"{faq_topic} 시간표·보완 방식은 현재 센터 운영 범위를 직접 확인해야 합니다.",
        "homework": f"{faq_topic} 과제량은 학생이 수행하고 다시 확인할 수 있는지를 기준으로 비교합니다.",
        "feedback": f"{faq_topic} 피드백은 기록과 다음 점검일이 연결되는지를 살펴봅니다.",
        "balance": f"{faq_topic} 학습 순서는 현재 병목과 학교 일정을 함께 놓고 정합니다.",
        "four_weeks": f"{faq_topic} 4주 흐름은 첫 기록과 마지막 재확인 결과를 이어 봅니다.",
        "consultation": f"{faq_topic} 첫 상담 질문은 학습 자료와 이용 조건을 나누어 준비합니다.",
        "results": f"{faq_topic} 상담 장면을 읽을 때는 출처와 성과 여부를 구분해야 합니다.",
    }
    required = ["diagnosis", "grades"]
    rotating = ["materials", "school", "tuition", "schedule", "homework", "feedback", "balance", "four_weeks", "consultation", "results"]
    start = stable_int(seed, "faq-intents") % len(rotating)
    # Every stride must be coprime with the ten rotating intents; otherwise a
    # stride of five would visit only two entries and never fill five slots.
    stride = (1, 3, 7, 9)[stable_int(seed, "faq-stride") % 4]
    chosen = list(required)
    cursor = start
    while len(chosen) < 5:
        intent = rotating[cursor % len(rotating)]
        if intent not in chosen:
            chosen.append(intent)
        cursor += stride
    if stable_int(seed, "faq-order") % 2:
        chosen[2:] = reversed(chosen[2:])
    return [
        (intents[intent][0], f"{answer_contexts[intent]} {intents[intent][1]}")
        for intent in chosen
    ]


def build_scenarios(config: CategoryConfig, center: dict[str, object], signals: tuple[TopicSignal, ...], seed: str) -> list[str]:
    locality = str(center["locality"])
    title = f"{locality} {config.label}"
    subject_text = "·".join(config.subjects)
    scenario_context = f"{locality} {config.level} {subject_text}"
    primary, secondary, support = signals[:3]
    material = material_label(config, relevant_schools(config, center))
    first = stable_pick(seed, "scenario-first", [
        f"{scenario_context} 상담을 준비하는 학부모가 {material}를 가져오는 장면을 가정했습니다. 학부모는 {primary.label}과 {secondary.label}이 드러난 부분을 보여 주고, 학생이 다음 점검까지 수행할 한 가지 행동이 무엇인지 질문합니다.",
        f"{locality}에서 {config.level} {subject_text} 학습을 하는 학생이 과제를 마쳐도 같은 어려움을 반복하는 상황을 가정했습니다. 학부모는 {primary.evidence}와 {support.evidence}, 두 자료를 나누어 제시하고 설명이 필요한 부분과 혼자 연습할 부분을 확인합니다.",
        f"학교 진도와 현재 실력 사이에서 우선순위를 정하기 어려운 {locality}의 {config.level} {subject_text} 학습 상담 장면입니다. 학부모는 {primary.label}을 먼저 보완할지 {secondary.label} 기록을 더 확인할지 실제 자료를 바탕으로 질문합니다.",
        f"가정 학습 시간이 일정하지 않은 {locality}의 {config.level} {subject_text} 학습 상황을 가정했습니다. 학부모는 많은 과제를 요청하기보다 ‘{support.home_action}’ 활동을 실천할 수 있는 분량과 다음 점검일을 상담에서 확인합니다.",
        f"{locality}에서 {config.level} {subject_text} 수업 설명을 비교하는 학부모 상황입니다. 학부모는 점수 향상 표현보다 {primary.check}와 {secondary.check}를 어떤 자료로 판단하는지 질문합니다.",
        f"{locality}에서 {config.level} {subject_text} 교재를 바꿔야 할지 고민하는 학부모의 상담 장면입니다. 학부모는 {primary.evidence}를 보여 주고 교재 난도와 학습 절차 중 무엇을 먼저 조정할지 확인합니다.",
    ])
    second = stable_pick(seed, "scenario-second", [
        f"{scenario_context} 상담에서 {center['center_name']}의 제공 주소와 가능 학년을 확인한 뒤 실제 시간표, 통학 시간과 교습비를 따로 비교하는 상황입니다. 학습 계획과 이용 조건을 한 문장으로 묶지 않고 확인된 사실과 추가 질문을 나누어 기록합니다.",
        f"{locality}에서 {config.level} {subject_text} 상담을 마친 학부모가 학생의 학교 일정과 가정 학습 시간을 대조하는 장면입니다. 첫 계획을 그대로 따르기보다 실행하지 못했을 때 분량과 순서를 어떻게 바꿀지 질문합니다.",
        f"{scenario_context} 학습 자료가 공개 학교 목록과 다를 수 있는 경우를 가정했습니다. 학부모는 실제 재학 학교와 현재 진도를 알려 주고 수업 가능 여부와 자료 반영 방식을 센터에서 다시 확인합니다.",
        f"{locality}에서 {config.level} {subject_text} 학습 시간이 겹치는 가정의 상담 상황입니다. 학부모는 {primary.label}에 집중할 시간과 {secondary.label}을 복습할 시간을 구분하고 일주일 뒤 어떤 기록으로 계획을 재검토할지 묻습니다.",
        f"{locality}에서 {config.level} {subject_text} 첫 상담을 마친 뒤 학생과 학부모가 내용을 다시 정리하는 상황입니다. ‘{support.practice}’를 실제로 해낼 수 있는지 살펴보고, 어려우면 다음 상담에서 분량과 피드백 주기를 조정합니다.",
        f"{scenario_context} 상담에서 센터 설명과 학생 자료가 일치하는지 비교하는 상황입니다. 확인되지 않은 차량·주차·보강 운영을 추정하지 않고 현재 개설 과목과 학년, 방문 조건을 직접 확인합니다.",
    ])
    return [first, second]


def offer_nodes(config: CategoryConfig, center: dict[str, object], service_id: str) -> list[dict[str, object]]:
    offers: list[dict[str, object]] = []
    for subject in config.subjects:
        grades = relevant_grades(config, center, subject)
        name = f"{config.level} {subject} 학습 상담" if config.grade_prefix else f"{subject} 학습 상담"
        offer: dict[str, object] = {
            "@type": "Offer",
            "name": name,
            "itemOffered": {"@id": service_id, "@type": "Service", "name": name},
        }
        if grades:
            offer["eligibleCustomerType"] = "·".join(grades)
        offers.append(offer)
    if center["tuition_url"]:
        offers.append({
            "@type": "Offer",
            "name": f"{center['center_name']} 교습과정·교습비 확인",
            "url": center["tuition_url"],
            "itemOffered": {"@type": "Service", "name": "센터별 교습과정 정보 확인"},
        })
    return offers


def related_links(config: CategoryConfig, center: dict[str, object], previous_slug: str, next_slug: str) -> list[tuple[str, str]]:
    locality = str(center["locality"])
    local_slug = str(center["slug"])
    related_slugs: list[str]
    if config.slug == "영수학원":
        related_slugs = ["초등학생국영수학원", "중학생국영수학원", "고등학생국영수학원", "중등영어학원", "중등수학학원"]
    elif config.grade_prefix == "초":
        related_slugs = ["초등영어학원", "초등수학학원", "초등학생국영수학원", "영수학원"]
    elif config.grade_prefix == "중":
        related_slugs = ["중등영어학원", "중등수학학원", "중학생국영수학원", "영수학원"]
    else:
        related_slugs = ["고등영어학원", "고등수학학원", "고등학생국영수학원", "영수학원"]
    links: list[tuple[str, str]] = [(f"{config.label} 전체 지역", route_for(config.slug))]
    for slug in related_slugs:
        if slug == config.slug:
            continue
        links.append((f"{locality} {ALL_CATEGORY_LABELS[slug]}", route_for(slug, local_slug)))
    links.extend([
        (f"이전 지역 · {previous_slug}", route_for(config.slug, previous_slug)),
        (f"다음 지역 · {next_slug}", route_for(config.slug, next_slug)),
        ("학습가이드", "/학습가이드/"),
        ("상담문의", "/상담문의/"),
    ])
    return links


def build_graph(
    config: CategoryConfig,
    center: dict[str, object],
    title: str,
    meta: str,
    rep_name: str,
    sections: list[dict[str, object]],
    faq: list[tuple[str, str]],
    links: list[tuple[str, str]],
) -> dict[str, object]:
    locality = str(center["locality"])
    slug = str(center["slug"])
    url = absolute_route(config.slug, slug)
    hub_url = absolute_route(config.slug)
    org_id = center_entity_id(center)
    service_id = url + "#service"
    rep_url = ORIGIN + f"/assets/representative/{quote(rep_name)}"
    body_url = ORIGIN + f"/assets/centers/common/{center['body_image']}"
    map_url = ORIGIN + "/assets/maps/" + quote(str(center["map_name"]))
    schools = relevant_schools(config, center)
    grades = list(dict.fromkeys(grade for subject in config.subjects for grade in relevant_grades(config, center, subject)))
    offers = offer_nodes(config, center, service_id)
    about = [
        {"@type": "Thing", "name": config.label},
        *({"@type": "Thing", "name": f"{config.level} {subject} 학습"} for subject in config.subjects),
        {"@type": "Thing", "name": "학습 진단"},
        {"@type": "Thing", "name": "오답 재학습"},
    ]
    mentions = [
        {"@type": "Place", "name": locality},
        *({"@type": "EducationalOrganization", "name": school} for school in schools),
        *({"@type": "Thing", "name": subject + " 학습"} for subject in config.subjects),
    ]
    section_parts = [
        {"@type": "WebPageElement", "@id": f"{url}#section-{index}", "name": str(section["heading"])}
        for index, section in enumerate(sections, 1)
    ]
    organization: dict[str, object] = {
        "@type": ["EducationalOrganization", "LocalBusiness"],
        "@id": org_id,
        "name": center["center_name"],
        "url": ORIGIN + "/",
        "telephone": PHONE,
        "address": {
            "@type": "PostalAddress",
            "streetAddress": center["address"],
            "addressRegion": center["address_region"],
            "addressCountry": "KR",
        },
        "areaServed": [{"@type": "Place", "name": locality}],
        "description": f"{center['center_name']}의 제공 주소·등록정보와 {locality} 학습 상담 확인 경로를 안내합니다.",
        "image": rep_url,
        "teaches": [f"{config.level} {subject}" for subject in config.subjects] + ["학습코칭"],
        "makesOffer": offers,
    }
    if center["registration"]:
        organization["identifier"] = {
            "@type": "PropertyValue",
            "name": center["office_name"] or "등록 학원명",
            "value": center["registration"],
        }
    if grades:
        organization["educationalLevel"] = grades
    breadcrumb = {
        "@type": "BreadcrumbList",
        "@id": url + "#breadcrumb",
        "itemListElement": [
            {"@type": "ListItem", "position": 1, "name": "홈", "item": ORIGIN + "/"},
            {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": ORIGIN + quote("/과목별학원/", safe="/")},
            {"@type": "ListItem", "position": 3, "name": config.label, "item": hub_url},
            {"@type": "ListItem", "position": 4, "name": title, "item": url},
        ],
    }
    image_object = {
        "@type": "ImageObject",
        "@id": url + "#primaryimage",
        "url": rep_url,
        "contentUrl": rep_url,
        "caption": f"{title} 대표 이미지",
    }
    article = {
        "@type": "Article",
        "@id": url + "#article",
        "url": url,
        "headline": title,
        "description": meta,
        "abstract": meta,
        "inLanguage": "ko-KR",
        "mainEntityOfPage": {"@id": url + "#webpage"},
        "author": {"@id": org_id},
        "publisher": {"@id": org_id},
        "datePublished": CONTENT_DATE,
        "dateModified": CONTENT_DATE,
        "articleSection": [config.label, str(center["region"]), str(center["district"]), locality],
        "about": about,
        "mentions": mentions,
        "hasPart": section_parts,
        "image": [rep_url, body_url, map_url],
    }
    if grades:
        article["educationalLevel"] = grades
    service = {
        "@type": "Service",
        "@id": service_id,
        "url": url,
        "name": f"{title} 학습 상담 안내",
        "serviceType": config.label,
        "provider": {"@id": org_id},
        "areaServed": {"@type": "Place", "name": locality},
        "audience": {"@type": "EducationalAudience", "educationalRole": "student", "audienceType": config.level},
        "description": meta,
        "offers": offers,
    }
    webpage = {
        "@type": "WebPage",
        "@id": url + "#webpage",
        "url": url,
        "name": title,
        "description": meta,
        "inLanguage": "ko-KR",
        "isPartOf": {"@id": hub_url + "#collection"},
        "breadcrumb": {"@id": breadcrumb["@id"]},
        "primaryImageOfPage": {"@id": image_object["@id"]},
        "about": about,
        "mentions": mentions,
        "hasPart": section_parts,
        "mainEntity": {"@id": article["@id"]},
    }
    faq_node = {
        "@type": "FAQPage",
        "@id": url + "#faq",
        "mainEntity": [
            {"@type": "Question", "name": question, "acceptedAnswer": {"@type": "Answer", "text": answer}}
            for question, answer in faq
        ],
    }
    item_list = {
        "@type": "ItemList",
        "@id": url + "#related",
        "name": f"{title} 관련 페이지",
        "numberOfItems": len(links),
        "itemListElement": [
            {"@type": "ListItem", "position": index, "name": name, "url": ORIGIN + quote(href, safe="/%:@")}
            for index, (name, href) in enumerate(links, 1)
        ],
    }
    return {
        "@context": "https://schema.org",
        "@graph": [organization, image_object, webpage, breadcrumb, article, service, faq_node, item_list],
    }


def header(prefix: str) -> str:
    items = [
        ("home", "홈", "index.html"),
        ("about", "학원소개", "학원소개/index.html"),
        ("guide", "학습가이드", "학습가이드/index.html"),
        ("contact", "상담문의", "상담문의/index.html"),
        ("subjects", "과목별학원", "과목별학원/index.html"),
    ]
    links = "".join(
        f'<a href="{prefix}{href}" data-nav="{key}"{" aria-current=\"page\"" if key == "subjects" else ""}>{label}</a>'
        for key, label, href in items
    )
    return f'''<a class="skip-link" href="#main">본문으로 건너뛰기</a>
  <header class="site-header"><div class="site-shell header-inner">
    <a class="brand" href="{prefix}index.html"><span class="brand-mark" aria-hidden="true">W</span><span class="brand-copy"><strong>{SITE_NAME}</strong><small>STUDY RECORD COACHING</small></span></a>
    <nav class="primary-nav" aria-label="주요 메뉴">{links}</nav>
  </div></header>'''


def footer(prefix: str) -> str:
    return f'''<footer class="site-footer"><div class="site-shell footer-grid">
    <div class="footer-brand"><h2>{SITE_NAME}</h2><p>학생별 진도와 교재, 실행 기록과 오답 재학습을 연결해 다음 공부 순서를 정리합니다. 개설 과목과 학년, 수업 방식은 센터별로 확인합니다.</p></div>
    <div><nav class="footer-links" aria-label="하단 메뉴"><a href="{prefix}학원소개/index.html">학원소개</a><a href="{prefix}학습가이드/index.html">학습가이드</a><a href="{prefix}과목별학원/index.html">과목별학원</a><a href="{prefix}상담문의/index.html">상담문의</a></nav><p class="footer-meta">대표 상담 {PHONE}<br>© {SITE_NAME}</p></div>
  </div></footer>
  <nav class="contact-dock" aria-label="빠른 상담"><a href="tel:010-6839-8283">전화문의</a><a href="https://blogsms.net/01068398283" target="_blank" rel="noopener">문자문의</a><a href="https://docs.google.com/forms/d/e/1FAIpQLSdb2oE5Qk5YS0TfYDxyV1w-IOTkhkjOCmmpAKTI9FmqpVj6Yg/viewform" target="_blank" rel="noopener">상담신청</a></nav>'''


def render_info(config: CategoryConfig, center: dict[str, object]) -> str:
    rows = [
        ("지역", display_region_label(center)),
        ("센터 기준", center["center_name"]),
        ("제공 주소", center["address"]),
    ]
    if center["registration"]:
        rows.append(("등록 정보", center["registration"]))
    html_rows = "".join(f"<div><dt>{esc(label)}</dt><dd>{esc(value)}</dd></div>" for label, value in rows if value)
    schools = relevant_schools(config, center)
    school_html = (
        f'<div><dt>{esc(config.school_name)} 참고</dt><dd><div class="local-tags">{"".join(f"<span>{esc(school)}</span>" for school in schools[:8])}</div></dd></div>'
        if schools else f'<div><dt>{esc(config.school_name)} 참고</dt><dd>제공 목록 없음 · 상담 시 실제 학교 자료 확인</dd></div>'
    )
    grades = "".join(
        f'<li><strong>{esc(subject)}</strong><span>{esc(value)}</span></li>'
        for subject, value in grade_summary(config, center)
    )
    grade_note = ", ".join(f"{subject} {value}" for subject, value in grade_summary(config, center))
    tuition = (
        f'<a class="button compact" href="{esc(center["tuition_url"])}" target="_blank" rel="noopener">센터별 교습비 확인 <span aria-hidden="true">↗</span></a>'
        if center["tuition_url"] else '<p class="info-note">교습비 자료는 희망 센터에서 확인합니다.</p>'
    )
    return (
        f'<dl class="local-facts">{html_rows}{school_html}</dl><ul class="grade-list">{grades}</ul>'
        f'<p class="center-verified-note"><strong>제공 자료 확인 기준</strong><span>표기된 학년은 제공 자료 기준이며, 시간표와 실제 개설 여부는 상담 시 확인합니다.</span></p>{tuition}'
    )


def render_page(
    config: CategoryConfig,
    record: dict[str, object],
    previous_record: dict[str, object],
    next_record: dict[str, object],
) -> str:
    center = record["center"]
    locality = str(center["locality"])
    title = str(record["title"])
    meta = str(record["meta"])
    rep_name = str(record["rep_name"])
    signals = record["signals"]
    sections = record["sections"]
    faq = record["faq"]
    scenarios = record["scenarios"]
    student = str(record["student"])
    quick = str(record["quick"])
    student_level = student_level_label(config)
    quick_student_note = (
        f"{locality} {student_level} 상담에서는 {signals[0].label}·{signals[1].label}의 현재 상태를 "
        "실제 자료에서 나누어 확인합니다."
    )
    links = related_links(config, center, str(previous_record["slug"]), str(next_record["slug"]))
    graph = build_graph(config, center, title, meta, rep_name, sections, faq, links)
    graph_json = json.dumps(graph, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    url = absolute_route(config.slug, str(center["slug"]))
    rep_url = ORIGIN + f"/assets/representative/{quote(rep_name)}"
    cards_html = "".join(
        f'<article><span>{esc(label)}</span><strong>{esc(value)}</strong></article>'
        for label, value in answer_cards(config, center, signals, student)
    )
    section_html_parts: list[str] = []
    for index, section in enumerate(sections, 1):
        paragraphs = "".join(f"<p>{esc(paragraph)}</p>" for paragraph in section["paragraphs"])
        checklist = section.get("checklist")
        if checklist:
            paragraphs += '<ol class="checklist">' + "".join(
                f'<li><div><strong>{esc(label)}</strong><span>{esc(value)}</span></div></li>'
                for label, value in checklist
            ) + "</ol>"
        section_html_parts.append(
            f'<section class="manuscript-section" id="section-{index}"><span class="section-kicker">{index:02d}</span><h2>{esc(section["heading"])}</h2>{paragraphs}</section>'
        )
    body_html = "".join(section_html_parts)
    faq_html = "".join(
        f'<details{" open" if index == 0 else ""}><summary>{esc(question)}</summary><p>{esc(answer)}</p></details>'
        for index, (question, answer) in enumerate(faq)
    )
    scenario_html = "".join(
        f'<article class="scenario-card"><span>학부모 관점 상담 예시 {index:02d}</span><p>{esc(value)}</p></article>'
        for index, value in enumerate(scenarios, 1)
    )
    related_html = "".join(
        f'<a class="local-nav-card{" is-parent" if index == 0 else ""}" href="{esc(href)}"><small>{"카테고리" if index == 0 else "관련 페이지"}</small><strong>{esc(name)}</strong><span>→</span></a>'
        for index, (name, href) in enumerate(links)
    )
    region_label = display_region_label(center)
    return f'''<!doctype html>
<html lang="ko"><head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)} | {SITE_NAME}</title><meta name="description" content="{esc(meta)}"><meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec">
  <link rel="canonical" href="{url}"><meta property="og:type" content="article"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="{esc(title)} | {SITE_NAME}"><meta property="og:description" content="{esc(meta)}"><meta property="og:url" content="{url}"><meta property="og:image" content="{rep_url}">
  <link rel="icon" href="../../../assets/favicon.png"><link rel="stylesheet" href="../../../assets/site14.css"><script type="application/ld+json">{graph_json}</script>
</head><body data-page="subjects">{header("../../../")}
  <main id="main">
    <section class="local-hero"><div class="site-shell">
      <nav class="breadcrumbs" aria-label="현재 위치"><a href="../../../index.html">홈</a><a href="../../index.html">과목별학원</a><a href="../index.html">{esc(config.label)}</a><span>{esc(title)}</span></nav>
      <p class="eyebrow">{esc(config.english)}</p><h1>{esc(title)}</h1><p class="local-lead">{esc(meta)}</p><div class="local-answer-grid">{cards_html}</div>
    </div></section>
    <section class="section local-overview"><div class="site-shell local-overview-grid">
      <div class="local-summary"><p class="chapter-label"><span>01</span> Quick answer</p><h2>{esc(locality)} {esc(config.level)} {esc('·'.join(config.subjects))} 학습에서 먼저 확인할 핵심 답변</h2><p>{esc(quick)}</p><div class="answer-note"><strong>확인 기준</strong><p>{esc(quick_student_note)}</p></div></div>
      <aside class="local-info-card"><p class="eyebrow">Center information</p><h2>지역·학년·센터 확인 정보</h2>{render_info(config, center)}</aside>
    </div></section>
    <section class="local-media-section"><div class="site-shell local-media-stack">
      <img src="../../../assets/representative/{esc(rep_name)}" alt="{esc(title)} {SITE_NAME} 대표" style="display:none;">
      <figure class="local-body-image"><img src="../../../assets/centers/common/{esc(center['body_image'])}" width="918" height="16116" alt="{esc(title)} 본문 학습 안내" loading="lazy" decoding="async"><figcaption>{esc(region_label)} {esc(config.label)} 학습 점검 안내</figcaption></figure>
      <figure class="local-map-image"><div class="map-art"><img src="../../../assets/maps/{esc(center['map_name'])}" alt="{esc(title)} 지도 {esc(center['center_name'])}" loading="lazy" decoding="async"></div><figcaption>센터 위치는 제공 주소를 기준으로 표시하며 방문 전 실제 운영 여부와 동선을 확인합니다.</figcaption></figure>
    </div></section>
    <section class="section manuscript-wrap"><article class="site-shell manuscript-article"><div class="manuscript-intro"><span>학습 답변 요약</span><p>{esc(locality)} {esc(config.level)} {esc('·'.join(config.subjects))} 학습 자료를 기준으로 진단, 학습 순서, 학교·센터 사실과 상담 질문을 차례로 살펴봅니다.</p></div>{body_html}</article></section>
    <section class="section blue-wash"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>02</span> Parent perspective</div><div><h2>{esc(locality)} {esc(config.level)} {esc('·'.join(config.subjects))} 학부모 상담 상황</h2><p>아래 {esc(locality)} {esc(config.level)} {esc('·'.join(config.subjects))} 상담 장면은 실제 이용 후기나 특정 학생의 성과가 아닌 준비용 가상 예시입니다. 학습 결과는 학생의 출발점과 실천 정도에 따라 달라질 수 있습니다.</p></div></div><div class="scenario-grid">{scenario_html}</div></div></section>
    <section class="section"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>03</span> FAQ</div><div><h2>{esc(locality)} {esc(config.level)} {esc('·'.join(config.subjects))} 자주 묻는 질문</h2><p>학생 자료, 가능 학년과 실제 이용 조건을 나누어 답했습니다.</p></div></div><div class="faq-list">{faq_html}</div></div></section>
    <section class="section local-links-section"><div class="site-shell"><div class="section-heading compact-heading"><div class="chapter-label"><span>04</span> Related</div><div><h2>{esc(locality)} 관련 학원·상담 페이지</h2><p>같은 지역의 과목 조합과 학년 카테고리, 앞뒤 지역 안내를 함께 확인하세요.</p></div></div><div class="local-navigation">{related_html}</div></div></section>
  </main>{footer("../../../")}<script src="../../../assets/site14.js" defer></script>
</body></html>'''


def grouped_records(records: list[dict[str, object]]) -> dict[str, dict[str, list[dict[str, object]]]]:
    grouped: dict[str, dict[str, list[dict[str, object]]]] = defaultdict(lambda: defaultdict(list))
    for record in records:
        center = record["center"]
        grouped[str(center["region"])][str(center["district"])].append(record)
    return grouped


def render_category_hub(config: CategoryConfig, records: list[dict[str, object]]) -> str:
    url = absolute_route(config.slug)
    items = [
        {
            "@type": "ListItem",
            "position": index,
            "name": record["title"],
            "url": absolute_route(config.slug, str(record["slug"])),
        }
        for index, record in enumerate(records, 1)
    ]
    graph = {
        "@context": "https://schema.org",
        "@graph": [
            {"@type": "EducationalOrganization", "@id": ORIGIN + "/#organization", "name": SITE_NAME, "url": ORIGIN + "/", "telephone": PHONE, "teaches": [f"{config.level} {subject}" for subject in config.subjects]},
            {"@type": "BreadcrumbList", "@id": url + "#breadcrumb", "itemListElement": [
                {"@type": "ListItem", "position": 1, "name": "홈", "item": ORIGIN + "/"},
                {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": ORIGIN + quote("/과목별학원/", safe="/")},
                {"@type": "ListItem", "position": 3, "name": config.label, "item": url},
            ]},
            {"@type": "CollectionPage", "@id": url + "#collection", "url": url, "name": f"{config.label} 지역 안내", "description": f"371개 동네별 {config.label} 학습 답변과 센터·학교·가능 학년 확인 정보를 제공합니다.", "inLanguage": "ko-KR", "breadcrumb": {"@id": url + "#breadcrumb"}, "hasPart": [{"@type": "WebPage", "name": record["title"], "url": absolute_route(config.slug, str(record["slug"]))} for record in records]},
            {"@type": "ItemList", "@id": url + "#directory", "name": f"{config.label} 371개 지역", "numberOfItems": len(items), "itemListElement": items},
        ],
    }
    graph_json = json.dumps(graph, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    grouped = grouped_records(records)
    region_filters = "".join(f'<button type="button" data-region-filter="{esc(region)}">{esc(region)}</button>' for region in REGION_ORDER if region in grouped)
    region_html: list[str] = []
    for region in REGION_ORDER:
        if region not in grouped:
            continue
        districts = grouped[region]
        district_html: list[str] = []
        for district, values in districts.items():
            cards = "".join(
                f'<a class="directory-card" href="{esc(record["slug"])}/index.html" data-locality="{esc(record["center"]["locality"])} {esc(record["title"])} {esc(record["center"]["center_name"])}"><strong>{esc(record["center"]["locality"])}</strong><span>{esc(config.label)} · {esc(record["signals"][0].label)}</span><i aria-hidden="true">→</i></a>'
                for record in values
            )
            district_html.append(f'<section class="directory-district"><div class="directory-district-head"><h2>{esc(district)}</h2><span>{len(values)}개 지역</span></div><div class="directory-grid">{cards}</div></section>')
        region_html.append(f'<details class="directory-region" data-region="{esc(region)}" open><summary><span><b>{esc(region)}</b><small>{sum(len(value) for value in districts.values())}개 지역</small></span><i aria-hidden="true">＋</i></summary><div class="directory-region-body">{"".join(district_html)}</div></details>')
    description = f"371개 동네별 {config.label} 페이지에서 {config.focus}, {config.school_name} 자료, 가능 학년과 센터 상담 체크리스트를 확인하세요."
    return f'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{esc(config.label)} 지역 안내 | {SITE_NAME}</title><meta name="description" content="{esc(description)}"><meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec"><link rel="canonical" href="{url}"><meta property="og:type" content="website"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="{esc(config.label)} 지역 안내 | {SITE_NAME}"><meta property="og:description" content="{esc(description)}"><meta property="og:url" content="{url}"><link rel="icon" href="../../assets/favicon.png"><link rel="stylesheet" href="../../assets/site14.css"><script type="application/ld+json">{graph_json}</script></head>
<body data-page="subjects">{header("../../")}<main id="main"><section class="directory-hero"><div class="site-shell"><nav class="breadcrumbs" aria-label="현재 위치"><a href="../../index.html">홈</a><a href="../index.html">과목별학원</a><span>{esc(config.label)}</span></nav><p class="eyebrow">{esc(config.english)} DIRECTORY</p><h1>동네별 {esc(config.label)}</h1><p>{esc(config.focus)}을 중심으로 학생 자료를 확인하고, 지역별 센터·학교·가능 학년과 상담 전 질문을 한 흐름에서 비교하세요.</p><div class="hub-metrics"><div><strong>371</strong><span>지역 페이지</span></div><div><strong>{esc(config.focus)}</strong><span>핵심 진단</span></div><div><strong>{esc(config.process)}</strong><span>상담 흐름</span></div></div></div></section>
<section class="section directory-section"><div class="site-shell"><div class="directory-toolbar"><label for="local-search-{esc(config.slug)}">동네·센터·학교 검색</label><div class="directory-search"><input id="local-search-{esc(config.slug)}" type="search" placeholder="예: 명일동, 강동구, 명일점" data-local-search><span data-directory-count>371개 지역</span></div><div class="region-filters"><button type="button" class="is-active" data-region-filter="all">전체</button>{region_filters}</div><div class="directory-actions"><button type="button" data-expand-all>모두 펼치기</button></div></div><p class="directory-empty" data-directory-empty hidden>검색 조건에 맞는 지역이 없습니다.</p><div class="directory-list">{"".join(region_html)}</div></div></section>
<section class="section ink"><div class="site-shell consult-cta"><div><h2>{esc(config.label)} 상담은 실제 자료에서 시작하세요</h2><p>현재 교재와 오답 기록, 학교 일정과 가능한 학습 시간을 준비하면 설명·연습·재확인 기준을 구체적으로 비교할 수 있습니다.</p></div><a class="button orange" href="../../상담문의/index.html">상담 방법 확인 <span aria-hidden="true">→</span></a></div></section></main>{footer("../../")}<script src="../../assets/site14.js" defer></script></body></html>'''


def build_records(config: CategoryConfig, centers: list[dict[str, object]], source_rows: list[str], reps: list[str]) -> list[dict[str, object]]:
    if not (len(centers) == len(source_rows) == len(reps) == EXPECTED_ROWS):
        raise ValueError(f"{config.slug}: row mismatch")
    records: list[dict[str, object]] = []
    for index, (center, raw, rep_name) in enumerate(zip(centers, source_rows, reps), 1):
        seed = f"{config.slug}|{center['locality']}|{index}"
        signals = rank_signals(config, raw, seed)
        title = f"{center['locality']} {config.label}"
        student = student_type(config, signals, seed, str(center["locality"]))
        tokens = page_particle_tokens(config, center, signals, student)
        student = normalize_particle_joins(student, tokens)
        sections = normalize_generated_value(build_sections(config, center, signals, seed), tokens)
        faq = normalize_generated_value(build_faq(config, center, signals, seed), tokens)
        scenarios = normalize_generated_value(build_scenarios(config, center, signals, seed), tokens)
        records.append({
            "center": center,
            "title": title,
            "slug": center["slug"],
            "rep_name": rep_name,
            "signals": signals,
            "student": student,
            "meta": normalize_particle_joins(meta_description(config, center, signals), tokens),
            "quick": normalize_particle_joins(quick_answer(config, center, signals, seed), tokens),
            "sections": sections,
            "faq": faq,
            "scenarios": scenarios,
        })
    return records


def contextualize_duplicate_paragraphs(records_by_category: dict[str, list[dict[str, object]]]) -> None:
    """Add page context only where a generated paragraph would otherwise repeat."""
    occurrences: dict[str, list[tuple[str, dict[str, object], dict[str, object], int]]] = defaultdict(list)
    occupied: set[str] = set()
    used_contexts: dict[tuple[str, str], set[str]] = defaultdict(set)
    for slug, records in records_by_category.items():
        for record in records:
            for section in record["sections"]:
                for paragraph_index, paragraph in enumerate(section["paragraphs"]):
                    value = clean(paragraph)
                    occupied.add(value)
                    occurrences[value].append((slug, record, section, paragraph_index))

    for paragraph, references in occurrences.items():
        if len(references) < 2:
            continue
        for slug, record, section, paragraph_index in references:
            config = CONFIG_BY_SLUG[slug]
            center = record["center"]
            locality = str(center["locality"])
            subject_text = "·".join(config.subjects)
            contexts = (
                f"이 기준은 {locality}에서 {config.level} {subject_text} 상담을 준비할 때 학생 자료와 함께 확인합니다.",
                f"{locality} {config.level} {subject_text} 계획에서는 이 내용을 실제 학습 기록과 대조합니다.",
                f"{locality} 학부모가 {config.level} {subject_text} 상담을 준비한다면 이 항목도 따로 메모해 두세요.",
                f"{locality} 학생의 {config.level} {subject_text} 흐름에 적용할 때는 현재 교재와 실행 기록을 함께 살핍니다.",
                f"이 내용은 {locality} {config.level} {subject_text} 수업 조건을 비교할 때 확인할 질문으로 사용할 수 있습니다.",
                f"{locality}에서 {config.level} {subject_text} 학습을 알아볼 때도 확인된 자료와 상담 질문을 구분합니다.",
                f"{locality}의 {config.level} {subject_text} 학습에서는 이 기준을 다음 점검일과 연결해 살핍니다.",
                f"상담에는 {locality} 학생의 {config.level} {subject_text} 기록도 함께 가져가세요.",
                f"이 확인 순서는 {locality} {config.level} {subject_text} 학습 자료를 비교할 때 사용할 수 있습니다.",
                f"{locality} {config.level} {subject_text} 상담 메모에서는 이 내용과 실제 이용 조건을 나누어 적습니다.",
                f"현재 자료를 비교하는 {locality} {config.level} {subject_text} 과정에서도 이 항목을 별도로 확인합니다.",
                f"이 질문은 {locality}에서 {config.level} {subject_text} 계획을 정하기 전에 학생과 함께 살펴봅니다.",
            )
            context_seed = f"{slug}|{locality}|{section['key']}|{paragraph_index}|paragraph-context"
            start = stable_int(context_seed, "context") % len(contexts)
            candidate = ""
            page_key = (slug, str(record["slug"]))
            for offset in range(len(contexts)):
                context = contexts[(start + offset) % len(contexts)]
                candidate = f"{paragraph} {context}"
                if context not in used_contexts[page_key] and candidate not in occupied:
                    used_contexts[page_key].add(context)
                    break
            else:
                raise ValueError(f"unable to contextualize duplicate paragraph: {record['title']}")
            section["paragraphs"][paragraph_index] = candidate
            occupied.add(candidate)


def contextualize_duplicate_faq_answers(records_by_category: dict[str, list[dict[str, object]]]) -> None:
    """Keep FAQ answers distinct without repeating the locality on every answer."""
    occurrences: dict[str, list[tuple[str, dict[str, object], int]]] = defaultdict(list)
    occupied: set[str] = set()
    used_contexts: dict[tuple[str, str], set[str]] = defaultdict(set)
    for slug, records in records_by_category.items():
        for record in records:
            for faq_index, (_question, answer) in enumerate(record["faq"]):
                value = clean(answer)
                occupied.add(value)
                occurrences[value].append((slug, record, faq_index))

    for answer, references in occurrences.items():
        if len(references) < 2:
            continue
        for slug, record, faq_index in references:
            config = CONFIG_BY_SLUG[slug]
            center = record["center"]
            locality = str(center["locality"])
            subject_text = "·".join(config.subjects)
            contexts = (
                f"{locality}의 {config.level} {subject_text} 상담에서는 학생 자료와 센터 조건을 함께 확인합니다.",
                f"이 답변은 {locality}에서 {config.level} {subject_text} 상담을 준비할 때 적용할 수 있습니다.",
                f"{locality} {config.level} {subject_text} 학습에서는 실제 자료를 확인한 뒤 판단해야 합니다.",
                f"{locality} 학부모는 {config.level} {subject_text} 질문과 이용 조건을 나누어 기록해 두세요.",
                f"{center['center_name']}의 {locality} {config.level} {subject_text} 운영 여부는 상담 시점에 다시 확인합니다.",
            )
            seed = f"{slug}|{locality}|{faq_index}|faq-context"
            start = stable_int(seed, "context") % len(contexts)
            candidate = ""
            page_key = (slug, str(record["slug"]))
            for offset in range(len(contexts)):
                context = contexts[(start + offset) % len(contexts)]
                candidate = f"{answer} {context}"
                if context not in used_contexts[page_key] and candidate not in occupied:
                    used_contexts[page_key].add(context)
                    break
            else:
                raise ValueError(f"unable to contextualize duplicate FAQ answer: {record['title']}")
            question, _old_answer = record["faq"][faq_index]
            record["faq"][faq_index] = (question, candidate)
            occupied.add(candidate)


def preflight(records_by_category: dict[str, list[dict[str, object]]]) -> dict[str, object]:
    titles: list[str] = []
    metas: list[str] = []
    paragraphs: dict[str, str] = {}
    sections: dict[str, str] = {}
    authored_blocks: dict[str, str] = {}
    faq_sets: set[str] = set()
    scenario_sets: set[str] = set()
    for slug, records in records_by_category.items():
        if len(records) != EXPECTED_ROWS:
            raise ValueError(f"{slug}: record count={len(records)}")
        for record in records:
            title = str(record["title"])
            titles.append(title)
            meta = str(record["meta"])
            if not 70 <= len(meta) <= 100 or not meta.endswith((".", "요.")):
                raise ValueError(f"invalid meta description: {title} ({len(meta)}): {meta}")
            metas.append(meta)
            page_blocks: list[str] = [str(record["student"]), str(record["quick"])]
            for section in record["sections"]:
                page_blocks.append(str(section["heading"]))
                section_text = clean(" ".join([str(section["heading"]), *section["paragraphs"]]))
                if section_text in sections:
                    raise ValueError(f"duplicate authored section: {title} / {sections[section_text]}")
                sections[section_text] = title
                for paragraph in section["paragraphs"]:
                    value = clean(paragraph)
                    if value in paragraphs:
                        raise ValueError(f"duplicate authored paragraph: {title} / {paragraphs[value]}")
                    paragraphs[value] = title
                    page_blocks.append(value)
                page_blocks.extend(clean(value) for _label, value in section.get("checklist", []))
            for question, answer in record["faq"]:
                page_blocks.extend((clean(question), clean(answer)))
            page_blocks.extend(clean(value) for value in record["scenarios"])
            for value in page_blocks:
                if value in authored_blocks:
                    raise ValueError(f"duplicate authored block: {title} / {authored_blocks[value]}: {value}")
                authored_blocks[value] = title
            faq_key = json.dumps(record["faq"], ensure_ascii=False)
            scenario_key = json.dumps(record["scenarios"], ensure_ascii=False)
            if faq_key in faq_sets or scenario_key in scenario_sets:
                raise ValueError(f"duplicate faq/scenario set: {title}")
            faq_sets.add(faq_key)
            scenario_sets.add(scenario_key)
    if len(titles) != len(set(titles)):
        raise ValueError("duplicate titles")
    if len(metas) != len(set(metas)):
        raise ValueError("duplicate meta descriptions")
    return {
        "detail_pages": len(titles),
        "unique_titles": len(set(titles)),
        "unique_meta": len(set(metas)),
        "unique_paragraphs": len(paragraphs),
        "unique_sections": len(sections),
        "unique_faq_sets": len(faq_sets),
        "unique_scenario_sets": len(scenario_sets),
        "unique_authored_blocks": len(authored_blocks),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate seven subject academy categories from XLSX source signals.")
    parser.add_argument("--category", action="append", choices=tuple(CONFIG_BY_SLUG), help="Generate only selected category; repeatable.")
    parser.add_argument("--skip-production-files", action="store_true", help="Do not rebuild canonical/sitemap/rss/social metadata.")
    args = parser.parse_args()
    selected = tuple(CONFIG_BY_SLUG[slug] for slug in args.category) if args.category else CONFIGS
    centers = load_centers()
    assignments = assign_representatives(selected)
    records_by_category: dict[str, list[dict[str, object]]] = {}
    for config in selected:
        records_by_category[config.slug] = build_records(config, centers, load_source_rows(config), assignments[config.slug])
    contextualize_duplicate_paragraphs(records_by_category)
    contextualize_duplicate_faq_answers(records_by_category)
    report = preflight(records_by_category)
    for config in selected:
        records = records_by_category[config.slug]
        category_root = TARGET_ROOT / config.slug
        category_root.mkdir(parents=True, exist_ok=True)
        for index, record in enumerate(records):
            target = category_root / str(record["slug"]) / "index.html"
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(render_page(config, record, records[(index - 1) % len(records)], records[(index + 1) % len(records)]), encoding="utf-8")
        (category_root / "index.html").write_text(render_category_hub(config, records), encoding="utf-8")

    # The legacy three-category generator owns the established root-hub
    # markup. Its hub now reads the shared ten-entry catalog, so reuse only
    # that pure renderer without invoking its page-generation main().
    original_argv = sys.argv[:]
    try:
        sys.argv = [sys.argv[0], "high"]
        import generate_highschool_korean_english_math as legacy_generator
    finally:
        sys.argv = original_argv
    (TARGET_ROOT / "index.html").write_text(legacy_generator.render_root_hub(), encoding="utf-8")

    if not args.skip_production_files:
        from prepare_production_domain import main as prepare_production_domain
        from normalize_internal_links_and_social_meta import normalize_site
        from postprocess_center_entities import process as postprocess_entities

        prepare_production_domain()
        normalization = normalize_site(ROOT, apply=True)
        postprocess_entities()
        prepare_production_domain()
        normalization = normalize_site(ROOT, apply=True)
        report["normalization"] = normalization
    report["category_hubs"] = len(selected)
    report["generated_categories"] = [config.slug for config in selected]
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
