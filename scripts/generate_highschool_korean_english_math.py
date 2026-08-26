from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import re
import struct
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path
from urllib.parse import quote
from zipfile import ZipFile

try:
    from .subject_catalog import SUBJECT_CATALOG
except ImportError:
    from subject_catalog import SUBJECT_CATALOG


ROOT = Path(__file__).resolve().parents[1]
COMMON = ROOT.parent / "참고자료" / "공통자료"
USED_MANUSCRIPTS = COMMON.parent / "사용한 원고" / "국영수학원.com 추가 원고"
PROFILE_ARGS = [value.lower() for value in sys.argv[1:] if not value.startswith("-")]
PROFILE = PROFILE_ARGS[0] if PROFILE_ARGS else "high"
PROFILES = {
    "high": {
        "source": USED_MANUSCRIPTS / "고등학생 국영수학원.zip",
        "category": "고등학생 국영수학원", "slug": "고등학생국영수학원",
        "level": "고등학생", "course": "고등", "grade_prefix": "고", "school": "고등학교", "school_column": "타깃학교\n(고)",
        "meta_focus": "고등 내신 자료", "english": "High school", "rep_prefix": "hs-kem",
        "school_materials": "시험 범위표·학교 프린트·수행평가 일정",
        "hub_intro": "고등 과정은 세 과목의 분량을 똑같이 늘리기보다 학교 내신 범위, 모의고사 오답과 수행평가 일정을 함께 놓고 우선순위를 정해야 합니다. 지역별 페이지에서 실제 센터 정보와 학교 참고 자료, 과목별 가능 학년을 확인할 수 있습니다.",
        "hub_focus": "내신·모의고사",
        "hub_process": "시험 뒤 재설계",
        "hub_cta": "고등 과정은 시험 결과를 다음 학습 계획으로 바꾸는 절차부터 확인하세요",
        "hub_cta_body": "최근 시험지와 오답 기록, 학교 시험 일정을 준비하면 국어·영어·수학의 보완 순서와 주간 시간 배분을 구체적으로 비교할 수 있습니다.",
        "root_card": "학교 내신 범위와 모의고사 오답, 수행평가 일정을 함께 살펴 세 과목의 우선순위를 정합니다.",
    },
    "middle": {
        "source": USED_MANUSCRIPTS / "중학생 국영수학원.zip",
        "category": "중학생 국영수학원", "slug": "중학생국영수학원",
        "level": "중학생", "course": "중등", "grade_prefix": "중", "school": "중학교", "school_column": "타깃학교\n(중)",
        "meta_focus": "중등 내신 자료", "english": "Middle school", "rep_prefix": "ms-kem",
        "school_materials": "시험 범위표·학교 프린트·수행평가 일정",
        "hub_intro": "중등 과정은 교과 개념의 빈틈과 학교 시험 준비를 구분하고, 과제 수행과 오답 복습이 스스로 이어지는지를 살펴야 합니다. 지역별 센터의 실제 운영 학년과 학교 참고 자료를 확인한 뒤 학생에게 맞는 관리 흐름을 비교하세요.",
        "hub_focus": "개념·내신 연결",
        "hub_process": "주간 실행 점검",
        "hub_cta": "중등 과정은 개념 이해와 학교 시험 준비가 한 주 안에서 연결되는지 확인하세요",
        "hub_cta_body": "최근 교재와 시험지, 완료하지 못한 과제를 준비하면 과목별 개념 빈틈과 복습 순서를 나누고 실행 가능한 주간 계획을 비교할 수 있습니다.",
        "root_card": "교과 개념의 빈틈, 학교 시험 준비와 과제 실행을 구분해 중등 학습의 주간 흐름을 확인합니다.",
    },
    "elementary": {
        "source": USED_MANUSCRIPTS / "초등학생 국영수학원.zip",
        "category": "초등학생 국영수학원", "slug": "초등학생국영수학원",
        "level": "초등학생", "course": "초등", "grade_prefix": "초", "school": "초등학교", "school_column": "타깃학교\n(초)",
        "meta_focus": "교과 진도와 기초 학습 자료", "english": "Elementary school", "rep_prefix": "es-kem",
        "school_materials": "교과 진도·알림장·단원평가 자료",
        "hub_intro": "초등 과정은 진도를 앞당기기보다 읽기 이해, 어휘 활용과 계산 과정을 확인하고 스스로 공부를 시작하는 습관을 만드는 것이 우선입니다. 지역별 페이지에서 실제 센터 정보와 교과 자료, 과목별 가능 학년을 확인하세요.",
        "hub_focus": "기초·교과 이해",
        "hub_process": "공부 습관 형성",
        "hub_cta": "초등 과정은 진도보다 기초 이해와 스스로 시작하는 공부 습관을 먼저 확인하세요",
        "hub_cta_body": "현재 교재와 알림장, 단원평가 자료를 준비하면 읽기·어휘·계산 과정에서 막히는 지점과 가정에서 확인할 짧은 학습 기록을 구체적으로 비교할 수 있습니다.",
        "root_card": "읽기 이해와 어휘, 계산 과정과 학습 시작 습관을 중심으로 초등 교과 기초를 점검합니다.",
    },
}
if PROFILE not in PROFILES:
    raise SystemExit(f"Unknown profile: {PROFILE}. Choose high, middle, or elementary.")
CONFIG = PROFILES[PROFILE]
SOURCE_ZIP = CONFIG["source"]
CENTER_CSV = COMMON / "센터정보 정리.csv"
IMAGE_CSV = COMMON / "이미지링크.csv"
REP_SOURCE = COMMON / "대표이미지"
REP_TARGET = ROOT / "assets" / "representative"
MAP_DIR = ROOT / "assets" / "maps"
TARGET_ROOT = ROOT / "과목별학원"
CATEGORY_NAME = CONFIG["category"]
CATEGORY_SLUG = CONFIG["slug"]
CATEGORY_ROOT = TARGET_ROOT / CATEGORY_SLUG
LEVEL_NAME = CONFIG["level"]
COURSE_NAME = CONFIG["course"]
GRADE_PREFIX = CONFIG["grade_prefix"]
SCHOOL_NAME = CONFIG["school"]
SCHOOL_COLUMN = CONFIG["school_column"]
META_FOCUS = CONFIG["meta_focus"]
ENGLISH_LEVEL = CONFIG["english"]
REP_PREFIX = CONFIG["rep_prefix"]
SCHOOL_MATERIALS = CONFIG["school_materials"]
HUB_INTRO = CONFIG["hub_intro"]
HUB_FOCUS = CONFIG["hub_focus"]
HUB_PROCESS = CONFIG["hub_process"]
HUB_CTA = CONFIG["hub_cta"]
HUB_CTA_BODY = CONFIG["hub_cta_body"]
SITE_NAME = "와와학습코칭센터"
TITLE_SUFFIX = "와와학습코칭센터 영어수학 전문학원"
PHONE = "010-6839-8283"
SITE_ORIGIN = "https://xn--3e0bz50b1zcyxat54c.com"
DATE_PUBLISHED = "2026-07-30"
DATE_MODIFIED = "2026-08-17"
REQUIRED = ("페이지타이틀", "메타설명", "본문", "FAQ", "학부모후기", "JSON-LD 요약")
REGION_ORDER = ["서울", "경기", "인천", "충청", "대전", "대구", "울산", "부산", "경상", "광주", "전라", "강원", "제주"]
OFFICIAL_REGION_NAMES = {
    "서울": "서울특별시",
    "경기": "경기도",
    "인천": "인천광역시",
    "대전": "대전광역시",
    "대구": "대구광역시",
    "울산": "울산광역시",
    "부산": "부산광역시",
    "광주": "광주광역시",
    "강원": "강원특별자치도",
    "제주": "제주특별자치도",
}
ADDRESS_REGION_NAMES = {
    "충북": "충청북도",
    "충청북도": "충청북도",
    "충남": "충청남도",
    "충청남도": "충청남도",
    "세종": "세종특별자치시",
    "세종특별자치시": "세종특별자치시",
    "경북": "경상북도",
    "경상북도": "경상북도",
    "경남": "경상남도",
    "경상남도": "경상남도",
    "전북": "전북특별자치도",
    "전북특별자치도": "전북특별자치도",
    "전남": "전라남도",
    "전라남도": "전라남도",
}
TRANSITION_PREFIXES = (
    "상담 기록을 기준으로 보면, ",
    "최근 자료를 함께 놓고 보면, ",
    "학생의 실제 실행을 살펴보면, ",
    "과목별 점검 순서를 정할 때, ",
    "다음 학습 계획을 세우려면, ",
    "가정에서 관찰한 내용을 더하면, ",
    "최근 교재와 오답을 대조하면, ",
    "주간 학습 흐름을 확인하면, ",
    "시험·교과 자료를 살펴보면, ",
    "완료 기록을 중심으로 보면, ",
    "상담 질문을 구체화하려면, ",
    "학생의 현재 시간표를 고려하면, ",
    "학교 자료를 대조할 때, ",
    "학생 자료를 먼저 펼쳐 보면, ",
    "실제 학습 범위를 확인하면, ",
    "상담 전에 정리해 보면, ",
    "센터 안내와 비교할 때, ",
    "과목별 기록을 살펴보면, ",
    "확인 순서를 세울 때, ",
    "상담 자료를 정리하면, ",
    "학교 일정까지 함께 보면, ",
    "다음 계획을 정하기 전에, ",
    "가정에서 기록을 준비하면, ",
    "학생의 현재 자료를 기준으로 보면, ",
)


def esc(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def stable_pick(seed: str, label: str, choices: list[str]) -> str:
    digest = hashlib.sha256(f"{seed}|{label}".encode("utf-8")).hexdigest()
    return choices[int(digest[:10], 16) % len(choices)]


def parse_sections(text: str) -> dict[str, str]:
    marker = re.compile(r"^\[([^\]]+)\]\s*$", re.MULTILINE)
    matches = list(marker.finditer(text))
    sections: dict[str, str] = {}
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        sections[match.group(1).strip()] = text[match.end():end].strip()
    return sections


def parse_body(body: str) -> tuple[str, list[tuple[str, list[str]]]]:
    heading = re.compile(r"^##\s+(.+?)\s*$", re.MULTILINE)
    matches = list(heading.finditer(body))
    intro = body[: matches[0].start()].strip() if matches else body.strip()
    sections: list[tuple[str, list[str]]] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        paragraphs = [clean(part) for part in re.split(r"\n\s*\n", body[match.end():end]) if clean(part)]
        sections.append((clean(match.group(1)), paragraphs))
    return clean(intro), sections


def locality_from_title(title: str) -> str:
    suffix = f" {CATEGORY_NAME}"
    if not title.endswith(suffix):
        raise ValueError(f"Unexpected title: {title}")
    return title[:-len(suffix)].strip()


def folder_slug(locality: str) -> str:
    return re.sub(r"\s+", "", locality)


def page_path(locality_slug: str | None = None) -> str:
    base = f"/과목별학원/{CATEGORY_SLUG}/"
    return base + (locality_slug + "/" if locality_slug else "")


def canonical_url(locality_slug: str | None = None) -> str:
    return SITE_ORIGIN + quote(page_path(locality_slug), safe="/")


def absolute_site_url(value: str) -> str:
    if value.startswith(("http://", "https://")):
        return value
    return SITE_ORIGIN + quote("/" + value.lstrip("/"), safe="/%:@?&=+-._~")


def list_values(value: str) -> list[str]:
    return list(dict.fromkeys(part.strip() for part in re.split(r"[,/\n]+", value or "") if part.strip()))


def school_values(value: str) -> list[str]:
    """Parse only explicit school names from the supplied center worksheet."""
    values: list[str] = []
    for group in re.split(r"[,，/·.\r\n]+", value or ""):
        group = clean(group)
        if not group or re.fullmatch(r"지역\s*내\s*모든\s*(?:초등|중|고등)?학교\s*가능", group):
            continue
        tokens = group.split()
        if len(tokens) > 1 and all(re.search(r"(?:초등학교|중학교|고등학교|초|중|고)$", token) for token in tokens):
            values.extend(tokens)
        else:
            values.append(group)
    return list(dict.fromkeys(values))


def image_dimensions(path: Path) -> tuple[int, int]:
    """Read published image dimensions without an optional imaging dependency."""
    data = path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:6] in {b"GIF87a", b"GIF89a"} and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data.startswith(b"\xff\xd8"):
        offset = 2
        while offset + 9 < len(data):
            if data[offset] != 0xFF:
                offset += 1
                continue
            marker = data[offset + 1]
            offset += 2
            if marker in {0xD8, 0xD9} or 0xD0 <= marker <= 0xD7:
                continue
            if offset + 2 > len(data):
                break
            size = int.from_bytes(data[offset:offset + 2], "big")
            if size < 2 or offset + size > len(data):
                break
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                return (
                    int.from_bytes(data[offset + 5:offset + 7], "big"),
                    int.from_bytes(data[offset + 3:offset + 5], "big"),
                )
            offset += size
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        kind = data[12:16]
        if kind == b"VP8X" and len(data) >= 30:
            return 1 + int.from_bytes(data[24:27], "little"), 1 + int.from_bytes(data[27:30], "little")
        if kind == b"VP8 " and len(data) >= 30:
            return int.from_bytes(data[26:28], "little") & 0x3FFF, int.from_bytes(data[28:30], "little") & 0x3FFF
        if kind == b"VP8L" and len(data) >= 25:
            bits = int.from_bytes(data[21:25], "little")
            return (bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1
    raise ValueError(f"Unsupported image header: {path}")


def image_mime_type(path: Path | str) -> str:
    suffix = Path(path).suffix.lower()
    mime = {
        ".gif": "image/gif",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".webp": "image/webp",
    }.get(suffix)
    if not mime:
        raise ValueError(f"Unsupported social image extension: {path}")
    return mime


def official_region_name(region: str, address: str) -> str:
    """Return a reader-facing official province or metropolitan-city name.

    The source worksheet intentionally groups several provinces under broad
    directory labels such as ``충청``, ``경상`` and ``전라``. Those labels
    remain untouched for hub grouping and source identity, but are not valid
    province names when joined to a city in prose. Split only those broad
    groups using the verified center address and use official names elsewhere.
    """
    if region in {"충청", "경상", "전라"}:
        prefix = clean(address).split(" ", 1)[0]
        official = ADDRESS_REGION_NAMES.get(prefix)
        if not official:
            raise ValueError(f"Cannot derive official region from address: {region!r} / {address!r}")
        return official
    official = OFFICIAL_REGION_NAMES.get(region)
    if not official:
        raise ValueError(f"Unknown source region: {region!r}")
    return official


def display_locality_name(locality: str, region: str, district: str) -> str:
    """Drop a repeated city prefix only in reader-facing administrative joins."""
    prefixes = [re.sub(r"(?:시|군|구)$", "", clean(district)), clean(region)]
    for prefix in dict.fromkeys(value for value in prefixes if value):
        match = re.fullmatch(rf"{re.escape(prefix)}\s+(.+)", clean(locality))
        if match:
            return clean(match.group(1))
    return clean(locality)


def target_grades(value: str) -> list[str]:
    return [grade for grade in list_values(value) if grade.startswith(GRADE_PREFIX)]


def load_centers() -> dict[str, dict]:
    with IMAGE_CSV.open(encoding="utf-8-sig", newline="") as handle:
        image_rows = {clean(row["제목"]): row for row in csv.DictReader(handle)}
    if len(image_rows) != 371:
        raise ValueError(f"Expected 371 image rows, found {len(image_rows)}")
    with CENTER_CSV.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    centers: dict[str, dict] = {}
    for row in rows:
        locality = clean(row["근처 수업가능 동네"])
        region = clean(row["지역"])
        address = clean(row["센터 주소"])
        english_slug = clean(row["동 영어"]).replace(" ", "-")
        candidates = sorted(MAP_DIR.glob(english_slug + ".*"))
        if not candidates:
            raise FileNotFoundError(f"Map missing: {locality} / {english_slug}")
        map_width, map_height = image_dimensions(candidates[0])
        image_row = image_rows.get(locality)
        if not image_row:
            raise ValueError(f"Image row missing: {locality}")
        body_image = "seoul6839.webp" if clean(image_row["본문"]).lower() == "seoul.jpg" else "local6839.webp"
        display_region = official_region_name(region, address)
        district = clean(row["시or구"])
        display_locality = display_locality_name(locality, region, district)
        centers[locality] = {
            "locality": locality,
            "english_slug": english_slug,
            "region": region,
            "display_region": display_region,
            "district": district,
            "display_district": "" if display_region == "세종특별자치시" else district,
            "display_locality": display_locality,
            "center_name": clean(row["센터명"]),
            "tuition_url": clean(row["센터 교습비"]),
            "office_name": clean(row["교육지원청명칭"]),
            "registration": clean(row["교육지원청 등록번호"]),
            "address": address,
            "schools": school_values(row[SCHOOL_COLUMN]),
            "korean_grades": target_grades(row["가능학년\n(국어)"]),
            "english_grades": target_grades(row["가능학년\n(영어)"]),
            "math_grades": target_grades(row["가능학년\n(수학)"]),
            "map_name": candidates[0].name,
            "map_width": map_width,
            "map_height": map_height,
            "body_image": body_image,
        }
    if len(centers) != 371:
        raise ValueError(f"Expected 371 centers, found {len(centers)}")
    return centers


def load_manuscripts() -> list[dict[str, str]]:
    if not SOURCE_ZIP.exists():
        raise FileNotFoundError(SOURCE_ZIP)
    manuscripts: list[dict[str, str]] = []
    with ZipFile(SOURCE_ZIP) as archive:
        for info in archive.infolist():
            if not info.filename.lower().endswith(".txt"):
                continue
            sections = parse_sections(archive.read(info).decode("utf-8-sig"))
            missing = [key for key in REQUIRED if not sections.get(key)]
            if missing:
                raise ValueError(f"{info.filename}: missing {missing}")
            title = clean(sections["페이지타이틀"])
            expected = Path(info.filename).stem
            if title != expected:
                raise ValueError(f"Title mismatch: {info.filename} / {title}")
            manuscripts.append(sections)
    manuscripts.sort(key=lambda item: clean(item["페이지타이틀"]))
    if len(manuscripts) != 371:
        raise ValueError(f"Expected 371 manuscripts, found {len(manuscripts)}")
    return manuscripts


def choose_representatives(count: int) -> list[str]:
    existing: list[str] = []
    for index in range(1, count + 1):
        matches = sorted(REP_TARGET.glob(f"{REP_PREFIX}-{index:03d}.*"))
        if len(matches) != 1:
            raise ValueError(
                f"Expected one stable representative for {REP_PREFIX}-{index:03d}, "
                f"found {len(matches)}"
            )
        existing.append(matches[0].name)
    return existing


def compact_meta(title: str, center: dict) -> str:
    value = clean(
        f"{title} 선택 전 {META_FOCUS}, 최근 오답과 과목별 가능 학년, "
        "센터 위치·비용 확인 기준을 안내합니다."
    )
    if not 65 <= len(value) <= 80:
        raise ValueError(f"Meta length outside 65..80: {len(value)} / {title}")
    return value


def hangul_jongseong(value: str) -> int | None:
    """Return the final-consonant index for a phrase without rewriting it.

    조사 교정은 문맥 없이 문장 전체에 적용하면 `흥덕마을`의 `을`이나
    `영창로`의 `로`를 조사로 잘못 인식한다. 이 함수는 새로 넣는 안전한
    상담 주제에 조사를 붙일 때만 사용한다.
    """
    last = next((char for char in reversed(value) if "가" <= char <= "힣"), None)
    return None if last is None else (ord(last) - 0xAC00) % 28


def attach_particle(value: str, particle: str) -> str:
    pairs = {
        "을": ("을", "를"), "를": ("을", "를"),
        "은": ("은", "는"), "는": ("은", "는"),
        "이": ("이", "가"), "가": ("이", "가"),
        "과": ("과", "와"), "와": ("과", "와"),
    }
    jongseong = hangul_jongseong(value)
    if particle in {"으로", "로"}:
        return value + ("로" if jongseong in {0, 8} else "으로")
    if particle not in pairs or jongseong is None:
        return value + particle
    consonant, vowel = pairs[particle]
    return value + (consonant if jongseong else vowel)


def replace_keyword_token(text: str, keyword: str, topic: str) -> str:
    """Replace one manuscript-only SEO keyword while preserving its particle."""
    # 원고에는 `온라인수업처럼`, `학원온라인수업를`처럼 키워드가 다른
    # 음절과 붙어 있는 변형도 있다. 먼저 토큰 자체를 빠짐없이 치환한 뒤,
    # 새 상담 주제에 바로 붙은 조사만 한정적으로 바로잡는다.
    text = text.replace(f"학원{keyword}", topic).replace(keyword, topic)
    for particle in ("으로", "로", "을", "를", "은", "는", "이", "가", "과", "와"):
        text = text.replace(f"{topic}{particle}", attach_particle(topic, particle))
    return text


def protect_verified_facts(text: str, center: dict) -> tuple[str, dict[str, str]]:
    """Hide authoritative literals while manuscript-only keywords are replaced."""
    values = {
        center.get("address", ""),
        center.get("center_name", ""),
        center.get("office_name", ""),
        center.get("registration", ""),
    }
    placeholders: dict[str, str] = {}
    index = 0
    for value in sorted((clean(item) for item in values if clean(item)), key=len, reverse=True):
        token = f"@@VERIFIED_FACT_{index:03d}@@"
        if value in text:
            text = text.replace(value, token)
            placeholders[token] = value
            index += 1
    for value in sorted((clean(item) for item in center.get("schools", []) if clean(item)), key=len, reverse=True):
        token = f"@@VERIFIED_FACT_{index:03d}@@"
        pattern = re.compile(rf"(?<![가-힣]){re.escape(value)}(?![가-힣])")
        if pattern.search(text):
            text = pattern.sub(token, text)
            placeholders[token] = value
            index += 1
    return text, placeholders


def restore_verified_facts(text: str, placeholders: dict[str, str]) -> str:
    for token, value in placeholders.items():
        text = text.replace(token, value)
    return text


def use_official_region_in_copy(text: str, center: dict) -> str:
    """Replace source taxonomy labels with official reader-facing geography."""
    region = center["region"]
    display_region = center["display_region"]
    locality = center.get("locality", "")
    locality_token = "@@VERIFIED_LOCALITY@@"
    if locality:
        text = text.replace(locality, locality_token)
    if region != display_region:
        text = re.sub(
            rf"(?<![가-힣]){re.escape(region)}(?![가-힣])",
            display_region,
            text,
        )
    # The worksheet's Sejong grouping column contains road names rather than
    # an administrative district. Keep that raw value for routing/grouping,
    # but never present it as ``세종특별자치시 새롬중앙로 다정동`` in prose.
    district = center.get("district", "")
    if district and not center.get("display_district"):
        text = re.sub(
            rf"{re.escape(display_region)}\s+{re.escape(district)}(?=\s)",
            display_region,
            text,
        )
    text = text.replace(locality_token, locality)
    full_raw_join = " ".join(
        part for part in (display_region, center.get("display_district", ""), locality) if part
    )
    full_display_join = " ".join(
        part for part in (
            display_region,
            center.get("display_district", ""),
            center.get("display_locality", locality),
        ) if part
    )
    if full_raw_join != full_display_join:
        text = text.replace(full_raw_join, full_display_join)
    return text


def normalize_address_sentences(text: str, center: dict, locality: str, seed: str) -> str:
    """Use the complete verified CSV address in every explicit address sentence."""
    address = clean(center.get("address", ""))
    if not address:
        return text
    pattern = re.compile(
        r"[^.!?\n]*(?:(?:방문\s+상담\s+)?주소는|제공된\s+주소는|"
        r"주소\s+표기(?:\([^)]*\))?|제공\s+주소|주소\s+기준|센터\s+주소|"
        rf"{re.escape(address)})[^.!?\n]*[.!?]"
    )
    occurrence = 0

    def scope_guidance(value: str) -> str:
        sentences = [part for part in re.split(r"(?<=[.!?])\s+", clean(value)) if part]
        return " ".join(
            sentence if address in sentence else f"{LEVEL_NAME}의 경우 {sentence}"
            for sentence in sentences
        )

    def replace(match: re.Match[str]) -> str:
        nonlocal occurrence
        if occurrence:
            follow_ups = [
                "이동 동선을 다시 확인할 때는 학생의 하교 시각과 귀가 뒤 복습 시간을 함께 계산하세요.",
                "방문 계획은 학교 일정과 실제 교통 시간을 반영해 무리 없는지 살펴보는 편이 좋습니다.",
                "같은 위치를 기준으로 평일과 시험 기간의 등원 가능 시각을 각각 확인해야 합니다.",
                "통학 조건을 비교할 때는 수업 전후 이동과 가정에서 확보할 공부 시간을 같이 보세요.",
                "실제 방문 전 건물 위치와 현재 개설 과목, 가능한 수업 시각을 다시 문의하세요.",
                "학생의 왕복 이동 시간이 과제와 오답 복습을 방해하지 않는지도 함께 점검해야 합니다.",
                "학교에서 수업 장소까지 이동한 뒤 귀가하는 데 필요한 시간을 주간표에 반영하세요.",
                "등록 전에는 통학 가능성과 더불어 과목별 수업 범위와 시간표를 한 번 더 확인하세요.",
            ]
            start = int(hashlib.sha256(f"{seed}|verified-location-follow-up".encode("utf-8")).hexdigest()[:8], 16)
            value = scope_guidance(follow_ups[(start + occurrence) % len(follow_ups)])
            occurrence += 1
            return (" " if match.start() and text[match.start() - 1] in ".!?" else "") + value
        choices = [
            f"제공된 센터 주소는 {address}입니다. 방문 전 학생의 실제 이동 시간과 귀가 뒤 복습 시간을 함께 확인하세요.",
            f"{locality} 방문 상담 주소는 {address}입니다. 학교 일정과 등하원 시간을 놓고 지속 가능한 시간표인지 살펴보는 편이 좋습니다.",
            f"센터 위치는 {address}입니다. 상담 전에는 학생의 통학 시간과 가정에서 확보할 복습 시간을 같이 계산해 보세요.",
            f"제공 자료의 센터 주소는 {address}입니다. 실제 방문 전 주소와 수업 가능 시간을 센터에 다시 확인하는 편이 정확합니다.",
            f"상담 장소의 전체 주소는 {address}입니다. 학생의 학교 일정과 이동 동선도 함께 점검해야 합니다.",
            f"{locality}에서 확인할 센터 주소는 {address}입니다. 등록 전 통학 가능 시간과 센터의 현재 운영 범위를 함께 문의하세요.",
            f"방문 전에 확인할 센터 주소는 {address}입니다. 수업 뒤 귀가와 복습까지 이어질 수 있는 동선인지도 살펴보세요.",
            f"주소는 제공 자료 기준 {address}입니다. 상담 예약 때 위치와 수업 시각을 다시 확인하고 학생의 실제 이동 시간을 계산하세요.",
            f"센터의 전체 주소는 {address}입니다. 학생이 학교에서 이동해 수업을 마친 뒤 귀가하는 데 필요한 시간을 따로 계산하세요.",
            f"상담 전 확인할 주소는 {address}입니다. 평일과 시험 기간의 이동 시간이 달라지는지도 함께 살펴보는 편이 좋습니다.",
            f"제공 주소는 {address}입니다. 학생의 하교 시각과 첫 복습 시작 시각을 놓고 실제 통학 가능성을 확인하세요.",
            f"센터 주소는 {address}입니다. 방문 예약 때 건물 위치와 현재 수업 시각을 다시 문의해야 합니다.",
            f"등록 자료에 적힌 주소는 {address}입니다. 주중 이동 동선과 수업 후 과제를 시작할 수 있는 시간을 같이 점검하세요.",
            f"위치 확인에 사용할 주소는 {address}입니다. 학교 일정과 교통 시간을 반영해 무리 없는 등원 시각인지 살펴보세요.",
            f"상담 장소는 {attach_particle(address, '로')} 안내되어 있습니다. 방문 전 센터 위치와 학생의 왕복 이동 시간을 다시 확인하세요.",
            f"현재 제공된 주소는 {address}입니다. 실제 상담을 예약할 때 위치, 개설 과목과 가능한 수업 시간을 함께 문의하세요.",
        ]
        start = int(hashlib.sha256(f"{seed}|verified-address".encode("utf-8")).hexdigest()[:8], 16)
        value = scope_guidance(choices[(start + occurrence) % len(choices)])
        occurrence += 1
        return (" " if match.start() and text[match.start() - 1] in ".!?" else "") + value

    return pattern.sub(replace, text)


def normalize_topic_particles(text: str, topic: str) -> str:
    """Normalize only particles attached to the generated safe topic."""
    pattern = re.compile(
        rf"{re.escape(topic)}(?P<particle>으로|로|을|를|은|는|이|가|과|와)"
    )
    return pattern.sub(
        lambda match: attach_particle(topic, match.group("particle")),
        text,
    )


def sanitize_unverified_operational_keyword(text: str, seed: str, locality: str) -> str:
    """Turn unsourced service-like keywords into verifiable consultation topics.

    원고의 개별화용 키워드에는 `온라인수업`, `방학캠프`, `입시성공사례`
    같이 센터 자료로 확인되지 않은 운영 표현이 섞여 있다. 서비스 사실로
    보일 수 있는 단어는 최근 시험지·오답·학교 일정처럼 학부모가 실제
    자료로 확인할 수 있는 상담 주제로 바꾼다.
    """
    unsafe_terms = (
        "녹화수업", "온라인수업", "방학캠프", "일대일수업", "야간수업",
        "입시성공사례", "학원자료실", "학습암기",
        "학원매출관리", "학원창업", "학원미납관리", "학원고객관리시스템",
        "학원전자계약", "학원관리솔루션", "학원결제시스템",
        "학원소수정예", "학원출입관리", "학원결제관리", "학원고객관리",
        "학원강사", "학원위치", "학원운영자", "학원운영", "학원행정",
    )
    keyword = next((value for value in unsafe_terms if value in text), "")
    patterns = (
        r"(?m)^##\s+(.+?)까지\s+고려한\s+.+?\s+관리\s+방식\s*$",
        r"수업\s+구조,\s*학교\s+연계,\s*(.{2,40}?)\s+관점의\s+확인사항",
        r"그리고\s+([가-힣A-Za-z0-9·]+)처럼\s+학부모님이\s+실제로\s+묻는\s+운영\s+요소",
        r"가정에서\s+확인할\s+([가-힣A-Za-z0-9·]+)과\s+학습\s+습관",
        r"확인해야\s+([가-힣A-Za-z0-9·]+)\s+운영도",
        r"([가-힣A-Za-z0-9·]+)(?:이|가)\s+상담\s+키워드로\s+제시된",
    )
    if not keyword:
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                keyword = match.group(1)
                break
    if not keyword:
        return text

    topic = stable_pick(seed, "safe-consultation-topic", [
        "최근 평가 기록",
        "오답 재확인 기록",
        "과제 완료 기준",
        "학교 일정 점검",
        "학습 플래너 기록",
        "과목별 복습 기록",
        "현재 교재 진도",
        "주간 학습 시간",
        "질문 기록",
        "재풀이 결과",
        "단원별 이해도",
        "수업 후 복습 시점",
        "시험·교과 범위",
        "과목별 완료 기록",
        "가정 복습 시간",
        "학습 시작 시각",
        "풀이 과정 기록",
        "틀린 이유 분류",
    ])
    text = replace_keyword_token(text, keyword, topic)
    topic_subject = attach_particle(topic, "은")
    topic_object = attach_particle(topic, "을")
    replacements = {
        f"{topic}처럼 학부모님이 실제로 묻는 운영 요소": f"{topic}처럼 상담에서 확인할 자료",
        f"{topic} 운영도 실질적으로 이어집니다": f"{topic}도 실제 학습 계획에 반영할 수 있습니다",
        f"{topic}이 상담 키워드로 제시된": f"{topic_object} 함께 확인하는",
        f"{topic} 안내를 통해": f"{topic_object} 점검해",
        f"{topic}과 연결해": f"{topic_object} 바탕으로",
        f"상담 시 확인할 {topic} 내용": f"상담 시 확인할 {topic}",
        f"{topic}까지 함께 알아보는": f"{topic}도 함께 확인하는",
        f"{topic}을 찾는다면": f"{topic_object} 확인한다면",
        f"{topic}를 찾는다면": f"{topic_object} 확인한다면",
        f"{topic}을 찾는": f"{topic_object} 확인하려는",
        f"{topic}를 찾는": f"{topic_object} 확인하려는",
        f"{topic}에 대한 안내는": f"{topic} 설명은",
        f"{topic}에 대한 안내": f"{topic}에 대한 설명",
        f"{topic}이라는 단어가 보여도": f"{topic_object} 확인할 때도",
        f"{topic}라는 단어가 보여도": f"{topic_object} 확인할 때도",
        f"{topic} 관점의 확인사항": f"{topic_object} 확인할 기준",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = re.sub(
        rf"{re.escape(topic)}(?:이|가)\s+상담\s+키워드로\s+제시된",
        f"{topic_object} 함께 확인하는",
        text,
    )
    text = re.sub(
        rf"{re.escape(topic)}(?:이|가)?라는\s+단어가\s+보여도",
        f"{topic_object} 확인할 때도",
        text,
    )
    keyword_explanation = re.compile(
        rf"{re.escape(topic)}(?:은|는)\s+[^.!?\n]*?핵심을\s+가장\s+직접적으로\s+"
        r"보여\s+주는\s+키워드입니다\."
    )
    text = keyword_explanation.sub(stable_pick(seed, "reader-facing-topic-evidence", [
        f"{topic_subject} {locality} {LEVEL_NAME}의 현재 학습 흐름을 구체적으로 확인하는 자료입니다. 다음 점검 날짜와 함께 살펴보세요.",
        f"{locality} 상담에서는 {topic_object} 학생 자료로 확인하고 과목별 계획에 어떻게 반영할지 질문하세요.",
        f"{topic_object} 살펴보면 {locality} {LEVEL_NAME}이 실제로 끝낸 내용과 다시 확인할 범위를 나눌 수 있습니다.",
        f"{topic_subject} {locality} 학생의 실행 상태를 판단할 근거입니다. 기록 내용이 다음 과제와 복습으로 이어지는지도 확인하세요.",
        f"상담 전 {topic_object} 준비하면 {locality} {LEVEL_NAME}의 현재 단원과 보완 순서를 더 구체적으로 설명할 수 있습니다.",
        f"{locality} 학부모는 {topic_object} 바탕으로 학생의 과목별 진행 상황과 다음 확인 시점을 물어볼 수 있습니다.",
    ]), text)
    search_query_explanation = re.compile(
        rf"{re.escape(topic)}(?:은|는)\s+[^.!?\n]*?보장\s+문구로\s+판단하기보다,\s*"
        r"진단과\s+복습이\s+반복되는\s+구조가\s+있는지\s+묻는\s+검색어로\s+"
        r"보는\s+것이\s+적절합니다\."
    )
    text = search_query_explanation.sub(stable_pick(seed, "reader-facing-topic-process", [
        f"{topic_object} 확인할 때는 진단 결과가 복습과 재점검으로 이어지는지 구체적으로 물어보세요.",
        f"{topic_subject} 결과를 약속하는 말이 아니라 상담에서 실제 기록과 확인 절차를 살필 항목입니다.",
        f"{locality} 상담에서는 {topic_object} 학생 자료와 대조하고 다음 학습 계획에 반영하는 방식을 확인해야 합니다.",
        f"학생의 {topic} 기록을 보면 진단 뒤 보완 과제와 재확인 날짜가 실제로 이어지는지 판단할 수 있습니다.",
        f"상담 전에는 {topic_object} 준비해 현재 상태, 복습 순서와 다음 점검 시점을 함께 확인하세요.",
        f"{topic_subject} {locality} {LEVEL_NAME}의 진단과 복습 과정을 연결해 살펴보는 구체적인 확인 자료입니다.",
    ]), text)
    expression_explanation = re.compile(
        rf"{re.escape(topic)}(?:이라는|라는)\s+표현은\s+[^.!?\n]*?결과를\s+"
        r"보장한다는\s+뜻으로\s+받아들이기보다,\s*과목별\s+준비\s+과정과\s+"
        r"상담\s+기준을\s+점검하는\s+말로\s+보는\s+것이\s+안전합니다\."
    )
    text = expression_explanation.sub(stable_pick(seed, "reader-facing-topic-scope", [
        f"{topic_subject} 결과를 보장하는 내용이 아니라 상담에서 학생의 준비 과정과 점검 시점을 확인할 자료입니다.",
        f"{locality} 상담에서는 {topic_object} 과목별 준비 기록과 대조해 실제 보완 순서를 정해야 합니다.",
        f"{topic_object} 확인할 때는 결과 예측보다 학생이 수행한 내용과 다음 점검 날짜를 구체적으로 살펴보세요.",
        f"학생의 {topic} 기록은 과목별 준비 과정과 복습 흐름을 확인하는 데 활용할 수 있습니다.",
        f"상담 전 {topic_object} 정리하면 {locality} {LEVEL_NAME}의 과목별 진행 상황을 더 정확히 설명할 수 있습니다.",
        f"{topic_subject} 학생의 현재 자료에서 확인하고 센터의 과목별 점검 방식과 대조해야 할 항목입니다.",
    ]), text)
    system_keyword_explanation = re.compile(
        rf"{re.escape(topic)}(?:이|가)?라는\s+키워드는\s+[^.!?\n]*?사람과\s+시스템이\s+"
        r"함께\s+움직이는지를\s+묻는\s+말로\s+바꿔\s+볼\s+수\s+있습니다\."
    )
    text = system_keyword_explanation.sub(stable_pick(seed, "reader-facing-topic-coordination", [
        f"{topic_object} 확인하면 학생의 과제, 복습과 다음 점검이 실제로 이어지는지 살펴볼 수 있습니다.",
        f"{locality} 상담에서는 {topic_object} 학생 자료와 대조하고 다음 학습 행동으로 연결하는 방식을 확인하세요.",
        f"{topic_subject} 학생의 현재 진행 상황과 수업 뒤 확인 절차를 함께 살필 자료입니다.",
        f"상담 전 {topic_object} 준비하면 학생과 담당자가 다음에 확인할 내용을 구체적으로 나눌 수 있습니다.",
        f"{locality} 학부모는 {topic_object} 바탕으로 과목별 실행 내용과 전달 방식을 함께 물어볼 수 있습니다.",
        f"학생의 {topic} 기록이 수업, 과제와 재확인 계획에 반영되는지 상담에서 살펴보세요.",
    ]), text)
    reason_pattern = re.compile(
        rf"{re.escape(locality)}\s+학부모님이\s+{re.escape(topic)}(?:을|를)\s+확인하려는\s+"
        r"이유는\s+단순한\s+수업보다\s+학생의\s+공부\s+방식이\s+바뀌기를\s+"
        r"기대하기\s+때문입니다\."
    )
    text = reason_pattern.sub(stable_pick(seed, "reader-facing-topic-reason", [
        f"{locality} 학부모는 {topic_object} 확인할 때 안내에 그치지 않고 학생의 실제 공부 행동에 어떻게 반영되는지 살펴야 합니다.",
        f"{topic_object} 묻는 목적은 {locality} 학생의 현재 기록이 다음 과제와 복습으로 이어지는지 확인하는 데 있습니다.",
        f"{locality} 상담에서는 {topic_object} 바탕으로 학생이 다음 주에 바꿀 학습 행동을 구체적으로 정해야 합니다.",
        f"학부모가 {topic_object} 확인하면 {locality} 학생의 공부 흐름에서 유지할 부분과 조정할 부분을 나눌 수 있습니다.",
        f"{topic_subject} 수업 설명보다 {locality} 학생의 실행 기록과 다음 점검 계획에서 구체적으로 확인해야 합니다.",
        f"{locality} 학부모에게 필요한 것은 {topic} 안내가 학생의 과제, 복습과 재확인으로 이어지는지 살피는 일입니다.",
    ]), text)
    meta_operational_pattern = re.compile(
        rf"{re.escape(topic)}처럼\s+운영(?:\s+관리)?와\s+관련된\s+표현은"
        r"[^.!?\n]*?(?:영역|투명성)으로\s+해석(?:될|할)\s+수\s+있습니다\."
    )
    text = meta_operational_pattern.sub(stable_pick(seed, "operational-topic-answer", [
        f"{topic_subject} 상담 전에 구체적으로 확인할 항목입니다. 학생의 실제 기록과 센터 안내를 나란히 살펴보세요.",
        f"상담에서는 {topic_object} 학생 자료로 확인하고 다음 학습 계획에 어떻게 반영하는지 질문하세요.",
        f"{topic_subject} 등록 판단보다 먼저 학생의 현재 자료에서 확인하고, 점검 시점과 전달 방식을 구체적으로 물어볼 내용입니다.",
        f"학생에게 필요한 {topic} 기준을 최근 교재와 과제 기록에서 확인한 뒤 센터의 설명과 대조하는 편이 좋습니다.",
        f"{topic_subject} 막연히 판단하지 말고, 상담에서 확인할 자료와 다음 점검 날짜로 구체화해야 합니다.",
        f"상담 전 {topic_object} 정리해 두면 학생의 현재 실행과 센터의 안내를 더 정확히 비교할 수 있습니다.",
    ]), text)
    heading = stable_pick(seed, "safe-seed-heading", [
        f"{locality}에서 {topic_object} 확인하는 방법",
        f"{locality} {LEVEL_NAME}의 주간 학습 기록 점검",
        f"과목별 계획을 조정할 때 보는 {topic}",
        f"{locality} 상담 전 확인할 학습 기록",
        f"{topic_object} 다음 학습에 반영하는 순서",
        f"{locality} 학부모가 상담에서 물어볼 {topic}",
    ])
    text = re.sub(
        rf"(?m)^##\s+{re.escape(topic)}까지\s+고려한\s+.+?\s+관리\s+방식\s*$",
        f"## {heading}",
        text,
    )
    return normalize_topic_particles(text, topic)


def remove_search_engine_language(text: str, locality: str, seed: str) -> str:
    """Replace copy that talks about the query itself with reader-facing guidance."""
    choices = [
        "세 과목을 함께 배우더라도 국어·영어·수학의 진단과 피드백 기준은 과목별로 달라야 합니다.",
        "국어는 답의 근거, 영어는 문장 해석, 수학은 풀이 과정을 따로 확인해야 합니다.",
        "같은 시간표 안에서도 과목별 오답 원인과 재확인 순서는 구분해야 합니다.",
        "세 과목의 분량을 같게 나누기보다 현재 단원과 시험 일정에 따라 우선순위를 정해야 합니다.",
        "과목을 묶어 계획해도 읽기, 해석, 풀이에서 막히는 지점은 각각 기록해야 합니다.",
        "중요한 것은 세 과목을 다룬다는 말보다 과목별 완료 기준과 복습 날짜가 남는지입니다.",
    ]
    pattern = re.compile(
        rf"(?m)^[^.!?\n]*?이\s+{re.escape(CATEGORY_NAME)}이라는\s+이름으로\s+"
        r"검색되더라도\s+세\s+과목을\s+같은\s+방식으로\s+가르친다는\s+뜻은\s+아닙니다\."
    )
    text = pattern.sub(stable_pick(seed, "reader-facing-subject-flow", choices), text)
    comparison_pattern = re.compile(
        rf"[^.!?\n]*?에서\s+{re.escape(CATEGORY_NAME)}을\s+비교할\s+때는\s+“세\s+과목을\s+모두\s+한다”는\s+말보다\s+국어·영어·수학의\s+약점\s+기록이\s+따로\s+남는지를\s+보셔야\s+합니다\."
    )
    comparison_choices = [
        f"{locality}에서는 세 과목의 이름보다 국어·영어·수학별 약점 기록과 다음 확인 날짜를 살펴보세요.",
        f"{locality} 수업을 비교할 때는 과목 수보다 각 과목의 오답 원인과 피드백이 따로 남는지가 중요합니다.",
        f"국어·영어·수학을 함께 계획하더라도 {locality} 학생의 과목별 병목과 복습 기록은 구분해야 합니다.",
        f"{locality} 상담에서는 세 과목을 모두 다룬다는 설명보다 진단 결과가 과목별 행동으로 이어지는지 확인하세요.",
        f"{locality} 학생에게 필요한 것은 과목 묶음보다 국어·영어·수학의 약점과 재점검 기준을 따로 기록하는 과정입니다.",
        f"세 과목 수업을 알아볼 때는 {locality} 학생의 최근 자료에서 과목별 보완 순서가 제시되는지 살펴야 합니다.",
        f"{locality} 학부모는 국어·영어·수학의 오답 기록과 피드백 방식이 과목별로 구분되는지 질문할 수 있습니다.",
        f"과목을 함께 관리해도 {locality} 학생의 읽기·해석·풀이 문제는 별도로 진단하고 기록해야 합니다.",
    ]
    return comparison_pattern.sub(stable_pick(seed, "reader-facing-comparison", comparison_choices), text)


def remove_meta_and_mechanical_copy(text: str, locality: str, seed: str) -> str:
    """Keep the manuscript reader-facing after source and template cleanup."""
    operational_meta = re.compile(
        r"[^.!?\n]*?처럼\s+운영(?:\s+관리)?(?:과|와)\s+관련된\s+표현은"
        r"[^.!?\n]*?(?:영역|투명성)으로\s+해석(?:될|할)\s+수\s+있습니다\."
    )
    text = operational_meta.sub(stable_pick(seed, "reader-facing-operational-check", [
        f"{locality} 상담에서는 비용 안내, 출결 기록과 상담 내용이 학생의 학습 계획에 어떻게 반영되는지 확인하세요.",
        f"등록 전에는 {locality} 센터의 비용 자료와 출결 전달 방식, 상담 기록의 확인 절차를 구체적으로 물어보세요.",
        f"{locality} 학부모는 비용, 출결과 학습 상담 내용이 언제 어떤 방식으로 전달되는지 확인할 수 있습니다.",
        f"센터 운영 정보를 볼 때는 {locality} 학생의 출결과 상담 기록을 학부모가 확인하는 절차를 함께 살펴야 합니다.",
        f"{locality} 상담 전에는 비용 안내의 기준과 출결·학습 기록의 전달 시점을 각각 확인하는 편이 좋습니다.",
        f"비용과 출결 정보를 확인한 뒤 {locality} 학생의 상담 내용이 다음 학습 계획으로 이어지는지도 질문하세요.",
    ]), text)
    location_meta = re.compile(
        r"[^.!?\n]*?페이지에서는\s+지역명을\s+바꾼\s+홍보\s+문구보다\s+"
        r"실제\s+생활\s+동선\s+안에서\s+국어·영어·수학\s+공부가\s+지속되는지를\s+"
        r"판단하는\s+데\s+초점을\s+둡니다\."
    )
    text = location_meta.sub(stable_pick(seed, "reader-facing-life-rhythm", [
        f"{locality} 학생의 국어·영어·수학 계획은 학교 일정과 이동 시간, 귀가 후 복습 시간을 함께 고려해야 합니다.",
        f"{locality}에서는 수업 내용과 함께 실제 등하원 동선과 가정 복습 시간을 놓고 지속 가능한 계획인지 살펴보세요.",
        f"세 과목 학습이 이어지려면 {locality} 학생의 학교 일정, 이동 시간과 과제 분량이 한 주 안에서 맞아야 합니다.",
        f"{locality} 상담에서는 국어·영어·수학 수업이 학생의 생활 리듬 안에서 무리 없이 이어지는지도 확인해야 합니다.",
        f"수업 계획을 정할 때는 {locality} 학생의 등하원 시간과 귀가 뒤 복습 가능 시간을 함께 계산하는 편이 좋습니다.",
        f"{locality} 학생에게 맞는 세 과목 계획인지 판단하려면 학교 일정과 실제 이동 동선을 같이 살펴야 합니다.",
    ]), text)
    middle_intro = re.compile(
        r"이\s+원고는\s+(.+?)을\s+기준으로\s+[^.!?\n]*?학부모가\s+상담\s+전에\s+"
        r"확인할\s+질문을\s+정보성\s+페이지\s+형태로\s+정리했습니다\."
    )

    def replace_middle_intro(match: re.Match[str]) -> str:
        student = clean(match.group(1))
        return stable_pick(seed, "reader-facing-middle-intro", [
            f"{locality} 상담에서는 {student}의 최근 자료를 바탕으로 과목별 약점과 다음 주 실행 순서를 확인해야 합니다.",
            f"{student}이라면 {locality} 상담 전에 최근 교재와 오답, 학교 일정을 준비해 보완 순서를 나누는 편이 좋습니다.",
            f"{locality}의 {student}에게는 현재 단원과 과제 기록을 함께 살펴 과목별로 바꿀 행동을 정하는 과정이 필요합니다.",
            f"상담 전에는 {student}의 최근 시험·과제 자료를 준비해 {locality}에서 확인할 질문을 구체적으로 정리하세요.",
            f"{locality} 학부모는 {student}의 실제 학습 기록을 바탕으로 진단, 복습과 다음 점검 날짜를 물어볼 수 있습니다.",
            f"{student}의 학습 계획은 {locality} 상담에서 과목별 병목과 주간 실행 시간을 나누어 확인하는 데서 시작합니다.",
        ])

    text = middle_intro.sub(replace_middle_intro, text)
    elementary_intro = re.compile(
        rf"이\s+원고는\s+{re.escape(locality)}\s+초등학생을\s+둔\s+학부모가\s+상담\s+전에\s+"
        r"묻는\s+질문에\s+답하도록\s+수업\s+구조,\s*학교\s+연계,\s*"
        r"(.+?)(?:을|를)\s+확인할\s+기준을\s+차례로\s+정리했습니다\."
    )

    def replace_elementary_intro(match: re.Match[str]) -> str:
        topic = clean(match.group(1))
        topic_object = attach_particle(topic, "을")
        return stable_pick(seed, "reader-facing-elementary-intro", [
            f"{locality} 초등 상담에서는 수업 흐름과 학교 자료, {topic_object} 함께 확인해야 합니다.",
            f"{locality} 학부모는 상담 전에 학교 자료와 {topic_object} 준비해 과목별 수업·복습 흐름을 살펴볼 수 있습니다.",
            f"초등 학습을 알아볼 때는 {locality} 학생의 현재 자료, 수업 구조와 {topic_object} 차례로 확인하세요.",
            f"{locality} 초등학생에게 맞는 계획인지 판단하려면 학교 진도와 수업 흐름, {topic_object} 함께 살펴야 합니다.",
            f"상담 전에는 {locality} 학생의 학교 자료와 {topic_object} 정리해 수업 뒤 복습까지 이어지는지 질문하세요.",
            f"{locality} 학부모가 준비할 질문은 과목별 수업 순서, 학교 자료와 {topic_object} 중심으로 정리할 수 있습니다.",
        ])

    text = elementary_intro.sub(replace_elementary_intro, text)
    page_openers = (
        (
            re.compile(r"[^.!?\n]*?페이지를\s+찾은\s+분들이\s+가장\s+궁금해하는\s+부분은\s+([^.!?\n]+?)입니다\."),
            lambda match: f"{locality} 상담에서 먼저 확인할 부분은 {clean(match.group(1))}입니다.",
        ),
        (
            re.compile(r"[^.!?\n]*?페이지의\s+핵심\s+답변은\s+명확합니다\."),
            lambda _match: f"{locality} 상담은 학생의 최근 학습 기록과 과목별 가능 범위를 확인하는 데서 시작합니다.",
        ),
        (
            re.compile(r"[^.!?\n]*?(?:을|를)\s+검색한\s+학부모가\s+바로\s+알고\s+싶은\s+답은\s+분명합니다\."),
            lambda _match: f"{locality} 상담에서 먼저 볼 내용은 현재 교재, 오답과 과목별 가능 범위입니다.",
        ),
        (
            re.compile(r"이\s+글은\s+[^.!?\n]*?(?:을|를)\s+검색한\s+학부모가\s+바로\s+확인할\s+수\s+있도록\s+수업\s+방향,\s*학교\s+진도,\s*학년별\s+체크포인트를\s+순서대로\s+설명합니다\."),
            lambda _match: f"{locality} 상담 전에는 수업 방향, 학교 진도와 학년별 확인 항목을 차례로 정리하세요.",
        ),
    )
    for pattern, replacement in page_openers:
        text = pattern.sub(replacement, text)
    # Search/query and page-production commentary addresses the writer rather
    # than the parent reading the guide. Replace the remaining bounded source
    # frames with direct consultation guidance.
    search_meta = re.compile(
        r"[^.!?\n]{0,180}?(?:을|를)\s+검색(?:한|하는|했다면)[^.!?\n]*[.!?]"
    )
    text = search_meta.sub(stable_pick(seed, "reader-facing-search-meta", [
        f"{locality} 상담에서는 학생의 최근 자료와 센터의 실제 가능 범위를 먼저 대조하세요.",
        f"{locality} 학부모는 현재 교재와 오답을 준비해 과목별 보완 순서를 질문할 수 있습니다.",
        f"상담 전에는 {locality} 학생의 학교 일정과 주간 복습 시간을 함께 정리하는 편이 좋습니다.",
        f"{locality} 학생에게 필요한 수업인지 판단하려면 최근 풀이와 과제 기록부터 살펴야 합니다.",
        f"센터 설명은 {locality} 학생의 현재 단원과 실행 가능한 시간표에 맞춰 확인하세요.",
        f"{locality} 상담의 출발점은 점수 예상보다 학생이 실제로 끝낸 내용과 막힌 지점을 나누는 일입니다.",
    ]), text)
    manuscript_claim = re.compile(
        r"[^.!?\n]*?원고는\s+성적\s+향상을\s+약속하는\s+표현보다\s+현재\s+학습\s+상태를\s+"
        r"어떻게\s+읽고\s+다음\s+주\s+학습으로\s+연결할지에\s+초점을\s+둡니다\."
    )
    text = manuscript_claim.sub(stable_pick(seed, "reader-facing-progress-scope", [
        f"{locality} 상담에서는 결과를 약속하는 말보다 현재 학습 기록을 읽고 다음 주 행동으로 연결하는 절차를 확인해야 합니다.",
        f"{locality} {LEVEL_NAME}의 최근 자료에서 약점을 찾고 다음 주 과제와 복습으로 이어지는지 살펴보세요.",
        f"중요한 것은 결과 예측보다 {locality} 학생의 현재 상태를 과목별 실행 계획으로 바꾸는 과정입니다.",
        f"{locality} 상담은 점수 약속보다 학생의 현재 기록과 다음 점검 계획을 구체적으로 확인하는 데 초점을 둡니다.",
        f"{locality} {LEVEL_NAME}에게는 현재 학습 상태를 읽고 다음 주에 바꿀 행동을 정하는 과정이 필요합니다.",
        f"상담에서는 {locality} 학생의 최근 교재와 오답을 바탕으로 다음 학습 순서를 구체화해야 합니다.",
        f"{locality} 학부모는 결과를 단정하는 설명보다 진단 내용이 다음 주 계획에 반영되는지 확인할 수 있습니다.",
        f"현재 상태를 과목별로 나누고 실행 가능한 다음 행동을 정하는 것이 {locality} 상담의 핵심입니다.",
    ]), text)
    replacements = {
        "학습 점검의 내신 대비는": "학습 점검에서 내신 대비는",
        "학습 계획의 내신 대비는": "학습 계획에서 내신 대비는",
        "학습 점검 시험 후에는": "학습 점검에서는",
        "학습 계획 시험 후에는": "학습 계획에서는",
        "학습 과정 시험 후에는": "학습 과정에서는",
        "세 과목 학습 상담 시험 후에는": "세 과목 학습 상담에서는",
        "관리 기준 시험 후에는": "관리 기준에서는",
        "고려하면, 그리고 ": "고려하면, ",
        "살펴보면, 그리고 ": "살펴보면, ",
        ", 그리고 ": ", ",
        "과제 완료 기준을 기준으로": "과제 완료 기준에 따라",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    advertising_replacements = {
        "광고처럼 빠른 변화를 약속하기보다": "빠른 변화를 기대하기보다",
        "광고 문구가 아니라": "막연한 설명이 아니라",
        "광고 문구보다": "막연한 설명보다",
        "광고의 크기가 아니라": "설명의 양이 아니라",
        "광고보다": "설명보다",
    }
    for source, target in advertising_replacements.items():
        text = text.replace(source, target)
    text = text.replace("페이지에서는", "상담에서는")
    text = text.replace("페이지에는", "제공 자료에는")
    text = text.replace("페이지에서", "상담에서")
    text = text.replace("페이지의", "제공 자료의")
    text = text.replace("페이지에", "제공 자료에")
    text = text.replace(
        f"{LEVEL_NAME}에게는 {LEVEL_NAME} 내신은",
        f"{LEVEL_NAME}의 내신은",
    )
    text = re.sub(
        r"학생의 현재 시간표를 고려하면,\s*([^.!?\n]+?)의 학습 환경을 고려하면",
        r"\1의 학습 환경과 학생의 현재 시간표를 함께 보면",
        text,
    )
    subject_consult = re.compile(
        r"(학습\s+(?:점검|계획|과정|상담))의\s+(국어|영어|수학)\s+상담에서는"
    )
    text = subject_consult.sub(
        lambda match: f"{match.group(1)}에서 {match.group(2)} 학습을 확인할 때는",
        text,
    )
    consultation_transition = stable_pick(seed, "natural-consultation-transition", [
        " 상담 전에는 학생의 실제 자료와 수업 흐름을 차례로 확인해야 합니다.",
        " 상담을 준비할 때는 학생의 최근 기록을 기준으로 질문을 정리하세요.",
        " 상담 전 확인 기준은 학생의 최근 교재와 오답에서 출발해야 합니다.",
        " 상담에서는 현재 학습 상태와 다음 점검 계획을 구분해 살펴야 합니다.",
        " 상담 전에는 과목별 약점과 주간 실행 시간을 함께 확인하는 편이 좋습니다.",
        " 상담을 앞두고 학생 자료와 센터의 과목·학년 범위를 먼저 대조하세요.",
        " 상담 전 질문은 최근 학습 기록과 다음 주 행동을 중심으로 정리할 수 있습니다.",
        " 상담에서는 학생이 끝낸 내용과 다시 확인할 범위를 차례로 물어보세요.",
    ])
    text = text.replace(" 상담 준비에서는 다음 기준을 적용합니다:", consultation_transition)
    comparison_transition = stable_pick(seed, "natural-comparison-transition", [
        "을 비교할 때는 학생의 실제 자료와 맞는지 먼저 살펴야 합니다.",
        "을 알아볼 때는 설명보다 학생의 현재 기록에 적용되는지를 확인하세요.",
        "을 비교한다면 최근 교재와 오답을 기준으로 수업 흐름을 살펴보세요.",
        "을 선택하기 전에는 진단 결과가 다음 계획으로 이어지는지 확인해야 합니다.",
        "을 살펴볼 때는 과목별 점검 내용과 재확인 날짜를 함께 물어보세요.",
        "을 비교할 때는 학생의 시간표 안에서 실행 가능한 계획인지 확인하세요.",
        "을 알아볼 때는 수업 설명이 실제 과제와 복습 기록에 반영되는지 살펴야 합니다.",
        "을 선택할 때는 현재 학습 상태와 센터의 가능 범위를 구체적으로 대조하세요.",
    ])
    text = text.replace(
        "을 비교할 때 화려한 설명보다 중요한 기준은 다음과 같습니다:",
        comparison_transition,
    )
    repeated_followups = [
        "첫 상담에서는 최근 시험지, 학교 프린트, 풀다 멈춘 문제집을 함께 가져오면 학생의 현재 위치를 더 구체적으로 볼 수 있습니다.",
        "학부모님이 자주 묻는 비용이나 일정도 중요하지만, 학생에게 맞는 반 배정 기준을 먼저 확인해야 합니다.",
        "상담 질문은 “몇 점까지 오르나요”보다 “어떤 단원부터 다시 확인하나요”에 가까울수록 실질적인 답을 얻기 쉽습니다.",
        "학생이 싫어하는 과목을 무조건 늘리기보다, 부담을 견딜 수 있는 주간 학습량부터 정하는 것이 현실적입니다.",
        "국어·영어·수학을 한곳에서 관리하더라도 과목별 선생님 피드백이 따로 기록되는지 살펴보는 편이 좋습니다.",
        "진단 결과가 단순 레벨명으로 끝나지 않고 교재, 과제량, 보충 방향으로 바뀌는지 살펴봐야 합니다.",
    ]
    localized_followups = [
        f"첫 상담에서는 {locality} 학생의 최근 시험지, 학교 프린트와 풀다 멈춘 문제집을 함께 보면 현재 위치를 더 구체적으로 알 수 있습니다.",
        f"학부모님이 자주 묻는 비용이나 일정도 중요하지만, {locality} 학생에게 맞는 반 배정 기준을 먼저 확인해야 합니다.",
        f"상담 질문은 ‘몇 점까지 오르나요’보다 ‘{locality} 학생은 어떤 단원부터 다시 확인해야 하나요’에 가까울수록 구체적인 답을 얻기 쉽습니다.",
        f"{locality} 학생이 싫어하는 과목을 무조건 늘리기보다 감당할 수 있는 주간 학습량부터 정하는 편이 현실적입니다.",
        f"{locality} 수업에서 국어·영어·수학을 함께 관리하더라도 과목별 피드백이 따로 기록되는지 살펴보는 편이 좋습니다.",
        f"{locality} 학생의 진단 결과가 레벨명으로 끝나지 않고 교재, 과제량과 보충 방향으로 이어지는지 살펴봐야 합니다.",
    ]
    for sentence, replacement in zip(repeated_followups, localized_followups):
        text = text.replace(sentence, replacement)
    text = re.sub(
        rf"(?:{re.escape(LEVEL_NAME)}\s*[·,/]\s*)+{re.escape(LEVEL_NAME)}",
        f"{COURSE_NAME} 과정",
        text,
    )
    return text


def reduce_phrase_repetition(text: str, phrase: str, seed: str, keep: int = 1) -> str:
    if not phrase:
        return text
    seen = 0
    references = [
        "이 학생",
        "이런 어려움을 보이는 학생",
        "같은 문제를 겪는 학생",
        "이 상담 사례의 학생",
        "비슷한 학습 흐름을 보이는 학생",
        "보완할 부분이 있는 학생",
        "같은 학습 고민이 있는 학생",
        "이와 비슷한 상황의 학생",
        "이 학습 장면에 해당하는 학생",
        "같은 실행 문제를 보이는 학생",
        "점검할 내용이 있는 학생",
        "비슷한 공부 습관을 가진 학생",
    ]
    start = int(hashlib.sha256(f"{seed}|student-reference".encode("utf-8")).hexdigest()[:8], 16)

    def replace(match: re.Match[str]) -> str:
        nonlocal seen
        seen += 1
        if seen <= keep:
            return match.group(0)
        return references[(start + seen - keep - 1) % len(references)]

    return re.sub(re.escape(phrase), replace, text)


def extract_elementary_student_core(student_type: str) -> str:
    """Return the concrete learning behavior after the source grade frame."""
    if PROFILE != "elementary":
        return ""
    value = clean(student_type).rstrip("., ")
    if " 중 " not in value or not value.endswith(" 학생"):
        return ""
    return clean(value.split(" 중 ", 1)[1][:-len(" 학생")])


def extract_elementary_grade_frame(student_type: str) -> str:
    """Return the grade-stage words preceding the source's concrete behavior."""
    if PROFILE != "elementary":
        return ""
    value = clean(student_type).rstrip("., ")
    return clean(value.split(" 중 ", 1)[0]) if " 중 " in value else ""


def reduce_elementary_grade_frame_repetition(
    text: str,
    frame: str,
    seed: str,
    keep: int = 1,
) -> str:
    """Avoid repeating a long grade-stage label around shortened situations."""
    if not frame:
        return text
    seen = 0
    references = [
        "이 시기의 초등",
        "비슷한 학년대의 초등",
        "현재 과정의 초등",
        "해당 학년대의 초등",
        "같은 성장 단계의 초등",
        "이 학습 단계의 초등",
        "비슷한 과정에 있는 초등",
        "현재 학년 흐름의 초등",
    ]
    start = int(hashlib.sha256(f"{seed}|elementary-grade-frame".encode("utf-8")).hexdigest()[:8], 16)

    def replace(match: re.Match[str]) -> str:
        nonlocal seen
        seen += 1
        if seen <= keep:
            return match.group(0)
        return references[(start + seen - keep - 1) % len(references)]

    return re.sub(re.escape(frame), replace, text)


def reduce_elementary_course_label_repetition(text: str, locality: str, seed: str) -> str:
    """Shorten repeated local category aliases without changing the H1 or URL."""
    if PROFILE != "elementary":
        return text
    phrase = f"{locality} {LEVEL_NAME} 세 과목 학습 상담"
    seen = 0
    references = [
        f"{locality} 초등 국영수 상담",
        f"{locality}의 세 과목 학습",
        f"{locality} 초등 학습 상담",
        f"{locality} 국어·영어·수학 점검",
        f"{locality}의 초등 국영수 계획",
        f"{locality} 세 과목 상담",
    ]
    start = int(hashlib.sha256(f"{seed}|elementary-course-label".encode("utf-8")).hexdigest()[:8], 16)

    def replace(match: re.Match[str]) -> str:
        nonlocal seen
        seen += 1
        if seen == 1:
            return match.group(0)
        return references[(start + seen - 2) % len(references)]

    return re.sub(re.escape(phrase), replace, text)


def reduce_elementary_student_core_repetition(
    text: str,
    core: str,
    seed: str,
    keep: int = 1,
) -> str:
    """Keep one full elementary learning situation and shorten later mentions.

    Elementary source drafts move the grade label around the same long behavior,
    so replacing only the full student phrase misses most repetitions.  These
    adjective-form references remain grammatical before ``학생``, ``아이`` and
    ``상태`` while avoiding another copy of the original long clause.
    """
    if not core:
        return text
    seen = 0
    references = [
        "비슷한 학습 어려움을 보이는",
        "같은 공부 고민이 있는",
        "이와 비슷한 학습 흐름을 보이는",
        "같은 실행 문제를 겪는",
        "이런 공부 습관을 보이는",
        "비슷한 보완 과제가 남은",
        "같은 지점에서 자주 멈추는",
        "이와 같은 학습 부담을 느끼는",
        "비슷한 복습 문제가 나타나는",
        "과목별로 비슷한 고민이 있는",
        "이런 점검이 도움이 되는",
        "비슷한 공부 장면이 반복되는",
    ]
    start = int(hashlib.sha256(f"{seed}|elementary-student-core".encode("utf-8")).hexdigest()[:8], 16)

    def replace(match: re.Match[str]) -> str:
        nonlocal seen
        seen += 1
        if seen <= keep:
            return match.group(0)
        return references[(start + seen - keep - 1) % len(references)]

    return re.sub(re.escape(core), replace, text)


def reduce_elementary_behavior_repetition(text: str, phrase: str) -> str:
    """Keep one concrete behavior example and shorten its later references."""
    if not phrase:
        return text
    seen = 0
    pattern = re.compile(
        re.escape(phrase)
        + r"(?P<context>\s+(?:상태라면|모습이\s+있다면|경우라면))?"
    )

    def replace(match: re.Match[str]) -> str:
        nonlocal seen
        seen += 1
        if seen == 1:
            return match.group(0)
        if match.group("context"):
            return "비슷한 어려움을 보인다면"
        return "같은 학습 행동을 보이는"

    return pattern.sub(replace, text)


def copy_tokens(value: str) -> list[str]:
    """Tokenize authored Korean copy exactly like the release auditor."""
    return re.findall(r"[가-힣A-Za-z0-9·]+", value)


def copy_windows(value: str, size: int = 8) -> list[tuple[str, ...]]:
    tokens = copy_tokens(value)
    return [
        tuple(tokens[index:index + size])
        for index in range(max(0, len(tokens) - size + 1))
    ]


def deduplicate_manuscript_windows(
    intro: str,
    sections: list[tuple[str, list[str]]],
    locality: str,
    student_type: str,
    seed: str,
) -> tuple[str, list[tuple[str, list[str]]]]:
    """Limit punctuation-normalized manuscript windows to one occurrence."""
    seen: set[tuple[str, ...]] = set(copy_windows(student_type))
    focuses = [
        "최근 교재의 풀이 표시",
        "과목별 오답 원인",
        "과제 시작과 완료 시각",
        "학교 시험 범위",
        "질문을 미룬 지점",
        "며칠 뒤 재풀이 결과",
        "주간 복습 기록",
        "현재 단원의 공백",
        "수업 뒤 설명 내용",
        "가정에서 가능한 공부 시간",
        "과목별 완료 기준",
        "다음 점검 날짜",
        "학교 자료와 교재 진도",
        "정답을 고른 근거",
        "풀이를 멈춘 단계",
        "복습을 시작한 시각",
    ]
    endings = [
        "실제 자료에서 확인해 이번 주에 바꿀 행동을 한 가지 정합니다.",
        "과목별로 나누어 우선순위를 기록합니다.",
        "학교 일정과 대조해 실행 가능한 분량을 정합니다.",
        "수업 전후 기록에서 살펴 재확인 날짜를 남깁니다.",
        "최근 결과와 비교해 다음 과제에 반영합니다.",
        "학생의 설명과 맞춰 보고 보완 순서를 정합니다.",
    ]
    start = int(hashlib.sha256(f"{seed}|manuscript-window".encode("utf-8")).hexdigest()[:8], 16)
    replacement_index = 0

    def replacement_sentence() -> str:
        nonlocal replacement_index
        for _ in range(len(focuses) * len(endings)):
            index = start + replacement_index
            replacement_index += 1
            focus = focuses[index % len(focuses)]
            ending = endings[(index // len(focuses) + replacement_index) % len(endings)]
            candidate = (
                f"{locality} {LEVEL_NAME} 상담에서는 "
                f"{attach_particle(focus, '을')} {ending}"
            )
            if not (set(copy_windows(candidate)) & seen):
                return candidate
        raise ValueError(f"Cannot create unique manuscript guidance for {seed}")

    def process_paragraph(paragraph: str) -> str:
        sentences = [
            clean(value)
            for value in re.split(r"(?<=[.!?])\s+", paragraph)
            if clean(value)
        ]
        result: list[str] = []
        for sentence in sentences:
            windows = copy_windows(sentence)
            if windows and set(windows) & seen:
                sentence = replacement_sentence()
                windows = copy_windows(sentence)
            seen.update(windows)
            result.append(sentence)
        return " ".join(result)

    intro = process_paragraph(intro)
    sections = [
        (heading, [process_paragraph(paragraph) for paragraph in paragraphs])
        for heading, paragraphs in sections
    ]
    return intro, sections


def reduce_locality_repetition(text: str, locality: str, center: dict, keep: int = 16) -> str:
    """Keep factual strings intact while replacing surplus locality mentions."""
    protected = [center.get("address", ""), *center.get("schools", [])]
    placeholders: dict[str, str] = {}
    for index, value in enumerate(sorted(set(item for item in protected if item), key=len, reverse=True)):
        token = f"@@FACT_{index}@@"
        if value in text:
            text = text.replace(value, token)
            placeholders[token] = value
    seen = 0
    particle_map = {
        "에서는": "이 지역에서는", "에서": "이 지역에서", "의": "이 지역의",
        "은": "이 지역은", "는": "이 지역은", "이": "이 지역이", "가": "이 지역이",
        "을": "이 지역을", "를": "이 지역을", "과": "이 지역과", "와": "이 지역과",
        "": "해당 지역",
    }
    geographic_prefix = ""
    if center.get("district"):
        geographic_prefix = (
            rf"(?:(?:{re.escape(center.get('region', ''))}\s+)?"
            rf"{re.escape(center['district'])}\s+)?"
        )
    pattern = re.compile(geographic_prefix + re.escape(locality) + r"(에서는|에서|의|은|는|이|가|을|를|과|와)?")

    def replace(match: re.Match[str]) -> str:
        nonlocal seen
        seen += 1
        return match.group(0) if seen <= keep else particle_map[match.group(1) or ""]

    text = pattern.sub(replace, text)
    for token, value in placeholders.items():
        text = text.replace(token, value)
    return text


def simplify_heading(value: str) -> str:
    replacements = {
        "의 국어·영어·수학 학습 점검 관리 방식": "의 국어·영어·수학 학습 점검",
        "학습 점검 관리 방식": "학습 점검 방식",
        "관리 기준 관리 방식": "관리 기준",
        "학습 과정 관리 방식": "학습 과정 점검",
        "학습 계획 관리 방식": "학습 계획 조정",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = value.replace(
        "고등학생의 국어·영어·수학 학습 점검의 국어·영어·수학 관리 흐름",
        "고등학생의 국어·영어·수학 관리 흐름",
    )
    value = value.replace(
        "국어·영어·수학 학습 계획의 국어·영어·수학 관리 흐름",
        "국어·영어·수학 관리 흐름",
    )
    value = value.replace(
        "고등 국어·영어·수학 관리 기준의 국어·영어·수학 관리 흐름",
        "고등 국어·영어·수학 관리 흐름",
    )
    value = re.sub(
        r"(?:학습 계획|학습 점검|학습 과정|세 과목 학습 상담|관리 기준)(?:이|가)\s+"
        r"필요한\s+학생\s+유형\s+정리",
        "학습 상황 점검",
        value,
    )
    return clean(value)


def shorten_elementary_sentences(text: str) -> str:
    if PROFILE != "elementary":
        return text
    text = re.sub(r"(학생|아이)(?:이고|이며|으로),\s*", r"\1입니다. 또한 ", text)
    text = text.replace("아이인지부터 확인해야 하며,", "아이인지도 확인해야 합니다. ")
    text = text.replace("학생인지부터 확인해야 하며,", "학생인지도 확인해야 합니다. ")
    text = text.replace("확인해야 하며,", "확인해야 합니다. ")
    return text


def representative_situation_sentence(locality: str, seed: str) -> str:
    return stable_pick(seed, "representative-situation-reference", [
        f"{locality} 학생의 현재 학습 상황을 상담 사례로 구체화해 살펴봅니다.",
        f"이 학습 상황을 {locality} 상담에서 확인할 사례로 살펴봅니다.",
        f"{locality}에서 비슷한 어려움을 겪는 학생의 상담 흐름을 살펴봅니다.",
        f"이 사례를 바탕으로 {locality} 학생에게 필요한 점검 순서를 확인합니다.",
        f"{locality} 학생의 실제 공부 장면을 중심으로 상담 기준을 살펴봅니다.",
        f"비슷한 학습 문제를 {locality} 상담에서 어떻게 나눌지 살펴봅니다.",
        f"이 학습 장면을 토대로 {locality} 상담의 확인 순서를 정리합니다.",
        f"{locality} 학생의 어려움을 실제 상담 질문으로 바꾸어 살펴봅니다.",
        f"이 사례에서는 {locality} 학생의 과목별 보완 순서를 확인합니다.",
        f"{locality} 상담에서 이 학습 상황을 어떻게 점검할지 살펴봅니다.",
        f"비슷한 상황의 {locality} 학생에게 필요한 상담 항목을 정리합니다.",
        f"이 학습 문제를 중심으로 {locality} 상담의 진행 흐름을 살펴봅니다.",
    ])


def polish_manual_review_families(text: str, locality: str, seed: str) -> str:
    """Remove reader-facing joins found during the final manual copy review."""
    if PROFILE == "middle":
        delayed_review = re.compile(
            r"[^.!?]*(?:학습 계획|학습 점검|학습 과정|세 과목 학습 상담|관리 기준)"
            r"(?:을|를)\s+진행한\s+뒤\s+복습에\s+더\s+잘\s+맞습니다\."
        )
        text = delayed_review.sub(
            stable_pick(seed, "middle-family-review-question", [
                f"{locality} 가정에서는 정답 수보다 학생이 틀린 이유를 설명하는지 확인하는 편이 좋습니다.",
                f"{locality} 학부모는 문제 수보다 자녀가 다시 볼 이유와 날짜를 말하는지 살펴보세요.",
                f"수업 뒤에는 {locality} 학생이 오답 원인과 다음 복습 순서를 설명할 수 있는지 확인하세요.",
                f"{locality} 가정의 질문은 풀이량보다 학생이 다시 확인할 문제를 고르는 데 초점을 두는 편이 좋습니다.",
                f"복습을 확인할 때는 {locality} 학생에게 틀린 이유와 다시 풀 날짜를 차례로 물어보세요.",
                f"{locality} 학부모는 학생이 배운 내용을 스스로 설명하고 다음 과제를 정하는지 살펴볼 수 있습니다.",
                f"가정에서는 {locality} 학생이 오답을 다시 설명하고 복습할 시점을 정했는지 확인하세요.",
                f"{locality} 학생의 복습 상태는 문제 수보다 풀이 근거와 재확인 계획을 묻는 방식으로 살펴보세요.",
            ]),
            text,
        )
        child_frame = re.compile(
            r"[^.!?]*우리\s+아이가\s+[^.!?]{1,140}?학생일\s+때"
            r"[^.!?]*\."
        )
        text = child_frame.sub(
            stable_pick(seed, "middle-parent-consultation-question", [
                f"{locality} 학부모는 자녀가 먼저 보완할 과목과 다음 확인 날짜를 상담에서 구체적으로 물어보세요.",
                f"상담에서는 {locality} 학생의 현재 자료를 보고 어느 과목부터 조정할지 질문하는 편이 좋습니다.",
                f"{locality} 학부모는 세 과목의 가능 여부와 함께 자녀의 우선 보완 순서를 확인하세요.",
                f"자녀의 반복 오답을 준비해 {locality} 상담에서 먼저 손볼 과목과 복습 방법을 물어보세요.",
                f"{locality} 상담 전에는 자녀의 최근 교재와 시험지를 바탕으로 과목별 질문을 정리하세요.",
                f"학부모는 {locality} 학생의 실제 공부 기록에 맞는 과목별 보완 순서를 질문할 수 있습니다.",
                f"{locality}에서는 자녀의 학습 상황을 설명한 뒤 첫 보완 과목과 주간 분량을 확인하세요.",
                f"세 과목을 함께 알아볼 때도 {locality} 학생에게 시급한 단원과 재확인 기준을 따로 물어보세요.",
            ]),
            text,
        )
        text = text.replace(
            "실제 독해에 적용하는 속도가 느린 방학 복습이 필요한",
            "실제 독해에 적용하는 속도가 느려 방학 복습이 필요한",
        )
        text = text.replace(
            "짧은 단위 복습에는 반응이 좋은 학습 공백을 줄이고 싶은",
            "짧은 단위 복습에는 잘 반응하지만 학습 공백을 줄이고 싶은",
        )
        text = text.replace(
            "집중 시간은 짧지만 짧은 단위 복습에는 잘 반응하지만 학습 공백을 줄이고 싶은",
            "집중 시간은 짧아도 짧은 단위 복습에는 잘 반응해 학습 공백을 줄이고 싶은",
        )

    if PROFILE == "elementary":
        student_child = re.compile(
            r"[^.!?]*초등[^.!?]{0,120}?학생이"
            r"[^.!?]{1,180}?아이일수록[^.!?]*\."
        )
        text = student_child.sub(
            stable_pick(seed, "elementary-student-child", [
                f"{locality}에서 비슷한 공부 습관을 보이는 아이에게는 실행 단위를 작게 나누는 편이 좋습니다.",
                f"{locality} 아이가 같은 지점에서 자주 멈춘다면 한 번에 할 과제를 짧고 분명하게 정하세요.",
                f"학습 흐름이 불안정한 {locality} 아이에게는 작은 과제를 직접 끝내는 경험이 필요합니다.",
                f"{locality}에서 복습을 미루는 아이는 국어·영어·수학의 다음 행동을 한 가지씩 정하는 편이 좋습니다.",
                f"비슷한 실수가 이어지는 {locality} 아이에게는 풀이, 채점과 재확인 순서를 짧게 안내하세요.",
                f"{locality} 아이의 공부 습관을 바꾸려면 오늘 완료할 분량과 확인 방법부터 구체적으로 정해야 합니다.",
                f"과제를 시작하기 어려운 {locality} 아이에게는 스스로 끝낼 수 있는 짧은 목표가 도움이 됩니다.",
                f"{locality} 초등 학습에서는 아이가 직접 시작하고 마칠 수 있는 과제 단위를 먼저 정하세요.",
                f"같은 학습 문제가 반복되는 {locality} 아이는 복습할 내용과 날짜를 작게 나누어 기록하는 편이 좋습니다.",
                f"{locality} 아이에게는 긴 계획보다 오늘 실행할 읽기, 어휘와 계산 과제를 분명히 제시해야 합니다.",
                f"학습 공백이 있는 {locality} 아이는 현재 단원에서 바로 실천할 행동부터 하나씩 정해야 합니다.",
                f"{locality}에서는 아이의 학년과 현재 습관을 함께 보고 감당할 수 있는 과제량을 정하세요.",
            ]),
            text,
        )
        nested_grade_traits = re.compile(
            r"[^.!?]*초등학생에게는[^.!?]{0,360}?초등학생의\s+특성에\s+맞춰[^.!?]*\."
        )
        nested_grade_rewrite = stable_pick(seed, "elementary-grade-trait-sentence", [
                f"{locality} 초등 학습에서는 세 과목을 같은 분량으로 밀기보다 현재 교재에서 드러난 보완 순서에 맞춰 설명 비중을 조절해야 합니다.",
                f"{locality} 학생에게는 국어·영어·수학의 현재 어려움을 나누고 필요한 과목부터 과제량을 조정하는 수업이 적합합니다.",
                f"세 과목의 분량을 똑같이 늘리기보다 {locality} 아이가 막히는 과정에 맞춰 교재와 복습 비중을 달리해야 합니다.",
                f"{locality} 초등학생은 현재 교재와 과제에서 확인한 약점에 따라 과목별 설명과 복습량을 조정하는 편이 좋습니다.",
                f"국어·영어·수학을 함께 계획해도 {locality} 아이의 기초 상태에 따라 과목별 과제와 보충 설명을 다르게 정해야 합니다.",
                f"{locality}에서는 같은 분량을 일괄 적용하기보다 아이의 읽기·어휘·계산 상태에 맞춰 학습 비중을 정합니다.",
                f"과목별 현재 수준이 다른 {locality} 아이에게는 교재, 과제와 보충 설명의 순서를 따로 정하는 편이 좋습니다.",
                f"{locality} 초등 과정은 세 과목의 기초 상태를 확인한 뒤 필요한 과목부터 설명과 복습 시간을 조정해야 합니다.",
                f"아이의 현재 단원과 공부 습관을 살펴 {locality} 상담에서 과목별 교재와 과제 비중을 구체적으로 정하세요.",
                f"{locality} 초등 학습에서는 획일적인 분량보다 과목별 어려움에 맞춘 교재와 복습 계획이 필요합니다.",
                f"세 과목을 함께 배우더라도 {locality} 아이가 멈추는 지점에 따라 설명, 과제와 재확인 비중을 달리해야 합니다.",
                f"{locality} 학생의 실제 풀이를 보고 국어·영어·수학의 보완 순서와 과제량을 각각 조정하는 편이 좋습니다.",
            ])
        text = nested_grade_traits.sub(nested_grade_rewrite, text)
        text = re.sub(
            r"같은\s+과목별\s+고민을\s+안고\s+있는\s+같은\s+성장\s+단계의\s+초등\s+아이",
            "과목별 고민이 반복되는 초등학생",
            text,
        )
        text = re.sub(
            r"(초[1-6]에서\s+초[1-6](?:으)?로\s+(?:이어지는|올라가는)\s+시기)\s+중",
            r"\1에",
            text,
        )
        text = re.sub(r"(초[1-6])\s+중\b", r"\1 학생 가운데", text)
        text = text.replace(
            "초등 과정이 되며 중등 준비 질문이 늘어나는",
            "중등 진학을 앞두고 준비 질문이 늘어나는",
        )
        grade_reference = (
            r"(?:이 시기의|비슷한 학년대의|현재 과정의|해당 학년대의|"
            r"같은 성장 단계의|이 학습 단계의|비슷한 과정에 있는|현재 학년 흐름의)"
        )
        text = re.sub(
            rf"({grade_reference})\s+초등\s+중\b",
            r"\1 초등학생 가운데",
            text,
        )
        text = re.sub(
            rf"({grade_reference})\s+초등에는",
            r"\1 초등학생에게는",
            text,
        )
        text = re.sub(
            rf"({grade_reference})\s+초등\s+특성",
            r"\1 초등학생의 특성",
            text,
        )
        text = re.sub(
            rf"{grade_reference}\s+초등에",
            "초등학생 가운데",
            text,
        )
        text = re.sub(r"초등\s+학생", "초등학생", text)
        text = re.sub(r"초등\s+중\b", "초등학생 가운데", text)
        text = text.replace("초등에는", "초등학생에게는")
        text = text.replace("초등에", "초등 과정에")
        text = re.sub(r"초등\s+특성", "초등학생의 특성", text)
        # The conversions immediately above can create a nested
        # `초등학생에게는 ... 초등학생의 특성` sentence. Rewrite after those
        # tokens have reached their final reader-facing form as well.
        text = nested_grade_traits.sub(nested_grade_rewrite, text)
    representative_tautology = re.compile(
        r"[^.!?]*학생(?:을|를)\s+대표\s+상담\s+사례로\s+살펴봅니다\."
    )
    text = representative_tautology.sub(
        representative_situation_sentence(locality, seed),
        text,
    )
    return text


def polish_residual_student_references(text: str, locality: str, seed: str) -> str:
    """Repair late template joins without changing verified facts or headings.

    Repetition reducers deliberately use short references after the first full
    student description. A possessive or an elementary grade-frame can make
    those references ungrammatical, so this final pass works on complete
    sentences and on a few exact, context-safe joins only.
    """
    text = text.replace(
        f"{locality}의 이 상담 사례의 학생에게는",
        f"{locality}에서 비슷한 학습 상황을 보이는 학생에게",
    )
    text = text.replace("이 상담 사례의 학생의", "이 학생의")
    text = text.replace("이 상담 사례의 학생에게는", "이 학생에게")
    text = re.sub(
        r"이\s+학생에게는(?=\s+내신\s+대비는)",
        "이 학생에게",
        text,
    )
    text = re.sub(
        r"이\s+학생에게는\s+(시험\s+3주\s+전|시험\s+직전|중간고사\s+직후)에는",
        r"이 학생은 \1에",
        text,
    )
    text = text.replace("이 학생에게는 오답은", "이 학생은 오답마다")

    if PROFILE == "high":
        repeated_topic_tail = (
            r"(?:(?:시험\s+3주\s+전|시험\s+직전|중간고사\s+직후|"
            r"기말\s+기간)에는|(?:내신\s+대비|오답)(?:은|는))"
        )
        # A student or consultation frame already carries the contrast marker.
        # Keep the topic marker on the concrete time/subject that follows.
        text = re.sub(
            rf"학생에게는(?=\s+{repeated_topic_tail})",
            "학생에게",
            text,
        )
        text = re.sub(
            rf"에서는(?=\s+{repeated_topic_tail})",
            "에서",
            text,
        )

    if PROFILE != "elementary":
        return text

    locality_forms = {locality, locality.split()[-1]}
    locality_grade_frames = (
        "현재 학년 흐름의",
        "현재 과정의",
        "이 학습 단계의",
        "같은 성장 단계의",
        "비슷한 과정에 있는",
        "해당 학년대의",
        "비슷한 학년대의",
    )
    for locality_form in sorted(locality_forms, key=len, reverse=True):
        for grade_frame in locality_grade_frames:
            replacement_frame = (
                "같은 성장 단계에 있는"
                if grade_frame == "같은 성장 단계의"
                else grade_frame
            )
            text = text.replace(
                f"{locality_form}의 {grade_frame}",
                f"{locality_form}에서 {replacement_frame}",
            )

    # Grade sanitization can legitimately fall back to `초등 과정`, but the
    # remaining noun joins should still read like ordinary Korean prose.
    text = text.replace("초등 과정 학생", "초등학생")
    text = text.replace("초등 과정 중", "초등학생 가운데")
    text = text.replace("초등 과정 시기", "초등 시기")
    text = text.replace("초등 과정 아이", "초등 아이")
    text = text.replace("현재 학년 흐름의 초등학생", "현재 학년의 초등학생")

    # Keep the representative frame specific without saying
    # `초등학생 가운데 ... 학생` in the same clause. Grade-frame variants
    # are already established earlier in the paragraph, so the condition can
    # directly modify `초등학생` here.
    representative_among = re.compile(
        r"(?:(?:현재\s+과정의|현재\s+학년\s+흐름의|해당\s+학년대의|"
        r"이\s+학습\s+단계의|같은\s+성장\s+단계에\s+있는|"
        r"비슷한\s+과정에\s+있는|비슷한\s+학년대의)\s+)?"
        r"초등학생\s+가운데\s+([^.!?\n]{1,100}?)\s+학생을\s+대표\s+사례로"
    )
    text = representative_among.sub(
        lambda match: f"{clean(match.group(1))} 초등학생을 대표 사례로",
        text,
    )

    # Remove a second `비슷한` only in the small set of late grade/student
    # joins. These are syntactic joins rather than comparisons that need both
    # words, so the meaning is preserved with `같은` or a direct description.
    text = re.sub(
        r"(비슷한\s+(?:과정에\s+있는|학년대의)\s+"
        r"(?:초등학생|초등\s+아이)(?:이|가))\s+비슷한\s+",
        r"\1 같은 ",
        text,
    )
    repeated_grade_phrases = (
        ("과목별로 비슷한 고민이 있는 비슷한 학년대의", "과목별 고민을 안고 있는"),
        ("비슷한 학습 어려움을 보이는 비슷한 학년대의", "학습 어려움을 보이는"),
        ("이와 비슷한 학습 흐름을 보이는 비슷한 학년대의", "비슷한 학습 흐름을 보이는"),
        ("비슷한 공부 장면이 반복되는 비슷한 학년대의", "같은 공부 장면을 반복해서 겪는"),
    )
    for old, new in repeated_grade_phrases:
        text = text.replace(old, new)
    text = text.replace(
        "단원평가와 수행 과제가 같은 주에 몰릴 수 있는",
        "단원평가와 수행 과제가 한 주에 몰릴 수 있는",
    )
    text = text.replace("같은 성장 단계의 같은 ", "같은 ")
    text = text.replace(
        "같은 성장 단계의 초등학생이 같은 ",
        "초등학생이 같은 ",
    )
    text = text.replace(
        "아이는 많은 문제를 풀어도 같은 실수가 남을 수 있으므로",
        "아이는 문제를 많이 풀어도 실수가 반복될 수 있으므로",
    )

    def rewrite_sentences(
        value: str,
        trigger: str,
        label: str,
        choices: list[str],
    ) -> str:
        pattern = re.compile(rf"[^.!?\n]*{trigger}[^.!?\n]*[.!?]")
        occurrence = 0

        def replace(_: re.Match[str]) -> str:
            nonlocal occurrence
            replacement = stable_pick(seed, f"{label}-{occurrence}", choices)
            occurrence += 1
            return replacement

        return pattern.sub(replace, value)

    textbook_level = [
        f"교재 수준은 앞서 나가는 분량보다 {locality} 아이가 현재 단원을 이해하고 다음 단계로 넘어갈 수 있는지에 맞춰야 합니다.",
        f"{locality} 초등 교재는 난도보다 아이가 배운 개념을 설명하고 다음 문제에 적용할 수 있는지를 기준으로 골라야 합니다.",
        f"높은 단계의 교재보다 {locality} 아이가 현재 단원을 스스로 풀고 부족한 부분을 보완할 수 있는 교재가 적절합니다.",
        f"{locality}에서는 교재의 권수보다 아이가 이해한 내용과 다시 확인할 문제를 구분할 수 있는지를 먼저 살펴야 합니다.",
        f"초등 교재의 난도는 {locality} 아이가 개념을 이해하고 한 단계씩 적용 범위를 넓힐 수 있도록 조정해야 합니다.",
        f"{locality} 아이에게는 어려운 교재를 서두르기보다 현재 단원의 빈칸을 채우며 다음 단계로 이어지는 구성이 필요합니다.",
        f"교재를 정할 때는 {locality} 학생이 혼자 설명할 수 있는 범위와 도움이 필요한 문제를 함께 확인하는 편이 좋습니다.",
        f"{locality} 초등 학습에서는 진도를 앞당기는 교재보다 현재 개념을 이해하고 복습할 수 있는 수준이 중요합니다.",
        f"아이의 교재는 {locality} 상담에서 확인한 기초 수준과 학습 속도에 맞춰 단계적으로 조정해야 합니다.",
        f"{locality} 학생의 교재 수준은 문제 수가 아니라 개념 이해, 풀이 설명과 재확인 가능성을 기준으로 판단하세요.",
        f"초등 교재를 비교할 때는 {locality} 아이가 현재 내용을 소화하고 다음 단원으로 자연스럽게 이어갈 수 있는지 살펴야 합니다.",
        f"{locality}에서는 아이의 실제 풀이를 보고 무리 없이 이해할 수 있는 교재와 보완 과제를 정하는 편이 좋습니다.",
    ]
    after_class_record = [
        f"{locality} 초등학생에게는 국어·영어·수학 수업 뒤에 복습한 내용과 다시 질문할 부분을 짧게 기록하는 관리가 필요합니다.",
        f"{locality} 아이는 세 과목 수업을 마친 뒤 이해한 내용, 다시 풀 문제와 다음 질문을 구분해 남기는 편이 좋습니다.",
        f"수업 뒤에는 {locality} 학생이 국어·영어·수학에서 복습할 내용과 질문할 지점을 직접 정리하도록 도와야 합니다.",
        f"{locality} 초등 학습에서는 수업이 끝난 뒤 과목별 복습 내용과 재확인할 문제를 기록하는 과정이 중요합니다.",
        f"국어·영어·수학 수업 후 {locality} 아이가 배운 내용과 막힌 부분을 나누어 적을 수 있도록 안내하세요.",
        f"{locality} 학생에게는 수업 뒤 세 과목의 복습 순서와 다음 시간에 물어볼 내용을 분명히 정해 주는 관리가 필요합니다.",
        f"세 과목을 함께 배우는 {locality} 아이는 수업 후 복습할 내용과 다시 질문할 부분을 과목별로 남기는 편이 좋습니다.",
        f"{locality} 초등학생의 수업 기록에는 국어·영어·수학에서 이해한 부분과 다시 확인할 부분이 구분되어야 합니다.",
        f"수업을 마친 {locality} 아이가 과목별 복습 범위와 질문을 스스로 말하고 기록하는지 확인하세요.",
        f"{locality} 초등 과정은 수업 뒤 배운 내용, 오답과 다음 질문을 세 과목별로 정리하는 흐름이 필요합니다.",
        f"국어·영어·수학을 배운 뒤 {locality} 학생이 무엇을 복습하고 언제 다시 물을지 정하도록 도와주세요.",
        f"{locality} 아이의 수업 후 관리에서는 과목별 복습 내용과 다음 질문을 짧고 구체적으로 남겨야 합니다.",
    ]
    grade_trait_copy = [
        f"{locality} 초등 학습에서는 세 과목을 같은 분량으로 밀기보다 아이의 현재 어려움에 맞춰 교재, 과제와 보충 설명의 비중을 조절해야 합니다.",
        f"{locality} 학생에게는 국어·영어·수학의 보완 순서를 나누고 필요한 과목부터 교재와 과제 비중을 조정하는 수업이 적합합니다.",
        f"세 과목의 분량을 똑같이 늘리기보다 {locality} 아이가 막히는 지점에 맞춰 교재와 복습 비중을 달리해야 합니다.",
        f"{locality} 초등학생은 현재 교재에서 확인한 약점에 따라 과목별 설명과 보충 과제의 비중을 조정하는 편이 좋습니다.",
        f"국어·영어·수학을 함께 계획해도 {locality} 아이의 기초 상태에 따라 교재, 과제와 보충 설명을 다르게 정해야 합니다.",
        f"{locality}에서는 획일적인 분량보다 아이의 읽기·어휘·계산 상태에 맞춘 교재와 복습 계획이 필요합니다.",
    ]
    text = rewrite_sentences(
        text,
        r"초등\s+과정\s+특성에\s+맞춰",
        "elementary-course-trait",
        grade_trait_copy,
    )
    text = rewrite_sentences(
        text,
        r"초등(?:\s+과정)?\s+상태(?=[^.!?\n]*(?:단계|교재|핵심))",
        "elementary-course-state-textbook",
        textbook_level,
    )
    text = rewrite_sentences(
        text,
        r"초등(?:\s+과정)?\s+상태(?=[^.!?\n]*수업\s+뒤)",
        "elementary-course-state-after-class",
        after_class_record,
    )

    student_subject = (
        r"초등학생이\s+[^.!?\n]{0,100}?"
        r"(?:문제|고민|과제)(?:이|가)\s+(?:나타나는|있는|남은)"
    )
    textbook_check = [
        f"{locality} 초등 교재를 점검할 때는 국어의 근거 표시, 영어의 단어와 문장 연결, 수학의 풀이 과정을 함께 살펴야 합니다.",
        f"교재 점검에서는 {locality} 아이가 국어 지문의 근거, 영어 문장 적용과 수학 풀이 순서를 직접 설명하는지 확인하세요.",
        f"{locality} 학생의 교재는 과목별 오답뿐 아니라 국어·영어·수학의 풀이 과정이 남아 있는지를 함께 보아야 합니다.",
        f"{locality} 초등 학습에서는 교재의 분량보다 국어의 읽기 근거, 영어의 문장 이해와 수학의 풀이 흔적을 확인하는 일이 중요합니다.",
        f"국어·영어·수학 교재를 볼 때는 {locality} 아이가 개념과 풀이 순서를 말로 설명할 수 있는지 살펴보세요.",
        f"{locality} 교재 상담에서는 아이가 표시한 근거, 어휘 적용과 계산 과정을 과목별로 나누어 확인해야 합니다.",
        f"교재를 고르기 전 {locality} 학생의 국어 읽기, 영어 문장 적용과 수학 풀이 기록에서 보완할 지점을 찾아야 합니다.",
        f"{locality} 아이의 세 과목 교재는 정답 수보다 근거를 찾고 풀이 과정을 남기는 습관을 확인하는 자료여야 합니다.",
    ]
    feedback_copy = [
        f"{locality} 초등 피드백은 점수보다 풀이 과정, 어휘 확인과 개념 연결을 중심으로 전달해야 합니다.",
        f"피드백에서는 {locality} 아이가 멈춘 지점과 다음에 복습할 행동을 국어·영어·수학별로 구분해 알려 주세요.",
        f"{locality} 학생에게는 결과만 말하기보다 읽기 근거, 문장 적용과 풀이 과정을 구체적으로 설명하는 피드백이 필요합니다.",
        f"국어·영어·수학 피드백은 {locality} 아이의 오답 원인과 다음 복습 순서를 함께 안내하는 방식이 적절합니다.",
        f"{locality} 초등 상담에서는 점수 변화보다 아이가 이해한 내용과 다시 확인할 부분을 과목별로 전달해야 합니다.",
        f"피드백을 받을 때는 {locality} 학생의 풀이 과정, 어휘 이해와 개념 적용이 어떻게 달라졌는지 확인하세요.",
        f"{locality} 아이의 세 과목 피드백에는 잘한 부분, 막힌 이유와 다음 주 복습 행동이 함께 담겨야 합니다.",
        f"결과를 설명할 때는 {locality} 학생의 국어·영어·수학 학습 과정과 다음 확인 시점을 구체적으로 제시해야 합니다.",
    ]
    diagnosis_copy = [
        f"{locality} 초등 진단에서는 선행 교재를 늘리기 전 지난 단원의 개념, 어휘와 풀이 흔적을 차례로 확인해야 합니다.",
        f"선행을 서두르기보다 {locality} 아이의 최근 교재에서 개념 공백과 반복 오답을 먼저 찾는 편이 안전합니다.",
        f"{locality} 학생의 첫 점검은 지난 단원의 이해도, 어휘 적용과 풀이 과정을 확인하는 데서 시작해야 합니다.",
        f"초등 진단에서는 {locality} 아이가 현재 단원을 얼마나 설명하고 적용할 수 있는지부터 살펴보세요.",
        f"{locality} 상담에서는 다음 교재를 정하기 전 아이의 개념 이해, 어휘와 계산 과정을 먼저 확인해야 합니다.",
        f"지난 단원의 풀이 흔적을 살펴 {locality} 학생이 선행 전에 보완할 개념과 학습 습관을 구분하세요.",
        f"{locality} 초등학생에게는 빠른 진도보다 현재 단원의 빈칸을 확인하고 복습 순서를 정하는 과정이 필요합니다.",
        f"새 교재를 시작하기 전 {locality} 아이의 최근 오답과 설명 과정을 보고 먼저 채울 기초를 찾아야 합니다.",
        f"{locality} 초등 학습의 출발점은 선행 분량이 아니라 지난 단원을 이해하고 다시 풀 수 있는지 확인하는 일입니다.",
        f"진단 결과는 {locality} 학생이 현재 교재에서 막힌 이유와 다음 복습 행동을 정하는 데 활용해야 합니다.",
    ]

    text = rewrite_sentences(
        text,
        rf"(?:교재\s+점검[^.!?\n]{{0,180}}?{student_subject}|"
        rf"{student_subject}[^.!?\n]{{0,180}}?교재\s+점검)",
        "elementary-student-subject-textbook",
        textbook_check,
    )
    text = rewrite_sentences(
        text,
        rf"{student_subject}(?=[^.!?\n]*피드백)",
        "elementary-student-subject-feedback",
        feedback_copy,
    )
    text = rewrite_sentences(
        text,
        student_subject,
        "elementary-student-subject-diagnosis",
        diagnosis_copy,
    )

    repeated_similar = re.compile(
        r"비슷한(?P<middle>[^.!?\n]{0,100}?)비슷한\s+"
        r"(?P<noun>과정|학습|보완|복습)"
    )
    similar_nouns = {
        "과정": "현재 과정",
        "학습": "학습",
        "보완": "보완",
        "복습": "복습",
    }
    while repeated_similar.search(text):
        text = repeated_similar.sub(
            lambda match: (
                "비슷한"
                + re.sub(r"이와\s*$", "이런 ", match.group("middle"))
                + similar_nouns[match.group("noun")]
            ),
            text,
        )
    repeated_same = re.compile(
        r"같은(?P<middle>[^.!?\n]{0,100}?)같은\s+"
        r"(?P<noun>성장|공부|학습)"
    )
    while repeated_same.search(text):
        text = repeated_same.sub(
            lambda match: "같은" + match.group("middle") + match.group("noun"),
            text,
        )
    text = re.sub(
        r"(?<![가-힣])이와\s+(?=(?:학습|과정|복습|보완))",
        "이런 ",
        text,
    )
    text = text.replace(
        "이런 학습 부담을 느끼는 초등 과정 특성에 맞춰",
        "이런 학습 부담을 고려해",
    )
    text = text.replace(
        "비슷한 복습 문제가 나타나는",
        "복습 문제가 반복되는",
    )
    text = text.replace(
        "복습 문제가 반복되는 현재 과정에 있는 초등학생",
        "현재 과정에서 복습 문제를 반복해서 겪는 초등학생",
    )
    return text


def diversify_common_copy(text: str, locality: str, seed: str) -> str:
    """Vary common guidance without changing any center or school fact."""
    common_banks: list[tuple[str, str, list[str]]] = [
        (
            "school-list-purpose",
            "이 목록은 수업 가능 여부를 단정하는 표현이 아니라 상담에서 학생의 실제 학교 자료를 확인하기 위한 참고 정보입니다.",
            [
                "학교 목록은 수업 가능 여부를 보장하지 않으며, 상담에서 학생이 가져온 실제 학교 자료를 대조하기 위한 참고 정보입니다.",
                "표시된 학교명만으로 수업 가능 여부를 판단하지 않고, 상담 때 학생의 시험·교과 자료를 함께 확인합니다.",
                "이 학교 정보는 상담 준비를 위한 참고 목록입니다. 실제 수업 여부는 학생 자료와 센터의 현재 개설 범위를 함께 확인해야 합니다.",
                "학교명은 현재 학습 자료를 확인하는 출발점으로만 활용하며, 센터별 수업 가능 여부를 뜻하지 않습니다.",
                "제공 목록은 학교 자료 확인을 돕기 위한 정보이고, 실제 개설 과목과 수업 가능 여부는 상담에서 별도로 확인합니다.",
                "목록에 포함된 학교는 상담 참고 대상이며, 학생의 실제 시험 범위와 센터 운영 범위를 확인한 뒤 판단해야 합니다.",
            ],
        ),
        (
            "school-material-priority",
            "학교명 반복보다 현재 자료를 어떻게 수업 계획에 반영하는지가 더 중요합니다.",
            [
                "학교명을 여러 번 언급하기보다 학생이 가져온 자료를 다음 수업 계획에 어떻게 반영하는지가 중요합니다.",
                "핵심은 학교명 자체가 아니라 현재 시험 범위와 과제 자료가 과목별 계획으로 이어지는 방식입니다.",
                "학교 정보보다 실제 자료에서 확인한 단원과 오답을 다음 학습 순서로 바꾸는 과정이 우선입니다.",
                "학교 이름을 강조하기보다 시험·교과 자료를 읽고 보완 순서를 정하는지를 확인해야 합니다.",
                "현재 학교 자료가 수업 진도와 복습 일정에 구체적으로 반영되는지를 살펴보는 편이 정확합니다.",
                "학교별 자료는 이름 나열보다 시험 범위, 과제와 수행평가 일정을 계획에 연결할 때 의미가 있습니다.",
            ],
        ),
        (
            "read-starting-point",
            "처음 필요한 단계는 성적표보다 최근 오답과 공부 시간을 함께 읽는 것입니다.",
            [
                "첫 단계에서는 성적표만 보기보다 최근 오답과 실제 공부 시간을 함께 살펴야 합니다.",
                "출발점을 정할 때는 점수보다 최근 틀린 문제와 주간 공부 시간을 나란히 확인하는 편이 좋습니다.",
                "처음에는 성적 한 줄보다 오답이 생긴 과정과 공부 시간의 사용 방식을 함께 확인합니다.",
                "현재 상태를 파악하려면 성적표와 함께 최근 오답, 과제 시작과 완료 시간을 살펴야 합니다.",
                "진단의 시작은 점수 비교가 아니라 최근 오답과 실제 학습 시간을 연결해 보는 과정입니다.",
                "우선 최근 풀이 기록과 공부 시간을 대조해 어느 과목에서 흐름이 끊기는지 확인해야 합니다.",
                "현재 위치는 성적만으로 정하지 않고 최근 풀이와 과제에 사용한 시간을 함께 보고 판단합니다.",
                "먼저 오답이 생긴 과정과 한 주의 공부 시간을 살펴 과목별 병목을 찾아야 합니다.",
            ],
        ),
        (
            "family-observation",
            f"{locality} 학부모가 집에서 할 수 있는 일은 문제를 대신 풀어 주는 것이 아니라 자녀가 국어·영어·수학 중 어느 과목에서 시간을 잃는지 관찰하는 것입니다.",
            [
                f"{locality} 가정에서는 문제를 대신 해결하기보다 자녀가 국어·영어·수학 중 어느 과목에서 오래 멈추는지 기록해 두는 것이 좋습니다.",
                f"{locality} 학부모가 확인할 부분은 정답을 알려 주는 일이 아니라 세 과목 중 시작과 완료가 늦어지는 지점을 살피는 것입니다.",
                f"가정에서는 {locality} 학생의 풀이를 대신하기보다 과목별 소요 시간과 질문을 미룬 장면을 짧게 남겨 주세요.",
                f"{locality} 학생의 집 공부를 볼 때는 문제를 대신 풀기보다 국어·영어·수학의 멈춘 지점과 이유를 확인하는 편이 좋습니다.",
                f"학부모는 {locality} 학생이 세 과목 중 어디에서 시간을 많이 쓰는지 관찰하고 상담 때 그 기록을 공유할 수 있습니다.",
                f"{locality} 가정의 역할은 풀이를 대신하는 것이 아니라 과목별 시작 시각, 완료 여부와 반복 질문을 확인하는 데 있습니다.",
            ],
        ),
        (
            "elementary-direct-intro",
            f"{locality} {LEVEL_NAME} 국영수학원을 검색한 학부모가 바로 알고 싶은 답은 분명합니다.",
            [
                f"{locality}에서 {LEVEL_NAME} 국어·영어·수학 학습을 알아볼 때 먼저 확인할 기준이 있습니다.",
                f"{locality} {LEVEL_NAME}의 세 과목 학습을 비교한다면 현재 기초와 공부 습관부터 살펴야 합니다.",
                f"{locality} 학부모가 {LEVEL_NAME} 국영수 수업을 찾을 때 첫 질문은 과목 수보다 현재 공백이 무엇인지입니다.",
                f"{locality} {LEVEL_NAME}에게 필요한 국영수 관리는 진도보다 과목별로 막히는 장면을 확인하는 데서 시작합니다.",
                f"{locality}에서 {LEVEL_NAME} 국영수학원을 선택하기 전에는 학생의 현재 교재와 풀이 과정을 먼저 확인해야 합니다.",
                f"{locality} {LEVEL_NAME}의 국어·영어·수학 계획은 세 과목을 모두 늘리기 전에 우선순위를 나누는 과정이 필요합니다.",
            ],
        ),
        (
            "high-weekly-balance",
            f"{locality} {LEVEL_NAME}에게는 한 과목의 과제량이 다른 과목 복습을 밀어내지 않도록 주간 균형을 조정하는 과정이 필요합니다.",
            [
                f"{locality} {LEVEL_NAME}은 한 과목의 과제가 다른 과목 복습 시간을 잠식하지 않도록 주간 분량을 나누어야 합니다.",
                f"{locality} {LEVEL_NAME}의 계획에서는 집중 과목을 정하되 나머지 과목의 최소 복습 시간을 함께 남겨야 합니다.",
                f"{locality}에서 세 과목 일정이 겹치는 {LEVEL_NAME}에게는 과제량과 복습 시간을 현실적으로 조정하는 과정이 필요합니다.",
                f"{locality} {LEVEL_NAME}은 시험 일정에 따라 집중 과목과 유지 과목을 구분해 주간 균형을 맞추는 편이 좋습니다.",
                f"{locality}에서 한 과목에 시간이 몰리는 {LEVEL_NAME}이라면 다른 과목의 복습이 끊기지 않도록 최소 실행량을 정해야 합니다.",
                f"{locality} {LEVEL_NAME}의 세 과목 계획은 과제 충돌을 줄이고 각 과목의 재확인 시간을 확보하는 방식이어야 합니다.",
            ],
        ),
        (
            "action-starting-point",
            f"{locality} {LEVEL_NAME}의 좋은 출발점은 학생에게 많은 약속을 주는 것이 아니라 오늘부터 바꿀 한두 가지 학습 행동을 정하는 데 있습니다.",
            [
                f"{locality} {LEVEL_NAME}의 학습 변화는 거창한 약속보다 이번 주에 실행할 한두 가지 행동을 정하는 데서 시작합니다.",
                f"{locality} {LEVEL_NAME}에게 필요한 첫 단계는 계획을 크게 잡기보다 바로 확인할 학습 행동을 구체화하는 것입니다.",
                f"처음부터 많은 목표를 제시하기보다 {locality} {LEVEL_NAME}이 오늘 바꿀 수 있는 공부 행동을 정해야 합니다.",
                f"{locality} {LEVEL_NAME}의 출발점은 약속의 수가 아니라 다음 수업까지 지킬 한두 가지 기준을 세우는 데 있습니다.",
                f"실행 가능한 변화는 {locality} {LEVEL_NAME}이 이번 주에 반복할 작은 행동을 정할 때 시작됩니다.",
                f"{locality} {LEVEL_NAME}의 계획은 많은 목표보다 바로 실천하고 확인할 행동 한두 가지를 먼저 담아야 합니다.",
            ],
        ),
    ]
    for label, source, choices in common_banks:
        if source in text:
            text = text.replace(source, stable_pick(seed, label, choices))
    repeated_selection = re.compile(
        rf"{re.escape(locality)}에서 (?:{re.escape(locality)} )?{re.escape(CATEGORY_NAME)}을 고를 때 상담표에 현재 단원, 틀린 유형, 과제 지속 시간이 같이 적히면 수업 방향을 더 현실적으로 세울 수 있습니다\."
    )
    if repeated_selection.search(text):
        text = repeated_selection.sub(stable_pick(seed, "selection-record", [
            f"{locality} 상담표에는 현재 단원과 반복 오답, 과제 소요 시간을 함께 적어 과목별 수업 방향을 구체화합니다.",
            f"{locality}에서 수업을 비교할 때는 최근 단원, 틀린 유형과 실제 과제 시간을 한 기록에서 확인하는 편이 좋습니다.",
            f"현재 단원과 오답 유형, 과제 지속 시간을 나란히 적으면 {locality} 학생의 다음 학습 순서를 현실적으로 정할 수 있습니다.",
            f"{locality} 학습 상담에서는 점수 외에도 단원별 오답과 과제에 쓴 시간을 기록해 수업 계획과 대조합니다.",
            f"상담 기록에 현재 범위, 반복해 틀린 문제와 완료 시간을 남겨 {locality} 학생에게 필요한 보완 순서를 찾습니다.",
            f"{locality} 학생의 최근 교재에서 단원, 오답 원인과 풀이 시간을 확인한 뒤 과목별 계획을 조정해야 합니다.",
            f"수업 방향은 이름만으로 정하지 않고 {locality} 학생의 현재 단원, 오답과 주간 실행 시간을 함께 보고 판단합니다.",
            f"{locality} 상담 전에는 최근 범위와 틀린 유형, 과제 완료 시간을 준비해 세 과목의 우선순위를 나눠 보세요.",
        ]), text)
    action_pattern = re.compile(
        r"[^.!?]*좋은 출발점은 학생에게 많은 약속을 주는 것이 아니라 오늘부터 바꿀 한두 가지 학습 행동을 정하는 데 있습니다\."
    )
    action_match = action_pattern.search(text)
    if action_match:
        separator = " " if action_match.start() and text[action_match.start() - 1] in ".!?" else ""
        text = text[:action_match.start()] + separator + stable_pick(seed, "action-start-generic", [
            f"{locality} 학생은 이번 주에 실행할 과목별 행동 한두 가지부터 구체적으로 정하는 편이 좋습니다.",
            f"많은 목표를 약속하기보다 {locality} 학생이 다음 수업까지 반복할 작은 행동을 먼저 정해야 합니다.",
            f"{locality} 학습 계획의 출발점은 오늘 실천하고 다음 점검에서 확인할 기준을 고르는 일입니다.",
            f"처음에는 {locality} 학생이 바로 바꿀 수 있는 공부 행동과 확인 날짜를 함께 정하는 것이 중요합니다.",
            f"실행 가능한 계획은 {locality} 학생의 이번 주 과제와 재풀이 기준을 작게 나눌 때 시작됩니다.",
            f"{locality} 학생에게는 거창한 목표보다 과목별로 지킬 최소 행동과 점검 시점이 먼저 필요합니다.",
            f"학습 변화를 시작하려면 {locality} 학생이 직접 수행할 한두 가지 행동을 주간표에 배치해야 합니다.",
            f"{locality} 상담에서는 약속의 수보다 학생이 이번 주에 끝내고 확인할 학습 행동을 정리합니다.",
        ]) + text[action_match.end():]
    return text


def diversify_repeated_sentences(text: str, seed: str) -> str:
    """Use transitions sparingly while preserving each manuscript claim."""
    fixed_frames = [
        (
            "학부모님은 국어·영어·수학을 모두 맡긴다는 표현보다 과목별 피드백이 서로 충돌하지 않는지 확인해야 합니다.",
            [
                "세 과목을 함께 맡길 때는 국어·영어·수학의 피드백과 과제가 서로 충돌하지 않는지 확인해야 합니다.",
                "국어·영어·수학을 한 시간표에 넣더라도 과목별 보완 순서와 피드백은 따로 살펴야 합니다.",
                "과목 수보다 중요한 것은 세 과목의 과제와 복습 시간이 한 주 안에서 무리 없이 이어지는지입니다.",
                "세 과목 학습을 계획할 때는 과목별 피드백이 다른 과목의 복습을 밀어내지 않는지 살펴보세요.",
                "국어·영어·수학을 함께 관리해도 오답 원인과 다음 확인 날짜는 과목마다 구분해야 합니다.",
                "상담에서는 세 과목을 모두 다룬다는 설명보다 과목별 피드백이 실제 계획에 반영되는지 물어보세요.",
                "한 과목의 과제량이 다른 과목의 오답 복습을 방해하지 않도록 과목별 최소 학습량을 정해야 합니다.",
                "세 과목의 시간표를 합치기 전 과목별 진단 기록과 보완 과제가 충돌하지 않는지 확인하는 편이 좋습니다.",
            ],
        ),
        (
            "그리고 선지 판단 이유를 말로 설명해 보게 하면 감으로 맞힌 문제와 알고 맞힌 문제가 나뉩니다.",
            [
                "선지 판단 이유를 설명하게 하면 감으로 고른 답과 근거를 알고 고른 답을 구분할 수 있습니다.",
                "국어 오답은 선택한 이유를 말로 확인해야 우연히 맞힌 문항까지 찾아낼 수 있습니다.",
                "학생이 선지의 근거를 직접 설명하면 읽기 과정에서 놓친 부분이 더 분명해집니다.",
                "정답만 확인하지 말고 선지를 고른 근거를 말하게 해 실제 이해 여부를 살펴보세요.",
                "선택지마다 맞고 틀린 이유를 설명하는 과정에서 감으로 푼 문제를 따로 찾을 수 있습니다.",
                "국어 점검에서는 학생이 답의 근거가 된 문장을 직접 가리키고 설명하도록 해야 합니다.",
            ],
        ),
        (
            "그리고 풀이가 길어질수록 어느 단계에서 흔들리는지 기록해야 보충 수업의 방향이 분명해집니다.",
            [
                "풀이가 길어질 때 멈춘 단계를 기록하면 수학 보완 순서를 더 구체적으로 정할 수 있습니다.",
                "수학 풀이에서는 조건 해석, 식 세우기와 계산 중 어느 단계에서 흔들렸는지 나눠 기록하세요.",
                "긴 풀이의 오류 지점을 단계별로 표시해야 다음 보충 학습의 범위가 분명해집니다.",
                "학생이 풀이를 이어가지 못한 지점을 남기면 다시 확인할 개념과 유형을 찾기 쉽습니다.",
                "수학 오답은 마지막 답보다 풀이가 끊긴 단계를 확인해 다음 과제에 반영해야 합니다.",
                "조건을 읽은 뒤 계산을 마칠 때까지 어느 과정에서 실수가 생겼는지 기록하는 편이 좋습니다.",
            ],
        ),
        (
            "그리고 문법을 독해 속에 넣어 점검하면 시험장에서 헷갈리는 유형을 줄이는 데 도움이 됩니다.",
            [
                "영어 문법은 독해 문장에 적용해 봐야 실제로 헷갈리는 유형을 구분할 수 있습니다.",
                "문법 개념을 본문 해석에 연결하면 시험에서 반복되는 오류를 더 정확히 찾을 수 있습니다.",
                "영어 점검에서는 문법 용어를 아는지보다 문장 안에서 구조를 적용하는지 살펴야 합니다.",
                "배운 문법을 실제 지문에 표시하고 해석해 보면 다시 확인할 유형이 분명해집니다.",
                "문법 문제와 독해를 따로 보지 말고 문장 구조를 읽는 과정에서 적용 여부를 확인하세요.",
                "학생이 문법 근거를 들어 문장을 해석하도록 하면 시험에서 흔들리는 부분을 찾기 쉽습니다.",
            ],
        ),
    ]
    for frame_index, (source, choices) in enumerate(fixed_frames):
        if source in text:
            text = text.replace(source, stable_pick(seed, f"fixed-frame-{frame_index}", choices))
    semantic_frames = [
        (
            "영어 수업은 단어 암기량만 늘리기보다 문장 구조와 해석 순서를 함께 확인합니다.",
            [
                "영어 학습에서는 단어 수와 함께 문장 구조를 찾고 해석하는 순서를 확인합니다.",
                "영어 수업은 암기한 단어를 실제 문장에 적용하고 해석 과정을 설명하는지 살펴봅니다.",
                "영어 점검에서는 단어량보다 문장의 뼈대를 찾고 의미를 연결하는 과정을 봅니다.",
                "영어는 어휘 암기와 별도로 문장 구조를 나누고 해석 근거를 말하는 연습이 필요합니다.",
                "단어를 많이 외웠는지만 보지 않고 문장 안에서 어휘와 구문을 적용하는지 확인합니다.",
                "영어 계획은 어휘 복습과 문장 분석, 해석 재확인 순서를 함께 담아야 합니다.",
            ],
        ),
        (
            "국어 수업은 지문을 오래 읽는 훈련보다 먼저 문제의 근거가 어느 문장에 있는지 표시하게 합니다.",
            [
                "국어 학습에서는 읽은 시간보다 답의 근거가 된 문장을 정확히 표시하는지 확인합니다.",
                "국어 수업은 지문을 반복해서 읽기 전에 문제와 연결되는 근거 문장을 찾게 합니다.",
                "국어 점검에서는 학생이 선택지의 근거를 본문에서 직접 가리킬 수 있는지 살펴봅니다.",
                "지문 읽기는 시간만 늘리기보다 핵심 문장과 답의 근거를 구분하는 연습이 필요합니다.",
                "국어 오답은 본문에서 놓친 문장과 선지를 고른 이유를 함께 기록해야 합니다.",
                "국어 계획에는 지문의 중심 내용, 근거 표시와 선택지 판단 과정을 따로 담습니다.",
            ],
        ),
        (
            "수학 수업은 공식 암기 다음에는 조건 해석, 식 세우기, 계산 검산을 분리해 봅니다.",
            [
                "수학 학습은 공식을 외운 뒤 조건 해석, 식 세우기와 검산 단계를 나누어 확인합니다.",
                "수학 수업에서는 답보다 조건을 읽고 식을 세운 뒤 계산을 확인하는 과정을 살펴봅니다.",
                "수학 오답은 개념 선택, 식 구성과 계산 실수를 각각 구분해 기록해야 합니다.",
                "공식 암기만 확인하지 않고 문제 조건을 식으로 옮기고 검산하는 순서를 점검합니다.",
                "수학 계획에는 조건 해석에서 계산 확인까지 학생이 멈춘 단계를 따로 표시합니다.",
                "수학은 풀이 과정을 개념, 식 세우기와 계산 단계로 나누어 다시 확인하는 편이 좋습니다.",
            ],
        ),
        (
            "영어가 흔들리는 학생은 문법을 독해 속에 넣어 점검하면 시험장에서 헷갈리는 유형을 줄이는 데 도움이 됩니다.",
            [
                "영어가 불안정하다면 배운 문법을 지문 해석에 적용해 반복 오류를 찾아야 합니다.",
                "영어 문법은 개념 확인에 그치지 않고 실제 문장에서 구조를 찾게 해야 합니다.",
                "독해에서 자주 멈추는 학생은 문법 근거를 들어 문장을 나누어 해석해 보세요.",
                "영어 오답에서는 어휘 문제와 구문 적용 문제를 나누어 다음 복습 범위를 정합니다.",
                "문법을 알고도 해석이 흔들린다면 본문에서 적용한 과정과 틀린 이유를 기록해야 합니다.",
                "영어 점검은 문법 용어보다 실제 문장을 분석하고 의미를 연결하는 과정에 초점을 둡니다.",
            ],
        ),
        (
            "국어가 흔들리는 학생은 선지 판단 이유를 말로 설명해 보게 하면 감으로 맞힌 문제와 알고 맞힌 문제가 나뉩니다.",
            [
                "국어가 불안정하다면 학생이 선택지의 근거를 직접 설명하도록 해야 합니다.",
                "국어 오답은 답을 고른 이유를 말하게 해 우연히 맞힌 문항까지 구분합니다.",
                "국어 점검에서는 선지마다 맞고 틀린 이유를 본문 근거와 연결해 봅니다.",
                "읽기에서 흔들리는 학생은 답의 근거가 된 문장을 표시하고 설명해야 합니다.",
                "선택지 판단 과정을 말로 확인하면 감으로 푼 문제와 이해한 문제를 나눌 수 있습니다.",
                "국어 학습 기록에는 정답뿐 아니라 근거 문장과 판단 이유를 함께 남겨야 합니다.",
            ],
        ),
        (
            "수학이 흔들리는 학생은 풀이가 길어질수록 어느 단계에서 흔들리는지 기록해야 보충 수업의 방향이 분명해집니다.",
            [
                "수학이 불안정하다면 긴 풀이에서 멈춘 단계를 표시해 보완 순서를 정해야 합니다.",
                "수학 오답은 조건 해석, 식 구성과 계산 중 오류가 난 지점을 따로 기록합니다.",
                "풀이가 길어질 때 흔들리는 학생은 문제를 단계로 나누어 다시 설명해 보세요.",
                "수학 점검에서는 마지막 답보다 풀이가 끊긴 과정과 그 이유를 먼저 살펴봅니다.",
                "긴 문제에서 실수가 이어진다면 개념 선택부터 검산까지 단계별로 확인해야 합니다.",
                "수학 보완 계획은 학생이 풀이를 이어가지 못한 지점에서 시작하는 편이 정확합니다.",
            ],
        ),
        (
            "국영수 진단은 점수만 보는 방식으로는 부족하고, 국어 지문 읽기, 영어 문장 해석, 수학 조건 해석을 따로 분리해야 합니다.",
            [
                "세 과목 진단은 점수와 함께 국어 읽기, 영어 해석과 수학 조건 파악 과정을 나누어 봅니다.",
                "국영수 점검에서는 결과보다 과목별로 풀이가 멈춘 장면과 이유를 구분해야 합니다.",
                "세 과목은 같은 기준으로 묶지 않고 읽기, 문장 해석과 수학 풀이 과정을 따로 확인합니다.",
                "국어·영어·수학 진단은 점수표와 더불어 과목별 사고 과정과 오답 원인을 살펴야 합니다.",
                "세 과목의 현재 상태를 보려면 국어 근거 찾기, 영어 구문 적용과 수학 식 세우기를 구분합니다.",
                "국영수 계획은 한 점수로 정하지 않고 과목마다 막힌 단계와 재확인 기준을 나누어야 합니다.",
            ],
        ),
        (
            "학습 환경을 고려하면 학교 일정, 과제 마감, 귀가 후 복습 시간을 함께 놓고 봐야 무리 없는 관리가 가능합니다.",
            [
                "학습 계획은 학교 일정과 과제 마감, 귀가 뒤 확보할 수 있는 복습 시간을 함께 반영해야 합니다.",
                "학교 일정, 과제 기한과 실제 귀가 시간을 놓고 한 주에 가능한 학습량을 계산합니다.",
                "무리 없는 시간표인지 판단하려면 등하원 동선과 과제 마감, 가정 복습 시간을 같이 봐야 합니다.",
                "주간 계획에는 학교 행사와 과제 기한뿐 아니라 집에서 다시 공부할 시간을 남겨야 합니다.",
                "학습량을 정할 때는 학교 일정, 이동 시간과 귀가 후 복습 가능성을 함께 살펴보세요.",
                "과제와 복습이 이어지려면 실제 생활 시간표 안에 세 과목의 실행 시간을 배치해야 합니다.",
            ],
        ),
        (
            "학생에게 맞는 처방은 선행 속도를 높이는 계획이 아니라 현재 단원에서 반복되는 실수를 줄이는 계획일 수 있습니다.",
            [
                "학생의 계획은 진도를 앞당기기보다 현재 단원에서 반복되는 오류를 줄이는 데서 시작할 수 있습니다.",
                "선행 분량을 늘리기 전에 학생이 현재 범위에서 다시 틀리는 이유를 확인해야 합니다.",
                "학생에게 필요한 학습량은 빠른 진도보다 현재 단원의 오답을 정확히 줄이는 수준일 수 있습니다.",
                "현재 단원에서 같은 실수가 이어진다면 선행보다 개념과 풀이 과정을 다시 확인하는 편이 좋습니다.",
                "학습 처방은 속도보다 학생이 반복해서 놓치는 내용과 재풀이 결과를 기준으로 정합니다.",
                "다음 진도로 넘어가기 전 현재 범위의 오답 원인과 완료 기준부터 분명히 해야 합니다.",
            ],
        ),
        (
            "중학생 국영수학원을 찾는 가정이라면 자녀가 모르는 문제를 틀리는지, 알면서도 실수하는지, 또는 질문을 미루는지부터 구분해야 합니다.",
            [
                "중학생의 세 과목 학습을 알아볼 때는 모르는 문제, 반복 실수와 미룬 질문을 먼저 구분해야 합니다.",
                "중등 국영수 상담에서는 개념 공백과 단순 실수, 질문을 미룬 장면을 나누어 확인합니다.",
                "중학생의 현재 상태는 몰라서 틀린 문제와 알고도 실수한 문제를 따로 기록해 살펴야 합니다.",
                "세 과목 계획을 정하기 전 학생이 막힌 문제와 질문하지 못한 내용을 먼저 구분하세요.",
                "중등 학습 점검은 오답의 수보다 개념 부족, 풀이 실수와 질문 지연을 나누는 데서 시작합니다.",
                "국영수 상담 전에는 자녀가 모르는 내용과 반복 실수, 미뤄 둔 질문을 과목별로 정리하세요.",
            ],
        ),
        (
            "국어 수업은 지문을 많이 푸는 것보다 중심 문장, 근거 표시, 서술형 표현을 확인해야",
            [
                "국어 학습은 문제 수보다 중심 문장, 답의 근거와 서술형 표현을 확인해야",
                "국어 수업에서는 많은 지문보다 핵심 내용과 선택지 근거, 서술 과정을 살펴야",
                "국어 점검은 지문 수를 늘리기 전에 중심 문장과 근거 표시, 서술형 답안을 봐야",
                "국어 오답에서는 핵심 문장, 선택지 판단 근거와 서술형 표현을 나누어야",
                "읽기 학습은 문제량보다 본문의 중심 내용과 답의 근거를 확인해야",
                "국어 계획에는 중심 문장 찾기, 근거 표시와 서술형 표현 연습이 들어가야",
            ],
        ),
        (
            "영어 학습에서는 단어 암기량만 보지 말고 문장 구조를 끊어 읽는 순서와 학교 교과서 표현을 함께 점검해야 중학생 내신에 도움이 됩니다.",
            [
                "영어는 단어 수와 함께 문장 구조를 나누고 교과서 표현을 적용하는 과정을 확인해야 합니다.",
                "중등 영어 점검에서는 어휘 암기보다 문장을 끊어 읽고 학교 본문을 해석하는 순서를 봅니다.",
                "영어 학습은 단어 복습, 구문 분석과 교과서 문장 적용을 한 흐름으로 이어야 합니다.",
                "영어 내신을 준비할 때는 단어량뿐 아니라 문장 구조와 학교 본문의 적용 여부를 살펴야 합니다.",
                "중학생 영어 계획에는 어휘, 문장 해석과 교과서 표현의 재확인 날짜가 필요합니다.",
                "영어 수업에서는 암기한 단어를 학교 본문과 문장 구조에 적용하는지 확인합니다.",
            ],
        ),
        (
            "중학생에게는 매번 새로운 문제를 많이 주는 것보다 이미 틀린 문제를 다른 조건으로 다시 설명하게 하는 확인 과정이 필요합니다.",
            [
                "중학생은 새 문제를 늘리기보다 틀린 문제를 조건을 바꿔 다시 설명하는 과정이 필요합니다.",
                "중등 학습에서는 문제량보다 이전 오답을 다른 방식으로 풀고 설명하는지 확인해야 합니다.",
                "새 과제를 추가하기 전 틀린 문제의 조건과 풀이 이유를 다시 말하게 하는 편이 좋습니다.",
                "중학생의 오답 복습은 같은 답을 외우기보다 조건이 달라져도 풀이를 설명하도록 해야 합니다.",
                "문제를 더 풀기 전에 이전 오답을 다시 해석하고 풀이 근거를 말하는 과정을 살펴보세요.",
                "중등 과제는 새 문제와 함께 이미 틀린 유형을 다시 설명하고 적용하는 내용을 담아야 합니다.",
            ],
        ),
        (
            "중학생 국영수학원을 알아볼 때 오답 분석표가 있다면 국어·영어·수학 각각의 실수 유형이 분리되어 있는지 확인해 보세요.",
            [
                "중등 국영수 상담에서는 오답 기록이 과목별 원인과 다음 행동으로 나뉘는지 확인하세요.",
                "세 과목 오답표를 볼 때는 국어·영어·수학의 실수 유형이 따로 기록되는지 살펴봅니다.",
                "오답 분석 자료에는 과목별로 틀린 이유와 다시 확인할 날짜가 구분되어야 합니다.",
                "중학생의 오답 기록은 세 과목의 실수를 같은 항목으로 묶지 않는 편이 정확합니다.",
                "국영수 학습을 비교한다면 오답 원인과 재풀이 결과가 과목마다 따로 남는지 질문하세요.",
                "세 과목을 함께 관리해도 국어·영어·수학의 오류 분류와 복습 기준은 구분해야 합니다.",
            ],
        ),
        (
            "학부모가 학년별로 질문을 나누면 같은 중학생 국영수학원이라도 아이에게 필요한 수업 밀도와 복습량을 더 구체적으로 비교할 수 있습니다.",
            [
                "학부모가 학년과 시험 일정에 맞춰 질문을 나누면 필요한 수업 밀도와 복습량을 구체화할 수 있습니다.",
                "중학생 상담은 학년별 범위와 현재 습관을 구분해 물어야 주간 학습량을 현실적으로 정할 수 있습니다.",
                "학년과 과목별 질문을 따로 준비하면 학생에게 맞는 수업 속도와 복습 기준을 비교하기 쉽습니다.",
                "학부모는 현재 학년의 시험 범위와 가정 복습 시간을 함께 놓고 수업 밀도를 확인할 수 있습니다.",
                "중등 과정은 학년만 묻기보다 학교 일정, 과목별 공백과 필요한 복습량을 함께 확인해야 합니다.",
                "상담 질문을 학년과 과목으로 나누면 학생이 감당할 수 있는 분량과 점검 주기를 정하기 쉽습니다.",
            ],
        ),
    ]
    for frame_index, (source, choices) in enumerate(semantic_frames):
        if source in text:
            text = text.replace(source, stable_pick(seed, f"semantic-frame-{frame_index}", choices))
    middle_grade_exam = re.compile(
        r"(?P<grade>중[1-3])\s+학생에게는\s+시험\s+범위가\s+본격적으로\s+넓어지는\s+"
        r"시기인\s+만큼\s+국어\s+문법,\s*영어\s+문법,\s*수학\s+함수·도형\s+같은\s+"
        r"누적\s+단원의\s+빈틈\s+확인이\s+중요합니다\."
    )

    def replace_middle_grade_exam(match: re.Match[str]) -> str:
        grade = match.group("grade")
        return stable_pick(seed, "middle-grade-exam", [
            f"{grade} 과정에서는 시험 범위가 넓어지므로 세 과목의 누적 단원을 따로 확인해야 합니다.",
            f"{grade} 학생은 국어 문법, 영어 구문과 수학 함수·도형의 공백을 시험 전에 나누어 점검해야 합니다.",
            f"시험 범위가 늘어나는 {grade} 시기에는 세 과목의 이전 단원과 현재 진도를 함께 살펴야 합니다.",
            f"{grade} 학습 계획에는 국어·영어·수학의 누적 공백을 확인할 날짜를 과목별로 남겨야 합니다.",
            f"{grade} 시험 준비는 현재 범위와 더불어 세 과목에서 이어지는 기초 단원의 빈틈을 확인하는 과정이 필요합니다.",
            f"국어 문법, 영어 문장 구조와 수학 누적 개념은 {grade} 시험 일정에 맞춰 따로 복습해야 합니다.",
            f"{grade} 학생의 시험 범위를 볼 때는 과목별 누적 단원과 반복 오답을 함께 표시하는 편이 좋습니다.",
            f"시험 과목과 범위가 늘어나는 {grade} 과정에서는 이전 단원의 빈칸부터 찾아 보완 순서를 정합니다.",
        ])

    text = middle_grade_exam.sub(replace_middle_grade_exam, text)
    parent_weekly = re.compile(
        r"부모가\s+매일\s+모든\s+내용을\s+관리하기는\s+어렵지만\s+"
        r"(?P<locality>[^.!?\n]{1,30})\s+중학생의\s+주간\s+학습\s+기록을\s+"
        r"한\s+번씩\s+같이\s+보면\s+수업\s+방향을\s+조정할\s+단서가\s+생깁니다\."
    )

    def replace_parent_weekly(match: re.Match[str]) -> str:
        area = clean(match.group("locality"))
        return stable_pick(seed, "middle-parent-weekly", [
            f"가정에서는 {area} 중학생의 주간 기록을 정기적으로 살펴 과제와 복습이 끊긴 지점을 확인할 수 있습니다.",
            f"{area} 학부모는 매일 개입하기보다 한 주의 완료 기록과 오답 재풀이 결과를 함께 확인하는 편이 좋습니다.",
            f"주간 학습표를 보면 {area} 학생의 과제 충돌과 복습 공백을 찾아 다음 수업 계획에 반영할 수 있습니다.",
            f"{area} 중학생의 공부를 도울 때는 매일 내용을 확인하기보다 주간 실행 기록을 놓고 변화한 부분을 살펴보세요.",
            f"가정에서 {area} 학생의 과목별 시작 시각과 완료 여부를 주 단위로 보면 조정할 학습량이 분명해집니다.",
            f"{area} 학부모는 한 주에 한 번 과제, 오답과 질문 기록을 함께 보며 다음 확인 항목을 정할 수 있습니다.",
            f"매일 공부를 지시하기보다 {area} 학생의 주간 기록에서 유지할 행동과 바꿀 행동을 나누는 편이 좋습니다.",
            f"{area} 중학생의 수업 방향은 주간 완료 내용과 다시 틀린 문제를 함께 볼 때 구체적으로 조정할 수 있습니다.",
        ])

    text = parent_weekly.sub(replace_parent_weekly, text)
    parent_difficulty = re.compile(
        r"[^.!?\n]*학부모님이\s+체감하는\s+어려움\s+중\s+하나는\s+"
        r"이\s+유형의\s+학생에게\s+같은\s+숙제를\s+반복해도\s+결과가\s+일정하지\s+않다는\s+점입니다\."
    )
    text = parent_difficulty.sub(stable_pick(seed, "parent-difficulty", [
        "같은 과제를 반복해도 결과가 달라진다면 학생이 멈춘 단계와 다시 푼 날짜를 함께 기록해야 합니다.",
        "학습 결과가 들쭉날쭉할 때는 과제량보다 틀린 이유와 재확인 시점을 먼저 살펴보세요.",
        "비슷한 문제에서 결과가 달라지는 학생은 풀이 과정과 복습 간격을 나누어 확인할 필요가 있습니다.",
        "숙제를 반복해도 정확도가 안정되지 않는다면 개념 이해와 실제 적용 단계를 따로 점검해야 합니다.",
        "학생의 수행이 일정하지 않을 때는 공부 시간, 오답 원인과 다시 푼 결과를 한 기록에서 비교하세요.",
        "반복 과제만 늘리기 전에 학생이 문제를 읽고 풀이를 시작하는 과정을 구체적으로 살펴야 합니다.",
        "같은 유형에서 실수가 이어진다면 설명 직후와 며칠 뒤의 재풀이 결과를 나누어 확인하세요.",
        "과제 결과의 편차가 크면 학생의 시작 시각과 완료 기준, 오답 복습 순서를 함께 조정해야 합니다.",
    ]), text)
    elementary_semantic_banks: list[tuple[str, list[str]]] = [
        (
            "과제의 양보다 오답을 다시 만나는 주기",
            [
                "초등 과제는 분량보다 틀린 문제를 며칠 뒤 다시 풀었는지 확인하는 주기가 중요합니다.",
                "학습 기록을 볼 때는 문제 수보다 오답 재풀이 날짜와 질문한 내용을 살펴보세요.",
                "과제를 많이 내는지보다 틀린 문제를 다시 설명하고 풀 기회가 있는지 확인해야 합니다.",
                "초등 복습에서는 과제량보다 오답을 다시 만나는 간격과 확인 방식이 더 중요합니다.",
                "학부모는 문제집 분량보다 아이가 틀린 이유를 남기고 다시 푼 기록을 확인할 수 있습니다.",
                "수업 뒤에는 문제 수를 늘리기보다 오답을 언제 다시 보고 질문했는지 기록하는 편이 좋습니다.",
                "과제의 효과는 양이 아니라 틀린 문제를 다시 풀고 설명하는 과정에서 확인할 수 있습니다.",
                "초등학생에게는 많은 숙제보다 오답을 짧게 반복하고 질문을 남기는 흐름이 도움이 됩니다.",
            ],
        ),
        (
            "영어는 단어장을 외우는 양만 확인하기보다",
            [
                "영어 어휘는 외운 개수보다 문장 안에서 뜻을 찾아 쓰는 과정까지 확인해야 합니다.",
                "단어장 분량만 세기보다 아이가 새 단어를 문장 속에서 설명할 수 있는지 살펴보세요.",
                "영어는 암기한 단어 수와 함께 읽기 지문에서 뜻을 다시 꺼내는 힘을 확인해야 합니다.",
                "어휘 학습은 목록을 외우는 데서 끝내지 말고 문장 해석과 짧은 쓰기로 이어져야 합니다.",
                "영어 단어를 점검할 때는 개수보다 문맥에서 의미를 찾고 활용하는 과정을 보아야 합니다.",
                "아이의 어휘력은 외운 양보다 읽은 문장에서 단어 뜻을 연결하는 모습으로 확인할 수 있습니다.",
                "단어 복습에서는 암기량과 더불어 예문을 읽고 직접 써 보는 단계가 필요합니다.",
                "영어 기초를 볼 때는 단어장 진도보다 배운 표현을 문장에 적용하는지 확인하세요.",
            ],
        ),
        (
            "이 방식은 광고처럼 빠른 변화를 약속하기보다",
            [
                "학습 변화는 단기간의 약속보다 아이가 다음 문제를 혼자 시작하는지 꾸준히 확인할 때 드러납니다.",
                "빠른 결과를 기대하기보다 아이가 배운 내용을 다음 과제에 적용하는 과정을 살펴야 합니다.",
                "초등 학습은 짧은 기간의 성과보다 스스로 시작하고 끝내는 행동이 이어지는지가 중요합니다.",
                "변화의 기준은 큰 약속이 아니라 아이가 다음 문제에서 같은 방법을 다시 사용하는지입니다.",
                "학습 효과는 즉각적인 결과보다 수업 뒤 혼자 풀 수 있는 범위가 넓어지는지로 확인합니다.",
                "한 번의 성과를 앞세우기보다 아이의 질문과 재풀이 행동이 꾸준히 이어지는지 살펴보세요.",
                "초등학생의 변화는 문제 수보다 배운 순서를 혼자 다시 실행하는 모습에서 확인할 수 있습니다.",
                "결과를 서두르기보다 아이가 도움받은 방법을 다음 학습에 적용하는 시간을 지켜봐야 합니다.",
            ],
        ),
        (
            "아이가 수업 후 집에서 무엇을 해야 하는지 짧게라도 말할 수 있다면",
            [
                "수업 뒤 아이가 집에서 할 일을 한 문장으로 설명할 수 있는지 확인하면 피드백의 실효성을 알 수 있습니다.",
                "아이가 귀가 후 복습 순서를 직접 말할 수 있다면 수업 안내가 실제 행동으로 이어진 것입니다.",
                "피드백은 아이가 다음에 풀 문제와 다시 볼 내용을 스스로 설명할 수 있을 때 도움이 됩니다.",
                "수업을 마친 아이에게 다음 공부 순서를 물어보면 안내 내용이 충분히 구체적이었는지 알 수 있습니다.",
                "아이 스스로 집에서 이어 갈 과제를 말할 수 있어야 학부모도 필요한 부분만 확인할 수 있습니다.",
                "수업 후에는 아이가 복습할 내용과 시작 시각을 짧게 설명하도록 해 실행 가능성을 살펴보세요.",
                "피드백의 기준은 아이가 귀가 뒤 무엇을 먼저 할지 알고 있는지로 확인할 수 있습니다.",
                "학부모는 수업 내용을 모두 묻기보다 아이가 다음 행동을 구체적으로 말하는지 살펴볼 수 있습니다.",
            ],
        ),
        (
            "중등 대비라는 말도",
            [
                "중학교 준비는 이름만 앞세우기보다 아이가 질문하고 오답을 다시 보는 생활 습관으로 이어져야 합니다.",
                "중등 준비는 선행 진도보다 읽기, 풀이와 복습을 혼자 시작하는 힘을 기르는 과정입니다.",
                "초등 고학년에는 중학교 과목 수보다 스스로 질문하고 틀린 문제를 다시 보는 습관을 먼저 살펴야 합니다.",
                "중학교를 준비할 때는 진도 계획과 함께 아이가 공부 순서를 정하고 마치는 연습이 필요합니다.",
                "중등 과정으로 넘어가기 전에는 과제량보다 질문 기록과 오답 복습이 생활 안에 자리 잡았는지 확인하세요.",
                "학년 전환기에는 선행 범위보다 읽은 내용을 설명하고 다시 푸는 습관을 안정시키는 편이 좋습니다.",
                "중학교 준비의 핵심은 많은 내용을 앞서가기보다 아이가 학습 계획을 직접 실행하는 데 있습니다.",
                "초등에서 중등으로 이어지는 시기에는 공부를 시작하고 점검하는 일상 루틴을 구체화해야 합니다.",
            ],
        ),
        (
            "자기주도학습은 혼자 두는 방식이 아니라",
            [
                "자기주도학습은 아이를 혼자 두는 일이 아니라 도움받을 단계와 스스로 할 단계를 구분하는 과정입니다.",
                "스스로 공부하는 힘은 모든 일을 맡기는 대신 시작과 점검 기준을 차례로 익힐 때 자랍니다.",
                "자기주도 습관은 아이가 혼자 버티게 하기보다 도움을 요청하고 다시 시도하는 순서를 배우는 데서 시작합니다.",
                "초등학생의 자율 학습은 필요한 안내를 받은 뒤 정해진 과제를 스스로 마치는 연습이어야 합니다.",
                "혼자 공부하게 두기보다 아이가 할 수 있는 부분과 확인이 필요한 부분을 나누는 과정이 중요합니다.",
                "자기주도성은 도움 없이 해결하는 능력보다 질문할 때와 스스로 시도할 때를 구분하는 힘에 가깝습니다.",
                "아이의 독립적인 학습은 작은 과제를 직접 시작하고 완료 여부를 확인하는 경험으로 길러집니다.",
                "스스로 공부하는 과정에도 명확한 시작 기준, 질문 방법과 마무리 점검이 함께 있어야 합니다.",
            ],
        ),
    ]
    if PROFILE == "elementary":
        for bank_index, (marker, choices) in enumerate(elementary_semantic_banks):
            pattern = re.compile(rf"[^.!?\n]*{re.escape(marker)}[^.!?\n]*\.")
            occurrence = 0

            def replace_elementary_semantic(match: re.Match[str]) -> str:
                nonlocal occurrence
                label = f"elementary-semantic-{bank_index}-{occurrence}"
                occurrence += 1
                return " " + stable_pick(seed, label, choices)

            text = pattern.sub(replace_elementary_semantic, text)
    prefixes = list(TRANSITION_PREFIXES[:12])
    markers = [
        "학부모님은 국어·영어·수학을 모두 맡긴다는 표현보다",
        "방문 상담을 잡을 때는 주소 표기",
        "방문 상담을 고려한다면 제공 주소",
        "학부모님이 체감하는 어려움 중 하나는",
        "그리고 선지 판단 이유를 말로 설명해 보게 하면",
        "그리고 풀이가 길어질수록 어느 단계에서 흔들리는지",
        "그리고 문법을 독해 속에 넣어 점검하면",
        "영어 수업은 단어 암기량만 늘리기보다",
        "수학이 흔들리는 학생은 풀이가 길어질수록",
        "국어 수업은 지문을 오래 읽는 훈련보다",
        "영어가 흔들리는 학생은 문법을 독해 속에",
        "국어가 흔들리는 학생은 선지 판단 이유를",
        "수학 수업은 공식 암기 다음에는",
        "국영수학원을 찾는 가정이라면 자녀가 모르는 문제를",
        "영어 학습에서는 단어 암기량만 보지 말고",
        "매번 새로운 문제를 많이 주는 것보다 이미 틀린 문제를",
        "국영수학원을 알아볼 때 오답 분석표가 있다면",
        "학부모가 학년별로 질문을 나누면",
        "시험 범위가 본격적으로 넓어지는 시기인 만큼",
        "부모가 매일 모든 내용을 관리하기는 어렵지만",
        "과제의 양보다 오답을 다시 만나는 주기",
        "영어는 단어장을 외우는 양만 확인하기보다",
        "아이가 수업 후 집에서 무엇을 해야 하는지 짧게라도 말할 수 있다면",
        "이 방식은 광고처럼 빠른 변화를 약속하기보다",
        "중등 대비라는 말도",
        "자기주도학습은 혼자 두는 방식이 아니라",
        "세 과목의 분량을 같게 나누기보다",
        "세 과목을 함께 배우더라도 국어·영어·수학의 진단과 피드백 기준은",
        "같은 시간표 안에서도 과목별 오답 원인과 재확인 순서는",
        "국어는 답의 근거, 영어는 문장 해석, 수학은 풀이 과정을",
        "과목을 묶어 계획해도 읽기, 해석, 풀이에서 막히는 지점은",
        "중요한 것은 세 과목을 다룬다는 말보다 과목별 완료 기준과 복습 날짜가",
        "국영수 진단은 점수만 보는 방식으로는 부족하고",
        "학습 환경을 고려하면 학교 일정, 과제 마감, 귀가 후 복습 시간을",
        "학생에게 맞는 처방은 선행 속도를 높이는 계획이 아니라",
        "학부모에게 필요한 피드백은 단순히 잘했다는 말보다",
        "학부모가 교재를 볼 때는 두꺼운 책의 권수보다",
    ]
    for marker_index, marker in enumerate(markers):
        pattern = re.compile(rf"[^.!?\n]*{re.escape(marker)}[^.!?\n]*\.")
        occurrence = 0

        def replace(match: re.Match[str]) -> str:
            nonlocal occurrence
            label = f"repeated-sentence-{marker_index}-{occurrence}"
            occurrence += 1
            digest = int(hashlib.sha256(f"{seed}|{label}".encode("utf-8")).hexdigest()[:8], 16)
            # Most source sentences already stand on their own. A transition
            # on every sentence created a stronger template signal than the
            # original copy, so retain one only for a small deterministic set.
            prefix = stable_pick(seed, label, prefixes) if digest % 3 == 0 else ""
            return " " + prefix + match.group(0).lstrip()

        text = pattern.sub(replace, text)
    return text


def limit_transition_prefixes(text: str, limit: int = 2) -> str:
    """Keep at most a few unique stock transitions in one manuscript."""
    alternation = "|".join(re.escape(prefix.strip()) for prefix in TRANSITION_PREFIXES)
    pattern = re.compile(
        rf"(?P<lead>^|[.!?]\s+|\n)(?P<prefix>{alternation})\s*",
        flags=re.MULTILINE,
    )
    used: set[str] = set()

    def replace(match: re.Match[str]) -> str:
        prefix = match.group("prefix")
        if prefix in used or len(used) >= limit:
            return match.group("lead")
        used.add(prefix)
        return f"{match.group('lead')}{prefix} "

    return pattern.sub(replace, text)


def diversify_school_paragraphs(paragraphs: list[str], seed: str) -> list[str]:
    prefixes = list(TRANSITION_PREFIXES[12:])
    result: list[str] = []
    used_prefixes: set[str] = set()
    for paragraph_index, paragraph in enumerate(paragraphs):
        sentences = [part for part in re.split(r"(?<=[.!?])\s+", paragraph) if part]
        varied: list[str] = []
        previous_prefix = ""
        for sentence_index, sentence in enumerate(sentences):
            prefix = stable_pick(seed, f"school-prefix-{paragraph_index}-{sentence_index}", prefixes)
            while prefix in used_prefixes and len(used_prefixes) < len(prefixes):
                prefix = prefixes[(prefixes.index(prefix) + 1) % len(prefixes)]
            varied.append(prefix + sentence)
            previous_prefix = prefix
            used_prefixes.add(prefix)
        result.append(" ".join(varied))
    return result


def polish_korean(text: str, locality: str = "") -> str:
    replacements = {
        "학원와": "학원과",
        "습관를": "습관을",
        "학원등원": "학원 등원",
        "학원차량": "학원 차량",
        "학원를": "학원을",
        "점검를": "점검을",
        "일정를": "일정을",
        "일정와": "일정과",
        "학원교통": "통학 여건",
        "학습동기관리": "학습 동기 관리",
        "국영수 학습 안내문": "국어·영어·수학 학습 계획",
        "방식를": "방식을",
        "내용를": "내용을",
        "과정를": "과정을",
        "문제집를": "문제집을",
        "학교 숙제 수행 시간를": "학교 숙제 수행 시간을",
        "최근 학교 단원를": "최근 학교 단원을",
        "반복되는 유형를": "반복되는 유형을",
        "학습성과관리은": "학습성과관리는",
        "학원안전관리을": "학원안전관리를",
        "와와학습코칭학원로": "와와학습코칭학원으로",
        "자료와 자료": "자료",
        "자료를 자료": "자료",
        "학원운영자": "학원 운영자",
        "고1식 공부법에서 고등 내신형 공부로 바꿔야": "중학교 때의 공부 습관을 고등 내신에 맞게 조정해야",
        "중등식 문제량": "현재 수준보다 많은 문제량",
        "수 있은": "수 있는",
        "묻은 질문": "묻는 질문",
        "찾은 자리": "찾는 자리",
        "함께 읽은 것입니다": "함께 읽는 것입니다",
        "맞은 순서": "맞는 순서",
        "찾은 가정": "찾는 가정",
        "찾은 학부모": "찾는 학부모",
        "에 대한 설명는": "에 대한 설명은",
        "기록라는 표현": "기록이라는 표현",
        "학부모님이 체감하는 현실은 학부모님이": "학부모님이 체감하는 현실은",
        "관리 기준 방문 전에는": "관리 기준을 확인하려면 방문 전",
        "하교 후 가능한 요일를": "하교 후 가능한 요일을",
        "가정에서 확인 가능한 복습 시간를": "가정에서 확인 가능한 복습 시간을",
        "현재 교재 진도을": "현재 교재 진도를",
        "시험·교과 범위을": "시험·교과 범위를",
        "단원별 이해도을": "단원별 이해도를",
        "재풀이 결과을": "재풀이 결과를",
        "틀린 이유 분류을": "틀린 이유 분류를",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    if locality:
        text = text.replace(
            f"{locality} 국어·영어·수학 학습 계획 상담 전 준비할 질문",
            f"{locality} 국어·영어·수학 상담 전 준비할 질문",
        )
    text = re.sub(r"(\d+)층로", r"\1층으로", text)
    text = re.sub(r"(?<![가-힣])([가-힣]+학원)로(?=\s|[,.!?)]|$)", r"\1으로", text)
    # 원고 생성 과정에서 같은 명사가 연달아 붙은 표현을 조사까지 보존해
    # 한 번만 남긴다. 예: `자료 자료를` -> `자료를`, `상담 상담에서는` -> `상담에서는`.
    repeated_terms = (
        "자료|상담|시기|기준|운영|계획|학습|학생|수업|방법|내용|과제|관리|확인|"
        "진단|과정|습관|시험|복습|오답|학교|학원|과목|설명|기록"
    )
    repeated_pattern = re.compile(rf"(?P<word>{repeated_terms})\s+(?P=word)")
    previous = None
    while previous != text:
        previous = text
        text = repeated_pattern.sub(r"\g<word>", text)
    double_need = re.compile(
        r"(?P<first>[^,.!?\n]{2,100}?)(?:이|가)\s+"
        r"(?P<priority>먼저\s+)?필요한\s+"
        r"(?P<second>[^,.!?\n]{2,50}?)(?:이|가)\s+필요한\s+학생"
    )

    def replace_double_need(match: re.Match[str]) -> str:
        first = clean(match.group("first"))
        second = clean(match.group("second"))
        priority = match.group("priority") or ""
        return (
            f"{attach_particle(first, '이')} {priority}필요하고 "
            f"{second}도 필요한 학생"
        )

    text = double_need.sub(replace_double_need, text)
    base = r"(?:학습 계획|학습 점검|학습 과정|세 과목 학습 상담|관리 기준)"
    text = re.sub(
        rf"(?P<base>{base})\s+방문\s+전에는",
        lambda match: f"{attach_particle(match.group('base'), '을')} 알아보려면 방문 전에",
        text,
    )
    text = re.sub(
        rf"(?P<base>{base})\s+상담에서\s+(?P<subject>국어|영어|수학)\s+계획",
        lambda match: (
            f"{attach_particle(match.group('base'), '을')} 살펴보는 상담에서 "
            f"{match.group('subject')} 계획"
        ),
        text,
    )
    text = re.sub(
        rf"(?P<base>{base})\s+수업\s+후\s+복습",
        lambda match: f"{attach_particle(match.group('base'), '을')} 진행한 뒤 복습",
        text,
    )
    text = re.sub(
        rf"(?P<base>{base})\s+수업은\s+진단",
        lambda match: f"{attach_particle(match.group('base'), '은')} 진단",
        text,
    )
    text = re.sub(
        rf"(?P<base>{base})\s+수업\s+뒤에\s+무엇",
        lambda match: f"{attach_particle(match.group('base'), '을')} 마친 뒤 무엇",
        text,
    )
    text = re.sub(
        rf"(?P<base>{base})\s+방문\s+상담\s+주소는",
        lambda match: f"{attach_particle(match.group('base'), '을')} 확인할 때 방문 상담 주소는",
        text,
    )
    text = re.sub(
        rf"(?P<base>{base})\s+교재\s+점검에서는",
        lambda match: f"{attach_particle(match.group('base'), '에')} 맞춘 교재 점검에서는",
        text,
    )
    # 일반 단어·지역명·도로명까지 훼손하는 전역 조사 정규식은 사용하지
    # 않는다. 위의 확인된 오류 표현만 보수적으로 교정한다.
    return text


def supported_grades(center: dict) -> set[str]:
    """Return grades verified for every advertised subject on a combined page."""
    subject_sets = [
        set(center["korean_grades"]),
        set(center["english_grades"]),
        set(center["math_grades"]),
    ]
    non_empty = [values for values in subject_sets if values]
    return set.intersection(*non_empty) if non_empty else set()


def verified_student_type(student_type: str, center: dict, seed: str) -> str:
    """Keep source individuality while removing grade claims not supported by center data."""
    value = naturalize_student_type(polish_korean(student_type, center["locality"]))
    mentioned = set(re.findall(rf"{re.escape(GRADE_PREFIX)}[1-6]", value))
    low_grade_claim = PROFILE == "elementary" and "저학년" in value
    permitted = supported_grades(center)
    unsupported = bool(mentioned - permitted)
    if low_grade_claim and not ({"초1", "초2"} & permitted):
        unsupported = True
    if not value or unsupported:
        alternatives = {
            "high": [
                "내신 범위와 모의고사 오답을 한 계획 안에 배치하기 어려운 고등학생",
                "과목별 시험 일정은 알고 있지만 복습 우선순위를 정하기 어려운 고등학생",
                "한 과목의 과제량 때문에 다른 과목의 오답 정리가 자주 밀리는 고등학생",
                "최근 시험 결과를 다음 주 학습 계획으로 바꾸는 과정이 필요한 고등학생",
                "국어·영어·수학의 공부 시간은 많지만 과목별 완료 기준이 불분명한 고등학생",
                "학교 자료와 현재 교재의 진도를 함께 관리하기 어려운 고등학생",
            ],
            "middle": [
                "교과 개념은 배웠지만 학교 시험 문제에 적용하는 과정이 불안정한 중학생",
                "과제는 시작해도 오답 복습과 재확인 날짜를 스스로 정하기 어려운 중학생",
                "세 과목의 시험 준비가 겹치면 우선순위를 정하지 못하는 중학생",
                "수업에서 이해한 내용을 혼자 다시 설명하고 적용하는 연습이 필요한 중학생",
                "학교 일정과 학원 과제를 한 주 계획 안에 배치하기 어려운 중학생",
                "국어·영어·수학마다 막히는 원인이 다른데 같은 방식으로 공부하는 중학생",
            ],
            "elementary": [
                "읽기 이해와 어휘, 계산 과정의 기초를 과목별로 점검할 필요가 있는 초등학생",
                "숙제를 시작하는 시각과 끝내는 기준을 스스로 정하기 어려운 초등학생",
                "배운 내용을 말로 설명하고 틀린 문제를 다시 푸는 습관이 필요한 초등학생",
                "교과 진도는 따라가지만 국어·영어·수학의 기초 과정이 고르지 않은 초등학생",
                "알림장과 교재를 보고 그날 공부할 순서를 정하는 연습이 필요한 초등학생",
                "정답을 맞히는 것보다 풀이 과정과 읽은 근거를 남기는 연습이 필요한 초등학생",
            ],
        }
        value = stable_pick(seed, "verified-student-type", alternatives[PROFILE])
    return value.rstrip("., ")


def direct_answer(locality: str, center: dict, student_type: str, seed: str) -> str:
    center_name = center["center_name"] or f"{locality} 지역 센터"
    school_context = (
        f"제공된 {SCHOOL_NAME} 참고 자료와 학생의 실제 교재를 대조하고"
        if center["schools"]
        else "학생이 사용하는 학교 자료와 현재 교재를 직접 확인하고"
    )
    endings = [
        "과목별로 무엇을 먼저 보완할지 정하는 순서가 필요합니다.",
        "국어·영어·수학의 막힌 원인을 나눈 뒤 실행 가능한 주간 계획을 정해야 합니다.",
        "세 과목을 같은 분량으로 늘리기보다 과목별 완료 기준과 재확인 날짜를 정해야 합니다.",
        "현재 자료에서 확인한 문제를 수업, 복습과 재점검으로 연결하는지 살펴봐야 합니다.",
    ]
    return (
        f"{locality}에서 {CATEGORY_NAME}을 찾는다면 {center_name}의 과목별 가능 학년을 먼저 확인하세요. "
        f"{school_context}, 학생의 실제 풀이 기록에 맞게 "
        f"{stable_pick(seed, 'direct-answer-ending', endings)}"
    )


def answer_cards(locality: str, center: dict, seed: str) -> list[tuple[str, str]]:
    school_label = f"{SCHOOL_NAME} 자료" if center["schools"] else "현재 교재"
    options = {
        "high": [
            ("01 / 내신 범위", f"{school_label}와 수행평가 일정을 과목별 계획으로 바꿉니다"),
            ("02 / 모의고사", "틀린 문항을 개념·해석·풀이 단계로 나눠 다시 확인합니다"),
            ("03 / 시간 배분", "시험 일정에 맞춰 집중 과목과 최소 복습량을 구분합니다"),
        ],
        "middle": [
            ("01 / 개념 연결", "배운 개념을 학교 시험 문제에 적용하는 과정을 확인합니다"),
            ("02 / 학교 일정", f"{school_label}와 시험 범위를 주간 과제에 연결합니다"),
            ("03 / 실행 점검", "과제 완료와 오답 재풀이가 스스로 이어지는지 살펴봅니다"),
        ],
        "elementary": [
            ("01 / 기초 이해", "읽기·어휘·계산 과정에서 멈추는 지점을 나눠 봅니다"),
            ("02 / 교과 자료", f"{school_label}에서 현재 단원과 복습 범위를 확인합니다"),
            ("03 / 공부 습관", "시작 시각과 완료 기준, 다시 풀 날짜를 짧게 기록합니다"),
        ],
    }
    cards = [
        (
            label,
            f"{answer}. " + stable_pick(seed, f"answer-card-detail-{index}", [
                "최근 교재와 대조해 확인합니다.",
                "학생의 실제 기록에서 살펴봅니다.",
                "다음 점검 날짜까지 함께 정합니다.",
                "학교 일정과 나란히 놓고 봅니다.",
                "주간 계획에 반영되는지 확인합니다.",
                "과제 완료 기록으로 다시 점검합니다.",
                "상담 때 구체적인 사례를 들어 묻습니다.",
                "수업 뒤 재확인 방식까지 살펴봅니다.",
            ])
        )
        for index, (label, answer) in enumerate(options[PROFILE])
    ]
    shift = int(hashlib.sha256(f"{seed}|answer-card".encode("utf-8")).hexdigest()[:4], 16) % len(cards)
    return cards[shift:] + cards[:shift]


def extract_student_type(body: str, locality: str) -> str:
    patterns = [
        r"이 페이지는 (.+?)을 기준 학생으로 두고,",
        r"대표 학생 유형은 (.+?)입니다\.",
        r"이 페이지에서 가정한 학생은 (.+?학생)(?=(?:이며|이고|으로|입니다))",
        r"상황이라면 (.+?학생)(?=(?:이며|이고|으로|입니다))",
        r"(?:^|\n|[.!?]\s+)([^.!?\n]{10,180}?학생)을 위한 진단은",
    ]
    for pattern in patterns:
        match = re.search(pattern, body, re.DOTALL)
        if match:
            value = polish_korean(clean(match.group(1)), locality)
            value = re.sub(rf"^{re.escape(locality)}\s*", "", value)
            if len(value) > 130:
                pieces = [piece.strip() for piece in value.split(",") if piece.strip()]
                value = ", ".join(pieces[:2])
            return value.rstrip("., ")
    return f"세 과목의 과제 순서와 오답 복습 시점을 스스로 정리하기 어려운 {LEVEL_NAME}"


def extract_elementary_behavior(body: str) -> str:
    """Return the second repeated behavior clause used by elementary drafts."""
    if PROFILE != "elementary":
        return ""
    match = re.search(
        r"상황이라면 .+?학생(?:이며|이고|으로),\s*(.+?)\s+(?:아이|학생)인지",
        body,
        flags=re.S,
    )
    return clean(match.group(1)) if match else ""


def naturalize_student_type(value: str) -> str:
    """Turn source-only grade joins into a natural standalone student phrase."""
    value = clean(value)
    value = value.replace(
        "실제 독해에 적용하는 속도가 느린 방학 복습이 필요한",
        "실제 독해에 적용하는 속도가 느려 방학 복습이 필요한",
    )
    value = value.replace(
        "짧은 단위 복습에는 반응이 좋은 학습 공백을 줄이고 싶은",
        "짧은 단위 복습에는 잘 반응하지만 학습 공백을 줄이고 싶은",
    )
    value = value.replace(
        "집중 시간은 짧지만 짧은 단위 복습에는 잘 반응하지만 학습 공백을 줄이고 싶은",
        "집중 시간은 짧아도 짧은 단위 복습에는 잘 반응해 학습 공백을 줄이고 싶은",
    )
    if PROFILE == "elementary":
        value = re.sub(r"(초[1-6])\s+중\b", r"\1 학생 가운데", value)
        value = re.sub(
            r"(초[1-6]에서\s+초[1-6](?:으)?로\s+(?:이어지는|올라가는)\s+시기)\s+중",
            r"\1에",
            value,
        )
        value = value.replace(
            "초등 과정이 되며 중등 준비 질문이 늘어나는",
            "중등 진학을 앞두고 준비 질문이 늘어나는",
        )
        match = re.fullmatch(
            r"초등\s+((?:[1-6](?:\s*[~～\-–—]\s*[1-6])?)학년)\s+중\s+(.+?)\s+학생",
            value,
        )
        if match:
            grade, description = clean(match.group(1)), clean(match.group(2))
            value = f"{description} 초등 {grade} 학생"
    return value


def normalize_student_intro(text: str, locality: str, seed: str) -> str:
    """Replace production-oriented subject frames with complete reader prose."""
    if PROFILE == "high":
        pattern = re.compile(
            rf"이\s+페이지는\s+(.+?고등학생)을\s+기준\s+학생으로\s+두고,\s*"
            rf"{re.escape(locality)}에서\s+국어·영어·수학을\s+어떻게\s+나누어\s+점검하면\s+좋은지\s+답변합니다\."
        )
        def replace_high(match: re.Match[str]) -> str:
            student = naturalize_student_type(match.group(1))
            follow_up = stable_pick(seed, "high-student-intro-follow-up", [
                f"{locality}에서는 최근 시험지로 세 과목의 보완 순서를 나누어 확인합니다.",
                f"{locality} 상담에서는 현재 교재와 오답을 과목별로 펼쳐 우선순위를 정합니다.",
                f"최근 학습 기록을 바탕으로 {locality} 학생의 국어·영어·수학 병목을 따로 살펴봅니다.",
                f"{locality} 학부모는 학교 일정과 복습 시간을 함께 놓고 세 과목 계획을 점검할 수 있습니다.",
                f"상담에서는 {locality} 학생이 끝낸 내용과 다시 확인할 범위를 과목별로 구분합니다.",
                f"{locality} 학생의 시험·과제 자료에서 먼저 조정할 과목과 재확인 날짜를 찾습니다.",
                f"세 과목을 함께 계획하더라도 {locality} 학생의 오답 원인과 완료 기준은 따로 확인합니다.",
                f"{locality} 상담의 첫 단계는 최근 풀이에서 과목별로 멈춘 지점을 나누는 일입니다.",
            ])
            return f"{attach_particle(student, '을')} 대표 상담 사례로 살펴봅니다. {follow_up}"

        text = pattern.sub(replace_high, text)
        configured_type = re.compile(
            r"[^.!?]{0,220}?에서\s+설정한\s+대표\s+학생\s+유형은\s+"
            r"(.+?학생)입니다\."
        )
        return configured_type.sub(replace_high, text)
    if PROFILE == "middle":
        pattern = re.compile(
            r"이\s+페이지에서\s+가정한\s+학생은\s+(.+?학생)(?:이며|이고|으로),\s*"
        )
        return pattern.sub(
            lambda match: (
                f"{attach_particle(naturalize_student_type(match.group(1)), '을')} "
                "대표 상담 사례로 살펴봅니다. "
            ),
            text,
        )
    pattern = re.compile(
        r"(?P<lead>[^.!?\n]{5,220}?상황이라면)\s+"
        r"(?P<student>.+?학생)(?:이며|이고|으로),\s*"
        r"(?P<behavior>.+?)\s+(?:아이|학생)인지부터\s+확인해야\s+하며,\s*"
    )

    def replace_elementary(match: re.Match[str]) -> str:
        lead = clean(match.group("lead")).replace("상황이라면", "상황에서는")
        student = naturalize_student_type(match.group("student"))
        behavior = clean(match.group("behavior"))
        return (
            f"{lead} {attach_particle(student, '을')} 대표 사례로 살펴봅니다. "
            f"먼저 {behavior} 모습이 있는지 확인해야 합니다. "
        )

    return pattern.sub(replace_elementary, text)


def reduce_title_repetition(text: str, title: str, locality: str, seed: str, keep: int = 4) -> str:
    alternatives = [
        f"{locality} {LEVEL_NAME}의 국어·영어·수학 학습 점검",
        f"{locality} 국어·영어·수학 학습 계획",
        f"{locality} {LEVEL_NAME} 세 과목 학습 상담",
        f"{locality} {LEVEL_NAME} 학습 과정",
        f"{locality} {LEVEL_NAME} 학습 점검",
        f"{COURSE_NAME} 국어·영어·수학 관리 기준",
    ]
    seen = 0
    index = 0
    def replace(match: re.Match[str]) -> str:
        nonlocal seen, index
        seen += 1
        if seen <= keep:
            return match.group(0)
        choice = stable_pick(seed, f"title-{index}", alternatives)
        index += 1
        return choice
    return re.sub(re.escape(title), replace, text)


def sanitize_grade_claims(text: str, center: dict) -> str:
    """Remove unsupported exact-grade scenarios without inventing a replacement grade."""
    permitted = supported_grades(center)
    all_profile_grades = {
        "high": ["고1", "고2", "고3"],
        "middle": ["중1", "중2", "중3"],
        "elementary": [f"초{number}" for number in range(1, 7)],
    }[PROFILE]
    for grade in all_profile_grades:
        if grade in permitted:
            continue
        text = text.replace(f"{grade} 첫 학기", f"{LEVEL_NAME} 과정 초반")
        text = text.replace(f"{grade} 진입", f"{LEVEL_NAME} 과정 전환")
        text = re.sub(rf"(?<![가-힣0-9]){re.escape(grade)}(?![가-힣0-9])", LEVEL_NAME, text)
    if PROFILE == "elementary":
        # The manuscripts also use ranges such as `초등 1~2학년` and broad
        # bands such as `저학년`. Keep them only when every referenced grade
        # is supported for all three subjects at the matched center.
        def replace_range(match: re.Match[str]) -> str:
            start, end = int(match.group(1)), int(match.group(2))
            claimed = {f"초{number}" for number in range(min(start, end), max(start, end) + 1)}
            return match.group(0) if claimed <= permitted else "초등 과정"

        text = re.sub(r"초등\s*([1-6])\s*[~～\-–—]\s*([1-6])학년", replace_range, text)

        def replace_single(match: re.Match[str]) -> str:
            claimed = f"초{match.group(1)}"
            return match.group(0) if claimed in permitted else "초등 과정"

        text = re.sub(r"초등\s*([1-6])학년", replace_single, text)
        bands = {
            "저학년": {"초1", "초2", "초3"},
            "중학년": {"초3", "초4"},
            "고학년": {"초4", "초5", "초6"},
        }
        for label, claimed in bands.items():
            if not claimed <= permitted:
                text = text.replace(f"초등 {label}", "초등 과정").replace(label, "초등 과정")
    return text


def personalize_body(body: str, title: str, locality: str, center: dict, student_type: str, seed: str) -> tuple[str, list[tuple[str, list[str]]]]:
    source_student_type = extract_student_type(body, locality)
    elementary_student_core = extract_elementary_student_core(source_student_type)
    elementary_grade_frame = extract_elementary_grade_frame(source_student_type)
    elementary_behavior = extract_elementary_behavior(body)
    body, protected_facts = protect_verified_facts(body, center)
    body = use_official_region_in_copy(body, center)
    body = sanitize_unverified_operational_keyword(body, seed, locality)
    body = restore_verified_facts(body, protected_facts)
    body = remove_search_engine_language(body, locality, seed)
    body = polish_korean(body, locality)
    body = body.replace("해당 지역", locality).replace("이 지역", locality)
    body = body.replace(f"{locality}에서 {locality} ", f"{locality}에서 ")
    body = sanitize_grade_claims(body, center)
    body = normalize_student_intro(body, locality, seed)
    body = body.replace("학생이며,", stable_pick(seed, "student-link", ["학생이고,", "학생으로,", "학생입니다. 또한 "]))
    body = body.replace(
        f"{locality} {LEVEL_NAME}에게 {LEVEL_NAME} 국영수학원이",
        f"{locality} {LEVEL_NAME}에게 국어·영어·수학 학습 점검이",
    )
    body = re.sub(r"니다,\s*그리고", "니다. 그리고", body)
    body = body.replace(
        "을 기준 학생으로 두고,",
        stable_pick(seed, "student-frame", ["을 대표 사례로 삼아,", "의 학습 장면을 중심으로,", "을 주요 상담 상황으로 보고,"]),
    )
    body = body.replace(
        "학원 선택은 위치와 시간표만으로 끝내기보다 시험 후 분석이 다음 계획으로 이어지는지 확인하는 과정이 필요합니다.",
        stable_pick(seed, "selection", [
            "학원을 비교할 때는 위치와 시간표뿐 아니라 시험 분석이 다음 주 계획에 실제로 반영되는지 확인해야 합니다.",
            "선택 기준에는 통학 조건과 함께 시험 뒤 분석, 보완 과제, 재점검 일정이 하나의 흐름으로 이어지는지가 포함되어야 합니다.",
            "위치가 편한지만 보기보다 시험 결과를 해석한 뒤 수업과 복습 계획이 어떻게 달라지는지 물어보는 편이 정확합니다.",
            "상담에서는 시간표 확인에 그치지 말고 시험 분석이 과목별 다음 행동으로 연결되는 절차를 살펴봐야 합니다.",
            "학원 선택 전에는 통학 가능성과 더불어 시험 뒤 약점 분류가 다음 계획을 바꾸는지 확인하는 과정이 필요합니다.",
            "비교표에는 거리와 수업 시각 외에도 시험 분석, 보완 순서, 재확인 시점을 함께 기록하는 것이 좋습니다.",
        ]),
    )
    body = body.replace(
        "수업 선택 전에는 정규 수업, 보강, 자습 관리, 과제 피드백이 서로 어떻게 이어지는지 확인해야 합니다.",
        stable_pick(seed, "operation", [
            "수업을 정하기 전에는 정규 수업과 보강, 자습, 과제 피드백이 각각 언제 연결되는지 확인해야 합니다.",
            "정규 진도만 묻지 말고 보강과 자습 관리, 과제 피드백이 다음 수업에 반영되는 순서를 살펴보세요.",
            "수업·보강·자습·과제 피드백이 따로 운영되는지, 하나의 주간 계획 안에서 이어지는지 질문하는 것이 좋습니다.",
            "과제 결과가 보강이나 자습 계획을 어떻게 바꾸는지까지 들어야 실제 관리 흐름을 판단할 수 있습니다.",
            "정규 수업 이후 보강과 자습, 과제 확인이 어떤 기록으로 남는지 상담에서 구체적으로 확인하세요.",
            "수업 형태보다 중요한 것은 정규 진도, 보완 학습과 과제 피드백이 다음 점검으로 이어지는 방식입니다.",
        ]),
    )
    body = reduce_title_repetition(body, title, locality, seed, keep=2)
    body = body.replace(f"{locality} 생활권에서 {locality} ", f"{locality} 생활권에서 ")
    body = diversify_common_copy(body, locality, seed)
    body = diversify_repeated_sentences(body, seed)
    body = limit_transition_prefixes(body, limit=0)
    body = remove_meta_and_mechanical_copy(body, locality, seed)
    long_frame = re.compile(r"(기준으로는 .*?한꺼번에 늘리는 계획보다),\s*(.+?)인지 먼저 확인하는 과정이 더 중요합니다\.")
    match = long_frame.search(body)
    if match:
        lead_ending = stable_pick(seed, "student-frame-lead-ending", [
            "는 최근 과제와 시험 범위를 먼저 살펴야 합니다.",
            "는 현재 교재에서 반복되는 오답부터 확인해야 합니다.",
            "는 실제 공부 시간과 과목별 마감 일정을 먼저 비교해야 합니다.",
            "는 학생이 혼자 끝낼 수 있는 분량부터 파악해야 합니다.",
            "는 최근 풀이에서 멈춘 단원과 원인을 먼저 나눠야 합니다.",
            "는 학교 일정과 재풀이 날짜를 먼저 대조해야 합니다.",
            "는 과목별 현재 수준과 완료 기록부터 살펴야 합니다.",
            "는 이번 주에 밀린 과제와 복습 시간을 먼저 확인해야 합니다.",
            "는 최근 평가 자료에서 보완 순서를 먼저 찾아야 합니다.",
            "는 학생의 시간표에서 확보할 수 있는 복습량부터 계산해야 합니다.",
            "는 과목마다 다시 확인할 내용과 날짜를 먼저 정해야 합니다.",
            "는 최근 학습 기록과 실행 가능성을 먼저 봐야 합니다.",
        ])
        follow_up = stable_pick(seed, "student-frame-follow-up", [
            f"{LEVEL_NAME}의 최근 풀이에서 과목별로 멈추는 지점을 먼저 구분해야 합니다.",
            f"{LEVEL_NAME}의 현재 교재와 오답 기록을 보고 세 과목의 보완 순서를 확인하는 편이 좋습니다.",
            f"{LEVEL_NAME}의 주간 실행 기록에서 우선 조정할 과목과 단원을 찾아야 합니다.",
            f"{LEVEL_NAME}의 최근 시험·과제 자료를 바탕으로 실제 학습 상황을 구체적으로 확인해야 합니다.",
            f"{LEVEL_NAME}의 과목별 완료 시간과 반복 오답을 살펴 현실적인 우선순위를 정해야 합니다.",
            f"{LEVEL_NAME}이 사용하는 자료에서 국어·영어·수학의 병목을 따로 확인해야 합니다.",
            f"{LEVEL_NAME}은 점수 한 줄보다 최근 풀이 과정과 복습 여부를 보고 계획을 조정해야 합니다.",
            f"{LEVEL_NAME}의 현재 범위와 공부 시간을 함께 살펴 이번 주에 바꿀 행동을 정해야 합니다.",
            f"{LEVEL_NAME}의 세 과목 과제 결과를 비교해 먼저 되짚을 내용을 정해야 합니다.",
            f"{LEVEL_NAME}의 실제 시간표와 풀이 기록을 바탕으로 실행 가능한 계획을 세워야 합니다.",
            f"{LEVEL_NAME}이 완료하지 못한 과제와 다시 틀린 문제를 나누어 다음 점검 순서를 잡아야 합니다.",
            f"{LEVEL_NAME}의 현재 단원과 재풀이 결과를 확인해 필요한 학습량을 조정해야 합니다.",
        ])
        body = body[:match.start()] + match.group(1) + lead_ending + " " + follow_up + body[match.end():]
    source_student_type = naturalize_student_type(
        sanitize_grade_claims(polish_korean(source_student_type, locality), center)
    )
    body = reduce_phrase_repetition(body, source_student_type, seed + "|source", keep=1)
    body = reduce_phrase_repetition(body, student_type, seed + "|verified", keep=1)
    body = reduce_elementary_student_core_repetition(
        body,
        polish_korean(elementary_student_core, locality),
        seed,
        keep=0,
    )
    body = reduce_elementary_grade_frame_repetition(
        body,
        polish_korean(elementary_grade_frame, locality),
        seed,
        keep=0,
    )
    body = reduce_elementary_course_label_repetition(body, locality, seed)
    if elementary_behavior:
        body = reduce_elementary_behavior_repetition(
            body,
            polish_korean(elementary_behavior, locality),
        )
    body = shorten_elementary_sentences(body)
    # Grade replacement can expose adjacent template fragments, so run the
    # conservative Korean cleanup once more before turning the source into HTML.
    body = polish_korean(body, locality)
    # `polish_korean` also repairs old noun chains. Some of those mechanical
    # repairs need a final full-sentence rewrite for reader-facing prose.
    body = polish_manual_review_families(body, locality, seed)
    body = polish_residual_student_references(body, locality, seed)
    body = re.sub(
        rf"[^.!?\n]*에서\s+{re.escape(CATEGORY_NAME)}을\s+비교할\s+때는[^.!?\n]*약점\s+기록이\s+따로\s+남는지를\s+보셔야\s+합니다\.",
        stable_pick(seed, "late-reader-facing-comparison", [
            f"{locality}에서는 과목 수보다 국어·영어·수학의 약점 기록과 재확인 기준을 살펴보세요.",
            f"{locality} 수업을 비교할 때는 과목별 오답 원인과 피드백이 구분되는지가 중요합니다.",
            f"세 과목을 함께 계획해도 {locality} 학생의 병목과 복습 기록은 따로 확인해야 합니다.",
            f"{locality} 상담에서는 진단 결과가 과목별 다음 행동으로 이어지는지 질문하세요.",
            f"{locality} 학생의 읽기·해석·풀이 문제를 구분해 기록하는 과정이 필요합니다.",
            f"세 과목 수업을 알아볼 때는 {locality} 학생의 자료에서 보완 순서가 제시되는지 살펴야 합니다.",
            f"{locality} 학부모는 오답 기록과 피드백 방식이 과목별로 나뉘는지 확인할 수 있습니다.",
            f"과목을 함께 관리해도 {locality} 학생의 현재 단원과 재점검 날짜는 따로 정해야 합니다.",
        ]),
        body,
    )
    body = re.sub(
        r"[^.!?\n]*(?:학습 과정|관리 기준)은 국어·영어·수학을 한꺼번에 밀어붙이는 곳보다 진단과 복습, 과제 확인이 연결되는지 살펴보는 편이 좋습니다\.",
        stable_pick(seed, "connected-elementary-process", [
            f"{locality} 초등 학습은 세 과목의 진단 결과가 복습과 과제 확인으로 이어지는지 살펴보세요.",
            f"{locality} 학생에게는 많은 분량보다 국어·영어·수학의 점검과 복습 흐름을 연결하는 일이 중요합니다.",
            f"세 과목을 한꺼번에 늘리기 전 {locality} 학생의 기초 진단과 과제 확인 절차를 비교해야 합니다.",
            f"{locality} 초등 과정은 진도보다 과목별 진단, 복습과 완료 확인이 이어지는지를 먼저 봅니다.",
            f"국어·영어·수학을 함께 계획할 때도 {locality} 학생의 오답과 다음 복습 날짜를 구분해야 합니다.",
            f"{locality} 학생의 세 과목 계획은 현재 기초를 확인하고 짧게 복습하는 흐름부터 갖춰야 합니다.",
            f"초등 학습을 비교한다면 {locality} 학생의 과목별 기초와 과제 완료 기록이 남는지 확인하세요.",
            f"{locality} 상담에서는 분량을 늘리는 약속보다 진단 결과가 다음 과제에 반영되는지 살펴봅니다.",
        ]),
        body,
    )
    body = re.sub(r"([.!?])(?=[가-힣])", r"\1 ", body)
    body = normalize_address_sentences(body, center, locality, seed)
    intro, sections = parse_body(body)
    intro, sections = deduplicate_manuscript_windows(
        intro,
        sections,
        locality,
        student_type,
        seed,
    )
    normalized_sections: list[tuple[str, list[str]]] = []
    seen_paragraphs: set[str] = {intro}
    for heading, paragraphs in sections:
        unique_paragraphs: list[str] = []
        for paragraph in paragraphs:
            if paragraph in seen_paragraphs:
                continue
            seen_paragraphs.add(paragraph)
            unique_paragraphs.append(paragraph)
        if unique_paragraphs:
            normalized_sections.append((simplify_heading(heading), unique_paragraphs))
    sections = normalized_sections
    return intro, sections


def school_section(center: dict, locality: str, seed: str) -> tuple[str, list[str]]:
    schools = center["schools"]
    material_label = SCHOOL_MATERIALS if SCHOOL_MATERIALS.endswith("자료") else f"{SCHOOL_MATERIALS} 자료"
    heading = stable_pick(seed, "school-heading", [
        f"{locality} {SCHOOL_NAME} 자료를 수업 계획에 반영하는 방법",
        f"{locality} 학생의 학교 자료를 확인하는 순서",
        f"{locality} {LEVEL_NAME} 학교 자료와 과목별 계획 연결",
        f"학교 자료로 살펴보는 {locality} {LEVEL_NAME} 학습 우선순위",
        f"{locality} 학교 일정과 현재 교재를 함께 확인하기",
        f"{locality} {SCHOOL_NAME} 참고 자료를 상담에 활용하는 기준",
    ])
    if schools:
        names = "·".join(schools)
        names_subject = attach_particle(names, "이")
        names_topic = attach_particle(names, "은")
        paragraphs = [
            stable_pick(seed, "school-list", [
                f"제공된 {SCHOOL_NAME} 참고 목록은 {names}입니다. 이 목록은 수업 가능 여부를 보장하지 않으며, 상담에서 학생이 가져온 실제 학교 자료를 대조하기 위한 정보입니다.",
                f"{locality} {SCHOOL_NAME} 참고 정보에는 {names_subject} 포함됩니다. 학교명만으로 수업 가능 여부를 판단하지 않고 센터의 현재 개설 범위를 함께 확인해야 합니다.",
                f"상담 준비를 위해 제공된 {SCHOOL_NAME} 목록은 {names}입니다. 이는 수업 가능 학교를 단정하는 자료가 아니라 학생의 실제 범위와 과제를 확인하기 위한 참고 정보입니다.",
                f"센터 자료에서 확인한 {SCHOOL_NAME} 참고 목록은 {names}입니다. 실제 수업 여부는 학생 자료와 센터별 과목·학년 운영 범위를 확인한 뒤 판단합니다.",
                f"{names_topic} 제공 자료에 포함된 {locality} {SCHOOL_NAME} 참고 목록입니다. 목록 포함 여부와 실제 수업 가능 여부는 같지 않으므로 상담에서 별도로 확인합니다.",
                f"제공 자료에는 {names_subject} {SCHOOL_NAME} 참고 정보로 정리되어 있습니다. 학교 이름은 상담의 출발점으로만 사용하고 실제 시험·교과 자료를 함께 확인합니다.",
            ]),
            stable_pick(seed, "school-material", [
                f"{locality} {LEVEL_NAME}은 {material_label}를 준비해 국어 읽기·문법, 영어 어휘·문장 적용, 수학 단원·풀이 과정을 구분하는 것이 좋습니다. 확인한 내용을 다음 수업과 복습 계획에 어떻게 반영하는지 살펴보세요.",
                f"상담에서는 {material_label}를 과목별로 나누어 현재 범위와 반복 오답을 확인합니다. {locality} 학생에게 필요한 것은 학교명 반복보다 자료에서 찾은 문제를 다음 계획으로 바꾸는 과정입니다.",
                f"{locality} 학생이 사용하는 {material_label}를 가져오면 국어·영어·수학의 현재 단원과 준비 일정을 구체적으로 확인할 수 있습니다. 자료가 수업 진도와 재확인 날짜에 반영되는지도 물어보세요.",
                f"학교 자료는 이름 나열보다 활용 방식이 중요합니다. {locality} {LEVEL_NAME}의 {material_label}에서 과목별 범위와 어려운 문제를 표시해 상담 계획과 대조하는 편이 좋습니다.",
                f"{material_label}를 준비한 뒤 국어의 읽기 근거, 영어의 문장 적용과 수학의 풀이 과정을 따로 확인하세요. {locality} 상담에서는 이 기록이 과목별 우선순위로 이어지는지가 핵심입니다.",
                f"{locality} {LEVEL_NAME}의 현재 자료를 국어·영어·수학으로 분류하면 시험·교과 범위와 복습할 단원을 더 명확히 볼 수 있습니다. 학교 일정이 주간 과제에 반영되는지도 함께 확인합니다.",
            ]),
        ]
    else:
        paragraphs = [
            stable_pick(seed, "school-none", [
                f"{locality}에 제공된 {SCHOOL_NAME} 목록은 없습니다. 확인되지 않은 학교명을 추가하지 않고 상담에서 학생이 실제 사용하는 {material_label}를 기준으로 확인합니다.",
                f"제공 자료에는 {locality} {SCHOOL_NAME} 목록이 따로 정리되어 있지 않습니다. 학생의 실제 {material_label}를 준비해 현재 범위와 과제 일정을 확인하는 것이 정확합니다.",
                f"{locality}의 확인된 {SCHOOL_NAME} 참고 목록이 없어 학교명을 임의로 제시하지 않습니다. 상담 시 학생이 사용하는 {material_label}와 센터 운영 범위를 직접 대조하세요.",
                f"{locality}의 학교 정보가 제공되지 않아 특정 학교를 추정하지 않습니다. 학생의 실제 {material_label}를 바탕으로 상담 질문을 준비합니다.",
                f"{locality} {SCHOOL_NAME} 목록은 제공 자료에서 확인되지 않았습니다. 실제 학교 자료와 희망 센터의 과목·학년 운영 여부를 함께 확인해야 합니다.",
                f"확인되지 않은 학교명을 넣지 않기 위해 {locality} 학교 목록은 표시하지 않습니다. 학생이 가져온 {material_label}를 기준으로 상담합니다.",
            ]),
            stable_pick(seed, "school-none-material", [
                f"{locality} 학생의 자료를 국어 읽기·문법, 영어 어휘·문장 적용, 수학 단원·풀이 과정으로 나누면 현재 우선순위를 더 분명하게 정리할 수 있습니다.",
                f"학교 목록이 없어도 현재 교재와 과제, 평가 자료를 과목별로 준비하면 {locality} 학생의 보완 순서를 구체적으로 확인할 수 있습니다.",
                f"{material_label}에서 어려웠던 부분을 표시한 뒤 국어·영어·수학의 원인을 나누어 질문하면 상담 내용이 더 분명해집니다.",
                f"{locality} 상담 전에는 학생이 실제 사용하는 자료의 범위와 오답을 과목별로 정리해 센터의 가능 학년과 함께 확인하세요.",
                f"학교명 대신 현재 자료를 기준으로 읽기, 문장 적용과 풀이 과정의 막힌 지점을 나누면 현실적인 주간 계획을 세울 수 있습니다.",
                f"학생이 가져온 {material_label}를 과목별로 살펴보고 다음 수업에서 확인할 단원과 재풀이 시점을 정하는 방식이 정확합니다.",
            ]),
        ]
    return heading, diversify_school_paragraphs(paragraphs, seed)


def build_faq(title: str, locality: str, center: dict, student_type: str, seed: str) -> list[tuple[str, str]]:
    material_label = SCHOOL_MATERIALS if SCHOOL_MATERIALS.endswith("자료") else f"{SCHOOL_MATERIALS} 자료"
    q1 = stable_pick(seed, "q1", [
        f"{title} 상담은 어떤 학생에게 도움이 될 수 있나요?",
        f"{title}을 알아볼 때 학생 상태를 어떻게 확인하나요?",
        f"{title} 상담 전에 먼저 점검할 학습 문제는 무엇인가요?",
        f"{title}이 필요한지 어떤 기록으로 판단하나요?",
        f"{title} 선택 전 학생의 어떤 습관을 살펴야 하나요?",
        f"{title} 상담에서 첫 번째로 확인하는 내용은 무엇인가요?",
    ])
    a1_options = {
        "high": [
            f"최근 시험 뒤 과목별 보완 순서를 정하기 어렵다면 시험지와 오답 기록을 함께 준비하세요. {locality} 상담에서는 내신 일정과 실제 복습 시간을 대조해 우선순위를 확인할 수 있습니다.",
            f"한 과목의 과제 때문에 다른 과목 복습이 자주 밀린다면 세 과목의 완료 시간과 재풀이 결과를 나누어 볼 필요가 있습니다. 센터의 실제 개설 범위는 등록 전에 확인해야 합니다.",
            "내신 범위와 모의고사 오답을 한 계획 안에 배치하기 어렵다면 국어 근거 찾기, 영어 해석과 수학 풀이 단계를 따로 기록해 상담 자료로 준비하는 편이 좋습니다.",
            f"시험 일정은 알고 있지만 실행 가능한 주간 분량을 정하기 어렵다면 최근 교재와 학교 자료를 가져가 보세요. {locality} 상담에서 집중 과목과 최소 복습량을 구분할 수 있습니다.",
            "최근 성적표만으로 보완 방향을 정하기 어려운 고등학생은 틀린 이유, 과제 완료 시각과 일정 시간이 지난 뒤의 재풀이 결과를 함께 살펴보는 점검이 필요합니다.",
            f"국어·영어·수학의 공부 시간은 많지만 완료 기준이 불분명하다면 과목별 기록과 학교 일정을 대조하세요. {locality} 센터의 수업 가능 학년도 별도로 확인해야 합니다.",
        ],
        "middle": [
            f"과제를 끝낸 뒤 오답 복습이 이어지지 않는다면 최근 시험지와 교재에서 막힌 지점을 과목별로 나누어 보세요. {locality} 상담에서는 학교 일정과 주간 실행 시간도 함께 확인합니다.",
            "배운 개념을 학교 시험 문제에 적용하기 어렵다면 국어 근거 설명, 영어 문장 해석과 수학 풀이 과정을 각각 확인하는 진단이 도움이 될 수 있습니다.",
            f"세 과목의 시험 준비가 겹쳐 우선순위를 정하기 어렵다면 완료하지 못한 과제와 반복 오답을 준비하세요. {locality} 센터의 개설 과목과 가능 학년은 등록 전에 확인해야 합니다.",
            "수업에서는 이해하지만 혼자 다시 설명하기 어렵다면 정답 수보다 풀이 과정, 질문을 미룬 지점과 재확인 날짜를 살펴보는 편이 좋습니다.",
            f"학교 일정과 학원 과제를 한 주 안에 배치하기 어렵다면 최근 시간표와 과제 기록을 대조해 보세요. {locality} 상담에서는 과목별 최소 복습량도 확인할 수 있습니다.",
            "국어·영어·수학마다 막히는 원인이 다른데 같은 방식으로 공부한다면 최근 자료를 과목별로 분류하고 보완 행동과 확인 날짜를 따로 정해야 합니다.",
        ],
        "elementary": [
            f"읽기·어휘·계산의 기초가 고르지 않다면 현재 교재와 단원평가 자료를 준비하세요. {locality} 상담에서는 아이가 스스로 설명하고 다시 푸는 과정도 함께 확인합니다.",
            "숙제를 시작하는 시각과 끝내는 기준을 스스로 정하기 어렵다면 과제량을 늘리기 전에 짧은 시작·완료 기록과 복습 주기를 살펴보는 편이 좋습니다.",
            f"배운 내용을 말로 설명하거나 틀린 문제를 다시 풀기 어렵다면 최근 풀이 흔적과 알림장을 함께 확인하세요. {locality} 센터의 가능 학년은 등록 전에 다시 확인해야 합니다.",
            "교과 진도는 따라가지만 세 과목의 기초 과정이 고르지 않다면 국어 읽기, 영어 어휘 활용과 수학 풀이 순서를 나누어 보는 점검이 필요합니다.",
            f"그날 공부할 순서를 정하는 연습이 필요하다면 현재 교재와 과제 기록을 준비하세요. {locality} 상담에서는 아이가 혼자 시작할 수 있는 분량부터 확인합니다.",
            "정답을 맞히는 것보다 풀이 과정과 읽은 근거를 남기는 연습이 필요하다면 점수만 보지 말고 설명 과정과 다시 풀 날짜를 함께 살펴보세요.",
        ],
    }
    a1 = stable_pick(seed, "a1", a1_options[PROFILE])
    q2 = stable_pick(seed, "q2", [
        f"{locality} {LEVEL_NAME}이 국어·영어·수학을 함께 관리할 때 무엇을 구분해야 하나요?",
        f"{locality}에서 세 과목 계획을 한꺼번에 세워도 괜찮을까요?",
        f"국어·영어·수학 과제량은 {locality} {LEVEL_NAME}에게 어떻게 배분하나요?",
        f"{locality} {LEVEL_NAME}의 국영수 오답은 같은 방식으로 관리하나요?",
        f"세 과목을 함께 배우더라도 {locality} 상담에서 따로 확인할 부분은 무엇인가요?",
        f"{locality} 국영수 학습에서 과목별 병목을 어떻게 나누나요?",
    ])
    a2 = f"{locality} 상담에서는 " + stable_pick(seed, "a2", [
        "국어는 지문과 선택지의 근거, 영어는 문장 구조와 해석 순서, 수학은 개념을 식으로 옮기는 과정을 따로 확인해야 합니다. 이후 시험 일정과 복습 시간을 보고 세 과목의 주간 분량을 조정합니다.",
        "한 과목의 과제가 다른 과목 복습을 밀어내지 않도록 최소 유지량과 집중 과목을 구분합니다. 국어·영어·수학의 오답 원인이 다르므로 재학습 방법과 확인 날짜도 과목별로 정해야 합니다.",
        "세 과목을 같은 분량으로 나누기보다 최근 시험과 학교 일정, 반복 오답을 기준으로 우선순위를 정합니다. 과목별 완료 기준이 있어야 계획과 실제 실행의 차이를 확인할 수 있습니다.",
        "국어는 답의 근거 설명, 영어는 어휘·구문·본문 적용, 수학은 개념·유형·계산 과정을 구분해 기록합니다. 주간 계획에서는 각 기록을 보고 다음 과제량을 조정해야 합니다.",
        "국영수를 함께 관리한다는 표현이 세 과목을 동일하게 다룬다는 뜻은 아닙니다. 학생이 막힌 장면과 시험 일정을 과목별로 확인한 뒤 공통 시간표 안에서 충돌하지 않게 배치합니다.",
        "먼저 최근 자료에서 가장 시급한 과목을 찾고, 나머지 과목의 최소 복습 시간을 남겨야 합니다. 오답은 설명 직후가 아니라 일정 시간이 지난 뒤 다시 풀어 적용 여부를 확인합니다.",
    ])
    schools = center["schools"]
    if schools:
        shown = "·".join(schools[:4])
        q3 = stable_pick(seed, "q3-school", [
            f"{locality} {SCHOOL_NAME} 참고 정보는 상담에서 어떻게 활용하나요?",
            f"{locality} 학교별 교과 자료는 어떤 방식으로 확인하나요?",
            f"제공된 {locality} {SCHOOL_NAME} 목록은 수업 가능 학교를 뜻하나요?",
            f"{locality} {LEVEL_NAME}은 상담 때 어떤 학교 자료를 준비해야 하나요?",
        ])
        a3 = stable_pick(seed, "a3-school", [
            f"제공된 {SCHOOL_NAME} 참고 목록에는 {shown} 등이 포함됩니다. 이는 수업 가능 여부를 보장하는 목록이 아니며, 실제 {material_label}와 센터별 개설 과목을 상담에서 함께 확인해야 합니다.",
            f"{shown} 등은 제공 자료의 {SCHOOL_NAME} 참고 목록입니다. 학생이 가져온 {material_label}를 대조하고 희망 센터의 과목·학년 운영 여부를 별도로 확인하세요.",
            f"학교명은 상담 준비를 위한 참고 정보입니다. {shown} 등과 관련한 실제 수업 여부는 학생의 {material_label}, 센터의 현재 개설 범위와 함께 확인해야 합니다.",
            f"제공 목록에서 {shown} 등을 확인할 수 있지만 모든 학교의 수업 가능 여부를 뜻하지는 않습니다. 실제 범위와 과제 자료를 준비해 상담에서 확인하는 편이 정확합니다.",
            f"{locality}의 {SCHOOL_NAME} 참고 정보에는 {shown} 등이 표시됩니다. 학교 목록보다 학생의 현재 {material_label}와 센터 운영 범위를 대조하는 과정이 우선입니다.",
            f"{shown} 등은 확인된 참고 학교명이며 수업 가능 학교를 단정하지 않습니다. 학생의 실제 자료와 센터별 개설 과목을 함께 놓고 상담하세요.",
        ])
    else:
        q3 = stable_pick(seed, "q3-none", [
            f"{locality} {SCHOOL_NAME} 목록이 없는 경우 상담은 어떻게 준비하나요?",
            f"{locality}에 학교명이 따로 표시되지 않으면 상담은 어떻게 준비하나요?",
            f"제공된 {locality} {SCHOOL_NAME} 정보가 없을 때 무엇을 확인해야 하나요?",
            f"{locality} 학생의 학교 자료는 상담 때 직접 가져가야 하나요?",
        ])
        a3 = stable_pick(seed, "a3-none", [
            f"{locality}에 제공된 {SCHOOL_NAME} 목록이 없어 임의로 학교명을 추가하지 않았습니다. 학생이 실제 사용하는 {material_label}를 준비하고 희망 센터의 과목·학년 운영 범위를 직접 확인하는 것이 정확합니다.",
            f"제공 자료에서 {locality} {SCHOOL_NAME} 목록을 확인할 수 없어 특정 학교를 추정하지 않습니다. 현재 {material_label}와 센터의 가능 학년을 상담에서 함께 확인하세요.",
            f"학교명이 표시되지 않아도 상담은 가능합니다. 학생이 사용하는 {material_label}를 가져와 현재 범위와 오답을 설명하고 센터의 실제 개설 과목을 확인하면 됩니다.",
            f"{locality}의 확인된 학교 목록이 없으므로 임의 정보를 넣지 않았습니다. 실제 교재·학교 자료와 희망 센터의 과목 운영 범위를 직접 대조하는 방식이 정확합니다.",
            f"제공 자료에는 {locality} {SCHOOL_NAME} 정보가 따로 없습니다. 학생의 {material_label}를 준비해 과목별 현재 진도와 센터 수업 가능 여부를 상담에서 확인하세요.",
            f"확인되지 않은 학교명을 추가하는 대신 학생이 실제 사용하는 {material_label}를 기준으로 안내합니다. 센터별 과목과 학년 운영은 등록 전에 다시 확인해야 합니다.",
        ])
    if PROFILE == "elementary":
        a2 += " 초등 교과에서는 읽기·어휘·계산 기초와 스스로 공부를 시작하는 습관을 함께 살핍니다."
        a3 += " 상담에는 알림장과 단원평가처럼 학생이 실제 사용하는 교과 자료를 준비하는 편이 좋습니다."
    q4 = stable_pick(seed, "q4", [
        f"{locality} 상담 전에 센터 위치와 교습비는 어디에서 확인하나요?",
        f"{locality} 센터 방문 전에 확인할 운영 정보는 무엇인가요?",
        f"{locality} {LEVEL_NAME} 상담을 예약할 때 어떤 정보를 준비하면 좋나요?",
        f"{locality} 학원 상담에서 수업 가능 학년과 비용을 어떻게 확인하나요?",
        f"{locality} 센터의 주소와 과목 운영은 모두 같은가요?",
        f"{locality} 센터에 문의할 때 주소와 가능 학년을 함께 물어봐야 하나요?",
        f"{locality} 수업 시간표와 교습비 자료는 상담 전에 볼 수 있나요?",
        f"{locality}에서 센터별 과목 운영 범위를 확인하는 방법은 무엇인가요?",
        f"{locality} 방문 상담 전 통학 조건과 비용은 어떻게 비교하나요?",
        f"{locality} 센터 등록 전에 다시 확인해야 할 항목은 무엇인가요?",
        f"{locality} 상담 시 개설 과목과 보강 방식은 어디에 문의하나요?",
        f"{locality} 학부모가 센터 운영 정보를 확인할 순서는 무엇인가요?",
    ])
    location_subject = "실제 방문 위치는 위 센터 확인 정보에서 확인할 수 있습니다. " if center.get("address") else ""
    location_object = "실제 방문 위치를 위 센터 확인 정보에서 확인할 수 있습니다. " if center.get("address") else ""
    tuition = "센터별 교습비 확인 버튼에서 제공 자료를 볼 수 있습니다. " if center["tuition_url"] else "교습비 자료는 상담 시 직접 확인해야 합니다. "
    a4 = stable_pick(seed, "a4", [
        f"{locality}에서 문의할 때는 {location_object}{tuition}개설 과목과 {LEVEL_NAME} 가능 학년, 시간표와 보강 방식은 센터별로 다를 수 있으므로 등록 전에 함께 확인하세요.",
        f"{location_subject}{tuition}{locality} 학생의 가능 학년과 개설 과목, 수업 시각과 보강 기준은 희망 센터에 직접 문의해야 합니다.",
        f"먼저 {location_object}{tuition}이후 {locality} 통학 시간과 {LEVEL_NAME} 과목 운영 범위를 센터 상담에서 대조하세요.",
        f"{locality} 센터 자료에서 주소와 교습비를 확인한 뒤, {LEVEL_NAME} 가능 과목·시간표·보강 여부를 등록 전에 다시 질문하세요. {location_subject}",
        f"{location_subject}{locality} 방문 전에는 {tuition}실제 개설 과목, 가능 학년과 수업 시간을 한 번 더 확인하는 편이 정확합니다.",
        f"센터마다 운영 조건이 다를 수 있습니다. {location_subject}{tuition}{locality} 상담에서 과목별 가능 학년과 보강 방식을 함께 확인하세요.",
        f"{tuition}{location_subject}{locality} 학생에게 맞는 시간표가 있는지와 {LEVEL_NAME} 수업 범위는 희망 센터의 현재 안내를 기준으로 판단합니다.",
        f"{locality} 상담 전에는 {location_object}{tuition}마지막으로 개설 과목, 가능 학년과 통학 가능한 수업 시각을 기록해 비교하세요.",
    ])
    return [(q1, a1), (q2, a2), (q3, a3), (q4, a4)]


def build_scenarios(locality: str, center: dict, student_type: str, seed: str) -> list[str]:
    material_label = SCHOOL_MATERIALS if SCHOOL_MATERIALS.endswith("자료") else f"{SCHOOL_MATERIALS} 자료"
    first = stable_pick(seed, "scenario-1", [
        f"{locality} {LEVEL_NAME} 자녀의 최근 시험지와 오답 노트를 학부모가 함께 준비한 상황입니다. 상담에서는 세 과목을 모두 늘리기보다 먼저 바꿀 과목과 복습 시점을 정합니다.",
        f"{locality} {LEVEL_NAME}의 상담을 가정한 예시입니다. 국어·영어·수학 점수만 비교하지 않고 시작이 늦어진 과제, 반복한 오답과 학교 일정을 나누어 본 뒤 이번 주에 실행할 한두 가지를 정리합니다.",
        f"{locality} 학부모가 자녀의 세 과목 학습량이 서로 충돌하는 문제를 질문한 상황을 예로 들었습니다. 상담 후에는 성적 약속보다 과목별 최소 복습량, 집중 단원과 재확인 날짜가 구체적인지 살펴봅니다.",
        f"최근 학교 학습 준비가 한 과목에 치우친 {locality} {LEVEL_NAME}을 가정했습니다. 학부모는 현재 교재와 학교 자료를 바탕으로 국어·영어·수학의 병목을 따로 듣고, 가정에서 확인할 기록을 정리합니다.",
        f"{locality} {LEVEL_NAME}의 하루 시간표를 바탕으로 만든 상담 상황 예시입니다. 과목별 과제와 오답 재풀이를 실제로 배치해 보고 무리한 분량은 줄이는 방향으로 질문을 정리합니다.",
        f"{locality}에서 세 과목을 함께 알아보는 학부모의 상담 장면을 가정했습니다. 최근 풀이를 보며 국어 근거 찾기, 영어 해석, 수학 풀이 단계 중 우선 점검할 부분을 구분해 듣는 상황입니다.",
    ])
    if center["schools"]:
        school_note = f"제공된 {SCHOOL_NAME} 참고 정보인 {'·'.join(center['schools'][:3])}{' 등' if len(center['schools']) > 3 else ''}의 목록과 학생의 실제 학교 자료를 대조해 질문하는"
    else:
        school_note = f"제공된 {SCHOOL_NAME} 목록이 없어 학생의 실제 {material_label}를 직접 준비해 질문하는"
    second = f"{locality}에서 " + stable_pick(seed, "scenario-2", [
        f"{school_note} 상황입니다. 학부모는 학교명 자체보다 시험 범위와 수행평가 일정이 수업 계획에 반영되는지, 다음 확인 시점이 언제인지 기록합니다.",
        f"{school_note} 상담 예시입니다. 센터의 개설 과목과 가능 학년을 확인한 뒤 학생의 귀가 시간, 과제량과 시험 기간 보강 기준을 같은 메모에 정리합니다.",
        f"{school_note} 학부모 상황을 가정했습니다. 상담 내용을 들은 뒤 국어·영어·수학의 피드백 방식이 각각 무엇인지와 가정에서 확인할 기록을 구분합니다.",
        f"{school_note} 경우입니다. 학부모는 교습비와 시간표뿐 아니라 과목별 진단 근거, 보완 순서와 일정 시간이 지난 뒤 재확인하는 방법까지 질문합니다.",
        f"{school_note} 장면을 예로 들었습니다. 실제 등록 판단 전에는 주소와 통학 시간, 센터별 운영 범위, 학교 시험 자료의 반영 방식을 차례로 확인합니다.",
        f"{school_note} 상담 상황입니다. 제공 자료와 학생이 가져온 자료를 구분하고, 확인되지 않은 학교 정보나 수업 조건은 임의로 단정하지 않습니다.",
    ])
    return [first, second]


def grade_summary(center: dict) -> list[tuple[str, str]]:
    return [
        ("국어", "·".join(center["korean_grades"]) or f"{COURSE_NAME} 과정 미기재"),
        ("영어", "·".join(center["english_grades"]) or f"{COURSE_NAME} 과정 미기재"),
        ("수학", "·".join(center["math_grades"]) or f"{COURSE_NAME} 과정 미기재"),
    ]


def offer_nodes(title: str, center: dict) -> list[dict]:
    if not center["tuition_url"]:
        return []
    return [{
        "@type": "Offer",
        "name": f"{center['center_name']} 교습비 확인",
        "url": center["tuition_url"],
        "itemOffered": {
            "@type": "Service",
            "name": f"{center['center_name']} 교습과정·교습비 자료",
            "serviceType": "센터별 교습과정 정보 확인",
        },
    }]


def build_graph(title: str, locality: str, slug: str, center: dict, meta: str, faq: list[tuple[str, str]], headings: list[str], rep_name: str) -> dict:
    url = canonical_url(slug)
    hub_url = canonical_url()
    level_subjects = [f"{COURSE_NAME} 국어", f"{COURSE_NAME} 영어", f"{COURSE_NAME} 수학"]
    subjects = [*level_subjects, "학교 학습 대비", "오답 재학습"]
    available_subjects = [
        label for label, grades in (
            (f"{COURSE_NAME} 국어", center["korean_grades"]),
            (f"{COURSE_NAME} 영어", center["english_grades"]),
            (f"{COURSE_NAME} 수학", center["math_grades"]),
        ) if grades
    ]
    level_grade_union = list(dict.fromkeys(center["korean_grades"] + center["english_grades"] + center["math_grades"]))
    schools = center["schools"]
    offers = offer_nodes(title, center)
    org = {
        "@type": "EducationalOrganization", "@id": url + "#organization", "name": center["center_name"],
        "url": url, "telephone": PHONE, "address": {"@type": "PostalAddress", "streetAddress": center["address"], "addressCountry": "KR"},
        "areaServed": [locality], "description": meta,
    }
    if available_subjects:
        org["teaches"] = [*available_subjects, "학습코칭"]
    if center["registration"]:
        org["identifier"] = {"@type": "PropertyValue", "name": center["office_name"] or "교육지원청 등록정보", "value": center["registration"]}
    if level_grade_union:
        org["educationalLevel"] = level_grade_union
    if offers:
        org["makesOffer"] = offers
    local_business = {
        "@type": "LocalBusiness", "@id": url + "#localbusiness", "name": center["center_name"], "url": url,
        "telephone": PHONE, "image": f"/assets/representative/{rep_name}", "address": org["address"],
        "areaServed": [locality], "parentOrganization": {"@id": org["@id"]},
    }
    if offers:
        local_business["makesOffer"] = offers
    breadcrumb = {
        "@type": "BreadcrumbList", "@id": url + "#breadcrumb", "itemListElement": [
            {"@type": "ListItem", "position": 1, "name": "홈", "item": SITE_ORIGIN + "/"},
            {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": absolute_site_url("/과목별학원/")},
            {"@type": "ListItem", "position": 3, "name": CATEGORY_NAME, "item": hub_url},
            {"@type": "ListItem", "position": 4, "name": title, "item": url},
        ],
    }
    about = [{"@type": "Thing", "name": value} for value in [CATEGORY_NAME, *level_subjects, "학교 학습 대비", "학습코칭"]]
    mentions = [{"@type": "Thing", "name": value} for value in dict.fromkeys([center["region"], center["district"], locality, *schools, *subjects]) if value]
    article = {
        "@type": "Article", "@id": url + "#article", "mainEntityOfPage": {"@id": url + "#webpage"},
        "headline": title, "description": meta, "abstract": meta, "inLanguage": "ko-KR",
        "articleSection": [CATEGORY_NAME, center["region"], center["district"], locality],
        "about": about, "mentions": mentions, "hasPart": [{"@type": "WebPageElement", "name": heading} for heading in headings],
        "author": {"@id": org["@id"]}, "publisher": {"@id": org["@id"]},
        "datePublished": DATE_PUBLISHED, "dateModified": DATE_MODIFIED,
        "image": [absolute_site_url(f"/assets/representative/{rep_name}"), absolute_site_url(f"/assets/centers/common/{center['body_image']}"), absolute_site_url(f"/assets/maps/{center['map_name']}")],
    }
    if level_grade_union:
        article["educationalLevel"] = level_grade_union
    service = {
        "@type": "Service", "@id": url + "#service", "name": f"{title} 학습상담 안내", "serviceType": CATEGORY_NAME,
        "provider": {"@id": org["@id"]}, "areaServed": [locality], "description": meta,
        "audience": {"@type": "EducationalAudience", "educationalRole": "student", "audienceType": LEVEL_NAME},
    }
    if offers:
        service["makesOffer"] = offers
    webpage = {
        "@type": "WebPage", "@id": url + "#webpage", "url": url, "name": title, "description": meta,
        "inLanguage": "ko-KR", "isPartOf": {"@id": hub_url + "#collection"}, "breadcrumb": {"@id": breadcrumb["@id"]},
        "primaryImageOfPage": {"@type": "ImageObject", "url": absolute_site_url(f"/assets/representative/{rep_name}")}, "about": about,
        "mainEntity": {"@id": article["@id"]},
    }
    faq_node = {"@type": "FAQPage", "@id": url + "#faq", "mainEntity": [{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in faq]}
    links = [
        (CATEGORY_NAME + " 전체 지역", hub_url), ("과목별학원", absolute_site_url("/과목별학원/")),
        ("학습가이드", absolute_site_url("/학습가이드/")), ("상담문의", absolute_site_url("/상담문의/")),
    ]
    item_list = {"@type": "ItemList", "@id": url + "#related", "name": f"{title} 관련 페이지", "itemListElement": [{"@type": "ListItem", "position": i, "name": name, "url": href} for i, (name, href) in enumerate(links, 1)]}
    return {"@context": "https://schema.org", "@graph": [org, local_business, webpage, breadcrumb, article, service, faq_node, item_list]}


def node_has_type(node: dict, expected: str) -> bool:
    value = node.get("@type", "")
    return expected in value if isinstance(value, list) else value == expected


def preserve_published_graph(
    current_html: str,
    fallback: dict,
    title: str,
    meta: str,
    faq: list[tuple[str, str]],
    headings: list[str],
) -> dict:
    """Keep consolidated center entities while refreshing page-owned copy."""
    match = re.search(r'<script\s+type="application/ld\+json">(.*?)</script>', current_html, re.DOTALL)
    if not match:
        return fallback
    try:
        graph = json.loads(match.group(1))
    except json.JSONDecodeError:
        return fallback
    if not isinstance(graph, dict) or not isinstance(graph.get("@graph"), list):
        return fallback
    graph = json.loads(json.dumps(graph, ensure_ascii=False))
    fallback_nodes = {
        kind: next((item for item in fallback.get("@graph", []) if isinstance(item, dict) and node_has_type(item, kind)), {})
        for kind in ("WebPage", "Article", "Service", "FAQPage", "BreadcrumbList", "ItemList")
    }
    for node in graph["@graph"]:
        if not isinstance(node, dict):
            continue
        if node_has_type(node, "WebPage"):
            source = fallback_nodes["WebPage"]
            node.update({key: source[key] for key in ("url", "name", "description", "primaryImageOfPage", "about", "mainEntity") if key in source})
        elif node_has_type(node, "Article"):
            source = fallback_nodes["Article"]
            node.update({
                "headline": title,
                "description": meta,
                "abstract": meta,
                "datePublished": DATE_PUBLISHED,
                "dateModified": DATE_MODIFIED,
                "hasPart": [{"@type": "WebPageElement", "name": heading} for heading in headings],
                "mentions": source.get("mentions", []),
                "about": source.get("about", []),
                "image": source.get("image", []),
            })
        elif node_has_type(node, "Service"):
            node["description"] = meta
        elif node_has_type(node, "FAQPage"):
            node["mainEntity"] = [
                {
                    "@type": "Question",
                    "name": question,
                    "acceptedAnswer": {"@type": "Answer", "text": answer},
                }
                for question, answer in faq
            ]
        elif node_has_type(node, "BreadcrumbList") and fallback_nodes["BreadcrumbList"]:
            node.update(fallback_nodes["BreadcrumbList"])
        elif node_has_type(node, "ItemList") and fallback_nodes["ItemList"]:
            node.update(fallback_nodes["ItemList"])
    return graph


def header(prefix: str, current: str = "subjects") -> str:
    items = [("home", "홈", "/"), ("about", "학원소개", "/학원소개/"), ("guide", "학습가이드", "/학습가이드/"), ("contact", "상담문의", "/상담문의/"), ("subjects", "과목별학원", "/과목별학원/")]
    links = "".join(f'<a href="{href}" data-nav="{key}"{" aria-current=\"page\"" if key == current else ""}>{label}</a>' for key, label, href in items)
    return f'''<a class="skip-link" href="#main">본문으로 건너뛰기</a>
  <header class="site-header"><div class="site-shell header-inner">
    <a class="brand" href="/"><span class="brand-mark" aria-hidden="true">W</span><span class="brand-copy"><strong>{SITE_NAME}</strong><small>STUDY RECORD COACHING</small></span></a>
    <nav class="primary-nav" aria-label="주요 메뉴">{links}</nav>
  </div></header>'''


def footer(prefix: str) -> str:
    return f'''<footer class="site-footer"><div class="site-shell footer-grid">
    <div class="footer-brand"><h2>{SITE_NAME}</h2><p>학생별 진도와 교재, 실행 기록과 오답 재학습을 연결해 다음 공부 순서를 정리합니다. 개설 과목과 학년, 수업 방식은 센터별로 확인합니다.</p></div>
    <div><nav class="footer-links" aria-label="하단 메뉴"><a href="/학원소개/">학원소개</a><a href="/학습가이드/">학습가이드</a><a href="/과목별학원/">과목별학원</a><a href="/상담문의/">상담문의</a></nav><p class="footer-meta">대표 상담 {PHONE}<br>© {SITE_NAME}</p></div>
  </div></footer>
  <nav class="contact-dock" aria-label="빠른 상담"><a href="tel:010-6839-8283">전화문의</a><a href="https://blogsms.net/01068398283" target="_blank" rel="noopener">문자문의</a><a href="https://docs.google.com/forms/d/e/1FAIpQLSdb2oE5Qk5YS0TfYDxyV1w-IOTkhkjOCmmpAKTI9FmqpVj6Yg/viewform" target="_blank" rel="noopener">상담신청</a></nav>'''


def render_info(center: dict) -> str:
    rows = [
        ("지역", " ".join(part for part in (center["display_region"], center["display_district"], center["display_locality"]) if part)),
        ("센터 기준", center["center_name"]),
        ("제공 주소", center["address"]),
    ]
    html_rows = "".join(f"<div><dt>{esc(label)}</dt><dd>{esc(value)}</dd></div>" for label, value in rows if value)
    grades = "".join(f'<li><strong>{esc(subject)}</strong><span>{esc(value)}</span></li>' for subject, value in grade_summary(center))
    schools = "".join(f"<span>{esc(school)}</span>" for school in center["schools"])
    school_html = f'<div><dt>{SCHOOL_NAME} 참고</dt><dd><div class="local-tags">{schools}</div></dd></div>' if schools else f'<div><dt>{SCHOOL_NAME} 참고</dt><dd>제공 목록 없음 · 상담 시 실제 학교 자료 확인</dd></div>'
    registration_html = f'<div><dt>등록 정보</dt><dd>{esc(center["registration"])}</dd></div>' if center["registration"] else ""
    phone_html = f'<div><dt>대표 상담</dt><dd><a href="tel:{PHONE}">{PHONE}</a></dd></div>'
    grade_text = ", ".join(f"{subject} {value}" for subject, value in grade_summary(center))
    fee_note = "교습비 자료는 아래 확인 버튼으로 연결됩니다." if center["tuition_url"] else "교습비 자료는 상담 시 확인합니다."
    verified_note = (
        '<p class="center-verified-note"><strong>제공 자료로 확인한 범위</strong>'
        f'<span>제공 자료에서 {esc(center["center_name"])}의 확인 가능한 범위는 '
        f'{esc(grade_text)}입니다. {fee_note} 시간표·보강·차량·주차는 센터에서 확인합니다.</span></p>'
    )
    tuition = f'<a class="button compact" href="{esc(center["tuition_url"])}" target="_blank" rel="noopener">센터별 교습비 확인 <span aria-hidden="true">↗</span></a>' if center["tuition_url"] else '<p class="info-note">교습비 자료는 희망 센터에서 확인합니다.</p>'
    return f'<dl class="local-facts">{html_rows}{school_html}{registration_html}{phone_html}</dl><ul class="grade-list">{grades}</ul>{verified_note}{tuition}'


def render_page(record: dict, previous_record: dict, next_record: dict) -> str:
    sections = record["sections"]
    title = record["title"]
    locality = record["locality"]
    slug = record["slug"]
    center = record["center"]
    meta = record["meta"]
    rep_name = record["rep_name"]
    student_type = verified_student_type(extract_student_type(sections["본문"], locality), center, slug)
    intro, body_sections = personalize_body(sections["본문"], title, locality, center, student_type, slug)
    replacement = school_section(center, locality, slug)
    normalized_sections: list[tuple[str, list[str]]] = []
    school_section_added = False
    for heading, paragraphs in body_sections:
        if "학교" in heading:
            if not school_section_added:
                normalized_sections.append(replacement)
                school_section_added = True
            continue
        normalized_sections.append((heading, paragraphs))
    if not school_section_added:
        insert_at = min(2, len(normalized_sections))
        normalized_sections.insert(insert_at, replacement)
    body_sections = normalized_sections
    faq = [(clean(question), clean(answer)) for question, answer in build_faq(title, locality, center, student_type, slug)]
    scenarios = build_scenarios(locality, center, student_type, slug)
    quick_answer = direct_answer(locality, center, student_type, slug)
    answer_card_html = "".join(
        f'<article><span>{esc(label)}</span><strong>{esc(answer)}</strong></article>'
        for label, answer in answer_cards(locality, center, slug)
    )
    current_path = CATEGORY_ROOT / slug / "index.html"
    current_html = current_path.read_text(encoding="utf-8") if current_path.is_file() else ""
    graph = preserve_published_graph(
        current_html,
        build_graph(title, locality, slug, center, meta, faq, [heading for heading, _ in body_sections], rep_name),
        title,
        meta,
        faq,
        [heading for heading, _ in body_sections],
    )
    graph_json = json.dumps(graph, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    body_html = "".join(f'<section class="manuscript-section"><span class="section-kicker">{index:02d}</span><h2>{esc(heading)}</h2>{"".join(f"<p>{esc(paragraph)}</p>" for paragraph in paragraphs)}</section>' for index, (heading, paragraphs) in enumerate(body_sections, 1))
    faq_html = "".join(f'<details{" open" if index == 0 else ""}><summary>{esc(question)}</summary><p>{esc(answer)}</p></details>' for index, (question, answer) in enumerate(faq))
    scenario_html = "".join(f'<article class="scenario-card"><span>상담 상황 예시 {index:02d}</span><p>{esc(value)}</p></article>' for index, value in enumerate(scenarios, 1))
    body_image = center["body_image"]
    region_label = " ".join(part for part in (center["display_region"], center["display_district"], center["display_locality"]) if part)
    prev_link = f'<a class="local-nav-card" href="{esc(page_path(previous_record["slug"]))}"><small>이전 지역</small><strong>{esc(previous_record["title"])}</strong><span>←</span></a>'
    next_link = f'<a class="local-nav-card" href="{esc(page_path(next_record["slug"]))}"><small>다음 지역</small><strong>{esc(next_record["title"])}</strong><span>→</span></a>'
    page_url = canonical_url(slug)
    rep_url = absolute_site_url(f"/assets/representative/{rep_name}")
    rep_width, rep_height = image_dimensions(REP_TARGET / rep_name)
    rep_mime = image_mime_type(rep_name)
    map_overlays = "".join(re.findall(r'<span class="map-contact-correction".*?</span>', current_html, re.DOTALL))
    map_id = ' id="center-map"' if map_overlays else ""
    map_width, map_height = center["map_width"], center["map_height"]
    return f'''<!doctype html>
<html lang="ko"><head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)} | {TITLE_SUFFIX}</title><meta name="description" content="{esc(meta)}">
  <meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec">
  <link rel="canonical" href="{page_url}"><meta property="og:url" content="{page_url}">
  <meta property="og:type" content="article"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="{esc(title)} | {TITLE_SUFFIX}"><meta property="og:description" content="{esc(meta)}"><meta property="og:image" content="{rep_url}"><meta property="og:image:secure_url" content="{rep_url}"><meta property="og:image:type" content="{rep_mime}"><meta property="og:image:width" content="{rep_width}"><meta property="og:image:height" content="{rep_height}"><meta property="og:image:alt" content="{esc(title)} 대표 이미지">
  <meta name="twitter:card" content="summary"><meta name="twitter:title" content="{esc(title)} | {TITLE_SUFFIX}"><meta name="twitter:description" content="{esc(meta)}"><meta name="twitter:image" content="{rep_url}"><meta name="twitter:image:alt" content="{esc(title)} 대표 이미지">
  <link rel="icon" href="../../../assets/favicon.png"><link rel="stylesheet" href="../../../assets/site14.css">
  <script type="application/ld+json">{graph_json}</script>
</head><body data-page="subjects">
  {header("../../../")}
  <main id="main">
    <section class="local-hero"><div class="site-shell">
      <nav class="breadcrumbs" aria-label="현재 위치"><a href="/">홈</a><a href="/과목별학원/">과목별학원</a><a href="{esc(page_path())}">{CATEGORY_NAME}</a><span>{esc(title)}</span></nav>
      <p class="eyebrow">{ENGLISH_LEVEL} Korean · English · Math</p><h1>{esc(title)}</h1><p class="local-lead">{esc(meta)}</p>
      <div class="local-answer-grid">{answer_card_html}</div>
    </div></section>
    <section class="section local-overview"><div class="site-shell local-overview-grid">
      <div class="local-summary"><p class="chapter-label"><span>01</span> Quick answer</p><h2>{esc(locality)}에서 먼저 확인할 내용</h2><p>{esc(quick_answer)}</p><div class="answer-note"><strong>대표 학생 상황</strong><p>{esc(student_type)}</p></div></div>
      <aside class="local-info-card"><p class="eyebrow">Center information</p><h2>수업·상담 확인 정보</h2>{render_info(center)}</aside>
    </div></section>
    <section class="local-media-section"><div class="site-shell local-media-stack">
      <figure class="local-body-image"><img src="../../../assets/centers/common/{body_image}" width="918" height="16116" alt="{esc(title)} 본문 {SITE_NAME}" loading="lazy" decoding="async"><figcaption>{esc(region_label)} {LEVEL_NAME}의 국어·영어·수학 학습 점검 안내</figcaption></figure>
      <figure class="local-map-image"{map_id}><div class="map-art"><img src="../../../assets/maps/{esc(center['map_name'])}" width="{map_width}" height="{map_height}" alt="{esc(title)} 지도 {SITE_NAME}" loading="lazy" decoding="async">{map_overlays}</div><figcaption>센터 위치는 제공된 주소 자료를 기준으로 표시하며 방문 전 실제 운영 여부를 확인합니다.</figcaption></figure>
    </div></section>
    <section class="section manuscript-wrap"><article class="site-shell manuscript-article"><div class="manuscript-intro"><span>상담 전 핵심 안내</span><p>{esc(intro)}</p></div>{body_html}</article></section>
    <section class="section blue-wash"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>02</span> Consultation examples</div><div><h2>{esc(locality)} 학부모 상담 상황 예시</h2><p>아래 내용은 실제 고객 후기나 성적 결과가 아니라, 상담에서 확인할 질문을 이해하기 위한 상황 예시입니다.</p></div></div><div class="scenario-grid">{scenario_html}</div></div></section>
    <section class="section"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>03</span> FAQ</div><div><h2>{esc(title)} 자주 묻는 질문</h2><p>학부모님이 상담 전에 자주 확인하는 내용을 질문과 답변으로 정리했습니다.</p></div></div><div class="faq-list">{faq_html}</div></div></section>
    <section class="section local-links-section"><div class="site-shell"><div class="section-heading compact-heading"><div class="chapter-label"><span>04</span> Continue</div><div><h2>{esc(locality)} 페이지 이동</h2><p>카테고리 전체 지역 또는 앞뒤 지역 안내로 이동할 수 있습니다.</p></div></div><div class="local-navigation"><a class="local-nav-card is-parent" href="{esc(page_path())}"><small>카테고리</small><strong>{CATEGORY_NAME} 전체 지역</strong><span>↑</span></a>{prev_link}{next_link}</div></div></section>
  </main>{footer("../../../")}<script src="../../../assets/site14.js" defer></script>
</body></html>'''


def grouped_records(records: list[dict]) -> dict[str, dict[str, list[dict]]]:
    grouped: dict[str, dict[str, list[dict]]] = defaultdict(lambda: defaultdict(list))
    for record in records:
        grouped[record["center"]["region"]][record["center"]["district"]].append(record)
    for districts in grouped.values():
        for values in districts.values():
            values.sort(key=lambda item: item["locality"])
    return grouped


def hub_graph(records: list[dict]) -> dict:
    url = canonical_url()
    items = [{"@type": "ListItem", "position": index, "name": record["title"], "url": canonical_url(record["slug"])} for index, record in enumerate(records, 1)]
    return {"@context": "https://schema.org", "@graph": [
        {"@type": "EducationalOrganization", "@id": SITE_ORIGIN + "/#organization", "name": SITE_NAME, "url": SITE_ORIGIN + "/", "telephone": PHONE, "teaches": [f"{COURSE_NAME} 국어", f"{COURSE_NAME} 영어", f"{COURSE_NAME} 수학", "학습코칭"]},
        {"@type": "BreadcrumbList", "@id": url + "#breadcrumb", "itemListElement": [{"@type": "ListItem", "position": 1, "name": "홈", "item": SITE_ORIGIN + "/"}, {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": absolute_site_url("/과목별학원/")}, {"@type": "ListItem", "position": 3, "name": CATEGORY_NAME, "item": url}]},
        {"@type": "CollectionPage", "@id": url + "#collection", "url": url, "name": f"{CATEGORY_NAME} 지역 안내", "description": f"371개 동네별 {LEVEL_NAME} 국어·영어·수학 학습 상담 정보와 센터 자료를 확인하는 지역 허브입니다.", "inLanguage": "ko-KR", "breadcrumb": {"@id": url + "#breadcrumb"}, "hasPart": [{"@type": "WebPage", "name": record["title"], "url": canonical_url(record["slug"])} for record in records]},
        {"@type": "ItemList", "@id": url + "#directory", "name": f"{CATEGORY_NAME} 371개 지역", "numberOfItems": len(items), "itemListElement": items},
    ]}


def render_category_hub(records: list[dict]) -> str:
    grouped = grouped_records(records)
    region_order = [region for region in REGION_ORDER if region in grouped] + sorted(set(grouped) - set(REGION_ORDER))
    region_buttons = '<button type="button" class="is-active" data-region-filter="all">전체</button>' + "".join(f'<button type="button" data-region-filter="{esc(region)}">{esc(region)}</button>' for region in region_order)
    blocks = []
    for r_index, region in enumerate(region_order):
        districts_html = []
        for district, values in sorted(grouped[region].items()):
            cards = "".join(f'<a class="directory-card" href="{esc(page_path(record["slug"]))}" data-locality="{esc(record["title"])}"><strong>{esc(record["locality"])}</strong><span>{LEVEL_NAME} 국어·영어·수학 안내</span><i aria-hidden="true">→</i></a>' for record in values)
            districts_html.append(f'<section class="directory-district"><div class="directory-district-head"><h2>{esc(district)}</h2><span>{len(values)}개 지역</span></div><div class="directory-grid">{cards}</div></section>')
        count = sum(len(values) for values in grouped[region].values())
        blocks.append(f'<details class="directory-region" data-region="{esc(region)}"{" open" if r_index == 0 else ""}><summary><span><b>{esc(region)}</b><small>{len(grouped[region])}개 시군구 · {count}개 동네</small></span><i aria-hidden="true">+</i></summary><div class="directory-region-body">{"".join(districts_html)}</div></details>')
    graph_json = json.dumps(hub_graph(records), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    hub_url = canonical_url()
    hub_title = f"{CATEGORY_NAME} 지역 안내 | {TITLE_SUFFIX}"
    hub_description = f"371개 동네별 {CATEGORY_NAME} 페이지에서 국어·영어·수학 진단, {SCHOOL_NAME} 자료, 가능 학년과 센터 상담 정보를 확인하세요."
    hub_social_description = f"광역지역과 시군구별로 371개 {CATEGORY_NAME} 안내를 찾을 수 있습니다."
    rep_name = records[0]["rep_name"]
    rep_url = absolute_site_url(f"/assets/representative/{rep_name}")
    rep_width, rep_height = image_dimensions(REP_TARGET / rep_name)
    rep_mime = image_mime_type(rep_name)
    return f'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{hub_title}</title><meta name="description" content="{hub_description}"><meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec"><link rel="canonical" href="{hub_url}"><meta property="og:url" content="{hub_url}"><meta property="og:type" content="website"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="{hub_title}"><meta property="og:description" content="{hub_social_description}"><meta property="og:image" content="{rep_url}"><meta property="og:image:secure_url" content="{rep_url}"><meta property="og:image:type" content="{rep_mime}"><meta property="og:image:width" content="{rep_width}"><meta property="og:image:height" content="{rep_height}"><meta property="og:image:alt" content="{CATEGORY_NAME} 지역 안내 대표 이미지"><meta name="twitter:card" content="summary"><meta name="twitter:title" content="{hub_title}"><meta name="twitter:description" content="{hub_social_description}"><meta name="twitter:image" content="{rep_url}"><meta name="twitter:image:alt" content="{CATEGORY_NAME} 지역 안내 대표 이미지"><link rel="icon" href="../../assets/favicon.png"><link rel="stylesheet" href="../../assets/site14.css"><script type="application/ld+json">{graph_json}</script></head>
<body data-page="subjects">{header("../../")}<main id="main"><section class="directory-hero"><div class="site-shell"><nav class="breadcrumbs" aria-label="현재 위치"><a href="/">홈</a><a href="/과목별학원/">과목별학원</a><span>{CATEGORY_NAME}</span></nav><p class="eyebrow">National subject directory</p><h1>동네별 {CATEGORY_NAME}</h1><p>{HUB_INTRO}</p><div class="hub-metrics"><div><strong>371</strong><span>지역 페이지</span></div><div><strong>{HUB_FOCUS}</strong><span>{LEVEL_NAME} 핵심 기준</span></div><div><strong>{HUB_PROCESS}</strong><span>상담 확인 흐름</span></div></div></div></section>
<section class="section directory-section"><div class="site-shell"><div class="directory-toolbar"><label for="local-search">동네 이름으로 찾기</label><div class="directory-search"><input id="local-search" type="search" placeholder="예: 명일동, 불당동, 중계동" autocomplete="off" data-local-search><span data-directory-count>전체 371개 지역</span></div><div class="region-filters" aria-label="광역지역 선택">{region_buttons}</div><div class="directory-actions"><button type="button" data-expand-all>모두 펼치기</button><button type="button" data-collapse-all>모두 접기</button></div></div><div class="directory-empty" data-directory-empty hidden>검색 조건에 맞는 동네가 없습니다.</div><div class="directory-list">{"".join(blocks)}</div></div></section>
<section class="section ink"><div class="site-shell consult-cta"><div><h2>{HUB_CTA}</h2><p>{HUB_CTA_BODY}</p></div><a class="button orange" href="/상담문의/">상담 방법 확인 <span aria-hidden="true">→</span></a></div></section></main>{footer("../../")}<script src="../../assets/site14.js" defer></script></body></html>'''


def root_hub_graph() -> dict:
    category_pages = [
        {"@type": "WebPage", "name": category["label"], "url": f"/과목별학원/{category['slug']}/"}
        for category in SUBJECT_CATALOG
    ]
    category_items = [
        {"@type": "ListItem", "position": category["order"], "name": category["label"], "url": f"/과목별학원/{category['slug']}/"}
        for category in SUBJECT_CATALOG
    ]
    return {"@context": "https://schema.org", "@graph": [
        {"@type": "EducationalOrganization", "@id": "/#organization", "name": SITE_NAME, "url": "/", "telephone": PHONE, "teaches": ["국어", "영어", "수학", "학습코칭"]},
        {"@type": "BreadcrumbList", "@id": "/과목별학원/#breadcrumb", "itemListElement": [{"@type": "ListItem", "position": 1, "name": "홈", "item": "/"}, {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": "/과목별학원/"}]},
        {"@type": "CollectionPage", "@id": "/과목별학원/#collection", "url": "/과목별학원/", "name": "과목별학원", "description": "학생의 학년과 과목 조합에 맞춰 지역별 학습 상담 정보를 찾는 과목별학원 허브입니다.", "inLanguage": "ko-KR", "breadcrumb": {"@id": "/과목별학원/#breadcrumb"}, "hasPart": category_pages},
        {"@type": "ItemList", "name": "과목별학원 카테고리", "numberOfItems": len(category_items), "itemListElement": category_items},
    ]}


def render_root_hub() -> str:
    graph_json = json.dumps(root_hub_graph(), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    cards = "".join(
        f'<a class="subject-hub-card" href="{esc(category["slug"])}/index.html"><span>{esc(category["english"].upper())} / {category["order"]:02d}</span><h3>{esc(category["label"])}</h3><p>{esc(category["description"])}</p><div><b>371개 지역</b><i aria-hidden="true">→</i></div></a>'
        for category in SUBJECT_CATALOG
    )
    return f'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>과목별학원 | {TITLE_SUFFIX}</title><meta name="description" content="학년과 과목 조합에 맞춰 지역별 학습 상담 정보를 찾을 수 있도록 과목별학원 카테고리와 선택 기준을 정리했습니다."><meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec"><meta property="og:type" content="website"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="과목별학원 | {TITLE_SUFFIX}"><meta property="og:description" content="학년·과목별 학습 목표와 지역 센터 정보를 한 흐름에서 확인하세요."><link rel="icon" href="../assets/favicon.png"><link rel="stylesheet" href="../assets/site14.css"><script type="application/ld+json">{graph_json}</script></head><body data-page="subjects">{header("../")}<main id="main"><section class="directory-hero subjects-root-hero"><div class="site-shell"><nav class="breadcrumbs" aria-label="현재 위치"><a href="../index.html">홈</a><span>과목별학원</span></nav><p class="eyebrow">Subject academy guide</p><h1>과목별학원</h1><p>같은 학년이라도 과목 조합과 막히는 지점에 따라 확인할 수업 기록이 다릅니다. 학생에게 필요한 카테고리를 선택한 뒤 지역별 센터 정보와 학습 안내를 확인하세요.</p></div></section><section class="section"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>01</span> Category</div><div><h2>현재 확인할 수 있는 학원 안내</h2><p>검증된 원고와 센터정보를 기준으로 학년과 과목 조합을 구분했습니다.</p></div></div><div class="subject-hub-grid">{cards}</div></div></section><section class="section blue-wash"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>02</span> How to choose</div><div><h2>카테고리를 고를 때 확인할 세 가지</h2><p>과목명만 보고 결정하지 않고 학생의 실제 자료와 센터 운영 범위를 함께 확인합니다.</p></div></div><div class="role-grid"><article class="role-card"><span class="icon">01</span><h3>학생의 현재 학년</h3><p>페이지 제목과 별개로 희망 센터에서 해당 학년과 과목을 실제로 운영하는지 확인합니다.</p></article><article class="role-card"><span class="icon">02</span><h3>최근 평가와 교재</h3><p>점수만 말하기보다 틀린 문제와 현재 교재를 준비해 과목별 병목을 구분합니다.</p></article><article class="role-card"><span class="icon">03</span><h3>주간 실행 가능성</h3><p>학교 일정, 귀가 시간과 복습 시간을 함께 놓고 무리하지 않는 과제량을 확인합니다.</p></article></div></div></section></main>{footer("../")}<script src="../assets/site14.js" defer></script></body></html>'''


def build_records() -> list[dict]:
    centers = load_centers()
    manuscripts = load_manuscripts()
    representatives = choose_representatives(len(manuscripts))
    records: list[dict] = []
    seen_slugs: set[str] = set()
    for sections, rep_name in zip(manuscripts, representatives):
        title = clean(sections["페이지타이틀"])
        locality = locality_from_title(title)
        slug = folder_slug(locality)
        if slug in seen_slugs:
            raise ValueError(f"Duplicate slug: {slug}")
        if locality not in centers:
            raise ValueError(f"Center mapping missing: {locality}")
        seen_slugs.add(slug)
        records.append({"sections": sections, "title": title, "locality": locality, "slug": slug, "center": centers[locality], "meta": compact_meta(title, centers[locality]), "rep_name": rep_name})
    return records


def render_outputs(records: list[dict]) -> dict[Path, str]:
    outputs: dict[Path, str] = {}
    for index, record in enumerate(records):
        previous_record = records[(index - 1) % len(records)]
        next_record = records[(index + 1) % len(records)]
        target = CATEGORY_ROOT / record["slug"] / "index.html"
        if not target.is_file():
            raise FileNotFoundError(f"Refusing to create an unreviewed route: {target}")
        outputs[target] = render_page(record, previous_record, next_record)
    hub = CATEGORY_ROOT / "index.html"
    if not hub.is_file():
        raise FileNotFoundError(f"Category hub missing: {hub}")
    outputs[hub] = render_category_hub(records)
    return outputs


def visible_text(source: str) -> str:
    source = re.sub(r"<(?:script|style)\b.*?</(?:script|style)>", " ", source, flags=re.I | re.S)
    return clean(html.unescape(re.sub(r"<[^>]+>", " ", source)))


def extract_one(pattern: str, source: str) -> str:
    match = re.search(pattern, source, flags=re.I | re.S)
    return html.unescape(match.group(1)).strip() if match else ""


def normalized_copy(value: str, record: dict) -> str:
    center = record["center"]
    replacements = {
        record["title"], record["locality"], center["region"], center["display_region"],
        center["district"], center["display_district"],
        center["center_name"], center["address"], LEVEL_NAME, COURSE_NAME, CATEGORY_SLUG,
        *center["schools"], *center["korean_grades"], *center["english_grades"], *center["math_grades"],
    }
    result = clean(value)
    for token in sorted((item for item in replacements if len(item) >= 2), key=len, reverse=True):
        result = result.replace(token, "{X}")
    result = re.sub(r"\d+(?:[.,]\d+)*", "{N}", result)
    return re.sub(r"\{X\}(?:[·,/]\{X\})+", "{X}", result)


def duplicate_metric(index: dict[str, set[str]]) -> dict[str, int]:
    duplicates = [pages for pages in index.values() if len(pages) > 1]
    return {
        "unique": len(index),
        "duplicate_patterns": len(duplicates),
        "affected_pages": len(set().union(*duplicates)) if duplicates else 0,
        "max_df": max((len(pages) for pages in index.values()), default=0),
    }


def preflight(records: list[dict], outputs: dict[Path, str], include_corpus: bool = False) -> dict:
    errors: list[str] = []
    details = {path: source for path, source in outputs.items() if path != CATEGORY_ROOT / "index.html"}
    if len(outputs) != 372 or len(details) != 371:
        errors.append(f"scope_count={len(outputs)} details={len(details)} expected=372/371")
    expected_paths = {CATEGORY_ROOT / record["slug"] / "index.html" for record in records} | {CATEGORY_ROOT / "index.html"}
    if set(outputs) != expected_paths:
        errors.append("rendered path set differs from the existing 371-detail-plus-hub scope")

    descriptions: set[str] = set()
    exact_paragraph: dict[str, set[str]] = defaultdict(set)
    normalized_paragraph: dict[str, set[str]] = defaultdict(set)
    exact_sentence: dict[str, set[str]] = defaultdict(set)
    normalized_sentence: dict[str, set[str]] = defaultdict(set)
    normalized_faq: dict[str, set[str]] = defaultdict(set)
    source_note_exact: dict[str, set[str]] = defaultdict(set)
    hero_fact_exact: dict[str, set[str]] = defaultdict(set)
    hero_fact_normalized: dict[str, set[str]] = defaultdict(set)
    title_counts: list[int] = []
    locality_density: list[float] = []
    known_errors = (
        "고이 포함됩니다", "고등학교이 포함됩니다", "중학교이 포함됩니다",
        "초이 포함됩니다", "초등학교이 포함됩니다", "에 대한 설명는", "기록라는 표현",
        "관리을", "관리이라는", "이라는 이름으로 검색되더라도",
        "학부모님이 체감하는 현실은 학부모님이", "관리 기준 방문 전에는",
        "처럼 운영과 관련된 표현은", "영역으로 해석할 수 있습니다",
        "학습 점검의 내신 대비", "학습 계획의 내신 대비",
        "학습 점검 시험 후에는", "학습 계획 시험 후에는",
        "학습 과정 시험 후에는", "세 과목 학습 상담 시험 후에는",
        "관리 기준 시험 후에는", ", 그리고",
        "고려하면, 그리고", "살펴보면, 그리고",
        "상담 준비에서는 다음 기준을 적용합니다:",
        "비교할 때 화려한 설명보다 중요한 기준은 다음과 같습니다:",
        "확인하려는 이유는 단순한 수업보다",
        "과제 완료 기준을 기준으로",
        "계획보다는 점수보다",
        "지역명을 바꾼 홍보 문구",
        "하교 후 가능한 요일를",
        "가정에서 확인 가능한 복습 시간를",
        "시험·교과 범위을",
        "단원별 이해도을",
        "재풀이 결과을",
        "현재 교재 진도을",
        "틀린 이유 분류을",
        "학습코칭학원로",
        "와와학교 일정 점검학원",
        "요일를", "시간를", "범위을", "이해도을", "결과을", "진도을", "분류을", "학원로",
        "처럼 운영 관리와 관련된 표현은", "투명성으로 해석될 수 있습니다",
        "상담 전에는 방문 위치는", "문의 기준으로 방문 위치는",
        "실제 독해에 적용하는 속도가 느린 방학 복습이 필요한",
        "짧은 단위 복습에는 반응이 좋은 학습 공백을 줄이고 싶은",
        "설정한 대표 학생 유형은",
        "집중 시간은 짧지만 짧은 단위 복습에는 잘 반응하지만",
        "이 상담 사례의 학생을 대표 상담 사례로 살펴봅니다",
        "상담에서는 국영수를 함께 관리한다는 말은",
        "같은 과목별 고민을",
        "초등 과정이 되며 중등 준비 질문이 늘어나는",
        "비슷한 복습 문제가 나타나는",
        "학습 부담을 느끼는 초등 과정 특성",
    )
    bad_elementary = (
        "학원매출관리", "학원창업", "학원미납관리", "학원고객관리시스템",
        "학원전자계약", "학원관리솔루션", "학원결제시스템", "까지 고려한",
    )
    bad_reader_seeds = (
        "녹화수업", "온라인수업", "방학캠프", "일대일수업", "야간수업",
        "입시성공사례", "학원자료실", "학습암기", "학원매출관리", "학원창업",
        "학원미납관리", "학원고객관리시스템", "학원전자계약", "학원관리솔루션",
        "학원결제시스템", "학원소수정예", "학원출입관리", "학원결제관리",
        "학원고객관리", "학원강사", "학원위치", "학원운영", "학원행정",
    )
    for record in records:
        path = CATEGORY_ROOT / record["slug"] / "index.html"
        source = details[path]
        rel = path.relative_to(ROOT).as_posix()
        title, locality = record["title"], record["locality"]
        expected_url = canonical_url(record["slug"])
        canonical = extract_one(r'<link\s+rel="canonical"\s+href="([^"]+)"', source)
        og_url = extract_one(r'<meta\s+property="og:url"\s+content="([^"]+)"', source)
        if canonical != expected_url or og_url != expected_url or source.count('rel="canonical"') != 1 or source.count('property="og:url"') != 1:
            errors.append(f"{rel}: canonical/og:url mismatch")
        expected_rep_url = absolute_site_url(f"/assets/representative/{record['rep_name']}")
        expected_rep_width, expected_rep_height = image_dimensions(REP_TARGET / record["rep_name"])
        detail_social = {
            "og:image": expected_rep_url,
            "og:image:secure_url": expected_rep_url,
            "og:image:type": image_mime_type(record["rep_name"]),
            "og:image:width": str(expected_rep_width),
            "og:image:height": str(expected_rep_height),
            "twitter:image": expected_rep_url,
        }
        for name, expected in detail_social.items():
            attribute = "property" if name.startswith("og:") else "name"
            actual = extract_one(rf'<meta\s+{attribute}="{re.escape(name)}"\s+content="([^"]+)"', source)
            if actual != expected:
                errors.append(f"{rel}: social meta {name}={actual!r} expected={expected!r}")
        h1s = [visible_text(value) for value in re.findall(r"<h1\b[^>]*>(.*?)</h1>", source, flags=re.I | re.S)]
        if h1s != [title]:
            errors.append(f"{rel}: H1={h1s!r} expected={title!r}")
        h2s = [visible_text(value) for value in re.findall(r"<h2\b[^>]*>(.*?)</h2>", source, flags=re.I | re.S)]
        if any(heading.count("국어·영어·수학") > 1 for heading in h2s):
            errors.append(f"{rel}: repeated subject phrase in H2")
        meta = extract_one(r'<meta\s+name="description"\s+content="([^"]+)"', source)
        if not 65 <= len(meta) <= 80:
            errors.append(f"{rel}: meta length={len(meta)}")
        if meta in descriptions:
            errors.append(f"{rel}: duplicate meta description")
        descriptions.add(meta)
        faq_pairs = [
            (visible_text(question), visible_text(answer))
            for question, answer in re.findall(r'<details(?:\s+[^>]*)?>\s*<summary>(.*?)</summary>\s*<p>(.*?)</p>\s*</details>', source, flags=re.I | re.S)
        ]
        if len(faq_pairs) != 4:
            errors.append(f"{rel}: FAQ count={len(faq_pairs)}")
        for question, answer in faq_pairs:
            normalized_faq[normalized_copy(question + "\t" + answer, record)].add(rel)
            if re.search(r"(?:상담\s+전에는|문의\s+기준으로)\s+방문\s+위치는", answer):
                errors.append(f"{rel}: malformed FAQ location topic")
        if len(re.findall(r'<article\s+class="scenario-card">', source)) != 2:
            errors.append(f"{rel}: scenario count is not 2")
        try:
            graph = json.loads(extract_one(r'<script\s+type="application/ld\+json">(.*?)</script>', source))
            nodes = graph.get("@graph", [])
            article = next(node for node in nodes if node_has_type(node, "Article"))
            faq_node = next(node for node in nodes if node_has_type(node, "FAQPage"))
            if article.get("datePublished") != DATE_PUBLISHED or article.get("dateModified") != DATE_MODIFIED:
                errors.append(f"{rel}: Article dates")
            schema_pairs = [(item.get("name", ""), item.get("acceptedAnswer", {}).get("text", "")) for item in faq_node.get("mainEntity", [])]
            if schema_pairs != faq_pairs:
                errors.append(f"{rel}: visible/schema FAQ mismatch")
            mentioned = {clean(item.get("name", "")) for item in article.get("mentions", []) if isinstance(item, dict)}
            if not set(record["center"]["schools"]).issubset(mentioned):
                errors.append(f"{rel}: Article school mentions")
        except (ValueError, StopIteration, AttributeError, TypeError, json.JSONDecodeError) as exc:
            errors.append(f"{rel}: JSON-LD {exc}")
        visible = visible_text(source)
        for phrase in known_errors:
            if phrase in visible:
                errors.append(f"{rel}: known copy error {phrase!r}")
        for phrase in bad_reader_seeds:
            if phrase in visible:
                errors.append(f"{rel}: unverified or irrelevant reader seed {phrase!r}")
        if PROFILE == "elementary":
            for phrase in bad_elementary:
                if phrase in visible:
                    errors.append(f"{rel}: irrelevant elementary seed {phrase!r}")
            entrance_count = visible.count("입시")
            school_exam_count = len(re.findall(r"(?<!별)내신", visible))
            if entrance_count or school_exam_count:
                errors.append(f"{rel}: elementary intent drift 입시={entrance_count} 내신={school_exam_count}")
        if re.search(r'<img\b[^>]*style="[^"]*display\s*:\s*none', source, flags=re.I):
            errors.append(f"{rel}: hidden body image")
        for attrs in re.findall(r"<img\b([^>]*)>", source, flags=re.I | re.S):
            if not re.search(r"\bwidth=", attrs) or not re.search(r"\bheight=", attrs):
                errors.append(f"{rel}: visible image missing dimensions")
                break
        manuscript = extract_one(
            r'<section\s+class="section manuscript-wrap"[^>]*>(.*?)</article>\s*</section>\s*<section\s+class="section blue-wash"',
            source,
        )
        manuscript_visible = visible_text(manuscript)
        meta_copy_terms = (
            "키워드", "검색어", "검색엔진", "SEO", "상위노출", "템플릿",
            "표현으로 보는", "영역으로 해석", "보장 문구", "단어가 보여도",
            "말로 보는 것이 안전", "정보성 페이지 형태", "지역명을 바꾼 홍보 문구",
            "광고 문구", "광고처럼", "광고의 크기",
        )
        for meta_term in meta_copy_terms:
            if meta_term in manuscript_visible:
                errors.append(f"{rel}: reader-facing meta copy {meta_term!r}")
        if re.search(r"(?<![가-힣])원고(?=(?:\s|는|를|가|에서|의|$))", manuscript_visible):
            errors.append(f"{rel}: reader-facing meta copy '원고'")
        if re.search(
            r"(?:이\s+(?:글|페이지)|페이지(?:에서는|에는|에서|의|에|를)|"
            r"(?:을|를)\s+검색(?:한|하는|했다면))",
            manuscript_visible,
        ):
            errors.append(f"{rel}: reader-facing page/search self-reference")
        if re.search(r"광고(?:처럼|\s*문구|의\s+크기|보다)", manuscript_visible):
            errors.append(f"{rel}: reader-facing advertising comparison")
        if re.search(
            r"처럼\s+운영(?:\s+관리)?와\s+관련된\s+표현은|투명성으로\s+해석",
            manuscript_visible,
        ):
            errors.append(f"{rel}: reader-facing operational meta copy")
        for generic_locality in ("해당 지역", "이 지역"):
            if generic_locality in manuscript_visible:
                errors.append(f"{rel}: generic locality replacement {generic_locality!r}")
        center = record["center"]
        locality_forms = {
            clean(value)
            for value in (
                locality,
                locality.split()[-1],
                center.get("display_locality", ""),
            )
            if clean(value)
        }
        locality_form_pattern = "|".join(
            re.escape(value) for value in sorted(locality_forms, key=len, reverse=True)
        )
        district_stem = re.sub(r"(?:시|군|구)$", "", center.get("district", ""))
        duplicate_city_patterns: list[str] = []
        if district_stem and re.match(rf"{re.escape(district_stem)}(?:\s|$)", locality):
            duplicate_city_patterns.append(
                rf"{re.escape(center['district'])}\s+{re.escape(district_stem)}(?=\s|[,.!?)]|$)"
            )
        if center.get("region") and center["region"] == district_stem:
            duplicate_city_patterns.append(
                rf"{re.escape(center['region'])}\s+{re.escape(center['district'])}(?=\s|[,.!?)]|$)"
            )
        elif center.get("region") and re.match(
            rf"{re.escape(center['region'])}(?:\s|$)", locality
        ):
            duplicate_city_patterns.append(
                rf"{re.escape(center['region'])}\s+{re.escape(center['district'])}\s+"
                rf"{re.escape(center['region'])}(?=\s|[,.!?)]|$)"
            )
        if any(re.search(pattern, visible) for pattern in duplicate_city_patterns):
            errors.append(f"{rel}: duplicated city prefix in reader-facing join")
        if center.get("district"):
            awkward_locality = re.compile(
                rf"(?:(?:{re.escape(center.get('region', ''))}\s+)?{re.escape(center['district'])}\s+)(?:해당\s+지역|이\s+지역)"
            )
            if awkward_locality.search(manuscript_visible):
                errors.append(f"{rel}: awkward locality replacement")
        if center["region"] in {"충청", "경상", "전라"} and center.get("district"):
            broad_join = f"{center['region']} {center['district']}"
            if broad_join in manuscript_visible:
                errors.append(f"{rel}: unofficial broad region join {broad_join!r}")
        if not center.get("display_district") and center.get("district"):
            invalid_display = f"{center['display_region']} {center['district']} {locality}"
            if invalid_display in manuscript_visible:
                errors.append(f"{rel}: non-administrative district in prose {invalid_display!r}")
        region_fact = extract_one(
            r'<dt>지역</dt><dd>(.*?)</dd>',
            source,
        )
        expected_region_fact = " ".join(
            part for part in (center["display_region"], center["display_district"], center["display_locality"]) if part
        )
        if visible_text(region_fact) != expected_region_fact:
            errors.append(
                f"{rel}: reader-facing region={visible_text(region_fact)!r} expected={expected_region_fact!r}"
            )
        center_fact = visible_text(extract_one(r'<dt>센터 기준</dt><dd>(.*?)</dd>', source))
        address_fact = visible_text(extract_one(r'<dt>제공 주소</dt><dd>(.*?)</dd>', source))
        registration_fact = visible_text(extract_one(r'<dt>등록 정보</dt><dd>(.*?)</dd>', source))
        if center_fact != center["center_name"] or address_fact != center["address"]:
            errors.append(f"{rel}: authoritative center/address fact mismatch")
        if registration_fact != center["registration"]:
            errors.append(f"{rel}: authoritative registration fact mismatch")
        source_note = visible_text(extract_one(r'<p class="center-verified-note">(.*?)</p>', source))
        source_note_exact[source_note].add(rel)
        hero_fact = visible_text(extract_one(r'<div class="local-answer-grid">(.*?)</div>', source))
        hero_fact_exact[hero_fact].add(rel)
        hero_fact_normalized[normalized_copy(hero_fact, record)].add(rel)
        address_context = re.compile(
            r"(?:(?:방문\s+상담\s+)?주소는|제공된\s+주소는|주소\s+표기|제공\s+주소|주소\s+기준|센터\s+주소)"
        )
        manuscript_address = clean(center.get("address", ""))
        for sentence in re.split(r"(?<=[.!?])\s+", manuscript_visible):
            if address_context.search(sentence) and manuscript_address not in sentence:
                errors.append(f"{rel}: partial or altered address context {sentence!r}")
        if re.search(r"\b[^.!?\n]{1,40}\s+기준\s+제공된", visible):
            errors.append(f"{rel}: malformed scenario locality join")
        if re.search(r"상담에서는\s+제공된\s+센터\s+주소는|기준\s+주소\s+기준", visible):
            errors.append(f"{rel}: malformed address consultation join")
        if PROFILE == "elementary" and re.search(
            r"상황이라면\s+[^.!?\n]{5,220}?학생입니다|초등\s*[1-6]학년\s+중",
            manuscript_visible,
        ):
            errors.append(f"{rel}: malformed elementary student frame")
        if re.search(
            r"[^.!?\n]{2,100}?(?:이|가)\s+필요한\s+"
            r"[^.!?\n]{2,50}?(?:이|가)\s+필요한\s+학생",
            manuscript_visible,
        ):
            errors.append(f"{rel}: repeated necessary-student frame")
        if re.search(r"대표\s+상담\s+사례는\s+[^.!?\n]{1,140}?학생입니다", manuscript_visible):
            errors.append(f"{rel}: malformed representative-student predicate")
        if "이 상담 사례의 학생의" in manuscript_visible:
            errors.append(f"{rel}: malformed representative-student possessive")
        if locality_form_pattern and re.search(
            rf"(?:{locality_form_pattern})의\s+이\s+상담\s+사례의\s+학생에게는",
            manuscript_visible,
        ):
            errors.append(f"{rel}: malformed locality representative-student join")
        if re.search(
            r"이\s+학생에게는\s+(?:내신\s+대비|국어|영어|수학|세\s+과목)(?:은|는)",
            manuscript_visible,
        ):
            errors.append(f"{rel}: duplicated student/topic particle")
        if re.search(
            r"이\s+학생에게는\s+(?:(?:시험\s+3주\s+전|시험\s+직전|중간고사\s+직후)에는|오답은)",
            manuscript_visible,
        ):
            errors.append(f"{rel}: duplicated student/time-or-error topic particle")
        if PROFILE == "high" and re.search(
            r"(?:학생에게는|에서는)\s+"
            r"(?:(?:시험\s+3주\s+전|시험\s+직전|중간고사\s+직후|기말\s+기간)에는|"
            r"(?:내신\s+대비|오답)(?:은|는))",
            manuscript_visible,
        ):
            errors.append(f"{rel}: duplicated high-school topic particle")
        if PROFILE == "middle":
            if re.search(
                r"(?:학습\s+계획|학습\s+점검|학습\s+과정|세\s+과목\s+학습\s+상담|관리\s+기준)"
                r"(?:을|를)\s+진행한\s+뒤\s+복습에\s+더\s+잘\s+맞습니다",
                manuscript_visible,
            ):
                errors.append(f"{rel}: malformed middle-school review join")
            if re.search(r"우리\s+아이가\s+[^.!?\n]{1,140}?학생일\s+때", manuscript_visible):
                errors.append(f"{rel}: malformed middle-school child frame")
        if PROFILE == "elementary":
            elementary_copy_errors = (
                r"초등\s+학생",
                r"초등\s+중\b",
                r"초등에는",
                r"초등\s+특성",
                r"초등[^.!?\n]{0,100}?학생이[^.!?\n]{0,180}?아이일수록",
                r"초[1-6]\s+중\b",
                r"초[1-6]에서\s+초[1-6](?:으)?로\s+(?:이어지는|올라가는)\s+시기\s+중",
                r"초등학생에게는[^.!?\n]{0,360}?초등학생의\s+특성에\s+맞춰",
                r"초등학생\s+가운데[^.!?\n]{1,120}?학생을\s+대표\s+사례로",
                r"초등에",
                r"초등(?:\s+과정)?\s+상태",
                r"초등\s+과정\s+(?:학생|중\b|시기|아이|특성)",
                r"현재\s+학년\s+흐름의\s+초등학생",
                r"초등학생이\s+[^.!?\n]{0,100}?(?:문제|고민|과제)(?:이|가)\s+"
                r"(?:나타나는|있는|남은)",
                r"비슷한\s+[^.!?\n]{0,100}?비슷한\s+"
                r"(?:과정|학습|보완|복습|어려움|공부|학년대)",
                r"같은\s+[^.!?\n]{0,100}?같은\s+"
                r"(?:성장|공부|학습|실수|지점|실행|주)",
                rf"(?:{locality_form_pattern})의\s+(?:"
                r"현재\s+학년(?:\s+흐름)?의|현재\s+과정의|"
                r"이\s+학습\s+단계의|같은\s+성장\s+단계의|"
                r"비슷한\s+과정에\s+있는|해당\s+학년대의|"
                r"비슷한\s+학년대의)",
                r"(?<![가-힣])이와\s+(?:학습|과정|복습|보완)",
            )
            if any(re.search(pattern, manuscript_visible) for pattern in elementary_copy_errors):
                errors.append(f"{rel}: malformed elementary-school prose")
        expected_reference = representative_situation_sentence(locality, record["slug"])
        for sentence in re.split(r"(?<=[.!?])\s+", manuscript_visible):
            if expected_reference in sentence and clean(sentence) != expected_reference:
                errors.append(f"{rel}: residual prefix before representative-situation rewrite")
                break
        title_count = manuscript_visible.count(title)
        title_counts.append(title_count)
        if title_count > 3:
            errors.append(f"{rel}: exact H1 phrase in manuscript={title_count}")
        density = 100 * manuscript_visible.count(locality) / max(1, len(manuscript_visible.split()))
        locality_density.append(density)
        if density > 8:
            errors.append(f"{rel}: locality density={density:.2f}/100 tokens")
        paragraph_texts = [visible_text(value) for value in re.findall(r"<p\b[^>]*>(.*?)</p>", manuscript, flags=re.I | re.S)]
        noun_base = r"(?:학습 계획|학습 점검|학습 과정|세 과목 학습 상담|관리 기준)"
        noun_chain_patterns = (
            rf"{noun_base}\s+방문\s+전에는",
            rf"{noun_base}\s+상담에서\s+(?:국어|영어|수학)\s+계획",
            rf"{noun_base}\s+수업\s+후\s+복습",
            rf"{noun_base}\s+수업은\s+진단",
            rf"{noun_base}\s+수업\s+뒤에\s+무엇",
            rf"{noun_base}\s+방문\s+상담\s+주소는",
            rf"{noun_base}\s+교재\s+점검에서는",
        )
        for noun_pattern in noun_chain_patterns:
            if re.search(noun_pattern, " ".join(paragraph_texts)):
                errors.append(f"{rel}: mechanical noun-chain frame")
                break
        transition_counts: dict[str, int] = defaultdict(int)
        for paragraph in paragraph_texts:
            for sentence in re.split(r"(?<=[.!?])\s+", paragraph):
                for prefix in TRANSITION_PREFIXES:
                    if sentence.startswith(prefix.strip()):
                        transition_counts[prefix.strip()] += 1
                        break
        if sum(transition_counts.values()) > 6 or max(transition_counts.values(), default=0) > 1:
            errors.append(
                f"{rel}: stock transitions total={sum(transition_counts.values())} "
                f"repeat={max(transition_counts.values(), default=0)}"
            )
        paragraphs = [value for value in paragraph_texts if len(value) >= 80]
        if len(paragraphs) != len(set(paragraphs)):
            errors.append(f"{rel}: within-page paragraph duplicate")
        for value in set(paragraphs):
            exact_paragraph[value].add(rel)
            normalized_paragraph[normalized_copy(value, record)].add(rel)
        sentences = [
            clean(value)
            for paragraph in paragraph_texts
            for value in re.split(r"(?<=[.!?])\s+", paragraph)
            if len(clean(value)) >= 30
        ]
        if len(sentences) != len(set(sentences)):
            errors.append(f"{rel}: within-page sentence duplicate")
        seven_word_windows: dict[tuple[str, ...], int] = defaultdict(int)
        for paragraph in paragraph_texts:
            words = paragraph.split()
            for index in range(max(0, len(words) - 6)):
                seven_word_windows[tuple(words[index:index + 7])] += 1
        repeated_seven = {
            words: count for words, count in seven_word_windows.items() if count >= 3
        }
        if repeated_seven:
            words, count = max(repeated_seven.items(), key=lambda item: item[1])
            errors.append(
                f"{rel}: within-page seven-word repetition={count} {' '.join(words)!r}"
            )
        answer_note = visible_text(extract_one(
            r'<div class="answer-note"><strong>대표 학생 상황</strong><p>(.*?)</p>',
            source,
        ))
        summary_html = extract_one(
            r'<div\s+class="local-summary"[^>]*>(.*?)</div>\s*</div>\s*</section>',
            source,
        )
        summary_blocks = [
            visible_text(value)
            for value in re.findall(r"<p\b[^>]*>(.*?)</p>", summary_html, flags=re.I | re.S)
        ]
        scenario_blocks = [
            visible_text(value)
            for value in re.findall(
                r'<article class="scenario-card">.*?<p>(.*?)</p>\s*</article>',
                source,
                flags=re.I | re.S,
            )
        ]
        authored_blocks = [
            *paragraph_texts,
            *summary_blocks,
            *scenario_blocks,
            *(value for pair in faq_pairs for value in pair),
        ]
        authored_eight_windows: dict[tuple[str, ...], int] = defaultdict(int)
        for block in authored_blocks:
            authored_eight_windows.update(Counter(copy_windows(block)))
        repeated_eight = {
            words: count for words, count in authored_eight_windows.items() if count > 2
        }
        if repeated_eight:
            words, count = max(repeated_eight.items(), key=lambda item: item[1])
            errors.append(
                f"{rel}: authored-block eight-word repetition={count} {' '.join(words)!r}"
            )
        main_visible = visible_text(extract_one(r"<main\b[^>]*>(.*?)</main>", source))
        main_words = main_visible.split()
        main_windows: dict[tuple[str, ...], int] = defaultdict(int)
        for index in range(max(0, len(main_words) - 6)):
            main_windows[tuple(main_words[index:index + 7])] += 1
        note_words = answer_note.split()
        note_repetition = max(
            (
                main_windows.get(tuple(note_words[index:index + 7]), 0)
                for index in range(max(0, len(note_words) - 6))
            ),
            default=0,
        )
        if note_repetition >= 3:
            errors.append(f"{rel}: representative-student phrase repeated {note_repetition} times")
        if max((len(value) for value in sentences), default=0) > 180:
            errors.append(f"{rel}: sentence longer than 180 characters")
        for value in set(sentences):
            exact_sentence[value].add(rel)
            normalized_sentence[normalized_copy(value, record)].add(rel)

    hub = outputs.get(CATEGORY_ROOT / "index.html", "")
    if extract_one(r'<link\s+rel="canonical"\s+href="([^"]+)"', hub) != canonical_url() or extract_one(r'<meta\s+property="og:url"\s+content="([^"]+)"', hub) != canonical_url():
        errors.append("category hub canonical/og:url mismatch")
    expected_hub_image = absolute_site_url(f"/assets/representative/{records[0]['rep_name']}")
    expected_width, expected_height = image_dimensions(REP_TARGET / records[0]["rep_name"])
    hub_social = {
        "og:image": expected_hub_image,
        "og:image:secure_url": expected_hub_image,
        "og:image:type": image_mime_type(records[0]["rep_name"]),
        "og:image:width": str(expected_width),
        "og:image:height": str(expected_height),
        "twitter:card": "summary",
        "twitter:image": expected_hub_image,
    }
    for name, expected in hub_social.items():
        attribute = "property" if name.startswith("og:") else "name"
        actual = extract_one(rf'<meta\s+{attribute}="{re.escape(name)}"\s+content="([^"]+)"', hub)
        if actual != expected:
            errors.append(f"category hub social meta {name}={actual!r} expected={expected!r}")
    metrics = {
        "exact_paragraph": duplicate_metric(exact_paragraph),
        "normalized_paragraph": duplicate_metric(normalized_paragraph),
        "exact_sentence": duplicate_metric(exact_sentence),
        "normalized_sentence": duplicate_metric(normalized_sentence),
        "normalized_faq_pair": duplicate_metric(normalized_faq),
        "source_note_exact": duplicate_metric(source_note_exact),
        "hero_fact_exact": duplicate_metric(hero_fact_exact),
        "hero_fact_normalized": duplicate_metric(hero_fact_normalized),
    }
    limits = {"exact_paragraph": 10, "normalized_paragraph": 50, "exact_sentence": 50, "normalized_sentence": 100, "normalized_faq_pair": 50}
    for name, limit in limits.items():
        if metrics[name]["max_df"] > limit:
            errors.append(f"cross-page {name} max_df={metrics[name]['max_df']} > {limit}")
    supplemental_limits = {
        "source_note_exact": 15,
        "hero_fact_exact": 3,
        "hero_fact_normalized": 15,
    }
    for name, limit in supplemental_limits.items():
        if metrics[name]["max_df"] > limit:
            errors.append(f"cross-page {name} max_df={metrics[name]['max_df']} > {limit}")
    report = {
        "profile": PROFILE,
        "mode": "check-only",
        "detail_pages": len(details),
        "category_hub": 1 if hub else 0,
        "description_length": {"min": min(map(len, descriptions), default=0), "max": max(map(len, descriptions), default=0), "unique": len(descriptions)},
        "exact_h1_phrase": {"max": max(title_counts, default=0), "over_3": sum(value > 3 for value in title_counts)},
        "locality_density": {"max": round(max(locality_density, default=0), 3), "over_8": sum(value > 8 for value in locality_density)},
        "duplicate_metrics": metrics,
        "errors": errors,
    }
    if include_corpus:
        report["_corpus"] = {
            name: {pattern: len(pages) for pattern, pages in index.items()}
            for name, index in (
                ("exact_paragraph", exact_paragraph),
                ("normalized_paragraph", normalized_paragraph),
                ("exact_sentence", exact_sentence),
                ("normalized_sentence", normalized_sentence),
                ("normalized_faq_pair", normalized_faq),
            )
        }
    if errors:
        raise ValueError(json.dumps(report, ensure_ascii=False, indent=2))
    return report


def outside_scope_hashes() -> dict[str, str]:
    hashes: dict[str, str] = {}
    selected = CATEGORY_ROOT.resolve()
    for path in ROOT.rglob("*.html"):
        resolved = path.resolve()
        try:
            resolved.relative_to(selected)
            continue
        except ValueError:
            pass
        hashes[path.relative_to(ROOT).as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    return hashes


def apply_outputs(outputs: dict[Path, str]) -> None:
    before = outside_scope_hashes()
    for path, source in outputs.items():
        if not path.is_file() or path.parent.parent != CATEGORY_ROOT and path.parent != CATEGORY_ROOT:
            raise RuntimeError(f"Refusing out-of-scope or new path: {path}")
        path.write_text(source, encoding="utf-8")
    after = outside_scope_hashes()
    if before != after:
        changed = sorted(set(before) | set(after), key=str)
        changed = [name for name in changed if before.get(name) != after.get(name)]
        raise RuntimeError(f"Out-of-scope HTML changed: {changed[:10]}")


def run_profile(apply: bool, include_corpus: bool = False) -> dict:
    records = build_records()
    outputs = render_outputs(records)
    report = preflight(records, outputs, include_corpus=include_corpus)
    if apply:
        apply_outputs(outputs)
        report["mode"] = "apply"
        report["written"] = len(outputs)
    return report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Safely check or apply one existing K/E/M category.")
    parser.add_argument("profile", nargs="?", choices=tuple(PROFILES), default="high")
    parser.add_argument("--all", action="store_true", help="preflight all three categories before any requested apply")
    parser.add_argument("--apply", action="store_true", help="write only the selected category's 371 details and hub")
    parser.add_argument("--check", "--dry-run", dest="check", action="store_true", help="explicit check-only alias (the default)")
    parser.add_argument("--corpus", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.all:
        reports: list[dict] = []
        child_env = dict(__import__("os").environ, PYTHONUTF8="1")
        for profile in PROFILES:
            command = [sys.executable, "-X", "utf8", str(Path(__file__).resolve()), profile, "--check", "--corpus"]
            completed = subprocess.run(command, cwd=ROOT, env=child_env, text=True, encoding="utf-8", stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if completed.returncode:
                raise SystemExit(completed.stderr or completed.stdout)
            reports.append(json.loads(completed.stdout))
        combined_limits = {"exact_paragraph": 10, "normalized_paragraph": 50, "exact_sentence": 50, "normalized_sentence": 100, "normalized_faq_pair": 50}
        combined_max_df: dict[str, int] = {}
        for name, limit in combined_limits.items():
            counts: dict[str, int] = defaultdict(int)
            for report in reports:
                for pattern, count in report["_corpus"][name].items():
                    counts[pattern] += count
            combined_max_df[name] = max(counts.values(), default=0)
            if combined_max_df[name] > limit:
                raise SystemExit(f"Combined cross-profile {name} max_df={combined_max_df[name]} > {limit}")
        for report in reports:
            report.pop("_corpus", None)
        if args.apply:
            applied: list[dict] = []
            for profile in PROFILES:
                command = [sys.executable, "-X", "utf8", str(Path(__file__).resolve()), profile, "--apply"]
                completed = subprocess.run(command, cwd=ROOT, env=child_env, text=True, encoding="utf-8", stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if completed.returncode:
                    raise SystemExit(completed.stderr or completed.stdout)
                applied.append(json.loads(completed.stdout))
            reports = applied
        print(json.dumps({"mode": "apply" if args.apply else "check-only", "combined_max_df": combined_max_df, "profiles": reports}, ensure_ascii=False))
        return
    report = run_profile(apply=args.apply, include_corpus=args.corpus)
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
