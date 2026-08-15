from __future__ import annotations

import csv
import hashlib
import html
import json
import random
import re
import shutil
import sys
from collections import defaultdict
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
PROFILE = sys.argv[1].lower() if len(sys.argv) > 1 else "high"
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
PHONE = "010-6839-8283"
DATE = "2026-07-30"
REQUIRED = ("페이지타이틀", "메타설명", "본문", "FAQ", "학부모후기", "JSON-LD 요약")
REGION_ORDER = ["서울", "경기", "인천", "충청", "대전", "대구", "울산", "부산", "경상", "광주", "전라", "강원", "제주"]


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
    return base + (quote(locality_slug) + "/" if locality_slug else "")


def list_values(value: str) -> list[str]:
    return list(dict.fromkeys(part.strip() for part in re.split(r"[,/\n]+", value or "") if part.strip()))


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
        english_slug = clean(row["동 영어"]).replace(" ", "-")
        candidates = sorted(MAP_DIR.glob(english_slug + ".*"))
        if not candidates:
            raise FileNotFoundError(f"Map missing: {locality} / {english_slug}")
        image_row = image_rows.get(locality)
        if not image_row:
            raise ValueError(f"Image row missing: {locality}")
        body_image = "seoul6839.webp" if clean(image_row["본문"]).lower() == "seoul.jpg" else "local6839.webp"
        centers[locality] = {
            "locality": locality,
            "english_slug": english_slug,
            "region": clean(row["지역"]),
            "district": clean(row["시or구"]),
            "center_name": clean(row["센터명"]),
            "tuition_url": clean(row["센터 교습비"]),
            "office_name": clean(row["교육지원청명칭"]),
            "registration": clean(row["교육지원청 등록번호"]),
            "address": clean(row["센터 주소"]),
            "schools": list_values(row[SCHOOL_COLUMN]),
            "korean_grades": target_grades(row["가능학년\n(국어)"]),
            "english_grades": target_grades(row["가능학년\n(영어)"]),
            "math_grades": target_grades(row["가능학년\n(수학)"]),
            "map_name": candidates[0].name,
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
    REP_TARGET.mkdir(parents=True, exist_ok=True)
    existing: list[str] = []
    for index in range(1, count + 1):
        matches = sorted(REP_TARGET.glob(f"{REP_PREFIX}-{index:03d}.*"))
        if len(matches) != 1:
            existing = []
            break
        existing.append(matches[0].name)
    if len(existing) == count:
        return existing

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
    if len(unique) < count:
        raise ValueError(f"Unique representative images: {len(unique)} < {count}")
    rng = random.Random(f"site14-{CATEGORY_SLUG}-2026-07-30")
    rng.shuffle(unique)
    result: list[str] = []
    for index, source in enumerate(unique[:count], 1):
        name = f"{REP_PREFIX}-{index:03d}{source.suffix.lower()}"
        target = REP_TARGET / name
        # Existing representative assignments are stable publication assets.
        # Regenerating page copy must not silently reshuffle/replace them when
        # the reference folder changes later.
        if not target.exists():
            shutil.copy2(source, target)
        result.append(name)
    return result


def compact_meta(title: str, center: dict) -> str:
    area = " ".join(part for part in (center["region"], center["district"]) if part)
    value = (
        f"{area} {title} 선택 전 국어·영어·수학 진단, {META_FOCUS}, "
        "과목별 오답 관리와 센터 상담 확인 항목을 정리했습니다."
    )
    value = clean(value)
    if len(value) > 100:
        value = clean(f"{title} 선택 전 국어·영어·수학 진단, {META_FOCUS}, 과목별 오답 관리와 센터 상담 기준을 정리했습니다.")
    if len(value) < 70:
        value = value.rstrip(".") + ". 최근 시험지와 학습 기록을 기준으로 확인할 내용을 안내합니다."
    return value[:100].rstrip(" ,·")


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


def sanitize_unverified_operational_keyword(text: str, seed: str) -> str:
    """Turn unsourced service-like keywords into verifiable consultation topics.

    원고의 개별화용 키워드에는 `온라인수업`, `방학캠프`, `입시성공사례`
    같이 센터 자료로 확인되지 않은 운영 표현이 섞여 있다. 서비스 사실로
    보일 수 있는 단어는 최근 시험지·오답·학교 일정처럼 학부모가 실제
    자료로 확인할 수 있는 상담 주제로 바꾼다.
    """
    unsafe_terms = (
        "녹화수업", "온라인수업", "방학캠프", "일대일수업", "야간수업",
        "입시성공사례", "학원자료실", "학습암기",
    )
    keyword = next((value for value in unsafe_terms if value in text), "")
    patterns = (
        r"그리고\s+([가-힣A-Za-z0-9·]+)처럼\s+학부모님이\s+실제로\s+묻는\s+운영\s+요소",
        r"가정에서\s+확인할\s+([가-힣A-Za-z0-9·]+)과\s+학습\s+습관",
        r"확인해야\s+([가-힣A-Za-z0-9·]+)\s+운영도",
        r"([가-힣A-Za-z0-9·]+)이\s+상담\s+키워드로\s+제시된",
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
    ])
    text = replace_keyword_token(text, keyword, topic)
    replacements = {
        f"{topic}처럼 학부모님이 실제로 묻는 운영 요소": f"{topic}처럼 상담에서 확인할 자료",
        f"{topic} 운영도 실질적으로 이어집니다": f"{topic}도 실제 학습 계획에 반영할 수 있습니다",
        f"{topic}이 상담 키워드로 제시된": f"{topic}을 함께 확인하는",
        f"{topic} 안내를 통해": f"{topic}을 점검해",
        f"{topic}과 연결해": f"{topic}을 바탕으로",
        f"상담 시 확인할 {topic} 내용": f"상담 시 확인할 {topic}",
        f"{topic}까지 함께 알아보는": f"{topic}도 함께 확인하는",
        f"{topic}을 찾는다면": f"{topic}을 확인한다면",
        f"{topic}에 대한 안내": f"{topic}에 대한 설명",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
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
                f"세 과목 일정이 겹치는 {locality} {LEVEL_NAME}에게는 과제량과 복습 시간을 현실적으로 조정하는 과정이 필요합니다.",
                f"{locality} {LEVEL_NAME}은 시험 일정에 따라 집중 과목과 유지 과목을 구분해 주간 균형을 맞추는 편이 좋습니다.",
                f"한 과목에 시간이 몰리는 {locality} {LEVEL_NAME}이라면 다른 과목의 복습이 끊기지 않도록 최소 실행량을 정해야 합니다.",
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
    return text


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
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    if locality:
        text = text.replace(
            f"{locality} 국어·영어·수학 학습 계획 상담 전 준비할 질문",
            f"{locality} 국어·영어·수학 상담 전 준비할 질문",
        )
    text = re.sub(r"(\d+)층로", r"\1층으로", text)
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
    value = polish_korean(student_type, center["locality"])
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
        f"{school_context}, {student_type}에게 맞게 "
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
    cards = options[PROFILE]
    shift = int(hashlib.sha256(f"{seed}|answer-card".encode("utf-8")).hexdigest()[:4], 16) % len(cards)
    return cards[shift:] + cards[:shift]


def extract_student_type(body: str, locality: str) -> str:
    patterns = [
        r"이 페이지는 (.+?)을 기준 학생으로 두고,",
        r"대표 학생 유형은 (.+?)입니다\.",
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


def reduce_title_repetition(text: str, title: str, locality: str, seed: str, keep: int = 4) -> str:
    alternatives = [
        f"{locality} {LEVEL_NAME}의 국어·영어·수학 학습 점검",
        f"{locality} 국어·영어·수학 학습 계획",
        f"{locality} {LEVEL_NAME} 세 과목 학습 상담",
        f"이 지역의 {LEVEL_NAME} 학습 과정",
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


def personalize_body(body: str, title: str, locality: str, center: dict, seed: str) -> tuple[str, list[tuple[str, list[str]]]]:
    body = sanitize_unverified_operational_keyword(body, seed)
    body = polish_korean(body, locality)
    body = body.replace(f"{locality}에서 {locality} ", f"{locality}에서 ")
    body = sanitize_grade_claims(body, center)
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
    body = reduce_title_repetition(body, title, locality, seed, keep=4)
    body = diversify_common_copy(body, locality, seed)
    # Grade replacement can expose adjacent template fragments, so run the
    # conservative Korean cleanup once more before turning the source into HTML.
    body = polish_korean(body, locality)
    intro, sections = parse_body(body)
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
        paragraphs = [
            stable_pick(seed, "school-list", [
                f"제공된 {SCHOOL_NAME} 참고 목록은 {names}입니다. 이 목록은 수업 가능 여부를 보장하지 않으며, 상담에서 학생이 가져온 실제 학교 자료를 대조하기 위한 정보입니다.",
                f"{locality} 페이지의 {SCHOOL_NAME} 참고 정보에는 {names}이 포함됩니다. 학교명만으로 수업 가능 여부를 판단하지 않고 센터의 현재 개설 범위를 함께 확인해야 합니다.",
                f"상담 준비를 위해 제공된 {SCHOOL_NAME} 목록은 {names}입니다. 이는 수업 가능 학교를 단정하는 자료가 아니라 학생의 실제 범위와 과제를 확인하기 위한 참고 정보입니다.",
                f"센터 자료에서 확인한 {SCHOOL_NAME} 참고 목록은 {names}입니다. 실제 수업 여부는 학생 자료와 센터별 과목·학년 운영 범위를 확인한 뒤 판단합니다.",
                f"{names}은 제공 자료에 포함된 {locality} {SCHOOL_NAME} 참고 목록입니다. 목록 포함 여부와 실제 수업 가능 여부는 같지 않으므로 상담에서 별도로 확인합니다.",
                f"제공 자료에는 {names}이 {SCHOOL_NAME} 참고 정보로 정리되어 있습니다. 학교 이름은 상담의 출발점으로만 사용하고 실제 시험·교과 자료를 함께 확인합니다.",
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
                f"{locality} 페이지에 제공된 {SCHOOL_NAME} 목록은 없습니다. 확인되지 않은 학교명을 추가하지 않고 상담에서 학생이 실제 사용하는 {material_label}를 기준으로 확인합니다.",
                f"제공 자료에는 {locality} {SCHOOL_NAME} 목록이 따로 정리되어 있지 않습니다. 학생의 실제 {material_label}를 준비해 현재 범위와 과제 일정을 확인하는 것이 정확합니다.",
                f"{locality}의 확인된 {SCHOOL_NAME} 참고 목록이 없어 학교명을 임의로 제시하지 않습니다. 상담 시 학생이 사용하는 {material_label}와 센터 운영 범위를 직접 대조하세요.",
                f"학교 정보가 제공되지 않은 {locality} 페이지입니다. 특정 학교를 추정하지 않고 학생의 실제 {material_label}를 바탕으로 상담 질문을 준비합니다.",
                f"{locality} {SCHOOL_NAME} 목록은 제공 자료에서 확인되지 않았습니다. 실제 학교 자료와 희망 센터의 과목·학년 운영 여부를 함께 확인해야 합니다.",
                f"확인되지 않은 학교명을 넣지 않기 위해 {locality} 페이지에는 학교 목록을 표시하지 않습니다. 학생이 가져온 {material_label}를 기준으로 상담합니다.",
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
    return heading, paragraphs


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
    a1 = stable_pick(seed, "a1", [
        f"{student_type}이라면 최근 국어·영어·수학 시험지와 교재를 함께 놓고 막힌 지점을 과목별로 나눠 보는 상담이 도움이 될 수 있습니다. 등록 여부는 실제 진단과 센터 운영 범위를 확인한 뒤 결정해야 합니다.",
        f"대표적으로 {student_type}의 경우 세 과목을 한꺼번에 늘리기보다 최근 풀이 기록에서 우선 보완할 과목과 단원을 구분해야 합니다. {locality} 상담에서는 학생 시간표와 센터별 개설 범위도 함께 확인합니다.",
        f"{student_type}에게는 과제량을 바로 늘리는 방식보다 국어 근거 찾기, 영어 해석 과정, 수학 풀이 단계 중 어디에서 멈추는지 먼저 나누어 보는 점검이 적합할 수 있습니다.",
        f"현재 모습이 {student_type}에 가깝다면 최근 시험지, 학교 자료와 반복 오답을 준비해 원인을 분류해 보세요. 상담 결과와 실제 수업 가능 여부는 희망 센터에서 별도로 확인해야 합니다.",
        f"{locality} 상담에서 살펴볼 학생 모습은 {student_type}입니다. 이때 과목별 점수보다 시작 시각, 완료 분량, 틀린 이유와 재풀이 결과를 먼저 살펴보는 편이 좋습니다.",
        f"세 과목의 문제가 모두 같다고 보지 않습니다. {student_type}이라면 각 과목의 병목과 주간 시간 배분을 따로 확인한 뒤 현실적인 우선순위를 정해야 합니다.",
    ])
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
        "국영수를 함께 관리한다는 말은 세 과목을 동일하게 다룬다는 뜻이 아닙니다. 학생이 막힌 장면과 시험 일정을 과목별로 확인한 뒤 공통 시간표 안에서 충돌하지 않게 배치합니다.",
        "먼저 최근 자료에서 가장 시급한 과목을 찾고, 나머지 과목의 최소 복습 시간을 남겨야 합니다. 오답은 설명 직후가 아니라 일정 시간이 지난 뒤 다시 풀어 적용 여부를 확인합니다.",
    ])
    schools = center["schools"]
    if schools:
        shown = "·".join(schools[:4])
        q3 = stable_pick(seed, "q3-school", [
            f"{locality} {SCHOOL_NAME} 참고 정보는 상담에서 어떻게 활용하나요?",
            f"{locality} 학교별 내신 자료는 어떤 방식으로 확인하나요?",
            f"제공된 {locality} {SCHOOL_NAME} 목록은 수업 가능 학교를 뜻하나요?",
            f"{locality} {LEVEL_NAME}은 상담 때 어떤 학교 자료를 준비해야 하나요?",
        ])
        a3 = stable_pick(seed, "a3-school", [
            f"제공된 {SCHOOL_NAME} 참고 목록에는 {shown} 등이 포함됩니다. 이는 수업 가능 여부를 보장하는 목록이 아니며, 실제 {material_label}와 센터별 개설 과목을 상담에서 함께 확인해야 합니다.",
            f"{shown} 등은 제공 자료의 {SCHOOL_NAME} 참고 목록입니다. 학생이 가져온 {material_label}를 대조하고 희망 센터의 과목·학년 운영 여부를 별도로 확인하세요.",
            f"학교명은 상담 준비를 위한 참고 정보입니다. {shown} 등과 관련한 실제 수업 여부는 학생의 {material_label}, 센터의 현재 개설 범위와 함께 확인해야 합니다.",
            f"제공 목록에서 {shown} 등을 확인할 수 있지만 모든 학교의 수업 가능 여부를 뜻하지는 않습니다. 실제 범위와 과제 자료를 준비해 상담에서 확인하는 편이 정확합니다.",
            f"{locality} 페이지에는 {shown} 등이 {SCHOOL_NAME} 참고 정보로 표시됩니다. 학교 목록보다 학생의 현재 {material_label}와 센터 운영 범위를 대조하는 과정이 우선입니다.",
            f"{shown} 등은 확인된 참고 학교명이며 수업 가능 학교를 단정하지 않습니다. 학생의 실제 자료와 센터별 개설 과목을 함께 놓고 상담하세요.",
        ])
    else:
        q3 = stable_pick(seed, "q3-none", [
            f"{locality} {SCHOOL_NAME} 목록이 없는 경우 상담은 어떻게 준비하나요?",
            f"{locality} 페이지에 학교명이 없으면 내신 상담이 어려운가요?",
            f"제공된 {locality} {SCHOOL_NAME} 정보가 없을 때 무엇을 확인해야 하나요?",
            f"{locality} 학생의 학교 자료는 상담 때 직접 가져가야 하나요?",
        ])
        a3 = stable_pick(seed, "a3-none", [
            f"{locality}에 제공된 {SCHOOL_NAME} 목록이 없어 임의로 학교명을 추가하지 않았습니다. 학생이 실제 사용하는 {material_label}를 준비하고 희망 센터의 과목·학년 운영 범위를 직접 확인하는 것이 정확합니다.",
            f"제공 자료에서 {locality} {SCHOOL_NAME} 목록을 확인할 수 없어 특정 학교를 추정하지 않습니다. 현재 {material_label}와 센터의 가능 학년을 상담에서 함께 확인하세요.",
            f"학교명이 표시되지 않아도 상담은 가능합니다. 학생이 사용하는 {material_label}를 가져와 현재 범위와 오답을 설명하고 센터의 실제 개설 과목을 확인하면 됩니다.",
            f"{locality}의 확인된 학교 목록이 없으므로 임의 정보를 넣지 않았습니다. 실제 교재·학교 자료와 희망 센터의 과목 운영 범위를 직접 대조하는 방식이 정확합니다.",
            f"이 페이지에는 제공된 {SCHOOL_NAME} 정보가 없습니다. 학생의 {material_label}를 준비해 과목별 현재 진도와 센터 수업 가능 여부를 상담에서 확인하세요.",
            f"확인되지 않은 학교명을 추가하는 대신 학생이 실제 사용하는 {material_label}를 기준으로 안내합니다. 센터별 과목과 학년 운영은 등록 전에 다시 확인해야 합니다.",
        ])
    q4 = stable_pick(seed, "q4", [
        f"{locality} 상담 전에 센터 위치와 교습비는 어디에서 확인하나요?",
        f"{locality} 센터 방문 전에 확인할 운영 정보는 무엇인가요?",
        f"{locality} {LEVEL_NAME} 상담을 예약할 때 어떤 정보를 준비하면 좋나요?",
        f"{locality} 학원 상담에서 수업 가능 학년과 비용을 어떻게 확인하나요?",
        f"{locality} 센터의 주소와 과목 운영은 모두 같은가요?",
    ])
    location = f"제공된 센터 주소는 {center['address']}입니다. " if center["address"] else ""
    tuition = "페이지의 센터별 교습비 안내 버튼에서 제공 자료를 확인할 수 있습니다. " if center["tuition_url"] else "교습비 자료는 상담 시 직접 확인해야 합니다. "
    a4 = f"{locality} 문의 기준으로 " + location + tuition + f"개설 과목과 {LEVEL_NAME} 가능 학년, 시간표와 보강 방식은 센터별로 다를 수 있으므로 등록 전에 함께 확인하세요."
    return [(q1, a1), (q2, a2), (q3, a3), (q4, a4)]


def build_scenarios(locality: str, center: dict, student_type: str, seed: str) -> list[str]:
    material_label = SCHOOL_MATERIALS if SCHOOL_MATERIALS.endswith("자료") else f"{SCHOOL_MATERIALS} 자료"
    first = stable_pick(seed, "scenario-1", [
        f"{locality}에서 {student_type} 자녀에 대해 학부모가 최근 시험지와 오답 노트를 함께 준비한 상황입니다. 상담에서는 세 과목을 모두 늘리기보다 먼저 바꿀 과목과 복습 시점을 정하고, 설명이 실제 주간 계획으로 이어지는지를 확인합니다.",
        f"{student_type}의 상담을 가정한 예시입니다. 국어·영어·수학 점수만 비교하지 않고 시작이 늦어진 과제, 반복한 오답과 학교 일정을 나누어 본 뒤 이번 주에 실행할 한두 가지를 정리합니다.",
        f"{locality} 학부모가 자녀의 세 과목 학습량이 서로 충돌하는 문제를 질문한 상황을 예로 들었습니다. 상담 후에는 성적 약속보다 과목별 최소 복습량, 집중 단원과 재확인 날짜가 구체적인지 살펴봅니다.",
        f"최근 학교 학습 준비가 한 과목에 치우친 {locality} {LEVEL_NAME}을 가정했습니다. 학부모는 현재 교재와 학교 자료를 바탕으로 국어·영어·수학의 병목을 따로 듣고, 가정에서 확인할 기록을 정리합니다.",
        f"{student_type}의 경우를 바탕으로 만든 상담 상황 예시입니다. 학생의 하루 시간표에 과목별 과제와 오답 재풀이를 실제로 배치해 보고 무리한 분량은 줄이는 방향으로 질문을 정리합니다.",
        f"{locality}에서 세 과목을 함께 알아보는 학부모의 상담 장면을 가정했습니다. 최근 풀이를 보며 국어 근거 찾기, 영어 해석, 수학 풀이 단계 중 우선 점검할 부분을 구분해 듣는 상황입니다.",
    ])
    if center["schools"]:
        school_note = f"제공된 {SCHOOL_NAME} 참고 정보인 {'·'.join(center['schools'][:3])}{' 등' if len(center['schools']) > 3 else ''}의 목록과 학생의 실제 학교 자료를 대조해 질문하는"
    else:
        school_note = f"제공된 {SCHOOL_NAME} 목록이 없어 학생의 실제 {material_label}를 직접 준비해 질문하는"
    second = f"{locality} 기준 " + stable_pick(seed, "scenario-2", [
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
    url = page_path(slug)
    hub_url = page_path()
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
            {"@type": "ListItem", "position": 1, "name": "홈", "item": "/"},
            {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": "/과목별학원/"},
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
        "author": {"@id": org["@id"]}, "publisher": {"@id": org["@id"]}, "datePublished": DATE, "dateModified": DATE,
        "image": [f"/assets/representative/{rep_name}", f"/assets/centers/common/{center['body_image']}", f"/assets/maps/{center['map_name']}"],
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
        "primaryImageOfPage": {"@type": "ImageObject", "url": f"/assets/representative/{rep_name}"}, "about": about,
        "mainEntity": {"@id": article["@id"]},
    }
    faq_node = {"@type": "FAQPage", "@id": url + "#faq", "mainEntity": [{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in faq]}
    links = [
        (CATEGORY_NAME + " 전체 지역", hub_url), ("과목별학원", "/과목별학원/"),
        ("학습가이드", "/학습가이드/"), ("상담문의", "/상담문의/"),
    ]
    item_list = {"@type": "ItemList", "@id": url + "#related", "name": f"{title} 관련 페이지", "itemListElement": [{"@type": "ListItem", "position": i, "name": name, "url": href} for i, (name, href) in enumerate(links, 1)]}
    return {"@context": "https://schema.org", "@graph": [org, local_business, webpage, breadcrumb, article, service, faq_node, item_list]}


def header(prefix: str, current: str = "subjects") -> str:
    items = [("home", "홈", "index.html"), ("about", "학원소개", "학원소개/index.html"), ("guide", "학습가이드", "학습가이드/index.html"), ("contact", "상담문의", "상담문의/index.html"), ("subjects", "과목별학원", "과목별학원/index.html")]
    links = "".join(f'<a href="{prefix}{href}" data-nav="{key}"{" aria-current=\"page\"" if key == current else ""}>{label}</a>' for key, label, href in items)
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


def render_info(center: dict) -> str:
    rows = [
        ("지역", " ".join(part for part in (center["region"], center["district"], center["locality"]) if part)),
        ("센터 기준", center["center_name"]),
        ("제공 주소", center["address"]),
    ]
    html_rows = "".join(f"<div><dt>{esc(label)}</dt><dd>{esc(value)}</dd></div>" for label, value in rows if value)
    grades = "".join(f'<li><strong>{esc(subject)}</strong><span>{esc(value)}</span></li>' for subject, value in grade_summary(center))
    schools = "".join(f"<span>{esc(school)}</span>" for school in center["schools"])
    school_html = f'<div><dt>{SCHOOL_NAME} 참고</dt><dd><div class="local-tags">{schools}</div></dd></div>' if schools else f'<div><dt>{SCHOOL_NAME} 참고</dt><dd>제공 목록 없음 · 상담 시 실제 학교 자료 확인</dd></div>'
    tuition = f'<a class="button compact" href="{esc(center["tuition_url"])}" target="_blank" rel="noopener">센터별 교습비 확인 <span aria-hidden="true">↗</span></a>' if center["tuition_url"] else '<p class="info-note">교습비 자료는 희망 센터에서 확인합니다.</p>'
    return f'<dl class="local-facts">{html_rows}{school_html}</dl><ul class="grade-list">{grades}</ul>{tuition}'


def render_page(record: dict, previous_record: dict, next_record: dict) -> str:
    sections = record["sections"]
    title = record["title"]
    locality = record["locality"]
    slug = record["slug"]
    center = record["center"]
    meta = record["meta"]
    rep_name = record["rep_name"]
    student_type = verified_student_type(extract_student_type(sections["본문"], locality), center, slug)
    intro, body_sections = personalize_body(sections["본문"], title, locality, center, slug)
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
    faq = build_faq(title, locality, center, student_type, slug)
    scenarios = build_scenarios(locality, center, student_type, slug)
    quick_answer = direct_answer(locality, center, student_type, slug)
    answer_card_html = "".join(
        f'<article><span>{esc(label)}</span><strong>{esc(answer)}</strong></article>'
        for label, answer in answer_cards(locality, center, slug)
    )
    graph = build_graph(title, locality, slug, center, meta, faq, [heading for heading, _ in body_sections], rep_name)
    graph_json = json.dumps(graph, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    body_html = "".join(f'<section class="manuscript-section"><span class="section-kicker">{index:02d}</span><h2>{esc(heading)}</h2>{"".join(f"<p>{esc(paragraph)}</p>" for paragraph in paragraphs)}</section>' for index, (heading, paragraphs) in enumerate(body_sections, 1))
    faq_html = "".join(f'<details{" open" if index == 0 else ""}><summary>{esc(question)}</summary><p>{esc(answer)}</p></details>' for index, (question, answer) in enumerate(faq))
    scenario_html = "".join(f'<article class="scenario-card"><span>상담 상황 예시 {index:02d}</span><p>{esc(value)}</p></article>' for index, value in enumerate(scenarios, 1))
    body_image = center["body_image"]
    region_label = " ".join(part for part in (center["region"], center["district"], locality) if part)
    prev_link = f'<a class="local-nav-card" href="../{esc(previous_record["slug"])}/index.html"><small>이전 지역</small><strong>{esc(previous_record["title"])}</strong><span>←</span></a>'
    next_link = f'<a class="local-nav-card" href="../{esc(next_record["slug"])}/index.html"><small>다음 지역</small><strong>{esc(next_record["title"])}</strong><span>→</span></a>'
    return f'''<!doctype html>
<html lang="ko"><head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)} | {SITE_NAME}</title><meta name="description" content="{esc(meta)}">
  <meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec">
  <meta property="og:type" content="article"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="{esc(title)} | {SITE_NAME}"><meta property="og:description" content="{esc(meta)}"><meta property="og:image" content="/assets/representative/{esc(rep_name)}">
  <link rel="icon" href="../../../assets/favicon.png"><link rel="stylesheet" href="../../../assets/site14.css">
  <script type="application/ld+json">{graph_json}</script>
</head><body data-page="subjects">
  {header("../../../")}
  <main id="main">
    <section class="local-hero"><div class="site-shell">
      <nav class="breadcrumbs" aria-label="현재 위치"><a href="../../../index.html">홈</a><a href="../../index.html">과목별학원</a><a href="../index.html">{CATEGORY_NAME}</a><span>{esc(title)}</span></nav>
      <p class="eyebrow">{ENGLISH_LEVEL} Korean · English · Math</p><h1>{esc(title)}</h1><p class="local-lead">{esc(meta)}</p>
      <div class="local-answer-grid">{answer_card_html}</div>
    </div></section>
    <section class="section local-overview"><div class="site-shell local-overview-grid">
      <div class="local-summary"><p class="chapter-label"><span>01</span> Quick answer</p><h2>{esc(locality)}에서 먼저 확인할 내용</h2><p>{esc(quick_answer)}</p><div class="answer-note"><strong>대표 학생 상황</strong><p>{esc(student_type)}</p></div></div>
      <aside class="local-info-card"><p class="eyebrow">Center information</p><h2>수업·상담 확인 정보</h2>{render_info(center)}</aside>
    </div></section>
    <section class="local-media-section"><div class="site-shell local-media-stack">
      <img src="../../../assets/representative/{esc(rep_name)}" alt="{esc(title)} {SITE_NAME} 대표" style="display:none;">
      <figure class="local-body-image"><img src="../../../assets/centers/common/{body_image}" width="918" height="16116" alt="{esc(title)} 본문 {SITE_NAME}" loading="lazy" decoding="async"><figcaption>{esc(region_label)} {LEVEL_NAME}의 국어·영어·수학 학습 점검 안내</figcaption></figure>
      <figure class="local-map-image"><img src="../../../assets/maps/{esc(center['map_name'])}" alt="{esc(title)} 지도 {SITE_NAME}" loading="lazy" decoding="async"><figcaption>센터 위치는 제공된 주소 자료를 기준으로 표시하며 방문 전 실제 운영 여부를 확인합니다.</figcaption></figure>
    </div></section>
    <section class="section manuscript-wrap"><article class="site-shell manuscript-article"><div class="manuscript-intro"><span>원고 핵심 답변</span><p>{esc(intro)}</p></div>{body_html}</article></section>
    <section class="section blue-wash"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>02</span> Consultation examples</div><div><h2>{esc(locality)} 학부모 상담 상황 예시</h2><p>아래 내용은 실제 고객 후기나 성적 결과가 아니라, 상담에서 확인할 질문을 이해하기 위한 상황 예시입니다.</p></div></div><div class="scenario-grid">{scenario_html}</div></div></section>
    <section class="section"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>03</span> FAQ</div><div><h2>{esc(title)} 자주 묻는 질문</h2><p>학부모님이 상담 전에 자주 확인하는 내용을 질문과 답변으로 정리했습니다.</p></div></div><div class="faq-list">{faq_html}</div></div></section>
    <section class="section local-links-section"><div class="site-shell"><div class="section-heading compact-heading"><div class="chapter-label"><span>04</span> Continue</div><div><h2>{esc(locality)} 페이지 이동</h2><p>카테고리 전체 지역 또는 앞뒤 지역 안내로 이동할 수 있습니다.</p></div></div><div class="local-navigation"><a class="local-nav-card is-parent" href="../index.html"><small>카테고리</small><strong>{CATEGORY_NAME} 전체 지역</strong><span>↑</span></a>{prev_link}{next_link}</div></div></section>
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
    url = page_path()
    items = [{"@type": "ListItem", "position": index, "name": record["title"], "url": page_path(record["slug"])} for index, record in enumerate(records, 1)]
    return {"@context": "https://schema.org", "@graph": [
        {"@type": "EducationalOrganization", "@id": "/#organization", "name": SITE_NAME, "url": "/", "telephone": PHONE, "teaches": [f"{COURSE_NAME} 국어", f"{COURSE_NAME} 영어", f"{COURSE_NAME} 수학", "학습코칭"]},
        {"@type": "BreadcrumbList", "@id": url + "#breadcrumb", "itemListElement": [{"@type": "ListItem", "position": 1, "name": "홈", "item": "/"}, {"@type": "ListItem", "position": 2, "name": "과목별학원", "item": "/과목별학원/"}, {"@type": "ListItem", "position": 3, "name": CATEGORY_NAME, "item": url}]},
        {"@type": "CollectionPage", "@id": url + "#collection", "url": url, "name": f"{CATEGORY_NAME} 지역 안내", "description": f"371개 동네별 {LEVEL_NAME} 국어·영어·수학 학습 상담 정보와 센터 자료를 확인하는 지역 허브입니다.", "inLanguage": "ko-KR", "breadcrumb": {"@id": url + "#breadcrumb"}, "hasPart": [{"@type": "WebPage", "name": record["title"], "url": page_path(record["slug"])} for record in records]},
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
            cards = "".join(f'<a class="directory-card" href="{esc(record["slug"])}/index.html" data-locality="{esc(record["locality"])} {esc(record["title"])}"><strong>{esc(record["locality"])}</strong><span>{LEVEL_NAME} 국어·영어·수학 안내</span><i aria-hidden="true">→</i></a>' for record in values)
            districts_html.append(f'<section class="directory-district"><div class="directory-district-head"><h2>{esc(district)}</h2><span>{len(values)}개 지역</span></div><div class="directory-grid">{cards}</div></section>')
        count = sum(len(values) for values in grouped[region].values())
        blocks.append(f'<details class="directory-region" data-region="{esc(region)}"{" open" if r_index == 0 else ""}><summary><span><b>{esc(region)}</b><small>{len(grouped[region])}개 시군구 · {count}개 동네</small></span><i aria-hidden="true">+</i></summary><div class="directory-region-body">{"".join(districts_html)}</div></details>')
    graph_json = json.dumps(hub_graph(records), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    return f'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{CATEGORY_NAME} 지역 안내 | {SITE_NAME}</title><meta name="description" content="371개 동네별 {CATEGORY_NAME} 페이지에서 국어·영어·수학 진단, {SCHOOL_NAME} 자료, 가능 학년과 센터 상담 정보를 확인하세요."><meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec"><meta property="og:type" content="website"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="{CATEGORY_NAME} 지역 안내 | {SITE_NAME}"><meta property="og:description" content="광역지역과 시군구별로 371개 {CATEGORY_NAME} 안내를 찾을 수 있습니다."><link rel="icon" href="../../assets/favicon.png"><link rel="stylesheet" href="../../assets/site14.css"><script type="application/ld+json">{graph_json}</script></head>
<body data-page="subjects">{header("../../")}<main id="main"><section class="directory-hero"><div class="site-shell"><nav class="breadcrumbs" aria-label="현재 위치"><a href="../../index.html">홈</a><a href="../index.html">과목별학원</a><span>{CATEGORY_NAME}</span></nav><p class="eyebrow">National subject directory</p><h1>동네별 {CATEGORY_NAME}</h1><p>{HUB_INTRO}</p><div class="hub-metrics"><div><strong>371</strong><span>지역 페이지</span></div><div><strong>{HUB_FOCUS}</strong><span>{LEVEL_NAME} 핵심 기준</span></div><div><strong>{HUB_PROCESS}</strong><span>상담 확인 흐름</span></div></div></div></section>
<section class="section directory-section"><div class="site-shell"><div class="directory-toolbar"><label for="local-search">동네 이름으로 찾기</label><div class="directory-search"><input id="local-search" type="search" placeholder="예: 명일동, 불당동, 중계동" autocomplete="off" data-local-search><span data-directory-count>전체 371개 지역</span></div><div class="region-filters" aria-label="광역지역 선택">{region_buttons}</div><div class="directory-actions"><button type="button" data-expand-all>모두 펼치기</button><button type="button" data-collapse-all>모두 접기</button></div></div><div class="directory-empty" data-directory-empty hidden>검색 조건에 맞는 동네가 없습니다.</div><div class="directory-list">{"".join(blocks)}</div></div></section>
<section class="section ink"><div class="site-shell consult-cta"><div><h2>{HUB_CTA}</h2><p>{HUB_CTA_BODY}</p></div><a class="button orange" href="../../상담문의/index.html">상담 방법 확인 <span aria-hidden="true">→</span></a></div></section></main>{footer("../../")}<script src="../../assets/site14.js" defer></script></body></html>'''


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
    return f'''<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>과목별학원 | {SITE_NAME}</title><meta name="description" content="학년과 과목 조합에 맞춰 지역별 학습 상담 정보를 찾을 수 있도록 과목별학원 카테고리와 선택 기준을 정리했습니다."><meta name="robots" content="index,follow,max-image-preview:large"><meta name="theme-color" content="#f7f3ec"><meta property="og:type" content="website"><meta property="og:locale" content="ko_KR"><meta property="og:title" content="과목별학원 | {SITE_NAME}"><meta property="og:description" content="학년·과목별 학습 목표와 지역 센터 정보를 한 흐름에서 확인하세요."><link rel="icon" href="../assets/favicon.png"><link rel="stylesheet" href="../assets/site14.css"><script type="application/ld+json">{graph_json}</script></head><body data-page="subjects">{header("../")}<main id="main"><section class="directory-hero subjects-root-hero"><div class="site-shell"><nav class="breadcrumbs" aria-label="현재 위치"><a href="../index.html">홈</a><span>과목별학원</span></nav><p class="eyebrow">Subject academy guide</p><h1>과목별학원</h1><p>같은 학년이라도 과목 조합과 막히는 지점에 따라 확인할 수업 기록이 다릅니다. 학생에게 필요한 카테고리를 선택한 뒤 지역별 센터 정보와 학습 안내를 확인하세요.</p></div></section><section class="section"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>01</span> Category</div><div><h2>현재 확인할 수 있는 학원 안내</h2><p>검증된 원고와 센터정보를 기준으로 학년과 과목 조합을 구분했습니다.</p></div></div><div class="subject-hub-grid">{cards}</div></div></section><section class="section blue-wash"><div class="site-shell"><div class="section-heading"><div class="chapter-label"><span>02</span> How to choose</div><div><h2>카테고리를 고를 때 확인할 세 가지</h2><p>과목명만 보고 결정하지 않고 학생의 실제 자료와 센터 운영 범위를 함께 확인합니다.</p></div></div><div class="role-grid"><article class="role-card"><span class="icon">01</span><h3>학생의 현재 학년</h3><p>페이지 제목과 별개로 희망 센터에서 해당 학년과 과목을 실제로 운영하는지 확인합니다.</p></article><article class="role-card"><span class="icon">02</span><h3>최근 평가와 교재</h3><p>점수만 말하기보다 틀린 문제와 현재 교재를 준비해 과목별 병목을 구분합니다.</p></article><article class="role-card"><span class="icon">03</span><h3>주간 실행 가능성</h3><p>학교 일정, 귀가 시간과 복습 시간을 함께 놓고 무리하지 않는 과제량을 확인합니다.</p></article></div></div></section></main>{footer("../")}<script src="../assets/site14.js" defer></script></body></html>'''


def update_base_navigation() -> None:
    page_settings = [(ROOT / "index.html", ""), (ROOT / "학원소개" / "index.html", "../"), (ROOT / "학습가이드" / "index.html", "../"), (ROOT / "상담문의" / "index.html", "../")]
    for path, prefix in page_settings:
        source = path.read_text(encoding="utf-8")
        replacements = {
            'href="/"': f'href="{prefix}index.html"',
            'href="/학원소개/"': f'href="{prefix}학원소개/index.html"',
            'href="/학습가이드/"': f'href="{prefix}학습가이드/index.html"',
            'href="/상담문의/"': f'href="{prefix}상담문의/index.html"',
        }
        for old, new in replacements.items():
            source = source.replace(old, new)
        if 'data-nav="subjects"' not in source:
            source = re.sub(r'(<a [^>]*data-nav="contact"[^>]*>상담문의</a>)', rf'\1<a href="{prefix}과목별학원/index.html" data-nav="subjects">과목별학원</a>', source, count=1)
        footer_match = re.search(r'(<nav class="footer-links"[^>]*>)(.*?)(</nav>)', source, re.DOTALL)
        if footer_match and "과목별학원" not in footer_match.group(2):
            contents = footer_match.group(2) + f'<a href="{prefix}과목별학원/index.html">과목별학원</a>'
            source = source[:footer_match.start()] + footer_match.group(1) + contents + footer_match.group(3) + source[footer_match.end():]
        path.write_text(source, encoding="utf-8")


def main() -> None:
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
    CATEGORY_ROOT.mkdir(parents=True, exist_ok=True)
    for index, record in enumerate(records):
        previous_record = records[(index - 1) % len(records)]
        next_record = records[(index + 1) % len(records)]
        target = CATEGORY_ROOT / record["slug"] / "index.html"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(render_page(record, previous_record, next_record), encoding="utf-8")
    (CATEGORY_ROOT / "index.html").write_text(render_category_hub(records), encoding="utf-8")
    TARGET_ROOT.mkdir(parents=True, exist_ok=True)
    (TARGET_ROOT / "index.html").write_text(render_root_hub(), encoding="utf-8")
    update_base_navigation()
    from prepare_production_domain import main as prepare_production_domain
    prepare_production_domain()
    # Keep generated links on their final canonical URLs and rebuild social cards
    # from the representative/hero image after production-domain metadata is ready.
    from normalize_internal_links_and_social_meta import normalize_site
    normalization = normalize_site(ROOT, apply=True)
    print(json.dumps({"detail_pages": len(records), "category_hub": str(CATEGORY_ROOT / 'index.html'), "subject_hub": str(TARGET_ROOT / 'index.html'), "unique_representatives": len(set(representatives)), "normalization": normalization}, ensure_ascii=False))


if __name__ == "__main__":
    main()
