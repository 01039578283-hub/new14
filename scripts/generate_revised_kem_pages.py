from __future__ import annotations

import argparse
import hashlib
import html
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable
from zipfile import BadZipFile, ZipFile


ROOT = Path(__file__).resolve().parents[1]
SUBJECT_ROOT = ROOT / "\uacfc\ubaa9\ubcc4\ud559\uc6d0"
DEFAULT_ARCHIVE = Path.home() / "Desktop" / "1.zip"
DEFAULT_COMMON = ROOT.parent / "참고자료" / "공통자료"
BASE_HELPER = ROOT / "scripts" / "generate_highschool_korean_english_math.py"
SITE_ORIGIN = "https://xn--3e0bz50b1zcyxat54c.com"
REVISION_DATE = "2026-08-22"
BASE_COMMIT = "9e58f271f6126db72d4eb10a363c9d3b4d163779"
OFFICIAL_REGION_PATTERN = (
    r"(?:서울특별시|부산광역시|대구광역시|인천광역시|광주광역시|대전광역시|"
    r"울산광역시|세종특별자치시|경기도|강원특별자치도|충청북도|충청남도|"
    r"전북특별자치도|전라남도|경상북도|경상남도|제주특별자치도)"
)
ADMIN_PREFIX_PATTERN = (
    rf"(?:(?:{OFFICIAL_REGION_PATTERN})\s+(?:[가-힣]+(?:시|군|구)\s+)?|"
    rf"[가-힣]+(?:시|군|구)\s+(?:[가-힣]+(?:시|군|구)\s+)?)"
)

EXPECTED_ARCHIVE_SHA256 = "20c268ce05bd48c18c659a629ae522ad9470eac8dba07ebf4d435b57d6a1d57f"
EXPECTED_ARCHIVE_ENTRIES = {
    "초등학생학원 원고.xlsx": "deb8a99bd51b9d9f5792cf8149e10c4a3c38579af870985e466b039939ac834d",
    "중학생학원 원고.xlsx": "336376c4186a3a5fea2137a3d2bc9a28f7876e9cde6137762b66917ff6c63558",
    "고등학생학원 원고.xlsx": "738a76a553efdb7af4e5e6cbe5f22de1b00847cd7b3b659f55f623cb3537797c",
}
EXPECTED_BASE_HELPER_SHA256 = "36f3d788760470f3984430b4f30fb6c9c630ac1b55387ead1f05df7f01fb5881"
EXPECTED_SUBJECT_CATALOG_SHA256 = "9ae6ae6757b4a7766717822bcd6183920f1381c5d2f643eaab82c6ec39550814"
EXPECTED_CENTER_CSV_SHA256 = "3ffbd7b70273b6dc1c8435c53a3a25e32d2a173ba1bf51840654389bd8954e1a"
EXPECTED_IMAGE_CSV_SHA256 = "c1b4f87b2b62f659107dbf0a79a1d566e213e008fc4b7f30cfa656ffae814100"
EXPECTED_SUPERSEDED_AFTER_MANIFEST = "f8828e61788be26d9d72e5ed619b450d626f5f0964052686730c4b6f50c8f451"
EXPECTED_DETAIL_COUNT = 371
EXPECTED_DETAIL_DOCUMENTS = 1113
EXPECTED_DOCUMENT_COUNT = 1114

PROFILE_SPECS = (
    ("elementary", "초등학생학원 원고.xlsx", "초등학생국영수학원"),
    ("middle", "중학생학원 원고.xlsx", "중학생국영수학원"),
    ("high", "고등학생학원 원고.xlsx", "고등학생국영수학원"),
)

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class GateError(RuntimeError):
    pass


@dataclass(frozen=True)
class SourceCell:
    workbook: str
    profile: str
    row: int
    locality: str
    raw_html: str
    raw_sha256: str
    source_h1: str
    source_intro: str
    source_sections: tuple[tuple[str, tuple[str, ...]], ...]
    corrected_localities: tuple[str, ...]


@dataclass(frozen=True)
class Document:
    path: Path
    before: bytes
    after: bytes
    profile: str
    locality: str
    source_sha256: str


@dataclass
class BuildPlan:
    documents: list[Document]
    source_manifest: dict[str, str]
    metrics: dict[str, object]

    @property
    def changed(self) -> list[Document]:
        return [doc for doc in self.documents if doc.before != doc.after]

    def candidate_sha256(self) -> str:
        digest = hashlib.sha256()
        for doc in sorted(self.documents, key=lambda item: item.path.as_posix()):
            rel = doc.path.relative_to(ROOT).as_posix().encode("utf-8")
            digest.update(rel + b"\0" + hashlib.sha256(doc.after).digest())
        for key, value in sorted(self.source_manifest.items()):
            digest.update(key.encode("utf-8") + b"\0" + value.encode("ascii"))
        return digest.hexdigest()

    def freeze_payload(self) -> dict[str, object]:
        return {
            "version": 1,
            "generator_sha256": sha256_file(Path(__file__)),
            "candidate_sha256": self.candidate_sha256(),
            "source_manifest": self.source_manifest,
            "documents": [
                {
                    "path": doc.path.relative_to(ROOT).as_posix(),
                    "before_sha256": sha256_bytes(doc.before),
                    "after_sha256": sha256_bytes(doc.after),
                    "source_sha256": doc.source_sha256,
                }
                for doc in sorted(self.documents, key=lambda item: item.path.as_posix())
            ],
        }


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    return sha256_bytes(path.read_bytes())


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", html.unescape(str(value or ""))).strip()


def visible_text(source: str) -> str:
    source = re.sub(r"<(?:script|style)\b.*?</(?:script|style)>", " ", source, flags=re.I | re.S)
    return clean(re.sub(r"<[^>]+>", " ", source))


def stable_pick(seed: str, label: str, choices: tuple[str, ...]) -> str:
    value = hashlib.sha256(f"{seed}|{label}".encode("utf-8")).hexdigest()
    return choices[int(value[:12], 16) % len(choices)]


def run_git(*args: str, root: Path = ROOT) -> str:
    result = subprocess.run(
        ["git", *args], cwd=root, check=True, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, text=True, encoding="utf-8",
    )
    return result.stdout.strip()


def assert_hash(path: Path, expected: str, label: str) -> None:
    if not path.is_file():
        raise GateError(f"missing {label}: {path}")
    actual = sha256_file(path)
    if actual != expected:
        raise GateError(f"{label} hash changed: {actual} != {expected}")


def normalize_zip_name(name: str) -> str:
    # ZipFile already honors the UTF-8 flag. Refuse names that only become safe
    # after rewriting rather than attempting a permissive repair.
    return name.replace("\\", "/")


def validate_outer_member(name: str) -> None:
    normalized = normalize_zip_name(name)
    path = PurePosixPath(normalized)
    if path.is_absolute() or ".." in path.parts or len(path.parts) != 1:
        raise GateError(f"unsafe archive member: {name!r}")
    if not normalized.lower().endswith(".xlsx"):
        raise GateError(f"unexpected archive content: {name!r}")


def workbook_cells(raw: bytes, workbook_name: str) -> list[str]:
    try:
        with ZipFile(io.BytesIO(raw)) as book:
            infos = book.infolist()
            names = {normalize_zip_name(info.filename) for info in infos}
            if any(info.flag_bits & 1 for info in infos):
                raise GateError(f"encrypted workbook part: {workbook_name}")
            if any(PurePosixPath(name).is_absolute() or ".." in PurePosixPath(name).parts for name in names):
                raise GateError(f"unsafe OOXML path: {workbook_name}")
            forbidden = (
                "vbaProject.bin", "externalLinks/", "embeddings/", "activeX/",
                "customUI/", "oleObject",
            )
            if any(any(token.lower() in name.lower() for token in forbidden) for name in names):
                raise GateError(f"active or external OOXML content: {workbook_name}")
            required = {
                "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
                "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml",
            }
            if not required.issubset(names):
                raise GateError(f"missing OOXML parts in {workbook_name}: {sorted(required - names)}")
            if sum(info.file_size for info in infos) > 80_000_000:
                raise GateError(f"workbook expansion too large: {workbook_name}")

            workbook = ET.fromstring(book.read("xl/workbook.xml"))
            sheets = workbook.find(NS_MAIN + "sheets")
            # These are Power Query exports. The first sheet contains the
            # frozen results we consume; the second is an unused source sheet.
            # The internal $Workbook$ connection is never opened or executed.
            if (
                sheets is None or len(sheets) != 2
                or sheets[0].attrib.get("name") != "A열_텍스트파일"
                or sheets[1].attrib.get("name") != "Sheet1"
            ):
                raise GateError(f"unexpected sheet contract: {workbook_name}")

            shared_root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            shared = [
                "".join(node.text or "" for node in item.iter(NS_MAIN + "t"))
                for item in shared_root.findall(NS_MAIN + "si")
            ]
            sheet = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
            dimension = sheet.find(NS_MAIN + "dimension")
            if dimension is None or dimension.attrib.get("ref") != "A1:A372":
                raise GateError(f"unexpected worksheet dimension: {workbook_name}")
            second_sheet = ET.fromstring(book.read("xl/worksheets/sheet2.xml"))
            if sheet.findall(".//" + NS_MAIN + "f") or second_sheet.findall(".//" + NS_MAIN + "f"):
                raise GateError(f"formulas are not allowed: {workbook_name}")
            if (
                sheet.findall(".//" + NS_MAIN + "hyperlink")
                or sheet.findall(".//" + NS_MAIN + "mergeCell")
                or second_sheet.findall(".//" + NS_MAIN + "hyperlink")
                or second_sheet.findall(".//" + NS_MAIN + "mergeCell")
            ):
                raise GateError(f"hyperlinks/merged cells are not allowed: {workbook_name}")

            values: list[str] = []
            refs: list[str] = []
            for cell in sheet.findall(".//" + NS_MAIN + "c"):
                ref = cell.attrib.get("r", "")
                value = cell.find(NS_MAIN + "v")
                if value is None:
                    raise GateError(f"empty cell {ref}: {workbook_name}")
                if cell.attrib.get("t") != "s":
                    raise GateError(f"non-shared-string cell {ref}: {workbook_name}")
                index = int(value.text or "-1")
                if not 0 <= index < len(shared):
                    raise GateError(f"shared string index out of bounds: {workbook_name} {ref}")
                refs.append(ref)
                values.append(shared[index].replace("_x000D_", ""))
            expected_refs = [f"A{row}" for row in range(1, 373)]
            if refs != expected_refs or values[0] != "사용자 지정":
                raise GateError(f"row/header contract mismatch: {workbook_name}")
            manuscripts = values[1:]
            if len(manuscripts) != EXPECTED_DETAIL_COUNT or len(set(manuscripts)) != EXPECTED_DETAIL_COUNT:
                raise GateError(f"manuscript cardinality/uniqueness: {workbook_name}")
            return manuscripts
    except (BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
        raise GateError(f"invalid workbook {workbook_name}: {exc}") from exc


def read_archive(path: Path) -> tuple[dict[str, list[str]], dict[str, str]]:
    assert_hash(path, EXPECTED_ARCHIVE_SHA256, "attached archive")
    workbooks: dict[str, list[str]] = {}
    manifest = {"archive": EXPECTED_ARCHIVE_SHA256}
    try:
        with ZipFile(path) as outer:
            infos = outer.infolist()
            if len(infos) != len(EXPECTED_ARCHIVE_ENTRIES):
                raise GateError(f"archive member count={len(infos)}")
            for info in infos:
                validate_outer_member(info.filename)
                if info.flag_bits & 1 or info.file_size <= 0 or info.file_size > 8_000_000:
                    raise GateError(f"unsafe archive member flags/size: {info.filename}")
                name = normalize_zip_name(info.filename)
                if name not in EXPECTED_ARCHIVE_ENTRIES:
                    raise GateError(f"unexpected workbook: {name}")
                raw = outer.read(info)
                actual = sha256_bytes(raw)
                if actual != EXPECTED_ARCHIVE_ENTRIES[name]:
                    raise GateError(f"workbook hash mismatch {name}: {actual}")
                workbooks[name] = workbook_cells(raw, name)
                manifest[f"workbook:{name}"] = actual
    except BadZipFile as exc:
        raise GateError(f"invalid outer ZIP: {exc}") from exc
    if set(workbooks) != set(EXPECTED_ARCHIVE_ENTRIES):
        raise GateError("workbook name set mismatch")
    return workbooks, manifest


def load_base(profile: str, common_dir: Path):
    helper_sha = sha256_file(BASE_HELPER)
    if helper_sha != EXPECTED_BASE_HELPER_SHA256:
        raise GateError(f"base helper hash changed: {helper_sha}")
    scripts_dir = str(BASE_HELPER.parent)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    old_argv = sys.argv[:]
    name = f"_revised_kem_base_{profile}_{uuid.uuid4().hex}"
    try:
        sys.argv = [str(BASE_HELPER), profile]
        spec = importlib.util.spec_from_file_location(name, BASE_HELPER)
        if spec is None or spec.loader is None:
            raise GateError("could not load base generator")
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
    finally:
        sys.argv = old_argv
        sys.modules.pop(name, None)
    module.ROOT = ROOT
    module.COMMON = common_dir
    module.CENTER_CSV = common_dir / "센터정보 정리.csv"
    module.IMAGE_CSV = common_dir / "이미지링크.csv"
    module.REP_SOURCE = common_dir / "대표이미지"
    module.REP_TARGET = ROOT / "assets" / "representative"
    module.MAP_DIR = ROOT / "assets" / "maps"
    module.TARGET_ROOT = ROOT / "과목별학원"
    module.CATEGORY_ROOT = module.TARGET_ROOT / module.CATEGORY_SLUG
    module.DATE_MODIFIED = REVISION_DATE
    return module


def extract_one(pattern: str, source: str) -> str:
    match = re.search(pattern, source, flags=re.I | re.S)
    return match.group(1).strip() if match else ""


def source_fragment(raw: str) -> tuple[str, str, list[tuple[str, list[str]]]]:
    if any(token in raw.lower() for token in ("<script", "<style", "<iframe", "javascript:", "onerror=", "onclick=")):
        raise GateError("active markup in manuscript cell")
    if raw.count('<main class="article-main">') != 1 or raw.count("</main>") != 1:
        raise GateError("unexpected article-main wrapper")
    h1 = clean(extract_one(r"<h1\b[^>]*>(.*?)</h1>", raw))
    intro = visible_text(extract_one(r'<p\s+class="article-intro"[^>]*>(.*?)</p>', raw))
    if not h1 or not intro:
        raise GateError("missing source H1 or article intro")
    sections: list[tuple[str, list[str]]] = []
    for section in re.findall(r'<section\s+class="article-section[^\"]*"[^>]*>(.*?)</section>', raw, flags=re.I | re.S):
        heading = visible_text(extract_one(r"<h2\b[^>]*>(.*?)</h2>", section))
        if not heading:
            continue
        paragraphs: list[str] = []
        cards = re.findall(r"<article\b[^>]*>(.*?)</article>", section, flags=re.I | re.S)
        for card in cards:
            lead = visible_text(extract_one(r"<(?:strong|h3)\b[^>]*>(.*?)</(?:strong|h3)>", card))
            pieces = [
                visible_text(item)
                for item in re.findall(r"<(?:p|li)\b[^>]*>(.*?)</(?:p|li)>", card, flags=re.I | re.S)
            ]
            pieces = [piece for piece in pieces if piece]
            if lead:
                lead = lead if lead.endswith((".", "?", "!")) else lead + "."
                pieces.insert(0, lead)
            paragraph = clean(" ".join(pieces))
            if paragraph:
                paragraphs.append(paragraph)
        if not paragraphs:
            paragraphs = [
                visible_text(item)
                for item in re.findall(r"<p\b[^>]*>(.*?)</p>", section, flags=re.I | re.S)
                if visible_text(item)
            ]
        if paragraphs:
            sections.append((heading, paragraphs))
    closing = visible_text(extract_one(r'<section\s+class="article-closing"[^>]*>(.*?)</section>', raw))
    if closing and sections:
        sections[-1][1].append(closing)
    if not 3 <= len(sections) <= 6:
        raise GateError(f"source section count={len(sections)}")
    return h1, intro, sections


HEADING_BANKS = {
    "focus": (
        "현재 자료에서 먼저 확인할 학습 신호",
        "과목 계획을 세우기 전에 볼 기준",
        "처음 상담에서 구분할 학습 과제",
        "지금의 공부 흐름을 읽는 기준",
        "학습 순서를 정할 때 확인할 내용",
    ),
    "teacher": (
        "설명을 듣고 풀이로 확인하는 과정",
        "이해와 실행을 함께 살피는 수업 기준",
        "질문과 기록으로 확인하는 학습 과정",
        "수업 안팎의 학습 태도를 살피는 방법",
        "풀이 과정과 복습을 연결하는 기준",
    ),
    "process": (
        "진단에서 재확인까지의 학습 순서",
        "수업 전후 기록을 이어 가는 방법",
        "개념 확인부터 오답 복습까지",
        "학습 계획을 실제 실행으로 옮기는 순서",
        "현재 수준에 맞춘 점검 단계",
    ),
    "grade": (
        "학년 단계에 맞춘 준비 기준",
        "교과 단계별로 나눠 보는 학습 과제",
        "학년 변화에 따라 달라지는 점검 항목",
        "현재 학년에 맞춘 과목별 준비",
        "다음 학습 단계로 이어지는 확인 기준",
    ),
    "other": (
        "상담 전에 함께 살펴볼 내용",
        "자료와 기록으로 확인할 기준",
        "과목별 계획을 조정하는 방법",
        "학습 흐름을 구체화하는 질문",
        "다음 계획에 반영할 점검 내용",
    ),
}


def classify_heading(value: str) -> str:
    value = clean(value)
    if any(token in value for token in ("학년", "전략", "학습 방향")):
        return "grade"
    if any(token in value for token in ("진행", "방식", "순서", "프로세스")):
        return "process"
    if any(token in value for token in ("선생", "특징", "지도")):
        return "teacher"
    if any(token in value for token in ("핵심", "포인트", "진단", "중점")):
        return "focus"
    return "other"


PHRASE_REWRITES = (
    ("성적 상승을 보장하지", "성적이나 진학 결과를 보장하지"),
    ("성적 향상을 보장하지", "성적이나 진학 결과를 보장하지"),
    ("성적이 오르게 돕습니다", "학습 변화가 기록에 남는지 확인합니다"),
    ("성적이 오르게 만듭니다", "복습이 이어지는지 확인합니다"),
    ("성적이 오르는 구조", "복습이 이어지는 구조"),
    ("성적이 오르게 만드는", "복습이 이어지게 하는"),
    ("성적이 오르는", "오답 원인이 줄어드는"),
    ("성적으로 연결되도록", "다음 풀이에 반영되도록"),
    ("시험에서 점수로 연결되도록", "시험 문제에 적용되는지 확인하도록"),
    ("점수로 연결되도록", "평가 문제에 적용되는지 확인하도록"),
    ("오답을 점수로 바꾸는", "오답 원인을 다음 풀이에 반영하는"),
    ("꾸준히 점수를 끌어올릴 수 있도록", "반복 오류가 줄어드는지 확인하도록"),
    ("꾸준히 점수를 끌어올리는", "반복 오류를 줄이는"),
    ("꾸준히 점수를 끌어올립니다", "반복 오류가 줄어드는지 확인합니다"),
    ("실전 점수를 끌어올립니다", "실전에서 반복되는 오류가 줄어드는지 확인합니다"),
    ("최종 점수를 끌어올립니다", "마지막 점검에서 반복 오류를 줄입니다"),
    ("점수를 끌어올릴 수 있도록", "반복 오류가 줄어드는지 확인하도록"),
    ("점수를 끌어올리면서", "반복 오류를 줄이면서"),
    ("점수를 끌어올리는", "반복 오류를 줄이는"),
    ("점수를 끌어올릴", "반복 오류를 줄일"),
    ("점수를 끌어올립니다", "반복 오류가 줄어드는지 확인합니다"),
    ("성적을 끌어올리는", "학습 변화를 기록하는"),
    ("성적을 끌어올릴", "학습 변화를 확인할"),
    ("성적을 끌어올립니다", "학습 변화를 기록합니다"),
    ("성적을 올리기보다", "점수만 비교하기보다"),
    ("성적을 올리기 위해서는", "학습 변화를 확인하려면"),
    ("성적을 올리기 위해", "학습 변화를 확인하려면"),
    ("성적을 올리고 싶은", "반복 오류를 줄이고 싶은"),
    ("성적을 올리고 싶다면", "반복 오류를 줄이려면"),
    ("성적을 올리고", "학습 변화를 기록하고"),
    ("성적을 올리는", "반복 오류를 줄이는"),
    ("성적을 올릴", "반복 오류를 줄일"),
    ("성적을 올립니다", "학습 변화가 기록에 남는지 확인합니다"),
    ("성적을 높이는", "학습 변화를 확인하는"),
    ("성적을 높일", "학습 변화를 확인할"),
    ("성적을 높입니다", "학습 변화가 기록에 남는지 확인합니다"),
    ("흔들리지 않는 실력 상승을 돕습니다", "과목별 학습 변화를 기록으로 확인합니다"),
    ("실력 상승을 돕습니다", "학습 변화를 기록으로 확인합니다"),
    ("실력이 누적되도록 코칭합니다", "학습 기록이 이어지는지 점검합니다"),
    ("실력이 누적되도록", "학습 기록이 이어지는지 확인하도록"),
    ("꾸준히 성장할 수 있게 돕습니다", "다음 점검 시점을 함께 정합니다"),
    ("성장할 수 있게", "학습 변화를 기록할 수 있게"),
    ("꾸준히 실력이 쌓이도록", "학습 기록이 이어지도록"),
    ("실력이 쌓이도록", "학습 기록이 이어지도록"),
    ("흔들리지 않게 만듭니다", "변화를 기록으로 확인합니다"),
    ("실력 상승", "학습 변화"),
    ("영어·수학 결과 관리", "영어·수학 학습 기록 확인"),
    ("시험에서 점수로 연결되는", "평가 문제에 적용되는"),
    ("점수로 연결되는", "평가 문제에 적용되는"),
    ("과목 간 실력 편차를 줄입니다", "과목별 이해 차이를 확인합니다"),
    ("학교 시험 범위와 출제 경향을 반영해", "학교 시험 범위표와 제공된 자료를 확인해"),
    ("출제 경향에 맞춘", "기존 시험지에서 확인한 유형을 기준으로 한"),
    ("출제 경향", "제공된 시험지의 문항 유형"),
    ("자주 나오는 유형", "반복해 틀린 유형"),
    ("점수 향상", "점수 변화"),
    ("점수 상승", "점수 변화"),
    ("성적 상승", "학습 변화"),
    ("성적 향상", "학습 변화"),
    ("확실하게", "구체적으로"),
    ("확실히", "구체적으로"),
    ("반드시", "우선"),
    ("최고의", "현재 상황에 맞는"),
    ("개별 맞춤", "학생별"),
    ("맞춤형", "현재 수준에 맞춘"),
    ("맞춤 학습", "현재 수준에 맞춘 학습"),
    ("체계적으로 지도하며", "학습 단계를 나눠 확인하며"),
    ("체계적으로 지도합니다", "학습 단계를 나눠 점검합니다"),
    ("정확히 파악합니다", "구체적으로 살펴봅니다"),
    ("정확하게 파악합니다", "자료를 기준으로 살펴봅니다"),
    ("실력을 끌어올립니다", "학습 변화가 기록에 남는지 확인합니다"),
    ("성취도를 끌어올립니다", "성취도 변화를 단원별로 확인합니다"),
    ("점수를 안정적으로 올립니다", "점수 변화의 근거를 오답 기록에서 확인합니다"),
    ("학습 효율을 높입니다", "과목별 시간 배분을 조정합니다"),
    ("효율을 높입니다", "시간 배분을 조정합니다"),
    ("빠르게 보완하고", "차근차근 보완하고"),
    ("빠르게 보완합니다", "차근차근 보완합니다"),
    ("상위 난도까지 확장합니다", "다음 난도 진입 시점을 확인합니다"),
    ("상위권 실력으로 확장합니다", "다음 단계에 필요한 조건을 확인합니다"),
    ("스스로 풀 수 있게 지도합니다", "혼자 다시 풀 수 있는지를 확인합니다"),
    ("스스로 풀 수 있도록 지도합니다", "혼자 다시 풀 수 있는지를 확인합니다"),
    ("실력이 유지되도록 관리합니다", "복습이 이어지는지를 기록으로 확인합니다"),
    ("꾸준히 실력이 쌓이도록 돕습니다", "학습 기록이 다음 주까지 이어지도록 점검합니다"),
    ("실력이 쌓이도록 돕습니다", "학습 기록이 이어지도록 점검합니다"),
    ("완성합니다", "확인합니다"),
    ("완성할 수 있습니다", "준비할 수 있습니다"),
    ("최적의", "현실적인"),
    ("무조건", "바로"),
)


DIRECT_QUESTION = "내신형 문제는 어떻게 준비해야 해요? 내신 출제"

OUTCOME_NOUN_PATTERN = (
    r"(?:성적|점수|실력|정확도|성취도|성취감|성취\s+경험|성취|성과|결과|성장|학습 변화|점수 변화"
    r"|수업(?:의)? 효과|학습(?:의)? 효과)"
)
ACADEMIC_ABILITY_PATTERN = (
    r"(?:독해력|문해력|사고력|응용력|문제\s*해결력|계산력|적용력|학습력|집중력|"
    r"어휘력|표현력|평가\s*대응력|실전\s*감각|기초\s*체력|자신감|이해력|해석력|"
    r"판단력|서술력|논리력|추론력|풀이력|기본기|기초력|학습\s*역량|과목\s*역량|"
    r"국어\s*역량|영어\s*역량|수학\s*역량|학업\s*역량|평가\s*역량|대응력|적응력|"
    r"풀이\s*역량|서술\s*역량|읽기\s*역량|해석\s*능력|풀이\s*능력|학습\s*능력)"
)
SAFE_NOMINAL_OUTCOME_RE = re.compile(
    r"(?:실력\s+향상을\s+(?:위한|위해)|실력\s+향상에\s+필요한|실력\s+향상\s+(?:계획|로드맵)"
    r"|성취도\s+향상\s+목표|점수대\s+상승을\s+목표로|개선\s+방향"
    r"|(?:학습|오답|복습)\s+기록|목표\s+성적\s+달성을\s+위한"
    r"|강점\s+확장|훈련\s+강화|쌓아가고\s+싶다면|성장\s+(?:방향|목표)"
    r"|성장을\s+목표로|성장하는\s+방향|(?:지속\s+)?성장\s+로드맵"
    r"|성취\s+목표(?:\([^)]*\))?|성취\s+중심)"
)
ABILITY_CROSS_OBJECT_RE = re.compile(
    ACADEMIC_ABILITY_PATTERN
    + r"(?:이|가|을|를|은|는|도|의|과|와)?\s*(?:통해|기준으로|바탕으로|근거로)"
)
ABILITY_PLANNING_CROSS_OBJECT_RE = re.compile(
    ACADEMIC_ABILITY_PATTERN
    + r"(?:이|가|을|를|은|는|도|의|과|와)?[^.!?]{0,25}?"
    + r"(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
    + r"(?:만들|확장|강화|완성)"
)
OUTCOME_CROSS_OBJECT_RE = re.compile(
    r"(?:성취도|성취)(?:만|를)?\s*(?:보지\s+않|통해|기준으로|바탕으로|근거로)"
    r"|(?:과정|기록|자료|내용|단계|흐름|습관|태도)(?:이|가)[^.!?]{0,12}(?:쌓|누적)"
)
RESIDUAL_OUTCOME_PATTERN = (
    r"(?:성적|점수|실력|정확도|성취도|성취|성과|학습\s*결과|결과|학습\s*변화|점수\s*변화)"
)
RESIDUAL_ABILITY_PATTERN = (
    ACADEMIC_ABILITY_PATTERN + r"|스스로\s*공부하는\s*힘|공부하는\s*힘"
)
RESIDUAL_SAFE_OUTCOME_RE = re.compile(
    r"결과\s*예측보다[^.!?]{0,45}(?:현재\s*상태|실행\s*계획)[^.!?]{0,18}바꾸는\s*과정"
    r"|진단\s*결과가\s*과목별\s*다음\s*행동으로\s*이어지는지"
    r"|(?:점수|성과)보다[^.!?]{0,55}(?:행동|풀이\s*과정|개념\s*연결)"
    r"|(?:향상|상승|개선|강화|확장|달성)(?:을|를|에|의)?\s*"
    r"(?:위해|위한|목표|로드맵|전략|방향)"
    r"|(?:결과|성취도?)(?:를|가)?[^.!?]{0,18}(?:대조|바탕|근거|점검|확인)"
    r"|(?:줄이도록|최소화하도록)[^.!?]{0,24}(?:점검|계획|전략을\s*세)"
    r"|학습\s*변화를\s*확인(?:할\s*수\s*있는|하는)\s*"
    r"(?:흐름|과정|구조|루틴)(?:을|를)\s*만들"
)
RESIDUAL_DIRECT_OUTCOME_RES = (
    re.compile(
        r"(?:꾸준함|오답|약점|성적)[^.!?]{0,35}(?:성적|실력|점수)[’”\"']?\s*"
        r"(?:으?로)?[^.!?]{0,12}(?:바꾸|전환)"
    ),
    re.compile(
        RESIDUAL_OUTCOME_PATTERN
        + r"(?:이|가|을|를|은|는|도|에|로|으로)?[^.!?]{0,25}"
        + r"(?:나게|나도록|나오게|나오도록|따라오|반영되|달라지|좋아지|회복|도달)"
    ),
    re.compile(r"(?:실력|학습\s*변화)(?:이|가)?[^.!?]{0,12}나오(?:게|도록)"),
    re.compile(
        r"(?:학습\s*변화|성적|점수\s*변화)[’”]?\s*(?:에|으로|로)\s*"
        r"(?:바로\s*|직접\s*)?(?:연결|이어지)"
        r"|(?:오답\s*노트|공부한\s*시간|공부\s*습관|공부|실력)[^.!?]{0,28}"
        r"(?:다음\s*점수|성적|학습\s*변화|다음\s*성과)[^.!?]{0,10}(?:이어지|연결)"
    ),
    re.compile(
        RESIDUAL_OUTCOME_PATTERN
        + r"(?:의|이|가|을|를|은|는|도)?[^.!?]{0,22}(?:변동|편차|격차|손실|요인)"
        + r"[^.!?]{0,14}(?:최소화|줄이)"
    ),
    re.compile(
        r"(?:개념|기초|학습|진도|내용)\s*누락[^.!?]{0,35}(?:최소화|줄이|줄여|막|방지|없애)"
        r"|(?:개념\s*누락|풀이\s*실수|독해\s*오류)[^.!?]{0,40}재발[^.!?]{0,20}(?:막|정착)"
    ),
    re.compile(
        r"(?:" + RESIDUAL_ABILITY_PATTERN + r")"
        r"(?:이|가|을|를|은|는|도|의|과|와)?[^.!?]{0,28}"
        r"(?:기르|길러|키우|회복|보완|다지|잡|만들|향상|높이|높여|올리|올려|"
        r"강화|확장|성장|쌓|누적|안정|정착|완성|확보)"
    ),
    re.compile(
        r"성장\s*(?:루틴|흐름|과정|기반|동력|습관)[^.!?]{0,30}(?:만들|정착|완성|구축|잡)"
    ),
    re.compile(
        RESIDUAL_OUTCOME_PATTERN
        + r"(?:이|가|을|를|은|는|도|의|에|으?로|까지)?[^.!?]{0,28}(?:만들|만듭)"
    ),
    re.compile(
        r"(?:국어\s*)?실력(?:도|을|은|는|까지|의\s*기반을)[^.!?]{0,18}(?:잡|다지)"
        r"|(?:기초\s*)?실력\s*다지기|실력까지\s*쌓|성취[^.!?]{0,16}유지"
        r"|(?:실력\s*격차|점수\s*변동\s*요인)[^.!?]{0,12}줄이"
    ),
)
AUDIT_SAFE_NOMINAL_SPANS = re.compile(
    r"(?:목표\s+성적\s+달성을\s+위한|"
    r"(?:향상|상승|개선|강화|확장)(?:을|를|에|은|이|과|와|의|까지)?\s*"
    r"(?:위해|위한|목표로|목표|계획|전략|방법|기준|과정|필요|중요|로드맵|방향|피드백)|"
    r"훈련\s+강화|성장\s+(?:방향|목표)|성장을\s+목표로|성장하는\s+방향|"
    r"(?:지속\s+)?성장\s+로드맵|성취\s+목표(?:\([^)]*\))?|성취\s+중심|"
    r"쌓아가고\s+싶다면)"
)
AUDIT_RESIDUAL_PATTERNS = {
    "score_conversion": re.compile(
        r"(?:오답|약점|실수|오류|학습\s*결과)(?:이|가|을|를)[^.!?]{0,45}?"
        r"[‘“\"']?(?:다음\s+)?(?:시험\s+)?점수[’”\"']?\s*로[^.!?]{0,10}?"
        r"(?:바꿉니다|바꾸(?:는|도록|게|어|며|고)|바꿔(?:집니다|지도록|주는|줍니다)?|"
        r"바뀝니다|바뀌(?:는|도록|게|어|며|고)|바뀐)"
        r"(?:\s+(?:방법|기준|과정|전략|질문|표현|피드백|관리))?"
    ),
    "outcome_transform": re.compile(
        r"(?:꾸준함|오답|약점|성적)(?:이|가|을|를|은|는)?[^.!?]{0,35}?"
        r"(?:성적|실력|점수)(?:이|가|을|를|은|는|으?로)?[^.!?]{0,12}?"
        r"(?:바꿉니다|바꾸(?:는|도록|게|어|며|고)|바뀝니다|바뀌(?:는|도록|게|어)|"
        r"전환(?:합니다|됩니다|시키|하도록|되도록|하는))"
    ),
    "outcome_appearance": re.compile(
        r"(?:성적|점수|실력|성과|결과|학습\s+변화)(?:이|가|에|은|는|도)?[^.!?]{0,25}?"
        r"(?:나게|나도록|나오게|나오도록|따라옵니다|따라오(?:게|도록|는|며|고)|"
        r"반영되(?:게|도록|는|며|고|었습니다|ㅂ니다)|"
        r"회복(?:합니다|됩니다|시키|하게|하도록|되도록|하는)|"
        r"도달(?:합니다|하게|하도록|하는))"
    ),
    "outcome_minimize": re.compile(
        r"(?:성적|점수|실력)(?:의|에서)?\s*(?:격차|차이|편차|변동|변화|하락|손실|흔들림)"
        r"(?:을|를|은|는)?[^.!?]{0,8}?최소화(?:합니다|됩니다|시키|하도록|되도록|하는)"
    ),
    "outcome_ability_finite": re.compile(
        r"(?:스스로\s+공부하는\s+힘|독해력|문해력|사고력|응용력|문제\s*해결력|계산력|"
        r"적용력|학습력|집중력|어휘력|표현력|평가\s*대응력|실전\s*감각|기초\s*체력|"
        r"자신감|이해력|해석력|판단력|서술력|논리력|추론력|풀이력|기본기|기초력|"
        r"학습\s*역량|과목\s*역량|국어\s*역량|영어\s*역량|수학\s*역량|학업\s*역량|"
        r"평가\s*역량|대응력|적응력|풀이\s*역량|서술\s*역량|읽기\s*역량|해석\s*능력|"
        r"풀이\s*능력|학습\s*능력)(?:이|가|을|를|은|는|도|과|와)?[^.!?]{0,18}?"
        r"(?:길러(?:냅니다|줍니다|지도록|가는|가도록)|기릅니다|키웁니다|키우(?:도록|게|는|며|고)|"
        r"회복(?:합니다|됩니다|시키|하도록|되도록|하는)|보완(?:합니다|되도록|하는)|"
        r"다집니다|다지(?:도록|게|는|며|고)|잡습니다|잡아(?:줍니다|가도록|주는)|잡도록)"
    ),
    "outcome_growth_flow_make": re.compile(
        r"(?:성장|학습\s+변화)[^.!?]{0,16}?(?:루틴|흐름)(?:이|가|을|를|은|는)?"
        r"[^.!?]{0,8}?(?:만듭니다|만들(?:어|도록|게|며|고|는))"
    ),
    "outcome_change_make": re.compile(
        r"(?:성적|점수|실력|성과|결과|학습\s+변화|점수\s+변화)"
        r"(?:이|가|을|를|은|는|도|으?로)?[^.!?]{0,20}?"
        r"(?:만듭니다|만들(?:어|도록|게|며|고|는))"
    ),
    "outcome_skill_build": re.compile(
        r"실력(?:이|가|을|를|은|는|도)?[^.!?]{0,14}?"
        r"(?:잡습니다|잡아(?:줍니다|가도록|주는)|잡도록|다집니다|다지(?:도록|게|는|며|고)|"
        r"쌓습니다|쌓이(?:도록|게|는)|쌓(?:도록|게|는))"
    ),
    "outcome_variance_reduce": re.compile(
        r"(?:격차|변동)(?:이|가|을|를|은|는|도)?[^.!?]{0,14}?"
        r"(?:줄입니다|줄이(?:도록|게|는|며|고)|줄어(?:들도록|듭니다|드는|들게))"
    ),
}
AUDIT_RESIDUAL_EXCLUSIONS = {
    "score_conversion": re.compile(r"바꾸(?:는|기\s+위한)\s+(?:방법|기준|과정|전략|질문|표현)"),
    "outcome_transform": re.compile(r"바꾸(?:는|기\s+위한)\s+(?:방법|기준|과정|전략|질문|표현)"),
    "outcome_appearance": re.compile(
        r"(?:기록|계획|상담|다음\s+풀이|학습\s+순서|자료|피드백)(?:에|으로)"
        r"[^.!?]{0,4}반영되|(?:나오|따라오|반영되|회복|도달)[^.!?]{0,12}"
        r"(?:는지|여부|원인|기준|확인|점검|분석)"
    ),
    "outcome_ability_finite": re.compile(
        r"(?:력|역량|능력|감각|체력|자신감|기본기)(?:을|를)\s*"
        r"(?:통해|기준으로|바탕으로|근거로)|"
        r"(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:만들|확장|강화|완성|잡|다지)"
    ),
    "outcome_change_make": re.compile(
        r"(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:만들|만듭|확장|강화|완성)"
    ),
    "outcome_skill_build": re.compile(
        r"(?:(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:잡|다지|쌓)|(?:쌓(?:는|을\s+수\s+있는)|다지는|잡는)\s+"
        r"(?:과정|연습|계획|전략|방법|기준|방향))"
    ),
}
DIRECT_OUTCOME_RES = (
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|과|와|도|은|는|의|로|으로)[^.!?]{0,48}?"
        + r"(?:향상|상승|오르|올라|올립|올려|끌어올|키우|키웁|높입|높여|개선|확보"
        + r"|만듭|나타나|나오는|쌓|누적|성장|안정|달라집|바뀌|일어나|확장|강화)"
    ),
    re.compile(
        r"(?:성적|점수|실력|정확도|성취도|성취|성과|결과|학습 변화|점수 변화)\s*"
        r"(?:향상|상승|개선|성장|누적|유지)"
    ),
    re.compile(r"오답[^.!?]{0,28}?(?:점수|성적)[^.!?]{0,20}?(?:바뀌|연결)"),
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|과|와|도|은|는|의)?[^.!?]{0,36}?"
        + r"(?:만듭|끌어올|오르|올라|키웁|높이|높여|확보|올리|나오는|나타나)"
    ),
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|과|와|도|은|는|의)?[^.!?]{0,36}?"
        + r"(?:향상(?:시킵|되|하도록|합니다|할\s+수|을\s+돕)"
        + r"|상승(?:되|하도록|합니다|시킵|할\s+수)"
        + r"|개선(?:합니다|되도록|시키|할\s+수|을\s+돕))"
    ),
    re.compile(
        OUTCOME_NOUN_PATTERN + r"(?:이|가|로|으로)[^.!?]{0,12}?이어지"
    ),
    re.compile(
        r"(?:성적|점수|실력|성취도|정확도|성취)"
        r"(?:이|가|을|를|과|와|도|은|는)[^.!?]{0,40}?"
        r"(?:"
        r"성장(?:하도록|되도록|하게|되게|합니다|됩니다|시킵니다|돕습니다|지도합니다|관리합니다|코칭합니다)"
        r"|쌓(?:이도록|도록|이게|게|입니다|습니다)"
        r"|누적(?:하도록|되도록|하게|되게|합니다|됩니다|시킵니다|돕습니다|지도합니다|관리합니다|코칭합니다)"
        r")"
    ),
    re.compile(r"(?:성적|점수)[^.!?]{0,32}?(?:하락을?\s+방지|떨어지지\s+않|유지하도록)"),
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|과|와|도|은|는|의|로|으로)?[^.!?]{0,36}?"
        + r"(?:높입|높여|올립|올려|안정화|안정되|달라집|바뀌)"
    ),
    re.compile(r"(?:학습 변화|점수 변화)(?:이|가|을|를)?[^.!?]{0,24}?일어나"),
    re.compile(
        r"(?:성적|점수|실력|성취도|정확도|성취)"
        r"(?:이|가|을|를|과|와|도|은|는)?[^.!?]{0,40}?"
        r"(?:쌓을\s+수\s+있도록|쌓아(?:갑니다|드립니다)"
        r"|쌓는[^.!?]{0,18}?(?:코칭|수업)[^.!?]{0,14}?제공"
        r"|누적[\"“”'‘’]*(?:합니다|시킵|되는\s+(?:방식|학습))"
        r"|성장(?:시켜|할\s+수\s+있도록|하도록))"
    ),
    re.compile(
        r"(?:국어|영어|수학|아이|학생|학습 흐름)(?:이|가|은|는|을|를|와|과|으로)[^.!?]{0,32}?"
        r"성장(?:시켜|할\s+수\s+있도록|하도록|하게|되도록)"
    ),
    re.compile(r"(?:점수|학습 변화|성과|점수 변화|성적|결과)(?:로|으로)[^.!?]{0,18}?(?:연결|이어지)"),
    re.compile(r"성장(?:시켜|시키|하도록|하게|되도록|을\s+돕|을\s+지원|을\s+관리|할\s+수\s+있도록)"),
    re.compile(r"(?:실력|성적|성과)(?:이|가|을|를|은|는)?[^.!?]{0,28}?(?:유지|안정)"),
    re.compile(r"(?:성적|점수|성과|결과)[^.!?]{0,28}?(?:만들|만듭|완성|달성)"),
    re.compile(r"직결"),
    re.compile(r"(?:실력|성적|점수|성과|정확도|역량)(?:이|가|을|를|은|는|로|으로)?[^.!?]{0,24}?(?:강화|확장|극대화)"),
    re.compile(ACADEMIC_ABILITY_PATTERN + r"(?:이|가|을|를|은|는|로|으로)?[^.!?]{0,24}?(?:강화|확장|극대화)"),
    re.compile(r"학습(?:이|가)[^.!?]{0,16}?누적(?:되도록|되게|됩니다|시키|합니다)"),
    re.compile(r"(?:성적|점수)[^.!?]{0,24}?(?:하락을?\s+(?:막|방지)|떨어지지\s+않)"),
    # Targeted high-recall families from the independent sentence audit.  The
    # predicates stay close to their academic subject, avoiding broad rewrites
    # of neutral process prose.
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|은|는|도|로|으로|까지)?\s*(?:바로\s+|곧\s+|다음\s+)?(?:이어지|연결)"
    ),
    re.compile(
        r"(?:성장을\s*(?:돕|지원|이끌|관리|지도|코칭)|"
        r"성장(?:하도록|할\s+수\s+있도록|하게|되게|합니다|됩니다|시킵|시켜))"
    ),
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|은|는|도|로|으로)?\s*(?:바로\s+|다음\s+|꾸준히\s+|"
        + r"안정적인\s+|안정적으로\s+|결국\s+|함께\s+)?(?:만들|만듭|완성|확보|달성|나오|나타나)"
    ),
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|은|는|도|의|과|와|로|으로)?[^.!?]{0,35}?"
        + r"(?:향상|상승|오르|올라|올리|올립|올려|끌어올|높이|높입|높여|키우|키웁|"
        + r"개선|확장|강화|극대화)(?:합니다|됩니다|시킵니다|시키는|시키며|하도록|"
        + r"되도록|되게|할\s+수|해|하는|되는)"
    ),
    re.compile(
        OUTCOME_NOUN_PATTERN
        + r"(?:이|가|을|를|은|는|도|의|과|와|로|으로|까지)?[^.!?]{0,20}?"
        + r"(?:쌓|누적)(?:습니다|합니다|됩니다|시킵니다|시키는|시키며|되도록|되게|"
        + r"되는|하도록|하게|을\s+수|아가|이게|이도록)"
    ),
    re.compile(
        r"(?:성적|점수|실력|성과|성취도|학습\s+변화)"
        r"(?:이|가|을|를|은|는|도|의|과|와|로|으로|까지)?[^.!?]{0,30}?"
        r"(?:안정(?:됩니다|되도록|되게|시키|화)|유지(?:됩니다|되도록|하도록|하게|합니다))"
    ),
    re.compile(
        r"(?:성취도|성취감|성취\s+경험|성취)(?:이|가|은|는|을|를|과|와|도|의)?"
        r"[^.!?]{0,60}?(?:향상|상승|높입|높이|올립|올리|끌어올|개선|확보|강화|확장|"
        r"성장|쌓|누적|만듭|달성)"
    ),
    # Academic-ability outcomes often avoid generic score nouns.
    re.compile(
        r"(?:" + ACADEMIC_ABILITY_PATTERN + r"|독해\s+속도|풀이\s+정확도)"
        r"(?:이|가|을|를|과|와|도|은|는|의|로|으로)?[^.!?]{0,25}?"
        r"(?:향상|상승|오르|올라|올리|올립|올려|끌어올|높이|높입|높여|키우|키웁|"
        r"개선|확보|만들|만듭|쌓|누적|성장|안정|확장|강화|완성|극대화)"
    ),
    # Short noun-only promotional/service fragments are not complete prose.
    # Replacing the whole sentence with a deterministic process statement is
    # safer than guessing a missing predicate.
    re.compile(
        r"^\s*[^.!?]{0,56}(?:실력|점수|성적|성과|성취|성장|학습)"
        r"[^.!?]{0,28}(?:연결|완성|관리|설계|진단|코칭)\s*[.!?]?\s*$"
    ),
    re.compile(
        r"^\s*[^.!?]{0,56}" + ACADEMIC_ABILITY_PATTERN
        + r"[^.!?]{0,28}(?:완성|연결|강화|확장|상승|향상|성장|코칭|관리|지도|설계|대비)\s*[.!?]?\s*$"
    ),
)


def has_direct_outcome_claim(value: str) -> bool:
    scrubbed = SAFE_NOMINAL_OUTCOME_RE.sub("", value)
    scrubbed = ABILITY_CROSS_OBJECT_RE.sub("", scrubbed)
    scrubbed = ABILITY_PLANNING_CROSS_OBJECT_RE.sub("", scrubbed)
    scrubbed = OUTCOME_CROSS_OBJECT_RE.sub("", scrubbed)
    if any(pattern.search(scrubbed) for pattern in DIRECT_OUTCOME_RES):
        return True
    audit_sentence = AUDIT_SAFE_NOMINAL_SPANS.sub(
        lambda match: " " * len(match.group(0)), value
    )
    for label, pattern in AUDIT_RESIDUAL_PATTERNS.items():
        matches = tuple(pattern.finditer(audit_sentence))
        if not matches:
            continue
        exclusion = AUDIT_RESIDUAL_EXCLUSIONS.get(label)
        if exclusion is not None and all(
            exclusion.search(audit_sentence[match.start():match.end() + 48])
            for match in matches
        ):
            continue
        return True
    residual = RESIDUAL_SAFE_OUTCOME_RE.sub("", scrubbed)
    return any(pattern.search(residual) for pattern in RESIDUAL_DIRECT_OUTCOME_RES)


def repair_corrupted_question_marks(value: str) -> str:
    marker = "\ue000DIRECT-QUESTION\ue001"
    text = value.replace(DIRECT_QUESTION, DIRECT_QUESTION.replace("?", marker))
    text = re.sub(
        r"내신\s*\?{2,}\s*실력을\s*만듭니다",
        "내신 범위와 오답 기록을 대조해 다음 학습 계획을 정합니다",
        text,
    )
    text = text.replace("의정부? 가능동", "의정부 가능동")
    text = re.sub(r"(?<=지)\s*\?\s*(?=[가-힣])", ", ", text)
    parallel_pairs = (
        ("어휘", "문장"), ("해석", "문법"), ("찾기", "표현"),
        ("이해", "짧은"), ("이해", "문단"), ("찾기", "해석"),
        ("해석", "정답"), ("구조", "해석"), ("해석", "근거"),
        ("설명", "확인"), ("확인", "적용"),
        ("문법(개념)", "독해(적용)"), ("독해(적용)", "서술형/수능형"),
    )
    for left, right in parallel_pairs:
        text = re.sub(rf"{re.escape(left)}\s*\?\s*{re.escape(right)}", f"{left}·{right}", text)
    # The remaining mojibake separators join ordered learning stages in the
    # attached source (diagnosis -> concept -> practice -> review).
    text = re.sub(r"(?<=[가-힣)])\s*\?\s*(?=[가-힣(])", "→", text)
    return text.replace(marker, "?")


def remove_unmatched_smart_quotes(value: str) -> str:
    chars = list(value)
    for opening, closing in (("‘", "’"), ("“", "”")):
        stack: list[int] = []
        for index, char in enumerate(chars):
            if char == opening:
                stack.append(index)
            elif char == closing:
                if stack:
                    stack.pop()
                else:
                    chars[index] = ""
        for index in stack:
            chars[index] = ""
    return "".join(chars)


def soften_outcome_claims(value: str, seed: str = "", locality: str = "") -> str:
    text = value
    noun = r"(?:(?:내신|시험|학업)\s+)?(?:점수|성적|실력)"
    modifier = r"(?:(?:빠르게|안정적으로|꾸준히|효율적으로|단계적으로|확실하게|확실히|함께)\s+)*"
    quote = r"[\"“”'‘’]*"
    subject_replacements = (
        (rf"{noun}(?:이|가)\s*{quote}{modifier}(?:올라갈\s+수\s+있도록|오를\s+수\s+있도록|올라가도록|오르도록){quote}", "학습 변화가 기록에 남도록"),
        (rf"{noun}(?:이|가)\s*{quote}{modifier}(?:올라가는|오르는){quote}", "학습 변화를 확인하는"),
        (rf"{noun}(?:이|가)\s*{quote}{modifier}오르게{quote}", "학습 변화가 기록에 남게"),
    )
    for pattern, replacement in subject_replacements:
        text = re.sub(pattern, replacement, text)
    broad_subject = (
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라갈\s+수\s+있도록|오를\s+수\s+있도록|올라가도록|오르도록)", "학습 변화가 기록에 남도록"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라가는|오르는)", "학습 변화를 확인하는"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?오르게", "학습 변화가 기록에 남게"),
    )
    for pattern, replacement in broad_subject:
        text = re.sub(pattern, replacement, text)
    fallback_subject = (
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라갑니다|오릅니다|올랐습니다)", "학습 변화가 기록에 남는지 확인합니다"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라가고|오르고)", "학습 변화를 기록하고"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라가며|오르며)", "학습 변화를 기록하며"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라갈|오를)", "학습 변화를 확인할"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라가기|오르기)", "학습 변화를 확인하기"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라간|오른)", "학습 변화를 확인한"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라가면|오르면)", "학습 변화를 확인하면"),
        (rf"{noun}(?:이|가)[^.!?]{{0,30}}?(?:올라가지|오르지)", "학습 변화가 기록되지"),
    )
    for pattern, replacement in fallback_subject:
        text = re.sub(pattern, replacement, text)
    text = text.replace("성적이 ‘오르는 과정’", "학습 변화가 기록되는 과정")
    replacements = (
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올리도록\s*돕습니다", "복습 결과를 다음 학습 계획에 반영합니다"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올려\s*드립니다", "학습 변화를 기록으로 안내합니다"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올립니다", "오답 기록의 변화를 확인합니다"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올리는", "오답 원인을 점검하는"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올리고", "학습 변화를 기록하고"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올리며", "학습 변화를 기록하며"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올릴", "학습 변화를 확인할"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올리기", "학습 변화를 확인하기"),
        (rf"{noun}(?:을|를)?\s*{modifier}(?:끌어)?올려", "반복 오류를 줄여"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    broad_object = (
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올리도록\s*돕습니다", "복습 결과를 다음 학습 계획에 반영합니다"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올려\s*드립니다", "학습 변화를 기록으로 안내합니다"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올립니다", "오답 기록의 변화를 확인합니다"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올리는", "오답 원인을 점검하는"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올리고", "학습 변화를 기록하고"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올리며", "학습 변화를 기록하며"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올리기", "학습 변화를 확인하기"),
        (rf"{noun}(?:을|를)[^.!?]{{0,30}}?(?:끌어)?올릴", "학습 변화를 확인할"),
    )
    for pattern, replacement in broad_object:
        text = re.sub(pattern, replacement, text)
    # Some reused sentences join the outcome noun to a second object (for
    # example, "성적과 자신감을 ... 끌어올립니다") instead of using 을/를.
    # Apply a final bounded clause rewrite that mirrors the release gate while
    # retaining the surrounding sentence.
    fallback_object = (
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올리도록\s*돕습니다", "학습 과정이 이어지도록 점검합니다"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올려\s*드립니다", "학습 변화를 기록으로 안내합니다"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올립니다", "학습 변화가 기록에 남는지 확인합니다"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올리는", "학습 과정을 점검하는"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올리고", "학습 변화를 기록하고"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올리며", "학습 변화를 기록하며"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올리기", "학습 변화를 확인하기"),
        (rf"{noun}(?:을|를)?[^.!?]{{0,30}}?(?:끌어)?올릴", "학습 변화를 확인할"),
    )
    for pattern, replacement in fallback_object:
        text = re.sub(pattern, replacement, text)
    # A small set of promotional source sentences uses less regular endings
    # such as "올리도록 지도합니다".  Replace only the residual sentence and
    # choose among deterministic process-focused alternatives so the cleanup
    # does not introduce a new repeated sentence across hundreds of pages.
    evidence_options = (
        "최근 풀이 기록을", "오답 원인을", "복습 이력을", "과제 이행 기록을",
        "단원별 이해도를", "다시 풀기 결과를", "진단 결과를", "학습 시간 배분을",
    )
    action_options = (
        "다음 보완 순서를 정합니다", "복습 시점을 정합니다",
        "학습 상태를 구체적으로 확인합니다", "다음 학습 계획을 조정합니다",
        "반복 오류를 점검합니다", "설명할 수 있는 범위를 확인합니다",
        "보완할 단원을 구분합니다", "상담에서 확인할 기준을 정리합니다",
    )
    prefix_options = (
        "학습 과정에서는", "학습 점검에서는", "수업 계획에서는", "복습 계획에서는",
        "과목별 점검에서는", "다음 학습 단계에서는", "학습 순서를 정할 때는", "수업 내용을 점검할 때는",
    )
    segments = re.split(r"(?<=[.!?])(\s+)", text)
    for index in range(0, len(segments), 2):
        sentence = segments[index]
        if not has_direct_outcome_claim(sentence):
            continue
        # Preserve explicit cautions that reject guarantees; they are useful
        # consumer guidance rather than an outcome promise.
        if (
            ("보장" in sentence and any(token in sentence for token in ("않", "아니", "없", "말보다")))
            or "약속보다" in sentence or "약속이 아니라" in sentence
        ):
            continue
        punctuation = sentence[-1] if sentence.endswith((".", "!", "?")) else ""
        digest = hashlib.sha256(f"{seed}|{index}|{sentence}".encode("utf-8")).digest()
        evidence = evidence_options[digest[0] % len(evidence_options)]
        action = action_options[digest[1] % len(action_options)]
        lead = prefix_options[digest[2] % len(prefix_options)]
        prefix = f"{locality} {lead} " if locality else lead + " "
        segments[index] = f"{prefix}{evidence} 바탕으로 {action}" + punctuation
    text = "".join(segments)
    text = text.replace("학습 변화을", "학습 변화를").replace("점수 변화을", "점수 변화를")
    return text


def collapse_adjacent_duplicates(value: str) -> str:
    text = (
        value.replace("유형 유형", "유형")
        .replace("진도 진도표", "진도표")
        .replace("진도 진도만", "진도만")
    )
    pattern = re.compile(r"(?<![가-힣0-9])([가-힣][가-힣0-9]{1,24})\s+\1(?![가-힣0-9])")
    for _ in range(3):
        revised = pattern.sub(r"\1", text)
        if revised == text:
            break
        text = revised
    # Reused rows also contain middle-dot duplicates.  Lexical boundaries
    # intentionally preserve legitimate compounds such as "비문학·문학".
    middle_dot_same = re.compile(r"(?<![가-힣])([가-힣]{2,})·\1(?![가-힣])")
    middle_dot_suffix = re.compile(r"(?<![가-힣])([가-힣]{2,})·\1(학습|형)(?![가-힣])")
    for _ in range(3):
        revised = middle_dot_suffix.sub(r"\1\2", middle_dot_same.sub(r"\1", text))
        if revised == text:
            break
        text = revised
    return text


def complete_short_copy(value: str, locality: str) -> str:
    """Turn inherited label-like paragraphs into complete, useful prose."""
    text = clean(value).replace("영수(영어+수학)", "영어·수학").replace("영수 통합", "영어·수학 통합")
    if len(text) >= 35:
        return text
    text = re.sub(r"^\d+\.\s*", "", text)
    if not text.endswith((".", "!", "?")):
        text += "."
    finite = re.search(
        r"(?:합니다|됩니다|입니다|있습니다|없습니다|봅니다|정합니다|확인합니다|점검합니다|"
        r"진행합니다|관리합니다|지도합니다|코칭합니다|돕습니다|중요합니다|필요합니다)\.$",
        text,
    )
    if not finite:
        stem = text[:-1].strip()
        if stem.endswith("훈련"):
            text = stem + " 순서를 정리합니다."
        elif stem.endswith("관리"):
            text = stem + " 기준을 정리합니다."
        elif stem.endswith("습관화"):
            text = stem + " 과정을 점검합니다."
        elif stem.endswith("습관"):
            text = stem + "을 점검합니다."
        elif stem.endswith(("코칭", "지도")):
            text = stem + "에서 확인할 항목을 정리합니다."
        elif stem.endswith("제공"):
            text = stem + " 범위는 상담에서 확인합니다."
        elif stem == "영어·수학":
            text = stem + " 학습 범위를 확인합니다."
        else:
            text = stem + " 내용을 구체적으로 확인합니다."
    if len(text) < 35:
        text += f" {locality} 상담에서는 관련 자료와 확인 항목을 함께 정리합니다."
    return clean(text)


def revise_copy(
    value: str,
    locality: str,
    all_localities: set[str],
    seed: str,
    profile: str = "",
    replace_center: bool = True,
    expected_area: str = "",
    admin_replacements: dict[str, str] | None = None,
) -> str:
    text = clean(value)
    text = re.sub(
        r"[^.]*의정부\?가 아니라,\s*진짜 요청하신 지역은 [“\"]상대원동[”\"]입니다\.\s*",
        "",
        text,
    )
    for token, replacement in sorted((admin_replacements or {}).items(), key=lambda item: len(item[0]), reverse=True):
        text = re.sub(
            rf"(?<![가-힣]){re.escape(token)}(?![가-힣])",
            replacement,
            text,
        )
    # The workbooks are reused copy. Only the authoritative row locality is
    # retained; accidental branch/locality carry-over is not promoted.
    for foreign in sorted(all_localities - {locality}, key=len, reverse=True):
        if foreign in locality or locality in foreign:
            continue
        if foreign and foreign in text:
            # Avoid corrupting ordinary words and compound place names: 교하
            # must not match 정교하게, and 송동 must not match 삼송동.
            # A real locality mention starts at a lexical boundary; Korean
            # particles after the locality are intentionally allowed.
            text = re.sub(
                rf"(?<![가-힣0-9]){re.escape(foreign)}"
                r"(?=(?:에서|의|은|는|이|가|을|를|과|와|으로|로|도|만|부터|까지|처럼|보다|에|[,.;:!?]|\s|$))",
                locality,
                text,
            )
    if expected_area:
        administrative_locality = re.compile(
            ADMIN_PREFIX_PATTERN
            + rf"(?:{re.escape(locality)}|[가-힣0-9]+?(?:동|읍|면|리|가|지구|신도시))"
            r"(?=(?:에서|의|은|는|이|가|을|를|과|와|으로|로|,|\.|\s|$))"
        )
        text = administrative_locality.sub(expected_area, text)
        text = text.replace(f"{expected_area}에서 {locality}", f"{expected_area}에서")
    if replace_center:
        text = re.sub(r"와와학습코칭(?:센터|학원)(?:\s+[가-힣A-Za-z0-9]+점)?", "지역 센터", text)
        text = (
            text.replace("지역 센터이 ", "지역 센터가 ")
            .replace("지역 센터은 ", "지역 센터는 ")
            .replace("지역 센터을 ", "지역 센터를 ")
            .replace("지역 센터과 ", "지역 센터와 ")
            .replace("지역 센터으로 ", "지역 센터로 ")
        )
    text = text.replace("LOCAL ACADEMY GUIDE", "")
    for before, after in PHRASE_REWRITES:
        text = text.replace(before, after)
    text = repair_corrupted_question_marks(text)
    text = soften_outcome_claims(text, seed, locality)
    text = re.sub(r"\b영수\(영어·수학\)", "영어·수학", text)
    text = re.sub(r"(?:국어·)?영수\s+학습", "국어·영어·수학 학습", text)
    text = text.replace("이 글에서는", "상담에서는").replace("이 글은", "이 안내는").replace("이 글", "이 안내")
    text = text.replace("검색되더라도", "표현되어도").replace("검색되는", "사용되는")
    text = text.replace("검색할 때", "찾을 때").replace("검색어", "찾는 표현")
    text = text.replace("키워드", "확인 표현").replace("페이지", "안내").replace("검색", "찾기")
    text = text.replace("까지 고려한", "도 함께 살핀")
    text = (
        text.replace("체계적으로", "단계별로")
        .replace("설계해드립니다", "설계합니다")
        .replace("잡아드립니다", "점검합니다")
        .replace("지도해드립니다", "지도합니다")
        .replace("자세히 안내드립니다", "자세히 정리합니다")
        .replace("안내드립니다", "정리합니다")
    )
    text = (
        text.replace("을 설계 기준을 정리", "의 설계 기준을 정리")
        .replace("를 설계 기준을 정리", "의 설계 기준을 정리")
        .replace("학습 변화이", "학습 변화가")
        .replace("점수 변화이", "점수 변화가")
        .replace("학습 변화으로", "학습 변화로")
        .replace("점수 변화으로", "점수 변화로")
        .replace("학습 변화과", "학습 변화와")
        .replace("점수 변화과", "점수 변화와")
        .replace("학습 변화은", "학습 변화는")
        .replace("점수 변화은", "점수 변화는")
        .replace("교과 평가이", "교과 평가가")
        .replace("교과 평가을", "교과 평가를")
        .replace("교과 평가은", "교과 평가는")
        .replace("교과 평가과", "교과 평가와")
        .replace("교과 평가으로", "교과 평가로")
        .replace("현재 수준에 맞춘으로", "현재 수준에 맞춰")
        .replace("내신 제공된 시험지", "제공된 내신 시험지")
        .replace("시험 제공된 시험지", "제공된 시험지")
        .replace("내신 제공된 시험지의 문항 유형", "제공된 내신 시험지의 문항 유형")
        .replace("현재 성적과 시험 제공된 시험지의 문항 유형", "현재 성적과 제공된 시험지의 문항 유형")
        .replace("영어·수학 영약(영역)별", "영어·수학 영역별")
        .replace("영어·수학 영약별", "영어·수학 영역별")
        .replace("유형 학습- 오답 관리", "유형 학습→오답 관리")
        .replace("유형 학습- 문제풀이", "유형 학습→문제풀이")
        .replace("영어·수학·국어·영어·수학", "국어·영어·수학")
        .replace("영어/수학/국어/영어·수학", "국어·영어·수학")
        .replace("국어/국어·영어·수학", "국어·영어·수학")
        .replace("국어와 국어·영어·수학", "국어·영어·수학")
        .replace("국어·영어·수학를", "국어·영어·수학을")
        .replace("영어·수학를", "영어·수학을")
        .replace("최적화된 로드맵", "현재 기록을 반영한 학습 계획")
        .replace("최적화된", "현재 기록을 반영한")
        .replace("성취 향상.", "성취 과정을 점검합니다.")
        .replace("성장 속도에 맞춘 지도.", "학습 속도에 맞춘 점검 기준을 정리합니다.")
        .replace("성장 체크로 동기 강화.", "학습 과정의 점검 기준과 동기 유지 방법을 확인합니다.")
        .replace("현재 수준에 맞춘 현재 수준에 맞춘", "현재 수준에 맞춘")
        .replace("과정”", "과정")
        .replace("과정’", "과정")
        .replace("흐름”", "흐름")
        .replace("국어·영어·수학 계열)를", "국어·영어·수학 계열을")
        .replace("국어는 핵심 독해와 문장력 향상으로 연결합니다.", "국어는 핵심 독해와 문장 표현 과정을 함께 점검합니다.")
        .replace("영어는 독해 정확도, 수학은 유형 전환 능력을 오답으로 관리합니다.", "영어는 독해 정확도를, 수학은 유형 전환 과정의 오답을 점검합니다.")
        .replace("국어는 문제 해결력 기반으로 보완합니다.", "국어는 문제 해결 과정에서 보완할 부분을 점검합니다.")
        .replace("‘학습 변화’으로", "‘학습 변화’로")
        .replace("‘시간은 했는데 학습 변화를 확인하는’ 상황", "‘시간은 들였지만 학습 변화를 확인하기 어려운’ 상황")
        .replace("유형 적응을 빠르게 합니다.", "유형별 적응 과정을 점검합니다.")
        .replace("국어·학습 코칭", "국어 학습과 코칭")
        .replace("영수(영수)", "영수")
    )
    text = re.sub(
        r"목표에\s+맞(?:춘|는)\s+현재\s+수준에\s+맞춘",
        "목표와 현재 수준을 반영한",
        text,
    )
    text = re.sub(
        r"영어·수학(?:은\s+물론|과)?\s+국어·영어·수학",
        "국어·영어·수학",
        text,
    )
    text = re.sub(
        r"(?:영어·수학(?:은\s+물론|뿐\s+아니라|과|\s+중심의)?|영어,\s*수학(?:뿐\s+아니라)?)"
        r"[^.!?]{0,28}?국어·영어·수학",
        "국어·영어·수학",
        text,
    )
    text = re.sub(
        r"영어·수학·국어[^.!?]{0,35}?국어·영어·수학",
        "국어·영어·수학",
        text,
    )
    text = re.sub(r"유형\s+학습\s*-\s*", "유형 학습→", text)
    if profile == "elementary":
        text = re.sub(r"(?<![가-힣0-9])내신", "교과 평가", text)
        text = re.sub(r"(?<![가-힣0-9])입시", "상급 학년 준비", text)
    text = re.sub(rf"(?:{re.escape(locality)}\s+){{2,}}", locality + " ", text)
    locality_parts = locality.split()
    if len(locality_parts) > 1:
        text = text.replace(f"{locality_parts[0]} {locality}", locality)
        text = text.replace(f"{locality_parts[0]}시 {locality}", locality)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    text = re.sub(r"([.!?])(?=[가-힣A-Za-z0-9])", r"\1 ", text)
    text = collapse_adjacent_duplicates(text)
    text = remove_unmatched_smart_quotes(text)
    text = clean(text)
    if text.endswith(")"):
        text += "."
    return complete_short_copy(text, locality)


def center_area(center: dict, locality: str) -> str:
    display = " ".join(
        part for part in (
            clean(center.get("display_region", "")),
            clean(center.get("display_district", "")),
            clean(center.get("display_locality", locality)),
        ) if part
    )
    if display:
        return display
    address = clean(center.get("address", ""))
    position = address.find(locality)
    if position < 0 and locality.split():
        terminal = locality.split()[-1]
        position = address.find(terminal)
        if position >= 0:
            return address[:position + len(terminal)].strip()
    if position < 0:
        return locality
    return address[:position + len(locality)].strip()


def administrative_replacements(centers: dict[str, dict], center: dict) -> dict[str, str]:
    allowed_regions = {clean(center.get("region", "")), clean(center.get("display_region", ""))}
    allowed_districts = {clean(center.get("district", "")), clean(center.get("display_district", ""))}
    expected_region = clean(center.get("display_region", ""))
    expected_district = clean(center.get("display_district", "")) or expected_region
    replacements: dict[str, str] = {}
    for item in centers.values():
        for token in {clean(item.get("region", "")), clean(item.get("display_region", ""))} - allowed_regions - {""}:
            replacements[token] = expected_region
        for token in {clean(item.get("district", "")), clean(item.get("display_district", ""))} - allowed_districts - {""}:
            replacements[token] = expected_district
    return replacements


def sentence_excerpt(value: str, *, max_sentences: int = 2, max_chars: int = 260) -> str:
    pieces = [clean(item) for item in re.split(r"(?<=[.!?])\s+", clean(value)) if clean(item)]
    selected: list[str] = []
    for piece in pieces:
        candidate = " ".join([*selected, piece])
        if selected and len(candidate) > max_chars:
            break
        selected.append(piece)
        if len(selected) >= max_sentences:
            break
    return " ".join(selected) if selected else clean(value)


_BASELINE_BLOBS: dict[str, bytes] = {}


def prepare_baseline_blobs(paths: Iterable[Path]) -> None:
    rels = sorted({path.relative_to(ROOT).as_posix() for path in paths})
    missing = [rel for rel in rels if rel not in _BASELINE_BLOBS]
    if not missing:
        return
    request = "".join(f"{BASE_COMMIT}:{rel}\n" for rel in missing).encode("utf-8")
    result = subprocess.run(
        ["git", "cat-file", "--batch"], cwd=ROOT, input=request,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True,
    )
    payload = result.stdout
    offset = 0
    for rel in missing:
        header_end = payload.find(b"\n", offset)
        if header_end < 0:
            raise GateError(f"truncated baseline blob header: {rel}")
        header = payload[offset:header_end]
        offset = header_end + 1
        if header.endswith(b" missing"):
            raise GateError(f"baseline blob missing: {rel}")
        parts = header.rsplit(b" ", 2)
        if len(parts) != 3 or parts[1] != b"blob":
            raise GateError(f"unexpected baseline blob header: {rel}: {header!r}")
        size = int(parts[2])
        end = offset + size
        if end >= len(payload) or payload[end:end + 1] != b"\n":
            raise GateError(f"truncated baseline blob body: {rel}")
        _BASELINE_BLOBS[rel] = payload[offset:end]
        offset = end + 1
    if offset != len(payload):
        raise GateError("unexpected trailing bytes from git cat-file")


def current_manuscript(path: Path) -> tuple[str, list[tuple[str, list[str]]]]:
    rel = path.relative_to(ROOT).as_posix()
    if rel not in _BASELINE_BLOBS:
        prepare_baseline_blobs([path])
    source = _BASELINE_BLOBS[rel].decode("utf-8")
    wrapper = extract_one(
        r'<section\s+class="section manuscript-wrap"[^>]*>\s*<article[^>]*>(.*?)</article>\s*</section>',
        source,
    )
    intro = visible_text(extract_one(r'<div\s+class="manuscript-intro"[^>]*>.*?<p>(.*?)</p>\s*</div>', wrapper))
    sections: list[tuple[str, list[str]]] = []
    for block in re.findall(r'<section\s+class="manuscript-section"[^>]*>(.*?)</section>', wrapper, flags=re.I | re.S):
        heading = visible_text(extract_one(r"<h2\b[^>]*>(.*?)</h2>", block))
        paragraphs = [visible_text(item) for item in re.findall(r"<p\b[^>]*>(.*?)</p>", block, flags=re.I | re.S)]
        paragraphs = [item for item in paragraphs if item]
        if heading and paragraphs:
            sections.append((heading, paragraphs))
    if not intro or not sections:
        raise GateError(f"could not parse current manuscript: {path}")
    return intro, sections


def token_set(value: str) -> set[str]:
    return {token for token in re.findall(r"[가-힣A-Za-z0-9]{2,}", clean(value).lower()) if token not in {"학습", "학생", "수업", "확인"}}


def select_context(
    source_text: str,
    current_sections: list[tuple[str, list[str]]],
    used: set[tuple[int, int]],
) -> str:
    source_tokens = token_set(source_text)
    candidates: list[tuple[float, int, int, str]] = []
    for section_index, (heading, paragraphs) in enumerate(current_sections):
        if "학교" in heading or "센터" in heading:
            continue
        for paragraph_index, paragraph in enumerate(paragraphs):
            key = (section_index, paragraph_index)
            if key in used or not 70 <= len(paragraph) <= 360:
                continue
            tokens = token_set(paragraph)
            score = len(source_tokens & tokens) / max(1, len(source_tokens | tokens))
            candidates.append((score, section_index, paragraph_index, paragraph))
    if not candidates:
        return ""
    _, section_index, paragraph_index, paragraph = max(candidates, key=lambda item: (item[0], -item[1], -item[2]))
    used.add((section_index, paragraph_index))
    return paragraph


def source_operational_claim(value: str) -> bool:
    text = clean(value)
    if any(token in text for token in ("등록번호", "교습비", "수업료", "주차", "차량", "셔틀", "운영시간")):
        return True
    if any(token in text for token in ("주소", "방문 위치", "센터 위치")) and re.search(r"\d", text):
        return True
    return False


def deduplicate_page_sentences(value: str, seen: set[str]) -> str:
    """Drop only byte-equivalent sentences already used on the same page."""
    kept: list[str] = []
    for sentence in re.split(r"(?<=[.!?])\s+", clean(value)):
        normalized = clean(sentence)
        process_core = re.search(
            r"(?:최근 풀이 기록을|오답 원인을|복습 이력을|과제 이행 기록을|단원별 이해도를|"
            r"다시 풀기 결과를|진단 결과를|학습 시간 배분을)\s+바탕으로\s+[^.!?]+[.!?]?\s*$",
            normalized,
        )
        signatures = {normalized}
        if process_core:
            signatures.add("process-core:" + process_core.group(0))
        if not normalized or signatures & seen:
            continue
        seen.update(signatures)
        kept.append(normalized)
    return " ".join(kept)


def revised_markdown(
    source_h1: str,
    source_intro: str,
    source_sections: list[tuple[str, list[str]]],
    locality: str,
    all_localities: set[str],
    current_path: Path,
    seed: str,
    profile: str,
    expected_area: str,
    admin_replacements: dict[str, str],
) -> str:
    current_intro, current_sections = current_manuscript(current_path)
    intro = revise_copy(
        source_intro, locality, all_localities, seed + "|intro", profile,
        expected_area=expected_area, admin_replacements=admin_replacements,
    )
    # A short, already fact-checked local sentence from the current page is kept
    # as context. This makes the result a genuine revision rather than a copy of
    # the reused workbook while retaining the workbook's topic and sequence.
    current_intro = revise_copy(
        current_intro, locality, {locality}, seed + "|current-intro", profile,
        replace_center=False,
    )
    current_intro = sentence_excerpt(current_intro)
    if current_intro != intro:
        intro = f"{intro} {current_intro}"

    seen_sentences: set[str] = set()
    intro = deduplicate_page_sentences(intro, seen_sentences)

    used_context: set[tuple[int, int]] = set()
    rendered: list[str] = [intro]
    seen_headings: set[str] = set()
    for section_index, (source_heading, paragraphs) in enumerate(source_sections):
        kind = classify_heading(source_heading)
        heading = stable_pick(seed, f"heading-{section_index}-{kind}", HEADING_BANKS[kind])
        if heading in seen_headings:
            heading = f"{heading} {section_index + 1}"
        seen_headings.add(heading)
        revised = [
            revise_copy(
                paragraph, locality, all_localities, f"{seed}|{section_index}|{index}", profile,
                expected_area=expected_area, admin_replacements=admin_replacements,
            )
            for index, paragraph in enumerate(paragraphs)
        ]
        revised = [paragraph for paragraph in revised if paragraph and not source_operational_claim(paragraph)]
        context = select_context(" ".join(revised), current_sections, used_context)
        if context:
            context = revise_copy(
                context, locality, {locality}, f"{seed}|context|{section_index}", profile,
                replace_center=False,
            )
            # The retained paragraph is already a complete, fact-checked
            # sentence.  Appending a stock lead-in can create stacked topics
            # ("...때는 OO 상담은...") or duplicate "대조하면" clauses.
            revised.append(context)
        revised = [deduplicate_page_sentences(paragraph, seen_sentences) for paragraph in revised]
        revised = [paragraph for paragraph in revised if paragraph]
        rendered.append("## " + heading + "\n\n" + "\n\n".join(revised))
    # Source and retained context are revised independently.  Collapse only
    # after composition as well so a boundary cannot produce forms such as
    # "유형 유형부터" or "진도 진도표".
    return collapse_adjacent_duplicates("\n\n".join(rendered).strip())


def add_revision_marker(source: str, cell: SourceCell) -> str:
    marker = (
        '<section class="section manuscript-wrap" '
        f'data-revision="composite-2026-08-22" data-source-row="{cell.row}" '
        f'data-source-sha256="{cell.raw_sha256}">'
    )
    if source.count('<section class="section manuscript-wrap">') != 1:
        raise GateError(f"manuscript wrapper count for {cell.profile}/{cell.locality}")
    return source.replace('<section class="section manuscript-wrap">', marker, 1)


def revise_sitemap(target_urls: set[str]) -> Document:
    path = ROOT / "sitemap.xml"
    before = path.read_bytes()
    source = before.decode("utf-8")
    pattern = re.compile(r"(<url><loc>([^<]+)</loc><lastmod>)([^<]+)(</lastmod></url>)")
    seen: set[str] = set()

    def replace(match: re.Match[str]) -> str:
        url = match.group(2)
        if url not in target_urls:
            return match.group(0)
        if url in seen:
            raise GateError(f"duplicate target URL in sitemap: {url}")
        seen.add(url)
        return match.group(1) + REVISION_DATE + match.group(4)

    revised = pattern.sub(replace, source)
    missing = target_urls - seen
    if missing or len(seen) != EXPECTED_DETAIL_DOCUMENTS:
        raise GateError(f"sitemap target mismatch seen={len(seen)} missing={list(sorted(missing))[:3]}")
    return Document(
        path=path, before=before, after=revised.encode("utf-8"),
        profile="discovery", locality="", source_sha256=sha256_bytes(before),
    )


def manifest_for_documents(documents: Iterable[Document], use_after: bool) -> str:
    digest = hashlib.sha256()
    for doc in sorted(documents, key=lambda item: item.path.as_posix()):
        digest.update(doc.path.relative_to(ROOT).as_posix().encode("utf-8") + b"\0")
        digest.update(hashlib.sha256(doc.after if use_after else doc.before).digest())
    return digest.hexdigest()


def page_graph(source: str) -> list[dict]:
    raw = extract_one(r'<script\s+type="application/ld\+json">(.*?)</script>', source)
    data = json.loads(raw)
    graph = data.get("@graph", []) if isinstance(data, dict) else []
    if not isinstance(graph, list):
        raise GateError("invalid JSON-LD graph")
    return graph


def node_types(node: dict) -> set[str]:
    value = node.get("@type", [])
    return {value} if isinstance(value, str) else set(value) if isinstance(value, list) else set()


def validate_document(doc: Document, source_cell: SourceCell, center: dict, base) -> dict[str, object]:
    source = doc.after.decode("utf-8")
    rel = doc.path.relative_to(ROOT).as_posix()
    if source.startswith("\ufeff") or "\x00" in source:
        raise GateError(f"encoding marker/control: {rel}")
    if source.count("<h1") != 1 or source.count('rel="canonical"') != 1:
        raise GateError(f"H1/canonical cardinality: {rel}")
    if f'data-source-row="{source_cell.row}"' not in source or f'data-source-sha256="{source_cell.raw_sha256}"' not in source:
        raise GateError(f"source hook mismatch: {rel}")
    canonical = extract_one(r'<link\s+rel="canonical"\s+href="([^"]+)"', source)
    expected = base.canonical_url(re.sub(r"\s+", "", source_cell.locality))
    if canonical != expected:
        raise GateError(f"canonical mismatch: {rel}")
    visible = visible_text(source)
    if source_cell.locality not in visible:
        raise GateError(f"locality absent: {rel}")
    forbidden = (
        "LOCAL ACADEMY GUIDE", "점수를 안정적으로 올립니다", "상위권 실력으로 확장합니다",
        "실력을 끌어올립니다", "최적의 맞춤형", "검색엔진", "상위노출", "SEO",
    )
    for phrase in forbidden:
        if phrase in visible:
            raise GateError(f"unrevised/promotional copy {phrase!r}: {rel}")
    graph = page_graph(source)
    articles = [node for node in graph if isinstance(node, dict) and "Article" in node_types(node)]
    faqs = [node for node in graph if isinstance(node, dict) and "FAQPage" in node_types(node)]
    if len(articles) != 1 or len(faqs) != 1 or articles[0].get("dateModified") != REVISION_DATE:
        raise GateError(f"Article/FAQ/date graph contract: {rel}")
    details = re.findall(r'<details(?:\s+[^>]*)?>\s*<summary>(.*?)</summary>\s*<p>(.*?)</p>\s*</details>', source, flags=re.I | re.S)
    visible_faq = [(visible_text(q), visible_text(a)) for q, a in details]
    schema_faq = [
        (item.get("name", ""), item.get("acceptedAnswer", {}).get("text", ""))
        for item in faqs[0].get("mainEntity", [])
    ]
    if visible_faq != schema_faq or len(visible_faq) != 4:
        raise GateError(f"FAQ visible/schema parity: {rel}")
    manuscript = extract_one(
        r'<section\s+class="section manuscript-wrap"[^>]*>(.*?)</article>\s*</section>', source,
    )
    paragraphs = [visible_text(item) for item in re.findall(r"<p\b[^>]*>(.*?)</p>", manuscript, flags=re.I | re.S)]
    manuscript_visible = visible_text(manuscript)
    question_count = manuscript_visible.count("?")
    direct_question_allowed = (
        source_cell.profile == "middle" and source_cell.locality == "풍동"
        and question_count == 1 and DIRECT_QUESTION in manuscript_visible
    )
    if question_count and not direct_question_allowed:
        raise GateError(f"unexpected question-mark residue count={question_count}: {rel}")
    claim_phrases = (
        "성적이 오르는", "성적이 오르게", "성적으로 연결되도록", "성적 상승",
        "성적 향상", "점수를 끌어올", "점수 향상", "확실히", "확실하게",
        "성적을 올", "성적을 높", "최고의", "최적의 맞춤", "오답을 점수로 바꾸",
        "실력 상승", "실력이 누적되도록", "성장할 수 있게", "점수로 연결되는",
        "출제 경향",
    )
    for phrase in claim_phrases:
        if phrase in manuscript_visible:
            raise GateError(f"unsoftened outcome claim {phrase!r}: {rel}")
    raise_score = re.search(r"(?:(?:내신|시험|학업)\s+)?(?:점수|성적|실력)(?:을|를)?[^.!?]{0,30}(?:끌어)?올", manuscript_visible)
    if raise_score:
        context = manuscript_visible[max(0, raise_score.start() - 60):raise_score.end() + 100]
        raise GateError(f"unsoftened raise-score family {context!r}: {rel}")
    rising_score = re.search(r"(?:(?:내신|시험|학업)\s+)?(?:점수|성적|실력)(?:이|가)[^.!?]{0,30}(?:오르|올라)", manuscript_visible)
    if rising_score:
        raise GateError(f"unsoftened rising-score family {rising_score.group(0)!r}: {rel}")
    duplicate_match = re.search(
        r"(?<![가-힣0-9])([가-힣][가-힣0-9]{1,24})\s+\1(?![가-힣0-9])",
        manuscript_visible,
    )
    if duplicate_match:
        raise GateError(f"adjacent duplicate token {duplicate_match.group(1)!r}: {rel}")
    middle_dot_duplicate = re.search(
        r"(?<![가-힣])([가-힣]{2,})·\1(?:학습|형)?(?![가-힣])",
        manuscript_visible,
    )
    if middle_dot_duplicate:
        raise GateError(f"middle-dot duplicate token {middle_dot_duplicate.group(0)!r}: {rel}")
    malformed_phrases = (
        "학습 변화이", "점수 변화이", "학습 변화으로", "점수 변화으로",
        "학습 변화과", "점수 변화과", "학습 변화은", "점수 변화은",
        "현재 수준에 맞춘으로", "내신 제공된 시험지", "시험 제공된 시험지",
        "영약", "유형 학습-", "을 설계 기준을", "를 설계 기준을",
        "국어·영어·수학를", "영어·수학를", "별교과 평가도시",
        "영어·수학·국어·영어·수학", "영어/수학/국어/영어·수학",
        "국어/국어·영어·수학", "국어와 국어·영어·수학",
        "최적화된", "삼삼송", "비전주 장동기보다",
        "성취 향상.",
        "교과 평가이", "교과 평가을", "교과 평가은", "교과 평가과", "교과 평가으로",
        "정재송동게", "정진월동게", "정미사게", "정금곡동게",
        "정소사벌게", "정영천동게", "정구월동게", "정칠곡게",
    )
    for phrase in malformed_phrases:
        if phrase in manuscript_visible:
            raise GateError(f"malformed revised copy {phrase!r}: {rel}")
    if re.search(r"목표에\s+맞(?:춘|는)\s+현재\s+수준에\s+맞춘", manuscript_visible):
        raise GateError(f"duplicated goal/level modifier: {rel}")
    if "현재 수준에 맞춘 현재 수준에 맞춘" in manuscript_visible:
        raise GateError(f"duplicated current-level modifier: {rel}")
    if re.search(r"(?:점수|성적)[^.!?]{0,24}(?:확보|만듭)", manuscript_visible):
        raise GateError(f"unsoftened score-acquisition claim: {rel}")
    for paragraph in paragraphs:
        for sentence in re.split(r"(?<=[.!?])\s+", paragraph):
            if not has_direct_outcome_claim(sentence):
                continue
            if (
                ("보장" in sentence and any(token in sentence for token in ("않", "아니", "없", "말보다")))
                or "약속보다" in sentence or "약속이 아니라" in sentence
            ):
                continue
            raise GateError(f"unsoftened direct outcome claim {sentence!r}: {rel}")
    for paragraph in paragraphs:
        if paragraph.endswith(")"):
            raise GateError(f"paragraph lacks terminal punctuation: {paragraph!r}: {rel}")
    subject_duplication = re.search(
        r"(?:영어·수학(?:은\s+물론|뿐\s+아니라|과|\s+중심의)?|영어,\s*수학(?:뿐\s+아니라)?)"
        r"[^.!?]{0,28}국어·영어·수학"
        r"|영어·수학·국어[^.!?]{0,35}국어·영어·수학",
        manuscript_visible,
    )
    if subject_duplication:
        raise GateError(f"duplicated subject sequence {subject_duplication.group(0)!r}: {rel}")
    if re.search(rf"{re.escape(source_cell.locality)}(?:해|하여|형인지)", manuscript_visible):
        raise GateError(f"locality substituted inside ordinary word: {rel}")
    if re.search(r"(?:을|를)\s+단계를\s+나눠", manuscript_visible):
        raise GateError(f"malformed staged-copy particle: {rel}")
    expected_area = center_area(center, source_cell.locality)
    area_pattern = re.compile(
        ADMIN_PREFIX_PATTERN + re.escape(source_cell.locality)
    )
    allowed_geo = {
        clean(center.get("region", "")), clean(center.get("display_region", "")),
        clean(center.get("district", "")), clean(center.get("display_district", "")),
    } - {""}
    wrong_areas: set[str] = set()
    for match in area_pattern.finditer(manuscript_visible):
        prefix = match.group(0)[:-len(source_cell.locality)].strip()
        tokens = prefix.split()
        has_allowed_city = any(
            token in allowed_geo and token.endswith(("시", "군")) for token in tokens
        )
        invalid = [
            token for token in tokens
            if token not in allowed_geo and not (has_allowed_city and token.endswith("구"))
        ]
        if invalid:
            wrong_areas.add(match.group(0))
    if wrong_areas:
        raise GateError(f"administrative locality mismatch {sorted(wrong_areas)!r}: {rel}")
    normalized = [re.sub(r"\W+", "", paragraph) for paragraph in paragraphs if paragraph]
    if len(normalized) != len(set(normalized)):
        raise GateError(f"within-page paragraph duplicate: {rel}")
    return {"paragraphs": paragraphs, "manuscript": manuscript_visible}


def build_plan(archive: Path, common_dir: Path, *, run_profile_preflight: bool = True) -> BuildPlan:
    if ROOT.resolve() != Path(__file__).resolve().parents[1]:
        raise GateError("generator/root mismatch")
    assert_hash(ROOT / "scripts" / "subject_catalog.py", EXPECTED_SUBJECT_CATALOG_SHA256, "subject catalog")
    assert_hash(common_dir / "센터정보 정리.csv", EXPECTED_CENTER_CSV_SHA256, "center CSV")
    assert_hash(common_dir / "이미지링크.csv", EXPECTED_IMAGE_CSV_SHA256, "image CSV")
    workbooks, source_manifest = read_archive(archive)
    source_manifest.update({
        "base_commit": BASE_COMMIT,
        "base_helper": EXPECTED_BASE_HELPER_SHA256,
        "subject_catalog": EXPECTED_SUBJECT_CATALOG_SHA256,
        "center_csv": EXPECTED_CENTER_CSV_SHA256,
        "image_csv": EXPECTED_IMAGE_CSV_SHA256,
    })

    if run_git("cat-file", "-t", BASE_COMMIT) != "commit":
        raise GateError(f"pinned base commit unavailable: {BASE_COMMIT}")
    baseline_paths: list[Path] = []
    for _, _, category in PROFILE_SPECS:
        category_root = SUBJECT_ROOT / category
        paths = sorted(category_root.glob("*/index.html"))
        if len(paths) != EXPECTED_DETAIL_COUNT:
            raise GateError(f"baseline detail path count for {category}: {len(paths)}")
        baseline_paths.extend(paths)
    prepare_baseline_blobs(baseline_paths)

    documents: list[Document] = []
    all_source_cells: dict[tuple[str, str], SourceCell] = {}
    profile_metrics: dict[str, object] = {}
    paragraph_df: dict[str, set[str]] = defaultdict(set)
    sentence_df: dict[str, set[str]] = defaultdict(set)
    foreign_source_mentions = 0
    target_canonicals: set[str] = set()

    for profile, workbook_name, expected_category in PROFILE_SPECS:
        base = load_base(profile, common_dir)
        if base.CATEGORY_SLUG != expected_category:
            raise GateError(f"profile/category mismatch: {profile}")
        centers = base.load_centers()
        locality_order = list(centers)
        if len(locality_order) != EXPECTED_DETAIL_COUNT:
            raise GateError(f"center count for {profile}: {len(locality_order)}")
        all_localities = set(locality_order)
        rows = workbooks[workbook_name]
        source_by_locality: dict[str, SourceCell] = {}
        for row, (locality, raw) in enumerate(zip(locality_order, rows), 2):
            h1, intro, sections = source_fragment(raw)
            raw_visible = visible_text(raw)
            foreign = {item for item in all_localities - {locality} if item and item in raw_visible}
            foreign_source_mentions += len(foreign)
            body_without_h1 = re.sub(r"<h1\b[^>]*>.*?</h1>", "", raw, flags=re.I | re.S)
            body_visible = visible_text(body_without_h1)
            intended_count = body_visible.count(locality)
            foreign_counts = sorted(
                ((body_visible.count(item), item) for item in foreign if body_visible.count(item)),
                reverse=True,
            )
            corrected_localities: tuple[str, ...] = ()
            if foreign_counts:
                top_count, top_locality = foreign_counts[0]
                if intended_count == 0 or (top_count >= 3 and top_count >= intended_count * 2):
                    corrected_localities = (top_locality,)
            cell = SourceCell(
                workbook=workbook_name, profile=profile, row=row, locality=locality,
                raw_html=raw, raw_sha256=sha256_bytes(raw.encode("utf-8")), source_h1=h1,
                source_intro=intro,
                source_sections=tuple((heading, tuple(paragraphs)) for heading, paragraphs in sections),
                corrected_localities=corrected_localities,
            )
            source_by_locality[locality] = cell
            all_source_cells[(profile, locality)] = cell

        records = base.build_records()
        if {record["locality"] for record in records} != all_localities:
            raise GateError(f"record/locality mismatch for {profile}")
        for record in records:
            locality = record["locality"]
            cell = source_by_locality[locality]
            current_path = base.CATEGORY_ROOT / record["slug"] / "index.html"
            if not current_path.is_file():
                raise GateError(f"existing route missing: {current_path}")
            record["sections"] = dict(record["sections"])
            record["sections"]["본문"] = revised_markdown(
                cell.source_h1, cell.source_intro,
                [(heading, list(paragraphs)) for heading, paragraphs in cell.source_sections],
                locality, all_localities, current_path,
                f"{profile}|{record['slug']}|{cell.raw_sha256}", profile,
                center_area(record["center"], locality),
                administrative_replacements(centers, record["center"]),
            )

        original_personalize = base.personalize_body

        def post_sanitize_personalized(
            body: str,
            title: str,
            locality: str,
            center: dict,
            student_type: str,
            seed: str,
        ):
            intro_value, section_values = original_personalize(
                body, title, locality, center, student_type, seed,
            )
            intro_value = revise_copy(
                intro_value, locality, {locality}, seed + "|post-intro", profile,
                replace_center=False,
            )
            post_seen_sentences: set[str] = set()
            intro_value = deduplicate_page_sentences(intro_value, post_seen_sentences)
            cleaned_sections: list[tuple[str, list[str]]] = []
            for section_index, (heading, paragraphs) in enumerate(section_values):
                cleaned_paragraphs = [
                    revise_copy(
                        paragraph, locality, {locality},
                        f"{seed}|post|{section_index}|{paragraph_index}", profile,
                        replace_center=False,
                    )
                    for paragraph_index, paragraph in enumerate(paragraphs)
                ]
                cleaned_paragraphs = [
                    paragraph for paragraph in cleaned_paragraphs
                    if paragraph and not source_operational_claim(paragraph)
                ]
                cleaned_paragraphs = [
                    deduplicate_page_sentences(paragraph, post_seen_sentences)
                    for paragraph in cleaned_paragraphs
                ]
                cleaned_paragraphs = [paragraph for paragraph in cleaned_paragraphs if paragraph]
                if not cleaned_paragraphs:
                    continue
                cleaned_sections.append((
                    heading,
                    cleaned_paragraphs,
                ))
            return intro_value, cleaned_sections

        base.personalize_body = post_sanitize_personalized

        outputs: dict[Path, str] = {}
        for index, record in enumerate(records):
            previous_record = records[(index - 1) % len(records)]
            next_record = records[(index + 1) % len(records)]
            cell = source_by_locality[record["locality"]]
            rendered = base.render_page(record, previous_record, next_record)
            rendered = add_revision_marker(rendered, cell)
            path = base.CATEGORY_ROOT / record["slug"] / "index.html"
            outputs[path] = rendered
            before = path.read_bytes()
            after = rendered.encode("utf-8")
            documents.append(Document(path, before, after, profile, record["locality"], cell.raw_sha256))
            target_canonicals.add(base.canonical_url(record["slug"]))

        if len(outputs) != EXPECTED_DETAIL_COUNT:
            raise GateError(f"render count for {profile}: {len(outputs)}")
        if run_profile_preflight:
            full_outputs = dict(outputs)
            full_outputs[base.CATEGORY_ROOT / "index.html"] = (base.CATEGORY_ROOT / "index.html").read_text(encoding="utf-8")
            try:
                report = base.preflight(records, full_outputs, include_corpus=False)
            except ValueError as exc:
                try:
                    failed_report = json.loads(str(exc))
                    failed_errors = failed_report.get("errors", [])
                except (json.JSONDecodeError, AttributeError):
                    raise GateError(f"base preflight failed for {profile}: {exc}") from exc
                raise GateError(f"base preflight failed for {profile}: {failed_errors[:25]}") from exc
            if report.get("errors"):
                sample = report["errors"][:20] if isinstance(report["errors"], list) else report["errors"]
                raise GateError(f"base preflight failed for {profile}: {sample}")
            profile_metrics[profile] = report

        for path, rendered in outputs.items():
            locality = path.parent.name
            cell = source_by_locality[next(item for item in source_by_locality if re.sub(r"\s+", "", item) == locality)]
            detail = next(doc for doc in documents if doc.path == path)
            observation = validate_document(detail, cell, centers[cell.locality], base)
            rel = path.relative_to(ROOT).as_posix()
            for paragraph in observation["paragraphs"]:
                norm = re.sub(r"\d+", "{N}", re.sub(r"\s+", " ", paragraph)).strip()
                for token in sorted(all_localities, key=len, reverse=True):
                    norm = norm.replace(token, "{지역}")
                if len(norm) >= 30:
                    paragraph_df[norm].add(rel)
                for sentence in re.split(r"(?<=[.!?])\s+", norm):
                    if len(sentence) >= 45:
                        sentence_df[sentence].add(rel)

    if len(target_canonicals) != EXPECTED_DETAIL_DOCUMENTS:
        raise GateError(f"target canonical scope={len(target_canonicals)} expected={EXPECTED_DETAIL_DOCUMENTS}")
    documents.append(revise_sitemap(target_canonicals))
    if len(documents) != EXPECTED_DOCUMENT_COUNT or len({doc.path for doc in documents}) != EXPECTED_DOCUMENT_COUNT:
        raise GateError(f"document scope={len(documents)} expected={EXPECTED_DOCUMENT_COUNT}")
    allowed_roots = (SUBJECT_ROOT.resolve(), (ROOT / "sitemap.xml").resolve())
    for doc in documents:
        resolved = doc.path.resolve()
        if resolved != allowed_roots[1] and not resolved.is_relative_to(allowed_roots[0]):
            raise GateError(f"authorized path escaped revision scope: {doc.path}")

    paragraph_max_df = max((len(paths) for paths in paragraph_df.values()), default=0)
    sentence_max_df = max((len(paths) for paths in sentence_df.values()), default=0)
    if paragraph_max_df > 80 or sentence_max_df > 100:
        raise GateError(f"revised corpus repetition too high paragraph={paragraph_max_df} sentence={sentence_max_df}")
    changed_count = sum(doc.before != doc.after for doc in documents)
    before_manifest = manifest_for_documents(documents, False)
    after_manifest = manifest_for_documents(documents, True)
    if changed_count not in {0, EXPECTED_DOCUMENT_COUNT} and before_manifest != EXPECTED_SUPERSEDED_AFTER_MANIFEST:
        raise GateError(
            f"partial materialization detected: changed={changed_count} "
            f"before={before_manifest} expected 0 or {EXPECTED_DOCUMENT_COUNT}"
        )

    metrics = {
        "documents": len(documents),
        "detail_documents": EXPECTED_DETAIL_DOCUMENTS,
        "sitemap_documents": 1,
        "changed": changed_count,
        "profiles": profile_metrics,
        "foreign_source_mentions_corrected": foreign_source_mentions,
        "paragraph_max_df": paragraph_max_df,
        "sentence_max_df": sentence_max_df,
        "before_manifest": before_manifest,
        "after_manifest": after_manifest,
    }
    return BuildPlan(documents, source_manifest, metrics)


def assert_repeat(plan: BuildPlan, archive: Path, common_dir: Path) -> None:
    repeat = build_plan(archive, common_dir, run_profile_preflight=False)
    if plan.candidate_sha256() != repeat.candidate_sha256():
        raise GateError("repeat build candidate mismatch")
    first = {doc.path: doc.after for doc in plan.documents}
    second = {doc.path: doc.after for doc in repeat.documents}
    if first != second:
        raise GateError("repeat build bytes mismatch")


def freeze_bytes(payload: dict[str, object]) -> bytes:
    return (json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def write_freeze(path: Path, plan: BuildPlan) -> None:
    data = freeze_bytes(plan.freeze_payload())
    path.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except Exception:
        path.unlink(missing_ok=True)
        raise


def authorized_pending_plan(plan: BuildPlan) -> bool:
    changed_count = len(plan.changed)
    if changed_count == EXPECTED_DOCUMENT_COUNT:
        return True
    return (
        0 < changed_count < EXPECTED_DOCUMENT_COUNT
        and plan.metrics.get("before_manifest") == EXPECTED_SUPERSEDED_AFTER_MANIFEST
        and plan.metrics.get("after_manifest") != EXPECTED_SUPERSEDED_AFTER_MANIFEST
    )


def verify_freeze(path: Path, plan: BuildPlan) -> None:
    expected = freeze_bytes(plan.freeze_payload())
    actual = path.read_bytes()
    if actual != expected:
        raise GateError(f"freeze mismatch: {sha256_bytes(actual)} != {sha256_bytes(expected)}")


def transaction_residue(root: Path = ROOT) -> list[Path]:
    patterns = (".revise-kem-*.tmp", ".revise-kem-*.bak", ".revise-kem-transaction.json", ".revise-kem.lock")
    found: list[Path] = []
    for pattern in patterns:
        found.extend(root.rglob(pattern))
    return sorted(set(found))


def apply_plan(plan: BuildPlan) -> None:
    if transaction_residue():
        raise GateError("transaction residue exists before apply")
    lock = ROOT / ".revise-kem.lock"
    journal = ROOT / ".revise-kem-transaction.json"
    lock_fd = os.open(lock, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    token = uuid.uuid4().hex
    staged: dict[Path, Path] = {}
    backups: dict[Path, Path] = {}
    committed: list[Path] = []
    try:
        with os.fdopen(lock_fd, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(json.dumps({"pid": os.getpid(), "token": token, "time": time.time()}) + "\n")
            stream.flush(); os.fsync(stream.fileno())
        for doc in plan.documents:
            if doc.path.read_bytes() != doc.before:
                raise GateError(f"concurrent edit before staging: {doc.path}")
            temp = doc.path.with_name(f".revise-kem-{token}.tmp")
            with temp.open("xb") as stream:
                stream.write(doc.after); stream.flush(); os.fsync(stream.fileno())
            if temp.read_bytes() != doc.after:
                raise GateError(f"staged bytes mismatch: {doc.path}")
            staged[doc.path] = temp
        journal.write_text(json.dumps({
            "token": token,
            "paths": [doc.path.relative_to(ROOT).as_posix() for doc in plan.documents],
            "committed": [],
        }, ensure_ascii=False), encoding="utf-8", newline="\n")
        for doc in plan.documents:
            backup = doc.path.with_name(f".revise-kem-{token}.bak")
            os.replace(doc.path, backup)
            backups[doc.path] = backup
            os.replace(staged[doc.path], doc.path)
            committed.append(doc.path)
            journal.write_text(json.dumps({
                "token": token,
                "paths": [item.path.relative_to(ROOT).as_posix() for item in plan.documents],
                "committed": [item.relative_to(ROOT).as_posix() for item in committed],
            }, ensure_ascii=False), encoding="utf-8", newline="\n")
        for doc in plan.documents:
            if doc.path.read_bytes() != doc.after:
                raise GateError(f"post-apply mismatch: {doc.path}")
        for backup in backups.values():
            backup.unlink()
        journal.unlink()
    except Exception:
        for path in reversed(committed):
            backup = backups.get(path)
            if backup and backup.exists():
                path.unlink(missing_ok=True)
                os.replace(backup, path)
        for path, backup in backups.items():
            if backup.exists() and not path.exists():
                os.replace(backup, path)
        raise
    finally:
        for temp in staged.values():
            temp.unlink(missing_ok=True)
        for backup in backups.values():
            backup.unlink(missing_ok=True)
        journal.unlink(missing_ok=True)
        lock.unlink(missing_ok=True)
    if transaction_residue():
        raise GateError("transaction residue after apply")


def report(plan: BuildPlan, mode: str) -> dict[str, object]:
    return {
        "status": "PASS",
        "mode": mode,
        "root": str(ROOT),
        "generator_sha256": sha256_file(Path(__file__)),
        "candidate_sha256": plan.candidate_sha256(),
        "documents": len(plan.documents),
        "changed": len(plan.changed),
        "metrics": plan.metrics,
        "source_manifest": plan.source_manifest,
        "residue": [str(path.relative_to(ROOT)) for path in transaction_residue()],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Revise the three existing KEM detail-page sets from an attached XLSX archive.")
    parser.add_argument("--archive", type=Path, default=DEFAULT_ARCHIVE)
    parser.add_argument("--common-dir", type=Path, default=DEFAULT_COMMON)
    parser.add_argument("--freeze-out", type=Path)
    parser.add_argument("--freeze-file", type=Path)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--go", default="")
    parser.add_argument("--skip-repeat", action="store_true")
    args = parser.parse_args()

    plan = build_plan(args.archive.resolve(), args.common_dir.resolve())
    if not args.skip_repeat:
        assert_repeat(plan, args.archive.resolve(), args.common_dir.resolve())
    if args.freeze_out:
        if args.apply:
            raise GateError("--freeze-out and --apply cannot be combined")
        if not authorized_pending_plan(plan):
            raise GateError(f"freeze requires an authorized pending plan: changed={len(plan.changed)}")
        write_freeze(args.freeze_out.resolve(), plan)
    if args.apply:
        if args.go != "APPLY-GO" or not args.freeze_file:
            raise GateError("apply requires --go APPLY-GO and --freeze-file")
        if not authorized_pending_plan(plan):
            raise GateError(f"apply requires an authorized pending plan: changed={len(plan.changed)}")
        verify_freeze(args.freeze_file.resolve(), plan)
        apply_plan(plan)
        post = build_plan(args.archive.resolve(), args.common_dir.resolve(), run_profile_preflight=True)
        if post.changed:
            raise GateError(f"post-apply idempotency failed: {len(post.changed)} changed")
        result = report(post, "apply")
    else:
        result = report(plan, "dry-run")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (GateError, FileExistsError, PermissionError) as exc:
        print(json.dumps({"status": "FAIL", "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(1)
