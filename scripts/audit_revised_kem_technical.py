from __future__ import annotations

"""Technical release gate for the revised K/E/M detail-page release.

The gate is deliberately independent of the generator.  It treats the attached
ZIP/XLSX files, the pre-release Git tree, a generator freeze, the candidate
working tree, HTTP deployments, browser evidence, and Vercel build provenance
as separate trust boundaries.

This program never rewrites site files.  Baselines, reports, and browser plans
may only be written outside the repository.  Network and Vercel checks are
read-only.
"""

import argparse
import ast
import concurrent.futures
import hashlib
import html
import io
import json
import os
import re
import ssl
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Iterable, Sequence
from zipfile import BadZipFile, ZipFile


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = Path(__file__).resolve()
GENERATOR_PATH = ROOT / "scripts" / "generate_revised_kem_pages.py"
CONTENT_AUDITOR_PATH = ROOT / "scripts" / "audit_revised_kem_content.py"
TECHNICAL_AUDITOR_PATH = SCRIPT_PATH
DEFAULT_ARCHIVE = Path.home() / "Desktop" / "1.zip"
DEFAULT_COMMON = ROOT.parent / "참고자료" / "공통자료"

SITE_ORIGIN = "https://xn--3e0bz50b1zcyxat54c.com"
DISPLAY_DOMAIN = "국영수학원.com"
VERCEL_PROJECT = "new14"
GITHUB_REPOSITORY = "01039578283-hub/new14"
EXPECTED_REMOTE_URLS = {
    "https://github.com/01039578283-hub/new14.git",
    "git@github.com:01039578283-hub/new14.git",
}
BASELINE_COMMIT = "9e58f271f6126db72d4eb10a363c9d3b4d163779"
REVISION_DATE = "2026-08-22"

# Final frozen release tool/candidate pins.  This transition is intentionally
# anchored to the exact, fully materialized superseded release rather than to
# either the original Git baseline or an arbitrary partially applied state.
EXPECTED_GENERATOR_SHA256 = "f145adec84c78a61bfcf30b8a137e5d741e0eb1f1ba59a4370fc2409533b296e"
EXPECTED_CANDIDATE_SHA256 = "33f212961e34e6978d9dfa5b0eeab9cad7916a0da6626afccadcd67a5f17b9e6"
EXPECTED_AFTER_MANIFEST_SHA256 = "081de645e104568f2b63e019093c5656ba585dc4204947b92098d938ee240cc9"
EXPECTED_CONTENT_AUDITOR_SHA256 = "3b0c7fcc96c3fd782c1f15e3eceac5b2dc1ddd046ccabbd33234cc7437d4b035"
EXPECTED_RAW_BEFORE_MANIFEST_SHA256 = "b158341e6e8cc27bb7951e4c9d02faae0d91651daf1c5d24ace64d2af7d3ce56"
EXPECTED_VERIFIED_GIT_MANIFEST_SHA256 = "69da99784e69a5d2d21fd69f3d6434872d00a455b2f907f64484b51a14617491"
EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256 = "f3c0ec8fd1a03da9f8ad04571c9cedcf0c0cc0964382ae6c62f31fa1bdb5b0aa"
EXPECTED_PROVENANCE_MANIFEST_SHA256 = "5bd2879042de2725d5e34fd3db4f86f14db1f2e51a1d542a8c9c88c58bd07f13"
EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256 = "f8828e61788be26d9d72e5ed619b450d626f5f0964052686730c4b6f50c8f451"
EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS = 290

EXPECTED_ARCHIVE_SHA256 = "20c268ce05bd48c18c659a629ae522ad9470eac8dba07ebf4d435b57d6a1d57f"
EXPECTED_WORKBOOK_SHA256 = {
    "초등학생학원 원고.xlsx": "deb8a99bd51b9d9f5792cf8149e10c4a3c38579af870985e466b039939ac834d",
    "중학생학원 원고.xlsx": "336376c4186a3a5fea2137a3d2bc9a28f7876e9cde6137762b66917ff6c63558",
    "고등학생학원 원고.xlsx": "738a76a553efdb7af4e5e6cbe5f22de1b00847cd7b3b659f55f623cb3537797c",
}
EXPECTED_SOURCE_SUPPORT = {
    "base_helper": (ROOT / "scripts" / "generate_highschool_korean_english_math.py", "36f3d788760470f3984430b4f30fb6c9c630ac1b55387ead1f05df7f01fb5881"),
    "subject_catalog": (ROOT / "scripts" / "subject_catalog.py", "9ae6ae6757b4a7766717822bcd6183920f1381c5d2f643eaab82c6ec39550814"),
    "center_csv": (DEFAULT_COMMON / "센터정보 정리.csv", "3ffbd7b70273b6dc1c8435c53a3a25e32d2a173ba1bf51840654389bd8954e1a"),
    "image_csv": (DEFAULT_COMMON / "이미지링크.csv", "c1b4f87b2b62f659107dbf0a79a1d566e213e008fc4b7f30cfa656ffae814100"),
}

PROFILE_SPECS = {
    "elementary": ("초등학생학원 원고.xlsx", "초등학생국영수학원"),
    "middle": ("중학생학원 원고.xlsx", "중학생국영수학원"),
    "high": ("고등학생학원 원고.xlsx", "고등학생국영수학원"),
}
EXPECTED_PER_PROFILE = 371
EXPECTED_DETAIL_COUNT = 1_113
EXPECTED_PRODUCT_DOCUMENT_COUNT = 1_114
EXPECTED_SITEMAP_COUNT = 3_725
EXPECTED_RELEASE_PATH_COUNT = 1_117
TOOL_PATHS = {
    "scripts/generate_revised_kem_pages.py",
    "scripts/audit_revised_kem_content.py",
    "scripts/audit_revised_kem_technical.py",
}
REQUIRED_SCHEMA_TYPES = {
    "EducationalOrganization",
    "LocalBusiness",
    "WebPage",
    "BreadcrumbList",
    "Article",
    "Service",
    "FAQPage",
    "ItemList",
}
FORBIDDEN_SCHEMA_TYPES = {"Review", "AggregateRating"}

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
JSONLD_RE = re.compile(
    r'<script\b[^>]*\btype=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
    re.I | re.S,
)
TAG_ATTR_RE = re.compile(r"([\w:-]+)\s*=\s*([\"'])(.*?)\2", re.I | re.S)
URL_BLOCK_RE = re.compile(r"<url\b[^>]*>.*?</url>", re.I | re.S)
LASTMOD_RE = re.compile(r"(<lastmod\b[^>]*>)(.*?)(</lastmod>)", re.I | re.S)
ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")
HTTP_SSL_CONTEXT = ssl.create_default_context()


class GateError(RuntimeError):
    pass


@dataclass
class Audit:
    errors: list[dict[str, str]] = field(default_factory=list)
    warnings: list[dict[str, str]] = field(default_factory=list)

    def error(self, code: str, detail: object) -> None:
        self.errors.append({"code": code, "detail": str(detail)})

    def warn(self, code: str, detail: object) -> None:
        self.warnings.append({"code": code, "detail": str(detail)})

    def capture(self, code: str, function, *args, **kwargs):
        try:
            return function(*args, **kwargs)
        except Exception as exc:  # A release gate reports all independent failures.
            self.error(code, exc)
            return None

    def summary(self) -> dict[str, object]:
        return {
            "error_count": len(self.errors),
            "warning_count": len(self.warnings),
            "error_codes": dict(sorted(Counter(item["code"] for item in self.errors).items())),
            "warning_codes": dict(sorted(Counter(item["code"] for item in self.warnings).items())),
            "errors": self.errors[:200],
            "warnings": self.warnings[:200],
        }


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    return sha256_bytes(path.read_bytes())


def canonical_json(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def set_sha(values: Iterable[str]) -> str:
    return sha256_bytes(("\n".join(sorted(values)) + "\n").encode("utf-8"))


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", html.unescape(str(value or ""))).strip()


def strip_tags(value: str) -> str:
    value = re.sub(r"<(?:script|style)\b.*?</(?:script|style)>", " ", value, flags=re.I | re.S)
    return clean(re.sub(r"<[^>]+>", " ", value))


def safe_resolve(path: Path) -> Path:
    resolved = path.resolve()
    if resolved == ROOT or ROOT in resolved.parents:
        return resolved
    raise GateError(f"path must remain inside repository: {path}")


def assert_external_output(path: Path) -> Path:
    resolved = path.expanduser().resolve()
    if resolved == ROOT or ROOT in resolved.parents:
        raise GateError(f"output must be outside repository: {resolved}")
    return resolved


def run_git(*args: str, text: bool = True, check: bool = True):
    return subprocess.run(
        ["git", *args], cwd=ROOT, check=check,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=text, encoding="utf-8" if text else None,
        errors="replace" if text else None,
    )


def git_text(*args: str) -> str:
    return run_git(*args).stdout.strip()


def git_blob(commit: str, relative: str) -> bytes:
    return run_git("cat-file", "blob", f"{commit}:{relative}", text=False).stdout


def git_blobs(commit: str, paths: Iterable[str]) -> dict[str, bytes]:
    """Read many pinned blobs through one Git process (important on Windows)."""
    ordered = sorted(paths)
    commands = b"".join(f"{commit}:{path}\n".encode("utf-8") for path in ordered)
    process = subprocess.run(
        ["git", "cat-file", "--batch"], cwd=ROOT, check=True,
        input=commands, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    output = process.stdout
    offset = 0
    result: dict[str, bytes] = {}
    for path in ordered:
        header_end = output.find(b"\n", offset)
        if header_end < 0:
            raise GateError(f"truncated git cat-file header: {path}")
        header = output[offset:header_end].decode("ascii", errors="replace").split()
        if len(header) != 3 or header[1] != "blob" or not header[2].isdigit():
            raise GateError(f"unexpected git cat-file header for {path}: {header}")
        size = int(header[2])
        start = header_end + 1
        end = start + size
        if end >= len(output) or output[end:end + 1] != b"\n":
            raise GateError(f"truncated git cat-file body: {path}")
        result[path] = output[start:end]
        offset = end + 1
    if output[offset:]:
        raise GateError(f"unexpected trailing git cat-file bytes: {len(output) - offset}")
    return result


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def profile_detail_paths(root: Path = ROOT) -> dict[str, list[Path]]:
    result: dict[str, list[Path]] = {}
    for profile, (_, slug) in PROFILE_SPECS.items():
        category = root / "과목별학원" / slug
        pages = sorted(path for path in category.glob("*/index.html") if path.parent != category)
        result[profile] = pages
    return result


def target_paths(root: Path = ROOT) -> list[Path]:
    return sorted(path for pages in profile_detail_paths(root).values() for path in pages)


def hub_paths(root: Path = ROOT) -> list[Path]:
    return sorted(root / "과목별학원" / slug / "index.html" for _, slug in PROFILE_SPECS.values())


def page_url(path: Path, root: Path = ROOT) -> str:
    parts = path.parent.relative_to(root).parts
    encoded = "/".join(urllib.parse.quote(part, safe="") for part in parts)
    return f"{SITE_ORIGIN}/{encoded}/" if encoded else f"{SITE_ORIGIN}/"


def semantic_url(value: str) -> str:
    parsed = urllib.parse.urlsplit(clean(value))
    path = re.sub(r"/+", "/", urllib.parse.unquote(parsed.path or "/"))
    if path != "/":
        path = path.rstrip("/") + "/"
    return f"{parsed.scheme.lower()}://{parsed.netloc.lower()}{path}"


def transport_url(value: str) -> str:
    """Normalize encoding/case without erasing path-shape redirects."""
    parsed = urllib.parse.urlsplit(clean(value))
    path = urllib.parse.unquote(parsed.path or "/")
    return urllib.parse.urlunsplit((parsed.scheme.lower(), parsed.netloc.lower(), path, parsed.query, ""))


def tag_attributes(tag: str) -> dict[str, str]:
    return {match.group(1).lower(): html.unescape(match.group(3)) for match in TAG_ATTR_RE.finditer(tag)}


def first_tag_value(source: str, tag: str) -> str:
    match = re.search(rf"<{tag}\b[^>]*>(.*?)</{tag}>", source, flags=re.I | re.S)
    return strip_tags(match.group(1)) if match else ""


def meta_content(source: str, *, name: str | None = None, prop: str | None = None) -> str:
    for match in re.finditer(r"<meta\b[^>]*>", source, flags=re.I | re.S):
        attrs = tag_attributes(match.group(0))
        if name and attrs.get("name", "").lower() == name.lower():
            return clean(attrs.get("content"))
        if prop and attrs.get("property", "").lower() == prop.lower():
            return clean(attrs.get("content"))
    return ""


def canonical_href(source: str) -> str:
    for match in re.finditer(r"<link\b[^>]*>", source, flags=re.I | re.S):
        attrs = tag_attributes(match.group(0))
        if "canonical" in attrs.get("rel", "").lower().split():
            return clean(attrs.get("href"))
    return ""


def node_types(node: dict) -> set[str]:
    value = node.get("@type", [])
    return set(value if isinstance(value, list) else [value])


def json_strings(value: object) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for item in value.values():
            yield from json_strings(item)
    elif isinstance(value, list):
        for item in value:
            yield from json_strings(item)


def parse_graph(source: str) -> tuple[list[dict], list[str]]:
    nodes: list[dict] = []
    errors: list[str] = []
    for index, match in enumerate(JSONLD_RE.finditer(source), 1):
        try:
            payload = json.loads(match.group(1))
        except json.JSONDecodeError as exc:
            errors.append(f"script {index}: {exc}")
            continue
        candidates = payload.get("@graph", []) if isinstance(payload, dict) else payload
        if isinstance(candidates, dict):
            candidates = [candidates]
        if not isinstance(candidates, list):
            errors.append(f"script {index}: graph is not a list")
            continue
        nodes.extend(item for item in candidates if isinstance(item, dict))
    if not list(JSONLD_RE.finditer(source)):
        errors.append("missing JSON-LD")
    return nodes, errors


def first_node(nodes: Sequence[dict], kind: str) -> dict:
    return next((node for node in nodes if kind in node_types(node)), {})


def stable_fact_payload(nodes: Sequence[dict]) -> dict[str, object]:
    organization = first_node(nodes, "EducationalOrganization") or first_node(nodes, "LocalBusiness")
    service = first_node(nodes, "Service")
    breadcrumb = first_node(nodes, "BreadcrumbList")
    related = first_node(nodes, "ItemList")
    return {
        "organization": {key: organization.get(key) for key in (
            "@id", "name", "url", "telephone", "address", "areaServed", "parentOrganization",
            "identifier", "educationalLevel", "teaches", "makesOffer",
        ) if key in organization},
        "service": {key: service.get(key) for key in (
            "@id", "name", "serviceType", "provider", "areaServed", "audience", "offers",
        ) if key in service},
        "breadcrumb": breadcrumb.get("itemListElement", []),
        "related": related.get("itemListElement", []),
    }


def visible_faq(source: str) -> list[tuple[str, str]]:
    return [
        (strip_tags(question), strip_tags(answer))
        for question, answer in re.findall(
            r"<details(?:\s+[^>]*)?>\s*<summary>(.*?)</summary>\s*<p>(.*?)</p>\s*</details>",
            source, flags=re.I | re.S,
        )
    ]


def reference_sets(source: str) -> dict[str, list[str]]:
    assets: list[str] = []
    links: list[str] = []
    external: list[str] = []
    for match in re.finditer(r"<(?:a|link|script|img|source)\b[^>]*>", source, flags=re.I | re.S):
        attrs = tag_attributes(match.group(0))
        for key in ("href", "src", "srcset"):
            value = clean(attrs.get(key))
            if not value:
                continue
            values = [value]
            if key == "srcset":
                values = [clean(item.split()[0]) for item in value.split(",") if clean(item)]
            for item in values:
                parsed = urllib.parse.urlsplit(item)
                if parsed.scheme in {"http", "https"}:
                    external.append(item)
                elif parsed.scheme in {"tel", "mailto"}:
                    external.append(item)
                elif parsed.scheme:
                    external.append(item)
                elif item.startswith("#"):
                    continue
                elif key in {"src", "srcset"} or item.lower().split("?", 1)[0].endswith(
                    (".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico")
                ):
                    assets.append(item)
                else:
                    links.append(item)
    return {
        "assets": sorted(dict.fromkeys(assets)),
        "links": sorted(dict.fromkeys(links)),
        "external": sorted(dict.fromkeys(external)),
    }


def page_record(path: Path, source: str | None = None) -> dict[str, object]:
    raw = path.read_bytes() if source is None else source.encode("utf-8")
    source = raw.decode("utf-8") if source is None else source
    nodes, graph_errors = parse_graph(source)
    article = first_node(nodes, "Article")
    return {
        "path": relative(path),
        "url": page_url(path),
        "sha256": sha256_bytes(raw),
        "lf_sha256": sha256_bytes(raw.replace(b"\r\n", b"\n")),
        "title": first_tag_value(source, "title"),
        "h1": first_tag_value(source, "h1"),
        "description": meta_content(source, name="description"),
        "canonical": canonical_href(source),
        "date_published": clean(article.get("datePublished")),
        "schema_types": sorted(set().union(*(node_types(node) for node in nodes))),
        "graph_errors": graph_errors,
        "facts_sha256": sha256_bytes(canonical_json(stable_fact_payload(nodes))),
        "references": reference_sets(source),
    }


def archive_contract(archive: Path) -> dict[str, object]:
    archive = archive.expanduser().resolve()
    if not archive.is_file():
        raise GateError(f"archive missing: {archive}")
    if sha256_file(archive) != EXPECTED_ARCHIVE_SHA256:
        raise GateError(f"archive SHA-256 mismatch: {sha256_file(archive)}")
    result: dict[str, object] = {
        "archive_sha256": EXPECTED_ARCHIVE_SHA256,
        "workbooks": {},
        "source_hashes": {},
    }
    try:
        with ZipFile(archive) as outer:
            infos = outer.infolist()
            names = [info.filename.replace("\\", "/") for info in infos]
            if set(names) != set(EXPECTED_WORKBOOK_SHA256) or len(names) != len(EXPECTED_WORKBOOK_SHA256):
                raise GateError(f"outer member set mismatch: {names}")
            if any(info.flag_bits & 1 for info in infos):
                raise GateError("encrypted outer ZIP member")
            for info in infos:
                name = info.filename.replace("\\", "/")
                pure = PurePosixPath(name)
                if pure.is_absolute() or ".." in pure.parts or len(pure.parts) != 1:
                    raise GateError(f"unsafe outer member: {name}")
                raw = outer.read(info)
                actual = sha256_bytes(raw)
                if actual != EXPECTED_WORKBOOK_SHA256[name]:
                    raise GateError(f"workbook SHA-256 mismatch {name}: {actual}")
                values = workbook_values(raw, name)
                profile = next(key for key, (workbook, _) in PROFILE_SPECS.items() if workbook == name)
                hashes = [sha256_bytes(value.encode("utf-8")) for value in values]
                result["workbooks"][name] = {
                    "sha256": actual,
                    "rows": len(values),
                    "unique_rows": len(set(values)),
                    "source_set_sha256": set_sha(hashes),
                }
                result["source_hashes"][profile] = sorted(hashes)
    except BadZipFile as exc:
        raise GateError(f"invalid archive: {exc}") from exc
    return result


def workbook_values(raw: bytes, name: str) -> list[str]:
    forbidden = ("vbaproject.bin", "externallinks/", "embeddings/", "activex/", "customui/", "oleobject")
    try:
        with ZipFile(io.BytesIO(raw)) as book:
            infos = book.infolist()
            names = {item.filename.replace("\\", "/") for item in infos}
            if any(item.flag_bits & 1 for item in infos):
                raise GateError(f"encrypted OOXML part: {name}")
            if any(PurePosixPath(item).is_absolute() or ".." in PurePosixPath(item).parts for item in names):
                raise GateError(f"unsafe OOXML path: {name}")
            if any(any(token in item.lower() for token in forbidden) for item in names):
                raise GateError(f"active/external OOXML content: {name}")
            required = {
                "xl/workbook.xml", "xl/sharedStrings.xml",
                "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml",
            }
            if not required.issubset(names):
                raise GateError(f"missing OOXML parts {name}: {sorted(required - names)}")
            if sum(item.file_size for item in infos) > 80_000_000:
                raise GateError(f"OOXML expansion too large: {name}")
            workbook = ET.fromstring(book.read("xl/workbook.xml"))
            sheets = workbook.find(NS_MAIN + "sheets")
            if sheets is None or [item.attrib.get("name") for item in sheets] != ["A열_텍스트파일", "Sheet1"]:
                raise GateError(f"sheet contract mismatch: {name}")
            shared_root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            shared = [
                "".join(node.text or "" for node in item.iter(NS_MAIN + "t"))
                for item in shared_root.findall(NS_MAIN + "si")
            ]
            first = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
            second = ET.fromstring(book.read("xl/worksheets/sheet2.xml"))
            dimension = first.find(NS_MAIN + "dimension")
            if dimension is None or dimension.attrib.get("ref") != "A1:A372":
                raise GateError(f"first-sheet dimension mismatch: {name}")
            if first.findall(".//" + NS_MAIN + "f") or second.findall(".//" + NS_MAIN + "f"):
                raise GateError(f"formula found: {name}")
            if first.findall(".//" + NS_MAIN + "hyperlink") or second.findall(".//" + NS_MAIN + "hyperlink"):
                raise GateError(f"workbook hyperlink found: {name}")
            refs: list[str] = []
            values: list[str] = []
            for cell in first.findall(".//" + NS_MAIN + "c"):
                ref = cell.attrib.get("r", "")
                value = cell.find(NS_MAIN + "v")
                if value is None or cell.attrib.get("t") != "s":
                    raise GateError(f"cell contract mismatch {name}:{ref}")
                index = int(value.text or "-1")
                if not 0 <= index < len(shared):
                    raise GateError(f"shared-string index mismatch {name}:{ref}")
                refs.append(ref)
                values.append(shared[index].replace("_x000D_", ""))
            if refs != [f"A{row}" for row in range(1, 373)] or values[0] != "사용자 지정":
                raise GateError(f"row/header contract mismatch: {name}")
            manuscripts = values[1:]
            if len(manuscripts) != EXPECTED_PER_PROFILE or len(set(manuscripts)) != EXPECTED_PER_PROFILE:
                raise GateError(f"manuscript cardinality/uniqueness mismatch: {name}")
            return manuscripts
    except (BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
        raise GateError(f"invalid workbook {name}: {exc}") from exc


def generator_contract() -> dict[str, object]:
    if not GENERATOR_PATH.is_file():
        raise GateError(f"generator missing: {GENERATOR_PATH}")
    source = GENERATOR_PATH.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(GENERATOR_PATH))
    functions = {
        node.name: node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }
    classes = {node.name: node for node in tree.body if isinstance(node, ast.ClassDef)}
    constants = {
        target.id: node.value.value
        for node in tree.body
        if isinstance(node, ast.Assign)
        and isinstance(node.value, ast.Constant)
        for target in node.targets
        if isinstance(target, ast.Name)
    }
    required_functions = {
        "build_plan", "assert_repeat", "freeze_bytes", "write_freeze", "verify_freeze",
        "transaction_residue", "apply_plan", "report", "main",
    }
    required_classes = {"GateError", "SourceCell", "Document", "BuildPlan"}
    missing_functions = sorted(required_functions - set(functions))
    missing_classes = sorted(required_classes - set(classes))
    flags = {
        constant.value
        for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "add_argument"
        for constant in node.args[:1]
        if isinstance(constant, ast.Constant) and isinstance(constant.value, str)
    }
    required_flags = {
        "--archive", "--common-dir", "--freeze-out", "--freeze-file",
        "--apply", "--go", "--skip-repeat",
    }
    if missing_functions or missing_classes or not required_flags.issubset(flags):
        raise GateError(
            f"generator API mismatch functions={missing_functions} classes={missing_classes} "
            f"flags={sorted(required_flags - flags)}"
        )
    expected_constants = {
        "BASE_COMMIT": BASELINE_COMMIT,
        "REVISION_DATE": REVISION_DATE,
        "EXPECTED_DETAIL_DOCUMENTS": EXPECTED_DETAIL_COUNT,
        "EXPECTED_DOCUMENT_COUNT": EXPECTED_PRODUCT_DOCUMENT_COUNT,
    }
    mismatched_constants = {
        key: {"actual": constants.get(key), "expected": value}
        for key, value in expected_constants.items()
        if constants.get(key) != value
    }
    if mismatched_constants:
        raise GateError(f"generator constant contract mismatch: {mismatched_constants}")
    if "APPLY-GO" not in source or "verify_freeze" not in source:
        raise GateError("generator apply/freeze fail-closed contract missing")
    actual = sha256_file(GENERATOR_PATH)
    if EXPECTED_GENERATOR_SHA256 != "PENDING" and actual != EXPECTED_GENERATOR_SHA256:
        raise GateError(
            f"generator SHA-256 pin mismatch: actual={actual} expected={EXPECTED_GENERATOR_SHA256}"
        )
    return {
        "path": relative(GENERATOR_PATH),
        "actual_sha256": actual,
        "expected_sha256": EXPECTED_GENERATOR_SHA256,
        "pin_status": "PENDING" if EXPECTED_GENERATOR_SHA256 == "PENDING" else "PINNED",
        "expected_candidate_sha256": EXPECTED_CANDIDATE_SHA256,
        "expected_after_manifest_sha256": EXPECTED_AFTER_MANIFEST_SHA256,
        "functions": sorted(required_functions),
        "classes": sorted(required_classes),
        "flags": sorted(required_flags),
        "constants": expected_constants,
    }


def release_toolchain_contract() -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    expected_pins = {
        GENERATOR_PATH: EXPECTED_GENERATOR_SHA256,
        CONTENT_AUDITOR_PATH: EXPECTED_CONTENT_AUDITOR_SHA256,
        TECHNICAL_AUDITOR_PATH: "SELF",
    }
    for path in (GENERATOR_PATH, CONTENT_AUDITOR_PATH, TECHNICAL_AUDITOR_PATH):
        if not path.is_file() or path.resolve().parent != (ROOT / "scripts").resolve():
            raise GateError(f"release tool missing or misplaced: {path}")
        source = path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(path))
        functions = {
            node.name for node in tree.body
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        }
        if "main" not in functions:
            raise GateError(f"release tool has no main(): {path}")
        actual = sha256_file(path)
        expected = expected_pins[path]
        if expected not in {"PENDING", "SELF"} and actual != expected:
            raise GateError(
                f"release tool SHA-256 pin mismatch {relative(path)}: actual={actual} expected={expected}"
            )
        result[relative(path)] = {
            "sha256": actual,
            "expected_sha256": expected,
            "pin_status": (
                "PENDING" if expected == "PENDING"
                else "SELF" if expected == "SELF"
                else "PINNED"
            ),
            "bytes": path.stat().st_size,
            "main": True,
        }
    return result


def sitemap_snapshot(path: Path, raw: bytes | None = None) -> dict[str, object]:
    raw = path.read_bytes() if raw is None else raw
    source = raw.decode("utf-8")
    lf_source = source.replace("\r\n", "\n")
    try:
        xml_root = ET.fromstring(source)
    except ET.ParseError as exc:
        raise GateError(f"malformed sitemap XML: {exc}") from exc
    sitemap_namespace = "http://www.sitemaps.org/schemas/sitemap/0.9"
    if xml_root.tag != f"{{{sitemap_namespace}}}urlset":
        raise GateError(f"unexpected sitemap root: {xml_root.tag}")
    xml_urls = xml_root.findall(f"{{{sitemap_namespace}}}url")
    blocks = URL_BLOCK_RE.findall(source)
    if len(xml_urls) != len(blocks) or len(xml_urls) != len(list(xml_root)):
        raise GateError(
            f"sitemap XML/byte block mismatch: xml={len(xml_urls)} blocks={len(blocks)} children={len(list(xml_root))}"
        )
    order: list[str] = []
    entries: dict[str, dict[str, str]] = {}
    for block in blocks:
        loc_match = re.search(r"<loc\b[^>]*>(.*?)</loc>", block, flags=re.I | re.S)
        if not loc_match:
            raise GateError("sitemap url block without loc")
        location = clean(loc_match.group(1))
        if location in entries:
            raise GateError(f"duplicate sitemap URL: {location}")
        modified_match = LASTMOD_RE.search(block)
        modified = clean(modified_match.group(2)) if modified_match else ""
        normalized = LASTMOD_RE.sub(r"\1{LASTMOD}\3", block, count=1)
        order.append(location)
        entries[location] = {
            "lastmod": modified,
            "raw_sha256": sha256_bytes(block.encode("utf-8")),
            "except_lastmod_sha256": sha256_bytes(normalized.encode("utf-8")),
        }
    return {
        "file_sha256": sha256_bytes(raw),
        "lf_file_sha256": sha256_bytes(lf_source.encode("utf-8")),
        "count": len(order),
        "order": order,
        "order_sha256": set_sha(f"{index:05d}\t{url}" for index, url in enumerate(order)),
        "outside_url_blocks_sha256": sha256_bytes(
            URL_BLOCK_RE.sub("{URL_BLOCK}", source).encode("utf-8")
        ),
        "outside_url_blocks_lf_sha256": sha256_bytes(
            URL_BLOCK_RE.sub("{URL_BLOCK}", lf_source).encode("utf-8")
        ),
        "entries": entries,
    }


def git_tree(commit: str) -> dict[str, dict[str, str]]:
    raw = run_git("ls-tree", "-r", "-z", "--long", commit, text=False).stdout
    result: dict[str, dict[str, str]] = {}
    for record in raw.split(b"\0"):
        if not record:
            continue
        header, name = record.split(b"\t", 1)
        mode, kind, oid, size = header.decode("ascii").split()
        path = name.decode("utf-8")
        result[path] = {"mode": mode, "type": kind, "oid": oid, "size": size}
    return result


def build_baseline(archive: Path) -> dict[str, object]:
    head = git_text("rev-parse", "HEAD")
    origin_main = git_text("rev-parse", "origin/main")
    remote = git_text("remote", "get-url", "origin")
    if head != BASELINE_COMMIT or origin_main != BASELINE_COMMIT:
        raise GateError(f"baseline commit mismatch head={head} origin/main={origin_main}")
    if remote not in EXPECTED_REMOTE_URLS:
        raise GateError(f"unexpected origin: {remote}")
    profiles = profile_detail_paths()
    if {profile: len(paths) for profile, paths in profiles.items()} != {
        profile: EXPECTED_PER_PROFILE for profile in PROFILE_SPECS
    }:
        raise GateError(f"target page count mismatch: { {k: len(v) for k, v in profiles.items()} }")
    pages = target_paths()
    if len(pages) != EXPECTED_DETAIL_COUNT or len(set(pages)) != EXPECTED_DETAIL_COUNT:
        raise GateError(f"target path cardinality mismatch: {len(pages)}")
    page_rels = {relative(path) for path in pages}
    expected_release_scope = page_rels | {"sitemap.xml"} | TOOL_PATHS
    actual_changes = changed_paths(head)
    if set(actual_changes) != expected_release_scope:
        raise GateError(
            "known-superseded Git scope mismatch: "
            f"missing={sorted(expected_release_scope-set(actual_changes))[:20]} "
            f"extra={sorted(set(actual_changes)-expected_release_scope)[:20]}"
        )
    if any(actual_changes.get(rel) != "M" for rel in page_rels | {"sitemap.xml"}):
        raise GateError("known-superseded product paths must all have Git status M")
    if any(actual_changes.get(rel) not in {"A", "M"} for rel in TOOL_PATHS):
        raise GateError("release tools must have Git status A or M")
    baseline_blobs = git_blobs(head, page_rels | {"sitemap.xml"})
    target_records = {
        rel: page_record(ROOT / rel, baseline_blobs[rel].decode("utf-8"))
        for rel in sorted(page_rels)
    }
    if len({item["url"] for item in target_records.values()}) != EXPECTED_DETAIL_COUNT:
        raise GateError("baseline target URL duplication")
    if any(semantic_url(str(item["canonical"])) != semantic_url(str(item["url"])) for item in target_records.values()):
        raise GateError("baseline canonical/route mismatch")
    tree = git_tree(head)
    release_content = set(target_records) | {"sitemap.xml"}
    immutable = {path: value for path, value in tree.items() if path not in release_content}
    immutable_worktree_sha256: dict[str, str] = {}
    for rel in immutable:
        local = ROOT / rel
        if not local.is_file():
            raise GateError(f"baseline immutable file missing: {rel}")
        immutable_worktree_sha256[rel] = sha256_file(local)
    all_html = sorted(path for path in tree if path.endswith(".html"))
    non_target_html = [path for path in all_html if path not in target_records]
    assets = [path for path in tree if path.startswith("assets/")]
    root_files = [path for path in tree if "/" not in path and path != "sitemap.xml"]
    sitemap = sitemap_snapshot(ROOT / "sitemap.xml", baseline_blobs["sitemap.xml"])
    product_worktree_sha256 = {
        rel: sha256_file(ROOT / rel) for rel in sorted(page_rels | {"sitemap.xml"})
    }
    raw_before_manifest = set_sha(
        f"{rel}\t{value}" for rel, value in product_worktree_sha256.items()
    )
    verified_git_manifest = set_sha(
        f"{rel}\t{sha256_bytes(raw)}" for rel, raw in baseline_blobs.items()
    )
    if raw_before_manifest != EXPECTED_RAW_BEFORE_MANIFEST_SHA256:
        raise GateError(
            f"raw-before manifest pin mismatch: {raw_before_manifest} != {EXPECTED_RAW_BEFORE_MANIFEST_SHA256}"
        )
    if verified_git_manifest != EXPECTED_VERIFIED_GIT_MANIFEST_SHA256:
        raise GateError(
            f"verified-Git manifest pin mismatch: {verified_git_manifest} != {EXPECTED_VERIFIED_GIT_MANIFEST_SHA256}"
        )
    target_urls = {str(item["url"]) for item in target_records.values()}
    if sitemap["count"] != EXPECTED_SITEMAP_COUNT or not target_urls.issubset(sitemap["entries"]):
        raise GateError("baseline sitemap coverage/count mismatch")
    support: dict[str, dict[str, str]] = {}
    for label, (path, expected) in EXPECTED_SOURCE_SUPPORT.items():
        if not path.is_file() or sha256_file(path) != expected:
            raise GateError(f"source support pin mismatch: {label}")
        support[label] = {"path": str(path), "sha256": expected}
    return {
        "schema_version": 1,
        "created_unix": int(time.time()),
        "root": str(ROOT),
        "site_origin": SITE_ORIGIN,
        "baseline_commit": head,
        "origin_main": origin_main,
        "origin_url": remote,
        "release_contract": {
            "detail_count": EXPECTED_DETAIL_COUNT,
            "product_document_count": EXPECTED_PRODUCT_DOCUMENT_COUNT,
            "generator_sha256": EXPECTED_GENERATOR_SHA256,
            "candidate_sha256": EXPECTED_CANDIDATE_SHA256,
            "after_manifest_sha256": EXPECTED_AFTER_MANIFEST_SHA256,
            "raw_before_manifest_sha256": EXPECTED_RAW_BEFORE_MANIFEST_SHA256,
            "verified_git_manifest_sha256": EXPECTED_VERIFIED_GIT_MANIFEST_SHA256,
            "candidate_after_manifest_sha256": EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256,
            "provenance_manifest_sha256": EXPECTED_PROVENANCE_MANIFEST_SHA256,
            "generator_before_manifest_sha256": EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256,
            "superseded_changed_documents": EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS,
            "sitemap_count": EXPECTED_SITEMAP_COUNT,
            "release_path_count": EXPECTED_RELEASE_PATH_COUNT,
            "revision_date": REVISION_DATE,
            "tool_paths": sorted(TOOL_PATHS),
            "target_path_set_sha256": set_sha(target_records),
            "target_url_set_sha256": set_sha(target_urls),
        },
        "generator": generator_contract(),
        "sources": archive_contract(archive),
        "source_support": support,
        "product_manifests": {
            "raw_before_sha256": raw_before_manifest,
            "verified_git_sha256": verified_git_manifest,
        },
        "product_worktree_sha256": product_worktree_sha256,
        "target_pages": target_records,
        "target_hubs": [relative(path) for path in hub_paths()],
        "immutable_git_blobs": immutable,
        "immutable_worktree_sha256": immutable_worktree_sha256,
        "immutable_manifests": {
            "all": set_sha(f"{path}\t{item['oid']}" for path, item in immutable.items()),
            "non_target_html": set_sha(f"{path}\t{tree[path]['oid']}" for path in non_target_html),
            "assets": set_sha(f"{path}\t{tree[path]['oid']}" for path in assets),
            "root_files": set_sha(f"{path}\t{tree[path]['oid']}" for path in root_files),
        },
        "sitemap": sitemap,
    }


def load_baseline(path: Path) -> dict[str, object]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != 1:
        raise GateError(f"unsupported baseline schema: {payload.get('schema_version')}")
    if payload.get("baseline_commit") != BASELINE_COMMIT:
        raise GateError(f"baseline provenance mismatch: {payload.get('baseline_commit')}")
    contract = payload.get("release_contract", {})
    expected = {
        "detail_count": EXPECTED_DETAIL_COUNT,
        "product_document_count": EXPECTED_PRODUCT_DOCUMENT_COUNT,
        "generator_sha256": EXPECTED_GENERATOR_SHA256,
        "candidate_sha256": EXPECTED_CANDIDATE_SHA256,
        "after_manifest_sha256": EXPECTED_AFTER_MANIFEST_SHA256,
        "raw_before_manifest_sha256": EXPECTED_RAW_BEFORE_MANIFEST_SHA256,
        "verified_git_manifest_sha256": EXPECTED_VERIFIED_GIT_MANIFEST_SHA256,
        "candidate_after_manifest_sha256": EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256,
        "provenance_manifest_sha256": EXPECTED_PROVENANCE_MANIFEST_SHA256,
        "generator_before_manifest_sha256": EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256,
        "superseded_changed_documents": EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS,
        "sitemap_count": EXPECTED_SITEMAP_COUNT,
        "release_path_count": EXPECTED_RELEASE_PATH_COUNT,
        "revision_date": REVISION_DATE,
        "tool_paths": sorted(TOOL_PATHS),
    }
    for key, value in expected.items():
        if contract.get(key) != value:
            raise GateError(f"baseline release contract mismatch {key}: {contract.get(key)!r}")
    if len(payload.get("target_pages", {})) != EXPECTED_DETAIL_COUNT:
        raise GateError("baseline target-page count mismatch")
    discovered_paths = {relative(item) for item in target_paths()}
    if set(payload["target_pages"]) != discovered_paths:
        raise GateError("baseline target path set differs from repository target set")
    if contract.get("target_path_set_sha256") != set_sha(discovered_paths):
        raise GateError("baseline target path-set seal mismatch")
    baseline_blobs = git_blobs(BASELINE_COMMIT, discovered_paths | {"sitemap.xml"})
    expected_records: dict[str, dict[str, object]] = {}
    for rel in sorted(discovered_paths):
        raw = baseline_blobs[rel]
        expected_records[rel] = page_record(ROOT / rel, raw.decode("utf-8"))
    for rel, expected_record in expected_records.items():
        stored_record = payload["target_pages"][rel]
        normalized_stored = dict(stored_record)
        normalized_stored["sha256"] = normalized_stored.get("lf_sha256")
        if normalized_stored != expected_record:
            raise GateError(f"baseline target record does not match pinned Git blob: {rel}")
    expected_urls = {str(item["url"]) for item in expected_records.values()}
    if contract.get("target_url_set_sha256") != set_sha(expected_urls):
        raise GateError("baseline target URL-set seal mismatch")
    expected_sitemap = sitemap_snapshot(
        ROOT / "sitemap.xml", baseline_blobs["sitemap.xml"]
    )
    stored_sitemap = dict(payload.get("sitemap", {}))
    stored_sitemap["file_sha256"] = stored_sitemap.get("lf_file_sha256")
    stored_sitemap["outside_url_blocks_sha256"] = stored_sitemap.get(
        "outside_url_blocks_lf_sha256"
    )
    if stored_sitemap != expected_sitemap:
        raise GateError("baseline sitemap does not match pinned Git blob after CRLF normalization")
    tree = git_tree(BASELINE_COMMIT)
    expected_immutable = {
        rel: item for rel, item in tree.items()
        if rel not in discovered_paths and rel != "sitemap.xml"
    }
    if payload.get("immutable_git_blobs") != expected_immutable:
        raise GateError("baseline immutable-tree manifest does not match pinned Git tree")
    if set(payload.get("immutable_worktree_sha256", {})) != set(expected_immutable):
        raise GateError("baseline immutable worktree-byte manifest path set mismatch")
    expected_product_manifests = {
        "raw_before_sha256": EXPECTED_RAW_BEFORE_MANIFEST_SHA256,
        "verified_git_sha256": EXPECTED_VERIFIED_GIT_MANIFEST_SHA256,
    }
    if payload.get("product_manifests") != expected_product_manifests:
        raise GateError("baseline product manifests do not match final frozen pins")
    product_worktree_sha256 = payload.get("product_worktree_sha256", {})
    expected_product_paths = discovered_paths | {"sitemap.xml"}
    if not isinstance(product_worktree_sha256, dict) or set(product_worktree_sha256) != expected_product_paths:
        raise GateError("baseline known-superseded product-byte path set mismatch")
    if set_sha(
        f"{rel}\t{value}" for rel, value in sorted(product_worktree_sha256.items())
    ) != EXPECTED_RAW_BEFORE_MANIFEST_SHA256:
        raise GateError("baseline known-superseded product-byte manifest mismatch")
    return payload


def changed_paths(base_commit: str) -> dict[str, str]:
    result: dict[str, str] = {}
    raw = run_git("diff", "--no-renames", "--name-status", "-z", base_commit, text=False).stdout
    parts = raw.split(b"\0")
    index = 0
    while index < len(parts) and parts[index]:
        status = parts[index].decode("ascii", errors="replace")
        if index + 1 >= len(parts):
            raise GateError("truncated git diff --name-status output")
        path = parts[index + 1].decode("utf-8")
        result[path] = status[0]
        index += 2
    untracked = run_git("ls-files", "--others", "--exclude-standard", "-z", text=False).stdout
    for item in untracked.split(b"\0"):
        if item:
            result[item.decode("utf-8")] = "A"
    return result


def validate_git_scope(audit: Audit, baseline: dict[str, object], phase: str) -> dict[str, object]:
    base = str(baseline["baseline_commit"])
    target = set(baseline["target_pages"])
    expected = target | {"sitemap.xml"} | TOOL_PATHS
    actual = changed_paths(base)
    actual_set = set(actual)
    missing = sorted(expected - actual_set)
    extra = sorted(actual_set - expected)
    if missing:
        audit.error("git_scope_missing", f"{len(missing)} paths; first={missing[:20]}")
    if extra:
        audit.error("git_scope_extra", f"{len(extra)} paths; first={extra[:20]}")
    if len(actual_set) != EXPECTED_RELEASE_PATH_COUNT:
        audit.error("git_scope_count", f"actual={len(actual_set)} expected={EXPECTED_RELEASE_PATH_COUNT}")
    for path in sorted(target | {"sitemap.xml"}):
        if path in actual and actual[path] != "M":
            audit.error("git_content_status", f"{path}: {actual[path]} expected M")
    for path in sorted(TOOL_PATHS):
        if path in actual and actual[path] not in {"A", "M"}:
            audit.error("git_tool_status", f"{path}: {actual[path]}")
    if phase in {"preview", "live"}:
        dirty = git_text("status", "--porcelain=v1", "--untracked-files=all")
        if dirty:
            audit.error("git_not_clean", dirty.splitlines()[:20])
    head = git_text("rev-parse", "HEAD")
    branch = git_text("branch", "--show-current")
    remote = git_text("remote", "get-url", "origin")
    if remote not in EXPECTED_REMOTE_URLS:
        audit.error("git_remote", remote)
    if phase == "preview" and branch == "main":
        audit.error("git_preview_branch", "preview must be built from a non-main branch")
    if phase == "live" and branch != "main":
        audit.error("git_live_branch", f"production audit branch={branch!r}, expected main")
    if phase == "live":
        origin_main = git_text("rev-parse", "origin/main")
        if head != origin_main:
            audit.error("git_live_origin_main", f"head={head} origin/main={origin_main}")
    return {
        "baseline": base,
        "head": head,
        "branch": branch,
        "remote": remote,
        "changed_count": len(actual),
        "changed_by_status": dict(sorted(Counter(actual.values()).items())),
        "missing": missing,
        "extra": extra,
    }


def validate_projected_git_scope(audit: Audit, baseline: dict[str, object]) -> dict[str, object]:
    base = str(baseline["baseline_commit"])
    actual = changed_paths(base)
    actual_set = set(actual)
    target = set(baseline["target_pages"])
    expected = target | {"sitemap.xml"} | TOOL_PATHS
    if actual_set != expected:
        audit.error(
            "projected_git_scope",
            f"missing={sorted(expected-actual_set)[:20]} extra={sorted(actual_set-expected)[:20]}",
        )
    if len(actual_set) != EXPECTED_RELEASE_PATH_COUNT:
        audit.error("projected_git_scope_count", f"actual={len(actual_set)} expected={EXPECTED_RELEASE_PATH_COUNT}")
    for path in sorted(target | {"sitemap.xml"}):
        if path in actual and actual[path] != "M":
            audit.error("projected_content_status", f"{path}: {actual[path]} expected M")
    for path in TOOL_PATHS:
        if path in actual and actual[path] not in {"A", "M"}:
            audit.error("projected_tool_status", f"{path}: {actual[path]}")
    head = git_text("rev-parse", "HEAD")
    origin_main = git_text("rev-parse", "origin/main")
    remote = git_text("remote", "get-url", "origin")
    if head != base or origin_main != base:
        audit.error("projected_git_baseline", f"head={head} origin/main={origin_main} expected={base}")
    if remote not in EXPECTED_REMOTE_URLS:
        audit.error("projected_git_remote", remote)
    return {
        "baseline": base,
        "head": head,
        "origin_main": origin_main,
        "remote": remote,
        "changed_count": len(actual),
        "changed_by_status": dict(sorted(Counter(actual.values()).items())),
        "missing": sorted(expected - actual_set),
        "extra": sorted(actual_set - expected),
    }


def validate_baseline_product_unchanged(
    audit: Audit,
    baseline: dict[str, object],
) -> dict[str, object]:
    expected_hashes: dict[str, str] = baseline["product_worktree_sha256"]
    rows: list[str] = []
    mismatches: list[str] = []
    for rel, expected in sorted(expected_hashes.items()):
        path = ROOT / rel
        actual = sha256_file(path) if path.is_file() else "MISSING"
        rows.append(f"{rel}\t{actual}")
        if actual != expected:
            mismatches.append(rel)
    manifest = set_sha(rows)
    if mismatches:
        audit.error("projected_superseded_product_changed", f"{len(mismatches)}; first={mismatches[:20]}")
    if manifest != EXPECTED_RAW_BEFORE_MANIFEST_SHA256:
        audit.error(
            "projected_superseded_raw_before_manifest",
            f"actual={manifest} expected={EXPECTED_RAW_BEFORE_MANIFEST_SHA256}",
        )
    return {
        "documents": len(rows),
        "state": "KNOWN_SUPERSEDED_RELEASE",
        "mismatches": mismatches,
        "raw_before_manifest_sha256": manifest,
    }


def validate_transaction_security(audit: Audit) -> dict[str, object]:
    residue_patterns = (
        ".revise-kem-*.tmp", ".revise-kem-*.bak",
        ".revise-kem-transaction.json", ".revise-kem.lock",
    )
    residue = sorted({
        relative(path) for pattern in residue_patterns for path in ROOT.rglob(pattern)
    })
    if residue:
        audit.error("transaction_residue", residue[:20])
    release_paths = [*target_paths(), ROOT / "sitemap.xml", GENERATOR_PATH, CONTENT_AUDITOR_PATH, SCRIPT_PATH]
    symlinks = sorted(relative(path) for path in release_paths if path.is_symlink())
    if symlinks:
        audit.error("release_symlink", symlinks[:20])
    git_lock = ROOT / ".git" / "index.lock"
    if git_lock.exists():
        audit.error("git_index_lock", git_lock)
    pycache = sorted(relative(path) for path in (ROOT / "scripts").glob("__pycache__/*") if path.is_file())
    if pycache:
        audit.error("release_pycache_residue", pycache)
    return {
        "transaction_residue": residue,
        "symlinks": symlinks,
        "git_index_lock": git_lock.exists(),
        "pycache_residue": pycache,
    }


def validate_immutable_tree(audit: Audit, baseline: dict[str, object]) -> dict[str, object]:
    base = str(baseline["baseline_commit"])
    immutable: dict[str, dict[str, str]] = baseline["immutable_git_blobs"]
    changed = changed_paths(base)
    violations = sorted(path for path in immutable if path in changed)
    if violations:
        audit.error("immutable_blob_changed", f"{len(violations)} paths; first={violations[:20]}")
    missing_files = sorted(path for path in immutable if not (ROOT / path).exists())
    if missing_files:
        audit.error("immutable_file_missing", f"{len(missing_files)} paths; first={missing_files[:20]}")
    expected_worktree_hashes: dict[str, str] = baseline.get("immutable_worktree_sha256", {})
    byte_violations: list[str] = []
    for rel, expected_hash in expected_worktree_hashes.items():
        local = ROOT / rel
        if local.is_file() and sha256_file(local) != expected_hash:
            byte_violations.append(rel)
    if byte_violations:
        audit.error(
            "immutable_worktree_bytes_changed",
            f"{len(byte_violations)} paths; first={byte_violations[:20]}",
        )
    hubs = set(baseline.get("target_hubs", []))
    hub_violations = sorted(hubs & set(changed))
    if hub_violations:
        audit.error("target_hub_changed", hub_violations)
    asset_violations = sorted(path for path in changed if path.startswith("assets/"))
    if asset_violations:
        audit.error("asset_changed", asset_violations[:20])
    non_target_html = sorted(path for path in changed if path.endswith(".html") and path not in baseline["target_pages"])
    if non_target_html:
        audit.error("non_target_html_changed", non_target_html[:20])
    root_violations = sorted(
        path for path in changed if "/" not in path and path != "sitemap.xml"
    )
    if root_violations:
        audit.error("root_file_changed", root_violations)
    return {
        "immutable_count": len(immutable),
        "violations": violations,
        "byte_violations": byte_violations,
        "hub_violations": hub_violations,
        "asset_violations": asset_violations,
        "non_target_html_violations": non_target_html,
        "root_violations": root_violations,
    }


def validate_sitemap_change(audit: Audit, baseline: dict[str, object]) -> dict[str, object]:
    before = baseline["sitemap"]
    after = sitemap_snapshot(ROOT / "sitemap.xml")
    target_urls = {str(item["url"]) for item in baseline["target_pages"].values()}
    if after["count"] != EXPECTED_SITEMAP_COUNT:
        audit.error("sitemap_count", f"actual={after['count']} expected={EXPECTED_SITEMAP_COUNT}")
    if after["order"] != before["order"]:
        audit.error("sitemap_url_order", "URL order or set changed")
    if set(after["entries"]) != set(before["entries"]):
        audit.error("sitemap_url_set", "URL set changed")
    if after["outside_url_blocks_lf_sha256"] != before.get("outside_url_blocks_lf_sha256"):
        audit.error("sitemap_outside_blocks_changed", "bytes outside <url> blocks changed")
    changed_blocks: list[str] = []
    non_target_changed: list[str] = []
    target_not_lastmod_only: list[str] = []
    target_bad_date: list[str] = []
    for url, old in before["entries"].items():
        new = after["entries"].get(url)
        if not new:
            continue
        if old["raw_sha256"] != new["raw_sha256"]:
            changed_blocks.append(url)
        if url in target_urls:
            if old["except_lastmod_sha256"] != new["except_lastmod_sha256"]:
                target_not_lastmod_only.append(url)
            if new["lastmod"] != REVISION_DATE:
                target_bad_date.append(url)
        elif old["raw_sha256"] != new["raw_sha256"]:
            non_target_changed.append(url)
    if set(changed_blocks) != target_urls:
        audit.error(
            "sitemap_changed_blocks",
            f"changed={len(changed_blocks)} expected={len(target_urls)} "
            f"missing={len(target_urls-set(changed_blocks))} extra={len(set(changed_blocks)-target_urls)}",
        )
    if non_target_changed:
        audit.error("sitemap_non_target_raw_changed", f"{len(non_target_changed)}; first={non_target_changed[:10]}")
    if target_not_lastmod_only:
        audit.error("sitemap_target_not_lastmod_only", f"{len(target_not_lastmod_only)}; first={target_not_lastmod_only[:10]}")
    if target_bad_date:
        audit.error("sitemap_target_lastmod", f"{len(target_bad_date)}; first={target_bad_date[:10]}")
    if after["file_sha256"] == before["file_sha256"]:
        audit.error("sitemap_file_unchanged", "sitemap.xml must change")
    return {
        "count": after["count"],
        "changed_blocks": len(changed_blocks),
        "non_target_changed": len(non_target_changed),
        "target_bad_date": len(target_bad_date),
        "order_sha256": after["order_sha256"],
    }


def resolve_local_reference(page: Path, value: str) -> Path | None:
    parsed = urllib.parse.urlsplit(html.unescape(value))
    if parsed.scheme or parsed.netloc or value.startswith("#"):
        return None
    decoded = urllib.parse.unquote(parsed.path)
    if not decoded:
        return None
    if decoded.startswith("/"):
        candidate = ROOT / decoded.lstrip("/")
    else:
        candidate = page.parent / decoded
    if decoded.endswith("/"):
        candidate = candidate / "index.html"
    elif not candidate.suffix:
        candidate = candidate / "index.html"
    try:
        candidate = candidate.resolve()
    except OSError as exc:
        raise GateError(f"cannot resolve reference {page}: {value}: {exc}") from exc
    if candidate != ROOT and ROOT not in candidate.parents:
        raise GateError(f"reference escapes repository {page}: {value}")
    return candidate


def validate_page(
    path: Path,
    baseline_record: dict[str, object],
    resource_exists: dict[Path, bool],
) -> tuple[list[tuple[str, str]], dict[str, object]]:
    errors: list[tuple[str, str]] = []
    source = path.read_text(encoding="utf-8")
    rel = relative(path)
    expected_url = str(baseline_record["url"])
    title = first_tag_value(source, "title")
    h1 = first_tag_value(source, "h1")
    description = meta_content(source, name="description")
    canonical = canonical_href(source)
    if not title or not h1 or not description:
        errors.append(("page_metadata", rel))
    html_tag = re.search(r"<html\b[^>]*>", source, flags=re.I)
    if not html_tag or tag_attributes(html_tag.group(0)).get("lang", "").lower() not in {"ko", "ko-kr"}:
        errors.append(("page_language", rel))
    if not re.search(r"<meta\b[^>]*\bcharset=[\"']?utf-8[\"']?[^>]*>", source, flags=re.I):
        errors.append(("page_charset", rel))
    if not meta_content(source, name="viewport"):
        errors.append(("page_viewport", rel))
    robots = {item.strip().lower() for item in meta_content(source, name="robots").split(",")}
    if not {"index", "follow"}.issubset(robots):
        errors.append(("page_robots", f"{rel}: {sorted(robots)}"))
    if meta_content(source, prop="og:type").lower() != "article":
        errors.append(("page_og_type", rel))
    if meta_content(source, prop="og:locale") != "ko_KR":
        errors.append(("page_og_locale", rel))
    if meta_content(source, prop="og:title") != title:
        errors.append(("page_og_title", rel))
    if meta_content(source, prop="og:description") != description:
        errors.append(("page_og_description", rel))
    if semantic_url(meta_content(source, prop="og:url")) != semantic_url(expected_url):
        errors.append(("page_og_url", rel))
    if semantic_url(canonical) != semantic_url(expected_url):
        errors.append(("page_canonical", f"{rel}: {canonical!r}"))
    canonical_count = sum(
        "canonical" in tag_attributes(match.group(0)).get("rel", "").lower().split()
        for match in re.finditer(r"<link\b[^>]*>", source, flags=re.I | re.S)
    )
    if canonical_count != 1:
        errors.append(("page_canonical_count", f"{rel}: {canonical_count}"))
    h1_count = len(re.findall(r"<h1\b", source, flags=re.I))
    if h1_count != 1:
        errors.append(("page_h1_count", f"{rel}: {h1_count}"))
    if "_x000D_" in source or "�" in source or "\x00" in source:
        errors.append(("page_encoding_marker", rel))
    if f'data-revision="composite-{REVISION_DATE}"' not in source:
        errors.append(("page_revision_marker", rel))
    row_match = re.search(r'data-source-row="(\d+)"', source)
    source_match = re.search(r'data-source-sha256="([0-9a-f]{64})"', source)
    if not row_match or not source_match:
        errors.append(("page_source_hook", rel))
    nodes, graph_errors = parse_graph(source)
    for detail in graph_errors:
        errors.append(("schema_json", f"{rel}: {detail}"))
    types = set().union(*(node_types(node) for node in nodes)) if nodes else set()
    missing = REQUIRED_SCHEMA_TYPES - types
    forbidden = FORBIDDEN_SCHEMA_TYPES & types
    if missing:
        errors.append(("schema_types_missing", f"{rel}: {sorted(missing)}"))
    if forbidden:
        errors.append(("schema_types_forbidden", f"{rel}: {sorted(forbidden)}"))
    ids = [clean(node.get("@id")) for node in nodes if clean(node.get("@id"))]
    if len(ids) != len(set(ids)):
        errors.append(("schema_duplicate_id", rel))
    site_host = urllib.parse.urlsplit(SITE_ORIGIN).hostname
    for value in sorted(set(json_strings(nodes))):
        parsed_value = urllib.parse.urlsplit(value)
        if parsed_value.scheme == "http":
            errors.append(("schema_insecure_url", f"{rel}: {value}"))
        if parsed_value.hostname == site_host:
            local = resolve_local_reference(path, parsed_value.path or "/")
            if local is not None:
                exists = resource_exists.setdefault(local, local.is_file())
                if not exists:
                    errors.append(("schema_broken_internal_url", f"{rel}: {value}"))
    webpage = first_node(nodes, "WebPage")
    article = first_node(nodes, "Article")
    faq = first_node(nodes, "FAQPage")
    if semantic_url(str(webpage.get("url", ""))) != semantic_url(expected_url):
        errors.append(("schema_webpage_url", rel))
    if clean(article.get("datePublished")) != clean(baseline_record.get("date_published")):
        errors.append(("schema_date_published", rel))
    if clean(article.get("dateModified")) != REVISION_DATE:
        errors.append(("schema_date_modified", rel))
    main_page_id = article.get("mainEntityOfPage", {})
    main_page_id = clean(main_page_id.get("@id")) if isinstance(main_page_id, dict) else clean(main_page_id)
    if main_page_id and main_page_id != clean(webpage.get("@id")):
        errors.append(("schema_article_webpage_link", rel))
    visible = visible_faq(source)
    schema_faq = [
        (clean(item.get("name")), clean(item.get("acceptedAnswer", {}).get("text")))
        for item in faq.get("mainEntity", []) if isinstance(item, dict)
    ]
    if len(visible) != 4 or visible != schema_faq:
        errors.append(("faq_visible_schema_parity", f"{rel}: visible={len(visible)} schema={len(schema_faq)}"))
    fact_hash = sha256_bytes(canonical_json(stable_fact_payload(nodes)))
    if fact_hash != baseline_record.get("facts_sha256"):
        errors.append(("locked_fact_schema_changed", rel))
    references = reference_sets(source)
    for key in ("assets", "links", "external"):
        if references[key] != baseline_record.get("references", {}).get(key, []):
            errors.append((f"locked_{key}_changed", rel))
    for value in references["assets"] + references["links"]:
        local = resolve_local_reference(path, value)
        if local is not None:
            exists = resource_exists.setdefault(local, local.is_file())
            if not exists:
                errors.append(("broken_local_reference", f"{rel}: {value}"))
    for match in re.finditer(r"<(?:a|link|script|img|source)\b[^>]*>", source, flags=re.I | re.S):
        attrs = tag_attributes(match.group(0))
        for key in ("href", "src"):
            value = clean(attrs.get(key))
            if not value:
                continue
            parsed = urllib.parse.urlsplit(value)
            if parsed.scheme == "http":
                errors.append(("mixed_content", f"{rel}: {value}"))
            if parsed.scheme.lower() == "javascript":
                errors.append(("javascript_url", f"{rel}: {value}"))
        if attrs.get("target", "").lower() == "_blank":
            rel_tokens = set(attrs.get("rel", "").lower().split())
            if "noopener" not in rel_tokens:
                errors.append(("blank_without_noopener", rel))
    return errors, {
        "path": rel,
        "title": title,
        "h1": h1,
        "description": description,
        "canonical": canonical,
        "source_row": int(row_match.group(1)) if row_match else None,
        "source_sha256": source_match.group(1) if source_match else "",
        "sha256": sha256_file(path),
        "schema_types": sorted(types),
    }


def validate_candidate_pages(
    audit: Audit,
    baseline: dict[str, object],
    source_contract: dict[str, object],
) -> dict[str, object]:
    target_records: dict[str, dict[str, object]] = baseline["target_pages"]
    observations: list[dict[str, object]] = []
    resource_exists: dict[Path, bool] = {}
    for rel, old in sorted(target_records.items()):
        path = ROOT / rel
        if not path.is_file():
            audit.error("target_page_missing", rel)
            continue
        try:
            errors, observation = validate_page(path, old, resource_exists)
        except Exception as exc:
            audit.error("page_validation_exception", f"{rel}: {type(exc).__name__}: {exc}")
            continue
        for code, detail in errors:
            audit.error(code, detail)
        if observation["sha256"] == old["sha256"]:
            audit.error("target_page_unchanged", rel)
        observations.append(observation)
    for key in ("title", "h1", "description", "canonical"):
        values = [clean(item[key]) for item in observations]
        if len(values) != EXPECTED_DETAIL_COUNT or len(set(values)) != EXPECTED_DETAIL_COUNT:
            audit.error(f"candidate_{key}_uniqueness", f"count={len(values)} unique={len(set(values))}")
    for profile, (_, slug) in PROFILE_SPECS.items():
        scoped = [item for item in observations if f"과목별학원/{slug}/" in str(item["path"])]
        rows = [item["source_row"] for item in scoped]
        hashes = [str(item["source_sha256"]) for item in scoped]
        expected_hashes = source_contract["source_hashes"][profile]
        valid_rows = [row for row in rows if isinstance(row, int)]
        if len(valid_rows) != len(rows) or sorted(valid_rows) != list(range(2, 373)):
            audit.error(
                "source_row_bijection",
                f"{profile}: count={len(rows)} valid={len(valid_rows)} unique={len(set(valid_rows))}",
            )
        if sorted(hashes) != sorted(expected_hashes):
            audit.error("source_hash_bijection", f"{profile}: source hash set mismatch")
    return {
        "pages": len(observations),
        "unique_titles": len({item["title"] for item in observations}),
        "unique_h1": len({item["h1"] for item in observations}),
        "unique_descriptions": len({item["description"] for item in observations}),
        "unique_canonicals": len({item["canonical"] for item in observations}),
        "resource_paths_checked": len(resource_exists),
        "missing_resources": sum(not exists for exists in resource_exists.values()),
        "candidate_manifest_sha256": set_sha(
            f"{item['path']}\t{item['sha256']}" for item in observations
        ),
    }


def generator_candidate_sha256(payload: dict[str, object]) -> str:
    digest = hashlib.sha256()
    documents = payload.get("documents", [])
    source_manifest = payload.get("source_manifest", {})
    if not isinstance(documents, list) or not isinstance(source_manifest, dict):
        raise GateError("freeze documents/source_manifest type mismatch")
    for document in sorted(documents, key=lambda item: str(item.get("path", ""))):
        path = str(document.get("path", ""))
        after = str(document.get("after_sha256", ""))
        if not re.fullmatch(r"[0-9a-f]{64}", after):
            raise GateError(f"invalid freeze after hash: {path}")
        digest.update(path.encode("utf-8") + b"\0" + bytes.fromhex(after))
    for key, value in sorted((str(key), str(value)) for key, value in source_manifest.items()):
        if not re.fullmatch(r"(?:[0-9a-f]{40}|[0-9a-f]{64})", value):
            raise GateError(f"invalid source manifest hash: {key}")
        digest.update(key.encode("utf-8") + b"\0" + value.encode("ascii"))
    return digest.hexdigest()


def generator_document_manifest_sha256(payload: dict[str, object], hash_field: str) -> str:
    documents = payload.get("documents", [])
    if not isinstance(documents, list):
        raise GateError("freeze documents type mismatch")
    digest = hashlib.sha256()
    for document in sorted(documents, key=lambda item: str(item.get("path", ""))):
        if not isinstance(document, dict):
            raise GateError(f"freeze document must be object: {type(document).__name__}")
        path = str(document.get("path", ""))
        value = str(document.get(hash_field, ""))
        if not path or not re.fullmatch(r"[0-9a-f]{64}", value):
            raise GateError(f"invalid {hash_field} for document: {path}")
        digest.update(path.encode("utf-8") + b"\0" + bytes.fromhex(value))
    return digest.hexdigest()


def expected_source_manifest() -> dict[str, str]:
    result = {"archive": EXPECTED_ARCHIVE_SHA256, "base_commit": BASELINE_COMMIT}
    result.update({f"workbook:{name}": value for name, value in EXPECTED_WORKBOOK_SHA256.items()})
    result.update({label: expected for label, (_, expected) in EXPECTED_SOURCE_SUPPORT.items()})
    return result


def validate_freeze(
    audit: Audit,
    path: Path,
    baseline: dict[str, object],
    source_contract: dict[str, object],
    repeat_path: Path | None,
    projected: bool = False,
) -> dict[str, object]:
    raw = path.read_bytes()
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise GateError(f"invalid freeze JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise GateError(f"freeze root must be an object, got {type(payload).__name__}")
    if raw != canonical_json(payload):
        audit.error("freeze_noncanonical_bytes", "freeze is not canonical generator JSON bytes")
    expected_freeze_keys = {
        "version", "generator_sha256", "candidate_sha256", "source_manifest", "documents",
    }
    if set(payload) != expected_freeze_keys:
        audit.error("freeze_key_set", f"actual={sorted(payload)} expected={sorted(expected_freeze_keys)}")
    if payload.get("version") != 1:
        audit.error("freeze_version", payload.get("version"))
    actual_generator = sha256_file(GENERATOR_PATH)
    freeze_generator = clean(payload.get("generator_sha256"))
    if EXPECTED_GENERATOR_SHA256 == "PENDING":
        audit.error("generator_pin_pending", "set EXPECTED_GENERATOR_SHA256 after final generator freeze")
    elif actual_generator != EXPECTED_GENERATOR_SHA256:
        audit.error("generator_pin_mismatch", f"actual={actual_generator} expected={EXPECTED_GENERATOR_SHA256}")
    if freeze_generator != actual_generator:
        audit.error("freeze_generator_hash", f"freeze={freeze_generator} actual={actual_generator}")
    expected_sources = expected_source_manifest()
    if payload.get("source_manifest") != expected_sources:
        audit.error("freeze_source_manifest", "freeze source pins differ from independent pins")
    documents = payload.get("documents", [])
    if not isinstance(documents, list):
        audit.error("freeze_documents_type", type(documents).__name__)
        documents = []
    by_path: dict[str, dict[str, str]] = {}
    duplicate_paths: list[str] = []
    for item in documents:
        if not isinstance(item, dict):
            audit.error("freeze_document_type", type(item).__name__)
            continue
        expected_document_keys = {"path", "before_sha256", "after_sha256", "source_sha256"}
        if set(item) != expected_document_keys:
            audit.error("freeze_document_key_set", f"path={item.get('path')}: {sorted(item)}")
        rel = clean(item.get("path"))
        if rel in by_path:
            duplicate_paths.append(rel)
        by_path[rel] = {key: clean(item.get(key)) for key in (
            "path", "before_sha256", "after_sha256", "source_sha256",
        )}
    expected_paths = set(baseline["target_pages"]) | {"sitemap.xml"}
    expected_before_hashes: dict[str, str] = baseline.get("product_worktree_sha256", {})
    if set(expected_before_hashes) != expected_paths:
        audit.error("freeze_superseded_before_path_set", "baseline superseded-byte map differs from product scope")
    if set(by_path) != expected_paths or len(documents) != EXPECTED_PRODUCT_DOCUMENT_COUNT:
        audit.error(
            "freeze_path_set",
            f"documents={len(documents)} unique={len(by_path)} missing={len(expected_paths-set(by_path))} "
            f"extra={len(set(by_path)-expected_paths)}",
        )
    if duplicate_paths:
        audit.error("freeze_duplicate_path", duplicate_paths[:20])
    reverse_rows: list[str] = []
    after_rows: list[str] = []
    source_by_profile: dict[str, list[str]] = defaultdict(list)
    unchanged: list[str] = []
    changed: list[str] = []
    for rel, item in sorted(by_path.items()):
        if rel not in expected_paths:
            continue
        current_path = ROOT / rel
        current_hash = sha256_file(current_path) if current_path.is_file() else "MISSING"
        before = item["before_sha256"]
        after = item["after_sha256"]
        expected_before = expected_before_hashes.get(rel, "MISSING")
        if before != expected_before:
            audit.error("freeze_superseded_before_hash", rel)
        if projected and current_hash != expected_before:
            audit.error(
                "freeze_projected_worktree_state",
                f"{rel}: current={current_hash} expected_superseded={expected_before}",
            )
        if not projected and after != current_hash:
            audit.error("freeze_after_hash", f"{rel}: freeze={after} current={current_hash}")
        if before == after:
            unchanged.append(rel)
        else:
            changed.append(rel)
        reverse_rows.append(f"{rel}\t{before}")
        after_rows.append(f"{rel}\t{after}")
        if rel == "sitemap.xml":
            if item["source_sha256"] != before:
                audit.error("freeze_sitemap_source_hash", f"freeze={item['source_sha256']} before={before}")
            continue
        for profile, (_, slug) in PROFILE_SPECS.items():
            if rel.startswith(f"과목별학원/{slug}/"):
                source_by_profile[profile].append(item["source_sha256"])
                break
    expected_unchanged = EXPECTED_PRODUCT_DOCUMENT_COUNT - EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS
    if len(changed) != EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS or len(unchanged) != expected_unchanged:
        audit.error(
            "freeze_superseded_transition_count",
            f"changed={len(changed)} expected={EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS}; "
            f"unchanged={len(unchanged)} expected={expected_unchanged}",
        )
    for profile in PROFILE_SPECS:
        expected_hashes = source_contract["source_hashes"][profile]
        if sorted(source_by_profile[profile]) != sorted(expected_hashes):
            audit.error("freeze_source_bijection", profile)
    calculated = generator_candidate_sha256(payload)
    after_manifest = generator_document_manifest_sha256(payload, "after_sha256")
    before_manifest = generator_document_manifest_sha256(payload, "before_sha256")
    if calculated != payload.get("candidate_sha256"):
        audit.error("freeze_candidate_sha256", f"calculated={calculated} freeze={payload.get('candidate_sha256')}")
    if EXPECTED_CANDIDATE_SHA256 == "PENDING":
        audit.error("candidate_pin_pending", "set EXPECTED_CANDIDATE_SHA256 after final repeat build")
    elif calculated != EXPECTED_CANDIDATE_SHA256:
        audit.error(
            "candidate_pin_mismatch",
            f"calculated={calculated} expected={EXPECTED_CANDIDATE_SHA256}",
        )
    if after_manifest != EXPECTED_AFTER_MANIFEST_SHA256:
        audit.error(
            "after_manifest_pin_mismatch",
            f"calculated={after_manifest} expected={EXPECTED_AFTER_MANIFEST_SHA256}",
        )
    if before_manifest != EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256:
        audit.error(
            "before_manifest_pin_mismatch",
            f"calculated={before_manifest} expected={EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256}",
        )
    if repeat_path:
        if repeat_path.resolve() == path.resolve():
            audit.error("generator_repeat_not_independent", "freeze and repeat paths are identical")
        repeated = repeat_path.read_bytes()
        if repeated != raw:
            audit.error("generator_repeat_freeze", f"first={sha256_bytes(raw)} repeat={sha256_bytes(repeated)}")
    else:
        audit.error("generator_repeat_evidence_missing", "provide --repeat-freeze-file from an independent dry run")
    superseded_reverse_manifest = set_sha(
        f"{rel}\t{value}" for rel, value in sorted(expected_before_hashes.items())
    )
    reverse_manifest = set_sha(reverse_rows)
    if reverse_manifest != superseded_reverse_manifest:
        audit.error(
            "freeze_reverse_manifest",
            f"reverse={reverse_manifest} superseded={superseded_reverse_manifest}",
        )
    if reverse_manifest != EXPECTED_RAW_BEFORE_MANIFEST_SHA256:
        audit.error(
            "raw_before_manifest_pin_mismatch",
            f"calculated={reverse_manifest} expected={EXPECTED_RAW_BEFORE_MANIFEST_SHA256}",
        )
    candidate_document_manifest = set_sha(after_rows)
    if candidate_document_manifest != EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256:
        audit.error(
            "candidate_after_manifest_pin_mismatch",
            f"calculated={candidate_document_manifest} expected={EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256}",
        )
    return {
        "path": str(path),
        "sha256": sha256_bytes(raw),
        "candidate_sha256": calculated,
        "after_manifest_sha256": after_manifest,
        "before_manifest_sha256": before_manifest,
        "documents": len(documents),
        "detail_documents": sum(path != "sitemap.xml" for path in by_path),
        "sitemap_documents": int("sitemap.xml" in by_path),
        "changed_documents": len(changed),
        "unchanged_documents": len(unchanged),
        "generator_sha256": actual_generator,
        "generator_pin": EXPECTED_GENERATOR_SHA256,
        "repeat_evidence": str(repeat_path) if repeat_path else None,
        "repeat_equal": bool(repeat_path and repeat_path.read_bytes() == raw),
        "projected": projected,
        "reverse_manifest_sha256": reverse_manifest,
        "candidate_document_manifest_sha256": candidate_document_manifest,
    }


def python_subprocess_environment() -> dict[str, str]:
    """Make captured Python JSON independent of the Windows console code page."""
    environment = os.environ.copy()
    environment["PYTHONIOENCODING"] = "utf-8"
    environment["PYTHONUTF8"] = "1"
    return environment


def parse_json_output(stdout: str) -> dict[str, object]:
    decoder = json.JSONDecoder()
    candidates: list[tuple[int, int, dict[str, object]]] = []
    for index, character in enumerate(stdout):
        if character != "{":
            continue
        try:
            value, length = decoder.raw_decode(stdout[index:])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            candidates.append((index + length, length, value))
    if not candidates:
        raise GateError("command produced no JSON object")
    # The subprocess may print progress before its final report.  A nested JSON
    # object begins later than its enclosing report, so selecting the last
    # appended candidate can silently return only the final nested dictionary.
    # The complete final report is the dictionary whose decoded span reaches
    # farthest into stdout; span length breaks the unlikely equal-end tie.
    return max(candidates, key=lambda candidate: (candidate[0], candidate[1]))[2]


def run_generator_dry(archive: Path, common_dir: Path, timeout: int) -> dict[str, object]:
    command = [
        sys.executable, "-B", str(GENERATOR_PATH),
        "--archive", str(archive), "--common-dir", str(common_dir), "--skip-repeat",
    ]
    result = subprocess.run(
        command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="strict", timeout=timeout,
        env=python_subprocess_environment(),
    )
    if result.returncode:
        raise GateError(f"generator dry-run failed rc={result.returncode}: {result.stderr[-2000:]}")
    payload = parse_json_output(result.stdout)
    if payload.get("status") != "PASS" or payload.get("mode") != "dry-run":
        raise GateError(f"generator dry-run payload mismatch: {payload}")
    return payload


def run_content_auditor(
    mode: str,
    archive: Path,
    common_dir: Path,
    timeout: int,
) -> dict[str, object]:
    command = [
        sys.executable, "-B", str(CONTENT_AUDITOR_PATH),
        "--mode", mode,
        "--generator", str(GENERATOR_PATH),
        "--archive", str(archive),
        "--common-dir", str(common_dir),
        "--generator-pin", EXPECTED_GENERATOR_SHA256,
    ]
    result = subprocess.run(
        command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="strict", timeout=timeout,
        env=python_subprocess_environment(),
    )
    if result.returncode:
        raise GateError(
            f"content auditor failed rc={result.returncode}: {result.stderr[-3000:]}\n{result.stdout[-3000:]}"
        )
    return parse_json_output(result.stdout)


def nested_value(payload: object, *keys: str) -> object:
    value = payload
    for key in keys:
        if not isinstance(value, dict):
            return None
        value = value.get(key)
    return value


def validate_content_auditor(
    audit: Audit,
    mode: str,
    archive: Path,
    common_dir: Path,
    timeout: int,
) -> dict[str, object]:
    product_paths = [*target_paths(), ROOT / "sitemap.xml"]
    before_product = set_sha(f"{relative(path)}\t{sha256_file(path)}" for path in product_paths)
    before_generator = sha256_file(GENERATOR_PATH)
    before_content = sha256_file(CONTENT_AUDITOR_PATH)
    payload = run_content_auditor(mode, archive, common_dir, timeout)
    after_product = set_sha(f"{relative(path)}\t{sha256_file(path)}" for path in product_paths)
    after_generator = sha256_file(GENERATOR_PATH)
    after_content = sha256_file(CONTENT_AUDITOR_PATH)
    if before_product != after_product:
        audit.error("content_auditor_mutated_product", f"before={before_product} after={after_product}")
    if before_generator != after_generator or before_generator != EXPECTED_GENERATOR_SHA256:
        audit.error(
            "content_auditor_generator_stability",
            f"before={before_generator} after={after_generator} expected={EXPECTED_GENERATOR_SHA256}",
        )
    if before_content != after_content or before_content != EXPECTED_CONTENT_AUDITOR_SHA256:
        audit.error(
            "content_auditor_self_pin",
            f"before={before_content} after={after_content} expected={EXPECTED_CONTENT_AUDITOR_SHA256}",
        )
    if payload.get("status") != "PASS" or payload.get("mode") != mode or payload.get("read_only") is not True:
        audit.error(
            "content_auditor_status",
            f"status={payload.get('status')} mode={payload.get('mode')} read_only={payload.get('read_only')}",
        )
    required_zero_findings = (
        ("contract", "findings", "error_count"),
        ("release_scope", "findings", "error_count"),
        ("projected", "findings", "error_count"),
    )
    if mode == "actual":
        required_zero_findings += (("actual", "findings", "error_count"),)
    for keys in required_zero_findings:
        if nested_value(payload, *keys) != 0:
            audit.error("content_auditor_findings", f"{'.'.join(keys)}={nested_value(payload, *keys)}")
    expected_values = {
        ("contract", "generator_pin", "observed_sha256"): EXPECTED_GENERATOR_SHA256,
        ("contract", "candidate_pin", "observed_sha256"): EXPECTED_CANDIDATE_SHA256,
        ("contract", "generator_after_manifest_pin", "observed_sha256"): EXPECTED_AFTER_MANIFEST_SHA256,
        ("projected", "generator_candidate_sha256"): EXPECTED_CANDIDATE_SHA256,
        ("projected", "verified_git_baseline_manifest_sha256"): EXPECTED_VERIFIED_GIT_MANIFEST_SHA256,
        ("projected", "candidate_after_manifest_sha256"): EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256,
        ("projected", "generator_metrics", "after_manifest"): EXPECTED_AFTER_MANIFEST_SHA256,
        ("projected", "repeat", "checked"): True,
        ("projected", "repeat", "deterministic"): True,
        ("projected", "repeat", "candidate_sha256"): EXPECTED_CANDIDATE_SHA256,
    }
    if mode == "projected":
        expected_values.update({
            ("contract", "working_product_state", "state"): "KNOWN_SUPERSEDED_RELEASE",
            ("contract", "working_product_state", "observed_before_manifest"): EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256,
            ("contract", "working_product_state", "generator_changed_documents"): EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS,
            ("contract", "working_product_state", "raw_pending_documents"): EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS,
            ("projected", "raw_worktree_pending_documents"): EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS,
            ("projected", "raw_before_manifest_sha256"): EXPECTED_RAW_BEFORE_MANIFEST_SHA256,
            ("projected", "provenance_manifest_sha256"): EXPECTED_PROVENANCE_MANIFEST_SHA256,
            ("projected", "generator_metrics", "before_manifest"): EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256,
        })
    else:
        expected_values.update({
            ("contract", "working_product_state", "state"): "CURRENT_CANDIDATE",
            ("contract", "working_product_state", "observed_before_manifest"): EXPECTED_AFTER_MANIFEST_SHA256,
            ("contract", "working_product_state", "generator_changed_documents"): 0,
            ("contract", "working_product_state", "raw_pending_documents"): 0,
            ("projected", "raw_worktree_pending_documents"): 0,
            ("projected", "raw_before_manifest_sha256"): EXPECTED_CANDIDATE_AFTER_MANIFEST_SHA256,
            ("projected", "generator_metrics", "before_manifest"): EXPECTED_AFTER_MANIFEST_SHA256,
        })
    mismatches = []
    for keys, expected in expected_values.items():
        actual = nested_value(payload, *keys)
        if actual != expected:
            mismatches.append(f"{'.'.join(keys)}={actual!r} expected={expected!r}")
    if mismatches:
        audit.error("content_auditor_manifest_pin", mismatches[:20])
    if nested_value(payload, "projected", "product_documents") != EXPECTED_PRODUCT_DOCUMENT_COUNT:
        audit.error("content_auditor_product_count", nested_value(payload, "projected", "product_documents"))
    expected_actual_state = "APPLIED"
    if nested_value(payload, "actual", "state") != expected_actual_state:
        audit.error(
            "content_auditor_actual_state",
            f"{nested_value(payload, 'actual', 'state')} expected={expected_actual_state}",
        )
    return {
        "status": payload.get("status"),
        "mode": payload.get("mode"),
        "actual_state": nested_value(payload, "actual", "state"),
        "build_seconds": nested_value(payload, "projected", "build_seconds"),
        "candidate_sha256": nested_value(payload, "projected", "generator_candidate_sha256"),
        "raw_before_manifest_sha256": nested_value(payload, "projected", "raw_before_manifest_sha256"),
        "verified_git_manifest_sha256": nested_value(payload, "projected", "verified_git_baseline_manifest_sha256"),
        "candidate_after_manifest_sha256": nested_value(payload, "projected", "candidate_after_manifest_sha256"),
        "provenance_manifest_sha256": nested_value(payload, "projected", "provenance_manifest_sha256"),
        "repeat": nested_value(payload, "projected", "repeat"),
        "product_unchanged": before_product == after_product,
    }


def validate_generator_idempotency(
    audit: Audit,
    archive: Path,
    common_dir: Path,
    timeout: int,
    expected_changed: int,
) -> dict[str, object]:
    before_status = git_text("status", "--porcelain=v1", "--untracked-files=all")
    product_paths = [*target_paths(), ROOT / "sitemap.xml"]
    before_hash = set_sha(f"{relative(path)}\t{sha256_file(path)}" for path in product_paths)
    first = audit.capture("generator_idempotency_execution", run_generator_dry, archive, common_dir, timeout)
    second = audit.capture("generator_repeat_execution", run_generator_dry, archive, common_dir, timeout)
    after_status = git_text("status", "--porcelain=v1", "--untracked-files=all")
    after_hash = set_sha(f"{relative(path)}\t{sha256_file(path)}" for path in product_paths)
    if before_status != after_status or before_hash != after_hash:
        audit.error("generator_dry_run_mutated_repo", "git status or target bytes changed")
    if first and second:
        if first.get("candidate_sha256") != second.get("candidate_sha256"):
            audit.error("generator_repeat_candidate", "two dry runs differ")
        if first.get("changed") != expected_changed or second.get("changed") != expected_changed:
            audit.error(
                "generator_changed_count",
                f"changed={first.get('changed')},{second.get('changed')} expected={expected_changed}",
            )
        for index, payload in enumerate((first, second), 1):
            if payload.get("documents") != EXPECTED_PRODUCT_DOCUMENT_COUNT:
                audit.error("generator_document_count", f"run={index}: {payload.get('documents')}")
            if payload.get("candidate_sha256") != EXPECTED_CANDIDATE_SHA256:
                audit.error(
                    "generator_candidate_pin",
                    f"run={index}: {payload.get('candidate_sha256')} expected={EXPECTED_CANDIDATE_SHA256}",
                )
            metrics = payload.get("metrics", {})
            actual_after_manifest = metrics.get("after_manifest") if isinstance(metrics, dict) else None
            if actual_after_manifest != EXPECTED_AFTER_MANIFEST_SHA256:
                audit.error(
                    "generator_after_manifest_pin",
                    f"run={index}: {actual_after_manifest} expected={EXPECTED_AFTER_MANIFEST_SHA256}",
                )
            expected_before_manifest = (
                EXPECTED_GENERATOR_BEFORE_MANIFEST_SHA256
                if expected_changed == EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS
                else EXPECTED_AFTER_MANIFEST_SHA256
            )
            actual_before_manifest = metrics.get("before_manifest") if isinstance(metrics, dict) else None
            if actual_before_manifest != expected_before_manifest:
                audit.error(
                    "generator_before_manifest_pin",
                    f"run={index}: {actual_before_manifest} expected={expected_before_manifest}",
                )
            if payload.get("source_manifest") != expected_source_manifest():
                audit.error("generator_source_manifest", f"run={index}")
            if payload.get("residue") != []:
                audit.error("generator_transaction_residue", f"run={index}: {payload.get('residue')}")
            actual_generator = sha256_file(GENERATOR_PATH)
            if payload.get("generator_sha256") != actual_generator:
                audit.error(
                    "generator_report_hash",
                    f"run={index}: report={payload.get('generator_sha256')} actual={actual_generator}",
                )
    return {
        "first": first,
        "second": second,
        "expected_changed": expected_changed,
        "repo_unchanged": before_status == after_status and before_hash == after_hash,
    }


def checked_base_url(phase: str, value: str) -> str:
    parsed = urllib.parse.urlsplit(value.rstrip("/") + "/")
    host = (parsed.hostname or "").lower()
    if parsed.scheme not in {"http", "https"}:
        raise GateError(f"unsupported base URL scheme: {value}")
    if phase == "local" and host not in {"127.0.0.1", "localhost", "::1"}:
        raise GateError(f"local phase requires loopback URL: {value}")
    if phase == "preview" and not host.endswith(".vercel.app"):
        raise GateError(f"preview phase requires a vercel.app URL: {value}")
    if phase == "live" and host != "xn--3e0bz50b1zcyxat54c.com":
        raise GateError(f"live phase requires {DISPLAY_DOMAIN}: {value}")
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), "", ""))


def deployment_url(base_url: str, canonical: str) -> str:
    parsed = urllib.parse.urlsplit(canonical)
    return base_url.rstrip("/") + parsed.path


def http_request(url: str, *, method: str = "GET", timeout: int = 30) -> tuple[int, str, dict[str, str], bytes]:
    request = urllib.request.Request(
        url,
        method=method,
        headers={"User-Agent": "new14-revised-kem-technical-audit/1.0", "Accept": "*/*"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout, context=HTTP_SSL_CONTEXT) as response:
            return (
                int(response.status),
                response.geturl(),
                {key.lower(): value for key, value in response.headers.items()},
                response.read() if method != "HEAD" else b"",
            )
    except urllib.error.HTTPError as exc:
        return int(exc.code), exc.geturl(), {key.lower(): value for key, value in exc.headers.items()}, exc.read()


def validate_http_deployment(
    audit: Audit,
    phase: str,
    base_url: str,
    baseline: dict[str, object],
    workers: int,
    timeout: int,
    full_resources: bool,
) -> dict[str, object]:
    base_url = checked_base_url(phase, base_url)
    records: dict[str, dict[str, object]] = baseline["target_pages"]
    control_paths = [ROOT / "sitemap.xml", *hub_paths()]
    deployment_relatives = set(records) | {relative(path) for path in control_paths}
    if phase == "local":
        expected_deployment_hashes = {
            rel: sha256_file(ROOT / rel) for rel in deployment_relatives
        }
    else:
        head_blobs = git_blobs(git_text("rev-parse", "HEAD"), deployment_relatives)
        expected_deployment_hashes = {
            rel: sha256_bytes(raw) for rel, raw in head_blobs.items()
        }

    def fetch_page(item: tuple[str, dict[str, object]]) -> tuple[str, list[tuple[str, str]], dict[str, object]]:
        rel, record = item
        url = deployment_url(base_url, str(record["url"]))
        errors: list[tuple[str, str]] = []
        try:
            status, final, headers, body = http_request(url, timeout=timeout)
        except Exception as exc:
            return rel, [("http_page_exception", f"{url}: {exc}")], {"url": url, "status": None}
        if status != 200:
            errors.append(("http_page_status", f"{url}: {status}"))
        if transport_url(final) != transport_url(url):
            errors.append(("http_page_redirect", f"{url} -> {final}"))
        content_type = headers.get("content-type", "")
        if "text/html" not in content_type:
            errors.append(("http_page_content_type", f"{url}: {content_type}"))
        try:
            source = body.decode("utf-8")
        except UnicodeDecodeError as exc:
            errors.append(("http_page_encoding", f"{url}: {exc}"))
            source = body.decode("utf-8", errors="replace")
        if first_tag_value(source, "title") != first_tag_value((ROOT / rel).read_text(encoding="utf-8"), "title"):
            errors.append(("http_page_title", rel))
        if first_tag_value(source, "h1") != first_tag_value((ROOT / rel).read_text(encoding="utf-8"), "h1"):
            errors.append(("http_page_h1", rel))
        if semantic_url(canonical_href(source)) != semantic_url(str(record["url"])):
            errors.append(("http_page_canonical", rel))
        if f'data-revision="composite-{REVISION_DATE}"' not in source:
            errors.append(("http_page_revision", rel))
        expected_sha = expected_deployment_hashes[rel]
        body_sha = sha256_bytes(body)
        if body_sha != expected_sha:
            errors.append(("http_page_bytes", f"{rel}: deployed={body_sha} local={expected_sha}"))
        return rel, errors, {
            "url": url, "final_url": final, "status": status,
            "bytes": len(body), "sha256": body_sha, "cache": headers.get("x-vercel-cache"),
        }

    page_results: list[dict[str, object]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(workers, 32))) as pool:
        futures = [pool.submit(fetch_page, item) for item in sorted(records.items())]
        for future in concurrent.futures.as_completed(futures):
            _, errors, result = future.result()
            page_results.append(result)
            for code, detail in errors:
                audit.error(code, detail)

    control_results: list[dict[str, object]] = []
    for path in control_paths:
        canonical = (
            SITE_ORIGIN + "/sitemap.xml"
            if path.name == "sitemap.xml"
            else page_url(path)
        )
        url = deployment_url(base_url, canonical)
        try:
            status, final, _, body = http_request(url, timeout=timeout)
            body_sha = sha256_bytes(body)
            expected_sha = expected_deployment_hashes[relative(path)]
            control_results.append({
                "path": relative(path), "url": url, "status": status,
                "final_url": final, "sha256": body_sha,
            })
            if status != 200:
                audit.error("http_control_status", f"{url}: {status}")
            if transport_url(final) != transport_url(url):
                audit.error("http_control_redirect", f"{url} -> {final}")
            if body_sha != expected_sha:
                audit.error(
                    "http_control_bytes",
                    f"{relative(path)}: deployed={body_sha} local={expected_sha}",
                )
        except Exception as exc:
            control_results.append({"path": relative(path), "url": url, "status": None, "error": str(exc)})
            audit.error("http_control_exception", f"{url}: {exc}")

    resource_results: list[dict[str, object]] = []
    if full_resources:
        resources: set[str] = set()
        for rel in records:
            refs = reference_sets((ROOT / rel).read_text(encoding="utf-8"))["assets"]
            for value in refs:
                local = resolve_local_reference(ROOT / rel, value)
                if local is not None:
                    resources.add(page_url(local) if local.name == "index.html" else SITE_ORIGIN + "/" + urllib.parse.quote(relative(local), safe="/"))

        def fetch_resource(canonical: str) -> dict[str, object]:
            url = deployment_url(base_url, canonical)
            try:
                status, final, headers, _ = http_request(url, method="HEAD", timeout=timeout)
                return {"url": url, "status": status, "final_url": final, "content_type": headers.get("content-type", "")}
            except Exception as exc:
                return {"url": url, "status": None, "error": str(exc)}

        with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, min(workers, 32))) as pool:
            resource_results = list(pool.map(fetch_resource, sorted(resources)))
        for result in resource_results:
            if result.get("status") != 200:
                audit.error("http_resource_status", f"{result.get('url')}: {result.get('status')} {result.get('error', '')}")
    return {
        "phase": phase,
        "base_url": base_url,
        "pages_checked": len(page_results),
        "page_statuses": dict(sorted(Counter(item.get("status") for item in page_results).items(), key=lambda item: str(item[0]))),
        "controls_checked": len(control_results),
        "control_failures": sum(item.get("status") != 200 for item in control_results),
        "resources_checked": len(resource_results),
        "resource_failures": sum(item.get("status") != 200 for item in resource_results),
    }


def browser_sample_routes(baseline: dict[str, object]) -> list[str]:
    paths = set(baseline["target_pages"])
    preferred = {
        "elementary": ("명일동", "개운동", "영천동", "반곡동"),
        "middle": ("명일동", "송강동", "화명동", "좌동", "개운동"),
        "high": ("명일동", "첨단", "연동", "병점동", "경산사동"),
    }
    selected: list[str] = []
    for profile, (_, slug) in PROFILE_SPECS.items():
        scoped = sorted(path for path in paths if path.startswith(f"과목별학원/{slug}/"))
        candidates = [scoped[0], scoped[len(scoped) // 2], scoped[-1]] if scoped else []
        for locality in preferred[profile]:
            match = next(
                (path for path in scoped if Path(path).parent.name == re.sub(r"\s+", "", locality)),
                None,
            )
            if match:
                candidates.append(match)
        for path in candidates:
            if path not in selected:
                selected.append(path)
    return selected


def build_browser_plan(phase: str, base_url: str, baseline: dict[str, object]) -> dict[str, object]:
    base_url = checked_base_url(phase, base_url)
    checks = []
    for rel in browser_sample_routes(baseline):
        canonical = str(baseline["target_pages"][rel]["url"])
        for viewport, width, height in (("mobile", 390, 844), ("desktop", 1440, 1000)):
            checks.append({
                "route": urllib.parse.urlsplit(canonical).path,
                "url": deployment_url(base_url, canonical),
                "viewport": viewport,
                "width": width,
                "height": height,
                "expected_canonical": canonical,
                "required_assertions": [
                    "status_200", "title_present", "single_h1", "canonical_match",
                    "header_visible", "breadcrumbs_visible", "manuscript_visible",
                    "map_loaded", "contact_dock_visible", "no_horizontal_overflow",
                    "no_console_errors", "no_failed_requests",
                ],
            })
    return {
        "schema_version": 1,
        "phase": phase,
        "base_url": base_url,
        "expected_commit": git_text("rev-parse", "HEAD"),
        "checks": checks,
    }


def validate_browser_contract_static(audit: Audit, baseline: dict[str, object]) -> dict[str, object]:
    plan = build_browser_plan("local", "http://127.0.0.1:8000", baseline)
    checks = plan["checks"]
    keys = [(item["route"], item["viewport"]) for item in checks]
    routes = {item["route"] for item in checks}
    expected_routes = {
        urllib.parse.urlsplit(str(baseline["target_pages"][rel]["url"])).path
        for rel in browser_sample_routes(baseline)
    }
    if len(keys) != len(set(keys)) or routes != expected_routes:
        audit.error(
            "browser_contract_coverage",
            f"checks={len(keys)} unique={len(set(keys))} routes={len(routes)} expected={len(expected_routes)}",
        )
    viewports = Counter(item["viewport"] for item in checks)
    if viewports != {"mobile": len(expected_routes), "desktop": len(expected_routes)}:
        audit.error("browser_contract_viewports", dict(viewports))
    required = {
        "status_200", "title_present", "single_h1", "canonical_match",
        "header_visible", "breadcrumbs_visible", "manuscript_visible", "map_loaded",
        "contact_dock_visible", "no_horizontal_overflow", "no_console_errors",
        "no_failed_requests",
    }
    if any(set(item["required_assertions"]) != required for item in checks):
        audit.error("browser_contract_assertions", "browser assertion set drifted")
    return {
        "routes": len(expected_routes),
        "checks": len(checks),
        "viewports": dict(sorted(viewports.items())),
        "plan_sha256": sha256_bytes(canonical_json(plan)),
    }


def validate_browser_evidence(
    audit: Audit,
    path: Path,
    phase: str,
    base_url: str,
    baseline: dict[str, object],
) -> dict[str, object]:
    expected = build_browser_plan(phase, base_url, baseline)
    evidence = json.loads(path.read_text(encoding="utf-8"))
    for key in ("schema_version", "phase", "base_url", "expected_commit"):
        if evidence.get(key) != expected.get(key):
            audit.error("browser_evidence_header", f"{key}: {evidence.get(key)!r} != {expected.get(key)!r}")
    expected_checks = {(item["route"], item["viewport"]): item for item in expected["checks"]}
    actual_checks: dict[tuple[str, str], dict] = {}
    for item in evidence.get("checks", []):
        if not isinstance(item, dict):
            audit.error("browser_evidence_record", type(item).__name__)
            continue
        key = (clean(item.get("route")), clean(item.get("viewport")))
        if key in actual_checks:
            audit.error("browser_evidence_duplicate", key)
        actual_checks[key] = item
    if set(actual_checks) != set(expected_checks):
        audit.error(
            "browser_evidence_coverage",
            f"missing={sorted(set(expected_checks)-set(actual_checks))[:10]} "
            f"extra={sorted(set(actual_checks)-set(expected_checks))[:10]}",
        )
    required_true = {
        "status_200", "title_present", "single_h1", "canonical_match",
        "header_visible", "breadcrumbs_visible", "manuscript_visible", "map_loaded",
        "contact_dock_visible", "no_horizontal_overflow", "no_console_errors",
        "no_failed_requests",
    }
    failures = 0
    for key, item in actual_checks.items():
        if key not in expected_checks:
            continue
        assertions = item.get("assertions", {})
        missing = sorted(name for name in required_true if assertions.get(name) is not True)
        screenshot = clean(item.get("screenshot_sha256"))
        if missing or not re.fullmatch(r"[0-9a-f]{64}", screenshot):
            failures += 1
            audit.error("browser_evidence_failure", f"{key}: assertions={missing} screenshot={screenshot!r}")
        if item.get("console_errors") not in ([], None) or item.get("failed_requests") not in ([], None):
            failures += 1
            audit.error("browser_runtime_error", key)
    return {
        "path": str(path),
        "expected_checks": len(expected_checks),
        "actual_checks": len(actual_checks),
        "failures": failures,
    }


def strip_cli_noise(value: str) -> str:
    value = ANSI_RE.sub("", value).replace("\b", "")
    return "\n".join(line.rstrip() for line in value.splitlines())


def parse_vercel_output(inspect: str, logs: str) -> dict[str, object]:
    inspect = strip_cli_noise(inspect)
    logs = strip_cli_noise(logs)
    try:
        inspect_json = parse_json_output(inspect)
    except GateError:
        inspect_json = None
    fields: dict[str, object] = {}
    aliases: list[str] = []
    if inspect_json:
        fields = {
            "id": inspect_json.get("id", ""),
            "name": inspect_json.get("name", ""),
            "target": inspect_json.get("target", ""),
            "status": inspect_json.get("readyState", ""),
            "url": inspect_json.get("url", ""),
            "created": inspect_json.get("createdAt", ""),
        }
        aliases = [
            value if str(value).startswith("http") else f"https://{value}"
            for value in inspect_json.get("aliases", [])
            if isinstance(value, str)
        ]
    else:
        for name in ("id", "name", "target", "status", "url", "created"):
            match = re.search(rf"(?mi)^\s*{name}\s+(.+?)\s*$", inspect)
            if match:
                fields[name] = clean(match.group(1).replace("●", ""))
        aliases = re.findall(r"https://[^\s]+", inspect)
    clone = re.search(
        r"Cloning\s+github\.com/([^\s]+)\s+\(Branch:\s*([^,]+),\s*Commit:\s*([0-9a-f]+)\)",
        logs, flags=re.I,
    )
    return {
        **fields,
        "aliases": sorted(dict.fromkeys(aliases)),
        "repository": clone.group(1) if clone else "",
        "branch": clean(clone.group(2)) if clone else "",
        "commit": clone.group(3).lower() if clone else "",
        "git_clone_found": bool(clone),
    }


def validate_vercel_deployment(
    audit: Audit,
    deployment: str,
    phase: str,
    timeout: int,
) -> dict[str, object]:
    if phase not in {"preview", "live"}:
        raise GateError("Vercel provenance is only valid for preview/live phases")
    inspect_result = subprocess.run(
        ["vercel", "inspect", deployment, "--json"], cwd=ROOT,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", timeout=timeout,
    )
    logs_result = subprocess.run(
        ["vercel", "inspect", deployment, "--logs"], cwd=ROOT,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", timeout=timeout,
    )
    if inspect_result.returncode or logs_result.returncode:
        audit.error("vercel_inspect", f"inspect={inspect_result.returncode} logs={logs_result.returncode}")
    parsed = parse_vercel_output(inspect_result.stdout, logs_result.stdout)
    expected_target = "preview" if phase == "preview" else "production"
    if clean(parsed.get("name")) != VERCEL_PROJECT:
        audit.error("vercel_project", parsed.get("name"))
    if clean(parsed.get("target")).lower() != expected_target:
        audit.error("vercel_target", f"{parsed.get('target')} expected {expected_target}")
    if "ready" not in clean(parsed.get("status")).lower():
        audit.error("vercel_status", parsed.get("status"))
    if parsed.get("repository") != GITHUB_REPOSITORY:
        audit.error("vercel_git_repository", parsed.get("repository"))
    head = git_text("rev-parse", "HEAD")
    commit = str(parsed.get("commit", ""))
    if not commit or not head.startswith(commit):
        audit.error("vercel_git_commit", f"deployment={commit} head={head}")
    branch = git_text("branch", "--show-current")
    if clean(parsed.get("branch")) != branch:
        audit.error("vercel_git_branch", f"deployment={parsed.get('branch')} local={branch}")
    aliases = set(parsed.get("aliases", []))
    live_alias = SITE_ORIGIN
    if phase == "live" and live_alias not in aliases:
        audit.error("vercel_live_alias", sorted(aliases))
    if phase == "preview" and live_alias in aliases:
        audit.error("vercel_preview_has_live_alias", live_alias)
    return parsed


def self_test(archive: Path) -> dict[str, object]:
    checks: list[str] = []
    if not 0 < EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS < EXPECTED_PRODUCT_DOCUMENT_COUNT:
        raise GateError("known-superseded transition cardinality self-test")
    checks.append("known_superseded_transition")
    parsed_output = parse_json_output('progress\n{"status":"PASS","outer":{"nested":1}}\n')
    if parsed_output.get("status") != "PASS" or nested_value(parsed_output, "outer", "nested") != 1:
        raise GateError("subprocess JSON parser self-test")
    environment = python_subprocess_environment()
    if environment.get("PYTHONIOENCODING") != "utf-8" or environment.get("PYTHONUTF8") != "1":
        raise GateError("subprocess UTF-8 environment self-test")
    checks.append("subprocess_json_utf8")
    if semantic_url(f"{SITE_ORIGIN}/%EA%B3%BC/") != f"{SITE_ORIGIN}/과/":
        raise GateError("semantic_url self-test")
    if transport_url(f"{SITE_ORIGIN}/x") == transport_url(f"{SITE_ORIGIN}/x/"):
        raise GateError("transport_url self-test")
    checks.append("semantic_url")
    sample = '<meta name="description" content="설명"><link href="/x/" rel="canonical"><h1>제목</h1>'
    if meta_content(sample, name="description") != "설명" or canonical_href(sample) != "/x/":
        raise GateError("HTML metadata self-test")
    checks.append("html_metadata")
    graph_sample = '<script type="application/ld+json">{"@context":"https://schema.org","@graph":[{"@type":["EducationalOrganization","LocalBusiness"],"@id":"x"}]}</script>'
    nodes, errors = parse_graph(graph_sample)
    if errors or node_types(nodes[0]) != {"EducationalOrganization", "LocalBusiness"}:
        raise GateError("JSON-LD self-test")
    checks.append("jsonld")
    payload = {
        "documents": [
            {"path": "b", "after_sha256": "11" * 32},
            {"path": "a", "after_sha256": "22" * 32},
        ],
        "source_manifest": {"base_commit": "aa" * 20, "z": "33" * 32, "a": "44" * 32},
    }
    if generator_candidate_sha256(payload) != generator_candidate_sha256(json.loads(json.dumps(payload))):
        raise GateError("candidate digest self-test")
    checks.append("candidate_digest")
    vercel_sample = parse_vercel_output(
        "name new14\ntarget preview\nstatus ● Ready\nurl https://new14-test.vercel.app\n",
        "Cloning github.com/01039578283-hub/new14 (Branch: codex/test, Commit: abcdef1)\n",
    )
    if vercel_sample["repository"] != GITHUB_REPOSITORY or vercel_sample["commit"] != "abcdef1":
        raise GateError("Vercel parser self-test")
    vercel_json_sample = parse_vercel_output(
        '{"id":"dpl_x","name":"new14","url":"preview.vercel.app","target":"preview",'
        '"readyState":"READY","createdAt":1,"aliases":["preview.vercel.app"]}',
        "Cloning github.com/01039578283-hub/new14 (Branch: codex/test, Commit: abcdef1)\n",
    )
    if vercel_json_sample["aliases"] != ["https://preview.vercel.app"]:
        raise GateError("Vercel JSON parser self-test")
    checks.append("vercel_parser")
    generator = generator_contract()
    checks.append("generator_api")
    toolchain = release_toolchain_contract()
    if (
        len(toolchain) != len(TOOL_PATHS)
        or toolchain[relative(GENERATOR_PATH)]["sha256"] != generator["actual_sha256"]
    ):
        raise GateError("release toolchain self-test")
    checks.append("release_toolchain")
    sources = archive_contract(archive)
    if any(item["rows"] != EXPECTED_PER_PROFILE for item in sources["workbooks"].values()):
        raise GateError("source row-count self-test")
    checks.append("source_archive")
    profiles = profile_detail_paths()
    if any(len(paths) != EXPECTED_PER_PROFILE for paths in profiles.values()):
        raise GateError("repository target count self-test")
    checks.append("repository_target_count")
    sitemap = sitemap_snapshot(ROOT / "sitemap.xml")
    if sitemap["count"] != EXPECTED_SITEMAP_COUNT:
        raise GateError("sitemap self-test")
    checks.append("sitemap")
    sitemap_lf = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        '  <url><loc>https://example.test/a/</loc><lastmod>2026-08-22</lastmod></url>\n'
        '</urlset>\n'
    ).encode("utf-8")
    sitemap_crlf = sitemap_lf.replace(b"\n", b"\r\n")
    lf_snapshot = sitemap_snapshot(ROOT / "sitemap.xml", sitemap_lf)
    crlf_snapshot = sitemap_snapshot(ROOT / "sitemap.xml", sitemap_crlf)
    if (
        lf_snapshot["outside_url_blocks_sha256"]
        == crlf_snapshot["outside_url_blocks_sha256"]
        or lf_snapshot["outside_url_blocks_lf_sha256"]
        != crlf_snapshot["outside_url_blocks_lf_sha256"]
        or lf_snapshot["order"] != crlf_snapshot["order"]
        or lf_snapshot["entries"] != crlf_snapshot["entries"]
    ):
        raise GateError("sitemap CRLF normalization self-test")
    checks.append("sitemap_crlf_normalization")
    return {
        "status": "PASS",
        "checks": checks,
        "generator": generator,
        "archive_sha256": sources["archive_sha256"],
        "target_counts": {key: len(value) for key, value in profiles.items()},
    }


def exclusive_json_write(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = canonical_json(payload)
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except Exception:
        path.unlink(missing_ok=True)
        raise


def source_report_summary(contract: object) -> object:
    """Keep console/report output compact while retaining full hashes in the baseline."""
    if not isinstance(contract, dict):
        return contract
    source_hashes = contract.get("source_hashes", {})
    return {
        "archive_sha256": contract.get("archive_sha256"),
        "workbooks": contract.get("workbooks", {}),
        "source_hash_counts": {
            profile: len(hashes) if isinstance(hashes, list) else None
            for profile, hashes in sorted(source_hashes.items())
        } if isinstance(source_hashes, dict) else {},
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Strict technical gate for the revised 1,113-page K/E/M release.")
    parser.add_argument(
        "--phase", choices=("baseline", "projected", "candidate", "local", "preview", "live"),
        default="baseline",
    )
    parser.add_argument("--archive", type=Path, default=DEFAULT_ARCHIVE)
    parser.add_argument("--common-dir", type=Path, default=DEFAULT_COMMON)
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--write-baseline", type=Path)
    parser.add_argument("--freeze-file", type=Path)
    parser.add_argument("--repeat-freeze-file", type=Path)
    parser.add_argument("--check-generator-idempotency", action="store_true")
    parser.add_argument("--check-content-auditor", action="store_true")
    parser.add_argument("--base-url")
    parser.add_argument("--http-workers", type=int, default=12)
    parser.add_argument("--http-timeout", type=int, default=30)
    parser.add_argument("--http-full-resources", action="store_true")
    parser.add_argument("--write-browser-plan", type=Path)
    parser.add_argument("--browser-evidence", type=Path)
    parser.add_argument("--vercel-deployment")
    parser.add_argument("--vercel-timeout", type=int, default=120)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    archive = args.archive.expanduser().resolve()
    common_dir = args.common_dir.expanduser().resolve()
    if args.self_test:
        result = self_test(archive)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    for output in (args.write_baseline, args.write_browser_plan, args.report):
        if output:
            assert_external_output(output)
    audit = Audit()
    source_contract = audit.capture("source_contract", archive_contract, archive)
    report: dict[str, object] = {
        "phase": args.phase,
        "root": str(ROOT),
        "technical_auditor_sha256": sha256_file(SCRIPT_PATH),
        "generator": audit.capture("generator_contract", generator_contract),
        "toolchain": audit.capture("release_toolchain_contract", release_toolchain_contract),
        "sources": source_report_summary(source_contract),
    }
    generator_report = report.get("generator")
    toolchain_report = report.get("toolchain")
    if isinstance(generator_report, dict) and isinstance(toolchain_report, dict):
        tool_generator = toolchain_report.get(relative(GENERATOR_PATH), {})
        if (
            not isinstance(tool_generator, dict)
            or tool_generator.get("sha256") != generator_report.get("actual_sha256")
        ):
            audit.error("generator_toolchain_race", "generator changed while its release contract was captured")
    pending_pins = [
        label for label, value in (
            ("generator", EXPECTED_GENERATOR_SHA256),
            ("candidate", EXPECTED_CANDIDATE_SHA256),
            ("after_manifest", EXPECTED_AFTER_MANIFEST_SHA256),
            ("content_auditor", EXPECTED_CONTENT_AUDITOR_SHA256),
        )
        if value == "PENDING"
    ]
    for label in pending_pins:
        if args.phase == "baseline":
            audit.warn(f"{label}_pin_pending", "baseline allowed; release phases remain fail-closed")
        else:
            audit.error(f"{label}_pin_pending", "release phase blocked until final SHA-256 pin")

    if args.phase == "baseline":
        baseline = audit.capture("baseline_build", build_baseline, archive)
        if baseline:
            if (
                isinstance(report.get("generator"), dict)
                and baseline.get("generator", {}).get("actual_sha256")
                != report["generator"].get("actual_sha256")
            ):
                audit.error("generator_baseline_race", "generator changed during baseline capture")
            report["baseline"] = {
                "baseline_commit": baseline["baseline_commit"],
                "detail_count": len(baseline["target_pages"]),
                "immutable_count": len(baseline["immutable_git_blobs"]),
                "sitemap_count": baseline["sitemap"]["count"],
                "target_path_set_sha256": baseline["release_contract"]["target_path_set_sha256"],
                "target_url_set_sha256": baseline["release_contract"]["target_url_set_sha256"],
            }
            if args.write_baseline:
                if audit.errors:
                    audit.error("baseline_write_blocked", "baseline capture has prior technical errors")
                else:
                    audit.capture(
                        "baseline_write", exclusive_json_write,
                        assert_external_output(args.write_baseline), baseline,
                    )
                    report["baseline_path"] = str(args.write_baseline.expanduser().resolve())
    else:
        if not args.baseline:
            audit.error("baseline_required", "--baseline is required outside baseline phase")
            baseline = None
        else:
            baseline = audit.capture("baseline_load", load_baseline, args.baseline.expanduser().resolve())
        if baseline:
            report["transaction_security"] = audit.capture(
                "transaction_security", validate_transaction_security, audit,
            )
            report["browser_contract"] = audit.capture(
                "browser_contract_static", validate_browser_contract_static, audit, baseline,
            )
            if args.phase == "projected":
                report["git"] = audit.capture(
                    "projected_git_scope", validate_projected_git_scope, audit, baseline,
                )
            else:
                report["git"] = audit.capture(
                    "git_scope", validate_git_scope, audit, baseline, args.phase,
                )
            report["immutability"] = audit.capture(
                "immutability", validate_immutable_tree, audit, baseline,
            )
            if args.phase == "projected":
                report["product"] = audit.capture(
                    "projected_product", validate_baseline_product_unchanged, audit, baseline,
                )
            else:
                report["sitemap"] = audit.capture(
                    "sitemap_change", validate_sitemap_change, audit, baseline,
                )
                if isinstance(source_contract, dict):
                    report["candidate"] = audit.capture(
                        "candidate_pages", validate_candidate_pages,
                        audit, baseline, source_contract,
                    )
            if not args.freeze_file:
                audit.error("freeze_required", "--freeze-file is required outside baseline phase")
            elif isinstance(source_contract, dict):
                report["freeze"] = audit.capture(
                    "freeze_validation", validate_freeze,
                    audit, args.freeze_file.expanduser().resolve(), baseline, source_contract,
                    args.repeat_freeze_file.expanduser().resolve() if args.repeat_freeze_file else None,
                    args.phase == "projected",
                )
            if args.check_generator_idempotency:
                report["generator_idempotency"] = audit.capture(
                    "generator_idempotency", validate_generator_idempotency,
                    audit, archive, common_dir, max(args.vercel_timeout, 600),
                    EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS if args.phase == "projected" else 0,
                )
            else:
                audit.error("generator_idempotency_evidence_missing", "pass --check-generator-idempotency")
            if args.check_content_auditor:
                report["content_auditor"] = audit.capture(
                    "content_auditor", validate_content_auditor,
                    audit, "projected" if args.phase == "projected" else "actual",
                    archive, common_dir, max(args.vercel_timeout, 1800),
                )
            else:
                audit.error("content_auditor_evidence_missing", "pass --check-content-auditor")
            if args.phase in {"local", "preview", "live"}:
                if not args.base_url:
                    audit.error("base_url_required", f"--base-url is required for {args.phase}")
                else:
                    report["http"] = audit.capture(
                        "http_deployment", validate_http_deployment,
                        audit, args.phase, args.base_url, baseline,
                        args.http_workers, args.http_timeout, args.http_full_resources,
                    )
                    if args.write_browser_plan:
                        plan = build_browser_plan(args.phase, args.base_url, baseline)
                        audit.capture(
                            "browser_plan_write", exclusive_json_write,
                            assert_external_output(args.write_browser_plan), plan,
                        )
                        report["browser_plan"] = str(args.write_browser_plan.expanduser().resolve())
                    if args.browser_evidence:
                        report["browser"] = audit.capture(
                            "browser_evidence", validate_browser_evidence,
                            audit, args.browser_evidence.expanduser().resolve(),
                            args.phase, args.base_url, baseline,
                        )
                    else:
                        audit.error("browser_evidence_missing", "provide --browser-evidence after Browser desktop/mobile checks")
            if args.phase in {"preview", "live"}:
                if not args.vercel_deployment:
                    audit.error("vercel_deployment_required", "provide --vercel-deployment")
                else:
                    report["vercel"] = audit.capture(
                        "vercel_deployment", validate_vercel_deployment,
                        audit, args.vercel_deployment, args.phase, args.vercel_timeout,
                    )

    report.update(audit.summary())
    report["status"] = "PASS" if not audit.errors else "FAIL"
    if args.report:
        try:
            exclusive_json_write(assert_external_output(args.report), report)
        except Exception as exc:  # report failure must affect the process status
            audit.error("report_write", f"{type(exc).__name__}: {exc}")
            report.update(audit.summary())
            report["status"] = "FAIL"
    # stdout is a machine-readable fallback in addition to the UTF-8 report
    # file.  ASCII escaping keeps it writable even under a legacy Windows
    # console code page such as cp949.
    print(json.dumps(report, ensure_ascii=True, indent=2))
    return 0 if not audit.errors else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (GateError, FileExistsError, PermissionError, OSError, subprocess.SubprocessError) as exc:
        print(json.dumps({"status": "FAIL", "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(1)
