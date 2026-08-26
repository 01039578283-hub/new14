from __future__ import annotations

"""Strict, read-only release audit for the three existing K/E/M categories.

The only write operations supported by this program are explicitly directed to
an output path outside the repository (baseline/report), or to an isolated
temporary copy when ``--check-idempotency`` is requested.  It never rewrites
the checked-out site.
"""

import argparse
import csv
import hashlib
import html
import json
import os
import re
import shutil
import statistics
import stat
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path
from urllib.parse import quote, unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
COMMON = ROOT.parent / "참고자료" / "공통자료"
CENTER_CSV = COMMON / "센터정보 정리.csv"
BASE_URL = "https://xn--3e0bz50b1zcyxat54c.com"
ROOT_ORG_ID = f"{BASE_URL}/#organization"
SUBJECT_ROOT = ROOT / "과목별학원"
PROFILES = {
    "elementary": {
        "slug": "초등학생국영수학원", "level": "초등학생", "course": "초등",
        "prefix": "초", "school": "초등학교", "school_column": "타깃학교\n(초)",
    },
    "middle": {
        "slug": "중학생국영수학원", "level": "중학생", "course": "중등",
        "prefix": "중", "school": "중학교", "school_column": "타깃학교\n(중)",
    },
    "high": {
        "slug": "고등학생국영수학원", "level": "고등학생", "course": "고등",
        "prefix": "고", "school": "고등학교", "school_column": "타깃학교\n(고)",
    },
}
SUBJECT_COLUMNS = {
    "국어": "가능학년\n(국어)", "영어": "가능학년\n(영어)", "수학": "가능학년\n(수학)",
}
REQUIRED_TYPES = {
    "EducationalOrganization", "LocalBusiness", "WebPage", "BreadcrumbList",
    "Article", "Service", "FAQPage", "ItemList",
}
KNOWN_COPY_ERRORS = (
    "고이 포함됩니다", "고등학교이 포함됩니다", "기록라는 표현",
    "에 대한 설명는", "중학교이 포함됩니다", "초이 포함됩니다",
    "초등학교이 포함됩니다", "관리을", "관리이라는",
    "이라는 이름으로 검색되더라도",
    "학부모님이 체감하는 현실은 학부모님이", "관리 기준 방문 전에는",
    "처럼 운영과 관련된 표현은", "영역으로 해석할 수 있습니다",
    "학습 점검의 내신 대비", "학습 계획의 내신 대비",
    "학습 점검 시험 후에는", "학습 계획 시험 후에는",
    "학습 과정 시험 후에는", "세 과목 학습 상담 시험 후에는",
    "관리 기준 시험 후에는", ", 그리고",
    "고려하면, 그리고", "살펴보면, 그리고",
    "상담 준비에서는 다음 기준을 적용합니다:",
    "비교할 때 화려한 설명보다 중요한 기준은 다음과 같습니다:",
    "확인하려는 이유는 단순한 수업보다", "계획보다는 점수보다",
    "과제 완료 기준을 기준으로",
    "요일를", "시간를", "범위을", "이해도을", "결과을", "진도을",
    "분류을", "학원로", "와와학교 일정 점검학원",
)
BAD_READER_SEEDS = (
    "녹화수업", "온라인수업", "방학캠프", "일대일수업", "야간수업",
    "입시성공사례", "학원자료실", "학습암기", "학원매출관리", "학원창업",
    "학원미납관리", "학원고객관리시스템", "학원전자계약", "학원관리솔루션",
    "학원결제시스템", "학원소수정예", "학원출입관리", "학원결제관리",
    "학원고객관리", "학원강사", "학원위치", "학원운영", "학원행정",
)
BAD_ELEMENTARY_SEEDS = (
    "학원매출관리", "학원창업", "학원미납관리", "학원고객관리시스템",
    "학원전자계약", "학원관리솔루션", "학원결제시스템", "까지 고려한",
)
READER_META_TERMS = (
    "키워드", "검색어", "검색엔진", "SEO", "상위노출", "템플릿",
    "표현으로 보는", "영역으로 해석", "보장 문구", "단어가 보여도",
    "말로 보는 것이 안전", "정보성 페이지 형태", "지역명을 바꾼 홍보 문구",
)
TRANSITION_PREFIXES = (
    "상담 기록을 기준으로 보면,", "최근 자료를 함께 놓고 보면,",
    "학생의 실제 실행을 살펴보면,", "과목별 점검 순서를 정할 때,",
    "다음 학습 계획을 세우려면,", "가정에서 관찰한 내용을 더하면,",
    "최근 교재와 오답을 대조하면,", "주간 학습 흐름을 확인하면,",
    "시험·교과 자료를 살펴보면,", "완료 기록을 중심으로 보면,",
    "상담 질문을 구체화하려면,", "학생의 현재 시간표를 고려하면,",
    "학교 자료를 대조할 때,", "학생 자료를 먼저 펼쳐 보면,",
    "실제 학습 범위를 확인하면,", "상담 전에 정리해 보면,",
    "센터 안내와 비교할 때,", "과목별 기록을 살펴보면,",
    "확인 순서를 세울 때,", "상담 자료를 정리하면,",
    "학교 일정까지 함께 보면,", "다음 계획을 정하기 전에,",
    "가정에서 기록을 준비하면,", "학생의 현재 자료를 기준으로 보면,",
)
AWKWARD_NOUN_BASE = r"(?:학습 계획|학습 점검|학습 과정|세 과목 학습 상담|관리 기준)"
AWKWARD_NOUN_PATTERNS = {
    "high": (
        rf"{AWKWARD_NOUN_BASE} 방문 전에는",
        rf"{AWKWARD_NOUN_BASE} 상담에서 (?:국어|영어|수학) 계획",
    ),
    "middle": (rf"{AWKWARD_NOUN_BASE} 수업 후 복습",),
    "elementary": (
        rf"{AWKWARD_NOUN_BASE} 수업은 진단",
        rf"{AWKWARD_NOUN_BASE} 수업 뒤에 무엇",
        rf"{AWKWARD_NOUN_BASE} 방문 상담 주소는",
        rf"{AWKWARD_NOUN_BASE} 교재 점검에서는",
    ),
}
REGION_NAMES = {
    "서울": "서울특별시", "서울특별시": "서울특별시",
    "경기": "경기도", "경기도": "경기도",
    "인천": "인천광역시", "인천광역시": "인천광역시",
    "부산": "부산광역시", "부산광역시": "부산광역시",
    "대구": "대구광역시", "대구광역시": "대구광역시",
    "대전": "대전광역시", "대전광역시": "대전광역시",
    "광주": "광주광역시", "광주광역시": "광주광역시",
    "울산": "울산광역시", "울산광역시": "울산광역시",
    "세종": "세종특별자치시", "세종시": "세종특별자치시",
    "세종특별자치시": "세종특별자치시",
    "강원": "강원특별자치도", "강원도": "강원특별자치도",
    "강원특별자치도": "강원특별자치도",
    "충북": "충청북도", "충청북도": "충청북도",
    "충남": "충청남도", "충청남도": "충청남도",
    "전북": "전북특별자치도", "전라북도": "전북특별자치도",
    "전북특별자치도": "전북특별자치도",
    "전남": "전라남도", "전라남도": "전라남도",
    "경북": "경상북도", "경상북도": "경상북도",
    "경남": "경상남도", "경상남도": "경상남도",
    "제주": "제주특별자치도", "제주도": "제주특별자치도",
    "제주특별자치도": "제주특별자치도",
}


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    return sha256_bytes(path.read_bytes())


def set_sha(values: list[str] | set[str]) -> str:
    return sha256_bytes(("\n".join(sorted(values)) + "\n").encode("utf-8"))


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", html.unescape(str(value or ""))).strip()


def strip_tags(value: str) -> str:
    value = re.sub(r"<(script|style)\b.*?</\1>", " ", value, flags=re.I | re.S)
    return clean(re.sub(r"<[^>]+>", " ", value))


def node_types(node: dict) -> set[str]:
    value = node.get("@type", [])
    return set(value if isinstance(value, list) else [value])


def first_node(nodes: list[dict], kind: str) -> dict:
    return next((node for node in nodes if kind in node_types(node)), {})


def as_list(value: object) -> list:
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def ref_id(value: object) -> str:
    return clean(value.get("@id", "")) if isinstance(value, dict) else ""


def split_values(value: str) -> list[str]:
    return list(dict.fromkeys(
        clean(part) for part in re.split(r"[,/·\n]+", value or "") if clean(part)
    ))


SCHOOL_SUFFIX_RE = re.compile(r"(?:초등학교|중학교|고등학교|초|중|고)$")


def split_schools(value: str) -> list[str]:
    """Parse CSV compound school cells without short-name substring matching."""
    value = clean(value)
    if not value or "지역내 모든 고등학교 가능" in value:
        return []
    parts = [clean(part) for part in re.split(r"[,/·.，\n]+", value) if clean(part)]
    result: list[str] = []
    for part in parts:
        tokens = part.split()
        if len(tokens) > 1 and all(SCHOOL_SUFFIX_RE.search(token) for token in tokens):
            result.extend(tokens)
        else:
            result.append(part)
    return list(dict.fromkeys(result))


def grades(value: str, prefix: str | None = None) -> list[str]:
    found = split_values(value)
    return [item for item in found if not prefix or item.startswith(prefix)]


def page_url(path: Path, root: Path = ROOT) -> str:
    parts = path.parent.relative_to(root).parts
    encoded = "/".join(quote(part, safe="") for part in parts)
    return f"{BASE_URL}/{encoded}/" if encoded else f"{BASE_URL}/"


def semantic_url(value: str) -> str:
    parsed = urlsplit(value)
    path = re.sub(r"/+", "/", unquote(parsed.path))
    if path != "/":
        path = path.rstrip("/") + "/"
    return f"{parsed.scheme.lower()}://{parsed.netloc.lower()}{path}"


def scoped_pages(root: Path = ROOT) -> list[Path]:
    pages: list[Path] = []
    for config in PROFILES.values():
        category = root / SUBJECT_ROOT.name / config["slug"]
        pages.append(category / "index.html")
        pages.extend(sorted(category.glob("*/index.html")))
    return sorted(pages)


def sitemap_urls(root: Path = ROOT) -> list[str]:
    return [
        clean(node.text) for node in ET.parse(root / "sitemap.xml").getroot().iter()
        if node.tag.endswith("loc") and clean(node.text)
    ]


def sitemap_entries(root: Path = ROOT) -> dict[str, str]:
    entries: dict[str, str] = {}
    for node in ET.parse(root / "sitemap.xml").getroot():
        location = next((clean(child.text) for child in node if child.tag.endswith("loc")), "")
        modified = next((clean(child.text) for child in node if child.tag.endswith("lastmod")), "")
        if location:
            entries[location] = modified
    return entries


def scoped_sitemap(urls: list[str]) -> list[str]:
    prefixes = [
        f"{BASE_URL}/{quote(SUBJECT_ROOT.name, safe='')}/{quote(c['slug'], safe='')}/"
        for c in PROFILES.values()
    ]
    return sorted(url for url in urls if any(url.startswith(prefix) for prefix in prefixes))


def extract_one(pattern: str, source: str) -> str:
    match = re.search(pattern, source, flags=re.I | re.S)
    return html.unescape(match.group(1)).strip() if match else ""


def page_record(path: Path, root: Path = ROOT) -> dict:
    source = path.read_text(encoding="utf-8")
    return {
        "path": path.relative_to(root).as_posix(),
        "url": page_url(path, root),
        "canonical": extract_one(r'<link\s+rel="canonical"\s+href="([^"]+)"', source),
        "og_url": extract_one(r'<meta\s+property="og:url"\s+content="([^"]+)"', source),
        "html_sha256": sha256_file(path),
    }


def git_value(*args: str, root: Path = ROOT) -> str:
    return subprocess.check_output(["git", *args], cwd=root, text=True, encoding="utf-8").strip()


def optional_git_value(*args: str, root: Path = ROOT) -> str:
    completed = subprocess.run(
        ["git", *args], cwd=root, text=True, encoding="utf-8",
        stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
    )
    return completed.stdout.strip() if completed.returncode == 0 else ""


def build_baseline(root: Path = ROOT) -> dict:
    pages = [page_record(path, root) for path in scoped_pages(root)]
    all_sitemap = sitemap_urls(root)
    scoped = scoped_sitemap(all_sitemap)
    lastmods = sitemap_entries(root)
    scoped_lastmods = {url: lastmods.get(url, "") for url in scoped}
    return {
        "schema_version": 1,
        "repo": str(root),
        "head": git_value("rev-parse", "HEAD", root=root),
        "base_url": BASE_URL,
        "categories": sorted(config["slug"] for config in PROFILES.values()),
        "counts": {
            "pages": len(pages), "details": sum(item["path"].count("/") == 3 for item in pages),
            "sitemap_total": len(all_sitemap), "sitemap_scoped": len(scoped),
        },
        "url_set_sha256": set_sha([item["url"] for item in pages]),
        "canonical_set_sha256": set_sha([item["canonical"] for item in pages]),
        "sitemap_scoped_set_sha256": set_sha(scoped),
        "sitemap_scoped_lastmod_sha256": set_sha([
            f"{url}\t{modified}" for url, modified in scoped_lastmods.items()
        ]),
        "html_manifest_sha256": set_sha([
            f'{item["path"]}\t{item["html_sha256"]}' for item in pages
        ]),
        "sitemap_file_sha256": sha256_file(root / "sitemap.xml"),
        "scoped_sitemap_urls": scoped,
        "scoped_sitemap_lastmods": scoped_lastmods,
        "pages": pages,
    }


class Audit:
    def __init__(self) -> None:
        self.counts: Counter[str] = Counter()
        self.samples: dict[str, list[str]] = defaultdict(list)

    def add(self, code: str, detail: str) -> None:
        self.counts[code] += 1
        if len(self.samples[code]) < 5:
            self.samples[code].append(detail)

    @property
    def total(self) -> int:
        return sum(self.counts.values())

    def payload(self) -> dict:
        return {
            "error_count": self.total,
            "error_codes": dict(sorted(self.counts.items())),
            "samples": dict(sorted(self.samples.items())),
        }


def load_source(audit: Audit) -> tuple[list[dict], dict[str, dict], dict[tuple[str, str, str], list[dict]]]:
    if not CENTER_CSV.exists():
        audit.add("source_missing", str(CENTER_CSV))
        return [], {}, {}
    with CENTER_CSV.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    if len(rows) != 371:
        audit.add("source_row_count", str(len(rows)))
    by_slug: dict[str, dict] = {}
    groups: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
    mandatory = (
        "근처 수업가능 동네", "센터명", "센터 주소", "교육지원청명칭", "교육지원청 등록번호",
    )
    for index, row in enumerate(rows, 2):
        for field in mandatory:
            if not clean(row.get(field)):
                audit.add("source_required_blank", f"row {index}: {field}")
        locality = clean(row.get("근처 수업가능 동네"))
        slug = re.sub(r"\s+", "", locality)
        if slug in by_slug:
            audit.add("source_slug_duplicate", slug)
        by_slug[slug] = row
        key = (
            clean(row.get("센터명")), clean(row.get("센터 주소")),
            clean(row.get("교육지원청 등록번호")),
        )
        groups[key].append(row)
        for subject, column in SUBJECT_COLUMNS.items():
            for grade in grades(row.get(column, "")):
                if not re.fullmatch(r"(?:초[1-6]|중[1-3]|고[1-3])", grade):
                    audit.add("source_invalid_grade", f"row {index}: {subject}={grade!r}")
        fee = clean(row.get("센터 교습비"))
        if fee and not fee.startswith("https://drive.google.com/"):
            audit.add("source_fee_url", f"row {index}: {fee}")
        for config in PROFILES.values():
            for school in split_schools(row.get(config["school_column"], "")):
                if not SCHOOL_SUFFIX_RE.search(school):
                    audit.add("source_invalid_school", f"row {index}: {school!r}")
    if len(by_slug) != 371:
        audit.add("source_locality_count", str(len(by_slug)))
    if len(groups) != 188:
        audit.add("source_physical_center_count", str(len(groups)))
    return rows, by_slug, groups


def profile_for(path: Path) -> tuple[str, dict] | tuple[None, None]:
    for name, config in PROFILES.items():
        if config["slug"] in path.parts:
            return name, config
    return None, None


def parse_graph(source: str) -> tuple[dict, list[dict]]:
    matches = re.findall(r'<script\s+type="application/ld\+json">(.*?)</script>', source, flags=re.I | re.S)
    if len(matches) != 1:
        raise ValueError(f"JSON-LD scripts={len(matches)}")
    graph = json.loads(matches[0])
    nodes = [node for node in graph.get("@graph", []) if isinstance(node, dict)]
    return graph, nodes


def official_address(row: dict) -> tuple[str, str]:
    tokens = clean(row.get("센터 주소")).replace(",", " ").split()
    region = REGION_NAMES.get(tokens[0], "") if tokens else ""
    locality = ""
    if region == "세종특별자치시":
        locality = clean(row.get("근처 수업가능 동네"))
    elif len(tokens) > 1:
        locality = tokens[1]
    return region, locality


def local_reference(page: Path, value: str) -> Path | None:
    parsed = urlsplit(html.unescape(value))
    if parsed.scheme or value.startswith(("#", "tel:", "mailto:", "javascript:")):
        return None
    value_path = unquote(parsed.path)
    if not value_path:
        return None
    candidate = ROOT / value_path.lstrip("/") if value_path.startswith("/") else page.parent / value_path
    candidate = candidate.resolve()
    try:
        candidate.relative_to(ROOT.resolve())
    except ValueError:
        return Path("__ESCAPES_REPOSITORY__")
    if candidate.is_dir():
        candidate /= "index.html"
    return candidate


def visible_faq(source: str) -> list[tuple[str, str]]:
    return [
        (strip_tags(question), strip_tags(answer))
        for question, answer in re.findall(
            r'<details(?:\s+[^>]*)?>\s*<summary>(.*?)</summary>\s*<p>(.*?)</p>\s*</details>',
            source, flags=re.I | re.S,
        )
    ]


def normalized_copy(value: str, row: dict, config: dict) -> str:
    replacements = {
        clean(row.get("근처 수업가능 동네")), clean(row.get("지역")),
        clean(row.get("시or구")), clean(row.get("센터명")), clean(row.get("센터 주소")),
        config["level"], config["course"], config["slug"],
    }
    for column in SUBJECT_COLUMNS.values():
        replacements.update(split_values(row.get(column, "")))
    replacements.update(split_schools(row.get(config["school_column"], "")))
    result = clean(value)
    for token in sorted((item for item in replacements if len(item) >= 2), key=len, reverse=True):
        result = result.replace(token, "{X}")
    result = re.sub(r"\d+(?:[.,]\d+)*", "{N}", result)
    result = re.sub(r"\{X\}(?:[·,/]\{X\})+", "{X}", result)
    return result


def expected_group_support(group: list[dict]) -> dict[str, list[str]]:
    support: dict[str, list[str]] = {}
    for subject, column in SUBJECT_COLUMNS.items():
        values = {grade for row in group for grade in grades(row.get(column, ""))}
        support[subject] = sorted(values, key=lambda item: ("초중고".find(item[:1]), item))
    return support


def validate_offer_claim(
    audit: Audit, rel: str, offer: dict, support: dict[str, list[str]], code: str,
    required_course: str | None = None,
) -> tuple[str, str] | None:
    if clean(offer.get("url")):
        return None
    name = clean(offer.get("name") or (offer.get("itemOffered") or {}).get("name"))
    match = re.fullmatch(r"(?:(초등|중등|고등)\s+)?(국어|영어|수학)\s+학습\s*상담", name)
    if not match:
        audit.add(code, f"{rel}: unrecognized non-fee offer {name!r}")
        return None
    course, subject = match.groups()
    allowed = support.get(subject, [])
    if required_course and course != required_course:
        audit.add(code, f"{rel}: {name!r}, expected course={required_course}")
    prefix = {"초등": "초", "중등": "중", "고등": "고"}.get(course or "")
    expected = [grade for grade in allowed if not prefix or grade.startswith(prefix)]
    actual = grades(clean(offer.get("eligibleCustomerType")))
    if not expected:
        audit.add(code, f"{rel}: unsupported {name!r}")
    elif actual != expected:
        audit.add(code, f"{rel}: {name!r} grades={actual}, expected={expected}")
    return course or "", subject


def distribution(values: list[int] | list[float]) -> dict:
    if not values:
        return {"min": 0, "median": 0, "p90": 0, "max": 0}
    ordered = sorted(values)
    return {
        "min": ordered[0], "median": statistics.median(ordered),
        "p90": ordered[min(len(ordered) - 1, int(len(ordered) * .9))], "max": ordered[-1],
    }


def compare_baseline(audit: Audit, baseline: dict, current: dict) -> dict:
    for key in (
        "url_set_sha256", "canonical_set_sha256", "sitemap_scoped_set_sha256",
    ):
        # Older baseline manifests remain readable; a missing optional lastmod
        # digest is not silently treated as a changed URL set.
        if key not in baseline:
            continue
        if baseline.get(key) != current.get(key):
            audit.add("baseline_url_set_changed", f"{key}: {baseline.get(key)} -> {current.get(key)}")
    old = {item["path"]: item for item in baseline.get("pages", [])}
    new = {item["path"]: item for item in current.get("pages", [])}
    if set(old) != set(new):
        audit.add("baseline_path_set_changed", f"removed={len(set(old)-set(new))}, added={len(set(new)-set(old))}")
    changed = [path for path in set(old) & set(new) if old[path].get("html_sha256") != new[path].get("html_sha256")]
    # HTML is expected to change during the authorized improvement.  The hash
    # manifest is a scope/accounting gate; URL/path/canonical sets remain hard
    # invariants while changed HTML paths are reported, not rejected.
    return {
        "html_changed": len(changed), "html_unchanged": len(set(old) & set(new)) - len(changed),
        "changed_samples": sorted(changed)[:10],
        "sitemap_lastmod_changed": (
            None if "sitemap_scoped_lastmod_sha256" not in baseline else
            baseline.get("sitemap_scoped_lastmod_sha256") != current.get("sitemap_scoped_lastmod_sha256")
        ),
    }


def audit_site(audit: Audit, source_by_slug: dict[str, dict], groups: dict) -> dict:
    pages = scoped_pages()
    details = [path for path in pages if path.parent.parent.name in {c["slug"] for c in PROFILES.values()}]
    if len(pages) != 1116:
        audit.add("scoped_page_count", str(len(pages)))
    if len(details) != 1113:
        audit.add("detail_page_count", str(len(details)))
    current_urls = {page_url(path) for path in pages}
    current_sitemap = set(scoped_sitemap(sitemap_urls()))
    sitemap_lastmods = sitemap_entries()
    if current_urls != current_sitemap:
        audit.add("sitemap_scope_mismatch", f"files-only={len(current_urls-current_sitemap)}, sitemap-only={len(current_sitemap-current_urls)}")
    expected_slugs = set(source_by_slug)
    for config in PROFILES.values():
        actual_slugs = {
            path.parent.name for path in (SUBJECT_ROOT / config["slug"]).glob("*/index.html")
        }
        if actual_slugs != expected_slugs:
            audit.add(
                "category_source_url_set",
                f'{config["slug"]}: source-only={len(expected_slugs-actual_slugs)}, html-only={len(actual_slugs-expected_slugs)}',
            )

    titles: list[str] = []
    descriptions: list[str] = []
    h1s: list[str] = []
    org_key_ids: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    org_id_keys: dict[str, set[tuple[str, str, str]]] = defaultdict(set)
    org_key_areas: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    paragraph_df: dict[str, set[str]] = defaultdict(set)
    normalized_paragraph_df: dict[str, set[str]] = defaultdict(set)
    sentence_df: dict[str, set[str]] = defaultdict(set)
    normalized_sentence_df: dict[str, set[str]] = defaultdict(set)
    normalized_faq_df: dict[str, set[str]] = defaultdict(set)
    exact_hero_df: dict[str, set[str]] = defaultdict(set)
    normalized_hero_df: dict[str, set[str]] = defaultdict(set)
    exact_note_df: dict[str, set[str]] = defaultdict(set)
    normalized_note_df: dict[str, set[str]] = defaultdict(set)
    keyword_counts: dict[str, list[int]] = defaultdict(list)
    locality_density: dict[str, list[float]] = defaultdict(list)
    hidden_image_bytes: set[tuple[str, int]] = set()
    broken_cache: dict[str, tuple[bool, int]] = {}
    known_schools = {
        school for candidate in source_by_slug.values()
        for field in ("타깃학교\n(초)", "타깃학교\n(중)", "타깃학교\n(고)")
        for school in split_schools(candidate.get(field, ""))
    }
    known_localities = sorted(
        {clean(candidate.get("근처 수업가능 동네")) for candidate in source_by_slug.values()},
        key=len, reverse=True,
    )

    for page in pages:
        rel = page.relative_to(ROOT).as_posix()
        source = page.read_text(encoding="utf-8")
        expected_url = page_url(page)
        canonical = extract_one(r'<link\s+rel="canonical"\s+href="([^"]+)"', source)
        og_url = extract_one(r'<meta\s+property="og:url"\s+content="([^"]+)"', source)
        title = strip_tags(extract_one(r"<title>(.*?)</title>", source)).removesuffix(" | 와와학습코칭센터 영어수학 전문학원")
        h1_matches = re.findall(r"<h1\b[^>]*>(.*?)</h1>", source, flags=re.I | re.S)
        h1 = strip_tags(h1_matches[0]) if len(h1_matches) == 1 else ""
        description = extract_one(r'<meta\s+name="description"\s+content="([^"]*)"', source)
        titles.append(title); h1s.append(h1); descriptions.append(description)
        if canonical != expected_url or og_url != expected_url:
            audit.add("canonical_og_url", f"{rel}: {canonical!r} / {og_url!r}")
        is_category_hub = page.parent.name in {config["slug"] for config in PROFILES.values()}
        if (not is_category_hub and title != h1) or not h1:
            audit.add("title_h1", f"{rel}: {title!r} / {h1!r}")
        robots = extract_one(r'<meta\s+name="robots"\s+content="([^"]*)"', source).lower()
        if "index" not in robots or "follow" not in robots or "noindex" in robots:
            audit.add("robots_meta", f"{rel}: {robots!r}")

        try:
            _, nodes = parse_graph(source)
        except (ValueError, json.JSONDecodeError) as exc:
            audit.add("jsonld", f"{rel}: {exc}")
            continue
        types = set().union(*(node_types(node) for node in nodes)) if nodes else set()
        webpage = first_node(nodes, "WebPage")
        if webpage and semantic_url(clean(webpage.get("url"))) != semantic_url(expected_url):
            audit.add("schema_page_url", f"{rel}: {webpage.get('url')!r}")

        # Hubs: their contract ends at complete 371-card Collection/ItemList coverage.
        profile_name, config = profile_for(page)
        if not config:
            audit.add("profile_resolution", rel)
            continue
        if page.parent.name == config["slug"]:
            required_hub = {"EducationalOrganization", "BreadcrumbList", "CollectionPage", "ItemList"}
            missing = required_hub - types
            if missing:
                audit.add("jsonld_required_types", f"{rel}: {sorted(missing)}")
            itemlist = first_node(nodes, "ItemList")
            items = as_list(itemlist.get("itemListElement"))
            if len(items) != 371 or clean(itemlist.get("numberOfItems")) not in ("", "371"):
                audit.add("hub_itemlist_count", f"{rel}: {len(items)}")
            continue

        missing = REQUIRED_TYPES - types
        if missing:
            audit.add("jsonld_required_types", f"{rel}: {sorted(missing)}")

        slug = page.parent.name
        row = source_by_slug.get(slug)
        if not row:
            audit.add("source_page_mapping", rel)
            continue
        locality = clean(row["근처 수업가능 동네"])
        expected_title = f'{locality} {config["level"]} 국영수학원'
        if title != expected_title:
            audit.add("title_source", f"{rel}: {title!r}, expected={expected_title!r}")

        hero_fact_html = extract_one(
            r'<div\s+class="local-answer-grid"[^>]*>(.*?)</div>',
            source,
        )
        hero_fact_text = strip_tags(hero_fact_html)
        if not hero_fact_text:
            audit.add("hero_fact_block_missing", rel)
        else:
            exact_hero_df[hero_fact_text].add(rel)
            normalized_hero_df[normalized_copy(hero_fact_text, row, config)].add(rel)

        org = first_node(nodes, "EducationalOrganization")
        article = first_node(nodes, "Article")
        service = first_node(nodes, "Service")
        faq = first_node(nodes, "FAQPage")
        if not all((org, article, service, faq)):
            audit.add("detail_schema_nodes", rel)
            continue
        published = clean(article.get("datePublished"))
        modified = clean(article.get("dateModified"))
        sitemap_modified = clean(sitemap_lastmods.get(expected_url))
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", published) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", modified):
            audit.add("article_dates", f"{rel}: published={published!r}, modified={modified!r}")
        elif published > modified:
            audit.add("article_dates", f"{rel}: published {published} > modified {modified}")
        if sitemap_modified != modified:
            audit.add("sitemap_lastmod_article", f"{rel}: sitemap={sitemap_modified!r}, article={modified!r}")
        org_id = clean(org.get("@id"))
        key = (clean(row["센터명"]), clean(row["센터 주소"]), clean(row["교육지원청 등록번호"]))
        org_key_ids[key].add(org_id); org_id_keys[org_id].add(key); org_key_areas[key].add(locality)
        address = org.get("address", {}) if isinstance(org.get("address"), dict) else {}
        identifier = org.get("identifier", {}) if isinstance(org.get("identifier"), dict) else {}
        if clean(org.get("name")) != key[0] or clean(address.get("streetAddress")) != key[1]:
            audit.add("org_center_fact", f"{rel}: name/address")
        if clean(identifier.get("value")) != key[2] or clean(identifier.get("name")) != clean(row["교육지원청명칭"]):
            audit.add("org_registration_fact", rel)
        region, address_locality = official_address(row)
        if clean(address.get("addressRegion")) != region or clean(address.get("addressLocality")) != address_locality:
            audit.add("postal_address", f"{rel}: {address.get('addressRegion')}/{address.get('addressLocality')} expected {region}/{address_locality}")
        if "telephone" in org:
            audit.add("branch_telephone", f"{rel}: {org.get('telephone')}")
        for unsupported_property in ("openingHours", "openingHoursSpecification", "contactPoint"):
            if unsupported_property in org:
                audit.add("branch_contact_claim", f"{rel}: {unsupported_property}")
        expected_areas = {clean(item["근처 수업가능 동네"]) for item in groups.get(key, [row])}
        actual_areas = {clean(item) for item in as_list(org.get("areaServed"))}
        if actual_areas != expected_areas:
            audit.add("org_area_served", f"{rel}: {sorted(actual_areas)} expected {sorted(expected_areas)}")
        if ref_id(article.get("author")) != ROOT_ORG_ID or ref_id(article.get("publisher")) != ROOT_ORG_ID:
            audit.add("article_author_publisher", f"{rel}: {ref_id(article.get('author'))} / {ref_id(article.get('publisher'))}")
        if ref_id(service.get("provider")) != org_id:
            audit.add("service_provider", rel)
        if set(clean(item) for item in as_list(service.get("areaServed"))) != {locality}:
            audit.add("service_area", f"{rel}: {service.get('areaServed')!r}")

        group_support = expected_group_support(groups.get(key, [row]))
        expected_teaches = {"학습코칭"}
        for course, prefix in (("초등", "초"), ("중등", "중"), ("고등", "고")):
            for subject, values in group_support.items():
                if any(value.startswith(prefix) for value in values):
                    expected_teaches.add(f"{course} {subject}")
        for subject in ("영어", "수학"):
            if {value[:1] for value in group_support[subject]} == {"초", "중", "고"}:
                expected_teaches.add(f"초·중·고 {subject}")
        actual_teaches = {clean(item) for item in as_list(org.get("teaches"))}
        if actual_teaches != expected_teaches:
            audit.add("org_teaches_exact", f"{rel}: actual={sorted(actual_teaches)}, expected={sorted(expected_teaches)}")
        expected_levels = sorted(
            {value for values in group_support.values() for value in values},
            key=lambda item: ("초중고".find(item[:1]), item),
        )
        actual_levels = grades("·".join(clean(item) for item in as_list(org.get("educationalLevel"))))
        if actual_levels != expected_levels:
            audit.add("org_educational_level", f"{rel}: {actual_levels} expected {expected_levels}")
        for teach in actual_teaches:
            match = re.fullmatch(r"(초등|중등|고등) (국어|영어|수학)", teach)
            if match:
                course, subject = match.groups()
                prefix = {"초등": "초", "중등": "중", "고등": "고"}[course]
                if not any(grade.startswith(prefix) for grade in group_support[subject]):
                    audit.add("org_teaches_unsupported", f"{rel}: {teach}")
            elif teach.startswith("초·중·고 "):
                subject = teach.removeprefix("초·중·고 ")
                prefixes = {grade[:1] for grade in group_support.get(subject, [])}
                if subject not in group_support or prefixes != {"초", "중", "고"}:
                    audit.add("org_teaches_unsupported", f"{rel}: {teach}")
            elif teach not in ("학습코칭",):
                audit.add("org_teaches_unknown", f"{rel}: {teach}")
        seen_org_offers: set[tuple[str, str]] = set()
        org_fee_urls: set[str] = set()
        for offer in as_list(org.get("makesOffer")):
            if isinstance(offer, dict):
                if clean(offer.get("url")):
                    org_fee_urls.add(clean(offer.get("url")))
                else:
                    parsed = validate_offer_claim(audit, rel, offer, group_support, "org_offer_unsupported")
                    if parsed:
                        seen_org_offers.add(parsed)
        expected_org_offers = {
            (course, subject)
            for course, prefix in (("초등", "초"), ("중등", "중"), ("고등", "고"))
            for subject, values in group_support.items() if any(value.startswith(prefix) for value in values)
        }
        expected_org_offers.update(("", subject) for subject in ("영어", "수학") if group_support[subject])
        if seen_org_offers != expected_org_offers:
            audit.add("org_offer_coverage", f"{rel}: seen={sorted(seen_org_offers)}, expected={sorted(expected_org_offers)}")
        expected_group_fees = {clean(item.get("센터 교습비")) for item in groups.get(key, [row]) if clean(item.get("센터 교습비"))}
        if org_fee_urls != expected_group_fees:
            audit.add("org_fee_offer", f"{rel}: {sorted(org_fee_urls)} expected {sorted(expected_group_fees)}")

        page_support = {
            subject: grades(row.get(column, ""), config["prefix"])
            for subject, column in SUBJECT_COLUMNS.items()
        }
        seen_service: set[tuple[str, str]] = set()
        fee_offers: list[dict] = []
        for offer in as_list(service.get("offers")):
            if not isinstance(offer, dict):
                audit.add("service_offer_invalid", rel)
                continue
            if clean(offer.get("url")):
                fee_offers.append(offer)
                continue
            parsed = validate_offer_claim(
                audit, rel, offer, page_support, "service_offer_unsupported", config["course"],
            )
            if parsed:
                seen_service.add(parsed)
            item_id = ref_id(offer.get("itemOffered"))
            if item_id and not semantic_url(item_id.split("#", 1)[0]) == semantic_url(expected_url):
                audit.add("service_offer_foreign_id", f"{rel}: {item_id}")
        expected_service = {(config["course"], subject) for subject, values in page_support.items() if values}
        if seen_service != expected_service:
            audit.add("service_offer_coverage", f"{rel}: seen={sorted(seen_service)}, expected={sorted(expected_service)}")

        fee_url = clean(row.get("센터 교습비"))
        actual_fee_urls = {clean(offer.get("url")) for offer in fee_offers}
        if fee_url:
            if actual_fee_urls != {fee_url} or f'href="{html.escape(fee_url, quote=True)}"' not in source:
                audit.add("fee_fact", f"{rel}: {sorted(actual_fee_urls)} expected {fee_url}")
        elif actual_fee_urls or "교습비 확인" in strip_tags(source):
            audit.add("fee_blank_claim", rel)

        card = extract_one(r'<aside\s+class="local-info-card"[^>]*>(.*?)</aside>', source)
        card_text = strip_tags(card)
        for field in ("센터명", "센터 주소", "교육지원청 등록번호"):
            if clean(row[field]) not in card_text:
                audit.add("visible_center_fact", f"{rel}: {field}")
        school_block = extract_one(
            rf'<dt>{re.escape(config["school"])} 참고</dt>\s*<dd>(.*?)</dd>', card,
        )
        visible_schools = [strip_tags(value) for value in re.findall(r"<span>(.*?)</span>", school_block, flags=re.S)]
        expected_schools = split_schools(row.get(config["school_column"], ""))
        if visible_schools != expected_schools:
            audit.add("visible_school_fact", f"{rel}: {visible_schools} expected {expected_schools}")
        mention_names = {
            clean(item.get("name")) for item in as_list(article.get("mentions"))
            if isinstance(item, dict) and clean(item.get("name")) in known_schools
        }
        if mention_names != set(expected_schools):
            audit.add("schema_school_fact", f"{rel}: {sorted(mention_names)} expected {expected_schools}")
        grade_block = extract_one(r'<ul\s+class="grade-list">(.*?)</ul>', source)
        visible_grades = {
            strip_tags(subject): grades(strip_tags(value))
            for subject, value in re.findall(
                r'<li><strong>(.*?)</strong><span>(.*?)</span></li>', grade_block, flags=re.S
            )
        }
        if visible_grades != page_support:
            normalized_visible = {
                subject: ([] if values == [f'{config["course"]} 과정 미기재'] else values)
                for subject, values in visible_grades.items()
            }
            if normalized_visible != page_support:
                audit.add("visible_grade_fact", f"{rel}: {visible_grades} expected {page_support}")

        article_levels = sorted(
            {value for values in page_support.values() for value in values},
            key=lambda item: ("초중고".find(item[:1]), item),
        )
        actual_article_levels = grades("·".join(clean(item) for item in as_list(article.get("educationalLevel"))))
        if actual_article_levels != article_levels:
            audit.add("article_educational_level", f"{rel}: {actual_article_levels} expected {article_levels}")
        audience = service.get("audience", {}) if isinstance(service.get("audience"), dict) else {}
        if clean(audience.get("audienceType")) != config["level"]:
            audit.add("service_audience", f"{rel}: {audience.get('audienceType')!r}")

        visible = strip_tags(extract_one(r"<main\b[^>]*>(.*?)</main>", source))
        for phrase in KNOWN_COPY_ERRORS:
            hits = visible.count(phrase)
            for _ in range(hits):
                audit.add("known_copy_error", f"{rel}: {phrase}")
        raw_region = clean(row.get("지역"))
        raw_district = clean(row.get("시or구"))
        if raw_region in {"충청", "경상", "전라"} and f"{raw_region} {raw_district}" in visible:
            audit.add(
                "visible_broad_region_join",
                f"{rel}: {raw_region} {raw_district}",
            )
        district_stem = re.sub(r"(?:시|군|구)$", "", raw_district)
        display_locality = locality
        for prefix in dict.fromkeys(
            value for value in (district_stem, raw_region) if value
        ):
            match = re.fullmatch(rf"{re.escape(prefix)}\s+(.+)", locality)
            if match:
                display_locality = clean(match.group(1))
                break
        locality_forms = {
            clean(value)
            for value in (locality, locality.split()[-1], display_locality)
            if clean(value)
        }
        locality_form_pattern = "(?:" + "|".join(
            re.escape(value)
            for value in sorted(locality_forms, key=len, reverse=True)
        ) + ")"
        duplicate_city_patterns: list[str] = []
        if district_stem and re.match(
            rf"{re.escape(district_stem)}(?:\s|$)",
            locality,
        ):
            duplicate_city_patterns.append(
                rf"{re.escape(raw_district)}\s+{re.escape(district_stem)}"
                rf"(?=\s|[,.!?)]|$)"
            )
        if raw_region and raw_region == district_stem:
            duplicate_city_patterns.append(
                rf"{re.escape(raw_region)}\s+{re.escape(raw_district)}"
                rf"(?=\s|[,.!?)]|$)"
            )
        elif raw_region and re.match(
            rf"{re.escape(raw_region)}(?:\s|$)",
            locality,
        ):
            duplicate_city_patterns.append(
                rf"{re.escape(raw_region)}\s+{re.escape(raw_district)}\s+"
                rf"{re.escape(raw_region)}(?=\s|[,.!?)]|$)"
            )
        for pattern in duplicate_city_patterns:
            if match := re.search(pattern, visible):
                audit.add("duplicate_city_join", f"{rel}: {match.group(0)}")
        scenario_copy = " ".join(
            strip_tags(value)
            for value in re.findall(
                r'<article\s+class="scenario-card"[^>]*>.*?<p>(.*?)</p>\s*</article>',
                source,
                flags=re.I | re.S,
            )
        )
        for match in re.finditer(
            r"\b[^.!?\n]{1,40}\s+기준\s+제공된",
            scenario_copy,
        ):
            audit.add("scenario_machine_copy", f"{rel}: {match.group(0)}")
        if profile_name == "elementary":
            for phrase in BAD_ELEMENTARY_SEEDS:
                if phrase in visible:
                    audit.add("irrelevant_seo_seed", f"{rel}: {phrase}")
            intent_copy = visible
            for known_locality in known_localities:
                intent_copy = intent_copy.replace(known_locality, "{지역}")
            if "입시" in intent_copy or "내신" in intent_copy:
                audit.add(
                    "elementary_intent_drift",
                    f"{rel}: 입시={intent_copy.count('입시')}, 내신={intent_copy.count('내신')}",
                )
        manuscript_html = extract_one(
            r'<section\s+class="(?:section\s+)?manuscript-wrap"[^>]*>'
            r'(.*?)(?=</article>\s*</section>\s*<section\s+class="section\s+blue-wash")',
            source,
        )
        if not manuscript_html:
            audit.add("manuscript_extract", rel)
        manuscript_visible = strip_tags(manuscript_html)
        manuscript_paragraph_texts = [
            strip_tags(value)
            for value in re.findall(
                r"<p\b[^>]*>(.*?)</p>",
                manuscript_html,
                flags=re.I | re.S,
            )
        ]
        manuscript_paragraph_copy = " ".join(manuscript_paragraph_texts)
        manuscript_sentences = [
            clean(sentence)
            for paragraph in manuscript_paragraph_texts
            for sentence in re.split(r"(?<=[.!?])\s+", paragraph)
            if clean(sentence)
        ]
        for heading in (
            strip_tags(value)
            for value in re.findall(
                r"<h2\b[^>]*>(.*?)</h2>",
                manuscript_html,
                flags=re.I | re.S,
            )
        ):
            if heading.count("국어·영어·수학") >= 2:
                audit.add("h2_kem_double_topic", f"{rel}: {heading}")
        for pattern in AWKWARD_NOUN_PATTERNS[profile_name]:
            for match in re.finditer(pattern, manuscript_paragraph_copy):
                audit.add("awkward_noun_join", f"{rel}: {match.group(0)}")
        transition_counts: Counter[str] = Counter()
        for sentence in manuscript_sentences:
            for prefix in TRANSITION_PREFIXES:
                if sentence.startswith(prefix):
                    transition_counts[prefix] += 1
                    break
        for prefix, count in transition_counts.items():
            if count > 1:
                audit.add(
                    "mechanical_transition_repeat",
                    f"{rel}: {count}x {prefix}",
                )
        if sum(transition_counts.values()) > 6:
            audit.add(
                "mechanical_transition_overuse",
                f"{rel}: {sum(transition_counts.values())} transitions",
            )
        authored_faq = visible_faq(source)
        authored_visible = " ".join(
            [manuscript_visible]
            + [question + " " + answer for question, answer in authored_faq]
        )
        summary_html = extract_one(
            r'<div\s+class="local-summary"[^>]*>'
            r'(.*?)(?=</div>\s*<aside\s+class="local-info-card")',
            source,
        )
        if not summary_html or "manuscript-wrap" in summary_html:
            audit.add("summary_extract", rel)
        repeat_blocks = list(manuscript_paragraph_texts)
        repeat_blocks.extend(
            strip_tags(value)
            for value in re.findall(
                r"<p\b[^>]*>(.*?)</p>",
                summary_html,
                flags=re.I | re.S,
            )
        )
        repeat_blocks.extend(
            strip_tags(value)
            for value in re.findall(
                r'<article\s+class="scenario-card"[^>]*>.*?<p>(.*?)</p>\s*</article>',
                source,
                flags=re.I | re.S,
            )
        )
        repeat_blocks.extend(
            value
            for question, answer in authored_faq
            for value in (question, answer)
        )
        for phrase in BAD_READER_SEEDS:
            if phrase in authored_visible:
                audit.add("unverified_reader_seed", f"{rel}: {phrase}")
        for phrase in READER_META_TERMS:
            if phrase in authored_visible:
                audit.add("reader_facing_meta_copy", f"{rel}: {phrase}")
        if re.search(r"(?<![가-힣])원고(?=(?:\s|는|를|가|에서|의|$))", authored_visible):
            audit.add("reader_facing_meta_copy", f"{rel}: 원고")
        for phrase in ("해당 지역", "이 지역"):
            if phrase in manuscript_visible:
                audit.add("generic_locality_copy", f"{rel}: {phrase}")
        authoritative_address = clean(row.get("센터 주소"))
        address_cue = re.compile(
            r"주소는|주소\s*표기|제공\s+주소|주소\s+기준|"
            r"센터\s+주소|센터\s+위치는"
        )
        for address_sentence in manuscript_sentences:
            if not address_cue.search(address_sentence):
                continue
            if authoritative_address not in address_sentence:
                audit.add(
                    "manuscript_address_mutation",
                    f"{rel}: {clean(address_sentence)!r} expected {authoritative_address!r}",
                )
        authored_patterns = {
            "elementary_situation_copula": (
                r"상황이라면\s+[^.!?\n]{5,220}?학생입니다"
            ),
            "elementary_grade_choice": r"초등\s*[1-6]학년\s+중",
            "repeated_needed_student": (
                r"(?P<a>[^.!?\n]{2,80}?)(?:이|가)\s+필요한\s+"
                r"(?P<b>[^.!?\n]{2,80}?)(?:이|가)\s+필요한\s+학생"
            ),
            "repeated_address_basis": r"기준\s+주소\s+기준",
            "sejong_geography_join": (
                r"(?:충청|세종특별자치시)\s+새롬중앙로\s+(?:다정동|새롬동)"
            ),
            "operation_meta_copy": (
                r"처럼\s+운영(?:\s+관리)?와\s+관련된\s+표현은"
                r"|투명성으로\s+해석"
            ),
            "promotion_meta_copy": r"지역명을\s+바꾼\s+홍보\s+문구",
            "representative_case_copula": (
                r"대표\s+상담\s+사례는\s+[^.!?\n]{1,140}?학생입니다"
            ),
            "nested_case_student_possessive": (
                r"이\s+상담\s+사례의\s+학생의"
            ),
            "locality_case_student_dative": (
                rf"{locality_form_pattern}의\s+이\s+상담\s+사례의\s+학생에게는"
            ),
            "student_topic_particle": (
                r"이\s+학생에게는\s+(?:내신\s+대비|국어|영어|수학|"
                r"세\s+과목)(?:은|는)"
            ),
            "student_time_topic_particle": (
                r"이\s+학생에게는\s+(?:(?:시험\s+3주\s+전|시험\s+직전|"
                r"중간고사\s+직후)에는|오답은)"
            ),
            "high_double_topic_join": (
                r"(?:학생에게는|에서는)\s+(?:(?:시험\s+3주\s+전|시험\s+직전|"
                r"중간고사\s+직후|기말\s+기간)에는|(?:내신\s+대비|오답)(?:은|는))"
            ),
            "middle_review_join": (
                rf"{AWKWARD_NOUN_BASE}(?:을|를)\s+진행한\s+뒤\s+복습에\s+"
                r"더\s+잘\s+맞습니다"
            ),
            "middle_child_case_join": (
                r"우리\s+아이가\s+[^.!?\n]{1,120}?학생일\s+때"
            ),
            "elementary_spaced_student": r"초등\s+학생",
            "elementary_grade_choice_join": r"초등\s+중",
            "elementary_plural_join": r"초등에는",
            "elementary_trait_join": r"초등\s+특성",
            "elementary_state_join": r"초등(?:\s+과정)?\s+상태",
            "elementary_course_noun_join": (
                r"초등\s+과정\s+(?:학생|중\b|시기|아이|특성)"
            ),
            "elementary_problem_join": (
                r"초등학생이\s+[^.!?\n]{0,100}?(?:문제|고민|과제)(?:이|가)\s+"
                r"(?:나타나는|있는|남은)"
            ),
            "elementary_similar_repeat": (
                r"비슷한\s+[^.!?\n]{0,100}?비슷한\s+"
                r"(?:과정|학습|보완|복습|어려움|공부|학년대)"
            ),
            "elementary_same_repeat": (
                r"같은\s+[^.!?\n]{0,100}?같은\s+"
                r"(?:성장|공부|학습|실수|지점|실행|주)"
            ),
            "elementary_locality_grade_frame": (
                rf"{locality_form_pattern}의\s+(?:현재\s+학년\s+흐름의|"
                r"현재\s+과정의|이\s+학습\s+단계의|"
                r"비슷한\s+과정(?:에\s+있는|의)?|해당\s+학년대의|"
                r"같은\s+성장\s+단계의|비슷한\s+학년대의)"
            ),
            "elementary_representative_double_student": (
                r"초등학생\s+가운데[^.!?\n]{1,120}?학생을\s+대표\s+사례로"
            ),
            "elementary_grade_token_join": (
                r"초[1-6]\s+중\b|초[1-6]에서\s+초[1-6](?:으)?로\s+"
                r"(?:이어지는|올라가는)\s+시기\s+중"
            ),
            "elementary_repeated_trait_frame": (
                r"초등학생에게는[^.!?\n]{0,360}?초등학생의\s+특성에\s+맞춰"
            ),
            "elementary_bare_dative": r"초등에",
            "elementary_meta_demonstrative": (
                r"(?<![가-힣])이와\s+(?:학습|과정|복습|보완)"
            ),
            "elementary_student_child_join": (
                r"초등[^.!?\n]{0,100}?학생이[^.!?\n]{0,180}?아이일수록"
            ),
        }
        for code, pattern in authored_patterns.items():
            if code.startswith("elementary_") and profile_name != "elementary":
                continue
            if code.startswith("middle_") and profile_name != "middle":
                continue
            if code.startswith("high_") and profile_name != "high":
                continue
            for match in re.finditer(pattern, manuscript_visible):
                audit.add(code, f"{rel}: {clean(match.group(0))}")
        authored_meta_patterns = {
            "search_meta_copy": (
                r"이\s+글은[^.!?]*검색한"
                r"|(?:을|를)\s+검색(?:한|하는|했다면)"
            ),
            "page_self_reference": (
                r"(?<![가-힣])(?:이\s+(?:글|페이지)"
                r"|페이지(?=(?:입니다|에서는|에는|에서|의|에|를|가|는|로|라는|"
                r"\s|[,.!?]|$)))"
            ),
            "advertising_meta_copy": r"광고(?:처럼|\s*문구|의\s+크기|보다)",
        }
        for code, pattern in authored_meta_patterns.items():
            for match in re.finditer(pattern, authored_visible):
                audit.add(code, f"{rel}: {clean(match.group(0))}")
        repeated_windows: Counter[tuple[str, ...]] = Counter()
        for repeat_block in repeat_blocks:
            copy_tokens = re.findall(r"[가-힣A-Za-z0-9·]+", repeat_block)
            repeated_windows.update(
                tuple(copy_tokens[index:index + 8])
                for index in range(max(0, len(copy_tokens) - 7))
            )
        repeated_window, repeated_count = max(
            repeated_windows.items(), key=lambda item: item[1], default=((), 0)
        )
        if repeated_count >= 3:
            audit.add(
                "within_page_8word_repeat",
                f"{rel}: {repeated_count}x {' '.join(repeated_window)}",
            )
        exact_phrase_count = manuscript_visible.count(expected_title)
        keyword_counts[profile_name].append(exact_phrase_count)
        tokens = max(1, len(manuscript_visible.split()))
        locality_density[profile_name].append(round(100 * manuscript_visible.count(locality) / tokens, 3))
        if exact_phrase_count > 3:
            audit.add("keyword_exact_h1_overuse", f"{rel}: {exact_phrase_count}")
        if 100 * manuscript_visible.count(locality) / tokens > 8:
            audit.add("keyword_locality_density", f"{rel}: {100 * manuscript_visible.count(locality) / tokens:.2f}/100 tokens")

        page_paragraphs: list[str] = []
        for paragraph in re.findall(r"<p\b[^>]*>(.*?)</p>", manuscript_html, flags=re.I | re.S):
            value = strip_tags(paragraph)
            if len(value) >= 80:
                page_paragraphs.append(value)
                paragraph_df[value].add(rel)
                normalized_paragraph_df[normalized_copy(value, row, config)].add(rel)
        if len(page_paragraphs) != len(set(page_paragraphs)):
            audit.add("within_page_paragraph_duplicate", rel)
        page_sentences = [clean(part) for part in re.split(r"(?<=[.!?])\s+", manuscript_visible) if len(clean(part)) >= 30]
        if len(page_sentences) != len(set(page_sentences)):
            audit.add("within_page_sentence_duplicate", rel)
        for sentence in set(page_sentences):
            sentence_df[sentence].add(rel)
            normalized_sentence_df[normalized_copy(sentence, row, config)].add(rel)

        actual_faq = authored_faq
        raw_schema_faq = [
            (str(item.get("name", "")), str((item.get("acceptedAnswer") or {}).get("text", "")))
            for item in as_list(faq.get("mainEntity")) if isinstance(item, dict)
        ]
        if any(question != clean(question) or answer != clean(answer) for question, answer in raw_schema_faq):
            audit.add("faq_schema_whitespace", rel)
        schema_faq = [
            (clean(question), clean(answer)) for question, answer in raw_schema_faq
        ]
        if len(actual_faq) != 4 or actual_faq != schema_faq:
            audit.add("faq_parity", f"{rel}: visible={len(actual_faq)}, schema={len(schema_faq)}")
        for question, answer in actual_faq:
            for match in re.finditer(
                r"상담에서는\s+제공된\s+센터\s+주소는",
                answer,
            ):
                audit.add("faq_address_machine_copy", f"{rel}: {match.group(0)}")
            for match in re.finditer(
                r"(?:상담\s+전에는|문의\s+기준으로)\s+방문\s+위치는",
                answer,
            ):
                audit.add("faq_double_topic_location", f"{rel}: {match.group(0)}")
            normalized_faq_df[normalized_copy(question + "\t" + answer, row, config)].add(rel)
        note = extract_one(r'<p\s+class="center-verified-note"[^>]*>(.*?)</p>', source)
        if note:
            note_text = strip_tags(note)
            exact_note_df[note_text].add(rel)
            normalized_note_df[normalized_copy(note_text, row, config)].add(rel)
        else:
            audit.add("source_note_missing", rel)

        # Local links and images.  Cache repeated assets to avoid thousands of stats.
        for tag, attrs in re.findall(r"<(a|img)\b([^>]*)>", source, flags=re.I | re.S):
            attribute = "href" if tag.lower() == "a" else "src"
            value = extract_one(rf'{attribute}="([^"]+)"', attrs)
            if not value:
                continue
            target = local_reference(page, value)
            if target is None:
                continue
            key_target = str(target)
            if key_target not in broken_cache:
                broken_cache[key_target] = (target.exists() and target.is_file(), target.stat().st_size if target.exists() and target.is_file() else 0)
            exists, size = broken_cache[key_target]
            if not exists:
                audit.add("broken_local_reference", f"{rel}: {value}")
            if tag.lower() == "img":
                if not extract_one(r'alt="([^"]*)"', attrs):
                    audit.add("image_alt", f"{rel}: {value}")
                if not re.search(r"\bwidth=", attrs) or not re.search(r"\bheight=", attrs):
                    audit.add("image_dimensions", f"{rel}: {value}")
                if re.search(r"display\s*:\s*none", attrs, flags=re.I):
                    audit.add("hidden_generated_image", f"{rel}: {value}")
                    if exists:
                        hidden_image_bytes.add((key_target, size))

    if len(set(titles)) != len(titles): audit.add("duplicate_title", str(len(titles) - len(set(titles))))
    if len(set(h1s)) != len(h1s): audit.add("duplicate_h1", str(len(h1s) - len(set(h1s))))
    if len(set(descriptions)) != len(descriptions): audit.add("duplicate_description", str(len(descriptions) - len(set(descriptions))))
    for key, ids in org_key_ids.items():
        if len(ids) != 1: audit.add("physical_org_id_unstable", f"{key[0]}: {sorted(ids)}")
    for org_id, keys in org_id_keys.items():
        if len(keys) != 1: audit.add("physical_org_id_collision", f"{org_id}: {len(keys)} keys")
    if len(org_key_ids) != 188 or len(org_id_keys) != 188:
        audit.add("physical_org_identity_count", f"keys={len(org_key_ids)}, ids={len(org_id_keys)}")
    for key, localities in org_key_areas.items():
        # Check every rendered copy of the shared organization has the complete area set.
        expected = {clean(row["근처 수업가능 동네"]) for row in groups.get(key, [])}
        if localities != expected:
            audit.add("physical_org_area_coverage", f"{key[0]}: {sorted(localities)} expected {sorted(expected)}")

    def duplicate_metric(index: dict[str, set[str]]) -> dict:
        duplicate = {key: value for key, value in index.items() if len(value) > 1}
        return {
            "unique": len(index), "duplicate_patterns": len(duplicate),
            "affected_pages": len(set().union(*duplicate.values())) if duplicate else 0,
            "max_df": max((len(value) for value in index.values()), default=0),
        }

    duplicate_metrics = {
        "exact_paragraph": duplicate_metric(paragraph_df),
        "normalized_paragraph": duplicate_metric(normalized_paragraph_df),
        "exact_sentence": duplicate_metric(sentence_df),
        "normalized_sentence": duplicate_metric(normalized_sentence_df),
        "normalized_faq_pair": duplicate_metric(normalized_faq_df),
        "exact_hero_fact": duplicate_metric(exact_hero_df),
        "normalized_hero_fact": duplicate_metric(normalized_hero_df),
        "exact_source_note": duplicate_metric(exact_note_df),
        "normalized_source_note": duplicate_metric(normalized_note_df),
    }
    limits = {
        "exact_paragraph": 10, "normalized_paragraph": 50,
        "exact_sentence": 50, "normalized_sentence": 100,
        "normalized_faq_pair": 50,
        "exact_hero_fact": 3, "normalized_hero_fact": 15,
        "exact_source_note": 15,
    }
    for name, limit in limits.items():
        if duplicate_metrics[name]["max_df"] > limit:
            audit.add("cross_page_copy_overuse", f"{name}: max_df={duplicate_metrics[name]['max_df']} > {limit}")
    return {
        "scoped_pages": len(pages), "detail_pages": len(details),
        "unique_titles": len(set(titles)), "unique_h1": len(set(h1s)),
        "unique_descriptions": len(set(descriptions)),
        "physical_org_keys": len(org_key_ids), "physical_org_ids": len(org_id_keys),
        "keyword_exact_h1": {key: distribution(value) for key, value in keyword_counts.items()},
        "locality_per_100_tokens": {key: distribution(value) for key, value in locality_density.items()},
        "duplicates": duplicate_metrics,
        "hidden_unique_bytes": sum(size for _, size in hidden_image_bytes),
        "hidden_unique_assets": len(hidden_image_bytes),
    }


def read_only_normalizer_check(audit: Audit) -> dict:
    command = [sys.executable, str(ROOT / "scripts" / "normalize_internal_links_and_social_meta.py"), "--check"]
    before = git_value("status", "--porcelain=v1", "--untracked-files=all")
    completed = subprocess.run(command, cwd=ROOT, text=True, encoding="utf-8", capture_output=True, timeout=180)
    after = git_value("status", "--porcelain=v1", "--untracked-files=all")
    if before != after:
        audit.add("readonly_check_mutated_repo", "normalize --check changed git status")
    if completed.returncode:
        audit.add("normalizer_check", (completed.stdout + completed.stderr)[-800:])
    return {"returncode": completed.returncode, "stdout_tail": completed.stdout[-500:]}


def copy_for_sandbox(source: str, destination: str) -> str:
    src = Path(source)
    # Generated/normalized text and representative targets must not share an inode.
    mutable_suffixes = {".html", ".py", ".css", ".js", ".json", ".xml", ".txt", ".csv"}
    if src.suffix.lower() in mutable_suffixes or "representative" in {part.lower() for part in src.parts}:
        return shutil.copy2(source, destination)
    try:
        os.link(source, destination)
        return destination
    except OSError:
        return shutil.copy2(source, destination)


def run_idempotency_sandbox(audit: Audit) -> dict:
    before_status = git_value("status", "--porcelain=v1", "--untracked-files=all")
    before_hash = set_sha([f"{path.relative_to(ROOT).as_posix()}\t{sha256_file(path)}" for path in scoped_pages()])
    base = Path(tempfile.mkdtemp(prefix="new14-kem-idempotency-"))
    sandbox = base / ROOT.name
    common_target = base / ROOT.parent.joinpath("참고자료").name
    try:
        baseline_path = base / "preapply-baseline.json"
        baseline_path.write_text(
            json.dumps(build_baseline(), ensure_ascii=False, sort_keys=True, separators=(",", ":")),
            encoding="utf-8",
        )
        shutil.copytree(
            ROOT, sandbox, copy_function=copy_for_sandbox,
            ignore=shutil.ignore_patterns(".git", ".vercel", "__pycache__", "*.pyc", ".env*"),
        )
        shutil.copytree(ROOT.parent / "참고자료", common_target, copy_function=copy_for_sandbox)
        env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8")
        scoped_relatives = {path.relative_to(sandbox).as_posix() for path in scoped_pages(sandbox)}
        outside_before = {
            path.relative_to(sandbox).as_posix(): sha256_file(path)
            for path in sandbox.rglob("index.html")
            if path.relative_to(sandbox).as_posix() not in scoped_relatives
        }
        sitemap_targets = {page_url(path, sandbox) for path in scoped_pages(sandbox)}
        initial_sitemap_source = (sandbox / "sitemap.xml").read_text(encoding="utf-8")

        def sitemap_pairs(value: str) -> list[tuple[str, str]]:
            return [
                (html.unescape(location), modified)
                for location, modified in re.findall(
                    r"<url><loc>(.*?)</loc><lastmod>(.*?)</lastmod></url>", value
                )
            ]

        initial_sitemap_pairs = sitemap_pairs(initial_sitemap_source)
        initial_sitemap_order = [location for location, _ in initial_sitemap_pairs]
        initial_non_target_lastmods = [
            pair for pair in initial_sitemap_pairs if pair[0] not in sitemap_targets
        ]

        def run_step(command: list[str], timeout: int = 600) -> None:
            completed = subprocess.run(
                command, cwd=sandbox, env=env, stdout=subprocess.PIPE,
                stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace", timeout=timeout,
            )
            if completed.returncode:
                output = ((completed.stdout or "") + "\n" + (completed.stderr or ""))[-4000:]
                raise RuntimeError(f"step failed ({completed.returncode}): {command!r}\n{output}")

        def finalize_scoped_lastmod() -> int:
            sitemap = sandbox / "sitemap.xml"
            source = sitemap.read_text(encoding="utf-8")
            seen: set[str] = set()

            def replace(match: re.Match[str]) -> str:
                location = html.unescape(match.group(1))
                if location not in sitemap_targets:
                    return match.group(0)
                seen.add(location)
                return match.group(0).replace(match.group(2), "2026-08-17", 1)

            updated = re.sub(
                r"<url><loc>(.*?)</loc><lastmod>(.*?)</lastmod></url>",
                replace, source,
            )
            if seen != sitemap_targets:
                raise RuntimeError(
                    f"scoped sitemap lastmod coverage: seen={len(seen)} targets={len(sitemap_targets)}"
                )
            if updated != source:
                sitemap.write_text(updated, encoding="utf-8")
            final_pairs = sitemap_pairs(updated)
            if [location for location, _ in final_pairs] != initial_sitemap_order:
                raise RuntimeError("sitemap URL order changed during scoped lastmod update")
            final_non_target = [pair for pair in final_pairs if pair[0] not in sitemap_targets]
            if final_non_target != initial_non_target_lastmods:
                raise RuntimeError("non-target sitemap lastmod changed")
            if any(modified != "2026-08-17" for location, modified in final_pairs if location in sitemap_targets):
                raise RuntimeError("target sitemap lastmod was not synchronized")
            return len(seen)

        def pipeline() -> None:
            for profile in ("high", "middle", "elementary"):
                run_step(
                    [sys.executable, "scripts/generate_highschool_korean_english_math.py", profile, "--apply"],
                )
            run_step(
                [sys.executable, "scripts/fix_existing_kem_schema.py", "--apply"],
            )
            finalize_scoped_lastmod()

        pipeline()
        pass_one = {path.relative_to(sandbox).as_posix(): sha256_file(path) for path in scoped_pages(sandbox)}
        outside_after_one = {
            path.relative_to(sandbox).as_posix(): sha256_file(path)
            for path in sandbox.rglob("index.html")
            if path.relative_to(sandbox).as_posix() not in scoped_relatives
        }
        outside_changed = sorted(
            path for path in set(outside_before) | set(outside_after_one)
            if outside_before.get(path) != outside_after_one.get(path)
        )
        if outside_changed:
            audit.add("pipeline_scope_escape", f"{len(outside_changed)} HTML files; first={outside_changed[:5]}")
        pipeline()
        pass_two = {path.relative_to(sandbox).as_posix(): sha256_file(path) for path in scoped_pages(sandbox)}
        changed = sorted(path for path in set(pass_one) | set(pass_two) if pass_one.get(path) != pass_two.get(path))
        if changed:
            audit.add("generator_postprocessor_not_idempotent", f"{len(changed)} files; first={changed[:5]}")
        current = {path.relative_to(ROOT).as_posix(): sha256_file(path) for path in scoped_pages()}
        drift = sorted(path for path in set(current) | set(pass_one) if current.get(path) != pass_one.get(path))
        if drift:
            audit.add("generated_output_drift", f"{len(drift)} files; first={drift[:5]}")
        # Reuse this exact auditor against the twice-generated sandbox.  A
        # private copy of .git makes its read-only provenance checks work while
        # keeping every Git write, if any, isolated from the real repository.
        shutil.copytree(ROOT / ".git", sandbox / ".git", copy_function=shutil.copy2)
        sandbox_report_path = base / "sandbox-strict.json"
        strict = subprocess.run(
            [
                sys.executable, "scripts/audit_existing_kem_release.py",
                "--baseline", str(baseline_path), "--report", str(sandbox_report_path),
            ],
            cwd=sandbox, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", errors="replace", timeout=600,
        )
        sandbox_report = (
            json.loads(sandbox_report_path.read_text(encoding="utf-8"))
            if sandbox_report_path.exists() else
            {"status": "missing-report", "stderr": strict.stderr[-1000:]}
        )
        if strict.returncode or sandbox_report.get("status") != "PASS":
            audit.add(
                "sandbox_strict_failed",
                f"returncode={strict.returncode}, codes={sandbox_report.get('error_codes', {})}",
            )
        repo_after = set_sha([f"{path}\t{digest}" for path, digest in current.items()])
        return {
            "pipeline": [
                "generator high --apply", "generator middle --apply",
                "generator elementary --apply", "targeted schema fixer --apply",
                "scoped sitemap lastmod 2026-08-17",
            ],
            "legacy_global_postprocessor": "excluded: writes outside the 1,113-page scope",
            "outside_scope_html_changed": len(outside_changed),
            "scoped_sitemap_lastmod_updated": len(sitemap_targets),
            "pass1_vs_pass2_changed": len(changed), "current_vs_pass1_changed": len(drift),
            "repo_scoped_hash_before": before_hash, "repo_scoped_hash_after": repo_after,
            "sandbox_pass1_hash": set_sha([f"{path}\t{digest}" for path, digest in pass_one.items()]),
            "sandbox_pass2_hash": set_sha([f"{path}\t{digest}" for path, digest in pass_two.items()]),
            "sandbox_strict": {
                "status": sandbox_report.get("status"),
                "error_count": sandbox_report.get("error_count"),
                "error_codes": sandbox_report.get("error_codes", {}),
                "manifest": sandbox_report.get("manifest", {}),
                "metrics": sandbox_report.get("metrics", {}),
                "freeze": {
                    key: sandbox_report.get("freeze", {}).get(key)
                    for key in (
                        "before_html_manifest_sha256",
                        "after_html_manifest_sha256",
                        "before_sitemap_file_sha256",
                        "after_sitemap_file_sha256",
                    )
                },
            },
        }
    except (RuntimeError, subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as exc:
        audit.add("idempotency_sandbox_failed", str(exc))
        return {"failed": str(exc)}
    finally:
        resolved = base.resolve()
        temp_root = Path(tempfile.gettempdir()).resolve()
        if resolved != temp_root and temp_root in resolved.parents:
            def unlock_and_retry(function, value, _exception) -> None:
                os.chmod(value, stat.S_IWRITE | stat.S_IREAD)
                function(value)
            try:
                shutil.rmtree(resolved, onexc=unlock_and_retry)
            except OSError as exc:
                audit.add("idempotency_sandbox_cleanup", f"{resolved}: {exc}")
        after_status = git_value("status", "--porcelain=v1", "--untracked-files=all")
        after_hash = set_sha([f"{path.relative_to(ROOT).as_posix()}\t{sha256_file(path)}" for path in scoped_pages()])
        if before_status != after_status or before_hash != after_hash:
            audit.add("idempotency_mutated_repo", "working tree status/hash changed")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline", type=Path, help="Existing manifest to enforce")
    parser.add_argument("--write-baseline", type=Path, help="Write manifest outside the repository")
    parser.add_argument("--report", type=Path, help="Optional JSON report outside the repository")
    parser.add_argument("--check-idempotency", action="store_true")
    args = parser.parse_args()

    for destination in (args.write_baseline, args.report):
        if destination:
            resolved = destination.resolve()
            try:
                resolved.relative_to(ROOT.resolve())
            except ValueError:
                pass
            else:
                raise SystemExit(f"Refusing to write audit output inside repository: {resolved}")

    initial_git_status = git_value("status", "--porcelain=v1", "--untracked-files=all")
    current_manifest = build_baseline()
    if args.write_baseline:
        args.write_baseline.parent.mkdir(parents=True, exist_ok=True)
        args.write_baseline.write_text(
            json.dumps(current_manifest, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
            encoding="utf-8",
        )

    audit = Audit()
    baseline_diff = {"status": "not_supplied"}
    if args.baseline:
        baseline = json.loads(args.baseline.read_text(encoding="utf-8"))
        baseline_diff = compare_baseline(audit, baseline, current_manifest)
    rows, source_by_slug, groups = load_source(audit)
    source_summary = {
        "rows": len(rows), "localities": len(source_by_slug), "physical_centers": len(groups),
        "physical_group_sizes": dict(sorted(Counter(len(value) for value in groups.values()).items())),
        "fee_blank_rows": sum(not clean(row.get("센터 교습비")) for row in rows),
        "school_blank_rows": {
            name: sum(not clean(row.get(config["school_column"], "")) for row in rows)
            for name, config in PROFILES.items()
        },
        "school_unusable_rows": {
            name: sum(not split_schools(row.get(config["school_column"], "")) for row in rows)
            for name, config in PROFILES.items()
        },
        "grade_blank_rows": {
            name: {
                subject: sum(not grades(row.get(column, ""), config["prefix"]) for row in rows)
                for subject, column in SUBJECT_COLUMNS.items()
            }
            for name, config in PROFILES.items()
        },
    }
    metrics = audit_site(audit, source_by_slug, groups)
    normalizer = read_only_normalizer_check(audit)
    idempotency = run_idempotency_sandbox(audit) if args.check_idempotency else {"status": "not_requested"}
    final_manifest = build_baseline()
    for key in (
        "url_set_sha256", "canonical_set_sha256", "sitemap_scoped_set_sha256",
        "sitemap_scoped_lastmod_sha256", "html_manifest_sha256", "sitemap_file_sha256",
    ):
        if current_manifest[key] != final_manifest[key]:
            audit.add("audit_freeze_changed", f"{key}: {current_manifest[key]} -> {final_manifest[key]}")
    status = git_value("status", "--porcelain=v1", "--untracked-files=all")
    report = {
        "status": "PASS" if not audit.total else "HOLD",
        "root": str(ROOT), "head": git_value("rev-parse", "HEAD"),
        "branch": git_value("branch", "--show-current"),
        "upstream": optional_git_value("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"),
        "origin_main": optional_git_value("rev-parse", "origin/main"),
        "origin_url": optional_git_value("remote", "get-url", "origin"),
        "source": source_summary,
        "manifest": {key: current_manifest[key] for key in (
            "counts", "url_set_sha256", "canonical_set_sha256", "sitemap_scoped_set_sha256",
            "sitemap_scoped_lastmod_sha256", "html_manifest_sha256", "sitemap_file_sha256",
        )},
        "freeze": {
            "before_html_manifest_sha256": current_manifest["html_manifest_sha256"],
            "after_html_manifest_sha256": final_manifest["html_manifest_sha256"],
            "before_sitemap_file_sha256": current_manifest["sitemap_file_sha256"],
            "after_sitemap_file_sha256": final_manifest["sitemap_file_sha256"],
            "git_status_before": initial_git_status.splitlines(),
            "git_status_after": status.splitlines(),
        },
        "baseline_diff": baseline_diff,
        "metrics": metrics, "normalizer": normalizer, "idempotency": idempotency,
        "git_status": status.splitlines(), **audit.payload(),
    }
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    print(rendered)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(rendered + "\n", encoding="utf-8")
    raise SystemExit(0 if not audit.total else 1)


if __name__ == "__main__":
    main()
