from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import os
import re
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
COMMON = ROOT.parent / "참고자료" / "공통자료"
CENTER_CSV = COMMON / "센터정보 정리.csv"
BASE_URL = "https://xn--3e0bz50b1zcyxat54c.com"
ROOT_ORGANIZATION_ID = f"{BASE_URL}/#organization"
EXPECTED_LOCALITIES = 371
EXPECTED_PAGES = EXPECTED_LOCALITIES * 3

GRAPH_RE = re.compile(
    r'(<script\s+type="application/ld\+json">)(.*?)(</script>)',
    re.DOTALL,
)


@dataclass(frozen=True)
class Profile:
    slug: str
    category: str
    level: str
    course: str
    grade_prefix: str
    school_label: str
    school_column: str


PROFILES = (
    Profile(
        "고등학생국영수학원",
        "고등학생 국영수학원",
        "고등학생",
        "고등",
        "고",
        "고등학교 참고",
        "타깃학교\n(고)",
    ),
    Profile(
        "중학생국영수학원",
        "중학생 국영수학원",
        "중학생",
        "중등",
        "중",
        "중학교 참고",
        "타깃학교\n(중)",
    ),
    Profile(
        "초등학생국영수학원",
        "초등학생 국영수학원",
        "초등학생",
        "초등",
        "초",
        "초등학교 참고",
        "타깃학교\n(초)",
    ),
)

SUBJECT_COLUMNS = (
    ("국어", "가능학년\n(국어)", "korean"),
    ("영어", "가능학년\n(영어)", "english"),
    ("수학", "가능학년\n(수학)", "math"),
)

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

REGION_NAMES = {
    "서울": "서울특별시",
    "서울특별시": "서울특별시",
    "경기": "경기도",
    "경기도": "경기도",
    "인천": "인천광역시",
    "인천광역시": "인천광역시",
    "부산": "부산광역시",
    "부산광역시": "부산광역시",
    "대구": "대구광역시",
    "대구광역시": "대구광역시",
    "대전": "대전광역시",
    "대전광역시": "대전광역시",
    "광주": "광주광역시",
    "광주광역시": "광주광역시",
    "울산": "울산광역시",
    "울산광역시": "울산광역시",
    "세종": "세종특별자치시",
    "세종시": "세종특별자치시",
    "세종특별자치시": "세종특별자치시",
    "강원": "강원특별자치도",
    "강원도": "강원특별자치도",
    "강원특별자치도": "강원특별자치도",
    "충북": "충청북도",
    "충청북도": "충청북도",
    "충남": "충청남도",
    "충청남도": "충청남도",
    "전북": "전북특별자치도",
    "전라북도": "전북특별자치도",
    "전북특별자치도": "전북특별자치도",
    "전남": "전라남도",
    "전라남도": "전라남도",
    "경북": "경상북도",
    "경상북도": "경상북도",
    "경남": "경상남도",
    "경상남도": "경상남도",
    "제주": "제주특별자치도",
    "제주도": "제주특별자치도",
    "제주특별자치도": "제주특별자치도",
}

SCHOOL_SUFFIX_RE = re.compile(r"(?:초등학교|중학교|고등학교|초|중|고)$")


class GateError(RuntimeError):
    pass


@dataclass(frozen=True)
class SourceRow:
    locality: str
    slug: str
    region: str
    district: str
    center_name: str
    fee_url: str
    office_name: str
    registration: str
    address: str
    schools: dict[str, tuple[str, ...]]
    school_legacy: dict[str, tuple[str, ...]]
    grades: dict[str, tuple[str, ...]]

    @property
    def center_key(self) -> str:
        return "|".join((self.center_name, self.address, self.registration))


@dataclass(frozen=True)
class PhysicalCenter:
    key: str
    center_name: str
    address: str
    office_name: str
    registration: str
    areas: tuple[str, ...]
    grades: dict[str, tuple[str, ...]]
    fee_urls: tuple[str, ...]

    @property
    def entity_id(self) -> str:
        token = hashlib.sha256(self.key.encode("utf-8")).hexdigest()[:16]
        return f"{BASE_URL}/#center-{token}"


@dataclass
class Candidate:
    path: Path
    relative_path: str
    profile: Profile
    source: SourceRow
    physical: PhysicalCenter
    original: str
    transformed: str
    source_sha256: str
    before_route_manifest: dict[str, object]
    after_route_manifest: dict[str, object]


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def clean_html(value: str) -> str:
    return clean(html.unescape(re.sub(r"<[^>]+>", "", value or "")))


def list_values(value: str) -> tuple[str, ...]:
    return tuple(
        dict.fromkeys(
            item.strip()
            for item in re.split(r"[,/\n]+", value or "")
            if item.strip()
        )
    )


def split_schools(value: str) -> tuple[str, ...]:
    """Mirror the release auditor's authoritative compound-school parser."""
    normalized = clean(value)
    if not normalized or "지역내 모든 고등학교 가능" in normalized:
        return ()
    parts = [
        clean(part)
        for part in re.split(r"[,/·.，\n]+", normalized)
        if clean(part)
    ]
    result: list[str] = []
    for part in parts:
        tokens = part.split()
        if len(tokens) > 1 and all(SCHOOL_SUFFIX_RE.search(token) for token in tokens):
            result.extend(tokens)
        else:
            result.append(part)
    return ordered_unique(result)


def locality_slug(value: str) -> str:
    return re.sub(r"\s+", "", value)


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def read_utf8_exact(path: Path) -> str:
    # Path.read_text() performs universal-newline translation. Reading bytes
    # keeps the non-schema hash gate sensitive to any original CRLF/LF bytes.
    return path.read_bytes().decode("utf-8")


def normalized_url_path(value: str) -> str:
    parsed = urlsplit(value)
    path = unquote(parsed.path).rstrip("/") + "/"
    return path


def node_types(node: dict) -> set[str]:
    value = node.get("@type", [])
    return set(value if isinstance(value, list) else [value])


def get_node(nodes: list[dict], expected_type: str, path: Path) -> dict:
    matches = [node for node in nodes if expected_type in node_types(node)]
    if len(matches) != 1:
        raise GateError(
            f"{path}: expected one {expected_type} node, found {len(matches)}"
        )
    return matches[0]


def grade_sort_key(value: str) -> tuple[int, str]:
    order = {
        **{f"초{number}": number for number in range(1, 7)},
        **{f"중{number}": 10 + number for number in range(1, 4)},
        **{f"고{number}": 20 + number for number in range(1, 4)},
    }
    return order.get(value, 99), value


def ordered_unique(values: Iterable[str]) -> tuple[str, ...]:
    return tuple(dict.fromkeys(value for value in values if value))


def load_sources() -> tuple[dict[str, SourceRow], dict[str, PhysicalCenter]]:
    if not CENTER_CSV.exists():
        raise GateError(f"authoritative source missing: {CENTER_CSV}")
    with CENTER_CSV.open(encoding="utf-8-sig", newline="") as handle:
        raw_rows = list(csv.DictReader(handle))
    if len(raw_rows) != EXPECTED_LOCALITIES:
        raise GateError(
            f"expected {EXPECTED_LOCALITIES} source rows, found {len(raw_rows)}"
        )

    required_columns = {
        "근처 수업가능 동네",
        "지역",
        "시or구",
        "센터명",
        "센터 교습비",
        "교육지원청명칭",
        "교육지원청 등록번호",
        "센터 주소",
        *(profile.school_column for profile in PROFILES),
        *(column for _, column, _ in SUBJECT_COLUMNS),
    }
    missing_columns = required_columns - set(raw_rows[0])
    if missing_columns:
        raise GateError(f"source columns missing: {sorted(missing_columns)}")

    rows: list[SourceRow] = []
    for raw in raw_rows:
        locality = clean(raw["근처 수업가능 동네"])
        row = SourceRow(
            locality=locality,
            slug=locality_slug(locality),
            region=clean(raw["지역"]),
            district=clean(raw["시or구"]),
            center_name=clean(raw["센터명"]),
            fee_url=clean(raw["센터 교습비"]),
            office_name=clean(raw["교육지원청명칭"]),
            registration=clean(raw["교육지원청 등록번호"]),
            address=clean(raw["센터 주소"]),
            schools={
                profile.slug: split_schools(raw[profile.school_column])
                for profile in PROFILES
            },
            school_legacy={
                profile.slug: list_values(raw[profile.school_column])
                for profile in PROFILES
            },
            grades={
                subject: list_values(raw[column])
                for subject, column, _ in SUBJECT_COLUMNS
            },
        )
        if not all(
            (row.locality, row.center_name, row.address, row.registration)
        ):
            raise GateError(
                f"source identity fact missing for locality {row.locality!r}"
            )
        rows.append(row)

    by_slug = {row.slug: row for row in rows}
    if len(by_slug) != EXPECTED_LOCALITIES:
        raise GateError("source locality slugs are not unique")
    if len({row.locality for row in rows}) != EXPECTED_LOCALITIES:
        raise GateError("source locality names are not unique")

    grouped: dict[str, list[SourceRow]] = defaultdict(list)
    for row in rows:
        grouped[row.center_key].append(row)

    physical_by_key: dict[str, PhysicalCenter] = {}
    for key, group in grouped.items():
        office_names = {row.office_name for row in group}
        if len(office_names) != 1:
            raise GateError(
                f"conflicting registration names for physical center {key}: "
                f"{sorted(office_names)}"
            )
        grades = {
            subject: tuple(
                sorted(
                    set(grade for row in group for grade in row.grades[subject]),
                    key=grade_sort_key,
                )
            )
            for subject, _, _ in SUBJECT_COLUMNS
        }
        physical_by_key[key] = PhysicalCenter(
            key=key,
            center_name=group[0].center_name,
            address=group[0].address,
            office_name=group[0].office_name,
            registration=group[0].registration,
            areas=ordered_unique(row.locality for row in group),
            grades=grades,
            fee_urls=ordered_unique(row.fee_url for row in group),
        )

    if len(physical_by_key) != 188:
        raise GateError(
            f"expected 188 physical centers, found {len(physical_by_key)}"
        )
    return by_slug, physical_by_key


def discover_pages(by_slug: dict[str, SourceRow]) -> list[tuple[Path, Profile, SourceRow]]:
    expected_slugs = set(by_slug)
    pages: list[tuple[Path, Profile, SourceRow]] = []
    for profile in PROFILES:
        category = ROOT / "과목별학원" / profile.slug
        actual = {
            path.parent.name: path
            for path in category.glob("*/index.html")
            if path.parent != category
        }
        missing = expected_slugs - set(actual)
        extra = set(actual) - expected_slugs
        if missing or extra:
            raise GateError(
                f"{profile.slug}: URL/locality manifest mismatch; "
                f"missing={sorted(missing)[:8]}, extra={sorted(extra)[:8]}"
            )
        if len(actual) != EXPECTED_LOCALITIES:
            raise GateError(
                f"{profile.slug}: expected {EXPECTED_LOCALITIES} details, "
                f"found {len(actual)}"
            )
        pages.extend(
            (actual[slug], profile, by_slug[slug])
            for slug in sorted(actual)
        )
    if len(pages) != EXPECTED_PAGES:
        raise GateError(f"expected {EXPECTED_PAGES} pages, found {len(pages)}")
    return pages


def extract_graph(text: str, path: Path) -> dict:
    matches = list(GRAPH_RE.finditer(text))
    if len(matches) != 1:
        raise GateError(
            f"{path}: expected one JSON-LD script, found {len(matches)}"
        )
    try:
        graph = json.loads(matches[0].group(2))
    except json.JSONDecodeError as error:
        raise GateError(f"{path}: invalid JSON-LD: {error}") from error
    if not isinstance(graph, dict) or not isinstance(graph.get("@graph"), list):
        raise GateError(f"{path}: JSON-LD @graph missing")
    return graph


def replace_graph(text: str, graph: dict, path: Path) -> str:
    payload = json.dumps(
        graph,
        ensure_ascii=False,
        separators=(",", ":"),
    ).replace("</", "<\\/")
    updated, count = GRAPH_RE.subn(
        lambda match: match.group(1) + payload + match.group(3),
        text,
        count=1,
    )
    if count != 1:
        raise GateError(f"{path}: JSON-LD replacement failed")
    return updated


def html_without_graph(text: str, path: Path) -> str:
    stripped, count = GRAPH_RE.subn(
        lambda match: match.group(1) + "__JSON_LD_PAYLOAD__" + match.group(3),
        text,
        count=1,
    )
    if count != 1:
        raise GateError(f"{path}: JSON-LD isolation failed")
    return stripped


def expected_path(profile: Profile, source: SourceRow) -> str:
    return f"/과목별학원/{profile.slug}/{source.slug}/"


def route_manifest(
    text: str,
    graph: dict,
    path: Path,
    profile: Profile,
    source: SourceRow,
) -> dict[str, object]:
    canonical_match = re.search(
        r'<link\s+rel="canonical"\s+href="([^"]+)"', text
    )
    if not canonical_match:
        raise GateError(f"{path}: canonical missing")
    canonical = canonical_match.group(1)
    route = expected_path(profile, source)
    if normalized_url_path(canonical) != route:
        raise GateError(
            f"{path}: canonical path mismatch: {canonical!r} != {route!r}"
        )

    nodes = graph["@graph"]
    webpage = get_node(nodes, "WebPage", path)
    article = get_node(nodes, "Article", path)
    service = get_node(nodes, "Service", path)
    breadcrumb = get_node(nodes, "BreadcrumbList", path)
    related = get_node(nodes, "ItemList", path)
    if normalized_url_path(str(webpage.get("url", ""))) != route:
        raise GateError(f"{path}: WebPage.url path mismatch")

    return {
        "relative_path": path.relative_to(ROOT).as_posix(),
        "canonical": canonical,
        "webpage_id": webpage.get("@id"),
        "webpage_url": webpage.get("url"),
        "article_id": article.get("@id"),
        "article_main": article.get("mainEntityOfPage"),
        "service_id": service.get("@id"),
        "breadcrumb": breadcrumb.get("itemListElement"),
        "related": related.get("itemListElement"),
        "non_schema_sha256": sha256_text(html_without_graph(text, path)),
    }


def expected_level_grades(
    grades: dict[str, tuple[str, ...]], profile: Profile
) -> dict[str, tuple[str, ...]]:
    return {
        subject: tuple(
            grade for grade in grades[subject]
            if grade.startswith(profile.grade_prefix)
        )
        for subject, _, _ in SUBJECT_COLUMNS
    }


def expected_article_levels(source: SourceRow, profile: Profile) -> list[str]:
    return sorted(
        {
            grade
            for values in expected_level_grades(source.grades, profile).values()
            for grade in values
        },
        key=grade_sort_key,
    )


def official_address(source: SourceRow) -> tuple[str, str]:
    tokens = source.address.replace(",", " ").split()
    region = REGION_NAMES.get(tokens[0], "") if tokens else ""
    if not region:
        raise GateError(
            f"unknown official address region for {source.locality}: {source.address}"
        )
    if region == "세종특별자치시":
        locality = source.locality
    elif len(tokens) > 1:
        locality = tokens[1]
    else:
        locality = ""
    if not locality:
        raise GateError(
            f"official address locality missing for {source.locality}: {source.address}"
        )
    return region, locality


def expected_visible_region(source: SourceRow) -> str:
    """Mirror the generator's reader-facing administrative-region label.

    The CSV's ``지역`` value remains the routing/taxonomy fact.  Only the
    visible label expands it to an official province/metropolitan-city name;
    the broad 충청/경상/전라 buckets are disambiguated by the verified street
    address.  Sejong's CSV district is a road field, so it is not displayed.
    A locality prefix already represented by its district/region is removed
    only from this reader-facing join; the source locality and URL stay raw.
    """
    if source.region in {"충청", "경상", "전라"}:
        display_region, _ = official_address(source)
    else:
        display_region = REGION_NAMES.get(source.region, "")
    if not display_region:
        raise GateError(
            f"unknown visible source region for {source.locality}: "
            f"{source.region!r} / {source.address!r}"
        )
    display_district = (
        "" if display_region == "세종특별자치시" else source.district
    )
    display_locality = source.locality
    prefixes = (
        re.sub(r"(?:시|군|구)$", "", clean(source.district)),
        clean(source.region),
    )
    for prefix in dict.fromkeys(value for value in prefixes if value):
        match = re.fullmatch(
            rf"{re.escape(prefix)}\s+(.+)", clean(source.locality)
        )
        if match:
            display_locality = clean(match.group(1))
            break
    return " ".join(
        value
        for value in (display_region, display_district, display_locality)
        if value
    )


def article_school_names(
    article: dict,
    profile: Profile,
    source: SourceRow,
    path: Path,
) -> tuple[str, ...]:
    mentions = article.get("mentions", [])
    if not isinstance(mentions, list):
        raise GateError(f"{path}: Article.mentions is not a list")
    locality_index = next(
        (
            index
            for index, item in enumerate(mentions)
            if isinstance(item, dict)
            and item.get("@type") == "Thing"
            and clean(item.get("name")) == source.locality
        ),
        None,
    )
    topic_names = {
        f"{profile.course} {subject}" for subject, _, _ in SUBJECT_COLUMNS
    } | {"학교 학습 대비", "오답 재학습", "학습코칭"}
    topic_index = next(
        (
            index
            for index, item in enumerate(mentions)
            if isinstance(item, dict)
            and clean(item.get("name")) in topic_names
        ),
        len(mentions),
    )
    if locality_index is None or topic_index < locality_index:
        raise GateError(f"{path}: Article school mention boundaries missing")
    return tuple(
        clean(item.get("name"))
        for item in mentions[locality_index + 1:topic_index]
        if isinstance(item, dict) and item.get("@type") == "Thing"
    )


def repair_article_school_mentions(
    article: dict,
    profile: Profile,
    source: SourceRow,
    path: Path,
) -> None:
    mentions = article.get("mentions", [])
    if not isinstance(mentions, list):
        raise GateError(f"{path}: Article.mentions is not a list")
    legacy_names = set(source.school_legacy[profile.slug]) | set(
        source.schools[profile.slug]
    )
    retained = [
        item
        for item in mentions
        if not (
            isinstance(item, dict)
            and item.get("@type") == "Thing"
            and clean(item.get("name")) in legacy_names
        )
    ]
    topic_names = {
        f"{profile.course} {subject}" for subject, _, _ in SUBJECT_COLUMNS
    } | {"학교 학습 대비", "오답 재학습", "학습코칭"}
    insert_at = next(
        (
            index
            for index, item in enumerate(retained)
            if isinstance(item, dict)
            and clean(item.get("name")) in topic_names
        ),
        len(retained),
    )
    schools = [
        {"@type": "Thing", "name": school}
        for school in source.schools[profile.slug]
    ]
    article["mentions"] = retained[:insert_at] + schools + retained[insert_at:]


def expected_organization_teaches(physical: PhysicalCenter) -> list[str]:
    teaches: list[str] = []
    for profile in PROFILES:
        level_grades = expected_level_grades(physical.grades, profile)
        teaches.extend(
            f"{profile.course} {subject}"
            for subject, _, _ in SUBJECT_COLUMNS
            if level_grades[subject]
        )
    # A literal "초·중·고" label is only source-grounded when the group has
    # at least one listed grade in every school level. Partial availability is
    # already represented by the precise level-specific labels above.
    if all(
        any(grade.startswith(profile.grade_prefix) for grade in physical.grades["영어"])
        for profile in PROFILES
    ):
        teaches.append("초·중·고 영어")
    if all(
        any(grade.startswith(profile.grade_prefix) for grade in physical.grades["수학"])
        for profile in PROFILES
    ):
        teaches.append("초·중·고 수학")
    teaches.append("학습코칭")
    return teaches


def subject_offer(
    name: str,
    grades: tuple[str, ...],
    *,
    item_id: str | None = None,
    provider_id: str | None = None,
    area: str | None = None,
) -> dict:
    item: dict[str, object] = {
        "@type": "Service",
        "name": name,
    }
    if item_id:
        item["@id"] = item_id
    if provider_id:
        item["provider"] = {"@id": provider_id}
    if area:
        item["areaServed"] = [area]
    offer: dict[str, object] = {
        "@type": "Offer",
        "name": name,
        "itemOffered": item,
        "eligibleCustomerType": "·".join(grades),
    }
    return offer


def fee_offer(center_name: str, url: str) -> dict:
    return {
        "@type": "Offer",
        "name": f"{center_name} 교습과정·교습비 확인",
        "url": url,
        "itemOffered": {
            "@type": "Service",
            "name": "센터별 교습과정 정보 확인",
        },
    }


def expected_organization_offers(physical: PhysicalCenter) -> list[dict]:
    offers: list[dict] = []
    for profile in PROFILES:
        level_grades = expected_level_grades(physical.grades, profile)
        for subject, _, _ in SUBJECT_COLUMNS:
            grades = level_grades[subject]
            if grades:
                name = f"{profile.course} {subject} 학습 상담"
                offers.append(subject_offer(name, grades))
    for subject in ("영어", "수학"):
        grades = physical.grades[subject]
        if grades:
            offers.append(subject_offer(f"{subject} 학습 상담", grades))
    offers.extend(
        fee_offer(physical.center_name, url)
        for url in physical.fee_urls
    )
    return offers


def expected_page_offers(
    source: SourceRow,
    physical: PhysicalCenter,
    profile: Profile,
    service_id: str,
) -> list[dict]:
    level_grades = expected_level_grades(source.grades, profile)
    offers: list[dict] = []
    codes = {code: subject for subject, _, code in SUBJECT_COLUMNS}
    for code, subject in codes.items():
        grades = level_grades[subject]
        if not grades:
            continue
        name = f"{profile.course} {subject} 학습 상담"
        offers.append(
            subject_offer(
                name,
                grades,
                item_id=f"{service_id}-{code}",
                provider_id=physical.entity_id,
                area=source.locality,
            )
        )
    if source.fee_url:
        offers.append(fee_offer(source.center_name, source.fee_url))
    return offers


def visible_facts(text: str, path: Path) -> dict[str, str]:
    match = re.search(
        r'<dl class="local-facts">(.*?)</dl>', text, re.DOTALL
    )
    if not match:
        raise GateError(f"{path}: visible local facts missing")
    return {
        clean_html(label): value
        for label, value in re.findall(
            r'<div><dt>(.*?)</dt><dd>(.*?)</dd></div>',
            match.group(1),
            re.DOTALL,
        )
    }


def validate_visible_source(
    text: str,
    path: Path,
    profile: Profile,
    source: SourceRow,
) -> None:
    facts = visible_facts(text, path)
    expected = {
        "지역": expected_visible_region(source),
        "센터 기준": source.center_name,
        "제공 주소": source.address,
        "등록 정보": source.registration,
    }
    for label, value in expected.items():
        actual = clean_html(facts.get(label, ""))
        if actual != value:
            raise GateError(
                f"{path}: visible {label} mismatch: {actual!r} != {value!r}"
            )

    school_html = facts.get(profile.school_label, "")
    visible_school_tokens = tuple(
        clean_html(value)
        for value in re.findall(r"<span>(.*?)</span>", school_html, re.DOTALL)
    )
    actual_schools = ordered_unique(
        school
        for token in visible_school_tokens
        for school in split_schools(token)
    )
    expected_schools = source.schools[profile.slug]
    if actual_schools != expected_schools:
        raise GateError(
            f"{path}: visible schools mismatch: "
            f"{actual_schools!r} != {expected_schools!r}"
        )
    legacy_all_school_claim = any(
        "지역내 모든 고등학교 가능" in token
        for token in visible_school_tokens
    )
    if (
        not expected_schools
        and "제공 목록 없음" not in clean_html(school_html)
        and not legacy_all_school_claim
    ):
        raise GateError(f"{path}: missing blank-school disclosure")

    grade_match = re.search(
        r'<ul class="grade-list">(.*?)</ul>', text, re.DOTALL
    )
    if not grade_match:
        raise GateError(f"{path}: visible grade list missing")
    actual_grades = {
        clean_html(subject): clean_html(value)
        for subject, value in re.findall(
            r'<li><strong>(.*?)</strong><span>(.*?)</span></li>',
            grade_match.group(1),
            re.DOTALL,
        )
    }
    level_grades = expected_level_grades(source.grades, profile)
    expected_grades = {
        subject: "·".join(level_grades[subject])
        or f"{profile.course} 과정 미기재"
        for subject, _, _ in SUBJECT_COLUMNS
    }
    if actual_grades != expected_grades:
        raise GateError(
            f"{path}: visible grades mismatch: "
            f"{actual_grades!r} != {expected_grades!r}"
        )

    actual_fee_links = re.findall(
        r'<a class="button compact" href="([^"]+)"[^>]*>'
        r'센터별 교습비 확인',
        text,
    )
    expected_fee_links = [source.fee_url] if source.fee_url else []
    if actual_fee_links != expected_fee_links:
        raise GateError(
            f"{path}: visible fee link mismatch: "
            f"{actual_fee_links!r} != {expected_fee_links!r}"
        )
    if not source.fee_url and "교습비 자료는 희망 센터에서 확인합니다" not in text:
        raise GateError(f"{path}: missing blank-fee disclosure")


def validate_faq_parity(text: str, graph: dict, path: Path) -> None:
    faq_block = re.search(
        r'<div class="faq-list">(.*?)</div>', text, re.DOTALL
    )
    if not faq_block:
        raise GateError(f"{path}: visible FAQ list missing")
    visible = [
        (clean_html(question), clean_html(answer))
        for question, answer in re.findall(
            r'<details(?: open)?>\s*<summary>(.*?)</summary>'
            r'<p>(.*?)</p></details>',
            faq_block.group(1),
            re.DOTALL,
        )
    ]
    faq = get_node(graph["@graph"], "FAQPage", path)
    schema = [
        (
            str(item.get("name", "")),
            str(item.get("acceptedAnswer", {}).get("text", "")),
        )
        for item in faq.get("mainEntity", [])
        if isinstance(item, dict)
    ]
    if len(visible) != 4 or visible != schema:
        raise GateError(
            f"{path}: FAQ visible/schema mismatch: {len(visible)}/{len(schema)}"
        )


def validate_source_schema(
    graph: dict,
    path: Path,
    profile: Profile,
    source: SourceRow,
    physical: PhysicalCenter,
) -> None:
    nodes = graph["@graph"]
    types = set().union(*(node_types(item) for item in nodes))
    missing = REQUIRED_SCHEMA_TYPES - types
    if missing:
        raise GateError(f"{path}: required schema types missing: {sorted(missing)}")

    organization = get_node(nodes, "EducationalOrganization", path)
    if "LocalBusiness" not in node_types(organization):
        raise GateError(f"{path}: physical entity is not a LocalBusiness")
    if organization.get("@id") != physical.entity_id:
        raise GateError(f"{path}: physical organization ID mismatch")
    if clean(organization.get("name")) != physical.center_name:
        raise GateError(f"{path}: organization name mismatch")
    if clean(organization.get("address", {}).get("streetAddress")) != physical.address:
        raise GateError(f"{path}: organization address mismatch")
    identifier = organization.get("identifier", {})
    if (
        clean(identifier.get("name")) != physical.office_name
        or clean(identifier.get("value")) != physical.registration
    ):
        raise GateError(f"{path}: organization registration mismatch")
    if set(organization.get("areaServed", [])) != set(physical.areas):
        raise GateError(f"{path}: organization areaServed mismatch")

    article = get_node(nodes, "Article", path)
    expected_sections = [
        profile.category,
        source.region,
        source.district,
        source.locality,
    ]
    if article.get("articleSection") != expected_sections:
        raise GateError(f"{path}: Article.articleSection mismatch")
    expected_levels = {
        grade
        for grades in expected_level_grades(source.grades, profile).values()
        for grade in grades
    }
    if set(article.get("educationalLevel", [])) != expected_levels:
        raise GateError(f"{path}: Article.educationalLevel mismatch")

def transform_graph(
    graph: dict,
    path: Path,
    profile: Profile,
    source: SourceRow,
    physical: PhysicalCenter,
) -> dict:
    # Round-trip through JSON so the caller's source graph remains immutable.
    updated = json.loads(json.dumps(graph, ensure_ascii=False))
    nodes = updated["@graph"]
    organization = get_node(nodes, "EducationalOrganization", path)
    article = get_node(nodes, "Article", path)
    service = get_node(nodes, "Service", path)

    address = organization.get("address")
    if not isinstance(address, dict):
        raise GateError(f"{path}: physical organization address is not an object")
    address_region, address_locality = official_address(source)
    address["addressRegion"] = address_region
    address["addressLocality"] = address_locality
    address["addressCountry"] = "KR"

    organization.pop("telephone", None)
    organization["teaches"] = expected_organization_teaches(physical)
    organization["educationalLevel"] = sorted(
        set(
            grade
            for grades in physical.grades.values()
            for grade in grades
        ),
        key=grade_sort_key,
    )
    organization["makesOffer"] = expected_organization_offers(physical)

    article["educationalLevel"] = expected_article_levels(source, profile)
    repair_article_school_mentions(article, profile, source, path)
    article["author"] = {"@id": ROOT_ORGANIZATION_ID}
    article["publisher"] = {"@id": ROOT_ORGANIZATION_ID}

    service["provider"] = {"@id": physical.entity_id}
    service["areaServed"] = [source.locality]
    service["offers"] = expected_page_offers(
        source,
        physical,
        profile,
        str(service.get("@id", "")),
    )
    service.pop("makesOffer", None)
    return updated


def validate_transformed_schema(
    graph: dict,
    path: Path,
    profile: Profile,
    source: SourceRow,
    physical: PhysicalCenter,
) -> None:
    nodes = graph["@graph"]
    organization = get_node(nodes, "EducationalOrganization", path)
    article = get_node(nodes, "Article", path)
    service = get_node(nodes, "Service", path)

    if "telephone" in organization:
        raise GateError(f"{path}: physical organization telephone was not removed")
    address = organization.get("address")
    if not isinstance(address, dict):
        raise GateError(f"{path}: physical organization address is not an object")
    expected_region, expected_locality = official_address(source)
    if clean(address.get("addressRegion")) != expected_region:
        raise GateError(f"{path}: PostalAddress.addressRegion mismatch")
    if clean(address.get("addressLocality")) != expected_locality:
        raise GateError(f"{path}: PostalAddress.addressLocality mismatch")
    if clean(address.get("addressCountry")) != "KR":
        raise GateError(f"{path}: PostalAddress.addressCountry mismatch")
    if organization.get("teaches") != expected_organization_teaches(physical):
        raise GateError(f"{path}: organization teaches is not source-grounded")
    expected_levels = sorted(
        {grade for grades in physical.grades.values() for grade in grades},
        key=grade_sort_key,
    )
    if organization.get("educationalLevel") != expected_levels:
        raise GateError(f"{path}: organization educationalLevel mismatch")
    if organization.get("makesOffer") != expected_organization_offers(physical):
        raise GateError(f"{path}: organization makesOffer is not source-grounded")

    for role in ("author", "publisher"):
        if article.get(role) != {"@id": ROOT_ORGANIZATION_ID}:
            raise GateError(f"{path}: Article.{role} is not the root organization")
    if article.get("educationalLevel") != expected_article_levels(source, profile):
        raise GateError(f"{path}: Article.educationalLevel order mismatch")
    if article_school_names(article, profile, source, path) != source.schools[profile.slug]:
        raise GateError(f"{path}: Article school mentions mismatch")
    if service.get("provider") != {"@id": physical.entity_id}:
        raise GateError(f"{path}: Service.provider is not the physical center")
    if service.get("areaServed") != [source.locality]:
        raise GateError(f"{path}: Service.areaServed mismatch")
    expected_offers = expected_page_offers(
        source,
        physical,
        profile,
        str(service.get("@id", "")),
    )
    if service.get("offers") != expected_offers:
        raise GateError(f"{path}: Service.offers mismatch")
    if "makesOffer" in service:
        raise GateError(f"{path}: stale Service.makesOffer remains")

    expected_names = {
        f"{profile.course} {subject} 학습 상담"
        for subject, grades in expected_level_grades(
            source.grades, profile
        ).items()
        if grades
    }
    actual_names = {
        str(offer.get("name", ""))
        for offer in service.get("offers", [])
        if isinstance(offer, dict) and not offer.get("url")
    }
    if actual_names != expected_names:
        raise GateError(
            f"{path}: Service subject offer set is not page-level/source-grounded"
        )
    for offer in service.get("offers", []):
        if not isinstance(offer, dict) or offer.get("url"):
            continue
        item = offer.get("itemOffered", {})
        item_id = str(item.get("@id", ""))
        if not item_id.startswith(str(service.get("@id", "")) + "-"):
            raise GateError(f"{path}: subject Offer leaks a foreign page @id")
        if item.get("provider") != {"@id": physical.entity_id}:
            raise GateError(f"{path}: nested subject Service provider mismatch")
        if item.get("areaServed") != [source.locality]:
            raise GateError(f"{path}: nested subject Service area mismatch")


def build_candidate(
    path: Path,
    profile: Profile,
    source: SourceRow,
    physical: PhysicalCenter,
) -> Candidate:
    original = read_utf8_exact(path)
    graph = extract_graph(original, path)
    validate_visible_source(original, path, profile, source)
    validate_faq_parity(original, graph, path)
    validate_source_schema(graph, path, profile, source, physical)
    before = route_manifest(original, graph, path, profile, source)

    updated_graph = transform_graph(
        graph, path, profile, source, physical
    )
    transformed = replace_graph(original, updated_graph, path)
    reparsed = extract_graph(transformed, path)
    validate_visible_source(transformed, path, profile, source)
    validate_faq_parity(transformed, reparsed, path)
    validate_source_schema(reparsed, path, profile, source, physical)
    validate_transformed_schema(reparsed, path, profile, source, physical)
    after = route_manifest(transformed, reparsed, path, profile, source)
    if before != after:
        raise GateError(
            f"{path}: URL/non-schema hash manifest changed: "
            f"before={before!r}, after={after!r}"
        )

    second_graph = transform_graph(
        reparsed, path, profile, source, physical
    )
    second = replace_graph(transformed, second_graph, path)
    if second != transformed:
        raise GateError(f"{path}: transform is not idempotent")

    return Candidate(
        path=path,
        relative_path=path.relative_to(ROOT).as_posix(),
        profile=profile,
        source=source,
        physical=physical,
        original=original,
        transformed=transformed,
        source_sha256=sha256_text(original),
        before_route_manifest=before,
        after_route_manifest=after,
    )


def manifest_digest(candidates: list[Candidate], *, full_source: bool) -> str:
    payload = []
    for candidate in sorted(candidates, key=lambda item: item.relative_path):
        entry: dict[str, object] = {
            "route": candidate.before_route_manifest,
        }
        if full_source:
            entry["source_sha256"] = candidate.source_sha256
        payload.append(entry)
    serialized = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return sha256_text(serialized)


def atomic_write(path: Path, text: str) -> None:
    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.kem-schema-",
        suffix=".tmp",
        dir=path.parent,
    )
    temp_path = Path(temp_name)
    try:
        with os.fdopen(descriptor, "wb") as handle:
            handle.write(text.encode("utf-8"))
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def apply_candidates(candidates: list[Candidate]) -> None:
    changed = [item for item in candidates if item.original != item.transformed]
    for candidate in candidates:
        current = read_utf8_exact(candidate.path)
        if sha256_text(current) != candidate.source_sha256:
            raise GateError(
                f"concurrent edit detected before apply: {candidate.path}"
            )

    written: list[Candidate] = []
    try:
        for candidate in changed:
            atomic_write(candidate.path, candidate.transformed)
            written.append(candidate)
    except Exception:
        for candidate in reversed(written):
            atomic_write(candidate.path, candidate.original)
        raise

    try:
        for candidate in candidates:
            current = read_utf8_exact(candidate.path)
            expected = candidate.transformed
            if current != expected:
                raise GateError(f"post-apply hash mismatch: {candidate.path}")
            graph = extract_graph(current, candidate.path)
            validate_visible_source(
                current,
                candidate.path,
                candidate.profile,
                candidate.source,
            )
            validate_faq_parity(current, graph, candidate.path)
            validate_source_schema(
                graph,
                candidate.path,
                candidate.profile,
                candidate.source,
                candidate.physical,
            )
            validate_transformed_schema(
                graph,
                candidate.path,
                candidate.profile,
                candidate.source,
                candidate.physical,
            )
            route = route_manifest(
                current,
                graph,
                candidate.path,
                candidate.profile,
                candidate.source,
            )
            if route != candidate.before_route_manifest:
                raise GateError(
                    f"post-apply route/hash manifest mismatch: {candidate.path}"
                )
            rerun_graph = transform_graph(
                graph,
                candidate.path,
                candidate.profile,
                candidate.source,
                candidate.physical,
            )
            if replace_graph(current, rerun_graph, candidate.path) != current:
                raise GateError(f"post-apply idempotency failed: {candidate.path}")
    except Exception:
        for candidate in changed:
            atomic_write(candidate.path, candidate.original)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Repair schema only for the 3x371 existing Korean/English/Math "
            "academy detail pages. Default mode is a read-only dry-run."
        )
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="write validated JSON-LD-only changes (default: dry-run)",
    )
    args = parser.parse_args()

    by_slug, physical_by_key = load_sources()
    discovered = discover_pages(by_slug)
    candidates = [
        build_candidate(
            path,
            profile,
            source,
            physical_by_key[source.center_key],
        )
        for path, profile, source in discovered
    ]
    changed = [item for item in candidates if item.original != item.transformed]
    counts = {
        profile.slug: sum(
            item.original != item.transformed
            for item in candidates
            if item.profile == profile
        )
        for profile in PROFILES
    }
    report = {
        "mode": "apply" if args.apply else "dry-run",
        "target_pages": len(candidates),
        "physical_centers": len(physical_by_key),
        "changed_pages": len(changed),
        "unchanged_pages": len(candidates) - len(changed),
        "changed_by_category": counts,
        "url_non_schema_manifest_sha256": manifest_digest(
            candidates, full_source=False
        ),
        "source_manifest_sha256": manifest_digest(
            candidates, full_source=True
        ),
        "url_canonical_visible_gate": "pass",
        "json_source_faq_gate": "pass",
        "idempotency_gate": "pass",
    }
    if args.apply:
        apply_candidates(candidates)
        report["written_pages"] = len(changed)
        report["post_apply_gate"] = "pass"
    else:
        report["written_pages"] = 0
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except GateError as error:
        print(json.dumps({"status": "fail", "error": str(error)}, ensure_ascii=False, indent=2))
        raise SystemExit(1)
