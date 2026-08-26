from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import random
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, unquote, urlsplit
from xml.etree import ElementTree as ET

from generate_subject_pages_from_xlsx import (
    CONFIGS,
    EXPECTED_ROWS,
    ORIGIN,
    ROOT,
    SITE_NAME,
    TITLE_SUFFIX,
    TARGET_ROOT,
    absolute_route,
    load_centers,
    load_source_rows,
    relevant_grades,
    relevant_schools,
)


EXPECTED_CATEGORIES = 7
EXPECTED_DETAILS = EXPECTED_CATEGORIES * EXPECTED_ROWS
EXPECTED_HUBS = EXPECTED_CATEGORIES
EXPECTED_GRAPH_SIGNATURES = Counter(
    {
        ("EducationalOrganization", "LocalBusiness"): 1,
        ("ImageObject",): 1,
        ("WebPage",): 1,
        ("BreadcrumbList",): 1,
        ("Article",): 1,
        ("Service",): 1,
        ("FAQPage",): 1,
        ("ItemList",): 1,
    }
)
AUTHORING_RE = re.compile(
    r"(?<![0-9A-Za-z가-힣])원고(?:를|가|는|의|에서|로|와|에|도|만)?(?![0-9A-Za-z가-힣])|"
    r"참고\s*문서|원문|재작성|리라이트|생성형\s*AI|인공지능으로\s*작성|"
    r"스프레드시트|엑셀|두\s*번\s*(?:사용|썼)|복사하지",
    re.IGNORECASE,
)
OVERCLAIM_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("unconditional", re.compile(r"무조건|100\s*%|백\s*퍼센트")),
    (
        "rank_or_unique",
        re.compile(
            r"(?:(?:전국|지역|동네)\s*)?(?<!\d)1\s*위(?![0-9A-Za-z가-힣])|"
            r"(?:전국|지역|동네)?\s*(?:최고의?\s*학원|유일한\s*학원)"
        ),
    ),
    (
        "guaranteed_result",
        re.compile(
            r"(?:성적|점수|등급|합격|진학|향상|상승).{0,16}(?:보장|확정)|"
            r"(?:보장|확정).{0,16}(?:성적|점수|등급|합격|진학|향상|상승)"
        ),
    ),
    ("certain_result", re.compile(r"확실한\s*(?:성적|점수|등급|합격|향상|상승)")),
    ("instant_result", re.compile(r"단기간.{0,12}(?:성적|점수|등급|향상|상승|완성)")),
    ("perfect_result", re.compile(r"완벽(?:한|하게).{0,12}(?:성적|학습|합격|완성)")),
)
WORD_RE = re.compile(r"[0-9A-Za-z가-힣]+")
SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?。！？])\s+|[\r\n]+")
TAG_RE_TEMPLATE = r"<%s\b(?P<attrs>[^>]*)>"
ATTR_RE = re.compile(
    r"(?P<name>[^\s=/>]+)(?:\s*=\s*(?P<quote>['\"])(?P<value>.*?)(?P=quote))?",
    re.DOTALL,
)
JSONLD_RE = re.compile(
    r'<script\b[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
    re.DOTALL | re.IGNORECASE,
)
MIN_SOURCE_SENTENCE_CHARS = 42
SOURCE_SHINGLE_WORDS = 12
SIMILARITY_SHINGLE_WORDS = 5
SIMILARITY_LIMIT = 0.75
WITHIN_PAGE_SENTENCE_CHARS = 24


@dataclass
class Finding:
    code: str
    path: str
    message: str
    details: object | None = None

    def as_dict(self) -> dict[str, object]:
        result: dict[str, object] = {
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
    check_counts: Counter[str] = field(default_factory=Counter)

    def check(self, code: str, condition: bool, path: Path | str, message: str, details: object | None = None) -> bool:
        self.check_counts[code] += 1
        if condition:
            return True
        display = path.relative_to(ROOT).as_posix() if isinstance(path, Path) and path.is_relative_to(ROOT) else str(path)
        self.findings.append(Finding(code, display, message, details))
        return False

    def add(self, code: str, path: Path | str, message: str, details: object | None = None) -> None:
        self.check(code, False, path, message, details)


@dataclass
class DetailDocument:
    path: Path
    category_slug: str
    locality_slug: str
    locality: str
    title: str
    canonical: str
    authored_text: str
    masked_text: str
    paragraphs: list[str]
    sections: list[str]
    representative_path: Path | None


def normalize(value: object) -> str:
    text = unicodedata.normalize("NFKC", html.unescape(str(value or "")))
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(value: object) -> str:
    return normalize(value).casefold()


def compact_korean(value: object) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]", "", unicodedata.normalize("NFKC", str(value or "")))


def validate_common_csv(centers: list[dict[str, object]], state: AuditState) -> dict[str, object]:
    source = ROOT.parent / "참고자료" / "공통자료" / "센터정보 정리.csv"
    if not state.check("common.file", source.is_file(), source, "common center CSV is missing"):
        return {}
    with source.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    state.check("common.row_count", len(rows) == EXPECTED_ROWS, source, "common center CSV must contain 371 rows", len(rows))
    state.check(
        "common.locality_order",
        [normalize(row.get("근처 수업가능 동네", "")) for row in rows] == [str(center["locality"]) for center in centers],
        source,
        "generator center order/localities differ from the common CSV",
    )
    school_columns = ("타깃학교\n(초)", "타깃학교\n(중)", "타깃학교\n(고)")
    grade_columns = {"국어": "가능학년\n(국어)", "영어": "가능학년\n(영어)", "수학": "가능학년\n(수학)"}
    for index, (row, center) in enumerate(zip(rows, centers), 1):
        locality = str(center["locality"])
        for subject, column in grade_columns.items():
            raw_values = [normalize(value) for value in re.split(r"[,/\n]+", row.get(column, "")) if normalize(value)]
            parsed_values = list(center["grades"][subject])  # type: ignore[index]
            state.check("common.grade_parse", parsed_values == raw_values, source, f"row {index} {locality} {subject} grade parsing differs from CSV", {"raw": raw_values, "parsed": parsed_values})
        for column in school_columns:
            raw_compact = compact_korean(row.get(column, ""))
            parsed_schools = list(center["schools"][column])  # type: ignore[index]
            unsupported = [school for school in parsed_schools if compact_korean(school) not in raw_compact]
            state.check("common.school_parse", not unsupported, source, f"row {index} {locality} parsed school is absent from its CSV cell", unsupported)
    return {"path": source.relative_to(ROOT.parent).as_posix(), "rows": len(rows)}


def parse_attrs(raw: str) -> dict[str, str]:
    attrs: dict[str, str] = {}
    for match in ATTR_RE.finditer(raw):
        name = match.group("name").lower()
        if not name or name.startswith(("<", "/")):
            continue
        attrs[name] = html.unescape(match.group("value") or "")
    return attrs


def tag_attrs(raw: str, tag: str) -> list[dict[str, str]]:
    pattern = re.compile(TAG_RE_TEMPLATE % re.escape(tag), re.DOTALL | re.IGNORECASE)
    return [parse_attrs(match.group("attrs")) for match in pattern.finditer(raw)]


def strip_visible(fragment: str) -> str:
    fragment = re.sub(r"(?is)<!--.*?-->", " ", fragment)
    fragment = re.sub(r"(?is)<(?:script|style|noscript|template)\b.*?</(?:script|style|noscript|template)>", " ", fragment)
    fragment = re.sub(r"(?is)<br\s*/?>|</(?:p|li|h[1-6]|div|section|article|details|summary|figcaption|dt|dd)>", "\n", fragment)
    fragment = re.sub(r"(?is)<[^>]+>", " ", fragment)
    lines = [normalize(line) for line in fragment.splitlines()]
    return "\n".join(line for line in lines if line)


def first_fragment(raw: str, pattern: str) -> str:
    match = re.search(pattern, raw, re.DOTALL | re.IGNORECASE)
    return match.group(1) if match else ""


def single_tag_text(raw: str, tag: str) -> list[str]:
    return [normalize(strip_visible(value)) for value in re.findall(rf"<{tag}\b[^>]*>(.*?)</{tag}>", raw, re.DOTALL | re.IGNORECASE)]


def meta_values(raw: str, *, name: str | None = None, prop: str | None = None) -> list[str]:
    result: list[str] = []
    for attrs in tag_attrs(raw, "meta"):
        if name is not None and attrs.get("name", "").casefold() == name.casefold():
            result.append(attrs.get("content", ""))
        if prop is not None and attrs.get("property", "").casefold() == prop.casefold():
            result.append(attrs.get("content", ""))
    return result


def canonical_values(raw: str) -> list[str]:
    return [
        attrs.get("href", "")
        for attrs in tag_attrs(raw, "link")
        if "canonical" in attrs.get("rel", "").casefold().split()
    ]


def graph_nodes(graph: object) -> list[dict[str, object]]:
    if not isinstance(graph, dict):
        return []
    value = graph.get("@graph", [])
    return [node for node in value if isinstance(node, dict)] if isinstance(value, list) else []


def node_types(node: dict[str, object]) -> tuple[str, ...]:
    value = node.get("@type", [])
    items = value if isinstance(value, list) else [value]
    return tuple(sorted(str(item) for item in items if item))


def node_for(nodes: Iterable[dict[str, object]], expected: str) -> dict[str, object] | None:
    return next((node for node in nodes if expected in node_types(node)), None)


def recursive_schema_types(value: object) -> set[str]:
    result: set[str] = set()
    if isinstance(value, dict):
        raw_type = value.get("@type", [])
        items = raw_type if isinstance(raw_type, list) else [raw_type]
        result.update(str(item) for item in items if item)
        for child in value.values():
            result.update(recursive_schema_types(child))
    elif isinstance(value, list):
        for child in value:
            result.update(recursive_schema_types(child))
    return result


def parse_graph(raw: str, path: Path, state: AuditState) -> tuple[dict[str, object], list[dict[str, object]]]:
    scripts = JSONLD_RE.findall(raw)
    if not state.check("jsonld.script_count", len(scripts) == 1, path, f"expected one JSON-LD script, found {len(scripts)}"):
        return {}, []
    try:
        graph = json.loads(scripts[0])
    except json.JSONDecodeError as exc:
        state.add("jsonld.parse", path, "invalid JSON-LD", str(exc))
        return {}, []
    if not isinstance(graph, dict):
        state.add("jsonld.root", path, "JSON-LD root is not an object")
        return {}, []
    state.check("jsonld.context", graph.get("@context") == "https://schema.org", path, "unexpected JSON-LD @context", graph.get("@context"))
    return graph, graph_nodes(graph)


def clean_href(value: str) -> str:
    parsed = urlsplit(value)
    if parsed.scheme:
        path = quote(unquote(parsed.path), safe="/%:@")
        return parsed._replace(path=path).geturl()
    path = quote(unquote(parsed.path), safe="/%:@")
    return ORIGIN + path


def resolve_local(page: Path, value: str) -> Path | None:
    parsed = urlsplit(html.unescape(value))
    if parsed.scheme or value.startswith(("#", "tel:", "mailto:", "javascript:")):
        return None
    path = unquote(parsed.path)
    if not path:
        return None
    target = (ROOT / path.lstrip("/")) if path.startswith("/") else (page.parent / path)
    target = target.resolve()
    if target.is_dir() or path.endswith("/"):
        target /= "index.html"
    return target


def sentence_values(value: str, minimum: int = 1) -> list[str]:
    result: list[str] = []
    for part in SENTENCE_SPLIT_RE.split(value):
        current = normalize_key(part).strip(" \t\r\n\"'“”‘’()[]{}<>.,!?。！？·")
        if len(current) >= minimum:
            result.append(current)
    return result


def word_values(value: str) -> list[str]:
    return [token.casefold() for token in WORD_RE.findall(unicodedata.normalize("NFKC", value))]


def negated_claim(text: str, start: int, end: int) -> bool:
    context = text[max(0, start - 20):min(len(text), end + 28)]
    return bool(
        re.search(
            r"보장(?:하지|되지|할\s*수\s*없)|보장하는.{0,12}(?:아니|아닌)|"
            r"확정(?:하지|되지|할\s*수\s*없)|확정하는.{0,12}(?:아니|아닌)",
            context,
        )
    )


def shingle_hashes(value: str, size: int) -> set[int]:
    words = word_values(value)
    if len(words) < size:
        return set()
    return {
        int.from_bytes(hashlib.blake2b("\x1f".join(words[index:index + size]).encode("utf-8"), digest_size=8).digest(), "big")
        for index in range(len(words) - size + 1)
    }


def shingle_hits(value: str, source_hashes: set[int], size: int, limit: int = 4) -> list[str]:
    words = word_values(value)
    hits: list[str] = []
    seen: set[int] = set()
    for index in range(len(words) - size + 1):
        phrase = " ".join(words[index:index + size])
        digest = int.from_bytes(hashlib.blake2b(phrase.replace(" ", "\x1f").encode("utf-8"), digest_size=8).digest(), "big")
        if digest in source_hashes and digest not in seen:
            seen.add(digest)
            hits.append(phrase)
            if len(hits) >= limit:
                break
    return hits


def authored_parts(raw: str) -> tuple[str, list[str], list[str]]:
    summary = first_fragment(raw, r'<div class="local-summary"[^>]*>(.*?)</div>\s*<aside\b')
    cards = first_fragment(raw, r'<div class="local-answer-grid"[^>]*>(.*?)</div>\s*</div>\s*</section>')
    manuscript = first_fragment(raw, r'<article class="site-shell manuscript-article"[^>]*>(.*?)</article>\s*</section>')
    scenarios = first_fragment(raw, r'<div class="scenario-grid"[^>]*>(.*?)</div>\s*</div>\s*</section>')
    faq = first_fragment(raw, r'<div class="faq-list"[^>]*>(.*?)</div>\s*</div>\s*</section>')
    fragments = [value for value in (cards, summary, manuscript, scenarios, faq) if value]
    text = "\n".join(strip_visible(value) for value in fragments)
    section_fragments = re.findall(r'<section class="manuscript-section"[^>]*>(.*?)</section>', manuscript, re.DOTALL | re.IGNORECASE)
    sections = [normalize_key(strip_visible(value)) for value in section_fragments if normalize(strip_visible(value))]
    paragraphs: list[str] = []
    for value in section_fragments:
        paragraphs.extend(
            normalize_key(strip_visible(paragraph))
            for _, paragraph in re.findall(r"<p\b([^>]*)>(.*?)</p>", value, re.DOTALL | re.IGNORECASE)
            if normalize(strip_visible(paragraph))
        )
    for fragment in (summary, scenarios, faq):
        paragraphs.extend(
            normalize_key(strip_visible(paragraph))
            for attrs_raw, paragraph in re.findall(r"<p\b([^>]*)>(.*?)</p>", fragment, re.DOTALL | re.IGNORECASE)
            if "chapter-label" not in parse_attrs(attrs_raw).get("class", "").split()
            and normalize(strip_visible(paragraph))
        )
    return text, paragraphs, sections


def mask_document(text: str, config: object, center: dict[str, object]) -> str:
    replacements: set[str] = {
        SITE_NAME,
        str(getattr(config, "label")),
        str(getattr(config, "level")),
        str(center.get("locality", "")),
        str(center.get("slug", "")),
        str(center.get("region", "")),
        str(center.get("district", "")),
        str(center.get("center_name", "")),
        str(center.get("address", "")),
        str(center.get("registration", "")),
        str(center.get("office_name", "")),
    }
    replacements.update(relevant_schools(config, center))
    for subject in getattr(config, "subjects"):
        replacements.update(relevant_grades(config, center, subject))
    result = unicodedata.normalize("NFKC", text)
    for value in sorted((value for value in replacements if len(value) >= 2), key=len, reverse=True):
        result = result.replace(value, " 엔터티 ")
    result = re.sub(r"(?<![0-9A-Za-z가-힣])(?:초[1-6]|중[1-3]|고[1-3])(?![0-9A-Za-z가-힣])", " 학년 ", result)
    result = re.sub(r"\d+(?:[.,]\d+)*", " 수치 ", result)
    return normalize_key(result)


def expected_meta(
    raw: str,
    title: str,
    canonical: str,
    path: Path,
    state: AuditState,
    *,
    h1: str | None = None,
    og_type: str = "article",
) -> None:
    titles = single_tag_text(raw, "title")
    h1s = single_tag_text(raw, "h1")
    canonicals = canonical_values(raw)
    og_urls = meta_values(raw, prop="og:url")
    og_titles = meta_values(raw, prop="og:title")
    descriptions = meta_values(raw, name="description")
    og_descriptions = meta_values(raw, prop="og:description")
    og_types = meta_values(raw, prop="og:type")
    og_locales = meta_values(raw, prop="og:locale")
    state.check("meta.title", titles == [f"{title} | {TITLE_SUFFIX}"], path, "title is not exact", titles)
    state.check("meta.h1", h1s == [h1 or title], path, "H1 is not exact", h1s)
    state.check("meta.canonical", canonicals == [canonical], path, "canonical is not exact", canonicals)
    state.check("meta.og_url", og_urls == [canonical], path, "og:url is not exact", og_urls)
    state.check("meta.og_title", og_titles == [f"{title} | {TITLE_SUFFIX}"], path, "og:title is not exact", og_titles)
    state.check("meta.description_pair", len(descriptions) == 1 and descriptions == og_descriptions and bool(descriptions[0]), path, "description and og:description are not one exact pair", {"description": descriptions, "og:description": og_descriptions})
    state.check("meta.og_type", og_types == [og_type], path, "og:type is not exact", og_types)
    state.check("meta.og_locale", og_locales == ["ko_KR"], path, "og:locale is not exact", og_locales)


def validate_visible_breadcrumb(
    raw: str,
    path: Path,
    expected_links: list[tuple[str, str]],
    current: str,
    state: AuditState,
) -> None:
    fragment = first_fragment(raw, r'<nav class="breadcrumbs"[^>]*>(.*?)</nav>')
    anchors = [
        (normalize(strip_visible(label)), html.unescape(attrs.get("href", "")))
        for attrs_raw, label in re.findall(r"<a\b([^>]*)>(.*?)</a>", fragment, re.DOTALL | re.IGNORECASE)
        for attrs in [parse_attrs(attrs_raw)]
    ]
    spans = [normalize(strip_visible(value)) for value in re.findall(r"<span\b[^>]*>(.*?)</span>", fragment, re.DOTALL | re.IGNORECASE)]
    state.check("breadcrumb.visible", anchors == expected_links and spans == [current], path, "visible breadcrumb is not exact", {"links": anchors, "current": spans})


def validate_schema_breadcrumb(
    nodes: list[dict[str, object]],
    path: Path,
    expected: list[tuple[str, str]],
    state: AuditState,
) -> None:
    node = node_for(nodes, "BreadcrumbList")
    if not state.check("breadcrumb.schema_node", node is not None, path, "BreadcrumbList is missing"):
        return
    items = node.get("itemListElement", [])
    actual: list[tuple[int, str, str]] = []
    if isinstance(items, list):
        for item in items:
            if isinstance(item, dict):
                actual.append((int(item.get("position", 0) or 0), str(item.get("name", "")), str(item.get("item", ""))))
    wanted = [(index, name, url) for index, (name, url) in enumerate(expected, 1)]
    state.check("breadcrumb.schema", actual == wanted, path, "schema breadcrumb is not exact", actual)


def expected_image_url(relative: str) -> str:
    return ORIGIN + quote(relative, safe="/%:@")


def validate_images(
    raw: str,
    path: Path,
    title: str,
    center: dict[str, object],
    nodes: list[dict[str, object]],
    state: AuditState,
) -> Path | None:
    images = tag_attrs(raw, "img")
    representative = [attrs for attrs in images if "/assets/representative/" in attrs.get("src", "")]
    body = [attrs for attrs in images if "/assets/centers/common/" in attrs.get("src", "")]
    maps = [attrs for attrs in images if "/assets/maps/" in attrs.get("src", "")]
    state.check("image.representative_count", len(representative) == 1, path, "representative image count is not one", len(representative))
    state.check("image.body_count", len(body) == 1, path, "body image count is not one", len(body))
    state.check("image.map_count", len(maps) == 1, path, "map image count is not one", len(maps))
    rep_path: Path | None = None
    if representative:
        attrs = representative[0]
        style = re.sub(r"\s+", "", attrs.get("style", "").casefold())
        state.check("image.representative_hidden", "display:none" in style, path, "representative image is not hidden")
        state.check("image.representative_alt", title in attrs.get("alt", ""), path, "representative alt does not contain title", attrs.get("alt"))
        rep_path = resolve_local(path, attrs.get("src", ""))
        state.check("image.representative_file", rep_path is not None and rep_path.is_file(), path, "representative file is missing", str(rep_path))
    if body:
        attrs = body[0]
        expected_src = f"../../../assets/centers/common/{center['body_image']}"
        state.check("image.body_src", attrs.get("src") == expected_src, path, "body image mapping is not exact", attrs.get("src"))
        state.check(
            "image.body_attributes",
            attrs.get("loading") == "lazy" and attrs.get("decoding") == "async" and attrs.get("width") == "918" and attrs.get("height") == "16116",
            path,
            "body image attributes are incomplete",
            attrs,
        )
        target = resolve_local(path, attrs.get("src", ""))
        state.check("image.body_file", target is not None and target.is_file(), path, "body image file is missing", str(target))
    if maps:
        attrs = maps[0]
        expected_src = f"../../../assets/maps/{center['map_name']}"
        state.check("image.map_src", attrs.get("src") == expected_src, path, "map image mapping is not exact", attrs.get("src"))
        state.check("image.map_attributes", attrs.get("loading") == "lazy" and attrs.get("decoding") == "async", path, "map image lazy/async attributes are missing", attrs)
        state.check("image.map_alt", title in attrs.get("alt", "") and str(center["center_name"]) in attrs.get("alt", ""), path, "map alt is incomplete", attrs.get("alt"))
        target = resolve_local(path, attrs.get("src", ""))
        state.check("image.map_file", target is not None and target.is_file(), path, "map image file is missing", str(target))

    image_object = node_for(nodes, "ImageObject")
    article = node_for(nodes, "Article")
    if representative and body and maps:
        rep_schema = expected_image_url("/assets/representative/" + Path(representative[0]["src"]).name)
        body_schema = expected_image_url("/assets/centers/common/" + Path(body[0]["src"]).name)
        map_schema = expected_image_url("/assets/maps/" + Path(maps[0]["src"]).name)
        state.check(
            "image.schema_object",
            image_object is not None and image_object.get("url") == rep_schema and image_object.get("contentUrl") == rep_schema,
            path,
            "ImageObject does not match representative image",
        )
        state.check("image.og", meta_values(raw, prop="og:image") == [rep_schema], path, "og:image does not match representative image", meta_values(raw, prop="og:image"))
        state.check(
            "image.schema_article",
            article is not None and article.get("image") == [rep_schema, body_schema, map_schema],
            path,
            "Article image list does not match visible image mapping",
            article.get("image") if article else None,
        )
    return rep_path


def validate_faq_and_related(
    raw: str,
    path: Path,
    nodes: list[dict[str, object]],
    state: AuditState,
) -> None:
    faq_fragment = first_fragment(raw, r'<div class="faq-list"[^>]*>(.*?)</div>\s*</div>\s*</section>')
    visible_faq = [
        (normalize(strip_visible(question)), normalize(strip_visible(answer)))
        for question, answer in re.findall(
            r"<details\b[^>]*>\s*<summary\b[^>]*>(.*?)</summary>\s*<p\b[^>]*>(.*?)</p>\s*</details>",
            faq_fragment,
            re.DOTALL | re.IGNORECASE,
        )
    ]
    faq_node = node_for(nodes, "FAQPage")
    schema_faq: list[tuple[str, str]] = []
    if faq_node and isinstance(faq_node.get("mainEntity"), list):
        for item in faq_node["mainEntity"]:
            if isinstance(item, dict):
                answer = item.get("acceptedAnswer", {})
                schema_faq.append((normalize(item.get("name", "")), normalize(answer.get("text", "") if isinstance(answer, dict) else "")))
    state.check("faq.count", len(visible_faq) == 5, path, "visible FAQ count is not five", len(visible_faq))
    state.check("faq.schema_match", visible_faq == schema_faq, path, "visible FAQ and FAQPage differ", {"visible": visible_faq, "schema": schema_faq})

    related_fragment = first_fragment(raw, r'<div class="local-navigation"[^>]*>(.*?)</div>\s*</div>\s*</section>')
    visible_links = [
        (
            normalize(strip_visible(first_fragment(label, r"<strong\b[^>]*>(.*?)</strong>"))),
            clean_href(attrs.get("href", "")),
        )
        for attrs_raw, label in re.findall(r"<a\b([^>]*)>(.*?)</a>", related_fragment, re.DOTALL | re.IGNORECASE)
        for attrs in [parse_attrs(attrs_raw)]
    ]
    item_list = node_for(nodes, "ItemList")
    schema_links: list[tuple[int, str, str]] = []
    if item_list and isinstance(item_list.get("itemListElement"), list):
        for item in item_list["itemListElement"]:
            if isinstance(item, dict):
                schema_links.append((int(item.get("position", 0) or 0), normalize(item.get("name", "")), str(item.get("url", ""))))
    expected_schema = [(index, name, url) for index, (name, url) in enumerate(visible_links, 1)]
    state.check("itemlist.visible_count", bool(visible_links), path, "visible related links are missing")
    state.check("itemlist.schema_match", schema_links == expected_schema, path, "ItemList and visible related links differ", {"visible": visible_links, "schema": schema_links})
    state.check("itemlist.number", item_list is not None and item_list.get("numberOfItems") == len(visible_links), path, "ItemList numberOfItems is wrong", item_list.get("numberOfItems") if item_list else None)
    for _, href in visible_links:
        target = resolve_local(path, href.replace(ORIGIN, "", 1)) if href.startswith(ORIGIN) else None
        state.check("itemlist.target", target is not None and target.is_file(), path, "related local target is missing", href)


def validate_grades_and_schools(
    raw: str,
    path: Path,
    config: object,
    center: dict[str, object],
    nodes: list[dict[str, object]],
    authored_text: str,
    state: AuditState,
) -> None:
    grade_fragment = first_fragment(raw, r'<ul class="grade-list"[^>]*>(.*?)</ul>')
    visible_grades = [
        (normalize(strip_visible(subject)), normalize(strip_visible(value)))
        for subject, value in re.findall(r"<li\b[^>]*>\s*<strong\b[^>]*>(.*?)</strong>\s*<span\b[^>]*>(.*?)</span>\s*</li>", grade_fragment, re.DOTALL | re.IGNORECASE)
    ]
    expected_grades = [
        (subject, "·".join(relevant_grades(config, center, subject)) or "상담 시 확인")
        for subject in getattr(config, "subjects")
    ]
    state.check("whitelist.visible_grades", visible_grades == expected_grades, path, "visible grades differ from common CSV", {"actual": visible_grades, "expected": expected_grades})
    if any(value == "상담 시 확인" for _, value in expected_grades):
        state.check("whitelist.empty_grade_label", "상담 시 확인" in strip_visible(grade_fragment), path, "blank grade is not labelled 상담 시 확인")

    expected_grade_set = {grade for subject in getattr(config, "subjects") for grade in relevant_grades(config, center, subject)}
    visible_grade_tokens = set(re.findall(r"(?<![0-9A-Za-z가-힣])(초[1-6]|중[1-3]|고[1-3])(?![0-9A-Za-z가-힣])", authored_text))
    state.check("whitelist.authored_grades", visible_grade_tokens <= expected_grade_set, path, "authored text contains unsupported grade", sorted(visible_grade_tokens - expected_grade_set))

    article = node_for(nodes, "Article")
    article_levels = set(article.get("educationalLevel", [])) if article and isinstance(article.get("educationalLevel"), list) else set()
    state.check("whitelist.article_levels", article_levels == expected_grade_set, path, "Article educationalLevel differs from page grades", {"actual": sorted(article_levels), "expected": sorted(expected_grade_set)})

    allowed_all_center_grades = {
        grade
        for values in center["grades"].values()  # type: ignore[union-attr]
        for grade in values
    }
    organization = node_for(nodes, "EducationalOrganization")
    organization_levels = set(organization.get("educationalLevel", [])) if organization and isinstance(organization.get("educationalLevel"), list) else set()
    state.check("whitelist.organization_levels", organization_levels <= allowed_all_center_grades, path, "organization claims a grade absent from common CSV", sorted(organization_levels - allowed_all_center_grades))

    school_fragment = first_fragment(raw, r'<dl class="local-facts"[^>]*>(.*?)</dl>')
    visible_school_tags = [normalize(strip_visible(value)) for value in re.findall(r'<div class="local-tags"[^>]*>.*?</div>', school_fragment, re.DOTALL | re.IGNORECASE) for value in re.findall(r"<span\b[^>]*>(.*?)</span>", value, re.DOTALL | re.IGNORECASE)]
    expected_schools = relevant_schools(config, center)
    state.check("whitelist.visible_schools", visible_school_tags == expected_schools[:8], path, "visible school tags differ from common CSV", {"actual": visible_school_tags, "expected": expected_schools[:8]})
    if not expected_schools:
        state.check("whitelist.empty_school_label", "상담 시 실제 학교 자료 확인" in strip_visible(school_fragment), path, "blank school list lacks consultation label")

    article_mentions = article.get("mentions", []) if article else []
    schema_schools = {
        normalize(item.get("name", ""))
        for item in article_mentions
        if isinstance(item, dict) and "EducationalOrganization" in node_types(item)
    } if isinstance(article_mentions, list) else set()
    state.check("whitelist.schema_schools", schema_schools == set(expected_schools), path, "Article school mentions differ from common CSV", {"actual": sorted(schema_schools), "expected": sorted(expected_schools)})


def validate_detail_schema(
    nodes: list[dict[str, object]],
    path: Path,
    config: object,
    center: dict[str, object],
    title: str,
    canonical: str,
    manuscript_section_count: int,
    state: AuditState,
) -> None:
    signatures = Counter(node_types(node) for node in nodes)
    state.check("jsonld.detail_nodes", signatures == EXPECTED_GRAPH_SIGNATURES, path, "detail graph is not the exact eight-node contract", {"actual": {"+".join(key): value for key, value in signatures.items()}, "expected": {"+".join(key): value for key, value in EXPECTED_GRAPH_SIGNATURES.items()}})
    every_type = recursive_schema_types(nodes)
    forbidden = [value for value in ("Review", "AggregateRating") if value in every_type]
    state.check("jsonld.no_review", not forbidden, path, "forbidden review/rating schema present", forbidden)
    article = node_for(nodes, "Article")
    webpage = node_for(nodes, "WebPage")
    organization = node_for(nodes, "EducationalOrganization")
    service = node_for(nodes, "Service")
    state.check("jsonld.article_identity", article is not None and article.get("headline") == title and article.get("url") == canonical, path, "Article identity is wrong")
    state.check("jsonld.webpage_identity", webpage is not None and webpage.get("name") == title and webpage.get("url") == canonical, path, "WebPage identity is wrong")
    for label, node in (("Article", article), ("WebPage", webpage)):
        state.check(f"jsonld.{label.lower()}_about", node is not None and isinstance(node.get("about"), list) and bool(node.get("about")), path, f"{label}.about is missing")
        state.check(f"jsonld.{label.lower()}_mentions", node is not None and isinstance(node.get("mentions"), list) and bool(node.get("mentions")), path, f"{label}.mentions is missing")
        parts = node.get("hasPart", []) if node else []
        state.check(f"jsonld.{label.lower()}_haspart", isinstance(parts, list) and len(parts) == manuscript_section_count and manuscript_section_count > 0, path, f"{label}.hasPart does not match visible sections", len(parts) if isinstance(parts, list) else None)
    state.check("jsonld.article_section", article is not None and isinstance(article.get("articleSection"), list) and len(article.get("articleSection", [])) >= 4, path, "Article.articleSection is incomplete", article.get("articleSection") if article else None)
    offers = organization.get("makesOffer", []) if organization else []
    state.check("jsonld.makes_offer", isinstance(offers, list) and bool(offers), path, "EducationalOrganization.makesOffer is missing")
    service_offers = service.get("offers", []) if service else []
    state.check("jsonld.service_offers", isinstance(service_offers, list) and bool(service_offers), path, "Service.offers is missing")
    validate_offer_contract(
        organization,
        service,
        path,
        config,
        center,
        canonical,
        state,
    )


def validate_offer_contract(
    organization: dict[str, object] | None,
    service: dict[str, object] | None,
    path: Path,
    config: object,
    center: dict[str, object],
    canonical: str,
    state: AuditState,
) -> None:
    level = str(getattr(config, "level"))
    grade_prefix = str(getattr(config, "grade_prefix"))
    subjects = tuple(str(subject) for subject in getattr(config, "subjects"))
    expected_names = [
        f"{level} {subject} 학습 상담" if grade_prefix else f"{subject} 학습 상담"
        for subject in subjects
    ]
    service_id = canonical + "#service"
    state.check(
        "offer.service_identity",
        service is not None
        and service.get("@id") == service_id
        and service.get("serviceType") == getattr(config, "label"),
        path,
        "Service identity/category is not page-specific",
        {
            "id": service.get("@id") if service else None,
            "serviceType": service.get("serviceType") if service else None,
            "expected_id": service_id,
            "expected_serviceType": getattr(config, "label"),
        },
    )

    raw_service_offers = service.get("offers", []) if service else []
    service_offers = [offer for offer in raw_service_offers if isinstance(offer, dict)] if isinstance(raw_service_offers, list) else []
    service_subject_offers = [offer for offer in service_offers if str(offer.get("name", "")) in expected_names]
    actual_subject_names = [str(offer.get("name", "")) for offer in service_subject_offers]
    state.check(
        "offer.service_subject_scope",
        actual_subject_names == expected_names,
        path,
        "Service.offers does not exactly match this page's subject category",
        {"actual": actual_subject_names, "expected": expected_names},
    )
    # No additional subject-learning offer may be inherited from a different
    # category. The only permitted extra is the center fee-information offer.
    unexpected_learning_names = [
        str(offer.get("name", ""))
        for offer in service_offers
        if str(offer.get("name", "")).endswith("학습 상담")
        and str(offer.get("name", "")) not in expected_names
    ]
    state.check(
        "offer.service_no_cross_subject",
        not unexpected_learning_names,
        path,
        "Service.offers contains a different category's subject offer",
        unexpected_learning_names,
    )
    offer_by_name = {str(offer.get("name", "")): offer for offer in service_subject_offers}
    for subject, expected_name in zip(subjects, expected_names):
        offer = offer_by_name.get(expected_name, {})
        item = offer.get("itemOffered", {}) if isinstance(offer, dict) else {}
        grades = relevant_grades(config, center, subject)
        expected_eligible = "·".join(grades)
        actual_eligible = str(offer.get("eligibleCustomerType", "")) if isinstance(offer, dict) else ""
        state.check(
            "offer.service_item",
            node_types(offer) == ("Offer",)
            and isinstance(item, dict)
            and "Service" in node_types(item)
            and item.get("@id") == service_id
            and item.get("name") == expected_name
            and actual_eligible == expected_eligible,
            path,
            f"Service offer contract is wrong for {expected_name}",
            {
                "offer": offer,
                "expected_item_id": service_id,
                "expected_eligibleCustomerType": expected_eligible,
            },
        )

    expected_fee_name = f"{center['center_name']} 교습과정·교습비 확인"
    fee_offers = [offer for offer in service_offers if str(offer.get("name", "")) == expected_fee_name]
    if center["tuition_url"]:
        fee_offer = fee_offers[0] if len(fee_offers) == 1 else {}
        fee_item = fee_offer.get("itemOffered", {}) if isinstance(fee_offer, dict) else {}
        state.check(
            "offer.service_fee",
            len(fee_offers) == 1
            and node_types(fee_offer) == ("Offer",)
            and fee_offer.get("url") == center["tuition_url"]
            and isinstance(fee_item, dict)
            and "Service" in node_types(fee_item)
            and fee_item.get("name") == "센터별 교습과정 정보 확인",
            path,
            "Service fee-information offer is missing or incorrect",
            fee_offers,
        )
    else:
        state.check(
            "offer.service_fee_absent",
            not fee_offers,
            path,
            "Service exposes a fee offer absent from the common CSV",
            fee_offers,
        )
    expected_service_offer_count = len(expected_names) + (1 if center["tuition_url"] else 0)
    state.check(
        "offer.service_count",
        len(service_offers) == expected_service_offer_count,
        path,
        "Service.offers count is not exact",
        {"actual": len(service_offers), "expected": expected_service_offer_count},
    )

    raw_org_offers = organization.get("makesOffer", []) if organization else []
    org_offers = [offer for offer in raw_org_offers if isinstance(offer, dict)] if isinstance(raw_org_offers, list) else []
    org_names = [str(offer.get("name", "")) for offer in org_offers]
    missing_org_names = [name for name in expected_names if name not in org_names]
    state.check(
        "offer.organization_includes_page_subject",
        not missing_org_names,
        path,
        "EducationalOrganization.makesOffer omits this page's subject offer",
        {"missing": missing_org_names, "organization_offers": org_names},
    )
    matching_org_offers = [offer for offer in org_offers if str(offer.get("name", "")) in expected_names]
    invalid_org_matches: list[dict[str, object]] = []
    for offer in matching_org_offers:
        item = offer.get("itemOffered", {})
        if (
            node_types(offer) != ("Offer",)
            or not isinstance(item, dict)
            or "Service" not in node_types(item)
            or item.get("name") != offer.get("name")
            or "@id" in item
        ):
            invalid_org_matches.append(offer)
    state.check(
        "offer.organization_neutral",
        not invalid_org_matches,
        path,
        "EducationalOrganization aggregate offer is not neutral or internally consistent",
        invalid_org_matches,
    )
    missing_org_contracts: list[dict[str, str]] = []
    for subject, expected_name in zip(subjects, expected_names):
        expected_eligible = "·".join(relevant_grades(config, center, subject))
        candidates = [offer for offer in matching_org_offers if offer.get("name") == expected_name]
        if not any(str(offer.get("eligibleCustomerType", "")) == expected_eligible for offer in candidates):
            missing_org_contracts.append({
                "name": expected_name,
                "eligibleCustomerType": expected_eligible,
            })
    state.check(
        "offer.organization_includes_page_contract",
        not missing_org_contracts,
        path,
        "EducationalOrganization.makesOffer omits this page's subject/grade contract",
        {"missing": missing_org_contracts, "organization_offers": matching_org_offers},
    )


def validate_hub(
    config: object,
    centers: list[dict[str, object]],
    path: Path,
    state: AuditState,
) -> None:
    raw = path.read_text(encoding="utf-8")
    label = str(getattr(config, "label"))
    canonical = absolute_route(str(getattr(config, "slug")))
    expected_meta(raw, f"{label} 지역 안내", canonical, path, state, h1=f"동네별 {label}", og_type="website")
    # Hub H1 deliberately includes the directory prefix while its title uses
    # the concise search label.
    h1s = single_tag_text(raw, "h1")
    state.check("hub.h1", h1s == [f"동네별 {label}"], path, "hub H1 is not exact", h1s)
    validate_visible_breadcrumb(raw, path, [("홈", "/"), ("과목별학원", "/과목별학원/")], label, state)
    _, nodes = parse_graph(raw, path, state)
    signatures = Counter(node_types(node) for node in nodes)
    expected_signatures = Counter({("EducationalOrganization",): 1, ("BreadcrumbList",): 1, ("CollectionPage",): 1, ("ItemList",): 1})
    state.check("hub.schema_nodes", signatures == expected_signatures, path, "hub graph is not the exact four-node contract", {"actual": {"+".join(key): value for key, value in signatures.items()}})
    expected_breadcrumb = [
        ("홈", ORIGIN + "/"),
        ("과목별학원", ORIGIN + quote("/과목별학원/", safe="/%:@")),
        (label, canonical),
    ]
    validate_schema_breadcrumb(nodes, path, expected_breadcrumb, state)
    item_list = node_for(nodes, "ItemList")
    items = item_list.get("itemListElement", []) if item_list else []
    expected_urls = [absolute_route(str(getattr(config, "slug")), str(center["slug"])) for center in centers]
    actual_urls = [str(item.get("url", "")) for item in items if isinstance(item, dict)] if isinstance(items, list) else []
    state.check("hub.itemlist_count", item_list is not None and item_list.get("numberOfItems") == EXPECTED_ROWS and len(actual_urls) == EXPECTED_ROWS, path, "hub ItemList does not contain 371 items", {"declared": item_list.get("numberOfItems") if item_list else None, "actual": len(actual_urls)})
    state.check("hub.itemlist_scope", actual_urls == expected_urls, path, "hub ItemList order/scope differs from common CSV")
    directory_links = [attrs.get("href", "") for attrs in tag_attrs(raw, "a") if "directory-card" in attrs.get("class", "").split()]
    expected_links = [f"/과목별학원/{config.slug}/{center['slug']}/" for center in centers]
    state.check(
        "hub.visible_directory",
        len(directory_links) == EXPECTED_ROWS
        and len(set(directory_links)) == EXPECTED_ROWS
        and set(directory_links) == set(expected_links),
        path,
        "visible directory URL scope differs from common CSV",
        {"actual_count": len(directory_links), "actual_unique": len(set(directory_links)), "expected_count": len(expected_links)},
    )


def build_source_index(state: AuditState) -> tuple[set[str], set[int], dict[str, object]]:
    sentence_index: set[str] = set()
    shingle_index: set[int] = set()
    workbook_metrics: dict[str, dict[str, int]] = {}
    for config in CONFIGS:
        rows = load_source_rows(config)
        sentence_count = 0
        shingle_count_before = len(shingle_index)
        for raw in rows:
            visible = strip_visible(raw)
            sentences = sentence_values(visible, MIN_SOURCE_SENTENCE_CHARS)
            sentence_index.update(sentences)
            sentence_count += len(sentences)
            shingle_index.update(shingle_hashes(visible, SOURCE_SHINGLE_WORDS))
        workbook_metrics[str(config.workbook)] = {
            "rows": len(rows),
            "long_sentences": sentence_count,
            "new_unique_12_word_shingles": len(shingle_index) - shingle_count_before,
        }
    state.check("source.workbook_count", len(workbook_metrics) == EXPECTED_CATEGORIES, "source XLSX", "expected seven source workbooks", len(workbook_metrics))
    return sentence_index, shingle_index, {
        "workbooks": workbook_metrics,
        "unique_sentences_min_42_chars": len(sentence_index),
        "unique_12_word_shingles": len(shingle_index),
    }


def minhash_signature(values: set[int], permutations: int = 48) -> tuple[int, ...]:
    if not values:
        return tuple([0] * permutations)
    prime = (1 << 61) - 1
    coefficients = [
        (
            1 + int.from_bytes(hashlib.blake2b(f"a:{index}".encode(), digest_size=8).digest(), "big") % (prime - 1),
            int.from_bytes(hashlib.blake2b(f"b:{index}".encode(), digest_size=8).digest(), "big") % prime,
        )
        for index in range(permutations)
    ]
    result = [prime] * permutations
    for raw_value in values:
        value = raw_value % prime
        for index, (a_value, b_value) in enumerate(coefficients):
            candidate = (a_value * value + b_value) % prime
            if candidate < result[index]:
                result[index] = candidate
    return tuple(result)


def similarity_candidates(signatures: list[tuple[int, ...]], shingle_sets: list[set[int]]) -> set[tuple[int, int]]:
    candidates: set[tuple[int, int]] = set()

    def add(left: int, right: int) -> None:
        if left == right:
            return
        candidates.add((left, right) if left < right else (right, left))

    bands = 12
    rows = 4
    buckets: dict[tuple[int, tuple[int, ...]], list[int]] = defaultdict(list)
    for index, signature in enumerate(signatures):
        for band in range(bands):
            start = band * rows
            buckets[(band, signature[start:start + rows])].append(index)
    for key, members in buckets.items():
        if len(members) <= 72:
            candidates.update(combinations(members, 2))
            continue
        # Highly generic bands can contain most of the corpus. Compare local
        # neighbours plus deterministic samples instead of materialising a
        # quadratic bucket; other independent bands still contribute pairs.
        ordered = sorted(members, key=lambda item: (len(shingle_sets[item]), item))
        seed = int.from_bytes(hashlib.blake2b(repr(key).encode(), digest_size=8).digest(), "big")
        rng = random.Random(seed)
        for position, left in enumerate(ordered):
            for offset in range(1, min(13, len(ordered))):
                add(left, ordered[(position + offset) % len(ordered)])
            for _ in range(8):
                add(left, ordered[rng.randrange(len(ordered))])

    # Candidate sampling also covers same-locality category variants, adjacent
    # rows, similar document sizes and deterministic corpus-wide comparisons.
    document_count = len(signatures)
    if document_count:
        category_size = EXPECTED_ROWS
        for locality_index in range(category_size):
            members = [category * category_size + locality_index for category in range(EXPECTED_CATEGORIES)]
            candidates.update(combinations(members, 2))
        for category in range(EXPECTED_CATEGORIES):
            start = category * category_size
            for offset in range(category_size):
                current = start + offset
                for stride in (1, 2, 7, 31, 97):
                    add(current, start + ((offset + stride) % category_size))
        by_size = sorted(range(document_count), key=lambda index: (len(shingle_sets[index]), index))
        for position, left in enumerate(by_size):
            for offset in range(1, 7):
                add(left, by_size[(position + offset) % document_count])
            rng = random.Random(0x6839 ^ left)
            for _ in range(8):
                add(left, rng.randrange(document_count))
    return candidates


def validate_similarity(documents: list[DetailDocument], state: AuditState) -> dict[str, object]:
    shingle_sets = [shingle_hashes(document.masked_text, SIMILARITY_SHINGLE_WORDS) for document in documents]
    signatures = [minhash_signature(values) for values in shingle_sets]
    candidates = similarity_candidates(signatures, shingle_sets)
    maximum = 0.0
    maximum_pair: tuple[int, int] | None = None
    violations: list[dict[str, object]] = []
    for left, right in sorted(candidates):
        first = shingle_sets[left]
        second = shingle_sets[right]
        union = len(first | second)
        similarity = (len(first & second) / union) if union else 1.0
        if similarity > maximum:
            maximum = similarity
            maximum_pair = (left, right)
        if similarity >= SIMILARITY_LIMIT:
            violations.append({
                "left": documents[left].path.relative_to(ROOT).as_posix(),
                "right": documents[right].path.relative_to(ROOT).as_posix(),
                "jaccard": round(similarity, 6),
            })
    state.check(
        "duplicates.masked_5_shingle",
        maximum < SIMILARITY_LIMIT,
        "all generated detail pages",
        f"masked five-word shingle maximum must be < {SIMILARITY_LIMIT}",
        {
            "maximum": round(maximum, 6),
            "pair": [documents[index].path.relative_to(ROOT).as_posix() for index in maximum_pair] if maximum_pair else [],
            "violations": violations,
        },
    )
    return {
        "method": "48-permutation MinHash, 12x4 LSH bands, oversized-bucket and deterministic candidate sampling, exact candidate Jaccard",
        "masked_entities": True,
        "shingle_words": SIMILARITY_SHINGLE_WORDS,
        "threshold_exclusive": SIMILARITY_LIMIT,
        "documents": len(documents),
        "candidate_pairs": len(candidates),
        "maximum_jaccard": round(maximum, 6),
        "maximum_pair": [documents[index].path.relative_to(ROOT).as_posix() for index in maximum_pair] if maximum_pair else [],
        "violating_pairs": len(violations),
        "violations": violations,
    }


def validate_sitemap(expected_target_urls: set[str], state: AuditState) -> dict[str, object]:
    sitemap = ROOT / "sitemap.xml"
    if not state.check("sitemap.file", sitemap.is_file(), sitemap, "sitemap.xml is missing"):
        return {}
    try:
        tree = ET.parse(sitemap)
        locs = [normalize(element.text) for element in tree.getroot().iter() if element.tag.rsplit("}", 1)[-1] == "loc"]
    except ET.ParseError as exc:
        state.add("sitemap.parse", sitemap, "invalid sitemap XML", str(exc))
        return {}
    loc_counts = Counter(locs)
    duplicates = sorted(url for url, count in loc_counts.items() if count > 1)
    state.check("sitemap.unique", not duplicates, sitemap, "duplicate sitemap locations", duplicates)

    prefixes = [absolute_route(str(config.slug)) for config in CONFIGS]
    target_actual = {url for url in locs if any(url.startswith(prefix) for prefix in prefixes)}
    state.check("sitemap.target_scope", target_actual == expected_target_urls, sitemap, "new subject sitemap scope is not exact", {"missing": sorted(expected_target_urls - target_actual), "extra": sorted(target_actual - expected_target_urls)})

    html_canonicals: set[str] = set()
    missing_canonical_pages: list[str] = []
    for page in ROOT.rglob("index.html"):
        values = canonical_values(page.read_text(encoding="utf-8"))
        if len(values) == 1:
            html_canonicals.add(values[0])
        else:
            missing_canonical_pages.append(page.relative_to(ROOT).as_posix())
    state.check("sitemap.all_html_have_canonical", not missing_canonical_pages, sitemap, "index pages without exactly one canonical", missing_canonical_pages)
    state.check("sitemap.global_scope", set(locs) == html_canonicals, sitemap, "sitemap does not exactly equal index-page canonicals", {"missing": sorted(html_canonicals - set(locs)), "extra": sorted(set(locs) - html_canonicals)})
    return {
        "loc_count": len(locs),
        "unique_loc_count": len(set(locs)),
        "target_expected": len(expected_target_urls),
        "target_actual": len(target_actual),
        "global_index_canonicals": len(html_canonicals),
        "duplicate_locs": duplicates,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Strict audit for the seven XLSX-derived subject academy categories.")
    parser.add_argument("--report", type=Path, help="Also write the complete JSON report to this path.")
    parser.add_argument("--skip-similarity", action="store_true", help="Diagnostic only: skip the expensive masked similarity check.")
    args = parser.parse_args()
    state = AuditState()

    state.check("contract.category_config_count", len(CONFIGS) == EXPECTED_CATEGORIES, "generator CONFIGS", "generator must define exactly seven categories", len(CONFIGS))
    centers = load_centers()
    state.check("contract.center_rows", len(centers) == EXPECTED_ROWS, "common CSV", "common CSV must contain 371 centers", len(centers))
    common_metrics = validate_common_csv(centers, state)

    source_sentences, source_shingles, source_metrics = build_source_index(state)
    documents: list[DetailDocument] = []
    expected_target_urls: set[str] = set()
    representative_digests_by_category: dict[str, list[str]] = defaultdict(list)
    representative_by_locality: dict[str, list[tuple[str, str]]] = defaultdict(list)
    paragraph_pages: dict[str, set[str]] = defaultdict(set)
    section_pages: dict[str, set[str]] = defaultdict(set)
    document_pages: dict[str, set[str]] = defaultdict(set)
    source_sentence_hits: list[dict[str, object]] = []
    source_shingle_hit_pages: list[dict[str, object]] = []
    within_page_sentence_hits: list[dict[str, object]] = []
    authored_term_hits: list[dict[str, object]] = []
    overclaim_hits: list[dict[str, object]] = []

    for config in CONFIGS:
        category_slug = str(config.slug)
        category_root = TARGET_ROOT / category_slug
        hub = category_root / "index.html"
        expected_target_urls.add(absolute_route(category_slug))
        state.check("filesystem.hub", hub.is_file(), hub, "category hub is missing")
        if hub.is_file():
            validate_hub(config, centers, hub, state)

        expected_paths = {category_root / str(center["slug"]) / "index.html" for center in centers}
        actual_paths = set(category_root.glob("*/index.html")) if category_root.is_dir() else set()
        state.check(
            "filesystem.detail_scope",
            actual_paths == expected_paths,
            category_root,
            "detail filesystem scope differs from the 371 common CSV rows",
            {
                "actual": len(actual_paths),
                "expected": len(expected_paths),
                "missing": [path.relative_to(ROOT).as_posix() for path in sorted(expected_paths - actual_paths)],
                "extra": [path.relative_to(ROOT).as_posix() for path in sorted(actual_paths - expected_paths)],
            },
        )

        for center in centers:
            locality = str(center["locality"])
            locality_slug = str(center["slug"])
            path = category_root / locality_slug / "index.html"
            canonical = absolute_route(category_slug, locality_slug)
            expected_target_urls.add(canonical)
            if not path.is_file():
                continue
            raw = path.read_text(encoding="utf-8")
            title = f"{locality} {config.label}"
            expected_meta(raw, title, canonical, path, state)
            validate_visible_breadcrumb(
                raw,
                path,
                [
                    ("홈", "/"),
                    ("과목별학원", "/과목별학원/"),
                    (str(config.label), f"/과목별학원/{config.slug}/"),
                ],
                title,
                state,
            )
            _, nodes = parse_graph(raw, path, state)
            manuscript_sections = re.findall(r'<section class="manuscript-section"[^>]*>', raw, re.IGNORECASE)
            validate_detail_schema(nodes, path, config, center, title, canonical, len(manuscript_sections), state)
            validate_schema_breadcrumb(
                nodes,
                path,
                [
                    ("홈", ORIGIN + "/"),
                    ("과목별학원", ORIGIN + quote("/과목별학원/", safe="/%:@")),
                    (str(config.label), absolute_route(category_slug)),
                    (title, canonical),
                ],
                state,
            )
            validate_faq_and_related(raw, path, nodes, state)
            rep_path = validate_images(raw, path, title, center, nodes, state)
            authored_text, paragraphs, sections = authored_parts(raw)
            validate_grades_and_schools(raw, path, config, center, nodes, authored_text, state)

            rel = path.relative_to(ROOT).as_posix()
            main_fragment = first_fragment(raw, r"<main\b[^>]*>(.*?)</main>")
            main_text = strip_visible(main_fragment)
            authoring = sorted(set(match.group(0) for match in AUTHORING_RE.finditer(main_text)))
            if authoring:
                authored_term_hits.append({"path": rel, "terms": authoring})
            claims: list[dict[str, str]] = []
            for label, pattern in OVERCLAIM_PATTERNS:
                for match in pattern.finditer(main_text):
                    if label == "guaranteed_result" and negated_claim(main_text, match.start(), match.end()):
                        continue
                    claims.append({"kind": label, "text": match.group(0)})
            if claims:
                overclaim_hits.append({"path": rel, "matches": claims})

            generated_sentences = sentence_values(authored_text)
            copied_sentences = sorted({sentence for sentence in generated_sentences if len(sentence) >= MIN_SOURCE_SENTENCE_CHARS and sentence in source_sentences})
            if copied_sentences:
                source_sentence_hits.append({"path": rel, "sentences": copied_sentences})
            copied_shingles = shingle_hits(authored_text, source_shingles, SOURCE_SHINGLE_WORDS)
            if copied_shingles:
                source_shingle_hit_pages.append({"path": rel, "shingles": copied_shingles})

            sentence_counts = Counter(sentence_values(authored_text, WITHIN_PAGE_SENTENCE_CHARS))
            repeated_sentences = [{"sentence": sentence, "count": count} for sentence, count in sentence_counts.items() if count > 1]
            if repeated_sentences:
                within_page_sentence_hits.append({"path": rel, "sentences": repeated_sentences})

            for paragraph in set(paragraphs):
                paragraph_pages[paragraph].add(rel)
            for section in set(sections):
                section_pages[section].add(rel)
            document_key = normalize_key("\n".join(sections))
            if document_key:
                document_pages[document_key].add(rel)

            if rep_path and rep_path.is_file():
                digest = hashlib.sha256(rep_path.read_bytes()).hexdigest()
                representative_digests_by_category[category_slug].append(digest)
                representative_by_locality[locality_slug].append((category_slug, digest))

            documents.append(
                DetailDocument(
                    path=path,
                    category_slug=category_slug,
                    locality_slug=locality_slug,
                    locality=locality,
                    title=title,
                    canonical=canonical,
                    authored_text=authored_text,
                    masked_text=mask_document(authored_text, config, center),
                    paragraphs=paragraphs,
                    sections=sections,
                    representative_path=rep_path,
                )
            )

    state.check("contract.detail_count", len(documents) == EXPECTED_DETAILS, "new subject categories", f"expected {EXPECTED_DETAILS} detail pages", len(documents))
    state.check("contract.target_url_count", len(expected_target_urls) == EXPECTED_HUBS + EXPECTED_DETAILS, "new subject categories", "expected 2,604 target URLs", len(expected_target_urls))
    state.check("content.no_authoring_terms", not authored_term_hits, "all generated detail pages", "authoring/source-production terms leaked into visible content", authored_term_hits)
    state.check("content.no_overclaims", not overclaim_hits, "all generated detail pages", "unsupported superlative/guarantee language found", overclaim_hits)
    state.check("source.no_exact_sentence", not source_sentence_hits, "all generated detail pages", f"exact source sentence of at least {MIN_SOURCE_SENTENCE_CHARS} characters reused", source_sentence_hits)
    state.check("source.no_12_word_shingle", not source_shingle_hit_pages, "all generated detail pages", "source twelve-word shingle reused", source_shingle_hit_pages)
    state.check("duplicates.within_page_sentence", not within_page_sentence_hits, "all generated detail pages", f"within-page sentence duplicate of at least {WITHIN_PAGE_SENTENCE_CHARS} characters found", within_page_sentence_hits)
    # The source indexes can contain millions of shingles. They are no longer
    # needed once every authored document has been compared, so release them
    # before building the independent five-word similarity index.
    del source_sentences, source_shingles

    duplicate_paragraphs = [
        {"text": text, "pages": sorted(pages)}
        for text, pages in paragraph_pages.items()
        if len(pages) > 1
    ]
    duplicate_sections = [
        {"text": text, "pages": sorted(pages)}
        for text, pages in section_pages.items()
        if len(pages) > 1
    ]
    duplicate_documents = [
        {"text": text, "pages": sorted(pages)}
        for text, pages in document_pages.items()
        if len(pages) > 1
    ]
    state.check("duplicates.exact_paragraph", not duplicate_paragraphs, "all generated detail pages", "exact authored paragraph reused across pages", duplicate_paragraphs)
    state.check("duplicates.exact_section", not duplicate_sections, "all generated detail pages", "exact authored section reused across pages", duplicate_sections)
    state.check("duplicates.exact_document", not duplicate_documents, "all generated detail pages", "exact authored document reused across pages", duplicate_documents)

    representative_metrics: dict[str, object] = {}
    for config in CONFIGS:
        digests = representative_digests_by_category[str(config.slug)]
        state.check("image.category_unique", len(digests) == EXPECTED_ROWS and len(set(digests)) == EXPECTED_ROWS, TARGET_ROOT / str(config.slug), "representative files are not content-unique within category", {"files": len(digests), "unique_digests": len(set(digests))})
        representative_metrics[str(config.slug)] = {"files": len(digests), "unique_content": len(set(digests))}
    locality_collisions = [
        {"locality_slug": slug, "categories": [category for category, _ in values]}
        for slug, values in representative_by_locality.items()
        if len({digest for _, digest in values}) != len(values)
    ]
    state.check("image.same_locality_unique", not locality_collisions, "all generated detail pages", "same-locality category pages reuse representative content", locality_collisions)

    similarity_metrics: dict[str, object]
    if args.skip_similarity:
        similarity_metrics = {"skipped": True, "strict_audit": False}
        state.add("duplicates.similarity_skipped", "all generated detail pages", "masked similarity check was skipped")
    elif len(documents) == EXPECTED_DETAILS:
        similarity_metrics = validate_similarity(documents, state)
    else:
        similarity_metrics = {"skipped": True, "reason": "detail page count is incomplete"}
        state.add("duplicates.similarity_incomplete", "all generated detail pages", "masked similarity cannot run until all detail pages exist")

    sitemap_metrics = validate_sitemap(expected_target_urls, state)
    error_counts = Counter(finding.code for finding in state.findings)
    report: dict[str, object] = {
        "status": "pass" if not state.findings else "fail",
        "contract": {
            "categories": EXPECTED_CATEGORIES,
            "hubs": EXPECTED_HUBS,
            "details": EXPECTED_DETAILS,
            "target_urls": EXPECTED_HUBS + EXPECTED_DETAILS,
            "detail_jsonld_nodes": 8,
            "visible_faq_items": 5,
            "source_sentence_min_chars": MIN_SOURCE_SENTENCE_CHARS,
            "source_shingle_words": SOURCE_SHINGLE_WORDS,
            "masked_similarity_shingle_words": SIMILARITY_SHINGLE_WORDS,
            "masked_similarity_limit_exclusive": SIMILARITY_LIMIT,
        },
        "counts": {
            "generated_hubs_found": sum((TARGET_ROOT / str(config.slug) / "index.html").is_file() for config in CONFIGS),
            "generated_details_found": len(documents),
            "expected_target_urls": len(expected_target_urls),
            "authored_paragraphs": sum(len(document.paragraphs) for document in documents),
            "unique_authored_paragraphs": len(paragraph_pages),
            "authored_sections": sum(len(document.sections) for document in documents),
            "unique_authored_sections": len(section_pages),
            "unique_authored_documents": len(document_pages),
        },
        "source_reuse": {
            **source_metrics,
            "exact_sentence_hit_pages": len(source_sentence_hits),
            "twelve_word_shingle_hit_pages": len(source_shingle_hit_pages),
            "exact_sentence_hits": source_sentence_hits,
            "twelve_word_shingle_hits": source_shingle_hit_pages,
        },
        "common_csv": common_metrics,
        "duplicates": {
            "exact_paragraph_duplicates": len(duplicate_paragraphs),
            "exact_section_duplicates": len(duplicate_sections),
            "exact_document_duplicates": len(duplicate_documents),
            "within_page_sentence_duplicate_pages": len(within_page_sentence_hits),
            "masked_similarity": similarity_metrics,
        },
        "images": {
            "by_category": representative_metrics,
            "same_locality_content_collisions": len(locality_collisions),
        },
        "sitemap": sitemap_metrics,
        "check_counts": dict(sorted(state.check_counts.items())),
        "error_count": len(state.findings),
        "error_counts_by_code": dict(sorted(error_counts.items())),
        "errors": [finding.as_dict() for finding in state.findings],
    }
    output = json.dumps(report, ensure_ascii=False, indent=2)
    print(output)
    if args.report:
        destination = args.report if args.report.is_absolute() else ROOT / args.report
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(output + "\n", encoding="utf-8")
    if state.findings:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
