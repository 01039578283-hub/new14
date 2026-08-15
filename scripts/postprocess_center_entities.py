from __future__ import annotations

import hashlib
import html
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
DETAIL_GLOB = "과목별학원/*/*/index.html"
BASE_URL = "https://xn--3e0bz50b1zcyxat54c.com"
PHONE = "010-6839-8283"
MINIMUM_DETAIL_PAGES = 1113
MINIMUM_PHYSICAL_CENTERS = 188
GRAPH_RE = re.compile(
    r'(<script\s+type="application/ld\+json">)(.*?)(</script>)', re.DOTALL
)


@dataclass
class Page:
    path: Path
    text: str
    graph: dict
    org: dict
    business: dict | None
    webpage: dict
    article: dict
    service: dict
    locality: str
    region: str
    district: str
    grades: list[tuple[str, str]]


def node_has_type(node: dict, expected: str) -> bool:
    value = node.get("@type", "")
    return expected in value if isinstance(value, list) else value == expected


def first_node(graph: dict, expected: str) -> dict | None:
    return next(
        (node for node in graph.get("@graph", []) if node_has_type(node, expected)),
        None,
    )


def extract_graph(text: str, path: Path) -> dict:
    match = GRAPH_RE.search(text)
    if not match:
        raise ValueError(f"JSON-LD script missing: {path}")
    return json.loads(match.group(2))


def center_key(org: dict) -> str:
    address = org.get("address", {})
    identifier = org.get("identifier", {})
    values = [
        str(org.get("name", "")).strip(),
        str(address.get("streetAddress", "")).strip(),
        str(identifier.get("value", "")).strip(),
    ]
    if not values[0] or not values[1]:
        raise ValueError(f"Incomplete center entity: {values}")
    return "|".join(values)


def natural_unique(values: list[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))


def has_fee_offer(offers: object) -> bool:
    return isinstance(offers, list) and any(
        isinstance(offer, dict) and bool(str(offer.get("url", "")).strip())
        for offer in offers
    )


def topic_marker(value: str) -> str:
    for character in reversed(value.strip()):
        codepoint = ord(character)
        if 0xAC00 <= codepoint <= 0xD7A3:
            return "은" if (codepoint - 0xAC00) % 28 else "는"
        if character.isalnum():
            break
    return "은"


def absolute_url(value: str) -> str:
    if value.startswith("http://") or value.startswith("https://"):
        return value
    return BASE_URL + "/" + value.lstrip("/")


def parse_page(path: Path) -> Page:
    text = path.read_text(encoding="utf-8")
    graph = extract_graph(text, path)
    org = first_node(graph, "EducationalOrganization")
    webpage = first_node(graph, "WebPage")
    article = first_node(graph, "Article")
    service = first_node(graph, "Service")
    if not all((org, webpage, article, service)):
        raise ValueError(f"Required entity missing: {path}")
    business = next(
        (
            node for node in graph.get("@graph", [])
            if node is not org and node_has_type(node, "LocalBusiness")
        ),
        None,
    )
    sections = article.get("articleSection", [])
    if not isinstance(sections, list):
        sections = [sections]
    # articleSection retains the page's locality even after the center entity
    # has been consolidated to all nearby service areas.
    locality = str(sections[-1]) if sections else str((org.get("areaServed") or [""])[0])
    region = str(sections[1]) if len(sections) >= 2 else ""
    district = str(sections[2]) if len(sections) >= 3 else ""
    grade_match = re.search(r'<ul class="grade-list">(.*?)</ul>', text, re.DOTALL)
    grades: list[tuple[str, str]] = []
    if grade_match:
        grades = [
            (html.unescape(label).strip(), html.unescape(value).strip())
            for label, value in re.findall(
                r'<li><strong>(.*?)</strong><span>(.*?)</span></li>',
                grade_match.group(1),
                re.DOTALL,
            )
        ]
    return Page(
        path, text, graph, org, business, webpage, article, service,
        locality, region, district, grades,
    )


def grade_sort(value: str) -> tuple[int, str]:
    order = {
        **{f"초{n}": n for n in range(1, 7)},
        **{f"중{n}": 10 + n for n in range(1, 4)},
        **{f"고{n}": 20 + n for n in range(1, 4)},
    }
    return order.get(value, 99), value


def stable_entity(group: list[Page], key: str) -> dict:
    sample = group[0].org
    token = hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]
    stable_id = f"{BASE_URL}/#center-{token}"
    areas = natural_unique([page.locality for page in group])
    regions = Counter(page.region for page in group if page.region)
    districts = Counter(page.district for page in group if page.district)
    teaches = natural_unique([
        item
        for page in group
        for item in page.org.get("teaches", [])
    ])
    levels = sorted(
        set(
            item
            for page in group
            for item in page.org.get("educationalLevel", [])
        ),
        key=grade_sort,
    )
    offers_by_key: dict[str, dict] = {}
    for page in group:
        for offer in page.org.get("makesOffer", []):
            normalized = json.loads(json.dumps(offer, ensure_ascii=False))
            item_offered = normalized.get("itemOffered")
            # Page-local Service @ids must not leak into the shared physical
            # center entity. Keep the semantic service name and grade scope.
            if isinstance(item_offered, dict):
                item_offered.pop("@id", None)
            key_data = {
                "url": normalized.get("url", ""),
                "name": normalized.get("name", ""),
                "eligibleCustomerType": normalized.get("eligibleCustomerType", ""),
                "itemOffered": normalized.get("itemOffered", {}),
            }
            offer_key = json.dumps(key_data, ensure_ascii=False, sort_keys=True)
            offers_by_key.setdefault(offer_key, normalized)
    image = ""
    for page in sorted(group, key=lambda value: str(value.path)):
        if page.business and page.business.get("image"):
            image = absolute_url(str(page.business["image"]))
            break
        if page.org.get("image"):
            image = absolute_url(str(page.org["image"]))
            break
    address = dict(sample.get("address", {}))
    if regions:
        address["addressRegion"] = regions.most_common(1)[0][0]
    if districts:
        address["addressLocality"] = districts.most_common(1)[0][0]
    address["addressCountry"] = "KR"
    area_text = "·".join(areas)
    entity = {
        "@type": ["EducationalOrganization", "LocalBusiness"],
        "@id": stable_id,
        "name": sample["name"],
        # A physical center can serve several localities and all three course
        # categories. Pointing the shared entity at whichever high-school page
        # sorts first made middle/elementary pages cite an unrelated canonical
        # URL. Keep the entity neutral until a dedicated center URL exists.
        "url": f"{BASE_URL}/",
        "telephone": PHONE,
        "address": address,
        "areaServed": areas,
        "description": (
            f"{sample['name']}{topic_marker(str(sample['name']))} 제공된 센터 자료에서 {address.get('streetAddress', '')} 소재로 확인되며, "
            f"{area_text} 인근 학생의 국어·영어·수학 학습 상담과 과목별 가능 학년·교습비 확인 경로를 안내합니다."
        ),
        "parentOrganization": {"@id": f"{BASE_URL}/#organization"},
    }
    if image:
        entity["image"] = image
    if teaches:
        entity["teaches"] = teaches
    if sample.get("identifier"):
        entity["identifier"] = sample["identifier"]
    if levels:
        entity["educationalLevel"] = levels
    if offers_by_key:
        entity["makesOffer"] = list(offers_by_key.values())
    return entity


def verified_answer(page: Page, entity: dict) -> str:
    address = entity.get("address", {}).get("streetAddress", "")
    grade_text = ", ".join(f"{subject} {grades}" for subject, grades in page.grades)
    missing_grade_subjects = [subject for subject, grades in page.grades if grades == "상담 시 확인"]
    listed_grade_items = [(subject, grades) for subject, grades in page.grades if grades != "상담 시 확인"]
    fee = has_fee_offer(page.service.get("offers", []))
    seed = page.path.relative_to(ROOT).as_posix()
    page_title = str(page.article.get("headline") or page.webpage.get("name") or page.locality)

    def pick(label: str, choices: list[str]) -> str:
        digest = hashlib.sha256(f"{seed}|{label}".encode("utf-8")).hexdigest()
        return choices[int(digest[:10], 16) % len(choices)]

    address_text = pick("verified-address", [
        f"제공된 센터 주소는 {address}입니다.",
        f"센터 자료에서 확인한 주소는 {address}입니다.",
        f"방문 전 확인할 제공 주소는 {address}입니다.",
        f"{page.locality} 상담 페이지의 주소 자료에 기재된 위치는 {address}입니다.",
        f"제공 자료상 해당 센터 위치는 {address}입니다.",
        f"주소 자료에는 센터가 {address}에 있는 것으로 기재되어 있습니다.",
    ])
    if missing_grade_subjects:
        if listed_grade_items:
            listed_text = ", ".join(f"{subject} {grades}" for subject, grades in listed_grade_items)
            grade_sentence = (
                f"제공 자료에는 {listed_text} 범위가 기재되어 있으며, {'·'.join(missing_grade_subjects)} 가능 학년은 상담에서 확인해야 합니다."
            )
        else:
            grade_sentence = (
                f"제공 자료에 {'·'.join(missing_grade_subjects)} 가능 학년이 기재되지 않아 상담 확인이 필요합니다."
            )
    else:
        grade_sentence = pick("verified-grades", [
            f"과목별 가능 학년은 {grade_text}입니다.",
            f"제공 자료에서 확인한 가능 학년은 {grade_text}입니다.",
            f"과목별 학년 정보는 {grade_text}입니다.",
            f"센터 자료의 과목별 학년 표기는 {grade_text}입니다.",
            f"현재 페이지에서 확인할 수 있는 학년 범위는 {grade_text}입니다.",
            f"제공 자료의 과목별 학년 정보는 {grade_text}입니다.",
        ])
    fee_text = pick("verified-fee", [
        "교습비 자료는 페이지의 센터별 교습비 확인 버튼에서 볼 수 있습니다.",
        "페이지에 연결된 센터별 교습비 자료도 함께 확인할 수 있습니다.",
        "비용 자료는 센터별 교습비 확인 버튼으로 연결됩니다.",
    ]) if fee else pick("verified-no-fee", [
        "제공된 교습비 링크가 없어 비용은 센터 상담에서 확인해야 합니다.",
        "교습비 자료가 연결되지 않아 실제 비용은 센터에 직접 확인해야 합니다.",
        "비용 정보는 제공 자료에 없어 상담 과정에서 별도로 확인합니다.",
    ])
    caution = pick("verified-caution", [
        "실제 개설 과목·시간표·보강·차량·주차는 변경되거나 제공 자료에 없을 수 있으므로 등록 전에 센터에서 다시 확인하세요.",
        "시간표와 보강, 차량·주차 운영은 제공 범위 밖이거나 달라질 수 있어 방문 전에 센터에 확인해야 합니다.",
        "과목 개설과 수업 시각, 보강·통학 관련 운영은 바뀔 수 있으므로 최종 등록 전에 직접 확인하세요.",
        "제공 학년 정보와 별개로 실제 개설 반, 시간표와 보강 방식은 센터 상담에서 다시 확인하는 것이 정확합니다.",
        "차량·주차와 구체적인 수업 시간은 페이지에서 단정하지 않으며 센터의 현재 안내를 확인해야 합니다.",
        "운영 조건은 시점에 따라 달라질 수 있으므로 과목·학년·시간표를 등록 전에 한 번 더 확인하세요.",
    ])
    context = f"{page_title} 페이지의 제공 센터 사실을 기준으로 답합니다."
    return " ".join((context, address_text, grade_sentence, fee_text, caution))


def update_visible_facts(text: str, page: Page, entity: dict) -> str:
    identifier = entity.get("identifier", {})
    registration = str(identifier.get("value", "")).strip()
    if 'class="center-verified-note"' not in text:
        facts = ""
        if registration:
            facts += f'<div><dt>등록 정보</dt><dd>{html.escape(registration)}</dd></div>'
        facts += (
            f'<div><dt>대표 상담</dt><dd><a href="tel:{PHONE}">{PHONE}</a></dd></div>'
        )
        text, count = re.subn(
            r'(</dl><ul class="grade-list">)',
            facts + r'\1',
            text,
            count=1,
        )
        if count != 1:
            raise ValueError(f"Center facts insertion failed: {page.path}")

        fee_text = (
            "교습비 자료는 아래 확인 버튼으로 연결됩니다."
            if has_fee_offer(page.service.get("offers", [])) else
            "제공된 교습비 링크가 없어 비용은 센터 상담에서 확인합니다."
        )
        note = (
            '<p class="center-verified-note"><strong>제공 자료 확인 기준</strong>'
            f'<span>표기된 학년은 제공 자료 기준입니다. {html.escape(fee_text)} '
            '시간표·보강·차량·주차는 센터에서 확인합니다.</span></p>'
        )
        text, count = re.subn(
            r'(</ul>)(?=<a class="button compact"|<p class="info-note")',
            r'\1' + note,
            text,
            count=1,
        )
        if count != 1:
            raise ValueError(f"Verified note insertion failed: {page.path}")
    return text


def update_faq(text: str, page: Page, new_answer: str) -> str:
    faq = first_node(page.graph, "FAQPage")
    if not faq:
        raise ValueError(f"FAQPage missing: {page.path}")
    candidates = [
        question for question in faq.get("mainEntity", [])
        if any(
            word in str(question.get("name", ""))
            for word in ("주소", "센터 위치", "교습비", "과목 운영", "가능 학년")
        )
    ]
    questions = faq.get("mainEntity", [])
    if not questions:
        raise ValueError(f"FAQ questions missing: {page.path}")
    # The generator reserves the fourth question for practical consultation
    # details, but the wording varies by page. Prefer an explicit center/fee
    # question and otherwise upgrade that final practical question.
    question = candidates[-1] if candidates else questions[-1]
    title = str(question.get("name", ""))
    old_answer = str(question.get("acceptedAnswer", {}).get("text", ""))
    question["acceptedAnswer"]["text"] = new_answer
    pattern = re.compile(
        r'(<details[^>]*>\s*<summary>'
        + re.escape(html.escape(title))
        + r'</summary>\s*<p>)'
        + re.escape(html.escape(old_answer))
        + r'(</p>)'
    )
    text, count = pattern.subn(
        lambda match: match.group(1) + html.escape(new_answer) + match.group(2),
        text,
        count=1,
    )
    if count != 1:
        raise ValueError(f"Visible FAQ replacement failed: {page.path}")
    return text


@lru_cache(maxsize=None)
def phone_banner_bands(image_path: Path) -> tuple[tuple[float, float], ...]:
    """Return the vertical percentages of the magenta contact strips.

    Some supplied map artworks contain one strip and some contain several
    stacked center maps. We leave the source file untouched and place the
    verified site phone over every detected strip in rendered HTML.
    """
    with Image.open(image_path) as source:
        image = source.convert("RGB")
        width, height = image.size
        sample_width = min(96, width)
        sample_height = min(900, height)
        sample = image.resize((sample_width, sample_height), Image.Resampling.BILINEAR)
        pixels = list(sample.get_flattened_data())
        active: list[int] = []
        for y in range(sample_height):
            row = pixels[y * sample_width:(y + 1) * sample_width]
            ratio = sum(
                red > 175 and green < 135 and blue > 60 and red > blue + 30
                for red, green, blue in row
            ) / sample_width
            if ratio > 0.18:
                active.append(y)
    groups: list[list[int]] = []
    join_gap = max(6, sample_height // 80)
    for y in active:
        if not groups or y > groups[-1][-1] + join_gap:
            groups.append([y])
        else:
            groups[-1].append(y)
    bands = tuple(
        (
            group[0] / sample_height * 100,
            (group[-1] + 1 - group[0]) / sample_height * 100,
        )
        for group in groups
        if group[-1] - group[0] + 1 >= 12
    )
    if not bands:
        raise ValueError(f"Phone banner not detected: {image_path}")
    return bands


def update_map_contact(text: str, path: Path) -> str:
    wrapped = re.compile(
        r'<figure class="local-map-image"(?: id="(?:센터지도|center-map)")?>'
        r'<div class="map-art">(<img\s+[^>]*>).*?</div>'
        r'(<figcaption>.*?</figcaption>)</figure>',
        re.DOTALL,
    )
    bare = re.compile(
        r'<figure class="local-map-image">(<img\s+[^>]*>)'
        r'(<figcaption>.*?</figcaption>)</figure>',
        re.DOTALL,
    )
    match = wrapped.search(text) or bare.search(text)
    if not match:
        raise ValueError(f"Map figure missing: {path}")
    image_tag, caption = match.group(1), match.group(2)
    src_match = re.search(r'\bsrc="([^"]+)"', image_tag)
    if not src_match:
        raise ValueError(f"Map source missing: {path}")
    image_path = (path.parent / html.unescape(src_match.group(1))).resolve()
    bands = phone_banner_bands(image_path)
    overlays = "".join(
        (
            f'<span class="map-contact-correction" style="top:{top:.4f}%;height:{height:.4f}%" '
            f'aria-label="대표 상담 전화 {PHONE}">{PHONE}</span>'
        )
        for top, height in bands
    )
    replacement = (
        '<figure class="local-map-image" id="center-map">'
        f'<div class="map-art">{image_tag}{overlays}</div>{caption}</figure>'
    )
    return text[:match.start()] + replacement + text[match.end():]


def replace_graph(text: str, graph: dict, path: Path) -> str:
    payload = json.dumps(graph, ensure_ascii=False, separators=(",", ":"))
    text, count = GRAPH_RE.subn(
        lambda match: match.group(1) + payload + match.group(3),
        text,
        count=1,
    )
    if count != 1:
        raise ValueError(f"JSON-LD replacement failed: {path}")
    return text


def process() -> dict:
    paths = sorted(ROOT.glob(DETAIL_GLOB))
    pages = [parse_page(path) for path in paths]
    if len(pages) < MINIMUM_DETAIL_PAGES:
        raise ValueError(
            f"Expected at least {MINIMUM_DETAIL_PAGES} detail pages, found {len(pages)}"
        )
    groups: dict[str, list[Page]] = defaultdict(list)
    for page in pages:
        groups[center_key(page.org)].append(page)
    if len(groups) < MINIMUM_PHYSICAL_CENTERS:
        raise ValueError(
            f"Expected at least {MINIMUM_PHYSICAL_CENTERS} physical centers, found {len(groups)}"
        )

    entity_by_key = {
        key: stable_entity(group, key) for key, group in groups.items()
    }
    updated = 0
    faq_updated = 0
    registration_rows = 0
    for page in pages:
        key = center_key(page.org)
        entity = json.loads(json.dumps(entity_by_key[key], ensure_ascii=False))
        page_service_offers = json.loads(json.dumps(page.service.get("offers", []), ensure_ascii=False))
        nodes = page.graph["@graph"]
        old_org_id = page.org["@id"]
        old_business_id = page.business.get("@id") if page.business else None
        page.org.clear()
        page.org.update(entity)
        if page.business:
            nodes.remove(page.business)
        for node in nodes:
            if node_has_type(node, "Article"):
                node["author"] = {"@id": entity["@id"]}
                node["publisher"] = {"@id": entity["@id"]}
                mentions = node.get("mentions", [])
                if {"@id": entity["@id"]} not in mentions:
                    node["mentions"] = [{"@id": entity["@id"]}, *mentions]
            elif node_has_type(node, "Service"):
                node["provider"] = {"@id": entity["@id"]}
                if page_service_offers:
                    node["offers"] = page_service_offers
                node.pop("makesOffer", None)
            elif node_has_type(node, "WebPage"):
                about = node.get("about", [])
                if not isinstance(about, list):
                    about = [about]
                node["about"] = [{"@id": entity["@id"]}, *[
                    item for item in about if item != {"@id": old_org_id}
                ]]
            # Prevent dangling references after merging the two entity nodes.
            serialized = json.dumps(node, ensure_ascii=False)
            if old_business_id and old_business_id in serialized:
                serialized = serialized.replace(old_business_id, entity["@id"])
                node.clear()
                node.update(json.loads(serialized))

        new_answer = verified_answer(page, entity)
        text = update_visible_facts(page.text, page, entity)
        text = update_faq(text, page, new_answer)
        faq_updated += 1
        text = update_map_contact(text, page.path)
        text = replace_graph(text, page.graph, page.path)
        page.path.write_text(text, encoding="utf-8", newline="\n")
        updated += 1
        if entity.get("identifier"):
            registration_rows += 1

    return {
        "detail_pages": len(pages),
        "physical_centers": len(groups),
        "updated_pages": updated,
        "faq_updated": faq_updated,
        "registration_rows": registration_rows,
        "stable_entity_ids": len({item["@id"] for item in entity_by_key.values()}),
    }


if __name__ == "__main__":
    print(json.dumps(process(), ensure_ascii=False, indent=2))
