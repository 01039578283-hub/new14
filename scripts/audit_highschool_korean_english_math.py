from __future__ import annotations

import hashlib
import html
import json
import re
import csv
import sys
from collections import Counter
from pathlib import Path
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
PROFILE = sys.argv[1].lower() if len(sys.argv) > 1 else "high"
PROFILES = {
    "high": ("고등학생국영수학원", "고등학생", "고등", "타깃학교\n(고)"),
    "middle": ("중학생국영수학원", "중학생", "중등", "타깃학교\n(중)"),
    "elementary": ("초등학생국영수학원", "초등학생", "초등", "타깃학교\n(초)"),
}
if PROFILE not in PROFILES:
    raise SystemExit(f"Unknown profile: {PROFILE}")
CATEGORY_SLUG, LEVEL_NAME, COURSE_NAME, SCHOOL_COLUMN = PROFILES[PROFILE]
CATEGORY = ROOT / "과목별학원" / CATEGORY_SLUG
DETAILS = sorted(path for path in CATEGORY.glob("*/index.html") if path.parent != CATEGORY)
REQUIRED_TYPES = {"EducationalOrganization", "LocalBusiness", "WebPage", "BreadcrumbList", "Article", "Service", "FAQPage", "ItemList"}
COMMON = ROOT.parent / "참고자료" / "공통자료"


def clean(value: str) -> str:
    return html.unescape(re.sub(r"<[^>]+>", "", value)).strip()


def incorrect_particle_forms(text: str) -> list[str]:
    # 도로명(영창로·사직로)과 고유 지명(흥덕마을)은 조사 규칙으로
    # 판별할 수 없다. 전역 정규식은 정상 명칭을 오탐하고 실제 콘텐츠까지
    # 훼손했으므로, 생성 과정에서 확인된 잘못된 결합만 명시적으로 검사한다.
    known_wrong = (
        "문제집를", "학교 숙제 수행 시간를", "최근 학교 단원를",
        "반복되는 유형를", "학습성과관리은", "학원안전관리을",
        "와와학습코칭학원로", "수 있은", "묻은 질문", "찾은 자리",
        "함께 읽은 것입니다", "맞은 순서", "반달마를", "산내마를",
        "후곡마를", "흥덕마를",
    )
    return [value for value in known_wrong if value in text]


def list_values(value: str) -> list[str]:
    return list(dict.fromkeys(part.strip() for part in re.split(r"[,/\n]+", value or "") if part.strip()))


def graph_nodes(data: dict) -> list[dict]:
    return [node for node in data.get("@graph", []) if isinstance(node, dict)]


def node_type(node: dict) -> set[str]:
    value = node.get("@type", [])
    return set(value if isinstance(value, list) else [value])


def local_target(page: Path, value: str) -> Path | None:
    parsed = urlsplit(value)
    if parsed.scheme or value.startswith(("#", "tel:", "mailto:", "javascript:")):
        return None
    path = unquote(parsed.path)
    if not path:
        return None
    if path.startswith("/"):
        target = ROOT / path.lstrip("/")
    else:
        target = page.parent / path
    target = target.resolve()
    if target.is_dir():
        target = target / "index.html"
    return target


def main() -> None:
    errors: list[str] = []
    titles: list[str] = []
    metas: list[str] = []
    faq_blocks: list[str] = []
    scenario_blocks: list[str] = []
    representative_paths: list[Path] = []
    paragraph_counter: Counter[str] = Counter()
    center_schools: dict[str, set[str]] = {}
    center_supported_grades: dict[str, set[str]] = {}
    known_schools: set[str] = set()
    with (COMMON / "센터정보 정리.csv").open(encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            locality = row["근처 수업가능 동네"].strip()
            center_schools[locality] = set(list_values(row[SCHOOL_COLUMN]))
            subject_grade_sets = [
                {grade for grade in list_values(row[column]) if grade.startswith(COURSE_NAME[0])}
                for column in ("가능학년\n(국어)", "가능학년\n(영어)", "가능학년\n(수학)")
            ]
            non_empty_grade_sets = [values for values in subject_grade_sets if values]
            center_supported_grades[locality] = (
                set.intersection(*non_empty_grade_sets) if non_empty_grade_sets else set()
            )
            for column in ("타깃학교\n(초)", "타깃학교\n(중)", "타깃학교\n(고)"):
                known_schools.update(list_values(row[column]))

    if len(DETAILS) != 371:
        errors.append(f"detail count={len(DETAILS)}")

    for page in DETAILS:
        raw = page.read_text(encoding="utf-8")
        rel = page.relative_to(ROOT).as_posix()
        title_match = re.search(r"<title>(.*?)</title>", raw, re.DOTALL)
        h1_match = re.findall(r"<h1[^>]*>(.*?)</h1>", raw, re.DOTALL)
        meta_match = re.search(r'<meta name="description" content="([^"]*)">', raw)
        if not title_match or len(h1_match) != 1 or not meta_match:
            errors.append(f"{rel}: title/h1/meta")
            continue
        title = clean(title_match.group(1)).removesuffix(" | 와와학습코칭센터 영어수학 전문학원")
        h1 = clean(h1_match[0])
        meta = html.unescape(meta_match.group(1))
        titles.append(title)
        metas.append(meta)
        if title != h1:
            errors.append(f"{rel}: title != h1")
        if not 70 <= len(meta) <= 100:
            errors.append(f"{rel}: meta={len(meta)}")
        if len(re.findall(r"<h1\b", raw)) != 1:
            errors.append(f"{rel}: h1 count")
        if any(value in raw for value in (
            "학원와", "니다, 그리고", "학생 학생", "안내이 ", "층로", "습관를", "학원등원", "학원차량",
            "학원를", "점검를", "일정를", "일정와", "학원교통", "학습동기관리", "국영수 학습 안내문",
            "FAQ 구조화 데이터", "자료 자료", "상담 상담", "시기 시기", "기준 기준",
            "운영 운영", "방식를", "내용를", "과정를", "고1식 공부법",
        )):
            errors.append(f"{rel}: awkward phrase")
        if 'data-nav="subjects" aria-current="page"' not in raw:
            errors.append(f"{rel}: nav current")

        scripts = re.findall(r'<script type="application/ld\+json">(.*?)</script>', raw, re.DOTALL)
        if len(scripts) != 1:
            errors.append(f"{rel}: jsonld scripts={len(scripts)}")
            continue
        try:
            graph = json.loads(scripts[0])
        except json.JSONDecodeError as exc:
            errors.append(f"{rel}: jsonld {exc}")
            continue
        nodes = graph_nodes(graph)
        types = set().union(*(node_type(node) for node in nodes))
        missing = REQUIRED_TYPES - types
        if missing:
            errors.append(f"{rel}: missing types={sorted(missing)}")
        if {"Review", "AggregateRating"} & types:
            errors.append(f"{rel}: review schema")

        organization = next((node for node in nodes if "EducationalOrganization" in node_type(node)), {})
        visible_grades = dict(re.findall(r'<li><strong>(국어|영어|수학)</strong><span>(.*?)</span></li>', raw, re.DOTALL))
        expected_teaches = {
            f"{COURSE_NAME} {subject}" for subject, value in visible_grades.items()
            if clean(value) != f"{COURSE_NAME} 과정 미기재"
        }
        actual_teaches = set(organization.get("teaches", [])) & {f"{COURSE_NAME} 국어", f"{COURSE_NAME} 영어", f"{COURSE_NAME} 수학"}
        if expected_teaches != actual_teaches:
            errors.append(f"{rel}: teaches mismatch {sorted(expected_teaches)} / {sorted(actual_teaches)}")

        visible_faq = [(clean(q), clean(a)) for q, a in re.findall(r'<details(?: open)?>\s*<summary>(.*?)</summary><p>(.*?)</p></details>', raw, re.DOTALL)]
        faq_node = next((node for node in nodes if "FAQPage" in node_type(node)), None)
        schema_faq = [] if not faq_node else [(item.get("name", ""), item.get("acceptedAnswer", {}).get("text", "")) for item in faq_node.get("mainEntity", [])]
        if visible_faq != schema_faq or len(visible_faq) != 4:
            errors.append(f"{rel}: faq mismatch {len(visible_faq)}/{len(schema_faq)}")
        faq_blocks.append(json.dumps(visible_faq, ensure_ascii=False))

        scenarios = [clean(value) for value in re.findall(r'<article class="scenario-card">.*?<p>(.*?)</p></article>', raw, re.DOTALL)]
        if len(scenarios) != 2:
            errors.append(f"{rel}: scenarios={len(scenarios)}")
        scenario_blocks.append(json.dumps(scenarios, ensure_ascii=False))

        breadcrumb_match = re.search(r'<nav class="breadcrumbs".*?</nav>', raw, re.DOTALL)
        if not breadcrumb_match or not re.search(rf'<span>{re.escape(html.escape(h1))}</span>\s*</nav>', breadcrumb_match.group(0)):
            errors.append(f"{rel}: breadcrumb label")

        for attr, value in re.findall(r'<(?:a|img|link|script)[^>]+(href|src)="([^"]+)"', raw):
            target = local_target(page, value)
            if target is not None and not target.exists():
                errors.append(f"{rel}: missing {attr}={value}")

        rep_match = re.search(r'<img src="(\.\./\.\./\.\./assets/representative/[^"]+)"[^>]+style="display:none;">', raw)
        if not rep_match:
            errors.append(f"{rel}: hidden representative")
        else:
            rep = (page.parent / rep_match.group(1)).resolve()
            representative_paths.append(rep)
        if not re.search(r'class="local-body-image"><img[^>]+loading="lazy"[^>]+decoding="async"', raw):
            errors.append(f"{rel}: body image loading")

        section_match = re.search(r'<section class="(?:section\s+)?manuscript-wrap">(.*?)</section>\s*<section class="section blue-wash">', raw, re.DOTALL)
        if section_match:
            section_headings = [clean(value) for value in re.findall(r'<h2[^>]*>(.*?)</h2>', section_match.group(1), re.DOTALL)]
            if len(section_headings) != len(set(section_headings)):
                errors.append(f"{rel}: duplicate manuscript headings")
            manuscript_text = clean(section_match.group(1))
            particle_errors = incorrect_particle_forms(manuscript_text)
            if particle_errors:
                errors.append(f"{rel}: particle errors={particle_errors[:6]}")
            if PROFILE == "elementary" and "중등식 문제량" in manuscript_text:
                errors.append(f"{rel}: middle-school phrasing in elementary manuscript")
            locality = title.removesuffix(f" {LEVEL_NAME} 국영수학원").strip()
            if PROFILE == "elementary":
                permitted = center_supported_grades.get(locality, set())
                for match in re.finditer(r"초등\s*([1-6])\s*[~～\-–—]\s*([1-6])학년", manuscript_text):
                    start, end = int(match.group(1)), int(match.group(2))
                    claimed = {f"초{number}" for number in range(min(start, end), max(start, end) + 1)}
                    if not claimed <= permitted:
                        errors.append(f"{rel}: unsupported grade range={match.group(0)}")
                for match in re.finditer(r"초등\s*([1-6])학년", manuscript_text):
                    if f"초{match.group(1)}" not in permitted:
                        errors.append(f"{rel}: unsupported grade={match.group(0)}")
                grade_bands = {
                    "저학년": {"초1", "초2", "초3"},
                    "중학년": {"초3", "초4"},
                    "고학년": {"초4", "초5", "초6"},
                }
                for label, claimed in grade_bands.items():
                    if label in manuscript_text and not claimed <= permitted:
                        errors.append(f"{rel}: unsupported grade band={label}")
            unexpected_schools = sorted(
                school for school in known_schools - center_schools.get(locality, set())
                if len(school) >= 2 and school in manuscript_text
            )
            if unexpected_schools:
                errors.append(f"{rel}: unrelated schools={unexpected_schools[:6]}")
            for paragraph in re.findall(r'<p>(.*?)</p>', section_match.group(1), re.DOTALL):
                value = clean(paragraph)
                if len(value) >= 80:
                    paragraph_counter[value] += 1

    if len(set(titles)) != len(titles):
        errors.append("duplicate titles")
    if len(set(metas)) != len(metas):
        errors.append("duplicate metas")
    if len(set(faq_blocks)) != len(faq_blocks):
        errors.append("duplicate faq blocks")
    if len(set(scenario_blocks)) != len(scenario_blocks):
        errors.append("duplicate scenario blocks")
    if len(representative_paths) != 371 or len(set(representative_paths)) != 371:
        errors.append("representative path uniqueness")
    digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in representative_paths if path.exists()]
    if len(set(digests)) != 371:
        errors.append(f"representative content unique={len(set(digests))}")

    repeated_paragraphs = [(count, value[:100]) for value, count in paragraph_counter.items() if count > 1]
    report = {
        "detail_pages": len(DETAILS),
        "unique_titles": len(set(titles)),
        "unique_meta_descriptions": len(set(metas)),
        "meta_length_range": [min(map(len, metas), default=0), max(map(len, metas), default=0)],
        "unique_faq_sets": len(set(faq_blocks)),
        "unique_scenario_sets": len(set(scenario_blocks)),
        "unique_representative_files": len(set(digests)),
        "repeated_long_manuscript_paragraphs": len(repeated_paragraphs),
        "errors": errors[:100],
        "error_count": len(errors),
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
