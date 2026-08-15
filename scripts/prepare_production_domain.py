from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from email.utils import format_datetime
from html import escape
from pathlib import Path
from urllib.parse import quote, urlsplit, urlunsplit

try:
    from .subject_catalog import SUBJECT_CATALOG
except ImportError:
    from subject_catalog import SUBJECT_CATALOG


ROOT = Path(__file__).resolve().parents[1]
BASE_URL = "https://xn--3e0bz50b1zcyxat54c.com"
SITE_NAME = "와와학습코칭센터"
JSON_URL_KEYS = {"@id", "url", "item", "image", "contentUrl", "thumbnailUrl"}


def page_path(path: Path) -> str:
    relative = path.relative_to(ROOT)
    if relative.as_posix() == "index.html":
        return "/"
    if relative.name == "index.html":
        return "/" + quote(relative.parent.as_posix(), safe="/") + "/"
    return "/" + quote(relative.as_posix(), safe="/")


def absolute_url(path: Path) -> str:
    return absolute_site_url(page_path(path))


def absolute_site_url(path: str) -> str:
    parts = urlsplit(path)
    encoded = urlunsplit((
        "",
        "",
        quote(parts.path, safe="/%:@"),
        quote(parts.query, safe="=&%/:;+?,@"),
        quote(parts.fragment, safe="%/:?&=;+,@"),
    ))
    return BASE_URL + encoded


def absolutize_json(value, key: str | None = None):
    if isinstance(value, dict):
        return {name: absolutize_json(item, name) for name, item in value.items()}
    if isinstance(value, list):
        return [absolutize_json(item, key) for item in value]
    if isinstance(value, str) and key in JSON_URL_KEYS and value.startswith("/"):
        return absolute_site_url(value)
    return value


def replace_head_url(source: str, pattern: str, markup: str) -> str:
    if re.search(pattern, source, flags=re.IGNORECASE):
        return re.sub(pattern, markup, source, count=1, flags=re.IGNORECASE)
    return source.replace("</head>", f"  {markup}\n</head>", 1)


def update_html(path: Path) -> None:
    if path.name == "404.html":
        return
    url = absolute_url(path)
    original = path.read_text(encoding="utf-8")
    source = original
    source = replace_head_url(
        source,
        r'<link\s+rel=["\']canonical["\'][^>]*>',
        f'<link rel="canonical" href="{url}">',
    )
    source = replace_head_url(
        source,
        r'<meta\s+property=["\']og:url["\'][^>]*>',
        f'<meta property="og:url" content="{url}">',
    )

    def rewrite_json(match: re.Match[str]) -> str:
        data = json.loads(match.group(1))
        output = json.dumps(absolutize_json(data), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        return f'<script type="application/ld+json">{output}</script>'

    source = re.sub(
        r'<script type="application/ld\+json">(.*?)</script>',
        rewrite_json,
        source,
        flags=re.DOTALL,
    )
    if source != original:
        path.write_text(source, encoding="utf-8")


def meta_description(path: Path) -> str:
    source = path.read_text(encoding="utf-8")
    match = re.search(r'<meta\s+name="description"\s+content="([^"]*)"', source, flags=re.IGNORECASE)
    return match.group(1) if match else "학생별 학습 진단과 실행 기록, 오답 재학습 과정을 안내합니다."


def page_title(path: Path) -> str:
    source = path.read_text(encoding="utf-8")
    match = re.search(r"<title>(.*?)</title>", source, flags=re.DOTALL | re.IGNORECASE)
    return re.sub(r"\s+", " ", match.group(1)).strip() if match else SITE_NAME


def write_sitemap(pages: list[Path]) -> None:
    entries = []
    for path in pages:
        modified = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc).date().isoformat()
        entries.append(f"  <url><loc>{escape(absolute_url(path))}</loc><lastmod>{modified}</lastmod></url>")
    content = '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + "\n".join(entries) + "\n</urlset>\n"
    (ROOT / "sitemap.xml").write_text(content, encoding="utf-8")


def write_robots() -> None:
    content = f"User-agent: *\nAllow: /\n\nSitemap: {BASE_URL}/sitemap.xml\n"
    (ROOT / "robots.txt").write_text(content, encoding="utf-8")


def write_rss() -> None:
    candidates = [
        ROOT / "index.html",
        ROOT / "학원소개" / "index.html",
        ROOT / "학습가이드" / "index.html",
        ROOT / "과목별학원" / "index.html",
        *(ROOT / "과목별학원" / category["slug"] / "index.html" for category in SUBJECT_CATALOG),
    ]
    items = []
    for path in candidates:
        if not path.exists():
            continue
        url = absolute_url(path)
        published = format_datetime(datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc))
        items.append(
            "    <item>"
            f"<title>{escape(page_title(path))}</title>"
            f"<link>{escape(url)}</link><guid isPermaLink=\"true\">{escape(url)}</guid>"
            f"<description>{escape(meta_description(path))}</description><pubDate>{published}</pubDate>"
            "</item>"
        )
    now = format_datetime(datetime.now(timezone.utc))
    content = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom"><channel>\n'
        f"    <title>{SITE_NAME}</title><link>{BASE_URL}/</link>"
        "<description>학년별 국어·영어·수학 학습관리와 지역별 상담 정보를 안내합니다.</description>"
        f'<atom:link href="{BASE_URL}/rss.xml" rel="self" type="application/rss+xml" />'
        f"<language>ko</language><lastBuildDate>{now}</lastBuildDate>\n"
        + "\n".join(items)
        + "\n</channel></rss>\n"
    )
    (ROOT / "rss.xml").write_text(content, encoding="utf-8")


def main() -> None:
    pages = sorted(path for path in ROOT.rglob("*.html") if path.name != "404.html")
    for path in pages:
        update_html(path)
    write_sitemap(pages)
    write_robots()
    write_rss()
    print(json.dumps({"updated_html": len(pages), "base_url": BASE_URL, "sitemap_urls": len(pages)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
