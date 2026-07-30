from __future__ import annotations

import argparse
import html
import json
import mimetypes
import posixpath
import re
from pathlib import Path
from urllib.parse import quote, unquote, urljoin, urlsplit, urlunsplit

try:
    from PIL import Image
except ImportError:  # Dimensions remain optional when Pillow is unavailable.
    Image = None


ROOT = Path(__file__).resolve().parents[1]
ORIGIN = "https://xn--3e0bz50b1zcyxat54c.com"
DEFAULT_IMAGE = "/assets/generated/site14-learning-hero.webp"

HREF_RE = re.compile(r"(?i)(\bhref\s*=\s*)([\"'])(.*?)\2")
META_RE = re.compile(r"(?i)<meta\b[^>]*>")
IMG_RE = re.compile(r"(?i)<img\b[^>]*>")
ATTR_RE = re.compile(r"(?i)([\w:-]+)\s*=\s*([\"'])(.*?)\2")


def attributes(tag: str) -> dict[str, str]:
    return {match.group(1).lower(): html.unescape(match.group(3)) for match in ATTR_RE.finditer(tag)}


def page_route(path: Path, root: Path) -> str:
    relative = path.relative_to(root)
    if relative.as_posix() == "index.html":
        return "/"
    if relative.name == "index.html":
        return "/" + quote(relative.parent.as_posix(), safe="/%:@") + "/"
    return "/" + quote(relative.as_posix(), safe="/%:@")


def rewrite_index_href(page: Path, root: Path, href: str) -> str:
    value = html.unescape(href).strip()
    if not value or value.startswith(("#", "//")):
        return href
    split = urlsplit(value)
    if split.scheme or split.netloc or split.path.lower().startswith(("mailto:", "tel:", "javascript:", "data:")):
        return href
    if not split.path.lower().endswith("index.html"):
        return href

    current_route = page_route(page, root)
    base_path = current_route if current_route.endswith("/") else posixpath.dirname(current_route) + "/"
    if split.path.startswith("/"):
        resolved = posixpath.normpath(split.path)
    else:
        resolved = posixpath.normpath(posixpath.join(unquote(base_path), split.path))
    resolved = re.sub(r"(?i)index\.html$", "", resolved)
    if not resolved.startswith("/"):
        resolved = "/" + resolved
    if resolved != "/" and not resolved.endswith("/"):
        resolved += "/"
    return urlunsplit(("", "", resolved, split.query, split.fragment))


def normalize_internal_hrefs(source: str, page: Path, root: Path) -> tuple[str, int]:
    changed = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal changed
        rewritten = rewrite_index_href(page, root, match.group(3))
        if rewritten != match.group(3):
            changed += 1
        return f"{match.group(1)}{match.group(2)}{html.escape(rewritten, quote=True)}{match.group(2)}"

    return HREF_RE.sub(replace, source), changed


def meta_value(source: str, *, property_name: str | None = None, name: str | None = None) -> str:
    for match in META_RE.finditer(source):
        attrs = attributes(match.group(0))
        if property_name and attrs.get("property", "").lower() == property_name.lower():
            return attrs.get("content", "")
        if name and attrs.get("name", "").lower() == name.lower():
            return attrs.get("content", "")
    return ""


def title_value(source: str) -> str:
    value = meta_value(source, property_name="og:title")
    if value:
        return value
    match = re.search(r"(?is)<title[^>]*>(.*?)</title>", source)
    return html.unescape(re.sub(r"<[^>]+>", "", match.group(1))).strip() if match else ""


def is_indexable(source: str) -> bool:
    robots = meta_value(source, name="robots").lower()
    return "noindex" not in robots


def absolute_url(value: str, page: Path, root: Path, source: str) -> str:
    value = html.unescape(value).strip()
    if not value:
        value = DEFAULT_IMAGE
    split = urlsplit(value)
    if split.scheme in {"http", "https"}:
        return value
    if value.startswith("//"):
        return "https:" + value
    canonical = meta_value(source, property_name="og:url")
    base = canonical if canonical.startswith("http") else ORIGIN + page_route(page, root)
    return urljoin(base, value)


def select_image(source: str) -> str:
    existing = meta_value(source, property_name="og:image")
    if existing:
        return existing
    candidates: list[str] = []
    for match in IMG_RE.finditer(source):
        src = attributes(match.group(0)).get("src", "")
        if src:
            candidates.append(src)
    for src in candidates:
        if "/representative/" in src or "assets/representative/" in src:
            return src
    for src in candidates:
        if "hero" in src.lower():
            return src
    return candidates[0] if candidates else DEFAULT_IMAGE


def local_media_info(image_url: str, root: Path) -> tuple[str, int | None, int | None]:
    split = urlsplit(image_url)
    mime = mimetypes.guess_type(split.path)[0] or ""
    if split.netloc and split.netloc != urlsplit(ORIGIN).netloc:
        return mime, None, None
    candidate = root / unquote(split.path).lstrip("/")
    if not candidate.is_file() or Image is None:
        return mime, None, None
    try:
        with Image.open(candidate) as image:
            detected = Image.MIME.get(image.format, mime)
            return detected, image.width, image.height
    except OSError:
        return mime, None, None


def remove_social_meta(source: str) -> str:
    properties = {
        "og:image",
        "og:image:secure_url",
        "og:image:type",
        "og:image:width",
        "og:image:height",
        "og:image:alt",
    }
    names = {"twitter:card", "twitter:title", "twitter:description", "twitter:image", "twitter:image:alt"}

    def replace(match: re.Match[str]) -> str:
        attrs = attributes(match.group(0))
        if attrs.get("property", "").lower() in properties or attrs.get("name", "").lower() in names:
            return ""
        return match.group(0)

    return META_RE.sub(replace, source)


def add_social_meta(source: str, page: Path, root: Path) -> tuple[str, bool, bool]:
    if not is_indexable(source):
        return source, False, False
    raw_image = select_image(source)
    image_url = absolute_url(raw_image, page, root, source)
    mime, width, height = local_media_info(image_url, root)
    title = title_value(source)
    description = meta_value(source, property_name="og:description") or meta_value(source, name="description")
    alt = title.removesuffix(" | 와와학습코칭센터").strip() + " 대표 이미지"

    source = remove_social_meta(source)
    # Keep repeated generator runs byte-stable after removing and rebuilding this block.
    source = re.sub(r"\s*</head>", "\n</head>", source, count=1, flags=re.IGNORECASE)
    tags = [
        f'<meta property="og:image" content="{html.escape(image_url, quote=True)}">',
        f'<meta property="og:image:secure_url" content="{html.escape(image_url, quote=True)}">',
    ]
    if mime:
        tags.append(f'<meta property="og:image:type" content="{html.escape(mime, quote=True)}">')
    if width and height:
        tags.extend((
            f'<meta property="og:image:width" content="{width}">',
            f'<meta property="og:image:height" content="{height}">',
        ))
    twitter_card = "summary_large_image" if width and height and width / height >= 1.5 else "summary"
    tags.extend((
        f'<meta property="og:image:alt" content="{html.escape(alt, quote=True)}">',
        f'<meta name="twitter:card" content="{twitter_card}">',
        f'<meta name="twitter:title" content="{html.escape(title, quote=True)}">',
        f'<meta name="twitter:description" content="{html.escape(description, quote=True)}">',
        f'<meta name="twitter:image" content="{html.escape(image_url, quote=True)}">',
        f'<meta name="twitter:image:alt" content="{html.escape(alt, quote=True)}">',
    ))
    block = "\n  " + "\n  ".join(tags) + "\n"
    updated, count = re.subn(r"(?i)</head>", block + "</head>", source, count=1)
    if count != 1:
        raise ValueError(f"Missing </head>: {page}")
    return updated, True, bool(width and height)


def transform_html(source: str, page: Path, root: Path = ROOT) -> tuple[str, dict[str, int]]:
    source, link_changes = normalize_internal_hrefs(source, page, root)
    source, social_added, dimensions_added = add_social_meta(source, page, root)
    return source, {
        "links_rewritten": link_changes,
        "social_pages": int(social_added),
        "dimension_pages": int(dimensions_added),
    }


def audit(root: Path) -> dict[str, int]:
    pages = sorted(root.rglob("*.html"))
    indexable = 0
    index_html_hrefs = 0
    social_failures = 0
    dimension_pages = 0
    broken_internal_hrefs = 0
    missing_og_image_files = 0
    for page in pages:
        source = page.read_text(encoding="utf-8")
        for match in HREF_RE.finditer(source):
            href = html.unescape(match.group(3)).strip()
            split_href = urlsplit(href)
            index_html_hrefs += int(split_href.path.lower().endswith("index.html"))
            if not href or href.startswith(("#", "//")) or split_href.scheme or split_href.netloc:
                continue
            if split_href.path.lower().startswith(("mailto:", "tel:", "javascript:", "data:")):
                continue
            route = page_route(page, root)
            resolved = urljoin(ORIGIN + route, href)
            resolved_parts = urlsplit(resolved)
            if resolved_parts.netloc != urlsplit(ORIGIN).netloc:
                continue
            local_path = unquote(resolved_parts.path)
            candidate = root / local_path.lstrip("/")
            if local_path.endswith("/"):
                candidate = candidate / "index.html"
            if not candidate.exists():
                broken_internal_hrefs += 1
        if not is_indexable(source):
            continue
        indexable += 1
        og_images = [attributes(tag.group(0)).get("content", "") for tag in META_RE.finditer(source) if attributes(tag.group(0)).get("property", "").lower() == "og:image"]
        twitter_cards = [tag for tag in META_RE.finditer(source) if attributes(tag.group(0)).get("name", "").lower() == "twitter:card"]
        twitter_images = [attributes(tag.group(0)).get("content", "") for tag in META_RE.finditer(source) if attributes(tag.group(0)).get("name", "").lower() == "twitter:image"]
        if len(og_images) != 1 or not og_images[0].startswith("https://") or len(twitter_cards) != 1 or len(twitter_images) != 1 or twitter_images[0] != og_images[0]:
            social_failures += 1
        elif urlsplit(og_images[0]).netloc == urlsplit(ORIGIN).netloc:
            image_file = root / unquote(urlsplit(og_images[0]).path).lstrip("/")
            missing_og_image_files += int(not image_file.is_file())
        has_width = any(attributes(tag.group(0)).get("property", "").lower() == "og:image:width" for tag in META_RE.finditer(source))
        has_height = any(attributes(tag.group(0)).get("property", "").lower() == "og:image:height" for tag in META_RE.finditer(source))
        dimension_pages += int(has_width and has_height)
    return {
        "html_pages": len(pages),
        "indexable_pages": indexable,
        "remaining_index_html_hrefs": index_html_hrefs,
        "broken_internal_hrefs": broken_internal_hrefs,
        "social_meta_failures": social_failures,
        "missing_og_image_files": missing_og_image_files,
        "pages_with_image_dimensions": dimension_pages,
    }


def normalize_site(root: Path = ROOT, *, apply: bool = True) -> dict[str, object]:
    totals = {"pages_changed": 0, "links_rewritten": 0, "social_pages": 0, "dimension_pages": 0}
    for page in sorted(root.rglob("*.html")):
        original = page.read_text(encoding="utf-8")
        updated, stats = transform_html(original, page, root)
        if updated != original:
            totals["pages_changed"] += 1
            if apply:
                page.write_text(updated, encoding="utf-8", newline="")
        for key, value in stats.items():
            totals[key] += value
    return {"changes": totals, "audit": audit(root) if apply else {}}


def main() -> None:
    parser = argparse.ArgumentParser(description="Normalize internal canonical links and social image metadata.")
    parser.add_argument("--check", action="store_true", help="Report current state without writing files.")
    args = parser.parse_args()
    result = {"audit": audit(ROOT)} if args.check else normalize_site(ROOT, apply=True)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
