#!/usr/bin/env python3
"""Legacy title utilities; CLI now applies the body-based title personalizer."""

from __future__ import annotations

import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TITLE_SUFFIX = "와와학습코칭센터 영어수학 전문학원"
TITLE_RE = re.compile(r"(<title\b[^>]*>)(.*?)(</title>)", re.IGNORECASE | re.DOTALL)
META_RE = re.compile(r"<meta\b[^>]*>", re.IGNORECASE | re.DOTALL)
KEY_RE = re.compile(r"\b(?:name|property)\s*=\s*([\"'])(og:title|twitter:title)\1", re.IGNORECASE)
CONTENT_RE = re.compile(r"(\bcontent\s*=\s*)([\"'])(.*?)\2", re.IGNORECASE | re.DOTALL)


def with_suffix(value: str) -> str:
    head, separator, _ = value.rpartition("|")
    if not separator:
        return f"{value.strip()} | {TITLE_SUFFIX}"
    return f"{head.rstrip()} | {TITLE_SUFFIX}"


def update_meta(match: re.Match[str]) -> str:
    tag = match.group(0)
    if not KEY_RE.search(tag):
        return tag

    def replace_content(content_match: re.Match[str]) -> str:
        quote = content_match.group(2)
        return content_match.group(1) + quote + with_suffix(content_match.group(3)) + quote

    return CONTENT_RE.sub(replace_content, tag, count=1)


def update_document(source: str) -> str:
    source = TITLE_RE.sub(
        lambda match: match.group(1) + with_suffix(match.group(2)) + match.group(3),
        source,
        count=1,
    )
    return META_RE.sub(update_meta, source)


def title_value(source: str) -> str:
    match = TITLE_RE.search(source)
    return re.sub(r"\s+", " ", match.group(2)).strip() if match else ""


def invalid_social_titles(source: str) -> list[str]:
    invalid: list[str] = []
    for tag_match in META_RE.finditer(source):
        tag = tag_match.group(0)
        key = KEY_RE.search(tag)
        if not key:
            continue
        content = CONTENT_RE.search(tag)
        if not content or not content.group(3).endswith(f" | {TITLE_SUFFIX}"):
            invalid.append(key.group(2).lower())
    return invalid


def main() -> None:
    """Keep the old entry point from restoring the former shared suffix."""
    import subprocess
    import sys

    helper = ROOT / "tools" / "personalize_title_suffixes.py"
    args = sys.argv[1:] or ["--write"]
    raise SystemExit(subprocess.call([sys.executable, str(helper), *args], cwd=ROOT))


if __name__ == "__main__":
    main()
