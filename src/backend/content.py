"""Validate content and preserve explicit lyric pages during pagination."""

import re


MARKER = re.compile(r"(?:间奏|前奏|尾声|结束|[（(]\s*(?:间奏|前奏|尾声|结束)\s*[）)])")


def is_marker(line):
    return bool(MARKER.fullmatch(line.strip()))


def split_lines(lines, size=4):
    pages = []
    closed = False
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if is_marker(line):
            if not pages:
                raise ValueError("A lyric marker must follow lyrics; its position is unknown.")
            pages[-1].append(line)
            closed = True
        else:
            if not pages or closed or len(pages[-1]) >= size:
                pages.append([])
            pages[-1].append(line)
            closed = False
    return pages


def validate_section(item, key):
    if not isinstance(item, dict):
        return f"{key} must be an object."
    if not isinstance(item.get("title"), str) or not item["title"].strip():
        return f"{key}.title must be a non-empty string."
    if "pages" in item:
        pages = item["pages"]
        if not isinstance(pages, list) or not pages:
            return f"{key}.pages must be a non-empty list."
        for index, page in enumerate(pages):
            path = f"{key}.pages[{index}]"
            if not isinstance(page, dict):
                return f"{path} must be an object."
            lines = page.get("lines")
            if not isinstance(lines, list) or not lines or not all(
                isinstance(line, str) and line.strip() and not is_marker(line)
                for line in lines
            ):
                return f"{path}.lines must contain non-empty lyric strings; put markers in footer."
            if "footer" in page and (
                not isinstance(page["footer"], str)
                or not is_marker(page["footer"])
            ):
                return f"{path}.footer must be 间奏, 前奏, 尾声 or 结束 (optional parentheses)."
        # Avoid silently discarding a second, conflicting copy of the lyrics.
        if "lines" in item and item["lines"] != [
            line for page in pages
            for line in page["lines"] + ([page["footer"]] if "footer" in page else [])
        ]:
            return f"{key}.lines conflicts with pages; omit lines when using pages."
    else:
        lines = item.get("lines")
        if not isinstance(lines, list) or not lines or not all(
            isinstance(line, str) for line in lines
        ) or not any(line.strip() for line in lines):
            return f"{key}.lines must be a non-empty list of strings."
        try:
            split_lines(lines)
        except ValueError as exc:
            return f"{key}: {exc}"
    return None


def page_lines(page):
    lines = [line.strip() for line in page["lines"]]
    if page.get("footer"):
        lines.append(page["footer"].strip())
    return lines


def section_pages(section):
    if "pages" in section:
        return [page_lines(page) for page in section["pages"]]
    return split_lines(section["lines"])
