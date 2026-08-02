"""Import Markdown into PPTX sections.

Parses a Markdown outline into a list of ``Section`` dicts that can be
fed directly to :func:`pptx_skill.auto_generate_ppt`.  This is the
reverse of :mod:`pptx_skill.markdown_export`.

Mapping rules
-------------
- ``# H1``  -> deck title (first one) / cover slide
- ``## H2`` -> section divider slide (or content slide if it has bullets)
- ``### H3`` -> content slide with a title
- ``- bullet`` / ``* bullet`` -> bullets on the current slide
- ``1. item`` -> ordered bullets (rendered as plain bullets)
- ``> quote`` -> quote slide
- GFM tables -> table slide (``table_headers`` / ``table_rows``)
- ``` code blocks -> bullets (one bullet per line, monospace inferred)
- YAML front-matter (``---`` delimited) -> title / subtitle / theme / lang

Quick start
-----------
>>> from pptx_skill.markdown_import import markdown_to_sections
>>> from pptx_skill import auto_generate_ppt
>>> sections = markdown_to_sections("outline.md")
>>> auto_generate_ppt(title=sections[0]["title"], sections=sections[1:],
...                    output_path="deck.pptx")
"""
from __future__ import annotations

import re
from typing import Any

__all__ = [
    "markdown_to_sections",
    "from_markdown",
    "import_markdown",
]

# ---------------------------------------------------------------------------
# Front-matter parser
# ---------------------------------------------------------------------------

_FRONT_MATTER_RE = re.compile(r"\A---\s*\n(.*?)\n---\s*\n?", re.DOTALL)


def _parse_front_matter(text: str) -> tuple[dict[str, str], str]:
    """Split a simple YAML-ish front-matter block from the body.

    Only supports flat ``key: value`` lines (no nesting).  This avoids a
    PyYAML dependency.
    """
    match = _FRONT_MATTER_RE.match(text)
    if not match:
        return {}, text
    raw = match.group(1)
    body = text[match.end():]
    meta: dict[str, str] = {}
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in line:
            continue
        key, _, value = line.partition(":")
        meta[key.strip()] = value.strip().strip('"').strip("'")
    return meta, body


# ---------------------------------------------------------------------------
# GFM table parser
# ---------------------------------------------------------------------------

_TABLE_ROW_RE = re.compile(r"^\|(.+)\|\s*$")
_TABLE_SEP_RE = re.compile(r"^\|[\s:|-]+\|\s*$")


def _is_table_separator(line: str) -> bool:
    return bool(_TABLE_SEP_RE.match(line))


def _parse_table_row(line: str) -> list[str]:
    m = _TABLE_ROW_RE.match(line)
    if not m:
        return []
    return [cell.strip() for cell in m.group(1).split("|")]


# ---------------------------------------------------------------------------
# Core parser
# ---------------------------------------------------------------------------

def markdown_to_sections(
    markdown: str,
    *,
    title: str | None = None,
    subtitle: str = "",
    lang: str = "zh",
) -> list[dict[str, Any]]:
    """Parse Markdown text into a list of section dicts.

    Parameters
    ----------
    markdown : str
        Markdown source text.
    title : str, optional
        Override deck title.  If None, uses the first ``# H1``.
    subtitle : str
        Deck subtitle (default empty).
    lang : str
        Language code passed through to the generation pipeline.

    Returns
    -------
    list[dict]
        Section dicts compatible with :func:`pptx_skill.auto_generate_ppt`.
        The first element is always the cover section.
    """
    meta, body = _parse_front_matter(markdown)
    deck_title = title or meta.get("title", "")
    deck_subtitle = subtitle or meta.get("subtitle", "")
    theme_key = meta.get("theme")
    # lang may be overridden by front-matter
    if meta.get("lang"):
        lang = meta["lang"]

    sections: list[dict[str, Any]] = []
    cover_done = False

    # State for the current section being built.
    current: dict[str, Any] | None = None
    current_level = 0

    def _flush() -> None:
        nonlocal current
        if current is not None:
            sections.append(current)
        current = None

    lines = body.splitlines()
    i = 0
    n = len(lines)
    in_code = False
    code_lines: list[str] = []

    while i < n:
        line = lines[i]

        # --- Code fence handling ---
        if line.strip().startswith("```"):
            if not in_code:
                in_code = True
                code_lines = []
            else:
                in_code = False
                if current is not None and code_lines:
                    # Render code as bullets (one per non-empty line).
                    current.setdefault("bullets", [])
                    current["bullets"].extend(code_lines)
                code_lines = []
            i += 1
            continue
        if in_code:
            code_lines.append(line)
            i += 1
            continue

        stripped = line.strip()

        # --- Headings ---
        if stripped.startswith("#"):
            # Count heading level
            m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
            if m:
                level = len(m.group(1))
                heading_text = m.group(2).strip()

                if level == 1:
                    # H1: deck title / cover
                    if not cover_done:
                        if not deck_title:
                            deck_title = heading_text
                        _flush()
                        current = {
                            "title": deck_title,
                            "subtitle": deck_subtitle,
                            "layout": "cover",
                        }
                        sections.append(current)
                        current = None
                        cover_done = True
                    else:
                        # Additional H1 -> treat as section divider
                        _flush()
                        current = {
                            "title": heading_text,
                            "layout": "section",
                            "section_number": str(len(sections)),
                        }
                    i += 1
                    continue
                elif level == 2:
                    # H2: section divider or content slide
                    _flush()
                    current = {
                        "title": heading_text,
                        "layout": "",  # let choose_layout decide
                    }
                else:
                    # H3+: content slide
                    _flush()
                    current = {
                        "title": heading_text,
                        "layout": "",
                    }
                i += 1
                continue

        # --- Block quote -> quote slide ---
        if stripped.startswith(">"):
            quote_lines = []
            while i < n and lines[i].strip().startswith(">"):
                q = lines[i].strip()[1:].strip()
                if q:
                    quote_lines.append(q)
                i += 1
            _flush()
            current = {
                "title": "引用",
                "quote": " ".join(quote_lines) if len(quote_lines) == 1 else quote_lines[0] if quote_lines else "",
                "layout": "quote",
            }
            # If there are multiple lines, join as the quote text.
            if len(quote_lines) > 1:
                current["quote"] = "\n".join(quote_lines)
            continue

        # --- GFM table ---
        if _TABLE_ROW_RE.match(stripped) and i + 1 < n and _is_table_separator(lines[i + 1]):
            header = _parse_table_row(stripped)
            i += 2  # skip header + separator
            rows: list[list[str]] = []
            while i < n and _TABLE_ROW_RE.match(lines[i].strip()):
                rows.append(_parse_table_row(lines[i].strip()))
                i += 1
            _flush()
            current = {
                "title": "数据表格",
                "table_headers": header,
                "table_rows": rows,
                "layout": "table",
            }
            continue

        # --- Bullet / numbered list ---
        bullet_match = re.match(r"^[-*+]\s+(.*)$", stripped)
        numbered_match = re.match(r"^\d+\.\s+(.*)$", stripped)
        if bullet_match or numbered_match:
            text = (bullet_match or numbered_match).group(1).strip()
            if current is None:
                current = {"title": "", "layout": ""}
            current.setdefault("bullets", []).append(text)
            i += 1
            continue

        # --- Horizontal rule -> section divider ---
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
            _flush()
            current = {"title": "", "layout": "section"}
            i += 1
            continue

        # --- Regular paragraph -> if no bullets yet, treat as subtitle/kicker ---
        if stripped:
            if current is None:
                current = {"title": stripped, "layout": ""}
            elif not current.get("bullets"):
                # First paragraph after a heading becomes the subtitle.
                if not current.get("subtitle"):
                    current["subtitle"] = stripped
                else:
                    current.setdefault("bullets", []).append(stripped)
            else:
                current.setdefault("bullets", []).append(stripped)
        i += 1

    _flush()

    # If no cover was created, make one.
    if not sections:
        sections.append({
            "title": deck_title or "Untitled",
            "subtitle": deck_subtitle,
            "layout": "cover",
        })
    elif sections[0].get("layout") != "cover":
        sections.insert(0, {
            "title": deck_title or sections[0].get("title", "Untitled"),
            "subtitle": deck_subtitle,
            "layout": "cover",
        })

    # Ensure an end slide.
    if not sections or sections[-1].get("layout") != "end":
        sections.append({"title": "谢谢", "subtitle": "", "layout": "end"})

    # Stash theme/lang on the cover section for the caller to read.
    sections[0].setdefault("layout_opts", {})
    if theme_key:
        sections[0]["layout_opts"]["theme_key"] = theme_key
    sections[0]["layout_opts"]["lang"] = lang

    return sections


# ---------------------------------------------------------------------------
# File-based convenience wrappers
# ---------------------------------------------------------------------------

def from_markdown(
    path: str,
    *,
    title: str | None = None,
    subtitle: str = "",
    lang: str = "zh",
) -> list[dict[str, Any]]:
    """Read a Markdown file and return section dicts.

    Parameters
    ----------
    path : str
        Path to a ``.md`` file.
    title, subtitle, lang
        See :func:`markdown_to_sections`.

    Returns
    -------
    list[dict]
        Section dicts ready for :func:`pptx_skill.auto_generate_ppt`.
    """
    with open(path, encoding="utf-8") as f:
        text = f.read()
    return markdown_to_sections(text, title=title, subtitle=subtitle, lang=lang)


def import_markdown(
    path: str,
    output_path: str,
    *,
    title: str | None = None,
    subtitle: str = "",
    lang: str = "zh",
    theme_key: str | None = None,
    auto_search_images: bool = False,
) -> str:
    """Import a Markdown file and generate a PPTX in one step.

    Parameters
    ----------
    path : str
        Path to a ``.md`` file.
    output_path : str
        Where to save the generated ``.pptx``.
    title, subtitle, lang
        See :func:`markdown_to_sections`.
    theme_key : str, optional
        Theme to use (overrides front-matter).
    auto_search_images : bool
        Whether to auto-search images (default False).

    Returns
    -------
    str
        The output path of the generated PPTX.
    """
    sections = from_markdown(path, title=title, subtitle=subtitle, lang=lang)
    cover = sections[0]
    deck_title = cover.get("title", "Untitled")
    deck_subtitle = cover.get("subtitle", "")
    opts = cover.get("layout_opts", {})
    effective_theme = theme_key or opts.get("theme_key")
    effective_lang = lang or opts.get("lang", "zh")

    # Lazy import to avoid circular dependency at module load.
    from pptx_skill import auto_generate_ppt

    return auto_generate_ppt(
        title=deck_title,
        subtitle=deck_subtitle,
        sections=sections[1:],  # cover is handled by auto_generate_ppt
        output_path=output_path,
        theme_key=effective_theme,
        lang=effective_lang,
        auto_search_images=auto_search_images,
    )
