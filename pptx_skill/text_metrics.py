"""Text metrics and CJK line breaking for adaptive layout (PR3).

This module provides real font measurement via Pillow, a simple CJK line
breaker and a binary-search fitter. It must not import ``python-pptx``.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class TextRun:
    text: str
    font_family: str = "Microsoft YaHei"
    size_pt: float = 16.0
    bold: bool = False
    italic: bool = False
    color: str | None = None


@dataclass
class ParagraphStyle:
    line_height: float = 1.35
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    first_line_indent_pt: float = 0.0
    bullet_indent_pt: float = 0.0
    alignment: str = "left"


@dataclass
class LineMetrics:
    text: str
    width_pt: float
    height_pt: float
    baseline_pt: float


@dataclass
class TextMetrics:
    width_pt: float
    height_pt: float
    num_lines: int
    required_font_size_pt: float
    lines: list[LineMetrics] = field(default_factory=list)
    overflow: bool = False


# ---------------------------------------------------------------------------
# Font resolver
# ---------------------------------------------------------------------------

_FONT_CACHE: dict[tuple[str, bool, bool, float], Any] = {}
_FONT_INDEX: dict[str, list[Path]] | None = None


def _index_system_fonts() -> dict[str, list[Path]]:
    global _FONT_INDEX
    if _FONT_INDEX is not None:
        return _FONT_INDEX
    candidates: dict[str, list[Path]] = {}
    dirs = []
    if os.name == "nt":
        dirs.append(Path(os.environ.get("SystemRoot", r"C:\Windows")) / "Fonts")
        dirs.append(Path.home() / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts")
    else:
        dirs.extend([Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), Path.home() / ".fonts"])
    for d in dirs:
        if not d.exists():
            continue
        for path in d.rglob("*"):
            if path.suffix.lower() not in {".ttf", ".ttc", ".otf"}:
                continue
            name = path.stem.lower()
            candidates.setdefault(name, []).append(path)
            # Also index by normalized family tokens.
            for token in re.split(r"[_\-+]", name):
                candidates.setdefault(token, []).append(path)
    _FONT_INDEX = candidates
    return candidates


def _normalize_family(family: str) -> str:
    return re.sub(r"[^a-z0-9]", "", family.lower())


def resolve_font(font_family: str, bold: bool = False, italic: bool = False) -> Path | None:
    """Return the best-matching font file path for a family/style."""
    index = _index_system_fonts()
    norm = _normalize_family(font_family)

    # Direct match by filename.
    if norm in index:
        paths = sorted(index[norm], key=lambda p: len(p.name))
        for p in paths:
            name = p.stem.lower()
            has_bold = "bold" in name
            has_italic = "italic" in name or "oblique" in name
            if has_bold == bold and has_italic == italic:
                return p
        return paths[0]

    # Fallback family aliases.
    aliases = {
        "microsoftyahei": ["msyh", "msyhbd"],
        "yahei": ["msyh", "msyhbd"],
        "simhei": ["simhei", "msyh"],
        "simsun": ["simsun", "nsimsun"],
        "hei": ["simhei", "msyh"],
        "song": ["simsun", "nsimsun"],
        "arial": ["arial"],
        "aptos": ["aptos", "calibri", "arial"],
        "georgia": ["georgia", "times"],
    }
    for alias in aliases.get(norm, [norm]):
        if alias in index:
            return index[alias][0]

    # Any font is better than none.
    if index:
        return next(iter(index.values()))[0]
    return None


def _load_font(font_family: str, size_pt: float, bold: bool, italic: bool):
    from PIL import ImageFont

    key = (font_family, bold, italic, size_pt)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    path = resolve_font(font_family, bold, italic)
    if path:
        try:
            font: Any = ImageFont.truetype(str(path), int(size_pt))
            _FONT_CACHE[key] = font
            return font
        except Exception:
            pass
    font = ImageFont.load_default()
    return font


# ---------------------------------------------------------------------------
# Line breaking
# ---------------------------------------------------------------------------

_CJK_OPEN_PUNCT = set("（「『《〈【〖〔“‘")
_CJK_CLOSE_PUNCT = set("）」』》〉】〗〕”’。，、；：？！")


def _is_cjk(char: str) -> bool:
    code = ord(char)
    return (
        0x4E00 <= code <= 0x9FFF
        or 0x3400 <= code <= 0x4DBF
        or 0x3040 <= code <= 0x309F
        or 0x30A0 <= code <= 0x30FF
        or 0xAC00 <= code <= 0xD7AF
    )


def _can_break_before(char: str) -> bool:
    return char not in _CJK_CLOSE_PUNCT and char != " "


def _can_break_after(char: str) -> bool:
    return char not in _CJK_OPEN_PUNCT


def _measure_run(run: TextRun) -> tuple[float, float]:
    """Return (width_pt, height_pt) for a run at its declared size."""
    from PIL import Image, ImageDraw

    font = _load_font(run.font_family, run.size_pt, run.bold, run.italic)
    # Pillow getlength returns pixel length at font size; 1 pt = 1 px at 72 dpi.
    draw = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    try:
        width = draw.textlength(run.text, font=font)
    except Exception:
        width = font.getlength(run.text)
    # Estimate height from ascent/descent.
    ascent, descent = font.getmetrics()
    height = float(ascent + descent)
    return float(width), float(height)


def _char_width(char: str, run: TextRun) -> float:
    from PIL import Image, ImageDraw

    font = _load_font(run.font_family, run.size_pt, run.bold, run.italic)
    draw = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    try:
        return float(draw.textlength(char, font=font))
    except Exception:
        return float(font.getlength(char))


def _break_text(text: str, width_pt: float, run: TextRun) -> list[str]:
    """Break text into lines that fit within width_pt."""
    if not text:
        return [""]

    lines: list[str] = []
    current = ""
    current_width = 0.0

    tokens: list[str] = []
    i = 0
    while i < len(text):
        if _is_cjk(text[i]):
            tokens.append(text[i])
            i += 1
        else:
            j = i
            while j < len(text) and not _is_cjk(text[j]):
                j += 1
            # Split Latin tokens on whitespace, keep whitespace as separate token.
            for word in re.split(r"( +)", text[i:j]):
                if word:
                    tokens.append(word)
            i = j

    for token in tokens:
        # Drop leading whitespace when we are at the start of a line.
        if not current and token.isspace():
            continue

        token_width = sum(_char_width(c, run) for c in token)

        # Prefer keeping a CJK closing punctuation on the previous line.
        if current and current_width + token_width <= width_pt:
            current += token
            current_width += token_width
            continue

        # Need a break before this token. Respect simple CJK line-head/line-tail rules.
        if current:
            last_char = current[-1]
            first_char = token[0]
            if (
                current_width + token_width <= width_pt * 1.05
                and (last_char in _CJK_OPEN_PUNCT or first_char in _CJK_CLOSE_PUNCT)
            ):
                # Squeeze this token onto current line to avoid prohibited break.
                current += token
                current_width += token_width
                continue
            lines.append(current)
            current = ""
            current_width = 0.0

        if token_width <= width_pt:
            current = token
            current_width = token_width
        else:
            # Token itself is too wide; break it character by character.
            for char in token:
                cw = _char_width(char, run)
                if current_width + cw > width_pt and current:
                    lines.append(current)
                    current = ""
                    current_width = 0.0
                current += char
                current_width += cw

    if current or not lines:
        lines.append(current)

    # Post-process: no CJK closing punctuation at the start of a line.
    fixed: list[str] = []
    for _idx, line in enumerate(lines):
        if line and line[0] in _CJK_CLOSE_PUNCT and fixed:
            prev = fixed[-1]
            if prev:
                moved = prev[-1]
                remainder = prev[:-1]
                if remainder:
                    fixed[-1] = remainder
                else:
                    fixed.pop()
                line = moved + line
        fixed.append(line)
    # Drop any empty lines that may have been created.
    fixed = [ln for ln in fixed if ln]
    return fixed if fixed else [""]


# ---------------------------------------------------------------------------
# Public measurer
# ---------------------------------------------------------------------------

def measure_text(
    text: str,
    width_pt: float,
    font_size_pt: float = 16.0,
    font_family: str = "Microsoft YaHei",
    paragraph_style: ParagraphStyle | None = None,
    bold: bool = False,
    italic: bool = False,
    min_font_size_pt: float = 9.0,
    locale: str = "zh-CN",
    height_pt: float | None = None,
) -> TextMetrics:
    """Measure plain text and fit it into a box, reducing size if needed."""
    style = paragraph_style or ParagraphStyle()
    run = TextRun(text=text, font_family=font_family, size_pt=font_size_pt, bold=bold, italic=italic)
    return measure_runs([run], width_pt, style, font_size_pt, min_font_size_pt, locale, height_pt=height_pt)


def measure_runs(
    runs: list[TextRun],
    width_pt: float,
    paragraph_style: ParagraphStyle,
    font_size_pt: float,
    min_font_size_pt: float = 9.0,
    locale: str = "zh-CN",
    height_pt: float | None = None,
) -> TextMetrics:
    """Measure a sequence of runs and search for a global scale that fits."""
    if not runs:
        return TextMetrics(width_pt=width_pt, height_pt=0.0, num_lines=0, required_font_size_pt=font_size_pt)

    def fits(size: float) -> tuple[bool, list[LineMetrics], float]:
        scaled_runs = [
            TextRun(
                text=r.text,
                font_family=r.font_family,
                size_pt=size,
                bold=r.bold,
                italic=r.italic,
                color=r.color,
            )
            for r in runs
        ]
        all_lines: list[LineMetrics] = []
        total_height = paragraph_style.space_before_pt
        for run in scaled_runs:
            lines = _break_text(run.text, width_pt, run)
            for line_text in lines:
                lw, lh = _measure_run(TextRun(text=line_text, font_family=run.font_family, size_pt=size, bold=run.bold, italic=run.italic))
                line_height = lh * paragraph_style.line_height
                all_lines.append(LineMetrics(text=line_text, width_pt=lw, height_pt=line_height, baseline_pt=lh))
                total_height += line_height
        total_height += paragraph_style.space_after_pt
        ok = True
        if height_pt is not None and total_height > height_pt:
            ok = False
        return ok, all_lines, total_height

    # Binary search for required size.
    lo, hi = min_font_size_pt, font_size_pt
    required = font_size_pt
    for _ in range(12):
        mid = (lo + hi) / 2
        ok, lines, height = fits(mid)
        if ok:
            required = mid
            lo = mid
        else:
            hi = mid

    ok, lines, height = fits(required)
    return TextMetrics(
        width_pt=width_pt,
        height_pt=height,
        num_lines=len(lines),
        required_font_size_pt=required,
        lines=lines,
        overflow=not ok,
    )


# ---------------------------------------------------------------------------
# Convenience class mirroring the blueprint protocol
# ---------------------------------------------------------------------------

class TextMeasurer:
    """Blueprint-compatible measurer wrapper."""

    def measure(
        self,
        runs: list[TextRun],
        width_pt: float,
        paragraph_style: ParagraphStyle,
        font_size_pt: float,
        locale: str = "zh-CN",
        min_font_size_pt: float = 9.0,
        height_pt: float | None = None,
    ) -> TextMetrics:
        return measure_runs(runs, width_pt, paragraph_style, font_size_pt, min_font_size_pt=min_font_size_pt, locale=locale, height_pt=height_pt)


def fit_text_to_height(
    text: str,
    width_pt: float,
    height_pt: float,
    font_size_pt: float = 16.0,
    font_family: str = "Microsoft YaHei",
    min_font_size_pt: float = 9.0,
    locale: str = "zh-CN",
) -> TextMetrics:
    """Find the largest font size so the text fits both width and height."""
    style = ParagraphStyle(line_height=1.35)
    TextRun(text=text, font_family=font_family, size_pt=font_size_pt)

    def fits(size: float) -> bool:
        r = TextRun(text=text, font_family=font_family, size_pt=size)
        lines = _break_text(text, width_pt, r)
        _, lh = _measure_run(r)
        total = lh * style.line_height * len(lines)
        return total <= height_pt

    lo, hi = min_font_size_pt, font_size_pt
    required = min_font_size_pt
    for _ in range(12):
        mid = (lo + hi) / 2
        if fits(mid):
            required = mid
            lo = mid
        else:
            hi = mid

    return measure_text(
        text,
        width_pt,
        font_size_pt=required,
        font_family=font_family,
        paragraph_style=style,
        min_font_size_pt=min_font_size_pt,
        locale=locale,
    )
