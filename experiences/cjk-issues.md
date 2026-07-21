# CJK Issues

## TL;DR

CJK (Chinese, Japanese, Korean) text has fundamentally different metrics and line-breaking rules from Latin text. Never assume character widths, always use Pillow for font metrics, use proper continuation markers, and handle encoding carefully on Windows.

---

## Rule 1: Font metrics must use Pillow for CJK characters

CJK characters are typically fullwidth (roughly 2x the width of a Latin character at the same font size). Hardcoded width tables only cover Latin glyphs and will severely underestimate CJK text width.

**Bad:**
```python
# Assumes all characters are the same width
text_width = len(text) * font_size * 0.6  # wrong for CJK
```

**Good:**
```python
from PIL import ImageFont

def measure_text_width(text: str, font_path: str, font_size: int) -> float:
    font = ImageFont.truetype(font_path, font_size)
    bbox = font.getbbox(text)
    return bbox[2] - bbox[0]  # right - left
```

The `text_metrics.py` module provides `measure_text_width()` and `bisect_font_size()` that handle CJK correctly using Pillow.

---

## Rule 2: Continuation markers use fullwidth parentheses

When a slide's content overflows and continues on the next slide, the continuation marker must use fullwidth CJK parentheses, not ASCII ellipsis or halfwidth parentheses.

**Bad:**
```python
# ASCII-style markers look wrong in Chinese context
title = f"{base_title}..."          # English ellipsis
title = f"{base_title}(...)"        # ASCII parentheses
title = f"{base_title}(cont.)"      # English abbreviation
```

**Good:**
```python
# Fullwidth CJK continuation marker
title = f"{base_title}（续）"       # Chinese: "continued"
title = f"{base_title}（续{page}）"  # With page number
```

The `（` and `）` are U+FF08 and U+FF09 (fullwidth left/right parentheses). The `续` is U+7EED.

---

## Rule 3: Microsoft YaHei is the safest CJK font on Windows

Font availability varies by OS. On Windows, the safest choices are:

| Priority | Font | Notes |
|---|---|---|
| 1 | Microsoft YaHei (微软雅黑) | Pre-installed on Windows Vista+, best CJK coverage |
| 2 | SimHei (黑体) | Pre-installed on all Windows versions, sans-serif |
| 3 | SimSun (宋体) | Pre-installed, serif, but poor at small sizes on screen |
| 4 | NSimSun | Same as SimSun but with fixed-width Latin |

**Never assume these fonts exist on macOS/Linux:**
- macOS: PingFang SC, Hiragino Sans GB
- Linux: Noto Sans CJK SC, WenQuanYi Micro Hei

The `THEMES` dict in `pptx_helper.py` pairs Latin and CJK fonts:
```python
FONTS = {
    "default": {"latin": "Calibri", "cjk": "Microsoft YaHei"},
    "serif":   {"latin": "Georgia", "cjk": "SimSun"},
    ...
}
```

Always specify both `latin` and `cjk` font names when creating text runs.

---

## Rule 4: CJK line-breaking does not require spaces

Latin text breaks on whitespace. CJK characters can break between any two characters without requiring a space. This means:

1. **Line-breaking algorithm must not require spaces** for CJK text
2. **CJK and Latin mixed text** needs special handling — break is allowed between CJK chars, but not mid-Latin-word

**Bad:**
```python
# Splits only on whitespace — CJK text never breaks
lines = text.split(" ")
```

**Good:**
```python
def cjk_aware_break(text: str, max_width: float, font, font_size: int) -> list[str]:
    """Break text into lines respecting CJK rules."""
    lines = []
    current = ""
    for char in text:
        test = current + char
        if measure_text_width(test, font, font_size) > max_width:
            if current:
                lines.append(current)
            current = char
        else:
            current = test
    if current:
        lines.append(current)
    return lines
```

The `text_metrics.py` module implements `cjk_line_break()` with proper CJK/Latin boundary handling.

---

## Rule 5: Windows GBK encoding — always use PYTHONIOENCODING=utf-8

Windows defaults to the system locale encoding (GBK/CP936 for Chinese Windows). Python's `print()` and `subprocess` output will fail or produce mojibake when encountering CJK characters outside the GBK range.

**Always set:**
```bash
PYTHONIOENCODING=utf-8 pytest tests/
```

In code:
```python
import os
os.environ["PYTHONIOENCODING"] = "utf-8"
```

This is especially critical for:
- Test output containing CJK strings
- Subprocess calls that may return CJK text
- Writing manifest JSON with CJK content

---

## Rule 6: Font size bisection must account for CJK width

When bisecting font size to fit text within a zone, CJK characters are approximately 1.0x the font size in width (fullwidth), while Latin characters are approximately 0.5-0.6x. A bisection algorithm that assumes Latin widths will overshoot for CJK-heavy text.

**Bad:**
```python
def bisect_font_size(text, max_width, max_height, font_path, min_size=10, max_size=36):
    """Find the largest font size that fits text in the given bounds."""
    lo, hi = min_size, max_size
    while lo < hi:
        mid = (lo + hi + 1) // 2
        w = measure_text_width(text, font_path, mid)
        h = measure_text_height(text, font_path, mid, max_width)
        if w <= max_width and h <= max_height:
            lo = mid
        else:
            hi = mid - 1
    return lo
```

The key is that `measure_text_width` and `measure_text_height` must use Pillow's actual glyph metrics, not character-count heuristics. CJK characters at 14pt are roughly 14pt wide; Latin characters at 14pt are roughly 7-8pt wide. A line of 40 CJK characters needs ~560pt, while 40 Latin characters need only ~280pt.
