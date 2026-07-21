"""HTML-to-PPTX import using the *local snapshot underlay* technique.

This module converts HTML slide decks (increasingly common from AI tools
like Marp, Slidev, reveal.js, and LLM-generated presentations) into
editable PPTX files.

**Local snapshot underlay** -- for each slide the module:

1. Extracts semantic content (text, images, tables) from the HTML.
2. Optionally renders the slide area as a raster image via a headless
   browser (Playwright, when available) and places it as a slide
   background.
3. Overlays editable native textboxes on top of the background image.

This keeps text fully editable while preserving visual fidelity for
complex CSS effects (gradients, shadows, transforms) that cannot be
expressed in OOXML.

Public API
----------
- :func:`import_from_html`       -- HTML file -> editable PPTX
- :func:`html_to_content_spec`   -- HTML file -> ContentSpec (pipeline)
- :func:`detect_html_slides`     -- probe HTML slide structure
"""
from __future__ import annotations

import base64
import logging
import os
import re
import tempfile
import urllib.request
import urllib.parse
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

from pptx_skill.content_model import (
    ContentSpec,
    ElementSpec,
    SlideSpec,
    StableIdGenerator,
)

__all__ = [
    "import_from_html",
    "html_to_content_spec",
    "detect_html_slides",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Standard 16:9 canvas in points (matches the rest of the codebase)
_CANVAS_WIDTH_PT = 959.976
_CANVAS_HEIGHT_PT = 540.0

# CSS pixel -> PowerPoint point conversion.
# Browsers render at 96 dpi; PowerPoint uses 72 dpi.
# 1 CSS px = 0.75 pt.
_CSS_PX_TO_PT = 0.75

# Default browser viewport width used for snapshot rendering (pixels).
_DEFAULT_VIEWPORT_WIDTH = 1280

# Mapping of common CSS font-family values to PowerPoint-safe fonts.
_CSS_FONT_MAP: dict[str, str] = {
    "arial": "Arial",
    "helvetica": "Arial",
    "helvetica neue": "Arial",
    "sans-serif": "Arial",
    "times new roman": "Times New Roman",
    "times": "Times New Roman",
    "serif": "Times New Roman",
    "georgia": "Georgia",
    "courier new": "Courier New",
    "courier": "Courier New",
    "monospace": "Courier New",
    "consolas": "Consolas",
    "verdana": "Verdana",
    "tahoma": "Tahoma",
    "calibri": "Calibri",
    "microsoft yahei": "Microsoft YaHei",
    "simhei": "SimHei",
    "simsun": "SimSun",
    "pingfang sc": "PingFang SC",
    "noto sans sc": "Noto Sans SC",
    "noto sans cjk sc": "Noto Sans SC",
    "source han sans sc": "Source Han Sans SC",
}

# HTML tags that represent text content.
_TEXT_TAGS = frozenset({"h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "span", "a", "div", "label", "figcaption", "blockquote", "strong", "em", "b", "i"})

# Heading tag -> default PowerPoint font size (pt).
_HEADING_SIZES: dict[str, float] = {
    "h1": 36.0,
    "h2": 28.0,
    "h3": 24.0,
    "h4": 20.0,
    "h5": 18.0,
    "h6": 16.0,
}

# Default font size for non-heading text.
_DEFAULT_FONT_SIZE_PT = 14.0


# ---------------------------------------------------------------------------
# CSS unit conversion
# ---------------------------------------------------------------------------

def _css_length_to_pt(value: str, base_font_pt: float = 16.0) -> float | None:
    """Convert a CSS length string to PowerPoint points.

    Supported units: px, pt, em, rem, %, vw, vh (approximate).
    Returns ``None`` for unparseable or unitless values.
    """
    value = value.strip().lower()
    if not value:
        return None

    # Percentage -- cannot resolve without context; treat as fraction of base.
    if value.endswith("%"):
        try:
            return base_font_pt * float(value[:-1]) / 100.0
        except ValueError:
            return None

    # px
    if value.endswith("px"):
        try:
            return float(value[:-2]) * _CSS_PX_TO_PT
        except ValueError:
            return None

    # pt
    if value.endswith("pt"):
        try:
            return float(value[:-2])
        except ValueError:
            return None

    # em (relative to base font)
    if value.endswith("em"):
        try:
            return float(value[:-2]) * base_font_pt
        except ValueError:
            return None

    # rem (relative to 16 pt browser default)
    if value.endswith("rem"):
        try:
            return float(value[:-3]) * 16.0
        except ValueError:
            return None

    # vw -- approximate as fraction of slide width.
    if value.endswith("vw"):
        try:
            return float(value[:-2]) / 100.0 * _CANVAS_WIDTH_PT
        except ValueError:
            return None

    # vh -- approximate as fraction of slide height.
    if value.endswith("vh"):
        try:
            return float(value[:-2]) / 100.0 * _CANVAS_HEIGHT_PT
        except ValueError:
            return None

    # Bare number (treat as px for inline style convenience).
    try:
        return float(value) * _CSS_PX_TO_PT
    except ValueError:
        return None


def _map_css_font(css_font: str) -> str:
    """Map a CSS font-family string to a PowerPoint-safe font name.

    Handles multi-family fallback strings like ``"Helvetica, Arial, sans-serif"``.
    Returns the first recognised font, or the first token if nothing matches.
    """
    if not css_font:
        return "Calibri"

    # Strip quotes, split on comma.
    families = re.split(r"\s*,\s*", css_font.strip().strip('"').strip("'"))
    for family in families:
        key = family.strip().strip('"').strip("'").lower()
        if key in _CSS_FONT_MAP:
            return _CSS_FONT_MAP[key]
    # Fallback: return the first family name, capitalised.
    first = families[0].strip().strip('"').strip("'")
    return first if first else "Calibri"


def _css_color_to_hex(color: str) -> str | None:
    """Convert a CSS color value to a ``#RRGGBB`` hex string.

    Supports ``#rgb``, ``#rrggbb``, ``rgb(r,g,b)``, and named colors
    (a small subset).  Returns ``None`` for unrecognised values.
    """
    color = color.strip().lower()
    if not color or color in ("inherit", "initial", "unset", "transparent", "none"):
        return None

    # #rrggbb
    if re.match(r"^#[0-9a-f]{6}$", color):
        return color

    # #rgb -> #rrggbb
    m = re.match(r"^#([0-9a-f])([0-9a-f])([0-9a-f])$", color)
    if m:
        return f"#{m.group(1)*2}{m.group(2)*2}{m.group(3)*2}"

    # rgb(r, g, b)
    m = re.match(r"^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$", color)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"#{r:02x}{g:02x}{b:02x}"

    # Small named-color lookup.
    _NAMED: dict[str, str] = {
        "black": "#000000", "white": "#ffffff", "red": "#ff0000",
        "green": "#008000", "blue": "#0000ff", "yellow": "#ffff00",
        "gray": "#808080", "grey": "#808080", "silver": "#c0c0c0",
        "navy": "#000080", "teal": "#008080", "maroon": "#800000",
        "olive": "#808000", "purple": "#800080", "lime": "#00ff00",
        "aqua": "#00ffff", "fuchsia": "#ff00ff", "orange": "#ffa500",
    }
    return _NAMED.get(color)


# ---------------------------------------------------------------------------
# Inline style parser
# ---------------------------------------------------------------------------

def _parse_inline_style(style_str: str) -> dict[str, str]:
    """Parse an HTML ``style`` attribute into a ``{property: value}`` dict."""
    result: dict[str, str] = {}
    if not style_str:
        return result
    for declaration in style_str.split(";"):
        declaration = declaration.strip()
        if ":" not in declaration:
            continue
        prop, _, val = declaration.partition(":")
        result[prop.strip().lower()] = val.strip()
    return result


def _extract_text_style(tag: str, style_str: str, *, font_family_override: str | None = None) -> dict[str, Any]:
    """Derive a renderer-compatible style dict from an HTML tag and inline style.

    Returns keys recognised by ``_apply_text_style`` in the renderer:
    ``size``, ``color``, ``bold``, ``italic``, ``font_family``, ``align``.
    """
    css = _parse_inline_style(style_str)
    style: dict[str, Any] = {}

    # Font size
    if "font-size" in css:
        pt = _css_length_to_pt(css["font-size"])
        if pt is not None and pt > 0:
            style["size"] = round(pt, 1)
    elif tag in _HEADING_SIZES:
        style["size"] = _HEADING_SIZES[tag]

    # Font family
    if font_family_override:
        style["font_family"] = font_family_override
    elif "font-family" in css:
        style["font_family"] = _map_css_font(css["font-family"])

    # Color
    if "color" in css:
        hex_color = _css_color_to_hex(css["color"])
        if hex_color:
            style["color"] = hex_color

    # Bold
    if tag in ("b", "strong") or css.get("font-weight") in ("bold", "bolder", "700", "800", "900"):
        style["bold"] = True
    elif css.get("font-weight"):
        try:
            if int(css["font-weight"]) >= 700:
                style["bold"] = True
        except ValueError:
            pass

    # Italic
    if tag in ("i", "em") or css.get("font-style") in ("italic", "oblique"):
        style["italic"] = True

    # Text alignment
    if "text-align" in css:
        align = css["text-align"].strip().lower()
        if align in ("left", "center", "right", "justify"):
            style["align"] = align

    return style


# ---------------------------------------------------------------------------
# HTML parser -- extract structured content from slide elements
# ---------------------------------------------------------------------------

class _SlideContentParser(HTMLParser):
    """Parse the inner HTML of a single slide element.

    Collects text runs of text with their associated styles, image references,
    and table data.
    """

    def __init__(self, *, font_family_override: str | None = None) -> None:
        super().__init__()
        self.font_family_override = font_family_override
        # Collected data
        self.text_runs: list[dict[str, Any]] = []  # {"text": str, "style": dict, "tag": str}
        self.images: list[dict[str, str]] = []      # {"src": str, "alt": str}
        self.tables: list[dict[str, Any]] = []       # {"headers": [...], "rows": [[...], ...]}
        # Parser state
        self._tag_stack: list[str] = []
        self._style_stack: list[dict[str, Any]] = []
        self._current_text: list[str] = []
        self._current_style: dict[str, Any] = {}
        self._current_tag: str = ""
        # Table state
        self._in_table: bool = False
        self._in_thead: bool = False
        self._in_tbody: bool = False
        self._in_row: bool = False
        self._in_cell: bool = False
        self._table_headers: list[str] = []
        self._table_rows: list[list[str]] = []
        self._current_row: list[str] = []
        self._current_cell_text: list[str] = []
        # Image state
        self._current_img_alt: str = ""

    # -- tag tracking --------------------------------------------------------

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag_lower = tag.lower()
        attr_dict = {k.lower(): (v or "") for k, v in attrs}

        # Image
        if tag_lower == "img":
            src = attr_dict.get("src", "")
            alt = attr_dict.get("alt", "")
            if src:
                self.images.append({"src": src, "alt": alt})
            return

        # Table
        if tag_lower == "table":
            self._in_table = True
            self._table_headers = []
            self._table_rows = []
            return
        if tag_lower == "thead":
            self._in_thead = True
            return
        if tag_lower == "tbody":
            self._in_tbody = True
            return
        if tag_lower == "tr":
            self._in_row = True
            self._current_row = []
            return
        if tag_lower in ("td", "th"):
            self._in_cell = True
            self._current_cell_text = []
            return

        # Text element
        if tag_lower in _TEXT_TAGS:
            # Flush any pending text from the previous element.
            self._flush_text()
            style_str = attr_dict.get("style", "")
            style = _extract_text_style(tag_lower, style_str, font_family_override=self.font_family_override)
            self._tag_stack.append(tag_lower)
            self._style_stack.append(style)
            self._current_tag = tag_lower
            self._current_style = style
            self._current_text = []

        # br -- line break within text
        if tag_lower == "br":
            self._current_text.append("\n")

    def handle_endtag(self, tag: str) -> None:
        tag_lower = tag.lower()

        # Table
        if tag_lower in ("td", "th"):
            cell_text = "".join(self._current_cell_text).strip()
            is_header = tag_lower == "th"
            self._in_cell = False
            if self._in_row:
                self._current_row.append(cell_text)
                if is_header and self._in_thead:
                    self._table_headers.append(cell_text)
            return
        if tag_lower == "tr":
            self._in_row = False
            if self._current_row:
                if self._in_thead and not self._table_headers:
                    self._table_headers = list(self._current_row)
                else:
                    self._table_rows.append(list(self._current_row))
            self._current_row = []
            return
        if tag_lower == "thead":
            self._in_thead = False
            return
        if tag_lower == "tbody":
            self._in_tbody = False
            return
        if tag_lower == "table":
            self._in_table = False
            if self._table_headers or self._table_rows:
                self.tables.append({
                    "headers": self._table_headers,
                    "rows": self._table_rows,
                })
            return

        # Text element
        if tag_lower in _TEXT_TAGS and self._tag_stack:
            self._flush_text()
            self._tag_stack.pop()
            self._style_stack.pop()
            if self._tag_stack:
                self._current_tag = self._tag_stack[-1]
                self._current_style = self._style_stack[-1] if self._style_stack else {}
            else:
                self._current_tag = ""
                self._current_style = {}

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._current_cell_text.append(data)
            return
        if self._tag_stack:
            self._current_text.append(data)

    def _flush_text(self) -> None:
        text = "".join(self._current_text).strip()
        if text:
            self.text_runs.append({
                "text": text,
                "style": dict(self._current_style),
                "tag": self._current_tag,
            })
        self._current_text = []


class _SlideDetectorParser(HTMLParser):
    """Lightweight parser to detect slide containers and count elements.

    Used by :func:`detect_html_slides`.
    """

    def __init__(self, *, slide_selector: str = "div.slide") -> None:
        super().__init__()
        self._selector_tag, self._selector_class = self._parse_selector(slide_selector)
        self._depth: int = 0
        self._in_slide: bool = False
        self._slide_depth: int = 0
        self._element_count: int = 0
        self._title_text: list[str] = []
        self._in_heading: bool = False
        self.slides: list[dict[str, Any]] = []
        self._slide_index: int = 0

    @staticmethod
    def _parse_selector(selector: str) -> tuple[str, str]:
        """Parse a simple CSS selector like ``div.slide`` or ``section``.

        Only supports ``tag.class`` and bare ``tag`` forms.
        """
        if "." in selector:
            tag, _, cls = selector.partition(".")
            return tag.lower(), cls.lower()
        return selector.lower(), ""

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag_lower = tag.lower()
        attr_dict = {k.lower(): (v or "") for k, v in attrs}

        # Detect slide container start
        if not self._in_slide:
            if tag_lower == self._selector_tag:
                if self._selector_class:
                    classes = attr_dict.get("class", "").lower().split()
                    if self._selector_class in classes:
                        self._in_slide = True
                        self._slide_depth = 0
                        self._element_count = 0
                        self._title_text = []
                else:
                    self._in_slide = True
                    self._slide_depth = 0
                    self._element_count = 0
                    self._title_text = []
        else:
            self._slide_depth += 1
            # Count meaningful elements
            if tag_lower in _TEXT_TAGS or tag_lower == "img" or tag_lower == "table":
                self._element_count += 1
            # Capture first heading as title
            if tag_lower in ("h1", "h2", "h3") and not self._title_text:
                self._in_heading = True

    def handle_endtag(self, tag: str) -> None:
        tag_lower = tag.lower()
        if self._in_heading and tag_lower in ("h1", "h2", "h3"):
            self._in_heading = False
        if self._in_slide:
            self._slide_depth -= 1
            if self._slide_depth < 0 and tag_lower == self._selector_tag:
                # Slide container closed
                title = "".join(self._title_text).strip() or f"Slide {self._slide_index + 1}"
                self.slides.append({
                    "index": self._slide_index,
                    "title": title,
                    "element_count": self._element_count,
                })
                self._slide_index += 1
                self._in_slide = False

    def handle_data(self, data: str) -> None:
        if self._in_heading:
            self._title_text.append(data)


# ---------------------------------------------------------------------------
# Image resolution
# ---------------------------------------------------------------------------

def _resolve_image_src(src: str, html_dir: str) -> str | None:
    """Resolve an image ``src`` to a local file path.

    Handles:
    - Relative paths (resolved against *html_dir*)
    - Absolute local paths
    - ``data:`` URIs (base64-decoded to a temp file)
    - ``http://`` / ``https://`` URLs (downloaded to a temp file)

    Returns the local file path, or ``None`` on failure.
    """
    if not src:
        return None

    # data: URI
    if src.startswith("data:"):
        return _decode_data_uri(src)

    # http/https URL
    if src.startswith(("http://", "https://")):
        return _download_url(src)

    # Local path (relative or absolute)
    if os.path.isabs(src):
        if os.path.isfile(src):
            return src
        return None

    # Relative path
    full_path = os.path.normpath(os.path.join(html_dir, src))
    if os.path.isfile(full_path):
        return full_path
    return None


def _decode_data_uri(uri: str) -> str | None:
    """Decode a ``data:image/...;base64,...`` URI to a temporary file."""
    try:
        # data:[<mediatype>][;base64],<data>
        if not uri.startswith("data:"):
            return None
        header, _, data = uri.partition(",")
        if not data:
            return None
        is_base64 = ";base64" in header
        # Extract extension from mediatype
        mediatype = header[5:]  # strip "data:"
        ext = _mediatype_to_ext(mediatype)
        tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
        try:
            if is_base64:
                tmp.write(base64.b64decode(data))
            else:
                tmp.write(data.encode("utf-8"))
            tmp.close()
            return tmp.name
        except Exception:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass
            return None
    except Exception:
        log.warning("Failed to decode data URI")
        return None


def _mediatype_to_ext(mediatype: str) -> str:
    """Return a file extension for a MIME mediatype string."""
    mt = mediatype.lower().split(";")[0].strip()
    _map = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/svg+xml": ".svg",
        "image/webp": ".webp",
        "image/bmp": ".bmp",
        "image/x-icon": ".ico",
    }
    return _map.get(mt, ".png")


def _download_url(url: str) -> str | None:
    """Download a URL to a temporary file. Returns path or ``None``."""
    try:
        ext = os.path.splitext(urllib.parse.urlparse(url).path)[1] or ".png"
        tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
        try:
            with urllib.request.urlopen(url, timeout=15) as resp:  # noqa: S310
                tmp.write(resp.read())
            tmp.close()
            return tmp.name
        except Exception:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass
            return None
    except Exception:
        log.warning("Failed to download image: %s", url[:120])
        return None


# ---------------------------------------------------------------------------
# Playwright snapshot rendering
# ---------------------------------------------------------------------------

def _render_slide_snapshot(
    html_path: str,
    slide_index: int,
    slide_selector: str,
    output_dir: str,
) -> str | None:
    """Render a single slide as a PNG image using Playwright.

    Returns the path to the rendered PNG, or ``None`` if Playwright is
    not available or rendering fails.
    """
    try:
        from playwright.sync_api import sync_playwright  # type: ignore[import-untyped]
    except ImportError:
        log.info("Playwright not available; skipping snapshot underlay")
        return None

    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page(
                viewport={"width": _DEFAULT_VIEWPORT_WIDTH, "height": int(_DEFAULT_VIEWPORT_WIDTH * 9 / 16)},
            )
            abs_path = os.path.abspath(html_path)
            file_url = f"file:///{abs_path.replace(os.sep, '/')}"
            page.goto(file_url, wait_until="networkidle", timeout=30_000)

            # Locate the slide element
            slides = page.query_selector_all(slide_selector)
            if slide_index >= len(slides):
                browser.close()
                return None

            slide_el = slides[slide_index]
            out_path = os.path.join(output_dir, f"_snapshot_{slide_index}.png")
            slide_el.screenshot(path=out_path)
            browser.close()
            return out_path
    except Exception:
        log.warning("Playwright snapshot rendering failed for slide %d", slide_index, exc_info=True)
        return None


# ---------------------------------------------------------------------------
# PPTX construction helpers
# ---------------------------------------------------------------------------

def _pt_to_inches(pt: float) -> float:
    """Convert points to inches (1 inch = 72 pt)."""
    return pt / 72.0


def _hex_to_rgb(color: str) -> tuple[int, int, int]:
    """Parse ``#RRGGBB`` to an (R, G, B) tuple."""
    color = color.lstrip("#")
    if len(color) == 3:
        color = "".join(c * 2 for c in color)
    return (int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))


def _add_textbox_to_slide(slide, left_pt: float, top_pt: float, width_pt: float, height_pt: float,
                          text: str, style: dict[str, Any]) -> None:
    """Add an editable textbox to a python-pptx slide."""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt

    left = Inches(_pt_to_inches(left_pt))
    top = Inches(_pt_to_inches(top_pt))
    width = Inches(_pt_to_inches(max(width_pt, 1.0)))
    height = Inches(_pt_to_inches(max(height_pt, 1.0)))

    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True

    # Split on newlines for multi-paragraph content
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        run = p.add_run()
        run.text = line

        # Apply style
        if "size" in style:
            run.font.size = Pt(style["size"])
        if "color" in style:
            try:
                run.font.color.rgb = RGBColor(*_hex_to_rgb(style["color"]))
            except Exception:
                pass
        if style.get("bold"):
            run.font.bold = True
        if style.get("italic"):
            run.font.italic = True
        font_family = style.get("font_family")
        if font_family:
            run.font.name = font_family
            # Set East-Asian typeface for CJK consistency
            try:
                rPr = run._r.get_or_add_rPr()
                ea = rPr.find(qn("a:ea"))
                if ea is None:
                    ea = rPr.makeelement(qn("a:ea"), {})
                    rPr.append(ea)
                ea.set("typeface", font_family)
            except Exception:
                pass

        # Alignment
        align = style.get("align", "left")
        align_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY,
        }
        p.alignment = align_map.get(align, PP_ALIGN.LEFT)


def _add_image_to_slide(slide, image_path: str, left_pt: float, top_pt: float,
                        width_pt: float, height_pt: float) -> None:
    """Add an image to a python-pptx slide."""
    from pptx.util import Inches

    left = Inches(_pt_to_inches(left_pt))
    top = Inches(_pt_to_inches(top_pt))
    width = Inches(_pt_to_inches(max(width_pt, 1.0)))
    height = Inches(_pt_to_inches(max(height_pt, 1.0)))

    if os.path.isfile(image_path):
        try:
            slide.shapes.add_picture(image_path, left, top, width, height)
        except Exception:
            log.warning("Failed to add image: %s", image_path[:120])
    else:
        # Placeholder rectangle
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE

        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(220, 220, 220)


def _add_background_image(slide, image_path: str, slide_width_pt: float, slide_height_pt: float) -> None:
    """Place an image as a full-slide background (z-order 0 underlay)."""
    from pptx.util import Inches

    if not image_path or not os.path.isfile(image_path):
        return

    left = Inches(0)
    top = Inches(0)
    width = Inches(_pt_to_inches(slide_width_pt))
    height = Inches(_pt_to_inches(slide_height_pt))

    try:
        pic = slide.shapes.add_picture(image_path, left, top, width, height)
        # Move to back (z-order 0) by moving element to first position in spTree
        sp_tree = slide.shapes._spTree
        sp_tree.remove(pic._element)
        sp_tree.insert(2, pic._element)  # index 2 = after the mandatory cSld>spTree children
    except Exception:
        log.warning("Failed to add background snapshot image")


def _add_table_to_slide(slide, headers: list[str], rows: list[list[str]],
                        left_pt: float, top_pt: float, width_pt: float, height_pt: float) -> None:
    """Add a table to a python-pptx slide."""
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    n_rows = len(rows) + (1 if headers else 0)
    n_cols = max(len(headers), max((len(r) for r in rows), default=0))
    if n_rows == 0 or n_cols == 0:
        return

    left = Inches(_pt_to_inches(left_pt))
    top = Inches(_pt_to_inches(top_pt))
    width = Inches(_pt_to_inches(max(width_pt, 1.0)))
    height = Inches(_pt_to_inches(max(height_pt, 1.0)))

    table_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = table_shape.table

    # Header row
    if headers:
        for col_idx, header_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = header_text
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(12)

    # Data rows
    start_row = 1 if headers else 0
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            if col_idx < n_cols:
                cell = table.cell(start_row + row_idx, col_idx)
                cell.text = cell_text
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(11)


# ---------------------------------------------------------------------------
# Slide layout estimation
# ---------------------------------------------------------------------------

def _estimate_text_position(
    run_index: int,
    total_runs: int,
    tag: str,
    canvas_width: float,
    canvas_height: float,
) -> tuple[float, float, float, float]:
    """Estimate (left, top, width, height) in points for a text run.

    Uses a simple heuristic layout:
    - Headings are placed at the top with large width.
    - Body text flows below headings.
    - Multiple runs are stacked vertically.
    """
    safe_left = 48.0
    safe_top = 36.0
    safe_right = 48.0
    safe_bottom = 32.0
    content_width = canvas_width - safe_left - safe_right

    if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
        # Headings: full width, near top
        top_offset = safe_top + run_index * 50.0
        height = 50.0
        return (safe_left, top_offset, content_width, height)
    else:
        # Body text: below headings, stacked
        heading_space = 60.0  # approximate space taken by headings
        line_height = 24.0
        top_offset = safe_top + heading_space + run_index * line_height
        height = line_height
        return (safe_left, top_offset, content_width, height)


def _estimate_image_position(
    img_index: int,
    total_images: int,
    canvas_width: float,
    canvas_height: float,
) -> tuple[float, float, float, float]:
    """Estimate position for an image on the slide."""
    safe_left = 48.0
    safe_top = 36.0
    safe_right = 48.0
    content_width = canvas_width - safe_left - safe_right

    # Place images in a grid-like arrangement
    img_height = (canvas_height - safe_top - 32.0) * 0.5
    img_width = content_width / max(total_images, 1)

    col = img_index % max(total_images, 1)
    row = img_index // max(total_images, 1)

    left = safe_left + col * img_width
    top = safe_top + 80.0 + row * (img_height + 8.0)  # offset below heading area

    return (left, top, img_width, img_height)


def _estimate_table_position(
    canvas_width: float,
    canvas_height: float,
) -> tuple[float, float, float, float]:
    """Estimate position for a table on the slide."""
    safe_left = 48.0
    safe_top = 36.0
    safe_right = 48.0
    content_width = canvas_width - safe_left - safe_right
    table_height = min(canvas_height * 0.6, 300.0)
    return (safe_left, safe_top + 80.0, content_width, table_height)


# ---------------------------------------------------------------------------
# Theme application
# ---------------------------------------------------------------------------

def _apply_theme_to_presentation(prs, theme_key: str | None) -> None:
    """Apply a theme key to the presentation's slide masters.

    This is a best-effort operation that sets background and default fonts
    when a known theme key is provided.
    """
    if not theme_key:
        return

    # Minimal theme application: set slide background color for known themes.
    # The full theme system lives in scripts/pptx_helper.py THEMES dict.
    _THEME_BACKGROUNDS: dict[str, str] = {
        "editorial": "#FFFFFF",
        "technical": "#1A1A2E",
        "minimal": "#FFFFFF",
        "corporate": "#FFFFFF",
        "dark": "#1A1A2E",
        "nature": "#F5F5DC",
        "creative": "#FFF8DC",
        "academic": "#FFFFFF",
        "consulting": "#FFFFFF",
        "startup": "#FFFFFF",
    }

    bg_color = _THEME_BACKGROUNDS.get(theme_key)
    if not bg_color:
        return

    try:
        from pptx.dml.color import RGBColor

        rgb = RGBColor(*_hex_to_rgb(bg_color))
        for slide_master in prs.slide_masters:
            background = slide_master.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = rgb
    except Exception:
        log.debug("Theme background application failed for %s", theme_key)


# ---------------------------------------------------------------------------
# Public API: detect_html_slides
# ---------------------------------------------------------------------------

def detect_html_slides(html_path: str | os.PathLike[str], *, slide_selector: str = "div.slide") -> list[dict[str, Any]]:
    """Detect slide structure in an HTML file.

    Parameters
    ----------
    html_path :
        Path to the HTML file.
    slide_selector :
        CSS selector for slide containers.  Only simple ``tag.class``
        and bare ``tag`` selectors are supported.

    Returns
    -------
    list[dict]
        Each dict has keys ``index`` (0-based), ``title`` (str), and
        ``element_count`` (int).  Returns an empty list if the file
        cannot be read or no slides are found.
    """
    html_path = os.fspath(html_path)
    if not os.path.isfile(html_path):
        return []

    try:
        with open(html_path, "r", encoding="utf-8", errors="replace") as fh:
            html_content = fh.read()
    except OSError:
        return []

    parser = _SlideDetectorParser(slide_selector=slide_selector)
    try:
        parser.feed(html_content)
    except Exception:
        log.warning("HTML parsing failed during slide detection")
        return []

    return parser.slides


# ---------------------------------------------------------------------------
# Public API: html_to_content_spec
# ---------------------------------------------------------------------------

def html_to_content_spec(
    html_path: str | os.PathLike[str],
    *,
    slide_selector: str = "div.slide",
) -> ContentSpec:
    """Convert an HTML slide deck to a :class:`ContentSpec`.

    This produces a content specification that can be fed into the
    generation pipeline (``plan_deck`` -> ``render_layout_plans``) for
    higher-quality layout than the direct import path.

    Parameters
    ----------
    html_path :
        Path to the HTML file.
    slide_selector :
        CSS selector for slide containers.

    Returns
    -------
    ContentSpec
        A content spec with one :class:`SlideSpec` per detected slide.
        If the file cannot be read or no slides are found, returns a
        ContentSpec with an empty slide list.
    """
    html_path = os.fspath(html_path)
    html_dir = os.path.dirname(os.path.abspath(html_path))

    if not os.path.isfile(html_path):
        return ContentSpec(id="html-import-empty", title="", subtitle="", slides=[])

    try:
        with open(html_path, "r", encoding="utf-8", errors="replace") as fh:
            html_content = fh.read()
    except OSError:
        return ContentSpec(id="html-import-empty", title="", subtitle="", slides=[])

    # Split HTML into slide chunks
    slide_chunks = _split_html_into_slides(html_content, slide_selector)
    if not slide_chunks:
        return ContentSpec(id="html-import-empty", title="", subtitle="", slides=[])

    gen = StableIdGenerator()
    deck_id = gen._join("html-import", os.path.basename(html_path))
    slides: list[SlideSpec] = []

    for slide_idx, chunk_html in enumerate(slide_chunks):
        parser = _SlideContentParser()
        try:
            parser.feed(chunk_html)
        except Exception:
            log.warning("HTML parsing failed for slide %d", slide_idx)
            continue

        # Flush any remaining text
        parser._flush_text()

        source_id = gen._join(deck_id, f"slide-{slide_idx}")
        elements: list[ElementSpec] = []

        # Determine role from content
        role = _infer_slide_role(parser)

        # Title element
        title_text = _extract_title(parser.text_runs)
        if title_text:
            elements.append(ElementSpec(
                id=gen.element_id(source_id, 0, "title", "title"),
                kind="text",
                role="title",
                content={"text": title_text},
                style_ref="component.title",
            ))

        # Body text elements
        body_items: list[str] = []
        for run in parser.text_runs:
            tag = run.get("tag", "")
            text = run.get("text", "")
            if tag in ("h1", "h2", "h3") and text == title_text:
                continue  # already captured as title
            if tag in ("p", "li", "span", "div", "a"):
                body_items.append(text)

        if body_items:
            elements.append(ElementSpec(
                id=gen.element_id(source_id, 0, "body", "body"),
                kind="text",
                role="body",
                content={"items": body_items},
                style_ref="component.body",
            ))

        # Subtitle / secondary heading
        for run in parser.text_runs:
            tag = run.get("tag", "")
            text = run.get("text", "")
            if tag in ("h2", "h3") and text != title_text and not any(
                e.role == "subtitle" for e in elements
            ):
                elements.append(ElementSpec(
                    id=gen.element_id(source_id, 0, "subtitle", "subtitle"),
                    kind="text",
                    role="subtitle",
                    content={"text": text},
                    style_ref="component.subtitle",
                ))
                break

        # Images
        for img_idx, img in enumerate(parser.images):
            resolved = _resolve_image_src(img["src"], html_dir)
            if resolved:
                elements.append(ElementSpec(
                    id=gen.element_id(source_id, 0, "hero" if img_idx == 0 else "supporting_image", f"image-{img_idx}"),
                    kind="image",
                    role="hero" if img_idx == 0 else "supporting_image",
                    content={"path": resolved, "fit": "smart-cover"},
                    style_ref="component.image",
                ))

        # Tables
        for tbl_idx, table in enumerate(parser.tables):
            elements.append(ElementSpec(
                id=gen.element_id(source_id, 0, "table", f"table-{tbl_idx}"),
                kind="table",
                role="table",
                content={"headers": table.get("headers", []), "rows": table.get("rows", [])},
                style_ref="component.table",
            ))

        if elements:
            slides.append(SlideSpec(
                id=gen.slide_id(source_id, 0),
                role=role,
                communication_goal=title_text or f"Slide {slide_idx + 1}",
                elements=elements,
                source_section_id=source_id,
            ))

    # Derive deck title from first slide
    deck_title = ""
    if slides:
        for elem in slides[0].elements:
            if elem.role == "title" and elem.content.get("text"):
                deck_title = elem.content["text"]
                break

    return ContentSpec(
        id=deck_id,
        title=deck_title,
        subtitle="",
        slides=slides,
        locale="zh-CN",
        metadata={"source": "html_import", "html_path": html_path},
    )


def _split_html_into_slides(html_content: str, slide_selector: str) -> list[str]:
    """Split full HTML into per-slide HTML chunks.

    Uses a simple regex-based approach to extract the inner HTML of each
    slide container element.  For complex HTML structures, the
    ``_SlideDetectorParser`` is used as a fallback.
    """
    # Parse selector
    if "." in slide_selector:
        tag, _, cls = slide_selector.partition(".")
        tag = tag.lower()
        cls = cls.lower()
    else:
        tag = slide_selector.lower()
        cls = ""

    # Try regex extraction for simple cases
    if cls:
        # Match <tag class="...cls...">...</tag>
        # This regex is intentionally simple; complex nesting may not be handled.
        pattern = re.compile(
            rf"<{tag}\s+[^>]*class=[\"'][^\"']*\b{re.escape(cls)}\b[^\"']*[\"'][^>]*>(.*?)</{tag}>",
            re.DOTALL | re.IGNORECASE,
        )
        matches = pattern.findall(html_content)
        if matches:
            return matches

    # Fallback: use the detector parser to find slide boundaries,
    # then extract substrings.
    return _extract_slides_by_parser(html_content, slide_selector)


def _extract_slides_by_parser(html_content: str, slide_selector: str) -> list[str]:
    """Extract slide HTML chunks using a position-tracking parser."""
    # We use a custom parser that records start/end positions of slide containers.
    positions: list[tuple[int, int]] = []

    if "." in slide_selector:
        sel_tag, _, sel_cls = slide_selector.partition(".")
        sel_tag = sel_tag.lower()
        sel_cls = sel_cls.lower()
    else:
        sel_tag = slide_selector.lower()
        sel_cls = ""

    class _PositionParser(HTMLParser):
        def __init__(self) -> None:
            super().__init__()
            self._depth = 0
            self._in_slide = False
            self._slide_start = 0
            self._slide_depth = 0

        def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
            nonlocal positions
            tag_lower = tag.lower()
            attr_dict = {k.lower(): (v or "") for k, v in attrs}

            if not self._in_slide:
                if tag_lower == sel_tag:
                    if sel_cls:
                        classes = attr_dict.get("class", "").lower().split()
                        if sel_cls in classes:
                            self._in_slide = True
                            self._slide_depth = 0
                            # Record approximate start position
                            self._slide_start = self.getpos()
                    else:
                        self._in_slide = True
                        self._slide_depth = 0
                        self._slide_start = self.getpos()
            else:
                self._slide_depth += 1

        def handle_endtag(self, tag: str) -> None:
            tag_lower = tag.lower()
            if self._in_slide:
                self._slide_depth -= 1
                if self._slide_depth < 0 and tag_lower == sel_tag:
                    end_pos = self.getpos()
                    positions.append((self._slide_start, end_pos))
                    self._in_slide = False

    parser = _PositionParser()
    parser.feed(html_content)

    # Convert line/col positions to character offsets and extract substrings
    lines = html_content.split("\n")
    line_offsets = [0]
    for line in lines:
        line_offsets.append(line_offsets[-1] + len(line) + 1)

    chunks: list[str] = []
    for start_pos, end_pos in positions:
        start_line, start_col = start_pos
        end_line, end_col = end_pos
        start_offset = line_offsets[start_line - 1] + start_col - 1 if start_line <= len(line_offsets) else 0
        end_offset = line_offsets[end_line - 1] + end_col - 1 if end_line <= len(line_offsets) else len(html_content)
        chunk = html_content[start_offset:end_offset]
        if chunk.strip():
            chunks.append(chunk)

    return chunks


def _infer_slide_role(parser: _SlideContentParser) -> str:
    """Infer the slide role from parsed content."""
    has_heading = any(r.get("tag") in ("h1", "h2") for r in parser.text_runs)
    n_images = len(parser.images)
    has_bullets = any(r.get("tag") == "li" for r in parser.text_runs)
    has_table = bool(parser.tables)

    if not parser.text_runs and n_images >= 1:
        return "full_image"
    if has_table:
        return "table"
    if n_images >= 3:
        return "image_grid"
    if n_images >= 1 and has_bullets:
        return "text_image"
    if n_images >= 1:
        return "full_image"
    if has_bullets:
        return "bullets"
    if has_heading and not has_bullets:
        return "section"
    return "bullets"


def _extract_title(text_runs: list[dict[str, Any]]) -> str:
    """Extract the title text from a list of parsed text runs."""
    # Prefer h1, then h2, then first run
    for tag in ("h1", "h2", "h3"):
        for run in text_runs:
            if run.get("tag") == tag and run.get("text", "").strip():
                return run["text"].strip()
    # Fallback: first non-empty run
    for run in text_runs:
        if run.get("text", "").strip():
            return run["text"].strip()
    return ""


# ---------------------------------------------------------------------------
# Public API: import_from_html
# ---------------------------------------------------------------------------

def import_from_html(
    html_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
    *,
    slide_selector: str = "div.slide",
    theme: str | None = None,
    font_family: str | None = None,
) -> str:
    """Import an HTML slide deck as an editable PPTX file.

    Uses the *local snapshot underlay* technique:

    1. Each slide element is parsed for text, images, and tables.
    2. If Playwright is available, the slide is also rendered as a
       raster image and placed as a background underlay.
    3. Editable native textboxes are overlaid on top, keeping text
       fully editable while preserving visual fidelity for complex CSS
       effects.

    Parameters
    ----------
    html_path :
        Path to the HTML file.
    output_path :
        Path where the PPTX file will be saved.
    slide_selector :
        CSS selector for slide containers (default ``"div.slide"``).
        Only simple ``tag.class`` and bare ``tag`` selectors are
        supported.
    theme :
        Optional theme key for styling (e.g. ``"editorial"``,
        ``"dark"``).
    font_family :
        Optional font family override applied to all text elements.

    Returns
    -------
    str
        The path to the generated PPTX file.

    Raises
    ------
    FileNotFoundError
        If *html_path* does not exist.
    ValueError
        If no slide elements are found in the HTML.
    """
    from pptx import Presentation
    from pptx.util import Inches

    html_path = os.fspath(html_path)
    output_path = os.fspath(output_path)
    html_dir = os.path.dirname(os.path.abspath(html_path))

    if not os.path.isfile(html_path):
        raise FileNotFoundError(f"HTML file not found: {html_path}")

    # Read HTML content
    with open(html_path, "r", encoding="utf-8", errors="replace") as fh:
        html_content = fh.read()

    # Split into slide chunks
    slide_chunks = _split_html_into_slides(html_content, slide_selector)
    if not slide_chunks:
        raise ValueError(
            f"No slide elements found with selector '{slide_selector}' in {html_path}"
        )

    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(_pt_to_inches(_CANVAS_WIDTH_PT))
    prs.slide_height = Inches(_pt_to_inches(_CANVAS_HEIGHT_PT))

    # Get blank layout
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]

    # Apply theme
    _apply_theme_to_presentation(prs, theme)

    # Resolve font family override
    resolved_font = _map_css_font(font_family) if font_family else None

    # Temporary directory for snapshot images and downloaded files
    with tempfile.TemporaryDirectory() as tmp_dir:
        for slide_idx, chunk_html in enumerate(slide_chunks):
            # Parse slide content
            parser = _SlideContentParser(font_family_override=resolved_font)
            try:
                parser.feed(chunk_html)
            except Exception:
                log.warning("HTML parsing failed for slide %d; creating blank slide", slide_idx)
                prs.slides.add_slide(blank_layout)
                continue

            # Flush remaining text
            parser._flush_text()

            # Create slide
            slide = prs.slides.add_slide(blank_layout)

            # --- Snapshot underlay (if Playwright available) ---
            snapshot_path = _render_slide_snapshot(html_path, slide_idx, slide_selector, tmp_dir)
            if snapshot_path:
                _add_background_image(slide, snapshot_path, _CANVAS_WIDTH_PT, _CANVAS_HEIGHT_PT)

            # --- Text elements ---
            text_run_idx = 0
            for run in parser.text_runs:
                text = run.get("text", "")
                tag = run.get("tag", "")
                style = run.get("style", {})

                if not text.strip():
                    text_run_idx += 1
                    continue

                # Override font family if specified
                if resolved_font:
                    style["font_family"] = resolved_font

                # Estimate position
                left, top, width, height = _estimate_text_position(
                    text_run_idx, len(parser.text_runs), tag,
                    _CANVAS_WIDTH_PT, _CANVAS_HEIGHT_PT,
                )

                _add_textbox_to_slide(slide, left, top, width, height, text, style)
                text_run_idx += 1

            # --- Images ---
            for img_idx, img in enumerate(parser.images):
                resolved_src = _resolve_image_src(img["src"], html_dir)
                if not resolved_src:
                    log.info("Could not resolve image src: %s", img["src"][:80])
                    continue

                left, top, width, height = _estimate_image_position(
                    img_idx, len(parser.images),
                    _CANVAS_WIDTH_PT, _CANVAS_HEIGHT_PT,
                )
                _add_image_to_slide(slide, resolved_src, left, top, width, height)

            # --- Tables ---
            for table in parser.tables:
                headers = table.get("headers", [])
                rows = table.get("rows", [])
                left, top, width, height = _estimate_table_position(
                    _CANVAS_WIDTH_PT, _CANVAS_HEIGHT_PT,
                )
                _add_table_to_slide(slide, headers, rows, left, top, width, height)

    # Ensure output directory exists
    out_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(out_dir, exist_ok=True)

    # Save
    prs.save(output_path)
    log.info("Imported %d slides from HTML -> %s", len(slide_chunks), output_path)

    return output_path
