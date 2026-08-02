"""Advanced color manipulation for PPTX — theme colors, palette extraction, recoloring, and gradients.

This module provides color utilities that go beyond simple hex replacement:

- **Theme color resolution**: resolve theme color references to concrete RGB values
- **Palette extraction**: extract the dominant color palette from a presentation
- **Color harmony**: generate complementary, analogous, triadic, and split-complementary palettes
- **Gradient builder**: create linear, radial, and path gradients for shapes
- **Batch recoloring**: replace colors across an entire presentation with mapping rules
- **Color conversion**: RGB ↔ HSL ↔ HSV ↔ CMYK conversions

Quick start
-----------
>>> from pptx_skill.color import extract_palette, complementary_colors
>>> palette = extract_palette("deck.pptx")
>>> comp = complementary_colors("#2D5016")
"""
from __future__ import annotations

import colorsys
import logging
import re
from dataclasses import dataclass
from typing import Any

__all__ = [
    "ColorInfo",
    "GradientStop",
    "GradientInfo",
    "extract_palette",
    "resolve_theme_color",
    "rgb_to_hsl",
    "hsl_to_rgb",
    "rgb_to_hsv",
    "hsv_to_rgb",
    "rgb_to_cmyk",
    "cmyk_to_rgb",
    "hex_to_rgb",
    "rgb_to_hex",
    "lighten",
    "darken",
    "complementary_colors",
    "analogous_colors",
    "triadic_colors",
    "split_complementary_colors",
    "tetradic_colors",
    "monochromatic_colors",
    "contrast_color",
    "blend_colors",
    "apply_gradient",
    "remove_gradient",
    "list_gradients",
    "recolor_presentation",
    "swap_colors",
]

log = logging.getLogger(__name__)

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class ColorInfo:
    """Info about a color occurrence in a presentation."""
    hex_value: str = ""
    count: int = 0
    context: str = ""  # "text", "fill", "line", "background"
    hsl: tuple[float, float, float] = (0.0, 0.0, 0.0)


@dataclass
class GradientStop:
    """A single gradient stop."""
    position: float = 0.0  # 0.0 to 1.0
    color: str = "#000000"  # hex
    alpha: int = 100  # 0-100


@dataclass
class GradientInfo:
    """Info about a gradient fill."""
    gradient_type: str = "linear"  # "linear", "radial", "path"
    angle: float = 0.0  # degrees, for linear
    stops: list[GradientStop] = None

    def __post_init__(self):
        if self.stops is None:
            self.stops = []


# ---------------------------------------------------------------------------
# Color conversion utilities
# ---------------------------------------------------------------------------

def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert hex color string to (R, G, B) tuple.

    Accepts ``"#RRGGBB"``, ``"RRGGBB"``, or ``"#RGB"`` (short form).
    """
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    if len(h) != 6:
        raise ValueError(f"Invalid hex color: {hex_color!r}")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert (R, G, B) to ``"#RRGGBB"`` hex string."""
    return f"#{r:02X}{g:02X}{b:02X}"


def rgb_to_hsl(r: int, g: int, b: int) -> tuple[float, float, float]:
    """Convert RGB (0-255) to HSL (0-360, 0-1, 0-1)."""
    r_n, g_n, b_n = r / 255.0, g / 255.0, b / 255.0
    h, l, s = colorsys.rgb_to_hls(r_n, g_n, b_n)
    return (h * 360, s, l)


def hsl_to_rgb(h: float, s: float, l: float) -> tuple[int, int, int]:
    """Convert HSL (0-360, 0-1, 0-1) to RGB (0-255)."""
    r_n, g_n, b_n = colorsys.hls_to_rgb(h / 360, l, s)
    return (int(round(r_n * 255)), int(round(g_n * 255)), int(round(b_n * 255)))


def rgb_to_hsv(r: int, g: int, b: int) -> tuple[float, float, float]:
    """Convert RGB (0-255) to HSV (0-360, 0-1, 0-1)."""
    r_n, g_n, b_n = r / 255.0, g / 255.0, b / 255.0
    h, s, v = colorsys.rgb_to_hsv(r_n, g_n, b_n)
    return (h * 360, s, v)


def hsv_to_rgb(h: float, s: float, v: float) -> tuple[int, int, int]:
    """Convert HSV (0-360, 0-1, 0-1) to RGB (0-255)."""
    r_n, g_n, b_n = colorsys.hsv_to_rgb(h / 360, s, v)
    return (int(round(r_n * 255)), int(round(g_n * 255)), int(round(b_n * 255)))


def rgb_to_cmyk(r: int, g: int, b: int) -> tuple[float, float, float, float]:
    """Convert RGB (0-255) to CMYK (0-1, 0-1, 0-1, 0-1)."""
    if r == 0 and g == 0 and b == 0:
        return (0.0, 0.0, 0.0, 1.0)
    r_n, g_n, b_n = r / 255.0, g / 255.0, b / 255.0
    k = 1.0 - max(r_n, g_n, b_n)
    c = (1.0 - r_n - k) / (1.0 - k)
    m = (1.0 - g_n - k) / (1.0 - k)
    y = (1.0 - b_n - k) / (1.0 - k)
    return (c, m, y, k)


def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> tuple[int, int, int]:
    """Convert CMYK (0-1) to RGB (0-255)."""
    r = int(round(255 * (1.0 - c) * (1.0 - k)))
    g = int(round(255 * (1.0 - m) * (1.0 - k)))
    b = int(round(255 * (1.0 - y) * (1.0 - k)))
    return (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))


# ---------------------------------------------------------------------------
# Color manipulation
# ---------------------------------------------------------------------------

def lighten(hex_color: str, amount: float = 0.2) -> str:
    """Lighten a color by mixing it with white.

    Parameters
    ----------
    amount : float
        0.0 = no change, 1.0 = pure white.
    """
    r, g, b = hex_to_rgb(hex_color)
    r = int(round(r + (255 - r) * amount))
    g = int(round(g + (255 - g) * amount))
    b = int(round(b + (255 - b) * amount))
    return rgb_to_hex(r, g, b)


def darken(hex_color: str, amount: float = 0.2) -> str:
    """Darken a color by mixing it with black.

    Parameters
    ----------
    amount : float
        0.0 = no change, 1.0 = pure black.
    """
    r, g, b = hex_to_rgb(hex_color)
    r = int(round(r * (1 - amount)))
    g = int(round(g * (1 - amount)))
    b = int(round(b * (1 - amount)))
    return rgb_to_hex(r, g, b)


def blend_colors(color1: str, color2: str, ratio: float = 0.5) -> str:
    """Blend two colors together.

    Parameters
    ----------
    ratio : float
        Blend ratio. 0.0 = color1, 1.0 = color2, 0.5 = equal mix.
    """
    r1, g1, b1 = hex_to_rgb(color1)
    r2, g2, b2 = hex_to_rgb(color2)
    r = int(round(r1 * (1 - ratio) + r2 * ratio))
    g = int(round(g1 * (1 - ratio) + g2 * ratio))
    b = int(round(b1 * (1 - ratio) + b2 * ratio))
    return rgb_to_hex(r, g, b)


def contrast_color(hex_color: str, threshold: float = 0.5) -> str:
    """Return black or white for maximum contrast against the given color.

    Uses the WCAG relative luminance formula.
    """
    r, g, b = hex_to_rgb(hex_color)
    # Relative luminance
    def linearize(c):
        c_n = c / 255.0
        return c_n / 12.92 if c_n <= 0.03928 else ((c_n + 0.055) / 1.055) ** 2.4

    luminance = 0.2126 * linearize(r) + 0.7152 * linearize(g) + 0.0722 * linearize(b)
    return "#000000" if luminance > threshold else "#FFFFFF"


# ---------------------------------------------------------------------------
# Color harmony generators
# ---------------------------------------------------------------------------

def complementary_colors(hex_color: str) -> list[str]:
    """Return the complementary color pair (original + opposite)."""
    h, s, l = rgb_to_hsl(*hex_to_rgb(hex_color))
    comp = hsl_to_rgb((h + 180) % 360, s, l)
    return [hex_color, rgb_to_hex(*comp)]


def analogous_colors(hex_color: str, angle: float = 30) -> list[str]:
    """Return 3 analogous colors (original ± angle on the color wheel)."""
    h, s, l = rgb_to_hsl(*hex_to_rgb(hex_color))
    return [
        rgb_to_hex(*hsl_to_rgb((h - angle) % 360, s, l)),
        hex_color,
        rgb_to_hex(*hsl_to_rgb((h + angle) % 360, s, l)),
    ]


def triadic_colors(hex_color: str) -> list[str]:
    """Return 3 triadic colors (120° apart on the color wheel)."""
    h, s, l = rgb_to_hsl(*hex_to_rgb(hex_color))
    return [
        hex_color,
        rgb_to_hex(*hsl_to_rgb((h + 120) % 360, s, l)),
        rgb_to_hex(*hsl_to_rgb((h + 240) % 360, s, l)),
    ]


def split_complementary_colors(hex_color: str, angle: float = 30) -> list[str]:
    """Return 3 split-complementary colors."""
    h, s, l = rgb_to_hsl(*hex_to_rgb(hex_color))
    return [
        hex_color,
        rgb_to_hex(*hsl_to_rgb((h + 180 - angle) % 360, s, l)),
        rgb_to_hex(*hsl_to_rgb((h + 180 + angle) % 360, s, l)),
    ]


def tetradic_colors(hex_color: str) -> list[str]:
    """Return 4 tetradic (rectangle) colors (90° apart)."""
    h, s, l = rgb_to_hsl(*hex_to_rgb(hex_color))
    return [
        hex_color,
        rgb_to_hex(*hsl_to_rgb((h + 90) % 360, s, l)),
        rgb_to_hex(*hsl_to_rgb((h + 180) % 360, s, l)),
        rgb_to_hex(*hsl_to_rgb((h + 270) % 360, s, l)),
    ]


def monochromatic_colors(hex_color: str, count: int = 5) -> list[str]:
    """Generate monochromatic variations by adjusting lightness."""
    h, s, l = rgb_to_hsl(*hex_to_rgb(hex_color))
    colors = []
    for i in range(count):
        new_l = max(0.05, min(0.95, l - 0.3 + (0.6 * i / max(1, count - 1))))
        colors.append(rgb_to_hex(*hsl_to_rgb(h, s, new_l)))
    return colors


# ---------------------------------------------------------------------------
# Palette extraction
# ---------------------------------------------------------------------------

def extract_palette(prs_or_path, *, top_n: int = 10) -> list[ColorInfo]:
    """Extract the dominant color palette from a presentation.

    Scans all text colors, shape fills, line colors, and backgrounds.

    Parameters
    ----------
    top_n : int
        Return the top N most-used colors.

    Returns
    -------
    list[ColorInfo]
        Sorted by frequency (most common first).
    """
    color_counts: dict[str, dict[str, int]] = {}

    def _record(hex_val: str, context: str):
        if not hex_val or len(hex_val) < 6:
            return
        hex_val = hex_val.upper()
        if not hex_val.startswith("#"):
            hex_val = f"#{hex_val}"
        if hex_val not in color_counts:
            color_counts[hex_val] = {"text": 0, "fill": 0, "line": 0, "background": 0}
        color_counts[hex_val][context] = color_counts[hex_val].get(context, 0) + 1

    prs = _open_prs(prs_or_path)
    try:
        for slide in prs.slides:
            # Background
            try:
                bg = slide.background
                fill = bg.fill
                if fill.type is not None:
                    try:
                        rgb = fill.fore_color.rgb
                        _record(f"#{rgb}", "background")
                    except Exception:
                        pass
            except Exception:
                pass

            for shape in slide.shapes:
                # Shape fill
                try:
                    fill = shape.fill
                    if fill.type is not None:
                        try:
                            rgb = fill.fore_color.rgb
                            _record(f"#{rgb}", "fill")
                        except Exception:
                            pass
                except Exception:
                    pass

                # Line color
                try:
                    line = shape.line
                    if line.fill.type is not None:
                        try:
                            rgb = line.color.rgb
                            _record(f"#{rgb}", "line")
                        except Exception:
                            pass
                except Exception:
                    pass

                # Text colors
                try:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                try:
                                    rgb = run.font.color.rgb
                                    _record(f"#{rgb}", "text")
                                except Exception:
                                    pass
                except Exception:
                    pass
    finally:
        pass

    # Sort by total count
    results = []
    for hex_val, contexts in color_counts.items():
        total = sum(contexts.values())
        dominant_context = max(contexts, key=contexts.get)
        r, g, b = hex_to_rgb(hex_val)
        results.append(ColorInfo(
            hex_value=hex_val,
            count=total,
            context=dominant_context,
            hsl=rgb_to_hsl(r, g, b),
        ))

    results.sort(key=lambda c: c.count, reverse=True)
    return results[:top_n]


# ---------------------------------------------------------------------------
# Theme color resolution
# ---------------------------------------------------------------------------

# PowerPoint theme color indices → typical names
_THEME_COLOR_NAMES = {
    "dk1": "Dark 1",
    "lt1": "Light 1",
    "dk2": "Dark 2",
    "lt2": "Light 2",
    "accent1": "Accent 1",
    "accent2": "Accent 2",
    "accent3": "Accent 3",
    "accent4": "Accent 4",
    "accent5": "Accent 5",
    "accent6": "Accent 6",
    "hlink": "Hyperlink",
    "folHlink": "Followed Hyperlink",
}


def resolve_theme_color(prs_or_path, theme_color: str) -> str | None:
    """Resolve a theme color reference to a concrete hex value.

    Parameters
    ----------
    theme_color : str
        Theme color name: ``"dk1"``, ``"lt1"``, ``"dk2"``, ``"lt2"``,
        ``"accent1"``–``"accent6"``, ``"hlink"``, ``"folHlink"``.

    Returns
    -------
    str or None
        Hex color string (e.g. ``"#4472C4"``), or None if not resolvable.
    """
    import zipfile

    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        with zipfile.ZipFile(path, "r") as zf:
            # Find theme file
            theme_path = None
            for name in zf.namelist():
                if name.startswith("ppt/theme/") and name.endswith(".xml"):
                    theme_path = name
                    break

            if theme_path is None:
                return None

            from xml.etree import ElementTree as ET
            theme_xml = zf.read(theme_path)
            root = ET.fromstring(theme_xml)

            # Navigate to themeElements/clrScheme
            ns_a = _NS_A
            for elem in root.iter(f"{{{ns_a}}}clrScheme"):
                for color_elem in elem:
                    tag = color_elem.tag.split("}")[-1] if "}" in color_elem.tag else color_elem.tag
                    if tag == theme_color:
                        # Find the actual color value
                        for child in color_elem:
                            child_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                            if child_tag == "srgbClr":
                                return f"#{child.get('val', '')}"
                            elif child_tag == "sysClr":
                                # System color — return lastClr if available
                                last_clr = child.get("lastClr")
                                if last_clr:
                                    return f"#{last_clr}"
                break

        return None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Gradient builder
# ---------------------------------------------------------------------------

def apply_gradient(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    gradient_type: str = "linear",
    angle: float = 0.0,
    stops: list[GradientStop] | None = None,
    colors: list[str] | None = None,
) -> bool:
    """Apply a gradient fill to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    gradient_type : str
        ``"linear"``, ``"radial"``, or ``"path"``.
    angle : float
        Angle in degrees (for linear gradients).
    stops : list[GradientStop], optional
        Explicit gradient stops with position, color, and alpha.
    colors : list[str], optional
        Simplified: just provide 2+ hex colors for equal-spaced stops.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        spPr = shape._element.find(f".//{{{_NS_A}}}spPr")
        if spPr is None:
            return False

        # Remove existing fill
        for fill in list(spPr):
            if fill.tag in (
                f"{{{_NS_A}}}solidFill", f"{{{_NS_A}}}noFill",
                f"{{{_NS_A}}}gradFill", f"{{{_NS_A}}}pattFill",
                f"{{{_NS_A}}}blipFill",
            ):
                spPr.remove(fill)

        # Build gradient stops
        if stops is None:
            if colors is None or len(colors) < 2:
                return False
            stops = []
            n = len(colors)
            for i, c in enumerate(colors):
                stops.append(GradientStop(
                    position=i / (n - 1),
                    color=c,
                    alpha=100,
                ))

        # Create gradFill
        gradFill = etree.SubElement(spPr, f"{{{_NS_A}}}gradFill")

        if gradient_type == "linear":
            lin = etree.SubElement(gradFill, f"{{{_NS_A}}}lin")
            lin.set("ang", str(int(angle * 60000)))  # EMU angle (1° = 60000)
            lin.set("scaled", "1")
        elif gradient_type == "radial":
            path = etree.SubElement(gradFill, f"{{{_NS_A}}}path")
            path.set("path", "circle")
            fillToRect = etree.SubElement(path, f"{{{_NS_A}}}fillToRect")
            fillToRect.set("l", "50000")
            fillToRect.set("t", "50000")
            fillToRect.set("r", "50000")
            fillToRect.set("b", "50000")
        elif gradient_type == "path":
            path = etree.SubElement(gradFill, f"{{{_NS_A}}}path")
            path.set("path", "shape")

        # Gradient stops
        gsLst = etree.SubElement(gradFill, f"{{{_NS_A}}}gsLst")
        for stop in stops:
            gs = etree.SubElement(gsLst, f"{{{_NS_A}}}gs")
            gs.set("pos", str(int(stop.position * 100000)))  # 0-100000

            solidFill = etree.SubElement(gs, f"{{{_NS_A}}}solidFill")
            hex_val = stop.color.lstrip("#")
            srgbClr = etree.SubElement(solidFill, f"{{{_NS_A}}}srgbClr")
            srgbClr.set("val", hex_val)

            if stop.alpha < 100:
                alpha = etree.SubElement(srgbClr, f"{{{_NS_A}}}alpha")
                alpha.set("val", str(int(stop.alpha * 1000)))  # 0-100000

        return True
    finally:
        _save_prs(prs, path)


def remove_gradient(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove gradient fill from a shape (reverts to no fill).

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        spPr = shape._element.find(f".//{{{_NS_A}}}spPr")
        if spPr is None:
            return False

        gradFill = spPr.find(f"{{{_NS_A}}}gradFill")
        if gradFill is not None:
            spPr.remove(gradFill)
            # Add noFill
            etree.SubElement(spPr, f"{{{_NS_A}}}noFill")
            return True
        return False
    finally:
        _save_prs(prs, path)


def list_gradients(prs_or_path, slide_index: int | None = None) -> list[GradientInfo]:
    """List all gradient fills in the presentation.

    Parameters
    ----------
    slide_index : int, optional
        1-based slide index (1 = first slide).  If provided, only list
        gradients on that slide.
    """
    prs = _open_prs(prs_or_path)
    try:
        results = []
        if slide_index is not None:
            if slide_index < 1 or slide_index > len(prs.slides):
                raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
            slides = [prs.slides[slide_index - 1]]
        else:
            slides = prs.slides

        for slide in slides:
            for shape in slide.shapes:
                try:
                    spPr = shape._element.find(f".//{{{_NS_A}}}spPr")
                    if spPr is None:
                        continue

                    gradFill = spPr.find(f"{{{_NS_A}}}gradFill")
                    if gradFill is None:
                        continue

                    info = GradientInfo()

                    # Type
                    lin = gradFill.find(f"{{{_NS_A}}}lin")
                    path_elem = gradFill.find(f"{{{_NS_A}}}path")

                    if lin is not None:
                        info.gradient_type = "linear"
                        ang = lin.get("ang", "0")
                        info.angle = int(ang) / 60000.0
                    elif path_elem is not None:
                        path_type = path_elem.get("path", "shape")
                        info.gradient_type = "radial" if path_type == "circle" else "path"

                    # Stops
                    gsLst = gradFill.find(f"{{{_NS_A}}}gsLst")
                    if gsLst is not None:
                        for gs in gsLst.findall(f"{{{_NS_A}}}gs"):
                            pos = int(gs.get("pos", "0")) / 100000.0
                            srgb = gs.find(f".//{{{_NS_A}}}srgbClr")
                            color = f"#{srgb.get('val', '000000')}" if srgb is not None else "#000000"
                            alpha_elem = srgb.find(f"{{{_NS_A}}}alpha") if srgb is not None else None
                            alpha = int(alpha_elem.get("val", "100000")) / 1000.0 if alpha_elem is not None else 100
                            info.stops.append(GradientStop(position=pos, color=color, alpha=alpha))

                    results.append(info)
                except Exception:
                    pass

        return results
    finally:
        pass


# ---------------------------------------------------------------------------
# Batch recoloring
# ---------------------------------------------------------------------------

def recolor_presentation(
    prs_or_path,
    *,
    color_map: dict[str, str] | None = None,
    theme_key: str | None = None,
) -> int:
    """Recolor an entire presentation using a color mapping.

    Parameters
    ----------
    color_map : dict[str, str], optional
        Map of old hex → new hex (e.g. ``{"#FF0000": "#0000FF"}``).
    theme_key : str, optional
        If provided, auto-extract palette and map to the given theme's colors.

    Returns
    -------
    int
        Number of color replacements made.
    """
    if color_map is None and theme_key is None:
        raise ValueError("Provide either color_map or theme_key")

    if theme_key is not None:
        # Auto-generate color map from theme
        color_map = _build_theme_color_map(prs_or_path, theme_key)
        if not color_map:
            return 0

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        count = 0
        # Normalize map keys
        normalized_map = {}
        for old, new in color_map.items():
            normalized_map[old.upper().lstrip("#")] = new.upper().lstrip("#")
            normalized_map[old.lower().lstrip("#")] = new.upper().lstrip("#")

        for slide in prs.slides:
            # Background
            try:
                bg = slide.background
                fill = bg.fill
                if fill.type is not None:
                    try:
                        rgb = str(fill.fore_color.rgb).upper()
                        if rgb in normalized_map:
                            fill.fore_color.rgb = normalized_map[rgb]
                            count += 1
                    except Exception:
                        pass
            except Exception:
                pass

            for shape in slide.shapes:
                # Shape fill
                try:
                    fill = shape.fill
                    if fill.type is not None:
                        try:
                            rgb = str(fill.fore_color.rgb).upper()
                            if rgb in normalized_map:
                                fill.fore_color.rgb = normalized_map[rgb]
                                count += 1
                        except Exception:
                            pass
                except Exception:
                    pass

                # Line color
                try:
                    line = shape.line
                    if line.fill.type is not None:
                        try:
                            rgb = str(line.color.rgb).upper()
                            if rgb in normalized_map:
                                line.color.rgb = normalized_map[rgb]
                                count += 1
                        except Exception:
                            pass
                except Exception:
                    pass

                # Text colors
                try:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                try:
                                    rgb = str(run.font.color.rgb).upper()
                                    if rgb in normalized_map:
                                        run.font.color.rgb = normalized_map[rgb]
                                        count += 1
                                except Exception:
                                    pass
                except Exception:
                    pass

        return count
    finally:
        _save_prs(prs, path)


def swap_colors(
    prs_or_path,
    color1: str,
    color2: str,
) -> int:
    """Swap two colors throughout a presentation.

    All occurrences of color1 become color2 and vice versa.
    """
    # Use a temporary intermediate color to avoid double-swapping
    temp = "#FEFEFE"  # Very unlikely to be used
    c1 = color1.upper().lstrip("#")
    c2 = color2.upper().lstrip("#")

    # First pass: color1 → temp
    count1 = recolor_presentation(prs_or_path, color_map={f"#{c1}": temp})
    # Second pass: color2 → color1
    count2 = recolor_presentation(prs_or_path, color_map={f"#{c2}": f"#{c1}"})
    # Third pass: temp → color2
    count3 = recolor_presentation(prs_or_path, color_map={temp: f"#{c2}"})

    return count1 + count2 + count3


def _build_theme_color_map(prs_or_path, theme_key: str) -> dict[str, str]:
    """Build a color mapping from the current palette to a target theme."""
    try:
        # Import theme colors from pptx_helper
        import importlib
        pptx_helper = importlib.import_module("pptx_helper")
        themes = getattr(pptx_helper, "THEMES", {})
        if theme_key not in themes:
            return {}

        target_theme = themes[theme_key]
        target_colors = set()
        for key in ("bg", "primary", "secondary", "accent", "text", "muted"):
            if key in target_theme:
                val = target_theme[key]
                if isinstance(val, str) and val.startswith("#"):
                    target_colors.add(val.upper().lstrip("#"))

        # Extract current palette
        palette = extract_palette(prs_or_path, top_n=20)
        if not palette:
            return {}

        # Simple mapping: map each palette color to the closest target color
        color_map = {}
        for info in palette:
            old_hex = info.hex_value.upper().lstrip("#")
            if old_hex in target_colors:
                continue  # Already a target color
            # Find closest target color
            best = None
            best_dist = float("inf")
            r1, g1, b1 = hex_to_rgb(f"#{old_hex}")
            for tc in target_colors:
                r2, g2, b2 = hex_to_rgb(f"#{tc}")
                dist = (r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2
                if dist < best_dist:
                    best_dist = dist
                    best = tc
            if best is not None and best_dist < 50000:  # Reasonable threshold
                color_map[f"#{old_hex}"] = f"#{best}"

        return color_map
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path):
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path.
    Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs, path):
    """Save *prs* back to *path* if *path* is not None."""
    if path is not None:
        prs.save(str(path))


def _find_shape(slide, shape_name: str):
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None
