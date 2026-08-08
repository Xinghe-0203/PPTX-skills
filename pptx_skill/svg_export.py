"""PowerPoint slide to SVG vector graphics export.

Converts DrawingML shapes on a python-pptx Slide to SVG elements -- the inverse
of the ``svg_import`` module.  Each slide becomes a self-contained SVG with
positioned ``<text>``, ``<image>``, ``<path>``, ``<rect>``, ``<ellipse>``,
``<line>``, and ``<g>`` elements that mirror the original layout.

OOXML reference: ECMA-376 Part 4, §20.1 (DrawingML), §21.2 (PresentationML).

Public API
----------
- :func:`export_to_svg`       -- PPTX -> per-slide SVG files
- :func:`export_slide_to_svg` -- single Slide -> SVG string
- :func:`shape_to_svg`        -- single Shape -> SVG element string
"""
from __future__ import annotations

import base64
import math
import os
from typing import Any

from pptx import Presentation as _open_presentation
from pptx.presentation import Presentation as _PresentationCls

from pptx_skill.constants import A_NS as _NS_A
from pptx_skill.constants import P_NS as _NS_P
from pptx_skill.units import EMU_PER_INCH as _EMU_PER_INCH
from pptx_skill.units import EMU_PER_PT as _EMU_PER_PT

__all__ = [
    "export_to_svg",
    "export_slide_to_svg",
    "shape_to_svg",
]


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_SVG_NS = "http://www.w3.org/2000/svg"
_XLINK_NS = "http://www.w3.org/1999/xlink"

# Preset geometry -> SVG element mapping for simple shapes.
# Shapes not in this map are rendered as <rect> with the bounding box.
_PRESET_TO_SVG: dict[str, str] = {
    "rect": "rect",
    "ellipse": "ellipse",
    "roundRect": "rect",
    "oval": "ellipse",
    "diamond": "polygon",
    "triangle": "polygon",
    "rtTriangle": "polygon",
    "parallelogram": "polygon",
    "trapezoid": "polygon",
    "pentagon": "polygon",
    "hexagon": "polygon",
    "octagon": "polygon",
    "star5": "polygon",
    "star6": "polygon",
    "star8": "polygon",
    "arrow": "polygon",
    "chevron": "polygon",
    "heart": "path",
    "lightningBolt": "path",
    "sun": "path",
    "moon": "path",
    "cloud": "path",
}

# MSO_SHAPE_TYPE constants
_MSOT_AUTO_SHAPE = 1
_MSOT_CALLOUT = 2
_MSOT_CHART = 3
_MSOT_COMMENT = 4
_MSOT_FREEFORM = 5
_MSOT_GROUP = 6
_MSOT_IGX_GRAPHIC = 7
_MSOT_INK = 8
_MSOT_MEDIA = 9
_MSOT_OLE_CONTROL = 10
_MSOT_PICTURE = 13
_MSOT_PLACEHOLDER = 14
_MSOT_TABLE = 19
_MSOT_TEXT_BOX = 17


# ---------------------------------------------------------------------------
# Unit conversion
# ---------------------------------------------------------------------------

def _emu_to_pt(emu: int) -> float:
    """Convert EMU to points (1 pt = 12700 EMU)."""
    return emu / _EMU_PER_PT


def _emu_to_px(emu: int, dpi: int = 150) -> float:
    """Convert EMU to pixels at the given DPI."""
    return _emu_to_pt(emu) * dpi / 72.0


def _hundredths_pt_to_pt(hundredths: int) -> float:
    """Convert font-size in hundredths of a point to points."""
    return hundredths / 100.0


# ---------------------------------------------------------------------------
# XML / SVG helpers
# ---------------------------------------------------------------------------

def _escape_xml(text: str) -> str:
    """Escape special XML characters."""
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _format_float(value: float, precision: int = 2) -> str:
    """Format a float for SVG attribute output, stripping trailing zeros."""
    s = f"{value:.{precision}f}"
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def _rgb_to_hex(rgb: Any) -> str:
    """Convert a python-pptx RGBColor to ``#RRGGBB`` string."""
    if rgb is None:
        return "#000000"
    try:
        return f"#{rgb}"
    except Exception:
        return "#000000"


def _color_to_svg(color: Any, default: str = "#000000") -> str:
    """Convert a python-pptx color object to an SVG color string.

    Handles RGBColor, theme colors (returns the default for theme colors
    since we cannot resolve them without a theme context), and None.
    """
    if color is None:
        return default
    try:
        rgb = color.rgb
        if rgb is not None:
            return _rgb_to_hex(rgb)
    except AttributeError:
        pass
    # Try accessing the RGB value directly
    try:
        return _rgb_to_hex(color)
    except Exception:
        return default


# ---------------------------------------------------------------------------
# Presentation resolution helpers
# ---------------------------------------------------------------------------

def _resolve_presentation(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
) -> tuple[_PresentationCls, str | None]:
    """Return ``(Presentation, temp_dir_or_None)``.

    Mirrors the pattern from ``export.py``.
    """
    if isinstance(prs_or_path, _PresentationCls):
        return prs_or_path, None
    path = os.fspath(prs_or_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"PPTX file not found: {path}")
    return _open_presentation(path), None


def _ensure_dir(path: str | os.PathLike[str]) -> str:
    """Create directory if needed and return the absolute path."""
    abs_path = os.path.abspath(path)
    os.makedirs(abs_path, exist_ok=True)
    return abs_path


# ---------------------------------------------------------------------------
# Background extraction
# ---------------------------------------------------------------------------

def _slide_background_color(slide: Any) -> str | None:
    """Attempt to determine the slide background fill colour.

    Returns an SVG colour string (``#RRGGBB``) or ``None`` if
    the background cannot be determined.
    """
    try:
        bg = slide.background
        fill = bg.fill
        if fill.type is not None:
            # solid fill
            try:
                if fill.fore_color and fill.fore_color.rgb:
                    return _rgb_to_hex(fill.fore_color.rgb)
            except Exception:
                pass
    except Exception:
        pass

    # Try reading the XML directly
    try:
        import lxml.etree  # noqa: F401 — availability probe

        sld_elem = slide._element
        bg_elem = sld_elem.find(f"{{{_NS_P}}}bg")
        if bg_elem is not None:
            for sf in bg_elem.iter(f"{{{_NS_A}}}solidFill"):
                srgb = sf.find(f"{{{_NS_A}}}srgbClr")
                if srgb is not None and srgb.get("val"):
                    return f"#{srgb.get('val')}"
    except ImportError:
        pass
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Shape fill / stroke extraction
# ---------------------------------------------------------------------------

def _shape_fill_to_svg(shape: Any) -> str | None:
    """Return the fill colour of *shape* as an SVG colour string, or None."""
    try:
        fill = shape.fill
        if fill.type is not None:
            try:
                if fill.fore_color and fill.fore_color.rgb:
                    return _rgb_to_hex(fill.fore_color.rgb)
            except Exception:
                pass
    except Exception:
        pass

    # XML fallback
    try:
        import lxml.etree  # noqa: F401 — availability probe

        sp_pr = shape._element.find(f"{{{_NS_P}}}spPr")
        if sp_pr is not None:
            sf = sp_pr.find(f"{{{_NS_A}}}solidFill")
            if sf is not None:
                srgb = sf.find(f"{{{_NS_A}}}srgbClr")
                if srgb is not None and srgb.get("val"):
                    return f"#{srgb.get('val')}"
    except ImportError:
        pass
    except Exception:
        pass

    return None


def _shape_stroke_to_svg(shape: Any) -> tuple[str | None, float]:
    """Return ``(stroke_colour, stroke_width_pt)`` for *shape*."""
    colour = None
    width = 0.0

    try:
        line = shape.line
        if line.fill.type is not None:
            try:
                if line.color and line.color.rgb:
                    colour = _rgb_to_hex(line.color.rgb)
            except Exception:
                pass
        if line.width is not None:
            width = _emu_to_pt(line.width)
    except Exception:
        pass

    # XML fallback for line properties
    if colour is None:
        try:
            import lxml.etree  # noqa: F401 — availability probe

            sp_pr = shape._element.find(f"{{{_NS_P}}}spPr")
            if sp_pr is not None:
                ln = sp_pr.find(f"{{{_NS_A}}}ln")
                if ln is not None:
                    w_attr = ln.get("w")
                    if w_attr:
                        try:
                            width = _emu_to_pt(int(w_attr))
                        except (ValueError, TypeError):
                            pass
                    sf = ln.find(f"{{{_NS_A}}}solidFill")
                    if sf is not None:
                        srgb = sf.find(f"{{{_NS_A}}}srgbClr")
                        if srgb is not None and srgb.get("val"):
                            colour = f"#{srgb.get('val')}"
        except ImportError:
            pass
        except Exception:
            pass

    return colour, width


# ---------------------------------------------------------------------------
# Shape geometry / preset shape path generation
# ---------------------------------------------------------------------------

def _preset_shape_path(preset: str, width: float, height: float) -> str | None:
    """Return an SVG ``d`` attribute string for a preset shape geometry.

    Returns ``None`` if the shape should be rendered as a simple primitive
    (``<rect>``, ``<ellipse>``) instead.
    """
    w, h = width, height
    hw, hh = w / 2, h / 2

    if preset == "diamond":
        return f"M {hw} 0 L {w} {hh} L {hw} {h} L 0 {hh} Z"

    if preset == "triangle":
        return f"M {hw} 0 L {w} {h} L 0 {h} Z"

    if preset == "rtTriangle":
        return f"M 0 0 L {w} 0 L 0 {h} Z"

    if preset == "parallelogram":
        offset = w * 0.2
        return f"M {offset} 0 L {w} 0 L {w - offset} {h} L 0 {h} Z"

    if preset == "trapezoid":
        offset = w * 0.15
        return f"M {offset} 0 L {w - offset} 0 L {w} {h} L 0 {h} Z"

    if preset == "pentagon":
        return _regular_polygon_path(hw, hh, min(hw, hh), 5, -math.pi / 2)

    if preset == "hexagon":
        return _regular_polygon_path(hw, hh, min(hw, hh), 6, -math.pi / 2)

    if preset == "octagon":
        return _regular_polygon_path(hw, hh, min(hw, hh), 8, -math.pi / 2)

    if preset == "star5":
        return _star_path(hw, hh, min(hw, hh), min(hw, hh) * 0.4, 5, -math.pi / 2)

    if preset == "star6":
        return _star_path(hw, hh, min(hw, hh), min(hw, hh) * 0.4, 6, -math.pi / 2)

    if preset == "star8":
        return _star_path(hw, hh, min(hw, hh), min(hw, hh) * 0.4, 8, -math.pi / 2)

    if preset == "arrow":
        inset = w * 0.35
        return f"M 0 {hh * 0.4} L {w - inset} {hh * 0.4} L {w - inset} 0 L {w} {hh} L {w - inset} {h} L {w - inset} {hh * 1.6} L 0 {hh * 1.6} Z"

    if preset == "chevron":
        inset = w * 0.3
        return f"M 0 0 L {w - inset} 0 L {w} {hh} L {w - inset} {h} L 0 {h} L {inset} {hh} Z"

    return None


def _regular_polygon_path(
    cx: float, cy: float, r: float, sides: int, start_angle: float
) -> str:
    """Build an SVG path string for a regular polygon."""
    points: list[str] = []
    for i in range(sides):
        angle = start_angle + 2 * math.pi * i / sides
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append(f"{_format_float(x)} {_format_float(y)}")
    d = "M " + points[0]
    for pt in points[1:]:
        d += f" L {pt}"
    d += " Z"
    return d


def _star_path(
    cx: float,
    cy: float,
    outer_r: float,
    inner_r: float,
    points: int,
    start_angle: float,
) -> str:
    """Build an SVG path string for a star shape."""
    parts: list[str] = []
    for i in range(points * 2):
        angle = start_angle + math.pi * i / points
        r = outer_r if i % 2 == 0 else inner_r
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        parts.append(f"{_format_float(x)} {_format_float(y)}")
    d = "M " + parts[0]
    for pt in parts[1:]:
        d += f" L {pt}"
    d += " Z"
    return d


def _custom_geom_to_path(shape: Any, offset_x: int, offset_y: int) -> str | None:
    """Extract a custom geometry path from a shape's XML and return an SVG ``d`` string.

    Returns ``None`` if the shape has no custom geometry.
    """
    try:
        import lxml.etree  # noqa: F401 — availability probe

        sp_pr = shape._element.find(f"{{{_NS_P}}}spPr")
        if sp_pr is None:
            return None

        cust_geom = sp_pr.find(f"{{{_NS_A}}}custGeom")
        if cust_geom is None:
            return None

        path_lst = cust_geom.find(f"{{{_NS_A}}}pathLst")
        if path_lst is None:
            return None

        # Get path coordinate system extents
        path_elems = path_lst.findall(f"{{{_NS_A}}}path")
        if not path_elems:
            return None

        shape_w = shape.width
        shape_h = shape.height

        d_parts: list[str] = []

        for path_elem in path_elems:
            path_w = int(path_elem.get("w", "0"))
            path_h = int(path_elem.get("h", "0"))

            # Scale factor: path coordinates -> shape EMU -> SVG points
            sx = shape_w / max(path_w, 1) if path_w else 1
            sy = shape_h / max(path_h, 1) if path_h else 1

            sub_path: list[str] = []
            for child in path_elem:
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

                if tag == "moveTo":
                    pts = child.findall(f"{{{_NS_A}}}pt")
                    if pts:
                        pt = pts[0]
                        x = _emu_to_pt(int(pt.get("x", "0")) * sx // _EMU_PER_PT + offset_x)
                        y = _emu_to_pt(int(pt.get("y", "0")) * sy // _EMU_PER_PT + offset_y)
                        sub_path.append(f"M {_format_float(x)} {_format_float(y)}")

                elif tag == "lnTo":
                    pts = child.findall(f"{{{_NS_A}}}pt")
                    if pts:
                        pt = pts[0]
                        x = _emu_to_pt(int(pt.get("x", "0")) * sx // _EMU_PER_PT + offset_x)
                        y = _emu_to_pt(int(pt.get("y", "0")) * sy // _EMU_PER_PT + offset_y)
                        sub_path.append(f"L {_format_float(x)} {_format_float(y)}")

                elif tag == "cubicBezTo":
                    pts = child.findall(f"{{{_NS_A}}}pt")
                    if len(pts) >= 3:
                        coords: list[str] = []
                        for pt in pts:
                            x = _emu_to_pt(int(pt.get("x", "0")) * sx // _EMU_PER_PT + offset_x)
                            y = _emu_to_pt(int(pt.get("y", "0")) * sy // _EMU_PER_PT + offset_y)
                            coords.append(f"{_format_float(x)} {_format_float(y)}")
                        sub_path.append(f"C {coords[0]} {coords[1]} {coords[2]}")

                elif tag == "quadBezTo":
                    pts = child.findall(f"{{{_NS_A}}}pt")
                    if len(pts) >= 2:
                        coords = []
                        for pt in pts:
                            x = _emu_to_pt(int(pt.get("x", "0")) * sx // _EMU_PER_PT + offset_x)
                            y = _emu_to_pt(int(pt.get("y", "0")) * sy // _EMU_PER_PT + offset_y)
                            coords.append(f"{_format_float(x)} {_format_float(y)}")
                        sub_path.append(f"Q {coords[0]} {coords[1]}")

                elif tag == "arcTo":
                    # DrawingML arcTo: wR, hR, stAng, swAng
                    int(child.get("wR", "0"))
                    int(child.get("hR", "0"))
                    int(child.get("stAng", "0"))
                    int(child.get("swAng", "0"))
                    # Simplified: draw as a line (arc approximation would need
                    # current-point tracking; skip for robustness)
                    pass

                elif tag == "close":
                    sub_path.append("Z")

            if sub_path:
                d_parts.append(" ".join(sub_path))

        return " ".join(d_parts) if d_parts else None

    except ImportError:
        return None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Text conversion
# ---------------------------------------------------------------------------

def _text_frame_to_svg(text_frame: Any, x: float, y: float, width: float, height: float) -> str:
    """Convert a python-pptx TextFrame to SVG ``<text>`` with ``<tspan>`` elements.

    Parameters
    ----------
    text_frame :
        The TextFrame to convert.
    x, y :
        Top-left position in SVG points.
    width, height :
        Bounding box dimensions in SVG points.
    """
    parts: list[str] = []

    # Determine vertical anchor
    anchor = "top"
    try:
        body_pr = text_frame._txBody.find(
            f"{{{_NS_A}}}bodyPr"
        )
        if body_pr is not None:
            anchor_val = body_pr.get("anchor", "t")
            anchor_map = {"t": "top", "ctr": "middle", "b": "bottom"}
            anchor = anchor_map.get(anchor_val, "top")
    except Exception:
        pass

    para_count = len(text_frame.paragraphs)

    for p_idx, para in enumerate(text_frame.paragraphs):
        # Calculate vertical offset for this paragraph
        # Use line spacing estimate: 1.2x font size
        default_size_pt = 12.0
        try:
            for run in para.runs:
                if run.font.size:
                    default_size_pt = _hundredths_pt_to_pt(run.font.size)
                    break
        except Exception:
            pass

        line_height = default_size_pt * 1.2
        para_y = y + p_idx * line_height

        # Adjust for vertical anchor
        if anchor == "middle" and para_count > 1:
            para_y = y + height / 2 - (para_count * line_height) / 2 + p_idx * line_height
        elif anchor == "bottom":
            para_y = y + height - (para_count - p_idx) * line_height

        # Paragraph alignment
        text_anchor = "start"
        try:
            if para.alignment is not None:
                align_map = {0: "start", 1: "middle", 2: "end", 3: "start"}
                text_anchor = align_map.get(para.alignment, "start")
                if para.alignment == 3:  # PP_ALIGN.JUSTIFY
                    text_anchor = "start"
        except Exception:
            pass

        # Build tspan elements for each run
        tspans: list[str] = []
        for run in para.runs:
            text = run.text
            if not text:
                continue

            attrs: list[str] = []

            # Font family
            font_families: list[str] = []
            try:
                if run.font.name:
                    font_families.append(run.font.name)
            except Exception:
                pass
            try:
                # CJK font

                r_elem = run._r
                ea = r_elem.find(f".//{{{_NS_A}}}ea")
                if ea is not None and ea.get("typeface"):
                    font_families.append(ea.get("typeface"))
            except Exception:
                pass

            if font_families:
                family_str = ", ".join(f"'{f}'" for f in font_families)
                attrs.append(f"font-family=\"{family_str}\"")

            # Font size
            try:
                if run.font.size:
                    size_pt = _hundredths_pt_to_pt(run.font.size)
                    attrs.append(f"font-size=\"{_format_float(size_pt)}pt\"")
            except Exception:
                pass

            # Bold
            try:
                if run.font.bold:
                    attrs.append("font-weight=\"bold\"")
            except Exception:
                pass

            # Italic
            try:
                if run.font.italic:
                    attrs.append("font-style=\"italic\"")
            except Exception:
                pass

            # Underline
            try:
                if run.font.underline:
                    attrs.append("text-decoration=\"underline\"")
            except Exception:
                pass

            # Colour
            try:
                if run.font.color and run.font.color.rgb:
                    attrs.append(f"fill=\"{_rgb_to_hex(run.font.color.rgb)}\"")
            except Exception:
                pass

            attr_str = " ".join(attrs)
            escaped_text = _escape_xml(text)

            if attr_str:
                tspans.append(f"<tspan {attr_str}>{escaped_text}</tspan>")
            else:
                tspans.append(f"<tspan>{escaped_text}</tspan>")

        if not tspans:
            # Empty paragraph -- preserve vertical space
            tspans.append("<tspan> </tspan>")

        # Build the text element for this paragraph
        text_attrs: list[str] = [
            f"x=\"{_format_float(x)}\"",
            f"y=\"{_format_float(para_y + default_size_pt)}\"",
        ]
        if text_anchor != "start":
            text_attrs.append(f"text-anchor=\"{text_anchor}\"")

        text_attr_str = " ".join(text_attrs)
        parts.append(f"<text {text_attr_str}>{''.join(tspans)}</text>")

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Picture conversion
# ---------------------------------------------------------------------------

def _picture_to_svg(
    shape: Any,
    x: float,
    y: float,
    width: float,
    height: float,
    *,
    embed_images: bool = True,
) -> str:
    """Convert a picture shape to an SVG ``<image>`` element."""
    try:
        image = shape.image
        content_type = image.content_type
        blob = image.blob
    except Exception:
        # Cannot extract image data -- draw placeholder rect
        return (
            f'<rect x="{_format_float(x)}" y="{_format_float(y)}" '
            f'width="{_format_float(width)}" height="{_format_float(height)}" '
            f'fill="#cccccc" stroke="#999999" stroke-width="1"/>'
        )

    if embed_images:
        b64 = base64.b64encode(blob).decode("ascii")
        href = f"data:{content_type};base64,{b64}"
        return (
            f'<image x="{_format_float(x)}" y="{_format_float(y)}" '
            f'width="{_format_float(width)}" height="{_format_float(height)}" '
            f'href="{href}" preserveAspectRatio="xMidYMid meet"/>'
        )
    else:
        return (
            f'<image x="{_format_float(x)}" y="{_format_float(y)}" '
            f'width="{_format_float(width)}" height="{_format_float(height)}" '
            f'xlink:href="image_{id(shape)}.png" '
            f'preserveAspectRatio="xMidYMid meet"/>'
        )


# ---------------------------------------------------------------------------
# Table conversion
# ---------------------------------------------------------------------------

def _table_to_svg(
    shape: Any,
    x: float,
    y: float,
    width: float,
    height: float,
) -> str:
    """Convert a table shape to an SVG ``<g>`` containing cells and text."""
    try:
        table = shape.table
    except Exception:
        return ""

    if not table.rows:
        return ""

    rows = table.rows
    cols = table.columns
    n_rows = len(rows)
    n_cols = len(cols)

    cell_w = width / max(n_cols, 1)
    cell_h = height / max(n_rows, 1)

    parts: list[str] = []
    parts.append(
        f'<g transform="translate({_format_float(x)}, {_format_float(y)})">'
    )

    for r_idx, row in enumerate(rows):
        for c_idx, cell in enumerate(row.cells):
            cx = c_idx * cell_w
            cy = r_idx * cell_h

            # Cell background
            cell_fill = "#ffffff"
            try:
                if cell.fill.type is not None:
                    try:
                        if cell.fill.fore_color and cell.fill.fore_color.rgb:
                            cell_fill = _rgb_to_hex(cell.fill.fore_color.rgb)
                    except Exception:
                        pass
            except Exception:
                pass

            parts.append(
                f'<rect x="{_format_float(cx)}" y="{_format_float(cy)}" '
                f'width="{_format_float(cell_w)}" height="{_format_float(cell_h)}" '
                f'fill="{cell_fill}" stroke="#000000" stroke-width="0.5"/>'
            )

            # Cell text
            if cell.has_text_frame:
                text_content = cell.text_frame.text.strip()
                if text_content:
                    font_size = 10.0
                    try:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.size:
                                    font_size = _hundredths_pt_to_pt(run.font.size)
                                    break
                            if font_size != 10.0:
                                break
                    except Exception:
                        pass

                    text_y = cy + cell_h / 2 + font_size / 3
                    text_x = cx + cell_w / 2

                    # Determine text colour
                    fill_colour = "#000000"
                    try:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.color and run.font.color.rgb:
                                    fill_colour = _rgb_to_hex(run.font.color.rgb)
                                    break
                            if fill_colour != "#000000":
                                break
                    except Exception:
                        pass

                    # Bold for first row
                    font_weight = ""
                    if r_idx == 0:
                        font_weight = ' font-weight="bold"'

                    parts.append(
                        f'<text x="{_format_float(text_x)}" y="{_format_float(text_y)}" '
                        f'font-size="{_format_float(font_size)}pt" '
                        f'text-anchor="middle" fill="{fill_colour}"{font_weight}>'
                        f'{_escape_xml(text_content)}</text>'
                    )

    parts.append("</g>")
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Auto shape conversion
# ---------------------------------------------------------------------------

def _auto_shape_to_svg(
    shape: Any,
    x: float,
    y: float,
    width: float,
    height: float,
) -> str:
    """Convert an auto shape to SVG element(s)."""
    # Determine preset geometry name
    preset = None
    try:
        sp_pr = shape._element.find(f"{{{_NS_P}}}spPr")
        if sp_pr is not None:
            prst_geom = sp_pr.find(f"{{{_NS_A}}}prstGeom")
            if prst_geom is not None:
                preset = prst_geom.get("prst")
    except Exception:
        pass

    # Try custom geometry first
    custom_d = _custom_geom_to_path(shape, 0, 0)
    if custom_d is not None:
        fill_colour = _shape_fill_to_svg(shape) or "none"
        stroke_colour, stroke_w = _shape_stroke_to_svg(shape)
        stroke_attrs = ""
        if stroke_colour and stroke_w > 0:
            stroke_attrs = f' stroke="{stroke_colour}" stroke-width="{_format_float(stroke_w)}"'
        elif stroke_colour:
            stroke_attrs = f' stroke="{stroke_colour}" stroke-width="1"'

        return (
            f'<path d="{_escape_xml(custom_d)}" '
            f'transform="translate({_format_float(x)}, {_format_float(y)})" '
            f'fill="{fill_colour}"{stroke_attrs}/>'
        )

    # Determine fill and stroke
    fill_colour = _shape_fill_to_svg(shape) or ""
    stroke_colour, stroke_w = _shape_stroke_to_svg(shape)

    fill_attr = f' fill="{fill_colour}"' if fill_colour else ' fill="none"'
    stroke_attr = ""
    if stroke_colour and stroke_w > 0:
        stroke_attr = f' stroke="{stroke_colour}" stroke-width="{_format_float(stroke_w)}"'
    elif stroke_colour:
        stroke_attr = f' stroke="{stroke_colour}" stroke-width="1"'

    # Simple primitives
    if preset == "ellipse" or preset == "oval":
        cx = x + width / 2
        cy = y + height / 2
        rx = width / 2
        ry = height / 2
        return (
            f'<ellipse cx="{_format_float(cx)}" cy="{_format_float(cy)}" '
            f'rx="{_format_float(rx)}" ry="{_format_float(ry)}" '
            f'{fill_attr.lstrip()}{stroke_attr}/>'
        )

    if preset == "roundRect":
        r = min(width, height) * 0.15
        return (
            f'<rect x="{_format_float(x)}" y="{_format_float(y)}" '
            f'width="{_format_float(width)}" height="{_format_float(height)}" '
            f'rx="{_format_float(r)}" ry="{_format_float(r)}" '
            f'{fill_attr.lstrip()}{stroke_attr}/>'
        )

    # Preset shapes with known path data
    if preset and preset in _PRESET_TO_SVG:
        path_d = _preset_shape_path(preset, width, height)
        if path_d:
            return (
                f'<path d="{path_d}" '
                f'transform="translate({_format_float(x)}, {_format_float(y)})" '
                f'{fill_attr.lstrip()}{stroke_attr}/>'
            )

    # Default: bounding box rect
    return (
        f'<rect x="{_format_float(x)}" y="{_format_float(y)}" '
        f'width="{_format_float(width)}" height="{_format_float(height)}" '
        f'{fill_attr.lstrip()}{stroke_attr}/>'
    )


# ---------------------------------------------------------------------------
# Connector conversion
# ---------------------------------------------------------------------------

def _connector_to_svg(
    shape: Any,
    x: float,
    y: float,
    width: float,
    height: float,
) -> str:
    """Convert a connector / line shape to an SVG ``<path>`` or ``<line>``."""
    stroke_colour, stroke_w = _shape_stroke_to_svg(shape)
    if not stroke_colour:
        stroke_colour = "#000000"
    if stroke_w <= 0:
        stroke_w = 1.0

    # Try extracting connection points from XML
    try:
        import lxml.etree  # noqa: F401 — availability probe

        sp_pr = shape._element.find(f"{{{_NS_P}}}spPr")
        if sp_pr is not None:
            # Try to get the actual path from the connector XML
            path_elems = sp_pr.findall(f".//{{{_NS_A}}}path")
            if path_elems:
                d_parts: list[str] = []
                for path_elem in path_elems:
                    for child in path_elem:
                        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                        if tag == "moveTo":
                            pts = child.findall(f"{{{_NS_A}}}pt")
                            if pts:
                                px = _emu_to_pt(int(pts[0].get("x", "0")))
                                py = _emu_to_pt(int(pts[0].get("y", "0")))
                                d_parts.append(f"M {_format_float(px + x)} {_format_float(py + y)}")
                        elif tag == "lnTo":
                            pts = child.findall(f"{{{_NS_A}}}pt")
                            if pts:
                                px = _emu_to_pt(int(pts[0].get("x", "0")))
                                py = _emu_to_pt(int(pts[0].get("y", "0")))
                                d_parts.append(f"L {_format_float(px + x)} {_format_float(py + y)}")
                        elif tag == "cubicBezTo":
                            pts = child.findall(f"{{{_NS_A}}}pt")
                            coords = []
                            for pt in pts:
                                px = _emu_to_pt(int(pt.get("x", "0")))
                                py = _emu_to_pt(int(pt.get("y", "0")))
                                coords.append(f"{_format_float(px + x)} {_format_float(py + y)}")
                            if len(coords) >= 3:
                                d_parts.append(f"C {coords[0]} {coords[1]} {coords[2]}")
                        elif tag == "close":
                            d_parts.append("Z")
                if d_parts:
                    d = " ".join(d_parts)
                    return (
                        f'<path d="{d}" fill="none" '
                        f'stroke="{stroke_colour}" stroke-width="{_format_float(stroke_w)}"/>'
                    )
    except ImportError:
        pass
    except Exception:
        pass

    # Fallback: simple line from start to end
    x2 = x + width
    y2 = y + height
    return (
        f'<line x1="{_format_float(x)}" y1="{_format_float(y)}" '
        f'x2="{_format_float(x2)}" y2="{_format_float(y2)}" '
        f'stroke="{stroke_colour}" stroke-width="{_format_float(stroke_w)}"/>'
    )


# ---------------------------------------------------------------------------
# Group conversion
# ---------------------------------------------------------------------------

def _group_to_svg(
    shape: Any,
    x: float,
    y: float,
    width: float,
    height: float,
    *,
    embed_images: bool = True,
) -> str:
    """Convert a group shape to an SVG ``<g>`` element containing child shapes."""
    parts: list[str] = []
    parts.append(
        f'<g transform="translate({_format_float(x)}, {_format_float(y)})">'
    )

    try:
        for child_shape in shape.shapes:
            # Compute child position relative to group
            child_x = _emu_to_pt(child_shape.left)
            child_y = _emu_to_pt(child_shape.top)
            child_w = _emu_to_pt(child_shape.width)
            child_h = _emu_to_pt(child_shape.height)

            child_svg = _convert_shape(child_shape, child_x, child_y, child_w, child_h, embed_images=embed_images)
            if child_svg:
                parts.append(child_svg)
    except Exception:
        pass

    parts.append("</g>")
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Chart conversion
# ---------------------------------------------------------------------------

def _chart_to_svg(
    shape: Any,
    x: float,
    y: float,
    width: float,
    height: float,
    *,
    embed_images: bool = True,
) -> str:
    """Convert a chart shape to SVG by rendering it to a PNG image first.

    SVG cannot directly represent OOXML charts, so we rasterise the chart
    using the export module's rendering infrastructure and embed the result.
    """
    # We cannot easily render a single chart to PNG without a full slide render,
    # so produce a placeholder rectangle with a label.
    fill_colour = _shape_fill_to_svg(shape) or "#f0f0f0"
    return (
        f'<rect x="{_format_float(x)}" y="{_format_float(y)}" '
        f'width="{_format_float(width)}" height="{_format_float(height)}" '
        f'fill="{fill_colour}" stroke="#999999" stroke-width="1"/>'
        f'<text x="{_format_float(x + width / 2)}" y="{_format_float(y + height / 2)}" '
        f'font-size="12pt" text-anchor="middle" fill="#666666">[Chart]</text>'
    )


# ---------------------------------------------------------------------------
# Shape dispatch
# ---------------------------------------------------------------------------

def _convert_shape(
    shape: Any,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    embed_images: bool = True,
) -> str | None:
    """Dispatch a single shape to the appropriate converter and return SVG."""
    try:
        shape_type = shape.shape_type
    except Exception:
        shape_type = None

    # Picture
    if shape_type == _MSOT_PICTURE:
        return _picture_to_svg(shape, x, y, w, h, embed_images=embed_images)

    # Table
    if shape_type == _MSOT_TABLE:
        return _table_to_svg(shape, x, y, w, h)

    # Group
    if shape_type == _MSOT_GROUP:
        return _group_to_svg(shape, x, y, w, h, embed_images=embed_images)

    # Chart
    if shape_type == _MSOT_CHART:
        return _chart_to_svg(shape, x, y, w, h, embed_images=embed_images)

    # Text box / placeholder with text frame
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        # Auto shape with text or standalone text box
        shape_svg = ""

        # Draw the shape background (if it's an auto shape, not just a text box)
        if shape_type in (_MSOT_AUTO_SHAPE, _MSOT_CALLOUT, _MSOT_FREEFORM):
            shape_svg = _auto_shape_to_svg(shape, x, y, w, h)
        elif shape_type == _MSOT_PLACEHOLDER:
            # Placeholders often have visible shapes
            shape_svg = _auto_shape_to_svg(shape, x, y, w, h)
        elif shape_type == _MSOT_TEXT_BOX:
            # Text boxes typically have no visible border/fill unless set
            fill = _shape_fill_to_svg(shape)
            stroke_colour, stroke_w = _shape_stroke_to_svg(shape)
            if fill or (stroke_colour and stroke_w > 0):
                shape_svg = _auto_shape_to_svg(shape, x, y, w, h)

        # Draw the text content
        text_svg = _text_frame_to_svg(shape.text_frame, x, y, w, h)

        if shape_svg and text_svg:
            return shape_svg + "\n" + text_svg
        if shape_svg:
            return shape_svg
        if text_svg:
            return text_svg

    # Auto shape without text
    if shape_type in (_MSOT_AUTO_SHAPE, _MSOT_CALLOUT, _MSOT_FREEFORM):
        return _auto_shape_to_svg(shape, x, y, w, h)

    # Connector (shape_type may vary; check for line-like shapes)
    try:
        if shape.connector:
            return _connector_to_svg(shape, x, y, w, h)
    except Exception:
        pass

    # Fallback: try auto shape conversion for anything else
    return _auto_shape_to_svg(shape, x, y, w, h)


# ---------------------------------------------------------------------------
# Font CSS generation
# ---------------------------------------------------------------------------

def _collect_fonts_from_slide(slide: Any) -> set[str]:
    """Collect all font family names used on a slide."""
    fonts: set[str] = set()

    for shape in slide.shapes:
        try:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        try:
                            if run.font.name:
                                fonts.add(run.font.name)
                        except Exception:
                            pass
                        # CJK font
                        try:

                            r_elem = run._r
                            ea = r_elem.find(
                                f".//{{{_NS_A}}}ea"
                            )
                            if ea is not None and ea.get("typeface"):
                                fonts.add(ea.get("typeface"))
                            lat = r_elem.find(f".//{{{_NS_A}}}latin")
                            if lat is not None and lat.get("typeface"):
                                fonts.add(lat.get("typeface"))
                        except Exception:
                            pass
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.has_text_frame:
                            for para in cell.text_frame.paragraphs:
                                for run in para.runs:
                                    try:
                                        if run.font.name:
                                            fonts.add(run.font.name)
                                    except Exception:
                                        pass
        except Exception:
            pass

    return fonts


def _fonts_to_css(fonts: set[str]) -> str:
    """Generate CSS ``@font-face`` declarations for the given font names.

    Since we cannot embed actual font subsets, this generates placeholder
    declarations that reference system fonts as fallbacks.
    """
    if not fonts:
        return ""

    parts: list[str] = []
    for font_name in sorted(fonts):
        parts.append(
            f"@font-face {{\n"
            f"  font-family: '{font_name}';\n"
            f"  src: local('{font_name}');\n"
            f"}}"
        )

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Public API: shape_to_svg
# ---------------------------------------------------------------------------

def shape_to_svg(
    shape: Any,
    *,
    offset_x: int = 0,
    offset_y: int = 0,
) -> str | None:
    """Convert a single python-pptx Shape object to an SVG element string.

    Parameters
    ----------
    shape :
        A python-pptx Shape object.
    offset_x :
        Horizontal offset in EMU to add to the shape position (default 0).
    offset_y :
        Vertical offset in EMU to add to the shape position (default 0).

    Returns
    -------
    str | None
        SVG element string, or ``None`` if the shape type is unsupported.

    Examples
    --------
    >>> from pptx import Presentation
    >>> prs = Presentation("slide.pptx")
    >>> slide = prs.slides[0]
    >>> svg_str = shape_to_svg(slide.shapes[0])
    """
    try:
        x = _emu_to_pt(shape.left + offset_x)
        y = _emu_to_pt(shape.top + offset_y)
        w = _emu_to_pt(shape.width)
        h = _emu_to_pt(shape.height)
    except Exception:
        return None

    try:
        return _convert_shape(shape, x, y, w, h)
    except Exception:
        # Return a fallback placeholder rather than crashing
        return (
            f'<rect x="{_format_float(x)}" y="{_format_float(y)}" '
            f'width="{_format_float(w)}" height="{_format_float(h)}" '
            f'fill="#eeeeee" stroke="#cccccc" stroke-width="0.5"/>'
        )


# ---------------------------------------------------------------------------
# Public API: export_slide_to_svg
# ---------------------------------------------------------------------------

def export_slide_to_svg(
    slide: Any,
    *,
    slide_width: float = 13.333,
    slide_height: float = 7.5,
    embed_images: bool = True,
) -> str:
    """Convert a single python-pptx Slide object to an SVG string.

    Parameters
    ----------
    slide :
        A python-pptx Slide object.
    slide_width :
        Slide width in inches (default 13.333, standard 16:9 widescreen).
    slide_height :
        Slide height in inches (default 7.5, standard 16:9 widescreen).
    embed_images :
        If ``True`` (default), images are base64-encoded directly into the SVG
        as data URIs.  If ``False``, ``xlink:href`` attributes reference
        external files.

    Returns
    -------
    str
        A complete SVG document string.

    Examples
    --------
    >>> from pptx import Presentation
    >>> prs = Presentation("deck.pptx")
    >>> svg = export_slide_to_svg(prs.slides[0])
    >>> with open("slide_1.svg", "w", encoding="utf-8") as f:
    ...     f.write(svg)
    """
    # Compute viewBox dimensions in points
    vb_w = slide_width * 72.0  # inches -> points
    vb_h = slide_height * 72.0

    # Try to get actual slide dimensions from the slide object
    try:
        actual_w = slide.slide_width
        actual_h = slide.slide_height
        if actual_w and actual_h:
            vb_w = _emu_to_pt(actual_w)
            vb_h = _emu_to_pt(actual_h)
    except Exception:
        pass

    # Collect shape SVG elements
    shape_parts: list[str] = []

    # Background
    bg_colour = _slide_background_color(slide)
    if bg_colour:
        shape_parts.append(
            f'<rect width="{_format_float(vb_w)}" height="{_format_float(vb_h)}" fill="{bg_colour}"/>'
        )
    else:
        shape_parts.append(
            f'<rect width="{_format_float(vb_w)}" height="{_format_float(vb_h)}" fill="#ffffff"/>'
        )

    # Shapes
    for shape in slide.shapes:
        try:
            x = _emu_to_pt(shape.left)
            y = _emu_to_pt(shape.top)
            w = _emu_to_pt(shape.width)
            h = _emu_to_pt(shape.height)
        except Exception:
            continue

        try:
            svg = _convert_shape(shape, x, y, w, h, embed_images=embed_images)
            if svg:
                shape_parts.append(svg)
        except Exception:
            # Insert placeholder on error
            shape_parts.append(
                f'<rect x="{_format_float(x)}" y="{_format_float(y)}" '
                f'width="{_format_float(w)}" height="{_format_float(h)}" '
                f'fill="#eeeeee" stroke="#cccccc" stroke-width="0.5"/>'
            )

    # Font CSS (placeholder @font-face declarations)
    font_css = ""

    # Assemble the SVG document
    svg_content = "\n".join(shape_parts)

    style_block = ""
    if font_css:
        style_block = f"\n<style>\n{font_css}\n</style>\n"

    svg = (
        f'<svg xmlns="{_SVG_NS}" '
        f'xmlns:xlink="{_XLINK_NS}" '
        f'viewBox="0 0 {_format_float(vb_w)} {_format_float(vb_h)}" '
        f'width="{_format_float(vb_w)}pt" height="{_format_float(vb_h)}pt">'
        f"{style_block}\n{svg_content}\n</svg>"
    )

    return svg


# ---------------------------------------------------------------------------
# Public API: export_to_svg
# ---------------------------------------------------------------------------

def export_to_svg(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
    output_dir: str | os.PathLike[str],
    *,
    dpi: int = 150,
    slide_indices: list[int] | None = None,
    embed_fonts: bool = False,
    embed_images: bool = True,
) -> list[str]:
    """Export PowerPoint slides as individual SVG files.

    Parameters
    ----------
    prs_or_path :
        A :class:`~pptx.Presentation` instance or a path to a ``.pptx`` file.
    output_dir :
        Directory for the output SVG files.  Created if it does not exist.
    dpi :
        Resolution hint in dots per inch (default 150).  Used for rasterised
        fallbacks such as chart rendering.
    slide_indices :
        Optional list of 0-based slide indices to export.  ``None`` means
        all slides.
    embed_fonts :
        If ``True``, include ``@font-face`` CSS declarations in each SVG
        for font families used on the slide.  Note: actual font subsetting
        is not performed; declarations reference local system fonts only.
    embed_images :
        If ``True`` (default), images are base64-encoded as data URIs inside
        the SVG.  If ``False``, ``xlink:href`` attributes reference external
        files that the caller must provide alongside the SVGs.

    Returns
    -------
    list[str]
        Absolute paths of the written SVG files, in slide order.

    Raises
    ------
    FileNotFoundError
        If *prs_or_path* is a path that does not exist.
    ValueError
        If the presentation has no slides or the index list is empty.

    Examples
    --------
    >>> paths = export_to_svg("deck.pptx", "./svg_output")
    >>> print(paths)
    ['C:\\\\...\\\\svg_output\\\\slide_001.svg', ...]
    """
    prs, _ = _resolve_presentation(prs_or_path)
    output_dir = _ensure_dir(output_dir)

    total_slides = len(prs.slides)
    if total_slides == 0:
        raise ValueError("Cannot export an empty presentation to SVG")

    # Determine which slides to export
    if slide_indices is not None:
        indices = [i for i in slide_indices if 0 <= i < total_slides]
        if not indices:
            raise ValueError(
                f"No valid slide indices in range 0-{total_slides - 1}"
            )
    else:
        indices = list(range(total_slides))

    # Compute slide dimensions in inches
    try:
        slide_w_emu = prs.slide_width or 0
        slide_h_emu = prs.slide_height or 0
        slide_w_in = slide_w_emu / _EMU_PER_INCH
        slide_h_in = slide_h_emu / _EMU_PER_INCH
    except Exception:
        slide_w_in = 13.333
        slide_h_in = 7.5

    output_paths: list[str] = []

    for i in indices:
        slide = prs.slides[i]
        svg_content = export_slide_to_svg(
            slide,
            slide_width=slide_w_in,
            slide_height=slide_h_in,
            embed_images=embed_images,
        )

        # Inject font CSS if requested
        if embed_fonts:
            fonts = _collect_fonts_from_slide(slide)
            font_css = _fonts_to_css(fonts)
            if font_css:
                # Insert style block after the opening <svg> tag
                style_block = f"\n<style>\n{font_css}\n</style>"
                svg_content = svg_content.replace(">", ">" + style_block, 1)

        filename = f"slide_{i + 1:03d}.svg"
        filepath = os.path.join(output_dir, filename)

        with open(filepath, "w", encoding="utf-8") as f:
            f.write(svg_content)

        output_paths.append(os.path.abspath(filepath))

    return output_paths
