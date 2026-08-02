"""Connectors, lines, arrows, and freeform paths for PowerPoint shapes.

Creates ``<p:cxnSp>`` (connector) and ``<p:sp>`` (line/freeform) elements
with full control over line style, dash patterns, and arrowhead markers.

OOXML reference: ECMA-376 Part 4, §21.3 (PresentationML — Connectors),
§20.1.2 (DrawingML — Line Properties), §20.1.9 (Custom Geometry).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

__all__ = [
    "ConnectorInfo",
    "LineInfo",
    "ARROW_NONE",
    "ARROW_TRIANGLE",
    "ARROW_STEALTH",
    "ARROW_DIAMOND",
    "ARROW_OVAL",
    "ARROW_OPEN",
    "DASH_SOLID",
    "DASH_DOT",
    "DASH_DASH",
    "DASH_LG_DASH",
    "DASH_DASH_DOT",
    "DASH_LG_DASH_DOT",
    "DASH_LG_DASH_DOT_DOT",
    "DASH_SYS_DASH",
    "DASH_SYS_DOT",
    "DASH_SYS_DASH_DOT",
    "add_connector",
    "add_curved_connector",
    "add_curve",
    "add_elbow_connector",
    "add_freeform",
    "add_line",
    "delete_connector",
    "list_connectors",
    "reroute_connector",
    "set_arrow_style",
    "set_line_style",
    "pt_to_emu",
    "inch_to_emu",
    "cm_to_emu",
]

# ---------------------------------------------------------------------------
# Unit helpers
# ---------------------------------------------------------------------------

def pt_to_emu(pt: float) -> int:
    return int(round(pt * 12700))

def inch_to_emu(inch: float) -> int:
    return int(round(inch * 914400))

def cm_to_emu(cm: float) -> int:
    return int(round(cm * 360000))


# ---------------------------------------------------------------------------
# Arrow / dash constants
# ---------------------------------------------------------------------------

ARROW_NONE = "none"
ARROW_TRIANGLE = "triangle"
ARROW_STEALTH = "stealth"
ARROW_DIAMOND = "diamond"
ARROW_OVAL = "oval"
ARROW_OPEN = "open"

DASH_SOLID = "solid"
DASH_DOT = "dot"
DASH_DASH = "dash"
DASH_LG_DASH = "lgDash"
DASH_DASH_DOT = "dashDot"
DASH_LG_DASH_DOT = "lgDashDot"
DASH_LG_DASH_DOT_DOT = "lgDashDotDot"
DASH_SYS_DASH = "sysDash"
DASH_SYS_DOT = "sysDot"
DASH_SYS_DASH_DOT = "sysDashDot"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class LineInfo:
    """Info about a line shape."""
    name: str
    start_x: int
    start_y: int
    end_x: int
    end_y: int
    line_color: str = ""
    line_width: int = 0
    dash_style: str = DASH_SOLID
    start_arrow: str = ARROW_NONE
    end_arrow: str = ARROW_NONE

@dataclass
class ConnectorInfo:
    """Info about a connector shape."""
    name: str
    connector_type: str  # "straight", "elbow", "curved"
    start_shape: str = ""
    start_side: str = ""
    end_shape: str = ""
    end_side: str = ""
    line_color: str = ""
    line_width: int = 0
    start_arrow: str = ARROW_NONE
    end_arrow: str = ARROW_NONE


# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------------------
# Internal helpers
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

def _get_shape_bounds(shape):
    """Get shape bounding box in EMU."""
    return shape.left, shape.top, shape.width, shape.height

def _get_connection_point(shape, side: str) -> tuple[int, int]:
    """Get the connection point (x, y) in EMU for a shape side."""
    left, top, width, height = _get_shape_bounds(shape)
    side_lower = side.lower()
    if side_lower == "top":
        return left + width // 2, top
    elif side_lower == "bottom":
        return left + width // 2, top + height
    elif side_lower == "left":
        return left, top + height // 2
    elif side_lower == "right":
        return left + width, top + height // 2
    else:
        return left + width // 2, top + height // 2

def _build_line_props(line_color: str | None, line_width: int | None,
                      dash_style: str | None, start_arrow: str | None,
                      end_arrow: str | None) -> Any:
    """Build <a:ln> element with styling."""
    from lxml import etree

    ln = etree.SubElement(etree.Element("dummy"), f"{{{_NS_A}}}ln")
    if line_width is not None:
        ln.set("w", str(line_width))
    else:
        ln.set("w", "12700")  # 1pt default

    if line_color:
        sf = etree.SubElement(ln, f"{{{_NS_A}}}solidFill")
        clr = etree.SubElement(sf, f"{{{_NS_A}}}srgbClr")
        clr.set("val", line_color.upper())

    if dash_style and dash_style != DASH_SOLID:
        prst_dash = etree.SubElement(ln, f"{{{_NS_A}}}prstDash")
        prst_dash.set("val", dash_style)

    if start_arrow and start_arrow != ARROW_NONE:
        he = etree.SubElement(ln, f"{{{_NS_A}}}headEnd")
        he.set("type", start_arrow)
        he.set("w", "med")
        he.set("len", "med")

    if end_arrow and end_arrow != ARROW_NONE:
        te = etree.SubElement(ln, f"{{{_NS_A}}}tailEnd")
        te.set("type", end_arrow)
        te.set("w", "med")
        te.set("len", "med")

    return ln


def _apply_line_props_to_sp_pr(sp_pr, line_color, line_width, dash_style,
                                start_arrow, end_arrow):
    """Apply line properties to an existing spPr element."""
    from lxml import etree

    # Remove existing ln
    for existing in sp_pr.findall(f"{{{_NS_A}}}ln"):
        sp_pr.remove(existing)

    ln = _build_line_props(line_color, line_width, dash_style, start_arrow, end_arrow)
    sp_pr.append(ln)


# ---------------------------------------------------------------------------
# Line creation
# ---------------------------------------------------------------------------

def add_line(prs_or_path, slide_index: int, *,
             start_x: int, start_y: int, end_x: int, end_y: int,
             line_color: str | None = None,
             line_width: int | None = None,
             line_style: str | None = None,
             dash_style: str | None = None,
             start_arrow: str | None = None,
             end_arrow: str | None = None,
             name: str | None = None) -> str:
    """Add a straight line shape to a slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    start_x, start_y : int
        Start point in EMU.
    end_x, end_y : int
        End point in EMU.
    line_color : str, optional
        Hex color. Default "000000".
    line_width : int, optional
        Width in EMU. Default 12700 (~1pt).
    start_arrow, end_arrow : str, optional
        Arrow type constant. Default ARROW_NONE.

    Returns
    -------
    str
        The shape name.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        sp_tree = slide.shapes._spTree

        _color = line_color or "000000"
        _width = line_width or 12700
        _name = name or f"Line {len(slide.shapes)}"

        cx = abs(end_x - start_x)
        cy = abs(end_y - start_y)
        off_x = min(start_x, end_x)
        off_y = min(start_y, end_y)

        sp = etree.SubElement(sp_tree, f"{{{_NS_P}}}sp")
        nvSpPr = etree.SubElement(sp, f"{{{_NS_P}}}nvSpPr")
        cNvPr = etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvSpPr")
        etree.SubElement(nvSpPr, f"{{{_NS_P}}}nvPr")

        spPr = etree.SubElement(sp, f"{{{_NS_A}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(off_x))
        off.set("y", str(off_y))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(max(cx, 1)))
        ext.set("cy", str(max(cy, 1)))

        prstGeom = etree.SubElement(spPr, f"{{{_NS_A}}}prstGeom")
        prstGeom.set("prst", "line")
        etree.SubElement(prstGeom, f"{{{_NS_A}}}avLst")

        _apply_line_props_to_sp_pr(spPr, _color, _width, dash_style,
                                    start_arrow, end_arrow)

        # No fill for lines
        noFill = etree.SubElement(spPr, f"{{{_NS_A}}}noFill")

        return _name
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Connector creation
# ---------------------------------------------------------------------------

def add_connector(prs_or_path, slide_index: int, *,
                  start_shape: str, start_side: str,
                  end_shape: str, end_side: str,
                  line_color: str | None = None,
                  line_width: int | None = None,
                  line_style: str | None = None,
                  dash_style: str | None = None,
                  start_arrow: str | None = None,
                  end_arrow: str | None = None,
                  name: str | None = None) -> str:
    """Add a straight connector between two shapes.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    start_shape, end_shape : str
        Shape names to connect.
    start_side, end_side : str
        "top", "bottom", "left", or "right".

    Returns
    -------
    str
        The connector shape name.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape1 = _find_shape(slide, start_shape)
        shape2 = _find_shape(slide, end_shape)
        if shape1 is None or shape2 is None:
            raise ValueError(f"Shape not found: {start_shape if shape1 is None else end_shape}")

        sp_tree = slide.shapes._spTree
        _color = line_color or "000000"
        _width = line_width or 12700
        _name = name or f"Connector {len(slide.shapes)}"

        sx, sy = _get_connection_point(shape1, start_side)
        ex, ey = _get_connection_point(shape2, end_side)

        cxnSp = etree.SubElement(sp_tree, f"{{{_NS_P}}}cxnSp")
        nvCxnSpPr = etree.SubElement(cxnSp, f"{{{_NS_P}}}nvCxnSpPr")
        cNvPr = etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        cNvCxnSpPr = etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}cNvCxnSpPr")
        stCxn = etree.SubElement(cNvCxnSpPr, f"{{{_NS_A}}}stCxn")
        stCxn.set("id", str(shape1.shape_id))
        endCxn = etree.SubElement(cNvCxnSpPr, f"{{{_NS_A}}}endCxn")
        endCxn.set("id", str(shape2.shape_id))
        etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}nvPr")

        spPr = etree.SubElement(cxnSp, f"{{{_NS_A}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(min(sx, ex)))
        off.set("y", str(min(sy, ey)))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(max(abs(ex - sx), 1)))
        ext.set("cy", str(max(abs(ey - sy), 1)))

        prstGeom = etree.SubElement(spPr, f"{{{_NS_A}}}prstGeom")
        prstGeom.set("prst", "line")
        etree.SubElement(prstGeom, f"{{{_NS_A}}}avLst")

        _apply_line_props_to_sp_pr(spPr, _color, _width, dash_style,
                                    start_arrow, end_arrow)

        return _name
    finally:
        _save_prs(prs, path)


def add_elbow_connector(prs_or_path, slide_index: int, *,
                        start_shape: str, start_side: str,
                        end_shape: str, end_side: str,
                        line_color: str | None = None,
                        line_width: int | None = None,
                        start_arrow: str | None = None,
                        end_arrow: str | None = None,
                        name: str | None = None) -> str:
    """Add an elbow (right-angle) connector between two shapes.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape1 = _find_shape(slide, start_shape)
        shape2 = _find_shape(slide, end_shape)
        if shape1 is None or shape2 is None:
            raise ValueError(f"Shape not found")

        sp_tree = slide.shapes._spTree
        _color = line_color or "000000"
        _width = line_width or 12700
        _name = name or f"Elbow {len(slide.shapes)}"

        sx, sy = _get_connection_point(shape1, start_side)
        ex, ey = _get_connection_point(shape2, end_side)

        cxnSp = etree.SubElement(sp_tree, f"{{{_NS_P}}}cxnSp")
        nvCxnSpPr = etree.SubElement(cxnSp, f"{{{_NS_P}}}nvCxnSpPr")
        cNvPr = etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        cNvCxnSpPr = etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}cNvCxnSpPr")
        stCxn = etree.SubElement(cNvCxnSpPr, f"{{{_NS_A}}}stCxn")
        stCxn.set("id", str(shape1.shape_id))
        endCxn = etree.SubElement(cNvCxnSpPr, f"{{{_NS_A}}}endCxn")
        endCxn.set("id", str(shape2.shape_id))
        etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}nvPr")

        spPr = etree.SubElement(cxnSp, f"{{{_NS_A}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(min(sx, ex)))
        off.set("y", str(min(sy, ey)))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(max(abs(ex - sx), 1)))
        ext.set("cy", str(max(abs(ey - sy), 1)))

        prstGeom = etree.SubElement(spPr, f"{{{_NS_A}}}prstGeom")
        prstGeom.set("prst", "bentConnector3")
        avLst = etree.SubElement(prstGeom, f"{{{_NS_A}}}avLst")
        # Adjust handle for the bend point (50% = default)
        gd = etree.SubElement(avLst, f"{{{_NS_A}}}gd")
        gd.set("name", "adj1")
        gd.set("fmla", "val 50000")

        _apply_line_props_to_sp_pr(spPr, _color, _width, None,
                                    start_arrow, end_arrow)
        return _name
    finally:
        _save_prs(prs, path)


def add_curved_connector(prs_or_path, slide_index: int, *,
                         start_shape: str, start_side: str,
                         end_shape: str, end_side: str,
                         line_color: str | None = None,
                         line_width: int | None = None,
                         start_arrow: str | None = None,
                         end_arrow: str | None = None,
                         name: str | None = None) -> str:
    """Add a curved connector between two shapes.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape1 = _find_shape(slide, start_shape)
        shape2 = _find_shape(slide, end_shape)
        if shape1 is None or shape2 is None:
            raise ValueError(f"Shape not found")

        sp_tree = slide.shapes._spTree
        _color = line_color or "000000"
        _width = line_width or 12700
        _name = name or f"CurvedConn {len(slide.shapes)}"

        sx, sy = _get_connection_point(shape1, start_side)
        ex, ey = _get_connection_point(shape2, end_side)

        cxnSp = etree.SubElement(sp_tree, f"{{{_NS_P}}}cxnSp")
        nvCxnSpPr = etree.SubElement(cxnSp, f"{{{_NS_P}}}nvCxnSpPr")
        cNvPr = etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        cNvCxnSpPr = etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}cNvCxnSpPr")
        stCxn = etree.SubElement(cNvCxnSpPr, f"{{{_NS_A}}}stCxn")
        stCxn.set("id", str(shape1.shape_id))
        endCxn = etree.SubElement(cNvCxnSpPr, f"{{{_NS_A}}}endCxn")
        endCxn.set("id", str(shape2.shape_id))
        etree.SubElement(nvCxnSpPr, f"{{{_NS_P}}}nvPr")

        spPr = etree.SubElement(cxnSp, f"{{{_NS_A}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(min(sx, ex)))
        off.set("y", str(min(sy, ey)))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(max(abs(ex - sx), 1)))
        ext.set("cy", str(max(abs(ey - sy), 1)))

        prstGeom = etree.SubElement(spPr, f"{{{_NS_A}}}prstGeom")
        prstGeom.set("prst", "curvedConnector3")
        avLst = etree.SubElement(prstGeom, f"{{{_NS_A}}}avLst")
        gd = etree.SubElement(avLst, f"{{{_NS_A}}}gd")
        gd.set("name", "adj1")
        gd.set("fmla", "val 50000")

        _apply_line_props_to_sp_pr(spPr, _color, _width, None,
                                    start_arrow, end_arrow)
        return _name
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Curve / freeform
# ---------------------------------------------------------------------------

def add_curve(prs_or_path, slide_index: int, *,
              points: list[tuple[int, int]],
              line_color: str | None = None,
              line_width: int | None = None,
              line_style: str | None = None,
              dash_style: str | None = None,
              start_arrow: str | None = None,
              end_arrow: str | None = None,
              name: str | None = None) -> str:
    """Add a curve (polyline) through the given points.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    points : list of (x, y) tuples
        Points in EMU. At least 2 points required.
    """
    from lxml import etree

    if len(points) < 2:
        raise ValueError("At least 2 points required for a curve")

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        sp_tree = slide.shapes._spTree
        _color = line_color or "000000"
        _width = line_width or 12700
        _name = name or f"Curve {len(slide.shapes)}"

        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        min_x, min_y = min(xs), min(ys)
        max_x, max_y = max(xs), max(ys)

        sp = etree.SubElement(sp_tree, f"{{{_NS_P}}}sp")
        nvSpPr = etree.SubElement(sp, f"{{{_NS_P}}}nvSpPr")
        cNvPr = etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvSpPr")
        etree.SubElement(nvSpPr, f"{{{_NS_P}}}nvPr")

        spPr = etree.SubElement(sp, f"{{{_NS_A}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(min_x))
        off.set("y", str(min_y))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(max(max_x - min_x, 1)))
        ext.set("cy", str(max(max_y - min_y, 1)))

        custGeom = etree.SubElement(spPr, f"{{{_NS_A}}}custGeom")
        avLst = etree.SubElement(custGeom, f"{{{_NS_A}}}avLst")
        gdLst = etree.SubElement(custGeom, f"{{{_NS_A}}}gdLst")
        ahLst = etree.SubElement(custGeom, f"{{{_NS_A}}}ahLst")
        cxnLst = etree.SubElement(custGeom, f"{{{_NS_A}}}cxnLst")
        rect = etree.SubElement(custGeom, f"{{{_NS_A}}}rect")
        rect.set("l", "0")
        rect.set("t", "0")
        rect.set("r", "0")
        rect.set("b", "0")

        pathLst = etree.SubElement(custGeom, f"{{{_NS_A}}}pathLst")
        path = etree.SubElement(pathLst, f"{{{_NS_A}}}path")

        # moveTo for first point
        moveTo = etree.SubElement(path, f"{{{_NS_A}}}moveTo")
        pt = etree.SubElement(moveTo, f"{{{_NS_A}}}pt")
        pt.set("x", str(points[0][0]))
        pt.set("y", str(points[0][1]))

        # lnTo for remaining points
        for px, py in points[1:]:
            lnTo = etree.SubElement(path, f"{{{_NS_A}}}lnTo")
            pt = etree.SubElement(lnTo, f"{{{_NS_A}}}pt")
            pt.set("x", str(px))
            pt.set("y", str(py))

        _apply_line_props_to_sp_pr(spPr, _color, _width, dash_style,
                                    start_arrow, end_arrow)
        noFill = etree.SubElement(spPr, f"{{{_NS_A}}}noFill")

        return _name
    finally:
        _save_prs(prs, path)


def add_freeform(prs_or_path, slide_index: int, *,
                 points: list[tuple[int, int]],
                 fill_color: str | None = None,
                 line_color: str | None = None,
                 line_width: int | None = None,
                 name: str | None = None) -> str:
    """Add a closed freeform shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    points : list of (x, y) tuples
        Points in EMU. At least 3 points required. Shape is auto-closed.
    fill_color : str, optional
        Fill hex color. Default no fill.
    """
    from lxml import etree

    if len(points) < 3:
        raise ValueError("At least 3 points required for a freeform shape")

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        sp_tree = slide.shapes._spTree
        _line_color = line_color or "000000"
        _width = line_width or 12700
        _name = name or f"Freeform {len(slide.shapes)}"

        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        min_x, min_y = min(xs), min(ys)
        max_x, max_y = max(xs), max(ys)

        sp = etree.SubElement(sp_tree, f"{{{_NS_P}}}sp")
        nvSpPr = etree.SubElement(sp, f"{{{_NS_P}}}nvSpPr")
        cNvPr = etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvSpPr")
        etree.SubElement(nvSpPr, f"{{{_NS_P}}}nvPr")

        spPr = etree.SubElement(sp, f"{{{_NS_A}}}spPr")
        xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(min_x))
        off.set("y", str(min_y))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(max(max_x - min_x, 1)))
        ext.set("cy", str(max(max_y - min_y, 1)))

        custGeom = etree.SubElement(spPr, f"{{{_NS_A}}}custGeom")
        avLst = etree.SubElement(custGeom, f"{{{_NS_A}}}avLst")
        gdLst = etree.SubElement(custGeom, f"{{{_NS_A}}}gdLst")
        ahLst = etree.SubElement(custGeom, f"{{{_NS_A}}}ahLst")
        cxnLst = etree.SubElement(custGeom, f"{{{_NS_A}}}cxnLst")
        rect = etree.SubElement(custGeom, f"{{{_NS_A}}}rect")
        rect.set("l", "0"); rect.set("t", "0"); rect.set("r", "0"); rect.set("b", "0")

        pathLst = etree.SubElement(custGeom, f"{{{_NS_A}}}pathLst")
        path = etree.SubElement(pathLst, f"{{{_NS_A}}}path")

        moveTo = etree.SubElement(path, f"{{{_NS_A}}}moveTo")
        pt = etree.SubElement(moveTo, f"{{{_NS_A}}}pt")
        pt.set("x", str(points[0][0]))
        pt.set("y", str(points[0][1]))

        for px, py in points[1:]:
            lnTo = etree.SubElement(path, f"{{{_NS_A}}}lnTo")
            pt = etree.SubElement(lnTo, f"{{{_NS_A}}}pt")
            pt.set("x", str(px))
            pt.set("y", str(py))

        close = etree.SubElement(path, f"{{{_NS_A}}}close")

        # Fill
        if fill_color:
            sf = etree.SubElement(spPr, f"{{{_NS_A}}}solidFill")
            clr = etree.SubElement(sf, f"{{{_NS_A}}}srgbClr")
            clr.set("val", fill_color.upper())
        else:
            etree.SubElement(spPr, f"{{{_NS_A}}}noFill")

        _apply_line_props_to_sp_pr(spPr, _line_color, _width, None, None, None)

        return _name
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Modification
# ---------------------------------------------------------------------------

def set_line_style(prs_or_path, slide_index: int, shape_name: str, *,
                   line_color: str | None = None,
                   line_width: int | None = None,
                   line_style: str | None = None,
                   dash_style: str | None = None) -> bool:
    """Set line style properties on an existing shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False
        _apply_line_props_to_sp_pr(sp_pr, line_color, line_width,
                                    dash_style, None, None)
        return True
    finally:
        _save_prs(prs, path)


def set_arrow_style(prs_or_path, slide_index: int, shape_name: str, *,
                    start_arrow: str | None = None,
                    end_arrow: str | None = None,
                    start_arrow_size: str | None = None,
                    end_arrow_size: str | None = None) -> bool:
    """Set arrowhead style on an existing line/connector shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False
        ln = sp_pr.find(f"{{{_NS_A}}}ln")
        if ln is None:
            ln = etree.SubElement(sp_pr, f"{{{_NS_A}}}ln")
            ln.set("w", "12700")

        if start_arrow is not None:
            for existing in ln.findall(f"{{{_NS_A}}}headEnd"):
                ln.remove(existing)
            if start_arrow != ARROW_NONE:
                he = etree.SubElement(ln, f"{{{_NS_A}}}headEnd")
                he.set("type", start_arrow)
                he.set("w", start_arrow_size or "med")
                he.set("len", "med")

        if end_arrow is not None:
            for existing in ln.findall(f"{{{_NS_A}}}tailEnd"):
                ln.remove(existing)
            if end_arrow != ARROW_NONE:
                te = etree.SubElement(ln, f"{{{_NS_A}}}tailEnd")
                te.set("type", end_arrow)
                te.set("w", end_arrow_size or "med")
                te.set("len", "med")

        return True
    finally:
        _save_prs(prs, path)


def reroute_connector(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Re-route a connector to its connected shapes' current positions.

    This recalculates the connector's bounding box based on the current
    positions of the start and end shapes.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        elem = shape._element
        # Find connection info
        nvCxnSpPr = elem.find(f".//{{{_NS_P}}}nvCxnSpPr")
        if nvCxnSpPr is None:
            return False

        cNvCxnSpPr = nvCxnSpPr.find(f"{{{_NS_P}}}cNvCxnSpPr")
        if cNvCxnSpPr is None:
            return False

        stCxn = cNvCxnSpPr.find(f"{{{_NS_A}}}stCxn")
        endCxn = cNvCxnSpPr.find(f"{{{_NS_A}}}endCxn")
        if stCxn is None or endCxn is None:
            return False

        start_id = int(stCxn.get("id", 0))
        end_id = int(endCxn.get("id", 0))

        # Find connected shapes
        start_shape = None
        end_shape = None
        for s in slide.shapes:
            if s.shape_id == start_id:
                start_shape = s
            if s.shape_id == end_id:
                end_shape = s

        if start_shape is None or end_shape is None:
            return False

        # Update xfrm
        sp_pr = elem.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        xfrm = sp_pr.find(f"{{{_NS_A}}}xfrm")
        if xfrm is None:
            return False

        sx = start_shape.left + start_shape.width // 2
        sy = start_shape.top + start_shape.height // 2
        ex = end_shape.left + end_shape.width // 2
        ey = end_shape.top + end_shape.height // 2

        off = xfrm.find(f"{{{_NS_A}}}off")
        ext = xfrm.find(f"{{{_NS_A}}}ext")
        if off is not None:
            off.set("x", str(min(sx, ex)))
            off.set("y", str(min(sy, ey)))
        if ext is not None:
            ext.set("cx", str(max(abs(ex - sx), 1)))
            ext.set("cy", str(max(abs(ey - sy), 1)))

        return True
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Deletion / listing
# ---------------------------------------------------------------------------

def delete_connector(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Delete a connector or line shape by name.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_tree = slide.shapes._spTree
        sp_tree.remove(shape._element)
        return True
    finally:
        _save_prs(prs, path)


def list_connectors(prs_or_path, slide_index: int) -> list[ConnectorInfo | LineInfo]:
    """List all connectors and lines on a slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        results: list[ConnectorInfo | LineInfo] = []

        for shape in slide.shapes:
            elem = shape._element
            tag = elem.tag

            # Check if it's a connector
            if tag == f"{{{_NS_P}}}cxnSp":
                info = ConnectorInfo(name=shape.name, connector_type="straight")
                nvCxnSpPr = elem.find(f".//{{{_NS_P}}}nvCxnSpPr")
                if nvCxnSpPr is not None:
                    cNvCxnSpPr = nvCxnSpPr.find(f"{{{_NS_P}}}cNvCxnSpPr")
                    if cNvCxnSpPr is not None:
                        stCxn = cNvCxnSpPr.find(f"{{{_NS_A}}}stCxn")
                        endCxn = cNvCxnSpPr.find(f"{{{_NS_A}}}endCxn")
                        if stCxn is not None:
                            info.start_shape = stCxn.get("id", "")
                        if endCxn is not None:
                            info.end_shape = endCxn.get("id", "")

                # Check geometry type
                sp_pr = elem.find(f"{{{_NS_A}}}spPr")
                if sp_pr is not None:
                    prstGeom = sp_pr.find(f"{{{_NS_A}}}prstGeom")
                    if prstGeom is not None:
                        prst = prstGeom.get("prst", "line")
                        if prst == "bentConnector3":
                            info.connector_type = "elbow"
                        elif prst == "curvedConnector3":
                            info.connector_type = "curved"

                    ln = sp_pr.find(f"{{{_NS_A}}}ln")
                    if ln is not None:
                        sf = ln.find(f"{{{_NS_A}}}solidFill")
                        if sf is not None:
                            clr = sf.find(f"{{{_NS_A}}}srgbClr")
                            if clr is not None:
                                info.line_color = clr.get("val", "")
                        info.line_width = int(ln.get("w", "0"))

                        he = ln.find(f"{{{_NS_A}}}headEnd")
                        if he is not None:
                            info.start_arrow = he.get("type", ARROW_NONE)
                        te = ln.find(f"{{{_NS_A}}}tailEnd")
                        if te is not None:
                            info.end_arrow = te.get("type", ARROW_NONE)

                results.append(info)

            # Check if it's a line shape
            elif tag == f"{{{_NS_P}}}sp":
                sp_pr = elem.find(f"{{{_NS_A}}}spPr")
                if sp_pr is not None:
                    prstGeom = sp_pr.find(f"{{{_NS_A}}}prstGeom")
                    if prstGeom is not None and prstGeom.get("prst") == "line":
                        info = LineInfo(
                            name=shape.name,
                            start_x=shape.left,
                            start_y=shape.top,
                            end_x=shape.left + shape.width,
                            end_y=shape.top + shape.height,
                        )
                        ln = sp_pr.find(f"{{{_NS_A}}}ln")
                        if ln is not None:
                            sf = ln.find(f"{{{_NS_A}}}solidFill")
                            if sf is not None:
                                clr = sf.find(f"{{{_NS_A}}}srgbClr")
                                if clr is not None:
                                    info.line_color = clr.get("val", "")
                            info.line_width = int(ln.get("w", "0"))
                            prstDash = ln.find(f"{{{_NS_A}}}prstDash")
                            if prstDash is not None:
                                info.dash_style = prstDash.get("val", DASH_SOLID)
                            he = ln.find(f"{{{_NS_A}}}headEnd")
                            if he is not None:
                                info.start_arrow = he.get("type", ARROW_NONE)
                            te = ln.find(f"{{{_NS_A}}}tailEnd")
                            if te is not None:
                                info.end_arrow = te.get("type", ARROW_NONE)

                        results.append(info)

        return results
    finally:
        pass  # Read-only operation
