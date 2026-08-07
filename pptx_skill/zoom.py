"""Slide Zoom, Section Zoom, and Summary Zoom for PPTX presentations.

PowerPoint (Office 365 / 2019+) supports three interactive navigation
features built on zoom thumbnails:

- **Slide Zoom**: a clickable thumbnail on one slide that jumps to another
  slide (with an optional return action on the target).
- **Section Zoom**: a clickable thumbnail that jumps to the first slide of a
  named section.
- **Summary Zoom**: a collection of section-zoom thumbnails arranged in a
  grid or list layout on a single summary slide.

Because python-pptx has no native API for zoom shapes, this module constructs
the OOXML elements directly using lxml.  Each zoom shape is a picture
(placeholder rectangle with the target slide's background colour) whose
``nvPr`` carries a ``<p15:zoom>`` extension and whose ``<a:hlinkClick>``
action jumps to the target slide.

Usage
-----
>>> from pptx_skill.zoom import add_slide_zoom, add_section_zoom, add_summary_zoom
>>> add_slide_zoom("deck.pptx", 1, target_slide_index=3)
>>> add_section_zoom("deck.pptx", 1, section_name="Results")
>>> add_summary_zoom("deck.pptx", 1, section_names=["Intro", "Results"])
"""
from __future__ import annotations

import logging
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

__all__ = [
    "ZoomInfo",
    "add_slide_zoom",
    "add_section_zoom",
    "add_summary_zoom",
    "remove_zoom",
    "list_zooms",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------

_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_NS_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"

# Clark-notation prefixes
_P = f"{{{_NS_P}}}"
_A = f"{{{_NS_A}}}"
_R = f"{{{_NS_R}}}"
_P14 = f"{{{_NS_P14}}}"
_P15 = f"{{{_NS_P15}}}"

# Known extension URIs for zoom features
_ZOOM_EXT_URI = "{BA7A70F8-3C3E-4F32-9C0E-417E468B1E77}"

# EMU conversion
_EMU_PER_PT = 12700
_EMU_PER_INCH = 914400

# ---------------------------------------------------------------------------
# Data class
# ---------------------------------------------------------------------------


@dataclass
class ZoomInfo:
    """Information about a zoom shape in a presentation.

    Attributes:
        name: Shape name assigned to the zoom thumbnail.
        slide_index: 1-based index of the slide that contains the zoom shape.
        zoom_type: One of ``"slide"``, ``"section"``, or ``"summary"``.
        target_slide_index: 1-based index of the target slide, or ``None``.
        target_section: Name of the target section, or ``None``.
        left: Left position in inches.
        top: Top position in inches.
        width: Width in inches.
        height: Height in inches.
    """

    name: str
    slide_index: int
    zoom_type: str  # "slide" | "section" | "summary"
    target_slide_index: int | None = None
    target_section: str | None = None
    left: float = 0.0
    top: float = 0.0
    width: float = 0.0
    height: float = 0.0


# ---------------------------------------------------------------------------
# Helpers – Presentation open / save
# ---------------------------------------------------------------------------


def _is_presentation(obj: Any) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path: Any) -> Any:
    """Return a ``Presentation`` from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path
    (``str | Path``).
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs: Any, path: str | Path | None) -> None:
    """Save *prs* back to *path*, creating a backup first."""
    if path is None:
        return
    import shutil

    p = Path(path)
    bak = p.with_suffix(".bak.pptx")
    if p.exists():
        shutil.copy2(str(p), str(bak))
    prs.save(str(p))


def _resolve_path(prs_or_path: Any) -> str | None:
    """Return the file path if *prs_or_path* is a path, else ``None``."""
    if _is_presentation(prs_or_path):
        return None
    return str(prs_or_path)


# ---------------------------------------------------------------------------
# Helpers – EMU conversion
# ---------------------------------------------------------------------------


def _pt_to_emu(pt: float) -> int:
    """Convert points to EMU (1 pt = 12700 EMU)."""
    return int(round(pt * _EMU_PER_PT))


def _inches_to_emu(inches: float) -> int:
    """Convert inches to EMU."""
    return int(round(inches * _EMU_PER_INCH))


# ---------------------------------------------------------------------------
# Helpers – section resolution
# ---------------------------------------------------------------------------


def _section_first_slide_index(prs: Any, section_name: str) -> int | None:
    """Return the 0-based index of the first slide in the named section.

    Returns ``None`` if the section does not exist or has no slides.
    """
    from pptx_skill.sections import list_sections

    for sec in list_sections(prs):
        if sec.name == section_name and sec.slide_indices:
            return sec.slide_indices[0]
    return None


def _all_section_names(prs: Any) -> list[str]:
    """Return the names of all sections in the presentation (in order)."""
    from pptx_skill.sections import list_sections

    return [sec.name for sec in list_sections(prs)]


# ---------------------------------------------------------------------------
# Helpers – slide background colour extraction
# ---------------------------------------------------------------------------

# Fallback colours for common themes
_BG_FALLBACK = "4472C4"  # A pleasant blue


def _slide_bg_color_hex(prs: Any, slide_index: int) -> str:
    """Extract the background colour of a slide as a hex string (no ``#``).

    Falls back to a default blue if no background fill is found.
    """
    try:
        slide = prs.slides[slide_index]
    except (IndexError, KeyError):
        return _BG_FALLBACK

    sld_elem = slide._element

    # Check <p:bg> -> <p:bgPr> -> <a:solidFill> -> <a:srgbClr>
    bg = sld_elem.find(f"{_P}bg")
    if bg is not None:
        bg_pr = bg.find(f"{_P}bgPr")
        if bg_pr is not None:
            solid_fill = bg_pr.find(f"{_P}solidFill")
            if solid_fill is None:
                # Also check under <a:solidFill> directly
                solid_fill = bg_pr.find(f"{_A}solidFill")
            if solid_fill is not None:
                srgb = solid_fill.find(f"{_A}srgbClr")
                if srgb is not None:
                    val = srgb.get("val")
                    if val:
                        return val

    # Check slide layout / master background (simplified — just use fallback)
    return _BG_FALLBACK


# ---------------------------------------------------------------------------
# Helpers – zoom shape XML construction
# ---------------------------------------------------------------------------


def _build_zoom_shape_xml(
    shape_id: int,
    shape_name: str,
    left_emu: int,
    top_emu: int,
    width_emu: int,
    height_emu: int,
    target_slide_index: int,
    bg_color_hex: str,
    zoom_type: str,
    return_to_slide: bool = False,
    source_slide_index: int | None = None,
) -> Any:
    """Build the lxml element tree for a zoom shape (placeholder rectangle).

    The shape is a ``<p:sp>`` with:
    - An ``<a:hlinkClick>`` action that jumps to the target slide.
    - A ``<p15:zoom>`` extension in ``nvPr`` marking it as a zoom shape.
    - A solid fill using the target slide's background colour (as thumbnail
      placeholder, since we cannot render actual thumbnails).

    Parameters
    ----------
    shape_id : int
        Unique shape ID within the slide.
    shape_name : str
        Human-readable shape name.
    left_emu, top_emu, width_emu, height_emu : int
        Position and size in EMU.
    target_slide_index : int
        0-based index of the target slide.
    bg_color_hex : str
        Hex colour string (no ``#``) for the placeholder fill.
    zoom_type : str
        One of ``"slide"``, ``"section"``, ``"summary"``.
    return_to_slide : bool
        If True, the target slide gets a return-to-source action.
    source_slide_index : int | None
        The 0-based index of the source slide (needed for return action).

    Returns
    -------
    lxml.etree._Element
        The ``<p:sp>`` element ready to insert into the shape tree.
    """
    from lxml import etree

    sp = etree.Element(f"{_P}sp")

    # --- nvSpPr (non-visual shape properties) ---
    nvSpPr = etree.SubElement(sp, f"{_P}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{_P}cNvPr")
    cNvPr.set("id", str(shape_id))
    cNvPr.set("name", shape_name)

    # Hyperlink click action — jump to target slide
    # We use an action hyperlink ("ppaction://hlinksldjump") which navigates
    # to a slide by 0-based index within the presentation.
    hlinkClick = etree.SubElement(cNvPr, f"{_A}hlinkClick")
    hlinkClick.set("action", f"ppaction://hlinksldjump,{target_slide_index}")

    cNvSpPr = etree.SubElement(nvSpPr, f"{_P}cNvSpPr")
    # Mark as a placeholder-like shape that does not appear in the outline
    sp_locks = etree.SubElement(cNvSpPr, f"{_A}spLocks")
    sp_locks.set("noGrp", "1")

    nvPr = etree.SubElement(nvSpPr, f"{_P}nvPr")

    # p15:zoom extension in nvPr
    extLst = etree.SubElement(nvPr, f"{_P}extLst")
    ext = etree.SubElement(extLst, f"{_P}ext")
    ext.set("uri", _ZOOM_EXT_URI)
    zoom_elem = etree.SubElement(ext, f"{_P15}zoom")
    # Store zoom type as an attribute for round-tripping
    zoom_elem.set("type", zoom_type)

    # --- spPr (shape properties) ---
    spPr = etree.SubElement(sp, f"{_P}spPr")

    # Xfrm (transform)
    xfrm = etree.SubElement(spPr, f"{_A}xfrm")
    off = etree.SubElement(xfrm, f"{_A}off")
    off.set("x", str(left_emu))
    off.set("y", str(top_emu))
    ext_elem = etree.SubElement(xfrm, f"{_A}ext")
    ext_elem.set("cx", str(width_emu))
    ext_elem.set("cy", str(height_emu))

    # Preset geometry — rectangle
    prstGeom = etree.SubElement(spPr, f"{_A}prstGeom")
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom, f"{_A}avLst")

    # Solid fill with the target slide's background colour
    solidFill = etree.SubElement(spPr, f"{_A}solidFill")
    srgbClr = etree.SubElement(solidFill, f"{_A}srgbClr")
    srgbClr.set("val", bg_color_hex)

    # Thin outline for visual definition
    ln = etree.SubElement(spPr, f"{_A}ln")
    ln.set("w", "12700")  # 1 pt
    solidFill_ln = etree.SubElement(ln, f"{_A}solidFill")
    srgbClr_ln = etree.SubElement(solidFill_ln, f"{_A}srgbClr")
    srgbClr_ln.set("val", "404040")

    # --- txBody (empty text body for completeness) ---
    txBody = etree.SubElement(sp, f"{_P}txBody")
    bodyPr = etree.SubElement(txBody, f"{_A}bodyPr")
    bodyPr.set("anchor", "ctr")
    etree.SubElement(txBody, f"{_A}lstStyle")

    # Add a paragraph with the zoom label
    p = etree.SubElement(txBody, f"{_A}p")
    r = etree.SubElement(p, f"{_A}r")
    rPr = etree.SubElement(r, f"{_A}rPr")
    rPr.set("lang", "en-US")
    rPr.set("sz", "1200")  # 12 pt
    rPr.set("b", "1")
    solidFill_r = etree.SubElement(rPr, f"{_A}solidFill")
    srgbClr_r = etree.SubElement(solidFill_r, f"{_A}srgbClr")
    srgbClr_r.set("val", "FFFFFF")
    t = etree.SubElement(r, f"{_A}t")
    t.text = shape_name

    return sp


# ---------------------------------------------------------------------------
# Helpers – insert shape into slide's shape tree
# ---------------------------------------------------------------------------


def _next_shape_id(slide: Any) -> int:
    """Return the next available shape ID for *slide*.

    Shape IDs must be unique within a slide. We find the current maximum and
    add one.
    """
    sp_tree = slide.shapes._spTree
    max_id = 1
    for elem in sp_tree.iter():
        id_attr = elem.get("id")
        if id_attr is not None:
            try:
                max_id = max(max_id, int(id_attr))
            except (ValueError, TypeError):
                pass
    return max_id + 1


def _insert_zoom_shape(slide: Any, shape_elem: Any) -> None:
    """Insert a zoom shape element into a slide's shape tree.

    The element is inserted before ``<p:extLst>`` if present, otherwise
    appended at the end of the shape tree.
    """
    sp_tree = slide.shapes._spTree
    ext_lst = sp_tree.find(f"{_P}extLst")
    if ext_lst is not None:
        ext_lst.addprevious(shape_elem)
    else:
        sp_tree.append(shape_elem)


# ---------------------------------------------------------------------------
# Helpers – return-to-source action
# ---------------------------------------------------------------------------


def _add_return_action(prs: Any, target_slide_index: int, source_slide_index: int) -> None:
    """Add a return-to-source click action on the target slide.

    When *return_to_slide* is True in :func:`add_slide_zoom`, clicking
    anywhere on the target slide returns to the source slide. This is
    implemented by adding a full-slide invisible shape with a
    ``ppaction://hlinksldjump`` action on the target.
    """
    from lxml import etree

    target_slide = prs.slides[target_slide_index]
    shape_id = _next_shape_id(target_slide)
    shape_name = f"ZoomReturn_{source_slide_index + 1}"

    sp = etree.Element(f"{_P}sp")

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f"{_P}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{_P}cNvPr")
    cNvPr.set("id", str(shape_id))
    cNvPr.set("name", shape_name)
    hlinkClick = etree.SubElement(cNvPr, f"{_A}hlinkClick")
    hlinkClick.set("action", f"ppaction://hlinksldjump,{source_slide_index}")

    cNvSpPr = etree.SubElement(nvSpPr, f"{_P}cNvSpPr")
    sp_locks = etree.SubElement(cNvSpPr, f"{_A}spLocks")
    sp_locks.set("noGrp", "1")
    sp_locks.set("noSelect", "1")

    nvPr = etree.SubElement(nvSpPr, f"{_P}nvPr")
    extLst = etree.SubElement(nvPr, f"{_P}extLst")
    ext = etree.SubElement(extLst, f"{_P}ext")
    ext.set("uri", _ZOOM_EXT_URI)
    zoom_elem = etree.SubElement(ext, f"{_P15}zoom")
    zoom_elem.set("type", "return")

    # spPr — full slide coverage, no fill, no line
    spPr = etree.SubElement(sp, f"{_P}spPr")
    xfrm = etree.SubElement(spPr, f"{_A}xfrm")
    off = etree.SubElement(xfrm, f"{_A}off")
    off.set("x", "0")
    off.set("y", "0")

    # Use the slide dimensions
    sld_w = prs.slide_width
    sld_h = prs.slide_height
    ext_elem = etree.SubElement(xfrm, f"{_A}ext")
    ext_elem.set("cx", str(sld_w))
    ext_elem.set("cy", str(sld_h))

    prstGeom = etree.SubElement(spPr, f"{_A}prstGeom")
    prstGeom.set("prst", "rect")
    etree.SubElement(prstGeom, f"{_A}avLst")

    # No fill (invisible shape)
    etree.SubElement(spPr, f"{_A}noFill")
    # No line
    ln = etree.SubElement(spPr, f"{_A}ln")
    ln.set("w", "0")
    etree.SubElement(ln, f"{_A}noFill")

    # Empty text body
    txBody = etree.SubElement(sp, f"{_P}txBody")
    etree.SubElement(txBody, f"{_A}bodyPr")
    etree.SubElement(txBody, f"{_A}lstStyle")
    etree.SubElement(txBody, f"{_A}p")

    # Insert at the bottom of the z-order so it sits behind other content
    _insert_zoom_shape(target_slide, sp)


# ---------------------------------------------------------------------------
# Helpers – grid layout computation
# ---------------------------------------------------------------------------


def _compute_grid_positions(
    n_items: int,
    left: float,
    top: float,
    total_width: float,
    total_height: float,
    item_width: float,
    item_height: float,
) -> list[tuple[float, float]]:
    """Compute (left, top) positions for *n_items* in a grid layout.

    Items are arranged in rows within the bounding box. The number of
    columns is chosen to best fill the available width.
    """
    if n_items == 0:
        return []

    gap = 0.2  # inches between items

    # How many columns fit?
    cols = max(1, int((total_width + gap) / (item_width + gap)))
    cols = min(cols, n_items)

    positions: list[tuple[float, float]] = []
    for i in range(n_items):
        row = i // cols
        col = i % cols
        x = left + col * (item_width + gap)
        y = top + row * (item_height + gap)
        positions.append((x, y))

    return positions


def _compute_list_positions(
    n_items: int,
    left: float,
    top: float,
    total_width: float,
    total_height: float,
    item_height: float,
) -> list[tuple[float, float, float]]:
    """Compute (left, top, width) positions for *n_items* in a vertical list.

    Each item spans the full width. Returns tuples of (left, top, width).
    """
    if n_items == 0:
        return []

    gap = 0.15  # inches between items
    item_width = total_width

    positions: list[tuple[float, float, float]] = []
    for i in range(n_items):
        y = top + i * (item_height + gap)
        positions.append((left, y, item_width))

    return positions


# ---------------------------------------------------------------------------
# Helpers – detect zoom shapes
# ---------------------------------------------------------------------------


def _is_zoom_shape(shape_elem: Any) -> bool:
    """Return True if the shape element contains a ``<p15:zoom>`` extension."""
    # Check in nvSpPr/nvPr/extLst
    for tag_prefix in (f"{_P}nvSpPr", f"{_P}nvGrpSpPr"):
        nv_tag = shape_elem.find(tag_prefix)
        if nv_tag is None:
            continue
        nvPr = nv_tag.find(f"{_P}nvPr")
        if nvPr is None:
            continue
        extLst = nvPr.find(f"{_P}extLst")
        if extLst is None:
            continue
        for ext in extLst.findall(f"{_P}ext"):
            zoom = ext.find(f"{_P15}zoom")
            if zoom is not None:
                return True
    return False


def _zoom_type_from_elem(shape_elem: Any) -> str:
    """Extract the zoom type from a shape element's ``<p15:zoom>`` attribute.

    Returns one of ``"slide"``, ``"section"``, ``"summary"``, or ``"return"``.
    Falls back to ``"slide"`` if the attribute is missing.
    """
    for tag_prefix in (f"{_P}nvSpPr", f"{_P}nvGrpSpPr"):
        nv_tag = shape_elem.find(tag_prefix)
        if nv_tag is None:
            continue
        nvPr = nv_tag.find(f"{_P}nvPr")
        if nvPr is None:
            continue
        extLst = nvPr.find(f"{_P}extLst")
        if extLst is None:
            continue
        for ext in extLst.findall(f"{_P}ext"):
            zoom = ext.find(f"{_P15}zoom")
            if zoom is not None:
                return zoom.get("type", "slide")
    return "slide"


def _target_slide_from_hlink(shape_elem: Any) -> int | None:
    """Extract the target slide index from a zoom shape's hlinkClick action.

    Returns the 0-based index or ``None``.
    """
    # Look for cNvPr/hlinkClick with action="ppaction://hlinksldjump,N"
    for tag_prefix in (f"{_P}nvSpPr", f"{_P}nvGrpSpPr"):
        nv_tag = shape_elem.find(tag_prefix)
        if nv_tag is None:
            continue
        cNvPr = nv_tag.find(f"{_P}cNvPr")
        if cNvPr is None:
            continue
        hlink = cNvPr.find(f"{_A}hlinkClick")
        if hlink is not None:
            action = hlink.get("action", "")
            if action.startswith("ppaction://hlinksldjump,"):
                try:
                    return int(action.split(",", 1)[1])
                except (ValueError, IndexError):
                    pass
    return None


def _shape_position_inches(shape_elem: Any) -> tuple[float, float, float, float]:
    """Extract (left, top, width, height) in inches from a shape element."""
    # Find xfrm under spPr or grpSpPr
    spPr = shape_elem.find(f"{_P}spPr")
    if spPr is None:
        spPr = shape_elem.find(f"{_P}grpSpPr")
    if spPr is None:
        return 0.0, 0.0, 0.0, 0.0

    xfrm = spPr.find(f"{_A}xfrm")
    if xfrm is None:
        return 0.0, 0.0, 0.0, 0.0

    off = xfrm.find(f"{_A}off")
    ext = xfrm.find(f"{_A}ext")

    left = int(off.get("x", "0")) / _EMU_PER_INCH if off is not None else 0.0
    top = int(off.get("y", "0")) / _EMU_PER_INCH if off is not None else 0.0
    width = int(ext.get("cx", "0")) / _EMU_PER_INCH if ext is not None else 0.0
    height = int(ext.get("cy", "0")) / _EMU_PER_INCH if ext is not None else 0.0

    return left, top, width, height


def _resolve_section_for_target(prs: Any, target_slide_index: int | None) -> str | None:
    """Find the section name that contains the target slide, if any."""
    if target_slide_index is None:
        return None
    from pptx_skill.sections import list_sections

    for sec in list_sections(prs):
        if target_slide_index in sec.slide_indices:
            return sec.name
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def add_slide_zoom(
    prs_or_path: Any,
    slide_index: int,
    *,
    target_slide_index: int,
    left: float = 1.0,
    top: float = 1.0,
    width: float = 3.0,
    height: float = 2.0,
    return_to_slide: bool = True,
) -> str:
    """Add a Slide Zoom — a clickable thumbnail that zooms into another slide.

    Creates a picture shape (filled with the target slide's background colour
    as a thumbnail placeholder) with a hyperlink action that jumps to the
    target slide.  When *return_to_slide* is True, an invisible full-slide
    shape is added on the target slide so that clicking it returns to the
    source.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        1-based slide index (1 = first slide).
    target_slide_index : int
        1-based index of the slide to zoom into.
    left : float
        Left position of the thumbnail in inches (default 1.0).
    top : float
        Top position of the thumbnail in inches (default 1.0).
    width : float
        Width of the thumbnail in inches (default 3.0).
    height : float
        Height of the thumbnail in inches (default 2.0).
    return_to_slide : bool
        If True (default), clicking the target slide returns to the source.

    Returns
    -------
    str
        The name of the created zoom shape.

    Raises
    ------
    IndexError
        If *slide_index* or *target_slide_index* is out of range.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    n_slides = len(prs.slides)
    if slide_index < 1 or slide_index > n_slides:
        raise IndexError(
            f"slide_index {slide_index} out of range (1..{n_slides})"
        )
    if target_slide_index < 1 or target_slide_index > n_slides:
        raise IndexError(
            f"target_slide_index {target_slide_index} out of range "
            f"(1..{n_slides})"
        )

    # Internal helpers use 0-based indices
    src_idx0 = slide_index - 1
    tgt_idx0 = target_slide_index - 1

    slide = prs.slides[src_idx0]
    shape_id = _next_shape_id(slide)
    shape_name = f"SlideZoom_{target_slide_index}"

    bg_color = _slide_bg_color_hex(prs, tgt_idx0)

    shape_elem = _build_zoom_shape_xml(
        shape_id=shape_id,
        shape_name=shape_name,
        left_emu=_inches_to_emu(left),
        top_emu=_inches_to_emu(top),
        width_emu=_inches_to_emu(width),
        height_emu=_inches_to_emu(height),
        target_slide_index=tgt_idx0,
        bg_color_hex=bg_color,
        zoom_type="slide",
        return_to_slide=return_to_slide,
        source_slide_index=src_idx0,
    )

    _insert_zoom_shape(slide, shape_elem)

    # Add return action on target slide if requested
    if return_to_slide and src_idx0 != tgt_idx0:
        _add_return_action(prs, tgt_idx0, src_idx0)

    _save_prs(prs, path)

    log.info(
        "Added slide zoom on slide %d targeting slide %d (shape=%r)",
        slide_index,
        target_slide_index,
        shape_name,
    )
    return shape_name


def add_section_zoom(
    prs_or_path: Any,
    slide_index: int,
    *,
    section_name: str,
    left: float = 1.0,
    top: float = 1.0,
    width: float = 3.0,
    height: float = 2.0,
) -> str:
    """Add a Section Zoom — a clickable thumbnail that zooms into a section.

    Creates a picture shape (filled with the section's first slide background
    colour) with a hyperlink action that jumps to the first slide of the named
    section.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        1-based slide index (1 = first slide).
    section_name : str
        The name of the section to zoom into.
    left : float
        Left position of the thumbnail in inches (default 1.0).
    top : float
        Top position of the thumbnail in inches (default 1.0).
    width : float
        Width of the thumbnail in inches (default 3.0).
    height : float
        Height of the thumbnail in inches (default 2.0).

    Returns
    -------
    str
        The name of the created zoom shape.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    ValueError
        If the named section does not exist or has no slides.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    n_slides = len(prs.slides)
    if slide_index < 1 or slide_index > n_slides:
        raise IndexError(
            f"slide_index {slide_index} out of range (1..{n_slides})"
        )

    # Internal helpers use 0-based indices
    src_idx0 = slide_index - 1

    target_slide = _section_first_slide_index(prs, section_name)
    if target_slide is None:
        raise ValueError(
            f"Section {section_name!r} not found or has no slides"
        )

    slide = prs.slides[src_idx0]
    shape_id = _next_shape_id(slide)

    # Sanitise section name for use as shape name
    safe_name = section_name.replace(" ", "_")[:28]
    shape_name = f"SectionZoom_{safe_name}"

    bg_color = _slide_bg_color_hex(prs, target_slide)

    shape_elem = _build_zoom_shape_xml(
        shape_id=shape_id,
        shape_name=shape_name,
        left_emu=_inches_to_emu(left),
        top_emu=_inches_to_emu(top),
        width_emu=_inches_to_emu(width),
        height_emu=_inches_to_emu(height),
        target_slide_index=target_slide,
        bg_color_hex=bg_color,
        zoom_type="section",
        return_to_slide=True,
        source_slide_index=src_idx0,
    )

    _insert_zoom_shape(slide, shape_elem)

    # Section zoom always has return-to-source
    if src_idx0 != target_slide:
        _add_return_action(prs, target_slide, src_idx0)

    _save_prs(prs, path)

    log.info(
        "Added section zoom on slide %d targeting section %r (slide %d, shape=%r)",
        slide_index,
        section_name,
        target_slide + 1,
        shape_name,
    )
    return shape_name


def add_summary_zoom(
    prs_or_path: Any,
    slide_index: int,
    *,
    section_names: list[str] | None = None,
    layout: str = "grid",
    left: float = 0.5,
    top: float = 0.5,
    width: float = 12.0,
    height: float = 6.5,
) -> list[str]:
    """Add a Summary Zoom — a collection of section-zoom thumbnails.

    Creates one clickable thumbnail per section (or per the explicitly listed
    *section_names*).  The thumbnails are arranged in a grid or vertical list
    layout within the specified bounding box.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        1-based slide index (1 = first slide).
    section_names : list[str] | None
        Explicit list of section names to include.  If ``None``, all sections
        in the presentation are included.
    layout : str
        ``"grid"`` (auto-arranged) or ``"list"`` (vertical stack).
    left : float
        Left position of the bounding box in inches (default 0.5).
    top : float
        Top position of the bounding box in inches (default 0.5).
    width : float
        Total width of the bounding box in inches (default 12.0).
    height : float
        Total height of the bounding box in inches (default 6.5).

    Returns
    -------
    list[str]
        A list of shape names for the created zoom thumbnails.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    ValueError
        If no sections are found, or any named section does not exist.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    n_slides = len(prs.slides)
    if slide_index < 1 or slide_index > n_slides:
        raise IndexError(
            f"slide_index {slide_index} out of range (1..{n_slides})"
        )

    # Internal helpers use 0-based indices
    src_idx0 = slide_index - 1

    # Resolve section names
    if section_names is None:
        section_names = _all_section_names(prs)

    if not section_names:
        raise ValueError("No sections found in the presentation")

    # Validate all sections exist and have slides
    targets: list[tuple[str, int]] = []  # (section_name, first_slide_index 0-based)
    for sec_name in section_names:
        idx = _section_first_slide_index(prs, sec_name)
        if idx is None:
            raise ValueError(
                f"Section {sec_name!r} not found or has no slides"
            )
        targets.append((sec_name, idx))

    slide = prs.slides[src_idx0]
    shape_names: list[str] = []

    if layout == "grid":
        # Compute per-item dimensions based on count
        n = len(targets)
        # Target ~3 items per row for grid, adjust item size
        cols = max(1, min(n, int(math.ceil(math.sqrt(n * width / height)))))
        item_w = (width - 0.2 * (cols - 1)) / cols
        rows = math.ceil(n / cols)
        item_h = min(item_w * 0.667, (height - 0.2 * max(0, rows - 1)) / max(1, rows))
        item_w = min(item_w, item_h / 0.5)  # cap width if height is constrained

        positions: list[Any] = _compute_grid_positions(n, left, top, width, height, item_w, item_h)

        for (sec_name, target_idx), (pos_left, pos_top) in zip(targets, positions, strict=True):
            shape_id = _next_shape_id(slide)
            safe_name = sec_name.replace(" ", "_")[:24]
            shape_name = f"SummaryZoom_{safe_name}"
            bg_color = _slide_bg_color_hex(prs, target_idx)

            shape_elem = _build_zoom_shape_xml(
                shape_id=shape_id,
                shape_name=shape_name,
                left_emu=_inches_to_emu(pos_left),
                top_emu=_inches_to_emu(pos_top),
                width_emu=_inches_to_emu(item_w),
                height_emu=_inches_to_emu(item_h),
                target_slide_index=target_idx,
                bg_color_hex=bg_color,
                zoom_type="summary",
                return_to_slide=True,
                source_slide_index=src_idx0,
            )

            _insert_zoom_shape(slide, shape_elem)
            shape_names.append(shape_name)

            # Add return action on target slide
            if src_idx0 != target_idx:
                _add_return_action(prs, target_idx, src_idx0)

    elif layout == "list":
        item_h = min(1.2, (height - 0.15 * max(0, len(targets) - 1)) / max(1, len(targets)))
        positions = _compute_list_positions(len(targets), left, top, width, height, item_h)

        for (sec_name, target_idx), (pos_left, pos_top, pos_w) in zip(targets, positions, strict=True):
            shape_id = _next_shape_id(slide)
            safe_name = sec_name.replace(" ", "_")[:24]
            shape_name = f"SummaryZoom_{safe_name}"
            bg_color = _slide_bg_color_hex(prs, target_idx)

            shape_elem = _build_zoom_shape_xml(
                shape_id=shape_id,
                shape_name=shape_name,
                left_emu=_inches_to_emu(pos_left),
                top_emu=_inches_to_emu(pos_top),
                width_emu=_inches_to_emu(pos_w),
                height_emu=_inches_to_emu(item_h),
                target_slide_index=target_idx,
                bg_color_hex=bg_color,
                zoom_type="summary",
                return_to_slide=True,
                source_slide_index=src_idx0,
            )

            _insert_zoom_shape(slide, shape_elem)
            shape_names.append(shape_name)

            # Add return action on target slide
            if src_idx0 != target_idx:
                _add_return_action(prs, target_idx, src_idx0)
    else:
        raise ValueError(f"Unknown layout {layout!r}; expected 'grid' or 'list'")

    _save_prs(prs, path)

    log.info(
        "Added summary zoom on slide %d with %d section(s) in %s layout",
        slide_index,
        len(targets),
        layout,
    )
    return shape_names


def remove_zoom(prs_or_path: Any, slide_index: int, shape_name: str) -> bool:
    """Remove a zoom shape from a slide.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        1-based slide index (1 = first slide).
    shape_name : str
        The name of the zoom shape to remove.

    Returns
    -------
    bool
        ``True`` if the shape was found and removed, ``False`` otherwise.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    n_slides = len(prs.slides)
    if slide_index < 1 or slide_index > n_slides:
        raise IndexError(
            f"slide_index {slide_index} out of range (1..{n_slides})"
        )

    slide = prs.slides[slide_index - 1]

    # Find the shape by name and verify it is a zoom shape
    for shape in slide.shapes:
        if shape.name == shape_name:
            if _is_zoom_shape(shape._element):
                sp_tree = slide.shapes._spTree
                sp_tree.remove(shape._element)
                _save_prs(prs, path)
                log.info("Removed zoom shape %r from slide %d", shape_name, slide_index)
                return True
            else:
                log.warning(
                    "Shape %r on slide %d is not a zoom shape; not removing",
                    shape_name,
                    slide_index,
                )
                return False

    log.warning("Shape %r not found on slide %d", shape_name, slide_index)
    return False


def list_zooms(prs_or_path: Any, slide_index: int | None = None) -> list[ZoomInfo]:
    """List all zoom shapes in the presentation.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int | None
        If provided, only return zoom shapes on this slide (1-based).
        If ``None``, return zoom shapes from all slides.

    Returns
    -------
    list[ZoomInfo]
        A list of :class:`ZoomInfo` objects describing each zoom shape.
    """
    prs = _open_prs(prs_or_path)

    n_slides = len(prs.slides)

    if slide_index is not None:
        if slide_index < 1 or slide_index > n_slides:
            raise IndexError(
                f"slide_index {slide_index} out of range (1..{n_slides})"
            )
        slide_range = [slide_index - 1]
    else:
        slide_range = list(range(n_slides))

    results: list[ZoomInfo] = []

    for idx in slide_range:
        slide = prs.slides[idx]
        for shape in slide.shapes:
            if not _is_zoom_shape(shape._element):
                continue

            zoom_type = _zoom_type_from_elem(shape._element)
            # Skip "return" type — those are invisible helper shapes
            if zoom_type == "return":
                continue

            target_idx = _target_slide_from_hlink(shape._element)
            left, top, w, h = _shape_position_inches(shape._element)
            target_section = _resolve_section_for_target(prs, target_idx)

            info = ZoomInfo(
                name=shape.name,
                slide_index=idx + 1,
                zoom_type=zoom_type,
                target_slide_index=target_idx + 1 if target_idx is not None else None,
                target_section=target_section,
                left=left,
                top=top,
                width=w,
                height=h,
            )
            results.append(info)

    return results
