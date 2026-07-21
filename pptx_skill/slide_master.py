"""Slide master and layout operations for PPTX presentations.

Provides query, modification, and placeholder operations on slide masters
and their layouts.  Uses python-pptx's ``SlideMasters`` / ``SlideLayouts``
API where available and falls back to lxml for XML manipulation where the
python-pptx API is insufficient (e.g. cloning layouts, adding placeholders,
setting backgrounds).

Public API
----------
Query:
- :func:`list_slide_masters`   -- list all slide masters
- :func:`list_layouts`         -- list layouts for a given master
- :func:`get_layout_by_name`   -- find a layout by name
- :func:`analyze_layout`       -- detailed placeholder analysis

Modification:
- :func:`clone_layout`         -- duplicate a layout under a new name
- :func:`delete_layout`        -- remove a layout from its master
- :func:`rename_layout`        -- rename a layout

Placeholder:
- :func:`add_placeholder_to_layout`    -- add a placeholder shape
- :func:`remove_placeholder_from_layout` -- remove a placeholder by index
- :func:`list_placeholders`            -- list placeholder details

Background:
- :func:`set_master_background`  -- set master background (color/image/gradient)
- :func:`set_layout_background`  -- set layout background (color/image/gradient)
"""
from __future__ import annotations

import logging
import re
import shutil
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

__all__ = [
    "MasterInfo",
    "LayoutInfo",
    "add_placeholder_to_layout",
    "analyze_layout",
    "clone_layout",
    "delete_layout",
    "get_layout_by_name",
    "list_layouts",
    "list_placeholders",
    "list_slide_masters",
    "remove_placeholder_from_layout",
    "rename_layout",
    "set_layout_background",
    "set_master_background",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------

_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_P_NS_PREFIX = f"{{{_P_NS}}}"
_A_NS_PREFIX = f"{{{_A_NS}}}"

# EMU conversion constants
_EMU_PER_PT = 12700
_EMU_PER_INCH = 914400

# Valid placeholder types (OOXML p:ph @type values)
_VALID_PH_TYPES = frozenset({
    "title", "body", "ctrTitle", "subTitle",
    "dt", "sldNum", "ftr", "hdr",
    "pic", "chart", "tbl", "dgm", "media", "sldImg",
    "obj",
})

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class MasterInfo:
    """Summary information about a slide master."""

    name: str
    layout_count: int
    slide_count: int


@dataclass(frozen=True)
class LayoutInfo:
    """Summary information about a slide layout."""

    name: str
    master_name: str
    placeholder_count: int
    placeholder_types: list[str] = field(default_factory=list)

    def __hash__(self) -> int:
        return hash((self.name, self.master_name))

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, LayoutInfo):
            return NotImplemented
        return self.name == other.name and self.master_name == other.master_name


# ---------------------------------------------------------------------------
# Helpers – Presentation open / save
# ---------------------------------------------------------------------------


def _is_presentation(obj: Any) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path: Any) -> Any:
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path
    (``str | Path``).  Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _resolve_path(prs_or_path: Any) -> str | None:
    """Return the file path if *prs_or_path* is a path, else ``None``."""
    if _is_presentation(prs_or_path):
        return None
    return str(prs_or_path)


def _save_prs(prs: Any, path: str | Path | None) -> None:
    """Save *prs* back to *path*, creating a backup first."""
    if path is None:
        return
    p = Path(path)
    bak = p.with_suffix(".bak.pptx")
    if p.exists():
        shutil.copy2(str(p), str(bak))
    prs.save(str(p))


def _get_master(prs: Any, master_index: int = 0) -> Any:
    """Return the slide master at *master_index* with bounds checking."""
    masters = prs.slide_masters
    if not 0 <= master_index < len(masters):
        raise IndexError(
            f"master_index {master_index} out of range "
            f"(presentation has {len(masters)} slide master(s))"
        )
    return masters[master_index]


def _pt_to_emu(pt: float) -> int:
    """Convert points to EMU."""
    return int(round(pt * _EMU_PER_PT))


# ---------------------------------------------------------------------------
# Placeholder helpers
# ---------------------------------------------------------------------------


def _ph_type_from_element(sp_elem: Any) -> str:
    """Extract placeholder type from a shape XML element.

    Returns the ``type`` attribute value from ``p:nvSpPr/p:nvPr/p:ph``,
    defaulting to ``"obj"`` if no type is specified.
    """
    from lxml import etree

    nvSpPr = sp_elem.find(f"{_P_NS_PREFIX}nvSpPr")
    if nvSpPr is None:
        return "obj"
    nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
    if nvPr is None:
        return "obj"
    ph = nvPr.find(f"{_P_NS_PREFIX}ph")
    if ph is None:
        return "obj"
    return ph.get("type", "obj")


def _ph_index_from_element(sp_elem: Any) -> int | None:
    """Extract placeholder index from a shape XML element.

    Returns the ``idx`` attribute value from ``p:nvSpPr/p:nvPr/p:ph``,
    or ``None`` if no index is specified.
    """
    from lxml import etree

    nvSpPr = sp_elem.find(f"{_P_NS_PREFIX}nvSpPr")
    if nvSpPr is None:
        return None
    nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
    if nvPr is None:
        return None
    ph = nvPr.find(f"{_P_NS_PREFIX}ph")
    if ph is None:
        return None
    idx_str = ph.get("idx")
    if idx_str is None:
        return None
    try:
        return int(idx_str)
    except (ValueError, TypeError):
        return None


def _layout_name(layout: Any) -> str:
    """Return the display name of a layout, falling back to the extent python-pptx exposes it."""
    # python-pptx SlideLayout has a .name property
    name = getattr(layout, "name", None)
    if name:
        return name
    # Fall back to parsing the XML
    from lxml import etree

    xml_elem = layout._element
    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is not None:
        name_attr = cSld.get("name")
        if name_attr:
            return name_attr
    return f"Layout_{id(layout)}"


def _master_name(master: Any) -> str:
    """Return the display name of a slide master."""
    name = getattr(master, "name", None)
    if name:
        return name
    from lxml import etree

    xml_elem = master._element
    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is not None:
        name_attr = cSld.get("name")
        if name_attr:
            return name_attr
    return f"SlideMaster_{id(master)}"


def _count_slides_using_layout(prs: Any, layout: Any) -> int:
    """Count how many slides use the given layout."""
    count = 0
    for slide in prs.slides:
        if slide.slide_layout is layout:
            count += 1
    return count


# ---------------------------------------------------------------------------
# Background helpers
# ---------------------------------------------------------------------------


def _set_background_on_element(xml_elem: Any, *, color: str | None = None,
                               image_path: str | None = None,
                               gradient: dict[str, Any] | None = None) -> None:
    """Set background properties on an XML element (master or layout).

    Manipulates the ``p:cSld/p:bg`` subtree directly.  Supports solid fill,
    image fill, and gradient fill.  Only one fill type should be specified;
    if multiple are given, the first non-None value wins in order:
    color > image_path > gradient.
    """
    from lxml import etree

    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is None:
        return

    # Remove existing background
    existing_bg = cSld.find(f"{_P_NS_PREFIX}bg")
    if existing_bg is not None:
        cSld.remove(existing_bg)

    # If all None, we've effectively cleared the background
    if color is None and image_path is None and gradient is None:
        return

    # Build new bg element
    bg = etree.SubElement(cSld, f"{_P_NS_PREFIX}bg")
    bgPr = etree.SubElement(bg, f"{_P_NS_PREFIX}bgPr")

    if color is not None:
        # Solid fill
        solidFill = etree.SubElement(bgPr, f"{_A_NS_PREFIX}solidFill")
        srgbClr = etree.SubElement(solidFill, f"{_A_NS_PREFIX}srgbClr")
        # Strip leading # if present
        hex_val = color.lstrip("#")
        srgbClr.set("val", hex_val)

    elif image_path is not None:
        # Image fill – we need to add the image to the package and create
        # a relationship.  This requires the parent part (master or layout).
        # We handle this in the calling function which has access to the part.
        # Here we just set a marker; the caller will fill in the blip reference.
        # Actually, we need the part context, so we'll handle image fills
        # in the public functions directly.
        raise NotImplementedError(
            "Image backgrounds must be set via set_master_background / "
            "set_layout_background which have access to the part object."
        )

    elif gradient is not None:
        # Gradient fill
        gradFill = etree.SubElement(bgPr, f"{_A_NS_PREFIX}gradFill")
        gsLst = etree.SubElement(gradFill, f"{_A_NS_PREFIX}gsLst")

        stops = gradient.get("stops", [])
        if not stops:
            # Default two-stop gradient
            stops = [
                {"position": 0, "color": "FFFFFF"},
                {"position": 100000, "color": "000000"},
            ]

        for stop in stops:
            gs = etree.SubElement(gsLst, f"{_A_NS_PREFIX}gs")
            pos = stop.get("position", 0)
            if isinstance(pos, float) and pos <= 1.0:
                pos = int(pos * 100000)
            gs.set("pos", str(pos))
            srgbClr = etree.SubElement(gs, f"{_A_NS_PREFIX}srgbClr")
            hex_val = str(stop.get("color", "000000")).lstrip("#")
            srgbClr.set("val", hex_val)

        # Linear direction
        angle = gradient.get("angle", 5400000)  # default: top-to-bottom
        lin = etree.SubElement(gradFill, f"{_A_NS_PREFIX}lin")
        lin.set("ang", str(angle))
        lin.set("scaled", "1")

    # Add effect list (required by spec, even if empty)
    etree.SubElement(bgPr, f"{_A_NS_PREFIX}effectLst")


def _set_image_background_on_part(part: Any, xml_elem: Any,
                                  image_path: str) -> None:
    """Set an image background on a part, adding the image to the package.

    *part* is the python-pptx part object (e.g. slide_master.part or
    slide_layout.part).  *xml_elem* is the underlying XML element.
    """
    from lxml import etree

    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is None:
        return

    # Remove existing background
    existing_bg = cSld.find(f"{_P_NS_PREFIX}bg")
    if existing_bg is not None:
        cSld.remove(existing_bg)

    # Add image to the package via the part's image management.
    # get_or_add_image_part returns (image_part, rId).
    image_part, rId = part.get_or_add_image_part(image_path)

    # Build bg element with blipFill
    bg = etree.SubElement(cSld, f"{_P_NS_PREFIX}bg")
    bgPr = etree.SubElement(bg, f"{_P_NS_PREFIX}bgPr")
    blipFill = etree.SubElement(bgPr, f"{_A_NS_PREFIX}blipFill")
    blip = etree.SubElement(blipFill, f"{_A_NS_PREFIX}blip")
    blip.set(f"{{{_R_NS}}}embed", rId)
    stretch = etree.SubElement(blipFill, f"{_A_NS_PREFIX}stretch")
    fillRect = etree.SubElement(stretch, f"{_A_NS_PREFIX}fillRect")

    etree.SubElement(bgPr, f"{_A_NS_PREFIX}effectLst")


# ---------------------------------------------------------------------------
# Query functions
# ---------------------------------------------------------------------------


def list_slide_masters(prs_or_path: Any) -> list[MasterInfo]:
    """List all slide masters in the presentation.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.

    Returns:
        A list of :class:`MasterInfo` objects, one per slide master.
    """
    prs = _open_prs(prs_or_path)
    result: list[MasterInfo] = []
    for master in prs.slide_masters:
        name = _master_name(master)
        layout_count = len(master.slide_layouts)
        # Count slides that belong to this master (via any of its layouts)
        slide_count = 0
        layout_set = set(id(lo) for lo in master.slide_layouts)
        for slide in prs.slides:
            if id(slide.slide_layout) in layout_set:
                slide_count += 1
        result.append(MasterInfo(
            name=name,
            layout_count=layout_count,
            slide_count=slide_count,
        ))
    return result


def list_layouts(prs_or_path: Any, *, master_index: int = 0) -> list[LayoutInfo]:
    """List all slide layouts for a given slide master.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        master_index: 0-based index of the slide master (default 0).

    Returns:
        A list of :class:`LayoutInfo` objects.

    Raises:
        IndexError: If *master_index* is out of range.
    """
    prs = _open_prs(prs_or_path)
    master = _get_master(prs, master_index)
    master_nm = _master_name(master)

    result: list[LayoutInfo] = []
    for layout in master.slide_layouts:
        name = _layout_name(layout)
        ph_types: list[str] = []
        ph_count = 0
        try:
            for ph in layout.placeholders:
                ph_count += 1
                ph_type = getattr(ph, "placeholder_format", None)
                if ph_type is not None:
                    type_name = getattr(ph_type, "type", None)
                    if type_name is not None:
                        ph_types.append(str(type_name))
                    else:
                        ph_types.append("obj")
                else:
                    ph_types.append("obj")
        except Exception:
            # Fallback: parse XML directly
            from lxml import etree

            xml_elem = layout._element
            cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
            if cSld is not None:
                spTree = cSld.find(f"{_P_NS_PREFIX}spTree")
                if spTree is not None:
                    for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
                        nvSpPr = sp.find(f"{_P_NS_PREFIX}nvSpPr")
                        if nvSpPr is not None:
                            nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
                            if nvPr is not None:
                                ph = nvPr.find(f"{_P_NS_PREFIX}ph")
                                if ph is not None:
                                    ph_count += 1
                                    ph_types.append(ph.get("type", "obj"))

        result.append(LayoutInfo(
            name=name,
            master_name=master_nm,
            placeholder_count=ph_count,
            placeholder_types=ph_types,
        ))
    return result


def get_layout_by_name(prs_or_path: Any, name: str) -> LayoutInfo | None:
    """Find a layout by name across all masters.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        name: The layout name to search for (case-sensitive).

    Returns:
        A :class:`LayoutInfo` if found, else ``None``.
    """
    prs = _open_prs(prs_or_path)
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if _layout_name(layout) == name:
                master_nm = _master_name(master)
                ph_types: list[str] = []
                ph_count = 0
                try:
                    for ph in layout.placeholders:
                        ph_count += 1
                        ph_type = getattr(ph, "placeholder_format", None)
                        if ph_type is not None:
                            type_name = getattr(ph_type, "type", None)
                            if type_name is not None:
                                ph_types.append(str(type_name))
                            else:
                                ph_types.append("obj")
                        else:
                            ph_types.append("obj")
                except Exception:
                    from lxml import etree

                    xml_elem = layout._element
                    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
                    if cSld is not None:
                        spTree = cSld.find(f"{_P_NS_PREFIX}spTree")
                        if spTree is not None:
                            for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
                                nvSpPr = sp.find(f"{_P_NS_PREFIX}nvSpPr")
                                if nvSpPr is not None:
                                    nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
                                    if nvPr is not None:
                                        ph = nvPr.find(f"{_P_NS_PREFIX}ph")
                                        if ph is not None:
                                            ph_count += 1
                                            ph_types.append(ph.get("type", "obj"))

                return LayoutInfo(
                    name=name,
                    master_name=master_nm,
                    placeholder_count=ph_count,
                    placeholder_types=ph_types,
                )
    return None


def analyze_layout(layout: Any) -> dict[str, Any]:
    """Analyze a layout's placeholders, positions, and sizes.

    Args:
        layout: A python-pptx ``SlideLayout`` object.

    Returns:
        A dict with keys ``name``, ``placeholders``, and ``master_name``.
        Each placeholder entry contains ``type``, ``index``, ``left``,
        ``top``, ``width``, ``height`` (all in EMU).
    """
    from lxml import etree

    name = _layout_name(layout)
    master = layout.slide_master
    master_nm = _master_name(master)

    placeholders: list[dict[str, Any]] = []

    # Try python-pptx API first
    try:
        for ph in layout.placeholders:
            entry: dict[str, Any] = {
                "type": "obj",
                "index": None,
                "left": ph.left,
                "top": ph.top,
                "width": ph.width,
                "height": ph.height,
            }
            ph_format = getattr(ph, "placeholder_format", None)
            if ph_format is not None:
                type_obj = getattr(ph_format, "type", None)
                if type_obj is not None:
                    entry["type"] = str(type_obj)
                idx = getattr(ph_format, "idx", None)
                if idx is not None:
                    entry["index"] = idx
            placeholders.append(entry)
    except Exception:
        # Fallback: parse XML directly
        xml_elem = layout._element
        cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
        if cSld is not None:
            spTree = cSld.find(f"{_P_NS_PREFIX}spTree")
            if spTree is not None:
                for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
                    nvSpPr = sp.find(f"{_P_NS_PREFIX}nvSpPr")
                    if nvSpPr is None:
                        continue
                    nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
                    if nvPr is None:
                        continue
                    ph_elem = nvPr.find(f"{_P_NS_PREFIX}ph")
                    if ph_elem is None:
                        continue

                    ph_type = ph_elem.get("type", "obj")
                    ph_idx = ph_elem.get("idx")
                    if ph_idx is not None:
                        try:
                            ph_idx = int(ph_idx)
                        except (ValueError, TypeError):
                            ph_idx = None

                    # Get position from spPr > xfrm
                    spPr = sp.find(f"{_P_NS_PREFIX}spPr")
                    left = top = width = height = 0
                    if spPr is not None:
                        xfrm = spPr.find(f"{_A_NS_PREFIX}xfrm")
                        if xfrm is not None:
                            off = xfrm.find(f"{_A_NS_PREFIX}off")
                            if off is not None:
                                left = int(off.get("x", "0"))
                                top = int(off.get("y", "0"))
                            ext = xfrm.find(f"{_A_NS_PREFIX}ext")
                            if ext is not None:
                                width = int(ext.get("cx", "0"))
                                height = int(ext.get("cy", "0"))

                    placeholders.append({
                        "type": ph_type,
                        "index": ph_idx,
                        "left": left,
                        "top": top,
                        "width": width,
                        "height": height,
                    })

    return {
        "name": name,
        "master_name": master_nm,
        "placeholders": placeholders,
    }


# ---------------------------------------------------------------------------
# Modification functions
# ---------------------------------------------------------------------------


def clone_layout(prs_or_path: Any, source_layout_name: str, new_name: str,
                 *, master_index: int = 0) -> LayoutInfo:
    """Clone an existing layout under a new name.

    The new layout is added to the same slide master as the source.  The
    source layout's XML is deep-copied and its relationship targets are
    re-wired so the clone is fully independent.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        source_layout_name: Name of the layout to clone.
        new_name: Name for the cloned layout.
        master_index: 0-based index of the slide master (default 0).

    Returns:
        A :class:`LayoutInfo` for the newly created layout.

    Raises:
        ValueError: If the source layout is not found or *new_name* already
            exists on the same master.
        IndexError: If *master_index* is out of range.
    """
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    prs = _open_prs(prs_or_path)
    path = _resolve_path(prs_or_path)
    master = _get_master(prs, master_index)
    master_nm = _master_name(master)

    # Find source layout
    source_layout = None
    for layout in master.slide_layouts:
        if _layout_name(layout) == source_layout_name:
            source_layout = layout
            break

    if source_layout is None:
        raise ValueError(
            f"Layout '{source_layout_name}' not found on master '{master_nm}'"
        )

    # Check for duplicate name
    for layout in master.slide_layouts:
        if _layout_name(layout) == new_name:
            raise ValueError(
                f"Layout '{new_name}' already exists on master '{master_nm}'"
            )

    # Deep-copy the source layout's XML element
    source_xml = source_layout._element
    clone_xml = deepcopy(source_xml)

    # Update the layout name in the XML
    cSld = clone_xml.find(f"{_P_NS_PREFIX}cSld")
    if cSld is not None:
        cSld.set("name", new_name)

    # Add the cloned layout part to the package.
    # We create a new SlideLayoutPart from the cloned XML element and
    # wire it into the package via the master part's relationships.
    master_part = master.part
    source_part = source_layout.part

    # Build the new part name (must be unique in the package)
    partname_str = str(source_part.partname)
    match = re.search(r"slideLayout(\d+)\.xml", partname_str)
    if match:
        base_num = int(match.group(1))
        # Find the next available number
        existing_nums = set()
        for lo in master.slide_layouts:
            lo_partname = str(lo.part.partname)
            m = re.search(r"slideLayout(\d+)\.xml", lo_partname)
            if m:
                existing_nums.add(int(m.group(1)))
        new_num = base_num
        while new_num in existing_nums:
            new_num += 1
        new_partname_str = partname_str.replace(
            f"slideLayout{match.group(1)}.xml",
            f"slideLayout{new_num}.xml",
        )
    else:
        new_partname_str = partname_str

    from pptx.opc.packuri import PackURI
    from pptx.parts.slide import SlideLayoutPart

    new_partname = PackURI(new_partname_str)

    # Create the new SlideLayoutPart from the cloned element
    new_part = SlideLayoutPart(
        new_partname,
        source_part.content_type,
        source_part.package,
        clone_xml,
    )

    # Copy relationships from source part to new part
    for rel in source_part.rels.values():
        new_part.relate_to(rel.target_part, rel.reltype)

    # Relate the master part to the new layout part
    master_part.relate_to(new_part, RT.SLIDE_LAYOUT)

    # Add the layout to the master's XML (sldLayoutIdLst)
    master_xml = master._element
    sldLayoutIdLst = master_xml.find(f"{_P_NS_PREFIX}sldLayoutIdLst")
    if sldLayoutIdLst is None:
        sldLayoutIdLst = etree.SubElement(master_xml, f"{_P_NS_PREFIX}sldLayoutIdLst")

    # Get the relationship ID for the new layout
    new_rId = None
    for rel in master_part.rels.values():
        if rel.target_part is new_part:
            new_rId = rel.rId
            break

    if new_rId is not None:
        # Find the max existing id to assign a new one
        max_id = 0
        for sldLayoutId in sldLayoutIdLst.findall(f"{_P_NS_PREFIX}sldLayoutId"):
            id_val = sldLayoutId.get("id")
            if id_val:
                try:
                    max_id = max(max_id, int(id_val))
                except (ValueError, TypeError):
                    pass

        sldLayoutId = etree.SubElement(sldLayoutIdLst, f"{_P_NS_PREFIX}sldLayoutId")
        sldLayoutId.set("id", str(max_id + 1))
        sldLayoutId.set(f"{{{_R_NS}}}id", new_rId)

    _save_prs(prs, path)

    # Re-open to get the fresh layout object
    if path is not None:
        prs = _open_prs(path)
        master = _get_master(prs, master_index)
        for layout in master.slide_layouts:
            if _layout_name(layout) == new_name:
                ph_types: list[str] = []
                ph_count = 0
                try:
                    for ph in layout.placeholders:
                        ph_count += 1
                        ph_type = getattr(ph, "placeholder_format", None)
                        if ph_type is not None:
                            type_name = getattr(ph_type, "type", None)
                            if type_name is not None:
                                ph_types.append(str(type_name))
                            else:
                                ph_types.append("obj")
                        else:
                            ph_types.append("obj")
                except Exception:
                    pass
                return LayoutInfo(
                    name=new_name,
                    master_name=master_nm,
                    placeholder_count=ph_count,
                    placeholder_types=ph_types,
                )

    # Fallback: return info from what we know (in-memory Presentation)
    ph_types_fb: list[str] = []
    ph_count_fb = 0
    spTree = cSld.find(f"{_P_NS_PREFIX}spTree") if cSld is not None else None
    if spTree is not None:
        for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
            nvSpPr = sp.find(f"{_P_NS_PREFIX}nvSpPr")
            if nvSpPr is not None:
                nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
                if nvPr is not None:
                    ph = nvPr.find(f"{_P_NS_PREFIX}ph")
                    if ph is not None:
                        ph_count_fb += 1
                        ph_types_fb.append(ph.get("type", "obj"))

    return LayoutInfo(
        name=new_name,
        master_name=master_nm,
        placeholder_count=ph_count_fb,
        placeholder_types=ph_types_fb,
    )


def delete_layout(prs_or_path: Any, layout_name: str, *,
                  master_index: int = 0) -> bool:
    """Delete a layout from its slide master.

    A layout that is currently in use by any slide cannot be deleted.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        layout_name: Name of the layout to delete.
        master_index: 0-based index of the slide master (default 0).

    Returns:
        ``True`` if the layout was deleted, ``False`` if it was not found.

    Raises:
        ValueError: If the layout is in use by one or more slides.
        IndexError: If *master_index* is out of range.
    """
    from lxml import etree

    prs = _open_prs(prs_or_path)
    path = _resolve_path(prs_or_path)
    master = _get_master(prs, master_index)

    # Find the layout
    target_layout = None
    for layout in master.slide_layouts:
        if _layout_name(layout) == layout_name:
            target_layout = layout
            break

    if target_layout is None:
        return False

    # Check if any slides use this layout
    use_count = _count_slides_using_layout(prs, target_layout)
    if use_count > 0:
        raise ValueError(
            f"Cannot delete layout '{layout_name}': "
            f"it is used by {use_count} slide(s). "
            f"Reassign those slides to a different layout first."
        )

    # Remove the layout from the master's sldLayoutIdLst
    master_xml = master._element
    sldLayoutIdLst = master_xml.find(f"{_P_NS_PREFIX}sldLayoutIdLst")
    if sldLayoutIdLst is not None:
        # Find the rId for this layout
        layout_part = target_layout.part
        rId = None
        for rel in master.part.rels.values():
            if rel.target_part is layout_part:
                rId = rel.rId
                break

        if rId is not None:
            # Remove the sldLayoutId entry
            for sldLayoutId in sldLayoutIdLst.findall(f"{_P_NS_PREFIX}sldLayoutId"):
                if sldLayoutId.get(f"{{{_R_NS}}}id") == rId:
                    sldLayoutIdLst.remove(sldLayoutId)
                    break

            # Drop the relationship
            master.part.drop_rel(rId)

    _save_prs(prs, path)
    return True


def rename_layout(prs_or_path: Any, old_name: str, new_name: str,
                  *, master_index: int = 0) -> LayoutInfo:
    """Rename a layout.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        old_name: Current name of the layout.
        new_name: New name for the layout.
        master_index: 0-based index of the slide master (default 0).

    Returns:
        A :class:`LayoutInfo` with the updated name.

    Raises:
        ValueError: If the layout is not found or *new_name* already exists.
        IndexError: If *master_index* is out of range.
    """
    from lxml import etree

    prs = _open_prs(prs_or_path)
    path = _resolve_path(prs_or_path)
    master = _get_master(prs, master_index)
    master_nm = _master_name(master)

    # Check for duplicate new_name
    for layout in master.slide_layouts:
        if _layout_name(layout) == new_name:
            raise ValueError(
                f"Layout '{new_name}' already exists on master '{master_nm}'"
            )

    # Find the layout
    target_layout = None
    for layout in master.slide_layouts:
        if _layout_name(layout) == old_name:
            target_layout = layout
            break

    if target_layout is None:
        raise ValueError(
            f"Layout '{old_name}' not found on master '{master_nm}'"
        )

    # Update the name in the XML
    xml_elem = target_layout._element
    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is not None:
        cSld.set("name", new_name)

    _save_prs(prs, path)

    # Build LayoutInfo
    ph_types: list[str] = []
    ph_count = 0
    try:
        for ph in target_layout.placeholders:
            ph_count += 1
            ph_type = getattr(ph, "placeholder_format", None)
            if ph_type is not None:
                type_name = getattr(ph_type, "type", None)
                if type_name is not None:
                    ph_types.append(str(type_name))
                else:
                    ph_types.append("obj")
            else:
                ph_types.append("obj")
    except Exception:
        pass

    return LayoutInfo(
        name=new_name,
        master_name=master_nm,
        placeholder_count=ph_count,
        placeholder_types=ph_types,
    )


# ---------------------------------------------------------------------------
# Placeholder operations
# ---------------------------------------------------------------------------


def add_placeholder_to_layout(layout: Any, ph_type: str,
                              left: float, top: float,
                              width: float, height: float,
                              *, ph_index: int | None = None) -> None:
    """Add a placeholder shape to a layout.

    Args:
        layout: A python-pptx ``SlideLayout`` object.
        ph_type: Placeholder type (e.g. ``"title"``, ``"body"``, ``"ctrTitle"``,
            ``"dt"``, ``"sldNum"``, ``"ftr"``, ``"hdr"``, ``"pic"``, ``"chart"``,
            ``"tbl"``, ``"dgm"``, ``"media"``).
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        ph_index: Optional placeholder index.  If ``None``, the next
            available index is used.

    Raises:
        ValueError: If *ph_type* is not a valid OOXML placeholder type.
    """
    from lxml import etree

    if ph_type not in _VALID_PH_TYPES:
        raise ValueError(
            f"Invalid placeholder type '{ph_type}'. "
            f"Valid types: {sorted(_VALID_PH_TYPES)}"
        )

    xml_elem = layout._element
    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is None:
        raise ValueError("Layout has no cSld element")
    spTree = cSld.find(f"{_P_NS_PREFIX}spTree")
    if spTree is None:
        raise ValueError("Layout has no spTree element")

    # Determine ph_index
    is_body = ph_type == "body"
    if ph_index is None:
        # Find the max existing idx
        max_idx = 0
        for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
            idx = _ph_index_from_element(sp)
            if idx is not None and idx > max_idx:
                max_idx = idx
        ph_index = max_idx + 1 if not is_body else 1

    # Get the next shape ID
    max_sp_id = 2  # minimum starting ID
    for sp in spTree.iter(f"{_P_NS_PREFIX}sp"):
        nvSpPr = sp.find(f"{_P_NS_PREFIX}nvSpPr")
        if nvSpPr is not None:
            cNvPr = nvSpPr.find(f"{_P_NS_PREFIX}cNvPr")
            if cNvPr is not None:
                sp_id = cNvPr.get("id")
                if sp_id:
                    try:
                        max_sp_id = max(max_sp_id, int(sp_id))
                    except (ValueError, TypeError):
                        pass

    new_sp_id = max_sp_id + 1

    # Convert points to EMU
    left_emu = _pt_to_emu(left)
    top_emu = _pt_to_emu(top)
    width_emu = _pt_to_emu(width)
    height_emu = _pt_to_emu(height)

    # Build the shape XML
    sp = etree.SubElement(spTree, f"{_P_NS_PREFIX}sp")

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f"{_P_NS_PREFIX}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{_P_NS_PREFIX}cNvPr")
    cNvPr.set("id", str(new_sp_id))
    cNvPr.set("name", f"{ph_type} {ph_index}")
    if ph_type not in ("dt", "sldNum", "ftr", "hdr"):
        cNvPr.set("hidden", "0")

    cNvSpPr = etree.SubElement(nvSpPr, f"{_P_NS_PREFIX}cNvSpPr")
    if is_body:
        cNvSpPr.set("txBox", "1")

    nvPr = etree.SubElement(nvSpPr, f"{_P_NS_PREFIX}nvPr")
    ph = etree.SubElement(nvPr, f"{_P_NS_PREFIX}ph")
    ph.set("type", ph_type)
    if not is_body:
        ph.set("idx", str(ph_index))

    # spPr
    spPr = etree.SubElement(sp, f"{_P_NS_PREFIX}spPr")
    xfrm = etree.SubElement(spPr, f"{_A_NS_PREFIX}xfrm")
    off = etree.SubElement(xfrm, f"{_A_NS_PREFIX}off")
    off.set("x", str(left_emu))
    off.set("y", str(top_emu))
    ext = etree.SubElement(xfrm, f"{_A_NS_PREFIX}ext")
    ext.set("cx", str(width_emu))
    ext.set("cy", str(height_emu))

    # Default text body for text-bearing placeholders
    if ph_type in ("title", "body", "ctrTitle", "subTitle"):
        txBody = etree.SubElement(sp, f"{_P_NS_PREFIX}txBody")
        bodyPr = etree.SubElement(txBody, f"{_A_NS_PREFIX}bodyPr")
        if ph_type == "title":
            bodyPr.set("anchor", "ctr")
        lstStyle = etree.SubElement(txBody, f"{_A_NS_PREFIX}lstStyle")
        p = etree.SubElement(txBody, f"{_A_NS_PREFIX}p")
        if ph_type == "title":
            endParaRPr = etree.SubElement(p, f"{_A_NS_PREFIX}endParaRPr")
            endParaRPr.set("lang", "en-US")


def remove_placeholder_from_layout(layout: Any, ph_index: int) -> None:
    """Remove a placeholder from a layout by its index.

    Args:
        layout: A python-pptx ``SlideLayout`` object.
        ph_index: The placeholder index (``idx`` attribute in the XML).

    Raises:
        ValueError: If no placeholder with the given index is found.
    """
    from lxml import etree

    xml_elem = layout._element
    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is None:
        raise ValueError("Layout has no cSld element")
    spTree = cSld.find(f"{_P_NS_PREFIX}spTree")
    if spTree is None:
        raise ValueError("Layout has no spTree element")

    # Find the placeholder with the matching index
    for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
        idx = _ph_index_from_element(sp)
        if idx == ph_index:
            spTree.remove(sp)
            return

    raise ValueError(
        f"No placeholder with index {ph_index} found in layout "
        f"'{_layout_name(layout)}'"
    )


def list_placeholders(layout: Any) -> list[dict[str, Any]]:
    """List all placeholders in a layout with their details.

    Args:
        layout: A python-pptx ``SlideLayout`` object.

    Returns:
        A list of dicts, each with keys ``type``, ``index``, ``left``,
        ``top``, ``width``, ``height`` (positions in EMU).
    """
    from lxml import etree

    result: list[dict[str, Any]] = []

    # Try python-pptx API first
    try:
        for ph in layout.placeholders:
            entry: dict[str, Any] = {
                "type": "obj",
                "index": None,
                "left": ph.left,
                "top": ph.top,
                "width": ph.width,
                "height": ph.height,
            }
            ph_format = getattr(ph, "placeholder_format", None)
            if ph_format is not None:
                type_obj = getattr(ph_format, "type", None)
                if type_obj is not None:
                    entry["type"] = str(type_obj)
                idx = getattr(ph_format, "idx", None)
                if idx is not None:
                    entry["index"] = idx
            result.append(entry)
        return result
    except Exception:
        pass

    # Fallback: parse XML directly
    xml_elem = layout._element
    cSld = xml_elem.find(f"{_P_NS_PREFIX}cSld")
    if cSld is None:
        return result
    spTree = cSld.find(f"{_P_NS_PREFIX}spTree")
    if spTree is None:
        return result

    for sp in spTree.findall(f"{_P_NS_PREFIX}sp"):
        nvSpPr = sp.find(f"{_P_NS_PREFIX}nvSpPr")
        if nvSpPr is None:
            continue
        nvPr = nvSpPr.find(f"{_P_NS_PREFIX}nvPr")
        if nvPr is None:
            continue
        ph_elem = nvPr.find(f"{_P_NS_PREFIX}ph")
        if ph_elem is None:
            continue

        ph_type = ph_elem.get("type", "obj")
        ph_idx = ph_elem.get("idx")
        if ph_idx is not None:
            try:
                ph_idx = int(ph_idx)
            except (ValueError, TypeError):
                ph_idx = None

        # Get position from spPr > xfrm
        spPr = sp.find(f"{_P_NS_PREFIX}spPr")
        left = top = width = height = 0
        if spPr is not None:
            xfrm = spPr.find(f"{_A_NS_PREFIX}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{_A_NS_PREFIX}off")
                if off is not None:
                    left = int(off.get("x", "0"))
                    top = int(off.get("y", "0"))
                ext = xfrm.find(f"{_A_NS_PREFIX}ext")
                if ext is not None:
                    width = int(ext.get("cx", "0"))
                    height = int(ext.get("cy", "0"))

        result.append({
            "type": ph_type,
            "index": ph_idx,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        })

    return result


# ---------------------------------------------------------------------------
# Background operations
# ---------------------------------------------------------------------------


def set_master_background(prs_or_path: Any, *, color: str | None = None,
                          image_path: str | None = None,
                          gradient: dict[str, Any] | None = None,
                          master_index: int = 0) -> None:
    """Set the background of a slide master.

    Only one fill type should be specified.  If multiple are given, the
    first non-None value wins in order: color > image_path > gradient.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        color: Hex color string (e.g. ``"FF0000"`` or ``"#FF0000"``).
        image_path: Path to an image file for the background.
        gradient: Dict with gradient parameters.  Supported keys:
            ``stops`` (list of ``{"position": int, "color": str}``),
            ``angle`` (int, in 60000ths of a degree; default 5400000 = 90deg).
        master_index: 0-based index of the slide master (default 0).

    Raises:
        IndexError: If *master_index* is out of range.
        FileNotFoundError: If *image_path* does not exist.
    """
    prs = _open_prs(prs_or_path)
    path = _resolve_path(prs_or_path)
    master = _get_master(prs, master_index)

    if image_path is not None and color is None and gradient is None:
        # Image background requires part access
        img = Path(image_path)
        if not img.exists():
            raise FileNotFoundError(f"Image file not found: {image_path}")
        _set_image_background_on_part(master.part, master._element, str(img))
    else:
        _set_background_on_element(
            master._element, color=color, gradient=gradient,
        )

    _save_prs(prs, path)


def set_layout_background(prs_or_path: Any, layout_name: str, *,
                          color: str | None = None,
                          image_path: str | None = None,
                          gradient: dict[str, Any] | None = None) -> None:
    """Set the background of a specific layout.

    Only one fill type should be specified.  If multiple are given, the
    first non-None value wins in order: color > image_path > gradient.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        layout_name: Name of the layout to modify.
        color: Hex color string (e.g. ``"FF0000"`` or ``"#FF0000"``).
        image_path: Path to an image file for the background.
        gradient: Dict with gradient parameters.  See :func:`set_master_background`.
    """
    prs = _open_prs(prs_or_path)
    path = _resolve_path(prs_or_path)

    # Find the layout across all masters
    target_layout = None
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if _layout_name(layout) == layout_name:
                target_layout = layout
                break
        if target_layout is not None:
            break

    if target_layout is None:
        raise ValueError(f"Layout '{layout_name}' not found in presentation")

    if image_path is not None and color is None and gradient is None:
        img = Path(image_path)
        if not img.exists():
            raise FileNotFoundError(f"Image file not found: {image_path}")
        _set_image_background_on_part(
            target_layout.part, target_layout._element, str(img),
        )
    else:
        _set_background_on_element(
            target_layout._element, color=color, gradient=gradient,
        )

    _save_prs(prs, path)
