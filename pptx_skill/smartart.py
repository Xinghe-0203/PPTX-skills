"""SmartArt detection, text extraction, editing, and preservation for PPTX.

python-pptx does not support SmartArt natively (issue #83 open since 2014).
SmartArt in OOXML is a ``graphicFrame`` containing ``graphicData`` with
``dgm:relIds`` pointing to four related parts:

- **diagramData**  (rId2) — the actual text/content in ``<dgm:dataModel>``
- **diagramLayout** (rId3) — the layout definition
- **diagramQuickStyle** (rId4) — the visual style
- **diagramColors** (rId5) — the color mapping

This module provides:

- :func:`detect_smartart`          -- find SmartArt shapes on a slide
- :func:`extract_smartart_text`    -- read text from a SmartArt shape
- :func:`populate_smartart_text`   -- edit SmartArt text in place
- :func:`preserve_smartart`        -- copy SmartArt with all parts to another slide
- :func:`list_smartart_layouts`    -- list known layout type identifiers
- :func:`smartart_layout_info`     -- get info about a specific layout type
"""
from __future__ import annotations

import copy
import logging
from typing import Any

from pptx_skill.constants import A_NS as _A_NS

__all__ = [
    "detect_smartart",
    "extract_smartart_text",
    "list_smartart_layouts",
    "populate_smartart_text",
    "preserve_smartart",
    "smartart_layout_info",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------
_DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

# Relationship types for the four SmartArt parts
_DGM_DATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
_DGM_LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
_DGM_STYLE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
_DGM_COLORS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"

# All four diagram relationship types as a set for quick membership checks
_DGM_REL_TYPES = frozenset({_DGM_DATA_REL, _DGM_LAYOUT_REL, _DGM_STYLE_REL, _DGM_COLORS_REL})

# ---------------------------------------------------------------------------
# Known SmartArt layout type identifiers
# ---------------------------------------------------------------------------
_SMARTART_LAYOUTS: dict[str, dict[str, str]] = {
    # Process layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/process1": {
        "category": "Process",
        "name": "Basic Chevron Process",
        "description": "A chevron-style horizontal process with sequential steps.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process2": {
        "category": "Process",
        "name": "Closed Chevron Process",
        "description": "A circular chevron process showing continuous flow.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process3": {
        "category": "Process",
        "name": "Accent Process",
        "description": "A process with accent circles for each step.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process4": {
        "category": "Process",
        "name": "Continuous Block Process",
        "description": "A horizontal block process with continuous arrows.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process5": {
        "category": "Process",
        "name": "Ascending Picture Accent Process",
        "description": "An ascending process with picture accents.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process6": {
        "category": "Process",
        "name": "Converging Arrows",
        "description": "Arrows converging to a central point.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process7": {
        "category": "Process",
        "name": "Diverging Arrows",
        "description": "Arrows diverging from a central point.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process8": {
        "category": "Process",
        "name": "Basic Timeline",
        "description": "A horizontal timeline with sequential events.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/process9": {
        "category": "Process",
        "name": "Segmented Timeline",
        "description": "A timeline divided into segments.",
    },
    # Hierarchy layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1": {
        "category": "Hierarchy",
        "name": "Organization Chart",
        "description": "A standard top-down organizational hierarchy.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy2": {
        "category": "Hierarchy",
        "name": "Hierarchy",
        "description": "A top-down hierarchy without assistant nodes.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy3": {
        "category": "Hierarchy",
        "name": "Horizontal Hierarchy",
        "description": "A left-to-right hierarchy layout.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy4": {
        "category": "Hierarchy",
        "name": "Horizontal Organization Chart",
        "description": "A left-to-right organizational chart.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy5": {
        "category": "Hierarchy",
        "name": "Name and Title Organization Chart",
        "description": "An org chart with name and title fields.",
    },
    # Cycle layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle1": {
        "category": "Cycle",
        "name": "Basic Cycle",
        "description": "A circular cycle with text in each segment.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle2": {
        "category": "Cycle",
        "name": "Text Cycle",
        "description": "A cycle with text wrapping around a circle.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle3": {
        "category": "Cycle",
        "name": "Block Cycle",
        "description": "A cycle with block-shaped segments.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/cycle4": {
        "category": "Cycle",
        "name": "Radial Cycle",
        "description": "A radial cycle with a central hub.",
    },
    # Relationship layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/rel1": {
        "category": "Relationship",
        "name": "Basic Venn",
        "description": "Overlapping circles showing relationships.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/rel2": {
        "category": "Relationship",
        "name": "Radial Cluster",
        "description": "A central concept with radial connections.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/rel3": {
        "category": "Relationship",
        "name": "Converging Radial",
        "description": "Radial elements converging to a center.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/rel4": {
        "category": "Relationship",
        "name": "Diverging Radial",
        "description": "A center element diverging outward.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/rel5": {
        "category": "Relationship",
        "name": "Balance",
        "description": "A balance scale showing opposing factors.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/rel6": {
        "category": "Relationship",
        "name": "Funnel",
        "description": "A funnel showing filtering or narrowing.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/rel7": {
        "category": "Relationship",
        "name": "Gear",
        "description": "Interlocking gears showing interrelated processes.",
    },
    # Matrix layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/matrix1": {
        "category": "Matrix",
        "name": "Basic Matrix",
        "description": "A 2x2 matrix with four quadrants.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/matrix2": {
        "category": "Matrix",
        "name": "Titled Matrix",
        "description": "A matrix with a title and quadrant labels.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/matrix3": {
        "category": "Matrix",
        "name": "Grid Matrix",
        "description": "A grid-based matrix layout.",
    },
    # Pyramid layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1": {
        "category": "Pyramid",
        "name": "Basic Pyramid",
        "description": "A standard pyramid with hierarchical levels.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid2": {
        "category": "Pyramid",
        "name": "Inverted Pyramid",
        "description": "An inverted pyramid layout.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/pyramid3": {
        "category": "Pyramid",
        "name": "Segmented Pyramid",
        "description": "A pyramid with distinct segment divisions.",
    },
    # Picture layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/picture1": {
        "category": "Picture",
        "name": "Picture Strips",
        "description": "Horizontal strips with pictures and captions.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/picture2": {
        "category": "Picture",
        "name": "Snapshot Picture List",
        "description": "A list with snapshot pictures beside text.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/picture3": {
        "category": "Picture",
        "name": "Framed Picture Matrix",
        "description": "A matrix of framed pictures with labels.",
    },
    # List layouts
    "urn:microsoft.com/office/officeart/2005/8/layout/list1": {
        "category": "List",
        "name": "Basic Block List",
        "description": "A vertical list of blocks with text.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/list2": {
        "category": "List",
        "name": "Vertical Box List",
        "description": "A vertical list with box containers.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/list3": {
        "category": "List",
        "name": "Vertical Chevron List",
        "description": "A vertical list with chevron markers.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/list4": {
        "category": "List",
        "name": "Vertical Bullet List",
        "description": "A vertical list with bullet markers.",
    },
    "urn:microsoft.com/office/officeart/2005/8/layout/list5": {
        "category": "List",
        "name": "Stacked List",
        "description": "A stacked horizontal list layout.",
    },
}


# ---------------------------------------------------------------------------
# Internal helpers — SmartArt shape identification
# ---------------------------------------------------------------------------
def _is_smartart_shape(shape: Any) -> bool:
    """Return True if *shape* is a SmartArt graphicFrame.

    A SmartArt shape is a ``graphicFrame`` whose ``graphicData`` element has
    a URI attribute starting with the diagram namespace prefix.
    """
    from pptx.oxml.ns import qn

    elem = getattr(shape, "_element", None)
    if elem is None:
        return False

    # SmartArt lives inside a <p:graphicFrame> element
    graphic_frame = elem.find(qn("p:graphicFrame"))
    if graphic_frame is None:
        # The shape element itself might be the graphicFrame
        tag_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag_local != "graphicFrame":
            return False
        graphic_frame = elem

    # Look for <a:graphic><a:graphicData> with diagram URI
    graphic = graphic_frame.find(qn("a:graphic"))
    if graphic is None:
        return False

    graphic_data = graphic.find(qn("a:graphicData"))
    if graphic_data is None:
        return False

    uri = graphic_data.get("uri", "")
    return "diagram" in uri or "dgm" in uri


def _get_graphic_frame_element(shape: Any) -> Any | None:
    """Return the ``<p:graphicFrame>`` element for a SmartArt shape, or None."""
    from pptx.oxml.ns import qn

    elem = getattr(shape, "_element", None)
    if elem is None:
        return None

    graphic_frame = elem.find(qn("p:graphicFrame"))
    if graphic_frame is not None:
        return graphic_frame

    tag_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
    if tag_local == "graphicFrame":
        return elem

    return None


def _get_rel_ids(shape: Any) -> dict[str, str] | None:
    """Extract the four diagram relationship IDs from a SmartArt shape.

    Returns a dict mapping keys ``data``, ``layout``, ``style``, ``colors``
    to their rId values, or None if the shape is not SmartArt.
    """
    from pptx.oxml.ns import qn

    gf = _get_graphic_frame_element(shape)
    if gf is None:
        return None

    graphic = gf.find(qn("a:graphic"))
    if graphic is None:
        return None

    graphic_data = graphic.find(qn("a:graphicData"))
    if graphic_data is None:
        return None

    # Find <dgm:relIds> element
    rel_ids_elem = None
    for child in graphic_data:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "relIds":
            rel_ids_elem = child
            break

    if rel_ids_elem is None:
        return None

    result: dict[str, str] = {}
    # The relIds element has attributes: r:dm (data), r:lo (layout), r:qs (style), r:cs (colors)
    dm = rel_ids_elem.get(qn("r:dm"))
    lo = rel_ids_elem.get(qn("r:lo"))
    qs = rel_ids_elem.get(qn("r:qs"))
    cs = rel_ids_elem.get(qn("r:cs"))

    if dm is not None:
        result["data"] = dm
    if lo is not None:
        result["layout"] = lo
    if qs is not None:
        result["style"] = qs
    if cs is not None:
        result["colors"] = cs

    return result if result else None


def _get_diagram_data_part(shape: Any) -> Any | None:
    """Return the diagramData part for a SmartArt shape, or None."""
    rel_ids = _get_rel_ids(shape)
    if rel_ids is None or "data" not in rel_ids:
        return None

    slide_part = getattr(shape, "part", None)
    if slide_part is None:
        return None

    try:
        return slide_part.related_part(rel_ids["data"])
    except Exception:
        return None


def _get_layout_type_from_part(diagram_layout_part: Any) -> str:
    """Extract the layout type URI from a diagramLayoutDefinition part.

    Returns the ``urn`` attribute of the ``<dgm:layoutDef>`` root element,
    or an empty string if not found.
    """
    try:
        element = getattr(diagram_layout_part, "_element", None)
        if element is None:
            return ""
        # The root element should be <dgm:layoutDef> with a "uri" attribute
        uri = element.get("uri", "")
        if uri:
            return uri
        # Also check for the unique identifier attribute
        uri = element.get("uniqueId", "")
        return uri
    except Exception:
        return ""


def _count_nodes_in_data(diagram_data_element: Any) -> int:
    """Count the number of point nodes in a diagramData element.

    Counts ``<dgm:pt>`` elements inside ``<dgm:ptLst>`` that contain
    text (i.e. have a ``<dgm:t>`` child).
    """
    try:
        # Find <dgm:ptLst>
        pt_lst = None
        for child in diagram_data_element:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag_local == "ptLst":
                pt_lst = child
                break

        if pt_lst is None:
            return 0

        count = 0
        for pt in pt_lst:
            tag_local = pt.tag.split("}")[-1] if "}" in pt.tag else pt.tag
            if tag_local == "pt":
                # Check if this point has a <dgm:t> child (text node)
                has_text = False
                for pt_child in pt:
                    pt_child_tag = pt_child.tag.split("}")[-1] if "}" in pt_child.tag else pt_child.tag
                    if pt_child_tag == "t":
                        has_text = True
                        break
                if has_text:
                    count += 1

        return count
    except Exception:
        return 0


def _has_fallback_image(shape: Any) -> bool:
    """Check if a SmartArt shape has a fallback image (blipFill).

    Some SmartArt shapes include a fallback image for applications that
    cannot render SmartArt.  This is stored as a ``<p:blipFill>`` element
    inside the graphicFrame.
    """
    from pptx.oxml.ns import qn

    gf = _get_graphic_frame_element(shape)
    if gf is None:
        return False

    # Check for <p:blipFill> or <p:txBody> with image
    blip_fill = gf.find(qn("p:blipFill"))
    return blip_fill is not None


# ---------------------------------------------------------------------------
# Public API — detect_smartart
# ---------------------------------------------------------------------------
def detect_smartart(slide: Any) -> list[dict[str, Any]]:
    """Detect SmartArt shapes on a slide.

    Parameters
    ----------
    slide : Slide
        A python-pptx ``Slide`` object.

    Returns
    -------
    list[dict]
        A list of dicts, one per SmartArt shape found.  Each dict contains:

        - ``shape_name`` (str): The name attribute of the shape.
        - ``shape_index`` (int): The 0-based index of the shape in the slide's
          shape collection.
        - ``layout_type`` (str): The layout type URI, or ``""`` if unknown.
        - ``node_count`` (int): Number of text-bearing nodes in the diagram data.
        - ``has_fallback_image`` (bool): Whether a fallback image is present.
    """
    results: list[dict[str, Any]] = []

    for idx, shape in enumerate(slide.shapes):
        if not _is_smartart_shape(shape):
            continue

        shape_name = getattr(shape, "name", f"SmartArt_{idx}")

        # Get layout type
        layout_type = ""
        rel_ids = _get_rel_ids(shape)
        if rel_ids and "layout" in rel_ids:
            slide_part = getattr(shape, "part", None)
            if slide_part is not None:
                try:
                    layout_part = slide_part.related_part(rel_ids["layout"])
                    layout_type = _get_layout_type_from_part(layout_part)
                except Exception:
                    log.debug("Could not read layout part for shape %s", shape_name)

        # Get node count from diagram data
        node_count = 0
        data_part = _get_diagram_data_part(shape)
        if data_part is not None:
            data_element = getattr(data_part, "_element", None)
            if data_element is not None:
                node_count = _count_nodes_in_data(data_element)

        has_fallback = _has_fallback_image(shape)

        results.append({
            "shape_name": shape_name,
            "shape_index": idx,
            "layout_type": layout_type,
            "node_count": node_count,
            "has_fallback_image": has_fallback,
        })

    return results


# ---------------------------------------------------------------------------
# Public API — extract_smartart_text
# ---------------------------------------------------------------------------
def extract_smartart_text(shape: Any) -> list[str]:
    """Extract text from a SmartArt shape's diagramData XML.

    Parameters
    ----------
    shape : Shape
        A python-pptx shape object that is a SmartArt graphicFrame.

    Returns
    -------
    list[str]
        A list of text strings, one per node in the diagram data.
        Nodes without text content produce an empty string in the list.
        Returns an empty list if the shape is not SmartArt or has no
        diagram data.
    """
    if not _is_smartart_shape(shape):
        return []

    data_part = _get_diagram_data_part(shape)
    if data_part is None:
        return []

    data_element = getattr(data_part, "_element", None)
    if data_element is None:
        return []

    texts: list[str] = []

    try:
        # Find <dgm:ptLst>
        pt_lst = None
        for child in data_element:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag_local == "ptLst":
                pt_lst = child
                break

        if pt_lst is None:
            return []

        for pt in pt_lst:
            tag_local = pt.tag.split("}")[-1] if "}" in pt.tag else pt.tag
            if tag_local != "pt":
                continue

            # Find <dgm:t> inside this <dgm:pt>
            t_elem = None
            for pt_child in pt:
                pt_child_tag = pt_child.tag.split("}")[-1] if "}" in pt_child.tag else pt_child.tag
                if pt_child_tag == "t":
                    t_elem = pt_child
                    break

            if t_elem is None:
                continue

            # Extract all <a:t> text from <dgm:t>
            # <dgm:t> contains <a:r><a:rPr>...<a:t>text</a:t></a:r> runs
            node_text_parts: list[str] = []
            for elem in t_elem.iter():
                elem_tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if elem_tag == "t" and elem.text:
                    # Only collect <a:t> elements (not the <dgm:t> wrapper itself)
                    parent_tag = ""
                    parent = elem.getparent()
                    if parent is not None:
                        parent_tag = parent.tag.split("}")[-1] if "}" in parent.tag else parent.tag
                    if parent_tag == "r":
                        node_text_parts.append(elem.text)

            # If no <a:r>/<a:t> runs found, try direct text content
            if not node_text_parts:
                direct_text = t_elem.text
                if direct_text and direct_text.strip():
                    node_text_parts.append(direct_text.strip())

            texts.append("".join(node_text_parts))

    except Exception as exc:
        log.warning("Error extracting SmartArt text: %s", exc)
        return []

    return texts


# ---------------------------------------------------------------------------
# Public API — populate_smartart_text
# ---------------------------------------------------------------------------
def populate_smartart_text(shape: Any, node_texts: list[str]) -> None:
    """Edit SmartArt text in place by modifying the diagramData XML.

    Parameters
    ----------
    shape : Shape
        A python-pptx shape object that is a SmartArt graphicFrame.
    node_texts : list[str]
        A list of strings, one per node.  Nodes are matched by position
        (the *i*-th string replaces the text of the *i*-th ``<dgm:pt>``
        element that contains a ``<dgm:t>`` child).  If the list is
        shorter than the number of nodes, remaining nodes are left
        unchanged.  Extra strings in the list are ignored.

    Raises
    ------
    ValueError
        If the shape is not a SmartArt shape.
    """
    from lxml import etree

    if not _is_smartart_shape(shape):
        raise ValueError("Shape is not a SmartArt graphicFrame")

    data_part = _get_diagram_data_part(shape)
    if data_part is None:
        raise ValueError("Cannot access diagramData part for this SmartArt shape")

    data_element = getattr(data_part, "_element", None)
    if data_element is None:
        raise ValueError("diagramData part has no XML element")

    try:
        # Find <dgm:ptLst>
        pt_lst = None
        for child in data_element:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag_local == "ptLst":
                pt_lst = child
                break

        if pt_lst is None:
            raise ValueError("diagramData has no ptLst element")

        text_idx = 0
        for pt in pt_lst:
            tag_local = pt.tag.split("}")[-1] if "}" in pt.tag else pt.tag
            if tag_local != "pt":
                continue

            # Find <dgm:t> inside this <dgm:pt>
            t_elem = None
            for pt_child in pt:
                pt_child_tag = pt_child.tag.split("}")[-1] if "}" in pt_child.tag else pt_child.tag
                if pt_child_tag == "t":
                    t_elem = pt_child
                    break

            if t_elem is None:
                continue

            if text_idx >= len(node_texts):
                break

            new_text = node_texts[text_idx]
            text_idx += 1

            # Replace text in all <a:t> elements within <dgm:t>
            a_t_elements = []
            for elem in t_elem.iter():
                elem_tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if elem_tag == "t":
                    parent = elem.getparent()
                    if parent is not None:
                        parent_tag = parent.tag.split("}")[-1] if "}" in parent.tag else parent.tag
                        if parent_tag == "r":
                            a_t_elements.append(elem)

            if a_t_elements:
                # Put all text in the first <a:t>, clear the rest
                a_t_elements[0].text = new_text
                for extra_t in a_t_elements[1:]:
                    extra_t.getparent().remove(extra_t)
            else:
                # No <a:r>/<a:t> runs exist; create one
                a_r = etree.SubElement(t_elem, f"{{{_A_NS}}}r")
                a_rPr = etree.SubElement(a_r, f"{{{_A_NS}}}rPr")
                a_rPr.set("lang", "en-US")
                a_rPr.set("dirty", "0")
                a_t = etree.SubElement(a_r, f"{{{_A_NS}}}t")
                a_t.text = new_text

    except ValueError:
        raise
    except Exception as exc:
        log.error("Error populating SmartArt text: %s", exc)
        raise RuntimeError(f"Failed to populate SmartArt text: {exc}") from exc


# ---------------------------------------------------------------------------
# Public API — preserve_smartart
# ---------------------------------------------------------------------------
def preserve_smartart(shape: Any, target_slide: Any) -> bool:
    """Copy a SmartArt shape with all 4 related parts to a target slide.

    This handles the full relationship graph: the four diagram parts
    (data, layout, style, colors) are copied to the target presentation
    and new relationship IDs are assigned.  The shape XML is deep-copied
    with remapped rIds.

    Parameters
    ----------
    shape : Shape
        A python-pptx shape object that is a SmartArt graphicFrame on the
        source slide.
    target_slide : Slide
        The target slide to copy the SmartArt into.

    Returns
    -------
    bool
        True if the SmartArt was successfully preserved, False if the
        shape is not SmartArt or an error occurred.
    """
    from pptx.oxml.ns import qn

    if not _is_smartart_shape(shape):
        return False

    try:
        source_slide_part = getattr(shape, "part", None)
        if source_slide_part is None:
            return False

        target_slide_part = target_slide.part

        # Get the four relationship IDs from the SmartArt shape
        rel_ids = _get_rel_ids(shape)
        if rel_ids is None:
            log.warning("SmartArt shape has no diagram relIds, cannot preserve")
            return False

        # Build a mapping from old rId to new rId for each diagram part
        rid_map: dict[str, str] = {}

        # Map of rel_id key -> relationship type
        rel_key_to_type = {
            "data": _DGM_DATA_REL,
            "layout": _DGM_LAYOUT_REL,
            "style": _DGM_STYLE_REL,
            "colors": _DGM_COLORS_REL,
        }

        for key, rel_type in rel_key_to_type.items():
            old_rid = rel_ids.get(key)
            if old_rid is None:
                continue

            try:
                source_part = source_slide_part.related_part(old_rid)
            except Exception as exc:
                log.warning(
                    "Could not access diagram part %s (rId=%s): %s",
                    key, old_rid, exc,
                )
                continue

            # Add the part to the target slide's relationships
            try:
                new_rid = target_slide_part.rels.get_or_add(rel_type, source_part)
                rid_map[old_rid] = new_rid
            except Exception as exc:
                log.warning(
                    "Could not add diagram part %s to target slide: %s",
                    key, exc,
                )
                return False

        # Deep-copy the shape XML and remap rIds
        gf = _get_graphic_frame_element(shape)
        if gf is None:
            return False

        cloned_gf = copy.deepcopy(gf)

        # Remap the four diagram rIds in <dgm:relIds>
        # Attributes: r:dm (data), r:lo (layout), r:qs (style), r:cs (colors)
        attr_to_key = {
            qn("r:dm"): "data",
            qn("r:lo"): "layout",
            qn("r:qs"): "style",
            qn("r:cs"): "colors",
        }

        for elem in cloned_gf.iter():
            for attr, _key in attr_to_key.items():
                old_val = elem.get(attr)
                if old_val is not None and old_val in rid_map:
                    elem.set(attr, rid_map[old_val])

        # Also remap any r:embed, r:id, r:link attributes (for fallback images etc.)
        _remap_general_rids(cloned_gf, rid_map)

        # Insert the cloned shape into the target slide's shape tree
        target_slide.shapes._spTree.insert_element_before(cloned_gf, "p:extLst")

        return True

    except Exception as exc:
        log.error("Error preserving SmartArt: %s", exc)
        return False


def _remap_general_rids(element: Any, rid_map: dict[str, str]) -> None:
    """Walk *element* and remap general relationship ID attributes.

    Handles ``r:embed``, ``r:id``, and ``r:link`` attributes that may
    appear in fallback images or other sub-elements of a SmartArt shape.
    """
    from pptx.oxml.ns import qn

    rel_attrs = frozenset({qn("r:embed"), qn("r:id"), qn("r:link")})
    for node in element.iter():
        for attr in rel_attrs:
            old = node.get(attr)
            if old is not None and old in rid_map:
                node.set(attr, rid_map[old])


# ---------------------------------------------------------------------------
# Public API — list_smartart_layouts
# ---------------------------------------------------------------------------
def list_smartart_layouts() -> list[str]:
    """Return known SmartArt layout type identifiers.

    Returns
    -------
    list[str]
        A list of layout type URIs (at least 30 common ones).  These are
        the ``uri`` or ``uniqueId`` values found in
        ``<dgm:layoutDef>`` elements.
    """
    return list(_SMARTART_LAYOUTS.keys())


# ---------------------------------------------------------------------------
# Public API — smartart_layout_info
# ---------------------------------------------------------------------------
def smartart_layout_info(layout_type: str) -> dict[str, str] | None:
    """Get information about a specific SmartArt layout type.

    Parameters
    ----------
    layout_type : str
        The layout type URI (e.g. the value from ``detect_smartart()``'s
        ``layout_type`` field).

    Returns
    -------
    dict or None
        A dict with keys ``category``, ``name``, and ``description``,
        or None if the layout type is not recognized.
    """
    return _SMARTART_LAYOUTS.get(layout_type)
