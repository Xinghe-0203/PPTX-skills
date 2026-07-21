"""Shape grouping and ungrouping for PowerPoint slides.

Creates ``<p:grpSp>`` (group shape) elements and provides APIs for
grouping, ungrouping, and manipulating grouped shapes.

OOXML reference: ECMA-376 Part 4, §19.3.1.25 (p:grpSp).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

__all__ = [
    "GroupInfo",
    "group_shapes",
    "ungroup_shapes",
    "list_groups",
    "list_group_children",
    "add_to_group",
    "remove_from_group",
    "move_in_group",
]

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class GroupInfo:
    """Info about a group shape."""
    name: str = ""
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0
    child_count: int = 0
    child_names: list[str] = None

    def __post_init__(self):
        if self.child_names is None:
            self.child_names = []


def _is_presentation(obj) -> bool:
    try:
        from pptx import Presentation
        return isinstance(obj, Presentation)
    except ImportError:
        return hasattr(obj, "slides")


def _open_prs(prs_or_path):
    if _is_presentation(prs_or_path):
        return prs_or_path, False
    from pptx import Presentation
    return Presentation(prs_or_path), True


def _save_prs(prs, path, is_path):
    if is_path and path:
        prs.save(path)


def _find_shape(slide, shape_name: str):
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None


def _compute_group_bounds(shape_elements):
    """Compute bounding box that encompasses all shapes."""
    min_x = min_y = float("inf")
    max_x = max_y = float("-inf")

    for elem in shape_elements:
        xfrm = elem.find(f".//{{{_NS_A}}}xfrm")
        if xfrm is not None:
            off = xfrm.find(f"{{{_NS_A}}}off")
            ext = xfrm.find(f"{{{_NS_A}}}ext")
            if off is not None and ext is not None:
                x = int(off.get("x", "0"))
                y = int(off.get("y", "0"))
                cx = int(ext.get("cx", "0"))
                cy = int(ext.get("cy", "0"))
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x + cx)
                max_y = max(max_y, y + cy)

    if min_x == float("inf"):
        min_x = min_y = 0
        max_x = max_y = 1

    return int(min_x), int(min_y), int(max_x - min_x), int(max_y - min_y)


def group_shapes(prs_or_path, slide_index: int, *,
                 shape_names: list[str],
                 group_name: str | None = None) -> str:
    """Group multiple shapes into a group shape.

    Parameters
    ----------
    shape_names : list of str
        Names of shapes to group. At least 2 shapes required.
    group_name : str, optional
        Name for the group shape.

    Returns
    -------
    str
        The group shape name.

    Raises
    ------
    ValueError
        If fewer than 2 shapes are provided or any shape is not found.
    """
    from lxml import etree

    if len(shape_names) < 2:
        raise ValueError("At least 2 shapes required for grouping")

    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        sp_tree = slide.shapes._spTree

        # Find shape elements
        shape_elements = []
        for name in shape_names:
            shape = _find_shape(slide, name)
            if shape is None:
                raise ValueError(f"Shape not found: {name!r}")
            shape_elements.append(shape._element)

        # Compute group bounds
        gx, gy, gcx, gcy = _compute_group_bounds(shape_elements)

        _name = group_name or f"Group {len(slide.shapes)}"

        # Create group shape element
        grpSp = etree.SubElement(sp_tree, f"{{{_NS_P}}}grpSp")

        # nvGrpSpPr
        nvGrpSpPr = etree.SubElement(grpSp, f"{{{_NS_P}}}nvGrpSpPr")
        cNvPr = etree.SubElement(nvGrpSpPr, f"{{{_NS_P}}}cNvPr")
        cNvPr.set("id", "0")
        cNvPr.set("name", _name)
        etree.SubElement(nvGrpSpPr, f"{{{_NS_P}}}cNvGrpSpPr")
        etree.SubElement(nvGrpSpPr, f"{{{_NS_P}}}nvPr")

        # grpSpPr
        grpSpPr = etree.SubElement(grpSp, f"{{{_NS_P}}}grpSpPr")

        # Group transform
        xfrm = etree.SubElement(grpSpPr, f"{{{_NS_A}}}xfrm")
        off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
        off.set("x", str(gx))
        off.set("y", str(gy))
        ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
        ext.set("cx", str(gcx))
        ext.set("cy", str(gcy))

        # Child offset (chOff) — same as group offset for simplicity
        chOff = etree.SubElement(xfrm, f"{{{_NS_A}}}chOff")
        chOff.set("x", str(gx))
        chOff.set("y", str(gy))

        # Child extent (chExt)
        chExt = etree.SubElement(xfrm, f"{{{_NS_A}}}chExt")
        chExt.set("cx", str(gcx))
        chExt.set("cy", str(gcy))

        # Move child shapes into the group
        for elem in shape_elements:
            sp_tree.remove(elem)
            grpSp.append(elem)

        return _name
    finally:
        _save_prs(prs, prs_or_path if is_path else None, is_path)


def ungroup_shapes(prs_or_path, slide_index: int, shape_name: str) -> list[str]:
    """Ungroup a group shape, returning its children to the slide.

    Parameters
    ----------
    shape_name : str
        Name of the group shape to ungroup.

    Returns
    -------
    list of str
        Names of the child shapes that were ungrouped.
    """
    from lxml import etree

    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        sp_tree = slide.shapes._spTree

        # Find the group shape
        group = _find_shape(slide, shape_name)
        if group is None:
            return []

        grpSp = group._element

        # Extract child shapes
        child_names = []
        children = []
        for child in list(grpSp):
            tag = child.tag
            # Skip nvGrpSpPr and grpSpPr
            if tag == f"{{{_NS_P}}}nvGrpSpPr" or tag == f"{{{_NS_P}}}grpSpPr":
                continue
            children.append(child)
            # Get child name
            cNvPr = child.find(f".//{{{_NS_P}}}cNvPr")
            if cNvPr is None:
                cNvPr = child.find(f".//{{{_NS_A}}}cNvPr")
            if cNvPr is not None:
                child_names.append(cNvPr.get("name", ""))

        # Remove group from spTree
        sp_tree.remove(grpSp)

        # Add children back to spTree
        for child in children:
            sp_tree.append(child)

        return child_names
    finally:
        _save_prs(prs, prs_or_path if is_path else None, is_path)


def list_groups(prs_or_path, slide_index: int) -> list[GroupInfo]:
    """List all group shapes on a slide."""
    from lxml import etree

    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        results = []

        for shape in slide.shapes:
            elem = shape._element
            if elem.tag == f"{{{_NS_P}}}grpSp":
                info = GroupInfo(name=shape.name)

                # Get bounds
                sp_pr = elem.find(f"{{{_NS_P}}}grpSpPr")
                if sp_pr is not None:
                    xfrm = sp_pr.find(f"{{{_NS_A}}}xfrm")
                    if xfrm is not None:
                        off = xfrm.find(f"{{{_NS_A}}}off")
                        ext = xfrm.find(f"{{{_NS_A}}}ext")
                        if off is not None:
                            info.left = int(off.get("x", "0"))
                            info.top = int(off.get("y", "0"))
                        if ext is not None:
                            info.width = int(ext.get("cx", "0"))
                            info.height = int(ext.get("cy", "0"))

                # Count and name children
                child_names = []
                for child in elem:
                    tag = child.tag
                    if tag in (f"{{{_NS_P}}}nvGrpSpPr", f"{{{_NS_P}}}grpSpPr"):
                        continue
                    cNvPr = child.find(f".//{{{_NS_P}}}cNvPr")
                    if cNvPr is None:
                        cNvPr = child.find(f".//{{{_NS_A}}}cNvPr")
                    if cNvPr is not None:
                        child_names.append(cNvPr.get("name", ""))

                info.child_count = len(child_names)
                info.child_names = child_names
                results.append(info)

        return results
    finally:
        pass


def list_group_children(prs_or_path, slide_index: int,
                        group_name: str) -> list[dict]:
    """List all children of a group shape with their properties."""
    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        group = _find_shape(slide, group_name)
        if group is None:
            return []

        results = []
        grpSp = group._element
        for child in grpSp:
            tag = child.tag
            if tag in (f"{{{_NS_P}}}nvGrpSpPr", f"{{{_NS_P}}}grpSpPr"):
                continue

            info = {"tag": tag.split("}")[-1] if "}" in tag else tag}

            cNvPr = child.find(f".//{{{_NS_P}}}cNvPr")
            if cNvPr is None:
                cNvPr = child.find(f".//{{{_NS_A}}}cNvPr")
            if cNvPr is not None:
                info["name"] = cNvPr.get("name", "")
                info["id"] = cNvPr.get("id", "")

            # Get position
            xfrm = child.find(f".//{{{_NS_A}}}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{{{_NS_A}}}off")
                ext = xfrm.find(f"{{{_NS_A}}}ext")
                if off is not None:
                    info["left"] = int(off.get("x", "0"))
                    info["top"] = int(off.get("y", "0"))
                if ext is not None:
                    info["width"] = int(ext.get("cx", "0"))
                    info["height"] = int(ext.get("cy", "0"))

            # Get text if available
            for t_elem in child.iter(f"{{{_NS_A}}}t"):
                if t_elem.text:
                    info["text"] = t_elem.text[:100]
                    break

            results.append(info)

        return results
    finally:
        pass


def add_to_group(prs_or_path, slide_index: int, *,
                 group_name: str, shape_name: str) -> bool:
    """Add an existing shape to a group.

    Parameters
    ----------
    group_name : str
        Name of the target group.
    shape_name : str
        Name of the shape to add.
    """
    from lxml import etree

    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        group = _find_shape(slide, group_name)
        shape = _find_shape(slide, shape_name)
        if group is None or shape is None:
            return False

        sp_tree = slide.shapes._spTree

        # Remove shape from spTree
        sp_tree.remove(shape._element)

        # Add to group
        group._element.append(shape._element)
        return True
    finally:
        _save_prs(prs, prs_or_path if is_path else None, is_path)


def remove_from_group(prs_or_path, slide_index: int, *,
                      group_name: str, shape_name: str) -> bool:
    """Remove a shape from a group and place it back on the slide.

    Parameters
    ----------
    group_name : str
        Name of the group.
    shape_name : str
        Name of the child shape to remove from the group.
    """
    from lxml import etree

    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        group = _find_shape(slide, group_name)
        if group is None:
            return False

        sp_tree = slide.shapes._spTree

        # Find child shape in group
        for child in list(group._element):
            cNvPr = child.find(f".//{{{_NS_P}}}cNvPr")
            if cNvPr is None:
                cNvPr = child.find(f".//{{{_NS_A}}}cNvPr")
            if cNvPr is not None and cNvPr.get("name") == shape_name:
                group._element.remove(child)
                sp_tree.append(child)
                return True

        return False
    finally:
        _save_prs(prs, prs_or_path if is_path else None, is_path)


def move_in_group(prs_or_path, slide_index: int, *,
                  group_name: str, shape_name: str,
                  new_index: int) -> bool:
    """Change the z-order of a shape within a group.

    Parameters
    ----------
    new_index : int
        New position index within the group (0-based).
    """
    from lxml import etree

    prs, is_path = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        group = _find_shape(slide, group_name)
        if group is None:
            return False

        # Find the child shape
        child_elem = None
        for child in list(group._element):
            cNvPr = child.find(f".//{{{_NS_P}}}cNvPr")
            if cNvPr is None:
                cNvPr = child.find(f".//{{{_NS_A}}}cNvPr")
            if cNvPr is not None and cNvPr.get("name") == shape_name:
                child_elem = child
                break

        if child_elem is None:
            return False

        # Remove and re-insert at new position
        group._element.remove(child_elem)

        # Count non-property children
        children = [c for c in group._element
                    if c.tag not in (f"{{{_NS_P}}}nvGrpSpPr", f"{{{_NS_P}}}grpSpPr")]

        # Clamp index
        idx = max(0, min(new_index, len(children)))

        # Find insertion point (after nvGrpSpPr and grpSpPr)
        insert_pos = 0
        for i, child in enumerate(group._element):
            if child.tag in (f"{{{_NS_P}}}nvGrpSpPr", f"{{{_NS_P}}}grpSpPr"):
                insert_pos = i + 1

        # Adjust for desired index
        actual_pos = insert_pos + idx
        actual_pos = min(actual_pos, len(group._element))

        group._element.insert(actual_pos, child_elem)
        return True
    finally:
        _save_prs(prs, prs_or_path if is_path else None, is_path)
