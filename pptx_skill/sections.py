"""Section group support for PPTX presentations.

PowerPoint sections (introduced in PPT 2010) organize slides into named groups.
This module provides CRUD operations on the ``<p:sectionLst>`` element stored
inside ``<p:presentation>``.

Because python-pptx has no native section API, all manipulation is performed
directly on the underlying lxml element tree.

Usage
-----
>>> from pptx_skill.sections import list_sections, add_section
>>> sections = list_sections("deck.pptx")
>>> add_section("deck.pptx", "Introduction", start_slide=0)
"""

from __future__ import annotations

import logging
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree

log = logging.getLogger(__name__)

__all__ = [
    "SectionInfo",
    "list_sections",
    "add_section",
    "remove_section",
    "rename_section",
    "move_section",
    "collapse_section",
]

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------

_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_P_NS_PREFIX = f"{{{_P_NS}}}"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_R_NS_PREFIX = f"{{{_R_NS}}}"

# Tag constants (Clark notation)
_TAG_SECTION_LST = f"{_P_NS_PREFIX}sectionLst"
_TAG_SECTION = f"{_P_NS_PREFIX}section"
_TAG_SLD_ID_LST = f"{_P_NS_PREFIX}sldIdLst"
_TAG_SLD_ID = f"{_P_NS_PREFIX}sldId"

# Relationship-id attribute in Clark notation
_R_ID_ATTR = f"{_R_NS_PREFIX}id"

# ---------------------------------------------------------------------------
# Data class
# ---------------------------------------------------------------------------


@dataclass
class SectionInfo:
    """A single presentation section.

    Attributes:
        name: Section display name.
        slide_indices: 0-based indices of slides belonging to this section.
        section_id: Unique identifier (GUID-style string) stored in the XML.
    """

    name: str
    slide_indices: list[int] = field(default_factory=list)
    section_id: str = ""


# ---------------------------------------------------------------------------
# Internal helpers — presentation open / save
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
    p = Path(path)
    bak = p.with_suffix(".bak.pptx")
    if p.exists():
        import shutil

        shutil.copy2(str(p), str(bak))
    prs.save(str(p))


def _resolve_path(prs_or_path: Any) -> str | None:
    """Return the file path if *prs_or_path* is a path, else ``None``."""
    if _is_presentation(prs_or_path):
        return None
    return str(prs_or_path)


# ---------------------------------------------------------------------------
# Internal helpers — XML access
# ---------------------------------------------------------------------------


def _get_section_lst(prs: Any) -> etree._Element:
    """Get or create the ``<p:sectionLst>`` element inside the presentation.

    The section list is placed after ``<p:sldIdLst>`` per OOXML convention.
    """
    pres = prs._element  # CT_Presentation (lxml element)
    section_lst = pres.find(_TAG_SECTION_LST)
    if section_lst is not None:
        return section_lst

    # Create and insert after <p:sldIdLst> (or at end if not found)
    section_lst = etree.SubElement(pres, _TAG_SECTION_LST)

    # Move to correct position: after sldIdLst
    sld_id_lst = pres.find(f"{_P_NS_PREFIX}sldIdLst")
    if sld_id_lst is not None:
        sld_id_lst.addnext(section_lst)

    return section_lst


def _sld_id_for_index(prs: Any, slide_index: int) -> etree._Element | None:
    """Return the ``<p:sldId>`` element for a 0-based *slide_index*.

    Returns ``None`` if the index is out of range.
    """
    sld_id_lst = prs._element.find(f"{_P_NS_PREFIX}sldIdLst")
    if sld_id_lst is None:
        return None
    sld_ids = list(sld_id_lst.findall(_TAG_SLD_ID))
    if 0 <= slide_index < len(sld_ids):
        return sld_ids[slide_index]
    return None


def _generate_section_id() -> str:
    """Generate a unique section ID (GUID-style with braces)."""
    return "{" + str(uuid.uuid4()).upper() + "}"


def _slide_index_for_sld_id(prs: Any, sld_id_elem: etree._Element) -> int | None:
    """Return the 0-based slide index for a given ``<p:sldId>`` element.

    Matches by comparing the ``id`` attribute value. Returns ``None`` if not
    found.
    """
    sld_id_lst = prs._element.find(f"{_P_NS_PREFIX}sldIdLst")
    if sld_id_lst is None:
        return None
    target_id = sld_id_elem.get("id")
    for idx, elem in enumerate(sld_id_lst.findall(_TAG_SLD_ID)):
        if elem.get("id") == target_id:
            return idx
    return None


def _section_elem_at(section_lst: etree._Element, section_index: int) -> etree._Element:
    """Return the ``<p:section>`` at *section_index*, raising ``IndexError``."""
    sections = list(section_lst.findall(_TAG_SECTION))
    if not 0 <= section_index < len(sections):
        if sections:
            msg = f"section_index {section_index} out of range (0..{len(sections) - 1})"
        else:
            msg = "no sections exist"
        raise IndexError(msg)
    return sections[section_index]


def _build_section_info(prs: Any, section_elem: etree._Element) -> SectionInfo:
    """Construct a :class:`SectionInfo` from a ``<p:section>`` element."""
    name = section_elem.get("name", "")
    section_id = section_elem.get("id", "")

    slide_indices: list[int] = []
    sld_id_lst = section_elem.find(_TAG_SLD_ID_LST)
    if sld_id_lst is not None:
        for sld_id in sld_id_lst.findall(_TAG_SLD_ID):
            idx = _slide_index_for_sld_id(prs, sld_id)
            if idx is not None:
                slide_indices.append(idx)

    return SectionInfo(
        name=name,
        slide_indices=slide_indices,
        section_id=section_id,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def list_sections(prs_or_path: Any) -> list[SectionInfo]:
    """List all sections with their slide ranges.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).

    Returns:
        A list of :class:`SectionInfo` objects in document order.
        Returns an empty list if the presentation has no sections.
    """
    prs = _open_prs(prs_or_path)
    section_lst = prs._element.find(_TAG_SECTION_LST)
    if section_lst is None:
        return []

    result: list[SectionInfo] = []
    for section_elem in section_lst.findall(_TAG_SECTION):
        result.append(_build_section_info(prs, section_elem))
    return result


def add_section(
    prs_or_path: Any,
    name: str,
    *,
    start_slide: int | None = None,
    before_section: int | None = None,
    after_section: int | None = None,
) -> SectionInfo:
    """Create a new section in the presentation.

    Exactly one positioning hint should be provided. If *start_slide* is given,
    the section starts at that slide (0-based). If *before_section* is given,
    the new section is inserted before the section at that index. If
    *after_section* is given, the new section is inserted after the section at
    that index. If no hint is provided, the section is appended at the end.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        name: The display name for the new section.
        start_slide: 0-based slide index where this section starts.
        before_section: Insert before this section index.
        after_section: Insert after this section index.

    Returns:
        The newly created :class:`SectionInfo`.

    Raises:
        ValueError: If conflicting positioning hints are provided.
        IndexError: If *before_section* or *after_section* is out of range.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    # Validate that at most one positioning hint is given
    hints = sum(x is not None for x in (start_slide, before_section, after_section))
    if hints > 1:
        raise ValueError(
            "At most one of start_slide, before_section, after_section may be specified"
        )

    section_lst = _get_section_lst(prs)
    section_id = _generate_section_id()

    # Build the <p:section> element
    section_elem = etree.Element(_TAG_SECTION)
    section_elem.set("name", name)
    section_elem.set("id", section_id)

    # Populate <p:sldIdLst> inside the section
    if start_slide is not None:
        sld_id_elem = _sld_id_for_index(prs, start_slide)
        if sld_id_elem is None:
            sld_id_lst_elem = prs._element.find(f"{_P_NS_PREFIX}sldIdLst")
            n_slides = len(list(sld_id_lst_elem)) if sld_id_lst_elem is not None else 0
            if n_slides:
                msg = f"start_slide {start_slide} out of range (0..{n_slides - 1})"
            else:
                msg = "start_slide out of range (no slides in presentation)"
            raise IndexError(msg)
        # Add this slide and all subsequent slides that are not already
        # claimed by another section — but per OOXML convention, a section
        # typically starts at the given slide and includes all slides until
        # the next section begins. For explicitness, we add only the
        # specified slide's sldId reference.
        section_sld_id_lst = etree.SubElement(section_elem, _TAG_SLD_ID_LST)
        sld_id_copy = etree.SubElement(section_sld_id_lst, _TAG_SLD_ID)
        sld_id_copy.set("id", sld_id_elem.get("id", ""))
        sld_id_copy.set(_R_ID_ATTR, sld_id_elem.get(_R_ID_ATTR, ""))

        # Also include slides after start_slide up to the next section boundary
        _populate_section_slides(prs, section_elem, start_slide)
    else:
        # Empty section (no slides yet)
        etree.SubElement(section_elem, _TAG_SLD_ID_LST)

    # Determine insert position
    if before_section is not None:
        target = _section_elem_at(section_lst, before_section)
        target.addprevious(section_elem)
    elif after_section is not None:
        target = _section_elem_at(section_lst, after_section)
        target.addnext(section_elem)
    else:
        section_lst.append(section_elem)

    # If start_slide was specified, trim the preceding section that may now
    # contain slides that belong to this new section.
    if start_slide is not None:
        _trim_preceding_section(prs, section_lst, section_elem, start_slide)

    # Save if opened from path
    _save_prs(prs, path)

    info = _build_section_info(prs, section_elem)
    log.info("Added section %r (id=%s) with %d slide(s)", name, section_id, len(info.slide_indices))
    return info


def _populate_section_slides(
    prs: Any,
    section_elem: etree._Element,
    start_index: int,
) -> None:
    """Populate a section's sldIdLst with slides from *start_index* onward.

    Slides are included until a slide that already belongs to a later section
    (in the section list after the new one) is encountered, or until the end
    of the presentation.
    """
    # Collect slide indices claimed by sections after the new one
    section_lst = prs._element.find(_TAG_SECTION_LST)
    if section_lst is None:
        return

    # Find all slides already claimed by other sections
    claimed_indices: set[int] = set()
    for sec in section_lst.findall(_TAG_SECTION):
        if sec is section_elem:
            continue
        sec_sld_lst = sec.find(_TAG_SLD_ID_LST)
        if sec_sld_lst is not None:
            for sld_id in sec_sld_lst.findall(_TAG_SLD_ID):
                idx = _slide_index_for_sld_id(prs, sld_id)
                if idx is not None:
                    claimed_indices.add(idx)

    # Ensure sldIdLst exists and is clean
    sld_id_lst = section_elem.find(_TAG_SLD_ID_LST)
    if sld_id_lst is None:
        sld_id_lst = etree.SubElement(section_elem, _TAG_SLD_ID_LST)
    else:
        # Remove existing children (we'll rebuild)
        for child in list(sld_id_lst):
            sld_id_lst.remove(child)

    # Add slides from start_index to end, skipping those already claimed
    # by a section that comes *after* this one in the section list
    next_section_start: int | None = None
    sections = list(section_lst.findall(_TAG_SECTION))
    try:
        my_pos = sections.index(section_elem)
    except ValueError:
        my_pos = len(sections) - 1  # appended at end

    # Find the first slide index of the section immediately after ours
    if my_pos + 1 < len(sections):
        next_sec = sections[my_pos + 1]
        next_sec_sld_lst = next_sec.find(_TAG_SLD_ID_LST)
        if next_sec_sld_lst is not None:
            next_sld_ids = next_sec_sld_lst.findall(_TAG_SLD_ID)
            if next_sld_ids:
                next_section_start = _slide_index_for_sld_id(prs, next_sld_ids[0])

    # Count slides directly from XML (prs.slides may be stale after removals)
    pres_sld_id_lst = prs._element.find(f"{_P_NS_PREFIX}sldIdLst")
    n_slides = len(list(pres_sld_id_lst)) if pres_sld_id_lst is not None else 0

    for idx in range(start_index, n_slides):
        if next_section_start is not None and idx >= next_section_start:
            break
        sld_id_elem = _sld_id_for_index(prs, idx)
        if sld_id_elem is not None:
            ref = etree.SubElement(sld_id_lst, _TAG_SLD_ID)
            ref.set("id", sld_id_elem.get("id", ""))
            ref.set(_R_ID_ATTR, sld_id_elem.get(_R_ID_ATTR, ""))


def _trim_preceding_section(
    prs: Any,
    section_lst: etree._Element,
    new_section_elem: etree._Element,
    new_start_index: int,
) -> None:
    """Remove slides from the section immediately before *new_section_elem*
    that now belong to the new section (i.e., slides at or after *new_start_index*).
    """
    sections = list(section_lst.findall(_TAG_SECTION))
    try:
        my_pos = sections.index(new_section_elem)
    except ValueError:
        return

    if my_pos == 0:
        return  # No preceding section to trim

    prev_elem = sections[my_pos - 1]
    prev_sld_lst = prev_elem.find(_TAG_SLD_ID_LST)
    if prev_sld_lst is None:
        return

    # Remove sldId entries whose slide index >= new_start_index
    to_remove: list[etree._Element] = []
    for sld_id in prev_sld_lst.findall(_TAG_SLD_ID):
        idx = _slide_index_for_sld_id(prs, sld_id)
        if idx is not None and idx >= new_start_index:
            to_remove.append(sld_id)

    for elem in to_remove:
        prev_sld_lst.remove(elem)


def remove_section(
    prs_or_path: Any,
    section_index: int,
    *,
    remove_slides: bool = False,
) -> int:
    """Remove a section from the presentation.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        section_index: 0-based index of the section to remove.
        remove_slides: If ``True``, also remove all slides that belong to the
            section from the presentation.

    Returns:
        The number of remaining sections.

    Raises:
        IndexError: If *section_index* is out of range.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    section_lst = prs._element.find(_TAG_SECTION_LST)
    if section_lst is None:
        raise IndexError("No sections exist in this presentation")

    section_elem = _section_elem_at(section_lst, section_index)

    # Collect slide indices before removing the element
    info = _build_section_info(prs, section_elem)

    # Remove slides if requested (iterate in reverse to keep indices stable)
    if remove_slides and info.slide_indices:
        _remove_slides_by_indices(prs, sorted(info.slide_indices, reverse=True))

    # Remove the section element
    section_lst.remove(section_elem)

    # Remove the sectionLst container if empty
    remaining = list(section_lst.findall(_TAG_SECTION))
    if not remaining:
        section_lst.getparent().remove(section_lst)

    # Save if opened from path
    _save_prs(prs, path)

    count = len(remaining)
    log.info(
        "Removed section %d (%r); %d section(s) remaining",
        section_index,
        info.name,
        count,
    )
    return count


def _remove_slides_by_indices(prs: Any, reverse_sorted_indices: list[int]) -> None:
    """Remove slides at the given indices (must be sorted in reverse order)."""
    from pptx.oxml.ns import qn

    sld_id_lst = prs.slides._sldIdLst
    for idx in reverse_sorted_indices:
        sld_id_elems = list(sld_id_lst)
        if 0 <= idx < len(sld_id_elems):
            sld_id = sld_id_elems[idx]
            rid = sld_id.get(qn("r:id"))
            if rid:
                prs.part.drop_rel(rid)
            sld_id_lst.remove(sld_id)


def rename_section(
    prs_or_path: Any,
    section_index: int,
    new_name: str,
) -> SectionInfo:
    """Rename a section.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        section_index: 0-based index of the section to rename.
        new_name: The new display name.

    Returns:
        The updated :class:`SectionInfo`.

    Raises:
        IndexError: If *section_index* is out of range.
        ValueError: If *new_name* is empty.
    """
    if not new_name:
        raise ValueError("new_name must not be empty")

    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    section_lst = prs._element.find(_TAG_SECTION_LST)
    if section_lst is None:
        raise IndexError("No sections exist in this presentation")

    section_elem = _section_elem_at(section_lst, section_index)
    old_name = section_elem.get("name", "")
    section_elem.set("name", new_name)

    _save_prs(prs, path)

    info = _build_section_info(prs, section_elem)
    log.info("Renamed section %d from %r to %r", section_index, old_name, new_name)
    return info


def move_section(
    prs_or_path: Any,
    section_index: int,
    new_position: int,
) -> None:
    """Move a section to a new position in the section list.

    This reorders the ``<p:section>`` elements within the section list.
    It does **not** move the actual slides — only the section grouping
    metadata is rearranged.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        section_index: Current 0-based index of the section to move.
        new_position: Target 0-based position.

    Raises:
        IndexError: If *section_index* or *new_position* is out of range.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    section_lst = prs._element.find(_TAG_SECTION_LST)
    if section_lst is None:
        raise IndexError("No sections exist in this presentation")

    sections = list(section_lst.findall(_TAG_SECTION))
    if not 0 <= section_index < len(sections):
        raise IndexError(
            f"section_index {section_index} out of range (0..{len(sections) - 1})"
        )
    if not 0 <= new_position < len(sections):
        raise IndexError(
            f"new_position {new_position} out of range (0..{len(sections) - 1})"
        )

    if section_index == new_position:
        return

    section_elem = sections[section_index]

    # Remove from current position
    section_lst.remove(section_elem)

    # Re-find sections after removal (indices shifted)
    sections = list(section_lst.findall(_TAG_SECTION))

    # Insert at new position
    if new_position >= len(sections):
        section_lst.append(section_elem)
    else:
        sections[new_position].addprevious(section_elem)

    _save_prs(prs, path)

    log.info("Moved section %d to position %d", section_index, new_position)


def collapse_section(
    prs_or_path: Any,
    section_index: int,
    collapsed: bool = True,
) -> None:
    """Set the collapsed state of a section for presentation mode.

    In PowerPoint, collapsed sections are minimized in the slide navigator.
    The collapsed state is stored as a ``sectPr`` attribute on the section.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        section_index: 0-based index of the section.
        collapsed: ``True`` to collapse, ``False`` to expand.

    Raises:
        IndexError: If *section_index* is out of range.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)

    section_lst = prs._element.find(_TAG_SECTION_LST)
    if section_lst is None:
        raise IndexError("No sections exist in this presentation")

    section_elem = _section_elem_at(section_lst, section_index)

    # The collapsed state is stored in an extLst / sectionPr extension.
    # PowerPoint stores this as a <p:extLst> containing a
    # <p:ext uri="{2A20BM13-B3ED-4496-BB5E-3B6BDAFD7F52}"> with
    # section properties. We use a simpler approach: set the
    # "collapsed" attribute directly on <p:section> which is the
    # convention used by LibreOffice and some OOXML variants.
    # PowerPoint itself uses an extension list, but for interop
    # we store it as a custom attribute.
    if collapsed:
        section_elem.set("collapsed", "1")
    else:
        # Remove the attribute if expanding
        if "collapsed" in section_elem.attrib:
            del section_elem.attrib["collapsed"]

    _save_prs(prs, path)

    state = "collapsed" if collapsed else "expanded"
    log.info("Section %d set to %s", section_index, state)
