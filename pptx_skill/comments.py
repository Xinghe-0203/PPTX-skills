"""Speaker notes and comments for PPTX presentations.

This module provides a unified API for reading, writing, and exporting speaker
notes (leveraging python-pptx's native ``notes_slide`` support) as well as
full OOXML comment support including threaded replies and resolution state.

Speaker notes are stored per-slide via the ``notesSlide`` part and are natively
supported by python-pptx.  Comments are stored in ``ppt/comments/commentN.xml``
and ``ppt/comments/authors.xml`` parts and require low-level OOXML manipulation
with lxml.

Public API
----------
Speaker notes:
- :func:`get_speaker_notes`     -- retrieve notes text for a slide
- :func:`set_speaker_notes`     -- overwrite notes text for a slide
- :func:`append_speaker_notes`  -- append text to existing notes
- :func:`remove_speaker_notes`  -- delete the notes slide entirely
- :func:`export_speaker_notes`  -- export all notes to txt / md / docx

Comments:
- :func:`add_comment`           -- add a comment (optionally as a reply)
- :func:`list_comments`         -- list comments with replies and positions
- :func:`delete_comment`        -- remove a comment by ID
- :func:`reply_to_comment`      -- add a reply to an existing comment
- :func:`resolve_comment`       -- mark a comment as resolved / unresolved
"""
from __future__ import annotations

import logging
import os
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal

__all__ = [
    "add_comment",
    "append_speaker_notes",
    "delete_comment",
    "export_speaker_notes",
    "get_speaker_notes",
    "list_comments",
    "remove_speaker_notes",
    "reply_to_comment",
    "resolve_comment",
    "set_speaker_notes",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------

_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_P_NS_PREFIX = f"{{{_P_NS}}}"

# EMU conversion constant
_EMU_PER_INCH = 914400

# ---------------------------------------------------------------------------
# Helpers – Presentation open / save
# ---------------------------------------------------------------------------


def _is_presentation(obj: Any) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path: Any) -> Any:
    """Open a Presentation from *prs_or_path*."""
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
        shutil.copy2(str(p), str(bak))
    prs.save(str(p))


def _resolve_path(prs_or_path: Any) -> str | None:
    """Return the file path if *prs_or_path* is a path, else ``None``."""
    if _is_presentation(prs_or_path):
        return None
    return str(prs_or_path)


def _ensure_path_on_disk(prs_or_path: Any) -> tuple[str, str | None]:
    """Return a file path on disk, saving to a temp file if needed.

    Returns a tuple of ``(path, tmp_dir)`` where *tmp_dir* is the temporary
    directory path that the caller must clean up, or ``None`` if no temporary
    directory was created (i.e. *prs_or_path* was already a file path).
    """
    path = _resolve_path(prs_or_path)
    if path is not None:
        return path, None

    prs = _open_prs(prs_or_path)
    tmp_dir = tempfile.mkdtemp(prefix="pptx_skill_comments_")
    tmp_path = os.path.join(tmp_dir, "work.pptx")
    prs.save(tmp_path)
    return tmp_path, tmp_dir


def _validate_slide_index(prs: Any, slide_index: int) -> None:
    """Raise ``IndexError`` if *slide_index* is out of range (1-based)."""
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(
            f"slide_index {slide_index} out of range "
            f"(1..{len(prs.slides)})"
        )


# ---------------------------------------------------------------------------
# Speaker notes
# ---------------------------------------------------------------------------


def get_speaker_notes(prs_or_path: Any, slide_index: int) -> str | None:
    """Return the speaker notes text for a slide, or ``None`` if none exist.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.

    Returns:
        The notes text as a string, or ``None`` when the slide has no notes.
    """
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)
    slide = prs.slides[slide_index - 1]
    if not slide.has_notes_slide:
        return None
    text = slide.notes_slide.notes_text_frame.text
    return text if text else None


def set_speaker_notes(prs_or_path: Any, slide_index: int, text: str) -> None:
    """Set the speaker notes text for a slide, replacing any existing notes.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.
        text: The notes text to set.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)
    slide = prs.slides[slide_index - 1]
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = text
    _save_prs(prs, path)


def append_speaker_notes(prs_or_path: Any, slide_index: int, text: str) -> None:
    """Append *text* to the existing speaker notes on a slide.

    If the slide has no existing notes a new notes slide is created.  A newline
    is inserted between the existing content and *text*.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.
        text: The text to append.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)
    slide = prs.slides[slide_index - 1]
    notes_slide = slide.notes_slide
    existing = notes_slide.notes_text_frame.text or ""
    if existing:
        notes_slide.notes_text_frame.text = existing + "\n" + text
    else:
        notes_slide.notes_text_frame.text = text
    _save_prs(prs, path)


def remove_speaker_notes(prs_or_path: Any, slide_index: int) -> bool:
    """Remove the speaker notes from a slide.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.

    Returns:
        ``True`` if notes were removed, ``False`` if none existed.
    """
    path = _resolve_path(prs_or_path)
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)
    slide = prs.slides[slide_index - 1]
    if not slide.has_notes_slide:
        return False
    # python-pptx doesn't expose a "remove notes slide" API directly.
    # We clear the text content instead; the notes part is preserved but empty.
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = ""
    _save_prs(prs, path)
    return True


def export_speaker_notes(
    prs_or_path: Any,
    output_path: str | Path,
    *,
    format: Literal["txt", "md", "docx"] = "txt",
) -> str:
    """Export all speaker notes to a file.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        output_path: Destination file path.
        format: Output format — ``"txt"`` (plain text), ``"md"`` (Markdown
            with slide headers), or ``"docx"`` (Word document).

    Returns:
        The absolute path of the written file.
    """
    prs = _open_prs(prs_or_path)
    output_path = str(output_path)

    # Collect all notes
    notes_list: list[tuple[int, str | None]] = []
    for i, slide in enumerate(prs.slides, start=1):
        text = None
        if slide.has_notes_slide:
            text = slide.notes_slide.notes_text_frame.text or None
        notes_list.append((i, text))

    if format == "txt":
        _export_notes_txt(notes_list, output_path)
    elif format == "md":
        _export_notes_md(notes_list, output_path)
    elif format == "docx":
        _export_notes_docx(notes_list, output_path)
    else:
        raise ValueError(f"Unsupported format: {format!r} (expected 'txt', 'md', or 'docx')")

    return os.path.abspath(output_path)


def _export_notes_txt(
    notes_list: list[tuple[int, str | None]], output_path: str
) -> None:
    """Write notes as plain text, one slide per block."""
    lines: list[str] = []
    for idx, text in notes_list:
        lines.append(f"--- Slide {idx} ---")
        lines.append(text if text else "(no notes)")
        lines.append("")
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def _export_notes_md(
    notes_list: list[tuple[int, str | None]], output_path: str
) -> None:
    """Write notes as Markdown with slide headings."""
    lines: list[str] = []
    for idx, text in notes_list:
        lines.append(f"## Slide {idx}")
        lines.append("")
        lines.append(text if text else "*No notes.*")
        lines.append("")
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def _export_notes_docx(
    notes_list: list[tuple[int, str | None]], output_path: str
) -> None:
    """Write notes as a Word document (requires python-docx)."""
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError as exc:
        raise ImportError(
            "python-docx is required for docx export. "
            "Install with: pip install python-docx"
        ) from exc

    doc = Document()
    for idx, text in notes_list:
        heading = doc.add_heading(f"Slide {idx}", level=2)
        for run in heading.runs:
            run.font.size = Pt(14)
        body = text if text else "(no notes)"
        para = doc.add_paragraph(body)
        for run in para.runs:
            run.font.size = Pt(11)
    doc.save(output_path)


# ---------------------------------------------------------------------------
# Comments – low-level OOXML helpers
# ---------------------------------------------------------------------------


def _read_zip_xml(zf: zipfile.ZipFile, arc_name: str) -> Any:
    """Read an XML entry from *zf* and return an lxml ``Element``."""
    from lxml import etree

    data = zf.read(arc_name)
    return etree.fromstring(data)


def _serialize_xml(elem: Any) -> bytes:
    """Serialize an lxml element to XML bytes with declaration."""
    from lxml import etree

    return etree.tostring(elem, xml_declaration=True, encoding="UTF-8", standalone=True)


def _next_comment_index(names: list[str]) -> int:
    """Return the next available comment file index (1-based).

    Looks for existing ``ppt/comments/commentN.xml`` entries in *names*.
    """
    import re

    pattern = re.compile(r"ppt/comments/comment(\d+)\.xml$")
    existing: set[int] = set()
    for name in names:
        m = pattern.match(name)
        if m:
            existing.add(int(m.group(1)))
    return max(existing, default=0) + 1


def _find_comment_index_for_slide(
    zf: zipfile.ZipFile, slide_index: int
) -> int | None:
    """Find the comment file index associated with *slide_index* (1-based).

    Walks the slide relationship file to find a ``comments`` relationship,
    then returns the numeric index from the target part name.
    """
    rels_name = f"ppt/slides/_rels/slide{slide_index}.xml.rels"
    if rels_name not in zf.namelist():
        return None

    rels_elem = _read_zip_xml(zf, rels_name)
    for rel in rels_elem:
        target = rel.get("Target", "")
        rel_type = rel.get("Type", "")
        if "/comments" in rel_type and "comment" in target.lower():
            # Target is like "../comments/comment1.xml"
            basename = target.rsplit("/", 1)[-1]
            try:
                return int(basename.replace("comment", "").replace(".xml", ""))
            except ValueError:
                continue
    return None


def _find_or_create_author(authors_elem: Any, author_name: str) -> int:
    """Find or create an author entry, returning the ``authorId``.

    Authors are stored in ``ppt/comments/authors.xml`` as::

        <p:cmAuthor id="0" name="User" .../>

    Returns the zero-based author ID.
    """
    from lxml import etree

    existing = authors_elem.findall(f"{_P_NS_PREFIX}cmAuthor")
    for author_elem in existing:
        if author_elem.get("name") == author_name:
            return int(author_elem.get("id"))

    # Create a new author entry
    new_id = max((int(a.get("id", "-1")) for a in existing), default=-1) + 1
    author_elem = etree.SubElement(authors_elem, f"{_P_NS_PREFIX}cmAuthor")
    author_elem.set("id", str(new_id))
    author_elem.set("name", author_name)
    author_elem.set("initials", author_name[:2] if author_name else "U")
    author_elem.set("color", "0")
    return new_id


def _ensure_authors_element(names: list[str], zip_data: dict[str, bytes]) -> tuple[Any, bool]:
    """Return ``(authors_element, was_created)`` for ``ppt/comments/authors.xml``.

    If the file does not exist yet it is created with the proper root element.
    """
    from lxml import etree

    authors_path = "ppt/comments/authors.xml"
    if authors_path in zip_data:
        elem = etree.fromstring(zip_data[authors_path])
        return elem, False

    # Create new authors element
    elem = etree.Element(f"{_P_NS_PREFIX}cmAuthorLst")
    return elem, True


def _ensure_content_type(
    content_types_elem: Any, part_name: str, content_type: str
) -> None:
    """Add an Override element to ``[Content_Types].xml`` if not present.

    *part_name* should NOT start with a leading ``/`` — this function adds it.
    """
    for override in content_types_elem:
        existing = override.get("PartName", "")
        # Normalise: both with and without leading /
        if existing.lstrip("/") == part_name.lstrip("/"):
            return
    from lxml import etree

    override = etree.SubElement(content_types_elem, "Override")
    override.set("PartName", f"/{part_name.lstrip('/')}")
    override.set("ContentType", content_type)


def _add_comment_relationship(
    rels_elem: Any, comment_index: int
) -> None:
    """Add a comments relationship to *rels_elem* if one does not exist."""
    from lxml import etree

    target = f"../comments/comment{comment_index}.xml"
    rel_type = f"{_R_NS}/comments"

    # Check if a comments rel already exists
    for rel in rels_elem:
        if rel.get("Type", "") == rel_type:
            return

    # Add the relationship
    max_rid = 0
    for rel in rels_elem:
        rid = rel.get("Id", "rId0")
        try:
            max_rid = max(max_rid, int(rid.replace("rId", "")))
        except ValueError:
            pass

    new_rel = etree.SubElement(rels_elem, "Relationship")
    new_rel.set("Id", f"rId{max_rid + 1}")
    new_rel.set("Type", rel_type)
    new_rel.set("Target", target)


def _parse_comment_tree(
    cm_elem: Any,
) -> dict:
    """Parse a ``<p:cm>`` element into a dict, recursively including replies."""
    replies: list[dict] = []
    # Replies are nested <p:cmLst> inside <p:cm>
    nested = cm_elem.find(f"{_P_NS_PREFIX}cmLst")
    if nested is not None:
        for reply_cm in nested.findall(f"{_P_NS_PREFIX}cm"):
            replies.append(_parse_comment_tree(reply_cm))

    pos_elem = cm_elem.find(f"{_P_NS_PREFIX}pos")
    position = None
    if pos_elem is not None:
        try:
            position = {
                "left": int(pos_elem.get("x", "0")),
                "top": int(pos_elem.get("y", "0")),
            }
        except (ValueError, TypeError):
            position = {"left": 0, "top": 0}

    text_elem = cm_elem.find(f"{_P_NS_PREFIX}text")
    text = text_elem.text if text_elem is not None else ""

    # Check for resolution status (extension attribute)
    resolved = cm_elem.get("resolved", "0") == "1"

    return {
        "id": cm_elem.get("id", ""),
        "author": cm_elem.get("authorId", ""),  # resolved to name by caller
        "text": text or "",
        "date": cm_elem.get("dt", ""),
        "replies": replies,
        "position": position,
        "resolved": resolved,
    }


def _resolve_author_names(
    comments: list[dict], authors_elem: Any
) -> list[dict]:
    """Replace ``authorId`` values with author names from the authors part."""
    author_map: dict[str, str] = {}
    for author_elem in authors_elem.findall(f"{_P_NS_PREFIX}cmAuthor"):
        aid = author_elem.get("id", "")
        name = author_elem.get("name", "")
        author_map[aid] = name

    def _resolve(comment: dict) -> dict:
        aid = str(comment.get("author", ""))
        comment["author"] = author_map.get(aid, aid)
        comment["replies"] = [_resolve(r) for r in comment.get("replies", [])]
        return comment

    return [_resolve(c) for c in comments]


def _find_comment_by_id(
    cm_lst: Any, comment_id: str
) -> tuple[Any, Any | None]:
    """Find a ``<p:cm>`` element by its ``id`` attribute.

    Returns ``(parent_cm_lst, cm_element)``.  The parent is the ``<p:cmLst>``
    that directly contains the found element, or ``None`` if not found.
    """
    for cm in cm_lst.findall(f"{_P_NS_PREFIX}cm"):
        if cm.get("id") == comment_id:
            return cm_lst, cm
        # Search replies
        nested = cm.find(f"{_P_NS_PREFIX}cmLst")
        if nested is not None:
            parent, found = _find_comment_by_id(nested, comment_id)
            if found is not None:
                return parent, found
    return cm_lst, None


def _max_comment_id(cm_lst: Any) -> int:
    """Return the maximum comment ID found in a ``<p:cmLst>``, recursively."""
    max_id = 0
    for cm in cm_lst.findall(f"{_P_NS_PREFIX}cm"):
        try:
            max_id = max(max_id, int(cm.get("id", "0")))
        except (ValueError, TypeError):
            pass
        # Check nested replies
        nested = cm.find(f"{_P_NS_PREFIX}cmLst")
        if nested is not None:
            max_id = max(max_id, _max_comment_id(nested))
    return max_id


def _build_comment_elem(
    comment_id: str,
    author_id: int,
    dt: str,
    pos_x: int | None,
    pos_y: int | None,
    text: str,
) -> Any:
    """Build a ``<p:cm>`` element."""
    from lxml import etree

    cm = etree.Element(f"{_P_NS_PREFIX}cm")
    cm.set("id", comment_id)
    cm.set("authorId", str(author_id))
    cm.set("dt", dt)

    if pos_x is not None and pos_y is not None:
        pos = etree.SubElement(cm, f"{_P_NS_PREFIX}pos")
        pos.set("x", str(pos_x))
        pos.set("y", str(pos_y))

    text_elem = etree.SubElement(cm, f"{_P_NS_PREFIX}text")
    text_elem.text = text

    return cm


# ---------------------------------------------------------------------------
# Comments – zip read/write helpers
# ---------------------------------------------------------------------------


def _read_zip_to_memory(path: str) -> dict[str, bytes]:
    """Read all entries from a PPTX zip into a dict."""
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _write_zip_from_memory(
    path: str, zip_data: dict[str, bytes], *, overwrite: dict[str, bytes] | None = None
) -> None:
    """Write a PPTX zip from in-memory entries.

    *overwrite* is a dict of ``{arc_name: bytes}`` that replaces or adds
    entries on top of *zip_data*.  This avoids duplicate entries.
    """
    merged = dict(zip_data)
    if overwrite:
        merged.update(overwrite)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(merged):
            zf.writestr(name, merged[name])


# ---------------------------------------------------------------------------
# Comments – public API
# ---------------------------------------------------------------------------


def add_comment(
    prs_or_path: Any,
    slide_index: int,
    text: str,
    *,
    author: str = "User",
    left: float = 0,
    top: float = 0,
    width: float = 2,
    height: float = 1,
    reply_to: str | None = None,
) -> str:
    """Add a comment to a slide.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.
        text: Comment text.
        author: Author display name.
        left: Horizontal position in inches from the left edge of the slide.
        top: Vertical position in inches from the top edge of the slide.
        width: Comment anchor width in inches (affects visual highlight area).
        height: Comment anchor height in inches (affects visual highlight area).
        reply_to: If given, the ID of a parent comment to reply to instead of
            creating a new top-level comment.

    Returns:
        The comment ID as a string.
    """
    from lxml import etree

    # Validate slide index first
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)

    path, tmp_dir = _ensure_path_on_disk(prs_or_path)
    pos_x = int(left * _EMU_PER_INCH)
    pos_y = int(top * _EMU_PER_INCH)
    now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    # Read all zip entries into memory
    zip_data = _read_zip_to_memory(path)
    updates: dict[str, bytes] = {}

    # --- Update authors ---
    authors_elem, _authors_created = _ensure_authors_element(list(zip_data), zip_data)
    author_id = _find_or_create_author(authors_elem, author)
    authors_path = "ppt/comments/authors.xml"
    updates[authors_path] = _serialize_xml(authors_elem)

    # --- Determine comment file ---
    # We need to peek at the current state to find existing relationships
    with zipfile.ZipFile(path, "r") as zf:
        comment_index = _find_comment_index_for_slide(zf, slide_index)

    if comment_index is None:
        # No comment part yet for this slide — create one
        comment_index = _next_comment_index(list(zip_data))
        comment_path = f"ppt/comments/comment{comment_index}.xml"

        # Create the root <p:cmLst> element
        cm_lst = etree.Element(f"{_P_NS_PREFIX}cmLst")

        # Add the comment
        new_id = "1"
        cm = _build_comment_elem(new_id, author_id, now_utc, pos_x, pos_y, text)
        cm_lst.append(cm)

        updates[comment_path] = _serialize_xml(cm_lst)

        # Update slide relationships
        rels_name = f"ppt/slides/_rels/slide{slide_index}.xml.rels"
        if rels_name in zip_data:
            rels_elem = etree.fromstring(zip_data[rels_name])
        else:
            rels_elem = etree.Element("Relationships", nsmap={None: _PR_NS})
        _add_comment_relationship(rels_elem, comment_index)
        updates[rels_name] = _serialize_xml(rels_elem)

        # Update [Content_Types].xml
        ct_elem = etree.fromstring(zip_data["[Content_Types].xml"])
        _ensure_content_type(ct_elem, comment_path,
                             "application/vnd.openxmlformats-officedocument.presentationml.comments+xml")
        _ensure_content_type(ct_elem, authors_path,
                             "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml")
        updates["[Content_Types].xml"] = _serialize_xml(ct_elem)

        # Write back
        _write_zip_from_memory(path, zip_data, overwrite=updates)
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return new_id

    # --- Existing comment file for this slide ---
    comment_path = f"ppt/comments/comment{comment_index}.xml"
    cm_lst = etree.fromstring(zip_data[comment_path])

    # Find the max comment ID (search recursively)
    max_id = _max_comment_id(cm_lst)

    if reply_to is not None:
        # Find the parent comment and add as a reply
        _, parent_cm = _find_comment_by_id(cm_lst, reply_to)
        if parent_cm is None:
            raise ValueError(f"Comment ID {reply_to!r} not found on slide {slide_index}")

        # Ensure parent has a nested <p:cmLst> for replies
        nested = parent_cm.find(f"{_P_NS_PREFIX}cmLst")
        if nested is None:
            nested = etree.SubElement(parent_cm, f"{_P_NS_PREFIX}cmLst")

        reply_id = str(max_id + 1)
        reply_cm = _build_comment_elem(reply_id, author_id, now_utc, None, None, text)
        nested.append(reply_cm)

        updates[comment_path] = _serialize_xml(cm_lst)

        # Update [Content_Types].xml for authors if not already there
        ct_elem = etree.fromstring(zip_data["[Content_Types].xml"])
        _ensure_content_type(ct_elem, authors_path,
                             "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml")
        updates["[Content_Types].xml"] = _serialize_xml(ct_elem)

        _write_zip_from_memory(path, zip_data, overwrite=updates)
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return reply_id

    # Add a new top-level comment
    new_id = str(max_id + 1)
    cm = _build_comment_elem(new_id, author_id, now_utc, pos_x, pos_y, text)
    cm_lst.append(cm)

    updates[comment_path] = _serialize_xml(cm_lst)

    # Update [Content_Types].xml for authors if not already there
    ct_elem = etree.fromstring(zip_data["[Content_Types].xml"])
    _ensure_content_type(ct_elem, authors_path,
                         "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml")
    updates["[Content_Types].xml"] = _serialize_xml(ct_elem)

    _write_zip_from_memory(path, zip_data, overwrite=updates)
    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return new_id


def list_comments(
    prs_or_path: Any,
    slide_index: int,
) -> list[dict]:
    """List all comments on a slide.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.

    Returns:
        A list of comment dicts.  Each dict has keys:
        ``"id"``, ``"author"``, ``"text"``, ``"date"``, ``"replies"``,
        ``"position"``, ``"resolved"``.  Replies are nested dicts with the
        same structure (without ``"position"``).
    """
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)

    path, tmp_dir = _ensure_path_on_disk(prs_or_path)

    with zipfile.ZipFile(path, "r") as zf:
        comment_index = _find_comment_index_for_slide(zf, slide_index)
        if comment_index is None:
            if tmp_dir is not None:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            return []

        comment_path = f"ppt/comments/comment{comment_index}.xml"
        if comment_path not in zf.namelist():
            if tmp_dir is not None:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            return []

        cm_lst = _read_zip_xml(zf, comment_path)

        # Parse comments
        comments: list[dict] = []
        for cm in cm_lst.findall(f"{_P_NS_PREFIX}cm"):
            comments.append(_parse_comment_tree(cm))

        # Resolve author IDs to names
        authors_path = "ppt/comments/authors.xml"
        if authors_path in zf.namelist():
            authors_elem = _read_zip_xml(zf, authors_path)
            comments = _resolve_author_names(comments, authors_elem)

    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return comments


def delete_comment(
    prs_or_path: Any,
    slide_index: int,
    comment_id: str,
) -> bool:
    """Delete a comment from a slide by its ID.

    If the comment has replies, the entire thread is removed.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.
        comment_id: The ID of the comment to delete.

    Returns:
        ``True`` if the comment was found and deleted, ``False`` otherwise.
    """
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)

    path, tmp_dir = _ensure_path_on_disk(prs_or_path)

    # Read zip into memory
    zip_data = _read_zip_to_memory(path)

    # Find comment index
    with zipfile.ZipFile(path, "r") as zf:
        comment_index = _find_comment_index_for_slide(zf, slide_index)
    if comment_index is None:
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return False

    comment_path = f"ppt/comments/comment{comment_index}.xml"
    if comment_path not in zip_data:
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return False

    from lxml import etree

    cm_lst = etree.fromstring(zip_data[comment_path])
    parent, cm_elem = _find_comment_by_id(cm_lst, comment_id)
    if cm_elem is None:
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return False

    parent.remove(cm_elem)

    # Write back only the modified comment file
    updates: dict[str, bytes] = {comment_path: _serialize_xml(cm_lst)}
    _write_zip_from_memory(path, zip_data, overwrite=updates)
    if tmp_dir is not None:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return True


def reply_to_comment(
    prs_or_path: Any,
    slide_index: int,
    comment_id: str,
    text: str,
    *,
    author: str = "User",
) -> str:
    """Add a reply to an existing comment.

    This is a convenience wrapper around :func:`add_comment` with
    ``reply_to=comment_id``.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.
        comment_id: The ID of the parent comment.
        text: Reply text.
        author: Author display name.

    Returns:
        The reply comment ID as a string.
    """
    return add_comment(
        prs_or_path,
        slide_index,
        text,
        author=author,
        reply_to=comment_id,
    )


def resolve_comment(
    prs_or_path: Any,
    slide_index: int,
    comment_id: str,
    resolved: bool = True,
) -> None:
    """Mark a comment as resolved or unresolved.

    Resolution state is stored as a ``resolved`` attribute on the ``<p:cm>``
    element (``"1"`` for resolved, ``"0"`` for unresolved).

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        slide_index: 1-based slide index.
        comment_id: The ID of the comment to resolve/unresolve.
        resolved: ``True`` to mark as resolved, ``False`` to unresolve.
    """
    prs = _open_prs(prs_or_path)
    _validate_slide_index(prs, slide_index)

    path, tmp_dir = _ensure_path_on_disk(prs_or_path)
    try:
        # Read zip into memory
        zip_data = _read_zip_to_memory(path)

        # Find comment index
        with zipfile.ZipFile(path, "r") as zf:
            comment_index = _find_comment_index_for_slide(zf, slide_index)
        if comment_index is None:
            raise ValueError(f"No comments found on slide {slide_index}")

        comment_path = f"ppt/comments/comment{comment_index}.xml"
        if comment_path not in zip_data:
            raise ValueError(f"No comments found on slide {slide_index}")

        from lxml import etree

        cm_lst = etree.fromstring(zip_data[comment_path])
        _, cm_elem = _find_comment_by_id(cm_lst, comment_id)
        if cm_elem is None:
            raise ValueError(f"Comment ID {comment_id!r} not found on slide {slide_index}")

        cm_elem.set("resolved", "1" if resolved else "0")

        # Write back only the modified comment file
        updates: dict[str, bytes] = {comment_path: _serialize_xml(cm_lst)}
        _write_zip_from_memory(path, zip_data, overwrite=updates)
    finally:
        if tmp_dir is not None:
            shutil.rmtree(tmp_dir, ignore_errors=True)
