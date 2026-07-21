"""VBA macro support for PPTX presentations.

VBA macros in PPTX files live inside ``ppt/vbaProject.bin``, an OLE compound
document embedded in the ZIP archive.  Full VBA *authoring* from Python is
impractical, but this module covers the practical operations: detection,
extraction, injection, macro-to-shape attachment, and cross-presentation
preservation.

The module works at the ZIP level (via ``zipfile``) for binary operations and
uses ``lxml`` for XML manipulation of ``[Content_Types].xml`` and
``.rels`` files.  All heavy dependencies are lazily imported.

Public API
----------
- :func:`has_vba_project`       -- check whether a PPTX contains a VBA project
- :func:`list_macro_names`      -- extract macro / module names from vbaProject.bin
- :func:`inject_vba_project`    -- inject a pre-built vbaProject.bin into a PPTX
- :func:`extract_vba_project`   -- extract vbaProject.bin to a file
- :func:`attach_macro_to_shape` -- set a shape click-action to call a VBA macro
- :func:`detach_macro_from_shape` -- remove macro attachment from a shape
- :func:`preserve_vba_project`  -- copy VBA project from one PPTX to another
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any

__all__ = [
    "attach_macro_to_shape",
    "detach_macro_from_shape",
    "extract_vba_project",
    "has_vba_project",
    "inject_vba_project",
    "list_macro_names",
    "preserve_vba_project",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_VBA_BIN_PATH = "ppt/vbaProject.bin"
_VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
_VBA_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"

# OOXML namespaces
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

_A_NS_PREFIX = f"{{{_A_NS}}}"
_R_NS_PREFIX = f"{{{_R_NS}}}"
_PR_NS_PREFIX = f"{{{_PR_NS}}}"
_P_NS_PREFIX = f"{{{_P_NS}}}"

# Regex fallback patterns for extracting macro names from the OLE binary.
# VBA stores module names as UTF-16LE strings preceded by a length byte.
# We look for common VBA module record markers.
_MODULE_NAME_RE = re.compile(
    rb"(?:\x00)([A-Za-z_][A-Za-z0-9_]{1,31})(?:\x00)"
)

# ---------------------------------------------------------------------------
# Helpers – Presentation / path resolution
# ---------------------------------------------------------------------------


def _is_presentation(obj: Any) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _resolve_path(prs_or_path: Any) -> str | None:
    """Return the file path if *prs_or_path* is a path, else ``None``."""
    if _is_presentation(prs_or_path):
        return None
    return str(prs_or_path)


def _ensure_path_on_disk(prs_or_path: Any) -> str:
    """Return a file path on disk, saving to a temp file if needed."""
    path = _resolve_path(prs_or_path)
    if path is not None:
        return path
    prs = prs_or_path  # already a Presentation object
    tmp_dir = tempfile.mkdtemp(prefix="pptx_skill_vba_")
    tmp_path = os.path.join(tmp_dir, "work.pptx")
    prs.save(tmp_path)
    return tmp_path


# ---------------------------------------------------------------------------
# Helpers – ZIP I/O
# ---------------------------------------------------------------------------


def _read_zip_to_memory(path: str) -> dict[str, bytes]:
    """Read all entries from a PPTX zip into a dict."""
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _write_zip_from_memory(
    path: str,
    zip_data: dict[str, bytes],
    *,
    overwrite: dict[str, bytes] | None = None,
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
# Helpers – XML manipulation (lxml, lazily imported)
# ---------------------------------------------------------------------------


def _parse_xml(data: bytes) -> Any:
    """Parse XML bytes into an lxml element tree."""
    from lxml import etree

    return etree.fromstring(data)


def _serialize_xml(elem: Any) -> bytes:
    """Serialize an lxml element to XML bytes with declaration."""
    from lxml import etree

    return etree.tostring(elem, xml_declaration=True, encoding="UTF-8", standalone=True)


def _ensure_content_type(content_types_elem: Any, part_name: str, content_type: str) -> None:
    """Add an Override element to ``[Content_Types].xml`` if not present.

    *part_name* should NOT start with a leading ``/`` — this function adds it.
    """
    for override in content_types_elem:
        existing = override.get("PartName", "")
        if existing.lstrip("/") == part_name.lstrip("/"):
            return
    from lxml import etree

    override = etree.SubElement(content_types_elem, "Override")
    override.set("PartName", f"/{part_name.lstrip('/')}")
    override.set("ContentType", content_type)


def _remove_content_type(content_types_elem: Any, part_name: str) -> bool:
    """Remove an Override element from ``[Content_Types].xml`` by part name.

    Returns ``True`` if an element was removed, ``False`` otherwise.
    """
    normalised = part_name.lstrip("/")
    for override in content_types_elem:
        existing = override.get("PartName", "").lstrip("/")
        if existing == normalised:
            content_types_elem.remove(override)
            return True
    return False


def _add_vba_relationship(rels_elem: Any) -> str:
    """Add a VBA project relationship to *rels_elem* if one does not exist.

    Returns the relationship ID (existing or new ``rId``.
    """
    from lxml import etree

    # Check if a VBA rel already exists
    for rel in rels_elem:
        if rel.get("Type", "") == _VBA_REL_TYPE:
            return rel.get("Id", "rId1")

    # Find the max rId to avoid collisions
    max_rid = 0
    for rel in rels_elem:
        rid = rel.get("Id", "rId0")
        try:
            max_rid = max(max_rid, int(rid.replace("rId", "")))
        except ValueError:
            pass

    new_id = f"rId{max_rid + 1}"
    new_rel = etree.SubElement(rels_elem, "Relationship")
    new_rel.set("Id", new_id)
    new_rel.set("Type", _VBA_REL_TYPE)
    new_rel.set("Target", "vbaProject.bin")
    return new_id


def _remove_vba_relationship(rels_elem: Any) -> bool:
    """Remove the VBA project relationship from *rels_elem*.

    Returns ``True`` if a relationship was removed.
    """
    for rel in list(rels_elem):
        if rel.get("Type", "") == _VBA_REL_TYPE:
            rels_elem.remove(rel)
            return True
    return False


# ---------------------------------------------------------------------------
# Helpers – OLE / VBA binary parsing
# ---------------------------------------------------------------------------


def _list_macro_names_olefile(vba_bytes: bytes) -> list[str] | None:
    """Extract macro names using the ``olefile`` library.

    Returns ``None`` if ``olefile`` is not available or parsing fails.
    """
    try:
        import olefile
    except ImportError:
        return None

    try:
        ole = olefile.OleFileIO(BytesIO(vba_bytes))
    except Exception:
        log.debug("olefile failed to parse vbaProject.bin", exc_info=True)
        return None

    names: list[str] = []
    try:
        # olefile exposes the VBA directory stream; module names appear as
        # stream entries under the "VBA" storage.
        vba_streams = [
            entry
            for entry in ole.listdir()
            if entry and entry[0].upper() == "VBA"
        ]
        for parts in vba_streams:
            # Stream names like ("VBA", "Module1"), ("VBA", "ThisWorkbook")
            if len(parts) >= 2:
                name = parts[-1]
                # Skip internal VBA streams
                if name.startswith("__"):
                    continue
                # The dir stream is metadata, not a user module
                if name.lower() in ("dir", "_dir"):
                    continue
                names.append(name)
    except Exception:
        log.debug("olefile stream enumeration failed", exc_info=True)
        return None
    finally:
        ole.close()

    return names if names else None


def _list_macro_names_regex(vba_bytes: bytes) -> list[str]:
    """Fallback: extract macro names via regex search on the OLE binary.

    This is a best-effort approach that looks for identifier-like strings
    surrounded by null bytes (UTF-16LE encoding typical in VBA storage).
    """
    # Try UTF-16LE decoding of the binary to find module names
    # VBA stores names as UTF-16LE in the dir stream
    seen: set[str] = set()
    names: list[str] = []

    # Strategy 1: look for UTF-16LE encoded identifiers
    # A VBA module name in UTF-16LE looks like: \x00A\x00u\x00t\x00o\x00O\x00p\x00e\x00n
    try:
        # Decode the binary, replacing non-decodable bytes
        text = vba_bytes.decode("utf-16-le", errors="ignore")
        # Find valid VBA identifiers (letters, digits, underscores, must start with letter)
        for match in re.finditer(r"\b([A-Za-z_][A-Za-z0-9_]{0,30})\b", text):
            name = match.group(1)
            # Filter out common non-macro identifiers
            if name.lower() in (
                "attribute", "vb_name", "vb_customizable", "option",
                "explicit", "private", "public", "sub", "function",
                "end", "dim", "as", "string", "long", "integer",
                "boolean", "variant", "object", "set", "nothing",
                "true", "false", "byval", "byref", "optional",
                "paramarray", "static", "const", "type", "enum",
                "class", "module", "property", "get", "let",
                "new", "me", "not", "and", "or", "if", "then",
                "else", "elseif", "select", "case", "for", "to",
                "step", "next", "do", "loop", "while", "wend",
                "until", "with", "on", "error", "goto", "gosub",
                "return", "call", "raiseevent", "implements",
                "friend", "global", "declare", "lib", "alias",
                "ptrsafe", "event", "handles",
            ):
                continue
            # Skip very short or very long names (unlikely to be module names)
            if len(name) < 2 or len(name) > 31:
                continue
            if name not in seen:
                seen.add(name)
                names.append(name)
    except Exception:
        log.debug("UTF-16LE regex extraction failed", exc_info=True)

    # Strategy 2: raw byte pattern search for module record markers
    if not names:
        for match in _MODULE_NAME_RE.finditer(vba_bytes):
            try:
                name = match.group(1).decode("ascii")
            except UnicodeDecodeError:
                continue
            if len(name) >= 2 and name not in seen:
                seen.add(name)
                names.append(name)

    return names


# ---------------------------------------------------------------------------
# Public API – Detection
# ---------------------------------------------------------------------------


def has_vba_project(prs_or_path: Any) -> bool:
    """Check whether a PPTX file contains a VBA project.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.

    Returns:
        ``True`` if ``ppt/vbaProject.bin`` exists in the archive.
    """
    path = _ensure_path_on_disk(prs_or_path)
    try:
        with zipfile.ZipFile(path, "r") as zf:
            return _VBA_BIN_PATH in zf.namelist()
    except (zipfile.BadZipFile, FileNotFoundError, OSError) as exc:
        log.warning("has_vba_project: cannot read %s: %s", path, exc)
        return False


def list_macro_names(prs_or_path: Any) -> list[str]:
    """Extract macro / module names from the VBA project in a PPTX.

    Uses ``olefile`` for structured OLE parsing when available, falling back
    to a regex-based binary search for module name markers.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.

    Returns:
        A list of macro / module name strings.  Returns an empty list if the
        PPTX has no VBA project or if the binary cannot be parsed.
    """
    path = _ensure_path_on_disk(prs_or_path)
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if _VBA_BIN_PATH not in zf.namelist():
                return []
            vba_bytes = zf.read(_VBA_BIN_PATH)
    except (zipfile.BadZipFile, FileNotFoundError, OSError) as exc:
        log.warning("list_macro_names: cannot read %s: %s", path, exc)
        return []

    if not vba_bytes:
        return []

    # Try olefile first
    names = _list_macro_names_olefile(vba_bytes)
    if names is not None:
        return names

    # Fallback to regex
    log.info("list_macro_names: olefile unavailable or failed, using regex fallback")
    return _list_macro_names_regex(vba_bytes)


# ---------------------------------------------------------------------------
# Public API – Injection
# ---------------------------------------------------------------------------


def inject_vba_project(prs_or_path: Any, vba_bin_path: str | Path) -> bool:
    """Inject a pre-built ``vbaProject.bin`` into a PPTX file.

    This adds the binary blob to ``ppt/vbaProject.bin`` inside the ZIP,
    registers the correct content type in ``[Content_Types].xml``, and adds
    a VBA relationship in ``ppt/presentation.xml.rels``.

    If the PPTX already contains a VBA project it is **replaced**.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        vba_bin_path: Path to the ``vbaProject.bin`` file to inject.

    Returns:
        ``True`` on success, ``False`` on failure.
    """
    vba_bin_path = Path(vba_bin_path).resolve()
    if not vba_bin_path.exists():
        log.error("inject_vba_project: vba_bin_path not found: %s", vba_bin_path)
        return False

    try:
        vba_bytes = vba_bin_path.read_bytes()
    except OSError as exc:
        log.error("inject_vba_project: cannot read vba bin: %s", exc)
        return False

    if not vba_bytes:
        log.error("inject_vba_project: vbaProject.bin is empty")
        return False

    path = _ensure_path_on_disk(prs_or_path)
    try:
        zip_data = _read_zip_to_memory(path)
    except (zipfile.BadZipFile, OSError) as exc:
        log.error("inject_vba_project: cannot read PPTX: %s", exc)
        return False

    # Create backup
    bak_path = path + ".bak.pptx"
    if os.path.exists(path):
        shutil.copy2(path, bak_path)

    try:
        from lxml import etree  # noqa: F811 – re-import in function scope
    except ImportError:
        log.error("inject_vba_project: lxml is required but not installed")
        return False

    updates: dict[str, bytes] = {}

    # 1. Add / replace the vbaProject.bin blob
    updates[_VBA_BIN_PATH] = vba_bytes

    # 2. Update [Content_Types].xml
    ct_bytes = zip_data.get("[Content_Types].xml")
    if ct_bytes is None:
        log.error("inject_vba_project: [Content_Types].xml not found in archive")
        return False
    ct_elem = _parse_xml(ct_bytes)
    _ensure_content_type(ct_elem, _VBA_BIN_PATH, _VBA_CONTENT_TYPE)
    updates["[Content_Types].xml"] = _serialize_xml(ct_elem)

    # 3. Update ppt/presentation.xml.rels
    rels_name = "ppt/_rels/presentation.xml.rels"
    if rels_name in zip_data:
        rels_elem = _parse_xml(zip_data[rels_name])
    else:
        rels_elem = etree.Element("Relationships", nsmap={None: _PR_NS})
    _add_vba_relationship(rels_elem)
    updates[rels_name] = _serialize_xml(rels_elem)

    # 4. Write back
    try:
        _write_zip_from_memory(path, zip_data, overwrite=updates)
    except OSError as exc:
        log.error("inject_vba_project: write failed: %s", exc)
        # Restore from backup
        if os.path.exists(bak_path):
            shutil.copy2(bak_path, path)
        return False

    return True


# ---------------------------------------------------------------------------
# Public API – Extraction
# ---------------------------------------------------------------------------


def extract_vba_project(prs_or_path: Any, output_path: str | Path) -> str:
    """Extract ``vbaProject.bin`` from a PPTX to a file.

    Args:
        prs_or_path: A ``Presentation`` object or path to a PPTX file.
        output_path: Destination file path for the extracted binary.

    Returns:
        The resolved output path as a string.

    Raises:
        ValueError: If the PPTX does not contain a VBA project.
        OSError: If the output path cannot be written.
    """
    output_path = Path(output_path).resolve()
    path = _ensure_path_on_disk(prs_or_path)

    try:
        with zipfile.ZipFile(path, "r") as zf:
            if _VBA_BIN_PATH not in zf.namelist():
                raise ValueError(
                    f"PPTX does not contain a VBA project: {_VBA_BIN_PATH} not found"
                )
            vba_bytes = zf.read(_VBA_BIN_PATH)
    except zipfile.BadZipFile as exc:
        raise ValueError(f"Cannot read PPTX as ZIP: {exc}") from exc

    # Ensure parent directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(vba_bytes)

    return str(output_path)


# ---------------------------------------------------------------------------
# Public API – Macro attachment / detachment on shapes
# ---------------------------------------------------------------------------


def attach_macro_to_shape(shape: Any, macro_name: str) -> None:
    """Set a shape's click action to call a VBA macro.

    This adds an ``<a:hlinkClick>`` element with
    ``action="ppaction://macro?name=MacroName"`` to the shape's XML tree.

    Args:
        shape: A python-pptx ``Shape`` object.
        macro_name: The VBA macro name (e.g. ``"Module1.MyMacro"``).

    Raises:
        ValueError: If *macro_name* is empty.
    """
    if not macro_name or not macro_name.strip():
        raise ValueError("macro_name must be a non-empty string")

    macro_name = macro_name.strip()

    # Access the shape's XML element via python-pptx
    sp = shape._element  # noqa: SLF001 – required to access the XML element

    # Find or create the <p:sp> or <p:pic> etc. -> <p:txBody> or direct
    # The click action goes on the shape's <p:cNvPr> (non-visual properties)
    # which is a child of <p:nvSpPr> or <p:nvPicPr> etc.
    from lxml import etree

    # Find cNvPr — it's always the second child of nvSpPr/nvPicPr/nvGrpSpPr
    cNvPr = sp.find(f"{_P_NS_PREFIX}nvSpPr/{_P_NS_PREFIX}cNvPr")
    if cNvPr is None:
        cNvPr = sp.find(f"{_P_NS_PREFIX}nvPicPr/{_P_NS_PREFIX}cNvPr")
    if cNvPr is None:
        cNvPr = sp.find(f"{_P_NS_PREFIX}nvGrpSpPr/{_P_NS_PREFIX}cNvPr")
    if cNvPr is None:
        # Try generic search for cNvPr in any child
        for child in sp:
            for sub in child:
                if sub.tag.endswith("cNvPr") or "cNvPr" in sub.tag:
                    cNvPr = sub
                    break
            if cNvPr is not None:
                break

    if cNvPr is None:
        log.warning("attach_macro_to_shape: cannot find cNvPr element on shape")
        return

    # Remove any existing hlinkClick with ppaction://macro
    for existing in cNvPr.findall(f"{_A_NS_PREFIX}hlinkClick"):
        action = existing.get("action", "")
        if "ppaction://macro" in action:
            cNvPr.remove(existing)

    # Add the new hlinkClick
    hlink = etree.SubElement(cNvPr, f"{_A_NS_PREFIX}hlinkClick")
    hlink.set("action", f"ppaction://macro?name={macro_name}")


def detach_macro_from_shape(shape: Any) -> None:
    """Remove macro attachment from a shape.

    Removes any ``<a:hlinkClick>`` element with a ``ppaction://macro`` action
    from the shape's non-visual properties.

    Args:
        shape: A python-pptx ``Shape`` object.
    """
    sp = shape._element  # noqa: SLF001 – required to access the XML element

    # Find cNvPr
    cNvPr = sp.find(f"{_P_NS_PREFIX}nvSpPr/{_P_NS_PREFIX}cNvPr")
    if cNvPr is None:
        cNvPr = sp.find(f"{_P_NS_PREFIX}nvPicPr/{_P_NS_PREFIX}cNvPr")
    if cNvPr is None:
        cNvPr = sp.find(f"{_P_NS_PREFIX}nvGrpSpPr/{_P_NS_PREFIX}cNvPr")
    if cNvPr is None:
        for child in sp:
            for sub in child:
                if sub.tag.endswith("cNvPr") or "cNvPr" in sub.tag:
                    cNvPr = sub
                    break
            if cNvPr is not None:
                break

    if cNvPr is None:
        return

    # Remove all hlinkClick elements with ppaction://macro
    removed = 0
    for existing in list(cNvPr.findall(f"{_A_NS_PREFIX}hlinkClick")):
        action = existing.get("action", "")
        if "ppaction://macro" in action:
            cNvPr.remove(existing)
            removed += 1

    if removed:
        log.debug("detach_macro_from_shape: removed %d macro attachment(s)", removed)


# ---------------------------------------------------------------------------
# Public API – Preservation
# ---------------------------------------------------------------------------


def preserve_vba_project(source_path: str | Path, target_path: str | Path) -> bool:
    """Copy the VBA project from *source_path* to *target_path*.

    This is used by merge operations to carry macros across presentations.
    If the source has no VBA project, the target is left unchanged and
    ``False`` is returned.

    Args:
        source_path: Path to the source PPTX containing a VBA project.
        target_path: Path to the target PPTX to receive the VBA project.

    Returns:
        ``True`` if a VBA project was copied, ``False`` otherwise.
    """
    source_path = Path(source_path).resolve()
    target_path = Path(target_path).resolve()

    if not source_path.exists():
        log.error("preserve_vba_project: source not found: %s", source_path)
        return False
    if not target_path.exists():
        log.error("preserve_vba_project: target not found: %s", target_path)
        return False

    # Read source VBA binary
    try:
        with zipfile.ZipFile(str(source_path), "r") as zf:
            if _VBA_BIN_PATH not in zf.namelist():
                log.info("preserve_vba_project: source has no VBA project")
                return False
            vba_bytes = zf.read(_VBA_BIN_PATH)
    except (zipfile.BadZipFile, OSError) as exc:
        log.error("preserve_vba_project: cannot read source: %s", exc)
        return False

    # Create backup of target
    bak_path = str(target_path) + ".bak.pptx"
    shutil.copy2(str(target_path), bak_path)

    # Read target into memory
    try:
        zip_data = _read_zip_to_memory(str(target_path))
    except (zipfile.BadZipFile, OSError) as exc:
        log.error("preserve_vba_project: cannot read target: %s", exc)
        return False

    try:
        from lxml import etree  # noqa: F811
    except ImportError:
        log.error("preserve_vba_project: lxml is required but not installed")
        return False

    updates: dict[str, bytes] = {}

    # 1. Add / replace vbaProject.bin
    updates[_VBA_BIN_PATH] = vba_bytes

    # 2. Update [Content_Types].xml
    ct_bytes = zip_data.get("[Content_Types].xml")
    if ct_bytes is None:
        log.error("preserve_vba_project: [Content_Types].xml not found in target")
        return False
    ct_elem = _parse_xml(ct_bytes)
    _ensure_content_type(ct_elem, _VBA_BIN_PATH, _VBA_CONTENT_TYPE)
    updates["[Content_Types].xml"] = _serialize_xml(ct_elem)

    # 3. Update ppt/presentation.xml.rels
    rels_name = "ppt/_rels/presentation.xml.rels"
    if rels_name in zip_data:
        rels_elem = _parse_xml(zip_data[rels_name])
    else:
        rels_elem = etree.Element("Relationships", nsmap={None: _PR_NS})
    _add_vba_relationship(rels_elem)
    updates[rels_name] = _serialize_xml(rels_elem)

    # 4. Write back
    try:
        _write_zip_from_memory(str(target_path), zip_data, overwrite=updates)
    except OSError as exc:
        log.error("preserve_vba_project: write failed: %s", exc)
        # Restore from backup
        shutil.copy2(bak_path, str(target_path))
        return False

    return True
