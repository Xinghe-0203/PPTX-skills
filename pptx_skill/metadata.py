"""Document metadata, custom properties, and embedded font management.

Provides APIs for reading/writing core document properties (title, author,
subject, keywords, etc.), custom key-value properties, and detecting/managing
embedded fonts.

OOXML reference: ECMA-376 Part 4, §15.2 (Core Properties),
§21.2 (Custom Properties), §15.5 (Font Embedding).
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from pptx_skill._io import is_presentation as _is_presentation
from pptx_skill._io import open_prs as _open_prs
from pptx_skill._io import save_prs as _save_prs_impl

__all__ = [
    "DocumentMetadata",
    "CustomProperty",
    "FontInfo",
    "get_metadata",
    "set_metadata",
    "get_custom_property",
    "set_custom_property",
    "delete_custom_property",
    "list_custom_properties",
    "list_embedded_fonts",
    "embed_font",
    "remove_embedded_font",
]

# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

_NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_NS_DC = "http://purl.org/dc/elements/1.1/"
_NS_DCTERMS = "http://purl.org/dc/terms/"
_NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
_NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class DocumentMetadata:
    """Core document properties."""
    title: str = ""
    subject: str = ""
    creator: str = ""
    last_modified_by: str = ""
    description: str = ""
    keywords: str = ""
    category: str = ""
    created: datetime | None = None
    modified: datetime | None = None
    revision: str = ""
    content_status: str = ""


@dataclass
class CustomProperty:
    """A custom document property."""
    name: str
    value: str | int | float | bool
    pid: int = 0
    type_name: str = "string"  # "string", "int", "float", "bool", "date"


@dataclass
class FontInfo:
    """Info about an embedded font."""
    name: str = ""
    part_path: str = ""
    font_key: str = ""
    is_embedded: bool = False
    is_subset: bool = False


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _save_prs(prs, path):
    _save_prs_impl(prs, path, backup=False)


def _vt_type_for_value(value) -> str:
    """Determine the VT element name for a Python value."""
    if isinstance(value, bool):
        return "bool"
    elif isinstance(value, int):
        return "i4"
    elif isinstance(value, float):
        return "r8"
    else:
        return "lpwstr"


def _vt_value_to_python(elem) -> tuple[Any, str]:
    """Convert a VT element to a Python value and type name."""
    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
    text = elem.text or ""
    if tag == "i4" or tag == "i8":
        try:
            return int(text), "int"
        except ValueError:
            return text, "string"
    elif tag == "r4" or tag == "r8":
        try:
            return float(text), "float"
        except ValueError:
            return text, "string"
    elif tag == "bool":
        return text.strip() == "-1" or text.strip() == "true", "bool"
    elif tag == "filetime":
        try:
            return datetime.fromisoformat(text), "date"
        except Exception:
            return text, "string"
    else:
        return text, "string"


# ---------------------------------------------------------------------------
# Core metadata
# ---------------------------------------------------------------------------

def get_metadata(prs_or_path) -> DocumentMetadata:
    """Read core document properties from a presentation.

    Returns a ``DocumentMetadata`` with all available fields populated.
    """
    prs = _open_prs(prs_or_path)
    meta = DocumentMetadata()
    cp = prs.core_properties
    meta.title = cp.title or ""
    meta.subject = cp.subject or ""
    meta.creator = cp.author or ""
    meta.last_modified_by = cp.last_modified_by or ""
    meta.description = cp.comments or ""
    meta.keywords = cp.keywords or ""
    meta.category = cp.category or ""
    meta.created = cp.created
    meta.modified = cp.modified
    meta.revision = cp.revision or ""
    meta.content_status = cp.content_status or ""
    return meta


def set_metadata(prs_or_path, **kwargs) -> bool:
    """Set core document properties.

    Accepts keyword arguments matching ``DocumentMetadata`` fields:
    ``title``, ``subject``, ``creator``, ``last_modified_by``,
    ``description``, ``keywords``, ``category``, ``revision``,
    ``content_status``.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    cp = prs.core_properties
    field_map = {
        "title": "title",
        "subject": "subject",
        "creator": "author",
        "last_modified_by": "last_modified_by",
        "description": "comments",
        "keywords": "keywords",
        "category": "category",
        "revision": "revision",
        "content_status": "content_status",
    }
    for kwarg, prop_name in field_map.items():
        if kwarg in kwargs:
            setattr(cp, prop_name, kwargs[kwarg])
    if is_path and path is not None:
        _save_prs(prs, path)
    return True


# ---------------------------------------------------------------------------
# Custom properties
# ---------------------------------------------------------------------------

def list_custom_properties(prs_or_path) -> list[CustomProperty]:
    """List all custom document properties.

    Custom properties are stored in ``docProps/custom.xml`` inside the PPTX.
    """
    import zipfile

    from lxml import etree

    props: list[CustomProperty] = []

    # Need to read from the ZIP directly since python-pptx doesn't expose custom props
    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        # If it's a Presentation object, we need to save it temporarily
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "docProps/custom.xml" not in zf.namelist():
                return props

            xml_bytes = zf.read("docProps/custom.xml")
            root = etree.fromstring(xml_bytes)

            for prop_elem in root.findall(f"{{{_NS_CUSTOM}}}property"):
                name = prop_elem.get("name", "")
                pid = int(prop_elem.get("pid", "0"))

                # Find the value child (vt:lpwstr, vt:i4, etc.)
                for child in prop_elem:
                    child_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    if child_tag in ("lpwstr", "i4", "i8", "r4", "r8", "bool", "filetime"):
                        value, type_name = _vt_value_to_python(child)
                        props.append(CustomProperty(
                            name=name, value=value, pid=pid, type_name=type_name
                        ))
                        break
    except Exception:
        pass

    return props


def get_custom_property(prs_or_path, name: str) -> CustomProperty | None:
    """Get a custom property by name. Returns None if not found."""
    for prop in list_custom_properties(prs_or_path):
        if prop.name == name:
            return prop
    return None


def set_custom_property(prs_or_path, name: str, value: str | int | float | bool) -> bool:
    """Set a custom document property. Creates it if it doesn't exist.

    Parameters
    ----------
    name : str
        Property name.
    value : str | int | float | bool
        Property value. Type is preserved.
    """
    import shutil
    import tempfile
    import zipfile

    from lxml import etree

    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        # Read existing custom.xml or create new
        with zipfile.ZipFile(path, "r") as zf:
            custom_xml = None
            if "docProps/custom.xml" in zf.namelist():
                custom_xml = zf.read("docProps/custom.xml")

        if custom_xml is not None:
            root = etree.fromstring(custom_xml)
        else:
            root = etree.Element(f"{{{_NS_CUSTOM}}}Properties")
            root.set("xmlns", _NS_CUSTOM)

        # Find or create the property
        existing = None
        max_pid = 1
        for prop_elem in root.findall(f"{{{_NS_CUSTOM}}}property"):
            pid = int(prop_elem.get("pid", "0"))
            max_pid = max(max_pid, pid)
            if prop_elem.get("name") == name:
                existing = prop_elem

        if existing is not None:
            # Remove old value children
            for child in list(existing):
                existing.remove(child)
            prop_elem = existing
        else:
            max_pid += 1
            prop_elem = etree.SubElement(root, f"{{{_NS_CUSTOM}}}property")
            prop_elem.set("fmtid", _FMTID)
            prop_elem.set("pid", str(max_pid))
            prop_elem.set("name", name)

        # Add the value
        vt_type = _vt_type_for_value(value)
        vt_elem = etree.SubElement(prop_elem, f"{{{_NS_VT}}}{vt_type}")
        if vt_type == "bool":
            vt_elem.text = "-1" if value else "0"
        else:
            vt_elem.text = str(value)

        # Write back to ZIP
        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        # Rewrite the ZIP with updated custom.xml
        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, "output.pptx")
        try:
            with zipfile.ZipFile(path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == "docProps/custom.xml":
                            zout.writestr(item, new_xml)
                        else:
                            zout.writestr(item, zin.read(item.filename))

                    # Add custom.xml if it didn't exist
                    if custom_xml is None:
                        # Also need to add relationship and content type
                        zout.writestr("docProps/custom.xml", new_xml)

            # Replace original with updated
            shutil.move(temp_zip, path)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        # Also add the relationship and content type if needed
        _ensure_custom_xml_relationship(path)
        _ensure_custom_xml_content_type(path)

        return True
    except Exception:
        return False


def delete_custom_property(prs_or_path, name: str) -> bool:
    """Delete a custom document property by name."""
    import shutil
    import tempfile
    import zipfile

    from lxml import etree

    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "docProps/custom.xml" not in zf.namelist():
                return False
            custom_xml = zf.read("docProps/custom.xml")

        root = etree.fromstring(custom_xml)
        found = False
        for prop_elem in root.findall(f"{{{_NS_CUSTOM}}}property"):
            if prop_elem.get("name") == name:
                root.remove(prop_elem)
                found = True

        if not found:
            return False

        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, "output.pptx")
        try:
            with zipfile.ZipFile(path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == "docProps/custom.xml":
                            zout.writestr(item, new_xml)
                        else:
                            zout.writestr(item, zin.read(item.filename))
            shutil.move(temp_zip, path)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        return True
    except Exception:
        return False


def _ensure_custom_xml_relationship(path: str):
    """Ensure _rels/.rels has a relationship to docProps/custom.xml."""
    import shutil
    import tempfile
    import zipfile

    from lxml import etree

    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "_rels/.rels" not in zf.namelist():
                return
            rels_xml = zf.read("_rels/.rels")

        root = etree.fromstring(rels_xml)

        # Check if relationship already exists
        target = "docProps/custom.xml"
        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        for rel in root:
            if rel.get("Target", "") == target:
                return

        # Add relationship
        max_id = 0
        for rel in root:
            rid = rel.get("Id", "")
            if rid.startswith("rId"):
                try:
                    max_id = max(max_id, int(rid[3:]))
                except ValueError:
                    pass

        new_rel = etree.SubElement(root, "Relationship")
        new_rel.set("Id", f"rId{max_id + 1}")
        new_rel.set("Type", rel_type)
        new_rel.set("Target", target)

        new_rels = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, "output.pptx")
        try:
            with zipfile.ZipFile(path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == "_rels/.rels":
                            zout.writestr(item, new_rels)
                        else:
                            zout.writestr(item, zin.read(item.filename))
            shutil.move(temp_zip, path)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass


def _ensure_custom_xml_content_type(path: str):
    """Ensure [Content_Types].xml has an entry for docProps/custom.xml."""
    import shutil
    import tempfile
    import zipfile

    from lxml import etree

    try:
        with zipfile.ZipFile(path, "r") as zf:
            ct_xml = zf.read("[Content_Types].xml")

        root = etree.fromstring(ct_xml)
        ct_ns = root.tag.split("}")[0] + "}" if "}" in root.tag else ""

        # Check if override exists
        for override in root.findall(f"{ct_ns}Override"):
            if override.get("PartName", "") == "/docProps/custom.xml":
                return

        # Add override
        override = etree.SubElement(root, f"{ct_ns}Override")
        override.set("PartName", "/docProps/custom.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")

        new_ct = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, "output.pptx")
        try:
            with zipfile.ZipFile(path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == "[Content_Types].xml":
                            zout.writestr(item, new_ct)
                        else:
                            zout.writestr(item, zin.read(item.filename))
            shutil.move(temp_zip, path)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Embedded fonts
# ---------------------------------------------------------------------------

def list_embedded_fonts(prs_or_path) -> list[FontInfo]:
    """List all embedded fonts in a presentation.

    Embedded fonts are stored in ``ppt/fonts/`` directory inside the PPTX.
    """
    import zipfile

    fonts: list[FontInfo] = []
    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if name.startswith("ppt/fonts/") and name.endswith(".ttf"):
                    font_name = os.path.splitext(os.path.basename(name))[0]
                    fonts.append(FontInfo(
                        name=font_name,
                        part_path=name,
                        is_embedded=True,
                        is_subset=font_name.startswith("fnt"),
                    ))
    except Exception:
        pass

    return fonts


def embed_font(prs_or_path, font_path: str, *, subset: bool = True) -> bool:
    """Embed a font file into the presentation.

    Parameters
    ----------
    font_path : str
        Path to the .ttf or .otf font file.
    subset : bool
        If True, mark as subset embed. If False, mark as full embed.

    Note
    ----
    This adds the font file to ``ppt/fonts/`` and registers it in the
    font table XML. PowerPoint will use the embedded font when the
    presentation is opened on a system without the font installed.

    Full font embedding may have licensing restrictions. Check the font's
    EULA before embedding.
    """
    import shutil
    import tempfile
    import zipfile

    from lxml import etree

    if not os.path.exists(font_path):
        raise FileNotFoundError(f"Font file not found: {font_path}")

    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        font_name = os.path.splitext(os.path.basename(font_path))[0]
        font_part = f"ppt/fonts/{font_name}.ttf"

        with open(font_path, "rb") as f:
            font_data = f.read()

        # Add font to ZIP and update fontTable.xml
        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, "output.pptx")

        try:
            with zipfile.ZipFile(path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
                    font_table_xml = None
                    for item in zin.infolist():
                        if item.filename == "ppt/fontTable.xml":
                            font_table_xml = zin.read(item.filename)
                            # Will update and write later
                        elif item.filename == font_part:
                            # Skip existing font with same name
                            zout.writestr(item, font_data)
                        else:
                            zout.writestr(item, zin.read(item.filename))

                    # Add font file
                    zout.writestr(font_part, font_data)

                    # Update fontTable.xml
                    if font_table_xml is not None:
                        ft_root = etree.fromstring(font_table_xml)
                    else:
                        ft_root = etree.Element("{{ft_ns}}fontTable")

                    ft_root.tag.split("}")[0] + "}" if "}" in ft_root.tag else \
                        "http://schemas.openxmlformats.org/drawingml/2006/main"

                    # Add font entry
                    font_elem = etree.SubElement(ft_root, "{{ft_ns}}font")
                    font_elem.set("charset", "00")
                    font_elem.set("panose", "00000000000000000000")
                    font_elem.set("pitchFamily", "00")

                    # Regular
                    regular = etree.SubElement(font_elem, "{{ft_ns}}regular")
                    regular.set("typeface", font_name)

                    embed_flag = "subset" if subset else "full"
                    regular.set("embed", embed_flag)

                    new_ft = etree.tostring(ft_root, xml_declaration=True, encoding="UTF-8", standalone=True)

                    if "ppt/fontTable.xml" in [item.filename for item in zin.infolist()]:
                        # Overwrite
                        pass  # Will be written as new entry
                    zout.writestr("ppt/fontTable.xml", new_ft)

            shutil.move(temp_zip, path)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        return True
    except Exception:
        return False


def remove_embedded_font(prs_or_path, font_name: str) -> bool:
    """Remove an embedded font from the presentation.

    Parameters
    ----------
    font_name : str
        Name of the font to remove (without .ttf extension).
    """
    import shutil
    import tempfile
    import zipfile

    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        font_part = f"ppt/fonts/{font_name}.ttf"
        found = False

        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, "output.pptx")
        try:
            with zipfile.ZipFile(path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == font_part:
                            found = True
                            # Skip this file (remove it)
                            continue
                        else:
                            zout.writestr(item, zin.read(item.filename))

            if found:
                shutil.move(temp_zip, path)
            else:
                os.remove(temp_zip)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        return found
    except Exception:
        return False
