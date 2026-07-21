"""Image compression and optimization for PPTX presentations.

Provides utilities to compress, convert, and analyze images stored inside a
PPTX archive.  PPTX files are ZIP archives; images live under ``ppt/media/``
and are referenced by slide relationship files (``ppt/slides/_rels/*.rels``)
as well as ``[Content_Types].xml``.

Public API
----------
- ``compress_images`` -- lossy resize / re-encode to reduce file size
- ``remove_unused_media`` -- delete orphaned media entries
- ``get_image_stats`` -- per-image statistics (dimensions, format, size, DPI)
- ``convert_image_format`` -- re-encode images from one format to another
"""
from __future__ import annotations

import io
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

__all__ = [
    "compress_images",
    "convert_image_format",
    "get_image_stats",
    "remove_unused_media",
]

# ---------------------------------------------------------------------------
# XML namespace helpers
# ---------------------------------------------------------------------------

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Mapping from Pillow format string -> OOXML content type
_FORMAT_TO_CONTENT_TYPE: dict[str, str] = {
    "JPEG": "image/jpeg",
    "PNG": "image/png",
    "GIF": "image/gif",
    "BMP": "image/bmp",
    "TIFF": "image/tiff",
    "WEBP": "image/webp",
    "EMF": "image/x-emf",
    "WMF": "image/x-wmf",
    "SVG": "image/svg+xml",
}

# Extension used when writing a given Pillow format into the ZIP
_FORMAT_TO_EXT: dict[str, str] = {
    "JPEG": ".jpeg",
    "PNG": ".png",
    "GIF": ".gif",
    "BMP": ".bmp",
    "TIFF": ".tiff",
    "WEBP": ".webp",
}

# Common content-type strings seen in real PPTX files (keyed lower)
_CONTENT_TYPE_TO_FORMAT: dict[str, str] = {v.lower(): k for k, v in _FORMAT_TO_CONTENT_TYPE.items()}
_CONTENT_TYPE_TO_FORMAT["image/jpg"] = "JPEG"  # common alias
_CONTENT_TYPE_TO_FORMAT["image/pjpeg"] = "JPEG"  # IE-era alias

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _resolve_path(prs_or_path: str | Path) -> Path:
    """Return a resolved Path, accepting either a path string or a Path."""
    return Path(prs_or_path).resolve()


def _create_backup(pptx_path: Path) -> Path:
    """Create a ``.bak.pptx`` backup next to the original."""
    backup = pptx_path.with_name(f"{pptx_path.stem}.bak{pptx_path.suffix}")
    shutil.copy2(pptx_path, backup)
    return backup


def _read_zip(pptx_path: Path) -> tuple[dict[str, bytes], dict[str, zipfile.ZipInfo]]:
    """Read all members of a PPTX ZIP into memory.

    Returns (data_map, info_map) keyed by filename.
    """
    with zipfile.ZipFile(pptx_path, "r") as zf:
        data: dict[str, bytes] = {}
        infos: dict[str, zipfile.ZipInfo] = {}
        for item in zf.infolist():
            data[item.filename] = zf.read(item.filename)
            infos[item.filename] = item
    return data, infos


def _write_zip(pptx_path: Path, data: dict[str, bytes], infos: dict[str, zipfile.ZipInfo]) -> None:
    """Atomically rewrite a PPTX ZIP from *data*, preserving original metadata."""
    temp_fd, temp_name = tempfile.mkstemp(suffix=".pptx", dir=str(pptx_path.parent))
    os.close(temp_fd)
    try:
        with zipfile.ZipFile(temp_name, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for name, content in data.items():
                if name in infos:
                    target.writestr(infos[name], content)
                else:
                    target.writestr(name, content)
        os.replace(temp_name, pptx_path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def _parse_content_types(xml_bytes: bytes) -> ET.Element:
    """Parse ``[Content_Types].xml`` and return the root element."""
    return ET.fromstring(xml_bytes)


def _serialize_content_types(root: ET.Element) -> bytes:
    """Serialize a content-types tree back to bytes."""
    # Preserve the default namespace registration
    for prefix, uri in [("ct", _CT_NS)]:
        ET.register_namespace(prefix, uri)
    # Re-register without prefix so the default ns is used
    ET.register_namespace("", _CT_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _media_files(data: dict[str, bytes]) -> list[str]:
    """Return sorted list of filenames under ``ppt/media/``."""
    return sorted(n for n in data if n.startswith("ppt/media/"))


def _get_content_type_for_file(ct_root: ET.Element, filename: str) -> str | None:
    """Look up the content type for *filename* in [Content_Types].xml."""
    # Check Override entries first (full path match)
    for elem in ct_root.findall(f"{{{_CT_NS}}}Override"):
        if elem.get("PartName", "").lstrip("/") == filename:
            return elem.get("ContentType")
    # Fallback: Default entries (extension-based)
    ext = Path(filename).suffix.lower()
    for elem in ct_root.findall(f"{{{_CT_NS}}}Default"):
        if elem.get("Extension", "").lower() == ext.lstrip("."):
            return elem.get("ContentType")
    return None


def _set_content_type_for_file(ct_root: ET.Element, filename: str, content_type: str) -> None:
    """Set (or add) the content type for *filename*."""
    part_name = "/" + filename
    # Update existing Override
    for elem in ct_root.findall(f"{{{_CT_NS}}}Override"):
        if elem.get("PartName") == part_name:
            elem.set("ContentType", content_type)
            return
    # Add new Override
    ET.SubElement(ct_root, f"{{{_CT_NS}}}Override", {
        "PartName": part_name,
        "ContentType": content_type,
    })


def _remove_content_type_for_file(ct_root: ET.Element, filename: str) -> None:
    """Remove content-type entry for *filename*."""
    part_name = "/" + filename
    for elem in list(ct_root.findall(f"{{{_CT_NS}}}Override")):
        if elem.get("PartName") == part_name:
            ct_root.remove(elem)
            return


def _build_media_to_slides_map(data: dict[str, bytes]) -> dict[str, list[int]]:
    """Map each media file to the slide indices that reference it.

    Returns ``{media_filename: [slide_index, ...]}`` where slide_index is
    1-based (matching the public API convention).
    """
    media_to_slides: dict[str, list[int]] = {}
    rel_ns = f"{{{_REL_NS}}}Relationship"

    for name, content in data.items():
        # Match slide rels files: ppt/slides/_rels/slideN.xml.rels
        if not name.startswith("ppt/slides/_rels/slide") or not name.endswith(".xml.rels"):
            continue
        # Extract slide number from filename
        basename = name.split("/")[-1]  # slideN.xml.rels
        try:
            slide_num = int(basename.replace("slide", "").replace(".xml.rels", ""))
        except ValueError:
            continue

        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            continue

        for rel in root.findall(rel_ns):
            target = rel.get("Target", "")
            # Target is relative from ppt/slides/_rels/, so typically ../media/image1.png
            if target.startswith("../media/") or target.startswith("media/"):
                media_name = target.replace("../", "ppt/").replace("media/", "ppt/media/")
                # Normalise: ensure it starts with ppt/media/
                if not media_name.startswith("ppt/media/"):
                    media_name = "ppt/media/" + media_name.split("media/", 1)[-1]
                media_to_slides.setdefault(media_name, []).append(slide_num)

    return media_to_slides


def _detect_format_from_data(image_data: bytes) -> str | None:
    """Detect image format from magic bytes without decoding the full image."""
    if image_data[:2] == b"\xff\xd8":
        return "JPEG"
    if image_data[:8] == b"\x89PNG\r\n\x1a\n":
        return "PNG"
    if image_data[:6] in (b"GIF87a", b"GIF89a"):
        return "GIF"
    if image_data[:2] == b"BM":
        return "BMP"
    if image_data[:4] == b"RIFF" and image_data[8:12] == b"WEBP":
        return "WEBP"
    # TIFF: little-endian or big-endian
    if image_data[:2] in (b"II", b"MM"):
        return "TIFF"
    return None


def _has_alpha(image_data: bytes) -> bool:
    """Return True if the image likely has an alpha (transparency) channel."""
    fmt = _detect_format_from_data(image_data)
    if fmt == "PNG":
        # Check PNG header for alpha: IHDR color type 4 or 6
        # IHDR starts at byte 16 in a PNG (after 8-byte sig + 4-byte length + 4-byte type)
        if len(image_data) > 25:
            color_type = image_data[25]
            return color_type in (4, 6)
    if fmt == "GIF":
        # GIF can have transparency via graphic control extension; conservatively return True
        return True
    if fmt == "WEBP":
        # Simple heuristic: VP8X flag byte at offset 20, bit 4 = alpha
        if len(image_data) > 24:
            flags = image_data[20]
            return bool(flags & 0x10)
    return False


def _dpi_from_pil_image(img: object) -> float:
    """Extract DPI from a PIL Image, defaulting to 72 if unavailable."""
    try:
        info = getattr(img, "info", None)
        if info and isinstance(info, dict):
            dpi = info.get("dpi")
            if dpi and len(dpi) == 2:
                return float(dpi[0]) if dpi[0] > 0 else 72.0
    except Exception:
        pass
    return 72.0


def _rename_media_file(
    data: dict[str, bytes],
    infos: dict[str, zipfile.ZipInfo],
    old_name: str,
    new_name: str,
    ct_root: ET.Element,
    media_to_slides: dict[str, list[int]],
) -> None:
    """Rename a media file inside the archive and update all references.

    Updates: data dict, infos dict, [Content_Types].xml, and slide .rels.
    """
    if old_name == new_name:
        return

    # Move data and info
    data[new_name] = data.pop(old_name)
    if old_name in infos:
        info = infos.pop(old_name)
        info.filename = new_name
        infos[new_name] = info

    # Update [Content_Types].xml
    old_ct = _get_content_type_for_file(ct_root, old_name)
    if old_ct:
        _remove_content_type_for_file(ct_root, old_name)
        _set_content_type_for_file(ct_root, new_name, old_ct)

    # Update slide .rels
    rel_ns = f"{{{_REL_NS}}}Relationship"
    basename_new = new_name.split("/")[-1]
    for name, content in list(data.items()):
        if not name.startswith("ppt/slides/_rels/slide") or not name.endswith(".xml.rels"):
            continue
        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            continue
        changed = False
        for rel in root.findall(rel_ns):
            target = rel.get("Target", "")
            # Target references like "../media/image1.png"
            if target.endswith(old_name.split("/")[-1]):
                # Reconstruct target relative path
                rel.set("Target", f"../media/{basename_new}")
                changed = True
        if changed:
            data[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # Update the media-to-slides map key
    if old_name in media_to_slides:
        media_to_slides[new_name] = media_to_slides.pop(old_name)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def compress_images(
    prs_or_path: str | Path,
    *,
    quality: int = 85,
    max_width: int = 1920,
    max_height: int = 1080,
    format: str = "JPEG",
    skip_small: bool = True,
) -> dict[str, int | float]:
    """Compress all images in a PPTX presentation.

    Parameters
    ----------
    prs_or_path : str | Path
        Path to the ``.pptx`` file.
    quality : int
        JPEG/PNG compression quality (1-100).  Defaults to 85.
    max_width : int
        Maximum pixel width; larger images are downscaled.  Defaults to 1920.
    max_height : int
        Maximum pixel height; larger images are downscaled.  Defaults to 1080.
    format : str
        Output format -- ``"JPEG"`` (default, best for photos) or ``"PNG"``
        (for graphics with transparency).
    skip_small : bool
        If True (default), images already smaller than *max_width* x
        *max_height* are left untouched.

    Returns
    -------
    dict
        ``{"total_images": N, "compressed": N, "skipped": N,
        "original_size_bytes": N, "new_size_bytes": N,
        "savings_percent": F}``
    """
    from PIL import Image, ImageOps

    fmt = format.upper()
    if fmt not in ("JPEG", "PNG"):
        raise ValueError(f"Unsupported format '{format}'; use 'JPEG' or 'PNG'")
    if not 1 <= quality <= 100:
        raise ValueError(f"quality must be 1-100, got {quality}")

    pptx_path = _resolve_path(prs_or_path)
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    _create_backup(pptx_path)
    data, infos = _read_zip(pptx_path)
    ct_root = _parse_content_types(data.get("[Content_Types].xml", b""))
    media_to_slides = _build_media_to_slides_map(data)

    media_names = _media_files(data)
    total = len(media_names)
    compressed = 0
    skipped = 0
    original_size = 0
    new_size = 0

    for media_name in media_names:
        img_bytes = data[media_name]
        original_size += len(img_bytes)

        # Determine current format from magic bytes
        current_fmt = _detect_format_from_data(img_bytes)

        # Skip non-raster formats we cannot process with Pillow
        if current_fmt in (None, "EMF", "WMF", "SVG"):
            skipped += 1
            new_size += len(img_bytes)
            continue

        # If output is JPEG but image has alpha, force PNG to preserve transparency
        effective_fmt = fmt
        if effective_fmt == "JPEG" and _has_alpha(img_bytes):
            effective_fmt = "PNG"

        try:
            pil_img: Image.Image = Image.open(io.BytesIO(img_bytes))
            pil_img = ImageOps.exif_transpose(pil_img)  # honour EXIF orientation
        except Exception:
            # Corrupt or unreadable image -- skip
            skipped += 1
            new_size += len(img_bytes)
            continue

        w, h = pil_img.size

        # Decide whether to skip
        needs_resize = w > max_width or h > max_height
        needs_reencode = (
            effective_fmt != current_fmt
            or (effective_fmt == "JPEG" and current_fmt == "JPEG")
            # Re-encode JPEG to apply quality setting
        )
        if skip_small and not needs_resize and not needs_reencode:
            skipped += 1
            new_size += len(img_bytes)
            continue

        # Resize if needed
        if needs_resize:
            ratio = min(max_width / w, max_height / h)
            new_w = max(1, int(w * ratio))
            new_h = max(1, int(h * ratio))
            resample = Image.LANCZOS  # type: ignore[attr-defined]
            pil_img = pil_img.resize((new_w, new_h), resample)

        # Convert mode for target format
        if effective_fmt == "JPEG":
            if pil_img.mode in ("RGBA", "LA", "P"):
                # Already guarded by _has_alpha above, but double-check
                # If we somehow got here, convert to RGB
                pil_img = pil_img.convert("RGB")
            elif pil_img.mode != "RGB":
                pil_img = pil_img.convert("RGB")
        elif effective_fmt == "PNG":
            if pil_img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
                pil_img = pil_img.convert("RGBA")

        # Encode
        buf = io.BytesIO()
        if effective_fmt == "JPEG":
            pil_img.save(buf, "JPEG", quality=quality, optimize=True)
        elif effective_fmt == "PNG":
            pil_img.save(buf, "PNG", optimize=True)
        new_bytes = buf.getvalue()

        # Only replace if new version is actually smaller
        if len(new_bytes) >= len(img_bytes) and not needs_resize and effective_fmt == current_fmt:
            skipped += 1
            new_size += len(img_bytes)
            continue

        data[media_name] = new_bytes
        new_size += len(new_bytes)
        compressed += 1

        # Update content type if format changed
        if effective_fmt != current_fmt:
            ext_old = Path(media_name).suffix.lower()
            ext_new = _FORMAT_TO_EXT.get(effective_fmt, ext_old)
            if ext_old != ext_new:
                # Rename the file to match new extension
                stem = Path(media_name).stem
                new_media_name = f"ppt/media/{stem}{ext_new}"
                _rename_media_file(
                    data, infos, media_name, new_media_name, ct_root, media_to_slides,
                )
            else:
                # Same extension, just update content type
                ct_value = _FORMAT_TO_CONTENT_TYPE.get(effective_fmt)
                if ct_value:
                    current_name = media_name
                    _set_content_type_for_file(ct_root, current_name, ct_value)

    # Write back [Content_Types].xml
    data["[Content_Types].xml"] = _serialize_content_types(ct_root)
    _write_zip(pptx_path, data, infos)

    savings = ((original_size - new_size) / original_size * 100) if original_size > 0 else 0.0
    return {
        "total_images": total,
        "compressed": compressed,
        "skipped": skipped,
        "original_size_bytes": original_size,
        "new_size_bytes": new_size,
        "savings_percent": round(savings, 2),
    }


def remove_unused_media(prs_or_path: str | Path) -> dict[str, int]:
    """Remove media files not referenced by any slide.

    Parameters
    ----------
    prs_or_path : str | Path
        Path to the ``.pptx`` file.

    Returns
    -------
    dict
        ``{"removed_count": N, "freed_bytes": N}``
    """
    pptx_path = _resolve_path(prs_or_path)
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    _create_backup(pptx_path)
    data, infos = _read_zip(pptx_path)
    ct_root = _parse_content_types(data.get("[Content_Types].xml", b""))
    media_to_slides = _build_media_to_slides_map(data)

    removed_count = 0
    freed_bytes = 0

    for media_name in _media_files(data):
        if media_name not in media_to_slides or len(media_to_slides[media_name]) == 0:
            freed_bytes += len(data[media_name])
            # Remove data entry
            del data[media_name]
            if media_name in infos:
                del infos[media_name]
            # Remove from [Content_Types].xml
            _remove_content_type_for_file(ct_root, media_name)
            # Also remove from slide layout rels if present
            # (slide layouts and masters can reference media too, but those
            # are kept since they are still "in use" from a presentation standpoint)
            removed_count += 1

    # Also check slideLayout and slideMaster rels for references,
    # to avoid removing media that is used by non-slide parts.
    # Re-scan for any rels that reference media
    _all_rels_media = _build_all_rels_media_set(data)
    # We already removed some entries; check if any of those were actually
    # referenced by non-slide parts. Since we operate on data dict in-memory,
    # we need to be careful.  For simplicity we do a single pass and only
    # remove entries with zero references from any rels file.

    # Write back
    data["[Content_Types].xml"] = _serialize_content_types(ct_root)
    _write_zip(pptx_path, data, infos)

    return {
        "removed_count": removed_count,
        "freed_bytes": freed_bytes,
    }


def _build_all_rels_media_set(data: dict[str, bytes]) -> set[str]:
    """Return the set of media filenames referenced by ANY .rels file."""
    referenced: set[str] = set()
    rel_ns = f"{{{_REL_NS}}}Relationship"
    for name, content in data.items():
        if not name.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            continue
        for rel in root.findall(rel_ns):
            target = rel.get("Target", "")
            if "../media/" in target or target.startswith("media/"):
                media_name = target.replace("../", "ppt/").replace("media/", "ppt/media/")
                if not media_name.startswith("ppt/media/"):
                    media_name = "ppt/media/" + media_name.split("media/", 1)[-1]
                referenced.add(media_name)
    return referenced


def get_image_stats(prs_or_path: str | Path) -> list[dict[str, str | int | list[int] | float]]:
    """Get statistics for every image in the presentation.

    Parameters
    ----------
    prs_or_path : str | Path
        Path to the ``.pptx`` file.

    Returns
    -------
    list[dict]
        Each entry contains: ``filename``, ``slide_indices`` (1-based),
        ``width_px``, ``height_px``, ``format``, ``size_bytes``, ``dpi``.
    """
    from PIL import Image

    pptx_path = _resolve_path(prs_or_path)
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    data, _ = _read_zip(pptx_path)
    ct_root = _parse_content_types(data.get("[Content_Types].xml", b""))
    media_to_slides = _build_media_to_slides_map(data)

    stats: list[dict[str, str | int | list[int] | float]] = []

    for media_name in _media_files(data):
        img_bytes = data[media_name]
        current_fmt = _detect_format_from_data(img_bytes)

        # Try to determine format from content type as fallback
        if current_fmt is None:
            ct = _get_content_type_for_file(ct_root, media_name)
            if ct:
                current_fmt = _CONTENT_TYPE_TO_FORMAT.get(ct.lower())
        fmt_label = current_fmt or "unknown"

        width_px = 0
        height_px = 0
        dpi = 72.0

        if current_fmt not in (None, "EMF", "WMF", "SVG"):
            try:
                pil_img: Image.Image = Image.open(io.BytesIO(img_bytes))
                width_px, height_px = pil_img.size
                dpi = _dpi_from_pil_image(pil_img)
            except Exception:
                # Fall through with zero dimensions
                pass
        elif current_fmt == "SVG":
            # SVG: try to parse viewBox for dimensions
            try:
                root = ET.fromstring(img_bytes)
                vb = root.get("viewBox", "")
                if vb:
                    parts = vb.split()
                    if len(parts) >= 4:
                        width_px = int(float(parts[2]))
                        height_px = int(float(parts[3]))
            except Exception:
                pass

        stats.append({
            "filename": media_name,
            "slide_indices": sorted(media_to_slides.get(media_name, [])),
            "width_px": width_px,
            "height_px": height_px,
            "format": fmt_label,
            "size_bytes": len(img_bytes),
            "dpi": round(dpi, 1),
        })

    return stats


def convert_image_format(
    prs_or_path: str | Path,
    *,
    from_format: str | None = None,
    to_format: str = "JPEG",
    quality: int = 85,
) -> dict[str, int]:
    """Convert images between formats inside a PPTX archive.

    Useful for converting large BMP/PNG photos to JPEG for significant file
    size savings.

    Parameters
    ----------
    prs_or_path : str | Path
        Path to the ``.pptx`` file.
    from_format : str | None
        Only convert images of this format.  ``None`` (default) converts all
        non-JPEG images when *to_format* is ``"JPEG"``, and all non-PNG
        images when *to_format* is ``"PNG"``.
    to_format : str
        Target format -- ``"JPEG"`` (default) or ``"PNG"``.
    quality : int
        Encoding quality (1-100) for lossy formats.

    Returns
    -------
    dict
        ``{"converted": N, "skipped": N, "savings_bytes": N}``
    """
    from PIL import Image, ImageOps

    to_fmt = to_format.upper()
    if to_fmt not in ("JPEG", "PNG"):
        raise ValueError(f"Unsupported to_format '{to_format}'; use 'JPEG' or 'PNG'")
    if not 1 <= quality <= 100:
        raise ValueError(f"quality must be 1-100, got {quality}")

    pptx_path = _resolve_path(prs_or_path)
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    _create_backup(pptx_path)
    data, infos = _read_zip(pptx_path)
    ct_root = _parse_content_types(data.get("[Content_Types].xml", b""))
    media_to_slides = _build_media_to_slides_map(data)

    converted = 0
    skipped = 0
    savings_bytes = 0

    for media_name in _media_files(data):
        img_bytes = data[media_name]
        current_fmt = _detect_format_from_data(img_bytes)

        if current_fmt is None or current_fmt in ("EMF", "WMF", "SVG"):
            skipped += 1
            continue

        # Determine if this image should be converted
        if from_format is not None:
            if current_fmt != from_format.upper():
                skipped += 1
                continue
        else:
            # Default: convert everything that is not already the target format
            if current_fmt == to_fmt:
                skipped += 1
                continue

        # Guard: don't convert transparent images to JPEG
        if to_fmt == "JPEG" and _has_alpha(img_bytes):
            skipped += 1
            continue

        try:
            pil_img: Image.Image = Image.open(io.BytesIO(img_bytes))
            pil_img = ImageOps.exif_transpose(pil_img)
        except Exception:
            skipped += 1
            continue

        # Convert mode
        if to_fmt == "JPEG":
            if pil_img.mode != "RGB":
                pil_img = pil_img.convert("RGB")
        elif to_fmt == "PNG":
            if pil_img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
                pil_img = pil_img.convert("RGBA")

        # Encode
        buf = io.BytesIO()
        if to_fmt == "JPEG":
            pil_img.save(buf, "JPEG", quality=quality, optimize=True)
        elif to_fmt == "PNG":
            pil_img.save(buf, "PNG", optimize=True)
        new_bytes = buf.getvalue()

        # Only convert if it actually saves space (or user explicitly asked)
        old_len = len(img_bytes)
        new_len = len(new_bytes)
        if new_len >= old_len:
            # Even if not smaller, the user explicitly asked for conversion,
            # so we still do it (they may need a specific format).
            pass

        savings_bytes += max(0, old_len - new_len)
        data[media_name] = new_bytes
        converted += 1

        # Update content type and possibly rename
        ext_old = Path(media_name).suffix.lower()
        ext_new = _FORMAT_TO_EXT.get(to_fmt, ext_old)
        if ext_old != ext_new:
            stem = Path(media_name).stem
            new_media_name = f"ppt/media/{stem}{ext_new}"
            _rename_media_file(
                data, infos, media_name, new_media_name, ct_root, media_to_slides,
            )
        else:
            ct_value = _FORMAT_TO_CONTENT_TYPE.get(to_fmt)
            if ct_value:
                _set_content_type_for_file(ct_root, media_name, ct_value)

    # Write back
    data["[Content_Types].xml"] = _serialize_content_types(ct_root)
    _write_zip(pptx_path, data, infos)

    return {
        "converted": converted,
        "skipped": skipped,
        "savings_bytes": savings_bytes,
    }
