"""PPTX export to multiple formats: PDF, images, HTML, thumbnails, and text.

This module provides export functions that convert a PPTX presentation into
various output formats.  It leverages the existing ``preview_renderer``
infrastructure (LibreOffice / PyMuPDF / Windows COM) for rasterisation-heavy
exports and builds HTML / text output directly from the python-pptx object
model.

Public API
----------
- :func:`export_to_pdf`        -- PPTX -> PDF
- :func:`export_to_images`     -- PPTX -> per-slide image files
- :func:`export_to_html`       -- PPTX -> self-contained HTML slideshow
- :func:`export_thumbnails`    -- PPTX -> small thumbnail images
- :func:`export_to_text`       -- PPTX -> plain-text / Markdown
"""
from __future__ import annotations

import base64
import os
import shutil
import tempfile
from typing import Literal

from pptx import Presentation as _open_presentation
from pptx.presentation import Presentation as _PresentationCls

from pptx_skill.preview_renderer import find_soffice, render_preview

__all__ = [
    "export_to_pdf",
    "export_to_images",
    "export_to_html",
    "export_thumbnails",
    "export_to_text",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _resolve_presentation(prs_or_path: str | os.PathLike[str] | _PresentationCls) -> tuple[_PresentationCls, str | None]:
    """Return ``(Presentation, temp_dir_or_None)``.

    If *prs_or_path* is already a :class:`Presentation` instance it is used
    directly and *temp_dir* is ``None``.  If it is a path, the file is opened
    and *temp_dir* remains ``None`` (the caller is not responsible for
    cleanup).  The helper exists mainly to centralise the isinstance branch.
    """
    if isinstance(prs_or_path, _PresentationCls):
        return prs_or_path, None
    path = os.fspath(prs_or_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"PPTX file not found: {path}")
    return _open_presentation(path), None


def _save_to_temp(prs: _PresentationCls) -> tuple[str, str]:
    """Save *prs* to a temporary PPTX file and return ``(path, tmp_dir)``.

    The caller should clean up *tmp_dir* when done.
    """
    tmp_dir = tempfile.mkdtemp(prefix="pptx_skill_export_")
    tmp_path = os.path.join(tmp_dir, "presentation.pptx")
    prs.save(tmp_path)
    return tmp_path, tmp_dir


def _slide_range_indices(total: int, slide_range: tuple[int, int] | None) -> list[int]:
    """Return 0-based slide indices for the given *slide_range* (1-based, inclusive).

    If *slide_range* is ``None``, returns ``list(range(total))``.
    """
    if slide_range is None:
        return list(range(total))
    start, end = slide_range
    # Clamp to valid bounds
    start = max(1, start)
    end = min(total, end)
    if start > end:
        return []
    return list(range(start - 1, end))


def _ensure_dir(path: str | os.PathLike[str]) -> str:
    """Create directory if needed and return the absolute path."""
    abs_path = os.path.abspath(path)
    os.makedirs(abs_path, exist_ok=True)
    return abs_path


# ---------------------------------------------------------------------------
# 1. PDF export
# ---------------------------------------------------------------------------

def export_to_pdf(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
    output_path: str | os.PathLike[str],
    *,
    dpi: int = 150,
    range: tuple[int, int] | None = None,
) -> str:
    """Export a PPTX presentation to PDF.

    Uses the existing ``preview_renderer`` infrastructure (LibreOffice /
    PyMuPDF / Windows COM) for conversion.  When LibreOffice is available it
    produces a high-fidelity PDF directly.  Otherwise PyMuPDF (fitz) is used
    to assemble a PDF from rendered page images.

    Parameters
    ----------
    prs_or_path :
        A :class:`~pptx.Presentation` instance or a path to a ``.pptx`` file.
    output_path :
        Destination path for the PDF file.
    dpi :
        Resolution for raster-based fallback paths (default 150).
    range :
        Optional ``(start, end)`` 1-based slide indices (inclusive) to
        export.  ``None`` means all slides.

    Returns
    -------
    str
        Absolute path of the written PDF file.

    Raises
    ------
    FileNotFoundError
        If *prs_or_path* is a path that does not exist.
    RuntimeError
        If no rendering backend is available or conversion fails.
    """
    prs, _ = _resolve_presentation(prs_or_path)
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    slide_indices = _slide_range_indices(len(prs.slides), range)
    if not slide_indices:
        raise ValueError("No slides to export (empty range or empty presentation)")

    # If we need a sub-range, we must create a trimmed copy first because
    # the rendering backends operate on whole files.
    need_trim = range is not None and len(slide_indices) < len(prs.slides)
    tmp_dir: str | None = None
    pptx_path: str

    if need_trim:
        pptx_path, tmp_dir = _save_to_temp_trimmed(prs, slide_indices)
    elif isinstance(prs_or_path, (str, os.PathLike)):
        pptx_path = os.fspath(prs_or_path)
    else:
        pptx_path, tmp_dir = _save_to_temp(prs)

    try:
        result = _convert_to_pdf(pptx_path, output_path, dpi)
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    return result


def _save_to_temp_trimmed(prs: _PresentationCls, indices: list[int]) -> tuple[str, str]:
    """Save only the slides at *indices* to a temporary PPTX and return ``(path, tmp_dir)``."""
    tmp_dir = tempfile.mkdtemp(prefix="pptx_skill_export_")
    tmp_path = os.path.join(tmp_dir, "trimmed.pptx")
    prs.save(tmp_path)

    # Re-open and prune slides we do not want
    trimmed = _open_presentation(tmp_path)
    index_set = set(indices)
    # Collect slides to remove (iterate in reverse to keep indices stable)
    rmlist = [
        i for i in range(len(trimmed.slides)) if i not in index_set
    ]
    for i in sorted(rmlist, reverse=True):
        sldIdLst = trimmed.slides._sldIdLst
        sldId = sldIdLst[i]
        # Find the relationship id for this slide
        rid = None
        for attr_name in sldId.attrib:
            if attr_name.endswith("}id") or attr_name == "id":
                rid = sldId.get(attr_name)
                break
        if rid is None:
            for k, v in sldId.attrib.items():
                if "id" in k.lower():
                    rid = v
                    break

        if rid:
            trimmed.part.drop_rel(rid)
        sldIdLst.remove(sldId)

    trimmed.save(tmp_path)
    return tmp_path, tmp_dir


def _convert_to_pdf(pptx_path: str, output_path: str, dpi: int) -> str:
    """Attempt to convert *pptx_path* to PDF at *output_path*.

    Tries LibreOffice first (direct PDF conversion), then falls back to
    PyMuPDF image-based PDF assembly.  Collects diagnostics from each
    failed strategy so the final error is actionable.
    """
    _errors: list[str] = []

    # --- Strategy 1: LibreOffice direct PDF conversion ----------------------
    soffice = find_soffice()
    if soffice:
        import subprocess
        abs_pptx = os.path.abspath(pptx_path)
        tmp_dir = tempfile.mkdtemp(prefix="pptx_skill_pdf_")
        user_install = os.path.join(tmp_dir, "profile")
        os.makedirs(user_install, exist_ok=True)
        try:
            cmd = [
                soffice,
                f"-env:UserInstallation=file:///{user_install.replace(os.sep, '/')}",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", tmp_dir,
                abs_pptx,
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if proc.returncode == 0:
                # Find the generated PDF
                import glob
                pdfs = glob.glob(os.path.join(tmp_dir, "*.pdf"))
                if pdfs:
                    # Pick the most recently created
                    pdfs.sort(key=os.path.getmtime, reverse=True)
                    shutil.copy2(pdfs[0], output_path)
                    return output_path
                else:
                    _errors.append("LibreOffice: conversion succeeded but no PDF found")
            else:
                _errors.append(f"LibreOffice exit {proc.returncode}: {(proc.stderr or '')[-300:]}")
        except Exception as e:
            _errors.append(f"LibreOffice: {e}")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
    else:
        _errors.append("LibreOffice: not found on PATH")

    # --- Strategy 2: Render to images, then assemble PDF via PyMuPDF -------
    try:
        import fitz  # PyMuPDF
    except ImportError:
        _errors.append("PyMuPDF: not installed (pip install PyMuPDF)")
    else:
        img_dir = tempfile.mkdtemp(prefix="pptx_skill_pdf_imgs_")
        try:
            render_result = render_preview(pptx_path, img_dir, dpi=dpi)
            if render_result.slide_pngs:
                with fitz.open() as doc:
                    for png_path in render_result.slide_pngs:
                        with fitz.open(png_path) as img_doc:
                            page = doc.new_page(width=img_doc[0].rect.width, height=img_doc[0].rect.height)
                            page.insert_image(page.rect, filename=png_path)
                    doc.save(output_path)
                return output_path
            else:
                _errors.append("PyMuPDF: render produced no slide images")
        except Exception as e:
            _errors.append(f"PyMuPDF: {e}")
        finally:
            shutil.rmtree(img_dir, ignore_errors=True)

    # --- Strategy 3: Pillow-based image-to-PDF assembly --------------------
    try:
        from PIL import Image
    except ImportError:
        _errors.append("Pillow: not installed")
    else:
        img_dir = tempfile.mkdtemp(prefix="pptx_skill_pdf_imgs_")
        try:
            render_result = render_preview(pptx_path, img_dir, dpi=dpi)
            if render_result.slide_pngs:
                images: list[Image.Image] = []
                for png_path in render_result.slide_pngs:
                    with Image.open(png_path) as img_src:
                        img: Image.Image = img_src
                        if img.mode == "RGBA":
                            # PDF does not support RGBA; convert to RGB with white bg
                            bg = Image.new("RGB", img.size, (255, 255, 255))
                            bg.paste(img, mask=img.split()[3])
                            img = bg
                        elif img.mode != "RGB":
                            img = img.convert("RGB")
                        images.append(img.copy())
                if images:
                    first = images[0]
                    rest = images[1:]
                    first.save(output_path, "PDF", save_all=True, append_images=rest)
                    return output_path
            else:
                _errors.append("Pillow: render produced no slide images")
        except Exception as e:
            _errors.append(f"Pillow: {e}")
        finally:
            shutil.rmtree(img_dir, ignore_errors=True)

    raise RuntimeError(
        "PDF export failed — all strategies exhausted. Diagnostics:\n  - " +
        "\n  - ".join(_errors) +
        "\nInstall LibreOffice, PyMuPDF (pip install PyMuPDF), or Pillow to enable PDF export."
    )


# ---------------------------------------------------------------------------
# 2. Image export
# ---------------------------------------------------------------------------

def export_to_images(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
    output_dir: str | os.PathLike[str],
    *,
    dpi: int = 150,
    format: Literal["PNG", "JPEG", "BMP"] = "PNG",
    range: tuple[int, int] | None = None,
) -> list[str]:
    """Export each slide as an image file.

    Uses the ``preview_renderer`` infrastructure to render slides, then
    optionally converts the output to the requested *format*.

    Parameters
    ----------
    prs_or_path :
        A :class:`~pptx.Presentation` instance or a path to a ``.pptx`` file.
    output_dir :
        Directory for the output image files.  Created if it does not exist.
    dpi :
        Resolution in dots per inch (default 150).
    format :
        Image format: ``"PNG"``, ``"JPEG"``, or ``"BMP"``.
    range :
        Optional ``(start, end)`` 1-based slide indices (inclusive).

    Returns
    -------
    list[str]
        Absolute paths of the written image files, in slide order.

    Raises
    ------
    FileNotFoundError
        If *prs_or_path* is a path that does not exist.
    RuntimeError
        If no rendering backend is available.
    """
    prs, _ = _resolve_presentation(prs_or_path)
    output_dir = _ensure_dir(output_dir)

    slide_indices = _slide_range_indices(len(prs.slides), range)
    if not slide_indices:
        return []

    # Determine file extension
    ext = format.lower()
    if ext == "jpeg":
        ext = "jpg"

    # Save to temp path for renderer
    tmp_dir: str | None = None
    if isinstance(prs_or_path, (str, os.PathLike)):
        pptx_path = os.fspath(prs_or_path)
    else:
        pptx_path, tmp_dir = _save_to_temp(prs)

    try:
        render_dir = tempfile.mkdtemp(prefix="pptx_skill_img_")
        render_result = render_preview(pptx_path, render_dir, dpi=dpi)
        if not render_result.slide_pngs:
            raise RuntimeError(
                "No rendering backend available for image export. "
                "Install LibreOffice, PyMuPDF, or use Windows COM."
            )

        output_paths: list[str] = []
        for i, idx in enumerate(slide_indices):
            if idx >= len(render_result.slide_pngs):
                break
            src = render_result.slide_pngs[idx]
            dst = os.path.join(output_dir, f"slide_{i + 1:03d}.{ext}")

            if format == "PNG" and src.lower().endswith(".png"):
                # Just copy / rename
                shutil.copy2(src, dst)
            else:
                # Convert via Pillow
                from PIL import Image as PILImage

                img: PILImage.Image = PILImage.open(src)
                with img:
                    if format == "JPEG" and img.mode == "RGBA":
                        bg = PILImage.new("RGB", img.size, (255, 255, 255))
                        bg.paste(img, mask=img.split()[3])
                        img = bg
                    elif format == "BMP" and img.mode == "RGBA":
                        bg = PILImage.new("RGB", img.size, (255, 255, 255))
                        bg.paste(img, mask=img.split()[3])
                        img = bg
                    elif format in ("JPEG", "BMP") and img.mode != "RGB":
                        img = img.convert("RGB")
                    img.save(dst, format=format)
            output_paths.append(dst)

        # Cleanup render dir
        shutil.rmtree(render_dir, ignore_errors=True)
        return output_paths

    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# 3. HTML export
# ---------------------------------------------------------------------------

_LIGHT_CSS = """
:root {
    --bg: #ffffff;
    --slide-bg: #ffffff;
    --text: #1a1a1a;
    --text-muted: #666666;
    --border: #e0e0e0;
    --accent: #2563eb;
    --nav-bg: #f5f5f5;
    --progress-bg: #e0e0e0;
    --progress-fg: #2563eb;
    --shadow: 0 2px 8px rgba(0,0,0,0.08);
}
"""

_DARK_CSS = """
:root {
    --bg: #111111;
    --slide-bg: #1e1e1e;
    --text: #e0e0e0;
    --text-muted: #999999;
    --border: #333333;
    --accent: #60a5fa;
    --nav-bg: #1a1a1a;
    --progress-bg: #333333;
    --progress-fg: #60a5fa;
    --shadow: 0 2px 8px rgba(0,0,0,0.3);
}
"""

_BASE_CSS = """
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body { width: 100%; height: 100%; overflow: hidden; background: var(--bg); color: var(--text); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }

.slide-container {
    position: relative;
    width: 100%;
    height: calc(100% - 48px);
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
}

.slide {
    position: absolute;
    width: 960px;
    height: 540px;
    background: var(--slide-bg);
    box-shadow: var(--shadow);
    transform-origin: center center;
    display: none;
    overflow: hidden;
    border-radius: 4px;
    border: 1px solid var(--border);
}
.slide.active { display: block; }

.slide .element {
    position: absolute;
    overflow: hidden;
    word-wrap: break-word;
    overflow-wrap: break-word;
}
.slide .element img {
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
}
.slide table {
    border-collapse: collapse;
    width: 100%;
    font-size: 14px;
}
.slide table th,
.slide table td {
    border: 1px solid var(--border);
    padding: 6px 10px;
    text-align: left;
}
.slide table th {
    background: var(--nav-bg);
    font-weight: 600;
}

.navbar {
    height: 48px;
    background: var(--nav-bg);
    border-top: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 20px;
    user-select: none;
}
.navbar .nav-btn {
    background: none;
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 6px 16px;
    cursor: pointer;
    color: var(--text);
    font-size: 14px;
    transition: background 0.15s;
}
.navbar .nav-btn:hover { background: var(--border); }
.navbar .nav-btn:disabled { opacity: 0.3; cursor: default; }
.navbar .slide-counter { font-size: 14px; color: var(--text-muted); min-width: 80px; text-align: center; }

.progress-bar {
    position: fixed;
    bottom: 48px;
    left: 0;
    height: 3px;
    background: var(--progress-fg);
    transition: width 0.3s ease;
    z-index: 10;
}

.notes-panel {
    position: fixed;
    bottom: 48px;
    left: 0;
    right: 0;
    max-height: 150px;
    background: var(--nav-bg);
    border-top: 1px solid var(--border);
    padding: 10px 20px;
    overflow-y: auto;
    font-size: 13px;
    color: var(--text-muted);
    display: none;
    z-index: 5;
}
.notes-panel.visible { display: block; }

.slide-nav {
    position: fixed;
    bottom: 56px;
    right: 12px;
    display: flex;
    flex-direction: column;
    gap: 4px;
    z-index: 10;
}
.slide-nav .thumb {
    width: 80px;
    height: 45px;
    border: 2px solid transparent;
    border-radius: 4px;
    cursor: pointer;
    opacity: 0.5;
    transition: opacity 0.2s, border-color 0.2s;
    background: var(--slide-bg);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    color: var(--text-muted);
}
.slide-nav .thumb:hover { opacity: 0.8; }
.slide-nav .thumb.active { border-color: var(--accent); opacity: 1; }

@media (max-width: 1100px) {
    .slide { width: 640px; height: 360px; }
}
@media (max-width: 720px) {
    .slide { width: 480px; height: 270px; }
    .slide-nav { display: none; }
}
"""

_JS = """
(function() {
    var current = 0;
    var total = 0;
    var slides = [];
    var thumbs = [];
    var progressBar = document.getElementById('progress-bar');
    var counterEl = document.getElementById('slide-counter');
    var notesPanel = document.getElementById('notes-panel');
    var prevBtn = document.getElementById('btn-prev');
    var nextBtn = document.getElementById('btn-next');

    function init() {
        slides = document.querySelectorAll('.slide');
        thumbs = document.querySelectorAll('.slide-nav .thumb');
        total = slides.length;
        if (total === 0) return;
        showSlide(0);
        window.addEventListener('keydown', onKey);
        // Touch support
        var touchStartX = 0;
        document.addEventListener('touchstart', function(e) { touchStartX = e.changedTouches[0].screenX; }, {passive: true});
        document.addEventListener('touchend', function(e) {
            var dx = e.changedTouches[0].screenX - touchStartX;
            if (Math.abs(dx) > 50) { dx > 0 ? prev() : next(); }
        }, {passive: true});
        // Click navigation: left half = prev, right half = next
        document.querySelector('.slide-container').addEventListener('click', function(e) {
            var rect = this.getBoundingClientRect();
            if (e.clientX < rect.left + rect.width / 2) { prev(); } else { next(); }
        });
        scaleSlides();
        window.addEventListener('resize', scaleSlides);
    }

    function scaleSlides() {
        var container = document.querySelector('.slide-container');
        if (!container || total === 0) return;
        var cw = container.clientWidth;
        var ch = container.clientHeight;
        var active = slides[current];
        if (!active) return;
        var sw = active.offsetWidth;
        var sh = active.offsetHeight;
        var scale = Math.min(cw / sw, ch / sh, 1) * 0.92;
        active.style.transform = 'scale(' + scale + ')';
    }

    function showSlide(n) {
        if (n < 0) n = 0;
        if (n >= total) n = total - 1;
        for (var i = 0; i < total; i++) {
            slides[i].classList.remove('active');
            slides[i].style.transform = '';
            if (thumbs[i]) thumbs[i].classList.remove('active');
        }
        slides[n].classList.add('active');
        if (thumbs[n]) thumbs[n].classList.add('active');
        current = n;
        updateUI();
        scaleSlides();
    }

    function next() { if (current < total - 1) showSlide(current + 1); }
    function prev() { if (current > 0) showSlide(current - 1); }

    function updateUI() {
        counterEl.textContent = (current + 1) + '/' + total;
        progressBar.style.width = ((current + 1) / total * 100) + '%';
        prevBtn.disabled = current === 0;
        nextBtn.disabled = current === total - 1;
        // Notes
        var noteEl = slides[current].getAttribute('data-notes') || '';
        if (notesPanel) {
            notesPanel.innerHTML = noteEl;
            notesPanel.classList.toggle('visible', !!noteEl);
        }
        // Scroll thumb into view
        if (thumbs[current]) {
            thumbs[current].scrollIntoView({block: 'nearest'});
        }
    }

    function onKey(e) {
        switch(e.key) {
            case 'ArrowRight': case ' ': case 'PageDown': e.preventDefault(); next(); break;
            case 'ArrowLeft': case 'PageUp': e.preventDefault(); prev(); break;
            case 'Home': e.preventDefault(); showSlide(0); break;
            case 'End': e.preventDefault(); showSlide(total - 1); break;
            case 'f': case 'F':
                if (!document.fullscreenElement) { document.documentElement.requestFullscreen().catch(function(){}); }
                else { document.exitFullscreen().catch(function(){}); }
                break;
        }
    }

    window.gotoSlide = function(n) { showSlide(n); };
    document.addEventListener('DOMContentLoaded', init);
})();
"""


def _emu_to_pt(emu: int) -> float:
    """Convert EMU to points (1 pt = 12700 EMU)."""
    return emu / 12700.0


def _extract_slide_html(
    slide,
    embed_images: bool,
    image_data_map: dict[str, str] | None = None,
) -> str:
    """Build the inner HTML for one slide.

    Iterates over shapes and produces positioned ``<div>`` elements with
    matching styles.
    """
    elements: list[str] = []
    # Standard slide dimensions in EMU (16:9)
    slide_width = 12192000   # 960 pt
    slide_height = 6858000   # 540 pt

    try:
        slide_width = slide.slide_width
        slide_height = slide.slide_height
    except Exception:
        pass

    for shape in slide.shapes:
        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height

        # Convert EMU to points, then to a percentage of the slide
        left_pct = (left / slide_width) * 100 if slide_width else 0
        top_pct = (top / slide_height) * 100 if slide_height else 0
        width_pct = (width / slide_width) * 100 if slide_width else 0
        height_pct = (height / slide_height) * 100 if slide_height else 0

        style = (
            f"position:absolute; left:{left_pct:.2f}%; top:{top_pct:.2f}%; "
            f"width:{width_pct:.2f}%; height:{height_pct:.2f}%; "
        )

        # Apply font styling from text frames
        inner = ""

        if shape.has_text_frame:
            inner = _text_frame_to_html(shape.text_frame, style)
            # Merge extra positioning style
            if inner.startswith("<div"):
                # Inject the position style into the first style attribute
                inner = inner.replace("style=\"", f"style=\"{style}", 1)
            else:
                inner = f'<div class="element" style="{style}">{inner}</div>'

        elif shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
            inner = _picture_to_html(shape, style, embed_images, image_data_map)

        elif shape.has_table:
            inner = _table_to_html(shape.table, style)

        else:
            # Generic shape: try to render text content
            try:
                if shape.has_text_frame:
                    inner = f'<div class="element" style="{style}">{_text_frame_to_html(shape.text_frame, "")}</div>'
                else:
                    inner = f'<div class="element" style="{style}"></div>'
            except Exception:
                inner = f'<div class="element" style="{style}"></div>'

        elements.append(inner)

    return "\n".join(elements)


def _text_frame_to_html(text_frame, extra_style: str = "") -> str:
    """Convert a python-pptx TextFrame to HTML."""
    paragraphs: list[str] = []
    for para in text_frame.paragraphs:
        runs_html: list[str] = []
        for run in para.runs:
            text = _escape_html(run.text)
            style_parts: list[str] = []
            if run.font.size:
                size_pt = _emu_to_pt(run.font.size)
                style_parts.append(f"font-size:{size_pt:.1f}pt")
            if run.font.bold:
                style_parts.append("font-weight:bold")
            if run.font.italic:
                style_parts.append("font-style:italic")
            if run.font.color:
                try:
                    rgb = run.font.color.rgb
                    if rgb:
                        style_parts.append(f"color:#{rgb}")
                except AttributeError:
                    pass
            if run.font.underline:
                style_parts.append("text-decoration:underline")
            style_attr = f' style="{";".join(style_parts)}"' if style_parts else ""
            runs_html.append(f"<span{style_attr}>{text}</span>")

        if not runs_html:
            # Empty paragraph
            paragraphs.append("<p>&nbsp;</p>")
            continue

        para_style = ""
        if para.alignment is not None:
            align_map = {0: "left", 1: "center", 2: "right", 3: "justify"}
            align = align_map.get(para.alignment, None)
            if align:
                para_style = f' style="text-align:{align}"'
        paragraphs.append(f"<p{para_style}>{''.join(runs_html)}</p>")

    return "".join(paragraphs) if paragraphs else ""


def _picture_to_html(
    shape,
    style: str,
    embed_images: bool,
    image_data_map: dict[str, str] | None = None,
) -> str:
    """Convert a picture shape to an HTML ``<img>`` tag."""
    try:
        image = shape.image
        content_type = image.content_type
        blob = image.blob
    except Exception:
        return f'<div class="element" style="{style}"></div>'

    if embed_images:
        b64 = base64.b64encode(blob).decode("ascii")
        src = f"data:{content_type};base64,{b64}"
    else:
        # Use a placeholder; the caller should extract images separately
        src = "data:image/png;base64,"

    return f'<div class="element" style="{style}"><img src="{src}" alt="image" style="width:100%;height:100%;object-fit:contain;"></div>'


def _table_to_html(table, style: str) -> str:
    """Convert a python-pptx Table to an HTML ``<table>``."""
    rows_html: list[str] = []
    for row_idx, row in enumerate(table.rows):
        cells_html: list[str] = []
        for cell in row.cells:
            cell_text = cell.text_frame.text if cell.has_text_frame else ""
            tag = "th" if row_idx == 0 else "td"
            cells_html.append(f"<{tag}>{_escape_html(cell_text)}</{tag}>")
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

    table_html = "<table>" + "\n".join(rows_html) + "</table>"
    return f'<div class="element" style="{style}">{table_html}</div>'


def _escape_html(text: str) -> str:
    """Escape special HTML characters."""
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _get_notes_text(slide) -> str:
    """Extract speaker notes from a slide as HTML."""
    try:
        notes_slide = slide.notes_slide
        if notes_slide and notes_slide.notes_text_frame:
            return _escape_html(notes_slide.notes_text_frame.text)
    except Exception:
        pass
    return ""


def export_to_html(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
    output_path: str | os.PathLike[str],
    *,
    theme: Literal["light", "dark"] = "light",
    include_notes: bool = False,
    embed_images: bool = True,
) -> str:
    """Export a PPTX presentation as a self-contained HTML slideshow.

    Generates an HTML file with embedded CSS and JavaScript that provides
    keyboard / click / swipe navigation, a progress bar, slide counter, and
    optional fullscreen mode.

    Parameters
    ----------
    prs_or_path :
        A :class:`~pptx.Presentation` instance or a path to a ``.pptx`` file.
    output_path :
        Destination path for the HTML file.
    theme :
        CSS colour theme: ``"light"`` (default) or ``"dark"``.
    include_notes :
        If ``True``, speaker notes are shown below each slide.
    embed_images :
        If ``True`` (default), images are base64-encoded directly into the
        HTML, producing a single self-contained file.  If ``False``, image
        sources are left as placeholders.
    include_notes :
        If ``True``, speaker notes are shown below each slide.

    Returns
    -------
    str
        Absolute path of the written HTML file.

    Raises
    ------
    FileNotFoundError
        If *prs_or_path* is a path that does not exist.
    ValueError
        If the presentation has no slides.
    """
    prs, _ = _resolve_presentation(prs_or_path)
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    slide_count = len(prs.slides)
    if slide_count == 0:
        raise ValueError("Cannot export an empty presentation to HTML")

    # Build slides HTML
    slides_html_parts: list[str] = []
    thumb_html_parts: list[str] = []

    for i, slide in enumerate(prs.slides):
        slide_inner = _extract_slide_html(slide, embed_images)
        notes_text = _get_notes_text(slide) if include_notes else ""
        notes_attr = f' data-notes="{notes_text}"' if notes_text else ""
        slides_html_parts.append(
            f'<div class="slide" data-index="{i}"{notes_attr}>\n{slide_inner}\n</div>'
        )
        thumb_html_parts.append(
            f'<div class="thumb{" active" if i == 0 else ""}" onclick="gotoSlide({i})">{i + 1}</div>'
        )

    # Assemble the full HTML document
    theme_css = _LIGHT_CSS if theme == "light" else _DARK_CSS
    slides_html = "\n".join(slides_html_parts)
    thumbs_html = "\n".join(thumb_html_parts)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Presentation</title>
<style>
{theme_css}
{_BASE_CSS}
</style>
</head>
<body>
<div class="slide-container">
{slides_html}
</div>

<div class="progress-bar" id="progress-bar" style="width:0%"></div>

<div class="notes-panel" id="notes-panel"></div>

<div class="slide-nav">
{thumbs_html}
</div>

<div class="navbar">
    <button class="nav-btn" id="btn-prev" onclick="gotoSlide(current-1)">&#9664; Prev</button>
    <span class="slide-counter" id="slide-counter">1/{slide_count}</span>
    <button class="nav-btn" id="btn-next" onclick="gotoSlide(current+1)">Next &#9654;</button>
</div>

<script>
{_JS}
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    return output_path


# ---------------------------------------------------------------------------
# 4. Thumbnail export
# ---------------------------------------------------------------------------

def export_thumbnails(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
    output_dir: str | os.PathLike[str],
    *,
    size: tuple[int, int] = (320, 180),
    format: Literal["PNG", "JPEG", "BMP"] = "PNG",
) -> list[str]:
    """Generate small thumbnail images for each slide.

    Renders slides via the ``preview_renderer`` infrastructure and then
    resizes each image to *size*.

    Parameters
    ----------
    prs_or_path :
        A :class:`~pptx.Presentation` instance or a path to a ``.pptx`` file.
    output_dir :
        Directory for the output thumbnail files.  Created if needed.
    size :
        ``(width, height)`` in pixels for each thumbnail (default 320x180).
    format :
        Image format: ``"PNG"``, ``"JPEG"``, or ``"BMP"``.

    Returns
    -------
    list[str]
        Absolute paths of the written thumbnail files, in slide order.

    Raises
    ------
    FileNotFoundError
        If *prs_or_path* is a path that does not exist.
    RuntimeError
        If no rendering backend is available.
    """
    prs, _ = _resolve_presentation(prs_or_path)
    output_dir = _ensure_dir(output_dir)

    if len(prs.slides) == 0:
        return []

    # We need a file path for the renderer
    tmp_dir: str | None = None
    if isinstance(prs_or_path, (str, os.PathLike)):
        pptx_path = os.fspath(prs_or_path)
    else:
        pptx_path, tmp_dir = _save_to_temp(prs)

    try:
        render_dir = tempfile.mkdtemp(prefix="pptx_skill_thumb_")
        render_result = render_preview(pptx_path, render_dir, dpi=150)
        if not render_result.slide_pngs:
            raise RuntimeError(
                "No rendering backend available for thumbnail export. "
                "Install LibreOffice, PyMuPDF, or use Windows COM."
            )

        from PIL import Image as PILImage

        ext = format.lower()
        if ext == "jpeg":
            ext = "jpg"

        output_paths: list[str] = []
        for i, src in enumerate(render_result.slide_pngs):
            dst = os.path.join(output_dir, f"thumb_{i + 1:03d}.{ext}")
            img: PILImage.Image = PILImage.open(src)
            with img:
                # Use high-quality downsampling
                img = img.resize(size, PILImage.Resampling.LANCZOS)
                if format == "JPEG" and img.mode == "RGBA":
                    bg = PILImage.new("RGB", img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[3])
                    img = bg
                elif format in ("JPEG", "BMP") and img.mode not in ("RGB", "L"):
                    img = img.convert("RGB")
                img.save(dst, format=format)
            output_paths.append(dst)

        shutil.rmtree(render_dir, ignore_errors=True)
        return output_paths

    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# 5. Text export
# ---------------------------------------------------------------------------

def export_to_text(
    prs_or_path: str | os.PathLike[str] | _PresentationCls,
    output_path: str | os.PathLike[str],
    *,
    include_notes: bool = True,
    include_tables: bool = True,
) -> str:
    """Export all text content from a PPTX as plain text / Markdown.

    Each slide becomes a section with a header.  Tables are rendered as
    Markdown tables (when *include_tables* is ``True``).  Speaker notes are
    included when *include_notes* is ``True``.

    Parameters
    ----------
    prs_or_path :
        A :class:`~pptx.Presentation` instance or a path to a ``.pptx`` file.
    output_path :
        Destination path for the text file.
    include_notes :
        If ``True`` (default), include speaker notes under each slide.
    include_tables :
        If ``True`` (default), include table content as Markdown tables.

    Returns
    -------
    str
        Absolute path of the written text file.

    Raises
    ------
    FileNotFoundError
        If *prs_or_path* is a path that does not exist.
    """
    prs, _ = _resolve_presentation(prs_or_path)
    output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    sections: list[str] = []

    for i, slide in enumerate(prs.slides):
        lines: list[str] = []
        lines.append(f"--- Slide {i + 1} ---\n")

        for shape in slide.shapes:
            # Text content
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        # Detect heading-like paragraphs (large font, bold, or first text)
                        is_heading = False
                        try:
                            for run in para.runs:
                                if run.font.bold or (run.font.size and run.font.size >= 200000):
                                    is_heading = True
                                    break
                        except Exception:
                            pass
                        if is_heading and not lines[1:]:
                            lines.append(f"# {text}\n")
                        else:
                            lines.append(text)

            # Table content
            if include_tables and shape.has_table:
                table = shape.table
                table_md = _table_to_markdown(table)
                if table_md:
                    lines.append("\n" + table_md)

        # Speaker notes
        if include_notes:
            try:
                notes_slide = slide.notes_slide
                if notes_slide and notes_slide.notes_text_frame:
                    notes_text = notes_slide.notes_text_frame.text.strip()
                    if notes_text:
                        lines.append(f"\n[Notes] {notes_text}")
            except Exception:
                pass

        sections.append("\n".join(lines))

    full_text = "\n\n".join(sections) + "\n"

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(full_text)

    return output_path


def _table_to_markdown(table) -> str:
    """Convert a python-pptx Table to a Markdown table string."""
    if not table.rows:
        return ""

    rows_md: list[str] = []
    for row_idx, row in enumerate(table.rows):
        cells: list[str] = []
        for cell in row.cells:
            cell_text = cell.text_frame.text.strip() if cell.has_text_frame else ""
            # Escape pipe characters
            cell_text = cell_text.replace("|", "\\|")
            cells.append(cell_text)
        row_str = "| " + " | ".join(cells) + " |"
        rows_md.append(row_str)
        # Add separator after header row
        if row_idx == 0:
            sep = "|" + "|".join([" --- " for _ in cells]) + "|"
            rows_md.append(sep)

    return "\n".join(rows_md)
