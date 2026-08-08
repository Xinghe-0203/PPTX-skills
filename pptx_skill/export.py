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

.loading {
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: 40px;
    height: 40px;
    border: 3px solid var(--border);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
    z-index: 9999;
}
@keyframes spin { to { transform: translate(-50%, -50%) rotate(360deg); } }

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
    transition: opacity 0.3s ease, transform 0.3s ease;
}
.slide.active { display: block; opacity: 1; }
.slide.leaving { opacity: 0; }

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
    transition: background 0.15s, transform 0.1s;
    outline: none;
}
.navbar .nav-btn:hover { background: var(--border); }
.navbar .nav-btn:focus-visible { box-shadow: 0 0 0 2px var(--accent); }
.navbar .nav-btn:active { transform: scale(0.96); }
.navbar .nav-btn:disabled { opacity: 0.3; cursor: default; }
.navbar .slide-counter { font-size: 14px; color: var(--text-muted); min-width: 80px; text-align: center; font-variant-numeric: tabular-nums; }

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
    transition: max-height 0.3s ease;
}
.notes-panel.visible { display: block; }
.notes-panel::-webkit-scrollbar { width: 6px; }
.notes-panel::-webkit-scrollbar-track { background: transparent; }
.notes-panel::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

.slide-nav {
    position: fixed;
    bottom: 56px;
    right: 12px;
    display: flex;
    flex-direction: column;
    gap: 4px;
    z-index: 10;
    max-height: 70vh;
    overflow-y: auto;
    padding: 4px;
}
.slide-nav::-webkit-scrollbar { width: 4px; }
.slide-nav::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
.slide-nav .thumb {
    width: 80px;
    height: 45px;
    border: 2px solid transparent;
    border-radius: 4px;
    cursor: pointer;
    opacity: 0.5;
    transition: opacity 0.2s, border-color 0.2s, transform 0.1s;
    background: var(--slide-bg);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    color: var(--text-muted);
}
.slide-nav .thumb:hover { opacity: 0.8; transform: scale(1.05); }
.slide-nav .thumb.active { border-color: var(--accent); opacity: 1; }
.slide-nav .thumb:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }

.play-indicator {
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    width: 80px;
    height: 80px;
    border-radius: 50%;
    background: rgba(0,0,0,0.6);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 100;
    animation: fadeInOut 1s ease;
}
.play-indicator.show { display: flex; }
.play-indicator svg { fill: white; width: 32px; height: 32px; }
@keyframes fadeInOut {
    0%, 100% { opacity: 0; }
    50% { opacity: 1; }
}

.slide-number-badge {
    position: fixed;
    top: 12px;
    right: 12px;
    background: var(--nav-bg);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 4px 12px;
    font-size: 13px;
    color: var(--text-muted);
    z-index: 20;
    font-variant-numeric: tabular-nums;
}

@media (max-width: 1100px) {
    .slide { width: 640px; height: 360px; }
}
@media (max-width: 720px) {
    .slide { width: 480px; height: 270px; }
    .slide-nav { display: none; }
    .navbar { padding: 0 10px; }
    .navbar .nav-btn { padding: 6px 10px; font-size: 12px; }
}

@media print {
    .navbar, .slide-nav, .progress-bar, .notes-panel, .slide-number-badge, .play-indicator, .loading {
        display: none !important;
    }
    .slide-container { height: 100vh; }
    .slide {
        display: block !important;
        position: relative;
        width: 100%;
        height: 100vh;
        box-shadow: none;
        border: none;
        page-break-after: always;
        break-after: page;
    }
    .slide.active { display: block !important; }
    .slide:last-child { page-break-after: avoid; }
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
    var slideBadge = document.getElementById('slide-badge');
    var playIndicator = document.getElementById('play-indicator');
    var slideContainer = document.querySelector('.slide-container');

    // Auto-play state
    var isPlaying = false;
    var playInterval = null;
    var playIntervalMs = 5000;

    // Debounce for resize
    var resizeTimer = null;

    function init() {
        slides = document.querySelectorAll('.slide');
        thumbs = document.querySelectorAll('.slide-nav .thumb');
        total = slides.length;
        if (total === 0) return;

        // Remove loading indicator
        var loading = document.getElementById('loading');
        if (loading) loading.remove();

        showSlide(0);

        // Keyboard navigation
        window.addEventListener('keydown', onKey);

        // Touch support with swipe detection
        var touchStartX = 0;
        var touchStartY = 0;
        var touchStartTime = 0;
        var isSwiping = false;

        document.addEventListener('touchstart', function(e) {
            touchStartX = e.changedTouches[0].screenX;
            touchStartY = e.changedTouches[0].screenY;
            touchStartTime = Date.now();
            isSwiping = false;
        }, {passive: true});

        document.addEventListener('touchmove', function(e) {
            var dx = e.changedTouches[0].screenX - touchStartX;
            var dy = e.changedTouches[0].screenY - touchStartY;
            if (Math.abs(dx) > 10 && Math.abs(dx) > Math.abs(dy)) {
                isSwiping = true;
            }
        }, {passive: true});

        document.addEventListener('touchend', function(e) {
            if (isSwiping) {
                var dx = e.changedTouches[0].screenX - touchStartX;
                var dt = Date.now() - touchStartTime;
                // Require minimum velocity: 50px and within reasonable time
                if (Math.abs(dx) > 50 && dt < 500) {
                    dx > 0 ? prev() : next();
                }
            }
        }, {passive: true});

        // Click navigation with swipe detection
        slideContainer.addEventListener('click', function(e) {
            if (isSwiping) return;
            var rect = this.getBoundingClientRect();
            if (e.clientX < rect.left + rect.width / 2) { prev(); } else { next(); }
        });

        // Debounced resize
        scaleSlides();
        window.addEventListener('resize', function() {
            clearTimeout(resizeTimer);
            resizeTimer = setTimeout(scaleSlides, 100);
        });

        // Click handlers for buttons (already in HTML but ensure they work)
        prevBtn.addEventListener('click', function() { prev(); });
        nextBtn.addEventListener('click', function() { next(); });

        // Event delegation for thumbs (click + keyboard)
        var nav = document.querySelector('.slide-nav');
        if (nav) {
            nav.addEventListener('click', function(e) {
                var thumb = e.target.closest('.thumb');
                if (thumb) {
                    var idx = parseInt(thumb.getAttribute('data-slide'), 10);
                    if (!isNaN(idx)) showSlide(idx);
                }
            });
            nav.addEventListener('keydown', function(e) {
                var thumb = e.target.closest('.thumb');
                if (thumb && (e.key === 'Enter' || e.key === ' ')) {
                    e.preventDefault();
                    var idx = parseInt(thumb.getAttribute('data-slide'), 10);
                    if (!isNaN(idx)) showSlide(idx);
                }
            });
        }
    }

    function scaleSlides() {
        var container = slideContainer;
        if (!container || total === 0) return;
        var cw = container.clientWidth;
        var ch = container.clientHeight;
        var active = slides[current];
        if (!active) return;
        var sw = active.offsetWidth;
        var sh = active.offsetHeight;
        var scale = Math.min(cw / sw, ch / sh, 1) * 0.92;
        active.style.transform = 'scale(' + scale + ')';
        active.style.transformOrigin = 'center center';
    }

    function showSlide(n) {
        if (n < 0) n = 0;
        if (n >= total) n = total - 1;

        var prevIndex = current;

        // Handle leaving slide with transition
        if (slides[prevIndex] && prevIndex !== n) {
            slides[prevIndex].classList.add('leaving');
            setTimeout(function() {
                slides[prevIndex].classList.remove('active', 'leaving');
                slides[prevIndex].style.transform = '';
            }, 300);
        } else {
            slides[prevIndex].classList.remove('active');
            slides[prevIndex].style.transform = '';
        }

        slides[n].classList.add('active');
        if (thumbs[n]) {
            thumbs[n].classList.add('active');
            // Smooth scroll thumb into view
            thumbs[n].scrollIntoView({block: 'nearest', behavior: 'smooth'});
        }

        current = n;
        updateUI();
        scaleSlides();
    }

    function next() { if (current < total - 1) showSlide(current + 1); else if (isPlaying) stopPlay(); }
    function prev() { if (current > 0) showSlide(current - 1); }

    function updateUI() {
        counterEl.textContent = (current + 1) + '/' + total;
        progressBar.style.width = ((current + 1) / total * 100) + '%';
        if (slideBadge) slideBadge.textContent = (current + 1) + ' / ' + total;
        prevBtn.disabled = current === 0;
        nextBtn.disabled = current === total - 1;

        // Update ARIA for active slide
        for (var i = 0; i < total; i++) {
            slides[i].setAttribute('aria-hidden', i !== current ? 'true' : 'false');
        }

        // Notes
        var noteEl = slides[current].getAttribute('data-notes') || '';
        if (notesPanel) {
            notesPanel.textContent = noteEl;
            notesPanel.classList.toggle('visible', !!noteEl);
        }
    }

    function onKey(e) {
        switch(e.key) {
            case 'ArrowRight':
            case 'PageDown':
                e.preventDefault();
                next();
                break;
            case 'ArrowLeft':
            case 'PageUp':
                e.preventDefault();
                prev();
                break;
            case ' ':
                e.preventDefault();
                if (isPlaying) stopPlay(); else startPlay();
                break;
            case 'Home':
                e.preventDefault();
                showSlide(0);
                break;
            case 'End':
                e.preventDefault();
                showSlide(total - 1);
                break;
            case 'f':
            case 'F':
                if (!document.fullscreenElement) {
                    document.documentElement.requestFullscreen().catch(function(){});
                } else {
                    document.exitFullscreen().catch(function(){});
                }
                break;
            case 'p':
            case 'P':
                e.preventDefault();
                if (isPlaying) stopPlay(); else startPlay();
                break;
            case 'n':
            case 'N':
                e.preventDefault();
                var noteVisible = notesPanel && notesPanel.classList.contains('visible');
                if (noteVisible) {
                    notesPanel.classList.remove('visible');
                } else {
                    var noteText = slides[current].getAttribute('data-notes') || '';
                    if (noteText) {
                        notesPanel.textContent = noteText;
                        notesPanel.classList.add('visible');
                    }
                }
                break;
            case 't':
            case 'T':
                e.preventDefault();
                toggleThumbnails();
                break;
        }
    }

    function startPlay() {
        if (total <= 1) return;
        isPlaying = true;
        showPlayIndicator('play');
        playInterval = setInterval(function() {
            if (current < total - 1) {
                showSlide(current + 1);
            } else {
                stopPlay();
            }
        }, playIntervalMs);
    }

    function stopPlay() {
        isPlaying = false;
        showPlayIndicator('pause');
        if (playInterval) {
            clearInterval(playInterval);
            playInterval = null;
        }
    }

    function showPlayIndicator(type) {
        if (!playIndicator) return;
        playIndicator.classList.remove('show');
        // Force reflow to restart animation
        void playIndicator.offsetWidth;
        playIndicator.classList.add('show');
        var svg = playIndicator.querySelector('svg');
        if (svg) {
            svg.innerHTML = type === 'play'
                ? '<polygon points="8,4 20,12 8,20"/>'
                : '<rect x="6" y="4" width="4" height="16"/><rect x="14" y="4" width="4" height="16"/>';
        }
        setTimeout(function() { playIndicator.classList.remove('show'); }, 1000);
    }

    function toggleThumbnails() {
        var nav = document.querySelector('.slide-nav');
        if (nav) {
            nav.style.display = nav.style.display === 'none' ? 'flex' : 'none';
        }
    }

    // Public API
    window.gotoSlide = function(n) { showSlide(n); };
    window.startPresentation = function() { startPlay(); };
    window.stopPresentation = function() { stopPlay(); };
    window.togglePresentation = function() { isPlaying ? stopPlay() : startPlay(); };

    // Init
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
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
        .replace("'", "&#39;")
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
            f'<div class="slide" data-index="{i}" role="region" aria-label="Slide {i + 1} of {slide_count}" aria-hidden="true"{notes_attr}>\n{slide_inner}\n</div>'
        )
        thumb_html_parts.append(
            f'<div class="thumb{" active" if i == 0 else ""}" data-slide="{i}" role="button" aria-label="Go to slide {i + 1}" tabindex="0">{i + 1}</div>'
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
<div class="loading" id="loading"></div>

<div class="slide-container" role="main" aria-label="Presentation slides">
{slides_html}
</div>

<div class="progress-bar" id="progress-bar" style="width:0%"></div>

<div class="notes-panel" id="notes-panel" role="note" aria-label="Speaker notes"></div>

<div class="slide-nav" aria-label="Slide navigation">
{thumbs_html}
</div>

<div class="slide-number-badge" id="slide-badge" aria-live="polite">1 / {slide_count}</div>

<div class="play-indicator" id="play-indicator" aria-hidden="true">
    <svg viewBox="0 0 24 24"><polygon points="8,4 20,12 8,20"/></svg>
</div>

<div class="navbar" role="toolbar" aria-label="Presentation controls">
    <button class="nav-btn" id="btn-prev" aria-label="Previous slide">&#9664; Prev</button>
    <span class="slide-counter" id="slide-counter" aria-live="polite">1/{slide_count}</span>
    <button class="nav-btn" id="btn-next" aria-label="Next slide">Next &#9654;</button>
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
