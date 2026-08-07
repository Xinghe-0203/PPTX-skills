"""Preview rendering adapter: PPTX -> PNG via LibreOffice or PowerPoint COM.

This is the authoritative renderer implementation used by ``scripts/render_slides.py``
and the V2 generation pipeline. It returns a structured ``PreviewRenderResult``
with environment metadata and per-engine attempt records.
"""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from typing import Any, Literal

from pptx_skill.pptx_renderer import PreviewRenderResult


@dataclass
class _EngineAttempt:
    engine: str
    success: bool
    error: dict[str, Any] | str | None = None


def find_soffice() -> str | None:
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    p = shutil.which("soffice")
    if p:
        candidates.insert(0, p)
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def find_powerpoint() -> str | None:
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _libreoffice_version(soffice_path: str) -> str:
    try:
        result = subprocess.run(
            [soffice_path, "--version"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        text = (result.stdout + result.stderr).strip()
        if result.returncode == 0 and text:
            return text
    except Exception:
        pass
    return "unknown"


def _fresh_pdfs(directory: str, since: float) -> list[str]:
    paths = glob.glob(os.path.join(directory, "*.pdf"))
    return [p for p in paths if os.path.getmtime(p) >= since]


def _cleanup_tmp(tmp_root: str) -> None:
    try:
        shutil.rmtree(tmp_root, ignore_errors=True)
    except OSError:
        pass


def _rename_pngs(pngs: list[str], output_dir: str) -> list[str]:
    renamed = []
    for i, p in enumerate(sorted(pngs)):
        new_name = os.path.join(output_dir, f"slide_{i+1:03d}.png")
        if p != new_name:
            if os.path.exists(new_name):
                os.remove(new_name)
            shutil.move(p, new_name)
        renamed.append(new_name)
    return renamed


def _pdf_to_pngs(pdf_path: str, output_dir: str, dpi: int) -> tuple[list[str], list[str]]:
    pngs: list[str] = []
    errors: list[str] = []
    try:
        import fitz
        doc = fitz.open(pdf_path)
        for page in doc:
            out = os.path.join(output_dir, f"slide_{page.number+1:03d}.png")
            pix = page.get_pixmap(dpi=dpi)
            pix.save(out)
            pngs.append(out)
        doc.close()
        return pngs, errors
    except ImportError:
        errors.append("PyMuPDF (fitz) not installed")
    except Exception as exc:
        errors.append(f"PyMuPDF error: {exc}")

    if pngs:
        return pngs, errors

    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path, dpi=dpi)
        for i, img in enumerate(images):
            out = os.path.join(output_dir, f"slide_{i+1:03d}.png")
            img.save(out, "PNG")
            pngs.append(out)
        return pngs, errors
    except ImportError:
        errors.append("pdf2image not installed (also needs poppler on system)")
    except Exception as exc:
        errors.append(f"pdf2image error: {exc}")

    return pngs, errors


def _pixel_sizes(pngs: list[str]) -> list[tuple[int, int]]:
    sizes: list[tuple[int, int]] = []
    try:
        from PIL import Image
        for p in pngs:
            with Image.open(p) as img:
                sizes.append(img.size)
    except Exception:
        pass
    return sizes


def render_with_libreoffice(
    pptx_path: str,
    output_dir: str,
    dpi: int = 150,
    soffice_path: str | None = None,
) -> PreviewRenderResult:
    soffice_path = soffice_path or find_soffice()
    result = PreviewRenderResult(
        renderer="libreoffice",
        renderer_version=_libreoffice_version(soffice_path or ""),
        target_dpi=dpi,
        environment={"soffice_path": soffice_path},
    )
    if not soffice_path:
        result.attempts.append({"engine": "libreoffice", "success": False, "error": "not found"})
        return result

    os.makedirs(output_dir, exist_ok=True)
    abs_pptx = os.path.abspath(pptx_path)

    tmp_root = tempfile.mkdtemp(prefix="pptx_skill_libreoffice_")
    pdf_dir = os.path.join(tmp_root, "pdf")
    user_install = os.path.join(tmp_root, "profile")
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(user_install, exist_ok=True)

    started = time.time()
    cmd = [
        soffice_path,
        f"-env:UserInstallation=file:///{user_install.replace(os.sep, '/')}",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", pdf_dir,
        abs_pptx,
    ]

    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    except subprocess.TimeoutExpired as exc:
        if exc.child is not None:  # type: ignore[attr-defined]
            try:
                exc.child.kill()  # type: ignore[attr-defined]
                exc.child.wait(timeout=10)  # type: ignore[attr-defined]
            except Exception:
                pass
        _cleanup_tmp(tmp_root)
        result.attempts.append({"engine": "libreoffice", "success": False, "error": f"timed out after {exc.timeout}s"})
        return result
    except Exception as exc:
        _cleanup_tmp(tmp_root)
        result.attempts.append({"engine": "libreoffice", "success": False, "error": f"subprocess error: {exc}"})
        return result

    if proc.returncode != 0:
        _cleanup_tmp(tmp_root)
        result.attempts.append(
            {
                "engine": "libreoffice",
                "success": False,
                "error": f"soffice exited {proc.returncode}",
                "stdout": proc.stdout,
                "stderr": proc.stderr,
            }
        )
        return result

    pdfs = _fresh_pdfs(pdf_dir, started)
    if not pdfs:
        _cleanup_tmp(tmp_root)
        result.attempts.append(
            {
                "engine": "libreoffice",
                "success": False,
                "error": "no PDF produced after conversion",
                "stdout": proc.stdout,
                "stderr": proc.stderr,
            }
        )
        return result

    pdf_path = pdfs[0]
    pngs, errors = _pdf_to_pngs(pdf_path, output_dir, dpi)
    _cleanup_tmp(tmp_root)

    if pngs:
        result.slide_pngs = _rename_pngs(pngs, output_dir)
        result.actual_pixel_sizes = _pixel_sizes(result.slide_pngs)
        result.attempts.append({"engine": "libreoffice", "success": True})
    else:
        result.attempts.append(
            {"engine": "libreoffice", "success": False, "error": "PDF produced but PNG conversion failed", "details": errors}
        )
    return result


def render_with_com(pptx_path: str, output_dir: str, dpi: int = 150) -> PreviewRenderResult:
    result = PreviewRenderResult(
        renderer="com",
        renderer_version="unknown",
        target_dpi=dpi,
        environment={"powerpoint_path": find_powerpoint()},
    )
    try:
        import win32com.client
    except ImportError:
        result.attempts.append({"engine": "com", "success": False, "error": "win32com not available"})
        return result

    pp = None
    pres = None
    try:
        pp = win32com.client.Dispatch("PowerPoint.Application")
        pp.Visible = 0
        abs_pptx = os.path.abspath(pptx_path)
        abs_out = os.path.abspath(output_dir)
        os.makedirs(abs_out, exist_ok=True)
        pres = pp.Presentations.Open(abs_pptx, WithWindow=False)
        page_setup = pres.PageSetup
        width_pt = float(page_setup.SlideSize.Width)
        height_pt = float(page_setup.SlideSize.Height)
        scale_width = round(width_pt / 72 * dpi)
        scale_height = round(height_pt / 72 * dpi)
        pres.Export(abs_out, "PNG", scale_width, scale_height)

        pngs = sorted(glob.glob(os.path.join(abs_out, "*.png")))
        if pngs:
            result.slide_pngs = _rename_pngs(pngs, abs_out)
            result.actual_pixel_sizes = _pixel_sizes(result.slide_pngs)
            result.attempts.append({"engine": "com", "success": True})
        else:
            result.attempts.append({"engine": "com", "success": False, "error": "No PNG exported"})
    except Exception as exc:
        result.attempts.append({"engine": "com", "success": False, "error": f"COM error: {exc}"})
    finally:
        try:
            if pres:
                pres.Close()
        except Exception:
            pass
        try:
            if pp:
                pp.Quit()
        except Exception:
            pass
    return result


def render_preview(
    pptx_path: str,
    output_dir: str = "./slides_preview",
    dpi: int = 150,
    engine: Literal["auto", "libreoffice", "com"] = "auto",
) -> PreviewRenderResult:
    """Render a PPTX to PNGs and return a structured result."""
    if not os.path.exists(pptx_path):
        return PreviewRenderResult(
            renderer="",
            renderer_version="",
            target_dpi=dpi,
            environment={},
            attempts=[{"engine": engine, "success": False, "error": f"File not found: {pptx_path}"}],
        )

    all_attempts: list[dict[str, Any]] = []

    if engine in ("auto", "libreoffice"):
        soffice = find_soffice()
        if soffice:
            result = render_with_libreoffice(pptx_path, output_dir, dpi, soffice)
            all_attempts.extend(result.attempts)
            if result.slide_pngs:
                result.attempts = all_attempts
                return result
        else:
            all_attempts.append({"engine": "libreoffice", "success": False, "error": "LibreOffice (soffice) not found on system"})

    if engine in ("auto", "com"):
        ppt = find_powerpoint()
        if ppt:
            result = render_with_com(pptx_path, output_dir, dpi)
            all_attempts.extend(result.attempts)
            if result.slide_pngs:
                result.attempts = all_attempts
                return result
        else:
            all_attempts.append({"engine": "com", "success": False, "error": "PowerPoint (POWERPNT.EXE) not found on system"})

    return PreviewRenderResult(
        renderer="",
        renderer_version="",
        target_dpi=dpi,
        environment={},
        attempts=all_attempts,
    )


