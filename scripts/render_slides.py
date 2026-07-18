"""
render_slides.py - PPTX slides to PNG for visual review.

PR0 fixes:
- LibreOffice uses a unique UserInstallation and temp directory per run.
- Only PDFs produced after the conversion call are accepted (stale-PDF guard).
- ``auto`` mode reports per-engine errors instead of a generic message.
- PowerPoint COM export honours the requested DPI.
"""
from __future__ import annotations

import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any


def find_soffice():
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


def find_powerpoint():
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _fresh_pdfs(directory: str, since: float) -> list[str]:
    """Return PDFs in directory whose mtime is >= since."""
    paths = glob.glob(os.path.join(directory, "*.pdf"))
    return [p for p in paths if os.path.getmtime(p) >= since]


def render_with_libreoffice(pptx_path, output_dir, dpi=150, soffice_path=None):
    soffice_path = soffice_path or find_soffice()
    if not soffice_path:
        return None, {"engine": "libreoffice", "error": "LibreOffice not found"}

    os.makedirs(output_dir, exist_ok=True)
    abs_pptx = os.path.abspath(pptx_path)
    abs_out = os.path.abspath(output_dir)

    # Unique temporary directories isolate both the user profile and the PDF output.
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
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    except subprocess.TimeoutExpired as exc:
        _cleanup_libreoffice_tmp(tmp_root)
        return None, {"engine": "libreoffice", "error": f"timed out after {exc.timeout}s"}
    except Exception as exc:
        _cleanup_libreoffice_tmp(tmp_root)
        return None, {"engine": "libreoffice", "error": f"subprocess error: {exc}"}

    if result.returncode != 0:
        _cleanup_libreoffice_tmp(tmp_root)
        return None, {
            "engine": "libreoffice",
            "error": f"soffice exited {result.returncode}",
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    pdfs = _fresh_pdfs(pdf_dir, started)
    if not pdfs:
        _cleanup_libreoffice_tmp(tmp_root)
        return None, {
            "engine": "libreoffice",
            "error": "no PDF produced after conversion",
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    pdf_path = pdfs[0]
    pngs = _pdf_to_pngs(pdf_path, abs_out, dpi)
    _cleanup_libreoffice_tmp(tmp_root)

    if pngs:
        return _rename_pngs(pngs, abs_out), None
    return None, {"engine": "libreoffice", "error": "PDF produced but PNG conversion failed"}


def _cleanup_libreoffice_tmp(tmp_root: str) -> None:
    try:
        shutil.rmtree(tmp_root, ignore_errors=True)
    except OSError as exc:
        print(f"  [警告] LibreOffice 临时目录清理失败 {tmp_root}: {exc}", file=sys.stderr)


def _rename_pngs(pngs: list[str], output_dir: str) -> list[str]:
    renamed = []
    for i, p in enumerate(sorted(pngs)):
        new_name = os.path.join(output_dir, f"slide_{i+1:03d}.png")
        if p != new_name:
            if os.path.exists(new_name):
                os.remove(new_name)
            os.rename(p, new_name)
        renamed.append(new_name)
    return renamed


def _pdf_to_pngs(pdf_path, output_dir, dpi=150):
    pngs = []
    errors = []
    try:
        import fitz
        doc = fitz.open(pdf_path)
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            out = os.path.join(output_dir, f"slide_{page.number+1:03d}.png")
            pix.save(out)
            pngs.append(out)
        doc.close()
        return pngs
    except ImportError:
        errors.append("PyMuPDF (fitz) not installed")
    except Exception as exc:
        errors.append(f"PyMuPDF error: {exc}")

    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path, dpi=dpi)
        for i, img in enumerate(images):
            out = os.path.join(output_dir, f"slide_{i+1:03d}.png")
            img.save(out, "PNG")
            pngs.append(out)
        return pngs
    except ImportError:
        errors.append("pdf2image not installed (also needs poppler on system)")
    except Exception as exc:
        errors.append(f"pdf2image error: {exc}")

    print(f"  [诊断] PDF→PNG 渲染失败原因: {'; '.join(errors)}", file=sys.stderr)
    return pngs


def render_with_com(pptx_path, output_dir, dpi=150):
    try:
        import win32com.client
    except ImportError:
        return None, {"engine": "com", "error": "win32com not available"}

    pp = None
    pres = None
    try:
        pp = win32com.client.Dispatch("PowerPoint.Application")
        pp.Visible = 0
        abs_pptx = os.path.abspath(pptx_path)
        abs_out = os.path.abspath(output_dir)
        os.makedirs(abs_out, exist_ok=True)
        pres = pp.Presentations.Open(abs_pptx, WithWindow=False)

        # Honour DPI by computing export pixel size from slide dimensions in points.
        page_setup = pres.PageSetup
        width_pt = float(page_setup.SlideSize.Width)
        height_pt = float(page_setup.SlideSize.Height)
        scale_width = round(width_pt / 72 * dpi)
        scale_height = round(height_pt / 72 * dpi)
        pres.Export(abs_out, "PNG", scale_width, scale_height)

        pngs = sorted(glob.glob(os.path.join(abs_out, "*.png")))
        if pngs:
            return _rename_pngs(pngs, abs_out), None
        return None, {"engine": "com", "error": "No PNG exported"}
    except Exception as exc:
        return None, {"engine": "com", "error": f"COM error: {exc}"}
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


def render_slides(pptx_path, output_dir="./slides_preview", dpi=150, engine="auto"):
    if not os.path.exists(pptx_path):
        return None, {"error": f"File not found: {pptx_path}"}

    attempts: list[dict[str, Any]] = []

    if engine in ("auto", "libreoffice"):
        soffice = find_soffice()
        if soffice:
            result, err = render_with_libreoffice(pptx_path, output_dir, dpi, soffice)
            attempts.append({"engine": "libreoffice", "success": result is not None, "error": err})
            if result:
                return result, None
        else:
            attempts.append({"engine": "libreoffice", "success": False, "error": "not found"})

    if engine in ("auto", "com"):
        ppt = find_powerpoint()
        if ppt:
            result, err = render_with_com(pptx_path, output_dir, dpi)
            attempts.append({"engine": "com", "success": result is not None, "error": err})
            if result:
                return result, None
        else:
            attempts.append({"engine": "com", "success": False, "error": "not found"})

    if engine == "auto":
        messages = [
            f"{a['engine']}: {a['error'].get('error', a['error']) if isinstance(a['error'], dict) else a['error']}"
            for a in attempts
            if not a["success"]
        ]
        return None, {"error": "No rendering engine succeeded", "attempts": attempts, "details": "; ".join(messages)}

    return None, {"error": f"Unknown engine: {engine}"}


def main():
    parser = argparse.ArgumentParser(description="Render PPTX slides to PNG")
    parser.add_argument("pptx_path")
    parser.add_argument("-o", "--output", default="./slides_preview")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--engine", default="auto", choices=["auto", "libreoffice", "com"])
    args = parser.parse_args()
    print(f"Rendering: {args.pptx_path}")
    pngs, err = render_slides(args.pptx_path, args.output, args.dpi, args.engine)
    if err:
        print(f"ERROR: {err.get('error', err) if isinstance(err, dict) else err}", file=sys.stderr)
        sys.exit(1)
    print(f"Rendered {len(pngs)} slides:")
    for p in pngs:
        print(f"  {p}")


if __name__ == "__main__":
    main()
