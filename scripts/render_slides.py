"""
render_slides.py - PPTX slides to PNG for visual review
"""
import argparse, os, sys, shutil, subprocess, glob

def find_soffice():
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    p = shutil.which("soffice")
    if p: candidates.insert(0, p)
    for c in candidates:
        if os.path.exists(c): return c
    return None

def find_powerpoint():
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    ]
    for c in candidates:
        if os.path.exists(c): return c
    return None

def render_with_libreoffice(pptx_path, output_dir, dpi=150, soffice_path=None):
    if not soffice_path: soffice_path = find_soffice()
    if not soffice_path: return None, "LibreOffice not found"
    os.makedirs(output_dir, exist_ok=True)
    abs_pptx = os.path.abspath(pptx_path)
    abs_out = os.path.abspath(output_dir)
    pdf_dir = os.path.join(abs_out, "_pdf_temp")
    os.makedirs(pdf_dir, exist_ok=True)
    cmd = [soffice_path, "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, abs_pptx]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        pdfs = glob.glob(os.path.join(pdf_dir, "*.pdf"))
        if not pdfs: return None, f"PDF conversion failed: {r.stderr}"
    except Exception as e:
        return None, f"PDF conversion error: {e}"
    pdf_path = pdfs[0]
    pngs = _pdf_to_pngs(pdf_path, abs_out, dpi)
    try:
        shutil.rmtree(pdf_dir, ignore_errors=True)
    except OSError as e:
        print(f"  [警告] 临时目录清理失败 {pdf_dir}: {e}", file=sys.stderr)
    if pngs:
        renamed = []
        for i, p in enumerate(sorted(pngs)):
            new_name = os.path.join(abs_out, f"slide_{i+1:03d}.png")
            if p != new_name:
                if os.path.exists(new_name): os.remove(new_name)
                os.rename(p, new_name)
            renamed.append(new_name)
        return renamed, None
    return None, "No PNG files generated"

def _pdf_to_pngs(pdf_path, output_dir, dpi=150):
    pngs = []
    errors = []
    try:
        import fitz
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=dpi)
            out = os.path.join(output_dir, f"slide_{i+1:03d}.png")
            pix.save(out); pngs.append(out)
        doc.close(); return pngs
    except ImportError:
        errors.append("PyMuPDF (fitz) not installed")
    except Exception as e:
        errors.append(f"PyMuPDF error: {e}")
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path, dpi=dpi)
        for i, img in enumerate(images):
            out = os.path.join(output_dir, f"slide_{i+1:03d}.png")
            img.save(out, "PNG"); pngs.append(out)
        return pngs
    except ImportError:
        errors.append("pdf2image not installed (also needs poppler on system)")
    except Exception as e:
        errors.append(f"pdf2image error: {e}")
    # Return errors as diagnostic info
    print(f"  [诊断] PDF→PNG 渲染失败原因: {'; '.join(errors)}", file=sys.stderr)
    return pngs

def render_with_com(pptx_path, output_dir, dpi=150):
    try: import win32com.client
    except ImportError: return None, "win32com not available"
    try:
        pp = win32com.client.Dispatch("PowerPoint.Application")
        pp.Visible = 0
        abs_pptx = os.path.abspath(pptx_path)
        abs_out = os.path.abspath(output_dir)
        os.makedirs(abs_out, exist_ok=True)
        pres = pp.Presentations.Open(abs_pptx, WithWindow=False)
        pres.Export(abs_out, "PNG")
        pres.Close(); pp.Quit()
        pngs = sorted(glob.glob(os.path.join(abs_out, "*.png")))
        if pngs:
            renamed = []
            for i, p in enumerate(pngs):
                new_name = os.path.join(abs_out, f"slide_{i+1:03d}.png")
                if p != new_name:
                    if os.path.exists(new_name): os.remove(new_name)
                    os.rename(p, new_name)
                renamed.append(new_name)
            return renamed, None
        return None, "No PNG exported"
    except Exception as e:
        return None, f"COM error: {e}"

def render_slides(pptx_path, output_dir="./slides_preview", dpi=150, engine="auto"):
    if not os.path.exists(pptx_path): return None, f"File not found: {pptx_path}"
    if engine == "auto":
        soffice = find_soffice()
        if soffice:
            result, err = render_with_libreoffice(pptx_path, output_dir, dpi, soffice)
            if result: return result, None
        ppt = find_powerpoint()
        if ppt:
            result, err = render_with_com(pptx_path, output_dir, dpi)
            if result: return result, None
        return None, "No rendering engine available"
    elif engine == "libreoffice":
        return render_with_libreoffice(pptx_path, output_dir, dpi)
    elif engine == "com":
        return render_with_com(pptx_path, output_dir, dpi)
    return None, f"Unknown engine: {engine}"

def main():
    parser = argparse.ArgumentParser(description="Render PPTX slides to PNG")
    parser.add_argument("pptx_path")
    parser.add_argument("-o", "--output", default="./slides_preview")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--engine", default="auto", choices=["auto", "libreoffice", "com"])
    args = parser.parse_args()
    print(f"Rendering: {args.pptx_path}")
    pngs, err = render_slides(args.pptx_path, args.output, args.dpi, args.engine)
    if err: print(f"ERROR: {err}", file=sys.stderr); sys.exit(1)
    print(f"Rendered {len(pngs)} slides:")
    for p in pngs: print(f"  {p}")

if __name__ == "__main__":
    main()
