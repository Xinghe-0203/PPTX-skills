"""render_slides.py - PPTX slides to PNG for visual review.

PR2 migration: the authoritative renderer implementation now lives in
``pptx_skill.preview_renderer``. This file remains a thin CLI and legacy
wrapper so existing commands and imports keep working.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Ensure pptx_skill is importable when running as a script (python scripts/render_slides.py).
# When installed via pip, this is unnecessary; when running from the repo, sys.path[0] is
# scripts/ and the package root is one level up.
_repo_root = str(Path(__file__).resolve().parent.parent)
if _repo_root not in sys.path:
    sys.path.insert(0, _repo_root)

from pptx_skill.preview_renderer import render_preview


def render_slides(pptx_path, output_dir="./slides_preview", dpi=150, engine="auto"):
    """Legacy wrapper returning ``(pngs, err)`` for compatibility."""
    result = render_preview(pptx_path, output_dir=output_dir, dpi=dpi, engine=engine)
    if result.slide_pngs:
        return result.slide_pngs, None
    messages = [
        f"{a.get('engine')}: {a.get('error')}"
        for a in result.attempts
        if not a.get("success")
    ]
    return None, {"error": "No rendering engine succeeded", "attempts": result.attempts, "details": "; ".join(messages)}


def main():
    parser = argparse.ArgumentParser(description="Render PPTX slides to PNG")
    parser.add_argument("pptx_path")
    parser.add_argument("-o", "--output", default="./slides_preview")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--engine", default="auto", choices=["auto", "libreoffice", "com"])
    args = parser.parse_args()
    print(f"Rendering: {args.pptx_path}")
    result = render_preview(
        args.pptx_path,
        output_dir=args.output,
        dpi=args.dpi,
        engine=args.engine,
    )
    if not result.slide_pngs:
        for a in result.attempts:
            if not a.get("success"):
                engine = a.get("engine", "unknown")
                error = a.get("error", "unknown error")
                print(f"ERROR [{engine}]: {error}", file=sys.stderr)
                if a.get("stdout"):
                    print(f"  stdout: {a['stdout'][:500]}", file=sys.stderr)
                if a.get("stderr"):
                    print(f"  stderr: {a['stderr'][:500]}", file=sys.stderr)
        sys.exit(1)
    print(f"Rendered {len(result.slide_pngs)} slides:")
    for p in result.slide_pngs:
        print(f"  {p}")


if __name__ == "__main__":
    main()
