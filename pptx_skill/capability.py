"""Runtime capability report for the PPTX skill.

Reports interpreter version, installed dependencies, optional extras,
rendering engines, and font/text shaping capabilities.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Any


def _module_version(name: str) -> str | None:
    try:
        mod = __import__(name)
    except ImportError:
        return None
    for attr in ("__version__", "VERSION", "version"):
        try:
            value = getattr(mod, attr)
            if isinstance(value, str):
                return value
        except AttributeError:
            continue
    return "installed"


def _pil_capabilities() -> dict[str, Any]:
    try:
        from PIL import features
    except ImportError:
        return {"available": False}
    return {
        "available": True,
        "raqm": features.check("raqm"),
    }


def _libreoffice_version(soffice_path: str | None) -> dict[str, Any] | None:
    if not soffice_path:
        return None
    import subprocess
    # On some Windows installs soffice --version hangs or launches a window;
    # a short timeout and fallback to path-only is acceptable for a capability report.
    try:
        result = subprocess.run(
            [soffice_path, "--version"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        text = (result.stdout + result.stderr).strip()
        if result.returncode == 0 and text:
            return {"path": soffice_path, "version_output": text}
        return {"path": soffice_path, "version_output": text or None}
    except subprocess.TimeoutExpired:
        return {"path": soffice_path, "version_output": None, "note": "--version timed out"}
    except Exception as exc:
        return {"path": soffice_path, "error": str(exc)}


def report_capabilities() -> dict[str, Any]:
    """Return a serializable dict describing the current runtime environment."""
    optional_deps = {
        "python-pptx": "pptx",
        "Pillow": "PIL",
        "kiwisolver": "kiwisolver",
        "pydantic": "pydantic",
        "numpy": "numpy",
        "scikit-image": "skimage",
        "opencv-python-headless": "cv2",
        "shapely": "shapely",
        "pytesseract": "pytesseract",
        "PyMuPDF": "fitz",
        "pdf2image": "pdf2image",
        "pywin32": "win32com",
        "fonttools": "fontTools",
        "uharfbuzz": "uharfbuzz",
    }

    deps = {}
    for distribution_name, import_name in optional_deps.items():
        deps[distribution_name] = _module_version(import_name)

    return {
        "interpreter": {
            "executable": sys.executable,
            "version": sys.version,
        },
        "dependencies": deps,
        "pil": _pil_capabilities(),
        "engines": {
            "libreoffice": _libreoffice_version(
                next(
                    (p for p in [
                        shutil.which("soffice"),
                        r"C:\Program Files\LibreOffice\program\soffice.exe",
                        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                    ] if p and Path(p).exists()),
                    None,
                )
            ),
            "powerpoint": {
                "path": next(
                    (p for p in [
                        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
                        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
                    ] if Path(p).exists()),
                    None,
                )
            },
        },
    }


def main() -> int:
    import json
    print(json.dumps(report_capabilities(), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
