"""Tests for the preview renderer adapter."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx_skill import auto_generate_ppt, render_preview


def _png_backend_available() -> bool:
    """True if at least one PDF→PNG backend is actually working."""
    try:
        import fitz  # noqa: F401  # PyMuPDF

        return True
    except ImportError:
        pass
    # pdf2image needs poppler on PATH; import alone is not enough.
    try:
        import shutil

        from pdf2image import pdfinfo_from_path  # noqa: F401

        if not shutil.which("pdfinfo"):
            return False
        return True
    except ImportError:
        pass
    return False


class PreviewRendererLogicTests(unittest.TestCase):
    """Tests that do not require a working PDF→PNG conversion pipeline."""

    def test_missing_file(self):
        result = render_preview("/nonexistent/file.pptx", output_dir="./out", engine="libreoffice")
        self.assertFalse(result.slide_pngs)
        self.assertTrue(any("not found" in str(a.get("error", "")).lower() for a in result.attempts))

    def test_auto_engine_preserves_concrete_errors(self):
        # Regression guard: when no engine succeeds, the returned attempts
        # must carry each backend's concrete error (e.g. "soffice not found")
        # rather than a generic "not found or failed" that hides the cause.
        from pptx_skill.preview_renderer import find_soffice

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            out = Path(tmp) / "preview"
            auto_generate_ppt(
                title="Errors",
                sections=[{"title": "A", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            result = render_preview(str(path), output_dir=str(out), dpi=96, engine="auto")
            if result.slide_pngs:
                self.skipTest("an engine succeeded; error-preservation path not exercised")
            self.assertFalse(result.slide_pngs)
            # Every attempt must have a non-generic error string.
            for attempt in result.attempts:
                self.assertNotIn("not found or failed", str(attempt.get("error", "")))
            # The LibreOffice attempt must mention soffice specifically when absent.
            if not find_soffice():
                lo = next((a for a in result.attempts if a.get("engine") == "libreoffice"), None)
                self.assertIsNotNone(lo)
                self.assertIn("soffice", str(lo.get("error", "")).lower())


@unittest.skipUnless(_png_backend_available(), "No PDF→PNG backend (PyMuPDF/poppler) available")
class PreviewRendererPngTests(unittest.TestCase):
    """Tests that require a working PDF→PNG conversion pipeline."""

    def test_libreoffice_render_result(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            out = Path(tmp) / "preview"
            auto_generate_ppt(
                title="Render",
                sections=[{"title": "A", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            result = render_preview(str(path), output_dir=str(out), dpi=150, engine="libreoffice")
            self.assertEqual(result.renderer, "libreoffice")
            self.assertTrue(len(result.slide_pngs) > 0)
            self.assertEqual(len(result.slide_pngs), len(result.actual_pixel_sizes))
            self.assertTrue(result.attempts)
            self.assertTrue(all(a["success"] for a in result.attempts))

    def test_legacy_wrapper(self):
        from pptx_skill import render_slides

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            out = Path(tmp) / "preview"
            auto_generate_ppt(
                title="Wrapper",
                sections=[{"title": "A", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            pngs, err = render_slides(str(path), output_dir=str(out), dpi=150, engine="libreoffice")
            self.assertIsNone(err)
            self.assertTrue(pngs)


if __name__ == "__main__":
    unittest.main()
