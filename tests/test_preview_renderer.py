"""Tests for the preview renderer adapter."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx_skill import auto_generate_ppt, render_preview


class PreviewRendererTests(unittest.TestCase):
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

    def test_missing_file(self):
        result = render_preview("/nonexistent/file.pptx", output_dir="./out", engine="libreoffice")
        self.assertFalse(result.slide_pngs)
        self.assertTrue(any("not found" in str(a.get("error", "")).lower() for a in result.attempts))

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
