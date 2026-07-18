"""Tests for image cropping utilities."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from PIL import Image

from pptx_skill.image_crop import crop_contain, crop_cover, crop_smart


class ImageCropTests(unittest.TestCase):
    def _make_image(self, directory: Path, size: tuple[int, int]) -> Path:
        path = directory / "img.png"
        Image.new("RGB", size, color=(100, 150, 200)).save(path)
        return path

    def _effective_ratio(self, src_w: float, src_h: float, result) -> float:
        effective_w = src_w * (1 - result.crop_fractions["left"] - result.crop_fractions["right"])
        effective_h = src_h * (1 - result.crop_fractions["top"] - result.crop_fractions["bottom"])
        return effective_w / effective_h

    def test_cover_preserves_aspect(self):
        src_w, src_h = 800, 600
        result = crop_cover(src_w, src_h, 16, 9)
        self.assertAlmostEqual(self._effective_ratio(src_w, src_h, result), 16 / 9, places=5)
        self.assertTrue(0 <= result.left < result.right <= 1)
        self.assertTrue(0 <= result.top < result.bottom <= 1)

    def test_contain_fits_inside(self):
        src_w, src_h = 800, 600
        result = crop_contain(src_w, src_h, 16, 9)
        # contain keeps the whole image visible, so effective ratio equals source ratio.
        self.assertAlmostEqual(self._effective_ratio(src_w, src_h, result), src_w / src_h, places=5)
        self.assertTrue(0 <= result.left < result.right <= 1)
        self.assertTrue(0 <= result.top < result.bottom <= 1)

    def test_crop_fractions_sum_consistently(self):
        result = crop_cover(800, 600, 16, 9)
        f = result.crop_fractions
        self.assertTrue(0 <= f["left"] + f["right"] <= 1)
        self.assertTrue(0 <= f["top"] + f["bottom"] <= 1)

    def test_smart_crop_returns_valid_rect(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_image(Path(tmp), (400, 300))
            src_w, src_h = 400, 300
            result = crop_smart(str(path), 16, 9)
            self.assertTrue(0 <= result.left < result.right <= 1)
            self.assertTrue(0 <= result.top < result.bottom <= 1)
            ratio = self._effective_ratio(src_w, src_h, result)
            self.assertAlmostEqual(ratio, 16 / 9, places=2)

    def test_smart_crop_with_boost(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_image(Path(tmp), (400, 300))
            result = crop_smart(str(path), 16, 9, boost_boxes=[(0.1, 0.1, 0.3, 0.3)])
            self.assertTrue(0 <= result.left < result.right <= 1)


if __name__ == "__main__":
    unittest.main()
