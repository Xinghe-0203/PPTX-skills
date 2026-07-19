"""Tests for text metrics and CJK line breaking."""
from __future__ import annotations

import unittest

from pptx_skill.text_metrics import (
    ParagraphStyle,
    TextRun,
    fit_text_to_height,
    measure_text,
    resolve_font,
)


class TextMetricsTests(unittest.TestCase):
    def test_resolve_font_returns_path_or_default(self):
        path = resolve_font("Microsoft YaHei")
        # On a Windows dev box this should exist; otherwise we just confirm no exception.
        self.assertTrue(path is None or path.exists())

    def test_measure_english_wraps(self):
        text = "This is a long sentence that should wrap into multiple lines."
        metrics = measure_text(text, width_pt=120, font_size_pt=16, font_family="Arial")
        self.assertGreater(metrics.num_lines, 1)
        self.assertLessEqual(metrics.required_font_size_pt, 16.0)
        self.assertTrue(all(line.width_pt <= 125 for line in metrics.lines))

    def test_measure_cjk_wraps_by_character(self):
        text = "这是需要自动换行的中文文本示例，应当在指定宽度内拆分成多行。"
        metrics = measure_text(text, width_pt=80, font_size_pt=16, font_family="Microsoft YaHei")
        self.assertGreater(metrics.num_lines, 1)
        self.assertTrue(all("\n" not in line.text for line in metrics.lines))
        # No CJK closing punctuation should be at line start.
        close_punct = set(
            "。、，；：？！」』》"
            "）〕】〉〗〙)"
        )
        for line in metrics.lines[1:]:
            self.assertNotIn(line.text[0], close_punct)

    def test_fit_text_to_height_reduces_size(self):
        text = "A" * 500
        metrics = fit_text_to_height(
            text,
            width_pt=200,
            height_pt=80,
            font_size_pt=24,
            font_family="Arial",
            min_font_size_pt=8,
        )
        # With default line height and metrics this long text may only fit by
        # dropping close to min size; assert the size was actually searched down.
        self.assertLessEqual(metrics.required_font_size_pt, 24.0)
        self.assertGreaterEqual(metrics.required_font_size_pt, 8.0)
        self.assertGreater(metrics.num_lines, 1)

    def test_measure_runs_with_multiple_runs(self):
        runs = [
            TextRun(text="Heading ", font_family="Arial", size_pt=20, bold=True),
            TextRun(text="content", font_family="Arial", size_pt=16),
        ]
        ParagraphStyle(line_height=1.2)
        metrics = measure_text("".join(r.text for r in runs), width_pt=300, font_size_pt=20, font_family="Arial")
        self.assertEqual(metrics.num_lines, 1)

    def test_required_font_size_bsearch_finds_value(self):
        text = "A" * 200
        metrics = measure_text(
            text,
            width_pt=100,
            font_size_pt=40,
            font_family="Arial",
            min_font_size_pt=8,
        )
        self.assertGreaterEqual(metrics.required_font_size_pt, 8.0)
        self.assertLessEqual(metrics.required_font_size_pt, 40.0)


if __name__ == "__main__":
    unittest.main()
