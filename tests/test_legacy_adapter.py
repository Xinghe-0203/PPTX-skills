"""Tests for the legacy Section -> ContentSpec adapter."""
from __future__ import annotations

import unittest

from pptx_skill.content_adapter import LegacyContentAdapter, _infer_layout


class LegacyAdapterTests(unittest.TestCase):
    def test_infer_layout_signals(self):
        self.assertEqual(_infer_layout({"metrics": [{"v": 1}]}), "dashboard")
        self.assertEqual(_infer_layout({"events": ["e"]}), "timeline")
        self.assertEqual(_infer_layout({"left": {"title": "L"}, "right": {"title": "R"}}), "comparison")
        self.assertEqual(_infer_layout({"quote": "q"}), "quote")
        self.assertEqual(_infer_layout({"steps": ["s"]}), "process")
        self.assertEqual(_infer_layout({"table_headers": ["h"]}), "table")
        self.assertEqual(_infer_layout({"images": ["a", "b", "c"]}), "image_grid")
        self.assertEqual(_infer_layout({"images": ["a"], "bullets": ["b"]}), "text_image")
        self.assertEqual(_infer_layout({"images": ["a"]}), "full_image")
        self.assertEqual(_infer_layout({"bullets": ["b"]}), "bullets")
        self.assertEqual(_infer_layout({}), "section")

    def test_adapt_basic_deck(self):
        adapter = LegacyContentAdapter(title="Report", subtitle="Q3")
        content = adapter.adapt(
            [
                {"title": "Cover", "kicker": "Intro"},
                {"title": "Summary", "bullets": ["A", "B"]},
                {"title": "Data", "metrics": [{"label": "Rev", "value": "10"}]},
            ]
        )
        self.assertEqual(content.title, "Report")
        self.assertEqual(len(content.slides), 3)
        self.assertEqual(content.slides[0].role, "section")  # no strong signal
        self.assertTrue(any(e.role == "kicker" for e in content.slides[0].elements))
        self.assertEqual(content.slides[1].role, "bullets")
        self.assertEqual(content.slides[2].role, "dashboard")

    def test_stable_ids(self):
        adapter = LegacyContentAdapter(title="Report")
        content1 = adapter.adapt([{"title": "A"}, {"title": "B"}])
        content2 = adapter.adapt([{"title": "A"}, {"title": "B"}])
        self.assertEqual(content1.slides[0].id, content2.slides[0].id)
        self.assertEqual(content1.slides[1].id, content2.slides[1].id)


if __name__ == "__main__":
    unittest.main()
