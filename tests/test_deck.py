"""Tests for the v6.0 ``pptx_skill.deck.Deck`` fluent API.

Covers generation, open/save lifecycle, chained calls, text editing,
notes, transitions, validation, inspection, context-manager support and
the ``__len__`` / ``__repr__`` dunders.
"""
from __future__ import annotations

import os
import tempfile
import unittest

from pptx import Presentation

from pptx_skill import Deck


def _make_deck(directory: str) -> str:
    """Generate a small deck into *directory* and return its path."""
    path = os.path.join(directory, "deck.pptx")
    Deck.generate(
        title="测试标题",
        subtitle="副标题",
        sections=[{"title": "第一节", "bullets": ["要点 A", "要点 B"]}],
        output_path=path,
    )
    return path


class DeckLifecycleTests(unittest.TestCase):
    # 1. generate + open + save ------------------------------------------
    def test_generate_open_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            self.assertTrue(os.path.exists(path))
            deck = Deck.open(path)
            self.assertGreater(deck.slide_count, 0)
            saved = deck.save()
            self.assertEqual(saved, path)

    # 2. Chained calls return Deck ---------------------------------------
    def test_chained_calls_return_deck(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            result = Deck.open(path).add_notes(1, "备注内容")
            self.assertIsInstance(result, Deck)
            result.save()

    # 3. find_replace_all ------------------------------------------------
    def test_find_replace_all_returns_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            deck = Deck.open(path)
            count = deck.find_replace_all("要点", "项目")
            self.assertIsInstance(count, int)
            self.assertGreaterEqual(count, 0)
            deck.save()

    # 4. add_notes --------------------------------------------------------
    def test_add_notes_persisted(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            deck = Deck.open(path)
            deck.add_notes(1, "这是演讲备注")
            deck.save()
            prs = Presentation(path)
            slide = prs.slides[0]
            self.assertTrue(slide.has_notes_slide)
            self.assertIn("这是演讲备注", slide.notes_slide.notes_text_frame.text)

    # 5. validate ---------------------------------------------------------
    def test_validate_returns_dict(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            deck = Deck.open(path)
            result = deck.validate()
            self.assertIsInstance(result, dict)
            self.assertIn("passed", result)

    # 6. inspect / info --------------------------------------------------
    def test_inspect_and_info_return_dicts(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            deck = Deck.open(path)
            insp = deck.inspect()
            self.assertIsInstance(insp, dict)
            self.assertTrue(insp)
            info = deck.info()
            self.assertIsInstance(info, dict)
            self.assertIn("slide_count", info)
            self.assertGreater(info["slide_count"], 0)

    # 7. Context manager --------------------------------------------------
    def test_context_manager(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            with Deck.open(path) as d:
                count = d.slide_count
            self.assertGreater(count, 0)

    # 8. __len__ and __repr__ --------------------------------------------
    def test_len_and_repr(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _make_deck(tmp)
            deck = Deck.open(path)
            self.assertEqual(len(deck), deck.slide_count)
            self.assertIn("Deck", repr(deck))

    # 9. from_markdown ----------------------------------------------------
    def test_from_markdown_generates_deck(self):
        with tempfile.TemporaryDirectory() as tmp:
            md_path = os.path.join(tmp, "outline.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write("# 标题\n## 节\n- 要点\n")
            out_path = os.path.join(tmp, "out.pptx")
            deck = Deck.from_markdown(md_path, output_path=out_path)
            self.assertGreaterEqual(deck.slide_count, 1)


if __name__ == "__main__":
    unittest.main()
