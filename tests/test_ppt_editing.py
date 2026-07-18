import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from ppt_edit import edit_text, recolor  # noqa: E402
from ppt_pages import delete_slide, duplicate_slide, insert_slide, move_slide, replace_layout  # noqa: E402
from ppt_project import load_project  # noqa: E402
from pptx_helper import auto_generate_ppt  # noqa: E402


def deck_text(path: Path) -> str:
    prs = Presentation(str(path))
    return "\n".join(
        shape.text for slide in prs.slides for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
    )


class PptEditingTests(unittest.TestCase):
    def _make_deck(self, directory: Path) -> Path:
        path = directory / "deck.pptx"
        auto_generate_ppt(
            title="原始标题",
            subtitle="副标题",
            sections=[
                {"title": "第一页", "bullets": ["要点 A", "要点 B"]},
                {"title": "第二页", "steps": ["发现", "设计", "交付"]},
            ],
            output_path=str(path),
            theme_key="editorial",
            auto_search_images=False,
        )
        return path

    def test_manifest_is_embedded_and_sidecar_fallback_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            deck = self._make_deck(directory)
            sidecar = deck.with_suffix(".manifest.json")
            self.assertTrue(sidecar.exists())
            project = load_project(deck)
            self.assertEqual(project["title"], "原始标题")
            sidecar.unlink()
            embedded = load_project(deck)
            self.assertEqual(embedded["sections"][0]["title"], "第一页")
            Presentation(str(deck))

    def test_text_color_and_page_operations_create_backups(self):
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            deck = self._make_deck(directory)
            self.assertEqual(edit_text(deck, 1, "原始标题", "修改后标题"), 1)
            self.assertIn("修改后标题", deck_text(deck))
            self.assertTrue(directory.joinpath("deck.bak.pptx").exists())
            self.assertGreater(recolor(deck, "#1A1A1A", "#123456"), 0)

            self.assertEqual(duplicate_slide(deck, 2, 3), 5)
            self.assertEqual(delete_slide(deck, 3), 4)
            move_slide(deck, 2, 3)
            self.assertEqual(insert_slide(
                deck, 2,
                {"title": "插入页", "bullets": ["新增内容"]},
                layout="bullets", auto_search_images=False,
            ), 5)
            self.assertIn("插入页", deck_text(deck))
            replace_layout(
                deck, 2, "quote",
                {"title": "引言", "quote": "重新设计这一页", "source": "Test"},
            )
            self.assertIn("重新设计这一页", deck_text(deck))


if __name__ == "__main__":
    unittest.main()
