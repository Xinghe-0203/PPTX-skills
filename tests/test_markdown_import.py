"""Tests for the v6.0 ``pptx_skill.markdown_import`` module.

Covers ``markdown_to_sections``, ``from_markdown`` and ``import_markdown``
following the project's testing conventions: ``unittest.TestCase``, no
shared fixtures, ``tempfile.TemporaryDirectory`` for all I/O.
"""
from __future__ import annotations

import os
import tempfile
import unittest

from pptx import Presentation

from pptx_skill.markdown_import import (
    from_markdown,
    import_markdown,
    markdown_to_sections,
)


class MarkdownToSectionsTests(unittest.TestCase):
    # 1. YAML front-matter ------------------------------------------------
    def test_yaml_front_matter_title_and_subtitle(self):
        md = "---\ntitle: T\nsubtitle: S\n---\n# H1"
        sections = markdown_to_sections(md)
        cover = sections[0]
        self.assertEqual(cover["layout"], "cover")
        self.assertEqual(cover["title"], "T")
        self.assertEqual(cover["subtitle"], "S")

    # 2. H1 as cover (no front-matter) -----------------------------------
    def test_h1_becomes_cover(self):
        md = "# 标题"
        sections = markdown_to_sections(md)
        cover = sections[0]
        self.assertEqual(cover["layout"], "cover")
        self.assertEqual(cover["title"], "标题")

    # 3. H2 as section with bullets --------------------------------------
    def test_h2_section_with_bullets(self):
        md = "# 封面\n## 第一节\n- 要点 1\n- 要点 2"
        sections = markdown_to_sections(md)
        content = [s for s in sections if s.get("title") == "第一节"]
        self.assertEqual(len(content), 1)
        self.assertEqual(content[0]["bullets"], ["要点 1", "要点 2"])

    # 4. GFM table --------------------------------------------------------
    def test_gfm_table_parsing(self):
        md = "# 数据\n| a | b |\n|---|---|\n| 1 | 2 |"
        sections = markdown_to_sections(md)
        tables = [s for s in sections if s.get("layout") == "table"]
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0]["table_headers"], ["a", "b"])
        self.assertEqual(tables[0]["table_rows"], [["1", "2"]])

    # 5. Code block -------------------------------------------------------
    def test_code_block_preserved_as_bullets(self):
        md = "# 代码\n## 示例\n```python\nprint('hi')\n```\n"
        sections = markdown_to_sections(md)
        code_sections = [s for s in sections if s.get("title") == "示例"]
        self.assertEqual(len(code_sections), 1)
        self.assertIn("print('hi')", code_sections[0].get("bullets", []))

    # 6. Blockquote -------------------------------------------------------
    def test_blockquote_parsed(self):
        md = "# 引用页\n> This is a quote"
        sections = markdown_to_sections(md)
        quotes = [s for s in sections if s.get("layout") == "quote"]
        self.assertEqual(len(quotes), 1)
        self.assertIn("This is a quote", quotes[0]["quote"])

    # 7. lang parameter is passed through --------------------------------
    def test_lang_passed_to_sections(self):
        md = "# Title"
        sections = markdown_to_sections(md, lang="en")
        opts = sections[0].get("layout_opts", {})
        self.assertEqual(opts.get("lang"), "en")

    # 8. Empty / whitespace-only input -----------------------------------
    def test_empty_and_whitespace_input(self):
        for md in ("", "   \n   \n"):
            sections = markdown_to_sections(md)
            self.assertIsInstance(sections, list)
            self.assertGreaterEqual(len(sections), 1)
            self.assertEqual(sections[0]["layout"], "cover")

    # 9. import_markdown end-to-end --------------------------------------
    def test_import_markdown_generates_pptx(self):
        with tempfile.TemporaryDirectory() as tmp:
            md_path = os.path.join(tmp, "input.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write("# 测试标题\n## 第一节\n- 要点一\n- 要点二\n")
            out_path = os.path.join(tmp, "output.pptx")
            result = import_markdown(md_path, out_path)
            self.assertTrue(os.path.exists(result))
            prs = Presentation(out_path)
            self.assertGreaterEqual(len(prs.slides), 1)


class FromMarkdownFileTests(unittest.TestCase):
    """Bonus coverage for the file-based ``from_markdown`` wrapper."""

    def test_from_markdown_returns_sections(self):
        with tempfile.TemporaryDirectory() as tmp:
            md_path = os.path.join(tmp, "outline.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write("# 标题\n## 节\n- 要点\n")
            sections = from_markdown(md_path)
            self.assertIsInstance(sections, list)
            self.assertGreaterEqual(len(sections), 1)
            self.assertEqual(sections[0]["layout"], "cover")


if __name__ == "__main__":
    unittest.main()
