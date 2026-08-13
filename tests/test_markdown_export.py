"""Regression tests for Markdown export helpers."""
from __future__ import annotations

import unittest

from pptx_skill.markdown_export import _guess_title


class MarkdownTitleTests(unittest.TestCase):
    def test_guess_title_removes_only_balanced_markdown_markers(self):
        cases = {
            "## **Bold title**": "Bold title",
            "*Italic title*": "Italic title",
            "***Bold italic***": "Bold italic",
            "Release A*": "Release A*",
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(_guess_title(source), expected)


if __name__ == "__main__":
    unittest.main()
