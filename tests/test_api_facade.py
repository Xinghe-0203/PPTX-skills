"""Tests for the api.py compatibility facade."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx_skill import (
    GenerationResult,
    auto_generate_ppt,
    auto_validate_ppt,
)
from pptx_skill.visual_qa import CheckOutcome


class ApiFacadeTests(unittest.TestCase):
    def test_legacy_string_return(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            result = auto_generate_ppt(
                title="兼容",
                sections=[{"title": "页", "bullets": ["1", "2"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            self.assertIsInstance(result, str)
            self.assertTrue(Path(result).exists())

    def test_return_result_structured(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            result = auto_generate_ppt(
                title="结构化",
                sections=[{"title": "页", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
                qa_mode="report",
                return_result=True,
            )
            self.assertIsInstance(result, GenerationResult)
            self.assertEqual(result.qa_status, CheckOutcome.PASS)
            self.assertTrue(Path(result.pptx_path).exists())

    def test_auto_validate_legacy_format(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            auto_generate_ppt(
                title="验证",
                sections=[{"title": "页", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            report = auto_validate_ppt(str(path))
            self.assertIn("passed", report)
            self.assertIn("checks", report)
            self.assertIn("warnings", report)
            self.assertIn("total_slides", report)

    def test_auto_validate_new_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            auto_generate_ppt(
                title="验证",
                sections=[{"title": "页", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            report = auto_validate_ppt(str(path), return_report=True)
            self.assertEqual(report.status, CheckOutcome.PASS)

    def test_facade_validate_matches_legacy_checks(self):
        # The facade validate must run the same five checks as the legacy
        # implementation (no silent loss of font-hierarchy / color-restraint).
        import sys
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
        from pptx_helper import auto_validate_ppt as legacy_validate

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            auto_generate_ppt(
                title="一致性",
                sections=[
                    {"title": "页一", "bullets": ["1", "2", "3"]},
                    {"title": "页二", "steps": ["发现", "设计", "交付"]},
                ],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            legacy = legacy_validate(str(path))
            facade = auto_validate_ppt(str(path))
            self.assertEqual(
                [c[0] for c in facade["checks"]],
                [c[0] for c in legacy["checks"]],
            )
            self.assertEqual(facade["passed"], legacy["passed"])
            self.assertEqual(facade["total_slides"], legacy["total_slides"])


if __name__ == "__main__":
    unittest.main()
