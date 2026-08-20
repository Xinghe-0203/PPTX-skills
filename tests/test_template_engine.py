import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from template_engine import (  # noqa: E402
    generate_template_preview,
    generate_template_profile,
    get_generated_template_dir,
    list_templates,
    load_template_profile,
    register_template_profile,
    validate_template_profile,
)


class TemplateEngineTests(unittest.TestCase):
    def test_builtin_catalog_has_multiple_reusable_templates(self):
        templates = list_templates()
        self.assertGreaterEqual(len(templates), 12)
        self.assertIn("strategy-consulting", {item["key"] for item in templates})
        self.assertEqual(
            {"editorial_grid", "technical_axis", "poster_column"},
            {item["layout_family"] for item in templates if not item.get("generated")},
        )

    def test_generate_template_profile_is_deterministic_and_valid(self):
        first = generate_template_profile("Aurora", "深色科技产品发布，克制、现代")
        second = generate_template_profile("Aurora", "深色科技产品发布，克制、现代")
        validate_template_profile(first)
        self.assertEqual(first, second)
        self.assertTrue(first["theme"]["on_dark"])
        self.assertEqual(first["layout_family"], "poster_column")

    def test_generated_profiles_choose_distinct_layout_families(self):
        self.assertEqual(
            generate_template_profile("Swiss", "瑞士编辑网格")['layout_family'],
            "editorial_grid",
        )
        self.assertEqual(
            generate_template_profile("Data", "数据工程技术报告")['layout_family'],
            "technical_axis",
        )

    def test_generate_editable_preview(self):
        profile = load_template_profile("academic-clean")
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "preview.pptx"
            generate_template_preview(profile, output)
            self.assertTrue(output.exists())
            self.assertGreater(output.stat().st_size, 10_000)

    def test_registered_profiles_use_dynamic_user_data_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            profile = generate_template_profile("User Scoped", "克制的编辑网格")
            with patch.dict(os.environ, {"PPTX_SKILL_TEMPLATE_DIR": tmp}):
                target = register_template_profile(profile)
                self.assertEqual(target.parent, Path(tmp))
                self.assertEqual(get_generated_template_dir(), Path(tmp))
                self.assertEqual(load_template_profile(profile["id"])["name"], "User Scoped")
                matching = [item for item in list_templates() if item["key"] == profile["id"]]
                self.assertEqual(len(matching), 1)
                self.assertTrue(matching[0]["generated"])


if __name__ == "__main__":
    unittest.main()
