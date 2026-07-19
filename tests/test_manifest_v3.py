"""Tests for Manifest V3 persistence and legacy v2 migration."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx_skill import auto_generate_ppt
from pptx_skill.manifest import (
    ManifestV3,
    load_manifest,
    manifest_from_dict,
    manifest_to_dict,
    migrate_v2_to_v3,
    save_manifest_v3,
    set_current_content,
)


class ManifestV3Tests(unittest.TestCase):
    def test_round_trip_via_sidecar(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "deck.pptx"
            auto_generate_ppt(
                title="RT",
                sections=[{"title": "A", "bullets": ["1"]}],
                output_path=str(path),
                theme_key="editorial",
                auto_search_images=False,
            )
            manifest = load_manifest(path)
            self.assertIsNotNone(manifest)
            self.assertEqual(manifest.manifest_schema_version, 3)

            save_manifest_v3(path, manifest)
            reloaded = load_manifest(path)
            self.assertEqual(reloaded.manifest_schema_version, 3)
            self.assertEqual(
                reloaded.current["content"]["title"],
                manifest.current["content"]["title"],
            )

    def test_legacy_v2_migration(self):
        v2 = {
            "skill_version": 2,
            "title": "Legacy",
            "subtitle": "Deck",
            "sections": [
                {"title": "Intro", "bullets": ["a", "b"]},
                {"title": "Data", "metrics": [{"label": "X", "value": "10"}]},
            ],
            "layouts": ["cover", "bullets", "dashboard", "end"],
        }
        manifest = migrate_v2_to_v3(v2)
        self.assertEqual(manifest.manifest_schema_version, 3)
        self.assertEqual(manifest.legacy["source_skill_version"], 2)
        self.assertEqual(len(manifest.current["content"]["slides"]), 2)
        self.assertEqual(
            manifest.current["content"]["slides"][1]["elements"][0]["role"],
            "title",
        )

    def test_dict_serialization(self):
        manifest = ManifestV3(current={"content": {"title": "T"}})
        manifest.record_attempt(0, plan_artifact="plan.json")
        data = manifest_to_dict(manifest)
        restored = manifest_from_dict(data)
        self.assertEqual(restored.manifest_schema_version, 3)
        self.assertEqual(len(restored.attempts), 1)
        self.assertTrue(restored.attempts[0].run_id)

    def test_set_current_content(self):
        from pptx_skill.content_model import ContentSpec, SlideSpec

        slide = SlideSpec(
            id="s1",
            role="bullets",
            communication_goal="intro",
            elements=[],
        )
        content = ContentSpec(
            id="deck", title="T", subtitle="S", slides=[slide]
        )
        manifest = ManifestV3()
        set_current_content(manifest, content)
        self.assertEqual(manifest.current["content"]["title"], "T")


if __name__ == "__main__":
    unittest.main()
