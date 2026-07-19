"""Tests for TemplateProfileV2 schema, compiler, migration and adapter (PR7)."""
from __future__ import annotations

import json
import unittest
from pathlib import Path

from pptx_skill.design_schema import (
    REQUIRED_ROLES,
    StyleIntent,
    bundle_to_dict,
    resolve_all_tokens,
    resolve_token,
    validate_source_profile,
)
from pptx_skill.template_compiler import (
    COMPILER_VERSION,
    compile_from_intent,
    geometry_similarity,
    passes_diversity_gate,
    source_profile_from_intent,
)
from pptx_skill.template_v2_adapter import (
    migrate_and_compile_v1,
    migrate_profile_v1_to_v2,
    v2_bundle_to_legacy_fonts,
    v2_bundle_to_legacy_profile,
    v2_bundle_to_legacy_theme,
)

SKILL_ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = SKILL_ROOT / "assets" / "templates" / "catalog.json"


def _load_v1_catalog() -> dict:
    with CATALOG_PATH.open("r", encoding="utf-8") as fh:
        return json.load(fh)


class TokenResolutionTests(unittest.TestCase):
    def test_resolve_primitive_value(self):
        tokens = {"primitive": {"space": {"4": 16}}}
        self.assertEqual(resolve_token("primitive.space.4", tokens), 16)

    def test_resolve_ref_chain(self):
        tokens = {
            "primitive": {"palette": {"ink": "#1A1A1A"}},
            "semantic": {"color": {"text_primary": {"$ref": "primitive.palette.ink"}}},
        }
        self.assertEqual(resolve_token("semantic.color.text_primary", tokens), "#1A1A1A")

    def test_resolve_circular_ref_raises(self):
        tokens = {
            "a": {"$ref": "b"},
            "b": {"$ref": "a"},
        }
        with self.assertRaises(ValueError):
            resolve_all_tokens(tokens)

    def test_resolve_all_tokens_no_refs_left(self):
        tokens = {
            "primitive": {"palette": {"ink": "#111"}},
            "semantic": {"c": {"$ref": "primitive.palette.ink"}},
            "component": {"x": {"$ref": "semantic.c"}},
        }
        resolved = resolve_all_tokens(tokens)
        self.assertEqual(resolved["component"]["x"], "#111")
        self.assertEqual(resolved["semantic"]["c"], "#111")


class SchemaValidationTests(unittest.TestCase):
    def _minimal_profile(self) -> dict:
        return source_profile_from_intent(StyleIntent(industry="test"))

    def test_valid_source_profile_passes(self):
        validate_source_profile(self._minimal_profile())

    def test_rejects_unknown_top_level_key(self):
        profile = self._minimal_profile()
        profile["bogus_field"] = 1
        with self.assertRaises(ValueError):
            validate_source_profile(profile)

    def test_rejects_wrong_schema_version(self):
        profile = self._minimal_profile()
        profile["schema_version"] = 1
        with self.assertRaises(ValueError):
            validate_source_profile(profile)

    def test_rejects_missing_role_coverage(self):
        profile = self._minimal_profile()
        profile["coverage"] = {"status": "partial", "roles": ["not_a_role"]}
        with self.assertRaises(ValueError):
            validate_source_profile(profile)

    def test_rejects_invalid_color_in_palette(self):
        profile = self._minimal_profile()
        profile["tokens"]["primitive"]["palette"]["ink_950"] = "red"
        with self.assertRaises(ValueError):
            validate_source_profile(profile)

    def test_extensions_allowed(self):
        profile = self._minimal_profile()
        profile["extensions"] = {"legacy_v1": {"anything": True}}
        validate_source_profile(profile)


class CompilerTests(unittest.TestCase):
    def test_compile_from_intent_completes_all_roles(self):
        bundle = compile_from_intent(StyleIntent(industry="precision instruments"))
        self.assertTrue(bundle.is_complete())
        for role in REQUIRED_ROLES:
            self.assertIn(role, bundle.layouts)
            self.assertGreaterEqual(len(bundle.layouts[role]), 1)

    def test_compiled_tokens_have_no_refs(self):
        bundle = compile_from_intent(StyleIntent(industry="test", brand_colors=["#184E77", "#F4A261"]))
        self.assertEqual(bundle.tokens["semantic"]["color"]["accent"], "#F4A261")
        self.assertNotIn("$ref", json.dumps(bundle.tokens))

    def test_determinism_same_intent_same_seed(self):
        intent = StyleIntent(industry="fintech", audience="board", tone=["precise", "restrained"])
        b1 = compile_from_intent(intent, seed=42)
        b2 = compile_from_intent(intent, seed=42)
        self.assertEqual(b1.id, b2.id)
        self.assertEqual(bundle_to_dict(b1)["tokens"], bundle_to_dict(b2)["tokens"])

    def test_different_intent_different_bundle(self):
        b1 = compile_from_intent(StyleIntent(industry="fintech"), seed=1)
        b2 = compile_from_intent(StyleIntent(industry="healthcare"), seed=2)
        self.assertNotEqual(b1.id, b2.id)

    def test_recipe_provenance_marks_generated_vs_inherited(self):
        bundle = compile_from_intent(StyleIntent(industry="test"))
        prov_values = set(bundle.recipe_provenance.values())
        self.assertTrue(prov_values.issubset({"generated", "inherited", "migrated"}))
        # Core roles declared by the source profile are 'generated'.
        for rid in bundle.layouts["bullets"]:
            self.assertEqual(bundle.recipe_provenance.get(rid), "generated")

    def test_provenance_records_compiler_and_seed(self):
        bundle = compile_from_intent(StyleIntent(industry="test"), seed=12345)
        self.assertEqual(bundle.provenance["generator"], COMPILER_VERSION)
        self.assertEqual(bundle.provenance["seed"], 12345)


class DiversityGateTests(unittest.TestCase):
    def test_identical_bundles_similar(self):
        b = compile_from_intent(StyleIntent(industry="test"), seed=1)
        self.assertGreater(geometry_similarity(b, b), 0.99)

    def test_diversity_gate_rejects_identical(self):
        b = compile_from_intent(StyleIntent(industry="test"), seed=1)
        passes, sim = passes_diversity_gate(b, [b])
        self.assertFalse(passes)
        self.assertGreater(sim, 0.85)

    def test_diversity_gate_passes_distinct(self):
        b1 = compile_from_intent(StyleIntent(industry="fintech"), seed=1)
        b2 = compile_from_intent(StyleIntent(industry="healthcare", geometry=["axes"]), seed=2)
        passes, sim = passes_diversity_gate(b2, [b1])
        self.assertTrue(passes or sim <= 0.85)


class V1MigrationTests(unittest.TestCase):
    def test_all_catalog_profiles_migrate(self):
        catalog = _load_v1_catalog()
        for key, v1 in catalog["templates"].items():
            source = migrate_profile_v1_to_v2(v1, key=key)
            self.assertEqual(source["schema_version"], 2)
            self.assertEqual(source["kind"], "template-source-profile")
            # The source profile must validate.
            validate_source_profile(source)
            # Round-trip: legacy V1 payload preserved under extensions.
            self.assertEqual(source["extensions"]["legacy_v1"]["id"], v1.get("id", key))

    def test_migration_and_compile_completes_roles(self):
        catalog = _load_v1_catalog()
        v1 = catalog["templates"]["strategy-consulting"]
        bundle = migrate_and_compile_v1(v1)
        self.assertTrue(bundle.is_complete())
        for role in REQUIRED_ROLES:
            self.assertIn(role, bundle.layouts)

    def test_v2_to_legacy_theme_has_required_fields(self):
        catalog = _load_v1_catalog()
        bundle = migrate_and_compile_v1(catalog["templates"]["strategy-consulting"])
        theme = v2_bundle_to_legacy_theme(bundle)
        for field in ("bg", "primary", "accent", "text", "on_dark"):
            self.assertIn(field, theme)
        # primary must be a hex color.
        self.assertTrue(theme["primary"].startswith("#"))

    def test_v2_to_legacy_fonts_preserves_families(self):
        catalog = _load_v1_catalog()
        bundle = migrate_and_compile_v1(catalog["templates"]["academic-clean"])
        fonts = v2_bundle_to_legacy_fonts(bundle)
        self.assertIn("cn", fonts)
        self.assertIn("en", fonts)

    def test_v2_to_legacy_profile_consumable_shape(self):
        catalog = _load_v1_catalog()
        bundle = migrate_and_compile_v1(catalog["templates"]["executive-dark"])
        profile = v2_bundle_to_legacy_profile(bundle)
        self.assertIn("theme", profile)
        self.assertIn("fonts", profile)
        self.assertIn("layout_family", profile)
        self.assertIn(profile["layout_family"], {"standard", "editorial_grid", "technical_axis", "poster_column"})

    def test_migration_preserves_preferred_sequence_in_provenance(self):
        catalog = _load_v1_catalog()
        v1 = catalog["templates"]["data-story"]
        source = migrate_profile_v1_to_v2(v1)
        self.assertEqual(
            source["provenance"]["legacy_preferred_sequence"],
            v1["preferred_sequence"],
        )


if __name__ == "__main__":
    unittest.main()
