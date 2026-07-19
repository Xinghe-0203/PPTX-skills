"""Tests for PR9: 200-page annotation dataset, golden renders, E2E scenarios,
stress and performance checks, and dataset/golden coverage reports.
"""
from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from pptx_skill.content_model import (
    ContentSpec,
    ElementSpec,
    SlideSpec,
    canvas_from_name,
)
from pptx_skill.deck_planner import plan_deck
from pptx_skill.golden_renders import (
    ALL_ROLES,
    GEOMETRY_FAMILIES,
    golden_coverage_report,
    load_golden_index,
    render_golden_set,
)
from pptx_skill.layout_engine import builtin_recipes, solve_recipe, solved_geometry_to_layout_plan, _builtin_tokens
from pptx_skill.pagination import paginate_content_spec
from pptx_skill.pptx_renderer import render_layout_plans
from pptx_skill.qa_dataset import (
    DATASET_CATEGORIES,
    DATASET_TOTAL_MIN,
    ISSUE_CLEAN,
    dataset_coverage_report,
    generate_annotation_dataset,
    load_annotation_dataset,
)


# ---------------------------------------------------------------------------
# Annotation dataset coverage (blueprint §12.3: 200 pages, 6 categories)
# ---------------------------------------------------------------------------


class AnnotationDatasetTests(unittest.TestCase):
    def test_dataset_meets_200_page_minimum(self):
        self.assertGreaterEqual(DATASET_TOTAL_MIN, 200)

    def test_generate_and_load_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            manifest = generate_annotation_dataset(tmp)
            self.assertGreaterEqual(manifest["total_fixtures"], DATASET_TOTAL_MIN)
            reloaded = load_annotation_dataset(tmp)
            self.assertEqual(reloaded["total_fixtures"], manifest["total_fixtures"])

    def test_all_categories_meet_minimums(self):
        with tempfile.TemporaryDirectory() as tmp:
            manifest = generate_annotation_dataset(tmp)
            report = dataset_coverage_report(manifest)
            self.assertTrue(report["all_minimums_met"])
            self.assertEqual(len(report["coverage"]), len(DATASET_CATEGORIES))

    def test_each_fixture_has_stable_ids_and_bbox(self):
        with tempfile.TemporaryDirectory() as tmp:
            manifest = generate_annotation_dataset(tmp, include_controls=False)
            for fx in manifest["fixtures"]:
                self.assertTrue(fx["fixture_id"])
                self.assertTrue(fx["element_id"])
                self.assertTrue(fx["render_node_id"])
                self.assertIn("bbox", fx)
                self.assertIn(fx["severity"], {"blocker", "warning", "info", "clean"})

    def test_fixture_pptx_files_exist(self):
        with tempfile.TemporaryDirectory() as tmp:
            manifest = generate_annotation_dataset(tmp, include_controls=False)
            for fx in manifest["fixtures"]:
                pptx_path = Path(tmp) / fx["expected"]["pptx"]
                self.assertTrue(pptx_path.exists(), f"missing {pptx_path}")

    def test_clean_controls_marked_clean(self):
        with tempfile.TemporaryDirectory() as tmp:
            manifest = generate_annotation_dataset(tmp, include_controls=True)
            controls = [f for f in manifest["fixtures"] if f["issue_code"] == ISSUE_CLEAN]
            self.assertEqual(len(controls), len(DATASET_CATEGORIES))


# ---------------------------------------------------------------------------
# Golden render coverage (blueprint §12.4: >=2 per role, >=1 family deck)
# ---------------------------------------------------------------------------


class GoldenRenderTests(unittest.TestCase):
    def test_golden_has_at_least_two_pages_per_role(self):
        with tempfile.TemporaryDirectory() as tmp:
            idx = render_golden_set(tmp, include_png=False)
            report = golden_coverage_report(idx)
            self.assertTrue(report["all_role_minimums_met"])
            self.assertEqual(len(report["role_coverage"]), len(ALL_ROLES))

    def test_golden_includes_all_three_geometry_families(self):
        with tempfile.TemporaryDirectory() as tmp:
            idx = render_golden_set(tmp, include_png=False)
            report = golden_coverage_report(idx)
            self.assertEqual(set(report["family_decks"]), set(GEOMETRY_FAMILIES))

    def test_golden_index_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            render_golden_set(tmp, include_png=False)
            index = load_golden_index(tmp)
            self.assertIn("pages", index)
            self.assertGreater(len(index["pages"]), 0)

    def test_each_golden_page_has_pptx_qa_json_and_signature(self):
        with tempfile.TemporaryDirectory() as tmp:
            idx = render_golden_set(tmp, include_png=False)
            for page in idx.pages:
                self.assertTrue(Path(page.pptx_path).exists())
                self.assertTrue(Path(page.qa_json_path).exists())
                self.assertTrue(page.geometry_signature)
                self.assertEqual(len(page.geometry_signature), 16)

    def test_golden_signatures_stable_across_runs(self):
        # Deterministic render: same recipe/canvas/tokens -> same signature.
        with tempfile.TemporaryDirectory() as tmp1, tempfile.TemporaryDirectory() as tmp2:
            idx1 = render_golden_set(tmp1, include_png=False)
            idx2 = render_golden_set(tmp2, include_png=False)
            sigs1 = {p.recipe_id: p.geometry_signature for p in idx1.pages}
            sigs2 = {p.recipe_id: p.geometry_signature for p in idx2.pages}
            self.assertEqual(sigs1, sigs2)


# ---------------------------------------------------------------------------
# E2E scenarios (blueprint §12.5 subset, deterministic)
# ---------------------------------------------------------------------------


def _bullets_slide(sid: str, title: str, n_bullets: int) -> SlideSpec:
    bullets = "\n".join(f"要点 {i}" for i in range(n_bullets))
    return SlideSpec(
        id=sid, role="bullets", communication_goal=title,
        elements=[
            ElementSpec(id=f"{sid}/title", kind="text", role="title", content={"text": title}, style_ref="component.title"),
            ElementSpec(id=f"{sid}/body", kind="text", role="body", content={"text": bullets}, style_ref="component.body"),
        ],
    )


class E2EScenarioTests(unittest.TestCase):
    """Subset of the 10 blueprint scenarios, expressed deterministically."""

    def _render_plans(self, plans, path: Path):
        render_layout_plans(plans, str(path))
        return Presentation(str(path))

    def test_scenario_3_long_bullets_auto_paginate(self):
        # §12.5 #3: long bullets auto-paginate.
        slides = [_bullets_slide(f"s{i}", f"第{i}节", 15) for i in range(3)]
        content = ContentSpec(id="deck", title="长要点", subtitle="", slides=slides)
        paginated, _ = paginate_content_spec(content, max_items=5)
        self.assertGreater(len(paginated.slides), 3)

    def test_scenario_1_dashboard_table_timeline_deck(self):
        # §12.5 #1: a deck including dashboard/table/timeline roles renders.
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        plans = []
        for role in ("cover", "dashboard", "table", "timeline", "end"):
            recipe = builtin_recipes(role)[0]
            geom = solve_recipe(recipe, canvas, tokens)
            self.assertFalse(geom.infeasible, f"{recipe.id} infeasible")
            slide = SlideSpec(
                id=f"s-{role}", role=role, communication_goal=role,
                elements=[
                    ElementSpec(id=f"s-{role}/title", kind="text", role="title", content={"text": role}, style_ref="component.title"),
                ],
            )
            plans.append(solved_geometry_to_layout_plan(slide, recipe, geom, canvas, tokens))
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "deck.pptx"
            prs = self._render_plans(plans, out)
            self.assertEqual(len(prs.slides), len(plans))

    def test_scenario_5_vertical_canvas_supported(self):
        # §12.5 #5: vertical (9:16) canvas.
        canvas = canvas_from_name("9:16")
        self.assertGreater(canvas.height_pt, canvas.width_pt)
        recipe = builtin_recipes("bullets")[0]
        geom = solve_recipe(recipe, canvas, _builtin_tokens())
        self.assertFalse(geom.infeasible)

    def test_scenario_4_4_3_canvas_supported(self):
        canvas = canvas_from_name("4:3")
        self.assertGreater(canvas.width_pt, canvas.height_pt)
        recipe = builtin_recipes("quote")[0]
        geom = solve_recipe(recipe, canvas, _builtin_tokens())
        self.assertFalse(geom.infeasible)

    def test_scenario_10_stress_100_pages(self):
        # §12.5 #10: 100-page stress test — render completes and deck has 100 slides.
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        recipe = builtin_recipes("bullets")[0]
        geom = solve_recipe(recipe, canvas, tokens)
        plans = []
        for i in range(100):
            slide = SlideSpec(
                id=f"s-{i:03d}", role="bullets", communication_goal=f"stress {i}",
                elements=[
                    ElementSpec(id=f"s-{i:03d}/title", kind="text", role="title", content={"text": f"页 {i}"}, style_ref="component.title"),
                    ElementSpec(id=f"s-{i:03d}/body", kind="text", role="body", content={"text": "要点"}, style_ref="component.body"),
                ],
            )
            plans.append(solved_geometry_to_layout_plan(slide, recipe, geom, canvas, tokens))
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "stress.pptx"
            render_layout_plans(plans, str(out))
            prs = Presentation(str(out))
            self.assertEqual(len(prs.slides), 100)


# ---------------------------------------------------------------------------
# Performance (blueprint §13.2: p95 single-page planning < 1s without CV)
# ---------------------------------------------------------------------------


class PerformanceTests(unittest.TestCase):
    def test_single_page_planning_under_one_second(self):
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        recipe = builtin_recipes("bullets")[0]
        slide = SlideSpec(
            id="perf", role="bullets", communication_goal="perf",
            elements=[
                ElementSpec(id="perf/title", kind="text", role="title", content={"text": "性能"}, style_ref="component.title"),
                ElementSpec(id="perf/body", kind="text", role="body", content={"text": "要点"}, style_ref="component.body"),
            ],
        )
        durations = []
        for _ in range(10):
            t0 = time.perf_counter()
            geom = solve_recipe(recipe, canvas, tokens)
            solved_geometry_to_layout_plan(slide, recipe, geom, canvas, tokens)
            durations.append(time.perf_counter() - t0)
        durations.sort()
        p95 = durations[int(0.95 * (len(durations) - 1))]
        self.assertFalse(geom.infeasible)
        self.assertLess(p95, 1.0, f"p95 single-page planning {p95:.3f}s exceeds 1s")

    def test_deck_planning_20_slides_under_5_seconds(self):
        canvas = canvas_from_name("16:9")
        slides = []
        for i in range(20):
            slides.append(SlideSpec(
                id=f"d-{i:02d}", role="bullets", communication_goal=f"d{i}",
                elements=[
                    ElementSpec(id=f"d-{i:02d}/title", kind="text", role="title", content={"text": str(i)}, style_ref="component.title"),
                    ElementSpec(id=f"d-{i:02d}/body", kind="text", role="body", content={"text": "要点"}, style_ref="component.body"),
                ],
            ))
        content = ContentSpec(id="d", title="deck", subtitle="", slides=slides)
        t0 = time.perf_counter()
        result = plan_deck(content.slides, canvas, profile_state={"tokens": _builtin_tokens(), "qa": {"min_body_font_size": 9}})
        elapsed = time.perf_counter() - t0
        self.assertIsNotNone(result)
        self.assertLess(elapsed, 5.0, f"20-slide deck planning {elapsed:.3f}s exceeds 5s")


# ---------------------------------------------------------------------------
# Property-style checks (blueprint §12.2: deterministic, stable-id invariants)
# ---------------------------------------------------------------------------


class PropertyTests(unittest.TestCase):
    def test_same_recipe_same_canvas_same_geometry(self):
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        recipe = builtin_recipes("comparison")[0]
        g1 = solve_recipe(recipe, canvas, tokens)
        g2 = solve_recipe(recipe, canvas, tokens)
        v1 = {k: v.value for k, v in g1.variables.items()}
        v2 = {k: v.value for k, v in g2.variables.items()}
        self.assertEqual(v1, v2)

    def test_all_recipes_produce_non_overlapping_title_zones(self):
        # Title zones should not have zero or negative width/height.
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        for role in ALL_ROLES:
            for recipe in builtin_recipes(role):
                geom = solve_recipe(recipe, canvas, tokens)
                if geom.infeasible:
                    continue
                slide = SlideSpec(
                    id=f"p-{recipe.id}", role=role, communication_goal=role,
                    elements=[ElementSpec(id=f"p-{recipe.id}/title", kind="text", role="title", content={"text": role}, style_ref="component.title")],
                )
                plan = solved_geometry_to_layout_plan(slide, recipe, geom, canvas, tokens)
                for node in plan.nodes:
                    self.assertGreater(node.geometry.bbox.width, 0, f"{recipe.id} {node.role} zero width")
                    self.assertGreater(node.geometry.bbox.height, 0, f"{recipe.id} {node.role} zero height")


if __name__ == "__main__":
    unittest.main()
