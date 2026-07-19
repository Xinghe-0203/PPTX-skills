"""Tests for rendered QA, repair engine, and generation pipeline (PR5)."""
from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from PIL import Image

from pptx_skill.content_model import CanvasSpec, ContentSpec, ElementSpec, SlideSpec
from pptx_skill.generation_pipeline import (
    GenerationResult,
    LayoutPlanningError,
    PresentationQualityError,
    plan_deck_layouts,
    run_generation_pipeline,
)
from pptx_skill.layout_engine import _builtin_tokens
from pptx_skill.repair_engine import (
    RepairAction,
    RepairActionKind,
    apply_repairs,
    merge_profile_overrides,
    propose_repairs,
)
from pptx_skill.render_qa import (
    RenderQAConfig,
    compare_slide_to_baseline,
    evaluate_render_against_baseline,
)
from pptx_skill.semantic_qa import DetectedIssue, IssueKind, IssueSeverity, SemanticQAEngine
from pptx_skill.visual_qa import CheckOutcome, QAIssue, QAReport, Severity


class RenderQATests(unittest.TestCase):
    def _make_png(self, width: int, height: int, color: tuple[int, int, int], path: str) -> str:
        img = Image.new("RGB", (width, height), color)
        img.save(path)
        return path

    def test_identical_images_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = self._make_png(160, 90, (255, 255, 255), os.path.join(tmp, "a.png"))
            b = self._make_png(160, 90, (255, 255, 255), os.path.join(tmp, "b.png"))
            diff = compare_slide_to_baseline(0, a, b, output_dir=tmp)
            self.assertAlmostEqual(diff.perceptual_diff_ratio, 0.0, places=3)
            self.assertEqual(diff.window_diff_density, 0.0)

    def test_different_images_fail_window_density(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = self._make_png(160, 90, (255, 255, 255), os.path.join(tmp, "a.png"))
            b = self._make_png(160, 90, (0, 0, 0), os.path.join(tmp, "b.png"))
            diff = compare_slide_to_baseline(0, a, b, output_dir=tmp)
            self.assertGreater(diff.window_diff_density, 0.0)
            self.assertIsNotNone(diff.diff_png_path)

    def test_evaluate_with_count_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            pngs = [self._make_png(160, 90, (255, 255, 255), os.path.join(tmp, f"s{i}.png")) for i in range(2)]
            baselines = [pngs[0]]
            report = evaluate_render_against_baseline(pngs, baselines)
            self.assertEqual(report.status, CheckOutcome.FAIL)
            self.assertTrue(any(i.code == "RENDER_BASELINE_COUNT_MISMATCH" for i in report.issues))

    def test_evaluate_same_baseline_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            pngs = [self._make_png(160, 90, (255, 255, 255), os.path.join(tmp, "s0.png"))]
            report = evaluate_render_against_baseline(pngs, list(pngs), output_dir=tmp, config=RenderQAConfig())
            self.assertEqual(report.status, CheckOutcome.PASS)


class RepairEngineTests(unittest.TestCase):
    def _make_slide(self, role: str = "bullets", body_text: str = "short") -> SlideSpec:
        return SlideSpec(
            id="s1",
            role=role,
            communication_goal="test",
            elements=[
                ElementSpec(id="title", kind="text", role="title", content={"text": "Title"}, style_ref="title"),
                ElementSpec(id="body", kind="text", role="body", content={"text": body_text}, style_ref="body"),
            ],
            preferred_layouts=["bullets.rail", "bullets.wide"],
        )

    def test_propose_reduce_font_for_overflow(self):
        # Build a plan manually with a small box to force overflow.
        from pptx_skill.content_model import BBox, GeometrySpec, LayoutPlan, PlannedNode, CanvasSpec
        from pptx_skill.semantic_qa import SemanticQAEngine, IssueKind
        node = PlannedNode(
            id="s1/body",
            element_id="body",
            recipe_node_id="body",
            kind="text",
            role="body",
            geometry=GeometrySpec(BBox(36, 120, 100, 30)),
            resolved_style={"size": 18, "font_family": "Microsoft YaHei"},
            content_binding={"text": "x" * 5000},
            z_order=1,
        )
        plan = LayoutPlan(canvas=CanvasSpec(959.976, 540), recipe_id="test", nodes=[node], local_score=0.0)
        report = SemanticQAEngine().check(plan)
        self.assertTrue(any(i.kind == IssueKind.TEXT_OVERFLOW for i in report.issues))
        slide = self._make_slide(body_text="x" * 5000)
        actions = propose_repairs(report, [slide], {"qa": {"min_body_font_size": 9}})
        self.assertTrue(any(a.kind == RepairActionKind.REDUCE_FONT_WITHIN_LIMIT for a in actions))

    def test_apply_reduce_font(self):
        from pptx_skill.content_model import BBox, GeometrySpec, LayoutPlan, PlannedNode, CanvasSpec
        from pptx_skill.semantic_qa import SemanticQAEngine
        node = PlannedNode(
            id="s1/body",
            element_id="body",
            recipe_node_id="body",
            kind="text",
            role="body",
            geometry=GeometrySpec(BBox(36, 120, 100, 30)),
            resolved_style={"size": 18, "font_family": "Microsoft YaHei"},
            content_binding={"text": "x" * 5000},
            z_order=1,
        )
        plan = LayoutPlan(canvas=CanvasSpec(959.976, 540), recipe_id="test", nodes=[node], local_score=0.0)
        profile = {"qa": {"min_body_font_size": 9}}
        report = SemanticQAEngine().check(plan)
        slide = self._make_slide(body_text="x" * 5000)
        actions = propose_repairs(report, [slide], profile)
        new_slides, overrides = apply_repairs(actions, [slide], profile)
        body = new_slides[0].elements[1]
        self.assertLess(body.content.get("font_size_pt", 18), 18)

    def test_merge_profile_overrides(self):
        profile = {"tokens": {"a": 1}}
        merged = merge_profile_overrides(profile, {"b": 2})
        self.assertEqual(merged["tokens"]["a"], 1)
        self.assertIn("repair_overrides", merged)

    def test_propose_switch_layout_for_empty_overflow_element(self):
        report = QAReport(
            status=CheckOutcome.FAIL,
            issues=[
                QAIssue(
                    code="TEXT_OVERFLOW_ESTIMATED",
                    severity=Severity.BLOCKER,
                    slide_index=0,
                    element_id="missing",
                    message="overflow",
                )
            ],
        )
        slide = self._make_slide()
        actions = propose_repairs(report, [slide], {})
        self.assertTrue(any(a.kind == RepairActionKind.SWITCH_LAYOUT_CANDIDATE for a in actions))


class GenerationPipelineTests(unittest.TestCase):
    def _make_content(self, body_text: str = "Short body") -> ContentSpec:
        return ContentSpec(
            id="deck-1",
            title="Test Deck",
            subtitle="",
            slides=[
                SlideSpec(
                    id="s1",
                    role="bullets",
                    communication_goal="test",
                    elements=[
                        ElementSpec(id="title", kind="text", role="title", content={"text": "Title"}, style_ref="title"),
                        ElementSpec(id="body", kind="text", role="body", content={"text": body_text}, style_ref="body"),
                    ],
                )
            ],
        )

    def test_plan_deck_layouts(self):
        content = self._make_content()
        profile = {"tokens": _builtin_tokens()}
        plans = plan_deck_layouts(content.slides, CanvasSpec(959.976, 540), profile)
        self.assertEqual(len(plans), 1)
        self.assertGreater(len(plans[0].nodes), 0)

    def test_pipeline_report_mode_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "deck.pptx")
            content = self._make_content()
            result = run_generation_pipeline(content, out, qa_mode="report")
            self.assertIsInstance(result, GenerationResult)
            self.assertIn(result.qa_status, {"pass", "inconclusive"})
            self.assertTrue(os.path.exists(out))

    def test_pipeline_strict_mode_raises_on_unrepairable(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "deck.pptx")
            # Patch plan_deck_layouts to return a plan with an overflow node.
            from pptx_skill.content_model import BBox, GeometrySpec, PlannedNode
            from pptx_skill.semantic_qa import SemanticQAEngine
            import pptx_skill.generation_pipeline as gp
            content = self._make_content(body_text="Short body")
            profile = {"tokens": _builtin_tokens()}
            bad_plans = plan_deck_layouts(content.slides, CanvasSpec(959.976, 540), profile)
            overflow_node = PlannedNode(
                id="s1/body",
                element_id="body",
                recipe_node_id="body",
                kind="text",
                role="body",
                geometry=GeometrySpec(BBox(36, 120, 100, 30)),
                resolved_style={"size": 18, "font_family": "Microsoft YaHei"},
                content_binding={"text": "x" * 5000},
                z_order=1,
            )
            bad_plans[0].nodes.append(overflow_node)
            self.assertFalse(SemanticQAEngine().check(bad_plans[0]).passed)
            orig = gp.plan_deck_layouts
            gp.plan_deck_layouts = lambda slides, canvas, profile: bad_plans
            try:
                with self.assertRaises(PresentationQualityError):
                    run_generation_pipeline(content, out, qa_mode="strict", max_repair_passes=0)
            finally:
                gp.plan_deck_layouts = orig

    def test_pipeline_off_mode_no_preview(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "deck.pptx")
            content = self._make_content()
            result = run_generation_pipeline(content, out, qa_mode="off")
            self.assertEqual(result.qa_status, "pass")
            self.assertIsNone(result.preview_dir)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _plan_for_slide(slide: SlideSpec):
    from pptx_skill.layout_engine import builtin_recipes, plan_slide_candidates, solved_geometry_to_layout_plan
    from pptx_skill.layout_engine import _builtin_tokens
    from pptx_skill.content_model import CanvasSpec, SafeInsets
    canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
    tokens = _builtin_tokens()
    recipes = builtin_recipes(slide.role)
    candidates, diagnostics = plan_slide_candidates(slide, recipes, canvas, tokens, max_candidates=1)
    if not candidates:
        return None
    bundle = candidates[0]
    return solved_geometry_to_layout_plan(slide, bundle.recipe, bundle.geometry, canvas, tokens)


if __name__ == "__main__":
    unittest.main()
