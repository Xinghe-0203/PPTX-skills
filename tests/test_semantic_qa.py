"""Tests for semantic QA engine."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from PIL import Image

from pptx_skill.content_model import BBox, CanvasSpec, GeometrySpec, LayoutPlan, PlannedNode
from pptx_skill.semantic_qa import (
    IssueKind,
    IssueSeverity,
    SemanticQAEngine,
    check_layout_plan,
    contrast_ratio,
)


class SemanticQATests(unittest.TestCase):
    def _make_plan(self, nodes: list[PlannedNode]) -> LayoutPlan:
        return LayoutPlan(
            canvas=CanvasSpec(960, 540),
            recipe_id="test.qa",
            nodes=nodes,
            local_score=0.0,
        )

    def _text_node(self, text: str, bbox: BBox, size: float = 16, color: str = "#1A1A1A", fill: str | None = None) -> PlannedNode:
        return PlannedNode(
            id="text",
            element_id="text",
            recipe_node_id="text",
            kind="text",
            role="body",
            geometry=GeometrySpec(bbox),
            resolved_style={"size": size, "color": color, "fill": fill},
            content_binding={"text": text},
            z_order=1,
        )

    def test_contrast_ratio_calculation(self):
        self.assertAlmostEqual(contrast_ratio("#000000", "#FFFFFF"), 21.0, places=1)
        self.assertGreater(contrast_ratio("#FFFFFF", "#000000"), 20.0)

    def test_text_overflow_detected(self):
        long_text = "A" * 800
        plan = self._make_plan([self._text_node(long_text, BBox(100, 100, 80, 40), size=16)])
        report = check_layout_plan(plan)
        overflow = [i for i in report.issues if i.kind == IssueKind.TEXT_OVERFLOW]
        self.assertTrue(overflow)
        self.assertEqual(overflow[0].severity, IssueSeverity.BLOCKER)

    def test_overlap_detected(self):
        plan = self._make_plan([
            PlannedNode(
                id="a",
                element_id="a",
                recipe_node_id="a",
                kind="shape",
                role="decoration",
                geometry=GeometrySpec(BBox(100, 100, 100, 100)),
                resolved_style={"fill": "#FF0000"},
                content_binding={},
                z_order=0,
            ),
            PlannedNode(
                id="b",
                element_id="b",
                recipe_node_id="b",
                kind="shape",
                role="body",
                geometry=GeometrySpec(BBox(150, 150, 100, 100)),
                resolved_style={"fill": "#0000FF"},
                content_binding={},
                z_order=1,
            ),
        ])
        report = check_layout_plan(plan)
        overlaps = [i for i in report.issues if i.kind == IssueKind.OVERLAP]
        self.assertTrue(overlaps)
        self.assertEqual(overlaps[0].severity, IssueSeverity.BLOCKER)

    def test_low_contrast_warning(self):
        plan = self._make_plan([self._text_node("Hello", BBox(100, 100, 200, 60), size=24, color="#DDDDDD", fill="#FFFFFF")])
        report = check_layout_plan(plan)
        contrast = [i for i in report.issues if i.kind == IssueKind.LOW_CONTRAST]
        self.assertTrue(contrast)
        self.assertEqual(contrast[0].severity, IssueSeverity.WARNING)

    def test_missing_font_warning(self):
        plan = self._make_plan([
            self._text_node("x", BBox(100, 100, 50, 50), size=12)
        ])
        # Use an empty index so the resolver cannot find any fallback.
        from pptx_skill import text_metrics
        original_index = text_metrics._FONT_INDEX
        text_metrics._FONT_INDEX = {}
        try:
            plan.nodes[0].resolved_style["font_family"] = "DefinitelyNotARealFontXYZ123"
            report = SemanticQAEngine().check(plan)
            missing = [i for i in report.issues if i.kind == IssueKind.MISSING_FONT]
            self.assertTrue(missing)
            self.assertEqual(missing[0].severity, IssueSeverity.WARNING)
        finally:
            text_metrics._FONT_INDEX = original_index

    def test_image_distortion_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            img_path = Path(tmp) / "wide.png"
            Image.new("RGB", (800, 400), color=(0, 128, 0)).save(img_path)
            plan = self._make_plan([
                PlannedNode(
                    id="img",
                    element_id="img",
                    recipe_node_id="img",
                    kind="image",
                    role="hero",
                    geometry=GeometrySpec(BBox(100, 100, 100, 200)),
                    resolved_style={},
                    content_binding={"path": str(img_path)},
                    z_order=0,
                )
            ])
            report = check_layout_plan(plan)
            dist = [i for i in report.issues if i.kind == IssueKind.IMAGE_DISTORTION]
            self.assertTrue(dist)
            self.assertEqual(dist[0].severity, IssueSeverity.WARNING)

    def test_missing_image_blocker(self):
        plan = self._make_plan([
            PlannedNode(
                id="img",
                element_id="img",
                recipe_node_id="img",
                kind="image",
                role="hero",
                geometry=GeometrySpec(BBox(100, 100, 100, 200)),
                resolved_style={},
                content_binding={"path": "/nonexistent/image.png"},
                z_order=0,
            )
        ])
        report = check_layout_plan(plan)
        self.assertFalse(report.passed)
        missing = [i for i in report.issues if i.kind == IssueKind.EMPTY_CONTENT]
        self.assertTrue(missing)

    def test_valid_plan_passes(self):
        plan = self._make_plan([self._text_node("Short text", BBox(100, 100, 400, 100), size=18)])
        report = check_layout_plan(plan)
        blockers = [i for i in report.issues if i.severity == IssueSeverity.BLOCKER]
        self.assertFalse(blockers)
        self.assertTrue(report.passed)


if __name__ == "__main__":
    unittest.main()
