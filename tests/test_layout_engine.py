"""Tests for the constraint-based adaptive layout engine."""
from __future__ import annotations

import math
import unittest

from pptx_skill.content_model import (
    BBox,
    CanvasSpec,
    ElementSpec,
    SafeInsets,
    SlideSpec,
)
from pptx_skill.layout_engine import (
    Constraint,
    LayoutRecipe,
    Term,
    _builtin_tokens,
    builtin_recipes,
    plan_slide_candidates,
    recipe_from_dict,
    recipe_to_dict,
    solve_recipe,
)


class LayoutEngineTests(unittest.TestCase):
    def _make_slide(self, role: str, elements: list[ElementSpec]) -> SlideSpec:
        return SlideSpec(
            id=f"s/{role}",
            role=role,
            communication_goal="test",
            elements=elements,
        )

    def test_recipe_round_trip(self):
        recipe = builtin_recipes("bullets")[0]
        data = recipe_to_dict(recipe)
        restored = recipe_from_dict(data)
        self.assertEqual(restored.id, recipe.id)
        self.assertEqual(len(restored.constraints), len(recipe.constraints))

    def test_invalid_constraint_op_rejected(self):
        with self.assertRaises(ValueError):
            recipe_from_dict({
                "id": "bad",
                "role": "bullets",
                "variant": "bad",
                "zones": {"body": {"kind": "text", "style": "body"}},
                "constraints": [{"terms": [], "op": "**", "rhs": {"const": 0}}],
            })

    def test_solve_bullets_recipe(self):
        recipe = builtin_recipes("bullets")[0]
        canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
        tokens = _builtin_tokens()
        geometry = solve_recipe(recipe, canvas, tokens)
        self.assertFalse(geometry.infeasible)
        self.assertIn("title.left", geometry.variables)
        # title should be within safe area
        title_left = geometry.variables["title.left"].value
        title_right = geometry.variables["title.right"].value
        self.assertGreaterEqual(title_left, canvas.safe.left)
        self.assertLessEqual(title_right, canvas.width_pt - canvas.safe.right)

    def test_solve_text_image_recipe(self):
        recipe = builtin_recipes("text_image")[0]
        canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
        tokens = _builtin_tokens()
        geometry = solve_recipe(recipe, canvas, tokens)
        self.assertFalse(geometry.infeasible)
        body_right = geometry.variables["body.right"].value
        hero_left = geometry.variables["hero.left"].value
        self.assertGreaterEqual(hero_left, body_right)

    def test_infeasible_constraint(self):
        recipe = LayoutRecipe(
            id="impossible",
            role="bullets",
            variant="impossible",
            zones={"body": {"kind": "text", "style": "body"}},
            constraints=[
                Constraint(terms=[Term("body.left")], op="==", rhs={"const": 0}),
                Constraint(terms=[Term("body.left")], op="==", rhs={"const": 500}, strength="required"),
            ],
        )
        canvas = CanvasSpec(959.976, 540)
        geometry = solve_recipe(recipe, canvas, {})
        self.assertTrue(geometry.infeasible)

    def test_plan_slide_candidates_topk(self):
        recipes = builtin_recipes("bullets")
        canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
        tokens = _builtin_tokens()
        elements = [
            ElementSpec(id="title", kind="text", role="title", content={"text": "Title"}, style_ref="title"),
            ElementSpec(id="body", kind="text", role="body", content={"text": "Short body."}, style_ref="body"),
        ]
        slide = self._make_slide("bullets", elements)
        candidates, diagnostics = plan_slide_candidates(slide, recipes, canvas, tokens, max_candidates=2)
        self.assertLessEqual(len(candidates), 2)
        if candidates:
            self.assertEqual(candidates[0].blocker_count, 0)

    def test_dashboard_recipe_solves(self):
        recipes = builtin_recipes("dashboard")
        canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
        tokens = _builtin_tokens()
        elements = [
            ElementSpec(id="title", kind="text", role="title", content={"text": "Dashboard"}, style_ref="title"),
            ElementSpec(id="lead", kind="text", role="lead", content={"text": "99.9%"}, style_ref="metric"),
            ElementSpec(id="support", kind="text", role="support", content={"text": "Details"}, style_ref="body"),
        ]
        slide = self._make_slide("dashboard", elements)
        candidates, _ = plan_slide_candidates(slide, recipes, canvas, tokens)
        self.assertTrue(len(candidates) >= 1)
        for c in candidates:
            self.assertEqual(c.blocker_count, 0)

    def test_variable_values_are_finite(self):
        recipe = builtin_recipes("text_image")[0]
        canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
        tokens = _builtin_tokens()
        geometry = solve_recipe(recipe, canvas, tokens)
        for var in geometry.variables.values():
            self.assertTrue(math.isfinite(var.value))

    def test_body_zone_within_safe_area(self):
        recipe = builtin_recipes("bullets")[0]
        canvas = CanvasSpec(959.976, 540, safe=SafeInsets(36, 48, 32, 48))
        tokens = _builtin_tokens()
        geometry = solve_recipe(recipe, canvas, tokens)
        body = geometry.variables
        # body.top comes from title.bottom + gap; allow small solver slack.
        self.assertGreaterEqual(body["body.left"].value, canvas.safe.left - 2)
        self.assertLessEqual(body["body.right"].value, canvas.width_pt - canvas.safe.right + 2)
        self.assertGreaterEqual(body["body.top"].value, canvas.safe.top - 2)
        self.assertLessEqual(body["body.bottom"].value, canvas.height_pt - canvas.safe.bottom + 2)


if __name__ == "__main__":
    unittest.main()
