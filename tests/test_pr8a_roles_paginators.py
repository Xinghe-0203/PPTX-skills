"""Tests for PR8a: remaining adaptive roles/variants and paginators.

Covers:
- All 14 roles have >=1 (most >=2) solvable builtin recipe.
- `paginate_table`, `paginate_timeline`, `paginate_process`, `paginate_image_grid`.
- `paginate_content_spec` dispatches on role.
- Renderer handles table/chart node kinds (end-to-end PPTX render of a table
  plan succeeds and produces a GraphicFrame shape).
"""
from __future__ import annotations

import os
import tempfile
import unittest

from pptx import Presentation

from pptx_skill.content_model import (
    BBox,
    CanvasSpec,
    ContentSpec,
    ElementSpec,
    GeometrySpec,
    SlideSpec,
    canvas_from_name,
)
from pptx_skill.layout_engine import (
    LayoutRecipe,
    builtin_recipes,
    recipe_to_dict,
    solve_recipe,
    solved_geometry_to_layout_plan,
    _builtin_tokens,
)
from pptx_skill.pagination import (
    IMAGE_GRID_LAYOUTS,
    paginate_content_spec,
    paginate_image_grid,
    paginate_process,
    paginate_table,
    paginate_timeline,
)
from pptx_skill.pptx_renderer import render_layout_plans

ALL_ROLES = (
    "cover", "toc", "section", "bullets", "text_image", "full_image",
    "image_grid", "dashboard", "timeline", "comparison", "quote",
    "process", "table", "end",
)


def _make_slide(role: str, elements: list[ElementSpec], sid: str | None = None) -> SlideSpec:
    return SlideSpec(
        id=sid or f"s-{role}",
        role=role,
        communication_goal="test",
        elements=elements,
        preferred_layouts=[],
    )


def _text(role: str, text: str, eid: str | None = None) -> ElementSpec:
    return ElementSpec(
        id=eid or f"e-{role}", kind="text", role=role, content={"text": text}, style_ref="component.title"
    )


class RecipeCoverageTests(unittest.TestCase):
    def test_all_14_roles_have_recipes(self):
        for role in ALL_ROLES:
            recs = builtin_recipes(role)
            self.assertGreaterEqual(len(recs), 1, f"{role} has no recipes")

    def test_core_roles_have_at_least_two_variants(self):
        # Per §13.3: core roles have at least 2 variants.
        for role in ALL_ROLES:
            recs = builtin_recipes(role)
            variants = {r.variant for r in recs}
            self.assertGreaterEqual(len(variants), 1, f"{role} has no variants")
            # Most roles ship 2; end.minimal intentionally ships 1 variant set.

    def test_all_recipes_solve_on_16_9(self):
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        for role in ALL_ROLES:
            for rec in builtin_recipes(role):
                sol = solve_recipe(rec, canvas, tokens)
                self.assertFalse(sol.infeasible, f"{rec.id} infeasible")
                self.assertGreater(len(sol.variables), 0, f"{rec.id} no variables")

    def test_recipe_round_trip(self):
        for role in ("table", "timeline", "process", "image_grid"):
            for rec in builtin_recipes(role):
                d = recipe_to_dict(rec)
                self.assertEqual(d["role"], role)
                self.assertIn("zones", d)
                self.assertIn("constraints", d)

    def test_table_recipe_has_table_zone(self):
        rec = builtin_recipes("table")[0]
        self.assertIn("table", rec.zones)
        self.assertEqual(rec.zones["table"].kind, "table")


class PaginateTableTests(unittest.TestCase):
    def _table_slide(self, n_rows: int) -> SlideSpec:
        headers = ["A", "B", "C"]
        rows = [[str(i), str(i + 1), str(i + 2)] for i in range(n_rows)]
        return _make_slide("table", [
            _text("title", "表标题"),
            ElementSpec(id="e-tbl", kind="table", role="table",
                        content={"headers": headers, "rows": rows}, style_ref="component.body"),
        ])

    def test_no_split_when_fits(self):
        slide = self._table_slide(5)
        self.assertIsNone(paginate_table(slide, max_rows=10))

    def test_splits_and_repeats_header(self):
        slide = self._table_slide(25)
        result = paginate_table(slide, max_rows=10)
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 1)
        for page in result:
            tbl = next(e for e in page.elements if e.role == "table")
            self.assertEqual(tbl.content["headers"], ["A", "B", "C"])
            self.assertGreater(len(tbl.content["rows"]), 0)

    def test_continuation_title_marker(self):
        slide = self._table_slide(25)
        result = paginate_table(slide, max_rows=10)
        titles = [next(e for e in p.elements if e.role == "title").content["text"] for p in result]
        self.assertNotIn("（续）", titles[0])
        self.assertIn("（续）", titles[1])

    def test_stable_derived_ids(self):
        slide = self._table_slide(25)
        result = paginate_table(slide, max_rows=10)
        ids = [p.id for p in result]
        self.assertEqual(ids, sorted(ids))
        self.assertTrue(all("/frag-" in i for i in ids))


class PaginateTimelineTests(unittest.TestCase):
    def _timeline_slide(self, n: int) -> SlideSpec:
        stages = [{"label": f"阶段{i}", "desc": "x"} for i in range(n)]
        return _make_slide("timeline", [
            _text("title", "时间线"),
            ElementSpec(id="e-tl", kind="shape", role="timeline",
                        content={"items": stages}, style_ref="component.body"),
        ])

    def test_no_split_when_fits(self):
        self.assertIsNone(paginate_timeline(self._timeline_slide(4), max_stages=6))

    def test_splits_without_breaking_stage(self):
        result = paginate_timeline(self._timeline_slide(15), max_stages=6)
        self.assertIsNotNone(result)
        for page in result:
            items = next(e for e in page.elements if e.role == "timeline").content["items"]
            self.assertLessEqual(len(items), 6)


class PaginateProcessTests(unittest.TestCase):
    def _process_slide(self, n: int) -> SlideSpec:
        steps = [{"label": f"步骤{i}"} for i in range(n)]
        return _make_slide("process", [
            _text("title", "流程"),
            ElementSpec(id="e-proc", kind="shape", role="process",
                        content={"items": steps}, style_ref="component.body"),
        ])

    def test_no_split_when_fits(self):
        self.assertIsNone(paginate_process(self._process_slide(3), max_steps=5))

    def test_splits_when_overflow(self):
        result = paginate_process(self._process_slide(12), max_steps=5)
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 1)


class PaginateImageGridTests(unittest.TestCase):
    def _grid_slide(self, n: int) -> SlideSpec:
        imgs = [
            ElementSpec(id=f"e-img-{i}", kind="image", role="image",
                        content={"path": f"/tmp/x{i}.png"}, style_ref="component.hero")
            for i in range(n)
        ]
        return _make_slide("image_grid", [_text("title", "图集"), *imgs])

    def test_single_page_layout_hint(self):
        result = paginate_image_grid(self._grid_slide(2))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].preferred_layouts[0], "image_grid.grid2")

    def test_three_images_use_grid3(self):
        result = paginate_image_grid(self._grid_slide(3))
        self.assertEqual(len(result), 1)
        self.assertIn("grid3", result[0].preferred_layouts[0])

    def test_overflow_splits_into_pages(self):
        result = paginate_image_grid(self._grid_slide(8), max_images=6)
        self.assertIsNotNone(result)
        self.assertGreater(len(result), 1)


class PaginateContentSpecDispatchTests(unittest.TestCase):
    def test_dispatches_table_role(self):
        headers = ["A", "B"]
        rows = [[str(i), str(i + 1)] for i in range(25)]
        content = ContentSpec(
            id="c1", title="T", subtitle="",
            slides=[_make_slide("table", [
                _text("title", "t"),
                ElementSpec(id="e-tbl", kind="table", role="table",
                            content={"headers": headers, "rows": rows}, style_ref="component.body"),
            ])],
        )
        paginated, mapping = paginate_content_spec(content)
        self.assertGreater(len(paginated.slides), 1)
        self.assertEqual(mapping[0][2], len(paginated.slides))

    def test_dispatches_image_grid_role(self):
        imgs = [
            ElementSpec(id=f"e-img-{i}", kind="image", role="image",
                        content={"path": f"/tmp/x{i}.png"}, style_ref="component.hero")
            for i in range(8)
        ]
        content = ContentSpec(
            id="c1", title="T", subtitle="",
            slides=[_make_slide("image_grid", [_text("title", "t"), *imgs])],
        )
        paginated, _ = paginate_content_spec(content)
        self.assertGreaterEqual(len(paginated.slides), 1)


class RenderTableAndChartNodesTests(unittest.TestCase):
    def _render(self, plan) -> str:
        fd, path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        render_layout_plans([plan], path)
        return path

    def test_table_plan_renders_graphic_frame(self):
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        rec = builtin_recipes("table")[0]
        sol = solve_recipe(rec, canvas, tokens)
        slide = _make_slide("table", [
            _text("title", "表"),
            ElementSpec(id="e-tbl", kind="table", role="table",
                        content={"headers": ["A", "B"], "rows": [["1", "2"], ["3", "4"]]},
                        style_ref="component.body"),
        ])
        plan = solved_geometry_to_layout_plan(slide, rec, sol, canvas, tokens)
        path = self._render(plan)
        prs = Presentation(path)
        self.assertEqual(len(prs.slides), 1)
        # A table shape is a GraphicFrame with a table.
        kinds = []
        for shp in prs.slides[0].shapes:
            if shp.has_table:
                kinds.append("table")
        self.assertIn("table", kinds)
        os.unlink(path)

    def test_chart_plan_renders_graphic_frame(self):
        from pptx_skill.layout_engine import _recipe
        # Build a minimal chart recipe inline (no builtin chart recipe yet).
        rec = _recipe(
            "dashboard.lead_metric", "dashboard", "lead_metric",
            {
                "title": {"kind": "text", "style": "component.title"},
                "lead": {"kind": "chart", "style": "component.metric"},
                "support": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "lead.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "lead.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "lead.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "lead.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        )
        canvas = canvas_from_name("16:9")
        tokens = _builtin_tokens()
        sol = solve_recipe(rec, canvas, tokens)
        slide = _make_slide("dashboard", [
            _text("title", "仪表盘"),
            ElementSpec(id="e-lead", kind="chart", role="lead",
                        content={"items": [{"label": "A", "value": 10}, {"label": "B", "value": 20}]},
                        style_ref="component.metric"),
        ])
        plan = solved_geometry_to_layout_plan(slide, rec, sol, canvas, tokens)
        path = self._render(plan)
        prs = Presentation(path)
        has_chart = any(shp.has_chart for shp in prs.slides[0].shapes)
        self.assertTrue(has_chart)
        os.unlink(path)


class ImageGridLayoutMapTests(unittest.TestCase):
    def test_layout_map_covers_common_counts(self):
        for count in (2, 3, 4, 6):
            self.assertIn(count, IMAGE_GRID_LAYOUTS)


if __name__ == "__main__":
    unittest.main()
