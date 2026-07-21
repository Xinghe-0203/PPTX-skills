"""Visual-rebuild adapter (PR8c).

Per blueprint §9.3 the goal for ``visual-rebuild`` is: parse the reference
deck's shapes/images, synthesize stable :class:`ElementSpec`/:class:`SlideSpec`
and pick a recipe, then hand the plan to the **adaptive renderer** — not the
clone branch. Output carries an editable render trace plus a reference-diff
report bounded by a similarity budget.

This module is the independent rebuild path. ``reference_adapter`` routed
``visual-rebuild`` to clone as an alias pending this PR; after PR8c the alias
is replaced by a real adaptive rebuild.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx_skill.content_model import (
    CanvasSpec,
    ElementSpec,
    LayoutPlan,
    SlideSpec,
)
from pptx_skill.layout_engine import (
    _builtin_tokens,
    builtin_recipes,
    solve_recipe,
    solved_geometry_to_layout_plan,
)
from pptx_skill.pptx_renderer import RenderResult, render_layout_plans

try:
    from reference_ppt import analyze_presentation  # type: ignore[import-not-found]
except Exception:  # pragma: no cover
    analyze_presentation = None  # type: ignore[assignment]


# Role -> recipe variant preference for the rebuild. Mirrors the 14-role
# coverage established in PR8a; the first variant is the default pick.
REBUILD_RECIPE_PREFERENCE: dict[str, str] = {
    "cover": "cover.axis",
    "toc": "toc.list",
    "section": "section.full",
    "bullets": "bullets.rail",
    "text_image": "text_image.asymmetric_left",
    "full_image": "full_image.full",
    "image_grid": "image_grid.grid3",
    "dashboard": "dashboard.lead_metric",
    "timeline": "timeline.horizontal",
    "comparison": "comparison.two_column",
    "quote": "quote.centered",
    "process": "process.horizontal",
    "table": "table.standard",
    "end": "end.thanks",
}

# Default similarity budget for the reference diff. Below this, the rebuild is
# considered too divergent and a warning is surfaced (never a silent pass).
DEFAULT_SIMILARITY_BUDGET = 0.45


@dataclass
class SlideRebuildPlan:
    """The rebuilt plan for one source slide."""

    source_slide_index: int
    detected_role: str
    chosen_recipe_id: str
    slide_spec: SlideSpec
    layout_plan: LayoutPlan
    diagnostics: dict[str, Any] = field(default_factory=dict)


@dataclass
class ReferenceDiffRecord:
    """Per-slide similarity between the rebuilt output and the reference."""

    slide_index: int
    role: str
    geometry_similarity: float
    within_budget: bool
    notes: list[str] = field(default_factory=list)


@dataclass
class VisualRebuildResult:
    output_path: str
    render_result: RenderResult
    canvas: CanvasSpec
    plans: list[SlideRebuildPlan] = field(default_factory=list)
    diff: list[ReferenceDiffRecord] = field(default_factory=list)
    mean_similarity: float = 0.0
    within_budget: bool = True
    diagnostics: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": self.output_path,
            "canvas": {"width_pt": self.canvas.width_pt, "height_pt": self.canvas.height_pt},
            "slides_rebuilt": len(self.plans),
            "mean_similarity": self.mean_similarity,
            "within_budget": self.within_budget,
            "diff": [
                {
                    "slide_index": d.slide_index,
                    "role": d.role,
                    "geometry_similarity": d.geometry_similarity,
                    "within_budget": d.within_budget,
                    "notes": d.notes,
                }
                for d in self.diff
            ],
            "diagnostics": self.diagnostics,
        }


def _shape_to_element(shape_record: dict, slide_id: str, role: str) -> ElementSpec | None:
    """Promote a reference shape record into a stable ElementSpec.

    Returns ``None`` for purely decorative shapes (no text, no image, no table).
    """
    shape_type = shape_record.get("shape_type", "")
    text = shape_record.get("text", "") or ""
    has_table = shape_record.get("has_table", False)
    has_chart = shape_record.get("has_chart", False)
    shape_id = shape_record.get("shape_id")
    eid = f"{slide_id}/shape-{shape_id}"

    if has_table:
        return ElementSpec(
            id=eid, kind="table", role="table",
            content={"headers": [], "rows": [], "placeholder": True},
            style_ref="component.body",
        )
    if has_chart:
        return ElementSpec(
            id=eid, kind="chart", role="lead",
            content={"items": [], "placeholder": True},
            style_ref="component.metric",
        )
    if shape_type == "PICTURE":
        return ElementSpec(
            id=eid, kind="image", role="hero",
            content={"path": "", "placeholder": True},
            style_ref="component.hero",
        )
    if text:
        # Heuristic role within the slide: largest font text -> title.
        sizes = shape_record.get("font_sizes") or []
        max_size = max(sizes) if sizes else 0
        is_title = max_size >= 28 or shape_record.get("placeholder_type") in {"TITLE", "CENTER_TITLE"}
        return ElementSpec(
            id=eid, kind="text", role="title" if is_title else "body",
            content={"text": text}, style_ref="component.title" if is_title else "component.body",
        )
    return None


def _shapes_to_slide_spec(
    slide_record: dict, index: int, role: str
) -> SlideSpec:
    """Build a SlideSpec from a reference slide's shape records."""
    sid = f"rebuild/slide-{index}"
    elements: list[ElementSpec] = []
    for shape in slide_record.get("shapes", []):
        elem = _shape_to_element(shape, sid, role)
        if elem is not None:
            elements.append(elem)
    # Guarantee a title element exists so recipes have something to bind.
    if not any(e.role == "title" for e in elements):
        title_text = slide_record.get("layout", "") or f"Slide {index}"
        elements.insert(0, ElementSpec(
            id=f"{sid}/title", kind="text", role="title",
            content={"text": title_text}, style_ref="component.title",
        ))
    return SlideSpec(
        id=sid,
        role=role,
        communication_goal=slide_record.get("layout", role),
        elements=elements,
        preferred_layouts=[REBUILD_RECIPE_PREFERENCE.get(role, "bullets.rail")],
    )


def _pick_recipe(role: str):
    """Choose the recipe for a role, preferring the named variant."""
    recipes = builtin_recipes(role)
    if not recipes:
        return None
    preferred_id = REBUILD_RECIPE_PREFERENCE.get(role)
    if preferred_id:
        for r in recipes:
            if r.id == preferred_id:
                return r
    return recipes[0]


def _bbox_iou(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    """IoU between two (x, y, w, h) fractional boxes."""
    ax1, ay1, aw, ah = a
    ax2, ay2 = ax1 + aw, ay1 + ah
    bx1, by1, bw, bh = b
    bx2, by2 = bx1 + bw, by1 + bh
    inter_x1 = max(ax1, bx1)
    inter_y1 = max(ay1, by1)
    inter_x2 = min(ax2, bx2)
    inter_y2 = min(ay2, by2)
    inter_w = max(0.0, inter_x2 - inter_x1)
    inter_h = max(0.0, inter_y2 - inter_y1)
    inter = inter_w * inter_h
    union = aw * ah + bw * bh - inter
    return inter / union if union > 0 else 0.0


def _slide_geometry_similarity(
    reference_shapes: list[dict], plan: Any
) -> float:
    """Mean IoU between reference shape bboxes and the rebuilt plan's node bboxes.

    Both are expressed as fractional (x, y, w, h) of the canvas so the comparison
    is resolution-independent. A low score means the rebuild diverged from the
    reference geometry — expected, since the adaptive renderer re-solves the
    recipe rather than cloning positions.
    """
    ref_boxes = [(s["x"], s["y"], s["w"], s["h"]) for s in reference_shapes
                 if all(k in s for k in ("x", "y", "w", "h"))
                 and (s.get("shape_type") != "PICTURE" or s.get("text"))]
    if not ref_boxes or not plan or not plan.nodes:
        return 0.0
    plan_w = plan.canvas.width_pt or 1.0
    plan_h = plan.canvas.height_pt or 1.0
    plan_boxes = []
    for node in plan.nodes:
        b = node.geometry.bbox
        plan_boxes.append((b.left / plan_w, b.top / plan_h, b.width / plan_w, b.height / plan_h))
    # Greedy best-match IoU.
    total = 0.0
    matched = 0
    used = [False] * len(plan_boxes)
    for rb in ref_boxes:
        best = 0.0
        best_j = -1
        for j, pb in enumerate(plan_boxes):
            if used[j]:
                continue
            iou = _bbox_iou(rb, pb)
            if iou > best:
                best = iou
                best_j = j
        if best_j >= 0:
            used[best_j] = True
            total += best
            matched += 1
    return total / max(matched, 1)


def analyze_reference_deck(
    reference_pptx: str | Path,
) -> tuple[CanvasSpec, list[dict]]:
    """Analyze a reference deck and return (canvas, per-slide records with role)."""
    if analyze_presentation is None:  # pragma: no cover
        raise RuntimeError("reference_ppt helpers not available; scripts/ not importable")
    analysis = analyze_presentation(reference_pptx)
    canvas_info = _analysis_to_canvas(analysis)
    slides = analysis.get("slides", [])
    return canvas_info, slides


def _analysis_to_canvas(analysis: dict) -> CanvasSpec:
    """Build a CanvasSpec from the analysis's slide_size (reference-native)."""
    from pptx_skill.content_model import SafeInsets

    size = analysis.get("slide_size", {})
    width_pt = float(size.get("width_in", 13.333)) * 72.0
    height_pt = float(size.get("height_in", 7.5)) * 72.0
    return CanvasSpec(
        width_pt=width_pt,
        height_pt=height_pt,
        safe=SafeInsets(top=36.0, right=36.0, bottom=32.0, left=36.0),
    )


def visual_rebuild_adapter(
    reference_pptx: str | Path,
    output_path: str | Path,
    *,
    similarity_budget: float = DEFAULT_SIMILARITY_BUDGET,
    max_slides: int | None = None,
) -> VisualRebuildResult:
    """Rebuild a reference deck through the adaptive renderer.

    Parses the reference deck into stable SlideSpecs, picks a recipe per slide,
    solves geometry, and renders via ``render_layout_plans`` (NOT clone). Returns
    a result with the render trace and a reference-diff report bounded by the
    similarity budget.
    """
    if analyze_presentation is None:  # pragma: no cover
        raise RuntimeError("reference_ppt helpers not available; scripts/ not importable")

    canvas, slide_records = analyze_reference_deck(reference_pptx)
    if max_slides is not None:
        slide_records = slide_records[:max_slides]

    tokens = _builtin_tokens()
    plans: list[Any] = []
    rebuild_plans: list[SlideRebuildPlan] = []
    diagnostics_log: list[dict] = []

    for record in slide_records:
        index = record.get("index", 0)
        role = record.get("role", "bullets")
        slide_spec = _shapes_to_slide_spec(record, index, role)
        preferred_id = REBUILD_RECIPE_PREFERENCE.get(role)
        recipe = _pick_recipe(role)
        if recipe is None:
            diagnostics_log.append({"slide": index, "role": role, "skipped": "no_recipe"})
            continue
        geometry = solve_recipe(recipe, canvas, tokens)
        if geometry.infeasible:
            # Fall back to the first available recipe for the role.
            fallback = builtin_recipes(role)
            if not fallback:
                diagnostics_log.append({"slide": index, "role": role, "skipped": "no_fallback_recipe"})
                continue
            recipe = fallback[0]
            geometry = solve_recipe(recipe, canvas, tokens)
        layout_plan = solved_geometry_to_layout_plan(slide_spec, recipe, geometry, canvas, tokens)
        plans.append(layout_plan)
        rebuild_plans.append(SlideRebuildPlan(
            source_slide_index=index,
            detected_role=role,
            chosen_recipe_id=recipe.id,
            slide_spec=slide_spec,
            layout_plan=layout_plan,
            diagnostics={"preferred_recipe": preferred_id, "solved_vars": len(geometry.variables)},
        ))

    if not plans:
        raise ValueError("visual-rebuild produced no plans from the reference deck")

    target = Path(output_path).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    render_result = render_layout_plans(plans, str(target))

    # Reference diff: per-slide geometry similarity vs the reference shapes.
    diff: list[ReferenceDiffRecord] = []
    sims: list[float] = []
    for plan, rebuild in zip(plans, rebuild_plans, strict=False):
        ref_shapes = next(
            (r.get("shapes", []) for r in slide_records if r.get("index") == rebuild.source_slide_index),
            [],
        )
        sim = _slide_geometry_similarity(ref_shapes, plan)
        sims.append(sim)
        within = sim >= similarity_budget
        notes: list[str] = []
        if not within:
            notes.append(f"similarity {sim:.3f} below budget {similarity_budget:.3f}")
        diff.append(ReferenceDiffRecord(
            slide_index=rebuild.source_slide_index,
            role=rebuild.detected_role,
            geometry_similarity=round(sim, 4),
            within_budget=within,
            notes=notes,
        ))

    mean_sim = sum(sims) / len(sims) if sims else 0.0
    all_within = all(d.within_budget for d in diff)

    return VisualRebuildResult(
        output_path=str(target),
        render_result=render_result,
        canvas=canvas,
        plans=rebuild_plans,
        diff=diff,
        mean_similarity=round(mean_sim, 4),
        within_budget=all_within,
        diagnostics={
            "slides_analyzed": len(slide_records),
            "slides_rebuilt": len(plans),
            "similarity_budget": similarity_budget,
            "per_slide_diagnostics": diagnostics_log,
        },
    )


def reference_rebuild_diff_report(result: VisualRebuildResult) -> dict[str, Any]:
    """Serialize a rebuild result into a stable diff report for the manifest."""
    return result.to_dict()
