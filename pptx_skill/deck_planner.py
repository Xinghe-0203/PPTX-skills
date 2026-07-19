"""Deck-level layout planner with beam search (PR6).

Plans a whole deck by collecting per-slide candidate bundles and selecting a
sequence that minimizes local scores plus transition penalties (adjacent
geometry similarity, rhythm rules, ``preferred_sequence`` and section
transitions). Pagination splits are produced by the ``pagination`` module;
the deck planner consumes already-derived slides and never re-splits.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from pptx_skill.content_model import (
    CanvasSpec,
    DeckPlanResult,
    LayoutPlan,
    PlannedNode,
    SlidePlanCandidate,
    SlidePlanResult,
    SlideSpec,
)
from pptx_skill.layout_engine import (
    builtin_recipes,
    plan_slide_candidates,
    solved_geometry_to_layout_plan,
)
from pptx_skill.pagination import paginate_bullets
from pptx_skill.semantic_qa import SemanticQAEngine


@dataclass
class LayoutScoringConfig:
    """Weights for transition scoring (see blueprint 7.5)."""

    adjacent_geometry_similarity: float = 8.0
    rhythm_rule_penalty: float = 6.0
    preferred_sequence_penalty: float = 6.0
    section_transition_penalty: float = 4.0
    adjacent_signature_threshold: float = 0.85
    split_score_threshold: float = 200.0
    max_candidates: int = 3
    beam_width: int = 8


@dataclass
class _BeamState:
    """A partial deck plan in the beam search."""

    chosen: list[SlidePlanCandidate]
    derived_slides: list[SlideSpec]
    plans: list[LayoutPlan]
    accumulated_score: float

    def branch(self) -> _BeamState:
        return _BeamState(
            chosen=list(self.chosen),
            derived_slides=list(self.derived_slides),
            plans=list(self.plans),
            accumulated_score=self.accumulated_score,
        )


def _geometry_signature(plan: LayoutPlan) -> tuple:
    """A normalized signature of a plan for similarity comparison.

    Captures role order, normalized bboxes and area ratios so two pages with
    the same skeleton compare as similar regardless of absolute coordinates.
    """
    sig: list[tuple] = []
    canvas_w = plan.canvas.width_pt or 1.0
    canvas_h = plan.canvas.height_pt or 1.0
    for node in plan.nodes:
        bbox = node.geometry.bbox if node.geometry else None
        if bbox is None:
            continue
        sig.append(
            (
                node.role,
                round(bbox.x / canvas_w, 3),
                round(bbox.y / canvas_h, 3),
                round(bbox.width / canvas_w, 3),
                round(bbox.height / canvas_h, 3),
            )
        )
    return tuple(sig)


def _signature_similarity(a: LayoutPlan, b: LayoutPlan) -> float:
    """Return 0..1 similarity between two plans' geometry signatures."""
    sig_a = _geometry_signature(a)
    sig_b = _geometry_signature(b)
    if not sig_a and not sig_b:
        return 1.0
    if not sig_a or not sig_b:
        return 0.0
    set_a = set(sig_a)
    set_b = set(sig_b)
    intersection = len(set_a & set_b)
    union = len(set_a | set_b)
    return intersection / union if union else 0.0


def _role_of_candidate(candidate: SlidePlanCandidate) -> str:
    if candidate.derived_slides:
        return candidate.derived_slides[0].role
    return ""


def _preferred_sequence_penalty(
    prev_role: str,
    cur_role: str,
    preferred_sequence: list[str] | None,
    config: LayoutScoringConfig,
) -> float:
    """Penalize transitions that violate the template's preferred_sequence."""
    if not preferred_sequence or not prev_role or not cur_role:
        return 0.0
    if prev_role not in preferred_sequence or cur_role not in preferred_sequence:
        return config.preferred_sequence_penalty * 0.5
    prev_idx = preferred_sequence.index(prev_role)
    cur_idx = preferred_sequence.index(cur_role)
    # Forward or same-position transitions are encouraged; large backward jumps penalized.
    if cur_idx >= prev_idx:
        return 0.0
    # Backward transition within the same emphasis family is mild; cross-family is heavier.
    return config.preferred_sequence_penalty * (prev_idx - cur_idx) / max(1, len(preferred_sequence))


def _rhythm_rule_penalty(
    prev_role: str,
    cur_role: str,
    rhythm_rules: list[dict] | None,
    config: LayoutScoringConfig,
) -> float:
    """Apply ``rhythm.rules`` from a profile (after X prefer/avoid Y)."""
    if not rhythm_rules or not prev_role or not cur_role:
        return 0.0
    penalty = 0.0
    for rule in rhythm_rules:
        if rule.get("after") != prev_role:
            continue
        avoid = rule.get("avoid", [])
        prefer = rule.get("prefer", [])
        if cur_role in avoid:
            penalty += config.rhythm_rule_penalty
        elif prefer and cur_role in prefer:
            penalty -= config.rhythm_rule_penalty * 0.25
    return max(0.0, penalty)


def _section_transition_penalty(
    prev: SlideSpec | None,
    cur: SlideSpec | None,
    config: LayoutScoringConfig,
) -> float:
    """Light penalty when consecutive slides come from different source sections."""
    if prev is None or cur is None:
        return 0.0
    if prev.source_section_id and cur.source_section_id:
        if prev.source_section_id != cur.source_section_id:
            return config.section_transition_penalty
    return 0.0


def _transition_score(
    prev_candidate: SlidePlanCandidate | None,
    cur_candidate: SlidePlanCandidate,
    preferred_sequence: list[str] | None,
    rhythm_rules: list[dict] | None,
    config: LayoutScoringConfig,
) -> float:
    """Compute the transition penalty between two adjacent candidate bundles."""
    if prev_candidate is None:
        return 0.0
    prev_plan = prev_candidate.plans[-1] if prev_candidate.plans else None
    cur_plan = cur_candidate.plans[0] if cur_candidate.plans else None
    score = 0.0
    if prev_plan and cur_plan:
        similarity = _signature_similarity(prev_plan, cur_plan)
        if similarity > config.adjacent_signature_threshold:
            score += config.adjacent_geometry_similarity * similarity
    prev_role = _role_of_candidate(prev_candidate)
    cur_role = _role_of_candidate(cur_candidate)
    score += _preferred_sequence_penalty(prev_role, cur_role, preferred_sequence, config)
    score += _rhythm_rule_penalty(prev_role, cur_role, rhythm_rules, config)
    prev_slide = prev_candidate.derived_slides[-1] if prev_candidate.derived_slides else None
    cur_slide = cur_candidate.derived_slides[0] if cur_candidate.derived_slides else None
    score += _section_transition_penalty(prev_slide, cur_slide, config)
    return score


def _plan_slide(
    slide: SlideSpec,
    canvas: CanvasSpec,
    tokens: dict[str, Any],
    config: LayoutScoringConfig,
    qa_engine: SemanticQAEngine,
    min_font_size_pt: float,
) -> SlidePlanResult:
    """Plan a single source slide: try single-page candidates, split if overloaded."""
    role = slide.role
    recipes = (
        builtin_recipes(role)
        if role in {"bullets", "text_image", "dashboard"}
        else builtin_recipes("bullets")
    )
    if not recipes:
        return SlidePlanResult(status="infeasible", candidates=[], blockers=[{"reason": "no recipes"}], diagnostics=[])

    candidates, diagnostics = plan_slide_candidates(
        slide, recipes, canvas, tokens, max_candidates=config.max_candidates
    )

    bundles: list[SlidePlanCandidate] = []
    for bundle in candidates:
        plans: list[LayoutPlan] = []
        derived = [slide]
        ok = True
        for derived_slide in derived:
            plan = solved_geometry_to_layout_plan(derived_slide, bundle.recipe, bundle.geometry, canvas, tokens)
            report = qa_engine.check(plan)
            if not report.passed:
                ok = False
                break
            plans.append(plan)
        if ok:
            bundles.append(
                SlidePlanCandidate(
                    derived_slides=derived,
                    plans=plans,
                    local_score=bundle.score,
                    diagnostics=list(bundle.diagnostics),
                    kind="single",
                )
            )

    # If best single candidate is too dense or none feasible, try splitting bullets.
    best_score = min((b.local_score for b in bundles), default=float("inf"))
    needs_split = (
        role == "bullets"
        and (not bundles or best_score > config.split_score_threshold)
    )
    if needs_split:
        split_slides = paginate_bullets(
            slide,
            max_items=int(tokens.get("content_limits", {}).get("bullets_max_items", 7)),
            min_font_size_pt=min_font_size_pt,
        )
        if split_slides:
            split_bundle = _build_split_candidate(
                split_slides, recipes[0], canvas, tokens, qa_engine
            )
            if split_bundle is not None:
                bundles.append(split_bundle)

    bundles.sort(key=lambda b: b.local_score)
    if not bundles:
        return SlidePlanResult(status="infeasible", candidates=[], blockers=diagnostics, diagnostics=diagnostics)
    return SlidePlanResult(status="feasible", candidates=bundles[: config.max_candidates], blockers=[], diagnostics=diagnostics)


def _build_split_candidate(
    split_slides: list[SlideSpec],
    recipe_template: Any,
    canvas: CanvasSpec,
    tokens: dict[str, Any],
    qa_engine: SemanticQAEngine,
) -> SlidePlanCandidate | None:
    """Build a split candidate bundle from paginated slides using the first recipe."""
    from pptx_skill.layout_engine import solve_recipe

    plans: list[LayoutPlan] = []
    total_score = 0.0
    for derived_slide in split_slides:
        geometry = solve_recipe(recipe_template, canvas, tokens)
        if geometry.infeasible:
            return None
        plan = solved_geometry_to_layout_plan(derived_slide, recipe_template, geometry, canvas, tokens)
        report = qa_engine.check(plan)
        if not report.passed:
            # Split pages that still fail are kept but penalized; do not drop silently.
            total_score += 1000.0
        plans.append(plan)
    return SlidePlanCandidate(
        derived_slides=split_slides,
        plans=plans,
        local_score=total_score / max(1, len(split_slides)),
        diagnostics=[{"split": True, "pages": len(split_slides)}],
        kind="split",
    )


def plan_deck(
    slides: list[SlideSpec],
    canvas: CanvasSpec,
    profile_state: dict[str, Any],
    config: LayoutScoringConfig | None = None,
) -> DeckPlanResult:
    """Plan layouts for an entire deck using beam search over per-slide candidates."""
    config = config or LayoutScoringConfig()
    tokens = profile_state.get("tokens", {})
    preferred_sequence = profile_state.get("preferred_sequence")
    rhythm_rules = profile_state.get("rhythm", {}).get("rules") if isinstance(profile_state.get("rhythm"), dict) else None
    qa_engine = SemanticQAEngine()
    min_font_size_pt = float(
        profile_state.get("qa", {}).get("min_body_font_size", 9)
    )

    blockers: list[dict] = []
    diagnostics: list[dict] = []

    # Per-slide SlidePlanResult.
    per_slide: list[SlidePlanResult] = []
    for idx, slide in enumerate(slides):
        result = _plan_slide(slide, canvas, tokens, config, qa_engine, min_font_size_pt)
        per_slide.append(result)
        diagnostics.append({"slide_index": idx, "status": result.status, "candidates": len(result.candidates)})
        if result.status == "infeasible":
            blockers.append({"slide_index": idx, "blockers": result.blockers})
            # Inject a passthrough candidate so the beam can still continue; it will be flagged.
            per_slide[-1] = SlidePlanResult(
                status="feasible",
                candidates=[
                    SlidePlanCandidate(
                        derived_slides=[slide],
                        plans=[_fallback_plan(slide, canvas)],
                        local_score=1e6,
                        diagnostics=result.diagnostics,
                        kind="single",
                    )
                ],
                blockers=result.blockers,
                diagnostics=result.diagnostics,
            )

    # Beam search.
    beam: list[_BeamState] = [_BeamState(chosen=[], derived_slides=[], plans=[], accumulated_score=0.0)]
    for idx, result in enumerate(per_slide):
        new_beam: list[_BeamState] = []
        for state in beam:
            prev_candidate = state.chosen[-1] if state.chosen else None
            for candidate in result.candidates:
                branch = state.branch()
                transition = _transition_score(
                    prev_candidate, candidate, preferred_sequence, rhythm_rules, config
                )
                branch.accumulated_score += candidate.local_score + transition
                branch.chosen.append(candidate)
                branch.derived_slides.extend(candidate.derived_slides)
                branch.plans.extend(candidate.plans)
                new_beam.append(branch)
        if not new_beam:
            break
        new_beam.sort(key=lambda s: s.accumulated_score)
        beam = new_beam[: config.beam_width]

    if not beam:
        return DeckPlanResult(
            status="infeasible",
            source_slides=slides,
            derived_slides=[],
            plans=[],
            blockers=blockers or [{"reason": "beam search produced no states"}],
            diagnostics=diagnostics,
        )

    best = beam[0]
    status = "infeasible" if blockers else "feasible"
    return DeckPlanResult(
        status=status,
        source_slides=slides,
        derived_slides=best.derived_slides,
        plans=best.plans,
        blockers=blockers,
        diagnostics=diagnostics + [{"beam_score": best.accumulated_score, "beam_size": len(beam)}],
    )


def _fallback_plan(slide: SlideSpec, canvas: CanvasSpec) -> LayoutPlan:
    """A minimal plan for an infeasible slide so the beam can continue."""
    from pptx_skill.content_model import BBox, GeometrySpec

    nodes: list[PlannedNode] = []
    for idx, element in enumerate(slide.elements):
        nodes.append(
            PlannedNode(
                id=f"{slide.id}/{element.role}",
                element_id=element.id,
                recipe_node_id=element.role,
                kind=element.kind,
                role=element.role,
                geometry=GeometrySpec(BBox(48, 48 + idx * 60, max(canvas.width_pt - 96, 24), 50)),
                resolved_style={"font_family": "Microsoft YaHei", "size": 16.0},
                content_binding=element.content or {},
                z_order=idx,
            )
        )
    return LayoutPlan(canvas=canvas, recipe_id="fallback", nodes=nodes, local_score=1e6)
