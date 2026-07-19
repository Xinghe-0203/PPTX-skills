"""Constraint-based adaptive layout engine (PR4).

Provides a structured linear constraint AST, a Kiwi solver wrapper, and a
candidate generator/scorer for bullets/text_image/dashboard roles.
"""
from __future__ import annotations

import json
import math
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any

from pptx_skill.content_model import BBox, CanvasSpec, GeometrySpec, LayoutPlan, PlannedNode, SlideSpec
from pptx_skill.semantic_qa import SemanticQAEngine
from pptx_skill.text_metrics import ParagraphStyle, TextRun, measure_runs


class Strength(str, Enum):
    REQUIRED = "required"
    STRONG = "strong"
    MEDIUM = "medium"
    WEAK = "weak"


@dataclass
class Term:
    var: str
    coef: float = 1.0


@dataclass
class Constraint:
    terms: list[Term]
    op: str  # ==, >=, <=
    rhs: dict  # one of: {"const": float}, {"ref": str}, {"token": str}
    strength: str = "required"


@dataclass
class ZoneSpec:
    kind: str
    style: str
    fit: str = "contain"


@dataclass
class LayoutRecipe:
    id: str
    role: str
    variant: str
    zones: dict[str, ZoneSpec]
    constraints: list[Constraint]
    content_limits: dict[str, Any] = field(default_factory=dict)
    fallbacks: list[str] = field(default_factory=list)


@dataclass
class SolvedVariable:
    name: str
    value: float


@dataclass
class SolvedGeometry:
    variables: dict[str, SolvedVariable]
    diagnostics: list[dict]
    infeasible: bool = False


@dataclass
class CandidateBundle:
    recipe: LayoutRecipe
    geometry: SolvedGeometry
    score: float
    blocker_count: int
    diagnostics: list[dict]


# ---------------------------------------------------------------------------
# Recipe loading / JSON schema
# ---------------------------------------------------------------------------


def _load_rhs(rhs: dict) -> dict:
    if not isinstance(rhs, dict):
        raise ValueError(f"RHS must be a dict, got {rhs!r}")
    allowed = {"const", "ref", "token"}
    keys = set(rhs.keys())
    if len(keys) != 1:
        raise ValueError(f"RHS must have exactly one of {allowed}, got {rhs!r}")
    return rhs


def recipe_from_dict(data: dict) -> LayoutRecipe:
    zones = {
        k: ZoneSpec(**v) if isinstance(v, dict) else v
        for k, v in data.get("zones", {}).items()
    }
    constraints = []
    for c in data.get("constraints", []):
        terms = [Term(var=t["var"], coef=t.get("coef", 1.0)) for t in c.get("terms", [])]
        op = c.get("op", "==")
        if op not in {"==", ">=", "<="}:
            raise ValueError(f"Unsupported constraint op: {op}")
        constraints.append(
            Constraint(
                terms=terms,
                op=op,
                rhs=_load_rhs(c.get("rhs", {"const": 0})),
                strength=c.get("strength", "required"),
            )
        )
    return LayoutRecipe(
        id=data["id"],
        role=data["role"],
        variant=data.get("variant", "default"),
        zones=zones,
        constraints=constraints,
        content_limits=data.get("content_limits", {}),
        fallbacks=data.get("fallbacks", []),
    )


def recipe_to_dict(recipe: LayoutRecipe) -> dict:
    return {
        "id": recipe.id,
        "role": recipe.role,
        "variant": recipe.variant,
        "zones": {k: {"kind": v.kind, "style": v.style, "fit": v.fit} for k, v in recipe.zones.items()},
        "constraints": [
            {
                "terms": [{"var": t.var, "coef": t.coef} for t in c.terms],
                "op": c.op,
                "rhs": c.rhs,
                "strength": c.strength,
            }
            for c in recipe.constraints
        ],
        "content_limits": recipe.content_limits,
        "fallbacks": recipe.fallbacks,
    }


# ---------------------------------------------------------------------------
# Token resolver
# ---------------------------------------------------------------------------


def _resolve_token(token: str, tokens: dict[str, Any]) -> float:
    parts = token.split(".")
    value = tokens
    for p in parts:
        if not isinstance(value, dict) or p not in value:
            raise ValueError(f"Token not found: {token}")
        value = value[p]
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, dict) and "$ref" in value:
        return _resolve_token(value["$ref"].lstrip("$").replace(".", "."), tokens)
    raise ValueError(f"Token value is not numeric: {token}")


# ---------------------------------------------------------------------------
# Constraint solver using Kiwi
# ---------------------------------------------------------------------------


def _strength_value(strength: str) -> Any:
    import kiwisolver as kiwi

    return {
        "required": kiwi.strength.required,
        "strong": kiwi.strength.strong,
        "medium": kiwi.strength.medium,
        "weak": kiwi.strength.weak,
    }.get(strength, kiwi.strength.required)


def solve_recipe(
    recipe: LayoutRecipe,
    canvas: CanvasSpec,
    tokens: dict[str, Any],
    hints: dict[str, float] | None = None,
) -> SolvedGeometry:
    """Solve a recipe's linear constraints using Kiwi."""
    import kiwisolver as kiwi

    diagnostics: list[dict] = []
    if _detect_infeasibility(recipe, canvas, tokens):
        diagnostics.append({"pre_check": "detected conflicting required constants"})
        return SolvedGeometry(variables={}, diagnostics=diagnostics, infeasible=True)

    solver = kiwi.Solver()
    variables: dict[str, Any] = {}
    diagnostics: list[dict] = []

    def get_var(name: str) -> Any:
        if name not in variables:
            variables[name] = kiwi.Variable(name)
        return variables[name]

    def rhs_value(rhs: dict) -> float:
        kind, value = next(iter(rhs.items()))
        if kind == "const":
            return float(value)
        if kind == "ref":
            ref = value
            if ref.startswith("canvas."):
                key = ref[len("canvas.") :]
                if key == "safe_left":
                    return canvas.safe.left
                if key == "safe_right":
                    return canvas.width_pt - canvas.safe.right
                if key == "safe_top":
                    return canvas.safe.top
                if key == "safe_bottom":
                    return canvas.height_pt - canvas.safe.bottom
                if key == "width":
                    return canvas.width_pt
                if key == "height":
                    return canvas.height_pt
                if key == "center_x":
                    return canvas.width_pt / 2.0
                if key == "center_y":
                    return canvas.height_pt / 2.0
            raise ValueError(f"Unknown ref: {ref}")
        if kind == "token":
            return _resolve_token(value, tokens)
        raise ValueError(f"Unknown rhs kind: {kind}")

    # Add constraints.
    for c in recipe.constraints:
        expr = sum(t.coef * get_var(t.var) for t in c.terms)
        rhs = rhs_value(c.rhs)
        strength = _strength_value(c.strength)
        if c.op == "==":
            solver.addConstraint((expr == rhs) | strength)
        elif c.op == ">=":
            solver.addConstraint((expr >= rhs) | strength)
        elif c.op == "<=":
            solver.addConstraint((expr <= rhs) | strength)

    # Add stay constraints for hints as weak objective.
    for name, value in (hints or {}).items():
        if name in variables:
            solver.addConstraint(get_var(name) == value | kiwi.strength.weak)

    try:
        solver.updateVariables()
    except Exception as exc:
        diagnostics.append({"solver_error": str(exc)})
        return SolvedGeometry(variables={}, diagnostics=diagnostics, infeasible=True)

    solved = {name: SolvedVariable(name, float(var.value())) for name, var in variables.items()}
    return SolvedGeometry(variables=solved, diagnostics=diagnostics, infeasible=False)


def _detect_infeasibility(recipe: LayoutRecipe, canvas: CanvasSpec, tokens: dict[str, Any]) -> bool:
    """Lightweight pre-check for obviously conflicting required constraints."""
    # Look for duplicate variables with different required constants.
    required_const: dict[str, float] = {}
    for c in recipe.constraints:
        if c.strength != "required":
            continue
        if len(c.terms) == 1 and c.terms[0].var and c.rhs.get("const") is not None:
            var = c.terms[0].var
            val = c.terms[0].coef * c.rhs["const"]
            if var in required_const and not math.isclose(required_const[var], val, rel_tol=1e-6):
                return True
            required_const[var] = val
    return False


# ---------------------------------------------------------------------------
# Candidate generation and scoring
# ---------------------------------------------------------------------------


def _zone_bboxes(variables: dict[str, SolvedVariable]) -> dict[str, BBox]:
    """Collect BBox for each zone from variables named {zone}.{side}."""
    vals: dict[str, float] = {name: var.value for name, var in variables.items()}
    zones: dict[str, dict[str, float]] = {}
    for name, value in vals.items():
        if "." not in name:
            continue
        zone, side = name.rsplit(".", 1)
        zones.setdefault(zone, {})[side] = value
    result: dict[str, BBox] = {}
    for zone, sides in zones.items():
        # Derive missing sides from known relations if possible.
        if "left" in sides and "width" in sides and "right" not in sides:
            sides["right"] = sides["left"] + sides["width"]
        if "right" in sides and "width" in sides and "left" not in sides:
            sides["left"] = sides["right"] - sides["width"]
        if "top" in sides and "height" in sides and "bottom" not in sides:
            sides["bottom"] = sides["top"] + sides["height"]
        if "bottom" in sides and "height" in sides and "top" not in sides:
            sides["top"] = sides["bottom"] - sides["height"]
        if {"left", "top", "right", "bottom"}.issubset(sides):
            result[zone] = BBox(
                x=sides["left"],
                y=sides["top"],
                width=sides["right"] - sides["left"],
                height=sides["bottom"] - sides["top"],
            )
    return result


def _estimate_text_height(text: str, bbox: BBox, style: dict, tokens: dict) -> tuple[float, bool]:
    if not text:
        return 0.0, False
    font_family = style.get("font_family", "Microsoft YaHei")
    size = float(style.get("size", 16))
    line_height = float(style.get("line_height", 1.35))
    min_size = float(style.get("min_size", 9))
    bold = style.get("bold", False)
    para = ParagraphStyle(line_height=line_height)
    run = TextRun(text=text, font_family=font_family, size_pt=size, bold=bold)
    try:
        metrics = measure_runs([run], bbox.width, para, size, min_font_size_pt=min_size)
    except Exception:
        # fallback: rough estimate
        lines = max(1, math.ceil(len(text) * size * 0.6 / bbox.width)) if bbox.width > 0 else 1
        return lines * size * line_height, False
    # Overflow only if the text cannot fit even at min font size.
    overflow = metrics.overflow
    return metrics.height_pt, overflow


def _score_candidate(
    recipe: LayoutRecipe,
    geometry: SolvedGeometry,
    canvas: CanvasSpec,
    slide: SlideSpec,
    tokens: dict[str, Any],
) -> CandidateBundle:
    diagnostics: list[dict] = []
    score = 0.0
    blockers = 0

    if geometry.infeasible:
        return CandidateBundle(recipe, geometry, score=1e9, blocker_count=1, diagnostics=diagnostics)

    bboxes = _zone_bboxes(geometry.variables)
    elements_by_role = {e.role: e for e in slide.elements}

    # Build LayoutPlan and run semantic QA.
    nodes: list[PlannedNode] = []
    for zone_name, zone in recipe.zones.items():
        bbox = bboxes.get(zone_name, BBox(0, 0, 100, 100))
        element = elements_by_role.get(zone_name)
        content: dict[str, Any] = {}
        resolved_style: dict[str, Any] = {}
        if element:
            content = element.content or {}
            resolved_style = {"font_family": "Microsoft YaHei", "size": 16.0}
            style_ref = element.style_ref or "component.body"
            token_size = tokens.get("component", {}).get(zone_name, {}).get("size", 16)
            resolved_style["size"] = float(token_size) if not isinstance(token_size, dict) else 16.0
        nodes.append(
            PlannedNode(
                id=f"{slide.id}/{zone_name}",
                element_id=element.id if element else None,
                recipe_node_id=zone_name,
                kind=zone.kind.split("-")[0] if "-" in zone.kind else zone.kind,
                role=zone_name,
                geometry=GeometrySpec(bbox),
                resolved_style=resolved_style,
                content_binding=content,
                z_order=0,
            )
        )

    plan = LayoutPlan(canvas=canvas, recipe_id=recipe.id, nodes=nodes, local_score=0.0)
    qa = SemanticQAEngine().check(plan)
    for issue in qa.issues:
        if issue.severity.value == "blocker":
            blockers += 1
            score += 1000
        elif issue.severity.value == "warning":
            score += 50
        diagnostics.append({"kind": issue.kind, "severity": issue.severity, "message": issue.message})

    # Text fit score.
    for node in nodes:
        if node.kind == "text":
            text = (node.content_binding or {}).get("text", "")
            h, overflow = _estimate_text_height(text, node.geometry.bbox, node.resolved_style or {}, tokens)
            if overflow:
                blockers += 1
                score += 120 * max(0.0, h / max(1.0, node.geometry.bbox.height) - 1.0)
            else:
                # prefer compact fit
                box_h = node.geometry.bbox.height
                if box_h > 0:
                    score += 20 * max(0.0, h / box_h)

    # Whitespace / density penalty.
    total_content_area = sum(b.width * b.height for b in bboxes.values())
    canvas_area = canvas.width_pt * canvas.height_pt
    density = total_content_area / canvas_area if canvas_area > 0 else 1.0
    target_density = tokens.get("grammar", {}).get("whitespace_ratio", {}).get("target", 0.72)
    score += 40 * abs(density - target_density)

    return CandidateBundle(recipe, geometry, score=score, blocker_count=blockers, diagnostics=diagnostics)


def plan_slide_candidates(
    slide: SlideSpec,
    recipes: list[LayoutRecipe],
    canvas: CanvasSpec,
    tokens: dict[str, Any],
    hints: dict[str, float] | None = None,
    max_candidates: int = 3,
) -> tuple[list[CandidateBundle], list[dict]]:
    """Solve and score a list of recipe candidates for a slide."""
    all_diagnostics: list[dict] = []
    bundles: list[CandidateBundle] = []
    for recipe in recipes:
        geometry = solve_recipe(recipe, canvas, tokens, hints=hints)
        bundle = _score_candidate(recipe, geometry, canvas, slide, tokens)
        all_diagnostics.append({"recipe": recipe.id, "score": bundle.score, "blockers": bundle.blocker_count})
        if bundle.blocker_count == 0:
            bundles.append(bundle)
    bundles.sort(key=lambda b: b.score)
    return bundles[:max_candidates], all_diagnostics


def solved_geometry_to_layout_plan(
    slide: SlideSpec,
    recipe: LayoutRecipe,
    geometry: SolvedGeometry,
    canvas: CanvasSpec,
    tokens: dict[str, Any],
) -> LayoutPlan:
    """Convert a solved recipe into a LayoutPlan."""
    bboxes = _zone_bboxes(geometry.variables)
    elements_by_role = {e.role: e for e in slide.elements}
    nodes: list[PlannedNode] = []
    for idx, (zone_name, zone) in enumerate(recipe.zones.items()):
        bbox = bboxes.get(zone_name, BBox(0, 0, 100, 100))
        element = elements_by_role.get(zone_name)
        style: dict[str, Any] = {"font_family": "Microsoft YaHei", "size": 16.0}
        if element:
            raw_size = tokens.get("component", {}).get(zone_name, {}).get("size", 16)
            if isinstance(raw_size, dict) and "$ref" in raw_size:
                raw_size = _resolve_token(raw_size["$ref"].lstrip("$").replace(".", "."), tokens)
            style["size"] = float(raw_size)
        nodes.append(
            PlannedNode(
                id=f"{slide.id}/{zone_name}",
                element_id=element.id if element else None,
                recipe_node_id=zone_name,
                kind=zone.kind.split("-")[0] if "-" in zone.kind else zone.kind,
                role=zone_name,
                geometry=GeometrySpec(bbox),
                resolved_style=style,
                content_binding=(element.content if element else {}) or {},
                z_order=idx,
            )
        )
    return LayoutPlan(canvas=canvas, recipe_id=recipe.id, nodes=nodes, local_score=0.0)


# ---------------------------------------------------------------------------
# Built-in recipe catalog for bullets/text_image/dashboard
# ---------------------------------------------------------------------------


def _builtin_tokens() -> dict[str, Any]:
    return {
        "primitive": {
            "palette": {"paper_050": "#FDFCF8", "ink_950": "#1A1A1A", "accent": "#184E77"},
            "space": {"4": 16, "6": 24, "8": 32, "12": 48},
            "font": {"size": {"title": 32, "body": 18}},
        },
        "semantic": {
            "space": {"component": {"$ref": "primitive.space.4"}, "section": {"$ref": "primitive.space.8"}},
            "color": {"text_primary": {"$ref": "primitive.palette.ink_950"}},
        },
        "component": {
            "title": {"size": {"$ref": "primitive.font.size.title"}},
            "body": {"size": {"$ref": "primitive.font.size.body"}},
        },
        "grammar": {"whitespace_ratio": {"target": 0.72}},
    }


def _recipe(name: str, role: str, variant: str, zones: dict, constraints: list, limits: dict | None = None) -> LayoutRecipe:
    # Auto-link closure constraints so every zone's bottom/right derive from
    # top+height / left+width. Without these, Kiwi leaves bottom/right
    # under-constrained and minimizes them to degenerate values.
    def has_link(zone: str, side: str) -> bool:
        """True if a constraint already links zone.side to the other two sides."""
        target = f"{zone}.{side}"
        for c in constraints:
            var_names = [t["var"] for t in c.get("terms", [])]
            if target in var_names and len(var_names) == 3 and c.get("op") == "==":
                return True
        return False

    for zone in zones:
        # bottom = top + height  (when height is set but bottom isn't linked)
        has_height = any(
            f"{zone}.height" in [t["var"] for t in c.get("terms", [])]
            for c in constraints
        )
        has_bottom = any(
            f"{zone}.bottom" in [t["var"] for t in c.get("terms", [])]
            for c in constraints
        )
        if has_height and has_bottom and not has_link(zone, "bottom"):
            constraints.append({
                "terms": [{"var": f"{zone}.bottom"}, {"var": f"{zone}.top", "coef": -1}, {"var": f"{zone}.height", "coef": -1}],
                "op": "==", "rhs": {"const": 0},
            })
        # right = left + width
        has_width = any(
            f"{zone}.width" in [t["var"] for t in c.get("terms", [])]
            for c in constraints
        )
        has_right = any(
            f"{zone}.right" in [t["var"] for t in c.get("terms", [])]
            for c in constraints
        )
        if has_width and has_right and not has_link(zone, "right"):
            constraints.append({
                "terms": [{"var": f"{zone}.right"}, {"var": f"{zone}.left", "coef": -1}, {"var": f"{zone}.width", "coef": -1}],
                "op": "==", "rhs": {"const": 0},
            })

    return LayoutRecipe(
        id=name,
        role=role,
        variant=variant,
        zones={k: ZoneSpec(**v) for k, v in zones.items()},
        constraints=[
            Constraint(
                terms=[Term(t["var"], t.get("coef", 1.0)) for t in c["terms"]],
                op=c["op"],
                rhs=_load_rhs(c["rhs"]),
                strength=c.get("strength", "required"),
            )
            for c in constraints
        ],
        content_limits=limits or {},
    )


def builtin_recipes(role: str) -> list[LayoutRecipe]:
    """Return built-in candidate recipes for a role."""
    recipes: list[LayoutRecipe] = []
    if role == "bullets":
        recipes.append(_recipe(
            "bullets.rail",
            "bullets",
            "rail",
            {
                "title": {"kind": "text", "style": "component.title"},
                "body": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"body.max_items": 7, "body.min_font_size": 15},
        ))
        recipes.append(_recipe(
            "bullets.wide",
            "bullets",
            "wide",
            {
                "title": {"kind": "text", "style": "component.title"},
                "body": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"body.max_items": 10, "body.min_font_size": 14},
        ))
    elif role == "text_image":
        recipes.append(_recipe(
            "text_image.asymmetric_left",
            "text_image",
            "asymmetric_left",
            {
                "title": {"kind": "text", "style": "component.title"},
                "body": {"kind": "text", "style": "component.body"},
                "hero": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.width"}], "op": "==", "rhs": {"const": 360}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.right"}, {"var": "title.left", "coef": -1}, {"var": "title.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.width"}], "op": "==", "rhs": {"const": 360}},
                {"terms": [{"var": "body.right"}, {"var": "body.left", "coef": -1}, {"var": "body.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "hero.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "hero.left"}, {"var": "body.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
            ],
        ))
        recipes.append(_recipe(
            "text_image.stacked",
            "text_image",
            "stacked",
            {
                "title": {"kind": "text", "style": "component.title"},
                "body": {"kind": "text", "style": "component.body"},
                "hero": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "hero.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.height"}], "op": "==", "rhs": {"const": 220}},
                {"terms": [{"var": "hero.bottom"}, {"var": "hero.top", "coef": -1}, {"var": "hero.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}, {"var": "hero.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        ))
    elif role == "dashboard":
        recipes.append(_recipe(
            "dashboard.lead_metric",
            "dashboard",
            "lead_metric",
            {
                "title": {"kind": "text", "style": "component.title"},
                "lead": {"kind": "text", "style": "component.metric"},
                "support": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "lead.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "lead.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "lead.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "lead.height"}], "op": "==", "rhs": {"const": 120}},
                {"terms": [{"var": "lead.bottom"}, {"var": "lead.top", "coef": -1}, {"var": "lead.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "support.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "support.top"}, {"var": "lead.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "support.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "support.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        ))
        recipes.append(_recipe(
            "dashboard.split",
            "dashboard",
            "split",
            {
                "title": {"kind": "text", "style": "component.title"},
                "left": {"kind": "text", "style": "component.metric"},
                "right": {"kind": "text", "style": "component.metric"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "left.width"}], "op": "==", "rhs": {"const": 420}},
                {"terms": [{"var": "left.right"}, {"var": "left.left", "coef": -1}, {"var": "left.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "left.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "right.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "right.width"}], "op": "==", "rhs": {"const": 420}},
                {"terms": [{"var": "right.left"}, {"var": "right.right", "coef": -1}, {"var": "right.width", "coef": 1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "right.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "right.left"}, {"var": "left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
            ],
        ))
    elif role == "cover":
        recipes.append(_recipe(
            "cover.axis",
            "cover",
            "axis",
            {
                "kicker": {"kind": "text", "style": "component.title"},
                "title": {"kind": "text", "style": "component.title"},
                "subtitle": {"kind": "text", "style": "component.body"},
                "hero": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "kicker.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "kicker.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "kicker.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "kicker.height"}], "op": "==", "rhs": {"const": 24}},
                {"terms": [{"var": "kicker.bottom"}, {"var": "kicker.top", "coef": -1}, {"var": "kicker.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}, {"var": "kicker.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 80}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "subtitle.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "subtitle.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "subtitle.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "subtitle.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.top"}, {"var": "kicker.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "hero.height"}], "op": ">=", "rhs": {"const": 200}},
                {"terms": [{"var": "hero.bottom"}, {"var": "hero.top", "coef": -1}, {"var": "hero.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 36},
        ))
        recipes.append(_recipe(
            "cover.poster",
            "cover",
            "poster",
            {
                "title": {"kind": "text", "style": "component.title"},
                "hero": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "footer": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "hero.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero.top"}], "op": ">=", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "footer.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "footer.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "footer.height"}], "op": "==", "rhs": {"const": 20}},
            ],
            {"title.max_lines": 3, "title.min_font_size": 32},
        ))
    elif role == "toc":
        recipes.append(_recipe(
            "toc.list",
            "toc",
            "list",
            {
                "title": {"kind": "text", "style": "component.title"},
                "items": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "items.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "items.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "items.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"items.max_items": 8, "items.min_font_size": 16},
        ))
        recipes.append(_recipe(
            "toc.rail",
            "toc",
            "rail",
            {
                "title": {"kind": "text", "style": "component.title"},
                "items": {"kind": "text-list", "style": "component.body"},
                "rail": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "rail.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "rail.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "rail.width"}], "op": "==", "rhs": {"const": 4}},
                {"terms": [{"var": "rail.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "items.left"}], "op": ">=", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "items.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "items.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"items.max_items": 8, "items.min_font_size": 16},
        ))
    elif role == "section":
        recipes.append(_recipe(
            "section.full",
            "section",
            "full",
            {
                "number": {"kind": "text", "style": "component.title"},
                "title": {"kind": "text", "style": "component.title"},
                "subtitle": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "number.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "number.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "number.height"}], "op": "==", "rhs": {"const": 80}},
                {"terms": [{"var": "number.bottom"}, {"var": "number.top", "coef": -1}, {"var": "number.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}, {"var": "number.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 80}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "subtitle.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "subtitle.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "subtitle.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "subtitle.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 36},
        ))
        recipes.append(_recipe(
            "section.split",
            "section",
            "split",
            {
                "title": {"kind": "text", "style": "component.title"},
                "kicker": {"kind": "text", "style": "component.body"},
                "rail": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "kicker.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "kicker.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "kicker.height"}], "op": "==", "rhs": {"const": 24}},
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}, {"var": "kicker.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "rail.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "rail.top"}], "op": ">=", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "rail.width"}], "op": "==", "rhs": {"const": 96}},
                {"terms": [{"var": "rail.height"}], "op": "==", "rhs": {"const": 4}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 36},
        ))
    elif role == "full_image":
        recipes.append(_recipe(
            "full_image.full",
            "full_image",
            "full",
            {
                "hero": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "caption": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "hero.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "caption.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "caption.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "caption.height"}], "op": "==", "rhs": {"const": 20}},
                {"terms": [{"var": "caption.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
            ],
        ))
        recipes.append(_recipe(
            "full_image.caption",
            "full_image",
            "caption",
            {
                "hero": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "title": {"kind": "text", "style": "component.title"},
                "caption": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "hero.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}, {"var": "hero.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "caption.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "caption.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "caption.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "caption.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 26},
        ))
    elif role == "image_grid":
        recipes.append(_recipe(
            "image_grid.grid2",
            "image_grid",
            "grid2",
            {
                "title": {"kind": "text", "style": "component.title"},
                "a": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "b": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "a.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "a.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "a.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "a.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "b.left"}, {"var": "a.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "b.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
        ))
        recipes.append(_recipe(
            "image_grid.grid3",
            "image_grid",
            "grid3",
            {
                "title": {"kind": "text", "style": "component.title"},
                "a": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "b": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "c": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "a.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "a.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "a.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "a.right"}, {"var": "b.left", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.right"}, {"var": "c.left", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "c.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Equal column widths: a.width == b.width == c.width.
                {"terms": [{"var": "a.width"}, {"var": "b.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.width"}, {"var": "c.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
        ))
    elif role == "timeline":
        recipes.append(_recipe(
            "timeline.horizontal",
            "timeline",
            "horizontal",
            {
                "title": {"kind": "text", "style": "component.title"},
                "items": {"kind": "text-list", "style": "component.body"},
                "axis": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "axis.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "axis.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "axis.top"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "axis.height"}], "op": "==", "rhs": {"const": 2}},
                {"terms": [{"var": "items.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "items.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "items.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"items.max_items": 6, "items.min_font_size": 14},
        ))
        recipes.append(_recipe(
            "timeline.vertical",
            "timeline",
            "vertical",
            {
                "title": {"kind": "text", "style": "component.title"},
                "items": {"kind": "text-list", "style": "component.body"},
                "axis": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "axis.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "axis.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "axis.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "axis.width"}], "op": "==", "rhs": {"const": 2}},
                {"terms": [{"var": "items.left"}, {"var": "axis.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "items.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"items.max_items": 6, "items.min_font_size": 14},
        ))
    elif role == "comparison":
        recipes.append(_recipe(
            "comparison.two_column",
            "comparison",
            "two_column",
            {
                "title": {"kind": "text", "style": "component.title"},
                "left": {"kind": "text", "style": "component.body"},
                "right": {"kind": "text", "style": "component.body"},
                "divider": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "divider.left"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "divider.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "divider.width"}], "op": "==", "rhs": {"const": 1}},
                {"terms": [{"var": "divider.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "left.right"}, {"var": "divider.left", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "left.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "right.left"}, {"var": "divider.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "right.top"}, {"var": "left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "right.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        ))
        recipes.append(_recipe(
            "comparison.stacked",
            "comparison",
            "stacked",
            {
                "title": {"kind": "text", "style": "component.title"},
                "left": {"kind": "text", "style": "component.body"},
                "right": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "left.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "left.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "right.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "right.top"}, {"var": "left.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "right.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        ))
    elif role == "quote":
        recipes.append(_recipe(
            "quote.centered",
            "quote",
            "centered",
            {
                "mark": {"kind": "text", "style": "component.title"},
                "body": {"kind": "text", "style": "component.title"},
                "attribution": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "mark.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "mark.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "mark.height"}], "op": "==", "rhs": {"const": 80}},
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}, {"var": "mark.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "attribution.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "attribution.top"}, {"var": "body.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "attribution.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "attribution.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "attribution.height"}], "op": "==", "rhs": {"const": 24}},
            ],
            {"body.max_lines": 4, "body.min_font_size": 24},
        ))
        recipes.append(_recipe(
            "quote.editorial",
            "quote",
            "editorial",
            {
                "body": {"kind": "text", "style": "component.title"},
                "attribution": {"kind": "text", "style": "component.body"},
                "rule": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "rule.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "rule.top"}], "op": ">=", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "rule.width"}], "op": "==", "rhs": {"const": 64}},
                {"terms": [{"var": "rule.height"}], "op": "==", "rhs": {"const": 2}},
                {"terms": [{"var": "attribution.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "attribution.top"}, {"var": "rule.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "attribution.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "attribution.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "attribution.height"}], "op": "==", "rhs": {"const": 24}},
            ],
            {"body.max_lines": 4, "body.min_font_size": 22},
        ))
    elif role == "process":
        recipes.append(_recipe(
            "process.horizontal",
            "process",
            "horizontal",
            {
                "title": {"kind": "text", "style": "component.title"},
                "steps": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "steps.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "steps.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "steps.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "steps.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"steps.max_items": 5, "steps.min_font_size": 14},
        ))
        recipes.append(_recipe(
            "process.vertical",
            "process",
            "vertical",
            {
                "title": {"kind": "text", "style": "component.title"},
                "steps": {"kind": "text-list", "style": "component.body"},
                "rail": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "rail.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "rail.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "rail.width"}], "op": "==", "rhs": {"const": 2}},
                {"terms": [{"var": "rail.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "steps.left"}, {"var": "rail.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "steps.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "steps.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "steps.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"steps.max_items": 6, "steps.min_font_size": 14},
        ))
    elif role == "table":
        recipes.append(_recipe(
            "table.standard",
            "table",
            "standard",
            {
                "title": {"kind": "text", "style": "component.title"},
                "table": {"kind": "table", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "table.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "table.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "table.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "table.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        ))
        recipes.append(_recipe(
            "table.striped",
            "table",
            "striped",
            {
                "title": {"kind": "text", "style": "component.title"},
                "table": {"kind": "table", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "table.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "table.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "table.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "table.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
        ))
    elif role == "end":
        recipes.append(_recipe(
            "end.thanks",
            "end",
            "thanks",
            {
                "title": {"kind": "text", "style": "component.title"},
                "subtitle": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "subtitle.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "subtitle.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "subtitle.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "subtitle.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 36},
        ))
        recipes.append(_recipe(
            "end.minimal",
            "end",
            "minimal",
            {
                "title": {"kind": "text", "style": "component.title"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 36},
        ))
    return recipes
