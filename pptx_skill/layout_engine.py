"""Constraint-based adaptive layout engine (PR4).

Provides a structured linear constraint AST, a Kiwi solver wrapper, and a
candidate generator/scorer for bullets/text_image/dashboard roles.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any

from pptx_skill._compat import StrEnum
from pptx_skill.content_model import BBox, CanvasSpec, GeometrySpec, LayoutPlan, PlannedNode, SlideSpec
from pptx_skill.semantic_qa import SemanticQAEngine
from pptx_skill.text_metrics import ParagraphStyle, TextRun, measure_runs


class Strength(StrEnum):
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
        k: ZoneSpec(**v) if isinstance(v, dict) else ZoneSpec(kind=str(v), style="default")
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


def _resolve_token(token: str, tokens: dict[str, Any], _visited: set[str] | None = None) -> float:
    _visited = _visited or set()
    if token in _visited:
        raise ValueError(f"Circular token reference detected: {token}")
    _visited.add(token)
    parts = token.split(".")
    value = tokens
    for p in parts:
        if not isinstance(value, dict) or p not in value:
            raise ValueError(f"Token not found: {token}")
        value = value[p]
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, dict) and "$ref" in value:
        return _resolve_token(value["$ref"].lstrip("$"), tokens, _visited)
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
            solver.addConstraint((get_var(name) == value) | kiwi.strength.weak)

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


def _normalize_node_kind(raw_kind: str) -> str:
    """Map a zone kind to a renderable node kind."""
    _KIND_MAP = {
        "text-list": "text",
        "chart-bar": "chart",
        "chart-line": "chart",
        "chart-pie": "chart",
        "chart-area": "chart",
    }
    return _KIND_MAP.get(raw_kind, raw_kind.split("-")[0] if "-" in raw_kind else raw_kind)


def _score_candidate(
    recipe: LayoutRecipe,
    geometry: SolvedGeometry,
    canvas: CanvasSpec,
    slide: SlideSpec,
    tokens: dict[str, Any],
    font_family: str | None = None,
) -> CandidateBundle:
    diagnostics: list[dict] = []
    score = 0.0
    blockers = 0

    if geometry.infeasible:
        return CandidateBundle(recipe, geometry, score=1e9, blocker_count=1, diagnostics=diagnostics)

    _default_font = font_family or tokens.get("primitive", {}).get("font", {}).get("family", {}).get("body", "Microsoft YaHei")
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
            resolved_style = {"font_family": _default_font, "size": 16.0}
            token_size = tokens.get("component", {}).get(zone_name, {}).get("size", 16)
            resolved_style["size"] = float(token_size) if not isinstance(token_size, dict) else 16.0
        nodes.append(
            PlannedNode(
                id=f"{slide.id}/{zone_name}",
                element_id=element.id if element else None,
                recipe_node_id=zone_name,
                kind=_normalize_node_kind(zone.kind),
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
    total_content_area = 0.0
    bbox_list = list(bboxes.values())
    for i, a in enumerate(bbox_list):
        area = a.width * a.height
        for j in range(i):
            b = bbox_list[j]
            ox = max(0.0, min(a.x + a.width, b.x + b.width) - max(a.x, b.x))
            oy = max(0.0, min(a.y + a.height, b.y + b.height) - max(a.y, b.y))
            area -= ox * oy
        total_content_area += area
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
    font_family: str | None = None,
) -> tuple[list[CandidateBundle], list[dict]]:
    """Solve and score a list of recipe candidates for a slide."""
    all_diagnostics: list[dict] = []
    bundles: list[CandidateBundle] = []
    all_bundles: list[CandidateBundle] = []
    for recipe in recipes:
        geometry = solve_recipe(recipe, canvas, tokens, hints=hints)
        bundle = _score_candidate(recipe, geometry, canvas, slide, tokens, font_family=font_family)
        all_diagnostics.append({"recipe": recipe.id, "score": bundle.score, "blockers": bundle.blocker_count})
        all_bundles.append(bundle)
        if bundle.blocker_count == 0:
            bundles.append(bundle)
    if not bundles and all_bundles:
        all_bundles.sort(key=lambda b: b.score)
        bundles = [all_bundles[0]]
    bundles.sort(key=lambda b: b.score)
    return bundles[:max_candidates], all_diagnostics


def solved_geometry_to_layout_plan(
    slide: SlideSpec,
    recipe: LayoutRecipe,
    geometry: SolvedGeometry,
    canvas: CanvasSpec,
    tokens: dict[str, Any],
    font_family: str | None = None,
) -> LayoutPlan:
    """Convert a solved recipe into a LayoutPlan."""
    _default_font = font_family or tokens.get("primitive", {}).get("font", {}).get("family", {}).get("body", "Microsoft YaHei")
    bboxes = _zone_bboxes(geometry.variables)
    elements_by_role = {e.role: e for e in slide.elements}
    nodes: list[PlannedNode] = []
    for idx, (zone_name, zone) in enumerate(recipe.zones.items()):
        bbox = bboxes.get(zone_name, BBox(0, 0, 100, 100))
        element = elements_by_role.get(zone_name)
        style: dict[str, Any] = {"font_family": _default_font, "size": 16.0}
        if element:
            raw_size = tokens.get("component", {}).get(zone_name, {}).get("size", 16)
            if isinstance(raw_size, dict) and "$ref" in raw_size:
                raw_size = _resolve_token(raw_size["$ref"].lstrip("$"), tokens)
            style["size"] = float(raw_size)
        nodes.append(
            PlannedNode(
                id=f"{slide.id}/{zone_name}",
                element_id=element.id if element else None,
                recipe_node_id=zone_name,
                kind=_normalize_node_kind(zone.kind),
                role=zone_name,
                geometry=GeometrySpec(bbox),
                resolved_style=style,
                content_binding=(element.content if element else {}) or {},
                z_order=idx,
            )
        )
    # Resolve background color from tokens if available.
    bg_color: str | None = None
    _bg_token = tokens.get("primitive", {}).get("palette", {}).get("bg")
    if isinstance(_bg_token, str):
        bg_color = _bg_token
    _bg_alt_token = tokens.get("primitive", {}).get("palette", {}).get("bg_alt")
    if bg_color is None and isinstance(_bg_alt_token, str):
        bg_color = _bg_alt_token

    return LayoutPlan(canvas=canvas, recipe_id=recipe.id, nodes=nodes, local_score=0.0, background_color=bg_color)


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
        recipes.append(_recipe(
            "bullets.sparse",
            "bullets",
            "sparse",
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
                {"terms": [{"var": "body.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.section"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"body.max_items": 4, "body.min_font_size": 18},
        ))
        recipes.append(_recipe(
            "bullets.with_icon",
            "bullets",
            "with_icon",
            {
                "title": {"kind": "text", "style": "component.title"},
                "icon_area": {"kind": "shape", "style": "component.body"},
                "body": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Icon area — narrow left column for number/icon prefixes
                {"terms": [{"var": "icon_area.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "icon_area.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "icon_area.width"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "icon_area.right"}, {"var": "icon_area.left", "coef": -1}, {"var": "icon_area.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "icon_area.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Body — text list to the right of icon column
                {"terms": [{"var": "body.left"}, {"var": "icon_area.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"body.max_items": 7, "body.min_font_size": 15},
        ))
        recipes.append(_recipe(
            "bullets.dense",
            "bullets",
            "dense",
            {
                "title": {"kind": "text", "style": "component.title"},
                "body": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 36}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "body.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "body.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "body.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "body.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"body.max_items": 12, "body.min_font_size": 12},
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
        recipes.append(_recipe(
            "dashboard.sparse",
            "dashboard",
            "sparse",
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
                {"terms": [{"var": "lead.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.section"}},
                {"terms": [{"var": "lead.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "lead.height"}], "op": "==", "rhs": {"const": 180}},
                {"terms": [{"var": "lead.bottom"}, {"var": "lead.top", "coef": -1}, {"var": "lead.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "support.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "support.top"}, {"var": "lead.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.section"}},
                {"terms": [{"var": "support.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "support.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"lead.max_metrics": 3, "support.min_font_size": 16},
        ))
        recipes.append(_recipe(
            "dashboard.metric_hero",
            "dashboard",
            "metric_hero",
            {
                "title": {"kind": "text", "style": "component.title"},
                "hero_metric": {"kind": "text", "style": "component.metric"},
                "hero_label": {"kind": "text", "style": "component.body"},
                "metrics_row": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Hero metric — large KPI value centered
                {"terms": [{"var": "hero_metric.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero_metric.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.section"}},
                {"terms": [{"var": "hero_metric.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero_metric.height"}], "op": "==", "rhs": {"const": 120}},
                {"terms": [{"var": "hero_metric.bottom"}, {"var": "hero_metric.top", "coef": -1}, {"var": "hero_metric.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Hero label — description below the large KPI
                {"terms": [{"var": "hero_label.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero_label.top"}, {"var": "hero_metric.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "hero_label.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero_label.height"}], "op": "==", "rhs": {"const": 24}},
                {"terms": [{"var": "hero_label.bottom"}, {"var": "hero_label.top", "coef": -1}, {"var": "hero_label.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Metrics row — supporting metrics at bottom
                {"terms": [{"var": "metrics_row.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "metrics_row.top"}, {"var": "hero_label.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.section"}},
                {"terms": [{"var": "metrics_row.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "metrics_row.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"hero_metric.max_lines": 1, "hero_metric.min_font_size": 60, "metrics_row.max_items": 4, "metrics_row.min_font_size": 14},
        ))
        recipes.append(_recipe(
            "dashboard.dense",
            "dashboard",
            "dense",
            {
                "title": {"kind": "text", "style": "component.title"},
                "grid": {"kind": "text-list", "style": "component.metric"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 36}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "grid.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "grid.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "grid.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "grid.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"grid.max_metrics": 8, "grid.min_font_size": 12},
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
                {"terms": [{"var": "hero.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "hero.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "hero.top"}, {"var": "kicker.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "hero.height"}], "op": "==", "rhs": {"const": 200}},
                {"terms": [{"var": "hero.bottom"}, {"var": "hero.top", "coef": -1}, {"var": "hero.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "title.min_font_size": 36},
        ))
        recipes.append(_recipe(
            "cover.statement",
            "cover",
            "statement",
            {
                "statement": {"kind": "text", "style": "component.title"},
                "attribution": {"kind": "text", "style": "component.body"},
                "footer": {"kind": "text", "style": "component.body"},
            },
            [
                # Statement — large bold text centered, occupying top ~60%
                {"terms": [{"var": "statement.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "statement.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "statement.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "statement.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Attribution — source/author below statement
                {"terms": [{"var": "attribution.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "attribution.top"}, {"var": "statement.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.section"}},
                {"terms": [{"var": "attribution.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "attribution.height"}], "op": "==", "rhs": {"const": 24}},
                {"terms": [{"var": "attribution.bottom"}, {"var": "attribution.top", "coef": -1}, {"var": "attribution.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Footer — small text at bottom
                {"terms": [{"var": "footer.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "footer.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "footer.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "footer.height"}], "op": "==", "rhs": {"const": 20}},
            ],
            {"statement.max_lines": 4, "statement.min_font_size": 36, "attribution.max_lines": 1, "footer.max_lines": 1},
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
                {"terms": [{"var": "b.left"}, {"var": "a.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.left"}, {"var": "b.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "c.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Equal column widths: a.width == b.width == c.width.
                {"terms": [{"var": "a.width"}, {"var": "b.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.width"}, {"var": "c.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
        ))
        recipes.append(_recipe(
            "image_grid.grid4",
            "image_grid",
            "grid4",
            {
                "title": {"kind": "text", "style": "component.title"},
                "a": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "b": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "c": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "d": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Top-left image
                {"terms": [{"var": "a.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "a.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "a.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "a.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Top-right image
                {"terms": [{"var": "b.left"}, {"var": "a.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "b.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Bottom-left image
                {"terms": [{"var": "c.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "c.top"}, {"var": "a.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "c.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "c.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Bottom-right image
                {"terms": [{"var": "d.left"}, {"var": "c.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "d.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "d.top"}, {"var": "c.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "d.bottom"}, {"var": "c.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
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
        recipes.append(_recipe(
            "timeline.vertical_dense",
            "timeline",
            "vertical_dense",
            {
                "title": {"kind": "text", "style": "component.title"},
                "items": {"kind": "text-list", "style": "component.body"},
                "axis": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 32}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "axis.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "axis.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "axis.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "axis.width"}], "op": "==", "rhs": {"const": 2}},
                {"terms": [{"var": "items.left"}, {"var": "axis.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "items.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "items.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"items.max_items": 10, "items.min_font_size": 11},
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
        recipes.append(_recipe(
            "comparison.matrix",
            "comparison",
            "matrix",
            {
                "title": {"kind": "text", "style": "component.title"},
                "left_header": {"kind": "text", "style": "component.title"},
                "right_header": {"kind": "text", "style": "component.title"},
                "criteria": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Left header — left half of header row
                {"terms": [{"var": "left_header.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "left_header.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "left_header.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "left_header.height"}], "op": "==", "rhs": {"const": 36}},
                {"terms": [{"var": "left_header.bottom"}, {"var": "left_header.top", "coef": -1}, {"var": "left_header.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Right header — right half of header row
                {"terms": [{"var": "right_header.left"}, {"var": "left_header.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "right_header.top"}, {"var": "left_header.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "right_header.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "right_header.height"}], "op": "==", "rhs": {"const": 36}},
                {"terms": [{"var": "right_header.bottom"}, {"var": "right_header.top", "coef": -1}, {"var": "right_header.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Criteria — full-width text list below headers
                {"terms": [{"var": "criteria.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "criteria.top"}, {"var": "left_header.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "criteria.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "criteria.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"criteria.max_items": 8, "criteria.min_font_size": 14, "left_header.max_lines": 2, "right_header.max_lines": 2},
        ))
        recipes.append(_recipe(
            "comparison.three_column",
            "comparison",
            "three_column",
            {
                "title": {"kind": "text", "style": "component.title"},
                "left": {"kind": "text", "style": "component.body"},
                "center": {"kind": "text", "style": "component.body"},
                "right": {"kind": "text", "style": "component.body"},
                "divider_a": {"kind": "shape", "style": "component.body"},
                "divider_b": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Left column
                {"terms": [{"var": "left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "left.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Center column
                {"terms": [{"var": "center.top"}, {"var": "left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "center.bottom"}, {"var": "left.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Right column
                {"terms": [{"var": "right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "right.top"}, {"var": "left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "right.bottom"}, {"var": "left.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Equal column widths
                {"terms": [{"var": "left.width"}, {"var": "center.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "center.width"}, {"var": "right.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Dividers
                {"terms": [{"var": "divider_a.left"}, {"var": "left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "divider_a.width"}], "op": "==", "rhs": {"const": 1}},
                {"terms": [{"var": "divider_a.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "divider_a.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "center.left"}, {"var": "divider_a.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "divider_b.left"}, {"var": "center.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "divider_b.width"}], "op": "==", "rhs": {"const": 1}},
                {"terms": [{"var": "divider_b.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "divider_b.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "right.left"}, {"var": "divider_b.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
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
        recipes.append(_recipe(
            "process.milestone",
            "process",
            "milestone",
            {
                "title": {"kind": "text", "style": "component.title"},
                "milestones": {"kind": "text-list", "style": "component.body"},
                "axis": {"kind": "shape", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 48}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Axis — horizontal timeline line at vertical center
                {"terms": [{"var": "axis.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "axis.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "axis.top"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "axis.height"}], "op": "==", "rhs": {"const": 4}},
                {"terms": [{"var": "axis.bottom"}, {"var": "axis.top", "coef": -1}, {"var": "axis.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Milestones — text list above and below axis
                {"terms": [{"var": "milestones.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "milestones.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "milestones.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "milestones.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"milestones.max_items": 6, "milestones.min_font_size": 14},
        ))
        recipes.append(_recipe(
            "process.horizontal_dense",
            "process",
            "horizontal_dense",
            {
                "title": {"kind": "text", "style": "component.title"},
                "steps": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 32}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "steps.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "steps.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "steps.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "steps.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"steps.max_items": 8, "steps.min_font_size": 11},
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
        recipes.append(_recipe(
            "table.wide",
            "table",
            "wide",
            {
                "title": {"kind": "text", "style": "component.title"},
                "table": {"kind": "table", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 32}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Table uses minimal margins for wider column space
                {"terms": [{"var": "table.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "table.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "table.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "table.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"table.min_col_width": 80, "table.min_font_size": 11},
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
    elif role == "matrix":
        recipes.append(_recipe(
            "matrix.quadrant",
            "matrix",
            "quadrant",
            {
                "title": {"kind": "text", "style": "component.title"},
                "top_left": {"kind": "text", "style": "component.body"},
                "top_right": {"kind": "text", "style": "component.body"},
                "bottom_left": {"kind": "text", "style": "component.body"},
                "bottom_right": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Top-left quadrant
                {"terms": [{"var": "top_left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "top_left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "top_left.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "top_left.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Top-right quadrant
                {"terms": [{"var": "top_right.left"}, {"var": "top_left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "top_right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "top_right.top"}, {"var": "top_left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "top_right.bottom"}, {"var": "top_left.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Bottom-left quadrant
                {"terms": [{"var": "bottom_left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "bottom_left.top"}, {"var": "top_left.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "bottom_left.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "bottom_left.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Bottom-right quadrant
                {"terms": [{"var": "bottom_right.left"}, {"var": "bottom_left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "bottom_right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "bottom_right.top"}, {"var": "bottom_left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "bottom_right.bottom"}, {"var": "bottom_left.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "top_left.max_lines": 4, "top_right.max_lines": 4, "bottom_left.max_lines": 4, "bottom_right.max_lines": 4},
        ))
        recipes.append(_recipe(
            "matrix.labeled",
            "matrix",
            "labeled",
            {
                "title": {"kind": "text", "style": "component.title"},
                "col_label_left": {"kind": "text", "style": "component.body"},
                "col_label_right": {"kind": "text", "style": "component.body"},
                "row_label_top": {"kind": "text", "style": "component.body"},
                "row_label_bottom": {"kind": "text", "style": "component.body"},
                "top_left": {"kind": "text", "style": "component.body"},
                "top_right": {"kind": "text", "style": "component.body"},
                "bottom_left": {"kind": "text", "style": "component.body"},
                "bottom_right": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Column labels
                {"terms": [{"var": "col_label_left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "col_label_left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "col_label_left.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "col_label_left.height"}], "op": "==", "rhs": {"const": 24}},
                {"terms": [{"var": "col_label_right.left"}, {"var": "col_label_left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "col_label_right.top"}, {"var": "col_label_left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "col_label_right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "col_label_right.height"}], "op": "==", "rhs": {"const": 24}},
                # Row labels (left column, narrow)
                {"terms": [{"var": "row_label_top.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "row_label_top.top"}, {"var": "col_label_left.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "row_label_top.width"}], "op": "==", "rhs": {"const": 80}},
                {"terms": [{"var": "row_label_top.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "row_label_bottom.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "row_label_bottom.top"}, {"var": "row_label_top.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "row_label_bottom.width"}], "op": "==", "rhs": {"const": 80}},
                {"terms": [{"var": "row_label_bottom.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Quadrant cells
                {"terms": [{"var": "top_left.left"}, {"var": "row_label_top.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "top_left.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "top_left.top"}, {"var": "col_label_left.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "top_left.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "top_right.left"}, {"var": "top_left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "top_right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "top_right.top"}, {"var": "top_left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "top_right.bottom"}, {"var": "top_left.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "bottom_left.left"}, {"var": "row_label_bottom.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "bottom_left.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "bottom_left.top"}, {"var": "top_left.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "bottom_left.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "bottom_right.left"}, {"var": "bottom_left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "bottom_right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "bottom_right.top"}, {"var": "bottom_left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "bottom_right.bottom"}, {"var": "bottom_left.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "top_left.max_lines": 3, "top_right.max_lines": 3, "bottom_left.max_lines": 3, "bottom_right.max_lines": 3},
        ))
    elif role == "kpi_hero":
        recipes.append(_recipe(
            "kpi_hero.split",
            "kpi_hero",
            "split",
            {
                "title": {"kind": "text", "style": "component.title"},
                "kpi_value": {"kind": "text", "style": "component.title"},
                "kpi_label": {"kind": "text", "style": "component.body"},
                "chart": {"kind": "chart", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # KPI value (large number, left side)
                {"terms": [{"var": "kpi_value.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "kpi_value.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "kpi_value.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "kpi_value.height"}], "op": "==", "rhs": {"const": 80}},
                # KPI label below value
                {"terms": [{"var": "kpi_label.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "kpi_label.top"}, {"var": "kpi_value.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "kpi_label.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "kpi_label.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Chart area (right side)
                {"terms": [{"var": "chart.left"}, {"var": "kpi_value.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "chart.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "chart.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "chart.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"kpi_value.max_lines": 1, "kpi_value.min_font_size": 48},
        ))
        recipes.append(_recipe(
            "kpi_hero.full",
            "kpi_hero",
            "full",
            {
                "title": {"kind": "text", "style": "component.title"},
                "kpi_value": {"kind": "text", "style": "component.title"},
                "kpi_label": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # KPI value centered, large
                {"terms": [{"var": "kpi_value.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "kpi_value.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "kpi_value.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "kpi_value.height"}], "op": "==", "rhs": {"const": 120}},
                # KPI label below
                {"terms": [{"var": "kpi_label.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "kpi_label.top"}, {"var": "kpi_value.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "kpi_label.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "kpi_label.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"kpi_value.max_lines": 1, "kpi_value.min_font_size": 72},
        ))
    elif role == "faq":
        recipes.append(_recipe(
            "faq.alternating",
            "faq",
            "alternating",
            {
                "title": {"kind": "text", "style": "component.title"},
                "questions": {"kind": "text-list", "style": "component.body"},
                "answers": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Questions on left half
                {"terms": [{"var": "questions.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "questions.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "questions.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "questions.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Answers on right half
                {"terms": [{"var": "answers.left"}, {"var": "questions.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "answers.top"}, {"var": "questions.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "answers.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "answers.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"questions.max_items": 6, "questions.min_font_size": 14, "answers.max_items": 6, "answers.min_font_size": 14},
        ))
        recipes.append(_recipe(
            "faq.stacked",
            "faq",
            "stacked",
            {
                "title": {"kind": "text", "style": "component.title"},
                "questions": {"kind": "text-list", "style": "component.body"},
                "answers": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Questions spanning full width, top half
                {"terms": [{"var": "questions.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "questions.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "questions.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "questions.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Answers spanning full width, bottom half
                {"terms": [{"var": "answers.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "answers.top"}, {"var": "questions.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "answers.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "answers.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"questions.max_items": 4, "questions.min_font_size": 14, "answers.max_items": 4, "answers.min_font_size": 14},
        ))
    elif role == "testimonial":
        recipes.append(_recipe(
            "testimonial.centered",
            "testimonial",
            "centered",
            {
                "quote_text": {"kind": "text", "style": "component.title"},
                "attribution": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "quote_text.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "quote_text.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "quote_text.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "quote_text.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "attribution.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "attribution.top"}, {"var": "quote_text.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "attribution.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "attribution.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "attribution.height"}], "op": "==", "rhs": {"const": 24}},
            ],
            {"quote_text.max_lines": 4, "quote_text.min_font_size": 24},
        ))
        recipes.append(_recipe(
            "testimonial.card",
            "testimonial",
            "card",
            {
                "quote_text": {"kind": "text", "style": "component.title"},
                "attribution": {"kind": "text", "style": "component.body"},
                "card_bg": {"kind": "shape", "style": "component.body"},
            },
            [
                # Card background
                {"terms": [{"var": "card_bg.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "card_bg.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "card_bg.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "card_bg.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Quote text inside card with padding
                {"terms": [{"var": "quote_text.left"}], "op": ">=", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "quote_text.top"}], "op": ">=", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "quote_text.right"}], "op": "<=", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "quote_text.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Attribution below
                {"terms": [{"var": "attribution.left"}], "op": ">=", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "attribution.top"}, {"var": "quote_text.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "attribution.right"}], "op": "<=", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "attribution.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "attribution.height"}], "op": "==", "rhs": {"const": 24}},
            ],
            {"quote_text.max_lines": 4, "quote_text.min_font_size": 22},
        ))
    elif role == "logo_wall":
        recipes.append(_recipe(
            "logo_wall.grid3",
            "logo_wall",
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
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Three logo columns
                {"terms": [{"var": "a.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "a.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "a.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "b.left"}, {"var": "a.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.left"}, {"var": "b.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "c.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "c.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Equal column widths
                {"terms": [{"var": "a.width"}, {"var": "b.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.width"}, {"var": "c.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
        ))
        recipes.append(_recipe(
            "logo_wall.grid4",
            "logo_wall",
            "grid4",
            {
                "title": {"kind": "text", "style": "component.title"},
                "a": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "b": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "c": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
                "d": {"kind": "image", "style": "component.hero", "fit": "smart-cover"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # 2x2 grid of logos
                {"terms": [{"var": "a.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "a.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "a.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "a.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "b.left"}, {"var": "a.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "b.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "b.top"}, {"var": "a.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "b.bottom"}, {"var": "a.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "c.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "c.top"}, {"var": "a.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "c.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "c.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                {"terms": [{"var": "d.left"}, {"var": "c.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "d.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "d.top"}, {"var": "c.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "d.bottom"}, {"var": "c.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
        ))
    elif role == "swot":
        recipes.append(_recipe(
            "swot.quadrant",
            "swot",
            "quadrant",
            {
                "title": {"kind": "text", "style": "component.title"},
                "strengths": {"kind": "text-list", "style": "component.body"},
                "weaknesses": {"kind": "text-list", "style": "component.body"},
                "opportunities": {"kind": "text-list", "style": "component.body"},
                "threats": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Strengths — top-left quadrant
                {"terms": [{"var": "strengths.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "strengths.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "strengths.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "strengths.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Weaknesses — top-right quadrant
                {"terms": [{"var": "weaknesses.left"}, {"var": "strengths.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "weaknesses.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "weaknesses.top"}, {"var": "strengths.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "weaknesses.bottom"}, {"var": "strengths.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Opportunities — bottom-left quadrant
                {"terms": [{"var": "opportunities.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "opportunities.top"}, {"var": "strengths.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "opportunities.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "opportunities.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Threats — bottom-right quadrant
                {"terms": [{"var": "threats.left"}, {"var": "opportunities.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "threats.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "threats.top"}, {"var": "opportunities.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "threats.bottom"}, {"var": "opportunities.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "strengths.max_items": 5, "weaknesses.max_items": 5, "opportunities.max_items": 5, "threats.max_items": 5},
        ))
        recipes.append(_recipe(
            "swot.labeled",
            "swot",
            "labeled",
            {
                "title": {"kind": "text", "style": "component.title"},
                "col_label_left": {"kind": "text", "style": "component.body"},
                "col_label_right": {"kind": "text", "style": "component.body"},
                "strengths": {"kind": "text-list", "style": "component.body"},
                "weaknesses": {"kind": "text-list", "style": "component.body"},
                "opportunities": {"kind": "text-list", "style": "component.body"},
                "threats": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Column labels
                {"terms": [{"var": "col_label_left.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "col_label_left.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "col_label_left.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "col_label_left.height"}], "op": "==", "rhs": {"const": 24}},
                {"terms": [{"var": "col_label_right.left"}, {"var": "col_label_left.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "col_label_right.top"}, {"var": "col_label_left.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "col_label_right.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "col_label_right.height"}], "op": "==", "rhs": {"const": 24}},
                # Strengths — top-left
                {"terms": [{"var": "strengths.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "strengths.top"}, {"var": "col_label_left.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "strengths.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "strengths.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Weaknesses — top-right
                {"terms": [{"var": "weaknesses.left"}, {"var": "strengths.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "weaknesses.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "weaknesses.top"}, {"var": "strengths.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "weaknesses.bottom"}, {"var": "strengths.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Opportunities — bottom-left
                {"terms": [{"var": "opportunities.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "opportunities.top"}, {"var": "strengths.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "opportunities.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "opportunities.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Threats — bottom-right
                {"terms": [{"var": "threats.left"}, {"var": "opportunities.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "threats.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "threats.top"}, {"var": "opportunities.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "threats.bottom"}, {"var": "opportunities.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "strengths.max_items": 5, "weaknesses.max_items": 5, "opportunities.max_items": 5, "threats.max_items": 5},
        ))
    elif role == "porter":
        recipes.append(_recipe(
            "porter.diamond",
            "porter",
            "diamond",
            {
                "title": {"kind": "text", "style": "component.title"},
                "rivalry": {"kind": "text-list", "style": "component.body"},
                "new_entrants": {"kind": "text-list", "style": "component.body"},
                "substitutes": {"kind": "text-list", "style": "component.body"},
                "buyer_power": {"kind": "text-list", "style": "component.body"},
                "supplier_power": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Rivalry — center zone, anchored to canvas center
                {"terms": [{"var": "rivalry.width"}], "op": "==", "rhs": {"const": 200}},
                {"terms": [{"var": "rivalry.height"}], "op": "==", "rhs": {"const": 120}},
                {"terms": [{"var": "rivalry.left"}], "op": "==", "rhs": {"const": 380}},
                {"terms": [{"var": "rivalry.top"}], "op": "==", "rhs": {"const": 210}},
                # New entrants — top zone
                {"terms": [{"var": "new_entrants.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "new_entrants.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "new_entrants.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "new_entrants.bottom"}, {"var": "rivalry.top", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                # Substitutes — bottom zone
                {"terms": [{"var": "substitutes.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "substitutes.top"}, {"var": "rivalry.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "substitutes.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "substitutes.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Buyer power — left zone
                {"terms": [{"var": "buyer_power.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "buyer_power.top"}, {"var": "new_entrants.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "buyer_power.right"}, {"var": "rivalry.left", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "buyer_power.bottom"}, {"var": "substitutes.top", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                # Supplier power — right zone
                {"terms": [{"var": "supplier_power.left"}, {"var": "rivalry.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "supplier_power.top"}, {"var": "buyer_power.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "supplier_power.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "supplier_power.bottom"}, {"var": "buyer_power.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "rivalry.max_items": 4, "new_entrants.max_items": 4, "substitutes.max_items": 4, "buyer_power.max_items": 4, "supplier_power.max_items": 4},
        ))
        recipes.append(_recipe(
            "porter.horizontal",
            "porter",
            "horizontal",
            {
                "title": {"kind": "text", "style": "component.title"},
                "rivalry": {"kind": "text-list", "style": "component.body"},
                "new_entrants": {"kind": "text-list", "style": "component.body"},
                "substitutes": {"kind": "text-list", "style": "component.body"},
                "buyer_power": {"kind": "text-list", "style": "component.body"},
                "supplier_power": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # 5 equal columns layout
                # new_entrants — col1
                {"terms": [{"var": "new_entrants.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "new_entrants.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "new_entrants.width"}], "op": "==", "rhs": {"const": 160}},
                {"terms": [{"var": "new_entrants.right"}, {"var": "new_entrants.left", "coef": -1}, {"var": "new_entrants.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "new_entrants.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # buyer_power — col2
                {"terms": [{"var": "buyer_power.left"}, {"var": "new_entrants.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "buyer_power.width"}], "op": "==", "rhs": {"const": 160}},
                {"terms": [{"var": "buyer_power.right"}, {"var": "buyer_power.left", "coef": -1}, {"var": "buyer_power.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "buyer_power.top"}, {"var": "new_entrants.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "buyer_power.bottom"}, {"var": "new_entrants.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # rivalry — col3 (center, wider)
                {"terms": [{"var": "rivalry.left"}, {"var": "buyer_power.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "rivalry.width"}], "op": "==", "rhs": {"const": 200}},
                {"terms": [{"var": "rivalry.right"}, {"var": "rivalry.left", "coef": -1}, {"var": "rivalry.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "rivalry.top"}, {"var": "new_entrants.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "rivalry.bottom"}, {"var": "new_entrants.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # supplier_power — col4
                {"terms": [{"var": "supplier_power.left"}, {"var": "rivalry.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "supplier_power.width"}], "op": "==", "rhs": {"const": 160}},
                {"terms": [{"var": "supplier_power.right"}, {"var": "supplier_power.left", "coef": -1}, {"var": "supplier_power.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "supplier_power.top"}, {"var": "new_entrants.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "supplier_power.bottom"}, {"var": "new_entrants.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # substitutes — col5
                {"terms": [{"var": "substitutes.left"}, {"var": "supplier_power.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "substitutes.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "substitutes.top"}, {"var": "new_entrants.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "substitutes.bottom"}, {"var": "new_entrants.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "rivalry.max_items": 4, "new_entrants.max_items": 4, "substitutes.max_items": 4, "buyer_power.max_items": 4, "supplier_power.max_items": 4},
        ))
    elif role == "pest":
        recipes.append(_recipe(
            "pest.grid",
            "pest",
            "grid",
            {
                "title": {"kind": "text", "style": "component.title"},
                "political": {"kind": "text-list", "style": "component.body"},
                "economic": {"kind": "text-list", "style": "component.body"},
                "social": {"kind": "text-list", "style": "component.body"},
                "technological": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Political — top-left quadrant
                {"terms": [{"var": "political.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "political.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "political.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "political.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                # Economic — top-right quadrant
                {"terms": [{"var": "economic.left"}, {"var": "political.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "economic.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "economic.top"}, {"var": "political.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "economic.bottom"}, {"var": "political.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Social — bottom-left quadrant
                {"terms": [{"var": "social.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "social.top"}, {"var": "political.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "social.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "social.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Technological — bottom-right quadrant
                {"terms": [{"var": "technological.left"}, {"var": "social.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "technological.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "technological.top"}, {"var": "social.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "technological.bottom"}, {"var": "social.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 2, "political.max_items": 5, "economic.max_items": 5, "social.max_items": 5, "technological.max_items": 5},
        ))
        recipes.append(_recipe(
            "pest.vertical",
            "pest",
            "vertical",
            {
                "title": {"kind": "text", "style": "component.title"},
                "political": {"kind": "text-list", "style": "component.body"},
                "economic": {"kind": "text-list", "style": "component.body"},
                "social": {"kind": "text-list", "style": "component.body"},
                "technological": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Political — row 1
                {"terms": [{"var": "political.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "political.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "political.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "political.height"}], "op": "==", "rhs": {"const": 90}},
                {"terms": [{"var": "political.bottom"}, {"var": "political.top", "coef": -1}, {"var": "political.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Economic — row 2
                {"terms": [{"var": "economic.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "economic.top"}, {"var": "political.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "economic.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "economic.height"}], "op": "==", "rhs": {"const": 90}},
                {"terms": [{"var": "economic.bottom"}, {"var": "economic.top", "coef": -1}, {"var": "economic.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Social — row 3
                {"terms": [{"var": "social.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "social.top"}, {"var": "economic.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "social.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "social.height"}], "op": "==", "rhs": {"const": 90}},
                {"terms": [{"var": "social.bottom"}, {"var": "social.top", "coef": -1}, {"var": "social.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Technological — row 4
                {"terms": [{"var": "technological.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "technological.top"}, {"var": "social.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "technological.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "technological.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "political.max_items": 5, "economic.max_items": 5, "social.max_items": 5, "technological.max_items": 5},
        ))
    elif role == "bmc":
        recipes.append(_recipe(
            "bmc.canvas",
            "bmc",
            "canvas",
            {
                "title": {"kind": "text", "style": "component.title"},
                "key_partners": {"kind": "text-list", "style": "component.body"},
                "key_activities": {"kind": "text-list", "style": "component.body"},
                "key_resources": {"kind": "text-list", "style": "component.body"},
                "value_propositions": {"kind": "text-list", "style": "component.body"},
                "customer_relationships": {"kind": "text-list", "style": "component.body"},
                "channels": {"kind": "text-list", "style": "component.body"},
                "customer_segments": {"kind": "text-list", "style": "component.body"},
                "cost_structure": {"kind": "text-list", "style": "component.body"},
                "revenue_streams": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 36}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Col1: key_partners (full main height)
                {"terms": [{"var": "key_partners.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "key_partners.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                # Equal column widths
                {"terms": [{"var": "key_activities.width"}, {"var": "key_partners.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "value_propositions.width"}, {"var": "key_partners.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_relationships.width"}, {"var": "key_partners.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_segments.width"}, {"var": "key_partners.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Column gaps
                {"terms": [{"var": "key_activities.left"}, {"var": "key_partners.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "value_propositions.left"}, {"var": "key_activities.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "customer_relationships.left"}, {"var": "value_propositions.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "customer_segments.left"}, {"var": "customer_relationships.right", "coef": -1}], "op": "==", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "customer_segments.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                # key_partners: full main height
                {"terms": [{"var": "key_partners.bottom"}, {"var": "cost_structure.top", "coef": -1}], "op": "<=", "rhs": {"token": "semantic.space.component"}},
                # key_activities: Col2 top half
                {"terms": [{"var": "key_activities.top"}, {"var": "key_partners.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "key_activities.bottom", "coef": 2}, {"var": "key_partners.top", "coef": -1}, {"var": "key_partners.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # key_resources: Col2 bottom half
                {"terms": [{"var": "key_resources.left"}, {"var": "key_activities.left", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "key_resources.width"}, {"var": "key_activities.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "key_resources.top"}, {"var": "key_activities.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "key_resources.bottom"}, {"var": "key_partners.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # value_propositions: Col3 full height
                {"terms": [{"var": "value_propositions.top"}, {"var": "key_partners.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "value_propositions.bottom"}, {"var": "key_partners.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # customer_relationships: Col4 top half
                {"terms": [{"var": "customer_relationships.top"}, {"var": "key_partners.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_relationships.bottom", "coef": 2}, {"var": "key_partners.top", "coef": -1}, {"var": "key_partners.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # channels: Col4 bottom half
                {"terms": [{"var": "channels.left"}, {"var": "customer_relationships.left", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "channels.width"}, {"var": "customer_relationships.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "channels.top"}, {"var": "customer_relationships.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "channels.bottom"}, {"var": "key_partners.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # customer_segments: Col5 full height
                {"terms": [{"var": "customer_segments.top"}, {"var": "key_partners.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_segments.bottom"}, {"var": "key_partners.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Bottom row: cost_structure + revenue_streams
                {"terms": [{"var": "cost_structure.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "cost_structure.top"}, {"var": "key_partners.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "cost_structure.height"}], "op": "==", "rhs": {"const": 72}},
                {"terms": [{"var": "cost_structure.bottom"}, {"var": "cost_structure.top", "coef": -1}, {"var": "cost_structure.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "cost_structure.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "revenue_streams.left"}, {"var": "cost_structure.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "revenue_streams.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "revenue_streams.top"}, {"var": "cost_structure.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "revenue_streams.bottom"}, {"var": "cost_structure.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "cost_structure.bottom"}], "op": "<=", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 1, "key_partners.max_items": 4, "key_activities.max_items": 3, "key_resources.max_items": 3, "value_propositions.max_items": 4, "customer_relationships.max_items": 3, "channels.max_items": 3, "customer_segments.max_items": 4, "cost_structure.max_items": 3, "revenue_streams.max_items": 3},
        ))
        recipes.append(_recipe(
            "bmc.compact",
            "bmc",
            "compact",
            {
                "title": {"kind": "text", "style": "component.title"},
                "key_partners": {"kind": "text-list", "style": "component.body"},
                "key_activities": {"kind": "text-list", "style": "component.body"},
                "key_resources": {"kind": "text-list", "style": "component.body"},
                "value_propositions": {"kind": "text-list", "style": "component.body"},
                "customer_relationships": {"kind": "text-list", "style": "component.body"},
                "channels": {"kind": "text-list", "style": "component.body"},
                "customer_segments": {"kind": "text-list", "style": "component.body"},
                "cost_structure": {"kind": "text-list", "style": "component.body"},
                "revenue_streams": {"kind": "text-list", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 36}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # 3-column layout: left / center / right
                # Left column: key_partners (top) + key_activities (bottom)
                {"terms": [{"var": "key_partners.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "key_partners.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "key_partners.right"}], "op": "==", "rhs": {"const": 312}},
                {"terms": [{"var": "key_partners.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "key_activities.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "key_activities.top"}, {"var": "key_partners.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "key_activities.right"}], "op": "==", "rhs": {"const": 312}},
                {"terms": [{"var": "key_activities.bottom"}], "op": "==", "rhs": {"ref": "canvas.safe_bottom"}},
                # Center column: key_resources (top) + value_propositions (middle) + customer_relationships (bottom)
                {"terms": [{"var": "key_resources.left"}], "op": "==", "rhs": {"const": 328}},
                {"terms": [{"var": "key_resources.top"}, {"var": "key_partners.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "key_resources.right"}], "op": "==", "rhs": {"const": 600}},
                {"terms": [{"var": "key_resources.height"}], "op": "==", "rhs": {"const": 100}},
                {"terms": [{"var": "key_resources.bottom"}, {"var": "key_resources.top", "coef": -1}, {"var": "key_resources.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "value_propositions.left"}, {"var": "key_resources.left", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "value_propositions.top"}, {"var": "key_resources.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "value_propositions.right"}, {"var": "key_resources.right", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "value_propositions.height"}], "op": "==", "rhs": {"const": 100}},
                {"terms": [{"var": "value_propositions.bottom"}, {"var": "value_propositions.top", "coef": -1}, {"var": "value_propositions.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_relationships.left"}, {"var": "key_resources.left", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_relationships.top"}, {"var": "value_propositions.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "customer_relationships.right"}, {"var": "key_resources.right", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_relationships.bottom"}, {"var": "key_activities.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Right column: channels (top) + customer_segments (bottom)
                {"terms": [{"var": "channels.left"}], "op": "==", "rhs": {"const": 616}},
                {"terms": [{"var": "channels.top"}, {"var": "key_partners.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "channels.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "channels.bottom"}], "op": "==", "rhs": {"ref": "canvas.center_y"}},
                {"terms": [{"var": "customer_segments.left"}, {"var": "channels.left", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_segments.top"}, {"var": "channels.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "customer_segments.right"}, {"var": "channels.right", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "customer_segments.bottom"}, {"var": "key_activities.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Bottom row: cost_structure (left 50%) + revenue_streams (right 50%)
                # These overlay the bottom of the grid
                {"terms": [{"var": "cost_structure.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "cost_structure.top"}, {"var": "customer_relationships.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "cost_structure.right"}], "op": "==", "rhs": {"ref": "canvas.center_x"}},
                {"terms": [{"var": "cost_structure.bottom"}, {"var": "customer_segments.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "revenue_streams.left"}, {"var": "cost_structure.right", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "revenue_streams.top"}, {"var": "cost_structure.top", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "revenue_streams.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "revenue_streams.bottom"}, {"var": "cost_structure.bottom", "coef": -1}], "op": "==", "rhs": {"const": 0}},
            ],
            {"title.max_lines": 1, "key_partners.max_items": 3, "key_activities.max_items": 3, "key_resources.max_items": 3, "value_propositions.max_items": 3, "customer_relationships.max_items": 3, "channels.max_items": 3, "customer_segments.max_items": 3, "cost_structure.max_items": 3, "revenue_streams.max_items": 3},
        ))
    elif role == "funnel":
        recipes.append(_recipe(
            "funnel.stacked",
            "funnel",
            "stacked",
            {
                "title": {"kind": "text", "style": "component.title"},
                "stage1": {"kind": "text", "style": "component.body"},
                "stage2": {"kind": "text", "style": "component.body"},
                "stage3": {"kind": "text", "style": "component.body"},
                "stage4": {"kind": "text", "style": "component.body"},
                "stage5": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Safe area width for reference
                # stage1 — 90% width, centered
                {"terms": [{"var": "stage1.width"}], "op": "==", "rhs": {"const": 783}},
                {"terms": [{"var": "stage1.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage1.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage1.right"}, {"var": "stage1.left", "coef": -1}, {"var": "stage1.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage1.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage1.height"}], "op": "==", "rhs": {"const": 64}},
                {"terms": [{"var": "stage1.bottom"}, {"var": "stage1.top", "coef": -1}, {"var": "stage1.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage2 — 75% width
                {"terms": [{"var": "stage2.width"}], "op": "==", "rhs": {"const": 652}},
                {"terms": [{"var": "stage2.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage2.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage2.right"}, {"var": "stage2.left", "coef": -1}, {"var": "stage2.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage2.top"}, {"var": "stage1.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage2.height"}], "op": "==", "rhs": {"const": 64}},
                {"terms": [{"var": "stage2.bottom"}, {"var": "stage2.top", "coef": -1}, {"var": "stage2.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage3 — 60% width
                {"terms": [{"var": "stage3.width"}], "op": "==", "rhs": {"const": 522}},
                {"terms": [{"var": "stage3.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage3.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage3.right"}, {"var": "stage3.left", "coef": -1}, {"var": "stage3.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage3.top"}, {"var": "stage2.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage3.height"}], "op": "==", "rhs": {"const": 64}},
                {"terms": [{"var": "stage3.bottom"}, {"var": "stage3.top", "coef": -1}, {"var": "stage3.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage4 — 45% width
                {"terms": [{"var": "stage4.width"}], "op": "==", "rhs": {"const": 392}},
                {"terms": [{"var": "stage4.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage4.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage4.right"}, {"var": "stage4.left", "coef": -1}, {"var": "stage4.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage4.top"}, {"var": "stage3.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage4.height"}], "op": "==", "rhs": {"const": 64}},
                {"terms": [{"var": "stage4.bottom"}, {"var": "stage4.top", "coef": -1}, {"var": "stage4.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage5 — 30% width
                {"terms": [{"var": "stage5.width"}], "op": "==", "rhs": {"const": 261}},
                {"terms": [{"var": "stage5.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage5.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage5.right"}, {"var": "stage5.left", "coef": -1}, {"var": "stage5.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage5.top"}, {"var": "stage4.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage5.height"}], "op": "==", "rhs": {"const": 64}},
                {"terms": [{"var": "stage5.bottom"}, {"var": "stage5.top", "coef": -1}, {"var": "stage5.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Must fit within safe area
                {"terms": [{"var": "stage5.bottom"}], "op": "<=", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "stage1.max_lines": 2, "stage2.max_lines": 2, "stage3.max_lines": 2, "stage4.max_lines": 2, "stage5.max_lines": 2},
        ))
        recipes.append(_recipe(
            "funnel.narrow",
            "funnel",
            "narrow",
            {
                "title": {"kind": "text", "style": "component.title"},
                "stage1": {"kind": "text", "style": "component.body"},
                "stage2": {"kind": "text", "style": "component.body"},
                "stage3": {"kind": "text", "style": "component.body"},
            },
            [
                {"terms": [{"var": "title.left"}], "op": "==", "rhs": {"ref": "canvas.safe_left"}},
                {"terms": [{"var": "title.top"}], "op": "==", "rhs": {"ref": "canvas.safe_top"}},
                {"terms": [{"var": "title.right"}], "op": "==", "rhs": {"ref": "canvas.safe_right"}},
                {"terms": [{"var": "title.height"}], "op": "==", "rhs": {"const": 40}},
                {"terms": [{"var": "title.bottom"}, {"var": "title.top", "coef": -1}, {"var": "title.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage1 — 80% width, centered
                {"terms": [{"var": "stage1.width"}], "op": "==", "rhs": {"const": 696}},
                {"terms": [{"var": "stage1.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage1.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage1.right"}, {"var": "stage1.left", "coef": -1}, {"var": "stage1.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage1.top"}, {"var": "title.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage1.height"}], "op": "==", "rhs": {"const": 100}},
                {"terms": [{"var": "stage1.bottom"}, {"var": "stage1.top", "coef": -1}, {"var": "stage1.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage2 — 55% width
                {"terms": [{"var": "stage2.width"}], "op": "==", "rhs": {"const": 478}},
                {"terms": [{"var": "stage2.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage2.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage2.right"}, {"var": "stage2.left", "coef": -1}, {"var": "stage2.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage2.top"}, {"var": "stage1.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage2.height"}], "op": "==", "rhs": {"const": 100}},
                {"terms": [{"var": "stage2.bottom"}, {"var": "stage2.top", "coef": -1}, {"var": "stage2.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # stage3 — 30% width
                {"terms": [{"var": "stage3.width"}], "op": "==", "rhs": {"const": 261}},
                {"terms": [{"var": "stage3.left"}, {"var": "canvas.center_x", "coef": 1}, {"var": "stage3.width", "coef": -0.5}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage3.right"}, {"var": "stage3.left", "coef": -1}, {"var": "stage3.width", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                {"terms": [{"var": "stage3.top"}, {"var": "stage2.bottom", "coef": -1}], "op": ">=", "rhs": {"token": "semantic.space.component"}},
                {"terms": [{"var": "stage3.height"}], "op": "==", "rhs": {"const": 100}},
                {"terms": [{"var": "stage3.bottom"}, {"var": "stage3.top", "coef": -1}, {"var": "stage3.height", "coef": -1}], "op": "==", "rhs": {"const": 0}},
                # Must fit within safe area
                {"terms": [{"var": "stage3.bottom"}], "op": "<=", "rhs": {"ref": "canvas.safe_bottom"}},
            ],
            {"title.max_lines": 2, "stage1.max_lines": 3, "stage2.max_lines": 3, "stage3.max_lines": 3},
        ))
    return recipes
