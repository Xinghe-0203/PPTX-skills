"""TemplateProfileV2 schema and compiled-bundle models (PR7).

This module defines the source profile that the V2 template compiler consumes
and the compiled bundle it produces. Token resolution uses a three-layer model
(primitive -> semantic -> component) with ``$ref`` aliases. The schema rejects
unknown fields outside ``extensions`` and validates colors, sizes and grammar.
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from typing import Any

from pptx_skill.content_model import CanvasSpec, SafeInsets

TEMPLATE_SCHEMA_VERSION = 2

REQUIRED_ROLES = (
    "cover",
    "toc",
    "section",
    "bullets",
    "text_image",
    "full_image",
    "image_grid",
    "dashboard",
    "timeline",
    "comparison",
    "quote",
    "process",
    "table",
    "end",
)

VALID_GRAMMAR_FAMILIES = {"standard", "editorial_grid", "technical_axis", "poster_column"}
VALID_GRAMMAR_NODES = {
    "grid", "axis", "column", "stack", "split", "overlay", "rail", "full-bleed",
    "text", "text-list", "metric", "table", "chart", "image", "line", "shape", "group",
}

_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
_REF_RE = re.compile(r"^[a-zA-Z_][a-zA-Z0-9_.]*$")


# ---------------------------------------------------------------------------
# Token resolution
# ---------------------------------------------------------------------------


def _lookup(token: str, tokens: dict[str, Any]) -> Any:
    parts = token.split(".")
    value: Any = tokens
    for part in parts:
        if not isinstance(value, dict) or part not in value:
            raise ValueError(f"Token reference not found: {token}")
        value = value[part]
    return value


def resolve_token(token: str, tokens: dict[str, Any]) -> Any:
    """Resolve a token path, following ``$ref`` chains until a concrete value."""
    seen: set[str] = set()
    current = token
    for _ in range(64):
        if current in seen:
            raise ValueError(f"Circular token reference: {current}")
        seen.add(current)
        value = _lookup(current, tokens)
        if isinstance(value, dict) and "$ref" in value:
            ref = value["$ref"]
            if isinstance(ref, str):
                ref = ref.lstrip("$")
            current = ref
            continue
        return value
    raise ValueError(f"Token reference too deep: {token}")


def _topo_sort_tokens(tokens: dict[str, Any]) -> list[str]:
    """Return token paths in dependency order (refs before referencers)."""
    edges: dict[str, list[str]] = {}
    paths: list[str] = []

    def walk(node: Any, prefix: str) -> None:
        if isinstance(node, dict):
            if "$ref" in node:
                ref = node["$ref"]
                if isinstance(ref, str):
                    target = ref.lstrip("$")
                    edges.setdefault(prefix, []).append(target)
                    paths.append(prefix)
                return
            for key, child in node.items():
                walk(child, f"{prefix}.{key}" if prefix else key)
        else:
            paths.append(prefix)

    walk(tokens, "")
    ordered: list[str] = []
    visited: set[str] = set()

    def visit(path: str, stack: set[str]) -> None:
        if path in visited:
            return
        if path in stack:
            raise ValueError(f"Circular token reference detected at {path}")
        stack.add(path)
        for dep in edges.get(path, []):
            if dep in paths:
                visit(dep, stack)
        stack.discard(path)
        visited.add(path)
        ordered.append(path)

    for path in paths:
        visit(path, set())
    return ordered


def resolve_all_tokens(tokens: dict[str, Any]) -> dict[str, Any]:
    """Return a fully-resolved copy of ``tokens`` with every ``$ref`` expanded."""
    ordered = _topo_sort_tokens(tokens)
    resolved: dict[str, Any] = copy.deepcopy(tokens)

    # Walk every leaf path in dependency order, replacing $ref dicts with values.
    def set_at(path: str, value: Any) -> None:
        parts = path.split(".")
        node = resolved
        for part in parts[:-1]:
            node = node[part]
        node[parts[-1]] = value

    def get_at(path: str) -> Any:
        parts = path.split(".")
        node: Any = resolved
        for part in parts:
            node = node[part]
        return node

    for path in ordered:
        value = get_at(path)
        if isinstance(value, dict) and "$ref" in value:
            ref = value["$ref"]
            if isinstance(ref, str):
                ref = ref.lstrip("$")
            concrete = resolve_token(ref, resolved)
            set_at(path, concrete)
    return resolved


# ---------------------------------------------------------------------------
# Schema validation
# ---------------------------------------------------------------------------


def _validate_color(path: str, value: Any) -> None:
    if not isinstance(value, str) or not _HEX_RE.match(value):
        raise ValueError(f"{path} must be a #RRGGBB color, got {value!r}")


def _validate_token_tree(path: str, node: Any) -> None:
    if isinstance(node, dict):
        if "$ref" in node:
            ref = node["$ref"]
            if not isinstance(ref, str) or not _REF_RE.match(ref.lstrip("$")):
                raise ValueError(f"{path}.$ref must be a dotted path, got {ref!r}")
            return
        for key, child in node.items():
            if not isinstance(key, str):
                raise ValueError(f"{path} has non-string key {key!r}")
            # Palette entries are always hex colors; enforce the shape.
            if path.endswith(".palette") or ".palette." in path:
                _validate_color(f"{path}.{key}", child)
                continue
            _validate_token_tree(f"{path}.{key}", child)
    elif not isinstance(node, (int, float, str, bool)):
        raise ValueError(f"{path} has unsupported token value type {type(node).__name__}")


def validate_source_profile(profile: dict) -> list[str]:
    """Validate a TemplateProfileV2 source profile.

    Returns a list of warnings; raises ``ValueError``
    on schema violations. Unknown keys outside ``extensions`` are rejected.
    """
    warnings: list[str] = []

    if not isinstance(profile, dict):
        raise TypeError("Template profile must be a JSON object")
    if profile.get("schema_version") != TEMPLATE_SCHEMA_VERSION:
        raise ValueError(
            f"schema_version must be {TEMPLATE_SCHEMA_VERSION}, got {profile.get('schema_version')!r}"
        )
    if profile.get("kind") != "template-source-profile":
        raise ValueError(f"kind must be 'template-source-profile', got {profile.get('kind')!r}")

    for required in ("id", "name", "description", "canvas", "tokens", "layouts", "coverage"):
        if required not in profile:
            raise ValueError(f"Template profile missing required field: {required}")

    if not isinstance(profile["id"], str) or not profile["id"]:
        raise ValueError("id must be a non-empty string")

    canvas = profile["canvas"]
    if not isinstance(canvas, dict):
        raise ValueError("canvas must be an object")
    preset = canvas.get("preset")
    if preset not in {"16:9", "4:3", "9:16", "custom"}:
        raise ValueError(f"canvas.preset must be 16:9 | 4:3 | 9:16 | custom, got {preset!r}")

    tokens = profile["tokens"]
    if not isinstance(tokens, dict):
        raise ValueError("tokens must be an object")
    for layer in ("primitive", "semantic", "component"):
        if layer not in tokens:
            raise ValueError(f"tokens.{layer} is required")
        _validate_token_tree(f"tokens.{layer}", tokens[layer])

    grammar = profile.get("grammar", {})
    if not isinstance(grammar, dict):
        raise ValueError("grammar must be an object")
    family = grammar.get("family", "standard")
    if family not in VALID_GRAMMAR_FAMILIES:
        raise ValueError(f"grammar.family must be one of {sorted(VALID_GRAMMAR_FAMILIES)}, got {family!r}")

    layouts = profile["layouts"]
    if not isinstance(layouts, dict):
        raise ValueError("layouts must be an object mapping role -> list[recipe_id]")
    for role, recipe_ids in layouts.items():
        if not isinstance(recipe_ids, list) or not recipe_ids:
            raise ValueError(f"layouts.{role} must be a non-empty list")
        if not all(isinstance(rid, str) for rid in recipe_ids):
            raise ValueError(f"layouts.{role} entries must be strings")

    coverage = profile["coverage"]
    if not isinstance(coverage, dict):
        raise ValueError("coverage must be an object")
    status = coverage.get("status")
    if status not in {"partial", "complete"}:
        raise ValueError(f"coverage.status must be partial|complete, got {status!r}")
    covered_roles = coverage.get("roles", [])
    if not isinstance(covered_roles, list):
        raise ValueError("coverage.roles must be a list")
    unknown = [r for r in covered_roles if r not in REQUIRED_ROLES]
    if unknown:
        raise ValueError(f"coverage.roles has unknown roles: {unknown}")

    rhythm = profile.get("rhythm", {})
    if not isinstance(rhythm, dict):
        raise ValueError("rhythm must be an object")
    for rule in rhythm.get("rules", []):
        if not isinstance(rule, dict):
            raise ValueError("rhythm.rules entries must be objects")
        if "after" not in rule:
            raise ValueError("rhythm.rules entries require an 'after' role")

    qa = profile.get("qa", {})
    if not isinstance(qa, dict):
        raise ValueError("qa must be an object")

    # Reject unknown top-level keys except extensions.
    allowed_keys = {
        "schema_version", "kind", "id", "name", "description", "canvas",
        "tokens", "grammar", "layouts", "coverage", "rhythm", "qa",
        "provenance", "extensions",
    }
    unknown_keys = set(profile.keys()) - allowed_keys
    if unknown_keys:
        raise ValueError(f"Unknown top-level keys: {sorted(unknown_keys)}")

    covered_in_layouts = set(profile["layouts"].keys())
    missing_roles = set(REQUIRED_ROLES) - covered_in_layouts
    if missing_roles:
        warnings.append(f"layouts missing required roles: {sorted(missing_roles)}")

    covered_in_coverage = set(coverage.get("roles", []))
    uncovered = covered_in_layouts - covered_in_coverage
    if uncovered:
        warnings.append(f"layouts has roles not declared in coverage.roles: {sorted(uncovered)}")

    return warnings


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass
class StyleIntent:
    """Structured intent produced by an agent or heuristic provider."""

    industry: str = ""
    audience: str = ""
    tone: list[str] = field(default_factory=list)
    density: str = "medium"
    image_style: str = ""
    geometry: list[str] = field(default_factory=list)
    avoid: list[str] = field(default_factory=list)
    brand_colors: list[str] = field(default_factory=list)
    canvas: str = "16:9"
    font_constraints: list[str] = field(default_factory=list)

    def normalize(self) -> dict:
        """Return a deterministic, comparable representation."""
        return {
            "industry": self.industry.strip().lower(),
            "audience": self.audience.strip().lower(),
            "tone": sorted(t.strip().lower() for t in self.tone),
            "density": self.density.strip().lower(),
            "image_style": self.image_style.strip().lower(),
            "geometry": sorted(g.strip().lower() for g in self.geometry),
            "avoid": sorted(a.strip().lower() for a in self.avoid),
            "brand_colors": [c.strip().upper() for c in self.brand_colors],
            "canvas": self.canvas.strip().lower(),
            "font_constraints": sorted(f.strip().lower() for f in self.font_constraints),
        }


@dataclass
class CompiledTemplateBundleV2:
    """A fully-resolved, complete template bundle ready for the catalog."""

    schema_version: int
    kind: str  # "compiled-template-bundle"
    id: str
    name: str
    description: str
    canvas: CanvasSpec
    tokens: dict[str, Any]  # fully resolved (no $ref)
    grammar: dict[str, Any]
    layouts: dict[str, list[str]]
    coverage: dict[str, Any]
    rhythm: dict[str, Any]
    qa: dict[str, Any]
    provenance: dict[str, Any]
    recipe_provenance: dict[str, str] = field(default_factory=dict)  # recipe_id -> generated|inherited|migrated

    def is_complete(self) -> bool:
        return self.coverage.get("status") == "complete" and all(
            role in self.layouts for role in REQUIRED_ROLES
        )


def bundle_to_dict(bundle: CompiledTemplateBundleV2) -> dict:
    return {
        "schema_version": bundle.schema_version,
        "kind": bundle.kind,
        "id": bundle.id,
        "name": bundle.name,
        "description": bundle.description,
        "canvas": {
            "preset": bundle.canvas.name,
            "width_pt": bundle.canvas.width_pt,
            "height_pt": bundle.canvas.height_pt,
            "safe_margin": {
                "top": bundle.canvas.safe.top,
                "right": bundle.canvas.safe.right,
                "bottom": bundle.canvas.safe.bottom,
                "left": bundle.canvas.safe.left,
            },
        },
        "tokens": bundle.tokens,
        "grammar": bundle.grammar,
        "layouts": bundle.layouts,
        "coverage": bundle.coverage,
        "rhythm": bundle.rhythm,
        "qa": bundle.qa,
        "provenance": bundle.provenance,
        "recipe_provenance": bundle.recipe_provenance,
    }


def canvas_from_profile(canvas_dict: dict) -> CanvasSpec:
    """Build a CanvasSpec from a V2 profile's ``canvas`` section."""
    preset = canvas_dict.get("preset", "16:9")
    width = float(canvas_dict.get("width_pt", 959.976))
    height = float(canvas_dict.get("height_pt", 540))
    safe = canvas_dict.get("safe_margin", {})
    insets = SafeInsets(
        top=float(safe.get("top", 36)),
        right=float(safe.get("right", 48)),
        bottom=float(safe.get("bottom", 32)),
        left=float(safe.get("left", 48)),
    )
    return CanvasSpec(width_pt=width, height_pt=height, name=preset, safe=insets)
