"""TemplateProfileV2 compiler (PR7).

Compiles a :class:`StyleIntent` into a source ``TemplateProfileV2`` and then
into a complete :class:`CompiledTemplateBundleV2` covering all 14 roles. The
compiler is deterministic for a given (normalized StyleIntent, compiler
version, catalog revision, seed) tuple. It only emits a constrained design
system; it never executes generated code.
"""
from __future__ import annotations

import copy
import hashlib
import re
from typing import Any

from pptx_skill.content_model import CanvasSpec
from pptx_skill.design_schema import (
    CompiledTemplateBundleV2,
    REQUIRED_ROLES,
    TEMPLATE_SCHEMA_VERSION,
    StyleIntent,
    canvas_from_profile,
    resolve_all_tokens,
    validate_source_profile,
)

COMPILER_VERSION = "template-compiler-v2"


# ---------------------------------------------------------------------------
# Deterministic helpers
# ---------------------------------------------------------------------------


_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _seed_from_intent(intent: StyleIntent) -> int:
    payload = repr(intent.normalize())
    digest = hashlib.sha256(payload.encode("utf-8")).digest()
    return int.from_bytes(digest[:4], "big")


def _mix(a: str, b: str, weight_b: float) -> str:
    def hx(c: str) -> tuple[int, int, int]:
        c = c.lstrip("#")
        return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)

    ra, ga, ba = hx(a)
    rb, gb, bb = hx(b)
    out = (
        round(ra * (1 - weight_b) + rb * weight_b),
        round(ga * (1 - weight_b) + gb * weight_b),
        round(ba * (1 - weight_b) + bb * weight_b),
    )
    return "#" + "".join(f"{v:02X}" for v in out)


def _is_dark_intent(intent: StyleIntent) -> bool:
    dark_tokens = ("dark", "深色", "night", "black")
    light_tokens = ("light", "浅色", "white", "bright")
    blob = " ".join([intent.industry, intent.audience, intent.image_style, *intent.tone]).lower()
    if any(t in blob for t in light_tokens):
        return False
    return any(t in blob for t in dark_tokens)


def _select_grammar_family(intent: StyleIntent) -> str:
    geo = {g.lower() for g in intent.geometry}
    if intent.geometry:
        if any(g in geo for g in ("axes", "hairlines", "axis", "technical")):
            return "technical_axis"
        if any(g in geo for g in ("poster", "columns", "bold")):
            return "poster_column"
        if any(g in geo for g in ("grid", "whitespace", "editorial")):
            return "editorial_grid"
    # Default by industry keyword.
    return "editorial_grid"


# ---------------------------------------------------------------------------
# StyleIntent -> source profile
# ---------------------------------------------------------------------------


def _palette_from_intent(intent: StyleIntent, seed: int) -> dict[str, str]:
    """Produce a primitive palette from brand colors or the deterministic seed."""
    if intent.brand_colors:
        colors = [c.upper() for c in intent.brand_colors]
        for c in colors:
            if not _HEX_RE.match(c):
                raise ValueError(f"Invalid brand color: {c}")
        primary = colors[0]
        accent = colors[1] if len(colors) > 1 else _mix(primary, "#C96845", 0.55)
    else:
        hue = (seed % 360) / 360.0
        import colorsys

        r, g, b = colorsys.hls_to_rgb(hue, 0.32, 0.55)
        primary = "#" + "".join(f"{round(v * 255):02X}" for v in (r, g, b))
        r, g, b = colorsys.hls_to_rgb((hue + 0.11) % 1.0, 0.52, 0.62)
        accent = "#" + "".join(f"{round(v * 255):02X}" for v in (r, g, b))
    dark = _is_dark_intent(intent)
    if dark:
        ink = _mix(primary, "#080A09", 0.88)
        paper = _mix(primary, "#111D1B", 0.74)
    else:
        ink = _mix(primary, "#1A1A1A", 0.78)
        paper = _mix(primary, "#FDFCF8", 0.92)
    muted = _mix(primary, paper, 0.55)
    return {
        "ink_950": ink,
        "paper_050": paper,
        "accent": accent,
        "muted": muted,
    }


def source_profile_from_intent(
    intent: StyleIntent,
    seed: int | None = None,
    compiler_version: str = COMPILER_VERSION,
    catalog_revision: int = 1,
) -> dict:
    """Build a partial V2 source profile from a StyleIntent.

    The source profile covers a small set of roles; the compiler fills the
    rest. Output is deterministic for a given normalized intent + seed.
    """
    seed_value = seed if seed is not None else _seed_from_intent(intent)
    dark = _is_dark_intent(intent)
    palette = _palette_from_intent(intent, seed_value)
    family = _select_grammar_family(intent)

    profile: dict[str, Any] = {
        "schema_version": TEMPLATE_SCHEMA_VERSION,
        "kind": "template-source-profile",
        "id": f"profile-{seed_value:08x}",
        "name": intent.industry or f"Generated {family}",
        "description": intent.audience or "Generated TemplateProfileV2",
        "canvas": {
            "preset": intent.canvas or "16:9",
            "width_pt": 959.976 if (intent.canvas or "16:9") == "16:9" else 720.0,
            "height_pt": 540.0,
            "safe_margin": {"top": 38, "right": 52, "bottom": 34, "left": 52},
        },
        "tokens": {
            "primitive": {
                "palette": palette,
                "font": {
                    "family": {"display": "Aptos Display", "body": "Microsoft YaHei"},
                    "size": {"hero": 58, "title": 32, "body": 17, "label": 10},
                    "min": {"title": 26, "body": 15, "label": 9},
                    "line_height": {"title": 1.08, "body": 1.35},
                },
                "space": {"4": 16, "6": 24, "8": 32, "12": 48},
                "stroke": {"hairline": 0.75, "regular": 1.25},
                "radius": {"none": 0, "small": 3},
            },
            "semantic": {
                "color": {
                    "canvas": {"$ref": "primitive.palette.ink_950" if dark else "primitive.palette.paper_050"},
                    "surface_raised": {"$ref": "primitive.palette.paper_050" if dark else "primitive.palette.muted"},
                    "text_primary": {"$ref": "primitive.palette.paper_050" if dark else "primitive.palette.ink_950"},
                    "text_muted": {"$ref": "primitive.palette.muted"},
                    "accent": {"$ref": "primitive.palette.accent"},
                },
                "type": {
                    "display_family": {"$ref": "primitive.font.family.display"},
                    "body_family": {"$ref": "primitive.font.family.body"},
                    "title_size": {"$ref": "primitive.font.size.title"},
                    "body_size": {"$ref": "primitive.font.size.body"},
                },
                "space": {
                    "section": {"$ref": "primitive.space.12"},
                    "component": {"$ref": "primitive.space.4"},
                },
            },
            "component": {
                "title": {
                    "color": {"$ref": "semantic.color.text_primary"},
                    "family": {"$ref": "semantic.type.display_family"},
                    "size": {"$ref": "semantic.type.title_size"},
                },
                "body": {
                    "color": {"$ref": "semantic.color.text_primary"},
                    "family": {"$ref": "semantic.type.body_family"},
                    "size": {"$ref": "semantic.type.body_size"},
                },
                "metric_value": {
                    "color": {"$ref": "semantic.color.accent"},
                    "size": {"$ref": "primitive.font.size.hero"},
                },
            },
        },
        "grammar": {
            "family": family,
            "grid": {"columns": 12, "gutter": 16},
            "primary_axes": ["left", "baseline"],
            "reading_paths": ["f", "column"],
            "whitespace_ratio": {"min": 0.20, "target": 0.28},
        },
        "layouts": {
            "cover": [f"{family}.cover.axis"],
            "bullets": [f"{family}.bullets.rail", f"{family}.bullets.wide"],
            "text_image": [f"{family}.text_image.asymmetric_left"],
            "dashboard": [f"{family}.dashboard.lead_metric"],
        },
        "coverage": {
            "status": "partial",
            "roles": ["cover", "bullets", "text_image", "dashboard"],
        },
        "rhythm": {
            "avoid_adjacent_same_signature": True,
            "emphasis_interval": [3, 5],
            "rules": [
                {"after": "dashboard", "prefer": ["section", "text_image"], "avoid": ["dashboard"]},
            ],
        },
        "qa": {
            "contrast": {"body_min": 4.5, "large_min": 3.0},
            "max_colors_per_slide": 7,
            "min_body_font_size": 15,
            "max_crop_subject_loss": 0.12,
        },
        "provenance": {
            "generator": compiler_version,
            "seed": seed_value,
            "catalog_revision": catalog_revision,
            "source_profiles": [],
            "style_intent": intent.normalize(),
        },
    }
    validate_source_profile(profile)
    return profile


# ---------------------------------------------------------------------------
# Source profile -> compiled bundle
# ---------------------------------------------------------------------------


def _recipe_id_for(family: str, role: str, variant: str) -> str:
    return f"{family}.{role}.{variant}"


def _default_recipe_variants(role: str) -> list[str]:
    """Return the default variant set for a role (>=1 per role)."""
    defaults = {
        "cover": ["axis", "poster"],
        "toc": ["list", "rail"],
        "section": ["full", "split"],
        "bullets": ["rail", "wide"],
        "text_image": ["asymmetric_left", "stacked"],
        "full_image": ["full", "caption"],
        "image_grid": ["grid2", "grid3"],
        "dashboard": ["lead_metric", "split"],
        "timeline": ["horizontal", "vertical"],
        "comparison": ["two_column", "stacked"],
        "quote": ["centered", "editorial"],
        "process": ["horizontal", "vertical"],
        "table": ["standard", "striped"],
        "end": ["centered", "split"],
    }
    return defaults.get(role, ["default"])


def _complete_layouts(source: dict) -> tuple[dict[str, list[str]], dict[str, str]]:
    """Fill missing roles with generated recipes, recording provenance."""
    family = source.get("grammar", {}).get("family", "standard")
    layouts = copy.deepcopy(source.get("layouts", {}))
    provenance: dict[str, str] = {}
    covered = set(source.get("coverage", {}).get("roles", []))
    for role in REQUIRED_ROLES:
        if role in layouts and layouts[role]:
            for rid in layouts[role]:
                provenance[rid] = "generated"
            continue
        variants = _default_recipe_variants(role)
        layouts[role] = [_recipe_id_for(family, role, v) for v in variants]
        for rid in layouts[role]:
            provenance[rid] = "generated" if role in covered else "inherited"
    return layouts, provenance


def compile_source_profile(source: dict) -> CompiledTemplateBundleV2:
    """Compile a partial source profile into a complete bundle.

    The compiler resolves all ``$ref`` tokens, fills the 14 roles, records
    recipe provenance, and marks ``coverage.status='complete'``. It never
    executes generated code; recipes are identifiers resolved by the layout
    engine's recipe registry.
    """
    warnings = validate_source_profile(source)
    resolved_tokens = resolve_all_tokens(source["tokens"])
    layouts, recipe_provenance = _complete_layouts(source)
    canvas = canvas_from_profile(source["canvas"])

    coverage = {
        "status": "complete",
        "roles": list(REQUIRED_ROLES),
    }
    provenance = dict(source.get("provenance", {}))
    provenance.setdefault("generator", COMPILER_VERSION)
    provenance["compiled_by"] = COMPILER_VERSION

    bundle = CompiledTemplateBundleV2(
        schema_version=TEMPLATE_SCHEMA_VERSION,
        kind="compiled-template-bundle",
        id=source["id"],
        name=source["name"],
        description=source["description"],
        canvas=canvas,
        tokens=resolved_tokens,
        grammar=source.get("grammar", {}),
        layouts=layouts,
        coverage=coverage,
        rhythm=source.get("rhythm", {}),
        qa=source.get("qa", {}),
        provenance=provenance,
        recipe_provenance=recipe_provenance,
    )
    return bundle


def compile_from_intent(
    intent: StyleIntent,
    seed: int | None = None,
    catalog_revision: int = 1,
) -> CompiledTemplateBundleV2:
    """One-shot: StyleIntent -> source profile -> compiled bundle."""
    source = source_profile_from_intent(intent, seed=seed, catalog_revision=catalog_revision)
    return compile_source_profile(source)


# ---------------------------------------------------------------------------
# Geometry signature / diversity gate
# ---------------------------------------------------------------------------


def recipe_geometry_signature(recipe_id: str, family: str) -> tuple:
    """A normalized signature for a recipe used by the diversity gate.

    The signature combines role, variant and family so two bundles that only
    differ in palette but share the same recipe set compare as identical.
    """
    parts = recipe_id.split(".")
    if len(parts) >= 3:
        return (parts[0], parts[1], parts[2])  # family, role, variant
    return (family, recipe_id, "default")


def bundle_diversity_signature(bundle: CompiledTemplateBundleV2) -> frozenset:
    """Return the set of (family, role, variant) triples covered by a bundle.

    ``family`` is part of the signature so two bundles that share recipe variants
    but use a different grammar family (e.g. editorial_grid vs technical_axis)
    are not collapsed to identical by the diversity gate.
    """
    sigs: set[tuple] = set()
    family = bundle.grammar.get("family", "standard")
    for role, recipe_ids in bundle.layouts.items():
        for rid in recipe_ids:
            _, _, variant = recipe_geometry_signature(rid, family)
            sigs.add((family, role, variant))
    return frozenset(sigs)


def geometry_similarity(bundle_a: CompiledTemplateBundleV2, bundle_b: CompiledTemplateBundleV2) -> float:
    """Return 0..1 Jaccard similarity of two bundles' recipe variant sets."""
    a = bundle_diversity_signature(bundle_a)
    b = bundle_diversity_signature(bundle_b)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def passes_diversity_gate(
    new_bundle: CompiledTemplateBundleV2,
    existing_bundles: list[CompiledTemplateBundleV2],
    max_similarity: float = 0.85,
) -> tuple[bool, float]:
    """Return (passes, max_similarity_to_existing).

    Rejects registration when a new bundle is too similar to an existing one.
    """
    if not existing_bundles:
        return True, 0.0
    sims = [geometry_similarity(new_bundle, other) for other in existing_bundles]
    peak = max(sims)
    return peak <= max_similarity, peak
