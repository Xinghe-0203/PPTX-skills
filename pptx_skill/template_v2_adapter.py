"""V1 -> V2 migration and V2 -> legacy theme adapter (PR7).

The 12 existing catalog profiles use a flat ``theme`` + ``fonts`` shape that
the legacy renderer consumes directly. Migrating to V2 promotes those fields
into the three-layer token model, fills the 14-role layout map and preserves
unknown fields under ``extensions.legacy_v1``. The reverse adapter converts a
compiled V2 bundle back into the legacy ``theme``/``fonts`` shape so existing
``choose_theme``/``draw_*`` functions keep working unchanged.
"""
from __future__ import annotations

import copy
from typing import Any

from pptx_skill.design_schema import (
    TEMPLATE_SCHEMA_VERSION,
    StyleIntent,
    canvas_from_profile,
    resolve_all_tokens,
    validate_source_profile,
)
from pptx_skill.template_compiler import compile_source_profile


LEGACY_THEME_FIELDS = (
    "bg", "bg_alt", "primary", "secondary", "accent",
    "text", "text_muted", "white", "dark", "on_dark",
)


def _palette_from_legacy_theme(theme: dict[str, Any]) -> dict[str, str]:
    """Pick a deterministic palette name from the legacy theme."""
    return {
        "ink_950": theme.get("text", "#1A1A1A"),
        "paper_050": theme.get("bg", "#FDFCF8"),
        "accent": theme.get("accent", "#C96845"),
        "muted": theme.get("text_muted", "#6E7C87"),
    }


def _family_from_layout_family(layout_family: str) -> str:
    if layout_family in {"technical_axis", "poster_column", "editorial_grid"}:
        return layout_family
    return "standard"


def migrate_profile_v1_to_v2(v1: dict[str, Any], key: str | None = None) -> dict[str, Any]:
    """Convert a V1 profile dict into a partial V2 source profile.

    The 14-role layout map is filled by the compiler in a second pass; this
    function only emits the partial structure with migrated tokens, grammar
    and rhythm, leaving unknown fields under ``extensions.legacy_v1``.

    ``key`` is the catalog key the profile lives under, used as the fallback id
    when the V1 payload does not carry one (the common case for the 12 catalog
    profiles, which are keyed by name rather than carrying an explicit ``id``).
    """
    if not isinstance(v1, dict):
        raise TypeError("V1 profile must be a JSON object")
    theme = v1.get("theme", {})
    fonts = v1.get("fonts", {})
    family = _family_from_layout_family(v1.get("layout_family", "standard"))
    layout_opts = v1.get("layout_opts", {})

    profile_id = v1.get("id", key or "v1-migrated")
    profile: dict[str, Any] = {
        "schema_version": TEMPLATE_SCHEMA_VERSION,
        "kind": "template-source-profile",
        "id": profile_id,
        "name": v1.get("name", profile_id),
        "description": v1.get("description", ""),
        "canvas": {
            "preset": "16:9",
            "width_pt": 959.976,
            "height_pt": 540.0,
            "safe_margin": {"top": 36, "right": 48, "bottom": 32, "left": 48},
        },
        "tokens": {
            "primitive": {
                "palette": _palette_from_legacy_theme(theme),
                "font": {
                    "family": {
                        "display": fonts.get("en", "Aptos Display"),
                        "body": fonts.get("cn", "Microsoft YaHei"),
                    },
                    "size": {
                        "hero": float(layout_opts.get("cover", {}).get("title_size", 56)),
                        "title": float(layout_opts.get("cover", {}).get("title_size", 36)),
                        "body": float(layout_opts.get("bullets", {}).get("body_size", 16)),
                        "label": 10,
                    },
                    "min": {"title": 26, "body": 13, "label": 9},
                    "line_height": {"title": 1.08, "body": 1.35},
                },
                "space": {"4": 16, "6": 24, "8": 32, "12": 48},
                "stroke": {"hairline": 0.75, "regular": 1.25},
                "radius": {"none": 0, "small": 3},
            },
            "semantic": {
                "color": {
                    "canvas": {"$ref": "primitive.palette.paper_050"},
                    "text_primary": {"$ref": "primitive.palette.ink_950"},
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
                "title": {"color": {"$ref": "semantic.color.text_primary"}, "family": {"$ref": "semantic.type.display_family"}, "size": {"$ref": "semantic.type.title_size"}},
                "body": {"color": {"$ref": "semantic.color.text_primary"}, "family": {"$ref": "semantic.type.body_family"}, "size": {"$ref": "semantic.type.body_size"}},
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
            "rules": [],
        },
        "qa": {
            "contrast": {"body_min": 4.5, "large_min": 3.0},
            "max_colors_per_slide": 7,
            "min_body_font_size": max(13, int(layout_opts.get("bullets", {}).get("body_size", 14))),
        },
        "provenance": {
            "generator": "migrate_profile_v1_to_v2",
            "migrated_from": 1,
            "manual_review_required": True,
        },
        "extensions": {"legacy_v1": copy.deepcopy(v1)},
    }
    # Some catalog profiles don't carry an explicit `id`; ensure the round-tripped
    # legacy payload exposes one so callers can recover it under extensions.
    profile["extensions"]["legacy_v1"].setdefault("id", profile_id)

    # Promote preferred_sequence into rhythm (referenced by deck planner).
    preferred = v1.get("preferred_sequence", [])
    if preferred:
        profile["provenance"]["legacy_preferred_sequence"] = list(preferred)

    validate_source_profile(profile)
    return profile


def migrate_and_compile_v1(v1: dict[str, Any]) -> "CompiledTemplateBundleV2":
    """One-shot V1 -> partial V2 source -> compiled bundle."""
    source = migrate_profile_v1_to_v2(v1)
    return compile_source_profile(source)


# ---------------------------------------------------------------------------
# V2 -> legacy theme adapter (backwards compat with draw_*)
# ---------------------------------------------------------------------------


def _lookup(path: str, tokens: dict[str, Any]) -> Any:
    parts = path.split(".")
    value: Any = tokens
    for part in parts:
        if not isinstance(value, dict) or part not in value:
            return None
        value = value[part]
    return value


def _relative_luminance(hex_color: str) -> float:
    """WCAG relative luminance for an ``#RRGGBB`` color (0=dark, 1=light)."""
    color = hex_color.lstrip("#")
    if len(color) != 6:
        return 1.0
    rgb = [int(color[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]

    def channel(c: float) -> float:
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return 0.2126 * channel(rgb[0]) + 0.7152 * channel(rgb[1]) + 0.0722 * channel(rgb[2])


def v2_bundle_to_legacy_theme(bundle) -> dict[str, Any]:
    """Convert a compiled V2 bundle back into the legacy theme/fonts shape.

    The result is consumable by ``choose_theme``/``draw_*`` even though it
    does not preserve all V2 grammar semantics (component tokens are dropped).
    """
    resolved = bundle.tokens if bundle.tokens else resolve_all_tokens(getattr(bundle, "tokens", {}))
    ink = _lookup("primitive.palette.ink_950", resolved) or "#1A1A1A"
    paper = _lookup("primitive.palette.paper_050", resolved) or "#FDFCF8"
    accent = _lookup("primitive.palette.accent", resolved) or "#C96845"
    muted = _lookup("primitive.palette.muted", resolved) or "#6E7C87"
    secondary = muted
    dark_color = ink
    text_color = ink
    # A dark canvas is one whose background is darker than its primary text.
    on_dark = _relative_luminance(paper) < _relative_luminance(ink)
    return {
        "bg": paper,
        "bg_alt": paper,
        "primary": ink,
        "secondary": secondary,
        "accent": accent,
        "text": text_color,
        "text_muted": muted,
        "white": "#FFFFFF",
        "dark": dark_color,
        "on_dark": on_dark,
    }


def v2_bundle_to_legacy_fonts(bundle) -> dict[str, str]:
    resolved = bundle.tokens
    display = _lookup("primitive.font.family.display", resolved) or "Aptos Display"
    body = _lookup("primitive.font.family.body", resolved) or "Microsoft YaHei"
    return {"cn": body, "en": display}


def v2_bundle_to_legacy_profile(bundle) -> dict[str, Any]:
    """Build the legacy ``profile`` dict (``theme`` + ``fonts`` + ``layout_opts``)."""
    family = (bundle.grammar or {}).get("family", "standard")
    layout_family = family if family in {"standard", "editorial_grid", "technical_axis", "poster_column"} else "editorial_grid"
    return {
        "id": bundle.id,
        "name": bundle.name,
        "description": bundle.description,
        "layout_family": layout_family,
        "theme": v2_bundle_to_legacy_theme(bundle),
        "fonts": v2_bundle_to_legacy_fonts(bundle),
        "layout_opts": {},
        "preferred_sequence": [role for role in REQUIRED_ROLES_FROM_BUNDLE if role in bundle.layouts],
    }


# Avoid hard-coupling to the design_schema constant name; redefine here.
REQUIRED_ROLES_FROM_BUNDLE = (
    "cover", "toc", "section", "bullets", "text_image", "full_image",
    "image_grid", "dashboard", "timeline", "comparison", "quote",
    "process", "table", "end",
)
