"""Advanced visual effects for PowerPoint shapes via OOXML manipulation.

Provides 3D formatting, shadow effects (outer/inner/perspective), glow,
reflection, soft edges, and effect presets. All effects are applied by
manipulating the ``<a:effectLst>`` and ``<a:sp3d>`` elements inside a
shape's ``<a:spPr>`` (shape properties).

OOXML reference: ECMA-376 Part 4, Section 20.1.2 (DrawingML - Effects)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pptx_skill._io import is_presentation as _is_presentation
from pptx_skill._io import open_prs as _open_prs
from pptx_skill._io import save_prs as _save_prs_impl
from pptx_skill.constants import A_NS as _NS_A

__all__ = [
    "EffectInfo",
    "BevelPreset",
    "LightingPreset",
    "MaterialPreset",
    "EFFECT_PRESETS",
    "apply_3d_format",
    "apply_effect_preset",
    "apply_glow",
    "apply_inner_shadow",
    "apply_perspective_shadow",
    "apply_reflection",
    "apply_shadow",
    "apply_soft_edges",
    "list_effects",
    "remove_3d_format",
    "remove_effect",
    "remove_effects",
    "remove_glow",
    "remove_reflection",
    "remove_shadow",
    "remove_soft_edges",
    "pt_to_emu",
    "inch_to_emu",
    "cm_to_emu",
]

# ---------------------------------------------------------------------------
# Unit helpers
# ---------------------------------------------------------------------------

def pt_to_emu(pt: float) -> int:
    """Convert points to EMU (1 pt = 12700 EMU)."""
    return int(round(pt * 12700))


def inch_to_emu(inch: float) -> int:
    """Convert inches to EMU (1 in = 914400 EMU)."""
    return int(round(inch * 914400))


def cm_to_emu(cm: float) -> int:
    """Convert centimeters to EMU (1 cm = 360000 EMU)."""
    return int(round(cm * 360000))


# ---------------------------------------------------------------------------
# Preset enums / constants
# ---------------------------------------------------------------------------

class BevelPreset:
    """Named bevel dimension presets (width, height in EMU)."""
    NONE = (0, 0)
    SUBTLE = (pt_to_emu(4), pt_to_emu(2))
    LIGHT = (pt_to_emu(6), pt_to_emu(3))
    MEDIUM = (pt_to_emu(8), pt_to_emu(4))
    HARD = (pt_to_emu(12), pt_to_emu(6))
    ANGLED = (pt_to_emu(6), pt_to_emu(6))
    SOFT_ROUND = (pt_to_emu(8), pt_to_emu(5))
    CROSS = (pt_to_emu(6), pt_to_emu(3))
    RELAXED_INSET = (pt_to_emu(8), pt_to_emu(4))
    DIVOT = (pt_to_emu(10), pt_to_emu(5))
    RIBLET = (pt_to_emu(6), pt_to_emu(2))
    CONVEX = (pt_to_emu(8), pt_to_emu(4))
    COOL_SLANT = (pt_to_emu(6), pt_to_emu(3))
    ART_DECO = (pt_to_emu(10), pt_to_emu(5))
    SLOPE = (pt_to_emu(6), pt_to_emu(3))


class LightingPreset:
    """Lighting rig names for 3D shapes (ECMA-376 §20.1.10.4)."""
    LEGACY_FLAT = "legacyFlat"
    LEGACY_HARSH = "legacyHarsh"
    LEGACY_NORMAL = "legacyNormal"
    THREE_POINT = "threePt"
    BALANCED = "balanced"
    SOFT = "soft"
    HARSH = "harsh"
    FLOOD = "flood"
    CONTRASTING = "contrasting"
    MORNING = "morning"
    SUNRISE = "sunrise"
    SUNSET = "sunset"
    CHILLY = "chilly"
    FREEZING = "freezing"
    FLAT = "flat"
    TWO_POINT = "twoPt"
    GLOW = "glow"
    BRIGHT_ROOM = "brightRoom"


class MaterialPreset:
    """Material type names for 3D shapes (ECMA-376 §20.1.10.5)."""
    LEGACY_MATTE = "legacyMatte"
    LEGACY_PLASTIC = "legacyPlastic"
    LEGACY_METAL = "legacyMetal"
    LEGACY_WIREFRAME = "legacyWireframe"
    MATTE = "matte"
    PLASTIC = "plastic"
    METAL = "metal"
    FLAT = "flat"
    WIREFRAME = "wireframe"
    POWDER = "powder"
    TRANSLUCENT = "translucent"
    CLEAR = "clear"
    SOFTEDGE = "softedge"
    DK_EDGE = "dkEdge"


# ---------------------------------------------------------------------------
# Effect presets
# ---------------------------------------------------------------------------

EFFECT_PRESETS: dict[str, dict[str, Any]] = {
    "subtle_shadow": {
        "shadow": {"shadow_type": "outer", "blur_rad": pt_to_emu(4), "dist": pt_to_emu(2),
                    "direction": 2700000, "alpha": 40, "color": "000000"},
    },
    "hard_shadow": {
        "shadow": {"shadow_type": "outer", "blur_rad": pt_to_emu(1), "dist": pt_to_emu(3),
                    "direction": 2700000, "alpha": 70, "color": "000000"},
    },
    "perspective_shadow": {
        "shadow": {"shadow_type": "perspective", "blur_rad": pt_to_emu(6), "dist": pt_to_emu(3),
                    "direction": 5400000, "alpha": 50, "color": "000000",
                    "sx": -100000, "sy": -100000, "kx": -5400000, "ky": 0, "algn": "bl"},
    },
    "inner_glow": {
        "glow": {"radius": pt_to_emu(4), "alpha": 30, "color": "000000"},
    },
    "neon_glow": {
        "glow": {"radius": pt_to_emu(8), "alpha": 60, "color": "00BFFF"},
    },
    "red_glow": {
        "glow": {"radius": pt_to_emu(6), "alpha": 50, "color": "FF0000"},
    },
    "mirror_reflection": {
        "reflection": {"blur_rad": pt_to_emu(0.5), "dist": pt_to_emu(3),
                        "alpha": 50, "end_alpha": 0, "direction": 5400000,
                        "st_pos": 0, "end_pos": 50000},
    },
    "faded_reflection": {
        "reflection": {"blur_rad": pt_to_emu(1), "dist": pt_to_emu(2),
                        "alpha": 30, "end_alpha": 0, "direction": 5400000,
                        "st_pos": 0, "end_pos": 40000},
    },
    "soft_bevel": {
        "3d": {"bevel_top": BevelPreset.SUBTLE, "bevel_bottom": BevelPreset.SUBTLE,
               "depth": pt_to_emu(2), "lighting": LightingPreset.SOFT,
               "material": MaterialPreset.MATTE},
    },
    "hard_bevel": {
        "3d": {"bevel_top": BevelPreset.HARD, "bevel_bottom": BevelPreset.HARD,
               "depth": pt_to_emu(6), "lighting": LightingPreset.HARSH,
               "material": MaterialPreset.METAL},
    },
    "metallic": {
        "3d": {"bevel_top": BevelPreset.COOL_SLANT, "bevel_bottom": BevelPreset.COOL_SLANT,
               "depth": pt_to_emu(4), "contour": pt_to_emu(0.5), "contour_color": "808080",
               "lighting": LightingPreset.THREE_POINT, "material": MaterialPreset.METAL},
    },
    "glass": {
        "3d": {"bevel_top": BevelPreset.SOFT_ROUND, "bevel_bottom": BevelPreset.SOFT_ROUND,
               "depth": pt_to_emu(3), "lighting": LightingPreset.FREEZING,
               "material": MaterialPreset.CLEAR},
    },
    "soft_edges": {
        "soft_edges": {"radius": pt_to_emu(4)},
    },
    "subtle_soft_edges": {
        "soft_edges": {"radius": pt_to_emu(2)},
    },
    "emboss": {
        "3d": {"bevel_top": BevelPreset.RELAXED_INSET, "bevel_bottom": BevelPreset.RELAXED_INSET,
               "depth": pt_to_emu(1), "lighting": LightingPreset.LEGACY_FLAT,
               "material": MaterialPreset.LEGACY_MATTE},
    },
    "deboss": {
        "3d": {"bevel_top": BevelPreset.CROSS, "bevel_bottom": BevelPreset.CROSS,
               "depth": pt_to_emu(1), "lighting": LightingPreset.LEGACY_FLAT,
               "material": MaterialPreset.LEGACY_MATTE},
    },
}


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class EffectInfo:
    """Structured info about effects applied to a shape."""
    has_outer_shadow: bool = False
    has_inner_shadow: bool = False
    has_perspective_shadow: bool = False
    has_glow: bool = False
    has_reflection: bool = False
    has_soft_edges: bool = False
    has_3d: bool = False
    shadow_params: dict = field(default_factory=dict)
    glow_params: dict = field(default_factory=dict)
    reflection_params: dict = field(default_factory=dict)
    soft_edges_params: dict = field(default_factory=dict)
    three_d_params: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _save_prs(prs, path):
    _save_prs_impl(prs, path, backup=False)


def _find_shape(slide, shape_name: str):
    """Find a shape by name on a slide. Returns the shape or None."""
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None


def _get_or_create_effect_lst(sp_pr) -> Any:
    """Get or create ``<a:effectLst>`` under ``<a:spPr>``."""
    from lxml import etree
    effect_lst = sp_pr.find(f"{{{_NS_A}}}effectLst")
    if effect_lst is None:
        # Insert before <a:scene3d> or <a:sp3d> or at end
        insert_before = sp_pr.find(f"{{{_NS_A}}}scene3d")
        if insert_before is None:
            insert_before = sp_pr.find(f"{{{_NS_A}}}sp3d")
        effect_lst = etree.SubElement(sp_pr, f"{{{_NS_A}}}effectLst")
        if insert_before is not None:
            sp_pr.remove(effect_lst)
            sp_pr.insert(list(sp_pr).index(insert_before), effect_lst)
    return effect_lst


def _get_or_create_sp3d(sp_pr) -> Any:
    """Get or create ``<a:sp3d>`` under ``<a:spPr>``."""
    from lxml import etree
    sp3d = sp_pr.find(f"{{{_NS_A}}}sp3d")
    if sp3d is None:
        sp3d = etree.SubElement(sp_pr, f"{{{_NS_A}}}sp3d")
    return sp3d


def _make_color_elem(color: str, alpha: int | None = None) -> Any:
    """Create ``<a:srgbClr>`` element with optional alpha."""
    from lxml import etree
    clr = etree.SubElement(etree.Element("dummy"), f"{{{_NS_A}}}srgbClr")
    clr.set("val", color.upper())
    if alpha is not None:
        alpha_pct = max(0, min(100, alpha))
        alpha_elem = etree.SubElement(clr, f"{{{_NS_A}}}alpha")
        alpha_elem.set("val", str(int(alpha_pct * 1000)))
    return clr


def _add_color_to_parent(parent, tag: str, color: str, alpha: int | None = None):
    """Add a color element under *parent* with the given *tag*."""
    from lxml import etree
    # Remove existing element with same tag
    for existing in parent.findall(f"{{{_NS_A}}}{tag}"):
        parent.remove(existing)
    elem = etree.SubElement(parent, f"{{{_NS_A}}}{tag}")
    clr = _make_color_elem(color, alpha)
    # Copy color children into elem
    for child in clr:
        elem.append(child)
    elem.set("val", clr.get("val", ""))
    # Actually, for shadow/glow the color is a child element
    # Re-do: the color element is a direct child of the effect element
    for existing in parent.findall(f"{{{_NS_A}}}{tag}"):
        parent.remove(existing)
    elem = etree.SubElement(parent, f"{{{_NS_A}}}{tag}")
    clr_elem = etree.SubElement(elem, f"{{{_NS_A}}}srgbClr")
    clr_elem.set("val", color.upper())
    if alpha is not None:
        alpha_pct = max(0, min(100, alpha))
        alpha_child = etree.SubElement(clr_elem, f"{{{_NS_A}}}alpha")
        alpha_child.set("val", str(int(alpha_pct * 1000)))
    return elem


def _remove_effect_by_tag(effect_lst, tag: str) -> bool:
    """Remove an effect element by tag name from effectLst."""
    found = False
    for elem in effect_lst.findall(f"{{{_NS_A}}}{tag}"):
        effect_lst.remove(elem)
        found = True
    # Clean up empty effectLst
    if found and len(effect_lst) == 0:
        effect_lst.getparent().remove(effect_lst)
    return found


# ---------------------------------------------------------------------------
# Shadow effects
# ---------------------------------------------------------------------------

def apply_shadow(prs_or_path, slide_index: int, shape_name: str, *,
                 shadow_type: str = "outer",
                 blur_rad: int | None = None,
                 dist: int | None = None,
                 direction: int | None = None,
                 alpha: int | None = None,
                 color: str | None = None,
                 rotate_with_shape: bool | None = None) -> bool:
    """Apply a shadow effect to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    shadow_type : str
        "outer", "inner", or "perspective".
    blur_rad : int, optional
        Blur radius in EMU. Default 76200 (~6pt).
    dist : int, optional
        Shadow distance in EMU. Default 38100 (~3pt).
    direction : int, optional
        Direction in 60000ths of a degree. Default 5400000 (90° = below).
    alpha : int, optional
        Opacity 0-100. Default 40.
    color : str, optional
        Hex color. Default "000000".
    rotate_with_shape : bool, optional
        Whether shadow rotates with shape. Default True.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        effect_lst = _get_or_create_effect_lst(sp_pr)

        # Remove existing shadows of same type
        tag_map = {"outer": "outerShdw", "inner": "innerShdw", "perspective": "outerShdw"}
        tag = tag_map.get(shadow_type, "outerShdw")
        _remove_effect_by_tag(effect_lst, tag)

        # Defaults
        _blur = blur_rad if blur_rad is not None else 76200
        _dist = dist if dist is not None else 38100
        _dir = direction if direction is not None else 5400000
        _alpha = alpha if alpha is not None else 40
        _color = color if color is not None else "000000"
        _rot = rotate_with_shape if rotate_with_shape is not None else True

        if shadow_type == "inner":
            elem = etree.SubElement(effect_lst, f"{{{_NS_A}}}innerShdw")
        else:
            elem = etree.SubElement(effect_lst, f"{{{_NS_A}}}outerShdw")

        elem.set("blurRad", str(_blur))
        elem.set("dist", str(_dist))
        elem.set("dir", str(_dir))
        elem.set("algn", "bl" if shadow_type == "perspective" else "tl")
        elem.set("rotWithShape", "1" if _rot else "0")

        if shadow_type == "perspective":
            sx = -100000
            sy = -100000
            kx = -5400000
            ky = 0
            elem.set("sx", str(sx))
            elem.set("sy", str(sy))
            elem.set("kx", str(kx))
            elem.set("ky", str(ky))

        clr = etree.SubElement(elem, f"{{{_NS_A}}}srgbClr")
        clr.set("val", _color.upper())
        alpha_elem = etree.SubElement(clr, f"{{{_NS_A}}}alpha")
        alpha_elem.set("val", str(int(_alpha * 1000)))

        return True
    finally:
        _save_prs(prs, path)


def apply_inner_shadow(prs_or_path, slide_index: int, shape_name: str, *,
                       blur_rad: int | None = None,
                       dist: int | None = None,
                       direction: int | None = None,
                       alpha: int | None = None,
                       color: str | None = None) -> bool:
    """Apply an inner shadow effect. Convenience wrapper for ``apply_shadow``.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    return apply_shadow(prs_or_path, slide_index, shape_name,
                        shadow_type="inner", blur_rad=blur_rad, dist=dist,
                        direction=direction, alpha=alpha, color=color)


def apply_perspective_shadow(prs_or_path, slide_index: int, shape_name: str, *,
                             blur_rad: int | None = None,
                             dist: int | None = None,
                             direction: int | None = None,
                             alpha: int | None = None,
                             color: str | None = None,
                             sx: int | None = None,
                             sy: int | None = None,
                             kx: int | None = None,
                             ky: int | None = None,
                             algn: str | None = None) -> bool:
    """Apply a perspective shadow effect with full transform control.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        effect_lst = _get_or_create_effect_lst(sp_pr)
        _remove_effect_by_tag(effect_lst, "outerShdw")

        _blur = blur_rad if blur_rad is not None else 76200
        _dist = dist if dist is not None else 38100
        _dir = direction if direction is not None else 5400000
        _alpha = alpha if alpha is not None else 50
        _color = color if color is not None else "000000"
        _sx = sx if sx is not None else -100000
        _sy = sy if sy is not None else -100000
        _kx = kx if kx is not None else -5400000
        _ky = ky if ky is not None else 0
        _algn = algn if algn is not None else "bl"

        elem = etree.SubElement(effect_lst, f"{{{_NS_A}}}outerShdw")
        elem.set("blurRad", str(_blur))
        elem.set("dist", str(_dist))
        elem.set("dir", str(_dir))
        elem.set("algn", _algn)
        elem.set("rotWithShape", "0")
        elem.set("sx", str(_sx))
        elem.set("sy", str(_sy))
        elem.set("kx", str(_kx))
        elem.set("ky", str(_ky))

        clr = etree.SubElement(elem, f"{{{_NS_A}}}srgbClr")
        clr.set("val", _color.upper())
        alpha_elem = etree.SubElement(clr, f"{{{_NS_A}}}alpha")
        alpha_elem.set("val", str(int(_alpha * 1000)))

        return True
    finally:
        _save_prs(prs, path)


def remove_shadow(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove all shadow effects from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        effect_lst = sp_pr.find(f"{{{_NS_A}}}effectLst")
        if effect_lst is None:
            return False

        found = False
        for tag in ("outerShdw", "innerShdw"):
            if _remove_effect_by_tag(effect_lst, tag):
                found = True
        return found
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Glow effect
# ---------------------------------------------------------------------------

def apply_glow(prs_or_path, slide_index: int, shape_name: str, *,
               radius: int | None = None,
               color: str | None = None,
               alpha: int | None = None) -> bool:
    """Apply a glow effect to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    radius : int, optional
        Glow radius in EMU. Default 25400 (~2pt).
    color : str, optional
        Hex color. Default "000000".
    alpha : int, optional
        Opacity 0-100. Default 40.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        effect_lst = _get_or_create_effect_lst(sp_pr)
        _remove_effect_by_tag(effect_lst, "glow")

        _rad = radius if radius is not None else 25400
        _color = color if color is not None else "000000"
        _alpha = alpha if alpha is not None else 40

        glow = etree.SubElement(effect_lst, f"{{{_NS_A}}}glow")
        glow.set("rad", str(_rad))

        clr = etree.SubElement(glow, f"{{{_NS_A}}}srgbClr")
        clr.set("val", _color.upper())
        alpha_elem = etree.SubElement(clr, f"{{{_NS_A}}}alpha")
        alpha_elem.set("val", str(int(_alpha * 1000)))

        return True
    finally:
        _save_prs(prs, path)


def remove_glow(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove glow effect from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False
        effect_lst = sp_pr.find(f"{{{_NS_A}}}effectLst")
        if effect_lst is None:
            return False
        return _remove_effect_by_tag(effect_lst, "glow")
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Reflection effect
# ---------------------------------------------------------------------------

def apply_reflection(prs_or_path, slide_index: int, shape_name: str, *,
                     blur_rad: int | None = None,
                     dist: int | None = None,
                     alpha: int | None = None,
                     end_alpha: int | None = None,
                     direction: int | None = None,
                     sx: int | None = None,
                     sy: int | None = None,
                     kx: int | None = None,
                     ky: int | None = None,
                     algn: str | None = None,
                     st_pos: int | None = None,
                     end_pos: int | None = None) -> bool:
    """Apply a reflection effect to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    blur_rad : int, optional
        Blur radius in EMU. Default 6350.
    dist : int, optional
        Reflection offset distance in EMU. Default 38100.
    alpha : int, optional
        Start opacity 0-100. Default 50.
    end_alpha : int, optional
        End opacity 0-100. Default 0.
    direction : int, optional
        Direction in 60000ths degree. Default 5400000.
    st_pos : int, optional
        Start position 0-100000. Default 0.
    end_pos : int, optional
        End position 0-100000. Default 50000.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        effect_lst = _get_or_create_effect_lst(sp_pr)
        _remove_effect_by_tag(effect_lst, "reflection")

        _blur = blur_rad if blur_rad is not None else 6350
        _dist = dist if dist is not None else 38100
        _alpha = alpha if alpha is not None else 50
        _end_alpha = end_alpha if end_alpha is not None else 0
        _dir = direction if direction is not None else 5400000
        _sx = sx if sx is not None else 100000
        _sy = sy if sy is not None else 100000
        _kx = kx if kx is not None else 0
        _ky = ky if ky is not None else 0
        _algn = algn if algn is not None else "bl"
        _st_pos = st_pos if st_pos is not None else 0
        _end_pos = end_pos if end_pos is not None else 50000

        refl = etree.SubElement(effect_lst, f"{{{_NS_A}}}reflection")
        refl.set("blurRad", str(_blur))
        refl.set("dist", str(_dist))
        refl.set("dir", str(_dir))
        refl.set("algn", _algn)
        refl.set("stA", str(int(_alpha * 1000)))
        refl.set("endA", str(int(_end_alpha * 1000)))
        refl.set("stPos", str(_st_pos))
        refl.set("endPos", str(_end_pos))
        refl.set("sx", str(_sx))
        refl.set("sy", str(_sy))
        refl.set("kx", str(_kx))
        refl.set("ky", str(_ky))

        return True
    finally:
        _save_prs(prs, path)


def remove_reflection(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove reflection effect from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False
        effect_lst = sp_pr.find(f"{{{_NS_A}}}effectLst")
        if effect_lst is None:
            return False
        return _remove_effect_by_tag(effect_lst, "reflection")
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Soft edges effect
# ---------------------------------------------------------------------------

def apply_soft_edges(prs_or_path, slide_index: int, shape_name: str, *,
                     radius: int | None = None) -> bool:
    """Apply soft edges effect to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    radius : int, optional
        Soft edge radius in EMU. Default 50800 (~4pt).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        effect_lst = _get_or_create_effect_lst(sp_pr)
        _remove_effect_by_tag(effect_lst, "softEdge")

        _rad = radius if radius is not None else 50800

        soft = etree.SubElement(effect_lst, f"{{{_NS_A}}}softEdge")
        soft.set("rad", str(_rad))

        return True
    finally:
        _save_prs(prs, path)


def remove_soft_edges(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove soft edges effect from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False
        effect_lst = sp_pr.find(f"{{{_NS_A}}}effectLst")
        if effect_lst is None:
            return False
        return _remove_effect_by_tag(effect_lst, "softEdge")
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# 3D format
# ---------------------------------------------------------------------------

def apply_3d_format(prs_or_path, slide_index: int, shape_name: str, *,
                    bevel_top: tuple | str | None = None,
                    bevel_bottom: tuple | str | None = None,
                    depth: int | None = None,
                    contour: int | None = None,
                    contour_color: str | None = None,
                    lighting: str | None = None,
                    material: str | None = None) -> bool:
    """Apply 3D formatting to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    bevel_top : tuple or str, optional
        Top bevel as (width_emu, height_emu) tuple or BevelPreset constant name.
    bevel_bottom : tuple or str, optional
        Bottom bevel as (width_emu, height_emu) tuple or BevelPreset constant name.
    depth : int, optional
        Extrusion depth in EMU. Default 0.
    contour : int, optional
        Contour width in EMU. Default 0.
    contour_color : str, optional
        Contour color hex. Default "000000".
    lighting : str, optional
        Lighting preset name. Default "threePt".
    material : str, optional
        Material preset name. Default "matte".
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False

        # Remove existing sp3d
        existing = sp_pr.find(f"{{{_NS_A}}}sp3d")
        if existing is not None:
            sp_pr.remove(existing)

        sp3d = _get_or_create_sp3d(sp_pr)

        # Resolve bevel presets
        def _resolve_bevel(val):
            if val is None:
                return None
            if isinstance(val, str):
                preset = getattr(BevelPreset, val.upper(), None)
                if preset is not None:
                    return preset
                return None
            if isinstance(val, (tuple, list)) and len(val) == 2:
                return tuple(val)
            return None

        top_bevel = _resolve_bevel(bevel_top)
        bottom_bevel = _resolve_bevel(bevel_bottom)

        if top_bevel is not None:
            # Remove existing bevelT
            for existing in sp3d.findall(f"{{{_NS_A}}}bevelT"):
                sp3d.remove(existing)
            bt = etree.SubElement(sp3d, f"{{{_NS_A}}}bevelT")
            bt.set("w", str(top_bevel[0]))
            bt.set("h", str(top_bevel[1]))

        if bottom_bevel is not None:
            for existing in sp3d.findall(f"{{{_NS_A}}}bevelB"):
                sp3d.remove(existing)
            bb = etree.SubElement(sp3d, f"{{{_NS_A}}}bevelB")
            bb.set("w", str(bottom_bevel[0]))
            bb.set("h", str(bottom_bevel[1]))

        _depth = depth if depth is not None else 0
        if _depth > 0:
            sp3d.set("depth", str(_depth))

        _contour = contour if contour is not None else 0
        if _contour > 0:
            sp3d.set("contourW", str(_contour))
            _cc = contour_color if contour_color is not None else "000000"
            # Remove existing contour color
            for existing in sp3d.findall(f"{{{_NS_A}}}contourClr"):
                sp3d.remove(existing)
            cc = etree.SubElement(sp3d, f"{{{_NS_A}}}contourClr")
            srgb = etree.SubElement(cc, f"{{{_NS_A}}}srgbClr")
            srgb.set("val", _cc.upper())

        _lighting = lighting if lighting is not None else LightingPreset.THREE_POINT
        sp3d.set("lit", _lighting)

        _material = material if material is not None else MaterialPreset.MATTE
        sp3d.set("prstMat", _material)

        return True
    finally:
        _save_prs(prs, path)


def remove_3d_format(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove 3D formatting from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False
        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return False
        sp3d = sp_pr.find(f"{{{_NS_A}}}sp3d")
        if sp3d is None:
            return False
        sp_pr.remove(sp3d)
        return True
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Generic remove
# ---------------------------------------------------------------------------

def remove_effect(prs_or_path, slide_index: int, shape_name: str,
                  effect_type: str) -> bool:
    """Remove a specific effect type from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    effect_type : str
        One of "shadow", "inner_shadow", "outer_shadow", "glow",
        "reflection", "soft_edges", "3d".
    """
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    if effect_type in ("shadow", "outer_shadow", "inner_shadow"):
        return remove_shadow(prs_or_path, slide_index, shape_name)
    if effect_type == "glow":
        return remove_glow(prs_or_path, slide_index, shape_name)
    if effect_type == "reflection":
        return remove_reflection(prs_or_path, slide_index, shape_name)
    if effect_type == "soft_edges":
        return remove_soft_edges(prs_or_path, slide_index, shape_name)
    if effect_type == "3d":
        return remove_3d_format(prs_or_path, slide_index, shape_name)
    return False


def remove_effects(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove all effects from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    results = []
    results.append(remove_shadow(prs_or_path, slide_index, shape_name))
    results.append(remove_glow(prs_or_path, slide_index, shape_name))
    results.append(remove_reflection(prs_or_path, slide_index, shape_name))
    results.append(remove_soft_edges(prs_or_path, slide_index, shape_name))
    results.append(remove_3d_format(prs_or_path, slide_index, shape_name))
    return any(results)


# ---------------------------------------------------------------------------
# Effect presets
# ---------------------------------------------------------------------------

def apply_effect_preset(prs_or_path, slide_index: int, shape_name: str,
                        preset_name: str) -> bool:
    """Apply a named effect preset to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).

    Available presets: subtle_shadow, hard_shadow, perspective_shadow,
    inner_glow, neon_glow, red_glow, mirror_reflection, faded_reflection,
    soft_bevel, hard_bevel, metallic, glass, soft_edges, subtle_soft_edges,
    emboss, deboss.
    """
    preset = EFFECT_PRESETS.get(preset_name)
    if preset is None:
        raise ValueError(f"Unknown effect preset: {preset_name!r}. "
                         f"Available: {', '.join(sorted(EFFECT_PRESETS))}")

    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")

    # Apply each effect in the preset
    success = True

    if "shadow" in preset:
        s = preset["shadow"]
        if not apply_shadow(prs_or_path, slide_index, shape_name, **s):
            success = False

    if "glow" in preset:
        g = preset["glow"]
        if not apply_glow(prs_or_path, slide_index, shape_name, **g):
            success = False

    if "reflection" in preset:
        r = preset["reflection"]
        if not apply_reflection(prs_or_path, slide_index, shape_name, **r):
            success = False

    if "soft_edges" in preset:
        se = preset["soft_edges"]
        if not apply_soft_edges(prs_or_path, slide_index, shape_name, **se):
            success = False

    if "3d" in preset:
        t = preset["3d"]
        if not apply_3d_format(prs_or_path, slide_index, shape_name, **t):
            success = False

    return success


# ---------------------------------------------------------------------------
# List effects
# ---------------------------------------------------------------------------

def list_effects(prs_or_path, slide_index: int, shape_name: str) -> EffectInfo:
    """List all effects applied to a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    info = EffectInfo()

    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range "
                         f"(1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return info

        sp_pr = shape._element.find(f"{{{_NS_A}}}spPr")
        if sp_pr is None:
            return info

        # Check effectLst
        effect_lst = sp_pr.find(f"{{{_NS_A}}}effectLst")
        if effect_lst is not None:
            for outer in effect_lst.findall(f"{{{_NS_A}}}outerShdw"):
                info.has_outer_shadow = True
                params = dict(outer.attrib)
                # Check if perspective (has sx/sy/kx/ky)
                if "sx" in params or "kx" in params:
                    info.has_perspective_shadow = True
                info.shadow_params = params

            for inner in effect_lst.findall(f"{{{_NS_A}}}innerShdw"):
                info.has_inner_shadow = True
                info.shadow_params.update(inner.attrib)

            for glow in effect_lst.findall(f"{{{_NS_A}}}glow"):
                info.has_glow = True
                info.glow_params = dict(glow.attrib)

            for refl in effect_lst.findall(f"{{{_NS_A}}}reflection"):
                info.has_reflection = True
                info.reflection_params = dict(refl.attrib)

            for soft in effect_lst.findall(f"{{{_NS_A}}}softEdge"):
                info.has_soft_edges = True
                info.soft_edges_params = dict(soft.attrib)

        # Check sp3d
        sp3d = sp_pr.find(f"{{{_NS_A}}}sp3d")
        if sp3d is not None:
            info.has_3d = True
            params = dict(sp3d.attrib)
            for bt in sp3d.findall(f"{{{_NS_A}}}bevelT"):
                params["bevelT"] = dict(bt.attrib)
            for bb in sp3d.findall(f"{{{_NS_A}}}bevelB"):
                params["bevelB"] = dict(bb.attrib)
            info.three_d_params = params

        return info
    finally:
        # Don't save for read-only list operation
        pass
