"""Slide animation support via OOXML manipulation.

python-pptx has no native API for slide animations, so this module
constructs the appropriate ``<p:timing>`` / ``<p:anim>`` /
``<p:animEffect>`` / ``<p:animMotion>`` XML elements directly using
lxml and inserts them into each slide's ``<p:sld>`` element.

Usage
-----
>>> from pptx_skill.animations import apply_animation, FADE_IN
>>> apply_animation(slide, shape, FADE_IN, duration_ms=500, trigger="onClick")

Convenience wrappers for each animation category:

>>> from pptx_skill.animations import apply_entrance_animation, FLOAT_UP
>>> apply_entrance_animation(slide, shape, FLOAT_UP, duration_ms=600)

>>> from pptx_skill.animations import apply_exit_animation, FADE_OUT
>>> apply_exit_animation(slide, shape, FADE_OUT, delay_ms=200)

>>> from pptx_skill.animations import apply_emphasis_animation, SPIN
>>> apply_emphasis_animation(slide, shape, SPIN, repeat=3)

>>> from pptx_skill.animations import apply_motion_path
>>> apply_motion_path(slide, shape, [(0, 0), (100, 50), (200, 0)])

Remove or inspect animations on a slide:

>>> from pptx_skill.animations import remove_animations, list_animations
>>> remove_animations(slide)
>>> anims = list_animations(slide)
"""
from __future__ import annotations

import logging
import warnings
from typing import Any

from lxml import etree

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 1. Animation type constants
# ---------------------------------------------------------------------------

# -- Entrance animations --
APPEAR = "appear"
FLY_IN = "fly_in"
FLOAT_UP = "float_up"
ZOOM = "zoom"
GROW_TURN = "grow_turn"
SWIVEL = "swivel"
BOUNCE = "bounce"
FADE_IN = "fade_in"
WIPE_IN = "wipe_in"
BLINDS_IN = "blinds_in"
BOX_IN = "box_in"
CHECKERBOARD_IN = "checkerboard_in"
SPLIT_IN = "split_in"
DIAGONAL_IN = "diagonal_in"
RANDOM_BARS_IN = "random_bars_in"
ASCEND = "ascend"
DESCEND = "descend"
SPIN_IN = "spin_in"
STRETCH_IN = "stretch_in"
WHEEL_IN = "wheel_in"

# -- Exit animations --
DISAPPEAR = "disappear"
FLY_OUT = "fly_out"
FLOAT_DOWN = "float_down"
ZOOM_OUT = "zoom_out"
SHRINK_TURN = "shrink_turn"
SWIVEL_OUT = "swivel_out"
BOUNCE_OUT = "bounce_out"
FADE_OUT = "fade_out"
WIPE_OUT = "wipe_out"
BLINDS_OUT = "blinds_out"
BOX_OUT = "box_out"
CHECKERBOARD_OUT = "checkerboard_out"
SPLIT_OUT = "split_out"
DIAGONAL_OUT = "diagonal_out"
RANDOM_BARS_OUT = "random_bars_out"

# -- Emphasis animations --
GROW_SHRINK = "grow_shrink"
SPIN = "spin"
PULSE = "pulse"
COLOR_CHANGE = "color_change"
TEETER = "teeter"
DESATURATE = "desaturate"
DARKEN = "darken"
LIGHTEN = "lighten"
TRANSPARENCY = "transparency"
OBJECT_COLOR = "object_color"
FONT_COLOR = "font_color"
BRUSH_ON_COLOR = "brush_on_color"
BRUSH_ON_UNDERLINE = "brush_on_underline"
WAVE = "wave"
FILL_COLOR = "fill_color"

# -- Motion path --
MOTION_PATH = "motion_path"

# Category sets for validation and convenience
ENTRANCE_TYPES: frozenset[str] = frozenset({
    APPEAR, FLY_IN, FLOAT_UP, ZOOM, GROW_TURN, SWIVEL, BOUNCE,
    FADE_IN, WIPE_IN, BLINDS_IN, BOX_IN, CHECKERBOARD_IN, SPLIT_IN,
    DIAGONAL_IN, RANDOM_BARS_IN, ASCEND, DESCEND, SPIN_IN,
    STRETCH_IN, WHEEL_IN,
})

EXIT_TYPES: frozenset[str] = frozenset({
    DISAPPEAR, FLY_OUT, FLOAT_DOWN, ZOOM_OUT, SHRINK_TURN, SWIVEL_OUT,
    BOUNCE_OUT, FADE_OUT, WIPE_OUT, BLINDS_OUT, BOX_OUT,
    CHECKERBOARD_OUT, SPLIT_OUT, DIAGONAL_OUT, RANDOM_BARS_OUT,
})

EMPHASIS_TYPES: frozenset[str] = frozenset({
    GROW_SHRINK, SPIN, PULSE, COLOR_CHANGE, TEETER, DESATURATE,
    DARKEN, LIGHTEN, TRANSPARENCY, OBJECT_COLOR, FONT_COLOR,
    BRUSH_ON_COLOR, BRUSH_ON_UNDERLINE, WAVE, FILL_COLOR,
})

ALL_ANIMATION_TYPES: frozenset[str] = ENTRANCE_TYPES | EXIT_TYPES | EMPHASIS_TYPES | {MOTION_PATH}

# ---------------------------------------------------------------------------
# 2. OOXML animation class identifiers
# ---------------------------------------------------------------------------

# Mapping from our constant names to OOXML animation preset class identifiers.
# These correspond to the "presetClass" / "presetID" / "presetSubtype" values
# used in <p:animEffect> and <p:anim> elements, as well as the filter
# attribute values for <p:animEffect>.
#
# Format: constant_name -> dict with keys:
#   category: "entrance" | "exit" | "emphasis"
#   filter:   OOXML filter string for <p:animEffect transition="in|out" filter="..."/>
#             Used for entrance/exit animations.
#   presetID: OOXML preset ID (for emphasis animations and some entrance/exit)
#   presetSubtype: OOXML preset subtype (direction/variant)
#
# Reference: ECMA-376 Part 1, Section 20.1.10 (Animation)

ANIMATION_TYPES: dict[str, dict[str, Any]] = {
    # --- Entrance ---
    APPEAR:            {"category": "entrance", "filter": "appear",           "presetID": 1,   "presetSubtype": 0},
    FLY_IN:            {"category": "entrance", "filter": "fly",              "presetID": 2,   "presetSubtype": 4},
    FLOAT_UP:          {"category": "entrance", "filter": "float",            "presetID": 42,  "presetSubtype": 8},
    ZOOM:              {"category": "entrance", "filter": "zoom",             "presetID": 53,  "presetSubtype": 0},
    GROW_TURN:         {"category": "entrance", "filter": "growTurn",         "presetID": 16,  "presetSubtype": 0},
    SWIVEL:            {"category": "entrance", "filter": "swivel",           "presetID": 46,  "presetSubtype": 0},
    BOUNCE:            {"category": "entrance", "filter": "bounce",           "presetID": 10,  "presetSubtype": 0},
    FADE_IN:           {"category": "entrance", "filter": "fade",             "presetID": 10,  "presetSubtype": 0},
    WIPE_IN:           {"category": "entrance", "filter": "wipe",             "presetID": 22,  "presetSubtype": 4},
    BLINDS_IN:         {"category": "entrance", "filter": "blinds",           "presetID": 3,   "presetSubtype": 4},
    BOX_IN:            {"category": "entrance", "filter": "box",              "presetID": 6,   "presetSubtype": 4},
    CHECKERBOARD_IN:   {"category": "entrance", "filter": "checkerboard",     "presetID": 11,  "presetSubtype": 4},
    SPLIT_IN:          {"category": "entrance", "filter": "split",            "presetID": 41,  "presetSubtype": 4},
    DIAGONAL_IN:       {"category": "entrance", "filter": "diagonal",         "presetID": 14,  "presetSubtype": 4},
    RANDOM_BARS_IN:    {"category": "entrance", "filter": "randomBars",       "presetID": 34,  "presetSubtype": 4},
    ASCEND:            {"category": "entrance", "filter": "ascend",           "presetID": 53,  "presetSubtype": 8},
    DESCEND:           {"category": "entrance", "filter": "descend",          "presetID": 53,  "presetSubtype": 16},
    SPIN_IN:           {"category": "entrance", "filter": "spin",             "presetID": 44,  "presetSubtype": 0},
    STRETCH_IN:        {"category": "entrance", "filter": "stretch",          "presetID": 45,  "presetSubtype": 4},
    WHEEL_IN:          {"category": "entrance", "filter": "wheel",            "presetID": 54,  "presetSubtype": 4},
    # --- Exit ---
    DISAPPEAR:         {"category": "exit",     "filter": "appear",           "presetID": 1,   "presetSubtype": 0},
    FLY_OUT:           {"category": "exit",     "filter": "fly",              "presetID": 2,   "presetSubtype": 4},
    FLOAT_DOWN:        {"category": "exit",     "filter": "float",            "presetID": 42,  "presetSubtype": 8},
    ZOOM_OUT:          {"category": "exit",     "filter": "zoom",             "presetID": 53,  "presetSubtype": 0},
    SHRINK_TURN:       {"category": "exit",     "filter": "shrinkTurn",       "presetID": 16,  "presetSubtype": 0},
    SWIVEL_OUT:        {"category": "exit",     "filter": "swivel",           "presetID": 46,  "presetSubtype": 0},
    BOUNCE_OUT:        {"category": "exit",     "filter": "bounce",           "presetID": 10,  "presetSubtype": 0},
    FADE_OUT:          {"category": "exit",     "filter": "fade",             "presetID": 10,  "presetSubtype": 0},
    WIPE_OUT:          {"category": "exit",     "filter": "wipe",             "presetID": 22,  "presetSubtype": 4},
    BLINDS_OUT:        {"category": "exit",     "filter": "blinds",           "presetID": 3,   "presetSubtype": 4},
    BOX_OUT:           {"category": "exit",     "filter": "box",              "presetID": 6,   "presetSubtype": 4},
    CHECKERBOARD_OUT:  {"category": "exit",     "filter": "checkerboard",     "presetID": 11,  "presetSubtype": 4},
    SPLIT_OUT:         {"category": "exit",     "filter": "split",            "presetID": 41,  "presetSubtype": 4},
    DIAGONAL_OUT:      {"category": "exit",     "filter": "diagonal",         "presetID": 14,  "presetSubtype": 4},
    RANDOM_BARS_OUT:   {"category": "exit",     "filter": "randomBars",       "presetID": 34,  "presetSubtype": 4},
    # --- Emphasis ---
    GROW_SHRINK:       {"category": "emphasis", "filter": None,               "presetID": 18,  "presetSubtype": 0},
    SPIN:              {"category": "emphasis", "filter": None,               "presetID": 44,  "presetSubtype": 0},
    PULSE:             {"category": "emphasis", "filter": None,               "presetID": 10,  "presetSubtype": 0},
    COLOR_CHANGE:      {"category": "emphasis", "filter": None,               "presetID": 32,  "presetSubtype": 0},
    TEETER:            {"category": "emphasis", "filter": None,               "presetID": 47,  "presetSubtype": 0},
    DESATURATE:        {"category": "emphasis", "filter": None,               "presetID": 30,  "presetSubtype": 0},
    DARKEN:            {"category": "emphasis", "filter": None,               "presetID": 12,  "presetSubtype": 0},
    LIGHTEN:           {"category": "emphasis", "filter": None,               "presetID": 28,  "presetSubtype": 0},
    TRANSPARENCY:      {"category": "emphasis", "filter": None,               "presetID": 54,  "presetSubtype": 0},
    OBJECT_COLOR:      {"category": "emphasis", "filter": None,               "presetID": 32,  "presetSubtype": 0},
    FONT_COLOR:        {"category": "emphasis", "filter": None,               "presetID": 26,  "presetSubtype": 0},
    BRUSH_ON_COLOR:    {"category": "emphasis", "filter": None,               "presetID": 7,   "presetSubtype": 0},
    BRUSH_ON_UNDERLINE:{"category": "emphasis", "filter": None,               "presetID": 8,   "presetSubtype": 0},
    WAVE:              {"category": "emphasis", "filter": None,               "presetID": 50,  "presetSubtype": 0},
    FILL_COLOR:        {"category": "emphasis", "filter": None,               "presetID": 22,  "presetSubtype": 0},
}

# ---------------------------------------------------------------------------
# 3. OOXML namespaces
# ---------------------------------------------------------------------------

_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ---------------------------------------------------------------------------
# 4. Internal helpers
# ---------------------------------------------------------------------------


def _get_shape_id(shape: Any) -> str:
    """Extract the shape ID string for OOXML references.

    Parameters
    ----------
    shape : pptx.shapes.base.BaseShape
        A python-pptx shape object.

    Returns
    -------
    str
        The shape ID as a string (e.g. ``"2"``).

    Raises
    ------
    ValueError
        If the shape has no valid ID.
    """
    # python-pptx exposes shape.shape_id as an integer
    shape_id = getattr(shape, "shape_id", None)
    if shape_id is None:
        # Fallback: try reading from the XML element directly
        elem = getattr(shape, "_element", None)
        if elem is not None:
            nvSpPr = elem.find(f"{{{_P_NS}}}nvSpPr")
            if nvSpPr is not None:
                cNvPr = nvSpPr.find(f"{{{_A_NS}}}cNvPr")
                if cNvPr is not None:
                    shape_id = cNvPr.get("id")
        if shape_id is None:
            raise ValueError("Shape has no valid ID; cannot apply animation")
    return str(shape_id)


def _ensure_timing(slide_elem: etree._Element) -> etree._Element:
    """Ensure a ``<p:timing>`` element exists on the slide and return it.

    If ``<p:timing>`` already exists, it is returned as-is.  Otherwise a
    new one is created with the standard sequence container structure and
    inserted in the correct position within ``<p:sld>``.

    The standard structure is::

        <p:timing>
          <p:tnLst>
            <p:par>
              <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                <p:childTnLst/>
              </p:cTn>
            </p:par>
          </p:tnLst>
        </p:timing>

    Parameters
    ----------
    slide_elem : lxml.etree._Element
        The ``<p:sld>`` root element.

    Returns
    -------
    lxml.etree._Element
        The ``<p:timing>`` element.
    """
    ns = _P_NS
    timing = slide_elem.find(f"{{{ns}}}timing")
    if timing is not None:
        return timing

    # Build the timing tree from scratch
    timing = etree.Element(f"{{{ns}}}timing")

    tnLst = etree.SubElement(timing, f"{{{ns}}}tnLst")
    root_par = etree.SubElement(tnLst, f"{{{ns}}}par")
    root_cTn = etree.SubElement(root_par, f"{{{ns}}}cTn")
    root_cTn.set("id", "1")
    root_cTn.set("dur", "indefinite")
    root_cTn.set("restart", "never")
    root_cTn.set("nodeType", "tmRoot")
    etree.SubElement(root_cTn, f"{{{ns}}}childTnLst")

    # Insert <p:timing> in the correct position:
    # after <p:transition> (if present), before <p:extLst> (if present),
    # or as the last child of <p:sld>.
    transition = slide_elem.find(f"{{{ns}}}transition")
    extLst = slide_elem.find(f"{{{ns}}}extLst")

    if extLst is not None:
        extLst.addprevious(timing)
    elif transition is not None:
        transition.addnext(timing)
    else:
        slide_elem.append(timing)

    return timing


def _next_id(timing: etree._Element) -> int:
    """Determine the next available animation ID in the timing tree.

    OOXML animation IDs must be unique within a slide.  This function
    scans all ``cTn`` elements in the timing tree and returns the
    maximum ID + 1.

    Parameters
    ----------
    timing : lxml.etree._Element
        The ``<p:timing>`` element.

    Returns
    -------
    int
        The next available ID (minimum 2, since 1 is the root).
    """
    ns = _P_NS
    max_id = 1
    for cTn in timing.iter(f"{{{ns}}}cTn"):
        id_val = cTn.get("id")
        if id_val is not None:
            try:
                max_id = max(max_id, int(id_val))
            except (ValueError, TypeError):
                pass
    return max_id + 1


def _trigger_attrs(trigger: str) -> dict[str, str]:
    """Return OOXML attributes for the given trigger mode.

    Parameters
    ----------
    trigger : str
        One of ``"onClick"``, ``"withPrevious"``, ``"afterPrevious"``.

    Returns
    -------
    dict[str, str]
        Attributes to set on the ``<p:cTn>`` element.
    """
    if trigger == "withPrevious":
        return {"nodeType": "withEffect"}
    elif trigger == "afterPrevious":
        return {"nodeType": "afterEffect"}
    else:
        # Default: onClick
        return {"nodeType": "clickEffect"}


def _build_on_click_condition(ns: str, shape_id: str, next_id: int) -> etree._Element:
    """Build a ``<p:cond>`` element for an on-click trigger.

    Parameters
    ----------
    ns : str
        The PresentationML namespace URI.
    shape_id : str
        The target shape ID.
    next_id : int
        The ID to assign to the condition's ``cTn``.

    Returns
    -------
    lxml.etree._Element
        The ``<p:stCondLst>`` element containing the click condition.
    """
    stCondLst = etree.Element(f"{{{ns}}}stCondLst")
    cond = etree.SubElement(stCondLst, f"{{{ns}}}cond")
    cond.set("delay", "0")
    tgtEl = etree.SubElement(cond, f"{{{ns}}}tgtEl")
    spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
    spTgt.set("spid", shape_id)
    return stCondLst


def _build_after_previous_condition(ns: str) -> etree._Element:
    """Build a ``<p:cond>`` element for an after-previous trigger.

    Parameters
    ----------
    ns : str
        The PresentationML namespace URI.

    Returns
    -------
    lxml.etree._Element
        The ``<p:stCondLst>`` element.
    """
    stCondLst = etree.Element(f"{{{ns}}}stCondLst")
    cond = etree.SubElement(stCondLst, f"{{{ns}}}cond")
    cond.set("delay", "0")
    return stCondLst


def _build_with_previous_condition(ns: str) -> etree._Element:
    """Build a ``<p:cond>`` element for a with-previous trigger.

    Parameters
    ----------
    ns : str
        The PresentationML namespace URI.

    Returns
    -------
    lxml.etree._Element
        The ``<p:stCondLst>`` element.
    """
    stCondLst = etree.Element(f"{{{ns}}}stCondLst")
    cond = etree.SubElement(stCondLst, f"{{{ns}}}cond")
    cond.set("delay", "0")
    return stCondLst


def _build_entrance_exit_xml(
    animation_type: str,
    shape_id: str,
    duration_ms: int,
    delay_ms: int,
    trigger: str,
    next_id_start: int,
) -> etree._Element:
    """Build a ``<p:par>`` element for an entrance or exit animation.

    The structure follows the OOXML pattern::

        <p:par>
          <p:cTn id="N" dur="ms" fill="hold" nodeType="...Effect">
            <p:stCondLst>
              <p:cond delay="ms"/>
            </p:stCondLst>
            <p:childTnLst>
              <p:par>
                <p:cTn id="N+1" fill="hold">
                  <p:stCondLst>
                    <p:cond delay="0"/>
                  </p:stCondLst>
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="N+2" presetID="..." presetClass="entrance|exit"
                             presetSubtype="..." fill="hold" nodeType="...Effect">
                        <p:stCondLst>
                          <p:cond delay="0"/>
                        </p:stCondLst>
                        <p:childTnLst>
                          <p:set .../>
                          <p:animEffect transition="in|out" filter="..."/>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
        </p:par>

    Parameters
    ----------
    animation_type : str
        One of the entrance or exit animation type constants.
    shape_id : str
        The target shape ID string.
    duration_ms : int
        Animation duration in milliseconds.
    delay_ms : int
        Delay before the animation starts, in milliseconds.
    trigger : str
        Trigger mode: ``"onClick"``, ``"withPrevious"``, or ``"afterPrevious"``.
    next_id_start : int
        The starting ID for this animation's ``cTn`` elements.

    Returns
    -------
    lxml.etree._Element
        The constructed ``<p:par>`` element.
    """
    ns = _P_NS
    info = ANIMATION_TYPES[animation_type]
    is_entrance = info["category"] == "entrance"
    transition_dir = "in" if is_entrance else "out"
    node_type = "entr" if is_entrance else "exit"
    filter_val = info["filter"]
    preset_id = info["presetID"]
    preset_subtype = info["presetSubtype"]

    # -- Outer par (sequence node) --
    outer_par = etree.Element(f"{{{ns}}}par")
    outer_cTn = etree.SubElement(outer_par, f"{{{ns}}}cTn")
    outer_cTn.set("id", str(next_id_start))
    outer_cTn.set("dur", str(duration_ms))
    outer_cTn.set("fill", "hold")

    trigger_attrs = _trigger_attrs(trigger)
    outer_cTn.set("nodeType", trigger_attrs["nodeType"])

    # Start condition (delay)
    if trigger == "onClick":
        outer_cTn.append(_build_on_click_condition(ns, shape_id, next_id_start))
    elif trigger == "afterPrevious":
        outer_cTn.append(_build_after_previous_condition(ns))
    else:
        outer_cTn.append(_build_with_previous_condition(ns))

    # -- Middle par --
    childTnLst1 = etree.SubElement(outer_cTn, f"{{{ns}}}childTnLst")
    mid_par = etree.SubElement(childTnLst1, f"{{{ns}}}par")
    mid_cTn = etree.SubElement(mid_par, f"{{{ns}}}cTn")
    mid_cTn.set("id", str(next_id_start + 1))
    mid_cTn.set("fill", "hold")

    mid_stCondLst = etree.SubElement(mid_cTn, f"{{{ns}}}stCondLst")
    mid_cond = etree.SubElement(mid_stCondLst, f"{{{ns}}}cond")
    mid_cond.set("delay", "0")

    # -- Inner par (the actual effect) --
    childTnLst2 = etree.SubElement(mid_cTn, f"{{{ns}}}childTnLst")
    inner_par = etree.SubElement(childTnLst2, f"{{{ns}}}par")
    inner_cTn = etree.SubElement(inner_par, f"{{{ns}}}cTn")
    inner_cTn.set("id", str(next_id_start + 2))
    inner_cTn.set("presetID", str(preset_id))
    inner_cTn.set("presetClass", "entr")
    inner_cTn.set("presetSubtype", str(preset_subtype))
    inner_cTn.set("fill", "hold")
    inner_cTn.set("nodeType", f"{node_type}Effect")

    inner_stCondLst = etree.SubElement(inner_cTn, f"{{{ns}}}stCondLst")
    inner_cond = etree.SubElement(inner_stCondLst, f"{{{ns}}}cond")
    inner_cond.set("delay", str(delay_ms))

    # -- Child list: set + animEffect --
    childTnLst3 = etree.SubElement(inner_cTn, f"{{{ns}}}childTnLst")

    # <p:set> to make shape visible (entrance) or handle visibility
    set_elem = etree.SubElement(childTnLst3, f"{{{ns}}}set")
    set_cBhvr = etree.SubElement(set_elem, f"{{{ns}}}cBhvr")
    set_cTn = etree.SubElement(set_cBhvr, f"{{{ns}}}cTn")
    set_cTn.set("id", str(next_id_start + 3))
    set_cTn.set("dur", "1")
    set_cTn.set("fill", "hold")
    set_stCondLst = etree.SubElement(set_cTn, f"{{{ns}}}stCondLst")
    set_cond = etree.SubElement(set_stCondLst, f"{{{ns}}}cond")
    set_cond.set("delay", "0")

    set_tgtEl = etree.SubElement(set_cBhvr, f"{{{ns}}}tgtEl")
    set_spTgt = etree.SubElement(set_tgtEl, f"{{{ns}}}spTgt")
    set_spTgt.set("spid", shape_id)

    set_attrNameLst = etree.SubElement(set_cBhvr, f"{{{ns}}}attrNameLst")
    set_attrName = etree.SubElement(set_attrNameLst, f"{{{ns}}}attrName")
    set_attrName.text = "style.visibility"

    set_to = etree.SubElement(set_elem, f"{{{ns}}}to")
    set_val = etree.SubElement(set_to, f"{{{ns}}}strVal")
    set_val.set("val", "visible" if is_entrance else "hidden")

    # <p:animEffect>
    animEffect = etree.SubElement(childTnLst3, f"{{{ns}}}animEffect")
    animEffect.set("transition", transition_dir)
    animEffect.set("filter", filter_val)

    animEffect_cBhvr = etree.SubElement(animEffect, f"{{{ns}}}cBhvr")
    animEffect_cTn = etree.SubElement(animEffect_cBhvr, f"{{{ns}}}cTn")
    animEffect_cTn.set("id", str(next_id_start + 4))
    animEffect_cTn.set("dur", str(duration_ms))

    animEffect_tgtEl = etree.SubElement(animEffect_cBhvr, f"{{{ns}}}tgtEl")
    animEffect_spTgt = etree.SubElement(animEffect_tgtEl, f"{{{ns}}}spTgt")
    animEffect_spTgt.set("spid", shape_id)

    return outer_par


def _build_emphasis_xml(
    animation_type: str,
    shape_id: str,
    duration_ms: int,
    delay_ms: int,
    trigger: str,
    repeat: int,
    next_id_start: int,
) -> etree._Element:
    """Build a ``<p:par>`` element for an emphasis animation.

    Emphasis animations use ``<p:anim>`` elements with attribute
    animations (color, scale, rotation, etc.) rather than
    ``<p:animEffect>``.

    Parameters
    ----------
    animation_type : str
        One of the emphasis animation type constants.
    shape_id : str
        The target shape ID string.
    duration_ms : int
        Animation duration in milliseconds.
    delay_ms : int
        Delay before the animation starts, in milliseconds.
    trigger : str
        Trigger mode.
    repeat : int
        Number of times to repeat the animation (0 = indefinite).
    next_id_start : int
        The starting ID for this animation's ``cTn`` elements.

    Returns
    -------
    lxml.etree._Element
        The constructed ``<p:par>`` element.
    """
    ns = _P_NS
    info = ANIMATION_TYPES[animation_type]
    preset_id = info["presetID"]
    preset_subtype = info["presetSubtype"]

    # -- Outer par (sequence node) --
    outer_par = etree.Element(f"{{{ns}}}par")
    outer_cTn = etree.SubElement(outer_par, f"{{{ns}}}cTn")
    outer_cTn.set("id", str(next_id_start))
    outer_cTn.set("dur", str(duration_ms))
    outer_cTn.set("fill", "hold")

    trigger_attrs = _trigger_attrs(trigger)
    outer_cTn.set("nodeType", trigger_attrs["nodeType"])

    if trigger == "onClick":
        outer_cTn.append(_build_on_click_condition(ns, shape_id, next_id_start))
    elif trigger == "afterPrevious":
        outer_cTn.append(_build_after_previous_condition(ns))
    else:
        outer_cTn.append(_build_with_previous_condition(ns))

    # -- Middle par --
    childTnLst1 = etree.SubElement(outer_cTn, f"{{{ns}}}childTnLst")
    mid_par = etree.SubElement(childTnLst1, f"{{{ns}}}par")
    mid_cTn = etree.SubElement(mid_par, f"{{{ns}}}cTn")
    mid_cTn.set("id", str(next_id_start + 1))
    mid_cTn.set("fill", "hold")

    mid_stCondLst = etree.SubElement(mid_cTn, f"{{{ns}}}stCondLst")
    mid_cond = etree.SubElement(mid_stCondLst, f"{{{ns}}}cond")
    mid_cond.set("delay", "0")

    # -- Inner par (the actual effect) --
    childTnLst2 = etree.SubElement(mid_cTn, f"{{{ns}}}childTnLst")
    inner_par = etree.SubElement(childTnLst2, f"{{{ns}}}par")
    inner_cTn = etree.SubElement(inner_par, f"{{{ns}}}cTn")
    inner_cTn.set("id", str(next_id_start + 2))
    inner_cTn.set("presetID", str(preset_id))
    inner_cTn.set("presetClass", "emph")
    inner_cTn.set("presetSubtype", str(preset_subtype))
    inner_cTn.set("fill", "hold")
    inner_cTn.set("nodeType", "withEffect")

    inner_stCondLst = etree.SubElement(inner_cTn, f"{{{ns}}}stCondLst")
    inner_cond = etree.SubElement(inner_stCondLst, f"{{{ns}}}cond")
    inner_cond.set("delay", str(delay_ms))

    if repeat != 1:
        repeat_val = "indefinite" if repeat == 0 else str(repeat)
        inner_cTn.set("repeatCount", repeat_val)

    # -- Child list: emphasis-specific <p:anim> elements --
    childTnLst3 = etree.SubElement(inner_cTn, f"{{{ns}}}childTnLst")

    # Build emphasis-specific animation content
    _build_emphasis_content(
        childTnLst3, animation_type, shape_id, duration_ms, next_id_start + 3
    )

    return outer_par


def _build_emphasis_content(
    parent: etree._Element,
    animation_type: str,
    shape_id: str,
    duration_ms: int,
    next_id: int,
) -> None:
    """Build the inner animation content for an emphasis effect.

    Parameters
    ----------
    parent : lxml.etree._Element
        The ``<p:childTnLst>`` element to append children to.
    animation_type : str
        The emphasis animation type constant.
    shape_id : str
        The target shape ID string.
    duration_ms : int
        Animation duration in milliseconds.
    next_id : int
        The next available cTn ID.
    """
    ns = _P_NS

    if animation_type in (GROW_SHRINK, PULSE):
        # Scale animation using <p:anim> on the transform scale
        anim = etree.SubElement(parent, f"{{{ns}}}anim")
        anim_cBhvr = etree.SubElement(anim, f"{{{ns}}}cBhvr")
        anim_cTn = etree.SubElement(anim_cBhvr, f"{{{ns}}}cTn")
        anim_cTn.set("id", str(next_id))
        anim_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(anim_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(anim_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "ppt_x"

        if animation_type == GROW_SHRINK:
            # Scale from 1.0 to 1.5 and back
            anim.set("by", "1.0,1.5,1.0")
        else:  # PULSE
            anim.set("by", "1.0,1.1,1.0")

    elif animation_type == SPIN:
        # Rotation animation
        anim = etree.SubElement(parent, f"{{{ns}}}anim")
        anim_cBhvr = etree.SubElement(anim, f"{{{ns}}}cBhvr")
        anim_cTn = etree.SubElement(anim_cBhvr, f"{{{ns}}}cTn")
        anim_cTn.set("id", str(next_id))
        anim_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(anim_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(anim_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "r"

        # Full 360-degree rotation
        anim.set("by", "0;360")

    elif animation_type in (COLOR_CHANGE, OBJECT_COLOR, FONT_COLOR, FILL_COLOR, BRUSH_ON_COLOR):
        # Color change animation using <p:animClr>
        animClr = etree.SubElement(parent, f"{{{ns}}}animClr")
        animClr.set("clrSpc", "rgb")

        animClr_cBhvr = etree.SubElement(animClr, f"{{{ns}}}cBhvr")
        animClr_cTn = etree.SubElement(animClr_cBhvr, f"{{{ns}}}cTn")
        animClr_cTn.set("id", str(next_id))
        animClr_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(animClr_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(animClr_cBhvr, f"{{{ns}}}attrNameLst")
        if animation_type == FONT_COLOR:
            attr_name = "style.color"
        elif animation_type == FILL_COLOR:
            attr_name = "fillcolor"
        else:
            attr_name = "style.color"
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = attr_name

        # Default: animate to a highlight color
        to_elem = etree.SubElement(animClr, f"{{{ns}}}to")
        srgbClr = etree.SubElement(to_elem, f"{{{_A_NS}}}srgbClr")
        srgbClr.set("val", "FF0000")

    elif animation_type == TRANSPARENCY:
        # Transparency animation
        anim = etree.SubElement(parent, f"{{{ns}}}anim")
        anim_cBhvr = etree.SubElement(anim, f"{{{ns}}}cBhvr")
        anim_cTn = etree.SubElement(anim_cBhvr, f"{{{ns}}}cTn")
        anim_cTn.set("id", str(next_id))
        anim_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(anim_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(anim_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "style.opacity"

        # Animate from 100% to 50% opacity
        anim.set("from", "1.0")
        anim.set("to", "0.5")

    elif animation_type in (DESATURATE, DARKEN, LIGHTEN):
        # Luminance/contrast animation using <p:animClr>
        animClr = etree.SubElement(parent, f"{{{ns}}}animClr")
        animClr.set("clrSpc", "rgb")

        animClr_cBhvr = etree.SubElement(animClr, f"{{{ns}}}cBhvr")
        animClr_cTn = etree.SubElement(animClr_cBhvr, f"{{{ns}}}cTn")
        animClr_cTn.set("id", str(next_id))
        animClr_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(animClr_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(animClr_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "style.color"

        to_elem = etree.SubElement(animClr, f"{{{ns}}}to")
        srgbClr = etree.SubElement(to_elem, f"{{{_A_NS}}}srgbClr")
        if animation_type == DARKEN:
            srgbClr.set("val", "800000")
        elif animation_type == LIGHTEN:
            srgbClr.set("val", "CCCCCC")
        else:  # DESATURATE
            srgbClr.set("val", "808080")

    elif animation_type == TEETER:
        # Slight rotation oscillation
        anim = etree.SubElement(parent, f"{{{ns}}}anim")
        anim_cBhvr = etree.SubElement(anim, f"{{{ns}}}cBhvr")
        anim_cTn = etree.SubElement(anim_cBhvr, f"{{{ns}}}cTn")
        anim_cTn.set("id", str(next_id))
        anim_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(anim_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(anim_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "r"

        # Oscillate between -5 and 5 degrees
        anim.set("by", "-5;5")

    elif animation_type == WAVE:
        # Wave effect: vertical displacement
        anim = etree.SubElement(parent, f"{{{ns}}}anim")
        anim_cBhvr = etree.SubElement(anim, f"{{{ns}}}cBhvr")
        anim_cTn = etree.SubElement(anim_cBhvr, f"{{{ns}}}cTn")
        anim_cTn.set("id", str(next_id))
        anim_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(anim_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(anim_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "ppt_y"

        anim.set("by", "0;-0.05;0")

    elif animation_type == BRUSH_ON_UNDERLINE:
        # Underline reveal effect
        animEffect = etree.SubElement(parent, f"{{{ns}}}animEffect")
        animEffect.set("transition", "in")
        animEffect.set("filter", "wipe")

        animEffect_cBhvr = etree.SubElement(animEffect, f"{{{ns}}}cBhvr")
        animEffect_cTn = etree.SubElement(animEffect_cBhvr, f"{{{ns}}}cTn")
        animEffect_cTn.set("id", str(next_id))
        animEffect_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(animEffect_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

    else:
        # Generic emphasis: use a simple <p:anim> with a scale pulse
        anim = etree.SubElement(parent, f"{{{ns}}}anim")
        anim_cBhvr = etree.SubElement(anim, f"{{{ns}}}cBhvr")
        anim_cTn = etree.SubElement(anim_cBhvr, f"{{{ns}}}cTn")
        anim_cTn.set("id", str(next_id))
        anim_cTn.set("dur", str(duration_ms))

        tgtEl = etree.SubElement(anim_cBhvr, f"{{{ns}}}tgtEl")
        spTgt = etree.SubElement(tgtEl, f"{{{ns}}}spTgt")
        spTgt.set("spid", shape_id)

        attrNameLst = etree.SubElement(anim_cBhvr, f"{{{ns}}}attrNameLst")
        attrName = etree.SubElement(attrNameLst, f"{{{ns}}}attrName")
        attrName.text = "ppt_x"


def _build_motion_path_xml(
    path_points: list[tuple[float, float]],
    shape_id: str,
    duration_ms: int,
    delay_ms: int,
    trigger: str,
    repeat: int,
    next_id_start: int,
) -> etree._Element:
    """Build a ``<p:par>`` element for a motion path animation.

    Parameters
    ----------
    path_points : list[tuple[float, float]]
        A list of (x, y) coordinate pairs defining the path.  The first
        point should typically be (0, 0) representing the shape's
        starting position.  Coordinates are in OOXML units (1/60000 of
        a slide width for relative, or EMU for absolute).
    shape_id : str
        The target shape ID string.
    duration_ms : int
        Animation duration in milliseconds.
    delay_ms : int
        Delay before the animation starts, in milliseconds.
    trigger : str
        Trigger mode.
    repeat : int
        Number of times to repeat (0 = indefinite).
    next_id_start : int
        The starting ID for this animation's ``cTn`` elements.

    Returns
    -------
    lxml.etree._Element
        The constructed ``<p:par>`` element.
    """
    ns = _P_NS

    # Build SVG-like path string from points
    if not path_points:
        log.warning("Empty path_points for motion path animation; skipping")
        return etree.Element(f"{{{ns}}}par")  # empty par

    path_str_parts: list[str] = []
    for i, (x, y) in enumerate(path_points):
        if i == 0:
            path_str_parts.append(f"M {x} {y}")
        else:
            path_str_parts.append(f"L {x} {y}")

    path_str = " ".join(path_str_parts)

    # -- Outer par --
    outer_par = etree.Element(f"{{{ns}}}par")
    outer_cTn = etree.SubElement(outer_par, f"{{{ns}}}cTn")
    outer_cTn.set("id", str(next_id_start))
    outer_cTn.set("dur", str(duration_ms))
    outer_cTn.set("fill", "hold")

    trigger_attrs = _trigger_attrs(trigger)
    outer_cTn.set("nodeType", trigger_attrs["nodeType"])

    if trigger == "onClick":
        outer_cTn.append(_build_on_click_condition(ns, shape_id, next_id_start))
    elif trigger == "afterPrevious":
        outer_cTn.append(_build_after_previous_condition(ns))
    else:
        outer_cTn.append(_build_with_previous_condition(ns))

    # -- Middle par --
    childTnLst1 = etree.SubElement(outer_cTn, f"{{{ns}}}childTnLst")
    mid_par = etree.SubElement(childTnLst1, f"{{{ns}}}par")
    mid_cTn = etree.SubElement(mid_par, f"{{{ns}}}cTn")
    mid_cTn.set("id", str(next_id_start + 1))
    mid_cTn.set("fill", "hold")

    mid_stCondLst = etree.SubElement(mid_cTn, f"{{{ns}}}stCondLst")
    mid_cond = etree.SubElement(mid_stCondLst, f"{{{ns}}}cond")
    mid_cond.set("delay", "0")

    # -- Inner par with <p:animMotion> --
    childTnLst2 = etree.SubElement(mid_cTn, f"{{{ns}}}childTnLst")
    inner_par = etree.SubElement(childTnLst2, f"{{{ns}}}par")
    inner_cTn = etree.SubElement(inner_par, f"{{{ns}}}cTn")
    inner_cTn.set("id", str(next_id_start + 2))
    inner_cTn.set("presetID", "1")
    inner_cTn.set("presetClass", "path")
    inner_cTn.set("presetSubtype", "0")
    inner_cTn.set("fill", "hold")
    inner_cTn.set("nodeType", "pathEffect")

    inner_stCondLst = etree.SubElement(inner_cTn, f"{{{ns}}}stCondLst")
    inner_cond = etree.SubElement(inner_stCondLst, f"{{{ns}}}cond")
    inner_cond.set("delay", str(delay_ms))

    if repeat != 1:
        repeat_val = "indefinite" if repeat == 0 else str(repeat)
        inner_cTn.set("repeatCount", repeat_val)

    childTnLst3 = etree.SubElement(inner_cTn, f"{{{ns}}}childTnLst")

    # <p:animMotion>
    animMotion = etree.SubElement(childTnLst3, f"{{{ns}}}animMotion")
    animMotion.set("path", path_str)
    animMotion.set("origin", "layout")

    animMotion_cBhvr = etree.SubElement(animMotion, f"{{{ns}}}cBhvr")
    animMotion_cTn = etree.SubElement(animMotion_cBhvr, f"{{{ns}}}cTn")
    animMotion_cTn.set("id", str(next_id_start + 3))
    animMotion_cTn.set("dur", str(duration_ms))

    animMotion_tgtEl = etree.SubElement(animMotion_cBhvr, f"{{{ns}}}tgtEl")
    animMotion_spTgt = etree.SubElement(animMotion_tgtEl, f"{{{ns}}}spTgt")
    animMotion_spTgt.set("spid", shape_id)

    return outer_par


# ---------------------------------------------------------------------------
# 5. Public API
# ---------------------------------------------------------------------------


def apply_animation(
    slide: Any,
    shape: Any,
    animation_type: str,
    *,
    duration_ms: int = 500,
    delay_ms: int = 0,
    trigger: str = "onClick",
    repeat: int = 1,
) -> None:
    """Apply an animation to a shape on a slide.

    This is the primary entry point for adding animations.  It
    determines the animation category (entrance, exit, emphasis, or
    motion path) and delegates to the appropriate builder.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide containing the shape.
    shape : pptx.shapes.base.BaseShape
        The shape to animate.
    animation_type : str
        One of the animation type constants defined in this module.
    duration_ms : int
        Animation duration in milliseconds (default 500).
    delay_ms : int
        Delay before the animation starts, in milliseconds (default 0).
    trigger : str
        Trigger mode: ``"onClick"`` (default), ``"withPrevious"``,
        or ``"afterPrevious"``.
    repeat : int
        Number of times to repeat the animation.  Use 0 for indefinite
        repeat.  Default is 1 (play once).

    Raises
    ------
    ValueError
        If the shape has no valid ID.
    """
    # Validate animation type
    if animation_type not in ALL_ANIMATION_TYPES:
        warnings.warn(
            f"Unknown animation type {animation_type!r}; skipping. "
            f"Valid types: {', '.join(sorted(ALL_ANIMATION_TYPES))}",
            stacklevel=2,
        )
        return

    # Validate trigger
    if trigger not in ("onClick", "withPrevious", "afterPrevious"):
        log.warning(
            "Unknown trigger %r; falling back to 'onClick'. "
            "Valid triggers: onClick, withPrevious, afterPrevious",
            trigger,
        )
        trigger = "onClick"

    # Get shape ID
    try:
        shape_id = _get_shape_id(shape)
    except ValueError:
        log.warning("Shape has no valid ID; cannot apply animation")
        return

    # Get or create the <p:timing> element
    sld = slide._element
    timing = _ensure_timing(sld)

    # Find the root childTnLst where animation sequences are added
    ns = _P_NS
    root_cTn = timing.find(f".//{{{ns}}}cTn[@nodeType='tmRoot']")
    if root_cTn is None:
        log.warning("Could not find root cTn in timing tree; cannot apply animation")
        return

    childTnLst = root_cTn.find(f"{{{ns}}}childTnLst")
    if childTnLst is None:
        childTnLst = etree.SubElement(root_cTn, f"{{{ns}}}childTnLst")

    # Determine next available ID
    next_id = _next_id(timing)

    # Build the appropriate animation XML
    if animation_type in ENTRANCE_TYPES or animation_type in EXIT_TYPES:
        anim_par = _build_entrance_exit_xml(
            animation_type, shape_id, duration_ms, delay_ms, trigger, next_id
        )
    elif animation_type in EMPHASIS_TYPES:
        anim_par = _build_emphasis_xml(
            animation_type, shape_id, duration_ms, delay_ms, trigger, repeat, next_id
        )
    elif animation_type == MOTION_PATH:
        # Motion path requires path_points; use a default if called directly
        log.warning(
            "Use apply_motion_path() for motion path animations; "
            "applying a default no-op path"
        )
        anim_par = _build_motion_path_xml(
            [(0, 0)], shape_id, duration_ms, delay_ms, trigger, repeat, next_id
        )
    else:
        # Should not reach here after validation, but be defensive
        warnings.warn(f"Unhandled animation type {animation_type!r}; skipping", stacklevel=2)
        return

    # Append the animation to the childTnLst
    childTnLst.append(anim_par)


def apply_entrance_animation(
    slide: Any,
    shape: Any,
    anim_type: str,
    **kwargs: Any,
) -> None:
    """Apply an entrance animation to a shape.

    Convenience wrapper that validates the animation type is an
    entrance animation before delegating to :func:`apply_animation`.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide containing the shape.
    shape : pptx.shapes.base.BaseShape
        The shape to animate.
    anim_type : str
        One of the entrance animation type constants.
    **kwargs
        Additional keyword arguments passed to :func:`apply_animation`
        (``duration_ms``, ``delay_ms``, ``trigger``, ``repeat``).
    """
    if anim_type not in ENTRANCE_TYPES:
        warnings.warn(
            f"{anim_type!r} is not an entrance animation; skipping. "
            f"Valid entrance types: {', '.join(sorted(ENTRANCE_TYPES))}",
            stacklevel=2,
        )
        return
    apply_animation(slide, shape, anim_type, **kwargs)


def apply_exit_animation(
    slide: Any,
    shape: Any,
    anim_type: str,
    **kwargs: Any,
) -> None:
    """Apply an exit animation to a shape.

    Convenience wrapper that validates the animation type is an
    exit animation before delegating to :func:`apply_animation`.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide containing the shape.
    shape : pptx.shapes.base.BaseShape
        The shape to animate.
    anim_type : str
        One of the exit animation type constants.
    **kwargs
        Additional keyword arguments passed to :func:`apply_animation`
        (``duration_ms``, ``delay_ms``, ``trigger``, ``repeat``).
    """
    if anim_type not in EXIT_TYPES:
        warnings.warn(
            f"{anim_type!r} is not an exit animation; skipping. "
            f"Valid exit types: {', '.join(sorted(EXIT_TYPES))}",
            stacklevel=2,
        )
        return
    apply_animation(slide, shape, anim_type, **kwargs)


def apply_emphasis_animation(
    slide: Any,
    shape: Any,
    anim_type: str,
    **kwargs: Any,
) -> None:
    """Apply an emphasis animation to a shape.

    Convenience wrapper that validates the animation type is an
    emphasis animation before delegating to :func:`apply_animation`.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide containing the shape.
    shape : pptx.shapes.base.BaseShape
        The shape to animate.
    anim_type : str
        One of the emphasis animation type constants.
    **kwargs
        Additional keyword arguments passed to :func:`apply_animation`
        (``duration_ms``, ``delay_ms``, ``trigger``, ``repeat``).
    """
    if anim_type not in EMPHASIS_TYPES:
        warnings.warn(
            f"{anim_type!r} is not an emphasis animation; skipping. "
            f"Valid emphasis types: {', '.join(sorted(EMPHASIS_TYPES))}",
            stacklevel=2,
        )
        return
    apply_animation(slide, shape, anim_type, **kwargs)


def apply_motion_path(
    slide: Any,
    shape: Any,
    path_points: list[tuple[float, float]],
    *,
    duration_ms: int = 500,
    delay_ms: int = 0,
    trigger: str = "onClick",
    repeat: int = 1,
) -> None:
    """Apply a motion path animation to a shape.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide containing the shape.
    shape : pptx.shapes.base.BaseShape
        The shape to animate.
    path_points : list[tuple[float, float]]
        A list of (x, y) coordinate pairs defining the motion path.
        The first point should typically be (0, 0) representing the
        shape's starting position.  Coordinates are in OOXML relative
        units where 1.0 = full slide width/height.
    duration_ms : int
        Animation duration in milliseconds (default 500).
    delay_ms : int
        Delay before the animation starts, in milliseconds (default 0).
    trigger : str
        Trigger mode: ``"onClick"`` (default), ``"withPrevious"``,
        or ``"afterPrevious"``.
    repeat : int
        Number of times to repeat the animation.  Use 0 for indefinite
        repeat.  Default is 1 (play once).

    Raises
    ------
    ValueError
        If the shape has no valid ID.
    """
    # Validate trigger
    if trigger not in ("onClick", "withPrevious", "afterPrevious"):
        log.warning(
            "Unknown trigger %r; falling back to 'onClick'. "
            "Valid triggers: onClick, withPrevious, afterPrevious",
            trigger,
        )
        trigger = "onClick"

    # Get shape ID
    try:
        shape_id = _get_shape_id(shape)
    except ValueError:
        log.warning("Shape has no valid ID; cannot apply motion path animation")
        return

    # Get or create the <p:timing> element
    sld = slide._element
    timing = _ensure_timing(sld)

    # Find the root childTnLst
    ns = _P_NS
    root_cTn = timing.find(f".//{{{ns}}}cTn[@nodeType='tmRoot']")
    if root_cTn is None:
        log.warning("Could not find root cTn in timing tree; cannot apply animation")
        return

    childTnLst = root_cTn.find(f"{{{ns}}}childTnLst")
    if childTnLst is None:
        childTnLst = etree.SubElement(root_cTn, f"{{{ns}}}childTnLst")

    # Determine next available ID
    next_id = _next_id(timing)

    # Build the motion path XML
    anim_par = _build_motion_path_xml(
        path_points, shape_id, duration_ms, delay_ms, trigger, repeat, next_id
    )

    # Append the animation to the childTnLst
    childTnLst.append(anim_par)


def remove_animations(slide: Any) -> None:
    """Remove all animations from a slide.

    This removes the entire ``<p:timing>`` element from the slide's
    XML, which eliminates all animation sequences.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide to remove animations from.
    """
    ns = _P_NS
    sld = slide._element
    timing = sld.find(f"{{{ns}}}timing")
    if timing is not None:
        sld.remove(timing)


def list_animations(slide: Any) -> list[dict[str, Any]]:
    """List all animations on a slide.

    Parses the ``<p:timing>`` element and returns a list of dictionaries
    describing each animation found.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide to inspect.

    Returns
    -------
    list[dict[str, Any]]
        A list of animation descriptors.  Each dictionary contains:

        - ``"shape_id"``: the target shape ID (str)
        - ``"type"``: the animation category (``"entrance"``, ``"exit"``,
          ``"emphasis"``, or ``"motion_path"``)
        - ``"filter"``: the OOXML filter string, if applicable
        - ``"duration_ms"``: the animation duration in milliseconds (int or None)
        - ``"delay_ms"``: the delay in milliseconds (int or None)
        - ``"trigger"``: the trigger mode (``"onClick"``,
          ``"withPrevious"``, or ``"afterPrevious"``)
        - ``"preset_id"``: the OOXML preset ID (int or None)
        - ``"path"``: the motion path string, if applicable
    """
    ns = _P_NS
    sld = slide._element
    timing = sld.find(f"{{{ns}}}timing")
    if timing is None:
        return []

    results: list[dict[str, Any]] = []

    # Walk through the timing tree looking for animation effects
    for animEffect in timing.iter(f"{{{ns}}}animEffect"):
        entry: dict[str, Any] = {
            "shape_id": None,
            "type": None,
            "filter": animEffect.get("filter"),
            "duration_ms": None,
            "delay_ms": None,
            "trigger": None,
            "preset_id": None,
            "path": None,
        }

        # Determine entrance vs exit from transition attribute
        transition = animEffect.get("transition", "in")
        if transition == "in":
            entry["type"] = "entrance"
        elif transition == "out":
            entry["type"] = "exit"

        # Get shape ID from cBhvr > tgtEl > spTgt
        cBhvr = animEffect.find(f"{{{ns}}}cBhvr")
        if cBhvr is not None:
            tgtEl = cBhvr.find(f"{{{ns}}}tgtEl")
            if tgtEl is not None:
                spTgt = tgtEl.find(f"{{{ns}}}spTgt")
                if spTgt is not None:
                    entry["shape_id"] = spTgt.get("spid")

            # Get duration from cTn
            cTn = cBhvr.find(f"{{{ns}}}cTn")
            if cTn is not None:
                dur = cTn.get("dur")
                if dur is not None:
                    try:
                        entry["duration_ms"] = int(dur)
                    except (ValueError, TypeError):
                        pass

        # Walk up to find the parent cTn with nodeType for trigger info
        parent = animEffect.getparent()
        while parent is not None:
            if parent.tag == f"{{{ns}}}cTn":
                node_type = parent.get("nodeType", "")
                if "click" in node_type:
                    entry["trigger"] = "onClick"
                elif "after" in node_type:
                    entry["trigger"] = "afterPrevious"
                elif "with" in node_type:
                    entry["trigger"] = "withPrevious"
                elif "entr" in node_type:
                    entry["type"] = "entrance"
                elif "exit" in node_type:
                    entry["type"] = "exit"

                # Get delay from stCondLst
                stCondLst = parent.find(f"{{{ns}}}stCondLst")
                if stCondLst is not None:
                    cond = stCondLst.find(f"{{{ns}}}cond")
                    if cond is not None:
                        delay = cond.get("delay")
                        if delay is not None:
                            try:
                                entry["delay_ms"] = int(delay)
                            except (ValueError, TypeError):
                                pass

                # Get presetID
                preset_id = parent.get("presetID")
                if preset_id is not None:
                    try:
                        entry["preset_id"] = int(preset_id)
                    except (ValueError, TypeError):
                        pass
                break
            parent = parent.getparent()

        results.append(entry)

    # Look for emphasis animations (<p:anim> and <p:animClr>)
    for anim in timing.iter(f"{{{ns}}}anim"):
        # Skip anim elements that are children of animEffect or animMotion
        parent = anim.getparent()
        if parent is not None and parent.tag == f"{{{ns}}}animEffect":
            continue

        entry = {
            "shape_id": None,
            "type": "emphasis",
            "filter": None,
            "duration_ms": None,
            "delay_ms": None,
            "trigger": None,
            "preset_id": None,
            "path": None,
        }

        cBhvr = anim.find(f"{{{ns}}}cBhvr")
        if cBhvr is not None:
            tgtEl = cBhvr.find(f"{{{ns}}}tgtEl")
            if tgtEl is not None:
                spTgt = tgtEl.find(f"{{{ns}}}spTgt")
                if spTgt is not None:
                    entry["shape_id"] = spTgt.get("spid")

            cTn = cBhvr.find(f"{{{ns}}}cTn")
            if cTn is not None:
                dur = cTn.get("dur")
                if dur is not None:
                    try:
                        entry["duration_ms"] = int(dur)
                    except (ValueError, TypeError):
                        pass

        # Walk up for trigger/delay/preset info
        parent = anim.getparent()
        while parent is not None:
            if parent.tag == f"{{{ns}}}cTn":
                node_type = parent.get("nodeType", "")
                if "click" in node_type:
                    entry["trigger"] = "onClick"
                elif "after" in node_type:
                    entry["trigger"] = "afterPrevious"
                elif "with" in node_type:
                    entry["trigger"] = "withPrevious"

                stCondLst = parent.find(f"{{{ns}}}stCondLst")
                if stCondLst is not None:
                    cond = stCondLst.find(f"{{{ns}}}cond")
                    if cond is not None:
                        delay = cond.get("delay")
                        if delay is not None:
                            try:
                                entry["delay_ms"] = int(delay)
                            except (ValueError, TypeError):
                                pass

                preset_id = parent.get("presetID")
                if preset_id is not None:
                    try:
                        entry["preset_id"] = int(preset_id)
                    except (ValueError, TypeError):
                        pass
                break
            parent = parent.getparent()

        results.append(entry)

    # Look for color animations (<p:animClr>)
    for animClr in timing.iter(f"{{{ns}}}animClr"):
        entry = {
            "shape_id": None,
            "type": "emphasis",
            "filter": None,
            "duration_ms": None,
            "delay_ms": None,
            "trigger": None,
            "preset_id": None,
            "path": None,
        }

        cBhvr = animClr.find(f"{{{ns}}}cBhvr")
        if cBhvr is not None:
            tgtEl = cBhvr.find(f"{{{ns}}}tgtEl")
            if tgtEl is not None:
                spTgt = tgtEl.find(f"{{{ns}}}spTgt")
                if spTgt is not None:
                    entry["shape_id"] = spTgt.get("spid")

            cTn = cBhvr.find(f"{{{ns}}}cTn")
            if cTn is not None:
                dur = cTn.get("dur")
                if dur is not None:
                    try:
                        entry["duration_ms"] = int(dur)
                    except (ValueError, TypeError):
                        pass

        # Walk up for trigger/delay/preset info
        parent = animClr.getparent()
        while parent is not None:
            if parent.tag == f"{{{ns}}}cTn":
                node_type = parent.get("nodeType", "")
                if "click" in node_type:
                    entry["trigger"] = "onClick"
                elif "after" in node_type:
                    entry["trigger"] = "afterPrevious"
                elif "with" in node_type:
                    entry["trigger"] = "withPrevious"

                stCondLst = parent.find(f"{{{ns}}}stCondLst")
                if stCondLst is not None:
                    cond = stCondLst.find(f"{{{ns}}}cond")
                    if cond is not None:
                        delay = cond.get("delay")
                        if delay is not None:
                            try:
                                entry["delay_ms"] = int(delay)
                            except (ValueError, TypeError):
                                pass

                preset_id = parent.get("presetID")
                if preset_id is not None:
                    try:
                        entry["preset_id"] = int(preset_id)
                    except (ValueError, TypeError):
                        pass
                break
            parent = parent.getparent()

        results.append(entry)

    # Look for motion path animations (<p:animMotion>)
    for animMotion in timing.iter(f"{{{ns}}}animMotion"):
        entry = {
            "shape_id": None,
            "type": "motion_path",
            "filter": None,
            "duration_ms": None,
            "delay_ms": None,
            "trigger": None,
            "preset_id": None,
            "path": animMotion.get("path"),
        }

        cBhvr = animMotion.find(f"{{{ns}}}cBhvr")
        if cBhvr is not None:
            tgtEl = cBhvr.find(f"{{{ns}}}tgtEl")
            if tgtEl is not None:
                spTgt = tgtEl.find(f"{{{ns}}}spTgt")
                if spTgt is not None:
                    entry["shape_id"] = spTgt.get("spid")

            cTn = cBhvr.find(f"{{{ns}}}cTn")
            if cTn is not None:
                dur = cTn.get("dur")
                if dur is not None:
                    try:
                        entry["duration_ms"] = int(dur)
                    except (ValueError, TypeError):
                        pass

        # Walk up for trigger/delay/preset info
        parent = animMotion.getparent()
        while parent is not None:
            if parent.tag == f"{{{ns}}}cTn":
                node_type = parent.get("nodeType", "")
                if "click" in node_type:
                    entry["trigger"] = "onClick"
                elif "after" in node_type:
                    entry["trigger"] = "afterPrevious"
                elif "with" in node_type:
                    entry["trigger"] = "withPrevious"

                stCondLst = parent.find(f"{{{ns}}}stCondLst")
                if stCondLst is not None:
                    cond = stCondLst.find(f"{{{ns}}}cond")
                    if cond is not None:
                        delay = cond.get("delay")
                        if delay is not None:
                            try:
                                entry["delay_ms"] = int(delay)
                            except (ValueError, TypeError):
                                pass

                preset_id = parent.get("presetID")
                if preset_id is not None:
                    try:
                        entry["preset_id"] = int(preset_id)
                    except (ValueError, TypeError):
                        pass
                break
            parent = parent.getparent()

        results.append(entry)

    return results


# ---------------------------------------------------------------------------
# 5b. Path-based convenience overloads (accept file path, 1-based slide_index)
# ---------------------------------------------------------------------------

def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path):
    """Open a Presentation from *prs_or_path* (path or Presentation)."""
    from pptx import Presentation
    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def add_animation(
    prs_or_path,
    slide_index: int,
    shape_name: str | None = None,
    anim_type: str = "fade_in",
    shape_index: int | None = None,
    **kwargs: Any,
) -> bool:
    """Apply an animation to a shape by 1-based slide index and shape name/index.

    Accepts either an open ``Presentation`` object or a file path string.
    When given a path, opens the file, applies the animation, and saves.

    Parameters
    ----------
    prs_or_path : Presentation or str
        An open Presentation object, or a path to a .pptx file.
    slide_index : int
        1-based slide index (1 = first slide).
    shape_name : str, optional
        Name of the shape to animate. If None, uses shape_index.
    anim_type : str
        Animation type constant (default ``"fade_in"``).
    shape_index : int, optional
        0-based index into slide.shapes if shape_name is None.
    **kwargs
        Additional animation parameters (``duration_ms``, ``delay_ms``, etc.).

    Returns
    -------
    bool
        True if animation was applied, False if shape was not found.
    """
    is_path = not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = None
        if shape_name is not None:
            for s in slide.shapes:
                if s.name == shape_name:
                    shape = s
                    break
        elif shape_index is not None:
            shapes = list(slide.shapes)
            if 0 <= shape_index < len(shapes):
                shape = shapes[shape_index]
        if shape is None:
            log.warning("Shape not found (name=%r, index=%r)", shape_name, shape_index)
            return False
        apply_animation(slide, shape, anim_type, **kwargs)
        return True
    finally:
        if is_path:
            prs.save(str(prs_or_path))


def add_entrance_animation(
    prs_or_path,
    slide_index: int,
    shape_name: str | None = None,
    anim_type: str = "fade_in",
    shape_index: int | None = None,
    **kwargs: Any,
) -> bool:
    """Apply an entrance animation by 1-based slide index (path or Presentation).

    Convenience path-based overload of :func:`apply_entrance_animation`.
    See :func:`add_animation` for parameter details.
    """
    is_path = not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = None
        if shape_name is not None:
            for s in slide.shapes:
                if s.name == shape_name:
                    shape = s
                    break
        elif shape_index is not None:
            shapes = list(slide.shapes)
            if 0 <= shape_index < len(shapes):
                shape = shapes[shape_index]
        if shape is None:
            return False
        apply_entrance_animation(slide, shape, anim_type, **kwargs)
        return True
    finally:
        if is_path:
            prs.save(str(prs_or_path))


# ---------------------------------------------------------------------------
# 6. __all__
# ---------------------------------------------------------------------------

__all__ = [
    # Core functions
    "apply_animation",
    "apply_entrance_animation",
    "apply_exit_animation",
    # Path-based overloads
    "add_animation",
    "add_entrance_animation",
    "apply_emphasis_animation",
    "apply_motion_path",
    "remove_animations",
    "list_animations",
    # Animation type mapping
    "ANIMATION_TYPES",
    # Category sets
    "ENTRANCE_TYPES",
    "EXIT_TYPES",
    "EMPHASIS_TYPES",
    "ALL_ANIMATION_TYPES",
    # Entrance constants
    "APPEAR",
    "FLY_IN",
    "FLOAT_UP",
    "ZOOM",
    "GROW_TURN",
    "SWIVEL",
    "BOUNCE",
    "FADE_IN",
    "WIPE_IN",
    "BLINDS_IN",
    "BOX_IN",
    "CHECKERBOARD_IN",
    "SPLIT_IN",
    "DIAGONAL_IN",
    "RANDOM_BARS_IN",
    "ASCEND",
    "DESCEND",
    "SPIN_IN",
    "STRETCH_IN",
    "WHEEL_IN",
    # Exit constants
    "DISAPPEAR",
    "FLY_OUT",
    "FLOAT_DOWN",
    "ZOOM_OUT",
    "SHRINK_TURN",
    "SWIVEL_OUT",
    "BOUNCE_OUT",
    "FADE_OUT",
    "WIPE_OUT",
    "BLINDS_OUT",
    "BOX_OUT",
    "CHECKERBOARD_OUT",
    "SPLIT_OUT",
    "DIAGONAL_OUT",
    "RANDOM_BARS_OUT",
    # Emphasis constants
    "GROW_SHRINK",
    "SPIN",
    "PULSE",
    "COLOR_CHANGE",
    "TEETER",
    "DESATURATE",
    "DARKEN",
    "LIGHTEN",
    "TRANSPARENCY",
    "OBJECT_COLOR",
    "FONT_COLOR",
    "BRUSH_ON_COLOR",
    "BRUSH_ON_UNDERLINE",
    "WAVE",
    "FILL_COLOR",
    # Motion path
    "MOTION_PATH",
]
