"""Slide transition support via OOXML manipulation.

python-pptx has no native API for slide transitions, so this module
constructs the appropriate ``<p:transition>`` XML elements directly
using lxml and inserts them into each slide's ``<p:sld>`` element.

Usage
-----
>>> from pptx_skill.transitions import apply_slide_transition, FADE
>>> apply_slide_transition(slide, FADE, duration_ms=700)

Or apply a consistent transition to every slide in a deck:

>>> from pptx_skill.transitions import apply_deck_transitions, PUSH_LEFT
>>> apply_deck_transitions(prs, PUSH_LEFT, duration_ms=500)
"""
from __future__ import annotations

import logging
from typing import Any

from lxml import etree

from pptx_skill._io import save_prs as _save_prs_impl
from pptx_skill.constants import P_NS as _P_NS

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 1. Transition type constants
# ---------------------------------------------------------------------------

FADE = "fade"
PUSH_LEFT = "push_left"
PUSH_RIGHT = "push_right"
PUSH_UP = "push_up"
PUSH_DOWN = "push_down"
WIPE_LEFT = "wipe_left"
WIPE_RIGHT = "wipe_right"
WIPE_UP = "wipe_up"
WIPE_DOWN = "wipe_down"
COVER_LEFT = "cover_left"
COVER_RIGHT = "cover_right"
SPLIT_HORIZONTAL_IN = "split_horizontal_in"
SPLIT_HORIZONTAL_OUT = "split_horizontal_out"
SPLIT_VERTICAL_IN = "split_vertical_in"
SPLIT_VERTICAL_OUT = "split_vertical_out"
DISSOLVE = "dissolve"
RANDOM = "random"
CUT = "cut"

# All valid transition types for validation
TRANSITION_TYPES: frozenset[str] = frozenset({
    FADE, PUSH_LEFT, PUSH_RIGHT, PUSH_UP, PUSH_DOWN,
    WIPE_LEFT, WIPE_RIGHT, WIPE_UP, WIPE_DOWN,
    COVER_LEFT, COVER_RIGHT,
    SPLIT_HORIZONTAL_IN, SPLIT_HORIZONTAL_OUT,
    SPLIT_VERTICAL_IN, SPLIT_VERTICAL_OUT,
    DISSOLVE, RANDOM, CUT,
})

# ---------------------------------------------------------------------------
# Internal mapping: transition_type -> (child_tag, extra_attrs)
# ---------------------------------------------------------------------------

# Speed presets mapped from approximate millisecond ranges
_SPEED_MAP: dict[int, str] = {
    1500: "slow",
    700: "med",
    300: "fast",
}


def _ms_to_spd(duration_ms: int) -> str | None:
    """Convert a duration in milliseconds to an OOXML ``spd`` attribute value.

    Returns one of ``"slow"``, ``"med"``, ``"fast"``, or ``None`` if the
    value does not match a standard preset (in which case the caller should
    use the ``dur`` attribute with the exact millisecond value instead).
    """
    # Find the closest preset
    closest = min(_SPEED_MAP.keys(), key=lambda s: abs(s - duration_ms))
    if abs(closest - duration_ms) <= 200:
        return _SPEED_MAP[closest]
    return None


def _build_transition_xml(
    transition_type: str,
    duration_ms: int,
    advance_ms: int | None,
) -> etree._Element:
    """Build a ``<p:transition>`` XML element.

    Parameters
    ----------
    transition_type : str
        One of the :data:`TRANSITION_TYPES` constants.
    duration_ms : int
        Transition duration in milliseconds.
    advance_ms : int | None
        Auto-advance time in milliseconds, or ``None`` for manual advance.

    Returns
    -------
    lxml.etree._Element
        The constructed ``<p:transition>`` element.
    """
    ns = _P_NS
    transition = etree.Element(f"{{{ns}}}transition")

    # --- Duration / speed ---
    spd = _ms_to_spd(duration_ms)
    if spd is not None:
        transition.set("spd", spd)
    else:
        # Use exact duration in milliseconds
        transition.set("dur", str(duration_ms))

    # --- Auto-advance ---
    if advance_ms is not None:
        transition.set("advTm", str(advance_ms))

    # --- Transition child element ---
    if transition_type == FADE:
        etree.SubElement(transition, f"{{{ns}}}fade")

    elif transition_type == CUT:
        # CUT is the absence of any child element — just the <p:transition>
        # container with no inner transition tag.  We can optionally set
        # the "click" attribute but leaving it bare is sufficient.
        pass

    elif transition_type.startswith("push_"):
        direction = transition_type.split("_", 1)[1]  # left, right, up, down
        dir_attr = {"left": "l", "right": "r", "up": "u", "down": "d"}[direction]
        push = etree.SubElement(transition, f"{{{ns}}}push")
        push.set("dir", dir_attr)

    elif transition_type.startswith("wipe_"):
        direction = transition_type.split("_", 1)[1]
        dir_attr = {"left": "l", "right": "r", "up": "u", "down": "d"}[direction]
        wipe = etree.SubElement(transition, f"{{{ns}}}wipe")
        wipe.set("dir", dir_attr)

    elif transition_type.startswith("cover_"):
        direction = transition_type.split("_", 1)[1]
        dir_attr = {"left": "l", "right": "r"}[direction]
        cover = etree.SubElement(transition, f"{{{ns}}}cover")
        cover.set("dir", dir_attr)

    elif transition_type.startswith("split_horizontal_"):
        orient = "horz"
        inout = "in" if transition_type.endswith("_in") else "out"
        split = etree.SubElement(transition, f"{{{ns}}}split")
        split.set("orient", orient)
        split.set("dir", inout)

    elif transition_type.startswith("split_vertical_"):
        orient = "vert"
        inout = "in" if transition_type.endswith("_in") else "out"
        split = etree.SubElement(transition, f"{{{ns}}}split")
        split.set("orient", orient)
        split.set("dir", inout)

    elif transition_type == DISSOLVE:
        etree.SubElement(transition, f"{{{ns}}}dissolve")

    elif transition_type == RANDOM:
        etree.SubElement(transition, f"{{{ns}}}random")

    return transition


# ---------------------------------------------------------------------------
# 2. Apply transition to a single slide
# ---------------------------------------------------------------------------

def apply_slide_transition(
    slide: Any,
    transition_type: str = FADE,
    duration_ms: int = 700,
    advance_ms: int | None = None,
) -> None:
    """Apply a slide transition to a slide via OOXML manipulation.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide to apply the transition to.
    transition_type : str
        One of the transition type constants (default ``"fade"``).
    duration_ms : int
        Transition duration in milliseconds (default 700).
    advance_ms : int, optional
        Auto-advance time in milliseconds. ``None`` means manual advance.
    """
    if transition_type not in TRANSITION_TYPES:
        log.warning(
            "Unknown transition type %r; falling back to %r. "
            "Valid types: %s",
            transition_type, FADE, ", ".join(sorted(TRANSITION_TYPES)),
        )
        transition_type = FADE

    # Build the <p:transition> element
    transition_elem = _build_transition_xml(transition_type, duration_ms, advance_ms)

    # Get the slide's root XML element (<p:sld>)
    sld = slide._element

    # Remove any existing <p:transition> to avoid duplicates
    ns = _P_NS
    existing = sld.find(f"{{{ns}}}transition")
    if existing is not None:
        sld.remove(existing)

    # Insert <p:transition> before <p:timing> if present, otherwise
    # before the last child (typically </p:sld>).  The OOXML spec
    # expects <p:transition> to appear after <p:clrMapOvr> and before
    # <p:timing> / <p:extLst>.
    timing = sld.find(f"{{{ns}}}timing")
    if timing is not None:
        timing.addprevious(transition_elem)
    else:
        # Append as the last element before the closing tag
        sld.append(transition_elem)


# ---------------------------------------------------------------------------
# 3. Deck-level transition helper
# ---------------------------------------------------------------------------

def apply_deck_transitions(
    prs: Any,
    transition_type: str = FADE,
    duration_ms: int = 700,
) -> None:
    """Apply a consistent transition to all slides in a presentation.

    Parameters
    ----------
    prs : pptx.Presentation
        The presentation object.
    transition_type : str
        One of the transition type constants (default ``"fade"``).
    duration_ms : int
        Transition duration in milliseconds (default 700).
    """
    for slide in prs.slides:
        apply_slide_transition(slide, transition_type, duration_ms)


# ---------------------------------------------------------------------------
# 3b. Path-based convenience overloads (accept file path, 1-based slide_index)
# ---------------------------------------------------------------------------

def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def add_transition(
    prs_or_path,
    slide_index: int,
    transition_type: str = FADE,
    duration_ms: int = 700,
    advance_ms: int | None = None,
) -> None:
    """Apply a transition to a specific slide by 1-based index.

    Accepts either an open ``Presentation`` object or a file path string.
    When given a path, opens the file, applies the transition, and saves.

    Parameters
    ----------
    prs_or_path : Presentation or str
        An open Presentation object, or a path to a .pptx file.
    slide_index : int
        1-based slide index (1 = first slide).
    transition_type : str
        One of the transition type constants (default ``"fade"``).
    duration_ms : int
        Transition duration in milliseconds (default 700).
    advance_ms : int, optional
        Auto-advance time in milliseconds. ``None`` means manual advance.
    """
    from pptx import Presentation

    is_path = not _is_presentation(prs_or_path)
    prs = prs_or_path if not is_path else Presentation(str(prs_or_path))
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        apply_slide_transition(slide, transition_type, duration_ms, advance_ms)
    finally:
        if is_path:
            _save_prs_impl(prs, prs_or_path, backup=False)


def add_deck_transitions(
    prs_or_path,
    transition_type: str = FADE,
    duration_ms: int = 700,
) -> None:
    """Apply a consistent transition to all slides (path or Presentation).

    Accepts either an open ``Presentation`` object or a file path string.
    When given a path, opens the file, applies transitions, and saves.

    Parameters
    ----------
    prs_or_path : Presentation or str
        An open Presentation object, or a path to a .pptx file.
    transition_type : str
        One of the transition type constants (default ``"fade"``).
    duration_ms : int
        Transition duration in milliseconds (default 700).
    """
    from pptx import Presentation

    is_path = not _is_presentation(prs_or_path)
    prs = prs_or_path if not is_path else Presentation(str(prs_or_path))
    try:
        for slide in prs.slides:
            apply_slide_transition(slide, transition_type, duration_ms)
    finally:
        if is_path:
            _save_prs_impl(prs, prs_or_path, backup=False)
