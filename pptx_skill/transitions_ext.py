"""Advanced slide transitions (Office 2010+) via OOXML p14 namespace.

python-pptx has no native API for slide transitions, and the original
:mod:`pptx_skill.transitions` module covers the 18 standard transitions
that use the ``p`` namespace.  This module adds 12 *advanced* transitions
introduced in Office 2010 that use the ``p14:prstTrans`` element inside
the standard ``<p:transition>`` container.

The p14 namespace is:
    ``http://schemas.microsoft.com/office/powerpoint/2010/main``

Usage
-----
>>> from pptx_skill.transitions_ext import set_advanced_transition
>>> set_advanced_transition("deck.pptx", 2, transition="wheel", duration=0.7)

Convenience helpers for parameterised transitions:

>>> from pptx_skill.transitions_ext import set_wheel_transition
>>> set_wheel_transition("deck.pptx", 3, spokes=6, duration=0.5)

>>> from pptx_skill.transitions_ext import set_flip_transition
>>> set_flip_transition("deck.pptx", 4, direction="right")
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pptx_skill._io import open_prs as _open_prs
from pptx_skill._io import resolve_path as _resolve_path
from pptx_skill._io import save_prs as _save_prs_impl
from pptx_skill.constants import P14_NS
from pptx_skill.constants import P_NS as _P_NS

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 1. Constants
# ---------------------------------------------------------------------------

# The 12 advanced transition names (valid values for prstTrans @prst)
ADVANCED_TRANSITIONS: tuple[str, ...] = (
    "wheel",
    "ripple",
    "honeycomb",
    "vortex",
    "shred",
    "flip",
    "gallery",
    "pan",
    "glitter",
    "warp",
    "wind",
    "curtain",
)

_VALID_ADVANCED: frozenset[str] = frozenset(ADVANCED_TRANSITIONS)

# Direction mappings for flip / pan
_DIR_MAP: dict[str, str] = {
    "left": "l",
    "right": "r",
    "up": "u",
    "down": "d",
}

# Speed presets (same as transitions.py)
_SPEED_MAP: dict[int, str] = {
    1500: "slow",
    700: "med",
    300: "fast",
}


# ---------------------------------------------------------------------------
# 2. Data class
# ---------------------------------------------------------------------------

@dataclass
class AdvancedTransitionInfo:
    """Describes an advanced transition applied to a slide.

    Attributes
    ----------
    slide_index : int
        1-based slide index.
    transition : str
        One of :data:`ADVANCED_TRANSITIONS`.
    duration : float or None
        Duration in seconds, if specified.
    advance_on_click : bool
        Whether the slide advances on mouse click.
    advance_after_ms : int or None
        Auto-advance time in milliseconds, or ``None`` for manual.
    """
    slide_index: int
    transition: str
    duration: float | None = None
    advance_on_click: bool = True
    advance_after_ms: int | None = None


# ---------------------------------------------------------------------------
# 3. Internal helpers
# ---------------------------------------------------------------------------

def _ms_to_spd(duration_ms: int) -> str | None:
    """Convert milliseconds to an OOXML ``spd`` preset.

    Returns ``"slow"``, ``"med"``, ``"fast"``, or ``None`` when no preset
    is close enough (caller should use ``dur`` attribute instead).
    """
    closest = min(_SPEED_MAP.keys(), key=lambda s: abs(s - duration_ms))
    if abs(closest - duration_ms) <= 200:
        return _SPEED_MAP[closest]
    return None


def _save_prs(prs: Any, prs_or_path: str | Path | Any) -> None:
    """Save *prs* back.  Only writes to disk when *prs_or_path* is a path."""
    path = _resolve_path(prs_or_path)
    if path is not None:
        _save_prs_impl(prs, path, backup=False)


def _build_advanced_transition_xml(
    transition: str,
    duration_ms: int | None = None,
    advance_on_click: bool = True,
    advance_after_ms: int | None = None,
    *,
    spokes: int | None = None,
    center: bool | None = None,
    direction: str | None = None,
) -> Any:
    """Build a ``<p:transition>`` element containing ``<p14:prstTrans>``.

    Parameters
    ----------
    transition : str
        One of :data:`ADVANCED_TRANSITIONS`.
    duration_ms : int or None
        Duration in milliseconds.  ``None`` uses the default (med).
    advance_on_click : bool
        Whether clicking advances the slide.
    advance_after_ms : int or None
        Auto-advance time in milliseconds.
    spokes : int or None
        For *wheel*: number of spokes (1-8).
    center : bool or None
        For *ripple*: ``True`` = ripple from center.
    direction : str or None
        For *flip* / *pan*: ``"left"``, ``"right"``, ``"up"``, ``"down"``.

    Returns
    -------
    lxml.etree._Element
        The constructed ``<p:transition>`` element.
    """
    from lxml import etree  # lazy import

    ns_p = _P_NS
    ns_p14 = P14_NS

    transition_elem = etree.Element(f"{{{ns_p}}}transition")

    # --- advClick ---
    if not advance_on_click:
        transition_elem.set("advClick", "0")

    # --- Duration / speed ---
    if duration_ms is not None:
        spd = _ms_to_spd(duration_ms)
        if spd is not None:
            transition_elem.set("spd", spd)
        else:
            transition_elem.set("dur", str(duration_ms))

    # --- Auto-advance ---
    if advance_after_ms is not None:
        transition_elem.set("advTm", str(advance_after_ms))

    # --- p14:prstTrans ---
    prst_trans = etree.SubElement(transition_elem, f"{{{ns_p14}}}prstTrans")
    prst_trans.set("prst", transition)

    # --- Optional child elements ---

    # Wheel: <p14:spokes spokes="4"/>
    if spokes is not None:
        if transition != "wheel":
            log.warning("spokes parameter only applies to 'wheel' transition; ignoring")
        else:
            spokes_clamped = max(1, min(8, spokes))
            spokes_elem = etree.SubElement(prst_trans, f"{{{ns_p14}}}spokes")
            spokes_elem.set("spokes", str(spokes_clamped))

    # Ripple: center attribute on prstTrans
    if center is not None:
        if transition != "ripple":
            log.warning("center parameter only applies to 'ripple' transition; ignoring")
        else:
            # When center=False, PowerPoint uses a corner origin; we signal
            # this by omitting the attribute (default is center).
            if center:
                prst_trans.set("dir", "ctr")

    # Flip / Pan: <p14:dir dir="l"/>
    if direction is not None:
        if transition not in ("flip", "pan"):
            log.warning("direction parameter only applies to 'flip'/'pan' transitions; ignoring")
        else:
            dir_code = _DIR_MAP.get(direction)
            if dir_code is None:
                log.warning(
                    "Invalid direction %r for %s transition; valid: %s. "
                    "Falling back to 'left'.",
                    direction, transition, ", ".join(_DIR_MAP),
                )
                dir_code = "l"
            dir_elem = etree.SubElement(prst_trans, f"{{{ns_p14}}}dir")
            dir_elem.set("dir", dir_code)

    return transition_elem


def _insert_transition_on_slide(slide: Any, transition_elem: Any) -> None:
    """Insert a ``<p:transition>`` element into a slide, replacing any existing one."""

    sld = slide._element
    ns_p = _P_NS

    # Remove any existing <p:transition>
    existing = sld.find(f"{{{ns_p}}}transition")
    if existing is not None:
        sld.remove(existing)

    # Insert before <p:timing> if present, otherwise append
    timing = sld.find(f"{{{ns_p}}}timing")
    if timing is not None:
        timing.addprevious(transition_elem)
    else:
        sld.append(transition_elem)


def _find_prst_trans(transition_elem: Any) -> Any | None:
    """Find the ``<p14:prstTrans>`` child of a ``<p:transition>`` element."""
    return transition_elem.find(f"{{{P14_NS}}}prstTrans")


def _read_transition_info(slide: Any, slide_index: int) -> AdvancedTransitionInfo | None:
    """Read advanced transition info from a slide, or ``None`` if not present."""
    sld = slide._element
    trans = sld.find(f"{{{_P_NS}}}transition")
    if trans is None:
        return None

    prst = _find_prst_trans(trans)
    if prst is None:
        return None

    prst_name = prst.get("prst", "")
    if prst_name not in _VALID_ADVANCED:
        return None

    # Duration
    duration: float | None = None
    dur_attr = trans.get("dur")
    spd_attr = trans.get("spd")
    if dur_attr is not None:
        try:
            duration = int(dur_attr) / 1000.0
        except (ValueError, TypeError):
            pass
    elif spd_attr is not None:
        _spd_to_ms: dict[str, int] = {"slow": 1500, "med": 700, "fast": 300}
        ms = _spd_to_ms.get(spd_attr)
        if ms is not None:
            duration = ms / 1000.0

    # advance_on_click
    adv_click_attr = trans.get("advClick")
    advance_on_click = adv_click_attr != "0" if adv_click_attr is not None else True

    # advance_after_ms
    advance_after_ms: int | None = None
    adv_tm_attr = trans.get("advTm")
    if adv_tm_attr is not None:
        try:
            advance_after_ms = int(adv_tm_attr)
        except (ValueError, TypeError):
            pass

    return AdvancedTransitionInfo(
        slide_index=slide_index,
        transition=prst_name,
        duration=duration,
        advance_on_click=advance_on_click,
        advance_after_ms=advance_after_ms,
    )


# ---------------------------------------------------------------------------
# 4. Public API
# ---------------------------------------------------------------------------

def set_advanced_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
    *,
    transition: str,
    duration: float | None = None,
    advance_on_click: bool = True,
    advance_after_ms: int | None = None,
) -> bool:
    """Apply an advanced (p14) transition to a slide.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.
    transition : str
        One of :data:`ADVANCED_TRANSITIONS` (e.g. ``"wheel"``, ``"ripple"``).
    duration : float or None
        Transition duration in seconds.  ``None`` uses the PowerPoint default.
    advance_on_click : bool
        Whether clicking advances the slide (default ``True``).
    advance_after_ms : int or None
        Auto-advance time in milliseconds, or ``None`` for manual.

    Returns
    -------
    bool
        ``True`` if the transition was applied successfully.
    """
    if transition not in _VALID_ADVANCED:
        log.error(
            "Invalid advanced transition %r. Valid: %s",
            transition, ", ".join(sorted(_VALID_ADVANCED)),
        )
        return False

    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        duration_ms = int(duration * 1000) if duration is not None else None
        elem = _build_advanced_transition_xml(
            transition,
            duration_ms=duration_ms,
            advance_on_click=advance_on_click,
            advance_after_ms=advance_after_ms,
        )
        _insert_transition_on_slide(slide, elem)
    except Exception:
        log.exception("Failed to set advanced transition on slide %d", slide_index)
        return False

    _save_prs(prs, prs_or_path)
    return True


def set_wheel_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
    *,
    spokes: int = 4,
    duration: float | None = None,
) -> bool:
    """Apply a wheel transition with configurable spoke count.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.
    spokes : int
        Number of spokes (1-8, default 4).
    duration : float or None
        Transition duration in seconds.

    Returns
    -------
    bool
        ``True`` if the transition was applied successfully.
    """
    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        duration_ms = int(duration * 1000) if duration is not None else None
        elem = _build_advanced_transition_xml(
            "wheel",
            duration_ms=duration_ms,
            spokes=spokes,
        )
        _insert_transition_on_slide(slide, elem)
    except Exception:
        log.exception("Failed to set wheel transition on slide %d wheel transition", slide_index)
        return False

    _save_prs(prs, prs_or_path)
    return True


def set_ripple_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
    *,
    center: bool = True,
    duration: float | None = None,
) -> bool:
    """Apply a ripple transition.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.
    center : bool
        ``True`` for ripple from center (default), ``False`` for corner.
    duration : float or None
        Transition duration in seconds.

    Returns
    -------
    bool
        ``True`` if the transition was applied successfully.
    """
    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        duration_ms = int(duration * 1000) if duration is not None else None
        elem = _build_advanced_transition_xml(
            "ripple",
            duration_ms=duration_ms,
            center=center,
        )
        _insert_transition_on_slide(slide, elem)
    except Exception:
        log.exception("Failed to set slide %d ripple transition", slide_index)
        return False

    _save_prs(prs, prs_or_path)
    return True


def set_flip_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
    *,
    direction: str = "left",
    duration: float | None = None,
) -> bool:
    """Apply a 3D flip transition.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.
    direction : str
        Flip direction: ``"left"``, ``"right"``, ``"up"``, ``"down"``
        (default ``"left"``).
    duration : float or None
        Transition duration in seconds.

    Returns
    -------
    bool
        ``True`` if the transition was applied successfully.
    """
    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        duration_ms = int(duration * 1000) if duration is not None else None
        elem = _build_advanced_transition_xml(
            "flip",
            duration_ms=duration_ms,
            direction=direction,
        )
        _insert_transition_on_slide(slide, elem)
    except Exception:
        log.exception("Failed to set slide %d flip transition", slide_index)
        return False

    _save_prs(prs, prs_or_path)
    return True


def set_pan_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
    *,
    direction: str = "left",
    duration: float | None = None,
) -> bool:
    """Apply a smooth pan transition.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.
    direction : str
        Pan direction: ``"left"``, ``"right"``, ``"up"``, ``"down"``
        (default ``"left"``).
    duration : float or None
        Transition duration in seconds.

    Returns
    -------
    bool
        ``True`` if the transition was applied successfully.
    """
    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        duration_ms = int(duration * 1000) if duration is not None else None
        elem = _build_advanced_transition_xml(
            "pan",
            duration_ms=duration_ms,
            direction=direction,
        )
        _insert_transition_on_slide(slide, elem)
    except Exception:
        log.exception("Failed to set slide %d pan transition", slide_index)
        return False

    _save_prs(prs, prs_or_path)
    return True


def remove_advanced_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
) -> bool:
    """Remove an advanced transition from a slide.

    This removes the entire ``<p:transition>`` element if it contains a
    ``<p14:prstTrans>`` child.  If the transition element contains other
    (non-p14) transition types, it is left intact.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.

    Returns
    -------
    bool
        ``True`` if an advanced transition was found and removed.
    """
    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        sld = slide._element
        trans = sld.find(f"{{{_P_NS}}}transition")
        if trans is None:
            return False

        prst = _find_prst_trans(trans)
        if prst is None:
            # Not an advanced transition — do not remove
            return False

        sld.remove(trans)
    except Exception:
        log.exception("Failed to remove advanced transition on slide %d", slide_index)
        return False

    _save_prs(prs, prs_or_path)
    return True


def list_advanced_transitions(
    prs_or_path: str | Path | Any,
) -> list[dict]:
    """List all advanced transitions in a presentation.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.

    Returns
    -------
    list[dict]
        A list of dicts, each with keys matching
        :class:`AdvancedTransitionInfo` fields, for every slide that has
        an advanced (p14) transition.
    """
    prs = _open_prs(prs_or_path)
    results: list[dict] = []

    try:
        for idx, slide in enumerate(prs.slides, start=1):
            info = _read_transition_info(slide, idx)
            if info is not None:
                results.append({
                    "slide_index": info.slide_index,
                    "transition": info.transition,
                    "duration": info.duration,
                    "advance_on_click": info.advance_on_click,
                    "advance_after_ms": info.advance_after_ms,
                })
    except Exception:
        log.exception("Failed to list advanced transitions")

    return results


def has_advanced_transition(
    prs_or_path: str | Path | Any,
    slide_index: int,
) -> bool | str:
    """Check whether a slide has an advanced (p14) transition.

    Parameters
    ----------
    prs_or_path : str, Path, or Presentation
        File path or an open ``pptx.Presentation`` object.
    slide_index : int
        1-based slide index.

    Returns
    -------
    bool or str
        ``False`` if no advanced transition is present, or the transition
        name string (e.g. ``"wheel"``) if one is found.
    """
    if slide_index < 1:
        log.error("slide_index must be >= 1, got %d", slide_index)
        return False

    prs = _open_prs(prs_or_path)
    try:
        slides = prs.slides
        if slide_index > len(slides):
            log.error(
                "slide_index %d exceeds slide count %d",
                slide_index, len(slides),
            )
            return False

        slide = slides[slide_index - 1]
        info = _read_transition_info(slide, slide_index)
        if info is not None:
            return info.transition
        return False
    except Exception:
        log.exception("Failed to check advanced transition on slide %d", slide_index)
        return False


# ---------------------------------------------------------------------------
# 5. Public names
# ---------------------------------------------------------------------------

__all__: list[str] = [
    # Constants
    "ADVANCED_TRANSITIONS",
    "P14_NS",
    # Data class
    "AdvancedTransitionInfo",
    # Public API
    "set_advanced_transition",
    "set_wheel_transition",
    "set_ripple_transition",
    "set_flip_transition",
    "set_pan_transition",
    "remove_advanced_transition",
    "list_advanced_transitions",
    "has_advanced_transition",
]
