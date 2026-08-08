"""PowerPoint Morph transition support via OOXML manipulation.

The Morph transition (introduced in PowerPoint 2016) creates smooth object
animations between slides.  It uses the ``p15`` namespace
(``http://schemas.microsoft.com/office/powerpoint/2012/main``) and is
expressed as a ``<p15:prstTrans prst="morph">`` element inside the
standard ``<p:transition>`` container.

Usage
-----
>>> from pptx_skill.morph import set_morph_transition, set_morph_options
>>> set_morph_transition("deck.pptx", 2, option="full")
>>> set_morph_options("deck.pptx", 2, morph_by="char")

Or with an already-opened Presentation object:

>>> from pptx import Presentation
>>> prs = Presentation("deck.pptx")
>>> set_morph_transition(prs, 1, option="full")
>>> # No save needed — caller controls the Presentation lifecycle
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

from pptx_skill._io import is_presentation as _is_presentation
from pptx_skill._io import open_prs as _open_prs
from pptx_skill._io import save_prs as _save_prs
from pptx_skill.constants import P15_NS as _P15_NS
from pptx_skill.constants import P_NS as _P_NS

log = logging.getLogger(__name__)

# Valid option values for set_morph_transition
_MORPH_OPTIONS: frozenset[str] = frozenset({"full", "none"})

# Valid morph_by values for set_morph_options
_MORPH_BY_VALUES: frozenset[str] = frozenset({"object", "word", "char"})

# Mapping from morph_by value to the p15:option @type attribute
_MORPH_BY_TYPE_MAP: dict[str, str] = {
    "object": "morphByObject",
    "word": "morphByWord",
    "char": "morphByChar",
}

# ---------------------------------------------------------------------------
# Public data class
# ---------------------------------------------------------------------------


@dataclass
class MorphInfo:
    """Information about a morph transition on a slide.

    Attributes
    ----------
    slide_index : int
        1-based slide index.
    option : str
        Morph option: ``"full"`` (morph all objects), ``"none"``
        (background only), or a specific shape name.
    morph_by : str
        Text morph mode: ``"object"``, ``"word"``, or ``"char"``.
    """

    slide_index: int
    option: str = "full"
    morph_by: str = "object"


# ---------------------------------------------------------------------------
# Internal helpers — Presentation open / save
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Internal helpers — XML construction
# ---------------------------------------------------------------------------


def _build_morph_transition_xml(
    option: str = "full",
    morph_by: str = "object",
) -> Any:
    """Build a ``<p:transition>`` element containing a morph transition.

    Parameters
    ----------
    option : str
        ``"full"`` — morph all objects (default).
        ``"none"`` — background-only morph (no effect options).
        Any other string — treated as a specific shape name; adds
        ``<p15:option type="morphByObject"/>`` to limit morphing to
        named objects.
    morph_by : str
        Controls how text morphs: ``"object"`` (default), ``"word"``,
        or ``"char"``.

    Returns
    -------
    lxml.etree._Element
        The constructed ``<p:transition>`` element.
    """
    from lxml import etree

    # Build <p:transition> with the p15 namespace in its nsmap so that
    # child elements can use it without per-element xmlns declarations.
    transition = etree.Element(
        f"{{{_P_NS}}}transition",
        nsmap={"p": _P_NS, "p15": _P15_NS},
    )
    transition.set("spd", "med")
    transition.set("advClick", "0")

    # <p15:prstTrans prst="morph">
    prst_trans = etree.SubElement(
        transition,
        f"{{{_P15_NS}}}prstTrans",
    )
    prst_trans.set("prst", "morph")

    # Add option elements based on parameters.
    # - option="full": always emit <p15:option type="morphByObject"/> (or
    #   the morph_by variant) so the XML is self-documenting and
    #   distinguishable from option="none".
    # - option="none": no <p15:option> at all (background-only morph).
    # - option=specific_name: add morphByObject option (same as "full"
    #   in OOXML; the caller is expected to set matching shape names).
    if option != "none":
        opt = etree.SubElement(prst_trans, f"{{{_P15_NS}}}option")
        opt.set("type", _MORPH_BY_TYPE_MAP.get(morph_by, "morphByObject"))

    return transition


def _find_transition_elem(sld: Any) -> Any | None:
    """Find the existing ``<p:transition>`` element on a slide, or None."""
    return sld.find(f"{{{_P_NS}}}transition")


def _find_prst_trans(transition: Any) -> Any | None:
    """Find the ``<p15:prstTrans>`` child of a transition element, or None."""
    return transition.find(f"{{{_P15_NS}}}prstTrans")


def _is_morph_prst_trans(prst_trans: Any) -> bool:
    """Return True if the prstTrans element is a morph transition."""
    return prst_trans is not None and prst_trans.get("prst") == "morph"


def _parse_morph_info(slide_index: int, transition: Any) -> MorphInfo | None:
    """Extract MorphInfo from a transition element, or None if not morph."""
    prst_trans = _find_prst_trans(transition)
    if not _is_morph_prst_trans(prst_trans):
        return None
    assert prst_trans is not None

    # Determine option
    options = prst_trans.findall(f"{{{_P15_NS}}}option")
    if not options:
        option = "none"
        morph_by = "object"
    else:
        # If there are option elements, check their types
        has_object_option = False
        morph_by = "object"
        for opt in options:
            opt_type = opt.get("type", "")
            if opt_type == "morphByObject":
                has_object_option = True
            elif opt_type == "morphByWord":
                has_object_option = True
                morph_by = "word"
            elif opt_type == "morphByChar":
                has_object_option = True
                morph_by = "char"

        if has_object_option:
            option = "full"
        else:
            option = "none"

    return MorphInfo(slide_index=slide_index, option=option, morph_by=morph_by)


def _insert_transition(sld: Any, transition_elem: Any) -> None:
    """Insert a transition element into the slide XML at the correct position.

    The OOXML spec expects ``<p:transition>`` to appear after
    ``<p:clrMapOvr>`` and before ``<p:timing>`` / ``<p:extLst>``.
    """
    # Remove any existing <p:transition> to avoid duplicates
    existing = _find_transition_elem(sld)
    if existing is not None:
        sld.remove(existing)

    # Insert before <p:timing> if present, otherwise append
    timing = sld.find(f"{{{_P_NS}}}timing")
    if timing is not None:
        timing.addprevious(transition_elem)
    else:
        sld.append(transition_elem)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def set_morph_transition(
    prs_or_path: Any,
    slide_index: int,
    *,
    option: str = "full",
) -> bool:
    """Apply a Morph transition to a slide.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.  When a path is
        provided the file is saved automatically (with a ``.bak.pptx``
        backup).
    slide_index : int
        1-based slide index (1 = first slide).
    option : str
        ``"full"`` (default) — morph all objects on the slide.
        ``"none"`` — background-only morph (no effect options emitted).
        Any other string — treated as a specific shape name; emits
        ``<p15:option type="morphByObject"/>`` to enable per-object
        morphing.

    Returns
    -------
    bool
        ``True`` if the transition was applied successfully.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)

    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(
            f"slide_index {slide_index} out of range "
            f"(1..{len(prs.slides)})"
        )

    slide = prs.slides[slide_index - 1]
    sld = slide._element

    # Determine morph_by from any existing morph transition (preserve it)
    existing = _find_transition_elem(sld)
    morph_by = "object"
    if existing is not None:
        info = _parse_morph_info(slide_index, existing)
        if info is not None:
            morph_by = info.morph_by

    transition_elem = _build_morph_transition_xml(option=option, morph_by=morph_by)
    _insert_transition(sld, transition_elem)

    if is_path and path is not None:
        _save_prs(prs, path)

    log.info("Morph transition set on slide %d (option=%r)", slide_index, option)
    return True


def set_morph_options(
    prs_or_path: Any,
    slide_index: int,
    *,
    morph_by: str = "object",
) -> bool:
    """Set the text-morph mode for a slide that already has a Morph transition.

    If the slide does not currently have a Morph transition, one is added
    with ``option="full"``.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        0-based index of the target slide.
    morph_by : str
        ``"object"`` (default) — morph by object.
        ``"word"`` — morph by word (text breaks into word-level animations).
        ``"char"`` — morph by character (text breaks into character-level
        animations).

    Returns
    -------
    bool
        ``True`` if the option was set successfully.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    ValueError
        If *morph_by* is not one of ``"object"``, ``"word"``, ``"char"``.
    """
    if morph_by not in _MORPH_BY_VALUES:
        raise ValueError(
            f"morph_by must be one of {sorted(_MORPH_BY_VALUES)}, "
            f"got {morph_by!r}"
        )

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)

    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(
            f"slide_index {slide_index} out of range "
            f"(0–{len(prs.slides) - 1})"
        )

    slide = prs.slides[slide_index]
    sld = slide._element

    existing = _find_transition_elem(sld)
    if existing is not None:
        prst_trans = _find_prst_trans(existing)
        if _is_morph_prst_trans(prst_trans):
            assert prst_trans is not None
            # Remove existing <p15:option> elements
            for opt in prst_trans.findall(f"{{{_P15_NS}}}option"):
                prst_trans.remove(opt)

            # Add the new option (morph_by="object" with option="full" is
            # the default — no element needed; but for "word"/"char" we
            # must emit one).
            if morph_by in ("word", "char"):
                from lxml import etree

                opt_elem = etree.SubElement(
                    prst_trans, f"{{{_P15_NS}}}option"
                )
                opt_elem.set("type", _MORPH_BY_TYPE_MAP[morph_by])
            elif morph_by == "object":
                # morphByObject is the default; still emit it for
                # explicitness so the XML is self-documenting.
                from lxml import etree

                opt_elem = etree.SubElement(
                    prst_trans, f"{{{_P15_NS}}}option"
                )
                opt_elem.set("type", "morphByObject")

            if is_path and path is not None:
                _save_prs(prs, path)

            log.info(
                "Morph option set on slide %d (morph_by=%r)",
                slide_index,
                morph_by,
            )
            return True

    # No existing morph transition — create one with the requested option
    transition_elem = _build_morph_transition_xml(option="full", morph_by=morph_by)
    _insert_transition(sld, transition_elem)

    if is_path and path is not None:
        _save_prs(prs, path)

    log.info(
        "Morph transition created on slide %d (morph_by=%r)",
        slide_index,
        morph_by,
    )
    return True


def remove_morph_transition(
    prs_or_path: Any,
    slide_index: int,
) -> bool:
    """Remove the Morph transition from a slide.

    If the slide has a transition element that contains a morph
    ``<p15:prstTrans>``, the entire ``<p:transition>`` element is
    removed.  If the transition is not a morph transition, nothing is
    changed.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        0-based index of the target slide.

    Returns
    -------
    bool
        ``True`` if a morph transition was found and removed,
        ``False`` otherwise.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)

    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(
            f"slide_index {slide_index} out of range "
            f"(0–{len(prs.slides) - 1})"
        )

    slide = prs.slides[slide_index]
    sld = slide._element

    existing = _find_transition_elem(sld)
    if existing is None:
        return False

    prst_trans = _find_prst_trans(existing)
    if not _is_morph_prst_trans(prst_trans):
        return False

    sld.remove(existing)

    if is_path and path is not None:
        _save_prs(prs, path)

    log.info("Morph transition removed from slide %d", slide_index)
    return True


def has_morph_transition(
    prs_or_path: Any,
    slide_index: int,
) -> bool:
    """Check whether a slide has a Morph transition.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.
    slide_index : int
        0-based index of the target slide.

    Returns
    -------
    bool
        ``True`` if the slide has a morph transition.

    Raises
    ------
    IndexError
        If *slide_index* is out of range.
    """
    prs = _open_prs(prs_or_path)

    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(
            f"slide_index {slide_index} out of range "
            f"(0–{len(prs.slides) - 1})"
        )

    slide = prs.slides[slide_index]
    sld = slide._element

    existing = _find_transition_elem(sld)
    if existing is None:
        return False

    prst_trans = _find_prst_trans(existing)
    return _is_morph_prst_trans(prst_trans)


def list_morph_transitions(prs_or_path: Any) -> list[dict]:
    """List all slides that have a Morph transition.

    Parameters
    ----------
    prs_or_path : Presentation | str | Path
        An open ``Presentation`` object or a file path.

    Returns
    -------
    list[dict]
        A list of dicts, each with keys ``slide_index``, ``option``,
        and ``morph_by``.  Only slides with a morph transition are
        included.
    """
    prs = _open_prs(prs_or_path)
    results: list[dict] = []

    for idx, slide in enumerate(prs.slides):
        sld = slide._element
        existing = _find_transition_elem(sld)
        if existing is None:
            continue

        info = _parse_morph_info(idx, existing)
        if info is not None:
            results.append(
                {
                    "slide_index": info.slide_index,
                    "option": info.option,
                    "morph_by": info.morph_by,
                }
            )

    return results


__all__ = [
    "MorphInfo",
    "has_morph_transition",
    "list_morph_transitions",
    "remove_morph_transition",
    "set_morph_options",
    "set_morph_transition",
]
