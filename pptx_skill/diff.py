"""Structural and content comparison/diff for PPTX presentations.

Compares two PPTX files at the slide, shape, text, image, and layout levels
and produces a structured diff report.  Designed for version-to-version change
detection and regression testing.

This module is intentionally standalone -- it only requires ``python-pptx`` at
call time (lazy import inside functions), keeping the import graph clean.
"""
from __future__ import annotations

import hashlib
import json
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from pptx_skill._compat import StrEnum

__all__ = [
    "Change",
    "ChangeKind",
    "PresentationDiff",
    "SlideDiff",
    "diff_presentations",
    "diff_presentations_visual",
    "diff_slides",
    "diff_text",
    "format_diff",
    "has_changes",
]

# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------

class ChangeKind(StrEnum):
    """Kinds of changes detectable between two shapes or slides."""

    TEXT_CHANGED = "text_changed"
    IMAGE_CHANGED = "image_changed"
    SHAPE_ADDED = "shape_added"
    SHAPE_REMOVED = "shape_removed"
    POSITION_CHANGED = "position_changed"
    SIZE_CHANGED = "size_changed"
    COLOR_CHANGED = "color_changed"
    FONT_CHANGED = "font_changed"
    LAYOUT_CHANGED = "layout_changed"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Change:
    """A single atomic difference between two presentation elements.

    Attributes:
        kind: What category of change occurred.
        old_value: The value in the older presentation (``None`` for additions).
        new_value: The value in the newer presentation (``None`` for removals).
        shape_name: Name of the shape where the change was detected, if any.
        details: Optional human-readable detail string.
    """

    kind: ChangeKind
    old_value: Any = None
    new_value: Any = None
    shape_name: str | None = None
    details: str | None = None


@dataclass
class SlideDiff:
    """Differences for a single slide pair.

    Attributes:
        slide_index: 0-based index of the slide in both presentations.
        changes: Ordered list of ``Change`` objects.
    """

    slide_index: int
    changes: list[Change] = field(default_factory=list)


@dataclass
class PresentationDiff:
    """Top-level diff result for two presentations.

    Attributes:
        slide_count_old: Number of slides in the older presentation.
        slide_count_new: Number of slides in the newer presentation.
        slides_added: 0-based indices of slides present only in the new file.
        slides_removed: 0-based indices of slides present only in the old file.
        slide_diffs: Per-slide diffs for slides that exist in both files.
        summary: Human-readable one-line summary.
    """

    slide_count_old: int = 0
    slide_count_new: int = 0
    slides_added: list[int] = field(default_factory=list)
    slides_removed: list[int] = field(default_factory=list)
    slide_diffs: list[SlideDiff] = field(default_factory=list)
    summary: str = ""


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

# Distance tolerance (points) for nearest-neighbour shape matching.
_POSITION_TOLERANCE_PT = 10.0


def _md5_of_blob(blob: bytes) -> str:
    """Return the hex MD5 digest of *blob*."""
    return hashlib.md5(blob).hexdigest()


def _extract_shape_info(shape: Any) -> dict[str, Any]:
    """Pull comparable information out of a python-pptx shape object.

    Returns a plain dict that is easy to diff.  Keys are deliberately flat
    so that simple equality checks work for most attributes.
    """
    info: dict[str, Any] = {
        "name": shape.name,
        "shape_type": str(shape.shape_type),
        "left": shape.left,
        "top": shape.top,
        "width": shape.width,
        "height": shape.height,
    }

    # -- Position in EMU -> pt for readability (914400 EMU = 72 pt) --
    info["left_pt"] = shape.left / 12700.0
    info["top_pt"] = shape.top / 12700.0
    info["width_pt"] = shape.width / 12700.0
    info["height_pt"] = shape.height / 12700.0

    # -- Text content --
    if shape.has_text_frame:
        paragraphs: list[dict[str, Any]] = []
        for para in shape.text_frame.paragraphs:
            runs: list[dict[str, Any]] = []
            for run in para.runs:
                run_info: dict[str, Any] = {"text": run.text}
                if run.font.size is not None:
                    run_info["font_size"] = run.font.size
                if run.font.name is not None:
                    run_info["font_name"] = run.font.name
                if run.font.bold is not None:
                    run_info["bold"] = run.font.bold
                if run.font.italic is not None:
                    run_info["italic"] = run.font.italic
                try:
                    if run.font.color is not None and run.font.color.type is not None:
                        run_info["font_color"] = str(run.font.color.rgb)
                except (AttributeError, TypeError):
                    # font.color.rgb raises AttributeError when the color
                    # type is _NoneColor (inherited theme / no explicit RGB).
                    pass
                runs.append(run_info)
            paragraphs.append({"text": para.text, "runs": runs})
        info["paragraphs"] = paragraphs
        info["full_text"] = shape.text_frame.text

    # -- Image content --
    if shape.shape_type is not None and hasattr(shape, "image"):
        try:
            img = shape.image
            info["image_hash"] = _md5_of_blob(img.blob)
            info["image_content_type"] = getattr(img, "content_type", "")
        except Exception:
            # Some shapes have an .image attribute but it is not always
            # readable (e.g. placeholders without an image fill).
            pass

    # -- Fill / colour --
    try:
        fill = shape.fill
        if fill.type is not None:
            info["fill_type"] = str(fill.type)
            if fill.type is not None and hasattr(fill, "fore_color"):
                try:
                    info["fill_color"] = str(fill.fore_color.rgb)
                except Exception:
                    pass
    except Exception:
        # Some shapes do not support fill access.
        pass

    return info


def _center_pt(info: dict[str, Any]) -> tuple[float, float]:
    """Return (cx, cy) of a shape info dict in points."""
    return (
        info["left_pt"] + info["width_pt"] / 2.0,
        info["top_pt"] + info["height_pt"] / 2.0,
    )


def _distance_pt(a: dict[str, Any], b: dict[str, Any]) -> float:
    """Euclidean distance between shape centres in points."""
    (ax, ay) = _center_pt(a)
    (bx, by) = _center_pt(b)
    return math.hypot(ax - bx, ay - by)


def _match_shapes(
    infos_a: list[dict[str, Any]],
    infos_b: list[dict[str, Any]],
) -> list[tuple[int | None, int | None]]:
    """Match shapes between two slides.

    Returns a list of pairs ``(index_a, index_b)``.  An index of ``None``
    means the shape is unmatched on that side.
    """
    matched_a: set[int] = set()
    matched_b: set[int] = set()
    pairs: list[tuple[int | None, int | None]] = []

    # Phase 1: match by name
    name_map_b: dict[str, int] = {}
    for i, info_b in enumerate(infos_b):
        nm = info_b.get("name", "")
        if nm:
            name_map_b.setdefault(nm, i)

    for ia, info_a in enumerate(infos_a):
        nm = info_a.get("name", "")
        if nm and nm in name_map_b:
            ib = name_map_b[nm]
            if ib not in matched_b:
                pairs.append((ia, ib))
                matched_a.add(ia)
                matched_b.add(ib)

    # Phase 2: match by nearest-neighbour position (within tolerance)
    unmatched_a = [i for i in range(len(infos_a)) if i not in matched_a]
    unmatched_b = [i for i in range(len(infos_b)) if i not in matched_b]

    for ia in list(unmatched_a):
        best_ib: int | None = None
        best_dist = float("inf")
        for ib in unmatched_b:
            d = _distance_pt(infos_a[ia], infos_b[ib])
            if d < best_dist:
                best_dist = d
                best_ib = ib
        if best_ib is not None and best_dist <= _POSITION_TOLERANCE_PT:
            pairs.append((ia, best_ib))
            matched_a.add(ia)
            matched_b.add(best_ib)
            unmatched_b.remove(best_ib)
            unmatched_a.remove(ia)

    # Phase 3: match by type + approximate position (looser tolerance)
    for ia in list(unmatched_a):
        best_ib = None
        best_dist = float("inf")
        for ib in unmatched_b:
            if infos_a[ia]["shape_type"] == infos_b[ib]["shape_type"]:
                d = _distance_pt(infos_a[ia], infos_b[ib])
                if d < best_dist:
                    best_dist = d
                    best_ib = ib
        if best_ib is not None and best_dist <= _POSITION_TOLERANCE_PT * 5:
            pairs.append((ia, best_ib))
            matched_a.add(ia)
            matched_b.add(best_ib)
            unmatched_b.remove(best_ib)
            unmatched_a.remove(ia)

    # Remaining unmatched in A -> removed
    for ia in unmatched_a:
        pairs.append((ia, None))

    # Remaining unmatched in B -> added
    for ib in unmatched_b:
        pairs.append((None, ib))

    return pairs


def _normalize_text(
    text: str,
    *,
    ignore_whitespace: bool = True,
    ignore_case: bool = False,
) -> str:
    """Normalise text for comparison."""
    if ignore_whitespace:
        text = re.sub(r"\s+", " ", text).strip()
    if ignore_case:
        text = text.lower()
    return text


def _compare_text(
    text_a: str,
    text_b: str,
    *,
    ignore_whitespace: bool = True,
    ignore_case: bool = False,
) -> bool:
    """Return True if two text strings are equal after normalisation."""
    return _normalize_text(text_a, ignore_whitespace=ignore_whitespace, ignore_case=ignore_case) == _normalize_text(
        text_b, ignore_whitespace=ignore_whitespace, ignore_case=ignore_case
    )


def _compare_shape_details(
    info_a: dict[str, Any],
    info_b: dict[str, Any],
    *,
    detail_level: str = "normal",
    ignore_whitespace: bool = True,
    ignore_case: bool = False,
) -> list[Change]:
    """Compare two shape info dicts and return a list of ``Change`` objects."""
    changes: list[Change] = []
    shape_name = info_a.get("name") or info_b.get("name")

    # -- Position --
    if abs(info_a["left_pt"] - info_b["left_pt"]) > 0.5 or abs(info_a["top_pt"] - info_b["top_pt"]) > 0.5:
        changes.append(
            Change(
                kind=ChangeKind.POSITION_CHANGED,
                old_value=f"({info_a['left_pt']:.1f}, {info_a['top_pt']:.1f})",
                new_value=f"({info_b['left_pt']:.1f}, {info_b['top_pt']:.1f})",
                shape_name=shape_name,
            )
        )

    # -- Size --
    if abs(info_a["width_pt"] - info_b["width_pt"]) > 0.5 or abs(info_a["height_pt"] - info_b["height_pt"]) > 0.5:
        changes.append(
            Change(
                kind=ChangeKind.SIZE_CHANGED,
                old_value=f"{info_a['width_pt']:.1f}x{info_a['height_pt']:.1f}",
                new_value=f"{info_b['width_pt']:.1f}x{info_b['height_pt']:.1f}",
                shape_name=shape_name,
            )
        )

    # -- Fill / colour --
    fill_a = info_a.get("fill_color")
    fill_b = info_b.get("fill_color")
    if fill_a != fill_b:
        changes.append(
            Change(
                kind=ChangeKind.COLOR_CHANGED,
                old_value=fill_a,
                new_value=fill_b,
                shape_name=shape_name,
                details="fill color",
            )
        )

    # -- Image hash --
    hash_a = info_a.get("image_hash")
    hash_b = info_b.get("image_hash")
    if hash_a is not None and hash_b is not None and hash_a != hash_b:
        changes.append(
            Change(
                kind=ChangeKind.IMAGE_CHANGED,
                old_value=hash_a,
                new_value=hash_b,
                shape_name=shape_name,
            )
        )

    # -- Text --
    text_a = info_a.get("full_text", "")
    text_b = info_b.get("full_text", "")
    if not _compare_text(text_a, text_b, ignore_whitespace=ignore_whitespace, ignore_case=ignore_case):
        changes.append(
            Change(
                kind=ChangeKind.TEXT_CHANGED,
                old_value=text_a,
                new_value=text_b,
                shape_name=shape_name,
            )
        )

    # -- Font-level details (only at "detailed" level) --
    if detail_level == "detailed":
        paras_a = info_a.get("paragraphs", [])
        paras_b = info_b.get("paragraphs", [])

        # Compare run-level formatting for common paragraphs
        for pi in range(min(len(paras_a), len(paras_b))):
            runs_a = paras_a[pi].get("runs", [])
            runs_b = paras_b[pi].get("runs", [])

            for ri in range(min(len(runs_a), len(runs_b))):
                ra = runs_a[ri]
                rb = runs_b[ri]

                font_attrs = ["font_name", "font_size", "bold", "italic", "font_color"]
                for attr in font_attrs:
                    va = ra.get(attr)
                    vb = rb.get(attr)
                    if va != vb:
                        changes.append(
                            Change(
                                kind=ChangeKind.FONT_CHANGED,
                                old_value=str(va),
                                new_value=str(vb),
                                shape_name=shape_name,
                                details=f"paragraph {pi} run {ri} {attr}",
                            )
                        )

    return changes


# ---------------------------------------------------------------------------
# Slide-level background comparison
# ---------------------------------------------------------------------------

def _extract_background(slide: Any) -> dict[str, Any] | None:
    """Try to extract background information from a slide."""
    try:
        bg = slide.background
        fill = bg.fill
        result: dict[str, Any] = {"fill_type": str(fill.type) if fill.type is not None else None}
        if fill.type is not None and hasattr(fill, "fore_color"):
            try:
                result["color"] = str(fill.fore_color.rgb)
            except Exception:
                pass
        return result
    except Exception:
        return None


def _slide_layout_name(slide: Any) -> str:
    """Return the layout name of a slide, or empty string."""
    try:
        return slide.slide_layout.name or ""
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def diff_text(
    text_a: str,
    text_b: str,
    *,
    ignore_whitespace: bool = True,
    ignore_case: bool = False,
) -> list[Change]:
    """Compare two text strings and return a list of ``Change`` objects.

    This is a convenience function for quick text-only comparison.
    """
    if _compare_text(text_a, text_b, ignore_whitespace=ignore_whitespace, ignore_case=ignore_case):
        return []
    return [
        Change(
            kind=ChangeKind.TEXT_CHANGED,
            old_value=text_a,
            new_value=text_b,
        )
    ]


def diff_slides(
    slide_a: Any,
    slide_b: Any,
    *,
    slide_index: int = 0,
    ignore_whitespace: bool = True,
    ignore_case: bool = False,
    ignore_order: bool = False,
    detail_level: str = "normal",
) -> SlideDiff:
    """Compare two individual python-pptx slide objects.

    Args:
        slide_a: The older slide object.
        slide_b: The newer slide object.
        slide_index: 0-based index (used to populate ``SlideDiff.slide_index``).
        ignore_whitespace: Collapse whitespace before comparing text.
        ignore_case: Case-insensitive text comparison.
        ignore_order: If True, sort shapes by name before matching to reduce
            false positives from re-ordered shapes.
        detail_level: ``"brief"`` (slide-level only), ``"normal"`` (shape-level),
            ``"detailed"`` (text-run level with font formatting).

    Returns:
        A ``SlideDiff`` containing all detected changes.
    """
    result = SlideDiff(slide_index=slide_index)
    changes: list[Change] = []

    # -- Layout name --
    layout_a = _slide_layout_name(slide_a)
    layout_b = _slide_layout_name(slide_b)
    if layout_a != layout_b:
        changes.append(
            Change(
                kind=ChangeKind.LAYOUT_CHANGED,
                old_value=layout_a,
                new_value=layout_b,
            )
        )

    # -- Background --
    bg_a = _extract_background(slide_a)
    bg_b = _extract_background(slide_b)
    if bg_a != bg_b:
        changes.append(
            Change(
                kind=ChangeKind.COLOR_CHANGED,
                old_value=bg_a,
                new_value=bg_b,
                details="slide background",
            )
        )

    # Brief mode stops here (only slide-level changes).
    if detail_level == "brief":
        result.changes = changes
        return result

    # -- Extract shape info --
    infos_a = [_extract_shape_info(s) for s in slide_a.shapes]
    infos_b = [_extract_shape_info(s) for s in slide_b.shapes]

    if ignore_order:
        infos_a.sort(key=lambda d: d.get("name", ""))
        infos_b.sort(key=lambda d: d.get("name", ""))

    # -- Match shapes --
    pairs = _match_shapes(infos_a, infos_b)

    for ia, ib in pairs:
        if ia is not None and ib is None:
            # Shape removed
            changes.append(
                Change(
                    kind=ChangeKind.SHAPE_REMOVED,
                    old_value=infos_a[ia].get("name", f"shape_{ia}"),
                    shape_name=infos_a[ia].get("name"),
                    details=f"type={infos_a[ia]['shape_type']}",
                )
            )
        elif ia is None and ib is not None:
            # Shape added
            changes.append(
                Change(
                    kind=ChangeKind.SHAPE_ADDED,
                    new_value=infos_b[ib].get("name", f"shape_{ib}"),
                    shape_name=infos_b[ib].get("name"),
                    details=f"type={infos_b[ib]['shape_type']}",
                )
            )
        else:
            # Both present -- compare details
            assert ia is not None and ib is not None
            shape_changes = _compare_shape_details(
                infos_a[ia],
                infos_b[ib],
                detail_level=detail_level,
                ignore_whitespace=ignore_whitespace,
                ignore_case=ignore_case,
            )
            changes.extend(shape_changes)

    result.changes = changes
    return result


def diff_presentations(
    path_a: str | Path,
    path_b: str | Path,
    *,
    ignore_whitespace: bool = True,
    ignore_case: bool = False,
    ignore_order: bool = False,
    detail_level: Literal["brief", "normal", "detailed"] = "normal",
) -> PresentationDiff:
    """Compare two PPTX files and return a structured diff.

    Args:
        path_a: Filesystem path to the older PPTX file.
        path_b: Filesystem path to the newer PPTX file.
        ignore_whitespace: Collapse whitespace before comparing text.
        ignore_case: Case-insensitive text comparison.
        ignore_order: If True, sort shapes by name before matching.
        detail_level:
            ``"brief"`` -- only slide-level changes (count, layout, background).
            ``"normal"`` -- shape-level comparison (position, size, text, images).
            ``"detailed"`` -- text-run level including font formatting.

    Returns:
        A ``PresentationDiff`` with all detected changes and a summary.

    Raises:
        FileNotFoundError: If either file does not exist.
        ValueError: If a file cannot be opened as a PPTX (corrupt or
            password-protected).
    """
    # Lazy import to keep module importable without python-pptx.
    from pptx import Presentation

    path_a = Path(path_a)
    path_b = Path(path_b)

    for p in (path_a, path_b):
        if not p.exists():
            raise FileNotFoundError(f"PPTX file not found: {p}")

    try:
        prs_a = Presentation(str(path_a))
    except Exception as exc:
        raise ValueError(f"Cannot open PPTX file {path_a}: {exc}") from exc

    try:
        prs_b = Presentation(str(path_b))
    except Exception as exc:
        raise ValueError(f"Cannot open PPTX file {path_b}: {exc}") from exc

    slides_a = list(prs_a.slides)
    slides_b = list(prs_b.slides)
    count_a = len(slides_a)
    count_b = len(slides_b)

    result = PresentationDiff(
        slide_count_old=count_a,
        slide_count_new=count_b,
    )

    # -- Determine added / removed slides --
    common = min(count_a, count_b)
    if count_b > count_a:
        result.slides_added = list(range(count_a, count_b))
    if count_a > count_b:
        result.slides_removed = list(range(count_b, count_a))

    # -- Diff common slides --
    for i in range(common):
        sd = diff_slides(
            slides_a[i],
            slides_b[i],
            slide_index=i,
            ignore_whitespace=ignore_whitespace,
            ignore_case=ignore_case,
            ignore_order=ignore_order,
            detail_level=detail_level,
        )
        result.slide_diffs.append(sd)

    # -- Build summary --
    total_changes = sum(len(sd.changes) for sd in result.slide_diffs)
    parts: list[str] = []
    if count_a != count_b:
        parts.append(f"slide count {count_a} -> {count_b}")
    if result.slides_added:
        parts.append(f"slides added: {result.slides_added}")
    if result.slides_removed:
        parts.append(f"slides removed: {result.slides_removed}")
    if total_changes:
        parts.append(f"{total_changes} change(s) across {len(result.slide_diffs)} slide(s)")
    result.summary = "; ".join(parts) if parts else "No differences detected"

    return result


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def has_changes(diff_result: PresentationDiff) -> bool:
    """Return True if *diff_result* contains any changes."""
    if diff_result.slides_added or diff_result.slides_removed:
        return True
    return any(sd.changes for sd in diff_result.slide_diffs)


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def format_diff(
    diff_result: PresentationDiff,
    *,
    format: Literal["text", "json", "markdown"] = "text",
) -> str:
    """Format a ``PresentationDiff`` for display.

    Args:
        diff_result: The diff result to format.
        format:
            ``"text"`` -- human-readable plain text.
            ``"json"`` -- structured JSON.
            ``"markdown"`` -- Markdown tables.

    Returns:
        The formatted string.
    """
    if format == "json":
        return _format_json(diff_result)
    if format == "markdown":
        return _format_markdown(diff_result)
    return _format_text(diff_result)


def _format_text(diff_result: PresentationDiff) -> str:
    """Human-readable plain-text diff."""
    lines: list[str] = []
    lines.append(f"PPTX Diff: {diff_result.slide_count_old} -> {diff_result.slide_count_new} slides")
    lines.append(diff_result.summary)
    lines.append("")

    if diff_result.slides_added:
        lines.append(f"  Slides added (indices): {diff_result.slides_added}")
    if diff_result.slides_removed:
        lines.append(f"  Slides removed (indices): {diff_result.slides_removed}")

    for sd in diff_result.slide_diffs:
        if not sd.changes:
            lines.append(f"  Slide {sd.slide_index}: no changes")
            continue
        lines.append(f"  Slide {sd.slide_index}: {len(sd.changes)} change(s)")
        for c in sd.changes:
            shape_tag = f" [{c.shape_name}]" if c.shape_name else ""
            detail_tag = f" ({c.details})" if c.details else ""
            old_tag = f" old={c.old_value!r}" if c.old_value is not None else ""
            new_tag = f" new={c.new_value!r}" if c.new_value is not None else ""
            lines.append(f"    - {c.kind.value}{shape_tag}{detail_tag}:{old_tag}{new_tag}")

    return "\n".join(lines)


def _format_json(diff_result: PresentationDiff) -> str:
    """Structured JSON diff."""
    data = {
        "slide_count_old": diff_result.slide_count_old,
        "slide_count_new": diff_result.slide_count_new,
        "slides_added": diff_result.slides_added,
        "slides_removed": diff_result.slides_removed,
        "summary": diff_result.summary,
        "slide_diffs": [
            {
                "slide_index": sd.slide_index,
                "changes": [
                    {
                        "kind": c.kind.value,
                        "old_value": c.old_value,
                        "new_value": c.new_value,
                        "shape_name": c.shape_name,
                        "details": c.details,
                    }
                    for c in sd.changes
                ],
            }
            for sd in diff_result.slide_diffs
        ],
    }
    return json.dumps(data, ensure_ascii=False, indent=2)


def _format_markdown(diff_result: PresentationDiff) -> str:
    """Markdown table diff."""
    lines: list[str] = []
    lines.append("# PPTX Diff Report")
    lines.append("")
    lines.append("| Metric | Value |")
    lines.append("|--------|-------|")
    lines.append(f"| Old slide count | {diff_result.slide_count_old} |")
    lines.append(f"| New slide count | {diff_result.slide_count_new} |")
    lines.append(f"| Slides added | {diff_result.slides_added or '-'} |")
    lines.append(f"| Slides removed | {diff_result.slides_removed or '-'} |")
    lines.append(f"| Summary | {diff_result.summary} |")
    lines.append("")

    for sd in diff_result.slide_diffs:
        if not sd.changes:
            lines.append(f"## Slide {sd.slide_index}: no changes")
            lines.append("")
            continue
        lines.append(f"## Slide {sd.slide_index}: {len(sd.changes)} change(s)")
        lines.append("")
        lines.append("| Kind | Shape | Details | Old | New |")
        lines.append("|------|-------|---------|-----|-----|")
        for c in sd.changes:
            shape = c.shape_name or "-"
            details = c.details or "-"
            old = str(c.old_value) if c.old_value is not None else "-"
            new = str(c.new_value) if c.new_value is not None else "-"
            # Escape pipe characters in cell content
            old = old.replace("|", "\\|")
            new = new.replace("|", "\\|")
            lines.append(f"| {c.kind.value} | {shape} | {details} | {old} | {new} |")
        lines.append("")

    return "\n".join(lines)


def diff_presentations_visual(
    path_a: str | Path,
    path_b: str | Path,
    *,
    dpi: int = 150,
    threshold: float = 0.02,
) -> dict[str, Any]:
    """Pixel-level visual diff between two PPTX files.

    Renders each slide to a PNG image, then compares corresponding slides
    using SSIM (structural similarity).  Falls back to a simple RMS pixel
    difference if scikit-image is not available.

    This is complementary to :func:`diff_presentations` -- it catches visual
    changes (font substitution, rendering differences) that structural diff
    cannot detect.

    Args:
        path_a: Path to the first (baseline) PPTX.
        path_b: Path to the second (modified) PPTX.
        dpi: Rendering resolution for the PNG images.
        threshold: SSIM threshold below which a slide is flagged as changed
                   (1.0 = identical, 0.0 = completely different).

    Returns:
        A dict with keys:
        - ``slide_count_a``, ``slide_count_b`` -- slide counts.
        - ``slides_compared`` -- number of slides compared.
        - ``slide_scores`` -- list of ``{"index": N, "ssim": F, "changed": bool}``.
        - ``visual_changes`` -- count of slides with SSIM below threshold.
        - ``rasterizer`` -- which backend was used (``"pymupdf"``/``"pillow"``/``"none"``).
    """
    from pathlib import Path as _P

    path_a, path_b = _P(path_a), _P(path_b)
    result: dict[str, Any] = {
        "slide_count_a": 0,
        "slide_count_b": 0,
        "slides_compared": 0,
        "slide_scores": [],
        "visual_changes": 0,
        "rasterizer": "none",
    }

    try:
        from pptx import Presentation  # noqa: E402
    except ImportError:
        return result

    try:
        prs_a = Presentation(str(path_a))
        prs_b = Presentation(str(path_b))
    except Exception:
        return result

    result["slide_count_a"] = len(prs_a.slides)
    result["slide_count_b"] = len(prs_b.slides)

    # Render slides to PNG using preview_renderer
    import tempfile

    try:
        from pptx_skill.preview_renderer import render_preview
    except ImportError:
        return result

    with tempfile.TemporaryDirectory() as tmp_a, tempfile.TemporaryDirectory() as tmp_b:
        try:
            render_preview(str(path_a), tmp_a, dpi=dpi)
            render_preview(str(path_b), tmp_b, dpi=dpi)
        except Exception:
            return result

        import glob

        pngs_a = sorted(glob.glob(str(_P(tmp_a) / "*.png")))
        pngs_b = sorted(glob.glob(str(_P(tmp_b) / "*.png")))

        n_compare = min(len(pngs_a), len(pngs_b))
        if n_compare == 0:
            return result

        result["rasterizer"] = "pymupdf"

        # Compare images
        try:
            import numpy as np
            from skimage.metrics import structural_similarity as ssim

            use_skimage = True
            result["rasterizer"] = "skimage"
        except ImportError:
            use_skimage = False

        for i in range(n_compare):
            try:
                from PIL import Image

                img_a = Image.open(pngs_a[i]).convert("L")
                img_b = Image.open(pngs_b[i]).convert("L")

                # Resize to same dimensions if needed
                if img_a.size != img_b.size:
                    img_b = img_b.resize(img_a.size)

                if use_skimage:
                    import numpy as np

                    arr_a = np.array(img_a)
                    arr_b = np.array(img_b)
                    score = float(ssim(arr_a, arr_b))
                else:
                    # Simple RMS difference as fallback
                    import numpy as np

                    arr_a = np.array(img_a, dtype=float)
                    arr_b = np.array(img_b, dtype=float)
                    diff = arr_a - arr_b
                    rms = float(np.sqrt(np.mean(diff ** 2)))
                    # Convert RMS (0-255) to pseudo-SSIM (0-1)
                    score = max(0.0, 1.0 - rms / 128.0)

                changed = score < threshold
                if changed:
                    result["visual_changes"] += 1
                result["slide_scores"].append({
                    "index": i,
                    "ssim": round(score, 4),
                    "changed": changed,
                })
            except Exception:
                result["slide_scores"].append({
                    "index": i,
                    "ssim": None,
                    "changed": None,
                    "error": "comparison failed",
                })

    result["slides_compared"] = n_compare
    return result
