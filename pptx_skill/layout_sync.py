"""Layout synchronization — align, distribute, and sync shapes across slides.

Provides tools for precisely positioning shapes and maintaining layout
consistency across slides:

- **Alignment**: align shapes to each other (left, right, center, top, bottom, middle)
- **Distribution**: evenly space shapes horizontally or vertically
- **Snap to grid**: snap shape positions to a configurable grid
- **Layout sync**: copy layout from one slide to another (position, size, style)
- **Master layout enforcement**: ensure all slides follow a consistent layout template
- **Z-order management**: bring to front, send to back, move up/down

Quick start
-----------
>>> from pptx_skill.layout_sync import align_shapes, distribute_shapes
>>> align_shapes("deck.pptx", 1, shape_names=["Title", "Subtitle"], alignment="center_h")
>>> distribute_shapes("deck.pptx", 1, shape_names=["Box1", "Box2", "Box3"], direction="horizontal")
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field

__all__ = [
    "LayoutTemplate",
    "align_shapes",
    "distribute_shapes",
    "snap_to_grid",
    "copy_layout",
    "apply_layout_template",
    "create_layout_template",
    "list_layout_templates",
    "bring_to_front",
    "send_to_back",
    "move_up",
    "move_down",
    "set_z_order",
    "match_size",
    "match_position",
    "center_on_slide",
]

log = logging.getLogger(__name__)

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Grid constants (in EMU)
_DEFAULT_GRID = 914400 // 8  # 1/8 inch = 114300 EMU


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class LayoutTemplate:
    """A layout template capturing shape positions from a reference slide."""
    name: str = ""
    shapes: list[dict] = field(default_factory=list)  # Each dict: {name, left, top, width, height, font_size, ...}

    def __post_init__(self):
        if self.shapes is None:
            self.shapes = []


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path):
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path.
    Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs, path):
    """Save *prs* back to *path* if *path* is not None."""
    if path is not None:
        prs.save(str(path))


def _find_shape(slide, shape_name: str):
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None


def _get_shape_bounds(shape) -> dict[str, int]:
    """Get shape bounds in EMU."""
    return {
        "left": shape.left,
        "top": shape.top,
        "width": shape.width,
        "height": shape.height,
        "right": shape.left + shape.width,
        "bottom": shape.top + shape.height,
        "center_h": shape.left + shape.width // 2,
        "center_v": shape.top + shape.height // 2,
    }


# ---------------------------------------------------------------------------
# Alignment
# ---------------------------------------------------------------------------

def align_shapes(
    prs_or_path,
    slide_index: int,
    *,
    shape_names: list[str],
    alignment: str,
    reference: str = "first",
) -> int:
    """Align multiple shapes relative to each other.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    shape_names : list[str]
        Names of shapes to align. At least 2 required.
    alignment : str
        Alignment type:
        - ``"left"``: align left edges
        - ``"right"``: align right edges
        - ``"center_h"``: align horizontal centers
        - ``"top"``: align top edges
        - ``"bottom"``: align bottom edges
        - ``"center_v"``: align vertical centers
    reference : str
        Reference shape: ``"first"`` (first in list), ``"last"``, ``"largest"``,
        or a specific shape name.

    Returns
    -------
    int
        Number of shapes aligned (excluding the reference).
    """
    if len(shape_names) < 2:
        raise ValueError("At least 2 shapes required for alignment")

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]

        # Find all shapes
        shapes = []
        for name in shape_names:
            shape = _find_shape(slide, name)
            if shape is None:
                log.warning("Shape not found: %r", name)
                continue
            shapes.append(shape)

        if len(shapes) < 2:
            return 0

        # Determine reference shape
        ref_shape = _determine_reference(shapes, reference, shape_names)
        if ref_shape is None:
            return 0

        ref_bounds = _get_shape_bounds(ref_shape)
        count = 0

        for shape in shapes:
            if shape is ref_shape:
                continue

            _get_shape_bounds(shape)

            if alignment == "left":
                shape.left = ref_bounds["left"]
            elif alignment == "right":
                shape.left = ref_bounds["right"] - shape.width
            elif alignment == "center_h":
                shape.left = ref_bounds["center_h"] - shape.width // 2
            elif alignment == "top":
                shape.top = ref_bounds["top"]
            elif alignment == "bottom":
                shape.top = ref_bounds["bottom"] - shape.height
            elif alignment == "center_v":
                shape.top = ref_bounds["center_v"] - shape.height // 2
            else:
                raise ValueError(f"Unknown alignment: {alignment!r}")

            count += 1

        return count
    finally:
        _save_prs(prs, path)


def _determine_reference(shapes, reference: str, shape_names: list[str]):
    """Determine the reference shape for alignment."""
    if reference == "first":
        return shapes[0]
    elif reference == "last":
        return shapes[-1]
    elif reference == "largest":
        return max(shapes, key=lambda s: s.width * s.height)
    elif reference in shape_names:
        idx = shape_names.index(reference)
        return shapes[idx] if idx < len(shapes) else shapes[0]
    else:
        return shapes[0]


# ---------------------------------------------------------------------------
# Distribution
# ---------------------------------------------------------------------------

def distribute_shapes(
    prs_or_path,
    slide_index: int,
    *,
    shape_names: list[str],
    direction: str = "horizontal",
    gap: int | None = None,
) -> int:
    """Distribute shapes evenly with equal spacing.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    direction : str
        ``"horizontal"`` or ``"vertical"``.
    gap : int, optional
        Gap between shapes in EMU. If None, shapes are distributed evenly
        across the span from first to last shape.

    Returns
    -------
    int
        Number of shapes moved.
    """
    if len(shape_names) < 2:
        return 0

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]

        shapes = []
        for name in shape_names:
            shape = _find_shape(slide, name)
            if shape is not None:
                shapes.append(shape)

        if len(shapes) < 2:
            return 0

        # Sort by position
        if direction == "horizontal":
            shapes.sort(key=lambda s: s.left)
        else:
            shapes.sort(key=lambda s: s.top)

        count = 0

        if gap is not None:
            # Fixed gap distribution
            if direction == "horizontal":
                current = shapes[0].left
                for shape in shapes:
                    shape.left = current
                    current += shape.width + gap
                    count += 1
            else:
                current = shapes[0].top
                for shape in shapes:
                    shape.top = current
                    current += shape.height + gap
                    count += 1
        else:
            # Even distribution across span
            if direction == "horizontal":
                first_left = shapes[0].left
                last_right = shapes[-1].left + shapes[-1].width
                total_width = sum(s.width for s in shapes)
                total_gap = last_right - first_left - total_width
                gap_size = total_gap / max(1, len(shapes) - 1)

                current = first_left
                for shape in shapes:
                    shape.left = int(current)
                    current += shape.width + gap_size
                    count += 1
            else:
                first_top = shapes[0].top
                last_bottom = shapes[-1].top + shapes[-1].height
                total_height = sum(s.height for s in shapes)
                total_gap = last_bottom - first_top - total_height
                gap_size = total_gap / max(1, len(shapes) - 1)

                current = first_top
                for shape in shapes:
                    shape.top = int(current)
                    current += shape.height + gap_size
                    count += 1

        return count
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Snap to grid
# ---------------------------------------------------------------------------

def snap_to_grid(
    prs_or_path,
    slide_index: int,
    *,
    grid_size: int = _DEFAULT_GRID,
    shape_names: list[str] | None = None,
) -> int:
    """Snap shape positions to a grid.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    grid_size : int
        Grid size in EMU. Default is 1/8 inch (114300 EMU).
    shape_names : list[str], optional
        Specific shapes to snap. If None, snap all shapes on the slide.

    Returns
    -------
    int
        Number of shapes snapped.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        count = 0

        for shape in slide.shapes:
            if shape_names is not None and shape.name not in shape_names:
                continue

            # Snap left and top
            new_left = round(shape.left / grid_size) * grid_size
            new_top = round(shape.top / grid_size) * grid_size

            if new_left != shape.left or new_top != shape.top:
                shape.left = new_left
                shape.top = new_top
                count += 1

        return count
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Layout sync
# ---------------------------------------------------------------------------

def copy_layout(
    prs_or_path,
    source_slide_index: int,
    target_slide_index: int,
    *,
    shape_mapping: dict[str, str] | None = None,
    copy_size: bool = True,
    copy_position: bool = True,
    copy_style: bool = False,
) -> int:
    """Copy layout from one slide to another.

    Parameters
    ----------
    source_slide_index : int
        1-based slide index of the source slide (1 = first slide).
    target_slide_index : int
        1-based slide index of the target slide (1 = first slide).
    shape_mapping : dict[str, str], optional
        Map of source shape name → target shape name. If None, match by name.
    copy_size : bool
        Copy width/height from source to target.
    copy_position : bool
        Copy left/top from source to target.
    copy_style : bool
        Copy font size, fill color, and line color.

    Returns
    -------
    int
        Number of shapes synchronized.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    n = len(prs.slides)
    if source_slide_index < 1 or source_slide_index > n:
        raise IndexError(f"source_slide_index {source_slide_index} out of range (1..{n})")
    if target_slide_index < 1 or target_slide_index > n:
        raise IndexError(f"target_slide_index {target_slide_index} out of range (1..{n})")
    try:
        source_slide = prs.slides[source_slide_index - 1]
        target_slide = prs.slides[target_slide_index - 1]
        count = 0

        # Build mapping
        if shape_mapping is None:
            # Auto-match by name
            target_names = {s.name for s in target_slide.shapes}
            shape_mapping = {}
            for shape in source_slide.shapes:
                if shape.name in target_names:
                    shape_mapping[shape.name] = shape.name

        for src_name, tgt_name in shape_mapping.items():
            src_shape = _find_shape(source_slide, src_name)
            tgt_shape = _find_shape(target_slide, tgt_name)
            if src_shape is None or tgt_shape is None:
                continue

            if copy_position:
                tgt_shape.left = src_shape.left
                tgt_shape.top = src_shape.top

            if copy_size:
                tgt_shape.width = src_shape.width
                tgt_shape.height = src_shape.height

            if copy_style:
                _copy_shape_style(src_shape, tgt_shape)

            count += 1

        return count
    finally:
        _save_prs(prs, path)


def _copy_shape_style(source, target):
    """Copy style properties from source shape to target shape."""
    try:
        # Fill
        if source.has_text_frame and target.has_text_frame:
            for src_para, tgt_para in zip(
                source.text_frame.paragraphs,
                target.text_frame.paragraphs, strict=False,
            ):
                for src_run, tgt_run in zip(src_para.runs, tgt_para.runs, strict=False):
                    try:
                        tgt_run.font.size = src_run.font.size
                    except Exception:
                        pass
                    try:
                        tgt_run.font.bold = src_run.font.bold
                    except Exception:
                        pass
                    try:
                        tgt_run.font.color.rgb = src_run.font.color.rgb
                    except Exception:
                        pass
    except Exception:
        pass


def create_layout_template(
    prs_or_path,
    slide_index: int,
    *,
    name: str = "",
    shape_names: list[str] | None = None,
) -> LayoutTemplate:
    """Create a layout template from a slide's shape positions.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    shape_names : list[str], optional
        Specific shapes to include. If None, include all shapes.

    Returns
    -------
    LayoutTemplate
    """
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        template = LayoutTemplate(name=name or f"Template from slide {slide_index}")

        for shape in slide.shapes:
            if shape_names is not None and shape.name not in shape_names:
                continue

            shape_info = {
                "name": shape.name,
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
            }

            # Capture font size if text frame
            try:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.size:
                                shape_info["font_size"] = run.font.size
                                break
                        break
            except Exception:
                pass

            template.shapes.append(shape_info)

        return template
    finally:
        pass


def apply_layout_template(
    prs_or_path,
    slide_index: int,
    template: LayoutTemplate,
    *,
    shape_mapping: dict[str, str] | None = None,
) -> int:
    """Apply a layout template to a slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    shape_mapping : dict[str, str], optional
        Map of template shape name → slide shape name. If None, match by name.

    Returns
    -------
    int
        Number of shapes positioned.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        count = 0

        for shape_info in template.shapes:
            tmpl_name = shape_info["name"]
            target_name = shape_mapping.get(tmpl_name, tmpl_name) if shape_mapping else tmpl_name

            shape = _find_shape(slide, target_name)
            if shape is None:
                continue

            shape.left = shape_info.get("left", shape.left)
            shape.top = shape_info.get("top", shape.top)
            shape.width = shape_info.get("width", shape.width)
            shape.height = shape_info.get("height", shape.height)

            # Font size
            if "font_size" in shape_info:
                try:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                run.font.size = shape_info["font_size"]
                except Exception:
                    pass

            count += 1

        return count
    finally:
        _save_prs(prs, path)


def list_layout_templates(prs_or_path) -> list[LayoutTemplate]:
    """Create layout templates from all slides (one template per slide)."""
    prs = _open_prs(prs_or_path)
    try:
        templates = []
        for idx in range(1, len(prs.slides) + 1):
            template = create_layout_template(prs, idx, name=f"Slide {idx}")
            templates.append(template)
        return templates
    finally:
        pass


# ---------------------------------------------------------------------------
# Z-order management
# ---------------------------------------------------------------------------

def bring_to_front(prs_or_path, slide_index: int, *, shape_name: str) -> bool:
    """Move a shape to the front (top of z-order).

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_tree = slide.shapes._spTree
        elem = shape._element
        sp_tree.remove(elem)
        sp_tree.append(elem)
        return True
    finally:
        _save_prs(prs, path)


def send_to_back(prs_or_path, slide_index: int, *, shape_name: str) -> bool:
    """Move a shape to the back (bottom of z-order).

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_tree = slide.shapes._spTree
        elem = shape._element
        sp_tree.remove(elem)
        # Insert after the first child (which is typically nvGrpSpPr/grpSpPr)
        children = list(sp_tree)
        if len(children) > 2:
            sp_tree.insert(2, elem)
        else:
            sp_tree.append(elem)
        return True
    finally:
        _save_prs(prs, path)


def move_up(prs_or_path, slide_index: int, *, shape_name: str) -> bool:
    """Move a shape one position up in z-order.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_tree = slide.shapes._spTree
        elem = shape._element

        # Find next sibling
        next_elem = elem.getnext()
        if next_elem is not None:
            sp_tree.remove(elem)
            next_elem.addnext(elem)
            return True
        return False
    finally:
        _save_prs(prs, path)


def move_down(prs_or_path, slide_index: int, *, shape_name: str) -> bool:
    """Move a shape one position down in z-order.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_tree = slide.shapes._spTree
        elem = shape._element

        # Find previous sibling (skip the first 2 metadata elements)
        prev_elem = elem.getprevious()
        if prev_elem is not None:
            sp_tree.remove(elem)
            prev_elem.addprevious(elem)
            return True
        return False
    finally:
        _save_prs(prs, path)


def set_z_order(prs_or_path, slide_index: int, *, shape_name: str, position: int) -> bool:
    """Set a shape to a specific z-order position.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    position : int
        0 = back, -1 = front.
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        sp_tree = slide.shapes._spTree
        elem = shape._element
        sp_tree.remove(elem)

        # Count shape children (skip first 2 metadata elements)
        shape_children = [c for c in sp_tree if c.tag not in (
            f"{{{_NS_P}}}nvGrpSpPr", f"{{{_NS_P}}}grpSpPr"
        )]

        if position < 0 or position >= len(shape_children):
            sp_tree.append(elem)
        else:
            # Insert at position (offset by 2 for metadata elements)
            insert_idx = min(position + 2, len(sp_tree))
            sp_tree.insert(insert_idx, elem)

        return True
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Size & position matching
# ---------------------------------------------------------------------------

def match_size(
    prs_or_path,
    slide_index: int,
    *,
    source_name: str,
    target_names: list[str],
) -> int:
    """Match the size of target shapes to a source shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        source = _find_shape(slide, source_name)
        if source is None:
            return 0

        count = 0
        for name in target_names:
            target = _find_shape(slide, name)
            if target is None:
                continue
            target.width = source.width
            target.height = source.height
            count += 1

        return count
    finally:
        _save_prs(prs, path)


def match_position(
    prs_or_path,
    slide_index: int,
    *,
    source_name: str,
    target_names: list[str],
) -> int:
    """Match the position of target shapes to a source shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        source = _find_shape(slide, source_name)
        if source is None:
            return 0

        count = 0
        for name in target_names:
            target = _find_shape(slide, name)
            if target is None:
                continue
            target.left = source.left
            target.top = source.top
            count += 1

        return count
    finally:
        _save_prs(prs, path)


def center_on_slide(
    prs_or_path,
    slide_index: int,
    *,
    shape_name: str,
    center: str = "both",
) -> bool:
    """Center a shape on the slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    center : str
        ``"horizontal"``, ``"vertical"``, or ``"both"``.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        slide_w = prs.slide_width
        slide_h = prs.slide_height

        if center in ("horizontal", "both"):
            shape.left = (slide_w - shape.width) // 2

        if center in ("vertical", "both"):
            shape.top = (slide_h - shape.height) // 2

        return True
    finally:
        _save_prs(prs, path)
