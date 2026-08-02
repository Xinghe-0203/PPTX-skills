"""Watermark support for PPTX presentations.

Provides text and image watermarks with opacity control, positional layout,
tiling, z-ordering, and removal/listing utilities. All watermark shapes are
tagged with a recognizable name prefix (``pptx_skill_watermark``) so they can
be identified and removed later.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

__all__ = [
    "add_text_watermark",
    "add_image_watermark",
    "remove_watermark",
    "list_watermarks",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_WATERMARK_PREFIX = "pptx_skill_watermark"
_WATERMARK_TEXT = f"{_WATERMARK_PREFIX}_text"
_WATERMARK_IMAGE = f"{_WATERMARK_PREFIX}_image"

# DrawingML namespace
_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_DML_NS_PREFIX = f"{{{_DML_NS}}}"

# Relationship namespace for blip references
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _is_presentation(obj: Any) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path: Any) -> Any:
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path
    (``str | Path``).  Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs: Any, path: str | Path | None) -> None:
    """Save *prs* back to *path*, creating a backup first."""
    if path is None:
        return

    p = Path(path)
    bak = p.with_suffix(".bak.pptx")
    if p.exists():
        import shutil

        shutil.copy2(str(p), str(bak))
    prs.save(str(p))


def _validate_opacity(opacity: float) -> None:
    """Raise ``ValueError`` if *opacity* is outside [0, 1]."""
    if not 0.0 <= opacity <= 1.0:
        raise ValueError(f"opacity must be between 0.0 and 1.0, got {opacity}")


def _alpha_val(opacity: float) -> str:
    """Convert 0-1 *opacity* to OOXML alpha value string (1/1000 percent).

    E.g. 0.3 -> "30000", 1.0 -> "100000".
    """
    return str(int(opacity * 100000))


def _resolve_slides(prs: Any, slides: list[int] | None) -> list[Any]:
    """Return a list of slide objects for the given 1-based *slides* indices.

    If *slides* is ``None``, return all slides in the presentation.
    """
    all_slides = list(prs.slides)
    if not all_slides:
        return []
    if slides is None:
        return all_slides
    # Validate and convert 1-based indices to 0-based
    n = len(all_slides)
    for i in slides:
        if i < 1 or i > n:
            raise IndexError(f"slide index {i} out of range (1..{n})")
    return [all_slides[i - 1] for i in slides]


def _position_rect(
    position: str,
    slide_width: int,
    slide_height: int,
    shape_width: int,
    shape_height: int,
) -> tuple[int, int]:
    """Return (left, top) Emu coordinates for *position*.

    Supported positions: ``center``, ``top-left``, ``top-right``,
    ``bottom-left``, ``bottom-right``, ``diagonal``.
    """
    margin = 914400  # ~1 cm margin
    sw, sh = slide_width, slide_height
    ww, hh = shape_width, shape_height

    if position == "center":
        return (sw - ww) // 2, (sh - hh) // 2
    elif position == "top-left":
        return margin, margin
    elif position == "top-right":
        return sw - ww - margin, margin
    elif position == "bottom-left":
        return margin, sh - hh - margin
    elif position == "bottom-right":
        return sw - ww - margin, sh - hh - margin
    elif position == "diagonal":
        # Place along the diagonal from top-left toward bottom-right
        return (sw - ww) // 3, (sh - hh) // 3
    else:
        log.warning("Unknown position %r, falling back to center", position)
        return (sw - ww) // 2, (sh - hh) // 2


def _tile_positions(
    position: str,
    slide_width: int,
    slide_height: int,
    shape_width: int,
    shape_height: int,
    tile: bool,
) -> list[tuple[int, int]]:
    """Return a list of (left, top) positions.

    If *tile* is True, compute a grid of positions covering the slide.
    Otherwise, return a single position.
    """
    if not tile:
        return [_position_rect(position, slide_width, slide_height, shape_width, shape_height)]

    positions: list[tuple[int, int]] = []
    # Step with 30% overlap so tiles are spaced with slight gap
    h_step = int(shape_width * 1.3) if shape_width > 0 else 1
    v_step = int(shape_height * 1.3) if shape_height > 0 else 1
    # Ensure minimum step so we don't create an enormous number of shapes
    h_step = max(h_step, 914400)  # at least ~1 cm
    v_step = max(v_step, 914400)

    y = 0
    while y < slide_height:
        x = 0
        while x < slide_width:
            positions.append((x, y))
            x += h_step
        y += v_step

    return positions


def _set_z_order(shape: Any, z_order: str) -> None:
    """Move *shape* to front or back of the slide's shape tree."""
    sp = shape._element
    parent = sp.getparent()
    if parent is None:
        return
    parent.remove(sp)
    if z_order == "front":
        parent.append(sp)
    else:
        # back: insert as first child
        parent.insert(0, sp)


def _apply_text_alpha(shape: Any, opacity: float) -> None:
    """Set alpha transparency on all text runs in *shape*."""
    from lxml import etree

    sp = shape._element
    # Find all srgbClr elements under text runs and set alpha
    for srgb in sp.findall(f".//{_DML_NS_PREFIX}srgbClr"):
        # Remove any existing alpha
        for old in srgb.findall(f"{_DML_NS_PREFIX}alpha"):
            srgb.remove(old)
        alpha_el = etree.SubElement(srgb, f"{_DML_NS_PREFIX}alpha")
        alpha_el.set("val", _alpha_val(opacity))


def _apply_shape_fill_alpha(shape: Any, opacity: float) -> None:
    """Set alpha transparency on the shape's solid fill color."""
    from lxml import etree

    sp = shape._element
    sp_pr = sp.find(f"{_DML_NS_PREFIX}spPr")
    if sp_pr is None:
        return
    for srgb in sp_pr.findall(f".//{_DML_NS_PREFIX}srgbClr"):
        for old in srgb.findall(f"{_DML_NS_PREFIX}alpha"):
            srgb.remove(old)
        alpha_el = etree.SubElement(srgb, f"{_DML_NS_PREFIX}alpha")
        alpha_el.set("val", _alpha_val(opacity))


def _apply_image_alpha(shape: Any, opacity: float) -> None:
    """Set alpha transparency on an image shape via alphaModFix."""
    from lxml import etree

    _PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    _PML_NS_PREFIX = f"{{{_PML_NS}}}"

    sp = shape._element
    # Picture shapes use p:blipFill (PresentationML namespace) containing a:blip
    blip_fill = sp.find(f"{_PML_NS_PREFIX}blipFill")
    if blip_fill is None:
        # Fallback: try DrawingML namespace (unlikely but defensive)
        blip_fill = sp.find(f"{_DML_NS_PREFIX}blipFill")
    if blip_fill is None:
        log.warning("Could not find blipFill element in image watermark shape")
        return
    blip = blip_fill.find(f"{_DML_NS_PREFIX}blip")
    if blip is None:
        log.warning("Could not find a:blip element in image watermark shape")
        return

    # Remove existing alphaModFix if any
    for old in blip.findall(f"{_DML_NS_PREFIX}alphaModFix"):
        blip.remove(old)

    alpha_mod = etree.SubElement(blip, f"{_DML_NS_PREFIX}alphaModFix")
    alpha_mod.set("amt", _alpha_val(opacity))


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def add_text_watermark(
    prs_or_path: Any,
    text: str,
    *,
    font_name: str = "Microsoft YaHei",
    font_size: int = 48,
    color: str = "C0C0C0",
    opacity: float = 0.3,
    rotation: float = -45,
    position: str = "center",
    bold: bool = False,
    italic: bool = False,
    slides: list[int] | None = None,
    z_order: str = "back",
    tile: bool = False,
) -> int:
    """Add a semi-transparent text watermark to slides.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        text: Watermark text content.
        font_name: Font family name.
        font_size: Font size in points.
        color: Hex color string (e.g. ``"C0C0C0"``).
        opacity: Transparency from 0.0 (invisible) to 1.0 (opaque).
        rotation: Rotation angle in degrees (default -45 for diagonal).
        position: Placement on the slide — ``center``, ``top-left``,
            ``top-right``, ``bottom-left``, ``bottom-right``, ``diagonal``.
        bold: Whether the watermark text is bold.
        italic: Whether the watermark text is italic.
        slides: 1-based slide indices to watermark. ``None`` means all slides.
        z_order: ``"front"`` or ``"back"`` (default ``"back"``).
        tile: If ``True``, repeat the watermark across the slide.

    Returns:
        Number of watermark shapes created.
    """
    from pptx.dml.color import RGBColor
    from pptx.util import Emu, Pt

    _validate_opacity(opacity)

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)

    target_slides = _resolve_slides(prs, slides)
    if not target_slides:
        return 0

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Estimate shape size based on text length and font size
    # Approximate: each character ~font_size * 0.6 pt wide
    char_width_pt = font_size * 0.6
    shape_width = int(Pt(max(len(text) * char_width_pt, font_size)))
    shape_height = int(Pt(font_size * 1.5))

    count = 0
    hex_color = color.lstrip("#")

    for slide in target_slides:
        positions = _tile_positions(position, slide_width, slide_height, shape_width, shape_height, tile)

        for left, top in positions:
            # Add text box
            txBox = slide.shapes.add_textbox(
                left=Emu(left),
                top=Emu(top),
                width=Emu(shape_width),
                height=Emu(shape_height),
            )

            # Tag the shape
            txBox.name = _WATERMARK_TEXT

            # Set text
            tf = txBox.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.text = text

            # Style the run
            run = p.runs[0]
            run.font.size = Pt(font_size)
            run.font.name = font_name
            run.font.color.rgb = RGBColor.from_string(hex_color)
            run.font.bold = bold
            run.font.italic = italic

            # Apply rotation
            txBox.rotation = rotation

            # Apply alpha transparency to text and fill
            _apply_text_alpha(txBox, opacity)
            _apply_shape_fill_alpha(txBox, opacity)

            # Z-order
            _set_z_order(txBox, z_order)

            count += 1

    if is_path and path is not None:
        _save_prs(prs, path)

    return count


def add_image_watermark(
    prs_or_path: Any,
    image_path: str | Path,
    *,
    opacity: float = 0.3,
    position: str = "center",
    scale: float = 0.5,
    slides: list[int] | None = None,
    z_order: str = "back",
    tile: bool = False,
) -> int:
    """Add a semi-transparent image watermark to slides.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        image_path: Path to the watermark image file.
        opacity: Transparency from 0.0 (invisible) to 1.0 (opaque).
        position: Placement on the slide — ``center``, ``top-left``,
            ``top-right``, ``bottom-left``, ``bottom-right``, ``diagonal``.
        scale: Scale factor relative to slide dimensions (0.0–1.0).
        slides: 1-based slide indices to watermark. ``None`` means all slides.
        z_order: ``"front"`` or ``"back"`` (default ``"back"``).
        tile: If ``True``, repeat the watermark across the slide.

    Returns:
        Number of watermark shapes created.
    """
    from pptx.util import Emu

    _validate_opacity(opacity)

    img = Path(image_path)
    if not img.exists():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)

    target_slides = _resolve_slides(prs, slides)
    if not target_slides:
        return 0

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Compute image shape size based on scale
    shape_width = int(slide_width * scale)
    shape_height = int(slide_height * scale)

    count = 0

    for slide in target_slides:
        positions = _tile_positions(position, slide_width, slide_height, shape_width, shape_height, tile)

        for left, top in positions:
            pic = slide.shapes.add_picture(
                image_file=str(img),
                left=Emu(left),
                top=Emu(top),
                width=Emu(shape_width),
                height=Emu(shape_height),
            )

            # Tag the shape
            pic.name = _WATERMARK_IMAGE

            # Apply alpha transparency to the image
            _apply_image_alpha(pic, opacity)

            # Z-order
            _set_z_order(pic, z_order)

            count += 1

    if is_path and path is not None:
        _save_prs(prs, path)

    return count


def remove_watermark(
    prs_or_path: Any,
    *,
    name_contains: str = "pptx_skill_watermark",
) -> int:
    """Remove all watermark shapes matching *name_contains* from the presentation.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        name_contains: Substring to match against shape names.

    Returns:
        Number of watermark shapes removed.
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)

    count = 0
    for slide in prs.slides:
        # Collect shapes to remove (cannot modify collection while iterating)
        to_remove = [
            shape for shape in slide.shapes if name_contains in shape.name
        ]
        for shape in to_remove:
            sp = shape._element
            parent = sp.getparent()
            if parent is not None:
                parent.remove(sp)
                count += 1

    if is_path and path is not None:
        _save_prs(prs, path)

    return count


def list_watermarks(
    prs_or_path: Any,
    *,
    name_contains: str = "pptx_skill_watermark",
) -> list[dict[str, Any]]:
    """List all watermark shapes in the presentation.

    Args:
        prs_or_path: A ``Presentation`` object or file path (``str | Path``).
        name_contains: Substring to match against shape names.

    Returns:
        A list of dicts with keys ``slide_index`` (1-based), ``shape_name``,
        ``shape_type``, ``left``, ``top``, ``width``, ``height``.
    """

    prs = _open_prs(prs_or_path)
    results: list[dict[str, Any]] = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if name_contains in shape.name:
                results.append(
                    {
                        "slide_index": slide_idx,
                        "shape_name": shape.name,
                        "shape_type": str(shape.shape_type),
                        "left": shape.left,
                        "top": shape.top,
                        "width": shape.width,
                        "height": shape.height,
                    }
                )

    return results
