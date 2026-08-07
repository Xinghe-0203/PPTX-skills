"""Renderer result models and the adaptive PPTX renderer.

At PR1 this module only defines the data contract for RenderTrace and render
results. The actual OOXML writing was added in PR2 and enhanced in subsequent
PRs with rich text, chart types, table styling, shape styling, slide meta,
slide notes, and media support.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx.dml.color import RGBColor

from pptx_skill.content_model import GeometrySpec, LayoutPlan, PlannedNode


@dataclass(frozen=True)
class RenderTraceEntry:
    """One-to-many mapping between a semantic element and a PPTX shape."""

    render_node_id: str
    element_id: str | None
    recipe_node_id: str
    shape_instance_index: int
    slide_index: int
    ppt_shape_id: int | None
    shape_name: str
    z_order: int
    geometry: GeometrySpec
    crop: dict[str, Any] | None
    parent_render_node_id: str | None = None


@dataclass
class PreviewRenderResult:
    """Result of exporting a PPTX to preview images."""

    renderer: str
    renderer_version: str
    target_dpi: int
    slide_pngs: list[str] = field(default_factory=list)
    actual_pixel_sizes: list[tuple[int, int]] = field(default_factory=list)
    environment: dict[str, Any] = field(default_factory=dict)
    attempts: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class RenderResult:
    """Result of generating a PPTX file."""

    pptx_path: str
    trace: list[RenderTraceEntry]
    generation_engine: str
    generation_engine_version: str
    artifacts: dict[str, str] = field(default_factory=dict)
    preview: PreviewRenderResult | None = None


class AdaptiveRendererError(Exception):
    """Raised when the adaptive renderer cannot produce a valid PPTX."""

    pass


# ---------------------------------------------------------------------------
# PR2 adaptive renderer (basic + enhancements)
# ---------------------------------------------------------------------------

log = logging.getLogger(__name__)

# Default deck options for render_layout_plans
_DEFAULT_DECK_OPTIONS: dict[str, Any] = {
    "show_slide_numbers": True,
    "footer_text": "",
    "header_text": "",
    "slide_number_format": "{current}/{total}",
    "theme": None,
}


def _hex_to_rgb(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    if len(color) == 3:
        color = "".join(c * 2 for c in color)
    return (
        int(color[0:2], 16),
        int(color[2:4], 16),
        int(color[4:6], 16),
    )


def _pt_to_inches(pt: float) -> float:
    return pt / 72.0


_SAFE_HYPERLINK_SCHEMES = ("http://", "https://", "mailto:", "slide://")


def _is_safe_hyperlink(href: str) -> bool:
    """Check that a hyperlink URL uses a safe scheme."""
    lower = href.lower().strip()
    return any(lower.startswith(scheme) for scheme in _SAFE_HYPERLINK_SCHEMES)


def _apply_text_style(run, style: dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from pptx.util import Pt

    if "size" in style:
        run.font.size = Pt(style["size"])
    if "color" in style:
        run.font.color.rgb = RGBColor(*_hex_to_rgb(style["color"]))
    if style.get("bold"):
        run.font.bold = True
    if style.get("italic"):
        run.font.italic = True
    # Apply font_family from resolved style (set by layout_engine / deck_planner)
    font_family = style.get("font_family")
    if font_family:
        run.font.name = font_family
        # Also set East-Asian typeface to keep CJK consistency
        try:
            rPr = run._r.get_or_add_rPr()
            ea = rPr.find(qn("a:ea"))
            if ea is None:
                ea = rPr.makeelement(qn("a:ea"), {})
                rPr.append(ea)
            ea.set("typeface", font_family)
        except Exception:
            pass


def _align_from_str(align_str: str):
    """Convert an alignment string to a PP_ALIGN enum value."""
    from pptx.enum.text import PP_ALIGN

    mapping = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }
    return mapping.get(str(align_str).lower(), PP_ALIGN.LEFT)


# ---------------------------------------------------------------------------
# 1. Slide meta: headers / footers / slide numbers
# ---------------------------------------------------------------------------

def _add_slide_meta(slide, slide_index: int, total_slides: int, deck_options: dict[str, Any],
                    slide_width_emu: int, slide_height_emu: int) -> None:
    """Add slide number, footer text, and header text to a slide."""
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt

    theme = deck_options.get("theme") or {}
    # Resolve colours: theme may provide text_muted / text keys; else gray.
    muted_hex = theme.get("text_muted", "#999999")
    theme.get("text", "#333333")

    show_numbers = deck_options.get("show_slide_numbers", True)
    footer_text = deck_options.get("footer_text", "")
    header_text = deck_options.get("header_text", "")
    fmt = deck_options.get("slide_number_format", "{current}/{total}")

    # Convert EMU dimensions to points for positioning calculations
    sw_pt = slide_width_emu / 914400 * 72
    sh_pt = slide_height_emu / 914400 * 72

    # --- Slide number (bottom-right) ---
    if show_numbers:
        num_text = fmt.format(current=slide_index + 1, total=total_slides)
        left = Inches(_pt_to_inches(sw_pt - 60))
        top = Inches(_pt_to_inches(sh_pt - 24))
        width = Inches(0.8)
        height = Inches(0.3)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = num_text
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(*_hex_to_rgb(muted_hex))

    # --- Footer (bottom-center) ---
    if footer_text:
        left = Inches(_pt_to_inches(sw_pt / 2 - 120))
        top = Inches(_pt_to_inches(sh_pt - 24))
        width = Inches(3.3)
        height = Inches(0.3)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = footer_text
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(*_hex_to_rgb(muted_hex))

    # --- Header (top area) ---
    if header_text:
        left = Inches(_pt_to_inches(36))
        top = Inches(_pt_to_inches(8))
        width = Inches(_pt_to_inches(sw_pt - 72))
        height = Inches(0.3)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = header_text
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(*_hex_to_rgb(muted_hex))


# ---------------------------------------------------------------------------
# 6. Slide notes
# ---------------------------------------------------------------------------

def _add_slide_notes(slide, notes_text: str) -> None:
    """Set speaker notes on a slide."""
    try:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text
    except Exception as exc:
        log.warning("Failed to add slide notes: %s", exc)


# ---------------------------------------------------------------------------
# 3. Rich text formatting
# ---------------------------------------------------------------------------

def _set_bullet_char(paragraph, bullet_char: str | bool) -> None:
    """Set a custom bullet character on a paragraph via XML manipulation.

    python-pptx does not expose a clean API for custom bullet characters,
    so we manipulate the ``pPr`` element directly.
    """
    from lxml import etree

    pPr = paragraph._p.get_or_add_pPr()

    # Enable bullet list
    buNone = pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone")
    if buNone is not None:
        pPr.remove(buNone)

    if bullet_char is True:
        # Default bullet
        buChar = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar")
        buChar.set("char", "•")
    elif isinstance(bullet_char, str) and bullet_char:
        buChar = etree.SubElement(pPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar")
        buChar.set("char", bullet_char)


def _add_text_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True

    binding = node.content_binding

    # --- Multi-paragraph text ---
    paragraphs_data = binding.get("paragraphs")
    runs_data = binding.get("runs")
    bullet = binding.get("bullet")

    if paragraphs_data and isinstance(paragraphs_data, list):
        # Warn if both paragraphs and runs are provided; runs are ignored
        if runs_data and isinstance(runs_data, list):
            log.warning(
                "Both 'paragraphs' and 'runs' provided for node %s; 'runs' will be ignored",
                node.id,
            )
        # Multiple paragraphs
        for p_idx, para_dict in enumerate(paragraphs_data):
            if p_idx == 0:
                p = tf.paragraphs[0]
                # Clear phantom default run that python-pptx creates
                for run in list(p.runs):
                    p._p.remove(run._r)
            else:
                p = tf.add_paragraph()

            para_text = para_dict.get("text", "")
            para_align = para_dict.get("align")
            para_level = para_dict.get("level", 0)
            para_bullet = para_dict.get("bullet")

            if para_align:
                p.alignment = _align_from_str(para_align)
            if para_level and isinstance(para_level, int):
                p.level = min(para_level, 8)

            # Handle bullet for this paragraph
            if para_bullet is not None and para_bullet is not False:
                _set_bullet_char(p, para_bullet)
            elif bullet is not None and bullet is not False and para_bullet is not False:
                _set_bullet_char(p, bullet)

            run = p.add_run()
            run.text = para_text
            _apply_text_style(run, node.resolved_style)

    elif runs_data and isinstance(runs_data, list):
        # Multi-run text in the first paragraph
        p = tf.paragraphs[0]
        # Clear phantom default run that python-pptx creates
        for run in list(p.runs):
            p._p.remove(run._r)
        for run_dict in runs_data:
            run = p.add_run()
            run.text = run_dict.get("text", "")
            if run_dict.get("bold"):
                run.font.bold = True
            if run_dict.get("italic"):
                run.font.italic = True
            if "color" in run_dict:
                run.font.color.rgb = RGBColor(*_hex_to_rgb(run_dict["color"]))
            if "size" in run_dict:
                run.font.size = Pt(run_dict["size"])
            if "href" in run_dict:
                href = run_dict["href"]
                if _is_safe_hyperlink(href):
                    try:
                        run.hyperlink.address = href
                    except Exception as exc:
                        log.warning("Failed to set hyperlink on run: %s", exc)
                else:
                    log.warning("Skipping unsafe hyperlink scheme: %s", href)
            # Also apply base resolved_style for any properties not overridden
            base_style = dict(node.resolved_style)
            # Don't override per-run settings
            for key in ("bold", "italic", "color", "size"):
                if key in run_dict:
                    base_style.pop(key, None)
            _apply_text_style(run, base_style)

        # Handle bullet for the paragraph
        if bullet is not None and bullet is not False:
            _set_bullet_char(p, bullet)

        # Alignment
        align = binding.get("align")
        if align:
            p.alignment = _align_from_str(align)

    else:
        # Fallback: original single-run behaviour
        text = binding.get("text", "")
        if not text and "items" in binding:
            text = "\n".join(str(item) for item in binding["items"])

        p = tf.paragraphs[0]
        # Clear phantom default run that python-pptx creates
        for run in list(p.runs):
            p._p.remove(run._r)
        run = p.add_run()
        run.text = text
        _apply_text_style(run, node.resolved_style)
        p.alignment = binding.get("align", PP_ALIGN.LEFT)

        # Handle bullet
        if bullet is not None and bullet is not False:
            _set_bullet_char(p, bullet)

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=textbox.shape_id,
        shape_name=textbox.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=None,
        parent_render_node_id=node.parent_node_id,
    )


# ---------------------------------------------------------------------------
# Image node (unchanged from PR2)
# ---------------------------------------------------------------------------

def _add_image_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    from pptx.util import Inches

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))

    path = node.content_binding.get("path", "")
    if not path or not Path(path).exists():
        # Placeholder rectangle when image is missing.
        from pptx.enum.shapes import MSO_SHAPE

        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(220, 220, 220)
        crop = None
    else:
        from pptx_skill.image_crop import crop_cover

        shape = slide.shapes.add_picture(path, left, top, width, height)
        # Apply real crop fractions so the picture covers the box without
        # overflowing neighbouring regions.
        with open(path, "rb") as fh:
            from PIL import Image

            with Image.open(fh) as img:
                src_w, src_h = img.size
        dst_ratio = geom.width / max(geom.height, 1)
        result = crop_cover(src_w, src_h, dst_ratio, 1.0)
        shape.crop_left = result.crop_fractions["left"]
        shape.crop_top = result.crop_fractions["top"]
        shape.crop_right = result.crop_fractions["right"]
        shape.crop_bottom = result.crop_fractions["bottom"]
        crop = result.crop_fractions

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=shape.shape_id,
        shape_name=shape.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=crop,
        parent_render_node_id=node.parent_node_id,
    )


# ---------------------------------------------------------------------------
# 5. Shape styling enhancements
# ---------------------------------------------------------------------------

def _add_shape_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches, Pt

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))

    shape_type = node.content_binding.get("shape_type", "rectangle")
    mso = getattr(MSO_SHAPE, shape_type.upper(), MSO_SHAPE.RECTANGLE)
    shape = slide.shapes.add_shape(mso, left, top, width, height)

    style = node.resolved_style
    if "fill" in style:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*_hex_to_rgb(style["fill"]))

    # --- Line / stroke ---
    if "line_color" in style:
        shape.line.color.rgb = RGBColor(*_hex_to_rgb(style["line_color"]))
    if "line_width" in style:
        shape.line.width = Pt(style["line_width"])

    # --- Rotation ---
    if node.geometry.rotation_deg:
        shape.rotation = node.geometry.rotation_deg

    # --- Gradient fill ---
    gradient = style.get("gradient")
    if gradient and isinstance(gradient, dict):
        try:
            from lxml import etree

            colors = gradient.get("colors", [])
            angle = gradient.get("angle", 0.0)
            if colors:
                # Remove solid fill if present, then add gradient
                sp_pr = shape._element.find(
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}spPr"
                )
                if sp_pr is None:
                    # Try the shape element directly
                    sp_pr = shape._element.find(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}spPr"
                    )
                if sp_pr is not None:
                    # Remove existing solidFill
                    for sf in sp_pr.findall(
                        "{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"
                    ):
                        sp_pr.remove(sf)
                    # Build gradient fill XML
                    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
                    grad_fill = etree.SubElement(sp_pr, f"{{{ns}}}gradFill")
                    gs_lst = etree.SubElement(grad_fill, f"{{{ns}}}gsLst")
                    for idx, hex_color in enumerate(colors):
                        pos_pct = int(idx / max(len(colors) - 1, 1) * 100000)
                        gs = etree.SubElement(gs_lst, f"{{{ns}}}gs")
                        gs.set("pos", str(pos_pct))
                        srgb = etree.SubElement(gs, f"{{{ns}}}srgbClr")
                        srgb.set("val", hex_color.lstrip("#"))
                    if angle:
                        lin = etree.SubElement(grad_fill, f"{{{ns}}}lin")
                        lin.set("ang", str(int(angle * 60000)))
                        lin.set("scaled", "1")
        except Exception as exc:
            log.warning("Failed to apply gradient fill to shape: %s", exc)

    # --- Shadow ---
    if style.get("shadow"):
        try:
            from lxml import etree

            ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
            sp_pr = shape._element.find(f"{{{ns}}}spPr")
            if sp_pr is None:
                sp_pr = shape._element.find(f".//{{{ns}}}spPr")
            if sp_pr is not None:
                effect_lst = etree.SubElement(sp_pr, f"{{{ns}}}effectLst")
                outer_shdw = etree.SubElement(effect_lst, f"{{{ns}}}outerShdw")
                outer_shdw.set("blurRad", "76200")
                outer_shdw.set("dist", "38100")
                outer_shdw.set("dir", "5400000")
                outer_shdw.set("algn", "bl")
                srgb = etree.SubElement(outer_shdw, f"{{{ns}}}srgbClr")
                srgb.set("val", "000000")
                alpha = etree.SubElement(srgb, f"{{{ns}}}alpha")
                alpha.set("val", "40000")
        except Exception as exc:
            log.warning("Failed to add shadow to shape: %s", exc)

    # --- Text inside shape ---
    binding = node.content_binding
    if binding.get("text"):
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = binding["text"]
        _apply_text_style(run, style)

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=shape.shape_id,
        shape_name=shape.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=None,
        parent_render_node_id=node.parent_node_id,
    )


# ---------------------------------------------------------------------------
# 4. Table styling enhancements
# ---------------------------------------------------------------------------

def _set_cell_border(cell, border_pt: float = 0.5) -> None:
    """Set thin borders on a table cell via XML manipulation."""
    from lxml import etree

    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    border_names = ["lnL", "lnR", "lnT", "lnB"]
    for name in border_names:
        # Remove any existing border element with the same tag to avoid duplicates
        existing = tcPr.find(f"{{{ns}}}{name}")
        if existing is not None:
            tcPr.remove(existing)
        ln = etree.SubElement(tcPr, f"{{{ns}}}{name}")
        ln.set("w", str(int(border_pt * 12700)))
        solid_fill = etree.SubElement(ln, f"{{{ns}}}solidFill")
        srgb = etree.SubElement(solid_fill, f"{{{ns}}}srgbClr")
        srgb.set("val", "CCCCCC")


def _add_table_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    """Render a table node with enhanced styling support."""
    from pptx.util import Emu, Inches, Pt

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))

    binding = node.content_binding or {}
    headers = list(binding.get("headers", []))
    rows = [list(r) for r in binding.get("rows", [])]
    all_rows = [headers] + rows if headers else rows
    n_rows = max(len(all_rows), 1)
    n_cols = max(len(headers), max((len(r) for r in rows), default=0), 1)

    table_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = table_shape.table

    style = node.resolved_style

    # --- Column widths ---
    col_widths = binding.get("col_widths")
    if col_widths and isinstance(col_widths, list):
        for c_idx, w in enumerate(col_widths):
            if c_idx < n_cols:
                try:
                    table.columns[c_idx].width = Pt(w)
                except Exception as exc:
                    log.warning("Failed to set column width %d: %s", c_idx, exc)

    # --- Cell padding (default 0.05" all sides) ---
    cell_margin = Emu(int(0.05 * 914400))

    # --- Fill cells ---
    header_fill_hex = style.get("table_header_fill")
    header_color_hex = style.get("table_header_color")
    banding = style.get("table_banding", False)
    borders = style.get("table_borders", False)
    table_align = binding.get("align")

    for r_idx, row in enumerate(all_rows):
        is_header = headers and r_idx == 0
        for c_idx in range(n_cols):
            cell_text = str(row[c_idx]) if c_idx < len(row) else ""
            cell = table.cell(r_idx, c_idx)

            # Set cell margins
            try:
                cell.margin_left = cell_margin
                cell.margin_right = cell_margin
                cell.margin_top = cell_margin
                cell.margin_bottom = cell_margin
            except Exception:
                pass

            cell.text = cell_text

            # Header row styling
            if is_header and cell_text:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                        if header_color_hex:
                            try:
                                run.font.color.rgb = RGBColor(*_hex_to_rgb(header_color_hex))
                            except Exception:
                                pass

            # Header fill
            if is_header and header_fill_hex:
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(*_hex_to_rgb(header_fill_hex))
                except Exception as exc:
                    log.warning("Failed to set header fill: %s", exc)

            # Banding: alternate row fills
            if banding and not is_header and r_idx % 2 == 0:
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(240, 240, 240)
                except Exception as exc:
                    log.warning("Failed to set banding fill: %s", exc)

            # Borders
            if borders:
                try:
                    _set_cell_border(cell)
                except Exception as exc:
                    log.warning("Failed to set cell border: %s", exc)

            # Cell text alignment
            align_val = table_align
            # Per-cell align override (if row items are dicts with align key)
            if c_idx < len(row) and isinstance(row[c_idx], dict):
                align_val = row[c_idx].get("align", align_val)
            if align_val:
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.alignment = _align_from_str(align_val)

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=table_shape.shape_id,
        shape_name=table_shape.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=None,
        parent_render_node_id=node.parent_node_id,
    )


# ---------------------------------------------------------------------------
# 2. Chart type enhancements
# ---------------------------------------------------------------------------

_CHART_TYPE_MAP: dict[str, Any] = {
    "column_clustered": "COLUMN_CLUSTERED",
    "column_stacked": "COLUMN_STACKED",
    "bar_clustered": "BAR_CLUSTERED",
    "bar_stacked": "BAR_STACKED",
    "line": "LINE",
    "line_markers": "LINE_MARKERS",
    "pie": "PIE",
    "doughnut": "DOUGHNUT",
    "scatter": "XY_SCATTER",
    "area": "AREA",
}


def _resolve_chart_type(chart_type_str: str):
    """Resolve a chart type string to an XL_CHART_TYPE enum value."""
    from pptx.enum.chart import XL_CHART_TYPE

    attr_name = _CHART_TYPE_MAP.get(chart_type_str, "COLUMN_CLUSTERED")
    return getattr(XL_CHART_TYPE, attr_name, XL_CHART_TYPE.COLUMN_CLUSTERED)


_LEGEND_POSITION_MAP: dict[str, Any] = {
    "bottom": "BOTTOM",
    "right": "RIGHT",
    "left": "LEFT",
    "top": "TOP",
}


def _build_embedded_xlsx(binding: dict, is_xy: bool) -> bytes | None:
    """Build a minimal xlsx workbook from chart binding data.

    Returns the xlsx bytes, or *None* if openpyxl is not available.
    """
    try:
        from io import BytesIO

        from openpyxl import Workbook
    except ImportError:
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    series_list = binding.get("series")

    if is_xy:
        # Scatter charts: columns are (X, Y) per series
        # Header row: "X (series-0)", "Y (series-0)", "X (series-1)", ...
        if series_list and isinstance(series_list, list):
            headers: list[str] = []
            data_rows: list[list[float | str]] = []
            max_len = 0
            for s_idx, s_dict in enumerate(series_list):
                s_name = s_dict.get("name", f"series-{s_idx}")
                headers.append(f"X ({s_name})")
                headers.append(f"Y ({s_name})")
                s_items = list(s_dict.get("items", []))
                max_len = max(max_len, len(s_items))
                for row_i in range(len(s_items)):
                    while len(data_rows) <= row_i:
                        data_rows.append([])
                    it = s_items[row_i]
                    x = float(it.get("x", it.get("value", 0)))
                    y = float(it.get("y", 0))
                    data_rows[row_i].extend([x, y])
            # Pad rows to match header count
            for row in data_rows:
                while len(row) < len(headers):
                    row.append("")
            ws.append(headers)
            for row in data_rows:
                ws.append(row)
        else:
            items = list(binding.get("items", []))
            ws.append(["X", "Y"])
            for i, it in enumerate(items):
                x = float(it.get("x", i))
                y = float(it.get("y", it.get("value", 0)))
                ws.append([x, y])
    else:
        # Category-based charts
        # Header row: "" | series-0-name | series-1-name | ...
        # Data rows:  category-label | value | value | ...
        if series_list and isinstance(series_list, list):
            headers = [""] + [
                s_dict.get("name", f"series-{s_idx}")
                for s_idx, s_dict in enumerate(series_list)
            ]
            ws.append(headers)
            first_items = list(series_list[0].get("items", []))
            for i in range(len(first_items)):
                label = str(first_items[i].get("label", f"item-{i}"))
                row = [label]
                for s_dict in series_list:
                    s_items = list(s_dict.get("items", []))
                    val = (
                        float(s_items[i].get("value", 0))
                        if i < len(s_items)
                        else 0.0
                    )
                    row.append(val)
                ws.append(row)
        else:
            items = list(binding.get("items", []))
            ws.append(["", "series"])
            for i, it in enumerate(items):
                label = str(it.get("label", f"item-{i}"))
                value = float(it.get("value", 0))
                ws.append([label, value])

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _embed_chart_workbook(chart_frame, binding: dict, is_xy: bool) -> None:
    """Embed an Excel workbook into the chart part so the chart is
    double-click-editable in PowerPoint.

    This creates a minimal xlsx from the chart data, adds it as an
    embedded-package relationship on the chart part, and updates the
    chart XML to reference the workbook.
    """
    from lxml import etree

    xlsx_bytes = _build_embedded_xlsx(binding, is_xy)
    if xlsx_bytes is None:
        log.info("openpyxl not available — skipping embedded workbook for chart")
        return

    chart_part = chart_frame.chart.part

    # Add the xlsx as an embedded package part
    xlsx_partname_str = chart_part.partname.replace("/chart", "/embeddings/chart_data")
    xlsx_partname_str = xlsx_partname_str.replace(".xml", ".xlsx")
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI

    xlsx_part = Part(
        partname=PackURI(xlsx_partname_str),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        blob=xlsx_bytes,
        package=chart_part.package,
    )

    # Create relationship from chart part to the embedded xlsx
    rel = chart_part.relate_to(
        xlsx_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    )

    # Update the chart XML to add the <c:externalData> reference.
    # The chart part's root element is <c:chartSpace>; we access it
    # via the private _element attribute (the public .element property
    # does not exist on ChartPart).
    chart_xml = chart_part._element
    ns_c = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    external_data = etree.SubElement(chart_xml, f"{{{ns_c}}}externalData")
    external_data.set(f"{{{ns_r}}}id", rel)
    # Add <c:autoUpdate val="1"/> so PowerPoint refreshes from the workbook
    etree.SubElement(
        external_data, f"{{{ns_c}}}autoUpdate"
    ).set("val", "1")


def _apply_data_labels(chart_frame, data_labels: dict) -> None:
    """Apply data-label settings to all series in the chart.

    *data_labels* schema:
        {
            "show_value": True,
            "show_category": False,
            "show_series_name": False,
            "position": "outEnd"  # optional
        }
    """
    from pptx.enum.chart import XL_LABEL_POSITION
    from pptx.util import Pt

    _LABEL_POSITION_MAP: dict[str, str] = {
        "outend": "OUTSIDE_END",
        "outsideend": "OUTSIDE_END",
        "inend": "INSIDE_END",
        "insideend": "INSIDE_END",
        "inbase": "INSIDE_BASE",
        "insidebase": "INSIDE_BASE",
        "center": "CENTER",
        "above": "ABOVE",
        "below": "BELOW",
        "left": "LEFT",
        "right": "RIGHT",
        "bestfit": "BEST_FIT",
    }

    position_str = (data_labels.get("position") or "outEnd").lower().replace("_", "")
    attr_name = _LABEL_POSITION_MAP.get(position_str, "OUTSIDE_END")
    position_enum = getattr(XL_LABEL_POSITION, attr_name, XL_LABEL_POSITION.OUTSIDE_END)

    chart = chart_frame.chart
    show_value = data_labels.get("show_value", True)
    show_category = data_labels.get("show_category", False)
    show_series_name = data_labels.get("show_series_name", False)

    try:
        plot = chart.plots[0]
        plot.has_data_labels = True
        data_labels_obj = plot.data_labels
        data_labels_obj.show_value = bool(show_value)
        data_labels_obj.show_category_name = bool(show_category)
        data_labels_obj.show_series_name = bool(show_series_name)
        try:
            data_labels_obj.number_format = "General"
        except Exception:
            pass
        try:
            data_labels_obj.font.size = Pt(9)
        except Exception:
            pass
        try:
            data_labels_obj.position = position_enum
        except Exception:
            # Some positions are invalid for certain chart types; ignore
            pass
    except Exception as exc:
        log.warning("Failed to apply data labels: %s", exc)


def _apply_insight_annotations(chart_frame, annotations: list[dict]) -> None:
    """Add callout text annotations to specific data points.

    *annotations* schema:
        [
            {"text": "▲ 23% YoY", "point_index": 0, "series_index": 0},
            ...
        ]
    If *series_index* is omitted the first series is used.

    This is implemented by setting the data-label text on the target point
    to the annotation text, which creates a visible callout in PowerPoint.
    """
    from lxml import etree

    if not annotations:
        return

    chart = chart_frame.chart
    nsmap = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
             "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

    try:
        plot = chart.plots[0]
    except (IndexError, Exception):
        log.warning("Cannot apply insight annotations — no plot found")
        return

    for ann in annotations:
        text = ann.get("text", "")
        if not text:
            continue
        point_idx = int(ann.get("point_index", 0))
        series_idx = int(ann.get("series_index", 0))

        try:
            series = plot.series[series_idx]
        except (IndexError, Exception):
            log.warning(
                "insight_annotations: series_index %d out of range", series_idx
            )
            continue

        # Enable data labels on the point via python-pptx if possible,
        # then patch the XML to set custom text.
        try:
            point = series.points[point_idx]
            point.has_data_label = True
            # Patch the XML: set custom text on the data label
            # python-pptx creates <c:dLbl> elements; we need to ensure
            # a <c:tx> with our annotation text exists.
            dlbl = point.data_label
            # Access the underlying XML element
            dlbl_elem = dlbl._element
            # Remove any existing <c:tx> to replace it
            for existing_tx in dlbl_elem.findall("{{{}}}tx".format(nsmap["c"])):
                dlbl_elem.remove(existing_tx)
            # Build <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>TEXT</a:t></a:r></a:p></c:rich></c:tx>
            tx = etree.SubElement(dlbl_elem, "{{{}}}tx".format(nsmap["c"]))
            rich = etree.SubElement(tx, "{{{}}}rich".format(nsmap["c"]))
            etree.SubElement(rich, "{{{}}}bodyPr".format(nsmap["a"]))
            etree.SubElement(rich, "{{{}}}lstStyle".format(nsmap["a"]))
            p = etree.SubElement(rich, "{{{}}}p".format(nsmap["a"]))
            pPr = etree.SubElement(p, "{{{}}}pPr".format(nsmap["a"]))
            etree.SubElement(pPr, "{{{}}}defRPr".format(nsmap["a"])).set("sz", "900")
            r = etree.SubElement(p, "{{{}}}r".format(nsmap["a"]))
            rPr = etree.SubElement(r, "{{{}}}rPr".format(nsmap["a"]))
            rPr.set("lang", "en-US")
            rPr.set("sz", "900")
            rPr.set("dirty", "0")
            t = etree.SubElement(r, "{{{}}}t".format(nsmap["a"]))
            t.text = text
            # Ensure <c:showVal> etc. are not overriding our custom text
            for tag in ("showVal", "showCatName", "showSerName", "showPercent"):
                elem = dlbl_elem.find("{{{}}}{}".format(nsmap["c"], tag))
                if elem is not None:
                    dlbl_elem.remove(elem)
        except Exception as exc:
            log.warning(
                "insight_annotations: failed on point %d series %d: %s",
                point_idx, series_idx, exc,
            )


def _add_chart_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    """Render a chart node with multiple chart types and multi-series support.

    Enhanced to support:
    - Embedded Excel workbook for double-click-editability
    - Data labels (show_value, show_category, show_series_name, position)
    - Insight annotations (callout text on specific data points)
    """
    from pptx.chart.data import CategoryChartData, XyChartData
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.util import Inches

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))

    binding = node.content_binding or {}
    style = node.resolved_style

    # Resolve chart type
    chart_type_str = binding.get("chart_type", "column_clustered")
    chart_type = _resolve_chart_type(chart_type_str)

    # Scatter charts require XyChartData instead of CategoryChartData
    is_xy = chart_type_str in ("scatter",)

    if is_xy:
        chart_data = XyChartData()
        series_list = binding.get("series")
        if series_list and isinstance(series_list, list):
            for s_idx, s_dict in enumerate(series_list):
                s_name = s_dict.get("name", f"series-{s_idx}")
                s_items = list(s_dict.get("items", []))
                series = chart_data.add_series(s_name)
                for it in s_items:
                    x = float(it.get("x", it.get("value", 0)))
                    y = float(it.get("y", 0))
                    series.add_data_point(x, y)
        else:
            items = list(binding.get("items", []))
            series = chart_data.add_series("series")
            for i, it in enumerate(items):
                x = float(it.get("x", i))
                y = float(it.get("y", it.get("value", 0)))
                series.add_data_point(x, y)
    else:
        # Build chart data for category-based charts
        chart_data = CategoryChartData()

        # Multi-series support
        series_list = binding.get("series")
        if series_list and isinstance(series_list, list):
            # Collect all category labels from the first series
            first_series = series_list[0] if series_list else {}
            items = list(first_series.get("items", []))
            labels = [str(it.get("label", f"item-{i}")) for i, it in enumerate(items)]
            chart_data.categories = labels or ["item"]

            for s_idx, s_dict in enumerate(series_list):
                s_name = s_dict.get("name", f"series-{s_idx}")
                s_items = list(s_dict.get("items", []))
                s_values = [float(it.get("value", 0)) for it in s_items] or [0.0]
                chart_data.add_series(s_name, s_values)
        else:
            # Single-series fallback (original behaviour)
            items = list(binding.get("items", []))
            labels = [str(it.get("label", f"item-{i}")) for i, it in enumerate(items)]
            values = [float(it.get("value", 0)) for it in items] or [0.0]
            chart_data.categories = labels or ["item"]
            chart_data.add_series("series", values)

    chart_frame = slide.shapes.add_chart(
        chart_type, left, top, width, height, chart_data
    )

    # --- Embedded Excel workbook for double-click-editability ---
    try:
        _embed_chart_workbook(chart_frame, binding, is_xy)
    except Exception as exc:
        log.warning("Failed to embed chart workbook (chart still works): %s", exc)

    # --- Chart styling ---
    # --- Chart style (int 1-48) ---
    chart_style_val = style.get("chart_style", 2)
    if isinstance(chart_style_val, int) and 1 <= chart_style_val <= 48:
        try:
            chart_frame.chart.style = chart_style_val
        except Exception as exc:
            log.warning("Failed to set chart style: %s", exc)

    # --- Legend ---
    has_legend = style.get("has_legend")
    if has_legend is None:
        # Default: show legend if multi-series
        has_legend = bool(series_list)
    if has_legend:
        try:
            chart_frame.chart.has_legend = True
            legend_pos_str = style.get("legend_position", "bottom")
            attr_name = _LEGEND_POSITION_MAP.get(legend_pos_str, "BOTTOM")
            chart_frame.chart.legend.position = getattr(
                XL_LEGEND_POSITION, attr_name, XL_LEGEND_POSITION.BOTTOM
            )
        except Exception as exc:
            log.warning("Failed to set chart legend: %s", exc)
    else:
        try:
            chart_frame.chart.has_legend = False
        except Exception:
            pass

    # --- Data labels ---
    data_labels = binding.get("data_labels") or style.get("data_labels")
    if data_labels and isinstance(data_labels, dict):
        try:
            _apply_data_labels(chart_frame, data_labels)
        except Exception as exc:
            log.warning("Failed to apply data labels: %s", exc)

    # --- Insight annotations ---
    insight_annotations = binding.get("insight_annotations")
    if insight_annotations and isinstance(insight_annotations, list):
        try:
            _apply_insight_annotations(chart_frame, insight_annotations)
        except Exception as exc:
            log.warning("Failed to apply insight annotations: %s", exc)

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=chart_frame.shape_id,
        shape_name=chart_frame.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=None,
        parent_render_node_id=node.parent_node_id,
    )


# ---------------------------------------------------------------------------
# 7. Video / Audio media support
# ---------------------------------------------------------------------------

_VIDEO_MIME_MAP: dict[str, str] = {
    ".mp4": "video/mp4",
    ".avi": "video/avi",
    ".mov": "video/quicktime",
    ".wmv": "video/x-ms-wmv",
    ".mkv": "video/x-matroska",
    ".webm": "video/webm",
}


def _resolve_video_mime(video_path: str) -> str:
    """Resolve a video MIME type from the file extension."""
    ext = Path(video_path).suffix.lower()
    return _VIDEO_MIME_MAP.get(ext, "video/mp4")


def _add_video_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    """Render a video node using slide.shapes.add_movie()."""
    from pptx.util import Inches, Pt

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))

    binding = node.content_binding or {}
    video_path = binding.get("path", "")
    poster_path = binding.get("poster")

    if not video_path or not Path(video_path).exists():
        # Fallback: placeholder rectangle
        from pptx.enum.shapes import MSO_SHAPE

        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(60, 60, 60)
        # Add a play icon text
        tf = shape.text_frame
        p = tf.paragraphs[0]
        p.alignment = _align_from_str("center")
        run = p.add_run()
        run.text = "▶ Video"
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(255, 255, 255)
    else:
        try:
            poster_image = poster_path if (poster_path and Path(poster_path).exists()) else None
            mime_type = binding.get("mime_type") or _resolve_video_mime(video_path)
            shape = slide.shapes.add_movie(
                video_path, left, top, width, height,
                poster_frame_image=poster_image,
                mime_type=mime_type,
            )
        except Exception as exc:
            log.warning("add_movie failed for %s, using placeholder: %s", video_path, exc)
            from pptx.enum.shapes import MSO_SHAPE

            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, left, top, width, height
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(60, 60, 60)
            tf = shape.text_frame
            p = tf.paragraphs[0]
            p.alignment = _align_from_str("center")
            run = p.add_run()
            run.text = "▶ Video"
            run.font.color.rgb = RGBColor(255, 255, 255)

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=shape.shape_id,
        shape_name=shape.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=None,
        parent_render_node_id=node.parent_node_id,
    )


_AUDIO_MIME_MAP: dict[str, str] = {
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".ogg": "audio/ogg",
    ".m4a": "audio/mp4",
    ".wma": "audio/x-ms-wma",
    ".aac": "audio/aac",
    ".flac": "audio/flac",
}


def _resolve_audio_mime(audio_path: str) -> str:
    """Resolve an audio MIME type from the file extension."""
    ext = Path(audio_path).suffix.lower()
    return _AUDIO_MIME_MAP.get(ext, "audio/mpeg")


def _add_audio_node(slide, node: PlannedNode, shape_index: int) -> RenderTraceEntry:
    """Render an audio node using add_movie with audio MIME type.

    python-pptx does not have a dedicated add_audio API, but add_movie
    supports embedding media files including audio when given an audio
    MIME type. Falls back to a placeholder shape on failure.
    """
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(max(geom.width, 1.0)))
    height = Inches(_pt_to_inches(max(geom.height, 1.0)))

    binding = node.content_binding or {}
    path = binding.get("path", "")

    if not path or not Path(path).exists():
        # Placeholder shape when audio file is missing
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(240, 240, 240)
        shape.line.fill.background()
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = _align_from_str("center")
        run = p.add_run()
        run.text = "♪ Audio"
        run.font.color.rgb = RGBColor(100, 100, 100)
    else:
        try:
            mime = binding.get("mime_type") or _resolve_audio_mime(path)
            shape = slide.shapes.add_movie(
                path, left, top, width, height, mime_type=mime
            )
        except Exception as exc:
            log.warning("Failed to embed audio %s: %s", path, exc)
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(240, 240, 240)
            shape.line.fill.background()
            tf = shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = _align_from_str("center")
            run = p.add_run()
            run.text = "♪ Audio"
            run.font.color.rgb = RGBColor(100, 100, 100)

    return RenderTraceEntry(
        render_node_id=node.id,
        element_id=node.element_id,
        recipe_node_id=node.recipe_node_id,
        shape_instance_index=shape_index,
        slide_index=0,
        ppt_shape_id=shape.shape_id,
        shape_name=shape.name,
        z_order=node.z_order,
        geometry=node.geometry,
        crop=None,
        parent_render_node_id=node.parent_node_id,
    )


# ---------------------------------------------------------------------------
# Main render entry points
# ---------------------------------------------------------------------------

def render_layout_plan(plan: LayoutPlan, output_path: str) -> RenderResult:
    """Render a single-slide LayoutPlan to a PPTX file.

    This is the PR2 minimal adaptive renderer: it supports text, image and
    shape nodes, applies real picture crop fractions, and returns a trace.
    """
    return render_layout_plans([plan], output_path)


def render_layout_plans(
    plans: list[LayoutPlan],
    output_path: str,
    deck_options: dict[str, Any] | None = None,
) -> RenderResult:
    """Render a list of LayoutPlans to a single multi-slide PPTX file.

    Each plan becomes one slide in the output deck. Slide index is recorded
    in every RenderTraceEntry so QA can map issues back to slides.

    Args:
        plans: List of LayoutPlan objects, one per slide.
        output_path: Filesystem path for the output .pptx file.
        deck_options: Optional dict controlling deck-level features:
            - show_slide_numbers (bool, default True)
            - footer_text (str, default "")
            - header_text (str, default "")
            - slide_number_format (str, default "{current}/{total}")
            - theme (dict, default None) — may contain text_muted, text keys
    """
    from pathlib import Path

    from pptx import Presentation
    from pptx.util import Inches

    if not plans:
        raise AdaptiveRendererError("Cannot render an empty plan list")

    # Merge deck_options with defaults
    opts = dict(_DEFAULT_DECK_OPTIONS)
    if deck_options:
        opts.update(deck_options)

    prs = Presentation()
    first_canvas = plans[0].canvas
    prs.slide_width = Inches(_pt_to_inches(first_canvas.width_pt))
    prs.slide_height = Inches(_pt_to_inches(first_canvas.height_pt))

    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]

    total_slides = len(plans)
    trace: list[RenderTraceEntry] = []
    for slide_index, plan in enumerate(plans):
        slide = prs.slides.add_slide(blank_layout)
        for i, node in enumerate(sorted(plan.nodes, key=lambda n: n.z_order)):
            try:
                if node.kind == "text":
                    entry = _add_text_node(slide, node, i)
                elif node.kind == "image":
                    entry = _add_image_node(slide, node, i)
                elif node.kind == "shape":
                    entry = _add_shape_node(slide, node, i)
                elif node.kind == "table":
                    entry = _add_table_node(slide, node, i)
                elif node.kind == "chart":
                    entry = _add_chart_node(slide, node, i)
                elif node.kind == "video":
                    entry = _add_video_node(slide, node, i)
                elif node.kind == "audio":
                    entry = _add_audio_node(slide, node, i)
                else:
                    continue
            except (TypeError, AttributeError, KeyError) as exc:
                logging.getLogger(__name__).error("Critical error rendering node %s: %s", node.id, exc, exc_info=True)
                raise
            except Exception as exc:
                logging.getLogger(__name__).warning("Skipping failed node %s: %s", node.id, exc)
                continue
            # Override slide_index on the entry to reflect this slide.
            trace.append(
                RenderTraceEntry(
                    render_node_id=entry.render_node_id,
                    element_id=entry.element_id,
                    recipe_node_id=entry.recipe_node_id,
                    shape_instance_index=entry.shape_instance_index,
                    slide_index=slide_index,
                    ppt_shape_id=entry.ppt_shape_id,
                    shape_name=entry.shape_name,
                    z_order=entry.z_order,
                    geometry=entry.geometry,
                    crop=entry.crop,
                    parent_render_node_id=entry.parent_render_node_id,
                )
            )

        # Add slide meta (headers / footers / slide numbers)
        try:
            _add_slide_meta(slide, slide_index, total_slides, opts,
                            prs.slide_width, prs.slide_height)
        except Exception as exc:
            log.warning("Failed to add slide meta for slide %d: %s", slide_index, exc)

        # Add slide notes from diagnostics
        try:
            for diag in plan.diagnostics:
                if isinstance(diag, dict) and "__slide_notes" in diag:
                    _add_slide_notes(slide, str(diag["__slide_notes"]))
                    break  # Only use the first notes entry
        except Exception as exc:
            log.warning("Failed to add slide notes for slide %d: %s", slide_index, exc)

    # Apply deck-level transitions if requested
    transition_cfg = opts.get("transition")
    if transition_cfg:
        from pptx_skill.transitions import TRANSITION_TYPES, apply_slide_transition

        if isinstance(transition_cfg, str):
            # Simple string: just the transition type
            t_type = transition_cfg if transition_cfg in TRANSITION_TYPES else "fade"
            for slide in prs.slides:
                apply_slide_transition(slide, t_type, 700)
        elif isinstance(transition_cfg, dict):
            # Dict with detailed options: type, duration_ms, advance_ms
            t_type = transition_cfg.get("type", "fade")
            if t_type not in TRANSITION_TYPES:
                log.warning("Unknown transition type %r in deck_options; falling back to 'fade'", t_type)
                t_type = "fade"
            t_duration = transition_cfg.get("duration_ms", 700)
            t_advance = transition_cfg.get("advance_ms")
            for slide in prs.slides:
                apply_slide_transition(slide, t_type, t_duration, t_advance)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)

    return RenderResult(
        pptx_path=output_path,
        trace=trace,
        generation_engine="pptx_skill.adaptive_renderer.v2",
        generation_engine_version="2.0.0-pr8a",
        artifacts={},
    )
