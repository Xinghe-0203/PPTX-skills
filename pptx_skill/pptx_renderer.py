"""Renderer result models and the eventual adaptive PPTX renderer.

At PR1 this module only defines the data contract for RenderTrace and render
results. The actual OOXML writing will be added in PR2.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx.dml.color import RGBColor

from pptx_skill.content_model import CanvasSpec, GeometrySpec, LayoutPlan, PlannedNode


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
# PR2 adaptive renderer (basic)
# ---------------------------------------------------------------------------

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


def _apply_text_style(run, style: dict[str, Any]) -> None:
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    if "size" in style:
        run.font.size = Pt(style["size"])
    if "color" in style:
        run.font.color.rgb = RGBColor(*_hex_to_rgb(style["color"]))
    if style.get("bold"):
        run.font.bold = True


def _add_text_node(slide, node: "PlannedNode", shape_index: int) -> "RenderTraceEntry":
    from pptx.util import Inches, Pt

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(geom.width))
    height = Inches(_pt_to_inches(geom.height))
    textbox = slide.shapes.add_textbox(left, top, width, height)
    tf = textbox.text_frame
    tf.word_wrap = True

    binding = node.content_binding
    text = binding.get("text", "")
    if not text and "items" in binding:
        text = "\n".join(str(item) for item in binding["items"])

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    _apply_text_style(run, node.resolved_style)
    p.alignment = binding.get("align", 1)  # 1 = center, 2 = right, 0 = left

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


def _add_image_node(slide, node: "PlannedNode", shape_index: int) -> "RenderTraceEntry":
    from pptx.util import Inches

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(geom.width))
    height = Inches(_pt_to_inches(geom.height))

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

            img = Image.open(fh)
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


def _add_shape_node(slide, node: "PlannedNode", shape_index: int) -> "RenderTraceEntry":
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    geom = node.geometry.bbox
    left = Inches(_pt_to_inches(geom.left))
    top = Inches(_pt_to_inches(geom.top))
    width = Inches(_pt_to_inches(geom.width))
    height = Inches(_pt_to_inches(geom.height))

    shape_type = node.content_binding.get("shape_type", "rectangle")
    mso = getattr(MSO_SHAPE, shape_type.upper(), MSO_SHAPE.RECTANGLE)
    shape = slide.shapes.add_shape(mso, left, top, width, height)

    style = node.resolved_style
    if "fill" in style:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*_hex_to_rgb(style["fill"]))

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


def render_layout_plan(plan: "LayoutPlan", output_path: str) -> RenderResult:
    """Render a single-slide LayoutPlan to a PPTX file.

    This is the PR2 minimal adaptive renderer: it supports text, image and
    shape nodes, applies real picture crop fractions, and returns a trace.
    """
    from pathlib import Path

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    canvas = plan.canvas
    prs.slide_width = Inches(_pt_to_inches(canvas.width_pt))
    prs.slide_height = Inches(_pt_to_inches(canvas.height_pt))

    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank_layout)

    trace: list[RenderTraceEntry] = []
    for i, node in enumerate(sorted(plan.nodes, key=lambda n: n.z_order)):
        try:
            if node.kind == "text":
                entry = _add_text_node(slide, node, i)
            elif node.kind == "image":
                entry = _add_image_node(slide, node, i)
            elif node.kind == "shape":
                entry = _add_shape_node(slide, node, i)
            else:
                continue
        except Exception as exc:
            raise AdaptiveRendererError(
                f"Failed to render node {node.id}: {exc}"
            ) from exc
        trace.append(entry)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)

    return RenderResult(
        pptx_path=output_path,
        trace=trace,
        generation_engine="pptx_skill.adaptive_renderer.v2",
        generation_engine_version="2.0.0-pr2",
        artifacts={},
    )
