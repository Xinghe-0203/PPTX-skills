"""Renderer result models and the eventual adaptive PPTX renderer.

At PR1 this module only defines the data contract for RenderTrace and render
results. The actual OOXML writing will be added in PR2.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pptx_skill.content_model import CanvasSpec, GeometrySpec


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
