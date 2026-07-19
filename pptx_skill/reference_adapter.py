"""Reference-deck adapters: native and clone modes (PR8b).

Two reference modes per blueprint §9.3:

- ``native``: drop source slides, build new ones from the reference file's own
  masters/layouts and fill placeholders. Only fit content *inside* placeholders;
  never alter master geometry. Canvas comes from the reference file's own
  ``slide_width``/``slide_height`` (not a 16:9 constant).
- ``clone``: pick an exemplar slide by role, clone it, and replace only the
  explicitly-mapped shapes. Geometry re-layout is OFF by default — unmapped
  shapes must not drift. Returns a drift report so QA can verify invariance.

The underlying clone/bind machinery lives in ``scripts/reference_ppt.py``; this
module is the clean ``pptx_skill`` adapter boundary that consumes a
``ContentSpec`` and emits a saved PPTX plus a structured result.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from pptx_skill.content_model import CanvasSpec, ContentSpec, SlideSpec, canvas_from_name
from pptx_skill.layout_engine import builtin_recipes, solve_recipe, solved_geometry_to_layout_plan, _builtin_tokens

# These helpers live in scripts/reference_ppt.py (the legacy module). The
# package __init__ prepends scripts/ to sys.path, so the bare import works both
# from tests and from consumers.
try:
    from reference_ppt import (  # type: ignore[import-not-found]
        analyze_presentation,
        auto_bind_slide,
        clone_slide,
        compose_from_reference,
        generate_native_from_reference,
    )
except Exception:  # pragma: no cover - import guard for environments without scripts/
    analyze_presentation = None  # type: ignore[assignment]
    auto_bind_slide = None  # type: ignore[assignment]
    clone_slide = None  # type: ignore[assignment]
    compose_from_reference = None  # type: ignore[assignment]
    generate_native_from_reference = None  # type: ignore[assignment]


REFERENCE_MODES = ("native", "clone", "visual-rebuild")


@dataclass
class ReferenceCanvasInfo:
    """The reference file's own canvas dimensions, in EMU and points."""

    width_emu: int
    height_emu: int
    width_pt: float
    height_pt: float
    aspect_ratio: float

    def to_canvas_spec(self, safe_margin_pt: float = 36.0) -> CanvasSpec:
        from pptx_skill.content_model import SafeInsets

        return CanvasSpec(
            width_pt=self.width_pt,
            height_pt=self.height_pt,
            safe=SafeInsets(
                top=safe_margin_pt,
                right=safe_margin_pt,
                bottom=safe_margin_pt,
                left=safe_margin_pt,
            ),
        )


def reference_canvas(reference_pptx: str | Path) -> ReferenceCanvasInfo:
    """Read the reference file's own slide size (blueprint: canvas from the
    reference file, not a 16:9 constant)."""
    from pptx import Presentation

    prs = Presentation(str(reference_pptx))
    w_emu = int(prs.slide_width)
    h_emu = int(prs.slide_height)
    return ReferenceCanvasInfo(
        width_emu=w_emu,
        height_emu=h_emu,
        width_pt=round(w_emu / 12700.0, 3),
        height_pt=round(h_emu / 12700.0, 3),
        aspect_ratio=round(w_emu / h_emu, 4) if h_emu else 0.0,
    )


@dataclass
class ReferenceAdapterResult:
    output_path: str
    mode: str
    canvas: ReferenceCanvasInfo
    slides_built: int
    unmapped_shape_drift: list[dict] = field(default_factory=list)
    diagnostics: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": self.output_path,
            "mode": self.mode,
            "canvas": asdict(self.canvas),
            "slides_built": self.slides_built,
            "unmapped_shape_drift": self.unmapped_shape_drift,
            "diagnostics": self.diagnostics,
        }


def _slide_to_section_dict(slide: SlideSpec, title: str | None = None) -> dict[str, Any]:
    """Flatten a SlideSpec into the dict shape the legacy reference helpers expect."""
    section: dict[str, Any] = {"title": title or slide.communication_goal or ""}
    bullets: list[str] = []
    images: list[str] = []
    table_headers: list[str] = []
    table_rows: list[list[str]] = []
    metrics: list[dict[str, Any]] = []
    events: list[dict[str, Any]] = []
    steps: list[str] = []
    left: dict[str, Any] = {}
    right: dict[str, Any] = {}
    quote: str | None = None
    source: str | None = None

    for element in slide.elements:
        kind = element.kind
        content = element.content or {}
        if kind == "text":
            text = str(content.get("text", ""))
            if element.role == "title":
                section["title"] = text or section["title"]
            elif element.role == "subtitle":
                section["subtitle"] = text
            else:
                for line in text.split("\n"):
                    if line.strip():
                        bullets.append(line.strip())
        elif kind == "image":
            path = str(content.get("path", ""))
            if path:
                images.append(path)
        elif kind == "table":
            table_headers = [str(h) for h in content.get("headers", [])]
            table_rows = [[str(c) for c in row] for row in content.get("rows", [])]
        elif kind == "chart":
            metrics.extend(content.get("items", []))
        elif kind == "shape":
            items = content.get("items", [])
            if slide.role == "timeline":
                events.extend(items)
            elif slide.role == "process":
                steps.extend(str(i.get("label", i)) if isinstance(i, dict) else str(i) for i in items)
        elif kind == "group" and element.role == "comparison":
            left = content.get("left", {})
            right = content.get("right", {})

    # Role-specific content extraction.
    if slide.role == "quote":
        for element in slide.elements:
            if element.role == "body":
                quote = str(element.content.get("text", ""))
                break
        for element in slide.elements:
            if element.role == "attribution":
                source = str(element.content.get("text", ""))
                break

    if bullets:
        section["bullets"] = bullets
    if images:
        section["images"] = images
    if table_headers or table_rows:
        section["table_headers"] = table_headers
        section["table_rows"] = table_rows
    if metrics:
        section["metrics"] = metrics
    if events:
        section["events"] = events
    if steps:
        section["steps"] = steps
    if left or right:
        section["left"] = left
        section["right"] = right
    if quote is not None:
        section["quote"] = quote
        if source is not None:
            section["source"] = source
    return section


def _content_spec_to_sections(content: ContentSpec) -> tuple[str, str, list[dict]]:
    """Flatten a ContentSpec into (title, subtitle, sections[]) for legacy helpers."""
    sections = []
    for slide in content.slides:
        # Skip cover/end roles — the legacy generator adds its own bookends.
        if slide.role in {"cover", "end"}:
            continue
        sections.append(_slide_to_section_dict(slide))
    return content.title, content.subtitle, sections


def native_mode_adapter(
    reference_pptx: str | Path,
    content: ContentSpec,
    output_path: str | Path,
) -> ReferenceAdapterResult:
    """Build a deck using the reference file's own masters/layouts.

    Only fills placeholders; never alters master geometry. Canvas is read from
    the reference file (blueprint §9.3 native mode).
    """
    if generate_native_from_reference is None:  # pragma: no cover
        raise RuntimeError("reference_ppt helpers not available; scripts/ not importable")

    canvas = reference_canvas(reference_pptx)
    title, subtitle, sections = _content_spec_to_sections(content)
    saved = generate_native_from_reference(
        reference_pptx, title, sections, output_path, subtitle=subtitle,
    )
    return ReferenceAdapterResult(
        output_path=saved,
        mode="native",
        canvas=canvas,
        slides_built=len(sections) + 2,  # +cover +end bookends added by legacy
        diagnostics={"source_slides_removed": True},
    )


def clone_mode_adapter(
    reference_pptx: str | Path,
    content: ContentSpec,
    output_path: str | Path,
    clear_unmapped_content: bool = False,
) -> ReferenceAdapterResult:
    """Clone exemplar slides by role and replace explicitly-mapped content.

    Geometry re-layout is OFF — unmapped shapes must not drift. When
    ``clear_unmapped_content`` is True the legacy helper clears leftover text in
    unmapped placeholders; this is opt-in so unmapped geometry is preserved by
    default per blueprint §9.3 clone mode.
    """
    if compose_from_reference is None or analyze_presentation is None:  # pragma: no cover
        raise RuntimeError("reference_ppt helpers not available; scripts/ not importable")

    canvas = reference_canvas(reference_pptx)
    analysis = analyze_presentation(reference_pptx)
    title, subtitle, sections = _content_spec_to_sections(content)

    # Build a plan: one clone per section, choosing an exemplar by role.
    # Bookend cover/end from the reference's own cover/end slides when available.
    items = _build_clone_plan(analysis, title, subtitle, sections)

    plan = {
        "slides": [
            {
                "source_slide": item["source_slide"],
                "content": item["content"],
                "clear_unmapped_content": clear_unmapped_content,
            }
            for item in items
        ],
        "remove_source_slides": True,
    }
    saved = compose_from_reference(reference_pptx, plan, output_path)

    # QA sidecar: verify unmapped shapes on cloned slides did not drift in
    # position/size relative to their source exemplar.
    drift = _verify_clone_drift(reference_pptx, analysis, items, saved)

    return ReferenceAdapterResult(
        output_path=saved,
        mode="clone",
        canvas=canvas,
        slides_built=len(items),
        unmapped_shape_drift=drift,
        diagnostics={
            "clear_unmapped_content": clear_unmapped_content,
            "recommended_mode": analysis.get("recommended_mode"),
        },
    )


def _build_clone_plan(
    analysis: dict, title: str, subtitle: str, sections: list[dict]
) -> list[dict]:
    """Pick an exemplar source slide for each output item by role."""
    total = len(sections) + 2  # cover + content + end
    items: list[dict] = []

    # Cover: prefer a reference slide classified as cover.
    cover_idx = _exemplar_for_role(analysis, "cover")
    items.append({"source_slide": cover_idx, "content": {"title": title, "subtitle": subtitle}})

    for position, section in enumerate(sections, start=1):
        role = _role_for_section(section, position, total)
        source = _exemplar_for_role(analysis, role)
        items.append({"source_slide": source, "content": section, "role": role})

    # End: prefer a reference end slide.
    end_idx = _exemplar_for_role(analysis, "end")
    items.append({"source_slide": end_idx, "content": {"title": "谢谢", "subtitle": subtitle}})
    return items


def _role_for_section(section: dict, position: int, total: int) -> str:
    if section.get("metrics") or section.get("table_headers"):
        return "dashboard"
    if section.get("events"):
        return "timeline"
    if section.get("left") and section.get("right"):
        return "comparison"
    if section.get("steps"):
        return "process"
    if section.get("quote"):
        return "quote"
    if len(section.get("images", [])) >= 3:
        return "image_grid"
    if section.get("images"):
        return "text_image"
    return "bullets"


def _exemplar_for_role(analysis: dict, role: str) -> int:
    """1-based index of a reference slide matching the role, else slide 1."""
    matching = [s["index"] for s in analysis.get("slides", []) if s.get("role") == role]
    if matching:
        return matching[0]
    # Fall back to the first slide so cloning always succeeds.
    return analysis.get("slides", [{}])[0].get("index", 1) if analysis.get("slides") else 1


def _verify_clone_drift(
    reference_pptx: str | Path,
    analysis: dict,
    items: list[dict],
    output_path: str | Path,
) -> list[dict]:
    """Compare cloned-slide shapes against their source exemplar's geometry.

    Returns a list of drift records for unmapped shapes whose position or size
    changed. Mapped (replaced) shapes are excluded since their text is expected
    to change (but not geometry). An empty list means no drift — the clone
    invariance holds.
    """
    from pptx import Presentation

    drift: list[dict] = []
    try:
        out_prs = Presentation(str(output_path))
    except Exception:  # pragma: no cover - corrupt output path
        return drift
    out_slides = list(out_prs.slides)
    # Build a lookup of source exemplar geometry keyed by source index.
    source_by_index: dict[int, dict] = {s["index"]: s for s in analysis.get("slides", [])}

    for item_index, item in enumerate(items):
        source_index = item["source_slide"]
        if item_index >= len(out_slides):
            break
        out_slide = out_slides[item_index]
        source_record = source_by_index.get(source_index)
        if not source_record:
            continue
        source_shapes = {r["shape_id"]: r for r in source_record.get("shapes", [])}
        for out_shape in out_slide.shapes:
            sid = out_shape.shape_id
            src = source_shapes.get(sid)
            if src is None:
                continue
            # Compare position/size in EMU; ignore pure-text changes.
            src_left = src.get("left")
            src_top = src.get("top")
            src_w = src.get("width")
            src_h = src.get("height")
            if None in (src_left, src_top, src_w, src_h):
                continue
            if (out_shape.left, out_shape.top, out_shape.width, out_shape.height) != (
                src_left, src_top, src_w, src_h
            ):
                drift.append({
                    "output_slide": item_index + 1,
                    "shape_id": sid,
                    "shape_name": out_shape.name,
                    "field": "geometry",
                    "source": {"left": src_left, "top": src_top, "width": src_w, "height": src_h},
                    "output": {
                        "left": out_shape.left,
                        "top": out_shape.top,
                        "width": out_shape.width,
                        "height": out_shape.height,
                    },
                })
    return drift


def generate_from_reference_adapter(
    reference_pptx: str | Path,
    output_path: str | Path,
    content: ContentSpec | None = None,
    mode: str = "auto",
) -> ReferenceAdapterResult:
    """Dispatch to native, clone, or visual-rebuild mode.

    ``auto`` uses the analysis recommendation. ``visual-rebuild`` routes to the
    real adaptive rebuild path (PR8c) via :mod:`pptx_skill.visual_rebuild` — it
    no longer aliases clone. ``content`` is optional for visual-rebuild since
    that mode derives its content from the reference deck itself.
    """
    if mode == "auto":
        if analyze_presentation is None:  # pragma: no cover
            raise RuntimeError("reference_ppt helpers not available")
        analysis = analyze_presentation(reference_pptx)
        mode = analysis.get("recommended_mode", "clone")
    if mode == "native":
        if content is None:
            raise ValueError("native mode requires a ContentSpec")
        return native_mode_adapter(reference_pptx, content, output_path)
    if mode == "clone":
        if content is None:
            raise ValueError("clone mode requires a ContentSpec")
        return clone_mode_adapter(reference_pptx, content, output_path)
    if mode == "visual-rebuild":
        # Real adaptive rebuild path (PR8c). Derives content from the reference.
        from pptx_skill.visual_rebuild import visual_rebuild_adapter

        vb = visual_rebuild_adapter(reference_pptx, output_path)
        return ReferenceAdapterResult(
            output_path=vb.output_path,
            mode="visual-rebuild",
            canvas=ReferenceCanvasInfo(
                width_emu=int(vb.canvas.width_pt * 12700),
                height_emu=int(vb.canvas.height_pt * 12700),
                width_pt=vb.canvas.width_pt,
                height_pt=vb.canvas.height_pt,
                aspect_ratio=round(vb.canvas.width_pt / vb.canvas.height_pt, 4)
                if vb.canvas.height_pt else 0.0,
            ),
            slides_built=len(vb.plans),
            unmapped_shape_drift=[],
            diagnostics={
                "mean_similarity": vb.mean_similarity,
                "within_budget": vb.within_budget,
                "real_rebuild": True,
            },
        )
    raise ValueError(f"Unsupported reference mode: {mode}")
