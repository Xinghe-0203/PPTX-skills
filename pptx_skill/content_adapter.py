"""Adapter from legacy Section dicts to the PR1 ContentSpec model.

This is a lossy but deterministic bridge: it turns the existing flat
``Section`` data structure into ``SlideSpec`` / ``ElementSpec`` so the new
layout engine, QA rules and repair engine can consume legacy content without
rewriting ``auto_generate_ppt`` callers.
"""
from __future__ import annotations

from typing import Any

from pptx_skill.content_model import ContentSpec, ElementSpec, SlideSpec, StableIdGenerator


def _infer_layout(section: dict[str, Any]) -> str:
    """Best-effort layout inference mirroring ``choose_layout`` signals."""
    if section.get("metrics"):
        return "dashboard"
    if section.get("events") or section.get("timeline"):
        return "timeline"
    if section.get("left") and section.get("right"):
        return "comparison"
    if section.get("quote"):
        return "quote"
    if section.get("steps") or section.get("process"):
        return "process"
    if section.get("table_headers"):
        return "table"
    images = section.get("images") or []
    n_images = len([im for im in images if im])
    has_bullets = bool(section.get("bullets"))
    if n_images >= 3:
        return "image_grid"
    if n_images >= 1 and has_bullets:
        return "text_image"
    if n_images >= 1 and not has_bullets:
        return "full_image"
    if has_bullets:
        return "bullets"
    return "section"


def _section_to_elements(
    section: dict[str, Any],
    gen: StableIdGenerator,
    source_section_id: str,
    fragment_index: int,
) -> list[ElementSpec]:
    """Create ElementSpecs from a legacy section dict."""
    elements: list[ElementSpec] = []

    def add(kind: str, role: str, content: dict[str, Any], item_key: str = "") -> None:
        elements.append(
            ElementSpec(
                id=gen.element_id(source_section_id, fragment_index, role, item_key or role),
                kind=kind,
                role=role,
                content=content,
                style_ref=f"component.{role}",
            )
        )

    if section.get("title"):
        add("text", "title", {"text": section["title"]}, "title")
    if section.get("subtitle"):
        add("text", "subtitle", {"text": section["subtitle"]}, "subtitle")
    if section.get("kicker"):
        add("text", "kicker", {"text": section["kicker"]}, "kicker")

    bullets = section.get("bullets") or []
    if bullets:
        add("text", "body", {"items": list(bullets)}, "body")

    quote = section.get("quote")
    if quote:
        add("text", "quote", {"text": quote, "source": section.get("source", "")}, "quote")

    images = section.get("images") or []
    for idx, image_path in enumerate(images):
        if image_path:
            add("image", "hero" if idx == 0 else "supporting_image",
                {"path": image_path, "fit": "smart-cover"}, f"image-{idx}")

    metrics = section.get("metrics") or []
    if metrics:
        add("chart", "metric_group", {"items": list(metrics)}, "metrics")

    steps = section.get("steps") or section.get("process") or []
    if steps:
        add("shape", "process", {"items": list(steps)}, "process")

    events = section.get("events") or section.get("timeline") or []
    if events:
        add("shape", "timeline", {"items": list(events)}, "timeline")

    table_headers = section.get("table_headers") or []
    table_rows = section.get("table_rows") or []
    if table_headers or table_rows:
        add("table", "table", {"headers": list(table_headers), "rows": [list(r) for r in table_rows]}, "table")

    left = section.get("left") or {}
    right = section.get("right") or {}
    if left or right:
        add("group", "comparison", {"left": dict(left), "right": dict(right)}, "comparison")

    return elements


class LegacyContentAdapter:
    """Convert legacy ``Section`` dicts into ``ContentSpec``."""

    def __init__(self, title: str = "", subtitle: str = ""):
        self.title = title
        self.subtitle = subtitle

    def adapt(
        self,
        sections: list[dict[str, Any] | Any],
        locale: str = "zh-CN",
        metadata: dict[str, Any] | None = None,
    ) -> ContentSpec:
        """Adapt a list of legacy sections to a ContentSpec.

        Each section may be a plain dict or a ``Section`` dataclass instance.
        """
        gen = StableIdGenerator.from_seed(self.title or "untitled")
        slides: list[SlideSpec] = []

        for ordinal, raw in enumerate(sections):
            if isinstance(raw, dict):
                section = raw
            else:
                # Support legacy Section dataclass.
                section = {
                    field: getattr(raw, field, None)
                    for field in getattr(raw, "__dataclass_fields__", {})
                }
            source_section_id = gen.source_section_id(ordinal)
            layout = section.get("layout") or _infer_layout(section)
            slide_id = gen.slide_id(source_section_id, fragment_index=0)
            slide = SlideSpec(
                id=slide_id,
                role=layout,
                communication_goal=section.get("title", ""),
                elements=_section_to_elements(section, gen, source_section_id, fragment_index=0),
                preferred_layouts=[layout],
                source_section_id=source_section_id,
                fragment_index=0,
            )
            slides.append(slide)

        return ContentSpec(
            id=gen.namespace,
            title=self.title,
            subtitle=self.subtitle,
            slides=slides,
            locale=locale,
            metadata=metadata or {},
        )


def adapt_legacy_sections(
    title: str,
    subtitle: str,
    sections: list[dict[str, Any] | Any],
    locale: str = "zh-CN",
    metadata: dict[str, Any] | None = None,
) -> ContentSpec:
    """Convenience wrapper around ``LegacyContentAdapter.adapt``."""
    return LegacyContentAdapter(title=title, subtitle=subtitle).adapt(
        sections, locale=locale, metadata=metadata
    )
