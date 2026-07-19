"""Pagination framework for adaptive layout (PR6).

Provides generic content splitting and role-specific paginators. Derived
slides/elements receive stable, deterministic IDs so QA and repair can trace
them across re-generation.
"""
from __future__ import annotations

import copy
import math
from dataclasses import dataclass
from typing import Any, Callable

from pptx_skill.content_model import ContentSpec, ElementSpec, SlideSpec


@dataclass
class FitResult:
    """Result of fitting a list of content items into a single page."""

    fits: bool
    overflow_items: int
    estimated_score: float
    diagnostics: dict[str, Any]


@dataclass
class ContentItem:
    """A normalized content item that can be paginated."""

    text: str
    level: int = 0
    note: str = ""
    kind: str = "bullet"
    metadata: dict[str, Any] | None = None

    @property
    def keep_with_next(self) -> bool:
        return (self.metadata or {}).get("keep_with_next", False)


def _derive_element_id(parent_id: str, kind: str, index: int) -> str:
    return f"{parent_id}/{kind}-{index}"


def _derive_slide_id(parent_id: str, fragment_index: int) -> str:
    return f"{parent_id}/frag-{fragment_index}"


def paginate_items(
    items: list[ContentItem],
    capacity_fn: Callable[[list[ContentItem]], FitResult],
    keep_with_next: set[int] | None = None,
    min_items_per_page: int = 1,
    max_items_per_page: int | None = None,
) -> list[list[ContentItem]]:
    """Split a list of content items into pages.

    Greedy packing that respects ``keep_with_next`` (0-indexed indices that must
    stay on the same page as the following item) and a per-page capacity
    predicate. ``min_items_per_page`` prevents degenerate pages; if even the
    first item cannot fit, it is emitted as a single-item page so the caller
    can decide whether to treat the original slide as infeasible.
    """
    if not items:
        return []

    keep_with_next = keep_with_next or set()
    pages: list[list[ContentItem]] = []
    current: list[ContentItem] = []

    def _commit() -> None:
        if current:
            pages.append(current[:])
            current.clear()

    i = 0
    while i < len(items):
        item = items[i]
        candidate = current + [item]
        fit = capacity_fn(candidate)
        within_max = max_items_per_page is None or len(candidate) <= max_items_per_page
        # Keep-with-next: if this item is marked, it must end a page (with next item).
        must_break_after = i in keep_with_next and current

        if fit.fits and within_max and not must_break_after:
            current.append(item)
            i += 1
            continue

        # If current page is empty, emit single-item page to avoid infinite loop.
        if not current:
            pages.append([item])
            current.clear()
            i += 1
            continue

        # Must start a new page. But if this item is keep_with_next of the
        # previous index, the previous item was already committed correctly
        # because must_break_after was true *before* adding it.
        _commit()
        # Do not advance i; retry the item on the new page.

    _commit()

    # Merge undersized pages forward when possible.
    merged: list[list[ContentItem]] = []
    for page in pages:
        if not merged:
            merged.append(page)
            continue
        if len(page) < min_items_per_page:
            combined = merged[-1] + page
            if capacity_fn(combined).fits and (
                max_items_per_page is None or len(combined) <= max_items_per_page
            ):
                merged[-1] = combined
                continue
        merged.append(page)
    return merged


def _extract_body_items(element: ElementSpec) -> list[ContentItem]:
    """Normalize an element's body text into paginatable ContentItems."""
    text = element.content.get("text", "")
    if not text:
        return []
    # Split on newlines, treating each non-empty line as a bullet.
    raw_lines = [ln.strip() for ln in str(text).split("\n") if ln.strip()]
    if not raw_lines:
        return []
    items: list[ContentItem] = []
    for idx, line in enumerate(raw_lines):
        level = 0
        # Markdown-ish list nesting: leading spaces/tabs or leading dash.
        stripped = line.lstrip(" ")
        level = (len(line) - len(stripped)) // 2
        if stripped.startswith(("- ", "* ", "• ")):
            stripped = stripped[2:]
        items.append(
            ContentItem(
                text=stripped,
                level=level,
                kind="bullet",
                metadata={"source_index": idx},
            )
        )
    return items


def _split_text_element(element: ElementSpec, items: list[ContentItem]) -> ElementSpec:
    """Return a copy of element with only the given items' text."""
    new_element = copy.deepcopy(element)
    lines = [("  " * it.level) + ("- " if it.level > 0 else "") + it.text for it in items]
    new_element.content["text"] = "\n".join(lines)
    return new_element


def _capacity_for_body(
    items: list[ContentItem],
    max_items: int,
    min_font_size_pt: float,
    text_measurer: Callable[[str, float], tuple[float, bool]] | None = None,
) -> FitResult:
    """Default capacity function for body bullets.

    ``max_items`` is the primary limit. An optional ``text_measurer`` can
    provide a more accurate overflow estimate (returns required_font_size,
    overflow).
    """
    if len(items) > max_items:
        return FitResult(fits=False, overflow_items=len(items) - max_items, estimated_score=0.0, diagnostics={})
    if text_measurer:
        text = "\n".join(it.text for it in items)
        required, overflow = text_measurer(text, float(max_items))
        return FitResult(
            fits=not overflow,
            overflow_items=len(items) if overflow else 0,
            estimated_score=required,
            diagnostics={"required_font_size_pt": required},
        )
    return FitResult(fits=True, overflow_items=0, estimated_score=0.0, diagnostics={})


def paginate_bullets(
    slide: SlideSpec,
    max_items: int = 7,
    min_items_per_page: int = 2,
    min_font_size_pt: float = 9.0,
    text_measurer: Callable[[str, float], tuple[float, bool]] | None = None,
) -> list[SlideSpec] | None:
    """Split a bullets slide by its body element.

    Returns ``None`` if no split is necessary or the body element is missing.
    Derived slides keep the title and split the body; fragment_index is
    incremented deterministically.
    """
    body_element = next((e for e in slide.elements if e.role == "body"), None)
    if body_element is None:
        return None
    items = _extract_body_items(body_element)
    if len(items) <= max_items:
        return None

    capacity_fn = lambda candidates: _capacity_for_body(  # noqa: E731
        candidates, max_items, min_font_size_pt, text_measurer
    )
    pages = paginate_items(
        items,
        capacity_fn,
        keep_with_next=set(),
        min_items_per_page=min_items_per_page,
        max_items_per_page=max_items,
    )
    if len(pages) <= 1:
        return None

    title_element = next((e for e in slide.elements if e.role == "title"), None)
    derived: list[SlideSpec] = []
    for idx, page_items in enumerate(pages):
        new_elements: list[ElementSpec] = []
        if title_element is not None:
            new_title = copy.deepcopy(title_element)
            if idx > 0:
                new_title.content["text"] = f"{title_element.content.get('text', '')}（续）"
                new_title.id = _derive_element_id(title_element.id, "continued", idx)
            new_elements.append(new_title)
        new_body = _split_text_element(body_element, page_items)
        new_body.id = _derive_element_id(body_element.id, "page", idx)
        new_elements.append(new_body)
        # Preserve any other elements on the first page only.
        if idx == 0:
            for elem in slide.elements:
                if elem.role not in {"title", "body"}:
                    new_elements.append(copy.deepcopy(elem))
        derived.append(
            SlideSpec(
                id=_derive_slide_id(slide.id, idx),
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=new_elements,
                preferred_layouts=list(slide.preferred_layouts),
                source_section_id=slide.source_section_id,
                fragment_index=idx,
            )
        )
    return derived


# ---------------------------------------------------------------------------
# Role-specific paginators (PR8a): table / timeline / process / image_grid
# ---------------------------------------------------------------------------


def _split_table_element(element: ElementSpec, headers: list, rows: list) -> ElementSpec:
    new_element = copy.deepcopy(element)
    new_element.content = {"headers": list(headers), "rows": [list(r) for r in rows]}
    return new_element


def paginate_table(
    slide: SlideSpec,
    max_rows: int = 10,
    min_rows_per_page: int = 3,
) -> list[SlideSpec] | None:
    """Split a table slide by rows, repeating the header on every page.

    Returns ``None`` if no split is needed or the table element is missing.
    Per blueprint §7.7: tables repeat the header row and paginate by row height.
    """
    table_element = next((e for e in slide.elements if e.role == "table"), None)
    if table_element is None:
        return None
    headers = list(table_element.content.get("headers", []))
    rows = [list(r) for r in table_element.content.get("rows", [])]
    if len(rows) <= max_rows:
        return None

    # Pack rows into pages respecting min/max bounds.
    pages: list[list[list]] = []
    current: list[list] = []
    for row in rows:
        current.append(row)
        if len(current) >= max_rows:
            pages.append(current)
            current = []
    if current:
        if pages and len(current) < min_rows_per_page:
            # Spill small tail into the previous page to avoid a stub page.
            pages[-1].extend(current)
        else:
            pages.append(current)

    if len(pages) <= 1:
        return None

    title_element = next((e for e in slide.elements if e.role == "title"), None)
    derived: list[SlideSpec] = []
    for idx, page_rows in enumerate(pages):
        new_elements: list[ElementSpec] = []
        if title_element is not None:
            new_title = copy.deepcopy(title_element)
            if idx > 0:
                new_title.content["text"] = f"{title_element.content.get('text', '')}（续）"
                new_title.id = _derive_element_id(title_element.id, "continued", idx)
            new_elements.append(new_title)
        new_table = _split_table_element(table_element, headers, page_rows)
        new_table.id = _derive_element_id(table_element.id, "page", idx)
        new_elements.append(new_table)
        if idx == 0:
            for elem in slide.elements:
                if elem.role not in {"title", "table"}:
                    new_elements.append(copy.deepcopy(elem))
        derived.append(
            SlideSpec(
                id=_derive_slide_id(slide.id, idx),
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=new_elements,
                preferred_layouts=list(slide.preferred_layouts),
                source_section_id=slide.source_section_id,
                fragment_index=idx,
            )
        )
    return derived


def _split_items_element(element: ElementSpec, items: list) -> ElementSpec:
    new_element = copy.deepcopy(element)
    new_element.content = {"items": list(items)}
    return new_element


def paginate_timeline(
    slide: SlideSpec,
    max_stages: int = 6,
    min_stages_per_page: int = 2,
) -> list[SlideSpec] | None:
    """Split a timeline slide by stages, never breaking a single stage.

    Returns ``None`` if no split is needed or the timeline element is missing.
    Per blueprint §7.7: timelines group by stage and do not split a single stage.
    """
    tl_element = next((e for e in slide.elements if e.role in {"timeline", "items"}), None)
    if tl_element is None:
        return None
    stages = list(tl_element.content.get("items", []))
    if len(stages) <= max_stages:
        return None

    pages: list[list] = []
    current: list = []
    for stage in stages:
        current.append(stage)
        if len(current) >= max_stages:
            pages.append(current)
            current = []
    if current:
        if pages and len(current) < min_stages_per_page:
            pages[-1].extend(current)
        else:
            pages.append(current)

    if len(pages) <= 1:
        return None

    title_element = next((e for e in slide.elements if e.role == "title"), None)
    derived: list[SlideSpec] = []
    for idx, page_stages in enumerate(pages):
        new_elements: list[ElementSpec] = []
        if title_element is not None:
            new_title = copy.deepcopy(title_element)
            if idx > 0:
                new_title.content["text"] = f"{title_element.content.get('text', '')}（续）"
                new_title.id = _derive_element_id(title_element.id, "continued", idx)
            new_elements.append(new_title)
        new_tl = _split_items_element(tl_element, page_stages)
        new_tl.id = _derive_element_id(tl_element.id, "page", idx)
        new_elements.append(new_tl)
        if idx == 0:
            for elem in slide.elements:
                if elem.role not in {"title", "timeline", "items"}:
                    new_elements.append(copy.deepcopy(elem))
        derived.append(
            SlideSpec(
                id=_derive_slide_id(slide.id, idx),
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=new_elements,
                preferred_layouts=list(slide.preferred_layouts),
                source_section_id=slide.source_section_id,
                fragment_index=idx,
            )
        )
    return derived


def paginate_process(
    slide: SlideSpec,
    max_steps: int = 5,
    min_steps_per_page: int = 2,
) -> list[SlideSpec] | None:
    """Split a process slide by steps.

    Per blueprint §7.7: process prefers switching to a two-row or vertical
    candidate before paginating; pagination is the last resort. This function
    only performs the pagination step — the candidate switch is handled by the
    layout engine's fallback chain. Returns ``None`` if no split is needed.
    """
    proc_element = next((e for e in slide.elements if e.role in {"process", "steps"}), None)
    if proc_element is None:
        return None
    steps = list(proc_element.content.get("items", []))
    if len(steps) <= max_steps:
        return None

    pages: list[list] = []
    current: list = []
    for step in steps:
        current.append(step)
        if len(current) >= max_steps:
            pages.append(current)
            current = []
    if current:
        if pages and len(current) < min_steps_per_page:
            pages[-1].extend(current)
        else:
            pages.append(current)

    if len(pages) <= 1:
        return None

    title_element = next((e for e in slide.elements if e.role == "title"), None)
    derived: list[SlideSpec] = []
    for idx, page_steps in enumerate(pages):
        new_elements: list[ElementSpec] = []
        if title_element is not None:
            new_title = copy.deepcopy(title_element)
            if idx > 0:
                new_title.content["text"] = f"{title_element.content.get('text', '')}（续）"
                new_title.id = _derive_element_id(title_element.id, "continued", idx)
            new_elements.append(new_title)
        new_proc = _split_items_element(proc_element, page_steps)
        new_proc.id = _derive_element_id(proc_element.id, "page", idx)
        new_elements.append(new_proc)
        if idx == 0:
            for elem in slide.elements:
                if elem.role not in {"title", "process", "steps"}:
                    new_elements.append(copy.deepcopy(elem))
        derived.append(
            SlideSpec(
                id=_derive_slide_id(slide.id, idx),
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=new_elements,
                preferred_layouts=list(slide.preferred_layouts),
                source_section_id=slide.source_section_id,
                fragment_index=idx,
            )
        )
    return derived


# Image-grid layout selection by image count (blueprint §7.7: 2/3/4/6 layouts).
IMAGE_GRID_LAYOUTS: dict[int, str] = {
    2: "image_grid.grid2",
    3: "image_grid.grid3",
    4: "image_grid.grid2",  # 2x2 reuses grid2 zone shape
    6: "image_grid.grid3",  # 2x3 reuses grid3 zone shape
}


def paginate_image_grid(
    slide: SlideSpec,
    max_images: int = 6,
) -> list[SlideSpec] | None:
    """Pick an image-grid layout by image count and paginate overflow.

    Per blueprint §7.7: choose 2/3/4/6 layouts based on image count to avoid
    tiny images. When the count exceeds ``max_images``, split into multiple
    grid pages. Returns ``None`` if the images fit a single grid.
    """
    image_elements = [e for e in slide.elements if e.role in {"image", "hero", "supporting_image"}]
    count = len(image_elements)
    if count == 0:
        return None
    if count <= max_images:
        # Single page: just hint the best layout.
        layout = IMAGE_GRID_LAYOUTS.get(count, "image_grid.grid3")
        return [
            SlideSpec(
                id=slide.id,
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=list(slide.elements),
                preferred_layouts=list(dict.fromkeys([layout, *slide.preferred_layouts])),
                source_section_id=slide.source_section_id,
                fragment_index=0,
            )
        ]

    # Overflow: split images into pages of max_images each.
    pages: list[list[ElementSpec]] = []
    for start in range(0, count, max_images):
        pages.append(image_elements[start:start + max_images])

    title_element = next((e for e in slide.elements if e.role == "title"), None)
    derived: list[SlideSpec] = []
    for idx, page_images in enumerate(pages):
        new_elements: list[ElementSpec] = []
        if title_element is not None:
            new_title = copy.deepcopy(title_element)
            if idx > 0:
                new_title.content["text"] = f"{title_element.content.get('text', '')}（续）"
                new_title.id = _derive_element_id(title_element.id, "continued", idx)
            new_elements.append(new_title)
        for img_idx, img in enumerate(page_images):
            new_img = copy.deepcopy(img)
            new_img.id = _derive_element_id(img.id, "page", idx) if img_idx == 0 else f"{img.id}/page-{idx}-{img_idx}"
            new_elements.append(new_img)
        if idx == 0:
            for elem in slide.elements:
                if elem.role not in {"title", "image", "hero", "supporting_image"}:
                    new_elements.append(copy.deepcopy(elem))
        layout = IMAGE_GRID_LAYOUTS.get(len(page_images), "image_grid.grid3")
        derived.append(
            SlideSpec(
                id=_derive_slide_id(slide.id, idx) if idx > 0 else slide.id,
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=new_elements,
                preferred_layouts=[layout],
                source_section_id=slide.source_section_id,
                fragment_index=idx,
            )
        )
    return derived


def paginate_content_spec(
    content: ContentSpec,
    max_items: int = 7,
    min_font_size_pt: float = 9.0,
    text_measurer: Callable[[str, float], tuple[float, bool]] | None = None,
) -> tuple[ContentSpec, list[tuple[str, int, int]]]:
    """Paginate an entire ContentSpec, splitting over-loaded slides as needed.

    Dispatches on slide role to the role-specific paginator (bullets/table/
    timeline/process/image_grid). Returns a new ``ContentSpec`` with derived
    slides and a mapping list of ``(source_slide_id, fragment_index,
    derived_count)`` for traceability.
    """
    derived_slides: list[SlideSpec] = []
    mapping: list[tuple[str, int, int]] = []
    for slide in content.slides:
        split: list[SlideSpec] | None = None
        if slide.role == "bullets":
            split = paginate_bullets(
                slide,
                max_items=max_items,
                min_font_size_pt=min_font_size_pt,
                text_measurer=text_measurer,
            )
        elif slide.role == "table":
            split = paginate_table(slide)
        elif slide.role == "timeline":
            split = paginate_timeline(slide)
        elif slide.role == "process":
            split = paginate_process(slide)
        elif slide.role == "image_grid":
            split = paginate_image_grid(slide)
        if split:
            derived_slides.extend(split)
            mapping.append((slide.id, slide.fragment_index, len(split)))
            continue
        derived_slides.append(slide)
        mapping.append((slide.id, slide.fragment_index, 1))
    return (
        ContentSpec(
            id=content.id,
            title=content.title,
            subtitle=content.subtitle,
            slides=derived_slides,
            locale=content.locale,
            metadata={**content.metadata, "pagination_mapping": mapping},
        ),
        mapping,
    )
