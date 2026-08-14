"""Core domain models for the PPTX skill (PR1 data contract).

All coordinates are in points (1 inch = 72 pt). This module must not import
``python-pptx``; it is consumed by the layout solver, QA rules and repair
engine as well as the renderer adapters.
"""
from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from typing import Any, Literal

# ---------------------------------------------------------------------------
# Geometry primitives
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class SafeInsets:
    """Safe-area insets from each canvas edge, in points."""

    top: float
    right: float
    bottom: float
    left: float


@dataclass(frozen=True)
class CanvasSpec:
    """Slide canvas specification."""

    width_pt: float
    height_pt: float
    name: str = "16:9"
    safe: SafeInsets = field(default_factory=lambda: SafeInsets(36, 48, 32, 48))

    @property
    def safe_width(self) -> float:
        return self.width_pt - self.safe.left - self.safe.right

    @property
    def safe_height(self) -> float:
        return self.height_pt - self.safe.top - self.safe.bottom


@dataclass(frozen=True)
class BBox:
    """Axis-aligned bounding box in points."""

    x: float
    y: float
    width: float
    height: float

    @property
    def left(self) -> float:
        return self.x

    @property
    def top(self) -> float:
        return self.y

    @property
    def right(self) -> float:
        return self.x + self.width

    @property
    def bottom(self) -> float:
        return self.y + self.height

    @property
    def center_x(self) -> float:
        return self.x + self.width / 2

    @property
    def center_y(self) -> float:
        return self.y + self.height / 2

    def area(self) -> float:
        return self.width * self.height

    def intersects(self, other: BBox) -> bool:
        return not (
            self.right <= other.left
            or self.left >= other.right
            or self.bottom <= other.top
            or self.top >= other.bottom
        )

    def intersection(self, other: BBox) -> BBox | None:
        if not self.intersects(other):
            return None
        x1 = max(self.left, other.left)
        y1 = max(self.top, other.top)
        x2 = min(self.right, other.right)
        y2 = min(self.bottom, other.bottom)
        return BBox(x1, y1, x2 - x1, y2 - y1)


@dataclass(frozen=True)
class GeometrySpec:
    """Resolved geometry of a render node."""

    bbox: BBox
    rotation_deg: float = 0.0
    polygon: tuple[tuple[float, float], ...] | None = None


# ---------------------------------------------------------------------------
# Stable identity
# ---------------------------------------------------------------------------

_NAMESPACE_SEED = uuid.UUID("c0c3d4e5-f6a7-8901-2345-6789abcdef01")


def _slugify(value: str) -> str:
    """Minimal slug for stable namespace seeds."""
    return "".join(
        c.lower() if c.isalnum() or c in {"-", "_", ".", " "} else "-"
        for c in value
    ).strip() or "untitled"


def make_namespace(seed: str) -> str:
    """Deterministically create a deck namespace from a seed string.

    The same seed yields the same namespace, so regeneration from an unchanged
    title produces a reproducible lineage.
    """
    return str(uuid.uuid5(_NAMESPACE_SEED, _slugify(seed)))


@dataclass
class StableIdGenerator:
    """Generate stable, lineage-based IDs for a deck.

    IDs are plain path-like strings so they are human-readable in manifest and
    QA reports. Uniqueness is enforced per generator instance; loading a deck
    must reuse the same namespace and source IDs to keep identities stable.
    """

    namespace: str
    _seen: set[str] = field(default_factory=set, repr=False)

    @classmethod
    def from_seed(cls, seed: str) -> StableIdGenerator:
        return cls(namespace=make_namespace(seed))

    @classmethod
    def from_manifest(cls, manifest: dict | None) -> StableIdGenerator:
        if manifest is None:
            return cls.from_seed("")
        namespace = manifest.get("deck_namespace") or manifest.get("namespace")
        if not namespace:
            # Derive deterministically from content title if no namespace stored.
            title = (manifest.get("title") or manifest.get("content", {}).get("title") or "")
            namespace = make_namespace(title)
        gen = cls(namespace=namespace)
        # Pre-register existing IDs from manifest so re-generation does not reuse them.
        for slide in manifest.get("current", {}).get("slides", []):
            gen._track(slide.get("id"))
            for elem in slide.get("elements", []):
                gen._track(elem.get("id"))
        return gen

    def _track(self, id_value: str | None) -> None:
        if id_value:
            self._seen.add(id_value)

    def _ensure_unique(self, candidate: str) -> str:
        """Return the candidate if unseen; deterministic callers must vary the path.

        If the same caller asks for the same lineage twice, the same ID is
        returned. Genuinely colliding independent generations will silently
        share an ID — callers that need guaranteed uniqueness must vary
        ``item_key`` or ``occurrence``.
        """
        if candidate in self._seen:
            return candidate
        self._seen.add(candidate)
        return candidate

    def _join(self, *parts: str) -> str:
        return "/".join(p.strip("/") for p in parts if p is not None and p != "")

    def source_section_id(self, ordinal: int, role: str = "") -> str:
        base = f"sec-{ordinal:02d}"
        if role:
            return self._join(self.namespace, base, role)
        return self._join(self.namespace, base)

    def slide_id(self, source_section_id: str, fragment_index: int = 0) -> str:
        base = self._join(source_section_id, f"frag-{fragment_index}")
        return self._ensure_unique(base)

    def element_id(
        self,
        source_section_id: str,
        fragment_index: int,
        role: str,
        item_key: str = "",
    ) -> str:
        base = self._join(
            source_section_id,
            f"frag-{fragment_index}",
            role,
            item_key or "main",
        )
        return self._ensure_unique(base)

    def render_node_id(
        self,
        slide_id: str,
        node_key: str,
        occurrence: int = 0,
    ) -> str:
        base = self._join(slide_id, "node", node_key)
        if occurrence:
            base = f"{base}@{occurrence}"
        return self._ensure_unique(base)

    def derive_id(self, parent_id: str, kind: str, index: int = 0) -> str:
        base = self._join(parent_id, kind, str(index))
        return self._ensure_unique(base)


# ---------------------------------------------------------------------------
# Content model
# ---------------------------------------------------------------------------

@dataclass
class ElementSpec:
    """A logical content element on a slide."""

    id: str
    kind: Literal["text", "image", "shape", "table", "chart", "group", "video", "audio"]
    role: str
    content: dict[str, Any]
    style_ref: str
    constraints: list[dict[str, Any]] = field(default_factory=list)
    decorative: bool = False


@dataclass
class SlideSpec:
    """A logical slide in the content plan."""

    id: str
    role: str
    communication_goal: str
    elements: list[ElementSpec]
    preferred_layouts: list[str] = field(default_factory=list)
    source_section_id: str | None = None
    fragment_index: int = 0


@dataclass
class ContentSpec:
    """Whole-deck content specification."""

    id: str
    title: str
    subtitle: str
    slides: list[SlideSpec]
    locale: str = "zh-CN"
    metadata: dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Layout planning model
# ---------------------------------------------------------------------------

def normalize_content_binding(content: Any) -> dict[str, Any]:
    """Normalize an element's ``content`` into a renderer-ready binding dict.

    ``content_binding`` is declared as ``dict[str, Any]`` (e.g. keys
    ``text``/``path``/``paragraphs``). When callers pass a plain string it is
    wrapped as ``{"text": ...}`` so renderers and QA engines can rely on
    ``.get()`` semantics.
    """
    if isinstance(content, dict):
        return content
    if isinstance(content, str):
        return {"text": content}
    return {}


@dataclass
class PlannedNode:
    """A resolved render node inside a LayoutPlan."""

    id: str
    element_id: str | None
    recipe_node_id: str
    kind: str
    role: str
    geometry: GeometrySpec
    resolved_style: dict[str, Any]
    content_binding: dict[str, Any]
    z_order: int
    crop: dict[str, Any] | None = None
    decorative: bool = False
    parent_node_id: str | None = None


@dataclass
class LayoutPlan:
    """Resolved layout for a single slide."""

    canvas: CanvasSpec
    recipe_id: str
    nodes: list[PlannedNode]
    local_score: float
    has_blocker: bool = False
    diagnostics: list[dict[str, Any]] = field(default_factory=list)
    background_color: str | None = None


@dataclass
class SlidePlanCandidate:
    """A candidate bundle for a source slide (one or more derived slides)."""

    derived_slides: list[SlideSpec]
    plans: list[LayoutPlan]
    local_score: float
    diagnostics: list[dict[str, Any]] = field(default_factory=list)
    kind: Literal["single", "split"] = "single"


@dataclass
class SlidePlanResult:
    """Result of planning a single source slide."""

    status: Literal["feasible", "infeasible"]
    candidates: list[SlidePlanCandidate]
    blockers: list[dict[str, Any]]
    diagnostics: list[dict[str, Any]]


@dataclass
class DeckPlanResult:
    """Result of planning a whole deck."""

    status: Literal["feasible", "infeasible"]
    source_slides: list[SlideSpec]
    derived_slides: list[SlideSpec]
    plans: list[LayoutPlan]
    blockers: list[dict[str, Any]]
    diagnostics: list[dict[str, Any]]


# ---------------------------------------------------------------------------
# Canvas presets
# ---------------------------------------------------------------------------

CANVAS_PRESETS: dict[str, CanvasSpec] = {
    "16:9": CanvasSpec(width_pt=959.976, height_pt=540.0),
    "4:3": CanvasSpec(width_pt=720.0, height_pt=540.0),
    "9:16": CanvasSpec(width_pt=540.0, height_pt=959.976),
}


def canvas_from_name(name: str) -> CanvasSpec:
    """Return a canvas preset by name, or raise ValueError."""
    if name in CANVAS_PRESETS:
        return CANVAS_PRESETS[name]
    raise ValueError(f"Unknown canvas preset: {name}")
