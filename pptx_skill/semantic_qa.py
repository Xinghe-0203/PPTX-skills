"""Semantic QA for planned layouts (PR3).

Checks for overflow, overlap, contrast, missing fonts and image distortion.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from PIL import Image

from pptx_skill.content_model import BBox, LayoutPlan, PlannedNode
from pptx_skill.text_metrics import ParagraphStyle, measure_text, resolve_font


class IssueSeverity(str, Enum):
    BLOCKER = "blocker"
    WARNING = "warning"
    INFO = "info"


class IssueKind(str, Enum):
    TEXT_OVERFLOW = "text_overflow"
    OVERLAP = "overlap"
    LOW_CONTRAST = "low_contrast"
    MISSING_FONT = "missing_font"
    IMAGE_DISTORTION = "image_distortion"
    EMPTY_CONTENT = "empty_content"


@dataclass
class DetectedIssue:
    kind: IssueKind
    severity: IssueSeverity
    node_id: str | None
    message: str
    details: dict = field(default_factory=dict)


@dataclass
class SemanticQAReport:
    issues: list[DetectedIssue] = field(default_factory=list)
    passed: bool = True

    def blocker_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == IssueSeverity.BLOCKER)

    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == IssueSeverity.WARNING)


# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    if len(value) == 3:
        value = "".join(c * 2 for c in value)
    if len(value) != 6:
        raise ValueError(f"Invalid hex color {value!r}")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def _relative_luminance(rgb: tuple[int, int, int]) -> float:
    def channel(c: int) -> float:
        s = c / 255.0
        return s / 12.92 if s <= 0.03928 else ((s + 0.055) / 1.055) ** 2.4

    r, g, b = channel(rgb[0]), channel(rgb[1]), channel(rgb[2])
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(foreground: str, background: str) -> float:
    """Compute WCAG 2.0 contrast ratio between two hex colors."""
    fg = _relative_luminance(_hex_to_rgb(foreground))
    bg = _relative_luminance(_hex_to_rgb(background))
    lighter = max(fg, bg)
    darker = min(fg, bg)
    return (lighter + 0.05) / (darker + 0.05)


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------


def _bbox_overlap(a: BBox, b: BBox) -> float:
    x_left = max(a.x, b.x)
    y_top = max(a.y, b.y)
    x_right = min(a.x + a.width, b.x + b.width)
    y_bottom = min(a.y + a.height, b.y + b.height)
    if x_right <= x_left or y_bottom <= y_top:
        return 0.0
    return float((x_right - x_left) * (y_bottom - y_top))


# ---------------------------------------------------------------------------
# QA Engine
# ---------------------------------------------------------------------------


class SemanticQAEngine:
    def __init__(
        self,
        min_contrast: float = 4.5,
        overlap_threshold_pt2: float = 1.0,
        distortion_threshold: float = 0.05,
        default_bg: str = "#FFFFFF",
    ):
        self.min_contrast = min_contrast
        self.overlap_threshold_pt2 = overlap_threshold_pt2
        self.distortion_threshold = distortion_threshold
        self.default_bg = default_bg

    def check(self, plan: LayoutPlan) -> SemanticQAReport:
        issues: list[DetectedIssue] = []
        issues.extend(self._check_overlaps(plan))
        for node in plan.nodes:
            issues.extend(self._check_node(node, plan))
        passed = not any(i.severity == IssueSeverity.BLOCKER for i in issues)
        return SemanticQAReport(issues=issues, passed=passed)

    def _check_node(self, node: PlannedNode, plan: LayoutPlan) -> list[DetectedIssue]:
        issues: list[DetectedIssue] = []
        if node.kind == "text":
            issues.extend(self._check_text_overflow(node))
            issues.extend(self._check_text_contrast(node))
            issues.extend(self._check_font(node))
            issues.extend(self._check_empty_text(node))
        elif node.kind == "image":
            issues.extend(self._check_image_distortion(node))
            issues.extend(self._check_missing_image(node))
        return issues

    def _check_text_overflow(self, node: PlannedNode) -> list[DetectedIssue]:
        text = (node.content_binding or {}).get("text", "")
        if not text:
            return []
        bbox = node.geometry.bbox if node.geometry else BBox(0, 0, 100, 100)
        style = node.resolved_style or {}
        font_size = style.get("size", 16)
        font_family = style.get("font_family", "Microsoft YaHei")
        color = style.get("color", "#1A1A1A")
        bold = style.get("bold", False)
        para = ParagraphStyle(line_height=style.get("line_height", 1.35))
        try:
            metrics = measure_text(
                text,
                width_pt=bbox.width,
                font_size_pt=font_size,
                font_family=font_family,
                paragraph_style=para,
                bold=bold,
                height_pt=bbox.height,
            )
        except Exception as exc:  # pragma: no cover - defensive
            return [
                DetectedIssue(
                    kind=IssueKind.TEXT_OVERFLOW,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=f"Could not measure text: {exc}",
                )
            ]
        if metrics.overflow:
            return [
                DetectedIssue(
                    kind=IssueKind.TEXT_OVERFLOW,
                    severity=IssueSeverity.BLOCKER,
                    node_id=node.id,
                    message=(
                        f"Text overflow: {metrics.num_lines} lines, "
                        f"measured {metrics.height_pt:.1f}pt > box {bbox.height:.1f}pt"
                    ),
                    details={
                        "measured_height_pt": metrics.height_pt,
                        "box_height_pt": bbox.height,
                        "required_font_size_pt": metrics.required_font_size_pt,
                    },
                )
            ]
        return []

    def _check_text_contrast(self, node: PlannedNode) -> list[DetectedIssue]:
        style = node.resolved_style or {}
        fg = style.get("color")
        bg = style.get("fill") or self.default_bg
        if not fg:
            return []
        try:
            ratio = contrast_ratio(fg, bg)
        except Exception:
            return []
        if ratio < self.min_contrast:
            return [
                DetectedIssue(
                    kind=IssueKind.LOW_CONTRAST,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=f"Contrast ratio {ratio:.2f} < {self.min_contrast}",
                    details={"foreground": fg, "background": bg, "ratio": ratio},
                )
            ]
        return []

    def _check_font(self, node: PlannedNode) -> list[DetectedIssue]:
        style = node.resolved_style or {}
        font_family = style.get("font_family", "Microsoft YaHei")
        bold = style.get("bold", False)
        italic = style.get("italic", False)
        path = resolve_font(font_family, bold, italic)
        if path is None:
            return [
                DetectedIssue(
                    kind=IssueKind.MISSING_FONT,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=f"Font not found: {font_family} bold={bold} italic={italic}",
                )
            ]
        return []

    def _check_empty_text(self, node: PlannedNode) -> list[DetectedIssue]:
        text = (node.content_binding or {}).get("text", "")
        if node.kind == "text" and (not text or not text.strip()):
            return [
                DetectedIssue(
                    kind=IssueKind.EMPTY_CONTENT,
                    severity=IssueSeverity.INFO,
                    node_id=node.id,
                    message="Text node has empty content",
                )
            ]
        return []

    def _check_image_distortion(self, node: PlannedNode) -> list[DetectedIssue]:
        path = (node.content_binding or {}).get("path", "")
        if not path or not Path(path).exists():
            return []
        bbox = node.geometry.bbox if node.geometry else BBox(0, 0, 1, 1)
        if bbox.width <= 0 or bbox.height <= 0:
            return []
        target_ratio = bbox.width / bbox.height
        try:
            with Image.open(path) as img:
                src_ratio = img.width / img.height
        except Exception:
            return []
        if target_ratio == 0:
            return []
        distortion = abs(src_ratio - target_ratio) / target_ratio
        if distortion > self.distortion_threshold:
            return [
                DetectedIssue(
                    kind=IssueKind.IMAGE_DISTORTION,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=(
                        f"Image aspect mismatch: source {src_ratio:.2f} vs "
                        f"box {target_ratio:.2f} ({distortion:.1%} distortion)"
                    ),
                    details={
                        "source_ratio": src_ratio,
                        "box_ratio": target_ratio,
                        "distortion": distortion,
                    },
                )
            ]
        return []

    def _check_missing_image(self, node: PlannedNode) -> list[DetectedIssue]:
        path = (node.content_binding or {}).get("path", "")
        if not path:
            return [
                DetectedIssue(
                    kind=IssueKind.EMPTY_CONTENT,
                    severity=IssueSeverity.BLOCKER,
                    node_id=node.id,
                    message="Image node has no source path",
                )
            ]
        if not Path(path).exists():
            return [
                DetectedIssue(
                    kind=IssueKind.EMPTY_CONTENT,
                    severity=IssueSeverity.BLOCKER,
                    node_id=node.id,
                    message=f"Image file not found: {path}",
                )
            ]
        return []

    def _check_overlaps(self, plan: LayoutPlan) -> list[DetectedIssue]:
        issues: list[DetectedIssue] = []
        nodes = [n for n in plan.nodes if n.geometry and not n.decorative]
        for i in range(len(nodes)):
            for j in range(i + 1, len(nodes)):
                a, b = nodes[i], nodes[j]
                area = _bbox_overlap(a.geometry.bbox, b.geometry.bbox)
                if area > self.overlap_threshold_pt2:
                    issues.append(
                        DetectedIssue(
                            kind=IssueKind.OVERLAP,
                            severity=IssueSeverity.BLOCKER,
                            node_id=f"{a.id},{b.id}",
                            message=f"Nodes {a.id} and {b.id} overlap by {area:.1f} pt²",
                            details={"overlap_area_pt2": area},
                        )
                    )
        return issues


def check_layout_plan(
    plan: LayoutPlan,
    min_contrast: float = 4.5,
    overlap_threshold_pt2: float = 1.0,
    distortion_threshold: float = 0.05,
) -> SemanticQAReport:
    """Run semantic QA checks on a layout plan."""
    return SemanticQAEngine(
        min_contrast=min_contrast,
        overlap_threshold_pt2=overlap_threshold_pt2,
        distortion_threshold=distortion_threshold,
    ).check(plan)
