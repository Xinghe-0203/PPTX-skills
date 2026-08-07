"""Semantic QA for planned layouts (PR3).

Checks for overflow, overlap, contrast, missing fonts, image distortion,
table issues, chart issues, slide-level issues, deck-level issues, and
typography hierarchy.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image

from pptx_skill._compat import StrEnum
from pptx_skill.content_model import BBox, CanvasSpec, LayoutPlan, PlannedNode
from pptx_skill.text_metrics import ParagraphStyle, measure_text, resolve_font


class IssueSeverity(StrEnum):
    BLOCKER = "blocker"
    WARNING = "warning"
    INFO = "info"


class IssueKind(StrEnum):
    TEXT_OVERFLOW = "text_overflow"
    OVERLAP = "overlap"
    LOW_CONTRAST = "low_contrast"
    MISSING_FONT = "missing_font"
    IMAGE_DISTORTION = "image_distortion"
    EMPTY_CONTENT = "empty_content"
    # Table QA
    TABLE_COLUMN_OVERFLOW = "table_column_overflow"
    TABLE_RAGGED_ROWS = "table_ragged_rows"
    TABLE_MISSING_HEADERS = "table_missing_headers"
    # Chart QA
    CHART_EMPTY_SERIES = "chart_empty_series"
    CHART_SINGLE_CATEGORY = "chart_single_category"
    CHART_LABEL_DENSITY = "chart_label_density"
    # Slide-level QA
    SAFE_MARGIN_VIOLATION = "safe_margin_violation"
    LOW_WHITESPACE_RATIO = "low_whitespace_ratio"
    MISSING_TITLE = "missing_title"
    EXCESSIVE_ELEMENT_COUNT = "excessive_element_count"
    # Deck-level QA
    FONT_SIZE_DRIFT = "font_size_drift"
    LAYOUT_REPETITION = "layout_repetition"
    MISSING_SECTION_BREAK = "missing_section_break"
    COLOR_PALETTE_DRIFT = "color_palette_drift"
    # Typography hierarchy
    TYPOGRAPHY_HIERARCHY_VIOLATION = "typography_hierarchy_violation"


@dataclass
class DetectedIssue:
    kind: IssueKind
    severity: IssueSeverity
    node_id: str | None
    message: str
    details: dict = field(default_factory=dict)
    slide_index: int = -1


@dataclass
class SemanticQAReport:
    issues: list[DetectedIssue] = field(default_factory=list)
    deck_issues: list[DetectedIssue] = field(default_factory=list)
    passed: bool = True

    def blocker_count(self) -> int:
        return sum(
            1
            for i in self.issues + self.deck_issues
            if i.severity == IssueSeverity.BLOCKER
        )

    def warning_count(self) -> int:
        return sum(
            1
            for i in self.issues + self.deck_issues
            if i.severity == IssueSeverity.WARNING
        )


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


def _bbox_outside_safe(bbox: BBox, canvas: CanvasSpec) -> bool:
    """Return True if bbox extends outside the canvas safe insets."""
    safe_left = canvas.safe.left
    safe_top = canvas.safe.top
    safe_right = canvas.width_pt - canvas.safe.right
    safe_bottom = canvas.height_pt - canvas.safe.bottom
    return (
        bbox.x < safe_left
        or bbox.y < safe_top
        or bbox.x + bbox.width > safe_right
        or bbox.y + bbox.height > safe_bottom
    )


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
        max_element_count: int = 15,
        min_whitespace_ratio: float = 0.20,
        typography_min_ratio: float = 1.5,
        chart_label_density_threshold: int = 12,
        font_drift_pt: float = 4.0,
        layout_repetition_run: int = 3,
        section_break_min_slides: int = 5,
        safe_margin_tolerance_pt: float = 2.0,
    ):
        self.min_contrast = min_contrast
        self.overlap_threshold_pt2 = overlap_threshold_pt2
        self.distortion_threshold = distortion_threshold
        self.default_bg = default_bg
        self.max_element_count = max_element_count
        self.min_whitespace_ratio = min_whitespace_ratio
        self.typography_min_ratio = typography_min_ratio
        self.chart_label_density_threshold = chart_label_density_threshold
        self.font_drift_pt = font_drift_pt
        self.layout_repetition_run = layout_repetition_run
        self.section_break_min_slides = section_break_min_slides
        self.safe_margin_tolerance_pt = safe_margin_tolerance_pt

    def check(self, plan: LayoutPlan) -> SemanticQAReport:
        issues: list[DetectedIssue] = []
        issues.extend(self._check_overlaps(plan))
        for node in plan.nodes:
            issues.extend(self._check_node(node, plan))
        # Slide-level checks
        issues.extend(self._check_safe_margins(plan))
        issues.extend(self._check_whitespace_ratio(plan))
        issues.extend(self._check_missing_title(plan))
        issues.extend(self._check_excessive_elements(plan))
        issues.extend(self._check_typography_hierarchy(plan))
        passed = not any(i.severity == IssueSeverity.BLOCKER for i in issues)
        return SemanticQAReport(issues=issues, passed=passed)

    def check_deck(self, plans: list[LayoutPlan]) -> SemanticQAReport:
        """Run deck-level QA checks across multiple slides."""
        deck_issues: list[DetectedIssue] = []

        # Build per-slide report first
        all_issues: list[DetectedIssue] = []
        for idx, plan in enumerate(plans):
            slide_report = self.check(plan)
            for issue in slide_report.issues:
                issue.slide_index = idx
            all_issues.extend(slide_report.issues)

        # Deck-level checks
        deck_issues.extend(self._check_font_size_drift(plans))
        deck_issues.extend(self._check_layout_repetition(plans))
        deck_issues.extend(self._check_missing_section_break(plans))
        deck_issues.extend(self._check_color_palette_drift(plans))

        passed = not any(
            i.severity == IssueSeverity.BLOCKER
            for i in all_issues + deck_issues
        )
        return SemanticQAReport(issues=all_issues, deck_issues=deck_issues, passed=passed)

    # -----------------------------------------------------------------------
    # Node-level checks (existing)
    # -----------------------------------------------------------------------

    def _check_node(self, node: PlannedNode, plan: LayoutPlan) -> list[DetectedIssue]:
        issues: list[DetectedIssue] = []
        if node.kind == "text":
            issues.extend(self._check_text_overflow(node))
            issues.extend(self._check_text_contrast(node, plan))
            issues.extend(self._check_font(node))
            issues.extend(self._check_empty_text(node))
        elif node.kind == "image":
            issues.extend(self._check_image_distortion(node))
            issues.extend(self._check_missing_image(node))
        elif node.kind == "table":
            issues.extend(self._check_table_missing_headers(node))
            issues.extend(self._check_table_ragged_rows(node))
            issues.extend(self._check_table_column_overflow(node))
        elif node.kind == "chart":
            issues.extend(self._check_chart_empty_series(node))
            issues.extend(self._check_chart_single_category(node))
            issues.extend(self._check_chart_label_density(node))
        return issues

    def _check_text_overflow(self, node: PlannedNode) -> list[DetectedIssue]:
        text = (node.content_binding or {}).get("text", "")
        if not text:
            return []
        bbox = node.geometry.bbox if node.geometry else BBox(0, 0, 100, 100)
        style = node.resolved_style or {}
        font_size = style.get("size", 16)
        font_family = style.get("font_family", "Microsoft YaHei")
        style.get("color", "#1A1A1A")
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

    def _check_text_contrast(self, node: PlannedNode, plan: LayoutPlan) -> list[DetectedIssue]:
        style = node.resolved_style or {}
        fg = style.get("color")
        if not fg:
            return []
        # Determine background: node fill > slide background > default
        bg = style.get("fill")
        if not bg:
            bg = plan.background_color or self.default_bg
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

    # -----------------------------------------------------------------------
    # Table QA checks
    # -----------------------------------------------------------------------

    def _check_table_missing_headers(self, node: PlannedNode) -> list[DetectedIssue]:
        binding = node.content_binding or {}
        headers = binding.get("headers", [])
        if not headers:
            return [
                DetectedIssue(
                    kind=IssueKind.TABLE_MISSING_HEADERS,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message="Table has no header row",
                )
            ]
        # Also check if all header cells are empty
        if all(not str(h).strip() for h in headers):
            return [
                DetectedIssue(
                    kind=IssueKind.TABLE_MISSING_HEADERS,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message="Table header row is entirely empty",
                )
            ]
        return []

    def _check_table_ragged_rows(self, node: PlannedNode) -> list[DetectedIssue]:
        binding = node.content_binding or {}
        rows = binding.get("rows", [])
        if not rows or len(rows) < 2:
            return []
        col_counts = [len(row) for row in rows if isinstance(row, (list, tuple))]
        if len(set(col_counts)) > 1:
            return [
                DetectedIssue(
                    kind=IssueKind.TABLE_RAGGED_ROWS,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=f"Table rows have inconsistent column counts: {col_counts}",
                    details={"column_counts": col_counts},
                )
            ]
        return []

    def _check_table_column_overflow(self, node: PlannedNode) -> list[DetectedIssue]:
        binding = node.content_binding or {}
        rows = binding.get("rows", [])
        col_widths = binding.get("col_widths", [])
        if not rows or not col_widths:
            return []

        def _has_cjk(text: str) -> bool:
            """Return True if text contains CJK or other wide characters."""
            return any(ord(ch) > 0x2E80 for ch in text)

        issues: list[DetectedIssue] = []
        for row_idx, row in enumerate(rows):
            if not isinstance(row, (list, tuple)):
                continue
            for col_idx, cell in enumerate(row):
                if col_idx >= len(col_widths):
                    continue
                text = str(cell) if cell is not None else ""
                # Use wider avg char width for CJK content
                avg_char_width = 14.0 if _has_cjk(text) else 7.0
                estimated_width = len(text) * avg_char_width
                col_width = col_widths[col_idx]
                if col_width > 0 and estimated_width > col_width:
                    issues.append(
                        DetectedIssue(
                            kind=IssueKind.TABLE_COLUMN_OVERFLOW,
                            severity=IssueSeverity.WARNING,
                            node_id=node.id,
                            message=(
                                f"Table cell [{row_idx},{col_idx}] may overflow: "
                                f"~{estimated_width:.0f}pt > col width {col_width:.0f}pt"
                            ),
                            details={
                                "row": row_idx,
                                "col": col_idx,
                                "estimated_width_pt": estimated_width,
                                "col_width_pt": col_width,
                            },
                        )
                    )
        return issues

    # -----------------------------------------------------------------------
    # Chart QA checks
    # -----------------------------------------------------------------------

    def _check_chart_empty_series(self, node: PlannedNode) -> list[DetectedIssue]:
        binding = node.content_binding or {}
        series = binding.get("series", [])
        if not series:
            return [
                DetectedIssue(
                    kind=IssueKind.CHART_EMPTY_SERIES,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message="Chart has no data series",
                )
            ]
        # Check if all values are zero
        all_zero = True
        for s in series:
            values = s.get("values", []) if isinstance(s, dict) else []
            if any(v != 0 for v in values):
                all_zero = False
                break
        if all_zero:
            return [
                DetectedIssue(
                    kind=IssueKind.CHART_EMPTY_SERIES,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message="Chart series contain all-zero values",
                )
            ]
        return []

    def _check_chart_single_category(self, node: PlannedNode) -> list[DetectedIssue]:
        binding = node.content_binding or {}
        categories = binding.get("categories", [])
        if not isinstance(categories, (list, tuple)):
            return []
        if len(categories) <= 1:
            return [
                DetectedIssue(
                    kind=IssueKind.CHART_SINGLE_CATEGORY,
                    severity=IssueSeverity.INFO,
                    node_id=node.id,
                    message=f"Chart has only {len(categories)} category (not very informative)",
                    details={"category_count": len(categories)},
                )
            ]
        return []

    def _check_chart_label_density(self, node: PlannedNode) -> list[DetectedIssue]:
        binding = node.content_binding or {}
        categories = binding.get("categories", [])
        if not isinstance(categories, (list, tuple)):
            return []
        num_categories = len(categories)
        if num_categories <= self.chart_label_density_threshold:
            return []
        # Estimate chart width from node geometry
        bbox = node.geometry.bbox if node.geometry else None
        chart_width = bbox.width if bbox else 0
        # Heuristic: small chart is < 400pt wide
        if chart_width > 0 and chart_width < 400:
            return [
                DetectedIssue(
                    kind=IssueKind.CHART_LABEL_DENSITY,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=(
                        f"Chart has {num_categories} categories in a small chart "
                        f"({chart_width:.0f}pt wide) — labels may overlap"
                    ),
                    details={
                        "category_count": num_categories,
                        "chart_width_pt": chart_width,
                    },
                )
            ]
        # Even for larger charts, warn if extremely dense
        if num_categories > 20:
            return [
                DetectedIssue(
                    kind=IssueKind.CHART_LABEL_DENSITY,
                    severity=IssueSeverity.WARNING,
                    node_id=node.id,
                    message=(
                        f"Chart has {num_categories} categories — labels may overlap"
                    ),
                    details={"category_count": num_categories},
                )
            ]
        return []

    # -----------------------------------------------------------------------
    # Slide-level QA checks
    # -----------------------------------------------------------------------

    def _check_safe_margins(self, plan: LayoutPlan) -> list[DetectedIssue]:
        issues: list[DetectedIssue] = []
        for node in plan.nodes:
            if node.decorative or not node.geometry:
                continue
            if _bbox_outside_safe(node.geometry.bbox, plan.canvas):
                # Compute how far outside the safe area the bbox extends
                bbox = node.geometry.bbox
                safe = plan.canvas.safe
                overflow = max(
                    safe.left - bbox.x,
                    safe.top - bbox.y,
                    (bbox.x + bbox.width) - (plan.canvas.width_pt - safe.right),
                    (bbox.y + bbox.height) - (plan.canvas.height_pt - safe.bottom),
                    0.0,
                )
                severity = (
                    IssueSeverity.BLOCKER
                    if overflow > self.safe_margin_tolerance_pt
                    else IssueSeverity.WARNING
                )
                issues.append(
                    DetectedIssue(
                        kind=IssueKind.SAFE_MARGIN_VIOLATION,
                        severity=severity,
                        node_id=node.id,
                        message=f"Node {node.id} extends outside safe margins",
                        details={
                            "bbox": {
                                "x": node.geometry.bbox.x,
                                "y": node.geometry.bbox.y,
                                "width": node.geometry.bbox.width,
                                "height": node.geometry.bbox.height,
                            },
                            "safe_insets": {
                                "top": plan.canvas.safe.top,
                                "right": plan.canvas.safe.right,
                                "bottom": plan.canvas.safe.bottom,
                                "left": plan.canvas.safe.left,
                            },
                            "overflow_pt": overflow,
                            "tolerance_pt": self.safe_margin_tolerance_pt,
                        },
                    )
                )
        return issues

    def _check_whitespace_ratio(self, plan: LayoutPlan) -> list[DetectedIssue]:
        safe_area = plan.canvas.safe_width * plan.canvas.safe_height
        if safe_area <= 0:
            return []
        content_nodes = [
            n for n in plan.nodes if not n.decorative and n.geometry
        ]
        # If overlaps exist, the simple area sum double-counts overlapping
        # regions, so skip the whitespace check to avoid false positives.
        has_overlaps = False
        for i in range(len(content_nodes)):
            for j in range(i + 1, len(content_nodes)):
                if _bbox_overlap(
                    content_nodes[i].geometry.bbox,
                    content_nodes[j].geometry.bbox,
                ) > self.overlap_threshold_pt2:
                    has_overlaps = True
                    break
            if has_overlaps:
                break
        if has_overlaps:
            return []
        total_node_area = 0.0
        for node in content_nodes:
            total_node_area += node.geometry.bbox.area()
        covered_ratio = total_node_area / safe_area
        whitespace_ratio = 1.0 - covered_ratio
        if whitespace_ratio < self.min_whitespace_ratio:
            return [
                DetectedIssue(
                    kind=IssueKind.LOW_WHITESPACE_RATIO,
                    severity=IssueSeverity.WARNING,
                    node_id=None,
                    message=(
                        f"Low whitespace: only {whitespace_ratio:.0%} of safe area is empty "
                        f"(threshold: {self.min_whitespace_ratio:.0%})"
                    ),
                    details={
                        "whitespace_ratio": whitespace_ratio,
                        "covered_ratio": covered_ratio,
                        "safe_area_pt2": safe_area,
                        "node_area_pt2": total_node_area,
                    },
                )
            ]
        return []

    # Recipe IDs for slides that intentionally have no title
    _NO_TITLE_RECIPES = ("cover.", "section.", "toc.", "end.", "full_image.")

    def _check_missing_title(self, plan: LayoutPlan) -> list[DetectedIssue]:
        # Skip slides whose recipe intentionally has no title
        recipe = plan.recipe_id or ""
        if any(recipe.startswith(prefix) for prefix in self._NO_TITLE_RECIPES):
            return []
        has_title = any(n.role == "title" for n in plan.nodes)
        if not has_title:
            return [
                DetectedIssue(
                    kind=IssueKind.MISSING_TITLE,
                    severity=IssueSeverity.WARNING,
                    node_id=None,
                    message="No node with role 'title' on this slide",
                )
            ]
        return []

    def _check_excessive_elements(self, plan: LayoutPlan) -> list[DetectedIssue]:
        node_count = sum(1 for n in plan.nodes if not n.decorative)
        if node_count > self.max_element_count:
            return [
                DetectedIssue(
                    kind=IssueKind.EXCESSIVE_ELEMENT_COUNT,
                    severity=IssueSeverity.WARNING,
                    node_id=None,
                    message=(
                        f"Slide has {node_count} nodes (threshold: {self.max_element_count})"
                    ),
                    details={"node_count": node_count},
                )
            ]
        return []

    # -----------------------------------------------------------------------
    # Typography hierarchy check
    # -----------------------------------------------------------------------

    def _check_typography_hierarchy(self, plan: LayoutPlan) -> list[DetectedIssue]:
        title_sizes: list[float] = []
        body_sizes: list[float] = []
        for node in plan.nodes:
            style = node.resolved_style or {}
            font_size = style.get("size")
            if font_size is None:
                continue
            if node.role == "title":
                title_sizes.append(float(font_size))
            elif node.role in ("body", "content", "text"):
                body_sizes.append(float(font_size))
        if not title_sizes or not body_sizes:
            return []
        max_title = max(title_sizes)
        max_body = max(body_sizes)
        if max_body <= 0:
            return []
        ratio = max_title / max_body
        if ratio < self.typography_min_ratio:
            return [
                DetectedIssue(
                    kind=IssueKind.TYPOGRAPHY_HIERARCHY_VIOLATION,
                    severity=IssueSeverity.WARNING,
                    node_id=None,
                    message=(
                        f"Title font size ({max_title:.1f}pt) is not significantly "
                        f"larger than body ({max_body:.1f}pt): ratio {ratio:.2f} < {self.typography_min_ratio}"
                    ),
                    details={
                        "title_size_pt": max_title,
                        "body_size_pt": max_body,
                        "ratio": ratio,
                        "min_ratio": self.typography_min_ratio,
                    },
                )
            ]
        return []

    # -----------------------------------------------------------------------
    # Deck-level QA checks
    # -----------------------------------------------------------------------

    def _check_font_size_drift(self, plans: list[LayoutPlan]) -> list[DetectedIssue]:
        """Check if same-role font sizes vary significantly across slides."""
        # Collect font sizes per role across all slides
        role_sizes: dict[str, list[tuple[int, float]]] = {}  # role -> [(slide_idx, size)]
        for idx, plan in enumerate(plans):
            for node in plan.nodes:
                style = node.resolved_style or {}
                font_size = style.get("size")
                if font_size is None:
                    continue
                role = node.role
                if role not in role_sizes:
                    role_sizes[role] = []
                role_sizes[role].append((idx, float(font_size)))

        issues: list[DetectedIssue] = []
        for role, size_list in role_sizes.items():
            if len(size_list) < 2:
                continue
            sizes = [s for _, s in size_list]
            min_size = min(sizes)
            max_size = max(sizes)
            drift = max_size - min_size
            if drift > self.font_drift_pt:
                issues.append(
                    DetectedIssue(
                        kind=IssueKind.FONT_SIZE_DRIFT,
                        severity=IssueSeverity.WARNING,
                        node_id=None,
                        message=(
                            f"Font size drift for role '{role}': "
                            f"{min_size:.1f}pt to {max_size:.1f}pt (drift: {drift:.1f}pt)"
                        ),
                        details={
                            "role": role,
                            "min_size_pt": min_size,
                            "max_size_pt": max_size,
                            "drift_pt": drift,
                        },
                    )
                )
        return issues

    def _check_layout_repetition(self, plans: list[LayoutPlan]) -> list[DetectedIssue]:
        """Check for 3+ consecutive slides with the same recipe."""
        issues: list[DetectedIssue] = []
        if len(plans) < self.layout_repetition_run:
            return issues
        run_start = 0
        for i in range(1, len(plans)):
            cur_id = plans[i].recipe_id
            start_id = plans[run_start].recipe_id
            # Skip None/empty recipe_ids — they break repetition runs
            if not cur_id or not start_id or cur_id != start_id:
                run_length = i - run_start
                if run_length >= self.layout_repetition_run:
                    issues.append(
                        DetectedIssue(
                            kind=IssueKind.LAYOUT_REPETITION,
                            severity=IssueSeverity.WARNING,
                            node_id=None,
                            message=(
                                f"{run_length} consecutive slides use recipe "
                                f"'{plans[run_start].recipe_id}' (slides {run_start}-{i - 1})"
                            ),
                            details={
                                "recipe_id": plans[run_start].recipe_id,
                                "run_length": run_length,
                                "start_slide": run_start,
                                "end_slide": i - 1,
                            },
                        )
                    )
                run_start = i
        # Check final run
        run_length = len(plans) - run_start
        if run_length >= self.layout_repetition_run:
            issues.append(
                DetectedIssue(
                    kind=IssueKind.LAYOUT_REPETITION,
                    severity=IssueSeverity.WARNING,
                    node_id=None,
                    message=(
                        f"{run_length} consecutive slides use recipe "
                        f"'{plans[run_start].recipe_id}' (slides {run_start}-{len(plans) - 1})"
                    ),
                    details={
                        "recipe_id": plans[run_start].recipe_id,
                        "run_length": run_length,
                        "start_slide": run_start,
                        "end_slide": len(plans) - 1,
                    },
                )
            )
        return issues

    def _check_missing_section_break(self, plans: list[LayoutPlan]) -> list[DetectedIssue]:
        """Check for missing section breaks in long decks."""
        if len(plans) <= self.section_break_min_slides:
            return []
        # Check if any slide has a role indicating a section break
        section_roles = {"section", "cover", "toc", "section_divider", "divider"}
        has_section = any(
            any(n.role in section_roles for n in plan.nodes)
            for plan in plans
        )
        if not has_section:
            return [
                DetectedIssue(
                    kind=IssueKind.MISSING_SECTION_BREAK,
                    severity=IssueSeverity.INFO,
                    node_id=None,
                    message=(
                        f"Deck has {len(plans)} slides but no section/cover/toc dividers"
                    ),
                    details={"slide_count": len(plans)},
                )
            ]
        return []

    def _check_color_palette_drift(self, plans: list[LayoutPlan]) -> list[DetectedIssue]:
        """Check if primary colors vary across slides."""
        # Collect primary (most-used) colors per slide
        slide_primary_colors: list[tuple[int, str]] = []  # (slide_idx, color)
        for idx, plan in enumerate(plans):
            color_counts: dict[str, int] = {}
            for node in plan.nodes:
                style = node.resolved_style or {}
                color = style.get("color")
                if color and isinstance(color, str) and color.startswith("#"):
                    # Normalize to uppercase 6-digit
                    normalized = color.lstrip("#").upper()
                    if len(normalized) == 3:
                        normalized = "".join(c * 2 for c in normalized)
                    if len(normalized) == 6:
                        color_counts[f"#{normalized}"] = color_counts.get(f"#{normalized}", 0) + 1
            if color_counts:
                primary = max(color_counts, key=color_counts.get)  # type: ignore[arg-type]
                slide_primary_colors.append((idx, primary))

        if len(slide_primary_colors) < 2:
            return []

        # Group colors by perceptual similarity (RGB distance ≤5 per channel)
        def _colors_equivalent(c1: str, c2: str) -> bool:
            """Treat two hex colors as equivalent if all RGB channels differ by ≤5."""
            try:
                r1, g1, b1 = _hex_to_rgb(c1)
                r2, g2, b2 = _hex_to_rgb(c2)
                return abs(r1 - r2) <= 5 and abs(g1 - g2) <= 5 and abs(b1 - b2) <= 5
            except (ValueError, TypeError):
                return c1 == c2

        # Map each color to a canonical representative
        canonical_map: dict[str, str] = {}
        representatives: list[str] = []
        for _, color in slide_primary_colors:
            if color in canonical_map:
                continue
            matched = None
            for rep in representatives:
                if _colors_equivalent(color, rep):
                    matched = rep
                    break
            if matched is not None:
                canonical_map[color] = matched
            else:
                representatives.append(color)
                canonical_map[color] = color

        # Use canonical colors for drift detection
        canonical_colors = [canonical_map[c] for _, c in slide_primary_colors]
        unique_colors = set(canonical_colors)
        if len(unique_colors) <= 1:
            return []

        # Simple heuristic: if more than half the slides use a different primary, warn
        color_counter = Counter(canonical_colors)
        most_common_color, most_common_count = color_counter.most_common(1)[0]
        slides_with_different = len(slide_primary_colors) - most_common_count

        if slides_with_different > len(slide_primary_colors) // 2:
            return [
                DetectedIssue(
                    kind=IssueKind.COLOR_PALETTE_DRIFT,
                    severity=IssueSeverity.WARNING,
                    node_id=None,
                    message=(
                        f"Primary text color varies across slides: "
                        f"{len(unique_colors)} different colors used"
                    ),
                    details={
                        "unique_colors": sorted(unique_colors),
                        "most_common": most_common_color,
                        "slides_with_different": slides_with_different,
                    },
                )
            ]
        return []


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
