"""Repair engine (PR5).

Produces whitelist RepairAction objects from a QA report and applies them to
slide specs or profile state. Only safe, reversible actions are applied
automatically.
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any

from pptx_skill.content_model import SlideSpec


class RepairActionKind(StrEnum):
    REDUCE_FONT_WITHIN_LIMIT = "reduce_font_within_limit"
    EXPAND_ZONE_WITHIN_RECIPE = "expand_zone_within_recipe"
    SWITCH_LAYOUT_CANDIDATE = "switch_layout_candidate"
    ADJUST_IMAGE_FOCUS = "adjust_image_focus"
    CHANGE_TEXT_COLOR_TO_TOKEN = "change_text_color_to_token"
    REMOVE_EMPTY_PLACEHOLDER = "remove_empty_placeholder"
    SPLIT_REPEATED_CONTENT = "split_repeated_content"  # PR6 only


@dataclass
class RepairAction:
    kind: RepairActionKind
    slide_index: int
    element_id: str | None
    render_node_id: str | None
    reason: str
    from_value: Any = None
    to_value: Any = None
    details: dict = field(default_factory=dict)


@dataclass
class _IssueView:
    code: str
    severity_value: str
    slide_index: int
    element_id: str | None
    render_node_id: str | None
    details: dict[str, Any]


def _view_issues(report: Any) -> list[_IssueView]:
    """Normalize SemanticQAReport or visual QAReport issues."""
    views: list[_IssueView] = []
    if not hasattr(report, "issues"):
        return views
    issues = report.issues

    for issue in issues:
        if hasattr(issue, "code"):
            # visual QAReport QAIssue
            views.append(
                _IssueView(
                    code=str(issue.code),
                    severity_value=issue.severity.value if hasattr(issue.severity, "value") else str(issue.severity),
                    slide_index=issue.slide_index,
                    element_id=issue.element_id,
                    render_node_id=issue.render_node_id,
                    details=issue.evidence or {},
                )
            )
        else:
            # SemanticQAReport DetectedIssue
            severity_value = issue.severity.value if hasattr(issue.severity, "value") else str(issue.severity)
            slide_index = issue.slide_index
            element_id = issue.node_id
            # node_id may be "slide/role"; try to extract element id from it.
            if element_id and "/" in element_id:
                element_id = element_id.split("/")[-1]
            views.append(
                _IssueView(
                    code=str(issue.kind.value) if hasattr(issue.kind, "value") else str(issue.kind),
                    severity_value=severity_value,
                    slide_index=slide_index,
                    element_id=element_id,
                    render_node_id=issue.node_id,
                    details=issue.details or {},
                )
            )
    return views


def propose_repairs(
    qa_report: Any,
    slide_specs: list[SlideSpec],
    profile_state: dict[str, Any],
) -> list[RepairAction]:
    """Propose whitelist repair actions from a QA report."""
    actions: list[RepairAction] = []
    issues = _view_issues(qa_report)

    text_overflow_blockers = [
        i for i in issues
        if i.code in {"text_overflow", "TEXT_OVERFLOW_ESTIMATED", "TEXT_OVERFLOW_CONFIRMED"}
        and i.severity_value == "blocker"
    ]
    low_contrast_warnings = [i for i in issues if i.code in {"low_contrast", "LOW_CONTRAST"}]
    empty_placeholders = [i for i in issues if i.code in {"empty_content", "EMPTY_PLACEHOLDER"}]

    for issue in text_overflow_blockers:
        slide_index = max(issue.slide_index, 0)
        element_id = issue.element_id
        if slide_index >= len(slide_specs):
            continue
        slide = slide_specs[slide_index]
        element = next((e for e in slide.elements if e.id == element_id or e.role == element_id), None)
        if element is None:
            # Fall back to switching layout candidate.
            actions.append(
                RepairAction(
                    kind=RepairActionKind.SWITCH_LAYOUT_CANDIDATE,
                    slide_index=slide_index,
                    element_id=element_id,
                    render_node_id=issue.render_node_id,
                    reason="Text overflow on unknown element; try next candidate",
                    from_value=slide.preferred_layouts[0] if slide.preferred_layouts else None,
                    to_value="next_candidate",
                )
            )
            continue
        # If element has text content, try reducing font first.
        if element.content.get("text"):
            required = issue.details.get("required_font_size_pt")
            current = element.content.get(
                "font_size_pt",
                profile_state.get("qa", {}).get("min_body_font_size", 16),
            )
            min_size = profile_state.get("qa", {}).get("min_body_font_size", 9)
            if required is not None and required < current:
                to_value = max(min_size, required - 0.5)
            else:
                to_value = max(min_size, current - 1)
            actions.append(
                RepairAction(
                    kind=RepairActionKind.REDUCE_FONT_WITHIN_LIMIT,
                    slide_index=slide_index,
                    element_id=element.id,
                    render_node_id=issue.render_node_id,
                    reason="Text overflow detected",
                    from_value=current,
                    to_value=to_value,
                    details={"issue": str(issue.details), "required_font_size": required},
                )
            )
        else:
            actions.append(
                RepairAction(
                    kind=RepairActionKind.SWITCH_LAYOUT_CANDIDATE,
                    slide_index=slide_index,
                    element_id=element.id,
                    render_node_id=issue.render_node_id,
                    reason="Text overflow on element without text",
                    from_value=slide.preferred_layouts[0] if slide.preferred_layouts else None,
                    to_value="next_candidate",
                )
            )

    for issue in low_contrast_warnings:
        slide_index = issue.slide_index if issue.slide_index >= 0 else 0
        if slide_index >= len(slide_specs):
            continue
        element_id = issue.element_id
        actions.append(
            RepairAction(
                kind=RepairActionKind.CHANGE_TEXT_COLOR_TO_TOKEN,
                slide_index=slide_index,
                element_id=element_id,
                render_node_id=issue.render_node_id,
                reason="Low contrast",
                from_value=issue.details.get("foreground"),
                to_value=profile_state.get("tokens", {}).get("semantic", {}).get("color", {}).get("text_primary", "#1A1A1A"),
                details={"background": issue.details.get("background"), "ratio": issue.details.get("ratio")},
            )
        )

    for issue in empty_placeholders:
        slide_index = issue.slide_index if issue.slide_index >= 0 else 0
        if slide_index >= len(slide_specs):
            continue
        element_id = issue.element_id
        if element_id:
            actions.append(
                RepairAction(
                    kind=RepairActionKind.REMOVE_EMPTY_PLACEHOLDER,
                    slide_index=slide_index,
                    element_id=element_id,
                    render_node_id=issue.render_node_id,
                    reason="Empty placeholder detected",
                )
            )

    return actions


def apply_repairs(
    actions: list[RepairAction],
    slide_specs: list[SlideSpec],
    profile_state: dict[str, Any],
) -> tuple[list[SlideSpec], dict[str, Any]]:
    """Apply whitelist actions to slide specs and profile overrides."""
    new_slides: list[SlideSpec] = []
    for slide in slide_specs:
        new_elements = [copy.deepcopy(e) for e in slide.elements]
        new_slides.append(
            SlideSpec(
                id=slide.id,
                role=slide.role,
                communication_goal=slide.communication_goal,
                elements=new_elements,
                preferred_layouts=list(slide.preferred_layouts),
                source_section_id=slide.source_section_id,
                fragment_index=slide.fragment_index,
            )
        )

    profile_overrides: dict[str, Any] = {}
    applied_layout_switch: set[int] = set()

    for action in actions:
        if action.slide_index < 0 or action.slide_index >= len(new_slides):
            continue
        slide = new_slides[action.slide_index]

        if action.kind == RepairActionKind.REDUCE_FONT_WITHIN_LIMIT:
            element = next((e for e in slide.elements if e.id == action.element_id), None)
            if element is not None:
                current = element.content.get("font_size_pt")
                if current is None:
                    current = profile_state.get("qa", {}).get("min_body_font_size", 16)
                new_size = action.to_value
                if new_size is not None and isinstance(new_size, (int, float)):
                    element.content["font_size_pt"] = float(new_size)
                elif new_size is None:
                    element.content["font_size_pt"] = max(
                        profile_state.get("qa", {}).get("min_body_font_size", 9),
                        current - 1,
                    )
                if action.render_node_id:
                    element.content["_repair_render_node_id"] = action.render_node_id

        elif action.kind == RepairActionKind.SWITCH_LAYOUT_CANDIDATE:
            if action.slide_index in applied_layout_switch:
                continue
            applied_layout_switch.add(action.slide_index)
            # Rotate preferred layouts so the next candidate is tried first.
            if len(slide.preferred_layouts) > 1:
                first = slide.preferred_layouts.pop(0)
                slide.preferred_layouts.append(first)
            elif slide.preferred_layouts:
                # If only one preferred layout, mark override to choose fallback.
                profile_overrides.setdefault("force_next_candidate", set()).add(action.slide_index)

        elif action.kind == RepairActionKind.CHANGE_TEXT_COLOR_TO_TOKEN:
            element = next((e for e in slide.elements if e.id == action.element_id), None)
            if element is not None:
                element.content["color"] = action.to_value

        elif action.kind == RepairActionKind.REMOVE_EMPTY_PLACEHOLDER:
            slide.elements = [e for e in slide.elements if e.id != action.element_id]

    return new_slides, profile_overrides


def merge_profile_overrides(
    profile_state: dict[str, Any],
    overrides: dict[str, Any],
) -> dict[str, Any]:
    """Merge repair overrides into profile state immutably."""
    merged = copy.deepcopy(profile_state)
    merged.setdefault("repair_overrides", []).append(overrides)
    return merged
