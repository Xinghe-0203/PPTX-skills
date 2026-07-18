"""Visual QA result models (PR1 data contract).

These classes describe QA outcomes, issues and evidence. They do not perform
any rendering or image processing themselves.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from pptx_skill.content_model import BBox


class Severity(str, Enum):
    INFO = "info"
    WARNING = "warning"
    BLOCKER = "blocker"


class CheckOutcome(str, Enum):
    PASS = "pass"
    FAIL = "fail"
    INCONCLUSIVE = "inconclusive"


class PresentationQualityError(Exception):
    """Raised in strict QA mode when blockers cannot be repaired."""

    def __init__(self, report: "QAReport", message: str = "Presentation quality check failed"):
        super().__init__(message)
        self.report = report


@dataclass
class QACheckResult:
    check_id: str
    outcome: CheckOutcome
    confidence: float
    evidence: dict[str, Any] = field(default_factory=dict)
    issue_codes: list[str] = field(default_factory=list)


@dataclass
class QAIssue:
    code: str
    severity: Severity
    slide_index: int
    element_id: str | None = None
    render_node_id: str | None = None
    ppt_shape_id: int | None = None
    bbox: BBox | None = None
    message: str = ""
    evidence: dict[str, Any] = field(default_factory=dict)
    suggested_repairs: list[str] = field(default_factory=list)
    confidence: float = 1.0


@dataclass
class QAReport:
    status: CheckOutcome
    checks: list[QACheckResult] = field(default_factory=list)
    issues: list[QAIssue] = field(default_factory=list)
    render_result: "RenderResult | None" = None
    metrics: dict[str, Any] = field(default_factory=dict)
    artifacts: dict[str, str] = field(default_factory=dict)

    def blockers(self) -> list[QAIssue]:
        return [i for i in self.issues if i.severity == Severity.BLOCKER]

    def warnings(self) -> list[QAIssue]:
        return [i for i in self.issues if i.severity == Severity.WARNING]

    def infos(self) -> list[QAIssue]:
        return [i for i in self.issues if i.severity == Severity.INFO]
