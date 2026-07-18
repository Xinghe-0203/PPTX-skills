"""Public API and compatibility facade for the PPTX skill.

This module preserves the historical function signatures of ``auto_generate_ppt``
and ``auto_validate_ppt`` while adding the new structured result types and
experimental switches for the V2 adaptive/QA pipeline.
"""
from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from pptx_skill.content_adapter import adapt_legacy_sections
from pptx_skill.manifest import ManifestV3, load_manifest, save_manifest_v3, set_current_content
from pptx_skill.visual_qa import CheckOutcome, QAReport, QACheckResult, Severity


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------

@dataclass
class GenerationResult:
    """Structured result returned when ``return_result=True``."""

    pptx_path: str
    manifest_path: str | None = None
    qa_report_path: str | None = None
    qa_status: CheckOutcome | None = None
    render_trace_path: str | None = None
    preview_dir: str | None = None
    repair_passes: int = 0
    manifest: ManifestV3 | None = None


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _ensure_dir(path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)


def _legacy_validate(pptx_path: str) -> dict[str, Any]:
    """Port the existing structural validation to the new outcome model."""
    from pptx import Presentation

    if not os.path.exists(pptx_path):
        return {"passed": False, "checks": [], "warnings": [f"文件不存在: {pptx_path}"]}

    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    n_slides = len(slides)
    checks: list[tuple[str, bool | None]] = []
    warnings: list[str] = []
    passed = True

    if n_slides >= 3:
        counts = [len(slide.shapes) for slide in slides]
        for i in range(2, n_slides):
            if counts[i] == counts[i - 1] == counts[i - 2]:
                warnings.append(
                    f"第{i-1}-{i+1}页shape数量相同({counts[i]})，可能版式重复"
                )
    checks.append(("相邻版式多样性", not any("版式重复" in w for w in warnings)))

    unique_counts = len({len(s.shapes) for s in slides})
    diversity_ratio = unique_counts / max(n_slides, 1)
    checks.append(("版式多样性", diversity_ratio >= 0.5))
    if diversity_ratio < 0.5:
        warnings.append(f"版式多样性不足: {unique_counts}/{n_slides}种不同的shape数量")
        passed = False

    if n_slides >= 1:
        cover_shapes = len(slides[0].shapes)
        checks.append(("封面有内容", cover_shapes >= 2))
        if cover_shapes < 2:
            warnings.append("封面内容过少")
            passed = False

    return {"passed": passed, "checks": checks, "warnings": warnings, "total_slides": n_slides}


def _qa_report_to_legacy(report: QAReport) -> dict[str, Any]:
    """Map a new QAReport back to the legacy validation dict shape."""
    blockers = [i for i in report.issues if i.severity == Severity.BLOCKER]
    warnings = [i for i in report.issues if i.severity == Severity.WARNING]
    return {
        "passed": report.status == CheckOutcome.PASS and not blockers,
        "checks": [(c.check_id, c.outcome == CheckOutcome.PASS) for c in report.checks],
        "warnings": [i.message for i in warnings + blockers],
        "total_slides": report.metrics.get("total_slides", 0),
        "_qa_report": report,
    }


def _build_qa_report(pptx_path: str, legacy_result: dict[str, Any] | None = None) -> QAReport:
    """Build a minimal QAReport from legacy structural validation.

    Real semantic/rendered checks are added in later PRs; this stub preserves
    the contract and lets callers switch to the new result type immediately.
    """
    if legacy_result is None:
        legacy_result = _legacy_validate(pptx_path)

    checks: list[QACheckResult] = []
    for name, ok in legacy_result.get("checks", []):
        checks.append(
            QACheckResult(
                check_id=name,
                outcome=CheckOutcome.PASS if ok else CheckOutcome.FAIL,
                confidence=1.0,
            )
        )

    issues = []
    for warning in legacy_result.get("warnings", []):
        issues.append(
            QAIssue(
                code="LEGACY_WARNING",
                severity=Severity.WARNING,
                slide_index=-1,
                message=warning,
            )
        )

    status = CheckOutcome.PASS if legacy_result.get("passed") else CheckOutcome.FAIL
    return QAReport(
        status=status,
        checks=checks,
        issues=issues,
        metrics={"total_slides": legacy_result.get("total_slides", 0)},
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def auto_generate_ppt(
    title,
    subtitle: str = "",
    sections: list[dict[str, Any] | Any] | None = None,
    output_path: str = "output.pptx",
    theme_key: str | None = None,
    image_dir: str = "./ppt_images",
    auto_search_images: bool = True,
    lang: str = "zh",
    template_key: str | None = None,
    template_profile: dict[str, Any] | None = None,
    template_catalog_path: str | None = None,
    layout_engine: Literal["legacy", "adaptive"] = "legacy",
    qa_mode: Literal["off", "report", "strict"] = "off",
    auto_repair: bool = True,
    max_repair_passes: int = 2,
    renderer_engine: Literal["auto", "libreoffice", "com"] = "auto",
    target_dpi: int = 150,
    reference_baseline: str | None = None,
    return_result: bool = False,
) -> str | GenerationResult:
    """Generate a PPTX. Default behavior is identical to the legacy function.

    New optional arguments are stubs for the V2 pipeline:

    - ``layout_engine="adaptive"`` is not implemented in PR1 and raises
      ``ValueError``.
    - ``qa_mode`` controls whether a QA report is produced; ``strict`` would
      block on blockers once real visual QA is wired in.
    - ``return_result=True`` returns a structured ``GenerationResult``.
    """
    if layout_engine not in ("legacy", "adaptive"):
        raise ValueError(f"Unknown layout_engine: {layout_engine}")
    if layout_engine == "adaptive":
        raise ValueError("layout_engine='adaptive' is not implemented in PR1")

    sections = sections or []
    abs_output = os.path.abspath(output_path)
    _ensure_dir(abs_output)

    # Legacy generation path.
    from pptx_helper import auto_generate_ppt as legacy_auto_generate_ppt

    pptx_path = legacy_auto_generate_ppt(
        title=title,
        subtitle=subtitle,
        sections=sections,
        output_path=abs_output,
        theme_key=theme_key,
        image_dir=image_dir,
        auto_search_images=auto_search_images,
        lang=lang,
        template_key=template_key,
        template_profile=template_profile,
        template_catalog_path=template_catalog_path,
    )

    # Build and persist V3 content model (best-effort; never breaks legacy path).
    try:
        content = adapt_legacy_sections(title, subtitle, sections, locale=lang or "zh-CN")
        manifest = load_manifest(pptx_path) or ManifestV3()
        set_current_content(manifest, content)
        manifest.legacy["theme_key"] = theme_key
        manifest.legacy["template_key"] = template_key
        save_manifest_v3(pptx_path, manifest)
    except Exception:
        # V3 manifest is optional in PR1; do not fail the legacy generation.
        manifest = None

    qa_report: QAReport | None = None
    if qa_mode != "off":
        qa_report = _build_qa_report(pptx_path)
        if qa_mode == "strict" and qa_report.status != CheckOutcome.PASS:
            # In PR1 strict only blocks on the legacy structural checks.
            from pptx_skill.visual_qa import PresentationQualityError

            raise PresentationQualityError(qa_report)

    if not return_result:
        return pptx_path

    return GenerationResult(
        pptx_path=pptx_path,
        manifest_path=_sidecar_path(pptx_path),
        qa_report_path=None,
        qa_status=qa_report.status if qa_report else None,
        render_trace_path=None,
        preview_dir=None,
        repair_passes=0,
        manifest=manifest,
    )


def auto_validate_ppt(pptx_path: str, return_report: bool = False) -> dict[str, Any] | QAReport:
    """Validate a PPTX.

    By default returns the legacy dict for backward compatibility. Pass
    ``return_report=True`` to receive the new ``QAReport``.
    """
    legacy = _legacy_validate(pptx_path)
    report = _build_qa_report(pptx_path, legacy)
    if return_report:
        return report
    return _qa_report_to_legacy(report)


def _sidecar_path(pptx_path: str) -> str:
    return str(Path(pptx_path).with_suffix(".manifest.json"))
