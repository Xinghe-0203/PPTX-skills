"""Generation pipeline orchestrator (PR5).

Wires adaptive layout planning, PPTX rendering, preview rendering, QA and a
limited repair loop together. Each pass is recorded in Manifest V3.
"""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx_skill.content_model import CanvasSpec, ContentSpec, LayoutPlan, SlideSpec
from pptx_skill.deck_planner import LayoutScoringConfig, plan_deck
from pptx_skill.pptx_renderer import render_layout_plans
from pptx_skill.preview_renderer import render_preview
from pptx_skill.render_qa import RenderQAConfig, evaluate_render_against_baseline
from pptx_skill.repair_engine import apply_repairs, merge_profile_overrides, propose_repairs
from pptx_skill.semantic_qa import SemanticQAEngine
from pptx_skill.visual_qa import CheckOutcome, PresentationQualityError, QACheckResult, QAReport


class LayoutPlanningError(Exception):
    pass


@dataclass
class GenerationResult:
    pptx_path: str
    manifest_path: str | None = None
    qa_report_path: str | None = None
    qa_status: str | None = None
    render_trace_path: str | None = None
    preview_dir: str | None = None
    repair_passes: int = 0
    repair_log: list[dict] = field(default_factory=list)
    manifest: Any | None = None


def plan_deck_layouts(
    slide_specs: list[SlideSpec],
    canvas: CanvasSpec,
    profile_state: dict[str, Any],
    config: LayoutScoringConfig | None = None,
) -> list[LayoutPlan]:
    """Plan layouts for every slide using deck-level beam search.

    Each slide spec in ``slide_specs`` may correspond to one or more derived
    slides after pagination, so the returned list of LayoutPlans is keyed
    against the *derived* slides rather than the source slides.
    """
    deck = plan_deck(slide_specs, canvas, profile_state, config=config)
    return deck.plans


def run_generation_pipeline(
    content: ContentSpec,
    output_path: str,
    canvas: CanvasSpec | None = None,
    profile_state: dict[str, Any] | None = None,
    qa_mode: str = "report",
    auto_repair: bool = True,
    max_repair_passes: int = 2,
    renderer_engine: str = "auto",
    target_dpi: int = 150,
    reference_baseline_pngs: list[str] | None = None,
) -> GenerationResult:
    """Run the full adaptive generation pipeline with optional QA/repair."""
    canvas = canvas or CanvasSpec(959.976, 540)
    profile_state = profile_state or {}
    tokens = profile_state.get("tokens", {})
    if not tokens:
        from pptx_skill.layout_engine import _builtin_tokens
        tokens = _builtin_tokens()
        profile_state["tokens"] = tokens

    slide_specs = list(content.slides)
    repair_log: list[dict] = []
    qa_report: QAReport | None = None
    final_pptx_path = output_path
    manifest_sidecar: str | None = None

    repair_budget = max_repair_passes if qa_mode != "off" and auto_repair else 0

    for pass_index in range(repair_budget + 1):
        plans = plan_deck_layouts(slide_specs, canvas, profile_state)

        # Validate plans with semantic QA.
        engine = SemanticQAEngine()
        for plan in plans:
            report = engine.check(plan)
            if not report.passed:
                if pass_index == repair_budget:
                    if qa_mode == "strict":
                        raise PresentationQualityError(report)
                    break
                actions = propose_repairs(report, slide_specs, profile_state)
                if not actions:
                    if qa_mode == "strict":
                        raise PresentationQualityError(report)
                    break
                slide_specs, overrides = apply_repairs(actions, slide_specs, profile_state)
                profile_state = merge_profile_overrides(profile_state, overrides)
                repair_log.append({
                    "pass": pass_index,
                    "trigger": "semantic_qa",
                    "actions": [{"kind": a.kind, "slide": a.slide_index, "element": a.element_id} for a in actions],
                })
                break
        else:
            # All plans passed semantic QA; render one slide per derived plan.
            render_layout_plans(plans, final_pptx_path)

            from pptx_skill.manifest import ManifestV3, save_manifest_v3
            manifest = ManifestV3()
            manifest.record_attempt(
                pass_index=pass_index,
                plan_artifact="layout_plans",
                trace_artifact="render_trace",
            )
            manifest.current["content"] = {"slides": len(plans)}
            save_manifest_v3(final_pptx_path, manifest)
            manifest_sidecar = str(Path(final_pptx_path).with_suffix(".manifest.json"))

            if qa_mode == "off":
                return GenerationResult(
                    pptx_path=final_pptx_path,
                    manifest_path=manifest_sidecar,
                    qa_report_path=None,
                    qa_status="pass",
                    preview_dir=None,
                    repair_passes=pass_index,
                    repair_log=repair_log,
                )

            preview_dir: str | None = None
            try:
                preview_dir = str(Path(output_path).parent / "preview")
                Path(preview_dir).mkdir(parents=True, exist_ok=True)
                preview = render_preview(final_pptx_path, output_dir=preview_dir, engine=renderer_engine, dpi=target_dpi)
                for idx, png_path in enumerate(preview.slide_pngs):
                    dest = Path(preview_dir) / f"slide_{idx:03d}.png"
                    shutil.copy(png_path, dest)
            except Exception as exc:
                qa_report = QAReport(
                    status=CheckOutcome.INCONCLUSIVE,
                    checks=[QACheckResult(check_id="render_preview", outcome=CheckOutcome.INCONCLUSIVE, confidence=0.5, evidence={"error": str(exc)}, issue_codes=[])],
                    issues=[],
                    render_result=None,
                    metrics={},
                    artifacts={},
                )
                if qa_mode == "strict":
                    raise PresentationQualityError(qa_report) from exc
                break

            if reference_baseline_pngs:
                diff_dir = str(Path(output_path).parent / "diff")
                qa_report = evaluate_render_against_baseline(
                    preview.slide_pngs,
                    reference_baseline_pngs,
                    output_dir=diff_dir,
                    config=RenderQAConfig(),
                    reference_mode=False,
                )
            else:
                # Without baseline, run semantic QA on rendered plans as report.
                issues = []
                checks = []
                for idx, plan in enumerate(plans):
                    report = engine.check(plan)
                    issues.extend(report.issues)
                    checks.append(QACheckResult(check_id=f"semantic_slide_{idx}", outcome=CheckOutcome.PASS if report.passed else CheckOutcome.FAIL, confidence=1.0, evidence={"issues": len(report.issues)}, issue_codes=[i.kind for i in report.issues]))
                status = CheckOutcome.PASS if all(i.severity.value != "blocker" for i in issues) else CheckOutcome.FAIL
                qa_report = QAReport(status=status, checks=checks, issues=issues, render_result=None, metrics={}, artifacts={"preview_dir": preview_dir} if preview_dir else {})

            if qa_report.status == CheckOutcome.PASS:
                return GenerationResult(
                    pptx_path=final_pptx_path,
                    manifest_path=manifest_sidecar,
                    qa_report_path=None,
                    qa_status="pass",
                    preview_dir=preview_dir,
                    repair_passes=pass_index,
                    repair_log=repair_log,
                )

            if pass_index == repair_budget:
                if qa_mode == "strict":
                    raise PresentationQualityError(qa_report)
                break

            actions = propose_repairs(qa_report, slide_specs, profile_state)
            if not actions:
                if qa_mode == "strict":
                    raise PresentationQualityError(qa_report)
                break
            slide_specs, overrides = apply_repairs(actions, slide_specs, profile_state)
            profile_state = merge_profile_overrides(profile_state, overrides)
            repair_log.append({
                "pass": pass_index,
                "trigger": "render_qa",
                "actions": [{"kind": a.kind, "slide": a.slide_index, "element": a.element_id} for a in actions],
            })
            continue

    # Exited without passing.
    status = "fail" if qa_report and qa_report.status == CheckOutcome.FAIL else "inconclusive"
    return GenerationResult(
        pptx_path=final_pptx_path,
        manifest_path=manifest_sidecar,
        qa_report_path=None,
        qa_status=status,
        preview_dir=None,
        repair_passes=len(repair_log),
        repair_log=repair_log,
    )
