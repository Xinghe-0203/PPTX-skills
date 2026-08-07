"""Rendered image QA (PR5).

Compares rendered PNGs against a baseline or reference deck using perceptual
pixel diff, windowed diff density and SSIM. Falls back to Pillow-only pixel
operations when numpy/scikit-image are unavailable.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops, ImageDraw

from pptx_skill._compat import StrEnum
from pptx_skill.visual_qa import CheckOutcome, QACheckResult, QAIssue, QAReport, Severity


class ImageBackend(StrEnum):
    SCIKIT_IMAGE = "scikit-image"
    PILLOW = "Pillow"
    UNAVAILABLE = "unavailable"


@dataclass
class RenderQAConfig:
    same_engine: dict[str, float] = field(default_factory=lambda: {
        "perceptual_threshold": 0.10,
        "max_window_diff_density": 0.015,
        "min_ssim": 0.985,
    })
    reference_clone: dict[str, float] = field(default_factory=lambda: {
        "max_window_diff_density": 0.08,
        "min_ssim": 0.94,
    })
    window_size: int = 16
    ignore_edges_px: int = 2


@dataclass
class SlideDiff:
    slide_index: int
    perceptual_diff_ratio: float
    window_diff_density: float
    ssim: float | None
    backend: ImageBackend
    diff_png_path: str | None = None
    annotated_png_path: str | None = None


def _detect_backend() -> ImageBackend:
    try:
        import numpy  # noqa: F401
        import skimage.metrics  # noqa: F401
        return ImageBackend.SCIKIT_IMAGE
    except Exception:
        return ImageBackend.PILLOW


def _load_image(path: str) -> Image.Image:
    fp = Image.open(path)
    img = fp.convert("RGB")
    fp.close()
    return img


def _same_size(img_a: Image.Image, img_b: Image.Image) -> bool:
    return img_a.size == img_b.size


def _resize_to_reference(
    img: Image.Image,
    target_size: tuple[int, int],
) -> Image.Image:
    return img.resize(target_size, Image.Resampling.LANCZOS)


def _perceptual_diff_pillow(
    a: Image.Image,
    b: Image.Image,
    threshold: int = 16,
    ignore_edges_px: int = 2,
) -> tuple[float, Image.Image]:
    """Return (diff_ratio, diff_mask) where diff_mask is white on differences."""
    width, height = a.size
    diff = ImageChops.difference(a, b)
    # Convert to grayscale and threshold.
    gray = diff.convert("L")
    # Ignore edges by drawing a black border.
    if ignore_edges_px > 0:
        draw = ImageDraw.Draw(gray)
        draw.rectangle([0, 0, width - 1, ignore_edges_px - 1], fill=0)
        draw.rectangle([0, height - ignore_edges_px, width - 1, height - 1], fill=0)
        draw.rectangle([0, 0, ignore_edges_px - 1, height - 1], fill=0)
        draw.rectangle([width - ignore_edges_px, 0, width - 1, height - 1], fill=0)
    # Pixels above threshold are considered different.
    mask = gray.point(lambda v: 255 if v > threshold else 0)
    data = list(mask.getdata() if not hasattr(mask, "get_flattened_data") else mask.get_flattened_data())
    diff_count = sum(1 for v in data if v > 0)
    total = len(data)
    ratio = diff_count / total if total > 0 else 0.0
    return ratio, mask


def _windowed_density_pillow(mask: Image.Image, window_size: int = 16) -> float:
    width, height = mask.size
    if width <= 0 or height <= 0 or window_size <= 0:
        return 0.0
    data = list(mask.getdata() if not hasattr(mask, "get_flattened_data") else mask.get_flattened_data())
    max_density = 0.0
    for y in range(0, height - window_size + 1, window_size):
        for x in range(0, width - window_size + 1, window_size):
            count = 0
            total = 0
            for dy in range(window_size):
                row_start = (y + dy) * width + x
                for dx in range(window_size):
                    if data[row_start + dx] > 0:
                        count += 1
                    total += 1
            density = count / total if total > 0 else 0.0
            if density > max_density:
                max_density = density
    return max_density


def _ssim_scikit(a: Image.Image, b: Image.Image) -> float:
    try:
        import numpy as np
        from skimage.metrics import structural_similarity as ssim
    except ImportError:
        return 0.0

    arr_a = np.array(a)
    arr_b = np.array(b)
    if arr_a.shape != arr_b.shape:
        raise ValueError("Images must have the same dimensions for SSIM")
    score, _ = ssim(arr_a, arr_b, channel_axis=2, full=True)
    return float(score)


def _annotated_diff(
    a: Image.Image,
    mask: Image.Image,
    issues: list[QAIssue],
) -> Image.Image:
    """Overlay red diff mask and issue bounding boxes on image A."""
    annotated = a.copy()
    # Overlay red mask at 40% opacity.
    red_mask = Image.new("RGB", a.size, (255, 0, 0))
    annotated = Image.blend(annotated, red_mask, alpha=0.4)
    draw = ImageDraw.Draw(annotated)
    for issue in issues:
        if issue.bbox is None:
            continue
        bbox = issue.bbox
        # Assume bbox is in points; for annotation we just use relative coords
        # if within image bounds, otherwise skip precise box.
        x1, y1 = int(bbox.x), int(bbox.y)
        x2, y2 = int(bbox.x + bbox.width), int(bbox.y + bbox.height)
        if x2 <= x1 or y2 <= y1:
            continue
        color = {
            Severity.BLOCKER: (255, 0, 0),
            Severity.WARNING: (255, 165, 0),
            Severity.INFO: (0, 112, 192),
        }.get(issue.severity, (255, 255, 255))
        draw.rectangle([x1, y1, x2, y2], outline=color, width=2)
        draw.text((x1 + 2, y1 - 12), issue.code, fill=color)
    return annotated


def compare_slide_to_baseline(
    slide_index: int,
    rendered_png: str,
    baseline_png: str,
    output_dir: str | None = None,
    config: RenderQAConfig | None = None,
) -> SlideDiff:
    """Compare a rendered slide to a baseline PNG."""
    cfg = config or RenderQAConfig()
    backend = _detect_backend()

    img_rendered = _load_image(rendered_png)
    img_baseline = _load_image(baseline_png)

    if not _same_size(img_rendered, img_baseline):
        img_rendered = _resize_to_reference(img_rendered, img_baseline.size)

    diff_ratio, mask = _perceptual_diff_pillow(
        img_rendered,
        img_baseline,
        ignore_edges_px=cfg.ignore_edges_px,
    )
    window_density = _windowed_density_pillow(mask, window_size=cfg.window_size)

    ssim_score: float | None = None
    if backend == ImageBackend.SCIKIT_IMAGE:
        try:
            ssim_score = _ssim_scikit(img_rendered, img_baseline)
        except Exception:
            backend = ImageBackend.PILLOW

    diff_path: str | None = None
    annotated_path: str | None = None
    if output_dir:
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        diff_path = str(out / f"slide_{slide_index:03d}_diff.png")
        mask.save(diff_path)
        red_mask = Image.new("RGB", img_rendered.size, (255, 0, 0))
        red_channel = red_mask.split()[0]
        r, g, b = img_rendered.split()
        mask_resized = mask.resize(img_rendered.size, Image.Resampling.NEAREST)
        mask_gray = mask_resized.convert("L")
        r_masked = Image.composite(red_channel, r, mask_gray)
        annotated = Image.merge("RGB", (r_masked, g, b))
        annotated_path = str(out / f"slide_{slide_index:03d}_annotated.png")
        annotated.save(annotated_path)

    return SlideDiff(
        slide_index=slide_index,
        perceptual_diff_ratio=diff_ratio,
        window_diff_density=window_density,
        ssim=ssim_score,
        backend=backend,
        diff_png_path=diff_path,
        annotated_png_path=annotated_path,
    )


def evaluate_render_against_baseline(
    rendered_pngs: list[str],
    baseline_pngs: list[str],
    output_dir: str | None = None,
    config: RenderQAConfig | None = None,
    reference_mode: bool = False,
) -> QAReport:
    """Compare all rendered slides to baseline/reference images and return a QAReport."""
    cfg = config or RenderQAConfig()
    thresholds = cfg.reference_clone if reference_mode else cfg.same_engine
    issues: list[QAIssue] = []
    checks: list[QACheckResult] = []
    metrics: dict[str, Any] = {"backend": _detect_backend().value}

    if len(rendered_pngs) != len(baseline_pngs):
        issues.append(
            QAIssue(
                code="RENDER_BASELINE_COUNT_MISMATCH",
                severity=Severity.BLOCKER,
                slide_index=-1,
                element_id=None,
                render_node_id=None,
                ppt_shape_id=None,
                bbox=None,
                message=f"Rendered {len(rendered_pngs)} slides but baseline has {len(baseline_pngs)}",
                evidence={"rendered_count": len(rendered_pngs), "baseline_count": len(baseline_pngs)},
                suggested_repairs=[],
                confidence=1.0,
            )
        )
        return QAReport(
            status=CheckOutcome.FAIL,
            checks=[QACheckResult(check_id="render_qa", outcome=CheckOutcome.FAIL, confidence=1.0, evidence={}, issue_codes=["RENDER_BASELINE_COUNT_MISMATCH"])],
            issues=issues,
            render_result=None,
            metrics=metrics,
            artifacts={},
        )

    diffs: list[SlideDiff] = []
    for idx, (rendered, baseline) in enumerate(zip(rendered_pngs, baseline_pngs, strict=False)):
        diff = compare_slide_to_baseline(idx, rendered, baseline, output_dir=output_dir, config=cfg)
        diffs.append(diff)

    status = CheckOutcome.PASS
    for diff in diffs:
        evidence = {
            "perceptual_diff_ratio": diff.perceptual_diff_ratio,
            "window_diff_density": diff.window_diff_density,
            "ssim": diff.ssim,
            "backend": diff.backend.value,
        }
        checks.append(
            QACheckResult(
                check_id=f"render_slide_{diff.slide_index}",
                outcome=CheckOutcome.PASS,
                confidence=1.0,
                evidence=evidence,
                issue_codes=[],
            )
        )

        if diff.window_diff_density > thresholds["max_window_diff_density"]:
            status = CheckOutcome.FAIL
            issues.append(
                QAIssue(
                    code="VISUAL_REGRESSION",
                    severity=Severity.BLOCKER,
                    slide_index=diff.slide_index,
                    element_id=None,
                    render_node_id=None,
                    ppt_shape_id=None,
                    bbox=None,
                    message=(
                        f"Window diff density {diff.window_diff_density:.4f} exceeds "
                        f"threshold {thresholds['max_window_diff_density']}"
                    ),
                    evidence=evidence,
                    suggested_repairs=["inspect_rendered_changes", "update_baseline_if_intentional"],
                    confidence=0.95,
                )
            )
        if diff.perceptual_diff_ratio > thresholds["perceptual_threshold"] and not reference_mode:
            status = CheckOutcome.FAIL
            issues.append(
                QAIssue(
                    code="VISUAL_REGRESSION",
                    severity=Severity.WARNING,
                    slide_index=diff.slide_index,
                    element_id=None,
                    render_node_id=None,
                    ppt_shape_id=None,
                    bbox=None,
                    message=(
                        f"Perceptual diff ratio {diff.perceptual_diff_ratio:.4f} exceeds "
                        f"threshold {thresholds['perceptual_threshold']}"
                    ),
                    evidence=evidence,
                    suggested_repairs=["inspect_rendered_changes"],
                    confidence=0.8,
                )
            )

        if diff.ssim is not None and diff.ssim < thresholds["min_ssim"]:
            status = CheckOutcome.FAIL
            issues.append(
                QAIssue(
                    code="STRUCTURAL_REGRESSION",
                    severity=Severity.WARNING,
                    slide_index=diff.slide_index,
                    element_id=None,
                    render_node_id=None,
                    ppt_shape_id=None,
                    bbox=None,
                    message=f"SSIM {diff.ssim:.4f} below threshold {thresholds['min_ssim']}",
                    evidence=evidence,
                    suggested_repairs=["inspect_rendered_changes"],
                    confidence=0.85,
                )
            )

    metrics.update(
        {
            "max_perceptual_diff_ratio": max(d.perceptual_diff_ratio for d in diffs) if diffs else 0.0,
            "max_window_diff_density": max(d.window_diff_density for d in diffs) if diffs else 0.0,
            "min_ssim": min((d.ssim for d in diffs if d.ssim is not None), default=None),
        }
    )

    artifacts = {}
    if output_dir:
        artifacts["diff_dir"] = output_dir

    return QAReport(
        status=status,
        checks=checks,
        issues=issues,
        render_result=None,
        metrics=metrics,
        artifacts=artifacts,
    )
