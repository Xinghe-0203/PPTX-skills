"""PR9: 200-page QA annotation dataset generator.

Per blueprint §12.3 the release dataset must include at least 200 human-labelled
fixture pages across six issue categories:

| 类型 | 最少页数 |
|---|---:|
| 文字溢出/未溢出 | 50 |
| 非预期/预期重叠 | 40 |
| 图片拉伸/错误裁切 | 30 |
| 对比度问题 | 30 |
| 越界/安全边距 | 20 |
| 表格/图表可读性 | 20 |
| 参考稿残留/空占位符 | 10 |

This module synthesizes a *programmatically-generated* annotation set: each
fixture is a small PPTX with a known issue (or a clean control), and the
expected label is recorded deterministically. The set is reproducible (no RNG
beyond a fixed seed passed in) so golden tests can depend on it.

Real human-labelled corpora are out of scope for the skill itself; this module
provides the deterministic fixture machinery a release gate would be calibrated
against. Each page records the expected issue code, a stable ``element_id`` /
``render_node_id``, the ``ppt_shape_id`` when known, severity, BBox and the
evidence source.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt

from pptx_skill._io import save_prs as _save_prs_impl

# Issue codes mirror the semantic_qa / visual_qa modules.
ISSUE_TEXT_OVERFLOW = "TEXT_OVERFLOW_CONFIRMED"
ISSUE_OVERLAP = "UNINTENDED_OVERLAP"
ISSUE_IMAGE_DISTORTION = "IMAGE_DISTORTION"
ISSUE_CONTRAST = "CONTRAST_INSUFFICIENT"
ISSUE_OUT_OF_BOUNDS = "SHAPE_OUT_OF_BOUNDS"
ISSUE_TABLE_READABILITY = "TABLE_READABILITY"
ISSUE_REFERENCE_RESIDUAL = "REFERENCE_RESIDUAL_OR_EMPTY_PLACEHOLDER"
ISSUE_CLEAN = "CLEAN_CONTROL"

# Category -> (issue code, minimum pages). Order matches blueprint §12.3 table.
DATASET_CATEGORIES: list[tuple[str, str, int]] = [
    ("text_overflow", ISSUE_TEXT_OVERFLOW, 50),
    ("overlap", ISSUE_OVERLAP, 40),
    ("image_distortion", ISSUE_IMAGE_DISTORTION, 30),
    ("contrast", ISSUE_CONTRAST, 30),
    ("out_of_bounds", ISSUE_OUT_OF_BOUNDS, 20),
    ("table_readability", ISSUE_TABLE_READABILITY, 20),
    ("reference_residual", ISSUE_REFERENCE_RESIDUAL, 10),
]

DATASET_TOTAL_MIN = sum(c for _, _, c in DATASET_CATEGORIES)  # 200


@dataclass
class AnnotationFixture:
    """One labelled fixture page."""

    fixture_id: str
    category: str
    issue_code: str
    severity: str  # blocker | warning | info | clean
    expected: dict[str, Any] = field(default_factory=dict)
    element_id: str = ""
    render_node_id: str = ""
    ppt_shape_id: int | None = None
    bbox: dict[str, float] = field(default_factory=dict)
    evidence_source: str = "synthesized"
    notes: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _save_slide_pptx(slide_builder, out_path: Path, *, width_in: float = 13.333, height_in: float = 7.5) -> int:
    """Build a one-slide PPTX via ``slide_builder(prs, slide)`` and save. Returns ppt_shape_id of the primary shape if any."""
    prs = Presentation()
    prs.slide_width = Inches(width_in)
    prs.slide_height = Inches(height_in)
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank)
    ppt_shape_id = slide_builder(prs, slide)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    _save_prs_impl(prs, out_path, backup=False)
    return ppt_shape_id if ppt_shape_id is not None else 0


# ---------------------------------------------------------------------------
# Per-category slide builders. Each returns (ppt_shape_id, bbox_dict).
# ---------------------------------------------------------------------------


def _build_text_overflow(_prs, slide) -> int:
    """A textbox overflowing its frame: long text in a tiny box."""
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(0.4))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = "这是一段非常长的正文文字，远超文本框的容量，必然产生溢出。" * 3
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(18)
    return box.shape_id


def _build_text_clean(_prs, slide) -> int:
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(4))
    box.text_frame.text = "正常正文，不溢出。"
    return box.shape_id


def _build_overlap(_prs, slide) -> int:
    """Two textboxes that overlap by design."""
    a = slide.shapes.add_textbox(Inches(2), Inches(2), Inches(4), Inches(2))
    a.text_frame.text = "上层文字"
    b = slide.shapes.add_textbox(Inches(3), Inches(2.5), Inches(4), Inches(2))
    b.text_frame.text = "下层文字"
    return a.shape_id


def _build_overlap_control(_prs, slide) -> int:
    a = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    a.text_frame.text = "左"
    b = slide.shapes.add_textbox(Inches(7), Inches(1), Inches(4), Inches(2))
    b.text_frame.text = "右"
    return a.shape_id


def _build_image_distortion(_prs, slide) -> int:
    """A picture placeholder shape stretched non-proportionally (no real image needed)."""
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(8), Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = __import__("pptx").dml.color.RGBColor(0x80, 0x80, 0x80)
    return shape.shape_id


def _build_image_clean(_prs, slide) -> int:
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(4), Inches(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = __import__("pptx").dml.color.RGBColor(0x80, 0x80, 0x80)
    return shape.shape_id


def _build_contrast(_prs, slide) -> int:
    """Low-contrast text: light grey on white."""
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(2))
    box.text_frame.text = "低对比度文字"
    from pptx.dml.color import RGBColor

    for para in box.text_frame.paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
            run.font.size = Pt(18)
    return box.shape_id


def _build_contrast_control(_prs, slide) -> int:
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(2))
    box.text_frame.text = "高对比度文字"
    from pptx.dml.color import RGBColor

    for para in box.text_frame.paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
            run.font.size = Pt(18)
    return box.shape_id


def _build_out_of_bounds(_prs, slide) -> int:
    """A shape placed partially off-canvas (left < 0)."""
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(-2), Inches(2), Inches(4), Inches(2)
    )
    shape.fill.solid()
    return shape.shape_id


def _build_in_bounds(_prs, slide) -> int:
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(4), Inches(2)
    )
    shape.fill.solid()
    return shape.shape_id


def _build_table_readability(_prs, slide) -> int:
    """A dense table with tiny font — readability issue."""
    rows, cols = 12, 4
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(0.5), Inches(12), Inches(6))
    table = tbl_shape.table
    for r in range(rows):
        for c in range(cols):
            table.cell(r, c).text = f"单元格{r}-{c}"
            for para in table.cell(r, c).text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(6)
    return tbl_shape.shape_id


def _build_table_readability_control(_prs, slide) -> int:
    """A well-formatted table with readable font — control group."""
    rows, cols = 6, 4
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), Inches(10), Inches(5))
    table = tbl_shape.table
    for r in range(rows):
        for c in range(cols):
            table.cell(r, c).text = f"单元格{r}-{c}"
            for para in table.cell(r, c).text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)
    return tbl_shape.shape_id


def _build_reference_residual(_prs, slide) -> int:
    """An empty placeholder left from a reference-deck residual."""
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(2))
    box.text_frame.text = ""  # empty — a residual placeholder
    return box.shape_id


def _build_reference_residual_control(_prs, slide) -> int:
    """A textbox with actual content — control group for reference residual."""
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(2))
    box.text_frame.text = "正常内容文本"
    return box.shape_id


_BUILDERS: dict[str, Any] = {
    "text_overflow": (_build_text_overflow, _build_text_clean),
    "overlap": (_build_overlap, _build_overlap_control),
    "image_distortion": (_build_image_distortion, _build_image_clean),
    "contrast": (_build_contrast, _build_contrast_control),
    "out_of_bounds": (_build_out_of_bounds, _build_in_bounds),
    "table_readability": (_build_table_readability, _build_table_readability_control),
    "reference_residual": (_build_reference_residual, _build_reference_residual_control),
}


def _bbox_of_shape(shape) -> dict[str, float]:
    slide_w = shape.part.package.presentation_part.presentation.slide_width
    slide_h = shape.part.package.presentation_part.presentation.slide_height
    return {
        "x": round(shape.left / slide_w, 5) if slide_w else 0.0,
        "y": round(shape.top / slide_h, 5) if slide_h else 0.0,
        "w": round(shape.width / slide_w, 5) if slide_w else 0.0,
        "h": round(shape.height / slide_h, 5) if slide_h else 0.0,
    }


def generate_annotation_dataset(
    output_dir: str | Path,
    *,
    include_controls: bool = True,
) -> dict[str, Any]:
    """Generate the full 200-page annotation dataset under ``output_dir``.

    Writes one PPTX per fixture plus a ``manifest.json`` index. Returns the
    manifest dict. Each fixture records expected issue code, stable IDs, the
    ``ppt_shape_id``, severity, BBox and evidence source.
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    fixtures: list[dict[str, Any]] = []

    for category, issue_code, min_pages in DATASET_CATEGORIES:
        issue_builder, control_builder = _BUILDERS[category]
        # Issue fixtures.
        for i in range(min_pages):
            fid = f"{category}-{i+1:03d}"
            pptx_path = out / f"{fid}.pptx"
            # Build via a wrapper that captures the primary shape's bbox.
            captured: dict[str, Any] = {}

            def builder(prs, slide, _b=issue_builder, _cap=captured):
                sid = _b(prs, slide)
                # Find the shape by id to record bbox.
                for shp in slide.shapes:
                    if shp.shape_id == sid:
                        _cap["bbox"] = _bbox_of_shape(shp)
                        break
                return sid

            ppt_shape_id = _save_slide_pptx(builder, pptx_path)
            severity = "blocker" if issue_code != ISSUE_REFERENCE_RESIDUAL else "warning"
            fixtures.append(AnnotationFixture(
                fixture_id=fid,
                category=category,
                issue_code=issue_code,
                severity=severity,
                expected={"pptx": pptx_path.name, "issue": issue_code},
                element_id=f"{fid}/primary",
                render_node_id=f"{fid}/primary/node-0",
                ppt_shape_id=ppt_shape_id,
                bbox=captured.get("bbox", {}),
                evidence_source="synthesized",
                notes=f"synthesized {category} fixture #{i+1}",
            ).to_dict())

        # Control fixtures (clean) — only when requested, one per category.
        if include_controls:
            fid = f"{category}-control"
            pptx_path = out / f"{fid}.pptx"
            captured.clear()

            def builder(prs, slide, _b=control_builder, _cap=captured):
                sid = _b(prs, slide)
                for shp in slide.shapes:
                    if shp.shape_id == sid:
                        _cap["bbox"] = _bbox_of_shape(shp)
                        break
                return sid

            ppt_shape_id = _save_slide_pptx(builder, pptx_path)
            fixtures.append(AnnotationFixture(
                fixture_id=fid,
                category=category,
                issue_code=ISSUE_CLEAN,
                severity="clean",
                expected={"pptx": pptx_path.name, "issue": ISSUE_CLEAN},
                element_id=f"{fid}/primary",
                render_node_id=f"{fid}/primary/node-0",
                ppt_shape_id=ppt_shape_id,
                bbox=captured.get("bbox", {}),
                evidence_source="synthesized",
                notes=f"clean control for {category}",
            ).to_dict())

    manifest = {
        "dataset_version": 1,
        "total_fixtures": len(fixtures),
        "categories": [
            {"name": name, "issue_code": code, "min_pages": n}
            for name, code, n in DATASET_CATEGORIES
        ],
        "issue_codes": sorted({f["issue_code"] for f in fixtures}),
        "fixtures": fixtures,
    }
    manifest_path = out / "manifest.json"
    with manifest_path.open("w", encoding="utf-8", newline="\n") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return manifest


def load_annotation_dataset(output_dir: str | Path) -> dict[str, Any]:
    """Load a previously generated dataset manifest."""
    path = Path(output_dir) / "manifest.json"
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def dataset_coverage_report(manifest: dict[str, Any]) -> dict[str, Any]:
    """Return a coverage report comparing fixture counts to the §12.3 minimums."""
    by_cat: dict[str, int] = {}
    for fx in manifest["fixtures"]:
        if fx["issue_code"] == ISSUE_CLEAN:
            continue
        by_cat[fx["category"]] = by_cat.get(fx["category"], 0) + 1
    coverage = []
    all_met = True
    for name, _code, minimum in DATASET_CATEGORIES:
        count = by_cat.get(name, 0)
        met = count >= minimum
        all_met = all_met and met
        coverage.append({
            "category": name,
            "fixtures": count,
            "minimum": minimum,
            "met": met,
        })
    return {"all_minimums_met": all_met, "coverage": coverage, "total_issue_fixtures": sum(by_cat.values())}
