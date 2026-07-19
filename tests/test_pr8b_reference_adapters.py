"""Tests for PR8b: native / clone reference-deck adapters.

Per blueprint §9.3 and PR8b acceptance:
- Two reference modes (native + clone) each have an E2E flow.
- clone mode must not drift unmapped shape geometry.
- Canvas dimensions come from the reference file, not a 16:9 constant.
- visual-rebuild is exposed as a clone alias pending PR8c.
"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from pptx_skill.content_model import (
    ContentSpec,
    ElementSpec,
    SlideSpec,
)
from pptx_skill.reference_adapter import (
    REFERENCE_MODES,
    ReferenceAdapterResult,
    clone_mode_adapter,
    generate_from_reference_adapter,
    native_mode_adapter,
    reference_canvas,
)


def _build_reference_pptx(tmpdir: Path) -> Path:
    """Build a small reference PPTX with cover + bullets + dashboard + end slides."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    title_layout = prs.slide_layouts[0]

    # Cover slide
    s = prs.slides.add_slide(title_layout)
    if s.shapes.title:
        s.shapes.title.text = "参考稿标题"
    box = s.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
    box.text_frame.text = "副标题"

    # Bullets slide
    s = prs.slides.add_slide(blank)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
    box.text_frame.text = "要点标题"
    bullets = s.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(4))
    tf = bullets.text_frame
    tf.text = "第一点"
    tf.add_paragraph().text = "第二点"
    tf.add_paragraph().text = "第三点"

    # Dashboard slide (with a chart-like metric box)
    s = prs.slides.add_slide(blank)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
    box.text_frame.text = "数据仪表盘"
    metric = s.shapes.add_textbox(Inches(1), Inches(2.5), Inches(4), Inches(2))
    metric.text_frame.text = "关键指标"

    # End slide
    s = prs.slides.add_slide(title_layout)
    if s.shapes.title:
        s.shapes.title.text = "谢谢"

    path = tmpdir / "reference.pptx"
    prs.save(str(path))
    return path


def _build_content_spec() -> ContentSpec:
    """A small content spec matching the reference roles."""
    slides = [
        SlideSpec(
            id="cover", role="cover", communication_goal="cover",
            elements=[ElementSpec(id="cover/title", kind="text", role="title",
                                  content={"text": "我们的演示"}, style_ref="component.title")],
        ),
        SlideSpec(
            id="bullets-1", role="bullets", communication_goal="要点",
            elements=[
                ElementSpec(id="bullets-1/title", kind="text", role="title",
                            content={"text": "核心观点"}, style_ref="component.title"),
                ElementSpec(id="bullets-1/body", kind="text", role="body",
                            content={"text": "观点一\n观点二\n观点三"}, style_ref="component.body"),
            ],
        ),
        SlideSpec(
            id="dash-1", role="dashboard", communication_goal="数据",
            elements=[
                ElementSpec(id="dash-1/title", kind="text", role="title",
                            content={"text": "关键指标"}, style_ref="component.title"),
            ],
        ),
        SlideSpec(
            id="end-1", role="end", communication_goal="end",
            elements=[ElementSpec(id="end-1/title", kind="text", role="title",
                                  content={"text": "谢谢"}, style_ref="component.title")],
        ),
    ]
    return ContentSpec(id="deck-1", title="我们的演示", subtitle="副标题", slides=slides)


class ReferenceCanvasTests(unittest.TestCase):
    def test_reference_canvas_returns_native_dimensions(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            info = reference_canvas(ref)
            self.assertEqual(info.width_emu, int(Inches(13.333)))
            self.assertEqual(info.height_emu, int(Inches(7.5)))
            self.assertAlmostEqual(info.aspect_ratio, 13.333 / 7.5, places=3)

    def test_to_canvas_spec_uses_native_dimensions(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            info = reference_canvas(ref)
            spec = info.to_canvas_spec()
            self.assertEqual(spec.width_pt, info.width_pt)
            self.assertEqual(spec.height_pt, info.height_pt)


class NativeModeTests(unittest.TestCase):
    def test_e2e_build(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            content = _build_content_spec()
            out = Path(tmp) / "out_native.pptx"
            result = native_mode_adapter(ref, content, out)
            self.assertIsInstance(result, ReferenceAdapterResult)
            self.assertEqual(result.mode, "native")
            self.assertTrue(out.exists())
            prs = Presentation(str(out))
            # Built from masters/layouts; has at least one slide per section.
            self.assertGreaterEqual(len(prs.slides), 1)
            # Canvas preserves reference dimensions (not 16:9-default).
            self.assertEqual(prs.slide_width, int(Inches(13.333)))


class CloneModeTests(unittest.TestCase):
    def test_e2e_clone_no_drift(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            content = _build_content_spec()
            out = Path(tmp) / "out_clone.pptx"
            result = clone_mode_adapter(ref, content, out, clear_unmapped_content=False)
            self.assertEqual(result.mode, "clone")
            self.assertTrue(out.exists())
            prs = Presentation(str(out))
            self.assertGreaterEqual(len(prs.slides), 1)
            # PR8b acceptance: unmapped shapes must not drift in clone mode.
            self.assertEqual(
                result.unmapped_shape_drift, [],
                f"clone mode drifted {len(result.unmapped_shape_drift)} unmapped shapes",
            )

    def test_e2e_clone_with_clear_unmapped(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            content = _build_content_spec()
            out = Path(tmp) / "out_clone_clear.pptx"
            result = clone_mode_adapter(ref, content, out, clear_unmapped_content=True)
            self.assertEqual(result.mode, "clone")
            self.assertEqual(result.unmapped_shape_drift, [])


class DispatchTests(unittest.TestCase):
    def test_auto_recommends_native_or_clone(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            content = _build_content_spec()
            out = Path(tmp) / "out_auto.pptx"
            result = generate_from_reference_adapter(ref, out, content, mode="auto")
            self.assertIn(result.mode, {"native", "clone"})

    def test_visual_rebuild_routes_to_real_rebuild(self):
        """After PR8c, visual-rebuild is a real adaptive path, not a clone alias."""
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            content = _build_content_spec()
            out = Path(tmp) / "out_vr.pptx"
            result = generate_from_reference_adapter(ref, out, content, mode="visual-rebuild")
            self.assertEqual(result.mode, "visual-rebuild")
            self.assertTrue(result.diagnostics.get("real_rebuild"))
            # Real rebuild derives mean_similarity; clone alias had none.
            self.assertIn("mean_similarity", result.diagnostics)
            self.assertNotIn("alias_for_clone_pending_pr8c", str(result.diagnostics))

    def test_unknown_mode_raises(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = _build_reference_pptx(Path(tmp))
            content = _build_content_spec()
            out = Path(tmp) / "out_bad.pptx"
            with self.assertRaises(ValueError):
                generate_from_reference_adapter(ref, out, content, mode="bogus")


class ReferenceModesConstantsTests(unittest.TestCase):
    def test_supported_modes(self):
        self.assertIn("native", REFERENCE_MODES)
        self.assertIn("clone", REFERENCE_MODES)
        self.assertIn("visual-rebuild", REFERENCE_MODES)


if __name__ == "__main__":
    unittest.main()
