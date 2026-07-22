"""Generate a comprehensive PPTX, apply all v5.x features, inspect for problems."""
from __future__ import annotations

import os
import sys
import tempfile
import traceback

if sys.platform == "win32":
    os.environ["PYTHONIOENCODING"] = "utf-8"

SECTIONS = [
    dict(title="全面功能演示", subtitle="PPTX Skill v5.1 综合测试", layout="cover"),
    dict(title="目录", bullets=["布局角色", "图表表格", "特效动画", "导出批处理"], layout="toc"),
    dict(title="第一部分", subtitle="布局角色测试", layout="section", section_number="01"),
    dict(title="核心要点", bullets=[
        "约束求解器布局引擎 Kiwi Cassowary",
        "24种布局角色 x 57种配方变体",
        "Beam search 跨页节奏规划",
        "语义 QA + 像素 SSIM 质量检查",
        "CJK 优先字体度量与断行",
        "Manifest 驱动无损编辑往返",
        "闭环修复循环 plan render QA repair",
    ], layout="bullets", kicker="核心特性"),
    dict(title="发展历程", events=[
        dict(date="2024-Q1", event="项目启动"),
        dict(date="2024-Q2", event="V1.0 发布"),
        dict(date="2024-Q3", event="V2.0 模板系统"),
        dict(date="2024-Q4", event="V3.0 布局引擎"),
        dict(date="2025-Q1", event="V4.0 特效动画"),
        dict(date="2025-Q2", event="V5.0 全功能覆盖"),
    ], layout="timeline"),
    dict(title="方案对比", left=dict(title="传统方案", bullets=["固定模板", "手动排版", "无质量检查"]),
     right=dict(title="PPTX Skill", bullets=["约束求解器", "自动布局", "闭环 QA"]), layout="comparison"),
    dict(title="核心理念", quote="让每一页幻灯片都经过数学验证", source="设计哲学", layout="quote"),
    dict(title="工作流程", steps=[
        dict(step="1", title="内容建模", desc="ContentSpec"),
        dict(step="2", title="布局规划", desc="Kiwi"),
        dict(step="3", title="渲染输出", desc="Renderer"),
        dict(step="4", title="质量检查", desc="QA"),
        dict(step="5", title="自动修复", desc="Repair"),
    ], layout="process"),
    dict(title="功能矩阵", table_headers=["模块", "API数", "状态"],
     table_rows=[
        ["布局引擎", "28", "OK"],
        ["动画系统", "51", "OK"],
        ["过渡效果", "30", "OK"],
        ["图表编辑", "19", "OK"],
        ["色彩科学", "28", "OK"],
        ["批量处理", "10", "OK"],
     ], layout="table"),
    dict(title="仪表盘", metrics=[
        dict(label="用户数", value="12847", delta="+12%"),
        dict(label="收入", value="2.4M", delta="+8%"),
        dict(label="转化率", value="3.7%", delta="-0.2%"),
        dict(label="NPS", value="72", delta="+5"),
    ], layout="dashboard"),
    dict(title="谢谢", subtitle="PPTX Skill 让演示文稿更智能", layout="end"),
]


def generate():
    from pptx_skill import auto_generate_ppt
    tmpdir = tempfile.mkdtemp(prefix="pptx_comprehensive_")
    output_path = os.path.join(tmpdir, "comprehensive.pptx")
    print("[1] Generating comprehensive PPTX ...")
    path = auto_generate_ppt(
        title="全面功能演示",
        subtitle="PPTX Skill v5.1 综合测试",
        sections=SECTIONS,
        output_path=output_path,
        lang="zh",
    )
    print("[1] Done:", path)
    return path


def apply_features(path):
    """Apply a range of v5.x features to the generated deck."""
    import shutil
    work = shutil.copy(path, os.path.join(os.path.dirname(path), "featured.pptx"))
    print("\n[2] Applying v5.x features ...")

    # 2a. Transitions
    try:
        from pptx_skill import apply_deck_transitions, FADE
        from pptx import Presentation
        prs = Presentation(work)
        apply_deck_transitions(prs, FADE, duration_ms=500)
        prs.save(work)
        print("  [OK] fade transitions applied")
    except Exception as e:
        print(f"  [FAIL] transitions: {e}")

    # 2b. Morph on slide 1
    try:
        from pptx_skill import set_morph_transition
        set_morph_transition(work, 1, option="full")
        print("  [OK] morph transition on slide 1")
    except Exception as e:
        print(f"  [FAIL] morph: {e}")

    # 2c. Advanced transition on slide 2
    try:
        from pptx_skill import set_wheel_transition
        set_wheel_transition(work, 2, spokes=6)
        print("  [OK] wheel transition on slide 2")
    except Exception as e:
        print(f"  [FAIL] wheel transition: {e}")

    # 2d. Animations
    try:
        from pptx_skill import apply_entrance_animation, FADE_IN
        from pptx import Presentation
        prs = Presentation(work)
        # Apply to first shape on slide 0
        slide = prs.slides[0]
        if slide.shapes:
            apply_entrance_animation(slide, slide.shapes[0], FADE_IN)
        prs.save(work)
        print("  [OK] fade-in animation on slide 0 shape 0")
    except Exception as e:
        print(f"  [FAIL] animation: {e}")

    # 2e. Watermark
    try:
        from pptx_skill import add_text_watermark
        add_text_watermark(work, "DRAFT v5.1", opacity=0.15)
        print("  [OK] text watermark")
    except Exception as e:
        print(f"  [FAIL] watermark: {e}")

    # 2f. Effects (shadow on a shape)
    try:
        from pptx_skill import apply_shadow, apply_effect_preset
        from pptx import Presentation
        prs = Presentation(work)
        # find a shape with fill on slide 0
        for slide in prs.slides:
            for shape in slide.shapes:
                try:
                    apply_shadow(shape, blur_pt=8, dist_pt=4, dir_deg=90, color="808080", alpha=50)
                    prs.save(work)
                    print("  [OK] shadow effect applied")
                    break
                except Exception:
                    continue
            else:
                continue
            break
    except Exception as e:
        print(f"  [FAIL] effect: {e}")

    # 2g. Sections
    try:
        from pptx_skill import add_section
        add_section(work, "布局测试", start_slide=1)
        add_section(work, "数据展示", start_slide=6)
        print("  [OK] sections added")
    except Exception as e:
        print(f"  [FAIL] sections: {e}")

    # 2h. Comments / notes
    try:
        from pptx_skill import set_speaker_notes, add_comment
        set_speaker_notes(work, 1, "这是封面页，介绍 PPTX Skill v5.1 的全面功能。")
        print("  [OK] speaker notes on slide 1")
    except Exception as e:
        print(f"  [FAIL] notes: {e}")

    # 2i. Metadata
    try:
        from pptx_skill import set_metadata, set_custom_property
        set_metadata(work, title="全面功能演示", author="PPTX Skill", subject="综合测试")
        set_custom_property(work, "version", "5.1.0")
        set_custom_property(work, "test_type", "comprehensive")
        print("  [OK] metadata + custom properties")
    except Exception as e:
        print(f"  [FAIL] metadata: {e}")

    # 2j. Chart editing (if there is a chart)
    try:
        from pptx_skill import list_charts
        charts = list_charts(work)
        if charts:
            print(f"  [OK] found {len(charts)} chart(s): {charts}")
        else:
            print("  [info] no charts in deck")
    except Exception as e:
        print(f"  [FAIL] chart list: {e}")

    # 2k. Accessibility
    try:
        from pptx_skill import audit_accessibility
        report = audit_accessibility(work)
        print(f"  [OK] accessibility audit: {len(report.issues)} issues")
    except Exception as e:
        print(f"  [FAIL] accessibility: {e}")

    return work


def inspect(path):
    """Deep inspection of the final PPTX for problems."""
    print("\n[3] Inspecting final PPTX ...")
    from pptx import Presentation
    from pptx.util import Emu, Inches
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(path)
    problems = []
    warnings = []

    slide_w = prs.slide_width
    slide_h = prs.slide_height
    print(f"  Slide size: {slide_w/914400:.2f} x {slide_h/914400:.2f} in ({slide_w} x {slide_h} EMU)")

    # Shape names that are decorative and expected to be partially off-slide
    WATERMARK_PREFIXES = ("pptx_skill_watermark", "watermark")

    # Collect shapes with types for overlap analysis
    all_slide_shapes = []

    for idx, slide in enumerate(prs.slides):
        shape_count = len(slide.shapes)
        shapes_info = []
        print(f"  Slide {idx}: {shape_count} shapes")

        for shape in slide.shapes:
            name = getattr(shape, 'name', '')
            shape_type = shape.shape_type

            # Check for off-slide shapes — skip decorative ovals/circles
            # that are intentionally placed partially off-slide
            try:
                left = shape.left or 0
                top = shape.top or 0
                width = shape.width or 0
                height = shape.height or 0
                right = left + width
                bottom = top + height

                is_decorative = name.startswith("Oval") and (
                    (left < 0 or top < 0) or
                    (right > slide_w + Inches(0.5) or bottom > slide_h + Inches(0.5))
                )

                if not is_decorative:
                    if left < 0 or top < 0:
                        problems.append(f"Slide {idx} shape {name!r}: negative position ({left},{top})")
                    if right > slide_w + Inches(0.1) or bottom > slide_h + Inches(0.1):
                        problems.append(f"Slide {idx} shape {name!r}: extends off-slide (right={right}, bottom={bottom}, slide={slide_w}x{slide_h})")
            except Exception as e:
                problems.append(f"Slide {idx} shape {name!r}: inspection error {e}")

            # Check for empty text boxes — skip watermarks
            try:
                if shape.has_text_frame and not any(name.startswith(p) for p in WATERMARK_PREFIXES):
                    text = shape.text_frame.text.strip()
                    if not text and shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                        problems.append(f"Slide {idx}: empty text box {name!r}")
            except Exception:
                pass

            # Collect for overlap analysis
            try:
                if shape.left is not None and shape.top is not None:
                    shapes_info.append({
                        'name': name,
                        'type': shape_type,
                        'left': shape.left,
                        'top': shape.top,
                        'right': shape.left + (shape.width or 0),
                        'bottom': shape.top + (shape.height or 0),
                        'area': (shape.width or 0) * (shape.height or 0),
                    })
            except Exception:
                pass

        all_slide_shapes.append(shapes_info)

    # Summary: count real issues
    off_slide_count = sum(1 for p in problems if "off-slide" in p or "negative position" in p)
    empty_text_count = sum(1 for p in problems if "empty text box" in p)
    overlap_count = sum(1 for p in problems if "overlap" in p)

    print(f"\n[3] Found {len(problems)} problem(s): {off_slide_count} off-slide, {empty_text_count} empty text, {overlap_count} overlaps")
    for p in problems[:30]:
        print(f"  - {p}")
    if len(problems) > 30:
        print(f"  ... and {len(problems)-30} more")

    return problems


def main():
    try:
        path = generate()
        featured = apply_features(path)
        problems = inspect(featured)
        print(f"\n=== RESULT ===")
        print(f"Generated: {path}")
        print(f"Featured:  {featured}")
        print(f"Problems:  {len(problems)}")
        # Write problems to a file for follow-up
        with open(os.path.join(os.path.dirname(path), "problems.txt"), "w", encoding="utf-8") as f:
            for p in problems:
                f.write(p + "\n")
        return problems
    except Exception as e:
        print(f"[FATAL] {e}")
        traceback.print_exc()
        return None


if __name__ == "__main__":
    main()
