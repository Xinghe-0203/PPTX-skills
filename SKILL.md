---
name: pptx
version: "6.0.0"
description: |
  创建、编辑、重构和验收可编辑 PowerPoint 演示文稿。用于从主题或资料生成 PPT、自动排版、选择或生成模板、按照用户提供的 .pptx 母版或样例页高保真复刻、替换指定文字/图片/图表、制作数据可视化、渲染截图并迭代修正。支持 24 种内容版式（含 5 种分析框架）、62 种版式配方、20 套模板档案、3 套非卡片式版式家族、20 套传统主题、字体配对系统（含 CJK）、10 种图表类型、富文本渲染、表格/形状增强、页眉页脚页码、备注页与媒体节点、模板包下载与档案提取、30 种转场动画（18 标准 + 12 高级，另有 Morph）、51 种入场/退出/强调动画、多文件合并与幻灯片提取、文字/图片水印、节管理、母版/版式查询与克隆、图片压缩与优化、演讲者备注与审阅批注、结构化差异比较、多格式导出（PDF/图片/HTML/文本）、SmartArt 检测与编辑、VBA 宏注入与提取、20 种内置表格样式与单元格操作、HTML 导入。
---

# PPTX 智能制作

> 一句话：始终交付可编辑的 `.pptx`，并执行渲染验收。不要只检查代码或 XML。

## 适用场景

- 用户给主题/资料，需要生成完整 PPT。
- 用户提供了企业母版或规范模板，需要按母版占位符生成。
- 用户提供了样例 PPT，需要高保真复刻并替换内容。
- 用户提供了扁平截图或复杂结构，需要重建可编辑近似版本。
- 用户需要修改已有 PPT 的文字、图片、颜色、页面顺序或版式。
- 用户需要把数据表格/指标/时间线/流程图做成可编辑图表或表格。
- 用户需要合并多个 PPT、添加水印、设置转场/动画、导出 PDF/图片。
- 用户需要从 HTML 幻灯片（Marp/Slidev/reveal.js）导入为可编辑 PPTX。

## 不适用场景

- 需要保留复杂 PowerPoint 动画、宏、OLE、嵌入工作簿或部分 SmartArt 的完整行为。
- 要求从扁平截图恢复原始矢量、不可见数据或原动画。
- 要求 LibreOffice 与 PowerPoint 逐像素一致（不同引擎使用不同基线）。
- 要求模型直接生成并执行未验证的 OOXML/Python/JavaScript。

## 能力地图

| 能力 | 状态 | 主要入口 |
|---|---|---|
| 从头生成演示文稿 | 已落地 | `pptx_skill.auto_generate_ppt()` / `scripts/pptx_helper.py` |
| 模板档案管理 | 已落地 | `scripts/template_engine.py list/generate/preview` |
| 模板包下载与档案提取 | 已落地 | `pptx_skill.template_downloader.download_template_pack()` / `import_template()` |
| 母版占位符生成（native） | 已落地 | `pptx_skill.reference_adapter.native_mode_adapter()` |
| 样例页逐形状克隆（clone） | 已落地 | `pptx_skill.reference_adapter.clone_mode_adapter()` |
| 扁平稿视觉重建（visual-rebuild） | 已落地 | `pptx_skill.visual_rebuild.visual_rebuild_adapter()` |
| 约束式自适应排版 | 已落地 | `pptx_skill.layout_engine.solve_recipe()` |
| 整本拆页与 beam search 节奏 | 已落地 | `pptx_skill.deck_planner.plan_deck()` |
| 文字度量与字号适配 | 已落地 | `pptx_skill.text_metrics.measure_runs()` |
| 语义 QA（四级：表格/图表/页面/整本） | 已落地 | `pptx_skill.semantic_qa.SemanticQAEngine` |
| 渲染 QA 与视觉回归 | 已落地 | `pptx_skill.render_qa.compare_slide_to_baseline()` |
| 白名单修复引擎 | 已落地 | `pptx_skill.repair_engine.propose_repairs()` |
| 生成 → QA → 修复闭环 | 已落地 | `pptx_skill.generation_pipeline.run_generation_pipeline()` |
| Manifest V3 读写 | 已落地 | `pptx_skill.manifest.load_manifest()` |
| 定点文字/图片/颜色编辑 | 已落地 | `pptx_skill.edit_text()` / `pptx_skill.swap_image()` / `pptx_skill.recolor()` |
| 页面增删移复制换版式 | 已落地 | `pptx_skill.insert_slide()` / `pptx_skill.delete_slide()` / `pptx_skill.move_slide()` |
| 10 种图表类型（多系列） | 已落地 | `pptx_skill.pptx_renderer` chart 节点 |
| 富文本渲染（多段落/多 run/项目符号/超链接） | 已落地 | `pptx_skill.pptx_renderer` rich-text 节点 |
| 表格增强样式（边框/条纹/填充/列宽） | 已落地 | `pptx_skill.pptx_renderer` table 节点 |
| 形状增强样式（线条/阴影/渐变/旋转/内嵌文字） | 已落地 | `pptx_skill.pptx_renderer` shape 节点 |
| 页眉页脚与页码 | 已落地 | `pptx_skill.pptx_renderer` deck_options |
| 备注页与媒体（video/audio） | 已落地 | `pptx_skill.pptx_renderer` notes/media 节点 |
| 字体配对系统（含 CJK） | 已落地 | `scripts/pptx_helper.py` 的 `FONTS` / `CJK_FONTS` |
| 200 页 QA 标注数据集 | 已落地 | `pptx_skill.qa_dataset.generate_annotation_dataset()` |
| Golden renders 基线 | 已落地 | `pptx_skill.golden_renders.render_golden_set()` |
| 30 种转场动画（18 标准 + 12 高级，另有 Morph） | 已落地 | `pptx_skill.transitions.apply_slide_transition()` / `pptx_skill.transitions_ext.set_advanced_transition()` / `pptx_skill.morph.set_morph_transition()` |
| 51 种幻灯片动画（20 入场/15 退出/15 强调/1 运动路径） | 已落地 | `pptx_skill.animations.apply_animation()` / `apply_entrance_animation()` / `apply_exit_animation()` / `apply_emphasis_animation()` / `apply_motion_path()` |
| 多文件合并与幻灯片提取 | 已落地 | `pptx_skill.merge.merge_presentations()` / `append_slides()` / `insert_slides()` / `extract_slides()` |
| 文字/图片水印（透明度/平铺/层级/移除） | 已落地 | `pptx_skill.watermark.add_text_watermark()` / `add_image_watermark()` / `remove_watermark()` |
| 节管理（增删改移折叠） | 已落地 | `pptx_skill.sections.add_section()` / `remove_section()` / `rename_section()` / `move_section()` / `collapse_section()` |
| 母版/版式查询、克隆、占位符、背景 | 已落地 | `pptx_skill.slide_master.list_slide_masters()` / `clone_layout()` / `add_placeholder_to_layout()` / `set_master_background()` |
| 图片压缩/格式转换/未用媒体清理 | 已落地 | `pptx_skill.image_optimize.compress_images()` / `convert_image_format()` / `remove_unused_media()` |
| 演讲者备注与审阅批注（线程回复/解决） | 已落地 | `pptx_skill.comments.set_speaker_notes()` / `add_comment()` / `reply_to_comment()` / `resolve_comment()` |
| 结构化差异比较（幻灯片/形状/文字/图片/版式） | 已落地 | `pptx_skill.diff.diff_presentations()` / `diff_slides()` / `diff_text()` |
| 多格式导出（PDF/图片/HTML/文本/缩略图） | 已落地 | `pptx_skill.export.export_to_pdf()` / `export_to_images()` / `export_to_html()` / `export_to_text()` / `export_thumbnails()` |
| SmartArt 检测/文本提取编辑/跨页保留 | 已落地 | `pptx_skill.smartart.detect_smartart()` / `extract_smartart_text()` / `populate_smartart_text()` / `preserve_smartart()` |
| VBA 宏检测/注入/提取/形状绑定 | 已落地 | `pptx_skill.vba.has_vba_project()` / `inject_vba_project()` / `extract_vba_project()` / `attach_macro_to_shape()` |
| 20 种内置表格样式 + 单元格合并/边框/填充/行列增删 | 已落地 | `pptx_skill.table_styles.apply_table_style()` / `merge_cells()` / `set_cell_border()` / `add_table_row()` |
| HTML 幻灯片导入（Marp/Slidev/reveal.js） | 已落地 | `pptx_skill.html_import.import_from_html()` / `html_to_content_spec()` / `detect_html_slides()` |
| 3D/斜面/发光/反射/柔边/内阴影/透视阴影 + 16 特效预设 | 已落地 | `pptx_skill.effects.apply_shadow()` / `apply_glow()` / `apply_reflection()` / `apply_effect_preset()` |
| 线条/肘形/曲线连接符/自由路径/箭头/虚线 | 已落地 | `pptx_skill.connectors.add_line()` / `add_elbow_connector()` / `add_freeform()` / `set_arrow_style()` |
| 写保护/密码加密/标记最终版 | 已落地 | `pptx_skill.protection.apply_write_protection()` / `encrypt_pptx()` / `mark_as_final()` |
| WCAG 2.1 无障碍审计/替代文本/对比度/自动修复 | 已落地 | `pptx_skill.accessibility.audit_accessibility()` / `set_alt_text()` / `check_contrast()` / `fix_accessibility()` |
| 核心文档属性/自定义属性/嵌入字体 | 已落地 | `pptx_skill.metadata.set_metadata()` / `set_custom_property()` / `embed_font()` |
| OMML 公式构建器 + LaTeX→OMML（分数/根号/n 元/矩阵） | 已落地 | `pptx_skill.equations.add_equation()` / `add_display_equation()` / `list_equations()` |
| 视频/音频嵌入与播放设置（循环/全屏/裁剪） | 已落地 | `pptx_skill.multimedia.add_video()` / `add_audio()` |
| 形状分组/取消分组/组内 z 序 | 已落地 | `pptx_skill.groups.group_shapes()` / `ungroup_shapes()` / `list_groups()` |
| Slide Zoom/Section Zoom/Summary Zoom | 已落地 | `pptx_skill.zoom.add_slide_zoom()` / `add_section_zoom()` / `add_summary_zoom()` |
| Morph 变形转场（p15，按对象/词/字） | 已落地 | `pptx_skill.morph.set_morph_transition()` / `set_morph_options()` |
| 12 种高级转场（轮/波纹/蜂窝/漩涡等，p14） | 已落地 | `pptx_skill.transitions_ext.set_advanced_transition()` / `set_wheel_transition()` / `set_flip_transition()` |
| 批量/并行处理（转换/水印/换色/检查/统计） | 已落地 | `pptx_skill.batch.batch_convert()` / `batch_watermark()` / `batch_recolor()` / `batch_inspect()` |
| SVG→DrawingML 转换（path/rect/circle/line/polygon） | 已落地 | `pptx_skill.svg_import.import_svg()` / `import_svg_as_image()` / `svg_to_drawingml()` |
| OCR 集成（Tesseract/EasyOCR/PaddleOCR）+ 自动替代文本 | 已落地 | `pptx_skill.ocr.ocr_slide()` / `ocr_presentation()` / `auto_caption_slide()` |
| 图表数据/样式/类型编辑 + 系列增删/轴控制/数据标签 | 已落地 | `pptx_skill.chart_edit.edit_chart_data()` / `add_chart_series()` / `set_chart_type()` / `set_data_labels()` |
| 色彩科学（RGB/HSL/HSV/CMYK）+ 调色板提取 + 和谐生成器 | 已落地 | `pptx_skill.color.extract_palette()` / `complementary_colors()` / `analogous_colors()` / `recolor_presentation()` |
| 形状对齐/分布/吸附网格/布局模板/z 序/居中 | 已落地 | `pptx_skill.layout_sync.align_shapes()` / `distribute_shapes()` / `snap_to_grid()` / `apply_layout_template()` |
| Markdown 导出（文本/表格/备注/图表数据/元数据） | 已落地 | `pptx_skill.markdown_export.export_to_markdown()` / `slide_to_markdown()` |
| MP4/GIF 视频导出（ffmpeg，交叉淡入/演讲计时） | 已落地 | `pptx_skill.video.export_to_video()` / `export_to_gif()` |
| DrawingML→SVG 导出（幻灯片→SVG 矢量输出） | 已落地 | `pptx_skill.svg_export.export_to_svg()` / `export_slide_to_svg()` |
| **统一 Deck 高层 API（v6.0，链式调用）** | 已落地 | `from pptx_skill import Deck` → `Deck.open(path).add_watermark("DRAFT").add_notes(1,"...").add_transition(1,"fade").save()` |
| **统一 CLI（v6.0）** | 已落地 | `pptx-skill info\|inspect\|generate\|from-markdown\|render\|edit\|watermark\|export\|pages\|merge\|validate\|template\|capability` |
| **Markdown 导入（v6.0）** | 已落地 | `pptx_skill.markdown_import.markdown_to_sections()` / `import_markdown()` |
| **全局查找替换（v6.0）** | 已落地 | `pptx_skill.find_replace_all(path, find, replace)` |
| **路径式转场/动画重载（v6.0）** | 已落地 | `pptx_skill.add_transition(path, slide_index, ...)` / `add_entrance_animation(path, slide_index, ...)` |

## 前置依赖

从项目根目录安装为可编辑包（推荐）：

```powershell
python -m pip install -e .
```

检查当前环境能力：

```powershell
python -m pptx_skill.capability
```

确认核心依赖可用：

```powershell
python -m pip show python-pptx Pillow
python scripts/render_slides.py --help
```

截图渲染优先使用 LibreOffice；Windows 也可回退到 PowerPoint COM。缺少渲染引擎时再安装 LibreOffice，不要无条件安装。

可选后端按需要安装（详见 `THIRD_PARTY_NOTICES.md`）：

```powershell
python -m pip install -e ".[adaptive,qa-image,render-pdf]"
```

## 按任务选择路径

1. **用户没有提供参考 PPT**
   - 直接使用已有模板档案生成；或
   - 根据行业、品牌色、受众和语气生成新模板档案，再生成演示文稿。
2. **用户提供了规范母版或企业模板**
   - 分析母版、版式和占位符；
   - 使用 `native` 模式从原生版式创建页面。
3. **用户提供了样例演示文稿，设计主要画在页面上**
   - 使用 `clone` 模式复制代表页；
   - 按形状名称精确替换文字、图片和内容区域。
4. **用户提供的是扁平图片、复杂 SmartArt/OLE 或不可复用结构**
   - 使用 `visual-rebuild` 路径；
   - 按截图重建可编辑元素并进行并排验收。
5. **用户提供了 HTML 幻灯片（Marp/Slidev/reveal.js）**
   - 使用 `html_import` 路径；
   - 提取语义内容并生成可编辑 PPTX，可选叠加光栅化底图。

处理用户提供的 PPT 时，完整阅读 [references/reference-ppt-workflow.md](references/reference-ppt-workflow.md)。
修改已有文件时，完整阅读 [references/editing-workflow.md](references/editing-workflow.md)。
调用生成 API 或管理模板档案时，阅读 [references/engine-api.md](references/engine-api.md)。
交付前阅读 [references/quality-checklist.md](references/quality-checklist.md)。

## 核心工作流

### 1. 结构化需求

提炼并确认：

- 受众、使用场景、页数、语言、画幅和演讲时长；
- 主标题、副标题和章节顺序；
- 每页唯一沟通任务；
- 数据、对比、流程、时间线、表格、图表、图片和引用；
- 模板来源、品牌色、字体、Logo 与必须保留的元素。

当用户只提供主题时，主动扩展为具体内容。避免空泛句子；对可能变化的数据进行互联网核验并记录来源。

### 2. 选择或生成模板

列出现有模板：

```powershell
python scripts/template_engine.py list
```

需要全新视觉语言时，从要求生成模板档案并输出可编辑预览：

```powershell
python scripts/template_engine.py generate `
  --name "模板名称" `
  --prompt "目标行业、风格、受众和视觉要求" `
  --brand-color "#184E77" `
  --register `
  --preview output/template-preview.pptx
```

渲染预览并检查后再正式生成。模板档案必须控制色彩、版式参数和推荐页面节奏，不能只是换一组颜色。

新模板优先从 `editorial_grid`、`technical_axis`、`poster_column` 三种几何语言中选择。除非用户明确要求，不要默认使用圆角卡片、装饰圆、渐变光斑、悬浮阴影、`VS` 徽章、巨大引号或无意义英文眉题。

### 3. 处理用户提供的 PPT

先渲染和分析：

```powershell
python scripts/render_slides.py reference.pptx `
  --output output/reference-preview --dpi 150

python scripts/reference_ppt.py analyze reference.pptx `
  --output output/reference-analysis.json
```

逐页查看参考图和分析结果，再选择模式：

- `native`：使用真实母版、版式和占位符；
- `clone`：复制样例页及其媒体/图表关系，再替换内容；
- `visual-rebuild`：提取设计语言并重建。

不要把"从参考稿提取主题色"当作"完全复刻"。当精确位置重要时，使用分析结果中的形状名称编写 shape-level plan。

### 4. 生成演示文稿

新建演示文稿时调用 `auto_generate_ppt()`（也可从 `pptx_skill` 包导入），按内容特征选择版式：

```python
from pptx_skill import auto_generate_ppt

auto_generate_ppt(
    title="汇报标题",
    subtitle="副标题",
    sections=sections,
    output_path="output/report.pptx",
    template_key="strategy-consulting",
    auto_search_images=False,
)
```

版式选择规则：

- 指标 → `dashboard` / `kpi_hero` / `dashboard.metric_hero`
- 时间节点 → `timeline`
- 两侧信息 → `comparison` / `comparison.matrix`
- 步骤 → `process` / `process.milestone`
- 表格 → `table`
- 引用 → `quote` / `testimonial`
- 多图 → `image_grid` / `logo_wall`
- 图文 → `text_image`
- 四象限 → `matrix` / `swot` / `swot.quadrant`
- 问答 → `faq`
- 普通要点 → `bullets` / `bullets.with_icon`
- 战略分析 → `porter` / `porter.diamond`
- 环境分析 → `pest` / `pest.grid`
- 商业模式 → `bmc` / `bmc.canvas`
- 转化漏斗 → `funnel` / `funnel.stacked`
- 封面宣言 → `cover.statement`

> **注意**：`kpi_hero`、`testimonial`、`logo_wall`、`matrix`、`faq`、`swot`、`porter`、`pest`、`bmc`、`funnel` 这 10 种 role 没有默认布局，使用时必须在 `SlideSpec` 或 `sections` 中显式指定 layout 名称（如 `kpi_hero.split`、`testimonial.centered`、`logo_wall.grid3`、`matrix.quadrant`、`faq.alternating`、`swot.quadrant`、`porter.diamond`、`pest.grid`、`bmc.canvas`、`funnel.stacked` 等）。

脚本入口仍可用：

```powershell
python scripts/reference_ppt.py generate reference.pptx `
  --content content.json --mode clone --output output/report.pptx
```

复杂页面改用显式计划：

```powershell
python scripts/reference_ppt.py compose reference.pptx `
  --plan plan.json --output output/report.pptx
```

### 5. 验证与迭代

运行结构检查：

```python
from pptx_skill import auto_validate_ppt
result = auto_validate_ppt("output/report.pptx")
```

渲染所有页面：

```powershell
python scripts/render_slides.py output/report.pptx `
  --output output/report-preview --dpi 150
```

逐页检查文字溢出、遮挡、裁图、对比度、层级、旧内容残留和模板一致性。至少执行两轮"生成 → 渲染 → 逐页检查 → 修正"，连续两轮没有阻断问题后再交付。

## 数据契约与 Manifest V3

生成和修改都围绕新的数据模型展开：

- `ContentSpec` 表示整本 PPT 内容，`SlideSpec` 表示一页，`ElementSpec` 表示页内元素；每个元素都有跨重生成的稳定 `element_id`。
- 本 Skill 生成的 PPTX 自动写入 Manifest V3（内嵌 XML + `.manifest.json` 侧写），记录 `content`、`layout_plans`、`render_trace` 和每次生成/QA/修复尝试。
- `from pptx_skill import auto_generate_ppt` 是兼容 facade：默认行为与旧 API 完全一致；传入 `return_result=True` 可获得结构化结果：

```python
from pptx_skill import auto_generate_ppt, GenerationResult

result = auto_generate_ppt(
    title="汇报标题",
    sections=sections,
    output_path="output/report.pptx",
    template_key="strategy-consulting",
    qa_mode="report",       # off | report | strict（strict 暂未启用阻断）
    return_result=True,
)
assert isinstance(result, GenerationResult)
print(result.qa_status)   # pass | fail | inconclusive
```

读取或迁移 manifest：

```python
from pptx_skill import load_manifest

manifest = load_manifest("output/report.pptx")
print(manifest.manifest_schema_version)  # 3
```

## 修改已有 PPT

- 对本 Skill 生成的文件，优先读取内嵌 manifest：
  - `from pptx_skill import load_project, edit_section, regenerate`（legacy v2 路径仍可用）
  - `from pptx_skill import load_manifest`（V3 路径）
  - 改 Section 后调用 `regenerate(project, output_path)` 重画。
- 对外部文件，使用 `pptx_skill.ppt_edit` 定点换字、换图或换色：
  - `edit_text` / `edit_text_by_role` / `swap_image` / `recolor` / `swap_theme`。
- 使用 `pptx_skill.ppt_pages` 增页、删页、移动、复制或重画单页：
  - `insert_slide` / `delete_slide` / `move_slide` / `duplicate_slide` / `replace_layout`。
- 合并多个 PPT 或提取幻灯片：
  - `merge_presentations()` / `append_slides()` / `insert_slides()` / `extract_slides()`。
- 添加水印、转场、动画：
  - `add_text_watermark()` / `add_image_watermark()` / `apply_slide_transition()` / `apply_animation()`。
- 管理节结构：
  - `add_section()` / `remove_section()` / `rename_section()` / `move_section()`。
- 导出为 PDF/图片/HTML/文本：
  - `export_to_pdf()` / `export_to_images()` / `export_to_html()` / `export_to_text()`。
- 所有写操作自动生成最近一次 `.bak.pptx` 备份；修改失败时调用 `restore_backup()`。
- 复杂页面先分析形状名称，再用 shape-level plan 或显式 Section 数据，不依赖模糊角色猜测。

## 错误处理

| 问题 | 典型原因 | 处理方式 |
|---|---|---|
| `ImportError: kiwisolver` | 未安装 adaptive 依赖 | `pip install -e ".[adaptive]"` 或回退 `layout_engine="legacy"` |
| `LayoutPlanningError` | 求解器不可行、内容超长、约束冲突或 recipe 不匹配 | 换 recipe、分页或放宽内容限制 |
| `AdaptiveRendererError` | 渲染节点类型不支持、画布参数异常或渲染过程出错 | 检查节点类型与参数，回退 legacy 渲染 |
| `ManifestError` | manifest 格式损坏、版本不兼容或读写失败 | 检查 manifest 版本，尝试 V2→V3 迁移或重新生成 |
| 参考稿模式推荐 `visual-rebuild` | 母版/占位符不可用，形状都在页面上 | 按视觉重建路径执行，并给出相似度预算报告 |
| QA 报 `blocker` | 溢出、重叠、越界、低对比度等 | 进入 repair 闭环，最多两轮；仍失败则明确返回 |
| 渲染引擎不可用 | 未安装 LibreOffice/PyMuPDF/PowerPoint COM | 标记 unavailable，不影响结构/语义 QA；需要时再安装 |
| 图片搜索失败 | Pixabay 网络或关键词问题 | 降级为本地图片或纯文字版式，不中断生成 |
| `ValueError` (animations) | 形状无有效 ID，无法绑定动画 | 确保形状已添加到幻灯片且有有效 shape_id |
| `FileNotFoundError` (merge/diff/export) | 源/目标 PPTX 文件路径不存在 | 检查文件路径，确认文件存在 |
| `IndexError` (merge/sections/comments) | 幻灯片索引越界或节不存在 | 检查 1-based 索引是否在有效范围内 |
| `ValueError` (watermark) | opacity 不在 [0, 1] 范围 | 传入 0.0–1.0 之间的浮点数 |
| `ValueError` (image_optimize) | 不支持的格式或 quality 超范围 | 使用 JPEG/PNG，quality 1–100 |
| `ValueError` (smartart) | 形状不是 SmartArt 或无法访问 diagramData | 先用 `detect_smartart()` 确认 |
| `ValueError` (vba) | PPTX 无法作为 ZIP 读取或宏名无效 | 确认文件是有效 PPTX，宏名非空 |
| `ValueError` / `RuntimeError` (table_styles) | 表格为空、无 tblGrid 或索引越界 | 确保表格有至少一行一列 |
| `ValueError` (html_import) | HTML 文件不存在或格式无法识别 | 检查 HTML 路径和内容格式 |
| `RuntimeError` (export) | 渲染后端不可用 | 安装 LibreOffice 或确认 COM 可用 |

## 模块索引

| 模块 | 职责 |
|---|---|
| `pptx_skill.api` | 兼容 facade 与 `GenerationResult` |
| `pptx_skill.content_model` | `ContentSpec`/`SlideSpec`/`ElementSpec` 数据模型 |
| `pptx_skill.content_adapter` | 适配 legacy `Section` 对象到 `ContentSpec` |
| `pptx_skill.layout_engine` | 声明式 `LayoutRecipe` + Kiwi 约束求解 |
| `pptx_skill.deck_planner` | 整本 beam search 节奏规划 |
| `pptx_skill.pagination` | bullets/table/timeline/process/image_grid 分页 |
| `pptx_skill.text_metrics` | Pillow 字体度量、CJK 换行、字号二分搜索 |
| `pptx_skill.semantic_qa` | 四级 QA：表格/图表/页面/整本（溢出/重叠/对比度/排版层级/色彩一致性） |
| `pptx_skill.render_qa` | 像素级 perceptual diff、窗口密度、SSIM |
| `pptx_skill.repair_engine` | 白名单修复动作与 profile override 合并 |
| `pptx_skill.generation_pipeline` | plan → render → QA → repair 闭环 |
| `pptx_skill.reference_adapter` | native/clone/visual-rebuild adapter |
| `pptx_skill.visual_rebuild` | 扁平稿解析与自适应重建 |
| `pptx_skill.qa_dataset` | 200 页标注数据集生成器 |
| `pptx_skill.golden_renders` | 每 role golden render 与 family deck |
| `pptx_skill.manifest` | Manifest V3 读写与迁移 |
| `pptx_skill.pptx_renderer` | 自适应渲染器（text/image/shape/table/chart/rich-text/media/notes/headers-footers） |
| `pptx_skill.template_compiler` | TemplateProfileV2 编译与多样性闸门 |
| `pptx_skill.template_downloader` | 模板包下载（GitHub/URL/本地）、PPTX 档案提取、远程包列表与搜索 |
| `pptx_skill.design_schema` | V2 schema、token 解析与校验 |
| `pptx_skill.image_crop` | contain/cover/smart crop + 真实 crop fractions |
| `pptx_skill.preview_renderer` | LibreOffice / PyMuPDF / Windows COM 渲染截图 |
| `pptx_skill.visual_qa` | 渲染后视觉 QA 检查 |
| `pptx_skill.capability` | 运行时环境能力检测与报告 |
| `pptx_skill.transitions` / `transitions_ext` / `morph` | 30 种转场动画（18 标准 + 12 高级，另有 Morph） |
| `pptx_skill.animations` | 51 种幻灯片动画（20 入场/15 退出/15 强调/1 运动路径）via OOXML timing XML |
| `pptx_skill.merge` | 多文件合并与幻灯片提取，布局/媒体/rel-ID 去重 |
| `pptx_skill.watermark` | 文字/图片水印，透明度/平铺/层级/移除 |
| `pptx_skill.sections` | 节管理（增删改移折叠）via OOXML sectionLst |
| `pptx_skill.slide_master` | 母版/版式查询、克隆、重命名、占位符操作、背景设置 |
| `pptx_skill.image_optimize` | 图片压缩、格式转换、未用媒体清理、统计 |
| `pptx_skill.comments` | 演讲者备注 + 审阅批注（线程回复/解决）via OOXML |
| `pptx_skill.diff` | 结构化差异比较（幻灯片/形状/文字/图片/版式级别） |
| `pptx_skill.export` | PDF/图片/HTML/文本/缩略图导出（LibreOffice/PyMuPDF/Pillow 后端） |
| `pptx_skill.smartart` | SmartArt 检测、文本提取/编辑、跨页保留 |
| `pptx_skill.vba` | VBA 宏检测、注入、提取、宏-形状绑定 |
| `pptx_skill.table_styles` | 20 种内置表格样式、单元格合并/拆分、边框/填充/行列增删 |
| `pptx_skill.html_import` | HTML 幻灯片导入（Marp/Slidev/reveal.js），语义提取 + 可选光栅化底图 |

## 版式与视觉原则

- 使用 24 种版式形成节奏，不连续堆叠同一种内容页。
- 让标题与正文产生明显字号跳跃，避免平均递减。
- 控制主色数量，使用一个识别色和克制的辅助色。
- 保留 20% 至 30% 留白，不用内容填满画布。
- 使用与论点直接相关的图片、图表和数据，不添加随机装饰图。
- 不把每组信息都包进卡片；优先使用网格、细线、轴线、留白、图像裁切和真实信息层级组织页面。
- 生成多套模板时必须改变几何骨架、信息入口和页面节奏，不能只换颜色或字体。
- 保持封面、章节页和结尾页在同一视觉家族内。
- 复刻参考稿时优先保持原始画幅、母版、Logo、安全边距和页脚规则。

## 支持边界

- `python-pptx` 适合静态可编辑页面、图片、表格和常见图表。
- 样例页克隆可保留常用形状、图片及图表关系，但复杂动画、宏、OLE、嵌入工作簿和部分 SmartArt 必须在 PowerPoint 中复核。
- 自动内容绑定适合简单封面和标题内容页；复杂仪表盘、流程图和多区域页面必须使用显式形状映射。
- 无法从扁平截图恢复原始矢量、动画或不可见数据；应重建可编辑近似版本并明确说明。
- Pixabay 搜索失败时降级为本地图片或纯文字版式，不中断生成。
- 动画和转场通过直接操作 OOXML XML 实现，复杂组合动画需在 PowerPoint 中验证。
- VBA 宏注入以二进制 `vbaProject.bin` 整体注入，无法编辑宏内容本身。
- SmartArt 文本编辑通过操作 diagramData XML 实现，布局/样式变更需在 PowerPoint 中完成。
- HTML 导入依赖可选的 Playwright 后端进行光栅化，无 Playwright 时仅提取语义文本。

## 研发状态

| PR | 内容 | 测试基线 |
|---|---|---|
| PR0 | unittest 基线恢复、`pptx_skill` 包壳、`pyproject.toml`、渲染器隔离、运行能力报告 | — |
| PR1 | 核心数据模型、稳定 ID、Manifest V3、legacy v2 迁移、`api.py` 兼容 facade | — |
| PR2 | preview_renderer、image_crop、adaptive pptx_renderer 基础 | 40/40 |
| PR3 | text_metrics、semantic_qa | 54/54 |
| PR4 | layout_engine 约束求解、候选评分 | 63/63 |
| PR5 | render_qa、repair_engine、generation_pipeline 闭环 | 75/75 |
| PR6 | pagination、deck_planner beam search、多页渲染 | 92/92 |
| PR7 | TemplateProfileV2、编译器、V1/V2 迁移与适配、多样性闸门 | 117/117 |
| PR8a | 14 role 全部 recipe、role-specific paginator、table/chart 节点 | 138/138 |
| PR8b | reference_adapter native/clone、clone drift QA sidecar | 147/147 |
| PR8c | visual_rebuild 真实自适应重建、参考差异预算 | 164/164 |
| PR9 | 200 页 QA 数据集、golden renders、E2E/压力/性能、第三方许可说明 | 217/217（历史基线；当前套件 250 个测试） |
| PR10 | 12 新模块：animations、merge、watermark、sections、slide_master、image_optimize、comments、diff、export、smartart、vba、table_styles | — |
| PR11 | 5 分析框架 role（swot/porter/pest/bmc/funnel）+ consulting 风格 recipe 变体 + html_import | — |

## 相关文档

- [references/reference-ppt-workflow.md](references/reference-ppt-workflow.md) — 处理用户提供的参考稿。
- [references/editing-workflow.md](references/editing-workflow.md) — 修改已有 PPT。
- [references/engine-api.md](references/engine-api.md) — 生成 API 与模板档案。
- [references/quality-checklist.md](references/quality-checklist.md) — 交付前检查清单。
- [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) — 第三方依赖与许可证。
