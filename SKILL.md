---
name: pptx
version: "2.0.0"
description: |
  创建、编辑、重构和验收可编辑 PowerPoint 演示文稿。用于从主题或资料生成 PPT、自动排版、选择或生成模板、按照用户提供的 .pptx 母版或样例页高保真复刻、替换指定文字/图片/图表、制作数据可视化、渲染截图并迭代修正。支持 14 种内容版式、12 套模板档案、3 套非卡片式版式家族、9 套传统主题、自然语言生成新模板、母版占位符填充和样例页逐形状克隆。
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
| 母版占位符生成（native） | 已落地 | `pptx_skill.reference_adapter.native_mode_adapter()` |
| 样例页逐形状克隆（clone） | 已落地 | `pptx_skill.reference_adapter.clone_mode_adapter()` |
| 扁平稿视觉重建（visual-rebuild） | 已落地 | `pptx_skill.visual_rebuild.visual_rebuild_adapter()` |
| 约束式自适应排版 | 已落地 | `pptx_skill.layout_engine.solve_recipe()` |
| 整本拆页与 beam search 节奏 | 已落地 | `pptx_skill.deck_planner.plan_deck()` |
| 文字度量与字号适配 | 已落地 | `pptx_skill.text_metrics.measure_runs()` |
| 语义 QA（溢出/重叠/对比度等） | 已落地 | `pptx_skill.semantic_qa.SemanticQAEngine` |
| 渲染 QA 与视觉回归 | 已落地 | `pptx_skill.render_qa.compare_slide_to_baseline()` |
| 白名单修复引擎 | 已落地 | `pptx_skill.repair_engine.propose_repairs()` |
| 生成 → QA → 修复闭环 | 已落地 | `pptx_skill.generation_pipeline.run_generation_pipeline()` |
| Manifest V3 读写 | 已落地 | `pptx_skill.manifest.load_manifest()` |
| 定点文字/图片/颜色编辑 | 已落地 | `pptx_skill.ppt_edit.*` |
| 页面增删移复制换版式 | 已落地 | `pptx_skill.ppt_pages.*` |
| 200 页 QA 标注数据集 | 已落地 | `pptx_skill.qa_dataset.generate_annotation_dataset()` |
| Golden renders 基线 | 已落地 | `pptx_skill.golden_renders.render_golden_set()` |

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
python -m pip install -e ".[adaptive,schema,qa-image,render-pdf]"
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

不要把“从参考稿提取主题色”当作“完全复刻”。当精确位置重要时，使用分析结果中的形状名称编写 shape-level plan。

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

- 指标 → `dashboard`
- 时间节点 → `timeline`
- 两侧信息 → `comparison`
- 步骤 → `process`
- 表格 → `table`
- 引用 → `quote`
- 多图 → `image_grid`
- 图文 → `text_image`
- 普通要点 → `bullets`

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

逐页检查文字溢出、遮挡、裁图、对比度、层级、旧内容残留和模板一致性。至少执行两轮“生成 → 渲染 → 逐页检查 → 修正”，连续两轮没有阻断问题后再交付。

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
- 所有写操作自动生成最近一次 `.bak.pptx` 备份；修改失败时调用 `restore_backup()`。
- 复杂页面先分析形状名称，再用 shape-level plan 或显式 Section 数据，不依赖模糊角色猜测。

## 错误处理

| 问题 | 典型原因 | 处理方式 |
|---|---|---|
| `ImportError: kiwisolver` | 未安装 adaptive 依赖 | `pip install -e ".[adaptive]"` 或回退 `layout_engine="legacy"` |
| 求解器不可行（infeasible） | 内容超长、约束冲突或 recipe 不匹配 | 换 recipe、分页或放宽内容限制 |
| 参考稿模式推荐 `visual-rebuild` | 母版/占位符不可用，形状都在页面上 | 按视觉重建路径执行，并给出相似度预算报告 |
| QA 报 `blocker` | 溢出、重叠、越界、低对比度等 | 进入 repair 闭环，最多两轮；仍失败则明确返回 |
| 渲染引擎不可用 | 未安装 LibreOffice/PyMuPDF/PowerPoint COM | 标记 unavailable，不影响结构/语义 QA；需要时再安装 |
| 图片搜索失败 | Pixabay 网络或关键词问题 | 降级为本地图片或纯文字版式，不中断生成 |

## 模块索引

| 模块 | 职责 |
|---|---|
| `pptx_skill.api` | 兼容 facade 与 `GenerationResult` |
| `pptx_skill.content_model` | `ContentSpec`/`SlideSpec`/`ElementSpec` 数据模型 |
| `pptx_skill.layout_engine` | 声明式 `LayoutRecipe` + Kiwi 约束求解 |
| `pptx_skill.deck_planner` | 整本 beam search 节奏规划 |
| `pptx_skill.pagination` | bullets/table/timeline/process/image_grid 分页 |
| `pptx_skill.text_metrics` | Pillow 字体度量、CJK 换行、字号二分搜索 |
| `pptx_skill.semantic_qa` | 溢出/重叠/对比度/缺字体/图片畸变/空内容 |
| `pptx_skill.render_qa` | 像素级 perceptual diff、窗口密度、SSIM |
| `pptx_skill.repair_engine` | 白名单修复动作与 profile override 合并 |
| `pptx_skill.generation_pipeline` | plan → render → QA → repair 闭环 |
| `pptx_skill.reference_adapter` | native/clone/visual-rebuild adapter |
| `pptx_skill.visual_rebuild` | 扁平稿解析与自适应重建 |
| `pptx_skill.qa_dataset` | 200 页标注数据集生成器 |
| `pptx_skill.golden_renders` | 每 role golden render 与 family deck |
| `pptx_skill.manifest` | Manifest V3 读写与迁移 |
| `pptx_skill.pptx_renderer` | 自适应渲染器（text/image/shape/table/chart） |
| `pptx_skill.template_compiler` | TemplateProfileV2 编译与多样性闸门 |
| `pptx_skill.design_schema` | V2 schema、token 解析与校验 |
| `pptx_skill.image_crop` | contain/cover/smart crop + 真实 crop fractions |

## 版式与视觉原则

- 使用 14 种版式形成节奏，不连续堆叠同一种内容页。
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
| PR9 | 200 页 QA 数据集、golden renders、E2E/压力/性能、第三方许可说明 | 184/184 |

## 相关文档

- [references/reference-ppt-workflow.md](references/reference-ppt-workflow.md) — 处理用户提供的参考稿。
- [references/editing-workflow.md](references/editing-workflow.md) — 修改已有 PPT。
- [references/engine-api.md](references/engine-api.md) — 生成 API 与模板档案。
- [references/quality-checklist.md](references/quality-checklist.md) — 交付前检查清单。
- [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) — 第三方依赖与许可证。
