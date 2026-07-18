---
name: pptx
description: |
  创建、编辑、重构和验收可编辑 PowerPoint 演示文稿。用于从主题或资料生成 PPT、自动排版、选择或生成模板、扩展模板库、按照用户提供的 .pptx 母版或样例页高保真复刻、替换指定文字/图片/图表、制作数据可视化、渲染截图并迭代修正。支持 14 种内容版式、12 套模板档案、3 套非卡片式版式家族、9 套传统主题、自然语言生成新模板、母版占位符填充和样例页逐形状克隆。
---

# PPTX 智能制作

始终交付可编辑 `.pptx`，并执行渲染验收。不要只检查代码或 XML。

## 按任务选择路径

1. 用户没有提供参考 PPT：
   - 直接使用已有模板档案生成；或
   - 根据行业、品牌色、受众和语气生成新模板档案，再生成演示文稿。
2. 用户提供了规范母版或企业模板：
   - 分析母版、版式和占位符；
   - 使用 `native` 模式从原生版式创建页面。
3. 用户提供了样例演示文稿，设计主要画在页面上：
   - 使用 `clone` 模式复制代表页；
   - 按形状名称精确替换文字、图片和内容区域。
4. 用户提供的是扁平图片、复杂 SmartArt/OLE 或不可复用结构：
   - 使用 `visual-rebuild` 路径；
   - 按截图重建可编辑元素并进行并排验收。

处理用户提供的 PPT 时，完整阅读 [references/reference-ppt-workflow.md](references/reference-ppt-workflow.md)。修改已有文件时，完整阅读 [references/editing-workflow.md](references/editing-workflow.md)。调用生成 API 或管理模板档案时，阅读 [references/engine-api.md](references/engine-api.md)。交付前阅读 [references/quality-checklist.md](references/quality-checklist.md)。

## 前置检查

从项目根目录安装为可编辑包（推荐）：

```powershell
python -m pip install -e .
```

然后检查当前环境能力：

```powershell
python -m pptx_skill.capability
```

确认以下依赖可用：

```powershell
python -m pip show python-pptx Pillow
python scripts/render_slides.py --help
```

截图渲染优先使用 LibreOffice；Windows 也可回退到 PowerPoint COM。缺少渲染引擎时再安装 LibreOffice，不要无条件安装。

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
from pptx_helper import auto_validate_ppt
result = auto_validate_ppt("output/report.pptx")
```

渲染所有页面：

```powershell
python scripts/render_slides.py output/report.pptx `
  --output output/report-preview --dpi 150
```

逐页检查文字溢出、遮挡、裁图、对比度、层级、旧内容残留和模板一致性。至少执行两轮“生成 → 渲染 → 逐页检查 → 修正”，连续两轮没有阻断问题后再交付。

## 修改已有 PPT

- 对本 Skill 生成的文件，优先读取内嵌 manifest：
  - `from pptx_skill import load_project, edit_section, regenerate`
  - 改 Section 后调用 `regenerate(project, output_path)` 重画。
- 对外部文件，使用 `pptx_skill.ppt_edit` 定点换字、换图或换色：
  - `edit_text` / `edit_text_by_role` / `swap_image` / `recolor` / `swap_theme`。
- 使用 `pptx_skill.ppt_pages` 增页、删页、移动、复制或重画单页：
  - `insert_slide` / `delete_slide` / `move_slide` / `duplicate_slide` / `replace_layout`。
- 所有写操作自动生成最近一次 `.bak.pptx` 备份；修改失败时调用 `restore_backup()`。
- 复杂页面先分析形状名称，再用 shape-level plan 或显式 Section 数据，不依赖模糊角色猜测。

## 研发状态（PR0 完成）

- PR0 已落地：unittest 基线恢复、`pptx_skill` 包壳、`pyproject.toml`、渲染器隔离、运行能力报告。
- 后续 PR（V2 数据模型、自适应排版、视觉 QA、模板引擎 V2）将按蓝图逐步实现。

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
