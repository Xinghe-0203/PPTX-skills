# pptx-skill

> 面向智能体工作流的 PowerPoint 生成、编辑与验收 Skill。
>
> 核心原则：**始终交付可编辑的 `.pptx`，并执行渲染验收**，而不只是检查代码或 XML。

## 功能特性

- **从零生成**：根据主题、章节、数据自动创建完整演示文稿，支持 14 种内容版式、12 套模板档案、3 套非卡片式几何语言。
- **母版占位符生成（native）**：按企业母版或规范模板创建页面。
- **样例页克隆（clone）**：高保真复制参考 PPT 的形状关系并替换内容。
- **视觉重建（visual-rebuild）**：从扁平截图或复杂结构重建可编辑近似版本。
- **约束式自适应排版**：基于 Kiwi 线性约束求解器的声明式 `LayoutRecipe`，覆盖 14 种 role 的全部 recipe。
- **整本规划**：`deck_planner` 使用 beam search 规划跨页节奏。
- **语义与渲染 QA**：溢出/重叠/对比度检查、像素级 perceptual diff、SSIM、窗口密度。
- **生成 → QA → 修复闭环**：`run_generation_pipeline` 自动迭代。
- **数据契约与 Manifest V3**：稳定 element ID、内嵌 XML + 侧写 `.manifest.json`。
- **200 页 QA 标注数据集与 Golden Renders**：用于回归测试与视觉基线。

## 安装

推荐以可编辑模式安装：

```powershell
python -m pip install -e .
```

检查当前环境能力：

```powershell
python -m pptx_skill.capability
```

可选后端按需安装：

```powershell
python -m pip install -e ".[adaptive,schema,qa-image,render-pdf]"
```

> 核心路径仅依赖 `python-pptx` 与 `Pillow`；其它后端均为可选。详见 [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)。

## 快速开始

### 1. 生成一份新 PPT

```python
from pptx_skill import auto_generate_ppt

auto_generate_ppt(
    title="市场拓展策略",
    subtitle="2026 年度规划",
    sections=[...],
    output_path="output/report.pptx",
    template_key="strategy-consulting",
    auto_search_images=False,
)
```

### 2. 编辑本 Skill 生成的 PPT（无损往返）

```python
from pptx_skill import load_project, edit_section, regenerate

project = load_project("output/report.pptx")
edit_section("output/report.pptx", index=2, changes={"bullets": ["新要点 A", "新要点 B"]})
```

### 3. 编辑外部 PPT（有损定点编辑）

```python
from pptx_skill.ppt_edit import edit_text, swap_image, recolor

edit_text("reference.pptx", slide_index=0, find="旧标题", replace="新标题")
recolor("reference.pptx", old_hex="#184E77", new_hex="#2D5016")
```

### 4. 复刻参考 PPT

```powershell
python scripts/render_slides.py reference.pptx --output output/reference-preview --dpi 150
python scripts/reference_ppt.py analyze reference.pptx --output output/reference-analysis.json
python scripts/reference_ppt.py generate reference.pptx --content content.json --mode clone --output output/report.pptx
```

## 核心概念

| 概念 | 说明 |
|---|---|
| `ContentSpec` / `SlideSpec` / `ElementSpec` | 整本 / 单页 / 页内元素的数据模型，元素带稳定 `element_id`。 |
| `LayoutRecipe` | 声明式版式配方，由变量、约束、几何签名组成，经 Kiwi 求解为 `SolvedGeometry`。 |
| `TemplateProfileV2` | 三层 token（design / semantic / component）描述的模板档案，支持 V1 迁移。 |
| `Manifest V3` | 每次生成、QA、修复的记录，可内嵌到 pptx XML part 或伴随 `.manifest.json`。 |
| `Reference Modes` | `native` / `clone` / `visual-rebuild` 三种参考稿处理策略。 |

## 项目结构

```
.
├── pptx_skill/               # 核心 Python 包
│   ├── api.py                # 兼容 facade 与 GenerationResult
│   ├── content_model.py      # ContentSpec / SlideSpec / ElementSpec
│   ├── layout_engine.py      # LayoutRecipe + Kiwi 约束求解
│   ├── deck_planner.py       # 整本 beam search 节奏规划
│   ├── pagination.py         # bullets / table / timeline / process / image_grid 分页
│   ├── text_metrics.py       # 字体度量、CJK 换行、字号适配
│   ├── semantic_qa.py        # 溢出 / 重叠 / 对比度等语义检查
│   ├── render_qa.py          # 像素级 perceptual diff
│   ├── repair_engine.py      # 白名单修复动作
│   ├── generation_pipeline.py # plan → render → QA → repair 闭环
│   ├── reference_adapter.py  # native / clone / visual-rebuild adapter
│   ├── visual_rebuild.py     # 扁平稿解析与自适应重建
│   ├── qa_dataset.py         # 200 页 QA 标注数据集
│   ├── golden_renders.py     # 每 role golden render 与 family deck
│   ├── manifest.py           # Manifest V3 读写
│   └── template_compiler.py  # TemplateProfileV2 编译与多样性闸门
├── scripts/                  # 命令行入口
│   ├── pptx_helper.py        # 生成引擎
│   ├── reference_ppt.py      # 参考稿分析与生成
│   ├── template_engine.py    # 模板档案管理
│   ├── render_slides.py      # 截图渲染
│   ├── ppt_project.py        # L1 manifest 往返
│   ├── ppt_inspect.py        # L2 外部 PPT 解析
│   ├── ppt_edit.py           # L3 定点编辑
│   └── ppt_pages.py          # L4 页级操作
├── references/               # 详细工作流文档
│   ├── reference-ppt-workflow.md
│   ├── editing-workflow.md
│   ├── engine-api.md
│   └── quality-checklist.md
├── tests/                    # 单元测试与 E2E 测试
├── SKILL.md                  # Skill 使用手册（面向智能体）
├── THIRD_PARTY_NOTICES.md    # 第三方依赖与许可证
└── pyproject.toml
```

## 测试

运行完整测试套件：

```powershell
python -m unittest discover -s tests -v
```

当前基线：**184 / 184 通过**。

PR 级测试基线：

| PR | 内容 | 测试基线 |
|---|---|---|
| PR0 | unittest 基线、`pptx_skill` 包壳、运行能力报告 | — |
| PR1 | 核心数据模型、Manifest V3、兼容 facade | — |
| PR2 | preview_renderer、image_crop、adaptive pptx_renderer | 40/40 |
| PR3 | text_metrics、semantic_qa | 54/54 |
| PR4 | layout_engine 约束求解 | 63/63 |
| PR5 | render_qa、repair_engine、generation_pipeline | 75/75 |
| PR6 | pagination、deck_planner beam search | 92/92 |
| PR7 | TemplateProfileV2、编译器、V1/V2 迁移 | 117/117 |
| PR8a | 14 role 全部 recipe、role-specific paginator | 138/138 |
| PR8b | reference_adapter native/clone | 147/147 |
| PR8c | visual_rebuild 真实自适应重建 | 164/164 |
| PR9 | 200 页 QA 数据集、golden renders、E2E/压力/性能 | 184/184 |

## 交付前检查

- 渲染所有页面并逐页检查溢出、遮挡、裁图、对比度、层级和模板一致性。
- 至少执行两轮“生成 → 渲染 → 检查 → 修正”。
- 详见 [`references/quality-checklist.md`](references/quality-checklist.md)。

## 支持边界

- 适合静态可编辑页面、图片、表格和常见图表。
- 不保证保留复杂动画、宏、OLE、嵌入工作簿及部分 SmartArt 的完整行为。
- 无法从扁平截图恢复原始矢量、动画或不可见数据；应重建可编辑近似版本。
- Pixabay 搜索失败时降级为本地图片或纯文字版式，不中断生成。

## 文档索引

- [`SKILL.md`](SKILL.md) — 面向智能体的完整使用手册。
- [`references/reference-ppt-workflow.md`](references/reference-ppt-workflow.md) — 处理用户提供的参考稿。
- [`references/editing-workflow.md`](references/editing-workflow.md) — 修改已有 PPT。
- [`references/engine-api.md`](references/engine-api.md) — 生成 API 与模板档案。
- [`references/quality-checklist.md`](references/quality-checklist.md) — 交付前检查清单。
- [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md) — 第三方依赖与许可证。

## 许可证

本项目核心路径使用 MIT 风格授权（详见仓库 LICENSE）。使用可选后端时，请遵守对应依赖的许可证要求，尤其是 PyMuPDF 的 AGPL/商业授权风险。

---

*Generated for pptx-skill v2.0.0.*
