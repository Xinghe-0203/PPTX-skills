# PPTX Skill 下一阶段能力升级提案

> 状态：Proposal — 待评审与方向决策
>
> 日期：2026-07-19
>
> 范围：v2.0.0（PR0–PR9 全部落地）之后的五个升级方向
>
> 目标读者：负责评审、排期和实现本 Skill 后续迭代的工程师或 Agent

## 0. 背景

`PPTX_ADVANCED_IMPLEMENTATION_PLAN.md` 规划的三大能力升级（真正的视觉 QA、约束式自适应排版、生成式模板引擎 V2）已随 PR0–PR9 全部落地，测试基线 184/184。当前 v2.0.0 已具备：

- 14 种内容版式、12 套模板档案、3 套非卡片式几何语言、9 套传统主题；
- native / clone / visual-rebuild 三种参考稿处理路径；
- LayoutRecipe + Kiwi 约束求解、deck_planner beam search 跨页节奏；
- text_metrics、semantic_qa、render_qa、repair_engine 与生成闭环；
- Manifest V3、L1–L4 修改层、200 页 QA 标注数据集与 golden renders。

本文列出下一阶段五个候选升级方向，给出每个方向的现状缺口、方案内容、成本风险和验收方式，并给出推荐排序。**本文是方向提案，不是实现设计**；方向确定后，每个方向应按本项目惯例单独出设计文档并按 PR 拆步。

## 1. 方向总览

| 方向 | 解决的核心缺口 | 预期收益 | 成本 | 风险 |
|---|---|---|---|---|
| 一、闭环严格化 | strict 只报告不阻断；无 PowerPoint 确认通道 | 已有投资直接变现，每份产出可信度提升 | 低 | 低 |
| 二、引入视觉模型 | visual_rebuild 启发式精度有限；规则 QA 抓不住构图/美观 | 复刻保真度与 QA 覆盖面上台阶 | 高 | 中（原则冲突，见 §5） |
| 三、内容 ingestion | 内容全靠调用方手写 sections，无资料入口 | 适用面从"给大纲"扩到"扔资料" | 中 | 低 |
| 四、图表与 OOXML 深水区 | 图表选型手工；SmartArt 只能整体放弃 | 数据页质量与外部文件兼容面 | 中 | 中（动画部分建议不做，见 §6） |
| 五、模板生态 | clone 是一次性复刻，不能沉淀为模板 | 企业模板资产化 | 中 | 低 |

## 2. 方向一：把 QA 闭环真正"关上"

### 2.1 现状缺口

- `qa_mode="strict"` 当前只输出报告，不阻断交付（SKILL.md 明确标注"strict 暂未启用阻断"）；repair 两轮失败后直接放行。安全网已织好但未通电。
- 当前环境无 PowerPoint COM。LibreOffice-only 路径下 `TEXT_OVERFLOW_CONFIRMED` 无法确证，只能返回 `inconclusive`，文字溢出这一最高频问题停留在"估测"级别。
- 两个渲染引擎（LibreOffice / PowerPoint）按设计不能共用 golden 基线，当前只有 LibreOffice 一侧的基线。

### 2.2 方案内容

1. 用已有的 200 页标注数据集校准各检查项阈值（重叠交集率、对比度、窗口差异密度、SSIM 等），把 `strict` 模式的阻断逻辑启用；`inconclusive` 是否升格为 blocker 由 profile 配置决定。
2. 在装有 Office 的 Windows 环境接入 PowerPoint COM：通过 `TextFrame2.TextRange.BoundWidth/BoundHeight` 读取真实排版尺寸，将估测溢出升级为确认溢出；修复 `render_with_com` 的 Export DPI 参数传递问题（`Presentation.Export(Path, FilterName, ScaleWidth, ScaleHeight)`）。
3. 建立 `tests/golden/powerpoint/<version>/<font-pack>/` 基线目录，与 LibreOffice 基线隔离。
4. 扩大 repair 白名单动作集（在保持"只允许白名单动作、每步写 manifest"约束的前提下）。

### 2.3 成本与风险

- 成本低：所有构件（数据集、QA 引擎、COM 路径骨架）均已存在，主要是校准、接线和阈值调参。
- 风险低：不改变现有数据契约；strict 默认行为变化需保留 `report` 模式作为兼容选项。

### 2.4 验收

- 标注数据集上 blocker 检出率 / 误报率达到文档化阈值；
- 含 PowerPoint 的环境中 `TEXT_OVERFLOW_CONFIRMED` 可由 COM 通道确证；
- 双引擎 golden 基线各自独立通过。

## 3. 方向二：引入视觉模型（VLM）

### 3.1 现状缺口

- `visual_rebuild` 当前是纯启发式：Laplacian 边缘、饱和度、肤色近似图 + 可选 OpenCV 人脸检测。对扁平截图的版面结构理解能力有限，重建保真度存在明显天花板。
- semantic_qa / render_qa 只能发现规则可描述的问题（溢出、重叠、对比度、回归）。构图失衡、视觉层级混乱、"看起来不美"这类问题没有检查手段。

### 3.2 方案内容

分两步走，先做只读、后做写入：

1. **VLM-QA 评委（只读，零风险）**：渲染 PNG 后，由多模态模型对构图、层级、留白、美观度打 advisory 分数并指出问题区域。产出只作为建议，不产生 blocker；blocker 仍全部由确定性检查产出。
2. **VLM 辅助视觉重建（写入，需验证）**：截图 → VLM 解析版面结构（区域检测 + OCR + 配色提取）→ 产出 `ElementSpec` 草稿树 → 交给现有 layout_engine 求解 + QA 验证。VLM 不直接产出坐标或 OOXML，只产出受 schema 约束的结构草稿。

### 3.3 与原设计原则的关系（必须先决策）

现有设计文档明确写着两个非目标：

- "用一个端到端黑盒模型替换可解释、可测试的排版和验收逻辑"；
- "让模型直接生成未经验证的 OOXML、Python 或 JavaScript 并执行"。

本方向**不触碰**这两条：VLM 定位为 advisory 评委和草稿生成器，最终排版仍由 solver 完成，最终质量仍由确定性 QA 把关。但"引入模型参与生成链路"本身是原则之外的新事物，是否接受需要项目所有者明确拍板，并同步修订设计文档的非目标章节。

### 3.4 成本与风险

- 成本高：需要模型接入、prompt/输出 schema 设计、解析失败的降级路径、以及新一批 VLM 输出的标注验收集。
- 风险中：模型输出不稳定；必须用 schema 校验 + solver 兜底 + QA 复验三层约束，不允许 VLM 输出绕过现有闭环。
- 新增运行时依赖（模型 API），应按本项目惯例做成可选 extra，核心路径保持 `python-pptx + Pillow` 可独立运行。

### 3.5 验收

- VLM-QA：在标注集上与人工评审结论的一致率达到文档化阈值；advisory 输出永不阻断交付。
- VLM 重建：对一组参考截图，重建稿与截图的相似度预算报告显著优于当前启发式路径，且全部通过确定性 QA。

## 4. 方向三：内容 ingestion 管线

### 4.1 现状缺口

Skill 的边界是"怎么画"，内容完全由调用方（agent）手写 sections。"扔一份资料生成 PPT"目前依赖 agent 即兴阅读和组织，没有契约化的入口，质量和结构不可复现。

### 4.2 方案内容

在现有 `LegacyContentAdapter` 同构的位置新增上游适配层：

1. **文档适配器**：docx / PDF / Markdown / 网页 → 抽取章节层级、要点、表格、数据 → 产出 `ContentSpec` 草稿（带 provenance，可追溯每页内容来自原文哪个位置）。
2. **数据适配器**：csv / xlsx → 按数据形状（维度、基数、时间性）判断该用 dashboard / table / chart 哪种版式，产出对应 `SlideSpec`。
3. 草稿进入现有生成管线，与手写 sections 走完全相同的 plan → render → QA → repair 路径，不新增特权通道。

### 4.3 成本与风险

- 成本中：解析器可用成熟库（python-docx、PyMuPDF 已在可选依赖中、markdown 解析），主要工作是内容结构化规则与 provenance 设计。
- 风险低：纯增量模块，不碰核心排版与 QA；失败时降级为"返回抽取结果由 agent 人工组织"，不中断流程。
- 注意许可证：PDF 解析若用 PyMuPDF 需遵守其 AGPL/商业授权约束，延续 `THIRD_PARTY_NOTICES.md` 的现有记录方式。

### 4.4 验收

- 三类来源资料各一组 fixture：产出的 `ContentSpec` 通过 schema 校验，生成稿通过 QA；
- provenance 字段可从成品 manifest 回溯到原文位置；
- 解析失败时按约定降级，不抛未处理异常。

## 5. 方向四：图表与 OOXML 深水区（有边界的增强）

### 5.1 建议做

- **图表自动选型**：按数据形状（维度数、基数、是否时间序列、是否构成占比）自动选择 bar / line / pie / scatter / 组合图；数据标签与坐标轴的默认策略随模板 token 配置。
- **SmartArt 近似重建**：遇到外部文件中的 SmartArt 时，不再整体放弃，而是重建为等价的 group shapes（可编辑近似），与 visual-rebuild 的"诚实降级 + 明确说明"策略保持一致。

### 5.2 建议不做

- **动画、宏、OLE、嵌入工作簿**：python-pptx 不具备写入能力，只能注入裸 OOXML，与"不生成并执行未经验证的 OOXML"原则直接冲突。建议只做一件事：**检测并在分析报告中明确警告**（`analyze` 输出中列出该文件包含的动画/宏/OLE/SmartArt 清单及其将丢失的风险），不尝试保留或重建。

### 5.3 验收

- 图表选型规则有单元测试覆盖；生成图表在 PowerPoint 中保持原生可编辑；
- SmartArt 近似重建产出"可编辑近似 + 差异说明"，通过渲染 QA；
- 含动画/宏/OLE 的外部文件在 analyze 阶段产出明确警告，不静默丢失。

## 6. 方向五：模板生态

### 6.1 现状缺口

clone 模式是一次性复刻：复刻完一份企业 PPT，设计知识（色板、字号阶梯、网格、页脚规则）不沉淀。下一个同企业的需求还要重新 clone。

### 6.2 方案内容

1. **模板归纳（template induction）**：从用户提供的真实 PPT 反向提取 design tokens（色板、字体、字号阶梯、间距）、几何骨架（网格/轴线/栏）和页面节奏规则，自动产出 `TemplateProfileV2` 草稿，经 template_compiler 编译为可复用 bundle。相当于 clone 模式的抽象化。
2. **品牌 kit 导入**：logo / 字体声明 / 色板文件 → tokens 的批量导入入口。
3. **暗色模式变体**：利用现有三层 token 结构（primitive → semantic → component），只替换 primitive 层生成暗色变体，验证 token 架构的设计收益。

### 6.3 成本与风险

- 成本中：归纳规则可复用 `reference_adapter` 的分析产物和 `template_compiler` 的校验闸门；多样性闸门（diversity gate）可直接复用，防止归纳出的模板与现有档案同质。
- 风险低：产出物是受 schema 约束的声明式 profile，不生成可执行代码，与本项目既有安全模型一致。

### 6.4 验收

- 对一组企业参考稿，归纳产出的 profile 通过 compiler 校验与多样性闸门；
- 用归纳模板生成的新稿与原参考稿在品牌元素上视觉一致（render QA 基线对比）。

## 7. 推荐实施顺序

| 顺序 | 方向 | 理由 |
|---|---|---|
| 1 | 方向一：QA 闭环严格化 | 收割已有投资，成本低、风险低、每份产出立即受益 |
| 2 | 方向三：内容 ingestion | 纯增量、不碰核心，直接扩大适用面 |
| 3 | 方向二：VLM（先 QA 评委，后视觉重建） | 收益最大但需先决策原则问题；只读先行可零风险验证价值 |
| 4 | 方向四：图表选型 + SmartArt 近似（动画不做） | 挑边界清晰的子项做 |
| 5 | 方向五：模板生态 | 依赖前序积累，余力再做 |

方向二与方向三、四、五之间无强依赖，评审后可按资源并行；方向一建议任何并行计划启动前先完成。

## 8. 非目标（本阶段继续不承诺）

- 完整保留或生成 PowerPoint 动画、宏、OLE、嵌入工作簿；
- LibreOffice 与 PowerPoint 逐像素一致；
- 从扁平截图恢复原始矢量、不可见数据或原动画；
- 让模型（LLM 或 VLM）直接生成并执行未经验证的 OOXML / Python / JavaScript；
- 用端到端黑盒模型替换可解释、可测试的排版与验收逻辑。

方向二若获批，仅在此清单中新增一条例外说明："模型可作为 advisory 评委与受 schema 约束的草稿生成器参与链路，最终排版与质量判定仍由确定性 solver 与 QA 完成"，其余条目不变。

## 9. 后续动作

1. 评审本文，确定方向取舍与排序（重点是方向二的原则决策）；
2. 每个获批方向按本项目惯例单独出实现设计文档（现状探查 → 数据契约 → 模块边界 → PR 拆步 → 测试基线）；
3. 方向一启动前，先确认目标环境是否具备 PowerPoint COM 条件，决定 COM 通道是否纳入首期。
