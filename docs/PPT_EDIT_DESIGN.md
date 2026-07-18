# PPT 修改能力设计文档（L1-L4）

> **状态**：主路径已实现（2026-07-18）。已新增 `ppt_project.py`、`ppt_inspect.py`、`ppt_edit.py`、`ppt_pages.py`，并与 `reference_ppt.py` 的母版/样例页复刻能力衔接。manifest 使用标准 custom properties 嵌入并保留 `.manifest.json` 兜底；所有写操作保留最近一次 `.bak.pptx`。
> **目标读者**：实现这套"修改层"的人（或 agent）。
> **核心原则**：agent 管"想改什么"，skill 管"怎么改到文件里"。skill 提供确定性原语，语义判断交给 agent。

---

## 1. 背景与目标

当前 skill 同时支持从零生成、模板档案生成、参考 PPT 复刻和确定性修改。消费者是**智能体**本身——给 agent 一个已有 PPT，让它"听我的修改"。

本设计新增"修改层"，支持两类来源的 PPT 修改：
- **本 skill 生成的 PPT**：无损往返修改（改内容 → 重画）。
- **外部现成 PPT**：有损解析 + 精确局部编辑。

---

## 2. 现状与约束（探查结论）

探查 `scripts/pptx_helper.py`（1400+ 行）得出以下硬约束，直接决定方案形态：

| 约束 | 影响 |
|---|---|
| 所有 layout 产出的 shape **都是匿名的**（无 `.name`、无自定义 marker、无 notes） | 外部 PPT 无法可靠反向识别"哪个 shape 是标题"，只能启发式猜 |
| 文本走 `_textbox`+`_add_para`，颜色挂 `run.font.color.rgb` | 颜色可读可写，L3 recolor/swap_theme 可行 |
| 背景挂 `slide.background.fill`，shape 填充挂 `shape.fill.fore_color.rgb` | 同上，全局换色可行 |
| 主题色是 9 套 `RGBColor` dict（`THEMES`） | swap_theme 需"旧色→新色"映射表 |
| python-pptx 原生**不支持**删除/重排页面，只支持 `add_slide` 追加到末尾 | L4 增删移需操作 `_sldIdLst` XML |
| `Section` 已是 dataclass（pptx_helper.py:1101） | manifest 可用 `dataclasses.asdict()` 序列化 |
| `auto_generate_ppt` 生成完一切都是临时的，无 manifest 存留 | L1 要补存取 |

**关键推论**：唯一能无损往返的是本 skill 自己生成的 PPT——存 manifest 即可改完重画。外部 PPT 只能有损解析。这把方案劈成两条路。

## 2.5 python-pptx 1.0.2 源码核验要点（本次检查产出）

读本地安装的 python-pptx 源码（`C:\Users\Xinghe\AppData\Roaming\Python\Python314\site-packages\pptx\`，v1.0.2）确认的关键 API，实现时直接照用：

| 关注点 | 源码位置 | 结论 |
|---|---|---|
| `Slides` 有无 delete/move | `slide.py:238-293` | **无**。只有 `add_slide`/`get`/`index`。增删移全靠 `_sldIdLst` |
| `CT_SlideIdList` 增删改 API | `oxml/presentation.py:51-94` | `sldId_lst` 是标准 list，可 `append/insert/remove/pop`；`add_sldId(rId)` 建新元素；`_next_id` 自动 ≥256 |
| 删 slide 的 relationship | `opc/package.py:84` vs `:409` | `OpcPackage.drop_rel`（包级，无条件）vs `XmlPart.drop_rel`（有 ref_count<2 保护）。删 slide 用 `prs.part.package.drop_rel(rId)` 最稳 |
| 官方删元素范式 | `slide.py:390-411` `SlideLayouts.remove` | `list.remove(elem)` + `part.drop_rel(rId)`，删 slide 照搬 |
| `background.fill` 副作用 | `slide.py:471-498` | **访问即设 NoFill 切断继承**。inspect 只读时必须绕开，直接查 `_element.cSld` 下的 `p:bg` |
| custom XML part API | `opc/package.py`、`parts/presentation.py` | **无公开 API**。嵌入需手搓 `Part`+Content-Type override，成本高 → manifest 走 json |
| `PresentationPart` 方法 | `parts/presentation.py` | `add_slide`/`related_slide`/`get_slide`/`slide_id`/`_next_slide_partname` 可用 |
| `Slide.slide_id` | `slide.py:223` | 不随重排变化，可作 manifest 的页标识（但 manifest 用 index 更直观） |

---

## 3. 架构总览

```
用户："改这个 PPT"
  │
  ├─ load_project(path) 成功（有 manifest）
  │    → 路线 A：本 skill PPT，无损往返
  │       · 改某页内容：edit_section → regenerate（只重画该页）
  │       · 增/删/移页：insert_slide / delete_slide / move_slide
  │       · 换某页版式：replace_layout
  │
  └─ load_project(path) 返回 None（外部 PPT）
       → 路线 B：外部 PPT，有损编辑
          · 理解：inspect_ppt → 结构化 dump（供 agent 读）
          · 精确改：edit_text / swap_image / recolor / swap_theme
          · 增页：insert_slide（新页走本 skill 版式）

  两条路改完都必走：render_slides + auto_validate_ppt 复验
```

---

## 4. 文件结构

```
scripts/
├── pptx_helper.py     # 生成引擎（最小侵入改造：加 _save_manifest 钩子 + 抽 _build_slides）
├── ppt_project.py     # L1 记忆往返：manifest 读写 / load_project / regenerate / edit_section
├── ppt_inspect.py     # L2 原始解析：inspect_ppt → 结构化 dump + 版式/主题猜测
├── ppt_edit.py        # L3 定点编辑：edit_text / swap_image / recolor / swap_theme
└── ppt_pages.py       # L4 页级操作：insert / delete / move / replace_layout
```

**依赖方向**（避免循环 import）：
- `ppt_project`、`ppt_inspect`、`ppt_edit`、`ppt_pages` 都从 `pptx_helper` import 共享件（`THEMES`/`LAYOUT_REGISTRY`/`choose_layout`/`_c`/`_text_on`/`Section`/`SLIDE_W` 等）。
- `pptx_helper` **不**在模块级 import 任何新文件。它对 manifest 的依赖通过**延迟导入**（在 `auto_generate_ppt` 函数内部 `from ppt_project import _save_manifest`）实现，彻底避免 `pptx_helper ↔ ppt_project` 循环。
- `ppt_pages.replace_layout` 和 `ppt_edit.swap_theme` 需要 `ppt_inspect`，单向依赖，无环。

import 用相对路径（SKILL.md 已用 `sys.path.insert(0,"scripts")` 模式，各新文件顶部同样处理，保证单独运行也能 import）。

---

## 5. L1 — 记忆往返（本 skill PPT，无损）

**文件**：`scripts/ppt_project.py`

### 5.1 manifest 存储（两层兜底，已据源码定路径）

python-pptx 1.0.2 **没有公开的 custom XML part API**（核验 `opc/package.py`、`parts/presentation.py`：无 `add_custom_xml_part` 之类）。两条可行路径，优先 json：

- **首选：伴随同名 `.manifest.json` 文件**。零风险、零依赖、可读可 diff。`_save_manifest` 写 `<pptx同名>.manifest.json`；`load_project` 读同路径。缺点：文件分离，用户只拷 pptx 会丢 manifest（此时降级为"外部 PPT"路线 B）。
- **备选：嵌入 pptx 内部自定义 Part**。可行但复杂：
  1. 构造 `Part(partname, "application/json", package, blob=json_bytes)`（`opc/package.py:281` 的 `Part` 基类可直接实例化）。
  2. `prs.part.relate_to(custom_part, RT.<自定义reltype>)`（`_RelatableMixin.relate_to`，`opc/package.py:41`）。
  3. 注册 `[Content_Types].xml` override：partname `/ppt/skill_manifest.json` → `application/json`（需操作 `_ContentTypeMap`，python-pptx 不直接暴露，得改写 serialized 层，**实现成本高**）。
  4. 保存时 `PackageWriter.write`（`opc/package.py:156`）会序列化所有 part。
  
  第 3 步是难点。**实现时若第 3 步卡住，直接放弃嵌入路径，全用 json**——功能不丢。

**决策**：第一版只做 json 路径（简单可靠）；嵌入路径列为"增强项"，等 json 跑通后再评估是否值得做。`_save_manifest` / `load_project` 接口不变，内部先 json，未来加嵌入也不影响调用方。

### 5.2 manifest schema

```json
{
  "skill_version": 1,
  "title": "...",
  "subtitle": "...",
  "theme_key": "editorial",
  "image_dir": "abs/path",
  "lang": "zh",
  "auto_search_images": true,
  "cover_image": "abs/path",
  "sections": [ { "Section 全字段 asdict" }, ... ],
  "layouts": ["cover","toc","dashboard",...],
  "generated_at": "由调用方传入的时间戳"
}
```

`sections` 每项 = `dataclasses.asdict(Section实例)`，含全部字段（title/subtitle/bullets/images/image_query/layout/kicker/metrics/events/steps/table_headers/table_rows/left/right/quote/source/section_number/page_number）。

### 5.3 图片可移植性对策（缺口 1）

`Section.images` 存的是生成时的本地绝对路径。换机器/换目录后 `choose_layout` 会因 `os.path.exists` 失败得到 `n_images=0`，从而**误选版式**。对策：

- manifest 里 `images` 同时存 `path`（原绝对路径）和 `rel_name`（相对 `image_dir` 的文件名）。
- `load_project` 读回时，按 `image_dir` + `rel_name` 重定位：若原 `path` 失效，用 `image_dir/rel_name` 兜底；仍失效则标记该图丢失（置空，触发 regenerate 时按 `image_query` 重搜）。
- `regenerate` 内部：若某 section 的图全部失效且有 `image_query`，调 `_auto_search_images` 补搜（复用 pptx_helper 已有逻辑）。

### 5.4 接口签名

```python
def _save_manifest(prs: "Presentation", project: dict) -> None:
    """写 manifest：先试 pptx 内 XML part，失败回退同名 .manifest.json。"""

def load_project(pptx_path: str) -> dict | None:
    """读 manifest：XML 失败试 .json，都没有返回 None。
    读回后做图片路径重定位（见 5.3）。"""

def regenerate(project: dict, output_path: str | None = None) -> str:
    """用 manifest 重放生成。layouts 用 manifest 记的序列（跳过 choose_layout），
    除非某 section 显式改了 layout 字段。返回保存路径。
    内部调 pptx_helper._build_slides 复用渲染逻辑。"""

def edit_section(pptx_path: str, index: int, changes: dict,
                 output_path: str | None = None) -> str:
    """便捷封装：load → 浅合并 changes 到 sections[index] → regenerate。
    changes 支持改任意 Section 字段。返回新文件路径。"""
```

### 5.5 pptx_helper.py 改造（最小侵入）

1. 抽出 `auto_generate_ppt` 里"解析 sections → 逐页 layout_fn"的循环为内部函数 `_build_slides(prs, theme, sections, layouts_override, image_dir, lang, auto_search, cover_image)`，供 `regenerate` 复用。
2. `auto_generate_ppt` 末尾 `prs.save` 前，**延迟导入**并调 `_save_manifest(prs, project_dict)`，让所有新生成 PPT 自动带 manifest。
3. `_build_slides` 接受 `layouts_override: list[str] | None`：非 None 时按它选版式（regenerate 走这条），None 时走 `choose_layout` + 防重复逻辑（原生成走这条）。

### 5.6 技术风险点（已降级）

原方案最大风险是"custom XML part 能否读写"。源码核验后：python-pptx 1.0.2 无公开 custom XML API，嵌入需手搓 Part + Content-Type override，成本高。**改用 json 伴随文件后此风险消除**——json 读写零依赖。

剩余风险仅：
- 图片跨目录重定位（5.3 已给对策，实测时覆盖）。
- 嵌入路径（增强项）若要做，第 3 步 Content-Type override 可能需补丁 python-pptx 序列化层——届时评估收益是否值得。

---

## 6. L2 — 原始解析（外部 PPT，有损）

**文件**：`scripts/ppt_inspect.py`

### 6.1 接口签名

```python
def inspect_ppt(pptx_path: str) -> dict:
    """返回结构化 dump，仅供 agent 阅读理解，不直接用于重画。"""
```

### 6.2 返回结构

```json
{
  "slide_count": 9,
  "width_in": 13.333,
  "height_in": 7.5,
  "guessed_theme": {"key":"nature","name":"自然质感","confidence":"medium",
                    "primary":"#2D5016","bg":"#FDFCF8"},
  "slides": [
    {
      "index": 0,
      "background": "#FAFAF8",
      "shapes": [
        {"kind":"text","left":0.9,"top":0.6,"width":11.53,"height":1.2,
         "paragraphs":[{"text":"标题","size_pt":34,"color":"#1A1A1A",
                        "bold":true,"align":"left","font":"Microsoft YaHei"}],
         "guessed_role":"title","role_confidence":"high"},
        {"kind":"image","left":7.5,"top":1.7,"width":4.8,"height":4.5},
        {"kind":"table","rows":[["..",".."]],"n_cols":2,"n_rows":3},
        {"kind":"shape","shape_type":"rounded_rect","fill":"#E8DDD0",
         "left":..,"top":..,"width":..,"height":..}
      ],
      "guessed_layout": "bullets",
      "layout_confidence": "high"
    }
  ]
}
```

### 6.3 实现要点

- 遍历 `prs.slides` → `slide.shapes`，按 `shape.shape_type`（`MSO_SHAPE_TYPE`：TEXT_BOX/AUTO_SHAPE/PICTURE/TABLE/PLACEHOLDER/CHART）分流。
- 文本：读 `text_frame.paragraphs` → 每个 run 的 `font.size`/`font.color.rgb`/`font.bold`/`paragraph.alignment`/`font.name`。
- 背景：⚠️ **`slide.background.fill` 有副作用**（`pptx/slide.py:471-498`：访问即把背景设为 NoFill、切断继承）。inspect **只读不写**时，不能直接访问 `.fill`，要先检查 `_element.cSld` 下有无 `p:bg/p:bgPr/a:solidFill`：
  ```python
  # 安全读取背景色，不破坏继承
  cSld = slide._element.cSld
  bg = cSld.find(qn('p:bg'))
  bg_color = None
  if bg is not None:
      sf = bg.find('.//' + qn('a:solidFill') + '/' + qn('a:srgbClr'))
      if sf is not None:
          bg_color = '#' + sf.get('val').upper()
  # bg is None → 背景继承自 master/layout，inspect 里标 "inherited"
  ```
- 主题色猜测：统计所有出现色值频次，取 top-N 对照 9 套主题的 bg/primary 找最接近的，给 `confidence`。
- shape 遍历注意：`slide.shapes` 返回 `SlideShapes`（`shapes/shapetree.py`），可迭代；`shape.shape_type` 是 `MSO_SHAPE_TYPE` 枚举。picture 读 `shape.left/top/width/height`（EMU，转英寸 `/914400`）；table 读 `shape.rows`/`shape.columns`。

### 6.4 版式猜测（全面猜 + 置信度）

14 种版式都尝试识别，`layout_confidence` 取 `high`/`medium`/`low`/`unknown`：

| 版式 | 识别信号 | 置信度 |
|---|---|---|
| cover | 首页 + 全屏图或单居中元素 | high |
| end | 末页 + 深底 + 居中单字 | high |
| toc | 第2页 + 列表编号 + 大标题 | medium |
| table | 含 table shape | high |
| image_grid | 3+ 图片网格布局 | high |
| dashboard | metrics 文本框成行 + chart/柱状 | medium |
| timeline | 竖轴 + 左右交替节点 | medium |
| comparison | 左右两列文本 + 中分隔线 | medium |
| quote | 单大文本 + 装饰引号字符 | medium |
| process | 横向卡片成行 + 箭头 | medium |
| section | 单大编号/标题占比高 | low |
| full_image | 全屏图 + 底部文字 | medium |
| text_image | 左文右图二分 | medium |
| bullets | 多文本框竖排 + 序号圆 | medium |

- 多特征交叉命中 = high，单特征 = medium，弱信号 = low，无信号 = unknown。
- 猜测规则封装在内部 `_guess_layout(slide_dump) -> tuple[str, str]`，便于单测。

### 6.5 字号陷阱对策（缺口 3）

`layout_bullets` 顶部有 **60pt 页码水印**、`layout_quote` 有 **200pt 装饰引号**，都比 34pt 标题大。按"字号最大→标题"猜 role 会全猜错。对策：

- role 猜测**不只看字号**，综合判断：
  - **位置**：标题通常在顶部 1/4 区（top < height*0.3）。
  - **字号 + 位置**：顶部区 + 较大字号 = title；顶部区 + 小字号 = kicker。
  - **装饰排除**：含单个特殊字符（引号/数字且字号 > 50pt 且非顶部居中标题位）的文本框标 `guessed_role="decoration"`，不参与 title 竞争。
  - **内容**：纯数字 + 大字号 + 角落位置 = page_number 水印，标 `decoration`。
- 所有 role 字段带 `guessed_` 前缀 + `role_confidence`，agent 据此判断可信度。

### 6.6 有损边界（明确告知）

外部 PPT 的装饰元素、自定义版式、动画、母版继承无法还原为 Section dict。inspect 产物**仅供 agent 阅读**，不直接喂给 `regenerate`。要改外部 PPT 只能走 L3 定点编辑。

---

## 7. L3 — 定点编辑（外部 PPT，精确改文字/图/色）

**文件**：`scripts/ppt_edit.py`

### 7.1 接口统一约定（缺口 7）

所有编辑函数**统一接受 `pptx_path: str`**（不接受 `prs` 对象），内部 `Presentation(path)` 打开 → 修改 → `save`。避免"对象 vs 路径"双模式导致的脏引用和未保存丢失。需要链式多步编辑时，agent 串行调用即可（每步落盘）。

### 7.2 接口签名

```python
def edit_text(pptx_path: str, slide_index: int, find: str, replace: str) -> int:
    """在指定页文本里做字符串替换，返回替换次数。
    优先按 run 内匹配；跨 run 时见 7.3 对策。"""

def edit_text_by_role(pptx_path: str, slide_index: int, role: str, new_text: str) -> int:
    """配合 inspect 的 guessed_role，改该页"标题/正文"。role 取 inspect 产出。"""

def swap_image(pptx_path: str, slide_index: int, image_index: int, new_path: str) -> bool:
    """替换某页第 N 张图片，保留原 bbox。成功返回 True。"""

def recolor(pptx_path: str, old_hex: str, new_hex: str) -> int:
    """全局色值替换（run.font.color + shape.fill + slide.background），返回替换数。"""

def swap_theme(pptx_path: str, new_theme_key: str) -> dict:
    """自动猜旧主题 + 套新主题色映射。见 7.4。
    返回 {"from_theme":str, "to_theme":str, "replaced":int, "guessed":bool}。"""

def edit_title(pptx_path: str, slide_index: int, new_title: str) -> int: ...
def edit_kicker(pptx_path: str, slide_index: int, new_kicker: str) -> int: ...
```

### 7.3 跨 run 格式丢失对策（缺口 4）

跨 run 的匹配（如 `**重要**的文字` 里"重要的文字"跨了粗体 run 和正常 run）若直接合并再 split，会抹平局部格式。对策：

- **优先 run 内匹配**：遍历每个 run，若 `find` 整串在单 run 内，直接替换，保留该 run 格式。
- **跨 run 时保守**：若 `find` 跨多个 run，替换为：把匹配段的第一个 run 文本改为 `replace`，后续匹配 run 文本清空（保留 run 结构但置空），**不合并 run**。这样替换文本继承第一个 run 的格式，原后续 run 格式虽不生效但不破坏文档。
- 返回时标注是否发生了跨 run 替换，让 agent 知道格式可能有损。

### 7.4 swap_theme 猜错保护（缺口 5）

外部 PPT 若不属于 9 套主题，自动猜旧主题会猜错，色值映射全乱却照样硬套。对策：

- `swap_theme` 内部先调 `inspect_ppt` 拿 `guessed_theme`。
- **置信度门槛**：若 `guessed_theme.confidence` 为 `low`/`unknown`，**不自动套用**，改为返回错误提示：
  ```
  {"error": "无法可靠识别原主题(confidence=low)，请先用 inspect_ppt 查看，
   或用 recolor(old_hex, new_hex) 手动逐色替换。"}
  ```
- confidence 为 medium/high 时才执行映射，且返回结果里带 `guessed=True` 和具体 `from_theme`，agent 可核对。
- 映射表 = 旧主题 9 色（bg/bg_alt/primary/secondary/accent/text/text_muted/white/dark）→ 新主题对应 9 色，全局替换。

### 7.5 swap_image 实现细节

- 用 `slide.shapes` 找第 `image_index` 个 `MSO_SHAPE_TYPE.PICTURE`。
- 记录原 `left/top/width/height`，删除旧 picture，`add_picture(new_path, ...)` 用原 bbox。
- 若新图比例不同，按 `cover` 裁剪适配（复用 pptx_helper 的 `_fit_image_in_box` 思路）。

---

## 8. L4 — 页级操作（增删移换，两条路线通用）

**文件**：`scripts/ppt_pages.py`

### 8.1 接口签名

```python
def insert_slide(pptx_path: str, index: int, section_dict: dict,
                 layout: str | None = None, theme_key: str | None = None,
                 auto_search_images: bool = True) -> int:
    """在指定位置插入一页，用本 skill 版式渲染。返回插入后页数。
    layout 为空则 choose_layout。见 8.3 搜图对策。"""

def delete_slide(pptx_path: str, index: int) -> int:
    """从 _sldIdLst 移除 + 删 relationship。返回删除后页数。"""

def move_slide(pptx_path: str, from_index: int, to_index: int) -> None:
    """重排 _sldIdLst 子元素顺序。"""

def replace_layout(pptx_path: str, index: int, new_layout: str,
                   section_dict: dict | None = None) -> None:
    """删旧页 + 用同内容按新版式重画。
    section_dict 为空时：有 manifest 从 manifest 取，无则先 inspect 提取（尽力）。"""
```

### 8.2 _sldIdLst 操作（已据源码确认）

python-pptx 1.0.2 的 `Slides`（`pptx/slide.py:238`）**只有 `add_slide`/`get`/`index`，没有 delete/move**。增删移需操作 `prs.slides._sldIdLst`（类型 `CT_SlideIdList`，`pptx/oxml/presentation.py:51`），它有：

- `sldId_lst: list[CT_SlideId]` —— 标准 list，可 `append`/`insert`/`remove`/切片赋值。每个 `CT_SlideId` 有 `.id`（≥256 int）和 `.rId`（str）属性。
- `add_sldId(rId) -> CT_SldId` —— 创建新 sldId 元素，自动分配 `_next_id`。

**官方同款范式**：`SlideLayouts.remove`（`pptx/slide.py:390-411`）= 从 list 移除元素 + `part.drop_rel(rId)`。删 slide 照搬：

```python
# delete_slide(i)：删第 i 页
sld_id_lst = prs.slides._sldIdLst.sldId_lst
sldId = sld_id_lst[i]
sld_id_lst.remove(sldId)
# 用包级 drop_rel（无条件删），而非 XmlPart.drop_rel（有 ref_count<2 保护可能删不掉）
prs.part.package.drop_rel(sldId.rId)
```

**drop_rel 选用要点**（源码核验）：
- `XmlPart.drop_rel`（`opc/package.py:409`）有保护：`if self._rel_ref_count(rId) < 2: pop`。slide 的 rId 通常只被 sldId 引用一次，ref_count=1，能删；但保险起见用包级。
- `OpcPackage.drop_rel`（`opc/package.py:84`）无条件 `pop`，最稳。删 slide 走 `prs.part.package.drop_rel(rId)`。

**move_slide(from,to)**：纯重排，不碰 relationship，只动 list 顺序：
```python
lst = prs.slides._sldIdLst.sldId_lst
elem = lst.pop(from_index)
lst.insert(to_index, elem)
```

**insert_slide(idx, ...)**：`add_slide` 只追加末尾（`Slides.add_slide` 调 `_sldIdLst.add_sldId(rId)` 末尾追加），插入到指定位置需追加后移动元素：
```python
slide = prs.slides.add_slide(blank_layout)   # 追加到末尾
# 渲染内容到 slide（调 layout_fn）
lst = prs.slides._sldIdLst.sldId_lst
lst.insert(index, lst.pop())   # 把末尾元素移到 index 位置
```

封装在内部工具 `_reorder_sld_id(prs, old_pos, new_pos)` / `_remove_sld_id(prs, index)` / `_insert_sld_id_at(prs, index)`。实现后用 render 验证页序正确（这是 python-pptx 不直接支持的操作，最易出错，必须每步 render 比对）。

### 8.3 insert_slide 搜图对策（缺口 8）

插一个声明了 `image_query` 却没 `images` 的 section，会生成空页（`choose_layout` 因 `n_images=0` 选纯文字版式，或 image_grid 页空白）。对策：

- `insert_slide` 默认 `auto_search_images=True`：若 `section_dict` 有 `image_query` 但无 `images`，调 `_auto_search_images` 补搜。
- 若搜图失败（网络/key 问题），降级：把该 section 的 `layout` 强制设为不依赖图的版式（bullets/section），并在返回结果里带 warning。
- 对外部 PPT（无 manifest）：`theme_key` 为空时，先 `inspect_ppt` 猜主题，让新插入页与原 PPT 风格尽量一致；猜不出用默认 `editorial`。

### 8.4 replace_layout 的内容提取

- **有 manifest**：从 manifest 取该 index 的 section，改 `layout` 字段后 regenerate（仅该页）。
- **无 manifest**：先 `inspect_ppt` 该页，尽力提取 title/bullets/metrics 等填进 section dict，再按新 layout 重画。提取不全的字段留空，agent 可后续补。

---

## 9. 横切关注点

### 9.1 备份 / undo（缺口 6）

"修改"场景改坏没有 Ctrl+Z 是硬伤。对策：

- **所有写操作的函数**（edit_section / regenerate / insert_slide / delete_slide / move_slide / replace_layout / edit_text / swap_image / recolor / swap_theme）在 `save` 前，把原文件复制为 `原名.bak.pptx`（同目录）。
- 提供 `restore_backup(pptx_path: str) -> bool`：把 `.bak.pptx` 还原回原文件。
- 备份只保留最近一次（覆盖式），避免堆积。
- agent 工作流里改完先 render 验收，不满意 `restore_backup` 回退。

### 9.2 编码（踩过的坑）

Windows 控制台 GBK 会乱码中文输出，误判验证结果。所有脚本运行用 `PYTHONIOENCODING=utf-8`。新文件的 print 输出统一 UTF-8，关键结果用结构化返回值（dict）而非纯 print，减少对终端编码的依赖。

### 9.3 错误处理约定

- 文件不存在、index 越界、图路径无效等：抛 `ValueError` 带清晰消息，不静默吞。
- 搜图失败、主题猜不出：不抛异常，返回带 `warning`/`error` 字段的 dict，让 agent 决策。
- manifest XML 读写失败：catch 后走 json 兜底，不中断主流程。

---

## 10. 数据结构汇总

### 10.1 manifest（L1）

见 5.2。

### 10.2 inspect dump（L2）

见 6.2。

### 10.3 函数签名总表

| 层 | 文件 | 函数 | 签名摘要 |
|---|---|---|---|
| L1 | ppt_project | `_save_manifest` | `(prs, project) -> None` |
| L1 | ppt_project | `load_project` | `(pptx_path) -> dict \| None` |
| L1 | ppt_project | `regenerate` | `(project, output_path?) -> str` |
| L1 | ppt_project | `edit_section` | `(pptx_path, index, changes, output_path?) -> str` |
| L1 | ppt_project | `restore_backup` | `(pptx_path) -> bool` |
| L2 | ppt_inspect | `inspect_ppt` | `(pptx_path) -> dict` |
| L3 | ppt_edit | `edit_text` | `(pptx_path, slide_index, find, replace) -> int` |
| L3 | ppt_edit | `edit_text_by_role` | `(pptx_path, slide_index, role, new_text) -> int` |
| L3 | ppt_edit | `swap_image` | `(pptx_path, slide_index, image_index, new_path) -> bool` |
| L3 | ppt_edit | `recolor` | `(pptx_path, old_hex, new_hex) -> int` |
| L3 | ppt_edit | `swap_theme` | `(pptx_path, new_theme_key) -> dict` |
| L3 | ppt_edit | `edit_title` / `edit_kicker` | `(pptx_path, slide_index, new) -> int` |
| L4 | ppt_pages | `insert_slide` | `(pptx_path, index, section_dict, layout?, theme_key?, auto_search?) -> int` |
| L4 | ppt_pages | `delete_slide` | `(pptx_path, index) -> int` |
| L4 | ppt_pages | `move_slide` | `(pptx_path, from_index, to_index) -> None` |
| L4 | ppt_pages | `replace_layout` | `(pptx_path, index, new_layout, section_dict?) -> None` |

---

## 11. 实施顺序与风险

| 步骤 | 内容 | 风险/验证 |
|---|---|---|
| 1 | L1 manifest 往返（json 路径）+ `_build_slides` 抽取 + regenerate | json 零风险；图片重定位需测跨目录 |
| 2 | L4 insert/delete/move/replace_layout | `_sldIdLst` 操作源码已确认（8.2），每步必 render 比对页序 |
| 3 | L2 inspect_ppt + 版式/主题猜测 | 背景读取避开 `.fill` 副作用（6.3）；字号陷阱对策用本 skill PPT 反测 |
| 4 | L3 定点编辑 + swap_theme | 跨 run 替换格式需用含粗体文本测；swap_theme 猜错保护（7.4） |
| 5 | SKILL.md 新增"修改 PPT 工作流"章节 + 5 个典型示例 | — |
| 6（可选） | 嵌入 pptx 的 manifest Part（增强项） | Content-Type override 难点，评估收益后再做 |

每步独立可验，建议一步一验再进下一步。所有 `_sldIdLst`/`drop_rel` 操作已据 python-pptx 1.0.2 源码确认 API 路径，不再是"参照社区方案"。

---

## 12. 验证方案

- **L1**：`auto_generate_ppt` 生成 → `load_project` 读回 → `edit_section` 改第 2 页 bullets → `regenerate` → render 对比，确认只有第 2 页变、其余页不变。再测跨目录拷贝（图片重定位）。
- **L4**：9 页 → `delete_slide(3)` → 8 页且第 4 页被删；`insert_slide(2,{...})` → 9 页且新页在第 2 位；`move_slide` 顺序正确。每步 render 验收页序。
- **L2**：拿外部 pptx → `inspect_ppt` → dump 字段完整；cover/bullets/table 几种明显版式猜对；用本 skill 生成的 PPT 反测，确认 60pt 水印/200pt 引号不被误判为标题。
- **L3**：`edit_text` 替换生效不破坏排版（含粗体跨 run 场景）；`swap_theme("editorial"→"neon_tech")` 全换色无残留；`swap_theme` 对非 9 套主题的外部 PPT 返回 error 而非乱改；`recolor` 计数正确。
- **备份**：任意编辑后 `.bak.pptx` 存在；`restore_backup` 能还原。
- 全程 `PYTHONIOENCODING=utf-8` 跑，关键结果用 dict 返回值核对，不只看 print。

---

## 13. SKILL.md 待补章节（草稿）

在"## 数据可视化"之后、"## 截图验收标准"之前插入"## 修改 PPT 工作流"一节，含：路线 A/B 决策树、改完复验要求、5 个典型示例（改某页文字 / 换主题 / 加一页 / 删一页 / 把某页改成时间线）。实现完成后补写。
