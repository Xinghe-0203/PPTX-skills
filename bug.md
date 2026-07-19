# PPTX Skill 代码问题报告

> 生成日期：2026-07-19
> 验证方式：20个子代理交叉验证
> 原始报告问题数：244 → 验证确认：133

---

## 问题总览

| 严重程度 | 数量 |
|---------|------|
| 🔴 严重 | 18 |
| 🟡 中等 | 53 |
| 🟢 轻微 | 62 |
| **合计** | **133** |

---

## 🔴 严重问题

### 1. `content_model.py:170-180` — `_ensure_unique` 逻辑反了

**类型**：Bug / 逻辑错误

```python
def _ensure_unique(self, candidate: str) -> str:
    if candidate in self._seen:
        return candidate        # 已存在 → 直接返回重复ID
    self._seen.add(candidate)
    return candidate
```

当 `candidate` 已在 `_seen` 中时直接返回相同值，从不添加后缀。真正的唯一性保证完全失效，会导致幻灯片/元素 ID 重复。

---

### 2. `layout_engine.py:161,449` — `.replace(".", ".")` 无操作死代码

**类型**：Bug

```python
_resolve_token(value["$ref"].lstrip("$").replace(".", "."), tokens)
```

`replace(".", ".")` 将 `.` 替换为 `.`，完全无意义。疑似未完成的笔误（可能本意是其他字符替换）。

---

### 3. `layout_engine.py:97-100` — zone 值非 dict 时类型错误

**类型**：Bug / 类型错误

```python
zones = {k: ZoneSpec(**v) if isinstance(v, dict) else v for k, v in ...}
```

非 dict 的 zone 值直接赋给 `ZoneSpec` 类型字段，下游访问 `.kind` 等属性必然 `AttributeError`。

---

### 4. `layout_engine.py:325` — `_estimate_text_height` 除零风险

**类型**：Bug / 边界条件

```python
math.ceil(len(text) * size * 0.6 / bbox.width)
```

当 `bbox.width == 0` 时除零崩溃。

---

### 5. `layout_engine.py:400-405` — 重叠区域密度计算错误

**类型**：Bug / 逻辑错误

density 计算中重叠区域被重复计入总面积，导致 `density > 1.0` 时（元素重叠）反而减轻惩罚。

---

### 6. `layout_engine.py:424-428` — `plan_slide_candidates` 返回空列表

**类型**：Bug / 边界条件

所有 recipe 都有 blocker 时返回空列表，依赖调用方自行防御。

---

### 7. `text_metrics.py:136-145` — `_load_font` 永久缓存失败字体

**类型**：Bug

字体加载失败（文件损坏）后回退到 `ImageFont.load_default()` 并缓存，后续永不重试加载真正字体。

---

### 8. `text_metrics.py:274-287` — CJK 后处理单字符行丢失内容

**类型**：Bug / 数据丢失

当 `prev` 只有单字符时，`fixed[-1] = prev[:-1]` 清空后 `fixed = [ln for ln in fixed if ln]` 过滤掉空行，字符永久丢失。

---

### 9. `text_metrics.py:175-189` — 高度计算恒等变换

**类型**：Bug

```python
height = (ascent + descent) / 72 * 72  # 乘除抵消，疑似未完成 DPI 转换
```

---

### 10. `text_metrics.py:382-390` — `TextMeasurer.measure` 接口不完整

**类型**：Bug

未传递 `min_font_size_pt` 和 `height_pt` 参数给 `measure_runs`，通过此接口无法利用高度约束。

---

### 11. `image_crop.py:56-83` — `crop_contain` / `crop_cover` 无除零检查

**类型**：Bug / 边界条件

```python
src_ratio = src_width / src_height  # src_height == 0 → ZeroDivisionError
```

---

### 12. `manifest.py:189-210` — `_presentation_rels_xml` 无异常处理

**类型**：Bug / 异常处理缺失

```python
root = ET.fromstring(existing)  # 损坏 XML 直接崩溃
```

同一文件内 `_read_embedded_manifest` 有 `except ET.ParseError`，此处遗漏。

---

### 13. `manifest.py:219` — rels 文件不存在时 KeyError

**类型**：Bug / 异常处理缺失

```python
members[presentation_rels] = _presentation_rels_xml(members[presentation_rels])
```

非标准 PPTX 文件缺少内部 rels 文件时直接崩溃。

---

### 14. `qa_dataset.py:221-228` — control builder 与 issue builder 完全相同

**类型**：Bug / 测试数据错误

`table_readability` 和 `reference_residual` 的 control = issue（同一函数），无法区分对照组与实验组。

---

### 15. `golden_renders.py:171-182` — `_try_render_png` 用 fitz 打开 PPTX

**类型**：Bug

PyMuPDF (`fitz`) 不支持 PPTX 格式，该函数必然失败，每次浪费一次异常处理。

---

### 16. `reference_ppt.py:332-356` — `clone_slide` 大量使用私有 API

**类型**：可靠性风险

使用 `_spTree`、`_element`、`cSld`、`p:extLst` 等 python-pptx 未公开内部 API，版本升级可能失效。

---

### 17. `pixabay_search.py:200-222` — download_images 重试逻辑缺陷

**类型**：Bug / 数据损坏

第一次下载写入部分数据后失败，文件留在磁盘，重跑时被误认为"已下载"而跳过。

---

### 18. `pptx_helper.py:757` — 柱状图负值处理有误

**类型**：Bug

```python
val = int(chg.replace("+", "").replace("%", "").replace("-", ""))  # 先移除负号
chart_data_vals.append(abs(val))                                    # 再 abs，双重丢失符号
```

---

## 🟡 中等问题

### 异常处理不足

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 19 | `reference_adapter.py` | 29-42 | `except Exception` 过宽，应只捕获 `ImportError` |
| 20 | `reference_adapter.py` | 350 | `except Exception: return drift` 静默吞掉所有 PPTX 加载异常 |
| 21 | `api.py` | 190 | `except Exception` 捕获过宽，无日志记录 |
| 22 | `api.py` | 183-192 | try 块范围过大（5 个操作共用） |
| 23 | `capability.py` | 59 | `except Exception as exc` 捕获过宽 |
| 24 | `semantic_qa.py` | 159 | `except Exception` + `# pragma: no cover`，分支未被测试 |
| 25 | `render_qa.py` | 62-63 | `_load_image` 返回 `Image.open()` 但未关闭 |
| 26 | `preview_renderer.py` | 176-186 | LibreOffice 超时子进程残留 |
| 27 | `preview_renderer.py` | 265-278 | COM 清理异常被静默吞掉 |
| 28 | `preview_renderer.py` | 94-125 | PyMuPDF 部分成功被丢弃 |
| 29 | `pptx_renderer.py` | 347-363 | 单节点失败整个 PPTX 生成中断 |

### 边界条件缺失

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 30 | `visual_rebuild.py` | 232-234 | `ref_boxes` 取 `s["x"]` 等键无防御访问 |
| 31 | `visual_rebuild.py` | 325-329 | fallback `builtin_recipes(role)[0]` 可能 IndexError |
| 32 | `template_compiler.py` | 49 | `hx()` 未验证颜色字符串长度 |
| 33 | `deck_planner.py` | 383-402 | canvas 尺寸小于 96 时 width 变负 |
| 34 | `golden_renders.py` | 98-145 | `_role_slide_spec` 为所有 zone 创建 element（含 recipe 未使用的） |
| 35 | `golden_renders.py` | 160-161 | `plan.canvas.width_pt` 无 None 防御 |
| 36 | `repair_engine.py` | 172 | `low_contrast_warnings` 缺少 `slide_index` 边界检查 |
| 37 | `repair_engine.py` | 188 | `empty_placeholders` 缺少 `slide_index` 边界检查 |
| 38 | `repair_engine.py` | 233-241 | `action.to_value` 为 None 时静默失败 |
| 39 | `semantic_qa.py` | 296-302 | `a.geometry.bbox` 无 None 检查 |
| 40 | `pptx_helper.py` | 393-394 | `_fit_image_in_box` 返回 None 时调用方未检查 |
| 41 | `pptx_helper.py` | 1528 | `shape.has_text_frame` 对特殊 shape 抛出 `AttributeError` |
| 42 | `ppt_project.py` | 96 | `members[presentation_rels]` 无防御 KeyError |
| 43 | `ppt_project.py` | 156 | `_relocate_images` 相对路径依赖当前工作目录 |

### 逻辑错误/不完整

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 44 | `generation_pipeline.py` | 163 | `i.severity.value != "blocker"` 应使用枚举比较 |
| 45 | `repair_engine.py` | 49-58 | `_view_issues` if/elif 分支逻辑完全冗余 |
| 46 | `repair_engine.py` | 113-116 | slide_index 钳制后检查 `< 0` 为死代码 |
| 47 | `repair_engine.py` | 135-141 | 字体缩减 current/min_size 默认值 16 和 9 不一致 |
| 48 | `semantic_qa.py` | 143-148 | `_check_text_overflow` 中 resolved_style 值无类型验证 |
| 49 | `semantic_qa.py` | 244-252 | `target_ratio` 除零保护顺序混乱（先除后查） |
| 50 | `render_qa.py` | 207 | `annotated` 用固定 40% 红色覆盖而非差异 mask 叠加 |
| 51 | `render_qa.py` | 304 | `elif` 替代 `if` 导致可能漏检不同维度回归 |
| 52 | `content_adapter.py` | 101-104 | `section.get("left") or {}` 抹掉 falsy 值 |
| 53 | `pptx_renderer.py` | 117 | 对齐映射注释与实际枚举值不一致 |
| 54 | `pagination.py` | 151-156 | `_split_text_element` 重建时丢失 level=0 行 `- ` 前缀 |
| 55 | `deck_planner.py` | 113-130 | 不在 `preferred_sequence` 中的 role 被 0.5 倍惩罚 |
| 56 | `golden_renders.py` | 330-333 | `load_golden_index` 无文件不存在异常处理 |
| 57 | `reference_ppt.py` | 359-365 | `_remove_slide` 使用 `prs.slides._sldIdLst` 等私有 API |
| 58 | `ppt_pages.py` | 57-61 | `_move_id` 操作 `prs.slides._sldIdLst` 私有 API |
| 59 | `ppt_project.py` | 202-203 | `zip` 长度不匹配时多余 sections 无 layout |

### 性能问题

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 60 | `text_metrics.py` | 61-84 | `_index_system_fonts` 全局变量线程不安全 |
| 61 | `text_metrics.py` | 72-84 | `rglob("*")` 递归扫描所有文件再过滤 |
| 62 | `text_metrics.py` | 81-82 | token 按 `_`、`-`、`+` 分割过于激进 |
| 63 | `text_metrics.py` | 98 | `resolve_font` 按文件名长度排序不可靠 |
| 64 | `text_metrics.py` | 175-201 | 每次测量都创建新 Image/ImageDraw |
| 65 | `text_metrics.py` | 192-200 | 逐字符宽度计算无缓存 |
| 66 | `text_metrics.py` | 413-414 | `fit_text_to_height` 二分搜索初始值无法区分无解 |
| 67 | `manifest.py` | 215-216 | `source.infolist()` 被调用两次 |
| 68 | `golden_renders.py` | 148-153 | `_recipe_by_id` 线性搜索无缓存 |
| 69 | `pptx_renderer.py` | 160-163 | PIL Image 对象未关闭 |

### 数据完整性

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 70 | `manifest.py` | 148-156 | `_deserialize_dataclass` 跳过未知字段，round-trip 数据丢失 |
| 71 | `manifest.py` | 134-135 | tuple 序列化为 list，反序列化无法还原 |

---

## 🟢 轻微问题

### 死代码/未使用导入

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 72 | `generation_pipeline.py` | 15-16 | `builtin_recipes`, `plan_slide_candidates`, `solved_geometry_to_layout_plan` 顶层导入未使用 |
| 73 | `reference_adapter.py` | 24 | 四个导入（含 `_builtin_tokens`）未使用 |
| 74 | `visual_rebuild.py` | 34 | `plan_slide_candidates` 导入未使用 |
| 75 | `visual_rebuild.py` | 15 | `asdict` 导入未使用 |
| 76 | `visual_rebuild.py` | 19-27 | `BBox`, `ContentSpec`, `GeometrySpec` 导入未使用 |
| 77 | `visual_rebuild.py` | 319-321 | `preferred_variant=None` 导致参数永远不生效 |

### 类型注解缺失

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 78 | `visual_qa.py` | 64 | `"RenderResult \| None"` 前向引用中 `RenderResult` 未导入 |
| 79 | `image_crop.py` | 116 | `_load_and_resize` 缺少返回类型注解 |
| 80 | `image_crop.py` | 129 | `_saliency_map` 缺少参数和返回类型注解 |
| 81 | `design_schema.py` | 182-271 | 函数声明返回 `list[str]` 但始终返回 `[]` |
| 82 | `render_slides.py` | 16-17 | 函数签名缺少类型注解 |
| 83 | `pptx_helper.py` | 1065 | `choose_layout` 签名声明 `section: dict` 但接受 Section dataclass |

### 代码风格/可读性

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 84 | `semantic_qa.py` | 64 | `# type: ignore[return-value]` 抑制类型检查，应显式构造 |
| 85 | `repair_engine.py` | 182 | 4 层嵌套 `.get()` 链可读性差 |
| 86 | `design_schema.py` | 294-307 | `brand_colors` 未做 `strip()` |
| 87 | `pptx_renderer.py` | 365-379 | 各 `_add_*_node` 硬编码 `slide_index=0` 然后重建覆盖 |
| 88 | `pptx_helper.py` | 303-305 | `_luminance` 函数无类型防御 |
| 89 | `pptx_helper.py` | 586 | `layout_bullets` 中 `gap=0.18` 硬编码 |
| 90 | `pptx_helper.py` | 624 | `layout_text_image` 中 `space_before=8` 硬编码 |
| 91 | `pptx_helper.py` | 773 | `* 0.85` 魔术数字无注释 |
| 92 | `pptx_helper.py` | 937 | `__import__('sys')` 内联导入不规范 |
| 93 | `pptx_helper.py` | 1311-1325 | `_load_pixabay_search` 每次调用都重新加载模块 |
| 94 | `pptx_helper.py` | 1284-1288 | `template_profile` 不可序列化时整个丢弃 |
| 95 | `ppt_edit.py` | 174 | `_color_hex` 异常捕获类型不全 |
| 96 | `ppt_edit.py` | 27-28 | `_slide_at` 1-based 索引设计（0 和 1 都指向第一页） |
| 97 | `ppt_edit.py` | 278-279 | 颜色频率差异导致主题匹配偏差 |
| 98 | `ppt_pages.py` | 29 | `_position` 函数名不够描述性 |
| 99 | `ppt_pages.py` | 106-107 | page_number/section_number 覆盖逻辑缺注释 |
| 100 | `ppt_pages.py` | 197 | 空 dict 不触发 fallback（边界条件） |
| 101 | `ppt_project.py` | 144-146 | 两个连续 `return None` 的死代码 |
| 102 | `ppt_project.py` | 228 | index=0 和 1 都映射到 0，API 混淆 |
| 103 | `reference_ppt.py` | 36-52 | `_rgb_hex` 不处理带 `#` 前缀的字符串 |
| 104 | `reference_ppt.py` | 460-479 | `_content_lines` 隐式优先级无文档 |
| 105 | `reference_ppt.py` | 150-157 | `_classify_slide` 判断顺序可能导致误分类 |
| 106 | `reference_ppt.py` | 686-709 | 结尾页"谢谢"硬编码 |
| 107 | `template_engine.py` | 134-135 | 验证只检查 4 个必需颜色，其他键缺失无提示 |
| 108 | `template_engine.py` | 209-212 | 色相只用 SHA256 第一个字节，分布有限 |
| 109 | `template_engine.py` | 253 | dark/light 优先级逻辑缺注释 |
| 110 | `template_engine.py` | 82-93 | glob 中非 JSON 文件导致崩溃 |
| 111 | `template_engine.py` | 66 | 版本号硬编码为 1 |
| 112 | `pixabay_search.py` | 146-147 | HTTPError 用 `from None` 丢失原始异常链 |
| 113 | `pixabay_search.py` | 252 | 函数内重复导入 time |
| 114 | `pixabay_search.py` | 346-347 | 无搜索结果用退出码 1（可能误判为错误） |
| 115 | `layout_variants.py` | 8-25 | 导入 `pptx_helper` 的 `_` 前缀私有函数 |
| 116 | `layout_variants.py` | 246-251 | `replace("-", "")` + `abs()` 双重去除负号 |
| 117 | `layout_variants.py` | 260 | value=0 仍显示 0.08 英寸小条 |
| 118 | `layout_variants.py` | 619-665 | `**` 展开继承模式可能意外传播新 key |

### 文档缺失

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 119 | `content_adapter.py` | 15-40 | `_infer_layout` 优先级顺序未在 docstring 中说明 |
| 120 | `template_compiler.py` | 92 | 错误信息未说明期望的颜色格式 |

### 其他

| # | 文件 | 行号 | 问题 |
|---|------|------|------|
| 121 | `api.py` | 166 | `from pptx_helper import` 导入路径脆弱 |
| 122 | `design_schema.py` | 226-233 | 未验证 layouts 的 role 是否在 REQUIRED_ROLES |
| 123 | `render_qa.py` | 179-181 | `_load_image` 返回 Image.open 但未关闭 |
| 124 | `manifest.py` | 106 | `run_id` 在同一秒内可能重复 |
| 125 | `manifest.py` | 133 | dict 的 key 未被序列化 |
| 126 | `manifest.py` | 224 | `ZIP_DEFLATED` 需要 zlib（某些 Python 构建不可用） |
| 127 | `qa_dataset.py` | 141 | 使用 `__import__` 动态导入而非顶部静态导入 |
| 128 | `qa_dataset.py` | 186 | `shape.fill.solid()` 未设置填充颜色 |
| 129 | `qa_dataset.py` | 232-240 | `_bbox_of_shape` 深层属性链可能失败 |
| 130 | `preview_renderer.py` | 87-91 | 跨驱动器 `os.rename` 失败，应用 `shutil.move` |
| 131 | `preview_renderer.py` | 281-323 | fallback 时丢失详细 attempt 信息 |
| 132 | `deck_planner.py` | 196-263 | `derived = [slide]` 循环暗示多 slide 但实际只有一个 |
| 133 | `deck_planner.py` | 323-338 | infeasible fallback 无 QA 检查 |

---

## 优先修复建议

1. **`content_model.py:170-180`** — `_ensure_unique` 逻辑反转（影响所有 ID 唯一性）
2. **`text_metrics.py:136-145`** — 字体加载失败被永久缓存
3. **`text_metrics.py:274-287`** — CJK 后处理导致字符丢失
4. **`qa_dataset.py:221-228`** — control 与 issue 同函数（QA 数据集无效）
5. **`layout_engine.py:97-100`** — zone 非 dict 类型错误
6. **`layout_engine.py:325`** — 除零风险
7. **`image_crop.py:56-83`** — 除零风险
8. **`manifest.py:189-210`** — XML 解析无异常处理
9. **`manifest.py:219`** — rels 文件 KeyError
10. **`golden_renders.py:171-182`** — fitz 打开 PPTX 必然失败
