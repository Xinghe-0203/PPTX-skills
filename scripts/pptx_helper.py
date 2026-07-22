"""
pptx_helper.py — 智能排版 PPT 生成库

设计目标：消除"模板味"和"AI 味"。
核心机制：
  1. 版式库 (LAYOUTS)：提供 14+ 种差异化版式，每种服务于一种信息结构。
  2. 智能版式选择器 (choose_layout)：根据每个章节的内容特征自动挑选最佳版式，
     而不是千篇一律套用同一个布局。
  3. 配色系统：每套主题色不超过 3 个主色 + 中性色，跨页统一。
  4. 字号层级跳跃：标题 / 副标题 / 正文 / 注释之间有明显落差，不做均匀递减。
  5. 大留白、强对比、数据可视化优先于堆文字。

依赖：python-pptx >= 0.6, Pillow (用于图片裁剪适配)。
"""

from __future__ import annotations

import json
import math
import os
from collections.abc import Callable
from dataclasses import dataclass

# lxml is a transitive dependency via python-pptx; not listed directly in
# pyproject.toml but always available when python-pptx is installed.
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

try:
    from PIL import Image  # noqa: F401  用于图片比例适配
except ImportError:  # Pillow 缺失时降级
    Image = None

# ============================================================
# 画布尺寸（16:9 宽屏，单位：英寸）
# ============================================================
SLIDE_W = 13.333
SLIDE_H = 7.5

# 安全边距：内容不贴边，呼吸感
MARGIN_X = 0.9
MARGIN_TOP = 0.6
MARGIN_BOTTOM = 0.5

# 内容区
CONTENT_W = SLIDE_W - 2 * MARGIN_X            # 11.533
CONTENT_H = SLIDE_H - MARGIN_TOP - MARGIN_BOTTOM  # 6.4

LAYOUT_CONSUMED_FIELDS = {
    "cover": {"title", "subtitle", "kicker", "cover_image"},
    "toc": {"title", "kicker", "toc_items", "bullets"},
    "section": {"title", "subtitle", "kicker", "section_number"},
    "bullets": {"title", "subtitle", "kicker", "bullets"},
    "text_image": {"title", "subtitle", "kicker", "bullets", "images"},
    "full_image": {"title", "kicker", "cover_image"},
    "image_grid": {"title", "kicker", "images"},
    "dashboard": {"title", "kicker", "metrics"},
    "timeline": {"title", "kicker", "events", "bullets"},
    "comparison": {"title", "kicker", "left", "right"},
    "quote": {"title", "kicker", "quote", "source"},
    "process": {"title", "kicker", "steps", "bullets"},
    "table": {"title", "kicker", "table_headers", "table_rows", "bullets"},
    "end": {"title", "subtitle", "kicker"},
}

def _warn_unconsumed(layout_name, ctx):
    consumed = LAYOUT_CONSUMED_FIELDS.get(layout_name)
    if not consumed:
        return
    content_fields = {"bullets", "images", "metrics", "events", "steps", "table_headers", "table_rows", "left", "right", "quote", "source", "cover_image"}
    provided = {k for k in content_fields if ctx.get(k)}
    unconsumed = provided - consumed
    if unconsumed:
        print(f"  [警告] 版式 '{layout_name}' 未消费字段 {unconsumed}，内容将被忽略", file=__import__('sys').stderr)

# ============================================================
# 配色系统：多套主题，每套 ≤3 主色 + 中性色，跨页统一
# ============================================================

def _c(hex_str: str) -> RGBColor:
    """#RRGGBB → RGBColor"""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


THEMES: dict[str, dict] = {
    "editorial": {
        "name": "杂志编辑",
        "bg": _c("FAFAF8"), "bg_alt": _c("F0EFEB"),
        "primary": _c("1A1A1A"), "secondary": _c("8C7B6B"), "accent": _c("E63946"),
        "text": _c("1A1A1A"), "text_muted": _c("7A7A7A"), "white": _c("FFFFFF"),
        "dark": _c("0D0D0D"),
        "on_dark": False,
    },
    "luxury": {
        "name": "奢侈品感",
        "bg": _c("1C1917"), "bg_alt": _c("292524"),
        "primary": _c("D4A574"), "secondary": _c("A8A29E"), "accent": _c("F5F0EB"),
        "text": _c("F5F0EB"), "text_muted": _c("A8A29E"), "white": _c("FFFFFF"),
        "dark": _c("0C0A09"),
        "on_dark": True,
    },
    "nature": {
        "name": "自然质感",
        "bg": _c("FDFCF8"), "bg_alt": _c("EDE8DD"),
        "primary": _c("2D5016"), "secondary": _c("B8860B"), "accent": _c("C45D2C"),
        "text": _c("2C2C2C"), "text_muted": _c("6B6256"), "white": _c("FFFFFF"),
        "dark": _c("1A1A1A"),
        "on_dark": False,
    },
    "neon_tech": {
        "name": "荧光科技",
        "bg": _c("0A0A0A"), "bg_alt": _c("1A1A1A"),
        "primary": _c("00FF87"), "secondary": _c("00D4FF"), "accent": _c("FF3366"),
        "text": _c("E0E0E0"), "text_muted": _c("888888"), "white": _c("FFFFFF"),
        "dark": _c("050505"),
        "on_dark": True,
    },
    "coral_breeze": {
        "name": "珊瑚微风",
        "bg": _c("FFFBF5"), "bg_alt": _c("F5EDE3"),
        "primary": _c("1B3A5C"), "secondary": _c("E07A5F"), "accent": _c("F2CC8F"),
        "text": _c("2B2B2B"), "text_muted": _c("7A7268"), "white": _c("FFFFFF"),
        "dark": _c("1B1B1B"),
        "on_dark": False,
    },
    "mono_red": {
        "name": "极简红",
        "bg": _c("FFFFFF"), "bg_alt": _c("F5F5F5"),
        "primary": _c("111111"), "secondary": _c("666666"), "accent": _c("FF2D2D"),
        "text": _c("111111"), "text_muted": _c("999999"), "white": _c("FFFFFF"),
        "dark": _c("000000"),
        "on_dark": False,
    },
    "terracotta": {
        "name": "赤陶暖沙",
        "bg": _c("FAF6F0"), "bg_alt": _c("E8DDD0"),
        "primary": _c("C45832"), "secondary": _c("5B7B6F"), "accent": _c("E8B86D"),
        "text": _c("2E2420"), "text_muted": _c("8C7E72"), "white": _c("FFFFFF"),
        "dark": _c("1E1614"),
        "on_dark": False,
    },
    "indigo_charcoal": {
        "name": "靛蓝炭灰",
        "bg": _c("F5F5F5"), "bg_alt": _c("E5E5E5"),
        "primary": _c("2E3A59"), "secondary": _c("6B7280"), "accent": _c("F59E0B"),
        "text": _c("1F2937"), "text_muted": _c("9CA3AF"), "white": _c("FFFFFF"),
        "dark": _c("111827"),
        "on_dark": False,
    },
    "forest_luxe": {
        "name": "深林赤金",
        "bg": _c("0F1A0F"), "bg_alt": _c("1A2B1A"),
        "primary": _c("C9A96E"), "secondary": _c("8B9D83"), "accent": _c("D4654A"),
        "text": _c("E8E4D9"), "text_muted": _c("8B9D83"), "white": _c("FFFFFF"),
        "dark": _c("050A05"),
        "on_dark": True,
    },
    "tech_blue": {
        "name": "科技蓝",
        "bg": _c("0A1628"), "bg_alt": _c("122240"),
        "primary": _c("00A3E0"), "secondary": _c("4A6FA5"), "accent": _c("00D4FF"),
        "text": _c("E8EDF3"), "text_muted": _c("7A8FA6"), "white": _c("FFFFFF"),
        "dark": _c("050B14"),
        "on_dark": True,
    },
    "warm_earth": {
        "name": "暖土",
        "bg": _c("F5F0E8"), "bg_alt": _c("E8DFD0"),
        "primary": _c("C45B28"), "secondary": _c("8B7355"), "accent": _c("7A8B6F"),
        "text": _c("2E2420"), "text_muted": _c("8C7E72"), "white": _c("FFFFFF"),
        "dark": _c("1E1614"),
        "on_dark": False,
    },
    "minimal_mono": {
        "name": "极简单色",
        "bg": _c("FFFFFF"), "bg_alt": _c("F5F5F5"),
        "primary": _c("1A1A1A"), "secondary": _c("555555"), "accent": _c("888888"),
        "text": _c("1A1A1A"), "text_muted": _c("999999"), "white": _c("FFFFFF"),
        "dark": _c("000000"),
        "on_dark": False,
    },
    "ocean_depth": {
        "name": "深海",
        "bg": _c("0B3D4C"), "bg_alt": _c("0F4F62"),
        "primary": _c("4ECDC4"), "secondary": _c("2E8B8B"), "accent": _c("FF6B6B"),
        "text": _c("E0F0EF"), "text_muted": _c("8BB8B5"), "white": _c("FFFFFF"),
        "dark": _c("062830"),
        "on_dark": True,
    },
    "sunset_glow": {
        "name": "日落暖光",
        "bg": _c("FFF8F0"), "bg_alt": _c("F5EDE0"),
        "primary": _c("E85D26"), "secondary": _c("B87333"), "accent": _c("F4A460"),
        "text": _c("2E2420"), "text_muted": _c("8C7E72"), "white": _c("FFFFFF"),
        "dark": _c("1E1614"),
        "on_dark": False,
    },
    "arctic_frost": {
        "name": "极地霜",
        "bg": _c("F0F4F8"), "bg_alt": _c("E0E8F0"),
        "primary": _c("4A7C9B"), "secondary": _c("6B8FA8"), "accent": _c("B0C4DE"),
        "text": _c("1F2937"), "text_muted": _c("7A8FA6"), "white": _c("FFFFFF"),
        "dark": _c("0F1A28"),
        "on_dark": False,
    },
    "cherry_blossom": {
        "name": "樱花",
        "bg": _c("FFF0F5"), "bg_alt": _c("F5E0EA"),
        "primary": _c("C4627A"), "secondary": _c("9B6B7A"), "accent": _c("E8A0BF"),
        "text": _c("2E2028"), "text_muted": _c("8C7280"), "white": _c("FFFFFF"),
        "dark": _c("1E1018"),
        "on_dark": False,
    },
    "cyber_punk": {
        "name": "赛博朋克",
        "bg": _c("0D0221"), "bg_alt": _c("1A0A3A"),
        "primary": _c("FF2A6D"), "secondary": _c("7B2D8E"), "accent": _c("05D9E8"),
        "text": _c("E8E0F0"), "text_muted": _c("8B7AA0"), "white": _c("FFFFFF"),
        "dark": _c("060114"),
        "on_dark": True,
    },
    "forest_canopy": {
        "name": "森林冠层",
        "bg": _c("1B3A2D"), "bg_alt": _c("244D3A"),
        "primary": _c("7BC950"), "secondary": _c("4A8B3C"), "accent": _c("D4A017"),
        "text": _c("E0F0E0"), "text_muted": _c("8BB88B"), "white": _c("FFFFFF"),
        "dark": _c("0D2820"),
        "on_dark": True,
    },
    "royal_purple": {
        "name": "皇家紫",
        "bg": _c("1A0A2E"), "bg_alt": _c("2A1A4E"),
        "primary": _c("7B2D8E"), "secondary": _c("5A3D7A"), "accent": _c("D4AF37"),
        "text": _c("E8E0F0"), "text_muted": _c("8B7AA0"), "white": _c("FFFFFF"),
        "dark": _c("0A0418"),
        "on_dark": True,
    },
    "desert_sand": {
        "name": "沙漠沙",
        "bg": _c("F4E4C1"), "bg_alt": _c("E8D4A0"),
        "primary": _c("A0522D"), "secondary": _c("8B7355"), "accent": _c("40B5A0"),
        "text": _c("2E2420"), "text_muted": _c("8C7E72"), "white": _c("FFFFFF"),
        "dark": _c("1E1614"),
        "on_dark": False,
    },
}


FONTS: dict[str, dict[str, str]] = {
    "editorial": {"heading": "Georgia", "body": "Calibri"},
    "luxury": {"heading": "Playfair Display", "body": "Lato"},
    "nature": {"heading": "Merriweather", "body": "Source Sans Pro"},
    "neon_tech": {"heading": "Orbitron", "body": "Roboto"},
    "coral_breeze": {"heading": "Nunito", "body": "Open Sans"},
    "mono_red": {"heading": "Helvetica Neue", "body": "Helvetica Neue"},
    "terracotta": {"heading": "Crimson Text", "body": "Lato"},
    "indigo_charcoal": {"heading": "IBM Plex Sans", "body": "IBM Plex Sans"},
    "forest_luxe": {"heading": "Cormorant Garamond", "body": "Montserrat"},
    "tech_blue": {"heading": "Rajdhani", "body": "Roboto"},
    "warm_earth": {"heading": "Lora", "body": "Source Sans Pro"},
    "minimal_mono": {"heading": "Helvetica Neue", "body": "Helvetica Neue"},
    "ocean_depth": {"heading": "Montserrat", "body": "Open Sans"},
    "sunset_glow": {"heading": "Poppins", "body": "Nunito"},
    "arctic_frost": {"heading": "IBM Plex Sans", "body": "IBM Plex Sans"},
    "cherry_blossom": {"heading": "Quicksand", "body": "Nunito"},
    "cyber_punk": {"heading": "Orbitron", "body": "Share Tech Mono"},
    "forest_canopy": {"heading": "Cormorant Garamond", "body": "Source Sans Pro"},
    "royal_purple": {"heading": "Cinzel", "body": "Lato"},
    "desert_sand": {"heading": "Lora", "body": "Open Sans"},
}

CJK_FONTS: dict[str, str] = {
    "heading": "Microsoft YaHei",
    "body": "Microsoft YaHei",
    "mono": "Consolas",
}


_THEME_COLOR_KEYS = (
    "bg", "bg_alt", "primary", "secondary", "accent",
    "text", "text_muted", "white", "dark",
)


def theme_from_profile(profile: dict) -> dict:
    """Convert a JSON-friendly template profile into a runtime theme."""
    source = profile.get("theme", profile)
    base = THEMES["editorial"]
    theme = {**base}
    theme["name"] = profile.get("name", source.get("name", "自定义模板"))
    for key in _THEME_COLOR_KEYS:
        value = source.get(key)
        if value is None:
            continue
        theme[key] = value if isinstance(value, RGBColor) else _c(str(value))
    theme["on_dark"] = bool(source.get("on_dark", _luminance(theme["bg"]) < 0.42))
    return theme


# ============================================================
# 版式参数系统：每个版式的核心参数可配置
# 通过 ctx["layout_opts"] 覆盖默认值
# ============================================================

LAYOUT_DEFAULTS = {
    "cover": {
        "title_size": 54, "subtitle_size": 20, "kicker_size": 14,
        "title_y_ratio": 0.32, "deco_alpha": 0.25,
    },
    "bullets": {
        "title_size": 34, "body_size": 17, "card_gap": 0.18,
        "card_radius": 0.12, "number_size": 16,
    },
    "text_image": {
        "text_ratio": 0.55, "image_padding": 0.15,
    },
    "dashboard": {
        "metric_value_size": 36, "card_gap": 0.25, "cards_count": 4,
    },
    "timeline": {
        "axis_x_ratio": 0.50, "node_size": 0.36,
    },
    "comparison": {
        "divider_x_ratio": 0.50, "vs_size": 14,
    },
    "process": {
        "max_steps": 5, "card_gap": 0.35, "arrow_w": 0.5,
    },
    "quote": {
        "quote_size": 28, "source_size": 16, "deco_alpha": 0.15,
    },
}


def _opts(ctx: dict, layout_name: str) -> dict:
    """获取版式参数，合并默认值和用户覆盖。"""
    defaults = LAYOUT_DEFAULTS.get(layout_name, {})
    overrides = ctx.get("layout_opts", {}).get(layout_name, {})
    return {**defaults, **overrides}


def _merge_layout_opts(base: dict, overrides: dict) -> dict:
    """Deep-merge per-layout option dictionaries."""
    result = {name: dict(values) for name, values in (base or {}).items()}
    for name, values in (overrides or {}).items():
        result[name] = {**result.get(name, {}), **values}
    return result


def _resolve_font_pair(theme_key: str | None, profile: dict | None = None) -> dict[str, str]:
    """Resolve heading/body font pair from a theme key or template profile.
    Returns {"heading": ..., "body": ...}."""
    # Profile fonts take priority (from V2 adapter or user-supplied)
    if profile:
        fonts = profile.get("fonts")
        if fonts and isinstance(fonts, dict):
            heading = fonts.get("en") or fonts.get("heading") or fonts.get("display")
            body = fonts.get("cn") or fonts.get("body")
            if heading or body:
                return {
                    "heading": heading or FONTS.get("editorial", {}).get("heading", FONT_EN),
                    "body": body or FONTS.get("editorial", {}).get("body", FONT_CN),
                }
    # Built-in FONTS dict lookup
    if theme_key and theme_key in FONTS:
        return dict(FONTS[theme_key])
    return dict(FONTS.get("editorial", {"heading": FONT_EN, "body": FONT_CN}))


def choose_theme(key: str | None = None, profile: dict | None = None) -> dict:
    """Resolve a built-in theme or a JSON-friendly template profile.
    The returned dict includes 'font_heading' and 'font_body' keys from FONTS."""
    if profile:
        theme = theme_from_profile(profile)
    elif key and key in THEMES:
        theme = THEMES[key]
    else:
        theme = THEMES["editorial"]
    # Attach font pair to the theme dict
    font_pair = _resolve_font_pair(key, profile)
    theme["font_heading"] = font_pair["heading"]
    theme["font_body"] = font_pair["body"]
    return theme


# 字体：跨平台自动检测，中英文混排
def _detect_font_cn():
    """跨平台检测可用的中文字体。"""
    import platform
    system = platform.system()
    if system == "Windows":
        return "Microsoft YaHei"
    elif system == "Darwin":  # macOS
        return "PingFang SC"
    else:  # Linux
        return "Noto Sans CJK SC"

def _detect_font_en():
    """跨平台检测可用的英文字体。"""
    import platform
    system = platform.system()
    if system == "Windows":
        return "Arial"
    elif system == "Darwin":
        return "Helvetica Neue"
    else:
        return "Noto Sans"

FONT_CN = _detect_font_cn()
FONT_EN = _detect_font_en()
_DEFAULT_FONT_SIZE = Pt(18)


def _blank_layout(prs):
    """安全获取空白版式：优先用索引6，回退到搜索名称或最后一个布局。"""
    layouts = prs.slide_layouts
    # 尝试索引6（python-pptx 默认模板的 blank layout）
    if len(layouts) > 6:
        return layouts[6]
    # 回退：搜索名称含 "blank" 的布局
    for layout in layouts:
        if "blank" in layout.name.lower():
            return layout
    # 最终回退：使用最后一个布局（通常是 blank）
    return layouts[-1] if len(layouts) > 0 else layouts[0]

# ============================================================
# 底层工具函数
# ============================================================

def _set_font(run, name=FONT_CN, size=None, bold=None, color=None, italic=None, font_family=None):
    """统一设置 run 字体，并强制东亚文字使用同一字体（避免 fallback 到宋体）。
    font_family: 优先级高于 name；传入时覆盖 name 用于拉丁字体，同时设置东亚字体。"""
    effective_name = font_family if font_family else name
    run.font.name = effective_name
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    # 强制东亚字体一致性
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    # When font_family is provided, use CJK_FONTS body font for East-Asian text;
    # otherwise fall back to the effective_name (original behaviour).
    ea_font = CJK_FONTS.get("body", effective_name) if font_family else effective_name
    ea.set("typeface", ea_font)


def _no_line(shape):
    """去掉形状边框。"""
    shape.line.fill.background()
    return shape


def _solid(shape, color):
    """形状填充纯色。"""
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    return shape


def _set_bg(slide, color):
    """整页背景填充。"""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color
    return slide


def _luminance(color):
    """返回 0-1 的相对亮度（越大越亮）。"""
    return (0.299 * color[0] + 0.587 * color[1] + 0.114 * color[2]) / 255


def _text_on(bg_color, dark=None, light=None):
    """根据载体背景色亮度自动选择可读的前景色。
    背景亮→用 dark（默认 text），背景暗→用 light（默认 white）。
    用于「文字压在色块/圆角卡片上」的场景，避免浅底白字。"""
    if dark is None:
        dark = _c("1A1A1A")
    if light is None:
        light = _c("FFFFFF")
    return light if _luminance(bg_color) < 0.55 else dark


def _rounded_rect(slide, left, t, w, h, fill=None, line=None, radius_frac=0.08):
    """圆角矩形（信息卡片用）。radius_frac 控制圆角占比。"""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(left), Inches(t), Inches(w), Inches(h))
    if fill is not None:
        _solid(shp, fill)
    if line is None:
        _no_line(shp)
    try:
        shp.adjustments[0] = radius_frac
    except Exception:
        pass
    return shp


def _rect(slide, left, t, w, h, fill=None):
    """直角矩形。"""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(left), Inches(t), Inches(w), Inches(h))
    if fill is not None:
        _solid(shp, fill)
    _no_line(shp)
    return shp


def _oval(slide, left, t, w, h, fill=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                 Inches(left), Inches(t), Inches(w), Inches(h))
    if fill is not None:
        _solid(shp, fill)
    _no_line(shp)
    return shp


def _textbox(slide, left, t, w, h, anchor=MSO_ANCHOR.TOP, wrap=True):
    """文本框，返回 text_frame。"""
    tb = slide.shapes.add_textbox(Inches(left), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    # 关闭自动调整，用固定尺寸保证留白
    try:
        tf.auto_size = None
    except Exception:
        pass
    return tf


def _add_para(tf, text, *, first=False, size=None, color=None, bold=False,
             name=FONT_CN, align=PP_ALIGN.LEFT, space_before=0, space_after=0,
             level=0, line_spacing=1.15, font_family=None):
    """添加一个段落并设置样式。first=True 时复用首段。
    font_family: 主题字体（拉丁），传入时覆盖 name 参数。"""
    if size is None:
        size = _DEFAULT_FONT_SIZE
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.text = text
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    if space_after:
        p.space_after = Pt(space_after)
    p.level = level
    try:
        p.line_spacing = line_spacing
    except Exception:
        pass
    if p.runs:
        _set_font(p.runs[0], name=name, size=size, bold=bold, color=color, font_family=font_family)
    return p


def _fit_image_in_box(slide, img_path, left, t, w, h, cover=True):
    """
    把图片放入 (left,t,w,h) 盒子，cover=True 裁剪填满，False 则居中适配留白。
    返回 picture shape。若图片缺失则跳过返回 None。
    """
    if not img_path or not os.path.exists(img_path):
        return None
    # 计算真实比例
    iw = ih = None
    if Image is not None:
        try:
            with Image.open(img_path) as im:
                iw, ih = im.size
        except Exception:
            iw = ih = None
    if iw and ih:
        box_ratio = w / h
        img_ratio = iw / ih
        if cover:
            # 填满盒子：图片严格落在盒子内，用真实 crop fractions 裁掉溢出部分。
            # 不再通过放大图片越出盒子模拟 cover（会遮挡相邻区域）。
            pic = slide.shapes.add_picture(img_path, Inches(left), Inches(t),
                                           Inches(w), Inches(h))
            if img_ratio > box_ratio:
                crop = 1.0 - box_ratio / img_ratio
                pic.crop_left = crop / 2
                pic.crop_right = crop / 2
            else:
                crop = 1.0 - img_ratio / box_ratio
                pic.crop_top = crop / 2
                pic.crop_bottom = crop / 2
            return pic
        else:
            new_w = min(w, h * img_ratio)
            new_h = new_w / img_ratio
            pl = left + (w - new_w) / 2
            pt = t + (h - new_h) / 2
            return slide.shapes.add_picture(img_path, Inches(pl), Inches(pt),
                                            Inches(new_w), Inches(new_h))
    # 无 Pillow 或读取失败，直接按盒子塞
    return slide.shapes.add_picture(img_path, Inches(left), Inches(t),
                                    Inches(w), Inches(h))

# ============================================================
# 版式实现：每个函数画一页，视觉风格差异化
# 命名约定: layout_<名字>, 签名 (prs, theme, ctx) -> slide
# ctx 为 dict, 含 title/subtitle/bullets/images/chart/kicker 等
# ============================================================

def _set_alpha_last(slide, alpha_frac):
    """给最后加入的形状设置透明度 (0~1)。通过 lxml 改 solidFill 的 alpha。"""
    try:
        shp = slide.shapes[-1]
        sp = shp._element
        srgb = sp.find(".//" + qn("a:srgbClr"))
        if srgb is not None:
            alpha_el = etree.SubElement(srgb, qn("a:alpha"))
            alpha_el.set("val", str(int(alpha_frac * 100000)))
    except Exception:
        pass


def _header_band(slide, theme, title, kicker=None):
    """通用页眉:左侧主色竖条 + 标题 + 可选眉标。不画顶部横条(太模板味)。"""
    _rect(slide, MARGIN_X, MARGIN_TOP, 0.12, 0.7, fill=theme["primary"])
    tx = MARGIN_X + 0.35
    if kicker:
        kf = _textbox(slide, tx, MARGIN_TOP - 0.05, CONTENT_W - 0.35, 0.3)
        _add_para(kf, kicker.upper(), first=True, size=Pt(12),
                  color=theme["accent"], bold=True, font_family=theme.get("font_heading", FONT_EN), space_after=0)
        ty = MARGIN_TOP + 0.28
    else:
        ty = MARGIN_TOP
    tf = _textbox(slide, tx, ty, CONTENT_W - 0.35, 0.7, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(tf, title, first=True, size=Pt(34), color=theme["text"],
              bold=True, font_family=theme.get("font_body", FONT_CN))


def layout_cover(prs, theme, ctx):
    """封面:全屏背景图 + 深色蒙层 + 大标题。最具视觉冲击。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    img = ctx.get("cover_image")
    if img and os.path.exists(img):
        slide.shapes.add_picture(img, 0, 0, Inches(SLIDE_W), Inches(SLIDE_H))
        _rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=theme["dark"])
        _set_alpha_last(slide, 0.45)
    else:
        _set_bg(slide, theme["bg"])
        _oval(slide, SLIDE_W - 3.5, -2, 5, 5, fill=theme["primary"])
        _set_alpha_last(slide, 0.25)
        _oval(slide, -1.5, SLIDE_H - 3, 4, 4, fill=theme["accent"])
        _set_alpha_last(slide, 0.30)

    on_dark = bool(img and os.path.exists(img))  # 有封面图时背景被深蒙层压暗，否则是浅底
    title_color = theme["white"] if on_dark else _text_on(theme["bg"])
    kicker = ctx.get("kicker")
    if kicker:
        kf = _textbox(slide, MARGIN_X, 1.9, CONTENT_W, 0.4)
        _add_para(kf, kicker.upper(), first=True, size=Pt(14),
                  color=theme["accent"] if on_dark else theme["primary"],
                  bold=True, font_family=theme.get('font_heading', FONT_EN))
    tf = _textbox(slide, MARGIN_X, 2.4, CONTENT_W, 2.0, anchor=MSO_ANCHOR.BOTTOM)
    _add_para(tf, ctx.get("title", ""), first=True, size=Pt(54),
              color=title_color, bold=True, font_family=theme.get('font_body', FONT_CN), line_spacing=1.1)
    sub = ctx.get("subtitle")
    if sub:
        sf = _textbox(slide, MARGIN_X, 4.5, CONTENT_W * 0.7, 0.8)
        _add_para(sf, sub, first=True, size=Pt(20),
                  color=theme["text_muted"] if on_dark else theme["secondary"],
                  font_family=theme.get('font_body', FONT_CN))
    # 底部分隔线
    line_color = theme["accent"] if on_dark else theme["primary"]
    _rect(slide, MARGIN_X, 6.7, 1.2, 0.06, fill=line_color)
    return slide

def layout_toc(prs, theme, ctx):
    """目录页:左侧大字'目录'+ 右侧两列章节列表,每项带编号。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    # 左侧标题区
    lf = _textbox(slide, MARGIN_X, MARGIN_TOP, 4.5, 2.0, anchor=MSO_ANCHOR.BOTTOM)
    _add_para(lf, ctx.get("kicker", "CONTENTS").upper(), first=True,
              size=Pt(13), color=theme["accent"], bold=True, font_family=theme.get('font_heading', FONT_EN))
    _add_para(lf, ctx.get("title", "目录"), size=Pt(44), color=theme["text"],
              bold=True, font_family=theme.get('font_body', FONT_CN), space_before=4)
    # 右侧目录项
    items = ctx.get("toc_items") or ctx.get("bullets") or []
    n = len(items)
    if n == 0:
        return slide
    cols = 1 if n <= 5 else 2
    col_w = (CONTENT_W - 5.0) / cols - 0.4
    rx = MARGIN_X + 5.0
    rows = math.ceil(n / cols)
    rh = min(0.95, CONTENT_H / max(rows, 1))
    for i, item in enumerate(items):
        col = i // rows
        row = i % rows
        ix = rx + col * (col_w + 0.4)
        iy = MARGIN_TOP + 0.2 + row * rh
        # 编号
        nf = _textbox(slide, ix, iy, 0.7, rh, anchor=MSO_ANCHOR.MIDDLE)
        _add_para(nf, f"{i+1:02d}", first=True, size=Pt(22),
                  color=theme["primary"], bold=True, font_family=theme.get('font_heading', FONT_EN))
        # 标题
        tf = _textbox(slide, ix + 0.8, iy, col_w - 0.8, rh,
                      anchor=MSO_ANCHOR.MIDDLE)
        _add_para(tf, item, first=True, size=Pt(16),
                  color=theme["text"], font_family=theme.get('font_body', FONT_CN))
        # 下划线
        _rect(slide, ix + 0.8, iy + rh - 0.12, col_w - 0.8, 0.02,
              fill=theme["bg_alt"])
    return slide


def layout_section(prs, theme, ctx):
    """章节分隔页:大号编号 + 章节名 + 装饰色块,过渡用。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    # 右侧大色块
    _rect(slide, SLIDE_W * 0.42, 0, SLIDE_W * 0.58, SLIDE_H, fill=theme["primary"])
    _set_alpha_last(slide, 0.92)
    num = ctx.get("section_number", "01")
    # 巨大编号(半透明衬底感)
    nf = _textbox(slide, SLIDE_W * 0.42 + 0.6, MARGIN_TOP, 5, 4,
                  anchor=MSO_ANCHOR.MIDDLE)
    _add_para(nf, num, first=True, size=Pt(160),
              color=_text_on(theme["primary"]),
              bold=True, font_family=theme.get('font_heading', FONT_EN))
    # 章节标题(在左侧深底)
    tf = _textbox(slide, MARGIN_X, SLIDE_H / 2 - 0.5, 4.5, 1.4,
                  anchor=MSO_ANCHOR.MIDDLE)
    kk = ctx.get("kicker")
    if kk:
        _add_para(tf, kk.upper(), first=True, size=Pt(13),
                  color=theme["accent"], bold=True, font_family=theme.get('font_heading', FONT_EN), space_after=6)
    _add_para(tf, ctx.get("title", ""), size=Pt(38), color=theme["text"],
              bold=True, font_family=theme.get('font_body', FONT_CN))
    return slide


def layout_bullets(prs, theme, ctx):
    """要点页:左侧大序号 + 右侧要点卡片列表,每点独立成卡。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    bullets = ctx.get("bullets") or []
    # 顶部装饰:一个大序号水印
    if ctx.get("page_number"):
        wf = _textbox(slide, SLIDE_W - 3.2, MARGIN_TOP - 0.15, 2.5, 1.2,
                      anchor=MSO_ANCHOR.MIDDLE)
        _add_para(wf, str(ctx["page_number"]).zfill(2), first=True,
                  size=Pt(60), color=theme["bg_alt"], bold=True, font_family=theme.get('font_heading', FONT_EN))

    top = MARGIN_TOP + 1.1
    n = len(bullets)
    if n == 0:
        return slide
    gap = 0.18
    card_h = min(1.0, (CONTENT_H - 1.1 - gap * (n - 1)) / n)
    for i, b in enumerate(bullets):
        cy = top + i * (card_h + gap)
        _rounded_rect(slide, MARGIN_X, cy, CONTENT_W, card_h,
                      fill=theme["bg_alt"], radius_frac=0.12)
        # 左侧主色小条
        _rect(slide, MARGIN_X, cy, 0.1, card_h, fill=theme["primary"])
        # 序号圆
        _oval(slide, MARGIN_X + 0.3, cy + card_h / 2 - 0.22, 0.44, 0.44,
              fill=theme["primary"])
        ntf = _textbox(slide, MARGIN_X + 0.3, cy + card_h / 2 - 0.22,
                       0.44, 0.44, anchor=MSO_ANCHOR.MIDDLE)
        _add_para(ntf, str(i + 1), first=True, size=Pt(16),
                  color=_text_on(theme["primary"]),
                  bold=True, font_family=theme.get('font_heading', FONT_EN), align=PP_ALIGN.CENTER)
        # 文字
        tf = _textbox(slide, MARGIN_X + 1.0, cy, CONTENT_W - 1.3, card_h,
                      anchor=MSO_ANCHOR.MIDDLE)
        _add_para(tf, b, first=True, size=Pt(17), color=theme["text"],
                  font_family=theme.get('font_body', FONT_CN))
    return slide

def layout_text_image(prs, theme, ctx):
    """左文右图:文字区占 55%,图片占 40%,中间留缝。图片用圆角卡片包裹。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    bullets = ctx.get("bullets") or []
    img = None
    imgs = ctx.get("images") or []
    if imgs and os.path.exists(imgs[0]):
        img = imgs[0]
    # 左侧文字
    left_w = 6.2
    tf = _textbox(slide, MARGIN_X, MARGIN_TOP + 1.1, left_w, CONTENT_H - 1.1)
    for i, b in enumerate(bullets):
        _add_para(tf, b, first=(i == 0), size=Pt(17), color=theme["text"],
                  font_family=theme.get('font_body', FONT_CN), space_before=8 if i > 0 else 0, line_spacing=1.4)
    # 右侧图片
    if img:
        img_l = MARGIN_X + left_w + 0.5
        img_w = CONTENT_W - left_w - 0.5
        img_h = CONTENT_H - 1.1
        # 圆角卡片底
        _rounded_rect(slide, img_l - 0.15, MARGIN_TOP + 1.0,
                      img_w + 0.3, img_h + 0.2,
                      fill=theme["bg_alt"], radius_frac=0.06)
        _fit_image_in_box(slide, img, img_l, MARGIN_TOP + 1.1,
                          img_w, img_h, cover=False)
    return slide


def layout_full_image(prs, theme, ctx):
    """全屏大图页:图片铺满 + 底部渐变蒙层 + 标题叠加。视觉冲击力最强。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    img = None
    imgs = ctx.get("images") or []
    if imgs and os.path.exists(imgs[0]):
        img = imgs[0]
    if img:
        slide.shapes.add_picture(img, 0, 0, Inches(SLIDE_W), Inches(SLIDE_H))
        # 底部渐变蒙层(用半透明深色矩形模拟)
        _rect(slide, 0, SLIDE_H * 0.45, SLIDE_W, SLIDE_H * 0.55,
              fill=theme["dark"])
        _set_alpha_last(slide, 0.7)
        # 有图时标题压在深色蒙层上，用白色保证可读
        title_color = theme["white"]
        sub_color = theme["text_muted"] if theme.get("on_dark") else _text_on(theme["dark"], light=theme["white"])
    else:
        # 无图时背景为 theme["bg"]（可能为浅色），用 _text_on 根据亮度选色
        _set_bg(slide, theme["bg"])
        title_color = _text_on(theme["bg"])
        sub_color = theme["text_muted"]
    # 标题叠加在底部
    tf = _textbox(slide, MARGIN_X, SLIDE_H * 0.55, CONTENT_W, 1.5,
                  anchor=MSO_ANCHOR.TOP)
    _add_para(tf, ctx.get("title", ""), first=True, size=Pt(42),
              color=title_color, bold=True, font_family=theme.get('font_body', FONT_CN), line_spacing=1.1)
    sub = ctx.get("subtitle") or ctx.get("kicker")
    if sub:
        _add_para(tf, sub, size=Pt(18), color=sub_color,
                  font_family=theme.get('font_body', FONT_CN), space_before=8)
    # 底部小装饰线
    _rect(slide, MARGIN_X, SLIDE_H * 0.55 - 0.15, 1.0, 0.06,
          fill=theme["accent"])
    return slide


def layout_image_grid(prs, theme, ctx):
    """图墙/九宫格:3x3 或 2x3 图片网格,每格圆角,统一间距。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    imgs = ctx.get("images") or []
    n = len(imgs)
    if n == 0:
        return slide
    # 决定网格: 2列 or 3列
    cols = 3 if n >= 4 else 2
    rows = math.ceil(n / cols)
    gap = 0.2
    top = MARGIN_TOP + 1.2
    avail_w = CONTENT_W
    avail_h = CONTENT_H - 1.2
    cell_w = (avail_w - gap * (cols - 1)) / cols
    cell_h = (avail_h - gap * (rows - 1)) / rows
    for i, img_path in enumerate(imgs[:cols * rows]):
        if not os.path.exists(img_path):
            continue
        col = i % cols
        row = i // cols
        cl = MARGIN_X + col * (cell_w + gap)
        ct = top + row * (cell_h + gap)
        # 圆角卡片底
        _rounded_rect(slide, cl, ct, cell_w, cell_h,
                      fill=theme["bg_alt"], radius_frac=0.06)
        _fit_image_in_box(slide, img_path, cl + 0.08, ct + 0.08,
                          cell_w - 0.16, cell_h - 0.16, cover=True)
    return slide

def layout_dashboard(prs, theme, ctx):
    """数据大屏/仪表盘:顶部标题 + 3-4 个指标卡片 + 底部图表区。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    opts = _opts(ctx, "dashboard")
    max_cards = max(1, int(opts.get("cards_count", 4)))
    metrics = (ctx.get("metrics") or [])[:max_cards]
    n = len(metrics)
    if n == 0:
        n = 4
        metrics = [{"label": f"指标{i+1}", "value": "--", "change": ""} for i in range(n)]
    # 卡片底色用 bg_alt：深色主题 bg_alt 深→白字；浅色主题 bg_alt 浅→用 primary 保证可读
    value_color = theme["white"] if theme.get("on_dark") else theme["primary"]
    # 指标卡片行
    card_gap = float(opts.get("card_gap", 0.25))
    card_w = (CONTENT_W - card_gap * (n - 1)) / n
    card_h = 1.6
    card_top = MARGIN_TOP + 1.2
    for i, m in enumerate(metrics[:4]):
        cl = MARGIN_X + i * (card_w + card_gap)
        _rounded_rect(slide, cl, card_top, card_w, card_h,
                      fill=theme["bg_alt"], radius_frac=0.08)
        # 指标值(大号)
        vf = _textbox(slide, cl + 0.3, card_top + 0.2, card_w - 0.6, 0.8,
                      anchor=MSO_ANCHOR.BOTTOM)
        _add_para(vf, m.get("value", "--"), first=True,
                  size=Pt(opts.get("metric_value_size", 36)),
                  color=value_color, bold=True, font_family=theme.get('font_heading', FONT_EN))
        # 标签
        lf = _textbox(slide, cl + 0.3, card_top + 1.0, card_w - 0.6, 0.3)
        _add_para(lf, m.get("label", ""), first=True, size=Pt(13),
                  color=theme["text_muted"], font_family=theme.get('font_body', FONT_CN))
        # 变化
        chg = m.get("change") or m.get("delta", "")
        if chg:
            cf = _textbox(slide, cl + 0.3, card_top + 1.25, card_w - 0.6, 0.25)
            chg_color = _c("22C55E") if chg.startswith("+") else _c("EF4444") if chg.startswith("-") else theme["text_muted"]
            _add_para(cf, chg, first=True, size=Pt(12), color=chg_color,
                      bold=True, font_family=theme.get('font_heading', FONT_EN))
    # 底部图表区：用形状模拟简单柱状图
    chart_top = card_top + card_h + 0.3
    chart_h = CONTENT_H - 1.2 - card_h - 0.3
    _rounded_rect(slide, MARGIN_X, chart_top, CONTENT_W, chart_h,
                  fill=theme["bg_alt"], radius_frac=0.06)
    # 从 metrics 的 change 值提取数据画柱状图
    chart_data_vals = []
    for m in metrics[:4]:
        chg = m.get("change", "")
        try:
            val = int(chg.replace("+", "").replace("%", ""))
            chart_data_vals.append(val)
        except (ValueError, AttributeError):
            chart_data_vals.append(20)
    if chart_data_vals:
        max_val = max(abs(v) for v in chart_data_vals) or 1
        bar_area_l = MARGIN_X + 1.5
        bar_area_w = CONTENT_W - 3.0
        bar_area_t = chart_top + 0.4
        bar_area_h = chart_h - 1.0
        n_bars = len(chart_data_vals)
        bar_gap = 0.3
        bar_w = (bar_area_w - bar_gap * (n_bars - 1)) / n_bars
        bar_colors = [theme["primary"], theme["secondary"],
                      theme["accent"], theme["primary"]]
        baseline_y = bar_area_t + bar_area_h * 0.5
        for bi, bv in enumerate(chart_data_vals):
            bh = (abs(bv) / max_val) * bar_area_h * 0.4
            bl = bar_area_l + bi * (bar_w + bar_gap)
            if bv >= 0:
                bt = baseline_y - bh
            else:
                bt = baseline_y
            _rounded_rect(slide, bl, bt, bar_w, bh,
                          fill=bar_colors[bi % len(bar_colors)], radius_frac=0.15)
            # 数值标签
            lf = _textbox(slide, bl, bt - 0.35, bar_w, 0.3)
            _add_para(lf, metrics[bi].get("change") or metrics[bi].get("delta", ""), first=True,
                      size=Pt(11), color=theme["text_muted"], font_family=theme.get('font_heading', FONT_EN),
                      align=PP_ALIGN.CENTER)
            # 底部标签
            blf = _textbox(slide, bl, bar_area_t + bar_area_h + 0.05, bar_w, 0.3)
            _add_para(blf, metrics[bi].get("label", ""), first=True,
                      size=Pt(10), color=theme["text_muted"], font_family=theme.get('font_body', FONT_CN),
                      align=PP_ALIGN.CENTER)
        # 底部基线
        _rect(slide, bar_area_l, bar_area_t + bar_area_h,
              bar_area_w, 0.02, fill=theme["text_muted"])
    return slide


def layout_timeline(prs, theme, ctx):
    """时间线:垂直时间轴 + 左右交替事件节点。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    events = ctx.get("events") or ctx.get("bullets") or []
    if len(events) > 6:
        print(f"  [警告] timeline 版式最多支持6个事件，当前{len(events)}个，超出的事件被截断", file=__import__('sys').stderr)
    n = len(events)
    if n == 0:
        return slide
    # 时间轴竖线
    axis_x = SLIDE_W / 2
    top = MARGIN_TOP + 1.3
    axis_h = CONTENT_H - 1.3
    _rect(slide, axis_x - 0.03, top, 0.06, axis_h, fill=theme["primary"])
    # 事件节点
    node_gap = axis_h / max(n, 1)
    for i, ev in enumerate(events[:6]):
        ny = top + i * node_gap + node_gap / 2
        # 圆点
        _oval(slide, axis_x - 0.18, ny - 0.18, 0.36, 0.36, fill=theme["accent"])
        # 交替左右
        if i % 2 == 0:
            # 左侧
            tf = _textbox(slide, MARGIN_X, ny - 0.35, axis_x - MARGIN_X - 0.5, 0.7,
                          anchor=MSO_ANCHOR.MIDDLE)
            align = PP_ALIGN.RIGHT
        else:
            tf = _textbox(slide, axis_x + 0.5, ny - 0.35,
                          SLIDE_W - MARGIN_X - axis_x - 0.5, 0.7,
                          anchor=MSO_ANCHOR.MIDDLE)
            align = PP_ALIGN.LEFT
        if isinstance(ev, dict):
            _add_para(tf, ev.get("date", ""), first=True, size=Pt(12),
                      color=theme["accent"], bold=True, font_family=theme.get('font_heading', FONT_EN), align=align)
            _add_para(tf, ev.get("title", ev.get("text", "")), size=Pt(15),
                      color=theme["text"], font_family=theme.get('font_body', FONT_CN), align=align, space_before=2)
        else:
            _add_para(tf, str(ev), first=True, size=Pt(15),
                      color=theme["text"], font_family=theme.get('font_body', FONT_CN), align=align)
    return slide

def layout_comparison(prs, theme, ctx):
    """对比页:左右两列,中间分隔线,各自标题+要点。用于竞品/方案对比。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    left_data = ctx.get("left") or {}
    right_data = ctx.get("right") or {}
    col_w = (CONTENT_W - 0.3) / 2
    top = MARGIN_TOP + 1.2
    # 中间分隔竖线
    mid_x = MARGIN_X + col_w + 0.15
    _rect(slide, mid_x - 0.015, top, 0.03, CONTENT_H - 1.2,
          fill=theme["primary"])
    # VS 标签
    _oval(slide, mid_x - 0.3, top + 0.1, 0.6, 0.6, fill=theme["accent"])
    vstf = _textbox(slide, mid_x - 0.3, top + 0.1, 0.6, 0.6,
                    anchor=MSO_ANCHOR.MIDDLE)
    _add_para(vstf, "VS", first=True, size=Pt(14),
              color=_text_on(theme["accent"]),
              bold=True, font_family=theme.get('font_heading', FONT_EN), align=PP_ALIGN.CENTER)
    # 左列
    ltitle = left_data.get("title", "")
    if ltitle:
        ltf = _textbox(slide, MARGIN_X, top, col_w, 0.5)
        _add_para(ltf, ltitle, first=True, size=Pt(22), color=theme["primary"],
                  bold=True, font_family=theme.get('font_body', FONT_CN))
    lbullets = left_data.get("bullets", [])
    lbf = _textbox(slide, MARGIN_X, top + 0.6, col_w, CONTENT_H - 1.8)
    for i, b in enumerate(lbullets):
        _add_para(lbf, b, first=(i == 0), size=Pt(16), color=theme["text"],
                  font_family=theme.get('font_body', FONT_CN), space_before=6 if i > 0 else 0, line_spacing=1.35)
    # 右列
    rx = mid_x + 0.15
    rtitle = right_data.get("title", "")
    if rtitle:
        rtf = _textbox(slide, rx, top, col_w, 0.5)
        _add_para(rtf, rtitle, first=True, size=Pt(22), color=theme["secondary"],
                  bold=True, font_family=theme.get('font_body', FONT_CN))
    rbullets = right_data.get("bullets", [])
    rbf = _textbox(slide, rx, top + 0.6, col_w, CONTENT_H - 1.8)
    for i, b in enumerate(rbullets):
        _add_para(rbf, b, first=(i == 0), size=Pt(16), color=theme["text"],
                  font_family=theme.get('font_body', FONT_CN), space_before=6 if i > 0 else 0, line_spacing=1.35)
    return slide


def layout_quote(prs, theme, ctx):
    """引用金句页:超大引号装饰 + 引文 + 出处,极简震撼。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    # 巨大装饰引号
    qtf = _textbox(slide, MARGIN_X - 0.1, 0.8, 3, 3)
    _add_para(qtf, "“", first=True, size=Pt(200), color=theme["primary"],
              bold=True, font_family=theme.get('font_heading', FONT_EN), align=PP_ALIGN.LEFT)
    _set_alpha_on_textbox(qtf, 0.15)
    # 引文
    quote_text = ctx.get("quote") or ctx.get("title", "")
    qf = _textbox(slide, MARGIN_X + 1.0, 2.2, CONTENT_W - 1.5, 2.5,
                  anchor=MSO_ANCHOR.MIDDLE)
    _add_para(qf, quote_text, first=True, size=Pt(28), color=theme["text"],
              font_family=theme.get('font_body', FONT_CN), line_spacing=1.5)
    # Set italic on the run
    if qf.paragraphs and qf.paragraphs[0].runs:
        qf.paragraphs[0].runs[0].font.italic = True
    # 出处
    source = ctx.get("source") or ctx.get("subtitle")
    if source:
        sf = _textbox(slide, MARGIN_X + 1.0, 5.0, CONTENT_W - 1.5, 0.5)
        _add_para(sf, f"—— {source}", first=True, size=Pt(16),
                  color=theme["accent"], font_family=theme.get('font_body', FONT_CN))
    # 装饰线
    _rect(slide, MARGIN_X + 1.0, 4.7, 2.0, 0.04, fill=theme["accent"])
    return slide


def _set_alpha_on_textbox(tf, alpha_frac):
    """给文本框内所有 run 的字体颜色设置透明度。"""
    try:
        for para in tf.paragraphs:
            for run in para.runs:
                rPr = run._r.get_or_add_rPr()
                srgb = rPr.find(".//" + qn("a:srgbClr"))
                if srgb is not None:
                    # 移除已有 alpha
                    for old in srgb.findall(qn("a:alpha")):
                        srgb.remove(old)
                    alpha_el = etree.SubElement(srgb, qn("a:alpha"))
                    alpha_el.set("val", str(int(alpha_frac * 100000)))
    except Exception:
        pass


def layout_process(prs, theme, ctx):
    """流程图页:横向箭头流程,3-5 步,每步一个圆角卡片。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    steps = ctx.get("steps") or ctx.get("bullets") or []
    n = len(steps)
    if n == 0:
        return slide
    if n > 5:
        print(f"  [警告] process 版式最多支持5步，当前{n}步，超出的步骤被截断", file=__import__('sys').stderr)
    n = min(n, 5)
    gap = 0.35
    arrow_w = 0.5
    card_w = (CONTENT_W - gap * (n - 1) - arrow_w * (n - 1)) / n
    card_h = 3.2
    top = MARGIN_TOP + 1.5
    colors_cycle = [theme["primary"], theme["secondary"],
                    theme["accent"], theme["primary"], theme["secondary"]]
    for i in range(n):
        cl = MARGIN_X + i * (card_w + gap + arrow_w)
        c = colors_cycle[i % len(colors_cycle)]
        # 卡片
        _rounded_rect(slide, cl, top, card_w, card_h, fill=theme["bg_alt"],
                      radius_frac=0.08)
        # 顶部色条
        _rect(slide, cl, top, card_w, 0.12, fill=c)
        # 步骤编号
        nf = _textbox(slide, cl, top + 0.3, card_w, 0.6, anchor=MSO_ANCHOR.MIDDLE)
        _add_para(nf, f"STEP {i+1}", first=True, size=Pt(13),
                  color=c, bold=True, font_family=theme.get('font_heading', FONT_EN), align=PP_ALIGN.CENTER)
        # 步骤文字
        step_text = steps[i] if isinstance(steps[i], str) else steps[i].get("title", "")
        sf = _textbox(slide, cl + 0.2, top + 1.0, card_w - 0.4, card_h - 1.2,
                      anchor=MSO_ANCHOR.TOP)
        _add_para(sf, step_text, first=True, size=Pt(16), color=theme["text"],
                  font_family=theme.get('font_body', FONT_CN), align=PP_ALIGN.CENTER, line_spacing=1.3)
        # 箭头(用右箭头形状)
        if i < n - 1:
            ax = cl + card_w + 0.05
            ay = top + card_h / 2 - 0.2
            arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                         Inches(ax), Inches(ay),
                                         Inches(arrow_w - 0.1), Inches(0.4))
            _solid(arr, theme["text_muted"])
            _no_line(arr)
    return slide

def layout_table(prs, theme, ctx):
    """表格页:标题 + 数据表格,表头主色,交替行色。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, ctx.get("title", ""), ctx.get("kicker"))
    headers = ctx.get("table_headers") or []
    rows = ctx.get("table_rows") or []
    if not headers or not rows:
        # 兜底:用 bullets 展示
        tf = _textbox(slide, MARGIN_X, MARGIN_TOP + 1.2, CONTENT_W, CONTENT_H - 1.2)
        for i, b in enumerate(ctx.get("bullets", [])):
            _add_para(tf, b, first=(i == 0), size=Pt(17), color=theme["text"],
                      font_family=theme.get('font_body', FONT_CN), space_before=6 if i > 0 else 0)
        return slide
    n_cols = len(headers)
    n_rows = len(rows)
    col_w = CONTENT_W / n_cols
    top = MARGIN_TOP + 1.2
    row_h = min(0.55, (CONTENT_H - 1.2) / (n_rows + 1))
    # 表头
    for j, h in enumerate(headers):
        _rect(slide, MARGIN_X + j * col_w, top, col_w, row_h, fill=theme["primary"])
        tf = _textbox(slide, MARGIN_X + j * col_w, top, col_w, row_h,
                      anchor=MSO_ANCHOR.MIDDLE)
        _add_para(tf, h, first=True, size=Pt(14),
                  color=_text_on(theme["primary"]),
                  bold=True, font_family=theme.get('font_body', FONT_CN), align=PP_ALIGN.CENTER)
    # 数据行
    for i, row in enumerate(rows):
        ry = top + (i + 1) * row_h
        bg = theme["bg_alt"] if i % 2 == 0 else theme["bg"]
        for j, cell in enumerate(row[:n_cols]):
            _rect(slide, MARGIN_X + j * col_w, ry, col_w, row_h, fill=bg)
            tf = _textbox(slide, MARGIN_X + j * col_w, ry, col_w, row_h,
                          anchor=MSO_ANCHOR.MIDDLE)
            _add_para(tf, str(cell), first=True, size=Pt(13),
                      color=theme["text"], font_family=theme.get('font_body', FONT_CN), align=PP_ALIGN.CENTER)
    return slide


def layout_end(prs, theme, ctx):
    """结尾页:深色底 + 居中大字 + 装饰元素,与封面呼应。"""
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["dark"])
    # 装饰圆(与封面呼应)
    _oval(slide, SLIDE_W - 4, SLIDE_H - 4, 6, 6, fill=theme["primary"])
    _set_alpha_last(slide, 0.08)
    _oval(slide, -2, -2, 5, 5, fill=theme["accent"])
    _set_alpha_last(slide, 0.06)
    # 主文字
    text = ctx.get("title") or ctx.get("end_text") or "谢谢"
    tf = _textbox(slide, MARGIN_X, SLIDE_H / 2 - 1, CONTENT_W, 2.0,
                  anchor=MSO_ANCHOR.MIDDLE)
    _add_para(tf, text, first=True, size=Pt(56), color=theme["white"],
              bold=True, font_family=theme.get('font_body', FONT_CN), align=PP_ALIGN.CENTER, line_spacing=1.1)
    # 副文字
    sub = ctx.get("subtitle") or ctx.get("source")
    if sub:
        sf = _textbox(slide, MARGIN_X, SLIDE_H / 2 + 1.2, CONTENT_W, 0.6)
        _add_para(sf, sub, first=True, size=Pt(16), color=theme["text_muted"],
                  font_family=theme.get('font_body', FONT_CN), align=PP_ALIGN.CENTER)
    # 底部装饰线
    _rect(slide, SLIDE_W / 2 - 0.6, SLIDE_H - 1.2, 1.2, 0.04,
          fill=theme["accent"])
    return slide


# ============================================================
# 版式注册表 & 智能版式选择器
# ============================================================

# 所有版式函数的注册表
LAYOUT_REGISTRY: dict[str, Callable] = {
    "cover": layout_cover,
    "toc": layout_toc,
    "section": layout_section,
    "bullets": layout_bullets,
    "text_image": layout_text_image,
    "full_image": layout_full_image,
    "image_grid": layout_image_grid,
    "dashboard": layout_dashboard,
    "timeline": layout_timeline,
    "comparison": layout_comparison,
    "quote": layout_quote,
    "process": layout_process,
    "table": layout_table,
    "end": layout_end,
}


def choose_layout(section: dict, position: int, total: int) -> str:
    """
    智能版式选择器：根据章节内容特征自动挑选最佳版式。

    决策逻辑（优先级从高到低）：
    1. 特殊位置：首页→cover, 末页→end
       （toc 目录页由 auto_generate_ppt 主流程单独生成，不在此处强制触发，
        以免占用用户传入的第一个内容章节）
    2. 内容特征信号：
       - 有 metrics → dashboard
       - 有 events/timeline → timeline
       - 有 left+right → comparison
       - 有 quote → quote
       - 有 steps/process → process
       - 有 table_headers → table
       - 有 3+ images → image_grid
       - 有 1 image + bullets → text_image
       - 有 1 image + 无 bullets → full_image
       - 有 bullets (3-7) → bullets
       - 无 bullets → section (章节分隔)

    注：连续 2 页同版式的防重复切换逻辑由 auto_generate_ppt 主循环负责，
        不在本函数内实现。

    参数:
        section: 章节数据 dict
        position: 当前章节在整本 PPT 中的位置 (0-based)
        total: 总章节数
    返回:
        版式名称字符串
    """
    # 1. 特殊位置
    # 注：toc 不在此处强制触发，由 auto_generate_ppt 主流程单独生成目录页，
    #    避免占用用户传入的第一个内容章节。
    if position == 0:
        return "cover"
    if position == total - 1:
        return "end"

    # 2. 内容特征信号 (兼容 dict 和 Section dataclass)
    def _get(obj, key, default=None):
        if isinstance(obj, dict):
            return obj.get(key, default)
        return getattr(obj, key, default)

    has_bullets = bool(_get(section, "bullets"))
    images = _get(section, "images") or []
    n_images = len([im for im in images if im and os.path.exists(im)])

    if _get(section, "metrics"):
        return "dashboard"
    if _get(section, "events") or _get(section, "timeline"):
        return "timeline"
    if _get(section, "left") and _get(section, "right"):
        return "comparison"
    if _get(section, "quote"):
        return "quote"
    if _get(section, "steps") or _get(section, "process"):
        return "process"
    if _get(section, "table_headers"):
        return "table"
    if n_images >= 3:
        return "image_grid"
    if n_images >= 1 and has_bullets:
        return "text_image"
    if n_images >= 1 and not has_bullets:
        return "full_image"
    if has_bullets:
        return "bullets"
    # 无内容信号 → 章节分隔页
    return "section"


# ============================================================
# 高级 API：自动生成完整 PPT（智能排版 + 配图 + 图表）
# ============================================================

from dataclasses import asdict
from dataclasses import field as dc_field


@dataclass
class Section:
    """PPT 章节数据结构。"""
    title: str = ""
    subtitle: str = ""
    bullets: list = dc_field(default_factory=list)
    images: list = dc_field(default_factory=list)
    image_query: str = ""
    layout: str = ""
    kicker: str = ""
    metrics: list = dc_field(default_factory=list)
    events: list = dc_field(default_factory=list)
    steps: list = dc_field(default_factory=list)
    table_headers: list = dc_field(default_factory=list)
    table_rows: list = dc_field(default_factory=list)
    left: dict = dc_field(default_factory=dict)
    right: dict = dc_field(default_factory=dict)
    quote: str = ""
    source: str = ""
    layout_opts: dict = dc_field(default_factory=dict)
    section_number: str = ""
    page_number: int = 0



def auto_generate_ppt(
    title, subtitle="", sections=None, output_path="output.pptx",
    theme_key=None, image_dir="./ppt_images",
    auto_search_images=True, lang="zh", template_key=None,
    template_profile=None, template_catalog_path=None,
):
    if sections is None:
        sections = []
    if template_key and template_profile is None:
        from template_engine import load_template_profile
        template_profile = load_template_profile(template_key, template_catalog_path)
    template_profile = template_profile or {}
    theme = choose_theme(theme_key, template_profile or None)
    template_layout_opts = template_profile.get("layout_opts", {})
    layout_family = template_profile.get("layout_family", "standard")

    def resolve_layout(layout_name, fallback):
        if layout_family != "standard":
            try:
                from layout_variants import get_layout_function
                variant = get_layout_function(layout_family, layout_name)
                if variant is not None:
                    return variant
            except (ImportError, ValueError):
                pass
        return LAYOUT_REGISTRY.get(layout_name, fallback)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    parsed = []
    for i, s in enumerate(sections):
        if isinstance(s, Section):
            sec = s
        else:
            valid_keys = set(Section.__dataclass_fields__.keys())
            sec = Section(**{k: v for k, v in s.items() if k in valid_keys})
        sec.section_number = sec.section_number or str(i).zfill(2)
        sec.page_number = i + 1
        parsed.append(sec)

    if auto_search_images:
        _auto_search_images(parsed, image_dir, lang)

    total = len(parsed) + 2
    used_layouts = []

    # 为封面单独搜索背景图
    cover_image = None
    if auto_search_images:
        cover_image = _search_cover_image(title, image_dir, lang)
    if not cover_image and parsed and parsed[0].images:
        cover_image = parsed[0].images[0]

    cover_ctx = {
        "title": title, "subtitle": subtitle,
        "cover_image": cover_image,
        "layout_opts": template_layout_opts,
    }
    resolve_layout("cover", layout_cover)(prs, theme, cover_ctx)
    used_layouts.append("cover")

    # 目录页：当章节数较多时（>3），在封面后单独生成目录页，
    # 使用各章节标题作为目录项，不占用用户传入的内容章节。
    if len(parsed) > 3:
        toc_ctx = {
            "title": "目录",
            "kicker": "CONTENTS",
            "toc_items": [s.title for s in parsed if s.title],
            "layout_opts": template_layout_opts,
        }
        resolve_layout("toc", layout_toc)(prs, theme, toc_ctx)
        used_layouts.append("toc")

    for i, sec in enumerate(parsed):
        ctx = {
            "title": sec.title, "subtitle": sec.subtitle,
            "bullets": sec.bullets, "images": sec.images,
            "kicker": sec.kicker, "section_number": sec.section_number,
            "page_number": sec.page_number, "metrics": sec.metrics,
            "events": sec.events, "steps": sec.steps,
            "table_headers": sec.table_headers, "table_rows": sec.table_rows,
            "left": sec.left, "right": sec.right,
            "quote": sec.quote, "source": sec.source,
            "layout_opts": _merge_layout_opts(template_layout_opts, sec.layout_opts),
        }
        if sec.layout and sec.layout in LAYOUT_REGISTRY:
            layout_name = sec.layout
        else:
            layout_name = choose_layout(sec, i + 1, total)

        if (len(used_layouts) >= 2
                and used_layouts[-1] == layout_name
                and used_layouts[-2] == layout_name):
            for alt in ["bullets", "text_image", "full_image", "process"]:
                if alt != layout_name and alt in LAYOUT_REGISTRY:
                    layout_name = alt
                    break

        layout_fn = resolve_layout(layout_name, layout_bullets)
        _warn_unconsumed(layout_name, ctx)
        layout_fn(prs, theme, ctx)
        used_layouts.append(layout_name)

    end_ctx = {
        "title": template_profile.get("end_title", "谢谢"),
        "subtitle": subtitle,
        "layout_opts": template_layout_opts,
    }
    resolve_layout("end", layout_end)(prs, theme, end_ctx)
    used_layouts.append("end")

    abs_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_path) or ".", exist_ok=True)
    prs.save(abs_path)
    try:
        json.dumps(template_profile)
        serializable_profile = template_profile or None
    except TypeError:
        serializable_profile = None
    try:
        from ppt_project import save_project_manifest
        save_project_manifest(abs_path, {
            "skill_version": 2,
            "title": title,
            "subtitle": subtitle,
            "theme_key": theme_key,
            "template_key": template_key,
            "template_profile": serializable_profile,
            "image_dir": os.path.abspath(image_dir),
            "lang": lang,
            "auto_search_images": auto_search_images,
            "sections": [asdict(section) for section in parsed],
            "layouts": used_layouts,
        })
    except Exception as exc:
        print(f"  [警告] 无法保存项目 manifest: {exc}")
    print(f"PPT 已保存至: {abs_path}")
    print(f"版式使用序列: {used_layouts}")
    return abs_path


def _load_pixabay_search():
    """延迟加载同目录下的 pixabay_search 模块，返回 search_and_download 函数。
    加载失败时返回 None。两个调用方（_auto_search_images / _search_cover_image）共用此函数。"""
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "pixabay_search",
            os.path.join(os.path.dirname(__file__), "pixabay_search.py"),
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod.search_and_download
    except Exception as e:
        print(f"  [警告] 无法加载 pixabay_search: {e}")
        return None


def _auto_search_images(sections, image_dir, lang):
    search_and_download = _load_pixabay_search()
    if search_and_download is None:
        return
    os.makedirs(image_dir, exist_ok=True)
    for sec in sections:
        if sec.images:
            continue
        query = sec.image_query or sec.title
        if not query:
            continue
        try:
            paths = _retry_search(
                search_and_download, query,
                count=3, output_dir=image_dir,
                orientation="horizontal", lang=lang,
                size="large", min_width=1280,
            )
            if paths:
                sec.images = paths
        except Exception as e:
            print(f"  [警告] 搜索 '{query}' 失败: {e}")


def _retry_search(search_and_download_fn, query, max_retries=3, **kwargs):
    """带指数退避重试的图片搜索。"""
    import time
    for attempt in range(max_retries):
        try:
            paths = search_and_download_fn(query=query, **kwargs)
            if paths:
                return paths
        except Exception as e:
            if attempt < max_retries - 1:
                wait = 2 ** attempt  # 1s, 2s, 4s
                print(f"  [重试] 搜索 '{query}' 失败 ({attempt+1}/{max_retries})，{wait}s 后重试: {e}")
                time.sleep(wait)
            else:
                print(f"  [警告] 搜索 '{query}' 最终失败 ({max_retries}次): {e}")
    return []


def _search_cover_image(title, image_dir, lang):
    """为封面单独搜索一张高质量背景图。"""
    search_and_download = _load_pixabay_search()
    if search_and_download is None:
        return None
    os.makedirs(image_dir, exist_ok=True)
    # 用标题关键词搜索
    query = title
    paths = _retry_search(
        search_and_download, query,
        count=1, output_dir=image_dir,
        orientation="horizontal", lang=lang,
        min_width=1920, size="large",
    )
    return paths[0] if paths else None


# ============================================================
# python-pptx chart wrapper
# ============================================================

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION


def add_chart_slide(prs, theme, title, chart_type, categories, series,
                    kicker="", position="full"):
    chart_map = {
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }
    xl_type = chart_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    _header_band(slide, theme, title, kicker)

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for s in series:
        chart_data.add_series(s.get("name", ""), s.get("values", []))

    if position == "full":
        cl, ct, cw, ch = MARGIN_X, MARGIN_TOP + 1.2, CONTENT_W, CONTENT_H - 1.2
    else:
        cl, ct, cw, ch = MARGIN_X, MARGIN_TOP + 1.2, CONTENT_W, (CONTENT_H - 1.2) * 0.55

    chart_frame = slide.shapes.add_chart(
        xl_type, Inches(cl), Inches(ct), Inches(cw), Inches(ch), chart_data)
    chart = chart_frame.chart
    chart.has_legend = len(series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    return slide


# ============================================================
# Legacy API compatibility
# ============================================================

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs

def add_title_slide(prs, main_title, subtitle="", background_color=None):
    theme = choose_theme()
    return layout_cover(prs, theme, {"title": main_title, "subtitle": subtitle})

def add_content_slide(prs, title, bullets, title_color=None):
    theme = choose_theme()
    return layout_bullets(prs, theme, {"title": title, "bullets": bullets})

def add_image_content_slide(prs, title, bullets, image_path, title_color=None):
    theme = choose_theme()
    ctx = {"title": title, "bullets": bullets, "images": [image_path] if image_path else []}
    return layout_text_image(prs, theme, ctx)

def add_section_slide(prs, section_title, section_number=""):
    theme = choose_theme()
    return layout_section(prs, theme, {"title": section_title, "section_number": section_number})

def add_end_slide(prs, text="谢谢观看"):
    theme = choose_theme()
    return layout_end(prs, theme, {"title": text})


def auto_validate_ppt(pptx_path):
    """
    自动化验收检查：分析 PPTX 文件，验证基本质量标准。
    返回 {"passed": bool, "checks": [...], "warnings": [...]}。
    """
    import os

    from pptx import Presentation

    if not os.path.exists(pptx_path):
        return {"passed": False, "checks": [], "warnings": [f"文件不存在: {pptx_path}"]}

    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    n_slides = len(slides)
    checks = []
    warnings = []
    passed = True

    # 1. 相邻版式不重复（通过shape数量和类型判断）
    if n_slides >= 3:
        prev_shape_count = None
        prev_prev_shape_count = None
        for i, slide in enumerate(slides):
            shape_count = len(slide.shapes)
            if (prev_shape_count is not None and prev_prev_shape_count is not None
                    and shape_count == prev_shape_count == prev_prev_shape_count):
                # 连续3页shape数量相同，可能重复布局
                warnings.append(f"第{i-1}-{i+1}页shape数量相同({shape_count})，可能版式重复")
            prev_prev_shape_count = prev_shape_count
            prev_shape_count = shape_count
    checks.append(("相邻版式多样性", True if not any("版式重复" in w for w in warnings) else None))

    # 2. 版式多样性：shape数量不重复率 > 50%
    shape_counts = [len(s.shapes) for s in slides]
    unique_counts = len(set(shape_counts))
    diversity_ratio = unique_counts / max(n_slides, 1)
    checks.append(("版式多样性", diversity_ratio >= 0.5))
    if diversity_ratio < 0.5:
        warnings.append(f"版式多样性不足: {unique_counts}/{n_slides}种不同的shape数量")
        passed = False

    # 3. 封面非空
    if n_slides >= 1:
        cover_shapes = len(slides[0].shapes)
        checks.append(("封面有内容", cover_shapes >= 2))
        if cover_shapes < 2:
            warnings.append("封面内容过少")
            passed = False

    # 4. 要点密度：按实际文字字符数判断内容是否臃肿，而非 shape 段落数。
    #    结构化版式（dashboard/timeline/process/comparison/image_grid/table）
    #    文本框多但每框字少，用段落数会误报；字符数才反映真实信息密度。
    #    阈值：单页正文 > 600 字（约 12-15 条长 bullet）才算臃肿。
    for i, slide in enumerate(slides):
        total_chars = sum(
            len(p.text) for s in slide.shapes if s.has_text_frame
            for p in s.text_frame.paragraphs
        )
        if total_chars > 600:
            warnings.append(f"第{i+1}页文字偏多({total_chars}字)，建议拆分")

    # 5. 字号层级
    max_font = 0
    min_font = 999
    for slide in slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size:
                        pt_val = run.font.size.pt
                        max_font = max(max_font, pt_val)
                        min_font = min(min_font, pt_val)
    if max_font > 0 and min_font < 999:
        ratio = max_font / max(min_font, 1)
        checks.append(("字号层级跳跃", ratio >= 2.5))
        if ratio < 2.5:
            warnings.append(f"字号层级不够跳跃: 最大{max_font}pt/最小{min_font}pt={ratio:.1f}x (需≥2.5x)")
            passed = False
    else:
        warnings.append("无法检测字号层级（可能未设置font.size）")

    # 6. 配色数量
    colors = set()
    for slide in slides:
        for shape in slide.shapes:
            try:
                if hasattr(shape, 'fill') and shape.fill.type is not None:
                    if shape.fill.fore_color and shape.fill.fore_color.rgb:
                        colors.add(str(shape.fill.fore_color.rgb))
            except Exception:
                pass
    checks.append(("配色克制", len(colors) <= 10))
    if len(colors) > 10:
        warnings.append(f"使用颜色过多: {len(colors)}种 (建议≤10)")
        passed = False

    # 汇总
    for _name, result in checks:
        if result is False:
            passed = False

    return {"passed": passed, "checks": checks, "warnings": warnings, "total_slides": n_slides}
