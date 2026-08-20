"""Less-template-like layout families for the PPTX generation engine."""

from __future__ import annotations

import math
import sys

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt
from pptx_helper import (
    FONT_CN,
    FONT_EN,
    _add_para,
    _blank_layout,
    _fit_image_in_box,
    _rect,
    _set_bg,
    _textbox,
)


def _new_slide(prs, theme):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _set_bg(slide, theme["bg"])
    return slide


def _folio(slide, theme, ctx):
    page = ctx.get("page_number")
    if page is None:
        return
    box = _textbox(slide, 11.9, 6.9, 0.5, 0.24)
    _add_para(
        box, f"{int(page):02d}", first=True, size=Pt(9),
        color=theme["text_muted"], name=FONT_EN, align=PP_ALIGN.RIGHT,
    )


def _header(slide, theme, ctx, compact=False):
    kicker = str(ctx.get("kicker") or "").strip()
    title = str(ctx.get("title") or "")
    if kicker:
        kicker_box = _textbox(slide, 0.9, 0.48, 2.4, 0.28)
        _add_para(
            kicker_box, kicker, first=True, size=Pt(9), color=theme["accent"],
            bold=True, name=FONT_EN,
        )
    title_y = 0.73 if kicker else 0.52
    title_box = _textbox(slide, 0.9, title_y, 10.9, 0.62)
    _add_para(
        title_box, title, first=True, size=Pt(26 if compact else 31),
        color=theme["text"], bold=True, name=FONT_CN,
    )
    _rect(slide, 0.9, 1.43, 11.55, 0.012, fill=theme["text_muted"])
    _rect(slide, 0.9, 1.40, 1.2, 0.04, fill=theme["accent"])
    _folio(slide, theme, ctx)


_TOC_MAX_ITEMS = 12


def _toc_items(ctx):
    items = list(ctx.get("toc_items") or ctx.get("bullets") or [])
    if len(items) > _TOC_MAX_ITEMS:
        print(
            f"  [警告] 目录版式最多支持{_TOC_MAX_ITEMS}项，当前{len(items)}项，超出的条目被截断",
            file=sys.stderr,
        )
    return items[:_TOC_MAX_ITEMS]


def layout_cover_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    image = ctx.get("cover_image")
    if image:
        _fit_image_in_box(slide, image, 8.45, 0, 4.883, 7.5, cover=True)
        panel_color = theme["bg"]
        _rect(slide, 0, 0, 8.7, 7.5, fill=panel_color)
    _rect(slide, 0.92, 0.75, 1.5, 0.045, fill=theme["accent"])
    title = _textbox(slide, 0.92, 1.35, 7.8 if image else 10.8, 2.2, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(
        title, ctx.get("title", ""), first=True, size=Pt(48),
        color=theme["text"], bold=True, name=FONT_CN,
    )
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        sub = _textbox(slide, 0.92, 4.2, 7.2 if image else 9.8, 1.0)
        _add_para(
            sub, subtitle, first=True, size=Pt(17), color=theme["text_muted"],
            name=FONT_CN,
        )
    _rect(slide, 0.92, 6.72, 3.1, 0.012, fill=theme["text_muted"])
    meta = _textbox(slide, 0.92, 6.82, 3.1, 0.26)
    _add_para(meta, "", first=True, size=Pt(9), color=theme["text_muted"], name=FONT_EN)
    return slide


def layout_toc_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    title = _textbox(slide, 0.9, 0.75, 4.2, 1.2)
    _add_para(
        title, ctx.get("title", "目录"), first=True, size=Pt(42),
        color=theme["text"], bold=True, name=FONT_CN,
    )
    kicker = str(ctx.get("kicker") or "").strip()
    if kicker:
        small = _textbox(slide, 0.92, 0.45, 2.5, 0.24)
        _add_para(small, kicker, first=True, size=Pt(9), color=theme["accent"], bold=True, name=FONT_EN)
    items = _toc_items(ctx)
    if len(items) <= 6:
        for index, item in enumerate(items):
            y = 1.0 + index * 0.82
            number = _textbox(slide, 5.35, y, 0.6, 0.34)
            _add_para(number, f"{index + 1:02d}", first=True, size=Pt(13), color=theme["accent"], bold=True, name=FONT_EN)
            text = _textbox(slide, 6.15, y - 0.02, 5.25, 0.42)
            _add_para(text, str(item), first=True, size=Pt(18), color=theme["text"], name=FONT_CN)
            _rect(slide, 5.35, y + 0.52, 6.0, 0.01, fill=theme["text_muted"])
    else:
        rows = math.ceil(len(items) / 2)
        col_w = 3.55
        row_h = min(0.9, 5.55 / rows)
        for index, item in enumerate(items):
            col = index // rows
            row = index % rows
            x = 5.15 + col * col_w
            y = 1.08 + row * row_h
            number = _textbox(slide, x, y, 0.45, 0.3)
            _add_para(number, f"{index + 1:02d}", first=True, size=Pt(11), color=theme["accent"], bold=True, name=FONT_EN)
            text = _textbox(slide, x + 0.52, y - 0.02, col_w - 0.65, 0.4)
            _add_para(text, str(item), first=True, size=Pt(14), color=theme["text"], name=FONT_CN)
            _rect(slide, x + 0.52, y + row_h - 0.2, col_w - 0.7, 0.01, fill=theme["text_muted"])
    _folio(slide, theme, ctx)
    return slide


def layout_section_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    number = str(ctx.get("section_number") or "01")
    number_box = _textbox(slide, 0.9, 0.75, 2.5, 1.3)
    _add_para(number_box, number, first=True, size=Pt(70), color=theme["accent"], bold=False, name=FONT_EN)
    _rect(slide, 0.9, 2.4, 2.0, 0.04, fill=theme["accent"])
    title = _textbox(slide, 4.15, 1.45, 7.6, 2.0, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(title, ctx.get("title", ""), first=True, size=Pt(42), color=theme["text"], bold=True, name=FONT_CN)
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        box = _textbox(slide, 4.18, 3.7, 6.8, 0.8)
        _add_para(box, subtitle, first=True, size=Pt(16), color=theme["text_muted"], name=FONT_CN)
    _folio(slide, theme, ctx)
    return slide


def layout_bullets_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx)
    bullets = ctx.get("bullets") or []
    top = 1.78
    gap = min(0.88, 4.85 / max(1, len(bullets)))
    for index, item in enumerate(bullets[:7]):
        y = top + index * gap
        number = _textbox(slide, 0.92, y, 0.48, 0.34)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN)
        line = _textbox(slide, 1.65, y - 0.03, 10.2, gap * 0.75)
        _add_para(line, str(item), first=True, size=Pt(18), color=theme["text"], name=FONT_CN)
        _rect(slide, 1.65, y + gap * 0.7, 10.2, 0.008, fill=theme["text_muted"])
    return slide


def layout_text_image_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    images = ctx.get("images") or []
    if images:
        _fit_image_in_box(slide, images[0], 7.55, 1.8, 4.85, 4.8, cover=True)
    bullets = ctx.get("bullets") or []
    top = 1.9
    bottom = 6.72
    text_w = 5.45
    line_h = 0.30
    gap = 0.14

    def _est_lines(text):
        # 贪心折行模拟：CJK 逐字可断（0.229"/字，16pt 全角加安全余量）；
        # 连续 ASCII 为不可断词：大写/数字/下划线 0.151"，其余 0.115"。
        # 扣除 python-pptx 文本框默认内边距（左右各 0.1"）。
        eff_w = text_w - 0.2
        units = []
        buf = 0.0
        for ch in str(text):
            if ord(ch) > 0x2E7F:
                if buf:
                    units.append(buf)
                    buf = 0.0
                units.append(0.229)
            else:
                buf += 0.151 if (ch.isupper() or ch.isdigit() or ch == "_") else 0.115
        if buf:
            units.append(buf)
        lines, cur = 1, 0.0
        for u in units:
            if cur > 0 and cur + u > eff_w:
                lines += 1
                cur = 0.0
            cur += u
        return lines

    y = top
    for index, item in enumerate(bullets[:6]):
        lines = _est_lines(item)
        item_h = lines * line_h
        if index > 0 and y + item_h > bottom:
            break  # 容量不足时停止堆叠，避免条目互相重叠（由 QA 报告内容密度）
        number = _textbox(slide, 0.92, y, 0.45, 0.3)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(9), color=theme["accent"], bold=True, name=FONT_EN)
        text = _textbox(slide, 1.55, y - 0.02, text_w, item_h + 0.1)
        _add_para(text, str(item), first=True, size=Pt(16), color=theme["text"], name=FONT_CN)
        y += item_h + gap
    _rect(slide, 7.33, 1.8, 0.025, 4.8, fill=theme["accent"])
    return slide


def layout_full_image_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    images = ctx.get("images") or []
    if images:
        _fit_image_in_box(slide, images[0], 0, 0, 13.333, 7.5, cover=True)
    _rect(slide, 0, 0, 5.25, 7.5, fill=theme["dark"])
    _rect(slide, 0.9, 0.85, 1.35, 0.045, fill=theme["accent"])
    title = _textbox(slide, 0.9, 1.55, 3.8, 2.2, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(title, ctx.get("title", ""), first=True, size=Pt(38), color=theme["white"], bold=True, name=FONT_CN)
    _folio(slide, theme, ctx)
    return slide


def layout_image_grid_editorial(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    images = ctx.get("images") or []
    if len(images) > 3:
        print(f"  [警告] image_grid_editorial 版式最多支持3张图，当前{len(images)}张，超出的图片被截断", file=__import__('sys').stderr)
    slots = [
        (0.9, 1.8, 7.2, 4.9),
        (8.3, 1.8, 4.1, 2.33),
        (8.3, 4.37, 4.1, 2.33),
    ]
    for image, slot in zip(images[:3], slots, strict=False):
        _fit_image_in_box(slide, image, *slot, cover=True)
    return slide


def layout_dashboard_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    metrics = (ctx.get("metrics") or [])[:4]
    n = max(1, len(metrics))
    left = 0.9
    width = 11.55 / n
    for index, metric in enumerate(metrics):
        x = left + index * width
        if index:
            _rect(slide, x, 1.85, 0.012, 1.45, fill=theme["text_muted"])
        label = _textbox(slide, x + 0.18, 1.87, width - 0.35, 0.3)
        _add_para(label, metric.get("label", ""), first=True, size=Pt(10), color=theme["text_muted"], name=FONT_CN)
        value = _textbox(slide, x + 0.18, 2.18, width - 0.35, 0.72)
        _add_para(value, metric.get("value", "--"), first=True, size=Pt(32), color=theme["text"], bold=True, name=FONT_EN)
        change = _textbox(slide, x + 0.18, 2.94, width - 0.35, 0.25)
        _add_para(change, metric.get("change", ""), first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN)
    _rect(slide, 0.9, 3.48, 11.55, 0.012, fill=theme["text_muted"])
    values = []
    for metric in metrics:
        raw = str(metric.get("change", ""))
        try:
            values.append(abs(float(raw.replace("+", "").replace("%", "").replace("pp", ""))))
        except ValueError:
            values.append(1)
    maximum = max(values, default=1) or 1
    bar_left = 2.4
    bar_width = 8.6
    for index, (metric, value) in enumerate(zip(metrics, values, strict=False)):
        y = 3.95 + index * 0.62
        label = _textbox(slide, 0.92, y - 0.02, 1.3, 0.28)
        _add_para(label, metric.get("label", ""), first=True, size=Pt(10), color=theme["text_muted"], name=FONT_CN)
        _rect(slide, bar_left, y, bar_width, 0.09, fill=theme["bg_alt"])
        _rect(slide, bar_left, y, bar_width * value / maximum, 0.09, fill=theme["accent"])
        number = _textbox(slide, 11.25, y - 0.07, 1.15, 0.25)
        _add_para(number, metric.get("change", ""), first=True, size=Pt(9), color=theme["text"], name=FONT_EN, align=PP_ALIGN.RIGHT)
    return slide


def layout_process_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    steps = (ctx.get("steps") or [])[:6]
    n = max(1, len(steps))
    start_x, end_x = 1.05, 12.1
    y = 3.65
    _rect(slide, start_x, y, end_x - start_x, 0.018, fill=theme["text_muted"])
    for index, step in enumerate(steps):
        x = start_x + (end_x - start_x) * (index / max(1, n - 1))
        _rect(slide, x - 0.05, y - 0.05, 0.1, 0.1, fill=theme["accent"])
        number_y = 2.35 if index % 2 == 0 else 3.95
        number = _textbox(slide, x - 0.42, number_y, 0.85, 0.28)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN, align=PP_ALIGN.CENTER)
        label = _textbox(slide, x - 0.78, number_y + 0.35, 1.56, 0.75)
        _add_para(label, str(step), first=True, size=Pt(15), color=theme["text"], name=FONT_CN, align=PP_ALIGN.CENTER)
    return slide


def _column_bullets(slide, theme, x, title, bullets, marker):
    mark = _textbox(slide, x, 1.95, 0.55, 0.28)
    _add_para(mark, marker, first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN)
    head = _textbox(slide, x, 2.28, 4.7, 0.62)
    _add_para(head, title, first=True, size=Pt(25), color=theme["text"], bold=True, name=FONT_CN)
    _rect(slide, x, 3.03, 4.75, 0.025, fill=theme["accent"])
    for index, item in enumerate((bullets or [])[:5]):
        y = 3.42 + index * 0.62
        number = _textbox(slide, x, y, 0.38, 0.28)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(9), color=theme["text_muted"], name=FONT_EN)
        text = _textbox(slide, x + 0.58, y - 0.02, 4.0, 0.4)
        _add_para(text, str(item), first=True, size=Pt(16), color=theme["text"], name=FONT_CN)


def layout_comparison_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    left = ctx.get("left") or {}
    right = ctx.get("right") or {}
    _column_bullets(slide, theme, 0.92, left.get("title", "A"), left.get("bullets", []), "A")
    _column_bullets(slide, theme, 6.82, right.get("title", "B"), right.get("bullets", []), "B")
    return slide


def layout_timeline_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    events = (ctx.get("events") or [])[:6]
    n = max(1, len(events))
    start_x, end_x, y = 1.1, 12.1, 3.75
    _rect(slide, start_x, y, end_x - start_x, 0.02, fill=theme["text_muted"])
    for index, event in enumerate(events):
        x = start_x + (end_x - start_x) * (index / max(1, n - 1))
        _rect(slide, x - 0.045, y - 0.045, 0.09, 0.09, fill=theme["accent"])
        above = index % 2 == 0
        date_y = 2.47 if above else 4.12
        date = _textbox(slide, x - 0.75, date_y, 1.5, 0.28)
        _add_para(date, event.get("date", ""), first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN, align=PP_ALIGN.CENTER)
        label = _textbox(slide, x - 0.9, date_y + 0.38, 1.8, 0.72)
        _add_para(label, event.get("title", ""), first=True, size=Pt(14), color=theme["text"], name=FONT_CN, align=PP_ALIGN.CENTER)
    return slide


def layout_quote_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _rect(slide, 0.92, 0.82, 1.45, 0.045, fill=theme["accent"])
    quote = _textbox(slide, 1.6, 1.65, 9.9, 3.2, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(quote, ctx.get("quote", ""), first=True, size=Pt(31), color=theme["text"], name=FONT_CN)
    _rect(slide, 1.6, 5.15, 2.1, 0.018, fill=theme["text_muted"])
    source = str(ctx.get("source") or "").strip()
    if source:
        box = _textbox(slide, 1.6, 5.38, 5.8, 0.36)
        _add_para(box, source, first=True, size=Pt(12), color=theme["accent"], name=FONT_CN)
    _folio(slide, theme, ctx)
    return slide


def layout_end_grid(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _rect(slide, 0.92, 0.88, 1.45, 0.045, fill=theme["accent"])
    title = _textbox(slide, 0.92, 2.25, 7.2, 1.35)
    _add_para(title, ctx.get("title", "谢谢"), first=True, size=Pt(44), color=theme["text"], bold=True, name=FONT_CN)
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        box = _textbox(slide, 0.95, 4.15, 7.8, 0.7)
        _add_para(box, subtitle, first=True, size=Pt(15), color=theme["text_muted"], name=FONT_CN)
    _rect(slide, 9.8, 0.75, 0.018, 5.85, fill=theme["text_muted"])
    return slide


def _technical_header(slide, theme, ctx):
    page = int(ctx.get("page_number") or 0)
    index = _textbox(slide, 0.82, 0.5, 1.15, 0.58)
    _add_para(index, f"{page:02d}", first=True, size=Pt(25), color=theme["accent"], name=FONT_EN)
    _rect(slide, 2.05, 0.48, 0.018, 0.72, fill=theme["text_muted"])
    title = _textbox(slide, 2.35, 0.48, 9.7, 0.64)
    _add_para(title, ctx.get("title", ""), first=True, size=Pt(29), color=theme["text"], bold=True, name=FONT_CN)
    _rect(slide, 0.82, 1.34, 11.7, 0.012, fill=theme["text_muted"])


def layout_cover_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    image = ctx.get("cover_image")
    if image:
        _fit_image_in_box(slide, image, 8.55, 0.0, 4.783, 7.5, cover=True)
        _rect(slide, 0, 0, 8.75, 7.5, fill=theme["bg"])
    _rect(slide, 0.82, 0.65, 0.028, 5.95, fill=theme["accent"])
    code = _textbox(slide, 1.2, 0.72, 3.2, 0.28)
    _add_para(code, "01", first=True, size=Pt(9), color=theme["text_muted"], name=FONT_EN)
    title = _textbox(slide, 1.2, 1.62, 6.75 if image else 10.5, 2.0, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(title, ctx.get("title", ""), first=True, size=Pt(46), color=theme["text"], bold=True, name=FONT_CN)
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        box = _textbox(slide, 1.22, 4.55, 6.25 if image else 9.0, 1.1)
        _add_para(box, subtitle, first=True, size=Pt(15), color=theme["text_muted"], name=FONT_CN)
    _rect(slide, 1.2, 6.55, 6.1, 0.012, fill=theme["text_muted"])
    return slide


def layout_toc_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _technical_header(slide, theme, ctx)
    items = _toc_items(ctx)
    cols = 3 if len(items) <= 9 else 4
    rows = max(1, math.ceil(len(items) / cols))
    col_w = 11.65 / cols
    row_h = 4.75 / rows
    font_size = 20 if len(items) <= 6 else 16 if len(items) <= 9 else 14
    for index, item in enumerate(items):
        row = index // cols
        col = index % cols
        x = 0.85 + col * col_w
        y = 1.75 + row * row_h
        number = _textbox(slide, x, y, 0.7, 0.32)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(11), color=theme["accent"], bold=True, name=FONT_EN)
        _rect(slide, x, y + 0.42, col_w - 0.45, 0.018, fill=theme["text_muted"])
        text = _textbox(slide, x, y + 0.62, col_w - 0.55, min(0.72, row_h - 0.68))
        _add_para(text, str(item), first=True, size=Pt(font_size), color=theme["text"], name=FONT_CN)
    return slide


def layout_bullets_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _technical_header(slide, theme, ctx)
    bullets = (ctx.get("bullets") or [])[:7]
    gap = min(0.68, 4.65 / max(1, len(bullets)))
    for index, item in enumerate(bullets):
        y = 1.72 + index * gap
        number = _textbox(slide, 0.86, y, 0.75, 0.32)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN)
        _rect(slide, 1.72, y + 0.13, 0.62, 0.012, fill=theme["text_muted"])
        text = _textbox(slide, 2.6, y - 0.04, 9.45, 0.46)
        _add_para(text, str(item), first=True, size=Pt(17), color=theme["text"], name=FONT_CN)
    return slide


def layout_dashboard_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _technical_header(slide, theme, ctx)
    metrics = (ctx.get("metrics") or [])[:4]
    lead = metrics[0] if metrics else {"label": "指标", "value": "--", "change": ""}
    label = _textbox(slide, 0.9, 1.92, 3.5, 0.35)
    _add_para(label, lead.get("label", ""), first=True, size=Pt(12), color=theme["text_muted"], name=FONT_CN)
    value = _textbox(slide, 0.88, 2.35, 4.2, 1.4)
    _add_para(value, lead.get("value", "--"), first=True, size=Pt(58), color=theme["text"], bold=True, name=FONT_EN)
    change = _textbox(slide, 0.92, 4.05, 3.8, 0.36)
    _add_para(change, lead.get("change", ""), first=True, size=Pt(12), color=theme["accent"], bold=True, name=FONT_EN)
    _rect(slide, 5.0, 1.76, 0.02, 4.65, fill=theme["text_muted"])
    rest = metrics[1:] or metrics[:1]
    for index, metric in enumerate(rest[:3]):
        y = 1.85 + index * 1.48
        _rect(slide, 5.42, y + 1.12, 6.75, 0.012, fill=theme["text_muted"])
        row_label = _textbox(slide, 5.45, y, 2.1, 0.34)
        _add_para(row_label, metric.get("label", ""), first=True, size=Pt(11), color=theme["text_muted"], name=FONT_CN)
        row_value = _textbox(slide, 7.55, y - 0.08, 2.55, 0.62)
        _add_para(row_value, metric.get("value", "--"), first=True, size=Pt(29), color=theme["text"], bold=True, name=FONT_EN)
        row_change = _textbox(slide, 10.25, y + 0.03, 1.85, 0.34)
        _add_para(row_change, metric.get("change", ""), first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN, align=PP_ALIGN.RIGHT)
    return slide


def layout_process_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _technical_header(slide, theme, ctx)
    steps = (ctx.get("steps") or [])[:6]
    _rect(slide, 1.25, 1.75, 0.02, 4.7, fill=theme["text_muted"])
    gap = 4.6 / max(1, len(steps))
    for index, step in enumerate(steps):
        y = 1.8 + index * gap
        _rect(slide, 1.17, y + 0.08, 0.18, 0.18, fill=theme["accent"])
        number = _textbox(slide, 1.75, y, 0.75, 0.35)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(10), color=theme["accent"], bold=True, name=FONT_EN)
        label = _textbox(slide, 2.75, y - 0.04, 8.75, 0.48)
        _add_para(label, str(step), first=True, size=Pt(18), color=theme["text"], name=FONT_CN)
        _rect(slide, 2.75, y + 0.55, 9.3, 0.008, fill=theme["text_muted"])
    return slide


def layout_timeline_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _technical_header(slide, theme, ctx)
    events = (ctx.get("events") or [])[:6]
    _rect(slide, 3.0, 1.7, 0.02, 4.85, fill=theme["text_muted"])
    gap = 4.65 / max(1, len(events))
    for index, event in enumerate(events):
        y = 1.78 + index * gap
        date = _textbox(slide, 0.88, y, 1.65, 0.32)
        _add_para(date, event.get("date", ""), first=True, size=Pt(11), color=theme["accent"], bold=True, name=FONT_EN, align=PP_ALIGN.RIGHT)
        _rect(slide, 2.91, y + 0.08, 0.2, 0.2, fill=theme["accent"])
        title = _textbox(slide, 3.55, y - 0.03, 8.35, 0.48)
        _add_para(title, event.get("title", ""), first=True, size=Pt(18), color=theme["text"], name=FONT_CN)
    return slide


def layout_end_technical(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _rect(slide, 0.82, 0.65, 0.028, 5.95, fill=theme["accent"])
    code = _textbox(slide, 1.22, 0.72, 3.0, 0.28)
    _add_para(code, "00", first=True, size=Pt(9), color=theme["text_muted"], name=FONT_EN)
    title = _textbox(slide, 1.2, 2.35, 7.7, 1.2)
    _add_para(title, ctx.get("title", "谢谢"), first=True, size=Pt(48), color=theme["text"], bold=True, name=FONT_CN)
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        box = _textbox(slide, 1.22, 4.45, 8.6, 0.7)
        _add_para(box, subtitle, first=True, size=Pt(15), color=theme["text_muted"], name=FONT_CN)
    _rect(slide, 1.2, 6.55, 10.9, 0.012, fill=theme["text_muted"])
    return slide


def layout_cover_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    image = ctx.get("cover_image")
    if image:
        _fit_image_in_box(slide, image, 7.8, 0, 5.533, 7.5, cover=True)
        _rect(slide, 0, 0, 8.05, 7.5, fill=theme["bg"])
    _rect(slide, 0, 0, 0.28, 7.5, fill=theme["accent"])
    title = _textbox(slide, 0.8, 0.88, 6.65 if image else 11.7, 3.35, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(title, ctx.get("title", ""), first=True, size=Pt(62), color=theme["text"], bold=True, name=FONT_CN)
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        box = _textbox(slide, 0.85, 5.35, 6.35 if image else 9.7, 1.0)
        _add_para(box, subtitle, first=True, size=Pt(14), color=theme["text_muted"], name=FONT_CN)
    _rect(slide, 0.85, 6.65, 2.2, 0.05, fill=theme["accent"])
    return slide


def layout_toc_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    title = _textbox(slide, 0.85, 0.55, 5.8, 0.85)
    _add_para(title, ctx.get("title", "目录"), first=True, size=Pt(36), color=theme["text"], bold=True, name=FONT_CN)
    items = _toc_items(ctx)
    cols = 2 if len(items) <= 6 else 3
    rows = max(1, math.ceil(len(items) / cols))
    col_w = 11.65 / cols
    row_h = 5.15 / rows
    number_size = 30 if cols == 2 else 23
    text_size = 18 if cols == 2 else 15
    for index, item in enumerate(items):
        col = index % cols
        row = index // cols
        x = 0.85 + col * col_w
        y = 1.55 + row * row_h
        number = _textbox(slide, x, y, 0.9, 0.62)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(number_size), color=theme["accent"], bold=True, name=FONT_EN)
        text = _textbox(slide, x + 1.0, y + 0.1, col_w - 1.25, 0.55)
        _add_para(text, str(item), first=True, size=Pt(text_size), color=theme["text"], name=FONT_CN)
        _rect(slide, x, y + row_h - 0.28, col_w - 0.35, 0.012, fill=theme["text_muted"])
    return slide


def layout_bullets_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx)
    bullets = (ctx.get("bullets") or [])[:6]
    for index, item in enumerate(bullets):
        col = index % 2
        row = index // 2
        x = 0.92 + col * 6.0
        y = 1.82 + row * 1.48
        number = _textbox(slide, x, y, 0.95, 0.55)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(23), color=theme["accent"], bold=True, name=FONT_EN)
        text = _textbox(slide, x + 1.15, y + 0.08, 4.55, 0.65)
        _add_para(text, str(item), first=True, size=Pt(17), color=theme["text"], name=FONT_CN)
    return slide


def layout_dashboard_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    metrics = (ctx.get("metrics") or [])[:4]
    lead = metrics[0] if metrics else {"label": "指标", "value": "--", "change": ""}
    lead_value = _textbox(slide, 0.82, 1.75, 6.15, 1.75)
    _add_para(lead_value, lead.get("value", "--"), first=True, size=Pt(72), color=theme["text"], bold=True, name=FONT_EN)
    lead_label = _textbox(slide, 0.92, 3.55, 5.6, 0.45)
    _add_para(lead_label, f"{lead.get('label', '')}  {lead.get('change', '')}".strip(), first=True, size=Pt(14), color=theme["accent"], bold=True, name=FONT_CN)
    _rect(slide, 6.65, 1.72, 0.025, 4.75, fill=theme["accent"])
    for index, metric in enumerate(metrics[1:4]):
        y = 1.82 + index * 1.45
        label = _textbox(slide, 7.15, y, 2.3, 0.34)
        _add_para(label, metric.get("label", ""), first=True, size=Pt(11), color=theme["text_muted"], name=FONT_CN)
        value = _textbox(slide, 9.25, y - 0.12, 2.0, 0.72)
        _add_para(value, metric.get("value", "--"), first=True, size=Pt(31), color=theme["text"], bold=True, name=FONT_EN)
        change = _textbox(slide, 11.25, y + 0.04, 1.0, 0.3)
        _add_para(change, metric.get("change", ""), first=True, size=Pt(9), color=theme["accent"], bold=True, name=FONT_EN, align=PP_ALIGN.RIGHT)
        _rect(slide, 7.15, y + 0.95, 5.1, 0.012, fill=theme["text_muted"])
    return slide


def layout_process_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    steps = (ctx.get("steps") or [])[:5]
    n = max(1, len(steps))
    width = 11.6 / n
    for index, step in enumerate(steps):
        x = 0.82 + index * width
        number = _textbox(slide, x, 2.0, width - 0.18, 1.05)
        _add_para(number, f"{index + 1:02d}", first=True, size=Pt(42), color=theme["accent"], bold=True, name=FONT_EN)
        _rect(slide, x, 3.22, width - 0.25, 0.025, fill=theme["text_muted"])
        label = _textbox(slide, x, 3.65, width - 0.3, 1.0)
        _add_para(label, str(step), first=True, size=Pt(17), color=theme["text"], name=FONT_CN)
    return slide


def layout_timeline_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _header(slide, theme, ctx, compact=True)
    events = (ctx.get("events") or [])[:4]
    for index, event in enumerate(events):
        col = index % 2
        row = index // 2
        x = 0.9 + col * 6.0
        y = 1.78 + row * 2.25
        date = _textbox(slide, x, y, 2.6, 0.5)
        _add_para(date, event.get("date", ""), first=True, size=Pt(20), color=theme["accent"], bold=True, name=FONT_EN)
        _rect(slide, x, y + 0.68, 5.55, 0.025, fill=theme["accent"])
        title = _textbox(slide, x, y + 1.02, 5.15, 0.72)
        _add_para(title, event.get("title", ""), first=True, size=Pt(20), color=theme["text"], name=FONT_CN)
    return slide


def layout_quote_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _rect(slide, 0, 0, 0.28, 7.5, fill=theme["accent"])
    quote = _textbox(slide, 0.9, 1.0, 11.35, 4.4, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(quote, ctx.get("quote", ""), first=True, size=Pt(38), color=theme["text"], bold=True, name=FONT_CN)
    source = str(ctx.get("source") or "").strip()
    if source:
        box = _textbox(slide, 0.95, 6.1, 5.5, 0.36)
        _add_para(box, source, first=True, size=Pt(11), color=theme["accent"], name=FONT_CN)
    return slide


def layout_end_poster(prs, theme, ctx):
    slide = _new_slide(prs, theme)
    _rect(slide, 0, 0, 13.333, 0.22, fill=theme["accent"])
    title = _textbox(slide, 0.82, 1.65, 11.7, 2.0, anchor=MSO_ANCHOR.MIDDLE)
    _add_para(title, ctx.get("title", "谢谢"), first=True, size=Pt(68), color=theme["text"], bold=True, name=FONT_CN, align=PP_ALIGN.CENTER)
    subtitle = str(ctx.get("subtitle") or "").strip()
    if subtitle:
        box = _textbox(slide, 2.0, 5.25, 9.3, 0.7)
        _add_para(box, subtitle, first=True, size=Pt(15), color=theme["text_muted"], name=FONT_CN, align=PP_ALIGN.CENTER)
    return slide


EDITORIAL_GRID = {
    "cover": layout_cover_grid,
    "toc": layout_toc_grid,
    "section": layout_section_grid,
    "bullets": layout_bullets_grid,
    "text_image": layout_text_image_grid,
    "full_image": layout_full_image_grid,
    "image_grid": layout_image_grid_editorial,
    "dashboard": layout_dashboard_grid,
    "timeline": layout_timeline_grid,
    "comparison": layout_comparison_grid,
    "quote": layout_quote_grid,
    "process": layout_process_grid,
    "end": layout_end_grid,
}


TECHNICAL_AXIS = {
    **EDITORIAL_GRID,
    "cover": layout_cover_technical,
    "toc": layout_toc_technical,
    "bullets": layout_bullets_technical,
    "dashboard": layout_dashboard_technical,
    "process": layout_process_technical,
    "timeline": layout_timeline_technical,
    "end": layout_end_technical,
}


POSTER_COLUMN = {
    **EDITORIAL_GRID,
    "cover": layout_cover_poster,
    "toc": layout_toc_poster,
    "bullets": layout_bullets_poster,
    "dashboard": layout_dashboard_poster,
    "process": layout_process_poster,
    "timeline": layout_timeline_poster,
    "quote": layout_quote_poster,
    "end": layout_end_poster,
}


FAMILIES = {
    "editorial_grid": EDITORIAL_GRID,
    "technical_axis": TECHNICAL_AXIS,
    "poster_column": POSTER_COLUMN,
}


def get_layout_function(family: str | None, layout_name: str):
    if not family:
        return None
    return FAMILIES.get(family, {}).get(layout_name)
