"""Chart editing — modify existing chart data, style, type, and layout after creation.

python-pptx creates charts via ``slide.shapes.add_chart()`` but offers limited
post-creation editing.  This module fills the gap with:

- **Data editing**: change series values, categories, add/remove series
- **Style editing**: colors, markers, line widths, data labels, legend
- **Type switching**: convert between chart types (column ↔ bar ↔ line ↔ area)
- **Axis control**: titles, ranges, number formats, gridlines
- **Layout control**: chart area size, plot area margins, legend position

All functions accept both ``Presentation`` objects and file paths.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field

from pptx_skill._io import is_presentation as _is_presentation
from pptx_skill._io import open_prs as _open_prs
from pptx_skill._io import save_prs as _save_prs_impl
from pptx_skill.constants import A_NS as _NS_A

__all__ = [
    "ChartInfo",
    "SeriesInfo",
    "edit_chart_data",
    "add_chart_series",
    "remove_chart_series",
    "rename_chart_series",
    "set_chart_type",
    "set_chart_colors",
    "set_chart_style",
    "set_data_labels",
    "set_legend",
    "set_axis_title",
    "set_axis_range",
    "set_axis_number_format",
    "toggle_gridlines",
    "list_charts",
    "get_chart_info",
    "export_chart_data",
    "import_chart_data",
]

log = logging.getLogger(__name__)

_NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# Chart type mapping: python-pptx enum value → OOXML chart tag
_CHART_TYPE_MAP = {
    "column_clustered": "barChart",
    "column_stacked": "barChart",
    "column_100_stacked": "barChart",
    "bar_clustered": "barChart",
    "bar_stacked": "barChart",
    "bar_100_stacked": "barChart",
    "line": "lineChart",
    "line_stacked": "lineChart",
    "line_100_stacked": "lineChart",
    "line_markers": "lineChart",
    "area": "areaChart",
    "area_stacked": "areaChart",
    "area_100_stacked": "areaChart",
    "pie": "pieChart",
    "doughnut": "doughnutChart",
    "scatter": "scatterChart",
    "scatter_lines": "scatterChart",
    "bubble": "bubbleChart",
    "radar": "radarChart",
    "radar_filled": "radarChart",
    "radar_markers": "radarChart",
}


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class SeriesInfo:
    """Info about a chart series."""
    name: str = ""
    values: list[float] = field(default_factory=list)
    index: int = 0

    def __post_init__(self):
        if self.values is None:
            self.values = []


@dataclass
class ChartInfo:
    """Comprehensive info about a chart shape."""
    name: str = ""
    chart_type: str = ""
    series_count: int = 0
    category_count: int = 0
    categories: list[str] = field(default_factory=list)
    series: list[SeriesInfo] = field(default_factory=list)
    has_legend: bool = True
    has_data_labels: bool = False
    left: float = 0
    top: float = 0
    width: float = 0
    height: float = 0

    def __post_init__(self):
        if self.categories is None:
            self.categories = []
        if self.series is None:
            self.series = []


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _save_prs(prs, path):
    _save_prs_impl(prs, path, backup=False)


def _find_chart_shape(slide, shape_name: str):
    """Find a chart shape by name on a slide."""
    for shape in slide.shapes:
        if shape.name == shape_name:
            if shape.has_chart:
                return shape
    return None


def _get_chart_element(shape):
    """Get the <c:chart> element from a chart shape."""
    return shape.chart._chartSpace


def _find_or_create(element, tag: str, ns: str = _NS_C):
    """Find a child element or create it if not found."""
    child = element.find(f"{{{ns}}}{tag}")
    if child is None:
        from lxml import etree
        child = etree.SubElement(element, f"{{{ns}}}{tag}")
    return child


# ---------------------------------------------------------------------------
# List & info
# ---------------------------------------------------------------------------

def list_charts(prs_or_path, slide_index: int | None = None) -> list[dict]:
    """List all chart shapes in the presentation.

    Parameters
    ----------
    slide_index : int, optional
        1-based slide index (1 = first slide).  If provided, only list charts
        on that slide.

    Returns
    -------
    list[dict]
        Each dict has keys: slide_index, name, chart_type, series_count.
    """
    not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    try:
        results = []
        if slide_index is not None:
            if slide_index < 1 or slide_index > len(prs.slides):
                raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
            slides = [prs.slides[slide_index - 1]]
        else:
            slides = prs.slides
        for idx, slide in enumerate(slides):
            for shape in slide.shapes:
                try:
                    if shape.has_chart:
                        chart = shape.chart
                        results.append({
                            "slide_index": idx if slide_index is None else slide_index,
                            "name": shape.name,
                            "chart_type": str(chart.chart_type),
                            "series_count": len(chart.series),
                            "left": shape.left / 914400,
                            "top": shape.top / 914400,
                            "width": shape.width / 914400,
                            "height": shape.height / 914400,
                        })
                except Exception:
                    pass
        return results
    finally:
        pass


def get_chart_info(prs_or_path, slide_index: int, shape_name: str) -> ChartInfo | None:
    """Get detailed info about a specific chart.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return None

        chart = shape.chart
        info = ChartInfo(
            name=shape.name,
            chart_type=str(chart.chart_type),
            left=shape.left / 914400,
            top=shape.top / 914400,
            width=shape.width / 914400,
            height=shape.height / 914400,
        )

        # Series
        for i, series in enumerate(chart.series):
            try:
                vals = []
                for pt in series.values:
                    try:
                        vals.append(float(pt))
                    except (TypeError, ValueError):
                        vals.append(0.0)
                info.series.append(SeriesInfo(
                    name=str(series.format_code) if hasattr(series, "format_code") else f"Series {i+1}",
                    values=vals,
                    index=i,
                ))
            except Exception:
                info.series.append(SeriesInfo(name=f"Series {i+1}", index=i))

        info.series_count = len(info.series)

        # Categories
        try:
            plot = chart.plots[0]
            info.categories = [str(cat) for cat in plot.categories]
            info.category_count = len(info.categories)
        except Exception:
            pass

        # Legend
        try:
            info.has_legend = chart.has_legend
        except Exception:
            pass

        return info
    finally:
        pass


# ---------------------------------------------------------------------------
# Data editing
# ---------------------------------------------------------------------------

def edit_chart_data(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    series_index: int = 0,
    values: list[float] | None = None,
    categories: list[str] | None = None,
) -> bool:
    """Edit chart data — change series values and/or category labels.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    series_index : int
        0-based index of the series to edit.
    values : list[float], optional
        New values for the series.
    categories : list[str], optional
        New category labels (applies to all series).

    Returns
    -------
    bool
        True if the chart was modified.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart

        # Edit series values
        if values is not None:
            try:
                chart.series[series_index]
                # Use OOXML manipulation for precise control
                _set_series_values(chart, series_index, values)
            except (IndexError, Exception) as exc:
                log.warning("Failed to edit series %d: %s", series_index, exc)
                return False

        # Edit categories
        if categories is not None:
            _set_categories(chart, categories)

        return True
    finally:
        _save_prs(prs, path)


def _set_series_values(chart, series_index: int, values: list[float]):
    """Set values for a chart series via OOXML."""
    from lxml import etree

    chart_space = chart._chartSpace
    # Find the plot area
    plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
    if plot_area is None:
        return

    # Find the first chart type element
    chart_elem = None
    for child in plot_area:
        if child.tag.endswith("Chart") or child.tag.endswith("chart"):
            chart_elem = child
            break

    if chart_elem is None:
        return

    # Find ser elements
    ser_list = chart_elem.findall(f"{{{_NS_C}}}ser")
    if series_index >= len(ser_list):
        return

    ser = ser_list[series_index]

    # Update val element
    val = ser.find(f"{{{_NS_C}}}val")
    if val is None:
        val = etree.SubElement(ser, f"{{{_NS_C}}}val")

    # Clear existing
    for child in list(val):
        val.remove(child)

    # Add numRef
    numRef = etree.SubElement(val, f"{{{_NS_C}}}numRef")
    etree.SubElement(numRef, f"{{{_NS_C}}}f").text = f"Sheet1!$${chr(65 + series_index)}$2:${chr(65 + series_index)}${len(values) + 1}"

    numCache = etree.SubElement(numRef, f"{{{_NS_C}}}numCache")
    etree.SubElement(numCache, f"{{{_NS_C}}}formatCode").text = "General"

    ptCount = etree.SubElement(numCache, f"{{{_NS_C}}}ptCount")
    ptCount.set("val", str(len(values)))

    for i, v in enumerate(values):
        pt = etree.SubElement(numCache, f"{{{_NS_C}}}pt")
        pt.set("idx", str(i))
        etree.SubElement(pt, f"{{{_NS_C}}}v").text = str(v)


def _set_categories(chart, categories: list[str]):
    """Set category labels for a chart via OOXML."""
    from lxml import etree

    chart_space = chart._chartSpace
    plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
    if plot_area is None:
        return

    # Find first chart element
    chart_elem = None
    for child in plot_area:
        if child.tag.endswith("Chart") or child.tag.endswith("chart"):
            chart_elem = child
            break

    if chart_elem is None:
        return

    # Update categories in each series
    for ser in chart_elem.findall(f"{{{_NS_C}}}ser"):
        cat = ser.find(f"{{{_NS_C}}}cat")
        if cat is None:
            cat = etree.SubElement(ser, f"{{{_NS_C}}}cat")

        for child in list(cat):
            cat.remove(child)

        strRef = etree.SubElement(cat, f"{{{_NS_C}}}strRef")
        etree.SubElement(strRef, f"{{{_NS_C}}}f").text = f"Sheet1!$A$2:$A${len(categories) + 1}"

        strCache = etree.SubElement(strRef, f"{{{_NS_C}}}strCache")
        ptCount = etree.SubElement(strCache, f"{{{_NS_C}}}ptCount")
        ptCount.set("val", str(len(categories)))

        for i, c in enumerate(categories):
            pt = etree.SubElement(strCache, f"{{{_NS_C}}}pt")
            pt.set("idx", str(i))
            etree.SubElement(pt, f"{{{_NS_C}}}v").text = str(c)


def add_chart_series(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    name: str,
    values: list[float],
) -> int:
    """Add a new series to an existing chart.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).

    Returns
    -------
    int
        The new series index, or -1 on failure.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return -1

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return -1

        # Find first chart element
        chart_elem = None
        for child in plot_area:
            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                chart_elem = child
                break

        if chart_elem is None:
            return -1

        # Determine new series index
        existing_ser = chart_elem.findall(f"{{{_NS_C}}}ser")
        new_idx = len(existing_ser)

        # Create new ser element
        ser = etree.SubElement(chart_elem, f"{{{_NS_C}}}ser")

        # idx
        idx = etree.SubElement(ser, f"{{{_NS_C}}}idx")
        idx.set("val", str(new_idx))

        # order
        order = etree.SubElement(ser, f"{{{_NS_C}}}order")
        order.set("val", str(new_idx))

        # tx (series name)
        tx = etree.SubElement(ser, f"{{{_NS_C}}}tx")
        strRef = etree.SubElement(tx, f"{{{_NS_C}}}strRef")
        strCache = etree.SubElement(strRef, f"{{{_NS_C}}}strCache")
        ptCount = etree.SubElement(strCache, f"{{{_NS_C}}}ptCount")
        ptCount.set("val", "1")
        pt = etree.SubElement(strCache, f"{{{_NS_C}}}pt")
        pt.set("idx", "0")
        etree.SubElement(pt, f"{{{_NS_C}}}v").text = name

        # val
        val = etree.SubElement(ser, f"{{{_NS_C}}}val")
        numRef = etree.SubElement(val, f"{{{_NS_C}}}numRef")
        etree.SubElement(numRef, f"{{{_NS_C}}}f").text = f"Sheet1!${chr(65 + new_idx)}$2:${chr(65 + new_idx)}${len(values) + 1}"
        numCache = etree.SubElement(numRef, f"{{{_NS_C}}}numCache")
        etree.SubElement(numCache, f"{{{_NS_C}}}formatCode").text = "General"
        ptCount = etree.SubElement(numCache, f"{{{_NS_C}}}ptCount")
        ptCount.set("val", str(len(values)))
        for i, v in enumerate(values):
            pt_elem = etree.SubElement(numCache, f"{{{_NS_C}}}pt")
            pt_elem.set("idx", str(i))
            etree.SubElement(pt_elem, f"{{{_NS_C}}}v").text = str(v)

        return new_idx
    finally:
        _save_prs(prs, path)


def remove_chart_series(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    series_index: int,
) -> bool:
    """Remove a series from a chart by index.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        for child in plot_area:
            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                ser_list = child.findall(f"{{{_NS_C}}}ser")
                if series_index < len(ser_list):
                    child.remove(ser_list[series_index])
                    # Re-index remaining series
                    for i, ser in enumerate(child.findall(f"{{{_NS_C}}}ser")):
                        idx_elem = ser.find(f"{{{_NS_C}}}idx")
                        if idx_elem is not None:
                            idx_elem.set("val", str(i))
                        order_elem = ser.find(f"{{{_NS_C}}}order")
                        if order_elem is not None:
                            order_elem.set("val", str(i))
                    return True
        return False
    finally:
        _save_prs(prs, path)


def rename_chart_series(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    series_index: int,
    new_name: str,
) -> bool:
    """Rename a chart series.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        for child in plot_area:
            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                ser_list = child.findall(f"{{{_NS_C}}}ser")
                if series_index < len(ser_list):
                    ser = ser_list[series_index]
                    # Find or create tx/strRef
                    tx = ser.find(f"{{{_NS_C}}}tx")
                    if tx is None:
                        tx = etree.SubElement(ser, f"{{{_NS_C}}}tx")

                    # Clear existing and add new name
                    for old in list(tx):
                        tx.remove(old)

                    strRef = etree.SubElement(tx, f"{{{_NS_C}}}strRef")
                    strCache = etree.SubElement(strRef, f"{{{_NS_C}}}strCache")
                    ptCount = etree.SubElement(strCache, f"{{{_NS_C}}}ptCount")
                    ptCount.set("val", "1")
                    pt = etree.SubElement(strCache, f"{{{_NS_C}}}pt")
                    pt.set("idx", "0")
                    etree.SubElement(pt, f"{{{_NS_C}}}v").text = new_name
                    return True
        return False
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Style editing
# ---------------------------------------------------------------------------

def set_chart_colors(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    colors: list[str],
) -> bool:
    """Set series colors for a chart.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    colors : list[str]
        Hex color strings (e.g. ``["#FF0000", "#00FF00"]``), one per series.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        for child in plot_area:
            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                ser_list = child.findall(f"{{{_NS_C}}}ser")
                for i, ser in enumerate(ser_list):
                    if i >= len(colors):
                        break
                    color_hex = colors[i].lstrip("#")

                    # Find or create spPr
                    spPr = ser.find(f"{{{_NS_C}}}spPr")
                    if spPr is None:
                        spPr = etree.SubElement(ser, f"{{{_NS_C}}}spPr")

                    # Remove existing fill
                    for fill in list(spPr):
                        if fill.tag in (f"{{{_NS_A}}}solidFill", f"{{{_NS_A}}}noFill", f"{{{_NS_A}}}gradFill"):
                            spPr.remove(fill)

                    # Add solid fill
                    solidFill = etree.SubElement(spPr, f"{{{_NS_A}}}solidFill")
                    srgbClr = etree.SubElement(solidFill, f"{{{_NS_A}}}srgbClr")
                    srgbClr.set("val", color_hex)

        return True
    finally:
        _save_prs(prs, path)


def set_chart_style(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    style_index: int,
) -> bool:
    """Set the chart style (built-in style index 1-48).

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    style_index : int
        Built-in chart style index (1–48).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace

        # Find or create style element
        style = chart_space.find(f"{{{_NS_C}}}style")
        if style is None:
            style = etree.SubElement(chart_space, f"{{{_NS_C}}}style")

        style.set("val", str(max(1, min(48, style_index))))
        return True
    finally:
        _save_prs(prs, path)


def set_data_labels(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    show: bool = True,
    show_value: bool = True,
    show_category: bool = False,
    show_series_name: bool = False,
    show_percent: bool = False,
    number_format: str | None = None,
    font_size: int | None = None,
) -> bool:
    """Configure data labels on a chart.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        for child in plot_area:
            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                if show:
                    dLbls = child.find(f"{{{_NS_C}}}dLbls")
                    if dLbls is None:
                        dLbls = etree.SubElement(child, f"{{{_NS_C}}}dLbls")

                    # Set show properties
                    show_props = {
                        "showVal": str(show_value).lower(),
                        "showCatName": str(show_category).lower(),
                        "showSerName": str(show_series_name).lower(),
                        "showPercent": str(show_percent).lower(),
                    }
                    for attr, val in show_props.items():
                        elem = dLbls.find(f"{{{_NS_C}}}{attr}")
                        if elem is None:
                            elem = etree.SubElement(dLbls, f"{{{_NS_C}}}{attr}")
                        elem.set("val", val)

                    # Number format
                    if number_format is not None:
                        numFmt = dLbls.find(f"{{{_NS_C}}}numFmt")
                        if numFmt is None:
                            numFmt = etree.SubElement(dLbls, f"{{{_NS_C}}}numFmt")
                        numFmt.set("formatCode", number_format)
                        numFmt.set("sourceLinked", "0")

                    # Font size
                    if font_size is not None:
                        txPr = dLbls.find(f"{{{_NS_C}}}txPr")
                        if txPr is None:
                            txPr = etree.SubElement(dLbls, f"{{{_NS_C}}}txPr")
                        bodyPr = txPr.find(f"{{{_NS_A}}}bodyPr")
                        if bodyPr is None:
                            bodyPr = etree.SubElement(txPr, f"{{{_NS_A}}}bodyPr")
                        lstStyle = txPr.find(f"{{{_NS_A}}}lstStyle")
                        if lstStyle is None:
                            etree.SubElement(txPr, f"{{{_NS_A}}}lstStyle")
                        p = txPr.find(f"{{{_NS_A}}}p")
                        if p is None:
                            p = etree.SubElement(txPr, f"{{{_NS_A}}}p")
                        rPr = p.find(f"{{{_NS_A}}}rPr")
                        if rPr is None:
                            rPr = etree.SubElement(p, f"{{{_NS_A}}}rPr")
                        rPr.set("sz", str(font_size * 100))  # in hundredths of a point
                else:
                    # Remove data labels
                    dLbls = child.find(f"{{{_NS_C}}}dLbls")
                    if dLbls is not None:
                        child.remove(dLbls)

        return True
    finally:
        _save_prs(prs, path)


def set_legend(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    show: bool = True,
    position: str = "b",
    overlay: bool = False,
    font_size: int | None = None,
) -> bool:
    """Configure chart legend.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    position : str
        Legend position: "b" (bottom), "t" (top), "l" (left), "r" (right), "tr" (top-right).
    overlay : bool
        Whether the legend overlays the chart area.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace

        if show:
            legend = chart_space.find(f".//{{{_NS_C}}}legend")
            if legend is None:
                # Insert after plotArea
                plot_area = chart_space.find(f"{{{_NS_C}}}plotArea")
                legend = etree.Element(f"{{{_NS_C}}}legend")
                if plot_area is not None:
                    plot_area.addnext(legend)
                else:
                    chart_space.append(legend)

            # Position
            legendPos = legend.find(f"{{{_NS_C}}}legendPos")
            if legendPos is None:
                legendPos = etree.SubElement(legend, f"{{{_NS_C}}}legendPos")
            legendPos.set("val", position)

            # Overlay
            overlay_elem = legend.find(f"{{{_NS_C}}}overlay")
            if overlay_elem is None:
                overlay_elem = etree.SubElement(legend, f"{{{_NS_C}}}overlay")
            overlay_elem.set("val", str(overlay).lower())

            # Font size
            if font_size is not None:
                txPr = legend.find(f"{{{_NS_C}}}txPr")
                if txPr is None:
                    txPr = etree.SubElement(legend, f"{{{_NS_C}}}txPr")
                    etree.SubElement(txPr, f"{{{_NS_A}}}bodyPr")
                    etree.SubElement(txPr, f"{{{_NS_A}}}lstStyle")
                p = txPr.find(f"{{{_NS_A}}}p")
                if p is None:
                    p = etree.SubElement(txPr, f"{{{_NS_A}}}p")
                rPr = p.find(f"{{{_NS_A}}}rPr")
                if rPr is None:
                    rPr = etree.SubElement(p, f"{{{_NS_A}}}rPr")
                rPr.set("sz", str(font_size * 100))
        else:
            # Remove legend
            legend = chart_space.find(f".//{{{_NS_C}}}legend")
            if legend is not None:
                chart_space.remove(legend)

        return True
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Axis control
# ---------------------------------------------------------------------------

def set_axis_title(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    axis: str = "value",
    title: str | None = None,
) -> bool:
    """Set or remove an axis title.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    axis : str
        ``"value"`` (Y-axis) or ``"category"`` (X-axis).
    title : str, optional
        The title text.  None removes the title.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")

        if plot_area is None:
            return False

        # Find the target axis
        axis_tag = f"{{{_NS_C}}}valAx" if axis == "value" else f"{{{_NS_C}}}catAx"
        axis_elem = plot_area.find(axis_tag)

        if axis_elem is None:
            return False

        if title is not None:
            # Find or create title
            title_elem = axis_elem.find(f"{{{_NS_C}}}title")
            if title_elem is None:
                title_elem = etree.SubElement(axis_elem, f"{{{_NS_C}}}title")

            # Clear and set
            for child in list(title_elem):
                title_elem.remove(child)

            tx = etree.SubElement(title_elem, f"{{{_NS_C}}}tx")
            rich = etree.SubElement(tx, f"{{{_NS_C}}}rich")
            etree.SubElement(rich, f"{{{_NS_A}}}bodyPr")
            etree.SubElement(rich, f"{{{_NS_A}}}lstStyle")
            p = etree.SubElement(rich, f"{{{_NS_A}}}p")
            r = etree.SubElement(p, f"{{{_NS_A}}}r")
            t = etree.SubElement(r, f"{{{_NS_A}}}t")
            t.text = title
        else:
            # Remove title
            title_elem = axis_elem.find(f"{{{_NS_C}}}title")
            if title_elem is not None:
                axis_elem.remove(title_elem)

        return True
    finally:
        _save_prs(prs, path)


def set_axis_range(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    axis: str = "value",
    min_val: float | None = None,
    max_val: float | None = None,
    major_unit: float | None = None,
    minor_unit: float | None = None,
) -> bool:
    """Set value axis range and units.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        axis_tag = f"{{{_NS_C}}}valAx" if axis == "value" else f"{{{_NS_C}}}catAx"
        axis_elem = plot_area.find(axis_tag)
        if axis_elem is None:
            return False

        if min_val is not None:
            scaling = _find_or_create(axis_elem, "scaling")
            min_max = _find_or_create(scaling, "min")
            min_max.set("val", str(min_val))

        if max_val is not None:
            scaling = _find_or_create(axis_elem, "scaling")
            max_elem = _find_or_create(scaling, "max")
            max_elem.set("val", str(max_val))

        if major_unit is not None and axis == "value":
            unit = axis_elem.find(f"{{{_NS_C}}}majorUnit")
            if unit is None:
                from lxml import etree
                unit = etree.SubElement(axis_elem, f"{{{_NS_C}}}majorUnit")
            unit.set("val", str(major_unit))

        if minor_unit is not None and axis == "value":
            unit = axis_elem.find(f"{{{_NS_C}}}minorUnit")
            if unit is None:
                from lxml import etree
                unit = etree.SubElement(axis_elem, f"{{{_NS_C}}}minorUnit")
            unit.set("val", str(minor_unit))

        return True
    finally:
        _save_prs(prs, path)


def set_axis_number_format(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    axis: str = "value",
    format_code: str = "General",
    linked_to_source: bool = False,
) -> bool:
    """Set the number format for an axis.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    format_code : str
        Excel-style format code, e.g. ``"#,##0"``, ``"0.0%"``, ``"$#,##0"``.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        axis_tag = f"{{{_NS_C}}}valAx" if axis == "value" else f"{{{_NS_C}}}catAx"
        axis_elem = plot_area.find(axis_tag)
        if axis_elem is None:
            return False

        numFmt = axis_elem.find(f"{{{_NS_C}}}numFmt")
        if numFmt is None:
            from lxml import etree
            numFmt = etree.SubElement(axis_elem, f"{{{_NS_C}}}numFmt")
        numFmt.set("formatCode", format_code)
        numFmt.set("sourceLinked", str(linked_to_source).lower())

        return True
    finally:
        _save_prs(prs, path)


def toggle_gridlines(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    axis: str = "value",
    major: bool | None = None,
    minor: bool | None = None,
) -> bool:
    """Show or hide axis gridlines.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    major : bool, optional
        Show major gridlines.
    minor : bool, optional
        Show minor gridlines.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        axis_tag = f"{{{_NS_C}}}valAx" if axis == "value" else f"{{{_NS_C}}}catAx"
        axis_elem = plot_area.find(axis_tag)
        if axis_elem is None:
            return False

        if major is not None:
            if major:
                gridlines = axis_elem.find(f"{{{_NS_C}}}majorGridlines")
                if gridlines is None:
                    etree.SubElement(axis_elem, f"{{{_NS_C}}}majorGridlines")
            else:
                gridlines = axis_elem.find(f"{{{_NS_C}}}majorGridlines")
                if gridlines is not None:
                    axis_elem.remove(gridlines)

        if minor is not None:
            if minor:
                gridlines = axis_elem.find(f"{{{_NS_C}}}minorGridlines")
                if gridlines is None:
                    etree.SubElement(axis_elem, f"{{{_NS_C}}}minorGridlines")
            else:
                gridlines = axis_elem.find(f"{{{_NS_C}}}minorGridlines")
                if gridlines is not None:
                    axis_elem.remove(gridlines)

        return True
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Chart type switching
# ---------------------------------------------------------------------------

def set_chart_type(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    chart_type: str,
) -> bool:
    """Change the chart type while preserving data.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    chart_type : str
        Target chart type: "column", "bar", "line", "area", "pie", "doughnut",
        "scatter", "radar", "bubble".

    Note
    ----
    Some type conversions may lose formatting (e.g., pie → scatter doesn't
    make visual sense).  The function preserves data and attempts to carry
    over styling where possible.
    """
    from lxml import etree

    TYPE_TO_TAG = {
        "column": ("barChart", {"barDir": "col", "grouping": "clustered"}),
        "column_stacked": ("barChart", {"barDir": "col", "grouping": "stacked"}),
        "bar": ("barChart", {"barDir": "bar", "grouping": "clustered"}),
        "bar_stacked": ("barChart", {"barDir": "bar", "grouping": "stacked"}),
        "line": ("lineChart", {}),
        "line_markers": ("lineChart", {}),
        "area": ("areaChart", {"grouping": "standard"}),
        "area_stacked": ("areaChart", {"grouping": "stacked"}),
        "pie": ("pieChart", {}),
        "doughnut": ("doughnutChart", {}),
        "scatter": ("scatterChart", {}),
        "radar": ("radarChart", {}),
        "radar_markers": ("radarChart", {}),
    }

    if chart_type not in TYPE_TO_TAG:
        raise ValueError(f"Unknown chart type: {chart_type!r}. Valid: {sorted(TYPE_TO_TAG.keys())}")

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart
        chart_space = chart._chartSpace
        plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
        if plot_area is None:
            return False

        # Find existing chart element and extract series data
        old_chart = None
        for child in plot_area:
            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                old_chart = child
                break

        if old_chart is None:
            return False

        # Extract series data from old chart
        series_data: list[dict] = []
        for ser in old_chart.findall(f"{{{_NS_C}}}ser"):
            ser_info = {}
            # Series name
            tx = ser.find(f"{{{_NS_C}}}tx")
            if tx is not None:
                v = tx.find(f".//{{{_NS_C}}}v")
                if v is not None and v.text:
                    ser_info["name"] = v.text

            # Values
            val = ser.find(f"{{{_NS_C}}}val")
            if val is not None:
                values = []
                for pt in val.findall(f".//{{{_NS_C}}}pt"):
                    v = pt.find(f"{{{_NS_C}}}v")
                    if v is not None and v.text:
                        try:
                            values.append(float(v.text))
                        except ValueError:
                            pass
                ser_info["values"] = values

            # Categories (from first series only)
            if not series_data:
                cat = ser.find(f"{{{_NS_C}}}cat")
                if cat is not None:
                    categories = []
                    for pt in cat.findall(f".//{{{_NS_C}}}pt"):
                        v = pt.find(f"{{{_NS_C}}}v")
                        if v is not None and v.text:
                            categories.append(v.text)
                    ser_info["categories"] = categories

            # Color
            spPr = ser.find(f"{{{_NS_C}}}spPr")
            if spPr is not None:
                srgb = spPr.find(f".//{{{_NS_A}}}srgbClr")
                if srgb is not None:
                    ser_info["color"] = srgb.get("val", "")

            series_data.append(ser_info)

        # Remove old chart element
        plot_area.remove(old_chart)

        # Create new chart element
        new_tag, extra_attrs = TYPE_TO_TAG[chart_type]
        new_chart = etree.SubElement(plot_area, f"{{{_NS_C}}}{new_tag}")
        for attr, val in extra_attrs.items():
            new_chart.set(attr, val)

        # Rebuild series
        categories = series_data[0].get("categories", []) if series_data else []
        for i, sd in enumerate(series_data):
            ser = etree.SubElement(new_chart, f"{{{_NS_C}}}ser")

            # idx & order
            idx = etree.SubElement(ser, f"{{{_NS_C}}}idx")
            idx.set("val", str(i))
            order = etree.SubElement(ser, f"{{{_NS_C}}}order")
            order.set("val", str(i))

            # Series name
            if "name" in sd:
                tx = etree.SubElement(ser, f"{{{_NS_C}}}tx")
                strRef = etree.SubElement(tx, f"{{{_NS_C}}}strRef")
                strCache = etree.SubElement(strRef, f"{{{_NS_C}}}strCache")
                ptCount = etree.SubElement(strCache, f"{{{_NS_C}}}ptCount")
                ptCount.set("val", "1")
                pt = etree.SubElement(strCache, f"{{{_NS_C}}}pt")
                pt.set("idx", "0")
                etree.SubElement(pt, f"{{{_NS_C}}}v").text = sd["name"]

            # Categories (first series only)
            if i == 0 and categories:
                cat = etree.SubElement(ser, f"{{{_NS_C}}}cat")
                strRef = etree.SubElement(cat, f"{{{_NS_C}}}strRef")
                strCache = etree.SubElement(strRef, f"{{{_NS_C}}}strCache")
                ptCount = etree.SubElement(strCache, f"{{{_NS_C}}}ptCount")
                ptCount.set("val", str(len(categories)))
                for j, c in enumerate(categories):
                    pt = etree.SubElement(strCache, f"{{{_NS_C}}}pt")
                    pt.set("idx", str(j))
                    etree.SubElement(pt, f"{{{_NS_C}}}v").text = str(c)

            # Values
            if "values" in sd:
                val = etree.SubElement(ser, f"{{{_NS_C}}}val")
                numRef = etree.SubElement(val, f"{{{_NS_C}}}numRef")
                numCache = etree.SubElement(numRef, f"{{{_NS_C}}}numCache")
                etree.SubElement(numCache, f"{{{_NS_C}}}formatCode").text = "General"
                ptCount = etree.SubElement(numCache, f"{{{_NS_C}}}ptCount")
                ptCount.set("val", str(len(sd["values"])))
                for j, v in enumerate(sd["values"]):
                    pt = etree.SubElement(numCache, f"{{{_NS_C}}}pt")
                    pt.set("idx", str(j))
                    etree.SubElement(pt, f"{{{_NS_C}}}v").text = str(v)

            # Color
            if "color" in sd:
                spPr = etree.SubElement(ser, f"{{{_NS_C}}}spPr")
                solidFill = etree.SubElement(spPr, f"{{{_NS_A}}}solidFill")
                srgbClr = etree.SubElement(solidFill, f"{{{_NS_A}}}srgbClr")
                srgbClr.set("val", sd["color"])

            # For scatter charts, add xValues
            if new_tag == "scatterChart" and "values" in sd:
                # Remove val, add xVal and yVal
                val_elem = ser.find(f"{{{_NS_C}}}val")
                if val_elem is not None:
                    ser.remove(val_elem)
                    # yVal
                    yVal = etree.SubElement(ser, f"{{{_NS_C}}}yVal")
                    numRef = etree.SubElement(yVal, f"{{{_NS_C}}}numRef")
                    numCache = etree.SubElement(numRef, f"{{{_NS_C}}}numCache")
                    etree.SubElement(numCache, f"{{{_NS_C}}}formatCode").text = "General"
                    ptCount = etree.SubElement(numCache, f"{{{_NS_C}}}ptCount")
                    ptCount.set("val", str(len(sd["values"])))
                    for j, v in enumerate(sd["values"]):
                        pt = etree.SubElement(numCache, f"{{{_NS_C}}}pt")
                        pt.set("idx", str(j))
                        etree.SubElement(pt, f"{{{_NS_C}}}v").text = str(v)

                    # xVal (sequential indices)
                    xVal = etree.SubElement(ser, f"{{{_NS_C}}}xVal")
                    numRef = etree.SubElement(xVal, f"{{{_NS_C}}}numRef")
                    numCache = etree.SubElement(numRef, f"{{{_NS_C}}}numCache")
                    etree.SubElement(numCache, f"{{{_NS_C}}}formatCode").text = "General"
                    ptCount = etree.SubElement(numCache, f"{{{_NS_C}}}ptCount")
                    ptCount.set("val", str(len(sd["values"])))
                    for j in range(len(sd["values"])):
                        pt = etree.SubElement(numCache, f"{{{_NS_C}}}pt")
                        pt.set("idx", str(j))
                        etree.SubElement(pt, f"{{{_NS_C}}}v").text = str(j + 1)

            # Line markers for line_markers
            if chart_type == "line_markers":
                marker = etree.SubElement(ser, f"{{{_NS_C}}}marker")
                symbol = etree.SubElement(marker, f"{{{_NS_C}}}symbol")
                symbol.set("val", "circle")

        # Add varyColors for pie/doughnut
        if new_tag in ("pieChart", "doughnutChart"):
            varyColors = etree.SubElement(new_chart, f"{{{_NS_C}}}varyColors")
            varyColors.set("val", "1")

        # Add axes for non-pie types
        if new_tag not in ("pieChart", "doughnutChart"):
            # Category axis
            catAx = etree.SubElement(plot_area, f"{{{_NS_C}}}catAx")
            axId = etree.SubElement(catAx, f"{{{_NS_C}}}axId")
            axId.set("val", "10")
            scaling = etree.SubElement(catAx, f"{{{_NS_C}}}scaling")
            orientation = etree.SubElement(scaling, f"{{{_NS_C}}}orientation")
            orientation.set("val", "minMax")
            delete = etree.SubElement(catAx, f"{{{_NS_C}}}delete")
            delete.set("val", "0")
            axPos = etree.SubElement(catAx, f"{{{_NS_C}}}axPos")
            axPos.set("val", "b")
            crossAx = etree.SubElement(catAx, f"{{{_NS_C}}}crossAx")
            crossAx.set("val", "20")

            # Value axis
            valAx = etree.SubElement(plot_area, f"{{{_NS_C}}}valAx")
            axId = etree.SubElement(valAx, f"{{{_NS_C}}}axId")
            axId.set("val", "20")
            scaling = etree.SubElement(valAx, f"{{{_NS_C}}}scaling")
            orientation = etree.SubElement(scaling, f"{{{_NS_C}}}orientation")
            orientation.set("val", "minMax")
            delete = etree.SubElement(valAx, f"{{{_NS_C}}}delete")
            delete.set("val", "0")
            axPos = etree.SubElement(valAx, f"{{{_NS_C}}}axPos")
            axPos.set("val", "l")
            crossAx = etree.SubElement(valAx, f"{{{_NS_C}}}crossAx")
            crossAx.set("val", "10")

        return True
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Data import/export
# ---------------------------------------------------------------------------

def export_chart_data(
    prs_or_path,
    slide_index: int,
    shape_name: str,
) -> dict:
    """Export chart data as a plain dict.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).

    Returns
    -------
    dict
        ``{"categories": [...], "series": [{"name": ..., "values": [...]}, ...]}``
    """
    not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        info = get_chart_info(prs, slide_index, shape_name)
        if info is None:
            return {}

        return {
            "categories": info.categories,
            "series": [
                {"name": s.name, "values": s.values}
                for s in info.series
            ],
        }
    finally:
        pass


def import_chart_data(
    prs_or_path,
    slide_index: int,
    shape_name: str,
    *,
    data: dict,
) -> bool:
    """Import chart data from a plain dict.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    data : dict
        ``{"categories": [...], "series": [{"name": ..., "values": [...]}, ...]}``
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        shape = _find_chart_shape(slide, shape_name)
        if shape is None:
            return False

        chart = shape.chart

        # Update categories
        categories = data.get("categories", [])
        if categories:
            _set_categories(chart, categories)

        # Update series
        series_list = data.get("series", [])
        for i, s_data in enumerate(series_list):
            name = s_data.get("name", f"Series {i+1}")
            values = s_data.get("values", [])

            # Try to edit existing series
            try:
                existing_series = chart.series
                if i < len(existing_series):
                    _set_series_values(chart, i, values)
                    # Rename
                    chart_space = chart._chartSpace
                    plot_area = chart_space.find(f".//{{{_NS_C}}}plotArea")
                    if plot_area is not None:
                        for child in plot_area:
                            if child.tag.endswith("Chart") or child.tag.endswith("chart"):
                                ser_list = child.findall(f"{{{_NS_C}}}ser")
                                if i < len(ser_list):
                                    tx = ser_list[i].find(f"{{{_NS_C}}}tx")
                                    if tx is not None:
                                        v = tx.find(f".//{{{_NS_C}}}v")
                                        if v is not None:
                                            v.text = name
                else:
                    # Add new series
                    add_chart_series(prs, slide_index, shape_name, name=name, values=values)
            except Exception as exc:
                log.warning("Failed to import series %d: %s", i, exc)

        return True
    finally:
        _save_prs(prs, path)
