# Chart Limits

## TL;DR

python-pptx supports only 10 chart types via `XL_CHART_TYPE`. Office 2016 modern charts are not supported. Data labels require python-pptx 1.0+. Embedded Excel workbooks need proper `PackURI` partnames. Double-click-editable charts need four coordinated parts.

---

## Rule 1: Only 10 chart types are supported

This repo supports exactly 10 chart types (`_CHART_TYPE_MAP` in `pptx_renderer.py`): `column_clustered`, `column_stacked`, `bar_clustered`, `bar_stacked`, `line`, `line_markers`, `pie`, `doughnut`, `scatter`, `area`. They map onto python-pptx's `XL_CHART_TYPE` enum:

| Type | Enum Value |
|---|---|
| Column clustered | `XL_CHART_TYPE.COLUMN_CLUSTERED` (51) |
| Column stacked | `XL_CHART_TYPE.COLUMN_STACKED` (52) |
| Column 100% stacked | `XL_CHART_TYPE.COLUMN_STACKED_100` (53) |
| Bar clustered | `XL_CHART_TYPE.BAR_CLUSTERED` (57) |
| Bar stacked | `XL_CHART_TYPE.BAR_STACKED` (58) |
| Bar 100% stacked | `XL_CHART_TYPE.BAR_STACKED_100` (59) |
| Line | `XL_CHART_TYPE.LINE` (4) |
| Pie | `XL_CHART_TYPE.PIE` (5) |
| Scatter | `XL_CHART_TYPE.XY_SCATTER` (-4169) |
| Area | `XL_CHART_TYPE.AREA` (1) |
| Doughnut | `XL_CHART_TYPE.DOUGHNUT` (-4120) |

The enum has more entries than the 10 this repo supports; stick to the `_CHART_TYPE_MAP` keys.

**Not supported** (Office 2016 modern charts):
- Sunburst
- Treemap
- Waterfall
- Funnel
- Box-and-whisker (boxplot)
- Histogram
- Map charts
- Filled map

Attempting to create these via python-pptx will raise `KeyError` or produce corrupt files.

---

## Rule 2: Data labels require python-pptx 1.0+

In python-pptx versions before 1.0, `plot.has_data_labels` and `series.has_data_labels` do not exist. Setting data labels on older versions silently fails or raises `AttributeError`.

**Check before using:**
```python
import pptx
from packaging.version import Version

if Version(pptx.__version__) >= Version("1.0.0"):
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(9)
else:
    # Fallback: no data labels, or manual XML manipulation
    pass
```

---

## Rule 3: Embedded Excel workbooks need PackURI, not bare strings

When adding a chart with an embedded Excel workbook, the partname must be a proper `PackURI` instance, not a plain string. Using a bare string causes `AttributeError` when python-pptx tries to normalize the URI.

**Bad:**
```python
# Bare string causes AttributeError
chart_part = ChartPart(
    partname="/ppt/charts/chart1.xml",  # string, not PackURI
    ...
)
```

**Good:**
```python
from pptx.opc.packuri import PackURI

chart_part = ChartPart(
    partname=PackURI("/ppt/charts/chart1.xml"),
    ...
)
```

---

## Rule 4: Double-click-editable charts need four coordinated parts

For a chart to be double-click-editable in PowerPoint (i.e., opening the embedded Excel data on double-click), four things must be present and correctly linked:

1. **Chart XML** (`/ppt/charts/chartN.xml`) — the chart definition with `<c:chart>` root
2. **Relationship** — from the chart part to the Excel workbook part (`rId` in `chartN.xml.rels`)
3. **Excel part** (`/ppt/charts/embedded/chartDataN.xlsx`) — the embedded workbook with chart data
4. **Content_Types update** — `[Content_Types].xml` must have an override entry for the xlsx part

Missing any of these causes:
- Chart renders but is not editable (missing relationship or xlsx)
- PowerPoint repair prompt on open (missing Content_Types entry)
- Chart shows "Data may have been changed" warning (stale relationship)

**Minimal wiring:**
```python
from pptx.opc.packuri import PackURI

# 1. Chart part
chart_partname = PackURI("/ppt/charts/chart1.xml")

# 2. Add relationship from chart to xlsx
chart_part.relate_to(xlsx_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")

# 3. XLSX part with proper PackURI
xlsx_partname = PackURI("/ppt/charts/embedded/chartData1.xlsx")

# 4. Content_Types override
pptx_part._package._content_types.add_override(
    xlsx_partname, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
```

---

## Rule 5: Chart category limits for readability

Even though python-pptx does not enforce category limits, practical readability degrades quickly:

- **8 categories max** for pie/doughnut charts — more slices become indistinguishable
- **12 categories max** for bar/column charts — more bars become too narrow
- **20 data points max** for line charts — more points need a wider slide or pagination
- **6 series max** for any chart type — more series create visual clutter

When content exceeds these limits, either:
- Aggregate small categories into "Other"
- Paginate across multiple slides (use the `wide` recipe variant for tables/charts)
- Switch to a different chart type (e.g., bar instead of pie for many categories)
