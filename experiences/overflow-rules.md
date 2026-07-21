# Overflow Rules

## TL;DR

When content exceeds its allocated zone, the system must handle it gracefully — never silently clip or corrupt. The strategy depends on content type: text reduces font size or paginates, tables switch to wider variants, images contain/cover, and charts limit categories. Always respect safe margins and minimum font sizes.

---

## Rule 1: Text overflow — reduce font size or paginate

If rendered text height exceeds the zone height, there are two strategies in priority order:

1. **Reduce font size** (down to the minimum, see Rule 6)
2. **Paginate** — split content across multiple slides with continuation markers

**Implementation in the pipeline:**
```python
# text_metrics.py: bisect_font_size finds the largest size that fits
font_size = bisect_font_size(
    text=content,
    max_width=zone.width,
    max_height=zone.height,
    font_path=font_path,
    min_size=10,  # body text minimum
    max_size=original_size,
)

if font_size < 10:
    # Cannot fit even at minimum size → paginate
    pages = paginate_content_spec(spec, role="bullets")
```

The `pagination.py` module handles splitting dense content into multiple `SlideSpec` objects. The `semantic_qa.py` module detects overflow after rendering via `check_text_overflow()`.

---

## Rule 2: Table overflow — switch to wide recipe variant

When a table has too many columns for the standard zone width, the `wide` recipe variant provides more horizontal space by using narrower margins and a full-width layout.

**Decision logic:**
```python
# In deck_planner.py or pagination.py
if column_count > 6 and role == "table":
    recipe_variant = "table_wide"
elif column_count > 8:
    # Even wide variant can't fit → split table across slides
    pages = paginate_table(table_spec, max_columns=8)
```

Table pagination splits by rows (not columns) — each continuation slide repeats the header row for context.

---

## Rule 3: Image overflow — always contain/cover, never stretch

Images must stay within their zone boundaries. Two strategies:

- **Contain**: scale image to fit entirely within the zone (may leave letterboxing)
- **Cover**: scale image to fill the zone entirely (may crop edges)

**Never stretch** an image beyond its zone — this causes distortion and is flagged by `semantic_qa.py` as a distortion error.

```python
# image_crop.py
def contain(image_size, zone_size):
    """Scale image to fit within zone, preserving aspect ratio."""
    scale = min(zone_size.w / image_size.w, zone_size.h / image_size.h)
    return (image_size.w * scale, image_size.h * scale)

def cover(image_size, zone_size):
    """Scale image to fill zone, preserving aspect ratio (may crop)."""
    scale = max(zone_size.w / image_size.w, zone_size.h / image_size.h)
    return (image_size.w * scale, image_size.h * scale)
```

The `image_crop.py` module also provides `smart_crop()` which uses entropy-based focus detection to choose the best crop region.

---

## Rule 4: Chart overflow — limit categories, paginate or aggregate

Charts with too many categories become unreadable. Enforce these limits:

| Chart Type | Max Categories | Action When Exceeded |
|---|---|---|
| Pie / Doughnut | 8 | Aggregate small slices into "Other" |
| Bar / Column | 12 | Switch to `wide` variant or paginate |
| Line | 20 | Paginate or reduce data granularity |
| Scatter | N/A (points) | Limit to 100 points for performance |

**Aggregation example:**
```python
def aggregate_small_slices(categories, values, threshold=0.05):
    """Merge categories below 5% into 'Other'."""
    total = sum(values)
    result_cats, result_vals = [], []
    other_total = 0
    for cat, val in zip(categories, values):
        if val / total < threshold:
            other_total += val
        else:
            result_cats.append(cat)
            result_vals.append(val)
    if other_total > 0:
        result_cats.append("Other")
        result_vals.append(other_total)
    return result_cats, result_vals
```

---

## Rule 5: Safe margin — always respect canvas insets

The canvas has safe insets that must not be violated by any content. These are defined in `CanvasSpec` and accessed via `canvas.safe`:

| Edge | Inset | Rationale |
|---|---|---|
| Top | 36pt | Slide title area, header |
| Left | 48pt | Binding margin, readability |
| Right | 48pt | Binding margin, readability |
| Bottom | 32pt | Footer, slide number area |

**Usage:**
```python
# All content zones must be within safe area
zone.left >= canvas.safe.left
zone.right <= canvas.safe.right
zone.top >= canvas.safe.top
zone.bottom <= canvas.safe.bottom
```

The `semantic_qa.py` module checks for safe margin violations via `check_margin_violation()`.

---

## Rule 6: Minimum font sizes — paginate below these thresholds

| Context | Minimum Size | Below This Action |
|---|---|---|
| Body text | 10pt | Paginate to additional slides |
| Title text | 14pt | Shorten title or use subtitle |
| Data labels | 8pt | Remove labels or use legend |
| Axis labels | 9pt | Rotate labels or reduce categories |

**Why these thresholds:**
- Below 10pt, body text is unreadable on projected presentations
- Below 14pt, titles lose their visual hierarchy
- Below 8pt, even on-screen reading is difficult

The `text_metrics.py` `bisect_font_size()` function takes `min_size` as a parameter. The `repair_engine.py` uses these thresholds to propose pagination repairs when font size would drop below minimum.

---

## Rule 7: Overflow detection in the QA loop

The generation pipeline (`generation_pipeline.py`) runs a closed loop: plan -> render -> QA -> repair. Overflow is detected at the QA stage:

```
SemanticQAEngine.check() → [OverflowIssue, OverlapIssue, ...]
    ↓
RepairEngine.propose_repairs() → [ReduceFontAction, PaginateAction, ...]
    ↓
Apply repairs → re-plan → re-render → re-QA (max 2 passes)
```

Overflow issues are classified by severity:
- **Critical**: text extends beyond slide boundaries
- **High**: text exceeds zone but stays on slide
- **Medium**: font size below recommended minimum
- **Low**: tight but within acceptable bounds
