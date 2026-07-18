# PPTX engine API

## Contents

1. Built-in generation
2. Template profiles
3. Section schema
4. Template catalog
5. Rendering and validation

## Built-in generation

Import `scripts/pptx_helper.py` and call:

```python
from pptx_helper import auto_generate_ppt

auto_generate_ppt(
    title="2026 行业趋势",
    subtitle="Executive Briefing",
    sections=[
        {"title": "核心判断", "bullets": ["判断一", "判断二", "判断三"]},
        {"title": "关键数据", "metrics": [
            {"label": "市场规模", "value": "¥820B", "change": "+18%"}
        ]},
    ],
    output_path="output/report.pptx",
    template_key="strategy-consulting",
    auto_search_images=False,
)
```

Use `theme_key` for the legacy nine-color themes. Use `template_key` for a reusable profile that also controls layout options and deck rhythm. Use `template_profile` to pass an in-memory profile.

## Template profiles

List profiles:

```powershell
python scripts/template_engine.py list
```

Generate and register a profile from style requirements:

```powershell
python scripts/template_engine.py generate `
  --name "Aurora Launch" `
  --prompt "深色、克制、精密仪器感的产品发布模板" `
  --brand-color "#184E77" `
  --brand-color "#F4A261" `
  --register `
  --preview output/aurora-preview.pptx
```

Generate a preview for an existing profile:

```powershell
python scripts/template_engine.py preview strategy-consulting `
  --output output/strategy-preview.pptx
```

Minimum profile structure:

```json
{
  "id": "brand-report",
  "name": "Brand Report",
  "description": "Warm editorial annual report",
  "layout_family": "editorial_grid",
  "theme": {
    "bg": "#FAF8F3",
    "primary": "#243B32",
    "accent": "#C96845",
    "text": "#1C2421"
  },
  "layout_opts": {
    "cover": {"title_size": 58},
    "dashboard": {"metric_value_size": 38}
  },
  "preferred_sequence": ["cover", "toc", "dashboard", "comparison", "end"]
}
```

All color values must use `#RRGGBB`. Valid layout names are `cover`, `toc`, `section`, `bullets`, `text_image`, `full_image`, `image_grid`, `dashboard`, `timeline`, `comparison`, `quote`, `process`, `table`, and `end`.

Catalog and generated profiles can use three non-card layout families: `editorial_grid`, `technical_axis`, and `poster_column`. Natural-language generation selects a family from style cues instead of only changing colors. These families use sharp grids, hairlines, whitespace, asymmetric typography, data axes, or poster columns instead of generic rounded cards, decorative circles, oversized quote marks, `VS` badges, and decorative English kickers. Set `layout_family` to `standard` only when the older card-oriented system is intentional.

## Section schema

| Field | Type | Purpose |
|---|---|---|
| `title` | string | Required slide title |
| `subtitle` | string | Supporting title |
| `bullets` | string list | Three to seven points |
| `images` | path list | Local images |
| `image_query` | string | Pixabay search query |
| `layout` | string | Force a layout |
| `kicker` | string | Small header label |
| `metrics` | object list | KPI cards with `label`, `value`, `change` |
| `events` | object list | Timeline nodes with `date`, `title` |
| `steps` | string list | Process steps |
| `table_headers` | string list | Table headers |
| `table_rows` | list list | Table rows |
| `left`, `right` | object | Two-sided comparison |
| `quote`, `source` | string | Quotation page |
| `layout_opts` | object | Per-slide layout overrides |

## Template catalog

The catalog contains these built-in profiles:

- `strategy-consulting`
- `executive-dark`
- `product-launch`
- `data-story`
- `startup-pitch`
- `academic-clean`
- `training-friendly`
- `government-formal`
- `healthcare-calm`
- `financial-luxe`
- `creative-editorial`
- `sustainability`

Generated profiles are stored in `assets/templates/generated/` and become available by their profile ID.

## Rendering and validation

Run structural validation:

```python
from pptx_helper import auto_validate_ppt
result = auto_validate_ppt("output/report.pptx")
```

Render every slide:

```powershell
python scripts/render_slides.py output/report.pptx `
  --output output/report-preview --dpi 150
```

Always inspect the PNGs. Structural validation cannot detect awkward crops, overflow, low contrast, or incorrect visual hierarchy.
