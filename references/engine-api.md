# PPTX engine API

## Contents

1. Built-in generation
2. Template profiles
3. Section schema
4. Template catalog
5. Rendering and validation
6. Adaptive pipeline and deck planning

## Built-in generation

Import `auto_generate_ppt` from the `pptx_skill` package and call:

```python
from pptx_skill import auto_generate_ppt

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

Use `theme_key` for the 20 legacy themes. Use `template_key` for a reusable profile that also controls layout options and deck rhythm. Use `template_profile` to pass an in-memory profile.

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

All color values must use `#RRGGBB`. Valid layout roles are `cover`, `toc`, `section`, `bullets`, `text_image`, `full_image`, `image_grid`, `dashboard`, `timeline`, `comparison`, `quote`, `process`, `table`, `end`, `matrix`, `kpi_hero`, `faq`, `testimonial`, `logo_wall`, and the five analytical framework roles `swot`, `porter`, `pest`, `bmc`, `funnel` (24 roles total).

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
| `chart_type` | string | Chart type: `column_clustered`, `column_stacked`, `bar_clustered`, `bar_stacked`, `line`, `line_markers`, `pie`, `doughnut`, `scatter`, `area` |
| `chart_categories` | string list | Chart category labels |
| `chart_series` | object list | Chart series with `name` and `values` |
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

## Template download and import

Download template packs from GitHub, direct URLs, or local directories:

```python
from pptx_skill import download_template_pack, import_template

# Download from GitHub
result = download_template_pack("github", "output/templates", repo="user/repo", branch="main")

# Download from a direct URL
result = download_template_pack("url", "output/templates", url="https://example.com/template.pptx")

# Copy from a local directory
result = download_template_pack("local", "output/templates", path="C:/templates")
```

Import a downloaded `.pptx` as a reusable V2 template profile:

```python
profile = import_template("output/templates/report.pptx", name="Client Brand")
# Profile is saved to assets/templates/generated/ and usable by template_key
```

List curated remote template sources:

```python
from pptx_skill.template_downloader import list_remote_packs
packs = list_remote_packs()
```

All download functions use stdlib `urllib` only (no `requests`). URLs are validated against private-network IPs for SSRF protection.

## Rendering and validation

Run structural validation:

```python
from pptx_skill import auto_validate_ppt
result = auto_validate_ppt("output/report.pptx")
```

Render every slide:

```powershell
python scripts/render_slides.py output/report.pptx `
  --output output/report-preview --dpi 150
```

Always inspect the PNGs. Structural validation cannot detect awkward crops, overflow, low contrast, or incorrect visual hierarchy.

## Adaptive pipeline and deck planning (PR4+)

The structured adaptive pipeline lives in `pptx_skill.generation_pipeline` and is reached via `auto_generate_ppt(..., layout_engine="adaptive")` once fully wired, or directly:

```python
from pptx_skill import run_generation_pipeline, ContentSpec, CanvasSpec

result = run_generation_pipeline(content, "output/report.pptx", qa_mode="report")
```

- `qa_mode`: `off` skips QA and repair; `report` records QA without blocking on quality failures; `strict` raises `PresentationQualityError` if blockers remain after the repair budget.
- Repair is capped at `max_repair_passes` (default 2); each pass re-plans, re-renders and re-QAs. Only whitelist actions (`reduce_font_within_limit`, `switch_layout_candidate`, `change_text_color_to_token`, `remove_empty_placeholder`) are applied automatically.

### Pagination and deck planning (PR6)

Long `bullets` content is split across derived slides by `pptx_skill.pagination.paginate_bullets`, with deterministic derived IDs (`<parent>/frag-<n>` for slides, `<parent>/page-<n>` for split body elements) so QA and repair keep stable mappings after re-generation.

Whole-deck planning is handled by `pptx_skill.deck_planner.plan_deck`, which:

1. builds per-slide candidate bundles (single-page or split);
2. runs a beam search over candidate bundles, accumulating `local_score` plus transition penalties (adjacent geometry-signature similarity, `preferred_sequence`, `rhythm.rules`, section transitions);
3. returns a `DeckPlanResult` with the chosen derived slides and their `LayoutPlan`s.

A profile supplies `preferred_sequence` (ordered layout names) and optional `rhythm.rules` (`{"after": <role>, "prefer": [...], "avoid": [...]}`); both now actually influence candidate selection instead of being metadata-only.

```python
from pptx_skill import plan_deck, LayoutScoringConfig, CanvasSpec

deck = plan_deck(slides, CanvasSpec(959.976, 540), profile_state)
# deck.derived_slides  -> list[SlideSpec] after pagination
# deck.plans            -> list[LayoutPlan] one per derived slide
```

The renderer writes one PPTX slide per derived `LayoutPlan` via `render_layout_plans`, and each `RenderTraceEntry` carries its `slide_index` so QA reports point back to the right slide.

### deck_options

`render_layout_plans` accepts a `deck_options` dict controlling deck-level features:

| Key | Type | Default | Purpose |
|---|---|---|---|
| `show_slide_numbers` | bool | `True` | Show page numbers bottom-right |
| `footer_text` | str | `""` | Centered footer text |
| `header_text` | str | `""` | Left-aligned header text |
| `slide_number_format` | str | `"{current}/{total}"` | Page number format string |
| `theme` | dict | `None` | Theme dict with `text_muted`, `text` color keys |

### New layout roles

Five additional layout roles are available in the adaptive pipeline:

| Role | Variants | Use |
|---|---|---|
| `matrix` | `matrix.quadrant`, `matrix.labeled` | Two-axis quadrant or labeled matrix |
| `kpi_hero` | `kpi_hero.split`, `kpi_hero.full` | Large KPI with supporting detail |
| `faq` | `faq.alternating`, `faq.stacked` | Question-answer pairs |
| `testimonial` | `testimonial.centered`, `testimonial.card` | Customer quote with attribution |
| `logo_wall` | `logo_wall.grid3`, `logo_wall.grid4` | Logo grid display |

### Rich text and media

Text nodes support multi-run formatting via the `runs` binding (list of run dicts with `text`, `bold`, `italic`, `color`, `size`, `href`). Paragraph-level formatting uses the `paragraphs` binding (list with `text`, `align`, `level`, `bullet`).

Video and audio nodes are supported as element kinds `video` and `audio`. Video uses `binding["path"]` and optional `binding["poster"]` for a poster frame. Audio uses `binding["path"]` and optional `binding["mime_type"]` (default `"audio/mpeg"`). Both fall back to placeholder shapes when files are missing.
