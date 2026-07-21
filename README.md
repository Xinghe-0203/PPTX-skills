# pptx-skill

> PowerPoint generation, editing, QA, and delivery skill for agent workflows.
>
> Core principle: **always ship an editable `.pptx` and run render verification** -- not just code or XML checks.

**v3.0** -- 50 Python files, 34,000+ lines, 390+ public APIs, 40 modules in `pptx_skill/`.

## Capabilities at a Glance

| Category | Count |
|---|---|
| Layout roles | 24 |
| Recipe variants | 62 |
| Animation types | 51 (20 entrance, 15 exit, 15 emphasis, 1 motion path) |
| Transition types | 18 |
| Themes | 20 |
| Template profiles | 20 |
| Chart types | 10 (with embedded Excel workbooks) |
| Table styles | 20 built-in |
| Analytical framework roles | 5 (SWOT, Porter, PEST, BMC, Funnel) |

## Feature Overview

- **Generation from scratch** -- auto-create full presentations from topics, sections, and data. 24 content roles, 62 recipe variants, 20 template profiles, 3 non-card geometric layout families, 20 traditional themes, font pairing with CJK support.
- **Slide master generation (native)** -- create pages conforming to corporate masters or template standards.
- **Reference slide cloning** -- high-fidelity copy of shape relationships from a reference PPT with content replacement.
- **Visual rebuild** -- reconstruct editable approximations from flattened screenshots or complex structures.
- **Constraint-based adaptive layout** -- declarative `LayoutRecipe` + Kiwi linear constraint solver covering all 24 roles with density-graded variants (sparse/dense/three_column/vertical_dense/horizontal_dense/grid4/wide, etc.).
- **Deck-level planning** -- `deck_planner` uses beam search with anti-repeat signature penalties for cross-slide rhythm.
- **Semantic + render QA** -- overflow/overlap/contrast/font checks, pixel-level perceptual diff, SSIM, window density, four-level QA (table/chart/page/deck), typographic hierarchy and color consistency.
- **Generation-QA-repair loop** -- `run_generation_pipeline` iterates automatically (max 2 passes).
- **Rich text** -- multi-paragraph, multi-run, bullet lists, hyperlinks.
- **10 chart types** -- column_clustered, column_stacked, bar_clustered, bar_stacked, line, line_markers, pie, doughnut, scatter, area, with multi-series and embedded Excel workbooks.
- **Table styling** -- borders, banding, cell fills, column widths, 20 built-in styles, cell merge/unmerge, add row/column.
- **Shape styling** -- line/stroke, shadow, gradient, rotation, text inside shapes.
- **Header, footer, and slide numbers** -- configured via `deck_options`.
- **Speaker notes and media** -- slide notes, video/audio media nodes.
- **Template download** -- GitHub/URL/local template packs, PPTX archive extraction, remote listing and search.
- **Data contract and Manifest V3** -- stable element IDs, embedded XML + sidecar `.manifest.json`.
- **200-page QA annotation dataset and golden renders** -- regression testing and visual baselines.

## New in v3.0

### Animations

51 animation types via OOXML timing XML:
- **Entrance** (20): appear, fly_in, float_up, zoom, grow_turn, swivel, bounce, fade_in, wipe_in, blinds_in, box_in, checkerboard_in, split_in, diagonal_in, random_bars_in, ascend, descend, spin_in, stretch_in, wheel_in
- **Exit** (15): disappear, fly_out, float_down, zoom_out, shrink_turn, swivel_out, bounce_out, fade_out, wipe_out, blinds_out, box_out, checkerboard_out, split_out, diagonal_out, random_bars_out
- **Emphasis** (15): grow_shrink, spin, pulse, color_change, teeter, desaturate, darken, lighten, transparency, object_color, font_color, brush_on_color, brush_on_underline, wave, fill_color
- **Motion path** (1): bounce_end

Animations support sequence ordering, delay, duration, click-trigger or auto-play, and can target text, shape, image, or table nodes.

### Transitions

18 slide transition types: fade, push (left/right/up/down), wipe (left/right/up/down), cover (left/right), split (horizontal_in/out, vertical_in/out), dissolve, random, cut. Configurable duration and advance timing.

### Merge

Multi-deck merge with layout/media/relationship-ID deduplication. Extract specific slides from source decks. Handles duplicate images, fonts, and theme conflicts automatically.

### Watermark

Text and image watermarks with configurable opacity, tiling, z-ordering (above/below content), and removal. Supports diagonal text watermarks and tiled image patterns.

### Sections

Section groups via OOXML `sectionLst`: add, remove, rename, move, and collapse sections. Sections organize the navigation pane in PowerPoint and support collapsible grouping.

### Slide Master

Slide master/layout query, clone, rename, placeholder operations (list, map content to placeholders), and background settings (solid fill, gradient, image). Enables programmatic control of master slides and layouts.

### Image Optimize

Image compression (JPEG quality, max dimension), format conversion (PNG/JPEG/WebP), unused media removal, and per-image stats reporting. Reduces file size without visual quality loss.

### Comments

Speaker notes and review comments via OOXML: add, list, delete, reply, and resolve comments. Full comment threading support with author and timestamp metadata.

### Diff

Structural diff at slide/shape/text/image/layout level with shape matching. Produces a structured diff report highlighting additions, removals, and modifications between two PPTX files.

### Export

PDF, image (PNG/JPEG), HTML, plain text, and thumbnail export. Supports LibreOffice, PyMuPDF, and Pillow backends with automatic fallback.

### SmartArt

SmartArt detection, text extraction and editing, and preservation across clone and merge operations. Reads SmartArt text from OOXML data model while preserving the visual diagram structure.

### VBA

VBA project detection, injection, extraction, and macro-to-shape attachment. Enables programmatic macro insertion for interactive presentations while preserving existing VBA projects.

### Table Styles

20 built-in table styles (light/medium/dark bands), cell merge/unmerge, custom borders, cell fills, add row/column operations. Full OOXML table style support.

### HTML Import

HTML-to-PPTX import using the *local snapshot underlay* technique: extracts semantic content from HTML (Marp, Slidev, reveal.js, LLM-generated slides), optionally renders as a background image via Playwright, and overlays editable native textboxes. Keeps text fully editable while preserving CSS fidelity.

### Analytical Frameworks

Five new layout roles for common business analysis frameworks:
- **SWOT** (quadrant/labeled) -- strengths, weaknesses, opportunities, threats
- **Porter** (diamond/horizontal) -- Porter's Five Forces
- **PEST** (grid/vertical) -- political, economic, social, technological
- **BMC** (canvas/compact) -- Business Model Canvas
- **Funnel** (stacked/narrow) -- sales/marketing funnels

Each role has at least 2 recipe variants.

## Installation

Install in editable mode:

```powershell
python -m pip install -e .
```

Check environment capabilities:

```powershell
python -m pptx_skill.capability
```

Optional backends (install as needed):

```powershell
python -m pip install -e ".[adaptive,schema,qa-image,render-pdf]"
```

> The core path requires only `python-pptx` and `Pillow`; all other backends are optional. See [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md).

## Quick Start

### 1. Generate a new PPT

```python
from pptx_skill import auto_generate_ppt

auto_generate_ppt(
    title="Market Expansion Strategy",
    subtitle="2026 Annual Plan",
    sections=[...],
    output_path="output/report.pptx",
    template_key="strategy-consulting",
    auto_search_images=False,
)
```

### 2. Add animations and transitions

```python
from pptx_skill.animations import add_entrance_animation, ENTRANCE_TYPES
from pptx_skill.transitions import add_transition, FADE

# Add fade-in animation to a shape
add_entrance_animation("output/report.pptx", slide_index=1, shape_name="Title 1",
                       animation_type="fade_in")

# Add fade transition between slides
add_transition("output/report.pptx", slide_index=1, transition_type=FADE, duration_ms=700)
```

### 3. Merge multiple decks

```python
from pptx_skill.merge import merge_decks

merge_decks(
    paths=["deck_a.pptx", "deck_b.pptx", "deck_c.pptx"],
    output_path="merged.pptx",
)
```

### 4. Add a watermark

```python
from pptx_skill.watermark import add_text_watermark

add_text_watermark("output/report.pptx", text="CONFIDENTIAL",
                   opacity=0.15, diagonal=True, font_size=48)
```

### 5. Export to PDF

```python
from pptx_skill.export import export_pdf

export_pdf("output/report.pptx", output_path="output/report.pdf")
```

### 6. Edit a skill-generated PPT (lossless round-trip)

```python
from pptx_skill import load_project, edit_section, regenerate

project = load_project("output/report.pptx")
edit_section("output/report.pptx", index=2, changes={"bullets": ["New point A", "New point B"]})
```

### 7. Edit an external PPT (best-effort targeted edits)

```python
from pptx_skill import edit_text, swap_image, recolor

edit_text("reference.pptx", slide_index=0, find="Old Title", replace="New Title")
recolor("reference.pptx", old_hex="#184E77", new_hex="#2D5016")
```

### 8. Import from HTML

```python
from pptx_skill.html_import import import_from_html

import_from_html("slides.html", output_path="output/from_html.pptx", snapshot_underlay=True)
```

### 9. Diff two presentations

```python
from pptx_skill.diff import diff_presentations

report = diff_presentations("v1.pptx", "v2.pptx")
for change in report["changes"]:
    print(f"{change['level']}: {change['description']}")
```

## Core Concepts

| Concept | Description |
|---|---|
| `ContentSpec` / `SlideSpec` / `ElementSpec` | Data model for deck/slide/element with stable `element_id`. |
| `LayoutRecipe` | Declarative layout formula: variables, constraints, geometry signature; solved by Kiwi into `SolvedGeometry`. |
| `TemplateProfileV2` | Three-layer token (design/semantic/component) template profile with V1 migration. |
| `Manifest V3` | Per-generation record of content, layout plans, render trace, and repair attempts; embedded in pptx XML or sidecar `.manifest.json`. |
| `Reference Modes` | `native` / `clone` / `visual-rebuild` strategies for handling reference decks. |
| `ANIMATION_TYPES` | Catalog of 51 animation presets with OOXML preset class/ID/subtype mappings. |
| `TRANSITION_TYPES` | Set of 18 transition presets with OOXML element mappings. |

## Project Structure

```
.
├── pptx_skill/                    # Core Python package (40 modules)
│   ├── api.py                     # Compatibility facade and GenerationResult
│   ├── content_model.py           # ContentSpec / SlideSpec / ElementSpec
│   ├── content_adapter.py         # Adapt legacy Section to ContentSpec
│   ├── layout_engine.py           # LayoutRecipe + Kiwi constraint solver (62 recipes, 24 roles)
│   ├── deck_planner.py            # Deck-level beam search planning
│   ├── pagination.py              # Role-specific paginators (bullets/table/timeline/process/image_grid)
│   ├── text_metrics.py            # Font metrics, CJK line-breaking, font-size bisection
│   ├── semantic_qa.py             # Overflow / overlap / contrast semantic checks
│   ├── render_qa.py               # Pixel-level perceptual diff, SSIM
│   ├── repair_engine.py           # Whitelist repair actions
│   ├── generation_pipeline.py     # Plan -> render -> QA -> repair closed loop
│   ├── manifest.py                # Manifest V3 read/write
│   ├── design_schema.py           # V2 schema, token resolution and validation
│   ├── template_compiler.py       # TemplateProfileV2 compilation with diversity gate
│   ├── template_v2_adapter.py     # V1->V2 migration and V2->legacy adaptation
│   ├── template_downloader.py     # Template pack download and archive extraction
│   ├── reference_adapter.py       # Native / clone / visual-rebuild adapter dispatch
│   ├── visual_rebuild.py          # Flattened deck parsing and adaptive rebuild
│   ├── pptx_renderer.py           # Adaptive renderer (text/image/shape/table/chart/rich-text/media)
│   ├── image_crop.py              # Contain / cover / smart crop
│   ├── preview_renderer.py        # LibreOffice / PyMuPDF / Windows COM render
│   ├── visual_qa.py               # Post-render visual QA checks
│   ├── qa_dataset.py              # 200-page QA annotation dataset generator
│   ├── golden_renders.py          # Per-role golden renders and family decks
│   ├── capability.py              # Runtime environment capability detection
│   ├── animations.py              # 51 animation types (entrance/exit/emphasis/motion path)
│   ├── transitions.py             # 18 slide transition types
│   ├── merge.py                   # Multi-deck merge with deduplication
│   ├── watermark.py               # Text/image watermarks with opacity, tiling, z-order
│   ├── sections.py                # Section groups (add/remove/rename/move/collapse)
│   ├── slide_master.py            # Slide master/layout query, clone, placeholder ops
│   ├── image_optimize.py          # Image compression, format conversion, unused media removal
│   ├── comments.py                # Speaker notes + review comments
│   ├── diff.py                    # Structural diff (slide/shape/text/image/layout)
│   ├── export.py                  # PDF/image/HTML/text/thumbnail export
│   ├── smartart.py                # SmartArt detection, text extraction/editing, preservation
│   ├── vba.py                     # VBA project detection, injection, extraction
│   ├── table_styles.py            # 20 built-in table styles, cell merge, borders, fills
│   ├── html_import.py             # HTML-to-PPTX with local snapshot underlay
│   └── __init__.py                # Package init, sys.path bridge, re-exports
├── scripts/                       # CLI entry points (legacy layer)
│   ├── pptx_helper.py             # Generation engine
│   ├── reference_ppt.py           # Reference deck analysis and generation
│   ├── template_engine.py         # Template profile management
│   ├── render_slides.py           # Screenshot rendering CLI
│   ├── ppt_project.py             # L1 manifest round-trip
│   ├── ppt_inspect.py             # L2 external PPT inspection
│   ├── ppt_edit.py                # L3 targeted edits
│   ├── ppt_pages.py               # L4 page-level operations
│   ├── layout_variants.py         # Non-card layout families
│   └── pixabay_search.py          # Pixabay image search
├── references/                    # Detailed workflow documentation
│   ├── reference-ppt-workflow.md
│   ├── editing-workflow.md
│   ├── engine-api.md
│   └── quality-checklist.md
├── tests/                         # Unit and E2E tests
├── SKILL.md                       # Skill usage manual (agent-facing)
├── THIRD_PARTY_NOTICES.md         # Third-party dependencies and licenses
└── pyproject.toml
```

## Data Flow

```
ContentSpec / SlideSpec / ElementSpec          (content_model.py)
        |
        v
adapt_legacy_sections() or direct construction (content_adapter.py)
        |
        v
paginate_content_spec() -> split dense slides  (pagination.py)
        |
        v
SemanticQAEngine.check() -> pre-render checks  (semantic_qa.py)
        |      (overflow / overlap / contrast / font / distortion)
        v
plan_deck() -> beam search                      (deck_planner.py)
  +- builtin_recipes(role) -> LayoutRecipe[]    (layout_engine.py)
  +- solve_recipe(recipe, canvas, tokens)       (layout_engine.py, kiwisolver)
  |      -> SolvedGeometry
  +- solved_geometry_to_layout_plan()
        |      -> LayoutPlan with PlannedNode[]
        v
render_layout_plans(plans, output_path, deck_options)  (pptx_renderer.py)
        |      -> RenderResult with RenderTraceEntry[]
        |      deck_options: show_slide_numbers, footer_text, header_text, theme
        v
add_transition() / add_animation()              (transitions.py / animations.py)
        v
add_watermark() / add_sections() / ...          (watermark.py / sections.py / ...)
        v
render_preview() -> PNG via LibreOffice/COM     (preview_renderer.py)
        |
        v
render_qa -> pixel-level perceptual diff, SSIM  (render_qa.py)
        |
        v  (if QA fails)
propose_repairs() -> apply_repairs()            (repair_engine.py)
        |      -> loop back to plan_deck (max 2 passes)
        v
save_manifest_v3()                             (manifest.py)
```

Canvas: 16:9 = 959.976 x 540 pt. Point-based coordinates throughout.

## Layout Roles and Recipes

24 roles with 62 recipe variants:

| Role | Variants | Description |
|---|---|---|
| `cover` | 3 | Title/cover slides |
| `toc` | 2 | Table of contents |
| `section` | 2 | Section dividers |
| `bullets` | 5 | Bullet-point content (rail/sparse/dense/split/compact) |
| `text_image` | 2 | Text with image side-by-side |
| `full_image` | 2 | Full-bleed image slides |
| `image_grid` | 3 | Image grid layouts (2x2/3x3/grid4) |
| `dashboard` | 5 | Dashboard/KPI layouts |
| `timeline` | 3 | Timeline layouts (horizontal/vertical_dense/compact) |
| `comparison` | 4 | Side-by-side comparisons (two_column/three_column) |
| `quote` | 2 | Quote slides |
| `process` | 4 | Process flow layouts (horizontal/horizontal_dense/vertical/stepper) |
| `table` | 3 | Data table slides (standard/wide/compact) |
| `end` | 2 | Closing slides |
| `matrix` | 2 | Matrix/2x2 grid |
| `kpi_hero` | 2 | KPI hero metric slides |
| `faq` | 2 | FAQ layouts |
| `testimonial` | 2 | Testimonial/quote with attribution |
| `logo_wall` | 2 | Logo wall / partner slides |
| `swot` | 2 | SWOT analysis (quadrant/labeled) |
| `porter` | 2 | Porter's Five Forces (diamond/horizontal) |
| `pest` | 2 | PEST analysis (grid/vertical) |
| `bmc` | 2 | Business Model Canvas (canvas/compact) |
| `funnel` | 2 | Sales/marketing funnel (stacked/narrow) |

## Editing Paths

| Source | Method | Fidelity |
|---|---|---|
| Skill-generated PPT (has manifest) | `load_project` -> `edit_section` -> `regenerate` | Lossless |
| Skill-generated PPT (has manifest) | `insert_slide` / `delete_slide` / `move_slide` / `replace_layout` | Lossless for new pages |
| External PPT | `inspect_ppt()` -> understand structure | Read-only |
| External PPT | `edit_text` / `swap_image` / `recolor` / `swap_theme` | Best-effort |
| HTML slides | `import_from_html()` -> editable PPTX | Underlay + editable text |

All editing functions create `.bak.pptx` backup before writing. Slide indices are 1-based in public APIs.

## Testing

Run the full test suite:

```powershell
python -m pytest tests/ -ra
```

Baseline: **185 tests passing** (including chart/rich-text/table-styling/media/template-downloader/animations/transitions/merge/watermark/sections/slide_master/image_optimize/comments/diff/export/smartart/vba/table_styles coverage).

## Pre-Delivery Checklist

- Render all slides and check each page for overflow, occlusion, image cropping, contrast, layer ordering, and template consistency.
- Run at least two rounds of "generate -> render -> check -> fix".
- See [`references/quality-checklist.md`](references/quality-checklist.md).

## Supported Boundaries

- Suited for static editable pages, images, tables, and common chart types.
- Complex animations, macros, OLE objects, and embedded workbooks may not be fully preserved.
- Flat screenshots cannot recover original vector data, animations, or invisible elements; visual rebuild produces editable approximations.
- When Pixabay search fails, generation falls back to local images or text-only layouts.

## Documentation Index

- [`SKILL.md`](SKILL.md) -- Complete agent-facing usage manual.
- [`references/reference-ppt-workflow.md`](references/reference-ppt-workflow.md) -- Handling user-provided reference decks.
- [`references/editing-workflow.md`](references/editing-workflow.md) -- Editing existing PPTs.
- [`references/engine-api.md`](references/engine-api.md) -- Generation API and template profiles.
- [`references/quality-checklist.md`](references/quality-checklist.md) -- Pre-delivery quality checklist.
- [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md) -- Third-party dependencies and licenses.

## License

Core path is MIT-style licensed (see repository LICENSE). When using optional backends, comply with their respective licenses -- especially PyMuPDF's AGPL/commercial terms.

---

*Generated for pptx-skill v3.0.0.*
