# pptx-skill

> PowerPoint generation, editing, QA, and delivery skill for agent workflows.
>
> Core principle: **always ship an editable `.pptx` and run render verification** -- not just code or XML checks.

![v6.0.0](https://img.shields.io/badge/version-6.0.0-blue) ![Python 3.10+](https://img.shields.io/badge/python-3.10+-green) ![267 tests](https://img.shields.io/badge/tests-267-brightgreen) ![64 modules](https://img.shields.io/badge/modules-64-orange) ![600+ APIs](https://img.shields.io/badge/APIs-600+-purple) [![CI](https://github.com/Xinghe-0203/PPTX-skills/actions/workflows/ci.yml/badge.svg)](https://github.com/Xinghe-0203/PPTX-skills/actions/workflows/ci.yml)

**v6.0.0** -- 64 Python modules, 600+ public APIs, 51 animations, 30 transitions, 24 layout roles, 20 themes, unified `Deck` class, unified CLI, Markdown import.

## Capabilities at a Glance

| Category | Count |
|---|---|
| Python modules | 64 |
| Public APIs | 600+ |
| Layout roles | 24 |
| Recipe variants | 62 |
| Animation types | 51 (20 entrance, 15 exit, 15 emphasis, 1 motion path) |
| Transition types | 30 (18 standard + 12 advanced) |
| Themes | 20 |
| Template profiles | 20 |
| Chart types | 10 (with embedded Excel workbooks) |
| Table styles | 20 built-in |
| Analytical framework roles | 5 (SWOT, Porter, PEST, BMC, Funnel) |
| Tests | 267 collected (266 passing, 1 skipped; 6 subtests passing) |

## Feature Overview

- **Unified `Deck` class** -- fluent high-level API wrapping 60+ modules into chainable methods: `Deck.open()`, `Deck.generate()`, `Deck.from_markdown()`, `.add_watermark()`, `.add_transition()`, `.add_animation()`, `.edit_text()`, `.save()`.
- **Unified CLI** -- `pptx-skill <command>` for info, inspect, generate, from-markdown, render, edit, watermark, export, pages, merge, validate, template, capability.
- **Markdown import** -- parse Markdown outlines into sections and generate PPTX in one step.
- **Global find-replace** -- `find_replace_all()` across all slides with run-formatting preservation.
- **Path-based overloads** -- `add_transition(path, slide_index, ...)` and `add_entrance_animation(path, slide_index, ...)` accept file paths or open `Presentation` objects, with standardized 1-based slide indexing.
- **Generation from scratch** -- auto-create full presentations from topics, sections, and data. 24 content roles, 62 recipe variants, 20 template profiles, 20 themes, font pairing with CJK support.
- **Constraint-based adaptive layout** -- declarative `LayoutRecipe` + Kiwi linear constraint solver covering all 24 roles with density-graded variants.
- **Deck-level planning** -- beam search with anti-repeat signature penalties for cross-slide rhythm.
- **Semantic + render QA** -- overflow/overlap/contrast/font checks, pixel-level perceptual diff, SSIM, window density.
- **Generation-QA-repair loop** -- `run_generation_pipeline` iterates automatically (max 2 passes).
- **10 chart types** -- column, bar, line, pie, doughnut, scatter, area variants with multi-series and embedded Excel workbooks.
- **Table styling** -- borders, banding, cell fills, 20 built-in styles, cell merge/unmerge, add row/column.
- **51 animations** -- entrance, exit, emphasis, and motion path via OOXML timing XML.
- **30 transitions** -- 18 standard + 12 advanced (wheel, ripple, honeycomb, vortex, shred, flip, gallery, pan, glitter, warp, wind, curtain) + morph transition.
- **Multi-deck merge** -- layout/media/relationship-ID deduplication with conflict resolution.
- **Watermarks** -- text and image watermarks with opacity, tiling, z-ordering, and removal.
- **Sections, slide masters, SmartArt, VBA** -- full OOXML-level control.
- **Export** -- PDF, images, HTML, Markdown, SVG, text, thumbnails, video (MP4/GIF via ffmpeg).
- **Color science** -- RGB/HSL/HSV/CMYK, palette extraction, harmony generators, gradient builder.
- **OCR** -- Tesseract/EasyOCR/PaddleOCR integration with auto-captioning.
- **Batch processing** -- parallel convert, watermark, recolor, inspect, stats.
- **Accessibility** -- WCAG 2.1 audit, alt-text management, contrast checking, auto-fix.
- **Protection** -- write-protection, password encryption, mark-as-final.
- **Template download** -- GitHub/URL/local template packs with remote listing and search.

## Installation

Install in editable mode (core only -- needs just `python-pptx` and `Pillow`):

```powershell
python -m pip install -e .
```

Install with optional backends:

```powershell
python -m pip install -e ".[adaptive,qa-image,render-pdf]"
```

Check environment capabilities:

```powershell
python -m pptx_skill.capability
```

Optional extras:

| Extra | Provides |
|---|---|
| `adaptive` | Kiwi constraint solver for adaptive layout engine |
| `qa-image` | NumPy + scikit-image for perceptual diff / SSIM |
| `render-pdf` | PyMuPDF for PDF export (AGPL -- see THIRD_PARTY_NOTICES.md) |
| `render-com` | Windows COM rendering via pywin32 |
| `chart-editable` | openpyxl for editable chart Excel workbooks |
| `encrypt` | msoffcrypto-tool for password encryption |
| `ocr-easy` | EasyOCR backend for OCR |
| `ocr-paddle` | PaddleOCR backend for OCR |
| `test` | pytest for running the test suite |
| `build` | build + Twine for validating release distributions |

> The core path requires only `python-pptx` and `Pillow`; all other backends are optional. See [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md).

## Quick Start

### 1. Generate a new PPT

```python
from pptx_skill import auto_generate_ppt

auto_generate_ppt(
    title="Market Expansion Strategy",
    subtitle="2026 Annual Plan",
    sections=[
        {"title": "Market Overview", "bullets": ["TAM $50B", "CAGR 12%"]},
        {"title": "Competitive Landscape", "bullets": ["Top 5 players", "Gap in mid-market"]},
    ],
    output_path="output/report.pptx",
    theme_key="corporate_blue",
    auto_search_images=False,
)
```

### 2. Deck fluent API (new in v6.0)

```python
from pptx_skill import Deck

# Generate from scratch, then chain edits
deck = Deck.generate(
    title="Q3 Review",
    sections=[{"title": "Highlights", "bullets": ["Revenue +15%", "New 3 clients"]}],
    output_path="output/q3.pptx",
)
deck.add_watermark("DRAFT", opacity=0.15, diagonal=True)
deck.add_notes(1, "Opening remarks for the board.")
deck.add_transition(1, "fade", duration_ms=500)
deck.add_animation(slide_index=1, shape_name="Title 1", anim_type="fade_in")
deck.save()

# Or open an existing file
deck = Deck.open("output/report.pptx")
deck.find_replace_all("Old Company", "New Company")
deck.recolor("#184E77", "#2D5016")
deck.save()
```

### 3. Import from Markdown (new in v6.0)

```python
from pptx_skill import import_markdown

# One-step: Markdown -> PPTX
import_markdown("outline.md", output_path="output/deck.pptx", theme_key="corporate_blue")

# Or two-step: parse sections, then generate
from pptx_skill import markdown_to_sections, auto_generate_ppt

sections = markdown_to_sections("outline.md")
auto_generate_ppt(
    title=sections[0]["title"],
    sections=sections[1:],
    output_path="output/deck.pptx",
)
```

Markdown mapping: `# H1` -> cover, `## H2` -> section/content slide, `### H3` -> content slide, `- bullet` -> bullets, `> quote` -> quote slide, GFM tables -> table slide, YAML front-matter -> title/subtitle/theme/lang.

### 4. Unified CLI (new in v6.0)

```bash
# Show summary info
pptx-skill info deck.pptx

# Generate from JSON sections
pptx-skill generate --title "My Deck" --sections sections.json -o out.pptx

# Generate from Markdown
pptx-skill from-markdown outline.md -o out.pptx

# Inspect structure
pptx-skill inspect deck.pptx

# Render slides as PNG
pptx-skill render deck.pptx -o ./previews --dpi 150

# Find-and-replace text
pptx-skill edit deck.pptx --find-replace "Old:New"

# Add watermark
pptx-skill watermark deck.pptx --text "DRAFT" --opacity 0.15 --diagonal

# Export to PDF
pptx-skill export deck.pptx --pdf out.pdf

# Delete a slide
pptx-skill pages deck.pptx --delete 3

# Merge multiple decks
pptx-skill merge a.pptx b.pptx -o merged.pptx

# Validate structure
pptx-skill validate deck.pptx

# List templates
pptx-skill template list

# Show capability report
pptx-skill capability
```

Also works as `python -m pptx_skill <command>`.

### 5. Edit a skill-generated PPT (lossless round-trip)

```python
from pptx_skill import load_project, edit_section, regenerate

project = load_project("output/report.pptx")
edit_section("output/report.pptx", index=2, changes={"bullets": ["New point A", "New point B"]})
regenerate(project)
```

### 6. Edit an external PPT (best-effort targeted edits)

```python
from pptx_skill import edit_text, recolor, swap_theme, find_replace_all

# Per-slide text replacement (1-based slide index)
edit_text("reference.pptx", slide_index=1, find="Old Title", replace="New Title")

# Global find-and-replace across all slides
count = find_replace_all("reference.pptx", "Old Company", "New Company")

# Recolor shapes
recolor("reference.pptx", old_hex="#184E77", new_hex="#2D5016")

# Swap theme
swap_theme("reference.pptx", new_theme_key="corporate_blue")
```

### 7. Add a watermark

```python
from pptx_skill import add_text_watermark

add_text_watermark(
    "output/report.pptx",
    text="CONFIDENTIAL",
    opacity=0.15,
    rotation=-45,
    font_size=48,
)
```

### 8. Export to PDF / images / Markdown

```python
from pptx_skill import export_to_pdf, export_to_images, export_to_markdown

# Export to PDF
export_to_pdf("output/report.pptx", "output/report.pdf", dpi=150)

# Export slides as PNG images
export_to_images("output/report.pptx", "output/slides/", dpi=150, format="PNG")

# Export to Markdown
export_to_markdown("output/report.pptx", output_path="output/report.md")
```

### 9. Merge multiple decks

```python
from pptx_skill import merge_presentations

result = merge_presentations(
    sources=["deck_a.pptx", "deck_b.pptx", "deck_c.pptx"],
    output_path="merged.pptx",
)
print(f"Merged {result.sources_merged} files -> {result.output_path} ({result.total_slides} slides)")
```

### 10. Add transitions and animations

```python
from pptx_skill import add_transition, add_entrance_animation, FADE, FADE_IN

# Path-based: accepts a file path or open Presentation (1-based slide index)
add_transition("output/report.pptx", slide_index=1, transition_type=FADE, duration_ms=700)
add_entrance_animation("output/report.pptx", slide_index=1, shape_name="Title 1", anim_type=FADE_IN)

# Or use the slide/object-based API directly
from pptx_skill import apply_slide_transition, apply_deck_transitions, apply_entrance_animation
from pptx import Presentation

prs = Presentation("output/report.pptx")
apply_deck_transitions(prs, "fade", duration_ms=500)
apply_entrance_animation(prs.slides[0], prs.slides[0].shapes[0], "fade_in")
prs.save("output/report.pptx")
```

### 11. Inspect an external PPT

```python
from pptx_skill import inspect_ppt

report = inspect_ppt("reference.pptx")
print(f"{report['total_slides']} slides")
for slide in report["slides"]:
    print(f"  Slide {slide['index']}: {slide['guessed_layout']}")
```

## Core Concepts

| Concept | Description |
|---|---|
| `Deck` | Unified fluent API wrapping 60+ modules. `Deck.open()`, `.generate()`, `.from_markdown()`, chainable methods. |
| `ContentSpec` / `SlideSpec` / `ElementSpec` | Data model for deck/slide/element with stable `element_id`. |
| `LayoutRecipe` | Declarative layout formula: variables, constraints, geometry signature; solved by Kiwi into `SolvedGeometry`. |
| `TemplateProfileV2` | Three-layer token (design/semantic/component) template profile with V1 migration. |
| `Manifest V3` | Per-generation record of content, layout plans, render trace, and repair attempts; embedded in pptx XML or sidecar `.manifest.json`. |
| `Reference Modes` | `native` / `clone` / `visual-rebuild` strategies for handling reference decks. |
| `ANIMATION_TYPES` | Catalog of 50 animation presets (20 entrance, 15 exit, 15 emphasis) with OOXML preset class/ID/subtype mappings; the motion-path type is applied via `apply_motion_path()` (51 types total). |
| `TRANSITION_TYPES` | Set of 18 standard transition presets; `ADVANCED_TRANSITIONS` adds 12 more. |
| 1-based slide indexing | All public APIs use 1-based slide indices (slide 1 = first slide) across every module. |

## Module List (64 modules)

| Module | Description |
|---|---|
| `api.py` | Compatibility facade wrapping both layers; `auto_generate_ppt`, `auto_validate_ppt`. |
| `deck.py` | **v6.0** Unified fluent `Deck` class and `open_deck()`. |
| `cli.py` | **v6.0** Unified CLI (`pptx-skill` command) with 13 subcommands. |
| `markdown_import.py` | **v6.0** Markdown-to-sections parser and `import_markdown()`. |
| `content_model.py` | Zero-dependency core: `ContentSpec`, `SlideSpec`, `ElementSpec`, `LayoutPlan`, `PlannedNode`, `GeometrySpec`, `BBox`, `CanvasSpec`, `StableIdGenerator`. |
| `content_adapter.py` | Adapt legacy `Section` objects to `ContentSpec`. |
| `layout_engine.py` | Declarative `LayoutRecipe` + Kiwi constraint solver -> `SolvedGeometry`. |
| `deck_planner.py` | Beam search over candidate bundles with transition penalties. |
| `pagination.py` | Role-specific paginators (bullets, table, timeline, process, image_grid). |
| `pptx_renderer.py` | Adaptive renderer dispatching text/image/shape/table/chart nodes. |
| `text_metrics.py` | Pillow font metrics, CJK line-breaking, font-size bisection. |
| `semantic_qa.py` | Overflow/overlap/contrast/font/distortion checks. |
| `render_qa.py` | Pixel-level perceptual diff, SSIM, window density. |
| `repair_engine.py` | Whitelist repair actions with profile override merging. |
| `generation_pipeline.py` | Plan -> render -> QA -> repair closed loop. |
| `manifest.py` | V3 manifest (embedded XML + sidecar JSON) with V2->V3 migration. |
| `design_schema.py` | V2 template schema, three-layer tokens, resolution. |
| `template_compiler.py` | TemplateProfileV2 compilation with diversity gate. |
| `template_v2_adapter.py` | V1->V2 migration, V2->legacy adaptation. |
| `reference_adapter.py` | Native/clone/visual-rebuild adapter dispatch. |
| `visual_rebuild.py` | Flattened-deck parsing and adaptive rebuild. |
| `qa_dataset.py` | 200-page annotation dataset generator. |
| `golden_renders.py` | Per-role golden renders and family decks. |
| `image_crop.py` | Contain/cover/smart crop. |
| `preview_renderer.py` | LibreOffice / PyMuPDF / Windows COM rendering. |
| `template_downloader.py` | Template pack download (GitHub/URL/local), PPTX archive extraction, remote pack listing & search. |
| `visual_qa.py` | Visual QA checks on rendered slides. |
| `capability.py` | Runtime environment capability detection and reporting. |
| `transitions.py` | 18 slide transition types (fade/push/wipe/cover/split/dissolve/random/cut) + path-based overloads. |
| `transitions_ext.py` | 12 advanced transitions (wheel, ripple, honeycomb, vortex, shred, flip, gallery, pan, glitter, warp, wind, curtain). |
| `animations.py` | 51 animation types (20 entrance, 15 exit, 15 emphasis, 1 motion path) via OOXML timing XML + path-based overloads. |
| `morph.py` | Morph transition (p15 namespace), morph-by-object/word/char options. |
| `merge.py` | Multi-deck merge with layout/media/rel-ID deduplication, slide extraction. |
| `watermark.py` | Text/image watermarks with opacity, tiling, z-ordering, removal. |
| `sections.py` | Section groups (add/remove/rename/move/collapse) via OOXML sectionLst. |
| `slide_master.py` | Slide master/layout query, clone, rename, placeholder ops, background settings. |
| `image_optimize.py` | Image compression, format conversion, unused media removal, stats. |
| `comments.py` | Speaker notes + review comments (add/list/delete/reply/resolve). |
| `diff.py` | Structural diff at slide/shape/text/image/layout level with shape matching. |
| `export.py` | PDF/image/HTML/text/thumbnail export (LibreOffice/PyMuPDF/Pillow backends). |
| `smartart.py` | SmartArt detection, text extraction/editing, preservation across clone/merge. |
| `vba.py` | VBA project detection, injection, extraction, macro-to-shape attachment. |
| `table_styles.py` | 20 built-in table styles, cell merge/unmerge, borders, fills, add row/column. |
| `html_import.py` | HTML-to-PPTX import with local snapshot underlay technique. |
| `effects.py` | 3D/bevel/glow/reflection/soft-edges/inner-shadow/perspective-shadow + 16 effect presets. |
| `connectors.py` | Lines, elbow/curved connectors, freeform paths, arrowheads, dash styles. |
| `protection.py` | Write-protection, password encryption (msoffcrypto), mark-as-final. |
| `accessibility.py` | WCAG 2.1 audit, alt-text management, contrast checking, auto-fix. |
| `metadata.py` | Core document properties, custom properties, embedded fonts. |
| `equations.py` | OMML builder + LaTeX-to-OMML converter (fractions, radicals, n-ary, matrices). |
| `multimedia.py` | Video/audio embedding with playback settings (loop, fullscreen, trim). |
| `groups.py` | Shape grouping/ungrouping, z-order within groups. |
| `zoom.py` | Slide Zoom, Section Zoom, Summary Zoom (interactive navigation thumbnails). |
| `batch.py` | Batch/parallel processing (convert, watermark, recolor, inspect, stats). |
| `svg_import.py` | SVG-to-DrawingML converter (path/rect/circle/line/polygon -> custGeom). |
| `svg_export.py` | DrawingML-to-SVG converter (slides -> SVG vector output). |
| `ocr.py` | OCR integration (Tesseract/EasyOCR/PaddleOCR), auto-caption. |
| `chart_edit.py` | Chart data/style/type editing, series add/remove/rename, axis control. |
| `color.py` | Color science (RGB/HSL/HSV/CMYK), palette extraction, harmony generators, gradient builder. |
| `layout_sync.py` | Alignment, distribution, snap-to-grid, layout templates, z-order. |
| `markdown_export.py` | Markdown export (text, tables, notes, chart data, metadata). |
| `video.py` | MP4/GIF video export via ffmpeg (crossfade transitions, speaker timing). |
| `constants.py` | Shared OOXML namespace/IRI constants. |
| `units.py` | EMU/point/inch conversion helpers. |

## Editing Paths

| Source | Method | Fidelity |
|---|---|---|
| Skill-generated PPT (has manifest) | `load_project` -> `edit_section` -> `regenerate` | Lossless |
| Skill-generated PPT (has manifest) | `insert_slide` / `delete_slide` / `move_slide` / `replace_layout` | Lossless for new pages |
| External PPT | `inspect_ppt()` -> understand structure | Read-only |
| External PPT | `edit_text` / `find_replace_all` / `swap_image` / `recolor` / `swap_theme` | Best-effort |
| Markdown file | `import_markdown()` -> editable PPTX | Full generation |
| HTML slides | `import_from_html()` -> editable PPTX | Underlay + editable text |

All editing functions create `.bak.pptx` backup before writing. Slide indices are 1-based in all public APIs.

## Testing

Run the full test suite:

```powershell
python -m pytest tests/ -ra
```

Baseline: **267 tests collected — 266 passing, 1 skipped, plus 6 passing subtests** (including shared atomic I/O, animations, transitions, merge, watermark, sections, slide_master, image_optimize, comments, diff, export, smartart, vba, table_styles, chart_edit, color, layout_sync, markdown_export, video, svg_export, svg_import, ocr, batch, morph, accessibility, effects, connectors, protection, metadata, equations, multimedia, groups, zoom, html_import, deck, markdown_import, cli, and core layout/QA/pipeline coverage).

Run a single test file:

```powershell
pytest tests/test_layout_engine.py
```

Lint and type check:

```powershell
ruff check .
mypy pptx_skill/
```

GitHub Actions runs Ruff, mypy, the full suite on Python 3.10/3.12/3.14,
LibreOffice render integration, and isolated wheel installation. Pushing a
version tag such as `v6.0.0` reruns validation and publishes the wheel and
source distribution to a GitHub Release. The tag must match `[project].version`.

## Pre-Delivery Checklist

- Render all slides and check each page for overflow, occlusion, image cropping, contrast, layer ordering, and template consistency.
- Run at least two rounds of "generate -> render -> check -> fix".
- See [`references/quality-checklist.md`](references/quality-checklist.md).

## Supported Boundaries

- Suited for static editable pages, images, tables, and common chart types.
- Complex animations, macros, OLE objects, and embedded workbooks may not be fully preserved.
- Flat screenshots cannot recover original vector data, animations, or invisible elements; visual rebuild produces editable approximations.
- When Pixabay search fails, generation falls back to local images or text-only layouts.
- Pixabay downloads write a `pixabay_assets.json` provenance sidecar; visually review location-specific results before using them.

## Documentation Index

- [`SKILL.md`](SKILL.md) -- Complete agent-facing usage manual.
- [`CLAUDE.md`](CLAUDE.md) -- Architecture, data flow, and developer guide.
- [`references/reference-ppt-workflow.md`](references/reference-ppt-workflow.md) -- Handling user-provided reference decks.
- [`references/editing-workflow.md`](references/editing-workflow.md) -- Editing existing PPTs.
- [`references/engine-api.md`](references/engine-api.md) -- Generation API and template profiles.
- [`references/quality-checklist.md`](references/quality-checklist.md) -- Pre-delivery quality checklist.
- [`experiences/layout-pitfalls.md`](experiences/layout-pitfalls.md) -- Common layout pitfalls and solutions.
- [`experiences/chart-limits.md`](experiences/chart-limits.md) -- Chart type limits and workarounds.
- [`experiences/cjk-issues.md`](experiences/cjk-issues.md) -- CJK font and text issues.
- [`experiences/overflow-rules.md`](experiences/overflow-rules.md) -- Text overflow rules and pagination.
- [`experiences/rendering-quirks.md`](experiences/rendering-quirks.md) -- Rendering engine quirks.
- [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md) -- Third-party dependencies and licenses.

## License

Core path is MIT-style licensed (see repository LICENSE). When using optional backends, comply with their respective licenses -- especially PyMuPDF's AGPL/commercial terms.

---

*Generated for pptx-skill v6.0.0.*
