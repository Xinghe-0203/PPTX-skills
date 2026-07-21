# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Subagent Policy

Prefer delegating work to subagents. Up to 20 agents may run concurrently. Use parallel agents for independent tasks (search, review, verification) and pipeline/sequential agents for dependent work. Always use the Agent tool rather than doing everything inline.

## Build & Test Commands

```bash
# Install (editable, core only)
python -m pip install -e .

# Install with optional backends
python -m pip install -e ".[adaptive,schema,qa-image,render-pdf]"

# Run full test suite (185 tests, ~300s)
python -m pytest tests/ -ra

# Run a single test file
pytest tests/test_layout_engine.py

# Run a single test method
pytest "tests/test_layout_engine.py::LayoutEngineTests::test_solve_bullets_recipe"

# Check environment capabilities
python -m pptx_skill.capability

# Lint
ruff check .

# Type check
mypy pptx_skill/
```

All tests use `unittest.TestCase` (no pytest fixtures or conftest). CJK content is common in fixtures.

## Architecture: Dual-Layer System

This project has two coexisting layers connected by `sys.path` manipulation:

```
pptx_skill/__init__.py
  └─ inserts scripts/ onto sys.path
  └─ re-exports legacy scripts as pptx_skill.* namespace
```

**Legacy layer** (`scripts/`) — the original monolithic engine, still the primary generation path:
- `pptx_helper.py` — generation engine, `auto_generate_ppt()`, `Section` dataclass, `LAYOUT_REGISTRY`, `THEMES`
- `ppt_project.py` — V2 manifest round-trip (`load_project`, `edit_section`, `regenerate`)
- `ppt_edit.py` — L3 in-place edits (`edit_text`, `swap_image`, `recolor`, `swap_theme`)
- `ppt_pages.py` — L4 page operations (`insert_slide`, `delete_slide`, `move_slide`, `replace_layout`)
- `ppt_inspect.py` — L2 external PPT inspection (`inspect_ppt`)
- `reference_ppt.py` — foundational adapter: analyze, clone, compose, native-mode generation
- `template_engine.py` — template catalog & profile generator
- `layout_variants.py` — three non-card layout families (editorial_grid, technical_axis, poster_column)
- `pixabay_search.py` — Pixabay image search (requires `PIXABAY_API_KEY` env var)
- `render_slides.py` — thin CLI, delegates to `pptx_skill.preview_renderer`

**Modular layer** (`pptx_skill/`) — the new package, PR1–PR9:
- `api.py` — compatibility facade wrapping both layers
- `content_model.py` — zero-dependency core: `ContentSpec`, `SlideSpec`, `ElementSpec`, `LayoutPlan`, `PlannedNode`, `GeometrySpec`, `BBox`, `CanvasSpec`, `StableIdGenerator`
- `content_adapter.py` — adapt legacy `Section` objects to `ContentSpec`
- `layout_engine.py` — declarative `LayoutRecipe` + Kiwi constraint solver → `SolvedGeometry`
- `deck_planner.py` — beam search over candidate bundles with transition penalties
- `pagination.py` — role-specific paginators (bullets, table, timeline, process, image_grid)
- `pptx_renderer.py` — adaptive renderer dispatching text/image/shape/table/chart nodes
- `text_metrics.py` — Pillow font metrics, CJK line-breaking, font-size bisection
- `semantic_qa.py` — overflow/overlap/contrast/font/distortion checks
- `render_qa.py` — pixel-level perceptual diff, SSIM, window density
- `repair_engine.py` — whitelist repair actions with profile override merging
- `generation_pipeline.py` — plan → render → QA → repair closed loop
- `manifest.py` — V3 manifest (embedded XML + sidecar JSON) with V2→V3 migration
- `design_schema.py` — V2 template schema, three-layer tokens, resolution
- `template_compiler.py` — TemplateProfileV2 compilation with diversity gate
- `template_v2_adapter.py` — V1→V2 migration, V2→legacy adaptation
- `reference_adapter.py` — native/clone/visual-rebuild adapter dispatch
- `visual_rebuild.py` — flattened-deck parsing and adaptive rebuild
- `qa_dataset.py` — 200-page annotation dataset generator
- `golden_renders.py` — per-role golden renders and family decks
- `image_crop.py` — contain/cover/smart crop
- `preview_renderer.py` — LibreOffice / PyMuPDF / Windows COM rendering
- `template_downloader.py` — template pack download (GitHub/URL/local), PPTX archive extraction, remote pack listing & search
- `visual_qa.py` — visual QA checks on rendered slides
- `capability.py` — runtime environment capability detection and reporting
- `transitions.py` — 18 slide transition types (fade/push/wipe/cover/split/dissolve/random/cut)
- `animations.py` — 51 animation types (20 entrance, 15 exit, 15 emphasis, 1 motion path) via OOXML timing XML
- `merge.py` — multi-deck merge with layout/media/rel-ID deduplication, slide extraction
- `watermark.py` — text/image watermarks with opacity, tiling, z-ordering, removal
- `sections.py` — section groups (add/remove/rename/move/collapse) via OOXML sectionLst
- `slide_master.py` — slide master/layout query, clone, rename, placeholder ops, background settings
- `image_optimize.py` — image compression, format conversion, unused media removal, stats
- `comments.py` — speaker notes + review comments (add/list/delete/reply/resolve) via OOXML
- `diff.py` — structural diff at slide/shape/text/image/layout level with shape matching
- `export.py` — PDF/image/HTML/text/thumbnail export (LibreOffice/PyMuPDF/Pillow backends)
- `smartart.py` — SmartArt detection, text extraction/editing, preservation across clone/merge
- `vba.py` — VBA project detection, injection, extraction, macro-to-shape attachment
- `table_styles.py` — 20 built-in table styles, cell merge/unmerge, borders, fills, add row/column
- `html_import.py` — HTML-to-PPTX import with local snapshot underlay technique

## Data Flow: Content → Rendered PPTX

```
ContentSpec / SlideSpec / ElementSpec          (content_model.py)
        │
        ▼
adapt_legacy_sections() or direct construction (content_adapter.py)
        │
        ▼
paginate_content_spec() → split dense slides   (pagination.py)
        │
        ▼
SemanticQAEngine.check() → pre-render checks  (semantic_qa.py)
        │      (overflow / overlap / contrast / font / distortion)
        ▼
plan_deck() → beam search                      (deck_planner.py)
  ├─ builtin_recipes(role) → LayoutRecipe[]    (layout_engine.py)
  ├─ solve_recipe(recipe, canvas, tokens)      (layout_engine.py, uses kiwisolver)
  │      → SolvedGeometry
  └─ solved_geometry_to_layout_plan()
        │      → LayoutPlan with PlannedNode[]
        ▼
render_layout_plans(plans, output_path, deck_options)  (pptx_renderer.py)
        │      → RenderResult with RenderTraceEntry[]
        │      deck_options: show_slide_numbers, footer_text, header_text, theme
        ▼
render_preview() → PNG via LibreOffice/COM     (preview_renderer.py)
        │
        ▼
render_qa → pixel-level perceptual diff, SSIM  (render_qa.py)
        │
        ▼  (if QA fails)
propose_repairs() → apply_repairs()            (repair_engine.py)
        │      → loop back to plan_deck (max 2 passes)
        ▼
save_manifest_v3()                             (manifest.py)
```

Canvas: 16:9 ≈ 959.976 × 540 pt. Point-based coordinates throughout. Tokens provide semantic spacing/sizing values; `_builtin_tokens()` is the standard profile.

## Manifest System

Two manifest versions coexist:
- **V2** (`scripts/ppt_project.py`): written by `auto_generate_ppt()`, flat `sections`/`layouts` payload, namespace `urn:openai:pptx-skill:manifest:v2`
- **V3** (`pptx_skill/manifest.py`): written by `run_generation_pipeline()`, structured `current.content`/`layout_plans`/`render_trace` + `attempts` history, namespace `urn:openai:pptx-skill:manifest:v3`

Both write to `customXml/pptxSkillManifest.xml` inside the pptx zip plus a `.manifest.json` sidecar fallback. `load_manifest()` auto-detects and migrates V2→V3 in memory.

## Editing Paths

| Source | Method | Fidelity |
|---|---|---|
| Skill-generated PPT (has manifest) | `load_project` → `edit_section` → `regenerate` | Lossless |
| Skill-generated PPT (has manifest) | `insert_slide` / `delete_slide` / `move_slide` / `replace_layout` | Lossless for new pages |
| External PPT | `inspect_ppt()` → understand structure | Read-only |
| External PPT | `edit_text` / `swap_image` / `recolor` / `swap_theme` | Best-effort, may break formatting |

All editing functions create `.bak.pptx` backup before writing. Slide indices are 1-based in public APIs.

## Layout Recipes & Roles

24 roles: `cover, toc, section, bullets, text_image, full_image, image_grid, dashboard, timeline, comparison, quote, process, table, end, matrix, kpi_hero, faq, testimonial, logo_wall, swot, porter, pest, bmc, funnel`

~57 recipe variants with density-graded options (sparse/dense for bullets, dashboard; three_column for comparison; vertical_dense for timeline; horizontal_dense for process; grid4 for image_grid; wide for table). Analytical framework roles: swot (quadrant/labeled), porter (diamond/horizontal), pest (grid/vertical), bmc (canvas/compact), funnel (stacked/narrow). Each role has ≥2 variants.

When adding a new recipe:
1. Define zones, constraints, content_limits in `layout_engine.py`
2. Ensure no redundant constraints (e.g. both `==` and `>=` on the same gap causes `UnsatisfiableConstraint`)
3. Add role-specific paginator in `pagination.py` if content can overflow
4. Add golden render coverage in `golden_renders.py`

## Key Design Constraints

- **Deterministic**: same recipe + canvas + tokens → identical geometry. Stable element IDs via `StableIdGenerator`. Tests assert this.
- **Lazy optional imports**: kiwisolver, numpy, PyMuPDF, fonttools, etc. are imported inside functions, never at module top-level. Core path needs only python-pptx + Pillow.
- **No AI-style decoration**: no rounded cards, decorative circles, gradient glows, floating shadows, VS badges, oversized quote marks, meaningless English kickers. Use grids, fine rules, axes, whitespace, and real information hierarchy.
- **AGPL risk**: `PyMuPDF` (render-pdf extra) is AGPL-3.0/commercial. Never make it a core dependency. See `THIRD_PARTY_NOTICES.md`.
- **CJK-first**: font metrics, line-breaking, and continuation markers (`（续）`) target Chinese content. Tests use CJK strings.
- **Anti-repeat layout**: `avoid_adjacent_same_signature` in deck_planner computes a geometry signature (role order + normalized bboxes + area ratios) for each plan; when two adjacent plans exceed the `adjacent_signature_threshold` (0.85 by default), a penalty is added to the beam search score, steering the planner away from consecutively similar layouts.
- **20 themes** with font pairing (FONTS + CJK_FONTS dicts in pptx_helper.py).
- **20 template profiles** in catalog (12 built-in + 8 new).
- **10 chart types** supported (column, bar, line, pie, doughnut, scatter, area variants).
- **Rich text**: multi-paragraph, multi-run, bullets, hyperlinks in text nodes.
- **Table styling**: borders, banding, cell fills, column widths.
- **Shape styling**: line/stroke, shadow, gradient, rotation, text inside shapes.
- **Template download**: `download_template_pack()`, `import_template()`, `search_github_templates()`.

## Test Patterns

- No conftest, no shared fixtures. Each file is self-contained with local `_make_*` helpers.
- `tempfile.TemporaryDirectory()` for all I/O tests.
- `canvas_from_name("16:9")` + `_builtin_tokens()` is the standard test profile.
- `solve_recipe(recipe, canvas, tokens)` → `solved_geometry_to_layout_plan(slide, recipe, geom, canvas, tokens)` is the canonical two-step for creating a `LayoutPlan`.
- `render_layout_plans(plans, path)` → open with `Presentation(path)` to assert slide count/shape properties.
- PR-tagged files (`test_pr8a_*`, `test_pr8b_*`, `test_pr8c_*`, `test_pr9_*`) contain multiple TestCase classes for that PR's acceptance criteria.
- Always use `PYTHONIOENCODING=utf-8` on Windows to avoid GBK encoding issues with CJK test output.
