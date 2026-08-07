# AGENTS.md — pptx-skill

## Setup

```powershell
python -m pip install -e .                    # core deps: python-pptx, Pillow
python -m pip install -e ".[adaptive,qa-image,render-pdf]"  # full pipeline
python -m pptx_skill.capability               # check what's available
```

Requires Python ≥ 3.10. The package adds `scripts/` to `sys.path` at import time — do not run scripts directly as modules (`python -m scripts.pptx_helper` will fail); use `python scripts/pptx_helper.py` or import from `pptx_skill`.

## Testing

```powershell
python -m pytest tests/ -x -q                 # full suite (217 tests; slow — includes rendering)
python -m pytest tests/test_content_model.py  # single file
python -m pytest tests/ -k "test_layout" -x   # by keyword
```

- Tests use `unittest.TestCase`; `pytest` is the runner (`python -m unittest discover -s tests` works without pytest installed).
- Rendering tests (pipeline, golden renders) are slow and need LibreOffice. Run fast unit tests first when iterating.
- Baseline: 216 passed + 1 skipped (one preview test skips when an engine is unavailable).

## Architecture: dual-layer

This repo has **two implementation layers** that coexist:

1. **`scripts/`** — the original monolithic scripts (`pptx_helper.py`, `ppt_edit.py`, etc.). These are importable because `pptx_skill/__init__.py` prepends `scripts/` to `sys.path`.
2. **`pptx_skill/`** — the new modular package with data models (`content_model.py`), constraint solver (`layout_engine.py`), QA engines, etc.

`pptx_skill.api.auto_generate_ppt` is the public facade. It delegates to the legacy `scripts/pptx_helper.py` generator by default and writes a V3 manifest. With `layout_engine="adaptive"` it dispatches to `run_generation_pipeline` directly (adaptive constraint-based layout).

## Key entrypoints

| What | Where |
|---|---|
| Generate a PPT | `from pptx_skill import auto_generate_ppt` |
| Validate structure | `from pptx_skill import auto_validate_ppt` |
| Full adaptive pipeline | `from pptx_skill import run_generation_pipeline` |
| List/generate templates | `python scripts/template_engine.py list\|generate\|preview` |
| Render slides to PNG | `python scripts/render_slides.py <file> -o <dir> --dpi 150` |
| Analyze a reference PPT | `python scripts/reference_ppt.py analyze <file>` |
| Edit text/image/color | `from pptx_skill.ppt_edit import edit_text, swap_image, recolor` |
| Page ops (insert/delete/move) | `from pptx_skill.ppt_pages import insert_slide, delete_slide, move_slide` |
| Round-trip edit generated PPT | `from pptx_skill import load_project, edit_section, regenerate` |

## Templates

12 built-in profiles stored in `assets/templates/catalog.json`. Generated profiles go to `assets/templates/generated/`. Three layout families: `editorial_grid`, `technical_axis`, `poster_column` — avoid `standard` unless the older card system is explicitly wanted.

## Rendering

- **LibreOffice** is the primary renderer (headless soffice). Windows can fall back to PowerPoint COM if `pywin32` is installed.
- `render_slides.py --engine auto` tries LibreOffice first.
- Do not install LibreOffice unconditionally — check capability first, install only if missing.

## Environment

- `.env` file holds `PIXABAY_API_KEY` for image search (optional; generation degrades gracefully without it).
- Output dirs `output/`, `slides_preview/`, `ppt_images/` are gitignored — generated artifacts go there.
- All edits auto-create a `deck.bak.pptx` backup; call `restore_backup()` on failure.

## Visual design rules (non-obvious)

- No rounded cards, decorative circles, gradient glows, floating shadows, `VS` badges, oversized quote marks, or meaningless English kickers.
- Max 3 primary colors per theme. Titles must have clear size jump over body text.
- 20–30% whitespace. Not every page should be a bullet list.
- Multiple generated template profiles must differ in **geometry and reading order**, not just palette.

## Data model

`ContentSpec` → `SlideSpec` → `ElementSpec` with stable `element_id`s. Manifest V3 is embedded in the PPTX XML and also written as a `.manifest.json` sidecar. Slide/image indices are one-based; `0` is accepted as an alias for the first item.

## Workflow docs

Task-specific detail lives in `references/`:
- `reference-ppt-workflow.md` — when the user provides a reference PPT
- `editing-workflow.md` — when modifying an existing PPT
- `engine-api.md` — generation API and template profile schema
- `quality-checklist.md` — pre-delivery checks

## Linting

```powershell
ruff check pptx_skill/ scripts/ tests/
```

Config in `pyproject.toml`: line-length 120, target Python 3.10, E501 ignored.
