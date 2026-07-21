# Editing existing PPTX files

## Contents

1. Safety and indexing
2. Round-trip edits for generated decks
3. Direct edits for external decks
4. Rich text editing
5. Media insertion
6. Page operations
7. Template download and import
8. Restore and verify

## Safety and indexing

All editing and page-operation functions create one overwrite-style backup named `deck.bak.pptx` before saving. Slide and image indices are one-based; `0` is accepted as an alias for the first item.

Render and inspect before and after every non-trivial edit. Use `reference_ppt.py analyze` to identify exact slide and shape targets.

## Round-trip edits for generated decks

Decks created by `auto_generate_ppt()` contain an embedded project manifest and a readable `.manifest.json` sidecar. Load and regenerate them without reverse-engineering the slide shapes:

```python
from pptx_skill import load_project, edit_section, regenerate

project = load_project("output/report.pptx")
edit_section(
    "output/report.pptx",
    2,
    {"title": "更新后的标题", "bullets": ["新要点一", "新要点二"]},
)
```

Use `regenerate(project, output_path)` after changing multiple section dictionaries in memory.

## Direct edits for external decks

Replace text while preserving run formatting where possible:

```powershell
python scripts/ppt_edit.py text deck.pptx `
  --slide 3 --find "旧文本" --replace "新文本"
```

Replace a semantic text region:

```powershell
python scripts/ppt_edit.py role deck.pptx `
  --slide 3 --role title --text "新标题"
```

Replace an image while preserving its bounding box:

```powershell
python scripts/ppt_edit.py image deck.pptx `
  --slide 4 --image 1 --path C:/absolute/path/new-image.png
```

Replace one exact color globally:

```powershell
python scripts/ppt_edit.py recolor deck.pptx `
  --old "#17324D" --new "#244A3D"
```

Switch between legacy built-in themes only when the old theme is recognized reliably:

```powershell
python scripts/ppt_edit.py theme deck.pptx --new forest_luxe
```

If theme recognition returns an error, inspect the color inventory and use explicit `recolor` mappings. Do not force a guessed theme onto an unrelated external deck.

## Rich text editing

The adaptive pipeline supports multi-run rich text in text nodes. Each run dict can specify `bold`, `italic`, `color` (hex), `size` (pt), and `href` (hyperlink URL). Paragraph-level formatting supports `align` (left/center/right/justify), `level` (indent level, max 8), and `bullet` (custom bullet character).

For round-trip edits on generated decks, modify the section content and regenerate. For direct edits on external decks, use `ppt_edit.py text` or `role` commands, which preserve existing run formatting where possible.

## Media insertion

Video and audio media are supported in the adaptive pipeline as element kinds `video` and `audio` in `SlideSpec.elements`. Specify the file path via `content["path"]`. Video supports an optional `content["poster"]` for a poster frame image. Audio supports an optional `content["mime_type"]` (default `"audio/mpeg"`).

For direct edits on external decks, media insertion is not yet available in `ppt_edit.py`. Use `python-pptx` directly via `slide.shapes.add_movie()` for video, or `slide.shapes.add_movie()` with an audio MIME type for audio.

## Page operations

Insert a page:

```powershell
python scripts/ppt_pages.py insert deck.pptx `
  --index 3 `
  --layout dashboard `
  --section '{"title":"新增数据页","metrics":[{"label":"收入","value":"¥8.2B"}]}'
```

Delete, move, or duplicate:

```powershell
python scripts/ppt_pages.py delete deck.pptx --index 5
python scripts/ppt_pages.py move deck.pptx --from 7 --to 3
python scripts/ppt_pages.py duplicate deck.pptx --source 2 --target 6
```

Redraw one page in another layout:

```powershell
python scripts/ppt_pages.py replace-layout deck.pptx `
  --index 4 --layout timeline `
  --section '{"title":"里程碑","events":[{"date":"Q1","title":"启动"}]}'
```

Supply an explicit section JSON for complex pages. Automatic extraction only recovers a basic title and body.

## Template download and import

Download template packs from GitHub, direct URLs, or local directories, then import them as reusable profiles:

```python
from pptx_skill import download_template_pack, import_template

# Download from GitHub
result = download_template_pack("github", "output/templates", repo="user/repo")

# Import a .pptx as a template profile
profile = import_template("output/templates/report.pptx", name="Client Brand")
```

Imported profiles are saved to `assets/templates/generated/` and can be used by their profile ID as a `template_key`.

## Restore and verify

Restore the most recent backup:

```python
from pptx_skill import restore_backup
restore_backup("deck.pptx")
```

After editing:

1. Reopen the file with `python-pptx`.
2. Render all slides.
3. Verify page order and requested changes.
4. Inspect untouched pages for regressions.
5. Open in PowerPoint when the deck contains animation, OLE, SmartArt, or embedded workbooks.
