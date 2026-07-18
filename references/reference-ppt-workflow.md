# User-provided PPTX workflow

## Contents

1. Analyze the reference
2. Choose a reuse mode
3. Native master mode
4. Exact exemplar clone mode
5. Visual rebuild mode
6. Shape-level plan schema
7. Fidelity rules

## Analyze the reference

Render and analyze before editing:

```powershell
python scripts/render_slides.py reference.pptx `
  --output output/reference-preview --dpi 150

python scripts/reference_ppt.py analyze reference.pptx `
  --output output/reference-analysis.json
```

Read the analysis JSON and inspect all preview PNGs. The analysis contains slide size, masters, layouts, placeholder geometry, slide roles, shape names, font inventory, color inventory, and a recommended mode.

## Choose a reuse mode

- Use `native` when the deck contains meaningful custom masters/layouts and placeholders. This produces new slides from those layouts and inherits theme formatting.
- Use `clone` when the visible design lives on exemplar slides. This duplicates selected pages, copies media/chart relationships, and replaces named shapes.
- Use `visual-rebuild` when the reference is flattened, corrupt, dominated by unsupported SmartArt/OLE, or has no reusable semantic structure. Rebuild editable elements after inspecting the rendered pages.

Prefer `clone` when exact visual similarity matters more than future layout switching. Prefer `native` when the organization maintains a proper corporate template.

## Native master mode

Prepare structured content JSON:

```json
{
  "title": "年度经营复盘",
  "subtitle": "FY2026",
  "sections": [
    {"title": "核心结论", "bullets": ["结论一", "结论二"]},
    {"title": "关键指标", "metrics": [{"label": "收入", "value": "¥8.2B"}]}
  ]
}
```

Then run:

```powershell
python scripts/reference_ppt.py generate reference.pptx `
  --content content.json --mode native --output output/native-report.pptx
```

## Exact exemplar clone mode

For automatic binding:

```powershell
python scripts/reference_ppt.py generate reference.pptx `
  --content content.json --mode clone --output output/cloned-report.pptx
```

For high fidelity, build an explicit plan from the shape names in `reference-analysis.json`:

```powershell
python scripts/reference_ppt.py compose reference.pptx `
  --plan plan.json --output output/exact-report.pptx
```

## Visual rebuild mode

1. Render every reference slide.
2. Select representative slides for cover, section, content, data, image, and ending roles.
3. Extract design tokens with `reference_ppt.py profile`.
4. Rebuild required layouts using `pptx_helper.py` or direct `python-pptx` shapes.
5. Keep every element editable unless the source itself is a flattened image.
6. Render the result beside the reference and iterate.

Extract a reusable profile:

```powershell
python scripts/reference_ppt.py profile reference.pptx `
  --name "Client Brand" --output output/client-brand.json
```

## Shape-level plan schema

Slide indices are one-based. Reference shapes by exact `name`, `shape_id`, or one-based shape index.

```json
{
  "remove_source_slides": true,
  "slides": [
    {
      "source_slide": 1,
      "replacements": {
        "Title 1": {"text": "新的封面标题"},
        "Subtitle 2": {"text": "新的副标题"},
        "Picture 5": {"image": "C:/absolute/path/hero.jpg"},
        "TextBox 8": {"bullets": ["要点一", "要点二"]},
        "Legacy Label": {"delete": true}
      }
    },
    {
      "source_slide": 4,
      "content": {
        "title": "自动绑定页面",
        "bullets": ["系统选择主要标题和正文框", "其余内容占位符可清空"]
      },
      "clear_unmapped_content": true
    }
  ]
}
```

Use explicit replacements for complex dashboards, diagrams, and slides with multiple text regions. Automatic binding is intended for simple cover, title-content, and image-content exemplars.

## Fidelity rules

- Do not assume the theme colors describe the visible design; many decks use slide-level shapes.
- Do not delete the reference slides before cloning; cloned media relationships must be created first.
- Preserve slide size and aspect ratio.
- Use absolute image paths.
- Replace content by shape name when exact placement matters.
- Re-render every generated page after text replacement because different text length can overflow.
- Treat animations, macros, OLE objects, embedded workbooks, and complex SmartArt as special cases. Preserve them by cloning where possible and verify in PowerPoint.
