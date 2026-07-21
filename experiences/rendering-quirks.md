# Rendering Quirks

## TL;DR

Rendering PowerPoint content to images/PDF involves multiple backends with different trade-offs. PyMuPDF is AGPL-3.0 and must never be a core dependency. LibreOffice has sub-pixel positioning differences. Windows COM gives best fidelity but requires PowerPoint. OOXML validity does not guarantee PowerPoint playback. Animation and SmartArt have limited empirical support.

---

## Rule 1: PyMuPDF is AGPL-3.0 — never make it a core dependency

PyMuPDF (the `fitz` package) is licensed under AGPL-3.0, which requires open-sourcing any software that links to it. This is incompatible with most commercial use.

**Allowed:**
- Optional dependency in the `render-pdf` extra: `pip install pptx_skill[render-pdf]`
- Lazy import inside functions, never at module top-level
- Feature detection via `capability.py` that gracefully degrades when absent

**Forbidden:**
- Import at module top-level
- Required dependency in `setup.cfg` / `pyproject.toml` core dependencies
- Any code path that fails without PyMuPDF installed

**Pattern:**
```python
def render_with_pymupdf(pptx_path, output_dir):
    """Render slides using PyMuPDF. Requires render-pdf extra."""
    try:
        import fitz  # lazy import
    except ImportError:
        raise ImportError(
            "PyMuPDF is required for PDF rendering. "
            "Install with: pip install pptx_skill[render-pdf]"
        )
    doc = fitz.open(pptx_path)
    ...
```

---

## Rule 2: LibreOffice rendering may differ from PowerPoint by +/-2px

LibreOffice's layout engine is not identical to PowerPoint's. Known differences:

- **Text positioning**: vertical baseline offset can differ by 1-2px
- **Font substitution**: if a font is missing, LibreOffice substitutes differently than PowerPoint
- **Shape rendering**: gradient angles and shadow offsets may vary slightly
- **Table borders**: border collapse rules differ between the two engines

**Implications for visual QA:**
- The `render_qa.py` module uses SSIM with a tolerance threshold (default 0.95) to account for these differences
- Pixel-perfect comparison between LibreOffice renders and PowerPoint renders will always show differences
- Use structural QA (`semantic_qa.py`) for layout correctness; use visual QA (`render_qa.py`) only for regression detection

---

## Rule 3: Windows COM automation gives best fidelity but requires PowerPoint

The Windows COM backend (via `win32com.client`) drives the actual PowerPoint application, producing renders identical to what a user sees. This is the gold standard for fidelity.

**Requirements:**
- Windows OS: Windows only
- PowerPoint must be installed (Office 2013+)
- `pywin32` package must be installed
- PowerPoint process runs in background (can conflict with user's open instances)

**Usage pattern:**
```python
def render_with_com(pptx_path, output_dir):
    """Render using PowerPoint COM automation. Windows only."""
    import win32com.client
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(pptx_path)
    for i, slide in enumerate(presentation.Slides):
        slide.Export(f"{output_dir}/slide_{i+1}.png", "PNG")
    presentation.Close()
    powerpoint.Quit()
```

The `preview_renderer.py` module tries backends in order: COM (Windows) -> LibreOffice -> PyMuPDF -> Pillow fallback.

---

## Rule 4: PNG transparency in watermarks needs alphaModFix on blipFill

When applying transparent PNG watermarks, the transparency must be set on the `blipFill` element using `alphaModFix`, not on a `solidFill` alpha. Using `solidFill` alpha creates a colored overlay, not a transparent image.

**Bad:**
```python
# This creates a semi-transparent colored rectangle, not a transparent image
fill = shape.fill
fill.solid()
fill.fore_color.rgb = RGBColor(255, 255, 255)
# Setting alpha on solidFill does not make the image transparent
```

**Good:**
```python
# Set alpha on the blip (image) fill
from lxml import etree

blip = shape.fill._fill.find(
    ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
)
alpha_mod = etree.SubElement(blip, qn("a:alphaModFix"))
alpha_mod.set("amt", str(int(opacity * 1000)))  # opacity 0.3 → amt="300"
```

The `watermark.py` module handles this correctly with `set_blip_alpha()`.

---

## Rule 5: OOXML spec compliance does not guarantee PowerPoint playback

The OOXML specification (ISO/IEC 29500) defines what is valid XML, but PowerPoint has its own interpretation. Many valid OOXML constructs are silently dropped or ignored by PowerPoint:

- **Custom geometry paths**: valid `<a:path>` elements may not render if they use unsupported commands
- **Animation timing**: valid `<p:timing>` XML may be ignored if the sequence doesn't match PowerPoint's expected structure
- **Theme overrides**: valid `<a:fmtScheme>` overrides may be ignored if they conflict with the slide master
- **Table styles**: custom `<a:tblStyle>` entries are ignored; PowerPoint only uses its built-in table style gallery

**Practical rule:** Always validate against PowerPoint playback, not just OOXML schema validation. The `semantic_qa.py` module checks for empirically known PowerPoint quirks.

---

## Rule 6: Animation filters — only 16 entrance/exit filters are empirically verified

The `animations.py` module defines 51 animation types (20 entrance, 15 exit, 15 emphasis, 1 motion path), but not all of these play correctly in PowerPoint slideshow mode. Empirically verified filters:

**Entrance (verified):**
- Appear, Fade, Fly In, Float Up, Float Down, Float Left, Float Right, Wipe, Box, Split, Wheel, Random Bars, Grow & Turn, Swivel, Bounce, Zoom

**Exit (verified):**
- Fade, Fly Out, Float Up, Float Down, Float Left, Float Right, Wipe, Box, Split, Shrink & Turn, Swivel, Bounce, Random Bars, Collapse, Disappear

**Emphasis (partially verified):**
- Grow/Shrink, Spin, Transparency, Color Change, Bold Flash

The remaining emphasis animations and the motion path animation generate valid OOXML but may not play in all PowerPoint versions. Always test animations in the target PowerPoint version before relying on them.

---

## Rule 7: SmartArt preservation requires copying all 4 diagram parts

SmartArt diagrams consist of four separate parts in the OOXML package:

1. **Data** (`/ppt/diagrams/data1.xml`) — the text content
2. **Layout** (`/ppt/diagrams/layout1.xml`) — the layout algorithm definition
3. **Colors** (`/ppt/diagrams/colors1.xml`) — the color scheme
4. **Quick Style** (`/ppt/diagrams/quickStyle1.xml`) — the visual style

When cloning or merging slides that contain SmartArt, all four parts must be copied along with their relationships. Missing any part causes:
- PowerPoint repair prompt on open
- SmartArt converted to a static image (loss of editability)
- Complete SmartArt disappearance

**The `smartart.py` module provides:**
- `detect_smartart(slide)` — find all SmartArt shapes on a slide
- `extract_smartart_parts(shape, pptx_package)` — collect all 4 parts + relationships
- `preserve_smartart(source_shape, target_slide, source_pkg, target_pkg)` — copy with full part preservation

**The `merge.py` module** uses SmartArt preservation automatically when merging decks that contain SmartArt diagrams.
