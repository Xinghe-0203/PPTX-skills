# Third-Party Notices

This project uses the following open-source packages and libraries. The core
PPTX generation path depends only on `python-pptx` and `Pillow`; everything
else is an optional backend and should be treated as such when evaluating
license obligations.

## Core dependencies

### python-pptx

- Repository: https://github.com/scanny/python-pptx
- Version range: `>=1.0,<2`
- License: MIT
- Usage: Programmatic reading and writing of `.pptx` files.

### Pillow

- Repository: https://github.com/python-pillow/Pillow
- Version range: `>=12,<13`
- License: HPND
- Usage: Font metrics, image analysis and raster QA.

## Optional adaptive layout backend

### Kiwi

- Repository: https://github.com/nucleic/kiwi
- Version range: `>=1.4,<2` (`kiwisolver`)
- License: BSD-3-Clause
- Usage: Linear constraint solver for the adaptive layout engine.

## Optional JSON schema validation

### Pydantic

- Repository: https://github.com/pydantic/pydantic
- Version range: `>=2,<3`
- License: MIT
- Usage: Declarative validation of `TemplateProfileV2` documents.

## Optional rendered-image QA backends

### NumPy

- Repository: https://github.com/numpy/numpy
- Version range: `>=2,<3`
- License: BSD-3-Clause
- Usage: Numerical arrays for image-based QA metrics.

### scikit-image

- Repository: https://github.com/scikit-image/scikit-image
- Version range: `>=0.25,<1`
- License: BSD-3-Clause
- Usage: Structural similarity and window-density QA helpers.

### OpenCV (headless)

- Repository: https://github.com/opencv/opencv
- Python package: https://github.com/opencv/opencv-python
- Version range: `>=4.10,<5`
- License: Apache-2.0
- Usage: Optional computer-vision QA helpers.

### Shapely

- Repository: https://github.com/shapely/shapely
- Version range: `>=2,<3`
- License: BSD-3-Clause
- Usage: Optional geometry helpers for overlap and IoU checks.

### pytesseract

- Repository: https://github.com/madmaze/pytesseract
- Version range: `>=0.3,<1`
- License: Apache-2.0
- Usage: OCR wrapper for rendered-slide text QA.

## Optional font/text backends

### fonttools

- Repository: https://github.com/fonttools/fonttools
- Version range: `>=4,<5`
- License: MIT
- Usage: Advanced font introspection.

### uharfbuzz

- Repository: https://github.com/harfbuzz/uharfbuzz
- Version range: `>=0.50,<1`
- License: Apache-2.0
- Usage: Optional complex-script text shaping.

## Optional render backends

### PyMuPDF (fitz)

- Repository: https://github.com/pymupdf/PyMuPDF
- Version range: `>=1.28,<2`
- License: AGPL-3.0 or commercial
- Usage: Optional PDF/PPTX rasterizer for rendered QA and golden previews.
  This is **not** a core dependency; projects redistributing the skill must
  independently choose a compliant license path before enabling this backend.

### pdf2image + Poppler

- pdf2image repository: https://github.com/Belval/pdf2image
- Version range: `>=1.17,<2`
- License: MIT
- Usage: Poppler wrapper for PDF rendering.
- Note: The system must provide Poppler utilities (e.g. `pdftoppm`) separately;
  their license terms are independent of this Python package.

### pywin32 (COM renderer)

- Repository: https://github.com/mhammond/pywin32
- Version range: `>=306,<400`
- License: PSF-2.0
- Usage: Optional PowerPoint COM automation on Windows.

## Test dependencies

### pytest

- Repository: https://github.com/pytest-dev/pytest
- Version range: `>=8,<10`
- License: MIT

### Hypothesis

- Repository: https://github.com/HypothesisWorks/hypothesis
- Version range: `>=6,<7`
- License: MPL-2.0

## Design inspirations and algorithms

The following resources informed the design of this project but no source code
was copied directly unless noted in the module-level comments:

- Pixelmatch anti-aliased perceptual diff: concept from `mapbox/pixelmatch`
  (ISC license). The implementation in `pptx_skill/render_qa.py` is an
  independent Python port; the research commit referenced during design was
  `pixelmatch@7.2.0`.
- Window-density and SSIM QA metrics follow widely published computer-vision
  formulas; scikit-image is used only as an optional backend.

## Attribution

When redistributing this skill, keep the license and NOTICE files of any
optional backend you actually ship. Core redistributions need only preserve the
MIT/HPND notices of `python-pptx` and `Pillow`.
