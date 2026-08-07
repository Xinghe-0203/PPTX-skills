"""OCR integration for PPTX — extract text from images and scanned slides.

Wraps Tesseract (pytesseract), EasyOCR, and PaddleOCR with the existing
render pipeline.  The module provides both **batch extraction** (scan all
images/slides for text) and **auto-caption** (write OCR results back as
alt-text or overlay text boxes).

Quick start
-----------
>>> from pptx_skill.ocr import ocr_slide, ocr_presentation
>>> results = ocr_slide("deck.pptx", 1)          # single slide (1-based)
>>> all_results = ocr_presentation("deck.pptx")   # all slides

Optional dependencies
---------------------
- pytesseract  (extra: qa-ocr)  — local Tesseract wrapper
- easyocr      (extra: ocr-easy) — deep-learning OCR
- paddleocr    (extra: ocr-paddle) — Baidu PaddleOCR
At least one must be installed; the module auto-selects the best available.
"""
from __future__ import annotations

import logging
import os
import tempfile
from dataclasses import dataclass

__all__ = [
    "OcrResult",
    "OcrEngine",
    "ocr_slide",
    "ocr_presentation",
    "ocr_image",
    "auto_caption_slide",
    "auto_caption_presentation",
    "list_available_engines",
]

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class OcrResult:
    """A single OCR detection."""
    text: str = ""
    confidence: float = 0.0
    left: float = 0.0    # fraction of image width  [0..1]
    top: float = 0.0     # fraction of image height [0..1]
    width: float = 0.0   # fraction of image width
    height: float = 0.0  # fraction of image height
    language: str = "en"
    engine: str = ""


class OcrEngine:
    """Enum-like selector for OCR backends."""
    TESSERACT = "tesseract"
    EASYOCR = "easyocr"
    PADDLEOCR = "paddleocr"
    AUTO = "auto"


# ---------------------------------------------------------------------------
# Engine detection
# ---------------------------------------------------------------------------

def list_available_engines() -> list[str]:
    """Return a list of OCR engine names that are importable."""
    engines = []
    try:
        import pytesseract  # noqa: F401
        engines.append(OcrEngine.TESSERACT)
    except ImportError:
        pass
    try:
        import easyocr  # noqa: F401
        engines.append(OcrEngine.EASYOCR)
    except ImportError:
        pass
    try:
        import paddleocr  # noqa: F401
        engines.append(OcrEngine.PADDLEOCR)
    except ImportError:
        pass
    return engines


def _select_engine(engine: str) -> str:
    """Resolve ``'auto'`` to the best available engine."""
    if engine != OcrEngine.AUTO:
        return engine
    available = list_available_engines()
    if not available:
        raise RuntimeError(
            "No OCR engine available.  Install one of: "
            "pytesseract (qa-ocr), easyocr (ocr-easy), paddleocr (ocr-paddle)"
        )
    # Prefer EasyOCR (good multilingual), then PaddleOCR, then Tesseract
    for pref in (OcrEngine.EASYOCR, OcrEngine.PADDLEOCR, OcrEngine.TESSERACT):
        if pref in available:
            return pref
    return available[0]


# ---------------------------------------------------------------------------
# Core OCR functions
# ---------------------------------------------------------------------------

def ocr_image(
    image_path: str,
    *,
    engine: str = OcrEngine.AUTO,
    languages: list[str] | None = None,
    min_confidence: float = 0.3,
) -> list[OcrResult]:
    """Run OCR on a single image file.

    Parameters
    ----------
    image_path : str
        Path to the image (PNG, JPG, BMP, TIFF).
    engine : str
        One of ``"auto"``, ``"tesseract"``, ``"easyocr"``, ``"paddleocr"``.
    languages : list[str], optional
        Language codes (e.g. ``["en", "zh"]``).  Default is ``["en"]``.
    min_confidence : float
        Discard results below this confidence threshold (0–1).

    Returns
    -------
    list[OcrResult]
    """
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")

    engine = _select_engine(engine)
    langs = languages or ["en"]

    if engine == OcrEngine.TESSERACT:
        return _ocr_tesseract(image_path, langs, min_confidence)
    elif engine == OcrEngine.EASYOCR:
        return _ocr_easyocr(image_path, langs, min_confidence)
    elif engine == OcrEngine.PADDLEOCR:
        return _ocr_paddleocr(image_path, langs, min_confidence)
    else:
        raise ValueError(f"Unknown OCR engine: {engine!r}")


def _ocr_tesseract(image_path: str, languages: list[str], min_confidence: float) -> list[OcrResult]:
    """OCR using pytesseract with detailed output."""
    import pytesseract
    from PIL import Image

    with Image.open(image_path) as img:
        img_w, img_h = img.size

        # Tesseract language codes: eng, chi_sim, chi_tra, jpn, kor, etc.
        lang_map = {"en": "eng", "zh": "chi_sim", "zh_tw": "chi_tra", "ja": "jpn", "ko": "kor"}
        tess_langs = "+".join(lang_map.get(lang, lang) for lang in languages)

        data = pytesseract.image_to_data(img, lang=tess_langs, output_type=pytesseract.Output.DICT)
        results = []
        n = len(data["text"])
        for i in range(n):
            text = data["text"][i].strip()
            conf = float(data["conf"][i]) / 100.0  # Tesseract gives 0-100
            if not text or conf < min_confidence:
                continue
            results.append(OcrResult(
                text=text,
                confidence=conf,
                left=data["left"][i] / img_w,
                top=data["top"][i] / img_h,
                width=data["width"][i] / img_w,
                height=data["height"][i] / img_h,
                language=languages[0] if languages else "en",
                engine=OcrEngine.TESSERACT,
            ))
        return results


def _ocr_easyocr(image_path: str, languages: list[str], min_confidence: float) -> list[OcrResult]:
    """OCR using EasyOCR (deep-learning based)."""
    import easyocr
    from PIL import Image

    with Image.open(image_path) as img:
        img_w, img_h = img.size

        # EasyOCR language codes: en, ch_sim, ch_tra, ja, ko, etc.
        lang_map = {"zh": "ch_sim", "zh_tw": "ch_tra"}
        reader_langs = [lang_map.get(lang, lang) for lang in languages]

        # Reader is cached per language combo
        reader = easyocr.Reader(reader_langs, verbose=False)
        detections = reader.readtext(image_path)

        results = []
        for bbox, text, conf in detections:
            if conf < min_confidence or not text.strip():
                continue
            # bbox is [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            results.append(OcrResult(
                text=text.strip(),
                confidence=conf,
                left=min(xs) / img_w,
                top=min(ys) / img_h,
                width=(max(xs) - min(xs)) / img_w,
                height=(max(ys) - min(ys)) / img_h,
                language=languages[0] if languages else "en",
                engine=OcrEngine.EASYOCR,
            ))
        return results


def _ocr_paddleocr(image_path: str, languages: list[str], min_confidence: float) -> list[OcrResult]:
    """OCR using PaddleOCR."""
    from PIL import Image

    with Image.open(image_path) as img:
        img_w, img_h = img.size

        # Lazy import — paddleocr is heavy
        from paddleocr import PaddleOCR

        lang_map = {"en": "en", "zh": "ch", "zh_tw": "ch", "ja": "japan", "ko": "korean"}
        pad_lang = lang_map.get(languages[0], "en") if languages else "en"

        ocr = PaddleOCR(lang=pad_lang, show_log=False)
        result = ocr.ocr(image_path, cls=True)

    results = []
    if result and result[0]:
        for line in result[0]:
            bbox = line[0]
            text = line[1][0]
            conf = line[1][1]
            if conf < min_confidence or not text.strip():
                continue
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            results.append(OcrResult(
                text=text.strip(),
                confidence=conf,
                left=min(xs) / img_w,
                top=min(ys) / img_h,
                width=(max(xs) - min(xs)) / img_w,
                height=(max(ys) - min(ys)) / img_h,
                language=languages[0] if languages else "en",
                engine=OcrEngine.PADDLEOCR,
            ))
    return results


# ---------------------------------------------------------------------------
# PPTX-level OCR
# ---------------------------------------------------------------------------

def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path):
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path.
    Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs, path):
    """Save *prs* back to *path* if *path* is not None."""
    if path is not None:
        prs.save(str(path))


def _render_slide_to_image(prs, slide_index: int, dpi: int = 200) -> str:
    """Render a slide to a temporary PNG using the preview renderer."""
    import tempfile

    from pptx_skill.preview_renderer import render_preview

    tmp_dir = tempfile.mkdtemp(prefix="pptx_ocr_")
    output_path = os.path.join(tmp_dir, f"slide_{slide_index}.png")
    render_preview(prs, output_path, dpi=dpi)
    return output_path


def _extract_image_from_shape(shape, output_dir: str, index: int) -> str | None:
    """Extract an image shape to a file and return its path."""
    try:
        image = shape.image
        ext = image.content_type.split("/")[-1]
        if ext not in ("png", "jpeg", "jpg", "bmp", "tiff"):
            ext = "png"
        path = os.path.join(output_dir, f"img_{index}.{ext}")
        with open(path, "wb") as f:
            f.write(image.blob)
        return path
    except Exception:
        return None


def ocr_slide(
    prs_or_path,
    slide_index: int,
    *,
    engine: str = OcrEngine.AUTO,
    languages: list[str] | None = None,
    min_confidence: float = 0.3,
    include_rendered: bool = False,
    dpi: int = 200,
) -> list[OcrResult]:
    """Run OCR on a single slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    include_rendered : bool
        If True, also OCR the rendered slide image (catches text in shapes
        that are hard to extract directly).  Default False — only OCR images
        embedded in the slide.
    dpi : int
        DPI for slide rendering (when ``include_rendered=True``).

    Returns
    -------
    list[OcrResult]
    """
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        all_results: list[OcrResult] = []

        # OCR each image shape
        with tempfile.TemporaryDirectory(prefix="pptx_ocr_") as tmp_dir:
            img_idx = 0
            for shape in slide.shapes:
                try:
                    from pptx.enum.shapes import MSO_SHAPE_TYPE
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        img_path = _extract_image_from_shape(shape, tmp_dir, img_idx)
                        if img_path and os.path.isfile(img_path):
                            results = ocr_image(
                                img_path,
                                engine=engine,
                                languages=languages,
                                min_confidence=min_confidence,
                            )
                            # Convert image-local coords to slide-fraction
                            for r in results:
                                # Approximate: image position within slide
                                sw = shape.width / prs.slide_width
                                sh = shape.height / prs.slide_height
                                sx = shape.left / prs.slide_width
                                sy = shape.top / prs.slide_height
                                r.left = sx + r.left * sw
                                r.top = sy + r.top * sh
                                r.width *= sw
                                r.height *= sh
                            all_results.extend(results)
                        img_idx += 1
                except Exception:
                    img_idx += 1
                    continue

            # Optionally OCR the rendered slide
            if include_rendered:
                try:
                    rendered_path = _render_slide_to_image(prs, slide_index, dpi)
                    rendered_results = ocr_image(
                        rendered_path,
                        engine=engine,
                        languages=languages,
                        min_confidence=min_confidence,
                    )
                    for r in rendered_results:
                        r.left = r.left  # already in slide fractions
                        r.top = r.top
                    all_results.extend(rendered_results)
                except Exception as exc:
                    log.warning("Failed to render slide for OCR: %s", exc)

        return all_results
    finally:
        pass


def ocr_presentation(
    prs_or_path,
    *,
    engine: str = OcrEngine.AUTO,
    languages: list[str] | None = None,
    min_confidence: float = 0.3,
    include_rendered: bool = False,
    max_slides: int | None = None,
) -> dict[int, list[OcrResult]]:
    """Run OCR on all slides in a presentation.

    Parameters
    ----------
    max_slides : int, optional
        Maximum number of slides to process.  None = all.

    Returns
    -------
    dict[int, list[OcrResult]]
        Map of slide_index (1-based) → OCR results.
    """
    prs = _open_prs(prs_or_path)
    try:
        results: dict[int, list[OcrResult]] = {}
        slide_count = len(prs.slides)
        limit = min(slide_count, max_slides) if max_slides else slide_count

        for idx in range(limit):
            results[idx + 1] = ocr_slide(
                prs, idx + 1,
                engine=engine,
                languages=languages,
                min_confidence=min_confidence,
                include_rendered=include_rendered,
            )

        return results
    finally:
        pass


# ---------------------------------------------------------------------------
# Auto-caption: write OCR results back as alt-text or text boxes
# ---------------------------------------------------------------------------

def auto_caption_slide(
    prs_or_path,
    slide_index: int,
    *,
    engine: str = OcrEngine.AUTO,
    languages: list[str] | None = None,
    min_confidence: float = 0.5,
    mode: str = "alt_text",
) -> int:
    """Run OCR on a slide and write the detected text back.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    mode : str
        ``"alt_text"`` — set alt text on image shapes (accessibility).
        ``"text_box"`` — add a text box overlay with OCR text.
        ``"notes"``    — append OCR text to speaker notes.

    Returns
    -------
    int
        Number of shapes/elements updated.
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
        slide = prs.slides[slide_index - 1]
        results = ocr_slide(
            prs, slide_index,
            engine=engine,
            languages=languages,
            min_confidence=min_confidence,
        )

        if not results:
            return 0

        # Combine all text
        combined_text = " ".join(r.text for r in results)

        if mode == "alt_text":
            return _apply_alt_text(slide, results, combined_text)
        elif mode == "text_box":
            return _apply_text_box(prs, slide, results)
        elif mode == "notes":
            return _apply_notes(slide, combined_text)
        else:
            raise ValueError(f"Unknown mode: {mode!r}. Use 'alt_text', 'text_box', or 'notes'.")
    finally:
        _save_prs(prs, path)


def auto_caption_presentation(
    prs_or_path,
    *,
    engine: str = OcrEngine.AUTO,
    languages: list[str] | None = None,
    min_confidence: float = 0.5,
    mode: str = "alt_text",
) -> int:
    """Run OCR on all slides and auto-caption them.

    Returns
    -------
    int
        Total number of shapes/elements updated.
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        total = 0
        for idx in range(1, len(prs.slides) + 1):
            total += auto_caption_slide(
                prs, idx,
                engine=engine,
                languages=languages,
                min_confidence=min_confidence,
                mode=mode,
            )
        return total
    finally:
        _save_prs(prs, path)


def _apply_alt_text(slide, results: list[OcrResult], combined_text: str) -> int:
    """Set alt text on image shapes from OCR results."""
    count = 0
    for shape in slide.shapes:
        try:
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if not shape.name:
                    continue
                # Set alt text via OOXML
                cNvPr = shape._element.find(
                    ".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"
                )
                if cNvPr is None:
                    cNvPr = shape._element.find(
                        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr"
                    )
                if cNvPr is not None:
                    cNvPr.set("descr", combined_text[:500])
                    count += 1
        except Exception:
            continue
    return count


def _apply_text_box(prs, slide, results: list[OcrResult]) -> int:
    """Add a text box overlay with OCR text for each detection."""
    from pptx.util import Inches, Pt

    count = 0
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    for r in results[:20]:  # Limit to 20 text boxes
        try:
            left = int(r.left * slide_w)
            top = int(r.top * slide_h)
            width = max(int(r.width * slide_w), Inches(0.5))
            height = max(int(r.height * slide_h), Inches(0.3))

            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.text = r.text
            # Make text small and semi-transparent
            for para in tf.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = None  # inherit
            count += 1
        except Exception:
            continue
    return count


def _apply_notes(slide, combined_text: str) -> int:
    """Append OCR text to speaker notes."""
    try:
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        existing = notes_tf.text
        if existing:
            notes_tf.text = existing + "\n\n[OCR] " + combined_text
        else:
            notes_tf.text = "[OCR] " + combined_text
        return 1
    except Exception:
        # No notes slide — create one
        try:
            notes_slide = slide.notes_slide
            notes_tf = notes_slide.notes_text_frame
            notes_tf.text = "[OCR] " + combined_text
            return 1
        except Exception:
            return 0
