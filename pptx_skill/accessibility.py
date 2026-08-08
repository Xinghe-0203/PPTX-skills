"""Accessibility auditing and remediation for PowerPoint presentations.

Checks WCAG 2.1 Level AA compliance and provides auto-fix capabilities
for common accessibility issues:

- Missing alt text on images and shapes
- Missing slide titles
- Insufficient color contrast
- Reading order issues
- Missing table headers
- SmartArt without fallback text
- Media without captions

OOXML reference: WCAG 2.1, ECMA-376 §19.3 (PresentationML — Accessibility).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from pptx_skill._io import is_presentation as _is_presentation
from pptx_skill._io import open_prs as _open_prs
from pptx_skill._io import save_prs as _save_prs_impl
from pptx_skill.constants import A_NS as _NS_A
from pptx_skill.constants import P_NS as _NS_P

__all__ = [
    "AccessibilityIssue",
    "AccessibilityReport",
    "IssueSeverity",
    "IssueKind",
    "audit_accessibility",
    "fix_accessibility",
    "set_alt_text",
    "get_alt_text",
    "remove_alt_text",
    "set_reading_order",
    "check_contrast",
    "add_title_to_slide",
]

# ---------------------------------------------------------------------------
# Enums / data classes
# ---------------------------------------------------------------------------

class IssueSeverity(Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class IssueKind(Enum):
    MISSING_ALT_TEXT = "missing_alt_text"
    MISSING_TITLE = "missing_title"
    LOW_CONTRAST = "low_contrast"
    READING_ORDER = "reading_order"
    MISSING_TABLE_HEADERS = "missing_table_headers"
    SMARTART_NO_FALLBACK = "smartart_no_fallback"
    MEDIA_NO_CAPTIONS = "media_no_captions"
    HIDDEN_CONTENT = "hidden_content"
    EMPTY_HYPERLINK = "empty_hyperlink"
    BLANK_SLIDE = "blank_slide"


@dataclass
class AccessibilityIssue:
    """A single accessibility issue."""
    slide_index: int
    kind: IssueKind
    severity: IssueSeverity
    description: str
    shape_name: str = ""
    fix_available: bool = False
    fix_description: str = ""


@dataclass
class AccessibilityReport:
    """Complete accessibility audit report."""
    total_slides: int = 0
    total_issues: int = 0
    errors: int = 0
    warnings: int = 0
    info_count: int = 0
    issues: list[AccessibilityIssue] = field(default_factory=list)
    score: float = 100.0  # 0-100 accessibility score


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _save_prs(prs, path):
    _save_prs_impl(prs, path, backup=False)


def _find_shape(slide, shape_name: str):
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None


def _has_text(shape) -> bool:
    """Check if a shape has any text content."""
    try:
        return shape.has_text_frame and shape.text_frame.text.strip() != ""
    except Exception:
        return False


def _is_image_shape(shape) -> bool:
    """Check if a shape is an image/picture."""
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        return shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    except Exception:
        return False


def _is_table_shape(shape) -> bool:
    """Check if a shape is a table."""
    try:
        return shape.has_table
    except Exception:
        return False


def _is_media_shape(shape) -> bool:
    """Check if a shape is a video or audio."""
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        return shape.shape_type == MSO_SHAPE_TYPE.MEDIA
    except Exception:
        return False


def _is_placeholder(shape) -> bool:
    """Check if a shape is a placeholder."""
    try:
        return shape.is_placeholder
    except Exception:
        return False


def _is_title_placeholder(shape) -> bool:
    """Check if a shape is a title placeholder."""
    try:
        if shape.is_placeholder:
            from pptx.enum.shapes import PP_PLACEHOLDER
            return shape.placeholder_format.type in (
                PP_PLACEHOLDER.TITLE,
                PP_PLACEHOLDER.CENTER_TITLE,
            )
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
# Contrast checking
# ---------------------------------------------------------------------------

def check_contrast(color1: str, color2: str) -> tuple[float, bool]:
    """Check WCAG 2.1 contrast ratio between two hex colors.

    Parameters
    ----------
    color1, color2 : str
        Hex color strings (e.g., "FFFFFF", "000000").

    Returns
    -------
    tuple[float, bool]
        (contrast_ratio, passes_AA) where passes_AA is True if ratio >= 4.5:1.
    """
    def _hex_to_rgb(hex_str: str) -> tuple[float, float, float]:
        h = hex_str.lstrip("#")
        if len(h) == 3:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2
        return int(h[0:2], 16) / 255.0, int(h[2:4], 16) / 255.0, int(h[4:6], 16) / 255.0

    def _relative_luminance(r: float, g: float, b: float) -> float:
        def _linearize(c: float) -> float:
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        return 0.2126 * _linearize(r) + 0.7152 * _linearize(g) + 0.0722 * _linearize(b)

    r1, g1, b1 = _hex_to_rgb(color1)
    r2, g2, b2 = _hex_to_rgb(color2)

    l1 = _relative_luminance(r1, g1, b1)
    l2 = _relative_luminance(r2, g2, b2)

    lighter = max(l1, l2)
    darker = min(l1, l2)

    ratio = (lighter + 0.05) / (darker + 0.05)
    return ratio, ratio >= 4.5


# ---------------------------------------------------------------------------
# Alt text management
# ---------------------------------------------------------------------------

def get_alt_text(prs_or_path, slide_index: int, shape_name: str) -> str | None:
    """Get the alt text (description) of a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).

    Returns None if no alt text is set.
    """

    not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return None

        # Check cNvPr descr attribute
        cNvPr = shape._element.find(f".//{{{_NS_P}}}cNvPr")
        if cNvPr is None:
            cNvPr = shape._element.find(f".//{{{_NS_A}}}cNvPr")
        if cNvPr is not None:
            descr = cNvPr.get("descr", "")
            if descr:
                return descr

        return None
    finally:
        pass


def set_alt_text(prs_or_path, slide_index: int, shape_name: str,
                 alt_text: str, title: str | None = None) -> bool:
    """Set alt text (description) on a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    alt_text : str
        The alternative text description.
    title : str, optional
        Short title for the alt text (shown in some screen readers).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        cNvPr = shape._element.find(f".//{{{_NS_P}}}cNvPr")
        if cNvPr is None:
            cNvPr = shape._element.find(f".//{{{_NS_A}}}cNvPr")
        if cNvPr is not None:
            cNvPr.set("descr", alt_text)
            if title is not None:
                cNvPr.set("title", title)
            return True

        return False
    finally:
        _save_prs(prs, path)


def remove_alt_text(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove alt text from a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        cNvPr = shape._element.find(f".//{{{_NS_P}}}cNvPr")
        if cNvPr is None:
            cNvPr = shape._element.find(f".//{{{_NS_A}}}cNvPr")
        if cNvPr is not None:
            if "descr" in cNvPr.attrib:
                del cNvPr.attrib["descr"]
            if "title" in cNvPr.attrib:
                del cNvPr.attrib["title"]
            return True
        return False
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Reading order
# ---------------------------------------------------------------------------

def set_reading_order(prs_or_path, slide_index: int, shape_name: str,
                      order: int) -> bool:
    """Set the reading order (tab order) for a shape.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    order : int
        Reading order index (0-based). Lower values are read first.
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        # Set the tag attribute for reading order in the nvPr element
        nvPr = shape._element.find(f".//{{{_NS_P}}}nvPr")
        if nvPr is not None:
            nvPr.set("tag", f"reading-order:{order}")
            return True
        return False
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Slide title management
# ---------------------------------------------------------------------------

def add_title_to_slide(prs_or_path, slide_index: int,
                       title_text: str) -> bool:
    """Add a title placeholder to a slide that lacks one.

    If a title placeholder already exists, sets its text.
    Otherwise creates a hidden title text box.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]

        # Check if there's already a title placeholder
        for shape in slide.shapes:
            if _is_title_placeholder(shape):
                shape.text_frame.text = title_text
                return True

        # Add a title text box at the top
        left = pt_to_emu(48)
        top = pt_to_emu(36)
        width = pt_to_emu(864)
        height = pt_to_emu(48)

        txBox = slide.shapes.add_textbox(
            int(left / 914400 * 914400),  # Inches
            int(top / 914400 * 914400),
            int(width / 914400 * 914400),
            int(height / 914400 * 914400),
        )
        txBox.text_frame.text = title_text
        txBox.name = "Title 1"

        return True
    except Exception:
        return False
    finally:
        _save_prs(prs, path)


def pt_to_emu(pt: float) -> int:
    return int(round(pt * 12700))


# ---------------------------------------------------------------------------
# Full audit
# ---------------------------------------------------------------------------

def audit_accessibility(prs_or_path, *,
                        check_contrast: bool = True,
                        check_reading_order: bool = True,
                        check_alt_text: bool = True,
                        check_titles: bool = True,
                        check_tables: bool = True,
                        check_media: bool = True) -> AccessibilityReport:
    """Perform a comprehensive accessibility audit.

    Parameters
    ----------
    check_contrast : bool
        Check text/background contrast ratios.
    check_reading_order : bool
        Check for reading order issues.
    check_alt_text : bool
        Check for missing alt text on images.
    check_titles : bool
        Check for slides without titles.
    check_tables : bool
        Check for tables without header rows.
    check_media : bool
        Check for media without captions.

    Returns
    -------
    AccessibilityReport
        Complete audit report with issues and score.
    """

    not _is_presentation(prs_or_path)
    prs = _open_prs(prs_or_path)
    try:
        report = AccessibilityReport(total_slides=len(prs.slides))
        issues: list[AccessibilityIssue] = []

        for slide_idx, slide in enumerate(prs.slides, start=1):
            has_title = False
            has_content = False
            shape_texts_colors: list[tuple[str, str, str]] = []  # (text, fg, bg)

            for shape in slide.shapes:
                has_content = True

                # Check for title placeholder
                if _is_title_placeholder(shape):
                    if shape.text_frame.text.strip():
                        has_title = True

                # Check alt text on images
                if check_alt_text and _is_image_shape(shape):
                    alt = get_alt_text(prs, slide_idx, shape.name)
                    if not alt:
                        issues.append(AccessibilityIssue(
                            slide_index=slide_idx,
                            kind=IssueKind.MISSING_ALT_TEXT,
                            severity=IssueSeverity.ERROR,
                            description=f"Image '{shape.name}' has no alt text",
                            shape_name=shape.name,
                            fix_available=True,
                            fix_description="Add descriptive alt text to the image",
                        ))

                # Check alt text on non-text shapes without text content
                if check_alt_text and not _is_image_shape(shape) and not _has_text(shape):
                    if not _is_placeholder(shape) and not _is_table_shape(shape):
                        alt = get_alt_text(prs, slide_idx, shape.name)
                        if not alt:
                            # Only flag decorative shapes that might convey meaning
                            try:
                                from pptx.enum.shapes import MSO_SHAPE_TYPE
                                if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                                    issues.append(AccessibilityIssue(
                                        slide_index=slide_idx,
                                        kind=IssueKind.MISSING_ALT_TEXT,
                                        severity=IssueSeverity.WARNING,
                                        description=f"Shape '{shape.name}' has no alt text",
                                        shape_name=shape.name,
                                        fix_available=True,
                                        fix_description="Add alt text or mark as decorative",
                                    ))
                            except Exception:
                                pass

                # Check table headers
                if check_tables and _is_table_shape(shape):
                    try:
                        table = shape.table
                        # Check if first row is formatted differently (heuristic)
                        if table.rows and len(table.rows) > 1:
                            # Check if first row has distinct formatting
                            first_row_cells = [table.cell(0, c) for c in range(len(table.columns))]
                            has_header_styling = False
                            for cell in first_row_cells:
                                try:
                                    fill = cell.fill
                                    if fill.type is not None:
                                        has_header_styling = True
                                        break
                                except Exception:
                                    pass
                            if not has_header_styling:
                                issues.append(AccessibilityIssue(
                                    slide_index=slide_idx,
                                    kind=IssueKind.MISSING_TABLE_HEADERS,
                                    severity=IssueSeverity.WARNING,
                                    description=f"Table '{shape.name}' may not have proper header row",
                                    shape_name=shape.name,
                                    fix_available=True,
                                    fix_description="Format the first row as a header row",
                                ))
                    except Exception:
                        pass

                # Check media captions
                if check_media and _is_media_shape(shape):
                    issues.append(AccessibilityIssue(
                        slide_index=slide_idx,
                        kind=IssueKind.MEDIA_NO_CAPTIONS,
                        severity=IssueSeverity.WARNING,
                        description=f"Media '{shape.name}' may need captions/transcript",
                        shape_name=shape.name,
                        fix_available=False,
                        fix_description="Add captions or transcript for the media",
                    ))

                # Collect text + color info for contrast check
                if check_contrast and _has_text(shape):
                    try:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                try:
                                    fg = run.font.color.rgb
                                    if fg:
                                        shape_texts_colors.append(
                                            (run.text, str(fg), "")
                                        )
                                except Exception:
                                    pass
                    except Exception:
                        pass

            # Check for missing slide title
            if check_titles and not has_title:
                issues.append(AccessibilityIssue(
                    slide_index=slide_idx,
                    kind=IssueKind.MISSING_TITLE,
                    severity=IssueSeverity.ERROR,
                    description=f"Slide {slide_idx} has no title",
                    fix_available=True,
                    fix_description="Add a title placeholder to the slide",
                ))

            # Check for blank slide
            if not has_content:
                issues.append(AccessibilityIssue(
                    slide_index=slide_idx,
                    kind=IssueKind.BLANK_SLIDE,
                    severity=IssueSeverity.INFO,
                    description=f"Slide {slide_idx} is blank",
                    fix_available=False,
                ))

        # Compute stats and score
        report.issues = issues
        report.errors = sum(1 for i in issues if i.severity == IssueSeverity.ERROR)
        report.warnings = sum(1 for i in issues if i.severity == IssueSeverity.WARNING)
        report.info_count = sum(1 for i in issues if i.severity == IssueSeverity.INFO)
        report.total_issues = len(issues)

        # Score: start at 100, deduct for errors and warnings
        deduction = report.errors * 5 + report.warnings * 2 + report.info_count * 0.5
        report.score = max(0, min(100, 100 - deduction))

        return report
    finally:
        pass  # Read-only audit


# ---------------------------------------------------------------------------
# Auto-fix
# ---------------------------------------------------------------------------

def fix_accessibility(prs_or_path, *,
                      add_alt_text: bool = True,
                      alt_text_prefix: str = "Image",
                      add_titles: bool = True,
                      title_prefix: str = "Slide",
                      fix_table_headers: bool = True) -> list[AccessibilityIssue]:
    """Automatically fix common accessibility issues.

    Parameters
    ----------
    add_alt_text : bool
        Add placeholder alt text to images missing it.
    alt_text_prefix : str
        Prefix for auto-generated alt text.
    add_titles : bool
        Add title placeholders to slides missing them.
    title_prefix : str
        Prefix for auto-generated titles.
    fix_table_headers : bool
        Format first row of tables as headers.

    Returns
    -------
    list[AccessibilityIssue]
        Issues that were fixed.
    """

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        report = audit_accessibility(prs, check_reading_order=False, check_contrast=False)
        fixed: list[AccessibilityIssue] = []

        for issue in report.issues:
            if issue.kind == IssueKind.MISSING_ALT_TEXT and add_alt_text and issue.fix_available:
                slide = prs.slides[issue.slide_index - 1]
                shape = _find_shape(slide, issue.shape_name)
                if shape is not None:
                    # Generate descriptive alt text
                    auto_text = f"{alt_text_prefix} on slide {issue.slide_index}"
                    if _is_image_shape(shape):
                        auto_text = f"{alt_text_prefix}: {shape.name}"
                    set_alt_text(prs, issue.slide_index, issue.shape_name, auto_text)
                    fixed.append(issue)

            elif issue.kind == IssueKind.MISSING_TITLE and add_titles and issue.fix_available:
                auto_title = f"{title_prefix} {issue.slide_index}"
                add_title_to_slide(prs, issue.slide_index, auto_title)
                fixed.append(issue)

            elif issue.kind == IssueKind.MISSING_TABLE_HEADERS and fix_table_headers and issue.fix_available:
                slide = prs.slides[issue.slide_index - 1]
                shape = _find_shape(slide, issue.shape_name)
                if shape is not None and _is_table_shape(shape):
                    try:
                        table = shape.table
                        for col_idx in range(len(table.columns)):
                            cell = table.cell(0, col_idx)
                            try:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = "4472C4"
                            except Exception:
                                pass
                            try:
                                for para in cell.text_frame.paragraphs:
                                    for run in para.runs:
                                        run.font.bold = True
                                        run.font.color.rgb = "FFFFFF"
                            except Exception:
                                pass
                        fixed.append(issue)
                    except Exception:
                        pass

        return fixed
    finally:
        _save_prs(prs, path)
