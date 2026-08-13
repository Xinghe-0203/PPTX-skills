"""Markdown export — convert PPTX presentations to Markdown.

Exports slide content (text, tables, speaker notes, chart data) as structured
Markdown.  Useful for documentation, wikis, AI/LLM ingestion, and version
control of presentation content.

Quick start
-----------
>>> from pptx_skill.markdown_export import export_to_markdown
>>> md = export_to_markdown("deck.pptx", "deck.md")
"""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass

from pptx_skill._io import open_prs as _open_prs

__all__ = [
    "MarkdownOptions",
    "export_to_markdown",
    "slide_to_markdown",
    "extract_markdown_content",
]

log = logging.getLogger(__name__)


@dataclass
class MarkdownOptions:
    """Options controlling Markdown export."""
    include_speaker_notes: bool = True
    include_slide_numbers: bool = True
    include_tables: bool = True
    include_images_as_links: bool = True
    image_output_dir: str | None = None
    heading_level: int = 2  # Slide titles start at this level (## = 2)
    bullets_style: str = "dash"  # "dash" (-), "asterisk" (*), "plus" (+)
    table_alignment: str = "left"  # "left", "center", "right"
    include_metadata: bool = True
    page_break_between_slides: bool = True  # Insert --- between slides


def _bullet_char(style: str) -> str:
    return {"dash": "-", "asterisk": "*", "plus": "+"}.get(style, "-")


def _table_alignment_colons(align: str, col_count: int) -> str:
    """Generate the alignment row for a Markdown table."""
    marker = {"left": ":---", "center": ":---:", "right": "---:"}.get(align, ":---")
    return "| " + " | ".join([marker] * col_count) + " |"


def _extract_shape_text(shape) -> str:
    """Extract text from a shape as Markdown."""
    if not shape.has_text_frame:
        return ""

    lines = []
    for para in shape.text_frame.paragraphs:
        if not para.runs:
            if para.text:
                lines.append(para.text)
            continue

        # Determine if it's a bullet
        is_bullet = para.level > 0

        # Collect run text
        run_texts = []
        for run in para.runs:
            text = run.text
            if not text:
                continue
            # Apply formatting
            if run.font.bold and run.font.italic:
                text = f"***{text}***"
            elif run.font.bold:
                text = f"**{text}**"
            elif run.font.italic:
                text = f"*{text}*"
            # Check for hyperlink
            try:
                if run.hyperlink and run.hyperlink.address:
                    text = f"[{text}]({run.hyperlink.address})"
            except Exception:
                pass
            run_texts.append(text)

        line = "".join(run_texts).rstrip()
        if not line:
            continue

        if is_bullet:
            indent = "  " * para.level
            lines.append(f"{indent}- {line}")
        else:
            lines.append(line)

    return "\n".join(lines)


def _extract_table_as_markdown(shape, align: str = "left") -> str:
    """Extract a table shape as a Markdown table."""
    if not shape.has_table:
        return ""

    table = shape.table
    rows = list(table.rows)
    if not rows:
        return ""

    col_count = len(rows[0].cells)
    md_lines = []

    # Header row
    header_cells = []
    for cell in rows[0].cells:
        cell_text = cell.text_frame.text.replace("|", "\\|").replace("\n", " ").strip()
        header_cells.append(cell_text)
    md_lines.append("| " + " | ".join(header_cells) + " |")
    md_lines.append(_table_alignment_colons(align, col_count))

    # Data rows
    for row in rows[1:]:
        row_cells = []
        for cell in row.cells:
            cell_text = cell.text_frame.text.replace("|", "\\|").replace("\n", " ").strip()
            row_cells.append(cell_text)
        md_lines.append("| " + " | ".join(row_cells) + " |")

    return "\n".join(md_lines)


def _extract_image_as_markdown(shape, slide_idx: int, shape_idx: int,
                                image_output_dir: str | None) -> str:
    """Extract an image shape as a Markdown image reference."""
    try:
        image = shape.image
        ext = image.content_type.split("/")[-1]
        if ext == "jpeg":
            ext = "jpg"

        if image_output_dir:
            os.makedirs(image_output_dir, exist_ok=True)
            filename = f"slide_{slide_idx + 1}_img_{shape_idx + 1}.{ext}"
            filepath = os.path.join(image_output_dir, filename)
            with open(filepath, "wb") as f:
                f.write(image.blob)
            return f"![image](./{filename})"
        else:
            # Just note the image exists
            return f"*![image: {ext}]*"
    except Exception:
        return ""


def _extract_chart_as_markdown(shape) -> str:
    """Extract chart data as a Markdown table."""
    try:
        if not shape.has_chart:
            return ""
        chart = shape.chart

        lines = [f"*Chart: {chart.chart_type}*", ""]

        # Try to get categories and series
        try:
            plot = chart.plots[0]
            categories = list(plot.categories)
        except Exception:
            categories = []

        series_data = []
        for series in chart.series:
            try:
                name = series.format_code if hasattr(series, "format_code") else "Series"
                values = []
                for v in series.values:
                    try:
                        values.append(float(v))
                    except (TypeError, ValueError):
                        values.append(0.0)
                series_data.append((str(name), values))
            except Exception:
                continue

        if not series_data:
            return ""

        # Build table
        header = ["Category"] + [name for name, _ in series_data]
        lines.append("| " + " | ".join(header) + " |")
        lines.append("| " + " | ".join([":---"] * len(header)) + " |")

        for i, cat in enumerate(categories):
            row = [str(cat)]
            for _, values in series_data:
                val = values[i] if i < len(values) else ""
                row.append(str(val))
            lines.append("| " + " | ".join(row) + " |")

        return "\n".join(lines)
    except Exception:
        return ""


def _extract_speaker_notes(slide) -> str:
    """Extract speaker notes as Markdown blockquote."""
    try:
        notes_slide = slide.notes_slide
        notes_text = notes_slide.notes_text_frame.text.strip()
        if not notes_text:
            return ""
        # Format as blockquote
        lines = notes_text.split("\n")
        quoted = "\n".join(f"> {line}" for line in lines)
        return f"**Speaker Notes:**\n\n{quoted}"
    except Exception:
        return ""


def _guess_title(shape_text: str) -> str:
    """Extract a title from shape text (first non-empty line)."""
    for line in shape_text.split("\n"):
        stripped = line.strip()
        if stripped:
            clean = stripped.lstrip("#").strip()
            for marker in ("***", "**", "*"):
                if clean.startswith(marker) and clean.endswith(marker) and len(clean) > len(marker) * 2:
                    clean = clean[len(marker):-len(marker)].strip()
                    break
            return clean
    return ""


def slide_to_markdown(slide, slide_index: int, *, options: MarkdownOptions | None = None,
                       prs=None) -> str:
    """Convert a single slide to Markdown.

    Parameters
    ----------
    slide : pptx.slide.Slide
        The slide object.
    slide_index : int
        0-based slide index (for numbering).
    options : MarkdownOptions, optional
        Export options.
    prs : Presentation, optional
        Parent presentation (for slide dimensions if needed).

    Returns
    -------
    str
        Markdown representation of the slide.
    """
    opts = options or MarkdownOptions()
    sections = []

    # Slide header
    title = ""
    body_parts = []
    image_parts = []
    table_parts = []
    chart_parts = []

    shape_idx = 0
    for shape in slide.shapes:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if opts.include_images_as_links:
                    img_md = _extract_image_as_markdown(
                        shape, slide_index, shape_idx, opts.image_output_dir
                    )
                    if img_md:
                        image_parts.append(img_md)
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                if opts.include_tables:
                    tbl_md = _extract_table_as_markdown(shape, opts.table_alignment)
                    if tbl_md:
                        table_parts.append(tbl_md)
            elif shape.has_chart:
                chart_md = _extract_chart_as_markdown(shape)
                if chart_md:
                    chart_parts.append(chart_md)
            elif shape.has_text_frame:
                text = _extract_shape_text(shape)
                if text:
                    if not title:
                        candidate = _guess_title(text)
                        if candidate:
                            title = candidate
                            # Add the rest as body if there's more
                            rest = text.split("\n", 1)
                            if len(rest) > 1 and rest[1].strip():
                                body_parts.append(rest[1])
                    else:
                        body_parts.append(text)
        except Exception as exc:
            log.debug("Error extracting shape %d: %s", shape_idx, exc)
        shape_idx += 1

    # Build slide markdown
    heading_prefix = "#" * opts.heading_level

    # Title
    if not title:
        title = f"Slide {slide_index + 1}"
    header_line = f"{heading_prefix} {title}"
    if opts.include_slide_numbers:
        header_line += f" *(Slide {slide_index + 1})*"
    sections.append(header_line)
    sections.append("")

    # Body text
    if body_parts:
        sections.append("\n\n".join(body_parts))
        sections.append("")

    # Tables
    for tbl in table_parts:
        sections.append(tbl)
        sections.append("")

    # Charts
    for chart in chart_parts:
        sections.append(chart)
        sections.append("")

    # Images
    for img in image_parts:
        sections.append(img)
        sections.append("")

    # Speaker notes
    if opts.include_speaker_notes:
        notes = _extract_speaker_notes(slide)
        if notes:
            sections.append(notes)
            sections.append("")

    # Page break
    if opts.page_break_between_slides:
        sections.append("---")

    return "\n".join(sections)


def extract_markdown_content(prs_or_path, *, options: MarkdownOptions | None = None) -> str:
    """Extract all content from a presentation as Markdown.

    Parameters
    ----------
    options : MarkdownOptions, optional
        Export options.

    Returns
    -------
    str
        Full Markdown string.
    """
    opts = options or MarkdownOptions()
    prs = _open_prs(prs_or_path)
    try:
        parts = []

        # Metadata header
        if opts.include_metadata:
            try:
                cp = prs.core_properties
                parts.append(f"# {cp.title or 'Presentation'}")
                parts.append("")
                if cp.author:
                    parts.append(f"**Author:** {cp.author}  ")
                if cp.subject:
                    parts.append(f"**Subject:** {cp.subject}  ")
                if cp.created:
                    parts.append(f"**Created:** {cp.created.strftime('%Y-%m-%d')}  ")
                if cp.modified:
                    parts.append(f"**Modified:** {cp.modified.strftime('%Y-%m-%d')}  ")
                parts.append(f"**Slides:** {len(prs.slides)}  ")
                parts.append("")
                parts.append("---")
                parts.append("")
            except Exception:
                parts.append("# Presentation")
                parts.append("")

        # Slides
        for idx, slide in enumerate(prs.slides):
            slide_md = slide_to_markdown(slide, idx, options=opts, prs=prs)
            parts.append(slide_md)
            parts.append("")

        return "\n".join(parts)
    finally:
        pass


def export_to_markdown(
    prs_or_path,
    output_path: str | None = None,
    *,
    options: MarkdownOptions | None = None,
) -> str:
    """Export a presentation to Markdown.

    Parameters
    ----------
    output_path : str, optional
        Path to write the Markdown file. If None, returns the string only.
    options : MarkdownOptions, optional
        Export options.

    Returns
    -------
    str
        The Markdown content.
    """
    content = extract_markdown_content(prs_or_path, options=options)

    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(content)

    return content
