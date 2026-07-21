"""Batch / parallel processing of PPTX files.

Provides high-level functions that operate on entire directories of PPTX files
in parallel, with progress tracking, error handling, and aggregate statistics.
CPU-bound work (rendering, image conversion) uses
:class:`~concurrent.futures.ProcessPoolExecutor`; I/O-bound work (file copy,
watermark application) uses :class:`~concurrent.futures.ThreadPoolExecutor`.

Public API
----------
- :func:`batch_convert`        -- Convert all PPTX files to PDF / PNG / etc.
- :func:`batch_apply`          -- Apply a custom operation to each PPTX file.
- :func:`batch_merge`          -- Merge multiple PPTX files into one.
- :func:`batch_watermark`      -- Apply watermarks to all PPTX files.
- :func:`batch_recolor`        -- Recolor all PPTX files.
- :func:`batch_export_images`  -- Export all slides as images.
- :func:`batch_inspect`        -- Inspect all PPTX files and return info dicts.
- :func:`batch_stats`          -- Aggregate statistics across all PPTX files.
"""
from __future__ import annotations

import logging
import os
import shutil
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

__all__ = [
    "BatchResult",
    "BatchStats",
    "batch_apply",
    "batch_convert",
    "batch_export_images",
    "batch_inspect",
    "batch_merge",
    "batch_recolor",
    "batch_stats",
    "batch_watermark",
]

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class BatchResult:
    """Summary of a batch operation.

    Attributes
    ----------
    total : int
        Total number of files that were attempted.
    succeeded : int
        Number of files processed successfully.
    failed : int
        Number of files that failed.
    errors : list[dict]
        Per-file error details.  Each dict has keys ``"file"`` and ``"error"``.
    output_paths : list[str]
        Absolute paths of successfully written output files.
    duration_seconds : float
        Wall-clock duration of the batch operation.
    """

    total: int
    succeeded: int
    failed: int
    errors: list[dict] = field(default_factory=list)
    output_paths: list[str] = field(default_factory=list)
    duration_seconds: float = 0.0

    def __repr__(self) -> str:
        return (
            f"BatchResult(total={self.total}, succeeded={self.succeeded}, "
            f"failed={self.failed}, duration={self.duration_seconds:.1f}s)"
        )


@dataclass
class BatchStats:
    """Aggregate statistics across a collection of PPTX files.

    Attributes
    ----------
    total_files : int
        Number of PPTX files scanned.
    total_slides : int
        Sum of slide counts across all files.
    total_shapes : int
        Sum of shape counts across all files.
    total_images : int
        Sum of image counts across all files.
    total_media : int
        Sum of media (audio/video) counts across all files.
    avg_slide_count : float
        Average number of slides per file.
    themes_used : dict
        Mapping of theme name to occurrence count.
    layouts_used : dict
        Mapping of layout name to occurrence count.
    """

    total_files: int = 0
    total_slides: int = 0
    total_shapes: int = 0
    total_images: int = 0
    total_media: int = 0
    avg_slide_count: float = 0.0
    themes_used: dict = field(default_factory=dict)
    layouts_used: dict = field(default_factory=dict)

    def __repr__(self) -> str:
        return (
            f"BatchStats(files={self.total_files}, slides={self.total_slides}, "
            f"shapes={self.total_shapes}, avg_slides={self.avg_slide_count:.1f})"
        )


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------

def _find_pptx_files(
    input_dir: str | os.PathLike[str],
    pattern: str = "*.pptx",
    recursive: bool = False,
) -> list[Path]:
    """Return a sorted list of PPTX file paths matching *pattern*.

    Parameters
    ----------
    input_dir : str or PathLike
        Directory to search.
    pattern : str
        Glob pattern (default ``"*.pptx"``).
    recursive : bool
        If ``True``, search subdirectories recursively.

    Returns
    -------
    list[Path]
        Sorted list of absolute file paths.
    """
    input_dir = Path(input_dir).resolve()
    if not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    if recursive:
        files = sorted(input_dir.rglob(pattern))
    else:
        files = sorted(input_dir.glob(pattern))

    # Filter out temporary files (starting with ~$)
    return [f for f in files if not f.name.startswith("~$")]


def _print_progress(current: int, total: int, filename: str) -> None:
    """Print a simple progress indicator to stdout.

    Parameters
    ----------
    current : int
        1-based index of the current file being processed.
    total : int
        Total number of files.
    filename : str
        Name of the current file (without directory).
    """
    pct = current / total * 100 if total > 0 else 0
    bar_width = 30
    filled = int(bar_width * current / total) if total > 0 else 0
    bar = "=" * filled + "-" * (bar_width - filled)
    sys.stdout.write(f"\r[{bar}] {pct:5.1f}% ({current}/{total}) {filename}")
    sys.stdout.flush()
    if current == total:
        sys.stdout.write("\n")
        sys.stdout.flush()


def _process_single(
    input_path: Path,
    output_dir: Path,
    operation: Callable[..., bool],
    on_error: str,
    **kwargs: Any,
) -> tuple[bool, str | None, dict | None]:
    """Process a single PPTX file with *operation*.

    Parameters
    ----------
    input_path : Path
        Absolute path to the input PPTX file.
    output_dir : Path
        Absolute path to the output directory.
    operation : Callable
        Function with signature ``(prs_or_path, output_path, **kwargs) -> bool``.
    on_error : str
        ``"skip"``, ``"raise"``, or ``"collect"``.
    **kwargs
        Extra keyword arguments forwarded to *operation*.

    Returns
    -------
    tuple[bool, str | None, dict | None]
        ``(success, output_path_or_None, error_dict_or_None)``
    """
    output_path = output_dir / input_path.name
    try:
        result = operation(str(input_path), str(output_path), **kwargs)
        # If the operation returns False, treat as failure
        if result is False:
            return False, None, {"file": str(input_path), "error": "Operation returned False"}
        # Discover the actual output path.  The operation may have written to
        # a different extension (e.g. .pdf instead of .pptx) or created a
        # subdirectory.  Try the exact path first, then look for alternatives.
        actual = _discover_output(output_path, output_dir, input_path.stem)
        return True, actual, None
    except Exception as exc:
        if on_error == "raise":
            raise
        return False, None, {"file": str(input_path), "error": str(exc)}


def _discover_output(
    nominal_path: Path,
    output_dir: Path,
    stem: str,
) -> str | None:
    """Find the actual output file after an operation.

    Checks the *nominal_path* first, then looks for files in *output_dir*
    whose stem matches *stem* but with a different extension (e.g. ``.pdf``,
    ``.html``, ``.txt``), or a subdirectory named *stem* (for image exports).

    Returns the absolute path string of the found output, or ``None``.
    """
    # 1. Exact match
    if nominal_path.exists() and nominal_path.is_file():
        return str(nominal_path)

    # 2. Same stem, different extension
    for ext in (".pdf", ".html", ".txt", ".jpg", ".png"):
        candidate = nominal_path.with_suffix(ext)
        if candidate.exists() and candidate.is_file():
            return str(candidate)

    # 3. Subdirectory named after the stem (image export)
    sub = output_dir / stem
    if sub.is_dir():
        return str(sub)

    # 4. Any file in output_dir that starts with stem
    for f in sorted(output_dir.iterdir()):
        if f.stem == stem and f.is_file():
            return str(f)

    return None


def _default_max_workers() -> int:
    """Return a sensible default worker count.

    Uses ``os.cpu_count()`` (minimum 1, maximum 8) to avoid overwhelming
    the system.
    """
    cpu = os.cpu_count() or 1
    return min(max(cpu, 1), 8)


def _ensure_dir(path: str | os.PathLike[str]) -> str:
    """Create directory if needed and return the absolute path."""
    abs_path = os.path.abspath(path)
    os.makedirs(abs_path, exist_ok=True)
    return abs_path


# ---------------------------------------------------------------------------
# Internal operation wrappers
# ---------------------------------------------------------------------------

def _convert_operation(
    prs_or_path: str,
    output_path: str,
    *,
    format: str = "pdf",
) -> bool:
    """Single-file conversion operation for :func:`batch_convert`.

    *output_path* arrives with a ``.pptx`` extension (because it mirrors the
    input filename).  This function replaces the extension with the appropriate
    one for the target *format* before calling the export function.
    """
    from pptx_skill.export import (
        export_to_html,
        export_to_images,
        export_to_pdf,
        export_to_text,
    )

    fmt = format.lower()
    p = Path(output_path)

    if fmt == "pdf":
        actual_output = str(p.with_suffix(".pdf"))
        export_to_pdf(prs_or_path, actual_output)
    elif fmt in ("png", "jpg", "jpeg", "bmp"):
        img_fmt = "PNG" if fmt == "png" else ("JPEG" if fmt in ("jpg", "jpeg") else "BMP")
        # For image export, create a subdirectory named after the file
        out_dir = p.with_suffix("")
        out_dir.mkdir(parents=True, exist_ok=True)
        export_to_images(prs_or_path, str(out_dir), format=img_fmt)
    elif fmt == "html":
        actual_output = str(p.with_suffix(".html"))
        export_to_html(prs_or_path, actual_output)
    elif fmt == "txt":
        actual_output = str(p.with_suffix(".txt"))
        export_to_text(prs_or_path, actual_output)
    else:
        raise ValueError(f"Unsupported export format: {format!r}")
    return True


def _watermark_operation(
    prs_or_path: str,
    output_path: str,
    *,
    text: str | None = None,
    image_path: str | None = None,
    opacity: float = 0.3,
) -> bool:
    """Single-file watermark operation for :func:`batch_watermark`."""
    from pptx_skill.watermark import add_image_watermark, add_text_watermark

    # Copy input to output first so we don't modify the original
    shutil.copy2(prs_or_path, output_path)

    if text is not None:
        add_text_watermark(output_path, text, opacity=opacity)
    elif image_path is not None:
        add_image_watermark(output_path, image_path, opacity=opacity)
    else:
        raise ValueError("Either text or image_path must be provided for watermark")
    return True


def _recolor_operation(
    prs_or_path: str,
    output_path: str,
    *,
    old_hex: str = "",
    new_hex: str = "",
) -> bool:
    """Single-file recolor operation for :func:`batch_recolor`."""
    from pptx import Presentation

    # Copy input to output first
    shutil.copy2(prs_or_path, output_path)

    # Validate hex colors
    for value in (old_hex, new_hex):
        if not value.startswith("#") or len(value) != 7:
            raise ValueError(f"Expected #RRGGBB color: {value}")

    prs = Presentation(str(output_path))
    replaced = _recolor_prs(prs, {old_hex: new_hex})
    if replaced:
        prs.save(str(output_path))
    return True


def _recolor_prs(prs: Any, color_map: dict[str, str]) -> int:
    """Replace colors throughout *prs* according to *color_map*.

    Returns the number of color replacements made.
    """
    replaced = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            # Shape fill (may raise if fill type is NoneFill)
            try:
                fill = shape.fill
                if fill is not None:
                    replaced += _replace_color_format(fill.fore_color, color_map)
            except Exception:
                pass
            # Shape line
            try:
                line = shape.line
                if line is not None:
                    replaced += _replace_color_format(line.color, color_map)
            except Exception:
                pass
            # Text runs
            if getattr(shape, "has_text_frame", False):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        try:
                            replaced += _replace_color_format(run.font.color, color_map)
                        except Exception:
                            pass
    return replaced


def _replace_color_format(color_format: Any, color_map: dict[str, str]) -> int:
    """Attempt to replace the color in *color_format* using *color_map*.

    Returns 1 if a replacement was made, 0 otherwise.
    """
    try:
        rgb = color_format.rgb
        if rgb is None:
            return 0
        hex_str = f"#{rgb}".upper()
        if hex_str in color_map:
            new_hex = color_map[hex_str].lstrip("#")
            from pptx.dml.color import RGBColor

            color_format.rgb = RGBColor.from_string(new_hex)
            return 1
    except Exception:
        pass
    return 0


def _export_images_operation(
    prs_or_path: str,
    output_path: str,
    *,
    dpi: int = 150,
    format: str = "png",
) -> bool:
    """Single-file image export operation for :func:`batch_export_images`."""
    from pptx_skill.export import export_to_images

    # output_path is used as a directory base
    out_dir = Path(output_path).with_suffix("")
    out_dir.mkdir(parents=True, exist_ok=True)

    img_fmt = format.upper()
    if img_fmt == "JPG":
        img_fmt = "JPEG"

    export_to_images(prs_or_path, str(out_dir), dpi=dpi, format=img_fmt)
    return True


def _inspect_single(input_path: Path) -> dict:
    """Inspect a single PPTX file and return structured info."""
    from pptx import Presentation

    prs = Presentation(str(input_path))
    slide_count = len(prs.slides)
    shape_count = 0
    image_count = 0
    media_count = 0
    layouts: dict[str, int] = {}

    for slide in prs.slides:
        for shape in slide.shapes:
            shape_count += 1
            # Picture type = 13
            if getattr(shape, "shape_type", None) == 13:
                image_count += 1
            # Media types (video=16, audio=19) -- approximate
            shape_type_val = getattr(shape, "shape_type", None)
            if shape_type_val is not None:
                try:
                    st = int(shape_type_val)
                    if st in (16, 19):
                        media_count += 1
                except (TypeError, ValueError):
                    pass
        # Layout name
        try:
            layout_name = slide.slide_layout.name
            layouts[layout_name] = layouts.get(layout_name, 0) + 1
        except Exception:
            pass

    # Theme name (best-effort from slide master)
    theme_name = "unknown"
    try:
        for master in prs.slide_masters:
            if master.slide_layouts:
                theme_name = master.slide_layouts[0].name.split(" ")[0]
                break
    except Exception:
        pass

    return {
        "file": str(input_path),
        "filename": input_path.name,
        "slide_count": slide_count,
        "shape_count": shape_count,
        "image_count": image_count,
        "media_count": media_count,
        "theme": theme_name,
        "layouts": layouts,
    }


# ---------------------------------------------------------------------------
# Core batch execution engine
# ---------------------------------------------------------------------------

def _run_batch(
    files: list[Path],
    output_dir: Path,
    operation: Callable[..., bool],
    *,
    on_error: str = "skip",
    max_workers: int | None = None,
    progress: bool = True,
    cpu_bound: bool = False,
    **kwargs: Any,
) -> BatchResult:
    """Execute *operation* on each file in *files*, collecting results.

    Parameters
    ----------
    files : list[Path]
        Input PPTX file paths.
    output_dir : Path
        Output directory (created if needed).
    operation : Callable
        Function with signature ``(input_path, output_path, **kwargs) -> bool``.
    on_error : str
        ``"skip"`` (log and continue), ``"raise"`` (stop on first error),
        ``"collect"`` (collect all errors, raise at end).
    max_workers : int or None
        Number of parallel workers.  ``None`` uses :func:`_default_max_workers`.
    progress : bool
        Print progress to stdout.
    cpu_bound : bool
        If ``True``, use :class:`ProcessPoolExecutor`; otherwise
        :class:`ThreadPoolExecutor`.
    **kwargs
        Extra keyword arguments forwarded to *operation*.

    Returns
    -------
    BatchResult
    """
    from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

    if on_error not in ("skip", "raise", "collect"):
        raise ValueError(
            f"Invalid on_error value: {on_error!r} "
            "(expected 'skip', 'raise', or 'collect')"
        )

    start_time = time.monotonic()
    total = len(files)
    if total == 0:
        return BatchResult(
            total=0, succeeded=0, failed=0,
            duration_seconds=time.monotonic() - start_time,
        )

    _ensure_dir(str(output_dir))
    workers = max_workers or _default_max_workers()
    workers = min(workers, total)  # No point having more workers than files

    succeeded = 0
    failed = 0
    errors: list[dict] = []
    output_paths: list[str] = []

    Executor = ProcessPoolExecutor if cpu_bound else ThreadPoolExecutor

    with Executor(max_workers=workers) as executor:
        # Submit all tasks
        future_to_file: dict[Any, Path] = {}
        for file_path in files:
            future = executor.submit(
                _process_single,
                file_path,
                output_dir,
                operation,
                on_error,
                **kwargs,
            )
            future_to_file[future] = file_path

        # Collect results as they complete
        completed = 0
        first_error: Exception | None = None

        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            completed += 1

            try:
                success, out_path, error_dict = future.result()
            except Exception as exc:
                # Unhandled exception from the worker
                success = False
                out_path = None
                error_dict = {"file": str(file_path), "error": str(exc)}
                if on_error == "raise" and first_error is None:
                    first_error = exc

            if success and out_path is not None:
                succeeded += 1
                output_paths.append(out_path)
            else:
                failed += 1
                if error_dict is not None:
                    errors.append(error_dict)
                    log.warning(
                        "Failed to process %s: %s",
                        error_dict.get("file", file_path),
                        error_dict.get("error", "unknown"),
                    )

            if progress:
                _print_progress(completed, total, file_path.name)

            # Stop early on "raise" mode
            if on_error == "raise" and first_error is not None:
                # Cancel remaining futures
                for f in future_to_file:
                    f.cancel()
                raise first_error

    duration = time.monotonic() - start_time

    # In "collect" mode, raise a summary exception if there were failures
    if on_error == "collect" and failed > 0:
        error_summary = "; ".join(
            f"{e['file']}: {e['error']}" for e in errors
        )
        raise RuntimeError(
            f"Batch completed with {failed}/{total} failures: {error_summary}"
        )

    return BatchResult(
        total=total,
        succeeded=succeeded,
        failed=failed,
        errors=errors,
        output_paths=output_paths,
        duration_seconds=duration,
    )


# ---------------------------------------------------------------------------
# Public API — batch_convert
# ---------------------------------------------------------------------------

def batch_convert(
    input_dir: str,
    output_dir: str,
    *,
    format: str = "pdf",
    pattern: str = "*.pptx",
    recursive: bool = False,
    max_workers: int | None = None,
    on_error: str = "skip",
    progress: bool = True,
) -> BatchResult:
    """Convert all PPTX files in *input_dir* to the specified *format*.

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    output_dir : str
        Directory for converted output files.  Created if it does not exist.
    format : str
        Output format: ``"pdf"``, ``"png"``, ``"jpg"``, ``"html"``, or ``"txt"``.
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    recursive : bool
        Search subdirectories recursively.
    max_workers : int or None
        Number of parallel workers.  ``None`` uses CPU count (capped at 8).
    on_error : str
        ``"skip"`` (log and continue), ``"raise"`` (stop on first error),
        ``"collect"`` (collect all errors, raise at end).
    progress : bool
        Print progress to stdout.

    Returns
    -------
    BatchResult
        Summary of the batch operation.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    ValueError
        If *format* is unsupported or *on_error* is invalid.
    """
    files = _find_pptx_files(input_dir, pattern, recursive)
    out_dir = Path(output_dir).resolve()

    # Conversion (especially PDF/image) is CPU-bound
    return _run_batch(
        files,
        out_dir,
        _convert_operation,
        on_error=on_error,
        max_workers=max_workers,
        progress=progress,
        cpu_bound=True,
        format=format,
    )


# ---------------------------------------------------------------------------
# Public API — batch_apply
# ---------------------------------------------------------------------------

def batch_apply(
    input_dir: str,
    output_dir: str,
    *,
    operation: Callable[..., bool],
    pattern: str = "*.pptx",
    recursive: bool = False,
    max_workers: int | None = None,
    on_error: str = "skip",
    progress: bool = True,
    **kwargs: Any,
) -> BatchResult:
    """Apply a custom *operation* function to each PPTX file.

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    output_dir : str
        Directory for output files.  Created if it does not exist.
    operation : Callable
        Function with signature ``(prs_or_path, output_path, **kwargs) -> bool``.
        Should return ``True`` on success, ``False`` on failure.
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    recursive : bool
        Search subdirectories recursively.
    max_workers : int or None
        Number of parallel workers.  ``None`` uses CPU count (capped at 8).
    on_error : str
        ``"skip"`` (log and continue), ``"raise"`` (stop on first error),
        ``"collect"`` (collect all errors, raise at end).
    progress : bool
        Print progress to stdout.
    **kwargs
        Extra keyword arguments forwarded to *operation*.

    Returns
    -------
    BatchResult
        Summary of the batch operation.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    ValueError
        If *on_error* is invalid.
    """
    files = _find_pptx_files(input_dir, pattern, recursive)
    out_dir = Path(output_dir).resolve()

    # Custom operations are assumed I/O-bound (file manipulation)
    return _run_batch(
        files,
        out_dir,
        operation,
        on_error=on_error,
        max_workers=max_workers,
        progress=progress,
        cpu_bound=False,
        **kwargs,
    )


# ---------------------------------------------------------------------------
# Public API — batch_merge
# ---------------------------------------------------------------------------

def batch_merge(
    input_paths: list[str],
    output_path: str,
    *,
    dedup_layouts: bool = True,
    dedup_media: bool = True,
) -> str:
    """Merge multiple PPTX files into one.

    Uses :func:`pptx_skill.merge.merge_presentations` internally.  Merging is
    inherently sequential, so no parallelism is applied.

    Parameters
    ----------
    input_paths : list[str]
        Paths to the PPTX files to merge.  Must contain at least one path.
    output_path : str
        Destination path for the merged PPTX file.
    dedup_layouts : bool
        If ``True`` (default), skip duplicate layout names during merge.
    dedup_media : bool
        If ``True`` (default), avoid copying duplicate media blobs.

    Returns
    -------
    str
        Absolute path of the merged PPTX file.

    Raises
    ------
    FileNotFoundError
        If any input file does not exist.
    ValueError
        If *input_paths* is empty.
    """
    from pptx_skill.merge import merge_presentations

    if not input_paths:
        raise ValueError("input_paths must contain at least one PPTX file path")

    # Validate all paths exist
    for p in input_paths:
        if not Path(p).exists():
            raise FileNotFoundError(f"Input file not found: {p}")

    on_conflict = "skip" if dedup_layouts else "rename"

    result = merge_presentations(
        sources=input_paths,
        output_path=output_path,
        on_conflict=on_conflict,
    )

    return result.output_path


# ---------------------------------------------------------------------------
# Public API — batch_watermark
# ---------------------------------------------------------------------------

def batch_watermark(
    input_dir: str,
    output_dir: str,
    *,
    text: str | None = None,
    image_path: str | None = None,
    opacity: float = 0.3,
    pattern: str = "*.pptx",
    max_workers: int | None = None,
) -> BatchResult:
    """Apply a watermark to all PPTX files in *input_dir*.

    Either *text* or *image_path* must be provided (but not both).

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    output_dir : str
        Directory for watermarked output files.  Created if it does not exist.
    text : str or None
        Watermark text.  Mutually exclusive with *image_path*.
    image_path : str or None
        Path to watermark image.  Mutually exclusive with *text*.
    opacity : float
        Watermark opacity from 0.0 (invisible) to 1.0 (opaque).
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    max_workers : int or None
        Number of parallel workers.

    Returns
    -------
    BatchResult
        Summary of the batch operation.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    ValueError
        If neither *text* nor *image_path* is provided.
    """
    if text is None and image_path is None:
        raise ValueError("Either text or image_path must be provided for watermark")

    files = _find_pptx_files(input_dir, pattern, recursive=False)
    out_dir = Path(output_dir).resolve()

    return _run_batch(
        files,
        out_dir,
        _watermark_operation,
        on_error="skip",
        max_workers=max_workers,
        progress=True,
        cpu_bound=False,
        text=text,
        image_path=image_path,
        opacity=opacity,
    )


# ---------------------------------------------------------------------------
# Public API — batch_recolor
# ---------------------------------------------------------------------------

def batch_recolor(
    input_dir: str,
    output_dir: str,
    *,
    old_hex: str,
    new_hex: str,
    pattern: str = "*.pptx",
    max_workers: int | None = None,
) -> BatchResult:
    """Recolor all PPTX files in *input_dir*.

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    output_dir : str
        Directory for recolored output files.  Created if it does not exist.
    old_hex : str
        Original color in ``#RRGGBB`` format.
    new_hex : str
        Replacement color in ``#RRGGBB`` format.
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    max_workers : int or None
        Number of parallel workers.

    Returns
    -------
    BatchResult
        Summary of the batch operation.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    ValueError
        If *old_hex* or *new_hex* is not in ``#RRGGBB`` format.
    """
    # Validate hex colors upfront
    for value in (old_hex, new_hex):
        if not value.startswith("#") or len(value) != 7:
            raise ValueError(f"Expected #RRGGBB color: {value}")

    files = _find_pptx_files(input_dir, pattern, recursive=False)
    out_dir = Path(output_dir).resolve()

    return _run_batch(
        files,
        out_dir,
        _recolor_operation,
        on_error="skip",
        max_workers=max_workers,
        progress=True,
        cpu_bound=False,
        old_hex=old_hex,
        new_hex=new_hex,
    )


# ---------------------------------------------------------------------------
# Public API — batch_export_images
# ---------------------------------------------------------------------------

def batch_export_images(
    input_dir: str,
    output_dir: str,
    *,
    dpi: int = 150,
    format: str = "png",
    pattern: str = "*.pptx",
    max_workers: int | None = None,
) -> BatchResult:
    """Export all slides from all PPTX files as images.

    For each input file ``report.pptx``, images are written to
    ``<output_dir>/report/slide_001.png``, etc.

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    output_dir : str
        Directory for image output.  Created if it does not exist.
    dpi : int
        Resolution in dots per inch (default 150).
    format : str
        Image format: ``"png"``, ``"jpg"``, or ``"bmp"``.
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    max_workers : int or None
        Number of parallel workers.

    Returns
    -------
    BatchResult
        Summary of the batch operation.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    """
    files = _find_pptx_files(input_dir, pattern, recursive=False)
    out_dir = Path(output_dir).resolve()

    # Image rendering is CPU-bound
    return _run_batch(
        files,
        out_dir,
        _export_images_operation,
        on_error="skip",
        max_workers=max_workers,
        progress=True,
        cpu_bound=True,
        dpi=dpi,
        format=format,
    )


# ---------------------------------------------------------------------------
# Public API — batch_inspect
# ---------------------------------------------------------------------------

def batch_inspect(
    input_dir: str,
    *,
    pattern: str = "*.pptx",
    recursive: bool = False,
    max_workers: int | None = None,
) -> list[dict]:
    """Inspect all PPTX files and return structured info dicts.

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    recursive : bool
        Search subdirectories recursively.
    max_workers : int or None
        Number of parallel workers.

    Returns
    -------
    list[dict]
        One dict per file with keys ``file``, ``filename``, ``slide_count``,
        ``shape_count``, ``image_count``, ``media_count``, ``theme``,
        ``layouts``.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    files = _find_pptx_files(input_dir, pattern, recursive)
    if not files:
        return []

    workers = max_workers or _default_max_workers()
    workers = min(workers, len(files))

    results: list[dict] = [None] * len(files)  # type: ignore[list-item]

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_idx = {
            executor.submit(_inspect_single, f): i
            for i, f in enumerate(files)
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results[idx] = future.result()
            except Exception as exc:
                results[idx] = {
                    "file": str(files[idx]),
                    "filename": files[idx].name,
                    "error": str(exc),
                }

    # Filter out None entries (should not happen, but defensive)
    return [r for r in results if r is not None]


# ---------------------------------------------------------------------------
# Public API — batch_stats
# ---------------------------------------------------------------------------

def batch_stats(
    input_dir: str,
    *,
    pattern: str = "*.pptx",
    recursive: bool = False,
) -> BatchStats:
    """Compute aggregate statistics across all PPTX files.

    Parameters
    ----------
    input_dir : str
        Directory containing PPTX files.
    pattern : str
        Glob pattern for input files (default ``"*.pptx"``).
    recursive : bool
        Search subdirectories recursively.

    Returns
    -------
    BatchStats
        Aggregated statistics.

    Raises
    ------
    FileNotFoundError
        If *input_dir* does not exist.
    """
    from pptx import Presentation

    files = _find_pptx_files(input_dir, pattern, recursive)
    if not files:
        return BatchStats()

    total_slides = 0
    total_shapes = 0
    total_images = 0
    total_media = 0
    themes: dict[str, int] = {}
    layouts: dict[str, int] = {}

    for file_path in files:
        try:
            prs = Presentation(str(file_path))
        except Exception as exc:
            log.warning("Could not open %s: %s", file_path, exc)
            continue

        total_slides += len(prs.slides)

        for slide in prs.slides:
            for shape in slide.shapes:
                total_shapes += 1
                # Picture type = 13
                if getattr(shape, "shape_type", None) == 13:
                    total_images += 1
                # Media types
                shape_type_val = getattr(shape, "shape_type", None)
                if shape_type_val is not None:
                    try:
                        st = int(shape_type_val)
                        if st in (16, 19):
                            total_media += 1
                    except (TypeError, ValueError):
                        pass
            # Layout name
            try:
                layout_name = slide.slide_layout.name
                layouts[layout_name] = layouts.get(layout_name, 0) + 1
            except Exception:
                pass

        # Theme name (best-effort)
        try:
            for master in prs.slide_masters:
                if master.slide_layouts:
                    theme_name = master.slide_layouts[0].name.split(" ")[0]
                    themes[theme_name] = themes.get(theme_name, 0) + 1
                    break
        except Exception:
            pass

    avg_slides = total_slides / len(files) if files else 0.0

    return BatchStats(
        total_files=len(files),
        total_slides=total_slides,
        total_shapes=total_shapes,
        total_images=total_images,
        total_media=total_media,
        avg_slide_count=round(avg_slides, 1),
        themes_used=themes,
        layouts_used=layouts,
    )
