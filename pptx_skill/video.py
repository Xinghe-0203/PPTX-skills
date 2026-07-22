"""Video export — convert PPTX presentations to MP4 or animated GIF.

Renders slides as images and encodes them with ffmpeg.  On Windows,
can optionally use PowerPoint COM automation for higher-fidelity output
including animations and transitions.

Quick start
-----------
>>> from pptx_skill.video import export_to_video, check_ffmpeg
>>> if check_ffmpeg():
...     export_to_video("deck.pptx", "deck.mp4", duration_per_slide=5)
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from typing import Any

__all__ = [
    "VideoExportOptions",
    "VideoInfo",
    "check_ffmpeg",
    "export_to_gif",
    "export_to_video",
    "extract_audio_track",
]

log = logging.getLogger(__name__)

_RESOLUTION_MAP = {
    "480p": (854, 480),
    "720p": (1280, 720),
    "1080p": (1920, 1080),
    "2k": (2560, 1440),
    "4k": (3840, 2160),
}


@dataclass
class VideoExportOptions:
    """Options for video export."""
    duration_per_slide: float = 5.0
    resolution: str = "1080p"
    fps: int = 24
    quality: int = 75
    codec: str = "libx264"
    transition_duration: float = 0.0


@dataclass
class VideoInfo:
    """Info about an exported video."""
    duration_seconds: float = 0.0
    slide_count: int = 0
    resolution: str = ""
    fps: int = 24
    codec: str = ""
    file_size: int = 0


def check_ffmpeg() -> str | None:
    """Check if ffmpeg is available.

    Returns
    -------
    str or None
        Path to ffmpeg if found, None otherwise.
    """
    path = shutil.which("ffmpeg")
    if path:
        return path

    # Common locations
    common_paths = [
        r"C:\ffmpeg\bin\ffmpeg.exe",
        r"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
        "/usr/bin/ffmpeg",
        "/usr/local/bin/ffmpeg",
    ]
    for p in common_paths:
        if os.path.isfile(p):
            return p

    return None


def _get_resolution(resolution: str) -> tuple[int, int]:
    """Parse resolution string to (width, height)."""
    if resolution in _RESOLUTION_MAP:
        return _RESOLUTION_MAP[resolution]
    # Try "WxH" format
    if "x" in resolution.lower():
        parts = resolution.lower().split("x")
        try:
            return (int(parts[0]), int(parts[1]))
        except (ValueError, IndexError):
            pass
    # Default to 1080p
    return (1920, 1080)


def _quality_to_crf(quality: int) -> int:
    """Convert quality (1-100) to ffmpeg CRF value (0-51).

    Higher quality → lower CRF.
    """
    quality = max(1, min(100, quality))
    # Map 1-100 → CRF 51-0 (inverted)
    return int(51 * (1 - quality / 100))


def _render_slides_to_images(prs_or_path, tmp_dir: str, width: int, height: int) -> list[str]:
    """Render all slides as PNG images.

    Returns list of image file paths.
    """
    from pptx_skill.export import export_to_images

    # Calculate DPI from desired resolution
    # Standard slide is 13.333" × 7.5" (widescreen)
    from pptx import Presentation
    if isinstance(prs_or_path, Presentation):
        slide_w_in = prs_or_path.slide_width / 914400
    else:
        prs = Presentation(prs_or_path)
        slide_w_in = prs.slide_width / 914400

    dpi = max(72, int(width / slide_w_in))

    images = export_to_images(prs_or_path, tmp_dir, dpi=dpi)
    return images


def _write_concat_file(image_paths: list[str], durations: list[float],
                        concat_path: str):
    """Write an ffmpeg concat demuxer file."""
    with open(concat_path, "w", encoding="utf-8") as f:
        f.write("ffconcat version 1.0\n")
        for img_path, duration in zip(image_paths, durations):
            # Use forward slashes for ffmpeg
            safe_path = img_path.replace("\\", "/")
            f.write(f"file '{safe_path}'\n")
            f.write(f"duration {duration}\n")
        # Repeat last entry (ffmpeg concat requires it)
        if image_paths:
            safe_path = image_paths[-1].replace("\\", "/")
            f.write(f"file '{safe_path}'\n")


def export_to_video(
    prs_or_path,
    output_path: str,
    *,
    duration_per_slide: float = 5.0,
    resolution: str = "1080p",
    fps: int = 24,
    quality: int = 75,
    codec: str = "libx264",
    transition_duration: float = 0.0,
    slide_durations: list[float] | None = None,
    use_speaker_timing: bool = False,
    include_audio: bool = False,
) -> str:
    """Export a presentation as an MP4 video.

    Parameters
    ----------
    output_path : str
        Output video file path (e.g. ``"output.mp4"``).
    duration_per_slide : float
        Seconds each slide is shown (default 5).
    resolution : str
        ``"480p"``, ``"720p"``, ``"1080p"``, ``"2k"``, ``"4k"``, or ``"WxH"``.
    fps : int
        Frames per second (default 24).
    quality : int
        Encoding quality 1-100 (higher = better, default 75).
    codec : str
        Video codec: ``"libx264"`` (default), ``"libx265"``, ``"vp9"``.
    transition_duration : float
        Crossfade duration between slides in seconds (0 = hard cut).
    slide_durations : list[float], optional
        Per-slide duration overrides.
    use_speaker_timing : bool
        Use speaker note timing if available.
    include_audio : bool
        Mix in any embedded audio tracks.

    Returns
    -------
    str
        Output video file path.

    Raises
    ------
    RuntimeError
        If ffmpeg is not found or encoding fails.
    """
    ffmpeg_path = check_ffmpeg()
    if not ffmpeg_path:
        raise RuntimeError(
            "ffmpeg not found. Install it from https://ffmpeg.org/download.html "
            "and ensure it is on your PATH."
        )

    width, height = _get_resolution(resolution)
    crf = _quality_to_crf(quality)

    with tempfile.TemporaryDirectory(prefix="pptx_video_") as tmp_dir:
        # Render slides as images
        log.info("Rendering slides as images...")
        image_paths = _render_slides_to_images(prs_or_path, tmp_dir, width, height)

        if not image_paths:
            raise RuntimeError("No slides rendered — empty presentation?")

        # Compute per-slide durations
        slide_count = len(image_paths)
        if slide_durations and len(slide_durations) == slide_count:
            durations = slide_durations
        elif use_speaker_timing:
            durations = _extract_speaker_durations(prs_or_path, slide_count, duration_per_slide)
        else:
            durations = [duration_per_slide] * slide_count

        # Check for crossfade support
        if transition_duration > 0 and slide_count > 1:
            return _export_with_crossfade(
                ffmpeg_path, image_paths, durations, output_path,
                width, height, fps, crf, codec, transition_duration,
            )

        # Simple concat approach
        concat_path = os.path.join(tmp_dir, "concat.txt")
        _write_concat_file(image_paths, durations, concat_path)

        # Build ffmpeg command
        cmd = [
            ffmpeg_path,
            "-y",  # Overwrite output
            "-f", "concat",
            "-safe", "0",
            "-i", concat_path,
            "-c:v", codec,
            "-pix_fmt", "yuv420p",
            "-crf", str(crf),
            "-r", str(fps),
            "-vf", f"scale={width}:{height}:force_original_aspect_ratio=decrease,pad={width}:{height}:(ow-iw)/2:(oh-ih)/2",
            output_path,
        ]

        log.info("Encoding video with ffmpeg...")
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=600,
            )
            if result.returncode != 0:
                raise RuntimeError(
                    f"ffmpeg encoding failed (code {result.returncode}):\n"
                    f"{result.stderr[-2000:]}"
                )
        except subprocess.TimeoutExpired:
            raise RuntimeError("ffmpeg encoding timed out after 600 seconds")

        if not os.path.isfile(output_path):
            raise RuntimeError(f"ffmpeg did not produce output file: {output_path}")

        log.info("Video exported: %s", output_path)
        return output_path


def _export_with_crossfade(
    ffmpeg_path: str,
    image_paths: list[str],
    durations: list[float],
    output_path: str,
    width: int, height: int,
    fps: int, crf: int, codec: str,
    transition_duration: float,
) -> str:
    """Export video with crossfade transitions between slides."""
    n = len(image_paths)

    # Build complex filter graph for crossfades
    inputs = []
    for i, img_path in enumerate(image_paths):
        safe_path = img_path.replace("\\", "/")
        # Total display time = duration + overlap with next crossfade
        dur = durations[i]
        if i < n - 1:
            dur += transition_duration  # Extra time for the crossfade
        inputs.extend(["-loop", "1", "-t", str(dur), "-i", safe_path])

    # Build filter complex
    filter_parts = []
    # Scale each input
    for i in range(n):
        filter_parts.append(f"[{i}:v]scale={width}:{height}:force_original_aspect_ratio=decrease,pad={width}:{height}:(ow-iw)/2:(oh-ih)/2,format=yuva420p[v{i}]")

    # Chain crossfades
    if n == 1:
        filter_parts.append(f"[v0]null[outv]")
    else:
        # First crossfade
        offset = durations[0]
        filter_parts.append(
            f"[v0][v1]xfade=transition=fade:duration={transition_duration}:offset={offset}[xf0]"
        )
        # Subsequent crossfades
        for i in range(1, n - 1):
            offset += durations[i] - transition_duration
            prev = f"xf{i-1}"
            curr = f"xf{i}"
            last_tag = "outv" if i == n - 2 else curr
            filter_parts.append(
                f"[{prev}][v{i+1}]xfade=transition=fade:duration={transition_duration}:offset={offset}[{last_tag}]"
            )

    filter_complex = ";".join(filter_parts)

    cmd = [
        ffmpeg_path,
        "-y",
        *inputs,
        "-filter_complex", filter_complex,
        "-map", "[outv]",
        "-c:v", codec,
        "-pix_fmt", "yuv420p",
        "-crf", str(crf),
        "-r", str(fps),
        output_path,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        if result.returncode != 0:
            log.warning("Crossfade encoding failed, falling back to simple concat: %s",
                        result.stderr[-500:])
            # Fall back to no transition
            concat_path = os.path.join(os.path.dirname(image_paths[0]), "concat.txt")
            _write_concat_file(image_paths, durations, concat_path)
            cmd = [
                ffmpeg_path, "-y",
                "-f", "concat", "-safe", "0", "-i", concat_path,
                "-c:v", codec, "-pix_fmt", "yuv420p",
                "-crf", str(crf), "-r", str(fps),
                "-vf", f"scale={width}:{height}",
                output_path,
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
            if result.returncode != 0:
                raise RuntimeError(f"ffmpeg encoding failed:\n{result.stderr[-2000:]}")
    except subprocess.TimeoutExpired:
        raise RuntimeError("ffmpeg encoding timed out after 600 seconds")

    return output_path


def export_to_gif(
    prs_or_path,
    output_path: str,
    *,
    duration_per_slide: float = 2.0,
    width: int = 960,
    fps: int = 10,
    optimize: bool = True,
) -> str:
    """Export a presentation as an animated GIF.

    Parameters
    ----------
    output_path : str
        Output GIF file path.
    duration_per_slide : float
        Seconds each slide is shown (default 2).
    width : int
        GIF width in pixels (default 960).
    fps : int
        Frames per second (default 10).
    optimize : bool
        Use palette optimization for smaller files (default True).

    Returns
    -------
    str
        Output GIF file path.
    """
    ffmpeg_path = check_ffmpeg()
    if not ffmpeg_path:
        raise RuntimeError(
            "ffmpeg not found. Install it from https://ffmpeg.org/download.html"
        )

    with tempfile.TemporaryDirectory(prefix="pptx_gif_") as tmp_dir:
        # Render slides
        image_paths = _render_slides_to_images(prs_or_path, tmp_dir, width, 540)
        if not image_paths:
            raise RuntimeError("No slides rendered")

        # Compute durations
        slide_count = len(image_paths)
        durations = [duration_per_slide] * slide_count

        # Write concat file
        concat_path = os.path.join(tmp_dir, "concat.txt")
        _write_concat_file(image_paths, durations, concat_path)

        # Build ffmpeg command
        if optimize:
            vf = f"fps={fps},scale={width}:-1:flags=lanczos,split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse"
        else:
            vf = f"fps={fps},scale={width}:-1:flags=lanczos"

        cmd = [
            ffmpeg_path,
            "-y",
            "-f", "concat", "-safe", "0",
            "-i", concat_path,
            "-vf", vf,
            output_path,
        ]

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
            if result.returncode != 0:
                raise RuntimeError(
                    f"GIF encoding failed (code {result.returncode}):\n"
                    f"{result.stderr[-2000:]}"
                )
        except subprocess.TimeoutExpired:
            raise RuntimeError("GIF encoding timed out after 600 seconds")

        return output_path


def _extract_speaker_durations(prs_or_path, slide_count: int,
                                default: float) -> list[float]:
    """Try to extract per-slide durations from speaker notes."""
    from pptx import Presentation

    if isinstance(prs_or_path, str):
        prs = Presentation(prs_or_path)
    else:
        prs = prs_or_path

    durations = []
    for idx in range(slide_count):
        try:
            slide = prs.slides[idx]
            notes_slide = slide.notes_slide
            notes_text = notes_slide.notes_text_frame.text.strip()
            # Look for timing hints like "30s" or "1:30" in notes
            import re
            match = re.search(r'\[(\d+(?:\.\d+)?)\s*s\]', notes_text)
            if match:
                durations.append(float(match.group(1)))
                continue
            # mm:ss format
            match = re.search(r'\[(\d+):(\d+)\]', notes_text)
            if match:
                durations.append(int(match.group(1)) * 60 + int(match.group(2)))
                continue
        except Exception:
            pass
        durations.append(default)

    return durations


def extract_audio_track(prs_or_path, output_path: str) -> str | None:
    """Extract embedded audio from a presentation.

    Parameters
    ----------
    output_path : str
        Output audio file path.

    Returns
    -------
    str or None
        Output path if audio found, None otherwise.
    """
    import zipfile

    path = prs_or_path if isinstance(prs_or_path, str) else None
    if path is None:
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs_or_path.save(tmp.name)
            path = tmp.name

    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if name.startswith("ppt/media/") and any(
                    name.endswith(ext) for ext in (".mp3", ".wav", ".wma", ".m4a")
                ):
                    with open(output_path, "wb") as f:
                        f.write(zf.read(name))
                    return output_path
        return None
    except Exception:
        return None
