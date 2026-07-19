"""Image cropping utilities for adaptive layout (PR2).

Provides ``contain``, ``cover`` and ``smart`` crop strategies. All results are
expressed as normalized crop fractions compatible with PowerPoint picture
crop properties.
"""
from __future__ import annotations

import math
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CropResult:
    """Normalized crop result.

    - ``focus`` is the salient point in normalized [0, 1] image coordinates.
    - ``crop_rect`` is ``(left, top, right, bottom)`` in normalized coordinates.
    - ``crop_fractions`` maps directly to python-pptx ``Picture.crop_*``.
    """

    focus: tuple[float, float]
    crop_rect: tuple[float, float, float, float]
    crop_fractions: dict[str, float]
    score: float

    @property
    def left(self) -> float:
        return self.crop_rect[0]

    @property
    def top(self) -> float:
        return self.crop_rect[1]

    @property
    def right(self) -> float:
        return self.crop_rect[2]

    @property
    def bottom(self) -> float:
        return self.crop_rect[3]


def _to_fractions(rect: tuple[float, float, float, float]) -> dict[str, float]:
    left, top, right, bottom = rect
    return {
        "left": max(0.0, left),
        "top": max(0.0, top),
        "right": max(0.0, 1.0 - right),
        "bottom": max(0.0, 1.0 - bottom),
    }


def crop_contain(
    src_width: float,
    src_height: float,
    dst_width: float,
    dst_height: float,
) -> CropResult:
    """Fit the whole image inside the destination box (letterbox)."""
    if src_height == 0 or src_width == 0 or dst_height == 0 or dst_width == 0:
        return CropResult(focus=(0.5, 0.5), crop_rect=(0.0, 0.0, 1.0, 1.0), crop_fractions=_to_fractions((0.0, 0.0, 1.0, 1.0)), score=0.0)
    src_ratio = src_width / src_height
    dst_ratio = dst_width / dst_height
    if dst_ratio > src_ratio:
        # Destination is wider: fit to height, black bars left/right.
        new_height = src_height
        new_width = src_height * dst_ratio
    else:
        # Destination is taller or equal: fit to width, black bars top/bottom.
        new_width = src_width
        new_height = src_width / dst_ratio
    left = max(0.0, (src_width - new_width) / 2 / src_width)
    top = max(0.0, (src_height - new_height) / 2 / src_height)
    right = min(1.0, left + new_width / src_width)
    bottom = min(1.0, top + new_height / src_height)
    rect = (left, top, right, bottom)
    return CropResult(
        focus=(0.5, 0.5),
        crop_rect=rect,
        crop_fractions=_to_fractions(rect),
        score=1.0,
    )


def crop_cover(
    src_width: float,
    src_height: float,
    dst_width: float,
    dst_height: float,
) -> CropResult:
    """Cover the destination box while preserving image aspect ratio."""
    if src_height == 0 or src_width == 0 or dst_height == 0 or dst_width == 0:
        return CropResult(focus=(0.5, 0.5), crop_rect=(0.0, 0.0, 1.0, 1.0), crop_fractions=_to_fractions((0.0, 0.0, 1.0, 1.0)), score=0.0)
    src_ratio = src_width / src_height
    dst_ratio = dst_width / dst_height
    if dst_ratio > src_ratio:
        # Destination is wider: cover width, crop top/bottom.
        new_width = src_width
        new_height = src_width / dst_ratio
    else:
        # Destination is taller or equal: cover height, crop left/right.
        new_height = src_height
        new_width = src_height * dst_ratio
    left = max(0.0, (src_width - new_width) / 2 / src_width)
    top = max(0.0, (src_height - new_height) / 2 / src_height)
    right = min(1.0, left + new_width / src_width)
    bottom = min(1.0, top + new_height / src_height)
    rect = (left, top, right, bottom)
    return CropResult(
        focus=(0.5, 0.5),
        crop_rect=rect,
        crop_fractions=_to_fractions(rect),
        score=1.0,
    )


def _load_and_resize(image_path: str | Path, max_side: int = 512):
    from PIL import Image

    img = Image.open(image_path).convert("RGB")
    original_size = img.size
    w, h = original_size
    if max(w, h) > max_side:
        scale = max_side / max(w, h)
        new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    return img, original_size


def _saliency_map(img):
    """Build a simple saliency map using edges, saturation and skin tones."""
    from PIL import ImageFilter

    w, h = img.size
    # Edge response
    edges = img.convert("L").filter(ImageFilter.FIND_EDGES)
    _gd = "get_flattened_data" if hasattr(edges, "get_flattened_data") else "getdata"
    edge_data = list(getattr(edges, _gd)())
    # Saturation
    hsv = img.convert("HSV")
    sat_data = [px[1] for px in getattr(hsv, _gd)()]
    # Skin tone approximation in YCbCr-ish RGB heuristic
    skin_data = []
    for r, g, b in getattr(img, _gd)():
        skin_score = 0
        if r > g and r > b and abs(r - g) > 15:
            skin_score = (r - g) + (r - b)
        skin_data.append(skin_score)

    max_edge = max(edge_data) or 1
    max_sat = max(sat_data) or 1
    max_skin = max(skin_data) or 1

    saliency = []
    for e, s, sk in zip(edge_data, sat_data, skin_data, strict=False):
        score = 0.5 * (e / max_edge) + 0.3 * (s / max_sat) + 0.2 * (sk / max_skin)
        saliency.append(score)
    return saliency, (w, h)


def _window_score(
    saliency: list[float],
    size: tuple[int, int],
    left: int,
    top: int,
    right: int,
    bottom: int,
    boost_boxes: Sequence[tuple[float, float, float, float]],
) -> float:
    w, h = size
    total = 0.0
    weights = 0.0
    for y in range(top, bottom):
        row_offset = y * w
        cy = (y + 0.5) / h
        for x in range(left, right):
            cx = (x + 0.5) / w
            base = saliency[row_offset + x]
            # Center bias
            dx = abs(cx - 0.5)
            dy = abs(cy - 0.5)
            center_weight = 1.0 - 0.5 * math.sqrt(dx * dx + dy * dy)
            # Rule of thirds intersection boost
            thirds = 0.0
            for tx in (1 / 3, 2 / 3):
                for ty in (1 / 3, 2 / 3):
                    d = math.hypot(cx - tx, cy - ty)
                    thirds += max(0, 1 - d * 3)
            thirds_weight = 1.0 + 0.3 * thirds
            # Boost boxes
            boost = 1.0
            for bx1, by1, bx2, by2 in boost_boxes:
                if bx1 <= cx <= bx2 and by1 <= cy <= by2:
                    boost += 2.0
            value = base * center_weight * thirds_weight * boost
            total += value
            weights += center_weight * thirds_weight * boost
    if weights == 0:
        return 0.0
    # Penalize cropping too close to edges or extreme aspect ratios.
    (right - left) * (bottom - top)
    edge_penalty = (
        (left / w) ** 2
        + (top / h) ** 2
        + ((w - right) / w) ** 2
        + ((h - bottom) / h) ** 2
    ) * 0.05
    return total / weights - edge_penalty


def crop_smart(
    image_path: str | Path,
    dst_width: float,
    dst_height: float,
    boost_boxes: Sequence[tuple[float, float, float, float]] | None = None,
    max_side: int = 512,
    candidates: int = 40,
) -> CropResult:
    """Smart crop an image to the destination aspect ratio.

    ``boost_boxes`` are normalized ``(left, top, right, bottom)`` regions that
    should be preserved (e.g. detected faces).
    """
    if dst_width <= 0 or dst_height <= 0:
        raise ValueError("Destination dimensions must be positive")

    img, (orig_w, orig_h) = _load_and_resize(image_path, max_side=max_side)
    saliency, (sw, sh) = _saliency_map(img)
    src_ratio = orig_w / orig_h
    dst_ratio = dst_width / dst_height

    # Candidate window sizes: from full image to a reasonable minimum.
    best_score = -1e9
    best = (0.0, 0.0, 1.0, 1.0)

    boost = list(boost_boxes) if boost_boxes else []

    for scale in [1.0, 0.9, 0.8, 0.7, 0.6, 0.5]:
        if dst_ratio > src_ratio:
            window_h = int(sh * scale)
            window_w = int(window_h * dst_ratio)
        else:
            window_w = int(sw * scale)
            window_h = int(window_w / dst_ratio)
        if window_w < 1 or window_h < 1 or window_w > sw or window_h > sh:
            continue
        # Scan positions with slight overlap.
        x_steps = max(1, candidates // 6)
        y_steps = max(1, candidates // 6)
        for xi in range(x_steps):
            left = int((sw - window_w) * xi / max(1, x_steps - 1))
            for yi in range(y_steps):
                top = int((sh - window_h) * yi / max(1, y_steps - 1))
                right = left + window_w
                bottom = top + window_h
                if right > sw or bottom > sh:
                    continue
                score = _window_score(saliency, (sw, sh), left, top, right, bottom, boost)
                if score > best_score:
                    best_score = score
                    best = (left / sw, top / sh, right / sw, bottom / sh)

    if best_score == -1e9:
        return crop_cover(orig_w, orig_h, dst_width, dst_height)

    left, top, right, bottom = best
    focus_x = (left + right) / 2
    focus_y = (top + bottom) / 2
    return CropResult(
        focus=(focus_x, focus_y),
        crop_rect=best,
        crop_fractions=_to_fractions(best),
        score=best_score,
    )
