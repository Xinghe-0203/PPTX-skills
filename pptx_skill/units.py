"""Shared EMU conversion constants and helpers.

EMU (English Metric Unit) is the internal coordinate system used by
OOXML / python-pptx.  These constants were previously scattered as
magic numbers (``914400``, ``12700``) across 30+ files.

Usage::

    from pptx_skill.units import EMU_PER_PT, EMU_PER_INCH, emu_to_pt, pt_to_emu
"""
from __future__ import annotations

__all__ = [
    "EMU_PER_PT",
    "EMU_PER_INCH",
    "emu_to_pt",
    "pt_to_emu",
    "emu_to_inch",
    "inch_to_emu",
    "emu_to_cm",
    "cm_to_emu",
]

# 1 inch = 914400 EMU
EMU_PER_INCH = 914400

# 1 point = 12700 EMU
EMU_PER_PT = 12700

# 1 cm = 360000 EMU
EMU_PER_CM = 360000


def emu_to_pt(emu: int) -> float:
    """Convert EMU to points (1 pt = 12700 EMU)."""
    return emu / EMU_PER_PT


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU (1 pt = 12700 EMU)."""
    return int(pt * EMU_PER_PT)


def emu_to_inch(emu: int) -> float:
    """Convert EMU to inches (1 inch = 914400 EMU)."""
    return emu / EMU_PER_INCH


def inch_to_emu(inch: float) -> int:
    """Convert inches to EMU (1 inch = 914400 EMU)."""
    return int(inch * EMU_PER_INCH)


def emu_to_cm(emu: int) -> float:
    """Convert EMU to centimetres."""
    return emu / EMU_PER_CM


def cm_to_emu(cm: float) -> int:
    """Convert centimetres to EMU."""
    return int(cm * EMU_PER_CM)
