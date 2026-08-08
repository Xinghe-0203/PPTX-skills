"""Shared OOXML namespace constants.

Previously duplicated across 30+ files with inconsistent naming
(``_P_NS`` vs ``_NS_P`` vs ``_PML_NS``).  This module provides a single
source of truth.

Usage::

    from pptx_skill.constants import P_NS, A_NS, R_NS, P_NS_PREFIX

All constants use bare names (no leading underscore) so they can be
imported cleanly.  ``*_PREFIX`` variants are Clark-notation braces
suitable for ``lxml`` tag matching.
"""
from __future__ import annotations

__all__ = [
    # Core OOXML namespaces
    "P_NS",
    "A_NS",
    "R_NS",
    "REL_NS",
    # Extended namespaces
    "P14_NS",
    "P15_NS",
    # Clark-notation prefixes
    "P_NS_PREFIX",
    "A_NS_PREFIX",
    "R_NS_PREFIX",
    "REL_NS_PREFIX",
    "P14_NS_PREFIX",
    "P15_NS_PREFIX",
]

# ---------------------------------------------------------------------------
# Core OOXML namespaces
# ---------------------------------------------------------------------------

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# ---------------------------------------------------------------------------
# Extended namespaces (PowerPoint 2010 / 2012)
# ---------------------------------------------------------------------------

P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P15_NS = "http://schemas.microsoft.com/office/powerpoint/2012/main"

# ---------------------------------------------------------------------------
# Clark-notation prefixes (for lxml tag matching)
# ---------------------------------------------------------------------------

P_NS_PREFIX = f"{{{P_NS}}}"
A_NS_PREFIX = f"{{{A_NS}}}"
R_NS_PREFIX = f"{{{R_NS}}}"
REL_NS_PREFIX = f"{{{REL_NS}}}"
P14_NS_PREFIX = f"{{{P14_NS}}}"
P15_NS_PREFIX = f"{{{P15_NS}}}"
