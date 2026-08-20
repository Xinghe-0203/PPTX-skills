"""Small helpers shared by directly executable command-line scripts."""

from __future__ import annotations

import sys
from typing import TextIO


def configure_utf8_console(
    stdout: TextIO | None = None,
    stderr: TextIO | None = None,
) -> None:
    """Make CLI output deterministic UTF-8 when the stream supports it.

    Windows commonly defaults Python's console streams to a legacy code page.
    That is fine in a classic terminal, but produces mojibake when another
    process captures the bytes as UTF-8.  ``reconfigure`` is unavailable on
    some wrapped streams, so this helper deliberately degrades to a no-op.
    """

    for stream in (stdout or sys.stdout, stderr or sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if not callable(reconfigure):
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError):
            continue
