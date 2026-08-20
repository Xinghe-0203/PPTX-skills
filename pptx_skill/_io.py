"""Shared I/O helpers for presentation open / save / path resolution.

These utilities were previously duplicated across 25+ extension modules.
Centralising them here eliminates ~70 copy-paste function definitions and
ensures consistent backup behaviour.

Public API
----------
- :func:`is_presentation` — duck-type check for ``pptx.Presentation``
- :func:`open_prs` — accept ``Presentation | str | Path``, return ``Presentation``
- :func:`save_prs` — save with optional ``.bak.pptx`` backup
- :func:`resolve_path` — return file path if *prs_or_path* is a path, else ``None``
- :func:`ensure_path_on_disk` — save to temp file if needed, return ``(path, tmp_dir)``
"""
from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from typing import Any

__all__ = [
    "is_presentation",
    "open_prs",
    "save_prs",
    "resolve_path",
    "ensure_path_on_disk",
]


def is_presentation(obj: Any) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import.

    Uses duck-typing on the class name and module to avoid importing
    ``python-pptx`` at module load time.
    """
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def open_prs(prs_or_path: Any) -> Any:
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path
    (``str | Path``).  Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def save_prs(prs: Any, path: str | Path | None, *, backup: bool = True) -> None:
    """Save *prs* back to *path*.

    When *backup* is ``True`` (the default), a ``.bak.pptx`` copy is created
    before overwriting.  The new package is first written to a temporary file
    in the destination directory and then atomically replaces the target, so
    a failed save cannot leave the original presentation partially written.
    When *path* is ``None`` the call is a no-op.
    """
    if path is None:
        return

    p = Path(path)
    if backup:
        bak = p.with_suffix(".bak.pptx")
        if p.exists():
            shutil.copy2(str(p), str(bak))

    suffix = p.suffix or ".pptx"
    with tempfile.NamedTemporaryFile(
        prefix=f".{p.stem}.",
        suffix=suffix,
        dir=str(p.parent),
        delete=False,
    ) as tmp_file:
        tmp_path = Path(tmp_file.name)

    try:
        prs.save(str(tmp_path))
        os.replace(str(tmp_path), str(p))
    finally:
        tmp_path.unlink(missing_ok=True)


def resolve_path(prs_or_path: Any) -> str | None:
    """Return the file path if *prs_or_path* is a path, else ``None``."""
    if is_presentation(prs_or_path):
        return None
    return str(prs_or_path)


def ensure_path_on_disk(prs_or_path: Any) -> tuple[str, str | None]:
    """Return a file path on disk, saving to a temp file if needed.

    Returns a tuple ``(path, tmp_dir)`` where *tmp_dir* is the temporary
    directory created (or ``None`` if no temp was needed).  The caller is
    responsible for cleaning up *tmp_dir* via ``shutil.rmtree(tmp_dir,
    ignore_errors=True)`` when it is not ``None``.
    """
    path = resolve_path(prs_or_path)
    if path is not None:
        return path, None

    prs = open_prs(prs_or_path)
    tmp_dir = tempfile.mkdtemp(prefix="pptx_skill_tmp_")
    tmp_path = os.path.join(tmp_dir, "work.pptx")
    try:
        prs.save(tmp_path)
    except BaseException:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise
    return tmp_path, tmp_dir
