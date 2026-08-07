"""Cross-version compatibility shims.

Provides ``StrEnum`` on Python < 3.11 so that the rest of the package can use
``from enum import StrEnum`` uniformly. On 3.11+ the stdlib version is re-exported
unchanged.

The fallback subclasses ``(str, Enum)`` which preserves the string value and
makes members comparable to plain strings -- the two behaviours this package
relies on.
"""
from __future__ import annotations

import sys

if sys.version_info >= (3, 11):  # pragma: no cover - version-gated
    from enum import StrEnum as StrEnum
else:  # pragma: no cover - version-gated
    from enum import Enum

    class StrEnum(str, Enum):
        """Backport of :class:`enum.StrEnum` for Python 3.10."""

        def __str__(self) -> str:
            return self.value

        @classmethod
        def _missing_(cls, value):
            if isinstance(value, str):
                for member in cls:
                    if member.value == value:
                        return member
            return None


__all__ = ["StrEnum"]
