"""PPTX skill public package (PR0 shim).

This package exposes the existing scripts/ modules through a stable namespace
while the V2 migration gradually moves authoritative implementations here.
"""
from __future__ import annotations

import sys
from pathlib import Path

_SKILL_ROOT = Path(__file__).resolve().parent.parent
_SCRIPTS_DIR = _SKILL_ROOT / "scripts"

if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from pptx_helper import (  # noqa: E402
    LAYOUT_REGISTRY,
    THEMES,
    Section,
    add_chart_slide,
    auto_generate_ppt,
    auto_validate_ppt,
    choose_layout,
    choose_theme,
)
from ppt_edit import (  # noqa: E402
    edit_kicker,
    edit_text,
    edit_text_by_role,
    edit_title,
    recolor,
    swap_image,
    swap_theme,
)
from ppt_inspect import inspect_ppt  # noqa: E402
from ppt_pages import (  # noqa: E402
    delete_slide,
    duplicate_slide,
    insert_slide,
    move_slide,
    replace_layout,
)
from ppt_project import (  # noqa: E402
    create_backup,
    edit_section,
    load_project,
    regenerate,
    restore_backup,
)
from reference_ppt import (  # noqa: E402
    analyze_presentation,
    compose_from_reference,
    extract_template_profile,
    generate_from_reference,
)
from render_slides import render_slides  # noqa: E402
from template_engine import (  # noqa: E402
    generate_template_preview,
    generate_template_profile,
    list_templates,
    load_template_profile,
    register_template_profile,
    validate_template_profile,
)

__version__ = "2.0.0"


def __getattr__(name: str):
    """Lazy import for submodules to avoid circular imports on ``python -m``."""
    if name == "report_capabilities":
        from pptx_skill.capability import report_capabilities
        return report_capabilities
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    "LAYOUT_REGISTRY",
    "Section",
    "THEMES",
    "add_chart_slide",
    "analyze_presentation",
    "auto_generate_ppt",
    "auto_validate_ppt",
    "choose_layout",
    "choose_theme",
    "compose_from_reference",
    "create_backup",
    "delete_slide",
    "duplicate_slide",
    "edit_kicker",
    "edit_section",
    "edit_text",
    "edit_text_by_role",
    "edit_title",
    "extract_template_profile",
    "generate_from_reference",
    "generate_template_preview",
    "generate_template_profile",
    "inspect_ppt",
    "insert_slide",
    "list_templates",
    "load_project",
    "load_template_profile",
    "move_slide",
    "recolor",
    "regenerate",
    "register_template_profile",
    "render_slides",
    "replace_layout",
    "report_capabilities",
    "restore_backup",
    "swap_image",
    "swap_theme",
    "validate_template_profile",
]
