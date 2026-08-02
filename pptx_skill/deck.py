"""Unified high-level Deck API.

Wraps the 60+ module functions into a single fluent class so users don't
need to juggle 600 free functions with inconsistent input contracts.

The ``Deck`` class accepts a file path or an open ``Presentation`` object,
normalises 1-based slide indexing internally, manages file lifecycle
(open / save / backup), and exposes the most common operations as
chainable methods.

Quick start
-----------
>>> from pptx_skill import Deck
>>> deck = Deck.open("report.pptx")
>>> deck.add_watermark("DRAFT", opacity=0.15)      # fluent
>>> deck.add_notes(1, "Opening remarks")           # 1-based
>>> deck.add_transition(1, "fade", duration_ms=500)
>>> deck.save()

Generate from scratch:

>>> from pptx_skill import Deck
>>> deck = Deck.generate(title="Q3 Review", sections=[...])
>>> deck.save("output/q3.pptx")

Import from Markdown:

>>> deck = Deck.from_markdown("outline.md")
>>> deck.add_watermark("CONFIDENTIAL")
>>> deck.save("output/deck.pptx")
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from pptx import Presentation

__all__ = [
    "Deck",
    "open_deck",
]


def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


class Deck:
    """A fluent wrapper around an open presentation.

    Parameters
    ----------
    path : str
        Path to a ``.pptx`` file.  The file is opened on construction.
        Call :meth:`save` to persist changes.
    """

    def __init__(self, path: str) -> None:
        from pptx import Presentation
        self._path: str | None = str(path)
        self._prs: Presentation = Presentation(str(path))
        self._dirty = False

    # ------------------------------------------------------------------
    # Construction helpers
    # ------------------------------------------------------------------

    @classmethod
    def open(cls, path: str) -> "Deck":
        """Open an existing PPTX file for editing.

        Creates a ``.bak.pptx`` backup alongside the original before any
        modification is saved.
        """
        return cls(path)

    @classmethod
    def generate(
        cls,
        title: str,
        subtitle: str = "",
        sections: list[dict[str, Any]] | None = None,
        *,
        theme_key: str | None = None,
        template_key: str | None = None,
        lang: str = "zh",
        auto_search_images: bool = False,
        output_path: str = "output.pptx",
    ) -> "Deck":
        """Generate a new deck from sections and open it for further editing.

        Delegates to :func:`pptx_skill.auto_generate_ppt`, then wraps the
        result in a ``Deck`` for fluent post-processing.
        """
        from pptx_skill import auto_generate_ppt

        path = auto_generate_ppt(
            title=title,
            subtitle=subtitle,
            sections=sections or [],
            output_path=output_path,
            theme_key=theme_key,
            template_key=template_key,
            lang=lang,
            auto_search_images=auto_search_images,
        )
        return cls(path)

    @classmethod
    def from_markdown(
        cls,
        path: str,
        output_path: str = "output.pptx",
        *,
        title: str | None = None,
        subtitle: str = "",
        lang: str = "zh",
        theme_key: str | None = None,
        auto_search_images: bool = False,
    ) -> "Deck":
        """Generate a deck from a Markdown file and open it for editing.

        Delegates to :func:`pptx_skill.import_markdown`.
        """
        from pptx_skill.markdown_import import import_markdown

        out = import_markdown(
            path,
            output_path=output_path,
            title=title,
            subtitle=subtitle,
            lang=lang,
            theme_key=theme_key,
            auto_search_images=auto_search_images,
        )
        return cls(out)

    # ------------------------------------------------------------------
    # Properties
    # ------------------------------------------------------------------

    @property
    def path(self) -> str | None:
        """The file path this deck was opened from (or None if unsaved)."""
        return self._path

    @property
    def slide_count(self) -> int:
        """Number of slides in the deck."""
        return len(self._prs.slides)

    @property
    def presentation(self) -> "Presentation":
        """The underlying python-pptx ``Presentation`` object."""
        return self._prs

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def save(self, path: str | None = None) -> str:
        """Save the deck to disk.

        Parameters
        ----------
        path : str, optional
            Output path.  If None, saves back to the original path.

        Returns
        -------
        str
            The path the file was saved to.
        """
        out = str(path) if path else self._path
        if out is None:
            raise ValueError("No output path — pass a path to save() or set one first.")
        out_dir = os.path.dirname(out)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        self._prs.save(out)
        self._path = out
        self._dirty = False
        return out

    def save_as(self, path: str) -> str:
        """Save a copy to a new path (alias for ``save(path)``)."""
        return self.save(path)

    # ------------------------------------------------------------------
    # Slide inspection (1-based)
    # ------------------------------------------------------------------

    def inspect(self) -> dict[str, Any]:
        """Return a structured inspection of the deck.

        Delegates to :func:`pptx_skill.inspect_ppt`.
        """
        from pptx_skill import inspect_ppt
        # inspect_ppt reads from the file — save first if dirty.
        if self._dirty and self._path:
            self._prs.save(self._path)
        return inspect_ppt(self._path or "")

    def info(self) -> dict[str, Any]:
        """Return summary info: slide count, size, shape counts per slide."""
        prs = self._prs
        slides_info = []
        for idx, slide in enumerate(prs.slides, 1):
            slides_info.append({
                "index": idx,
                "shape_count": len(slide.shapes),
                "has_notes": slide.has_notes_slide and bool(slide.notes_slide.notes_text_frame.text.strip()),
            })
        return {
            "path": self._path,
            "slide_count": len(prs.slides),
            "width_in": prs.slide_width / 914400,
            "height_in": prs.slide_height / 914400,
            "slides": slides_info,
        }

    # ------------------------------------------------------------------
    # Text editing
    # ------------------------------------------------------------------

    def edit_text(self, slide_index: int, find: str, replace: str) -> int:
        """Replace text on a slide (1-based index). Returns replacement count."""
        from pptx_skill import edit_text
        self._dirty = True
        return edit_text(self._prs, slide_index, find, replace)

    def find_replace_all(self, find: str, replace: str) -> int:
        """Global find-and-replace across all slides. Returns total count."""
        from pptx_skill import find_replace_all
        self._dirty = True
        return find_replace_all(self._prs, find, replace)

    def edit_title(self, slide_index: int, new_title: str) -> int:
        """Set the title text on a slide (1-based). Returns count."""
        from pptx_skill import edit_title
        self._dirty = True
        return edit_title(self._prs, slide_index, new_title)

    # ------------------------------------------------------------------
    # Images
    # ------------------------------------------------------------------

    def swap_image(self, slide_index: int, image_index: int, new_path: str) -> bool:
        """Replace an image on a slide (1-based slide, 0-based image)."""
        from pptx_skill import swap_image
        self._dirty = True
        return swap_image(self._prs, slide_index, image_index, new_path)

    # ------------------------------------------------------------------
    # Color / theme
    # ------------------------------------------------------------------

    def recolor(self, old_hex: str, new_hex: str) -> int:
        """Global color replacement. Returns replacement count."""
        from pptx_skill import recolor
        self._dirty = True
        return recolor(self._prs, old_hex, new_hex)

    def swap_theme(self, new_theme_key: str) -> dict[str, Any]:
        """Swap the theme across the entire deck."""
        from pptx_skill import swap_theme
        self._dirty = True
        return swap_theme(self._prs, new_theme_key)

    # ------------------------------------------------------------------
    # Watermark
    # ------------------------------------------------------------------

    def add_watermark(
        self,
        text: str | None = None,
        *,
        image_path: str | None = None,
        opacity: float = 0.15,
        diagonal: bool = False,
        font_size: int = 48,
        slides: list[int] | None = None,
    ) -> "Deck":
        """Add a text or image watermark. Chainable.

        Parameters
        ----------
        text : str
            Watermark text (if text watermark).
        image_path : str
            Image path (if image watermark).
        opacity : float
            Opacity 0-1 (default 0.15).
        diagonal : bool
            Diagonal text watermark (default False).  When True, sets
            rotation=-45 degrees.
        font_size : int
            Font size for text watermark (default 48).
        slides : list[int], optional
            1-based slide indices to watermark.  None = all slides.
        """
        from pptx_skill import add_text_watermark, add_image_watermark
        # Save to a temp path first if needed, since watermark functions
        # accept a path or Presentation.
        rotation = -45.0 if diagonal else 0.0
        if text:
            add_text_watermark(
                self._prs, text=text, opacity=opacity,
                rotation=rotation, font_size=font_size,
                slides=slides,
            )
        elif image_path:
            add_image_watermark(
                self._prs, image_path=image_path, opacity=opacity,
                slides=slides,
            )
        else:
            raise ValueError("Either text or image_path must be provided")
        self._dirty = True
        return self

    # ------------------------------------------------------------------
    # Transitions & animations
    # ------------------------------------------------------------------

    def add_transition(
        self,
        slide_index: int | None = None,
        transition_type: str = "fade",
        duration_ms: int = 700,
    ) -> "Deck":
        """Add a slide transition (1-based index, or None for all slides).

        Chainable.
        """
        from pptx_skill.transitions import apply_slide_transition, apply_deck_transitions
        if slide_index is None:
            apply_deck_transitions(self._prs, transition_type, duration_ms)
        else:
            if slide_index < 1 or slide_index > len(self._prs.slides):
                raise IndexError(f"slide_index {slide_index} out of range (1..{len(self._prs.slides)})")
            apply_slide_transition(self._prs.slides[slide_index - 1], transition_type, duration_ms)
        self._dirty = True
        return self

    def add_animation(
        self,
        slide_index: int,
        shape_name: str | None = None,
        anim_type: str = "fade_in",
        shape_index: int | None = None,
    ) -> "Deck":
        """Add an animation to a shape (1-based slide index). Chainable."""
        from pptx_skill.animations import apply_entrance_animation
        if slide_index < 1 or slide_index > len(self._prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(self._prs.slides)})")
        slide = self._prs.slides[slide_index - 1]
        shape = None
        if shape_name is not None:
            for s in slide.shapes:
                if s.name == shape_name:
                    shape = s
                    break
        elif shape_index is not None:
            shapes = list(slide.shapes)
            if 0 <= shape_index < len(shapes):
                shape = shapes[shape_index]
        if shape is not None:
            apply_entrance_animation(slide, shape, anim_type)
        self._dirty = True
        return self

    # ------------------------------------------------------------------
    # Notes & comments
    # ------------------------------------------------------------------

    def add_notes(self, slide_index: int, text: str) -> "Deck":
        """Set speaker notes on a slide (1-based index). Chainable."""
        from pptx_skill import set_speaker_notes
        set_speaker_notes(self._prs, slide_index, text)
        self._dirty = True
        return self

    def add_comment(
        self,
        slide_index: int,
        text: str,
        *,
        author: str = "Claude",
    ) -> "Deck":
        """Add a review comment to a slide (1-based index). Chainable."""
        from pptx_skill import add_comment
        add_comment(self._prs, slide_index, text, author=author)
        self._dirty = True
        return self

    # ------------------------------------------------------------------
    # Sections
    # ------------------------------------------------------------------

    def add_section(self, name: str, start_slide: int = 1) -> "Deck":
        """Add a section group (1-based start slide). Chainable."""
        from pptx_skill import add_section
        add_section(self._prs, name, start_slide=start_slide)
        self._dirty = True
        return self

    # ------------------------------------------------------------------
    # Page operations (ppt_pages functions take paths, not Presentation)
    # ------------------------------------------------------------------

    def _save_if_dirty(self) -> None:
        """Persist to self._path if there are unsaved changes."""
        if self._dirty and self._path:
            self._prs.save(self._path)
            self._dirty = False

    def insert_slide(self, index: int, section: dict, layout: str | None = None) -> "Deck":
        """Insert a slide at 1-based index. Chainable."""
        from pptx_skill import insert_slide
        self._save_if_dirty()
        insert_slide(self._path, index, section, layout=layout)
        # Re-open to reflect the new slide.
        from pptx import Presentation
        self._prs = Presentation(self._path)
        return self

    def delete_slide(self, index: int) -> "Deck":
        """Delete a slide at 1-based index. Chainable."""
        from pptx_skill import delete_slide
        self._save_if_dirty()
        delete_slide(self._path, index)
        from pptx import Presentation
        self._prs = Presentation(self._path)
        return self

    def move_slide(self, from_index: int, to_index: int) -> "Deck":
        """Move a slide (1-based indices). Chainable."""
        from pptx_skill import move_slide
        self._save_if_dirty()
        move_slide(self._path, from_index, to_index)
        from pptx import Presentation
        self._prs = Presentation(self._path)
        return self

    # ------------------------------------------------------------------
    # Metadata
    # ------------------------------------------------------------------

    def set_metadata(self, **kwargs) -> "Deck":
        """Set core document properties (title, author, subject, ...). Chainable."""
        from pptx_skill import set_metadata
        set_metadata(self._prs, **kwargs)
        self._dirty = True
        return self

    def set_custom_property(self, name: str, value) -> "Deck":
        """Set a custom document property. Chainable."""
        from pptx_skill import set_custom_property
        set_custom_property(self._prs, name, value)
        self._dirty = True
        return self

    # ------------------------------------------------------------------
    # Effects
    # ------------------------------------------------------------------

    def add_shadow(
        self,
        slide_index: int,
        shape_name: str | None = None,
        shape_index: int | None = None,
        *,
        blur_pt: float = 8,
        dist_pt: float = 4,
        dir_deg: float = 90,
        color: str = "808080",
        alpha: int = 50,
    ) -> "Deck":
        """Apply a shadow to a shape (1-based slide index). Chainable."""
        from pptx_skill import apply_shadow
        if slide_index < 1 or slide_index > len(self._prs.slides):
            raise IndexError(f"slide_index {slide_index} out of range (1..{len(self._prs.slides)})")
        slide = self._prs.slides[slide_index - 1]
        shape = None
        if shape_name is not None:
            for s in slide.shapes:
                if s.name == shape_name:
                    shape = s
                    break
        elif shape_index is not None:
            shapes = list(slide.shapes)
            if 0 <= shape_index < len(shapes):
                shape = shapes[shape_index]
        if shape is not None:
            apply_shadow(shape, blur_pt=blur_pt, dist_pt=dist_pt,
                         dir_deg=dir_deg, color=color, alpha=alpha)
        self._dirty = True
        return self

    # ------------------------------------------------------------------
    # QA / accessibility
    # ------------------------------------------------------------------

    def audit_accessibility(self):
        """Run a WCAG accessibility audit. Returns an ``AccessibilityReport``."""
        from pptx_skill import audit_accessibility
        return audit_accessibility(self._prs)

    def validate(self) -> dict[str, Any]:
        """Run the structural validation checks. Returns a dict."""
        from pptx_skill import auto_validate_ppt
        # auto_validate_ppt takes a path — save first if dirty.
        if self._dirty and self._path:
            self._prs.save(self._path)
        return auto_validate_ppt(self._path or "")

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def export_pdf(self, output_path: str, *, dpi: int = 150) -> str:
        """Export to PDF. Returns the output path."""
        from pptx_skill import export_to_pdf
        if self._dirty and self._path:
            self._prs.save(self._path)
        return export_to_pdf(self._path or "", output_path, dpi=dpi)

    def export_images(self, output_dir: str, *, dpi: int = 150, fmt: str = "PNG") -> list[str]:
        """Export slides as images. Returns list of image paths."""
        from pptx_skill import export_to_images
        return export_to_images(self._prs, output_dir, dpi=dpi, format=fmt)

    def export_markdown(self, output_path: str | None = None) -> str:
        """Export to Markdown. Returns the markdown text (or writes to path)."""
        from pptx_skill import export_to_markdown
        return export_to_markdown(self._prs, output_path=output_path)

    def export_html(self, output_path: str) -> str:
        """Export to HTML. Returns the output path."""
        from pptx_skill.export import export_to_html
        return export_to_html(self._prs, output_path)

    # ------------------------------------------------------------------
    # Image optimization
    # ------------------------------------------------------------------

    def compress_images(self, *, quality: int = 85, max_dim: int | None = None) -> dict[str, Any]:
        """Compress all images in the deck. Returns stats."""
        from pptx_skill import compress_images
        # compress_images takes a path — save first.
        if self._dirty and self._path:
            self._prs.save(self._path)
        result = compress_images(self._path or "", quality=quality, max_dimension=max_dim)
        # Re-open the compressed file.
        from pptx import Presentation
        self._prs = Presentation(self._path)
        return result

    # ------------------------------------------------------------------
    # Dunder
    # ------------------------------------------------------------------

    def __len__(self) -> int:
        return len(self._prs.slides)

    def __repr__(self) -> str:
        return f"Deck(path={self._path!r}, slides={len(self._prs.slides)})"

    def __enter__(self) -> "Deck":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if exc_type is None and self._dirty:
            self.save()


def open_deck(path: str) -> Deck:
    """Open a PPTX file for editing (convenience function).

    Alias for ``Deck.open(path)``.
    """
    return Deck.open(path)
