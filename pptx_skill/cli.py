"""Unified command-line interface for pptx-skill.

Provides a single ``pptx-skill`` command with subcommands for the most
common operations: generate, inspect, edit, render, export, batch, merge,
template, pages, and info.

Usage
-----
::

    pptx-skill info deck.pptx
    pptx-skill generate --title "My Deck" --sections sections.json -o out.pptx
    pptx-skill from-markdown outline.md -o out.pptx
    pptx-skill inspect deck.pptx
    pptx-skill render deck.pptx -o ./previews --dpi 150
    pptx-skill edit deck.pptx --find-replace "Old:New"
    pptx-skill watermark deck.pptx --text "DRAFT" --opacity 0.15
    pptx-skill export deck.pptx --pdf out.pdf
    pptx-skill pages deck.pptx --delete 3
    pptx-skill merge a.pptx b.pptx -o merged.pptx
    pptx-skill template list
    pptx-skill validate deck.pptx

Also supports ``python -m pptx_skill <command>``.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections.abc import Sequence


def _ensure_utf8() -> None:
    """Ensure UTF-8 output on Windows to avoid GBK encoding errors."""
    if sys.platform == "win32":
        os.environ.setdefault("PYTHONIOENCODING", "utf-8")


# ---------------------------------------------------------------------------
# Command: info
# ---------------------------------------------------------------------------

def cmd_info(args: argparse.Namespace) -> int:
    """Print summary info about a PPTX file."""
    from pptx_skill import Deck
    deck = Deck.open(args.path)
    info = deck.info()
    print(f"Path: {info['path']}")
    print(f"Slides: {info['slide_count']}")
    print(f"Size: {info['width_in']:.2f} x {info['height_in']:.2f} in")
    print()
    print(f"{'#':>3}  {'Shapes':>7}  Notes")
    print(f"{'─'*3}  {'─'*7}  {'─'*5}")
    for s in info["slides"]:
        notes = "yes" if s["has_notes"] else ""
        print(f"{s['index']:>3}  {s['shape_count']:>7}  {notes}")
    return 0


# ---------------------------------------------------------------------------
# Command: inspect
# ---------------------------------------------------------------------------

def cmd_inspect(args: argparse.Namespace) -> int:
    """Print a structured inspection of a PPTX file."""
    from pptx_skill import inspect_ppt
    result = inspect_ppt(args.path)
    print(json.dumps(result, indent=2, ensure_ascii=False, default=str))
    return 0


# ---------------------------------------------------------------------------
# Command: generate
# ---------------------------------------------------------------------------

def cmd_generate(args: argparse.Namespace) -> int:
    """Generate a PPTX from a JSON sections file."""
    from pptx_skill import auto_generate_ppt

    if args.sections:
        with open(args.sections, encoding="utf-8") as f:
            sections = json.load(f)
    else:
        sections = []

    path = auto_generate_ppt(
        title=args.title or "Untitled",
        subtitle=args.subtitle or "",
        sections=sections,
        output_path=args.output,
        theme_key=args.theme,
        lang=args.lang,
        auto_search_images=args.search_images,
    )
    print(f"Generated: {path}")
    return 0


# ---------------------------------------------------------------------------
# Command: from-markdown
# ---------------------------------------------------------------------------

def cmd_from_markdown(args: argparse.Namespace) -> int:
    """Generate a PPTX from a Markdown file."""
    from pptx_skill import import_markdown
    path = import_markdown(
        args.markdown_path,
        output_path=args.output,
        title=args.title,
        lang=args.lang,
        theme_key=args.theme,
        auto_search_images=args.search_images,
    )
    print(f"Generated: {path}")
    return 0


# ---------------------------------------------------------------------------
# Command: render
# ---------------------------------------------------------------------------

def cmd_render(args: argparse.Namespace) -> int:
    """Render slides as PNG images."""
    from pptx_skill import export_to_images
    images = export_to_images(args.path, args.output, dpi=args.dpi, format=args.format)
    print(f"Rendered {len(images)} slides to {args.output}")
    for img in images:
        print(f"  {img}")
    return 0


# ---------------------------------------------------------------------------
# Command: edit
# ---------------------------------------------------------------------------

def cmd_edit(args: argparse.Namespace) -> int:
    """Edit text in a PPTX file (find-replace, per-slide or global)."""
    from pptx_skill import Deck
    deck = Deck.open(args.path)
    total = 0
    if args.find_replace:
        for pair in args.find_replace:
            if ":" not in pair:
                print(f"Error: --find-replace expects 'Old:New' format, got {pair!r}")
                return 1
            old, new = pair.split(":", 1)
            count = deck.find_replace_all(old, new)
            total += count
            print(f"  '{old}' -> '{new}': {count} replacements")
    if args.title:
        # Set slide 1 title
        deck.edit_title(1, args.title)
        print(f"  Set slide 1 title to: {args.title}")
    deck.save()
    print(f"Total replacements: {total}")
    print(f"Saved: {args.path}")
    return 0


# ---------------------------------------------------------------------------
# Command: watermark
# ---------------------------------------------------------------------------

def cmd_watermark(args: argparse.Namespace) -> int:
    """Add a watermark to a PPTX file."""
    from pptx_skill import Deck
    deck = Deck.open(args.path)
    if args.text:
        deck.add_watermark(text=args.text, opacity=args.opacity,
                           diagonal=args.diagonal, font_size=args.font_size)
        print(f"Added text watermark: {args.text!r}")
    elif args.image:
        deck.add_watermark(image_path=args.image, opacity=args.opacity)
        print(f"Added image watermark: {args.image}")
    else:
        print("Error: --text or --image required")
        return 1
    deck.save()
    print(f"Saved: {args.path}")
    return 0


# ---------------------------------------------------------------------------
# Command: export
# ---------------------------------------------------------------------------

def cmd_export(args: argparse.Namespace) -> int:
    """Export a PPTX to PDF, images, HTML, or Markdown."""
    from pptx_skill import Deck
    deck = Deck.open(args.path)
    if args.pdf:
        deck.export_pdf(args.pdf, dpi=args.dpi)
        print(f"Exported PDF: {args.pdf}")
    if args.images:
        imgs = deck.export_images(args.images, dpi=args.dpi)
        print(f"Exported {len(imgs)} images to {args.images}")
    if args.html:
        deck.export_html(args.html)
        print(f"Exported HTML: {args.html}")
    if args.markdown:
        md = deck.export_markdown(args.markdown)
        print(f"Exported Markdown: {args.markdown} ({len(md)} chars)")
    return 0


# ---------------------------------------------------------------------------
# Command: pages
# ---------------------------------------------------------------------------

def cmd_pages(args: argparse.Namespace) -> int:
    """Page-level operations: delete, move."""
    from pptx_skill import Deck
    deck = Deck.open(args.path)
    if args.delete is not None:
        deck.delete_slide(args.delete)
        print(f"Deleted slide {args.delete}")
    if args.move:
        if len(args.move) != 2:
            print("Error: --move expects FROM TO (1-based)")
            return 1
        deck.move_slide(args.move[0], args.move[1])
        print(f"Moved slide {args.move[0]} -> {args.move[1]}")
    deck.save()
    print(f"Saved: {args.path} ({deck.slide_count} slides)")
    return 0


# ---------------------------------------------------------------------------
# Command: merge
# ---------------------------------------------------------------------------

def cmd_merge(args: argparse.Namespace) -> int:
    """Merge multiple PPTX files into one."""
    from pptx_skill import merge_presentations
    path = merge_presentations(args.paths, output_path=args.output)
    print(f"Merged {len(args.paths)} files -> {path}")
    return 0


# ---------------------------------------------------------------------------
# Command: validate
# ---------------------------------------------------------------------------

def cmd_validate(args: argparse.Namespace) -> int:
    """Run structural validation checks on a PPTX file."""
    from pptx_skill import auto_validate_ppt
    result = auto_validate_ppt(args.path)
    assert isinstance(result, dict), "Expected dict from auto_validate_ppt"
    passed = result.get("passed", False)
    print(f"Validation: {'PASSED' if passed else 'FAILED'}")
    for check_name, check_passed in result.get("checks", []):
        status = "OK" if check_passed else "FAIL"
        print(f"  [{status}] {check_name}")
    for w in result.get("warnings", []):
        print(f"  [WARN] {w}")
    return 0 if passed else 1


# ---------------------------------------------------------------------------
# Command: template
# ---------------------------------------------------------------------------

def cmd_template(args: argparse.Namespace) -> int:
    """List available templates."""
    from pptx_skill import list_templates
    templates = list_templates()
    if not templates:
        print("No templates found.")
        return 0
    print(f"{'Name':<30} {'Family':<20} {'Description'}")
    print(f"{'─'*30} {'─'*20} {'─'*40}")
    for t in templates:
        name = t.get("name", "?")
        family = t.get("layout_family", "")
        desc = t.get("description", "")[:60]
        print(f"{name:<30} {family:<20} {desc}")
    print(f"\n{len(templates)} templates available.")
    return 0


# ---------------------------------------------------------------------------
# Command: capability
# ---------------------------------------------------------------------------

def cmd_capability(args: argparse.Namespace) -> int:
    """Print environment capability report."""
    from pptx_skill.capability import main as cap_main
    return cap_main()


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    """Build the top-level argument parser."""
    parser = argparse.ArgumentParser(
        prog="pptx-skill",
        description="The world's most powerful PPTX toolkit — generate, edit, inspect, export.",
    )
    sub = parser.add_subparsers(dest="command", help="Available commands")

    # info
    p = sub.add_parser("info", help="Show summary info about a PPTX file")
    p.add_argument("path", help="Path to .pptx file")
    p.set_defaults(func=cmd_info)

    # inspect
    p = sub.add_parser("inspect", help="Structured inspection of a PPTX file")
    p.add_argument("path", help="Path to .pptx file")
    p.set_defaults(func=cmd_inspect)

    # generate
    p = sub.add_parser("generate", help="Generate a PPTX from a JSON sections file")
    p.add_argument("--title", help="Deck title")
    p.add_argument("--subtitle", help="Deck subtitle")
    p.add_argument("--sections", help="Path to JSON file with sections array")
    p.add_argument("-o", "--output", default="output.pptx", help="Output path")
    p.add_argument("--theme", help="Theme key")
    p.add_argument("--lang", default="zh", help="Language (default zh)")
    p.add_argument("--search-images", action="store_true", help="Auto-search images")
    p.set_defaults(func=cmd_generate)

    # from-markdown
    p = sub.add_parser("from-markdown", help="Generate a PPTX from a Markdown file")
    p.add_argument("markdown_path", help="Path to .md file")
    p.add_argument("-o", "--output", default="output.pptx", help="Output path")
    p.add_argument("--title", help="Override deck title")
    p.add_argument("--theme", help="Theme key")
    p.add_argument("--lang", default="zh", help="Language")
    p.add_argument("--search-images", action="store_true", help="Auto-search images")
    p.set_defaults(func=cmd_from_markdown)

    # render
    p = sub.add_parser("render", help="Render slides as PNG images")
    p.add_argument("path", help="Path to .pptx file")
    p.add_argument("-o", "--output", default="./preview", help="Output directory")
    p.add_argument("--dpi", type=int, default=150, help="DPI (default 150)")
    p.add_argument("--format", default="PNG", choices=["PNG", "JPEG", "BMP"])
    p.set_defaults(func=cmd_render)

    # edit
    p = sub.add_parser("edit", help="Edit text in a PPTX (find-replace)")
    p.add_argument("path", help="Path to .pptx file")
    p.add_argument("--find-replace", action="append", help="Old:New pair (can repeat)")
    p.add_argument("--title", help="Set slide 1 title")
    p.set_defaults(func=cmd_edit)

    # watermark
    p = sub.add_parser("watermark", help="Add a watermark to a PPTX")
    p.add_argument("path", help="Path to .pptx file")
    p.add_argument("--text", help="Watermark text")
    p.add_argument("--image", help="Watermark image path")
    p.add_argument("--opacity", type=float, default=0.15, help="Opacity 0-1")
    p.add_argument("--diagonal", action="store_true", help="Diagonal text")
    p.add_argument("--font-size", type=int, default=48, help="Font size")
    p.set_defaults(func=cmd_watermark)

    # export
    p = sub.add_parser("export", help="Export PPTX to PDF/images/HTML/Markdown")
    p.add_argument("path", help="Path to .pptx file")
    p.add_argument("--pdf", help="Output PDF path")
    p.add_argument("--images", help="Output images directory")
    p.add_argument("--html", help="Output HTML path")
    p.add_argument("--markdown", help="Output Markdown path")
    p.add_argument("--dpi", type=int, default=150, help="DPI for images")
    p.set_defaults(func=cmd_export)

    # pages
    p = sub.add_parser("pages", help="Page operations (delete/move)")
    p.add_argument("path", help="Path to .pptx file")
    p.add_argument("--delete", type=int, help="Delete slide at 1-based index")
    p.add_argument("--move", nargs=2, type=int, metavar=("FROM", "TO"),
                   help="Move slide FROM -> TO (1-based)")
    p.set_defaults(func=cmd_pages)

    # merge
    p = sub.add_parser("merge", help="Merge multiple PPTX files")
    p.add_argument("paths", nargs="+", help="Paths to .pptx files")
    p.add_argument("-o", "--output", default="merged.pptx", help="Output path")
    p.set_defaults(func=cmd_merge)

    # validate
    p = sub.add_parser("validate", help="Run structural validation")
    p.add_argument("path", help="Path to .pptx file")
    p.set_defaults(func=cmd_validate)

    # template
    p = sub.add_parser("template", help="List available templates")
    p.set_defaults(func=cmd_template)

    # capability
    p = sub.add_parser("capability", help="Show environment capability report")
    p.set_defaults(func=cmd_capability)

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Entry point for the ``pptx-skill`` console script."""
    _ensure_utf8()
    parser = build_parser()
    args = parser.parse_args(argv)

    if not hasattr(args, "func"):
        parser.print_help()
        return 0

    try:
        return args.func(args)
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        return 130
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if os.environ.get("PPTX_SKILL_DEBUG"):
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
