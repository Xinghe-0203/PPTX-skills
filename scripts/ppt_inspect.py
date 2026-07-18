"""Compatibility inspection layer for external PPTX files."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from reference_ppt import analyze_presentation, extract_template_profile


def _guess_shape_role(shape: dict) -> tuple[str, str]:
    placeholder = shape.get("placeholder_type")
    if placeholder in {"TITLE", "CENTER_TITLE"}:
        return "title", "high"
    if placeholder == "SUBTITLE":
        return "subtitle", "high"
    if placeholder in {"BODY", "OBJECT"}:
        return "body", "high"
    text = str(shape.get("text") or "").strip()
    sizes = shape.get("font_sizes") or []
    max_size = max(sizes, default=0)
    if not text:
        return "decoration", "medium"
    if (len(text) <= 3 and max_size >= 50) or (text.isdigit() and max_size >= 36):
        return "decoration", "high"
    if shape.get("y", 1) < 0.3 and max_size >= 26:
        return "title", "medium"
    if shape.get("y", 1) < 0.22 and max_size <= 16:
        return "kicker", "medium"
    return "body", "low"


def inspect_ppt(pptx_path: str | Path) -> dict:
    result = analyze_presentation(pptx_path)
    profile = extract_template_profile(pptx_path)
    for slide in result["slides"]:
        slide["guessed_layout"] = slide["role"]
        slide["layout_confidence"] = "high" if slide["role"] in {
            "cover", "end", "dashboard", "comparison", "image_grid",
        } else "medium"
        for shape in slide["shapes"]:
            role, confidence = _guess_shape_role(shape)
            shape["guessed_role"] = role
            shape["role_confidence"] = confidence
    result["width_in"] = result["slide_size"]["width_in"]
    result["height_in"] = result["slide_size"]["height_in"]
    result["guessed_theme"] = {
        "name": profile["name"],
        "confidence": "medium" if result["design_tokens"]["colors"] else "low",
        "primary": profile["theme"]["primary"],
        "bg": profile["theme"]["bg"],
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect an external PPTX")
    parser.add_argument("pptx")
    parser.add_argument("--output")
    args = parser.parse_args()
    result = inspect_ppt(args.pptx)
    if args.output:
        target = Path(args.output)
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(result, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        print(str(target.resolve()))
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
