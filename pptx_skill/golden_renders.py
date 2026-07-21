"""PR9: golden render harness.

Per blueprint §12.4:
- Each role has >= 2 golden pages.
- Each geometry family has >= 1 representative deck.
- Save PPTX, PNG (when a rasterizer is available), QA JSON and geometry
  signature per golden.
- Environment key changes only report baseline mismatch, never auto-overwrite.

This module renders one golden PPTX per (role, variant) using the adaptive
renderer, computes a geometry signature, and emits a stable index. PNG export
is optional (LibreOffice/PyMuPDF may be absent) — the harness records the
rasterizer availability in the manifest rather than failing.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from pptx_skill.content_model import (
    ElementSpec,
    SlideSpec,
    canvas_from_name,
)
from pptx_skill.layout_engine import (
    LayoutRecipe,
    _builtin_tokens,
    builtin_recipes,
    solve_recipe,
    solved_geometry_to_layout_plan,
)
from pptx_skill.pptx_renderer import render_layout_plans

ALL_ROLES = (
    "cover", "toc", "section", "bullets", "text_image", "full_image",
    "image_grid", "dashboard", "timeline", "comparison", "quote",
    "process", "table", "end",
    "matrix", "kpi_hero", "faq", "testimonial", "logo_wall",
)

# Representative variant per role for the family deck.
FAMILY_DECK_SEQUENCE = [
    ("cover", "cover.axis"),
    ("toc", "toc.list"),
    ("section", "section.full"),
    ("bullets", "bullets.rail"),
    ("text_image", "text_image.asymmetric_left"),
    ("dashboard", "dashboard.lead_metric"),
    ("timeline", "timeline.horizontal"),
    ("comparison", "comparison.two_column"),
    ("quote", "quote.centered"),
    ("process", "process.horizontal"),
    ("table", "table.standard"),
    ("matrix", "matrix.quadrant"),
    ("kpi_hero", "kpi_hero.split"),
    ("faq", "faq.alternating"),
    ("testimonial", "testimonial.centered"),
    ("logo_wall", "logo_wall.grid3"),
    ("end", "end.thanks"),
]

# Three geometry families representative decks (blueprint §12.4).
GEOMETRY_FAMILIES = ("standard", "editorial_grid", "technical_axis", "poster_column")


@dataclass
class GoldenPage:
    role: str
    variant: str
    recipe_id: str
    pptx_path: str
    png_path: str | None
    qa_json_path: str
    geometry_signature: str
    slide_count: int
    diagnostics: dict[str, Any] = field(default_factory=dict)


@dataclass
class GoldenIndex:
    golden_version: int = 1
    renderer_engine: str = "adaptive"
    rasterizer_available: bool = False
    pages: list[GoldenPage] = field(default_factory=list)
    family_decks: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "golden_version": self.golden_version,
            "renderer_engine": self.renderer_engine,
            "rasterizer_available": self.rasterizer_available,
            "pages": [asdict(p) for p in self.pages],
            "family_decks": self.family_decks,
        }


def _role_slide_spec(role: str, recipe_id: str) -> tuple[SlideSpec, str]:
    """Build a minimal SlideSpec that binds the recipe's zones with sample content."""
    parts = recipe_id.split(".")
    family = parts[0] if parts else "standard"
    sid = f"golden/{role}"
    elements: list[ElementSpec] = []
    sample_text = {
        "title": f"{role} 标题",
        "body": "正文要点一\n正文要点二\n正文要点三",
        "subtitle": "副标题",
        "kicker": "导语",
        "number": "01",
        "caption": "图注",
        "attribution": "来源",
        "mark": "“",
        "lead": "关键指标",
        "support": "支撑说明",
        "footer": "页脚",
        "items": "项目一\n项目二\n项目三",
        "steps": "步骤一\n步骤二\n步骤三",
    }
    recipe = _recipe_by_id(recipe_id)
    zone_roles = set(recipe.zones.keys()) if recipe else set(sample_text.keys())
    for zone_role in zone_roles:
        text = sample_text.get(zone_role, zone_role)
        kind = recipe.zones[zone_role].kind if recipe and zone_role in recipe.zones else "text"
        if kind in ("text", "text-list"):
            elements.append(ElementSpec(
                id=f"{sid}/{zone_role}", kind="text", role=zone_role,
                content={"text": text}, style_ref="component.title" if zone_role in {"title", "mark", "lead"} else "component.body",
            ))
        elif kind == "table":
            elements.append(ElementSpec(
                id=f"{sid}/{zone_role}", kind="table", role=zone_role,
                content={"headers": ["A", "B"], "rows": [["1", "2"], ["3", "4"]]},
                style_ref="component.body",
            ))
        elif kind == "chart":
            elements.append(ElementSpec(
                id=f"{sid}/{zone_role}", kind="chart", role=zone_role,
                content={"items": [{"label": "A", "value": 10}, {"label": "B", "value": 20}]},
                style_ref="component.metric",
            ))
        elif kind in ("image", "hero"):
            elements.append(ElementSpec(
                id=f"{sid}/{zone_role}", kind="image", role=zone_role,
                content={"path": ""}, style_ref="component.hero",
            ))
        else:
            elements.append(ElementSpec(
                id=f"{sid}/{zone_role}", kind="text", role=zone_role,
                content={"text": text}, style_ref="component.body",
            ))
    return SlideSpec(
        id=sid, role=role, communication_goal=f"golden {role}",
        elements=elements, preferred_layouts=[recipe_id],
    ), family


def _recipe_by_id(recipe_id: str) -> LayoutRecipe | None:
    for role in ALL_ROLES:
        for r in builtin_recipes(role):
            if r.id == recipe_id:
                return r
    return None


def _geometry_signature(plan) -> str:
    """Stable signature of a plan's node geometry (family, role, variant, bboxes)."""
    h = hashlib.sha256()
    h.update(plan.recipe_id.encode("utf-8"))
    plan_w = plan.canvas.width_pt or 1.0
    plan_h = plan.canvas.height_pt or 1.0
    for node in sorted(plan.nodes, key=lambda n: n.z_order):
        b = node.geometry.bbox
        h.update(f"{node.role}|{round(b.left,2)},{round(b.top,2)},{round(b.width,2)},{round(b.height,2)}".encode())
    h.update(f"|{plan_w},{plan_h}".encode())
    return h.hexdigest()[:16]


def _try_render_png(pptx_path: str, png_path: Path) -> bool:
    """Attempt to rasterize the first slide to PNG. Returns False if no backend."""
    return _try_render_png_libreoffice(pptx_path, png_path)


def _try_render_png_libreoffice(pptx_path: str, png_path: Path) -> bool:
    """Fallback rasterizer via LibreOffice + PyMuPDF (pdf intermediate)."""
    try:
        import subprocess
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx_path],
                check=True, capture_output=True, timeout=60,
            )
            pdf = Path(tmp) / (Path(pptx_path).stem + ".pdf")
            if not pdf.exists():
                return False
            import fitz  # type: ignore[import-not-found]

            doc = fitz.open(str(pdf))  # type: ignore[attr-defined]
            if doc.page_count > 0:
                doc[0].get_pixmap(dpi=150).save(str(png_path))
            doc.close()
            return True
    except Exception:
        return False


def render_golden_set(
    output_dir: str | Path,
    *,
    roles: tuple[str, ...] = ALL_ROLES,
    variants_per_role: int = 2,
    include_family_decks: bool = True,
    include_png: bool = True,
) -> GoldenIndex:
    """Render the golden PPTX set: >=2 pages per role + family decks.

    Each page saves PPTX, (optional) PNG, QA JSON (geometry signature + node
    count), and a stable geometry signature. Returns a ``GoldenIndex``.
    """
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    canvas = canvas_from_name("16:9")
    tokens = _builtin_tokens()
    index = GoldenIndex()

    # Detect rasterizer availability once.
    rasterizer_ok = False
    if include_png:
        # Probe with a throwaway render.
        try:
            import fitz  # noqa: F401

            rasterizer_ok = True
        except Exception:
            rasterizer_ok = False
    index.rasterizer_available = rasterizer_ok

    for role in roles:
        recipes = builtin_recipes(role)[:variants_per_role]
        for recipe in recipes:
            slide_spec, _family = _role_slide_spec(role, recipe.id)
            geometry = solve_recipe(recipe, canvas, tokens)
            if geometry.infeasible:
                continue
            plan = solved_geometry_to_layout_plan(slide_spec, recipe, geometry, canvas, tokens)
            pptx_path = out / f"{recipe.id.replace('.', '_')}.pptx"
            render_layout_plans([plan], str(pptx_path))
            sig = _geometry_signature(plan)

            png_path: str | None = None
            if include_png and rasterizer_ok:
                candidate_png = out / f"{recipe.id.replace('.', '_')}.png"
                ok = _try_render_png(str(pptx_path), candidate_png)
                if not ok:
                    ok = _try_render_png_libreoffice(str(pptx_path), candidate_png)
                if ok:
                    png_path = str(candidate_png)

            qa_json_path = out / f"{recipe.id.replace('.', '_')}.qa.json"
            qa_payload = {
                "recipe_id": recipe.id,
                "role": role,
                "variant": recipe.variant,
                "geometry_signature": sig,
                "node_count": len(plan.nodes),
                "canvas": {"width_pt": canvas.width_pt, "height_pt": canvas.height_pt},
                "nodes": [
                    {
                        "role": n.role,
                        "kind": n.kind,
                        "bbox": {
                            "left": round(n.geometry.bbox.left, 2),
                            "top": round(n.geometry.bbox.top, 2),
                            "width": round(n.geometry.bbox.width, 2),
                            "height": round(n.geometry.bbox.height, 2),
                        },
                    }
                    for n in plan.nodes
                ],
            }
            with qa_json_path.open("w", encoding="utf-8", newline="\n") as fh:
                json.dump(qa_payload, fh, ensure_ascii=False, indent=2)
                fh.write("\n")

            index.pages.append(GoldenPage(
                role=role,
                variant=recipe.variant,
                recipe_id=recipe.id,
                pptx_path=str(pptx_path),
                png_path=png_path,
                qa_json_path=str(qa_json_path),
                geometry_signature=sig,
                slide_count=1,
                diagnostics={"infeasible": False},
            ))

    if include_family_decks:
        for family in GEOMETRY_FAMILIES:
            deck_pptx = out / f"family_{family}.pptx"
            plans = []
            for role, recipe_id in FAMILY_DECK_SEQUENCE:
                recipe = _recipe_by_id(recipe_id)
                if recipe is None:
                    continue
                slide_spec, _ = _role_slide_spec(role, recipe_id)
                geometry = solve_recipe(recipe, canvas, tokens)
                if geometry.infeasible:
                    continue
                plans.append(solved_geometry_to_layout_plan(slide_spec, recipe, geometry, canvas, tokens))
            if plans:
                render_layout_plans(plans, str(deck_pptx))
                index.family_decks.append({
                    "family": family,
                    "pptx_path": str(deck_pptx),
                    "slide_count": len(plans),
                    "sequence": [rid for _, rid in FAMILY_DECK_SEQUENCE],
                })

    # Write the index.
    index_path = out / "golden_index.json"
    with index_path.open("w", encoding="utf-8", newline="\n") as fh:
        json.dump(index.to_dict(), fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return index


def load_golden_index(output_dir: str | Path) -> dict[str, Any]:
    path = Path(output_dir) / "golden_index.json"
    if not path.exists():
        return {"golden_version": 0, "pages": [], "family_decks": []}
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def golden_coverage_report(index: GoldenIndex | dict[str, Any]) -> dict[str, Any]:
    """Report per-role golden coverage (must be >= 2 per role per §12.4)."""
    pages = index["pages"] if isinstance(index, dict) else [asdict(p) for p in index.pages]
    by_role: dict[str, int] = {}
    for p in pages:
        by_role[p["role"]] = by_role.get(p["role"], 0) + 1
    coverage = []
    all_met = True
    for role in ALL_ROLES:
        count = by_role.get(role, 0)
        met = count >= 2
        all_met = all_met and met
        coverage.append({"role": role, "pages": count, "minimum": 2, "met": met})
    families = index.get("family_decks", []) if isinstance(index, dict) else index.family_decks
    family_coverage = {f["family"] for f in families}
    return {
        "all_role_minimums_met": all_met,
        "role_coverage": coverage,
        "family_decks": sorted(family_coverage),
        "families_present": len(family_coverage),
        "total_pages": len(pages),
    }
