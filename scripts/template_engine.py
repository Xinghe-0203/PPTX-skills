"""Template catalog and deterministic template-profile generator for the PPTX skill."""

from __future__ import annotations

import argparse
import colorsys
import copy
import hashlib
import json
import os
import re
import sys
from collections.abc import Iterable
from pathlib import Path

try:
    from .cli_utils import configure_utf8_console
except ImportError:
    from cli_utils import configure_utf8_console

SKILL_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CATALOG = SKILL_ROOT / "assets" / "templates" / "catalog.json"
PACKAGED_GENERATED_TEMPLATE_DIR = SKILL_ROOT / "assets" / "templates" / "generated"
TEMPLATE_DIR_ENV = "PPTX_SKILL_TEMPLATE_DIR"


def get_generated_template_dir() -> Path:
    """Return the writable directory for user-generated template profiles.

    ``PPTX_SKILL_TEMPLATE_DIR`` takes precedence.  Otherwise the path follows
    the platform's conventional per-user data location, keeping runtime writes
    out of the installed package and source checkout.
    """
    override = os.environ.get(TEMPLATE_DIR_ENV, "").strip()
    if override:
        return Path(override).expanduser()
    if os.name == "nt":
        local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
        if local_app_data:
            return Path(local_app_data) / "pptx-skill" / "templates"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "pptx-skill" / "templates"
    xdg_data_home = os.environ.get("XDG_DATA_HOME", "").strip()
    data_home = Path(xdg_data_home).expanduser() if xdg_data_home else Path.home() / ".local" / "share"
    return data_home / "pptx-skill" / "templates"


# Backwards-compatible public constant.  Internal callers use the function so
# an environment override set after import still takes effect.
GENERATED_TEMPLATE_DIR = get_generated_template_dir()


def _generated_template_dirs() -> list[Path]:
    """Return writable then packaged/legacy profile directories, deduplicated."""
    result: list[Path] = []
    seen: set[str] = set()
    for directory in (get_generated_template_dir(), PACKAGED_GENERATED_TEMPLATE_DIR):
        normalized = os.path.normcase(str(directory.expanduser().resolve(strict=False)))
        if normalized not in seen:
            seen.add(normalized)
            result.append(directory)
    return result

VALID_LAYOUTS = {
    "cover", "toc", "section", "bullets", "text_image", "full_image",
    "image_grid", "dashboard", "timeline", "comparison", "quote",
    "process", "table", "end",
    "matrix", "kpi_hero", "faq", "testimonial", "logo_wall",
    "swot", "porter", "pest", "bmc", "funnel",
}

STYLE_KEYWORDS = {
    "strategy-consulting": ("consulting", "strategy", "咨询", "战略", "管理层", "swot", "porter", "pest", "bmc", "funnel", "分析框架", "商业模式", "五力", "漏斗"),
    "executive-dark": ("executive", "board", "董事会", "高管", "高端", "奢华"),
    "product-launch": ("product", "launch", "tech", "产品", "发布", "科技"),
    "data-story": ("data", "analytics", "dashboard", "数据", "经营", "分析"),
    "startup-pitch": ("startup", "pitch", "funding", "创业", "融资", "路演"),
    "academic-clean": ("academic", "research", "thesis", "学术", "科研", "答辩"),
    "training-friendly": ("training", "course", "education", "培训", "课程", "教学"),
    "government-formal": ("government", "policy", "政务", "政府", "政策", "党建"),
    "healthcare-calm": ("health", "medical", "医疗", "健康", "生命科学"),
    "financial-luxe": ("finance", "legal", "金融", "财务", "法律", "投委会"),
    "creative-editorial": ("creative", "editorial", "brand", "创意", "品牌", "杂志"),
    "sustainability": ("sustainability", "environment", "esg", "可持续", "环保", "自然"),
    "tech-startup": ("tech", "startup", "saas", "科技", "创业", "技术架构", "SaaS"),
    "health-wellness": ("wellness", "养生", "心理", "关怀", "健康科普", "养生方案"),
    "cyber-security": ("security", "cyber", "hack", "安全", "渗透", "防护", "黑客"),
    "environmental": ("environmental", "ecology", "green", "生态", "绿色", "环保报告", "绿色能源"),
    "luxury-brand": ("luxury", "premium", "brand", "奢侈", "精品", "高端定制", "尊贵"),
    "education-vibrant": ("education", "vibrant", "课程", "教育", "知识", "培训", "活力"),
    "government-modern": ("modern", "digital", "智慧", "数字化", "政务公开", "智慧城市"),
    "creative-agency": ("agency", "visual", "前卫", "机构", "视觉传达", "大胆"),
}

LAYOUT_FAMILY_KEYWORDS = {
    "editorial_grid": ("editorial", "swiss", "magazine", "grid", "编辑", "瑞士", "杂志", "网格", "出版"),
    "technical_axis": ("technical", "data", "instrument", "engineering", "precise", "技术", "数据", "仪器", "工程", "精密", "严谨"),
    "poster_column": ("poster", "campaign", "launch", "creative", "bold", "海报", "活动", "发布", "创意", "大胆", "品牌"),
}


def _read_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise TypeError(f"JSON file must contain an object: {path}")
    return payload


def _write_json(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return path


def load_catalog(catalog_path: str | Path | None = None) -> dict:
    path = Path(catalog_path) if catalog_path else DEFAULT_CATALOG
    catalog = _read_json(path)
    if catalog.get("version") != 1 or not isinstance(catalog.get("templates"), dict):
        raise ValueError(f"Unsupported template catalog: {path}")
    return catalog


def list_templates(catalog_path: str | Path | None = None) -> list[dict]:
    catalog = load_catalog(catalog_path)
    result = []
    for key, profile in catalog["templates"].items():
        result.append({
            "key": key,
            "name": profile.get("name", key),
            "description": profile.get("description", ""),
            "use_cases": profile.get("use_cases", []),
            "layout_family": profile.get("layout_family", "editorial_grid"),
        })
    generated_profiles: dict[str, dict] = {}
    # Packaged/legacy profiles load first; user profiles with the same id win.
    for directory in reversed(_generated_template_dirs()):
        if not directory.exists():
            continue
        for path in sorted(directory.glob("*.json")):
            try:
                profile = _read_json(path)
            except (json.JSONDecodeError, TypeError, ValueError):
                continue
            key = profile.get("id", path.stem)
            if not isinstance(key, str) or not key:
                key = path.stem
            generated_profiles[key] = {
                "key": key,
                "name": profile.get("name", path.stem),
                "description": profile.get("description", ""),
                "use_cases": profile.get("use_cases", []),
                "layout_family": profile.get("layout_family", "editorial_grid"),
                "generated": True,
            }
    result.extend(generated_profiles[key] for key in sorted(generated_profiles))
    return result


def load_template_profile(
    key_or_path: str | Path,
    catalog_path: str | Path | None = None,
) -> dict:
    candidate = Path(key_or_path)
    if candidate.exists():
        profile = _read_json(candidate)
        profile.setdefault("layout_family", "editorial_grid")
        validate_template_profile(profile)
        return profile

    key = str(key_or_path)
    catalog = load_catalog(catalog_path)
    if key in catalog["templates"]:
        profile = copy.deepcopy(catalog["templates"][key])
        profile.setdefault("id", key)
        profile.setdefault("layout_family", "editorial_grid")
        validate_template_profile(profile)
        return profile

    for directory in _generated_template_dirs():
        generated = directory / f"{key}.json"
        if generated.exists():
            profile = _read_json(generated)
            profile.setdefault("layout_family", "editorial_grid")
            validate_template_profile(profile)
            return profile
    builtin = sorted(catalog["templates"])
    generated_keys = sorted({
        path.stem
        for directory in _generated_template_dirs()
        if directory.exists()
        for path in directory.glob("*.json")
    })
    choices = ", ".join(sorted(set(builtin + generated_keys)))
    raise KeyError(f"Unknown template '{key}'. Available: {choices}")


def validate_template_profile(profile: dict) -> None:
    if not isinstance(profile, dict):
        raise TypeError("Template profile must be a JSON object")
    theme = profile.get("theme")
    if not isinstance(theme, dict):
        raise ValueError("Template profile requires a 'theme' object")
    # Validate all 9 color swatches
    required_colors = ("bg", "bg_alt", "primary", "secondary", "accent", "text", "text_muted", "white", "dark")
    for key in required_colors:
        value = theme.get(key)
        if not isinstance(value, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", value):
            raise ValueError(f"theme.{key} must be a #RRGGBB color")
    # Validate layout_opts values are within reasonable ranges
    layout_opts = profile.get("layout_opts", {})
    if not isinstance(layout_opts, dict):
        raise ValueError("layout_opts must be a dict")
    for layout_name, opts in layout_opts.items():
        if layout_name not in VALID_LAYOUTS:
            raise ValueError(f"Unknown layout in layout_opts: {layout_name}")
        if not isinstance(opts, dict):
            raise ValueError(f"layout_opts.{layout_name} must be a dict")
        for opt_key, opt_val in opts.items():
            if isinstance(opt_val, (int, float)):
                if opt_val < 0:
                    raise ValueError(f"layout_opts.{layout_name}.{opt_key} must be non-negative, got {opt_val}")
                if opt_val > 1000:
                    raise ValueError(f"layout_opts.{layout_name}.{opt_key} seems unreasonably large: {opt_val}")
    # Validate preferred_sequence only contains valid layout names
    preferred_sequence = profile.get("preferred_sequence", [])
    if not isinstance(preferred_sequence, list):
        raise ValueError("preferred_sequence must be a list")
    for layout in preferred_sequence:
        if layout not in VALID_LAYOUTS:
            raise ValueError(f"Unknown layout in preferred_sequence: {layout}")
    layout_family = profile.get("layout_family", "editorial_grid")
    if layout_family not in {"standard", "editorial_grid", "technical_axis", "poster_column"}:
        raise ValueError(f"Unknown layout_family: {layout_family}")
    # Warn on missing optional fields (don't error)
    optional_fields = ("id", "name", "description", "use_cases", "fonts")
    for field in optional_fields:
        if field not in profile:
            import warnings
            warnings.warn(f"Template profile is missing optional field '{field}'", stacklevel=2)


def _slugify(value: str) -> str:
    ascii_slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    if ascii_slug:
        return ascii_slug[:48]
    digest = hashlib.sha256(value.encode("utf-8")).hexdigest()[:10]
    return f"custom-{digest}"


def _select_base_template(prompt: str, catalog: dict) -> str:
    normalized = prompt.casefold()
    scores = {
        key: sum(1 for keyword in keywords if keyword.casefold() in normalized)
        for key, keywords in STYLE_KEYWORDS.items()
    }
    winner = max(scores, key=scores.get)
    return winner if scores[winner] else "strategy-consulting"


def _select_layout_family(prompt: str, base_profile: dict) -> str:
    normalized = prompt.casefold()
    scores = {
        family: sum(1 for keyword in keywords if keyword.casefold() in normalized)
        for family, keywords in LAYOUT_FAMILY_KEYWORDS.items()
    }
    winner = max(scores, key=scores.get)
    if scores[winner]:
        return winner
    return base_profile.get("layout_family", "editorial_grid")


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb: Iterable[int]) -> str:
    return "#" + "".join(f"{max(0, min(255, int(v))):02X}" for v in rgb)


def _hls_color(hue: float, lightness: float, saturation: float) -> str:
    rgb = colorsys.hls_to_rgb(hue % 1.0, lightness, saturation)
    return _rgb_to_hex(round(channel * 255) for channel in rgb)


def _mix(color_a: str, color_b: str, weight_b: float) -> str:
    a = _hex_to_rgb(color_a)
    b = _hex_to_rgb(color_b)
    return _rgb_to_hex(round(x * (1 - weight_b) + y * weight_b) for x, y in zip(a, b, strict=False))


def _generated_palette(prompt: str, dark: bool, brand_colors: list[str] | None) -> dict:
    if brand_colors:
        colors = [color.upper() for color in brand_colors]
        for color in colors:
            if not re.fullmatch(r"#[0-9A-F]{6}", color):
                raise ValueError(f"Invalid brand color: {color}")
        primary = _mix(colors[0], "#FFFFFF", 0.34) if dark else colors[0]
        accent = colors[1] if len(colors) > 1 else _mix(primary, "#C96845", 0.58)
    else:
        digest = hashlib.sha256(prompt.encode("utf-8")).digest()
        hue = digest[0] / 255
        primary = _hls_color(hue, 0.68 if dark else 0.29, 0.42)
        accent = _hls_color(hue + 0.11, 0.62 if dark else 0.51, 0.60)

    if dark:
        bg = _mix(primary, "#080A09", 0.84)
        bg_alt = _mix(primary, "#161A18", 0.68)
        text = "#F2EEE6"
        muted = "#A3AAA5"
        dark_color = "#050706"
    else:
        bg = _mix(primary, "#FBF9F4", 0.94)
        bg_alt = _mix(primary, "#ECE8DF", 0.86)
        text = _mix(primary, "#171A19", 0.78)
        muted = _mix(primary, "#747A77", 0.82)
        dark_color = _mix(primary, "#101413", 0.70)
    return {
        "bg": bg,
        "bg_alt": bg_alt,
        "primary": primary,
        "secondary": muted,
        "accent": accent,
        "text": text,
        "text_muted": muted,
        "white": "#FFFFFF",
        "dark": dark_color,
        "on_dark": dark,
    }


def generate_template_profile(
    name: str,
    prompt: str,
    brand_colors: list[str] | None = None,
    template_id: str | None = None,
) -> dict:
    """Generate a deterministic, reusable template profile from style requirements."""
    catalog = load_catalog()
    base_key = _select_base_template(prompt, catalog)
    profile = copy.deepcopy(catalog["templates"][base_key])
    dark_tokens = ("dark", "night", "black", "深色", "暗色", "黑色", "夜间")
    light_tokens = ("light", "white", "bright", "浅色", "白色", "明亮")
    normalized = prompt.casefold()
    base_dark = bool(profile.get("theme", {}).get("on_dark")) if isinstance(profile.get("theme"), dict) else False
    dark = any(token in normalized for token in dark_tokens)
    if any(token in normalized for token in light_tokens):
        dark = False
    elif not dark:
        dark = base_dark

    profile["id"] = template_id or _slugify(name)
    profile["name"] = name
    profile["description"] = prompt.strip()
    profile["generated_from"] = {"base_template": base_key, "prompt": prompt}
    profile["theme"] = _generated_palette(prompt, dark, brand_colors)
    profile["layout_family"] = _select_layout_family(prompt, profile)
    profile["use_cases"] = profile.get("use_cases", [])
    validate_template_profile(profile)
    return profile


def register_template_profile(
    profile: dict,
    output_dir: str | Path | None = None,
    overwrite: bool = False,
) -> Path:
    validate_template_profile(profile)
    target_dir = Path(output_dir) if output_dir else get_generated_template_dir()
    template_id = profile.get("id") or _slugify(profile.get("name", "custom-template"))
    target = target_dir / f"{template_id}.json"
    if target.exists() and not overwrite:
        raise FileExistsError(f"Template already exists: {target}")
    payload = copy.deepcopy(profile)
    payload["id"] = template_id
    return _write_json(target, payload)


def generate_template_preview(profile: dict, output_path: str | Path) -> str:
    """Create an editable gallery deck that demonstrates a template profile."""
    validate_template_profile(profile)
    try:
        from .pptx_helper import auto_generate_ppt
    except ImportError:
        from pptx_helper import auto_generate_ppt

    sections = [
        {"title": "内容结构", "bullets": ["结论先行", "信息分组", "视觉证据"]},
        {"title": "关键指标", "metrics": [
            {"label": "增长率", "value": "+28%", "change": "同比"},
            {"label": "覆盖率", "value": "76%", "change": "+9pp"},
            {"label": "效率", "value": "2.4×", "change": "提升"},
            {"label": "满意度", "value": "92", "change": "+6"}
        ]},
        {"title": "推进路径", "steps": ["洞察", "设计", "验证", "发布", "复盘"]},
        {"title": "方案选择", "left": {"title": "方案 A", "bullets": ["快速", "轻量", "低风险"]}, "right": {"title": "方案 B", "bullets": ["深入", "灵活", "高上限"]}},
        {"title": "发展节奏", "events": [
            {"date": "Q1", "title": "发现机会"}, {"date": "Q2", "title": "完成验证"},
            {"date": "Q3", "title": "规模推广"}, {"date": "Q4", "title": "持续优化"}
        ]},
        {"title": "让设计服务于结论", "quote": "每一页只承担一个清晰的沟通任务。", "source": "版式原则"}
    ]
    return auto_generate_ppt(
        title=profile.get("name", "自定义模板"),
        subtitle=profile.get("description", "Template Preview"),
        sections=sections,
        output_path=str(output_path),
        auto_search_images=False,
        template_profile=profile,
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Manage reusable PPTX template profiles")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list", help="List built-in and generated template profiles")

    show = sub.add_parser("show", help="Print one template profile")
    show.add_argument("template")

    validate = sub.add_parser("validate", help="Validate a template profile JSON file")
    validate.add_argument("profile")

    preview = sub.add_parser("preview", help="Generate an editable preview deck")
    preview.add_argument("template")
    preview.add_argument("--output", required=True)

    generate = sub.add_parser("generate", help="Generate a reusable profile from style requirements")
    generate.add_argument("--name", required=True)
    generate.add_argument("--prompt", required=True)
    generate.add_argument("--brand-color", action="append", dest="brand_colors")
    generate.add_argument("--output", help="Write profile to this JSON file")
    generate.add_argument("--register", action="store_true")
    generate.add_argument("--preview", help="Also create an editable preview PPTX")
    generate.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    configure_utf8_console()
    args = _build_parser().parse_args()
    if args.command == "list":
        print(json.dumps(list_templates(), ensure_ascii=False, indent=2))
        return 0
    if args.command == "show":
        print(json.dumps(load_template_profile(args.template), ensure_ascii=False, indent=2))
        return 0
    if args.command == "validate":
        profile = _read_json(Path(args.profile))
        validate_template_profile(profile)
        print(f"OK: {args.profile}")
        return 0
    if args.command == "preview":
        generate_template_preview(load_template_profile(args.template), args.output)
        return 0

    profile = generate_template_profile(args.name, args.prompt, args.brand_colors)
    written = None
    if args.output:
        written = _write_json(Path(args.output), profile)
    if args.register:
        written = register_template_profile(profile, overwrite=args.overwrite)
    if args.preview:
        generate_template_preview(profile, args.preview)
    if written:
        print(f"Template profile saved: {written.resolve()}")
    else:
        print(json.dumps(profile, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
