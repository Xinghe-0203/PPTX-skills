"""Download, import and search for PPTX template packs.

Provides utilities to fetch .pptx files from GitHub repos, direct URLs or
local directories, then import them as reusable template profiles that can
be consumed by the V2 template compiler and renderer.

Only stdlib networking (urllib, json, pathlib) is used -- no ``requests``
dependency.
"""
from __future__ import annotations

import ipaddress
import json
import logging
import re
import shutil
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

_SKILL_ROOT = Path(__file__).resolve().parent.parent
_SCRIPTS_DIR = _SKILL_ROOT / "scripts"
_DEFAULT_CATALOG_DIR = _SKILL_ROOT / "assets" / "templates" / "generated"

_GITHUB_API_BASE = "https://api.github.com"
_GITHUB_RAW_BASE = "https://raw.githubusercontent.com"
_REQUEST_TIMEOUT = 30  # seconds


# ---------------------------------------------------------------------------
# Curated remote packs
# ---------------------------------------------------------------------------

REMOTE_PACKS: list[dict[str, Any]] = [
    {
        "id": "slidesgo-free",
        "name": "SlidesGo Free Templates",
        "url": "https://slidesgo.com",
        "type": "web",
        "license": "Custom (attribution required)",
        "note": "Manual download; programmatic access not available",
    },
    {
        "id": "github-pptx-templates",
        "name": "GitHub PPTX Template Collections",
        "url": "https://github.com/search?q=pptx+template",
        "type": "github_search",
        "license": "Varies per repo",
        "note": "Search GitHub for MIT/BSD licensed .pptx template repos",
    },
    {
        "id": "slidescarnival",
        "name": "SlidesCarnival Free Templates",
        "url": "https://www.slidescarnival.com",
        "type": "web",
        "license": "Free for personal & commercial use (attribution appreciated)",
        "note": "Manual download; high-quality free templates",
    },
    {
        "id": "fppt-free",
        "name": "Free PowerPoint Templates (FPPT)",
        "url": "https://www.free-power-point-templates.com",
        "type": "web",
        "license": "Free for personal use",
        "note": "Manual download; large collection of .pptx files",
    },
    {
        "id": "googleslides-templates",
        "name": "Google Slides Template Gallery",
        "url": "https://slides.google.com",
        "type": "web",
        "license": "Free",
        "note": "Web-based; export to .pptx manually",
    },
]


# ---------------------------------------------------------------------------
# Security helpers
# ---------------------------------------------------------------------------

# Private-network CIDRs that must never be reached from user-supplied URLs.
_PRIVATE_NETWORKS = [
    ipaddress.ip_network("127.0.0.0/8"),
    ipaddress.ip_network("10.0.0.0/8"),
    ipaddress.ip_network("172.16.0.0/12"),
    ipaddress.ip_network("192.168.0.0/16"),
]


def _validate_url(url: str) -> None:
    """Reject dangerous URLs: non-HTTP(S) schemes and private/reserved IPs."""
    parsed = urllib.parse.urlparse(url)
    scheme = parsed.scheme.lower()
    if scheme not in ("http", "https"):
        raise ValueError(
            f"Blocked URL scheme {scheme!r}; only http and https are allowed"
        )
    if scheme == "file":
        # Redundant (already blocked above), but explicit for clarity.
        raise ValueError("file:// scheme is not allowed for downloads")

    hostname = parsed.hostname
    if not hostname:
        return  # relative / malformed; let urllib surface the real error later

    try:
        # Resolve the hostname to check for private IPs.
        import socket
        addr_infos = socket.getaddrinfo(hostname, None, socket.AF_UNSPEC, socket.SOCK_STREAM)
    except socket.gaierror:
        return  # unresolvable hostname; let urllib surface the real error later

    for _family, _type, _proto, _canon, sockaddr in addr_infos:
        ip_str = sockaddr[0]
        try:
            ip = ipaddress.ip_address(ip_str)
        except ValueError:
            continue
        for network in _PRIVATE_NETWORKS:
            if ip in network:
                raise ValueError(
                    f"Blocked private IP {ip_str} for hostname {hostname!r}"
                )


def _sanitize_filename(name: str) -> str:
    """Strip path separators and traversal sequences from a filename."""
    # Keep only the last path component (defence against embedded / or \).
    name = name.replace("\\", "/").rsplit("/", 1)[-1]
    # Reject traversal patterns.
    if ".." in name:
        raise ValueError(f"Filename contains path-traversal sequence: {name!r}")
    # Remove leading dots (hidden files / relative tricks).
    name = name.lstrip(".")
    if not name:
        raise ValueError("Filename is empty after sanitization")
    return name


def _validate_output_dir(output_dir: str) -> Path:
    """Validate *output_dir* does not contain path-traversal sequences."""
    if ".." in Path(output_dir).parts:
        raise ValueError(
            f"output_dir contains path-traversal sequence '..': {output_dir!r}"
        )
    p = Path(output_dir).resolve()
    return p


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _ensure_scripts_on_path() -> None:
    """Make sure ``scripts/`` is importable for ``reference_ppt`` etc."""
    import sys
    if str(_SCRIPTS_DIR) not in sys.path:
        sys.path.insert(0, str(_SCRIPTS_DIR))


def _fetch_json(url: str, timeout: int = _REQUEST_TIMEOUT) -> Any:
    """Fetch JSON from *url*, raising on HTTP errors."""
    _validate_url(url)
    req = urllib.request.Request(url, headers={"User-Agent": "pptx-skill/2.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
    return json.loads(raw.decode("utf-8", errors="replace"))


def _download_file(url: str, dest: Path, timeout: int = _REQUEST_TIMEOUT) -> None:
    """Download a single file from *url* to *dest*."""
    _validate_url(url)
    req = urllib.request.Request(url, headers={"User-Agent": "pptx-skill/2.0"})
    dest.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(req, timeout=timeout) as resp, \
         dest.open("wb") as out:
        while True:
            chunk = resp.read(64 * 1024)
            if not chunk:
                break
            out.write(chunk)


_REPO_RE = re.compile(r'^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$')
_BRANCH_RE = re.compile(r'^[A-Za-z0-9_.:/-]+$')


def _validate_repo_branch(repo: str, branch: str) -> None:
    """Validate that *repo* and *branch* match expected formats."""
    if not _REPO_RE.match(repo):
        raise ValueError(
            f"Invalid repo format {repo!r}; expected 'owner/repo' with "
            f"alphanumeric characters, dots, hyphens and underscores"
        )
    if not _BRANCH_RE.match(branch):
        raise ValueError(
            f"Invalid branch format {branch!r}; expected alphanumeric "
            f"characters, dots, hyphens, underscores, colons and slashes"
        )


def _github_api_list_dir(repo: str, subdir: str, branch: str) -> list[dict[str, Any]]:
    """List files in a GitHub repo directory via the Contents API."""
    _validate_repo_branch(repo, branch)
    encoded_subdir = urllib.parse.quote(subdir.strip("/"), safe="/")
    url = f"{_GITHUB_API_BASE}/repos/{repo}/contents/{encoded_subdir}?ref={branch}"
    result = _fetch_json(url)
    # GitHub API returns a single dict when the path points to a file;
    # wrap it in a list so callers can iterate uniformly.
    if isinstance(result, dict):
        return [result]
    return result


def _github_raw_url(repo: str, path: str, branch: str) -> str:
    """Build a raw.githubusercontent.com download URL."""
    encoded_path = urllib.parse.quote(path.lstrip("/"), safe="/")
    return f"{_GITHUB_RAW_BASE}/{repo}/{branch}/{encoded_path}"


def _walk_github_dir(repo: str, subdir: str, branch: str, depth: int = 3) -> list[str]:
    """Recursively find .pptx files in a GitHub repo directory.

    Returns a list of repo-relative paths.
    """
    if depth <= 0:
        return []
    try:
        entries = _github_api_list_dir(repo, subdir, branch)
    except (urllib.error.HTTPError, urllib.error.URLError, json.JSONDecodeError) as exc:
        logger.warning("Failed to list GitHub dir %s/%s: %s", repo, subdir, exc)
        return []

    pptx_paths: list[str] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        entry_type = entry.get("type", "")
        entry_path = entry.get("path", "")
        if entry_type == "file" and entry_path.lower().endswith(".pptx"):
            pptx_paths.append(entry_path)
        elif entry_type == "dir" and depth > 1:
            pptx_paths.extend(_walk_github_dir(repo, entry_path, branch, depth - 1))
    return pptx_paths


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def download_template_pack(source: str, output_dir: str, **kwargs: Any) -> dict[str, Any]:
    """Download a template pack from a known source.

    Parameters
    ----------
    source :
        One of ``"github"``, ``"url"``, or ``"local"``.
    output_dir :
        Directory to save downloaded files into.
    kwargs :
        Source-specific options (see below).

    GitHub kwargs: ``repo`` (e.g. "user/repo"), ``subdir`` (optional),
    ``branch`` (default "main").
    URL kwargs: ``url`` (direct .pptx link).
    Local kwargs: ``path`` (local directory to scan).

    Returns
    -------
    dict with ``downloaded`` (list of file paths), ``errors`` (list of error
    messages), and ``source`` (the *source* argument).
    """
    out = _validate_output_dir(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    result: dict[str, Any] = {"downloaded": [], "errors": [], "source": source}

    if source == "github":
        repo = kwargs.get("repo", "")
        if not repo or "/" not in repo:
            result["errors"].append("github source requires 'repo' kwarg (e.g. 'user/repo')")
            return result
        subdir = kwargs.get("subdir", "")
        branch = kwargs.get("branch", "main")
        try:
            _validate_repo_branch(repo, branch)
        except ValueError as exc:
            result["errors"].append(str(exc))
            return result
        try:
            pptx_paths = _walk_github_dir(repo, subdir, branch)
        except (urllib.error.HTTPError, urllib.error.URLError) as exc:
            reason = getattr(exc, "reason", str(exc))
            code = getattr(exc, "code", None)
            if code == 403:
                result["errors"].append(f"GitHub API rate limit hit: {reason}")
            elif code == 404:
                result["errors"].append(f"GitHub repo/dir not found: {repo}/{subdir}")
            else:
                result["errors"].append(f"GitHub API error: {exc}")
            return result

        if not pptx_paths:
            result["errors"].append(f"No .pptx files found in {repo}/{subdir or '.'}")
            return result

        for rel_path in pptx_paths:
            filename = _sanitize_filename(Path(rel_path).name)
            dest = out / filename
            try:
                raw_url = _github_raw_url(repo, rel_path, branch)
                _download_file(raw_url, dest)
                result["downloaded"].append(str(dest.resolve()))
                logger.info("Downloaded %s -> %s", rel_path, dest)
            except (urllib.error.HTTPError, urllib.error.URLError, OSError) as exc:
                msg = f"Failed to download {rel_path}: {exc}"
                result["errors"].append(msg)
                logger.warning(msg)

    elif source == "url":
        url = kwargs.get("url", "")
        if not url:
            result["errors"].append("url source requires 'url' kwarg")
            return result
        # Derive filename from URL path, fallback to "downloaded.pptx"
        parsed = urllib.parse.urlparse(url)
        raw_name = Path(parsed.path).name or "downloaded.pptx"
        filename = _sanitize_filename(raw_name)
        if not filename.lower().endswith(".pptx"):
            filename += ".pptx"
        dest = out / filename
        try:
            _download_file(url, dest)
            result["downloaded"].append(str(dest.resolve()))
            logger.info("Downloaded %s -> %s", url, dest)
        except (urllib.error.HTTPError, urllib.error.URLError, OSError) as exc:
            result["errors"].append(f"Failed to download {url}: {exc}")
            logger.warning("Failed to download %s: %s", url, exc)

    elif source == "local":
        local_path = kwargs.get("path", "")
        if not local_path:
            result["errors"].append("local source requires 'path' kwarg")
            return result
        src_dir = Path(local_path)
        if not src_dir.is_dir():
            result["errors"].append(f"Path is not a directory: {local_path}")
            return result
        for pptx_file in src_dir.rglob("*.pptx"):
            dest = out / pptx_file.name
            try:
                shutil.copy2(str(pptx_file), str(dest))
                result["downloaded"].append(str(dest.resolve()))
                logger.info("Copied %s -> %s", pptx_file, dest)
            except OSError as exc:
                msg = f"Failed to copy {pptx_file}: {exc}"
                result["errors"].append(msg)
                logger.warning(msg)

    else:
        result["errors"].append(f"Unknown source: {source!r}. Use 'github', 'url', or 'local'.")

    return result


def import_template(
    pptx_path: str,
    name: str | None = None,
    catalog_dir: str | None = None,
) -> dict[str, Any]:
    """Import a .pptx file as a template profile.

    1. Extracts colors, fonts and layout info via
       ``reference_ppt.extract_template_profile``.
    2. Runs ``reference_ppt.analyze_presentation`` for slide roles and design
       tokens.
    3. Converts the result to a TemplateProfileV2-compatible format with
       design_schema tokens.
    4. Saves the profile JSON to *catalog_dir* (or the default generated
       templates directory).
    5. Returns the profile dict.
    """
    _ensure_scripts_on_path()
    import reference_ppt  # noqa: E402

    pptx = Path(pptx_path).resolve()
    if not pptx.exists():
        raise FileNotFoundError(f"PPTX file not found: {pptx}")
    if not pptx.is_file():
        raise ValueError(f"PPTX path is not a regular file: {pptx}")
    if not pptx.suffix.lower() == ".pptx":
        raise ValueError(f"Expected a .pptx file, got: {pptx.suffix}")
    _MAX_IMPORT_SIZE = 50 * 1024 * 1024  # 50 MB
    if pptx.stat().st_size > _MAX_IMPORT_SIZE:
        raise ValueError(
            f"PPTX file too large ({pptx.stat().st_size / (1024 * 1024):.1f} MB); "
            f"maximum is {_MAX_IMPORT_SIZE // (1024 * 1024)} MB"
        )

    profile_name = name or pptx.stem

    # Step 1: extract the V1-style profile (colors, fonts, reference info)
    v1_profile = reference_ppt.extract_template_profile(str(pptx), profile_name)

    # Step 2: get the full analysis for slide roles and design tokens
    analysis = reference_ppt.analyze_presentation(str(pptx))

    # Step 3: build a V2-compatible profile with design_schema tokens
    theme = v1_profile.get("theme", {})
    fonts = v1_profile.get("fonts", {})
    on_dark = theme.get("on_dark", False)

    # Build three-layer tokens (primitive -> semantic -> component)
    primitive_tokens = {
        "palette": {
            "ink_950": theme.get("dark", "#1A1A1A"),
            "paper_050": theme.get("bg", "#FDFCF8"),
            "accent": theme.get("accent", "#C96845"),
            "muted": theme.get("text_muted", "#6E7C87"),
            "primary": theme.get("primary", "#2D5F8A"),
            "secondary": theme.get("secondary", "#5A8DB8"),
        },
        "font": {
            "family": {
                "display": fonts.get("primary", "Microsoft YaHei"),
                "body": fonts.get("primary", "Microsoft YaHei"),
            },
        },
    }

    semantic_tokens = {
        "color": {
            "background": {"$ref": "primitive.palette.paper_050"},
            "background_alt": {"$ref": "primitive.palette.muted"},
            "foreground": {"$ref": "primitive.palette.ink_950"},
            "accent": {"$ref": "primitive.palette.accent"},
            "muted": {"$ref": "primitive.palette.muted"},
        },
        "typography": {
            "display": {"$ref": "primitive.font.family.display"},
            "body": {"$ref": "primitive.font.family.body"},
        },
    }

    component_tokens = {
        "slide": {
            "background": {"$ref": "semantic.color.background"},
            "title_color": {"$ref": "semantic.color.foreground"},
            "body_color": {"$ref": "semantic.color.foreground"},
        },
        "cover": {
            "background": {"$ref": "semantic.color.accent"},
            "title_color": {"$ref": "semantic.color.background"},
        },
    }

    # Determine layout family from analysis
    layout_family = v1_profile.get("layout_family", "editorial_grid")
    if layout_family not in {"editorial_grid", "technical_axis", "poster_column"}:
        layout_family = "editorial_grid"

    # Build slide role -> layout mapping from analysis
    slide_roles = v1_profile.get("reference", {}).get("slide_roles", {})
    role_set: set[str] = set()
    for role in slide_roles.values():
        role_set.add(role)
    # Ensure at least the common roles are covered
    for fallback_role in ("cover", "bullets", "end"):
        role_set.add(fallback_role)

    layouts_map: dict[str, list[str]] = {}
    for role in sorted(role_set):
        layouts_map[role] = [f"{layout_family}-{role}-default"]

    # Build the V2 profile
    profile_id = v1_profile.get("id", re.sub(r"[^a-z0-9]+", "-", profile_name.lower()).strip("-") or "imported-template")

    v2_profile: dict[str, Any] = {
        "schema_version": 2,
        "kind": "template-source-profile",
        "id": profile_id,
        "name": profile_name,
        "description": f"Imported from {pptx.name}",
        "canvas": {
            "preset": "16:9",
            "width_pt": round(analysis.get("slide_size", {}).get("width_in", 13.333) * 72, 3),
            "height_pt": round(analysis.get("slide_size", {}).get("height_in", 7.5) * 72, 3),
            "safe_margin": {"top": 36, "right": 48, "bottom": 32, "left": 48},
        },
        "tokens": {
            "primitive": primitive_tokens,
            "semantic": semantic_tokens,
            "component": component_tokens,
        },
        "grammar": {
            "family": layout_family,
        },
        "layouts": layouts_map,
        "coverage": {
            "status": "partial",
            "roles": sorted(role_set),
        },
        "rhythm": {"rules": []},
        "qa": {},
        "provenance": {
            "source_file": str(pptx),
            "recommended_mode": analysis.get("recommended_mode", "clone"),
            "imported_at": _iso_now(),
        },
        "extensions": {
            "legacy_v1": v1_profile,
            "design_tokens": analysis.get("design_tokens", {}),
        },
    }

    # Step 4: save to catalog
    target_dir = Path(catalog_dir) if catalog_dir else _DEFAULT_CATALOG_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    target_file = target_dir / f"{profile_id}.json"
    with target_file.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(v2_profile, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    logger.info("Template profile saved: %s", target_file.resolve())

    # Step 5: return the profile
    v2_profile["_saved_to"] = str(target_file.resolve())
    return v2_profile


def list_remote_packs() -> list[dict[str, Any]]:
    """Return a curated list of known open-source PPTX template packs."""
    return [dict(pack) for pack in REMOTE_PACKS]


def search_github_templates(query: str, max_results: int = 10) -> list[dict[str, Any]]:
    """Search GitHub for PPTX template repositories.

    Uses the GitHub Search API (no authentication required, but rate-limited
    to ~10 requests/minute for unauthenticated users).

    Returns a list of dicts with: ``full_name``, ``description``, ``stars``,
    ``license``, ``url``.
    """
    encoded_query = urllib.parse.quote(f"{query} pptx template")
    url = f"{_GITHUB_API_BASE}/search/repositories?q={encoded_query}&sort=stars&per_page={max_results}"
    try:
        data = _fetch_json(url)
    except (urllib.error.HTTPError, urllib.error.URLError, json.JSONDecodeError) as exc:
        code = getattr(exc, "code", None)
        if code == 403:
            logger.warning("GitHub API rate limit exceeded")
            raise RuntimeError("GitHub API rate limit exceeded. Try again in a few minutes.") from exc
        raise RuntimeError(f"GitHub search failed: {exc}") from exc

    items = data.get("items", [])
    results: list[dict[str, Any]] = []
    for item in items[:max_results]:
        license_info = item.get("license")
        license_name = license_info.get("spdx_id") if isinstance(license_info, dict) and license_info else None
        results.append({
            "full_name": item.get("full_name", ""),
            "description": item.get("description", "") or "",
            "stars": item.get("stargazers_count", 0),
            "license": license_name or "Unknown",
            "url": item.get("html_url", ""),
        })
    return results


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _iso_now() -> str:
    """Return current UTC time as ISO-8601 string (no external deps)."""
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).isoformat()
