"""Manifest V3 model, persistence and legacy v2 migration.

Manifest V3 stores the content model, layout plans, render trace and every
plan/QA/repair attempt. Legacy v2 manifests (embedded XML wrapper with
``skill_version=2`` JSON payload) are migrated to V3 in memory only; existing
files are not rewritten until an explicit V3 save is requested.
"""
from __future__ import annotations

import copy
import json
import os
import tempfile
import zipfile
from dataclasses import MISSING, asdict, dataclass, field, is_dataclass
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pptx_skill.content_model import (
    ContentSpec,
    LayoutPlan,
)
from pptx_skill.pptx_renderer import RenderTraceEntry

# ---------------------------------------------------------------------------
# XML constants
# ---------------------------------------------------------------------------

CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_XML_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
MANIFEST_PART = "customXml/pptxSkillManifest.xml"

V2_MANIFEST_NS = "urn:openai:pptx-skill:manifest:v2"
V3_MANIFEST_NS = "urn:openai:pptx-skill:manifest:v3"

ET.register_namespace("", CUSTOM_NS)
ET.register_namespace("vt", VT_NS)


# ---------------------------------------------------------------------------
# V3 model
# ---------------------------------------------------------------------------

@dataclass
class AttemptRecord:
    """One generation/QA/repair pass."""

    run_id: str
    pass_index: int
    plan_artifact: str | None = None
    trace_artifact: str | None = None
    qa_artifact: str | None = None
    repairs_applied: list[dict[str, Any]] = field(default_factory=list)
    sha256: dict[str, str] = field(default_factory=dict)


@dataclass
class ManifestV3:
    """In-memory Manifest V3 representation."""

    manifest_schema_version: int = 3
    template_schema_version: int = 2
    current: dict[str, Any] = field(default_factory=dict)
    attempts: list[AttemptRecord] = field(default_factory=list)
    renderer_environment: dict[str, Any] = field(default_factory=dict)
    repair_log: list[dict[str, Any]] = field(default_factory=list)
    manifest_stale: bool = False
    legacy: dict[str, Any] = field(default_factory=dict)

    @property
    def content(self) -> dict[str, Any]:
        return self.current.setdefault("content", {})

    @property
    def layout_plans(self) -> list[dict[str, Any]]:
        return self.current.setdefault("layout_plans", [])

    @property
    def render_trace(self) -> list[dict[str, Any]]:
        return self.current.setdefault("render_trace", [])

    def record_attempt(
        self,
        pass_index: int,
        plan_artifact: str | None = None,
        trace_artifact: str | None = None,
        qa_artifact: str | None = None,
        repairs: list[dict[str, Any]] | None = None,
        sha256: dict[str, str] | None = None,
    ) -> AttemptRecord:
        run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + f"-{pass_index:04d}"
        attempt = AttemptRecord(
            run_id=run_id,
            pass_index=pass_index,
            plan_artifact=plan_artifact,
            trace_artifact=trace_artifact,
            qa_artifact=qa_artifact,
            repairs_applied=repairs or [],
            sha256=sha256 or {},
        )
        self.attempts.append(attempt)
        return attempt


# ---------------------------------------------------------------------------
# Serialization helpers
# ---------------------------------------------------------------------------

def _serialize_value(value: Any) -> Any:
    """Recursively turn dataclasses/enums into JSON-serializable values."""
    if is_dataclass(value) and not isinstance(value, type):
        result = {k: _serialize_value(v) for k, v in asdict(value).items()}
        extra = getattr(value, "_extra", None)
        if extra:
            result.update({k: _serialize_value(v) for k, v in extra.items()})
        return result
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, list):
        return [_serialize_value(v) for v in value]
    if isinstance(value, dict):
        return {k: _serialize_value(v) for k, v in value.items()}
    if isinstance(value, tuple):
        return {"__tuple__": [_serialize_value(v) for v in value]}
    return value


def _to_plain_dict(obj: Any) -> Any:
    """Serialize a dataclass instance to a plain dict/list."""
    return _serialize_value(obj)


def manifest_to_dict(manifest: ManifestV3) -> dict[str, Any]:
    return _to_plain_dict(manifest)


def _deserialize_value(value: Any) -> Any:
    """Inverse of _serialize_value: restore tuples and other wrapped types."""
    if isinstance(value, dict):
        if "__tuple__" in value and len(value) == 1:
            return tuple(_deserialize_value(v) for v in value["__tuple__"])
        return {k: _deserialize_value(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_deserialize_value(v) for v in value]
    return value


def _deserialize_dataclass(cls: type, data: dict[str, Any]) -> Any:
    """Best-effort dataclass reconstruction from a plain dict.

    Unknown fields (not present in the dataclass definition) are preserved
    via a ``_extra`` attribute so that round-trip serialization does not
    lose data when the schema evolves.
    """
    fields_map: Any = getattr(cls, "__dataclass_fields__", None) or {}
    field_types = {f.name: f.type for f in fields_map.values()}
    kwargs: dict[str, Any] = {}
    extra: dict[str, Any] = {}
    for name, value in data.items():
        value = _deserialize_value(value)
        if name in field_types:
            kwargs[name] = value
        else:
            extra[name] = value
    try:
        obj = cls(**kwargs)
    except TypeError:
        defaults = {f.name: f.default for f in fields_map.values() if f.default is not MISSING}
        defaults.update({f.name: f.default_factory() for f in fields_map.values() if f.default_factory is not MISSING})
        for k in fields_map:
            if k not in kwargs and k in defaults:
                kwargs[k] = defaults[k]
        obj = cls(**kwargs)
    if extra:
        obj._extra = extra
    return obj


def manifest_from_dict(data: dict[str, Any]) -> ManifestV3:
    data = copy.deepcopy(data)
    data = _deserialize_value(data)
    attempts = [
        _deserialize_dataclass(AttemptRecord, a)
        for a in data.get("attempts", [])
    ]
    manifest = ManifestV3(
        manifest_schema_version=data.get("manifest_schema_version", 3),
        template_schema_version=data.get("template_schema_version", 2),
        current=data.get("current", {}),
        attempts=attempts,
        renderer_environment=data.get("renderer_environment", {}),
        repair_log=data.get("repair_log", []),
        manifest_stale=data.get("manifest_stale", False),
        legacy=data.get("legacy", {}),
    )
    return manifest


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def _manifest_xml(manifest_json: str, namespace: str, version: str) -> bytes:
    root = ET.Element(f"{{{namespace}}}manifest", {"version": version})
    payload = ET.SubElement(root, f"{{{namespace}}}json")
    payload.text = manifest_json
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _presentation_rels_xml(existing: bytes) -> bytes:
    try:
        root = ET.fromstring(existing)
    except ET.ParseError:
        return existing
    relationship_tag = f"{{{REL_NS}}}Relationship"
    target_name = "../customXml/pptxSkillManifest.xml"
    if not any(
        rel.get("Type") == CUSTOM_XML_REL_TYPE and rel.get("Target") == target_name
        for rel in root.findall(relationship_tag)
    ):
        used = {rel.get("Id") for rel in root.findall(relationship_tag)}
        number = 1
        while f"rId{number}" in used:
            number += 1
        ET.SubElement(
            root,
            relationship_tag,
            {
                "Id": f"rId{number}",
                "Type": CUSTOM_XML_REL_TYPE,
                "Target": target_name,
            },
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _embed_manifest_xml(pptx_path: Path, manifest_json: str, namespace: str, version: str) -> None:
    with zipfile.ZipFile(pptx_path, "r") as source:
        info_list = source.infolist()
        members = {item.filename: source.read(item.filename) for item in info_list}
        infos = {item.filename: item for item in info_list}
    members[MANIFEST_PART] = _manifest_xml(manifest_json, namespace, version)
    presentation_rels = "ppt/_rels/presentation.xml.rels"
    if presentation_rels in members:
        members[presentation_rels] = _presentation_rels_xml(members[presentation_rels])

    temp_fd, temp_name = tempfile.mkstemp(suffix=".pptx", dir=str(pptx_path.parent))
    os.close(temp_fd)
    try:
        with zipfile.ZipFile(temp_name, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for name, data in members.items():
                if name in infos:
                    target.writestr(infos[name], data)
                else:
                    target.writestr(name, data)
        os.replace(temp_name, pptx_path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def _read_embedded_manifest(pptx_path: Path) -> tuple[dict[str, Any] | None, str | None]:
    """Return (payload, namespace_version) or (None, None) if not found/unreadable."""
    try:
        with zipfile.ZipFile(pptx_path, "r") as package:
            data = package.read(MANIFEST_PART)
    except (KeyError, zipfile.BadZipFile):
        return None, None
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return None, None
    namespace = None
    version = None
    if root.tag.startswith(f"{{{V2_MANIFEST_NS}}}"):
        namespace = V2_MANIFEST_NS
        version = root.get("version") or "2"
    elif root.tag.startswith(f"{{{V3_MANIFEST_NS}}}"):
        namespace = V3_MANIFEST_NS
        version = root.get("version") or "3"
    else:
        return None, None
    value = root.find(f"{{{namespace}}}json")
    if value is None or not value.text:
        return None, None
    try:
        return json.loads(value.text), f"{namespace}#{version}"
    except json.JSONDecodeError:
        return None, None


# ---------------------------------------------------------------------------
# Sidecar helpers
# ---------------------------------------------------------------------------

def _sidecar_path(pptx_path: str | Path) -> Path:
    path = Path(pptx_path)
    return path.with_suffix(".manifest.json")


def _write_sidecar(pptx_path: Path, payload: dict[str, Any]) -> None:
    sidecar = _sidecar_path(pptx_path)
    with sidecar.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def _read_sidecar(pptx_path: Path) -> dict[str, Any] | None:
    sidecar = _sidecar_path(pptx_path)
    if not sidecar.exists():
        return None
    try:
        with sidecar.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None


# ---------------------------------------------------------------------------
# Legacy v2 -> V3 migration
# ---------------------------------------------------------------------------

def migrate_v2_to_v3(v2_payload: dict[str, Any]) -> ManifestV3:
    """Migrate a legacy v2 manifest payload to Manifest V3 in memory."""
    v2 = copy.deepcopy(v2_payload)
    sections = v2.get("sections", [])
    layouts = v2.get("layouts", [])

    # Build a minimal ContentSpec-shaped current.content from legacy sections.
    slides: list[dict[str, Any]] = []
    for idx, section in enumerate(sections):
        slides.append(
            {
                "id": f"legacy/sec-{idx:02d}",
                "role": section.get("layout") or "bullets",
                "communication_goal": section.get("title", ""),
                "elements": [
                    {
                        "id": f"legacy/sec-{idx:02d}/title",
                        "kind": "text",
                        "role": "title",
                        "content": {"text": section.get("title", "")},
                        "style_ref": "component.title",
                        "constraints": [],
                        "decorative": False,
                    }
                ],
                "preferred_layouts": [section.get("layout")] if section.get("layout") else [],
                "source_section_id": f"legacy/sec-{idx:02d}",
                "fragment_index": 0,
            }
        )

    manifest = ManifestV3(
        manifest_schema_version=3,
        template_schema_version=2,
        current={
            "content": {
                "id": v2.get("title", "untitled"),
                "title": v2.get("title", ""),
                "subtitle": v2.get("subtitle", ""),
                "slides": slides,
                "locale": v2.get("lang", "zh-CN"),
                "metadata": {
                    "migrated_from": "legacy_v2",
                    "original_generated_at": v2.get("generated_at"),
                    "original_layouts": layouts,
                },
            },
            "layout_plans": [],
            "render_trace": [],
        },
        legacy={
            "theme_key": v2.get("theme_key"),
            "template_key": v2.get("template_key"),
            "source_skill_version": v2.get("skill_version", 2),
            "original_image_dir": v2.get("image_dir"),
            "original_auto_search_images": v2.get("auto_search_images"),
        },
    )
    return manifest


# ---------------------------------------------------------------------------
# Public load/save API
# ---------------------------------------------------------------------------

class ManifestError(Exception):
    """Raised when a manifest cannot be read, parsed or migrated."""

    pass


def load_manifest(pptx_path: str | Path) -> ManifestV3 | None:
    """Load a manifest from embedded XML or sidecar, migrating legacy v2 to V3.

    Returns ``None`` when no manifest exists. Raises ``ManifestError`` when a
    manifest is present but unreadable or malformed.
    """
    path = Path(pptx_path).resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    payload: dict[str, Any] | None = None
    source_version: str | None = None

    embedded, embedded_version = _read_embedded_manifest(path)
    if embedded is not None:
        payload = embedded
        source_version = embedded_version

    if payload is None:
        payload = _read_sidecar(path)
        source_version = "sidecar"

    if payload is None:
        return None

    if not isinstance(payload, dict):
        raise ManifestError(f"Manifest must contain a JSON object: {path}")

    schema_version = payload.get("manifest_schema_version")
    if schema_version == 3:
        manifest = manifest_from_dict(payload)
        manifest.legacy.setdefault("loaded_from", source_version)
        return manifest

    if schema_version is not None and schema_version > 3:
        raise ManifestError(f"Unsupported manifest schema version {schema_version}: {path}")

    # Treat payload without manifest_schema_version as legacy v2.
    manifest = migrate_v2_to_v3(payload)
    manifest.legacy["loaded_from"] = source_version
    manifest.legacy["migrated_at"] = datetime.now(timezone.utc).isoformat()
    return manifest


def save_manifest_v3(
    pptx_path: str | Path,
    manifest: ManifestV3,
    write_embedded: bool = True,
) -> None:
    """Save a Manifest V3 as sidecar and optionally embedded XML."""
    path = Path(pptx_path).resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    manifest.manifest_schema_version = 3
    payload = manifest_to_dict(manifest)
    _write_sidecar(path, payload)

    if write_embedded:
        manifest_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        _embed_manifest_xml(path, manifest_json, V3_MANIFEST_NS, "3")


# ---------------------------------------------------------------------------
# Convenience: convert domain objects into manifest.current
# ---------------------------------------------------------------------------

def set_current_content(manifest: ManifestV3, content: ContentSpec) -> None:
    manifest.current["content"] = _to_plain_dict(content)


def set_current_plans(manifest: ManifestV3, plans: list[LayoutPlan]) -> None:
    manifest.current["layout_plans"] = [_to_plain_dict(p) for p in plans]


def set_current_trace(manifest: ManifestV3, trace: list[RenderTraceEntry]) -> None:
    manifest.current["render_trace"] = [_to_plain_dict(t) for t in trace]
