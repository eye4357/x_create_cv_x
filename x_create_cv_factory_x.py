#!/usr/bin/env python3
"""Command-line CRUD factory for app-native CV JSON data."""

from __future__ import annotations

import argparse
import base64
import hashlib
import html
import json
import os
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Sequence
from pathlib import Path
from typing import Any

Relationship = tuple[str, str, str] | tuple[str, str, str, str]
MASTER_FILE = "master_profile.json"
DEFAULT_DB_DIR = Path("data/private/cv_factory_output")
SCHEMA_DIR = Path(__file__).resolve().parent / "schemas"
AUDIT_POLICY_DIR = Path(__file__).resolve().parent / "audit_policies"
DEFAULT_AUDIT_POLICY = AUDIT_POLICY_DIR / "default_office_audit_policy.json"
RESUME_APP_MODEL = "resume_document_crud"
MASTER_APP_MODEL = "master_profile_crud"
SCHEMA_VERSION = "1.0.0"
VERSION = "0.0.3"
CONTRACT_SCHEMA_FILES = {
    "master_profile": "master_profile.schema.json",
    "resume": "resume.schema.json",
    "workbook_layout": "workbook_layout.schema.json",
    "document_layout": "document_layout.schema.json",
    "audit_policy": "audit_policy.schema.json",
}
DEFAULT_GOLDEN_EVIDENCE_DIR = Path("../x_create_cv_test_data_x/evidence")
DEFAULT_EVIDENCE_MANIFEST = DEFAULT_GOLDEN_EVIDENCE_DIR / "a_priori_manifest.json"
DEFAULT_CHAIN_MANIFEST = DEFAULT_GOLDEN_EVIDENCE_DIR / "chain_of_evidence_manifest.json"
HASH_CHUNK_SIZE = 1024 * 1024
GOLDEN_SCRIPT_FILES = [
    "01_build_master_data.py",
    "02_build_old_resume.py",
    "03_build_new_resume.py",
    "04_build_current_resume.py",
]
GOLDEN_JSON_EXPECTATIONS = {
    MASTER_FILE: "generated/json/a_posteriori/master_profile_a_posteriori.json",
    "resume_2017.json": "generated/json/a_posteriori/resume_2017_a_posteriori.json",
    "resume_2023.json": "generated/json/a_posteriori/resume_2023_a_posteriori.json",
    "resume_2024.json": "generated/json/a_posteriori/resume_2024_a_posteriori.json",
}
GOLDEN_OFFICE_EXPECTATIONS = {
    "master_profile_a_posteriori.xlsx": "generated/office/a_posteriori/master_profile_a_posteriori.xlsx",
    "resume_2017_a_posteriori.docx": "generated/office/a_posteriori/resume_2017_a_posteriori.docx",
    "resume_2023_a_posteriori.docx": "generated/office/a_posteriori/resume_2023_a_posteriori.docx",
    "resume_2024_a_posteriori.docx": "generated/office/a_posteriori/resume_2024_a_posteriori.docx",
}
GOLDEN_SOURCE_OFFICE = {
    "master_profile_a_posteriori.xlsx": "source_office/a_priori/R_cv_2023_0501_1427_a_priori.xlsx",
    "resume_2017_a_posteriori.docx": "source_office/a_priori/R_cv_2017_1129_0848_a_priori.docx",
    "resume_2023_a_posteriori.docx": "source_office/a_priori/R_cv_2023_0315_2158_a_priori.docx",
    "resume_2024_a_posteriori.docx": "source_office/a_priori/R_cv_2024_1206_0000_a_priori.docx",
}
GOLDEN_OFFICE_REPORT = "reports/a_posteriori_office_comparison.json"
GOLDEN_OFFICE_AUDIT_REPORT = "reports/a_posteriori_office_audit.md"
ZIP_TIMESTAMP = (2026, 7, 1, 0, 0, 0)
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DEFAULT_WORKBOOK_SHEETS = [
    ("Highlights", "highlights", ["id", "label", "highlights"]),
    (
        "Jobs",
        "jobs",
        [
            "id",
            "label",
            "job_title",
            "job_focus",
            "job_highlights",
            "vehicle_program",
            "employer",
            "address",
            "reason_for_leaving",
            "salary_usd",
            "start_date",
            "end_date",
            "time_spent",
        ],
    ),
    (
        "Standards Development",
        "standards",
        ["id", "label", "organization", "title_of_standard", "role", "start_date", "end_date", "time_spent"],
    ),
    (
        "School",
        "education",
        ["id", "label", "school", "address", "degree", "start_date", "end_date", "time_spent"],
    ),
    (
        "Certifications",
        "certifications",
        ["id", "label", "certificate_subject", "certificate_name", "certificate_number"],
    ),
    ("Patents", "patents", ["id", "label", "patent_name", "patent_number", "patent_url"]),
    (
        "Publications",
        "publications",
        ["id", "label", "year", "publication_name", "publisher", "publication_reference", "publication_url"],
    ),
    (
        "Lectures",
        "lectures",
        ["id", "label", "year", "publication_or_lecture_name", "publisher_name"],
    ),
    (
        "Residences",
        "residences",
        ["id", "label", "residence", "start_date", "end_date", "time_spent"],
    ),
]
DEFAULT_DOCUMENT_MARGINS: dict[str, dict[str, Any]] = {
    "resume_2017_a_posteriori.docx": {
        "margins": {"top": "1440", "right": "1440", "bottom": "1440", "left": "1440", "header": "720", "footer": "720"},
    },
    "resume_2023_a_posteriori.docx": {
        "margins": {"top": "1440", "right": "1800", "bottom": "1440", "left": "1800", "header": "720", "footer": "720"},
    },
}
DEFAULT_DOCX_STYLES: list[dict[str, Any]] = [
    {"id": "Normal", "name": "Normal", "default": True, "font": "Calibri", "size": "22"},
    {"id": "Title", "name": "Title", "based_on": "Normal", "bold": True, "size": "32", "spacing_after": "120"},
    {
        "id": "Heading1",
        "name": "heading 1",
        "based_on": "Normal",
        "bold": True,
        "size": "24",
        "spacing_before": "160",
        "spacing_after": "80",
    },
    {
        "id": "Heading2",
        "name": "heading 2",
        "based_on": "Normal",
        "bold": True,
        "size": "22",
        "spacing_before": "120",
        "spacing_after": "60",
    },
    {"id": "ListParagraph", "name": "List Paragraph", "based_on": "Normal", "indent_left": "720"},
    {
        "id": "ListBullet",
        "name": "List Bullet",
        "based_on": "ListParagraph",
        "numbering": {"level": "0", "num_id": "1"},
    },
]

MASTER_COLLECTIONS = [
    "people",
    "contact_methods",
    "addresses",
    "jobs",
    "achievements",
    "projects",
    "highlights",
    "standards",
    "education",
    "certifications",
    "patents",
    "publications",
    "lectures",
    "residences",
    "skills",
    "text_blocks",
]

MASTER_COMMANDS = {
    "add-user": "people",
    "add-contact-method": "contact_methods",
    "add-address": "addresses",
    "add-job": "jobs",
    "add-achievement": "achievements",
    "add-project": "projects",
    "add-highlight": "highlights",
    "add-standard": "standards",
    "add-education": "education",
    "add-certification": "certifications",
    "add-patent": "patents",
    "add-publication": "publications",
    "add-lecture": "lectures",
    "add-residence": "residences",
    "add-skill": "skills",
    "add-text-block": "text_blocks",
}

EXPECTED_FILES = [MASTER_FILE, "resume_2017.json", "resume_2023.json"]

MASTER_FIELD_ORDER = {
    "people": ["id", "display_name", "preferred_name"],
    "contact_methods": ["id", "person_id", "kind", "value", "is_primary"],
    "addresses": ["id", "label", "address_text"],
    "jobs": [
        "id",
        "label",
        "a",
        "b",
        "c",
        "d",
        "e",
        "f",
        "g",
        "h",
        "i",
        "j",
        "k",
        "job_title",
        "job_focus",
        "job_highlights",
        "vehicle_program",
        "employer",
        "address",
        "reason_for_leaving",
        "salary_usd",
        "start_date",
        "end_date",
        "time_spent",
        "address_id",
        "project_ids",
        "achievement_ids",
    ],
    "achievements": ["id", "text", "job_id"],
    "projects": ["id", "name"],
    "highlights": ["id", "label", "a", "highlights"],
    "standards": [
        "id",
        "label",
        "a",
        "b",
        "c",
        "d",
        "e",
        "f",
        "organization",
        "title_of_standard",
        "role",
        "start_date",
        "end_date",
        "time_spent",
    ],
    "education": [
        "id",
        "label",
        "a",
        "b",
        "c",
        "d",
        "e",
        "f",
        "school",
        "address",
        "degree",
        "start_date",
        "end_date",
        "time_spent",
        "address_id",
    ],
    "certifications": ["id", "label", "a", "b", "c", "certificate_subject", "certificate_name", "certificate_number"],
    "patents": ["id", "label", "a", "b", "c", "patent_name", "patent_number", "patent_url"],
    "publications": [
        "id",
        "label",
        "a",
        "b",
        "c",
        "d",
        "e",
        "year",
        "publication_name",
        "publisher",
        "publication_reference",
        "publication_url",
    ],
    "lectures": ["id", "label", "a", "b", "c", "year", "publication_or_lecture_name", "publisher_name"],
    "residences": [
        "id",
        "label",
        "a",
        "b",
        "c",
        "d",
        "residence",
        "start_date",
        "end_date",
        "time_spent",
        "address_id",
    ],
    "skills": ["id", "name", "category"],
    "text_blocks": ["id", "text", "section_label", "used_in_resume_ids", "is_job_related"],
}

RESUME_FIELD_ORDER = ["id", "label", "status", "created_at", "updated_at"]
SECTION_FIELD_ORDER = ["id", "title", "kind", "sort_order", "item_ids", "is_visible"]
ITEM_FIELD_ORDER = [
    "id",
    "section_id",
    "kind",
    "sort_order",
    "master_record_refs",
    "text_override",
    "formatting",
    "is_visible",
]
BOOL_FIELDS = {"is_primary", "is_job_related", "is_visible"}
LIST_FIELDS = {"achievement_ids", "item_ids", "master_record_refs", "project_ids", "used_in_resume_ids"}
DICT_FIELDS = {"formatting"}


def parse_json_payload(json_text: str | None, json_base64: str | None, payload_name: str) -> dict[str, Any]:
    if json_text is None and json_base64 is None:
        raise ValueError(f"{payload_name} JSON is required")
    if json_text is not None and json_base64 is not None:
        raise ValueError(f"Pass either --json or --json-base64 for {payload_name}, not both")

    if json_base64 is not None:
        try:
            json_text = base64.b64decode(json_base64).decode("utf-8")
        except ValueError as exc:
            raise ValueError(f"Invalid base64 for {payload_name}") from exc

    try:
        value = json.loads(json_text or "{}")
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON for {payload_name}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"{payload_name} JSON must be an object")
    return value


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise ValueError(f"Missing JSON file: {path}")
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"Expected JSON object in {path}")
    return value


def schema_type_names(schema_type: Any) -> list[str]:
    if isinstance(schema_type, str):
        return [schema_type]
    if isinstance(schema_type, list):
        return [item for item in schema_type if isinstance(item, str)]
    return []


def json_type_name(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, int):
        return "integer"
    if isinstance(value, float):
        return "number"
    if isinstance(value, str):
        return "string"
    if isinstance(value, list):
        return "array"
    if isinstance(value, dict):
        return "object"
    return type(value).__name__


def value_matches_json_type(value: Any, type_name: str) -> bool:
    if type_name == "null":
        return value is None
    if type_name == "boolean":
        return isinstance(value, bool)
    if type_name == "integer":
        return isinstance(value, int) and not isinstance(value, bool)
    if type_name == "number":
        return isinstance(value, int | float) and not isinstance(value, bool)
    if type_name == "string":
        return isinstance(value, str)
    if type_name == "array":
        return isinstance(value, list)
    if type_name == "object":
        return isinstance(value, dict)
    return True


def child_json_path(path: str, key: str) -> str:
    return f"{path}.{key}" if key.isidentifier() else f"{path}[{key!r}]"


def validate_json_schema_subset(value: Any, schema: dict[str, Any], path: str = "$") -> list[str]:
    failures: list[str] = []
    expected_types = schema_type_names(schema.get("type"))
    if expected_types and not any(value_matches_json_type(value, expected_type) for expected_type in expected_types):
        failures.append(f"{path}: expected {' or '.join(expected_types)}, got {json_type_name(value)}")
        return failures

    enum_values = schema.get("enum")
    if isinstance(enum_values, list) and value not in enum_values:
        failures.append(f"{path}: expected one of {enum_values!r}, got {value!r}")

    if isinstance(value, dict):
        properties = schema.get("properties")
        property_schemas = properties if isinstance(properties, dict) else {}
        required = schema.get("required")
        if isinstance(required, list):
            for key in required:
                if isinstance(key, str) and key not in value:
                    failures.append(f"{child_json_path(path, key)}: required property is missing")
        for key, child_schema in property_schemas.items():
            if key in value and isinstance(key, str) and isinstance(child_schema, dict):
                failures.extend(validate_json_schema_subset(value[key], child_schema, child_json_path(path, key)))
        additional_properties = schema.get("additionalProperties", True)
        if additional_properties is False:
            for key in value:
                if key not in property_schemas:
                    failures.append(f"{child_json_path(path, str(key))}: additional property is not allowed")
        elif isinstance(additional_properties, dict):
            for key, child_value in value.items():
                if key not in property_schemas:
                    failures.extend(
                        validate_json_schema_subset(child_value, additional_properties, child_json_path(path, str(key)))
                    )

    if isinstance(value, list):
        items = schema.get("items")
        if isinstance(items, dict):
            for index, item in enumerate(value):
                failures.extend(validate_json_schema_subset(item, items, f"{path}[{index}]"))

    return failures


def contract_schema_path(schema_name: str) -> Path:
    schema_file = CONTRACT_SCHEMA_FILES.get(schema_name)
    if schema_file is None:
        known = ", ".join(sorted(CONTRACT_SCHEMA_FILES))
        raise ValueError(f"Unknown contract schema {schema_name!r}; expected one of: {known}")
    return SCHEMA_DIR / schema_file


def load_contract_schema(schema_name: str) -> dict[str, Any]:
    return read_json(contract_schema_path(schema_name))


def contract_validation_failures(data: dict[str, Any], schema_name: str, path: str = "$") -> list[str]:
    failures = validate_json_schema_subset(data, load_contract_schema(schema_name), path)
    office_layout = data.get("office_layout")
    if schema_name == "master_profile" and isinstance(office_layout, dict):
        workbook = office_layout.get("workbook")
        if isinstance(workbook, dict):
            failures.extend(contract_validation_failures(workbook, "workbook_layout", f"{path}.office_layout.workbook"))
    if schema_name == "resume" and isinstance(office_layout, dict):
        document = office_layout.get("document")
        if isinstance(document, dict):
            failures.extend(contract_validation_failures(document, "document_layout", f"{path}.office_layout.document"))
    return failures


def validate_contract(data: dict[str, Any], schema_name: str) -> None:
    failures = contract_validation_failures(data, schema_name)
    if failures:
        raise ValueError(f"Schema validation failed for {schema_name}:\n" + "\n".join(failures))


def infer_contract_schema_name(data: dict[str, Any], path: Path) -> str:
    app_model = data.get("app_model")
    if app_model == MASTER_APP_MODEL:
        return "master_profile"
    if app_model == RESUME_APP_MODEL:
        return "resume"
    if "policy_version" in data and "accepted_differences" in data:
        return "audit_policy"
    if "sheets" in data:
        return "workbook_layout"
    if any(
        key in data for key in ["margins", "page_size", "package", "styles", "numbering", "header", "footer", "flow"]
    ):
        return "document_layout"
    raise ValueError(f"Could not infer contract schema for {path}")


def validate_contract_file(path: Path, schema_name: str | None = None) -> str:
    data = read_json(path)
    resolved_schema_name = schema_name or infer_contract_schema_name(data, path)
    validate_contract(data, resolved_schema_name)
    return resolved_schema_name


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(HASH_CHUNK_SIZE), b""):
            digest.update(chunk)
    return digest.hexdigest()


def check_evidence_manifest(manifest_path: Path) -> int:
    manifest = read_json(manifest_path)
    if manifest.get("algorithm") != "sha256":
        raise ValueError("Evidence manifest must use sha256")

    entries = manifest.get("files")
    if not isinstance(entries, list) or not entries:
        raise ValueError("Evidence manifest must contain a non-empty files list")

    manifest_dir = manifest_path.parent.resolve()
    for entry in entries:
        if not isinstance(entry, dict):
            raise ValueError("Evidence manifest entries must be JSON objects")
        relative_path = entry.get("path")
        expected_sha256 = entry.get("sha256")
        expected_bytes = entry.get("bytes")
        if not isinstance(relative_path, str) or not relative_path:
            raise ValueError("Evidence manifest entries must contain a non-empty path")
        if not isinstance(expected_sha256, str) or len(expected_sha256) != 64:
            raise ValueError(f"Evidence manifest entry has an invalid sha256: {relative_path}")
        if not isinstance(expected_bytes, int) or expected_bytes < 0:
            raise ValueError(f"Evidence manifest entry has an invalid byte count: {relative_path}")

        evidence_path = (manifest_dir / relative_path).resolve()
        try:
            evidence_path.relative_to(manifest_dir)
        except ValueError as exc:
            raise ValueError(f"Evidence path escapes manifest directory: {relative_path}") from exc
        if not evidence_path.exists():
            raise ValueError(f"Missing evidence file: {relative_path}")

        actual_bytes = evidence_path.stat().st_size
        if actual_bytes != expected_bytes:
            raise ValueError(
                f"Evidence size mismatch {relative_path}: expected {expected_bytes} bytes, got {actual_bytes} bytes"
            )

        actual_sha256 = sha256_file(evidence_path)
        if actual_sha256 != expected_sha256.lower():
            raise ValueError(f"Evidence hash mismatch {relative_path}: expected {expected_sha256}, got {actual_sha256}")

    return len(entries)


def scalar_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, dict | list):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return str(value)


def clean_xml_text(value: Any) -> str:
    text = scalar_text(value)
    return "".join(character if character in "\t\n\r" or ord(character) >= 32 else " " for character in text)


def xml_text(value: Any) -> str:
    return html.escape(clean_xml_text(value), quote=False)


def xml_attr(value: Any) -> str:
    return html.escape(clean_xml_text(value), quote=True)


def write_zip_entries(path: Path, entries: list[tuple[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in entries:
            info = zipfile.ZipInfo(name, ZIP_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, content.encode("utf-8"))


def relationships_xml(relationships: Sequence[Relationship]) -> str:
    entries = ""
    for relationship in relationships:
        relationship_id, relationship_type, target = relationship[:3]
        target_mode = f' TargetMode="{xml_attr(relationship[3])}"' if len(relationship) == 4 else ""
        entries += (
            f'<Relationship Id="{xml_attr(relationship_id)}" Type="{xml_attr(relationship_type)}" '
            f'Target="{xml_attr(target)}"{target_mode}/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{entries}</Relationships>"
    )


def column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def header_label(field: str) -> str:
    if len(field) == 1 and field.isalpha():
        return field.upper()
    return " ".join(part.upper() if part == "id" else part.capitalize() for part in field.split("_"))


def default_workbook_layout() -> dict[str, Any]:
    return {
        "styles": {
            "freeze_header": True,
            "auto_filter": True,
            "header_fill": "FF1F4E79",
            "header_font_color": "FFFFFFFF",
            "body_border_color": "FFBFBFBF",
            "page_margins": {
                "left": "0.7",
                "right": "0.7",
                "top": "0.75",
                "bottom": "0.75",
                "header": "0.3",
                "footer": "0.3",
            },
        },
        "sheets": [
            {
                "name": sheet_name,
                "collection": collection_name,
                "skip_placeholder_rows": True,
                "columns": [{"field": field, "header": header_label(field)} for field in fields],
            }
            for sheet_name, collection_name, fields in DEFAULT_WORKBOOK_SHEETS
        ],
    }


def default_document_layout(file_name: str) -> dict[str, Any]:
    layout = DEFAULT_DOCUMENT_MARGINS.get(file_name, DEFAULT_DOCUMENT_MARGINS["resume_2017_a_posteriori.docx"])
    return {
        "margins": dict(layout["margins"]),
        "package": {
            "core_properties": True,
            "extended_properties": True,
            "theme": True,
            "font_table": True,
            "web_settings": True,
            "footnotes": True,
            "endnotes": True,
            "custom_xml": True,
            "fonts": ["Calibri", "Symbol"],
        },
        "styles": [dict(style) for style in DEFAULT_DOCX_STYLES],
        "numbering": {"bullet_text": "•", "font": "Symbol", "left": "720", "hanging": "360"},
    }


def workbook_layout(master: dict[str, Any]) -> dict[str, Any]:
    layout = master.get("office_layout")
    workbook = layout.get("workbook") if isinstance(layout, dict) else None
    if isinstance(workbook, dict) and isinstance(workbook.get("sheets"), list):
        return workbook
    return default_workbook_layout()


def workbook_style_options(layout: dict[str, Any]) -> dict[str, Any]:
    styles = layout.get("styles")
    return styles if isinstance(styles, dict) else {}


def workbook_sheet_columns(sheet: dict[str, Any]) -> list[tuple[str, str, str]]:
    columns = sheet.get("columns")
    column_defs: list[tuple[str, str, str]] = []
    if isinstance(columns, list):
        for column in columns:
            if not isinstance(column, dict):
                continue
            field = column.get("field")
            if not isinstance(field, str) or not field:
                continue
            header = scalar_text(column.get("header")) or header_label(field)
            value_type = scalar_text(column.get("value_type"))
            if value_type not in {"string", "number", "boolean"}:
                value_type = "string"
            column_defs.append((field, header, value_type))
    return column_defs


def placeholder_collection_label(collection_name: str) -> str:
    return " ".join(part.capitalize() for part in collection_name.split("_"))


def is_placeholder_workbook_record(
    record: dict[str, Any], columns: list[tuple[str, str, str]], collection_name: str
) -> bool:
    label = scalar_text(record.get("label"))
    if label != placeholder_collection_label(collection_name):
        return False
    return all(
        field in {"id", "label"} or not scalar_text(record.get(field)) for field, _header, _value_type in columns
    )


def sheet_rows(master: dict[str, Any], sheet: dict[str, Any]) -> list[list[str]]:
    columns = workbook_sheet_columns(sheet)
    rows = [[header for _field, header, _value_type in columns]]
    collections = master.get("collections")
    collection_name = sheet.get("collection")
    if not isinstance(collection_name, str) or not collection_name:
        return rows
    records = collections.get(collection_name) if isinstance(collections, dict) else []
    if not isinstance(records, list):
        return rows
    for record in records:
        if isinstance(record, dict):
            if sheet.get("skip_placeholder_rows") is not False and is_placeholder_workbook_record(
                record, columns, collection_name
            ):
                continue
            rows.append([scalar_text(record.get(field)) for field, _header, _value_type in columns])
    return rows


def workbook_sheets(master: dict[str, Any]) -> list[tuple[str, list[list[str]], dict[str, Any]]]:
    sheets = workbook_layout(master).get("sheets")
    if not isinstance(sheets, list):
        return []
    rendered_sheets: list[tuple[str, list[list[str]], dict[str, Any]]] = []
    for sheet in sheets:
        if not isinstance(sheet, dict):
            continue
        sheet_name = scalar_text(sheet.get("name") or sheet.get("collection") or "Sheet")
        rendered_sheets.append((sheet_name, sheet_rows(master, sheet), sheet))
    return rendered_sheets


def worksheet_dimension(rows: list[list[str]]) -> str:
    row_count = max(len(rows), 1)
    column_count = max((len(row) for row in rows), default=1)
    return f"A1:{column_name(column_count)}{row_count}"


def configured_column_width(sheet: dict[str, Any], column_index: int) -> float | None:
    widths = sheet.get("column_widths")
    if not isinstance(widths, list) or column_index >= len(widths):
        return None
    width = widths[column_index]
    if isinstance(width, int | float) and not isinstance(width, bool):
        return float(max(1, min(width, 255)))
    return None


def column_widths(rows: list[list[str]], sheet: dict[str, Any]) -> list[float]:
    column_count = max((len(row) for row in rows), default=1)
    widths: list[float] = []
    for column_index in range(column_count):
        configured_width = configured_column_width(sheet, column_index)
        if configured_width is not None:
            widths.append(configured_width)
            continue
        width = max((len(row[column_index]) for row in rows if column_index < len(row)), default=8)
        widths.append(float(max(10, min(width + 2, 42))))
    return widths


def style_bool(styles: dict[str, Any], key: str, default: bool) -> bool:
    value = styles.get(key)
    return value if isinstance(value, bool) else default


def sheet_bool(sheet: dict[str, Any], styles: dict[str, Any], key: str, default: bool) -> bool:
    value = sheet.get(key)
    return value if isinstance(value, bool) else style_bool(styles, key, default)


def workbook_column_value_types(sheet: dict[str, Any]) -> list[str]:
    return [value_type for _field, _header, value_type in workbook_sheet_columns(sheet)]


def xlsx_number_value(value: str) -> str | None:
    stripped = value.strip()
    if not stripped:
        return None
    try:
        number = float(stripped)
    except ValueError:
        return None
    if number != number or number in {float("inf"), float("-inf")}:
        return None
    return stripped


def xlsx_boolean_value(value: str) -> str | None:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes"}:
        return "1"
    if normalized in {"0", "false", "no"}:
        return "0"
    return None


def xlsx_cell_xml(cell_ref: str, style_id: str, value: str, value_type: str) -> str:
    if value_type == "number":
        number_value = xlsx_number_value(value)
        if number_value is not None:
            return f'<c r="{cell_ref}" s="{style_id}"><v>{xml_text(number_value)}</v></c>'
    if value_type == "boolean":
        boolean_value = xlsx_boolean_value(value)
        if boolean_value is not None:
            return f'<c r="{cell_ref}" s="{style_id}" t="b"><v>{boolean_value}</v></c>'
    return (
        f'<c r="{cell_ref}" s="{style_id}" t="inlineStr"><is><t xml:space="preserve">' f"{xml_text(value)}</t></is></c>"
    )


def style_text(styles: dict[str, Any], key: str, default: str) -> str:
    return scalar_text(styles.get(key)) or default


def workbook_page_margins(styles: dict[str, Any]) -> dict[str, str]:
    value = styles.get("page_margins")
    margins = value if isinstance(value, dict) else {}
    return {
        "left": scalar_text(margins.get("left") if isinstance(margins, dict) else None) or "0.7",
        "right": scalar_text(margins.get("right") if isinstance(margins, dict) else None) or "0.7",
        "top": scalar_text(margins.get("top") if isinstance(margins, dict) else None) or "0.75",
        "bottom": scalar_text(margins.get("bottom") if isinstance(margins, dict) else None) or "0.75",
        "header": scalar_text(margins.get("header") if isinstance(margins, dict) else None) or "0.3",
        "footer": scalar_text(margins.get("footer") if isinstance(margins, dict) else None) or "0.3",
    }


def worksheet_xml(rows: list[list[str]], layout: dict[str, Any], sheet: dict[str, Any]) -> str:
    styles = workbook_style_options(layout)
    value_types = workbook_column_value_types(sheet)
    rendered_rows: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells: list[str] = []
        for column_index, value in enumerate(row, start=1):
            cell_ref = f"{column_name(column_index)}{row_index}"
            style_id = "1" if row_index == 1 else "2"
            value_type = "string"
            if row_index != 1 and column_index <= len(value_types):
                value_type = value_types[column_index - 1]
            cells.append(xlsx_cell_xml(cell_ref, style_id, value, value_type))
        rendered_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    widths = "".join(
        f'<col min="{index}" max="{index}" width="{width:.2f}" customWidth="1"/>'
        for index, width in enumerate(column_widths(rows, sheet), start=1)
    )
    dimension = worksheet_dimension(rows)
    sheet_view_xml = '<sheetViews><sheetView tabSelected="0" workbookViewId="0">'
    if sheet_bool(sheet, styles, "freeze_header", True):
        sheet_view_xml += '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        sheet_view_xml += '<selection pane="bottomLeft"/>'
    sheet_view_xml += "</sheetView></sheetViews>"
    auto_filter_xml = f'<autoFilter ref="{dimension}"/>' if sheet_bool(sheet, styles, "auto_filter", True) else ""
    margins = workbook_page_margins(styles)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{SHEET_NS}" xmlns:r="{REL_NS}">'
        f'<dimension ref="{dimension}"/>'
        f"{sheet_view_xml}"
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"<cols>{widths}</cols>"
        f'<sheetData>{"".join(rendered_rows)}</sheetData>'
        f"{auto_filter_xml}"
        f'<pageMargins left="{xml_attr(margins["left"])}" right="{xml_attr(margins["right"])}" '
        f'top="{xml_attr(margins["top"])}" bottom="{xml_attr(margins["bottom"])}" '
        f'header="{xml_attr(margins["header"])}" footer="{xml_attr(margins["footer"])}"/>'
        "</worksheet>"
    )


def workbook_content_types(sheet_count: int) -> str:
    sheet_overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{sheet_overrides}"
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = "".join(
        f'<sheet name="{xml_attr(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, sheet_name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{SHEET_NS}" xmlns:r="{REL_NS}"><sheets>{sheets}</sheets></workbook>'
    )


def workbook_relationships(sheet_count: int) -> str:
    relationships = [
        (
            f"rId{index}",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            f"worksheets/sheet{index}.xml",
        )
        for index in range(1, sheet_count + 1)
    ]
    relationships.append(
        (
            f"rId{sheet_count + 1}",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml",
        )
    )
    return relationships_xml(relationships)


def workbook_styles_xml(layout: dict[str, Any]) -> str:
    styles = workbook_style_options(layout)
    header_fill = style_text(styles, "header_fill", "FF1F4E79")
    header_font_color = style_text(styles, "header_font_color", "FFFFFFFF")
    body_border_color = style_text(styles, "body_border_color", "FFBFBFBF")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{SHEET_NS}">'
        '<fonts count="3">'
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
        f'<font><b/><sz val="11"/><color rgb="{xml_attr(header_font_color)}"/>'
        '<name val="Calibri"/><family val="2"/></font>'
        '<font><i/><sz val="11"/><color rgb="FF666666"/><name val="Calibri"/><family val="2"/></font>'
        '</fonts><fills count="4">'
        '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        f'<fill><patternFill patternType="solid"><fgColor rgb="{xml_attr(header_fill)}"/>'
        '<bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        f'<border><left style="thin"><color rgb="{xml_attr(body_border_color)}"/></left>'
        f'<right style="thin"><color rgb="{xml_attr(body_border_color)}"/></right>'
        f'<top style="thin"><color rgb="{xml_attr(body_border_color)}"/></top>'
        f'<bottom style="thin"><color rgb="{xml_attr(body_border_color)}"/></bottom>'
        "<diagonal/></border></borders>"
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def write_master_workbook(master: dict[str, Any], path: Path) -> None:
    layout = workbook_layout(master)
    sheets = workbook_sheets(master)
    sheet_names = [sheet_name for sheet_name, _rows, _sheet in sheets]
    entries = [
        ("[Content_Types].xml", workbook_content_types(len(sheets))),
        (
            "_rels/.rels",
            relationships_xml(
                [
                    (
                        "rId1",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                        "xl/workbook.xml",
                    ),
                    (
                        "rId2",
                        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                        "docProps/core.xml",
                    ),
                    (
                        "rId3",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                        "docProps/app.xml",
                    ),
                ]
            ),
        ),
        ("xl/workbook.xml", workbook_xml(sheet_names)),
        ("xl/_rels/workbook.xml.rels", workbook_relationships(len(sheets))),
        ("xl/styles.xml", workbook_styles_xml(layout)),
        *[
            (f"xl/worksheets/sheet{index}.xml", worksheet_xml(rows, layout, sheet))
            for index, (_sheet_name, rows, sheet) in enumerate(sheets, start=1)
        ],
        (
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:title>Master Profile A Posteriori</dc:title><dc:creator>x_create_cv_x</dc:creator>"
            "<cp:lastModifiedBy>x_create_cv_x</cp:lastModifiedBy>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:modified>'
            "</cp:coreProperties>",
        ),
        (
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>x_create_cv_x</Application></Properties>",
        ),
    ]
    write_zip_entries(path, entries)


def master_records_by_id(master: dict[str, Any]) -> dict[str, dict[str, Any]]:
    records_by_id: dict[str, dict[str, Any]] = {}
    collections = master.get("collections")
    if not isinstance(collections, dict):
        return records_by_id
    for records in collections.values():
        if not isinstance(records, list):
            continue
        for record in records:
            if isinstance(record, dict) and isinstance(record.get("id"), str):
                records_by_id[str(record["id"])] = record
    return records_by_id


def sort_order(record: dict[str, Any]) -> int:
    value = record.get("sort_order")
    return value if isinstance(value, int) else 0


def record_summary(record: dict[str, Any]) -> str:
    preferred_fields = [
        "text",
        "name",
        "label",
        "job_title",
        "employer",
        "degree",
        "school",
        "patent_name",
        "publication_name",
        "certificate_name",
    ]
    parts = [scalar_text(record[field]) for field in preferred_fields if scalar_text(record.get(field))]
    if parts:
        return " - ".join(parts[:3])
    for field, value in record.items():
        if field != "id" and scalar_text(value):
            return scalar_text(value)
    return scalar_text(record.get("id"))


def resume_item_text(item: dict[str, Any], records_by_id: dict[str, dict[str, Any]]) -> str:
    text_override = item.get("text_override")
    if isinstance(text_override, str) and text_override.strip():
        return text_override
    refs = item.get("master_record_refs")
    if isinstance(refs, list):
        summaries = [
            record_summary(records_by_id[ref]) for ref in refs if isinstance(ref, str) and ref in records_by_id
        ]
        return "\n".join(summary for summary in summaries if summary)
    return ""


def document_layout(resume: dict[str, Any], file_name: str) -> dict[str, Any]:
    default_layout = default_document_layout(file_name)
    layout = resume.get("office_layout")
    document = layout.get("document") if isinstance(layout, dict) else None
    if not isinstance(document, dict):
        return default_layout
    merged = dict(default_layout)
    merged.update(document)
    default_margins = default_layout.get("margins")
    document_margins = document.get("margins")
    if isinstance(default_margins, dict) and isinstance(document_margins, dict):
        margins = dict(default_margins)
        margins.update(document_margins)
        merged["margins"] = margins
    return merged


def item_formatting(item: dict[str, Any]) -> dict[str, Any]:
    formatting = item.get("formatting")
    return formatting if isinstance(formatting, dict) else {}


def formatting_style(formatting: dict[str, Any]) -> str | None:
    block_style = formatting.get("block_style")
    return block_style if isinstance(block_style, str) and block_style else None


def formatting_numbering(formatting: dict[str, Any]) -> dict[str, str] | None:
    numbering = formatting.get("numbering")
    if not isinstance(numbering, dict):
        return None
    level = scalar_text(numbering.get("level")) or "0"
    num_id = scalar_text(numbering.get("num_id"))
    if not num_id:
        return None
    return {"level": level, "num_id": num_id}


def formatting_runs(formatting: dict[str, Any]) -> list[dict[str, Any]]:
    runs = formatting.get("runs")
    rendered_runs: list[dict[str, Any]] = []
    if not isinstance(runs, list):
        return rendered_runs
    for run in runs:
        if not isinstance(run, dict):
            continue
        text = scalar_text(run.get("text"))
        has_tab = run.get("tab") is True or run.get("tabs") is not None
        if not text and not has_tab:
            continue
        rendered_run: dict[str, Any] = {
            "text": text,
            "bold": run.get("bold") is True,
            "italic": run.get("italic") is True,
            "underline": run.get("underline") is True,
        }
        for key in ["style", "font", "size", "color", "link"]:
            value = scalar_text(run.get(key))
            if value:
                rendered_run[key] = value
        if has_tab:
            rendered_run["tabs"] = scalar_text(run.get("tabs")) or "1"
        rendered_runs.append(rendered_run)
    return rendered_runs


def paragraph_layout_options(value: dict[str, Any]) -> dict[str, Any]:
    options: dict[str, Any] = {}
    alignment = scalar_text(value.get("alignment")) or scalar_text(value.get("align"))
    if alignment:
        options["alignment"] = alignment
    spacing = value.get("spacing")
    if isinstance(spacing, dict):
        rendered_spacing = {key: scalar_text(item) for key, item in spacing.items() if scalar_text(item)}
        if rendered_spacing:
            options["spacing"] = rendered_spacing
    indent = value.get("indent")
    if isinstance(indent, dict):
        rendered_indent = {key: scalar_text(item) for key, item in indent.items() if scalar_text(item)}
        if rendered_indent:
            options["indent"] = rendered_indent
    tab_stops = value.get("tab_stops") or value.get("tabs")
    if isinstance(tab_stops, list):
        rendered_tab_stops = [
            {key: scalar_text(item) for key, item in tab_stop.items() if scalar_text(item)}
            for tab_stop in tab_stops
            if isinstance(tab_stop, dict)
        ]
        rendered_tab_stops = [tab_stop for tab_stop in rendered_tab_stops if tab_stop]
        if rendered_tab_stops:
            options["tab_stops"] = rendered_tab_stops
    return options


def paragraph_contract(
    *, style: str | None, numbering: dict[str, Any] | None, runs: list[dict[str, Any]], layout: dict[str, Any]
) -> dict[str, Any]:
    paragraph: dict[str, Any] = {"style": style, "numbering": numbering, "runs": runs}
    paragraph.update(paragraph_layout_options(layout))
    return paragraph


def item_paragraphs(item: dict[str, Any], records_by_id: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    formatting = item_formatting(item)
    style = formatting_style(formatting)
    numbering = formatting_numbering(formatting)
    runs = formatting_runs(formatting)
    if runs:
        return [paragraph_contract(style=style, numbering=numbering, runs=runs, layout=formatting)]

    text = resume_item_text(item, records_by_id)
    paragraphs: list[dict[str, Any]] = []
    for paragraph in text.splitlines() or [text]:
        if paragraph.strip():
            paragraphs.append(
                paragraph_contract(style=style, numbering=numbering, runs=[{"text": paragraph}], layout=formatting)
            )
    return paragraphs


def layout_paragraphs(value: dict[str, Any]) -> list[dict[str, Any]]:
    style = formatting_style(value)
    if style is None:
        explicit_style = value.get("style")
        style = explicit_style if isinstance(explicit_style, str) and explicit_style else None
    numbering = formatting_numbering(value)
    runs = formatting_runs(value)
    if runs:
        return [paragraph_contract(style=style, numbering=numbering, runs=runs, layout=value)]
    if value.get("empty") is True:
        return [paragraph_contract(style=style, numbering=numbering, runs=[{"text": ""}], layout=value)]
    text = scalar_text(value.get("text"))
    return [paragraph_contract(style=style, numbering=numbering, runs=[{"text": text}], layout=value)] if text else []


def resume_item_records(resume: dict[str, Any]) -> list[dict[str, Any]]:
    collections_value = resume.get("collections")
    collections = collections_value if isinstance(collections_value, dict) else {}
    items = collections.get("items")
    return [item for item in items if isinstance(item, dict)] if isinstance(items, list) else []


def visible_resume_item_lookup(resume: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {
        item_id: item
        for item in resume_item_records(resume)
        if item.get("is_visible") is not False and isinstance((item_id := item.get("id")), str)
    }


def flow_item_paragraphs(
    item_ids: Any, item_lookup: dict[str, dict[str, Any]], records_by_id: dict[str, dict[str, Any]]
) -> list[dict[str, Any]]:
    if not isinstance(item_ids, list):
        return []
    paragraphs: list[dict[str, Any]] = []
    for item_id in item_ids:
        if isinstance(item_id, str) and item_id in item_lookup:
            paragraphs.extend(item_paragraphs(item_lookup[item_id], records_by_id))
    return paragraphs


def flow_cell_paragraphs(
    cell: Any, item_lookup: dict[str, dict[str, Any]], records_by_id: dict[str, dict[str, Any]]
) -> list[dict[str, Any]]:
    if not isinstance(cell, dict):
        return []
    paragraphs = flow_item_paragraphs(cell.get("item_ids"), item_lookup, records_by_id)
    extra_paragraphs = cell.get("paragraphs")
    if isinstance(extra_paragraphs, list):
        for paragraph in extra_paragraphs:
            if isinstance(paragraph, dict):
                paragraphs.extend(layout_paragraphs(paragraph))
    if paragraphs:
        return paragraphs
    return layout_paragraphs(cell)


def resume_flow_blocks(master: dict[str, Any], resume: dict[str, Any], layout: dict[str, Any]) -> list[dict[str, Any]]:
    flow = layout.get("flow")
    if not isinstance(flow, list):
        return []
    records_by_id = master_records_by_id(master)
    item_lookup = visible_resume_item_lookup(resume)
    blocks: list[dict[str, Any]] = []
    for entry in flow:
        if not isinstance(entry, dict):
            continue
        entry_type = entry.get("type")
        if entry_type == "item":
            paragraphs = flow_item_paragraphs([entry.get("item_id")], item_lookup, records_by_id)
            blocks.extend({"type": "paragraph", "paragraph": paragraph} for paragraph in paragraphs)
        elif entry_type == "items":
            paragraphs = flow_item_paragraphs(entry.get("item_ids"), item_lookup, records_by_id)
            blocks.extend({"type": "paragraph", "paragraph": paragraph} for paragraph in paragraphs)
        elif entry_type == "paragraph":
            blocks.extend({"type": "paragraph", "paragraph": paragraph} for paragraph in layout_paragraphs(entry))
        elif entry_type == "table":
            rows_value = entry.get("rows")
            if not isinstance(rows_value, list):
                continue
            rows: list[list[list[dict[str, Any]]]] = []
            for row_value in rows_value:
                if not isinstance(row_value, list):
                    continue
                row = [flow_cell_paragraphs(cell, item_lookup, records_by_id) for cell in row_value]
                rows.append(row)
            blocks.append(
                {
                    "type": "table",
                    "rows": rows,
                    "column_widths": entry.get("column_widths"),
                    "style": entry.get("style"),
                }
            )
    return blocks


def resume_paragraphs(master: dict[str, Any], resume: dict[str, Any]) -> list[dict[str, Any]]:
    records_by_id = master_records_by_id(master)
    resume_value = resume.get("resume")
    resume_meta = resume_value if isinstance(resume_value, dict) else {}

    collections_value = resume.get("collections")
    collections = collections_value if isinstance(collections_value, dict) else {}
    sections = collections.get("sections")
    items = collections.get("items")
    section_records = (
        [section for section in sections if isinstance(section, dict)] if isinstance(sections, list) else []
    )
    item_records = [item for item in items if isinstance(item, dict)] if isinstance(items, list) else []
    visible_section_ids = {section.get("id") for section in section_records if section.get("is_visible") is not False}
    section_order = {
        section.get("id"): sort_order(section)
        for section in section_records
        if isinstance(section.get("id"), str) and section.get("is_visible") is not False
    }

    paragraphs: list[dict[str, Any]] = []
    for item in sorted(
        item_records,
        key=lambda record: (section_order.get(record.get("section_id"), 1_000_000), sort_order(record)),
    ):
        if item.get("is_visible") is False:
            continue
        if visible_section_ids and item.get("section_id") not in visible_section_ids:
            continue
        paragraphs.extend(item_paragraphs(item, records_by_id))

    if paragraphs:
        return paragraphs

    label = scalar_text(resume_meta.get("label") or resume_meta.get("id") or "Resume")
    return [{"style": None, "numbering": None, "runs": [{"text": label}]}]


def resume_blocks(master: dict[str, Any], resume: dict[str, Any], layout: dict[str, Any]) -> list[dict[str, Any]]:
    flow_blocks = resume_flow_blocks(master, resume, layout)
    if flow_blocks:
        return flow_blocks
    return [{"type": "paragraph", "paragraph": paragraph} for paragraph in resume_paragraphs(master, resume)]


def docx_run_properties(run: dict[str, Any]) -> str:
    properties: list[str] = []
    style = scalar_text(run.get("style"))
    if style:
        properties.append(f'<w:rStyle w:val="{xml_attr(style)}"/>')
    font = scalar_text(run.get("font"))
    if font:
        properties.append(f'<w:rFonts w:ascii="{xml_attr(font)}" w:hAnsi="{xml_attr(font)}" w:cs="{xml_attr(font)}"/>')
    if run.get("bold") is True:
        properties.append("<w:b/>")
    if run.get("italic") is True:
        properties.append("<w:i/>")
    if run.get("underline") is True:
        properties.append('<w:u w:val="single"/>')
    color = scalar_text(run.get("color"))
    if color:
        properties.append(f'<w:color w:val="{xml_attr(color)}"/>')
    size = scalar_text(run.get("size"))
    if size:
        properties.append(f'<w:sz w:val="{xml_attr(size)}"/>')
    return f'<w:rPr>{"".join(properties)}</w:rPr>' if properties else ""


def run_tabs(run: dict[str, Any]) -> int:
    tabs = scalar_text(run.get("tabs"))
    if tabs.isdecimal():
        return max(0, int(tabs))
    return 1 if run.get("tab") is True else 0


def run_link_target(run: dict[str, Any]) -> str | None:
    target = scalar_text(run.get("link")) or scalar_text(run.get("url"))
    return target or None


def docx_run(run: dict[str, Any]) -> str:
    properties_xml = docx_run_properties(run)
    tabs_xml = "<w:tab/>" * run_tabs(run)
    text = scalar_text(run.get("text"))
    text_xml = f'<w:t xml:space="preserve">{xml_text(text)}</w:t>' if text else ""
    return f"<w:r>{properties_xml}{tabs_xml}{text_xml}</w:r>"


def docx_paragraph_run(run: dict[str, Any], hyperlink_ids: dict[str, str]) -> str:
    run_xml = docx_run(run)
    link_target = run_link_target(run)
    relationship_id = hyperlink_ids.get(link_target or "")
    if relationship_id:
        return f'<w:hyperlink r:id="{xml_attr(relationship_id)}" w:history="1">{run_xml}</w:hyperlink>'
    return run_xml


def docx_paragraph_spacing(paragraph: dict[str, Any]) -> str:
    spacing = paragraph.get("spacing")
    if not isinstance(spacing, dict):
        return ""
    attrs: list[str] = []
    for key, word_key in [("before", "before"), ("after", "after"), ("line", "line"), ("line_rule", "lineRule")]:
        value = scalar_text(spacing.get(key)) or scalar_text(spacing.get(word_key))
        if value:
            attrs.append(f'w:{word_key}="{xml_attr(value)}"')
    return f'<w:spacing {" ".join(attrs)}/>' if attrs else ""


def docx_paragraph_indent(paragraph: dict[str, Any]) -> str:
    indent = paragraph.get("indent")
    if not isinstance(indent, dict):
        return ""
    attrs: list[str] = []
    for key, word_key in [
        ("left", "left"),
        ("right", "right"),
        ("hanging", "hanging"),
        ("first_line", "firstLine"),
    ]:
        value = scalar_text(indent.get(key)) or scalar_text(indent.get(word_key))
        if value:
            attrs.append(f'w:{word_key}="{xml_attr(value)}"')
    return f'<w:ind {" ".join(attrs)}/>' if attrs else ""


def docx_paragraph_tab_stops(paragraph: dict[str, Any]) -> str:
    tab_stops = paragraph.get("tab_stops")
    if not isinstance(tab_stops, list):
        return ""
    rendered: list[str] = []
    for tab_stop in tab_stops:
        if not isinstance(tab_stop, dict):
            continue
        value = scalar_text(tab_stop.get("value") or tab_stop.get("val"))
        position = scalar_text(tab_stop.get("position") or tab_stop.get("pos"))
        leader = scalar_text(tab_stop.get("leader"))
        attrs = []
        if value:
            attrs.append(f'w:val="{xml_attr(value)}"')
        if leader:
            attrs.append(f'w:leader="{xml_attr(leader)}"')
        if position:
            attrs.append(f'w:pos="{xml_attr(position)}"')
        if attrs:
            rendered.append(f'<w:tab {" ".join(attrs)}/>')
    return f'<w:tabs>{"".join(rendered)}</w:tabs>' if rendered else ""


def docx_paragraph(paragraph: dict[str, Any], hyperlink_ids: dict[str, str] | None = None) -> str:
    paragraph_properties: list[str] = []
    style = paragraph.get("style")
    if isinstance(style, str) and style:
        paragraph_properties.append(f'<w:pStyle w:val="{xml_attr(style)}"/>')
    for property_xml in [
        docx_paragraph_tab_stops(paragraph),
        docx_paragraph_spacing(paragraph),
        docx_paragraph_indent(paragraph),
    ]:
        if property_xml:
            paragraph_properties.append(property_xml)
    alignment = scalar_text(paragraph.get("alignment"))
    if alignment:
        paragraph_properties.append(f'<w:jc w:val="{xml_attr(alignment)}"/>')
    numbering = paragraph.get("numbering")
    if isinstance(numbering, dict):
        level = scalar_text(numbering.get("level")) or "0"
        num_id = scalar_text(numbering.get("num_id"))
        if num_id:
            paragraph_properties.append(
                f'<w:numPr><w:ilvl w:val="{xml_attr(level)}"/><w:numId w:val="{xml_attr(num_id)}"/></w:numPr>'
            )
    properties_xml = f'<w:pPr>{"".join(paragraph_properties)}</w:pPr>' if paragraph_properties else ""
    runs = paragraph.get("runs")
    rendered_runs = [docx_run(run) for run in runs if isinstance(run, dict)] if isinstance(runs, list) else []
    if hyperlink_ids is not None and isinstance(runs, list):
        rendered_runs = [docx_paragraph_run(run, hyperlink_ids) for run in runs if isinstance(run, dict)]
    if not rendered_runs:
        rendered_runs = [docx_run({"text": ""})]
    return f'<w:p>{properties_xml}{"".join(rendered_runs)}</w:p>'


def table_column_widths(table: dict[str, Any], column_count: int) -> list[str]:
    widths = table.get("column_widths")
    if isinstance(widths, list):
        rendered = [scalar_text(width) for width in widths[:column_count]]
        if len(rendered) == column_count and all(rendered):
            return rendered
    if column_count <= 0:
        return []
    default_width = str(max(1, 8640 // column_count))
    return [default_width] * column_count


def docx_table_cell(paragraphs: list[dict[str, Any]], width: str, hyperlink_ids: dict[str, str]) -> str:
    rendered_paragraphs = "".join(
        docx_paragraph(paragraph, hyperlink_ids) for paragraph in paragraphs
    ) or docx_paragraph({"runs": [{"text": ""}]})
    return f'<w:tc><w:tcPr><w:tcW w:w="{xml_attr(width)}" w:type="dxa"/></w:tcPr>{rendered_paragraphs}</w:tc>'


def docx_table(table: dict[str, Any], hyperlink_ids: dict[str, str]) -> str:
    rows = table.get("rows")
    if not isinstance(rows, list):
        return ""
    column_count = max((len(row) for row in rows if isinstance(row, list)), default=0)
    widths = table_column_widths(table, column_count)
    grid = "".join(f'<w:gridCol w:w="{xml_attr(width)}"/>' for width in widths)
    rendered_rows: list[str] = []
    for row in rows:
        if not isinstance(row, list):
            continue
        cells = "".join(
            docx_table_cell(
                cell if isinstance(cell, list) else [], widths[index] if index < len(widths) else "2160", hyperlink_ids
            )
            for index, cell in enumerate(row)
        )
        rendered_rows.append(f"<w:tr>{cells}</w:tr>")
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>{''.join(rendered_rows)}</w:tbl>"
    )


def docx_body_block(block: dict[str, Any], hyperlink_ids: dict[str, str]) -> str:
    block_type = block.get("type")
    if block_type == "table":
        return docx_table(block, hyperlink_ids)
    paragraph = block.get("paragraph")
    return docx_paragraph(paragraph, hyperlink_ids) if isinstance(paragraph, dict) else ""


def block_paragraphs(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    paragraphs: list[dict[str, Any]] = []
    for block in blocks:
        if block.get("type") == "table":
            rows = block.get("rows")
            if isinstance(rows, list):
                for row in rows:
                    if isinstance(row, list):
                        for cell in row:
                            if isinstance(cell, list):
                                paragraphs.extend(paragraph for paragraph in cell if isinstance(paragraph, dict))
        else:
            paragraph = block.get("paragraph")
            if isinstance(paragraph, dict):
                paragraphs.append(paragraph)
    return paragraphs


def docx_margins(layout: dict[str, Any]) -> dict[str, str]:
    value = layout.get("margins")
    if not isinstance(value, dict):
        value = {}
    return {
        "top": scalar_text(value.get("top") or "1440"),
        "right": scalar_text(value.get("right") or "1440"),
        "bottom": scalar_text(value.get("bottom") or "1440"),
        "left": scalar_text(value.get("left") or "1440"),
        "header": scalar_text(value.get("header") or "720"),
        "footer": scalar_text(value.get("footer") or "720"),
    }


def docx_page_size(layout: dict[str, Any]) -> dict[str, str]:
    value = layout.get("page_size")
    if not isinstance(value, dict):
        value = {}
    page_size = {
        "w": scalar_text(value.get("w") or "12240"),
        "h": scalar_text(value.get("h") or "15840"),
    }
    orientation = scalar_text(value.get("orient") or value.get("orientation"))
    if orientation:
        page_size["orient"] = orientation
    return page_size


def docx_section_references(layout: dict[str, Any]) -> str:
    references = ""
    if isinstance(layout.get("header"), dict):
        references += '<w:headerReference w:type="default" r:id="rId7"/>'
    if isinstance(layout.get("footer"), dict):
        references += '<w:footerReference w:type="default" r:id="rId8"/>'
    return references


def docx_document_xml(blocks: list[dict[str, Any]], layout: dict[str, Any], hyperlink_ids: dict[str, str]) -> str:
    body_xml = "".join(docx_body_block(block, hyperlink_ids) for block in blocks)
    margins = docx_margins(layout)
    page_size = docx_page_size(layout)
    orientation_xml = f' w:orient="{xml_attr(page_size["orient"])}"' if "orient" in page_size else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}"><w:body>{body_xml}'
        f"<w:sectPr>{docx_section_references(layout)}"
        f'<w:pgSz w:w="{xml_attr(page_size["w"])}" w:h="{xml_attr(page_size["h"])}"{orientation_xml}/>'
        f'<w:pgMar w:top="{margins["top"]}" w:right="{margins["right"]}" '
        f'w:bottom="{margins["bottom"]}" w:left="{margins["left"]}" '
        f'w:header="{margins["header"]}" w:footer="{margins["footer"]}" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


def document_package_options(layout: dict[str, Any]) -> dict[str, Any]:
    package = layout.get("package")
    return package if isinstance(package, dict) else {}


def package_enabled(layout: dict[str, Any], key: str) -> bool:
    value = document_package_options(layout).get(key)
    return value is True


def document_content_types(layout: dict[str, Any]) -> str:
    optional_overrides = ""
    if package_enabled(layout, "theme"):
        optional_overrides += (
            '<Override PartName="/word/theme/theme1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        )
    if package_enabled(layout, "font_table"):
        optional_overrides += (
            '<Override PartName="/word/fontTable.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        )
    if package_enabled(layout, "web_settings"):
        optional_overrides += (
            '<Override PartName="/word/webSettings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>'
        )
    if package_enabled(layout, "footnotes"):
        optional_overrides += (
            '<Override PartName="/word/footnotes.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        )
    if package_enabled(layout, "endnotes"):
        optional_overrides += (
            '<Override PartName="/word/endnotes.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        )
    if package_enabled(layout, "custom_xml"):
        optional_overrides += (
            '<Override PartName="/customXml/itemProps1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
        )
    if isinstance(layout.get("header"), dict):
        optional_overrides += (
            '<Override PartName="/word/header1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        )
    if isinstance(layout.get("footer"), dict):
        optional_overrides += (
            '<Override PartName="/word/footer1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        )
    if package_enabled(layout, "core_properties"):
        optional_overrides += (
            '<Override PartName="/docProps/core.xml" '
            'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        )
    if package_enabled(layout, "extended_properties"):
        optional_overrides += (
            '<Override PartName="/docProps/app.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/numbering.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        '<Override PartName="/word/settings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        f"{optional_overrides}"
        "</Types>"
    )


def document_styles(layout: dict[str, Any]) -> list[dict[str, Any]]:
    styles = layout.get("styles")
    if isinstance(styles, list) and all(isinstance(style, dict) for style in styles):
        return [style for style in styles if isinstance(style.get("id"), str)]
    return [dict(style) for style in DEFAULT_DOCX_STYLES]


def docx_style_numbering_xml(numbering: dict[str, Any]) -> str:
    level = scalar_text(numbering.get("level")) or "0"
    num_id = scalar_text(numbering.get("num_id"))
    if not num_id:
        return ""
    return f'<w:numPr><w:ilvl w:val="{xml_attr(level)}"/><w:numId w:val="{xml_attr(num_id)}"/></w:numPr>'


def docx_style_xml(style: dict[str, Any]) -> str:
    style_id = scalar_text(style.get("id"))
    style_name = scalar_text(style.get("name")) or style_id
    default_attr = ' w:default="1"' if style.get("default") is True else ""
    based_on = scalar_text(style.get("based_on"))
    based_on_xml = f'<w:basedOn w:val="{xml_attr(based_on)}"/>' if based_on else ""

    paragraph_properties: list[str] = []
    spacing_attrs: list[str] = []
    spacing_before = scalar_text(style.get("spacing_before"))
    spacing_after = scalar_text(style.get("spacing_after"))
    if spacing_before:
        spacing_attrs.append(f'w:before="{xml_attr(spacing_before)}"')
    if spacing_after:
        spacing_attrs.append(f'w:after="{xml_attr(spacing_after)}"')
    if spacing_attrs:
        paragraph_properties.append(f'<w:spacing {" ".join(spacing_attrs)}/>')
    indent_left = scalar_text(style.get("indent_left"))
    indent_hanging = scalar_text(style.get("indent_hanging"))
    indent_attrs: list[str] = []
    if indent_left:
        indent_attrs.append(f'w:left="{xml_attr(indent_left)}"')
    if indent_hanging:
        indent_attrs.append(f'w:hanging="{xml_attr(indent_hanging)}"')
    if indent_attrs:
        paragraph_properties.append(f'<w:ind {" ".join(indent_attrs)}/>')
    numbering = style.get("numbering")
    if isinstance(numbering, dict):
        paragraph_properties.append(docx_style_numbering_xml(numbering))
    paragraph_properties_xml = f'<w:pPr>{"".join(paragraph_properties)}</w:pPr>' if paragraph_properties else ""

    run_properties: list[str] = []
    font = scalar_text(style.get("font"))
    if font:
        run_properties.append(f'<w:rFonts w:ascii="{xml_attr(font)}" w:hAnsi="{xml_attr(font)}"/>')
    if style.get("bold") is True:
        run_properties.append("<w:b/>")
    if style.get("italic") is True:
        run_properties.append("<w:i/>")
    if style.get("underline") is True:
        run_properties.append('<w:u w:val="single"/>')
    size = scalar_text(style.get("size"))
    if size:
        run_properties.append(f'<w:sz w:val="{xml_attr(size)}"/>')
    run_properties_xml = f'<w:rPr>{"".join(run_properties)}</w:rPr>' if run_properties else ""

    return (
        f'<w:style w:type="paragraph"{default_attr} w:styleId="{xml_attr(style_id)}">'
        f'<w:name w:val="{xml_attr(style_name)}"/>'
        f"{based_on_xml}{paragraph_properties_xml}{run_properties_xml}</w:style>"
    )


def docx_styles_xml(layout: dict[str, Any]) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}">'
        f'{"".join(docx_style_xml(style) for style in document_styles(layout))}'
        "</w:styles>"
    )


def docx_numbering_ids(paragraphs: list[dict[str, Any]], layout: dict[str, Any]) -> list[str]:
    num_ids: set[str] = set()
    for paragraph in paragraphs:
        numbering = paragraph.get("numbering")
        if isinstance(numbering, dict):
            num_id = scalar_text(numbering.get("num_id"))
            if num_id:
                num_ids.add(num_id)
    for style in document_styles(layout):
        numbering = style.get("numbering")
        if isinstance(numbering, dict):
            num_id = scalar_text(numbering.get("num_id"))
            if num_id:
                num_ids.add(num_id)
    return sorted(num_ids, key=lambda value: int(value) if value.isdecimal() else 1_000_000)


def docx_numbering_xml(paragraphs: list[dict[str, Any]], layout: dict[str, Any]) -> str:
    numbering_options = layout.get("numbering")
    numbering = numbering_options if isinstance(numbering_options, dict) else {}
    bullet_text = scalar_text(numbering.get("bullet_text")) or "•"
    font = scalar_text(numbering.get("font")) or "Symbol"
    left = scalar_text(numbering.get("left")) or "720"
    hanging = scalar_text(numbering.get("hanging")) or "360"
    abstract_nums: list[str] = []
    nums: list[str] = []
    for fallback_index, num_id in enumerate(docx_numbering_ids(paragraphs, layout), start=1):
        abstract_num_id = num_id if num_id.isdecimal() else str(fallback_index)
        abstract_nums.append(
            f'<w:abstractNum w:abstractNumId="{xml_attr(abstract_num_id)}"><w:multiLevelType w:val="hybridMultilevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
            f'<w:lvlText w:val="{xml_attr(bullet_text)}"/>'
            f'<w:pPr><w:ind w:left="{xml_attr(left)}" w:hanging="{xml_attr(hanging)}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{xml_attr(font)}" w:hAnsi="{xml_attr(font)}" w:hint="default"/></w:rPr></w:lvl>'
            '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
            f'<w:lvlText w:val="{xml_attr(bullet_text)}"/>'
            f'<w:pPr><w:ind w:left="{xml_attr(str(int(left) * 2 if left.isdecimal() else 1440))}" '
            f'w:hanging="{xml_attr(hanging)}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{xml_attr(font)}" w:hAnsi="{xml_attr(font)}" w:hint="default"/></w:rPr></w:lvl>'
            "</w:abstractNum>"
        )
        nums.append(
            f'<w:num w:numId="{xml_attr(num_id)}"><w:abstractNumId w:val="{xml_attr(abstract_num_id)}"/></w:num>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{WORD_NS}">{"".join(abstract_nums)}{"".join(nums)}</w:numbering>'
    )


def docx_document_relationships(layout: dict[str, Any]) -> list[Relationship]:
    relationships: list[Relationship] = [
        ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"),
        ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml"),
        ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml"),
    ]
    if package_enabled(layout, "theme"):
        relationships.append(
            ("rId4", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml")
        )
    if package_enabled(layout, "font_table"):
        relationships.append(
            ("rId5", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable", "fontTable.xml")
        )
    if package_enabled(layout, "web_settings"):
        relationships.append(
            (
                "rId6",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
                "webSettings.xml",
            )
        )
    if isinstance(layout.get("header"), dict):
        relationships.append(
            ("rId7", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header1.xml")
        )
    if isinstance(layout.get("footer"), dict):
        relationships.append(
            ("rId8", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer1.xml")
        )
    if package_enabled(layout, "footnotes"):
        relationships.append(
            ("rId9", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", "footnotes.xml")
        )
    if package_enabled(layout, "endnotes"):
        relationships.append(
            ("rId10", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", "endnotes.xml")
        )
    if package_enabled(layout, "custom_xml"):
        relationships.append(
            (
                "rId11",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                "../customXml/item1.xml",
            )
        )
    return relationships


def docx_document_relationships_xml(layout: dict[str, Any]) -> str:
    return relationships_xml(docx_document_relationships(layout))


def docx_hyperlink_relationship_ids(blocks: list[dict[str, Any]]) -> dict[str, str]:
    targets: list[str] = []
    for paragraph in block_paragraphs(blocks):
        runs = paragraph.get("runs")
        if not isinstance(runs, list):
            continue
        for run in runs:
            if not isinstance(run, dict):
                continue
            target = run_link_target(run)
            if target and target not in targets:
                targets.append(target)
    return {target: f"rId{101 + index}" for index, target in enumerate(targets)}


def docx_document_relationships_with_links_xml(layout: dict[str, Any], blocks: list[dict[str, Any]]) -> str:
    relationships = docx_document_relationships(layout)
    relationships.extend(
        (
            relationship_id,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            target,
            "External",
        )
        for target, relationship_id in docx_hyperlink_relationship_ids(blocks).items()
    )
    return relationships_xml(relationships)


def docx_settings_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{WORD_NS}"><w:defaultTabStop w:val="720"/></w:settings>'
    )


def document_font_names(layout: dict[str, Any]) -> list[str]:
    fonts = document_package_options(layout).get("fonts")
    if not isinstance(fonts, list):
        return ["Calibri", "Symbol"]
    names = [scalar_text(font) for font in fonts]
    return [name for index, name in enumerate(names) if name and name not in names[:index]] or ["Calibri", "Symbol"]


def docx_font_table_xml(layout: dict[str, Any]) -> str:
    fonts = "".join(f'<w:font w:name="{xml_attr(font)}"/>' for font in document_font_names(layout))
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' f'<w:fonts xmlns:w="{WORD_NS}">{fonts}</w:fonts>'


def docx_web_settings_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' f'<w:webSettings xmlns:w="{WORD_NS}"/>'


def docx_footnotes_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{WORD_NS}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/>'
        '</w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
        "<w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>"
    )


def docx_endnotes_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:endnotes xmlns:w="{WORD_NS}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/>'
        '</w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
        "<w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>"
    )


def custom_xml_item_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cv:metadata xmlns:cv="urn:x-create-cv-x"/>'


def custom_xml_props_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<ds:datastoreItem ds:itemID="{00000000-0000-0000-0000-000000000001}" '
        'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
        "<ds:schemaRefs/></ds:datastoreItem>"
    )


def custom_xml_relationships_xml() -> str:
    return relationships_xml(
        [
            (
                "rId1",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
                "itemProps1.xml",
            )
        ]
    )


def docx_theme_xml() -> str:
    fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    line_styles = "".join(
        f'<a:ln w="{width}" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        for width in ("6350", "12700", "19050")
    )
    effect_styles = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    background_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Generated Office Theme">'
        '<a:themeElements><a:clrScheme name="Generated">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1F4E79"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
        '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
        '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>'
        '</a:clrScheme><a:fontScheme name="Generated"><a:majorFont><a:latin typeface="Calibri"/>'
        '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/>'
        '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Generated">'
        f"<a:fillStyleLst>{fill_styles}</a:fillStyleLst>"
        f"<a:lnStyleLst>{line_styles}</a:lnStyleLst>"
        f"<a:effectStyleLst>{effect_styles}</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{background_fill_styles}</a:bgFillStyleLst>"
        "</a:fmtScheme></a:themeElements></a:theme>"
    )


def docx_header_footer_xml(tag: str, content: dict[str, Any]) -> str:
    paragraphs = layout_paragraphs(content) or [{"runs": [{"text": ""}]}]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:{tag} xmlns:w="{WORD_NS}">{"".join(docx_paragraph(paragraph) for paragraph in paragraphs)}</w:{tag}>'
    )


def write_resume_document(master: dict[str, Any], resume: dict[str, Any], path: Path) -> None:
    layout = document_layout(resume, path.name)
    blocks = resume_blocks(master, resume, layout)
    paragraphs = block_paragraphs(blocks)
    resume_value = resume.get("resume")
    resume_meta = resume_value if isinstance(resume_value, dict) else {}
    title = scalar_text(resume_meta.get("label") or "Resume")
    root_relationships: list[Relationship] = [
        (
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "word/document.xml",
        )
    ]
    if package_enabled(layout, "core_properties"):
        root_relationships.append(
            (
                "rId2",
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                "docProps/core.xml",
            )
        )
    if package_enabled(layout, "extended_properties"):
        root_relationships.append(
            (
                "rId3",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                "docProps/app.xml",
            )
        )
    entries = [
        ("[Content_Types].xml", document_content_types(layout)),
        ("_rels/.rels", relationships_xml(root_relationships)),
        ("word/_rels/document.xml.rels", docx_document_relationships_with_links_xml(layout, blocks)),
        ("word/document.xml", docx_document_xml(blocks, layout, docx_hyperlink_relationship_ids(blocks))),
        ("word/styles.xml", docx_styles_xml(layout)),
        ("word/numbering.xml", docx_numbering_xml(paragraphs, layout)),
        ("word/settings.xml", docx_settings_xml()),
    ]
    if package_enabled(layout, "core_properties"):
        entries.append(
            (
                "docProps/core.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" '
                'xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f"<dc:title>{xml_text(title)}</dc:title><dc:creator>x_create_cv_x</dc:creator>"
                "<cp:lastModifiedBy>x_create_cv_x</cp:lastModifiedBy>"
                '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:modified>'
                "</cp:coreProperties>",
            )
        )
    if package_enabled(layout, "extended_properties"):
        entries.append(
            (
                "docProps/app.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                "<Application>x_create_cv_x</Application></Properties>",
            )
        )
    if package_enabled(layout, "theme"):
        entries.append(("word/theme/theme1.xml", docx_theme_xml()))
    if package_enabled(layout, "font_table"):
        entries.append(("word/fontTable.xml", docx_font_table_xml(layout)))
    if package_enabled(layout, "web_settings"):
        entries.append(("word/webSettings.xml", docx_web_settings_xml()))
    if package_enabled(layout, "footnotes"):
        entries.append(("word/footnotes.xml", docx_footnotes_xml()))
    if package_enabled(layout, "endnotes"):
        entries.append(("word/endnotes.xml", docx_endnotes_xml()))
    if package_enabled(layout, "custom_xml"):
        entries.extend(
            [
                ("customXml/item1.xml", custom_xml_item_xml()),
                ("customXml/itemProps1.xml", custom_xml_props_xml()),
                ("customXml/_rels/item1.xml.rels", custom_xml_relationships_xml()),
            ]
        )
    header = layout.get("header")
    if isinstance(header, dict):
        entries.append(("word/header1.xml", docx_header_footer_xml("hdr", header)))
    footer = layout.get("footer")
    if isinstance(footer, dict):
        entries.append(("word/footer1.xml", docx_header_footer_xml("ftr", footer)))
    write_zip_entries(path, entries)


def load_golden_json(evidence_dir: Path) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    json_dir = evidence_dir / "generated" / "json" / "a_posteriori"
    master = read_json(json_dir / "master_profile_a_posteriori.json")
    validate_contract(master, "master_profile")
    resumes: dict[str, dict[str, Any]] = {}
    for json_name, relative_path in GOLDEN_JSON_EXPECTATIONS.items():
        if json_name == MASTER_FILE:
            continue
        json_path = Path(relative_path)
        resume = read_json(json_dir / json_path.name)
        validate_contract(resume, "resume")
        resumes[json_path.stem] = resume
    return master, resumes


def docx_text_lines(path: Path) -> list[str]:
    namespace = {"w": WORD_NS}
    with zipfile.ZipFile(path) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
    lines: list[str] = []
    for paragraph in document.findall(".//w:p", namespace):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", namespace))
        normalized = " ".join(text.split())
        if normalized:
            lines.append(normalized)
    return lines


def xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    namespace = {"s": SHEET_NS}
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in item.findall(".//s:t", namespace)) for item in root]


def xlsx_cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    namespace = {"s": SHEET_NS}
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//s:t", namespace))
    value = cell.find("s:v", namespace)
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (IndexError, ValueError):
            return ""
    return value.text


def xlsx_text_lines(path: Path) -> list[str]:
    namespace = {"s": SHEET_NS}
    lines: list[str] = []
    with zipfile.ZipFile(path) as archive:
        shared_strings = xlsx_shared_strings(archive)
        worksheet_names = sorted(name for name in archive.namelist() if name.startswith("xl/worksheets/sheet"))
        for worksheet_name in worksheet_names:
            root = ET.fromstring(archive.read(worksheet_name))
            for row in root.findall(".//s:row", namespace):
                values = [xlsx_cell_text(cell, shared_strings) for cell in row.findall("s:c", namespace)]
                line = "\t".join(value for value in values if value).strip()
                if line:
                    lines.append(" ".join(line.split()))
    return lines


def normalized_text_summary(path: Path) -> dict[str, Any]:
    if path.suffix.lower() == ".docx":
        lines = docx_text_lines(path)
    elif path.suffix.lower() == ".xlsx":
        lines = xlsx_text_lines(path)
    else:
        lines = []
    payload = "\n".join(lines).encode("utf-8")
    return {"line_count": len(lines), "sha256": hashlib.sha256(payload).hexdigest()}


def docx_structure_summary(path: Path) -> dict[str, Any]:
    namespace = {"w": WORD_NS, "r": REL_NS, "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
    with zipfile.ZipFile(path) as archive:
        part_names = archive.namelist()
        document = ET.fromstring(archive.read("word/document.xml"))
        relationships = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        style_definitions = ET.fromstring(archive.read("word/styles.xml")) if "word/styles.xml" in part_names else None
        font_table = ET.fromstring(archive.read("word/fontTable.xml")) if "word/fontTable.xml" in part_names else None
        numbering = ET.fromstring(archive.read("word/numbering.xml")) if "word/numbering.xml" in part_names else None
    tables = document.findall(".//w:tbl", namespace)
    paragraphs = document.findall(".//w:p", namespace)
    runs = document.findall(".//w:r", namespace)
    body = document.find("w:body", namespace)
    section = document.find(".//w:sectPr", namespace)
    page_size: dict[str, str] = {}
    page_margins: dict[str, str] = {}
    section_header_references: list[dict[str, str]] = []
    section_footer_references: list[dict[str, str]] = []
    if section is not None:
        page_size_element = section.find("w:pgSz", namespace)
        if page_size_element is not None:
            page_size = {key.rsplit("}", 1)[-1]: value for key, value in page_size_element.attrib.items()}
        page_margins_element = section.find("w:pgMar", namespace)
        if page_margins_element is not None:
            page_margins = {key.rsplit("}", 1)[-1]: value for key, value in page_margins_element.attrib.items()}
        for reference in section.findall("w:headerReference", namespace):
            section_header_references.append(
                {
                    "type": reference.attrib.get(f"{{{WORD_NS}}}type", ""),
                    "id": reference.attrib.get(f"{{{REL_NS}}}id", ""),
                }
            )
        for reference in section.findall("w:footerReference", namespace):
            section_footer_references.append(
                {
                    "type": reference.attrib.get(f"{{{WORD_NS}}}type", ""),
                    "id": reference.attrib.get(f"{{{REL_NS}}}id", ""),
                }
            )
    body_child_counts: dict[str, int] = {}
    if body is not None:
        for child in body:
            child_name = child.tag.rsplit("}", 1)[-1]
            body_child_counts[child_name] = body_child_counts.get(child_name, 0) + 1
    relationship_type_counts: dict[str, int] = {}
    relationship_target_mode_counts: dict[str, int] = {}
    relationship_count = 0
    external_relationship_count = 0
    external_hyperlink_relationship_count = 0
    for relationship in relationships.findall("rel:Relationship", namespace):
        relationship_count += 1
        relationship_type = relationship.attrib.get("Type", "").rsplit("/", 1)[-1]
        relationship_type_counts[relationship_type] = relationship_type_counts.get(relationship_type, 0) + 1
        target_mode = relationship.attrib.get("TargetMode", "")
        if target_mode:
            relationship_target_mode_counts[target_mode] = relationship_target_mode_counts.get(target_mode, 0) + 1
        if target_mode == "External":
            external_relationship_count += 1
            if relationship_type == "hyperlink":
                external_hyperlink_relationship_count += 1
    font_names: list[str] = []
    if font_table is not None:
        font_names = [
            font.attrib[f"{{{WORD_NS}}}name"]
            for font in font_table.findall("w:font", namespace)
            if f"{{{WORD_NS}}}name" in font.attrib
        ]
    run_property_counts = {
        "bold_run_count": 0,
        "italic_run_count": 0,
        "underline_run_count": 0,
        "colored_run_count": 0,
        "fonted_run_count": 0,
    }
    for run in runs:
        properties = run.find("w:rPr", namespace)
        if properties is None:
            continue
        if properties.find("w:b", namespace) is not None:
            run_property_counts["bold_run_count"] += 1
        if properties.find("w:i", namespace) is not None:
            run_property_counts["italic_run_count"] += 1
        if properties.find("w:u", namespace) is not None:
            run_property_counts["underline_run_count"] += 1
        if properties.find("w:color", namespace) is not None:
            run_property_counts["colored_run_count"] += 1
        if properties.find("w:rFonts", namespace) is not None:
            run_property_counts["fonted_run_count"] += 1
    paragraph_property_counts = {
        "styled_paragraph_count": 0,
        "aligned_paragraph_count": 0,
        "spaced_paragraph_count": 0,
        "indented_paragraph_count": 0,
        "tab_stopped_paragraph_count": 0,
    }
    for paragraph in paragraphs:
        properties = paragraph.find("w:pPr", namespace)
        if properties is None:
            continue
        if properties.find("w:pStyle", namespace) is not None:
            paragraph_property_counts["styled_paragraph_count"] += 1
        if properties.find("w:jc", namespace) is not None:
            paragraph_property_counts["aligned_paragraph_count"] += 1
        if properties.find("w:spacing", namespace) is not None:
            paragraph_property_counts["spaced_paragraph_count"] += 1
        if properties.find("w:ind", namespace) is not None:
            paragraph_property_counts["indented_paragraph_count"] += 1
        if properties.find("w:tabs", namespace) is not None:
            paragraph_property_counts["tab_stopped_paragraph_count"] += 1
    table_grid_widths: list[list[str]] = []
    table_cell_widths: list[list[list[str]]] = []
    for table in tables:
        table_grid_widths.append(
            [grid_col.attrib.get(f"{{{WORD_NS}}}w", "") for grid_col in table.findall("w:tblGrid/w:gridCol", namespace)]
        )
        table_rows: list[list[str]] = []
        for row in table.findall("w:tr", namespace):
            cell_widths: list[str] = []
            for cell in row.findall("w:tc", namespace):
                width = cell.find("w:tcPr/w:tcW", namespace)
                cell_widths.append(width.attrib.get(f"{{{WORD_NS}}}w", "") if width is not None else "")
            table_rows.append(cell_widths)
        table_cell_widths.append(table_rows)
    numbering_abstract_ids: list[str] = []
    numbering_num_ids: list[str] = []
    numbering_level_texts: set[str] = set()
    numbering_level_fonts: set[str] = set()
    numbering_level_count = 0
    if numbering is not None:
        for abstract_num in numbering.findall("w:abstractNum", namespace):
            abstract_num_id = abstract_num.attrib.get(f"{{{WORD_NS}}}abstractNumId")
            if abstract_num_id:
                numbering_abstract_ids.append(abstract_num_id)
            for level in abstract_num.findall("w:lvl", namespace):
                numbering_level_count += 1
                level_text = level.find("w:lvlText", namespace)
                if level_text is not None:
                    text_value = level_text.attrib.get(f"{{{WORD_NS}}}val")
                    if text_value:
                        numbering_level_texts.add(text_value)
                fonts = level.find("w:rPr/w:rFonts", namespace)
                if fonts is not None:
                    font_value = fonts.attrib.get(f"{{{WORD_NS}}}ascii") or fonts.attrib.get(f"{{{WORD_NS}}}hAnsi")
                    if font_value:
                        numbering_level_fonts.add(font_value)
        for num in numbering.findall("w:num", namespace):
            num_id = num.attrib.get(f"{{{WORD_NS}}}numId")
            if num_id:
                numbering_num_ids.append(num_id)
    style_definition_ids: list[str] = []
    default_style_ids: list[str] = []
    style_based_on: dict[str, str] = {}
    style_run_fonts: dict[str, str] = {}
    style_run_sizes: dict[str, str] = {}
    style_bold_ids: list[str] = []
    style_paragraph_spacing: dict[str, dict[str, str]] = {}
    style_paragraph_indents: dict[str, dict[str, str]] = {}
    style_numbering: dict[str, dict[str, str]] = {}
    if style_definitions is not None:
        for style in style_definitions.findall("w:style", namespace):
            style_id = style.attrib.get(f"{{{WORD_NS}}}styleId")
            if not style_id:
                continue
            style_definition_ids.append(style_id)
            if style.attrib.get(f"{{{WORD_NS}}}default") == "1":
                default_style_ids.append(style_id)
            based_on = style.find("w:basedOn", namespace)
            if based_on is not None:
                based_on_value = based_on.attrib.get(f"{{{WORD_NS}}}val")
                if based_on_value:
                    style_based_on[style_id] = based_on_value
            font = style.find("w:rPr/w:rFonts", namespace)
            if font is not None:
                font_value = font.attrib.get(f"{{{WORD_NS}}}ascii") or font.attrib.get(f"{{{WORD_NS}}}hAnsi")
                if font_value:
                    style_run_fonts[style_id] = font_value
            size = style.find("w:rPr/w:sz", namespace)
            if size is not None:
                size_value = size.attrib.get(f"{{{WORD_NS}}}val")
                if size_value:
                    style_run_sizes[style_id] = size_value
            if style.find("w:rPr/w:b", namespace) is not None:
                style_bold_ids.append(style_id)
            spacing = style.find("w:pPr/w:spacing", namespace)
            if spacing is not None:
                style_paragraph_spacing[style_id] = {
                    key.rsplit("}", 1)[-1]: value for key, value in spacing.attrib.items()
                }
            indent = style.find("w:pPr/w:ind", namespace)
            if indent is not None:
                style_paragraph_indents[style_id] = {
                    key.rsplit("}", 1)[-1]: value for key, value in indent.attrib.items()
                }
            style_num_id = style.find("w:pPr/w:numPr/w:numId", namespace)
            if style_num_id is not None:
                style_level = style.find("w:pPr/w:numPr/w:ilvl", namespace)
                style_numbering[style_id] = {
                    "level": style_level.attrib.get(f"{{{WORD_NS}}}val", "0") if style_level is not None else "0",
                    "num_id": style_num_id.attrib.get(f"{{{WORD_NS}}}val", ""),
                }
    styles = sorted(
        {
            style.attrib[f"{{{WORD_NS}}}val"]
            for style in document.findall(".//w:pStyle", namespace)
            if f"{{{WORD_NS}}}val" in style.attrib
        }
    )
    return {
        "part_names": sorted(part_names),
        "part_count": len(part_names),
        "has_theme": "word/theme/theme1.xml" in part_names,
        "has_font_table": "word/fontTable.xml" in part_names,
        "has_web_settings": "word/webSettings.xml" in part_names,
        "has_footnotes": "word/footnotes.xml" in part_names,
        "has_endnotes": "word/endnotes.xml" in part_names,
        "has_custom_xml": "customXml/item1.xml" in part_names,
        "header_count": len([name for name in part_names if name.startswith("word/header")]),
        "footer_count": len([name for name in part_names if name.startswith("word/footer")]),
        "page_size": page_size,
        "page_margins": page_margins,
        "section_header_references": section_header_references,
        "section_footer_references": section_footer_references,
        "body_child_counts": body_child_counts,
        "relationship_count": relationship_count,
        "relationship_type_counts": relationship_type_counts,
        "relationship_target_mode_counts": relationship_target_mode_counts,
        "external_relationship_count": external_relationship_count,
        "external_hyperlink_relationship_count": external_hyperlink_relationship_count,
        "paragraph_count": len(paragraphs),
        "run_count": len(runs),
        **paragraph_property_counts,
        **run_property_counts,
        "hyperlink_count": len(document.findall(".//w:hyperlink", namespace)),
        "tab_count": len(document.findall(".//w:tab", namespace)),
        "table_count": len(tables),
        "table_row_count": len(document.findall(".//w:tr", namespace)),
        "table_cell_count": len(document.findall(".//w:tc", namespace)),
        "table_paragraph_count": sum(len(table.findall(".//w:p", namespace)) for table in tables),
        "table_grid_widths": table_grid_widths,
        "table_cell_widths": table_cell_widths,
        "numbered_paragraph_count": len(document.findall(".//w:numPr", namespace)),
        "numbering_abstract_count": len(numbering_abstract_ids),
        "numbering_num_count": len(numbering_num_ids),
        "numbering_level_count": numbering_level_count,
        "numbering_abstract_ids": numbering_abstract_ids,
        "numbering_num_ids": numbering_num_ids,
        "numbering_level_texts": sorted(numbering_level_texts),
        "numbering_level_fonts": sorted(numbering_level_fonts),
        "style_definition_count": len(style_definition_ids),
        "style_definition_ids": style_definition_ids,
        "default_style_ids": default_style_ids,
        "style_based_on": style_based_on,
        "style_run_fonts": style_run_fonts,
        "style_run_sizes": style_run_sizes,
        "style_bold_ids": style_bold_ids,
        "style_paragraph_spacing": style_paragraph_spacing,
        "style_paragraph_indents": style_paragraph_indents,
        "style_numbering": style_numbering,
        "font_names": font_names,
        "styles": styles,
    }


def xlsx_column_index(cell_reference: str) -> int:
    column_index_value = 0
    for character in cell_reference:
        if not character.isalpha():
            break
        column_index_value = column_index_value * 26 + ord(character.upper()) - ord("A") + 1
    return column_index_value


def xlsx_relationship_targets(archive: zipfile.ZipFile) -> dict[str, str]:
    relationship_path = "xl/_rels/workbook.xml.rels"
    if relationship_path not in archive.namelist():
        return {}
    namespace = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
    root = ET.fromstring(archive.read(relationship_path))
    targets: dict[str, str] = {}
    for relationship in root.findall("rel:Relationship", namespace):
        relationship_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if relationship_id is None or target is None:
            continue
        normalized_target = target.lstrip("/")
        if not normalized_target.startswith("xl/"):
            normalized_target = f"xl/{normalized_target}"
        targets[relationship_id] = normalized_target
    return targets


def xlsx_style_summary(archive: zipfile.ZipFile) -> dict[str, int]:
    if "xl/styles.xml" not in archive.namelist():
        return {}
    namespace = {"s": SHEET_NS}
    root = ET.fromstring(archive.read("xl/styles.xml"))
    cell_xfs = root.find("s:cellXfs", namespace)
    cell_styles = root.find("s:cellStyles", namespace)
    return {
        "cell_xfs_count": len(list(cell_xfs)) if cell_xfs is not None else 0,
        "cell_style_count": len(list(cell_styles)) if cell_styles is not None else 0,
    }


def xlsx_sheet_summary(archive: zipfile.ZipFile, worksheet_path: str, shared_strings: list[str]) -> dict[str, Any]:
    namespace = {"s": SHEET_NS}
    root = ET.fromstring(archive.read(worksheet_path))
    dimension = root.find("s:dimension", namespace)
    auto_filter = root.find("s:autoFilter", namespace)
    pane = root.find(".//s:pane", namespace)
    rows = root.findall(".//s:row", namespace)
    max_column = 0
    styled_cell_count = 0
    cell_type_counts: dict[str, int] = {}
    for row in rows:
        for cell in row.findall("s:c", namespace):
            max_column = max(max_column, xlsx_column_index(cell.attrib.get("r", "")))
            if "s" in cell.attrib:
                styled_cell_count += 1
            cell_type = cell.attrib.get("t", "number")
            cell_type_counts[cell_type] = cell_type_counts.get(cell_type, 0) + 1
    header_row = rows[0] if rows else None
    headers = (
        []
        if header_row is None
        else [xlsx_cell_text(cell, shared_strings) for cell in header_row.findall("s:c", namespace)]
    )
    return {
        "path": worksheet_path,
        "dimension": dimension.attrib.get("ref", "") if dimension is not None else "",
        "row_count": len(rows),
        "column_count": max_column,
        "headers": headers,
        "styled_cell_count": styled_cell_count,
        "cell_type_counts": cell_type_counts,
        "auto_filter_ref": auto_filter.attrib.get("ref", "") if auto_filter is not None else "",
        "freeze_pane": dict(pane.attrib) if pane is not None else {},
        "column_widths": [col.attrib.get("width", "") for col in root.findall("s:cols/s:col", namespace)],
    }


def xlsx_structure_summary(path: Path) -> dict[str, Any]:
    namespace = {"s": SHEET_NS}
    with zipfile.ZipFile(path) as archive:
        relationship_targets = xlsx_relationship_targets(archive)
        shared_strings = xlsx_shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        style_summary = xlsx_style_summary(archive)
        sheet_names: list[str] = []
        sheets: list[dict[str, Any]] = []
        for index, sheet in enumerate(workbook.findall(".//s:sheet", namespace), start=1):
            sheet_name = str(sheet.attrib["name"])
            relationship_id = sheet.attrib.get(f"{{{REL_NS}}}id")
            worksheet_path = relationship_targets.get(relationship_id or "", f"xl/worksheets/sheet{index}.xml")
            sheet_names.append(sheet_name)
            if worksheet_path in archive.namelist():
                sheet_summary = xlsx_sheet_summary(archive, worksheet_path, shared_strings)
                sheet_summary["name"] = sheet_name
                sheets.append(sheet_summary)
    return {"sheet_count": len(sheet_names), "sheet_names": sheet_names, "sheets": sheets, "styles": style_summary}


def office_structure_summary(path: Path) -> dict[str, Any]:
    if path.suffix.lower() == ".docx":
        return docx_structure_summary(path)
    if path.suffix.lower() == ".xlsx":
        return xlsx_structure_summary(path)
    return {}


def office_file_summary(relative_path: str, path: Path) -> dict[str, Any]:
    summary: dict[str, Any] = {"path": relative_path, "bytes": path.stat().st_size, "sha256": sha256_file(path)}
    try:
        summary["normalized_text"] = normalized_text_summary(path)
        summary["structure"] = office_structure_summary(path)
    except (KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        summary["normalized_text"] = {"error": exc.__class__.__name__}
        summary["structure"] = {"error": exc.__class__.__name__}
    return summary


def load_audit_policy(path: Path) -> dict[str, Any]:
    policy = read_json(path)
    validate_contract(policy, "audit_policy")
    return policy


def accepted_difference_for(comparison: dict[str, Any], audit_policy: dict[str, Any]) -> dict[str, Any] | None:
    generated = comparison.get("generated")
    generated_path = str(generated.get("path", "")) if isinstance(generated, dict) else ""
    generated_path_normalized = generated_path.replace("\\", "/")
    generated_name = office_report_name(generated_path)
    generated_candidates = {generated_path, generated_path_normalized, generated_name, "*"}
    entries = audit_policy.get("accepted_differences")
    if not isinstance(entries, list):
        return None
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        policy_path = entry.get("generated_path")
        if not isinstance(policy_path, str):
            continue
        if policy_path in generated_candidates or policy_path.replace("\\", "/") in generated_candidates:
            return entry
    return None


def apply_audit_policy(comparison: dict[str, Any], audit_policy: dict[str, Any]) -> dict[str, Any]:
    classified = dict(comparison)
    if comparison.get("byte_identical") is True:
        classified["status"] = "pass"
        classified["status_reason"] = "Generated file is byte-identical to source evidence."
        return classified
    if comparison.get("normalized_text_match") is True:
        classified["status"] = "pass"
        classified["status_reason"] = "Normalized text matches source evidence."
        return classified
    accepted_difference = accepted_difference_for(comparison, audit_policy)
    if accepted_difference is not None:
        classified["status"] = "accepted_drift"
        classified["status_reason"] = scalar_text(accepted_difference.get("reason"))
        classified["accepted_difference"] = accepted_difference
        return classified
    classified["status"] = "review_required"
    classified["status_reason"] = "No policy entry accepts this generated/source difference."
    return classified


def office_comparisons(evidence_dir: Path, audit_policy: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    policy = audit_policy or load_audit_policy(DEFAULT_AUDIT_POLICY)
    comparisons: list[dict[str, Any]] = []
    for generated_name, generated_relative_path in GOLDEN_OFFICE_EXPECTATIONS.items():
        source_relative_path = GOLDEN_SOURCE_OFFICE[generated_name]
        generated_summary = office_file_summary(generated_relative_path, evidence_dir / generated_relative_path)
        source_summary = office_file_summary(source_relative_path, evidence_dir / source_relative_path)
        generated_normalized = generated_summary.get("normalized_text")
        source_normalized = source_summary.get("normalized_text")
        normalized_match = isinstance(generated_normalized, dict) and generated_normalized == source_normalized
        comparison = {
            "generated": generated_summary,
            "source": source_summary,
            "byte_identical": generated_summary["sha256"] == source_summary["sha256"],
            "normalized_text_match": normalized_match,
        }
        comparisons.append(apply_audit_policy(comparison, policy))
    return comparisons


def audit_policy_summary(audit_policy: dict[str, Any], policy_path: Path | None = None) -> dict[str, Any]:
    summary = {
        "schema_version": audit_policy.get("schema_version"),
        "policy_version": audit_policy.get("policy_version"),
        "name": audit_policy.get("name"),
        "accepted_difference_count": len(audit_policy.get("accepted_differences", [])),
    }
    if policy_path is not None:
        summary["path"] = str(policy_path).replace("\\", "/")
    return summary


def write_office_comparison_report(evidence_dir: Path, policy_path: Path = DEFAULT_AUDIT_POLICY) -> Path:
    audit_policy = load_audit_policy(policy_path)
    comparisons = office_comparisons(evidence_dir, audit_policy)
    report_path = evidence_dir / GOLDEN_OFFICE_REPORT
    write_json(
        report_path,
        {
            "schema_version": SCHEMA_VERSION,
            "generator": f"x_create_cv_x {VERSION}",
            "purpose": "Private comparison report for generated _a_posteriori Office evidence",
            "audit_policy": audit_policy_summary(audit_policy, policy_path),
            "comparisons": comparisons,
        },
    )
    return report_path


def markdown_cell(value: Any) -> str:
    return scalar_text(value).replace("\n", " ").replace("|", "\\|")


def office_report_name(relative_path: str) -> str:
    return relative_path.replace("\\", "/").rsplit("/", 1)[-1]


def office_report_suffix(relative_path: str) -> str:
    name = office_report_name(relative_path)
    if "." not in name:
        return ""
    return "." + name.rsplit(".", 1)[-1].lower()


def structure_metric(summary: dict[str, Any], key: str) -> Any:
    structure = summary.get("structure")
    if not isinstance(structure, dict):
        return ""
    return structure.get(key, "")


def normalized_metric(summary: dict[str, Any], key: str) -> Any:
    normalized = summary.get("normalized_text")
    if not isinstance(normalized, dict):
        return ""
    return normalized.get(key, "")


def append_metric_table(
    lines: list[str], generated: dict[str, Any], source: dict[str, Any], metrics: list[str]
) -> None:
    lines.extend(["| Metric | Source | Generated |", "| --- | ---: | ---: |"])
    for metric in metrics:
        lines.append(
            f"| {markdown_cell(metric)} | {markdown_cell(structure_metric(source, metric))} | "
            f"{markdown_cell(structure_metric(generated, metric))} |"
        )
    lines.append("")


def write_office_audit_report(evidence_dir: Path, policy_path: Path = DEFAULT_AUDIT_POLICY) -> Path:
    audit_policy = load_audit_policy(policy_path)
    comparisons = office_comparisons(evidence_dir, audit_policy)
    lines = [
        "# A Posteriori Office Audit",
        "",
        f"Generator: `x_create_cv_x {VERSION}`",
        f"Audit policy: `{audit_policy.get('name')}` version `{audit_policy.get('policy_version')}`",
        "",
        "## Summary",
        "",
        "| File | Type | Byte Identical | Normalized Text Match | Status | Reason | Source Lines | Generated Lines |",
        "| --- | --- | ---: | ---: | --- | --- | ---: | ---: |",
    ]
    for comparison in comparisons:
        generated = comparison["generated"]
        source = comparison["source"]
        generated_relative_path = str(generated["path"])
        lines.append(
            f"| {markdown_cell(office_report_name(generated_relative_path))} | "
            f"{markdown_cell(office_report_suffix(generated_relative_path))} | "
            f"{markdown_cell(comparison['byte_identical'])} | {markdown_cell(comparison['normalized_text_match'])} | "
            f"{markdown_cell(comparison['status'])} | {markdown_cell(comparison.get('status_reason', ''))} | "
            f"{markdown_cell(normalized_metric(source, 'line_count'))} | "
            f"{markdown_cell(normalized_metric(generated, 'line_count'))} |"
        )

    lines.extend(["", "## DOCX Structure", ""])
    docx_metrics = [
        "part_count",
        "relationship_count",
        "external_relationship_count",
        "external_hyperlink_relationship_count",
        "page_size",
        "page_margins",
        "section_header_references",
        "section_footer_references",
        "paragraph_count",
        "run_count",
        "styled_paragraph_count",
        "aligned_paragraph_count",
        "spaced_paragraph_count",
        "indented_paragraph_count",
        "tab_stopped_paragraph_count",
        "bold_run_count",
        "italic_run_count",
        "underline_run_count",
        "colored_run_count",
        "fonted_run_count",
        "hyperlink_count",
        "tab_count",
        "table_count",
        "table_row_count",
        "table_cell_count",
        "table_paragraph_count",
        "table_grid_widths",
        "table_cell_widths",
        "numbered_paragraph_count",
        "numbering_abstract_count",
        "numbering_num_count",
        "numbering_level_count",
        "numbering_abstract_ids",
        "numbering_num_ids",
        "numbering_level_texts",
        "numbering_level_fonts",
        "style_definition_count",
        "style_definition_ids",
        "default_style_ids",
        "style_based_on",
        "style_run_fonts",
        "style_run_sizes",
        "style_bold_ids",
        "style_paragraph_spacing",
        "style_paragraph_indents",
        "style_numbering",
    ]
    for comparison in comparisons:
        generated = comparison["generated"]
        generated_relative_path = str(generated["path"])
        if office_report_suffix(generated_relative_path) != ".docx":
            continue
        lines.extend([f"### {office_report_name(generated_relative_path)}", ""])
        append_metric_table(lines, generated, comparison["source"], docx_metrics)

    lines.extend(["## XLSX Structure", ""])
    for comparison in comparisons:
        generated = comparison["generated"]
        source = comparison["source"]
        generated_relative_path = str(generated["path"])
        if office_report_suffix(generated_relative_path) != ".xlsx":
            continue
        lines.extend([f"### {office_report_name(generated_relative_path)}", ""])
        append_metric_table(lines, generated, source, ["sheet_count"])
        generated_sheets = structure_metric(generated, "sheets")
        source_sheets = structure_metric(source, "sheets")
        if isinstance(generated_sheets, list) and isinstance(source_sheets, list):
            lines.extend(
                [
                    "| Sheet | Source Dimension | Generated Dimension | Generated Freeze Pane | "
                    "Generated Filter | Generated Cell Types | Source Headers | Generated Headers |",
                    "| --- | --- | --- | --- | --- | --- | --- | --- |",
                ]
            )
            for index, generated_sheet in enumerate(generated_sheets):
                if not isinstance(generated_sheet, dict):
                    continue
                source_sheet = source_sheets[index] if index < len(source_sheets) else {}
                if not isinstance(source_sheet, dict):
                    source_sheet = {}
                lines.append(
                    f"| {markdown_cell(generated_sheet.get('name', ''))} | "
                    f"{markdown_cell(source_sheet.get('dimension', ''))} | "
                    f"{markdown_cell(generated_sheet.get('dimension', ''))} | "
                    f"{markdown_cell(generated_sheet.get('freeze_pane', {}))} | "
                    f"{markdown_cell(generated_sheet.get('auto_filter_ref', ''))} | "
                    f"{markdown_cell(generated_sheet.get('cell_type_counts', {}))} | "
                    f"{markdown_cell(source_sheet.get('headers', []))} | "
                    f"{markdown_cell(generated_sheet.get('headers', []))} |"
                )
            lines.append("")

    lines.extend(
        [
            "## Known Acceptable Differences",
            "",
        ]
    )
    accepted_differences = audit_policy.get("accepted_differences")
    if isinstance(accepted_differences, list) and accepted_differences:
        lines.extend(["| Generated Path | Scope | Reason |", "| --- | --- | --- |"])
        for entry in accepted_differences:
            if not isinstance(entry, dict):
                continue
            lines.append(
                f"| {markdown_cell(entry.get('generated_path', ''))} | {markdown_cell(entry.get('scope', ''))} | "
                f"{markdown_cell(entry.get('reason', ''))} |"
            )
    else:
        lines.append(
            "The active policy contains no accepted-drift entries; unmatched differences remain review-required."
        )
    lines.append("")
    report_path = evidence_dir / GOLDEN_OFFICE_AUDIT_REPORT
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


def write_golden_office_outputs(
    evidence_dir: Path, output_root: Path | None = None, *, write_report: bool = True
) -> tuple[int, int]:
    master, resumes = load_golden_json(evidence_dir)
    target_root = output_root or evidence_dir
    output_dir = target_root / "generated" / "office" / "a_posteriori"
    write_master_workbook(master, output_dir / "master_profile_a_posteriori.xlsx")
    for office_name, relative_path in GOLDEN_OFFICE_EXPECTATIONS.items():
        if office_name == "master_profile_a_posteriori.xlsx":
            continue
        resume = resumes[office_name.removesuffix(".docx")]
        write_resume_document(master, resume, output_dir / Path(relative_path).name)
    report_count = 0
    if write_report:
        write_office_comparison_report(evidence_dir)
        report_count = 1
    return len(GOLDEN_OFFICE_EXPECTATIONS), report_count


def run_golden_office_generation(evidence_dir: Path) -> int:
    with tempfile.TemporaryDirectory(prefix="x_create_cv_office_") as temp_dir:
        temp_root = Path(temp_dir)
        generated_count, _report_count = write_golden_office_outputs(evidence_dir, temp_root, write_report=False)
        for generated_relative_path in GOLDEN_OFFICE_EXPECTATIONS.values():
            generated_path = temp_root / generated_relative_path
            expected_path = evidence_dir / generated_relative_path
            if not expected_path.exists():
                raise ValueError(f"Missing expected a posteriori Office evidence: {generated_relative_path}")
            if generated_path.read_bytes() != expected_path.read_bytes():
                raise ValueError(f"Golden generated Office mismatch: {generated_relative_path}")
        return generated_count


def run_golden_json_scripts(evidence_dir: Path) -> int:
    script_dir = evidence_dir / "scripts"
    public_repo_root = Path(__file__).resolve().parent
    with tempfile.TemporaryDirectory(prefix="x_create_cv_golden_") as temp_dir:
        output_dir = Path(temp_dir)
        env = os.environ.copy()
        env["CV_FACTORY_DB_DIR"] = str(output_dir)
        env["PYTHONPATH"] = str(public_repo_root) + os.pathsep + env.get("PYTHONPATH", "")

        for script_name in GOLDEN_SCRIPT_FILES:
            script_path = script_dir / script_name
            if not script_path.exists():
                raise ValueError(f"Missing golden script: {script_path}")
            completed = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=public_repo_root,
                env=env,
                capture_output=True,
                text=True,
                check=False,
            )
            if completed.returncode != 0:
                raise ValueError(f"Golden script failed: {script_path.name}")

        for generated_name, expected_relative_path in GOLDEN_JSON_EXPECTATIONS.items():
            generated_path = output_dir / generated_name
            expected_path = evidence_dir / expected_relative_path
            if not generated_path.exists():
                raise ValueError(f"Missing generated JSON from golden scripts: {generated_name}")
            if not expected_path.exists():
                raise ValueError(f"Missing expected a posteriori JSON evidence: {expected_relative_path}")
            schema_name = "master_profile" if generated_name == MASTER_FILE else "resume"
            validate_contract_file(generated_path, schema_name)
            validate_contract_file(expected_path, schema_name)
            if generated_path.read_bytes() != expected_path.read_bytes():
                raise ValueError(f"Golden generated JSON mismatch: {generated_name}")

    return len(GOLDEN_JSON_EXPECTATIONS)


def check_golden_evidence(evidence_dir: Path) -> tuple[int, int, int, int]:
    a_priori_count = check_evidence_manifest(evidence_dir / "a_priori_manifest.json")
    chain_count = check_evidence_manifest(evidence_dir / "chain_of_evidence_manifest.json")
    generated_json_count = run_golden_json_scripts(evidence_dir)
    generated_office_count = run_golden_office_generation(evidence_dir)
    return a_priori_count, chain_count, generated_json_count, generated_office_count


def master_path(db_dir: Path) -> Path:
    return db_dir / MASTER_FILE


def resume_path(db_dir: Path, resume_id: str) -> Path:
    return db_dir / f"{resume_id}.json"


def empty_master(profile: dict[str, Any], workbook_layout_value: dict[str, Any] | None = None) -> dict[str, Any]:
    master = {
        "schema_version": SCHEMA_VERSION,
        "app_model": MASTER_APP_MODEL,
        "profile": profile,
        "collections": {collection: [] for collection in MASTER_COLLECTIONS},
    }
    if workbook_layout_value is not None:
        master["office_layout"] = {"workbook": workbook_layout_value}
    return master


def empty_resume(resume: dict[str, Any], document_layout_value: dict[str, Any] | None = None) -> dict[str, Any]:
    resume_data = {
        "schema_version": SCHEMA_VERSION,
        "app_model": RESUME_APP_MODEL,
        "resume": resume,
        "collections": {
            "sections": [],
            "items": [],
        },
    }
    if document_layout_value is not None:
        resume_data["office_layout"] = {"document": document_layout_value}
    return resume_data


def require_record_id(record: dict[str, Any]) -> str:
    record_id = record.get("id")
    if not isinstance(record_id, str) or not record_id:
        raise ValueError("Record JSON must contain a non-empty string id")
    return record_id


def append_unique(records: list[Any], record: dict[str, Any], replace: bool) -> None:
    record_id = require_record_id(record)
    for index, existing in enumerate(records):
        if isinstance(existing, dict) and existing.get("id") == record_id:
            if not replace:
                raise ValueError(f"Record already exists: {record_id}")
            records[index] = record
            return
    records.append(record)


def init_database(
    db_dir: Path, profile: dict[str, Any], force: bool, workbook_layout_value: dict[str, Any] | None = None
) -> None:
    if db_dir.exists() and force:
        shutil.rmtree(db_dir)
    if master_path(db_dir).exists():
        raise ValueError(f"Database already exists at {db_dir}; pass --force to replace it")
    master = empty_master(profile, workbook_layout_value)
    validate_contract(master, "master_profile")
    write_json(master_path(db_dir), master)


def set_profile(db_dir: Path, profile: dict[str, Any]) -> None:
    master = read_json(master_path(db_dir))
    master["profile"] = profile
    validate_contract(master, "master_profile")
    write_json(master_path(db_dir), master)


def add_master_record(db_dir: Path, collection: str, record: dict[str, Any], replace: bool) -> None:
    master = read_json(master_path(db_dir))
    collections = master.get("collections")
    if not isinstance(collections, dict):
        raise ValueError("Master profile collections are missing or invalid")
    records = collections.get(collection)
    if not isinstance(records, list):
        raise ValueError(f"Unknown master collection: {collection}")
    append_unique(records, record, replace)
    validate_contract(master, "master_profile")
    write_json(master_path(db_dir), master)


def add_resume(
    db_dir: Path, resume: dict[str, Any], replace: bool, document_layout_value: dict[str, Any] | None = None
) -> None:
    resume_id = resume.get("id")
    if not isinstance(resume_id, str) or not resume_id:
        raise ValueError("Resume JSON must contain a non-empty string id")
    path = resume_path(db_dir, resume_id)
    if path.exists() and not replace:
        raise ValueError(f"Resume already exists: {resume_id}")
    resume_data = empty_resume(resume, document_layout_value)
    validate_contract(resume_data, "resume")
    write_json(path, resume_data)


def add_resume_record(db_dir: Path, resume_id: str, collection: str, record: dict[str, Any], replace: bool) -> None:
    resume = read_json(resume_path(db_dir, resume_id))
    collections = resume.get("collections")
    if not isinstance(collections, dict):
        raise ValueError(f"Resume collections are missing or invalid: {resume_id}")
    records = collections.get(collection)
    if not isinstance(records, list):
        raise ValueError(f"Unknown resume collection: {collection}")
    append_unique(records, record, replace)
    validate_contract(resume, "resume")
    write_json(resume_path(db_dir, resume_id), resume)


def ordered_record(field_order: list[str], values: dict[str, Any]) -> dict[str, Any]:
    return {field: values[field] for field in field_order if field in values and values[field] is not None}


def add_ordered_master_record(db_dir: Path, collection: str, values: dict[str, Any], replace: bool) -> str:
    record = ordered_record(MASTER_FIELD_ORDER[collection], values)
    add_master_record(db_dir, collection, record, replace)
    return str(record["id"])


def init_master_profile(
    db_dir: Path,
    *,
    profile_id: str,
    person_id: str,
    label: str,
    created_at: str,
    updated_at: str,
    workbook_layout: dict[str, Any] | None = None,
    force: bool = False,
) -> None:
    init_database(
        db_dir,
        ordered_record(
            ["id", "person_id", "label", "created_at", "updated_at"],
            {
                "id": profile_id,
                "person_id": person_id,
                "label": label,
                "created_at": created_at,
                "updated_at": updated_at,
            },
        ),
        force,
        workbook_layout,
    )


def add_user(db_dir: Path, *, record_id: str, display_name: str, preferred_name: str, replace: bool = False) -> str:
    return add_ordered_master_record(
        db_dir, "people", {"id": record_id, "display_name": display_name, "preferred_name": preferred_name}, replace
    )


def add_contact_method(
    db_dir: Path,
    *,
    record_id: str,
    person_id: str,
    kind: str,
    value: str,
    is_primary: bool,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(
        db_dir,
        "contact_methods",
        {"id": record_id, "person_id": person_id, "kind": kind, "value": value, "is_primary": is_primary},
        replace,
    )


def add_address(db_dir: Path, *, record_id: str, label: str, address_text: str, replace: bool = False) -> str:
    return add_ordered_master_record(
        db_dir, "addresses", {"id": record_id, "label": label, "address_text": address_text}, replace
    )


def add_job(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    d: str | None = None,
    e: str | None = None,
    f: str | None = None,
    g: str | None = None,
    h: str | None = None,
    i: str | None = None,
    j: str | None = None,
    k: str | None = None,
    job_title: str | None = None,
    job_focus: str | None = None,
    job_highlights: str | None = None,
    vehicle_program: str | None = None,
    employer: str | None = None,
    address: str | None = None,
    reason_for_leaving: str | None = None,
    salary_usd: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    time_spent: str | None = None,
    address_id: str | None = None,
    project_ids: list[str] | None = None,
    achievement_ids: list[str] | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "jobs", locals() | {"id": record_id}, replace)


def add_achievement(db_dir: Path, *, record_id: str, text: str, job_id: str, replace: bool = False) -> str:
    return add_ordered_master_record(db_dir, "achievements", {"id": record_id, "text": text, "job_id": job_id}, replace)


def add_project(db_dir: Path, *, record_id: str, name: str, replace: bool = False) -> str:
    return add_ordered_master_record(db_dir, "projects", {"id": record_id, "name": name}, replace)


def add_highlight(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    highlights: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "highlights", locals() | {"id": record_id}, replace)


def add_standard(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    d: str | None = None,
    e: str | None = None,
    f: str | None = None,
    organization: str | None = None,
    title_of_standard: str | None = None,
    role: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    time_spent: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "standards", locals() | {"id": record_id}, replace)


def add_education(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    d: str | None = None,
    e: str | None = None,
    f: str | None = None,
    school: str | None = None,
    address: str | None = None,
    degree: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    time_spent: str | None = None,
    address_id: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "education", locals() | {"id": record_id}, replace)


def add_certification(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    certificate_subject: str | None = None,
    certificate_name: str | None = None,
    certificate_number: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "certifications", locals() | {"id": record_id}, replace)


def add_patent(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    patent_name: str | None = None,
    patent_number: str | None = None,
    patent_url: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "patents", locals() | {"id": record_id}, replace)


def add_publication(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    d: str | None = None,
    e: str | None = None,
    year: str | None = None,
    publication_name: str | None = None,
    publisher: str | None = None,
    publication_reference: str | None = None,
    publication_url: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "publications", locals() | {"id": record_id}, replace)


def add_lecture(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    year: str | None = None,
    publication_or_lecture_name: str | None = None,
    publisher_name: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "lectures", locals() | {"id": record_id}, replace)


def add_residence(
    db_dir: Path,
    *,
    record_id: str,
    label: str,
    a: str | None = None,
    b: str | None = None,
    c: str | None = None,
    d: str | None = None,
    residence: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    time_spent: str | None = None,
    address_id: str | None = None,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(db_dir, "residences", locals() | {"id": record_id}, replace)


def add_skill(db_dir: Path, *, record_id: str, name: str, category: str, replace: bool = False) -> str:
    return add_ordered_master_record(db_dir, "skills", {"id": record_id, "name": name, "category": category}, replace)


def add_text_block(
    db_dir: Path,
    *,
    record_id: str,
    text: str,
    section_label: str,
    used_in_resume_ids: list[str],
    is_job_related: bool,
    replace: bool = False,
) -> str:
    return add_ordered_master_record(
        db_dir,
        "text_blocks",
        {
            "id": record_id,
            "text": text,
            "section_label": section_label,
            "used_in_resume_ids": used_in_resume_ids,
            "is_job_related": is_job_related,
        },
        replace,
    )


def create_resume(
    db_dir: Path,
    *,
    resume_id: str,
    label: str,
    status: str,
    created_at: str,
    updated_at: str,
    document_layout: dict[str, Any] | None = None,
    replace: bool = False,
) -> str:
    resume = ordered_record(
        RESUME_FIELD_ORDER,
        {"id": resume_id, "label": label, "status": status, "created_at": created_at, "updated_at": updated_at},
    )
    add_resume(db_dir, resume, replace, document_layout)
    return resume_id


def add_resume_section(
    db_dir: Path,
    *,
    resume_id: str,
    section_id: str,
    title: str,
    kind: str,
    sort_order: int,
    item_ids: list[str],
    is_visible: bool,
    replace: bool = False,
) -> str:
    section = ordered_record(
        SECTION_FIELD_ORDER,
        {
            "id": section_id,
            "title": title,
            "kind": kind,
            "sort_order": sort_order,
            "item_ids": item_ids,
            "is_visible": is_visible,
        },
    )
    add_resume_record(db_dir, resume_id, "sections", section, replace)
    return section_id


def add_resume_item(
    db_dir: Path,
    *,
    resume_id: str,
    item_id: str,
    section_id: str,
    kind: str,
    sort_order: int,
    master_record_refs: list[str],
    text_override: str,
    formatting: dict[str, Any],
    is_visible: bool,
    replace: bool = False,
) -> str:
    item = ordered_record(
        ITEM_FIELD_ORDER,
        {
            "id": item_id,
            "section_id": section_id,
            "kind": kind,
            "sort_order": sort_order,
            "master_record_refs": master_record_refs,
            "text_override": text_override,
            "formatting": formatting,
            "is_visible": is_visible,
        },
    )
    add_resume_record(db_dir, resume_id, "items", item, replace)
    return item_id


def compare_zip_entry(zip_file: zipfile.ZipFile, entry_name: str, actual_path: Path) -> str | None:
    try:
        expected = zip_file.read(entry_name)
    except KeyError:
        return f"missing expected zip entry {entry_name}"
    if not actual_path.exists():
        return f"missing generated file {actual_path}"
    actual = actual_path.read_bytes()
    if actual != expected:
        return f"mismatch {actual_path.name}: expected {len(expected)} bytes, got {len(actual)} bytes"
    return None


def validate_against_zip(db_dir: Path, expected_zip: Path, zip_prefix: str) -> None:
    prefix = zip_prefix.strip("/")
    failures: list[str] = []
    with zipfile.ZipFile(expected_zip) as zip_file:
        for filename in EXPECTED_FILES:
            entry_name = f"{prefix}/{filename}" if prefix else filename
            failure = compare_zip_entry(zip_file, entry_name, db_dir / filename)
            if failure is not None:
                failures.append(failure)
    if failures:
        raise ValueError("Validation failed:\n" + "\n".join(failures))


def add_payload_options(parser: argparse.ArgumentParser, payload_label: str) -> None:
    parser.add_argument("--json", dest="json_text", help=f"{payload_label} JSON object")
    parser.add_argument("--json-base64", dest="json_base64", help=f"Base64-encoded {payload_label} JSON object")


def add_common_db_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--db-dir", type=Path, default=DEFAULT_DB_DIR, help="Directory containing generated private CV JSON"
    )


def add_replace_option(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--replace", action="store_true", help="Replace an existing record with the same id")


def parse_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    raise argparse.ArgumentTypeError(f"Expected a boolean value, got {value!r}")


def parse_object(value: str) -> dict[str, Any]:
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError as exc:
        raise argparse.ArgumentTypeError(f"Invalid JSON object: {exc}") from exc
    if not isinstance(parsed, dict):
        raise argparse.ArgumentTypeError("Expected a JSON object")
    return parsed


def list_flag(field: str) -> str:
    if field.endswith("_ids"):
        return "--" + field.removesuffix("s").replace("_", "-")
    if field.endswith("_refs"):
        return "--" + field.removesuffix("s").replace("_", "-")
    return "--" + field.replace("_", "-")


def add_field_options(parser: argparse.ArgumentParser, field_order: list[str]) -> None:
    for field in field_order:
        destination = f"field_{field}"
        if field in LIST_FIELDS:
            parser.add_argument(list_flag(field), dest=destination, action="append", help=f"Add one {field} value")
        elif field in BOOL_FIELDS:
            parser.add_argument("--" + field.replace("_", "-"), dest=destination, type=parse_bool, help=f"Set {field}")
        elif field in DICT_FIELDS:
            parser.add_argument(
                "--" + field.replace("_", "-") + "-json",
                dest=destination,
                type=parse_object,
                help=f"Set {field} from a JSON object",
            )
        elif field == "sort_order":
            parser.add_argument("--sort-order", dest=destination, type=int, help="Set sort_order")
        else:
            parser.add_argument("--" + field.replace("_", "-"), dest=destination, help=f"Set {field}")


def typed_payload(args: argparse.Namespace, field_order: list[str], payload_name: str) -> dict[str, Any]:
    if getattr(args, "json_text", None) or getattr(args, "json_base64", None):
        return parse_json_payload(args.json_text, args.json_base64, payload_name)
    payload = {
        field: getattr(args, f"field_{field}")
        for field in field_order
        if getattr(args, f"field_{field}", None) is not None
    }
    if "id" not in payload:
        raise ValueError(f"{payload_name} requires --id when --json is not used")
    return payload


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create app-native private CV CRUD JSON from CLI commands.")
    parser.add_argument("--version", action="version", version=f"x_create_cv_x {VERSION}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    init_parser = subparsers.add_parser("init", help="Create an empty master profile database")
    add_common_db_options(init_parser)
    add_payload_options(init_parser, "profile")
    add_field_options(init_parser, ["id", "person_id", "label", "created_at", "updated_at"])
    init_parser.add_argument("--workbook-layout-json", type=parse_object, help="Set workbook office_layout from JSON")
    init_parser.add_argument("--force", action="store_true", help="Delete and recreate the target database directory")

    profile_parser = subparsers.add_parser("set-profile", help="Replace the master profile metadata")
    add_common_db_options(profile_parser)
    add_payload_options(profile_parser, "profile")
    add_field_options(profile_parser, ["id", "person_id", "label", "created_at", "updated_at"])

    for command_name, collection_name in MASTER_COMMANDS.items():
        record_parser = subparsers.add_parser(command_name, help=f"Add a record via {command_name}")
        add_common_db_options(record_parser)
        add_payload_options(record_parser, "record")
        add_field_options(record_parser, MASTER_FIELD_ORDER[collection_name])
        add_replace_option(record_parser)

    add_resume_parser = subparsers.add_parser("add-resume", help="Create a resume document")
    add_common_db_options(add_resume_parser)
    add_payload_options(add_resume_parser, "resume")
    add_field_options(add_resume_parser, RESUME_FIELD_ORDER)
    add_resume_parser.add_argument(
        "--document-layout-json", type=parse_object, help="Set document office_layout from JSON"
    )
    add_replace_option(add_resume_parser)

    create_resume_parser = subparsers.add_parser("create-resume", help="Create a resume document")
    add_common_db_options(create_resume_parser)
    add_payload_options(create_resume_parser, "resume")
    add_field_options(create_resume_parser, RESUME_FIELD_ORDER)
    create_resume_parser.add_argument(
        "--document-layout-json", type=parse_object, help="Set document office_layout from JSON"
    )
    add_replace_option(create_resume_parser)

    add_section_parser = subparsers.add_parser("add-section", help="Add a section to a resume")
    add_common_db_options(add_section_parser)
    add_section_parser.add_argument("--resume-id", required=True)
    add_payload_options(add_section_parser, "section")
    add_field_options(add_section_parser, SECTION_FIELD_ORDER)
    add_replace_option(add_section_parser)

    add_resume_section_parser = subparsers.add_parser("add-resume-section", help="Add a section to a resume")
    add_common_db_options(add_resume_section_parser)
    add_resume_section_parser.add_argument("--resume-id", required=True)
    add_payload_options(add_resume_section_parser, "section")
    add_field_options(add_resume_section_parser, SECTION_FIELD_ORDER)
    add_replace_option(add_resume_section_parser)

    add_item_parser = subparsers.add_parser("add-item", help="Add an item to a resume")
    add_common_db_options(add_item_parser)
    add_item_parser.add_argument("--resume-id", required=True)
    add_payload_options(add_item_parser, "item")
    add_field_options(add_item_parser, ITEM_FIELD_ORDER)
    add_replace_option(add_item_parser)

    add_resume_item_parser = subparsers.add_parser("add-resume-item", help="Add an item to a resume")
    add_common_db_options(add_resume_item_parser)
    add_resume_item_parser.add_argument("--resume-id", required=True)
    add_payload_options(add_resume_item_parser, "item")
    add_field_options(add_resume_item_parser, ITEM_FIELD_ORDER)
    add_replace_option(add_resume_item_parser)

    validate_parser = subparsers.add_parser("validate", help="Compare generated JSON against a private zip reference")
    add_common_db_options(validate_parser)
    validate_parser.add_argument("--expected-zip", type=Path, required=True)
    validate_parser.add_argument("--zip-prefix", default="private/cv_app_crud")

    schema_parser = subparsers.add_parser("validate-schema", help="Validate JSON files against public contracts")
    schema_parser.add_argument(
        "--schema", choices=sorted(CONTRACT_SCHEMA_FILES), help="Schema name; inferred if omitted"
    )
    schema_parser.add_argument("paths", nargs="+", type=Path, help="JSON contract files to validate")

    check_evidence_parser = subparsers.add_parser(
        "check-evidence", help="Fast-fail if private a priori Office evidence has changed"
    )
    check_evidence_parser.add_argument("--manifest", type=Path, default=DEFAULT_EVIDENCE_MANIFEST)

    exercise_golden_parser = subparsers.add_parser(
        "exercise-golden", help="Run the side-by-side private golden evidence smoke test"
    )
    exercise_golden_parser.add_argument("--evidence-dir", type=Path, default=DEFAULT_GOLDEN_EVIDENCE_DIR)

    generate_office_parser = subparsers.add_parser(
        "generate-golden-office", help="Generate private _a_posteriori Office evidence from golden JSON"
    )
    generate_office_parser.add_argument("--evidence-dir", type=Path, default=DEFAULT_GOLDEN_EVIDENCE_DIR)
    generate_office_parser.add_argument("--no-report", action="store_true", help="Skip the comparison report")

    audit_parser = subparsers.add_parser("audit", help="Write private-safe Office parity audit reports")
    audit_parser.add_argument("--evidence-dir", type=Path, default=DEFAULT_GOLDEN_EVIDENCE_DIR)
    audit_parser.add_argument("--policy", type=Path, default=DEFAULT_AUDIT_POLICY, help="Office audit policy JSON")
    audit_parser.add_argument("--no-json", action="store_true", help="Skip the machine-readable JSON report")
    return parser


def run(args: argparse.Namespace) -> int:
    command = args.command
    if command == "init":
        profile = typed_payload(args, ["id", "person_id", "label", "created_at", "updated_at"], "profile")
        init_database(args.db_dir, profile, args.force, args.workbook_layout_json)
        return 0
    if command == "set-profile":
        set_profile(
            args.db_dir, typed_payload(args, ["id", "person_id", "label", "created_at", "updated_at"], "profile")
        )
        return 0
    if command in MASTER_COMMANDS:
        collection = MASTER_COMMANDS[command]
        record = typed_payload(args, MASTER_FIELD_ORDER[collection], "record")
        add_master_record(args.db_dir, collection, record, args.replace)
        return 0
    if command in {"add-resume", "create-resume"}:
        resume = typed_payload(args, RESUME_FIELD_ORDER, "resume")
        add_resume(args.db_dir, resume, args.replace, args.document_layout_json)
        return 0
    if command in {"add-section", "add-resume-section"}:
        section = typed_payload(args, SECTION_FIELD_ORDER, "section")
        add_resume_record(args.db_dir, args.resume_id, "sections", section, args.replace)
        return 0
    if command in {"add-item", "add-resume-item"}:
        item = typed_payload(args, ITEM_FIELD_ORDER, "item")
        add_resume_record(args.db_dir, args.resume_id, "items", item, args.replace)
        return 0
    if command == "validate":
        validate_against_zip(args.db_dir, args.expected_zip, args.zip_prefix)
        print("Validation passed")
        return 0
    if command == "validate-schema":
        schema_counts: dict[str, int] = {}
        for path in args.paths:
            schema_name = validate_contract_file(path, args.schema)
            schema_counts[schema_name] = schema_counts.get(schema_name, 0) + 1
        count_summary = ", ".join(f"{count} {schema_name}" for schema_name, count in sorted(schema_counts.items()))
        print(f"Schema validation passed: {count_summary}")
        return 0
    if command == "check-evidence":
        checked_count = check_evidence_manifest(args.manifest)
        print(f"Evidence integrity passed: {checked_count} files")
        return 0
    if command == "exercise-golden":
        a_priori_count, chain_count, generated_json_count, generated_office_count = check_golden_evidence(
            args.evidence_dir
        )
        print(
            f"Golden exercise passed: {a_priori_count} a priori evidence files; "
            f"{chain_count} chain files; {generated_json_count} generated JSON files; "
            f"{generated_office_count} generated Office files"
        )
        return 0
    if command == "generate-golden-office":
        generated_office_count, report_count = write_golden_office_outputs(
            args.evidence_dir, write_report=not args.no_report
        )
        print(f"Generated golden Office evidence: {generated_office_count} files; {report_count} reports")
        return 0
    if command == "audit":
        json_report = None if args.no_json else write_office_comparison_report(args.evidence_dir, args.policy)
        markdown_report = write_office_audit_report(args.evidence_dir, args.policy)
        reports = [str(markdown_report)]
        if json_report is not None:
            reports.append(str(json_report))
        print("Office audit written: " + "; ".join(reports))
        return 0
    raise ValueError(f"Unknown command: {command}")


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return run(args)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
