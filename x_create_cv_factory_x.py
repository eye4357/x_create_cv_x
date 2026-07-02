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
from pathlib import Path
from typing import Any

MASTER_FILE = "master_profile.json"
DEFAULT_DB_DIR = Path("data/private/cv_factory_output")
RESUME_APP_MODEL = "resume_document_crud"
MASTER_APP_MODEL = "master_profile_crud"
SCHEMA_VERSION = "1.0.0"
VERSION = "0.0.2"
DEFAULT_GOLDEN_EVIDENCE_DIR = Path("../x_create_cv_test_data_x/evidence")
DEFAULT_EVIDENCE_MANIFEST = DEFAULT_GOLDEN_EVIDENCE_DIR / "a_priori_manifest.json"
DEFAULT_CHAIN_MANIFEST = DEFAULT_GOLDEN_EVIDENCE_DIR / "chain_of_evidence_manifest.json"
HASH_CHUNK_SIZE = 1024 * 1024
GOLDEN_SCRIPT_FILES = [
    "01_build_master_data.py",
    "02_build_old_resume.py",
    "03_build_new_resume.py",
]
GOLDEN_JSON_EXPECTATIONS = {
    MASTER_FILE: "generated/json/a_posteriori/master_profile_a_posteriori.json",
    "resume_2017.json": "generated/json/a_posteriori/resume_2017_a_posteriori.json",
    "resume_2023.json": "generated/json/a_posteriori/resume_2023_a_posteriori.json",
}
GOLDEN_OFFICE_EXPECTATIONS = {
    "master_profile_a_posteriori.xlsx": "generated/office/a_posteriori/master_profile_a_posteriori.xlsx",
    "resume_2017_a_posteriori.docx": "generated/office/a_posteriori/resume_2017_a_posteriori.docx",
    "resume_2023_a_posteriori.docx": "generated/office/a_posteriori/resume_2023_a_posteriori.docx",
}
GOLDEN_SOURCE_OFFICE = {
    "master_profile_a_posteriori.xlsx": "source_office/a_priori/R_cv_2023_0501_1427_a_priori.xlsx",
    "resume_2017_a_posteriori.docx": "source_office/a_priori/R_cv_2017_1129_0848_a_priori.docx",
    "resume_2023_a_posteriori.docx": "source_office/a_priori/R_cv_2023_0315_2158_a_priori.docx",
}
GOLDEN_OFFICE_REPORT = "reports/a_posteriori_office_comparison.json"
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


def relationships_xml(relationships: list[tuple[str, str, str]]) -> str:
    entries = "".join(
        f'<Relationship Id="{xml_attr(relationship_id)}" Type="{xml_attr(relationship_type)}" '
        f'Target="{xml_attr(target)}"/>'
        for relationship_id, relationship_type, target in relationships
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
                "columns": [{"field": field, "header": header_label(field)} for field in fields],
            }
            for sheet_name, collection_name, fields in DEFAULT_WORKBOOK_SHEETS
        ],
    }


def default_document_layout(file_name: str) -> dict[str, Any]:
    layout = DEFAULT_DOCUMENT_MARGINS.get(file_name, DEFAULT_DOCUMENT_MARGINS["resume_2017_a_posteriori.docx"])
    return {
        "margins": dict(layout["margins"]),
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


def workbook_sheet_columns(sheet: dict[str, Any]) -> list[tuple[str, str]]:
    columns = sheet.get("columns")
    column_defs: list[tuple[str, str]] = []
    if isinstance(columns, list):
        for column in columns:
            if not isinstance(column, dict):
                continue
            field = column.get("field")
            if not isinstance(field, str) or not field:
                continue
            header = scalar_text(column.get("header")) or header_label(field)
            column_defs.append((field, header))
    return column_defs


def sheet_rows(master: dict[str, Any], sheet: dict[str, Any]) -> list[list[str]]:
    columns = workbook_sheet_columns(sheet)
    rows = [[header for _field, header in columns]]
    collections = master.get("collections")
    collection_name = sheet.get("collection")
    if not isinstance(collection_name, str) or not collection_name:
        return rows
    records = collections.get(collection_name) if isinstance(collections, dict) else []
    if not isinstance(records, list):
        return rows
    for record in records:
        if isinstance(record, dict):
            rows.append([scalar_text(record.get(field)) for field, _header in columns])
    return rows


def workbook_sheets(master: dict[str, Any]) -> list[tuple[str, list[list[str]]]]:
    sheets = workbook_layout(master).get("sheets")
    if not isinstance(sheets, list):
        return []
    rendered_sheets: list[tuple[str, list[list[str]]]] = []
    for sheet in sheets:
        if not isinstance(sheet, dict):
            continue
        sheet_name = scalar_text(sheet.get("name") or sheet.get("collection") or "Sheet")
        rendered_sheets.append((sheet_name, sheet_rows(master, sheet)))
    return rendered_sheets


def worksheet_dimension(rows: list[list[str]]) -> str:
    row_count = max(len(rows), 1)
    column_count = max((len(row) for row in rows), default=1)
    return f"A1:{column_name(column_count)}{row_count}"


def column_widths(rows: list[list[str]]) -> list[float]:
    column_count = max((len(row) for row in rows), default=1)
    widths: list[float] = []
    for column_index in range(column_count):
        width = max((len(row[column_index]) for row in rows if column_index < len(row)), default=8)
        widths.append(float(max(10, min(width + 2, 42))))
    return widths


def style_bool(styles: dict[str, Any], key: str, default: bool) -> bool:
    value = styles.get(key)
    return value if isinstance(value, bool) else default


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


def worksheet_xml(rows: list[list[str]], layout: dict[str, Any]) -> str:
    styles = workbook_style_options(layout)
    rendered_rows: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells: list[str] = []
        for column_index, value in enumerate(row, start=1):
            cell_ref = f"{column_name(column_index)}{row_index}"
            style_id = "1" if row_index == 1 else "2"
            cells.append(
                f'<c r="{cell_ref}" s="{style_id}" t="inlineStr"><is><t xml:space="preserve">'
                f"{xml_text(value)}</t></is></c>"
            )
        rendered_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    widths = "".join(
        f'<col min="{index}" max="{index}" width="{width:.2f}" customWidth="1"/>'
        for index, width in enumerate(column_widths(rows), start=1)
    )
    dimension = worksheet_dimension(rows)
    sheet_view_xml = '<sheetViews><sheetView tabSelected="0" workbookViewId="0">'
    if style_bool(styles, "freeze_header", True):
        sheet_view_xml += '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        sheet_view_xml += '<selection pane="bottomLeft"/>'
    sheet_view_xml += "</sheetView></sheetViews>"
    auto_filter_xml = f'<autoFilter ref="{dimension}"/>' if style_bool(styles, "auto_filter", True) else ""
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
    sheet_names = [sheet_name for sheet_name, _rows in sheets]
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
            (f"xl/worksheets/sheet{index}.xml", worksheet_xml(rows, layout))
            for index, (_sheet_name, rows) in enumerate(sheets, start=1)
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
        if not text:
            continue
        rendered_runs.append(
            {
                "text": text,
                "bold": run.get("bold") is True,
                "italic": run.get("italic") is True,
                "underline": run.get("underline") is True,
            }
        )
    return rendered_runs


def item_paragraphs(item: dict[str, Any], records_by_id: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    formatting = item_formatting(item)
    style = formatting_style(formatting)
    numbering = formatting_numbering(formatting)
    runs = formatting_runs(formatting)
    if runs:
        return [{"style": style, "numbering": numbering, "runs": runs}]

    text = resume_item_text(item, records_by_id)
    paragraphs: list[dict[str, Any]] = []
    for paragraph in text.splitlines() or [text]:
        if paragraph.strip():
            paragraphs.append({"style": style, "numbering": numbering, "runs": [{"text": paragraph}]})
    return paragraphs


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


def docx_run(run: dict[str, Any]) -> str:
    properties: list[str] = []
    if run.get("bold") is True:
        properties.append("<w:b/>")
    if run.get("italic") is True:
        properties.append("<w:i/>")
    if run.get("underline") is True:
        properties.append('<w:u w:val="single"/>')
    properties_xml = f'<w:rPr>{"".join(properties)}</w:rPr>' if properties else ""
    return f'<w:r>{properties_xml}<w:t xml:space="preserve">{xml_text(scalar_text(run.get("text")))}</w:t></w:r>'


def docx_paragraph(paragraph: dict[str, Any]) -> str:
    paragraph_properties: list[str] = []
    style = paragraph.get("style")
    if isinstance(style, str) and style:
        paragraph_properties.append(f'<w:pStyle w:val="{xml_attr(style)}"/>')
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
    if not rendered_runs:
        rendered_runs = [docx_run({"text": ""})]
    return f'<w:p>{properties_xml}{"".join(rendered_runs)}</w:p>'


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


def docx_document_xml(paragraphs: list[dict[str, Any]], layout: dict[str, Any]) -> str:
    paragraph_xml = "".join(docx_paragraph(paragraph) for paragraph in paragraphs)
    margins = docx_margins(layout)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{paragraph_xml}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        f'<w:pgMar w:top="{margins["top"]}" w:right="{margins["right"]}" '
        f'w:bottom="{margins["bottom"]}" w:left="{margins["left"]}" '
        f'w:header="{margins["header"]}" w:footer="{margins["footer"]}" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


def document_content_types() -> str:
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
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
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


def docx_document_relationships_xml() -> str:
    return relationships_xml(
        [
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"),
            ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml"),
            ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml"),
        ]
    )


def docx_settings_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{WORD_NS}"><w:defaultTabStop w:val="720"/></w:settings>'
    )


def write_resume_document(master: dict[str, Any], resume: dict[str, Any], path: Path) -> None:
    layout = document_layout(resume, path.name)
    paragraphs = resume_paragraphs(master, resume)
    resume_value = resume.get("resume")
    resume_meta = resume_value if isinstance(resume_value, dict) else {}
    title = scalar_text(resume_meta.get("label") or "Resume")
    entries = [
        ("[Content_Types].xml", document_content_types()),
        (
            "_rels/.rels",
            relationships_xml(
                [
                    (
                        "rId1",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                        "word/document.xml",
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
        ("word/_rels/document.xml.rels", docx_document_relationships_xml()),
        ("word/document.xml", docx_document_xml(paragraphs, layout)),
        ("word/styles.xml", docx_styles_xml(layout)),
        ("word/numbering.xml", docx_numbering_xml(paragraphs, layout)),
        ("word/settings.xml", docx_settings_xml()),
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


def load_golden_json(evidence_dir: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    json_dir = evidence_dir / "generated" / "json" / "a_posteriori"
    return (
        read_json(json_dir / "master_profile_a_posteriori.json"),
        read_json(json_dir / "resume_2017_a_posteriori.json"),
        read_json(json_dir / "resume_2023_a_posteriori.json"),
    )


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


def office_file_summary(relative_path: str, path: Path) -> dict[str, Any]:
    summary: dict[str, Any] = {"path": relative_path, "bytes": path.stat().st_size, "sha256": sha256_file(path)}
    try:
        summary["normalized_text"] = normalized_text_summary(path)
    except (KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        summary["normalized_text"] = {"error": exc.__class__.__name__}
    return summary


def write_office_comparison_report(evidence_dir: Path) -> Path:
    comparisons: list[dict[str, Any]] = []
    for generated_name, generated_relative_path in GOLDEN_OFFICE_EXPECTATIONS.items():
        source_relative_path = GOLDEN_SOURCE_OFFICE[generated_name]
        generated_summary = office_file_summary(generated_relative_path, evidence_dir / generated_relative_path)
        source_summary = office_file_summary(source_relative_path, evidence_dir / source_relative_path)
        generated_normalized = generated_summary.get("normalized_text")
        source_normalized = source_summary.get("normalized_text")
        normalized_match = isinstance(generated_normalized, dict) and generated_normalized == source_normalized
        comparisons.append(
            {
                "generated": generated_summary,
                "source": source_summary,
                "byte_identical": generated_summary["sha256"] == source_summary["sha256"],
                "normalized_text_match": normalized_match,
                "status": "normalized_match" if normalized_match else "review_required",
            }
        )
    report_path = evidence_dir / GOLDEN_OFFICE_REPORT
    write_json(
        report_path,
        {
            "schema_version": SCHEMA_VERSION,
            "generator": f"x_create_cv_x {VERSION}",
            "purpose": "Private comparison report for generated _a_posteriori Office evidence",
            "comparisons": comparisons,
        },
    )
    return report_path


def write_golden_office_outputs(
    evidence_dir: Path, output_root: Path | None = None, *, write_report: bool = True
) -> tuple[int, int]:
    master, resume_2017, resume_2023 = load_golden_json(evidence_dir)
    target_root = output_root or evidence_dir
    output_dir = target_root / "generated" / "office" / "a_posteriori"
    write_master_workbook(master, output_dir / "master_profile_a_posteriori.xlsx")
    write_resume_document(master, resume_2017, output_dir / "resume_2017_a_posteriori.docx")
    write_resume_document(master, resume_2023, output_dir / "resume_2023_a_posteriori.docx")
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
    write_json(master_path(db_dir), empty_master(profile, workbook_layout_value))


def set_profile(db_dir: Path, profile: dict[str, Any]) -> None:
    master = read_json(master_path(db_dir))
    master["profile"] = profile
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
    write_json(path, empty_resume(resume, document_layout_value))


def add_resume_record(db_dir: Path, resume_id: str, collection: str, record: dict[str, Any], replace: bool) -> None:
    resume = read_json(resume_path(db_dir, resume_id))
    collections = resume.get("collections")
    if not isinstance(collections, dict):
        raise ValueError(f"Resume collections are missing or invalid: {resume_id}")
    records = collections.get(collection)
    if not isinstance(records, list):
        raise ValueError(f"Unknown resume collection: {collection}")
    append_unique(records, record, replace)
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
