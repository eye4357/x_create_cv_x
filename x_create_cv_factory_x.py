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


def master_profile_rows(master: dict[str, Any]) -> list[list[str]]:
    rows = [["scope", "collection", "record_id", "field", "value"]]
    profile = master.get("profile")
    if isinstance(profile, dict):
        for field, value in profile.items():
            rows.append(["profile", "profile", "", str(field), scalar_text(value)])

    collections = master.get("collections")
    if isinstance(collections, dict):
        collection_names = [name for name in MASTER_COLLECTIONS if name in collections]
        collection_names.extend(sorted(name for name in collections if name not in MASTER_COLLECTIONS))
        for collection_name in collection_names:
            records = collections.get(collection_name)
            if not isinstance(records, list):
                continue
            for record in records:
                if not isinstance(record, dict):
                    continue
                record_id = scalar_text(record.get("id"))
                for field, value in record.items():
                    rows.append(["collection", str(collection_name), record_id, str(field), scalar_text(value)])
    return rows


def worksheet_xml(rows: list[list[str]]) -> str:
    rendered_rows: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells: list[str] = []
        for column_index, value in enumerate(row, start=1):
            cell_ref = f"{column_name(column_index)}{row_index}"
            cells.append(f'<c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{xml_text(value)}</t></is></c>')
        rendered_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    last_row = max(len(rows), 1)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:E{last_row}"/><sheetData>{"".join(rendered_rows)}</sheetData></worksheet>'
    )


def workbook_content_types() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def write_master_workbook(master: dict[str, Any], path: Path) -> None:
    entries = [
        ("[Content_Types].xml", workbook_content_types()),
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
        (
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Master Profile" sheetId="1" r:id="rId1"/></sheets></workbook>',
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships_xml(
                [
                    (
                        "rId1",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        "worksheets/sheet1.xml",
                    ),
                    (
                        "rId2",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                        "styles.xml",
                    ),
                ]
            ),
        ),
        (
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            "</styleSheet>",
        ),
        ("xl/worksheets/sheet1.xml", worksheet_xml(master_profile_rows(master))),
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


def resume_lines(master: dict[str, Any], resume: dict[str, Any]) -> list[tuple[str, str | None]]:
    records_by_id = master_records_by_id(master)
    resume_value = resume.get("resume")
    resume_meta = resume_value if isinstance(resume_value, dict) else {}
    label = scalar_text(resume_meta.get("label") or resume_meta.get("id") or "Resume")
    lines: list[tuple[str, str | None]] = [(label, "Heading1")]
    status = scalar_text(resume_meta.get("status"))
    if status:
        lines.append((f"Status: {status}", None))

    collections_value = resume.get("collections")
    collections = collections_value if isinstance(collections_value, dict) else {}
    sections = collections.get("sections")
    items = collections.get("items")
    section_records = (
        [section for section in sections if isinstance(section, dict)] if isinstance(sections, list) else []
    )
    item_records = [item for item in items if isinstance(item, dict)] if isinstance(items, list) else []

    for section in sorted(section_records, key=sort_order):
        if section.get("is_visible") is False:
            continue
        title = scalar_text(section.get("title") or section.get("id") or "Section")
        lines.append((title, "Heading2"))
        section_id = section.get("id")
        section_items = [item for item in item_records if item.get("section_id") == section_id]
        for item in sorted(section_items, key=sort_order):
            if item.get("is_visible") is False:
                continue
            item_text = resume_item_text(item, records_by_id)
            for paragraph in item_text.splitlines() or [item_text]:
                if paragraph.strip():
                    lines.append((paragraph, None))
    return lines


def docx_paragraph(text: str, style: str | None = None) -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{xml_attr(style)}"/></w:pPr>' if style else ""
    return f'<w:p>{style_xml}<w:r><w:t xml:space="preserve">{xml_text(text)}</w:t></w:r></w:p>'


def docx_document_xml(lines: list[tuple[str, str | None]]) -> str:
    paragraphs = "".join(docx_paragraph(text, style) for text, style in lines)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}"><w:body>{paragraphs}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" '
        'w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>'
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
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def write_resume_document(master: dict[str, Any], resume: dict[str, Any], path: Path) -> None:
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
        ("word/document.xml", docx_document_xml(resume_lines(master, resume))),
        (
            "word/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:styles xmlns:w="{WORD_NS}">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
            '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
            '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
            "</w:styles>",
        ),
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


def empty_master(profile: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "app_model": MASTER_APP_MODEL,
        "profile": profile,
        "collections": {collection: [] for collection in MASTER_COLLECTIONS},
    }


def empty_resume(resume: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "app_model": RESUME_APP_MODEL,
        "resume": resume,
        "collections": {
            "sections": [],
            "items": [],
        },
    }


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


def init_database(db_dir: Path, profile: dict[str, Any], force: bool) -> None:
    if db_dir.exists() and force:
        shutil.rmtree(db_dir)
    if master_path(db_dir).exists():
        raise ValueError(f"Database already exists at {db_dir}; pass --force to replace it")
    write_json(master_path(db_dir), empty_master(profile))


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


def add_resume(db_dir: Path, resume: dict[str, Any], replace: bool) -> None:
    resume_id = resume.get("id")
    if not isinstance(resume_id, str) or not resume_id:
        raise ValueError("Resume JSON must contain a non-empty string id")
    path = resume_path(db_dir, resume_id)
    if path.exists() and not replace:
        raise ValueError(f"Resume already exists: {resume_id}")
    write_json(path, empty_resume(resume))


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
    replace: bool = False,
) -> str:
    resume = ordered_record(
        RESUME_FIELD_ORDER,
        {"id": resume_id, "label": label, "status": status, "created_at": created_at, "updated_at": updated_at},
    )
    add_resume(db_dir, resume, replace)
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
    add_replace_option(add_resume_parser)

    create_resume_parser = subparsers.add_parser("create-resume", help="Create a resume document")
    add_common_db_options(create_resume_parser)
    add_payload_options(create_resume_parser, "resume")
    add_field_options(create_resume_parser, RESUME_FIELD_ORDER)
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
        init_database(args.db_dir, profile, args.force)
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
        add_resume(args.db_dir, resume, args.replace)
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
