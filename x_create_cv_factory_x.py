#!/usr/bin/env python3
"""Command-line CRUD factory for app-native CV JSON data."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import shutil
import sys
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
HASH_CHUNK_SIZE = 1024 * 1024

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


def check_golden_evidence(evidence_dir: Path) -> int:
    return check_evidence_manifest(evidence_dir / "a_priori_manifest.json")


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
        checked_count = check_golden_evidence(args.evidence_dir)
        print(f"Golden exercise passed: {checked_count} a priori evidence files")
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
