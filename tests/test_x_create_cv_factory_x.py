from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from typing import Any, cast

import pytest

import x_create_cv_factory_x as app

FIXTURE_PATH = Path(__file__).parent / "fixtures" / "fake_profile_seed.json"


def load_seed() -> dict[str, Any]:
    value = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise AssertionError("fake seed fixture must be a JSON object")
    return cast(dict[str, Any], value)


def object_at(container: dict[str, Any], key: str) -> dict[str, Any]:
    value = container[key]
    if not isinstance(value, dict):
        raise AssertionError(f"{key} must be a JSON object")
    return cast(dict[str, Any], value)


def objects_at(container: dict[str, Any], key: str) -> list[dict[str, Any]]:
    value = container[key]
    if not isinstance(value, list) or not all(isinstance(item, dict) for item in value):
        raise AssertionError(f"{key} must be a list of JSON objects")
    return cast(list[dict[str, Any]], value)


def string_list(value: Any) -> list[str]:
    if not isinstance(value, list):
        raise AssertionError("expected a JSON list")
    return [str(item) for item in value]


def read_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise AssertionError(f"expected object JSON in {path}")
    return cast(dict[str, Any], value)


def build_fake_database(db_dir: Path) -> None:
    seed = load_seed()
    profile = object_at(seed, "profile")
    person = object_at(seed, "person")
    contact = object_at(seed, "contact_method")
    address = object_at(seed, "address")
    project = object_at(seed, "project")
    achievement = object_at(seed, "achievement")
    job = object_at(seed, "job")
    skill = object_at(seed, "skill")
    text_block = object_at(seed, "text_block")
    formatting = object_at(seed, "formatting")

    app.init_master_profile(
        db_dir,
        profile_id=str(profile["id"]),
        person_id=str(profile["person_id"]),
        label=str(profile["label"]),
        created_at=str(profile["created_at"]),
        updated_at=str(profile["updated_at"]),
        force=True,
    )
    app.add_user(
        db_dir,
        record_id=str(person["id"]),
        display_name=str(person["display_name"]),
        preferred_name=str(person["preferred_name"]),
    )
    app.add_contact_method(
        db_dir,
        record_id=str(contact["id"]),
        person_id=str(contact["person_id"]),
        kind=str(contact["kind"]),
        value=str(contact["value"]),
        is_primary=bool(contact["is_primary"]),
    )
    app.add_address(
        db_dir, record_id=str(address["id"]), label=str(address["label"]), address_text=str(address["address_text"])
    )
    app.add_project(db_dir, record_id=str(project["id"]), name=str(project["name"]))
    app.add_achievement(
        db_dir, record_id=str(achievement["id"]), text=str(achievement["text"]), job_id=str(achievement["job_id"])
    )
    app.add_job(
        db_dir,
        record_id=str(job["id"]),
        label=str(job["label"]),
        job_title=str(job["job_title"]),
        employer=str(job["employer"]),
        start_date=str(job["start_date"]),
        end_date=str(job["end_date"]),
        address_id=str(job["address_id"]),
        project_ids=string_list(job["project_ids"]),
        achievement_ids=string_list(job["achievement_ids"]),
    )
    app.add_skill(db_dir, record_id=str(skill["id"]), name=str(skill["name"]), category=str(skill["category"]))
    app.add_text_block(
        db_dir,
        record_id=str(text_block["id"]),
        text=str(text_block["text"]),
        section_label=str(text_block["section_label"]),
        used_in_resume_ids=string_list(text_block["used_in_resume_ids"]),
        is_job_related=bool(text_block["is_job_related"]),
    )

    for resume in objects_at(seed, "resumes"):
        resume_id = str(resume["id"])
        section_id = str(resume["section_id"])
        item_id = str(resume["item_id"])
        app.create_resume(
            db_dir,
            resume_id=resume_id,
            label=str(resume["label"]),
            status=str(resume["status"]),
            created_at=str(resume["created_at"]),
            updated_at=str(resume["updated_at"]),
        )
        app.add_resume_section(
            db_dir,
            resume_id=resume_id,
            section_id=section_id,
            title="Summary",
            kind="summary",
            sort_order=1,
            item_ids=[item_id],
            is_visible=True,
        )
        app.add_resume_item(
            db_dir,
            resume_id=resume_id,
            item_id=item_id,
            section_id=section_id,
            kind="text",
            sort_order=1,
            master_record_refs=[str(text_block["id"])],
            text_override=str(text_block["text"]),
            formatting=formatting,
            is_visible=True,
        )


def make_expected_zip(db_dir: Path, zip_path: Path, prefix: str) -> None:
    with zipfile.ZipFile(zip_path, "w") as zip_file:
        for filename in app.EXPECTED_FILES:
            zip_file.write(db_dir / filename, f"{prefix}/{filename}")


def test_version_constant_matches_cli(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exit_info:
        app.main(["--version"])

    assert exit_info.value.code == 0
    assert f"x_create_cv_x {app.VERSION}" in capsys.readouterr().out


def test_fake_seed_rebuilds_master_and_resume_json(tmp_path: Path) -> None:
    build_fake_database(tmp_path)

    master = read_json(tmp_path / app.MASTER_FILE)
    assert master["schema_version"] == app.SCHEMA_VERSION
    assert master["app_model"] == app.MASTER_APP_MODEL
    assert master["collections"]["people"][0]["display_name"] == "Ada Example"
    assert master["collections"]["contact_methods"][0]["value"] == "ada.example@example.test"
    assert master["collections"]["jobs"][0]["project_ids"] == ["project_fake"]

    resume = read_json(tmp_path / "resume_2023.json")
    assert resume["schema_version"] == app.SCHEMA_VERSION
    assert resume["app_model"] == app.RESUME_APP_MODEL
    assert resume["collections"]["sections"][0]["item_ids"] == ["item_summary"]
    assert resume["collections"]["items"][0]["master_record_refs"] == ["text_fake_summary"]


def test_validate_against_zip_passes_for_fake_database(tmp_path: Path) -> None:
    db_dir = tmp_path / "db"
    zip_path = tmp_path / "expected.zip"
    build_fake_database(db_dir)
    make_expected_zip(db_dir, zip_path, "expected")

    app.validate_against_zip(db_dir, zip_path, "expected")


def test_validate_against_zip_reports_mismatch(tmp_path: Path) -> None:
    db_dir = tmp_path / "db"
    zip_path = tmp_path / "expected.zip"
    build_fake_database(db_dir)
    make_expected_zip(db_dir, zip_path, "expected")
    (db_dir / "resume_2023.json").write_text("{}", encoding="utf-8")

    with pytest.raises(ValueError, match="mismatch resume_2023.json"):
        app.validate_against_zip(db_dir, zip_path, "expected")


def test_check_evidence_manifest_passes_for_fake_files(tmp_path: Path) -> None:
    evidence_file = tmp_path / "source_office" / "a_priori" / "fake_a_priori.docx"
    evidence_file.parent.mkdir(parents=True)
    evidence_file.write_bytes(b"fake office package")
    manifest_path = tmp_path / "a_priori_manifest.json"
    app.write_json(
        manifest_path,
        {
            "schema_version": "1.0.0",
            "algorithm": "sha256",
            "files": [
                {
                    "path": "source_office/a_priori/fake_a_priori.docx",
                    "bytes": evidence_file.stat().st_size,
                    "sha256": app.sha256_file(evidence_file),
                }
            ],
        },
    )

    assert app.check_evidence_manifest(manifest_path) == 1


def test_check_evidence_manifest_fails_fast_on_hash_mismatch(tmp_path: Path) -> None:
    evidence_file = tmp_path / "source_office" / "a_priori" / "fake_a_priori.xlsx"
    evidence_file.parent.mkdir(parents=True)
    evidence_file.write_bytes(b"fake office package")
    manifest_path = tmp_path / "a_priori_manifest.json"
    app.write_json(
        manifest_path,
        {
            "schema_version": "1.0.0",
            "algorithm": "sha256",
            "files": [
                {
                    "path": "source_office/a_priori/fake_a_priori.xlsx",
                    "bytes": evidence_file.stat().st_size,
                    "sha256": "0" * 64,
                }
            ],
        },
    )

    with pytest.raises(ValueError, match="Evidence hash mismatch"):
        app.check_evidence_manifest(manifest_path)


def test_cli_check_evidence_reports_missing_manifest(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    result = app.main(["check-evidence", "--manifest", str(tmp_path / "missing.json")])

    assert result == 1
    assert "Missing JSON file" in capsys.readouterr().err


def test_cli_exercise_golden_uses_evidence_dir(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    evidence_dir = tmp_path / "evidence"
    evidence_file = evidence_dir / "source_office" / "a_priori" / "fake_a_priori.docx"
    script_dir = evidence_dir / "scripts"
    expected_json_dir = evidence_dir / "generated" / "json" / "a_posteriori"
    evidence_file.parent.mkdir(parents=True)
    script_dir.mkdir(parents=True)
    expected_json_dir.mkdir(parents=True)
    evidence_file.write_bytes(b"fake office package")
    expected_json = {
        "master_profile.json": ('{"file":"master"}\n', "master_profile_a_posteriori.json"),
        "resume_2017.json": ('{"file":"resume_2017"}\n', "resume_2017_a_posteriori.json"),
        "resume_2023.json": ('{"file":"resume_2023"}\n', "resume_2023_a_posteriori.json"),
    }
    script_outputs = [
        ("01_build_master_data.py", "master_profile.json", '{"file":"master"}'),
        ("02_build_old_resume.py", "resume_2017.json", '{"file":"resume_2017"}'),
        ("03_build_new_resume.py", "resume_2023.json", '{"file":"resume_2023"}'),
    ]
    for script_name, output_name, output_json in script_outputs:
        (script_dir / script_name).write_text(
            "from pathlib import Path\n"
            "import os\n"
            "out = Path(os.environ['CV_FACTORY_DB_DIR'])\n"
            "out.mkdir(parents=True, exist_ok=True)\n"
            f"(out / {output_name!r}).write_text({output_json!r} + '\\n', encoding='utf-8')\n",
            encoding="utf-8",
        )
    for _generated_name, (content, expected_name) in expected_json.items():
        (expected_json_dir / expected_name).write_text(content, encoding="utf-8")
    generated_office_count, report_count = app.write_golden_office_outputs(evidence_dir, write_report=False)
    assert generated_office_count == 3
    assert report_count == 0
    a_priori_manifest = tmp_path / "evidence" / "a_priori_manifest.json"
    app.write_json(
        a_priori_manifest,
        {
            "schema_version": "1.0.0",
            "algorithm": "sha256",
            "files": [
                {
                    "path": "source_office/a_priori/fake_a_priori.docx",
                    "bytes": evidence_file.stat().st_size,
                    "sha256": app.sha256_file(evidence_file),
                }
            ],
        },
    )
    app.write_json(
        tmp_path / "evidence" / "chain_of_evidence_manifest.json",
        {
            "schema_version": "1.0.0",
            "algorithm": "sha256",
            "files": [
                {
                    "path": "a_priori_manifest.json",
                    "bytes": a_priori_manifest.stat().st_size,
                    "sha256": app.sha256_file(a_priori_manifest),
                },
                {
                    "path": "source_office/a_priori/fake_a_priori.docx",
                    "bytes": evidence_file.stat().st_size,
                    "sha256": app.sha256_file(evidence_file),
                },
                *[
                    {
                        "path": f"scripts/{script_name}",
                        "bytes": (script_dir / script_name).stat().st_size,
                        "sha256": app.sha256_file(script_dir / script_name),
                    }
                    for script_name, _output_name, _output_json in script_outputs
                ],
                *[
                    {
                        "path": f"generated/json/a_posteriori/{expected_name}",
                        "bytes": (expected_json_dir / expected_name).stat().st_size,
                        "sha256": app.sha256_file(expected_json_dir / expected_name),
                    }
                    for _generated_name, (_content, expected_name) in expected_json.items()
                ],
                *[
                    {
                        "path": expected_path,
                        "bytes": (evidence_dir / expected_path).stat().st_size,
                        "sha256": app.sha256_file(evidence_dir / expected_path),
                    }
                    for expected_path in app.GOLDEN_OFFICE_EXPECTATIONS.values()
                ],
            ],
        },
    )

    assert app.main(["exercise-golden", "--evidence-dir", str(evidence_dir)]) == 0
    assert (
        "Golden exercise passed: 1 a priori evidence files; 11 chain files; 3 generated JSON files; "
        "3 generated Office files" in capsys.readouterr().out
    )


def test_cli_typed_flags_build_expected_payload(tmp_path: Path) -> None:
    assert (
        app.main(
            [
                "init",
                "--db-dir",
                str(tmp_path),
                "--force",
                "--id",
                "profile_cli",
                "--person-id",
                "person_cli",
                "--label",
                "CLI Test",
                "--created-at",
                "2026-07-01T00:00:00Z",
                "--updated-at",
                "2026-07-01T00:00:00Z",
            ]
        )
        == 0
    )
    assert (
        app.main(
            [
                "add-user",
                "--db-dir",
                str(tmp_path),
                "--id",
                "person_cli",
                "--display-name",
                "CLI Example",
                "--preferred-name",
                "CLI",
            ]
        )
        == 0
    )

    master = read_json(tmp_path / app.MASTER_FILE)
    assert master["profile"]["id"] == "profile_cli"
    assert master["collections"]["people"][0]["preferred_name"] == "CLI"


def test_cli_rejects_invalid_json_payload(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    result = app.main(["init", "--db-dir", str(tmp_path), "--json", "[]"])

    assert result == 1
    assert "profile JSON must be an object" in capsys.readouterr().err


def test_parse_bool_rejects_unknown_value() -> None:
    with pytest.raises(argparse.ArgumentTypeError, match="Expected a boolean value"):
        app.parse_bool("sometimes")
