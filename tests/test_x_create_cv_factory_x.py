from __future__ import annotations

import argparse
import json
import xml.etree.ElementTree as ET
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


def workbook_sheet_names(path: Path) -> list[str]:
    namespace = {"s": app.SHEET_NS}
    with zipfile.ZipFile(path) as workbook:
        root = ET.fromstring(workbook.read("xl/workbook.xml"))
    return [str(sheet.attrib["name"]) for sheet in root.findall(".//s:sheet", namespace)]


def worksheet_row_values(path: Path, row_index: int) -> list[str]:
    namespace = {"s": app.SHEET_NS}
    with zipfile.ZipFile(path) as workbook:
        root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
    row = root.find(f".//s:row[@r='{row_index}']", namespace)
    if row is None:
        raise AssertionError(f"generated XLSX must contain row {row_index}")
    return [
        "".join(node.text or "" for node in cell.findall(".//s:t", namespace)) for cell in row.findall("s:c", namespace)
    ]


def first_worksheet_summary(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path) as workbook:
        return app.xlsx_sheet_summary(workbook, "xl/worksheets/sheet1.xml", [])


def first_worksheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as workbook:
        return workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")


def docx_document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as document:
        return document.read("word/document.xml").decode("utf-8")


def docx_relationships_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as document:
        return document.read("word/_rels/document.xml.rels").decode("utf-8")


def docx_part_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as document:
        return sorted(document.namelist())


def docx_theme_style_counts(path: Path) -> dict[str, int]:
    namespace = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    with zipfile.ZipFile(path) as document:
        theme = ET.fromstring(document.read("word/theme/theme1.xml"))
    counts: dict[str, int] = {}
    for list_name in ["fillStyleLst", "lnStyleLst", "effectStyleLst", "bgFillStyleLst"]:
        node = theme.find(f".//a:{list_name}", namespace)
        counts[list_name] = len(list(node)) if node is not None else 0
    return counts


def docx_section_margins(path: Path) -> dict[str, str]:
    namespace = {"w": app.WORD_NS}
    with zipfile.ZipFile(path) as document:
        root = ET.fromstring(document.read("word/document.xml"))
    margins = root.find(".//w:pgMar", namespace)
    if margins is None:
        raise AssertionError("generated DOCX must contain section margins")
    return {key.split("}")[-1]: value for key, value in margins.attrib.items()}


def test_version_constant_matches_cli(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exit_info:
        app.main(["--version"])

    assert exit_info.value.code == 0
    assert f"x_create_cv_x {app.VERSION}" in capsys.readouterr().out


def test_fake_seed_rebuilds_master_and_resume_json(tmp_path: Path) -> None:
    build_fake_database(tmp_path)

    master = read_json(tmp_path / app.MASTER_FILE)
    app.validate_contract(master, "master_profile")
    assert master["schema_version"] == app.SCHEMA_VERSION
    assert master["app_model"] == app.MASTER_APP_MODEL
    assert master["collections"]["people"][0]["display_name"] == "Ada Example"
    assert master["collections"]["contact_methods"][0]["value"] == "ada.example@example.test"
    assert master["collections"]["jobs"][0]["project_ids"] == ["project_fake"]

    resume = read_json(tmp_path / "resume_2023.json")
    app.validate_contract(resume, "resume")
    assert resume["schema_version"] == app.SCHEMA_VERSION
    assert resume["app_model"] == app.RESUME_APP_MODEL
    assert resume["collections"]["sections"][0]["item_ids"] == ["item_summary"]
    assert resume["collections"]["items"][0]["master_record_refs"] == ["text_fake_summary"]


def test_cli_validate_schema_accepts_generated_fake_json(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    build_fake_database(tmp_path)

    assert app.main(["validate-schema", str(tmp_path / app.MASTER_FILE), str(tmp_path / "resume_2023.json")]) == 0

    assert "Schema validation passed: 1 master_profile, 1 resume" in capsys.readouterr().out


def test_schema_validation_reports_contract_path() -> None:
    invalid_resume = app.empty_resume(
        {
            "id": "resume_invalid",
            "label": "Invalid Resume",
            "status": "draft",
            "created_at": "2026-07-02T00:00:00Z",
            "updated_at": "2026-07-02T00:00:00Z",
        }
    )
    invalid_resume["collections"]["sections"].append(
        {"id": "section_invalid", "title": "Broken", "kind": "summary", "sort_order": "1", "item_ids": []}
    )

    with pytest.raises(ValueError, match=r"\$\.collections\.sections\[0\]\.sort_order"):
        app.validate_contract(invalid_resume, "resume")


def test_layout_contract_schemas_accept_default_layouts() -> None:
    app.validate_contract(app.default_workbook_layout(), "workbook_layout")
    app.validate_contract(app.default_document_layout("resume_2017_a_posteriori.docx"), "document_layout")
    app.validate_contract_file(app.DEFAULT_AUDIT_POLICY, "audit_policy")


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


def test_office_generation_consumes_layout_contracts(tmp_path: Path) -> None:
    master: dict[str, Any] = {
        "office_layout": {
            "workbook": {
                "styles": {
                    "freeze_header": False,
                    "auto_filter": False,
                    "header_fill": "FFABCDEF",
                    "header_font_color": "FF102030",
                    "body_border_color": "FF405060",
                    "page_margins": {
                        "left": "0.2",
                        "right": "0.4",
                        "top": "0.6",
                        "bottom": "0.8",
                        "header": "0.1",
                        "footer": "0.3",
                    },
                },
                "sheets": [
                    {
                        "name": "Highlights",
                        "collection": "highlights",
                        "skip_placeholder_rows": True,
                        "freeze_header": True,
                        "auto_filter": True,
                        "column_widths": [8.5, 13, 28.25],
                        "columns": [
                            {"field": "id", "header": "ID"},
                            {"field": "label", "header": "Label"},
                            {"field": "highlights", "header": "Highlights"},
                            {"field": "score", "header": "Score", "value_type": "number"},
                            {"field": "is_current", "header": "Current", "value_type": "boolean"},
                        ],
                    }
                ],
            }
        },
        "collections": {
            "highlights": [
                {"id": "highlight_000", "label": "Highlights", "highlights": ""},
                {
                    "id": "highlight_001",
                    "label": "One",
                    "a": "raw value",
                    "highlights": "Structured value",
                    "score": "98.5",
                    "is_current": True,
                },
            ]
        },
    }
    app.validate_contract(master["office_layout"]["workbook"], "workbook_layout")
    workbook_path = tmp_path / "layout_contract.xlsx"
    app.write_master_workbook(master, workbook_path)
    with zipfile.ZipFile(workbook_path) as workbook:
        workbook_part_names = sorted(workbook.namelist())
        content_types_xml = workbook.read("[Content_Types].xml").decode("utf-8")
        root_relationships_xml = workbook.read("_rels/.rels").decode("utf-8")
        core_properties_xml = workbook.read("docProps/core.xml").decode("utf-8")
        app_properties_xml = workbook.read("docProps/app.xml").decode("utf-8")
        workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
        workbook_relationships_xml = workbook.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        styles_xml = workbook.read("xl/styles.xml").decode("utf-8")

    assert workbook_sheet_names(workbook_path) == ["Highlights"]
    assert workbook_part_names == [
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/app.xml",
        "docProps/core.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
    ]
    assert content_types_xml == (
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
    assert (
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>' in content_types_xml
    )
    assert (
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' in content_types_xml
    )
    assert root_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/><Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/></Relationships>'
    )
    assert (
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"' in root_relationships_xml
    )
    assert (
        'Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"' in root_relationships_xml
    )
    assert (
        'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"' in root_relationships_xml
    )
    assert core_properties_xml == (
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
        "</cp:coreProperties>"
    )
    assert app_properties_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>x_create_cv_x</Application></Properties>"
    )
    assert workbook_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{app.SHEET_NS}" xmlns:r="{app.REL_NS}"><sheets>'
        '<sheet name="Highlights" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    assert '<sheet name="Highlights" sheetId="1" r:id="rId1"/>' in workbook_xml
    assert workbook_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/></Relationships>'
    )
    assert (
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"' in workbook_relationships_xml
    )
    assert (
        'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"' in workbook_relationships_xml
    )
    assert styles_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{app.SHEET_NS}">'
        '<fonts count="3">'
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
        '<font><b/><sz val="11"/><color rgb="FF102030"/><name val="Calibri"/><family val="2"/></font>'
        '<font><i/><sz val="11"/><color rgb="FF666666"/><name val="Calibri"/><family val="2"/></font>'
        '</fonts><fills count="4">'
        '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFABCDEF"/>'
        '<bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FF405060"/></left>'
        '<right style="thin"><color rgb="FF405060"/></right>'
        '<top style="thin"><color rgb="FF405060"/></top>'
        '<bottom style="thin"><color rgb="FF405060"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" '
        'applyFont="1" applyFill="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )
    assert worksheet_row_values(workbook_path, 1) == ["ID", "Label", "Highlights", "Score", "Current"]
    assert worksheet_row_values(workbook_path, 2) == ["highlight_001", "One", "Structured value", "", ""]
    sheet_summary = first_worksheet_summary(workbook_path)
    assert sheet_summary["path"] == "xl/worksheets/sheet1.xml"
    assert sheet_summary["dimension"] == "A1:E2"
    assert sheet_summary["headers"] == ["ID", "Label", "Highlights", "Score", "Current"]
    assert sheet_summary["freeze_pane"] == {
        "ySplit": "1",
        "topLeftCell": "A2",
        "activePane": "bottomLeft",
        "state": "frozen",
    }
    assert sheet_summary["auto_filter_ref"] == "A1:E2"
    assert sheet_summary["row_count"] == 2
    assert sheet_summary["column_count"] == 5
    assert sheet_summary["styled_cell_count"] == 10
    assert sheet_summary["page_margins"] == {
        "left": "0.2",
        "right": "0.4",
        "top": "0.6",
        "bottom": "0.8",
        "header": "0.1",
        "footer": "0.3",
    }
    assert sheet_summary["column_widths"] == ["8.50", "13.00", "28.25", "10.00", "10.00"]
    assert sheet_summary["cell_type_counts"] == {"inlineStr": 8, "number": 1, "b": 1}
    workbook_summary = app.xlsx_structure_summary(workbook_path)
    assert workbook_summary["part_names"] == [
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/app.xml",
        "docProps/core.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
    ]
    assert workbook_summary["part_count"] == 8
    assert workbook_summary["worksheet_part_count"] == 1
    assert workbook_summary["has_core_properties"] is True
    assert workbook_summary["has_extended_properties"] is True
    assert workbook_summary["has_styles"] is True
    assert workbook_summary["content_type_overrides"] == [
        "/docProps/app.xml",
        "/docProps/core.xml",
        "/xl/styles.xml",
        "/xl/workbook.xml",
        "/xl/worksheets/sheet1.xml",
    ]
    assert workbook_summary["root_relationship_type_counts"] == {
        "officeDocument": 1,
        "core-properties": 1,
        "extended-properties": 1,
    }
    assert workbook_summary["workbook_relationship_type_counts"] == {"worksheet": 1, "styles": 1}
    assert workbook_summary["workbook_relationship_targets"] == {
        "rId1": "xl/worksheets/sheet1.xml",
        "rId2": "xl/styles.xml",
    }
    assert workbook_summary["sheet_count"] == 1
    assert workbook_summary["sheet_names"] == ["Highlights"]
    assert workbook_summary["sheets"] == [
        {
            "path": "xl/worksheets/sheet1.xml",
            "dimension": "A1:E2",
            "row_count": 2,
            "column_count": 5,
            "headers": ["ID", "Label", "Highlights", "Score", "Current"],
            "styled_cell_count": 10,
            "cell_type_counts": {"inlineStr": 8, "number": 1, "b": 1},
            "auto_filter_ref": "A1:E2",
            "freeze_pane": {
                "ySplit": "1",
                "topLeftCell": "A2",
                "activePane": "bottomLeft",
                "state": "frozen",
            },
            "page_margins": {
                "left": "0.2",
                "right": "0.4",
                "top": "0.6",
                "bottom": "0.8",
                "header": "0.1",
                "footer": "0.3",
            },
            "column_widths": ["8.50", "13.00", "28.25", "10.00", "10.00"],
            "name": "Highlights",
        }
    ]
    style_summary = workbook_summary["styles"]
    assert style_summary["font_count"] == 3
    assert style_summary["fill_count"] == 4
    assert style_summary["border_count"] == 2
    assert style_summary["cell_xfs_count"] == 3
    assert style_summary["cell_style_count"] == 1
    assert style_summary["font_names"] == ["Calibri", "Calibri", "Calibri"]
    assert style_summary["font_colors"] == ["theme:1", "FF102030", "FF666666"]
    assert style_summary["fill_fg_colors"] == ["FFABCDEF", "FFD9EAF7"]
    assert style_summary["border_colors"] == ["FF405060"]
    assert style_summary["bold_font_indexes"] == [1]
    assert style_summary["italic_font_indexes"] == [2]
    assert style_summary["cell_xfs"] == [
        {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"},
        {
            "numFmtId": "0",
            "fontId": "1",
            "fillId": "2",
            "borderId": "1",
            "xfId": "0",
            "applyFont": "1",
            "applyFill": "1",
            "applyBorder": "1",
        },
        {
            "numFmtId": "0",
            "fontId": "0",
            "fillId": "0",
            "borderId": "1",
            "xfId": "0",
            "applyBorder": "1",
            "applyAlignment": "1",
        },
    ]
    worksheet_xml = first_worksheet_xml(workbook_path)
    assert worksheet_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{app.SHEET_NS}" xmlns:r="{app.REL_NS}">'
        '<dimension ref="A1:E2"/>'
        '<sheetViews><sheetView tabSelected="0" workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="8.50" customWidth="1"/>'
        '<col min="2" max="2" width="13.00" customWidth="1"/>'
        '<col min="3" max="3" width="28.25" customWidth="1"/>'
        '<col min="4" max="4" width="10.00" customWidth="1"/>'
        '<col min="5" max="5" width="10.00" customWidth="1"/></cols>'
        '<sheetData><row r="1">'
        '<c r="A1" s="1" t="inlineStr"><is><t xml:space="preserve">ID</t></is></c>'
        '<c r="B1" s="1" t="inlineStr"><is><t xml:space="preserve">Label</t></is></c>'
        '<c r="C1" s="1" t="inlineStr"><is><t xml:space="preserve">Highlights</t></is></c>'
        '<c r="D1" s="1" t="inlineStr"><is><t xml:space="preserve">Score</t></is></c>'
        '<c r="E1" s="1" t="inlineStr"><is><t xml:space="preserve">Current</t></is></c></row>'
        '<row r="2"><c r="A2" s="2" t="inlineStr"><is><t xml:space="preserve">highlight_001</t></is></c>'
        '<c r="B2" s="2" t="inlineStr"><is><t xml:space="preserve">One</t></is></c>'
        '<c r="C2" s="2" t="inlineStr"><is><t xml:space="preserve">Structured value</t></is></c>'
        '<c r="D2" s="2"><v>98.5</v></c><c r="E2" s="2" t="b"><v>1</v></c></row></sheetData>'
        '<autoFilter ref="A1:E2"/>'
        '<pageMargins left="0.2" right="0.4" top="0.6" bottom="0.8" header="0.1" footer="0.3"/>'
        "</worksheet>"
    )

    resume = {
        "office_layout": {
            "document": {
                "margins": {
                    "top": "1440",
                    "right": "1800",
                    "bottom": "1440",
                    "left": "1800",
                    "header": "720",
                    "footer": "720",
                },
                "page_size": {"w": "12240", "h": "15840", "orient": "portrait"},
            }
        },
        "resume": {"id": "resume_api", "label": "API Resume"},
        "collections": {
            "sections": [
                {
                    "id": "section_001",
                    "title": "API-backed section",
                    "kind": "summary",
                    "sort_order": 1,
                    "item_ids": ["item_001"],
                    "is_visible": True,
                }
            ],
            "items": [
                {
                    "id": "item_001",
                    "section_id": "section_001",
                    "kind": "text",
                    "sort_order": 1,
                    "master_record_refs": [],
                    "text_override": "fallback text",
                    "formatting": {
                        "block_style": "Heading1",
                        "numbering": {"level": "0", "num_id": "7"},
                        "runs": [
                            {"text": "Styled", "bold": True, "italic": False, "underline": False},
                            {"text": " run", "bold": False, "italic": True, "underline": True},
                        ],
                    },
                    "is_visible": True,
                }
            ],
        },
    }
    document_path = tmp_path / "layout_contract.docx"
    app.write_resume_document(master, resume, document_path)
    part_names = docx_part_names(document_path)
    document_xml = docx_document_xml(document_path)
    with zipfile.ZipFile(document_path) as document:
        content_types_xml = document.read("[Content_Types].xml").decode("utf-8")
        root_relationships_xml = document.read("_rels/.rels").decode("utf-8")
        core_properties_xml = document.read("docProps/core.xml").decode("utf-8")
        app_properties_xml = document.read("docProps/app.xml").decode("utf-8")
        docx_relationships_xml = document.read("word/_rels/document.xml.rels").decode("utf-8")
        docx_theme_xml = document.read("word/theme/theme1.xml").decode("utf-8")
        docx_styles_xml = document.read("word/styles.xml").decode("utf-8")
        docx_numbering_xml = document.read("word/numbering.xml").decode("utf-8")
        docx_settings_xml = document.read("word/settings.xml").decode("utf-8")
        docx_font_table_xml = document.read("word/fontTable.xml").decode("utf-8")
        docx_web_settings_xml = document.read("word/webSettings.xml").decode("utf-8")
        docx_footnotes_xml = document.read("word/footnotes.xml").decode("utf-8")
        docx_endnotes_xml = document.read("word/endnotes.xml").decode("utf-8")
        custom_xml = document.read("customXml/item1.xml").decode("utf-8")
        custom_props_xml = document.read("customXml/itemProps1.xml").decode("utf-8")
        custom_relationships_xml = document.read("customXml/_rels/item1.xml.rels").decode("utf-8")
    structure_summary = app.docx_structure_summary(document_path)

    assert part_names == [
        "[Content_Types].xml",
        "_rels/.rels",
        "customXml/_rels/item1.xml.rels",
        "customXml/item1.xml",
        "customXml/itemProps1.xml",
        "docProps/app.xml",
        "docProps/core.xml",
        "word/_rels/document.xml.rels",
        "word/document.xml",
        "word/endnotes.xml",
        "word/fontTable.xml",
        "word/footnotes.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/webSettings.xml",
    ]
    assert structure_summary["part_names"] == part_names
    assert structure_summary["part_count"] == len(part_names)
    assert document_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{app.WORD_NS}" xmlns:r="{app.REL_NS}"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>'
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Styled</w:t></w:r>'
        '<w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr>'
        '<w:t xml:space="preserve"> run</w:t></w:r></w:p><w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
        '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
    )
    assert '<w:pStyle w:val="Heading1"/>' in document_xml
    assert '<w:numId w:val="7"/>' in document_xml
    assert '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>' in document_xml
    assert '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>' in document_xml
    assert (
        '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
        'w:header="720" w:footer="720" w:gutter="0"/>' in document_xml
    )
    assert structure_summary["page_size"] == {"w": "12240", "h": "15840", "orient": "portrait"}
    assert structure_summary["page_margins"] == {
        "top": "1440",
        "right": "1800",
        "bottom": "1440",
        "left": "1800",
        "header": "720",
        "footer": "720",
        "gutter": "0",
    }
    assert "<w:b/>" in document_xml
    assert "<w:i/>" in document_xml
    assert '<w:u w:val="single"/>' in document_xml
    assert core_properties_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>API Resume</dc:title><dc:creator>x_create_cv_x</dc:creator>"
        "<cp:lastModifiedBy>x_create_cv_x</cp:lastModifiedBy>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:modified>'
        "</cp:coreProperties>"
    )
    assert app_properties_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>x_create_cv_x</Application></Properties>"
    )
    assert content_types_xml == (
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
        '<Override PartName="/word/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/word/fontTable.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        '<Override PartName="/word/webSettings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>'
        '<Override PartName="/word/footnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        '<Override PartName="/word/endnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        '<Override PartName="/customXml/itemProps1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )
    assert root_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/><Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/></Relationships>'
    )
    assert docx_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
        'Target="numbering.xml"/><Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
        'Target="settings.xml"/><Relationship Id="rId4" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/><Relationship Id="rId5" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" '
        'Target="fontTable.xml"/><Relationship Id="rId6" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" '
        'Target="webSettings.xml"/><Relationship Id="rId9" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
        'Target="footnotes.xml"/><Relationship Id="rId10" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" '
        'Target="endnotes.xml"/><Relationship Id="rId11" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" '
        'Target="../customXml/item1.xml"/></Relationships>'
    )
    assert docx_settings_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{app.WORD_NS}"><w:defaultTabStop w:val="720"/></w:settings>'
    )
    assert docx_font_table_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts xmlns:w="{app.WORD_NS}"><w:font w:name="Calibri"/><w:font w:name="Symbol"/></w:fonts>'
    )
    assert docx_web_settings_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' f'<w:webSettings xmlns:w="{app.WORD_NS}"/>'
    )
    assert docx_footnotes_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{app.WORD_NS}"><w:footnote w:type="separator" w:id="-1">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
        "<w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>"
    )
    assert docx_endnotes_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:endnotes xmlns:w="{app.WORD_NS}"><w:endnote w:type="separator" w:id="-1">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:endnote>"
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
        "<w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>"
    )
    assert (
        custom_xml
        == '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cv:metadata xmlns:cv="urn:x-create-cv-x"/>'
    )
    assert custom_props_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<ds:datastoreItem ds:itemID="{00000000-0000-0000-0000-000000000001}" '
        'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
        "<ds:schemaRefs/></ds:datastoreItem>"
    )
    assert custom_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" '
        'Target="itemProps1.xml"/></Relationships>'
    )
    expected_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    expected_line_styles = "".join(
        f'<a:ln w="{expected_width}" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        for expected_width in ("6350", "12700", "19050")
    )
    expected_effect_styles = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    expected_background_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    assert docx_theme_xml == (
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
        f"<a:fillStyleLst>{expected_fill_styles}</a:fillStyleLst>"
        f"<a:lnStyleLst>{expected_line_styles}</a:lnStyleLst>"
        f"<a:effectStyleLst>{expected_effect_styles}</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{expected_background_fill_styles}</a:bgFillStyleLst>"
        "</a:fmtScheme></a:themeElements></a:theme>"
    )
    assert docx_styles_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{app.WORD_NS}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/>'
        '<w:basedOn w:val="ListParagraph"/><w:pPr><w:numPr><w:ilvl w:val="0"/>'
        '<w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
        "</w:styles>"
    )
    assert '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Styled</w:t></w:r>' in document_xml
    assert (
        '<w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr>' '<w:t xml:space="preserve"> run</w:t></w:r>' in document_xml
    )
    expected_abstract_num_xml = (
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
    )
    assert docx_numbering_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{app.WORD_NS}">'
        f'<w:abstractNum w:abstractNumId="1">{expected_abstract_num_xml}</w:abstractNum>'
        f'<w:abstractNum w:abstractNumId="7">{expected_abstract_num_xml}</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
        '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>'
        "</w:numbering>"
    )
    assert structure_summary["numbered_paragraph_count"] == 1
    assert structure_summary["numbering_abstract_count"] == 2
    assert structure_summary["numbering_num_count"] == 2
    assert structure_summary["numbering_level_count"] == 4
    assert structure_summary["numbering_abstract_ids"] == ["1", "7"]
    assert structure_summary["numbering_num_ids"] == ["1", "7"]
    assert structure_summary["numbering_level_texts"] == ["\u2022"]
    assert structure_summary["numbering_level_fonts"] == ["Symbol"]
    assert structure_summary["style_definition_count"] == 6
    assert structure_summary["style_definition_ids"] == [
        "Normal",
        "Title",
        "Heading1",
        "Heading2",
        "ListParagraph",
        "ListBullet",
    ]
    assert structure_summary["default_style_ids"] == ["Normal"]
    assert structure_summary["style_based_on"] == {
        "Title": "Normal",
        "Heading1": "Normal",
        "Heading2": "Normal",
        "ListParagraph": "Normal",
        "ListBullet": "ListParagraph",
    }
    assert structure_summary["style_run_fonts"] == {"Normal": "Calibri"}
    assert structure_summary["style_run_sizes"] == {
        "Normal": "22",
        "Title": "32",
        "Heading1": "24",
        "Heading2": "22",
    }
    assert structure_summary["style_bold_ids"] == ["Title", "Heading1", "Heading2"]
    assert structure_summary["style_paragraph_spacing"] == {
        "Title": {"after": "120"},
        "Heading1": {"before": "160", "after": "80"},
        "Heading2": {"before": "120", "after": "60"},
    }
    assert structure_summary["style_paragraph_indents"]["ListParagraph"] == {"left": "720"}
    assert structure_summary["style_numbering"] == {"ListBullet": {"level": "0", "num_id": "1"}}


def test_docx_generation_consumes_flow_and_package_contracts(tmp_path: Path) -> None:
    master: dict[str, Any] = {"collections": {}}
    resume: dict[str, Any] = {
        "office_layout": {
            "document": {
                "package": {
                    "theme": True,
                    "font_table": True,
                    "web_settings": True,
                    "footnotes": True,
                    "endnotes": True,
                    "custom_xml": True,
                    "fonts": ["Calibri", "Symbol", "Courier New"],
                },
                "header": {"runs": [{"text": "Header", "bold": True}]},
                "footer": {"runs": [{"text": "1"}]},
                "flow": [
                    {"type": "items", "item_ids": ["item_intro"]},
                    {
                        "type": "table",
                        "column_widths": ["2400", "2400"],
                        "rows": [
                            [{"item_ids": ["item_left"]}, {"item_ids": ["item_right"]}],
                            [
                                {"item_ids": ["item_bottom"], "paragraphs": [{"empty": True}]},
                                {
                                    "runs": [
                                        {
                                            "text": "Inline cell",
                                            "italic": True,
                                            "font": "Courier New",
                                            "size": "22",
                                            "color": "0563C1",
                                            "link": "https://example.test/cv",
                                        },
                                        {"tab": True},
                                        {"text": "after tab"},
                                    ]
                                },
                            ],
                        ],
                    },
                ],
            }
        },
        "resume": {"id": "resume_api", "label": "API Resume"},
        "collections": {
            "sections": [
                {
                    "id": "section_001",
                    "title": "API-backed section",
                    "kind": "summary",
                    "sort_order": 1,
                    "item_ids": ["item_intro", "item_left", "item_right", "item_bottom"],
                    "is_visible": True,
                }
            ],
            "items": [
                {
                    "id": "item_intro",
                    "section_id": "section_001",
                    "kind": "text",
                    "sort_order": 1,
                    "master_record_refs": [],
                    "text_override": "Intro",
                    "formatting": {
                        "block_style": "Title",
                        "alignment": "center",
                        "spacing": {"after": "0", "line": "240", "line_rule": "auto"},
                        "indent": {"left": "720", "hanging": "360"},
                        "tab_stops": [{"val": "right", "leader": "none", "pos": "10800"}],
                        "numbering": None,
                        "runs": [{"text": "Intro"}],
                    },
                    "is_visible": True,
                },
                {
                    "id": "item_left",
                    "section_id": "section_001",
                    "kind": "text",
                    "sort_order": 2,
                    "master_record_refs": [],
                    "text_override": "Left",
                    "formatting": {"block_style": None, "numbering": None, "runs": [{"text": "Left"}]},
                    "is_visible": True,
                },
                {
                    "id": "item_right",
                    "section_id": "section_001",
                    "kind": "text",
                    "sort_order": 3,
                    "master_record_refs": [],
                    "text_override": "Right",
                    "formatting": {"block_style": None, "numbering": None, "runs": [{"text": "Right"}]},
                    "is_visible": True,
                },
                {
                    "id": "item_bottom",
                    "section_id": "section_001",
                    "kind": "text",
                    "sort_order": 4,
                    "master_record_refs": [],
                    "text_override": "Bottom",
                    "formatting": {"block_style": None, "numbering": None, "runs": [{"text": "Bottom"}]},
                    "is_visible": True,
                },
            ],
        },
    }
    app.validate_contract(object_at(object_at(resume, "office_layout"), "document"), "document_layout")
    document_path = tmp_path / "flow_contract.docx"
    app.write_resume_document(master, resume, document_path)
    part_names = docx_part_names(document_path)
    document_xml = docx_document_xml(document_path)
    relationships_xml = docx_relationships_xml(document_path)
    with zipfile.ZipFile(document_path) as document:
        content_types_xml = document.read("[Content_Types].xml").decode("utf-8")
        root_relationships_xml = document.read("_rels/.rels").decode("utf-8")
        theme_xml = document.read("word/theme/theme1.xml").decode("utf-8")
        styles_xml = document.read("word/styles.xml").decode("utf-8")
        numbering_xml = document.read("word/numbering.xml").decode("utf-8")
        settings_xml = document.read("word/settings.xml").decode("utf-8")
        font_table_xml = document.read("word/fontTable.xml").decode("utf-8")
        web_settings_xml = document.read("word/webSettings.xml").decode("utf-8")
        footnotes_xml = document.read("word/footnotes.xml").decode("utf-8")
        endnotes_xml = document.read("word/endnotes.xml").decode("utf-8")
        custom_xml = document.read("customXml/item1.xml").decode("utf-8")
        custom_props_xml = document.read("customXml/itemProps1.xml").decode("utf-8")
        custom_relationships_xml = document.read("customXml/_rels/item1.xml.rels").decode("utf-8")
        header_xml = document.read("word/header1.xml").decode("utf-8")
        footer_xml = document.read("word/footer1.xml").decode("utf-8")
    structure_summary = app.docx_structure_summary(document_path)

    assert part_names == [
        "[Content_Types].xml",
        "_rels/.rels",
        "customXml/_rels/item1.xml.rels",
        "customXml/item1.xml",
        "customXml/itemProps1.xml",
        "word/_rels/document.xml.rels",
        "word/document.xml",
        "word/endnotes.xml",
        "word/fontTable.xml",
        "word/footer1.xml",
        "word/footnotes.xml",
        "word/header1.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/webSettings.xml",
    ]
    assert structure_summary["part_names"] == part_names
    assert structure_summary["part_count"] == len(part_names)
    assert "word/theme/theme1.xml" in part_names
    assert "word/settings.xml" in part_names
    assert "word/fontTable.xml" in part_names
    assert "word/webSettings.xml" in part_names
    assert "word/footnotes.xml" in part_names
    assert "word/endnotes.xml" in part_names
    assert "customXml/item1.xml" in part_names
    assert "customXml/itemProps1.xml" in part_names
    assert "customXml/_rels/item1.xml.rels" in part_names
    assert "word/header1.xml" in part_names
    assert "word/footer1.xml" in part_names
    assert content_types_xml == (
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
        '<Override PartName="/word/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/word/fontTable.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        '<Override PartName="/word/webSettings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>'
        '<Override PartName="/word/footnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        '<Override PartName="/word/endnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        '<Override PartName="/customXml/itemProps1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
        '<Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        "</Types>"
    )
    assert (
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>' in content_types_xml
    )
    assert (
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' in content_types_xml
    )
    assert (
        '<Override PartName="/word/numbering.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/word/settings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/word/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' in content_types_xml
    )
    assert (
        '<Override PartName="/word/fontTable.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/word/webSettings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/word/footnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/word/endnotes.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        in content_types_xml
    )
    assert (
        '<Override PartName="/customXml/itemProps1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' in content_types_xml
    )
    assert (
        '<Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' in content_types_xml
    )
    assert (
        '<Override PartName="/word/footer1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' in content_types_xml
    )
    expected_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    expected_line_styles = "".join(
        f'<a:ln w="{expected_width}" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        for expected_width in ("6350", "12700", "19050")
    )
    expected_effect_styles = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    expected_background_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    assert theme_xml == (
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
        f"<a:fillStyleLst>{expected_fill_styles}</a:fillStyleLst>"
        f"<a:lnStyleLst>{expected_line_styles}</a:lnStyleLst>"
        f"<a:effectStyleLst>{expected_effect_styles}</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{expected_background_fill_styles}</a:bgFillStyleLst>"
        "</a:fmtScheme></a:themeElements></a:theme>"
    )
    assert root_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    )
    assert (
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"' in root_relationships_xml
    )
    assert "core-properties" not in root_relationships_xml
    assert "extended-properties" not in root_relationships_xml
    expected_abstract_num_xml = (
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
    )
    assert numbering_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{app.WORD_NS}">'
        f'<w:abstractNum w:abstractNumId="1">{expected_abstract_num_xml}</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
    )
    assert settings_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{app.WORD_NS}"><w:defaultTabStop w:val="720"/></w:settings>'
    )
    assert font_table_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts xmlns:w="{app.WORD_NS}">'
        '<w:font w:name="Calibri"/><w:font w:name="Symbol"/><w:font w:name="Courier New"/></w:fonts>'
    )
    assert web_settings_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' f'<w:webSettings xmlns:w="{app.WORD_NS}"/>'
    )
    assert '<w:font w:name="Calibri"/><w:font w:name="Symbol"/><w:font w:name="Courier New"/>' in font_table_xml
    assert f'<w:webSettings xmlns:w="{app.WORD_NS}"/>' in web_settings_xml
    assert footnotes_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{app.WORD_NS}"><w:footnote w:type="separator" w:id="-1">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
        "<w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>"
    )
    assert endnotes_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:endnotes xmlns:w="{app.WORD_NS}"><w:endnote w:type="separator" w:id="-1">'
        "<w:p><w:r><w:separator/></w:r></w:p></w:endnote>"
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r>'
        "<w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotes>"
    )
    assert '<w:footnote w:type="separator" w:id="-1">' in footnotes_xml
    assert "<w:continuationSeparator/>" in footnotes_xml
    assert '<w:endnote w:type="separator" w:id="-1">' in endnotes_xml
    assert "<w:continuationSeparator/>" in endnotes_xml
    assert (
        custom_xml
        == '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cv:metadata xmlns:cv="urn:x-create-cv-x"/>'
    )
    assert custom_props_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<ds:datastoreItem ds:itemID="{00000000-0000-0000-0000-000000000001}" '
        'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
        "<ds:schemaRefs/></ds:datastoreItem>"
    )
    assert custom_relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" '
        'Target="itemProps1.xml"/></Relationships>'
    )
    assert '<cv:metadata xmlns:cv="urn:x-create-cv-x"/>' in custom_xml
    assert 'ds:itemID="{00000000-0000-0000-0000-000000000001}"' in custom_props_xml
    assert "<ds:schemaRefs/>" in custom_props_xml
    assert (
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" '
        'Target="itemProps1.xml"' in custom_relationships_xml
    )
    assert header_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{app.WORD_NS}"><w:p><w:r><w:rPr><w:b/></w:rPr>'
        '<w:t xml:space="preserve">Header</w:t></w:r></w:p></w:hdr>'
    )
    assert footer_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{app.WORD_NS}"><w:p><w:r><w:t xml:space="preserve">1</w:t></w:r></w:p></w:ftr>'
    )
    assert '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Header</w:t></w:r>' in header_xml
    assert '<w:r><w:t xml:space="preserve">1</w:t></w:r>' in footer_xml
    assert relationships_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
        'Target="numbering.xml"/><Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
        'Target="settings.xml"/><Relationship Id="rId4" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/><Relationship Id="rId5" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" '
        'Target="fontTable.xml"/><Relationship Id="rId6" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" '
        'Target="webSettings.xml"/><Relationship Id="rId7" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header1.xml"/><Relationship Id="rId8" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer1.xml"/><Relationship Id="rId9" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
        'Target="footnotes.xml"/><Relationship Id="rId10" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" '
        'Target="endnotes.xml"/><Relationship Id="rId11" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" '
        'Target="../customXml/item1.xml"/><Relationship Id="rId101" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        'Target="https://example.test/cv" TargetMode="External"/></Relationships>'
    )
    expected_relationships = [
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"',
        'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
        'Target="numbering.xml"',
        'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
        'Target="settings.xml"',
        'Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"',
        'Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" '
        'Target="fontTable.xml"',
        'Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" '
        'Target="webSettings.xml"',
        'Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header1.xml"',
        'Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer1.xml"',
        'Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
        'Target="footnotes.xml"',
        'Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" '
        'Target="endnotes.xml"',
        'Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" '
        'Target="../customXml/item1.xml"',
        'Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        'Target="https://example.test/cv" TargetMode="External"',
    ]
    for expected_relationship in expected_relationships:
        assert expected_relationship in relationships_xml
    assert structure_summary["has_theme"] is True
    assert structure_summary["has_font_table"] is True
    assert structure_summary["has_web_settings"] is True
    assert structure_summary["has_footnotes"] is True
    assert structure_summary["has_endnotes"] is True
    assert structure_summary["has_custom_xml"] is True
    assert structure_summary["header_count"] == 1
    assert structure_summary["footer_count"] == 1
    assert structure_summary["body_child_counts"] == {"p": 1, "tbl": 1, "sectPr": 1}
    assert structure_summary["font_names"] == ["Calibri", "Symbol", "Courier New"]
    assert structure_summary["style_definition_count"] == 6
    assert structure_summary["style_definition_ids"] == [
        "Normal",
        "Title",
        "Heading1",
        "Heading2",
        "ListParagraph",
        "ListBullet",
    ]
    assert structure_summary["default_style_ids"] == ["Normal"]
    assert structure_summary["style_based_on"] == {
        "Title": "Normal",
        "Heading1": "Normal",
        "Heading2": "Normal",
        "ListParagraph": "Normal",
        "ListBullet": "ListParagraph",
    }
    assert structure_summary["style_run_fonts"] == {"Normal": "Calibri"}
    assert structure_summary["style_run_sizes"] == {
        "Normal": "22",
        "Title": "32",
        "Heading1": "24",
        "Heading2": "22",
    }
    assert structure_summary["style_bold_ids"] == ["Title", "Heading1", "Heading2"]
    assert structure_summary["style_paragraph_spacing"] == {
        "Title": {"after": "120"},
        "Heading1": {"before": "160", "after": "80"},
        "Heading2": {"before": "120", "after": "60"},
    }
    assert structure_summary["style_paragraph_indents"] == {"ListParagraph": {"left": "720"}}
    assert structure_summary["style_numbering"] == {"ListBullet": {"level": "0", "num_id": "1"}}
    assert structure_summary["styles"] == ["Title"]
    assert styles_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{app.WORD_NS}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/>'
        '<w:basedOn w:val="ListParagraph"/><w:pPr><w:numPr><w:ilvl w:val="0"/>'
        '<w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
        "</w:styles>"
    )
    assert document_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{app.WORD_NS}" xmlns:r="{app.REL_NS}"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Title"/>'
        '<w:tabs><w:tab w:val="right" w:leader="none" w:pos="10800"/></w:tabs>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:left="720" w:hanging="360"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">Intro</w:t></w:r></w:p>'
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t xml:space="preserve">Left</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t xml:space="preserve">Right</w:t></w:r></w:p></w:tc></w:tr>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t xml:space="preserve">Bottom</w:t></w:r></w:p><w:p><w:r></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p>'
        '<w:hyperlink r:id="rId101" w:history="1"><w:r><w:rPr>'
        '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
        '<w:i/><w:color w:val="0563C1"/><w:sz w:val="22"/></w:rPr>'
        '<w:t xml:space="preserve">Inline cell</w:t></w:r></w:hyperlink>'
        '<w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">after tab</w:t></w:r>'
        "</w:p></w:tc></w:tr></w:tbl><w:sectPr>"
        '<w:headerReference w:type="default" r:id="rId7"/>'
        '<w:footerReference w:type="default" r:id="rId8"/>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
    )
    assert document_xml.startswith(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{app.WORD_NS}" xmlns:r="{app.REL_NS}"><w:body>'
    )
    assert document_xml.endswith("</w:body></w:document>")
    assert '<w:tabs><w:tab w:val="right" w:leader="none" w:pos="10800"/></w:tabs>' in document_xml
    assert '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' in document_xml
    assert '<w:ind w:left="720" w:hanging="360"/>' in document_xml
    assert '<w:jc w:val="center"/>' in document_xml
    assert (
        '<w:pPr><w:pStyle w:val="Title"/><w:tabs><w:tab w:val="right" w:leader="none" w:pos="10800"/>'
        '</w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:left="720" w:hanging="360"/><w:jc w:val="center"/></w:pPr>' in document_xml
    )
    assert docx_theme_style_counts(document_path) == {
        "fillStyleLst": 3,
        "lnStyleLst": 3,
        "effectStyleLst": 3,
        "bgFillStyleLst": 3,
    }
    assert '<w:hyperlink r:id="rId101" w:history="1">' in document_xml
    assert (
        '<w:hyperlink r:id="rId101" w:history="1"><w:r><w:rPr>'
        '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
        '<w:i/><w:color w:val="0563C1"/><w:sz w:val="22"/></w:rPr>'
        '<w:t xml:space="preserve">Inline cell</w:t></w:r></w:hyperlink>' in document_xml
    )
    assert "<w:tab/>" in document_xml
    assert '<w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">after tab</w:t></w:r>' in document_xml
    assert '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>' in document_xml
    assert '<w:color w:val="0563C1"/>' in document_xml
    assert '<w:sz w:val="22"/>' in document_xml
    assert 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"' in relationships_xml
    assert 'Target="https://example.test/cv" TargetMode="External"' in relationships_xml
    assert structure_summary["relationship_count"] == 12
    assert structure_summary["relationship_type_counts"] == {
        "styles": 1,
        "numbering": 1,
        "settings": 1,
        "theme": 1,
        "fontTable": 1,
        "webSettings": 1,
        "header": 1,
        "footer": 1,
        "footnotes": 1,
        "endnotes": 1,
        "customXml": 1,
        "hyperlink": 1,
    }
    assert structure_summary["relationship_type_counts"]["hyperlink"] == 1
    assert structure_summary["relationship_target_mode_counts"] == {"External": 1}
    assert structure_summary["external_relationship_count"] == 1
    assert structure_summary["external_hyperlink_relationship_count"] == 1
    assert structure_summary["styled_paragraph_count"] == 1
    assert structure_summary["aligned_paragraph_count"] == 1
    assert structure_summary["spaced_paragraph_count"] == 1
    assert structure_summary["indented_paragraph_count"] == 1
    assert structure_summary["tab_stopped_paragraph_count"] == 1
    assert {
        summary_key: structure_summary[summary_key]
        for summary_key in [
            "paragraph_count",
            "run_count",
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
            "numbered_paragraph_count",
        ]
    } == {
        "paragraph_count": 6,
        "run_count": 8,
        "bold_run_count": 0,
        "italic_run_count": 1,
        "underline_run_count": 0,
        "colored_run_count": 1,
        "fonted_run_count": 1,
        "hyperlink_count": 1,
        "tab_count": 2,
        "table_count": 1,
        "table_row_count": 2,
        "table_cell_count": 4,
        "table_paragraph_count": 5,
        "numbered_paragraph_count": 0,
    }
    assert structure_summary["numbering_abstract_count"] == 1
    assert structure_summary["numbering_num_count"] == 1
    assert structure_summary["numbering_level_count"] == 2
    assert structure_summary["numbering_abstract_ids"] == ["1"]
    assert structure_summary["numbering_num_ids"] == ["1"]
    assert structure_summary["numbering_level_texts"] == ["\u2022"]
    assert structure_summary["numbering_level_fonts"] == ["Symbol"]
    assert "<w:tbl>" in document_xml
    assert (
        '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>' in document_xml
    )
    assert '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' in document_xml
    assert document_xml.count('<w:tcW w:w="2400" w:type="dxa"/>') == 4
    assert (
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t xml:space="preserve">Left</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
        '<w:p><w:r><w:t xml:space="preserve">Right</w:t></w:r></w:p></w:tc></w:tr>' in document_xml
    )
    assert '<w:p><w:r><w:t xml:space="preserve">Bottom</w:t></w:r></w:p><w:p><w:r></w:r></w:p>' in document_xml
    assert document_xml.count("<w:p>") == 6
    assert document_xml.count("<w:tr>") == 2
    assert document_xml.count("<w:tc>") == 4
    assert structure_summary["table_grid_widths"] == [["2400", "2400"]]
    assert structure_summary["table_cell_widths"] == [[["2400", "2400"], ["2400", "2400"]]]
    assert '<w:pgSz w:w="12240" w:h="15840"/>' in document_xml
    assert (
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>' in document_xml
    )
    assert '<w:headerReference w:type="default" r:id="rId7"/>' in document_xml
    assert '<w:footerReference w:type="default" r:id="rId8"/>' in document_xml
    assert structure_summary["page_size"] == {"w": "12240", "h": "15840"}
    assert structure_summary["page_margins"] == {
        "top": "1440",
        "right": "1440",
        "bottom": "1440",
        "left": "1440",
        "header": "720",
        "footer": "720",
        "gutter": "0",
    }
    assert structure_summary["section_header_references"] == [{"type": "default", "id": "rId7"}]
    assert structure_summary["section_footer_references"] == [{"type": "default", "id": "rId8"}]


def test_docx_generation_can_omit_optional_package_properties(tmp_path: Path) -> None:
    master: dict[str, Any] = {"collections": {}}
    resume = {
        "office_layout": {
            "document": {
                "package": {
                    "core_properties": False,
                    "extended_properties": False,
                    "theme": True,
                    "font_table": True,
                    "web_settings": False,
                    "footnotes": False,
                    "endnotes": False,
                    "custom_xml": False,
                },
                "flow": [{"type": "paragraph", "runs": [{"text": "Lean package"}]}],
            }
        },
        "resume": {"id": "resume_lean", "label": "Lean Resume"},
        "collections": {"sections": [], "items": []},
    }
    document_path = tmp_path / "lean.docx"

    app.write_resume_document(master, resume, document_path)

    part_names = docx_part_names(document_path)
    structure_summary = app.docx_structure_summary(document_path)
    assert part_names == [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/_rels/document.xml.rels",
        "word/document.xml",
        "word/fontTable.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
    ]
    assert "docProps/core.xml" not in part_names
    assert "docProps/app.xml" not in part_names
    assert "word/webSettings.xml" not in part_names
    assert "word/footnotes.xml" not in part_names
    assert "word/endnotes.xml" not in part_names
    assert "customXml/item1.xml" not in part_names
    assert "customXml/itemProps1.xml" not in part_names
    assert "customXml/_rels/item1.xml.rels" not in part_names
    assert "word/header1.xml" not in part_names
    assert "word/footer1.xml" not in part_names
    assert structure_summary["part_names"] == part_names
    assert structure_summary["part_count"] == len(part_names)
    assert structure_summary["has_theme"] is True
    assert structure_summary["has_font_table"] is True
    assert structure_summary["has_web_settings"] is False
    assert structure_summary["has_footnotes"] is False
    assert structure_summary["has_endnotes"] is False
    assert structure_summary["has_custom_xml"] is False
    assert structure_summary["header_count"] == 0
    assert structure_summary["footer_count"] == 0
    assert structure_summary["page_size"] == {"w": "12240", "h": "15840"}
    assert structure_summary["page_margins"] == {
        "top": "1440",
        "right": "1440",
        "bottom": "1440",
        "left": "1440",
        "header": "720",
        "footer": "720",
        "gutter": "0",
    }
    assert structure_summary["section_header_references"] == []
    assert structure_summary["section_footer_references"] == []
    assert structure_summary["body_child_counts"] == {"p": 1, "sectPr": 1}
    assert structure_summary["font_names"] == ["Calibri", "Symbol"]
    assert structure_summary["relationship_count"] == 5
    assert structure_summary["relationship_type_counts"] == {
        "styles": 1,
        "numbering": 1,
        "settings": 1,
        "theme": 1,
        "fontTable": 1,
    }
    assert structure_summary["relationship_target_mode_counts"] == {}
    assert structure_summary["external_relationship_count"] == 0
    assert structure_summary["external_hyperlink_relationship_count"] == 0
    assert structure_summary["paragraph_count"] == 1
    assert structure_summary["run_count"] == 1
    assert structure_summary["hyperlink_count"] == 0
    assert structure_summary["tab_count"] == 0
    assert structure_summary["table_count"] == 0
    assert structure_summary["table_row_count"] == 0
    assert structure_summary["table_cell_count"] == 0
    assert structure_summary["table_paragraph_count"] == 0
    assert structure_summary["table_grid_widths"] == []
    assert structure_summary["table_cell_widths"] == []
    assert structure_summary["numbered_paragraph_count"] == 0
    assert structure_summary["numbering_abstract_count"] == 1
    assert structure_summary["numbering_num_count"] == 1
    assert structure_summary["numbering_level_count"] == 2
    assert structure_summary["numbering_abstract_ids"] == ["1"]
    assert structure_summary["numbering_num_ids"] == ["1"]
    assert structure_summary["numbering_level_texts"] == ["\u2022"]
    assert structure_summary["numbering_level_fonts"] == ["Symbol"]
    assert structure_summary["styled_paragraph_count"] == 0
    assert structure_summary["aligned_paragraph_count"] == 0
    assert structure_summary["spaced_paragraph_count"] == 0
    assert structure_summary["indented_paragraph_count"] == 0
    assert structure_summary["tab_stopped_paragraph_count"] == 0
    assert structure_summary["bold_run_count"] == 0
    assert structure_summary["italic_run_count"] == 0
    assert structure_summary["underline_run_count"] == 0
    assert structure_summary["colored_run_count"] == 0
    assert structure_summary["fonted_run_count"] == 0
    assert structure_summary["style_definition_count"] == 6
    assert structure_summary["style_definition_ids"] == [
        "Normal",
        "Title",
        "Heading1",
        "Heading2",
        "ListParagraph",
        "ListBullet",
    ]
    assert structure_summary["default_style_ids"] == ["Normal"]
    assert structure_summary["style_based_on"] == {
        "Title": "Normal",
        "Heading1": "Normal",
        "Heading2": "Normal",
        "ListParagraph": "Normal",
        "ListBullet": "ListParagraph",
    }
    assert structure_summary["style_run_fonts"] == {"Normal": "Calibri"}
    assert structure_summary["style_run_sizes"] == {
        "Normal": "22",
        "Title": "32",
        "Heading1": "24",
        "Heading2": "22",
    }
    assert structure_summary["style_bold_ids"] == ["Title", "Heading1", "Heading2"]
    assert structure_summary["style_paragraph_spacing"] == {
        "Title": {"after": "120"},
        "Heading1": {"before": "160", "after": "80"},
        "Heading2": {"before": "120", "after": "60"},
    }
    assert structure_summary["style_paragraph_indents"] == {"ListParagraph": {"left": "720"}}
    assert structure_summary["style_numbering"] == {"ListBullet": {"level": "0", "num_id": "1"}}
    assert structure_summary["styles"] == []
    with zipfile.ZipFile(document_path) as document:
        root_relationships = document.read("_rels/.rels").decode("utf-8")
        document_relationships = document.read("word/_rels/document.xml.rels").decode("utf-8")
        document_xml = document.read("word/document.xml").decode("utf-8")
        numbering_xml = document.read("word/numbering.xml").decode("utf-8")
        styles_xml = document.read("word/styles.xml").decode("utf-8")
        font_table_xml = document.read("word/fontTable.xml").decode("utf-8")
        settings_xml = document.read("word/settings.xml").decode("utf-8")
        theme_xml = document.read("word/theme/theme1.xml").decode("utf-8")
        content_types = document.read("[Content_Types].xml").decode("utf-8")
    assert root_relationships == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    )
    assert document_relationships == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
        'Target="numbering.xml"/><Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
        'Target="settings.xml"/><Relationship Id="rId4" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/><Relationship Id="rId5" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" '
        'Target="fontTable.xml"/></Relationships>'
    )
    assert document_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{app.WORD_NS}" xmlns:r="{app.REL_NS}"><w:body>'
        '<w:p><w:r><w:t xml:space="preserve">Lean package</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
    )
    expected_abstract_num_xml = (
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
    )
    assert numbering_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{app.WORD_NS}">'
        f'<w:abstractNum w:abstractNumId="1">{expected_abstract_num_xml}</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
    )
    assert styles_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{app.WORD_NS}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/>'
        '<w:basedOn w:val="ListParagraph"/><w:pPr><w:numPr><w:ilvl w:val="0"/>'
        '<w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
        "</w:styles>"
    )
    assert font_table_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts xmlns:w="{app.WORD_NS}"><w:font w:name="Calibri"/><w:font w:name="Symbol"/></w:fonts>'
    )
    assert settings_xml == (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{app.WORD_NS}"><w:defaultTabStop w:val="720"/></w:settings>'
    )
    expected_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    expected_line_styles = "".join(
        f'<a:ln w="{expected_width}" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        for expected_width in ("6350", "12700", "19050")
    )
    expected_effect_styles = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    expected_background_fill_styles = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    assert theme_xml == (
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
        f"<a:fillStyleLst>{expected_fill_styles}</a:fillStyleLst>"
        f"<a:lnStyleLst>{expected_line_styles}</a:lnStyleLst>"
        f"<a:effectStyleLst>{expected_effect_styles}</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{expected_background_fill_styles}</a:bgFillStyleLst>"
        "</a:fmtScheme></a:themeElements></a:theme>"
    )
    assert content_types == (
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
        '<Override PartName="/word/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/word/fontTable.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        "</Types>"
    )
    assert "/docProps/core.xml" not in content_types
    assert "/docProps/app.xml" not in content_types
    assert "/word/webSettings.xml" not in content_types
    assert "/word/footnotes.xml" not in content_types
    assert "/word/endnotes.xml" not in content_types
    assert "/customXml/itemProps1.xml" not in content_types
    assert "/word/header1.xml" not in content_types
    assert "/word/footer1.xml" not in content_types


def test_cli_exercise_golden_uses_evidence_dir(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    evidence_dir = tmp_path / "evidence"
    evidence_file = evidence_dir / "source_office" / "a_priori" / "fake_a_priori.docx"
    script_dir = evidence_dir / "scripts"
    expected_json_dir = evidence_dir / "generated" / "json" / "a_posteriori"
    evidence_file.parent.mkdir(parents=True)
    script_dir.mkdir(parents=True)
    expected_json_dir.mkdir(parents=True)
    evidence_file.write_bytes(b"fake office package")
    master_contract = app.empty_master(
        {
            "id": "profile_fake",
            "person_id": "person_fake",
            "label": "Fake Golden Profile",
            "created_at": "2026-07-01T00:00:00Z",
            "updated_at": "2026-07-01T00:00:00Z",
        }
    )
    resume_contracts = {
        resume_id: app.empty_resume(
            {
                "id": resume_id,
                "label": f"Fake {resume_id}",
                "status": "active",
                "created_at": "2026-07-01T00:00:00Z",
                "updated_at": "2026-07-01T00:00:00Z",
            }
        )
        for resume_id in ["resume_2017", "resume_2023", "resume_2024"]
    }
    expected_json = {
        "master_profile.json": (master_contract, "master_profile_a_posteriori.json"),
        "resume_2017.json": (resume_contracts["resume_2017"], "resume_2017_a_posteriori.json"),
        "resume_2023.json": (resume_contracts["resume_2023"], "resume_2023_a_posteriori.json"),
        "resume_2024.json": (resume_contracts["resume_2024"], "resume_2024_a_posteriori.json"),
    }
    script_outputs = [
        ("01_build_master_data.py", "master_profile.json", master_contract),
        ("02_build_old_resume.py", "resume_2017.json", resume_contracts["resume_2017"]),
        ("03_build_new_resume.py", "resume_2023.json", resume_contracts["resume_2023"]),
        ("04_build_current_resume.py", "resume_2024.json", resume_contracts["resume_2024"]),
    ]
    for script_name, output_name, output_json in script_outputs:
        (script_dir / script_name).write_text(
            "from pathlib import Path\n"
            "import os\n"
            "import x_create_cv_factory_x as app\n"
            "out = Path(os.environ['CV_FACTORY_DB_DIR'])\n"
            f"app.write_json(out / {output_name!r}, {output_json!r})\n",
            encoding="utf-8",
        )
    for _generated_name, (content, expected_name) in expected_json.items():
        app.write_json(expected_json_dir / expected_name, content)
    generated_office_count, report_count = app.write_golden_office_outputs(evidence_dir, write_report=False)
    assert generated_office_count == 4
    assert report_count == 0
    assert workbook_sheet_names(evidence_dir / app.GOLDEN_OFFICE_EXPECTATIONS["master_profile_a_posteriori.xlsx"]) == [
        str(sheet["name"]) for sheet in app.default_workbook_layout()["sheets"]
    ]
    assert (
        docx_section_margins(evidence_dir / app.GOLDEN_OFFICE_EXPECTATIONS["resume_2023_a_posteriori.docx"])["right"]
        == "1800"
    )
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
        "Golden exercise passed: 1 a priori evidence files; 14 chain files; 4 generated JSON files; "
        "4 generated Office files" in capsys.readouterr().out
    )


def test_cli_audit_writes_human_readable_office_report(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    evidence_dir = tmp_path / "evidence"
    expected_json_dir = evidence_dir / "generated" / "json" / "a_posteriori"
    expected_json_dir.mkdir(parents=True)
    app.write_json(
        expected_json_dir / "master_profile_a_posteriori.json",
        app.empty_master(
            {
                "id": "profile_fake",
                "person_id": "person_fake",
                "label": "Fake Golden Profile",
                "created_at": "2026-07-01T00:00:00Z",
                "updated_at": "2026-07-01T00:00:00Z",
            }
        ),
    )
    for resume_id in ["resume_2017", "resume_2023", "resume_2024"]:
        app.write_json(
            expected_json_dir / f"{resume_id}_a_posteriori.json",
            app.empty_resume(
                {
                    "id": resume_id,
                    "label": f"Fake {resume_id}",
                    "status": "active",
                    "created_at": "2026-07-01T00:00:00Z",
                    "updated_at": "2026-07-01T00:00:00Z",
                }
            ),
        )
    app.write_golden_office_outputs(evidence_dir, write_report=False)
    for generated_name, source_relative_path in app.GOLDEN_SOURCE_OFFICE.items():
        generated_path = evidence_dir / app.GOLDEN_OFFICE_EXPECTATIONS[generated_name]
        source_path = evidence_dir / source_relative_path
        source_path.parent.mkdir(parents=True, exist_ok=True)
        source_path.write_bytes(generated_path.read_bytes())

    comparisons = app.office_comparisons(evidence_dir)
    assert [comparison["status"] for comparison in comparisons] == ["pass"] * 4
    assert app.office_report_name("generated\\office\\a_posteriori\\resume.docx") == "resume.docx"

    assert app.main(["audit", "--evidence-dir", str(evidence_dir)]) == 0

    output = capsys.readouterr().out
    audit_path = evidence_dir / app.GOLDEN_OFFICE_AUDIT_REPORT
    json_path = evidence_dir / app.GOLDEN_OFFICE_REPORT
    assert "Office audit written" in output
    assert audit_path.exists()
    assert json_path.exists()
    report = app.read_json(json_path)
    assert report["schema_version"] == app.SCHEMA_VERSION
    assert report["generator"] == f"x_create_cv_x {app.VERSION}"
    assert report["purpose"] == "Private comparison report for generated _a_posteriori Office evidence"
    assert report["audit_policy"] == {
        "schema_version": app.SCHEMA_VERSION,
        "policy_version": "0.0.2",
        "name": "default_office_audit_policy",
        "accepted_difference_count": 1,
        "path": str(app.DEFAULT_AUDIT_POLICY).replace("\\", "/"),
    }
    assert [comparison["status"] for comparison in report["comparisons"]] == ["pass"] * 4
    assert [comparison["byte_identical"] for comparison in report["comparisons"]] == [True] * 4
    assert [comparison["normalized_text_match"] for comparison in report["comparisons"]] == [True] * 4
    for comparison in report["comparisons"]:
        assert comparison["generated"]["normalized_text"] == comparison["source"]["normalized_text"]
        assert comparison["generated"]["structure"] == comparison["source"]["structure"]
    assert report["comparisons"][0]["generated"]["normalized_text"]["line_count"] == 9
    assert report["comparisons"][1]["generated"]["normalized_text"]["line_count"] == 1
    assert report["comparisons"][0]["generated"]["structure"]["sheet_count"] == 9
    assert report["comparisons"][1]["generated"]["structure"]["style_definition_count"] == 6
    audit_text = audit_path.read_text(encoding="utf-8")
    assert "# A Posteriori Office Audit" in audit_text
    assert f"Generator: `x_create_cv_x {app.VERSION}`" in audit_text
    assert "Audit policy: `default_office_audit_policy` version `0.0.2`" in audit_text
    assert (
        "| master_profile_a_posteriori.xlsx | .xlsx | true | true | pass | "
        "Generated file is byte-identical to source evidence. | 9 | 9 |" in audit_text
    )
    assert (
        "| resume_2017_a_posteriori.docx | .docx | true | true | pass | "
        "Generated file is byte-identical to source evidence. | 1 | 1 |" in audit_text
    )
    assert (
        "| resume_2023_a_posteriori.docx | .docx | true | true | pass | "
        "Generated file is byte-identical to source evidence. | 1 | 1 |" in audit_text
    )
    assert (
        "| resume_2024_a_posteriori.docx | .docx | true | true | pass | "
        "Generated file is byte-identical to source evidence. | 1 | 1 |" in audit_text
    )
    assert "| style_definition_count | 6 | 6 |" in audit_text
    assert "| sheet_count | 9 | 9 |" in audit_text
    assert "## DOCX Structure" in audit_text
    assert "## XLSX Structure" in audit_text
    assert "### resume_2017_a_posteriori.docx" in audit_text
    assert "### resume_2023_a_posteriori.docx" in audit_text
    assert "### resume_2024_a_posteriori.docx" in audit_text
    resume_2017_docx_section = audit_text.split("### resume_2017_a_posteriori.docx", 1)[1].split(
        "### resume_2023_a_posteriori.docx", 1
    )[0]
    resume_2023_docx_section = audit_text.split("### resume_2023_a_posteriori.docx", 1)[1].split(
        "### resume_2024_a_posteriori.docx", 1
    )[0]
    resume_2024_docx_section = audit_text.split("### resume_2024_a_posteriori.docx", 1)[1].split(
        "### master_profile_a_posteriori.xlsx", 1
    )[0]
    assert (
        "### resume_2017_a_posteriori.docx\n\n"
        "| Metric | Source | Generated |\n"
        "| --- | ---: | ---: |\n"
        "| part_names |" in audit_text
    )
    assert "| style_definition_count | 6 | 6 |" in resume_2017_docx_section
    assert "| part_count | 17 | 17 |" in resume_2017_docx_section
    assert "| has_theme | true | true |" in resume_2017_docx_section
    assert "| has_font_table | true | true |" in resume_2017_docx_section
    assert "| has_web_settings | true | true |" in resume_2017_docx_section
    assert "| has_footnotes | true | true |" in resume_2017_docx_section
    assert "| relationship_count | 9 | 9 |" in resume_2017_docx_section
    assert "| external_hyperlink_relationship_count | 0 | 0 |" in resume_2017_docx_section
    assert '| page_size | {"h": "15840", "w": "12240"} | {"h": "15840", "w": "12240"} |' in resume_2017_docx_section
    assert (
        '| page_margins | {"bottom": "1440", "footer": "720", "gutter": "0", "header": "720", '
        '"left": "1440", "right": "1440", "top": "1440"} | {"bottom": "1440", "footer": "720", '
        '"gutter": "0", "header": "720", "left": "1440", "right": "1440", "top": "1440"} |' in resume_2017_docx_section
    )
    assert '| numbering_num_ids | ["1"] | ["1"] |' in resume_2017_docx_section
    assert '| numbering_level_fonts | ["Symbol"] | ["Symbol"] |' in resume_2017_docx_section
    assert '| font_names | ["Calibri", "Symbol"] | ["Calibri", "Symbol"] |' in resume_2017_docx_section
    assert '| default_style_ids | ["Normal"] | ["Normal"] |' in resume_2017_docx_section
    assert '| style_run_fonts | {"Normal": "Calibri"} | {"Normal": "Calibri"} |' in resume_2017_docx_section
    assert (
        '| style_run_sizes | {"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} | '
        '{"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} |' in resume_2017_docx_section
    )
    assert (
        '| style_bold_ids | ["Title", "Heading1", "Heading2"] | ["Title", "Heading1", "Heading2"] |'
        in resume_2017_docx_section
    )
    assert (
        '| style_paragraph_indents | {"ListParagraph": {"left": "720"}} | '
        '{"ListParagraph": {"left": "720"}} |' in resume_2017_docx_section
    )
    assert (
        '| style_numbering | {"ListBullet": {"level": "0", "num_id": "1"}} | '
        '{"ListBullet": {"level": "0", "num_id": "1"}} |' in resume_2017_docx_section
    )
    assert (
        '| style_definition_ids | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", '
        '"ListBullet"] | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", "ListBullet"] |'
        in resume_2017_docx_section
    )
    assert (
        '| style_based_on | {"Heading1": "Normal", "Heading2": "Normal", "ListBullet": '
        '"ListParagraph", "ListParagraph": "Normal", "Title": "Normal"} | {"Heading1": "Normal", '
        '"Heading2": "Normal", "ListBullet": "ListParagraph", "ListParagraph": "Normal", '
        '"Title": "Normal"} |' in resume_2017_docx_section
    )
    assert (
        '| style_paragraph_spacing | {"Heading1": {"after": "80", "before": "160"}, "Heading2": '
        '{"after": "60", "before": "120"}, "Title": {"after": "120"}} | {"Heading1": '
        '{"after": "80", "before": "160"}, "Heading2": {"after": "60", "before": "120"}, '
        '"Title": {"after": "120"}} |' in resume_2017_docx_section
    )
    assert (
        "### resume_2023_a_posteriori.docx\n\n"
        "| Metric | Source | Generated |\n"
        "| --- | ---: | ---: |\n"
        "| part_names |" in audit_text
    )
    assert "| style_definition_count | 6 | 6 |" in resume_2023_docx_section
    assert "| part_count | 17 | 17 |" in resume_2023_docx_section
    assert "| has_theme | true | true |" in resume_2023_docx_section
    assert "| has_font_table | true | true |" in resume_2023_docx_section
    assert "| has_web_settings | true | true |" in resume_2023_docx_section
    assert "| relationship_count | 9 | 9 |" in resume_2023_docx_section
    assert "| external_hyperlink_relationship_count | 0 | 0 |" in resume_2023_docx_section
    assert '| page_size | {"h": "15840", "w": "12240"} | {"h": "15840", "w": "12240"} |' in resume_2023_docx_section
    assert (
        '| page_margins | {"bottom": "1440", "footer": "720", "gutter": "0", "header": "720", '
        '"left": "1800", "right": "1800", "top": "1440"} | {"bottom": "1440", "footer": "720", '
        '"gutter": "0", "header": "720", "left": "1800", "right": "1800", "top": "1440"} |' in resume_2023_docx_section
    )
    assert '| numbering_num_ids | ["1"] | ["1"] |' in resume_2023_docx_section
    assert '| numbering_level_fonts | ["Symbol"] | ["Symbol"] |' in resume_2023_docx_section
    assert '| font_names | ["Calibri", "Symbol"] | ["Calibri", "Symbol"] |' in resume_2023_docx_section
    assert '| default_style_ids | ["Normal"] | ["Normal"] |' in resume_2023_docx_section
    assert '| style_run_fonts | {"Normal": "Calibri"} | {"Normal": "Calibri"} |' in resume_2023_docx_section
    assert (
        '| style_run_sizes | {"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} | '
        '{"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} |' in resume_2023_docx_section
    )
    assert (
        '| style_bold_ids | ["Title", "Heading1", "Heading2"] | ["Title", "Heading1", "Heading2"] |'
        in resume_2023_docx_section
    )
    assert (
        '| style_paragraph_indents | {"ListParagraph": {"left": "720"}} | '
        '{"ListParagraph": {"left": "720"}} |' in resume_2023_docx_section
    )
    assert (
        '| style_numbering | {"ListBullet": {"level": "0", "num_id": "1"}} | '
        '{"ListBullet": {"level": "0", "num_id": "1"}} |' in resume_2023_docx_section
    )
    assert (
        '| style_definition_ids | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", '
        '"ListBullet"] | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", "ListBullet"] |'
        in resume_2023_docx_section
    )
    assert (
        '| style_based_on | {"Heading1": "Normal", "Heading2": "Normal", "ListBullet": '
        '"ListParagraph", "ListParagraph": "Normal", "Title": "Normal"} | {"Heading1": "Normal", '
        '"Heading2": "Normal", "ListBullet": "ListParagraph", "ListParagraph": "Normal", '
        '"Title": "Normal"} |' in resume_2023_docx_section
    )
    assert (
        '| style_paragraph_spacing | {"Heading1": {"after": "80", "before": "160"}, "Heading2": '
        '{"after": "60", "before": "120"}, "Title": {"after": "120"}} | {"Heading1": '
        '{"after": "80", "before": "160"}, "Heading2": {"after": "60", "before": "120"}, '
        '"Title": {"after": "120"}} |' in resume_2023_docx_section
    )
    assert (
        "### resume_2024_a_posteriori.docx\n\n"
        "| Metric | Source | Generated |\n"
        "| --- | ---: | ---: |\n"
        "| part_names |" in audit_text
    )
    assert "| style_definition_count | 6 | 6 |" in resume_2024_docx_section
    assert "| part_count | 17 | 17 |" in resume_2024_docx_section
    assert "| has_theme | true | true |" in resume_2024_docx_section
    assert "| has_font_table | true | true |" in resume_2024_docx_section
    assert "| has_web_settings | true | true |" in resume_2024_docx_section
    assert "| relationship_count | 9 | 9 |" in resume_2024_docx_section
    assert "| external_hyperlink_relationship_count | 0 | 0 |" in resume_2024_docx_section
    assert '| page_size | {"h": "15840", "w": "12240"} | {"h": "15840", "w": "12240"} |' in resume_2024_docx_section
    assert (
        '| page_margins | {"bottom": "1440", "footer": "720", "gutter": "0", "header": "720", '
        '"left": "1440", "right": "1440", "top": "1440"} | {"bottom": "1440", "footer": "720", '
        '"gutter": "0", "header": "720", "left": "1440", "right": "1440", "top": "1440"} |' in resume_2024_docx_section
    )
    assert '| numbering_num_ids | ["1"] | ["1"] |' in resume_2024_docx_section
    assert '| numbering_level_fonts | ["Symbol"] | ["Symbol"] |' in resume_2024_docx_section
    assert '| font_names | ["Calibri", "Symbol"] | ["Calibri", "Symbol"] |' in resume_2024_docx_section
    assert '| default_style_ids | ["Normal"] | ["Normal"] |' in resume_2024_docx_section
    assert '| style_run_fonts | {"Normal": "Calibri"} | {"Normal": "Calibri"} |' in resume_2024_docx_section
    assert (
        '| style_run_sizes | {"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} | '
        '{"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} |' in resume_2024_docx_section
    )
    assert (
        '| style_bold_ids | ["Title", "Heading1", "Heading2"] | ["Title", "Heading1", "Heading2"] |'
        in resume_2024_docx_section
    )
    assert (
        '| style_paragraph_indents | {"ListParagraph": {"left": "720"}} | '
        '{"ListParagraph": {"left": "720"}} |' in resume_2024_docx_section
    )
    assert (
        '| style_numbering | {"ListBullet": {"level": "0", "num_id": "1"}} | '
        '{"ListBullet": {"level": "0", "num_id": "1"}} |' in resume_2024_docx_section
    )
    assert (
        '| style_definition_ids | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", '
        '"ListBullet"] | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", "ListBullet"] |'
        in resume_2024_docx_section
    )
    assert (
        '| style_based_on | {"Heading1": "Normal", "Heading2": "Normal", "ListBullet": '
        '"ListParagraph", "ListParagraph": "Normal", "Title": "Normal"} | {"Heading1": "Normal", '
        '"Heading2": "Normal", "ListBullet": "ListParagraph", "ListParagraph": "Normal", '
        '"Title": "Normal"} |' in resume_2024_docx_section
    )
    assert (
        '| style_paragraph_spacing | {"Heading1": {"after": "80", "before": "160"}, "Heading2": '
        '{"after": "60", "before": "120"}, "Title": {"after": "120"}} | {"Heading1": '
        '{"after": "80", "before": "160"}, "Heading2": {"after": "60", "before": "120"}, '
        '"Title": {"after": "120"}} |' in resume_2024_docx_section
    )
    assert "| part_count | 17 | 17 |" in audit_text
    assert "| has_theme | true | true |" in audit_text
    assert "| has_font_table | true | true |" in audit_text
    assert "| relationship_count | 9 | 9 |" in audit_text
    assert "| external_hyperlink_relationship_count | 0 | 0 |" in audit_text
    assert '| page_size | {"h": "15840", "w": "12240"} | {"h": "15840", "w": "12240"} |' in audit_text
    assert (
        '| page_margins | {"bottom": "1440", "footer": "720", "gutter": "0", "header": "720", '
        '"left": "1440", "right": "1440", "top": "1440"} | {"bottom": "1440", "footer": "720", '
        '"gutter": "0", "header": "720", "left": "1440", "right": "1440", "top": "1440"} |' in audit_text
    )
    assert '| numbering_num_ids | ["1"] | ["1"] |' in audit_text
    assert '| numbering_level_fonts | ["Symbol"] | ["Symbol"] |' in audit_text
    assert '| font_names | ["Calibri", "Symbol"] | ["Calibri", "Symbol"] |' in audit_text
    assert '| default_style_ids | ["Normal"] | ["Normal"] |' in audit_text
    assert '| style_run_fonts | {"Normal": "Calibri"} | {"Normal": "Calibri"} |' in audit_text
    assert (
        '| style_run_sizes | {"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} | '
        '{"Heading1": "24", "Heading2": "22", "Normal": "22", "Title": "32"} |' in audit_text
    )
    assert '| style_bold_ids | ["Title", "Heading1", "Heading2"] | ["Title", "Heading1", "Heading2"] |' in audit_text
    assert (
        '| style_paragraph_indents | {"ListParagraph": {"left": "720"}} | '
        '{"ListParagraph": {"left": "720"}} |' in audit_text
    )
    assert (
        '| style_numbering | {"ListBullet": {"level": "0", "num_id": "1"}} | '
        '{"ListBullet": {"level": "0", "num_id": "1"}} |' in audit_text
    )
    assert (
        '| style_definition_ids | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", '
        '"ListBullet"] | ["Normal", "Title", "Heading1", "Heading2", "ListParagraph", "ListBullet"] |' in audit_text
    )
    assert "master_profile_a_posteriori.xlsx" in audit_text
    assert "### master_profile_a_posteriori.xlsx" in audit_text
    assert "| part_count | 16 | 16 |" in audit_text
    assert "| worksheet_part_count | 9 | 9 |" in audit_text
    assert "| has_core_properties | true | true |" in audit_text
    assert "| has_extended_properties | true | true |" in audit_text
    assert "| has_styles | true | true |" in audit_text
    assert (
        '| root_relationship_type_counts | {"core-properties": 1, "extended-properties": 1, '
        '"officeDocument": 1} | {"core-properties": 1, "extended-properties": 1, "officeDocument": 1} |' in audit_text
    )
    assert (
        '| workbook_relationship_type_counts | {"styles": 1, "worksheet": 9} | '
        '{"styles": 1, "worksheet": 9} |' in audit_text
    )
    assert (
        '| sheet_names | ["Highlights", "Jobs", "Standards Development", "School", "Certifications", '
        '"Patents", "Publications", "Lectures", "Residences"] | ["Highlights", "Jobs", '
        '"Standards Development", "School", "Certifications", "Patents", "Publications", "Lectures", '
        '"Residences"] |' in audit_text
    )
    assert "| Source Sheet | Generated Sheet | Source Path | Generated Path |" in audit_text
    assert "Source Rows | Generated Rows | Source Columns | Generated Columns" in audit_text
    assert "Source Styled Cells | Generated Styled Cells" in audit_text
    assert "Source Cell Types | Generated Cell Types" in audit_text
    assert "Source Freeze Pane | Generated Freeze Pane | Source Filter | Generated Filter" in audit_text
    assert "Source Page Margins | Generated Page Margins | Source Column Widths | Generated Column Widths" in audit_text
    assert (
        "| Highlights | Highlights | xl/worksheets/sheet1.xml | xl/worksheets/sheet1.xml | A1:C1 | "
        "A1:C1 | 1 | 1 | 3 | 3 | 3 | 3 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:C1 | A1:C1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "12.00"] | ["10.00", "10.00", "12.00"] | '
        '{"inlineStr": 3} | {"inlineStr": 3} | ["ID", "Label", "Highlights"] | '
        '["ID", "Label", "Highlights"] |' in audit_text
    )
    assert (
        "| Jobs | Jobs | xl/worksheets/sheet2.xml | xl/worksheets/sheet2.xml | A1:M1 | "
        "A1:M1 | 1 | 1 | 13 | 13 | 13 | 13 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:M1 | A1:M1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "11.00", "11.00", "16.00", "17.00", "10.00", "10.00", "20.00", '
        '"12.00", "12.00", "10.00", "12.00"] | ["10.00", "10.00", "11.00", "11.00", '
        '"16.00", "17.00", "10.00", "10.00", "20.00", "12.00", "12.00", "10.00", "12.00"] | '
        '{"inlineStr": 13} | {"inlineStr": 13} | ["ID", "Label", "Job Title", "Job Focus", '
        '"Job Highlights", "Vehicle Program", "Employer", "Address", "Reason For Leaving", "Salary Usd", '
        '"Start Date", "End Date", "Time Spent"] | ["ID", "Label", "Job Title", "Job Focus", '
        '"Job Highlights", "Vehicle Program", "Employer", "Address", "Reason For Leaving", "Salary Usd", '
        '"Start Date", "End Date", "Time Spent"] |' in audit_text
    )
    assert (
        "| Standards Development | Standards Development | xl/worksheets/sheet3.xml | "
        "xl/worksheets/sheet3.xml | A1:H1 | A1:H1 | 1 | 1 | 8 | 8 | 8 | 8 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:H1 | A1:H1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "14.00", "19.00", "10.00", "12.00", "10.00", "12.00"] | '
        '["10.00", "10.00", "14.00", "19.00", "10.00", "12.00", "10.00", "12.00"] | '
        '{"inlineStr": 8} | {"inlineStr": 8} | ["ID", "Label", "Organization", "Title Of Standard", '
        '"Role", "Start Date", "End Date", "Time Spent"] | ["ID", "Label", "Organization", '
        '"Title Of Standard", "Role", "Start Date", "End Date", "Time Spent"] |' in audit_text
    )
    assert (
        "| School | School | xl/worksheets/sheet4.xml | xl/worksheets/sheet4.xml | "
        "A1:H1 | A1:H1 | 1 | 1 | 8 | 8 | 8 | 8 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:H1 | A1:H1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "10.00", "10.00", "10.00", "12.00", "10.00", "12.00"] | '
        '["10.00", "10.00", "10.00", "10.00", "10.00", "12.00", "10.00", "12.00"] | '
        '{"inlineStr": 8} | {"inlineStr": 8} | ["ID", "Label", "School", "Address", '
        '"Degree", "Start Date", "End Date", "Time Spent"] | ["ID", "Label", "School", '
        '"Address", "Degree", "Start Date", "End Date", "Time Spent"] |' in audit_text
    )
    assert (
        "| Certifications | Certifications | xl/worksheets/sheet5.xml | "
        "xl/worksheets/sheet5.xml | A1:E1 | A1:E1 | 1 | 1 | 5 | 5 | 5 | 5 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:E1 | A1:E1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "21.00", "18.00", "20.00"] | '
        '["10.00", "10.00", "21.00", "18.00", "20.00"] | '
        '{"inlineStr": 5} | {"inlineStr": 5} | ["ID", "Label", "Certificate Subject", '
        '"Certificate Name", "Certificate Number"] | ["ID", "Label", "Certificate Subject", '
        '"Certificate Name", "Certificate Number"] |' in audit_text
    )
    assert (
        "| Patents | Patents | xl/worksheets/sheet6.xml | xl/worksheets/sheet6.xml | "
        "A1:E1 | A1:E1 | 1 | 1 | 5 | 5 | 5 | 5 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:E1 | A1:E1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "13.00", "15.00", "12.00"] | '
        '["10.00", "10.00", "13.00", "15.00", "12.00"] | '
        '{"inlineStr": 5} | {"inlineStr": 5} | ["ID", "Label", "Patent Name", "Patent Number", '
        '"Patent Url"] | ["ID", "Label", "Patent Name", "Patent Number", "Patent Url"] |' in audit_text
    )
    assert (
        "| Publications | Publications | xl/worksheets/sheet7.xml | "
        "xl/worksheets/sheet7.xml | A1:G1 | A1:G1 | 1 | 1 | 7 | 7 | 7 | 7 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:G1 | A1:G1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "10.00", "18.00", "11.00", "23.00", "17.00"] | '
        '["10.00", "10.00", "10.00", "18.00", "11.00", "23.00", "17.00"] | '
        '{"inlineStr": 7} | {"inlineStr": 7} | ["ID", "Label", "Year", "Publication Name", '
        '"Publisher", "Publication Reference", "Publication Url"] | ["ID", "Label", "Year", '
        '"Publication Name", "Publisher", "Publication Reference", "Publication Url"] |' in audit_text
    )
    assert (
        "| Lectures | Lectures | xl/worksheets/sheet8.xml | xl/worksheets/sheet8.xml | "
        "A1:E1 | A1:E1 | 1 | 1 | 5 | 5 | 5 | 5 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:E1 | A1:E1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "10.00", "29.00", "16.00"] | '
        '["10.00", "10.00", "10.00", "29.00", "16.00"] | '
        '{"inlineStr": 5} | {"inlineStr": 5} | ["ID", "Label", "Year", '
        '"Publication Or Lecture Name", "Publisher Name"] | ["ID", "Label", "Year", '
        '"Publication Or Lecture Name", "Publisher Name"] |' in audit_text
    )
    assert (
        "| Residences | Residences | xl/worksheets/sheet9.xml | xl/worksheets/sheet9.xml | "
        "A1:F1 | A1:F1 | 1 | 1 | 6 | 6 | 6 | 6 | "
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        '{"activePane": "bottomLeft", "state": "frozen", "topLeftCell": "A2", "ySplit": "1"} | '
        "A1:F1 | A1:F1 | "
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '{"bottom": "0.75", "footer": "0.3", "header": "0.3", "left": "0.7", '
        '"right": "0.7", "top": "0.75"} | '
        '["10.00", "10.00", "11.00", "12.00", "10.00", "12.00"] | '
        '["10.00", "10.00", "11.00", "12.00", "10.00", "12.00"] | '
        '{"inlineStr": 6} | {"inlineStr": 6} | ["ID", "Label", "Residence", '
        '"Start Date", "End Date", "Time Spent"] | ["ID", "Label", "Residence", '
        '"Start Date", "End Date", "Time Spent"] |' in audit_text
    )
    assert "## Known Acceptable Differences" in audit_text
    assert "| Generated Path | Scope | Reason |" in audit_text
    assert (
        "| generated/office/a_posteriori/master_profile_a_posteriori.xlsx | workbook_app_native_columns | "
        "Reviewed v0.0.2 workbook drift: the generated XLSX preserves the canonical nine-sheet app-native "
        "workbook contract with ID/Label columns, while the legacy source workbook omits those helper columns. |"
        in audit_text
    )


def test_audit_policy_can_accept_explicit_drift() -> None:
    comparison = {
        "generated": {"path": "generated/office/a_posteriori/resume_fake.docx"},
        "source": {"path": "source_office/a_priori/resume_fake.docx"},
        "byte_identical": False,
        "normalized_text_match": False,
    }
    audit_policy = {
        "schema_version": app.SCHEMA_VERSION,
        "policy_version": "test",
        "name": "test_policy",
        "description": "Test policy",
        "accepted_differences": [
            {
                "generated_path": r"generated\office\a_posteriori\resume_fake.docx",
                "scope": "package",
                "reason": "Known package-only drift in fake fixture.",
            }
        ],
    }

    classified = app.apply_audit_policy(comparison, audit_policy)

    assert classified["status"] == "accepted_drift"
    assert classified["status_reason"] == "Known package-only drift in fake fixture."
    assert app.audit_policy_summary(audit_policy, Path(r"audit_policies\test_policy.json"))["path"] == (
        "audit_policies/test_policy.json"
    )
    unaccepted_comparison = dict(comparison)
    unaccepted_comparison["generated"] = {"path": "generated/office/a_posteriori/unlisted.docx"}

    unaccepted = app.apply_audit_policy(unaccepted_comparison, audit_policy)

    assert unaccepted["status"] == "review_required"
    assert unaccepted["status_reason"] == "No policy entry accepts this generated/source difference."


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
