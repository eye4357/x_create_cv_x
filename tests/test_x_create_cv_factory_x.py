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
        content_types_xml = workbook.read("[Content_Types].xml").decode("utf-8")
        root_relationships_xml = workbook.read("_rels/.rels").decode("utf-8")
        core_properties_xml = workbook.read("docProps/core.xml").decode("utf-8")
        app_properties_xml = workbook.read("docProps/app.xml").decode("utf-8")
        workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
        workbook_relationships_xml = workbook.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        styles_xml = workbook.read("xl/styles.xml").decode("utf-8")

    assert workbook_sheet_names(workbook_path) == ["Highlights"]
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
    assert "<dc:title>Master Profile A Posteriori</dc:title>" in core_properties_xml
    assert "<dc:creator>x_create_cv_x</dc:creator>" in core_properties_xml
    assert "<cp:lastModifiedBy>x_create_cv_x</cp:lastModifiedBy>" in core_properties_xml
    assert '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:created>' in core_properties_xml
    assert '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:modified>' in core_properties_xml
    assert "<Application>x_create_cv_x</Application>" in app_properties_xml
    assert '<sheet name="Highlights" sheetId="1" r:id="rId1"/>' in workbook_xml
    assert (
        'Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"' in workbook_relationships_xml
    )
    assert (
        'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"' in workbook_relationships_xml
    )
    assert f'<styleSheet xmlns="{app.SHEET_NS}">' in styles_xml
    assert '<fonts count="3">' in styles_xml
    assert '<font><b/><sz val="11"/><color rgb="FF102030"/><name val="Calibri"/><family val="2"/></font>' in styles_xml
    assert '<fills count="4">' in styles_xml
    assert '<fgColor rgb="FFABCDEF"/>' in styles_xml
    assert '<borders count="2">' in styles_xml
    assert '<left style="thin"><color rgb="FF405060"/></left>' in styles_xml
    assert '<right style="thin"><color rgb="FF405060"/></right>' in styles_xml
    assert '<top style="thin"><color rgb="FF405060"/></top>' in styles_xml
    assert '<bottom style="thin"><color rgb="FF405060"/></bottom>' in styles_xml
    assert '<cellXfs count="3">' in styles_xml
    assert (
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>' in styles_xml
    )
    assert worksheet_row_values(workbook_path, 1) == ["ID", "Label", "Highlights", "Score", "Current"]
    assert worksheet_row_values(workbook_path, 2) == ["highlight_001", "One", "Structured value", "", ""]
    sheet_summary = first_worksheet_summary(workbook_path)
    assert sheet_summary["path"] == "xl/worksheets/sheet1.xml"
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
    assert workbook_summary["sheet_names"] == ["Highlights"]
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
    assert style_summary["cell_xfs"][0] == {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"}
    assert style_summary["cell_xfs"][1]["fontId"] == "1"
    assert style_summary["cell_xfs"][1]["fillId"] == "2"
    assert style_summary["cell_xfs"][1]["borderId"] == "1"
    assert style_summary["cell_xfs"][1]["applyFont"] == "1"
    assert style_summary["cell_xfs"][1]["applyFill"] == "1"
    assert style_summary["cell_xfs"][1]["applyBorder"] == "1"
    assert style_summary["cell_xfs"][2]["borderId"] == "1"
    assert style_summary["cell_xfs"][2]["applyBorder"] == "1"
    assert style_summary["cell_xfs"][2]["applyAlignment"] == "1"
    worksheet_xml = first_worksheet_xml(workbook_path)
    assert f'<worksheet xmlns="{app.SHEET_NS}" xmlns:r="{app.REL_NS}">' in worksheet_xml
    assert '<dimension ref="A1:E2"/>' in worksheet_xml
    assert '<sheetViews><sheetView tabSelected="0" workbookViewId="0">' in worksheet_xml
    assert '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' in worksheet_xml
    assert '<selection pane="bottomLeft"/>' in worksheet_xml
    assert '<sheetFormatPr defaultRowHeight="15"/>' in worksheet_xml
    assert '<sheetData><row r="1">' in worksheet_xml
    assert '</row><row r="2">' in worksheet_xml
    assert '<autoFilter ref="A1:E2"/>' in worksheet_xml
    assert '<pageMargins left="0.2" right="0.4" top="0.6" bottom="0.8" header="0.1" footer="0.3"/>' in worksheet_xml
    assert '<col min="1" max="1" width="8.50" customWidth="1"/>' in worksheet_xml
    assert '<col min="4" max="4" width="10.00" customWidth="1"/>' in worksheet_xml
    assert '<c r="A1" s="1" t="inlineStr"><is><t xml:space="preserve">ID</t></is></c>' in worksheet_xml
    assert '<c r="C2" s="2" t="inlineStr"><is><t xml:space="preserve">Structured value</t></is></c>' in worksheet_xml
    assert '<c r="D2" s="2"><v>98.5</v></c>' in worksheet_xml
    assert '<c r="E2" s="2" t="b"><v>1</v></c>' in worksheet_xml

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
    document_xml = docx_document_xml(document_path)
    with zipfile.ZipFile(document_path) as document:
        core_properties_xml = document.read("docProps/core.xml").decode("utf-8")
        app_properties_xml = document.read("docProps/app.xml").decode("utf-8")
        docx_styles_xml = document.read("word/styles.xml").decode("utf-8")
        docx_numbering_xml = document.read("word/numbering.xml").decode("utf-8")
    structure_summary = app.docx_structure_summary(document_path)

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
    assert "<dc:title>API Resume</dc:title>" in core_properties_xml
    assert "<dc:creator>x_create_cv_x</dc:creator>" in core_properties_xml
    assert "<cp:lastModifiedBy>x_create_cv_x</cp:lastModifiedBy>" in core_properties_xml
    assert '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:created>' in core_properties_xml
    assert '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-01T00:00:00Z</dcterms:modified>' in core_properties_xml
    assert "<Application>x_create_cv_x</Application>" in app_properties_xml
    assert f'<w:styles xmlns:w="{app.WORD_NS}">' in docx_styles_xml
    assert (
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>' in docx_styles_xml
    )
    assert (
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>' in docx_styles_xml
    )
    assert (
        '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>' in docx_styles_xml
    )
    assert (
        '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/>'
        '<w:basedOn w:val="ListParagraph"/><w:pPr><w:numPr><w:ilvl w:val="0"/>'
        '<w:numId w:val="1"/></w:numPr></w:pPr></w:style>' in docx_styles_xml
    )
    assert '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Styled</w:t></w:r>' in document_xml
    assert (
        '<w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr>' '<w:t xml:space="preserve"> run</w:t></w:r>' in document_xml
    )
    assert f'<w:numbering xmlns:w="{app.WORD_NS}">' in docx_numbering_xml
    assert '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' in docx_numbering_xml
    assert '<w:abstractNum w:abstractNumId="7"><w:multiLevelType w:val="hybridMultilevel"/>' in docx_numbering_xml
    assert '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' in docx_numbering_xml
    assert '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' in docx_numbering_xml
    assert '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>' in docx_numbering_xml
    assert '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' in docx_numbering_xml
    assert '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>' in docx_numbering_xml
    assert '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>' in docx_numbering_xml
    assert structure_summary["numbered_paragraph_count"] == 1
    assert structure_summary["numbering_abstract_count"] == 2
    assert structure_summary["numbering_num_count"] == 2
    assert structure_summary["numbering_level_count"] == 4
    assert structure_summary["numbering_num_ids"] == ["1", "7"]
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
    assert structure_summary["style_based_on"]["ListBullet"] == "ListParagraph"
    assert structure_summary["style_run_fonts"] == {"Normal": "Calibri"}
    assert structure_summary["style_run_sizes"]["Heading1"] == "24"
    assert structure_summary["style_bold_ids"] == ["Title", "Heading1", "Heading2"]
    assert structure_summary["style_paragraph_spacing"]["Heading1"] == {"before": "160", "after": "80"}
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
    assert theme_xml.startswith(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'name="Generated Office Theme"><a:themeElements><a:clrScheme name="Generated">'
    )
    expected_theme_fragments = [
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>',
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>',
        '<a:dk2><a:srgbClr val="1F4E79"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>',
        '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>',
        '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>',
        '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>',
        '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>',
        '<a:fontScheme name="Generated"><a:majorFont><a:latin typeface="Calibri"/>',
        '<a:minorFont><a:latin typeface="Calibri"/>',
        '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">',
        '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">',
        '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">',
    ]
    for expected_theme_fragment in expected_theme_fragments:
        assert expected_theme_fragment in theme_xml
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
    assert f'<w:settings xmlns:w="{app.WORD_NS}"><w:defaultTabStop w:val="720"/></w:settings>' in settings_xml
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
    assert structure_summary["styles"] == ["Title"]
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
    assert structure_summary["relationship_type_counts"]["hyperlink"] == 1
    assert structure_summary["relationship_target_mode_counts"] == {"External": 1}
    assert structure_summary["external_hyperlink_relationship_count"] == 1
    assert structure_summary["styled_paragraph_count"] == 1
    assert structure_summary["aligned_paragraph_count"] == 1
    assert structure_summary["spaced_paragraph_count"] == 1
    assert structure_summary["indented_paragraph_count"] == 1
    assert structure_summary["tab_stopped_paragraph_count"] == 1
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
    assert "docProps/core.xml" not in part_names
    assert "docProps/app.xml" not in part_names
    assert "word/webSettings.xml" not in part_names
    assert "word/footnotes.xml" not in part_names
    assert "word/endnotes.xml" not in part_names
    assert "customXml/item1.xml" not in part_names
    with zipfile.ZipFile(document_path) as document:
        root_relationships = document.read("_rels/.rels").decode("utf-8")
        content_types = document.read("[Content_Types].xml").decode("utf-8")
    assert "core-properties" not in root_relationships
    assert "extended-properties" not in root_relationships
    assert "/docProps/core.xml" not in content_types
    assert "/docProps/app.xml" not in content_types


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
    audit_text = audit_path.read_text(encoding="utf-8")
    assert "# A Posteriori Office Audit" in audit_text
    assert "## DOCX Structure" in audit_text
    assert "## XLSX Structure" in audit_text
    assert "master_profile_a_posteriori.xlsx" in audit_text
    assert "| Source Sheet | Generated Sheet | Source Path | Generated Path |" in audit_text
    assert "Source Rows | Generated Rows | Source Columns | Generated Columns" in audit_text
    assert "Source Styled Cells | Generated Styled Cells" in audit_text
    assert "Source Cell Types | Generated Cell Types" in audit_text
    assert "Source Freeze Pane | Generated Freeze Pane | Source Filter | Generated Filter" in audit_text
    assert "Source Page Margins | Generated Page Margins | Source Column Widths | Generated Column Widths" in audit_text


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
