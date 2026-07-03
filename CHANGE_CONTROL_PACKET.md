# Change Control Packet

Tool: `x_create_cv_x`
Version: `0.0.3`
Packet ID: `x_create_cv_x-0.0.3-2026-07-02`
Date: `2026-07-02`
Status: Updated

## Scope

Start renderer conformance hardening on top of the `0.0.2` contract-and-audit lock while preserving the side-by-side private golden evidence workflow.

## Controlled Files

- `x_create_cv_factory_x.py`
- `README.md`
- `WORLD_CLASS_CV_GLIDEPATH.md`
- `RELEASE_NOTES_0.0.3.md`
- `schemas/master_profile.schema.json`
- `schemas/resume.schema.json`
- `schemas/workbook_layout.schema.json`
- `schemas/document_layout.schema.json`
- `schemas/audit_policy.schema.json`
- `audit_policies/default_office_audit_policy.json`
- `CV_OFFICE_REGENERATION_PLAN.md`
- `ONLINE_SERVICE_STRATEGY.md`
- `CHANGELOG.md`
- `CHANGE_CONTROL_PACKET.md`
- `SECURITY.md`
- `pyproject.toml`
- `requirements-dev.txt`
- `.github/workflows/ci.yml`
- `.vscode/settings.json`
- `.gitignore`
- `tests/test_x_create_cv_factory_x.py`
- `tests/fixtures/fake_profile_seed.json`

## Change Summary

- Bumped the development version from `0.0.1` to `0.0.2`.
- Bumped the development version from `0.0.2` to `0.0.3`.
- Added schema-backed per-sheet XLSX freeze-header, autofilter, and column-width controls.
- Added schema-backed XLSX column `value_type` controls for string, number, and boolean cell rendering.
- Added XLSX style definition audit metrics for fonts, fills, borders, and cell formats.
- Added XLSX package audit metrics for parts, content type overrides, root relationships, and workbook relationships.
- Added XLSX worksheet page-margin audit metrics from the schema-backed workbook style contract.
- Added XLSX sheet row, column, styled-cell, and column-width audit report metrics.
- Added XLSX package part-name and workbook sheet-name audit report metrics.
- Added XLSX worksheet-path audit report metrics.
- Added XLSX source worksheet-path audit report metrics.
- Added XLSX source row-count and column-count audit report metrics.
- Added XLSX source styled-cell-count audit report metrics.
- Added XLSX source cell-type-count audit report metrics.
- Added XLSX source freeze-pane and autofilter audit report metrics.
- Added XLSX source page-margin and column-width audit report metrics.
- Added XLSX source worksheet-name audit report metrics.
- Expanded XLSX structure summaries and Markdown audit reports to expose generated freeze-pane and autofilter state.
- Expanded XLSX structure summaries and Markdown audit reports to expose cell-type counts.
- Expanded XLSX structure summaries and Markdown audit reports to expose style color and cell-format details.
- Expanded XLSX structure summaries and Markdown audit reports to expose package and relationship details.
- Expanded XLSX sheet summaries and Markdown audit reports to expose generated worksheet page margins.
- Expanded XLSX Markdown audit reports to expose generated row counts, column counts, styled-cell counts, and column widths.
- Expanded XLSX Markdown audit reports to expose generated package part names and workbook sheet names.
- Expanded XLSX Markdown audit reports to expose generated worksheet paths.
- Expanded XLSX Markdown audit reports to expose source worksheet paths alongside generated worksheet paths.
- Expanded XLSX Markdown audit reports to expose source row and column counts alongside generated counts.
- Expanded XLSX Markdown audit reports to expose source styled-cell counts alongside generated counts.
- Expanded XLSX Markdown audit reports to expose source cell-type counts alongside generated counts.
- Expanded XLSX Markdown audit reports to expose source freeze-pane and autofilter state alongside generated state.
- Expanded XLSX Markdown audit reports to expose source page margins and column widths alongside generated values.
- Expanded XLSX Markdown audit reports to expose source worksheet names alongside generated worksheet names.
- Extended public fake Office tests to assert workbook conformance controls without private data.
- Extended public fake Office tests to assert XLSX package relationship metrics without private data.
- Extended public fake Office tests to assert XLSX page-margin metrics without private data.
- Extended public fake Office tests to assert XLSX row, column, styled-cell, and column-width metrics without private data.
- Extended public fake Office tests to assert XLSX package part-name and sheet-name metrics without private data.
- Extended public fake Office tests to assert XLSX worksheet-path metrics without private data.
- Extended public fake Office tests to assert XLSX source worksheet-path metrics without private data.
- Extended public fake Office tests to assert XLSX source row-count and column-count metrics without private data.
- Extended public fake Office tests to assert XLSX source styled-cell-count metrics without private data.
- Extended public fake Office tests to assert XLSX source cell-type-count metrics without private data.
- Extended public fake Office tests to assert XLSX source freeze-pane and autofilter metrics without private data.
- Extended public fake Office tests to assert XLSX source page-margin and column-width metrics without private data.
- Extended public fake Office tests to assert XLSX source worksheet-name metrics without private data.
- Added explicit DOCX run hyperlink/style fields to the public document layout schema.
- Expanded DOCX Markdown audit reports to expose package flags, body child counts, relationship type maps, font table names, and used styles.
- Extended public fake Office tests to assert DOCX package, body-shape, font-table, and used-style metrics without private data.
- Extended public fake Office tests to assert XLSX style font-name, cell format count, and cell style count metrics without private data.
- Extended public fake Office tests to assert XLSX default, header, and body cell format attributes without private data.
- Extended public fake Office tests to assert XLSX header and inline-string body style application without private data.
- Extended public fake Office tests to assert XLSX worksheet XML controls for widths, freeze panes, filters, and margins without private data.
- Extended public fake Office tests to assert XLSX workbook sheet and relationship binding without private data.
- Extended public fake Office tests to assert XLSX content-type and root relationship package binding without private data.
- Extended public fake Office tests to assert XLSX core and app document property metadata without private data.
- Extended public fake Office tests to assert XLSX style XML color, border, and alignment binding without private data.
- Extended public fake Office tests to assert XLSX worksheet-level XML structure and row ordering without private data.
- Extended public fake Office tests to assert DOCX styles XML definitions and list-style binding without private data.
- Extended public fake Office tests to assert DOCX numbering XML abstract and concrete numbering bindings without private data.
- Extended public fake Office tests to assert DOCX content-type and root relationship package binding without private data.
- Extended public fake Office tests to assert DOCX optional content-type package binding without private data.
- Extended public fake Office tests to assert lean DOCX optional-package part manifest without private data.
- Extended public fake Office tests to assert lean DOCX optional-package content-type XML without private data.
- Extended public fake Office tests to assert lean DOCX optional-package document relationship XML without private data.
- Extended public fake Office tests to assert lean DOCX optional-package root relationship XML without private data.
- Extended public fake Office tests to assert lean DOCX optional-package structure summary package flags and relationship counters without private data.
- Extended public fake Office tests to assert DOCX baseline and optional document relationship binding without private data.
- Extended public fake Office tests to assert exact DOCX document relationship XML with optional package parts and hyperlink targets without private data.
- Extended public fake Office tests to assert DOCX settings part and default tab stop XML without private data.
- Extended public fake Office tests to assert exact DOCX settings part XML without private data.
- Extended public fake Office tests to assert exact DOCX package part manifest without private data.
- Extended public fake Office tests to assert DOCX structure summary package manifest without private data.
- Extended public fake Office tests to assert DOCX structure summary relationship counts without private data.
- Extended public fake Office tests to assert DOCX structure summary content counters without private data.
- Extended public fake Office tests to assert DOCX structure summary page setup values without private data.
- Extended public fake Office tests to assert DOCX header and footer part XML without private data.
- Extended public fake Office tests to assert DOCX custom XML item, properties, and relationship part XML without private data.
- Extended public fake Office tests to assert DOCX footnotes and endnotes part XML without private data.
- Extended public fake Office tests to assert DOCX font table and web settings part XML without private data.
- Extended public fake Office tests to assert DOCX theme color, font, and line-style XML without private data.
- Extended public fake Office tests to assert exact DOCX theme XML without private data.
- Extended public fake Office tests to assert DOCX package root relationship XML without private data.
- Extended public fake Office tests to assert DOCX content-type default declarations without private data.
- Extended public fake Office tests to assert enabled DOCX content-type override XML without private data.
- Extended public fake Office tests to assert DOCX document root namespace envelope XML without private data.
- Extended public fake Office tests to assert DOCX default section page size and margin XML without private data.
- Extended public fake Office tests to assert DOCX table width and border property XML without private data.
- Extended public fake Office tests to assert XLSX content-type default declarations without private data.
- Extended public fake Office tests to assert XLSX package root relationship XML without private data.
- Extended public fake Office tests to assert XLSX workbook relationship XML without private data.
- Extended public fake Office tests to assert XLSX workbook root and sheet XML without private data.
- Extended public fake Office tests to assert XLSX content-type override XML without private data.
- Extended public fake Office tests to assert DOCX core and app document property metadata without private data.
- Extended public fake Office tests to assert DOCX table grid and cell width XML without private data.
- Extended public fake Office tests to assert DOCX section page margin XML without private data.
- Extended public fake Office tests to assert DOCX run property and text binding without private data.
- Extended public fake Office tests to assert DOCX numbering level and numbering ID binding without private data.
- Extended public fake Office tests to assert DOCX hyperlink relationship and formatted run binding without private data.
- Extended public fake Office tests to assert DOCX package relationship ID, type, and target binding without private data.
- Extended public fake Office tests to assert DOCX header and footer part content without private data.
- Extended public fake Office tests to assert DOCX font table part content without private data.
- Extended public fake Office tests to assert optional DOCX package part XML content without private data.
- Extended public fake Office tests to assert DOCX paragraph property binding without private data.
- Extended public fake Office tests to assert DOCX tab run and following text adjacency without private data.
- Extended public fake Office tests to assert DOCX table cell content binding without private data.
- Extended public fake Office tests to assert explicit DOCX empty table paragraph binding without private data.
- Expanded DOCX structure summaries and Markdown audit reports to expose relationship counts and external hyperlink relationship counts.
- Extended public fake Office tests to assert external hyperlink relationship conformance without private data.
- Expanded DOCX structure summaries and Markdown audit reports to expose paragraph style, alignment, spacing, indentation, and tab-stop counts.
- Extended public fake Office tests to assert DOCX paragraph property audit metrics without private data.
- Expanded DOCX structure summaries and Markdown audit reports to expose table row, cell, paragraph, grid-width, and cell-width geometry.
- Extended public fake Office tests to assert DOCX table width audit metrics without private data.
- Expanded DOCX structure summaries and Markdown audit reports to expose numbering definitions, IDs, levels, bullet text patterns, and fonts.
- Extended public fake Office tests to assert DOCX numbering definition audit metrics without private data.
- Expanded DOCX structure summaries and Markdown audit reports to expose section page size, margins, and header/footer references.
- Extended public fake Office tests to assert DOCX page setup audit metrics without private data.
- Expanded DOCX structure summaries and Markdown audit reports to expose style definitions, defaults, based-on links, run properties, paragraph properties, and style numbering.
- Extended public fake Office tests to assert DOCX style definition audit metrics without private data.
- Added `CV_OFFICE_REGENERATION_PLAN.md` as the controlled roadmap for replacing or downgrading `private.zip` with Office evidence.
- Extracted the three local source archives into original `_a_priori` DOCX/XLSX files under the ignored private evidence boundary, then removed the ZIP archives locally.
- Added a private SHA-256 manifest for the three `_a_priori` Office files.
- Added `check-evidence` to fast-fail when private `_a_priori` Office evidence is missing, size-mismatched, or hash-mismatched.
- Added `exercise-golden` as the simple local smoke test for the side-by-side private repository `x_create_cv_test_data_x`.
- Documented the `_a_priori` suffix for original source evidence and the `_a_posteriori` suffix for future whole-cloth generated Office files.
- Moved the real `_a_priori` Office evidence and SHA-256 manifest ownership into `x_create_cv_test_data_x`.
- Moved private seed scripts, generated JSON, legacy JSON, and the legacy private zip ownership into `x_create_cv_test_data_x`.
- Added full chain-of-evidence manifest validation to `exercise-golden`.
- Added private script execution and generated JSON comparison to `exercise-golden`.
- Documented that `_a_posteriori` Office files must be generated by scripts from JSON, not copied from `_a_priori` evidence.
- Added deterministic XLSX/DOCX generation from golden JSON for the private `_a_posteriori` Office evidence.
- Added `generate-golden-office` to write private Office evidence and a comparison report under `x_create_cv_test_data_x/evidence`.
- Extended `exercise-golden` to regenerate Office files into a temporary directory and byte-compare them against stored `_a_posteriori` Office evidence.
- Improved XLSX generation to preserve the original nine-sheet workbook shape while adding structured app-native data.
- Tuned DOCX generation toward `_a_priori` page setup and paragraph style parity.
- Added explicit generated JSON `office_layout` contracts for workbook and document layout.
- Changed the Office renderer to consume workbook sheet/column/style layout and DOCX `block_style`, `numbering`, and `runs` from generated JSON.
- Removed raw `a`/`b`/`c` columns from generated XLSX files in favor of named app-native columns.
- Added DOCX flow layout support for table blocks plus optional header/footer, theme, font table, and web settings package parts.
- Hardened generated DOCX theme packages for stricter importers such as Google Drive.
- Added DOCX contract support for hyperlinks, tabs, run font/size/color/style, footnotes/endnotes/custom XML package parts, and explicit font tables.
- Added DOCX contract support for paragraph alignment, spacing, indentation, tab stops, page size, and optional core/extended package properties.
- Added the fourth golden resume path for the current 2024 private source resume.
- Expanded DOCX structure reporting to include relationship types, body child counts, hyperlinks, tabs, package scaffolding, and font names.
- Optimized generated workbook sheets to skip redundant collection placeholder rows while preserving named app-native columns.
- Added non-sensitive structural Office metrics to the private comparison report.
- Added `WORLD_CLASS_CV_GLIDEPATH.md` as the public versioned workplan from the current `0.0.2` checkpoint through a stable `1.0.0` CV compiler baseline.
- Updated `CV_OFFICE_REGENERATION_PLAN.md` to include the current 2024 resume as the fourth golden source.
- Added public JSON Schema contracts for master profile, resume, workbook layout, and document layout JSON.
- Added dependency-free schema validation for generated master/resume JSON and private golden JSON checks.
- Added `validate-schema` as the public CLI surface for checking schema-backed JSON contracts.
- Added `audit` as the first-class CLI surface for private-safe Office parity reports.
- Expanded comparison summaries with DOCX package part names, run formatting counts, and XLSX sheet dimensions, headers, and style counts.
- Added versioned Office audit policy files and schema-backed accepted-drift classification.
- Added a reviewed accepted-drift policy entry for the generated app-native master workbook columns.
- Updated README, changelog, and security policy notes for the Office-regeneration evidence workflow.
- Preserved public CI as fake-fixture only; private Office evidence is exercised locally only when the private repo is present.

## Validation

- Run `python .\x_create_cv_factory_x.py --version`.
- Run `python .\x_create_cv_factory_x.py --help`.
- Run `python .\x_create_cv_factory_x.py validate-schema <json-files>` for generated master/resume JSON contract checks.
- Run `python .\x_create_cv_factory_x.py validate-schema .\audit_policies\default_office_audit_policy.json` for audit policy contract checks.
- Run `python .\x_create_cv_factory_x.py audit` when `..\x_create_cv_test_data_x\evidence` is available to produce JSON and Markdown Office parity reports.
- Run `python -m py_compile .\x_create_cv_factory_x.py`.
- Run `python -m pytest`.
- Run `ruff check .`.
- Run `black --check .`.
- Run `mypy`.
- Run `python .\x_create_cv_factory_x.py generate-golden-office` when `..\x_create_cv_test_data_x\evidence` is available to regenerate the private Office evidence and report.
- Run `python .\x_create_cv_factory_x.py exercise-golden` when `..\x_create_cv_test_data_x\evidence` is available; expected output includes `_a_priori` count, chain file count, generated JSON count, and generated Office count.
- Run `python .\x_create_cv_factory_x.py check-evidence` when `..\x_create_cv_test_data_x\evidence\a_priori_manifest.json` is available.
- Confirm `x_create_cv_x` does not track private Office evidence and `x_create_cv_test_data_x` owns the private `evidence/` tree.

## Rollback

Restore the `0.0.1` version values and remove `CV_OFFICE_REGENERATION_PLAN.md` from public tracking if reverting this development step. The released `v0.0.1` tag remains the rollback reference for the productionized CLI.

## Operator Notes

Real CV data, private rebuild scripts, generated JSON, `private.zip`, original `_a_priori` Office evidence, private evidence manifests, comparison reports, human approval notes, and generated `_a_posteriori` Office evidence must remain out of the public repository. CI must use fake fixtures only and must never require private CV data. Version `0.0.2` records the private-repo evidence location, chain integrity gate, no-copy evidence rule, and baseline XLSX/DOCX regeneration from golden JSON.