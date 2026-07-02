# Release Notes: 0.0.3 Development Checkpoint

Date: 2026-07-02
Status: Development checkpoint, not a public production release

## Summary

`0.0.3` starts renderer conformance hardening after the `0.0.2` contract-and-audit lock. This checkpoint keeps the public/private evidence split intact while making XLSX sheet behavior more explicit, testable, and visible in audit reports.

## Highlights

- Added schema-backed per-sheet XLSX `freeze_header`, `auto_filter`, and `column_widths` controls.
- Added schema-backed XLSX column `value_type` controls for string, number, and boolean cell rendering.
- Made generated XLSX worksheets honor sheet-level controls over workbook defaults.
- Expanded XLSX structure summaries with freeze-pane, autofilter, and column-width metadata.
- Expanded XLSX audit metrics with cell-type counts.
- Expanded XLSX audit metrics with font, fill, border, and cell-format style details.
- Expanded XLSX audit metrics with package parts, content type overrides, and relationship summaries.
- Expanded XLSX audit metrics with generated worksheet page margins.
- Expanded the Markdown Office audit report to show generated XLSX freeze-pane and autofilter state.
- Added explicit DOCX run hyperlink/style fields to the public document layout schema.
- Expanded DOCX audit metrics with relationship counts and external hyperlink relationship counts.
- Expanded DOCX audit metrics with paragraph style, alignment, spacing, indentation, and tab-stop counts.
- Expanded DOCX audit metrics with table row, cell, paragraph, grid-width, and cell-width geometry.
- Expanded DOCX audit metrics with numbering definitions, IDs, levels, bullet text patterns, and numbering fonts.
- Expanded DOCX audit metrics with section page size, margins, and header/footer references.
- Expanded DOCX audit metrics with style definitions, defaults, based-on links, run properties, paragraph properties, and style numbering.
- Added public fake-fixture coverage for the new workbook conformance controls, including typed number and boolean cells.
- Added public fake-fixture coverage for XLSX style color and cell-format audit metrics.
- Added public fake-fixture coverage for XLSX package part and relationship audit metrics.
- Added public fake-fixture coverage for XLSX page-margin audit metrics.
- Added public fake-fixture coverage for external hyperlink relationship conformance.
- Added public fake-fixture coverage for DOCX paragraph property audit metrics.
- Added public fake-fixture coverage for DOCX table width audit metrics.
- Added public fake-fixture coverage for DOCX numbering definition audit metrics.
- Added public fake-fixture coverage for DOCX page setup audit metrics.
- Added public fake-fixture coverage for DOCX style definition audit metrics.

## Validation

- `validate-schema`: default audit policy passed.
- `check-evidence`: 4 source Office files.
- `exercise-golden`: 4 `_a_priori` files, 23 chain files, 4 generated JSON files, 4 generated Office files.
- `pytest`: 19 passed.
- `ruff check .`: passed.
- `black --check x_create_cv_factory_x.py tests/test_x_create_cv_factory_x.py`: passed.
- `mypy`: passed.

## Evidence Checkpoint

Private evidence is checkpointed in `x_create_cv_test_data_x` with regenerated policy-backed audit reports for the `0.0.3` renderer-conformance metrics.

## Release Decision

This checkpoint remains in development. No public `v0.0.3` production tag is cut from `x_create_cv_x` yet.