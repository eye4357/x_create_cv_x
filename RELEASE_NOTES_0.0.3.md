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
- Expanded XLSX audit reports with generated row counts, column counts, styled-cell counts, and column widths.
- Expanded XLSX audit reports with generated package part names and workbook sheet names.
- Expanded XLSX audit reports with generated worksheet paths.
- Expanded XLSX audit reports with source worksheet paths alongside generated worksheet paths.
- Expanded XLSX audit reports with source row and column counts alongside generated counts.
- Expanded XLSX audit reports with source styled-cell counts alongside generated counts.
- Expanded XLSX audit reports with source cell-type counts alongside generated counts.
- Expanded XLSX audit reports with source freeze-pane and autofilter state alongside generated state.
- Expanded XLSX audit reports with source page margins and column widths alongside generated values.
- Expanded XLSX audit reports with source worksheet names alongside generated worksheet names.
- Expanded the Markdown Office audit report to show generated XLSX freeze-pane and autofilter state.
- Added explicit DOCX run hyperlink/style fields to the public document layout schema.
- Expanded DOCX audit reports with package flags, body child counts, font table names, and used styles.
- Expanded DOCX audit metrics with relationship counts and external hyperlink relationship counts.
- Expanded DOCX audit metrics with paragraph style, alignment, spacing, indentation, and tab-stop counts.
- Expanded DOCX audit metrics with table row, cell, paragraph, grid-width, and cell-width geometry.
- Expanded DOCX audit metrics with numbering definitions, IDs, levels, bullet text patterns, and numbering fonts.
- Expanded DOCX audit metrics with section page size, margins, and header/footer references.
- Expanded DOCX audit metrics with style definitions, defaults, based-on links, run properties, paragraph properties, and style numbering.
- Added public fake-fixture coverage for the new workbook conformance controls, including typed number and boolean cells.
- Added public fake-fixture coverage for XLSX style color and cell-format audit metrics.
- Added public fake-fixture coverage for XLSX style font-name, cell format count, and cell style count audit metrics.
- Added public fake-fixture coverage for XLSX default, header, and body cell format attributes.
- Added public fake-fixture coverage for XLSX header and inline-string body style application.
- Added public fake-fixture coverage for XLSX worksheet XML controls including widths, freeze panes, filters, and margins.
- Added public fake-fixture coverage for XLSX workbook sheet and relationship binding.
- Added public fake-fixture coverage for XLSX content-type and root relationship package binding.
- Added public fake-fixture coverage for XLSX core and app document property metadata.
- Added public fake-fixture coverage for XLSX style XML color, border, and alignment binding.
- Added public fake-fixture coverage for XLSX worksheet-level XML structure and row ordering.
- Added public fake-fixture coverage for DOCX styles XML definitions and list-style binding.
- Added public fake-fixture coverage for DOCX numbering XML abstract and concrete numbering bindings.
- Added public fake-fixture coverage for DOCX content-type and root relationship package binding.
- Added public fake-fixture coverage for DOCX optional content-type package binding.
- Added public fake-fixture coverage for DOCX baseline and optional document relationship binding.
- Added public fake-fixture coverage for exact DOCX document relationship XML with optional package parts and hyperlink targets.
- Added public fake-fixture coverage for DOCX settings part and default tab stop XML.
- Added public fake-fixture coverage for DOCX header and footer part XML.
- Added public fake-fixture coverage for DOCX custom XML item, properties, and relationship part XML.
- Added public fake-fixture coverage for DOCX footnotes and endnotes part XML.
- Added public fake-fixture coverage for DOCX font table and web settings part XML.
- Added public fake-fixture coverage for DOCX theme color, font, and line-style XML.
- Added public fake-fixture coverage for DOCX package root relationship XML.
- Added public fake-fixture coverage for DOCX content-type default declarations.
- Added public fake-fixture coverage for enabled DOCX content-type override XML.
- Added public fake-fixture coverage for DOCX document root namespace envelope XML.
- Added public fake-fixture coverage for DOCX default section page size and margin XML.
- Added public fake-fixture coverage for DOCX table width and border property XML.
- Added public fake-fixture coverage for XLSX content-type default declarations.
- Added public fake-fixture coverage for XLSX package root relationship XML.
- Added public fake-fixture coverage for XLSX workbook relationship XML.
- Added public fake-fixture coverage for XLSX workbook root and sheet XML.
- Added public fake-fixture coverage for XLSX content-type override XML.
- Added public fake-fixture coverage for DOCX core and app document property metadata.
- Added public fake-fixture coverage for DOCX table grid and cell width XML.
- Added public fake-fixture coverage for DOCX section page margin XML.
- Added public fake-fixture coverage for DOCX run property and text binding.
- Added public fake-fixture coverage for DOCX numbering level and numbering ID binding.
- Added public fake-fixture coverage for DOCX hyperlink relationship and formatted run binding.
- Added public fake-fixture coverage for DOCX package relationship ID, type, and target binding.
- Added public fake-fixture coverage for DOCX header and footer part content.
- Added public fake-fixture coverage for DOCX font table part content.
- Added public fake-fixture coverage for optional DOCX package part XML content.
- Added public fake-fixture coverage for DOCX paragraph property binding.
- Added public fake-fixture coverage for DOCX tab run and following text adjacency.
- Added public fake-fixture coverage for DOCX table cell content binding.
- Added public fake-fixture coverage for explicit DOCX empty table paragraph binding.
- Added public fake-fixture coverage for XLSX package part and relationship audit metrics.
- Added public fake-fixture coverage for XLSX page-margin audit metrics.
- Added public fake-fixture coverage for XLSX row, column, styled-cell, and column-width report metrics.
- Added public fake-fixture coverage for XLSX package part-name and sheet-name report metrics.
- Added public fake-fixture coverage for XLSX worksheet-path report metrics.
- Added public fake-fixture coverage for XLSX source worksheet-path report metrics.
- Added public fake-fixture coverage for XLSX source row-count and column-count report metrics.
- Added public fake-fixture coverage for XLSX source styled-cell-count report metrics.
- Added public fake-fixture coverage for XLSX source cell-type-count report metrics.
- Added public fake-fixture coverage for XLSX source freeze-pane and autofilter report metrics.
- Added public fake-fixture coverage for XLSX source page-margin and column-width report metrics.
- Added public fake-fixture coverage for XLSX source worksheet-name report metrics.
- Added public fake-fixture coverage for DOCX package flags, body child counts, font table names, and used styles.
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