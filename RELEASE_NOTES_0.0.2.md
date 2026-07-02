# Release Notes: 0.0.2 Development Checkpoint

Date: 2026-07-02
Status: Development checkpoint, not a public production release

## Summary

`0.0.2` establishes the private golden-evidence workflow for regenerating CV Office artifacts from JSON contracts, with schema-backed contracts, policy-backed Office audits, and the current 2024 resume as the fourth golden source.

## Highlights

- Added side-by-side private evidence validation through `exercise-golden`.
- Added full private chain-of-evidence validation for source Office files, rebuild scripts, generated JSON, generated Office outputs, reports, and legacy references.
- Added deterministic XLSX/DOCX generation from generated JSON evidence.
- Added the current 2024 resume as private `_a_priori` source evidence and generated a fourth `_a_posteriori` JSON/DOCX pair.
- Improved XLSX output to keep the nine source workbook sheets, use named app-native columns, and skip redundant collection placeholder rows.
- Expanded DOCX layout contracts for document flow, tables, headers/footers, package parts, fonts, hyperlinks, tabs, paragraph alignment, spacing, indentation, tab stops, page size, and rich runs.
- Regenerated private comparison reports with non-sensitive structure metrics.
- Added public JSON Schema contracts, the `validate-schema` CLI, the `audit` CLI, and a versioned accepted-drift policy for reviewed Office parity differences.

## Validation

- `exercise-golden`: 4 `_a_priori` files, 23 chain files, 4 generated JSON files, 4 generated Office files.
- `check-evidence`: 4 source Office files.
- `pytest`: 19 passed.
- `ruff check .`: passed.
- `black --check .`: passed.
- `mypy`: passed.
- Public CI: required on the final pushed checkpoint before any production tag.

## Evidence Checkpoint

Private evidence is checkpointed in `x_create_cv_test_data_x` with the matching evidence tag `v0.0.2-evidence`.

## Release Decision

This checkpoint remains in development. No public `v0.0.2` production tag is cut from `x_create_cv_x` yet.