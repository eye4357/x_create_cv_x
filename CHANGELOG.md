# Changelog

## 0.0.2 - 2026-07-01

### Added

- Added the Office regeneration roadmap for `_a_priori` evidence, XLSX generation, DOCX generation, and normalized comparison reports.
- Added `check-evidence` to fast-fail when private `_a_priori` Office evidence does not match its SHA-256 manifest.
- Added `exercise-golden` as the simple side-by-side smoke test for `x_create_cv_test_data_x` private golden evidence.
- Extended `exercise-golden` to validate the full private chain-of-evidence manifest, including scripts, generated JSON, legacy references, and original Office evidence.
- Extended `exercise-golden` to run private rebuild scripts in a temporary directory and compare generated JSON against stored `_a_posteriori` JSON evidence.
- Added deterministic stdlib XLSX/DOCX generation for private golden `_a_posteriori` Office evidence.
- Added `generate-golden-office` to create private Office evidence and a comparison report from golden JSON.
- Improved generated XLSX evidence to use the same nine workbook sheets as the `_a_priori` spreadsheet, with structured data added after the source-shaped columns.
- Tuned generated DOCX evidence toward `_a_priori` page setup and paragraph styles, including source-like margins and list/title/heading styles.
- Promoted workbook/document layout into explicit generated JSON `office_layout` contracts consumed by the Office renderer.
- Removed confusing raw `a`/`b`/`c` spreadsheet columns from generated XLSX output in favor of named app-native fields.
- Updated DOCX generation to consume item `formatting.block_style`, `formatting.numbering`, and rich `formatting.runs` from the JSON contract.
- Added DOCX document-flow layout support for JSON-backed paragraph, item, and table blocks.
- Added DOCX header/footer, theme, font table, and web settings package parts driven by the document layout contract.
- Hardened generated DOCX theme packages for stricter importers such as Google Drive.
- Added non-sensitive Office structure metrics to the private comparison report.
- Documented the ignored private evidence location for the original `_a_priori` Office files.
- Linked the Office regeneration plan from the README.

### Changed

- Bumped the development version to `0.0.2` after the `v0.0.1` release.
- Reframed `private.zip` as a legacy/private validation reference to be replaced or downgraded by `_a_priori` Office evidence.
- Replaced the local archive evidence layout with extracted `_a_priori` DOCX/XLSX files; generated Office files use `_a_posteriori`.
- Moved the real `_a_priori` Office evidence contract from the public repo's ignored local folder to the private sibling repository `x_create_cv_test_data_x`.
- Moved private seed scripts, generated JSON, legacy JSON, and the legacy private zip out of the public repo and into `x_create_cv_test_data_x`.

### Security

- Kept original `_a_priori` Office files and their SHA-256 manifest under the ignored `data/private/evidence/` boundary.
- Clarified that `_a_priori` and `_a_posteriori` Office files are private local evidence and are not public fixtures or CI inputs.
- Kept the public repository free of real golden evidence while allowing local validation against the private test-data repository.
- Documented that `_a_posteriori` Office files must be generated from scripts and JSON, not copied from `_a_priori` files.
- Extended `exercise-golden` to regenerate private Office evidence into a temporary directory and byte-compare it against stored `_a_posteriori` Office evidence.

## 0.0.1 - 2026-07-01

### Added

- Initial productionized release of the CV factory CLI.
- Explicit `--version` output.
- Public-safe fake fixture coverage for deterministic JSON generation and zip validation.
- Pinned development tooling for Ruff, Black, mypy, and pytest.
- GitHub Actions quality gate.
- Security notes for private CV data handling.
- Online service strategy for future showcase/service work.

### Security

- Documented that real CV data, private seed scripts, generated private JSON, and golden private archives stay out of Git.
- Added fake fixture workflow for CI so public checks do not require private data.