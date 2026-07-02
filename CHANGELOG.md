# Changelog

## 0.0.2 - 2026-07-01

### Added

- Added the Office regeneration roadmap for source-ZIP evidence, XLSX generation, DOCX generation, and normalized comparison reports.
- Documented the ignored private evidence location for the original source ZIP archives.
- Linked the Office regeneration plan from the README.

### Changed

- Bumped the development version to `0.0.2` after the `v0.0.1` release.
- Reframed `private.zip` as a legacy/private validation reference to be replaced or downgraded by source-ZIP evidence.

### Security

- Kept original source ZIP archives under the ignored `data/private/evidence/source_zips/` boundary.
- Clarified that source ZIP archives are private local evidence and are not public fixtures or CI inputs.

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