# World-Class CV Creation Glidepath

Date: 2026-07-02
Starting point: `0.0.2` development checkpoint

## Product Thesis

`x_create_cv_x` should become a privacy-first CV compiler: it ingests career evidence, normalizes it into explicit app-native contracts, renders polished Office outputs, and proves every generated claim and formatting choice through auditable evidence.

The project should keep the public/private split that already works:

- Public repository: engine, schemas, fake fixtures, docs, tests, CI, and examples that contain no real CV data.
- Private evidence repository: real `_a_priori` sources, rebuild scripts, generated `_a_posteriori` outputs, manifests, reports, and human approval notes.
- Release rule: public releases must be reproducible with fake data; private evidence checkpoints may be tagged separately when real evidence is regenerated.

## Non-Negotiable Design Rules

- Formatting promises belong in explicit JSON contracts and schemas, not in hidden renderer shortcuts.
- Generated Office files must be created from JSON and renderer code; copying `_a_priori` files is not valid evidence.
- Public tests must exercise the same code paths as private evidence using fake fixtures only.
- Every release must include a human-readable audit report for the relevant evidence level.
- Private evidence must never be required by public CI.

## Version Glidepath

### `0.0.2` - Contract And Audit Lock

Goal: finish the current development checkpoint as the first honest API-and-parity baseline.

Scope:

- Add explicit JSON Schemas for master profile, resume, workbook layout, document layout, sections, items, rich runs, numbering, tables, page setup, and package options.
- Add a first-class `audit` command that emits human-readable DOCX/XLSX comparison reports.
- Promote known acceptable differences into versioned audit policy files instead of ad hoc prose.
- Keep `exercise-golden`, `check-evidence`, and fake public tests green.
- Cut a public `v0.0.2` tag only after the contract and audit gates pass; until then, keep the current checkpoint as development-only.

Acceptance gate:

- JSON Schema validation passes for all fake fixtures and private golden JSON.
- `audit` reports DOCX package parts, body shape, paragraph counts, text equality, fonts, runs, hyperlinks, tabs, tables, numbering, XLSX sheets, headers, dimensions, and styles.
- Private evidence has a matching private evidence tag for the final public `v0.0.2` candidate.

### `0.0.3` - Renderer Conformance Hardening

Goal: make the Office renderers visibly dependable before adding new product surface area.

Scope:

- Expand renderer conformance tests for the schema fields added in `0.0.2`.
- Add compatibility checks for Word, Google Drive, and LibreOffice import behavior where practical.
- Add stricter normalized OpenXML comparisons for volatile package metadata.
- Improve DOCX numbering, run, paragraph, table, and relationship parity using only schema-backed inputs.
- Improve XLSX style, width, freeze-pane, filter, data-type, and sheet-level parity using only schema-backed inputs.

Acceptance gate:

- No known renderer behavior relies on undocumented layout defaults.
- Public fake fixtures include enough layout complexity to catch regressions without private data.
- Private parity reports classify every mismatch as pass, fail, or accepted drift.

### `0.1.0` - Template And Theme System

Goal: separate core career data from intentional output design.

Scope:

- Introduce named templates that generate or transform layout contracts without mutating core profile data.
- Ship initial template modes: faithful reproduction, modern executive CV, technical resume, academic long-form CV, and ATS-safe plain resume.
- Add template metadata: intended audience, supported formats, page targets, typography assumptions, and known limitations.
- Add CLI support such as `render --template faithful` and `render --template ats-safe`.

Acceptance gate:

- The same profile can render multiple outputs by changing only the template/layout contract.
- Template behavior is covered by fake fixtures and schema validation.
- The faithful template remains parity-auditable against private `_a_priori` evidence.

### `0.1.1` - Template Polish Pass

Goal: raise visual quality without expanding the core product contract.

Scope:

- Tune spacing, typography, hierarchy, lists, and table styles for each template.
- Add visual approval notes to private evidence reports where real outputs are reviewed by a human.
- Add focused regression fixtures for template edge cases such as long roles, dense skills, patents, publications, and multi-page resumes.

Acceptance gate:

- Each template has a clear use case and does not look like a thin restyle of the others.
- Long content remains readable and does not cause obvious layout collapse.

### `0.2.0` - Authoring Workflow CLI

Goal: make the tool pleasant to use without hand-editing large JSON files.

Scope:

- Add import and authoring commands such as `import-docx`, `import-xlsx`, `create-profile`, `create-resume`, `render`, and `audit --against`.
- Support editable intermediate files for review before committing generated data.
- Add dry-run, diff, and explain modes for commands that mutate profile or resume JSON.
- Keep commands shaped like future UI actions.

Acceptance gate:

- A user can create or update a profile from source documents through CLI commands and review the resulting changes.
- Generated JSON remains deterministic and schema-valid.
- CLI errors identify the file, field, and recovery step when possible.

### `0.2.1` - Import Quality Pass

Goal: improve messy evidence intake.

Scope:

- Improve DOCX text, run, hyperlink, tab, section, and page-layout extraction.
- Improve XLSX sheet, table, header, and record extraction.
- Add import confidence warnings for ambiguous records, duplicate facts, and unsupported formatting.
- Add private golden import reports for the known source documents.

Acceptance gate:

- Import reports separate extracted facts, inferred structure, unsupported formatting, and manual-review warnings.
- Unsupported features do not disappear silently.

### `0.3.0` - Role-Targeted Resume Generation

Goal: turn the compiler into a resume strategist while preserving evidence discipline.

Scope:

- Add structured target-role input with job title, industry, keywords, emphasis, max pages, tone, must-include evidence, and must-exclude evidence.
- Add scoring and selection rules for choosing profile evidence for a target resume.
- Add commands such as `create-resume --target role.yml` and `explain-targeting`.
- Preserve provenance from every generated section and bullet back to source profile records.

Acceptance gate:

- A generated targeted resume includes an explanation of why each major item was selected.
- The user can override selections without breaking provenance or schema validation.
- Target-role logic is tested with public fake profiles.

### `0.3.1` - Targeting Quality Pass

Goal: make role targeting useful for real revision cycles.

Scope:

- Add ranking diagnostics for omitted-but-relevant evidence.
- Add keyword coverage reports that distinguish exact matches, synonyms, and unsupported claims.
- Add style controls for concise, executive, technical, academic, and ATS-safe language.

Acceptance gate:

- Reports make it clear what changed between two targeted resumes and why.
- The tool does not invent unsupported accomplishments.

### `0.4.0` - Field-Level Provenance And Privacy Guarantees

Goal: make trust a product feature.

Scope:

- Add field-level provenance links from generated claims, bullets, and summary statements back to source records.
- Add output manifests that include code commit, schema version, template version, source evidence IDs, generated file hashes, and audit status.
- Add privacy leak checks for public fixtures, docs, test outputs, and generated reports.
- Add separate public and private release checklists.

Acceptance gate:

- Every substantive generated claim can be traced to source evidence or marked as user-authored.
- Public CI includes a privacy guard that prevents accidental fixture or report leaks.
- Private evidence reports can be shared internally without exposing more than intended.

### `0.5.0` - Local Review Experience

Goal: make review and iteration faster before building a broader UI.

Scope:

- Add local review artifacts: HTML audit reports, side-by-side evidence summaries, and generated output indexes.
- Add commands for opening the latest generated outputs and reports from the evidence directory.
- Add a review checklist that records human approval decisions without committing private content publicly.

Acceptance gate:

- A reviewer can inspect source/generation/audit status from one local report entry point.
- Human approval notes are tracked in private evidence without leaking into public fixtures.

### `0.6.0` - Interface Beta

Goal: put the proven engine behind a simple usable interface.

Scope:

- Choose one interface path: local desktop-style app around files, or a simple web UI that runs locally first.
- Expose core flows: choose evidence, import, edit profile records, choose template, target a role, render, audit, export.
- Keep the CLI as the underlying automation boundary.

Acceptance gate:

- The interface can complete a real local end-to-end flow without requiring private data in the public repo.
- CLI and UI outputs are equivalent for the same inputs.

### `0.7.0` - Export And Compatibility Expansion

Goal: support the formats and environments users actually need.

Scope:

- Add PDF export strategy after DOCX is stable.
- Add ATS-safe text or markdown export.
- Add compatibility notes and smoke checks for common import paths.
- Add packaging instructions for non-developer use.

Acceptance gate:

- DOCX, XLSX, PDF, and ATS-safe exports are generated from the same app-native data model.
- Export limitations are explicit and tested where possible.

### `0.8.0` - Productization And Distribution

Goal: make installation, upgrades, and evidence migration boring.

Scope:

- Add data migration tooling for schema upgrades.
- Add project initialization commands for new users and private evidence stores.
- Add packaged installation path and release artifacts.
- Add upgrade notes for each persisted schema version.

Acceptance gate:

- Existing local profiles can be migrated forward safely.
- A new user can initialize a clean private workspace without learning the internal repository layout first.

### `0.9.0` - Release Candidate

Goal: freeze the world-class baseline and focus on defects.

Scope:

- Freeze public schema contracts for `1.0.0`.
- Freeze template contract versions.
- Run full fake public gate and private evidence gate.
- Resolve all release-blocking audit, privacy, import, render, and interface defects.

Acceptance gate:

- No known critical or high release blockers remain.
- Documentation covers install, authoring, rendering, auditing, privacy, and evidence workflows.

### `1.0.0` - World-Class Baseline

Goal: release the first stable version of the CV compiler.

Scope:

- Stable public schema and CLI contracts.
- Faithful Office regeneration for golden evidence within documented accepted drift.
- Multiple high-quality templates.
- Role-targeted generation with provenance.
- Strong privacy posture and public/private evidence separation.
- Local review workflow and at least one practical interface path.

Acceptance gate:

- Public CI is green on fake fixtures.
- Private evidence gate passes on the real golden corpus.
- Release notes include known limitations, accepted drift, schema versions, template versions, and evidence tag references.
- The tool can take a real evidence set from import through targeted export with auditable output.

## Release Discipline

For every public version:

- Update `pyproject.toml`, `README.md`, `CHANGELOG.md`, and release notes.
- Run public tests, Ruff, Black, mypy, and the relevant CLI smoke checks.
- Keep public fixtures synthetic.
- Record whether a matching private evidence tag exists.
- Do not cut a public production tag until the version acceptance gate is satisfied.

For every private evidence checkpoint:

- Regenerate private JSON and Office outputs from scripts.
- Update evidence manifests and audit reports.
- Tag the private evidence repository with the matching evidence tag.
- Keep real evidence out of the public repository.

## Next Engineering Move

The next concrete implementation slice should be `0.0.2` contract-and-audit lock:

1. Add JSON Schema files for the public data and layout contracts.
2. Add schema validation to fake tests and private golden exercise.
3. Add a first-class `audit` CLI command that emits readable DOCX/XLSX parity reports.
4. Version known acceptable differences so release decisions are evidence-backed rather than vibe-backed.