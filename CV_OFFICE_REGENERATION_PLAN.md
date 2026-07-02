# CV Office Regeneration Plan

Date: 2026-07-01

Version note: `0.0.2` records this roadmap, the private sibling-repo `_a_priori` Office evidence location, the full chain-of-evidence store, the evidence integrity gate, and baseline XLSX/DOCX regeneration from golden JSON. Deeper normalized equivalence and human approval remain the next engineering milestone.

## Goal

Replace the current private `private.zip` golden reference with a stronger evidence workflow based on three original Office files extracted from the source archives and stored under `../x_create_cv_test_data_x/evidence/source_office/a_priori/`:

- `R_cv_2017_1129_0848_a_priori.docx`
- `R_cv_2023_0315_2158_a_priori.docx`
- `R_cv_2023_0501_1427_a_priori.xlsx`

The `_a_priori` suffix marks original source evidence known before regeneration. Whole-cloth generated XLSX and DOCX outputs use an `_a_posteriori` suffix.

Copies are not valid `_a_posteriori` evidence. The generated Office outputs must be produced by private scripts from the app-native JSON chain, then compared against the `_a_priori` Office files.

The target end state is that `x_create_cv_x` can rebuild the app-native CV database, regenerate the spreadsheet database, regenerate the 2017 resume DOCX, regenerate the 2023 resume DOCX, and compare those outputs against the original source evidence.

## Position

This is the right next milestone. The current `private.zip` proves that the new app-native JSON can match a private snapshot, but it is still a derived artifact. The original source archives are better evidence because they represent the handwritten historical materials. Once the code can regenerate the JSON, XLSX, and DOCX outputs from whole cloth and compare them against those originals, the project moves from extraction/prototype confidence to reproducible-document confidence.

The XLSX comparison should allow known historical drift because the original spreadsheet may be missing records that are present elsewhere. The DOCX comparison should aim for practical identity first: identical text, section order, table/list structure, styles that affect visible output, document properties, and normalized Office Open XML package content. Byte-for-byte DOCX identity may require controlling ZIP member order, timestamps, relationship IDs, generated revision IDs, and document metadata, so it should be treated as a stretch target after normalized equivalence passes.

## Evidence Model

Keep four evidence classes together locally, but do not commit private content to public Git:

1. Source evidence: the three original `_a_priori` Office files under `../x_create_cv_test_data_x/evidence/source_office/a_priori/`.
2. Rebuild scripts: present-tense Python scripts under `evidence/scripts/` that create master data, the spreadsheet database, and each resume document from code.
3. Generated JSON evidence: `_a_posteriori` master/profile/resume JSON produced by the scripts.
4. Regenerated Office evidence: new whole-cloth `_a_posteriori` XLSX and DOCX files produced by the scripts from JSON.
5. Comparison evidence: SHA-256 manifests, normalized manifests, hashes, diffs, validation reports, and human approval notes proving what matches exactly and what intentionally differs.

Legacy ignored local layout inside `x_create_cv_x`:

```text
data/private/evidence/generated/
data/private/evidence/normalized/
data/private/evidence/reports/
data/private/evidence/scripts/
```

Recommended private test-data repo layout inside `x_create_cv_test_data_x`:

```text
evidence/a_priori_manifest.json
evidence/chain_of_evidence_manifest.json
evidence/source_office/a_priori/
evidence/scripts/
evidence/generated/
evidence/normalized/
evidence/reports/
evidence/legacy/
```

The public repo should keep only fake fixtures, public code, public tests, and documentation that describes the workflow without exposing real CV data.

## Target Outputs

The whole-cloth rebuild produces `_a_posteriori` outputs:

- `master_profile.json`
- `resume_2017.json`
- `resume_2023.json`
- `master_profile_a_posteriori.xlsx` generated from `master_profile.json`
- `resume_2017_a_posteriori.docx` generated from `resume_2017.json` plus shared profile data
- `resume_2023_a_posteriori.docx` generated from `resume_2023.json` plus shared profile data

The `_a_priori` source documents should be used as comparison evidence, not as runtime dependencies for normal generation.

## Acceptance Criteria

The full Office-regeneration milestone is complete when all of the following are true:

- The current `private.zip` validation path is replaced or downgraded to a legacy compatibility check.
- `exercise-golden` passes against both private SHA-256 manifests in the side-by-side `x_create_cv_test_data_x` checkout before any Office regeneration exercise proceeds.
- `exercise-golden` runs the private rebuild scripts into a temporary directory and compares generated JSON against stored `_a_posteriori` JSON evidence.
- Private seed scripts, generated JSON, legacy references, and Office files are all represented in `chain_of_evidence_manifest.json`.
- The original Office files use the `_a_priori` suffix and remain under ignored local evidence folders.
- The factory can generate the XLSX database from `master_profile.json` without manual Office editing.
- The factory can generate the 2017 DOCX from app-native JSON without manual Office editing.
- The factory can generate the 2023 DOCX from app-native JSON without manual Office editing.
- DOCX comparison passes under a normalized Office Open XML comparison.
- Any remaining DOCX byte-level differences are explained in a generated report.
- XLSX comparison passes with an explicit allowed-drift report for records missing from the historical spreadsheet.
- Public CI still runs only fake fixtures and does not require real source zips or private documents.
- Private evidence and generated Office files remain ignored by Git.

## Implementation Phases

### Phase 1: Evidence Inventory

- Use the three `_a_priori` Office files under `x_create_cv_test_data_x/evidence/source_office/a_priori/`.
- Maintain private SHA-256 manifests with file names, sizes, hashes, Office document types, generated JSON, scripts, legacy references, and selected package metadata.
- Run `exercise-golden` before any private Office regeneration exercise.

### Phase 2: Normalizers And Comparators

- Add private-safe comparison helpers for DOCX and XLSX Office Open XML packages.
- Normalize volatile package data such as ZIP timestamps, relationship ordering, document properties that change on save, and generated application metadata.
- Produce readable comparison reports that separate visible-content mismatches from harmless package-level drift.
- Keep public tests on synthetic DOCX/XLSX fixtures or small generated fake documents.

### Phase 3: XLSX Generation

- Define the workbook schema generated from `master_profile.json`.
- Generate worksheets for the main master-profile collections.
- Preserve stable IDs, ordering fields, source notes, and cross-record relationships.
- Compare generated workbook data against the original spreadsheet with an allowed-drift mechanism for known missing historical records.

### Phase 4: DOCX Generation

- Define DOCX templates or code-driven layout builders for the 2017 and 2023 resume variants.
- Generate visible text, section order, list/table structure, styles, margins, headers/footers, and document properties from JSON.
- Compare generated DOCX files against the original DOCX files using normalized Office Open XML comparison.
- Promote any required hand-authored layout knowledge into explicit code or private templates.

### Phase 5: Factory Integration

- Add CLI commands that match the current button-like domain-action style:
  - `check-evidence`
  - `generate-xlsx`
  - `generate-docx`
  - `compare-office`
  - `rebuild-evidence`
- Keep generation deterministic: stable ordering, stable IDs, stable timestamps where the file format allows it, and stable metadata.
- Keep commands composable so the future UI can expose them as buttons without hidden state.

### Phase 6: Evidence Bundle

- Keep the private local evidence bundle in `x_create_cv_test_data_x`; it contains `_a_priori` Office files, rebuild scripts, regenerated `_a_posteriori` XLSX/DOCX files, normalized manifests, SHA-256 manifests, and comparison reports.
- Store enough metadata to prove which code commit generated the evidence.
- Keep the evidence bundle ignored locally unless a separate private repository is intentionally created for it.

## Recommended Technical Direction

- Use `zipfile` and XML parsing from the Python standard library for initial Office package inspection and normalized comparison.
- Add `openpyxl` only when XLSX writing becomes necessary.
- Add `python-docx` only if code-driven DOCX generation is sufficient for the required layout fidelity.
- If exact DOCX layout fidelity exceeds what `python-docx` can preserve, use private DOCX templates plus deterministic XML patching instead of trying to model all Word layout behavior from scratch.

## Risks

- DOCX byte-for-byte identity may not be realistic without preserving or recreating volatile Word metadata and package details.
- The original XLSX may not be a complete source of truth, so comparison must classify missing records as expected drift rather than failure.
- Office tools can rewrite files with different package structure even when visible content is unchanged.
- Any private evidence workflow must stay outside public Git and public CI.

## First Concrete Task

Use `exercise-golden` as the first step of every private Office regeneration exercise. The next implementation task is to add normalized Office Open XML comparison for `_a_priori` and future `_a_posteriori` files.