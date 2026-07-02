# x_create_cv_x

Privacy-first CLI tooling for building app-native CV and resume data as deterministic JSON.

Version: `0.0.3`

The project keeps source code, docs, tests, and fake fixtures public while real CV data stays in ignored local folders.

## Factory CLI

`x_create_cv_factory_x.py` maintains one master profile JSON file plus one JSON file per resume document:

- `master_profile.json`
- `resume_2017.json`
- `resume_2023.json`
- `resume_2024.json`

The master profile is built with granular CRUD-style commands such as `add-user`, `add-address`, `add-job`, `add-achievement`, `add-patent`, `add-skill`, and `add-residence`. Resume documents are built with `create-resume`, `add-resume-section`, and `add-resume-item`.

These commands are intentionally shaped like future GUI actions: one command corresponds to one button or form submission.

Show the tool version:

```powershell
python .\x_create_cv_factory_x.py --version
```

Create a small local fake database:

```powershell
$db = Join-Path $PWD 'data\private\cv_factory_output'
python .\x_create_cv_factory_x.py init --db-dir $db --force --id profile_001 --person-id person_001 --label Example --created-at 2026-07-01T00:00:00Z --updated-at 2026-07-01T00:00:00Z
python .\x_create_cv_factory_x.py add-user --db-dir $db --id person_001 --display-name Example --preferred-name Example
python .\x_create_cv_factory_x.py add-job --db-dir $db --id job_001 --label Example --job-title Engineer --employer ExampleCorp
python .\x_create_cv_factory_x.py create-resume --db-dir $db --id resume_2023 --label "2023 Resume" --status active --created-at 2026-07-01T00:00:00Z --updated-at 2026-07-01T00:00:00Z
python .\x_create_cv_factory_x.py add-resume-section --db-dir $db --resume-id resume_2023 --id section_001 --title Summary --kind summary --sort-order 1 --is-visible true
```

The low-level `--json` and `--json-base64` inputs remain available as escape hatches, but the preferred surface is the typed command flags above.

Validate generated JSON against the public contract schemas:

```powershell
python .\x_create_cv_factory_x.py validate-schema data\private\cv_factory_output\master_profile.json data\private\cv_factory_output\resume_2023.json
```

## Public Tests And Fake Data

CI and local public validation use only synthetic data from `tests/fixtures/`. The fake fixture proves the schema and CLI workflow without exposing the private CV.

Run the public test suite:

```powershell
python -m pytest
```

## Private Golden Evidence

The private Python rebuild scripts live in the side-by-side private repository `x_create_cv_test_data_x` under `evidence/scripts/`. They rebuild the master profile and resume documents as readable action logs using granular functions such as `add_job(...)`, `add_patent(...)`, `add_resume_section(...)`, and `add_resume_item(...)`.

The real golden Office evidence lives in the side-by-side private repository `x_create_cv_test_data_x`, under `evidence/source_office/a_priori/`, with an `_a_priori` filename suffix. These are private evidence files, not public fixtures. Future whole-cloth regenerated Office files should use an `_a_posteriori` suffix so original evidence and generated evidence are visually distinct.

When `x_create_cv_test_data_x` is checked out next to this repository, run the private golden exercise before Office-regeneration work:

```powershell
python .\x_create_cv_factory_x.py exercise-golden
```

That command fast-fails on evidence corruption by reading both private SHA-256 manifests from `..\x_create_cv_test_data_x\evidence`: the `_a_priori` Office manifest and the full chain-of-evidence manifest covering scripts, generated JSON, legacy references, and original Office files. It then runs the private rebuild scripts into a temporary folder and byte-compares the generated JSON against the stored `_a_posteriori` JSON evidence. The lower-level `_a_priori` check is also available:

```powershell
python .\x_create_cv_factory_x.py check-evidence
```

Copies do not count as generated evidence. The private `_a_posteriori` DOCX/XLSX outputs are produced from JSON, then reported against the `_a_priori` Office files.

Generate the private `_a_posteriori` Office evidence and comparison report from the side-by-side golden JSON:

```powershell
python .\x_create_cv_factory_x.py generate-golden-office
```

Write the private-safe Office audit reports without regenerating Office outputs:

```powershell
python .\x_create_cv_factory_x.py audit
```

Use an explicit accepted-drift policy when a release has reviewed, versioned differences:

```powershell
python .\x_create_cv_factory_x.py audit --policy .\audit_policies\default_office_audit_policy.json
```

The generated workbook keeps the nine original spreadsheet sheet names, skips redundant collection placeholder rows, uses named app-native fields instead of raw `a`/`b`/`c` columns, and reads sheet/column/style settings from generated JSON `office_layout`, including sheet-level freeze panes, autofilters, explicit column widths, and column value types. The generated DOCX files read document margins/styles, page size, document flow, tables, paragraph alignment/spacing/indentation/tab stops, optional package parts, and per-item `block_style`, `numbering`, and rich `runs` from the same JSON-backed contract.

The `audit` command writes both the machine-readable comparison JSON and a Markdown report with private-safe structure metrics: DOCX package parts, relationship counts, external hyperlink relationship counts, paragraph property counts, run/style counts, hyperlinks, tabs, tables, numbering, XLSX sheet names, dimensions, headers, cell-type counts, and style counts. Audit policy files live under `audit_policies/`; they version accepted drift so review-required differences must be explicitly classified before a release. The default `0.0.2` policy includes one reviewed workbook drift entry for the generated app-native ID/Label columns; all other unmatched differences remain review-required.

The older `validate --expected-zip` command remains available for legacy archive checks, but the active 0.0.2 golden path is `exercise-golden` against `x_create_cv_test_data_x`.

Do not commit real private data to `x_create_cv_x`.

## Development And Release Checks

Install the local development toolchain:

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements-dev.txt
```

Run the full local quality gate before release:

```powershell
python -m pytest
ruff check .
black --check .
mypy
```

Additional CLI checks:

```powershell
python .\x_create_cv_factory_x.py --version
python .\x_create_cv_factory_x.py --help
python -m py_compile .\x_create_cv_factory_x.py
```

The repository includes:

- `pyproject.toml` for project metadata, pytest discovery, Ruff rules, Black formatting, and strict mypy settings.
- `schemas/` with the public JSON Schema contracts for master profile, resume, workbook layout, document layout, and audit policy JSON.
- `audit_policies/` with versioned Office audit policies for explicit accepted-drift classification.
- `.vscode/settings.json` with Pylance strict type checking and pytest discovery.
- `.github/workflows/ci.yml` to run tests, Ruff, Black, and mypy on every push and pull request.
- Unit tests that rebuild fake CV JSON and validate zip comparison behavior without private data.
- Unit tests that generate fake XLSX/DOCX Office evidence without private data.

## Documentation

- See [WORLD_CLASS_CV_GLIDEPATH.md](WORLD_CLASS_CV_GLIDEPATH.md) for the versioned workplan from the current `0.0.2` checkpoint to a stable world-class CV compiler.
- See [RELEASE_NOTES_0.0.3.md](RELEASE_NOTES_0.0.3.md) for the current renderer-conformance development checkpoint.
- See [ONLINE_SERVICE_STRATEGY.md](ONLINE_SERVICE_STRATEGY.md) for the future showcase/service path.
- See [CV_OFFICE_REGENERATION_PLAN.md](CV_OFFICE_REGENERATION_PLAN.md) for the `_a_priori` evidence, `_a_posteriori` generation, XLSX, and DOCX regeneration roadmap.
- See [SECURITY.md](SECURITY.md) for private-data handling rules.
- See [CHANGE_CONTROL_PACKET.md](CHANGE_CONTROL_PACKET.md) for the current controlled release record.
- See [CHANGELOG.md](CHANGELOG.md) for release history.

## Privacy Rules

- Public repo files must contain only source, docs, tests, and fake fixtures.
- Ad hoc local private output under `data/private/` remains ignored.
- Real golden evidence, including scripts, generated JSON, manifests, original Office files, legacy references, and generated Office outputs, stays in the private sibling repo `x_create_cv_test_data_x`.
- CI must never require private data.
