# x_create_cv_x

Privacy-first CLI tooling for building app-native CV and resume data as deterministic JSON.

Version: `0.0.2`

The project keeps source code, docs, tests, and fake fixtures public while real CV data stays in ignored local folders.

## Factory CLI

`x_create_cv_factory_x.py` maintains one master profile JSON file plus one JSON file per resume document:

- `master_profile.json`
- `resume_2017.json`
- `resume_2023.json`

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

## Public Tests And Fake Data

CI and local public validation use only synthetic data from `tests/fixtures/`. The fake fixture proves the schema and CLI workflow without exposing the private CV.

Run the public test suite:

```powershell
python -m pytest
```

## Private Seeds

The private Python seed scripts live under `data/private/` and are intentionally ignored by Git because they contain private CV data. They rebuild the master profile and both resume documents as readable action logs using granular functions such as `add_job(...)`, `add_patent(...)`, `add_resume_section(...)`, and `add_resume_item(...)`.

The real golden Office evidence lives in the side-by-side private repository `x_create_cv_test_data_x`, under `evidence/source_office/a_priori/`, with an `_a_priori` filename suffix. These are private evidence files, not public fixtures. Future whole-cloth regenerated Office files should use an `_a_posteriori` suffix so original evidence and generated evidence are visually distinct.

When `x_create_cv_test_data_x` is checked out next to this repository, run the private golden exercise before Office-regeneration work:

```powershell
python .\x_create_cv_factory_x.py exercise-golden
```

That command fast-fails on evidence corruption by reading both private SHA-256 manifests from `..\x_create_cv_test_data_x\evidence`: the `_a_priori` Office manifest and the full chain-of-evidence manifest covering scripts, generated JSON, legacy references, and original Office files. It then runs the private rebuild scripts into a temporary folder and byte-compares the generated JSON against the stored `_a_posteriori` JSON evidence. The lower-level `_a_priori` check is also available:

```powershell
python .\x_create_cv_factory_x.py check-evidence
```

Copies do not count as generated evidence. The private `_a_posteriori` DOCX/XLSX outputs must eventually be produced by scripts from JSON, then compared against the `_a_priori` Office files.

When private data is available locally, validate the generated JSON against the private golden archive:

```powershell
$env:CV_FACTORY_DB_DIR = Join-Path $PWD 'data\private\cv_factory_output'
python .\data\private\cv_factory_seed_scripts\01_build_master_data.py
python .\data\private\cv_factory_seed_scripts\02_build_old_resume.py
python .\data\private\cv_factory_seed_scripts\03_build_new_resume.py
python .\x_create_cv_factory_x.py validate --db-dir $env:CV_FACTORY_DB_DIR --expected-zip .\data\private\private.zip --zip-prefix private/cv_app_crud
```

Do not commit anything under `data/private/`.

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
- `.vscode/settings.json` with Pylance strict type checking and pytest discovery.
- `.github/workflows/ci.yml` to run tests, Ruff, Black, and mypy on every push and pull request.
- Unit tests that rebuild fake CV JSON and validate zip comparison behavior without private data.

## Documentation

- See [ONLINE_SERVICE_STRATEGY.md](ONLINE_SERVICE_STRATEGY.md) for the future showcase/service path.
- See [CV_OFFICE_REGENERATION_PLAN.md](CV_OFFICE_REGENERATION_PLAN.md) for the `_a_priori` evidence, `_a_posteriori` generation, XLSX, and DOCX regeneration roadmap.
- See [SECURITY.md](SECURITY.md) for private-data handling rules.
- See [CHANGE_CONTROL_PACKET.md](CHANGE_CONTROL_PACKET.md) for the current controlled release record.
- See [CHANGELOG.md](CHANGELOG.md) for release history.

## Privacy Rules

- Public repo files must contain only source, docs, tests, and fake fixtures.
- Real CV data stays under `data/private/`.
- Private seed scripts stay under `data/private/`.
- Private golden archives stay under `data/private/`.
- Real golden evidence, including scripts, generated JSON, manifests, original Office files, legacy references, and future generated Office outputs, stays in the private sibling repo `x_create_cv_test_data_x`.
- CI must never require private data.
