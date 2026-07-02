# Change Control Packet

Tool: `x_create_cv_x`
Version: `0.0.2`
Packet ID: `x_create_cv_x-0.0.2-2026-07-01`
Date: `2026-07-01`
Status: Updated

## Scope

Establish the next controlled development version for the private Office source-evidence workflow and whole-cloth XLSX/DOCX regeneration roadmap.

## Controlled Files

- `x_create_cv_factory_x.py`
- `README.md`
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
- Added `CV_OFFICE_REGENERATION_PLAN.md` as the controlled roadmap for replacing or downgrading `private.zip` with source-ZIP evidence.
- Documented the ignored private evidence location `data/private/evidence/source_zips/` for the original source ZIP archives.
- Moved the three original source ZIP archives into the ignored private evidence boundary locally.
- Updated README, changelog, and security policy notes for the Office-regeneration evidence workflow.
- Preserved public CI as fake-fixture only; private Office source archives remain local validation evidence.

## Validation

- Run `python .\x_create_cv_factory_x.py --version`.
- Run `python .\x_create_cv_factory_x.py --help`.
- Run `python -m py_compile .\x_create_cv_factory_x.py`.
- Run `python -m pytest`.
- Run `ruff check .`.
- Run `black --check .`.
- Run `mypy`.
- Run private validation locally with `python .\x_create_cv_factory_x.py validate --db-dir .\data\private\cv_factory_output --expected-zip .\data\private\private.zip --zip-prefix private/cv_app_crud` when private data is available.
- Run `git check-ignore -v data/private/evidence/source_zips/R_cv_2017_1129_0848.zip data/private/evidence/source_zips/R_cv_2023_0315_2158.zip data/private/evidence/source_zips/R_cv_2023_0501_1427.zip`.

## Rollback

Restore the `0.0.1` version values and remove `CV_OFFICE_REGENERATION_PLAN.md` from public tracking if reverting this development step. The released `v0.0.1` tag remains the rollback reference for the productionized CLI.

## Operator Notes

Real CV data, private seed scripts, `private.zip`, original source ZIP archives, and generated Office evidence must remain under ignored local paths. CI must use fake fixtures only and must never require private CV data. Version `0.0.2` records the roadmap and evidence location; the actual XLSX/DOCX regeneration implementation remains future work.