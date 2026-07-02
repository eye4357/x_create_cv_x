# Change Control Packet

Tool: `x_create_cv_x`
Version: `0.0.2`
Packet ID: `x_create_cv_x-0.0.2-2026-07-01`
Date: `2026-07-01`
Status: Updated

## Scope

Establish the next controlled development version for the side-by-side private golden evidence workflow, `_a_priori` source-document integrity checks, and whole-cloth `_a_posteriori` XLSX/DOCX regeneration roadmap.

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
- Added `CV_OFFICE_REGENERATION_PLAN.md` as the controlled roadmap for replacing or downgrading `private.zip` with Office evidence.
- Extracted the three local source archives into original `_a_priori` DOCX/XLSX files under the ignored private evidence boundary, then removed the ZIP archives locally.
- Added a private SHA-256 manifest for the three `_a_priori` Office files.
- Added `check-evidence` to fast-fail when private `_a_priori` Office evidence is missing, size-mismatched, or hash-mismatched.
- Added `exercise-golden` as the simple local smoke test for the side-by-side private repository `x_create_cv_test_data_x`.
- Documented the `_a_priori` suffix for original source evidence and the `_a_posteriori` suffix for future whole-cloth generated Office files.
- Moved the real `_a_priori` Office evidence and SHA-256 manifest ownership into `x_create_cv_test_data_x`.
- Updated README, changelog, and security policy notes for the Office-regeneration evidence workflow.
- Preserved public CI as fake-fixture only; private Office evidence is exercised locally only when the private repo is present.

## Validation

- Run `python .\x_create_cv_factory_x.py --version`.
- Run `python .\x_create_cv_factory_x.py --help`.
- Run `python -m py_compile .\x_create_cv_factory_x.py`.
- Run `python -m pytest`.
- Run `ruff check .`.
- Run `black --check .`.
- Run `mypy`.
- Run `python .\x_create_cv_factory_x.py exercise-golden` when `..\x_create_cv_test_data_x\evidence` is available.
- Run `python .\x_create_cv_factory_x.py check-evidence` when `..\x_create_cv_test_data_x\evidence\a_priori_manifest.json` is available.
- Run private validation locally with `python .\x_create_cv_factory_x.py validate --db-dir .\data\private\cv_factory_output --expected-zip .\data\private\private.zip --zip-prefix private/cv_app_crud` when private data is available.
- Confirm `x_create_cv_x` does not track private Office evidence and `x_create_cv_test_data_x` owns the private `evidence/` tree.

## Rollback

Restore the `0.0.1` version values and remove `CV_OFFICE_REGENERATION_PLAN.md` from public tracking if reverting this development step. The released `v0.0.1` tag remains the rollback reference for the productionized CLI.

## Operator Notes

Real CV data, private seed scripts, `private.zip`, original `_a_priori` Office evidence, private evidence manifests, and generated `_a_posteriori` Office evidence must remain out of the public repository. CI must use fake fixtures only and must never require private CV data. Version `0.0.2` records the roadmap, private-repo evidence location, and integrity gate; the actual XLSX/DOCX regeneration implementation remains future work.