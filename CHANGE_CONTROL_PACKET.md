# Change Control Packet

Tool: `x_create_cv_x`
Version: `0.0.1`
Packet ID: `x_create_cv_x-0.0.1-2026-07-01`
Date: `2026-07-01`
Status: Updated

## Scope

Establish the first controlled version of the privacy-first CV factory CLI and its public productionization scaffolding.

## Controlled Files

- `x_create_cv_factory_x.py`
- `README.md`
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

- Added explicit tool version `0.0.1`.
- Added `--version` CLI output.
- Added pinned Ruff, Black, mypy, pytest development tooling.
- Added strict Pylance and pytest VS Code settings.
- Added GitHub Actions CI for Ruff, Black, mypy, and pytest.
- Added public-safe fake fixture data and tests for factory output, CLI typed flags, validation, and error handling.
- Added security documentation for private CV data handling.
- Added changelog and online-service strategy documentation.
- Updated the README with public fake-data workflow, local private-data workflow, and release validation commands.

## Validation

- Run `python .\x_create_cv_factory_x.py --version`.
- Run `python .\x_create_cv_factory_x.py --help`.
- Run `python -m py_compile .\x_create_cv_factory_x.py`.
- Run `python -m pytest`.
- Run `ruff check .`.
- Run `black --check .`.
- Run `mypy`.
- Run private validation locally with `python .\x_create_cv_factory_x.py validate --db-dir .\data\private\cv_factory_output --expected-zip .\data\private\private.zip --zip-prefix private/cv_app_crud` when private data is available.

## Rollback

Restore the previous README, `.gitignore`, and factory script versions, then remove the productionization scaffolding files if reverting before release.

## Operator Notes

Real CV data, private seed scripts, and `private.zip` must remain under ignored local paths. CI must use fake fixtures only and must never require private CV data.