# Security Policy

## Supported Versions

The current supported development version is `0.0.3`.

## Private Data Boundary

This repository is designed around a strict public/private split.

Public files may contain:

- Source code.
- Documentation.
- Tests.
- Synthetic examples and fixtures.
- Public-safe architecture and strategy notes.

Public files must not contain:

- Real CV or resume data.
- Real names, addresses, phone numbers, emails, employment history, publication identifiers tied to the private profile, or other PII.
- Private rebuild scripts.
- `data/private/private.zip`.
- Original `_a_priori` Office evidence from the private sibling repo `x_create_cv_test_data_x`.
- Private Office evidence integrity manifests from `x_create_cv_test_data_x`.
- Private chain-of-evidence scripts, generated JSON, legacy references, comparison reports, and human approval notes from `x_create_cv_test_data_x`.
- Generated `_a_posteriori` Office evidence from private evidence folders or private repos.
- Generated private JSON from `data/private/`.
- Raw extraction output or legacy source archives.
- Tokens, passwords, API keys, or local machine secrets.

## Local Private Workflow

Real CV data belongs outside the public repository. Ad hoc local output may use ignored paths such as `data/private/`, while golden evidence belongs in the private sibling repository `x_create_cv_test_data_x`. The private golden zip, original `_a_priori` Office evidence, private SHA-256 manifests, generated `_a_posteriori` Office evidence, generated JSON, comparison reports, human approval notes, and private rebuild scripts are local validation tools only. They are not part of public CI and must not be committed to the public repository.

## CI And Tests

CI uses fake fixtures only. Tests must not require real private data, network credentials, or local machine-specific files.

## Reporting Issues

If you find a private-data exposure risk, remove the sensitive file from the working tree immediately, rotate any exposed secrets if applicable, and treat the incident as a repository hygiene/security issue before continuing feature work.