# Security Policy

## Supported Versions

The current supported development version is `0.0.2`.

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
- Private seed scripts.
- `data/private/private.zip`.
- Original `_a_priori` Office evidence from `data/private/evidence/source_office/a_priori/`.
- Private Office evidence integrity manifests from `data/private/evidence/`.
- Generated `_a_posteriori` Office evidence from `data/private/evidence/generated/`.
- Generated private JSON from `data/private/`.
- Raw extraction output or legacy source archives.
- Tokens, passwords, API keys, or local machine secrets.

## Local Private Workflow

Real CV data belongs under ignored local paths such as `data/private/`. The private golden zip, original `_a_priori` Office evidence, private SHA-256 manifests, generated `_a_posteriori` Office evidence, and private seed scripts are local validation tools only. They are not part of CI and must not be committed.

## CI And Tests

CI uses fake fixtures only. Tests must not require real private data, network credentials, or local machine-specific files.

## Reporting Issues

If you find a private-data exposure risk, remove the sensitive file from the working tree immediately, rotate any exposed secrets if applicable, and treat the incident as a repository hygiene/security issue before continuing feature work.