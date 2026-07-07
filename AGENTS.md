# x_create_cv_x Agent Readme

This readme contains repository-specific guidance for agents working on `x_create_cv_x`. Keep general model-selection and productionization guidance in `x_agent_helpers_x`; keep tool-specific lessons here.

## Default Execution

- Use GPT 5.3 Codex with the normal coding agent for deterministic implementation, test, docs, commit, push, and CI-closure slices unless the work crosses a design boundary.
- Stop for a higher-capability model or explicit human decision when the work involves renderer behavior changes, schema/default policy changes, template compatibility, privacy boundaries, release strategy, packaging, or ambiguous user-facing behavior.
- For deterministic no-design slices, preserve the closure sequence: local focused check, full gates, commit, push, CI terminal success, repo memory, and clean-state verification.

## Local Validation

Use the workspace interpreter explicitly:

```powershell
c:/Users/primu/OneDrive/Desktop/ppnw_2026_07/.venv/Scripts/python.exe -m pytest -q -k cli_audit_writes_human_readable_office_report
c:/Users/primu/OneDrive/Desktop/ppnw_2026_07/.venv/Scripts/python.exe -m ruff check .
c:/Users/primu/OneDrive/Desktop/ppnw_2026_07/.venv/Scripts/python.exe -m black --check .
c:/Users/primu/OneDrive/Desktop/ppnw_2026_07/.venv/Scripts/python.exe -m mypy --strict .
c:/Users/primu/OneDrive/Desktop/ppnw_2026_07/.venv/Scripts/python.exe -m pytest -q
```

Use the absolute Git executable for unattended commits and pushes:

```powershell
C:/Program Files/Git/cmd/git.exe -c gc.auto=0 -c maintenance.auto=false status --short
```

## CI Closure

- Use the full 40-character commit SHA for GitHub commit checks.
- Derive the run id and Quality Gates job id from the latest commit check result; never reuse stale ids.
- Close a slice only when Quality Gates is `completed/success` and the checks success marker is present.
- If CI remains `in_progress`, keep polling while run/job ids are stable and no failure marker appears.
- Treat non-blocking runner deprecation annotations as maintenance follow-up unless a required gate fails.

## Privacy And Public Fixtures

- Public source must not include real CV data, raw extraction files, private seed scripts, private zips, private JSON, addresses, emails, phone numbers, or personal history details.
- Keep `data/private/`, `private.zip`, real seed scripts, private archive contents, and real private validation outputs out of Git.
- CI must validate fake CV data only.
- Public examples should demonstrate the model without exposing the real person.
- README and docs should distinguish public fake-data flow from private local-only flow.
- Validation with real private data remains local and ignored.

## Productionization Notes

- Keep fake fixtures for CI.
- Keep tests for factory commands, JSON writing, zip validation, resume record references, and fake fixture rebuilds.
- Keep `CHANGE_CONTROL_PACKET.md`, `CHANGELOG.md`, and `SECURITY.md` synchronized when release-hardening slices touch public behavior, privacy, or operational evidence.
- Split the CLI into modules only if tests and future web/API work benefit from it.
- Public fake fixtures must rebuild expected output deterministically and must not derive from private archive contents.

## Office Audit Assertion Lessons

- For scoped Office audit Markdown assertions, probe rendered report rows in the exact audit code path before adding assertions; values can differ from lower-level structure fixtures.
- For list-valued scoped Office audit metrics, assert the exact Markdown-serialized row value from the rendered audit report instead of inferring formatting from lower-level summaries.
- For deterministic metric glidepaths, derive the next row from the full ordered metric list rather than filtered keyword searches; filtered slices can skip intermediate metrics.
- For deterministic JSON contract glidepaths, close adjacent scalars/booleans one metric family at a time with minimal churn.
- When a scalar is already literal-asserted on `generated`, close the symmetric `source` contract before moving to larger payloads.
- For per-comparison sibling closures, advance comparisons in order (`1`, then `2`, then `3`) for the same metric family before switching families.
- After closing the final sibling index for a family, start the next adjacent family at the first sibling index.
- When a rendered Markdown row is exact and deterministic, use that row value as the no-assumption seed for JSON value contracts in the same family.
- Reuse one expected constant name across a metric family unless sibling values genuinely differ.
- For nested-list metric starts, declare the expected constant with an explicit strict type shape, for example `list[list[list[str]]]`, so `mypy --strict` remains deterministic.
- Keep parity assertions outside dict/list literal blocks to avoid syntax-break inserts during small deterministic patches.

## DOCX Glidepath Lessons

- Continue DOCX structure metric families in existing structure-key order without opening parallel families.
- At a new boolean/scalar family start, introduce the matching expected constant once and reuse it across sibling closures.
- For sibling-specific maps such as page margins, preserve assertion shape but use sibling-specific expected maps when rendered values differ.
- For `normalized_text.line_count`, keep the scalar expectation fixed at `1` across comparisons and advance one sibling index per slice.
- For `normalized_text.sha256` length contracts, use fixed length `64` and advance one sibling index at a time.
- For `normalized_text.sha256` equality contracts, add only the next generated/source hash equality per comparison.
- When a family already has older parity assertions, use the live assertion block and current deterministic progression state rather than assuming historical blocks are the active frontier.
- If a handoff line conflicts with the live assertion block, use the next in-file adjacent metric missing comparison-level generated/source parity; do not reopen already-closed families.
- In repeated sibling assertion blocks, keep each assertion bound to the matching section variable (`resume_2017_docx_section`, `resume_2023_docx_section`, `resume_2024_docx_section`).

## XLSX Glidepath Lessons

- After closing DOCX sibling equality families, close adjacent XLSX singleton contracts at comparison `0` explicitly before introducing new multi-sibling families.
- For XLSX singleton `normalized_text.sha256` length contracts, keep fixed hash length `64` and close generated/source length parity in one slice.
- After XLSX singleton sha256-length closure, close explicit generated/source sha256 equality in the next adjacent slice.
- For XLSX singleton structure keys at comparison `0`, close adjacent keys by explicit generated/source parity assertions before literal value-shape contracts.
- Continue singleton structure-key closure at comparison `0` in adjacent order, one explicit generated/source assertion per slice.
- For remaining unpaired structure map/scalar contracts, close one explicit generated/source equality per slice.
- After top-level XLSX structure parity is explicit, continue into nested `styles` subkeys by sorted order with one generated/source parity assertion per slice.
- Within nested XLSX `sheets[0]` parity closure, continue adjacent scalar subkeys first (`name`, then `path`) before list/map-heavy fields.

## Mirrored Release Docs

- Mirror every release-hardening slice across `CHANGELOG.md`, `RELEASE_NOTES_0.0.3.md`, and `CHANGE_CONTROL_PACKET.md` in the same commit.
- When closing a final sibling in a metric family, mirror that final-index addition in all three docs.
- Normalize ambiguous wording such as `both generated and source` to explicit comparison indices such as `comparison 1`.
- Normalize singleton bullets that omit comparison qualifiers to explicit singleton indices such as XLSX `comparison 0` generated/source.
- Keep mirrored docs machine-checkable so deterministic completeness scans can find true gaps.

## Git And Shell Lessons

- On Windows/OneDrive worktrees, `git commit` or `git push` can trigger interactive cleanup prompts; use `-c gc.auto=0 -c maintenance.auto=false` for unattended slices.
- Prefer absolute/path-pinned Git commands through slice closure to avoid shell cwd drift.
- If `git` resolves as missing in multi-command runs, switch to `C:/Program Files/Git/cmd/git.exe` immediately.
- If quoted absolute paths are rejected after terminal parser drift, use the short-path executable form (`C:/Progra~1/Git/cmd/git.exe`) until closure is complete.

## Current No-Design Frontier Pattern

The current deterministic docs-normalization lane converts old unqualified XLSX singleton value-contract bullets to explicit `comparison 0 generated and source structures` wording across the three mirrored release docs, one metric per slice, with full local gates and CI closure.
