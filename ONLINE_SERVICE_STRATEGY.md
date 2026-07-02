# Online Service Strategy

## Product Vision

`x_create_cv_x` can become a showcase-grade CV factory: a private, structured career-data service that lets a user maintain one master professional profile and generate multiple tailored resumes from it.

The online version should demonstrate three strengths:

- A clean master-data model for jobs, achievements, skills, education, publications, patents, residences, and contact methods.
- Resume-builder workflows that assemble documents from reusable master records instead of copying and pasting text.
- A privacy-first architecture where real personal data is never committed to the public repository.

## Core Positioning

This should not look like a generic resume template site. The differentiator is structured career intelligence.

The service should answer:

- What do I know about this person?
- Which facts are reusable across resumes?
- Which resume uses which facts?
- How can a user tune a resume without corrupting the master profile?

The public demo can use fake data, but the workflow should be strong enough for real private use.

## Target Architecture

### 1. Domain Core

Keep `x_create_cv_factory_x.py` as the first domain core. It already models useful commands such as:

- `add-user`
- `add-address`
- `add-job`
- `add-achievement`
- `add-patent`
- `add-publication`
- `create-resume`
- `add-resume-section`
- `add-resume-item`

Next, split this into a small Python package so the CLI, API, and future UI can all call the same functions.

Recommended structure:

```text
x_create_cv_x/
  x_create_cv_factory_x.py
  x_create_cv_x/
    __init__.py
    models.py
    store.py
    commands.py
    validation.py
    render.py
  tests/
```

The CLI should become a thin wrapper around `commands.py`.

### 2. Storage

Start with JSON files for local-first simplicity:

```text
data/private/
  master_profile.json
  resumes/
    resume_2017.json
    resume_2023.json
```

For the hosted version, move to one of these:

- SQLite for a single-user/local hosted service.
- Postgres for a multi-user SaaS version.
- Encrypted object storage for export/import snapshots.

Do not store real user PII in the public repo. Public examples should use fake data only.

### 3. API Layer

Use FastAPI for the first web service layer. It fits the Python project and gives automatic OpenAPI docs.

Suggested endpoints:

```text
POST   /profiles
GET    /profiles/{profile_id}
POST   /profiles/{profile_id}/addresses
POST   /profiles/{profile_id}/jobs
POST   /profiles/{profile_id}/jobs/{job_id}/achievements
POST   /profiles/{profile_id}/skills
POST   /profiles/{profile_id}/patents
POST   /profiles/{profile_id}/publications

POST   /profiles/{profile_id}/resumes
POST   /profiles/{profile_id}/resumes/{resume_id}/sections
POST   /profiles/{profile_id}/resumes/{resume_id}/items
GET    /profiles/{profile_id}/resumes/{resume_id}/preview
POST   /profiles/{profile_id}/resumes/{resume_id}/export
```

Each endpoint should map to a granular domain command. That preserves the “button-like” design and makes later GUI work straightforward.

### 4. Frontend

The frontend should be an app, not a landing page.

Primary screens:

- Master Profile: editable collections for person, contact methods, jobs, achievements, skills, education, patents, publications, and residences.
- Resume Builder: sections on the left, selected items in the middle, master-data picker on the right.
- Preview: rendered resume view with export controls.
- Compare: show which master records are used by each resume.

Good first stack:

- React or Next.js for the UI.
- FastAPI backend.
- JSON or SQLite storage for the prototype.
- Postgres only when authentication and multi-user hosting are needed.

## Privacy Strategy

Privacy is part of the product, not a deployment footnote.

Rules:

- Keep `data/private/` ignored forever.
- Keep real seed scripts ignored forever.
- Commit only fake example profiles.
- Add a `data/examples/` folder with synthetic data for demos and tests.
- Add an export/import feature so users can own their data.
- Avoid logging request bodies that contain profile data.

Hosted version options:

- Local-first: user runs the service locally and no PII leaves the machine.
- Private hosted: user deploys their own instance.
- SaaS: requires authentication, encrypted backups, deletion workflows, audit logs, and a privacy policy.

For a showcase project, start with local-first plus a fake-data hosted demo.

## Validation Strategy

The current private zip workflow is valuable as a temporary golden-master test. Keep it private. The Office-regeneration workflow should first run `exercise-golden` against the private `_a_priori` and full chain-of-evidence SHA-256 manifests in `x_create_cv_test_data_x`, then generate `_a_posteriori` outputs from scripts and JSON before comparing them against the original evidence.

Public validation should use fake data:

```text
tests/fixtures/fake_profile_seed.py
tests/fixtures/expected_master_profile.json
tests/fixtures/expected_resume_backend.json
tests/fixtures/expected_resume_public.json
```

Test layers:

- Unit tests for each command.
- Golden JSON tests for deterministic output.
- API tests for FastAPI endpoints.
- Render tests for Markdown/HTML resume output.
- Security tests confirming private paths are ignored and examples contain no real PII.

## Deployment Path

### Phase 1: Package The Core

- Move reusable logic out of the CLI into importable modules.
- Add `pyproject.toml`.
- Add Ruff, Black, mypy, pytest, and GitHub Actions.
- Add fake example data.
- Keep private seed scripts local only.

### Phase 2: Local Web App

- Add FastAPI.
- Add endpoints for master-data CRUD.
- Add endpoints for resume CRUD.
- Add a simple frontend or server-rendered UI.
- Store data in local JSON files or SQLite.

### Phase 3: Resume Rendering

- Render Markdown from resume JSON.
- Add HTML preview.
- Add PDF export later, likely through Playwright or WeasyPrint.
- Add templates, but keep the data model template-independent.

### Phase 4: Hosted Demo

- Deploy a fake-data demo only.
- Good hosting options:
  - Render or Fly.io for FastAPI.
  - Vercel for a separate frontend.
  - Railway/Supabase if Postgres becomes necessary.
- Use synthetic data and clear demo labeling.

### Phase 5: Private User Mode

- Add auth.
- Add encrypted backups.
- Add account deletion.
- Add import/export.
- Add user-owned deployment docs.

## Showcase Features

For a portfolio-quality project, prioritize features that demonstrate engineering judgment:

- Master data reused across multiple resumes.
- Resume-specific overrides without mutating master records.
- Diff view between resumes.
- “Unused facts” view showing master records not used in a resume.
- “Where used?” view for any job, skill, publication, or achievement.
- Deterministic JSON output.
- Fake demo profile that can be rebuilt from readable seed commands.
- Validation command that proves generated data matches expected JSON.

## Public Repo Hygiene

Before publishing as a showcase:

- Remove raw extraction code if it is no longer needed.
- Keep `private.zip` out of Git.
- Keep private `_a_priori` and `_a_posteriori` Office evidence out of Git.
- Keep `data/private/` out of Git.
- Add fake fixtures and examples.
- Add a `SECURITY.md` explaining private-data handling.
- Add a `docs/architecture.md` after the package/API split.
- Add CI checks.

## Recommended Next Step

The next concrete step is to professionalize the current CLI project while preserving the private data workflow:

1. Create a proper Python package.
2. Move command logic into reusable modules.
3. Add fake example seed scripts.
4. Add tests proving the fake seed scripts rebuild expected JSON.
5. Add FastAPI only after the command layer is clean and tested.

That path keeps the project grounded while moving it toward a credible online service.