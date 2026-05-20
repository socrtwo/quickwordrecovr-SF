# CLAUDE.md

Recovers `.docx` files that Word refuses to open with "unspecified error" —
typically a truncated or malformed `word/document.xml` inside the ZIP.
Two implementations: a **cross-platform PWA under `web/`** (canonical
user-facing app) and a **legacy Windows-only VB.NET WinForms app** under
`Unspecified Error DOCX Recovery/`. Assume edits target `web/` unless told
otherwise.

## Repo map

- `web/` — the PWA. Repair runs entirely in the browser; files never
  upload. Powers the live app and every cross-platform release.
- `Unspecified Error DOCX Recovery/`,
  `Unspecified Error DOCX Recovery.sln` — legacy VB.NET WinForms app
  (Visual Studio solution). Windows-only.
- `setupVII-without-adware-offers.iss` — Inno Setup script for the
  Windows installer.
- `note on the inno installer script.txt` — read this before touching the
  Inno script.
- `releases/` — pre-packaged release archives committed to the repo.
- `.github/workflows/` — `build.yml` (CI), `pages.yml` (deploy `web/` to
  Pages on push to `main`), `release.yml` (build per-platform bundles on
  `v*` tag).

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. CI runs on the PR; Pages and Release pipelines fire from `main` only.

## Releasing

- Push a `v*` tag to `main` to produce: Windows `.exe` + Inno installer,
  plus PWA bundles for macOS / Linux / ChromeOS / Android / iOS / Web.
  All non-Windows bundles wrap the same `web/` source.

## Verifying changes

- PWA: serve `web/` locally, drop a known-broken `.docx`, confirm the
  four-step pipeline runs (open as ZIP → read `word/document.xml` →
  truncate at last well-formed paragraph → repack).
- VB.NET app: build via `Unspecified Error DOCX Recovery.sln` in Visual
  Studio. CI on `build.yml` validates this build.
- Test both happy-path (slightly broken `.docx`) and fallback
  (plain-text extraction when XML is unsalvageable).

## Gotchas

- The recovery pipeline must keep working when the central directory is
  missing — that's a common failure mode for `.docx` corruption.
- 100% client-side is a hard requirement for the PWA. No file uploads.
- The Inno installer script has known quirks documented in
  `note on the inno installer script.txt`. Read it before editing.
- The plain-text fallback is the last-resort path; don't "improve" it to
  also try XML repair — that defeats its purpose as a guaranteed-exit
  recovery mode.
