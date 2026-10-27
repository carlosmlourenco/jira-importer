# Changelog

All notable changes to the Jira Importer script are documented here.

## v3.3 — current

### Added
- **DTS Service Request required fields.** Creating a Service Request (routed to DTS
  since v3.2) now also sends the three fields the DTS project requires:
  - **DT Service Catalogue** (`customfield_25906`) — chosen **per row** in a new
    **column N**. The cell takes a catalogue label (matched case-insensitively) or a
    bare option id; mapped to `{ id }` via `DT_SERVICE_CATALOGUE_MAP` (17 options).
  - **Cluster Team** (`customfield_11073`) and **Roadmap** (`customfield_11477`) —
    fixed defaults (`18627` / `19215`) applied to every DTS Service Request.
  All three are sent **only** when routing to DTS, create-only, and are protected from
  the field-drop fallback (dropping a required field guarantees failure).
- **`showFieldOptions` diagnostic** — lists the allowed options (label + id) of an
  option/select field from createmeta `allowedValues`. Used to capture the DT Service
  Catalogue option list. Editor-only (not on the menu), like the other diagnostics.

### Changed
- `Bulk Import` column layout extended: **N = DT Service Catalogue**; Jira Key / Jira
  Notes move to **O / P** (still found by header name, so no positional change needed).

## v3.2

### Added
- **Project routing by issue type (create only).** Issues whose Issue Type is
  **Service Request** are always created in the **DTS** project, regardless of the
  project chosen at the prompt. Every other issue type uses the prompted project as
  before. The parent link is still attempted for DTS Service Requests; if Jira rejects
  a cross-project parent, the existing field-drop fallback handles it. Matched
  case-insensitively; configurable via `SERVICE_REQUEST_ISSUE_TYPE` /
  `SERVICE_REQUEST_PROJECT`.

## v3.1

### Added
- **Labels field (column M).** New `labels` mapping on the import side. The cell is
  comma-separated (e.g. `dt_cs, dt_bluefcu`); the script splits, trims, drops empties,
  and sends Jira's standard `labels` field. Applied on **both create and update**, and
  left in the field-drop fallback so a rejected label doesn't fail the whole row.
  Labels cannot contain spaces in Jira — type them with underscores (e.g.
  `DT_Activation_Team`).

### Changed
- **Write modes read from the "Bulk Import" sheet by name.** Import / Update / Sync no
  longer act on whatever tab is active — they always read rows from a sheet named
  `Bulk Import`, with a clear error if it's missing. "Selected Rows" actions still
  require you to be on the `Bulk Import` sheet with the rows selected.
- Help text updated to describe the `Bulk Import` source sheet and the Labels column.

## v3.0

### Added
- **Get Jira Data (read from Jira).** New one-click export that pulls issues OUT of
  Jira into a sheet named `Jira`, replacing the "Jira Cloud for Sheets" extension.
  Reads a JQL query (or a bare project key) from `Assets!B3`, pages through the
  enhanced `/rest/api/3/search/jql` endpoint, and writes a fixed 22-column layout.
  Formatting matches the extension: dates as `DD/MM/YYYY HH:MM:SS`, time estimates as
  raw seconds (so downstream `/3600` formulas keep working). Clears only columns A–V,
  leaving user helper columns to the right intact.
- **Last-run timestamp** written to `Assets!B5` (a real datetime value) on each
  successful export.

### Changed
- Migrated JQL search from the removed legacy `/rest/api/3/search` endpoint
  (Atlassian HTTP 410 / CHANGE-2046) to the enhanced `/rest/api/3/search/jql`
  endpoint with cursor-based pagination (`nextPageToken`).
- Help dialog reworked to cover both directions (write to Jira + read from Jira).

## v2

### Fixed
- **Account field accepts a JSON object in the cell.** Previously only plain-text
  labels mapped via `ACCOUNT_MAP` worked; a cell containing
  `{"id": 4, "value": "..."}` was treated as unknown and skipped. New
  `parseAccountValue` helper accepts a JSON object, a mapped label, or a bare id,
  and coerces the id to a string (Jira requires `id` as a string).

- **Never drop required fields in the fallback.** Previously the field-drop retry
  could remove a field that Jira reported as required (e.g. Epic Name), which
  guaranteed failure. `parseJiraError` now flags fields whose message says they are
  required, and both the create and update fallbacks exclude those from the drop set.

### Added
- **Epic Name on epic create** — when Issue Type is `Epic`, the Epic Name field
  (`customfield_10371` for this instance) is set to the summary and protected from
  the field-drop fallback. This field is required by a validator even though it is
  not reported by `createmeta`.
- **Diagnostics**: `🔎 Show Required Fields for Issue Type…` (lists required fields +
  schema types from createmeta) and `🔬 Inspect Existing Issue Fields…` (reads a live
  issue and shows field ids/values by name match — used to find the real Epic Name id).
- **Clickable Jira links** — the "Jira Key" cell is now written as a rich-text
  hyperlink to the issue (`{base}/browse/{KEY}`) on create and update, so the key
  can be clicked to open the issue. Falls back to plain text if rich text is
  unavailable. Reading the cell still returns the key, so skip/update detection is
  unaffected.

- **Update mode** — pushes sheet values onto existing Jira issues (rows that already
  have a Jira Key). Uses `PUT /rest/api/3/issue/{key}` and treats `204 No Content` as success.
- **Sync mode** — single pass that creates rows without a key and updates rows with one.
- **Help menu item** (`❓ Help — What do these modes do?`) explaining Create / Update / Sync
  and the key behavioral notes.
- **"Jira Notes" column** — per-row short status plus a cell comment with full detail
  (browse URL on success, full error on failure).
- **"Import Log" tab** — timestamped run log with one row per issue and a totals footer.
- Structured, human-readable error reporting: field IDs mapped to names, error categories
  (auth / permission / field / network), and tailored guidance per category.
- Generic field-fallback: any field Jira rejects is dropped and retried (previously only
  two hardcoded fields), with the dropped fields reported back.

### Changed
- **Original Estimate** is now sent via Jira's `timetracking.originalEstimate` duration
  string (e.g. `"2400m"`) instead of the bare `timeoriginalestimate` number, which the
  create/edit endpoints reject. The column holds seconds; the script converts to minutes.
- On update, `project`, `issuetype`, and `parent` are omitted (project can't change,
  type/parent moves need a dedicated flow).
- Update failures preserve the existing Jira Key in the cell instead of overwriting it
  with `ERROR`.

## v1.1
- Intermediate iteration.

## v1
- Initial version: create-only import from the sheet to Jira, credential setup,
  custom menu, per-row Jira Key status.
