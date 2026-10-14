# Changelog

All notable changes to the Jira Importer script are documented here.

## v2 — current

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
