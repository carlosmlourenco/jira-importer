# Changelog

All notable changes to the Jira Importer script are documented here.

## v2 — current

### Added
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
