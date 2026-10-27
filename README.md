# 🚀 Jira Importer (Google Apps Script)

A Google Sheets add-on that moves work items **both ways** between a spreadsheet and
Jira Cloud (REST API v3): push rows to Jira (create / update / sync) and pull issues
back from Jira into the sheet. Built for bulk importing and for PM reporting.

## What it does

- **Write to Jira** — reads issue data from the `Bulk Import` sheet (one issue per row),
  creates new Jira issues, updates existing ones, or both in a single pass. Writes the
  resulting Jira key back as a clickable link and tracks per-row status.
- **Read from Jira** — `Get Jira Data` pulls issues matching a JQL query into a `Jira`
  sheet, in a fixed column layout, in one click.
- Reports errors in a readable way, including a run log on a dedicated tab.

## Files

| File | Description |
|------|-------------|
| `jiraimporter v3.2.js` | Current version. Routes **Service Request** issues to the **DTS** project on create. |
| `jiraimporter v3.1.js` | Adds the Labels field and sources write modes from the `Bulk Import` sheet. |
| `jiraimporter v3.0.js` | Adds **Get Jira Data** (read from Jira); migrates JQL search to the enhanced endpoint. |
| `jiraimporter v2.js`   | Create / Update / Sync modes, richer error reporting, help menu. |
| `jiraimporter v1.1.js` | Earlier iteration. |
| `jiraimporter v1.js`   | First version. |

The `.js` files are Google Apps Script — paste them into the Apps Script editor
bound to your Google Sheet (Extensions → Apps Script).

## Setup

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste the contents of `jiraimporter v3.2.js`.
3. Reload the sheet. A **🚀 Jira Importer by Carloszai** menu appears.
4. Run **⚙️ Setup My Jira Credentials** and enter your Jira email + API token.
   - Get an API token at: https://id.atlassian.com/manage-profile/security/api-tokens
   - The token is stored privately per user via Apps Script `PropertiesService`.
     It is **not** stored in the code or the sheet.
5. Run **🧪 Test Connection** to confirm.

## Write modes (sheet → Jira)

Import / Update / Sync read rows from the **`Bulk Import`** sheet (by name — you don't
need it to be the active tab). What differs is which rows they act on, based on whether
a row already has a value in the **Jira Key** column.

| Mode | Row has a key | Row has no key |
|------|---------------|----------------|
| **Import** (Create) | skip | create |
| **Update**          | update | skip |
| **Sync**            | update | create |

Notes:
- Updates send whatever is in the sheet (not a diff) — every mapped field with a
  value is pushed on each run.
- A **blank** cell does **not** clear the field in Jira; it is simply omitted.
- Project and Issue Type are only set when creating, not when updating.
- Issue Type **Service Request** is always created in the **DTS** project, regardless
  of the project you pick at the prompt. All other types use the prompted project.
- **Selected Rows** actions require you to be on the `Bulk Import` sheet with the rows
  selected.
- Use the **❓ Help** menu item inside the sheet for the same explanation.

## Read mode (Jira → sheet)

**⬇️ Get Jira Data** pulls issues OUT of Jira for reporting/PM work:

- Reads a JQL query from the **`Assets`** sheet, cell **B3** (a bare project key like
  `RBI` also works).
- Pulls every matching issue into a sheet named **`Jira`**, creating it if needed and
  replacing the previous contents (columns A–V) each run. Helper columns you keep to the
  right of column V are left untouched.
- Writes a real datetime of the last successful run into **`Assets!B5`**.
- One click, no prompts. Dates render as `DD/MM/YYYY HH:MM:SS`; time estimates come
  through as **raw seconds** so downstream `/3600` formulas keep working.

## Column layout — `Bulk Import` sheet (0-based)

| Col | Field |
|-----|-------|
| A | Initiative |
| B | Parent ID |
| C | Summary |
| D | Description |
| E | Acceptance Criteria |
| F | CS Functional Group |
| G | # of Approved Hours |
| H | Start date |
| I | Due Date |
| J | Original Estimate (**seconds**) |
| K | Issue Type |
| L | Account |
| M | Labels (comma-separated; no spaces per label — use underscores) |

The tool adds two columns on first run: **Jira Key** and **Jira Notes**.

## Configuration

Edit these near the top of the script:

- `JIRA_BASE_URL` — your Jira site URL.
- `CUSTOM_FIELDS` — custom field IDs for your Jira instance.
- `CS_FUNCTIONAL_GROUP_MAP`, `ACCOUNT_MAP` — text-to-id value mappings.

## Author

Carlos Lourenço — carlos.lourenco@feedzai.com
