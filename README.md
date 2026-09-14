# 🚀 Jira Importer (Google Apps Script)

A Google Sheets add-on that reads rows from the active sheet and creates or updates
Jira issues through the Jira Cloud REST API (v3). Built for bulk importing and
syncing work items straight from a spreadsheet.

## What it does

- Reads issue data from the active sheet (one issue per row).
- Creates new Jira issues, updates existing ones, or both in a single pass.
- Writes the resulting Jira key back into the sheet as a clickable link to the issue,
  and tracks per-row status.
- Reports errors in a readable way, including a run log on a dedicated tab.

## Files

| File | Description |
|------|-------------|
| `jiraimporter v2.js`   | Current version. Create / Update / Sync modes, richer error reporting, help menu. |
| `jiraimporter v1.1.js` | Earlier iteration. |
| `jiraimporter v1.js`   | First version. |

The `.js` files are Google Apps Script — paste them into the Apps Script editor
bound to your Google Sheet (Extensions → Apps Script).

## Setup

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste the contents of `jiraimporter v2.js`.
3. Reload the sheet. A **🚀 Jira Importer by Carloszai** menu appears.
4. Run **⚙️ Setup My Jira Credentials** and enter your Jira email + API token.
   - Get an API token at: https://id.atlassian.com/manage-profile/security/api-tokens
   - The token is stored privately per user via Apps Script `PropertiesService`.
     It is **not** stored in the code or the sheet.
5. Run **🧪 Test Connection** to confirm.

## Modes

All modes write **from the sheet to Jira**. What differs is which rows they act on,
based on whether a row already has a value in the **Jira Key** column.

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
- Use the **❓ Help** menu item inside the sheet for the same explanation.

## Column layout (0-based)

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

The tool adds two columns on first run: **Jira Key** and **Jira Notes**.

## Configuration

Edit these near the top of the script:

- `JIRA_BASE_URL` — your Jira site URL.
- `CUSTOM_FIELDS` — custom field IDs for your Jira instance.
- `CS_FUNCTIONAL_GROUP_MAP`, `ACCOUNT_MAP` — text-to-id value mappings.

## Author

Carlos Lourenço — carlos.lourenco@feedzai.com
