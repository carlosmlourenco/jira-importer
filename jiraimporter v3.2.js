// =============================================
// 🚀 Jira Importer by Carloszai — v3.2
// Setup, modes, and version history: see README.md and CHANGELOG.md
// =============================================

// =============================================
// CONFIGURATION — shared settings (not sensitive)
// =============================================
const JIRA_BASE_URL = "https://fdz.atlassian.net";

// =============================================
// CUSTOM FIELD IDs
// =============================================
const CUSTOM_FIELDS = {
  ORIGINAL_ESTIMATE:   "timeoriginalestimate",
  INITIATIVE:          "customfield_16455",
  ACCEPTANCE_CRITERIA: "customfield_11371",
  CS_FUNCTIONAL_GROUP: "customfield_14391",
  ACCOUNT:             "customfield_16433",
  APPROVED_HOURS:      "customfield_13750",
  START_DATE:          "customfield_16414",
  // "Epic Name" — required when creating Epics in this instance. Confirmed id for
  // the fdz.atlassian.net Jira via the Inspect Issue Fields diagnostic (TWIN-464).
  // Note: createmeta does NOT list this field, but a validator still requires it.
  EPIC_NAME:           "customfield_10371",
};

// =============================================
// FIELD ID → HUMAN NAME (for friendly error messages)
// =============================================
const FIELD_LABELS = {
  "summary":             "Summary",
  "description":         "Description",
  "issuetype":           "Issue Type",
  "project":             "Project",
  "parent":              "Parent",
  "labels":              "Labels",
  "duedate":             "Due Date",
  "timeoriginalestimate":"Original Estimate",
  "timetracking":        "Original Estimate (time tracking)",
  "customfield_10371":   "Epic Name",
  "customfield_16455":   "Initiative",
  "customfield_11371":   "Acceptance Criteria",
  "customfield_14391":   "CS Functional Group",
  "customfield_16433":   "Account",
  "customfield_13750":   "# of Approved Hours",
  "customfield_16414":   "Start Date",
};

// Turn a field id into a readable label, e.g. "customfield_16414" → "Start Date (customfield_16414)"
function fieldLabel(fieldId) {
  const name = FIELD_LABELS[fieldId];
  return name ? `${name} (${fieldId})` : fieldId;
}


// =============================================
// SOURCE SHEET for write modes (Import / Update / Sync)
// =============================================
// Import/Update/Sync always read rows FROM this sheet, looked up by name —
// not from whatever tab happens to be active.
const BULK_IMPORT_SHEET = "Bulk Import";

// =============================================
// PROJECT ROUTING BY ISSUE TYPE (create only)
// =============================================
// "Service Request" issues must ALWAYS be created in this project, regardless of
// the project chosen at the prompt. Every other issue type uses the prompted
// project (askProjectKey) as before.
const SERVICE_REQUEST_ISSUE_TYPE = "service request";   // matched case-insensitively
const SERVICE_REQUEST_PROJECT    = "DTS";               // forced project key for Service Requests

// =============================================
// COLUMN INDEX MAP (0-based) — the "Bulk Import" sheet
// A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12  (N/O = Jira Key/Notes)
// =============================================
const COL = {
  INITIATIVE:          0,  // A — Initiative
  PARENT:              1,  // B — Parent ID
  SUMMARY:             2,  // C — Summary
  DESCRIPTION:         3,  // D — Description
  ACCEPTANCE_CRITERIA: 4,  // E — Acceptance Criteria
  CS_FUNCTIONAL_GROUP: 5,  // F — CS Functional Group
  APPROVED_HOURS:      6,  // G — # of Approved Hours
  START_DATE:          7,  // H — Start date
  DUE_DATE:            8,  // I — Due Date
  ORIGINAL_ESTIMATE:   9,  // J — Original Estimate (seconds)
  ISSUE_TYPE:          10, // K — Issue Type
  ACCOUNT:             11, // L — Account
  LABELS:              12, // M — Labels (comma-separated; no spaces per label)
};

// =============================================
// FIELD VALUE MAPPINGS
// =============================================
const CS_FUNCTIONAL_GROUP_MAP = {
  "technical solutions": "15533",
  "data science":        "15534",
  "project management":  "15537",
  // Add more as needed:
  // "engagement manager": "XXXXX",
  // "risk consultant":    "XXXXX",
  // "csm":                "XXXXX",
  // "dt activation team": "XXXXX",
};

const ACCOUNT_MAP = {
  "standard billable":         { id: "6",  value: "Standard billable" },
  "post golive billable":      { id: "4",  value: "Post GoLive billable" },
  "l2 billable (maintenance)": { id: "20", value: "L2 billable (Maintenance)" },
  // Add more as discovered
};

// =============================================
// FOOTER — appended to all dialogs
// =============================================
const FOOTER = "\n\n─────────────────────────\n🚀 Jira Importer by Carloszai · v3.2\n📧 carlos.lourenco@feedzai.com";

// =============================================
// CREDENTIALS — stored privately per user
// =============================================
function getCredentials() {
  const props = PropertiesService.getUserProperties();
  return {
    email: props.getProperty("JIRA_EMAIL"),
    token: props.getProperty("JIRA_API_TOKEN"),
  };
}

function saveCredentials(email, token) {
  const props = PropertiesService.getUserProperties();
  props.setProperty("JIRA_EMAIL", email);
  props.setProperty("JIRA_API_TOKEN", token);
}

function clearCredentials() {
  PropertiesService.getUserProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert(
    "✅ Credentials cleared",
    "Your Jira credentials have been removed." + FOOTER,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupCredentials() {
  const ui = SpreadsheetApp.getUi();

  // Step 1 — Email
  const emailResponse = ui.prompt(
    "Jira Setup (1/2)",
    "Enter your Jira email address:\n(e.g. yourname@feedzai.com)" + FOOTER,
    ui.ButtonSet.OK_CANCEL
  );
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
  const email = emailResponse.getResponseText().trim();
  if (!email) {
    ui.alert("❌ Setup cancelled", "No email entered." + FOOTER, ui.ButtonSet.OK);
    return;
  }

  // Step 2 — API Token
  const tokenResponse = ui.prompt(
    "Jira Setup (2/2)",
    "Enter your Jira API token:\n\n" +
    "⚠️ Stored privately for your account only. No one else can see it.\n\n" +
    "👉 Get your token at:\nhttps://id.atlassian.com/manage-profile/security/api-tokens\n" +
    "   → Click 'Create API token'" +
    FOOTER,
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenResponse.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenResponse.getResponseText().trim();
  if (!token) {
    ui.alert("❌ Setup cancelled", "No token entered." + FOOTER, ui.ButtonSet.OK);
    return;
  }

  saveCredentials(email, token);

  const ok = testConnectionSilent(email, token);
  if (ok.success) {
    ui.alert(
      "✅ Setup Complete",
      `Connected as: ${ok.displayName} (${ok.email})\n\nYou're ready to import to Jira.` + FOOTER,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      "⚠️ Credentials saved but connection failed",
      `Error: ${ok.error}\n\nDouble-check your email and token and run Setup again.` + FOOTER,
      ui.ButtonSet.OK
    );
  }
}

// =============================================
// CREDENTIAL GUARD
// =============================================
function requireCredentials() {
  const { email, token } = getCredentials();
  if (!email || !token) {
    const ui       = SpreadsheetApp.getUi();
    const response = ui.alert(
      "⚙️ Credentials Required",
      "You haven't set up your Jira credentials yet.\n\nClick OK to set them up now." + FOOTER,
      ui.ButtonSet.OK_CANCEL
    );
    if (response === ui.Button.OK) setupCredentials();
    return null;
  }
  return { email, token };
}

// =============================================
// PROJECT KEY PROMPT
// =============================================
function askProjectKey() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();

  const lastKey    = props.getProperty("JIRA_LAST_PROJECT_KEY") || "";
  const suggestion = lastKey ? `Last used: ${lastKey}\n\n` : "";

  const response = ui.prompt(
    "Select Jira Project",
    `${suggestion}Enter the Jira project key to import into:\n(e.g. BACKBASE, LLOYDS, ARGENTA)` + FOOTER,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const projectKey = response.getResponseText().trim().toUpperCase();
  if (!projectKey) {
    ui.alert("❌ Cancelled", "No project key entered." + FOOTER, ui.ButtonSet.OK);
    return null;
  }

  props.setProperty("JIRA_LAST_PROJECT_KEY", projectKey);
  return projectKey;
}

// =============================================
// MENU SETUP
// =============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Jira Importer by Carloszai")
    .addItem("❓ Help — How this tool works", "showModeHelp")
    .addSeparator()
    .addItem("⬇️ Get Jira Data",             "exportFromJira")
    .addSeparator()
    .addItem("⬆️ Import ALL Rows to Jira",      "importAllToJira")
    .addItem("⬆️ Import Selected Rows to Jira", "importSelectedToJira")
    .addSeparator()
    .addItem("🔄 Sync ALL Rows (create + update)",      "syncAllToJira")
    .addItem("🔄 Sync Selected Rows (create + update)", "syncSelectedToJira")
    .addSeparator()
    .addItem("✏️ Update ALL Rows (with Jira Key)",      "updateAllInJira")
    .addItem("✏️ Update Selected Rows (with Jira Key)", "updateSelectedInJira")
    .addSeparator()
    .addItem("🧪 Test Connection",           "testConnection")
    // Diagnostics kept in code but off the menu to reduce clutter. To use them,
    // run showRequiredFields / inspectIssueFields from the Apps Script editor
    // (Run → select function), or temporarily re-add the lines below:
    // .addItem("🔎 Show Required Fields for Issue Type…", "showRequiredFields")
    // .addItem("🔬 Inspect Existing Issue Fields…",       "inspectIssueFields")
    .addSeparator()
    .addItem("⚙️ Setup My Jira Credentials", "setupCredentials")
    .addItem("🗑️ Clear My Credentials",      "clearCredentials")
    .addToUi();
}

// =============================================
// DIAGNOSTIC — list the required fields (with real IDs) for an issue type
// =============================================
// Uses GET /rest/api/3/issue/createmeta so you can see the exact field IDs your
// Jira instance expects (e.g. the real "Epic Name" custom field id).
// Fetch an existing issue and show the real IDs + current values of any field
// whose NAME contains the search term (default "epic"). This finds fields that
// createmeta hides but that validators still require.
function inspectIssueFields() {
  const creds = requireCredentials();
  if (!creds) return;
  const ui = SpreadsheetApp.getUi();

  const keyResp = ui.prompt(
    "Inspect an existing issue",
    "Enter an existing issue key to read (ideally an existing Epic):\n(e.g. TWIN-123)" + FOOTER,
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResp.getSelectedButton() !== ui.Button.OK) return;
  const issueKey = keyResp.getResponseText().trim();
  if (!issueKey) return;

  const termResp = ui.prompt(
    "Filter fields",
    "Show fields whose NAME contains (leave blank for 'epic'):" + FOOTER,
    ui.ButtonSet.OK_CANCEL
  );
  if (termResp.getSelectedButton() !== ui.Button.OK) return;
  const term = (termResp.getResponseText().trim() || "epic").toLowerCase();

  const credentials = Utilities.base64Encode(`${creds.email}:${creds.token}`);
  // expand=names gives us a fieldId → human name map alongside the values.
  const url = `${JIRA_BASE_URL}/rest/api/3/issue/${encodeURIComponent(issueKey)}?expand=names`;

  let body;
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": `Basic ${credentials}`, "Accept": "application/json" },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200) {
      ui.alert("❌ Could not read issue",
        `HTTP ${resp.getResponseCode()}\n${resp.getContentText().slice(0, 400)}` + FOOTER,
        ui.ButtonSet.OK);
      return;
    }
    body = JSON.parse(resp.getContentText());
  } catch (e) {
    ui.alert("❌ Error", e.message + FOOTER, ui.ButtonSet.OK);
    return;
  }

  const names  = body.names || {};
  const fields = body.fields || {};
  const matches = [];
  for (const fieldId in names) {
    const name = String(names[fieldId] || "");
    if (name.toLowerCase().indexOf(term) !== -1) {
      let val = fields[fieldId];
      val = (val === null || val === undefined) ? "(empty)" : JSON.stringify(val);
      if (val.length > 80) val = val.slice(0, 80) + "…";
      matches.push(`${name}  →  ${fieldId}\n     value: ${val}`);
    }
  }

  let msg = `Issue: ${issueKey}\nFields whose name contains "${term}":\n\n`;
  msg += matches.length ? matches.join("\n") : "(no matching fields)";
  msg += `\n\nThe field that is NOT empty on an Epic is your Epic Name field.\nSet CUSTOM_FIELDS to that id.`;

  ui.alert("🔬 Issue field inspector", msg + FOOTER, ui.ButtonSet.OK);
}

function showRequiredFields() {
  const creds = requireCredentials();
  if (!creds) return;

  const ui = SpreadsheetApp.getUi();

  const projectKey = askProjectKey();
  if (!projectKey) return;

  const typeResp = ui.prompt(
    "Required Fields",
    "Enter the issue type to inspect:\n(e.g. Epic, Story, Task)" + FOOTER,
    ui.ButtonSet.OK_CANCEL
  );
  if (typeResp.getSelectedButton() !== ui.Button.OK) return;
  const issueType = typeResp.getResponseText().trim();
  if (!issueType) return;

  const credentials = Utilities.base64Encode(`${creds.email}:${creds.token}`);
  const url =
    `${JIRA_BASE_URL}/rest/api/3/issue/createmeta` +
    `?projectKeys=${encodeURIComponent(projectKey)}` +
    `&issuetypeNames=${encodeURIComponent(issueType)}` +
    `&expand=projects.issuetypes.fields`;

  let body;
  try {
    const response = UrlFetchApp.fetch(url, {
      method:             "get",
      headers:            { "Authorization": `Basic ${credentials}`, "Accept": "application/json" },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      ui.alert("❌ Could not fetch metadata",
        `HTTP ${response.getResponseCode()}\n${response.getContentText().slice(0, 400)}` + FOOTER,
        ui.ButtonSet.OK);
      return;
    }
    body = JSON.parse(response.getContentText());
  } catch (e) {
    ui.alert("❌ Error", e.message + FOOTER, ui.ButtonSet.OK);
    return;
  }

  // Drill into projects[0].issuetypes[0].fields
  const project = body.projects && body.projects[0];
  const itype   = project && project.issuetypes && project.issuetypes[0];
  if (!itype || !itype.fields) {
    ui.alert("No data",
      `No create metadata found for "${issueType}" in "${projectKey}".\n` +
      `Check the type name and your permissions.` + FOOTER,
      ui.ButtonSet.OK);
    return;
  }

  const required = [];
  let optionalCount = 0;
  for (const fieldId in itype.fields) {
    const f = itype.fields[fieldId];
    if (!f.required) { optionalCount++; continue; }
    // Describe the expected shape so we know how to build the payload.
    const schema = f.schema || {};
    const type   = schema.type || "?";
    const items  = schema.items ? `<${schema.items}>` : "";
    required.push(`${f.name}  →  ${fieldId}\n     type: ${type}${items}`);
  }

  let msg = `Issue type: ${issueType}   Project: ${projectKey}\n\n`;
  msg += `REQUIRED fields (${required.length}):\n\n` + (required.length ? required.join("\n") : "  (none)");
  msg += `\n\nOptional fields: ${optionalCount} (not listed)`;

  ui.alert("🔎 Required fields", msg + FOOTER, ui.ButtonSet.OK);
}

// =============================================
// HELP — explain what each mode does
// =============================================
function showModeHelp() {
  const msg =
    "This tool works in BOTH directions:\n" +
    "   • WRITE to Jira  → Import / Update / Sync (push rows from the sheet).\n" +
    "   • READ from Jira → Get Jira Data (pull issues into a \"Jira\" tab).\n" +
    "\n" +
    "════════ WRITE TO JIRA ════════\n" +
    "These modes always read rows FROM the \"Bulk Import\" sheet (by name) — you\n" +
    "don't need to have it open. \"Selected Rows\" actions do require you to be on\n" +
    "the \"Bulk Import\" sheet with the rows selected.\n" +
    "\n" +
    "What differs between the three is which rows they touch, based on whether a\n" +
    "row already has a value in the \"Jira Key\" column.\n" +
    "\n" +
    "📥 IMPORT (Create)\n" +
    "   • Row WITHOUT a key → creates a new Jira issue, writes the new key back.\n" +
    "   • Row WITH a key    → skipped (already imported).\n" +
    "   Use when: adding brand-new issues.\n" +
    "\n" +
    "✏️ UPDATE\n" +
    "   • Row WITH a key    → pushes the sheet's values onto that existing issue.\n" +
    "   • Row WITHOUT a key → skipped (never creates anything).\n" +
    "   Use when: editing existing issues only.\n" +
    "\n" +
    "🔄 SYNC (Import + Update in one pass)\n" +
    "   • Row WITH a key    → updates it.\n" +
    "   • Row WITHOUT a key → creates it.\n" +
    "   Use when: your sheet mixes new and existing issues.\n" +
    "\n" +
    "Good to know (write modes):\n" +
    "   • Updates send whatever is in the sheet, not a diff — every mapped field\n" +
    "     with a value is pushed on each run.\n" +
    "   • A BLANK cell does NOT clear the field in Jira; it is simply left as-is.\n" +
    "   • Project & Issue Type are only set when creating, not when updating.\n" +
    "   • Issue Type \"Service Request\" is ALWAYS created in the DTS project,\n" +
    "     no matter which project you pick at the prompt. All other types use\n" +
    "     the project you enter.\n" +
    "   • Labels (column M): comma-separated, e.g. dt_cs, dt_bluefcu. Labels\n" +
    "     can't contain spaces — use underscores (e.g. DT_Activation_Team).\n" +
    "\n" +
    "════════ READ FROM JIRA ════════\n" +
    "⬇️ GET JIRA DATA\n" +
    "   • Reads a JQL query from the \"Assets\" sheet, cell B3\n" +
    "     (a bare project key like \"RBI\" also works).\n" +
    "   • Pulls every matching issue into a sheet called \"Jira\" — creating it if\n" +
    "     needed, and REPLACING the previous contents (columns A–V) each run.\n" +
    "   • Any helper columns you add to the RIGHT of column V are left untouched.\n" +
    "   • On success, stamps the run date/time into Assets!B5.\n" +
    "   • Runs in one click — no prompts. If B3 is empty or the \"Assets\" sheet is\n" +
    "     missing, it tells you exactly what to fix.";

  SpreadsheetApp.getUi().alert(
    "❓ How this tool works",
    msg + FOOTER,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================
// ENTRY POINTS — mode is one of: "create" | "update" | "sync"
// =============================================
function importAllToJira()      { runImport("create", false); }
function importSelectedToJira()  { runImport("create", true); }
function syncAllToJira()         { runImport("sync",   false); }
function syncSelectedToJira()    { runImport("sync",   true); }
function updateAllInJira()       { runImport("update", false); }
function updateSelectedInJira()  { runImport("update", true); }

// Shared launcher for every menu action.
function runImport(mode, selectionOnly) {
  const creds = requireCredentials();
  if (!creds) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Always read from the "Bulk Import" sheet, regardless of which tab is active.
  const sheet = ss.getSheetByName(BULK_IMPORT_SHEET);
  if (!sheet) {
    ui.alert(
      "❌ Missing \"" + BULK_IMPORT_SHEET + "\" sheet",
      `Import, Sync and Update read rows from a sheet named "${BULK_IMPORT_SHEET}", ` +
      `but it wasn't found.\n\nCreate/rename your import sheet to "${BULK_IMPORT_SHEET}" ` +
      `and try again.` + FOOTER,
      ui.ButtonSet.OK
    );
    return;
  }

  const projectKey = askProjectKey();
  if (!projectKey) return;

  const allData = sheet.getDataRange().getValues();

  let from, to;
  if (selectionOnly) {
    // Selection only makes sense when the user is actually on the Bulk Import
    // sheet. If they ran a "Selected Rows" action from another tab, fall back
    // to all rows (there is no meaningful selection on this sheet otherwise).
    const activeSheet = ss.getActiveSheet();
    if (activeSheet.getName() === BULK_IMPORT_SHEET) {
      const selection = sheet.getActiveRange();
      from = Math.max(selection.getRow() - 1, 1);
      to   = selection.getLastRow() - 1;
    } else {
      ui.alert(
        "ℹ️ No selection on \"" + BULK_IMPORT_SHEET + "\"",
        `You ran a "Selected Rows" action, but the active tab isn't ` +
        `"${BULK_IMPORT_SHEET}".\n\nOpen the "${BULK_IMPORT_SHEET}" sheet and select ` +
        `the rows you want, then run it again.` + FOOTER,
        ui.ButtonSet.OK
      );
      return;
    }
  } else {
    from = 1;
    to   = allData.length - 1;
  }

  processRows(sheet, allData, from, to, creds, projectKey, mode);
}

// =============================================
// CORE PROCESSING LOGIC
// =============================================
function processRows(sheet, data, fromIndex, toIndex, creds, projectKey, mode) {
  mode = mode || "create";
  const ui    = SpreadsheetApp.getUi();
  const count = toIndex - fromIndex + 1;

  const modeLabel = {
    create: "Create new issues",
    update: "Update existing issues (rows with a Jira Key)",
    sync:   "Sync (create rows without a key, update rows with one)",
  }[mode];

  const confirm = ui.alert(
    "Jira " + (mode === "update" ? "Update" : mode === "sync" ? "Sync" : "Import"),
    `Mode: ${modeLabel}\n\n` +
    `About to process ${count} row(s) against project "${projectKey}".\n\nProceed?` + FOOTER,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Add "Jira Key" status column if it doesn't exist
  const headers    = data[0];
  let statusColIdx = headers.indexOf("Jira Key");
  if (statusColIdx === -1) {
    statusColIdx = headers.length;
    sheet.getRange(1, statusColIdx + 1).setValue("Jira Key");
  }

  // Add "Jira Notes" column right after "Jira Key" if it doesn't exist
  let notesColIdx = headers.indexOf("Jira Notes");
  if (notesColIdx === -1) {
    notesColIdx = statusColIdx + 1;
    sheet.getRange(1, notesColIdx + 1).setValue("Jira Notes");
  }

  // Start a fresh run log
  const runStarted = new Date();
  const log        = startImportLog(projectKey, runStarted, mode);

  const results = { created: 0, updated: 0, warnings: 0, skipped: 0, errors: [] };

  for (let i = fromIndex; i <= toIndex; i++) {
    const row     = data[i];
    const rowNum  = i + 1;
    const summary = row[COL.SUMMARY] ? String(row[COL.SUMMARY]).trim() : "";

    // Skip empty rows
    if (!summary) {
      results.skipped++;
      continue;
    }

    const statusCell  = sheet.getRange(rowNum, statusColIdx + 1);
    const notesCell   = sheet.getRange(rowNum, notesColIdx + 1);
    const existingKey = extractIssueKey(statusCell.getValue());

    // Decide the action for this row based on mode + whether a key already exists.
    const action = decideAction(mode, existingKey);

    if (action === "skip-create") {
      // create mode, row already imported — leave it alone
      results.skipped++;
      Logger.log(`⏭️ Row ${rowNum}: already imported as ${existingKey}`);
      appendImportLog(log, rowNum, summary, "SKIPPED", existingKey, "Already imported");
      continue;
    }

    if (action === "skip-update") {
      // update mode, but no key to update against
      results.skipped++;
      Logger.log(`⏭️ Row ${rowNum}: no Jira Key — nothing to update`);
      appendImportLog(log, rowNum, summary, "SKIPPED", "", "No Jira Key to update");
      continue;
    }

    if (action === "update") {
      const outcome = updateIssueWithFallback(row, i, creds, projectKey, existingKey);
      const url     = `${JIRA_BASE_URL}/browse/${existingKey}`;

      if (outcome.ok) {
        if (outcome.warning) {
          results.warnings++;
          setKeyLink(statusCell, existingKey);
          statusCell.setBackground("#ffeb9c");
          writeNote(notesCell, `⚠️ Updated · ${outcome.warning}`, `${existingKey}\n${outcome.warning}\n${url}`);
          Logger.log(`⚠️ Row ${rowNum}: Updated ${existingKey} (${outcome.warning})`);
          appendImportLog(log, rowNum, summary, "UPDATED (partial)", existingKey, outcome.warning);
        } else {
          results.updated++;
          setKeyLink(statusCell, existingKey);
          statusCell.setBackground("#c6efce");
          writeNote(notesCell, "✏️ Updated", `Updated ${existingKey}\n${url}`);
          Logger.log(`✏️ Row ${rowNum}: Updated ${existingKey}`);
          appendImportLog(log, rowNum, summary, "UPDATED", existingKey, "");
        }
      } else {
        recordRowError(results, log, sheet, statusCell, notesCell, rowNum, summary, outcome, existingKey);
      }

      Utilities.sleep(300);
      continue;
    }

    // action === "create"
    const outcome = createIssueWithFallback(row, i, creds, projectKey);

    if (outcome.key) {
      const url = `${JIRA_BASE_URL}/browse/${outcome.key}`;

      if (outcome.warning) {
        results.warnings++;
        setKeyLink(statusCell, outcome.key);
        statusCell.setBackground("#ffeb9c");
        writeNote(notesCell, `⚠️ ${outcome.warning}`, `${outcome.key}\n${outcome.warning}\n${url}`);
        Logger.log(`⚠️ Row ${rowNum}: Created ${outcome.key} (${outcome.warning}) — ${summary}`);
        appendImportLog(log, rowNum, summary, "CREATED (partial)", outcome.key, outcome.warning);
      } else {
        results.created++;
        setKeyLink(statusCell, outcome.key);
        statusCell.setBackground("#c6efce");
        writeNote(notesCell, "✅ Imported", `Created ${outcome.key}\n${url}`);
        Logger.log(`✅ Row ${rowNum}: Created ${outcome.key} — ${summary}`);
        appendImportLog(log, rowNum, summary, "CREATED", outcome.key, "");
      }
    } else {
      recordRowError(results, log, sheet, statusCell, notesCell, rowNum, summary, outcome, "");
    }

    Utilities.sleep(300);
  }

  finishImportLog(log, results, new Date() - runStarted);

  // Summary dialog
  let msg =
    `✅ Created:   ${results.created}\n` +
    `✏️ Updated:   ${results.updated}\n` +
    `⚠️ Partial:   ${results.warnings}   (some fields dropped)\n` +
    `⏭️ Skipped:  ${results.skipped}\n` +
    `❌ Errors:    ${results.errors.length}`;

  if (results.errors.length > 0) {
    msg += `\n\n❌ Error details:\n` + results.errors.slice(0, 8).join("\n");
    if (results.errors.length > 8) msg += `\n…and ${results.errors.length - 8} more (see "Import Log" tab)`;
  }
  msg += `\n\n📋 Full run recorded in the "Import Log" tab.`;

  ui.alert("Done", msg + FOOTER, ui.ButtonSet.OK);
}

// =============================================
// PER-ROW ACTION DECISION
// =============================================
// Returns: "create" | "update" | "skip-create" | "skip-update"
function decideAction(mode, existingKey) {
  if (mode === "update") {
    return existingKey ? "update" : "skip-update";
  }
  if (mode === "sync") {
    return existingKey ? "update" : "create";
  }
  // create mode
  return existingKey ? "skip-create" : "create";
}

// Pull a valid-looking Jira key (e.g. ABC-123) out of a cell value.
// Returns "" if the cell is empty or holds an ERROR marker.
function extractIssueKey(cellValue) {
  if (!cellValue) return "";
  const s = String(cellValue).trim();
  if (s.toUpperCase().indexOf("ERROR") !== -1) return "";
  const m = s.match(/[A-Z][A-Z0-9]+-\d+/i);
  return m ? m[0].toUpperCase() : "";
}

// Shared error recorder for both create and update paths.
function recordRowError(results, log, sheet, statusCell, notesCell, rowNum, summary, outcome, keepKey) {
  results.errors.push(`Row ${rowNum} ("${summary}"): ${firstLine(outcome.error)}`);

  // Preserve an existing key on update failures (keep it clickable); otherwise mark ERROR.
  if (keepKey) {
    setKeyLink(statusCell, keepKey);
    statusCell.setBackground("#ffc7ce");
  } else {
    statusCell.setValue("ERROR").setBackground("#ffc7ce");
  }

  writeNote(notesCell, errorShortLabel(outcome), buildErrorDetail(outcome));
  Logger.log(`❌ Row ${rowNum}: ${outcome.error}`);
  appendImportLog(log, rowNum, summary, "ERROR", keepKey || "", firstLine(outcome.error));
}

// =============================================
// LOGGING / REPORTING HELPERS
// =============================================

// Write a Jira key into a cell as a clickable hyperlink to the issue.
// Falls back to a plain value if RichText isn't available for any reason.
function setKeyLink(cell, key) {
  const url = `${JIRA_BASE_URL}/browse/${key}`;
  try {
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(key)
      .setLinkUrl(url)
      .build();
    cell.setRichTextValue(rich);
  } catch (e) {
    cell.setValue(key);
  }
}

// Write a short value into the notes cell and attach the full detail as a cell comment/note.
function writeNote(cell, shortText, fullDetail) {
  cell.setValue(shortText || "");
  if (fullDetail) {
    cell.setNote(String(fullDetail));
  } else {
    cell.clearNote();
  }
}

// First line only, trimmed — handy for compact one-line summaries.
function firstLine(text) {
  if (!text) return "";
  return String(text).split("\n")[0].trim();
}

// A short cell label for a failed row, based on error category.
function errorShortLabel(outcome) {
  switch (outcome.category) {
    case "auth":       return "❌ Auth failed";
    case "permission": return "❌ No permission";
    case "network":    return "❌ Network error";
    case "field":      return "❌ Field rejected";
    default:           return "❌ Failed";
  }
}

// The full, human-readable detail shown in the cell comment.
function buildErrorDetail(outcome) {
  let detail = outcome.error || "Unknown error";
  if (outcome.httpCode) detail = `HTTP ${outcome.httpCode}\n` + detail;
  if (outcome.dropped && outcome.dropped.length) {
    detail += `\n\nAlready dropped before failing: ${outcome.dropped.map(fieldLabel).join(", ")}`;
  }
  return detail;
}

// Create (or reset) the "Import Log" sheet and write the header rows.
function startImportLog(projectKey, startedAt, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Import Log");
  if (!sheet) {
    sheet = ss.insertSheet("Import Log");
  }
  sheet.clear();

  const stamp    = Utilities.formatDate(startedAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const modeText = (mode || "create").toUpperCase();
  sheet.getRange(1, 1).setValue(`🚀 Jira ${modeText} — project ${projectKey} — started ${stamp}`)
       .setFontWeight("bold");

  const header = ["Time", "Row", "Summary", "Result", "Jira Key", "Detail"];
  sheet.getRange(3, 1, 1, header.length).setValues([header]).setFontWeight("bold");
  sheet.setFrozenRows(3);

  return { sheet: sheet, nextRow: 4 };
}

// Append one row to the run log.
function appendImportLog(log, rowNum, summary, result, key, detail) {
  const t = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
  log.sheet.getRange(log.nextRow, 1, 1, 6).setValues([[
    t, rowNum, summary, result, key || "", detail || ""
  ]]);
  log.nextRow++;
}

// Write a totals footer and tidy up the log sheet.
function finishImportLog(log, results, elapsedMs) {
  const secs = Math.round(elapsedMs / 1000);
  const summary =
    `Done — created ${results.created}, updated ${results.updated || 0}, ` +
    `partial ${results.warnings}, skipped ${results.skipped}, ` +
    `errors ${results.errors.length}  (${secs}s)`;

  log.sheet.getRange(log.nextRow + 1, 1).setValue(summary).setFontWeight("bold");

  // Auto-size the columns for readability.
  for (let c = 1; c <= 6; c++) log.sheet.autoResizeColumn(c);
}

// =============================================
// CREATE ISSUE WITH AUTO-FALLBACK
// =============================================
function createIssueWithFallback(row, rowIndex, creds, projectKey) {
  const skipFields  = {};   // set of field ids to drop on retry
  const dropped     = [];   // field ids we successfully dropped, for reporting
  const maxRetries  = 5;    // guard against loops

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const payload  = buildIssuePayload(row, skipFields, projectKey, false);
      const response = createJiraIssue(payload, creds);

      if (dropped.length > 0) {
        const names = dropped.map(fieldLabel).join(", ");
        return {
          key:     response.key,
          warning: `Created without: ${names}`,
          dropped: dropped.slice(),
        };
      }
      return { key: response.key };

    } catch (e) {
      // Only field errors are recoverable by dropping fields.
      const recoverable = e.category === "field"
        && Array.isArray(e.badFields)
        && e.badFields.length > 0;

      // Never drop a field Jira says is REQUIRED — dropping it guarantees failure.
      const required = Array.isArray(e.requiredFields) ? e.requiredFields : [];

      // Find rejected fields we haven't already dropped, that we may skip, and that
      // aren't required by Jira.
      const newlyBad = recoverable
        ? e.badFields.filter(f => !skipFields[f] && isSkippableField(f) && required.indexOf(f) === -1)
        : [];

      if (newlyBad.length === 0) {
        // Nothing left to try — surface the full structured error.
        return {
          error:    e.message,
          category: e.category || "http",
          httpCode: e.httpCode || null,
          raw:      e.rawResponse || "",
          dropped:  dropped.slice(),
        };
      }

      newlyBad.forEach(f => { skipFields[f] = true; dropped.push(f); });
      Logger.log(`⚠️ Row ${rowIndex + 1}: retrying without ${newlyBad.map(fieldLabel).join(", ")}`);
    }
  }

  return {
    error:    `Gave up after ${maxRetries} retries; Jira kept rejecting fields.`,
    category: "field",
    dropped:  dropped.slice(),
  };
}

// Fields we are allowed to drop and still produce a useful issue.
// Core fields (summary, project, issuetype) and Epic Name must never be dropped —
// dropping Epic Name when creating an Epic guarantees a "required" failure.
function isSkippableField(fieldId) {
  const NEVER_DROP = ["summary", "project", "issuetype", CUSTOM_FIELDS.EPIC_NAME];
  return NEVER_DROP.indexOf(fieldId) === -1;
}

// =============================================
// UPDATE ISSUE WITH AUTO-FALLBACK
// =============================================
function updateIssueWithFallback(row, rowIndex, creds, projectKey, issueKey) {
  const skipFields = {};
  const dropped    = [];
  const maxRetries = 5;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const payload = buildIssuePayload(row, skipFields, projectKey, true);
      updateJiraIssue(issueKey, payload, creds);   // 204 No Content on success

      if (dropped.length > 0) {
        return { ok: true, warning: `Updated without: ${dropped.map(fieldLabel).join(", ")}`, dropped: dropped.slice() };
      }
      return { ok: true };

    } catch (e) {
      const recoverable = e.category === "field"
        && Array.isArray(e.badFields)
        && e.badFields.length > 0;

      const newlyBad = recoverable
        ? e.badFields.filter(f => !skipFields[f] && isSkippableField(f))
        : [];

      if (newlyBad.length === 0) {
        return {
          ok:       false,
          error:    e.message,
          category: e.category || "http",
          httpCode: e.httpCode || null,
          raw:      e.rawResponse || "",
          dropped:  dropped.slice(),
        };
      }

      newlyBad.forEach(f => { skipFields[f] = true; dropped.push(f); });
      Logger.log(`⚠️ Row ${rowIndex + 1}: retry update of ${issueKey} without ${newlyBad.map(fieldLabel).join(", ")}`);
    }
  }

  return {
    ok:       false,
    error:    `Gave up after ${maxRetries} retries; Jira kept rejecting fields.`,
    category: "field",
    dropped:  dropped.slice(),
  };
}

// =============================================
// BUILD JIRA ISSUE PAYLOAD
// =============================================
// isUpdate=true omits fields the edit endpoint won't accept the same way
// (project can't change; issuetype/parent moves need a dedicated API).
function buildIssuePayload(row, skipFields, projectKey, isUpdate) {
  skipFields = skipFields || {};
  const skip = (fieldId) => skipFields[fieldId] === true;

  const fields = {
    summary: String(row[COL.SUMMARY]).trim(),
  };

  if (!isUpdate) {
    // Create-only fields: project & issue type are set at creation time.
    const issueType  = row[COL.ISSUE_TYPE] ? String(row[COL.ISSUE_TYPE]).trim() : "Task";

    // Project routing: Service Requests are ALWAYS created in DTS; everything
    // else uses the project chosen at the prompt. (Create-only — project can't
    // change on an existing issue.)
    const targetProject = (issueType.toLowerCase() === SERVICE_REQUEST_ISSUE_TYPE)
      ? SERVICE_REQUEST_PROJECT
      : projectKey;

    fields.project   = { key: targetProject };
    fields.issuetype = { name: issueType };

    // Epic Name — required when creating an Epic on this instance (customfield_10371).
    // createmeta hides it, but a validator enforces it, so always send it for Epics,
    // defaulting to the summary.
    if (!skip(CUSTOM_FIELDS.EPIC_NAME) && issueType.toLowerCase() === "epic") {
      fields[CUSTOM_FIELDS.EPIC_NAME] = String(row[COL.SUMMARY]).trim();
    }
  }

  // Description
  if (!skip("description") && row[COL.DESCRIPTION]) {
    fields.description = toADF(String(row[COL.DESCRIPTION]));
  }

  // Acceptance Criteria
  if (!skip(CUSTOM_FIELDS.ACCEPTANCE_CRITERIA) && row[COL.ACCEPTANCE_CRITERIA]) {
    fields[CUSTOM_FIELDS.ACCEPTANCE_CRITERIA] = toADF(String(row[COL.ACCEPTANCE_CRITERIA]));
  }

  // Original Estimate — cell holds SECONDS. Jira's edit/create endpoints reject a bare
  // "timeoriginalestimate" number; the reliable way is the "timetracking" structure with
  // a duration string. Minutes ("Nm") avoid Jira's working-day/week conversion surprises.
  if (!skip("timetracking") && !skip(CUSTOM_FIELDS.ORIGINAL_ESTIMATE) && row[COL.ORIGINAL_ESTIMATE]) {
    const seconds = Number(row[COL.ORIGINAL_ESTIMATE]);
    if (!isNaN(seconds) && seconds > 0) {
      const minutes = Math.round(seconds / 60);
      fields.timetracking = { originalEstimate: `${minutes}m` };
    }
  }

  // Parent issue — set on create only (re-parenting via edit is unreliable across Jira configs)
  if (!isUpdate && !skip("parent") && row[COL.PARENT]) {
    fields.parent = { key: String(row[COL.PARENT]).trim() };
  }

  // Initiative — array of strings
  if (!skip(CUSTOM_FIELDS.INITIATIVE) && row[COL.INITIATIVE] && String(row[COL.INITIATIVE]).trim() !== "") {
    fields[CUSTOM_FIELDS.INITIATIVE] = [String(row[COL.INITIATIVE]).trim()];
  }

    // Start date → yyyy-MM-dd (handles both DD/MM/YYYY text and Google Sheets date)
  if (!skip(CUSTOM_FIELDS.START_DATE) && row[COL.START_DATE] && String(row[COL.START_DATE]).trim() !== "") {
    let startDateFormatted = null;

    if (row[COL.START_DATE] instanceof Date) {
      // Google Sheets date cell
      startDateFormatted = Utilities.formatDate(row[COL.START_DATE], "UTC", "yyyy-MM-dd");
    } else {
      // Plain text — handle DD/MM/YYYY
      const parts = String(row[COL.START_DATE]).trim().split("/");
      if (parts.length === 3) {
        // parts[0]=DD, parts[1]=MM, parts[2]=YYYY
        startDateFormatted = `${parts[2]}-${parts[1].padStart(2,"0")}-${parts[0].padStart(2,"0")}`;
      }
    }

    if (startDateFormatted) {
      fields[CUSTOM_FIELDS.START_DATE] = startDateFormatted;
    } else {
      Logger.log(`⚠️ Could not parse Start Date: "${row[COL.START_DATE]}" — skipping`);
    }
  }


  // Due date → yyyy-MM-dd
  if (!skip("duedate") && row[COL.DUE_DATE]) {
    const d = new Date(row[COL.DUE_DATE]);
    if (!isNaN(d)) {
      fields.duedate = Utilities.formatDate(d, "UTC", "yyyy-MM-dd");
    }
  }

  // CS Functional Group — array of {id}
  if (!skip(CUSTOM_FIELDS.CS_FUNCTIONAL_GROUP) && row[COL.CS_FUNCTIONAL_GROUP] && String(row[COL.CS_FUNCTIONAL_GROUP]).trim() !== "") {
    const key = String(row[COL.CS_FUNCTIONAL_GROUP]).trim().toLowerCase();
    const id  = CS_FUNCTIONAL_GROUP_MAP[key];
    if (id) {
      fields[CUSTOM_FIELDS.CS_FUNCTIONAL_GROUP] = [{ id: id }];
    } else {
      Logger.log(`⚠️ Unknown CS Functional Group: "${row[COL.CS_FUNCTIONAL_GROUP]}" — skipping`);
    }
  }

  // Approved Hours — number
  if (!skip(CUSTOM_FIELDS.APPROVED_HOURS) && row[COL.APPROVED_HOURS] !== "" && row[COL.APPROVED_HOURS] !== null && row[COL.APPROVED_HOURS] !== undefined) {
    const hours = Number(row[COL.APPROVED_HOURS]);
    if (!isNaN(hours)) {
      fields[CUSTOM_FIELDS.APPROVED_HOURS] = hours;
    }
  }

  // Account — map plain text to {id, value}
  if (!skip(CUSTOM_FIELDS.ACCOUNT) && row[COL.ACCOUNT] && String(row[COL.ACCOUNT]).trim() !== "") {
    const accountVal = parseAccountValue(row[COL.ACCOUNT]);
    if (accountVal) {
      fields[CUSTOM_FIELDS.ACCOUNT] = accountVal;
    } else {
      Logger.log(`⚠️ Unknown Account value: "${row[COL.ACCOUNT]}" — skipping`);
    }
  }

  // Labels — Jira's standard "labels" field (array of strings). The cell is
  // comma-separated; split, trim, and drop empties. Labels cannot contain
  // spaces in Jira, so type them with underscores (e.g. DT_Activation_Team).
  // Editable on both create AND update, so it is NOT gated by isUpdate.
  if (!skip("labels") && row[COL.LABELS] && String(row[COL.LABELS]).trim() !== "") {
    const labels = String(row[COL.LABELS])
      .split(",")
      .map(s => s.trim())
      .filter(s => s !== "");
    if (labels.length) {
      fields.labels = labels;
    }
  }

  return { fields };
}

// =============================================
// HELPER: normalise an Account cell into Jira's { id, value } select format
// =============================================
// Accepts any of:
//   • a JSON object pasted in the cell, e.g. {"id": 4, "value": "Post GoLive billable"}
//   • plain text that matches ACCOUNT_MAP, e.g. "Post GoLive billable"
//   • a bare id, e.g. "4"
// Jira expects the id as a STRING, so numeric ids are coerced.
function parseAccountValue(raw) {
  const text = String(raw).trim();

  // 1) JSON object in the cell
  if (text.charAt(0) === "{") {
    try {
      const obj = JSON.parse(text);
      const out = {};
      if (obj.id !== undefined && obj.id !== null) out.id = String(obj.id);
      if (obj.value !== undefined && obj.value !== null) out.value = String(obj.value);
      // A single-select field needs at least one of id/value.
      if (out.id || out.value) return out;
    } catch (e) {
      Logger.log(`⚠️ Account looked like JSON but failed to parse: ${e.message}`);
    }
  }

  // 2) Plain-text label mapped via ACCOUNT_MAP
  const mapped = ACCOUNT_MAP[text.toLowerCase()];
  if (mapped) return mapped;

  // 3) Bare numeric id
  if (/^\d+$/.test(text)) return { id: text };

  return null;
}

// =============================================
// HELPER: Convert text to Atlassian Doc Format
// =============================================
function toADF(text) {
  return {
    type: "doc",
    version: 1,
    content: [{
      type: "paragraph",
      content: [{ type: "text", text: text }]
    }]
  };
}

// =============================================
// JIRA REST API — Create Issue
// =============================================
function createJiraIssue(payload, creds) {
  const url         = `${JIRA_BASE_URL}/rest/api/3/issue`;
  const credentials = Utilities.base64Encode(`${creds.email}:${creds.token}`);

  const options = {
    method:             "post",
    contentType:        "application/json",
    headers:            {
      "Authorization": `Basic ${credentials}`,
      "Accept":        "application/json"
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    // Network / transport-level failure (DNS, timeout, SSL, etc.)
    const err = new Error(`Network error: ${e.message}`);
    err.category    = "network";
    err.httpCode    = 0;
    err.badFields   = [];
    err.rawResponse = "";
    throw err;
  }

  const code    = response.getResponseCode();
  const rawText = response.getContentText();

  if (code === 201) {
    return JSON.parse(rawText);
  }

  // ---- Non-success: build a structured, readable error ----
  let body = {};
  try { body = JSON.parse(rawText); } catch (_) { /* non-JSON body */ }

  const parsed = parseJiraError(code, body, rawText);
  const err    = new Error(parsed.message);
  err.category    = parsed.category;   // "auth" | "permission" | "field" | "http"
  err.httpCode    = code;
  err.badFields   = parsed.badFields;  // array of field ids that Jira rejected
  err.requiredFields = parsed.requiredFields;  // subset Jira says are required
  err.rawResponse = rawText;
  throw err;
}

// =============================================
// JIRA REST API — Update Issue (PUT /issue/{key})
// =============================================
// Returns true on success. A successful edit responds 204 No Content.
function updateJiraIssue(issueKey, payload, creds) {
  const url         = `${JIRA_BASE_URL}/rest/api/3/issue/${encodeURIComponent(issueKey)}`;
  const credentials = Utilities.base64Encode(`${creds.email}:${creds.token}`);

  const options = {
    method:             "put",
    contentType:        "application/json",
    headers:            {
      "Authorization": `Basic ${credentials}`,
      "Accept":        "application/json"
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    const err = new Error(`Network error: ${e.message}`);
    err.category    = "network";
    err.httpCode    = 0;
    err.badFields   = [];
    err.rawResponse = "";
    throw err;
  }

  const code    = response.getResponseCode();
  const rawText = response.getContentText();

  // 204 (No Content) is the normal success; some proxies return 200.
  if (code === 204 || code === 200) {
    return true;
  }

  // 404 → the key doesn't exist / not visible to this account.
  if (code === 404) {
    const err = new Error(`Issue "${issueKey}" not found (HTTP 404). Check the key and your access.`);
    err.category    = "http";
    err.httpCode    = 404;
    err.badFields   = [];
    err.rawResponse = rawText;
    throw err;
  }

  let body = {};
  try { body = JSON.parse(rawText); } catch (_) { /* non-JSON body */ }

  const parsed = parseJiraError(code, body, rawText);
  const err    = new Error(parsed.message);
  err.category    = parsed.category;
  err.httpCode    = code;
  err.badFields   = parsed.badFields;
  err.requiredFields = parsed.requiredFields;
  err.rawResponse = rawText;
  throw err;
}

// =============================================
// PARSE A JIRA ERROR RESPONSE INTO SOMETHING READABLE
// =============================================
// Returns: { message, category, badFields: [fieldId,...] }
function parseJiraError(code, body, rawText) {
  const badFields      = [];
  const requiredFields = [];   // fields Jira says are REQUIRED — must never be dropped
  const parts          = [];

  // Field-specific errors: { errors: { "customfield_16414": "msg", ... } }
  if (body && body.errors && typeof body.errors === "object") {
    for (const fieldId in body.errors) {
      badFields.push(fieldId);
      const detail = body.errors[fieldId] || "(no message — field may be required or not on the create screen)";
      if (/is required|must be provided|cannot be empty/i.test(String(detail))) {
        requiredFields.push(fieldId);
      }
      parts.push(`• ${fieldLabel(fieldId)}: ${detail}`);
    }
  }

  // Global messages: { errorMessages: ["...", "..."] }
  if (body && Array.isArray(body.errorMessages) && body.errorMessages.length) {
    body.errorMessages.forEach(m => parts.push(`• ${m}`));
  }

  // Some 400s return a single top-level message instead of the above shapes.
  if (!parts.length && body && typeof body.message === "string" && body.message) {
    parts.push(`• ${body.message}`);
  }

  // Categorise for friendlier top-line guidance
  let category = "http";
  if (code === 401)                     category = "auth";
  else if (code === 403)                category = "permission";
  else if (badFields.length > 0)        category = "field";

  let headline;
  switch (category) {
    case "auth":
      headline = "Authentication failed (HTTP 401). Check your email + API token in Setup.";
      break;
    case "permission":
      headline = "Permission denied (HTTP 403). Your account may lack rights to create issues in this project.";
      break;
    case "field":
      headline = `Jira rejected ${badFields.length} field(s) (HTTP ${code}).`;
      break;
    default:
      headline = `Jira returned HTTP ${code}.`;
  }

  let message = headline;
  if (parts.length) message += "\n" + parts.join("\n");
  // Fall back to raw text if Jira gave us nothing structured
  if (!parts.length && rawText) message += `\n${String(rawText).slice(0, 600)}`;

  return { message, category, badFields, requiredFields };
}

// =============================================
// TEST CONNECTION (menu item)
// =============================================
function testConnection() {
  const creds = requireCredentials();
  if (!creds) return;

  const ok = testConnectionSilent(creds.email, creds.token);
  if (ok.success) {
    SpreadsheetApp.getUi().alert(
      "✅ Connection OK",
      `Connected as: ${ok.displayName} (${ok.email})` + FOOTER,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "❌ Connection Failed",
      ok.error + FOOTER,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// TEST CONNECTION (silent — returns result object)
function testConnectionSilent(email, token) {
  try {
    const credentials = Utilities.base64Encode(`${email}:${token}`);
    const response    = UrlFetchApp.fetch(`${JIRA_BASE_URL}/rest/api/3/myself`, {
      method:             "get",
      headers:            {
        "Authorization": `Basic ${credentials}`,
        "Accept":        "application/json"
      },
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    const body = JSON.parse(response.getContentText());

    if (code === 200) {
      return { success: true, displayName: body.displayName, email: body.emailAddress };
    } else {
      return { success: false, error: `HTTP ${code}: ${response.getContentText()}` };
    }
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// =============================================
// EXPORT FROM JIRA → SHEET (JQL)
// =============================================
// Pulls issues OUT of Jira into a dedicated "Jira Export" sheet for PM work,
// replacing the "Jira Cloud for Sheets" extension. The column set and order are
// fixed here so you always get the exact same layout without manual setup.
//
// Each column declares:
//   header — the human column title written to row 1
//   field  — the Jira field id requested via /search (null for derived columns)
//   get    — (fields) => cell value; formats the raw field into something readable
// =============================================
// NOTE: values are formatted to match the "Jira Cloud for Sheets" extension output
// so this export is a drop-in replacement:
//   • Date/time → "DD/MM/YYYY HH:MM:SS" text (Created/Updated); "DD/MM/YYYY" for date-only.
//   • Time estimates → RAW SECONDS (so downstream /3600 formulas keep working).
//   • Parent / Initiative (parent) → the plain issue key (e.g. "RBI-785").
const EXPORT_COLUMNS = [
  { header: "Issue Type",           field: "issuetype",             get: f => name(f.issuetype) },
  { header: "Key",                  field: "__key__",               get: (f, issue) => issue.key },
  { header: "Summary",              field: "summary",               get: f => f.summary || "" },
  { header: "Assignee",             field: "assignee",              get: f => displayName(f.assignee) },
  { header: "Reporter",             field: "reporter",              get: f => displayName(f.reporter) },
  { header: "Priority",             field: "priority",              get: f => name(f.priority) },
  { header: "Status",               field: "status",                get: f => name(f.status) },
  { header: "Created",              field: "created",               get: f => fmtDateTime(f.created) },
  { header: "Updated",              field: "updated",               get: f => fmtDateTime(f.updated) },
  { header: "Start date",           field: "customfield_16414",     get: f => fmtDate(f.customfield_16414) },
  { header: "Due date",             field: "duedate",               get: f => fmtDate(f.duedate) },
  { header: "parent",               field: "parent",                get: f => f.parent ? f.parent.key : "" },
  { header: "Remaining Estimate",   field: "timeestimate",          get: f => secondsValue(f.timeestimate) },
  { header: "Account",              field: "customfield_16433",     get: f => optionValue(f.customfield_16433) },
  { header: "Σ Remaining Estimate", field: "aggregatetimeestimate", get: f => secondsValue(f.aggregatetimeestimate) },
  { header: "Visibility",           field: "customfield_11791",     get: f => anyToText(f.customfield_11791) },
  { header: "Who (Company)",        field: "customfield_12013",     get: f => anyToText(f.customfield_12013) },
  { header: "Who (Person)",         field: "customfield_12014",     get: f => anyToText(f.customfield_12014) },
  { header: "Agreed By",            field: "customfield_12015",     get: f => anyToText(f.customfield_12015) },
  { header: "CS Functional Group",  field: "customfield_14391",     get: f => anyToText(f.customfield_14391) },
  { header: "Original estimate",    field: "timeoriginalestimate",  get: f => secondsValue(f.timeoriginalestimate) },
  { header: "Initiative (parent)",  field: "customfield_16455",     get: f => issueKeyOf(f.customfield_16455) },
];

// Where the JQL query is read from: the "Assets" sheet, cell B3.
const JQL_SOURCE_SHEET = "Assets";
const JQL_SOURCE_CELL  = "B3";
// Where the timestamp of the last successful import is written: Assets!B5.
const LAST_RUN_CELL    = "B5";

// Menu entry point. Reads the JQL from Assets!B3 and pulls matching issues.
function exportFromJira() {
  const creds = requireCredentials();
  if (!creds) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read the JQL from Assets!B3.
  const assets = ss.getSheetByName(JQL_SOURCE_SHEET);
  if (!assets) {
    ui.alert(
      "❌ Missing \"" + JQL_SOURCE_SHEET + "\" sheet",
      `Create a sheet named "${JQL_SOURCE_SHEET}" and put your JQL query in cell ` +
      `${JQL_SOURCE_CELL}.` + FOOTER,
      ui.ButtonSet.OK
    );
    return;
  }

  let jql = String(assets.getRange(JQL_SOURCE_CELL).getValue() || "").trim();
  if (!jql) {
    ui.alert(
      "❌ No JQL to run",
      `There's no query to run.\n\n` +
      `Cell ${JQL_SOURCE_SHEET}!${JQL_SOURCE_CELL} is empty, so Get Jira Data has ` +
      `nothing to fetch.\n\n` +
      `Fix: type a JQL query into ${JQL_SOURCE_SHEET}!${JQL_SOURCE_CELL}, for example\n` +
      `  project = RBI ORDER BY created ASC\n` +
      `(a bare project key like "RBI" also works), then run Get Jira Data again.` + FOOTER,
      ui.ButtonSet.OK
    );
    return;
  }

  // Allow a bare project key as a shortcut.
  if (/^[A-Z][A-Z0-9]+$/i.test(jql)) {
    jql = `project = ${jql.toUpperCase()} ORDER BY created DESC`;
  }

  // One-click: no confirmation — fetch straight from Assets!B3.
  let issues;
  try {
    issues = searchAllIssues(jql, creds);
  } catch (e) {
    ui.alert("❌ Get Jira Data failed", (e.message || String(e)) + FOOTER, ui.ButtonSet.OK);
    return;
  }

  if (!issues.length) {
    ui.alert("No results", `JQL returned 0 issues:\n${jql}` + FOOTER, ui.ButtonSet.OK);
    return;
  }

  writeExportSheet(issues, jql);

  // Stamp the last successful run into Assets!B5 as a real datetime value
  // (so it sorts/filters and supports date math), formatted for display.
  const ranAt = new Date();
  assets.getRange(LAST_RUN_CELL)
        .setValue(ranAt)
        .setNumberFormat("dd/MM/yyyy HH:mm:ss");

  const ranAtText = Utilities.formatDate(ranAt, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  ui.alert(
    "✅ Done",
    `Pulled ${issues.length} issue(s) into the "Jira" tab.\n\n` +
    `Last run stamped in ${JQL_SOURCE_SHEET}!${LAST_RUN_CELL}: ${ranAtText}` + FOOTER,
    ui.ButtonSet.OK
  );
}

// =============================================
// JIRA REST API — paged JQL search (enhanced endpoint)
// =============================================
// Uses POST /rest/api/3/search/jql. The legacy /rest/api/3/search was removed by
// Atlassian (HTTP 410, CHANGE-2046). This endpoint uses CURSOR pagination:
//   • no startAt / total
//   • each response returns a nextPageToken; absent/null means "last page"
//   • maxResults is capped at 100
// Requests only the fields our columns need, to keep responses small.
function searchAllIssues(jql, creds) {
  const url         = `${JIRA_BASE_URL}/rest/api/3/search/jql`;
  const credentials = Utilities.base64Encode(`${creds.email}:${creds.token}`);

  // Collect the distinct field ids we actually need (skip derived columns).
  const fields = EXPORT_COLUMNS
    .map(c => c.field)
    .filter(f => f && f !== "__key__");

  const all       = [];
  const pageSize  = 100;   // enhanced endpoint hard cap
  const maxPages  = 200;   // safety cap: up to 20,000 issues
  let nextToken   = null;

  for (let page = 0; page < maxPages; page++) {
    const payload = {
      jql:        jql,
      maxResults: pageSize,
      fields:     fields,
    };
    if (nextToken) payload.nextPageToken = nextToken;

    const response = UrlFetchApp.fetch(url, {
      method:             "post",
      contentType:        "application/json",
      headers:            { "Authorization": `Basic ${credentials}`, "Accept": "application/json" },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code    = response.getResponseCode();
    const rawText = response.getContentText();

    if (code !== 200) {
      let body = {};
      try { body = JSON.parse(rawText); } catch (_) {}
      const parsed = parseJiraError(code, body, rawText);
      throw new Error(parsed.message);
    }

    const body = JSON.parse(rawText);
    (body.issues || []).forEach(issue => all.push(issue));

    // Advance the cursor. Stop when there's no next page token.
    nextToken = body.nextPageToken || null;
    if (!nextToken || body.isLast === true) break;

    Utilities.sleep(200);   // be gentle on the API between pages
  }

  return all;
}

// =============================================
// WRITE RESULTS INTO THE "Jira Export" SHEET
// =============================================
// Regenerates a dedicated tab so repeated exports never accumulate stale rows.
function writeExportSheet(issues, jql) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Always land on a sheet called "Jira": reuse it (cleaning old data below) if
  // it exists, otherwise create it.
  let sheet = ss.getSheetByName("Jira");
  if (!sheet) {
    sheet = ss.insertSheet("Jira");
  }

  const headers = EXPORT_COLUMNS.map(c => c.header);
  const numCols = headers.length;   // 22 — columns A..V

  // Only clear the 22 data columns we own (A..V). Any helper columns the user
  // adds to the RIGHT (e.g. the =ARRAYFORMULA(.../3600) column) are left intact.
  const maxRows = Math.max(sheet.getMaxRows(), 1);
  sheet.getRange(1, 1, maxRows, numCols).clearContent().clearNote();

  // Build the 2D value matrix: one row per issue, one cell per column.
  const rows = issues.map(issue => {
    const f = issue.fields || {};
    return EXPORT_COLUMNS.map(col => {
      try {
        const v = col.get(f, issue);
        return (v === null || v === undefined) ? "" : v;
      } catch (e) {
        return "";   // never let one odd field break the whole export
      }
    });
  });

  // Header row.
  sheet.getRange(1, 1, 1, numCols).setValues([headers]).setFontWeight("bold");

  // Data rows (issues only — no trailing blank padding).
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
  }

  // Make the Key column a clickable link to the issue.
  const keyColIdx = EXPORT_COLUMNS.findIndex(c => c.field === "__key__");
  if (keyColIdx !== -1) {
    for (let r = 0; r < issues.length; r++) {
      setKeyLink(sheet.getRange(r + 2, keyColIdx + 1), issues[r].key);
    }
  }

  sheet.setFrozenRows(1);
  for (let c = 1; c <= numCols; c++) sheet.autoResizeColumn(c);
  SpreadsheetApp.setActiveSheet(sheet);
}

// =============================================
// FIELD FORMATTERS — turn raw Jira values into readable cell text
// =============================================

// A field that exposes a ".name" (issuetype, status, priority).
function name(obj) {
  return obj && obj.name ? obj.name : "";
}

// A user field → display name.
function displayName(user) {
  return user && user.displayName ? user.displayName : "";
}

// A single-select custom field → its "value".
function optionValue(opt) {
  if (!opt) return "";
  if (Array.isArray(opt)) return opt.map(o => o && (o.value || o.name) || "").filter(Boolean).join(", ");
  return opt.value || opt.name || "";
}

// ISO timestamp → "DD/MM/YYYY HH:MM:SS" in the script's timezone.
// Matches the Jira Cloud for Sheets extension's date/time output.
function fmtDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

// Date-only field → "DD/MM/YYYY". Matches the extension's output.
function fmtDate(val) {
  if (!val) return "";
  const d = new Date(val);
  if (isNaN(d)) return String(val);
  return Utilities.formatDate(d, "UTC", "dd/MM/yyyy");
}

// Time estimate → RAW SECONDS as a number (e.g. 28800), matching the extension.
// Returns "" when the field is absent so empty rows stay blank; returns 0 for a
// real zero estimate. Downstream sheet formulas (e.g. =value/3600) rely on this.
function secondsValue(seconds) {
  if (seconds === null || seconds === undefined || seconds === "") return "";
  const s = Number(seconds);
  return isNaN(s) ? "" : s;
}

// Pull an issue key out of a "parent link"-style field. The value may arrive as
// a plain string ("RBI-785") or an object with a .key; normalise to the key text.
function issueKeyOf(val) {
  if (!val) return "";
  if (typeof val === "string") return val;
  if (Array.isArray(val)) return val.map(issueKeyOf).filter(Boolean).join(", ");
  if (typeof val === "object") {
    if (val.key !== undefined && val.key !== null) return String(val.key);
    if (val.value !== undefined && val.value !== null) return String(val.value);
  }
  return "";
}

// Best-effort stringifier for custom fields of unknown/various shapes:
// strings, numbers, {value}/{name}/{displayName} objects, or arrays of those.
function anyToText(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "string" || typeof val === "number" || typeof val === "boolean") {
    return String(val);
  }
  if (Array.isArray(val)) {
    return val.map(anyToText).filter(Boolean).join(", ");
  }
  if (typeof val === "object") {
    if (val.value       !== undefined) return String(val.value);
    if (val.name        !== undefined) return String(val.name);
    if (val.displayName !== undefined) return String(val.displayName);
    if (val.key         !== undefined) return String(val.key);
    // ADF rich text → flatten to plain text.
    if (val.type === "doc" && Array.isArray(val.content)) return adfToText(val);
  }
  return "";
}

// Flatten Atlassian Document Format into plain text (for rich-text custom fields).
function adfToText(node) {
  if (!node) return "";
  if (node.type === "text") return node.text || "";
  if (!Array.isArray(node.content)) return "";
  const sep = node.type === "paragraph" ? "\n" : "";
  return node.content.map(adfToText).join("") + sep;
}
