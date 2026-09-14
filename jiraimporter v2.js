// =============================================
// v2-gantt template
// 🚀 Jira Importer by Carloszai
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
  "duedate":             "Due Date",
  "timeoriginalestimate":"Original Estimate",
  "timetracking":        "Original Estimate (time tracking)",
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
// COLUMN INDEX MAP (0-based)
// A=0  B=1  C=2  D=3  E=4  F=5  G=6  H=7  I=8  J=9  K=10  L=11
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
const FOOTER = "\n\n─────────────────────────\n🚀 Jira Importer by Carloszai\n📧 carlos.lourenco@feedzai.com";

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
    .addItem("❓ Help — What do these modes do?", "showModeHelp")
    .addSeparator()
    .addItem("Import ALL Rows to Jira",      "importAllToJira")
    .addItem("Import Selected Rows to Jira", "importSelectedToJira")
    .addSeparator()
    .addItem("🔄 Sync ALL Rows (create + update)",      "syncAllToJira")
    .addItem("🔄 Sync Selected Rows (create + update)", "syncSelectedToJira")
    .addSeparator()
    .addItem("✏️ Update ALL Rows (with Jira Key)",      "updateAllInJira")
    .addItem("✏️ Update Selected Rows (with Jira Key)", "updateSelectedInJira")
    .addSeparator()
    .addItem("🧪 Test Connection",           "testConnection")
    .addSeparator()
    .addItem("⚙️ Setup My Jira Credentials", "setupCredentials")
    .addItem("🗑️ Clear My Credentials",      "clearCredentials")
    .addToUi();
}

// =============================================
// HELP — explain what each mode does
// =============================================
function showModeHelp() {
  const msg =
    "All modes write from this sheet TO Jira. What differs is which rows they act on,\n" +
    "based on whether a row already has a value in the \"Jira Key\" column.\n" +
    "\n" +
    "──────────────────────────\n" +
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
    "──────────────────────────\n" +
    "Good to know:\n" +
    "   • Updates send whatever is in the sheet, not a diff — every mapped field\n" +
    "     with a value is pushed on each run.\n" +
    "   • A BLANK cell does NOT clear the field in Jira; it is simply left as-is.\n" +
    "   • Project & Issue Type are only set when creating, not when updating.";

  SpreadsheetApp.getUi().alert(
    "❓ How the modes work",
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

  const projectKey = askProjectKey();
  if (!projectKey) return;

  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const allData = sheet.getDataRange().getValues();

  let from, to;
  if (selectionOnly) {
    const selection = sheet.getActiveRange();
    from = Math.max(selection.getRow() - 1, 1);
    to   = selection.getLastRow() - 1;
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
          statusCell.setValue(existingKey).setBackground("#ffeb9c");
          writeNote(notesCell, `⚠️ Updated · ${outcome.warning}`, `${existingKey}\n${outcome.warning}\n${url}`);
          Logger.log(`⚠️ Row ${rowNum}: Updated ${existingKey} (${outcome.warning})`);
          appendImportLog(log, rowNum, summary, "UPDATED (partial)", existingKey, outcome.warning);
        } else {
          results.updated++;
          statusCell.setValue(existingKey).setBackground("#c6efce");
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
        statusCell.setValue(outcome.key).setBackground("#ffeb9c");
        writeNote(notesCell, `⚠️ ${outcome.warning}`, `${outcome.key}\n${outcome.warning}\n${url}`);
        Logger.log(`⚠️ Row ${rowNum}: Created ${outcome.key} (${outcome.warning}) — ${summary}`);
        appendImportLog(log, rowNum, summary, "CREATED (partial)", outcome.key, outcome.warning);
      } else {
        results.created++;
        statusCell.setValue(outcome.key).setBackground("#c6efce");
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

  // Preserve an existing key on update failures; otherwise mark ERROR.
  if (keepKey) {
    statusCell.setValue(keepKey).setBackground("#ffc7ce");
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

      // Find rejected fields we haven't already dropped and that we know how to skip.
      const newlyBad = recoverable
        ? e.badFields.filter(f => !skipFields[f] && isSkippableField(f))
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
// Core fields (summary, project, issuetype) must never be dropped.
function isSkippableField(fieldId) {
  const REQUIRED = ["summary", "project", "issuetype"];
  return REQUIRED.indexOf(fieldId) === -1;
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
    fields.project   = { key: projectKey };
    fields.issuetype = { name: row[COL.ISSUE_TYPE] ? String(row[COL.ISSUE_TYPE]).trim() : "Task" };
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
    const key        = String(row[COL.ACCOUNT]).trim().toLowerCase();
    const accountVal = ACCOUNT_MAP[key];
    if (accountVal) {
      fields[CUSTOM_FIELDS.ACCOUNT] = accountVal;
    } else {
      Logger.log(`⚠️ Unknown Account value: "${row[COL.ACCOUNT]}" — skipping`);
    }
  }

  return { fields };
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
  err.rawResponse = rawText;
  throw err;
}

// =============================================
// PARSE A JIRA ERROR RESPONSE INTO SOMETHING READABLE
// =============================================
// Returns: { message, category, badFields: [fieldId,...] }
function parseJiraError(code, body, rawText) {
  const badFields = [];
  const parts     = [];

  // Field-specific errors: { errors: { "customfield_16414": "msg", ... } }
  if (body && body.errors && typeof body.errors === "object") {
    for (const fieldId in body.errors) {
      badFields.push(fieldId);
      parts.push(`• ${fieldLabel(fieldId)}: ${body.errors[fieldId]}`);
    }
  }

  // Global messages: { errorMessages: ["...", "..."] }
  if (body && Array.isArray(body.errorMessages) && body.errorMessages.length) {
    body.errorMessages.forEach(m => parts.push(`• ${m}`));
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
  if (!parts.length && rawText) message += `\n${String(rawText).slice(0, 300)}`;

  return { message, category, badFields };
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
