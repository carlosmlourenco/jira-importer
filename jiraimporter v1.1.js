// =============================================
// v2-gantt template
// 🚀 Jira Importer by Carloszai
// =============================================

// =============================================
// FIELD REFERENCE — available fields for future use
// =============================================
// # of Approved Hours          customfield_13750   number
// # of Requested Hours         customfield_14083   number
// Total Remaining Hours        customfield_20108   number
// Non-billable hours helper    customfield_14471   string
// Approvers                    customfield_11905   array
// Pre-Approved                 customfield_12644   option
// Data Science Approver        customfield_11786   user
// Engineering Approver         customfield_11785   user
// Cloud Operations Approvers   customfield_12580   array
// COPS Financial Approvers     customfield_15471   array
// Approval from Groups         customfield_16403   array
// Approved CSSM Depl. Request  customfield_16434   array
// Approvals                    customfield_11870   sd-approvals
// Approved by                  customfield_11170   user
// Approved by (read only)      customfield_11171   string
// Approved on                  customfield_11172   datetime
// Approved on (read only)      customfield_11173   string
// Funding Proof of Approval    customfield_22700   string
// Expedite Priority            customfield_20340   string
// Expedite Approved By         customfield_26408   option
// Approvers (migrated)         customfield_20726   array
// Approver groups (migrated)   customfield_20790   array

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
    .addItem("Import ALL Rows to Jira",      "importAllToJira")
    .addItem("Import Selected Rows to Jira", "importSelectedToJira")
    .addSeparator()
    .addItem("🧪 Test Connection",           "testConnection")
    .addSeparator()
    .addItem("⚙️ Setup My Jira Credentials", "setupCredentials")
    .addItem("🗑️ Clear My Credentials",      "clearCredentials")
    .addToUi();
}

// =============================================
// IMPORT ALL ROWS
// =============================================
function importAllToJira() {
  const creds = requireCredentials();
  if (!creds) return;

  const projectKey = askProjectKey();
  if (!projectKey) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  processRows(sheet, data, 1, data.length - 1, creds, projectKey);
}

// =============================================
// IMPORT SELECTED ROWS
// =============================================
function importSelectedToJira() {
  const creds = requireCredentials();
  if (!creds) return;

  const projectKey = askProjectKey();
  if (!projectKey) return;

  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const allData   = sheet.getDataRange().getValues();

  const from = Math.max(selection.getRow() - 1, 1);
  const to   = selection.getLastRow() - 1;

  processRows(sheet, allData, from, to, creds, projectKey);
}

// =============================================
// CORE PROCESSING LOGIC
// =============================================
function processRows(sheet, data, fromIndex, toIndex, creds, projectKey) {
  const ui    = SpreadsheetApp.getUi();
  const count = toIndex - fromIndex + 1;

  const confirm = ui.alert(
    "Import to Jira",
    `About to import ${count} issue(s) to project "${projectKey}".\n\nProceed?` + FOOTER,
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
  const log        = startImportLog(projectKey, runStarted);

  const results = { created: 0, warnings: 0, skipped: 0, errors: [] };

  for (let i = fromIndex; i <= toIndex; i++) {
    const row     = data[i];
    const rowNum  = i + 1;
    const summary = row[COL.SUMMARY] ? String(row[COL.SUMMARY]).trim() : "";

    // Skip empty rows
    if (!summary) {
      results.skipped++;
      continue;
    }

    const statusCell = sheet.getRange(rowNum, statusColIdx + 1);
    const notesCell  = sheet.getRange(rowNum, notesColIdx + 1);
    const statusValue = statusCell.getValue();

    // Skip already successfully imported rows
    if (statusValue && !String(statusValue).includes("ERROR")) {
      results.skipped++;
      Logger.log(`⏭️ Row ${rowNum}: already imported as ${statusValue}`);
      appendImportLog(log, rowNum, summary, "SKIPPED", String(statusValue), "Already imported");
      continue;
    }

    const outcome = createIssueWithFallback(row, i, creds, projectKey);

    if (outcome.key) {
      const url = `${JIRA_BASE_URL}/browse/${outcome.key}`;

      if (outcome.warning) {
        // Created, but with dropped fields — amber
        results.warnings++;
        statusCell.setValue(outcome.key).setBackground("#ffeb9c");
        writeNote(notesCell, `⚠️ ${outcome.warning}`, outcome.warning);
        Logger.log(`⚠️ Row ${rowNum}: Created ${outcome.key} (${outcome.warning}) — ${summary}`);
        appendImportLog(log, rowNum, summary, "CREATED (partial)", outcome.key, outcome.warning);
      } else {
        // Clean success — green
        results.created++;
        statusCell.setValue(outcome.key).setBackground("#c6efce");
        writeNote(notesCell, "✅ Imported", `Created ${outcome.key}\n${url}`);
        Logger.log(`✅ Row ${rowNum}: Created ${outcome.key} — ${summary}`);
        appendImportLog(log, rowNum, summary, "CREATED", outcome.key, "");
      }

    } else {
      // Failure — red. Store short reason in cell, full detail in a comment.
      results.errors.push(`Row ${rowNum} ("${summary}"): ${firstLine(outcome.error)}`);
      statusCell.setValue("ERROR").setBackground("#ffc7ce");

      const shortReason = errorShortLabel(outcome);
      const fullDetail  = buildErrorDetail(outcome);
      writeNote(notesCell, shortReason, fullDetail);

      Logger.log(`❌ Row ${rowNum}: ${outcome.error}`);
      appendImportLog(log, rowNum, summary, "ERROR", "", firstLine(outcome.error));
    }

    Utilities.sleep(300);
  }

  finishImportLog(log, results, new Date() - runStarted);

  // Summary dialog
  let msg =
    `✅ Created:   ${results.created}\n` +
    `⚠️ Partial:   ${results.warnings}   (created, some fields dropped)\n` +
    `⏭️ Skipped:  ${results.skipped}\n` +
    `❌ Errors:    ${results.errors.length}`;

  if (results.errors.length > 0) {
    msg += `\n\n❌ Error details:\n` + results.errors.slice(0, 8).join("\n");
    if (results.errors.length > 8) msg += `\n…and ${results.errors.length - 8} more (see "Import Log" tab)`;
  }
  msg += `\n\n📋 Full run recorded in the "Import Log" tab.`;

  ui.alert("Import Complete", msg + FOOTER, ui.ButtonSet.OK);
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
function startImportLog(projectKey, startedAt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Import Log");
  if (!sheet) {
    sheet = ss.insertSheet("Import Log");
  }
  sheet.clear();

  const stamp = Utilities.formatDate(startedAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.getRange(1, 1).setValue(`🚀 Jira Import — project ${projectKey} — started ${stamp}`)
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
    `Done — created ${results.created}, partial ${results.warnings}, ` +
    `skipped ${results.skipped}, errors ${results.errors.length}  (${secs}s)`;

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
      const payload  = buildIssuePayload(row, skipFields, projectKey);
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
// BUILD JIRA ISSUE PAYLOAD
// =============================================
function buildIssuePayload(row, skipFields, projectKey) {
  skipFields = skipFields || {};
  const skip = (fieldId) => skipFields[fieldId] === true;

  const fields = {
    project:   { key: projectKey },
    summary:   String(row[COL.SUMMARY]).trim(),
    issuetype: { name: row[COL.ISSUE_TYPE] ? String(row[COL.ISSUE_TYPE]).trim() : "Task" },
  };

  // Description
  if (!skip("description") && row[COL.DESCRIPTION]) {
    fields.description = toADF(String(row[COL.DESCRIPTION]));
  }

  // Acceptance Criteria
  if (!skip(CUSTOM_FIELDS.ACCEPTANCE_CRITERIA) && row[COL.ACCEPTANCE_CRITERIA]) {
    fields[CUSTOM_FIELDS.ACCEPTANCE_CRITERIA] = toADF(String(row[COL.ACCEPTANCE_CRITERIA]));
  }

  // Original Estimate (seconds)
  if (!skip(CUSTOM_FIELDS.ORIGINAL_ESTIMATE) && row[COL.ORIGINAL_ESTIMATE]) {
    fields[CUSTOM_FIELDS.ORIGINAL_ESTIMATE] = Number(row[COL.ORIGINAL_ESTIMATE]);
  }

  // Parent issue
  if (!skip("parent") && row[COL.PARENT]) {
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
