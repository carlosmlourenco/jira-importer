// =============================================
// CONFIGURATION — update these 3 values
// =============================================
const JIRA_BASE_URL   = "https://fdz.atlassian.net"; // e.g. https://acme.atlassian.net
const JIRA_EMAIL      = "carlos.lourenco@feedzai.com";
const JIRA_API_TOKEN  = "ATATT3xFfGF0FohehMoPn_D9xrXnxN7fLhyhGQrS3UIJNv1NqDyTElAT6dlKTXy8e1QrAtK2qSYYWkVphYJLVMrPtbDWLYagtc2LK0gUPCZbtlCM7dsDoHLCUi1rHt94pxA_Y7vFttca0ahyKXpnqxQSJS_GNMMHYj0g8LGPwYjTI31328gMWCM=3D6EC0EB";  // https://id.atlassian.com/manage-profile/security/api-tokens

const JIRA_PROJECT_KEY = "BACKBASE";

// =============================================
// CUSTOM FIELD IDs — from your BulkCreate config
// =============================================
const CUSTOM_FIELDS = {
  EPIC_LINK:           "customfield_10370",
  ORIGINAL_ESTIMATE:   "timeoriginalestimate",
  INITIATIVE:          "customfield_16455",
  NATURE_TYPE:         "customfield_12604",
  ACCEPTANCE_CRITERIA: "customfield_11371",
  CS_FUNCTIONAL_GROUP: "customfield_14391",
  CAPITALIZABLE:       "customfield_15670",
  CAP_BUCKET:          "customfield_16072",
  CATEGORY:            "customfield_12233",
  ACCOUNT:             "customfield_16433",
  PROJECT_PLAN:        "customfield_16491",
};

// =============================================
// COLUMN INDEX MAP (0-based, matches your CSV)
// =============================================
const COL = {
  EPIC_LINK:           0,
  ISSUE_TYPE:          1,
  // EPIC_NAME:        2,  // "do not map" — skipped
  SUMMARY:             3,
  // ALPHANUMERIC:     4,  // "do not map" — skipped
  ORIGINAL_ESTIMATE:   5,
  INITIATIVE:          6,
  NATURE_TYPE:         7,
  DESCRIPTION:         8,
  ACCEPTANCE_CRITERIA: 9,
  PRIORITY:            10,
  LABELS:              11,
  DUE_DATE:            12,
  CS_FUNCTIONAL_GROUP: 13,
  ACCOUNT:             14,
  CAPITALIZABLE:       15,
  CAP_BUCKET:          16,
  CATEGORY:            17,
  // PROJECT_PLAN:     18,  // handled separately
  PARENT:              19,
};

// =============================================
// FIELD VALUE MAPPINGS — option IDs from Jira
// =============================================

// CS Functional Group — array of {id}
const CS_FUNCTIONAL_GROUP_MAP = {
  "technical solutions": "15533",
  "data science":        "15534",
  "project management":  "15537",
  // Add more as you discover them:
  // "engagement manager": "XXXXX",
  // "risk consultant":    "XXXXX",
  // "csm":                "XXXXX",
  // "dt activation team": "XXXXX",
};

// Nature Type — single {id}
const NATURE_TYPE_MAP = {
  "new development": "13243",
  "technical debt":  "13245",
};

// Capitalizable — single {id}
const CAPITALIZABLE_MAP = {
  "no":  "17121",
  // "yes": "XXXXX",  // add when discovered
};

// Capitalization Bucket — single {id}
const CAP_BUCKET_MAP = {
  "n/a": "17901",
  // add more as discovered
};

// Project Plan Template — single {id}
// Your CSV value is {"id": 6, "value": "Standard billable"}
// Jira option IDs found: 18168 = Yes, 18169 = No
// Using id "6" as per your original CSV — update if Jira rejects it
const PROJECT_PLAN_VALUE = { id: "6" };

// =============================================
// MENU SETUP
// =============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Jira")
    .addItem("Import ALL Rows to Jira",      "importAllToJira")
    .addItem("Import Selected Rows to Jira", "importSelectedToJira")
    .addSeparator()
    .addItem("🧪 Test Connection",           "testConnection")
    .addToUi();
}

// =============================================
// IMPORT ALL ROWS
// =============================================
function importAllToJira() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  processRows(sheet, data, 1, data.length - 1);
}

// =============================================
// IMPORT SELECTED ROWS
// =============================================
function importSelectedToJira() {
  const sheet     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const allData   = sheet.getDataRange().getValues();

  const from = Math.max(selection.getRow() - 1, 1);
  const to   = selection.getLastRow() - 1;

  processRows(sheet, allData, from, to);
}

// =============================================
// CORE PROCESSING LOGIC
// =============================================
function processRows(sheet, data, fromIndex, toIndex) {
  const ui    = SpreadsheetApp.getUi();
  const count = toIndex - fromIndex + 1;

  const confirm = ui.alert(
    "Import to Jira",
    `About to import ${count} issue(s) to project "${JIRA_PROJECT_KEY}".\n\nProceed?`,
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

  const results = { created: 0, skipped: 0, errors: [] };

  for (let i = fromIndex; i <= toIndex; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[COL.SUMMARY] || String(row[COL.SUMMARY]).trim() === "") {
      results.skipped++;
      continue;
    }

    // Skip rows already successfully imported (has a Jira key like BACKBASE-123)
    const statusCell  = sheet.getRange(i + 1, statusColIdx + 1);
    const statusValue = statusCell.getValue();
    if (statusValue && !String(statusValue).includes("ERROR")) {
      results.skipped++;
      Logger.log(`⏭️ Row ${i + 1}: already imported as ${statusValue}`);
      continue;
    }

    const outcome = createIssueWithFallback(row, i);

    if (outcome.key) {
      results.created++;
      statusCell.setValue(outcome.key);
      statusCell.setBackground(outcome.warning ? "#ffeb9c" : "#c6efce"); // yellow = fallback, green = full
      Logger.log(`✅ Row ${i + 1}: Created ${outcome.key} — ${row[COL.SUMMARY]}${outcome.warning ? " (" + outcome.warning + ")" : ""}`);
    } else {
      results.errors.push(`Row ${i + 1} ("${row[COL.SUMMARY]}"): ${outcome.error}`);
      statusCell.setValue("ERROR").setBackground("#ffc7ce");
      Logger.log(`❌ Row ${i + 1}: ${outcome.error}`);
    }

    Utilities.sleep(300); // respect Jira rate limits
  }

  // Summary dialog
  let msg = `✅ Created:  ${results.created}\n⏭️ Skipped:  ${results.skipped}`;
  if (results.errors.length > 0) {
    msg += `\n\n❌ Errors (${results.errors.length}):\n` + results.errors.slice(0, 10).join("\n");
    if (results.errors.length > 10) msg += `\n...and ${results.errors.length - 10} more (check logs)`;
  }
  ui.alert("Import Complete", msg, ui.ButtonSet.OK);
}

// =============================================
// CREATE ISSUE WITH AUTO-FALLBACK
// Tries with estimate first, retries without if blocked
// =============================================
function createIssueWithFallback(row, rowIndex) {
  // First attempt — with original estimate
  try {
    const payload  = buildIssuePayload(row, false);
    const response = createJiraIssue(payload);
    return { key: response.key };
  } catch (e) {
    // If estimate field is not on screen, retry without it
    if (e.message.includes("timeoriginalestimate")) {
      Logger.log(`⚠️ Row ${rowIndex + 1}: estimate field blocked, retrying without it...`);
      try {
        const payload  = buildIssuePayload(row, true);
        const response = createJiraIssue(payload);
        return { key: response.key, warning: "estimate skipped" };
      } catch (e2) {
        return { error: e2.message };
      }
    }
    return { error: e.message };
  }
}

// =============================================
// BUILD JIRA ISSUE PAYLOAD
// =============================================
function buildIssuePayload(row, skipEstimate) {
  const fields = {
    project:   { key: JIRA_PROJECT_KEY },
    summary:   String(row[COL.SUMMARY]).trim(),
    issuetype: { name: String(row[COL.ISSUE_TYPE]).trim() },
    priority:  { name: row[COL.PRIORITY] ? String(row[COL.PRIORITY]).trim() : "Medium" },
  };

  // Description
  if (row[COL.DESCRIPTION]) {
    fields.description = toADF(String(row[COL.DESCRIPTION]));
  }

  // Acceptance Criteria
  if (row[COL.ACCEPTANCE_CRITERIA]) {
    fields[CUSTOM_FIELDS.ACCEPTANCE_CRITERIA] = toADF(String(row[COL.ACCEPTANCE_CRITERIA]));
  }

  // Original estimate (seconds) — skipped if field not on Jira screen
  if (!skipEstimate && row[COL.ORIGINAL_ESTIMATE]) {
    fields[CUSTOM_FIELDS.ORIGINAL_ESTIMATE] = Number(row[COL.ORIGINAL_ESTIMATE]);
  }

  // Epic link
  if (row[COL.EPIC_LINK]) {
    fields[CUSTOM_FIELDS.EPIC_LINK] = String(row[COL.EPIC_LINK]).trim();
  }

  // Parent issue
  if (row[COL.PARENT]) {
    fields.parent = { key: String(row[COL.PARENT]).trim() };
  }

  // Labels
  if (row[COL.LABELS]) {
    fields.labels = String(row[COL.LABELS]).split(/[,\s]+/).filter(Boolean);
  }

  // Due date → yyyy-MM-dd
  if (row[COL.DUE_DATE]) {
    const d = new Date(row[COL.DUE_DATE]);
    if (!isNaN(d)) {
      fields.duedate = Utilities.formatDate(d, "UTC", "yyyy-MM-dd");
    }
  }

  // Initiative — array of strings
  if (row[COL.INITIATIVE] && String(row[COL.INITIATIVE]).trim() !== "") {
    fields[CUSTOM_FIELDS.INITIATIVE] = [String(row[COL.INITIATIVE]).trim()];
  }

  // CS Functional Group — array of {id}
  if (row[COL.CS_FUNCTIONAL_GROUP] && String(row[COL.CS_FUNCTIONAL_GROUP]).trim() !== "") {
    const key = String(row[COL.CS_FUNCTIONAL_GROUP]).trim().toLowerCase();
    const id  = CS_FUNCTIONAL_GROUP_MAP[key];
    if (id) {
      fields[CUSTOM_FIELDS.CS_FUNCTIONAL_GROUP] = [{ id: id }];
    } else {
      Logger.log(`⚠️ Unknown CS Functional Group: "${row[COL.CS_FUNCTIONAL_GROUP]}" — skipping`);
    }
  }

  // Nature Type — single {id}
  if (row[COL.NATURE_TYPE] && String(row[COL.NATURE_TYPE]).trim() !== "") {
    const key = String(row[COL.NATURE_TYPE]).trim().toLowerCase();
    const id  = NATURE_TYPE_MAP[key];
    if (id) {
      fields[CUSTOM_FIELDS.NATURE_TYPE] = { id: id };
    } else {
      Logger.log(`⚠️ Unknown Nature Type: "${row[COL.NATURE_TYPE]}" — skipping`);
    }
  }

  // Capitalizable — single {id}
  if (row[COL.CAPITALIZABLE] && String(row[COL.CAPITALIZABLE]).trim() !== "") {
    const key = String(row[COL.CAPITALIZABLE]).trim().toLowerCase();
    const id  = CAPITALIZABLE_MAP[key];
    if (id) {
      fields[CUSTOM_FIELDS.CAPITALIZABLE] = { id: id };
    } else {
      Logger.log(`⚠️ Unknown Capitalizable: "${row[COL.CAPITALIZABLE]}" — skipping`);
    }
  }

  // Capitalization Bucket — single {id}
  if (row[COL.CAP_BUCKET] && String(row[COL.CAP_BUCKET]).trim() !== "") {
    const key = String(row[COL.CAP_BUCKET]).trim().toLowerCase();
    const id  = CAP_BUCKET_MAP[key];
    if (id) {
      fields[CUSTOM_FIELDS.CAP_BUCKET] = { id: id };
    } else {
      Logger.log(`⚠️ Unknown Cap Bucket: "${row[COL.CAP_BUCKET]}" — skipping`);
    }
  }

  // Category — plain string
  setField(fields, CUSTOM_FIELDS.CATEGORY, row[COL.CATEGORY]);

  // Account — parse JSON {"id": 6, "value": "Standard billable"} from column O
  if (row[COL.ACCOUNT] && String(row[COL.ACCOUNT]).trim() !== "") {
    try {
      const accountVal = typeof row[COL.ACCOUNT] === "string"
        ? JSON.parse(row[COL.ACCOUNT])
        : row[COL.ACCOUNT];
      fields[CUSTOM_FIELDS.ACCOUNT] = { id: String(accountVal.id), value: accountVal.value };
    } catch(e) {
      Logger.log(`⚠️ Could not parse Account value: ${row[COL.ACCOUNT]}`);
    }
  }

  // Project Plan Template — Yes/No field (column S, index 18)
  // 18168 = Yes, 18169 = No
  const projectPlanRaw = String(row[18] || "").trim().toLowerCase();
  if (projectPlanRaw === "yes") {
    fields[CUSTOM_FIELDS.PROJECT_PLAN] = { id: "18168" };
  } else if (projectPlanRaw === "no") {
    fields[CUSTOM_FIELDS.PROJECT_PLAN] = { id: "18169" };
  } else {
    Logger.log(`⚠️ Unknown Project Plan value: "${row[18]}" — skipping`);
  }

  return { fields };
}


// =============================================
// HELPER: Set field only if value is non-empty
// =============================================
function setField(fields, fieldId, value) {
  if (fieldId && value && String(value).trim() !== "") {
    fields[fieldId] = String(value).trim();
  }
}

// =============================================
// HELPER: Convert plain text to Atlassian Doc Format
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
function createJiraIssue(payload) {
  const url         = `${JIRA_BASE_URL}/rest/api/3/issue`;
  const credentials = Utilities.base64Encode(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`);

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

  const response = UrlFetchApp.fetch(url, options);
  const code     = response.getResponseCode();
  const body     = JSON.parse(response.getContentText());

  if (code !== 201) {
    const errors = body.errors ? JSON.stringify(body.errors) : JSON.stringify(body.errorMessages);
    throw new Error(`HTTP ${code}: ${errors}`);
  }

  return body;
}

// =============================================
// TEST CONNECTION
// =============================================
function testConnection() {
  const url         = `${JIRA_BASE_URL}/rest/api/3/myself`;
  const credentials = Utilities.base64Encode(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`);

  const response = UrlFetchApp.fetch(url, {
    method:  "get",
    headers: {
      "Authorization": `Basic ${credentials}`,
      "Accept":        "application/json"
    },
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const body = JSON.parse(response.getContentText());

  if (code === 200) {
    SpreadsheetApp.getUi().alert(
      "✅ Connection OK",
      `Connected as: ${body.displayName} (${body.emailAddress})`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "❌ Connection Failed",
      `HTTP ${code}: ${response.getContentText()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// =============================================
// DISCOVER FIELD VALUES — run to find option IDs
// =============================================
function discoverFieldValues() {
  const credentials = Utilities.base64Encode(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`);
  const headers     = { "Authorization": `Basic ${credentials}`, "Accept": "application/json" };

  const url  = `${JIRA_BASE_URL}/rest/api/3/search/jql`;
  const body = {
    jql:        `issueType = "Technical Task" ORDER BY created DESC`,
    maxResults: 50,
    fields:     ["summary", "customfield_14391", "customfield_16491", "customfield_12604", "customfield_16072", "customfield_15670"]
  };

  const response = UrlFetchApp.fetch(url, {
    method:             "post",
    contentType:        "application/json",
    headers:            headers,
    payload:            JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  const issues = result.issues || [];

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("_Field Options");
  if (!sheet) sheet = ss.insertSheet("_Field Options");
  sheet.clearContents();
  sheet.appendRow(["Issue Key", "Summary", "CS Func Group (14391)", "Project Plan (16491)", "Nature Type (12604)", "Cap Bucket (16072)", "Capitalizable (15670)"]);

  let found = 0;
  issues.forEach(function(issue) {
    const f = issue.fields;
    const hasData = f.customfield_14391 || f.customfield_16491 ||
                    f.customfield_12604  || f.customfield_16072 ||
                    f.customfield_15670;
    if (hasData) {
      found++;
      sheet.appendRow([
        issue.key,
        f.summary ? f.summary.substring(0, 50) : "",
        JSON.stringify(f.customfield_14391),
        JSON.stringify(f.customfield_16491),
        JSON.stringify(f.customfield_12604),
        JSON.stringify(f.customfield_16072),
        JSON.stringify(f.customfield_15670),
      ]);
    }
  });

  SpreadsheetApp.getUi().alert(
    "✅ Done",
    `Scanned ${issues.length} issues, found ${found} with field data.\nCheck '_Field Options' sheet.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

