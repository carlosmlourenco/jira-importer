# Jira Field Reference

Custom fields available in the Jira instance, kept here for wiring up future
mappings in the importer. These were previously an inline comment block at the
top of `jiraimporter v2.js`.

| Field | ID | Type |
|-------|----|------|
| # of Approved Hours | `customfield_13750` | number |
| # of Requested Hours | `customfield_14083` | number |
| Total Remaining Hours | `customfield_20108` | number |
| Non-billable hours helper | `customfield_14471` | string |
| Approvers | `customfield_11905` | array |
| Pre-Approved | `customfield_12644` | option |
| Data Science Approver | `customfield_11786` | user |
| Engineering Approver | `customfield_11785` | user |
| Cloud Operations Approvers | `customfield_12580` | array |
| COPS Financial Approvers | `customfield_15471` | array |
| Approval from Groups | `customfield_16403` | array |
| Approved CSSM Depl. Request | `customfield_16434` | array |
| Approvals | `customfield_11870` | sd-approvals |
| Approved by | `customfield_11170` | user |
| Approved by (read only) | `customfield_11171` | string |
| Approved on | `customfield_11172` | datetime |
| Approved on (read only) | `customfield_11173` | string |
| Funding Proof of Approval | `customfield_22700` | string |
| Expedite Priority | `customfield_20340` | string |
| Expedite Approved By | `customfield_26408` | option |
| Approvers (migrated) | `customfield_20726` | array |
| Approver groups (migrated) | `customfield_20790` | array |

## Fields already wired into the importer

See `CUSTOM_FIELDS` / `COL` in `jiraimporter v3.1.js`:

| Field | ID |
|-------|----|
| Original Estimate | `timeoriginalestimate` (sent via `timetracking`) |
| Initiative | `customfield_16455` |
| Acceptance Criteria | `customfield_11371` |
| CS Functional Group | `customfield_14391` |
| Account | `customfield_16433` |
| # of Approved Hours | `customfield_13750` |
| Start Date | `customfield_16414` |
| Epic Name | `customfield_10371` (set to summary on Epic create) |
| Labels | `labels` (standard field; comma-separated cell, column M) |
