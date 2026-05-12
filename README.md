# DAR Procurement Monitoring System

Web-based procurement workflow tracker for DAR, built on Google Apps Script + Google Sheets.

## Overview

This system tracks procurement requests end-to-end across BAC, Supply, Budget, Accounting, Cash, RCAO, ARDA, and End User touchpoints. It stores transactional data in Google Sheets and serves the UI through Apps Script HtmlService.

## Current Features

### Workflow and Tracking
- Central transaction list keyed by TRACKING NO.
- Department page data with completed-copy behavior for previously processed departments.
- Return-aware workflow: receive, return, continue returned, receive-and-forward, and complete flows.
- Timeline/search view that derives effective section and status from history.

### Security and Access
- Google Workspace identity check via Session.getActiveUser().getEmail().
- Multi-role login handling with role selection when one email has multiple active roles.
- Login rate limiting:
  - Max attempts: 5
  - Lockout duration: 15 minutes
  - Attempt window: 30 minutes
- PIN lifecycle:
  - Per-user 6-digit PIN setup
  - SHA-256 + salt pin hashing (hex storage)
  - PIN verification and admin reset support

### Admin and Data Maintenance
- User management (active/inactive, role, end-user flag, PIN reset).
- Supplier and end-user master list management.
- Audit log write/read/delete operations.
- Auto-create/repair core sheets and required headers.

## Core Data Sheets

The app ensures these sheets exist:
- TRANSACTIONS
- TRANSACTION_HISTORY
- USERS
- SUPPLIERS
- END_USERS
- AUDIT_LOGS

## Deployment Guide

### Prerequisites

1. Google account with access to the target spreadsheet.
2. A Google Sheet that will be used as the datastore.
3. Apps Script project bound to that sheet, or standalone script with the correct Spreadsheet ID.
4. Node.js + npm if you will use clasp locally.

### A. Bind and Configure the Script

1. Open your spreadsheet.
2. Go to Extensions > Apps Script.
3. Paste/upload the project files.
4. Confirm SPREADSHEET_ID in Code.js points to your active spreadsheet.
5. Run setupSheets once from the Apps Script editor to initialize missing tabs/headers.

### B. Link Your Google Cloud Console Project (Standard Project)

Since you already created a Google Cloud project, link it to this Apps Script project:

1. In Apps Script editor, open Project Settings.
2. Under Google Cloud Platform (GCP) Project, click Change project.
3. Enter your GCP Project Number and link it.
4. In Google Cloud Console for that project:
   - Enable APIs used by your deployment (at minimum Apps Script API for clasp operations).
   - Configure OAuth consent screen if users outside your internal workspace will access it.
   - Add authorized test users if app is in testing mode.
5. Return to Apps Script and verify the linked project is shown in settings.

Notes:
- For internal DAR use on Workspace, Internal user type is typically enough.
- If you deploy to broader audiences, complete consent screen and verification requirements as needed.

### C. Deploy as Web App (Apps Script UI)

1. Click Deploy > New deployment.
2. Type: Web app.
3. Execute as: User accessing the web app (current manifest setting) or Me, based on your security model.
4. Who has access: choose appropriate audience (domain/internal/public as required).
5. Deploy, authorize scopes, then copy the web app URL.

If you update code, use Manage deployments to edit and redeploy.

### D. Deploy from Local with clasp

1. Install clasp:

```bash
npm install -g @google/clasp
```

2. Login:

```bash
clasp login
```

3. Clone or set script:

```bash
clasp clone <SCRIPT_ID>
```

4. Push local updates:

```bash
clasp push
```

5. Create a version and deploy:

```bash
clasp version "Docs/feature update"
clasp deploy --description "Web app update"
```

## Runtime Notes

- Runtime: V8
- Timezone: Asia/Manila
- Web app manifest defaults:
  - access: ANYONE
  - executeAs: USER_DEPLOYING

Review these values in appsscript.json before production release to ensure they match DAR policy.

## Operational Checklist

Before go-live:

1. Confirm USERS sheet has active accounts and roles.
2. Confirm each required sheet exists with headers.
3. Test login with a single-role account and a multi-role account.
4. Test PIN set, verify, and admin reset.
5. Test forward, return, receive, and complete paths.
6. Confirm audit logs are written for key actions.

## License

Internal DAR use only.
