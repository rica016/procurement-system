# Clasp Cheatsheet (DARPMSv2)

Quick command reference for this Apps Script project.

## 1) Authentication

```bash
clasp login
clasp logout
```

If you manage multiple Google accounts, verify active auth before deploy.

## 2) Project Linking and Inspection

```bash
clasp clone <SCRIPT_ID>
clasp open-script
clasp open
clasp status
```

- clasp open-script: opens Apps Script editor
- clasp open: opens project in Google Drive
- clasp status: shows changed local files

## 3) Daily Development Flow

```bash
clasp pull
# edit locally
clasp push
```

Optional auto-push while editing:

```bash
clasp push --watch
```

## 4) Versioning and Deployment

Create immutable version first, then deploy/update web app:

```bash
clasp version "Describe release"
clasp deploy --description "Web app release"
clasp deployments
clasp undeploy <DEPLOYMENT_ID>
```

Update an existing deployment:

```bash
clasp deploy --deploymentId <DEPLOYMENT_ID> --description "Update existing web app"
```

## 5) Logs and Troubleshooting

```bash
clasp logs
```

Useful checks:

```bash
clasp status
clasp deployments
```

## 6) Google Cloud Project Notes (Important)

clasp commands work against the Apps Script project, but OAuth and API settings are controlled by the linked Google Cloud project.

Checklist:

1. Apps Script Project Settings -> Google Cloud Platform project is linked to your standard GCP project.
2. In Google Cloud Console, enable Apps Script API.
3. Configure OAuth consent screen (Internal or External based on audience).
4. Add test users if app is still in testing mode.
5. Redeploy web app after consent/scope changes.

## 7) Recommended Release Sequence

```bash
git pull
clasp pull
# edit and test
git add .
git commit -m "Describe change"
git push
clasp push
clasp version "Release note"
clasp deploy --description "Production update"
```