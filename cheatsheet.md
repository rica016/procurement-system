# 🗂️ Clasp Cheatsheet

> Google Apps Script CLI — quick reference for VS Code

---

## 🔐 Authentication

```bash
clasp login     # Login to Google account
clasp logout    # Logout
```

---

## 📂 Project Setup

```bash
clasp create --title "My Project"   # Create new Apps Script project
clasp clone <SCRIPT_ID>             # Clone existing project
```

---

## 📁 Open Project

```bash
clasp open-script   # Open Apps Script in browser
```

---

## 🚀 Deployment

```bash
clasp deploy                # Create a new deployment
clasp deployments           # List all deployments
clasp undeploy <DEPLOY_ID>  # Remove a deployment
```

---

## 📜 Logs

```bash
clasp logs   # View execution logs
```

---

## ⚙️ Config

```bash
clasp status   # Check file changes
clasp version  # Create a new version
```

---

## 🧠 Common Workflow (Daily Use)

```bash
clasp pull        # Sync latest from remote
# edit code in VS Code
clasp push        # Upload local changes
clasp open-script # Open in browser if needed
```

---

> 💡 **Tip:** Use `clasp push --watch` to auto-push every time you save a file in VS Code.