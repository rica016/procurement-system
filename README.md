<div align="center">

# 📊 DAR Procurement Monitoring System
### A web-based procurement tracking solution for the Department of Agrarian Reform

![Status](https://img.shields.io/badge/Status-Active-brightgreen)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-4285F4?logo=google)
![Database](https://img.shields.io/badge/Database-Google%20Sheets-34A853?logo=googlesheets)
![License](https://img.shields.io/badge/License-Internal%20Use-blue)

</div>

---

## 📌 Overview

The **DAR Procurement Monitoring System (PMS)** is a web-based platform built entirely on **Google Apps Script**, designed to streamline and monitor procurement transactions across multiple departments within the Department of Agrarian Reform (DAR). It centralizes transaction data, improves inter-department coordination, and ensures transparency throughout the procurement lifecycle — all without requiring external hosting or a separate database.

---

## 🚀 Features

### 📄 Transaction Management
- Create, update, and track procurement transactions in real time
- View all department-related data from a unified dashboard

### 🔄 Workflow Tracking
Monitor transaction movement across the following departments:

| Department | Role |
|---|---|
| BAC | Bids and Awards Committee — initiates and manages bidding |
| Supply | Handles supply-related processing |
| Budget | Budget allocation and clearance |
| Accounting | Financial review and recording |
| Cashier | Final payment processing |

**Transaction Statuses:**
- `New` — Freshly submitted transaction
- `Active / In Progress` — Currently being processed
- `Returned` — Sent back to originating department
- `Completed` — Fully processed and closed
- `Cancelled` — Voided transaction

### 🔁 Return Handling
- Send transactions back to originating departments with detailed remarks
- Tracks: **Returned By**, **Return Remarks**, and **Forward Remarks**

### 📝 BAC Resolution Tracking
Manage BAC Resolution for Award with the following statuses:

| Status | Description |
|---|---|
| Draft | Initial preparation stage |
| For Signature | Awaiting signatories |
| Partially Signed | Some signatures obtained |
| Fully Signed | All signatures complete |
| Approved | Officially approved |
| Returned | Sent back for revision |

### 🔐 User & Role Management
- Admin-controlled user access
- Department-based role assignments

### 📊 Google Sheets as Database
- All data stored directly in Google Sheets
- No external database or server required
- Column ranges mapped per department (see [System Structure](#️-system-structure))

---

## 🏗️ System Structure

### Department-to-Column Mapping

| Department | Column Range | Purpose |
|---|---|---|
| BAC | A – P | Core procurement data |
| Supply | Q – AA | Supply chain details |
| Budget | AB – AD | Budget allocation info |
| Accounting | AE – AG | Financial records |
| Cashier | AH – AI | Payment processing |

---

## 🔄 Procurement Workflow

```
[New] ──► [Active / In Progress] ──► [Completed / Cancelled]
                  │                           ▲
                  ▼                           │
             [Returned] ──► [Received] ──► [Active]
```

- A returned transaction includes **Return Remarks** from the sending department
- Once acknowledged, it re-enters the **Active / In Progress** state
- Forwarding includes **Forward Remarks** for the receiving department

---

## 🖥️ Tech Stack

| Layer | Technology |
|---|---|
| Backend | Google Apps Script |
| Frontend | HTML, CSS, JavaScript (served via GAS `HtmlService`) |
| Database | Google Sheets |
| Hosting | Google Apps Script Web App (no external server needed) |
| Dev Tools | Apps Script IDE / Clasp (CLI), Git |
| AI Assistance | Claude Code / GitHub Copilot |

---

## ⚙️ Setup & Deployment

### Prerequisites
- A **Google Account** with access to Google Drive
- The target **Google Sheets** file set up with the correct column structure
- *(Optional)* [Clasp](https://github.com/google/clasp) for local development

---

### Option A — Direct via Apps Script IDE

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Paste or upload the project source files
4. Click **Deploy → New Deployment**
   - Type: **Web App**
   - Execute as: `Me`
   - Who has access: `Anyone within [your organization]` *(or as needed)*
5. Click **Deploy** and copy the Web App URL
6. Share the URL with your department users

---

### Option B — Using Clasp (Local Development)

**1. Install Clasp**
```bash
npm install -g @google/clasp
```

**2. Log in to your Google account**
```bash
clasp login
```

**3. Clone the Apps Script project**
```bash
clasp clone <your-script-id>
```
> Find your Script ID under **Project Settings** in the Apps Script IDE.

**4. Make changes locally, then push**
```bash
clasp push
```

**5. Deploy as Web App**
```bash
clasp deploy --description "v1.0 Initial Release"
```

---

### 🔑 Configuration

In your `Code.gs` (or equivalent config file), update the following:

```javascript
const SPREADSHEET_ID = 'your-google-sheet-id-here';
const SHEET_NAME = 'Transactions'; // or your actual sheet name
```

---

## 👤 User Roles

| Role | Access Level |
|---|---|
| **Admin** | Full access to all departments, users, and data |
| **Department User** | Limited to assigned department; can update statuses and remarks |

---

## 📌 Terminology / Naming Conventions

| Term | Definition |
|---|---|
| **Returned By** | The department that sent the transaction back |
| **Return Remarks** | Reason or notes for the return |
| **Forward Remarks** | Notes added when forwarding to the next department |
| **Status** | Current state of the transaction in the workflow |

---

## 📈 Planned Improvements

- [ ] 📧 Email notifications per department upon status change
- [ ] 📊 Dashboard with analytics and charts
- [ ] 🔑 Google OAuth-based authentication
- [ ] 🗂️ Audit logs for all transaction changes
- [ ] 📱 Mobile-responsive UI improvements

---

## 🤝 Contributing

Contributions are welcome for authorized collaborators.

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Commit your changes: `git commit -m "Add: your feature description"`
4. Push to the branch: `git push origin feature/your-feature-name`
5. Open a Pull Request

---

## 📄 License

This project is developed for **internal use within the Department of Agrarian Reform (DAR)**. Unauthorized distribution or use outside DAR is not permitted.

---

<div align="center">
  Built with ❤️ for the Department of Agrarian Reform
</div>
