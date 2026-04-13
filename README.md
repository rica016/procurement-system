<div align="center">

# 📊 DAR Procurement Monitoring System
### A web-based procurement tracking solution for the Department of Agrarian Reform

![Status](https://img.shields.io/badge/Status-Active-brightgreen)
![Backend](https://img.shields.io/badge/Backend-CodeIgniter-EF4223?logo=codeigniter)
![License](https://img.shields.io/badge/License-Internal%20Use-blue)

</div>

---

## 📌 Overview

The **DAR Procurement Monitoring System (PMS)** is a web-based platform built to streamline and monitor procurement transactions across multiple departments within the Department of Agrarian Reform (DAR). It centralizes transaction data, improves inter-department coordination, and ensures transparency throughout the procurement lifecycle.

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

### 📊 Google Sheets Integration
- Data stored and synced via structured Google Sheets
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
| Backend | CodeIgniter 4 |
| Frontend | HTML, CSS, JavaScript |
| Database | Google Sheets (via API) |
| Dev Tools | VS Code, Git |
| AI Assistance | Claude Code / GitHub Copilot |

---

## ⚙️ Installation & Setup

### Prerequisites
- PHP >= 8.1
- Composer
- Google Sheets API credentials

### Steps

**1. Clone the repository**
```bash
git clone https://github.com/your-username/dar-procurement-system.git
cd dar-procurement-system
```

**2. Install dependencies**
```bash
composer install
```

**3. Configure environment**
```bash
cp env .env
```
Update `.env` with your settings:
```env
app.baseURL = 'http://localhost:8080'
```

**4. Set up Google Sheets API**
- Go to [Google Cloud Console](https://console.cloud.google.com/)
- Enable the **Google Sheets API**
- Download your `credentials.json` and place it in the `/config` or `/writable` directory
- Update your config file with the target **Spreadsheet ID** and **Sheet name**

**5. Run the development server**
```bash
php spark serve
```

Visit `http://localhost:8080` in your browser.

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
