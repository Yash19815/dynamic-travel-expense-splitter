<div align="center">

# ✈️ Dynamic Group Travel Expense Splitter

**A zero-dependency, open-source Splitwise alternative built directly inside Google Sheets.**  
Automated per-person matrix calculations, greedy debt minimization, and 1-click personalized PDF email reports.

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)](#)
[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white)](#)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](LICENSE)
[![PRs Welcome](https://img.shields.io/badge/PRs-Welcome-brightgreen.svg?style=for-the-badge)](CONTRIBUTING.md)

<br />

<!-- ================================================================= -->
<!-- SCREENSHOT PLACEHOLDER 1: HERO DEMO GIF -->
<!-- Replace 'assets/hero-demo.gif' with your recorded workflow GIF -->
<!-- ================================================================= -->
<a href="#-quick-start">
  <img src="assets/hero-demo.gif" alt="Expense Splitter Interactive Demo" width="850" />
</a>

<br /><br />

[**🚀 Click Here to Make a Copy in Google Drive**](https://docs.google.com/spreadsheets/d/11dY2B2j1dNIRlWfagUtbX13LRmW894mhnhCGBVsx9oQ/copy)  
*(No installation required for non-technical users)*

</div>

---

## 🌟 Why Use This Over Splitwise?

* **100% Free & Self-Hosted:** No subscription limits, item limits, or ads. Your trip data stays inside your Google Drive.
* **Greedy Settlement Engine:** Automatically computes the mathematically optimal minimum number of transactions to clear group debts.
* **Mobile-Friendly Confirmation:** Enter raw expenses on the mobile app, then tick a single **"All Expenses Confirmed"** checkbox to generate real-time balances.
* **Automated PDF Email Statements:** Generates custom, cleanly formatted HTML/PDF breakdown reports and dispatches them via Gmail to every group member in one click.
* **Dynamic Roster Scaling:** Scale your group seamlessly from 2 to 15 people—the grid, headers, and validations auto-adjust on the fly.

---

## 📸 Screenshots & Workflow

<table width="100%">
  <tr>
    <td width="50%" align="center">
      <b>1. Trip & Roster Setup</b><br /><br />
      <!-- SCREENSHOT PLACEHOLDER 2: SETUP TAB -->
      <img src="assets/setup-tab.png" alt="Setup Tab" width="100%" />
    </td>
    <td width="50%" align="center">
      <b>2. Smart Expense Tracker</b><br /><br />
      <!-- SCREENSHOT PLACEHOLDER 3: EXPENSES TAB -->
      <img src="assets/expenses-tab.png" alt="Expenses Tab" width="100%" />
    </td>
  </tr>
  <tr>
    <td width="50%" align="center">
      <b>3. Simplified Settlements</b><br /><br />
      <!-- SCREENSHOT PLACEHOLDER 4: SETTLEMENT PLAN -->
      <img src="assets/settlements-tab.png" alt="Settlement Plan" width="100%" />
    </td>
    <td width="50%" align="center">
      <b>4. Generated Email PDF Report</b><br /><br />
      <!-- SCREENSHOT PLACEHOLDER 5: PDF EMAIL REPORT -->
      <img src="assets/pdf-report.png" alt="PDF Report Sample" width="100%" />
    </td>
  </tr>
</table>

---

## ⚡ Quick Start

### Option A: 1-Click Google Sheet Template (Easiest)
1. Click the **[Make a Copy Template Link](https://docs.google.com/spreadsheets/d/11dY2B2j1dNIRlWfagUtbX13LRmW894mhnhCGBVsx9oQ/copy)**.
2. Open the newly created spreadsheet in your Google Drive.
3. Open the custom **Expense Splitter** menu at the top and select `1. Initialize / Reset Tracker`.
4. Authorize the Google Apps Script permissions when prompted.

---

### Option B: Manual Setup via Google Apps Script
1. Create a blank **Google Sheet**.
2. Navigate to **Extensions** -> **Apps Script**.
3. Replace all content in `Code.gs` with the code from [`main_optimized.gs`](main_optimized.gs).
4. Save the project and click **Run** on the `setup` function to authorize scopes.
5. Reload your Google Sheet—the custom **Expense Splitter** menu will appear in the top navigation toolbar.

---

⚙️ How It Works
1. Equal & Partial Splits
In the Expenses tab, enter the expenditure details:
 * "All" or Blank: Splits the cost equally among all roster members.
 * Comma-Separated Names (Rahul, Priya): Splits the cost exclusively among the specified subgroup.
 * Remainder Handling: Fractional cents/paisa remainders from division are mathematically allocated to the payer to prevent rounding balance drift.
2. Minimum Transaction Engine (Greedy Algorithm)
Instead of N \times (N-1) individual payments, the system matches net creditors and net debtors to minimize settlement transactions:
```
[ Debtor List (Sorted Desc) ] <---> [ Creditor List (Sorted Desc) ]
                 |                                    |
                 +-----> Pay Minimum Share Balance ---+
```
📂 Project Architecture
```
dynamic-travel-expense-splitter/
├── .clasp.json               # Google Apps Script Clasp CLI config
├── appsscript.json           # Apps Script manifest & scope authorizations
├── main_optimized.gs         # Production-ready Apps Script engine
├── README.md                 # Project documentation
├── LICENSE                   # MIT License
└── assets/                   # Screenshots, diagrams, and hero GIFs
    ├── hero-demo.gif
    ├── setup-tab.png
    ├── expenses-tab.png
    ├── settlements-tab.png
    └── pdf-report.png
```
🛠️ Configuration Options
```
| Setting Cell | Description | Default Value |
| Setup!B2 | Trip Name / Destination | Goa Vacation 2026 |
| Setup!B3 | Number of Participants (Dropdown 2–15) | 4 |
| Setup!B4 | Currency Symbol | ₹ |
| Setup!B6 | Master Expense Confirmation Toggle | FALSE |
```
🤝 Contributing
Contributions are welcome! Please feel free to open an issue or submit a pull request:
 * Fork the Repository.
 * Create your Feature Branch (git checkout -b feature/AmazingFeature).
 * Commit your Changes (git commit -m 'Add some AmazingFeature').
 * Push to the Branch (git push origin feature/AmazingFeature).
 * Open a Pull Request.
📜 License
Distributed under the MIT License. See LICENSE for more information.
