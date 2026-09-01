# Dynamic Group Travel Expense Splitter

A Google Apps Script-powered spreadsheet application that automatically splits travel expenses among group members, calculates fair settlements, and sends personalized PDF expense reports via email.

## Overview

This tool solves the common problem of tracking and settling shared expenses during group travel. Instead of keeping receipts and arguing over calculations at the end of the trip, everyone can:

- Enter expenses in real-time as they happen
- See who paid what and who owes whom
- Get a personalized breakdown of their own expenses
- Receive PDF reports via email for their records

---

## Features

### Core Functionality

| Feature | Description |
|---------|-------------|
| **Dynamic Roster** | Supports 2-15 people per trip |
| **Flexible Splitting** | Split equally among everyone or specific people |
| **Auto Calculations** | Per-person shares and net balances update automatically |
| **Settlement Optimization** | Greedy algorithm minimizes the number of transactions needed to settle |
| **PDF Reports** | Personalized email reports with itemized expenses per person |
| **Mobile Support** | Works on Google Sheets mobile app via Setup sheet controls |

### Sheet Structure

The script automatically creates three sheets:

#### 1. Setup Sheet
- **Trip Name / Destination**: Name your trip (e.g., "Goa Vacation 2026")
- **Number of People**: Dropdown to select 2-15 people
- **Currency Symbol**: Customize for ₹, $, €, £, etc.
- **Setup Status**: Auto-validates configuration
- **All Expenses Confirmed**: Master checkbox (mobile-friendly)
- **Roster Table**: Names and optional emails for each participant

#### 2. Expenses Sheet
| Column | Purpose |
|--------|---------|
| Date | When the expense occurred |
| Description | What was purchased |
| Amount | Total cost |
| Paid By | Who paid (dropdown from roster) |
| Split Among | Who shares this expense (leave blank for all) |
| Person Columns | Auto-calculated net balance per person |

#### 3. Send Reports Sheet
- **Personal Balances Summary**: Each person's total paid, share, and net balance
- **Simplified Settlement Plan**: Minimal transactions needed to settle all debts
- **Report Status**: Tracks email delivery status

---

## Installation Guide

### Prerequisites

- A Google account with access to Google Sheets
- Internet connection
- Basic familiarity with Google Sheets

### Step 1: Create a New Google Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Click **Blank** to create a new spreadsheet
3. Give it a meaningful name (e.g., "Trip Expenses")

### Step 2: Open the Apps Script Editor

1. In your new spreadsheet, click **Extensions** → **Apps Script**
2. This opens the Apps Script editor in a new tab
3. Delete any existing code in the editor

### Step 3: Install the Script

1. Copy the entire contents of `main_optimized.gs` (recommended) or `main.gs`
2. Paste it into the Apps Script editor
3. Click the **Save** icon (💾) or press `Ctrl + S`

### Step 4: Authorize the Script

The first time you run the script, it needs permission to access your spreadsheet and send emails:

1. In the Apps Script editor, select **`setup`** from the function dropdown
2. Click the **Run** button (▶️)
3. Google will show an authorization prompt:
   - Click **Continue**
   - Select your Google account
   - Click **Advanced** → **Go to [project name] (unsafe)**
   - Click **Allow**
4. The `setup()` function will create the three sheets and configure everything

### Step 5: Reload the Spreadsheet

1. Go back to your Google Spreadsheet
2. Refresh the page (press `F5` or `Ctrl + R`)
3. A new menu **"Expense Splitter"** will appear in the toolbar

---

## User Guide

### Initial Setup (One Time)

1. Click **Expense Splitter** menu → **1. Initialize / Reset Tracker**
2. Edit the **Setup** sheet:
   - Set your **Trip Name**
   - Verify **Number of People** (adjust if needed)
   - Set your **Currency Symbol**
   - Enter each person's **Name** and **Email** (email is optional but needed for PDF reports)

### Recording Expenses

1. Go to the **Expenses** sheet
2. For each expense, fill in:
   - **Date**: Click to pick a date
   - **Description**: What was bought (e.g., "Dinner at Beach Restaurant")
   - **Amount**: The total cost
   - **Paid By**: Select from the dropdown
   - **Split Among**: Leave blank for everyone, or type specific names (comma-separated)

### Confirming Expenses

**Why is this step needed?**
The confirm checkbox prevents accidental edits to settled expenses and hides calculations until you're ready.

**On Mobile:**
1. Go to the **Setup** sheet
2. Check the **All Expenses Confirmed** checkbox (B6)

**On Desktop:**
1. Click **Expense Splitter** menu
2. Select **Confirm All Expenses**

Once confirmed:
- Per-person balance columns appear on the Expenses sheet
- A TOTAL row is added
- The Send Reports sheet shows balances and settlements

**To make changes:**
- Click **Expense Splitter** menu → **Re-open for Editing (Unconfirm)**

### Sending Reports

1. Make sure all expenses are **confirmed**
2. Click **Expense Splitter** menu → **3. Send PDF Reports via Email**
3. Each participant with an email address receives:
   - Their total paid vs. share
   - Their net balance (positive = owed money, negative = owes money)
   - Itemized list of relevant expenses
   - Their specific settlement instructions

---

## Menu Reference

| Menu Item | Function | Description |
|-----------|----------|-------------|
| 1. Initialize / Reset Tracker | `setup()` | Creates/recreates all sheets and configuration |
| 2. Recalculate Expenses & Settlements | `recalculateAll()` | Force recalculate all values |
| Confirm All Expenses | `confirmAllExpenses()` | Lock expenses and show totals |
| Re-open for Editing (Unconfirm) | `unconfirmExpenses()` | Unlock expenses for editing |
| 3. Send PDF Reports via Email | `sendPdfReports()` | Email PDF reports to all participants |
| Test Email Permission | `testEmailPermission()` | Verify email sending works |
| Fix Expenses Sheet Formatting | `fixExpensesSheet()` | Reset sheet if formatting gets corrupted |
| Diagnose | `diagnose()` | Health check for troubleshooting |

---

## How the Calculations Work

### Equal Split Algorithm

When an expense is split among N people:
```
baseShare = floor(amount / N * 100) / 100
remainder = round((amount - baseShare * N) * 100) / 100
```

The remainder (from rounding) is given to the person who paid.

**Example:** ₹100 split among 3 people
- Base share: ₹33.33
- Remainder: ₹0.01
- Payer pays ₹33.34, others pay ₹33.33 each

### Settlement Optimization

The greedy algorithm minimizes transactions by:
1. Identifying who is owed money (creditors)
2. Identifying who owes money (debtors)
3. Matching largest debtors with largest creditors
4. Repeating until all debts are settled

**Example:** A owes ₹100, B owes ₹50, C is owed ₹150
- A → C: ₹100
- B → C: ₹50
- Total: 2 transactions instead of 3

---

## Split Among Options

| Value | Meaning |
|-------|---------|
| *(blank)* | Split equally among ALL participants |
| `All` | Same as blank - split among everyone |
| `Rahul, Priya` | Split only between Rahul and Priya |
| `Rahul` | Rahul pays the full amount (e.g., for a personal item) |

---

## Troubleshooting

### Menu Doesn't Appear
- Make sure you've reloaded the spreadsheet after running `setup()`
- Check if you're using Google Sheets mobile app (menus don't appear on mobile)

### "Fix Expenses Sheet Formatting" Resets My Data
- Yes, this function clears all expense data intentionally
- Use it when sheet formatting gets corrupted or you want a fresh start

### Email Permission Denied
- Click **Expense Splitter** → **Test Email Permission**
- Make sure you're logged into the correct Google account
- Check spam folder

### Calculations Look Wrong
- Ensure **All Expenses Confirmed** is checked
- Verify "Paid By" and "Split Among" are spelled exactly as names appear in Setup

### People Columns Are Missing
- Make sure all names are filled in the Setup roster
- Run **Diagnose** from the menu to check sheet state

---

## File Structure

```
travel-expense-excel/
├── main.gs              # Original script version
├── main_optimized.gs     # Optimized version (recommended)
├── README.md            # This file
└── .gitignore           # Git ignore file
```

---

## Technical Notes

- **Maximum data rows**: 500 expense entries
- **Maximum people**: 15 (limited by column space in Expenses sheet)
- **Currency**: Configurable symbol stored in Setup!B4
- **Email limits**: Google limits ~100 emails/day for free accounts
- **PDF generation**: Uses HTML to PDF conversion via `HtmlService`

---

## Credits

Developed for group travel expense management. Uses Google Apps Script with:
- Spreadsheet API for data management
- MailApp for email delivery
- HtmlService for PDF generation
