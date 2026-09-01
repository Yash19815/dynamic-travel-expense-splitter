/**
 * DYNAMIC GROUP TRAVEL EXPENSE SPLITTER
 * ------------------------------------------------------------------
 * Custom Menu: "Expense Splitter" (PC only - menus do NOT appear in
 * the Google Sheets mobile app. On mobile, use the
 * "All Expenses Confirmed" checkbox on the Setup sheet.)
 *
 * INSTALL:
 *   1. Extensions -> Apps Script -> paste this file -> Save
 *   2. Run setup() once and authorize ALL requested scopes
 *      (including "Send email as you" - required for reports)
 *   3. Reload the spreadsheet to see the "Expense Splitter" menu
 *
 * CONFIRM FLOW (single switch, not per-row):
 *   - Enter all expense rows on "Expenses" (no totals shown yet)
 *   - When done: tick Setup -> "All Expenses Confirmed" (B6)
 *     OR PC menu: Expense Splitter -> Confirm All Expenses
 *     OR assign confirmAllExpenses to an Insert -> Drawing button
 *   - Per-person columns + TOTAL row appear on the Expenses tab
 *   - Untick / "Re-open for Editing" hides totals again
 *
 * NOTE: the TOTAL label lives in the Description column because
 * column A has strict date validation which also blocks script writes.
 *
 * TROUBLESHOOTING:
 *   - "Fix Expenses Sheet Formatting" wipes stray formatting/validation
 *   - "Diagnose" reports the state of every moving part
 */

const SHEET_SETUP = "Setup";
const SHEET_EXPENSES = "Expenses";
const SHEET_REPORTS = "Send Reports";

const COL = { DATE: 1, DESC: 2, AMOUNT: 3, PAID_BY: 4, SPLIT: 5 };
const FIRST_PERSON_COL = 6; // person share columns start at column F
const MAX_DATA_ROWS = 500;
const CONFIRM_CELL = "B6"; // Setup sheet master confirm checkbox
const TOTAL_LABEL = "TOTAL"; // written to Description col (B), never Date col (A)

const COLOR = {
  HEADER: "#0F766E",
  HEADER_DARK: "#1E293B",
  HEADER_MID: "#334155",
  TOTAL_BG: "#E2E8F0",
  GREEN_BG: "#D1FAE5", GREEN_TX: "#065F46",
  RED_BG: "#FEE2E2",   RED_TX: "#991B1B",
  GRAY_BG: "#F3F4F6",  GRAY_TX: "#374151",
  WARN_TX: "#B45309",
  PENDING: "#FEF3C7"
};

/**
 * UI-safe alert: getUi() throws without an open spreadsheet UI.
 * Falls back to a toast, then to the execution log.
 */
function safeAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
    return;
  } catch (e) { /* no UI context — fall through */ }
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Expense Splitter", 10);
  } catch (e2) {
    Logger.log(message);
  }
}

/**
 * Creates custom menu when the spreadsheet opens (PC / desktop only).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Expense Splitter')
    .addItem('1. Initialize / Reset Tracker', 'setup')
    .addItem('2. Recalculate Expenses & Settlements', 'recalculateAll')
    .addSeparator()
    .addItem('Confirm All Expenses', 'confirmAllExpenses')
    .addItem('Re-open for Editing (Unconfirm)', 'unconfirmExpenses')
    .addSeparator()
    .addItem('3. Send PDF Reports via Email', 'sendPdfReports')
    .addItem('Test Email Permission', 'testEmailPermission')
    .addSeparator()
    .addItem('Fix Expenses Sheet Formatting', 'fixExpensesSheet')
    .addItem('Diagnose', 'diagnose')
    .addToUi();
}

/**
 * Health check: verifies the state of every piece and reports back.
 */
function diagnose() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = [];

  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  lines.push(`Setup sheet: ${setupSheet ? "found" : "MISSING"}`);
  lines.push(`Expenses sheet: ${expenseSheet ? "found" : "MISSING"}`);

  if (setupSheet) {
    lines.push(`Confirmed checkbox (B6): ${setupSheet.getRange(CONFIRM_CELL).getValue()}`);
    lines.push(`Setup status: ${setupSheet.getRange("B5").getValue()}`);
  }

  const roster = getRoster();
  lines.push(`Roster names: ${roster.length ? roster.map(r => r.name).join(", ") : "NONE"}`);

  if (expenseSheet) {
    lines.push(`Expense data rows: ${countDataRows_(expenseSheet)}`);

    if (expenseSheet.getLastRow() > 1) {
      const r = expenseSheet.getRange(2, 1, 1, 5).getValues()[0];
      lines.push(`Row 2 — Amount: "${r[2]}", Paid By: "${r[3]}", Split Among: "${r[4]}"`);
    }

    // Write test: can the script write to the person columns?
    try {
      const testCell = expenseSheet.getRange(MAX_DATA_ROWS, FIRST_PERSON_COL);
      const old = testCell.getValue();
      testCell.setValue("__test__");
      testCell.setValue(old);
      lines.push("Write test on person columns: OK");
    } catch (err) {
      lines.push(`Write test on person columns: FAILED — ${err.message}`);
    }
  }

  safeAlert("DIAGNOSIS:\n\n" + lines.join("\n"));
}

/**
 * Master initialization function.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const setupSheet = ss.getSheetByName(SHEET_SETUP) || ss.insertSheet(SHEET_SETUP);
  ss.getSheetByName(SHEET_EXPENSES) || ss.insertSheet(SHEET_EXPENSES);
  ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);

  // --- SETUP SHEET CONFIGURATION ---
  setupSheet.clear();
  setupSheet.getRange("A1:B1").merge().setValue("TRIP CONFIGURATION")
    .setBackground(COLOR.HEADER_DARK).setFontColor("#FFFFFF").setFontWeight("bold");

  const configLabels = [
    ["Trip Name / Destination:", "Goa Vacation 2026"],
    ["Number of People:", 4],
    ["Currency Symbol:", "₹"],
    ["Setup Status:", "Checking..."],
    ["All Expenses Confirmed:", false]
  ];
  setupSheet.getRange("A2:B6").setValues(configLabels);
  setupSheet.getRange("A2:A6").setFontWeight("bold");

  // Master confirm checkbox (mobile-friendly; menus are PC-only)
  setupSheet.getRange(CONFIRM_CELL).insertCheckboxes();
  setupSheet.getRange("A6").setNote(
    "Tick this AFTER entering all expenses. Totals and per-person splits appear only while this is ticked.");

  // Number of people: true dropdown listing 2 to 15
  const peopleOptions = [];
  for (let i = 2; i <= 15; i++) peopleOptions.push(String(i));
  const numPeopleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(peopleOptions, true)
    .setAllowInvalid(false)
    .setHelpText("Pick the number of people (2 to 15).")
    .build();
  setupSheet.getRange("B3").setDataValidation(numPeopleRule);

  // Roster header — Email is OPTIONAL (needed only for PDF reports)
  setupSheet.getRange("A8:C8").setValues([["#", "Name", "Email (optional)"]])
    .setBackground(COLOR.HEADER_MID).setFontColor("#FFFFFF").setFontWeight("bold");

  const initialRoster = [
    [1, "Rahul", ""],
    [2, "Priya", ""],
    [3, "Amit", ""],
    [4, "Neha", ""]
  ];
  setupSheet.getRange(9, 1, initialRoster.length, 3).setValues(initialRoster);

  applyRobotoFont(setupSheet);
  setupSheet.autoResizeColumns(1, 3);

  validateSetupStatus();
  updateExpensesStructure();
  recalculateAll();

  safeAlert("Setup complete! Edit your Trip details and Roster in the 'Setup' tab.");
}

function applyRobotoFont(sheet) {
  const range = sheet.getDataRange();
  if (range.getNumRows() && range.getNumColumns()) {
    range.setFontFamily("Roboto").setFontSize(12);
  }
}

/**
 * True only when the master "All Expenses Confirmed" checkbox is ticked.
 */
function isExpensesConfirmed() {
  const setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETUP);
  return setupSheet ? setupSheet.getRange(CONFIRM_CELL).getValue() === true : false;
}

/**
 * Validates trip config & roster names (emails are OPTIONAL).
 */
function validateSetupStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  if (!setupSheet) return false;

  const tripName = setupSheet.getRange("B2").getValue();
  const numPeople = parseInt(setupSheet.getRange("B3").getValue(), 10);
  const currency = setupSheet.getRange("B4").getValue();
  const statusCell = setupSheet.getRange("B5");

  if (!tripName) {
    statusCell.setValue("❌ Missing Trip Name").setFontColor(COLOR.RED_TX).setFontWeight("normal");
    return false;
  }
  if (!currency) {
    statusCell.setValue("❌ Missing Currency Symbol").setFontColor(COLOR.RED_TX).setFontWeight("normal");
    return false;
  }
  if (!numPeople || numPeople < 2) {
    statusCell.setValue("❌ Select Number of People").setFontColor(COLOR.RED_TX).setFontWeight("normal");
    return false;
  }

  const rosterData = setupSheet.getRange(9, 1, numPeople, 3).getValues();
  const names = [];
  let namedCount = 0;

  for (let i = 0; i < numPeople; i++) {
    const name = String(rosterData[i][1]).trim();
    if (!name) continue;
    namedCount++;
    if (names.includes(name.toLowerCase())) {
      statusCell.setValue(`❌ Duplicate name found: "${name}"`).setFontColor(COLOR.RED_TX).setFontWeight("normal");
      return false;
    }
    names.push(name.toLowerCase());
  }

  if (namedCount < numPeople) {
    statusCell.setValue(`⚠️ ${namedCount} of ${numPeople} people named — fill all names`)
      .setFontColor(COLOR.WARN_TX).setFontWeight("normal");
    return false;
  }

  statusCell.setValue("Setup complete ✅").setFontColor("#166534").setFontWeight("bold");
  return true;
}

/**
 * Roster from Setup (rows 9+). Only rows WITH a name are returned.
 */
function getRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  if (!setupSheet) return [];

  const numPeople = parseInt(setupSheet.getRange("B3").getValue(), 10) || 0;
  if (numPeople === 0) return [];

  const rawData = setupSheet.getRange(9, 1, numPeople, 3).getValues();
  const roster = [];
  rawData.forEach(row => {
    const name = String(row[1]).trim();
    if (name) roster.push({ id: row[0], name: name, email: String(row[2]).trim() });
  });
  return roster;
}

/**
 * Parses "Split Among": blank or "All" = everyone, else comma-separated names.
 */
function parseIncluded(splitStr, names) {
  const s = String(splitStr).trim();
  if (!s || s.toLowerCase() === "all") return [...names];
  const splitList = s.split(",").map(x => x.trim().toLowerCase());
  return names.filter(n => splitList.includes(n.toLowerCase()));
}

/**
 * Simple trigger: reacts to edits on Setup and Expenses.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const a1 = e.range.getA1Notation();

  if (sheetName === SHEET_SETUP) {
    // The master checkbox is the only Setup edit that needs an immediate
    // calculation; this works on both mobile and desktop.
    if (a1 === CONFIRM_CELL) {
      recalculateAll();
      return;
    }

    // Rebuild headers / validations only when trip configuration or roster
    // content changes — never for arbitrary Setup formatting edits.
    const setupConfigChanged = a1 === "B2" || a1 === "B3" || a1 === "B4";
    const rosterChanged = row >= 9 && col >= 1 && col <= 3;

    if (a1 === "B3") adjustRosterRows();

    if (setupConfigChanged || rosterChanged) {
      validateSetupStatus();
      updateExpensesStructure();
      recalculateAll();
    }
    return;
  }

  // Recalculate only for editable expense input fields A:E. Ignore clicks,
  // calculated person columns, formatting changes and report-sheet changes.
  if (sheetName === SHEET_EXPENSES && row >= 2 && row <= MAX_DATA_ROWS + 1 && col >= COL.DATE && col <= COL.SPLIT) {
    recalculateAll();
  }
}

/**
 * Adjusts Roster rows in Setup to match N (cell B3).
 * New rows start with a BLANK name (no expense column until named).
 */
function adjustRosterRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const targetN = parseInt(setupSheet.getRange("B3").getValue(), 10);
  if (!targetN) return;

  const currentRosterCount = Math.max(0, setupSheet.getLastRow() - 8);

  if (targetN > currentRosterCount) {
    const rowsToAdd = targetN - currentRosterCount;
    const newRows = [];
    for (let i = 1; i <= rowsToAdd; i++) {
      newRows.push([currentRosterCount + i, "", ""]);
    }
    setupSheet.getRange(9 + currentRosterCount, 1, rowsToAdd, 3).setValues(newRows);
  } else if (targetN < currentRosterCount) {
    setupSheet.deleteRows(9 + targetN, currentRosterCount - targetN);
  }

  applyRobotoFont(setupSheet);
  setupSheet.autoResizeColumns(1, 3);
}

/**
 * ONE-TIME CLEANER: wipes ALL content, formatting and validations
 * below the Expenses header row, then rebuilds.
 * (Keeps the header row; re-enter your rows afterwards.)
 */
function fixExpensesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!expenseSheet) return;

  clearAllProtections(expenseSheet);

  if (expenseSheet.getMaxRows() > 1) {
    expenseSheet
      .getRange(2, 1, expenseSheet.getMaxRows() - 1, expenseSheet.getMaxColumns())
      .clearContent().clearFormat().clearDataValidations();
  }

  updateExpensesStructure();
  recalculateAll();
  safeAlert("Expenses sheet cleaned. Re-enter your expense rows.");
}

/**
 * Removes every range protection on a sheet (best effort).
 */
function clearAllProtections(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    try { p.remove(); } catch (e) { /* not ours — ignore */ }
  });
}

/**
 * Rebuilds headers, validations and formats on Expenses.
 * Removes leftover person columns when the roster shrinks (5 -> 3 etc),
 * and normalizes person columns (clears stale checkbox validation and
 * any colored fills left by older versions).
 */
function updateExpensesStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!expenseSheet) return;

  const roster = getRoster();
  const names = roster.map(r => r.name);
  const headers = ["Date", "Description", "Amount", "Paid By", "Split Among", ...names];
  const newWidth = headers.length;
  const currencySymbol = ss.getSheetByName(SHEET_SETUP).getRange("B4").getValue() || "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;

  // This function runs only when setup/roster changes, not after normal
  // expense typing. Keep structural work here rather than in onEdit's hot path.
  clearAllProtections(expenseSheet);

  const oldWidth = expenseSheet.getLastColumn();
  if (oldWidth > newWidth) {
    expenseSheet.getRange(1, newWidth + 1, expenseSheet.getMaxRows(), oldWidth - newWidth)
      .clearContent()
      .clearFormat()
      .clearDataValidations();
  }

  expenseSheet.getRange(1, 1, 1, newWidth).setValues([headers])
    .setBackground(COLOR.HEADER)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  expenseSheet.setFrozenRows(1);

  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Enter a valid date, e.g. 09-Jan-2026.")
    .build();
  expenseSheet.getRange(2, COL.DATE, MAX_DATA_ROWS, 1)
    .setDataValidation(dateRule)
    .setNumberFormat("dd-mmm-yyyy");

  if (names.length) {
    const paidByRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true)
      .setAllowInvalid(false)
      .build();

    expenseSheet.getRange(2, COL.PAID_BY, MAX_DATA_ROWS, 1).setDataValidation(paidByRule);
    expenseSheet.getRange(2, FIRST_PERSON_COL, MAX_DATA_ROWS, names.length)
      .clearDataValidations()
      .setBackground("#FFFFFF")
      .setFontColor("#000000")
      .setFontWeight("normal")
      .setNumberFormat(currencyFormat);
  } else {
    expenseSheet.getRange(2, COL.PAID_BY, MAX_DATA_ROWS, 1).clearDataValidations();
  }

  expenseSheet.getRange(2, COL.AMOUNT, MAX_DATA_ROWS, 1).setNumberFormat(currencyFormat);
}

/**
 * True if the given row is the TOTAL summary row.
 * The label sits in column B because column A enforces date validation
 * (which rejects script writes too).
 */
function isTotalRow_(sheet, rowIndex) {
  const a = String(sheet.getRange(rowIndex, COL.DATE).getValue()).toUpperCase();
  const b = String(sheet.getRange(rowIndex, COL.DESC).getValue()).toUpperCase();
  return a.indexOf(TOTAL_LABEL) === 0 || b.indexOf(TOTAL_LABEL) === 0;
}

/**
 * Number of real data rows (excludes header and any TOTAL row).
 */
function countDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  return isTotalRow_(sheet, lastRow) ? lastRow - 2 : lastRow - 1;
}

/**
 * Returns last data row, deleting any existing TOTAL row first.
 */
function stripTotalRow(expenseSheet) {
  let lastRow = expenseSheet.getLastRow();
  if (lastRow > 1 && isTotalRow_(expenseSheet, lastRow)) {
    expenseSheet.deleteRow(lastRow);
    lastRow--;
  }
  return lastRow;
}

/**
 * Core recalculation. Person shares + TOTAL row appear ONLY when the
 * master "All Expenses Confirmed" checkbox (Setup!B6) is ticked.
 */
function recalculateAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const roster = getRoster();
  if (!expenseSheet || !setupSheet) return;

  const names = roster.map(r => r.name);
  if (names.length === 0) { updateSendReportsSheet(); return; }

  const currencySymbol = setupSheet.getRange("B4").getValue() || "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;
  const width = COL.SPLIT + names.length;

  const dataLastRow = stripTotalRow(expenseSheet);
  const confirmed = isExpensesConfirmed();

  if (dataLastRow >= 2 && confirmed) {
    const data = expenseSheet.getRange(2, 1, dataLastRow - 1, width).getValues();
    const shareMatrix = [];
    let totalAmount = 0;
    const personTotals = new Array(names.length).fill(0);

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const amount = parseFloat(row[COL.AMOUNT - 1]) || 0;
      const paidBy = String(row[COL.PAID_BY - 1]).trim();
      const rowNet = new Array(names.length).fill("");

      if (amount > 0 && paidBy) {
        const included = parseIncluded(row[COL.SPLIT - 1], names);
        if (included.length > 0) {
          // Equal split, 2-decimal precision, remainder to the payer
          const baseShare = Math.floor((amount / included.length) * 100) / 100;
          const remainder = Math.round((amount - baseShare * included.length) * 100) / 100;

          names.forEach((name, idx) => {
            const consumed = included.includes(name)
              ? baseShare + (name === paidBy ? remainder : 0) : 0;
            const paid = (name === paidBy) ? amount : 0;
            rowNet[idx] = Math.round((paid - consumed) * 100) / 100;
            personTotals[idx] = Math.round((personTotals[idx] + rowNet[idx]) * 100) / 100;
          });
          totalAmount = Math.round((totalAmount + amount) * 100) / 100;
        }
      }
      shareMatrix.push(rowNet);
    }

    const shareRange = expenseSheet.getRange(2, FIRST_PERSON_COL, shareMatrix.length, names.length);
    shareRange.setValues(shareMatrix).setNumberFormat(currencyFormat);

    // TOTAL row — label in Description (B); Date col (A) stays empty
    // because its date validation rejects non-date writes, even from scripts.
    const totalsRowIndex = dataLastRow + 1;
    const totalsRow = ["", TOTAL_LABEL, totalAmount, "", ""].concat(personTotals);
    expenseSheet.getRange(totalsRowIndex, 1, 1, width).setValues([totalsRow])
      .setBackground(COLOR.TOTAL_BG).setFontWeight("bold");
    expenseSheet.getRange(totalsRowIndex, COL.AMOUNT).setNumberFormat(currencyFormat);
    expenseSheet.getRange(totalsRowIndex, FIRST_PERSON_COL, 1, names.length).setNumberFormat(currencyFormat);

  } else if (dataLastRow >= 2 && !confirmed) {
    // Not confirmed: keep rows, but hide all computed values
    expenseSheet.getRange(2, FIRST_PERSON_COL, dataLastRow - 1, names.length).clearContent();
  }

  // Avoid rebuilding the reports sheet after every unconfirmed expense edit.
  if (confirmed || dataLastRow === 1) updateSendReportsSheet();
}

/**
 * CONFIRM HELPERS — one switch for everything.
 * confirmAllExpenses is assignable to an Insert -> Drawing button.
 */
function confirmAllExpenses() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETUP)
    .getRange(CONFIRM_CELL).setValue(true);
  recalculateAll();
}

function unconfirmExpenses() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETUP)
    .getRange(CONFIRM_CELL).setValue(false);
  recalculateAll();
}

/**
 * Expense rows as objects. Returns ZERO rows while unconfirmed,
 * so totals/reports stay hidden until the master switch is ticked.
 */
function getConfirmedExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const roster = getRoster();
  const names = roster.map(r => r.name);
  if (!expenseSheet || names.length === 0) return { expenses: [], roster, names };
  if (!isExpensesConfirmed()) return { expenses: [], roster, names };

  const dataLastRow = stripTotalRow(expenseSheet);
  if (dataLastRow < 2) return { expenses: [], roster, names };

  const data = expenseSheet.getRange(2, 1, dataLastRow - 1, COL.SPLIT + names.length).getValues();
  const tz = ss.getSpreadsheetTimeZone();
  const expenses = [];

  data.forEach(row => {
    const amount = parseFloat(row[COL.AMOUNT - 1]) || 0;
    const paidBy = String(row[COL.PAID_BY - 1]).trim();
    if (amount <= 0 || !paidBy) return;

    const dateVal = row[COL.DATE - 1] instanceof Date
      ? Utilities.formatDate(row[COL.DATE - 1], tz, "dd-MMM-yyyy")
      : String(row[COL.DATE - 1]);

    expenses.push({
      date: dateVal,
      desc: String(row[COL.DESC - 1]),
      amount: amount,
      paidBy: paidBy,
      included: parseIncluded(row[COL.SPLIT - 1], names)
    });
  });

  return { expenses, roster, names };
}

/**
 * Per-person totals across expenses.
 */
function computeTotals(expenses, names) {
  const totals = {};
  names.forEach(n => totals[n] = { paid: 0, share: 0 });

  expenses.forEach(e => {
    if (e.included.length === 0) return;
    const baseShare = Math.floor((e.amount / e.included.length) * 100) / 100;
    const remainder = Math.round((e.amount - baseShare * e.included.length) * 100) / 100;

    names.forEach(name => {
      const isPayer = e.paidBy === name;
      const isIncluded = e.included.includes(name);
      if (isPayer) totals[name].paid += e.amount;
      if (isIncluded) totals[name].share += baseShare + (isPayer ? remainder : 0);
    });
  });

  names.forEach(n => {
    totals[n].paid = Math.round(totals[n].paid * 100) / 100;
    totals[n].share = Math.round(totals[n].share * 100) / 100;
    totals[n].net = Math.round((totals[n].paid - totals[n].share) * 100) / 100;
  });
  return totals;
}

/**
 * Populates Summary & Simplified Settlements on "Send Reports".
 */
function updateSendReportsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const currencySymbol = setupSheet ? (setupSheet.getRange("B4").getValue() || "₹") : "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;
  const { expenses, roster, names } = getConfirmedExpenses();

  // Avoid clear() over the entire sheet and avoid autoresizeColumns() on every
  // edit. Only clear the small report area actually managed by this script.
  const requiredRows = Math.max(20, roster.length + 20);
  if (reportsSheet.getMaxRows() < requiredRows) {
    reportsSheet.insertRowsAfter(reportsSheet.getMaxRows(), requiredRows - reportsSheet.getMaxRows());
  }
  if (reportsSheet.getMaxColumns() < 6) {
    reportsSheet.insertColumnsAfter(reportsSheet.getMaxColumns(), 6 - reportsSheet.getMaxColumns());
  }

  const managedRange = reportsSheet.getRange(1, 1, requiredRows, 6);
  managedRange.clearContent();

  if (!isExpensesConfirmed()) {
    reportsSheet.getRange("A1:F1").merge()
      .setValue("⚠️ Expenses not confirmed yet — tick 'All Expenses Confirmed' on the Setup sheet to see totals and send reports.")
      .setBackground(COLOR.PENDING)
      .setFontColor(COLOR.WARN_TX)
      .setFontWeight("bold");
    return;
  }

  const totals = computeTotals(expenses, names);
  reportsSheet.getRange("A1:F1").merge()
    .setValue("PERSONAL BALANCES SUMMARY")
    .setBackground(COLOR.HEADER_DARK)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  reportsSheet.getRange("A2:F2")
    .setValues([["Name", "Email", "Total Paid", "Total Share", "Net Balance", "Report Status"]])
    .setBackground(COLOR.HEADER_MID)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  const summaryData = roster.map(person => {
    const t = totals[person.name] || { paid: 0, share: 0, net: 0 };
    return [
      person.name,
      person.email || "—",
      t.paid,
      t.share,
      t.net,
      person.email ? "Not Sent" : "No email — will be skipped"
    ];
  });

  if (summaryData.length) {
    reportsSheet.getRange(3, 1, summaryData.length, 6).setValues(summaryData);
    reportsSheet.getRange(3, 3, summaryData.length, 3).setNumberFormat(currencyFormat);
  }

  const balancesMap = {};
  names.forEach(name => balancesMap[name] = totals[name].net);
  const settlements = calculateGreedySettlements(balancesMap);
  const startRow = summaryData.length + 5;

  reportsSheet.getRange(startRow, 1, 1, 4).merge()
    .setValue("SIMPLIFIED SETTLEMENT PLAN")
    .setBackground(COLOR.HEADER)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  reportsSheet.getRange(startRow + 1, 1, 1, 4)
    .setValues([["From (Debtor)", "To (Creditor)", "Amount", "Instruction"]])
    .setBackground("#115E59")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  if (settlements.length) {
    const settlementRows = settlements.map(s => [
      s.from,
      s.to,
      s.amount,
      `${s.from} pays ${s.to} ${currencySymbol}${s.amount.toFixed(2)}`
    ]);
    reportsSheet.getRange(startRow + 2, 1, settlementRows.length, 4).setValues(settlementRows);
    reportsSheet.getRange(startRow + 2, 3, settlementRows.length, 1).setNumberFormat(currencyFormat);
  } else {
    reportsSheet.getRange(startRow + 2, 1, 1, 4).merge()
      .setValue("All debts settled!")
      .setFontStyle("italic");
  }
}

/**
 * Greedy algorithm to minimize total settlement transactions.
 */
function calculateGreedySettlements(balances) {
  const debtors = [];
  const creditors = [];

  for (const person in balances) {
    const bal = Math.round(balances[person] * 100) / 100;
    if (bal < -0.01) debtors.push({ name: person, amount: Math.abs(bal) });
    else if (bal > 0.01) creditors.push({ name: person, amount: bal });
  }

  debtors.sort((a, b) => b.amount - a.amount);
  creditors.sort((a, b) => b.amount - a.amount);

  const settlements = [];
  let d = 0, c = 0;

  while (d < debtors.length && c < creditors.length) {
    const debtor = debtors[d];
    const creditor = creditors[c];
    const amount = Math.min(debtor.amount, creditor.amount);

    settlements.push({ from: debtor.name, to: creditor.name, amount: Math.round(amount * 100) / 100 });

    debtor.amount = Math.round((debtor.amount - amount) * 100) / 100;
    creditor.amount = Math.round((creditor.amount - amount) * 100) / 100;

    if (debtor.amount === 0) d++;
    if (creditor.amount === 0) c++;
  }
  return settlements;
}

/**
 * Sends a plain test email to the account running the script.
 * Use this FIRST to verify the "Send email as you" permission is granted.
 */
function testEmailPermission() {
  let addr = Session.getActiveUser().getEmail();

  if (!addr) {
    try {
      const ui = SpreadsheetApp.getUi();
      const resp = ui.prompt("Test Email", "Enter your email address:", ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) return;
      addr = resp.getResponseText().trim();
    } catch (e) {
      safeAlert("Could not detect your email automatically. Open the spreadsheet on a computer and use the menu: Expense Splitter → Test Email Permission.");
      return;
    }
  }

  try {
    MailApp.sendEmail(addr, "Expense Splitter — Test",
      "If you can read this, the script has permission to send email as you.");
    safeAlert(`✅ Test email sent to ${addr}.\n\nSender = the Google account you are logged in with.\nCheck inbox AND spam.`);
  } catch (err) {
    safeAlert(`❌ Test FAILED:\n\n${err.message}\n\nFix: run any function from the Apps Script editor and approve the "Send email as you" permission.`);
  }
}

/**
 * Generates a personalized PDF per person and emails it.
 * Sender = the Google account running this function.
 * People without a valid email are skipped (logged in Sheet 3).
 */
function sendPdfReports() {
  if (!validateSetupStatus()) {
    safeAlert("Cannot send reports: Setup is incomplete. Check the 'Setup' tab status cell.");
    return;
  }
  if (!isExpensesConfirmed()) {
    safeAlert("Expenses are not confirmed yet. Tick 'All Expenses Confirmed' on the Setup sheet first.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const reportsSheet = ss.getSheetByName(SHEET_REPORTS);

  const tripName = setupSheet.getRange("B2").getValue();
  const currencySymbol = setupSheet.getRange("B4").getValue() || "₹";
  const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy");

  const { expenses, roster, names } = getConfirmedExpenses();
  if (expenses.length === 0) {
    safeAlert("No expense rows found on the 'Expenses' sheet.");
    return;
  }

  const totals = computeTotals(expenses, names);
  const balancesMap = {};
  names.forEach(n => balancesMap[n] = totals[n].net);
  const allSettlements = calculateGreedySettlements(balancesMap);

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  let sentCount = 0, skippedCount = 0;
  const failures = [];

  roster.forEach((person, idx) => {
    const statusCell = reportsSheet.getRange(3 + idx, 6);
    try {
      if (!person.email || !emailRegex.test(person.email)) {
        statusCell.setValue("Skipped (no email)").setFontColor(COLOR.WARN_TX);
        skippedCount++;
        return;
      }

      // Only expenses where this person paid or is included
      const relevant = [];
      expenses.forEach(e => {
        const isPayer = e.paidBy === person.name;
        const isIncluded = e.included.includes(person.name);
        if (!isPayer && !isIncluded) return;

        const baseShare = Math.floor((e.amount / e.included.length) * 100) / 100;
        const remainder = Math.round((e.amount - baseShare * e.included.length) * 100) / 100;
        const myShare = isIncluded ? baseShare + (isPayer ? remainder : 0) : 0;

        relevant.push({ date: e.date, desc: e.desc, amount: e.amount, paidBy: e.paidBy, myShare: myShare });
      });

      const myTotals = totals[person.name];
      const mySettlements = allSettlements.filter(s => s.from === person.name || s.to === person.name);

      const htmlContent = `
        <html>
          <head>
            <style>
              body { font-family: 'Roboto', sans-serif; margin: 20px; color: #333; }
              h1 { color: #0F766E; margin-bottom: 2px; }
              h3 { color: #64748B; margin-top: 0; font-weight: normal; }
              .card-container { display: flex; margin: 20px 0; gap: 10px; }
              .card { background: #F8FAFC; border: 1px solid #E2E8F0; padding: 12px; border-radius: 6px; flex: 1; text-align: center; }
              .card-title { font-size: 11px; color: #64748B; text-transform: uppercase; }
              .card-val { font-size: 18px; font-weight: bold; margin-top: 4px; }
              table { width: 100%; border-collapse: collapse; margin-top: 15px; }
              th { background: #1E293B; color: white; padding: 8px; font-size: 12px; text-align: left; }
              td { padding: 8px; border-bottom: 1px solid #E2E8F0; font-size: 11px; }
              .box { background: #F0FDF4; border: 1px solid #BBF7D0; padding: 12px; border-radius: 6px; margin-top: 20px; }
              .badge-green { color: #065F46; }
              .badge-red { color: #991B1B; }
            </style>
          </head>
          <body>
            <h1>${tripName}</h1>
            <h3>Expense Report for <b>${person.name}</b> &bull; Generated on ${todayStr}</h3>

            <div class="card-container">
              <div class="card"><div class="card-title">Total Paid</div><div class="card-val">${currencySymbol}${myTotals.paid.toFixed(2)}</div></div>
              <div class="card"><div class="card-title">Total Share</div><div class="card-val">${currencySymbol}${myTotals.share.toFixed(2)}</div></div>
              <div class="card"><div class="card-title">Net Balance</div><div class="card-val ${myTotals.net >= 0 ? 'badge-green' : 'badge-red'}">${currencySymbol}${myTotals.net.toFixed(2)}</div></div>
            </div>

            <h4>Your Relevant Expenses</h4>
            <table>
              <thead>
                <tr><th>Date</th><th>Description</th><th>Total Amount</th><th>Paid By</th><th>Your Share</th></tr>
              </thead>
              <tbody>
                ${relevant.length > 0 ? relevant.map(e => `
                  <tr>
                    <td>${e.date}</td>
                    <td>${e.desc}</td>
                    <td>${currencySymbol}${e.amount.toFixed(2)}</td>
                    <td>${e.paidBy}</td>
                    <td>${currencySymbol}${e.myShare.toFixed(2)}</td>
                  </tr>
                `).join('') : '<tr><td colspan="5">No expenses involve you.</td></tr>'}
              </tbody>
            </table>

            <div class="box">
              <h4 style="margin-top:0;">Your Settlement Action Items</h4>
              ${mySettlements.length > 0 ? `
                <ul>
                  ${mySettlements.map(s => `
                    <li>${s.from === person.name
                      ? `You owe <b>${s.to}</b> <b>${currencySymbol}${s.amount.toFixed(2)}</b>`
                      : `<b>${s.from}</b> owes you <b>${currencySymbol}${s.amount.toFixed(2)}</b>`}</li>
                  `).join('')}
                </ul>
              ` : '<p>You have no pending settlement actions!</p>'}
            </div>

            <h4 style="margin-top:25px;">Full Group Settlement Plan (For Transparency)</h4>
            <ul>
              ${allSettlements.length > 0
                ? allSettlements.map(s => `<li>${s.from} pays ${s.to} ${currencySymbol}${s.amount.toFixed(2)}</li>`).join('')
                : '<li>All settled up!</li>'}
            </ul>
          </body>
        </html>
      `;

      const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
        .getAs('application/pdf')
        .setName(`${tripName} - ${person.name} Expense Report.pdf`);

      const subject = `[${tripName}] — Your Expense Report`;
      const body = `Hi ${person.name},\n\nHere is your itemized expense report for "${tripName}". Please find your PDF breakdown attached.\n\nBest regards,\nGroup Expense Splitter`;

      MailApp.sendEmail(person.email, subject, body, { attachments: [pdfBlob] });

      const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMM HH:mm");
      statusCell.setValue(`Sent (${timestamp})`).setFontColor("#166534");
      sentCount++;

    } catch (err) {
      statusCell.setValue(`Failed: ${err.message}`).setFontColor(COLOR.RED_TX);
      failures.push(`${person.name}: ${err.message}`);
    }
  });

  let msg = `Done! Sent: ${sentCount}, Skipped (no email): ${skippedCount}, Failed: ${failures.length}.`;
  if (failures.length > 0) msg += `\n\nFailure reasons:\n• ${failures.join("\n• ")}`;
  msg += `\n\nNote: emails are sent FROM the Google account that ran this function — check that account's Sent folder.`;
  safeAlert(msg);
}
