/**
 * DYNAMIC GROUP TRAVEL EXPENSE SPLITTER
 * ------------------------------------------------------------------
 * Custom Menu: "Expense Splitter" (PC only - menus do NOT appear in
 * the Google Sheets mobile app. On mobile, everything works via
 * checkboxes + automatic onEdit triggers.)
 *
 * INSTALL:
 *   1. Extensions -> Apps Script -> paste this file -> Save
 *   2. Run setup() once and authorize the requested scopes
 *      (MailApp scope is required to send emails; emails are sent
 *      from the Google account that runs the function)
 *   3. Reload the spreadsheet to see the "Expense Splitter" menu
 *
 * EXPENSE FLOW:
 *   - Add expense rows on the "Expenses" sheet
 *   - Tick the "Confirm?" checkbox (works on mobile) OR use menu
 *     "Confirm Selected Rows" / "Confirm All Rows" (PC)
 *   - Per-person columns + TOTAL row only count CONFIRMED rows
 *   - Menu "Send PDF Reports via Email" emails each person only the
 *     confirmed expenses that involve them + their settlement plan
 */

const SHEET_SETUP = "Setup";
const SHEET_EXPENSES = "Expenses";
const SHEET_REPORTS = "Send Reports";

const COL = { DATE: 1, DESC: 2, AMOUNT: 3, PAID_BY: 4, SPLIT: 5, CONFIRM: 6 };
const FIRST_PERSON_COL = 7; // person share columns start at column G
const MAX_DATA_ROWS = 500;

const COLOR = {
  HEADER: "#0F766E",
  HEADER_DARK: "#1E293B",
  HEADER_MID: "#334155",
  PENDING: "#FEF3C7",   // amber for unconfirmed rows
  TOTAL_BG: "#E2E8F0",
  GREEN_BG: "#D1FAE5", GREEN_TX: "#065F46",
  RED_BG: "#FEE2E2",   RED_TX: "#991B1B",
  GRAY_BG: "#F3F4F6",  GRAY_TX: "#374151"
};

/**
 * Creates custom menu when the spreadsheet opens (PC / desktop only).
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Splitter')
    .addItem('1. Initialize / Reset Tracker', 'setup')
    .addItem('2. Recalculate Expenses & Settlements', 'recalculateAll')
    .addSeparator()
    .addItem('Confirm Selected Rows', 'confirmSelectedExpenses')
    .addItem('Confirm All Rows', 'confirmAllExpenses')
    .addSeparator()
    .addItem('3. Send PDF Reports via Email', 'sendPdfReports')
    .addToUi();
}

/**
 * Master initialization function.
 * Builds all sheets, labels, roster table, defaults, and formatting.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const setupSheet = ss.getSheetByName(SHEET_SETUP) || ss.insertSheet(SHEET_SETUP);
  ss.getSheetByName(SHEET_EXPENSES) || ss.insertSheet(SHEET_EXPENSES);
  ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);

  // --- 1. SETUP SHEET CONFIGURATION ---
  setupSheet.clear();
  setupSheet.getRange("A1:B1").merge().setValue("TRIP CONFIGURATION")
    .setBackground(COLOR.HEADER_DARK).setFontColor("#FFFFFF").setFontWeight("bold");

  const configLabels = [
    ["Trip Name / Destination:", "Goa Vacation 2026"],
    ["Number of People:", 4],
    ["Currency Symbol:", "₹"],
    ["Setup Status:", "Checking..."]
  ];
  setupSheet.getRange("A2:B5").setValues(configLabels);
  setupSheet.getRange("A2:A5").setFontWeight("bold");

  // Number of people: true dropdown listing 2 to 15
  const peopleOptions = [];
  for (let i = 2; i <= 15; i++) peopleOptions.push(String(i));
  const numPeopleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(peopleOptions, true)
    .setAllowInvalid(false)
    .setHelpText("Pick the number of people (2 to 15).")
    .build();
  setupSheet.getRange("B3").setDataValidation(numPeopleRule);

  // Roster header — Email is OPTIONAL (only needed if you want PDF reports emailed)
  setupSheet.getRange("A7:C7").setValues([["#", "Name", "Email (optional)"]])
    .setBackground(COLOR.HEADER_MID).setFontColor("#FFFFFF").setFontWeight("bold");

  const initialRoster = [
    [1, "Rahul", ""],
    [2, "Priya", ""],
    [3, "Amit", ""],
    [4, "Neha", ""]
  ];
  setupSheet.getRange(8, 1, initialRoster.length, 3).setValues(initialRoster);

  applyRobotoFont(setupSheet);
  setupSheet.autoResizeColumns(1, 3);

  // --- 2. INITIALIZE EXPENSES & REPORTS ---
  validateSetupStatus();
  updateExpensesStructure();
  recalculateAll();

  SpreadsheetApp.getUi().alert("Setup complete! Edit your Trip details and Roster in the 'Setup' tab.");
}

/**
 * Helper to apply Roboto 12pt across a given sheet.
 */
function applyRobotoFont(sheet) {
  sheet.getDataRange().setFontFamily("Roboto").setFontSize(12);
}

/**
 * Validates trip config & roster names (emails are OPTIONAL).
 * Updates the status cell B5 in Setup. Returns true if usable.
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

  const rosterData = setupSheet.getRange(8, 1, numPeople, 3).getValues();
  const names = [];
  let namedCount = 0;

  for (let i = 0; i < numPeople; i++) {
    const name = String(rosterData[i][1]).trim();
    if (!name) continue; // unnamed rows simply don't get a column yet
    namedCount++;
    if (names.includes(name.toLowerCase())) {
      statusCell.setValue(`❌ Duplicate name found: "${name}"`).setFontColor(COLOR.RED_TX).setFontWeight("normal");
      return false;
    }
    names.push(name.toLowerCase());
  }

  if (namedCount < numPeople) {
    statusCell.setValue(`⚠️ ${namedCount} of ${numPeople} people named — fill all names`)
      .setFontColor("#B45309").setFontWeight("normal");
    return false;
  }

  statusCell.setValue("Setup complete ✅").setFontColor("#166534").setFontWeight("bold");
  return true;
}

/**
 * Reads the roster from Setup. Only rows WITH a name are returned,
 * so unnamed placeholder rows never become expense columns.
 */
function getRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  if (!setupSheet) return [];

  const numPeople = parseInt(setupSheet.getRange("B3").getValue(), 10) || 0;
  if (numPeople === 0) return [];

  const rawData = setupSheet.getRange(8, 1, numPeople, 3).getValues();
  const roster = [];
  rawData.forEach(row => {
    const name = String(row[1]).trim();
    if (name) roster.push({ id: row[0], name: name, email: String(row[2]).trim() });
  });
  return roster;
}

/**
 * Parses the "Split Among" cell into a list of roster names.
 * Blank or "All" = everyone. Otherwise comma-separated names.
 */
function parseIncluded(splitStr, names) {
  const s = String(splitStr).trim();
  if (!s || s.toLowerCase() === "all") return [...names];
  const splitList = s.split(",").map(x => x.trim().toLowerCase());
  return names.filter(n => splitList.includes(n.toLowerCase()));
}

/**
 * Simple trigger: reacts to edits on Setup and Expenses.
 * NOTE: UI dialogs (alerts/msgBox) are NOT allowed in simple triggers,
 * so structural changes apply immediately without a prompt.
 */
function onEdit(e) {
  if (!e) return;
  const sheetName = e.range.getSheet().getName();

  if (sheetName === SHEET_SETUP) {
    if (e.range.getA1Notation() === "B3") {
      adjustRosterRows(); // also removes extra person columns via updateExpensesStructure below
    }
    validateSetupStatus();
    updateExpensesStructure();
    recalculateAll();
  } else if (sheetName === SHEET_EXPENSES) {
    // Don't rebuild structure if the edit was in a computed person column
    recalculateAll();
  }
}

/**
 * Adjusts the Roster table row count in Setup to match N (cell B3).
 * New rows start with a BLANK name so they don't create expense columns
 * until a real name is typed.
 */
function adjustRosterRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const targetN = parseInt(setupSheet.getRange("B3").getValue(), 10);
  if (!targetN) return;

  const currentRosterCount = Math.max(0, setupSheet.getLastRow() - 7);

  if (targetN > currentRosterCount) {
    const rowsToAdd = targetN - currentRosterCount;
    const newRows = [];
    for (let i = 1; i <= rowsToAdd; i++) {
      const nextId = currentRosterCount + i;
      newRows.push([nextId, "", ""]); // blank name on purpose
    }
    setupSheet.getRange(8 + currentRosterCount, 1, rowsToAdd, 3).setValues(newRows);
  } else if (targetN < currentRosterCount) {
    setupSheet.deleteRows(8 + targetN, currentRosterCount - targetN);
  }

  applyRobotoFont(setupSheet);
  setupSheet.autoResizeColumns(1, 3);
}

/**
 * Rebuilds headers, validations, checkboxes and formats on Expenses.
 * Also REMOVES leftover person columns when the roster shrinks
 * (e.g. count changed 5 -> 3 removes the last 2 columns entirely).
 */
function updateExpensesStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!expenseSheet) return;

  const roster = getRoster();
  const names = roster.map(r => r.name);
  const headers = ["Date", "Description", "Amount", "Paid By", "Split Among", "Confirm?", ...names];
  const newWidth = headers.length;

  // Remove protections so structure can be modified
  expenseSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  // Clear leftover columns from a previous wider roster
  const oldWidth = expenseSheet.getLastColumn();
  if (oldWidth > newWidth) {
    expenseSheet
      .getRange(1, newWidth + 1, expenseSheet.getMaxRows(), oldWidth - newWidth)
      .clearContent().clearFormat().clearDataValidations();
  }

  // Headers
  expenseSheet.getRange(1, 1, 1, newWidth).setValues([headers])
    .setBackground(COLOR.HEADER).setFontColor("#FFFFFF").setFontWeight("bold");
  expenseSheet.setFrozenRows(1);

  // Date validation on the Date column (works on mobile + PC)
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("Enter a valid date, e.g. 09-Jan-2026.")
    .build();
  expenseSheet.getRange(2, COL.DATE, MAX_DATA_ROWS, 1).setDataValidation(dateRule);

  // "Paid By" dropdown from current roster names
  if (names.length > 0) {
    const paidByRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true)
      .setAllowInvalid(false)
      .build();
    expenseSheet.getRange(2, COL.PAID_BY, MAX_DATA_ROWS, 1).setDataValidation(paidByRule);
  } else {
    expenseSheet.getRange(2, COL.PAID_BY, MAX_DATA_ROWS, 1).clearDataValidations();
  }

  // "Confirm?" checkboxes — the mobile-friendly confirm button
  expenseSheet.getRange(2, COL.CONFIRM, MAX_DATA_ROWS, 1).insertCheckboxes();

  // Formats
  const currencySymbol = ss.getSheetByName(SHEET_SETUP).getRange("B4").getValue() || "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;
  expenseSheet.getRange(2, COL.DATE, MAX_DATA_ROWS, 1).setNumberFormat("dd-mmm-yyyy");
  expenseSheet.getRange(2, COL.AMOUNT, MAX_DATA_ROWS, 1).setNumberFormat(currencyFormat);

  applyRobotoFont(expenseSheet);
  expenseSheet.autoResizeColumns(1, newWidth);
}

/**
 * Returns the last data row on the Expenses sheet, and removes any
 * existing TOTAL row. Person column count is respected.
 */
function stripTotalRow(expenseSheet) {
  let lastRow = expenseSheet.getLastRow();
  if (lastRow > 1) {
    const firstCell = String(expenseSheet.getRange(lastRow, 1).getValue()).toUpperCase();
    if (firstCell.indexOf("TOTAL") === 0) {
      expenseSheet.deleteRow(lastRow);
      lastRow--;
    }
  }
  return lastRow; // last row of actual data (1 = no data)
}

/**
 * Core recalculation: per-person share columns + TOTAL row.
 * ONLY rows with Confirm? = TRUE are calculated and totalled.
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

  const dataLastRow = stripTotalRow(expenseSheet);

  let totalAmount = 0;
  const personTotals = new Array(names.length).fill(0);

  if (dataLastRow >= 2) {
    const width = COL.CONFIRM + names.length; // A..F + person cols
    const data = expenseSheet.getRange(2, 1, dataLastRow - 1, width).getValues();
    const shareMatrix = [];
    const rowBackgrounds = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const amount = parseFloat(row[COL.AMOUNT - 1]) || 0;
      const paidBy = String(row[COL.PAID_BY - 1]).trim();
      const confirmed = row[COL.CONFIRM - 1] === true;

      const rowNet = new Array(names.length).fill("");

      if (confirmed && amount > 0 && paidBy) {
        const included = parseIncluded(row[COL.SPLIT - 1], names);

        if (included.length > 0) {
          // Equal split, 2-decimal precision, remainder goes to the payer
          const baseShare = Math.floor((amount / included.length) * 100) / 100;
          const remainder = Math.round((amount - baseShare * included.length) * 100) / 100;

          names.forEach((name, idx) => {
            const consumed = included.includes(name)
              ? baseShare + (name === paidBy ? remainder : 0)
              : 0;
            const paid = (name === paidBy) ? amount : 0;
            rowNet[idx] = Math.round((paid - consumed) * 100) / 100;
            personTotals[idx] = Math.round((personTotals[idx] + rowNet[idx]) * 100) / 100;
          });
          totalAmount = Math.round((totalAmount + amount) * 100) / 100;
        }
      }

      shareMatrix.push(rowNet);
      // Amber background = pending confirmation, white = confirmed
      const bg = confirmed ? "#FFFFFF" : COLOR.PENDING;
      rowBackgrounds.push(new Array(COL.CONFIRM).fill(bg));
    }

    // Write per-person shares
    const shareRange = expenseSheet.getRange(2, FIRST_PERSON_COL, shareMatrix.length, names.length);
    shareRange.setValues(shareMatrix).setNumberFormat(currencyFormat).setBackground("#FFFFFF");

    // Highlight input area by confirmation state
    expenseSheet.getRange(2, 1, shareMatrix.length, COL.CONFIRM).setBackgrounds(rowBackgrounds);

    // TOTAL row (confirmed rows only) — written as static values
    const totalsRowIndex = dataLastRow + 1;
    const totalsRow = ["TOTAL (Confirmed)", "", totalAmount, "", "", ""].concat(personTotals);
    expenseSheet.getRange(totalsRowIndex, 1, 1, width).setValues([totalsRow])
      .setBackground(COLOR.TOTAL_BG).setFontWeight("bold");
    expenseSheet.getRange(totalsRowIndex, COL.AMOUNT).setNumberFormat(currencyFormat);
    expenseSheet.getRange(totalsRowIndex, FIRST_PERSON_COL, 1, names.length).setNumberFormat(currencyFormat);
  }

  protectExpensesSheet(expenseSheet, names.length);
  updateSendReportsSheet();
}

/**
 * Protects header row, computed person columns and the TOTAL row.
 * Input columns (Date..Confirm?) stay editable for everyone.
 */
function protectExpensesSheet(sheet, numPeople) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  const headerProt = sheet.getRange(1, 1, 1, COL.CONFIRM + numPeople).protect()
    .setDescription("Header Protection");
  headerProt.removeEditors(headerProt.getEditors());

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const shareProt = sheet.getRange(2, FIRST_PERSON_COL, lastRow - 1, numPeople).protect()
    .setDescription("Per-person Formula Protection");
  shareProt.removeEditors(shareProt.getEditors());
}

/**
 * CONFIRM HELPERS
 * confirmAllExpenses: ticks every data row (also assignable to a
 * drawing/image button via right-click -> Assign script).
 */
function confirmAllExpenses() {
  const expenseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXPENSES);
  if (!expenseSheet) return;
  const dataLastRow = stripTotalRow(expenseSheet);
  if (dataLastRow < 2) return;
  expenseSheet.getRange(2, COL.CONFIRM, dataLastRow - 1, 1).setValue(true);
  recalculateAll();
}

/**
 * Ticks only the currently selected rows on the Expenses sheet.
 */
function confirmSelectedExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getActiveRange();
  if (!range || range.getSheet().getName() !== SHEET_EXPENSES) {
    SpreadsheetApp.getUi().alert("Select one or more expense rows on the 'Expenses' sheet first.");
    return;
  }
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const dataLastRow = stripTotalRow(expenseSheet);
  const startRow = Math.max(2, range.getRow());
  const endRow = Math.min(dataLastRow, range.getLastRow());
  if (endRow < startRow) return;
  expenseSheet.getRange(startRow, COL.CONFIRM, endRow - startRow + 1, 1).setValue(true);
  recalculateAll();
}

/**
 * Reads CONFIRMED expense rows as objects shared by reports & emails.
 */
function getConfirmedExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const roster = getRoster();
  const names = roster.map(r => r.name);
  if (!expenseSheet || names.length === 0) return { expenses: [], roster, names };

  const dataLastRow = stripTotalRow(expenseSheet);
  if (dataLastRow < 2) return { expenses: [], roster, names };

  const data = expenseSheet.getRange(2, 1, dataLastRow - 1, COL.CONFIRM + names.length).getValues();
  const tz = ss.getSpreadsheetTimeZone();
  const expenses = [];

  data.forEach(row => {
    if (row[COL.CONFIRM - 1] !== true) return; // skip unconfirmed
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
 * Per-person totals across confirmed expenses.
 * share = what they consumed, paid = what they fronted.
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
 * Populates Summary & Simplified Settlements on the "Send Reports" sheet.
 */
function updateSendReportsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const currencySymbol = setupSheet ? (setupSheet.getRange("B4").getValue() || "₹") : "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;

  reportsSheet.clear();

  const { expenses, roster, names } = getConfirmedExpenses();
  const totals = computeTotals(expenses, names);

  // --- 1. SUMMARY TABLE ---
  reportsSheet.getRange("A1:F1").merge().setValue("PERSONAL BALANCES SUMMARY (confirmed expenses only)")
    .setBackground(COLOR.HEADER_DARK).setFontColor("#FFFFFF").setFontWeight("bold");

  reportsSheet.getRange("A2:F2")
    .setValues([["Name", "Email", "Total Paid", "Total Share", "Net Balance", "Report Status"]])
    .setBackground(COLOR.HEADER_MID).setFontColor("#FFFFFF").setFontWeight("bold");

  const summaryData = roster.map(person => {
    const t = totals[person.name];
    return [person.name, person.email || "—", t.paid, t.share, t.net,
      person.email ? "Not Sent" : "No email — will be skipped"];
  });

  if (summaryData.length > 0) {
    const summaryRange = reportsSheet.getRange(3, 1, summaryData.length, 6);
    summaryRange.setValues(summaryData);
    reportsSheet.getRange(3, 3, summaryData.length, 3).setNumberFormat(currencyFormat);

    for (let i = 0; i < summaryData.length; i++) {
      const cell = reportsSheet.getRange(3 + i, 5);
      const val = summaryData[i][4];
      if (val > 0.01) cell.setBackground(COLOR.GREEN_BG).setFontColor(COLOR.GREEN_TX);
      else if (val < -0.01) cell.setBackground(COLOR.RED_BG).setFontColor(COLOR.RED_TX);
      else cell.setBackground(COLOR.GRAY_BG).setFontColor(COLOR.GRAY_TX);
    }
  }

  // --- 2. SIMPLIFIED SETTLEMENTS (greedy, minimizes transactions) ---
  const balancesMap = {};
  names.forEach(n => balancesMap[n] = totals[n].net);
  const settlements = calculateGreedySettlements(balancesMap);

  const startRow = summaryData.length + 5;
  reportsSheet.getRange(startRow, 1, 1, 4).merge().setValue("SIMPLIFIED SETTLEMENT PLAN")
    .setBackground(COLOR.HEADER).setFontColor("#FFFFFF").setFontWeight("bold");

  reportsSheet.getRange(startRow + 1, 1, 1, 4)
    .setValues([["From (Debtor)", "To (Creditor)", "Amount", "Instruction"]])
    .setBackground("#115E59").setFontColor("#FFFFFF").setFontWeight("bold");

  if (settlements.length > 0) {
    const rows = settlements.map(s => [
      s.from, s.to, s.amount,
      `${s.from} pays ${s.to} ${currencySymbol}${s.amount.toFixed(2)}`
    ]);
    reportsSheet.getRange(startRow + 2, 1, rows.length, 4).setValues(rows);
    reportsSheet.getRange(startRow + 2, 3, rows.length, 1).setNumberFormat(currencyFormat);
  } else {
    reportsSheet.getRange(startRow + 2, 1, 1, 4).merge()
      .setValue("All debts settled! Net balance is 0 for everyone.").setFontStyle("italic");
  }

  applyRobotoFont(reportsSheet);
  reportsSheet.autoResizeColumns(1, 6);
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
 * Generates a personalized PDF per person and emails it.
 * Emails are sent FROM the Google account running this function.
 * People without a valid email are skipped and logged in Sheet 3.
 */
function sendPdfReports() {
  const ui = SpreadsheetApp.getUi();
  if (!validateSetupStatus()) {
    ui.alert("Cannot send reports: Setup is incomplete. Check the 'Setup' tab status cell.");
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
    ui.alert("No confirmed expenses found. Tick the Confirm? checkbox on at least one row first.");
    return;
  }

  const totals = computeTotals(expenses, names);
  const balancesMap = {};
  names.forEach(n => balancesMap[n] = totals[n].net);
  const allSettlements = calculateGreedySettlements(balancesMap);

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  let sentCount = 0, skippedCount = 0;

  roster.forEach((person, idx) => {
    const statusCell = reportsSheet.getRange(3 + idx, 6);
    try {
      // Skip people without a usable email instead of failing everyone
      if (!person.email || !emailRegex.test(person.email)) {
        statusCell.setValue("Skipped (no email)").setFontColor("#B45309");
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
    }
  });

  ui.alert(`Done! Sent: ${sentCount}, Skipped (no email): ${skippedCount}, Failed: ${roster.length - sentCount - skippedCount}.\n\nNote: emails are sent FROM the Google account that ran this function — check that account's Sent folder.`);
}
