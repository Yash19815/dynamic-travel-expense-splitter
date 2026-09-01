/**
 * DYNAMIC GROUP TRAVEL EXPENSE SPLITTER
 * ------------------------------------------------------------------
 * Custom Menu: "Expense Splitter" -> "Initialize / Setup Tracker"
 */

const SHEET_SETUP = "Setup";
const SHEET_EXPENSES = "Expenses";
const SHEET_REPORTS = "Send Reports";

/**
 * Creates custom menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Splitter')
    .addItem('1. Initialize / Reset Tracker', 'setup')
    .addItem('2. Recalculate Expenses & Settlements', 'recalculateAll')
    .addItem('3. Send PDF Reports via Email', 'sendPdfReports')
    .addToUi();
}

/**
 * Master initialization function.
 * Builds all sheets, labels, roster tables, default settings, and formatting.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or reset sheets
  let setupSheet = ss.getSheetByName(SHEET_SETUP) || ss.insertSheet(SHEET_SETUP);
  let expenseSheet = ss.getSheetByName(SHEET_EXPENSES) || ss.insertSheet(SHEET_EXPENSES);
  let reportsSheet = ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);

  // --- 1. SETUP SHEET CONFIGURATION ---
  setupSheet.clear();
  setupSheet.getRange("A1:B1").merge().setValue("TRIP CONFIGURATION")
    .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");

  const configLabels = [
    ["Trip Name / Destination:", "Goa Vacation 2026"],
    ["Number of People:", 4],
    ["Currency Symbol:", "₹"],
    ["Setup Status:", "Checking..."]
  ];
  setupSheet.getRange("A2:B5").setValues(configLabels);
  setupSheet.getRange("A2:A5").setFontWeight("bold");

  // Number of people dropdown (2 to 15)
  const numPeopleRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(2, 15)
    .setAllowInvalid(false)
    .setHelpText("Please enter a number between 2 and 15.")
    .build();
  setupSheet.getRange("B3").setDataValidation(numPeopleRule);

  // Default Roster Table Header
  setupSheet.getRange("A7:C7").setValues([["#", "Name", "Email"]])
    .setBackground("#334155").setFontColor("#FFFFFF").setFontWeight("bold");

  // Seed default roster rows
  const initialRoster = [
    [1, "Rahul", "rahul@example.com"],
    [2, "Priya", "priya@example.com"],
    [3, "Amit", "amit@example.com"],
    [4, "Neha", "neha@example.com"]
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
  const range = sheet.getDataRange();
  range.setFontFamily("Roboto");
  range.setFontSize(12);
}

/**
 * Validates roster names & emails, updating cell B5 in Setup.
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
    statusCell.setValue("❌ Missing Trip Name").setFontColor("#DC2626");
    return false;
  }
  if (!currency) {
    statusCell.setValue("❌ Missing Currency Symbol").setFontColor("#DC2626");
    return false;
  }

  const rosterData = setupSheet.getRange(8, 1, numPeople, 3).getValues();
  const names = [];
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  for (let i = 0; i < numPeople; i++) {
    let name = String(rosterData[i][1]).trim();
    let email = String(rosterData[i][2]).trim();

    if (!name) {
      statusCell.setValue(`❌ Missing name for Person #${i+1}`).setFontColor("#DC2626");
      return false;
    }
    if (names.includes(name.toLowerCase())) {
      statusCell.setValue(`❌ Duplicate name found: "${name}"`).setFontColor("#DC2626");
      return false;
    }
    names.push(name.toLowerCase());

    if (!email || !emailRegex.test(email)) {
      statusCell.setValue(`❌ Invalid email for "${name}"`).setFontColor("#DC2626");
      return false;
    }
  }

  statusCell.setValue("Setup complete ✅").setFontColor("#166534").setFontWeight("bold");
  return true;
}

/**
 * Reads configured Roster array from Setup sheet.
 */
function getRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const numPeople = parseInt(setupSheet.getRange("B3").getValue(), 10) || 0;
  
  if (numPeople === 0) return [];
  
  const rawData = setupSheet.getRange(8, 1, numPeople, 3).getValues();
  return rawData.map(row => ({
    id: row[0],
    name: String(row[1]).trim(),
    email: String(row[2]).trim()
  }));
}

/**
 * Listens for edits to trigger dynamic updates in roster size or expense calculations.
 */
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_SETUP) {
    // If Number of People (B3) edited
    if (range.getA1Notation() === "B3") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
      if (expenseSheet && expenseSheet.getLastRow() > 1) {
        const confirm = Browser.msgBox("Warning", "Changing person count will adjust your expense headers. Continue?", Browser.Buttons.YES_NO);
        if (confirm !== "yes") return;
      }
      adjustRosterRows();
    }
    validateSetupStatus();
    updateExpensesStructure();
    recalculateAll();
  } else if (sheetName === SHEET_EXPENSES) {
    recalculateAll();
  }
}

/**
 * Adjusts the row count of the Roster table in Setup to match N (B3).
 */
function adjustRosterRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const targetN = parseInt(setupSheet.getRange("B3").getValue(), 10);
  
  const currentLastRow = setupSheet.getLastRow();
  const currentRosterCount = Math.max(0, currentLastRow - 7);

  if (targetN > currentRosterCount) {
    const rowsToAdd = targetN - currentRosterCount;
    const newRows = [];
    for (let i = 1; i <= rowsToAdd; i++) {
      const nextId = currentRosterCount + i;
      newRows.push([nextId, `Person ${nextId}`, `person${nextId}@example.com`]);
    }
    setupSheet.getRange(8 + currentRosterCount, 1, rowsToAdd, 3).setValues(newRows);
  } else if (targetN < currentRosterCount) {
    const rowsToDelete = currentRosterCount - targetN;
    setupSheet.deleteRows(8 + targetN, rowsToDelete);
  }

  applyRobotoFont(setupSheet);
  setupSheet.autoResizeColumns(1, 3);
}

/**
 * Rebuilds headers, data validations, and protections on the Expenses sheet.
 */
function updateExpensesStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const roster = getRoster();
  if (roster.length === 0) return;

  const names = roster.map(r => r.name);
  const headers = ["Date", "Description", "Amount", "Paid By", "Split Among", ...names];

  // Unprotect sheet temporarily to modify structure
  const protections = expenseSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => p.remove());

  // Set Headers
  expenseSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold");

  // Data Validation for "Paid By" column
  const paidByRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(false)
    .build();
  expenseSheet.getRange("D2:D500").setDataValidation(paidByRule);

  // Set column formats
  const currencySymbol = ss.getSheetByName(SHEET_SETUP).getRange("B4").getValue() || "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;
  
  expenseSheet.getRange("A2:A500").setNumberFormat("dd-mmm-yyyy");
  expenseSheet.getRange("C2:C500").setNumberFormat(currencyFormat);

  applyRobotoFont(expenseSheet);
  expenseSheet.autoResizeColumns(1, headers.length);
}

/**
 * Master recalculation engine for Expenses and Send Reports sheets.
 */
function recalculateAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const roster = getRoster();
  
  if (!expenseSheet || roster.length === 0) return;

  const currencySymbol = setupSheet.getRange("B4").getValue() || "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;

  const lastRow = expenseSheet.getLastRow();
  const names = roster.map(r => r.name);

  // Clear existing totals row if present
  let dataLastRow = lastRow;
  if (lastRow > 1) {
    const firstCellInLastRow = expenseSheet.getRange(lastRow, 1).getValue();
    if (String(firstCellInLastRow).toUpperCase() === "TOTAL") {
      expenseSheet.deleteRow(lastRow);
      dataLastRow = lastRow - 1;
    }
  }

  // Calculate per-person shares if there are data rows
  if (dataLastRow >= 2) {
    const expenseData = expenseSheet.getRange(2, 1, dataLastRow - 1, 5 + names.length).getValues();
    const computedShares = [];

    for (let i = 0; i < expenseData.length; i++) {
      const row = expenseData[i];
      const amount = parseFloat(row[2]) || 0;
      const paidBy = String(row[3]).trim();
      const splitStr = String(row[4]).trim();

      const rowNet = new Array(names.length).fill(0);

      if (amount > 0 && paidBy) {
        // Determine included people
        let included = [];
        if (!splitStr || splitStr.toLowerCase() === "all") {
          included = [...names];
        } else {
          const splitList = splitStr.split(",").map(s => s.trim().toLowerCase());
          included = names.filter(n => splitList.includes(n.toLowerCase()));
        }

        if (included.length > 0) {
          // Equal split with 2 decimal precision & remainder to payer
          const baseShare = Math.floor((amount / included.length) * 100) / 100;
          const totalBaseShares = baseShare * included.length;
          const remainder = Math.round((amount - totalBaseShares) * 100) / 100;

          names.forEach((name, idx) => {
            let consumedShare = 0;
            if (included.includes(name)) {
              consumedShare = baseShare + (name === paidBy ? remainder : 0);
            }
            const paidAmount = (name === paidBy) ? amount : 0;
            rowNet[idx] = paidAmount - consumedShare;
          });
        }
      }
      computedShares.push(rowNet);
    }

    // Write computed shares back to sheet
    if (computedShares.length > 0) {
      const shareRange = expenseSheet.getRange(2, 6, computedShares.length, names.length);
      shareRange.setValues(computedShares);
      shareRange.setNumberFormat(currencyFormat);
    }

    // Add Totals Row
    const totalsRowIndex = dataLastRow + 1;
    expenseSheet.getRange(totalsRowIndex, 1).setValue("TOTAL").setFontWeight("bold");
    expenseSheet.getRange(totalsRowIndex, 3).setFormula(`=SUM(C2:C${dataLastRow})`).setFontWeight("bold").setNumberFormat(currencyFormat);

    names.forEach((_, idx) => {
      const colLetter = String.fromCharCode(70 + idx); // F = 70
      const totalCell = expenseSheet.getRange(totalsRowIndex, 6 + idx);
      totalCell.setFormula(`=SUM(${colLetter}2:${colLetter}${dataLastRow})`)
        .setFontWeight("bold")
        .setNumberFormat(currencyFormat);
    });

    expenseSheet.getRange(totalsRowIndex, 1, 1, 5 + names.length)
      .setBackground("#E2E8F0")
      .setFontWeight("bold");
  }

  // Refresh protections on Expenses sheet
  protectExpensesSheet(expenseSheet, names.length);

  // Update Sheet 3 (Send Reports)
  updateSendReportsSheet();
}

/**
 * Protects formula columns and total rows on the Expenses sheet.
 */
function protectExpensesSheet(sheet, numPeople) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => p.remove());

  // Protect Header row
  const headerProt = sheet.getRange(1, 1, 1, 5 + numPeople).protect().setDescription("Header Protection");
  headerProt.removeEditors(headerProt.getEditors());

  // Protect calculated columns (Cols F onwards)
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const shareProt = sheet.getRange(2, 6, lastRow, numPeople).protect().setDescription("Per-person Formula Protection");
  shareProt.removeEditors(shareProt.getEditors());
}

/**
 * Populates Summary & Simplified Settlements on Sheet 3 ("Send Reports").
 */
function updateSendReportsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName(SHEET_REPORTS) || ss.insertSheet(SHEET_REPORTS);
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const roster = getRoster();

  reportsSheet.clear();
  const currencySymbol = ss.getSheetByName(SHEET_SETUP).getRange("B4").getValue() || "₹";
  const currencyFormat = `"${currencySymbol}"#,##0.00`;

  // --- 1. SUMMARY TABLE ---
  reportsSheet.getRange("A1:F1").merge().setValue("PERSONAL BALANCES SUMMARY")
    .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");

  const headers = [["Name", "Email", "Total Paid", "Total Share", "Net Balance", "Report Status"]];
  reportsSheet.getRange("A2:F2").setValues(headers)
    .setBackground("#334155").setFontColor("#FFFFFF").setFontWeight("bold");

  const summaryData = [];
  const balancesMap = {};

  const expenseLastRow = expenseSheet.getLastRow();
  const dataLastRow = (expenseLastRow > 1 && expenseSheet.getRange(expenseLastRow, 1).getValue() === "TOTAL") 
    ? expenseLastRow - 1 : expenseLastRow;

  roster.forEach((person, idx) => {
    let totalPaid = 0;
    let totalShare = 0;

    if (dataLastRow >= 2) {
      const expenses = expenseSheet.getRange(2, 1, dataLastRow - 1, 5 + roster.length).getValues();
      expenses.forEach(row => {
        const amt = parseFloat(row[2]) || 0;
        const paidBy = String(row[3]).trim();
        const splitStr = String(row[4]).trim();
        
        let included = [];
        if (!splitStr || splitStr.toLowerCase() === "all") {
          included = roster.map(r => r.name);
        } else {
          const splitList = splitStr.split(",").map(s => s.trim().toLowerCase());
          included = roster.map(r => r.name).filter(n => splitList.includes(n.toLowerCase()));
        }

        if (paidBy.toLowerCase() === person.name.toLowerCase()) {
          totalPaid += amt;
        }

        if (included.map(n => n.toLowerCase()).includes(person.name.toLowerCase())) {
          const baseShare = Math.floor((amt / included.length) * 100) / 100;
          const remainder = Math.round((amt - (baseShare * included.length)) * 100) / 100;
          totalShare += baseShare + (paidBy.toLowerCase() === person.name.toLowerCase() ? remainder : 0);
        }
      });
    }

    const netBalance = totalPaid - totalShare;
    balancesMap[person.name] = netBalance;
    summaryData.push([person.name, person.email, totalPaid, totalShare, netBalance, "Not Sent"]);
  });

  if (summaryData.length > 0) {
    const summaryRange = reportsSheet.getRange(3, 1, summaryData.length, 6);
    summaryRange.setValues(summaryData);
    
    // Formatting numbers
    reportsSheet.getRange(3, 3, summaryData.length, 3).setNumberFormat(currencyFormat);
    
    // Color coding Net Balances
    for (let i = 0; i < summaryData.length; i++) {
      const cell = reportsSheet.getRange(3 + i, 5);
      const val = summaryData[i][4];
      if (val > 0.01) cell.setBackground("#D1FAE5").setFontColor("#065F46"); // Green
      else if (val < -0.01) cell.setBackground("#FEE2E2").setFontColor("#991B1B"); // Red
      else cell.setBackground("#F3F4F6").setFontColor("#374151");
    }
  }

  // --- 2. SIMPLIFIED SETTLEMENTS (GREEDY ALGORITHM) ---
  const settlements = calculateGreedySettlements(balancesMap);

  const startSettlementRow = summaryData.length + 5;
  reportsSheet.getRange(startSettlementRow, 1, 1, 4).merge().setValue("SIMPLIFIED SETTLEMENT PLAN")
    .setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold");

  reportsSheet.getRange(startSettlementRow + 1, 1, 1, 4)
    .setValues([["From (Debtor)", "To (Creditor)", "Amount", "Instruction"]])
    .setBackground("#115E59").setFontColor("#FFFFFF").setFontWeight("bold");

  if (settlements.length > 0) {
    const settlementRows = settlements.map(s => [
      s.from,
      s.to,
      s.amount,
      `${s.from} pays ${s.to} ${currencySymbol}${s.amount.toFixed(2)}`
    ]);
    const setRange = reportsSheet.getRange(startSettlementRow + 2, 1, settlementRows.length, 4);
    setRange.setValues(settlementRows);
    reportsSheet.getRange(startSettlementRow + 2, 3, settlementRows.length, 1).setNumberFormat(currencyFormat);
  } else {
    reportsSheet.getRange(startSettlementRow + 2, 1, 1, 4).merge()
      .setValue("All debts settled! Net balance is 0 for everyone.").setFontStyle("italic");
  }

  applyRobotoFont(reportsSheet);
  reportsSheet.autoResizeColumns(1, 6);
}

/**
 * Greedy algorithm to minimize total settlement transactions.
 */
function calculateGreedySettlements(balances) {
  let debtors = [];
  let creditors = [];

  for (let person in balances) {
    let bal = Math.round(balances[person] * 100) / 100;
    if (bal < -0.01) debtors.push({ name: person, amount: Math.abs(bal) });
    else if (bal > 0.01) creditors.push({ name: person, amount: bal });
  }

  debtors.sort((a, b) => b.amount - a.amount);
  creditors.sort((a, b) => b.amount - a.amount);

  let settlements = [];
  let d = 0, c = 0;

  while (d < debtors.length && c < creditors.length) {
    let debtor = debtors[d];
    let creditor = creditors[c];
    let settlementAmount = Math.min(debtor.amount, creditor.amount);

    settlements.push({
      from: debtor.name,
      to: creditor.name,
      amount: Math.round(settlementAmount * 100) / 100
    });

    debtor.amount = Math.round((debtor.amount - settlementAmount) * 100) / 100;
    creditor.amount = Math.round((creditor.amount - settlementAmount) * 100) / 100;

    if (debtor.amount === 0) d++;
    if (creditor.amount === 0) c++;
  }

  return settlements;
}

/**
 * Generates custom HTML-based PDF reports for each person and emails them out.
 */
function sendPdfReports() {
  const ui = SpreadsheetApp.getUi();
  if (!validateSetupStatus()) {
    ui.alert("Cannot send reports: Setup is incomplete. Check the 'Setup' tab status cell.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName(SHEET_SETUP);
  const expenseSheet = ss.getSheetByName(SHEET_EXPENSES);
  const reportsSheet = ss.getSheetByName(SHEET_REPORTS);
  
  const tripName = setupSheet.getRange("B2").getValue();
  const currencySymbol = setupSheet.getRange("B4").getValue() || "₹";
  const roster = getRoster();

  // Get all expenses
  const expenseLastRow = expenseSheet.getLastRow();
  const dataLastRow = (expenseLastRow > 1 && expenseSheet.getRange(expenseLastRow, 1).getValue() === "TOTAL") 
    ? expenseLastRow - 1 : expenseLastRow;

  const expenses = (dataLastRow >= 2) 
    ? expenseSheet.getRange(2, 1, dataLastRow - 1, 5 + roster.length).getValues() 
    : [];

  // Get current global settlements map
  const balancesMap = {};
  roster.forEach(r => {
    const cellVal = reportsSheet.getRange(3 + roster.findIndex(x => x.name === r.name), 5).getValue();
    balancesMap[r.name] = parseFloat(cellVal) || 0;
  });
  const allSettlements = calculateGreedySettlements(balancesMap);

  const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy");

  roster.forEach((person, idx) => {
    try {
      // 1. Filter relevant expenses for this person
      const relevantExpenses = [];
      let totalPaid = 0;
      let totalShare = 0;

      expenses.forEach(row => {
        const dateVal = row[0] instanceof Date 
          ? Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy")
          : String(row[0]);
        const desc = row[1];
        const amt = parseFloat(row[2]) || 0;
        const paidBy = String(row[3]).trim();
        const splitStr = String(row[4]).trim();

        let included = [];
        if (!splitStr || splitStr.toLowerCase() === "all") {
          included = roster.map(r => r.name);
        } else {
          const splitList = splitStr.split(",").map(s => s.trim().toLowerCase());
          included = roster.map(r => r.name).filter(n => splitList.includes(n.toLowerCase()));
        }

        const isPayer = paidBy.toLowerCase() === person.name.toLowerCase();
        const isIncluded = included.map(n => n.toLowerCase()).includes(person.name.toLowerCase());

        if (isPayer || isIncluded) {
          const baseShare = Math.floor((amt / included.length) * 100) / 100;
          const remainder = Math.round((amt - (baseShare * included.length)) * 100) / 100;
          const myShare = isIncluded ? (baseShare + (isPayer ? remainder : 0)) : 0;

          if (isPayer) totalPaid += amt;
          if (isIncluded) totalShare += myShare;

          relevantExpenses.push({
            date: dateVal,
            desc: desc,
            amount: amt,
            paidBy: paidBy,
            myShare: myShare
          });
        }
      });

      const netBalance = totalPaid - totalShare;

      // 2. Build personalized settlements text
      const myPersonalSettlements = allSettlements.filter(s => s.from === person.name || s.to === person.name);

      // 3. Construct HTML Document for PDF Blob
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
              <div class="card"><div class="card-title">Total Paid</div><div class="card-val">${currencySymbol}${totalPaid.toFixed(2)}</div></div>
              <div class="card"><div class="card-title">Total Share</div><div class="card-val">${currencySymbol}${totalShare.toFixed(2)}</div></div>
              <div class="card"><div class="card-title">Net Balance</div><div class="card-val ${netBalance >= 0 ? 'badge-green' : 'badge-red'}">${currencySymbol}${netBalance.toFixed(2)}</div></div>
            </div>

            <h4>Your Relevant Expenses</h4>
            <table>
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Description</th>
                  <th>Total Amount</th>
                  <th>Paid By</th>
                  <th>Your Share</th>
                </tr>
              </thead>
              <tbody>
                ${relevantExpenses.map(e => `
                  <tr>
                    <td>${e.date}</td>
                    <td>${e.desc}</td>
                    <td>${currencySymbol}${e.amount.toFixed(2)}</td>
                    <td>${e.paidBy}</td>
                    <td>${currencySymbol}${e.myShare.toFixed(2)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>

            <div class="box">
              <h4 style="margin-top:0;">Your Settlement Action Items</h4>
              ${myPersonalSettlements.length > 0 ? `
                <ul>
                  ${myPersonalSettlements.map(s => `
                    <li>${s.from === person.name 
                      ? `You owe <b>${s.to}</b> <b>${currencySymbol}${s.amount.toFixed(2)}</b>` 
                      : `<b>${s.from}</b> owes you <b>${currencySymbol}${s.amount.toFixed(2)}</b>`}</li>
                  `).join('')}
                </ul>
              ` : '<p>You have no pending settlement actions!</p>'}
            </div>

            <h4 style="margin-top:25px;">Full Group Settlement Plan (For Transparency)</h4>
            <ul>
              ${allSettlements.map(s => `<li>${s.from} pays ${s.to} ${currencySymbol}${s.amount.toFixed(2)}</li>`).join('')}
            </ul>
          </body>
        </html>
      `;

      // 4. Convert HTML directly to PDF Blob (Zero garbage files in Drive)
      const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
        .getAs('application/pdf')
        .setName(`${tripName} - ${person.name} Expense Report.pdf`);

      // 5. Email PDF to Person
      const subject = `[${tripName}] — Your Expense Report`;
      const body = `Hi ${person.name},\n\nHere is your itemized expense report for "${tripName}". Please find your PDF breakdown attached.\n\nBest regards,\nGroup Expense Splitter`;

      MailApp.sendEmail(person.email, subject, body, {
        attachments: [pdfBlob]
      });

      // 6. Log success timestamp in Sheet 3
      const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMM HH:mm");
      reportsSheet.getRange(3 + idx, 6).setValue(`Sent (${timestamp})`).setFontColor("#166534");

    } catch (err) {
      reportsSheet.getRange(3 + idx, 6).setValue(`Failed: ${err.message}`).setFontColor("#DC2626");
    }
  });

  ui.alert("Expense reports have been generated and emailed to all participants!");
}