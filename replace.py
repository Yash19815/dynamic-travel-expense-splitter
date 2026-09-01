"""Patch main.gs with performance optimizations for the expense splitter.

Usage:
    python optimize_expense_splitter.py
    python optimize_expense_splitter.py path/to/main.gs
    python optimize_expense_splitter.py path/to/main.gs -o main_optimized.gs

The script creates a timestamped backup next to the input file and writes an
optimized copy. It replaces these functions:
- onEdit
- applyRobotoFont
- updateExpensesStructure
- updateSendReportsSheet

It also changes the final updateSendReportsSheet() call in recalculateAll()
to avoid rebuilding reports repeatedly while expenses are unconfirmed.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path


ON_EDIT = r'''function onEdit(e) {
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
}'''

APPLY_ROBOTO = r'''function applyRobotoFont(sheet) {
  const range = sheet.getDataRange();
  if (range.getNumRows() && range.getNumColumns()) {
    range.setFontFamily("Roboto").setFontSize(12);
  }
}'''

UPDATE_EXPENSES = r'''function updateExpensesStructure() {
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
}'''

UPDATE_REPORTS = r'''function updateSendReportsSheet() {
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
}'''


def find_function_span(source: str, function_name: str) -> tuple[int, int]:
    """Return [start, end) character offsets for a top-level JS function."""
    match = re.search(rf"(?m)^function\s+{re.escape(function_name)}\s*\(", source)
    if not match:
        raise ValueError(f"Could not find function {function_name}().")

    start = match.start()
    brace_start = source.find("{", match.end())
    if brace_start < 0:
        raise ValueError(f"Could not find opening brace for {function_name}().")

    depth = 0
    in_single = in_double = in_template = False
    escape = False
    line_comment = block_comment = False
    i = brace_start

    while i < len(source):
        ch = source[i]
        nxt = source[i + 1] if i + 1 < len(source) else ""

        if line_comment:
            if ch == "\n":
                line_comment = False
            i += 1
            continue
        if block_comment:
            if ch == "*" and nxt == "/":
                block_comment = False
                i += 2
                continue
            i += 1
            continue
        if in_single:
            if not escape and ch == "'":
                in_single = False
            escape = (ch == "\\" and not escape)
            if ch != "\\":
                escape = False
            i += 1
            continue
        if in_double:
            if not escape and ch == '"':
                in_double = False
            escape = (ch == "\\" and not escape)
            if ch != "\\":
                escape = False
            i += 1
            continue
        if in_template:
            if not escape and ch == "`":
                in_template = False
            escape = (ch == "\\" and not escape)
            if ch != "\\":
                escape = False
            i += 1
            continue

        if ch == "/" and nxt == "/":
            line_comment = True
            i += 2
            continue
        if ch == "/" and nxt == "*":
            block_comment = True
            i += 2
            continue
        if ch == "'":
            in_single = True
            i += 1
            continue
        if ch == '"':
            in_double = True
            i += 1
            continue
        if ch == "`":
            in_template = True
            i += 1
            continue

        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return start, i + 1
        i += 1

    raise ValueError(f"Could not find closing brace for {function_name}().")


def replace_function(source: str, name: str, replacement: str) -> str:
    start, end = find_function_span(source, name)
    return source[:start] + replacement + source[end:]


def patch_final_report_call(source: str) -> str:
    """Make recalculateAll skip report refreshes while unconfirmed.

    Only changes the final direct call inside recalculateAll(), leaving calls in
    other functions untouched.
    """
    start, end = find_function_span(source, "recalculateAll")
    fn = source[start:end]
    old = "  updateSendReportsSheet();"
    new = "  // Avoid rebuilding the reports sheet after every unconfirmed expense edit.\n  if (confirmed || dataLastRow === 1) updateSendReportsSheet();"
    position = fn.rfind(old)
    if position < 0:
        raise ValueError("Could not find the final updateSendReportsSheet() call inside recalculateAll().")
    fn = fn[:position] + new + fn[position + len(old):]
    return source[:start] + fn + source[end:]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Optimize the travel expense splitter Apps Script locally.")
    parser.add_argument("input", nargs="?", default="main.gs", help="Input Apps Script file (default: main.gs)")
    parser.add_argument("-o", "--output", help="Output file. Default: <input-stem>_optimized.gs")
    parser.add_argument("--no-backup", action="store_true", help="Do not create a timestamped backup")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.is_file():
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 2

    output_path = Path(args.output).expanduser().resolve() if args.output else input_path.with_name(f"{input_path.stem}_optimized{input_path.suffix}")
    source = input_path.read_text(encoding="utf-8")

    required = ["onEdit", "applyRobotoFont", "updateExpensesStructure", "updateSendReportsSheet", "recalculateAll"]
    missing = [name for name in required if not re.search(rf"(?m)^function\s+{re.escape(name)}\s*\(", source)]
    if missing:
        print("Error: this does not look like the expected main_v7.gs file. Missing: " + ", ".join(missing), file=sys.stderr)
        return 3

    if not args.no_backup:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = input_path.with_name(f"{input_path.stem}.backup_{stamp}{input_path.suffix}")
        shutil.copy2(input_path, backup_path)
    else:
        backup_path = None

    try:
        patched = source
        patched = replace_function(patched, "onEdit", ON_EDIT)
        patched = replace_function(patched, "applyRobotoFont", APPLY_ROBOTO)
        patched = replace_function(patched, "updateExpensesStructure", UPDATE_EXPENSES)
        patched = replace_function(patched, "updateSendReportsSheet", UPDATE_REPORTS)
        patched = patch_final_report_call(patched)
    except ValueError as exc:
        print(f"Patch failed safely: {exc}", file=sys.stderr)
        return 4

    output_path.write_text(patched, encoding="utf-8", newline="\n")
    print(f"Optimized file created: {output_path}")
    if backup_path:
        print(f"Backup created:         {backup_path}")
    print("\nNext steps:")
    print("1. Review the optimized .gs file.")
    print("2. Replace Code.gs contents in Apps Script with it.")
    print("3. Save and reload the spreadsheet.")
    print("4. Do not run setup() unless you intentionally want to reset the tracker.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())