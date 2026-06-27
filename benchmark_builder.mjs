import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
const workbookName = "benchmark_complex_lbo_dcf_stress_test.xlsx";
const truthSetName = "benchmark_complex_lbo_dcf_truth_set.json";
const periods = 84;
const startCol = 2; // B

const workbook = Workbook.create();
const sheets = {};
for (const name of [
  "Control",
  "Revenue Build",
  "Operating Forecast",
  "Working Capital",
  "Debt Schedule",
  "Income Statement",
  "Cash Flow",
  "Balance Sheet",
  "Valuation",
  "Returns Summary",
  "Checks",
  "Sensitivity",
]) {
  sheets[name] = workbook.worksheets.add(name);
}

const truthSet = [];

function colLabel(n) {
  let s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function addr(col, row) {
  return `${colLabel(col)}${row}`;
}

function qSheet(name) {
  return /[^A-Za-z0-9_]/.test(name) ? `'${name}'` : name;
}

function sheetAddr(sheet, col, row) {
  return `${qSheet(sheet)}!${addr(col, row)}`;
}

function monthEndDates(count) {
  const dates = [];
  let d = new Date("2024-01-31T00:00:00");
  for (let i = 0; i < count; i += 1) {
    dates.push(new Date(d));
    d = new Date(d.getFullYear(), d.getMonth() + 2, 0);
  }
  return dates;
}

function fillDates(sheet) {
  const dates = monthEndDates(periods);
  sheet.getRange(`B2:${colLabel(startCol + periods - 1)}2`).values = [dates];
}

function writeLabels(sheet, startRow, labels) {
  sheet.getRange(`A${startRow}:A${startRow + labels.length - 1}`).values = labels.map((label) => [label]);
}

function writeMonthlyBlock(sheet, startRow, rowDefs) {
  writeLabels(sheet, startRow, rowDefs.map((row) => row.label));
  const endCol = colLabel(startCol + periods - 1);
  const formulas = rowDefs.map((rowDef) =>
    Array.from({ length: periods }, (_, idx) => {
      const col = startCol + idx;
      return rowDef.formula(col, idx);
    }),
  );
  sheet.getRange(`B${startRow}:${endCol}${startRow + rowDefs.length - 1}`).formulas = formulas;
}

function writeMonthlyValues(sheet, startRow, rowDefs) {
  writeLabels(sheet, startRow, rowDefs.map((row) => row.label));
  const endCol = colLabel(startCol + periods - 1);
  const values = rowDefs.map((rowDef) =>
    Array.from({ length: periods }, (_, idx) => rowDef.value(idx)),
  );
  sheet.getRange(`B${startRow}:${endCol}${startRow + rowDefs.length - 1}`).values = values;
}

function addIssue(type, cell, detail, detectorHint) {
  truthSet.push({ type, cell, detail, detectorHint });
}

for (const sheet of Object.values(sheets)) {
  fillDates(sheet);
  sheet.getRange("A1").values = [[sheet.name]];
}

const control = sheets["Control"];
control.getRange("A4:B23").values = [
  ["Annual Revenue Growth", 0.12],
  ["Annual Price Growth", 0.015],
  ["Returns % of Gross Revenue", 0.03],
  ["COGS % of Revenue", 0.42],
  ["S&M % of Revenue", 0.18],
  ["R&D % of Revenue", 0.12],
  ["G&A % of Revenue", 0.08],
  ["Payroll Inflation", 0.05],
  ["DSO", 45],
  ["Inventory Days", 18],
  ["AP Days", 40],
  ["Interest Rate", 0.09],
  ["Tax Rate", 0.24],
  ["WACC", 0.11],
  ["Terminal Growth", 0.03],
  ["Exit Multiple", 8.0],
  ["Initial Debt", 250],
  ["Sponsor Equity", 120],
  ["Opening Cash", 25],
  ["Monthly Capex", 4],
];

writeMonthlyBlock(sheets["Revenue Build"], 4, [
  {
    label: "Units Sold",
    formula: (col) => (col === startCol ? "=10000" : `=${addr(col - 1, 4)}*(1+${sheetAddr("Control", 2, 4)}/12)`),
  },
  {
    label: "Price / Unit",
    formula: (col) => (col === startCol ? "=52" : `=${addr(col - 1, 5)}*(1+${sheetAddr("Control", 2, 5)}/12)`),
  },
  { label: "Gross Revenue", formula: (col) => `=${addr(col, 4)}*${addr(col, 5)}` },
  { label: "Returns %", formula: () => `=${sheetAddr("Control", 2, 6)}` },
  { label: "Returns", formula: (col) => `=${addr(col, 6)}*${addr(col, 7)}` },
  { label: "Net Product Revenue", formula: (col) => `=${addr(col, 6)}-${addr(col, 8)}` },
  {
    label: "Implementation Revenue",
    formula: (col, idx) => `=${idx < 12 ? 120 : 180}+ROUND(${addr(col, 4)}/150,0)`,
  },
  { label: "Total Revenue", formula: (col) => `=${addr(col, 9)}+${addr(col, 10)}` },
  { label: "COGS %", formula: () => `=${sheetAddr("Control", 2, 7)}` },
  { label: "COGS", formula: (col) => `=${addr(col, 11)}*${addr(col, 12)}` },
  { label: "Gross Profit", formula: (col) => `=${addr(col, 11)}-${addr(col, 13)}` },
  { label: "Gross Margin", formula: (col) => `=${addr(col, 14)}/${addr(col, 11)}` },
]);

writeMonthlyBlock(sheets["Operating Forecast"], 4, [
  {
    label: "Headcount",
    formula: (col) => (col === startCol ? "=85" : `=${addr(col - 1, 4)}+IF(MOD(COLUMN(),6)=0,2,1)`),
  },
  {
    label: "Avg Salary",
    formula: (col) => (col === startCol ? "=9.5" : `=${addr(col - 1, 5)}*(1+${sheetAddr("Control", 2, 11)}/12)`),
  },
  { label: "Payroll", formula: (col) => `=${addr(col, 4)}*${addr(col, 5)}` },
  { label: "S&M %", formula: () => `=${sheetAddr("Control", 2, 8)}` },
  { label: "S&M Expense", formula: (col) => `='Revenue Build'!${addr(col, 11)}*${addr(col, 7)}` },
  { label: "R&D %", formula: () => `=${sheetAddr("Control", 2, 9)}` },
  { label: "R&D Expense", formula: (col) => `='Revenue Build'!${addr(col, 11)}*${addr(col, 9)}` },
  { label: "G&A %", formula: () => `=${sheetAddr("Control", 2, 10)}` },
  { label: "G&A Expense", formula: (col) => `='Revenue Build'!${addr(col, 11)}*${addr(col, 11)}` },
  { label: "Hosting", formula: (col, idx) => `=${20 + idx * 0.2}+('Revenue Build'!${addr(col, 4)}/1000)` },
  { label: "Rent", formula: (col) => (col === startCol ? "=35" : `=${addr(col - 1, 14)}`) },
  { label: "Total Opex", formula: (col) => `=SUM(${addr(col, 6)}:${addr(col, 14)})` },
  { label: "EBITDA", formula: (col) => `='Revenue Build'!${addr(col, 14)}-${addr(col, 15)}` },
  { label: "EBITDA Margin", formula: (col) => `=${addr(col, 16)}/'Revenue Build'!${addr(col, 11)}` },
  { label: "Exception Ratio", formula: (col) => `=${addr(col, 16)}/'Revenue Build'!${addr(col, 11)}` },
  { label: "Capex", formula: (col) => `=${sheetAddr("Control", 2, 23)}+IF(MOD(COLUMN(),12)=0,2,0)` },
  { label: "Depreciation", formula: (col) => (col === startCol ? `=${addr(col, 19)}/24` : `=${addr(col - 1, 20)}*0.92+${addr(col, 19)}/24`) },
  { label: "EBIT", formula: (col) => `=${addr(col, 16)}-${addr(col, 20)}` },
]);

// Seed divide by zero
sheets["Operating Forecast"].getRange("N18").formulas = [["=N16/0"]];
addIssue("divide_by_zero", "Operating Forecast!N18", "Literal divide by zero in exception ratio", "divide_by_zero");

// Seed hardcoded repeated constants
for (const cell of ["Q19", "R19", "S19", "T19"]) {
  sheets["Operating Forecast"].getRange(cell).formulas = [[`=${cell.replace("19", "13")}*0+25`]];
  addIssue("hardcoded_number", `Operating Forecast!${cell}`, "Hardcoded 25 embedded in capex line", "hardcoded_numbers");
}

writeMonthlyBlock(sheets["Working Capital"], 4, [
  { label: "DSO", formula: () => `=${sheetAddr("Control", 2, 12)}` },
  { label: "Accounts Receivable", formula: (col) => `='Revenue Build'!${addr(col, 11)}*${addr(col, 4)}/30` },
  { label: "Inventory Days", formula: () => `=${sheetAddr("Control", 2, 13)}` },
  { label: "Inventory", formula: (col) => `='Revenue Build'!${addr(col, 13)}*${addr(col, 6)}/30` },
  { label: "AP Days", formula: () => `=${sheetAddr("Control", 2, 14)}` },
  { label: "Accounts Payable", formula: (col) => `='Revenue Build'!${addr(col, 13)}*${addr(col, 8)}/30` },
  { label: "Deferred Revenue", formula: (col) => `='Revenue Build'!${addr(col, 10)}*0.4` },
  { label: "Net Working Capital", formula: (col) => `=${addr(col, 5)}+${addr(col, 7)}-${addr(col, 9)}-${addr(col, 10)}` },
  {
    label: "Change in NWC",
    formula: (col) => (col === startCol ? `=${addr(col, 11)}` : `=${addr(col, 11)}-${addr(col - 1, 11)}`),
  },
]);

writeMonthlyBlock(sheets["Debt Schedule"], 4, [
  { label: "Beginning Balance", formula: (col) => (col === startCol ? `=${sheetAddr("Control", 2, 20)}` : `=${addr(col - 1, 10)}`) },
  { label: "Mandatory Amort %", formula: () => "=0.005" },
  { label: "Mandatory Amortization", formula: (col) => `=${addr(col, 4)}*${addr(col, 5)}` },
  { label: "Cash Sweep %", formula: () => "=0.30" },
  { label: "Cash Sweep", formula: (col) => `=MAX(0,'Cash Flow'!${addr(col, 8)}*${addr(col, 7)})` },
  { label: "Optional Repayment", formula: () => "=0" },
  { label: "Ending Balance", formula: (col) => `=${addr(col, 4)}-${addr(col, 6)}-${addr(col, 8)}-${addr(col, 9)}` },
  { label: "Interest Rate", formula: () => `=${sheetAddr("Control", 2, 15)}` },
  { label: "Interest Expense", formula: (col) => `=AVERAGE(${addr(col, 4)},${addr(col, 10)})*${addr(col, 11)}/12` },
  { label: "Leverage Ratio", formula: (col) => `=${addr(col, 10)}/MAX(1,'Operating Forecast'!${addr(col, 16)}*12)` },
  { label: "Coverage Ratio", formula: (col) => `='Operating Forecast'!${addr(col, 16)}/${addr(col, 12)}` },
]);

// Seed hidden formula inconsistency in debt roll-forward
sheets["Debt Schedule"].getRange("AZ10").formulas = [["=AZ4-AZ6+AZ8-AZ9"]];
addIssue("formula_inconsistency", "Debt Schedule!AZ10", "Ending balance adds cash sweep instead of subtracting it", "formula_inconsistency");

writeMonthlyBlock(sheets["Income Statement"], 4, [
  { label: "Revenue", formula: (col) => `='Revenue Build'!${addr(col, 11)}` },
  { label: "COGS", formula: (col) => `='Revenue Build'!${addr(col, 13)}` },
  { label: "Gross Profit", formula: (col) => `='Revenue Build'!${addr(col, 14)}` },
  { label: "Payroll", formula: (col) => `='Operating Forecast'!${addr(col, 6)}` },
  { label: "S&M", formula: (col) => `='Operating Forecast'!${addr(col, 8)}` },
  { label: "R&D", formula: (col) => `='Operating Forecast'!${addr(col, 10)}` },
  { label: "G&A", formula: (col) => `='Operating Forecast'!${addr(col, 12)}` },
  { label: "Hosting", formula: (col) => `='Operating Forecast'!${addr(col, 13)}` },
  { label: "Rent", formula: (col) => `='Operating Forecast'!${addr(col, 14)}` },
  { label: "Total Opex", formula: (col) => `=SUM(${addr(col, 7)}:${addr(col, 12)})` },
  { label: "EBITDA", formula: (col) => `=${addr(col, 6)}-${addr(col, 13)}` },
  { label: "Depreciation", formula: (col) => `='Operating Forecast'!${addr(col, 20)}` },
  { label: "EBIT", formula: (col) => `=${addr(col, 14)}-${addr(col, 15)}` },
  { label: "Interest Expense", formula: (col) => `='Debt Schedule'!${addr(col, 12)}` },
  { label: "EBT", formula: (col) => `=${addr(col, 16)}-${addr(col, 17)}` },
  { label: "Taxes", formula: (col) => `=MAX(0,${addr(col, 18)}*${sheetAddr("Control", 2, 16)})` },
  { label: "Net Income", formula: (col) => `=${addr(col, 18)}-${addr(col, 19)}` },
]);

writeMonthlyBlock(sheets["Cash Flow"], 4, [
  { label: "Net Income", formula: (col) => `='Income Statement'!${addr(col, 20)}` },
  { label: "Depreciation", formula: (col) => `='Income Statement'!${addr(col, 15)}` },
  { label: "Change in NWC", formula: (col) => `=-'Working Capital'!${addr(col, 12)}` },
  { label: "Capex", formula: (col) => `=-'Operating Forecast'!${addr(col, 19)}` },
  { label: "Unlevered FCF", formula: (col) => `=SUM(${addr(col, 4)}:${addr(col, 7)})` },
  { label: "Interest Expense", formula: (col) => `=-'Debt Schedule'!${addr(col, 12)}` },
  { label: "Mandatory Amortization", formula: (col) => `=-'Debt Schedule'!${addr(col, 6)}` },
  { label: "Cash Sweep", formula: (col) => `=-'Debt Schedule'!${addr(col, 8)}` },
  { label: "Net Debt Change", formula: (col) => `=SUM(${addr(col, 9)}:${addr(col, 11)})` },
  { label: "Beginning Cash", formula: (col) => (col === startCol ? `=${sheetAddr("Control", 2, 22)}` : `=${addr(col - 1, 14)}`) },
  { label: "Ending Cash", formula: (col) => `=${addr(col, 13)}+${addr(col, 8)}+${addr(col, 12)}` },
]);

writeMonthlyBlock(sheets["Balance Sheet"], 4, [
  { label: "Cash", formula: (col) => `='Cash Flow'!${addr(col, 14)}` },
  { label: "Accounts Receivable", formula: (col) => `='Working Capital'!${addr(col, 5)}` },
  { label: "Inventory", formula: (col) => `='Working Capital'!${addr(col, 7)}` },
  {
    label: "PP&E",
    formula: (col) => (col === startCol ? "=60+'Operating Forecast'!B19-'Operating Forecast'!B20" : `=${addr(col - 1, 7)}+'Operating Forecast'!${addr(col, 19)}-'Operating Forecast'!${addr(col, 20)}`),
  },
  { label: "Other Assets", formula: () => "=10" },
  { label: "Total Assets", formula: (col) => `=SUM(${addr(col, 4)}:${addr(col, 8)})` },
  { label: "Debt", formula: (col) => `='Debt Schedule'!${addr(col, 10)}` },
  { label: "Accounts Payable", formula: (col) => `='Working Capital'!${addr(col, 9)}` },
  { label: "Deferred Revenue", formula: (col) => `='Working Capital'!${addr(col, 10)}` },
  { label: "Other Liabilities", formula: () => "=8" },
  {
    label: "Shareholder Equity",
    formula: (col) => (col === startCol ? `=${sheetAddr("Control", 2, 21)}+'Income Statement'!${addr(col, 20)}` : `=${addr(col - 1, 14)}+'Income Statement'!${addr(col, 20)}`),
  },
  { label: "Total Liab & Equity", formula: (col) => `=SUM(${addr(col, 10)}:${addr(col, 14)})` },
  { label: "Balance Check", formula: (col) => `=${addr(col, 9)}+${addr(col, 15)}` },
]);
addIssue("balance_check_sign", "Balance Sheet!B16:CG16", "Balance check adds assets and liabilities instead of subtracting", "balance_check");

writeMonthlyBlock(sheets["Valuation"], 4, [
  { label: "Unlevered FCF", formula: (col) => `='Cash Flow'!${addr(col, 8)}` },
  { label: "Discount Rate", formula: () => `=${sheetAddr("Control", 2, 17)}` },
  { label: "Discount Factor", formula: (col, idx) => `=1/(1+${addr(col, 5)}/12)^${idx + 1}` },
  { label: "PV of FCF", formula: (col) => `=${addr(col, 4)}*${addr(col, 6)}` },
  { label: "Terminal Growth", formula: () => `=${sheetAddr("Control", 2, 18)}` },
  { label: "Terminal Value", formula: (col) => `=(${addr(col, 4)}*12*(1+${addr(col, 8)}))/(${addr(col, 5)}-${addr(col, 8)})` },
  { label: "PV Terminal Value", formula: (col) => `=${addr(col, 9)}*${addr(col, 6)}` },
  { label: "Enterprise Value", formula: (col) => `=${addr(col, 7)}+${addr(col, 10)}` },
  { label: "Net Debt", formula: (col) => `='Debt Schedule'!${addr(col, 10)}-'Cash Flow'!${addr(col, 14)}` },
  { label: "Equity Value", formula: (col) => `=${addr(col, 11)}-${addr(col, 12)}` },
]);

// Seed Gordon growth sign issue
sheets["Valuation"].getRange("CG9").formulas = [["=(CG4*12*(1+CG8))/(CG5+CG8)"]];
addIssue("terminal_value_sign", "Valuation!CG9", "Terminal value uses r + g instead of r - g", "label_semantics");

writeMonthlyBlock(sheets["Returns Summary"], 4, [
  { label: "Sponsor Equity Invested", formula: () => `=${sheetAddr("Control", 2, 21)}` },
  { label: "Debt Funded", formula: () => `=${sheetAddr("Control", 2, 20)}` },
  { label: "Exit Equity Value", formula: () => `='Valuation'!CG13` },
  { label: "MOIC", formula: () => "=B6/B4" },
  { label: "IRR", formula: () => "=(B7/B4)^(1/5)-1" },
  { label: "Cash-on-Cash", formula: () => "=B6/B4" },
]);
addIssue("irr_label_mismatch", "Returns Summary!B8", "IRR row does not use IRR/XIRR/MIRR", "label_semantics");

writeMonthlyBlock(sheets["Checks"], 4, [
  { label: "Revenue Check", formula: (col) => `='Revenue Build'!${addr(col, 11)}-'Income Statement'!${addr(col, 4)}` },
  { label: "Cash Flow Check", formula: (col) => `='Cash Flow'!${addr(col, 14)}-'Balance Sheet'!${addr(col, 4)}` },
  { label: "Balance Check", formula: (col) => `='Balance Sheet'!${addr(col, 9)}` },
  { label: "Debt Check", formula: (col) => `='Debt Schedule'!${addr(col, 10)}-'Balance Sheet'!${addr(col, 10)}` },
  { label: "Sources & Uses Check", formula: () => `=${sheetAddr("Control", 2, 20)}+${sheetAddr("Control", 2, 21)}-370` },
]);

writeLabels(sheets["Sensitivity"], 4, ["WACC / Exit", "7.0x", "8.0x", "9.0x", "10.0x", "11.0x", "Bad Label"]);
sheets["Sensitivity"].getRange("B4:F4").values = [[0.09, 0.10, 0.11, 0.12, 0.13]];
for (let r = 5; r <= 9; r += 1) {
  for (let c = 2; c <= 6; c += 1) {
    sheets["Sensitivity"].getRange(addr(c, r)).formulas = [[`=('Valuation'!CG13*(1+${addr(c, 4)}-${sheetAddr("Control", 2, 17)})*(${addr(1, r)}+0))/1`]];
  }
}
sheets["Sensitivity"].getRange("A10:B11").values = [["Year 1", "Text Divider Test"], ["Metric", null]];
sheets["Sensitivity"].getRange("B11").formulas = [["='Valuation'!CG13/A10"]];
addIssue("text_math", "Sensitivity!B11", "Formula divides by literal text label Year 1", "text_math");

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/${workbookName}`);
await fs.writeFile(`${outputDir}/${truthSetName}`, JSON.stringify(truthSet, null, 2));
console.log(JSON.stringify({ workbook: `${outputDir}/${workbookName}`, truthSet: `${outputDir}/${truthSetName}`, seededIssues: truthSet.length }, null, 2));
