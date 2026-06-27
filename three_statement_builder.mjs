import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
const workbookName = "benchmark_three_statement_operating_stress_test.xlsx";
const truthSetName = "benchmark_three_statement_operating_truth_set.json";
const periods = 96;
const startCol = 2;

const workbook = Workbook.create();
const sheetNames = [
  "Control",
  "Revenue Drivers",
  "Operating Model",
  "Working Capital",
  "Income Statement",
  "Cash Flow",
  "Balance Sheet",
  "Debt Schedule",
  "Equity Bridge",
  "Valuation",
  "Checks",
  "Sensitivity",
];
const sheets = Object.fromEntries(sheetNames.map((name) => [name, workbook.worksheets.add(name)]));
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
function sAddr(sheet, col, row) {
  return `${qSheet(sheet)}!${addr(col, row)}`;
}
function addIssue(type, cell, detail, detectorHint) {
  truthSet.push({ type, cell, detail, detectorHint });
}
function monthlyDates(count) {
  const dates = [];
  let d = new Date("2025-01-31T00:00:00");
  for (let i = 0; i < count; i += 1) {
    dates.push(new Date(d));
    d = new Date(d.getFullYear(), d.getMonth() + 2, 0);
  }
  return dates;
}
function setupSheet(sheet) {
  sheet.getRange("A1").values = [[sheet.name]];
  sheet.getRange(`B2:${colLabel(startCol + periods - 1)}2`).values = [monthlyDates(periods)];
}
for (const sheet of Object.values(sheets)) setupSheet(sheet);

function writeLabels(sheet, startRow, labels) {
  sheet.getRange(`A${startRow}:A${startRow + labels.length - 1}`).values = labels.map((x) => [x]);
}
function writeFormulas(sheet, startRow, rowDefs) {
  writeLabels(sheet, startRow, rowDefs.map((r) => r.label));
  const endCol = colLabel(startCol + periods - 1);
  const formulas = rowDefs.map((rowDef) =>
    Array.from({ length: periods }, (_, idx) => rowDef.formula(startCol + idx, idx)),
  );
  sheet.getRange(`B${startRow}:${endCol}${startRow + rowDefs.length - 1}`).formulas = formulas;
}

sheets["Control"].getRange("A4:B27").values = [
  ["Customer Growth", 0.10],
  ["ARPU Growth", 0.03],
  ["Churn %", 0.02],
  ["COGS % Revenue", 0.38],
  ["S&M % Revenue", 0.17],
  ["R&D % Revenue", 0.11],
  ["G&A % Revenue", 0.08],
  ["Payroll Inflation", 0.045],
  ["DSO", 42],
  ["Inventory Days", 22],
  ["AP Days", 39],
  ["Tax Rate", 0.25],
  ["Interest Rate", 0.082],
  ["WACC", 0.10],
  ["Exit Multiple", 11],
  ["Opening Cash", 35],
  ["Opening Debt", 140],
  ["Share Count", 52],
  ["Maintenance Capex", 3.5],
  ["Depreciation Life", 36],
  ["Deferred Revenue %", 0.09],
  ["Terminal Growth", 0.025],
  ["Repurchase Rate", 0.004],
  ["Other Assets", 12],
];

writeFormulas(sheets["Revenue Drivers"], 4, [
  { label: "Beginning Customers", formula: (c, i) => `${i === 0 ? "=1200" : `=${addr(c - 1, 7)}`}` },
  { label: "New Customers", formula: (c, i) => `${i === 0 ? "=85" : `=${addr(c - 1, 5)}*(1+${sAddr("Control", 2, 5)}/12)`}` },
  { label: "Churn %", formula: () => `=${sAddr("Control", 2, 7)}` },
  { label: "Churned Customers", formula: (c) => `=${addr(c, 4)}*${addr(c, 6)}` },
  { label: "Ending Customers", formula: (c) => `=${addr(c, 4)}+${addr(c, 5)}-${addr(c, 7)}` },
  { label: "ARPU", formula: (c, i) => `${i === 0 ? "=2100" : `=${addr(c - 1, 9)}*(1+${sAddr("Control", 2, 6)}/12)`}` },
  { label: "Subscription Revenue", formula: (c) => `=${addr(c, 8)}*${addr(c, 9)}` },
  { label: "Services Revenue", formula: (c, i) => `${i < 12 ? "=250" : `=250+ROUND(${addr(c, 8)}/20,0)`}` },
  { label: "Total Revenue", formula: (c) => `=SUM(${addr(c, 10)}:${addr(c, 11)})` },
]);

writeFormulas(sheets["Operating Model"], 4, [
  { label: "Headcount", formula: (c, i) => `${i === 0 ? "=180" : `=${addr(c - 1, 4)}+IF(MOD(COLUMN(),4)=0,2,1)`}` },
  { label: "Avg Salary", formula: (c, i) => `${i === 0 ? "=10.5" : `=${addr(c - 1, 5)}*(1+${sAddr("Control", 2, 11)}/12)`}` },
  { label: "Payroll", formula: (c) => `=${addr(c, 4)}*${addr(c, 5)}` },
  { label: "COGS %", formula: () => `=${sAddr("Control", 2, 8)}` },
  { label: "COGS", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}*${addr(c, 7)}` },
  { label: "S&M %", formula: () => `=${sAddr("Control", 2, 9)}` },
  { label: "S&M", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}*${addr(c, 9)}` },
  { label: "R&D %", formula: () => `=${sAddr("Control", 2, 10)}` },
  { label: "R&D", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}*${addr(c, 11)}` },
  { label: "G&A %", formula: () => `=${sAddr("Control", 2, 11)}` },
  { label: "G&A", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}*${addr(c, 13)}` },
  { label: "Cloud Hosting", formula: (c, i) => `${i === 0 ? "=55" : `=${addr(c - 1, 15)}*(1+0.01)`}` },
  { label: "Maintenance Capex", formula: () => `=${sAddr("Control", 2, 23)}` },
  { label: "Depreciation", formula: (c, i) => `${i === 0 ? `=${addr(c, 16)}/${sAddr("Control", 2, 24)}` : `=${addr(c - 1, 17)}*0.94+${addr(c, 16)}/${sAddr("Control", 2, 24)}`}` },
  { label: "EBITDA", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}-${addr(c, 8)}-${addr(c, 10)}-${addr(c, 12)}-${addr(c, 14)}-${addr(c, 15)}` },
  { label: "EBIT", formula: (c) => `=${addr(c, 18)}-${addr(c, 17)}` },
  { label: "EBITDA Margin", formula: (c) => `=${addr(c, 18)}/'Revenue Drivers'!${addr(c, 12)}` },
]);

writeFormulas(sheets["Working Capital"], 4, [
  { label: "DSO", formula: () => `=${sAddr("Control", 2, 12)}` },
  { label: "Accounts Receivable", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}*${addr(c, 5)}/30` },
  { label: "Inventory Days", formula: () => `=${sAddr("Control", 2, 13)}` },
  { label: "Inventory", formula: (c) => `='Operating Model'!${addr(c, 8)}*${addr(c, 7)}/30` },
  { label: "AP Days", formula: () => `=${sAddr("Control", 2, 14)}` },
  { label: "Accounts Payable", formula: (c) => `='Operating Model'!${addr(c, 8)}*${addr(c, 9)}/30` },
  { label: "Deferred Revenue %", formula: () => `=${sAddr("Control", 2, 25)}` },
  { label: "Deferred Revenue", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}*${addr(c, 11)}` },
  { label: "Net Working Capital", formula: (c) => `=${addr(c, 5)}+${addr(c, 7)}-${addr(c, 9)}-${addr(c, 11)}` },
  { label: "Change in NWC", formula: (c, i) => `${i === 0 ? `=${addr(c, 12)}` : `=${addr(c, 12)}-${addr(c - 1, 12)}`}` },
]);

writeFormulas(sheets["Income Statement"], 4, [
  { label: "Revenue", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}` },
  { label: "COGS", formula: (c) => `='Operating Model'!${addr(c, 8)}` },
  { label: "Gross Profit", formula: (c) => `=${addr(c, 4)}-${addr(c, 5)}` },
  { label: "S&M", formula: (c) => `='Operating Model'!${addr(c, 10)}` },
  { label: "R&D", formula: (c) => `='Operating Model'!${addr(c, 12)}` },
  { label: "G&A", formula: (c) => `='Operating Model'!${addr(c, 14)}` },
  { label: "Cloud Hosting", formula: (c) => `='Operating Model'!${addr(c, 15)}` },
  { label: "EBITDA", formula: (c) => `='Operating Model'!${addr(c, 18)}` },
  { label: "Depreciation", formula: (c) => `='Operating Model'!${addr(c, 17)}` },
  { label: "EBIT", formula: (c) => `='Operating Model'!${addr(c, 19)}` },
  { label: "Interest Expense", formula: (c) => `='Debt Schedule'!${addr(c, 8)}` },
  { label: "EBT", formula: (c) => `=${addr(c, 13)}-${addr(c, 14)}` },
  { label: "Taxes", formula: (c) => `=MAX(0,${addr(c, 15)}*${sAddr("Control", 2, 15)})` },
  { label: "Net Income", formula: (c) => `=${addr(c, 15)}-${addr(c, 16)}` },
]);

writeFormulas(sheets["Cash Flow"], 4, [
  { label: "Net Income", formula: (c) => `='Income Statement'!${addr(c, 17)}` },
  { label: "Depreciation", formula: (c) => `='Income Statement'!${addr(c, 12)}` },
  { label: "Change in NWC", formula: (c) => `=-'Working Capital'!${addr(c, 13)}` },
  { label: "Capex", formula: (c) => `=-'Operating Model'!${addr(c, 16)}` },
  { label: "Operating Cash Flow", formula: (c) => `=SUM(${addr(c, 4)}:${addr(c, 7)})` },
  { label: "Debt Draw / (Repayment)", formula: (c) => `='Debt Schedule'!${addr(c, 7)}` },
  { label: "Share Repurchase", formula: (c) => `=-'Equity Bridge'!${addr(c, 8)}` },
  { label: "Ending Cash", formula: (c, i) => `${i === 0 ? `=${sAddr("Control", 2, 19)}+SUM(${addr(c, 8)}:${addr(c, 10)})` : `=${addr(c - 1, 11)}+SUM(${addr(c, 8)}:${addr(c, 10)})`}` },
]);

writeFormulas(sheets["Balance Sheet"], 4, [
  { label: "Cash", formula: (c) => `='Cash Flow'!${addr(c, 11)}` },
  { label: "Accounts Receivable", formula: (c) => `='Working Capital'!${addr(c, 5)}` },
  { label: "Inventory", formula: (c) => `='Working Capital'!${addr(c, 7)}` },
  { label: "PP&E", formula: (c, i) => `${i === 0 ? `=65+'Operating Model'!${addr(c, 16)}-'Operating Model'!${addr(c, 17)}` : `=${addr(c - 1, 7)}+'Operating Model'!${addr(c, 16)}-'Operating Model'!${addr(c, 17)}`}` },
  { label: "Other Assets", formula: () => `=${sAddr("Control", 2, 27)}` },
  { label: "Total Assets", formula: (c) => `=SUM(${addr(c, 4)}:${addr(c, 8)})` },
  { label: "Debt", formula: (c) => `='Debt Schedule'!${addr(c, 7)}` },
  { label: "Accounts Payable", formula: (c) => `='Working Capital'!${addr(c, 9)}` },
  { label: "Deferred Revenue", formula: (c) => `='Working Capital'!${addr(c, 11)}` },
  { label: "Shareholders Equity", formula: (c, i) => `${i === 0 ? `=150+'Income Statement'!${addr(c, 17)}` : `=${addr(c - 1, 13)}+'Income Statement'!${addr(c, 17)}-'Equity Bridge'!${addr(c, 8)}`}` },
  { label: "Total Liab & Equity", formula: (c) => `=SUM(${addr(c, 10)}:${addr(c, 13)})` },
  { label: "Balance Check", formula: (c) => `=${addr(c, 9)}+${addr(c, 14)}` },
]);

writeFormulas(sheets["Debt Schedule"], 4, [
  { label: "Opening Balance", formula: (c, i) => `${i === 0 ? `=${sAddr("Control", 2, 20)}` : `=${addr(c - 1, 7)}`}` },
  { label: "Interest Rate", formula: () => `=${sAddr("Control", 2, 16)}` },
  { label: "Interest Expense", formula: (c) => `=AVERAGE(${addr(c, 4)},${addr(c, 7)})*${addr(c, 5)}/12` },
  { label: "Mandatory Amortization", formula: (c) => `=MAX(0,${addr(c, 4)}*0.01)` },
  { label: "Cash Sweep", formula: (c) => `=MAX(0,'Cash Flow'!${addr(c, 9)}*0.2)` },
  { label: "Draw / (Repayment)", formula: (c) => `=-${addr(c, 6)}-${addr(c, 7)}` },
  { label: "Closing Balance", formula: (c) => `=${addr(c, 4)}-${addr(c, 6)}-${addr(c, 7)}` },
]);

writeFormulas(sheets["Equity Bridge"], 4, [
  { label: "Share Count", formula: (c, i) => `${i === 0 ? `=${sAddr("Control", 2, 21)}` : `=${addr(c - 1, 4)}-${addr(c, 8)}`}` },
  { label: "Repurchase Rate", formula: () => `=${sAddr("Control", 2, 26)}` },
  { label: "Shares Repurchased", formula: (c) => `=${addr(c, 4)}*${addr(c, 5)}` },
  { label: "EPS", formula: (c) => `='Income Statement'!${addr(c, 17)}/${addr(c, 4)}` },
]);

writeFormulas(sheets["Valuation"], 4, [
  { label: "EBITDA", formula: (c) => `='Income Statement'!${addr(c, 11)}` },
  { label: "Exit Multiple", formula: () => `=${sAddr("Control", 2, 18)}` },
  { label: "Enterprise Value", formula: (c) => `=${addr(c, 4)}*${addr(c, 5)}` },
  { label: "Net Debt", formula: (c) => `='Balance Sheet'!${addr(c, 10)}-'Balance Sheet'!${addr(c, 4)}` },
  { label: "Equity Value", formula: (c) => `=${addr(c, 6)}-${addr(c, 7)}` },
  { label: "Value Per Share", formula: (c) => `=${addr(c, 8)}/'Equity Bridge'!${addr(c, 4)}` },
]);

writeFormulas(sheets["Checks"], 4, [
  { label: "Revenue Link Check", formula: (c) => `='Revenue Drivers'!${addr(c, 12)}-'Income Statement'!${addr(c, 4)}` },
  { label: "Cash Check", formula: (c) => `='Cash Flow'!${addr(c, 11)}-'Balance Sheet'!${addr(c, 4)}` },
  { label: "Balance Check", formula: (c) => `='Balance Sheet'!${addr(c, 9)}` },
  { label: "EPS Check", formula: (c) => `='Valuation'!${addr(c, 9)}-'Income Statement'!${addr(c, 17)}/'Equity Bridge'!${addr(c, 4)}` },
]);

writeFormulas(sheets["Sensitivity"], 4, [
  { label: "Scenario Label", formula: () => `="Year 3"` },
  { label: "Upside Revenue", formula: () => `='Valuation'!M8/B4` },
  { label: "Price Case", formula: () => "=0.1" },
  { label: "EPS Case", formula: () => `='Valuation'!M9/B6` },
]);

// Seed issues
sheets["Revenue Drivers"].getRange("CT10").formulas = [["=CS10*(1+0.18)"]];
addIssue("hardcoded_jump", "Revenue Drivers!CT10", "One hidden revenue month uses 18% hardcoded jump", "hardcoded_jump");

sheets["Operating Model"].getRange("BN18").formulas = [["='Revenue Drivers'!BN12-'Operating Model'!BN8-'Operating Model'!BN10-'Operating Model'!BN12-'Operating Model'!BN14+'Operating Model'!BN15"]];
addIssue("formula_inconsistency", "Operating Model!BN18", "EBITDA adds cloud hosting instead of subtracting in one hidden month", "formula_inconsistency");

sheets["Working Capital"].getRange("DG13").formulas = [["=DG12"]];
addIssue("working_capital_change", "Working Capital!DG13", "Change in NWC line drops prior-period comparison in one hidden month", "formula_inconsistency");

sheets["Income Statement"].getRange("BR16").formulas = [["=BR15*0"]];
addIssue("tax_suppression", "Income Statement!BR16", "Taxes zeroed out in profitable month", "tax_anomaly");

sheets["Cash Flow"].getRange("CW9").formulas = [["=SUM(CW4:CW6)"]];
addIssue("cash_flow_omission", "Cash Flow!CW9", "Operating cash flow omits capex line in one hidden month", "formula_inconsistency");

sheets["Balance Sheet"].getRange("B15").formulas = [["=B9+B14"]];
addIssue("balance_check_sign", "Balance Sheet!B15", "Balance check adds instead of subtracting", "balance_check");

sheets["Debt Schedule"].getRange("CM7").formulas = [["=CM4+CM6-CM7"]];
addIssue("debt_rollforward", "Debt Schedule!CM7", "Closing balance uses circular/self reference style bad pattern", "formula_inconsistency");

sheets["Valuation"].getRange("CZ9").formulas = [["=CZ8*1000"]];
addIssue("per_share_mismatch", "Valuation!CZ9", "Value per share row replaced with enterprise-value-style hardcode", "label_semantics");

sheets["Sensitivity"].getRange("B5").formulas = [["='Valuation'!M8/B4"]];
addIssue("text_divisor", "Sensitivity!B5", "Formula divides by text label Year 3", "text_math");

sheets["Equity Bridge"].getRange("CH6").formulas = [["=CH4*1.5"]];
addIssue("hardcoded_share_change", "Equity Bridge!CH6", "Repurchased shares line uses 150% hardcoded multiplier", "hardcoded_jump");

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/${workbookName}`);
await fs.writeFile(`${outputDir}/${truthSetName}`, JSON.stringify(truthSet, null, 2));
console.log(JSON.stringify({ workbook: `${outputDir}/${workbookName}`, truthSet: `${outputDir}/${truthSetName}`, seededIssues: truthSet.length }, null, 2));
