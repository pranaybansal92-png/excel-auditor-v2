import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs";
const workbookName = "benchmark_project_finance_stress_test.xlsx";
const truthSetName = "benchmark_project_finance_truth_set.json";
const periods = 120;
const startCol = 2;

const workbook = Workbook.create();
const sheetNames = [
  "Control",
  "Construction",
  "Operations",
  "Revenue",
  "Opex",
  "Working Capital",
  "Debt",
  "Tax",
  "Cash Flow",
  "DSCR",
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

const control = sheets["Control"];
control.getRange("A4:B24").values = [
  ["Construction Months", 24],
  ["Commercial Operation Date Month", 25],
  ["Base Capacity MW", 200],
  ["Capacity Factor", 0.51],
  ["Tariff $/MWh", 87],
  ["Tariff Escalation", 0.02],
  ["Availability", 0.96],
  ["Curtailment", 0.03],
  ["Fixed Opex / Month", 0.7],
  ["Variable Opex / MWh", 9],
  ["DSRA Months", 6],
  ["Interest Rate", 0.075],
  ["Tax Rate", 0.25],
  ["Discount Rate", 0.09],
  ["Terminal Multiple", 10],
  ["Initial Debt", 420],
  ["Initial Equity", 180],
  ["NWC Days", 30],
  ["Maintenance Capex / Month", 0.2],
  ["Degradation %", 0.004],
  ["Inflation", 0.025],
];

writeFormulas(sheets["Construction"], 4, [
  { label: "EPC Capex", formula: (c, i) => `=${i < 24 ? (8 + (i % 3)) : 0}` },
  { label: "Development Cost", formula: (c, i) => `=${i < 18 ? 0.5 : 0}` },
  { label: "Owner Costs", formula: (c, i) => `=${i < 24 ? 0.25 : 0}` },
  { label: "IDC Drawn", formula: (c) => `='Debt'!${addr(c, 8)}` },
  { label: "Total Construction Spend", formula: (c) => `=SUM(${addr(c, 4)}:${addr(c, 7)})` },
]);

writeFormulas(sheets["Operations"], 4, [
  {
    label: "Available Capacity MW",
    formula: (c, i) =>
      `${i < 24 ? "=0" : i === 24 ? `=${sAddr("Control", 2, 6)}*${sAddr("Control", 2, 9)}` : `=${addr(c - 1, 4)}*(1-${sAddr("Control", 2, 23)}/12)`}`,
  },
  { label: "Hours", formula: () => "=730" },
  { label: "Gross MWh", formula: (c) => `=${addr(c, 4)}*${addr(c, 5)}*${sAddr("Control", 2, 7)}` },
  { label: "Curtailment %", formula: () => `=${sAddr("Control", 2, 11)}` },
  { label: "Net MWh", formula: (c) => `=${addr(c, 6)}*(1-${addr(c, 7)})` },
]);

writeFormulas(sheets["Revenue"], 4, [
  { label: "Tariff $/MWh", formula: (c, i) => `${i < 24 ? "=0" : i === 24 ? `=${sAddr("Control", 2, 8)}` : `=${addr(c - 1, 4)}*(1+${sAddr("Control", 2, 9)}/12)`}` },
  { label: "Net MWh", formula: (c) => `='Operations'!${addr(c, 8)}` },
  { label: "Energy Revenue", formula: (c) => `=${addr(c, 4)}*${addr(c, 5)}` },
  { label: "Other Revenue", formula: (c, i) => `${i < 24 ? "=0" : "=0.15"}` },
  { label: "Total Revenue", formula: (c) => `=SUM(${addr(c, 6)}:${addr(c, 7)})` },
]);

writeFormulas(sheets["Opex"], 4, [
  { label: "Fixed Opex", formula: (c, i) => `${i < 24 ? "=0" : i === 24 ? `=${sAddr("Control", 2, 12)}` : `=${addr(c - 1, 4)}*(1+${sAddr("Control", 2, 23)}/12)`}` },
  { label: "Variable Opex / MWh", formula: () => `=${sAddr("Control", 2, 13)}` },
  { label: "Variable Opex", formula: (c) => `='Operations'!${addr(c, 8)}*${addr(c, 5)}` },
  { label: "Maintenance Capex", formula: (c, i) => `${i < 24 ? "=0" : `=${sAddr("Control", 2, 22)}`}` },
  { label: "Total Opex", formula: (c) => `=SUM(${addr(c, 4)}:${addr(c, 7)})` },
]);

writeFormulas(sheets["Working Capital"], 4, [
  { label: "NWC Days", formula: () => `=${sAddr("Control", 2, 21)}` },
  { label: "Receivables", formula: (c) => `='Revenue'!${addr(c, 8)}*${addr(c, 4)}/30` },
  { label: "Payables", formula: (c) => `='Opex'!${addr(c, 8)}*20/30` },
  { label: "Net Working Capital", formula: (c) => `=${addr(c, 5)}-${addr(c, 6)}` },
  { label: "Change in NWC", formula: (c, i) => `${i === 0 ? `=${addr(c, 7)}` : `=${addr(c, 7)}-${addr(c - 1, 7)}`}` },
]);

writeFormulas(sheets["Debt"], 4, [
  { label: "Opening Balance", formula: (c, i) => `${i === 0 ? `=${sAddr("Control", 2, 19)}` : `=${addr(c - 1, 10)}`}` },
  { label: "Construction Draw", formula: (c, i) => `${i < 24 ? `='Construction'!${addr(c, 8)}*0.7` : "=0"}` },
  { label: "Scheduled Principal", formula: (c, i) => `${i < 24 ? "=0" : `=MAX(0,${addr(c, 4)}/96)`}` },
  { label: "Cash Sweep", formula: (c, i) => `${i < 24 ? "=0" : `=MAX(0,'Cash Flow'!${addr(c, 10)}*0.5)`}` },
  { label: "Closing Balance", formula: (c) => `=${addr(c, 4)}+${addr(c, 5)}-${addr(c, 6)}-${addr(c, 7)}` },
  { label: "Interest Rate", formula: () => `=${sAddr("Control", 2, 15)}` },
  { label: "Interest Expense", formula: (c) => `=AVERAGE(${addr(c, 4)},${addr(c, 8)})*${addr(c, 9)}/12` },
  { label: "DSRA Required", formula: (c) => `=('Debt'!${addr(c, 7)}+'Debt'!${addr(c, 10)})*${sAddr("Control", 2, 14)}` },
  { label: "DSRA Balance", formula: (c, i) => `${i === 0 ? "=0" : `=MIN(${addr(c, 11)},'Cash Flow'!${addr(c, 11)}+${addr(c - 1, 12)})`}` },
]);

writeFormulas(sheets["Tax"], 4, [
  { label: "EBITDA", formula: (c) => `='Revenue'!${addr(c, 8)}-'Opex'!${addr(c, 8)}` },
  { label: "Depreciation", formula: (c, i) => `${i < 24 ? "=0" : `='Construction'!${addr(c, 8)}/180`}` },
  { label: "EBIT", formula: (c) => `=${addr(c, 4)}-${addr(c, 5)}` },
  { label: "Interest", formula: (c) => `='Debt'!${addr(c, 10)}` },
  { label: "Taxable Income", formula: (c) => `=${addr(c, 6)}-${addr(c, 7)}` },
  { label: "Taxes", formula: (c) => `=MAX(0,${addr(c, 8)}*${sAddr("Control", 2, 16)})` },
]);

writeFormulas(sheets["Cash Flow"], 4, [
  { label: "Revenue", formula: (c) => `='Revenue'!${addr(c, 8)}` },
  { label: "Opex", formula: (c) => `=-'Opex'!${addr(c, 8)}` },
  { label: "Taxes", formula: (c) => `=-'Tax'!${addr(c, 9)}` },
  { label: "Change in NWC", formula: (c) => `=-'Working Capital'!${addr(c, 8)}` },
  { label: "Maintenance Capex", formula: (c) => `=-'Opex'!${addr(c, 7)}` },
  { label: "CFADS", formula: (c) => `=SUM(${addr(c, 4)}:${addr(c, 8)})` },
  { label: "Debt Service", formula: (c) => `='Debt'!${addr(c, 6)}+'Debt'!${addr(c, 10)}` },
  { label: "Cash Sweep", formula: (c) => `=-'Debt'!${addr(c, 7)}` },
  { label: "Equity Distribution", formula: (c) => `=MAX(0,${addr(c, 9)}-${addr(c, 10)}+'Debt'!${addr(c, 12)})` },
]);

writeFormulas(sheets["DSCR"], 4, [
  { label: "CFADS", formula: (c) => `='Cash Flow'!${addr(c, 9)}` },
  { label: "Debt Service", formula: (c) => `='Cash Flow'!${addr(c, 10)}` },
  { label: "DSCR", formula: (c) => `=${addr(c, 4)}/${addr(c, 5)}` },
  { label: "Min DSCR", formula: (c) => `=MIN(${addr(c, 6)},1.20)` },
]);

writeFormulas(sheets["Valuation"], 4, [
  { label: "CFADS", formula: (c) => `='Cash Flow'!${addr(c, 9)}` },
  { label: "Discount Rate", formula: () => `=${sAddr("Control", 2, 17)}` },
  { label: "Discount Factor", formula: (c, i) => `=1/(1+${addr(c, 5)}/12)^${i + 1}` },
  { label: "PV of CFADS", formula: (c) => `=${addr(c, 4)}*${addr(c, 6)}` },
  { label: "Exit Multiple", formula: () => `=${sAddr("Control", 2, 18)}` },
  { label: "Terminal Value", formula: (c) => `=('Revenue'!${addr(c, 8)}-'Opex'!${addr(c, 8)})*${addr(c, 8)}` },
  { label: "Enterprise Value", formula: (c) => `=${addr(c, 7)}+${addr(c, 9)}` },
]);

writeFormulas(sheets["Checks"], 4, [
  { label: "Revenue Link Check", formula: (c) => `='Revenue'!${addr(c, 8)}-'Cash Flow'!${addr(c, 4)}` },
  { label: "Debt Roll Check", formula: (c) => `='Debt'!${addr(c, 8)}-'Debt'!${addr(c, 4)}-'Debt'!${addr(c, 5)}+'Debt'!${addr(c, 6)}+'Debt'!${addr(c, 7)}` },
  { label: "DSCR Check", formula: (c) => `='DSCR'!${addr(c, 6)}-'Cash Flow'!${addr(c, 9)}/'Cash Flow'!${addr(c, 10)}` },
  { label: "Balance Check", formula: (c) => `='Cash Flow'!${addr(c, 9)}+'Cash Flow'!${addr(c, 10)}` },
]);

writeFormulas(sheets["Sensitivity"], 4, [
  { label: "Case Label", formula: () => `="Base"` },
  { label: "Price Upside", formula: () => "=0.05" },
  { label: "DSCR Upside", formula: () => `='DSCR'!M6/(1-B5)` },
  { label: "Text Divisor", formula: () => `="Year 2"` },
  { label: "Bad Ratio", formula: () => `='Valuation'!M10/B8` },
]);

// Seed issues
sheets["Debt"].getRange("AZ8").formulas = [["=AZ4+AZ5+AZ6-AZ7"]];
addIssue("formula_inconsistency", "Debt!AZ8", "Closing balance adds scheduled principal instead of subtracting it", "formula_inconsistency");

sheets["Cash Flow"].getRange("AR9").formulas = [["=SUM(AR4:AR7)"]];
addIssue("working_capital_omission", "Cash Flow!AR9", "CFADS omits maintenance capex in one hidden period", "formula_inconsistency");

sheets["DSCR"].getRange("BH6").formulas = [["=BH4+BH5"]];
addIssue("ratio_mismatch", "DSCR!BH6", "DSCR adds instead of divides in one hidden month", "label_semantics");

sheets["Valuation"].getRange("CW9").formulas = [["=('Revenue'!CW8-'Opex'!CW8)*(CW8+1)"]];
addIssue("terminal_value_bad_multiple", "Valuation!CW9", "Terminal value uses month index instead of exit multiple", "formula_inconsistency");

sheets["Checks"].getRange("B7").formulas = [["='Cash Flow'!B9+'Cash Flow'!B10"]];
addIssue("balance_check_sign", "Checks!B7", "Balance check adds components instead of offsetting", "balance_check");

sheets["Sensitivity"].getRange("B8").formulas = [["='Valuation'!M10/B7"]];
addIssue("text_math", "Sensitivity!B8", "Formula divides by text label Year 2", "text_math");

sheets["Revenue"].getRange("CF4").formulas = [["=CE4*(1+0.2)"]];
addIssue("hardcoded_escalation", "Revenue!CF4", "One tariff month uses 20% hardcoded escalation", "hardcoded_numbers");

sheets["Opex"].getRange("BP4").formulas = [["=BO4*(1+0.5)"]];
addIssue("hardcoded_opex_jump", "Opex!BP4", "One opex month uses 50% hardcoded jump", "hardcoded_numbers");

sheets["Tax"].getRange("BY9").formulas = [["=BY8*0"]];
addIssue("tax_logic", "Tax!BY9", "Taxes zeroed out in one profitable month", "hardcoded_numbers");

sheets["Debt"].getRange("N11").formulas = [["=('Debt'!N10+'Debt'!N7)*'Control'!B14"]];
addIssue("dsra_months_mismatch", "Debt!N11", "DSRA required uses months count directly instead of dividing by 12", "hardcoded_numbers");

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/${workbookName}`);
await fs.writeFile(`${outputDir}/${truthSetName}`, JSON.stringify(truthSet, null, 2));
console.log(JSON.stringify({ workbook: `${outputDir}/${workbookName}`, truthSet: `${outputDir}/${truthSetName}`, seededIssues: truthSet.length }, null, 2));
