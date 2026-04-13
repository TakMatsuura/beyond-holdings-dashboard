/**
 * BEYOND Holdings - Budget Excel Parser
 * Reads per-company budget Excel files from Box and outputs budget JSON.
 * Each company file has a different format, so extraction is configured per-company.
 *
 * Output format per company:
 *   { companyId: number, budgets: [{ month: "YYYY/MM", revenueBudget: number, opBudget: number }] }
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const SCRIPT_DIR = __dirname;
const OUTPUT_FILE = path.join(SCRIPT_DIR, 'budget_data.json');

// Box base path for budget files
const BOX_BASE = 'C:/Users/t-mat/Box/001_BEYOND/001_予算/FY2025';

// FY2025 months: May 2025 - April 2026
const FY_MONTHS = [
  '2025/05','2025/06','2025/07','2025/08','2025/09','2025/10',
  '2025/11','2025/12','2026/01','2026/02','2026/03','2026/04'
];

// ── Per-company extraction configs ──

const COMPANIES = [
  {
    name: 'デンコー',
    companyId: 3115888,
    file: '01. DENKO/DENKO_予算 202505 - 202604_加藤修正.xlsx',
    extract: extractGroupA,
    // Group A: YYYYMM codes in header row 0, label col 0, monthly cols 2-13
    sheetName: '予算 202405 - 202504',
    headerRow: 0,
    labelCol: 0,
    revenueLabel: '売上 // 売上高',
    opLabel: '営業利益',
    monthStartCol: 2,
  },
  {
    name: 'SAFARI',
    companyId: 1980825,
    file: '02. SAFARI/SAFARI_2025年度月次予算_ver0.1.xlsx',
    extract: extractSafari,
    sheetName: '26年4月期予算',
  },
  {
    name: 'HANABI',
    companyId: 2619462,
    file: '03. HANABI/HANABI 予算draft 202505-202604_ver0.2.xlsx',
    extract: extractGroupA,
    sheetName: '予算 202505 - 502604',
    headerRow: 3,
    labelCol: 1,
    revenueLabel: '売上 // 売上高',
    opLabel: '営業利益',
    monthStartCol: 3,
  },
  {
    name: 'K2',
    companyId: 2619418,
    file: '04. K2/FY25 K2 予算 202505 - 202604 ver0.2.xlsx',
    extract: extractGroupA,
    sheetName: '予算 202405 - 202504',
    headerRow: 0,
    labelCol: 0,
    revenueLabel: '売上高合計',
    opLabel: '営業利益',
    monthStartCol: 2,
  },
  {
    name: 'ライフプロ',
    companyId: 11006999,
    file: '05. ライフプロ/ライフプロ_2025年度月次予算ver0.4.xlsx',
    extract: extractLifePro,
    sheetName: '26年4月期予算',
  },
  {
    name: 'M7Logi',
    companyId: 10713894,
    file: '06. M7Logi/M7logi予算202505-202604v5.xlsx',
    extract: extractM7Logi,
    sheetName: 'V5(20250709)',
  },
  {
    name: 'BEYOND Holdings',
    companyId: 10815529,
    file: '00. BEYOND/FY25 BEYOND 予算 202505 - 202604_ver0.2.xlsx',
    extract: extractGroupA,
    sheetName: 'Final',
    headerRow: 0,
    labelCol: 0,
    revenueLabel: '売上',
    opLabel: '営業利益',
    monthStartCol: 2,
  },
];

// ── Helper: read sheet as array of arrays ──
function readSheet(filePath, sheetName) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    console.log(`  Available sheets: ${wb.SheetNames.join(', ')}`);
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}

// ── Helper: find row index by label ──
function findRow(data, labelCol, label) {
  for (let i = 0; i < data.length; i++) {
    const cell = data[i]?.[labelCol];
    if (cell != null && String(cell).trim() === label) return i;
  }
  // Try partial match
  for (let i = 0; i < data.length; i++) {
    const cell = data[i]?.[labelCol];
    if (cell != null && String(cell).trim().startsWith(label)) return i;
  }
  return -1;
}

// ── Group A: YYYYMM codes in header row ──
// Used by: DENKO, HANABI, K2, BEYOND
function extractGroupA(config) {
  const filePath = path.join(BOX_BASE, config.file);
  const data = readSheet(filePath, config.sheetName);
  const { headerRow, labelCol, monthStartCol } = config;

  // Verify month headers (YYYYMM codes)
  const headers = data[headerRow];
  const monthCols = [];
  for (let c = monthStartCol; c < monthStartCol + 12; c++) {
    const h = headers?.[c];
    monthCols.push(c);
    // Verify it looks like a YYYYMM code
    if (h && h >= 202500 && h <= 202700) continue;
    console.log(`  Warning: col ${c} header = ${h} (expected YYYYMM code)`);
  }

  const revRow = findRow(data, labelCol, config.revenueLabel);
  const opRow = findRow(data, labelCol, config.opLabel);
  console.log(`  Revenue row: ${revRow} (${data[revRow]?.[labelCol]})`);
  console.log(`  OP row: ${opRow} (${data[opRow]?.[labelCol]})`);

  if (revRow === -1 || opRow === -1) throw new Error('Could not find revenue/OP rows');

  const budgets = [];
  for (let i = 0; i < 12; i++) {
    const col = monthCols[i];
    const rev = Number(data[revRow]?.[col]) || 0;
    const op = Number(data[opRow]?.[col]) || 0;
    budgets.push({ month: FY_MONTHS[i], revenueBudget: Math.round(rev), opBudget: Math.round(op) });
  }
  return budgets;
}

// ── SAFARI: Date serial headers, flat labels ──
function extractSafari(config) {
  const filePath = path.join(BOX_BASE, config.file);
  const data = readSheet(filePath, config.sheetName);

  // Row 7: 売上高計, Row 35: 営業利益, cols 1-12 for months
  const revRow = findRow(data, 0, '売上高計');
  const opRow = findRow(data, 0, '営業利益');
  console.log(`  Revenue row: ${revRow} (${data[revRow]?.[0]})`);
  console.log(`  OP row: ${opRow} (${data[opRow]?.[0]})`);

  if (revRow === -1 || opRow === -1) throw new Error('Could not find revenue/OP rows');

  const budgets = [];
  for (let i = 0; i < 12; i++) {
    const col = i + 1; // cols 1-12
    const rev = Number(data[revRow]?.[col]) || 0;
    const op = Number(data[opRow]?.[col]) || 0;
    budgets.push({ month: FY_MONTHS[i], revenueBudget: Math.round(rev), opBudget: Math.round(op) });
  }
  return budgets;
}

// ── ライフプロ: Irregular columns with 実績/差分 interleaved ──
function extractLifePro(config) {
  const filePath = path.join(BOX_BASE, config.file);
  const data = readSheet(filePath, config.sheetName);

  // Cols: 1=May予算, 2=May実績, 3=May差分, 4=Jun予算, 5=Jun実績, 6=Jun差分,
  //       7=Jul, 8=Aug, 9=Sep, 10=Oct, 11=Nov, 12=Dec, 13=Jan, 14=Feb, 15=Mar, 16=Apr
  // Budget columns for 12 months:
  const budgetCols = [1, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];

  const revRow = findRow(data, 0, '売上高計');
  const opRow = findRow(data, 0, '営業利益');
  console.log(`  Revenue row: ${revRow} (${data[revRow]?.[0]})`);
  console.log(`  OP row: ${opRow} (${data[opRow]?.[0]})`);

  if (revRow === -1 || opRow === -1) throw new Error('Could not find revenue/OP rows');

  const budgets = [];
  for (let i = 0; i < 12; i++) {
    const col = budgetCols[i];
    const rev = Number(data[revRow]?.[col]) || 0;
    const op = Number(data[opRow]?.[col]) || 0;
    budgets.push({ month: FY_MONTHS[i], revenueBudget: Math.round(rev), opBudget: Math.round(op) });
  }
  return budgets;
}

// ── M7Logi: Wide format, Japanese date headers, FY25 budget in cols 34-45 ──
function extractM7Logi(config) {
  const filePath = path.join(BOX_BASE, config.file);
  const data = readSheet(filePath, config.sheetName);

  // Row 3: 売上, Row 41: 営業利益, FY25 budget cols 34-45 (25年5月~26年4月)
  const revRow = findRow(data, 0, '売上');
  const opRow = findRow(data, 0, '営業利益');
  console.log(`  Revenue row: ${revRow} (${data[revRow]?.[0]})`);
  console.log(`  OP row: ${opRow} (${data[opRow]?.[0]})`);

  if (revRow === -1 || opRow === -1) throw new Error('Could not find revenue/OP rows');

  const budgets = [];
  for (let i = 0; i < 12; i++) {
    const col = 34 + i;
    const rev = Number(data[revRow]?.[col]) || 0;
    const op = Number(data[opRow]?.[col]) || 0;
    budgets.push({ month: FY_MONTHS[i], revenueBudget: Math.round(rev), opBudget: Math.round(op) });
  }
  return budgets;
}

// ── Main ──
function main() {
  console.log('=== BEYOND Holdings Budget Parser ===\n');

  const results = [];

  for (const config of COMPANIES) {
    console.log(`\n--- ${config.name} (${config.companyId}) ---`);
    const filePath = path.join(BOX_BASE, config.file);

    if (!fs.existsSync(filePath)) {
      console.log(`  File not found: ${filePath}`);
      continue;
    }

    try {
      const budgets = config.extract(config);
      const totalRev = budgets.reduce((s, b) => s + b.revenueBudget, 0);
      const totalOP = budgets.reduce((s, b) => s + b.opBudget, 0);
      console.log(`  Annual budget: Rev=${(totalRev / 10000).toFixed(0)}万, OP=${(totalOP / 10000).toFixed(0)}万`);
      results.push({ companyId: config.companyId, name: config.name, budgets });
    } catch (e) {
      console.log(`  ERROR: ${e.message}`);
    }
  }

  // Save
  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(results, null, 2));
  console.log(`\n=== Complete: ${results.length} companies → ${OUTPUT_FILE} ===`);
}

main();
