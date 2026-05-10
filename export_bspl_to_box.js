/**
 * 各社のBS/PLをExcel(.xlsx)で出力し、Boxの月次フォルダに保存
 *
 * 出力先: C:\Users\t-mat\Box\001_BEYOND\010_Meeting\FY2026\020_ BEYOND Board\BSPL
 * ファイル名: YYYY.MM.DD 事業会社名 PL.xlsx / BS.xlsx
 *
 * Usage:
 *   node export_bspl_to_box.js               # 最新締め月分を出力
 *   node export_bspl_to_box.js 2026-03-31    # 指定日(月末)分を出力
 */
const https = require('https');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const RATE_LIMIT_DELAY = 600;
const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const BOX_DIR = 'C:/Users/t-mat/Box/001_BEYOND/010_Meeting/FY2026/020_ BEYOND Board/BSPL';

// Determine target close date
const arg = process.argv[2];
let closeYear, closeMonth, closeDay;
if (arg) {
  const [y, m, d] = arg.split('-').map(Number);
  closeYear = y; closeMonth = m; closeDay = d;
} else {
  // Default: latest closed month (today < 20: 2 months back, else 1 back)
  const now = new Date();
  const monthsBack = now.getDate() < 20 ? 2 : 1;
  let cy = now.getFullYear();
  let cm = now.getMonth() + 1 - monthsBack;
  if (cm <= 0) { cm += 12; cy--; }
  closeYear = cy; closeMonth = cm;
  closeDay = new Date(cy, cm, 0).getDate();
}
const closeDateStr = `${closeYear}-${String(closeMonth).padStart(2,'0')}-${String(closeDay).padStart(2,'0')}`;
const fileDateStr = `${closeYear}.${String(closeMonth).padStart(2,'0')}.${String(closeDay).padStart(2,'0')}`;

// Fiscal year start (May start, with 184 override)
const FY_START_DATE_OVERRIDES = {
  12243427: '2025-11-13'  // 184: founded 2025-11-13
};
function getFYStart(companyId) {
  const ov = FY_START_DATE_OVERRIDES[companyId];
  if (ov) return ov;
  // Determine FY: if closeMonth >= 5, FY = closeYear; else FY = closeYear - 1
  const fyYear = closeMonth >= 5 ? closeYear : closeYear - 1;
  return `${fyYear}-05-01`;
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function httpsRequest(urlStr, options, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlStr);
    const opts = { hostname: url.hostname, path: url.pathname + url.search, method: options.method || 'GET', headers: options.headers || {} };
    if (body) opts.headers['Content-Length'] = Buffer.byteLength(body);
    const req = https.request(opts, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => { try { resolve({ statusCode: res.statusCode, data: JSON.parse(data) }); } catch (e) { resolve({ statusCode: res.statusCode, data: { raw: data } }); } });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}
async function refresh(rt) {
  const body = new URLSearchParams({ grant_type: 'refresh_token', client_id: CLIENT_ID, client_secret: CLIENT_SECRET, refresh_token: rt }).toString();
  const r = await httpsRequest('https://accounts.secure.freee.co.jp/public_api/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }, body);
  if (r.data.error) throw new Error(r.data.error);
  return r.data;
}
async function apiGet(token, ep, params) {
  const url = new URL('https://api.freee.co.jp/api/1' + ep);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, String(v)));
  await sleep(RATE_LIMIT_DELAY);
  const r = await httpsRequest(url.toString(), { method: 'GET', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' } });
  if (r.statusCode === 429) { console.log('    Rate limited, waiting 60s'); await sleep(60000); return apiGet(token, ep, params); }
  if (r.statusCode === 401) throw new Error('Unauthorized');
  if (r.statusCode >= 400) throw new Error('HTTP ' + r.statusCode + ': ' + JSON.stringify(r.data).substring(0, 200));
  return r.data;
}

// Apply Meiryo UI font + alignment to all cells
function applyStyling(ws, range) {
  const dec = XLSX.utils.decode_range(range || ws['!ref']);
  for (let R = dec.s.r; R <= dec.e.r; R++) {
    for (let C = dec.s.c; C <= dec.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell) {
        cell.s = cell.s || {};
        cell.s.font = { name: 'Meiryo UI', sz: 10, ...(cell.s.font || {}) };
      }
    }
  }
}

// Write PL Excel
async function writePLFile(token, companyId, companyName) {
  const fyStart = getFYStart(companyId);
  // Fetch month-by-month cumulative, then convert to monthly
  const fyStartDate = new Date(fyStart);
  const fyStartYear = fyStartDate.getFullYear();
  const fyStartMonth = fyStartDate.getMonth() + 1;
  const months = [];
  for (let i = 0; i < 12; i++) {
    let y = fyStartYear;
    let m = fyStartMonth + i;
    if (m > 12) { m -= 12; y++; }
    // Skip months after close date
    if (y > closeYear || (y === closeYear && m > closeMonth)) break;
    months.push({ year: y, month: m });
  }

  // Cumulative fetch per month
  const cumByMonth = {};
  for (const { year, month } of months) {
    const lastDay = new Date(year, month, 0).getDate();
    const end = `${year}-${String(month).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
    const data = await apiGet(token, '/reports/trial_pl', {
      company_id: companyId,
      start_date: fyStart,
      end_date: end
    });
    const balances = data.trial_pl?.balances || [];
    const monthKey = `${year}/${String(month).padStart(2,'0')}`;
    cumByMonth[monthKey] = balances;
  }

  // Build account item map (account_item_name -> {month: cumValue})
  const items = {}; // name -> { category, hierarchy_level, cumValues: { '2025/05': 100, ... } }
  for (const [monthKey, balances] of Object.entries(cumByMonth)) {
    for (const b of balances) {
      const name = b.account_item_name || ('[' + b.account_category_name + ']');
      if (!items[name]) items[name] = { category: b.account_category_name, level: b.hierarchy_level, cum: {} };
      items[name].cum[monthKey] = b.closing_balance || 0;
    }
  }

  // Compute monthly values from cumulative
  const monthKeys = months.map(m => `${m.year}/${String(m.month).padStart(2,'0')}`);
  const itemRows = []; // [{name, category, level, monthly: {key: val}, total: val}]
  for (const [name, data] of Object.entries(items)) {
    const monthly = {};
    let prevCum = 0;
    let total = 0;
    for (const mk of monthKeys) {
      const cum = data.cum[mk] !== undefined ? data.cum[mk] : prevCum;
      const val = cum - prevCum;
      monthly[mk] = val;
      total = cum;
      prevCum = cum;
    }
    itemRows.push({ name, category: data.category, level: data.level, monthly, total });
  }
  // Sort: by category order in trial_pl, then hierarchy
  itemRows.sort((a, b) => (a.level || 99) - (b.level || 99));

  // Build Excel sheet: rows = account items, columns = months + 累計
  const headers = ['勘定科目', 'カテゴリ', ...monthKeys, '累計'];
  const data = [headers];
  for (const r of itemRows) {
    const row = [r.name, r.category];
    for (const mk of monthKeys) row.push(r.monthly[mk] || 0);
    row.push(r.total);
    data.push(row);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  // Column widths
  ws['!cols'] = [
    { wch: 30 }, { wch: 16 }, ...monthKeys.map(() => ({ wch: 13 })), { wch: 14 }
  ];
  applyStyling(ws);
  XLSX.utils.book_append_sheet(wb, ws, 'PL');

  const outPath = path.join(BOX_DIR, `${fileDateStr} ${companyName} PL.xlsx`);
  XLSX.writeFile(wb, outPath);
  return outPath;
}

// Write BS Excel
async function writeBSFile(token, companyId, companyName) {
  const fyStart = getFYStart(companyId);
  const data = await apiGet(token, '/reports/trial_bs', {
    company_id: companyId,
    start_date: fyStart,
    end_date: closeDateStr
  });
  const balances = data.trial_bs?.balances || [];

  const wb = XLSX.utils.book_new();
  const headers = ['勘定科目', 'カテゴリ', '階層', '期首残高', '借方', '貸方', '期末残高', '構成比%'];
  const rows = [headers];
  for (const b of balances) {
    rows.push([
      b.account_item_name || '',
      b.account_category_name || '',
      b.hierarchy_level || '',
      b.opening_balance || 0,
      b.debit_amount || 0,
      b.credit_amount || 0,
      b.closing_balance || 0,
      b.composition_ratio !== undefined ? b.composition_ratio : ''
    ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 28 }, { wch: 18 }, { wch: 6 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 10 }];
  applyStyling(ws);
  XLSX.utils.book_append_sheet(wb, ws, 'BS');

  const outPath = path.join(BOX_DIR, `${fileDateStr} ${companyName} BS.xlsx`);
  XLSX.writeFile(wb, outPath);
  return outPath;
}

(async () => {
  console.log(`=== BS/PL Export to Box ===`);
  console.log(`Close date: ${closeDateStr}`);
  console.log(`Output dir: ${BOX_DIR}`);
  if (!fs.existsSync(BOX_DIR)) {
    console.log('Creating BSPL directory...');
    fs.mkdirSync(BOX_DIR, { recursive: true });
  }
  console.log();

  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  const updated = [];
  for (const t of tokens) {
    console.log(`[${t.name}]`);
    let access = t.access_token;
    try {
      const r = await refresh(t.refresh_token);
      access = r.access_token;
      updated.push({ ...t, access_token: r.access_token, refresh_token: r.refresh_token });
    } catch (e) { updated.push(t); console.log('  refresh fail:', e.message); continue; }

    // Skip 184 if close date is before its founding
    if (t.company_id === 12243427 && closeDateStr < '2025-11-13') {
      console.log('  Skip: 184 not founded yet'); continue;
    }

    try {
      const plPath = await writePLFile(access, t.company_id, t.name);
      console.log('  ✓ PL:', path.basename(plPath));
    } catch (e) { console.log('  PL ERROR:', e.message); }
    try {
      const bsPath = await writeBSFile(access, t.company_id, t.name);
      console.log('  ✓ BS:', path.basename(bsPath));
    } catch (e) { console.log('  BS ERROR:', e.message); }
  }
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updated, null, 2));
  console.log('\n=== Complete ===');
})();
