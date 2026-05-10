/**
 * Fetch Freee data for a specific fiscal year (May start, April end)
 * Usage: node fetch_year.js 2024
 *
 * Saves to: freee_data_FY{YEAR}.json
 */
const https = require('https');
const fs = require('fs');
const path = require('path');

const FY = parseInt(process.argv[2]);
if (!FY || FY < 2018 || FY > 2030) {
  console.error('Usage: node fetch_year.js <YEAR>');
  console.error('Example: node fetch_year.js 2024  (= FY2024: 2024/05 - 2025/04)');
  process.exit(1);
}

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const RATE_LIMIT_DELAY = 600;
const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const OUTPUT_FILE = path.join(SCRIPT_DIR, `freee_data_FY${FY}.json`);

const FY_START_MONTH = 5;
const FY_START_DATE_DEFAULT = `${FY}-05-01`;
const FY_END_DATE = `${FY + 1}-04-30`;

// 184: founded 2025-11-13. For FY2024 and earlier, skip.
const FY_START_DATE_OVERRIDES = {
  12243427: { month: 11, startDate: '2025-11-13', firstFY: 2025 }
};
const COMPANY_FIRST_FY = {
  // Companies that didn't exist in older years (approximate; will fail gracefully)
  12243427: 2025,  // 184
  10713669: 2024,  // DNK (2023 setup)
  10713894: 2024,  // M7Logi
  10815529: 2024,  // BEYOND HD
  11006999: 2024,  // ライフプロ
};

// Determine cutoff: only include months where closing is complete
// FY2025 is current; for past FYs, the entire year is closed
const now = new Date();
const todayYear = now.getFullYear();
const todayMonth = now.getMonth() + 1;
const todayDay = now.getDate();
const monthsBack = todayDay < 20 ? 2 : 1;
let cutoffYear = todayYear;
let cutoffMonth = todayMonth - monthsBack;
if (cutoffMonth <= 0) { cutoffMonth += 12; cutoffYear--; }

console.log(`=== Fetching FY${FY} (${FY_START_DATE_DEFAULT} ~ ${FY_END_DATE}) ===`);
console.log(`Today: ${todayYear}/${todayMonth}/${todayDay}, cutoff: ${cutoffYear}/${String(cutoffMonth).padStart(2,'0')}`);

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function httpsRequest(urlStr, options, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlStr);
    const opts = { hostname: url.hostname, path: url.pathname + url.search, method: options.method || 'GET', headers: options.headers || {} };
    if (body) opts.headers['Content-Length'] = Buffer.byteLength(body);
    const req = https.request(opts, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve({ statusCode: res.statusCode, data: JSON.parse(data) }); }
        catch (e) { resolve({ statusCode: res.statusCode, data: { raw: data.substring(0, 300) } }); }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}
async function refreshToken(rt) {
  const body = new URLSearchParams({ grant_type: 'refresh_token', client_id: CLIENT_ID, client_secret: CLIENT_SECRET, refresh_token: rt }).toString();
  const r = await httpsRequest('https://accounts.secure.freee.co.jp/public_api/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }, body);
  if (r.data.error) throw new Error('refresh: ' + r.data.error);
  return r.data;
}
async function apiGet(token, ep, params, opts={}) {
  const url = new URL('https://api.freee.co.jp/api/1' + ep);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, String(v)));
  await sleep(RATE_LIMIT_DELAY);
  const r = await httpsRequest(url.toString(), { method: 'GET', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' } });
  if (r.statusCode === 429) { console.log('    Rate limited 60s'); await sleep(60000); return apiGet(token, ep, params, opts); }
  if (r.statusCode === 401) throw new Error('Unauthorized');
  if (r.statusCode >= 400) {
    if (opts.allowError) return null;
    throw new Error('HTTP ' + r.statusCode + ': ' + JSON.stringify(r.data).substring(0, 200));
  }
  return r.data;
}

function isMonthClosed(year, month) {
  if (year < cutoffYear) return true;
  if (year === cutoffYear && month <= cutoffMonth) return true;
  return false;
}

async function fetchMonthlyPL(token, companyId) {
  const override = FY_START_DATE_OVERRIDES[companyId];
  const fyStartMonth = override ? override.month : FY_START_MONTH;
  const fyStartDate = override && override.startDate && FY === override.firstFY ? override.startDate : FY_START_DATE_DEFAULT;

  const months = [];
  for (let i = 0; i < 12; i++) {
    let y = FY;
    let m = fyStartMonth + i;
    if (m > 12) { m -= 12; y++; }
    months.push({ year: y, month: m });
  }

  const cumulative = [];
  for (const { year, month } of months) {
    if (!isMonthClosed(year, month)) { cumulative.push(null); continue; }
    const lastDay = new Date(year, month, 0).getDate();
    const end = `${year}-${String(month).padStart(2, '0')}-${lastDay}`;
    try {
      const data = await apiGet(token, '/reports/trial_pl', {
        company_id: companyId,
        start_date: fyStartDate,
        end_date: end
      }, { allowError: true });
      if (!data) { cumulative.push(null); continue; }
      const balances = data.trial_pl?.balances || [];
      const findH1 = (name) => {
        const item = balances.find(b => b.account_category_name === name && b.hierarchy_level === 1);
        return item ? (item.closing_balance || 0) : 0;
      };
      const revenue = Math.abs(findH1('売上高'));
      const grossProfit = findH1('売上総損益金額');
      const cogs = revenue - grossProfit;
      const operatingProfit = findH1('営業損益金額');
      const sga = grossProfit - operatingProfit;
      const dep = balances
        .filter(b => b.account_item_name && b.account_item_name.includes('減価償却'))
        .reduce((sum, b) => sum + Math.abs(b.closing_balance || 0), 0);
      cumulative.push({ year, month, revenue, cogs, grossProfit, sga, operatingProfit, depreciation: dep });
    } catch (e) {
      cumulative.push(null);
    }
  }

  // cumulative -> monthly
  const monthly = [];
  for (let i = 0; i < cumulative.length; i++) {
    const cur = cumulative[i];
    if (!cur) continue;
    const monthStr = `${cur.year}/${String(cur.month).padStart(2, '0')}`;
    if (i === 0 || !cumulative[i - 1]) {
      monthly.push({ month: monthStr, revenue: cur.revenue, cogs: cur.cogs, grossProfit: cur.grossProfit, sga: cur.sga, operatingProfit: cur.operatingProfit, depreciation: cur.depreciation, ebitda: cur.operatingProfit + cur.depreciation });
    } else {
      const prev = cumulative[i - 1];
      const op = cur.operatingProfit - prev.operatingProfit;
      const dep = cur.depreciation - prev.depreciation;
      monthly.push({
        month: monthStr,
        revenue: cur.revenue - prev.revenue,
        cogs: cur.cogs - prev.cogs,
        grossProfit: cur.grossProfit - prev.grossProfit,
        sga: cur.sga - prev.sga,
        operatingProfit: op,
        depreciation: dep,
        ebitda: op + dep
      });
    }
  }
  return monthly;
}

async function fetchBSWithPartners(token, companyId) {
  const override = FY_START_DATE_OVERRIDES[companyId];
  const startDate = override && override.startDate && FY === override.firstFY ? override.startDate : FY_START_DATE_DEFAULT;

  // Try fiscal year-end BS; if year not yet ended, use latest closed month
  let endDate = FY_END_DATE;
  if (FY === todayYear || (FY + 1 === todayYear && todayMonth <= 4)) {
    // Current FY: use latest closed month-end
    const lastDay = new Date(cutoffYear, cutoffMonth, 0).getDate();
    endDate = `${cutoffYear}-${String(cutoffMonth).padStart(2,'0')}-${lastDay}`;
  }

  const data = await apiGet(token, '/reports/trial_bs', {
    company_id: companyId,
    start_date: startDate,
    end_date: endDate,
    breakdown_display_type: 'partner'
  }, { allowError: true });
  if (!data || !data.trial_bs) return null;
  const balances = data.trial_bs.balances || [];

  const find = (name, level) => {
    const item = balances.find(b => b.account_category_name === name && (level == null || b.hierarchy_level === level));
    return item ? (item.closing_balance || 0) : 0;
  };
  const cash = balances.filter(b => b.account_category_name === '現金・預金').reduce((s, b) => s + (b.closing_balance || 0), 0);

  const totalAssets = find('資産', 1);
  const totalLiabilities = find('負債', 1);
  const netAssets = find('純資産', 1);

  const LOAN_LIAB = ['短期借入金','長期借入金','役員借入金'];
  const LOAN_ASSET = ['短期貸付金','長期貸付金','立替金'];
  const isDenko = (p) => /デンコー|DENKO|電工/i.test(p.name || '');
  let denkoLiab = 0, denkoAssetReverse = 0;
  let extShort = 0, extLong = 0, extOff = 0;
  const denkoDetail = [];

  for (const b of balances) {
    const name = b.account_item_name || '';
    const isLoanLiab = LOAN_LIAB.some(k => name.includes(k));
    const isLoanAsset = LOAN_ASSET.some(k => name.includes(k));
    if (!isLoanLiab && !isLoanAsset) continue;
    const partners = b.partners || [];
    let denkoBal = 0, extBal = 0;
    for (const p of partners) {
      if (isDenko(p)) denkoBal += (p.closing_balance || 0);
      else if (isLoanLiab) extBal += (p.closing_balance || 0);
    }
    if (denkoBal !== 0) {
      denkoDetail.push({ account: name, side: isLoanLiab ? 'liability' : 'asset', amount: denkoBal });
      if (isLoanLiab) denkoLiab += denkoBal; else denkoAssetReverse += denkoBal;
    }
    if (isLoanLiab) {
      if (name.includes('短期借入金')) extShort += extBal;
      else if (name.includes('長期借入金')) extLong += extBal;
      else if (name.includes('役員借入金')) extOff += extBal;
    }
  }
  return {
    bs: { totalAssets, cash, totalLiabilities, netAssets },
    bsExtended: {
      denko: { liability: denkoLiab, assetReverse: denkoAssetReverse, net: denkoLiab - denkoAssetReverse, detail: denkoDetail },
      external: { shortTerm: extShort, longTerm: extLong, officer: extOff, total: extShort + extLong + extOff }
    }
  };
}

async function fetchPartners(token, companyId) {
  // Top customers/suppliers based on transactions in this FY
  // For now skip for historical years to save API calls; can be added later
  return { topCustomers: [], topSuppliers: [] };
}

(async () => {
  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  const updatedTokens = [];
  const companies = [];

  for (const t of tokens) {
    console.log(`\n[${t.name}] (id=${t.company_id})`);
    // Skip if company didn't exist in this FY
    const firstFY = COMPANY_FIRST_FY[t.company_id];
    if (firstFY && FY < firstFY) {
      console.log(`  SKIP: ${t.name} 設立前 (firstFY=${firstFY})`);
      updatedTokens.push(t);
      continue;
    }
    let access = t.access_token;
    let refresh = t.refresh_token;
    try {
      const r = await refreshToken(t.refresh_token);
      access = r.access_token; refresh = r.refresh_token;
    } catch (e) { console.log('  refresh failed:', e.message); }
    updatedTokens.push({ ...t, access_token: access, refresh_token: refresh });

    try {
      console.log('  Monthly PL...');
      const monthlyPL = await fetchMonthlyPL(access, t.company_id);
      const ytdRev = monthlyPL.reduce((s,r)=>s+r.revenue, 0);
      const ytdOp = monthlyPL.reduce((s,r)=>s+r.operatingProfit, 0);
      console.log(`    months=${monthlyPL.length} 売上累計=${Math.round(ytdRev/10000).toLocaleString()}万 営業利益=${Math.round(ytdOp/10000).toLocaleString()}万`);

      console.log('  BS w/ partners...');
      const bsData = await fetchBSWithPartners(access, t.company_id);
      if (bsData) {
        console.log(`    資産=${bsData.bs.totalAssets.toLocaleString()} 純資産=${bsData.bs.netAssets.toLocaleString()} 現預金=${bsData.bs.cash.toLocaleString()}`);
        if (bsData.bsExtended.denko.net !== 0) console.log(`    DENKO純額=${bsData.bsExtended.denko.net.toLocaleString()}`);
      } else {
        console.log('    BS: no data');
      }

      const co = {
        id: t.company_id,
        name: t.name,
        displayName: t.name,
        monthlyPL,
        budgets: [],
        topCustomers: [],
        topSuppliers: [],
        bs: bsData ? bsData.bs : { totalAssets: 0, cash: 0, totalLiabilities: 0, netAssets: 0 },
        bsExtended: bsData ? bsData.bsExtended : null
      };
      companies.push(co);
    } catch (e) {
      console.log('  ERROR:', e.message);
    }
  }

  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updatedTokens, null, 2));
  const output = {
    fetchedAt: new Date().toISOString(),
    fiscalYear: FY,
    fiscalYearLabel: `FY${FY} (${FY}/05 - ${FY+1}/04)`,
    cutoffMonth: `${cutoffYear}/${String(cutoffMonth).padStart(2,'0')}`,
    companies
  };
  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(output, null, 2));
  console.log(`\n=== Saved: ${OUTPUT_FILE} (${companies.length} companies) ===`);
})();
