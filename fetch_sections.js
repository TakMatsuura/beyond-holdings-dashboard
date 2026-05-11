/**
 * 各社の事業部別PLを取得し、コストセンターを売上比で按分
 *
 * 仕組み:
 *   1. section_mapping.json でグループ定義（事業部・除外・コストセンター）
 *   2. Freee API で section_id 別の月次累計PLを取得
 *   3. cumulative -> monthly 変換
 *   4. コストセンターの販管費を、各事業部の売上比で按分
 *   5. freee_data.json に各社の `sectionGroups` 配列を追記
 *
 * Usage:
 *   node fetch_sections.js              # 全社（mappingがある会社）
 *   node fetch_sections.js 3115888      # 指定会社のみ
 */
const https = require('https');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const RATE_LIMIT_DELAY = 600;
const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const DATA_FILE = path.join(SCRIPT_DIR, 'freee_data.json');
const MAPPING_FILE = path.join(SCRIPT_DIR, 'section_mapping.json');

// Target FY - current FY (May 2025 - April 2026)
const FY_START_DATE = '2025-05-01';
const FY_START_MONTH = 5;
const FY_YEAR = 2025;

// Cutoff: latest closed month
const now = new Date();
const monthsBack = now.getDate() < 20 ? 2 : 1;
let cutoffYear = now.getFullYear();
let cutoffMonth = now.getMonth() + 1 - monthsBack;
if (cutoffMonth <= 0) { cutoffMonth += 12; cutoffYear--; }

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
  if (r.statusCode === 429) { console.log('    rate limited 60s'); await sleep(60000); return apiGet(token, ep, params); }
  if (r.statusCode === 401) throw new Error('Unauthorized');
  if (r.statusCode >= 400) throw new Error('HTTP ' + r.statusCode + ': ' + JSON.stringify(r.data).substring(0, 200));
  return r.data;
}

// Fetch cumulative PL for a single section_id (or no section = all)
async function fetchCumulativePLSingle(token, companyId, endDate, sectionId) {
  const params = { company_id: companyId, start_date: FY_START_DATE, end_date: endDate };
  if (sectionId) params.section_id = sectionId;
  const data = await apiGet(token, '/reports/trial_pl', params);
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
  return { revenue, cogs, grossProfit, sga, operatingProfit, depreciation: dep };
}

// Fetch cumulative PL for multiple section_ids, sum them
async function fetchCumulativePL(token, companyId, endDate, sectionIds) {
  const sum = { revenue: 0, cogs: 0, grossProfit: 0, sga: 0, operatingProfit: 0, depreciation: 0 };
  for (const sid of sectionIds) {
    const pl = await fetchCumulativePLSingle(token, companyId, endDate, sid);
    sum.revenue += pl.revenue;
    sum.cogs += pl.cogs;
    sum.grossProfit += pl.grossProfit;
    sum.sga += pl.sga;
    sum.operatingProfit += pl.operatingProfit;
    sum.depreciation += pl.depreciation;
  }
  return sum;
}

// Cumulative -> Monthly conversion
function cumToMonthly(cumByMonth, months) {
  const monthly = [];
  let prev = null;
  for (const mk of months) {
    const cur = cumByMonth[mk];
    if (!cur) continue;
    const m = prev ? {
      month: mk,
      revenue: cur.revenue - prev.revenue,
      cogs: cur.cogs - prev.cogs,
      grossProfit: cur.grossProfit - prev.grossProfit,
      sga: cur.sga - prev.sga,
      operatingProfit: cur.operatingProfit - prev.operatingProfit,
      depreciation: cur.depreciation - prev.depreciation,
    } : { month: mk, ...cur };
    m.ebitda = m.operatingProfit + m.depreciation;
    monthly.push(m);
    prev = cur;
  }
  return monthly;
}

// Process one company
async function processCompany(token, companyId, mapping) {
  console.log(`\n[${mapping.displayName}] (id=${companyId})`);

  // Determine month list (FY 12 months capped at cutoff)
  const months = [];
  for (let i = 0; i < 12; i++) {
    let y = FY_YEAR, m = FY_START_MONTH + i;
    if (m > 12) { m -= 12; y++; }
    if (y > cutoffYear || (y === cutoffYear && m > cutoffMonth)) break;
    months.push({ year: y, month: m, key: `${y}/${String(m).padStart(2,'0')}` });
  }

  // For each group's section_ids, fetch cumulative PL per month, then convert to monthly
  const fetchGroupMonthly = async (sectionIds) => {
    const cum = {};
    for (const m of months) {
      const lastDay = new Date(m.year, m.month, 0).getDate();
      const end = `${m.year}-${String(m.month).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
      const pl = await fetchCumulativePL(token, companyId, end, sectionIds);
      cum[m.key] = pl;
    }
    return cumToMonthly(cum, months.map(m => m.key));
  };

  // Process each group
  const groups = [];
  for (const g of mapping.groups) {
    console.log(`  → ${g.name} (sections: ${g.sectionIds.join(',')})...`);
    const monthlyPL = await fetchGroupMonthly(g.sectionIds);
    const ytdRev = monthlyPL.reduce((s,r)=>s+r.revenue, 0);
    const ytdOp = monthlyPL.reduce((s,r)=>s+r.operatingProfit, 0);
    console.log(`    YTD: 売上=${Math.round(ytdRev/10000).toLocaleString()}万 営業利益=${Math.round(ytdOp/10000).toLocaleString()}万`);
    groups.push({ ...g, monthlyPL });
  }

  // Cost centers (按分対象)
  let costCenterMonthly = null;
  if (mapping.costCenters && mapping.costCenters.length) {
    const ccIds = mapping.costCenters.map(c => c.sectionId);
    console.log(`  → コストセンター (sections: ${ccIds.join(',')})...`);
    costCenterMonthly = await fetchGroupMonthly(ccIds);
    const ccSGA = costCenterMonthly.reduce((s,r)=>s+r.sga, 0);
    console.log(`    YTD 販管費=${Math.round(ccSGA/10000).toLocaleString()}万 (按分対象)`);
  }

  // Allocation: 売上比で按分
  // For each month, allocate cost center's expenses (sga, operatingProfit) proportionally by group revenue
  if (costCenterMonthly && mapping.allocationMethod === 'revenue') {
    for (const m of months) {
      const ccM = costCenterMonthly.find(r => r.month === m.key);
      if (!ccM) continue;
      const groupRevs = groups.map(g => {
        const gM = g.monthlyPL.find(r => r.month === m.key);
        return gM ? gM.revenue : 0;
      });
      const totalRev = groupRevs.reduce((s, v) => s + v, 0);
      if (totalRev <= 0) continue; // 売上ゼロなら按分しない（均等按分する選択肢もあるが安全側で）
      groups.forEach((g, i) => {
        const ratio = groupRevs[i] / totalRev;
        const allocSga = ccM.sga * ratio;
        const allocCogs = ccM.cogs * ratio;
        const allocDep = ccM.depreciation * ratio;
        const gM = g.monthlyPL.find(r => r.month === m.key);
        if (!gM) return;
        gM.allocatedSga = allocSga;
        gM.allocatedCogs = allocCogs;
        gM.allocatedDep = allocDep;
        // 配賦後の営業利益
        gM.operatingProfitAllocated = gM.operatingProfit - allocSga - allocCogs;
        gM.ebitdaAllocated = gM.operatingProfitAllocated + (gM.depreciation + allocDep);
      });
    }
  }

  // Return sectionGroups data
  return {
    sectionGroups: groups.map(g => ({
      name: g.name,
      color: g.color,
      note: g.note,
      sectionIds: g.sectionIds,
      monthlyPL: g.monthlyPL
    })),
    costCenterMonthly,
    excluded: mapping.exclude || [],
    allocationMethod: mapping.allocationMethod || 'none'
  };
}

(async () => {
  const targetId = process.argv[2] ? parseInt(process.argv[2]) : null;
  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  const dashData = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  const allMapping = JSON.parse(fs.readFileSync(MAPPING_FILE, 'utf8'));
  const updatedTokens = [];

  for (const t of tokens) {
    const mapping = allMapping[String(t.company_id)];
    if (!mapping) { updatedTokens.push(t); continue; }
    if (targetId && t.company_id !== targetId) { updatedTokens.push(t); continue; }

    let access = t.access_token;
    try {
      const r = await refresh(t.refresh_token);
      access = r.access_token;
      updatedTokens.push({ ...t, access_token: r.access_token, refresh_token: r.refresh_token });
    } catch (e) {
      console.log(`[${t.name}] refresh fail, using existing:`, e.message);
      updatedTokens.push(t);
    }

    try {
      const result = await processCompany(access, t.company_id, mapping);
      const co = dashData.companies.find(c => c.id === t.company_id);
      if (co) {
        co.sectionGroups = result.sectionGroups;
        co.costCenterMonthly = result.costCenterMonthly;
        co.sectionMeta = { excluded: result.excluded, allocationMethod: result.allocationMethod };
      }
    } catch (e) {
      console.log(`[${t.name}] ERROR:`, e.message);
    }
  }

  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updatedTokens, null, 2));
  dashData.fetchedAt = new Date().toISOString();
  fs.writeFileSync(DATA_FILE, JSON.stringify(dashData, null, 2));
  console.log('\n=== Saved ===');
})();
