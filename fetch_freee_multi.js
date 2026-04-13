/**
 * BEYOND Holdings Dashboard - Multi-Company Freee Data Fetcher
 * Uses per-company access tokens from freee_tokens.json
 *
 * Key behaviors:
 * - Freee trial_pl returns CUMULATIVE values from fiscal year start
 * - We convert cumulative -> monthly by subtracting consecutive months
 * - Only includes months where closing (締め) is complete
 *   Rule: if today < 20th, last closed month = 2 months ago
 *          if today >= 20th, last closed month = 1 month ago
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const RATE_LIMIT_DELAY = 600;
const MAX_RETRIES = 3;

// FY start month overrides (when Freee API returns incorrect values)
// 184: established mid-FY, Freee returns wrong FY start. Actual: May start (4月末締め)
const FY_START_OVERRIDES = {
  12243427: 5  // 184: 5月スタート
};

const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const OUTPUT_FILE = path.join(SCRIPT_DIR, 'freee_data.json');

// Determine fiscal year and cutoff
const now = new Date();
const todayYear = now.getFullYear();
const todayMonth = now.getMonth() + 1;
const todayDay = now.getDate();

// Cutoff: if before 20th, last closed = 2 months ago; if >= 20th, 1 month ago
const monthsBack = todayDay < 20 ? 2 : 1;
let cutoffYear = todayYear;
let cutoffMonth = todayMonth - monthsBack;
if (cutoffMonth <= 0) { cutoffMonth += 12; cutoffYear--; }

// Fiscal year that contains the cutoff month (April start)
const fiscalYear = cutoffMonth >= 4 ? cutoffYear : cutoffYear - 1;
const fyStartMonth = 4;

console.log('=== BEYOND Holdings Freee Multi-Company Fetch ===');
console.log(`Today: ${todayYear}/${todayMonth}/${todayDay}`);
console.log(`Last closed month: ${cutoffYear}/${String(cutoffMonth).padStart(2, '0')}`);
console.log(`Fiscal Year: ${fiscalYear} (${fiscalYear}/04 - ${fiscalYear + 1}/03)`);

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function httpsRequest(urlStr, options, body) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlStr);
    const opts = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: options.method || 'GET',
      headers: options.headers || {}
    };
    if (body) opts.headers['Content-Length'] = Buffer.byteLength(body);

    const req = https.request(opts, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          resolve({ statusCode: res.statusCode, data: JSON.parse(data) });
        } catch (e) {
          resolve({ statusCode: res.statusCode, data: { raw: data.substring(0, 300) } });
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function refreshAccessToken(refreshToken) {
  const postData = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    refresh_token: refreshToken
  }).toString();

  const res = await httpsRequest('https://accounts.secure.freee.co.jp/public_api/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  }, postData);

  if (res.data.error) throw new Error('Token refresh failed: ' + res.data.error);
  return { access_token: res.data.access_token, refresh_token: res.data.refresh_token };
}

async function apiGet(accessToken, endpoint, params) {
  const url = new URL('https://api.freee.co.jp/api/1' + endpoint);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, String(v)));

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    await sleep(RATE_LIMIT_DELAY);
    const res = await httpsRequest(url.toString(), {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Accept': 'application/json' }
    });

    if (res.statusCode === 429) {
      console.log('    Rate limited, waiting 60s...');
      await sleep(60000);
      continue;
    }
    if (res.statusCode === 401) throw new Error('Unauthorized');
    if (res.statusCode >= 400) throw new Error('HTTP ' + res.statusCode + ': ' + JSON.stringify(res.data).substring(0, 200));
    return res.data;
  }
  throw new Error('Rate limit exceeded');
}

function getMonthRange(year, month) {
  const start = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const end = `${year}-${String(month).padStart(2, '0')}-${lastDay}`;
  return { start, end };
}

// Check if a month is within the cutoff
function isMonthClosed(year, month) {
  if (year < cutoffYear) return true;
  if (year === cutoffYear && month <= cutoffMonth) return true;
  return false;
}

// Get company's fiscal year start month from Freee
async function getCompanyFYStart(accessToken, companyId) {
  // Check override first
  if (FY_START_OVERRIDES[companyId]) {
    console.log(`    Using FY start override: month ${FY_START_OVERRIDES[companyId]}`);
    return FY_START_OVERRIDES[companyId];
  }
  try {
    const data = await apiGet(accessToken, '/companies/' + companyId, {});
    const fy = data.company?.fiscal_years;
    if (fy && fy.length > 0) {
      const startMonth = parseInt(fy[0].start_date?.split('-')[1]) || 4;
      return startMonth;
    }
  } catch (e) {
    console.log('    Could not get FY start, defaulting to April');
  }
  return 4; // Default April
}

async function fetchMonthlyPL(token, companyId, companyFYStart) {
  // Determine the 12 months of this fiscal year
  const months = [];
  for (let i = 0; i < 12; i++) {
    let y = fiscalYear;
    let m = companyFYStart + i;
    if (m > 12) { m -= 12; y++; }
    months.push({ year: y, month: m });
  }

  // Fetch cumulative values for each month
  const cumulativeValues = [];
  for (const { year, month } of months) {
    if (!isMonthClosed(year, month)) {
      console.log(`    ${year}/${String(month).padStart(2, '0')}: skipped (not yet closed)`);
      cumulativeValues.push(null);
      continue;
    }

    const range = getMonthRange(year, month);
    try {
      // Query from FY start to end of this month to get cumulative
      const fyStartDate = `${fiscalYear}-${String(companyFYStart).padStart(2, '0')}-01`;
      const data = await apiGet(token, '/reports/trial_pl', {
        company_id: companyId,
        start_date: fyStartDate,
        end_date: range.end
      });

      const balances = data.trial_pl?.balances || [];
      const find = (name) => {
        const item = balances.find(b => b.account_category_name === name);
        return item ? Math.abs(item.closing_balance || 0) : 0;
      };

      cumulativeValues.push({
        revenue: find('売上高'),
        cogs: find('売上原価'),
        sga: find('販売費及び一般管理費')
      });
      console.log(`    ${range.start}: cumRev=${(cumulativeValues[cumulativeValues.length - 1].revenue / 10000).toFixed(0)}万`);
    } catch (e) {
      console.log(`    ${range.start}: ERROR - ${e.message}`);
      cumulativeValues.push(null);
    }
  }

  // Convert cumulative to monthly by subtracting previous month
  const monthlyPL = [];
  for (let i = 0; i < months.length; i++) {
    const { year, month } = months[i];
    const monthStr = `${year}/${String(month).padStart(2, '0')}`;

    if (cumulativeValues[i] === null) {
      // Skip unclosed months entirely
      continue;
    }

    let revenue, cogs, sga;
    if (i === 0 || cumulativeValues[i - 1] === null) {
      // First month of FY or first available month
      revenue = cumulativeValues[i].revenue;
      cogs = cumulativeValues[i].cogs;
      sga = cumulativeValues[i].sga;
    } else {
      // Monthly = current cumulative - previous cumulative
      revenue = cumulativeValues[i].revenue - cumulativeValues[i - 1].revenue;
      cogs = cumulativeValues[i].cogs - cumulativeValues[i - 1].cogs;
      sga = cumulativeValues[i].sga - cumulativeValues[i - 1].sga;
    }

    const grossProfit = revenue - cogs;
    const operatingProfit = grossProfit - sga;

    monthlyPL.push({ month: monthStr, revenue, cogs, grossProfit, sga, operatingProfit });
    console.log(`    → ${monthStr}: monthly rev=${(revenue / 10000).toFixed(0)}万 op=${(operatingProfit / 10000).toFixed(0)}万`);
  }

  return monthlyPL;
}

async function fetchPartners(token, companyId) {
  try {
    const data = await apiGet(token, '/reports/trial_pl', {
      company_id: companyId,
      fiscal_year: fiscalYear,
      breakdown_display_type: 'partner'
    });

    const balances = data.trial_pl?.balances || [];
    const revItem = balances.find(b => b.account_category_name === '売上高');
    const cogsItem = balances.find(b => b.account_category_name === '売上原価');

    const topCustomers = (revItem?.partners || [])
      .filter(p => Math.abs(p.closing_balance) > 0)
      .sort((a, b) => Math.abs(b.closing_balance) - Math.abs(a.closing_balance))
      .slice(0, 10)
      .map(p => ({ name: p.partner_name, revenue: Math.abs(p.closing_balance) }));

    const topSuppliers = (cogsItem?.partners || [])
      .filter(p => Math.abs(p.closing_balance) > 0)
      .sort((a, b) => Math.abs(b.closing_balance) - Math.abs(a.closing_balance))
      .slice(0, 10)
      .map(p => ({ name: p.partner_name, cost: Math.abs(p.closing_balance) }));

    return { topCustomers, topSuppliers };
  } catch (e) {
    console.log(`    Partners: ERROR - ${e.message}`);
    return { topCustomers: [], topSuppliers: [] };
  }
}

async function fetchBS(token, companyId) {
  try {
    const data = await apiGet(token, '/reports/trial_bs', {
      company_id: companyId,
      fiscal_year: fiscalYear
    });

    const balances = data.trial_bs?.balances || [];
    const find = (name) => {
      const item = balances.find(b => b.account_category_name === name);
      return item ? Math.abs(item.closing_balance || 0) : 0;
    };

    return {
      totalAssets: find('資産'),
      cash: find('現金・預金'),
      totalLiabilities: find('負債'),
      netAssets: find('純資産')
    };
  } catch (e) {
    console.log(`    BS: ERROR - ${e.message}`);
    return { totalAssets: 0, cash: 0, totalLiabilities: 0, netAssets: 0 };
  }
}

async function main() {
  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  console.log(`\nLoaded ${tokens.length} company tokens\n`);

  const companies = [];
  const updatedTokens = [];

  for (const t of tokens) {
    console.log(`\n=== ${t.name} (company_id=${t.company_id}) ===`);

    // Refresh token
    let accessToken = t.access_token;
    let newRefreshToken = t.refresh_token;
    try {
      const refreshed = await refreshAccessToken(t.refresh_token);
      accessToken = refreshed.access_token;
      newRefreshToken = refreshed.refresh_token;
      console.log('  Token refreshed OK');
    } catch (e) {
      console.log('  Token refresh failed, using existing:', e.message);
    }
    updatedTokens.push({ ...t, access_token: accessToken, refresh_token: newRefreshToken });

    // Get company FY start month
    console.log('  Getting FY info...');
    const companyFYStart = await getCompanyFYStart(accessToken, t.company_id);
    console.log(`  FY starts: month ${companyFYStart}`);

    // Fetch monthly PL (cumulative -> monthly conversion)
    console.log('  Fetching monthly PL...');
    const monthlyPL = await fetchMonthlyPL(accessToken, t.company_id, companyFYStart);

    // Fetch partners
    console.log('  Fetching partners...');
    const { topCustomers, topSuppliers } = await fetchPartners(accessToken, t.company_id);
    console.log(`  Customers: ${topCustomers.length}, Suppliers: ${topSuppliers.length}`);

    // Fetch BS
    console.log('  Fetching BS...');
    const bs = await fetchBS(accessToken, t.company_id);

    companies.push({
      id: t.company_id,
      name: t.name,
      displayName: t.name,
      monthlyPL,
      budgets: [],
      topCustomers,
      topSuppliers,
      bs
    });
  }

  // Save updated tokens
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updatedTokens, null, 2));

  // Save dashboard data
  const output = {
    fetchedAt: new Date().toISOString(),
    fiscalYear,
    fiscalYearLabel: `FY${fiscalYear} (${fiscalYear}/04 - ${fiscalYear + 1}/03)`,
    cutoffMonth: `${cutoffYear}/${String(cutoffMonth).padStart(2, '0')}`,
    companies
  };

  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(output, null, 2));
  console.log(`\n=== Complete ===`);
  console.log(`Output: ${OUTPUT_FILE}`);
  console.log(`Companies: ${companies.length}`);
  console.log(`Data through: ${cutoffYear}/${String(cutoffMonth).padStart(2, '0')}`);
}

main().catch(err => {
  console.error('FATAL:', err.message);
  process.exit(1);
});
