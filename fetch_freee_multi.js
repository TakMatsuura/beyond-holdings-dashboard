/**
 * BEYOND Holdings Dashboard - Multi-Company Freee Data Fetcher
 * Uses per-company access tokens from freee_tokens.json
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const API_BASE = 'https://api.freee.co.jp/api/1';
const RATE_LIMIT_DELAY = 600;
const MAX_RETRIES = 3;

const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const OUTPUT_FILE = path.join(SCRIPT_DIR, 'freee_data.json');

// Fiscal year: April start
// Use PREVIOUS completed fiscal year for full 12-month data
// FY2025 = 2025/04 - 2026/03
const now = new Date();
const currentYear = now.getFullYear();
const currentMonth = now.getMonth() + 1;
const fiscalYear = (currentMonth >= 4 ? currentYear : currentYear - 1) - 1; // Previous FY

console.log('=== BEYOND Holdings Freee Multi-Company Fetch ===');
console.log('Fiscal Year:', fiscalYear, `(${fiscalYear}/04 - ${fiscalYear + 1}/03)`);

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
  const url = new URL(API_BASE + endpoint);
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
    if (res.statusCode === 401) {
      throw new Error('Unauthorized (token expired?)');
    }
    if (res.statusCode >= 400) {
      throw new Error('HTTP ' + res.statusCode + ': ' + JSON.stringify(res.data).substring(0, 200));
    }
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

async function fetchMonthlyPL(token, companyId) {
  const results = [];
  for (let i = 0; i < 12; i++) {
    let year = fiscalYear;
    let month = 4 + i; // April start
    if (month > 12) { month -= 12; year += 1; }
    const range = getMonthRange(year, month);

    try {
      const data = await apiGet(token, '/reports/trial_pl', {
        company_id: companyId,
        start_date: range.start,
        end_date: range.end
      });

      const balances = data.trial_pl?.balances || [];
      const find = (name) => {
        const item = balances.find(b => b.account_category_name === name);
        return item ? Math.abs(item.closing_balance || 0) : 0;
      };

      const revenue = find('売上高');
      const cogs = find('売上原価');
      const grossProfit = revenue - cogs;
      const sga = find('販売費及び一般管理費');
      const operatingProfit = grossProfit - sga;

      results.push({
        month: `${year}/${String(month).padStart(2, '0')}`,
        revenue, cogs, grossProfit, sga, operatingProfit
      });
      console.log(`    ${range.start}: rev=${(revenue/10000).toFixed(0)}万 op=${(operatingProfit/10000).toFixed(0)}万`);
    } catch (e) {
      console.log(`    ${range.start}: ERROR - ${e.message}`);
      results.push({
        month: `${year}/${String(month).padStart(2, '0')}`,
        revenue: 0, cogs: 0, grossProfit: 0, sga: 0, operatingProfit: 0
      });
    }
  }
  return results;
}

async function fetchPartners(token, companyId) {
  try {
    // Use fiscal_year parameter instead of date range to avoid cross-year error
    console.log(`    Partner fiscal_year: ${fiscalYear}`);
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
  console.log(`Loaded ${tokens.length} company tokens\n`);

  const companies = [];
  const updatedTokens = [];

  for (const t of tokens) {
    console.log(`\n--- ${t.name} (company_id=${t.company_id}) ---`);

    // Refresh token to get fresh access
    let accessToken = t.access_token;
    let newRefreshToken = t.refresh_token;
    try {
      console.log('  Refreshing token...');
      const refreshed = await refreshAccessToken(t.refresh_token);
      accessToken = refreshed.access_token;
      newRefreshToken = refreshed.refresh_token;
      console.log('  Token refreshed OK');
    } catch (e) {
      console.log('  Token refresh failed, using existing access token:', e.message);
    }

    updatedTokens.push({ ...t, access_token: accessToken, refresh_token: newRefreshToken });

    // Fetch PL
    console.log('  Fetching monthly PL...');
    const monthlyPL = await fetchMonthlyPL(accessToken, t.company_id);

    // Fetch partners
    console.log('  Fetching partners...');
    const { topCustomers, topSuppliers } = await fetchPartners(accessToken, t.company_id);

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

  // Save updated tokens (with refreshed refresh_tokens)
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updatedTokens, null, 2));
  console.log('\nTokens updated');

  // Save dashboard data
  const output = {
    fetchedAt: new Date().toISOString(),
    fiscalYear,
    fiscalYearLabel: `FY${fiscalYear} (${fiscalYear}/04 - ${fiscalYear + 1}/03)`,
    companies
  };

  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(output, null, 2));
  console.log(`Dashboard data saved: ${OUTPUT_FILE}`);
  console.log(`Companies: ${companies.length}`);
  console.log('=== Done ===');
}

main().catch(err => {
  console.error('FATAL:', err.message);
  process.exit(1);
});
