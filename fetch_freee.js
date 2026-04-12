/**
 * BEYOND Holdings Dashboard - Freee API Data Fetcher
 *
 * Fetches financial data from Freee Accounting API for all subsidiaries.
 * Requires OAuth2 credentials as environment variables.
 *
 * Environment variables:
 *   FREEE_CLIENT_ID     - OAuth2 client ID
 *   FREEE_CLIENT_SECRET - OAuth2 client secret
 *   FREEE_REFRESH_TOKEN - OAuth2 refresh token
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const TOKEN_URL = 'https://accounts.secure.freee.co.jp/public_api/token';
const API_BASE = 'https://api.freee.co.jp/api/1';
const OUTPUT_DIR = '/tmp/freee_data';
const RATE_LIMIT_DELAY = 500; // ms between API calls
const MAX_RETRIES = 3;

// Fiscal year config
const FISCAL_YEAR_START_MONTH = 4; // April
const now = new Date();
const currentYear = now.getFullYear();
const currentMonth = now.getMonth() + 1;
const fiscalYear = currentMonth >= FISCAL_YEAR_START_MONTH ? currentYear : currentYear - 1;
const fyStart = `${fiscalYear}-${String(FISCAL_YEAR_START_MONTH).padStart(2, '0')}-01`;
const fyEnd = `${fiscalYear + 1}-${String(FISCAL_YEAR_START_MONTH - 1).padStart(2, '0')}-31`;

console.log(`=== BEYOND Holdings Freee Data Fetch ===`);
console.log(`Fiscal Year: ${fiscalYear} (${fyStart} ~ ${fyEnd})`);

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function httpsRequest(url, options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(url, options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode === 429) {
          resolve({ statusCode: 429, data: null });
        } else if (res.statusCode >= 200 && res.statusCode < 300) {
          try { resolve({ statusCode: res.statusCode, data: JSON.parse(data) }); }
          catch (e) { reject(new Error(`JSON parse error: ${e.message}`)); }
        } else {
          reject(new Error(`HTTP ${res.statusCode}: ${data.substring(0, 200)}`));
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function refreshToken() {
  const clientId = process.env.FREEE_CLIENT_ID;
  const clientSecret = process.env.FREEE_CLIENT_SECRET;
  const refreshToken = process.env.FREEE_REFRESH_TOKEN;

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('Missing FREEE_CLIENT_ID, FREEE_CLIENT_SECRET, or FREEE_REFRESH_TOKEN');
  }

  const postData = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken
  }).toString();

  const url = new URL(TOKEN_URL);
  const res = await httpsRequest(url, {
    method: 'POST',
    hostname: url.hostname,
    path: url.pathname,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(postData) }
  }, postData);

  const newRefreshToken = res.data.refresh_token;
  if (newRefreshToken) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    fs.writeFileSync(path.join(OUTPUT_DIR, 'new_refresh_token.txt'), newRefreshToken, 'utf8');
    console.log('New refresh token saved');
  }

  return res.data.access_token;
}

async function apiCall(accessToken, endpoint, params) {
  const url = new URL(API_BASE + endpoint);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, v));

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    const res = await httpsRequest(url, {
      method: 'GET',
      hostname: url.hostname,
      path: url.pathname + url.search,
      headers: { 'Authorization': `Bearer ${accessToken}`, 'Accept': 'application/json' }
    });

    if (res.statusCode === 429) {
      console.log(`  Rate limited, waiting 60s (attempt ${attempt + 1}/${MAX_RETRIES})...`);
      await sleep(60000);
      continue;
    }
    return res.data;
  }
  throw new Error(`Rate limit exceeded after ${MAX_RETRIES} retries for ${endpoint}`);
}

function getMonthRange(year, month) {
  const start = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const end = `${year}-${String(month).padStart(2, '0')}-${lastDay}`;
  return { start, end };
}

function extractPLValue(balances, categoryName) {
  const item = balances.find(b =>
    b.account_category_name === categoryName ||
    b.parent_account_category_name === categoryName
  );
  return item ? (item.closing_balance || 0) : 0;
}

async function fetchCompanyPL(accessToken, companyId) {
  const monthlyPL = [];
  for (let i = 0; i < 12; i++) {
    let year = fiscalYear;
    let month = FISCAL_YEAR_START_MONTH + i;
    if (month > 12) { month -= 12; year += 1; }
    const range = getMonthRange(year, month);

    await sleep(RATE_LIMIT_DELAY);
    try {
      const data = await apiCall(accessToken, '/reports/trial_pl', {
        company_id: companyId,
        start_date: range.start,
        end_date: range.end
      });

      const balances = data.trial_pl?.balances || [];
      const revenue = extractPLValue(balances, '売上高');
      const cogs = extractPLValue(balances, '売上原価');
      const grossProfit = revenue - cogs;
      const sga = extractPLValue(balances, '販売費及び一般管理費');
      const operatingProfit = grossProfit - sga;

      monthlyPL.push({
        month: `${year}/${String(month).padStart(2, '0')}`,
        revenue, cogs, grossProfit, sga, operatingProfit
      });
      console.log(`    ${range.start}: rev=${revenue}, op=${operatingProfit}`);
    } catch (e) {
      console.error(`    ${range.start}: ERROR - ${e.message}`);
      monthlyPL.push({
        month: `${year}/${String(month).padStart(2, '0')}`,
        revenue: 0, cogs: 0, grossProfit: 0, sga: 0, operatingProfit: 0
      });
    }
  }
  return monthlyPL;
}

async function fetchPartnerBreakdown(accessToken, companyId) {
  await sleep(RATE_LIMIT_DELAY);
  try {
    const data = await apiCall(accessToken, '/reports/trial_pl', {
      company_id: companyId,
      start_date: fyStart,
      end_date: fyEnd,
      breakdown_display_type: 'partner'
    });

    const balances = data.trial_pl?.balances || [];
    const revenueItem = balances.find(b => b.account_category_name === '売上高');
    const cogsItem = balances.find(b => b.account_category_name === '売上原価');

    const topCustomers = (revenueItem?.partners || [])
      .filter(p => p.closing_balance > 0)
      .sort((a, b) => b.closing_balance - a.closing_balance)
      .slice(0, 10)
      .map(p => ({ name: p.partner_name, revenue: p.closing_balance }));

    const topSuppliers = (cogsItem?.partners || [])
      .filter(p => p.closing_balance > 0)
      .sort((a, b) => b.closing_balance - a.closing_balance)
      .slice(0, 10)
      .map(p => ({ name: p.partner_name, cost: p.closing_balance }));

    return { topCustomers, topSuppliers };
  } catch (e) {
    console.error(`    Partner breakdown: ERROR - ${e.message}`);
    return { topCustomers: [], topSuppliers: [] };
  }
}

async function fetchBS(accessToken, companyId) {
  await sleep(RATE_LIMIT_DELAY);
  try {
    const data = await apiCall(accessToken, '/reports/trial_bs', {
      company_id: companyId,
      fiscal_year: fiscalYear
    });

    const balances = data.trial_bs?.balances || [];
    const findVal = (name) => {
      const item = balances.find(b => b.account_category_name === name);
      return item ? (item.closing_balance || 0) : 0;
    };

    return {
      totalAssets: findVal('資産'),
      cash: findVal('現金・預金') || findVal('流動資産') * 0.3,
      totalLiabilities: findVal('負債'),
      netAssets: findVal('純資産')
    };
  } catch (e) {
    console.error(`    BS: ERROR - ${e.message}`);
    return { totalAssets: 0, cash: 0, totalLiabilities: 0, netAssets: 0 };
  }
}

async function main() {
  // Step 1: Get access token
  console.log('\n[1/4] Refreshing OAuth2 token...');
  const accessToken = await refreshToken();
  console.log('Token obtained');

  // Step 2: Get companies
  console.log('\n[2/4] Fetching companies...');
  await sleep(RATE_LIMIT_DELAY);
  const companiesData = await apiCall(accessToken, '/companies', {});
  const companies = companiesData.companies || [];
  console.log(`Found ${companies.length} companies`);

  // Step 3: Fetch data for each company
  console.log('\n[3/4] Fetching financial data...');
  const results = [];
  for (const co of companies) {
    console.log(`\n  Processing: ${co.display_name || co.name} (ID: ${co.id})`);

    const monthlyPL = await fetchCompanyPL(accessToken, co.id);
    const { topCustomers, topSuppliers } = await fetchPartnerBreakdown(accessToken, co.id);
    const bs = await fetchBS(accessToken, co.id);

    results.push({
      id: co.id,
      name: co.name || co.display_name,
      displayName: co.display_name || co.name,
      monthlyPL,
      budgets: [], // Budgets loaded from budget.csv by build_html.js
      topCustomers,
      topSuppliers,
      bs
    });
  }

  // Step 4: Write output
  console.log('\n[4/4] Writing output...');
  const output = {
    fetchedAt: new Date().toISOString(),
    fiscalYear,
    fiscalYearLabel: `FY${fiscalYear} (${fiscalYear}/${String(FISCAL_YEAR_START_MONTH).padStart(2, '0')} - ${fiscalYear + 1}/${String(FISCAL_YEAR_START_MONTH - 1).padStart(2, '0')})`,
    companies: results
  };

  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(path.join(OUTPUT_DIR, 'dashboard_data.json'), JSON.stringify(output, null, 2), 'utf8');
  console.log(`\nOutput: ${OUTPUT_DIR}/dashboard_data.json`);
  console.log(`Companies: ${results.length}`);
  console.log('=== Fetch complete ===');
}

main().catch(err => {
  console.error('FATAL:', err.message);
  process.exit(1);
});
