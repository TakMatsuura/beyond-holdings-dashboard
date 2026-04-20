/**
 * Partial fetcher: Update only specific companies in freee_data.json
 * Usage: node fetch_freee_partial.js <company_id_1> <company_id_2> ...
 *
 * Reuses all logic from fetch_freee_multi.js but only processes
 * the specified company IDs, then merges results into existing freee_data.json.
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const RATE_LIMIT_DELAY = 600;
const MAX_RETRIES = 3;

const DEFAULT_FY_START_MONTH = 5;
const FY_START_DATE_OVERRIDES = {
  12243427: { month: 11, startDate: '2025-11-13' }
};

const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const OUTPUT_FILE = path.join(SCRIPT_DIR, 'freee_data.json');

// Parse target company IDs from args
const targetIds = process.argv.slice(2).map(Number).filter(n => !isNaN(n));
if (targetIds.length === 0) {
  console.error('Usage: node fetch_freee_partial.js <company_id_1> <company_id_2> ...');
  process.exit(1);
}

const now = new Date();
const todayYear = now.getFullYear();
const todayMonth = now.getMonth() + 1;
const todayDay = now.getDate();
const monthsBack = todayDay < 20 ? 2 : 1;
let cutoffYear = todayYear;
let cutoffMonth = todayMonth - monthsBack;
if (cutoffMonth <= 0) { cutoffMonth += 12; cutoffYear--; }
const fiscalYear = cutoffMonth >= 4 ? cutoffYear : cutoffYear - 1;

console.log('=== BEYOND Holdings Freee Partial Fetch ===');
console.log(`Today: ${todayYear}/${todayMonth}/${todayDay}`);
console.log(`Last closed month: ${cutoffYear}/${String(cutoffMonth).padStart(2, '0')}`);
console.log(`Target company IDs: ${targetIds.join(', ')}`);

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

async function refreshAccessToken(refreshToken) {
  const postData = new URLSearchParams({
    grant_type: 'refresh_token', client_id: CLIENT_ID, client_secret: CLIENT_SECRET, refresh_token: refreshToken
  }).toString();
  const res = await httpsRequest('https://accounts.secure.freee.co.jp/public_api/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
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
      method: 'GET', headers: { 'Authorization': 'Bearer ' + accessToken, 'Accept': 'application/json' }
    });
    if (res.statusCode === 429) { console.log('    Rate limited, waiting 60s...'); await sleep(60000); continue; }
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

function isMonthClosed(year, month) {
  if (year < cutoffYear) return true;
  if (year === cutoffYear && month <= cutoffMonth) return true;
  return false;
}

function getCompanyFYStart(companyId) {
  const override = FY_START_DATE_OVERRIDES[companyId];
  if (override) return override.month;
  return DEFAULT_FY_START_MONTH;
}

// Import the full logic from fetch_freee_multi.js by requiring pieces
// We need to duplicate fetchMonthlyPL, fetchPartners, fetchBS here.
// To avoid duplicating ~200 lines, we'll load and eval the main script.

const multiSrc = fs.readFileSync(path.join(SCRIPT_DIR, 'fetch_freee_multi.js'), 'utf8');
// Extract the three fetch functions as strings
function extractFn(src, name) {
  const re = new RegExp(`async function ${name}[\\s\\S]*?\\n\\}\\n`, 'm');
  const m = src.match(re);
  if (!m) throw new Error('Function not found: ' + name);
  return m[0];
}

// eval the three functions into scope
eval(extractFn(multiSrc, 'fetchMonthlyPL'));
eval(extractFn(multiSrc, 'fetchPartners'));
eval(extractFn(multiSrc, 'fetchBS'));

async function main() {
  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  const existing = JSON.parse(fs.readFileSync(OUTPUT_FILE, 'utf8'));

  const targetTokens = tokens.filter(t => targetIds.includes(t.company_id));
  if (targetTokens.length === 0) {
    console.error('No matching companies found in freee_tokens.json');
    process.exit(1);
  }
  console.log(`\nProcessing ${targetTokens.length} companies: ${targetTokens.map(t => t.name).join(', ')}\n`);

  const updatedTokens = [...tokens];

  for (const t of targetTokens) {
    console.log(`\n=== ${t.name} (company_id=${t.company_id}) ===`);
    let accessToken = t.access_token;
    let newRefreshToken = t.refresh_token;
    try {
      const refreshed = await refreshAccessToken(t.refresh_token);
      accessToken = refreshed.access_token;
      newRefreshToken = refreshed.refresh_token;
      console.log('  Token refreshed OK');
    } catch (e) {
      console.log('  Token refresh failed:', e.message);
    }
    // update token in array
    const tokIdx = updatedTokens.findIndex(x => x.company_id === t.company_id);
    updatedTokens[tokIdx] = { ...t, access_token: accessToken, refresh_token: newRefreshToken };

    const companyFYStart = getCompanyFYStart(t.company_id);
    console.log('  Fetching monthly PL...');
    const monthlyPL = await fetchMonthlyPL(accessToken, t.company_id, companyFYStart);
    console.log('  Fetching partners...');
    const { topCustomers, topSuppliers } = await fetchPartners(accessToken, t.company_id);
    console.log('  Fetching BS...');
    const bs = await fetchBS(accessToken, t.company_id);

    // Preserve existing budgets
    const existingCo = existing.companies.find(c => c.id === t.company_id);
    const newRecord = {
      id: t.company_id,
      name: t.name,
      displayName: t.name,
      monthlyPL,
      budgets: existingCo ? (existingCo.budgets || []) : [],
      topCustomers,
      topSuppliers,
      bs
    };

    // Merge into existing.companies
    const coIdx = existing.companies.findIndex(c => c.id === t.company_id);
    if (coIdx >= 0) existing.companies[coIdx] = newRecord;
    else existing.companies.push(newRecord);
  }

  // Save tokens
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updatedTokens, null, 2));

  // Save merged data (update fetchedAt and cutoff if newer)
  existing.fetchedAt = new Date().toISOString();
  existing.cutoffMonth = `${cutoffYear}/${String(cutoffMonth).padStart(2, '0')}`;
  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(existing, null, 2));

  console.log(`\n=== Partial fetch complete ===`);
  console.log(`Updated ${targetTokens.length} companies through ${existing.cutoffMonth}`);
}

main().catch(err => { console.error('FATAL:', err.message); process.exit(1); });
