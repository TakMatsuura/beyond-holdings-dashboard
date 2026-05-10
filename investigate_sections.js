/**
 * 各社のFreee部門(sections)を調査
 */
const https = require('https');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const TOKENS_FILE = path.join(__dirname, 'freee_tokens.json');

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
  await sleep(600);
  const r = await httpsRequest(url.toString(), { method: 'GET', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' } });
  if (r.statusCode >= 400) return { error: r.statusCode, data: r.data };
  return r.data;
}

(async () => {
  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  const updated = [];
  console.log('=== 各社の事業部(sections)調査 ===\n');
  for (const t of tokens) {
    let access = t.access_token;
    try {
      const r = await refresh(t.refresh_token);
      access = r.access_token;
      updated.push({ ...t, access_token: r.access_token, refresh_token: r.refresh_token });
    } catch (e) { updated.push(t); console.log(`[${t.name}] refresh fail`); continue; }

    const data = await apiGet(access, '/sections', { company_id: t.company_id });
    if (data.error) { console.log(`[${t.name}] API error ${data.error}`); continue; }
    const sections = data.sections || [];
    console.log(`[${t.name}] ${sections.length}部門:`);
    sections.forEach(s => {
      const indent = s.parent_id ? '    └ ' : '  ';
      console.log(`${indent}${s.id}: ${s.name}${s.shortcut1 ? ' (' + s.shortcut1 + ')' : ''}${s.parent_id ? ' [親=' + s.parent_id + ']' : ''}`);
    });
    if (!sections.length) console.log('  (部門設定なし)');
    console.log();
  }
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updated, null, 2));
})();
