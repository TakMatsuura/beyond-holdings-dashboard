/**
 * Extended fetch: 取引先別BS（DENKO識別） + 前年同月PL
 * 既存の freee_data.json にマージ更新する
 */
const https = require('https');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = '705227068726896';
const CLIENT_SECRET = '1YLBoILMxQUtYkj45xI4lFqy7VKP8I91Z4CY3Y_RI8rQz9ShSr1Kyxr_0MTIk1M4nqlLeryQOjvFWfcCsRqkYA';
const RATE_LIMIT_DELAY = 600;
const SCRIPT_DIR = __dirname;
const TOKENS_FILE = path.join(SCRIPT_DIR, 'freee_tokens.json');
const OUTPUT_FILE = path.join(SCRIPT_DIR, 'freee_data.json');

// BS基準日: 最新締め月末
const BS_END_DATE = '2026-03-31';
const FY_START_DATE_MAIN = '2025-05-01';     // 通常会社
const FY_START_DATE_184 = '2025-11-13';      // 184のみ
const PREV_FY_START = '2024-05-01';
const PREV_FY_END   = '2025-04-30';

// DENKO本体は除外（社内貸付の対象外）
const DENKO_COMPANY_ID = 3115888;

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
async function apiGet(token, endpoint, params, opts={}) {
  const url = new URL('https://api.freee.co.jp/api/1' + endpoint);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, String(v)));
  await sleep(RATE_LIMIT_DELAY);
  const r = await httpsRequest(url.toString(), { method: 'GET', headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' } });
  if (r.statusCode === 429) { console.log('    Rate limited, retry...'); await sleep(60000); return apiGet(token, endpoint, params, opts); }
  if (r.statusCode === 401) throw new Error('Unauthorized (token expired)');
  if (r.statusCode >= 400) {
    if (opts.allowError) return null;
    throw new Error('HTTP ' + r.statusCode + ': ' + JSON.stringify(r.data).substring(0, 300));
  }
  return r.data;
}

// 取引先別BSを取得し、DENKO関連／社外借入を分離
async function fetchBSWithPartners(token, companyId) {
  const startDate = (companyId === 12243427) ? FY_START_DATE_184 : FY_START_DATE_MAIN;
  const data = await apiGet(token, '/reports/trial_bs', {
    company_id: companyId,
    start_date: startDate,
    end_date: BS_END_DATE,
    breakdown_display_type: 'partner'
  });
  const balances = data.trial_bs.balances || [];

  // 集計関数
  const find = (name, level) => {
    const item = balances.find(b => b.account_category_name === name && (level == null || b.hierarchy_level === level));
    return item ? (item.closing_balance || 0) : 0;
  };

  // 自己資本比率系: 資産・負債・純資産 (符号保持)
  const totalAssets = find('資産', 1);
  const totalLiabilities = find('負債', 1);
  const netAssets = find('純資産', 1);
  // 現預金は集計行がないので、account_category_name='現金・預金'の全項目を合計
  const cash = balances
    .filter(b => b.account_category_name === '現金・預金')
    .reduce((s, b) => s + (b.closing_balance || 0), 0);

  // 借入/貸付/立替 関連の取引先別残高
  const LOAN_ACCOUNTS_LIABILITY = ['短期借入金','長期借入金','役員借入金'];
  const LOAN_ACCOUNTS_ASSET = ['短期貸付金','長期貸付金','立替金'];
  const isDenkoPartner = (p) => /デンコー|DENKO|電工/i.test(p.name || '');

  // companyIdがDENKO本体の場合、自分自身を除外する必要があるが、
  // DENKOがDENKOの取引先になることはないので問題なし

  let denkoLiability = 0;       // DENKOが当社へ貸している額 (負債側)
  let denkoAssetReverse = 0;     // 当社がDENKOへ貸している/立替 (資産側、逆向き)
  let externalShortTerm = 0;
  let externalLongTerm = 0;
  let externalOfficer = 0;

  const denkoDetail = [];

  for (const b of balances) {
    const name = b.account_item_name || '';
    const isLoanLiab = LOAN_ACCOUNTS_LIABILITY.some(k => name.includes(k));
    const isLoanAsset = LOAN_ACCOUNTS_ASSET.some(k => name.includes(k));
    if (!isLoanLiab && !isLoanAsset) continue;

    const partners = b.partners || [];
    let denkoBalanceForThisAccount = 0;
    let externalBalanceForThisAccount = 0;
    for (const p of partners) {
      if (isDenkoPartner(p)) {
        denkoBalanceForThisAccount += (p.closing_balance || 0);
      } else if (isLoanLiab) {
        externalBalanceForThisAccount += (p.closing_balance || 0);
      }
    }

    if (denkoBalanceForThisAccount !== 0) {
      denkoDetail.push({
        account: name,
        side: isLoanLiab ? 'liability' : 'asset',
        amount: denkoBalanceForThisAccount
      });
      if (isLoanLiab) denkoLiability += denkoBalanceForThisAccount;
      else denkoAssetReverse += denkoBalanceForThisAccount;
    }
    if (isLoanLiab) {
      if (name.includes('短期借入金')) externalShortTerm += externalBalanceForThisAccount;
      else if (name.includes('長期借入金')) externalLongTerm += externalBalanceForThisAccount;
      else if (name.includes('役員借入金')) externalOfficer += externalBalanceForThisAccount;
    }
  }

  // DENKO貸付残（純額）= 借入金系(DENKOから)  -  資産側(DENKOへ/立替)
  const denkoNet = denkoLiability - denkoAssetReverse;

  return {
    totalAssets, cash, totalLiabilities, netAssets,
    denko: {
      liability: denkoLiability,        // 借入金系のDENKO分(当社が借りている)
      assetReverse: denkoAssetReverse,  // 貸付金/立替金のDENKO分(逆方向)
      net: denkoNet,                    // 純額: 当社のDENKO純借入残
      detail: denkoDetail
    },
    external: {
      shortTerm: externalShortTerm,
      longTerm: externalLongTerm,
      officer: externalOfficer,
      total: externalShortTerm + externalLongTerm + externalOfficer
    }
  };
}

// 前年同期間のPL (前年同月比計算用)
async function fetchPrevFYMonthly(token, companyId, fyStartMonth) {
  // FY2024: 2024/05 - 2025/04 (通常)
  // 184は前年なし
  if (companyId === 12243427) return [];

  const months = [];
  for (let i = 0; i < 12; i++) {
    let y = 2024;
    let m = fyStartMonth + i;
    if (m > 12) { m -= 12; y++; }
    months.push({ year: y, month: m });
  }

  const cumulative = [];
  for (const { year, month } of months) {
    const lastDay = new Date(year, month, 0).getDate();
    const end = `${year}-${String(month).padStart(2, '0')}-${lastDay}`;
    try {
      const data = await apiGet(token, '/reports/trial_pl', {
        company_id: companyId,
        start_date: PREV_FY_START,
        end_date: end
      }, { allowError: true });
      if (!data) { cumulative.push(null); continue; }
      const balances = data.trial_pl?.balances || [];
      const findH1 = (name) => {
        const item = balances.find(b => b.account_category_name === name && b.hierarchy_level === 1);
        return item ? (item.closing_balance || 0) : 0;
      };
      const findH2 = (name) => {
        const item = balances.find(b => b.account_category_name === name && b.hierarchy_level === 2);
        return item ? (item.closing_balance || 0) : 0;
      };
      cumulative.push({
        year, month,
        revenue: findH1('売上高'),
        operatingProfit: findH1('営業利益')
      });
    } catch (e) {
      cumulative.push(null);
    }
  }

  // cumulative -> monthly
  const monthly = [];
  for (let i = 0; i < cumulative.length; i++) {
    const cur = cumulative[i];
    if (!cur) { monthly.push(null); continue; }
    if (i === 0) {
      monthly.push({ month: `${cur.year}/${String(cur.month).padStart(2, '0')}`, revenue: cur.revenue, operatingProfit: cur.operatingProfit });
    } else {
      const prev = cumulative[i - 1];
      if (!prev) { monthly.push(null); continue; }
      monthly.push({
        month: `${cur.year}/${String(cur.month).padStart(2, '0')}`,
        revenue: cur.revenue - prev.revenue,
        operatingProfit: cur.operatingProfit - prev.operatingProfit
      });
    }
  }
  return monthly.filter(x => x !== null);
}

(async () => {
  console.log('=== Extended fetch: 取引先別BS + 前年同月PL ===\n');
  const tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
  const existing = JSON.parse(fs.readFileSync(OUTPUT_FILE, 'utf8'));
  const updatedTokens = [];

  for (const t of tokens) {
    console.log(`\n[${t.name}] (id=${t.company_id})`);
    let access = t.access_token;
    let refresh = t.refresh_token;
    try {
      const r = await refreshToken(t.refresh_token);
      access = r.access_token; refresh = r.refresh_token;
    } catch (e) {
      console.log(`  refresh failed: ${e.message} -- 既存トークン使用`);
    }
    updatedTokens.push({ ...t, access_token: access, refresh_token: refresh });

    // BS取得
    let bsExt = null;
    try {
      console.log('  Fetching BS with partners...');
      bsExt = await fetchBSWithPartners(access, t.company_id);
      console.log(`    資産=${bsExt.totalAssets.toLocaleString()} 負債=${bsExt.totalLiabilities.toLocaleString()} 純資産=${bsExt.netAssets.toLocaleString()}`);
      console.log(`    DENKO 借入(負債側)=${bsExt.denko.liability.toLocaleString()} 立替/貸付(資産側)=${bsExt.denko.assetReverse.toLocaleString()} 純額=${bsExt.denko.net.toLocaleString()}`);
      if (bsExt.denko.detail.length) {
        bsExt.denko.detail.forEach(d => console.log(`      - ${d.account} (${d.side}): ${d.amount.toLocaleString()}`));
      }
      console.log(`    外部借入: 短期=${bsExt.external.shortTerm.toLocaleString()} 長期=${bsExt.external.longTerm.toLocaleString()} 役員=${bsExt.external.officer.toLocaleString()}`);
    } catch (e) {
      console.log(`  BS fetch FAILED: ${e.message}`);
    }

    // 前年PL取得
    let prevFYPL = [];
    try {
      console.log('  Fetching prev FY PL...');
      prevFYPL = await fetchPrevFYMonthly(access, t.company_id, t.company_id === 12243427 ? 11 : 5);
      console.log(`    Months fetched: ${prevFYPL.length}`);
    } catch (e) {
      console.log(`  Prev FY fetch FAILED: ${e.message}`);
    }

    // existing にマージ
    const co = existing.companies.find(c => c.id === t.company_id);
    if (co) {
      if (bsExt) {
        co.bs = {
          totalAssets: bsExt.totalAssets,
          cash: bsExt.cash,
          totalLiabilities: bsExt.totalLiabilities,
          netAssets: bsExt.netAssets
        };
        co.bsExtended = {
          denko: bsExt.denko,
          external: bsExt.external
        };
      }
      co.prevFYMonthlyPL = prevFYPL;
    }
  }

  // save
  fs.writeFileSync(TOKENS_FILE, JSON.stringify(updatedTokens, null, 2));
  existing.fetchedAt = new Date().toISOString();
  fs.writeFileSync(OUTPUT_FILE, JSON.stringify(existing, null, 2));
  console.log('\n=== Complete ===');
  console.log('Saved to:', OUTPUT_FILE);
})();
