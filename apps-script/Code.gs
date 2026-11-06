/**
 * 股票報價 API（Google Apps Script，綁在一張 Google 試算表上）
 *
 * 用法：GET <部署網址>?key=<授權碼>&symbols=AAPL,MSFT,BRK/B
 * 回傳：{"asOf":"...","quotes":{"AAPL":{"price":343.63,"prevClose":349.51,"time":"..."}, ...}}
 *
 * 價格來自 GOOGLEFINANCE（美股約延遲 15–20 分鐘），每個代號快取 5 分鐘。
 * 只回傳股價，不接收也不保存任何持倉資料。
 */

const CACHE_SECONDS = 300;
const STALE_SECONDS = 21600;
const MAX_SYMBOLS = 100;
const SHEET_NAME = 'quotes';

// ---------- 分享短連結 ----------
// 存：POST {key, data}（需要授權碼）→ {id}
// 讀：GET ?share=<id>（不需要授權碼，朋友打開連結時用）→ {data}
const SHARE_SHEET = 'shares';
const SHARE_MAX_CHARS = 30000;

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json({ error: 'bad_request' }); }
  const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!apiKey || String(body.key || '') !== apiKey) return json({ error: 'unauthorized' });
  const text = JSON.stringify(body.data || null);
  if (!body.data || text.length > SHARE_MAX_CHARS) return json({ error: 'bad_request' });

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHARE_SHEET) || ss.insertSheet(SHARE_SHEET);
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  sh.appendRow([id, text, new Date()]);
  return json({ id: id });
}

function readShare(id) {
  if (!/^[A-Za-z0-9]{6,24}$/.test(id)) return json({ error: 'not_found' });
  const cache = CacheService.getScriptCache();
  const hit = cache.get('share_' + id);
  if (hit) return json({ data: JSON.parse(hit) });
  const sh = SpreadsheetApp.getActive().getSheetByName(SHARE_SHEET);
  if (!sh || sh.getLastRow() < 1) return json({ error: 'not_found' });
  const found = sh.getRange(1, 1, sh.getLastRow(), 1).createTextFinder(id).matchEntireCell(true).findNext();
  if (!found) return json({ error: 'not_found' });
  const text = sh.getRange(found.getRow(), 2).getValue();
  if (String(text).length < 90000) cache.put('share_' + id, String(text), 21600);
  return json({ data: JSON.parse(text) });
}

function doGet(e) {
  if (e && e.parameter && e.parameter.share) return readShare(String(e.parameter.share));

  // 只有帶正確授權碼的請求才回應。授權碼設在「專案設定 → 指令碼屬性」的 API_KEY。
  const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  const given = String((e && e.parameter && e.parameter.key) || '');
  if (!apiKey || given !== apiKey) return json({ error: 'unauthorized' });

  const raw = String((e && e.parameter && e.parameter.symbols) || '').toUpperCase();
  // 只允許代號常見字元，防止被塞入公式
  const symbols = Array.from(new Set(
    raw.split(',').map(function (s) { return s.trim(); })
       .filter(function (s) { return /^[A-Z0-9.\-\/:]{1,20}$/.test(s); })
  )).slice(0, MAX_SYMBOLS);

  const cache = CacheService.getScriptCache();
  const quotes = {};
  const missing = [];
  const cached = symbols.length ? cache.getAll(symbols.map(function (s) { return 'q_' + s; })) : {};
  symbols.forEach(function (s) {
    const hit = cached['q_' + s];
    if (hit) quotes[s] = JSON.parse(hit);
    else missing.push(s);
  });

  if (missing.length) {
    let fresh = null;
    try { fresh = fetchQuotes(missing); } catch (err) { fresh = null; }
    if (fresh) {
      const toCache = {};
      Object.keys(fresh).forEach(function (s) {
        quotes[s] = fresh[s];
        if (fresh[s].price != null) {
          toCache['q_' + s] = JSON.stringify(fresh[s]);
          toCache['s_' + s] = JSON.stringify(fresh[s]);
        }
      });
      // q_ = 新鮮報價（5 分鐘）；s_ = 備用的上一筆報價（6 小時），忙碌或出錯時拿來回傳
      const fresh5 = {}, stale6h = {};
      Object.keys(toCache).forEach(function (k) { (k.indexOf('q_') === 0 ? fresh5 : stale6h)[k] = toCache[k]; });
      if (Object.keys(fresh5).length) cache.putAll(fresh5, CACHE_SECONDS);
      if (Object.keys(stale6h).length) cache.putAll(stale6h, STALE_SECONDS);
    } else {
      const stale = cache.getAll(missing.map(function (s) { return 's_' + s; }));
      missing.forEach(function (s) {
        const hit = stale['s_' + s];
        if (hit) { const q = JSON.parse(hit); q.stale = true; quotes[s] = q; }
      });
    }
  }

  return json({ asOf: new Date().toISOString(), quotes: quotes });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fetchQuotes(symbols) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    sh.clear();
    const rows = symbols.map(function (s) {
      const g = toGoogleSymbol(s);
      return [
        s,
        '=IFERROR(GOOGLEFINANCE("' + g + '","price"),"")',
        '=IFERROR(GOOGLEFINANCE("' + g + '","closeyest"),"")',
        '=IFERROR(TEXT(GOOGLEFINANCE("' + g + '","tradetime"),"yyyy-mm-dd hh:mm"),"")',
      ];
    });
    const range = sh.getRange(1, 1, rows.length, 4);
    range.setValues(rows);

    let values = [];
    for (let attempt = 0; attempt < 5; attempt++) {
      SpreadsheetApp.flush();
      values = range.getValues();
      const loading = values.some(function (r) { return String(r[1]).indexOf('Loading') === 0; });
      if (!loading) break;
      Utilities.sleep(1000);
    }

    const out = {};
    values.forEach(function (r) {
      const price = typeof r[1] === 'number' ? r[1] : null;
      const prev = typeof r[2] === 'number' ? r[2] : null;
      out[r[0]] = { price: price, prevClose: prev, time: r[3] || null };
    });
    return out;
  } finally {
    lock.releaseLock();
  }
}

// 券商寫法 → Google Finance 寫法，例如 BRK/B → BRK.B
function toGoogleSymbol(s) {
  return s.replace('/', '.');
}

// 在編輯器裡按「執行」測試用
function testQuotes() {
  Logger.log(JSON.stringify(fetchQuotes(['AAPL', 'BRK/B', 'IVV', 'SPCX'])));
}
