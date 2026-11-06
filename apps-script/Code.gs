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

  if (e.parameter.history) {
    const from = /^\d{4}-\d{2}-\d{2}$/.test(String(e.parameter.from || '')) ? String(e.parameter.from) : '2026-06-01';
    return readHistory(parseSymbols(e.parameter.history), from);
  }

  const symbols = parseSymbols(e.parameter.symbols);

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

// 只允許代號常見字元，防止被塞入公式
function parseSymbols(raw) {
  return Array.from(new Set(
    String(raw || '').toUpperCase().split(',').map(function (s) { return s.trim(); })
       .filter(function (s) { return /^[A-Z0-9.\-\/:]{1,20}$/.test(s); })
  )).slice(0, MAX_SYMBOLS);
}

// ---------- 歷史收盤價（總價值走勢圖用） ----------
// GOOGLEFINANCE 的歷史資料無法從 Apps Script 讀取，所以改向 Yahoo Finance 取每日收盤價，快取 3 小時。
// 回傳 {history: {AAPL: [["2026-06-01", 201.5], ...], ...}}
function readHistory(symbols, from) {
  const cache = CacheService.getScriptCache();
  const start = Math.floor(new Date(from + 'T00:00:00Z').getTime() / 1000);
  const now = Math.floor(Date.now() / 1000);
  const out = {};
  const need = [];
  symbols.forEach(function (s) {
    const hit = cache.get('h_' + from + '_' + s);
    if (hit) out[s] = JSON.parse(hit); else need.push(s);
  });
  if (need.length) {
    const reqs = need.map(function (s) {
      return {
        url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(toYahooSymbol(s)) +
             '?period1=' + start + '&period2=' + now + '&interval=1d',
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' },
      };
    });
    const res = UrlFetchApp.fetchAll(reqs);
    res.forEach(function (r, i) {
      const s = need[i];
      const rows = [];
      try {
        const d = JSON.parse(r.getContentText()).chart.result[0];
        const ts = d.timestamp || [];
        const close = (d.indicators.quote[0] || {}).close || [];
        ts.forEach(function (t, j) {
          if (close[j] != null) rows.push([Utilities.formatDate(new Date(t * 1000), 'America/New_York', 'yyyy-MM-dd'), Math.round(close[j] * 100) / 100]);
        });
      } catch (err) { /* 沒有資料的代號回傳空陣列 */ }
      out[s] = rows;
      if (rows.length) cache.put('h_' + from + '_' + s, JSON.stringify(rows), 10800);
    });
  }
  return json({ history: out });
}

// 券商寫法 → Yahoo 寫法，例如 BRK/B → BRK-B
function toYahooSymbol(s) {
  return s.replace('/', '-').replace('.', '-');
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

// 在編輯器裡按「執行」：第一次會要求「連接到外部服務」的授權，允許後走勢圖才抓得到歷史股價
function testHistory() {
  Logger.log(readHistory(['SPY', 'AAPL'], '2026-06-01').getContent().slice(0, 300));
}

// 在編輯器裡按「執行」測試用
function testQuotes() {
  Logger.log(JSON.stringify(fetchQuotes(['AAPL', 'BRK/B', 'IVV', 'SPCX'])));
}
