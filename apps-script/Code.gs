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
const MAX_SYMBOLS = 100;
const SHEET_NAME = 'quotes';

function doGet(e) {
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
    const fresh = fetchQuotes(missing);
    const toCache = {};
    Object.keys(fresh).forEach(function (s) {
      quotes[s] = fresh[s];
      if (fresh[s].price != null) toCache['q_' + s] = JSON.stringify(fresh[s]);
    });
    if (Object.keys(toCache).length) cache.putAll(toCache, CACHE_SECONDS);
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
