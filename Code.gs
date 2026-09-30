/**
 * 売上見込ラップ速報 ― Google Apps Script 版
 * データ元：公開スプレッドシートのCSV（/pub?output=csv）
 * 構成：Code.gs（サーバー側：データ取得・整形） + Index.html（画面）
 * デプロイ：デプロイ → 新しいデプロイ → 種類「ウェブアプリ」 → アクセスできるユーザーを選択
 */
const PUB_CSV = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTRUlQFACBOztP-ljb3MJCssZ5ZD69p-ccLzotZgNepZaY99QB3ulqp0ftOeWPvc7GlUUBhg8j-DVho/pub?output=csv';
const CACHE_SEC = 300; // 5分キャッシュ

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('売上見込ラップ速報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 画面から google.script.run.getData() で呼ばれる */
function getData() {
  try {
    const cache = CacheService.getScriptCache();
    const hit = cache.get('lap_csv');
    if (hit) return { ok: true, data: normalize(parseCSV(hit)), cached: true };
    const res = UrlFetchApp.fetch(PUB_CSV, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    const text = res.getContentText('UTF-8');
    cache.put('lap_csv', text, CACHE_SEC);
    return { ok: true, data: normalize(parseCSV(text)), cached: false };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/** RFC4180 準拠の簡易CSVパーサ（改行を含む引用符フィールド対応） */
function parseCSV(text) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else { q = false; } }
      else { f += c; }
    } else {
      if (c === '"') q = true;
      else if (c === ',') { row.push(f); f = ''; }
      else if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; }
      else if (c !== '\r') f += c;
    }
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}

function toNum(s) {
  const v = parseFloat(String(s == null ? '' : s).replace(/,/g, '').replace(/%/g, '').trim());
  return isNaN(v) ? 0 : v;
}

/** 元シートの並び：ヘッダー行 → （店舗行 ×N → エリア合計行）×14 → 末尾にヘッダー再掲 */
function normalize(rows) {
  let hi = -1;
  for (let i = 0; i < Math.min(rows.length, 6); i++) {
    if (String(rows[i][0] || '').trim() === 'エリア名称') { hi = i; break; }
  }
  if (hi < 0) throw new Error('ヘッダー行（エリア名称）が見つかりません');
  const stores = [], areaTotals = [];
  let day = '', time = '';
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || r.length < 6) continue;
    const area = String(r[0] || '').trim(), name = String(r[2] || '').trim();
    if (area === 'エリア名称' || name === '店舗名称') continue;
    if (!area && !name) {
      if (String(r[5] || '').trim()) areaTotals.push({
        name: '', sales: toNum(r[5]), cust: toNum(r[6]), work: toNum(r[7]), groups: toNum(r[8]),
        food: toNum(r[9]), funit: toNum(r[10]), drink: toNum(r[11]), dunit: toNum(r[12]),
        nomi: toNum(r[13]), nomiRate: toNum(r[14]), nomiSales: toNum(r[15]), nomiRate2: toNum(r[16])
      });
      continue;
    }
    if (!day && String(r[3] || '').trim()) day = String(r[3]).trim();
    if (!time && String(r[4] || '').trim()) time = String(r[4]).trim();
    stores.push({
      area: area, no: toNum(r[1]), name: name, day: String(r[3] || '').trim(), time: String(r[4] || '').trim(),
      sales: toNum(r[5]), cust: toNum(r[6]), work: toNum(r[7]), groups: toNum(r[8]),
      food: toNum(r[9]), funit: toNum(r[10]), drink: toNum(r[11]), dunit: toNum(r[12]),
      nomi: toNum(r[13]), nomiRate: toNum(r[14]), nomiSales: toNum(r[15]), nomiRate2: toNum(r[16])
    });
  }
  const order = [];
  stores.forEach(function (s) { if (order.indexOf(s.area) < 0) order.push(s.area); });
  areaTotals.forEach(function (t, i) { t.name = order[i] || ('エリア ' + (i + 1)); });
  return { day: day, time: time, stores: stores, areaTotals: areaTotals };
}
