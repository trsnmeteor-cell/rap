// ============================================================
//  売上見込 転記ツール（GAS ウェブアプリ） v3 マルチタブ版
//  元シート: ラップ集計(店舗) ブックの店舗タブ（73/53/104/105/248/267）
//  目標:     同ブック 'd' タブ（B3:C8）を参照して各タブ最上部に表示
//  転記先:   店舗ごとに異なる gid のシートへ G列/L列 転記
//
//  - データ行: 各タブ A5:Q30（ヘッダが5行目、データは6行目以降。
//    旧レイアウトのように5行目がデータの場合も自動判定で対応）
//  - 転記先行: 11:00→13行目 〜 5:00→31行目（全タブ共通・規定通り）
//  - 転記済み判定: 「営業日|店舗コード|時刻」キーで PropertiesService に記録
// ============================================================

// ===== 設定 =====
var SOURCE_SHEET_ID = '1QtvvJb80-sablHgWhaYESCpSxX_sEhwS2P6P96VHgGI';
var TARGET_SHEET_ID = '1JG7COuIVeuFABl6xYZCeL2F2pfbaU1i8hx1EZfHhMVo';
var GOAL_SHEET_NAME = 'd';      // 目標シート（B列=店舗名, C列=目標金額, 3行目から）
var DATA_START_ROW  = 5;        // 各タブの読み込み開始行（A5:Q30）
var DATA_END_ROW    = 30;
var DONE_PROP_KEY   = 'TENKI_DONE_MAP_v3';

// タブ定義：tab=元シートのタブ名 / goalName=dタブB列の店舗名 / targetGid=転記先シートのgid
var TABS = [
  { tab: '73',  goalName: '三条',  targetGid: 0 },          // 73 → gid=0（先頭シート）
  { tab: '53',  goalName: '四条',  targetGid: 352300064 },
  { tab: '104', goalName: '本店',  targetGid: 2020072788 },
  { tab: '105', goalName: 'SJK',   targetGid: 1183350844 },
  { tab: '248', goalName: '近江屋', targetGid: 1000553340 },
  { tab: '267', goalName: '木屋町', targetGid: 31976571 }
];

// 時刻 → 転記先の行番号（19枠。23:00の次は0:00に折り返し）
var TIME_TO_ROW = {
  '11:00': 13, '12:00': 14, '13:00': 15, '14:00': 16, '15:00': 17,
  '16:00': 18, '17:00': 19, '18:00': 20, '19:00': 21, '20:00': 22,
  '21:00': 23, '22:00': 24, '23:00': 25, '0:00':  26, '1:00':  27,
  '2:00':  28, '3:00':  29, '4:00':  30, '5:00':  31
};

// フォールバック用ヘッダ名
var FALLBACK_HEADERS = ['エリア名称','店舗コード','店舗名称','営業日','時刻','売上見込金額',
  '累計客数','稼働客数','稼働組数','フード売上','フ客単','ドリンク売上','ド客単',
  '飲み放題客数','飲み放題獲得率（客数）','飲食売上','飲み放題獲得率（売上）'];

// 列インデックス（0始まり）：E=時刻, F=売上見込金額, G=累計客数
var COL_TIME = 4, COL_SALES = 5, COL_CUSTOMERS = 6;

// ---------- ウェブアプリのエントリポイント ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('売上見込 転記ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- 目標一覧の取得（dタブ B3:C8） ----------
function getGoals_() {
  var sheet = SpreadsheetApp.openById(SOURCE_SHEET_ID).getSheetByName(GOAL_SHEET_NAME);
  var goals = {};
  if (!sheet) return goals;
  var values = sheet.getRange('B3:C8').getDisplayValues();
  values.forEach(function(r) {
    var name = String(r[0] || '').trim();
    var num = toNumber_(r[1]);
    if (name) goals[name] = num;
  });
  return goals;
}

// ---------- タブ情報取得（指定タブのデータ＋目標＋転記済みフラグ） ----------
function getStoreData(tabName) {
  var cfg = null;
  TABS.forEach(function(t) { if (t.tab === String(tabName)) cfg = t; });
  if (!cfg) return { error: '不明なタブ: ' + tabName };

  var ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  var sheet = ss.getSheetByName(cfg.tab);
  if (!sheet) return { error: 'タブが見つかりません: ' + cfg.tab };

  // A5:Q30 を読み込み（先頭行がヘッダかデータかを自動判定）
  var values = sheet.getRange('A' + DATA_START_ROW + ':Q' + DATA_END_ROW).getDisplayValues();
  var startIdx = 0;
  if (values.length && /時刻|エリア名称|店舗コード/.test(values[0].join(''))) startIdx = 1;

  var headers = (startIdx === 1 && values[0].join('') !== '')
    ? values[0] : FALLBACK_HEADERS;

  var doneMap = getDoneMap_();
  var rows = [], storeName = '';
  for (var i = startIdx; i < values.length; i++) {
    var r = values[i];
    var norm = normalizeTime_(r[COL_TIME]);
    if (!norm) continue; // 時刻が空の行はスキップ
    if (!storeName) storeName = r[2];
    var key = makeKey_(r[3], r[1], r[COL_TIME]);
    rows.push({
      key: key,
      time: norm,
      rowNo: DATA_START_ROW + i,               // 実際のシート行番号
      tenkiable: !!TIME_TO_ROW.hasOwnProperty.call(TIME_TO_ROW, norm) || !!TIME_TO_ROW[norm],
      done: !!doneMap[key],
      cells: r
    });
  }

  // 最新データ（最後の行）＝現在値
  var latest = null;
  if (rows.length) {
    var last = rows[rows.length - 1];
    latest = {
      time: last.time,
      sales: toNumber_(last.cells[COL_SALES]),
      customers: toNumber_(last.cells[COL_CUSTOMERS])
    };
  }

  var goal = getGoals_()[cfg.goalName] || null;
  return {
    tab: cfg.tab,
    storeName: storeName,
    goal: goal,
    latest: latest,
    headers: headers,
    rows: rows
  };
}

// ---------- 単一時刻の転記（プルダウン用） ----------
function tenki(tabName, selectedTime) {
  var norm = normalizeTime_(selectedTime);
  if (!norm || !TIME_TO_ROW[norm]) {
    return { ok: false, message: '⚠️ 不正な時刻です: ' + selectedTime };
  }
  var data = getStoreData(tabName);
  if (data.error) return { ok: false, message: '⚠️ ' + data.error };

  var found = null;
  data.rows.forEach(function(r) { if (r.time === norm && !found) found = r; });
  if (!found) {
    return { ok: false, message: '⚠️ タブ ' + tabName + ' に時刻 ' + norm + ' のデータが見つかりません' };
  }
  return writeAndMark_(tabName, found);
}

// ---------- チェックした行をすべて転記（一括） ----------
function tenkiAll(tabName, keys) {
  if (!keys || !keys.length) {
    return { ok: false, message: '⚠️ チェックされた行がありません', results: [] };
  }
  var data = getStoreData(tabName);
  if (data.error) return { ok: false, message: '⚠️ ' + data.error, results: [] };

  var results = [], okCount = 0, skipCount = 0, failCount = 0;
  var handled = {};

  keys.forEach(function(key) {
    key = String(key);
    if (handled[key]) return; // 重複スキップ
    handled[key] = true;

    var found = null;
    data.rows.forEach(function(r) { if (r.key === key && !found) found = r; });
    if (!found) {
      results.push({ key: key, status: 'fail', message: '元シートに該当行が見つかりません' });
      failCount++;
      return;
    }
    if (found.done) {
      results.push({ key: key, status: 'skip', message: found.time + ' は転記済みのためスキップ' });
      skipCount++;
      return;
    }
    var res = writeAndMark_(tabName, found);
    results.push({ key: key, status: res.ok ? 'ok' : 'fail', message: res.message });
    if (res.ok) okCount++; else failCount++;
  });

  var summary = '✅ ' + okCount + '件転記';
  if (skipCount) summary += ' ／ スキップ ' + skipCount + '件';
  if (failCount) summary += ' ／ 失敗 ' + failCount + '件';
  return { ok: failCount === 0, message: summary, results: results };
}

// ---------- 転記の実行＋転記済み記録 ----------
function writeAndMark_(tabName, row) {
  var targetGid = null;
  TABS.forEach(function(t) { if (t.tab === String(tabName)) targetGid = t.targetGid; });
  if (!targetGid && targetGid !== 0) {
    return { ok: false, message: '⚠️ 転記先が未定義のタブ: ' + tabName };
  }

  var target = getTargetSheet_(targetGid);
  if (!target) {
    return { ok: false, message: '⚠️ 転記先シート（gid=' + targetGid + '）が見つかりません' };
  }

  var targetRow = TIME_TO_ROW[row.time];
  target.getRange('G' + targetRow).setValue(toNumber_(row.cells[COL_SALES]));
  target.getRange('L' + targetRow).setValue(toNumber_(row.cells[COL_CUSTOMERS]));

  var doneMap = getDoneMap_();
  doneMap[row.key] = new Date().getTime();
  saveDoneMap_(doneMap);

  return {
    ok: true,
    message: '✅ ' + tabName + ' ' + row.time + ' → ' + targetRow + '行目 ｜ 売上見込 ' +
             toNumber_(row.cells[COL_SALES]).toLocaleString() + '円 → G' + targetRow +
             ' ／ 累計客数 ' + toNumber_(row.cells[COL_CUSTOMERS]).toLocaleString() + '名 → L' + targetRow
  };
}

// ---------- 内部ユーティリティ ----------
function getTargetSheet_(gid) {
  var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  if (gid === 0 || gid === '0') return ss.getSheets()[0]; // gid=0 は先頭シート
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
}

function getDoneMap_() {
  var raw = PropertiesService.getScriptProperties().getProperty(DONE_PROP_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function saveDoneMap_(map) {
  PropertiesService.getScriptProperties().setProperty(DONE_PROP_KEY, JSON.stringify(map));
}

// キー生成：営業日｜店舗コード｜時刻
function makeKey_(date, storeCode, time) {
  return String(date || '').trim() + '|' +
         String(storeCode || '').trim() + '|' +
         (normalizeTime_(time) || String(time || '').trim());
}

// "6:00" / "06:00" / Date型 をすべて "h:mm" に正規化
function normalizeTime_(v) {
  if (v instanceof Date) {
    return v.getHours() + ':' + ('0' + v.getMinutes()).slice(-2);
  }
  var m = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return parseInt(m[1], 10) + ':' + m[2];
}

// "31,135" などの表示文字列を数値へ
function toNumber_(v) {
  if (typeof v === 'number') return v;
  var s = String(v).replace(/,/g, '').replace(/[¥￥]/g, '').trim();
  if (s === '' || s === '--') return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}
