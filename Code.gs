// ============================================================
//  売上見込 転記ツール（GAS ウェブアプリ） v3.1 マルチタブ版
//  元シート: ラップ集計(店舗) ブックの店舗タブ（73/53/104/105/248/267）
//  目標:     同ブック 'd' タブ（B3:C8）を参照して各タブ最上部に表示
//  転記先:   店舗ごとに異なる gid のシートへ G列/L列 転記
//
//  v3.1 変更点:
//   - doGet / 各関数を try/catch で包み、エラーを画面に表示
//   - ?debug=1 でJSON診断出力（ブラウザで直接開いて確認できる）
//   - google.script.run 用ラッパー関数を導入（サーバーエラーを
//     失敗ハンドラではなく正常応答として返し、必ず画面に表示させる）
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
function doGet(e) {
  try {
    // ?debug=1 を付けるとJSONで生データを確認できる（動作診断用）
    if (e && e.parameter && e.parameter.debug) {
      var out = ContentService.createTextOutput(
        JSON.stringify({
          ok: true,
          version: 'v3.1',
          time: new Date().toISOString(),
          tabs: TABS,
          sample73: getStoreData('73')
        }, null, 2)
      );
      out.setMimeType(ContentService.MimeType.JSON);
      return out;
    }
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('売上見込 転記ツール')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    // doGet が失敗すると原因不明の画面になるため、原因を画面に出す
    var msg = String(err && err.stack ? err.stack : err);
    return HtmlService.createHtmlOutput(
      '<h2 style="color:#c00;font-family:sans-serif;">⚠️ アプリ初期化エラー</h2>' +
      '<pre style="white-space:pre-wrap;border:1px solid #ccc;padding:10px;font-size:12px;">' + msg + '</pre>' +
      '<p style="font-family:sans-serif;">確認ポイント：HTMLファイル名が「index」か／Code.gsとindex.htmlが同じバージョンか／新しいバージョンでデプロイしたか</p>'
    ).setTitle('エラー');
  }
}

// ============================================================
//  google.script.run 用ラッパー（サーバーエラーを必ず画面に返す）
//  ※ 画面側は getStoreData / tenki / tenkiAll を呼び出す
// ============================================================
function getStoreData(tabName) {
  try {
    return getStoreData_(tabName);
  } catch (err) {
    return { error: 'サーバーエラー(getStoreData): ' + errMsg_(err) };
  }
}
function tenki(tabName, selectedTime) {
  try {
    return tenki_(tabName, selectedTime);
  } catch (err) {
    return { ok: false, message: '❌ サーバーエラー(tenki): ' + errMsg_(err) };
  }
}
function tenkiAll(tabName, keys) {
  try {
    return tenkiAll_(tabName, keys);
  } catch (err) {
    return { ok: false, message: '❌ サーバーエラー(tenkiAll): ' + errMsg_(err), results: [] };
  }
}

function errMsg_(err) {
  return String(err && err.message ? err.message : err);
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
function getStoreData_(tabName) {
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

  // 先に全行の営業日を読んで「最新の営業日」を求める
  // （日次クリアの実行前に前日の残骸行が残っている場合、
  //   そちらを誤って転記・表示しないための判定）
  var dayKeys = [];
  for (var d = 0; d < values.length; d++) {
    dayKeys.push(toDayKey_(values[d][3]));
  }
  var maxDay = '';
  for (var d = 0; d < dayKeys.length; d++) {
    if (dayKeys[d] > maxDay) maxDay = dayKeys[d];
  }

  var rows = [], storeName = '';
  for (var i = startIdx; i < values.length; i++) {
    var r = values[i];
    var norm = normalizeTime_(r[COL_TIME]);
    if (!norm) continue; // 時刻が空の行はスキップ
    if (dayKeys[i] === maxDay && !storeName) storeName = r[2];
    var key = makeKey_(r[3], r[1], r[COL_TIME]);
    rows.push({
      isToday: dayKeys[i] === maxDay,
      key: key,
      time: norm,
      rowNo: DATA_START_ROW + i,   // 実際のシート行番号
      tenkiable: !!TIME_TO_ROW[norm],
      done: !!doneMap[key],
      cells: r
    });
  }

  // 最新データ＝「最新の営業日」の中で時刻が最も進んでいる行
  // （24:00→0:00 の折り返しも含めて正しい時刻順で比較する）
  var TIME_ORDER = ['6:00','7:00','8:00','9:00','10:00','11:00','12:00','13:00','14:00','15:00',
                    '16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00',
                    '0:00','1:00','2:00','3:00','4:00','5:00'];
  var latest = null;
  for (var i = 0; i < rows.length; i++) {
    var cand = rows[i];
    if (!cand.isToday) continue; // 前日以前の行は「現在値」に使わない
    if (!latest || TIME_ORDER.indexOf(cand.time) > TIME_ORDER.indexOf(latest.time)) {
      latest = {
        time: cand.time,
        sales: toNumber_(cand.cells[COL_SALES]),
        customers: toNumber_(cand.cells[COL_CUSTOMERS])
      };
    }
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
function tenki_(tabName, selectedTime) {
  var norm = normalizeTime_(selectedTime);
  if (!norm || !TIME_TO_ROW[norm]) {
    return { ok: false, message: '⚠️ 不正な時刻です: ' + selectedTime };
  }
  var data = getStoreData_(tabName);
  if (data.error) return { ok: false, message: '⚠️ ' + data.error };

  // 該当時刻の行を探す。
  // 前日の残骸行が残っている場合は「最新の営業日」の行を優先する
  // （旧版は最初に一致した行＝前日の行を転記してしまうバグがあった）
  var found = null;
  data.rows.forEach(function(r) {
    if (r.time !== norm) return;
    if (!found || (r.isToday && !found.isToday)) found = r;
  });
  if (!found) {
    return { ok: false, message: '⚠️ タブ ' + tabName + ' に時刻 ' + norm + ' のデータが見つかりません' };
  }
  return writeAndMark_(tabName, found);
}

// ---------- チェックした行をすべて転記（一括） ----------
function tenkiAll_(tabName, keys) {
  if (!keys || !keys.length) {
    return { ok: false, message: '⚠️ チェックされた行がありません', results: [] };
  }
  var data = getStoreData_(tabName);
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
    if (!found.isToday) {
      results.push({ key: key, status: 'skip', message: found.time + '（前日以前のデータ）はスキップ' });
      skipCount++;
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

// "6:00" / "06:00" / "24:00" / Date型 をすべて "h:mm" に正規化
// ※ 24時以降の表記（"24:00"＝翌日0:00）は "0:00" に変換する
//   （旧版は "24:00" が正規化されず、0:00の転記が失敗するバグがあった）
function normalizeTime_(v) {
  var h, m;
  if (v instanceof Date) {
    h = v.getHours();
    m = v.getMinutes();
  } else {
    var mt = String(v).trim().match(/^(\d{1,2}):(\d{2})/);
    if (!mt) return null;
    h = parseInt(mt[1], 10);
    m = parseInt(mt[2], 10);
  }
  if (h >= 24) h = h - 24; // "24:00" → "0:00"
  return h + ':' + ('0' + m).slice(-2);
}

// 営業日の表示（"2026/11/04" / "2026-11-04" / Date型 など）を
// 比較可能な "YYYY/MM/DD" 文字列に正規化。
// ※ 日付として解釈できない文字列（ヘッダ行の「営業日」など）は
//   空文字を返す。これを怠ると文字コード順で「営業日」が最大と
//   なり、全行が「前日扱い」になるバグの原因になる。
function toDayKey_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + ('0' + v.getDate()).slice(-2);
  }
  var m = String(v || '').trim().match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (m) {
    return m[1] + '/' + ('0' + parseInt(m[2], 10)).slice(-2) + '/' + ('0' + parseInt(m[3], 10)).slice(-2);
  }
  return ''; // 日付でない文字（ヘッダ等）は無視
}

// "31,135" などの表示文字列を数値へ
function toNumber_(v) {
  if (typeof v === 'number') return v;
  var s = String(v).replace(/,/g, '').replace(/[¥￥]/g, '').trim();
  if (s === '' || s === '--') return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}
