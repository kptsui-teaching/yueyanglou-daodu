/**
 * 《岳陽樓記》範文導讀 — 學生成績接收（Google Apps Script）
 *
 * 使用方式：
 * 1. 在老師的 Google 試算表中：擴充功能 → Apps Script，貼上本檔全部內容。
 * 2. （可選）若本腳本已綁定該試算表，可把 SPREADSHEET_ID 留空字串；
 *    否則填入試算表 ID（網址 /d/ 與 /edit 之間那一段）。
 * 3. 在編輯器執行 setup() 一次，建立「交卷紀錄」分頁與表頭。
 * 4. 部署 → 新增部署 → 類型選「網頁應用程式」：
 *    - 執行身分：我（Execute as: Me）
 *    - 具有存取權的使用者：任何人（Anyone）
 * 5. 把部署後的網址貼入 index.html 的 const GAS_WEBAPP_URL = "...";
 */

// 若腳本綁定於目標試算表，可留空；否則填入試算表 ID
var SPREADSHEET_ID = '122b_4B_g5x1e8EHPUl6Y-u-xmUT4n3aFIE6usQUCQdA';
var SHEET_NAME = '交卷紀錄';

var HEADERS = [
  '時間', '班別', '學號', '姓名',
  '內容理解得分', '字詞解釋得分', '寫作手法得分',
  '總分', '滿分', '百分比', '備註'
];

function ss_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  return sh;
}

/** 首次使用：在編輯器選 setup 並執行。可重複執行，不會清掉既有列。 */
function setup() {
  var sh = sheet_();
  var first = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var empty = first.every(function (c) { return c === '' || c === null; });
  if (empty || String(first[0]) !== HEADERS[0]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 健康檢查：瀏覽器開啟網頁應用程式網址應見到 ok */
function doGet() {
  return out_({ ok: true, msg: '岳陽樓記導讀・交卷服務運作中' });
}

/**
 * 接收學生交卷 JSON。
 * 前端以 text/plain 送出，避免 CORS preflight；此處仍以 JSON 解析。
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var raw = '';
    if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    } else if (e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }
    var d = {};
    try {
      d = JSON.parse(raw);
    } catch (err) {
      d = (e && e.parameter) ? e.parameter : {};
    }

    var cls = String(d.class || d.cls || '').trim().toUpperCase();
    var sid = String(d.studentId || d.studentNo || d.sid || '').trim();
    var name = String(d.name || d.studentName || '').trim();
    var sc = Number(d.scoreContent);
    var sv = Number(d.scoreVocab);
    var sr = Number(d.scoreRhetorical);
    var total = Number(d.scoreTotal);
    var maxT = Number(d.maxTotal) || 60;
    var pct = d.percent != null ? Number(d.percent) : (maxT ? Math.round(total / maxT * 1000) / 10 : 0);
    var note = String(d.answersSummary || d.note || '').trim();
    var ts = d.timestamp ? new Date(d.timestamp) : new Date();
    if (isNaN(ts.getTime())) ts = new Date();

    if (!cls || !sid || !name) {
      return out_({ ok: false, error: 'missing student fields' });
    }
    if (!(sc >= 0 && sc <= 20) || !(sv >= 0 && sv <= 20) || !(sr >= 0 && sr <= 20)) {
      return out_({ ok: false, error: 'invalid section scores' });
    }
    if (isNaN(total)) total = sc + sv + sr;

    setup();
    var sh = sheet_();
    sh.appendRow([
      ts,
      cls,
      "'" + sid,
      name,
      sc,
      sv,
      sr,
      total,
      maxT,
      pct,
      note
    ]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
