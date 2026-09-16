/**
 * mindful-app 用 GAS（Google Apps Script）
 * 役割：① doPost で記録をスプレッドシートに追記 ② 毎晩21時に「今日の記録が0件」ならメール
 *
 * 初回セットアップは docs/setup.md を参照。
 * 合言葉は「プロジェクトの設定 → スクリプト プロパティ」に SECRET という名前で入れる（コードに書かない）。
 */

const SHEET_NAME = 'log';
const HEADER = ['timestamp', 'type', 'score', 'session_id', 'duration_min',
                'label_kangae', 'label_fuan', 'label_keikaku', 'memo'];

// ---- ① 記録の受け取り ----
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }

  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret || body.secret !== secret) {
    return json_({ ok: false, error: 'unauthorized' });
  }

  const records = Array.isArray(body.records) ? body.records : [];
  if (records.length === 0) return json_({ ok: false, error: 'no_records' });

  const sheet = getSheet_();
  records.forEach(r => {
    sheet.appendRow(HEADER.map(k => (r[k] === undefined || r[k] === null) ? '' : r[k]));
  });
  return json_({ ok: true, count: records.length });
}

// ---- ② 毎晩の未記録チェック（時間トリガーから呼ばれる） ----
function checkTodayAndNotify() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const values = getSheet_().getDataRange().getValues().slice(1);   // ヘッダーを除く
  const count = values.filter(row => String(row[0]).slice(0, 10) === today).length;
  if (count === 0) {
    notify_('【mindful】今日の記録が0件です', '気分だけでも1タップで記録しておくと、比較データになります。');
  }
}

// 通知手段はこの関数だけ差し替えれば LINE 等に変えられる（仕様書 §7）
function notify_(subject, message) {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, message);
}

// ---- 初回に1回だけ手動実行：21時のトリガーを作る ----
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkTodayAndNotify')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkTodayAndNotify').timeBased().atHour(21).everyDays(1).create();
}

// ---- 内部ユーティリティ ----
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
