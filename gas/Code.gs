// マインドフルネス継続アプリ：記録をスプレッドシートに保存するGAS Web App
//
// 列の並び（1行目に見出しを作っておくこと）：
// timestamp | type | score | session_id | duration_min | label_kangae | label_fuan | label_keikaku | memo

var SHEET_NAME = "記録";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var secret = PropertiesService.getScriptProperties().getProperty("SECRET");
    if (!secret || data.secret !== secret) {
      return jsonResponse({ ok: false, error: "合言葉が違います" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ ok: false, error: "シート「" + SHEET_NAME + "」が見つかりません" });
    }

    if (data.quick) {
      appendRow(sheet, {
        timestamp: data.quick.timestamp,
        type: "quick",
        score: data.quick.score
      });
    }

    if (data.pre) {
      appendRow(sheet, {
        timestamp: data.pre.timestamp,
        type: "pre",
        score: data.pre.score,
        session_id: data.session_id,
        duration_min: data.duration_min
      });
    }

    if (data.post) {
      appendRow(sheet, {
        timestamp: data.post.timestamp,
        type: "post",
        score: data.post.score,
        session_id: data.session_id,
        duration_min: data.duration_min,
        label_kangae: data.label_kangae,
        label_fuan: data.label_fuan,
        label_keikaku: data.label_keikaku
      });
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// ホーム画面の「今日の記録件数」用
function doGet(e) {
  try {
    var secret = PropertiesService.getScriptProperties().getProperty("SECRET");
    if (!secret || e.parameter.secret !== secret) {
      return jsonResponse({ ok: false, error: "合言葉が違います" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ ok: false, error: "シート「" + SHEET_NAME + "」が見つかりません" });
    }

    return jsonResponse({ ok: true, todayCount: countTodayRows(sheet) });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function countTodayRows(sheet) {
  var values = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var count = 0;
  for (var i = 1; i < values.length; i++) {
    var raw = values[i][0];
    if (!raw) continue;
    var d = raw instanceof Date ? raw : new Date(raw);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, tz, "yyyy-MM-dd") === todayStr) count++;
  }
  return count;
}

function appendRow(sheet, row) {
  sheet.appendRow([
    row.timestamp || "",
    row.type || "",
    row.score === undefined || row.score === null ? "" : row.score,
    row.session_id || "",
    row.duration_min === undefined ? "" : row.duration_min,
    row.label_kangae === undefined ? "" : row.label_kangae,
    row.label_fuan === undefined ? "" : row.label_fuan,
    row.label_keikaku === undefined ? "" : row.label_keikaku,
    ""
  ]);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 毎晩21時のトリガーから呼ぶ関数（Apps Scriptの「トリガー」画面で手動設定する）
function checkDailyRecord() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  if (countTodayRows(sheet) === 0) {
    notifyNoRecordToday();
  }
}

// 通知を送る部分だけを切り出した関数。将来LINE等に差し替える場合はここだけ直せばよい
function notifyNoRecordToday() {
  sendReminderEmail(
    "【マインドフルネス】今日の記録がまだありません",
    "今日はまだ気分の記録がありません。\n1回だけでも「気分を記録」してみましょう。"
  );
}

// 毎時のトリガーから呼ぶ関数（8〜21時の間だけ、ランダムな確率で気分記録を促すメールを送る）
// 1時間ごとに呼ばれる前提。13時間(8〜20時台)×MOOD_PROMPT_PROBABILITYが1日あたりの平均送信回数
// 回数を増減したいときはこの数値だけ変えればよい（0.2なら平均2〜3回/日）
var MOOD_PROMPT_PROBABILITY = 0.2;

function maybeSendMoodPrompt() {
  var hour = new Date().getHours();
  if (hour < 8 || hour >= 21) return;
  if (Math.random() > MOOD_PROMPT_PROBABILITY) return;

  sendReminderEmail(
    "【マインドフルネス】今の気分はいかがですか？",
    "少し手を止めて、今の気分を記録してみましょう。"
  );
}

// メール送信そのものを切り出した関数。将来LINE等に差し替える場合はここだけ直せばよい
function sendReminderEmail(subject, body) {
  var to = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(to, subject, body);
}
