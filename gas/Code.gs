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
