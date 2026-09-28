// 走行メーター帳 → スプレッドシート受け取り用（バックアップ・月別明細）
// 使い方：スプレッドシートの「拡張機能」→「Apps Script」にこのコードを貼り付け、
// 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」（実行：自分／アクセス：全員）で公開する。
// 最初に届いた合言葉だけを受け付けます。合言葉は「バックアップ」シートのB1に書き出されます。

var BACKUP_SHEET = 'バックアップ';

function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  if (d.app !== 'meterbook') return out_({ ok: false, error: 'app' });
  if (!checkKey_(d.key)) return out_({ ok: false, error: 'key' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (d.month && d.rows && d.rows.length) writeMonthSheet_(ss, d);
    if (d.month && d.entries) writeBackup_(ss, 'm:' + d.month, d.entries);
    if (d.settings) writeBackup_(ss, 'settings', d.settings);
  } finally {
    lock.releaseLock();
  }
  return out_({ ok: true });
}

// 復元用：?key=合言葉 で全データを返す
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!checkKey_(p.key, true)) return out_({ ok: false, error: 'key' });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKUP_SHEET);
  var months = {}, settings = null;
  if (sh && sh.getLastRow() >= 3) {
    sh.getRange(3, 1, sh.getLastRow() - 2, 2).getValues().forEach(function (r) {
      var k = String(r[0]);
      if (k === 'settings') settings = JSON.parse(r[1]);
      else if (k.indexOf('m:') === 0) months[k.slice(2)] = JSON.parse(r[1]);
    });
  }
  return out_({ ok: true, app: 'meterbook', version: 1, settings: settings, months: months });
}

function checkKey_(key, readOnly) {
  if (!key) return false;
  var props = PropertiesService.getScriptProperties();
  var saved = props.getProperty('KEY');
  if (!saved) {
    if (readOnly) return false;
    props.setProperty('KEY', key);
    var sh = backupSheet_(SpreadsheetApp.getActiveSpreadsheet());
    sh.getRange(1, 1, 1, 3).setValues([['合言葉', key, '機種変更の時にアプリの設定へ入れてください']]);
    return true;
  }
  return saved === key;
}

function backupSheet_(ss) {
  var sh = ss.getSheetByName(BACKUP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BACKUP_SHEET);
    sh.getRange(2, 1, 1, 3).setValues([['項目', 'データ（アプリ用・編集しないでください）', '更新日時']]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');
    sh.setColumnWidth(2, 400);
  }
  return sh;
}

function writeBackup_(ss, key, value) {
  var sh = backupSheet_(ss);
  var json = JSON.stringify(value);
  var last = sh.getLastRow();
  var row = 0;
  if (last >= 3) {
    var keys = sh.getRange(3, 1, last - 2, 1).getValues();
    for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === key) { row = i + 3; break; }
  }
  if (!row) row = Math.max(last, 2) + 1;
  sh.getRange(row, 1, 1, 3).setValues([[key, json, new Date()]]);
}

function writeMonthSheet_(ss, d) {
  var title = (d.month + '_' + String(d.name || '未設定').replace(/[\\\/?*\[\]:]/g, '')).slice(0, 90);
  var sh = ss.getSheetByName(title) || ss.insertSheet(title);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['氏名', d.name || '', '更新', new Date()]]);
  sh.getRange(3, 1, d.rows.length, d.rows[0].length).setValues(d.rows);
  sh.getRange(3, 1, 1, d.rows[0].length).setFontWeight('bold').setBackground('#eeeeee');
  sh.getRange(2 + d.rows.length, 1, 1, d.rows[0].length).setFontWeight('bold');
  if (d.rows.length > 1) sh.getRange(4, 9, d.rows.length - 1, 3).setNumberFormat('¥#,##0');
  sh.setFrozenRows(3);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
