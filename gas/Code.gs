// 青木くん、離席できません。 ランキング用 Google Apps Script
// スプレッドシートの「拡張機能 → Apps Script」に、このファイルの中身をまるごと貼り付けて使います。

const SHEET_NAME = 'scores';
const TZ = 'Asia/Tokyo';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['登録日時', '日付', 'id', '名前', 'スコア', '死因', '時刻', 'Lv', '最高体重', '稼いだ額', '称号']);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 日付は文字列のまま保存
    sh.getRange('G:G').setNumberFormat('@'); // 時刻も文字列のまま保存
  }
  return sh;
}

function today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

// 先頭が = + - @ の文字はスプレッドシートで数式扱いされるので無害化する
function text_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ランキング取得: ?scope=today（今日） / ?scope=all（歴代）、?sort=score / weight / earned
function doGet(e) {
  const scope = (e && e.parameter && e.parameter.scope) === 'all' ? 'all' : 'today';
  const day = today_();
  // 見出し行の有無にかかわらず、スコアが数値の行だけを記録として扱う
  const values = sheet_().getDataRange().getValues().filter(function (r) { return r[2] !== '' && typeof r[4] === 'number'; });
  let rows = values.map(function (r) {
    const d = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]);
    return { day: d, id: String(r[2]), name: String(r[3]).replace(/^'/, ''), score: Number(r[4]) || 0,
             cause: String(r[5]).replace(/^'/, ''), clock: r[6] instanceof Date ? Utilities.formatDate(r[6], TZ, 'HH:mm') : String(r[6]), level: Number(r[7]) || 1, maxW: Number(r[8]) || 0,
             earned: Number(r[9]) || 0, title: String(r[10] == null ? '' : r[10]).replace(/^'/, '') };
  });
  if (scope === 'today') rows = rows.filter(function (r) { return r.day === day; });
  const sort = e && e.parameter && e.parameter.sort;
  const key = sort === 'weight' ? 'maxW' : sort === 'earned' ? 'earned' : 'score';
  rows.sort(function (a, b) { return b[key] - a[key]; });
  return json_(rows.slice(0, 50));
}

// スコア登録（ゲームから自動で呼ばれます）
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const score = Math.max(0, Math.min(9999999, Math.round(Number(d.score) || 0)));
    sheet_().appendRow([
      new Date(),
      today_(),
      text_(d.id, 24),
      text_(d.name, 12) || '名無しの青木ファン',
      score,
      text_(d.cause, 40),
      text_(d.clock, 5),
      Math.max(1, Math.min(99, Math.round(Number(d.level) || 1))),
      Math.max(0, Math.min(999, Math.round(Number(d.maxW) || 0))),
      Math.max(0, Math.min(99999999, Math.round(Number(d.earned) || 0))),
      text_(d.title, 20)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}
