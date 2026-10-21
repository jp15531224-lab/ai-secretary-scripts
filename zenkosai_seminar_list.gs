/**
 * 全厚済セミナー一覧：取り込み＋自動削除（Google Apps Script / スプレッドシートに紐づけて使う）
 *
 * 表示ルール（固定）:
 *   列順 = 開催日｜時間｜種別｜講師名｜都道府県｜会場｜主催者
 *   時間 = 上段に開始、下段に終了（セル内改行）
 *   開催日 = yyyy/mm/dd(曜) で保存（年が無いと自動削除の判定ができないため）
 *   中止の回は会場の後ろに「（中止）」、画面で切れた会場名は「…※」のまま（推測で補わない）
 *
 * 自動削除:
 *   月が終わったら、その終わった月の2ヶ月前の月の行を削除する。
 *   例）11月が終わった 12/1 に 9月以前を削除。10月・11月・12月以降は残す。
 *   年が読めない行は削除しない。元の画像は一切触らない（このスクリプトは一覧シートの行だけを扱う）。
 *
 * 取り込み:
 *   一覧と同じフォルダの「取り込み待ち」フォルダに CSV（上と同じ列順、1行目は見出し）を置くと、
 *   毎日の処理で一覧に追記し、CSV は「取り込み済み」フォルダへ移す。
 *   同じ回（開催日＋開始時刻＋種別＋講師名）が既にあれば、新しい内容で上書きする（中止の反映など）。
 *
 * 初回だけ: このファイルの中身を「拡張機能 → Apps Script」に貼り付け → setup を実行 → 権限を許可。
 */

const SHEET_NAME = '一覧';
const RULE_SHEET_NAME = 'ルール';
const HEADERS = ['開催日', '時間', '種別', '講師名', '都道府県', '会場', '主催者'];
const KEEP_MONTHS = 2;          // 終わった月を含めて残す過去の月数（今月・来月以降は常に残す）
const INBOX_FOLDER = '取り込み待ち';
const DONE_FOLDER = '取り込み済み';
const TZ = 'Asia/Tokyo';
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  const sh = getListSheet_(ss);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  writeRules_(ss);
  getSubFolder_(ss, INBOX_FOLDER);
  getSubFolder_(ss, DONE_FOLDER);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'daily')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('daily').timeBased().everyDays(1).atHour(5).inTimezone(TZ).create();

  daily();
}

function daily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  importInbox_(ss);
  pruneOld_(ss);
  sortAndFormat_(ss);
}

function importInbox_(ss) {
  const sh = getListSheet_(ss);
  const inbox = getSubFolder_(ss, INBOX_FOLDER);
  const done = getSubFolder_(ss, DONE_FOLDER);
  const files = inbox.getFiles();
  if (!files.hasNext()) return;

  const rows = readRows_(sh);
  const index = new Map(rows.map((r, i) => [rowKey_(r), i]));

  while (files.hasNext()) {
    const file = files.next();
    if (!/\.csv$/i.test(file.getName()) && file.getMimeType() !== MimeType.CSV) continue;
    const csv = Utilities.parseCsv(file.getBlob().getDataAsString('UTF-8'));
    csv.forEach(line => {
      if (line.length < HEADERS.length || line[0] === HEADERS[0]) return;
      const r = normalizeRow_(line.slice(0, HEADERS.length));
      const key = rowKey_(r);
      if (index.has(key)) {
        rows[index.get(key)] = r;
      } else {
        index.set(key, rows.length);
        rows.push(r);
      }
    });
    file.moveTo(done);
  }
  writeRows_(sh, rows);
}

function pruneOld_(ss) {
  const sh = getListSheet_(ss);
  const now = new Date();
  const y = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  const m = Number(Utilities.formatDate(now, TZ, 'M'));
  // 例）12月なら cutoff = 10/1 → 9月以前を削除
  const cutoff = new Date(y, m - 1 - KEEP_MONTHS, 1);
  const kept = readRows_(sh).filter(r => {
    const d = parseDate_(r[0]);
    return d === null || d >= cutoff;
  });
  writeRows_(sh, kept);
}

function sortAndFormat_(ss) {
  const sh = getListSheet_(ss);
  const rows = readRows_(sh);
  rows.sort((a, b) => sortKey_(a).localeCompare(sortKey_(b)));
  writeRows_(sh, rows);
  const n = Math.max(rows.length, 1);
  sh.getRange(1, 1, n + 1, HEADERS.length).setVerticalAlignment('middle');
  sh.getRange(2, 2, n, 1).setWrap(true).setHorizontalAlignment('center');
  sh.autoResizeColumns(1, HEADERS.length);
}

function normalizeRow_(r) {
  const out = r.map(v => String(v == null ? '' : v).trim());
  const d = parseDate_(out[0]);
  if (d) out[0] = formatDate_(d);
  // 「11:30-16:00」「11:30～16:00」なども 上段開始／下段終了 にそろえる
  const t = out[1].match(/(\d{1,2}:\d{2})\D+(\d{1,2}:\d{2})/);
  if (t) out[1] = t[1] + '\n' + t[2];
  return out;
}

function parseDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v).match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function formatDate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd') + '(' + WEEK[d.getDay()] + ')';
}

function rowKey_(r) {
  const d = parseDate_(r[0]);
  const start = (String(r[1]).match(/\d{1,2}:\d{2}/) || [''])[0].padStart(5, '0');
  // 同じ講師・同じ時刻で種別だけ違う回がある（例：中止のBTと開催のPBS）ため種別もキーに含める
  return [d ? formatDate_(d) : r[0], start, r[2], String(r[3]).replace(/\s/g, '')].join('|');
}

function sortKey_(r) {
  const d = parseDate_(r[0]);
  const start = (String(r[1]).match(/\d{1,2}:\d{2}/) || [''])[0].padStart(5, '0');
  return (d ? Utilities.formatDate(d, TZ, 'yyyyMMdd') : '99999999') + start;
}

function readRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => r.some(v => v !== ''))
    .map(r => normalizeRow_(r));
}

function writeRows_(sh, rows) {
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS.length).setNumberFormat('@').setValues(rows);
}

function getListSheet_(ss) {
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0].setName(SHEET_NAME);
}

function getSubFolder_(ss, name) {
  const parent = DriveApp.getFileById(ss.getId()).getParents().next();
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function writeRules_(ss) {
  const sh = ss.getSheetByName(RULE_SHEET_NAME) || ss.insertSheet(RULE_SHEET_NAME);
  const rules = [
    ['項目', '内容'],
    ['列順', HEADERS.join('｜')],
    ['時間', '上段に開始、下段に終了（セル内改行）'],
    ['開催日', 'yyyy/mm/dd(曜)。年が無い行は自動削除の対象外'],
    ['会場', '中止は「（中止）」を付記。画面で切れた名称は「…※」のまま（推測で補わない）'],
    ['自動削除', '毎日5時に実行。月が終わったら、その2ヶ月前の月の行を削除（例：12/1に9月以前を削除）'],
    ['取り込み', '「' + INBOX_FOLDER + '」フォルダに同じ列順のCSVを置くと毎日追記し、「' + DONE_FOLDER + '」へ移動'],
    ['重複', '開催日＋開始時刻＋種別＋講師名が同じなら上書き'],
  ];
  sh.clearContents();
  sh.getRange(1, 1, rules.length, 2).setValues(rules);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}
