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
 *   1分ごとに確認して一覧に反映し（置いてから1分ほどで使える）、CSV は「取り込み済み」フォルダへ移す。
 *   同じ回（開催日＋開始時刻＋種別＋講師名）が既にあれば、新しい内容で上書きする（中止の反映など）。
 *
 * 入れ替え（撮り直し用）:
 *   ファイル名を replace__<年-月>__<種別>__<都道府県>.csv にすると、その範囲の行を一覧から
 *   いったん全部消してから CSV の内容を入れる（講師変更・削除された回・表記ゆれの残りを防ぐ）。
 *   種別・都道府県は「+」区切り、絞り込み無しは ALL。
 *   例）replace__2027-01__BE__ALL.csv / replace__2026-11__PBS__京都府+大阪府+兵庫県.csv
 *   CSV が見出しだけ（0件）でも、その範囲は空になる。
 *
 * ChatGPT連携:
 *   ウェブアプリとして公開すると、ChatGPT（カスタムGPTのアクション）から
 *   「PBSだけ」「大阪だけ」「12月分」などの条件で一覧を取得し、PDFも作れる。
 *   PDFは「PDF出力」フォルダに保存され、共有設定は変えない（本人だけが開ける）。
 *   呼び出しには合言葉（token）が必要。合言葉は「ルール」シートに表示される。
 *
 * 初回だけ: このファイルの中身を「拡張機能 → Apps Script」に貼り付け → setup を実行 → 権限を許可。
 * ChatGPT連携: デプロイ → 新しいデプロイ → ウェブアプリ（実行ユーザー：自分／アクセス：全員）
 *   → 表示されたURLを「ChatGPT設定」シートのB1に貼る → chatgptSettings を実行。
 */

const SHEET_NAME = '一覧';
const RULE_SHEET_NAME = 'ルール';
const HEADERS = ['開催日', '時間', '種別', '講師名', '都道府県', '会場', '主催者'];
const KEEP_MONTHS = 2;          // 終わった月を含めて残す過去の月数（今月・来月以降は常に残す）
const INBOX_FOLDER = '取り込み待ち';
const DONE_FOLDER = '取り込み済み';
const TZ = 'Asia/Tokyo';
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];
const PDF_FOLDER = 'PDF出力';
const GPT_SHEET_NAME = 'ChatGPT設定';
const MAX_ROWS = 300;           // 1回の応答で返す最大件数

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  if (!props.getProperty('API_TOKEN')) props.setProperty('API_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  const sh = getListSheet_(ss);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  writeRules_(ss);
  getSubFolder_(ss, INBOX_FOLDER);
  getSubFolder_(ss, DONE_FOLDER);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'daily')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('daily').timeBased().everyMinutes(1).create();

  daily();
}

// 1分ごとに動く。CSVが無い時は取り込み待ちを見るだけで終わり、
// 古い月の削除と並べ替えは「取り込みがあった時」と「1日1回」だけ行う（実行時間の上限対策）。
function daily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  const imported = importInbox_(ss);
  if (!imported && props.getProperty('LAST_MAINT') === today) return;
  pruneOld_(ss);
  sortAndFormat_(ss);
  props.setProperty('LAST_MAINT', today);
}

function importInbox_(ss) {
  const sh = getListSheet_(ss);
  const inbox = getSubFolder_(ss, INBOX_FOLDER);
  const done = getSubFolder_(ss, DONE_FOLDER);
  const files = inbox.getFiles();
  if (!files.hasNext()) return false;

  const list = [];
  while (files.hasNext()) {
    const file = files.next();
    if (/\.csv$/i.test(file.getName()) || file.getMimeType() === MimeType.CSV) list.push(file);
  }
  // 撮り直しの新しい方が後から効くよう、置いた順に処理する
  list.sort((a, b) => a.getDateCreated() - b.getDateCreated());

  let rows = readRows_(sh);
  list.forEach(file => {
    const scope = replaceScope_(file.getName());
    if (scope) rows = rows.filter(r => !inScope_(r, scope));
    const index = new Map(rows.map((r, i) => [rowKey_(r), i]));
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
  });
  writeRows_(sh, rows);
  return list.length > 0;
}

// replace__2027-01__BE__ALL.csv → { month: '2027/01', types: ['BE'], prefs: null }
function replaceScope_(name) {
  const m = String(name).match(/^replace__(\d{4})-(\d{1,2})__([^_]+)__(.+)\.csv$/i);
  if (!m) return null;
  const types = m[3].split('+').map(t => t.trim().toUpperCase()).filter(Boolean);
  const prefs = m[4].toUpperCase() === 'ALL' ? null : m[4].split('+').map(x => x.trim()).filter(Boolean);
  return { month: m[1] + '/' + ('0' + m[2]).slice(-2), types: types, prefs: prefs };
}

function inScope_(r, scope) {
  const d = parseDate_(r[0]);
  if (!d || Utilities.formatDate(d, TZ, 'yyyy/MM') !== scope.month) return false;
  if (scope.types.indexOf(String(r[2]).toUpperCase()) < 0) return false;
  return !scope.prefs || scope.prefs.indexOf(String(r[4])) >= 0;
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
    ['自動削除', '1日1回（取り込み時も）実行。月が終わったら、その2ヶ月前の月の行を削除（例：12/1に9月以前を削除）'],
    ['取り込み', '「' + INBOX_FOLDER + '」フォルダに同じ列順のCSVを置くと1分ほどで反映し、「' + DONE_FOLDER + '」へ移動'],
    ['重複', '開催日＋開始時刻＋種別＋講師名が同じなら上書き'],
    ['入れ替え', 'ファイル名 replace__年-月__種別__都道府県.csv はその範囲を全部消してから入れる（例 replace__2027-01__BE__ALL.csv）'],
    ['ChatGPT連携の合言葉', PropertiesService.getScriptProperties().getProperty('API_TOKEN') || '（setup実行後に表示）'],
  ];
  sh.clearContents();
  sh.getRange(1, 1, rules.length, 2).setValues(rules);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}

// ===== ChatGPT連携（ウェブアプリ） =====
// 例）?token=xxx&type=PBS&pref=大阪&month=2026-12&format=pdf

function doGet(e) {
  const p = (e && e.parameter) || {};
  const props = PropertiesService.getScriptProperties();
  if (!p.token || p.token !== props.getProperty('API_TOKEN')) return json_({ error: 'unauthorized' });

  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const rows = filterRows_(readRows_(getListSheet_(ss)), p);
  const items = rows.slice(0, MAX_ROWS).map(toItem_);
  const res = { count: rows.length, returned: items.length, conditions: describe_(p), items: items };
  if (p.format === 'pdf') {
    const file = makePdf_(ss, rows, res.conditions);
    res.pdf_url = file.getUrl();
    res.pdf_name = file.getName();
  }
  return json_(res);
}

function filterRows_(rows, p) {
  const types = splitList_(p.type).map(t => t.toUpperCase());
  const prefs = splitList_(p.pref);
  const lecturer = (p.lecturer || '').replace(/\s/g, '');
  const venue = (p.venue || '').replace(/\s/g, '');
  const from = p.month ? monthStart_(p.month) : parseDate_(p.from || '');
  const to = p.month ? monthEnd_(p.month) : parseDate_(p.to || '');
  const includeCanceled = String(p.include_canceled || 'true') !== 'false';
  return rows.filter(r => {
    const d = parseDate_(r[0]);
    if (from && (!d || d < from)) return false;
    if (to && (!d || d > to)) return false;
    if (types.length && types.indexOf(String(r[2]).toUpperCase()) < 0) return false;
    // 「大阪」「大阪府」どちらでも当たるよう部分一致
    if (prefs.length && !prefs.some(x => String(r[4]).indexOf(x.replace(/[都道府県]$/, '')) === 0)) return false;
    if (lecturer && String(r[3]).replace(/\s/g, '').indexOf(lecturer) < 0) return false;
    if (venue && String(r[5]).replace(/\s/g, '').indexOf(venue) < 0) return false;
    if (!includeCanceled && String(r[5]).indexOf('（中止）') >= 0) return false;
    return true;
  });
}

function toItem_(r) {
  const t = String(r[1]).split('\n');
  return { date: r[0], start: t[0] || '', end: t[1] || '', type: r[2], lecturer: r[3],
           pref: r[4], venue: r[5], organizer: r[6], canceled: String(r[5]).indexOf('（中止）') >= 0 };
}

function makePdf_(ss, rows, title) {
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const body = rows.map(r => '<tr>' + [r[0], esc(r[1]).replace('\n', '<br>'), r[2], r[3], r[4], r[5], r[6]]
    .map((v, i) => '<td' + (i === 1 ? ' class="t"' : '') + '>' + (i === 1 ? v : esc(v)) + '</td>').join('') + '</tr>').join('');
  const html = '<html><head><meta charset="utf-8"><style>'
    + 'body{font-family:sans-serif;font-size:9pt}h1{font-size:13pt}'
    + 'table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:3px 4px;vertical-align:middle}'
    + 'th{background:#eee}td.t{text-align:center;white-space:nowrap}</style></head><body>'
    + '<h1>全厚済セミナー一覧（' + esc(title) + '）</h1>'
    + '<p>' + rows.length + '件　作成：' + Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') + '</p>'
    + '<table><tr>' + HEADERS.map(h => '<th>' + h + '</th>').join('') + '</tr>' + body + '</table></body></html>';
  const name = 'セミナー一覧_' + title.replace(/[\\\/:*?"<>|\s]/g, '_') + '_'
    + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '.pdf';
  const blob = Utilities.newBlob(html, 'text/html', 'list.html').getAs(MimeType.PDF).setName(name);
  return getSubFolder_(ss, PDF_FOLDER).createFile(blob);
}

function describe_(p) {
  const parts = [];
  if (p.month) parts.push(p.month.replace('-', '年') + '月');
  if (p.from || p.to) parts.push((p.from || '') + '〜' + (p.to || ''));
  if (p.type) parts.push(p.type);
  if (p.pref) parts.push(p.pref);
  if (p.lecturer) parts.push('講師:' + p.lecturer);
  if (p.venue) parts.push('会場:' + p.venue);
  if (String(p.include_canceled) === 'false') parts.push('中止除く');
  return parts.length ? parts.join('・') : '全件';
}

function splitList_(v) {
  return String(v || '').split(/[,、，\s]+/).filter(Boolean);
}

function monthStart_(ym) {
  const m = String(ym).match(/(\d{4})\D+(\d{1,2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, 1) : null;
}

function monthEnd_(ym) {
  const m = String(ym).match(/(\d{4})\D+(\d{1,2})/);
  return m ? new Date(Number(m[1]), Number(m[2]), 0) : null;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ウェブアプリのURLを「ChatGPT設定」シートのB1に貼ってから実行すると、
// カスタムGPTにそのまま貼れる「指示文」と「スキーマ」を作る。
function chatgptSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GPT_SHEET_NAME) || ss.insertSheet(GPT_SHEET_NAME);
  const url = String(sh.getRange('B1').getValue() || '').trim();
  sh.getRange('A1').setValue('ウェブアプリURL（ここに貼る→）');
  const m = url.match(/^https:\/\/script\.google\.com(\/macros\/s\/[^\/]+\/exec)/);
  if (!m) {
    sh.getRange('A3').setValue('B1にウェブアプリのURL（https://script.google.com/macros/s/…/exec）を貼ってから再実行してください。');
    return;
  }
  const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  const schema = {
    openapi: '3.1.0',
    info: { title: '全厚済セミナー一覧', version: '1.0.0' },
    servers: [{ url: 'https://script.google.com' }],
    paths: {},
  };
  schema.paths[m[1]] = { get: {
    operationId: 'searchSeminars',
    summary: '全厚済セミナー一覧を条件で検索し、必要ならPDFを作る',
    parameters: [
      { name: 'token', in: 'query', required: true, schema: { type: 'string', enum: [token] }, description: '合言葉（常にこの値）' },
      { name: 'type', in: 'query', schema: { type: 'string' }, description: '種別。PBS, BE, BT, コンベンション。複数はカンマ区切り' },
      { name: 'pref', in: 'query', schema: { type: 'string' }, description: '都道府県。例: 大阪, 兵庫。複数はカンマ区切り' },
      { name: 'month', in: 'query', schema: { type: 'string' }, description: '対象月 yyyy-mm。例: 2026-12' },
      { name: 'from', in: 'query', schema: { type: 'string' }, description: '開始日 yyyy/mm/dd（monthと併用しない）' },
      { name: 'to', in: 'query', schema: { type: 'string' }, description: '終了日 yyyy/mm/dd（monthと併用しない）' },
      { name: 'lecturer', in: 'query', schema: { type: 'string' }, description: '講師名（部分一致）' },
      { name: 'venue', in: 'query', schema: { type: 'string' }, description: '会場名（部分一致）' },
      { name: 'include_canceled', in: 'query', schema: { type: 'string', enum: ['true', 'false'] }, description: '中止の回を含めるか。既定 true' },
      { name: 'format', in: 'query', schema: { type: 'string', enum: ['json', 'pdf'] }, description: 'pdf を指定するとPDFを作成しURLを返す' },
    ],
    responses: { '200': { description: '検索結果' } },
  } };
  const instructions = [
    'あなたはJPの秘書です。全厚済セミナーの予定を searchSeminars アクションで調べて答えます。',
    '・「PBSだけ」「大阪だけ」「12月」「来月」などの条件を type / pref / month に変換して呼び出す。月は今日の日付から yyyy-mm に直す。',
    '・「1月」のように月だけ言われたら、今日以降で一番近いその月（例：今が2026年10月なら2027年1月）として扱う。',
    '・「PDFにして」「PDFで」と言われたら format=pdf で呼び出し、返ってきた pdf_url を伝える。',
    '・回答の並びは必ず「開催日｜時間（開始〜終了）｜種別｜講師名｜都道府県｜会場｜主催者」。',
    '・音声で読み上げるときは、件数を先に言い、そのあと日付順に「日付・時間・講師・会場」を簡潔に読む。10件を超えるときは最初の10件を読み、残りは画面のリストかPDFを案内する。',
    '・中止の回（canceled=true）は必ず「中止」と伝える。',
    '・会場名の「…※」は元画面で名前が切れている印。推測で補わない。',
    '・データに無いことは推測せず「一覧に無い」と答える。',
  ].join('\n');
  sh.getRange('A3:B5').setValues([
    ['使い方', 'ChatGPT → GPTを作成 → 構成：「指示」にB4、アクション →「スキーマ」にB5を貼る（認証：なし）'],
    ['指示（Instructions）', instructions],
    ['スキーマ（Schema）', JSON.stringify(schema, null, 2)],
  ]);
  sh.getRange('B3:B5').setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 700);
}
