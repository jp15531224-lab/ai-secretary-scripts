// ===== Apps Script用 1ファイル版（コード.gs ＋ 参加者管理.gs）=====
// プロジェクトの .gs ファイルをこの1つだけにして使う（他の .gs があると二重定義のエラーになる）。

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
  if (e.pathInfo || (p.action && p.action !== 'search')) return handleAttendance_(e);  // 参加者管理（参加者管理.gs）

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
  return { seminar_id: seminarId_(r), date: r[0], start: t[0] || '', end: t[1] || '', type: r[2], lecturer: r[3],
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
    info: { title: '全厚済セミナー一覧', version: '1.1.0' },
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
  addAttendanceParams_(schema.paths[m[1]].get);  // 参加者管理（参加者管理.gs）
  const instructions = [
    'あなたはJPの秘書です。全厚済セミナーの予定を seminarApi アクションで調べて答えます。',
    '・「PBSだけ」「大阪だけ」「12月」「来月」などの条件を type / pref / month に変換して呼び出す。月は今日の日付から yyyy-mm に直す。',
    '・「1月」のように月だけ言われたら、今日以降で一番近いその月（例：今が2026年10月なら2027年1月）として扱う。',
    '・「PDFにして」「PDFで」と言われたら format=pdf で呼び出し、返ってきた pdf_url を伝える。',
    '・回答の並びは必ず「開催日｜時間（開始〜終了）｜種別｜講師名｜都道府県｜会場｜主催者」。',
    '・音声で読み上げるときは、件数を先に言い、そのあと日付順に「日付・時間・講師・会場」を簡潔に読む。10件を超えるときは最初の10件を読み、残りは画面のリストかPDFを案内する。',
    '・中止の回（canceled=true）は必ず「中止」と伝える。',
    '・会場名の「…※」は元画面で名前が切れている印。推測で補わない。',
    '・データに無いことは推測せず「一覧に無い」と答える。',
  ].join('\n') + attendanceInstructions_();
  sh.getRange('A3:B5').setValues([
    ['使い方', 'ChatGPT → GPTを作成 → 構成：「指示」にB4、アクション →「スキーマ」にB5を貼る（認証：なし）'],
    ['指示（Instructions）', instructions],
    ['スキーマ（Schema）', JSON.stringify(schema, null, 2)],
  ]);
  sh.getRange('B3:B5').setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 700);
}

/**
 * 全厚済セミナー：参加者管理（「コード.gs」と同じプロジェクトに追加するファイル）
 *
 * 保存先はこのスプレッドシートだけ。追加するシート:
 *   人物 / 参加記録 / 確認待ち候補
 * 「一覧」シートは読むだけ（取り込み・自動削除・並べ替えには一切触らない）。
 *
 * 紐付け:
 *   セミナーID = 開催日-開始時刻-種別-講師名（例 20261101-1200-PBS-青木弘幸）。
 *   行番号を使わないので、一覧の並べ替え・取り込み・削除で紐付けはずれない。
 *   参加記録には登録時点のセミナー情報（開催日・時間・種別・講師・地域・会場）も写して保存し、
 *   元の一覧から消えても過去の参加履歴は残る。
 *
 * 参加者の増減（2026-10-23 追加）:
 *   action=apply       … join=追加/再参加する人、leave=不参加にする人、undecided=判断できない人（確認待ちへ）
 *   action=roster_sync … names=名簿の全員。complete=true の時だけ、名簿に無い参加予定者を不参加にする
 *   action=changes     … 変更履歴（変更前後・日時・理由・出典）を見る
 *   action=person_rename … 人物の氏名の訂正（参加記録の氏名も直し、履歴に残す）
 *   不参加は「キャンセル」（参加予定だった人）か「辞退」（まだ返事前だった人）で保存し、行は消さない。
 *   書き込みはすべて「書く→読み直して一致を確認→合わなければ最大3回書き直す」。行はIDで探して書くので再試行で重複しない。
 *
 * ChatGPT（カスタムGPTのアクション）からは、ウェブアプリURLに action=register などを付けて呼ぶ。
 *   （URLの後ろに /register を付ける形はログインしていない呼び出しで使えないため、action で振り分ける）
 *   合言葉（token）の確認は「コード.gs」の doGet で済ませてからここに来る。
 */

const PERSON_SHEET = '人物';
const RECORD_SHEET = '参加記録';
const CANDIDATE_SHEET = '確認待ち候補';
const PERSON_HEADERS = ['人物ID', '氏名', '読み', '区別メモ', '登録日時'];
const RECORD_HEADERS = ['記録ID', '人物ID', '氏名', 'セミナーID', '開催日', '時間', '種別', '講師', '地域', '会場',
  '案内日', '連絡手段', '返事', '状態', '次回連絡日', 'フォロー状況', 'メモ', '登録日時', '更新日時'];
const CANDIDATE_HEADERS = ['候補ID', '入口', '資料', '重複キー', '根拠', '氏名', '人物ID', 'セミナーID', 'セミナー表示',
  '提案状態', '次回連絡日', 'メモ', '確認状態', '記録ID', '作成日時'];
const HISTORY_SHEET = '変更履歴';
const HISTORY_HEADERS = ['履歴ID', '日時', '操作', '記録ID', '人物ID', '氏名', 'セミナーID', '開催日', '変更前', '変更後', '理由', '出典'];
const MEMO_DONE_SHEET = '処理済み資料';
const MEMO_DONE_HEADERS = ['資料', '処理日時', '候補数'];
const STATUSES = ['案内予定', '案内済み', '返事待ち', '参加予定', '未定', '辞退', 'キャンセル', '参加済み', 'フォロー済み'];
const CLOSED_STATUSES = ['辞退', 'キャンセル', 'フォロー済み'];
const ATTENDING = ['参加予定', '参加済み', 'フォロー済み'];
const NOT_ATTENDING = ['辞退', 'キャンセル'];
const LAST_SEMINAR_HOURS = 24;  // セミナー指定なしの追加は、この時間内に扱ったセミナーとみなす
const STATUS_ORDER = ['参加済み', 'フォロー済み', '参加予定', '返事待ち', '案内済み', '未定', '案内予定', '辞退', 'キャンセル'];

function handleAttendance_(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || e.pathInfo || '').replace(/^\/+|\/+$/g, '');
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const handlers = {
    register: registerAttendance_, update: updateAttendance_, roster: seminarRoster_,
    history: personHistory_, followups: followups_, people: searchPeople_,
    candidate_add: addCandidate_, candidates: listCandidates_, candidate_confirm: confirmCandidates_,
    memos: listMemos_, memo_done: markMemoDone_,
    apply: applyChanges_, roster_sync: rosterSync_, changes: listChanges_, person_rename: renamePerson_,
  };
  if (!handlers[action]) return json_({ error: 'unknown_action', action: action });
  const write = ['register', 'update', 'candidate_add', 'candidate_confirm', 'memo_done',
    'apply', 'roster_sync', 'person_rename'].indexOf(action) >= 0;
  const lock = LockService.getScriptLock();
  syncReset_();
  if (write && !lock.tryLock(30000)) {
    return json_({ error: 'busy', message: '台帳を別の処理が更新中です。少し待ってもう一度呼んでください（まだ何も書き込んでいません）。' });
  }
  try {
    const res = handlers[action](ss, p);
    if (write) res.sync = syncOut_();
    return json_(res);
  } catch (err) {
    const out = { error: String(err && err.message || err) };
    if (write) out.sync = syncOut_();
    return json_(out);
  } finally {
    if (write) lock.releaseLock();
  }
}

// ===== セミナー =====

function seminarId_(r) {
  const d = parseDate_(r[0]);
  const start = (String(r[1]).match(/\d{1,2}:\d{2}/) || [''])[0].padStart(5, '0').replace(':', '');
  return [d ? Utilities.formatDate(d, TZ, 'yyyyMMdd') : r[0], start, r[2], String(r[3]).replace(/\s/g, '')].join('-');
}

function seminarFromRow_(r) {
  const t = String(r[1]).split('\n');
  return { seminar_id: seminarId_(r), date: r[0], start: t[0] || '', end: t[1] || '', type: r[2], lecturer: r[3],
           pref: r[4], venue: r[5], organizer: r[6], canceled: String(r[5]).indexOf('（中止）') >= 0, in_list: true };
}

function seminarFromRecord_(rec) {
  const t = String(rec['時間']).split('\n');
  return { seminar_id: rec['セミナーID'], date: rec['開催日'], start: t[0] || '', end: t[1] || '', type: rec['種別'],
           lecturer: rec['講師'], pref: rec['地域'], venue: rec['会場'], organizer: '', canceled: false, in_list: false };
}

function seminarLabel_(s) {
  return s.date + ' ' + s.start + (s.end ? '〜' + s.end : '') + '｜' + s.lecturer + '｜' + s.type + '｜' + s.pref + '｜' + s.venue
    + (s.canceled ? '｜中止' : '') + (s.in_list ? '' : '｜元一覧に無し');
}

function listSeminars_(ss) {
  return readRows_(getListSheet_(ss)).map(seminarFromRow_);
}

// セミナーIDか、日付・講師・種別・地域・会場・開始時刻の組み合わせで1件に絞る。
function resolveSeminar_(ss, p) {
  const r = resolveSeminarRaw_(ss, p);
  if (r.seminar) rememberSeminar_(r.seminar.seminar_id);
  return r;
}

function resolveSeminarRaw_(ss, p) {
  const all = listSeminars_(ss);
  if (p.seminar_id) {
    const hit = all.filter(s => s.seminar_id === p.seminar_id)[0];
    if (hit) return { seminar: hit };
    const rec = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS).filter(r => r['セミナーID'] === p.seminar_id)[0];
    if (rec) return { seminar: currentSeminar_(all, seminarFromRecord_(rec)) };
    return { error: 'seminar_not_found', message: 'このセミナーIDは一覧にも参加記録にもありません。' };
  }
  const date = parseDate_(p.date || '');
  const lecturer = norm_(p.lecturer);
  if (!date && !lecturer) return { error: 'seminar_criteria_missing', message: '開催日か講師名を指定してください。' };
  const types = splitList_(p.type).map(t => t.toUpperCase());
  const pref = String(p.pref || '').replace(/[都道府県]$/, '');
  const venue = norm_(p.venue);
  const start = String(p.start || '').match(/\d{1,2}:\d{2}/);
  const hits = all.filter(s => {
    const d = parseDate_(s.date);
    if (date && (!d || d.getTime() !== date.getTime())) return false;
    if (lecturer && norm_(s.lecturer).indexOf(lecturer) < 0) return false;
    if (types.length && types.indexOf(String(s.type).toUpperCase()) < 0) return false;
    if (pref && String(s.pref).indexOf(pref) !== 0) return false;
    if (venue && norm_(s.venue).indexOf(venue) < 0) return false;
    if (start && s.start.padStart(5, '0') !== start[0].padStart(5, '0')) return false;
    return true;
  });
  if (hits.length === 1) return { seminar: hits[0] };
  if (!hits.length) return { error: 'seminar_not_found', message: '条件に合うセミナーが一覧にありません。日付・講師・地域を確認してください。' };
  return { needs_confirmation: 'seminar', message: '候補が' + hits.length + '件あります。時間か会場で選んでください。',
           choices: hits.slice(0, 20).map(s => ({ seminar_id: s.seminar_id, label: seminarLabel_(s) })) };
}

// 参加記録に写したセミナーを、今の一覧の内容（会場変更・中止）で補う。
// 講師変更でIDが変わった時は、同じ日・開始時刻・種別・地域が1件だけなら同じ回とみなす。
function currentSeminar_(all, snap) {
  const exact = all.filter(s => s.seminar_id === snap.seminar_id)[0];
  if (exact) return exact;
  const same = all.filter(s => s.date === snap.date && s.start === snap.start && s.type === snap.type && s.pref === snap.pref);
  if (same.length === 1) return Object.assign({}, same[0], { seminar_id: snap.seminar_id, note: '講師などが元一覧で変更されています' });
  return snap;
}

// ===== 人物 =====

function norm_(v) {
  return String(v || '').replace(/[\s　]/g, '').replace(/(さん|様|さま|君|くん|ちゃん|先生)$/, '');
}

function findPeople_(ss, p) {
  const people = readObjects_(ss, PERSON_SHEET, PERSON_HEADERS);
  if (p.person_id) return people.filter(x => x['人物ID'] === p.person_id);
  const n = norm_(p.name);
  if (!n) return [];
  return people.filter(x => norm_(x['氏名']) === n || (x['読み'] && norm_(x['読み']) === n));
}

function personLabel_(x) {
  return x['氏名'] + (x['読み'] ? '（' + x['読み'] + '）' : '') + (x['区別メモ'] ? '・' + x['区別メモ'] : '');
}

// 1人に決まれば person、同名が複数なら確認、いなければ（create=true の時だけ）新規登録。
function resolvePerson_(ss, p, create) {
  const hits = findPeople_(ss, p);
  if (p.new_person !== 'true') {
    if (hits.length === 1) return { person: hits[0] };
    if (hits.length > 1) return { needs_confirmation: 'person', message: '同じ名前の人が' + hits.length + '人います。どちらですか。',
      choices: hits.map(x => ({ person_id: x['人物ID'], label: personLabel_(x) })) };
    if (p.person_id) return { error: 'person_not_found' };
  }
  if (!create) return { error: 'person_not_found', message: '「' + (p.name || '') + '」は人物に登録されていません。' };
  const name = String(p.name || '').trim().replace(/(さん|様|さま|君|くん|ちゃん)$/, '');
  if (!validName_(name)) return { error: 'name_unclear', message: '名前「' + name + '」が確定していません。正しい名前を確認してください。' };
  const person = { '人物ID': newId_('P'), '氏名': name, '読み': p.kana || '', '区別メモ': p.person_note || '', '登録日時': now_() };
  appendObject_(ss, PERSON_SHEET, PERSON_HEADERS, person);
  return { person: person, new_person: true };
}

function searchPeople_(ss, p) {
  const n = norm_(p.name);
  const people = readObjects_(ss, PERSON_SHEET, PERSON_HEADERS)
    .filter(x => !n || norm_(x['氏名']).indexOf(n) >= 0 || norm_(x['読み']).indexOf(n) >= 0);
  return { count: people.length, people: people.slice(0, 50).map(x => ({ person_id: x['人物ID'], label: personLabel_(x) })) };
}

// ===== 参加記録 =====

function registerAttendance_(ss, p) {
  const status = p.status || '案内済み';
  if (STATUSES.indexOf(status) < 0) return { error: 'bad_status', allowed: STATUSES };
  const sem = resolveSeminar_(ss, p);
  if (!sem.seminar) return sem;
  const per = resolvePerson_(ss, p, true);
  if (!per.person) return per;
  const records = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS);
  const dup = records.filter(r => r['人物ID'] === per.person['人物ID'] && r['セミナーID'] === sem.seminar.seminar_id)[0];
  if (dup && p.status && p.status !== dup['状態']) {
    const before = dup['状態'];
    setStatus_(ss, dup, p.status, { op: '状態変更', reason: p.reason || p.reply || '', source: p.source_ref || '' });
    return { ok: true, updated: true, status_before: before, rejoined: NOT_ATTENDING.indexOf(before) >= 0 && ATTENDING.indexOf(p.status) >= 0,
             message: '登録済みの記録の状態を「' + before + '」から「' + p.status + '」に変更しました（行は増やしていません）。',
             record: recordOut_(dup, sem.seminar) };
  }
  if (dup) return { duplicate: true, message: 'この人はこのセミナーに登録済みです（重複登録しません）。変更なら update を使ってください。',
                    record: recordOut_(dup, sem.seminar) };
  const s = sem.seminar;
  const rec = {
    '記録ID': newId_('R'), '人物ID': per.person['人物ID'], '氏名': per.person['氏名'], 'セミナーID': s.seminar_id,
    '開催日': s.date, '時間': s.start + (s.end ? '\n' + s.end : ''), '種別': s.type, '講師': s.lecturer, '地域': s.pref, '会場': s.venue,
    '案内日': dateOrEmpty_(p.invited_on), '連絡手段': p.channel || '', '返事': p.reply || '', '状態': status,
    '次回連絡日': dateOrEmpty_(p.next_contact), 'フォロー状況': followFor_(status, p.follow || ''), 'メモ': p.memo || '',
    '登録日時': now_(), '更新日時': now_(),
  };
  appendObject_(ss, RECORD_SHEET, RECORD_HEADERS, rec);
  logChange_(ss, '登録', rec, '', status, { reason: p.reason || p.reply || '', source: p.source_ref || p.memo || '' });
  return { ok: true, created: true, new_person: !!per.new_person, record: recordOut_(rec, s) };
}

function updateAttendance_(ss, p) {
  if (p.status && STATUSES.indexOf(p.status) < 0) return { error: 'bad_status', allowed: STATUSES };
  const sh = sheet_(ss, RECORD_SHEET, RECORD_HEADERS);
  const records = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS);
  let targets;
  if (p.record_id) {
    targets = records.filter(r => r['記録ID'] === p.record_id);
  } else {
    const per = resolvePerson_(ss, p, false);
    if (!per.person) return per;
    targets = records.filter(r => r['人物ID'] === per.person['人物ID']);
    if (p.seminar_id || p.date || p.lecturer) {
      const sem = resolveSeminar_(ss, p);
      if (!sem.seminar) return sem;
      targets = targets.filter(r => r['セミナーID'] === sem.seminar.seminar_id);
    } else if (targets.length > 1) {
      const open = targets.filter(r => CLOSED_STATUSES.indexOf(r['状態']) < 0);
      if (open.length === 1) targets = open;
    }
  }
  if (!targets.length) return { error: 'record_not_found', message: '対象の参加記録が見つかりません。' };
  const all = listSeminars_(ss);
  if (targets.length > 1) return { needs_confirmation: 'record', message: '対象のセミナーが複数あります。どれを変更しますか。',
    choices: targets.map(r => ({ record_id: r['記録ID'], label: r['氏名'] + '｜' + seminarLabel_(currentSeminar_(all, seminarFromRecord_(r))) + '｜' + r['状態'] })) };
  const rec = targets[0];
  const before = rec['状態'];
  if (p.status) rec['状態'] = p.status;
  if (p.reply) rec['返事'] = p.reply;
  if (p.channel) rec['連絡手段'] = p.channel;
  if (p.invited_on) rec['案内日'] = dateOrEmpty_(p.invited_on);
  if (p.next_contact) rec['次回連絡日'] = p.next_contact === 'clear' ? '' : dateOrEmpty_(p.next_contact);
  if (p.follow) rec['フォロー状況'] = p.follow;
  else if (p.status) rec['フォロー状況'] = followFor_(p.status, rec['フォロー状況']);
  if (p.memo) rec['メモ'] = p.memo_mode === 'replace' ? p.memo
    : (rec['メモ'] ? rec['メモ'] + '\n' : '') + Utilities.formatDate(new Date(), TZ, 'M/d') + ' ' + p.memo;
  rec['更新日時'] = now_();
  saveRow_(sh, RECORD_HEADERS, rec);
  if (p.status && p.status !== before) logChange_(ss, '状態変更', rec, before, p.status, { reason: p.reason || p.reply || '', source: p.source_ref || '' });
  return { ok: true, updated: true, status_before: before, record: recordOut_(rec, currentSeminar_(all, seminarFromRecord_(rec))) };
}

function followFor_(status, current) {
  if (status === 'フォロー済み') return '済み';
  if (status === '参加済み') return current && current !== '済み' ? current : '未フォロー';
  return current || '';
}

function recordOut_(r, s) {
  return { record_id: r['記録ID'], person_id: r['人物ID'], name: r['氏名'], seminar_id: r['セミナーID'],
           seminar: s ? seminarLabel_(s) : '', status: r['状態'], reply: r['返事'], channel: r['連絡手段'],
           invited_on: r['案内日'], next_contact: r['次回連絡日'], follow: r['フォロー状況'], memo: r['メモ'],
           organizer: s && s.organizer ? s.organizer : undefined, note: s && s.note ? s.note : undefined };
}

// ===== 検索・名簿 =====

function seminarRoster_(ss, p) {
  const sem = resolveSeminar_(ss, p);
  if (!sem.seminar) return sem;
  const s = sem.seminar;
  const rows = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS).filter(r => r['セミナーID'] === s.seminar_id)
    .sort((a, b) => STATUS_ORDER.indexOf(a['状態']) - STATUS_ORDER.indexOf(b['状態']) || String(a['氏名']).localeCompare(String(b['氏名']), 'ja'));
  const header = { date: s.date, time: s.start + (s.end ? '〜' + s.end : ''), lecturer: s.lecturer, type: s.type,
                   pref: s.pref, venue: s.venue, organizer: s.organizer, canceled: s.canceled, in_list: s.in_list };
  const res = { seminar_id: s.seminar_id, header: header, count: rows.length,
                participants: rows.map(r => ({ name: r['氏名'], status: r['状態'], record_id: r['記録ID'], next_contact: r['次回連絡日'] })) };
  if (p.format === 'pdf') {
    const head = [['開催日', s.date], ['時間', header.time], ['講師', s.lecturer], ['種別', s.type], ['地域', s.pref],
                  ['会場', s.venue + (s.canceled ? '（中止）' : '')]];
    const table = '<table class="h">' + head.map(h => '<tr><th>' + h[0] + '</th><td>' + esc_(h[1]) + '</td></tr>').join('') + '</table>'
      + '<h2>参加者（' + rows.length + '名）</h2><table><tr><th>No</th><th>氏名</th><th>状態</th></tr>'
      + rows.map((r, i) => '<tr><td>' + (i + 1) + '</td><td>' + esc_(r['氏名']) + '</td><td>' + esc_(r['状態']) + '</td></tr>').join('') + '</table>';
    const f = savePdf_(ss, '参加者名簿（' + s.date + ' ' + s.type + ' ' + s.lecturer + '）', table,
      '名簿_' + Utilities.formatDate(parseDate_(s.date) || new Date(), TZ, 'yyyyMMdd') + '_' + s.type + '_' + norm_(s.lecturer));
    res.pdf_url = f.getUrl(); res.pdf_name = f.getName();
  }
  return res;
}

function personHistory_(ss, p) {
  const per = resolvePerson_(ss, p, false);
  if (!per.person) return per;
  const all = listSeminars_(ss);
  const rows = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS).filter(r => r['人物ID'] === per.person['人物ID'])
    .map(r => ({ r: r, s: currentSeminar_(all, seminarFromRecord_(r)) }))
    .sort((a, b) => dateKey_(a.s.date).localeCompare(dateKey_(b.s.date)));
  const want = splitList_(p.status);
  const list = rows.filter(x => !want.length || want.indexOf(x.r['状態']) >= 0);
  const res = { person: personLabel_(per.person), person_id: per.person['人物ID'], count: list.length,
                history: list.map(x => recordOut_(x.r, x.s)) };
  if (p.format === 'pdf') {
    const table = '<table><tr><th>開催日</th><th>時間</th><th>講師</th><th>種別</th><th>地域</th><th>会場</th><th>状態</th></tr>'
      + list.map(x => '<tr>' + [x.s.date, x.s.start + '〜' + x.s.end, x.s.lecturer, x.s.type, x.s.pref, x.s.venue, x.r['状態']]
        .map(v => '<td>' + esc_(v) + '</td>').join('') + '</tr>').join('') + '</table>';
    const f = savePdf_(ss, '参加履歴（' + per.person['氏名'] + '）', table, '参加履歴_' + per.person['氏名']);
    res.pdf_url = f.getUrl(); res.pdf_name = f.getName();
  }
  return res;
}

// kind: waiting（返事待ち）/ today（今日までに連絡する人）/ not_followed（参加後まだフォローしていない）/ status（状態で絞る）
function followups_(ss, p) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd');
  const all = listSeminars_(ss);
  const kind = p.kind || 'today';
  const want = splitList_(p.status);
  const rows = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS).filter(r => {
    if (kind === 'waiting') return r['状態'] === '返事待ち' || r['状態'] === '案内済み';
    if (kind === 'not_followed') return r['状態'] === '参加済み' && r['フォロー状況'] !== '済み';
    if (kind === 'status') return want.indexOf(r['状態']) >= 0;
    return r['次回連絡日'] && r['次回連絡日'] <= today && CLOSED_STATUSES.indexOf(r['状態']) < 0;
  }).map(r => {
    const out = recordOut_(r, currentSeminar_(all, seminarFromRecord_(r)));
    if (kind === 'today' && r['次回連絡日'] < today) out.overdue = true;
    return out;
  }).sort((a, b) => String(a.next_contact || '9').localeCompare(String(b.next_contact || '9')));
  const res = { kind: kind, today: today, count: rows.length, items: rows };
  if (p.format === 'pdf') {
    const title = { waiting: '返事待ち', today: '今日連絡する人', not_followed: '参加後 未フォロー', status: '状態：' + want.join('・') }[kind] || kind;
    const table = '<table><tr><th>氏名</th><th>セミナー</th><th>状態</th><th>次回連絡日</th><th>連絡手段</th><th>メモ</th></tr>'
      + rows.map(x => '<tr>' + [x.name, x.seminar, x.status, x.next_contact, x.channel, x.memo]
        .map(v => '<td>' + esc_(v || '') + '</td>').join('') + '</tr>').join('') + '</table>';
    const f = savePdf_(ss, title + '（' + today + '）', table, title);
    res.pdf_url = f.getUrl(); res.pdf_name = f.getName();
  }
  return res;
}

// ===== 確認待ち候補（PLAUD・ボイスメモ・LINE画像・画像メモから） =====
// 候補は参加記録に入らない。candidate_confirm で承認された時だけ登録する。

function addCandidate_(ss, p) {
  const status = p.status || '未定';
  if (STATUSES.indexOf(status) < 0) return { error: 'bad_status', allowed: STATUSES };
  if (!p.source_ref) return { error: 'source_ref_missing', message: '資料名（録音タイトル・画像名・日時など）を指定してください。' };
  let seminarId = '', seminarText = p.seminar_text || '', choices;
  if (p.seminar_id || p.date || p.lecturer) {
    const sem = resolveSeminar_(ss, p);
    if (sem.seminar) { seminarId = sem.seminar.seminar_id; seminarText = seminarLabel_(sem.seminar); }
    else if (sem.choices) { choices = sem.choices; seminarText = seminarText || '候補複数：' + sem.choices.map(c => c.label).join(' / '); }
  }
  const nameUnclear = !p.name || /[?？]|判読|不明/.test(p.name);
  // 同じ資料の同じ人は1件だけ（再処理でセミナーの伝え方が変わっても重複させない）
  const key = [norm_(p.source_ref), norm_(p.name)].join('|');
  const existing = readObjects_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS)
    .filter(c => norm_(c['資料']) === norm_(p.source_ref) && norm_(c['氏名']) === norm_(p.name || '（判読不可）'))[0];
  if (existing) return { duplicate: true, message: '同じ資料から同じ候補が登録済みです（重複しません）。', candidate: candidateOut_(existing) };
  const people = nameUnclear ? [] : findPeople_(ss, p);
  const c = {
    '候補ID': newId_('C'), '入口': p.source_type || '', '資料': p.source_ref, '重複キー': key, '根拠': p.evidence || '',
    '氏名': p.name || '（判読不可）', '人物ID': people.length === 1 ? people[0]['人物ID'] : '', 'セミナーID': seminarId,
    'セミナー表示': seminarText, '提案状態': status, '次回連絡日': dateOrEmpty_(p.next_contact), 'メモ': p.memo || '',
    '確認状態': '確認待ち', '記録ID': '', '作成日時': now_(),
  };
  appendObject_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS, c);
  const out = { ok: true, created: true, candidate: candidateOut_(c) };
  if (nameUnclear) out.warning = '名前が確定していません。確認時に正しい名前を指定してください。';
  if (!seminarId) out.warning = (out.warning ? out.warning + ' ' : '') + 'セミナーが1件に決まっていません。確認時に指定してください。';
  if (people.length > 1) out.warning = (out.warning ? out.warning + ' ' : '') + '同名の人物が複数います。確認時に指定してください。';
  if (choices) out.seminar_choices = choices;
  return out;
}

function candidateOut_(c) {
  return { candidate_id: c['候補ID'], source_type: c['入口'], source_ref: c['資料'], name: c['氏名'], person_id: c['人物ID'],
           seminar_id: c['セミナーID'], seminar: c['セミナー表示'], proposed_status: c['提案状態'], next_contact: c['次回連絡日'],
           evidence: c['根拠'], memo: c['メモ'], state: c['確認状態'], record_id: c['記録ID'] };
}

function listCandidates_(ss, p) {
  const state = p.state || '確認待ち';
  const list = readObjects_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS).filter(c => state === 'all' || c['確認状態'] === state);
  return { count: list.length, candidates: list.map(candidateOut_) };
}

// decision=approve：参加記録へ登録（重複は登録しない）。decision=reject：却下。
// name / person_id / seminar_id / status / next_contact を渡すと、候補の内容を直してから登録する。
function confirmCandidates_(ss, p) {
  const ids = splitList_(p.candidate_id);
  if (!ids.length) return { error: 'candidate_id_missing' };
  const sh = sheet_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS);
  const all = readObjects_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS);
  const results = ids.map(id => {
    const c = all.filter(x => x['候補ID'] === id)[0];
    if (!c) return { candidate_id: id, error: 'candidate_not_found' };
    if (c['確認状態'] !== '確認待ち') return { candidate_id: id, error: 'already_' + c['確認状態'], record_id: c['記録ID'] };
    if (p.decision === 'reject') {
      c['確認状態'] = '却下';
    } else if (p.decision === 'approve') {
      const one = ids.length === 1;
      const req = {
        name: (one && p.name) || c['氏名'], person_id: (one && p.person_id) || c['人物ID'], new_person: one ? p.new_person : '',
        kana: one ? p.kana : '', person_note: one ? p.person_note : '',
        seminar_id: (one && p.seminar_id) || c['セミナーID'], status: (one && p.status) || c['提案状態'],
        next_contact: (one && p.next_contact) || c['次回連絡日'], channel: one ? p.channel : '',
        memo: [c['メモ'], '出典：' + c['入口'] + ' ' + c['資料']].filter(Boolean).join(' / '),
      };
      if (!validName_(req.name)) return { candidate_id: id, error: 'name_unclear', message: '正しい名前を name で指定してください。' };
      if (!req.seminar_id) return { candidate_id: id, error: 'seminar_unresolved', message: 'seminar_id を指定してください。', seminar: c['セミナー表示'] };
      const r = registerAttendance_(ss, req);
      if (r.needs_confirmation || r.error) return Object.assign({ candidate_id: id }, r);
      c['確認状態'] = '登録済み';
      c['記録ID'] = r.record.record_id;
      c['人物ID'] = r.record.person_id;
      writeCandidate_(sh, c);
      return { candidate_id: id, registered: !r.duplicate, duplicate: !!r.duplicate, record: r.record };
    } else {
      return { candidate_id: id, error: 'decision_must_be_approve_or_reject' };
    }
    writeCandidate_(sh, c);
    return { candidate_id: id, state: c['確認状態'] };
  });
  return { results: results };
}

function writeCandidate_(sh, c) {
  saveRow_(sh, CANDIDATE_HEADERS, c);
}

// ===== ボイスメモ・PLAUDの文字起こし（Googleドライブ） =====
// AI秘書が保存しているノート（…__voicememo-ID.md / …__plaud-ID.md）を新しい順に読み、本文をChatGPTに渡す。
// 同じIDが複数のフォルダにある時は1つだけ（大きい方）にする。人や参加意思を拾うのはChatGPT側で、候補は確認待ちに入るだけ。

const MEMO_PATTERN = /__(voicememo|plaud)-([A-Za-z0-9_-]+)\.md$/;
const MEMO_TEXT_LIMIT = 8000;    // 1件あたりの文字数
const MEMO_TOTAL_LIMIT = 30000;  // 1回の応答の合計文字数

function findMemos_(p) {
  const kw = norm_(p.keyword);
  const days = Math.min(Number(p.days) || (kw ? 60 : 3), 365);
  const since = Utilities.formatDate(new Date(Date.now() - days * 86400000), 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
  const kinds = p.source === 'voicememo' || p.source === 'plaud' ? [p.source] : ['voicememo', 'plaud'];
  const byKey = {};
  kinds.forEach(k => {
    const it = DriveApp.searchFiles("title contains '__" + k + "-' and trashed = false and modifiedDate > '" + since + "'");
    while (it.hasNext()) {
      const f = it.next();
      const m = f.getName().match(MEMO_PATTERN);
      if (!m) continue;
      const key = m[1] + '-' + m[2];
      if (!byKey[key] || f.getSize() > byKey[key].getSize()) byKey[key] = f;
    }
  });
  return Object.keys(byKey).map(k => ({ key: k, file: byKey[k], source_ref: byKey[k].getName().replace(/\.md$/, '') }))
    .filter(x => !kw || norm_(x.source_ref).indexOf(kw) >= 0)
    .sort((a, b) => b.source_ref.localeCompare(a.source_ref));  // ファイル名は録音日時から始まる
}

function memoDoneSet_(ss) {
  return new Set(readObjects_(ss, MEMO_DONE_SHEET, MEMO_DONE_HEADERS).map(x => norm_(x['資料'])));
}

function listMemos_(ss, p) {
  const done = memoDoneSet_(ss);
  const limit = Math.min(Number(p.limit) || 3, 5);
  let memos = findMemos_(p);
  memos.forEach(x => x.processed = done.has(norm_(x.source_ref)));
  const found = memos.length;
  if (p.include_processed !== 'true') memos = memos.filter(x => !x.processed);
  const remaining = memos.length;
  memos = memos.slice(0, limit);
  const per = Math.min(MEMO_TEXT_LIMIT, Math.floor(MEMO_TOTAL_LIMIT / Math.max(memos.length, 1)));
  return { found: found, unprocessed: remaining, returned: memos.length, items: memos.map(x => {
    const text = x.file.getBlob().getDataAsString('UTF-8');
    return { source_ref: x.source_ref, source_type: x.key.indexOf('plaud') === 0 ? 'PLAUD' : 'ボイスメモ',
             processed: x.processed, text: text.slice(0, per), truncated: text.length > per };
  }) };
}

// 1件の資料を読み終えたら呼ぶ（候補0件でも）。次からは memos に出てこない。
function markMemoDone_(ss, p) {
  if (!p.source_ref) return { error: 'source_ref_missing' };
  if (memoDoneSet_(ss).has(norm_(p.source_ref))) return { ok: true, already: true };
  appendObject_(ss, MEMO_DONE_SHEET, MEMO_DONE_HEADERS, { '資料': p.source_ref, '処理日時': now_(), '候補数': p.count || '' });
  return { ok: true, marked: true };
}

// ===== 参加者の増減（追加・不参加・再参加・名簿の照合） =====

// 「山田、田中さん、佐藤 様」→ ['山田', '田中', '佐藤']（空白は姓名の間なので区切りにしない）
function names_(v) {
  const seen = {};
  return String(v || '').split(/[,、，;；\/／\n]+/)
    .map(x => x.trim().replace(/(さん|様|さま|君|くん|ちゃん)$/, '').trim())
    .filter(x => x && !seen[norm_(x)] && (seen[norm_(x)] = true));
}

// 名前として使えるか（判読不可・？付き・IDの貼り間違いを弾く）。
// 「P山」「石原GM」「東谷け」のような呼び名はJPの実際の書き方なのでそのまま受け付ける。
function validName_(name) {
  const n = String(name || '').trim();
  if (!n || /[?？]|判読|不明/.test(n)) return false;
  if (/^[PRCH]\d{6,}/.test(n)) return false;
  return true;
}

function rememberSeminar_(id) {
  PropertiesService.getScriptProperties().setProperty('LAST_SEMINAR', JSON.stringify({ id: id, at: Date.now() }));
}

function lastSeminar_(ss) {
  try {
    const v = JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_SEMINAR') || 'null');
    if (!v || Date.now() - v.at > LAST_SEMINAR_HOURS * 3600000) return null;
    const r = resolveSeminarRaw_(ss, { seminar_id: v.id });
    return r.seminar || null;
  } catch (err) {
    return null;
  }
}

// セミナーの決め方：明示（seminar_id・日付・講師など）> 直近24時間に扱ったセミナー
function pickSeminar_(ss, p) {
  if (p.seminar_id || p.date || p.lecturer) {
    const r = resolveSeminar_(ss, p);
    return r.seminar ? { seminar: r.seminar, from: 'specified' } : r;
  }
  const last = lastSeminar_(ss);
  if (last) { rememberSeminar_(last.seminar_id); return { seminar: last, from: 'last_used' }; }
  return { error: 'seminar_missing', message: 'どのセミナーか分かりません。開催日・講師・種別などを教えてください。' };
}

function newResult_() {
  return { added: [], rejoined: [], left: [], unchanged: [], pending: [], not_registered: [], needs_choice: [] };
}

function changeInfo_(p) {
  return { reason: p.reason || '', source: [p.source_type, p.source_ref].filter(Boolean).join(' ') };
}

// 判断できない人は確認待ち候補に入れる（参加記録は変えない）。同じ資料・同じ人は増えない。
function toPending_(ss, p, s, name, proposed, why, res) {
  const c = addCandidate_(ss, {
    source_type: p.source_type || '貼り付け',
    source_ref: p.source_ref || ('増減_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '_' + (s ? s.seminar_id : '')),
    name: name, seminar_id: s ? s.seminar_id : '', status: proposed,
    evidence: [why, p.reason].filter(Boolean).join(' / '), memo: why,
  });
  res.pending.push({ name: name, why: why, proposed_status: proposed, candidate_id: c.candidate && c.candidate.candidate_id,
                     duplicate: !!c.duplicate });
}

// 1人分の変更。intent: join（参加予定にする）/ leave（不参加にする）
function changeOne_(ss, p, s, name, intent, res, records) {
  if (!validName_(name)) return toPending_(ss, p, s, name, intent === 'join' ? '参加予定' : 'キャンセル', '名前が確定していません', res);
  const people = findPeople_(ss, { name: name });
  if (people.length > 1) return toPending_(ss, p, s, name, intent === 'join' ? '参加予定' : 'キャンセル', '同じ名前の人が' + people.length + '人います', res);
  let person = people[0];
  if (!person) {
    if (intent === 'leave') return res.not_registered.push({ name: name, why: '人物に登録されていません（変更なし）' });
    const created = resolvePerson_(ss, { name: name, new_person: 'true' }, true);
    if (!created.person) return toPending_(ss, p, s, name, '参加予定', created.message || '登録できませんでした', res);
    person = created.person;
  }
  const rec = records.filter(r => r['人物ID'] === person['人物ID'] && r['セミナーID'] === s.seminar_id)[0];
  const info = changeInfo_(p);
  if (intent === 'join') {
    if (!rec) {
      const r = {
        '記録ID': newId_('R'), '人物ID': person['人物ID'], '氏名': person['氏名'], 'セミナーID': s.seminar_id,
        '開催日': s.date, '時間': s.start + (s.end ? '\n' + s.end : ''), '種別': s.type, '講師': s.lecturer, '地域': s.pref, '会場': s.venue,
        '案内日': '', '連絡手段': p.channel || '', '返事': '', '状態': '参加予定', '次回連絡日': '', 'フォロー状況': '',
        'メモ': info.source ? '出典：' + info.source : '', '登録日時': now_(), '更新日時': now_(),
      };
      appendObject_(ss, RECORD_SHEET, RECORD_HEADERS, r);
      records.push(r);
      logChange_(ss, '追加', r, '', '参加予定', info);
      return res.added.push({ name: person['氏名'], record_id: r['記録ID'], status_before: '' });
    }
    const before = rec['状態'];
    if (ATTENDING.indexOf(before) >= 0) return res.unchanged.push({ name: person['氏名'], status: before });
    setStatus_(ss, rec, '参加予定', Object.assign({ op: NOT_ATTENDING.indexOf(before) >= 0 ? '再参加' : '追加' }, info));
    return (NOT_ATTENDING.indexOf(before) >= 0 ? res.rejoined : res.added)
      .push({ name: person['氏名'], record_id: rec['記録ID'], status_before: before });
  }
  // leave
  if (!rec) return res.not_registered.push({ name: person['氏名'], why: 'このセミナーの参加記録がありません（変更なし）' });
  const before = rec['状態'];
  if (NOT_ATTENDING.indexOf(before) >= 0) return res.unchanged.push({ name: person['氏名'], status: before });
  if (before === '参加済み' || before === 'フォロー済み') {
    return toPending_(ss, p, s, person['氏名'], 'キャンセル', 'すでに「' + before + '」のため自動では不参加にしません', res);
  }
  const after = before === '参加予定' ? 'キャンセル' : '辞退';
  setStatus_(ss, rec, after, Object.assign({ op: '不参加' }, info));
  res.left.push({ name: person['氏名'], record_id: rec['記録ID'], status_before: before, status: after });
}

function setStatus_(ss, rec, status, info) {
  const before = rec['状態'];
  rec['状態'] = status;
  rec['フォロー状況'] = followFor_(status, rec['フォロー状況']);
  rec['更新日時'] = now_();
  saveRow_(sheet_(ss, RECORD_SHEET, RECORD_HEADERS), RECORD_HEADERS, rec);
  logChange_(ss, info.op || '状態変更', rec, before, status, info);
}

// 参加予定のセミナーが本人に1件だけなら、それを不参加の対象にする（「田中さんが来られなくなった」用）
function leaveWithoutSeminar_(ss, p, name, res, records) {
  const people = findPeople_(ss, { name: name });
  if (people.length !== 1) {
    if (people.length > 1) return res.needs_choice.push({ name: name, message: '同じ名前の人が複数います。セミナーか人を指定してください。',
      choices: people.map(x => ({ person_id: x['人物ID'], label: personLabel_(x) })) });
    return res.not_registered.push({ name: name, why: '人物に登録されていません（変更なし）' });
  }
  const today = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  const all = listSeminars_(ss);
  const open = records.filter(r => r['人物ID'] === people[0]['人物ID'] && NOT_ATTENDING.indexOf(r['状態']) < 0
    && ['参加済み', 'フォロー済み'].indexOf(r['状態']) < 0 && dateKey_(r['開催日']) >= today);
  if (!open.length) return res.not_registered.push({ name: people[0]['氏名'], why: '今後の参加予定がありません（変更なし）' });
  if (open.length > 1) return res.needs_choice.push({ name: people[0]['氏名'], message: '今後の予定が' + open.length + '件あります。どれを不参加にしますか。',
    choices: open.map(r => ({ seminar_id: r['セミナーID'], label: seminarLabel_(currentSeminar_(all, seminarFromRecord_(r))) + '｜' + r['状態'] })) });
  const s = currentSeminar_(all, seminarFromRecord_(open[0]));
  const before = res.left.length;
  changeOne_(ss, p, s, people[0]['氏名'], 'leave', res, records);
  if (res.left.length > before) res.left[res.left.length - 1].seminar = seminarLabel_(s);
}

function summary_(ss, s, res, extra) {
  const recs = s ? readObjects_(ss, RECORD_SHEET, RECORD_HEADERS).filter(r => r['セミナーID'] === s.seminar_id) : [];
  const cands = s ? readObjects_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS)
    .filter(c => c['セミナーID'] === s.seminar_id && c['確認状態'] === '確認待ち') : [];
  return Object.assign({
    ok: true,
    seminar_id: s ? s.seminar_id : '', seminar: s ? seminarLabel_(s) : '',
    counts: {
      added: res.added.length, left: res.left.length, rejoined: res.rejoined.length, unchanged: res.unchanged.length,
      pending: res.pending.length, not_registered: res.not_registered.length, needs_choice: res.needs_choice.length,
      attending_now: s ? recs.filter(r => r['状態'] === '参加予定' || r['状態'] === '参加済み').length : null,
      pending_now: s ? cands.length : null,
    },
    attending_names: s ? recs.filter(r => r['状態'] === '参加予定' || r['状態'] === '参加済み').map(r => r['氏名']) : [],
  }, res, extra || {});
}

// join / leave / undecided に名前を「、」区切りで入れる。理由は reason、資料は source_type / source_ref。
function applyChanges_(ss, p) {
  const join = names_(p.join), leave = names_(p.leave), undecided = names_(p.undecided);
  if (!join.length && !leave.length && !undecided.length) {
    return { error: 'names_missing', message: 'join（参加）・leave（不参加）・undecided（判断できない）のどれかに名前を入れてください。' };
  }
  const records = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS);
  const res = newResult_();
  const explicit = p.seminar_id || p.date || p.lecturer;
  if (!explicit && !join.length && !undecided.length) {
    leave.forEach(n => leaveWithoutSeminar_(ss, p, n, res, records));
    return summary_(ss, null, res, { seminar_from: 'per_person' });
  }
  const sem = pickSeminar_(ss, p);
  if (!sem.seminar) return sem;
  const s = sem.seminar;
  join.forEach(n => changeOne_(ss, p, s, n, 'join', res, records));
  leave.forEach(n => changeOne_(ss, p, s, n, 'leave', res, records));
  undecided.forEach(n => toPending_(ss, p, s, n, p.status || '未定', '参加か不参加か判断できません', res));
  return summary_(ss, s, res, { seminar_from: sem.from });
}

// 名簿との照合。names = 名簿の全員。complete=true（最新の完全な名簿と確認できた時）だけ、名簿に無い参加予定者を不参加にする。
function rosterSync_(ss, p) {
  const list = names_(p.names);
  if (!list.length) return { error: 'names_missing', message: 'names に名簿の名前を「、」区切りで入れてください。' };
  const sem = pickSeminar_(ss, p);
  if (!sem.seminar) return sem;
  const s = sem.seminar;
  const complete = String(p.complete) === 'true';
  const records = readObjects_(ss, RECORD_SHEET, RECORD_HEADERS);
  const people = readObjects_(ss, PERSON_SHEET, PERSON_HEADERS);
  const listed = list.map(norm_);
  const isListed = r => {
    const per = people.filter(x => x['人物ID'] === r['人物ID'])[0] || {};
    return listed.indexOf(norm_(r['氏名'])) >= 0 || (per['読み'] && listed.indexOf(norm_(per['読み'])) >= 0);
  };
  const missing = records.filter(r => r['セミナーID'] === s.seminar_id && r['状態'] === '参加予定' && !isListed(r));
  const attendingBefore = records.filter(r => r['セミナーID'] === s.seminar_id && r['状態'] === '参加予定').length;
  // 安全装置：完全な名簿でも、参加予定の半分以上（3人以上）が一度に消える時は確認を求める
  if (complete && missing.length >= 3 && missing.length * 2 >= attendingBefore && String(p.force) !== 'true') {
    return { needs_confirmation: 'mass_leave', seminar_id: s.seminar_id, seminar: seminarLabel_(s),
             message: '名簿に無い参加予定者が' + missing.length + '人（参加予定' + attendingBefore + '人中）います。本当に全員を不参加にしてよいか確認してください。よければ force=true で呼び直します。',
             would_leave: missing.map(r => r['氏名']) };
  }
  const res = newResult_();
  list.forEach(n => changeOne_(ss, p, s, n, 'join', res, records));
  const kept = [];
  missing.forEach(r => {
    if (complete) changeOne_(ss, p, s, r['氏名'], 'leave', res, records);
    else kept.push(r['氏名']);
  });
  return summary_(ss, s, res, { seminar_from: sem.from, complete: complete,
    not_in_list_kept: kept, note: complete ? '' : '最新の完全な名簿と確認できないため、名簿に名前が無い人は不参加にしていません。' });
}

// ===== 変更履歴 =====

function logChange_(ss, op, rec, before, after, info) {
  appendObject_(ss, HISTORY_SHEET, HISTORY_HEADERS, {
    '履歴ID': newId_('H'), '日時': Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss'), '操作': op,
    '記録ID': rec['記録ID'] || '', '人物ID': rec['人物ID'] || '', '氏名': rec['氏名'] || '',
    'セミナーID': rec['セミナーID'] || '', '開催日': rec['開催日'] || '',
    '変更前': before || '', '変更後': after || '', '理由': (info && info.reason) || '', '出典': (info && info.source) || '',
  });
}

function listChanges_(ss, p) {
  const n = norm_(p.name);
  const limit = Math.min(Number(p.limit) || 20, 100);
  const list = readObjects_(ss, HISTORY_SHEET, HISTORY_HEADERS)
    .filter(h => (!n || norm_(h['氏名']).indexOf(n) >= 0) && (!p.seminar_id || h['セミナーID'] === p.seminar_id))
    .reverse().slice(0, limit);
  return { count: list.length, changes: list.map(h => ({ at: h['日時'], op: h['操作'], name: h['氏名'], seminar_id: h['セミナーID'],
    date: h['開催日'], before: h['変更前'], after: h['変更後'], reason: h['理由'], source: h['出典'] })) };
}

// 氏名の訂正：人物シートと、その人の参加記録の氏名を直す（行は消さない・履歴に残す）
function renamePerson_(ss, p) {
  const newName = String(p.new_name || '').trim().replace(/(さん|様|さま|君|くん|ちゃん)$/, '');
  if (!validName_(newName)) return { error: 'name_unclear', message: '新しい名前 new_name を指定してください。' };
  const per = resolvePerson_(ss, p, false);
  if (!per.person) return per;
  const old = per.person['氏名'];
  if (old === newName) return { ok: true, unchanged: true, person_id: per.person['人物ID'], name: newName };
  const dup = findPeople_(ss, { name: newName }).filter(x => x['人物ID'] !== per.person['人物ID']);
  if (dup.length && String(p.force) !== 'true') return { needs_confirmation: 'name_exists',
    message: '「' + newName + '」は別の人物として登録済みです。同じ人なら統合が必要なので、JPに確認してください。', choices: dup.map(x => ({ person_id: x['人物ID'], label: personLabel_(x) })) };
  per.person['氏名'] = newName;
  saveRow_(sheet_(ss, PERSON_SHEET, PERSON_HEADERS), PERSON_HEADERS, per.person);
  const info = { reason: p.reason || '氏名の訂正', source: p.source_ref || '' };
  logChange_(ss, '氏名訂正', { '人物ID': per.person['人物ID'], '氏名': newName }, old, newName, info);
  const sh = sheet_(ss, RECORD_SHEET, RECORD_HEADERS);
  let n = 0;
  readObjects_(ss, RECORD_SHEET, RECORD_HEADERS).filter(r => r['人物ID'] === per.person['人物ID']).forEach(r => {
    r['氏名'] = newName; r['更新日時'] = now_();
    saveRow_(sh, RECORD_HEADERS, r);
    n++;
  });
  return { ok: true, renamed: true, person_id: per.person['人物ID'], before: old, after: newName, records_updated: n };
}

// ===== シート操作 =====

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  }
  return sh;
}

function readObjects_(ss, name, headers) {
  const sh = sheet_(ss, name, headers);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, headers.length).getDisplayValues()
    .filter(r => r[0] !== '')
    .map(r => headers.reduce((o, h, i) => (o[h] = r[i], o), {}));
}

function appendObject_(ss, name, headers, obj) {
  saveRow_(sheet_(ss, name, headers), headers, obj);
}

// ===== 保存の検証（書く→読み直す→合わなければ書き直す） =====

const SYNC_ = { writes: 0, verified: 0, retries: 0, errors: [] };

function syncReset_() {
  SYNC_.writes = 0; SYNC_.verified = 0; SYNC_.retries = 0; SYNC_.errors = [];
}

function syncOut_() {
  return { ok: SYNC_.errors.length === 0, writes: SYNC_.writes, verified: SYNC_.verified, retries: SYNC_.retries,
           errors: SYNC_.errors.slice(0, 5) };
}

// 1列目のIDで行を探して書く（無ければ末尾に追加）。書いた後に読み直して全列が一致するか確かめ、
// 合わなければ最大3回まで書き直す。再試行でも同じIDの行に書くので、重複行はできない。
function saveRow_(sh, headers, obj) {
  const id = String(obj[headers[0]] == null ? '' : obj[headers[0]]);
  if (!id) throw new Error('save_without_id');
  const values = headers.map(h => obj[h] == null ? '' : String(obj[h]));
  SYNC_.writes++;
  let lastErr = 'verify_mismatch';
  let row = 0;  // 1回目に決めた行を覚えておき、書き損じた時も同じ行に書き直す（行が増えない）
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      row = findRow_(sh, id) || row || sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, headers.length).setNumberFormat('@').setValues([values]);
      SpreadsheetApp.flush();
      const back = sh.getRange(row, 1, 1, headers.length).getDisplayValues()[0];
      if (back.every((v, i) => v === values[i]) && findRow_(sh, id) === row) {
        SYNC_.verified++;
        return row;
      }
    } catch (err) {
      lastErr = String(err && err.message || err);
    }
    if (attempt < 3) { SYNC_.retries++; Utilities.sleep(400 * attempt); }
  }
  SYNC_.errors.push(sh.getName() + ' ' + id + ': ' + lastErr);
  throw new Error('sync_failed: ' + sh.getName() + ' ' + id + '（3回書き直しても保存を確認できませんでした）');
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Utilities.getUuid().slice(0, 4);
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
}

function dateOrEmpty_(v) {
  const d = parseDate_(v || '');
  return d ? Utilities.formatDate(d, TZ, 'yyyy/MM/dd') : '';
}

function dateKey_(v) {
  const d = parseDate_(v);
  return d ? Utilities.formatDate(d, TZ, 'yyyyMMdd') : '99999999';
}

function esc_(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

function savePdf_(ss, title, bodyHtml, baseName) {
  const html = '<html><head><meta charset="utf-8"><style>'
    + 'body{font-family:sans-serif;font-size:10pt}h1{font-size:14pt}h2{font-size:12pt;margin-top:14px}'
    + 'table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:4px 6px;vertical-align:middle;text-align:left}'
    + 'th{background:#eee;white-space:nowrap}table.h th{width:70px}</style></head><body>'
    + '<h1>' + esc_(title) + '</h1>' + bodyHtml
    + '<p style="color:#666">作成：' + Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') + '</p></body></html>';
  const name = baseName.replace(/[\\\/:*?"<>|\s]/g, '_') + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '.pdf';
  const blob = Utilities.newBlob(html, 'text/html', 'p.html').getAs(MimeType.PDF).setName(name);
  return getSubFolder_(ss, PDF_FOLDER).createFile(blob);
}

// 初回だけ実行：3つのシートを作る（既にあれば何もしない）。
function setupAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, PERSON_SHEET, PERSON_HEADERS);
  sheet_(ss, RECORD_SHEET, RECORD_HEADERS);
  sheet_(ss, CANDIDATE_SHEET, CANDIDATE_HEADERS);
  sheet_(ss, HISTORY_SHEET, HISTORY_HEADERS);
}

// ===== ChatGPT設定（chatgptSettings から呼ぶ）=====
// 既存の検索アクション（GET /exec）に action と参加者管理の項目を足し、1つのアクションで全部を扱う。

function addAttendanceParams_(get) {
  const q = (name, desc, extra) => Object.assign({ name: name, in: 'query', schema: { type: 'string' }, description: desc }, extra || {});
  get.operationId = 'seminarApi';
  get.summary = '全厚済セミナー一覧の検索・PDF、参加者の登録・変更・名簿・履歴・連絡リスト・確認待ち候補';
  get.parameters = get.parameters.concat([
    q('action', 'apply=参加者の追加・不参加・再参加（まとめて可）/ roster_sync=名簿と照合 / changes=変更履歴 / person_rename=氏名の訂正 / search=一覧検索(既定) / register=参加登録 / update=変更 / roster=セミナー別名簿 / history=人別履歴 / followups=連絡リスト / people=人物検索 / candidate_add=確認待ち候補を保存 / candidates=候補一覧 / candidate_confirm=候補の承認・却下 / memos=ドライブのボイスメモ・PLAUDを読む / memo_done=資料を処理済みにする',
      { schema: { type: 'string', enum: ['search', 'register', 'update', 'roster', 'history', 'followups', 'people', 'candidate_add', 'candidates', 'candidate_confirm', 'memos', 'memo_done',
        'apply', 'roster_sync', 'changes', 'person_rename'] } }),
    q('seminar_id', 'セミナーID（検索結果や前の応答の seminar_id。分かっていれば優先）'), q('date', '開催日 yyyy-mm-dd（参加者管理でセミナーを探す時）'),
    q('start', '開始時刻 hh:mm'), q('name', '氏名（さん付け不要）'), q('person_id', '人物ID（同名で選び直す時）'),
    q('status', '状態: ' + STATUSES.join('/') + '（history・followups では絞り込み、カンマ区切り可）'),
    q('new_person', '同名の別人として新しく登録する時だけ true'), q('kana', '読み'), q('person_note', '同名を区別するメモ'),
    q('invited_on', '案内日 yyyy-mm-dd'), q('channel', '連絡手段 LINE/電話'), q('reply', '返事の内容'),
    q('next_contact', '次回連絡日 yyyy-mm-dd。update で消す時は clear'), q('memo', 'メモ（update では追記）'), q('memo_mode', 'replace でメモを上書き'),
    q('record_id', '記録ID（確認で選ばれた時）'), q('follow', 'フォロー状況 未フォロー/済み'),
    q('kind', 'followups の種類: waiting=返事待ち, today=今日（期限切れ含む）連絡する人, not_followed=参加後未フォロー, status=状態で絞る',
      { schema: { type: 'string', enum: ['waiting', 'today', 'not_followed', 'status'] } }),
    q('source_type', '候補の入口 PLAUD/ボイスメモ/LINE画像/画像メモ/貼り付け'),
    q('source_ref', '候補の資料名（録音タイトル・ファイル名・日時）。同じ資料の再処理では同じ値'),
    q('evidence', '候補の根拠となる発言・文面（短く）'), q('seminar_text', 'セミナーが決まらない時の手がかり'),
    q('state', 'candidates の絞り込み 確認待ち/登録済み/却下/all'), q('candidate_id', '候補ID（カンマ区切りで複数可）'),
    q('decision', 'candidate_confirm の approve か reject', { schema: { type: 'string', enum: ['approve', 'reject'] } }),
    q('source', 'memos の対象 voicememo/plaud/all（既定 all）', { schema: { type: 'string', enum: ['all', 'voicememo', 'plaud'] } }),
    q('days', 'memos で何日前までを見るか（既定3、今日だけなら1）'), q('limit', 'memos で読む件数（既定3、最大5）'),
    q('keyword', 'memos のタイトル絞り込み（人名・話題など）'), q('include_processed', 'memos で処理済みも含める時 true'),
    q('count', 'memo_done の候補数'),
    q('join', 'apply：参加予定にする人（追加・再参加）。「、」区切りで複数可'),
    q('leave', 'apply：不参加にする人（来られない・外す・キャンセル）。「、」区切り'),
    q('undecided', 'apply：参加か不参加か判断できない人（確認待ちに入る）。「、」区切り'),
    q('names', 'roster_sync：名簿に載っている全員。「、」区切り'),
    q('complete', 'roster_sync：最新の完全な名簿だとJPが確認した時だけ true（名簿に無い参加予定者を不参加にする）',
      { schema: { type: 'string', enum: ['true', 'false'] } }),
    q('reason', '変更理由（例：LINE「仕事で行けなくなりました」）。履歴に残る'),
    q('force', '確認済みで大量の不参加や同名への訂正を実行する時だけ true'),
    q('new_name', 'person_rename の正しい氏名'),
  ]);
}

function attendanceInstructions_() {
  return [
    '',
    '【最優先ルール】',
    '・参加者の登録・変更・確認を頼まれたら、必ず seminarApi を呼ぶ。自分の記憶や会話だけで「登録しました」と言わない。',
    '・seminarApi がこの会話で使えない時だけ「このチャットには台帳の登録機能（seminarApi）が付いていません。セミナー用GPTの文字入力（マイクの音声入力も可）で話しかけてください」と答える。エラーが返った時は「呼び出せない」と言わず、返ってきた message をそのまま伝える。',
    '・「今日」「明日」「今週土曜」なども必ず今日を基準に yyyy-mm-dd に直してから呼ぶ。',
    '【参加者管理】すべて seminarApi の action で切り替える（一覧検索は action=search か省略）。',
    '・登録「11月1日 大阪 青木講師のPBSに田中さんを参加予定」→ action=register, date, lecturer, pref, type, name, status。セミナーは一覧から自動で探すので、セミナー情報を手入力させない。',
    '・講師名は姓だけでよい（「又吉さん」→ lecturer=又吉）。市名（加古川・姫路など）は都道府県にして pref=兵庫 とし、venue には入れない。',
    '・複数人「今日の又吉さんのBEに牛谷、山下、玉田が参加します」→ name は1人ずつ。1人目を register で呼び、返ってきた record.seminar_id を2人目以降に使って順に呼ぶ（name に複数人を入れない）。状態を言われなければ参加予定。',
    '・needs_confirmation が返ったら choices を短く読み上げて選んでもらい、選ばれた seminar_id / person_id / record_id で呼び直す。勝手に選ばない。',
    '・seminar_not_found が返ったら推測で別の回に登録しない。同じ講師・種別で action=search を呼んで近い回を示し、「その日の回は一覧にまだ入っていません（スクショ未取り込み）」と伝える。JPが「とりあえず残して」と言ったら、1人ずつ action=candidate_add（source_type=貼り付け, source_ref=今日の日付＋講師＋種別, seminar_text=日付・地域・講師・種別, status）で確認待ちに保存し、回が一覧に入った後で candidate_confirm で承認する。',
    '・duplicate=true は「登録済み」と伝える（重複登録しない）。new_person=true で新しい人を作った時は「新規登録しました」と伝える。',
    '・登録結果の報告は、セミナーを1行（開催日・時間・講師・種別・地域・会場）、その下に「氏名｜新規登録/登録済み/要確認/未登録｜状態」を1人1行、最後に件数。',
    '【参加者の増減（最優先でこちらを使う）】',
    '・「山田さんを追加して」「佐藤さんがやっぱり参加する」→ action=apply, join=山田。「田中さんが来られなくなった」「山下さんを外して」「キャンセル」→ action=apply, leave=田中。複数人・混在は1回で join と leave に分けて入れる（「、」区切り）。',
    '・セミナーは直前の話題の seminar_id を付ける。不明なら付けずに呼ぶ（leave だけなら本人の今後の予定から、join は直近24時間に扱ったセミナーで処理し、seminar に結果が出る）。seminar_missing なら日付・講師を聞く。',
    '・LINEの文章が貼られて「この内容で更新」→ 文面から人ごとに参加/不参加/判断できないを分け、apply（join/leave/undecided、source_type=LINE、source_ref=送信者＋日時、reason=該当の一文）で1回呼ぶ。迷う人は必ず undecided。',
    '・名簿が送られたら action=roster_sync, names=全員。JPが「最新の全員の名簿」と確認した時だけ complete=true。それ以外は complete=false（名簿に無いだけの人を不参加にしない）。mass_leave が返ったら would_leave を読み上げて確認し、承認後に force=true で呼び直す。',
    '・報告は必ず返答の値で：セミナー1行、追加○人・不参加○人・再参加○人（名前付き）、現在の参加予定 counts.attending_now 人、確認待ち counts.pending_now 人、同期結果（sync.ok=true なら「保存を確認済み」、false なら errors を伝えて同じ内容で1回だけ呼び直す）。unchanged・not_registered・needs_choice があれば一言添え、needs_choice は choices から選んでもらう。',
    '・busy が返ったら数秒おいて同じ内容で1回呼び直す。不参加にしても記録は消えず、変更履歴は action=changes で見られる。名前の誤り（例「P山」→山下）は action=person_rename, name, new_name で直す。',
    '・変更「田中さんを参加済みに」→ action=update, name, status。「次は11月5日に連絡」→ next_contact。',
    '・名簿「このセミナーの参加者」→ action=roster。直前に話題にしたセミナーの seminar_id を使う。表示は見出し（開催日・時間・講師・種別・地域・会場）の下に参加者を1人1行で、状態を併記。PDFと言われたら format=pdf。',
    '・人別「田中さんが参加予定のセミナー」→ action=history, name, status=参加予定。',
    '・「返事待ち」→ action=followups, kind=waiting。「今日連絡する人」→ kind=today。「参加後まだフォローしていない人」→ kind=not_followed。',
    '・状態は 案内予定(誘いたい)/案内済み/返事待ち/参加予定/未定/辞退/キャンセル/参加済み/フォロー済み。日付は今日を基準に yyyy-mm-dd に直す（月日だけなら今日以降で一番近い日）。',
    '【録音・文字起こし・画像から】',
    '・PLAUDやボイスメモの文字起こし（貼り付け・添付）、LINEのスクショ、画像メモから人とセミナーを拾う時は、register を使わず、1人ずつ action=candidate_add で確認待ちに保存する。source_ref は資料のタイトルやファイル名・日時で、同じ資料なら毎回同じ値にする。',
    '・「誘いたい」は案内予定、「行けたら」「検討」は未定、「行けない」は辞退。はっきり「参加します」でない限り参加予定にしない。',
    '・読めない名前は 判読不可、不明な日付は推測せずに seminar_text に手がかりだけ書く。',
    '・保存後、候補を「名前｜セミナー｜提案状態｜根拠」で一覧にして確認を求める。本人が承認したものだけ action=candidate_confirm, decision=approve、違うものは reject。訂正があれば name/seminar_id/status を付けて承認。',
    '【Googleドライブのボイスメモ・PLAUD】',
    '・「今日のボイスメモから候補を拾って」「最近のPLAUDから」→ action=memos（source=voicememo/plaud、今日だけなら days=1、人名や話題は keyword）。返った items の text からセミナーに関係する人・参加意思・次回連絡を拾い、source_type と source_ref は返った値をそのまま使って candidate_add する。',
    '・1件読み終えたら候補が0件でも action=memo_done, source_ref, count で処理済みにする。処理済みは次から出ない（見直す時は include_processed=true）。unprocessed が returned より多い時は残りがあると伝える。',
    '・ノートの「聞き取り要確認」の名前は推測で直さず、そのまま候補にして確認を求める。セミナーに関係する人がいなければ「候補なし」と伝える。',
  ].join('\n');
}
