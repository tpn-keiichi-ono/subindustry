/**
 * 会議の日程（Google カレンダー）と、会議資料を作る画面
 * 会議サマリ（MeetingSummary.gs。RX Business Meeting）・RX Service Meeting の事前資料（ServiceMeeting.gs）と同じプロジェクトに置くファイル。
 *
 * 仕組み
 * - 管理者がエディタで importMeetingSchedules() を実行すると、Google カレンダー（MSCH_OPTIONS.calendarId。空なら実行する人の既定のカレンダー）の予定から、
 *   RX Business Meeting・RX Service Meeting の予定（予定の名前に MSCH_OPTIONS.titles の言葉を含むもの）を探し、会議資料のインデックス
 *   （MTG_OPTIONS.index.sheet）に、会議日と会議名の行を先に足す（ある行はそのまま。資料ができたら、その行のセルに書く）。
 * - 資料を作る関数（createBusinessMtgSummaries()・previewBusinessMtgSummaries()・createBusinessMtgCombinedSummary()・
 *   createBusinessMtgReportingSummary()・createServiceMtgSummaries()・previewServiceMtgSummaries()）は、画面（MeetingRunDialog）を開く。画面で、インデックスにある会議日程を
 *   プルダウンで選び、作るシート（サブインダストリー）を選んで、確かめてから作る。前回の会議からの変化を出す（前回の資料の記録が無ければ、
 *   一覧の前の会議日から。RX Business Meeting）。比べる時点（保存先のフォルダの時点コピー・会議日・日時。RX Service Meeting は記録した会議）も選べる。
 *   画面の「カレンダーから読み込む」でも importMeetingSchedules() と同じことをする。
 * - 画面から呼ぶ関数（getMeetingRunData など）は末尾に _ を付けない（google.script.run から呼ぶため）。
 */

const MSCH_OPTIONS = {
  calendarId: '',          // 読む Google カレンダーの ID（空なら、実行する人の既定のカレンダー。招待された予定も入る）
  pastDays: 365,           // 今日から何日前までの予定を読むか
  aheadDays: 400,          // 今日から何日後までの予定を読むか
  // 会議ごとの、カレンダーの予定の名前に含まれる言葉（全角・半角・大文字小文字・空白の違いは無視）。どれかを含む予定を、その会議とみなす
  titles: {
    business: ['RX Business Meeting', 'RX Business MTG', 'RXビジネスMTG', 'RXビジネスミーティング', 'パイプライン会議'],
    service: ['RX Service Meeting', 'RX Service MTG', 'RXサービスMTG', 'RXサービスミーティング']
  },
  excludeWords: ['中止', 'キャンセル'],   // 予定の名前にこの言葉を含む予定は読まない（会議が無くなったとき）
  dialogTemplate: 'MeetingRunDialog',
  dialogSize: {width: 640, height: 600},
  // 画面の「作成する」の状態（作っている途中・結果）を、作った人のユーザーのプロパティに控える。画面が読み込み直されて初めの画面に戻っても、
  // 作っている途中なら待って結果を出し、終わっていれば結果を出す（同じ資料を2回作らないように）
  runStateKey: 'MSCH_RUN_STATE',
  runningMinutes: 7,       // 作り始めてからこの分数を過ぎても「作っている途中」のままなら、止まったとみなす（1回の実行は6分まで）
  resultMinutes: 30,       // 作り終えてからこの分数のあいだは、画面を開くと結果を出す
  pollSeconds: 10          // 作っている途中のとき、画面が結果を問い合わせる間隔
};

/* ---------------- エディタから実行する ---------------- */

/**
 * Google カレンダーから RX Business Meeting・RX Service Meeting の日程を読み、「INDEX_会議資料」に会議日と会議名の行を足す（エディタから実行）。
 * 何度実行してもよい（ある行はそのまま。カレンダーで日を変えた会議は、新しい日の行を足し、前の日の行は残す（資料が無ければ消してかまいません））
 */
function importMeetingSchedules() {
  const result = mschImport_();
  console.log(result.message);
  const ui = mtgUiOrNull_();
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, result.message, ui.ButtonSet.OK);
}

/* ---------------- 読む ---------------- */

/** 会議（key：'business'・'service'）の名前（インデックスの会議名の列に書く値） */
function mschKindName_(key) {
  return key === 'service' ? SMTG_OPTIONS.index.kind : MTG_OPTIONS.index.kind;
}

/** 予定の名前の、会議の見分け（全角・半角・大文字小文字・空白の違いは無視）。当てはまらなければ空 */
function mschKeyOfTitle_(title) {
  const norm = v => String(v || '').normalize('NFKC').toLowerCase().replace(/\s/g, '');
  const t = norm(title);
  if (MSCH_OPTIONS.excludeWords.some(w => t.indexOf(norm(w)) >= 0)) return '';
  return ['service', 'business'].find(key => MSCH_OPTIONS.titles[key].some(w => t.indexOf(norm(w)) >= 0)) || '';
}

/** 読むカレンダー（MSCH_OPTIONS.calendarId。開けなければ止める） */
function mschCalendar_() {
  const id = MSCH_OPTIONS.calendarId;
  const calendar = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
  if (!calendar) {
    throw new Error('Google カレンダー（ID：' + id + '）を開けません。実行するアカウントがそのカレンダーを見られるか、MeetingSchedule.gs の MSCH_OPTIONS.calendarId を確かめてください。');
  }
  return calendar;
}

/**
 * カレンダーの、RX Business Meeting・RX Service Meeting の予定（読むだけ）。断った予定・名前に excludeWords を含む予定は読まない。
 * [{key, kind: 会議名, date: 'yyyy/MM/dd', title, start}]（会議日の順。同じ会議・同じ日は1つ）
 */
function mschReadCalendar_(now) {
  const o = MSCH_OPTIONS;
  const tz = Session.getScriptTimeZone();
  const start = new Date(now.getTime() - o.pastDays * 86400000);
  const end = new Date(now.getTime() + o.aheadDays * 86400000);
  const seen = new Set();
  const out = [];
  mschCalendar_().getEvents(start, end).forEach(event => {
    if (event.getMyStatus && event.getMyStatus() === CalendarApp.GuestStatus.NO) return;
    const title = String(event.getTitle() || '');
    const key = mschKeyOfTitle_(title);
    if (!key) return;
    const at = event.getStartTime();
    const date = Utilities.formatDate(at, tz, 'yyyy/MM/dd');
    if (seen.has(key + date)) return;
    seen.add(key + date);
    out.push({key, kind: mschKindName_(key), date, title, start: at});
  });
  return out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}

/**
 * カレンダーの日程をインデックスに足す。{message, added, found: {business, service}}（知らせる文。書けなかったときも止めずに文で知らせる）
 */
function mschImport_() {
  const now = new Date();
  const events = mschReadCalendar_(now);
  const count = key => events.filter(e => e.key === key).length;
  const calendar = mschCalendar_().getName();
  const head = 'Google カレンダー「' + calendar + '」の ' + MSCH_OPTIONS.pastDays + '日前〜' + MSCH_OPTIONS.aheadDays + '日後の予定から、' +
    mschKindName_('business') + ' ' + count('business') + '件・' + mschKindName_('service') + ' ' + count('service') + '件の日程を読みました。';
  if (!events.length) {
    return {message: head + '\n予定の名前に、MeetingSchedule.gs の MSCH_OPTIONS.titles の言葉（「' + MSCH_OPTIONS.titles.business[0] + '」など）が入っているか確かめてください。',
      added: 0, found: {business: 0, service: 0}};
  }
  const result = mtgIndexEnsureRows_(events.map(e => ({date: e.date, kind: e.kind})));
  if (result.error) return {message: head + '\n' + result.error, added: 0, found: {business: count('business'), service: count('service')}};
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const next = key => events.filter(e => e.key === key && e.date >= today)[0];
  const lines = ['business', 'service'].map(key => {
    const n = next(key);
    return '・' + mschKindName_(key) + '：次回 ' + (n ? mtgDateLabel_(n.date) + '（' + n.title + '）' : 'なし');
  });
  return {message: head + '\n「' + MTG_OPTIONS.index.sheet + '」に ' + result.added + '行を足しました（すでにある ' + result.existing + '行はそのまま）：' + result.url +
    '\n\n' + lines.join('\n'), added: result.added, found: {business: count('business'), service: count('service')}};
}

/* ---------------- 資料を作る画面 ---------------- */

/**
 * 会議資料を作る画面の種類（kind）：'business' 会議サマリ・'report' 会議サマリの報告案件版（MeetingReport.gs。画面と作り方は business と同じ）・
 * 'combined' 統合版・'compact' 統合版（報告用。MeetingCompact.gs。画面と作り方は combined と同じ）・'service' RX Service Meeting。それ以外は止める
 */
function mschCheckKind_(kind) {
  if (['business', 'report', 'combined', 'compact', 'service'].indexOf(kind) < 0) throw new Error('資料の種類「' + kind + '」が分かりません。');
  return kind;
}

/** 会議資料を作る画面（MeetingRunDialog）を開く（エディタから実行する資料作成の関数から。スプレッドシートの画面に出す） */
function mschOpenRunDialog_(kind, final) {
  const ui = mtgUi_();
  mtgRequireSlides_();
  const o = MSCH_OPTIONS;
  const template = HtmlService.createTemplateFromFile(o.dialogTemplate);
  template.kind = mschCheckKind_(kind);
  template.final = !!final;
  ui.showModalDialog(template.evaluate().setWidth(o.dialogSize.width).setHeight(o.dialogSize.height), ' ');
}

/**
 * 画面が読む内容（google.script.run から呼ぶ）。
 * {kind, final, title, meetingName, color, meetings: [{date, label, filled, past}], selected: 初めに選ぶ会議日, sheets: [{name, checked, note}],
 *  points: 比べる時点の候補 [{key, label, day, group: プルダウンの見出し（無ければ見出しなし）}]（空の key は前回の会議。画面で、選んだ会議日より前の
 *  day のものだけ出す）, canPickTime: 日時を指定できるか,
 *  job: 途中の仕事の知らせ}
 */
function getMeetingRunData(kind, final) {
  mschCheckKind_(kind);
  // どこで時間がかかったかを実行ログに残す（画面が「読み込んでいます…」から進まないときに、実行数のログで確かめる）
  const timer = mschTimer_('資料を作る画面の読み込み');
  const service = kind === 'service';
  const meetingName = mschKindName_(service ? 'service' : 'business');
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const meetings = mtgIndexMeetings_(meetingName).map(m => ({date: m.date, label: mtgDateLabel_(m.date), filled: m.filled, past: m.date < today}));
  timer.lap('会議日程（' + MTG_OPTIONS.index.sheet + '）');
  // 初めに選ぶ会議日：今日以降でいちばん近い会議日（無ければ、いちばん新しい会議日）
  const upcoming = meetings.filter(m => !m.past)[0];
  const selected = upcoming ? upcoming.date : meetings.length ? meetings[meetings.length - 1].date : '';
  let sheets = [];
  let points = [];
  let job = '';
  if (kind === 'business' || kind === 'report') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getDocumentProperties();
    sheets = mtgTrackedNames_(ss).map(name => {
      const last = mtgMeetings_(props, name).filter(m => !m.baseline).pop();
      return {name, checked: true, note: last ? '前回の資料 ' + mtgShortDate_(last.date) : '前回の資料なし'};
    });
    job = mtgJobPendingText_();
    timer.lap('シートと前回の資料');
  } else if (service) {
    const data = smtgRunData_();
    sheets = data.sheets;
    points = data.points;
    job = data.job;
    timer.lap('RX Service Meeting の対象');
  }
  if (!service) {
    points = mtgComparePoints_(SpreadsheetApp.getActiveSpreadsheet()).map(p => ({key: p.key, label: p.label, day: p.day, group: p.group}));
    timer.lap('比べる時点（' + MTG_OPTIONS.copyName + ' ' + points.length + '件）');
  }
  console.log('資料を作る画面の読み込み：' + timer.text());
  const titles = {business: '会議サマリ（' + meetingName + '）', report: '会議サマリ・報告案件版（' + meetingName + '）', combined: '統合版（' + MTG_OPTIONS.combined.name + '）',
    compact: mtgCompactLabel_(), service: meetingName + ' の事前資料'};
  // 直前の「作成する」（画面が読み込み直されたとき・開き直したときに、作っている途中なら待ち、終わっていれば結果を出す）
  const run = getMeetingRunState(kind);
  return {kind, final: !!final, title: titles[kind] + (final ? 'を作る' : 'を試しに作る'), meetingName,
    color: service ? SMTG_OPTIONS.colors.green : MTG_OPTIONS.colors.cover, meetings, selected, sheets, points, canPickTime: !service, job,
    run, pollSeconds: MSCH_OPTIONS.pollSeconds};
}

/**
 * 手順ごとの時間を測る（lap(名前) でそこまでの時間を記録、text() で「名前 1.2秒・…（合計 3.4秒）」）。
 * name を渡すと、始めたときと手順が終わるたびにすぐ実行ログに書く（途中で止まっても、どこまで進んだかが実行数のログで分かる）。
 * progress なら、「作成する」の控え（mschRunState_）にも書く（mschRunProgress_）
 */
function mschTimer_(name, progress) {
  const began = Date.now();
  let last = began;
  const laps = [];
  const sec = ms => (Math.round(ms / 100) / 10) + '秒';
  if (name) console.log('[' + name + '] 開始');
  return {
    lap: label => {
      const now = Date.now();
      laps.push(label + ' ' + sec(now - last));
      if (name) console.log('[' + name + '] ' + label + '：済み（' + sec(now - last) + '）');
      last = now;
      // progress：「作成する」の控えにも、済んだ手順を書く（6分の上限で止められても、画面にどこまで進んだかを出せるように）
      if (progress) mschRunProgress_(name, laps.join('・'));
    },
    text: () => laps.join('・') + '（合計 ' + sec(Date.now() - began) + '）',
    laps
  };
}

/**
 * 資料を作る画面（MeetingRunDialog）が「読み込んでいます…」から進まないときの確かめ（エディタから実行する。読むだけ）。
 * 画面が読み込むものを同じ順に読み、手順ごとの時間と件数を出す（時点コピーの保存先のフォルダ・ファイルの数も）
 */
function diagnoseMeetingRunDialog() {
  const lines = [];
  const timer = mschTimer_();
  try {
    const data = getMeetingRunData('business', false);
    lines.push('読み込み：' + timer.text());
    lines.push('会議日程：' + data.meetings.length + '件・シート：' + data.sheets.length + '・比べる時点：' + data.points.length + '件' + (data.job ? '・途中の仕事あり' : ''));
  } catch (error) {
    lines.push('✗ 読み込めませんでした：' + (error && error.message || error) + '（' + timer.text() + '）');
  }
  const t2 = mschTimer_('時点コピーの保存先');
  const folder = mtgCopyFolder_(SpreadsheetApp.getActiveSpreadsheet(), true);
  t2.lap('時点コピーの保存先を開く');
  let count = 0;
  if (folder) {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) { files.next(); count++; }
    t2.lap('ファイルを数える');
  }
  lines.push('時点コピーの保存先：' + (folder ? folder.getName() + '（スプレッドシート ' + count + '件）' : 'なし') + '　' + t2.text());
  svcShowLines_(['資料を作る画面の読み込みの確かめ'].concat(lines));
}

/**
 * 資料の中身の読み方ごとの時間を比べる（エディタから実行する。読むだけ。大勢が編集している時間帯に実行する）。
 * ① このスプレッドシートを Sheets API（拡張サービス）で読む ② このスプレッドシートを SpreadsheetApp で読む（前の版の読み方）
 * ③ シートの値だけを CSV で書き出して読む（スプレッドシートの書き出し。数式の計算し直しが無い）。
 * 読むのは、いちばん初めの追跡シート（見出しの行から最後の行まで）と、変更履歴_差分の最後の 5,000行。
 * 作り始めたときのコピーは、作ったばかりのものを初めて開くと数式の計算し直しで数分かかったため（2026/11/04 に 274秒）、コピーを作らずに読めるかを確かめる
 */
function diagnoseMeetingRead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = Object.keys(DIFF_RULES).find(n => ss.getSheetByName(n));
  const lines = [];
  if (!name) {
    svcShowLines_(['資料の読み方の確かめ', '追跡シートがありません。']);
    return;
  }
  const quote = n => "'" + n.replace(/'/g, "''") + "'";
  const logRows = 5000;
  // ① Sheets API
  const t1 = mschTimer_('読み方の確かめ：Sheets API');
  try {
    if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets) throw new Error('Sheets API（拡張サービス）が使えません');
    const meta = Sheets.Spreadsheets.get(ss.getId(), {fields: 'sheets.properties(title,gridProperties(rowCount,columnCount))'});
    const size = title => ((meta.sheets || []).find(x => x.properties.title === title) || {properties: {gridProperties: {}}}).properties.gridProperties;
    t1.lap('シートの大きさ');
    const deals = Sheets.Spreadsheets.Values.batchGet(ss.getId(), {ranges: [quote(name) + '!A' + DIFF_RULES[name].headerRow + ':' + diffColumnLetter_(size(name).columnCount || 26)],
      valueRenderOption: 'FORMATTED_VALUE'});
    const dealRows = ((deals.valueRanges || [])[0].values || []).length;
    t1.lap('「' + name + '」' + dealRows + '行');
    // 変更履歴_差分は全行（シートの行数は、記録の無い行も数えるため、最後の何行かを指定すると空になることがある）
    const log = Sheets.Spreadsheets.Values.batchGet(ss.getId(), {ranges: [quote(DIFF_OPTIONS.logSheet) + '!A1:I'], valueRenderOption: 'UNFORMATTED_VALUE'});
    t1.lap('「' + DIFF_OPTIONS.logSheet + '」' + (((log.valueRanges || [])[0] || {}).values || []).length + '行');
    lines.push('① Sheets API：' + t1.text());
  } catch (error) {
    lines.push('① Sheets API：✗ ' + (error && error.message || error) + '（' + t1.text() + '）');
  }
  // ② SpreadsheetApp（前の版の読み方）
  const t2 = mschTimer_('読み方の確かめ：SpreadsheetApp');
  try {
    const sheet = ss.getSheetByName(name);
    const rule = DIFF_RULES[name];
    const rows = Math.max(sheet.getLastRow() - rule.headerRow + 1, 1);
    sheet.getRange(rule.headerRow, 1, rows, sheet.getLastColumn()).getDisplayValues();
    t2.lap('「' + name + '」' + rows + '行');
    const log = ss.getSheetByName(DIFF_OPTIONS.logSheet);
    if (log) {
      const total = log.getLastRow();
      const from = Math.max(1, total - logRows + 1);
      log.getRange(from, 1, total - from + 1, log.getLastColumn()).getValues();
      t2.lap('「' + DIFF_OPTIONS.logSheet + '」' + (total - from + 1) + '行');
    }
    lines.push('② SpreadsheetApp：' + t2.text());
  } catch (error) {
    lines.push('② SpreadsheetApp：✗ ' + (error && error.message || error) + '（' + t2.text() + '）');
  }
  // ③ 値だけを CSV で書き出す（シートごと。表示の文字）
  const t3 = mschTimer_('読み方の確かめ：CSV');
  try {
    [name, DIFF_OPTIONS.logSheet].forEach(title => {
      const sheet = ss.getSheetByName(title);
      if (!sheet) return;
      const response = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=csv&gid=' + sheet.getSheetId(),
        {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
      if (response.getResponseCode() !== 200) throw new Error('「' + title + '」を書き出せません（' + response.getResponseCode() + '）');
      const rows = Utilities.parseCsv(response.getContentText()).length;
      t3.lap('「' + title + '」' + rows + '行');
    });
    lines.push('③ 値だけ（CSV）：' + t3.text());
  } catch (error) {
    lines.push('③ 値だけ（CSV）：✗ ' + (error && error.message || error) + '（' + t3.text() + '）');
  }
  svcShowLines_(['資料の読み方の確かめ（このスプレッドシートを、コピーを作らずに読む時間）'].concat(lines));
}

/**
 * スプレッドシートの計算を重くしている数式・書式を探す（エディタから実行する。読むだけ）。
 * diagnoseMeetingRead() で、225行のシートを読むだけでもどの読み方でも約100秒かかった（2026/11/04）。読み方ではなく、スプレッドシートの数式の計算を
 * 待たされている。どの数式が重いかを、計算の済んだいちばん新しい時点コピーで調べる（このスプレッドシートを読むと、それ自体を待たされるため）。
 * シートごとのセル・数式・条件付き書式の数と、重くなりやすい数式（ほかのファイルを読む・編集のたびに計算し直す・範囲の広い計算・列全体の参照・
 * 記録のシートを参照）の数と、例のセルを出す
 */
function diagnoseSpreadsheetLoad() {
  const timer = mschTimer_('計算の重さの確かめ');
  const copies = mtgListCopies_();
  let book = null;
  let where = '';
  for (let i = 0; i < copies.length && !book; i++) {
    try {
      book = SpreadsheetApp.openById(copies[i].id);
      where = '「' + copies[i].title + '」（いちばん新しい時点コピー。中身はこのスプレッドシートと同じ）';
    } catch (_) {}
  }
  if (!book) {
    book = SpreadsheetApp.getActiveSpreadsheet();
    where = 'このスプレッドシート（時点コピーが無いため。読むのに時間がかかります）';
  }
  timer.lap('開く');
  const kinds = [
    {label: 'ほかのファイルを読む（コピーでは 0・エラーになる）', names: ['IMPORTRANGE', 'IMPORTDATA', 'IMPORTXML', 'IMPORTHTML', 'IMPORTFEED', 'GOOGLEFINANCE', 'IMAGE']},
    {label: '編集のたびに計算し直す', names: ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'INDIRECT', 'OFFSET', 'CELL', 'INFO']},
    {label: '範囲の広い計算', names: ['ARRAYFORMULA', 'QUERY', 'FILTER', 'SORT', 'SORTN', 'UNIQUE', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'LOOKUP', 'MATCH', 'XMATCH',
      'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'SUMPRODUCT', 'MAP', 'BYROW', 'BYCOL', 'LAMBDA', 'REGEXMATCH', 'REGEXEXTRACT', 'SPLIT', 'TEXTJOIN']}
  ];
  const records = [DIFF_OPTIONS.logSheet, DIFF_OPTIONS.rowSnapshotSheet, DIFF_OPTIONS.snapshotPrefix];
  const found = {};   // 関数 → {count, sheets: {シート: 件数}, example}
  const note = (key, sheet, a1) => {
    const f = found[key] = found[key] || {count: 0, sheets: {}, example: ''};
    f.count++;
    f.sheets[sheet] = (f.sheets[sheet] || 0) + 1;
    if (!f.example) f.example = sheet + '!' + a1;
  };
  const sheets = [];
  let cells = 0;
  book.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const size = sheet.getMaxRows() * sheet.getMaxColumns();
    cells += size;
    const rows = sheet.getLastRow();
    const width = sheet.getLastColumn();
    let formulas = 0;
    const distinct = new Set();
    if (rows > 0 && width > 0) {
      sheet.getRange(1, 1, rows, width).getFormulas().forEach((line, r) => line.forEach((f, c) => {
        if (!f) return;
        formulas++;
        const text = String(f).toUpperCase();
        distinct.add(text.replace(/\d+/g, '#'));   // 同じ形の数式（行だけ違う）は1種類と数える
        const a1 = diffColumnLetter_(c + 1) + (r + 1);
        (text.match(/[A-Z][A-Z0-9.]*(?=\s*\()/g) || []).filter((fn, i, list) => list.indexOf(fn) === i)
          .forEach(fn => { if (kinds.some(k => k.names.indexOf(fn) >= 0)) note(fn, name, a1); });
        if (/(^|[^A-Z$])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![0-9A-Z])/.test(text)) note('列全体の参照（A:A など）', name, a1);
        if (records.some(n => String(f).indexOf(n) >= 0)) note('記録のシートを参照', name, a1);
      }));
    }
    let rules = 0;
    let custom = 0;
    try {
      const list = sheet.getConditionalFormatRules();
      rules = list.length;
      custom = list.filter(rule => { try { return !!rule.getBooleanCondition() && String(rule.getBooleanCondition().getCriteriaType()) === 'CUSTOM_FORMULA'; } catch (_) { return false; } }).length;
    } catch (_) {}
    sheets.push({name, size, rows, width, formulas, distinct: distinct.size, rules, custom});
  });
  timer.lap('数式と書式を数える');
  const lines = ['調べたファイル：' + where,
    'セルの数（空のセルを含む）：' + cells.toLocaleString() + '（上限 10,000,000）・シート ' + sheets.length + '枚', '', '▼ シートごと（数式の多い順）'];
  sheets.filter(x => x.formulas || x.rules).sort((a, b) => b.formulas - a.formulas || b.rules - a.rules).forEach(x => {
    lines.push('・' + x.name + '：数式 ' + x.formulas.toLocaleString() + '（形は ' + x.distinct + '種類）・セル ' + x.size.toLocaleString() +
      '（使っている範囲 ' + x.rows + '行×' + x.width + '列）' + (x.rules ? '・条件付き書式 ' + x.rules + '（数式 ' + x.custom + '）' : ''));
  });
  lines.push('', '▼ 重くなりやすい数式（件数・どのシートか・例のセル）');
  const show = key => {
    const f = found[key];
    if (!f || !f.count) return;
    const where2 = Object.keys(f.sheets).map(n => n + ' ' + f.sheets[n]).join('・');
    lines.push('・' + key + '：' + f.count.toLocaleString() + '件（' + where2 + '）　例 ' + f.example);
  };
  kinds.forEach(k => {
    if (!k.names.some(n => found[n] && found[n].count)) return;
    lines.push('〔' + k.label + '〕');
    k.names.forEach(show);
  });
  if (found['列全体の参照（A:A など）'] || found['記録のシートを参照']) lines.push('〔そのほか〕');
  show('列全体の参照（A:A など）');
  show('記録のシートを参照');
  lines.push('', '（' + timer.text() + '）');
  svcShowLines_(['スプレッドシートの計算の重さの確かめ'].concat(lines));
}

/** 画面の「カレンダーから読み込む」（google.script.run から呼ぶ）。{message, data: getMeetingRunData と同じ形} */
function importMeetingSchedulesFromDialog(kind, final) {
  const result = mschImport_();
  return {message: result.message, data: getMeetingRunData(kind, final)};
}

/** 画面で選んだ会議日・シートを確かめる（google.script.run から呼ぶ）。会議日の形・シートの名前が正しくなければ止める。戻り値は会議日（yyyy/MM/dd） */
function mschCheckInput_(kind, meeting, names) {
  const date = mtgParseDate_(String(meeting || ''), new Date());
  if (!date) throw new Error('会議日を選んでください。');
  if (kind === 'combined' || kind === 'compact') return date;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allowed = kind === 'service' ? svcSourceSheets_().filter(name => ss.getSheetByName(name)).concat([MTG_OPTIONS.combined.name])
    : mtgTrackedNames_(ss);
  if (!Array.isArray(names) || !names.length) throw new Error('作るシート（サブインダストリー）を1つ以上選んでください。');
  names.forEach(name => { if (allowed.indexOf(name) < 0) throw new Error('「' + name + '」は選べません。'); });
  return date;
}

/**
 * 作る前の確かめ（google.script.run から呼ぶ）。point は比べる時点（画面のプルダウンの key。空なら前回の会議）。
 * {text: 確かめの文} か {error: 作れない理由}
 */
function checkMeetingRun(kind, final, meeting, names, point) {
  mschCheckKind_(kind);
  const date = mschCheckInput_(kind, meeting, names);
  if (kind === 'service') return {text: smtgRunCheckText_(!!final, date, names, smtgCheckPoint_(point))};
  const resolved = mtgResolvePoint_(point);
  if (kind === 'combined') return mtgCombinedCheck_(date, resolved);
  if (kind === 'compact') return mtgCompactCheck_(date, resolved);
  const text = mtgRunCheckText_(!!final, names, date, resolved);
  return {text: kind === 'report' ? mtgReportCheckText_() + '\n\n' + text : text};
}

/**
 * 資料を作る（google.script.run から呼ぶ。画面の「作成する」）。cancelJob なら作っている途中の仕事をやめて作る。overwrite なら、前に作った同じ資料
 * （同じ会議日・同じ名前の PDF）の中身を書き換えて、リンクを変えない（mtgExistingPdf_・mtgOverwriteFile_）。{message: 終わりの知らせ}。
 * 作っている途中・結果を控える（mschRunState_）。同じ人が作っている途中なら作らない（画面が読み込み直されたあとに、もう一度押されても2回作らない）
 */
function runMeetingMaterials(kind, final, meeting, names, cancelJob, point, overwrite) {
  mschCheckKind_(kind);
  const date = mschCheckInput_(kind, meeting, names);
  const running = mschRunState_();
  if (running && running.status === 'running') {
    throw new Error(Utilities.formatDate(new Date(running.at), Session.getScriptTimeZone(), 'H:mm') +
      ' から資料を作っている途中です。終わるまでお待ちください（画面を開き直すと、作っている途中の知らせか結果が出ます）。');
  }
  const base = {kind, final: !!final, meeting: date, at: Date.now()};
  mschRunStateSave_(Object.assign({status: 'running'}, base));
  try {
    let message;
    if (kind === 'service') {
      message = smtgRunStart_(!!final, date, names, !!cancelJob, smtgCheckPoint_(point), !!overwrite);
    } else {
      const resolved = mtgResolvePoint_(point);
      message = kind === 'combined' ? mtgCombinedStart_(date, resolved, !!overwrite)
        : kind === 'compact' ? mtgCompactStart_(date, resolved, !!overwrite)
        : mtgRunStart_(!!final, date, names, !!cancelJob, resolved, !!overwrite, kind === 'report' ? 'report' : '');
    }
    mschRunStateSave_(Object.assign({status: 'done', message, finishedAt: Date.now()}, base));
    return {message};
  } catch (error) {
    const current = mschRunState_();
    mschRunStateSave_(Object.assign({status: 'error', message: String(error && error.message || error), finishedAt: Date.now(),
      progress: current && current.progress || null}, base));
    throw error;
  }
}

/**
 * 画面が結果を問い合わせる（google.script.run から呼ぶ。読み込み直された画面が、作っている途中のときに MSCH_OPTIONS.pollSeconds ごとに）。
 * その資料の種類の控え（mschRunState_）を {status, final, meeting, message, started: 作り始めた時刻「H:mm」, finished: 終えた時刻} にしたもの。無ければ null
 */
function getMeetingRunState(kind) {
  mschCheckKind_(kind);
  const state = mschRunState_();
  if (!state || state.kind !== kind) return null;
  const time = ms => ms ? Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'H:mm') : '';
  // 作れなかったときは、どこまで進んだか（済んだ手順と時間）も出す（6分の上限で止められたときは、mschRunState_ が文に足している）
  const message = (state.message || '') + (state.status === 'error' && !/止まる前に済んだ手順/.test(state.message || '') ? mschRunProgressText_(state) : '');
  return {status: state.status, final: !!state.final, meeting: state.meeting || '', message, started: time(state.at), finished: time(state.finishedAt)};
}

/**
 * 「作成する」の控え {status: 'running'|'done'|'error', kind, final, meeting, at: 作り始めた時刻, message, finishedAt}（作った人のユーザーのプロパティ）。
 * 作っている途中のまま runningMinutes を過ぎたものは error（1回の実行の上限で止まった）にし、終わってから resultMinutes を過ぎたものは null
 */
function mschRunState_() {
  const o = MSCH_OPTIONS;
  let state = null;
  try {
    const value = PropertiesService.getUserProperties().getProperty(o.runStateKey);
    state = value ? JSON.parse(value) : null;
  } catch (_) {
    return null;
  }
  if (!state || !state.at) return null;
  const now = Date.now();
  if (state.status === 'running') {
    if (now - state.at <= o.runningMinutes * 60000) return state;
    state = Object.assign({}, state, {status: 'error', finishedAt: state.at + o.runningMinutes * 60000,
      message: '資料を作っている途中で止まったようです（1回の実行の上限の6分を過ぎた など）。作られた資料は、出力先のフォルダで確かめてください。' +
        mschRunProgressText_(state)});
  }
  return now - (state.finishedAt || state.at) <= o.resultMinutes * 60000 ? state : null;
}

/** 作っている途中の「作成する」の控えに、済んだ手順と時間を書く（name ごとに上書き。作っていなければ何もしない） */
function mschRunProgress_(name, text) {
  try {
    const value = PropertiesService.getUserProperties().getProperty(MSCH_OPTIONS.runStateKey);
    const state = value ? JSON.parse(value) : null;
    if (!state || state.status !== 'running') return;
    state.progress = Object.assign({}, state.progress, {[name]: text});
    mschRunStateSave_(state);
  } catch (error) {
    console.warn('作成の進み具合を控えられませんでした：' + (error && error.message || error));
  }
}

/** 控えの済んだ手順と時間の文（無ければ空） */
function mschRunProgressText_(state) {
  const entries = Object.entries(state && state.progress || {});
  return entries.length ? '\n\n止まる前に済んだ手順と時間：\n' + entries.map(([name, text]) => '・' + name + '：' + text).join('\n') : '';
}

/** 「作成する」の控えを保存する（控えが書けなくても資料作りは止めない） */
function mschRunStateSave_(state) {
  try {
    PropertiesService.getUserProperties().setProperty(MSCH_OPTIONS.runStateKey, JSON.stringify(state));
  } catch (error) {
    console.warn('作成の状態を控えられませんでした：' + (error && error.message || error));
  }
}
