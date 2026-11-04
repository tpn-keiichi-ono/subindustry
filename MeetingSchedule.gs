/**
 * 会議の日程（Google カレンダー）と、会議資料を作る画面
 * 会議サマリ（MeetingSummary.gs。RX Business Meeting）・RX Service Meeting の事前資料（ServiceMeeting.gs）と同じプロジェクトに置くファイル。
 *
 * 仕組み
 * - 管理者がエディタで importMeetingSchedules() を実行すると、Google カレンダー（MSCH_OPTIONS.calendarId。空なら実行する人の既定のカレンダー）の予定から、
 *   RX Business Meeting・RX Service Meeting の予定（予定の名前に MSCH_OPTIONS.titles の言葉を含むもの）を探し、会議資料のインデックス
 *   （MTG_OPTIONS.index.sheet）に、会議日と会議名の行を先に足す（ある行はそのまま。資料ができたら、その行のセルに書く）。
 * - 資料を作る関数（createBusinessMtgSummaries()・previewBusinessMtgSummaries()・createBusinessMtgCombinedSummary()・
 *   createServiceMtgSummaries()・previewServiceMtgSummaries()）は、画面（MeetingRunDialog）を開く。画面で、インデックスにある会議日程を
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
  dialogSize: {width: 640, height: 600}
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

/** 会議資料を作る画面の種類（kind）：'business' 会議サマリ・'combined' 統合版・'service' RX Service Meeting。それ以外は止める */
function mschCheckKind_(kind) {
  if (['business', 'combined', 'service'].indexOf(kind) < 0) throw new Error('資料の種類「' + kind + '」が分かりません。');
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
  if (kind === 'business') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getDocumentProperties();
    sheets = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name)).map(name => {
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
  const titles = {business: '会議サマリ（' + meetingName + '）', combined: '統合版（' + MTG_OPTIONS.combined.name + '）', service: meetingName + ' の事前資料'};
  return {kind, final: !!final, title: titles[kind] + (final ? 'を作る' : 'を試しに作る'), meetingName,
    color: service ? SMTG_OPTIONS.colors.green : MTG_OPTIONS.colors.cover, meetings, selected, sheets, points, canPickTime: !service, job};
}

/**
 * 手順ごとの時間を測る（lap(名前) でそこまでの時間を記録、text() で「名前 1.2秒・…（合計 3.4秒）」）。
 * name を渡すと、始めたときと手順が終わるたびにすぐ実行ログに書く（途中で止まっても、どこまで進んだかが実行数のログで分かる）
 */
function mschTimer_(name) {
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

/** 画面の「カレンダーから読み込む」（google.script.run から呼ぶ）。{message, data: getMeetingRunData と同じ形} */
function importMeetingSchedulesFromDialog(kind, final) {
  const result = mschImport_();
  return {message: result.message, data: getMeetingRunData(kind, final)};
}

/** 画面で選んだ会議日・シートを確かめる（google.script.run から呼ぶ）。会議日の形・シートの名前が正しくなければ止める。戻り値は会議日（yyyy/MM/dd） */
function mschCheckInput_(kind, meeting, names) {
  const date = mtgParseDate_(String(meeting || ''), new Date());
  if (!date) throw new Error('会議日を選んでください。');
  if (kind === 'combined') return date;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allowed = kind === 'service' ? svcSourceSheets_().filter(name => ss.getSheetByName(name)).concat([MTG_OPTIONS.combined.name])
    : Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
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
  return {text: mtgRunCheckText_(!!final, names, date, resolved)};
}

/** 資料を作る（google.script.run から呼ぶ。画面の「作成する」）。cancelJob なら作っている途中の仕事をやめて作る。{message: 終わりの知らせ} */
function runMeetingMaterials(kind, final, meeting, names, cancelJob, point) {
  mschCheckKind_(kind);
  const date = mschCheckInput_(kind, meeting, names);
  if (kind === 'service') return {message: smtgRunStart_(!!final, date, names, !!cancelJob, smtgCheckPoint_(point))};
  const resolved = mtgResolvePoint_(point);
  if (kind === 'combined') return {message: mtgCombinedStart_(date, resolved)};
  return {message: mtgRunStart_(!!final, date, names, !!cancelJob, resolved)};
}
