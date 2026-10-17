/**
 * RX Service Meeting の事前資料（サービス改善の会議。16:9 のスライドと PDF。紫を基調にしたトンマナ）
 * サービス管理（ServiceManagement.gs）・タスク管理（TaskManagement.gs）・会議サマリ（MeetingSummary.gs）と同じプロジェクトに置くファイル。
 *
 * 目的
 * - RX Service Meeting（サービス改善。月1回・第二週の金曜 9:00〜9:30）の事前資料。会議の目的（顧客課題と提供価値の仮説をすり合わせ、
 *   提案・実行のシナリオを具体化する）と議題（シナリオの有効性検証・顧客・業界との適合・マネタイズの検討・サービスの磨き込み）に向けて、
 *   サービス案ごとの状況を事前に共有する。
 * - 資料はサブインダストリー（リクエストを読むシート）ごとに作る（2026/10/17 から。RX Business Meeting と同じ）。そのサブインダストリーのリクエストと、
 *   それを含むサービス案を載せる（サービス案のタスク・サブインダストリーの広がりは、サービス案全体）。すべてのリクエストをまとめた資料
 *   （全サブインダストリー。MTG_OPTIONS.combined.name）も作れる。
 * - RX Business Meeting の資料と同じく、シートから分かる事実だけを書く（サービス案の良し悪し・決めてほしいことなどの判断は書かない）。
 *
 * 資料の構成（RX Business Meeting の資料と同じ形・同じ部品（MeetingSummary.gs の mtg*）で、色を紫（SMTG_OPTIONS.colors）にする）
 *   表紙（ひな形の3枚目＝紫。MTG_OPTIONS.templatePages.service。ひな形が無ければ同じ絵柄の紫の画像に描く）
 *   1 サマリー：数字（登録リクエスト・判断待ち・サービス化検討・サービス案・タスクの完了）と前回比、前回以降の動き、議題と資料のページ
 *   2 サービス案の一覧：リクエストの数・得意先・サブインダストリー・元になった案件の期待値・主管本部と担当者・タスクの状況
 *   3 顧客・業界の広がり：サービス案 × サブインダストリーのリクエストの件数
 *   4 サービス案（1つずつ）：元になったリクエスト（得意先・リクエスト・サービス部門からのフィードバック・案件のステータスと期待値）とタスク
 *   5 判断待ちのリクエスト：判断が「未判断」のリクエスト
 *   扉（Appendix。グレー）
 *   Appendix ① リクエスト（一覧）　② タスク（一覧）　③ 入力内容の確認事項（一覧）
 *
 * 読むもの（書き込まない）
 * - サービスリクエスト（SVC_OPTIONS.requestSheet）：登録したリクエスト。「サービス案」に同じ名前を入れたリクエストを1つのサービス案にまとめる
 *   （判断が「棄却」のリクエストは、サービス案に数えない）
 * - サービス案タスク管理（TASK_OPTIONS.sheet）：「サービス案」（「SR-0001 …」）の ID で、リクエスト → サービス案に結び付ける
 * - リクエストを読むシート（DIFF_RULES の requests: true）：得意先＋案件名の案件のステータス・期待値（案件全体の値。サービスの売上ではない）。
 *   リクエストがどのサブインダストリーのものかは、案件のあるシート（案件が無ければ、リクエストの「サブインダストリー」と同じ名前のシート）で決める
 *
 * 仕組み
 * - 管理者がエディタで createServiceMtgSummaries() を実行し、作るサブインダストリーと会議日を入力する（試すときは previewServiceMtgSummaries()。
 *   会議日は記録しない）。Google スライドで作り、PDF を会議サマリの保存先（MTG_OPTIONS.folderId）に保存する。配る（共有する）のは手動。
 *   1回の実行の6分に収まらなければ区切り、1回きりのトリガー（continueServiceMtgSummaries）で続ける（結果は showServiceMtgSummariesResult()）。
 * - 本番は、会議日とその時点のリクエストの判断・タスクの状況・サブインダストリーごとの件数をドキュメントのプロパティ（SMTG_OPTIONS.meetingsKey）に残し、
 *   次の会議の資料に「前回以降の動き」（新規リクエスト・判断の変更・完了したタスク・新しいタスク）と前回比を出す。
 *   本番の PDF は会議資料のインデックス（MTG_OPTIONS.index.sheet）の、種別が SMTG_OPTIONS.index.kind の行の、そのサブインダストリーの列に書く。
 * - データが無いうちは、サンプルデータ.gs の addServiceMtgSampleData() でサンプルのリクエスト・タスクを入れて試す
 *   （removeServiceMtgSampleData() で消す。サンプルを含む資料には【サンプルデータを含む】と出す）。
 */

const SMTG_OPTIONS = {
  meetingName: 'RX Service Meeting',          // 表紙・出典の行・会議のフォルダの名前に出す会議の名前
  formerMeetingNames: ['RX Service MTG'],      // 前の名前（2026/10/17 まで）。この名前の会議のフォルダも、この会議のフォルダとして使う
  // ファイル名（スライド・PDF）。{date} は会議日（yyyyMMdd）、{name} はサブインダストリー（全体は MTG_OPTIONS.combined.name）。試作は末尾に「_試作」
  fileName: '{date}_{name}_RX Service Meeting用資料',
  // 会議の記録（ドキュメントのプロパティ）：会議日の一覧と、会議日ごとのリクエストの判断・タスクの状況・件数（後ろに yyyyMMdd）
  meetingsKey: 'SMTG_MEETINGS',
  meetingDetailPrefix: 'SMTG_MEETING_',
  keepMeetings: 24,
  detailMaxLength: 8000,                  // 判断・状況の記録の長さの上限（プロパティの1つの値の上限 9KB に収める。越えたら件数だけ残す）
  // 会議資料のインデックス（MTG_OPTIONS.index）の種別の列に書く値（その行の、サブインダストリー・全サブインダストリーの列に PDF を書く）
  // formerKinds は前の版で書いた種別（同じ会議の行として使い、書くときに今の種別に書き換える）
  index: {kind: 'RX Service Meeting', formerKinds: ['RX Service MTG']},
  // Apps Script の1回の実行は6分まで。作り始めてから limitMs を過ぎそうなら区切り、continueAfterMs あとに続ける（トリガー handler。1回きり）。
  // 見積もり：readMs シートを読む・deckMs 資料1つ・storeMs 格納とインデックス1つ
  run: {limitMs: 320000, readMs: 20000, deckMs: 40000, storeMs: 10000, continueAfterMs: 60000,
    handler: 'continueServiceMtgSummaries', jobKey: 'SMTG_JOB', resultKey: 'SMTG_JOB_RESULT'},
  // 議題と、それを話すときに見るページ（サマリーの「議題と資料のページ」。pages は SMTG_TITLES のキー）
  agenda: [
    // note は改行で行を分けられる（2行まで）
    {label: 'シナリオの有効性検証', pages: ['services', 'detail'], note: '元になったリクエストと\nサービス部門のフィードバック'},
    {label: '顧客・業界との適合', pages: ['fit', 'detail'], note: 'サービス案ごとの得意先と\nサブインダストリー'},
    {label: 'マネタイズの検討', pages: ['services', 'detail'], note: '元になった案件のステータスと\n期待値（案件全体の値）'},
    {label: 'サービスの磨き込み', pages: ['detail', 'tasks'], note: 'タスクの状況・メモ'}
  ],
  fitColumns: 10,     // 顧客・業界の広がりに並べるサブインダストリーの数の上限（多ければリクエストの多い順。残りは「その他」にまとめる）
  movesMax: 6,        // サマリーの「前回以降の動き」に名前を出す数（残りは「他 N件」）
  // 色（紫を基調にしたトンマナ）。MTG_OPTIONS.colors の同じキーを置き換えて、会議サマリの部品で描く（キーの名前は green のままで、中身が紫）。
  // 表の見出し・帯は green、表の1列目は greenSoft、大きな数字は greenDark、ラベル（サービス化検討・完了）は greenDeep
  colors: {green: '#7030A0', greenDark: '#5B2483', greenDeep: '#7030A0', greenSoft: '#EFE6F6', greenText: '#3F1A5C', up: '#B58BD8',
    cover: '#7030A0'}
};

/* スライドの見出し（Appendix の見出しは「Appendix ①　…」の形にする。扉の一覧に出す。mtgDividerSlide_） */
const SMTG_TITLES = {
  cover: '表紙',
  summary: 'サマリー',
  services: 'サービス案の一覧',
  fit: '顧客・業界の広がり',
  detail: 'サービス案',
  pending: '判断待ちのリクエスト',
  appendix: 'Appendix',
  requests: 'Appendix ①　リクエスト（一覧）',
  tasks: 'Appendix ②　タスク（一覧）',
  checks: 'Appendix ③　入力内容の確認事項（一覧）'
};

/* 入力内容の確認事項の種類（この順に出す。意味は docs/DESIGN.md の「RX Service Meeting の事前資料」） */
const SMTG_CHECKS = ['サービス案未設定', '判断とサービス案の不整合', '選択肢外の判断', '主管本部・担当者未設定', 'タスクなし', 'タスクの担当者未入力',
  'タスクの状況未入力', 'リクエストが見つからない', '元の案件が見つからない'];

/* ---------------- エディタから実行する ---------------- */

/**
 * RX Service Meeting の事前資料（Google スライドと PDF）を作る（作るサブインダストリーと今回の会議日を聞く）。
 * 会議日とその時点の状況を記録し、次の資料は今からの変化を出す。
 */
function createServiceMtgSummaries() {
  mschOpenRunDialog_('service', true);
}

/** 試しに RX Service Meeting の事前資料を作る（保存するが、会議日は記録しない。ファイル名に「試作」が付く）。 */
function previewServiceMtgSummaries() {
  mschOpenRunDialog_('service', false);
}

/**
 * 画面（MeetingRunDialog）に出す、作る資料の候補：サブインダストリー（リクエストの件数。0件は選ばない）と全サブインダストリー。
 * {sheets: [{name, checked, note}], job: 作っている途中の仕事の知らせ（無ければ空）}
 */
function smtgRunData_() {
  const base = smtgRead_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
  const all = MTG_OPTIONS.combined.name;
  const sheets = smtgSheetCounts_(base).map(c => ({name: c.name, checked: c.count > 0, note: 'リクエスト ' + c.count + '件'}))
    .concat([{name: all, checked: true, note: 'リクエスト ' + base.requests.length + '件'}]);
  return {sheets, job: smtgJobPendingText_()};
}

/** 作る前の確かめの文（会議日・前回の会議・作る資料・サンプルの有無・作成済み）。本番なら記録の注意も */
function smtgRunCheckText_(final, meeting, names) {
  const base = smtgRead_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
  const prev = smtgPrevious_(base.list, meeting);
  return (final ? '本番の資料を作り、会議日を記録します。' : '試しに資料を作ります（会議日は記録しません。ファイル名に「試作」が付きます）。') +
    '\n\n会議日：' + mtgDateLabel_(meeting) + '\n' +
    '前回の会議：' + (prev ? mtgDateLabel_(prev.date) + '（その時点からの変化を出します）' : '記録なし（前回からの変化は出しません）') + '\n' +
    '作る資料：' + names.join('・') + '（' + names.length + '件）' +
    (base.sample ? '\n\n※ サンプルデータ（' + SMP_OPTIONS.namePrefix + '）が入っています。資料に【サンプルデータを含む】と出ます。' +
      (final ? '配る資料なら、先に removeServiceMtgSampleData() で消してください。' : '') : '') +
    (final && base.list.some(m => m.date === meeting) ? '\n\n※ この会議日の資料は作成済みです。作り直すと、次の会議の起点が今になります。' : '') +
    (final ? '\n\n記録した会議日は、次の会議の資料の起点になります。配布したあとは作り直さないでください。' : '');
}

/**
 * 資料を作り始める（画面の「作成する」から。MeetingSchedule.gs の runMeetingMaterials）。作っている途中の仕事があれば、cancelJob ならやめて作り、
 * そうでなければ止める。リクエストの無いサブインダストリーは作らない（選んだときを除く）。戻り値は終わり（区切ったとき）の知らせ
 */
function smtgRunStart_(final, meeting, names, cancelJob) {
  const started = mtgNow_();   // 6分の上限は、この実行を始めたときから
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  mtgRequireSlides_();
  if (smtgJobLoad_()) {
    if (!cancelJob) throw new Error(smtgJobPendingText_());
    smtgJobClear_();
    smtgJobDeleteTriggers_();
  }
  const now = new Date();
  const base = smtgRead_(ss, now);
  const job = {final, meeting, now: now.getTime(), names, skipped: [], items: [], recorded: false, folderUrl: '', sample: base.sample};
  const done = smtgJobRun_(ss, job, started, base);
  const message = smtgJobMessage_(job, done);
  console.log(message);
  return message;
}

/** リクエストを読むシート（サブインダストリー）ごとのリクエストの件数 [{name, count}]（DIFF_RULES の順） */
function smtgSheetCounts_(base) {
  return base.sources.map(name => ({name, count: base.requests.filter(r => r.sheet === name).length}));
}




/* ---------------- 作る仕事（6分の上限に掛からないよう、区切って続ける） ---------------- */

/**
 * 資料を作る仕事を、時間の許す限り進める。job = {final, meeting, now: 作り始めた時刻（ms）, names: 作る資料（サブインダストリー・全サブインダストリー）,
 * skipped: リクエストが無いため作らないサブインダストリー, items: 作った資料 [{name, line}], recorded: 会議日を記録したか, folderUrl}。
 * 資料を1つずつ作り（本番は格納とインデックスも）、すべて作ったら本番は会議日を記録する。base は読んだシート（無ければ読む）。
 * 次の資料が SMTG_OPTIONS.run の見積もりで上限を越えそうなら、仕事を保存して続きを予約し、false を返す。全部済んだら true
 */
function smtgJobRun_(ss, job, started, base) {
  const R = SMTG_OPTIONS.run;
  const fits = ms => mtgNow_() - started + ms <= R.limitMs;
  const folder = mtgFolder_(ss);
  job.folderUrl = folder.getUrl();
  const now = new Date(job.now);
  let data = base || null;
  let made = 0;   // この実行で作った資料の数（1つ目は見積もりにかかわらず作る。止まらずに進むように）
  while (job.items.length < job.names.length) {
    const need = (data ? 0 : R.readMs) + R.deckMs + (job.final ? R.storeMs : 0);
    if (made && !fits(need)) return smtgJobPause_(job);
    if (!data) data = smtgRead_(ss, now);
    job.sample = job.sample || data.sample;
    job.items.push(smtgCreateOne_(data, job, job.names[job.items.length], folder));
    made++;
  }
  if (job.final && !job.recorded) {
    smtgRecordMeeting_(data || smtgRead_(ss, now), job.meeting);
    job.recorded = true;
  }
  return true;
}

/**
 * 1つの資料（name：サブインダストリーか全サブインダストリー）を作って保存する。本番は会議のフォルダ（yyMMdd_RX Service Meeting_第N回）に格納し、
 * インデックスの種別「RX Service Meeting」の行の name の列に書く。戻り値 {name, line: 知らせる文}
 */
function smtgCreateOne_(base, job, name, folder) {
  const data = smtgView_(base, name, job.meeting);
  const slides = smtgSlides_(data, job.meeting, job.final);
  const out = smtgWithPalette_(() => mtgExportDeck_(smtgFileName_(job.meeting, name, job.final), slides, folder));
  const k = data.kpi;
  const lines = ['「' + name + '」：' + out.pdf.getName() + '（リクエスト ' + k.requests + '・判断待ち ' + k.undecided + '・サービス案 ' + k.services +
    '・確認事項 ' + data.checks.length + '。' + slides.length + 'ページ）', '　' + out.pdf.getUrl()];
  if (!job.final) return {name, line: lines.join('\n')};
  // 会議のフォルダ：サブインダストリーのフォルダ（DIFF_RULES の meetingFolder。全サブインダストリーは MTG_OPTIONS.combined.meetingFolder）の中の
  // 「yyMMdd_RX Service Meeting_第N回」（無ければ作る。mtgMeetingFolder_）
  let stored = {file: out.pdf, folder};
  const moved = mtgStoreToMeetingFolder_(name, job.meeting, out.pdf, SMTG_OPTIONS.meetingName);
  if (moved.error) lines.push('　会議のフォルダに格納できませんでした（' + moved.error + '）。PDF は保存先に残しています。');
  else {
    stored = moved;
    lines.push('　会議のフォルダ「' + moved.folder.getName() + '」に' + (moved.created ? '作って' : '') + '格納しました：' + moved.folder.getUrl());
  }
  const index = mtgWriteIndex_(name, job.meeting, stored, SMTG_OPTIONS.index.kind);
  lines.push('　' + (index.error || '「' + index.title + '」シートの「' + SMTG_OPTIONS.index.kind + '」の行に書きました。'));
  return {name, line: lines.join('\n')};
}

/** 終わり（区切ったとき）の知らせ */
function smtgJobMessage_(job, done) {
  const lines = job.items.map(item => item.line);
  if (!done) {
    const rest = job.names.slice(job.items.length);
    return ['時間がかかるため、ここで区切りました（Apps Script の1回の実行は6分まで）。残りの資料 ' + rest.length + '件（' + rest.join('・') + '）は、約' +
      Math.round(SMTG_OPTIONS.run.continueAfterMs / 60000) + '分後に自動で続けます。',
    '続きが終わったら、エディタで showServiceMtgSummariesResult() を実行すると結果が出ます（実行ログにも出ます）。終わるまで、RX Service Meeting の資料の関数は実行しないでください。',
    ''].concat(lines, ['', '保存先：' + job.folderUrl]).join('\n');
  }
  if (job.skipped && job.skipped.length) lines.push('', 'リクエストが無いため作らなかったサブインダストリー：' + job.skipped.join('・'));
  if (job.sample) lines.push('', '※ サンプルデータ（' + SMP_OPTIONS.namePrefix + '）を含みます。配る資料なら、removeServiceMtgSampleData() で消してから作り直してください。');
  lines.push('', '保存先：' + job.folderUrl);
  lines.push(job.final
    ? '会議日（' + job.meeting + '）を記録しました。次の会議の資料は、今からの変化（新規リクエスト・判断の変更・タスクの完了など）を出します。' +
      '配るときは、PDF を参加者に共有してください（保存先のフォルダは共有しない）。'
    : '試作のため、会議日は記録していません。');
  return lines.join('\n');
}

/** 仕事を保存し（スクリプトのプロパティ）、続きのトリガーを1つだけ作る。false を返す */
function smtgJobPause_(job) {
  const R = SMTG_OPTIONS.run;
  PropertiesService.getScriptProperties().setProperty(R.jobKey, JSON.stringify(Object.assign({}, job, {savedAt: Date.now()})));
  smtgJobDeleteTriggers_();
  ScriptApp.newTrigger(R.handler).timeBased().after(R.continueAfterMs).create();
  console.log('RX Service Meeting の資料：時間がかかるため区切りました（資料 ' + job.items.length + '/' + job.names.length + '）。続きを予約しました。');
  return false;
}

/** 保存した仕事（無ければ null） */
function smtgJobLoad_() {
  return mtgReadJson_(PropertiesService.getScriptProperties(), SMTG_OPTIONS.run.jobKey);
}

function smtgJobClear_() {
  PropertiesService.getScriptProperties().deleteProperty(SMTG_OPTIONS.run.jobKey);
}

function smtgJobDeleteTriggers_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === SMTG_OPTIONS.run.handler).forEach(t => ScriptApp.deleteTrigger(t));
}

/** 結果の知らせを保存する（スクリプトのプロパティ。2,000文字ずつ分ける） */
function smtgJobSaveResult_(message) {
  const R = SMTG_OPTIONS.run;
  const props = PropertiesService.getScriptProperties();
  props.getKeys().filter(key => key.indexOf(R.resultKey + '_') === 0).forEach(key => props.deleteProperty(key));
  const text = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' に終わりました。\n\n' + message;
  for (let i = 0; i * 2000 < text.length; i++) props.setProperty(R.resultKey + '_' + i, text.slice(i * 2000, (i + 1) * 2000));
}

/**
 * 区切った RX Service Meeting の資料の続き（時間主導のトリガーから。1回きり）。保存した仕事を続け、まだ残れば、また区切って次を予約する。
 * 終わったら（止まったら）知らせを保存する（showServiceMtgSummariesResult() で出す。実行ログにも出す）
 */
function continueServiceMtgSummaries() {
  const started = mtgNow_();
  smtgJobDeleteTriggers_();
  const job = smtgJobLoad_();
  if (!job) return;
  let message;
  try {
    if (!smtgJobRun_(SpreadsheetApp.getActiveSpreadsheet(), job, started, null)) return;   // また区切った（次を予約した）
    message = smtgJobMessage_(job, true);
  } catch (error) {
    console.error(error && error.stack || error);
    message = '途中で止まりました：' + (error && error.message || error) +
      '\ncreateServiceMtgSummaries()（試作なら previewServiceMtgSummaries()）で、作り直してください。\n\n' + smtgJobMessage_(job, true);
  }
  smtgJobClear_();
  smtgJobSaveResult_(message);
  console.log(message);
}

/** 自動で続けた RX Service Meeting の資料の結果・進み具合を出す（エディタから実行） */
function showServiceMtgSummariesResult() {
  const R = SMTG_OPTIONS.run;
  const ui = mtgUi_();
  const job = smtgJobLoad_();
  let message;
  if (job) {
    const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === R.handler).length;
    message = 'RX Service Meeting の資料を作っている途中です（会議日 ' + job.meeting + '。資料 ' + job.items.length + '/' + job.names.length + '件）。' +
      (triggers ? '自動で続けています。少し待ってから、もう一度実行してください。'
        : '続きの予約がありません。createServiceMtgSummaries()（試作なら previewServiceMtgSummaries()）で作り直してください。');
  } else {
    const props = PropertiesService.getScriptProperties();
    const keys = props.getKeys().filter(key => key.indexOf(R.resultKey + '_') === 0).sort((a, b) => Number(a.split('_').pop()) - Number(b.split('_').pop()));
    message = keys.length ? keys.map(key => props.getProperty(key)).join('') : '自動で続けた RX Service Meeting の資料の結果はありません。';
  }
  console.log(message);
  ui.alert(SVC_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** 作っている途中の仕事があれば、その知らせ（画面に出し、「やめて新しく作る」を選んでもらう）。無ければ空 */
function smtgJobPendingText_() {
  const job = smtgJobLoad_();
  if (!job) return '';
  return '前に始めた RX Service Meeting の資料（会議日 ' + job.meeting + '・' + (job.final ? '本番' : '試作') + '。資料 ' + job.items.length + '/' +
    job.names.length + '件）を、自動で続けている途中です。やめて新しく作るときは「やめて新しく作る」にチェックを入れてください' +
    '（続きが終わったら showServiceMtgSummariesResult() で結果を見られます）。';
}

/** ファイル名（20261113_スーパー・GMS_RX Service Meeting用資料。試作は「_試作」つき） */
function smtgFileName_(meeting, name, final) {
  return SMTG_OPTIONS.fileName.replace('{date}', meeting.replace(/\//g, '')).replace('{name}', name) + (final ? '' : '_試作');
}

/**
 * fn を、会議サマリの色（MTG_OPTIONS.colors）を紫（SMTG_OPTIONS.colors）に置き換えて実行する（終わったら元に戻す）。
 * 区分などのラベル（mtgPills_）も紫で描く（mtgWithColors_ と違い、元の色 baseColors を残さない）
 */
function smtgWithPalette_(fn) {
  const o = MTG_OPTIONS;
  const base = o.colors;
  o.colors = Object.assign({}, base, SMTG_OPTIONS.colors);
  try {
    return fn();
  } finally {
    o.colors = base;
  }
}

/* ---------------- 読む・まとめる ---------------- */

/** 判断の選択肢（サービスリクエストの「判断」のプルダウン。[未判断, サービス化検討, 棄却]） */
function smtgDecisionOptions_() {
  return SVC_OPTIONS.requestColumns.find(c => c.key === 'decision').options;
}

/** タスクの状況の選択肢（タスク管理の「状況」のプルダウン。[未着手, 対応中, 完了]） */
function smtgStatusOptions_() {
  return TASK_OPTIONS.columns.find(c => c.key === 'status').options;
}

/** 判断の分類：'undecided' 未判断（空も）・'considering' サービス化検討・'rejected' 棄却・'other' 選択肢に無い値 */
function smtgDecisionKind_(decision) {
  const i = smtgDecisionOptions_().indexOf(decision);
  return ['undecided', 'considering', 'rejected'][i] || 'other';
}

/** 同じものを1つにした並び（全角・半角と空白の違いは同じとみなす。空は除く。最初の書き方を残す） */
function smtgDistinct_(list) {
  const seen = new Set();
  return list.filter(v => {
    const key = credNormalize_(v);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/** 1つの資料の中身（name はサブインダストリーか全サブインダストリー。省けば全サブインダストリー）。smtgRead_ と smtgView_ をまとめたもの */
function smtgCollect_(ss, now, meeting, name) {
  return smtgView_(smtgRead_(ss, now), name || MTG_OPTIONS.combined.name, meeting);
}

/**
 * シートを読む（すべてのリクエスト・タスク・案件。どの資料にも使う）。リクエストには、どのサブインダストリー（リクエストを読むシート）のものか（sheet）を付ける：
 * 案件のあるシート。案件が無ければ、リクエストの「サブインダストリー」と同じ名前（前の名前も）のシート。どちらも無ければ空（全サブインダストリーの資料だけに出す）。
 * 戻り値 {now, list: 会議の記録, requests, tasks, services: [{key, name, requests, tasks}], hasWeighted, dealGaps, sources, sample}
 */
function smtgRead_(ss, now) {
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) {
    throw new Error('「' + SVC_OPTIONS.requestSheet + '」シートがありません。setupRequestSheet() を実行してから、リクエストを登録してください' +
      '（試すときは addServiceMtgSampleData() でサンプルを入れられます）。');
  }
  const marker = SMP_OPTIONS.namePrefix;
  const undecided = smtgDecisionOptions_()[0];
  const sources = svcSourceSheets_().filter(name => ss.getSheetByName(name));
  const resolve = diffSheetNameResolver_();
  const sheetOf = text => {
    const name = resolve(String(text || '').trim());
    return sources.indexOf(name) >= 0 ? name : (sources.find(s => credNormalize_(s) === credNormalize_(name)) || '');
  };
  const requests = svcRegistered_(sheet).rows.map(r => {
    const v = r.values;
    const decision = v.decision || undecided;
    return {row: r.row, id: v.id || '', number: svcIdNumber_(v.id), subIndustry: v.subIndustry, customer: v.customer, project: v.project,
      accountDept: v.accountDept, accountOwner: v.accountOwner, bxDept: v.bxDept, bxOwner: v.bxOwner, request: v.request,
      decision, kind: smtgDecisionKind_(decision), feedback: v.feedback || '', service: v.service || '', department: v.department || '',
      owner: v.owner || '', sample: v.request.indexOf(marker) === 0, deal: null, serviceKey: '', sheet: ''};
  });
  const byNumber = new Map();
  requests.forEach(r => { if (r.number && !byNumber.has(r.number)) byNumber.set(r.number, r); });

  // 案件（得意先＋案件名）：リクエストを読むシートから。読めないシートは飛ばして知らせる
  const deals = new Map();
  const dealGaps = [];
  let hasWeighted = false;
  sources.forEach(name => {
    try {
      const table = mtgReadDeals_(ss.getSheetByName(name), DIFF_RULES[name]);
      if (table.has.weighted) hasWeighted = true;
      table.deals.forEach(d => { if (!deals.has(d.key)) deals.set(d.key, Object.assign({sheet: name}, d)); });
    } catch (e) {
      dealGaps.push(name + '：' + svcErrorText_(e));
    }
  });
  requests.forEach(r => {
    r.deal = deals.get(mtgDealKey_(r.customer, r.project)) || null;
    r.sheet = r.deal ? r.deal.sheet : sheetOf(r.subIndustry);
  });

  const tasks = smtgReadTasks_(ss, byNumber, marker);

  // サービス案：「サービス案」に同じ名前を入れたリクエスト（棄却は除く）
  const services = [];
  const byService = new Map();
  requests.forEach(r => {
    if (!r.service || r.kind === 'rejected') return;
    const key = credNormalize_(r.service);
    let s = byService.get(key);
    if (!s) {
      s = {key, name: r.service, requests: [], tasks: []};
      byService.set(key, s);
      services.push(s);
    }
    s.requests.push(r);
    r.serviceKey = key;
  });
  tasks.forEach(t => {
    const s = t.request && t.request.serviceKey ? byService.get(t.request.serviceKey) : null;
    t.serviceName = s ? s.name : '';
    if (s) s.tasks.push(t);
  });
  return {now, list: smtgMeetings_(), requests, tasks, services, hasWeighted, dealGaps, sources,
    sample: requests.some(r => r.sample) || tasks.some(t => t.sample)};
}

/**
 * 1つの資料の中身。name がサブインダストリーなら、そのサブインダストリーのリクエストと、それを含むサービス案（タスクはサービス案全体）にする。
 * 全サブインダストリー（MTG_OPTIONS.combined.name）ならすべて。サービス案の得意先・期待値・件数はこの資料のリクエストの分、
 * サブインダストリー・主管本部と担当者・タスク・顧客・業界の広がりはサービス案全体。
 * 戻り値 {name, all, now, meeting, list, previous, previousCounts, requests, tasks, services, kpi, moves, checks, fit, hasWeighted, dealGaps, sources, sample}
 */
function smtgView_(base, name, meeting) {
  const all = name === MTG_OPTIONS.combined.name;
  const requests = all ? base.requests : base.requests.filter(r => r.sheet === name);
  const inView = new Set(requests);
  const services = base.services.filter(s => s.requests.some(r => inView.has(r))).map(s => {
    const own = s.requests.filter(r => inView.has(r));
    const dealsOf = new Map();
    own.forEach(r => { if (r.deal) dealsOf.set(r.deal.key, r.deal); });
    const deals = Array.from(dealsOf.values());
    const active = deals.filter(d => d.kind === 'active');
    return {key: s.key, name: s.name, requests: own, allRequests: s.requests, tasks: s.tasks,
      customers: smtgDistinct_(own.map(r => r.customer)),
      subIndustries: smtgDistinct_(s.requests.map(r => r.subIndustry)),
      owners: smtgDistinct_(s.requests.map(r => [r.department, r.owner].filter(Boolean).join(' '))),
      deals, active, weighted: active.reduce((sum, d) => sum + (d.weighted || 0), 0), counts: smtgStatusCounts_(s.tasks),
      firstNumber: Math.min.apply(null, own.map(r => r.number || Infinity))};
  });
  // リクエストの多い順 → 元になった案件の期待値の大きい順 → 最初のリクエストの ID の順
  services.sort((a, b) => (b.requests.length - a.requests.length) || (b.weighted - a.weighted) || (a.firstNumber - b.firstNumber));
  const keys = new Set(services.map(s => s.key));
  // タスク：この資料のリクエストか、この資料のサービス案のタスク（全サブインダストリーは、リクエストに結び付かないタスクも）
  const tasks = base.tasks.filter(t => (t.request ? inView.has(t.request) || keys.has(t.request.serviceKey) : all));
  const counts = smtgStatusCounts_(tasks);
  const kpi = {
    requests: requests.length,
    undecided: requests.filter(r => r.kind === 'undecided').length,
    considering: requests.filter(r => r.kind === 'considering').length,
    rejected: requests.filter(r => r.kind === 'rejected').length,
    services: services.length,
    serviceRequests: services.reduce((sum, s) => sum + s.requests.length, 0),
    customers: smtgDistinct_([].concat.apply([], services.map(s => s.customers))).length,
    tasks: tasks.length,
    done: counts.done
  };
  const previous = smtgPrevious_(base.list, meeting);
  const data = {name, all, now: base.now, meeting, list: base.list, previous,
    previousCounts: previous && previous.counts ? previous.counts[name] || null : null,
    requests, tasks, services, kpi, hasWeighted: base.hasWeighted, dealGaps: base.dealGaps, sources: base.sources,
    sample: requests.some(r => r.sample) || tasks.some(t => t.sample)};
  data.moves = smtgMoves_(data);
  data.checks = smtgChecks_(data);
  data.fit = smtgFit_(services);
  return data;
}

/**
 * タスク管理のタスク（「サービス案」か「タスク」が入っている行）。[{row, label: サービス案の値, number: その ID の番号, request: そのリクエスト（無ければ null）,
 * task, department, owner, status, note, key: 前回と比べる印, sample}]。シートが無ければ空
 */
function smtgReadTasks_(ss, byNumber, marker) {
  const sheet = ss.getSheetByName(TASK_OPTIONS.sheet);
  const headerRow = TASK_OPTIONS.headerRow;
  if (!sheet || sheet.getLastRow() <= headerRow) return [];
  const cols = taskColumnMap_(sheet, TASK_OPTIONS.columns, headerRow);
  const tasks = [];
  sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, Math.max(sheet.getLastColumn(), 1)).getDisplayValues().forEach((line, i) => {
    const get = key => (cols[key] ? String(line[cols[key] - 1] == null ? '' : line[cols[key] - 1]).trim() : '');
    const t = {row: headerRow + 1 + i, label: get('service'), task: get('task'), department: get('department'), owner: get('owner'),
      status: get('status'), note: get('note')};
    if (!t.label && !t.task) return;
    t.number = svcLabelIdNumber_(t.label);
    t.request = (t.number && byNumber.get(t.number)) || null;
    t.key = mtgHash_(t.number + '|' + credNormalize_(t.task));
    t.sample = t.label.indexOf(marker) >= 0 || t.task.indexOf(marker) === 0;
    tasks.push(t);
  });
  return tasks;
}

/** タスクの状況の件数 {total, todo: 未着手, doing: 対応中, done: 完了, blank: 未入力, other: 選択肢に無い値} */
function smtgStatusCounts_(tasks) {
  const [todo, doing, done] = smtgStatusOptions_();
  const out = {total: tasks.length, todo: 0, doing: 0, done: 0, blank: 0, other: 0};
  tasks.forEach(t => {
    const key = t.status === todo ? 'todo' : t.status === doing ? 'doing' : t.status === done ? 'done' : t.status ? 'other' : 'blank';
    out[key]++;
  });
  return out;
}

/**
 * 前回の会議からの動き {first: 前回の記録が無い, partial: 前回の判断・状況の記録が無い（件数だけ）, newRequests, decided: [{request, from}],
 * completed: 完了になったタスク, newTasks: 新しいタスク}
 */
function smtgMoves_(data) {
  const prev = data.previous;
  const out = {first: !prev, partial: false, newRequests: [], decided: [], completed: [], newTasks: []};
  if (!prev) return out;
  out.newRequests = data.requests.filter(r => r.number > (prev.lastId || 0));
  if (!prev.decisions || !prev.tasks) {
    out.partial = true;
    return out;
  }
  const decisions = smtgParseRecord_(prev.decisions);
  const statuses = smtgParseRecord_(prev.tasks);
  const options = smtgDecisionOptions_();
  data.requests.forEach(r => {
    const before = decisions[r.number];
    if (!r.number || before == null || before === smtgDecisionCode_(r.decision)) return;
    out.decided.push({request: r, from: smtgDecisionOf_(before, options)});
  });
  const done = smtgStatusOptions_()[2];
  data.tasks.forEach(t => {
    const before = statuses[t.key];
    if (before == null) out.newTasks.push(t);
    else if (t.status === done && before !== smtgStatusCode_(done)) out.completed.push(t);
  });
  return out;
}

/**
 * 入力内容の確認事項 [{type: SMTG_CHECKS の種類, target: 対象（ID とリクエスト・サービス案・タスク）, text: 内容}]（種類の順、同じ種類は ID・行の順）
 */
function smtgChecks_(data) {
  const out = [];
  const add = (type, target, text) => out.push({type, target, text});
  const [, considering, rejected] = smtgDecisionOptions_();
  const label = r => (r.id || '（IDなし）') + ' ' + r.request;
  const readable = data.sources.length > data.dealGaps.length;   // 案件を読めたシートがある
  data.requests.forEach(r => {
    if (r.kind === 'considering' && !r.service) add('サービス案未設定', label(r), '判断が「' + considering + '」で、サービス案が空');
    if (r.kind === 'rejected' && r.service) {
      add('判断とサービス案の不整合', label(r), '判断が「' + rejected + '」で、サービス案「' + r.service + '」が入っている（サービス案には数えない）');
    }
    if (r.kind === 'other') add('選択肢外の判断', label(r), '判断「' + r.decision + '」がプルダウンの選択肢に無い');
    if (!r.deal && readable) {
      add('元の案件が見つからない', label(r), '得意先「' + (r.customer || '未入力') + '」・案件名「' + (r.project || '未入力') + '」の案件が、' +
        svcSourceLabel_() + 'に無い');
    }
  });
  data.services.forEach(s => {
    if (!s.owners.length) add('主管本部・担当者未設定', s.name, 'サービス案のリクエストに、主管本部・担当者が入っていない');
    if (!s.tasks.length) add('タスクなし', s.name, '「' + TASK_OPTIONS.sheet + '」に、このサービス案のタスクが無い');
  });
  data.tasks.forEach(t => {
    const target = (t.task || '（タスク未入力）') + '（' + TASK_OPTIONS.sheet + ' ' + t.row + '行目）';
    if (!t.owner) add('タスクの担当者未入力', target, '担当者が空');
    if (!t.status) add('タスクの状況未入力', target, '状況が空');
    if (!t.request) {
      add('リクエストが見つからない', target, t.label ? 'サービス案「' + t.label + '」の ID のリクエストが、「' + SVC_OPTIONS.requestSheet + '」に無い'
        : 'サービス案が空');
    }
  });
  return out.map((item, i) => Object.assign({order: i}, item))
    .sort((a, b) => (SMTG_CHECKS.indexOf(a.type) - SMTG_CHECKS.indexOf(b.type)) || (a.order - b.order));
}

/**
 * 顧客・業界の広がり（サービス案のすべてのリクエスト。サブインダストリーの資料でも、ほかのサブインダストリーの分を数える）{columns: サブインダストリー（リクエストの多い順。fitColumns を越えた分は「その他」）, rows: [{service, counts: 列ごとの件数, total}]}
 */
function smtgFit_(services) {
  const totals = new Map();
  const names = new Map();
  services.forEach(s => (s.allRequests || s.requests).forEach(r => {
    const key = credNormalize_(r.subIndustry) || '';
    if (!names.has(key)) names.set(key, r.subIndustry || '（未入力）');
    totals.set(key, (totals.get(key) || 0) + 1);
  }));
  const keys = Array.from(totals.keys()).sort((a, b) => totals.get(b) - totals.get(a));
  const max = SMTG_OPTIONS.fitColumns;
  const shown = keys.length > max ? keys.slice(0, max - 1) : keys;
  const columns = shown.map(k => names.get(k)).concat(keys.length > shown.length ? ['その他'] : []);
  const rows = services.map(s => {
    const counts = columns.map(() => 0);
    const list = s.allRequests || s.requests;
    list.forEach(r => {
      const i = shown.indexOf(credNormalize_(r.subIndustry) || '');
      counts[i >= 0 ? i : columns.length - 1]++;
    });
    return {service: s, counts, total: list.length};
  });
  return {columns, rows};
}

/* ---------------- 会議の記録 ---------------- */

/** 記録した会議 [{date, at, lastId}]（会議日の順） */
function smtgMeetings_() {
  const list = mtgReadJson_(PropertiesService.getDocumentProperties(), SMTG_OPTIONS.meetingsKey);
  return (Array.isArray(list) ? list : []).filter(m => m && /^\d{4}\/\d{2}\/\d{2}$/.test(m.date))
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}

/**
 * 前回の会議（meeting より前で、いちばん新しい会議日）の記録と、その時点の判断・状況・資料ごとの件数 {…, decisions, tasks, counts: {資料の名前: 件数}}。無ければ null
 */
function smtgPrevious_(list, meeting) {
  const prev = list.filter(m => m.date < meeting).pop();
  if (!prev) return null;
  const detail = mtgReadJson_(PropertiesService.getDocumentProperties(), SMTG_OPTIONS.meetingDetailPrefix + prev.date.replace(/\//g, '')) || {};
  return Object.assign({}, prev, {decisions: detail.decisions || '', tasks: detail.tasks || '', counts: detail.counts || null});
}

/**
 * 本番の会議日と、その時点の状況を記録する（同じ会議日は置き換える）。一覧は keepMeetings まで。判断・状況と、資料（サブインダストリー・全サブインダストリー）
 * ごとの件数は会議日ごとの別のキー（長すぎれば件数だけ残し、次の資料の判断の変更・タスクの完了は出さない）。base は smtgRead_ の戻り値
 */
function smtgRecordMeeting_(base, meeting) {
  const props = PropertiesService.getDocumentProperties();
  const o = SMTG_OPTIONS;
  const data = base;
  const entry = {date: meeting, at: data.now.getTime(), lastId: Math.max.apply(null, [0].concat(data.requests.map(r => r.number)))};
  const counts = {};
  base.sources.concat([MTG_OPTIONS.combined.name]).forEach(name => { counts[name] = smtgView_(base, name, meeting).kpi; });
  const all = smtgMeetings_().filter(m => m.date !== meeting).concat([entry])
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  const keep = all.slice(-o.keepMeetings);
  all.slice(0, all.length - keep.length).forEach(m => props.deleteProperty(o.meetingDetailPrefix + m.date.replace(/\//g, '')));
  let detail = JSON.stringify({
    decisions: data.requests.filter(r => r.number).map(r => r.number + ':' + smtgDecisionCode_(r.decision)).join(','),
    tasks: data.tasks.map(t => t.key + ':' + smtgStatusCode_(t.status)).join(','),
    counts
  });
  if (detail.length > o.detailMaxLength) detail = JSON.stringify({counts});
  props.setProperty(o.meetingDetailPrefix + meeting.replace(/\//g, ''), detail);
  props.setProperty(o.meetingsKey, JSON.stringify(keep));
}

/** 「12:1,13:0」→ {12: '1', 13: '0'} */
function smtgParseRecord_(text) {
  const out = {};
  String(text || '').split(',').filter(Boolean).forEach(pair => {
    const at = pair.lastIndexOf(':');
    if (at > 0) out[pair.slice(0, at)] = pair.slice(at + 1);
  });
  return out;
}

/** 判断の記録の印（選択肢の番号。選択肢に無い値は x） */
function smtgDecisionCode_(decision) {
  const i = smtgDecisionOptions_().indexOf(decision);
  return i >= 0 ? String(i) : 'x';
}

/** 記録の印から判断の名前（選択肢に無い値だったときは「選択肢外」） */
function smtgDecisionOf_(code, options) {
  return options[Number(code)] || (code === 'x' ? '選択肢外' : '');
}

/** 状況の記録の印（選択肢の番号。空は b、選択肢に無い値は x） */
function smtgStatusCode_(status) {
  if (!status) return 'b';
  const i = smtgStatusOptions_().indexOf(status);
  return i >= 0 ? String(i) : 'x';
}

/* ---------------- スライド ---------------- */

/**
 * スライドの並び：表紙・サマリー・サービス案の一覧・顧客・業界の広がり・サービス案（1つずつ）・判断待ちのリクエスト・扉・Appendix ①〜③。
 * 紫の色で作る（smtgWithPalette_。Appendix はグレー）。ページ番号を付けて返す（mtgNumberSlides_）
 */
function smtgSlides_(data, meeting, final) {
  return smtgWithPalette_(() => {
    const cover = smtgCoverSlide_(data, meeting, final);
    const fit = smtgFitSlides_(data);
    const pending = smtgPendingSlides_(data);
    const details = data.services.map(s => smtgDetailSlides_(data, s));
    // サービス案の一覧のページの数は、ページ番号の文字に左右されないので、先に数える
    const listPages = smtgServicesSlides_(data, data.services.map(() => 0)).length;
    const pages = {summary: 2, services: 3};
    pages.fit = pages.services + listPages;
    pages.detail = pages.fit + fit.length;
    const starts = [];
    let page = pages.detail;
    details.forEach(list => { starts.push(page); page += list.length; });
    pages.pending = page;
    const list = smtgServicesSlides_(data, starts);
    const main = list.concat(fit, [].concat.apply([], details), pending);
    const appendix = mtgWithColors_(MTG_OPTIONS.appendixColors, () => {
      const parts = {requests: smtgRequestsAppendix_(data), tasks: smtgTasksAppendix_(data), checks: smtgChecksAppendix_(data)};
      pages.appendix = 2 + main.length + 1;   // 表紙・サマリーのあとに本編、その次が扉
      pages.requests = pages.appendix + 1;
      pages.tasks = pages.requests + parts.requests.length;
      pages.checks = pages.tasks + parts.tasks.length;
      return [].concat(parts.requests, parts.tasks, parts.checks);
    });
    const ranges = {services: [pages.services, pages.fit - 1], fit: [pages.fit, pages.detail - 1],
      detail: details.length ? [pages.detail, pages.pending - 1] : null, pending: [pages.pending, pages.appendix - 1],
      requests: [pages.requests, pages.tasks - 1], tasks: [pages.tasks, pages.checks - 1], checks: [pages.checks, pages.appendix + appendix.length]};
    const summary = smtgSummarySlide_(data, ranges);
    const slides = [cover, summary].concat(main, [mtgDividerSlide_(appendix, pages.requests)], appendix);
    return mtgNumberSlides_(smtgFrame_(data, meeting, final), slides);
  });
}

/** いちばん下の行（「スーパー・GMS（11/13 RX Service Meeting 資料）」。試作は【試作】、サンプルを含めば【サンプルデータを含む】）{source} */
function smtgFrame_(data, meeting, final) {
  return {source: (final ? '' : '【試作】') + (data.sample ? '【サンプルデータを含む】' : '') + data.name + '（' + mtgShortDate_(meeting) + ' ' +
    SMTG_OPTIONS.meetingName + ' 資料）　金額の単位：百万円'};
}

/** 表紙（ひな形の3枚目＝紫）：見出しにサブインダストリー（全サブインダストリー）、その下に会議の名前と「事前資料」、日付に会議日。その下に出典と対象を小さく */
function smtgCoverSlide_(data, meeting, final) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = data.kpi;
  const line = (y, text) => mtgBox_(S.x, y, S.w, 16, mtgWrap_(text, S.w, 9, 1)[0], {size: 9, color: c.white, align: 'CENTER'});
  return {
    title: SMTG_TITLES.cover,
    cover: true,
    service: true,
    template: {title: data.name, subtitle: SMTG_OPTIONS.meetingName + '　事前資料' + (final ? '' : '【試作】'),
      date: mtgDateLabel_(meeting), name: ''},
    elements: [
      line(350, '出典：「' + SVC_OPTIONS.requestSheet + '」「' + TASK_OPTIONS.sheet + '」シートおよび' + (data.all ? svcSourceLabel_() : '「' + data.name + '」') +
        'の案件（' +
        (data.sample ? 'サンプルデータを含む。' : '') + Utilities.formatDate(data.now, Session.getScriptTimeZone(), 'yyyy/M/d HH:mm') + ' 時点）'),
      line(365, '対象：登録リクエスト ' + k.requests + '件・サービス案 ' + k.services + '件・タスク ' + k.tasks + '件' +
        (data.previous ? '（前回の会議 ' + mtgShortDate_(data.previous.date) + ' からの変化）' : ''))
    ]
  };
}

/** 「P.3」「P.5〜9」（range = [始まり, 終わり]。無ければ「—」） */
function smtgPagesText_(range) {
  if (!range) return '—';
  return 'P.' + range[0] + (range[1] > range[0] ? '〜' + range[1] : '');
}

/**
 * 1 サマリー：リード文にサービス案・タスク・判断待ちの数と前回以降の動き。その下に
 * - 数字（登録リクエスト・判断待ち・サービス化検討・サービス案・タスクの完了）と前回比
 * - 前回以降の動き（新規リクエスト・判断の変更・完了したタスク・新しいタスク。前回の記録が無ければ灰色の「データなし」）
 * - 議題と資料のページ（SMTG_OPTIONS.agenda。議題ごとに、見るページとそこに書いたこと）
 */
function smtgSummarySlide_(data, ranges) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const G = MTG_OPTIONS.gapText;
  const k = data.kpi;
  const m = data.moves;
  const prev = data.previous;
  const pc = data.previousCounts;   // 前回の会議の時点の、この資料の件数
  const lead = 'サービス案 ' + k.services + '件（リクエスト ' + k.serviceRequests + '件・得意先 ' + k.customers + '社）、タスク ' + k.tasks + '件（完了 ' +
    k.done + '件）。登録リクエスト ' + k.requests + '件のうち、判断待ち ' + k.undecided + '件。\n' +
    (!prev ? '前回の会議の記録が無いため、前回からの変化は出していません（本番の資料を作ると、次の会議の資料から出ます）。'
      : '前回の会議（' + mtgShortDate_(prev.date) + '）以降：新規リクエスト ' + m.newRequests.length + '件' +
        (m.partial ? '（判断の変更・タスクの完了は、前回の記録が無いため出していません）。'
          : '・判断の変更 ' + m.decided.length + '件・完了したタスク ' + m.completed.length + '件・新しいタスク ' + m.newTasks.length + '件。'));
  const top = Math.max(S.summaryTop, Math.max.apply(null, mtgLead_(lead).map(e => e.y + e.h)) + 2);
  const el = [];

  // 数字：紫の帯の見出しと、大きな数字・小さな前回比
  const versus = (key, value) => (pc && pc[key] != null ? '前回比 ' + mtgSigned_(value - pc[key]) : '');
  const tiles = [
    {label: '登録リクエスト', value: k.requests + '件', sub: versus('requests', k.requests)},
    {label: '判断待ち（未判断）', value: k.undecided + '件', sub: versus('undecided', k.undecided)},
    {label: 'サービス化検討', value: k.considering + '件', sub: versus('considering', k.considering)},
    {label: 'サービス案', value: k.services + '件', sub: versus('services', k.services)},
    {label: 'タスク（完了／全体）', value: k.done + '／' + k.tasks + '件', sub: pc && pc.done != null ? '完了 前回比 ' + mtgSigned_(k.done - pc.done) : ''}
  ];
  const th = 34;
  const tw = (S.w - (tiles.length - 1) * 8) / tiles.length;
  tiles.forEach((t, i) => {
    const x = S.x + i * (tw + 8);
    const text = t.value + (t.sub ? '　' + t.sub : '');
    el.push({type: 'box', x, y: top, w: tw, h: th, fill: c.white, line: c.frame, lineWeight: 0.75, shape: 'RECTANGLE'});
    el.push(mtgBox_(x, top, tw, 14, mtgWrap_(t.label, tw, 8.5, 1)[0],
      {size: 8.5, bold: true, color: c.white, fill: c.green, shape: 'RECTANGLE', align: 'CENTER', role: 'tile'}));
    el.push(mtgBox_(x, top + 14, tw, th - 14, mtgWrap_(text, tw, 8, 1)[0], {size: 8, color: c.muted, align: 'CENTER', role: 'tile-value',
      spans: [{start: 0, end: t.value.length, size: 15, bold: true, color: c.greenDark}]}));
  });

  // 議題と資料のページ（いちばん下）
  const agenda = SMTG_OPTIONS.agenda;
  const ah = 62;
  const agendaY = S.bottom - ah;
  const aw = (S.w - (agenda.length - 1) * 8) / agenda.length;
  const heading = (y, text) => mtgBox_(S.x, y, S.w, 16, text, {size: 10.5, bold: true, color: c.ink, role: 'heading'});
  el.push(heading(agendaY - 18, '議題と資料のページ'));
  agenda.forEach((a, i) => {
    const x = S.x + i * (aw + 8);
    const pages = smtgDistinct_(a.pages.map(key => (ranges[key] ? smtgPagesText_(ranges[key]) : ''))).join('・') || '—';
    const note = (a.note.indexOf('\n') >= 0 ? a.note.split('\n').map(part => mtgWrap_(part, aw, 8, 1)[0]) : mtgWrap_(a.note, aw, 8, 2))
      .slice(0, 2).join('\n');
    el.push({type: 'box', x, y: agendaY, w: aw, h: ah, fill: c.white, line: c.frame, lineWeight: 0.75, shape: 'RECTANGLE'});
    el.push(mtgBox_(x, agendaY, aw, 14, mtgWrap_(a.label, aw, 8.5, 1)[0],
      {size: 8.5, bold: true, color: c.white, fill: c.green, shape: 'RECTANGLE', align: 'CENTER', role: 'tile'}));
    el.push(mtgBox_(x, agendaY + 14, aw, ah - 14, mtgWrap_(pages, aw, 10, 1)[0] + '\n' + note,
      {size: 8, color: c.muted, align: 'CENTER', role: 'agenda', spans: [{start: 0, end: mtgWrap_(pages, aw, 10, 1)[0].length, size: 10, bold: true,
        color: c.greenDark}]}));
  });

  // 前回以降の動き（数字と議題の間。入る行の数で名前を出す）
  const columns = mtgColumns_([132, 56, 0], 2);
  const shown = list => list.slice(0, SMTG_OPTIONS.movesMax);
  const more = list => (list.length > SMTG_OPTIONS.movesMax ? '　他 ' + (list.length - SMTG_OPTIONS.movesMax) + '件' : '');
  const reqText = r => (r.id || '（IDなし）') + ' ' + mtgClip_(r.request, 16);
  const taskText = t => mtgClip_(t.task || '（タスク未入力）', 14) + (t.serviceName ? '（' + mtgClip_(t.serviceName, 10) + '）' : '');
  const lines = [
    {label: '新規リクエスト', items: m.newRequests, text: r => reqText(r)},
    {label: '判断の変更', items: m.decided, partial: true, text: d => (d.request.id || '（IDなし）') + ' ' + (d.from || '—') + ' ▶ ' + d.request.decision},
    {label: '完了したタスク', items: m.completed, partial: true, text: taskText},
    {label: '新しいタスク', items: m.newTasks, partial: true, text: taskText}
  ];
  const tableY = top + th + 8 + 16;
  const room = agendaY - 18 - 6 - tableY;
  const build = n => [mtgHeadRow_(columns, ['区分', '件数', '内容'], S.bodySize)].concat(lines.map(line => {
    const label = {text: line.label, size: S.bodySize, bold: true, color: c.ink, fill: c.greenSoft};
    const reason = m.first ? '前回の会議の記録なし' : line.partial && m.partial ? '前回の判断・状況の記録なし' : '';
    if (reason) {
      return mtgRow_(columns, [label, {text: '—', size: S.bodySize, color: c.muted, fill: c.white, align: 'CENTER'},
        mtgGapCell_({size: S.bodySize, lines: n}, G + '：' + reason)], 0);
    }
    return mtgRow_(columns, [
      label,
      {text: line.items.length + '件', size: 10, bold: true, color: c.greenDark, fill: c.white, align: 'CENTER'},
      {text: line.items.length ? shown(line.items).map(line.text).join('、') + more(line.items) : MTG_OPTIONS.noneText, size: S.bodySize, lines: n,
        fill: c.white, color: line.items.length ? c.body : c.muted}
    ], 0);
  }));
  let rows = build(2);
  if (rows.reduce((h, r) => h + r.h, 0) > room) rows = build(1);
  el.push(heading(tableY - 16, prev ? '前回の会議（' + mtgShortDate_(prev.date) + '）以降の動き' : '前回以降の動き'),
    {type: 'table', x: S.x, y: tableY, columns, rows});
  return {title: SMTG_TITLES.summary, lead, elements: el};
}

/**
 * 表をページに並べる（入りきらなければ次のページに続ける。続きのページの見出しに「（続き）」）。
 * spec = {title, lead, size, sections: [{heading: 表の上の見出し（無くてもよい）, columns, head: 見出しの行の文字, headRow: 作った見出しの行（あれば head の代わり）, rows: [セルの並び],
 *   pills: {col, labels: i => [{text, tone}], width}, empty: {kind: 'none'|'gap', text}}]}。戻り値はスライドの並び。
 * spec.shrink なら、続きのページになるときは文字を 0.5pt ずつ小さくし（8pt まで）、1ページに入る大きさがあればそれにする
 */
function smtgPages_(spec) {
  if (!spec.shrink) return smtgLayPages_(spec);
  const first = smtgLayPages_(spec);
  for (let size = (spec.size || MTG_SLIDE.bodySize) - 0.5; first.length > 1 && size >= 8; size -= 0.5) {
    const smaller = smtgLayPages_(Object.assign({}, spec, {size}));
    if (smaller.length === 1) return smaller;
  }
  return first;
}

/** smtgPages_ の中身（spec.size の文字で並べる） */
function smtgLayPages_(spec) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const size = spec.size || S.bodySize;
  const headingH = 16;
  const gap = 8;
  // 中身はリード文の下から（リード文が短ければ上に詰める。会議サマリの「全サブインダストリーの状況」と同じ）
  const top = spec.lead ? Math.max(S.summaryTop, Math.max.apply(null, mtgLead_(spec.lead).map(e => e.y + e.h)) + 6) : S.top;
  const pages = [];
  let page = null;
  const newPage = () => {
    page = {elements: [], y: top, empty: null};
    pages.push(page);
  };
  newPage();
  spec.sections.forEach(sec => {
    const w = sec.columns.reduce((a, b) => a + b, 0);
    const head = sec.headRow || mtgHeadRow_(sec.columns, sec.head, size);
    const rows = sec.rows.map(cells => mtgRow_(sec.columns, cells.map(cell => Object.assign({size}, cell)), 0));
    let table = null;
    const tables = [];
    const open = cont => {
      if (sec.heading) {
        page.elements.push(mtgBox_(S.x, page.y, S.w, headingH, sec.heading + (cont ? '（続き）' : ''),
          {size: 10.5, bold: true, color: c.ink, role: 'heading'}));
        page.y += headingH;
      }
      table = {type: 'table', x: S.x, y: page.y, columns: sec.columns, rows: [head]};
      tables.push({table, index: []});
      page.elements.push(table);
      page.y += head.h;
    };
    const lead = (sec.heading ? headingH : 0) + head.h;
    if (!rows.length) {
      const empty = sec.empty || {kind: 'none'};
      const lines = empty.text ? mtgWrap_(empty.text, w, 9, 3).length : 0;
      const h = Math.max(40, mtgTextHeight_(1, 10.5) + lines * 9 * S.lineFactor + 8);
      if (page.y > top && page.y + lead + 8 + h > S.bottom) newPage();
      open(false);
      page.elements.push(mtgEmptyBox_(S.x, page.y + 6, w, h, empty.kind, empty.text || ''));
      page.y += 6 + h + gap;
      if (spec.sections.length === 1) page.empty = empty.kind;
      return;
    }
    rows.forEach((row, i) => {
      if (!table || page.y + row.h > S.bottom) {
        // 見出しと最初の行が入らなければ次のページから（ページの頭なら、そのまま置く）
        if (table || (page.y > top && page.y + lead + row.h > S.bottom)) newPage();
        open(i > 0);
      }
      table.rows.push(row);
      tables[tables.length - 1].index.push(i);
      page.y += row.h;
    });
    page.y += gap;
    if (sec.pills) {
      tables.forEach(t => t.index.forEach((index, n) => {
        const target = pages.find(p => p.elements.indexOf(t.table) >= 0);
        target.elements.push.apply(target.elements, mtgPills_(t.table, n + 1, sec.pills.col, sec.pills.labels(index), sec.pills.width));
      }));
    }
  });
  return pages.map((p, i) => Object.assign({title: spec.title + (i ? '（続き）' : ''), lead: spec.lead, elements: p.elements},
    p.empty ? {empty: p.empty} : {}));
}

/** 判断のラベル（サービス化検討は紫・棄却は濃いグレー・未判断などは薄いグレー） */
function smtgDecisionPill_(r) {
  const tone = {considering: 'up', rejected: 'down'}[r.kind] || 'note';
  return [{text: r.decision, tone}];
}

/** 状況のラベル（完了は紫・対応中は薄い紫・未着手は薄いグレー。空なら「未入力」） */
function smtgStatusPill_(t) {
  const [, doing, done] = smtgStatusOptions_();
  if (!t.status) return [{text: MTG_OPTIONS.blankText, tone: 'note'}];
  return [{text: t.status, tone: t.status === done ? 'up' : t.status === doing ? 'stall' : 'note'}];
}

/** ラベルの幅（判断・状況の選択肢のいちばん長いもの） */
function smtgPillWidth_(options, cw) {
  return mtgPillWidth_(options.concat([MTG_OPTIONS.blankText]), cw);
}

/** 元になった案件のセル（ステータス・期待値。取り込み元に無ければ灰色の「案件なし」） */
function smtgDealCell_(data, r, width, size) {
  const d = r.deal;
  const pt = size || MTG_SLIDE.bodySize;
  if (!d) return mtgGapCell_({size: pt, lines: 2}, '案件なし');
  const value = data.hasWeighted ? '期待値 ' + (d.weighted == null ? MTG_OPTIONS.blankText : mtgAmount_(d.weighted)) : '';
  return {text: mtgStatusLines_(d.status || MTG_OPTIONS.blankText, width, pt) + (value ? '\n' + value : ''), size: pt, lines: 3,
    fill: MTG_OPTIONS.colors.white};
}

/** 得意先（太字）・案件名・サブインダストリー（小さく灰色）のセル */
function smtgCustomerCell_(r, width, size) {
  const c = MTG_OPTIONS.colors;
  const pt = size || MTG_SLIDE.bodySize;
  const line = (text, w) => mtgWrap_(text, width, w, 1)[0];
  return Object.assign({text: [line(r.customer || '（得意先未入力）', pt), line(r.project || '（案件名未入力）', pt), line(r.subIndustry || '（サブインダストリー未入力）', 8)]
    .join('\n'), lines: 3, color: c.ink, fill: c.white, parts: [{bold: true}, {}, {size: 8, color: c.muted}]}, size ? {size} : {});
}

/** 営業のセル（アカウント責任者部署・アカウント責任者。BX がいれば「BX：」で続ける） */
function smtgSalesText_(r) {
  const account = [r.accountDept, r.accountOwner].filter(Boolean).join(' ');
  const bx = [r.bxDept, r.bxOwner].filter(Boolean).join(' ');
  return [account, bx ? 'BX：' + bx : ''].filter(Boolean).join('\n') || MTG_OPTIONS.blankText;
}

/** 空なら「未入力」（灰色の字） */
function smtgTextCell_(text, extra) {
  const c = MTG_OPTIONS.colors;
  return Object.assign({text: text || MTG_OPTIONS.blankText, color: text ? c.body : c.muted, fill: c.white}, extra || {});
}

/**
 * 2 サービス案の一覧：サービス案ごとのリクエストの数・得意先・サブインダストリー・元になった案件の期待値（進行中）・主管本部と担当者・タスク・詳細のページ。
 * starts はサービス案ごとの詳細のページ（0 なら空。ページの数を先に数えるとき）
 */
function smtgServicesSlides_(data, starts) {
  const c = MTG_OPTIONS.colors;
  const columns = mtgColumns_([0, 44, 124, 96, 66, 100, 104, 44], 0);
  const rows = data.services.map((s, i) => {
    const n = s.counts;
    return [
      {text: s.name, lines: 3, bold: true, color: c.ink, fill: c.greenSoft},
      // 件数：この資料のリクエスト（サブインダストリーの資料で、ほかのサブインダストリーにもあれば、下に全体の件数）
      s.allRequests.length > s.requests.length
        ? {text: s.requests.length + '件\n全体' + s.allRequests.length, lines: 2, color: c.greenDark, fill: c.white, align: 'CENTER',
          parts: [{bold: true, size: 10}, {size: 7.5, color: c.muted}]}
        : {text: s.requests.length + '件', size: 10, bold: true, color: c.greenDark, fill: c.white, align: 'CENTER'},
      {text: s.customers.length + '社\n' + s.customers.join('・'), lines: 3, fill: c.white, parts: [{bold: true}, {}]},
      smtgTextCell_(s.subIndustries.join('・'), {lines: 3}),
      data.hasWeighted ? {text: s.active.length ? mtgAmount_(s.weighted) + '\n（' + s.active.length + '案件）' : '進行中の\n案件なし', lines: 2,
        align: 'END', fill: c.white, color: s.active.length ? c.body : c.muted} : mtgGapCell_({}),
      smtgTextCell_(s.owners.join('\n'), {lines: 3}),
      n.total ? {text: '全' + n.total + '件（完了' + n.done + '）\n対応中' + n.doing + '・未着手' + n.todo, lines: 3, fill: c.white,
        parts: [{bold: true}, {}]} : {text: 'タスクなし', color: c.muted, fill: c.white},
      {text: starts[i] ? String(starts[i]) : '', bold: true, color: c.greenDark, fill: c.white, align: 'CENTER'}
    ];
  });
  const lead = (data.all ? 'サービス案 ' + data.services.length + '件。'
    : 'このサブインダストリーのリクエストがあるサービス案 ' + data.services.length + '件（件数・得意先・期待値はこのサブインダストリーの分）。') +
    'リクエストの多い順。\n期待値は、元になった案件（進行中）の期待値調整済の想定売上の合計（案件全体の値で、サービスの売上ではありません）。';
  return smtgPages_({title: SMTG_TITLES.services, lead, sections: [{columns, rows,
    head: ['サービス案', '件数', '得意先', 'サブインダストリー', '元の案件の\n期待値', '主管本部・担当者', 'タスク', 'ページ'],
    empty: {kind: 'none', text: '「' + SVC_OPTIONS.requestSheet + '」の「サービス案」に名前を入れたリクエストがありません（同じ名前のリクエストが1つのサービス案になります）。'}}]});
}

/** 3 顧客・業界の広がり：サービス案 × サブインダストリーのリクエストの件数（1件以上のセルは薄い紫に太字） */
function smtgFitSlides_(data) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const fit = data.fit;
  const n = fit.columns.length;
  const cw = n ? Math.min(60, Math.floor((S.w - 150 - 44) / n)) : 0;
  const columns = mtgColumns_([0].concat(fit.columns.map(() => cw), [44]), 0);
  const multi = fit.rows.filter(r => r.counts.filter(Boolean).length > 1).length;
  const rows = fit.rows.map(r => [{text: r.service.name, lines: 2, bold: true, color: c.ink, fill: c.greenSoft}]
    .concat(r.counts.map(v => (v ? {text: String(v), bold: true, color: c.greenDark, fill: c.greenSoft, align: 'CENTER'}
      : {text: '', fill: c.white})), [{text: String(r.total), bold: true, color: c.ink, fill: c.white, align: 'CENTER'}]));
  const lead = 'サービス案ごとの、リクエストのあったサブインダストリー（数字はリクエストの件数）。' +
    (data.services.length ? '\n2つ以上のサブインダストリーにまたがるサービス案 ' + multi + '件。' : '') +
    (n >= SMTG_OPTIONS.fitColumns ? 'サブインダストリーはリクエストの多い順に ' + (SMTG_OPTIONS.fitColumns - 1) + 'つまで（残りは「その他」）。' : '');
  // 見出しはサブインダストリーの名前が切れないよう、4行まで・小さめの字（7.5pt）
  // 1行に入らない名前は、まん中に近い「・」で改行する（「化粧品・ト／イレタリー」のように言葉の途中で切らない）
  const split = text => {
    if (mtgWrap_(text, cw, 7.5, 9).length <= 1 || text.indexOf('・') < 0) return text;
    const at = Array.from(text.matchAll(/・/g)).map(m => m.index + 1).sort((a, b) => Math.abs(a - text.length / 2) - Math.abs(b - text.length / 2))[0];
    return text.slice(0, at) + '\n' + text.slice(at);
  };
  const labels = ['サービス案'].concat(fit.columns.map(split), ['合計']);
  const headRow = mtgRow_(columns, labels.map(text => ({text, size: 7.5, bold: true, color: c.white, fill: c.green, align: 'CENTER', lines: 4})),
    S.headRow);
  return smtgPages_({title: SMTG_TITLES.fit, lead, sections: [{columns, rows, head: labels, headRow,
    empty: {kind: 'none', text: 'サービス案がありません。'}}]});
}

/**
 * 4 サービス案（1つずつ）：元になったリクエスト（ID・得意先と案件名・サブインダストリー・リクエスト・サービス部門からのフィードバック・案件のステータスと期待値）と、
 * そのサービス案のタスク（タスク・本部と担当者・状況・メモ）。長い文は全文入れる（小さくし、入らなければ行を伸ばす）。入りきらなければ続きのページ
 */
function smtgDetailSlides_(data, s) {
  const c = MTG_OPTIONS.colors;
  const n = s.counts;
  const reqColumns = mtgColumns_([64, 132, 0, 150, 96], 2);
  const taskColumns = mtgColumns_([0, 120, 64, 230], 0);
  const statuses = smtgStatusOptions_();
  const requests = s.requests.map(r => [
    {text: r.id || '（IDなし）', bold: true, color: c.ink, fill: c.greenSoft, align: 'CENTER'},
    smtgCustomerCell_(r, reqColumns[1]),
    {text: r.request, lines: 4, fit: true, grow: true, fill: c.white},
    smtgTextCell_(r.feedback, {lines: 4, fit: true, grow: true}),
    smtgDealCell_(data, r, reqColumns[4])
  ]);
  const tasks = s.tasks.map(t => [
    smtgTextCell_(t.task, {lines: 3, fit: true, grow: true}),
    smtgTextCell_([t.department, t.owner].filter(Boolean).join(' '), {lines: 2}),
    {text: '', fill: c.white},
    smtgTextCell_(t.note, {lines: 3, fit: true, grow: true})
  ]);
  const amount = data.hasWeighted ? '元の案件の期待値 ' + mtgAmount_(s.weighted) + '百万円（進行中 ' + s.active.length + '案件）' : '';
  const others = s.allRequests.length - s.requests.length;
  const lead = (data.all ? 'リクエスト ' + s.requests.length + '件（得意先 ' + s.customers.length + '社・' + (s.subIndustries.join('・') || 'サブインダストリー未入力') + '）'
    : 'リクエスト ' + s.requests.length + '件' + (others ? '（ほかのサブインダストリーに ' + others + '件）' : '') + '・得意先 ' + s.customers.length + '社') +
    (amount ? '、' + amount : '') + '。\nタスク ' + n.total + '件（完了 ' + n.done + '・対応中 ' + n.doing + '・未着手 ' + n.todo + '）。主管本部・担当者：' +
    (s.owners.join('、') || MTG_OPTIONS.blankText) + '。';
  return smtgPages_({title: SMTG_TITLES.detail + '：' + s.name, lead, shrink: true, sections: [
    {heading: '元になったリクエスト', columns: reqColumns, rows: requests,
      head: ['ID', '得意先・案件名\n（サブインダストリー）', 'リクエスト', 'サービス部門からのフィードバック', '案件のステータス・期待値']},
    {heading: 'タスク', columns: taskColumns, rows: tasks, head: ['タスク', '本部・担当者', '状況', 'メモ'],
      pills: {col: 2, labels: i => smtgStatusPill_(s.tasks[i]), width: smtgPillWidth_(statuses, taskColumns[2])},
      empty: {kind: 'none', text: 'このサービス案のタスクは、まだありません（「' + TASK_OPTIONS.sheet + '」に追加します）。'}}
  ]});
}

/** 5 判断待ちのリクエスト：判断が「未判断」（空も）のリクエスト（ID の順。前回以降に登録したものは ID の下に「新規」） */
function smtgPendingSlides_(data) {
  const c = MTG_OPTIONS.colors;
  const columns = mtgColumns_([64, 132, 0, 120, 96], 2);
  const fresh = new Set(data.moves.newRequests);
  const list = data.requests.filter(r => r.kind === 'undecided').sort((a, b) => (a.number || Infinity) - (b.number || Infinity));
  const rows = list.map(r => [
    {text: (r.id || '（IDなし）') + (fresh.has(r) ? '\n新規' : ''), bold: true, color: c.ink, fill: c.greenSoft, align: 'CENTER', lines: 2,
      parts: [{}, {color: c.greenDark, size: 8}]},
    smtgCustomerCell_(r, columns[1]),
    {text: r.request, lines: 4, fit: true, grow: true, fill: c.white},
    smtgTextCell_(smtgSalesText_(r), {lines: 3}),
    smtgDealCell_(data, r, columns[4])
  ]);
  const lead = '判断が「' + smtgDecisionOptions_()[0] + '」のリクエスト ' + list.length + '件' +
    (data.previous ? '（うち前回以降に登録 ' + list.filter(r => fresh.has(r)).length + '件）' : '') + '。ID の順。';
  return smtgPages_({title: SMTG_TITLES.pending, lead, sections: [{columns, rows,
    head: ['ID', '得意先・案件名\n（サブインダストリー）', 'リクエスト', '営業', '案件のステータス・期待値'],
    empty: {kind: 'none', text: 'すべてのリクエストに判断が入っています。'}}]});
}

/** Appendix ① リクエスト（一覧。棄却も。ID の順） */
function smtgRequestsAppendix_(data) {
  const c = MTG_OPTIONS.colors;
  const columns = mtgColumns_([64, 84, 104, 0, 92, 100, 112], 3);
  const list = data.requests.slice().sort((a, b) => (a.number || Infinity) - (b.number || Infinity) || a.row - b.row);
  const rows = list.map(r => [
    {text: r.id || '（IDなし）', bold: true, color: c.ink, fill: c.greenSoft, align: 'CENTER'},
    smtgTextCell_(r.subIndustry, {lines: 2}),
    Object.assign(mtgDealCell_({customer: r.customer, project: r.project}, columns[2], MTG_SLIDE.appendixSize), {fill: c.white}),
    {text: r.request, lines: 3, fit: true, fill: c.white},
    {text: '', fill: c.white},
    smtgTextCell_(r.service, {lines: 2}),
    smtgTextCell_(r.feedback, {lines: 3, fit: true})
  ]);
  const k = data.kpi;
  const lead = '登録リクエスト ' + k.requests + '件（' + smtgDecisionOptions_().map(d => d + ' ' + data.requests.filter(r => r.decision === d).length + '件')
    .join('・') + '）。ID の順。';
  return smtgPages_({title: SMTG_TITLES.requests, lead, size: MTG_SLIDE.appendixSize, sections: [{columns, rows,
    head: ['ID', 'サブ\nインダストリー', '得意先・案件名', 'リクエスト', '判断', 'サービス案', 'サービス部門からのフィードバック'],
    pills: {col: 4, labels: i => smtgDecisionPill_(list[i]), width: smtgPillWidth_(smtgDecisionOptions_(), columns[4])},
    empty: {kind: 'none', text: '「' + SVC_OPTIONS.requestSheet + '」にリクエストが登録されていません。'}}]});
}

/** Appendix ② タスク（一覧。サービス案の並び（一覧と同じ）→ シートの行の順。サービス案に結び付かないタスクは最後） */
function smtgTasksAppendix_(data) {
  const c = MTG_OPTIONS.colors;
  const columns = mtgColumns_([150, 0, 110, 62, 200], 1);
  const order = new Map(data.services.map((s, i) => [s.key, i]));
  const rank = t => (t.request && order.has(t.request.serviceKey) ? order.get(t.request.serviceKey) : Infinity);
  const list = data.tasks.slice().sort((a, b) => (rank(a) - rank(b)) || (a.row - b.row));
  const rows = list.map(t => [
    {text: (t.serviceName || '（サービス案なし）') + '\n' + (t.request ? t.request.id : t.label || '（サービス案未入力）'), lines: 3, color: c.ink,
      fill: c.greenSoft, parts: [{bold: true}, {color: c.muted}]},
    smtgTextCell_(t.task, {lines: 3, fit: true}),
    smtgTextCell_([t.department, t.owner].filter(Boolean).join(' '), {lines: 2}),
    {text: '', fill: c.white},
    smtgTextCell_(t.note, {lines: 3, fit: true})
  ]);
  const n = smtgStatusCounts_(data.tasks);
  const lead = 'タスク ' + n.total + '件（完了 ' + n.done + '・対応中 ' + n.doing + '・未着手 ' + n.todo + (n.blank ? '・状況未入力 ' + n.blank : '') +
    '）。サービス案の一覧の順。';
  return smtgPages_({title: SMTG_TITLES.tasks, lead, size: MTG_SLIDE.appendixSize, sections: [{columns, rows,
    head: ['サービス案（リクエストの ID）', 'タスク', '本部・担当者', '状況', 'メモ'],
    pills: {col: 3, labels: i => smtgStatusPill_(list[i]), width: smtgPillWidth_(smtgStatusOptions_(), columns[3])},
    empty: {kind: 'none', text: '「' + TASK_OPTIONS.sheet + '」にタスクがありません。'}}]});
}

/** Appendix ③ 入力内容の確認事項（一覧。種類の順） */
function smtgChecksAppendix_(data) {
  const c = MTG_OPTIONS.colors;
  const columns = mtgColumns_([140, 220, 0], 2);
  const rows = data.checks.map(item => [
    {text: item.type, bold: true, color: c.ink, fill: c.greenSoft, lines: 2},
    {text: item.target, lines: 2, fill: c.white},
    {text: item.text, lines: 3, fill: c.white}
  ]);
  const types = SMTG_CHECKS.map(type => [type, data.checks.filter(x => x.type === type).length]).filter(x => x[1]);
  const lead = '確認事項 ' + data.checks.length + '件' + (types.length ? '（' + types.map(x => x[0] + ' ' + x[1]).join('・') + '）' : '') + '。' +
    (data.dealGaps.length ? '\n案件を読めなかったシート：' + data.dealGaps.join('、') : '');
  return smtgPages_({title: SMTG_TITLES.checks, lead, size: MTG_SLIDE.appendixSize, sections: [{columns, rows, head: ['種類', '対象', '内容'],
    empty: {kind: 'none', text: '確認事項はありません。'}}]});
}
