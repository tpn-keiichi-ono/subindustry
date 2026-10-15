/**
 * 会議サマリ（追跡シートごとの会議資料。16:9 のスライドと PDF）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs・HistorySidebar.gs と同じプロジェクトに置くファイル。
 *
 * 目的
 * - RX Business MTG（パイプライン会議。月1回）の会議資料。追跡シート（スーパー・GMS など。DIFF_RULES のシート）ごとの推進リーダーが説明に使い、
 *   本部長クラスがレビューする。発表は1シートあたり数分（報告＋フィードバック）なので、参加者は事前に通読してから会議に臨む。
 *   全リストをスクロールしなくても、全体感・前回からの動き・重点案件の状況と課題・不備をつかめるようにする。
 * - 「動いた案件・止まった案件・不備のある案件」のふるい分けに絞る。個別の案件を深く見るのは、オポチュニティキャンバスなど別の画面で行う。
 * - 文はすべてシートと変更の記録から作る（記入欄は無い）。スライドの見出しはどのシートでも同じ（MTG_TITLES）で、状態はリード文に書く。
 *
 * 資料の構成（16:9。社内資料のトンマナ：左上の紫のタグ・紫の見出し・リード文・表・赤枠・結論の帯・ページ番号）
 *   シートごとの資料（会議の議題の順）
 *   1 サマリー：今期の受注と着地見込み・4つの数字（期待値と前回比・受注／失注・前進／後退・停滞）・本日の論点
 *     （論点 ▶ 状況 ▶ 決めてほしいこと）
 *   2 パイプラインの変化：前回からの期待値の増減（新規・増額・復帰・減額・受注・失注・保留・削除）・ステータス別・受注月（四半期）別の期待値
 *   3 前回からの進捗：受注・失注・後退・受注月の延期などの大きな動き（上位。悪い動きは赤枠）
 *   4 重点案件の状況・課題：★の案件と期待値の大きい進行中の案件（支援の要否を決めたい案件は赤枠。★はオポチュニティキャンバスで深掘りする候補）
 *   付録（全件。入りきらなければ続きのページ）
 *   A 前回からの変化　B 停滞している案件　C 要対応（責任者別・一覧）　D ステータスの分類
 *   全体サマリー（追跡シートが2枚以上のとき。すべての追跡シートをまとめる。事務局が会議の冒頭で報告する）
 *   1 全体サマリー：シートごとの数字と入力状況（前回から更新した案件・最終更新日時）　2 本日の論点（サブインダストリー別）
 *
 * ことばの決まり（どのページでも同じ数え方にする）
 * - ステータスの分類：進行中・受注・失注・保留（MTG_OPTIONS のことば・statusKinds で決める）。期待値・パイプラインは進行中だけを数える。
 *   前進・後退は、案件ステータスのプルダウンの並び順で決める。
 * - 前回からの変化は、案件ごとに1つの区分にする（新規 > 受注 > 失注 > 保留 > 復帰 > 前進 > 後退 > 延期 > 前倒し > 増額 > 減額 > 変更 > 更新のみ）。
 *   行を削除した案件は「削除」（失注には数えない）。
 * - 停滞：進行中で、ステータスが最後に変わってから stallDays 日以上たった案件（記録より前から変わっていなければ「N日以上」）。
 * - 重点案件：★の案件を先に入れ、期待値の大きい進行中の案件で mainRows.focus 件まで。支援の要否を決めたい案件（停滞・受注月超過・
 *   受注月の延期が canvas.slips 回以上・課題に canvas.words の語）は赤枠。
 *
 * 仕組み
 * - 管理者がエディタで createMeetingSummaries() を実行し、作るシートと今回の会議日を入力する（試すときは previewMeetingSummaries()）。
 *   シートごとの Google スライドと PDF、全体サマリーをドライブのフォルダに保存する（MTG_OPTIONS.folderId。空なら、このスプレッドシートと
 *   同じフォルダの中に MTG_OPTIONS.folderName のフォルダを作る）。配る（共有する）のは手動。
 * - 「前回から」は、前回の会議（今回の会議日より前で、いちばん新しい会議日）の資料を作った日時から。会議日と作った日時は
 *   シートごとにドキュメントのプロパティ（MTG_MEETINGS_シート名）に残す（同じ会議日で作り直したら、作った日時を上書きする）。
 *   前回の会議の記録が無ければ、MTG_OPTIONS.firstDays 日前から。previewMeetingSummaries() は会議日を記録しないので、何度でも試せる
 *   （ファイル名に「試作」が付く）。
 * - 会議日ごとに最初の試作の要対応を覚えておき（MTG_DRAFT_シート名）、同じ会議日の資料で「試作から解消した件数」を出す
 *   （試作を責任者に見せて、会議の前に直してもらう運用）。
 * - スプレッドシートは読むだけで、書き込まない。変更の記録は、変更時点スナップショット（その時の行の内容）で得意先＋案件名に結び付ける
 *   （行の追加・削除で行番号がずれても合う。得意先・案件名を書き換えた案件・あとから入れた案件も、前の名前の記録をつなげる）。
 *   どちらの記録も追記なので、下（新しいほう）から、読む期間（前回と MTG_OPTIONS.lookbackDays 日前の古いほう）より前に届くまでだけ読む。
 * - スライドは Slides API（拡張サービス「Google Slides API」。appsscript.json に書いてある）で作り、PDF に書き出す。
 */

const MTG_OPTIONS = {
  title: 'RXビジネスMTG',             // 左上のタグ（「タグ｜シート名（会議日）」と出す）
  fileName: '会議サマリ',             // ファイル名の先頭（会議サマリ_スーパー・GMS_2026-10-23.pdf。日付は会議日）
  folderId: '',                      // 保存先のフォルダの ID（空なら、このスプレッドシートと同じフォルダの中に folderName を作る）
  folderName: '会議サマリ',
  folderKey: 'MTG_FOLDER_ID',         // 作ったフォルダの ID（ドキュメントのプロパティ）
  meetingsKeyPrefix: 'MTG_MEETINGS_', // シートごとの会議日と資料を作った日時（ドキュメントのプロパティ。後ろにシート名）
  draftKeyPrefix: 'MTG_DRAFT_',       // シートごとの、会議日ごとの最初の試作の要対応（解消した件数を出す用）
  lastKeyPrefix: 'MTG_LAST_',         // 前の版の「前回の作成日時」（会議日の記録が無いときだけ使う）
  keepMeetings: 24,                  // 会議日の記録をいくつ残すか
  keepSlides: true,                  // Google スライドもフォルダに残す（推進リーダーが書き足せる。false なら PDF だけ）
  firstDays: 14,                     // 初回（前回の会議の記録が無い）は、何日前からの変化を出すか
  lookbackDays: 180,                 // ステータスの滞留・受注月の延期の回数を数えるために、何日前までの記録を読むか
  stallDays: 45,                     // 停滞：ステータスがこの日数以上変わっていない進行中の案件
  noUpdateDays: 30,                  // 要対応「更新なし」：最終更新日時がこの日数より前の進行中の案件
  overdueDays: 90,                   // 要対応「予定超過」：予定日を過ぎてこの日数以内のもの（古いものは出さない）
  // ★（オポチュニティキャンバスで深掘りする候補）：停滞の日数・受注月の延期の回数・活動における課題に含まれる語のどれかに当たる案件の、
  // 想定の大きい max 件
  canvas: {days: 60, slips: 2, words: ['決裁者', '上席'], max: 3},
  fiscalStartMonth: 4,               // 期の始まりの月（今期の受注・着地見込み・四半期）
  targets: {},                       // シートごとの今期の受注の目標（百万円）。例 {'スーパー・GMS': 500}。空なら達成率を出さない
  // 読む列の見出し（改行・空白、全角・半角の違いは無視。customer・project・status はどのシートにも必要。ほかは無ければ使わない）
  columns: {
    customer: '得意先',
    project: '案件名',
    status: '案件ステータス',
    amount: '想定売上規模（百万）',
    weighted: '期待値調整済\n想定売上規模（百万）',
    orderMonth: '受注月',
    startMonth: '売上開始月',
    activity: '活動状況',
    issue: '活動における課題',
    owner: 'アカウント責任者',
    delivery: 'デリバリー担当'
  },
  // 必須項目：進行中の案件で空なら要対応に出す（シートに無い見出しは見ない）
  required: ['想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月', '活動状況', 'アカウント責任者'],
  // ステータスの分類。statusKinds に書いたステータスはそのとおりに（'active' 進行中・'won' 受注・'lost' 失注・'hold' 保留）。
  // ほかは、その語を含むかで決める（失注 → 保留 → 受注 の順に見る。受注は notWonWords を含むもの（受注見込み など）を除く）
  statusKinds: {},
  lostWords: ['失注', '辞退', '見送', '中止', 'クローズ', 'ロスト', '消滅'],
  holdWords: ['保留', 'ペンディング', '凍結', '休眠'],
  wonWords: ['受注', '成約', '契約済'],
  notWonWords: ['見込', '予定', '前', '待'],
  logChunkRows: 2000,                // 変更の記録を下から読むときの1回の行数
  batchRequests: 1500,               // Slides API の1回の batchUpdate に入れるリクエストの数の目安（呼ぶ回数を減らす。1分あたりの回数に上限がある）
  mainRows: {changes: 7, focus: 5, owners: 9},     // 本編の表に出す行の数（残りは付録）。owners は付録C の責任者別の表の行
  overallName: '全体',               // 全体サマリー（すべての追跡シートを1枚の表にした資料）のファイル名に入れる名前
  // 色（社内資料のトンマナ：紫を基調に、薄い紫・薄いグレーの面、強調は赤枠、結論はピンクの帯）
  colors: {
    purple: '#7030A0', purpleSoft: '#EEE3F5', up: '#A77BCA', down: '#BFBFBF', gray: '#F2F2F2', grayHead: '#A6A6A6',
    arrow: '#A6A6A6', red: '#E60012', highlight: '#F6D5F7', ink: '#1B1B1C', body: '#262626', muted: '#595959', faint: '#8C8C8C',
    white: '#FFFFFF'
  }
};

/* 前回からの変化の区分（上ほど優先。案件ごとに1つ）。bad は赤枠で囲む悪い動き */
const MTG_CATEGORIES = [
  {key: 'new', label: '新規'},
  {key: 'won', label: '受注'},
  {key: 'lost', label: '失注', bad: true},
  {key: 'hold', label: '保留'},
  {key: 'revived', label: '復帰'},
  {key: 'forward', label: '前進'},
  {key: 'backward', label: '後退', bad: true},
  {key: 'slipped', label: '延期', bad: true},
  {key: 'pulled', label: '前倒し'},
  {key: 'up', label: '増額'},
  {key: 'down', label: '減額'},
  {key: 'status', label: '変更'},
  {key: 'removed', label: '削除'},
  {key: 'updated', label: '更新のみ'}
];

/*
 * スライドの見出し（タイトル）。どのサブインダストリーでも同じにする（状態はリード文に書く）。
 * RX Business MTG（パイプライン会議）の議題の順：サマリー → パイプラインの変化 → 前回からの進捗 → 重点案件の状況・課題。付録は全件の一覧
 */
const MTG_TITLES = {
  overall: '全体サマリー',
  overallIssues: '本日の論点（サブインダストリー別）',
  summary: 'サマリー',
  pipeline: 'パイプラインの変化',
  progress: '前回からの進捗',
  focus: '重点案件の状況・課題',
  changes: '付録A　前回からの変化（一覧）',
  stalled: '付録B　停滞している案件（一覧）',
  owners: '付録C　要対応（責任者別）',
  alerts: '付録C　要対応（一覧）',
  status: '付録D　ステータスの分類'
};

/* 本編の「前回からの進捗」に出す順（更新のみは出さない） */
const MTG_MOVE_ORDER = ['won', 'lost', 'backward', 'slipped', 'forward', 'new', 'revived', 'hold', 'removed', 'pulled', 'up', 'down', 'status'];

/* 期待値の増減の内訳（増える → 減るの順） */
const MTG_BRIDGE = [
  {key: 'new', label: '新規'}, {key: 'up', label: '増額'}, {key: 'revived', label: '復帰'},
  {key: 'down', label: '減額'}, {key: 'won', label: '受注'}, {key: 'lost', label: '失注'}, {key: 'hold', label: '保留'},
  {key: 'removed', label: '削除'}
];

/* 要対応の種類（この順に出す） */
const MTG_ALERT_TYPES = ['予定超過', '受注月超過', '未入力', '更新なし', '金額の矛盾', '日付の矛盾', '未定義ステータス', '体制未登録'];

/*
 * スライドの形（pt。16:9 = 720 × 405）。本編は上から タグ・タイトル・リード文・中身（top〜bottom）・結論の帯・出典とページ番号。
 * 文字の幅と行の高さは見積もり（mtgWrap_・mtgLineHeight_）。本物より大きめに見積もって、はみ出さないようにする。
 */
const MTG_SLIDE = {
  width: 720,
  height: 405,
  font: 'Noto Sans JP',
  x: 24,                 // 左の余白
  w: 672,                // 中身の幅
  top: 98,               // 中身の上端
  bottom: 348,           // 本編の中身の下端（下に結論の帯）
  appendixBottom: 378,   // 付録の中身の下端（帯なし）
  titleSize: 20,
  leadSize: 9.5,
  bodySize: 8,
  appendixSize: 7.5,
  lineFactor: 1.5,       // 行の高さ = 文字の大きさ × lineFactor
  padX: 14.4,            // 箱・表のセルの左右の余白（合わせて）
  padY: 7.2,             // 上下の余白（合わせて）
  headRow: 18            // 表の見出しの行の高さ（1行）
};

/* ---------------- エディタから実行する ---------------- */

/**
 * 会議サマリ（Google スライドと PDF）を作って保存する（作るシートと今回の会議日を聞く）。
 * 会議日と作った日時を記録し、次の会議の資料は今以降の変化を出す。
 */
function createMeetingSummaries() {
  mtgRun_(true);
}

/** 試しに会議サマリを作る（保存するが、会議日は記録しない。ファイル名に「試作」が付く）。 */
function previewMeetingSummaries() {
  mtgRun_(false);
}

function mtgRun_(final) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const names = mtgAskSheets_(ss, ui, final);
  if (!names) return;
  const now = new Date();
  const meeting = mtgAskMeeting_(ui, names, now);
  if (!meeting) return;
  const created = mtgCreate_(ss, names, now, final, meeting);
  const lines = created.results.map(r => '「' + r.summary.name + '」：' + r.file.getName() + '（' + mtgCountsText_(r.summary) + '）');
  if (created.overall) lines.push('全体サマリー：' + created.overall.pdf.getName() + '（追跡シート ' + created.overall.sheets + '枚）');
  lines.push('', '保存先：' + created.folder.getUrl());
  if (MTG_OPTIONS.keepSlides) lines.push('同じフォルダに、Google スライド（同じ名前）も保存しました。');
  lines.push(final
    ? '会議日（' + meeting + '）を記録しました。次の会議の資料は、今以降の変化を出します。配るときは、保存先のフォルダから共有してください。'
    : '試作のため、会議日は記録していません。');
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/** 作るシートを聞く（空欄ならすべて。「1,3」のように複数も）。やめたら null。 */
function mtgAskSheets_(ss, ui, final) {
  const all = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  if (!all.length) throw new Error('追跡しているシート（test.gs の DIFF_RULES）が見つかりません。');
  const message = (final ? '会議サマリを作ります。' : '会議サマリを試しに作ります（会議日は記録しません）。') +
    '\n作るシートの番号を入力してください（空欄ならすべて。複数は「1,3」のように）。\n\n' +
    all.map((name, i) => (i + 1) + '. ' + name).join('\n');
  const response = ui.prompt(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const text = credNormalize_(response.getResponseText());
  if (!text) return all;
  const picked = [];
  text.split(/[,、・]/).filter(Boolean).forEach(part => {
    const index = Number(part) - 1;
    if (!(index >= 0 && index < all.length)) throw new Error('「' + part + '」に当たるシートがありません。1〜' + all.length + ' の番号で入力してください。');
    if (picked.indexOf(all[index]) < 0) picked.push(all[index]);
  });
  return picked;
}

/** 今回の会議日を聞く（「10/23」「2026/10/23」）。「yyyy/MM/dd」を返す。やめたら null。 */
function mtgAskMeeting_(ui, names, now) {
  const props = PropertiesService.getDocumentProperties();
  const last = names.map(name => {
    const list = mtgMeetings_(props, name);
    return '・' + name + '：' + (list.length ? list[list.length - 1].date : '記録なし');
  });
  const message = '今回の会議日を入力してください（例：10/23）。\n' +
    '前回の会議（今回より前で、いちばん新しい会議日）の資料を作った時点からの変化を出します。\n\n記録してある最後の会議日\n' + last.join('\n');
  const response = ui.prompt(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const text = String(response.getResponseText() || '').trim();
  if (!text) throw new Error('会議日を入力してください（例：10/23）。');
  const date = mtgParseDate_(text, now);
  if (!date) throw new Error('「' + text + '」を会議日として読めません。10/23 か 2026/10/23 のように入力してください。');
  return date;
}

/** 完成した結果の件数の知らせ（「動いた 12・停滞 5・要対応 7」） */
function mtgCountsText_(summary) {
  return '動いた ' + summary.kpi.changed + '・停滞 ' + summary.stalled.length + '・要対応 ' + summary.alerts.length;
}

/**
 * names のシートの会議サマリを作り、スライドと PDF を保存する。meeting は今回の会議日（yyyy/MM/dd）。
 * 追跡シートが2枚以上あれば、すべての追跡シート（names 以外も）をまとめた全体サマリーも作る（事務局のサマリ報告用）。
 * final なら names のシートに会議日と今の日時を記録する。試作なら、その会議日の最初の試作の要対応を覚える。
 * 戻り値 {results: [{summary, file: PDF, deck: Google スライド（残さないときは null）}], overall: {pdf, deck, sheets}|null, folder}
 */
function mtgCreate_(ss, names, now, final, meeting) {
  const o = MTG_OPTIONS;
  const props = PropertiesService.getDocumentProperties();
  const all = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  const targets = all.length > 1 ? all : names;
  const plans = targets.map(name => {
    const meetings = mtgMeetings_(props, name);
    const previous = meetings.filter(m => m.date < meeting).pop() || null;
    const legacy = previous ? 0 : Number(props.getProperty(o.lastKeyPrefix + name)) || 0;
    const since = previous ? previous.cutoff : (legacy || now.getTime() - o.firstDays * 86400000);
    const draft = mtgReadJson_(props, o.draftKeyPrefix + name);
    return {name, selected: names.indexOf(name) >= 0, meetings, previous, legacy: legacy || null, since: new Date(since),
      draft: draft && draft.meeting === meeting && Array.isArray(draft.keys) ? draft : null};
  });
  const windowStart = new Date(Math.min.apply(null,
    plans.map(p => p.since.getTime()).concat([now.getTime() - o.lookbackDays * 86400000])));
  const context = {
    now,
    meeting,
    windowStart,
    trackedSince: mtgLogStart_(ss),
    log: mtgReadLog_(ss, windowStart),
    snapshots: mtgReadSnapshots_(ss, windowStart),
    credentials: hsCredentialIndex_(ss)
  };
  const summaries = plans.map(p => mtgSummarize_(ss, p.name, Object.assign({
    since: p.since, previous: p.previous, legacy: p.legacy, draft: p.draft, final
  }, context)));
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  const results = [];
  plans.forEach((p, i) => {
    if (!p.selected) return;
    const out = mtgExportDeck_(mtgFileName_(p.name, meeting, final), mtgSlides_(summaries[i]), folder);
    results.push({summary: summaries[i], file: out.pdf, deck: out.deck});
  });
  let overall = null;
  if (summaries.length > 1) {
    overall = mtgExportDeck_(mtgFileName_(o.overallName, meeting, final), mtgOverallSlides_(summaries, {meeting, final, now}), folder);
    overall.sheets = summaries.length;
  }
  plans.forEach((p, i) => {
    if (!p.selected) return;
    if (final) {
      const list = p.meetings.filter(m => m.date !== meeting).concat([{date: meeting, cutoff: now.getTime()}])
        .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).slice(-o.keepMeetings);
      props.setProperty(o.meetingsKeyPrefix + p.name, JSON.stringify(list));
    } else if (!p.draft) {
      props.setProperty(o.draftKeyPrefix + p.name, JSON.stringify({
        meeting, at: now.getTime(), keys: summaries[i].alerts.map(a => mtgHash_(a.key))
      }));
    }
  });
  return {results, overall, folder};
}

/** シートの会議の記録 [{date: 'yyyy/MM/dd', cutoff: 資料を作った時刻（ms）}]（会議日の古い順） */
function mtgMeetings_(props, name) {
  const list = mtgReadJson_(props, MTG_OPTIONS.meetingsKeyPrefix + name);
  return (Array.isArray(list) ? list : [])
    .filter(m => m && /^\d{4}\/\d{2}\/\d{2}$/.test(m.date) && typeof m.cutoff === 'number')
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}

function mtgReadJson_(props, key) {
  try {
    return JSON.parse(props.getProperty(key) || 'null');
  } catch (_) {
    return null;
  }
}

/**
 * 会議日の文字を「yyyy/MM/dd」に（「10/23」「10月23日」は、今日から半年より前になるなら来年とみなす）。読めなければ空。
 */
function mtgParseDate_(text, now) {
  const t = String(text || '').normalize('NFKC').trim();
  const full = /^(\d{4})\s*[\/\-.年]\s*(\d{1,2})\s*[\/\-.月]\s*(\d{1,2})\s*日?$/.exec(t);
  const short = /^(\d{1,2})\s*[\/\-.月]\s*(\d{1,2})\s*日?$/.exec(t);
  const pad = n => String(n).padStart(2, '0');
  let y;
  let m;
  let d;
  if (full) {
    y = Number(full[1]); m = Number(full[2]); d = Number(full[3]);
  } else if (short) {
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    y = Number(today.slice(0, 4)); m = Number(short[1]); d = Number(short[2]);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31 && mtgDay_(y + '/' + pad(m) + '/' + pad(d)) < mtgDay_(today) - 183 * 86400000) y++;
  } else {
    return '';
  }
  if (m < 1 || m > 12 || d < 1 || new Date(Date.UTC(y, m - 1, d)).getUTCDate() !== d) return '';
  return y + '/' + pad(m) + '/' + pad(d);
}

/** 短い文字のハッシュ（要対応を覚えておく用。プロパティの大きさを抑える） */
function mtgHash_(text) {
  let h = 2166136261;
  Array.from(String(text)).forEach(ch => {
    h ^= ch.codePointAt(0);
    h = Math.imul(h, 16777619) >>> 0;
  });
  return h.toString(36);
}

/* ---------------- 読む ---------------- */

/**
 * 記録のシート（変更履歴_差分・変更時点スナップショット）を下から読み、since 以降の行を返す（古い順）。
 * 記録は追記なので、サンプル（イベントID が SMP-。過去の日時で足される）以外で since より前の行に届いたら止める。
 * {index: 見出し → 列の位置（0 から）, rows}
 */
function mtgReadRecent_(sheet, since) {
  const out = {index: {}, rows: []};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const width = sheet.getLastColumn();
  sheet.getRange(1, 1, 1, width).getDisplayValues()[0].forEach((h, i) => { out.index[String(h).trim()] = i; });
  const atCol = out.index['記録日時'];
  const idCol = out.index['イベントID'];
  if (atCol == null) return out;
  const limit = since.getTime();
  const step = MTG_OPTIONS.logChunkRows;
  for (let end = sheet.getLastRow(); end >= 2; end -= step) {
    const start = Math.max(2, end - step + 1);
    const rows = sheet.getRange(start, 1, end - start + 1, width).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      const sample = idCol != null && /^SMP-/.test(String(rows[i][idCol]));
      const at = mtgTime_(rows[i][atCol]);
      if (sample) continue;
      if (at != null && at < limit) return Object.assign(out, {rows: out.rows.reverse()});
      if (at != null) out.rows.push(rows[i]);
    }
  }
  out.rows.reverse();
  return out;
}

/** 変更履歴_差分 の since 以降の記録 [{at, eventId, sheet, cell, column, type, before, after}]（古い順） */
function mtgReadLog_(ss, since) {
  const read = mtgReadRecent_(ss.getSheetByName(DIFF_OPTIONS.logSheet), since);
  const i = read.index;
  const need = ['記録日時', 'イベントID', 'シート', 'セル', '列名', '種類', '変更前', '変更後'];
  if (read.rows.length && need.some(h => i[h] == null)) {
    throw new Error('「' + DIFF_OPTIONS.logSheet + '」の見出しが今の形ではありません。エディタで compactDiffRecords() を実行してから、もう一度お試しください。');
  }
  const text = v => mtgCellText_(v);
  return read.rows.map(r => ({
    at: mtgTime_(r[i['記録日時']]),
    eventId: text(r[i['イベントID']]),
    sheet: text(r[i['シート']]),
    cell: text(r[i['セル']]),
    column: text(r[i['列名']]),
    type: text(r[i['種類']]),
    before: text(r[i['変更前']]),
    after: text(r[i['変更後']])
  }));
}

/** 変更の記録を始めた時刻（変更履歴_差分 の最初の行。記録が無ければ null）。 */
function mtgLogStart_(ss) {
  const log = ss.getSheetByName(DIFF_OPTIONS.logSheet);
  if (!log || log.getLastRow() < 2) return null;
  const headers = log.getRange(1, 1, 1, log.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
  const col = headers.indexOf('記録日時');
  return col < 0 ? null : mtgTime_(log.getRange(2, col + 1).getValue());
}

/**
 * 変更時点スナップショットの since 以降の行を、「イベントID#行番号」→ {customer, project, status} にまとめる
 * （その時の行の内容。変更の記録を得意先＋案件名に結び付けるのに使う）。
 */
function mtgReadSnapshots_(ss, since) {
  const read = mtgReadRecent_(ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet), since);
  const i = read.index;
  const map = new Map();
  if (i['イベントID'] == null || i['行番号'] == null) return map;
  const cols = MTG_OPTIONS.columns;
  read.rows.forEach(r => {
    const content = new Map();
    if (i['行の内容'] != null) {
      diffParseSnapshotContent_(r[i['行の内容']]).forEach(([k, v]) => content.set(credNormalize_(k), v));
    }
    const pick = label => String(content.get(credNormalize_(label)) || '').trim();
    map.set(mtgCellText_(r[i['イベントID']]) + '#' + Number(r[i['行番号']]), {
      customer: pick(cols.customer) || (i['得意先'] != null ? mtgCellText_(r[i['得意先']]).trim() : ''),
      project: pick(cols.project),
      status: pick(cols.status)
    });
  });
  return map;
}

/**
 * 追跡シートの案件（得意先か案件名のある行）を読む。
 * {headers, has: {見出しのキー: シートにあるか}, statusOrder: ステータスのプルダウンの選択肢（無ければ空）,
 *  deals: [{row, key, customer, project, status, kind, amount, weighted, orderMonth: {y, m}|null, orderMonthText,
 *  startMonth: {y, m}|null, startMonthText, activity, issue, owner, delivery, updatedAt: 時刻（ms）|null, missing: [見出し]}]}
 */
function mtgReadDeals_(sheet, rule) {
  const headerRow = rule.headerRow;
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, width).getDisplayValues()[0];
  const find = label => headers.findIndex(h => credNormalize_(h) === credNormalize_(label));
  const cols = {};
  const has = {};
  Object.keys(MTG_OPTIONS.columns).forEach(key => {
    cols[key] = find(MTG_OPTIONS.columns[key]);
    has[key] = cols[key] >= 0;
  });
  ['customer', 'project', 'status'].forEach(key => {
    if (cols[key] < 0) {
      throw new Error('「' + sheet.getName() + '」の ' + headerRow + '行目に「' + MTG_OPTIONS.columns[key] +
        '」の見出しがありません（MeetingSummary.gs の MTG_OPTIONS.columns を確かめてください）。');
    }
  });
  const stamp = diffNormalizeStampHeader_(DIFF_OPTIONS.stampHeader);
  const stampCol = headers.findIndex(h => diffNormalizeStampHeader_(h) === stamp);
  const required = MTG_OPTIONS.required.map(label => ({label, col: find(label)})).filter(r => r.col >= 0);

  const out = {headers, has, statusOrder: [], deals: []};
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return out;
  const rule0 = sheet.getRange(headerRow + 1, cols.status + 1).getDataValidation();
  if (rule0 && rule0.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    out.statusOrder = rule0.getCriteriaValues()[0].map(v => String(v).trim()).filter(Boolean);
  }
  const range = sheet.getRange(headerRow + 1, 1, count, width);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  shown.forEach((line, n) => {
    const text = key => (cols[key] >= 0 ? String(line[cols[key]]).trim() : '');
    const value = key => (cols[key] >= 0 ? values[n][cols[key]] : null);
    const customer = text('customer');
    const project = text('project');
    if (!customer && !project) return;
    const status = text('status');
    out.deals.push({
      row: headerRow + 1 + n,
      key: mtgDealKey_(customer, project),
      customer, project, status,
      kind: mtgStatusKind_(status),
      amount: mtgNumber_(value('amount')),
      weighted: mtgNumber_(value('weighted')),
      orderMonth: mtgMonth_(value('orderMonth')),
      orderMonthText: text('orderMonth'),
      startMonth: mtgMonth_(value('startMonth')),
      startMonthText: text('startMonth'),
      activity: text('activity'),
      issue: text('issue'),
      owner: text('owner'),
      delivery: text('delivery'),
      updatedAt: stampCol >= 0 ? mtgTime_(values[n][stampCol]) : null,
      missing: required.filter(r => !String(line[r.col]).trim()).map(r => r.label)
    });
  });
  return out;
}

/* ---------------- まとめる ---------------- */

/**
 * 1つのシートの会議サマリの中身。
 * context = {since, previous, legacy, draft, final, now, meeting, windowStart, trackedSince, log, snapshots, credentials}
 * 戻り値 {name, meeting, previous, legacy, since, first, final, now, windowStart, kpi, fiscal, bridge, stages, quarters,
 *         changes, stalled, focus: 重点案件, alerts, resolved, unmatched, lastUpdated: いちばん新しい最終更新日時（ms）}
 */
function mtgSummarize_(ss, name, context) {
  const o = MTG_OPTIONS;
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('「' + name + '」シートが見つかりません。');
  const table = mtgReadDeals_(sheet, DIFF_RULES[name]);
  const deals = table.deals;
  const byKey = new Map();
  deals.forEach(d => { if (!byKey.has(d.key)) byKey.set(d.key, d); });
  const now = context.now.getTime();
  const since = context.since.getTime();
  const history = mtgHistory_(name, context);
  const labels = mtgLabels_();
  const order = table.statusOrder.map(credNormalize_);
  const rank = status => order.indexOf(credNormalize_(status));

  // 前回からの変化（案件ごとに1つの区分）。行を削除した案件は「削除」
  const changes = [];
  const changed = new Map();
  byKey.forEach((deal, key) => {
    const item = mtgChange_(key, deal, null, history.events.get(key) || [], since, labels, rank);
    if (item) { changes.push(item); changed.set(key, item); }
  });
  history.removed.forEach(line => {
    const r = Object.assign(mtgParseRemoved_(line.text, table.headers), {at: line.at});
    r.key = mtgDealKey_(r.customer, r.project);
    if (r.at <= since || !(r.customer || r.project) || byKey.has(r.key) || changed.has(r.key)) return;
    const item = mtgChange_(r.key, null, r, history.events.get(r.key) || [], since, labels, rank);
    if (item) { changes.push(item); changed.set(r.key, item); }
  });
  const priority = MTG_MOVE_ORDER.concat(['updated']);
  changes.sort((a, b) => (priority.indexOf(a.category) - priority.indexOf(b.category)) ||
    ((b.amount || 0) - (a.amount || 0)) || (b.at - a.at));

  // 期待値の増減（前回の時点で進行中だった案件の期待値 → 今の進行中の期待値）。進行中から外れた分・入った分も内訳にする
  const bridge = mtgBridge_(deals, byKey, changed, changes);

  // 進行中の案件の滞留：ステータスが最後に変わってから（または案件ができてから）の日数。変わった記録が無ければ、
  // 記録を読んだ期間の始まり・最終更新日時の古いほうからの「N日以上」（どちらも無ければ分からない）
  const tracked = context.trackedSince == null ? null : Math.max(context.trackedSince, context.windowStart.getTime());
  const month = mtgMonthIndex_(mtgMonth_(Utilities.formatDate(context.now, Session.getScriptTimeZone(), 'yyyy/MM')));
  const slipsOf = new Map();
  const items = [];
  byKey.forEach((deal, key) => {
    const events = history.events.get(key) || [];
    const slips = events.filter(e => {
      const from = mtgMonth_(e.before);
      const to = mtgMonth_(e.after);
      return e.label === labels.orderMonth && from && to && mtgMonthIndex_(to) > mtgMonthIndex_(from);
    }).length;
    slipsOf.set(key, slips);
    if (deal.kind !== 'active') return;
    const moved = events.filter(e => e.label === labels.status || e.born).map(e => e.at);
    let at = moved.length ? Math.max.apply(null, moved) : null;
    const atLeast = at == null;
    if (at == null) {
      const bounds = [tracked, deal.updatedAt].filter(v => v != null);
      at = bounds.length ? Math.min.apply(null, bounds) : null;
    }
    const days = at == null ? null : Math.floor((now - at) / 86400000);
    const words = o.canvas.words.filter(w => credNormalize_(deal.issue).indexOf(credNormalize_(w)) >= 0);
    const stalled = days != null && days >= o.stallDays;
    const overdue = !!deal.orderMonth && mtgMonthIndex_(deal.orderMonth) < month;
    // 支援の要否を決めたい理由（重点案件の赤枠）
    const reasons = [stalled && '停滞', overdue && '受注月超過', slips >= o.canvas.slips && '延期' + slips + '回',
      words.length && '課題に' + words.join('・')].filter(Boolean);
    items.push(Object.assign({}, deal, {
      days, atLeast, slips, overdue, stalled, reasons, support: reasons.length > 0, star: false,
      candidate: stalled && (days >= o.canvas.days || slips >= o.canvas.slips || words.length > 0),
      decide: words.length ? '上席の同行の要否' : overdue || slips >= o.canvas.slips ? '受注月の見直しか、\nクローズか'
        : stalled && days >= o.canvas.days ? '継続か、クローズか' : stalled ? '次の打ち手の確認' : '受注に向けた\n次のアクション'
    }));
  });
  // 停滞：滞留が stallDays 日以上（想定の大きい順）。★は当てはまる案件の想定の大きい canvas.max 件まで
  const stalled = items.filter(d => d.stalled)
    .sort((a, b) => ((b.amount || 0) - (a.amount || 0)) || ((b.weighted || 0) - (a.weighted || 0)) || (b.days - a.days));
  stalled.filter(d => d.candidate).forEach((d, i) => { d.star = i < o.canvas.max; });
  // 重点案件：★の案件と、期待値の大きい進行中の案件（★を先に入れて mainRows.focus 件まで。並びは期待値の大きい順）
  const byWeighted = (a, b) => ((b.weighted || 0) - (a.weighted || 0)) || ((b.amount || 0) - (a.amount || 0));
  const ranked = items.slice().sort(byWeighted);
  const focus = ranked.filter(d => d.star).concat(ranked.filter(d => !d.star)).slice(0, o.mainRows.focus).sort(byWeighted);

  // 要対応と、試作から解消した件数
  const alerts = mtgAlerts_(table, context, {slipsOf, changes, tracked, rank});
  let resolved = null;
  if (context.draft) {
    const hashes = alerts.map(a => mtgHash_(a.key));
    resolved = {
      at: context.draft.at,
      fixed: context.draft.keys.filter(k => hashes.indexOf(k) < 0).length,
      added: hashes.filter(h => context.draft.keys.indexOf(h) < 0).length
    };
  }

  const active = deals.filter(d => d.kind === 'active');
  const count = key => changes.filter(c => c.category === key).length;
  const owners = new Set(alerts.map(a => a.owner || '（責任者が空）'));
  return {
    name,
    meeting: context.meeting,
    previous: context.previous || null,
    legacy: context.legacy || null,
    since: context.since,
    first: !context.previous && !context.legacy,
    final: context.final,
    now: context.now,
    windowStart: context.windowStart,
    kpi: {
      deals: deals.length,
      active: active.length,
      amount: mtgSum_(active, 'amount'),
      weighted: mtgSum_(active, 'weighted'),
      missingWeighted: active.filter(d => d.weighted == null).length,
      changed: changes.filter(c => c.category !== 'updated').length,
      updated: count('updated'),
      new: count('new'),
      won: count('won'),
      wonAmount: mtgSum_(changes.filter(c => c.category === 'won'), 'amount'),
      lost: count('lost'),
      hold: count('hold'),
      revived: count('revived'),
      forward: count('forward'),
      backward: count('backward'),
      removed: count('removed'),
      slipped: count('slipped'),
      slips: changes.filter(c => c.flags.indexOf('slipped') >= 0).length,
      bad: changes.filter(c => mtgCategory_(c.category).bad).length,
      stalled: stalled.length,
      stalledAmount: mtgSum_(stalled, 'amount'),
      stars: stalled.filter(d => d.star).length,
      alerts: alerts.length,
      owners: owners.size,
      support: focus.filter(d => d.support).length
    },
    fiscal: mtgFiscal_(name, deals, context.now),
    bridge,
    stages: mtgStages_(deals, table.statusOrder),
    quarters: mtgQuarters_(active, context.now),
    changes,
    stalled,
    focus,
    alerts,
    resolved,
    unmatched: history.unmatched,
    lastUpdated: deals.reduce((max, d) => (d.updatedAt != null && (max == null || d.updatedAt > max) ? d.updatedAt : max), null)
  };
}

/** 正規化した見出し（変更の記録の列名と比べる） */
function mtgLabels_() {
  const cols = MTG_OPTIONS.columns;
  const out = {};
  Object.keys(cols).forEach(key => { out[key] = credNormalize_(cols[key]); });
  return out;
}

/**
 * 変更の記録を案件（得意先＋案件名）ごとにまとめる。得意先・案件名を書き換えた・あとから入れた案件は、前の名前の記録もつなげる。
 * {events: Map(キー → [{at, label, column, before, after, born}]（古い順））, removed: [{at, text: 削除した行の「見出し: 値」}],
 *  unmatched: 前回以降で結び付けられなかった記録の数}
 * born は、空の行に得意先か案件名を入れた（案件ができた）記録・行の追加。
 */
function mtgHistory_(name, context) {
  const resolve = diffSheetNameResolver_();
  const labels = mtgLabels_();
  const since = context.since.getTime();
  const empty = v => (v === CHG_OPTIONS.emptyMark ? '' : String(v == null ? '' : v).trim());
  const isKey = label => label === labels.customer || label === labels.project;
  // 同じイベントで得意先・案件名を一緒に変えたとき（貼り付けなど）の、もう一方の変更前
  const keyBefore = new Map();
  context.log.forEach(e => {
    const label = credNormalize_(e.column);
    if (isKey(label)) keyBefore.set(e.eventId + '#' + mtgCellRow_(e.cell) + '#' + label, empty(e.before));
  });
  const alias = new Map();
  const items = [];
  const removed = [];
  let unmatched = 0;
  context.log.forEach(e => {
    if (resolve(e.sheet) !== name || e.at == null || /^(SMP|CRD)-/.test(e.eventId)) return;
    if (e.type === DIFF_ROW_OPTIONS.kindDelete) {
      removed.push({at: e.at, text: e.before});
      return;
    }
    const row = mtgCellRow_(e.cell);
    const snap = context.snapshots.get(e.eventId + '#' + row);
    if (!snap || !(snap.customer || snap.project)) {
      if (e.at > since) unmatched++;
      return;
    }
    const key = mtgDealKey_(snap.customer, snap.project);
    const label = credNormalize_(e.column);
    const item = {key, at: e.at, label, column: e.column, before: empty(e.before), after: empty(e.after),
      born: e.type === DIFF_ROW_OPTIONS.kindInsert};
    if (isKey(label)) {
      const isCustomer = label === labels.customer;
      const other = isCustomer ? labels.project : labels.customer;
      const otherKey = e.eventId + '#' + row + '#' + other;
      const otherBefore = keyBefore.has(otherKey) ? keyBefore.get(otherKey) : (isCustomer ? snap.project : snap.customer);
      if (!item.before && !otherBefore) item.born = true;
      const old = isCustomer ? mtgDealKey_(item.before, otherBefore) : mtgDealKey_(otherBefore, item.before);
      if (!item.born && old !== key) alias.set(old, key);
    }
    items.push(item);
  });
  const root = key => {
    const seen = new Set();
    while (alias.has(key) && !seen.has(key)) {
      seen.add(key);
      key = alias.get(key);
    }
    return key;
  };
  const events = new Map();
  items.forEach(item => {
    const key = root(item.key);
    if (!events.has(key)) events.set(key, []);
    events.get(key).push(item);
  });
  return {events, removed, unmatched};
}

/**
 * 1つの案件の前回からの変化（変化が無ければ null）。deal は今の行、removed は削除した行（どちらか一方）。
 * 前回の時点の値は、前回以降で最初の変更の「変更前」（その項目の変更が無ければ今の値）。
 */
function mtgChange_(key, deal, removed, events, since, labels, rank) {
  const recent = events.filter(e => e.at > since);   // 前回の資料を作った時刻ちょうどの記録は、前回の資料に入っている
  if (!recent.length && !removed) return null;
  const born = recent.some(e => e.born);
  if (born && removed) return null;   // 前回より後に足して、消した行
  const cur = deal || removed;
  const pick = (label, value, parse) => {
    const e = recent.find(x => x.label === label);
    return e ? parse(e.before) : value;
  };
  const base = born ? null : {
    status: pick(labels.status, cur.status, String),
    amount: pick(labels.amount, cur.amount, mtgNumber_),
    weighted: pick(labels.weighted, cur.weighted, mtgNumber_),
    monthText: pick(labels.orderMonth, cur.orderMonthText, String)
  };
  const flags = [];
  if (!base) {
    flags.push('new');
  } else if (removed) {
    flags.push('removed');
  } else {
    const from = mtgStatusKind_(base.status);
    if (credNormalize_(base.status) !== credNormalize_(deal.status)) {
      if (from !== deal.kind) {
        flags.push({won: 'won', lost: 'lost', hold: 'hold', active: 'revived'}[deal.kind]);
      } else if (deal.kind === 'active' && rank(base.status) >= 0 && rank(deal.status) >= 0) {
        flags.push(rank(deal.status) > rank(base.status) ? 'forward' : 'backward');
      } else {
        flags.push('status');
      }
    }
    const months = mtgMonthIndex_(deal.orderMonth) - mtgMonthIndex_(mtgMonth_(base.monthText));
    if (deal.orderMonth && mtgMonth_(base.monthText) && months) flags.push(months > 0 ? 'slipped' : 'pulled');
    const delta = (deal.weighted || 0) - (base.weighted || 0);
    if (Math.abs(delta) > 1e-9) flags.push(delta > 0 ? 'up' : 'down');
    if (!flags.length) flags.push('updated');
  }
  const category = MTG_CATEGORIES.find(c => flags.indexOf(c.key) >= 0).key;
  const columns = [];
  recent.forEach(e => {
    if (e.column && e.column !== '（行全体）' && columns.indexOf(e.column) < 0) columns.push(e.column);
  });
  return {
    key, category, flags, base,
    customer: cur.customer, project: cur.project,
    deal: deal || null,
    removed: !!removed,
    statusFrom: base ? base.status : '',
    statusTo: deal ? deal.status : '',
    weightedFrom: base ? base.weighted : null,
    weightedTo: deal ? deal.weighted : null,
    monthFrom: base ? base.monthText : '',
    monthTo: deal ? deal.orderMonthText : '',
    amount: cur.amount,
    owner: cur.owner || '',
    activity: deal ? deal.activity : '',
    issue: deal ? deal.issue : '',
    columns,
    at: Math.max.apply(null, recent.map(e => e.at).concat([removed ? removed.at : 0]))
  };
}

/**
 * 期待値の増減。前回の時点で進行中だった案件の期待値（start）から、今の進行中の期待値（end）までを内訳にする（start + 内訳の合計 = end）。
 * 進行中のまま：増額・減額。進行中に入った：新規・復帰。進行中から外れた：受注・失注・保留・削除（前回の期待値の分だけ減る）。
 * {start, end, startCount, endCount, steps: [{key, label, value, count}]}
 */
function mtgBridge_(deals, byKey, changed, changes) {
  const steps = {};
  MTG_BRIDGE.forEach(b => { steps[b.key] = {key: b.key, label: b.label, value: 0, count: 0}; });
  const out = {start: 0, end: 0, startCount: 0, endCount: 0};
  const add = (item, was, wStart, now, wEnd, kindNow) => {
    if (was) { out.start += wStart; out.startCount++; }
    if (now) { out.end += wEnd; out.endCount++; }
    let key = null;
    if (!was && now) key = item && item.flags.indexOf('new') >= 0 ? 'new' : 'revived';
    else if (was && !now) key = item && item.removed ? 'removed' : ({won: 'won', lost: 'lost', hold: 'hold'}[kindNow] || 'down');
    else if (was && now && wEnd !== wStart) key = wEnd > wStart ? 'up' : 'down';
    if (!key) return;
    steps[key].value += (now ? wEnd : 0) - (was ? wStart : 0);
    steps[key].count++;
  };
  deals.forEach(d => {
    const item = byKey.get(d.key) === d ? changed.get(d.key) : null;
    const was = item ? !!item.base && mtgStatusKind_(item.base.status) === 'active' : d.kind === 'active';
    const wStart = item ? (item.base && item.base.weighted) || 0 : d.weighted || 0;
    add(item, was, wStart, d.kind === 'active', d.weighted || 0, d.kind);
  });
  changes.filter(c => c.removed).forEach(c => {
    add(c, mtgStatusKind_(c.base.status) === 'active', c.base.weighted || 0, false, 0, null);
  });
  out.steps = MTG_BRIDGE.map(b => steps[b.key]);
  return out;
}

/**
 * 要対応 [{type, key, dealKey, customer, project, text, amount, owner}]（MTG_ALERT_TYPES の順。種類の中は想定の大きい順。予定超過は過ぎた日数の多い順）。
 * info = {slipsOf: キー → 受注月の延期の回数, changes, tracked: 変更の記録がそろっている始まり（ms）, rank}
 */
function mtgAlerts_(table, context, info) {
  const o = MTG_OPTIONS;
  const tz = Session.getScriptTimeZone();
  const now = context.now.getTime();
  const today = Utilities.formatDate(context.now, tz, 'yyyy/MM/dd');
  const deals = table.deals;
  const active = deals.filter(d => d.kind === 'active');
  const out = [];
  const push = (type, d, text) => out.push({type, key: type + '|' + d.key, dealKey: d.key, customer: d.customer, project: d.project,
    text, amount: d.amount, owner: d.owner || ''});
  const month = mtgMonthIndex_(mtgMonth_(today));
  const label = m => m.y + '/' + m.m;

  // 予定超過：このシートに進行中の案件がある得意先の、予定日を過ぎても実施日が空のクレデンシャル・オファリング
  const customers = new Map();
  active.forEach(d => {
    const key = credNormalize_(d.customer);
    if (d.customer && !customers.has(key)) customers.set(key, d.customer);
  });
  const overdue = [];
  customers.forEach((customer, key) => {
    (context.credentials[key] || []).forEach(c => {
      const plan = mtgDateText_(c.values.plan);
      if (!plan || mtgDateText_(c.values.date) || plan >= today) return;
      const days = Math.round((mtgDay_(today) - mtgDay_(plan)) / 86400000);
      if (days > o.overdueDays) return;
      overdue.push({type: '予定超過', key: '予定超過|' + c.id, dealKey: '', customer, project: '', owner: c.values.owner || '', amount: null, days,
        text: (c.values.kind || '予定') + '「' + c.values.name + '」の予定日（' + plan.slice(5) + '）から ' + days + '日、実施日が空'});
    });
  });
  overdue.sort((a, b) => b.days - a.days);
  active.forEach(d => {
    const slips = info.slipsOf.get(d.key) || 0;
    if (d.orderMonth && mtgMonthIndex_(d.orderMonth) < month) {
      push('受注月超過', d, '受注月 ' + label(d.orderMonth) + ' を過ぎても「' + (d.status || '未設定') + '」' + (slips ? '（延期 ' + slips + '回）' : ''));
    }
    if (d.missing.length) push('未入力', d, '空：' + d.missing.map(h => h.replace(/\s+/g, '')).join('・'));
    if (d.updatedAt != null) {
      const days = Math.floor((now - d.updatedAt) / 86400000);
      if (days >= o.noUpdateDays) push('更新なし', d, '最終更新から ' + days + '日');
    } else if (info.tracked != null && now - info.tracked >= o.noUpdateDays * 86400000) {
      push('更新なし', d, '変更の記録を始めてから更新なし（' + Math.floor((now - info.tracked) / 86400000) + '日以上）');
    }
    if (d.amount != null && d.weighted != null && d.weighted > d.amount) {
      push('金額の矛盾', d, '期待値 ' + mtgAmount_(d.weighted) + ' が想定 ' + mtgAmount_(d.amount) + ' より大きい');
    }
  });
  deals.forEach(d => {
    if ((d.kind === 'active' || d.kind === 'won') && d.orderMonth && d.startMonth &&
        mtgMonthIndex_(d.startMonth) < mtgMonthIndex_(d.orderMonth)) {
      push('日付の矛盾', d, '売上開始月 ' + label(d.startMonth) + ' が受注月 ' + label(d.orderMonth) + ' より前');
    }
    if (table.statusOrder.length && d.status && info.rank(d.status) < 0) {
      push('未定義ステータス', d, '「' + d.status + '」はプルダウンの選択肢にない（' + {active: '進行中', won: '受注', lost: '失注', hold: '保留'}[d.kind] + 'として数える）');
    }
  });
  if (table.has.delivery) {
    info.changes.filter(c => c.category === 'won' && c.deal && !c.deal.delivery)
      .forEach(c => push('体制未登録', c.deal, '受注したが「' + MTG_OPTIONS.columns.delivery + '」が空'));
  }
  const byAmount = (a, b) => (b.amount || 0) - (a.amount || 0);
  return overdue.concat(MTG_ALERT_TYPES.slice(1).reduce((list, type) => list.concat(out.filter(a => a.type === type).sort(byAmount)), []));
}

/** 要対応の種類の説明（要対応のページの右に出す） */
function mtgAlertNote_(type) {
  const o = MTG_OPTIONS;
  return {
    '予定超過': 'クレデンシャル・オファリングの予定日を過ぎても実施日が空（過ぎて' + o.overdueDays + '日以内）',
    '受注月超過': '受注月を過ぎても進行中のまま',
    '未入力': '進行中の案件で、' + o.required.map(h => h.replace(/\s+/g, '')).join('・') + ' のどれかが空',
    '更新なし': '進行中で、最終更新日時が' + o.noUpdateDays + '日以上前',
    '金額の矛盾': '期待値調整済が想定売上規模より大きい',
    '日付の矛盾': '売上開始月が受注月より前',
    '未定義ステータス': '案件ステータスが、プルダウンの選択肢にない値',
    '体制未登録': '前回から受注になったが、' + o.columns.delivery + 'が空'
  }[type] || '';
}

/** 今期（MTG_OPTIONS.fiscalStartMonth から12か月）の受注と着地見込み。{label, wonCount, wonAmount, planned, landing, target} */
function mtgFiscal_(name, deals, now) {
  const start = mtgFiscalStart_(now);
  const inYear = m => m && mtgMonthIndex_(m) >= start && mtgMonthIndex_(m) < start + 12;
  const won = deals.filter(d => d.kind === 'won' && inYear(d.orderMonth));
  const planned = mtgSum_(deals.filter(d => d.kind === 'active' && inYear(d.orderMonth)), 'weighted');
  const wonAmount = mtgSum_(won, 'amount');
  const from = mtgMonthOf_(start);
  const to = mtgMonthOf_(start + 11);
  const target = Number(MTG_OPTIONS.targets[name]) || null;
  return {
    label: from.y + '/' + from.m + '〜' + to.y + '/' + to.m,
    wonCount: won.length,
    wonAmount,
    planned,
    landing: wonAmount + planned,
    target
  };
}

/** 今期の始まりの月の番号（年 × 12 + 月） */
function mtgFiscalStart_(now) {
  const m = mtgMonth_(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM'));
  const fs = MTG_OPTIONS.fiscalStartMonth;
  return (m.m >= fs ? m.y : m.y - 1) * 12 + fs;
}

/** 進行中の案件の受注月ごと（今期の四半期・来期以降・前期以前・受注月なし）の期待値 [{label, value, count}] */
function mtgQuarters_(active, now) {
  const start = mtgFiscalStart_(now);
  const buckets = [{label: '前期以前', value: 0, count: 0, optional: true}];
  for (let q = 0; q < 4; q++) {
    const a = mtgMonthOf_(start + q * 3);
    const b = mtgMonthOf_(start + q * 3 + 2);
    buckets.push({label: (q + 1) + 'Q（' + a.m + '〜' + b.m + '月）', value: 0, count: 0});
  }
  buckets.push({label: '来期以降', value: 0, count: 0, optional: true}, {label: '受注月なし', value: 0, count: 0, optional: true});
  active.forEach(d => {
    const i = d.orderMonth ? mtgMonthIndex_(d.orderMonth) : null;
    const bucket = i == null ? buckets[6] : i < start ? buckets[0] : i >= start + 12 ? buckets[5] : buckets[1 + Math.floor((i - start) / 3)];
    bucket.value += d.weighted || 0;
    bucket.count++;
  });
  return buckets.filter(b => !b.optional || b.count).map(b => ({label: b.label, value: b.value, count: b.count}));
}

/** ステータスごとの件数・金額 [{status, kind, known, count, amount, weighted}]（進行中 → 保留 → 受注 → 失注。同じ分類の中はプルダウンの順） */
function mtgStages_(deals, statusOrder) {
  const map = new Map();
  const order = statusOrder.slice();
  deals.forEach(d => {
    const label = d.status || '（未設定）';
    if (order.indexOf(label) < 0) order.push(label);
    const item = map.get(label) || {status: label, kind: d.kind, known: !statusOrder.length || statusOrder.indexOf(label) >= 0 || !d.status,
      count: 0, amount: 0, weighted: 0};
    item.count++;
    item.amount += d.amount || 0;
    item.weighted += d.weighted || 0;
    map.set(label, item);
  });
  const rankOf = {active: 0, hold: 1, won: 2, lost: 3};
  return order.filter(label => map.has(label)).map(label => map.get(label))
    .sort((a, b) => rankOf[a.kind] - rankOf[b.kind]);   // 同じ分類の中はプルダウンの順（安定ソート）
}

/**
 * 削除した行の「変更前」（「見出し: 値」の行）から {customer, project, status, amount, weighted, orderMonth, orderMonthText, owner}。
 * 見出し・値に改行が入ることがあるので、シートの見出し（headers）で始まる行を区切りにする（それ以外の行は前の値の続き）。
 */
function mtgParseRemoved_(text, headers) {
  const s = String(text || '');
  const heads = (headers || []).map(h => String(h).trim()).filter(Boolean).sort((a, b) => b.length - a.length);
  const values = new Map();
  let current = null;
  let pos = 0;
  while (pos < s.length) {
    const head = heads.find(h => s.startsWith(h + ': ', pos));
    if (head) {
      current = credNormalize_(head);
      values.set(current, '');
      pos += head.length + 2;
    } else if (current) {
      values.set(current, values.get(current) + '\n');
    }
    const end = s.indexOf('\n', pos) < 0 ? s.length : s.indexOf('\n', pos);
    if (current) values.set(current, values.get(current) + s.slice(pos, end));
    pos = end + 1;
  }
  const cols = MTG_OPTIONS.columns;
  const get = label => String(values.get(credNormalize_(label)) || '').trim();
  return {customer: get(cols.customer), project: get(cols.project), status: get(cols.status),
    amount: mtgNumber_(get(cols.amount)), weighted: mtgNumber_(get(cols.weighted)),
    orderMonth: mtgMonth_(get(cols.orderMonth)), orderMonthText: get(cols.orderMonth), owner: get(cols.owner)};
}

/* ---------------- 小さな補助 ---------------- */

/** 得意先＋案件名のキー（全角・半角と空白の違いは無視）。 */
function mtgDealKey_(customer, project) {
  return credNormalize_(customer) + '\u0001' + credNormalize_(project);
}

/** ステータスの分類：'won'（受注）・'lost'（失注・クローズなど）・'hold'（保留）・'active'（進行中。空も） */
function mtgStatusKind_(status) {
  const text = credNormalize_(status);
  if (!text) return 'active';
  const o = MTG_OPTIONS;
  const fixed = Object.keys(o.statusKinds || {}).find(k => credNormalize_(k) === text);
  if (fixed) return o.statusKinds[fixed];
  const has = words => words.some(w => text.indexOf(credNormalize_(w)) >= 0);
  if (has(o.lostWords)) return 'lost';
  if (has(o.holdWords)) return 'hold';
  if (has(o.wonWords) && !has(o.notWonWords)) return 'won';
  return 'active';
}

/** 区分の定義（MTG_CATEGORIES の1つ） */
function mtgCategory_(key) {
  return MTG_CATEGORIES.find(c => c.key === key) || {key, label: key};
}

/** セルの値の時刻（ms）。日時の値か「yyyy/MM/dd HH:mm」の文字。読めなければ null。 */
function mtgTime_(value) {
  if (value && typeof value.getTime === 'function') return value.getTime();
  const m = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(String(value || '').trim());
  if (!m) return null;
  const text = m[1] + '/' + m[2].padStart(2, '0') + '/' + m[3].padStart(2, '0') + ' ' +
    (m[4] || '0').padStart(2, '0') + ':' + (m[5] || '00') + ':' + (m[6] || '00');
  return Utilities.parseDate(text, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss').getTime();
}

/** 日付の文字（「2026/10/01」の形。読めなければ空）。 */
function mtgDateText_(value) {
  const m = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/.exec(String(value || '').trim());
  return m ? m[1] + '/' + m[2].padStart(2, '0') + '/' + m[3].padStart(2, '0') : '';
}

/** 「2026/10/01」の日の ms（日数の差を出す用。時差の影響を受けないよう UTC で数える）。 */
function mtgDay_(text) {
  const p = text.split('/').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2]);
}

/** 「2026/10/23」→「10/23」 */
function mtgShortDate_(text) {
  const p = String(text || '').split('/');
  return p.length === 3 ? Number(p[1]) + '/' + Number(p[2]) : String(text || '');
}

/** 数（「1,234.5」「12百万」も読む）。空・読めなければ null。 */
function mtgNumber_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : null;
  const text = String(value == null ? '' : value).normalize('NFKC').replace(/[,\s]|百万|円/g, '');
  return /^-?\d+(\.\d+)?$/.test(text) ? Number(text) : null;
}

/** 受注月 {y, m}（日付の値・「2026/10」「2026年10月」「2026-10-01」）。読めなければ null。 */
function mtgMonth_(value) {
  if (value && typeof value.getTime === 'function') {
    const p = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM').split('-').map(Number);
    return {y: p[0], m: p[1]};
  }
  const m = /(\d{4})\s*[\/\-年.]\s*(\d{1,2})/.exec(String(value || '').normalize('NFKC'));
  return m && Number(m[2]) >= 1 && Number(m[2]) <= 12 ? {y: Number(m[1]), m: Number(m[2])} : null;
}

/** 月の番号（年 × 12 + 月。比べる用）。月が無ければ null */
function mtgMonthIndex_(month) {
  return month ? month.y * 12 + month.m : null;
}

/** 月の番号 → {y, m} */
function mtgMonthOf_(index) {
  return {y: Math.floor((index - 1) / 12), m: (index - 1) % 12 + 1};
}

/** 「F12」「12行」「12行（削除前の行番号）」の行番号（読めなければ 0）。 */
function mtgCellRow_(cell) {
  const m = /^[A-Z]+(\d+)$/.exec(String(cell)) || /^(\d+)行/.exec(String(cell));
  return m ? Number(m[1]) : 0;
}

/** セルの値を文字に（日時は記録の形）。 */
function mtgCellText_(value) {
  if (value && typeof value.getTime === 'function') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), DIFF_OPTIONS.dateFormat);
  }
  return String(value == null ? '' : value);
}

function mtgSum_(list, key) {
  return list.reduce((sum, d) => sum + (d[key] || 0), 0);
}

/** 金額（百万）の表示：「1,234.5」（整数なら小数なし）。null は「—」。 */
function mtgAmount_(value) {
  if (value == null) return '—';
  const rounded = Math.round(value * 10) / 10;
  const parts = String(Math.abs(rounded)).split('.');
  return (rounded < 0 ? '-' : '') + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] ? '.' + parts[1] : '');
}

/** 増減の表示：「+12.5」「▲3」「±0」 */
function mtgSigned_(value) {
  const rounded = Math.round((value || 0) * 10) / 10;
  if (!rounded) return '±0';
  return (rounded > 0 ? '+' : '▲') + mtgAmount_(Math.abs(rounded));
}

/** 得意先／案件名 */
function mtgDealLabel_(item) {
  return [item.customer, item.project].filter(Boolean).join('／');
}

/* ---------------- 資料（スライド）にする ---------------- */

/**
 * 1つのシートの会議サマリをスライドの並びにする（Slides API に送る前の形。テストと見た目の確認にも使う）。
 * 本編（サマリー → パイプラインの変化 → 前回からの進捗 → 重点案件の状況・課題）と付録 A〜D。見出しは MTG_TITLES（どのシートでも同じ）。
 * 戻り値 [{title, elements}]（elements は mtgSlideRequests_ の部品。タグ・タイトル・リード文・結論の帯・出典・ページ番号も含む）
 */
function mtgSlides_(summary) {
  const slides = [mtgSummarySlide_(summary), mtgBridgeSlide_(summary), mtgMovesSlide_(summary), mtgFocusSlide_(summary)]
    .concat(mtgAppendixChanges_(summary), mtgAppendixStalled_(summary), mtgAppendixOwners_(summary), mtgAppendixAlerts_(summary),
      mtgAppendixStatus_(summary));
  return mtgNumberSlides_(mtgSheetFrame_(summary), slides);
}

/** スライドに、どのスライドにも付ける部品（mtgFrame_）とページ番号を付ける */
function mtgNumberSlides_(frame, slides) {
  return slides.map((slide, i) => ({title: slide.title, elements: mtgFrame_(frame, slide, i + 1).concat(slide.elements)}));
}

/** シートの資料の左上のタグと出典 {tag, source} */
function mtgSheetFrame_(s) {
  return {
    tag: MTG_OPTIONS.title + '｜' + s.name + '（' + mtgShortDate_(s.meeting) + ' 会議）' + (s.final ? '' : '　試作'),
    source: '出典：「' + s.name + '」シート・' + DIFF_OPTIONS.logSheet + '（' +
      Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' 時点）　期間：' + mtgPeriodText_(s)
  };
}

/** どのスライドにも付ける部品：左上のタグ・タイトル・リード文・結論の帯（本編）・出典・ページ番号。frame = {tag, source} */
function mtgFrame_(frame, slide, page) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const title = mtgTitleFit_(slide.title);
  const el = [
    mtgBox_(0, 0, Math.min(440, mtgTextWidth_(frame.tag, 10) / 0.9 + S.padX + 12), 22, frame.tag,
      {size: 10, bold: true, color: c.white, fill: c.purple, shape: 'RECTANGLE'}),
    mtgBox_(S.x, 26, S.w, 36, title.text, {size: title.size, bold: true, color: c.purple})
  ];
  if (slide.lead) {
    el.push(mtgBox_(S.x, 62, S.w, 34, mtgWrap_(slide.lead, S.w, S.leadSize, 2).join('\n'),
      {size: S.leadSize, bold: true, color: c.ink, valign: 'TOP'}));
  }
  if (slide.band) {
    el.push(mtgBox_(S.x, 354, S.w, 24, mtgWrap_(slide.band, S.w, 10.5, 1)[0],
      {size: 10.5, bold: true, color: c.ink, fill: c.highlight, shape: 'RECTANGLE', align: 'CENTER'}));
  }
  el.push(mtgBox_(S.x, 383, S.w - 60, 16, mtgWrap_(frame.source, S.w - 60, 6.5, 1)[0], {size: 6.5, color: c.faint}));
  el.push(mtgBox_(S.x + S.w - 60, 381, 60, 18, String(page), {size: 9, color: c.muted, align: 'END'}));
  return el;
}

/** 「前回から」「直近14日で」（リード文の書き出し） */
function mtgPeriodShort_(s) {
  return s.first ? '直近' + MTG_OPTIONS.firstDays + '日で' : '前回から';
}

/** 1 サマリー：今期の帯・4つの数字・本日の論点。リード文は期待値の前回比・動いた件数・最優先の論点 */
function mtgSummarySlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const b = s.bridge;
  const f = s.fiscal;
  const delta = b.end - b.start;
  const issues = mtgIssues_(s);
  const el = [];

  const fy = '今期（' + f.label + '）　受注 ' + f.wonCount + '件・' + mtgAmount_(f.wonAmount) + '百万円' +
    (f.target ? '（目標 ' + mtgAmount_(f.target) + '・達成率 ' + Math.round(f.wonAmount / f.target * 100) + '%）' : '') +
    '　｜　今期に受注予定の期待値 ' + mtgAmount_(f.planned) + '百万円　｜　着地見込み ' + mtgAmount_(f.landing) + '百万円';
  el.push(mtgBox_(S.x, S.top, S.w, 20, mtgWrap_(fy, S.w, 9, 1)[0],
    {size: 9, bold: true, color: c.white, fill: c.purple, shape: 'RECTANGLE', align: 'CENTER'}));

  const tiles = [
    {label: '進行中の期待値（百万円）', value: mtgAmount_(b.end),
      sub: (s.first ? '直近' + o.firstDays + '日で ' : '前回比 ') + mtgSigned_(delta) + '・' + k.active + '件（想定 ' + mtgAmount_(k.amount) + '）'},
    {label: '受注／失注', value: k.won + '件／' + k.lost + '件', sub: '受注の想定 ' + mtgAmount_(k.wonAmount) + '百万円・削除 ' + k.removed + '件'},
    {label: '前進／後退', value: k.forward + '件／' + k.backward + '件', sub: '新規 ' + k.new + '件・受注月の延期 ' + k.slips + '件'},
    {label: '停滞（' + o.stallDays + '日以上）', value: k.stalled + '件',
      sub: '想定 ' + mtgAmount_(k.stalledAmount) + '百万円・★ ' + k.stars + '件'}
  ];
  const tw = (S.w - 3 * 8) / 4;
  tiles.forEach((t, i) => {
    const x = S.x + i * (tw + 8);
    el.push({type: 'box', x, y: 124, w: tw, h: 54, fill: c.purpleSoft, shape: 'RECTANGLE'});
    el.push(mtgBox_(x, 125, tw, 15, mtgWrap_(t.label, tw, 7.5, 1)[0], {size: 7.5, bold: true, color: c.muted}));
    el.push(mtgBox_(x, 139, tw, 24, mtgWrap_(t.value, tw, 17, 1)[0], {size: 17, bold: true, color: c.purple}));
    el.push(mtgBox_(x, 162, tw, 15, mtgWrap_(t.sub, tw, 7, 1)[0], {size: 7, color: c.muted}));
  });

  const columns = [128, 32, 282, 32, 198];
  const rows = [mtgHeadRow_(columns, ['本日の論点', '', '状況', '', '決めてほしいこと'])];
  issues.forEach(issue => rows.push(mtgRow_(columns, [
    {text: issue.title + (issue.count ? '\n' + issue.count + '件' : ''), lines: 3, size: 9, bold: true, color: c.white, fill: c.purple},
    mtgArrowCell_(),
    {text: issue.status, lines: 3, fill: c.gray},
    mtgArrowCell_(),
    {text: issue.ask, lines: 3, bold: true, color: c.ink, fill: c.gray}
  ], 46)));
  el.push({type: 'table', x: S.x, y: 186, columns, rows,
    borders: issues[0] && issues[0].focus ? [{row: 1, col: 0, rows: 1, cols: columns.length, color: c.red}] : []});

  return {
    title: MTG_TITLES.summary,
    lead: '期待値は ' + mtgAmount_(b.end) + '百万円（' + (s.first ? '直近' + o.firstDays + '日で ' : '前回比 ') + mtgSigned_(delta) + '）。' +
      mtgPeriodShort_(s) + '受注 ' + k.won + '件・失注 ' + k.lost + '件・前進 ' + k.forward + '件・後退 ' + k.backward + '件・新規 ' + k.new + '件。\n' +
      issues[0].lead + '。',
    band: '本日決めたいこと：' + issues[0].ask.split('\n')[0],
    elements: el
  };
}

/**
 * 本日の論点（上から3つまで）[{title, count, status, ask, lead: サマリーのリード文の2行目, focus}]。
 * 停滞している大型案件 → 受注月の遅れ → 失注・後退・削除 → 受注 → 要対応 の順に、中身のあるものを出す。
 */
function mtgIssues_(s) {
  const o = MTG_OPTIONS;
  const k = s.kpi;
  const names = (list, text) => list.slice(0, 3).map(text).join('、') + (list.length > 3 ? ' ほか' : '');
  const out = [];
  if (s.stalled.length) {
    out.push({title: '停滞している\n大型案件', count: s.stalled.length, focus: true,
      lead: '最優先の論点は、ステータスが' + o.stallDays + '日以上動いていない案件 ' + s.stalled.length + '件（想定 ' +
        mtgAmount_(k.stalledAmount) + '百万円）の打開',
      status: '想定 ' + mtgAmount_(k.stalledAmount) + '百万円が、ステータス' + o.stallDays + '日以上そのまま\n' +
        names(s.stalled, d => mtgDealLabel_(d) + '（' + mtgAmount_(d.amount) + '・' + d.days + '日' + (d.atLeast ? '以上' : '') + '）'),
      ask: '上席の同行と、継続かクローズかの判断' + (k.stars ? '\n（★' + k.stars + '件はオポチュニティキャンバスで深掘り）' : '')});
  }
  const slipped = s.changes.filter(c => c.flags.indexOf('slipped') >= 0);
  const overdue = s.alerts.filter(a => a.type === '受注月超過' && !slipped.some(c => c.key === a.dealKey));
  if (slipped.length || overdue.length) {
    const count = slipped.length + overdue.length;
    out.push({title: '受注月の遅れ', count, lead: '最優先の論点は、受注月の遅れ ' + count + '件の見通しと、今期の着地への影響',
      status: '受注月を延ばした ' + slipped.length + '件・受注月を過ぎても進行中 ' + s.alerts.filter(a => a.type === '受注月超過').length + '件\n' +
        names(slipped.map(c => ({label: mtgDealLabel_(c), text: mtgMonthLabel_(c.monthFrom) + '→' + mtgMonthLabel_(c.monthTo)}))
          .concat(overdue.map(a => ({label: mtgDealLabel_(a), text: '超過'}))), x => x.label + '（' + x.text + '）'),
      ask: '受注月の見通しと、今期の着地への影響'});
  }
  const losses = s.changes.filter(c => ['lost', 'backward', 'removed'].indexOf(c.category) >= 0);
  if (losses.length) {
    out.push({title: '失注・後退', count: losses.length, lead: '最優先の論点は、失注・後退 ' + losses.length + '件の要因と再アプローチの要否',
      status: '失注 ' + k.lost + '件・後退 ' + k.backward + '件・削除 ' + k.removed + '件\n' +
        names(losses, c => mtgDealLabel_(c) + '（' + mtgCategory_(c.category).label + '）'),
      ask: '要因と、再アプローチの要否'});
  }
  const won = s.changes.filter(c => c.category === 'won');
  if (won.length) {
    out.push({title: '受注', count: won.length,
      lead: '最優先の論点は、受注 ' + won.length + '件（想定 ' + mtgAmount_(k.wonAmount) + '百万円）の立ち上げと横展開',
      status: '想定 ' + mtgAmount_(k.wonAmount) + '百万円\n' + names(won, c => mtgDealLabel_(c) + '（' + mtgAmount_(c.amount) + '）'),
      ask: '立ち上げの体制と、ほかの得意先への横展開'});
  }
  if (s.alerts.length) {
    const types = MTG_ALERT_TYPES.map(t => [t, s.alerts.filter(a => a.type === t).length]).filter(p => p[1]);
    out.push({title: '要対応', count: s.alerts.length, lead: '最優先の論点は、要対応 ' + s.alerts.length + '件（責任者 ' + k.owners + '人）の解消',
      status: '責任者 ' + k.owners + '人\n' + types.map(p => p[0] + ' ' + p[1]).join('・'),
      ask: '各責任者への修正の依頼（次の会議まで）'});
  }
  if (!out.length) {
    out.push({title: '大きな動きなし', count: 0, lead: mtgPeriodShort_(s) + '大きな動きは無く、各案件の次のアクションを確認したい',
      status: mtgPeriodShort_(s) + '大きな動きはありません', ask: '各案件の次のアクションの確認'});
  }
  return out.slice(0, 3);
}

/** 2 パイプラインの変化：期待値の増減（滝の図）・ステータス別・受注月（四半期）別の期待値 */
function mtgBridgeSlide_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const b = s.bridge;
  const delta = b.end - b.start;
  const moves = b.steps.filter(st => st.count);
  const biggest = moves.slice().sort((x, y) => Math.abs(y.value) - Math.abs(x.value))[0];
  const worst = moves.filter(st => st.value < 0).sort((x, y) => x.value - y.value)[0];
  const plus = mtgSum_(b.steps.filter(st => st.value > 0), 'value');
  const minus = mtgSum_(b.steps.filter(st => st.value < 0), 'value');
  const peak = s.quarters.filter(q => /Q/.test(q.label) && q.value > 0).sort((x, y) => y.value - x.value)[0];
  const el = mtgWaterfall_(s, S.x, S.top, 360, 244);
  const stages = s.stages.filter(st => st.kind === 'active');
  const shown = stages.slice(0, 8);
  if (stages.length > 8) {
    const rest = stages.slice(7);
    shown.splice(7, 1, {status: 'ほか ' + rest.length + '種類', count: mtgSum_(rest, 'count'), weighted: mtgSum_(rest, 'weighted')});
  }
  const right = S.x + 372;
  el.push.apply(el, mtgBars_('ステータス別の期待値（進行中）', shown.map(st => ({label: st.status, value: st.weighted, count: st.count})),
    right, S.top, S.w - 372, 13));
  el.push.apply(el, mtgBars_('受注月別の期待値（進行中・今期の四半期）', s.quarters, right, S.top + 26 + Math.max(shown.length, 1) * 13, S.w - 372, 13));
  return {
    title: MTG_TITLES.pipeline,
    lead: '期待値は ' + mtgAmount_(b.start) + ' → ' + mtgAmount_(b.end) + '百万円（' + mtgSigned_(delta) + '）。増えた分 ' + mtgSigned_(plus) +
      '・減った分 ' + mtgSigned_(minus) + (biggest && biggest.value ? '、最大の要因は' + biggest.label + '（' + biggest.count + '件・' + mtgSigned_(biggest.value) + '）' : '') +
      '。\n今期に受注予定の期待値は ' + mtgAmount_(s.fiscal.planned) + '百万円' + (peak ? '（' + peak.label.replace(/（.*/, '') + 'が最も多い ' + mtgAmount_(peak.value) + '）' : '') +
      (s.kpi.missingWeighted ? '。期待値が空の ' + s.kpi.missingWeighted + '件は0として計算' : '') + '。',
    band: worst && -worst.value >= plus ? worst.label + 'の ' + worst.count + '件（' + mtgSigned_(worst.value) + '）の要因を確認したい'
      : plus > 0 ? '増えた分（新規・増額・復帰）の確度と受注月を確認したい' : '前回から期待値の増減はありません',
    elements: el
  };
}

/** 期待値の増減の滝の図（前回 → 内訳 → 今回）。左上 (x, y)・幅 w・高さ h */
function mtgWaterfall_(s, x, y, w, h) {
  const c = MTG_OPTIONS.colors;
  const b = s.bridge;
  const el = [mtgBox_(x, y, w, 16, '期待値の増減（百万円）', {size: 9, bold: true, color: c.purple})];
  const bars = [{label: s.first ? MTG_OPTIONS.firstDays + '日前' : '前回', value: b.start, count: b.startCount, total: true}]
    .concat(b.steps.filter(st => st.count))
    .concat([{label: '今回', value: b.end, count: b.endCount, total: true}]);
  const top = y + 36;
  const axis = y + h - 30;
  let level = 0;
  let max = 0;
  const spans = bars.map(bar => {
    const from = bar.total ? 0 : level;
    level = bar.total ? bar.value : level + bar.value;
    max = Math.max(max, from, level);
    return {lo: Math.min(from, level), hi: Math.max(from, level)};
  });
  const scale = max > 0 ? (axis - top) / max : 0;
  const slot = w / bars.length;
  const bw = Math.min(34, slot * 0.6);
  el.push({type: 'box', x, y: axis, w, h: 0.75, fill: c.grayHead, shape: 'RECTANGLE'});
  bars.forEach((bar, i) => {
    const span = spans[i];
    const bx = x + slot * i + (slot - bw) / 2;
    const bh = Math.max((span.hi - span.lo) * scale, 1);
    const by = axis - span.lo * scale - bh;
    el.push({type: 'box', x: bx, y: by, w: bw, h: bh, fill: bar.total ? c.purple : bar.value >= 0 ? c.up : c.down, shape: 'RECTANGLE'});
    const text = bar.total ? mtgAmount_(bar.value) : mtgSigned_(bar.value);
    el.push(mtgBox_(x + slot * i - 8, by - 15, slot + 16, 14, text,
      {size: 7.5, bold: true, color: bar.total ? c.purple : c.ink, align: 'CENTER'}));
    el.push(mtgBox_(x + slot * i - 8, axis + 2, slot + 16, 26, bar.label + '\n' + bar.count + '件',
      {size: 7, color: c.muted, align: 'CENTER', valign: 'TOP'}));
  });
  if (!bars.some(bar => bar.value)) {
    el.push(mtgBox_(x, top + 40, w, 20, '進行中の案件の期待値はありません', {size: 8, color: c.muted, align: 'CENTER'}));
  }
  return el;
}

/** 横棒の一覧（見出し・ラベル・棒・「n件・金額」）。items = [{label, value, count}]。左上 (x, y)・幅 w・1行の高さ rowH */
function mtgBars_(title, items, x, y, w, rowH) {
  const c = MTG_OPTIONS.colors;
  const labelW = 92;
  const valueW = 70;
  const barMax = w - labelW - valueW - 4;
  const max = Math.max.apply(null, items.map(it => it.value).concat([0]));
  const el = [mtgBox_(x, y, w, 16, title, {size: 8.5, bold: true, color: c.purple})];
  if (!items.length) el.push(mtgBox_(x, y + 18, w, rowH, '（なし）', {size: 7, color: c.muted}));
  items.forEach((it, n) => {
    const ry = y + 18 + n * rowH;
    const bw = max > 0 && it.value > 0 ? Math.max(barMax * it.value / max, 1) : 0;
    el.push(mtgBox_(x, ry, labelW, rowH, mtgWrap_(it.label, labelW, 7, 1)[0], {size: 7, color: c.ink, align: 'END'}));
    if (bw) el.push({type: 'box', x: x + labelW + 2, y: ry + 3, w: bw, h: rowH - 6, fill: c.purple, shape: 'RECTANGLE'});
    el.push(mtgBox_(x + labelW + 2 + bw, ry, valueW, rowH, it.count + '件・' + mtgAmount_(it.value), {size: 7, color: c.muted}));
  });
  return el;
}

/** 3 前回からの進捗：ステータス・期待値・受注月が動いた案件（大きな動きの順に上位。悪い動きは赤枠） */
function mtgMovesSlide_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const moves = s.changes.filter(ch => ch.category !== 'updated');
  const top = moves.slice(0, MTG_OPTIONS.mainRows.changes);
  const columns = [146, 44, 130, 66, 80, 154, 52];
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '区分', 'ステータス（前▶今）', '期待値\n（前▶今）', '受注月\n（前▶今）', '活動状況（今）', '責任者'])];
  const borders = [];
  top.forEach(ch => {
    const bad = mtgCategory_(ch.category).bad;
    if (bad) borders.push({row: rows.length, col: 0, rows: 1, cols: columns.length, color: c.red});
    rows.push(mtgRow_(columns, [
      mtgDealCell_(ch, columns[0], false),
      {text: mtgCategory_(ch.category).label, bold: true, color: bad ? c.red : c.purple, fill: c.gray, align: 'CENTER'},
      {text: mtgStatusMove_(ch), lines: 2, fill: c.gray},
      {text: mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: mtgMonthMove_(ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: ch.activity || '—', lines: 2, fill: c.gray},
      {text: ch.owner || '—', lines: 2, fill: c.gray, align: 'CENTER'}
    ]));
  });
  const el = [];
  if (top.length) {
    el.push({type: 'table', x: S.x, y: S.top, columns, rows, borders});
  } else {
    el.push(mtgBox_(S.x, S.top + 20, S.w, 30, '前回から、ステータス・期待値・受注月が動いた案件はありません', {size: 10, color: c.muted, align: 'CENTER'}));
  }
  const lead = moves.length
    ? mtgPeriodShort_(s) + '動いた案件は ' + moves.length + '件（受注 ' + k.won + '・失注 ' + k.lost + '・前進 ' + k.forward + '・後退 ' + k.backward +
      '・延期 ' + k.slipped + '・新規 ' + k.new + '）。' + (k.bad ? '悪い動き（失注・後退・延期）の ' + k.bad + '件は赤枠。' : '悪い動きはなし。') +
      '\n' + (moves.length > top.length ? '大きな動きの順に上位' + top.length + '件。' : '') + '全件と、ほかの項目だけ更新した ' + k.updated + '件は付録A。'
    : mtgPeriodShort_(s) + 'ステータス・期待値・受注月が動いた案件はありません。\nほかの項目だけ更新した案件は ' + k.updated + '件（付録A）。';
  return {
    title: MTG_TITLES.progress,
    lead,
    band: k.bad ? '失注・後退・延期の ' + k.bad + '件は、要因と次の打ち手を確認したい'
      : moves.length ? '前進した案件の次のアクションを確認したい' : '各案件の次のアクションを確認したい',
    elements: el
  };
}

/**
 * 4 重点案件の状況・課題：★の案件と、期待値の大きい進行中の案件（summary.focus）。
 * 支援の要否を決めたい案件（停滞・受注月超過・受注月の延期が多い・課題に決裁者／上席）は赤枠。ラップアップMTG（BX本部内会議）の「支援が必要な案件」にもつながる
 */
function mtgFocusSlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const columns = [130, 50, 84, 52, 98, 32, 92, 32, 102];
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '期待値／\n想定', 'ステータス\n（滞留）', '受注月', '活動状況', '', '活動における課題', '', '決めること'])];
  const borders = [];
  s.focus.forEach(d => {
    if (d.support) borders.push({row: rows.length, col: 0, rows: 1, cols: columns.length, color: c.red});
    rows.push(mtgRow_(columns, [
      mtgDealCell_(d, columns[0], d.star),
      {text: mtgAmount_(d.weighted) + '\n' + mtgAmount_(d.amount), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: (d.status || '（空欄）') + '\n' + (d.days == null ? '—' : d.days + '日' + (d.atLeast ? '以上' : '')), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: mtgMonthLabel_(d.orderMonthText) + (d.slips ? '\n延期' + d.slips + '回' : d.overdue ? '\n超過' : ''), lines: 2, fill: c.gray,
        align: 'CENTER', color: d.overdue ? c.red : c.body},
      {text: d.activity || '—', lines: 3, fill: c.gray},
      mtgArrowCell_(),
      {text: d.issue || '—', lines: 3, fill: c.gray},
      mtgArrowCell_(),
      {text: d.decide, lines: 3, bold: true, color: c.ink, fill: c.gray}
    ], 44));
  });
  const el = [];
  if (s.focus.length) {
    el.push({type: 'table', x: S.x, y: S.top, columns, rows, borders});
  } else {
    el.push(mtgBox_(S.x, S.top + 20, S.w, 30, '進行中の案件はありません', {size: 10, color: c.muted, align: 'CENTER'}));
  }
  const share = k.weighted > 0 ? Math.round(mtgSum_(s.focus, 'weighted') / k.weighted * 100) : 0;
  return {
    title: MTG_TITLES.focus,
    lead: s.focus.length
      ? '期待値の大きい進行中の案件' + (k.stars ? 'と★の案件' : '') + 'の ' + s.focus.length + '件（進行中の期待値の ' + share + '%）。' +
        (k.stalled ? '停滞の全 ' + k.stalled + '件は付録B。' : '停滞している案件はなし。') +
        '\n赤枠の ' + k.support + '件は支援の要否を決めたい案件（停滞・受注月の遅れ・課題に上席など）。' + (k.stars ? '★はキャンバスで深掘りする候補。' : '')
      : '進行中の案件はありません。',
    band: k.support ? '赤枠の ' + k.support + '件は、支援の要否（上席の同行・クローズ）を決めたい'
      : s.focus.length ? '重点案件の次のアクションと受注月の見通しを確認したい' : '進行中の案件はありません',
    elements: el
  };
}

/** 付録C 要対応（責任者別）：責任者 × 種類の件数・種類の説明・試作から解消した件数（要対応が無ければ作らない） */
function mtgAppendixOwners_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  if (!s.alerts.length) return [];
  const types = MTG_ALERT_TYPES.filter(t => s.alerts.some(a => a.type === t));
  const groups = new Map();
  s.alerts.forEach(a => {
    const owner = a.owner || '（責任者が空）';
    if (!groups.has(owner)) groups.set(owner, []);
    groups.get(owner).push(a);
  });
  const owners = Array.from(groups.keys()).sort((a, b) => groups.get(b).length - groups.get(a).length);
  const limit = MTG_OPTIONS.mainRows.owners;
  const lines = owners.slice(0, owners.length > limit ? limit - 1 : limit).map(owner => ({label: owner, list: groups.get(owner)}));
  if (owners.length > limit) {
    const rest = owners.slice(limit - 1);
    lines.push({label: 'ほか ' + rest.length + '人', list: [].concat.apply([], rest.map(owner => groups.get(owner)))});
  }
  lines.push({label: '合計', list: s.alerts, total: true});
  const columns = [96].concat(types.map(() => 52), [40]);
  const short = {'予定超過': '予定\n超過', '受注月超過': '受注月\n超過', '更新なし': '更新\nなし', '金額の矛盾': '金額の\n矛盾', '日付の矛盾': '日付の\n矛盾',
    '未定義ステータス': '未定義の\nステータス', '体制未登録': '体制\n未登録'};
  const rows = [mtgHeadRow_(columns, ['責任者'].concat(types.map(t => short[t] || t), ['計']), 6.5)];
  lines.forEach(line => {
    const fill = line.total ? c.purpleSoft : c.gray;
    rows.push(mtgRow_(columns, [{text: line.label, bold: true, color: c.ink, fill: c.purpleSoft}]
      .concat(types.map(t => {
        const n = line.list.filter(a => a.type === t).length;
        return {text: n ? String(n) : '', fill, align: 'CENTER', bold: !!line.total};
      }), [{text: String(line.list.length), fill, align: 'CENTER', bold: true, color: c.ink}]), 20));
  });
  const el = [{type: 'table', x: S.x, y: S.top, columns, rows}];
  const tableW = columns.reduce((a, b) => a + b, 0);
  const noteX = S.x + tableW + 16;
  const noteW = S.w - tableW - 16;
  el.push(mtgBox_(noteX, S.top, noteW, 16, '要対応の種類', {size: 8.5, bold: true, color: c.purple}));
  let y = S.top + 18;
  types.forEach(t => {
    const text = mtgWrap_(t + '：' + mtgAlertNote_(t), noteW, 7, 3);
    const h = mtgLineHeight_(text.length, 7);
    if (y + h > S.appendixBottom) return;
    el.push(mtgBox_(noteX, y, noteW, h, text.join('\n'), {size: 7, color: c.body, valign: 'TOP'}));
    y += h;
  });
  const r = s.resolved;
  return [{
    title: MTG_TITLES.owners,
    lead: '要対応は ' + s.alerts.length + '件（責任者 ' + k.owners + '人）。' +
      (r ? '試作（' + Utilities.formatDate(new Date(r.at), Session.getScriptTimeZone(), 'MM/dd') + '）の時点から ' + r.fixed + '件を解消・新たに ' + r.added + '件。' : '') +
      '\n会議の前に、責任者ごとに修正を依頼する（案件ごとの一覧は次のページ）。',
    elements: el
  }];
}

/** 付録A 前回からの変化（一覧。更新のみ・削除も） */
function mtgAppendixChanges_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [150, 50, 118, 72, 84, 146, 52];
  return mtgPagedTable_({
    title: MTG_TITLES.changes,
    lead: mtgPeriodShort_(s) + '変わった案件は ' + s.changes.length + '件。' +
      (s.unmatched ? '案件に結び付けられなかった変更の記録 ' + s.unmatched + '件は数えていない。' : '') +
      '\n区分は案件ごとに1つ（新規 > 受注 > 失注 > 保留 > 復帰 > 前進 > 後退 > 延期 > 前倒し > 増額 > 減額 > 変更 > 更新のみ）。前進・後退はプルダウンの並び順で判定。',
    columns,
    head: ['得意先／案件名', '区分', 'ステータス（前▶今）', '期待値\n（前▶今）', '受注月\n（前▶今）', '更新した項目', '責任者'],
    rows: s.changes.map(ch => ({bad: mtgCategory_(ch.category).bad, cells: [
      {text: mtgDealLabel_(ch), lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: mtgCategory_(ch.category).label, bold: true, color: mtgCategory_(ch.category).bad ? c.red : c.purple, fill: c.gray, align: 'CENTER'},
      {text: mtgStatusMove_(ch), lines: 2, fill: c.gray},
      {text: mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: mtgMonthMove_(ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: ch.columns.map(mtgShortHeader_).join('・') || (ch.removed ? '（行を削除）' : '—'), lines: 2, fill: c.gray},
      {text: ch.owner || '—', fill: c.gray, align: 'CENTER'}
    ]}))
  });
}

/** 付録B 停滞している案件（一覧） */
function mtgAppendixStalled_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [150, 44, 44, 84, 58, 56, 34, 56, 146];
  return mtgPagedTable_({
    title: MTG_TITLES.stalled,
    lead: '進行中で、ステータスが' + MTG_OPTIONS.stallDays + '日以上変わっていない案件は ' + s.stalled.length + '件（想定 ' +
      mtgAmount_(s.kpi.stalledAmount) + '百万円。想定の大きい順）。\n「以上」は記録より前から変わっていないもの。★（赤枠）はオポチュニティキャンバスで深掘りする候補。',
    columns,
    head: ['得意先／案件名', '想定', '期待値', 'ステータス', '滞留', '受注月', '延期', '責任者', '活動における課題'],
    rows: s.stalled.map(d => ({bad: d.star, cells: [
      {text: (d.star ? '★ ' : '') + mtgDealLabel_(d), lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: mtgAmount_(d.amount), fill: c.gray, align: 'END'},
      {text: mtgAmount_(d.weighted), fill: c.gray, align: 'END'},
      {text: d.status || '（空欄）', fill: c.gray},
      {text: d.days + '日' + (d.atLeast ? '以上' : ''), fill: c.gray, align: 'CENTER'},
      {text: mtgMonthLabel_(d.orderMonthText), fill: c.gray, align: 'CENTER'},
      {text: d.slips ? d.slips + '回' : '', fill: c.gray, align: 'CENTER'},
      {text: d.owner || '—', fill: c.gray, align: 'CENTER'},
      {text: d.issue || '—', lines: 2, fill: c.gray}
    ]}))
  });
}

/** 付録C 要対応（一覧） */
function mtgAppendixAlerts_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [170, 70, 296, 52, 84];
  return mtgPagedTable_({
    title: MTG_TITLES.alerts,
    lead: '要対応は ' + s.alerts.length + '件。\n種類の順（' + MTG_ALERT_TYPES.join('・') + '）。種類の中は想定の大きい順。',
    columns,
    head: ['得意先／案件名', '種類', '内容', '想定', '責任者'],
    rows: s.alerts.map(a => ({cells: [
      {text: mtgDealLabel_(a), lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: a.type, bold: true, color: c.red, fill: c.gray, align: 'CENTER'},
      {text: a.text, lines: 2, fill: c.gray},
      {text: mtgAmount_(a.amount), fill: c.gray, align: 'END'},
      {text: a.owner || '—', fill: c.gray, align: 'CENTER'}
    ]}))
  });
}

/** 付録D ステータスの分類（件数・金額。期待値・前進／後退の数え方） */
function mtgAppendixStatus_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [150, 60, 44, 70, 70, 278];
  const kinds = {active: '進行中', won: '受注', lost: '失注', hold: '保留'};
  const notes = {
    active: '期待値・パイプラインに数える。プルダウンの順で前進・後退を判定',
    won: 'パイプラインから外す（今期の受注に数える）',
    lost: 'パイプラインから外す',
    hold: 'パイプラインから外す（進行中に戻ったら「復帰」）'
  };
  const rows = s.stages.map(st => ({bad: !st.known, cells: [
    {text: st.status, bold: true, color: c.ink, fill: c.purpleSoft},
    {text: kinds[st.kind], fill: c.gray, align: 'CENTER'},
    {text: String(st.count), fill: c.gray, align: 'END'},
    {text: mtgAmount_(st.amount), fill: c.gray, align: 'END'},
    {text: mtgAmount_(st.weighted), fill: c.gray, align: 'END'},
    {text: (st.known ? '' : 'プルダウンにない値。') + notes[st.kind], fill: c.gray}
  ]}));
  rows.push({cells: [
    {text: '合計', bold: true, color: c.ink, fill: c.purpleSoft}, {text: '', fill: c.purpleSoft},
    {text: String(mtgSum_(s.stages, 'count')), bold: true, fill: c.purpleSoft, align: 'END'},
    {text: mtgAmount_(mtgSum_(s.stages, 'amount')), bold: true, fill: c.purpleSoft, align: 'END'},
    {text: mtgAmount_(mtgSum_(s.stages, 'weighted')), bold: true, fill: c.purpleSoft, align: 'END'},
    {text: '', fill: c.purpleSoft}
  ]});
  return mtgPagedTable_({
    title: MTG_TITLES.status,
    lead: '案件 ' + s.kpi.deals + '件のうち、進行中 ' + s.kpi.active + '件（期待値 ' + mtgAmount_(s.kpi.weighted) + '百万円）。' +
      (s.stages.some(st => !st.known) ? 'プルダウンにない値（赤枠）がある。' : '') +
      '\n分類はステータスのことばで判定（MeetingSummary.gs の MTG_OPTIONS で変えられる）。並びは案件ステータスのプルダウンの順。',
    columns,
    head: ['案件ステータス', '分類', '件数', '想定', '期待値', '数え方'],
    rows
  });
}

/**
 * 表を、入りきる行ずつのページに分ける（見出しの行は各ページに）。行が無ければページを作らない。
 * spec = {title, lead, columns, head, rows: [{cells, bad}], band}。bad の行は赤枠。band（結論の帯）があれば帯の上まで（無ければ付録の下端まで）
 */
function mtgPagedTable_(spec) {
  const S = MTG_SLIDE;
  const red = MTG_OPTIONS.colors.red;
  const head = mtgHeadRow_(spec.columns, spec.head);
  const bottom = spec.band ? S.bottom : S.appendixBottom;
  const pages = [];
  let page = null;
  spec.rows.forEach(line => {
    const row = mtgRow_(spec.columns, line.cells.map(cell => Object.assign({size: S.appendixSize}, cell)), 0);
    if (!page || page.y + row.h > bottom) {
      page = {rows: [head], borders: [], y: S.top + head.h};
      pages.push(page);
    }
    if (line.bad) page.borders.push({row: page.rows.length, col: 0, rows: 1, cols: spec.columns.length, color: red});
    page.rows.push(row);
    page.y += row.h;
  });
  return pages.map((p, i) => ({
    title: spec.title + (i ? '（続き）' : ''),
    lead: spec.lead,
    band: spec.band,
    elements: [{type: 'table', x: S.x, y: S.top, columns: spec.columns, rows: p.rows, borders: p.borders}]
  }));
}

/* ---------------- 全体サマリー（事務局のサマリ報告） ---------------- */

/**
 * すべての追跡シートを1つにまとめた資料（事務局が会議の冒頭で報告する全体サマリー）。summaries はシートの並び（DIFF_RULES の順）。
 * 1 全体サマリー：シートごとの数字の表（入力状況として、前回から更新した案件の数と最終更新日時も）
 * 2 本日の論点（サブインダストリー別）：シートごとの、いちばん上の論点と決めてほしいこと
 * info = {meeting, final, now}
 */
function mtgOverallSlides_(summaries, info) {
  const o = MTG_OPTIONS;
  const frame = {
    tag: o.title + '｜' + o.overallName + '（' + mtgShortDate_(info.meeting) + ' 会議）' + (info.final ? '' : '　試作'),
    source: '出典：追跡シート ' + summaries.length + '枚・' + DIFF_OPTIONS.logSheet + '（' +
      Utilities.formatDate(info.now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' 時点）　期間：シートごとに、前回の会議の資料を作った時点以降'
  };
  return mtgNumberSlides_(frame, mtgOverallTable_(summaries).concat(mtgOverallIssues_(summaries)));
}

/** 全体サマリーの表（シートごとの1行と合計）。前回比がマイナス・前回から更新の無いシートは赤字 */
function mtgOverallTable_(summaries) {
  const c = MTG_OPTIONS.colors;
  const tz = Session.getScriptTimeZone();
  const sum = (key, part) => summaries.reduce((total, s) => total + (s[part][key] || 0), 0);
  const delta = s => s.bridge.end - s.bridge.start;
  const quiet = s => !s.changes.length;
  const red = bad => (bad ? c.red : c.body);
  const line = (name, v, total) => ({cells: [
    {text: name, bold: true, color: c.ink, fill: c.purpleSoft},
    {text: String(v.active), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total},
    {text: mtgAmount_(v.weighted), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total},
    {text: mtgSigned_(v.delta), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total, color: red(v.delta < -0.05)},
    {text: mtgAmount_(v.won), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total},
    {text: mtgAmount_(v.landing), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total},
    {text: v.wonCount + '／' + v.lost, fill: total ? c.purpleSoft : c.gray, align: 'CENTER', bold: total},
    {text: v.forward + '／' + v.backward, fill: total ? c.purpleSoft : c.gray, align: 'CENTER', bold: total},
    {text: String(v.stalled), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total},
    {text: String(v.alerts), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total},
    {text: String(v.updated), fill: total ? c.purpleSoft : c.gray, align: 'END', bold: total || v.updated === 0, color: red(v.updated === 0)},
    {text: v.last == null ? '—' : Utilities.formatDate(new Date(v.last), tz, 'MM/dd HH:mm'), fill: total ? c.purpleSoft : c.gray, align: 'CENTER',
      bold: total, color: red(v.stale)}
  ]});
  const values = s => ({active: s.kpi.active, weighted: s.bridge.end, delta: delta(s), won: s.fiscal.wonAmount, landing: s.fiscal.landing,
    wonCount: s.kpi.won, lost: s.kpi.lost, forward: s.kpi.forward, backward: s.kpi.backward, stalled: s.kpi.stalled, alerts: s.kpi.alerts,
    updated: s.changes.length, last: s.lastUpdated, stale: quiet(s)});
  const totals = {active: sum('active', 'kpi'), weighted: summaries.reduce((t, s) => t + s.bridge.end, 0),
    delta: summaries.reduce((t, s) => t + delta(s), 0), won: sum('wonAmount', 'fiscal'), landing: sum('landing', 'fiscal'),
    wonCount: sum('won', 'kpi'), lost: sum('lost', 'kpi'), forward: sum('forward', 'kpi'), backward: sum('backward', 'kpi'),
    stalled: sum('stalled', 'kpi'), alerts: sum('alerts', 'kpi'), updated: summaries.reduce((t, s) => t + s.changes.length, 0),
    last: summaries.reduce((m, s) => (s.lastUpdated != null && (m == null || s.lastUpdated > m) ? s.lastUpdated : m), null), stale: false};
  const rows = summaries.map(s => line(s.name, values(s), false)).concat([line('合計', totals, true)]);
  const most = key => summaries.filter(s => s.kpi[key] > 0).sort((a, b) => b.kpi[key] - a.kpi[key])[0];
  const stalled = most('stalled');
  const alerts = most('alerts');
  const quietNames = summaries.filter(quiet).map(s => s.name);
  return mtgPagedTable_({
    title: MTG_TITLES.overall,
    lead: '全体（' + summaries.length + 'シート）の期待値は ' + mtgAmount_(totals.weighted) + '百万円（前回比 ' + mtgSigned_(totals.delta) + '）。今期の受注 ' +
      mtgAmount_(totals.won) + '百万円・着地見込み ' + mtgAmount_(totals.landing) + '百万円。\n受注 ' + totals.wonCount + '件・失注 ' + totals.lost +
      '件・停滞 ' + totals.stalled + '件・要対応 ' + totals.alerts + '件' + (quietNames.length ? '。前回から更新の無いシート：' + quietNames.join('・') : '') + '。',
    columns: [104, 40, 54, 50, 52, 56, 52, 52, 40, 44, 52, 76],
    head: ['サブインダストリー', '進行中', '期待値', '前回比', '今期の\n受注', '着地\n見込み', '受注／\n失注', '前進／\n後退', '停滞', '要対応',
      '更新した\n案件', '最終更新'],
    rows,
    band: stalled && alerts && stalled !== alerts
      ? '停滞の多い「' + stalled.name + '」（' + stalled.kpi.stalled + '件）と、要対応の多い「' + alerts.name + '」（' + alerts.kpi.alerts + '件）を重点的に確認したい'
      : stalled || alerts ? '「' + (stalled || alerts).name + '」の停滞 ' + (stalled || alerts).kpi.stalled + '件・要対応 ' + (stalled || alerts).kpi.alerts +
        '件を重点的に確認したい'
        : '各サブインダストリーの論点を確認したい'
  });
}

/** 本日の論点（サブインダストリー別）：シートごとの、いちばん上の論点 ▶ 状況 ▶ 決めてほしいこと */
function mtgOverallIssues_(summaries) {
  const c = MTG_OPTIONS.colors;
  return mtgPagedTable_({
    title: MTG_TITLES.overallIssues,
    lead: 'サブインダストリーごとの、いちばん上の論点と決めてほしいこと（' + summaries.length + 'シート。発表の順）。' +
      '\n状況の詳しい内容は、各サブインダストリーの資料の「サマリー」。',
    columns: [104, 120, 268, 180],
    head: ['サブインダストリー', '論点', '状況', '決めてほしいこと'],
    rows: summaries.map(s => {
      const issue = mtgIssues_(s)[0];
      return {cells: [
        {text: s.name, lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
        {text: issue.title.replace(/\n/g, '') + (issue.count ? ' ' + issue.count + '件' : ''), lines: 2, bold: true, color: c.white, fill: c.purple},
        {text: issue.status, lines: 2, fill: c.gray},
        {text: issue.ask.split('\n')[0], lines: 2, bold: true, color: c.ink, fill: c.gray}
      ]};
    })
  });
}

/** ステータスの動き（「提案中 ▶ 受注」。変わっていなければ今の値） */
function mtgStatusMove_(ch) {
  const blank = v => v || '（空欄）';
  if (ch.removed) return blank(ch.statusFrom) + ' ▶ 行を削除';
  if (!ch.base) return '新規 ▶ ' + blank(ch.statusTo);
  if (credNormalize_(ch.statusFrom) === credNormalize_(ch.statusTo)) return blank(ch.statusTo);
  return blank(ch.statusFrom) + ' ▶ ' + blank(ch.statusTo);
}

/** 金額の動き（「10 ▶ 15」。変わっていなければ今の値） */
function mtgValueMove_(from, to, ch) {
  if (ch.removed) return mtgAmount_(from) + ' ▶ —';
  if (!ch.base || (from || 0) === (to || 0)) return mtgAmount_(to);
  return mtgAmount_(from) + ' ▶ ' + mtgAmount_(to);
}

/** 受注月の動き（「2026/11 ▶ 2026/12」） */
function mtgMonthMove_(ch) {
  if (ch.removed) return mtgMonthLabel_(ch.monthFrom);
  if (!ch.base || credNormalize_(ch.monthFrom) === credNormalize_(ch.monthTo)) return mtgMonthLabel_(ch.monthTo);
  return mtgMonthLabel_(ch.monthFrom) + ' ▶ ' + mtgMonthLabel_(ch.monthTo);
}

/** 受注月の表示（「2026年11月」「2026/11/01」→「2026/11」。読めなければそのまま。空は「—」） */
function mtgMonthLabel_(text) {
  const m = mtgMonth_(text);
  return m ? m.y + '/' + m.m : (String(text || '').trim() || '—');
}

/** 変更した項目の見出しを短く（「期待値調整済 想定売上規模（百万）」→「期待値」。MTG_OPTIONS.columns にないものは空白を除いた見出し） */
function mtgShortHeader_(header) {
  const names = {customer: '得意先', project: '案件名', status: 'ステータス', amount: '想定', weighted: '期待値', orderMonth: '受注月',
    startMonth: '売上開始月', activity: '活動状況', issue: '課題', owner: '責任者', delivery: 'デリバリー担当'};
  const labels = mtgLabels_();
  const key = Object.keys(labels).find(k => labels[k] === credNormalize_(header));
  return key ? names[key] : String(header).replace(/\s+/g, '');
}

/** 「前回の会議（10/9）の資料（10/07 18:00）以降」「初回のため直近14日（09/30〜）」 */
function mtgPeriodText_(s) {
  const tz = Session.getScriptTimeZone();
  if (s.previous) {
    return '前回の会議（' + mtgShortDate_(s.previous.date) + '）の資料を作った ' + Utilities.formatDate(new Date(s.previous.cutoff), tz, 'MM/dd HH:mm') + ' 以降';
  }
  if (s.legacy) return '前回の作成（' + Utilities.formatDate(s.since, tz, 'MM/dd HH:mm') + '）以降';
  return '初回のため直近' + MTG_OPTIONS.firstDays + '日（' + Utilities.formatDate(s.since, tz, 'MM/dd') + '〜）';
}

/* ---------------- スライドの部品 ---------------- */

/** 文字の箱 */
function mtgBox_(x, y, w, h, text, style) {
  return Object.assign({type: 'box', x, y, w, h, text}, style || {});
}

/** 表の矢印のセル */
function mtgArrowCell_() {
  return {text: '▶', size: 9, color: MTG_OPTIONS.colors.arrow, align: 'CENTER'};
}

/** 得意先（1行目）と案件名（2行目）のセル */
function mtgDealCell_(item, width, star) {
  const c = MTG_OPTIONS.colors;
  const size = MTG_SLIDE.bodySize;
  const customer = (star ? '★ ' : '') + (item.customer || '（得意先なし）');
  return {text: mtgWrap_(customer, width, size, 1)[0] + '\n' + mtgWrap_(item.project || '—', width, size, 1)[0],
    lines: 2, bold: true, color: c.ink, fill: c.purpleSoft};
}

/**
 * 表の1行。cells: [{text, lines, size, bold, color, fill, align}]（文字は列の幅で折り返し、lines 行（既定は1）を超えたら「…」）。
 * 行の高さは、いちばん行数の多いセルに合わせる（min 以上）。
 */
function mtgRow_(columns, cells, min) {
  let h = min || 0;
  const out = cells.map((cell, i) => {
    const size = cell.size || MTG_SLIDE.bodySize;
    const lines = mtgWrap_(cell.text, columns[i], size, cell.lines || 1);
    h = Math.max(h, mtgLineHeight_(lines.length, size));
    return Object.assign({}, cell, {text: lines.join('\n'), size});
  });
  return {h: Math.ceil(h * 10) / 10, cells: out};
}

/** 表の見出しの行（グレーに白の太字） */
function mtgHeadRow_(columns, labels, size) {
  const c = MTG_OPTIONS.colors;
  return mtgRow_(columns, labels.map(text => ({text, size: size || 7.5, bold: true, color: c.white, fill: c.grayHead, align: 'CENTER', lines: 3})),
    MTG_SLIDE.headRow);
}

/** タイトルを1行に収める（20 → 18 → 16 ポイント。それでも入らなければ「…」で切る） */
function mtgTitleFit_(text) {
  const S = MTG_SLIDE;
  const sizes = [S.titleSize, S.titleSize - 2, S.titleSize - 4];
  const size = sizes.find(pt => mtgWrap_(text, S.w, pt, 2).length === 1) || sizes[sizes.length - 1];
  return {text: mtgWrap_(text, S.w, size, 1)[0], size};
}

/** 文字の幅の見積もり（pt。全角は文字の大きさ、半角は 0.6 倍） */
function mtgTextWidth_(text, size) {
  return Array.from(String(text == null ? '' : text)).reduce((sum, ch) => sum + (/[\u0000-\u00ff\uff61-\uff9f]/.test(ch) ? 0.6 : 1), 0) * size;
}

/**
 * 文を、幅 widthPt の箱（セル）の中に収まる行に分ける（改行はそのまま行の区切り。pt で数える）。max 行を超えたら最後の行を「…」の分をあけて切る。
 * 文字の幅は見積もり（全角は文字の大きさ、半角は 0.6 倍。箱の左右の余白 MTG_SLIDE.padX を引く）なので、本物の折り返しより早めに分ける。
 * 行の頭に来てはいけない文字（、。）」など）は、本当の幅に収まるなら前の行の終わりにぶら下げ、収まらなければ前の文字と一緒に次の行へ送る。
 * 半角の語（「2026/12」など）の途中では切らず、語の前の空白で折り返す。
 */
function mtgWrap_(text, widthPt, size, max) {
  const room = (widthPt - MTG_SLIDE.padX) / (size || MTG_SLIDE.bodySize);
  const limit = room * 0.92;
  const width = ch => (/[\u0000-\u00ff\uff61-\uff9f]/.test(ch) ? 0.6 : 1);
  const hang = /[、。，．,.)）」』】〕〉》・ーぁぃぅぇぉっゃゅょァィゥェォッャュョ！？!?%％:：;；]/;
  const out = [];
  String(text == null ? '' : text).split('\n').forEach(part => {
    let line = '';
    let used = 0;
    Array.from(part.replace(/[ \t]+/g, ' ').trim()).forEach(ch => {
      if (used + width(ch) > limit && line && !(hang.test(ch) && used + width(ch) <= room)) {
        const chars = Array.from(line);
        const word = hang.test(ch) ? null : /[!-~]/.test(ch) ? /^(.*\S) ([!-~]+)$/.exec(line) : null;
        const keep = word ? word[1] : hang.test(ch) && chars.length > 1 ? chars.slice(0, -1).join('') : line;
        out.push(keep);
        line = word ? word[2] : line.slice(keep.length);
        used = Array.from(line).reduce((sum, c) => sum + width(c), 0);
      }
      line += ch;
      used += width(ch);
    });
    out.push(line);
  });
  const keep = max || 1;
  if (out.length <= keep) return out;
  const kept = out.slice(0, keep);
  let last = '';
  let used = 0;
  Array.from(kept[keep - 1]).some(ch => {
    if (used + width(ch) > limit - 1) return true;
    last += ch;
    used += width(ch);
    return false;
  });
  kept[keep - 1] = last + '…';
  return kept;
}

/** 行の高さ（pt）の見積もり：lines 行・文字の大きさ size（上下の余白を含む） */
function mtgLineHeight_(lines, size) {
  return lines * (size || MTG_SLIDE.bodySize) * MTG_SLIDE.lineFactor + MTG_SLIDE.padY;
}

/* ---------------- スライド（16:9）に書く ---------------- */

/*
 * スライドの部品（位置・大きさは pt。スライドは MTG_SLIDE.width × height）
 * - 箱：{type: 'box', x, y, w, h, text, size, bold, color, fill, line, lineWeight, align: 'START'|'CENTER'|'END', valign: 'TOP'|'MIDDLE'|'BOTTOM',
 *        shape: 'TEXT_BOX'|'RECTANGLE'}
 * - 表：{type: 'table', x, y, columns: [列の幅], rows: [{h, cells: [{text, size, bold, color, fill, align}]}],
 *        borders: [{row, col, rows, cols, color, weight}]}（区切り線は白。borders はそのあとに外枠として付ける）
 *   列の幅は 32pt 以上にする（Slides API は、それより狭い列の幅をエラーにする）。
 */

/** 色（#7030A0）を Slides API の形に。 */
function mtgRgb_(hex) {
  const n = parseInt(String(hex).replace('#', ''), 16);
  return {red: ((n >> 16) & 255) / 255, green: ((n >> 8) & 255) / 255, blue: (n & 255) / 255};
}

function mtgPt_(value) {
  return {magnitude: Math.round(value * 100) / 100, unit: 'PT'};
}

/** 文字の書式（Slides API の updateTextStyle の style と fields）。 */
function mtgTextStyle_(item) {
  return {
    style: {
      fontFamily: MTG_SLIDE.font,
      fontSize: mtgPt_(item.size || MTG_SLIDE.bodySize),
      bold: !!item.bold,
      foregroundColor: {opaqueColor: {rgbColor: mtgRgb_(item.color || MTG_OPTIONS.colors.body)}}
    },
    fields: 'fontFamily,fontSize,bold,foregroundColor'
  };
}

/** 段落の書式（寄せ・行間）。 */
function mtgParagraphStyle_(item) {
  return {
    style: {alignment: item.align || 'START', lineSpacing: 100, spaceAbove: mtgPt_(0), spaceBelow: mtgPt_(0)},
    fields: 'alignment,lineSpacing,spaceAbove,spaceBelow'
  };
}

/**
 * 1枚のスライドを作る Slides API のリクエスト（batchUpdate に渡す）。pageId はスライドの ID（5文字以上）、index は何枚目か（0 から）。
 * 表の空のセルにも空白を1つ入れて文字の大きさをそろえる（空のままだと既定の大きな文字の高さで行が伸びる）。
 */
function mtgSlideRequests_(slide, pageId, index) {
  const requests = [{createSlide: {objectId: pageId, insertionIndex: index, slideLayoutReference: {predefinedLayout: 'BLANK'}}}];
  const styleText = (target, item, text) => {
    const font = mtgTextStyle_(item);
    const para = mtgParagraphStyle_(item);
    requests.push({insertText: Object.assign({text, insertionIndex: 0}, target)});
    requests.push({updateTextStyle: Object.assign({textRange: {type: 'ALL'}, style: font.style, fields: font.fields}, target)});
    requests.push({updateParagraphStyle: Object.assign({textRange: {type: 'ALL'}, style: para.style, fields: para.fields}, target)});
  };
  const fill = color => (color ? {solidFill: {color: {rgbColor: mtgRgb_(color)}}} : {propertyState: 'NOT_RENDERED'});
  slide.elements.forEach((el, n) => {
    const id = pageId + '_' + n;
    const width = el.type === 'table' ? el.columns.reduce((a, b) => a + b, 0) : el.w;
    const height = el.type === 'table' ? el.rows.reduce((sum, r) => sum + r.h, 0) : el.h;
    const elementProperties = {
      pageObjectId: pageId,
      size: {width: mtgPt_(width), height: mtgPt_(height)},
      transform: {scaleX: 1, scaleY: 1, translateX: Math.round(el.x * 100) / 100, translateY: Math.round(el.y * 100) / 100, unit: 'PT'}
    };
    if (el.type === 'table') {
      requests.push({createTable: {objectId: id, elementProperties, rows: el.rows.length, columns: el.columns.length}});
      el.columns.forEach((w, c) => requests.push({updateTableColumnProperties: {objectId: id, columnIndices: [c],
        tableColumnProperties: {columnWidth: mtgPt_(w)}, fields: 'columnWidth'}}));
      el.rows.forEach((row, r) => {
        requests.push({updateTableRowProperties: {objectId: id, rowIndices: [r], tableRowProperties: {minRowHeight: mtgPt_(row.h)}, fields: 'minRowHeight'}});
        row.cells.forEach((cell, c) => {
          const location = {rowIndex: r, columnIndex: c};
          requests.push({updateTableCellProperties: {objectId: id, tableRange: {location, rowSpan: 1, columnSpan: 1},
            tableCellProperties: {tableCellBackgroundFill: fill(cell.fill), contentAlignment: 'MIDDLE'},
            fields: 'tableCellBackgroundFill,contentAlignment'}});
          styleText({objectId: id, cellLocation: location}, cell, cell.text == null || cell.text === '' ? ' ' : String(cell.text));
        });
      });
      const all = {location: {rowIndex: 0, columnIndex: 0}, rowSpan: el.rows.length, columnSpan: el.columns.length};
      [{range: all, position: 'ALL', color: MTG_OPTIONS.colors.white, weight: 1.5}].concat((el.borders || []).map(b => ({
        range: {location: {rowIndex: b.row, columnIndex: b.col}, rowSpan: b.rows, columnSpan: b.cols}, position: 'OUTER',
        color: b.color, weight: b.weight || 2.25
      }))).forEach(b => requests.push({updateTableBorderProperties: {objectId: id, tableRange: b.range, borderPosition: b.position,
        tableBorderProperties: {tableBorderFill: {solidFill: {color: {rgbColor: mtgRgb_(b.color)}}}, weight: mtgPt_(b.weight), dashStyle: 'SOLID'},
        fields: 'tableBorderFill,weight,dashStyle'}}));
      return;
    }
    requests.push({createShape: {objectId: id, shapeType: el.shape || 'TEXT_BOX', elementProperties}});
    requests.push({updateShapeProperties: {objectId: id, shapeProperties: {
      shapeBackgroundFill: fill(el.fill),
      outline: el.line
        ? {outlineFill: {solidFill: {color: {rgbColor: mtgRgb_(el.line)}}}, weight: mtgPt_(el.lineWeight || 2.25), dashStyle: 'SOLID'}
        : {propertyState: 'NOT_RENDERED'},
      contentAlignment: el.valign || 'MIDDLE'
    }, fields: 'shapeBackgroundFill,outline,contentAlignment'}});
    if (el.text != null && el.text !== '') styleText({objectId: id}, el, String(el.text));
  });
  return requests;
}

/** Google Slides API（拡張サービス）が使えるか（使えなければ、追加のしかたを知らせて止める） */
function mtgRequireSlides_() {
  if (typeof Slides === 'undefined') {
    throw new Error('Google Slides API（拡張サービス）が有効になっていません。エディタの左の「サービス」の＋から「Google Slides API」を追加してから、もう一度お試しください。');
  }
}

/** ファイル名（会議サマリ_スーパー・GMS_2026-10-23。試作は「_試作」つき） */
function mtgFileName_(name, meeting, final) {
  return MTG_OPTIONS.fileName + '_' + name + '_' + meeting.replace(/\//g, '-') + (final ? '' : '_試作');
}

/**
 * スライドの並び（mtgSlides_・mtgOverallSlides_）を、Google スライド（16:9）と PDF にしてフォルダに保存する。
 * スライドは Slides API（拡張サービス「Google Slides API」）で作り、できたスライドを PDF に書き出す。
 * MTG_OPTIONS.keepSlides なら Google スライドもフォルダに残す（推進リーダーが書き足せる）。うまくいかなかったときはゴミ箱に移す。
 * 戻り値 {pdf, deck}（deck は残さないときは null）
 */
function mtgExportDeck_(name, slides, folder) {
  const deck = Slides.Presentations.create({title: name,
    pageSize: {width: mtgPt_(MTG_SLIDE.width), height: mtgPt_(MTG_SLIDE.height)}});
  const file = DriveApp.getFileById(deck.presentationId);
  try {
    // 何枚かずつまとめて送る（作ったときの白紙のスライドは、1回目の終わりに消す）
    let requests = [];
    let first = true;
    const flush = () => {
      if (first) (deck.slides || []).forEach(s => requests.push({deleteObject: {objectId: s.objectId}}));
      Slides.Presentations.batchUpdate({requests}, deck.presentationId);
      requests = [];
      first = false;
    };
    slides.forEach((slide, i) => {
      const add = mtgSlideRequests_(slide, 'mtg_p' + (i + 1), i);
      if (requests.length && requests.length + add.length > MTG_OPTIONS.batchRequests) flush();
      requests = requests.concat(add);
    });
    flush();
    file.moveTo(folder);
    const pdf = folder.createFile(file.getAs('application/pdf').setName(name + '.pdf'));
    if (!MTG_OPTIONS.keepSlides) file.setTrashed(true);
    return {pdf, deck: MTG_OPTIONS.keepSlides ? file : null};
  } catch (e) {
    try { file.setTrashed(true); } catch (_) {}
    throw e;
  }
}

/** 保存先のフォルダ（MTG_OPTIONS.folderId。空なら、このスプレッドシートと同じフォルダの中に folderName を作り、次からも使う）。 */
function mtgFolder_(ss) {
  if (MTG_OPTIONS.folderId) return DriveApp.getFolderById(MTG_OPTIONS.folderId);
  const props = PropertiesService.getDocumentProperties();
  const saved = props.getProperty(MTG_OPTIONS.folderKey);
  if (saved) {
    try {
      const folder = DriveApp.getFolderById(saved);
      if (!folder.isTrashed()) return folder;
    } catch (_) {}
  }
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder(MTG_OPTIONS.folderName);
  props.setProperty(MTG_OPTIONS.folderKey, folder.getId());
  return folder;
}
