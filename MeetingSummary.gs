/**
 * 会議サマリ（追跡シートごとの会議資料。16:9 のスライドと PDF）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs・HistorySidebar.gs と同じプロジェクトに置くファイル。
 *
 * 目的
 * - RX Business MTG（パイプライン会議。月1回）の事前資料。会議の前に配り、参加者（営業の統括・リーダー、BX、サブインダストリーのリーダー）が
 *   通読してから会議に臨む。発表は1サブインダストリー数分（報告＋フィードバック）なので、会議の目的（案件の進捗・受注見込みを確かめ、
 *   優先順位と次のアクションを決める）に向けて、状況を事前に共有しておく。
 * - 資料には、シートと変更の記録から分かる事実だけを書く。論点・決めてほしいこと・支援の要否などの判断は、各サブインダストリーのリーダーが
 *   会議で話すことなので、資料には入れない（資料が判断を先に書くと、読み手を誤った方向に導くおそれがある）。
 * - 資料はサブインダストリー（追跡シート）ごとに1つ。スライドの見出しはどのシートでも同じ（MTG_TITLES）で、状態は見出しの下のリード文に書く。
 *   使う言葉はAppendix ⑤（用語の定義）にまとめ、資料の中で同じ意味で使う。
 *
 * 資料の構成（16:9。社内資料のトンマナ：紫の見出し・リード文・表・ページ番号）
 *   表紙：会議の名前と「事前資料」・サブインダストリー・会議日・データの時点・期間
 *   1 サマリー：今期の受注と着地見込み・4つの数字（期待値と前回比・受注／失注・前進／後退・停滞）・前回からの状況の内訳（区分ごとの件数と案件）
 *   2 パイプラインの変化：前回からの期待値の増減（新規・増額・復帰・減額・受注・失注・保留・削除）・ステータス別・受注月（四半期）別の期待値
 *   3 前回から変わった案件：ステータス・期待値・受注月・活動状況が変わった案件（区分の順に上位）
 *   4 期待値の大きい案件の状況・課題：進行中の案件の期待値の大きい順（ステータスが同じ日数・受注月・活動状況・課題・当てはまる状況）
 *   扉（Appendix。Appendix の見出しとページ）
 *   Appendix（全件。入りきらなければ続きのページ）
 *   A 前回から変わった案件　B 停滞している案件　C 入力内容のチェック（責任者別・一覧）　D ステータスの分類　E 用語の定義
 *   左上のラベルは置かない（どのサブインダストリー・会議日の資料かは、表紙・ファイル名・下の出典の行に書く。試作は表紙と出典の行に【試作】）
 *
 * ことばの決まり（どのページでも同じ数え方にする。Appendix ⑤ にも出す）
 * - ステータスの分類：進行中・受注・失注・保留（MTG_OPTIONS のことば・statusKinds で決める）。期待値は進行中の案件だけを数える。
 *   前進・後退は、案件ステータスのプルダウンの並び順で決める。
 * - 前回からの変化は、案件ごとに1つの区分にする（新規 > 受注 > 失注 > 保留 > 復帰 > 前進 > 後退 > 延期 > 前倒し > 増額 > 減額 > 変更 > 活動の更新 > 更新のみ）。
 *   行を削除した案件は「削除」（失注には数えない）。
 * - 停滞：進行中で、ステータスが最後に変わってから stallDays 日以上同じ案件（変更の記録・基準より前から同じなら「N日以上」）。
 *
 * 仕組み
 * - 管理者がエディタで createMeetingSummaries() を実行し、作るシートと今回の会議日を入力する（試すときは previewMeetingSummaries()）。
 *   シートごとに Google スライドと PDF をドライブのフォルダに保存する（MTG_OPTIONS.folderId。空なら、このスプレッドシートと
 *   同じフォルダの中に MTG_OPTIONS.folderName のフォルダを作る）。配る（共有する）のは手動。
 * - 「前回から」は、前回の会議（今回の会議日より前で、いちばん新しい会議日）の資料を作った日時から。会議日と作った日時は
 *   シートごとにドキュメントのプロパティ（MTG_MEETINGS_シート名）に残す（同じ会議日で作り直したら、作った日時を上書きする）。
 *   前回の会議の記録が無ければ、MTG_OPTIONS.firstDays 日前から。previewMeetingSummaries() は会議日を記録しないので、何度でも試せる
 *   （ファイル名に「試作」が付く）。
 * - 旧フォーマットからの移行など、案件をまとめて貼り付けたあとは setMeetingBaseline() で今を「基準」として記録する（会議の記録と同じ
 *   プロパティに {baseline: true} で残す。シートごとに1つ）。基準が前回の会議より後なら、次の資料は基準より後の変化だけを出す
 *   （貼り付けた案件を「新規」に数えない）。
 * - 会議日ごとに最初の試作の入力内容のチェックの結果を覚えておき（MTG_DRAFT_シート名）、同じ会議日の資料に、試作から解消した件数を出す
 *   （試作を各リーダーに見せて、会議の前に入力を直してもらう運用）。
 * - スプレッドシートは読むだけで、書き込まない。変更の記録は、変更時点スナップショット（その時の行の内容）で得意先＋案件名に結び付ける
 *   （行の追加・削除で行番号がずれても合う。得意先・案件名を書き換えた案件・あとから入れた案件も、前の名前の記録をつなげる）。
 *   どちらの記録も追記なので、下（新しいほう）から、読む期間（前回と MTG_OPTIONS.lookbackDays 日前の古いほう）より前に届くまでだけ読む。
 * - スライドは Slides API（拡張サービス「Google Slides API」。appsscript.json に書いてある）で作り、PDF に書き出す。
 */

const MTG_OPTIONS = {
  meetingName: 'RX Business MTG',    // 表紙に出す会議の名前
  fileName: '会議サマリ',             // ファイル名の先頭（会議サマリ_スーパー・GMS_2026-10-23.pdf。日付は会議日）
  folderId: '',                      // 保存先のフォルダの ID（空なら、このスプレッドシートと同じフォルダの中に folderName を作る）
  folderName: '会議サマリ',
  folderKey: 'MTG_FOLDER_ID',         // 作ったフォルダの ID（ドキュメントのプロパティ）
  meetingsKeyPrefix: 'MTG_MEETINGS_', // シートごとの会議日と資料を作った日時（ドキュメントのプロパティ。後ろにシート名）
  draftKeyPrefix: 'MTG_DRAFT_',       // シートごとの、会議日ごとの最初の試作の入力内容のチェックの結果（解消した件数を出す用）
  lastKeyPrefix: 'MTG_LAST_',         // 前の版の「前回の作成日時」（会議日の記録が無いときだけ使う）
  keepMeetings: 24,                  // 会議日の記録をいくつ残すか
  keepSlides: true,                  // Google スライドもフォルダに残す（推進リーダーが書き足せる。false なら PDF だけ）
  firstDays: 14,                     // 初回（前回の会議の記録が無い）は、何日前からの変化を出すか
  lookbackDays: 180,                 // ステータスの滞留・受注月の延期の回数を数えるために、何日前までの記録を読むか
  stallDays: 45,                     // 停滞：ステータスがこの日数以上同じ進行中の案件
  noUpdateDays: 30,                  // 入力内容のチェック「更新なし」：最終更新日時がこの日数より前の進行中の案件
  overdueDays: 90,                   // 入力内容のチェック「予定日超過」：予定日を過ぎてこの日数以内のもの（古いものは出さない）
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
    delivery: 'デリバリー担当',
    record: '商談記録（ＵＲＬ）'          // 提案資料・商談の音声の格納先（入力内容のチェック「資料なし」）
  },
  // 必須項目：進行中の案件で空なら、入力内容のチェックの「未入力」に出す（シートに無い見出しは見ない）
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
  mainRows: {changes: 7, focus: 5, owners: 9, deals: 4},   // 本編の表に出す行の数（残りは Appendix）。owners は Appendix ③ の責任者別の表の行、
                                                           // deals はサマリーの内訳の1区分に名前を出す案件の数
  // 色（社内資料のトンマナ：紫を基調に、薄い紫・薄いグレーの面。赤は失注・後退・延期・期限を過ぎたものの文字）
  colors: {
    purple: '#7030A0', purpleSoft: '#EEE3F5', up: '#A77BCA', down: '#BFBFBF', gray: '#F2F2F2', grayHead: '#A6A6A6',
    red: '#E60012', ink: '#1B1B1C', body: '#262626', muted: '#595959', faint: '#8C8C8C', white: '#FFFFFF'
  }
};

/* 前回からの変化の区分（上ほど優先。案件ごとに1つ）。red は赤字で出す区分（失注・後退・延期）。cell は表の狭い列での書き方。意味は Appendix ⑤（mtgGlossary_） */
const MTG_CATEGORIES = [
  {key: 'new', label: '新規'},
  {key: 'won', label: '受注'},
  {key: 'lost', label: '失注', red: true},
  {key: 'hold', label: '保留'},
  {key: 'revived', label: '復帰'},
  {key: 'forward', label: '前進'},
  {key: 'backward', label: '後退', red: true},
  {key: 'slipped', label: '延期', red: true},
  {key: 'pulled', label: '前倒し'},
  {key: 'up', label: '増額'},
  {key: 'down', label: '減額'},
  {key: 'status', label: '変更'},
  {key: 'removed', label: '削除'},
  {key: 'activity', label: '活動の更新', cell: '活動の\n更新'},
  {key: 'updated', label: '更新のみ'}
];

/*
 * スライドの見出し（タイトル）。どのサブインダストリーでも同じにする（状態はリード文に書く）。
 * 表紙のあとに本編：サマリー → パイプラインの変化 → 前回から変わった案件 → 期待値の大きい案件の状況・課題。扉（Appendix）のあとに、全件の一覧と言葉の定義
 */
const MTG_TITLES = {
  cover: '表紙',
  summary: 'サマリー',
  pipeline: 'パイプラインの変化',
  progress: '前回から変わった案件',
  focus: '期待値の大きい案件の状況・課題',
  appendix: 'Appendix',
  changes: 'Appendix ①　前回から変わった案件（一覧）',
  stalled: 'Appendix ②　停滞している案件（一覧）',
  owners: 'Appendix ③　入力内容のチェック（責任者別）',
  alerts: 'Appendix ③　入力内容のチェック（一覧）',
  status: 'Appendix ④　ステータスの分類',
  glossary: 'Appendix ⑤　用語の定義'
};

/* 本編の「前回から変わった案件」に出す区分の順（更新のみは出さない） */
const MTG_MOVE_ORDER = ['won', 'lost', 'backward', 'slipped', 'forward', 'new', 'revived', 'hold', 'removed', 'pulled', 'up', 'down', 'status',
  'activity'];

/* 期待値の増減の内訳（増える → 減るの順） */
const MTG_BRIDGE = [
  {key: 'new', label: '新規'}, {key: 'up', label: '増額'}, {key: 'revived', label: '復帰'},
  {key: 'down', label: '減額'}, {key: 'won', label: '受注'}, {key: 'lost', label: '失注'}, {key: 'hold', label: '保留'},
  {key: 'removed', label: '削除'}
];

/* 入力内容のチェックの種類（この順に出す。意味は mtgAlertNote_） */
const MTG_ALERT_TYPES = ['予定日超過', '受注月超過', '未入力', '資料なし', '更新なし', '金額の不整合', '月の不整合', '選択肢外のステータス', 'デリバリー担当なし'];

/*
 * スライドの形（pt。16:9 = 720 × 405）。上から 見出し（titleY）・リード文（leadY）・中身（top〜bottom）・出典の行とページ番号（footY）。
 * 表紙と扉は、見出し・リード文の代わりに大きな文字を置く（表紙には出典の行・ページ番号も置かない）。
 * 文字の幅と行の高さは見積もり（mtgWrap_・mtgLineHeight_）。本物より大きめに見積もって、はみ出さないようにする。
 */
const MTG_SLIDE = {
  width: 720,
  height: 405,
  font: 'BIZ UDPGothic',  // BIZ UDPゴシック（Google スライド・Google Fonts での名前）
  x: 24,                 // 左の余白
  w: 672,                // 中身の幅
  titleY: 16,            // 見出しの上端
  leadY: 50,             // リード文の上端
  top: 88,               // 中身の上端
  bottom: 376,           // 中身の下端（下に出典の行とページ番号）
  footY: 384,            // 出典の行の上端
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

/**
 * 今を会議サマリの「基準」として記録する（作るシートを聞く）。旧フォーマットからの移行など、案件をまとめて貼り付けたあとに実行する。
 * 次に作る資料は、基準より後の変化だけを出す（貼り付けた案件を「新規」に数えない）。会議日は記録しない。もう一度実行したら基準を今に置き換える。
 */
function setMeetingBaseline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const names = mtgAskSheets_(ss, ui, '会議サマリの基準を記録します（会議日は記録しません）。');
  if (!names) return;
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const at = Utilities.formatDate(now, tz, 'MM/dd HH:mm');
  const answer = ui.alert(CRED_OPTIONS.menuTitle,
    '「' + names.join('」「') + '」に、今（' + at + '）を基準として記録します。\n' +
    '次に作る会議サマリは、基準より後の変化だけを出します（今までに貼り付けた案件を「新規」に数えません）。\n' +
    '移行の貼り付けがすべて終わってから実行してください。記録してよろしいですか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const props = PropertiesService.getDocumentProperties();
  const date = Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  names.forEach(name => {
    const list = mtgMeetings_(props, name).filter(m => !m.baseline).concat([{date, cutoff: now.getTime(), baseline: true}])
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).slice(-MTG_OPTIONS.keepMeetings);
    props.setProperty(MTG_OPTIONS.meetingsKeyPrefix + name, JSON.stringify(list));
  });
  const message = '「' + names.join('」「') + '」に、基準（' + at + '）を記録しました。次に作る会議サマリは、基準より後の変化を出します。';
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

function mtgRun_(final) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const names = mtgAskSheets_(ss, ui, final ? '会議サマリを作ります。' : '会議サマリを試しに作ります（会議日は記録しません）。');
  if (!names) return;
  const now = new Date();
  const meeting = mtgAskMeeting_(ui, names, now);
  if (!meeting) return;
  const results = mtgCreate_(ss, names, now, final, meeting);
  const lines = results.map(r => '「' + r.summary.name + '」：' + r.file.getName() + '（' + mtgCountsText_(r.summary) + '）');
  lines.push('', '保存先：' + results[0].folder.getUrl());
  if (MTG_OPTIONS.keepSlides) lines.push('同じフォルダに、Google スライド（同じ名前）も保存しました。');
  lines.push(final
    ? '会議日（' + meeting + '）を記録しました。次の会議の資料は、今以降の変化を出します。配るときは、保存先のフォルダから共有してください。'
    : '試作のため、会議日は記録していません。');
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/** シートを聞く（空欄ならすべて。「1,3」のように複数も）。heading は最初の行。やめたら null。 */
function mtgAskSheets_(ss, ui, heading) {
  const all = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  if (!all.length) throw new Error('追跡しているシート（test.gs の DIFF_RULES）が見つかりません。');
  const message = heading +
    '\nシートの番号を入力してください（空欄ならすべて。複数は「1,3」のように）。\n\n' +
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
  const tz = Session.getScriptTimeZone();
  const last = names.map(name => {
    const list = mtgMeetings_(props, name);
    const held = list.filter(m => !m.baseline).pop();
    const base = list.filter(m => m.baseline).pop();
    return '・' + name + '：' + (held ? held.date : '記録なし') +
      (base && (!held || base.cutoff > held.cutoff)
        ? '（' + (held ? 'そのあと ' : '') + Utilities.formatDate(new Date(base.cutoff), tz, 'MM/dd HH:mm') + ' に基準を記録）' : '');
  });
  const message = '今回の会議日を入力してください（例：10/23）。\n' +
    '前回の会議（今回より前で、いちばん新しい会議日）の資料を作った時点（そのあとに基準を記録していれば基準）からの変化を出します。\n\n' +
    '記録してある最後の会議日\n' + last.join('\n');
  const response = ui.prompt(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const text = String(response.getResponseText() || '').trim();
  if (!text) throw new Error('会議日を入力してください（例：10/23）。');
  const date = mtgParseDate_(text, now);
  if (!date) throw new Error('「' + text + '」を会議日として読めません。10/23 か 2026/10/23 のように入力してください。');
  return date;
}

/** 完成した結果の件数の知らせ（「変わった案件 12・停滞 5・チェック 7」） */
function mtgCountsText_(summary) {
  return '変わった案件 ' + summary.kpi.changed + '・停滞 ' + summary.stalled.length + '・チェック ' + summary.alerts.length;
}

/**
 * names のシート（サブインダストリー）ごとに会議サマリを作り、スライドと PDF を保存する。meeting は今回の会議日（yyyy/MM/dd）。
 * final なら会議日と今の日時を記録する。試作なら、その会議日の最初の試作の入力内容のチェックの結果を覚える。
 * 戻り値 [{summary, file: PDF, deck: Google スライド（残さないときは null）, folder}]
 */
function mtgCreate_(ss, names, now, final, meeting) {
  const o = MTG_OPTIONS;
  const props = PropertiesService.getDocumentProperties();
  const plans = names.map(name => {
    const meetings = mtgMeetings_(props, name);
    // 前回：今回より前でいちばん新しい会議。そのあとに基準を記録していれば（今回の会議日まで）基準
    const held = meetings.filter(m => !m.baseline && m.date < meeting).pop() || null;
    const base = meetings.filter(m => m.baseline && m.date <= meeting).pop() || null;
    const previous = base && (!held || base.cutoff > held.cutoff) ? base : held;
    const legacy = previous ? 0 : Number(props.getProperty(o.lastKeyPrefix + name)) || 0;
    const since = previous ? previous.cutoff : (legacy || now.getTime() - o.firstDays * 86400000);
    const draft = mtgReadJson_(props, o.draftKeyPrefix + name);
    const baseline = meetings.filter(m => m.baseline && m.cutoff <= now.getTime()).pop();
    return {name, meetings, previous, legacy: legacy || null, since: new Date(since), baseline: baseline ? baseline.cutoff : null,
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
    since: p.since, previous: p.previous, legacy: p.legacy, baseline: p.baseline, draft: p.draft, final
  }, context)));
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  const results = summaries.map(summary => {
    const out = mtgExportDeck_(mtgFileName_(summary.name, meeting, final), mtgSlides_(summary), folder);
    return {summary, file: out.pdf, deck: out.deck, folder};
  });
  plans.forEach((p, i) => {
    if (final) {
      const list = p.meetings.filter(m => m.baseline || m.date !== meeting).concat([{date: meeting, cutoff: now.getTime()}])
        .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).slice(-o.keepMeetings);
      props.setProperty(o.meetingsKeyPrefix + p.name, JSON.stringify(list));
    } else if (!p.draft) {
      props.setProperty(o.draftKeyPrefix + p.name, JSON.stringify({
        meeting, at: now.getTime(), keys: summaries[i].alerts.map(a => mtgHash_(a.key))
      }));
    }
  });
  return results;
}

/**
 * シートの会議の記録 [{date: 'yyyy/MM/dd', cutoff: 資料を作った時刻（ms）}]（会議日の古い順）。
 * setMeetingBaseline() の基準も同じ形で {date: 記録した日, cutoff: 記録した時刻, baseline: true}（1つだけ）
 */
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

/** 短い文字のハッシュ（入力内容のチェックの結果を覚えておく用。プロパティの大きさを抑える） */
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
 *  startMonth: {y, m}|null, startMonthText, activity, issue, owner, delivery, record, updatedAt: 時刻（ms）|null, missing: [見出し]}]}
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
      record: text('record'),
      updatedAt: stampCol >= 0 ? mtgTime_(values[n][stampCol]) : null,
      missing: required.filter(r => !String(line[r.col]).trim()).map(r => r.label)
    });
  });
  return out;
}

/* ---------------- まとめる ---------------- */

/**
 * 1つのシートの会議サマリの中身。
 * context = {since, previous, legacy, baseline: 基準を記録した時刻|null, draft, final, now, meeting, windowStart, trackedSince, log, snapshots, credentials}
 * 戻り値 {name, meeting, previous, legacy, since, first, final, now, windowStart, kpi, fiscal, bridge, stages, quarters,
 *         changes, stalled, focus: 期待値の大きい進行中の案件, overdue: 受注月を過ぎた進行中の案件, alerts: 入力内容のチェック,
 *         checks: このシートで見た入力内容のチェックの種類, resolved, unmatched}
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

  // 進行中の案件の、ステータスが同じ日数：最後に変わってから（または案件ができてから）。変わった記録が無ければ、
  // 記録を読んだ期間の始まり・最終更新日時の古いほうからの「N日以上」（どちらも無ければ分からない）。
  // 最後に変わった記録が基準より前なら「N日以上」（移行で貼り付けた行は、貼り付けた時に「案件ができた」記録になるため）
  const baseline = context.baseline == null ? null : context.baseline;
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
    const atLeast = at == null || (baseline != null && at <= baseline);
    if (at == null) {
      const bounds = [tracked, deal.updatedAt].filter(v => v != null);
      at = bounds.length ? Math.min.apply(null, bounds) : null;
    }
    const days = at == null ? null : Math.floor((now - at) / 86400000);
    const stalled = days != null && days >= o.stallDays;
    const overdue = !!deal.orderMonth && mtgMonthIndex_(deal.orderMonth) < month;
    items.push(Object.assign({}, deal, {days, atLeast, slips, overdue, stalled}));
  });
  // 停滞：ステータスが stallDays 日以上同じ（想定売上の大きい順）
  const stalled = items.filter(d => d.stalled)
    .sort((a, b) => ((b.amount || 0) - (a.amount || 0)) || ((b.weighted || 0) - (a.weighted || 0)) || (b.days - a.days));
  // 期待値の大きい進行中の案件（mainRows.focus 件まで）
  const focus = items.slice().sort((a, b) => ((b.weighted || 0) - (a.weighted || 0)) || ((b.amount || 0) - (a.amount || 0)))
    .slice(0, o.mainRows.focus);

  // 入力内容のチェックと、試作から解消した件数
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
      activity: count('activity'),
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
      stalled: stalled.length,
      stalledAmount: mtgSum_(stalled, 'amount'),
      overdue: items.filter(d => d.overdue).length,
      alerts: alerts.length,
      owners: owners.size
    },
    fiscal: mtgFiscal_(name, deals, context.now),
    bridge,
    stages: mtgStages_(deals, table.statusOrder),
    quarters: mtgQuarters_(active, context.now),
    changes,
    stalled,
    focus,
    overdue: items.filter(d => d.overdue).sort((a, b) => (b.amount || 0) - (a.amount || 0)),
    alerts,
    checks: MTG_ALERT_TYPES.filter(t => (t !== '資料なし' || table.has.record) && (t !== 'デリバリー担当なし' || table.has.delivery)),
    resolved,
    unmatched: history.unmatched
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
    monthText: pick(labels.orderMonth, cur.orderMonthText, String),
    activity: pick(labels.activity, cur.activity || '', String)
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
    if (credNormalize_(base.activity) !== credNormalize_(deal.activity)) flags.push('activity');
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
 * 入力内容のチェック [{type, key, dealKey, customer, project, text, amount, owner}]（MTG_ALERT_TYPES の順。種類の中は想定売上の大きい順。
 * 予定日超過は過ぎた日数の多い順）。
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

  // 予定日超過：このシートに進行中の案件がある得意先の、予定日を過ぎても実施日が空のクレデンシャル・オファリング
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
      overdue.push({type: '予定日超過', key: '予定日超過|' + c.id, dealKey: '', customer, project: '', owner: c.values.owner || '', amount: null, days,
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
    if (table.has.record && !d.record) push('資料なし', d, '「' + o.columns.record + '」が空');
    if (d.updatedAt != null) {
      const days = Math.floor((now - d.updatedAt) / 86400000);
      if (days >= o.noUpdateDays) push('更新なし', d, '最終更新から ' + days + '日');
    } else if (info.tracked != null && now - info.tracked >= o.noUpdateDays * 86400000) {
      push('更新なし', d, '変更の記録を始めてから更新なし（' + Math.floor((now - info.tracked) / 86400000) + '日以上）');
    }
    if (d.amount != null && d.weighted != null && d.weighted > d.amount) {
      push('金額の不整合', d, '期待値 ' + mtgAmount_(d.weighted) + ' が想定売上 ' + mtgAmount_(d.amount) + ' より大きい');
    }
  });
  deals.forEach(d => {
    if ((d.kind === 'active' || d.kind === 'won') && d.orderMonth && d.startMonth &&
        mtgMonthIndex_(d.startMonth) < mtgMonthIndex_(d.orderMonth)) {
      push('月の不整合', d, '売上開始月 ' + label(d.startMonth) + ' が受注月 ' + label(d.orderMonth) + ' より前');
    }
    if (table.statusOrder.length && d.status && info.rank(d.status) < 0) {
      push('選択肢外のステータス', d, '「' + d.status + '」はプルダウンの選択肢にない（' + {active: '進行中', won: '受注', lost: '失注', hold: '保留'}[d.kind] + 'として数える）');
    }
  });
  if (table.has.delivery) {
    info.changes.filter(c => c.category === 'won' && c.deal && !c.deal.delivery)
      .forEach(c => push('デリバリー担当なし', c.deal, '受注したが「' + MTG_OPTIONS.columns.delivery + '」が空'));
  }
  const byAmount = (a, b) => (b.amount || 0) - (a.amount || 0);
  return overdue.concat(MTG_ALERT_TYPES.slice(1).reduce((list, type) => list.concat(out.filter(a => a.type === type).sort(byAmount)), []));
}

/** 入力内容のチェックの種類の説明（Appendix ③ の右に出す） */
function mtgAlertNote_(type) {
  const o = MTG_OPTIONS;
  return {
    '予定日超過': 'クレデンシャル・オファリングの予定日を過ぎても実施日が空（過ぎて' + o.overdueDays + '日以内）',
    '受注月超過': '受注月を過ぎても進行中のまま',
    '未入力': '進行中の案件で、' + o.required.map(h => h.replace(/\s+/g, '')).join('・') + ' のどれかが空',
    '資料なし': '進行中の案件で、「' + o.columns.record + '」が空（提案資料・商談の音声の格納先の URL）',
    '更新なし': '進行中で、最終更新日時が' + o.noUpdateDays + '日以上前',
    '金額の不整合': '期待値調整済 想定売上規模が、想定売上規模より大きい',
    '月の不整合': '売上開始月が受注月より前',
    '選択肢外のステータス': '案件ステータスが、プルダウンの選択肢にない値',
    'デリバリー担当なし': '前回から受注になったが、' + o.columns.delivery + 'が空'
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
 * 1つのシート（サブインダストリー）の会議サマリをスライドの並びにする（Slides API に送る前の形。テストと見た目の確認にも使う）。
 * 表紙、本編（サマリー → パイプラインの変化 → 前回から変わった案件 → 期待値の大きい案件の状況・課題）、扉（Appendix）、Appendix ①〜⑤。
 * ページ番号は表紙を1枚目として数える（表紙には出さない。PDF のページと同じ番号になる）。
 * 見出しは MTG_TITLES（どのシートでも同じ）。書くのはシートと変更の記録から分かる事実だけ（論点・決めてほしいことなどの判断は、会議でリーダーが話す）。
 * 戻り値 [{title, elements}]（elements は mtgSlideRequests_ の部品。タイトル・リード文・出典・ページ番号も含む）
 */
function mtgSlides_(summary) {
  const cover = mtgCoverSlide_(summary);
  const main = [mtgSummarySlide_(summary), mtgBridgeSlide_(summary), mtgMovesSlide_(summary), mtgFocusSlide_(summary)];
  const appendix = [].concat(mtgAppendixChanges_(summary), mtgAppendixStalled_(summary), mtgAppendixOwners_(summary),
    mtgAppendixAlerts_(summary), mtgAppendixStatus_(summary), mtgAppendixGlossary_(summary));
  const divider = appendix.length ? [mtgDividerSlide_(appendix, 1 + main.length + 1 + 1)] : [];
  return mtgNumberSlides_(mtgSheetFrame_(summary), [cover].concat(main, divider, appendix));
}

/** 表紙：会議の名前と「事前資料」・サブインダストリー・会議日・データの時点・期間（試作は【試作】） */
function mtgCoverSlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const p = s.meeting.split('/').map(Number);
  const week = '日月火水木金土'.charAt(new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay());
  const columns = [72, S.w - 72];
  const rows = [
    ['会議日', p[0] + '年' + p[1] + '月' + p[2] + '日（' + week + '）'],
    ['データ', Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' 時点の「' + s.name + '」シートと' + DIFF_OPTIONS.logSheet],
    ['期間', mtgPeriodText_(s) + 'の変化']
  ].map(([label, text]) => mtgRow_(columns, [{text: label, bold: true, color: c.purple}, {text, color: c.body}], 0));
  return {
    title: MTG_TITLES.cover,
    cover: true,
    elements: [
      mtgBox_(S.x, 118, S.w, 24, o.meetingName + '　事前資料' + (s.final ? '' : '　【試作】'), {size: 13, bold: true, color: c.purple}),
      mtgBox_(S.x, 142, S.w, 52, mtgWrap_(s.name, S.w, 32, 1)[0], {size: 32, bold: true, color: c.ink}),
      {type: 'box', x: S.x + 7.2, y: 200, w: S.w - 7.2, h: 2, fill: c.purple, shape: 'RECTANGLE'},
      {type: 'table', x: S.x, y: 212, columns, rows}
    ]
  };
}

/**
 * 扉（Appendix）：大きな見出しと、Appendix の見出し・始まりのページの一覧。appendix は Appendix のスライド、firstPage はその1枚目のページ番号
 */
function mtgDividerSlide_(appendix, firstPage) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const columns = [96, 516, 60];
  const rows = [];
  appendix.forEach((slide, i) => {
    if (/（続き）$/.test(slide.title)) return;
    const m = /^(\S+ \S+)　(.*)$/.exec(slide.title) || [slide.title, slide.title, ''];
    rows.push(mtgRow_(columns, [
      {text: m[1], bold: true, color: c.purple, fill: c.purpleSoft},
      {text: m[2], fill: c.gray},
      {text: String(firstPage + i), fill: c.gray, align: 'END'}
    ], 22));
  });
  return {
    title: MTG_TITLES.appendix,
    divider: true,
    elements: [
      mtgBox_(S.x, 120, S.w, 56, MTG_TITLES.appendix, {size: 36, bold: true, color: c.purple, valign: 'BOTTOM'}),
      {type: 'box', x: S.x + 7.2, y: 180, w: S.w - 7.2, h: 2, fill: c.purple, shape: 'RECTANGLE'},
      {type: 'table', x: S.x, y: 194, columns, rows}
    ]
  };
}

/** スライドに、どのスライドにも付ける部品（mtgFrame_）とページ番号を付ける */
function mtgNumberSlides_(frame, slides) {
  return slides.map((slide, i) => ({title: slide.title, elements: mtgFrame_(frame, slide, i + 1).concat(slide.elements)}));
}

/** シートの資料のいちばん下の行（どのサブインダストリー・会議日の資料か・出典・期間。試作は先頭に【試作】）{source} */
function mtgSheetFrame_(s) {
  return {
    source: (s.final ? '' : '【試作】') + s.name + '（' + mtgShortDate_(s.meeting) + ' 会議）　データ：シートと' + DIFF_OPTIONS.logSheet + '（' +
      Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' 時点）　期間：' + mtgPeriodText_(s)
  };
}

/** どのスライドにも付ける部品：タイトル・リード文（扉には無い）・いちばん下の出典の行・ページ番号（表紙には何も付けない）。frame = {source} */
function mtgFrame_(frame, slide, page) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const el = [];
  if (slide.cover) return el;
  if (!slide.divider) {
    const title = mtgTitleFit_(slide.title);
    el.push(mtgBox_(S.x, S.titleY, S.w, 34, title.text, {size: title.size, bold: true, color: c.purple}));
  }
  if (slide.lead) {
    el.push(mtgBox_(S.x, S.leadY, S.w, 34, mtgWrap_(slide.lead, S.w, S.leadSize, 2).join('\n'),
      {size: S.leadSize, bold: true, color: c.ink, valign: 'TOP'}));
  }
  el.push(mtgBox_(S.x, S.footY, S.w - 60, 16, mtgWrap_(frame.source, S.w - 60, 6.5, 1)[0], {size: 6.5, color: c.faint}));
  el.push(mtgBox_(S.x + S.w - 60, S.footY - 2, 60, 18, String(page), {size: 9, color: c.muted, align: 'END'}));
  return el;
}

/** 「前回から」「直近14日で」（リード文の書き出し） */
function mtgPeriodShort_(s) {
  return s.first ? '直近' + MTG_OPTIONS.firstDays + '日で' : '前回から';
}

/** 1 サマリー：今期の帯・4つの数字・前回からの状況の内訳（区分ごとの件数・想定売上・案件） */
function mtgSummarySlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const b = s.bridge;
  const f = s.fiscal;
  const delta = b.end - b.start;
  const el = [];

  const fy = '今期（' + f.label + '）　受注 ' + f.wonCount + '件・' + mtgAmount_(f.wonAmount) + '百万円' +
    (f.target ? '（目標 ' + mtgAmount_(f.target) + '・達成率 ' + Math.round(f.wonAmount / f.target * 100) + '%）' : '') +
    '　｜　今期に受注予定の期待値 ' + mtgAmount_(f.planned) + '百万円　｜　着地見込み ' + mtgAmount_(f.landing) + '百万円';
  el.push(mtgBox_(S.x, S.top, S.w, 20, mtgWrap_(fy, S.w, 9, 1)[0],
    {size: 9, bold: true, color: c.white, fill: c.purple, shape: 'RECTANGLE', align: 'CENTER'}));

  const tiles = [
    {label: '進行中の期待値（百万円）', value: mtgAmount_(b.end),
      sub: (s.first ? '直近' + o.firstDays + '日で ' : '前回比 ') + mtgSigned_(delta) + '・' + k.active + '件（想定売上 ' + mtgAmount_(k.amount) + '）'},
    {label: '受注／失注（' + mtgPeriodShort_(s).replace(/で$/, '') + '）', value: k.won + '件／' + k.lost + '件',
      sub: '受注した案件の想定売上 ' + mtgAmount_(k.wonAmount) + '百万円'},
    {label: '前進／後退（' + mtgPeriodShort_(s).replace(/で$/, '') + '）', value: k.forward + '件／' + k.backward + '件',
      sub: '新規 ' + k.new + '件・受注月の延期 ' + k.slips + '件'},
    {label: '停滞（ステータスが' + o.stallDays + '日以上同じ）', value: k.stalled + '件', sub: '想定売上 ' + mtgAmount_(k.stalledAmount) + '百万円'}
  ];
  const tw = (S.w - 3 * 8) / 4;
  tiles.forEach((t, i) => {
    const x = S.x + i * (tw + 8);
    const y = S.top + 28;
    el.push({type: 'box', x, y, w: tw, h: 58, fill: c.purpleSoft, shape: 'RECTANGLE'});
    el.push(mtgBox_(x, y + 2, tw, 15, mtgWrap_(t.label, tw, 7.5, 1)[0], {size: 7.5, bold: true, color: c.muted}));
    el.push(mtgBox_(x, y + 17, tw, 24, mtgWrap_(t.value, tw, 17, 1)[0], {size: 17, bold: true, color: c.purple}));
    el.push(mtgBox_(x, y + 41, tw, 15, mtgWrap_(t.sub, tw, 7, 1)[0], {size: 7, color: c.muted}));
  });

  const columns = [150, 44, 64, 414];
  const rows = [mtgHeadRow_(columns, ['区分', '件数', '想定売上', '案件（想定売上の大きい順）'])];
  mtgBreakdown_(s).forEach(line => rows.push(mtgRow_(columns, [
    {text: line.label, lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
    {text: line.items.length + '件', fill: c.gray, align: 'CENTER'},
    {text: line.items.length ? mtgAmount_(mtgSum_(line.items, 'amount')) : '—', fill: c.gray, align: 'END'},
    {text: mtgNames_(line.items, line.note), lines: 2, fill: c.gray}
  ], 0)));
  el.push({type: 'table', x: S.x, y: S.top + 96, columns, rows});

  return {
    title: MTG_TITLES.summary,
    lead: '期待値は ' + mtgAmount_(b.end) + '百万円（' + (s.first ? '直近' + o.firstDays + '日で ' : '前回比 ') + mtgSigned_(delta) + '）、進行中 ' +
      k.active + '件。' + mtgPeriodShort_(s) + '受注 ' + k.won + '件・失注 ' + k.lost + '件・前進 ' + k.forward + '件・後退 ' + k.backward +
      '件・新規 ' + k.new + '件。\nステータスが' + o.stallDays + '日以上同じ案件は ' + k.stalled + '件（想定売上 ' + mtgAmount_(k.stalledAmount) +
      '百万円）、受注月を過ぎた案件は ' + k.overdue + '件。',
    elements: el
  };
}

/**
 * サマリーの内訳の区分（どのシートでも同じ区分を、この順に出す。該当が無ければ0件）[{label, items, note: 案件の後ろの（）の中}]。
 * items は想定売上の大きい順
 */
function mtgBreakdown_(s) {
  const o = MTG_OPTIONS;
  const byAmount = list => list.slice().sort((a, b) => (b.amount || 0) - (a.amount || 0));
  const of = keys => byAmount(s.changes.filter(ch => keys.indexOf(ch.category) >= 0));
  const slipped = s.changes.filter(ch => ch.flags.indexOf('slipped') >= 0 && !ch.removed);
  const late = slipped.map(ch => Object.assign({}, ch, {note: mtgMonthLabel_(ch.monthFrom) + '→' + mtgMonthLabel_(ch.monthTo)}))
    .concat(s.overdue.filter(d => !slipped.some(ch => ch.key === d.key))
      .map(d => Object.assign({}, d, {note: mtgMonthLabel_(d.orderMonthText) + 'を過ぎて進行中'})));
  return [
    {label: '新規・前進', items: of(['new', 'forward']),
      note: ch => (ch.category === 'new' ? '新規' : ch.statusFrom + '→' + ch.statusTo)},
    {label: '受注', items: of(['won'])},
    {label: '失注・後退・削除', items: of(['lost', 'backward', 'removed']),
      note: ch => (ch.category === 'backward' ? ch.statusFrom + '→' + ch.statusTo : mtgCategory_(ch.category).label)},
    {label: '受注月の延期・超過', items: byAmount(late), note: d => d.note},
    {label: '停滞（' + o.stallDays + '日以上）', items: s.stalled, note: d => d.days + '日' + (d.atLeast ? '以上' : '')}
  ];
}

/** 案件の名前を並べる（mainRows.deals 件まで。残りは「ほか N件」）。note(案件) があれば名前の後ろに（）で付ける */
function mtgNames_(items, note) {
  if (!items.length) return '—';
  const n = MTG_OPTIONS.mainRows.deals;
  return items.slice(0, n).map(d => mtgDealLabel_(d) + (note ? '（' + note(d) + '）' : '')).join('、') +
    (items.length > n ? '　ほか ' + (items.length - n) + '件' : '');
}

/** 2 パイプラインの変化：期待値の増減（滝の図）・ステータス別・受注月（四半期）別の期待値 */
function mtgBridgeSlide_(s) {
  const S = MTG_SLIDE;
  const b = s.bridge;
  const delta = b.end - b.start;
  const moves = b.steps.filter(st => st.count);
  const biggest = moves.slice().sort((x, y) => Math.abs(y.value) - Math.abs(x.value))[0];
  const plus = mtgSum_(b.steps.filter(st => st.value > 0), 'value');
  const minus = mtgSum_(b.steps.filter(st => st.value < 0), 'value');
  const peak = s.quarters.filter(q => /Q/.test(q.label) && q.value > 0).sort((x, y) => y.value - x.value)[0];
  const el = mtgWaterfall_(s, S.x, S.top, 360, S.bottom - S.top - 8);
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
      '・減った分 ' + mtgSigned_(minus) + (biggest && biggest.value ? '、内訳で最も大きいのは' + biggest.label + '（' + biggest.count + '件・' +
        mtgSigned_(biggest.value) + '）' : '') +
      '。\n今期に受注予定の期待値は ' + mtgAmount_(s.fiscal.planned) + '百万円' + (peak ? '（' + peak.label.replace(/（.*/, '') + 'が最も多く ' + mtgAmount_(peak.value) + '）' : '') +
      (s.kpi.missingWeighted ? '。期待値が空の ' + s.kpi.missingWeighted + '件は0として計算' : '') + '。',
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

/** 3 前回から変わった案件：ステータス・期待値・受注月・活動状況が変わった案件（区分の順に上位。活動状況だけ変わった案件は最後） */
function mtgMovesSlide_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const moves = s.changes.filter(ch => ch.category !== 'updated');
  const top = moves.slice(0, MTG_OPTIONS.mainRows.changes);
  const columns = [146, 44, 130, 66, 80, 154, 52];
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '区分', 'ステータス（前▶今）', '期待値\n（前▶今）', '受注月\n（前▶今）', '活動状況（現在）', '責任者'])];
  top.forEach(ch => {
    rows.push(mtgRow_(columns, [
      mtgDealCell_(ch, columns[0]),
      {text: mtgCategory_(ch.category).cell || mtgCategory_(ch.category).label, lines: 2, bold: true,
        color: mtgCategory_(ch.category).red ? c.red : c.purple, fill: c.gray, align: 'CENTER'},
      {text: mtgStatusMove_(ch), lines: 2, fill: c.gray},
      {text: mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: mtgMonthMove_(ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: ch.activity || '—', lines: 2, fill: c.gray},
      {text: ch.owner || '—', lines: 2, fill: c.gray, align: 'CENTER'}
    ]));
  });
  const el = [];
  if (top.length) {
    el.push({type: 'table', x: S.x, y: S.top, columns, rows});
  } else {
    el.push(mtgBox_(S.x, S.top + 20, S.w, 30, mtgPeriodShort_(s) + 'ステータス・期待値・受注月・活動状況が変わった案件はありません',
      {size: 10, color: c.muted, align: 'CENTER'}));
  }
  return {
    title: MTG_TITLES.progress,
    lead: moves.length
      ? mtgPeriodShort_(s) + '変わった案件は ' + moves.length + '件（受注 ' + k.won + '・失注 ' + k.lost + '・前進 ' + k.forward + '・後退 ' +
        k.backward + '・延期 ' + k.slipped + '・新規 ' + k.new + '・活動の更新 ' + k.activity + '）。\n' +
        (moves.length > top.length ? '区分の順（受注・失注・後退・延期・前進・新規…）に上位' + top.length + '件。' : '') +
        '全件と、ほかの項目だけ更新した ' + k.updated + '件は Appendix ①。'
      : mtgPeriodShort_(s) + 'ステータス・期待値・受注月・活動状況が変わった案件はありません。\nほかの項目だけ更新した案件は ' + k.updated + '件（Appendix ①）。',
    elements: el
  };
}

/** 4 期待値の大きい案件の状況・課題：進行中の案件の期待値の大きい順（summary.focus）。当てはまる状況（停滞・受注月超過・延期）は事実だけ */
function mtgFocusSlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const columns = [140, 56, 88, 56, 136, 124, 72];
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '期待値／\n想定売上', 'ステータス\n（同じ日数）', '受注月', '活動状況', '活動における課題',
    '当てはまる\n状況'])];
  s.focus.forEach(d => rows.push(mtgRow_(columns, [
    mtgDealCell_(d, columns[0]),
    {text: mtgAmount_(d.weighted) + '\n' + mtgAmount_(d.amount), lines: 2, fill: c.gray, align: 'CENTER'},
    {text: (d.status || '（空欄）') + '\n' + (d.days == null ? '—' : d.days + '日' + (d.atLeast ? '以上' : '')), lines: 2, fill: c.gray, align: 'CENTER'},
    {text: mtgMonthLabel_(d.orderMonthText), fill: c.gray, align: 'CENTER', color: d.overdue ? c.red : c.body},
    {text: d.activity || '—', lines: 3, fill: c.gray},
    {text: d.issue || '—', lines: 3, fill: c.gray},
    {text: mtgFlags_(d).join('\n') || '—', lines: 3, fill: c.gray, align: 'CENTER', color: mtgFlags_(d).length ? c.red : c.body}
  ], 44)));
  const el = [];
  if (s.focus.length) {
    el.push({type: 'table', x: S.x, y: S.top, columns, rows});
  } else {
    el.push(mtgBox_(S.x, S.top + 20, S.w, 30, '進行中の案件はありません', {size: 10, color: c.muted, align: 'CENTER'}));
  }
  const share = k.weighted > 0 ? Math.round(mtgSum_(s.focus, 'weighted') / k.weighted * 100) : 0;
  const count = test => s.focus.filter(test).length;
  return {
    title: MTG_TITLES.focus,
    lead: s.focus.length
      ? '進行中の案件のうち、期待値の大きい ' + s.focus.length + '件（進行中の期待値の ' + share + '%）。\nこのうち停滞 ' + count(d => d.stalled) +
        '件・受注月超過 ' + count(d => d.overdue) + '件・受注月の延期あり ' + count(d => d.slips > 0) + '件。停滞の全 ' + k.stalled + '件は Appendix ②。'
      : '進行中の案件はありません。',
    elements: el
  };
}

/** 案件に当てはまる状況（停滞・受注月超過・延期n回） */
function mtgFlags_(d) {
  return [d.stalled && '停滞', d.overdue && '受注月超過', d.slips && '延期' + d.slips + '回'].filter(Boolean);
}

/** Appendix ① 前回から変わった案件（一覧。更新のみ・削除も） */
function mtgAppendixChanges_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [130, 50, 104, 60, 70, 132, 74, 52];
  return mtgPagedTable_({
    title: MTG_TITLES.changes,
    lead: mtgPeriodShort_(s) + '変わった案件は ' + s.changes.length + '件。' +
      (s.unmatched ? '案件に結び付けられなかった変更の記録 ' + s.unmatched + '件は数えていない。' : '') +
      '\n区分は案件ごとに1つ（区分の意味と、当てはまるものが複数あるときの順は Appendix ⑤）。',
    columns,
    head: ['得意先／案件名', '区分', 'ステータス（前▶今）', '期待値\n（前▶今）', '受注月\n（前▶今）', '活動状況（現在）', '更新した項目', '責任者'],
    rows: s.changes.map(ch => [
      {text: mtgDealLabel_(ch), lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: mtgCategory_(ch.category).cell || mtgCategory_(ch.category).label, lines: 2, bold: true,
        color: mtgCategory_(ch.category).red ? c.red : c.purple, fill: c.gray, align: 'CENTER'},
      {text: mtgStatusMove_(ch), lines: 2, fill: c.gray},
      {text: mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: mtgMonthMove_(ch), lines: 2, fill: c.gray, align: 'CENTER'},
      {text: ch.removed ? '（行を削除）' : ch.activity || '—', lines: 2, fill: c.gray},
      {text: ch.category === 'new' ? '（案件を追加）' : ch.columns.map(mtgShortHeader_).join('・') || (ch.removed ? '（行を削除）' : '—'), lines: 2, fill: c.gray},
      {text: ch.owner || '—', fill: c.gray, align: 'CENTER'}
    ])
  });
}

/** Appendix ② 停滞している案件（一覧） */
function mtgAppendixStalled_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [150, 44, 44, 84, 58, 56, 34, 56, 146];
  return mtgPagedTable_({
    title: MTG_TITLES.stalled,
    lead: '進行中で、ステータスが' + MTG_OPTIONS.stallDays + '日以上同じ案件は ' + s.stalled.length + '件（想定売上 ' +
      mtgAmount_(s.kpi.stalledAmount) + '百万円。想定売上の大きい順）。\n「以上」は、変更の記録を始める前（基準を記録したときは基準より前）から同じもの。',
    columns,
    head: ['得意先／案件名', '想定\n売上', '期待値', 'ステータス', '同じ日数', '受注月', '延期', '責任者', '活動における課題'],
    rows: s.stalled.map(d => [
      {text: mtgDealLabel_(d), lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: mtgAmount_(d.amount), fill: c.gray, align: 'END'},
      {text: mtgAmount_(d.weighted), fill: c.gray, align: 'END'},
      {text: d.status || '（空欄）', fill: c.gray},
      {text: d.days + '日' + (d.atLeast ? '以上' : ''), fill: c.gray, align: 'CENTER'},
      {text: mtgMonthLabel_(d.orderMonthText), fill: c.gray, align: 'CENTER', color: d.overdue ? c.red : c.body},
      {text: d.slips ? d.slips + '回' : '', fill: c.gray, align: 'CENTER'},
      {text: d.owner || '—', fill: c.gray, align: 'CENTER'},
      {text: d.issue || '—', lines: 2, fill: c.gray}
    ])
  });
}

/**
 * Appendix ③ 入力内容のチェック（責任者別）：責任者 × 種類の件数・種類の説明・試作から解消した件数（当てはまるものが無ければ作らない）。
 * 種類の説明は、表の右にすべて入るときだけ置く（入らなければ置かず、Appendix ⑤ を案内する。Appendix ⑤ にはいつも書く）
 */
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
  // 種類の説明（表の右）。すべて切れずに入るときだけ置く
  const noteX = S.x + 96 + types.length * 52 + 40 + 16;
  const noteW = S.x + S.w - noteX;
  const notes = [];
  let y = S.top + 18;
  const fits = noteW >= 100 && types.every(t => {
    const text = mtgWrap_(t + '：' + mtgAlertNote_(t), noteW, 7, 4);
    const h = mtgLineHeight_(text.length, 7);
    if (/…$/.test(text[text.length - 1]) || y + h > S.bottom) return false;
    notes.push(mtgBox_(noteX, y, noteW, h, text.join('\n'), {size: 7, color: c.body, valign: 'TOP'}));
    y += h;
    return true;
  });
  const columns = [fits ? 96 : Math.min(160, S.w - types.length * 52 - 40)].concat(types.map(() => 52), [40]);
  const short = {'予定日超過': '予定日\n超過', '受注月超過': '受注月\n超過', '資料なし': '資料\nなし', '更新なし': '更新\nなし', '金額の不整合': '金額の\n不整合',
    '月の不整合': '月の\n不整合', '選択肢外のステータス': '選択肢外の\nステータス', 'デリバリー担当なし': 'デリバリー\n担当なし'};
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
  if (fits) el.push(mtgBox_(noteX, S.top, noteW, 16, 'チェックの種類', {size: 8.5, bold: true, color: c.purple}), ...notes);
  const r = s.resolved;
  return [{
    title: MTG_TITLES.owners,
    lead: '入力内容のチェックに当てはまったのは ' + s.alerts.length + '件（責任者 ' + k.owners + '人）。' +
      (r ? '試作（' + Utilities.formatDate(new Date(r.at), Session.getScriptTimeZone(), 'MM/dd') + '）の時点から ' + r.fixed + '件が解消、新たに ' + r.added + '件。' : '') +
      '\n責任者ごとの件数' + (fits ? 'と、チェックの種類（右）' : '（チェックの種類の意味は Appendix ⑤）') + '。案件ごとの一覧は次のページ。',
    elements: el
  }];
}

/** Appendix ③ 入力内容のチェック（一覧） */
function mtgAppendixAlerts_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [170, 70, 296, 52, 84];
  return mtgPagedTable_({
    title: MTG_TITLES.alerts,
    lead: '入力内容のチェックに当てはまったのは ' + s.alerts.length + '件。\n並びは種類の順（前のページの表の列の順）、種類の中は想定売上の大きい順。',
    columns,
    head: ['得意先／案件名', '種類', '内容', '想定\n売上', '責任者'],
    rows: s.alerts.map(a => [
      {text: mtgDealLabel_(a), lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: a.type, lines: 2, bold: true, color: c.red, fill: c.gray, align: 'CENTER'},
      {text: a.text, lines: 2, fill: c.gray},
      {text: mtgAmount_(a.amount), fill: c.gray, align: 'END'},
      {text: a.owner || '—', fill: c.gray, align: 'CENTER'}
    ])
  });
}

/** Appendix ④ ステータスの分類（件数・金額・数え方） */
function mtgAppendixStatus_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = [150, 60, 44, 70, 70, 278];
  const kinds = {active: '進行中', won: '受注', lost: '失注', hold: '保留'};
  const notes = {
    active: '期待値に数える。プルダウンの順で前進・後退を判定',
    won: '期待値に数えない（今期の受注に数える）',
    lost: '期待値に数えない',
    hold: '期待値に数えない（進行中に戻ったら「復帰」）'
  };
  const rows = s.stages.map(st => [
    {text: st.status, bold: true, color: c.ink, fill: c.purpleSoft},
    {text: kinds[st.kind], fill: c.gray, align: 'CENTER'},
    {text: String(st.count), fill: c.gray, align: 'END'},
    {text: mtgAmount_(st.amount), fill: c.gray, align: 'END'},
    {text: mtgAmount_(st.weighted), fill: c.gray, align: 'END'},
    {text: (st.known ? '' : '選択肢外（プルダウンにない値）。') + notes[st.kind], fill: c.gray, color: st.known ? c.body : c.red}
  ]);
  rows.push([
    {text: '合計', bold: true, color: c.ink, fill: c.purpleSoft}, {text: '', fill: c.purpleSoft},
    {text: String(mtgSum_(s.stages, 'count')), bold: true, fill: c.purpleSoft, align: 'END'},
    {text: mtgAmount_(mtgSum_(s.stages, 'amount')), bold: true, fill: c.purpleSoft, align: 'END'},
    {text: mtgAmount_(mtgSum_(s.stages, 'weighted')), bold: true, fill: c.purpleSoft, align: 'END'},
    {text: '', fill: c.purpleSoft}
  ]);
  return mtgPagedTable_({
    title: MTG_TITLES.status,
    lead: '案件 ' + s.kpi.deals + '件のうち、進行中 ' + s.kpi.active + '件（期待値 ' + mtgAmount_(s.kpi.weighted) + '百万円）。' +
      '\n分類はステータスのことばで判定（MeetingSummary.gs の MTG_OPTIONS で変えられる）。並びは案件ステータスのプルダウンの順。',
    columns,
    head: ['案件ステータス', '分類', '件数', '想定売上', '期待値', '数え方'],
    rows
  });
}

/** Appendix ⑤ 用語の定義（資料で使う言葉の意味と数え方。mtgGlossary_） */
function mtgAppendixGlossary_(s) {
  const c = MTG_OPTIONS.colors;
  return mtgPagedTable_({
    title: MTG_TITLES.glossary,
    lead: '資料で使っている言葉の意味と数え方。\n日数・期の始まりなどは MeetingSummary.gs の MTG_OPTIONS の設定による。',
    columns: [132, 540],
    head: ['用語', '意味'],
    rows: mtgGlossary_(s).map(([term, meaning]) => [
      {text: term, lines: 2, bold: true, color: c.ink, fill: c.purpleSoft},
      {text: meaning, lines: Math.max(2, meaning.split('\n').length * 2), fill: c.gray}   // 改行で分けた意味（種類の一覧）は、それぞれ2行まで
    ])
  });
}

/** 資料で使う言葉の意味 [[用語, 意味]]（Appendix ⑤。設定の日数などを入れる） */
function mtgGlossary_(s) {
  const o = MTG_OPTIONS;
  return [
    ['前回', '前回の会議の資料を作った日時（そのあとに基準を記録していれば基準の日時。どちらの記録も無ければ、直近' + o.firstDays + '日）。' +
      'この資料の期間：' + mtgPeriodText_(s)],
    ['想定売上', '「想定売上規模（百万）」の値（百万円）'],
    ['期待値', '「期待値調整済 想定売上規模（百万）」の値（百万円）。合計は進行中の案件だけ（空は0として計算）'],
    ['進行中・受注・失注・保留', '案件ステータスの分類（Appendix ④）。受注・失注・保留のどれにも当たらないステータス（空も）は進行中'],
    ['区分', '前回からの変化の分け方。案件ごとに1つで、当てはまるものが複数あれば ' +
      MTG_CATEGORIES.map(c => c.label).filter(l => l !== '削除').join(' > ') + ' の順に選ぶ（行の削除は「削除」）'],
    ['新規', '前回より後に追加された案件'],
    ['受注・失注・保留（区分）', '前回より後に、ステータスがその分類に変わった案件'],
    ['復帰', '前回より後に、受注・失注・保留から進行中に戻った案件'],
    ['前進・後退', '進行中のまま、ステータスがプルダウンの並びで後ろ（前進）・前（後退）の値に変わった案件'],
    ['延期・前倒し', '受注月が後ろ（延期）・前（前倒し）の月に変わった案件。「延期n回」は、変更の記録がある期間の回数'],
    ['増額・減額', '期待値が増えた・減った案件'],
    ['変更', 'ステータスが変わったが、前進・後退で比べられない案件（プルダウンにない値など）'],
    ['削除', '前回より後に行が削除された案件（失注には数えない）'],
    ['活動の更新', 'ステータス・期待値・受注月は変わらず、「活動状況」が変わった案件'],
    ['更新のみ', '上のどれにも当たらず、ほかの項目だけ変わった案件'],
    ['停滞', '進行中で、ステータスが' + o.stallDays + '日以上同じ案件（変更の記録を始める前・基準を記録する前から同じなら「N日以上」）'],
    ['受注月超過', '進行中で、受注月が今月より前の案件'],
    ['今期', s.fiscal.label],
    ['着地見込み', '今期の受注（受注月が今期の、受注した案件の想定売上）＋ 今期に受注予定の期待値（受注月が今期の、進行中の案件の期待値）'],
    ['入力内容のチェック', '入力の漏れ・予定日を過ぎた予定・値の不整合など（Appendix ③）。このシートで見る種類は次のとおり\n' +
      (s.checks || MTG_ALERT_TYPES).map(type => '・' + type + '：' + mtgAlertNote_(type)).join('\n')]
  ];
}

/**
 * Appendix の表を、入りきる行ずつのページに分ける（見出しの行は各ページに）。行が無ければページを作らない。
 * spec = {title, lead, columns, head, rows: [[セル]]}
 */
function mtgPagedTable_(spec) {
  const S = MTG_SLIDE;
  const head = mtgHeadRow_(spec.columns, spec.head);
  const pages = [];
  let page = null;
  spec.rows.forEach(cells => {
    const row = mtgRow_(spec.columns, cells.map(cell => Object.assign({size: S.appendixSize}, cell)), 0);
    if (!page || page.y + row.h > S.bottom) {
      page = {rows: [head], y: S.top + head.h};
      pages.push(page);
    }
    page.rows.push(row);
    page.y += row.h;
  });
  return pages.map((p, i) => ({
    title: spec.title + (i ? '（続き）' : ''),
    lead: spec.lead,
    elements: [{type: 'table', x: S.x, y: S.top, columns: spec.columns, rows: p.rows}]
  }));
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
    startMonth: '売上開始月', activity: '活動状況', issue: '課題', owner: '責任者', delivery: 'デリバリー担当', record: '商談記録'};
  const labels = mtgLabels_();
  const key = Object.keys(labels).find(k => labels[k] === credNormalize_(header));
  return key ? names[key] : String(header).replace(/\s+/g, '');
}

/** 「前回の会議（10/9）の資料（10/07 18:00）以降」「初回のため直近14日（09/30〜）」 */
function mtgPeriodText_(s) {
  const tz = Session.getScriptTimeZone();
  if (s.previous && s.previous.baseline) {
    return '基準（' + Utilities.formatDate(new Date(s.previous.cutoff), tz, 'MM/dd HH:mm') + ' に記録）以降';
  }
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

/** 得意先（1行目）と案件名（2行目）のセル */
function mtgDealCell_(item, width) {
  const c = MTG_OPTIONS.colors;
  const size = MTG_SLIDE.bodySize;
  const customer = item.customer || '（得意先なし）';
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

/**
 * 文字1つの幅の見積もり（文字の大きさに対する割合）。BIZ UDPゴシックで測った幅より少し大きめ：
 * 全角 1（かなは 1.02 のものもあるが、mtgWrap_ の余裕に収まる）、半角の英大文字 0.75（実際は 0.72 前後）、ほかの半角 0.6（数字は 0.5）
 */
function mtgCharWidth_(ch) {
  if (/[A-Z]/.test(ch)) return 0.75;
  return /[\u0000-\u00ff\uff61-\uff9f]/.test(ch) ? 0.6 : 1;
}

/**
 * 文を、幅 widthPt の箱（セル）の中に収まる行に分ける（改行はそのまま行の区切り。pt で数える）。max 行を超えたら最後の行を「…」の分をあけて切る。
 * 文字の幅は見積もり（mtgCharWidth_。箱の左右の余白 MTG_SLIDE.padX を引き、さらに 8% の余裕をとる）なので、本物の折り返しより早めに分ける。
 * 行の頭に来てはいけない文字（、。）」など）は、本当の幅に収まるなら前の行の終わりにぶら下げ、収まらなければ前の文字と一緒に次の行へ送る。
 * 半角の語（「2026/12」など）の途中では切らず、語の前の空白で折り返す。
 */
function mtgWrap_(text, widthPt, size, max) {
  const room = (widthPt - MTG_SLIDE.padX) / (size || MTG_SLIDE.bodySize);
  const limit = room * 0.92;
  const width = mtgCharWidth_;
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
 * スライドの並び（mtgSlides_）を、Google スライド（16:9）と PDF にしてフォルダに保存する。
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
