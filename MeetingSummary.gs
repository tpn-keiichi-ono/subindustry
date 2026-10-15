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
 *   使う言葉は docs/DESIGN.md（会議サマリの数え方）で決め、資料の中で同じ意味で使う。
 *
 * 資料の構成（16:9。緑を基調にしたトンマナ：黒の見出し・リード文・緑の見出しの表・区分のラベル・区切り線と出典・ページ番号。
 *   表紙と扉は、いただいた表紙のひな形（Google スライド。setMeetingTemplate() で登録）をそのまま使う
 *   表紙：会議の名前と「事前資料」・サブインダストリー・会議日・データの時点・期間
 *   1 サマリー：期待値と進行中案件の数・前回以降の動き（案件ごとに1回だけ数える区分の件数と案件）・現在の状態（受注月超過・停滞）
 *   2 パイプラインの変化：前回からの期待値の増減（新規・増額・復帰・減額・受注・失注・保留・削除）・ステータス別・受注月（四半期）別の期待値
 *   3 前回から変化のある案件：ステータス・期待値・受注月・活動状況が変わった案件（区分の順に上位）
 *   4 期待値上位案件の状況・課題：進行中の案件の期待値の大きい順（ステータスの滞留日数・受注月・活動状況・課題・該当事項）
 *   扉（Appendix。Appendix の見出しとページ）
 *   Appendix（全件。入りきらなければ続きのページ）
 *   ① 前回から変化のある案件　② 停滞案件　③ 入力内容のチェック（責任者別・一覧）
 *   左上のラベルは置かない（どのサブインダストリー・会議日の資料かは、表紙・ファイル名・下の出典の行に書く。試作は表紙と出典の行に【試作】）
 *
 * ことばの決まり（どのページでも同じ数え方にする）
 * - ステータスの分類：進行中・受注・失注・保留（MTG_OPTIONS のことば・statusKinds で決める）。期待値は進行中の案件だけを数える。
 *   前進・後退は、案件ステータスのプルダウンの並び順で決める。
 * - 前回からの変化は、案件ごとに1つの区分にする（新規 > 受注 > 失注 > 保留 > 復帰 > 前進 > 後退 > 延期 > 前倒し > 増額 > 減額 > 変更 > 活動更新 > その他更新）。
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
 *   表紙・扉のひな形を登録してあれば（setMeetingTemplate()）、ひな形をコピーした資料に作り、ひな形の表紙の決まった文字を置き換える。
 */

const MTG_OPTIONS = {
  meetingName: 'RX Business MTG',    // 表紙に出す会議の名前
  fileName: '会議サマリ',             // ファイル名の先頭（会議サマリ_スーパー・GMS_2026-10-23.pdf。日付は会議日）
  folderId: '',                      // 保存先のフォルダの ID（空なら、このスプレッドシートと同じフォルダの中に folderName を作る）
  folderName: '会議サマリ',
  folderKey: 'MTG_FOLDER_ID',         // 作ったフォルダの ID（ドキュメントのプロパティ）
  // 表紙・扉のひな形（Google スライド。いただいた .pptx を Google スライドにしたもの）。資料はひな形のコピーに作る（ページの大きさもひな形のまま）。
  // ID は setMeetingTemplate() で登録する（ドキュメントのプロパティ templateKey）。templateId を書けば、そちらを使う。どちらも無ければ、ひな形なしで作る
  templateId: '',
  templateKey: 'MTG_TEMPLATE_ID',
  // ひな形の1枚目（表紙）の、置き換える文字（長いものから置き換える）。扉は表紙を複製して同じように置き換える
  templateTexts: {title: 'XXXXXX', subtitle: 'XXXXX', date: '2026年 4月 17日', name: 'XXXX'},
  meetingsKeyPrefix: 'MTG_MEETINGS_', // シートごとの会議日と資料を作った日時（ドキュメントのプロパティ。後ろにシート名）
  draftKeyPrefix: 'MTG_DRAFT_',       // シートごとの、会議日ごとの最初の試作の入力内容のチェックの結果（解消した件数を出す用）
  lastKeyPrefix: 'MTG_LAST_',         // 前の版の「前回の作成日時」（会議日の記録が無いときだけ使う）
  keepMeetings: 24,                  // 会議日の記録をいくつ残すか
  keepSlides: true,                  // Google スライドもフォルダに残す（推進リーダーが書き足せる。false なら PDF だけ）
  // Google チャットへの通知（webhook）。setMeetingChat() で webhook の URL と資料の格納先の URL を登録し、
  // notifyMeetingSummaries() を実行すると、その場で格納先の URL をチャットに送る（自動では送らない）
  chat: {
    webhookKey: 'MTG_CHAT_WEBHOOK',  // スクリプトのプロパティ：webhook の URL（鍵を含むので、コードやシートに書かない）
    urlKey: 'MTG_CHAT_URL',          // スクリプトのプロパティ：送る資料の格納先の URL（ドライブのフォルダ・ファイルなど）
    legacyHandler: 'mtgNotifyScheduled' // 前の版の自動通知（2日前）のトリガーの関数。残っていれば setMeetingChat()・notifyMeetingSummaries() で外す
  },
  firstDays: 14,                     // 初回（前回の会議の記録が無い）は、何日前からの変化を出すか
  lookbackDays: 180,                 // ステータスの滞留・受注月の延期の回数を数えるために、何日前までの記録を読むか
  stallDays: 45,                     // 停滞：ステータスがこの日数以上同じ進行中の案件
  noUpdateDays: 30,                  // 入力内容のチェック「更新なし」：最終更新日時がこの日数より前の進行中の案件
  overdueDays: 90,                   // 入力内容のチェック「予定日超過」：予定日を過ぎてこの日数以内のもの（古いものは出さない）
  fiscalStartMonth: 4,               // 期の始まりの月（受注月別の期待値の四半期）
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
  // 色（緑を基調にしたトンマナ）。表の見出し・帯は green、表の1列目・合計は greenSoft、大きな数字・期待値の棒は greenDark
  colors: {
    green: '#70AD47', greenDark: '#38761D', greenDeep: '#548235', greenSoft: '#E2EFDA', greenText: '#274E13',
    up: '#8FD36B', down: '#A6A6A6', dark: '#595959',
    ink: '#1A1A1A', body: '#1A1A1A', sub: '#404040', muted: '#595959', faint: '#7F7F7F',
    line: '#D0D0D0', frame: '#BFBFBF', rule: '#D9D9D9', white: '#FFFFFF', graySoft: '#EDEDED',
    cover: '#3A7D22'                 // ひな形が無いときの表紙・扉の背景
  }
};

/*
 * 前回からの変化の区分（上ほど優先。案件ごとに1つ）。意味は docs/DESIGN.md（会議サマリの数え方）。
 * tone は表のラベルの色（up 濃い緑：前向き・down 濃いグレー：後ろ向き・note 薄いグレー：そのほか。mtgPills_）
 */
const MTG_CATEGORIES = [
  {key: 'new', label: '新規', tone: 'up'},
  {key: 'won', label: '受注', tone: 'up'},
  {key: 'lost', label: '失注', tone: 'down'},
  {key: 'hold', label: '保留', tone: 'down'},
  {key: 'revived', label: '復帰', tone: 'up'},
  {key: 'forward', label: '前進', tone: 'up'},
  {key: 'backward', label: '後退', tone: 'down'},
  {key: 'slipped', label: '延期', tone: 'note'},
  {key: 'pulled', label: '前倒し', tone: 'up'},
  {key: 'up', label: '増額', tone: 'up'},
  {key: 'down', label: '減額', tone: 'down'},
  {key: 'status', label: '変更', tone: 'note'},
  {key: 'removed', label: '削除', tone: 'down'},
  {key: 'activity', label: '活動更新', tone: 'note'},
  {key: 'updated', label: 'その他更新', tone: 'note'}
];

/*
 * スライドの見出し（タイトル）。どのサブインダストリーでも同じにする（状態はリード文に書く）。
 * 表紙のあとに本編：サマリー → パイプラインの変化 → 前回から変化のある案件 → 期待値上位案件の状況・課題。扉（Appendix）のあとに、全件の一覧と言葉の定義
 */
const MTG_TITLES = {
  cover: '表紙',
  summary: 'サマリー',
  pipeline: 'パイプラインの変化',
  progress: '前回から変化のある案件',
  focus: '期待値上位案件の状況・課題',
  appendix: 'Appendix',
  changes: 'Appendix ①　前回から変化のある案件（一覧）',
  stalled: 'Appendix ②　停滞案件（一覧）',
  owners: 'Appendix ③　入力内容の確認事項（責任者別）',
  alerts: 'Appendix ③　入力内容の確認事項（一覧）'
};

/*
 * 本編の「前回から変化のある案件」に出す区分の順（その他更新は出さない）。サマリーの「前回以降の動き」の行の順と同じ
 * （受注／失注・保留・後退・削除／新規・前進・復帰／活動更新／受注月・期待値などの変更）
 */
const MTG_MOVE_ORDER = ['won', 'lost', 'hold', 'backward', 'removed', 'new', 'forward', 'revived', 'activity', 'slipped', 'pulled', 'up', 'down',
  'status'];

/* 期待値の増減の内訳（増える → 減るの順） */
const MTG_BRIDGE = [
  {key: 'new', label: '新規'}, {key: 'up', label: '期待値増'}, {key: 'revived', label: '復帰'},
  {key: 'down', label: '期待値減'}, {key: 'won', label: '受注'}, {key: 'lost', label: '失注'}, {key: 'hold', label: '保留'},
  {key: 'removed', label: '削除'}
];

/* 入力内容のチェックの種類（この順に出す。意味は docs/DESIGN.md の「入力内容のチェック」） */
const MTG_ALERT_TYPES = ['予定日超過', '受注月超過', '未入力', '資料なし', '更新なし', '金額の不整合', '月の不整合', '選択肢外のステータス', 'デリバリー担当なし'];
/* 資料に出す種類の名前（資料では「入力内容の確認事項」。種類の名前は試作の結果の記録にも使うので、出す名前だけを変える） */
const MTG_ALERT_LABELS = {'予定日超過': '予定日超過', '受注月超過': '受注月超過', '未入力': '必須項目未入力', '資料なし': '商談記録未登録',
  '更新なし': '長期未更新', '金額の不整合': '金額不整合', '月の不整合': '売上開始月不整合', '選択肢外のステータス': '選択肢外ステータス',
  'デリバリー担当なし': 'デリバリー担当未設定'};

/*
 * スライドの形（pt。16:9 = 720 × 405）。上から 見出し（titleY）・リード文（leadY）・中身（top〜bottom）・出典の行とページ番号（footY）。
 * 表紙と扉は、見出し・リード文の代わりに大きな文字を置く（表紙には出典の行・ページ番号も置かない）。
 * 文字の幅と行の高さは見積もり（mtgWrap_・mtgLineHeight_）。本物より大きめに見積もって、はみ出さないようにする。
 */
const MTG_SLIDE = {
  width: 720,
  height: 405,
  font: 'BIZ UDPGothic',  // BIZ UDPゴシック（Google スライド・Google Fonts での名前）
  x: 22,                 // 左右の余白
  w: 676,                // 中身の幅
  titleY: 12,            // 見出しの上端
  leadY: 47,             // リード文の上端（見出しとの間を 10pt ほどあける。補足はその下の行）
  top: 98,               // 中身の上端
  summaryTop: 70,        // サマリーの中身の上端（リード文が1行のとき。2行以上ならリード文の下から）
  bottom: 379,           // 中身の下端（下に区切り線・出典の行・ページ番号）
  ruleY: 383,            // いちばん下の区切り線
  footY: 386,            // 出典の行の上端
  titleSize: 21,
  leadSize: 10.5,        // リード文（1つ目の文も補足も同じ大きさ・同じ色）
  leadLine: 1.35,        // リード文の行の送り = 文字の大きさ × leadLine
  bodySize: 9,
  appendixSize: 8.5,
  lineFactor: 1.5,       // 表の行の高さ = 文字の大きさ × lineFactor（Google スライドで行が伸びないよう、多めに見積もる）
  padX: 14.4,            // 箱・表のセルの左右の余白（合わせて）
  padY: 7.2,             // 上下の余白（合わせて）
  headRow: 20,           // 表の見出しの行の高さ（1行）
  // ラベル（区分・当てはまる状況）：高さ・文字の大きさ・縦に並べるときの間・セルの上下の余白。幅は列ごとに同じ（mtgPillWidth_）
  pill: {h: 12, size: 8, gap: 2, margin: 4}
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

/**
 * 記録した会議日を取り消す（会議日を打ち間違えたとき・本番を作り直す前に戻すとき。エディタから実行する）。
 * シートと会議日を聞き、その会議日の記録を消す。基準の記録と、保存した資料のファイルは消さない
 */
function removeMeetingRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const names = mtgAskSheets_(ss, ui, '会議サマリの会議日の記録を取り消します（資料のファイルは消しません）。');
  if (!names) return;
  const props = PropertiesService.getDocumentProperties();
  const recorded = names.map(name => '・' + name + '：' +
    (mtgMeetings_(props, name).filter(m => !m.baseline).map(m => mtgShortDate_(m.date)).join('、') || '記録なし'));
  const response = ui.prompt(CRED_OPTIONS.menuTitle, '取り消す会議日を入力してください（例：10/23）。\n\n記録してある会議日\n' + recorded.join('\n'),
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const text = String(response.getResponseText() || '').trim();
  const date = text && mtgParseDate_(text, new Date());
  if (!date) throw new Error('「' + text + '」を会議日として読めません。10/23 か 2026/10/23 のように入力してください。');
  const hit = names.filter(name => mtgMeetings_(props, name).some(m => !m.baseline && m.date === date));
  if (!hit.length) {
    ui.alert(CRED_OPTIONS.menuTitle, '「' + names.join('」「') + '」に、会議日 ' + date + ' の記録はありません。', ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert(CRED_OPTIONS.menuTitle, '「' + hit.join('」「') + '」の会議日（' + date + '）の記録を取り消します。\n' +
    '次に作る資料は、その前の会議（または基準）からの変化を出します。よろしいですか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  hit.forEach(name => props.setProperty(MTG_OPTIONS.meetingsKeyPrefix + name,
    JSON.stringify(mtgMeetings_(props, name).filter(m => m.baseline || m.date !== date))));
  const message = '「' + hit.join('」「') + '」の会議日（' + date + '）の記録を取り消しました。';
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/**
 * 本番：会議日を記録する前に、会議日（曜日）・シートごとの対象期間・同じ会議日の資料を作成済みかを出して確かめる。
 * 記録した会議日は次の会議の資料の起点になるので、配布したあとは作り直さないよう知らせる。はいのときだけ true
 */
function mtgConfirmFinal_(ui, names, meeting) {
  const props = PropertiesService.getDocumentProperties();
  const tz = Session.getScriptTimeZone();
  const fmt = at => Utilities.formatDate(new Date(at), tz, 'M/d HH:mm');
  const p = meeting.split('/').map(Number);
  const week = '日月火水木金土'.charAt(new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay());
  const lines = names.map(name => {
    const list = mtgMeetings_(props, name);
    const held = list.filter(m => !m.baseline && m.date < meeting).pop();
    const base = list.filter(m => m.baseline && m.date <= meeting).pop();
    const same = list.find(m => !m.baseline && m.date === meeting);
    const legacy = Number(props.getProperty(MTG_OPTIONS.lastKeyPrefix + name)) || 0;   // 前の版の記録（mtgCreate_ と同じ順で選ぶ）
    const from = base && (!held || base.cutoff > held.cutoff) ? '基準（' + fmt(base.cutoff) + '）以降'
      : held ? '前回会議（' + mtgShortDate_(held.date) + '）の資料作成時（' + fmt(held.cutoff) + '）以降'
        : legacy ? '前回の資料作成時（' + fmt(legacy) + '）以降'
          : '直近' + MTG_OPTIONS.firstDays + '日間（前回の会議の記録なし）';
    return '・' + name + '：' + from +
      (same ? '\n　※ この会議日の資料は ' + fmt(same.cutoff) + ' に作成済み。作り直すと、次の会議の起点が今になります。' : '');
  });
  const answer = ui.alert(CRED_OPTIONS.menuTitle,
    '本番の資料を作り、会議日を記録します。\n\n会議日：' + p[0] + '/' + p[1] + '/' + p[2] + '（' + week + '）\n対象期間\n' + lines.join('\n') +
    '\n\n記録した会議日は、次の会議の資料の起点になります。配布したあとは作り直さないでください。\n作成してよろしいですか？', ui.ButtonSet.YES_NO);
  return answer === ui.Button.YES;
}

function mtgRun_(final) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const names = mtgAskSheets_(ss, ui, final ? '会議サマリを作ります。' : '会議サマリを試しに作ります（会議日は記録しません）。');
  if (!names) return;
  const now = new Date();
  const meeting = mtgAskMeeting_(ui, names, now);
  if (!meeting) return;
  if (final && !mtgConfirmFinal_(ui, names, meeting)) return;
  const results = mtgCreate_(ss, names, now, final, meeting);
  const lines = results.map(r => '「' + r.summary.name + '」：' + r.file.getName() + '（' + mtgCountsText_(r.summary) + '）\n　' + r.file.getUrl());
  lines.push('', '保存先：' + results[0].folder.getUrl());
  if (MTG_OPTIONS.keepSlides) lines.push('同じフォルダに、Google スライド（同じ名前）も保存しました。');
  lines.push(final
    ? '会議日（' + meeting + '）を記録しました。次の会議の資料は、今以降の変化を出します。配るときは、PDF を参加者に共有してください（保存先のフォルダは共有しない）。'
    : '試作のため、会議日は記録していません。');
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/* ---------------- Google チャットへの通知（webhook） ---------------- */

/**
 * Google チャットへの通知を設定する（エディタから実行する。管理者が1回）。webhook の URL と、送る資料の格納先の URL を聞き、スクリプトのプロパティに残す。
 * webhook の URL を空欄で OK にすると、設定を外す。前の版の自動通知のトリガーが残っていれば外す
 */
function setMeetingChat() {
  const ui = SpreadsheetApp.getUi();
  const o = MTG_OPTIONS.chat;
  const props = PropertiesService.getScriptProperties();
  const current = mtgChatSettings_();
  const response = ui.prompt(CRED_OPTIONS.menuTitle,
    'Google チャットの webhook の URL を貼り付けてください（スペースの「アプリと統合」→「Webhook」で作ったもの）。\n' +
    '空欄で OK を押すと、設定を外します。\n\n今の設定：' + (current ? mtgChatSpace_(current.webhook) + ' に送る' : '未設定'),
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const webhook = String(response.getResponseText() || '').trim();
  if (!webhook) {
    props.deleteProperty(o.webhookKey);
    props.deleteProperty(o.urlKey);
    mtgRemoveLegacyTrigger_();
    const message = 'Google チャットへの通知の設定を外しました。';
    console.log(message);
    ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
    return;
  }
  if (!/^https:\/\/chat\.googleapis\.com\/v1\/spaces\/[^/?]+\/messages\?\S*key=\S+&token=\S+$/.test(webhook)) {
    throw new Error('Google チャットの webhook の URL ではありません。https://chat.googleapis.com/v1/spaces/…/messages?key=…&token=… の形の URL を貼り付けてください。');
  }
  const answer = ui.prompt(CRED_OPTIONS.menuTitle,
    '送る資料の格納先の URL を貼り付けてください（ドライブのフォルダ・ファイルなど。参加者が開けるように共有しておく）。' +
    (current && current.url ? '\n空欄なら今の URL のまま：' + current.url : ''), ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const url = String(answer.getResponseText() || '').trim() || (current && current.url) || '';
  if (!/^https:\/\/\S+$/.test(url)) throw new Error('資料の格納先の URL を読めません。https:// で始まる URL を貼り付けてください。');
  props.setProperty(o.webhookKey, webhook);
  props.setProperty(o.urlKey, url);
  const removed = mtgRemoveLegacyTrigger_();
  const lines = ['Google チャットへの通知を設定しました。', '・送り先：' + mtgChatSpace_(webhook), '・送る URL：' + url,
    '・notifyMeetingSummaries() を実行すると、この URL をすぐに送ります。']
    .concat(removed ? ['・前の版の自動通知（2日前）のトリガーを外しました。'] : []);
  const test = ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n') + '\n\n確かめのメッセージを、いまチャットに送りますか？', ui.ButtonSet.YES_NO);
  if (test === ui.Button.YES) {
    mtgPostChat_(webhook, '会議サマリの通知の設定を確かめるためのメッセージです。' + MTG_OPTIONS.meetingName + ' の事前資料の URL をここに送ります。');
    lines.push('確かめのメッセージを送りました。チャットに届いたか確かめてください。');
  }
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/** 登録した資料の格納先の URL を、すぐに Google チャットに送る（エディタから実行する。「はい」のときだけ） */
function notifyMeetingSummaries() {
  const chat = mtgChatSettings_();
  if (!chat || !chat.url) throw new Error('Google チャットへの通知が設定されていません。先に setMeetingChat() で webhook の URL と資料の格納先の URL を登録してください。');
  mtgRemoveLegacyTrigger_();
  const ui = SpreadsheetApp.getUi();
  const text = mtgChatText_(chat.url);
  const answer = ui.alert(CRED_OPTIONS.menuTitle,
    'Google チャット（' + mtgChatSpace_(chat.webhook) + '）に、次のメッセージをすぐに送ります。よろしいですか？\n\n' + text, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  mtgPostChat_(chat.webhook, text);
  const message = 'Google チャットに送りました：' + chat.url;
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** チャットに送る文（会議の名前と、資料の格納先の URL） */
function mtgChatText_(url) {
  return '*' + MTG_OPTIONS.meetingName + '　事前資料*\n会議の前に、資料を通読してください。\n' + url;
}

/** webhook にメッセージを送る（混んでいる・一時的な失敗は1回だけやり直す）。失敗したら止める（URL は鍵を含むので出さない） */
function mtgPostChat_(webhook, text) {
  const params = {method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify({text}), muteHttpExceptions: true};
  let res = UrlFetchApp.fetch(webhook, params);
  if (res.getResponseCode() === 429 || res.getResponseCode() >= 500) {
    Utilities.sleep(2000);
    res = UrlFetchApp.fetch(webhook, params);
  }
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('Google チャットに送れませんでした（HTTP ' + code + '）：' + String(res.getContentText() || '').slice(0, 200) +
      '\nwebhook の URL が消されていないか、setMeetingChat() で登録し直してください。');
  }
}

/** 通知の設定 {webhook, url（資料の格納先。無ければ null）}（スクリプトのプロパティ）。webhook が無ければ null */
function mtgChatSettings_() {
  const props = PropertiesService.getScriptProperties();
  const webhook = props.getProperty(MTG_OPTIONS.chat.webhookKey);
  return webhook ? {webhook, url: props.getProperty(MTG_OPTIONS.chat.urlKey) || null} : null;
}

/** 送り先の見せ方（「スペース AAAA…」。鍵・トークンは出さない） */
function mtgChatSpace_(webhook) {
  const m = /\/spaces\/([^/?]+)/.exec(webhook);
  return 'スペース ' + (m ? m[1] : '（不明）');
}

/**
 * 前の版で、会議日の2日前に自動で送っていた時間主導のトリガーの関数。今は自動では送らない（notifyMeetingSummaries() ですぐ送る）。
 * トリガーが残っていても「関数が見つからない」エラーにならないよう名前だけを残し、動いたらそのトリガーを外す（何も送らない）。
 */
function mtgNotifyScheduled() {
  try {
    mtgRemoveLegacyTrigger_();
  } catch (error) {
    console.warn('前の版の自動通知のトリガーを外せませんでした：' + (error && error.message));
  }
}

/**
 * 前の版の自動通知のトリガー（mtgNotifyScheduled）を外す。外した数を返す。
 * トリガーは持ち主にしか見えないので、前の版で setMeetingChat() を実行した人が実行したとき（またはトリガーが動いたとき）に外れる
 */
function mtgRemoveLegacyTrigger_() {
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === MTG_OPTIONS.chat.legacyHandler);
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  return triggers.length;
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

/** 完成した結果の件数の知らせ（「変化のある案件 12・停滞 5・確認事項 7」） */
function mtgCountsText_(summary) {
  return '変化のある案件 ' + summary.kpi.changed + '・停滞 ' + summary.stalled.length + '・確認事項 ' + summary.alerts.length;
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
 * 戻り値 {name, meeting, previous, legacy, since, first, final, now, windowStart, kpi, recordColumn: 商談記録の列があるか,
 *         stallPending: 停滞をまだ数えられないとき {at: 記録の始まりか基準の時刻, baseline}, bridge, stages, quarters,
 *         changes, stalled, focus: 期待値の大きい進行中の案件, overdue: 受注月を過ぎた進行中の案件, alerts: 入力内容のチェック,
 *         resolved, unmatched}
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
    // touched：前回以降に何かしらの更新（その他の項目のみの更新も含む）がある
    items.push(Object.assign({}, deal, {days, atLeast, slips, overdue, stalled, touched: changed.has(key)}));
  });
  // 停滞をまだ数えられない：記録の始まり・基準から stallDays 日たっていない（その前からの日数は分からないため）
  const stallFrom = [tracked, baseline].filter(v => v != null);
  const from = stallFrom.length ? Math.max.apply(null, stallFrom) : null;
  const stallPending = from != null && now - from < o.stallDays * 86400000
    ? {at: from, baseline: baseline != null && from === baseline} : null;
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
  const owners = new Set(alerts.map(a => a.owner || '（責任者未設定）'));
  const alertDeals = new Set(alerts.map(a => a.dealKey || a.key));
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
      stalledQuiet: stalled.filter(d => !d.touched).length,
      overdue: items.filter(d => d.overdue).length,
      alerts: alerts.length,
      alertDeals: alertDeals.size,
      owners: owners.size
    },
    recordColumn: !!table.has.record,
    stallPending,
    bridge,
    stages: mtgStages_(deals, table.statusOrder),
    quarters: mtgQuarters_(active, context.now),
    changes,
    stalled,
    focus,
    overdue: items.filter(d => d.overdue).sort((a, b) => (b.amount || 0) - (a.amount || 0)),
    alerts,
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
 * {start, end, startCount, endCount, amountStart, amountEnd（進行中の想定売上の合計。前回・今回）, steps: [{key, label, value, count}]}
 */
function mtgBridge_(deals, byKey, changed, changes) {
  const steps = {};
  MTG_BRIDGE.forEach(b => { steps[b.key] = {key: b.key, label: b.label, value: 0, count: 0}; });
  const out = {start: 0, end: 0, startCount: 0, endCount: 0, amountStart: 0, amountEnd: 0};
  const add = (item, was, wStart, now, wEnd, kindNow, aStart, aEnd) => {
    if (was) { out.start += wStart; out.startCount++; out.amountStart += aStart || 0; }
    if (now) { out.end += wEnd; out.endCount++; out.amountEnd += aEnd || 0; }
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
    const aStart = item ? (item.base && item.base.amount) || 0 : d.amount || 0;
    add(item, was, wStart, d.kind === 'active', d.weighted || 0, d.kind, aStart, d.amount || 0);
  });
  changes.filter(c => c.removed).forEach(c => {
    add(c, mtgStatusKind_(c.base.status) === 'active', c.base.weighted || 0, false, 0, null, c.base.amount || 0, 0);
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
        text: (c.values.kind || '予定') + '「' + c.values.name + '」：予定日（' + mtgShortDate_(plan) + '）から ' + days + '日経過、実施日未入力'});
    });
  });
  overdue.sort((a, b) => b.days - a.days);
  active.forEach(d => {
    const slips = info.slipsOf.get(d.key) || 0;
    if (d.orderMonth && mtgMonthIndex_(d.orderMonth) < month) {
      push('受注月超過', d, '受注月（' + label(d.orderMonth) + '）経過、ステータス「' + (d.status || '未入力') + '」' + (slips ? '（延期 ' + slips + '回）' : ''));
    }
    if (d.missing.length) push('未入力', d, '未入力：' + d.missing.map(h => h.replace(/\s+/g, '')).join('・'));
    if (table.has.record && !d.record) push('資料なし', d, '「' + o.columns.record + '」未入力');
    if (d.updatedAt != null) {
      const days = Math.floor((now - d.updatedAt) / 86400000);
      if (days >= o.noUpdateDays) push('更新なし', d, '最終更新から ' + days + '日経過');
    } else if (info.tracked != null && now - info.tracked >= o.noUpdateDays * 86400000) {
      push('更新なし', d, '記録開始以降の更新なし（' + Math.floor((now - info.tracked) / 86400000) + '日以上）');
    }
    if (d.amount != null && d.weighted != null && d.weighted > d.amount) {
      push('金額の不整合', d, '期待値（' + mtgAmount_(d.weighted) + '）が想定売上（' + mtgAmount_(d.amount) + '）を上回る');
    }
  });
  deals.forEach(d => {
    if ((d.kind === 'active' || d.kind === 'won') && d.orderMonth && d.startMonth &&
        mtgMonthIndex_(d.startMonth) < mtgMonthIndex_(d.orderMonth)) {
      push('月の不整合', d, '売上開始月（' + label(d.startMonth) + '）が受注月（' + label(d.orderMonth) + '）より前');
    }
    if (table.statusOrder.length && d.status && info.rank(d.status) < 0) {
      push('選択肢外のステータス', d, '「' + d.status + '」は選択肢外の値（' + {active: '進行中', won: '受注', lost: '失注', hold: '保留'}[d.kind] + 'として集計）');
    }
  });
  if (table.has.delivery) {
    info.changes.filter(c => c.category === 'won' && c.deal && !c.deal.delivery)
      .forEach(c => push('デリバリー担当なし', c.deal, '受注済み、「' + MTG_OPTIONS.columns.delivery + '」未入力'));
  }
  const byAmount = (a, b) => (b.amount || 0) - (a.amount || 0);
  return overdue.concat(MTG_ALERT_TYPES.slice(1).reduce((list, type) => list.concat(out.filter(a => a.type === type).sort(byAmount)), []));
}

/** 今期の始まりの月の番号（年 × 12 + 月） */
function mtgFiscalStart_(now) {
  const m = mtgMonth_(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM'));
  const fs = MTG_OPTIONS.fiscalStartMonth;
  return (m.m >= fs ? m.y : m.y - 1) * 12 + fs;
}

/**
 * 進行中の案件の受注月ごとの期待値 [{label, value, count}]：受注月超過（受注月が今月より前）・今期の今の四半期から4Q まで・来期以降・受注月未入力。
 * 受注月を過ぎた案件は四半期に入れない（「受注月が今期の期待値」に過ぎた案件を混ぜないため。過ぎた四半期の行は出さない）
 */
function mtgQuarters_(active, now) {
  const start = mtgFiscalStart_(now);
  const month = mtgMonthIndex_(mtgMonth_(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM')));
  const first = Math.floor((month - start) / 3);
  const over = {label: '受注月超過', value: 0, count: 0, optional: true};
  const quarters = [];
  for (let q = first; q < 4; q++) {
    const a = mtgMonthOf_(start + q * 3);
    const b = mtgMonthOf_(start + q * 3 + 2);
    quarters.push({label: (q + 1) + 'Q（' + a.m + '〜' + b.m + '月）', value: 0, count: 0});
  }
  const later = {label: '来期以降', value: 0, count: 0, optional: true};
  const none = {label: '受注月未入力', value: 0, count: 0, optional: true};
  active.forEach(d => {
    const i = d.orderMonth ? mtgMonthIndex_(d.orderMonth) : null;
    const bucket = i == null ? none : i < month ? over : i >= start + 12 ? later : quarters[Math.floor((i - start) / 3) - first];
    bucket.value += d.weighted || 0;
    bucket.count++;
  });
  return [over].concat(quarters, [later, none]).filter(b => !b.optional || b.count).map(b => ({label: b.label, value: b.value, count: b.count}));
}

/** ステータスごとの件数・金額 [{status, kind, known, count, amount, weighted}]（進行中 → 保留 → 受注 → 失注。同じ分類の中はプルダウンの順） */
function mtgStages_(deals, statusOrder) {
  const map = new Map();
  const order = statusOrder.slice();
  deals.forEach(d => {
    const label = d.status || '（未入力）';
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
 * 表紙、本編（サマリー → パイプラインの変化 → 前回から変化のある案件 → 期待値上位案件の状況・課題）、扉（Appendix）、Appendix ①〜③。
 * ページ番号は表紙を1枚目として数える（表紙・扉には出さない。PDF のページと同じ番号になる）。
 * 見出しは MTG_TITLES（どのシートでも同じ）。書くのはシートと変更の記録から分かる事実だけ（論点・決めてほしいことなどの判断は、会議でリーダーが話す）。
 * 戻り値 [{title, elements}]（elements は mtgSlideRequests_ の部品。タイトル・リード文・出典・ページ番号も含む）
 */
function mtgSlides_(summary) {
  const cover = mtgCoverSlide_(summary);
  const main = [mtgSummarySlide_(summary), mtgBridgeSlide_(summary), mtgMovesSlide_(summary), mtgFocusSlide_(summary)];
  const appendix = [].concat(mtgAppendixChanges_(summary), mtgAppendixStalled_(summary), mtgAppendixOwners_(summary),
    mtgAppendixAlerts_(summary));
  const divider = appendix.length ? [mtgDividerSlide_(appendix, 1 + main.length + 1 + 1)] : [];
  return mtgNumberSlides_(mtgSheetFrame_(summary), [cover].concat(main, divider, appendix));
}

/**
 * 表紙（ひな形の1枚目）：ひな形の決まった文字を置き換える（template。見出しにサブインダストリー、その下に会議の名前と「事前資料」
 * （試作は【試作】）、日付に会議日、名前の欄は空）。その下にデータの時点・期間を小さく足す（elements）
 */
function mtgCoverSlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const p = s.meeting.split('/').map(Number);
  const week = '日月火水木金土'.charAt(new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay());
  const line = (y, text) => mtgBox_(S.x, y, S.w, 16, mtgWrap_(text, S.w, 9, 1)[0], {size: 9, color: c.white, align: 'CENTER'});
  return {
    title: MTG_TITLES.cover,
    cover: true,
    template: {title: s.name, subtitle: o.meetingName + '　事前資料' + (s.final ? '' : '【試作】'),
      date: p[0] + '年' + p[1] + '月' + p[2] + '日（' + week + '）', name: ''},
    elements: [
      line(350, '出典：「' + s.name + '」シートおよび変更履歴（' + Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/M/d HH:mm') + ' 時点）'),
      line(365, '対象期間：' + mtgPeriodText_(s))
    ]
  };
}

/**
 * 扉（Appendix。ひな形の1枚目を複製）：見出しを「Appendix」にし（ほかの決まった文字は空）、その下に Appendix の見出し・始まりのページの一覧
 * （白い字と白い線）。appendix は Appendix のスライド、firstPage はその1枚目のページ番号
 */
function mtgDividerSlide_(appendix, firstPage) {
  const c = MTG_OPTIONS.colors;
  const items = [];
  appendix.forEach((slide, i) => {
    const m = /^(\S+ \S+)　(.*)$/.exec(slide.title) || [slide.title, slide.title, ''];
    if (items.some(it => it.label === m[1])) return;
    items.push({label: m[1], title: m[2].replace(/（[^（）]*）$/, ''), page: firstPage + i});
  });
  const el = [];
  const x = 180;
  const rowH = 26;
  items.forEach((it, i) => {
    const y = 228 + i * rowH;
    el.push(mtgBox_(x, y, 104, 22, it.label, {size: 12, bold: true, color: c.white, role: 'toc'}));
    el.push(mtgBox_(x + 98, y, 222, 22, mtgWrap_(it.title, 222, 12, 1)[0], {size: 12, color: c.white, role: 'toc'}));
    el.push(mtgBox_(x + 320, y, 40, 22, String(it.page), {size: 12, color: c.white, align: 'END', role: 'toc'}));
    el.push({type: 'box', x, y: y + 23, w: 360, h: 0.75, fill: c.white, alpha: 0.5, shape: 'RECTANGLE'});
  });
  return {title: MTG_TITLES.appendix, divider: true, template: {title: MTG_TITLES.appendix, subtitle: '', date: '', name: ''}, elements: el};
}

/** スライドに、どのスライドにも付ける部品（mtgFrame_）とページ番号を付ける（表紙・扉の印と、ひな形の置き換える文字は残す） */
function mtgNumberSlides_(frame, slides) {
  return slides.map((slide, i) => {
    const out = {title: slide.title, elements: mtgFrame_(frame, slide, i + 1).concat(slide.elements)};
    ['cover', 'divider', 'template'].forEach(key => { if (slide[key]) out[key] = slide[key]; });
    return out;
  });
}

/** シートの資料のいちばん下の行（どのサブインダストリー・会議日の資料か。試作は先頭に【試作】）{source}。データの時点・期間は表紙に書く */
function mtgSheetFrame_(s) {
  return {source: (s.final ? '' : '【試作】') + s.name + '（' + mtgShortDate_(s.meeting) + ' 会議資料）　金額の単位：百万円'};
}

/**
 * どのスライドにも付ける部品：タイトル・リード文・いちばん下の区切り線と出典の行・ページ番号（表紙・扉には付けない）。
 * frame = {source}。slide.note があれば出典の行の後ろに付ける
 */
function mtgFrame_(frame, slide, page) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  if (slide.cover || slide.divider) return [];
  const title = mtgTitleFit_(slide.title);
  const el = [mtgBox_(S.x, S.titleY, S.w, 30, title.text, {size: title.size, bold: true, color: c.ink, role: 'title'})];
  if (slide.lead) el.push.apply(el, mtgLead_(slide.lead));
  el.push({type: 'box', x: S.x, y: S.ruleY, w: S.w, h: 0.75, fill: c.rule, shape: 'RECTANGLE'});
  el.push(mtgBox_(S.x, S.footY, S.w - 60, 16, mtgWrap_(frame.source + (slide.note ? '　' + slide.note : ''), S.w - 60, 9, 1)[0],
    {size: 9, color: c.faint, role: 'footer'}));
  el.push(mtgBox_(S.x + S.w - 60, S.footY - 1, 60, 18, String(page), {size: 10.5, color: c.ink, align: 'END', role: 'page'}));
  return el;
}

/**
 * リード文（タイトルの下）。どのページも同じ大きさ（MTG_SLIDE.leadSize）・同じ色（黒）。改行の前（1つ目の文）と改行のあと（補足）は行を分ける。
 * 1つ目の文は2行まで、補足は残りの行（合わせて3行まで）。入らなければ「…」で切る
 */
function mtgLead_(text) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const parts = String(text).split('\n');
  const rest = parts.slice(1).join('');
  const size = S.leadSize;
  const step = lines => lines * size * S.leadLine;
  const first = mtgWrap_(parts[0], S.w, size, 2);
  const el = [mtgBox_(S.x, S.leadY, S.w, step(first.length) + S.padY, first.join('\n'), {size, color: c.ink, valign: 'TOP', role: 'lead'})];
  if (rest) {
    const second = mtgWrap_(rest, S.w, size, 3 - first.length);
    el.push(mtgBox_(S.x, S.leadY + step(first.length), S.w, step(second.length) + S.padY, second.join('\n'),
      {size, color: c.ink, valign: 'TOP', role: 'lead'}));
  }
  return el;
}

/** 「前回以降」「基準以降」「直近14日間」（数字の見出し・リード文の期間） */
function mtgPeriodShort_(s) {
  if (s.first) return '直近' + MTG_OPTIONS.firstDays + '日間';
  return s.previous && s.previous.baseline ? '基準以降' : '前回以降';
}

/** 「うち前回以降の更新なし」「うち基準以降の更新なし」「うち直近14日間の更新なし」（停滞のうち、対象期間に変化の無い案件） */
function mtgQuietLabel_(s) {
  return 'うち' + mtgPeriodShort_(s) + 'の更新なし';
}

/** 「前回比 」「基準比 」「直近14日間で 」（前回比の書き出し） */
function mtgVersus_(s) {
  if (s.first) return mtgPeriodShort_(s) + 'で ';
  return s.previous && s.previous.baseline ? '基準比 ' : '前回比 ';
}

/** 「（ほかに、その他の項目のみの更新 N件）」（その他更新は「変化のある案件」に数えないので、別に書く。無ければ空） */
function mtgUpdatedNote_(s) {
  return s.kpi.updated ? '（ほかに、その他の項目のみの更新 ' + s.kpi.updated + '件）' : '';
}

/** 「前回から変化のある案件」「直近14日間に変化のある案件」 */
function mtgChangedLabel_(s) {
  if (s.first) return '直近' + MTG_OPTIONS.firstDays + '日間に変化のある案件';
  return (s.previous && s.previous.baseline ? '基準以降に' : '前回から') + '変化のある案件';
}

/**
 * 1 サマリー：リード文に進行中案件の数と期待値合計（前回比）。その下に
 * - 合計の数字（進行中案件・想定売上合計・期待値合計（どれも前回比）・入力内容の確認事項（項目・案件・責任者の数））。
 *   どれも進行中案件全体の数で、下の表の区分とは重ならない
 * - 「前回以降の動き」（変化のある案件を区分ごとに1回だけ数える。行の件数を合わせると見出しの件数・「前回から変化のある案件」の件数になる）
 * - 「現在の状態」（進行中案件の受注月超過・停滞・商談記録未登録。重なることがあるので合計しない）
 * 表の件数は表の中に大きく1回だけ書く（前の版は上に受注／失注・前進／後退・停滞の箱を置いていたが、表と区分のまとめ方が違い、
 * 同じ数に見えて別のものを数えていたため、表にまとめた）
 */
function mtgSummarySlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const b = s.bridge;
  const period = mtgPeriodShort_(s);
  const vs = mtgVersus_(s);
  const lead = '進行中案件 ' + k.active + '件（' + vs + mtgSigned_(k.active - b.startCount) + '件）、期待値合計 ' + mtgAmount_(b.end) + '百万円（' +
    vs + mtgSigned_(b.end - b.start) + '）。';
  // 中身はリード文の下から（リード文が2行になれば下げる）
  const top = Math.max(S.summaryTop, Math.max.apply(null, mtgLead_(lead).map(e => e.y + e.h)) + 2);
  const el = [];

  // 合計の数字：緑の帯の見出しと、大きな数字・小さな前回比（1行）
  const tiles = [
    {label: '進行中案件', value: k.active + '件', sub: vs + mtgSigned_(k.active - b.startCount) + '件'},
    {label: '想定売上合計（進行中・百万円）', value: mtgAmount_(b.amountEnd), sub: vs + mtgSigned_(b.amountEnd - b.amountStart)},
    {label: '期待値合計（進行中・百万円）', value: mtgAmount_(b.end), sub: vs + mtgSigned_(b.end - b.start)},
    {label: '入力内容の確認事項', value: k.alerts + '項目', sub: k.alerts ? k.alertDeals + '案件・責任者 ' + k.owners + '名' : ''}
  ];
  const tw = (S.w - 3 * 8) / 4;
  const th = 36;
  tiles.forEach((t, i) => {
    const x = S.x + i * (tw + 8);
    const text = t.value + (t.sub ? '　' + t.sub : '');
    el.push({type: 'box', x, y: top, w: tw, h: th, fill: c.white, line: c.frame, lineWeight: 0.75, shape: 'RECTANGLE'});
    el.push(mtgBox_(x, top, tw, 14, mtgWrap_(t.label, tw, 8.5, 1)[0],
      {size: 8.5, bold: true, color: c.white, fill: c.green, shape: 'RECTANGLE', align: 'CENTER', role: 'tile'}));
    el.push(mtgBox_(x, top + 14, tw, th - 14, mtgWrap_(text, tw, 8, 1)[0], {size: 8, color: c.muted, align: 'CENTER', role: 'tile-value',
      spans: [{start: 0, end: t.value.length, size: 15, bold: true, color: c.greenDark}]}));
  });

  // 表：案件の名前は1行から始め、ページに入る限り上の行から2行にする。入る分だけ並べ、残りは「他 N件」。件数は大きな緑の字。
  // 区分の名前の下に小さな字の補足（停滞の定義）を付けられる
  const columns = [136, 60, 62, S.w - 258];
  const rowOf = (line, lines) => mtgRow_(columns, [
    line.sub
      ? {text: line.label + '\n' + line.sub, size: 7.5, lines: 2, color: c.muted, fill: c.greenSoft, parts: [{bold: true, size: S.bodySize, color: c.ink}, {}]}
      : {text: line.label, lines: 2, bold: true, color: c.ink, fill: c.greenSoft},
    {text: line.items.length + '件', size: 11, bold: true, color: c.greenDark, fill: c.white, align: 'CENTER'},
    {text: line.items.length ? mtgAmount_(mtgSum_(line.items, 'amount')) : '—', fill: c.white, align: 'END'},
    {text: line.items.length || line.prefix ? mtgNames_(line.shown || line.items, line.note, columns[3], lines, line.prefix) : line.empty || '—', lines,
      fill: c.white}
  ], 0);
  const groups = mtgBreakdown_(s);
  const all = groups.moves.concat(groups.states);
  const headH = 18;
  const gap = 6;
  const tablesTop = top + th + 8;
  const height = rows => rows.reduce((sum, r) => sum + r.h, 0);
  const head = mtgHeadRow_(columns, ['区分', '件数', '想定売上', '主な案件（想定売上順）']);
  const room = S.bottom - tablesTop - 2 * headH - gap - head.h;
  let rows = all.map(line => rowOf(line, 1));
  all.forEach((line, i) => {
    const two = rowOf(line, 2);
    if (height(rows) - rows[i].h + two.h <= room) rows[i] = two;
  });
  // 下にあいた分は、表の行に分けて高くする（1行あたり 10pt まで。ページの下が空かないように）
  const extra = Math.floor(Math.max(0, Math.min(10, (room - height(rows)) / rows.length)) * 10) / 10;
  if (extra > 0) rows = rows.map(r => Object.assign({}, r, {h: r.h + extra}));
  const moves = [head].concat(rows.slice(0, groups.moves.length));
  const states = rows.slice(groups.moves.length);
  const heading = (y, text) => mtgBox_(S.x, y, S.w, headH, text, {size: 10.5, bold: true, color: c.ink, role: 'heading'});
  const statesY = tablesTop + headH + height(moves) + gap;
  el.push(heading(tablesTop, period + 'の動き（変化のある案件 ' + k.changed + '件）'),
    {type: 'table', x: S.x, y: tablesTop + headH, columns, rows: moves},
    heading(statesY, '現在の状態（進行中案件）'),
    {type: 'table', x: S.x, y: statesY + headH, columns, rows: states});
  return {title: MTG_TITLES.summary, lead, elements: el};
}

/**
 * サマリーの表の行 {moves, states}。どの行も {label, sub: 区分の名前の下の小さな補足, items: 数える案件（想定売上の大きい順）,
 * shown: 名前を並べる案件（無ければ items）, prefix: 名前の前に書く文, empty: 該当が無いときの文, note: 案件の名前の後ろの（）の中}。
 * moves：前回以降の動き。案件ごとの区分（MTG_CATEGORIES）で1回だけ数え、すべての区分（その他更新を除く）をどれかの行に入れる
 * （受注／失注・保留・後退・削除／新規・前進・復帰／活動更新／受注月・期待値などの変更。P4 の並び（MTG_MOVE_ORDER）と同じ順）。
 * 活動更新は、最新の活動状況を（）に書く。
 * states：現在の状態（受注月超過・停滞・商談記録未登録）。停滞は、前回以降に更新の無い案件を「うち前回以降の更新なし N件」として名前を並べる
 * （初回は「直近N日間」、基準の後は「基準以降」。mtgQuietLabel_）。
 * 商談記録未登録は、シートに商談記録の列があるときだけ
 */
function mtgBreakdown_(s) {
  const o = MTG_OPTIONS;
  const tz = Session.getScriptTimeZone();
  const byAmount = list => list.slice().sort((a, b) => (b.amount || 0) - (a.amount || 0));
  const of = keys => byAmount(s.changes.filter(ch => keys.indexOf(ch.category) >= 0));
  const label = ch => mtgCategory_(ch.category).label;
  const status = ch => (ch.statusFrom || '（未入力）') + ' ▶ ' + (ch.statusTo || '（未入力）');
  const month = ch => mtgMonthLabel_(ch.monthFrom) + ' ▶ ' + mtgMonthLabel_(ch.monthTo);
  const value = ch => mtgAmount_(ch.weightedFrom) + ' ▶ ' + mtgAmount_(ch.weightedTo);
  const quiet = s.stalled.filter(d => !d.touched);
  const pending = s.stallPending;
  const states = [
    {label: '受注月超過', items: byAmount(s.overdue), note: d => '受注月 ' + mtgMonthLabel_(d.orderMonthText)},
    {label: '停滞', sub: '（同一ステータス' + o.stallDays + '日以上）', items: s.stalled, shown: quiet,
      prefix: s.stalled.length ? mtgQuietLabel_(s) + ' ' + quiet.length + '件' + (quiet.length ? '：' : '') : '',
      empty: pending ? (pending.baseline ? '基準' : '記録の開始') + '（' + Utilities.formatDate(new Date(pending.at), tz, 'M/d') + '）から ' +
        o.stallDays + '日未満のため、まだ数えられない' : '—',
      note: d => d.days + '日' + (d.atLeast ? '以上' : '')}
  ];
  if (s.recordColumn) {
    states.push({label: '商談記録未登録', items: byAmount(s.alerts.filter(a => a.type === '資料なし'))});
  }
  return {
    moves: [
      {label: '受注', items: of(['won'])},
      {label: '失注・保留・後退・削除', items: of(['lost', 'hold', 'backward', 'removed']), note: ch => (ch.category === 'backward' ? status(ch) : label(ch))},
      {label: '新規・前進・復帰', items: of(['new', 'forward', 'revived']), note: ch => (ch.category === 'forward' ? status(ch) : label(ch))},
      {label: '活動更新', items: of(['activity']), note: ch => mtgClip_(ch.activity || '—', 20)},
      {label: '受注月・期待値などの変更', items: of(['slipped', 'pulled', 'up', 'down', 'status']),
        note: ch => label(ch) + ({slipped: ' ' + month(ch), pulled: ' ' + month(ch), up: ' ' + value(ch), down: ' ' + value(ch), status: ' ' + status(ch)}[ch.category] || '')}
    ],
    states
  };
}

/** 文を全角 max 文字ぶんまでに切る（長ければ「…」） */
function mtgClip_(text, max) {
  let used = 0;
  let out = '';
  const chars = Array.from(String(text).replace(/\s+/g, ' ').trim());
  for (let i = 0; i < chars.length; i++) {
    if (used + mtgCharWidth_(chars[i]) > max) return out + '…';
    out += chars[i];
    used += mtgCharWidth_(chars[i]);
  }
  return out;
}

/**
 * 案件の名前を並べる（mainRows.deals 件まで。残りは「他 N件」）。note(案件) があれば名前の後ろに（）で付ける。
 * width・lines を渡すと、その幅・行数に切れずに入る件数まで減らす（「他 N件」が「…」で切れないように）。prefix は名前の前に書く文
 */
function mtgNames_(items, note, width, lines, prefix) {
  const head = prefix || '';
  if (!items.length) return head || '—';
  const text = n => head + items.slice(0, n).map(d => mtgDealLabel_(d) + (note ? '（' + note(d) + '）' : '')).join('、') +
    (items.length > n ? '　他 ' + (items.length - n) + '件' : '');
  let n = Math.min(MTG_OPTIONS.mainRows.deals, items.length);
  while (width && n > 1 && mtgWrap_(text(n), width, MTG_SLIDE.bodySize, 99).length > lines) n--;
  return text(n);
}

/** 2 パイプラインの変化：期待値の増減（滝の図・凡例つき）と、ステータス別・受注月（四半期）別の期待値（横棒のグラフ。同じ目盛り） */
function mtgBridgeSlide_(s) {
  const S = MTG_SLIDE;
  const b = s.bridge;
  const delta = b.end - b.start;
  const moves = b.steps.filter(st => st.count);
  const biggest = moves.slice().sort((x, y) => Math.abs(y.value) - Math.abs(x.value))[0];
  const plus = mtgSum_(b.steps.filter(st => st.value > 0), 'value');
  const minus = mtgSum_(b.steps.filter(st => st.value < 0), 'value');
  const peak = s.quarters.filter(q => /Q/.test(q.label) && q.value > 0).sort((x, y) => y.value - x.value)[0];
  const over = s.quarters.find(q => q.label === '受注月超過');
  const won = b.steps.find(st => st.key === 'won' && st.count);
  const el = mtgWaterfall_(s, S.x, S.top, 360, S.bottom - S.top);
  const stages = s.stages.filter(st => st.kind === 'active');
  const shown = stages.slice(0, 8);
  if (stages.length > 8) {
    const rest = stages.slice(7);
    shown.splice(7, 1, {status: 'その他 ' + rest.length + '種類', count: mtgSum_(rest, 'count'), weighted: mtgSum_(rest, 'weighted')});
  }
  const right = S.x + 376;
  const width = S.w - 376;
  const statusRows = shown.map(st => ({label: st.status, count: st.count, value: st.weighted || 0}));
  const monthRows = s.quarters.map(q => ({label: q.label, count: q.count, value: q.value || 0}));
  // 2つのグラフは同じ目盛り（どちらも進行中の期待値の内訳なので、棒の長さで比べられるように）。行の高さは、下まで入るように
  const max = Math.max.apply(null, statusRows.concat(monthRows).map(r => r.value).concat([0]));
  const gap = 14;
  const rowH = Math.min(15, (S.bottom - S.top - 2 * 22 - gap) / Math.max(1, statusRows.length + monthRows.length));
  const byStatus = mtgBarChart_('ステータス別期待値（進行中・百万円）', statusRows, right, S.top, width, max, rowH, '進行中案件なし');
  const byMonth = mtgBarChart_('受注月別期待値（進行中・百万円）', monthRows, right, byStatus.bottom + gap, width, max, rowH, '進行中案件なし');
  el.push.apply(el, byStatus.elements.concat(byMonth.elements));
  return {
    title: MTG_TITLES.pipeline,
    lead: '進行中の期待値は ' + mtgAmount_(b.start) + ' → ' + mtgAmount_(b.end) + '百万円（' + mtgSigned_(delta) + '）。増加 ' + mtgSigned_(plus) +
      '・減少 ' + mtgSigned_(minus) + (won && won.value ? '（うち受注 ' + mtgSigned_(won.value) + '）' : '') +
      (biggest && biggest.value ? '、内訳の最大は' + biggest.label + '（' + biggest.count + '件・' + mtgSigned_(biggest.value) + '）' : '') +
      '。\n受注月が今期の期待値は ' + mtgAmount_(mtgSum_(s.quarters.filter(q => /Q/.test(q.label)), 'value')) + '百万円' +
      (peak ? '（最大は' + peak.label.replace(/（.*/, '') + 'の ' + mtgAmount_(peak.value) + '）' : '') +
      (over ? '、受注月超過は ' + mtgAmount_(over.value) + '百万円' : '') +
      (s.kpi.missingWeighted ? '。期待値未入力の ' + s.kpi.missingWeighted + '件は0として集計' : '') + '。',
    elements: el
  };
}

/** 期待値の増減の滝の図（前回 → 内訳 → 今回）。見出しと凡例（期待値・増加・減少・受注）つき。左上 (x, y)・幅 w・高さ h。縦軸は0から */
function mtgWaterfall_(s, x, y, w, h) {
  const c = MTG_OPTIONS.colors;
  const b = s.bridge;
  const el = [mtgBox_(x, y - 2, 190, 20, '期待値の増減（進行中・百万円）', {size: 10.5, bold: true, color: c.ink})];
  // 凡例：色の四角と名前。名前の箱は文字の幅＋左右の余白（14.4）＋少し。右端を図の右端にそろえ、項目の間は 8pt あける
  const legend = [['期待値', c.greenDark], ['増加', c.up], ['減少', c.down], ['受注', c.white, c.greenDark]]
    .map(([label, color, line]) => ({label, color, line, bw: Array.from(label).length * 8.5 + 16}));
  let lx = x + w - legend.reduce((sum, item) => sum + 6 + item.bw, 0) - (legend.length - 1) * (8 - 7.2 - 1.6);
  legend.forEach(item => {
    el.push({type: 'box', x: lx, y: y + 5, w: 7, h: 7, fill: item.color, line: item.line, lineWeight: item.line ? 1 : undefined, shape: 'RECTANGLE', role: 'legend'});
    el.push(mtgBox_(lx + 6, y - 1, item.bw, 18, item.label, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += 6 + item.bw + 8 - 7.2 - 1.6;   // 名前の箱の右の余白（7.2）と、文字のあとの少し（1.6）の分を詰めて、文字の終わりから 8pt
  });
  const bars = [{label: s.first ? MTG_OPTIONS.firstDays + '日前' : '前回', value: b.start, count: b.startCount, total: true}]
    .concat(b.steps.filter(st => st.count))
    .concat([{label: '今回', value: b.end, count: b.endCount, total: true}]);
  const top = y + 42;
  const axis = y + h - 34;
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
  const bw = Math.min(34, slot * 0.7);
  el.push({type: 'box', x, y: axis, w, h: 0.75, fill: c.faint, shape: 'RECTANGLE'});
  bars.forEach((bar, i) => {
    const span = spans[i];
    const bx = x + slot * i + (slot - bw) / 2;
    const bh = Math.max((span.hi - span.lo) * scale, 1);
    const by = axis - span.lo * scale - bh;
    // 受注は進行中から外れるので減少だが、失注などと分けて白地に濃い緑の枠で描く
    const won = bar.key === 'won';
    el.push({type: 'box', x: bx, y: by, w: bw, h: bh, fill: bar.total ? c.greenDark : won ? c.white : bar.value >= 0 ? c.up : c.down,
      line: won ? c.greenDark : undefined, lineWeight: won ? 1 : undefined, shape: 'RECTANGLE', role: 'bar-' + (bar.key || 'total')});
    el.push(mtgBox_(x + slot * i - 8, by - 15, slot + 16, 14, bar.total ? mtgAmount_(bar.value) : mtgSigned_(bar.value),
      {size: 9, bold: true, color: c.ink, align: 'CENTER'}));
    const name = bar.label;
    el.push(mtgBox_(x + slot * i - 8, axis + 3, slot + 16, 28, name + '\n' + bar.count + '件',
      {size: 8.5, color: c.muted, align: 'CENTER', valign: 'TOP', spans: [{start: 0, end: name.length, bold: true, color: c.ink}]}));
  });
  if (!bars.some(bar => bar.value)) {
    el.push(mtgBox_(x, top + 40, w, 20, '進行中案件の期待値なし', {size: 9, color: c.muted, align: 'CENTER'}));
  }
  return el;
}

/**
 * 横棒のグラフ（パイプラインの変化の右。ステータス別・受注月別の期待値）。見出しは滝の図と同じ形で、行ごとに 名前・件数・棒・値。
 * 棒は期待値の色（greenDark。滝の図の「期待値」と同じ）。値は棒の先に書くので、目盛りは置かない（0 の線だけ）。
 * rows = [{label, count, value}]（この順に上から）。max は棒の長さの基準（2つのグラフで同じにする）。左上 (x, y)・幅 w・行の高さ rowH。
 * 戻り値 {elements, bottom}
 */
function mtgBarChart_(title, rows, x, y, w, max, rowH, empty) {
  const c = MTG_OPTIONS.colors;
  const labelW = 92;
  const countW = 40;                         // 件数の箱（右寄せ。「123件」が折り返さない幅。名前の箱に 10pt 重ねる）
  const valueW = 51;                         // 値の箱（「1,234.5」が折り返さない幅）
  const base = x + labelW + countW - 6;      // 棒の始まり（0 の線）
  const room = x + w - base - (valueW - 3);  // 棒のいちばん長い長さ（右に値を書く分をあける。値は棒の先から 4pt あける）
  const top = y + 22;
  const el = [mtgBox_(x, y - 2, w, 20, title, {size: 10.5, bold: true, color: c.ink})];
  if (!rows.length) {
    el.push(mtgBox_(x, top, w, 18, empty, {size: 8.5, color: c.muted}));
    return {elements: el, bottom: top + 18};
  }
  const bar = Math.min(9, rowH - 5);
  el.push({type: 'box', x: base - 0.75, y: top, w: 0.75, h: rows.length * rowH, fill: c.faint, shape: 'RECTANGLE'});
  rows.forEach((r, i) => {
    const ry = top + i * rowH;
    const len = max > 0 ? Math.max(0, r.value) / max * room : 0;
    el.push(mtgBox_(x, ry, labelW, rowH, mtgWrap_(r.label, labelW, 8.5, 1)[0], {size: 8.5, color: c.ink, role: 'bar-label'}));
    el.push(mtgBox_(x + labelW - 10, ry, countW, rowH, r.count + '件', {size: 8, color: c.muted, align: 'END'}));
    if (len > 0) el.push({type: 'box', x: base, y: ry + (rowH - bar) / 2, w: Math.max(len, 1), h: bar, fill: c.greenDark, shape: 'RECTANGLE', role: 'bar'});
    el.push(mtgBox_(base + len - 3, ry, valueW, rowH, mtgAmount_(r.value), {size: 8.5, bold: true, color: c.ink, role: 'bar-value'}));
  });
  return {elements: el, bottom: top + rows.length * rowH};
}

/** 3 前回から変化のある案件：ステータス・期待値・受注月・活動状況が変わった案件（区分の順に上位。活動更新は最後）。区分はラベルで出す */
function mtgMovesSlide_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const moves = s.changes.filter(ch => ch.category !== 'updated');
  const top = moves.slice(0, MTG_OPTIONS.mainRows.changes);
  const columns = mtgColumns_([138, 66, 118, 66, 76, 144, 52], 5);
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '区分', 'ステータス\n前回▶今回', '期待値\n前回▶今回', '受注月\n前回▶今回', '活動状況（最新）', '責任者'])];
  top.forEach(ch => {
    rows.push(mtgRow_(columns, [
      mtgDealCell_(ch, columns[0]),
      {text: '', fill: c.white},
      {text: mtgStatusMove_(ch), lines: 2, fill: c.white},
      {text: mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), lines: 2, fill: c.white, align: 'CENTER'},
      {text: mtgMonthMove_(ch), lines: 2, fill: c.white, align: 'CENTER'},
      {text: ch.activity || '—', lines: 2, fill: c.white},
      {text: ch.owner || '—', lines: 2, fill: c.white, align: 'CENTER'}
    ]));
  });
  const el = [];
  if (top.length) {
    const table = {type: 'table', x: S.x, y: S.top, columns, rows};
    el.push(table);
    top.forEach((ch, i) => el.push.apply(el, mtgPills_(table, i + 1, 1, [mtgCategoryPill_(ch.category)], mtgCategoryPillWidth_(columns[1]))));
  } else {
    el.push(mtgBox_(S.x, S.top + 20, S.w, 30, mtgPeriodShort_(s) + '、ステータス・期待値・受注月・活動状況に変化のある案件なし',
      {size: 10, color: c.muted, align: 'CENTER'}));
  }
  return {
    title: MTG_TITLES.progress,
    lead: (moves.length ? mtgChangedLabel_(s) + 'は ' + moves.length + '件' + mtgUpdatedNote_(s) + '。'
      : mtgPeriodShort_(s) + '、ステータス・期待値・受注月・活動状況に変化のある案件なし' + mtgUpdatedNote_(s) + '。') +
      (moves.length > top.length ? '\n区分の順（受注・失注・保留・後退・削除・新規・前進…）に ' + top.length + '件を掲載。全件は Appendix ① を参照。'
        : k.updated ? '\n全件は Appendix ① を参照。' : ''),
    elements: el
  };
}

/** 4 期待値上位案件の状況・課題：進行中の案件の期待値の大きい順（summary.focus）。当てはまる状況（停滞・受注月超過・延期）は事実だけをラベルで */
function mtgFocusSlide_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const columns = mtgColumns_([146, 62, 84, 56, 98, 142, 72], 4);
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '期待値／\n想定売上', 'ステータス\n（滞留日数）', '受注月', '活動状況', '活動における課題',
    '該当事項'])];
  s.focus.forEach(d => rows.push(mtgRow_(columns, [
    mtgDealCell_(d, columns[0]),
    {text: mtgAmount_(d.weighted) + '\n' + mtgAmount_(d.amount), parts: [{bold: true, color: c.greenDark, size: 11}, {}], lines: 2, fill: c.white,
      align: 'CENTER'},
    {text: (d.status || '（未入力）') + '\n' + (d.days == null ? '—' : d.days + '日' + (d.atLeast ? '以上' : '')), parts: [{}, {color: c.muted}],
      lines: 3, fill: c.white, align: 'CENTER'},
    {text: mtgMonthLabel_(d.orderMonthText), fill: c.white, align: 'CENTER'},
    {text: d.activity || '—', lines: 3, fill: c.white},
    {text: d.issue || '—', lines: 3, fill: c.white},
    {text: mtgFlags_(d).length ? '' : '—', fill: c.white, align: 'CENTER'}
  ], Math.max(44, mtgPillsHeight_(mtgFlags_(d).length)))));
  const el = [];
  if (s.focus.length) {
    const table = {type: 'table', x: S.x, y: S.top, columns, rows};
    el.push(table);
    const width = mtgPillWidth_(['停滞', '受注月超過', '延期9回'], columns[6]);
    s.focus.forEach((d, i) => el.push.apply(el, mtgPills_(table, i + 1, 6, mtgFlags_(d).map(flag => ({text: flag, tone: flag === '停滞' ? 'stall' : 'note'})),
      width)));
  } else {
    el.push(mtgBox_(S.x, S.top + 20, S.w, 30, '進行中案件なし', {size: 10, color: c.muted, align: 'CENTER'}));
  }
  const share = k.weighted > 0 ? Math.round(mtgSum_(s.focus, 'weighted') / k.weighted * 100) : 0;
  return {
    title: MTG_TITLES.focus,
    lead: s.focus.length ? '進行中案件のうち期待値上位 ' + s.focus.length + '件（進行中案件の期待値合計の ' + share + '%）。' : '進行中案件なし。',
    elements: el
  };
}

/** 列の幅（合わせて 660pt で決めた幅）を中身の幅（MTG_SLIDE.w）に合わせる。差は flex 番目の列（文の列）で調整する */
function mtgColumns_(columns, flex) {
  const out = columns.slice();
  out[flex] += MTG_SLIDE.w - columns.reduce((a, b) => a + b, 0);
  return out;
}

/** 案件に当てはまる状況（停滞・受注月超過・延期n回） */
function mtgFlags_(d) {
  return [d.stalled && '停滞', d.overdue && '受注月超過', d.slips && '延期' + d.slips + '回'].filter(Boolean);
}

/** 区分のラベル {text, tone}（mtgPills_ に渡す） */
function mtgCategoryPill_(key) {
  const cat = mtgCategory_(key);
  return {text: cat.label, tone: cat.tone || 'note'};
}

/**
 * 表のセルの上に重ねるラベル（角の丸い四角。区分・当てはまる状況）。table は表の部品（x・y・columns・rows）、row・col はセルの位置
 * （見出しの行が 0）、labels は [{text, tone}]（tone：up 前向き＝濃い緑・down 後ろ向き＝濃いグレー・note そのほか＝薄いグレー・stall 停滞＝薄い緑。
 * どれも枠なし）。width はラベルの幅（列の中で同じにする。mtgPillWidth_）。複数は縦に並べ、セルの上下の中央に置く。セルの文字は空にしておく
 */
function mtgPills_(table, row, col, labels, width) {
  const c = MTG_OPTIONS.colors;
  const P = MTG_SLIDE.pill;
  const styles = {
    up: {fill: c.greenDeep, color: c.white},
    down: {fill: c.dark, color: c.white},
    note: {fill: c.graySoft, color: c.sub},
    stall: {fill: c.greenSoft, color: c.greenText}
  };
  const x0 = table.x + table.columns.slice(0, col).reduce((a, b) => a + b, 0);
  const y0 = table.y + table.rows.slice(0, row).reduce((a, r) => a + r.h, 0);
  const cw = table.columns[col];
  const pw = Math.min(width || mtgPillWidth_(labels.map(label => label.text), cw), cw - 8);
  const total = mtgPillsHeight_(labels.length) - 2 * P.margin;
  return labels.map((label, i) => mtgBox_(x0 + (cw - pw) / 2, y0 + (table.rows[row].h - total) / 2 + i * (P.h + P.gap), pw, P.h, label.text,
    Object.assign({size: P.size, bold: true, align: 'CENTER', shape: 'ROUND_RECTANGLE', role: 'pill'}, styles[label.tone] || styles.note)));
}

/** ラベルの幅：texts のうちいちばん長いものが1行に入る幅（列の幅 cw から左右 4pt ずつあける）。列の中のラベルをそろえるため、出しうる文字をすべて渡す */
function mtgPillWidth_(texts, cw) {
  const P = MTG_SLIDE.pill;
  const need = Math.max.apply(null, texts.map(text => Array.from(text).reduce((sum, ch) => sum + mtgCharWidth_(ch), 0)));
  return Math.min(cw - 8, Math.ceil(need * P.size / 0.92 + MTG_SLIDE.padX + 4));
}

/** ラベルを n 個縦に並べたときに要るセルの高さ（上下の余白を含む） */
function mtgPillsHeight_(n) {
  const P = MTG_SLIDE.pill;
  return n ? n * P.h + (n - 1) * P.gap + 2 * P.margin : 0;
}

/** 区分のラベルの幅（どの区分でも同じ幅） */
function mtgCategoryPillWidth_(cw) {
  return mtgPillWidth_(MTG_CATEGORIES.map(cat => cat.label), cw);
}

/** Appendix ① 前回から変化のある案件（一覧。その他更新・削除も）。区分はラベルで出す */
function mtgAppendixChanges_(s) {
  const c = MTG_OPTIONS.colors;
  const columns = mtgColumns_([120, 66, 112, 62, 72, 104, 74, 50], 5);
  return mtgPagedTable_({
    title: MTG_TITLES.changes,
    lead: mtgChangedLabel_(s) + 'は ' + s.kpi.changed + '件' + mtgUpdatedNote_(s) + '。' +
      (s.unmatched ? '\n案件と照合できなかった変更記録 ' + s.unmatched + '件は集計対象外。' : ''),
    columns,
    head: ['得意先／案件名', '区分', 'ステータス\n前回▶今回', '期待値\n前回▶今回', '受注月\n前回▶今回', '活動状況（最新）', '更新項目', '責任者'],
    rows: s.changes.map(ch => [
      mtgDealCell_(ch, columns[0], MTG_SLIDE.appendixSize),
      {text: '', fill: c.white},
      {text: mtgStatusMove_(ch), lines: 2, fill: c.white},
      {text: mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), lines: 2, fill: c.white, align: 'CENTER'},
      {text: mtgMonthMove_(ch), lines: 2, fill: c.white, align: 'CENTER'},
      {text: ch.removed ? '（削除）' : ch.activity || '—', lines: 2, fill: c.white},
      {text: ch.category === 'new' ? '（新規追加）' : ch.columns.map(mtgShortHeader_).join('・') || (ch.removed ? '（削除）' : '—'), lines: 2,
        fill: c.white},
      {text: ch.owner || '—', fill: c.white, align: 'CENTER'}
    ]),
    pills: {col: 1, labels: i => [mtgCategoryPill_(s.changes[i].category)], width: mtgCategoryPillWidth_(columns[1])}
  });
}

/** Appendix ② 停滞案件（一覧）。同じステータスの日数と、最新の活動状況（前回以降に更新があるかが分かるように） */
function mtgAppendixStalled_(s) {
  const c = MTG_OPTIONS.colors;
  const k = s.kpi;
  const columns = mtgColumns_([128, 40, 44, 80, 62, 54, 34, 44, 92, 82], 9);
  return mtgPagedTable_({
    title: MTG_TITLES.stalled,
    lead: '進行中案件のうち、同一ステータスに ' + MTG_OPTIONS.stallDays + '日以上滞留している案件は ' + s.stalled.length + '件（想定売上合計 ' +
      mtgAmount_(k.stalledAmount) + '百万円）。\n' + mtgQuietLabel_(s) + ' ' + k.stalledQuiet + '件。',
    columns,
    head: ['得意先／案件名', '想定\n売上', '期待値', 'ステータス', '滞留日数', '受注月', '延期\n回数', '責任者', '活動における課題', '活動状況（最新）'],
    rows: s.stalled.map(d => [
      mtgDealCell_(d, columns[0], MTG_SLIDE.appendixSize),
      {text: mtgAmount_(d.amount), fill: c.white, align: 'END'},
      {text: mtgAmount_(d.weighted), fill: c.white, align: 'END'},
      {text: d.status || '（未入力）', lines: 2, fill: c.white},
      {text: d.days + '日' + (d.atLeast ? '以上' : ''), fill: c.white, align: 'CENTER'},
      {text: mtgMonthLabel_(d.orderMonthText), fill: c.white, align: 'CENTER'},
      {text: d.slips ? d.slips + '回' : '', fill: c.white, align: 'CENTER'},
      {text: d.owner || '—', fill: c.white, align: 'CENTER'},
      {text: d.issue || '—', lines: 2, fill: c.white},
      {text: d.activity || '—', lines: 2, fill: c.white}
    ])
  });
}

/**
 * Appendix ③ 入力内容のチェック（責任者別）：責任者 × 種類の件数・試作から解消した件数（当てはまるものが無ければ作らない）。
 * 種類の説明は置かない（案件ごとの一覧に内容を書く）
 */
function mtgAppendixOwners_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  if (!s.alerts.length) return [];
  const types = MTG_ALERT_TYPES.filter(t => s.alerts.some(a => a.type === t));
  const groups = new Map();
  s.alerts.forEach(a => {
    const owner = a.owner || '（責任者未設定）';
    if (!groups.has(owner)) groups.set(owner, []);
    groups.get(owner).push(a);
  });
  // 責任者は名前の順（件数の順に並べると順位表に見えるため）。入りきらないときは、件数の少ない責任者を「その他 N名」にまとめる
  const limit = MTG_OPTIONS.mainRows.owners;
  const byCount = Array.from(groups.keys()).sort((a, b) => groups.get(b).length - groups.get(a).length);
  const shown = byCount.slice(0, byCount.length > limit ? limit - 1 : limit);
  const owners = mtgOwnerOrder_(shown);
  const lines = owners.map(owner => ({label: owner, list: groups.get(owner)}));
  if (byCount.length > limit) {
    const rest = byCount.slice(limit - 1);
    lines.push({label: 'その他 ' + rest.length + '名', list: [].concat.apply([], rest.map(owner => groups.get(owner)))});
  }
  lines.push({label: '合計', list: s.alerts, total: true});

  // 表はページの幅まで（種類が少なければ、種類の列は 80pt までにして左に寄せる）
  const ownerW = 120;
  const totalW = 40;
  const typeW = Math.min(80, (S.w - ownerW - totalW) / types.length);
  const columns = [ownerW].concat(types.map(() => typeW), [totalW]);
  const short = {'予定日超過': '予定日\n超過', '受注月超過': '受注月\n超過', '未入力': '必須項目\n未入力', '資料なし': '商談記録\n未登録', '更新なし': '長期\n未更新',
    '金額の不整合': '金額\n不整合', '月の不整合': '売上開始月\n不整合', '選択肢外のステータス': '選択肢外\nステータス', 'デリバリー担当なし': 'デリバリー\n担当未設定'};
  const head = mtgHeadRow_(columns, ['責任者'].concat(types.map(t => short[t] || t), ['計']), typeW >= 64 ? 8.5 : 7);
  head.cells[head.cells.length - 1].fill = c.greenDeep;
  const rows = [head];
  lines.forEach(line => {
    const fill = line.total ? c.greenSoft : c.white;
    rows.push(mtgRow_(columns, [{text: line.label, lines: 2, bold: true, color: c.ink, fill: c.greenSoft, align: 'CENTER'}]
      .concat(types.map(t => {
        const n = line.list.filter(a => a.type === t).length;
        return {text: n ? String(n) : '', fill, align: 'CENTER', bold: !!line.total};
      }), [{text: String(line.list.length), fill, align: 'CENTER', bold: true, color: c.ink}]), 20));
  });
  const r = s.resolved;
  return [{
    title: MTG_TITLES.owners,
    lead: '入力内容の確認事項は ' + s.alerts.length + '項目（' + k.alertDeals + '案件・責任者 ' + k.owners + '名）。' +
      (r ? '試作時（' + Utilities.formatDate(new Date(r.at), Session.getScriptTimeZone(), 'M/d') + '）から ' + r.fixed + '項目解消、' + r.added + '項目追加。' : ''),
    elements: [{type: 'table', x: S.x, y: S.top, columns, rows}]
  }];
}

/** 責任者の並び：名前の順（localeCompare の日本語の順。漢字は JIS の順で、読みの50音順ではない。責任者が空は最後） */
function mtgOwnerOrder_(owners) {
  const blank = '（責任者未設定）';
  return owners.filter(o => o !== blank).sort((a, b) => a.localeCompare(b, 'ja')).concat(owners.indexOf(blank) >= 0 ? [blank] : []);
}

/**
 * Appendix ③ 入力内容の確認事項（一覧）：責任者（名前の順）→ 案件（想定売上の大きい順）に並べ、1案件を1行にまとめる
 * （責任者が自分の分を一か所で見られるように）。項目と内容は1項目ずつ行を分け、項目の名前と内容の行をそろえる
 */
function mtgAppendixAlerts_(s) {
  const c = MTG_OPTIONS.colors;
  const size = MTG_SLIDE.appendixSize;
  const columns = mtgColumns_([64, 150, 104, 286, 56], 3);
  const groups = new Map();
  s.alerts.forEach(a => {
    const owner = a.owner || '（責任者未設定）';
    if (!groups.has(owner)) groups.set(owner, new Map());
    const deals = groups.get(owner);
    const key = a.dealKey || a.key;
    if (!deals.has(key)) deals.set(key, []);
    deals.get(key).push(a);
  });
  const rows = [];
  mtgOwnerOrder_(Array.from(groups.keys())).forEach(owner => {
    Array.from(groups.get(owner).values()).sort((x, y) => (y[0].amount || 0) - (x[0].amount || 0)).forEach(list => {
      const types = [];
      const texts = [];
      list.forEach(a => {
        const t = mtgWrap_(MTG_ALERT_LABELS[a.type] || a.type, columns[2], size, 2);
        const x = mtgWrap_(a.text, columns[3], size, 2);
        for (let n = 0; n < Math.max(t.length, x.length); n++) {
          types.push(t[n] || '');
          texts.push(x[n] || '');
        }
      });
      rows.push([
        {text: owner, lines: 2, bold: true, color: c.ink, fill: c.greenSoft, align: 'CENTER'},
        Object.assign(mtgDealCell_(list[0], columns[1], size), {fill: c.white}),
        {text: types.join('\n'), lines: types.length, bold: true, color: c.ink, fill: c.white},
        {text: texts.join('\n'), lines: texts.length, fill: c.white},
        {text: mtgAmount_(list[0].amount), fill: c.white, align: 'END'}
      ]);
    });
  });
  return mtgPagedTable_({
    title: MTG_TITLES.alerts,
    lead: '入力内容の確認事項は ' + s.alerts.length + '項目（' + s.kpi.alertDeals + '案件）。',
    columns,
    head: ['責任者', '得意先／案件名', '項目', '内容', '想定\n売上'],
    rows
  });
}

/**
 * Appendix の表を、入りきる行ずつのページに分ける（見出しの行は各ページに）。行が無ければページを作らない。
 * spec = {title, lead, columns, head, rows: [[セル]], pills: {col, labels: 行の番号 → [{text, tone}]}}（pills はセルに重ねるラベル。mtgPills_）
 */
function mtgPagedTable_(spec) {
  const S = MTG_SLIDE;
  const head = mtgHeadRow_(spec.columns, spec.head, S.appendixSize);
  const pages = [];
  let page = null;
  spec.rows.forEach((cells, index) => {
    const row = mtgRow_(spec.columns, cells.map(cell => Object.assign({size: S.appendixSize}, cell)), 0);
    if (!page || page.y + row.h > S.bottom) {
      page = {rows: [head], y: S.top + head.h, index: []};
      pages.push(page);
    }
    page.rows.push(row);
    page.index.push(index);
    page.y += row.h;
  });
  return pages.map((p, i) => {
    const table = {type: 'table', x: S.x, y: S.top, columns: spec.columns, rows: p.rows};
    const el = [table];
    if (spec.pills) p.index.forEach((index, n) => el.push.apply(el, mtgPills_(table, n + 1, spec.pills.col, spec.pills.labels(index), spec.pills.width)));
    return {title: spec.title + (i ? '（続き）' : ''), lead: spec.lead, elements: el};
  });
}

/** ステータスの動き（「提案中 ▶ 受注」。変わっていなければ今の値） */
function mtgStatusMove_(ch) {
  const blank = v => v || '（未入力）';
  if (ch.removed) return blank(ch.statusFrom) + ' ▶ （削除）';
  if (!ch.base) return '— ▶ ' + blank(ch.statusTo);
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
  const names = {customer: '得意先', project: '案件名', status: 'ステータス', amount: '想定売上', weighted: '期待値', orderMonth: '受注月',
    startMonth: '売上開始月', activity: '活動状況', issue: '課題', owner: '責任者', delivery: 'デリバリー担当', record: '商談記録'};
  const labels = mtgLabels_();
  const key = Object.keys(labels).find(k => labels[k] === credNormalize_(header));
  return key ? names[key] : String(header).replace(/\s+/g, '');
}

/** 「前回会議（10/9）の資料作成時（10/07 18:00）以降」「直近14日間（09/30〜。初回のため）」 */
function mtgPeriodText_(s) {
  const tz = Session.getScriptTimeZone();
  if (s.previous && s.previous.baseline) {
    return '基準時点（' + Utilities.formatDate(new Date(s.previous.cutoff), tz, 'M/d HH:mm') + '）以降';
  }
  if (s.previous) {
    return '前回会議（' + mtgShortDate_(s.previous.date) + '）の資料作成時（' + Utilities.formatDate(new Date(s.previous.cutoff), tz, 'M/d HH:mm') + '）以降';
  }
  if (s.legacy) return '前回作成時（' + Utilities.formatDate(s.since, tz, 'M/d HH:mm') + '）以降';
  return '直近' + MTG_OPTIONS.firstDays + '日間（' + Utilities.formatDate(s.since, tz, 'M/d') + '〜。初回のため）';
}

/* ---------------- スライドの部品 ---------------- */

/** 文字の箱 */
function mtgBox_(x, y, w, h, text, style) {
  return Object.assign({type: 'box', x, y, w, h, text}, style || {});
}

/** 得意先（1行目・太字）と案件名（2行目）のセル（1列目なので薄い緑）。size は文字の大きさ（既定は bodySize） */
function mtgDealCell_(item, width, size) {
  const c = MTG_OPTIONS.colors;
  const pt = size || MTG_SLIDE.bodySize;
  return {text: mtgWrap_(item.customer || '（得意先未入力）', width, pt, 1)[0] + '\n' + mtgWrap_(item.project || '—', width, pt, 1)[0],
    size: pt, parts: [{bold: true}, {}], lines: 2, color: c.ink, fill: c.greenSoft};
}

/**
 * 表の1行。cells: [{text, lines, size, bold, color, fill, align, parts}]（文字は列の幅で折り返し、lines 行（既定は1）を超えたら「…」）。
 * parts は改行で分けた部分ごとの書式 [{bold, color, size}]（スライドでは spans にする）。行の高さは、いちばん行数の多いセルに合わせる（min 以上）。
 */
function mtgRow_(columns, cells, min) {
  let h = min || 0;
  const out = cells.map((cell, i) => {
    const size = cell.size || MTG_SLIDE.bodySize;
    const lines = mtgBalance_(cell.text, columns[i], size, cell.lines || 1);
    h = Math.max(h, mtgLineHeight_(lines.length, size));
    const text = lines.join('\n');
    return Object.assign({}, cell, {text, size}, cell.parts ? {spans: mtgPartSpans_(String(cell.text), cell.parts, columns[i], size, lines)} : {});
  });
  return {h: Math.ceil(h * 10) / 10, cells: out};
}

/**
 * 改行で分けた部分ごとの書式（parts）を、折り返したあとの文字の位置の spans [{start, end, bold, color, size}] にする。
 * 部分ごとに折り返した行の数を数え、lines（切ったあとの行）に順に当てる
 */
function mtgPartSpans_(text, parts, width, size, lines) {
  const spans = [];
  let line = 0;
  let offset = 0;
  text.split('\n').forEach((part, k) => {
    const count = mtgWrap_(part, width, size, 99).length;
    const from = Math.min(line, lines.length);
    const to = Math.min(line + count, lines.length);
    const length = lines.slice(from, to).join('\n').length;
    if (to > from && parts[k] && Object.keys(parts[k]).length) spans.push(Object.assign({start: offset, end: offset + length}, parts[k]));
    offset += length + (to > from ? 1 : 0);
    line += count;
  });
  return spans;
}

/** 表の見出しの行（緑に白の太字）。size は文字の大きさ（既定は bodySize） */
function mtgHeadRow_(columns, labels, size) {
  const c = MTG_OPTIONS.colors;
  return mtgRow_(columns, labels.map(text => ({text, size: size || MTG_SLIDE.bodySize, bold: true, color: c.white, fill: c.green, align: 'CENTER',
    lines: 3})), MTG_SLIDE.headRow);
}

/** タイトルを1行に収める（21 → 19 → 17 ポイント。それでも入らなければ「…」で切る） */
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
 * 半角の語（「2026/12」など）の途中では切らず、語の前の空白で折り返す。「前 ▶ 今」は矢印のあとで折り返す。開きかっこは行の終わりに残さない。
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
        // 「前 ▶ 今」は矢印のあとで、半角の語は語の前で改行する。空白の無い数字・英字の並び（「95日」の 95 など）も途中で切らない
        const arrow = hang.test(ch) ? null : /^(.*[▶▸]) (\S.*)$/.exec(line);
        const run = /[0-9A-Za-z]/.test(ch) ? /^(.*[^0-9A-Za-z.,:/%-])([0-9A-Za-z.,:/%-]+)$/.exec(line) : null;
        const word = arrow || (hang.test(ch) ? null : /[!-~]/.test(ch) ? /^(.*\S) ([!-~]+)$/.exec(line) : null) || (run ? [run[0], run[1], run[2]] : null);
        let keep = word ? word[1] : hang.test(ch) && chars.length > 1 ? chars.slice(0, -1).join('') : line;
        let rest = word ? word[2] : line.slice(keep.length);
        // 行の終わりに来ない文字（開きかっこ）は次の行へ
        const open = /[（「『【〔〈《(［\[]$/.exec(keep);
        if (open && Array.from(keep).length > 1) {
          keep = keep.slice(0, -open[0].length);
          rest = open[0] + rest;
        }
        out.push(keep);
        line = rest;
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

/**
 * 最後の行が短い（いちばん長い行の 1/3 未満。「い」だけが残るなど）ときは、行の長さをそろえて折り返す
 * （2行で、空白で分ければどちらも1行に入るなら空白で分ける。ほかは mtgWrap_ と同じ行数のまま、いちばん狭い幅で折り返す）。
 * 「…」で切れるときと、最後の行が十分に長いときは、そのまま
 */
function mtgBalance_(text, widthPt, size, max) {
  const lines = mtgWrap_(text, widthPt, size, max);
  if (lines.length < 2 || /…$/.test(lines[lines.length - 1])) return lines;
  const units = line => Array.from(line).reduce((sum, ch) => sum + mtgCharWidth_(ch), 0);
  if (units(lines[lines.length - 1]) >= Math.max.apply(null, lines.map(units)) / 3) return lines;
  // 2行で、空白で分ければどちらも1行に入るなら、空白で分ける（「10/13 価／格の再提示」ではなく「10/13／価格の再提示」）
  if (lines.length === 2) {
    let best = null;
    Array.from(text).forEach((ch, i, chars) => {
      if (ch !== ' ') return;
      const pair = [chars.slice(0, i).join('').trim(), chars.slice(i + 1).join('').trim()];
      if (!pair[0] || !pair[1] || pair.some(part => mtgWrap_(part, widthPt, size, 99).length > 1)) return;
      const longest = Math.max(units(pair[0]), units(pair[1]));
      if (!best || longest < best.longest) best = {longest, pair};
    });
    if (best) return best.pair;
  }
  let lo = MTG_SLIDE.padX + (size || MTG_SLIDE.bodySize);
  let hi = widthPt;
  for (let k = 0; k < 12; k++) {
    const mid = (lo + hi) / 2;
    if (mtgWrap_(text, mid, size, 99).length > lines.length) lo = mid;
    else hi = mid;
  }
  return mtgWrap_(text, hi, size, lines.length);
}

/** 行の高さ（pt）の見積もり：lines 行・文字の大きさ size（上下の余白を含む） */
function mtgLineHeight_(lines, size) {
  return lines * (size || MTG_SLIDE.bodySize) * MTG_SLIDE.lineFactor + MTG_SLIDE.padY;
}

/* ---------------- スライド（16:9）に書く ---------------- */

/*
 * スライドの部品（位置・大きさは pt。スライドは MTG_SLIDE.width × height）
 * - 箱：{type: 'box', x, y, w, h, text, size, bold, color, font（既定は MTG_SLIDE.font）, fill, alpha（塗りの濃さ 0〜1）, line, lineWeight, align: 'START'|'CENTER'|'END',
 *        valign: 'TOP'|'MIDDLE'|'BOTTOM', shape: 'TEXT_BOX'|'RECTANGLE'|'ROUND_RECTANGLE', spans: [{start, end, bold, color, size}]（一部の文字の書式）}
 * - 表：{type: 'table', x, y, columns: [列の幅], rows: [{h, cells: [{text, size, bold, color, fill, align, spans}]}],
 *        borders: [{row, col, rows, cols, color, weight}]}（区切り線は薄いグレー。borders はそのあとに外枠として付ける）
 *   列の幅は 32pt 以上にする（Slides API は、それより狭い列の幅をエラーにする）。
 * - role（'title'・'lead'・'footer'・'page'・'pill'・'toc'）は部品の役割の目印（スライドには使わない。テスト・確認用）
 * - 表紙・扉の slide.template（{title, subtitle, date, name}）は、ひな形の決まった文字を置き換える文字（mtgSlideRequests_）
 */

/** 色（#7030A0）を Slides API の形に。 */
function mtgRgb_(hex) {
  const n = parseInt(String(hex).replace('#', ''), 16);
  return {red: ((n >> 16) & 255) / 255, green: ((n >> 8) & 255) / 255, blue: (n & 255) / 255};
}

function mtgPt_(value) {
  return {magnitude: Math.round(value * 100) / 100, unit: 'PT'};
}

/** 文字の書式（Slides API の updateTextStyle の style と fields）。scale は大きさの倍率（ひな形のページの大きさに合わせる） */
function mtgTextStyle_(item, scale) {
  return {
    style: {
      fontFamily: item.font || MTG_SLIDE.font,
      fontSize: mtgPt_((item.size || MTG_SLIDE.bodySize) * (scale || 1)),
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
 * options = {scale: 位置・大きさ・文字の倍率（ひな形のページの大きさ ÷ MTG_SLIDE.width。既定 1）, existing: true ならスライドは作らない（ひな形のページ）,
 *            idPrefix: 部品の ID の先頭（既定は pageId）}。
 * slide.template（表紙・扉）は、ひな形のページなら決まった文字（MTG_OPTIONS.templateTexts）を置き換え、ひな形が無ければ同じ位置に描く（mtgTemplateFallback_）
 */
function mtgSlideRequests_(slide, pageId, index, options) {
  const opt = options || {};
  const k = opt.scale || 1;
  const pt = v => mtgPt_(v * k);
  const requests = [];
  let elements = slide.elements;
  if (opt.existing) {
    mtgTemplateReplacements_(slide.template).forEach(([from, to]) => requests.push({replaceAllText: {
      containsText: {text: from, matchCase: true}, replaceText: to, pageObjectIds: [pageId]}}));
  } else {
    requests.push({createSlide: {objectId: pageId, insertionIndex: index, slideLayoutReference: {predefinedLayout: 'BLANK'}}});
    if (slide.template) elements = mtgTemplateFallback_(slide.template).concat(elements);
  }
  const styleText = (target, item, text) => {
    const font = mtgTextStyle_(item, k);
    const para = mtgParagraphStyle_(item);
    requests.push({insertText: Object.assign({text, insertionIndex: 0}, target)});
    requests.push({updateTextStyle: Object.assign({textRange: {type: 'ALL'}, style: font.style, fields: font.fields}, target)});
    requests.push({updateParagraphStyle: Object.assign({textRange: {type: 'ALL'}, style: para.style, fields: para.fields}, target)});
    (item.spans || []).forEach(span => {
      const end = Math.min(span.end, text.length);
      if (!(end > span.start)) return;
      const style = {};
      const fields = [];
      if ('bold' in span) { style.bold = !!span.bold; fields.push('bold'); }
      if (span.color) { style.foregroundColor = {opaqueColor: {rgbColor: mtgRgb_(span.color)}}; fields.push('foregroundColor'); }
      if (span.size) { style.fontSize = pt(span.size); fields.push('fontSize'); }
      if (fields.length) {
        requests.push({updateTextStyle: Object.assign({textRange: {type: 'FIXED_RANGE', startIndex: span.start, endIndex: end}, style,
          fields: fields.join(',')}, target)});
      }
    });
  };
  const fill = (color, alpha) => (color
    ? {solidFill: Object.assign({color: {rgbColor: mtgRgb_(color)}}, alpha != null ? {alpha} : {})}
    : {propertyState: 'NOT_RENDERED'});
  elements.forEach((el, n) => {
    const id = (opt.idPrefix || pageId) + '_' + n;
    const width = el.type === 'table' ? el.columns.reduce((a, b) => a + b, 0) : el.w;
    const height = el.type === 'table' ? el.rows.reduce((sum, r) => sum + r.h, 0) : el.h;
    const elementProperties = {
      pageObjectId: pageId,
      size: {width: pt(width), height: pt(height)},
      transform: {scaleX: 1, scaleY: 1, translateX: Math.round(el.x * k * 100) / 100, translateY: Math.round(el.y * k * 100) / 100, unit: 'PT'}
    };
    if (el.type === 'table') {
      requests.push({createTable: {objectId: id, elementProperties, rows: el.rows.length, columns: el.columns.length}});
      el.columns.forEach((w, c) => requests.push({updateTableColumnProperties: {objectId: id, columnIndices: [c],
        tableColumnProperties: {columnWidth: pt(w)}, fields: 'columnWidth'}}));
      el.rows.forEach((row, r) => {
        requests.push({updateTableRowProperties: {objectId: id, rowIndices: [r], tableRowProperties: {minRowHeight: pt(row.h)}, fields: 'minRowHeight'}});
        row.cells.forEach((cell, c) => {
          const location = {rowIndex: r, columnIndex: c};
          requests.push({updateTableCellProperties: {objectId: id, tableRange: {location, rowSpan: 1, columnSpan: 1},
            tableCellProperties: {tableCellBackgroundFill: fill(cell.fill), contentAlignment: 'MIDDLE'},
            fields: 'tableCellBackgroundFill,contentAlignment'}});
          styleText({objectId: id, cellLocation: location}, cell, cell.text == null || cell.text === '' ? ' ' : String(cell.text));
        });
      });
      const all = {location: {rowIndex: 0, columnIndex: 0}, rowSpan: el.rows.length, columnSpan: el.columns.length};
      [{range: all, position: 'ALL', color: MTG_OPTIONS.colors.line, weight: 0.75}].concat((el.borders || []).map(b => ({
        range: {location: {rowIndex: b.row, columnIndex: b.col}, rowSpan: b.rows, columnSpan: b.cols}, position: 'OUTER',
        color: b.color, weight: b.weight || 2.25
      }))).forEach(b => requests.push({updateTableBorderProperties: {objectId: id, tableRange: b.range, borderPosition: b.position,
        tableBorderProperties: {tableBorderFill: {solidFill: {color: {rgbColor: mtgRgb_(b.color)}}}, weight: pt(b.weight), dashStyle: 'SOLID'},
        fields: 'tableBorderFill,weight,dashStyle'}}));
      return;
    }
    requests.push({createShape: {objectId: id, shapeType: el.shape || 'TEXT_BOX', elementProperties}});
    requests.push({updateShapeProperties: {objectId: id, shapeProperties: {
      shapeBackgroundFill: fill(el.fill, el.alpha),
      outline: el.line
        ? {outlineFill: {solidFill: {color: {rgbColor: mtgRgb_(el.line)}}}, weight: pt(el.lineWeight || 2.25), dashStyle: 'SOLID'}
        : {propertyState: 'NOT_RENDERED'},
      contentAlignment: el.valign || 'MIDDLE'
    }, fields: 'shapeBackgroundFill,outline,contentAlignment'}});
    if (el.text != null && el.text !== '') styleText({objectId: id}, el, String(el.text));
  });
  return requests;
}

/**
 * ひな形の決まった文字の置き換え [[ひな形の文字, 置き換える文字]]（ひな形の文字の長い順。「XXXXXX」を「XXXXX」より先に置き換えるため）。
 * template = {title, subtitle, date, name}
 */
function mtgTemplateReplacements_(template) {
  const texts = MTG_OPTIONS.templateTexts;
  return Object.keys(texts).filter(key => texts[key] && template && key in template)
    .sort((a, b) => texts[b].length - texts[a].length)
    .map(key => [texts[key], String(template[key] == null ? '' : template[key])]);
}

/**
 * ひな形が無いときの表紙・扉：緑の背景に、ひな形と同じ位置（ひな形の 960 × 540pt を 0.75 倍）で見出し・その下の文字・日付・名前を白で描く
 */
function mtgTemplateFallback_(template) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const line = (y, h, text, size, bold) => mtgBox_(S.x, y, S.w, h, mtgWrap_(text, S.w, size, 1)[0], {size, bold: !!bold, color: c.white, align: 'CENTER'});
  return [
    {type: 'box', x: 0, y: 0, w: S.width, h: S.height, fill: c.cover, shape: 'RECTANGLE'},
    line(170, 46, template.title, 30, true),
    line(226, 26, template.subtitle, 15, true),
    line(326, 20, template.date, 12),
    line(344, 20, template.name, 12)
  ].filter(el => el.fill || el.text);
}

/** 表紙・扉のひな形の ID（MTG_OPTIONS.templateId か、setMeetingTemplate() で登録したもの。無ければ空） */
function mtgTemplateId_() {
  return MTG_OPTIONS.templateId || PropertiesService.getDocumentProperties().getProperty(MTG_OPTIONS.templateKey) || '';
}

/** 「https://docs.google.com/presentation/d/ID/edit」や ID だけの文字から ID を取り出す（読めなければ空） */
function mtgFileIdOf_(text) {
  const t = String(text || '').trim();
  const m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(t) || /^([A-Za-z0-9_-]{20,})$/.exec(t);
  return m ? m[1] : '';
}

/**
 * 表紙・扉のひな形（Google スライド）を登録する（URL を聞く。空欄なら登録を外し、ひな形なしで作る）。
 * いただいた .pptx は、ドライブにアップロードして Google スライドで開き、「ファイル」→「Google スライドとして保存」したものを登録する
 */
function setMeetingTemplate() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const current = props.getProperty(MTG_OPTIONS.templateKey);
  const response = ui.prompt(CRED_OPTIONS.menuTitle,
    '会議サマリの表紙・扉のひな形（Google スライド）の URL を入力してください。\n空欄なら登録を外します（ひな形なしで作ります）。\n\n' +
    '今の登録：' + (current ? 'https://docs.google.com/presentation/d/' + current + '/edit' : 'なし'), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const text = String(response.getResponseText() || '').trim();
  if (!text) {
    props.deleteProperty(MTG_OPTIONS.templateKey);
    ui.alert(CRED_OPTIONS.menuTitle, 'ひな形の登録を外しました。会議サマリは、ひな形なしで作ります。', ui.ButtonSet.OK);
    return;
  }
  const id = mtgFileIdOf_(text);
  if (!id) throw new Error('「' + text + '」から Google スライドの ID を読めません。Google スライドを開いたときのアドレス（URL）を貼り付けてください。');
  mtgRequireSlides_();
  mtgTemplateCheck_(id);
  props.setProperty(MTG_OPTIONS.templateKey, id);
  const deck = Slides.Presentations.get(id);
  const message = 'ひな形を登録しました（' + (deck.slides || []).length + '枚。1枚目を表紙・扉に使います）。次に作る会議サマリから使います。';
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** ひな形が Google スライドか確かめる（.pptx のままなら、Google スライドにしてもらう） */
function mtgTemplateCheck_(id) {
  let file;
  try {
    file = DriveApp.getFileById(id);
  } catch (e) {
    throw new Error('ひな形（ID：' + id + '）を開けません。ファイルがあるか、この管理者のアカウントで開けるかを確かめてください（' + e.message + '）。');
  }
  if (file.getMimeType() !== MimeType.GOOGLE_SLIDES) {
    throw new Error('ひな形が Google スライドではありません（.pptx のままなど）。ドライブで開いて「ファイル」→「Google スライドとして保存」し、' +
      'できた Google スライドを setMeetingTemplate() で登録してください。');
  }
  return file;
}

/** 資料のページの幅（pt）。Slides API の pageSize（EMU か PT） */
function mtgPageWidth_(deck) {
  const w = deck.pageSize && deck.pageSize.width;
  if (!w) return MTG_SLIDE.width;
  return w.unit === 'EMU' ? w.magnitude / 12700 : w.magnitude;
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
 * 表紙・扉のひな形（mtgTemplateId_）があれば、ひな形をコピーした資料に作る：1枚目（表紙のひな形）の決まった文字を置き換えて表紙にし、
 * 1枚目を複製して扉にする（2枚目以降は消す）。ほかのスライドは白紙のスライドを足して描き、位置・大きさ・文字はひな形のページの大きさに合わせて
 * 拡大する（MTG_SLIDE の 720 × 405pt を基準に描いているため）。ひな形が無ければ、新しい資料（720 × 405pt）に作る。
 * スライドは Slides API（拡張サービス「Google Slides API」）で作り、できたスライドを PDF に書き出す。
 * MTG_OPTIONS.keepSlides なら Google スライドもフォルダに残す（推進リーダーが書き足せる）。うまくいかなかったときはゴミ箱に移す。
 * 戻り値 {pdf, deck}（deck は残さないときは null）
 */
function mtgExportDeck_(name, slides, folder) {
  const templateId = mtgTemplateId_();
  let file;
  let deck;
  if (templateId) {
    file = mtgTemplateCheck_(templateId).makeCopy(name, folder);
    deck = Slides.Presentations.get(file.getId());
  } else {
    deck = Slides.Presentations.create({title: name, pageSize: {width: mtgPt_(MTG_SLIDE.width), height: mtgPt_(MTG_SLIDE.height)}});
    file = DriveApp.getFileById(deck.presentationId);
  }
  try {
    const scale = mtgPageWidth_(deck) / MTG_SLIDE.width;
    const pages = (deck.slides || []).map(sl => sl.objectId);
    const cover = templateId ? pages[0] : null;
    // ひな形：2枚目以降を消し、1枚目を複製して扉にする（表紙の次に入る。本編を表紙の次から順に入れると、扉は本編のあとになる）
    let requests = [];
    if (cover) {
      pages.slice(1).forEach(id => requests.push({deleteObject: {objectId: id}}));
      if (slides.some(sl => sl.divider)) requests.push({duplicateObject: {objectId: cover, objectIds: {[cover]: 'mtg_divider'}}});
    }
    let first = true;
    const flush = () => {
      if (first && !cover) pages.forEach(id => requests.push({deleteObject: {objectId: id}}));   // 作ったときの白紙のスライド
      Slides.Presentations.batchUpdate({requests}, deck.presentationId);
      requests = [];
      first = false;
    };
    slides.forEach((slide, i) => {
      const own = 'mtg_p' + (i + 1);
      const add = cover && slide.cover ? mtgSlideRequests_(slide, cover, i, {scale, existing: true, idPrefix: own})
        : cover && slide.divider ? mtgSlideRequests_(slide, 'mtg_divider', i, {scale, existing: true, idPrefix: own})
          : mtgSlideRequests_(slide, own, i, {scale});
      if (requests.length && requests.length + add.length > MTG_OPTIONS.batchRequests) flush();
      requests = requests.concat(add);
    });
    flush();
    const extra = mtgPlaceholderRequests_(Slides.Presentations.get(deck.presentationId), slides.map((slide, i) => 'mtg_p' + (i + 1)));
    if (extra.length) Slides.Presentations.batchUpdate({requests: extra}, deck.presentationId);
    if (!templateId) file.moveTo(folder);
    const pdf = folder.createFile(file.getAs('application/pdf').setName(name + '.pdf'));
    mtgTrashOld_(folder, [name, name + '.pdf'], [file.getId(), pdf.getId()]);
    if (!MTG_OPTIONS.keepSlides) file.setTrashed(true);
    return {pdf, deck: MTG_OPTIONS.keepSlides ? file : null};
  } catch (e) {
    try { file.setTrashed(true); } catch (_) {}
    throw e;
  }
}

/**
 * 足したスライド（pageIds）に付いたレイアウトのプレースホルダ（ページ番号・日付・フッターの枠）を消すリクエスト。
 * ひな形のレイアウトにはページ番号の枠があり、白紙のスライドにも付くと、mtgFrame_ のページ番号と2つ出るため。表紙・扉（ひな形のページ）はそのまま
 */
function mtgPlaceholderRequests_(deck, pageIds) {
  const own = new Set(pageIds);
  const requests = [];
  (deck.slides || []).filter(sl => own.has(sl.objectId)).forEach(sl => (sl.pageElements || [])
    .filter(el => el.shape && el.shape.placeholder).forEach(el => requests.push({deleteObject: {objectId: el.objectId}})));
  return requests;
}

/**
 * 保存先のフォルダの、同じ名前の古いファイル（作り直す前の版）をゴミ箱に移す（古い版を配らないように）。keep の ID は残す。移した数を返す。
 * 移せないファイル（ほかの人のファイルなど）はログに残して続ける（作った資料は残す）
 */
function mtgTrashOld_(folder, names, keep) {
  let count = 0;
  names.forEach(name => {
    const files = folder.getFilesByName(name);
    while (files.hasNext()) {
      const file = files.next();
      if (keep.indexOf(file.getId()) >= 0) continue;
      try {
        file.setTrashed(true);
        count++;
      } catch (error) {
        console.warn('古い版をゴミ箱に移せませんでした（' + name + '）：' + (error && error.message || error));
      }
    }
  });
  return count;
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
