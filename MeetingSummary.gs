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
 *   表紙と扉は、いただいた表紙のひな形（Google スライド。setMeetingTemplate() で登録）をそのまま使う。登録が無ければ、ひな形と同じ絵柄
 *   （緑のグラデーションに大きな「RX」）の画像をスクリプトで描いて背景にし、同じ位置に文字を描く（mtgCoverPng_）
 *   表紙：会議の名前と「事前資料」・サブインダストリー・会議日・データの時点・期間
 *   1 サマリー：期待値と進行中案件の数・前回以降の動き（案件ごとに1回だけ数える区分の件数と案件）・現在の状態（受注月超過・停滞）
 *   2 パイプラインの変化：前回からの期待値の増減（新規・増額・復帰・減額・受注・失注・保留・削除）・ステータス別・受注月（四半期）別の期待値
 *   3 前回から変化のある案件：ステータス・期待値・受注月・活動状況が変わった案件（区分の順に上位）
 *   4 期待値上位案件の状況・課題：進行中の案件の期待値の大きい順（ステータスの滞留日数・受注月・活動状況・課題・該当事項）
 *   扉（Appendix。Appendix の見出しとページ）
 *   Appendix（全件。入りきらなければ続きのページ。中身が無くても ①〜③ を出す）
 *   作れないところ（シートに列が無い・変更の記録が無い・停滞をまだ数えられないなど）は、資料の型を変えずに、その場所に灰色の「データなし」と理由（mtgGaps_）
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
 * - スライドは Slides API（拡張サービス「Google Slides API」。appsscript.json に書いてある）で作り、PDF に書き出す
 *   （ひな形が無いときの表紙・扉の背景の画像だけは SlidesApp で付ける）。
 *   表紙・扉のひな形を登録してあれば（setMeetingTemplate()）、ひな形をコピーした資料に作り、ひな形の表紙の決まった文字を置き換える。
 */

const MTG_OPTIONS = {
  meetingName: 'RX Business MTG',    // 表紙に出す会議の名前
  // 空欄・該当なしの書き方。資料には推定で値を補わず、空欄は blankText と書く（「—」で済ませない）。該当する案件が無いときは noneText。
  // シートに列が無い・記録が無いなどで作れない（数えられない）ところは gapText と理由を、灰色の枠・セルで書く（資料の型は変えない）
  blankText: '未入力',
  noneText: '該当なし',
  gapText: 'データなし',
  fileName: '会議サマリ',             // ファイル名の先頭（会議サマリ_スーパー・GMS_2026-10-23.pdf。日付は会議日）
  folderId: '1A0rkoFLICOK4KV25dUTeRdgoOaM0vJEK',   // 保存先のフォルダの ID（空なら、このスプレッドシートと同じフォルダの中に folderName を作る）
  folderName: '会議サマリ',
  folderKey: 'MTG_FOLDER_ID',         // 作ったフォルダの ID（ドキュメントのプロパティ）
  // 会議時点のコピー：本番の資料を作ったとき・基準を記録したときに、追跡シートの値を別のスプレッドシート（タイトルは実行日時）に保存する。
  // 次の資料の「前回からの変化」は、前回のコピーと今のシートを比べて出す（差分追跡が止まっていた間・2,000セルを超える貼り付けの変更も出る）。
  // 保存先は、会議サマリのフォルダの中の copyFolderName のフォルダ
  copyName: '会議時点のコピー',       // タイトル（会議時点のコピー_2026-10-16 18:30（会議日 2026-10-23））
  copyFolderName: '会議時点のコピー',
  copyFolderKey: 'MTG_COPY_FOLDER_ID', // 作ったフォルダの ID（ドキュメントのプロパティ）
  copyChunkRows: 500,                 // コピーするときの1回の行数
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
    urlKey: 'MTG_CHAT_URL',          // スクリプトのプロパティ：送る資料の格納先の URL（ドライブのフォルダかファイル）
    meetingTime: '9:00〜10:30',      // 会議の時刻（会議日程の日付のあとに書く。送るときに時刻を入力すれば、そちら。空なら日付だけ）
    // 送る文（1行ずつ。目的・議題もここで変える）。{meeting} は会議の名前（meetingName）、{date} は会議日程（2026年10月23日（金）9:00〜10:30）、
    // {file} は資料のファイルのリンク（ファイル名で表示）、{url} はファイルの URL に置き換える。
    // 格納先がフォルダなら、送るときにその中でいちばん新しく更新した PDF（PDF が無ければいちばん新しいファイル）を送る。*〜* は太字
    message: [
      '*【{meeting}】事前資料のご案内*',
      '{meeting} の事前資料を、下記に格納いたしました。',
      'お手数ですが、会議の前にご一読いただけますと幸いです。',
      '',
      '▼ 会議日程',
      '{date}',
      '',
      '▼ 目的',
      '案件の進捗・受注見込みを確認し、優先順位と次のアクションを決める',
      '',
      '▼ 議題',
      '・更新シートで進捗報告',
      '・重点案件の状況・課題',
      '・質疑・フィードバック',
      '・サービス連携テーマ',
      '',
      '▼ 事前資料',
      '{file}'
    ],
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
  required: ['案件ステータス', '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月', '活動状況', 'アカウント責任者'],
  // ステータスの定義のシート（期待値の率）。ステータスの進み具合（前進・後退）は率の大きい順、率が同じなら定義の下の行ほど進んでいる
  // （C-提案中 → C-選定対応中 → C-稟議対応中）。期待値の列が数式（想定売上 × 率）のとき、前回の期待値を記録から取れない案件はこの率で計算する。
  // シートが無ければ、進み具合はプルダウンの順（下ほど進んでいる）
  statusTable: {sheet: '定義_ステータス', headerRow: 1, status: 'ステータス', ratio: '期待値'},
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
    gap: '#F2F2F2', gapLine: '#A6A6A6',   // データなし（作れなかったところ）の塗りと点線の枠
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
/* 資料に出す種類の名前を2行に分けるときの区切り（言葉の途中で切らない。責任者別の表の見出し・一覧で1行に入らないとき） */
const MTG_ALERT_SHORT = {'予定日超過': '予定日\n超過', '受注月超過': '受注月\n超過', '未入力': '必須項目\n未入力', '資料なし': '商談記録\n未登録',
  '更新なし': '長期\n未更新', '金額の不整合': '金額\n不整合', '月の不整合': '売上開始月\n不整合', '選択肢外のステータス': '選択肢外\nステータス',
  'デリバリー担当なし': 'デリバリー\n担当未設定'};
const MTG_ALERT_LABELS = {'予定日超過': '予定日超過', '受注月超過': '受注月超過', '未入力': '必須項目未入力', '資料なし': '商談記録未登録',
  '更新なし': '長期未更新', '金額の不整合': '金額不整合', '月の不整合': '売上開始月不整合', '選択肢外のステータス': '選択肢外ステータス',
  'デリバリー担当なし': 'デリバリー担当未設定'};

/*
 * スライドの形（pt。16:9 = 720 × 405）。上から 見出し（titleY）・リード文（leadY）・中身（top〜bottom）・出典の行とページ番号（footY）。
 * 表紙と扉は、見出し・リード文の代わりに大きな文字を置く（表紙には出典の行・ページ番号も置かない）。
 * 文字の幅と行の高さは見積もり（mtgWrap_・mtgTextHeight_）。Google スライドで書き出した PDF で測った値に合わせ、少し大きめに見積もって、はみ出さないようにする
 * （表のセルの余白は上下左右とも 7.2pt で、API では変えられない。行の送りは文字の大きさ × 1.2。表の行は、中身が入りきらなければ伸びる）。
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
  lineFactor: 1.2,       // 行の送り = 文字の大きさ × lineFactor（行間は「1」。Google スライドで測った値）
  padX: 14.4,            // 箱・表のセルの左右の余白（合わせて。7.2pt ずつ）
  padY: 14.4,            // 上下の余白（合わせて。7.2pt ずつ）。表の行の高さ = 行の送りの合計 + padY
  fit: 0.95,             // 文字の幅の見積もりの余裕（セルの幅の 95% までに収める）
  headRow: 20,           // 表の見出しの行の高さの下限
  // ラベル（区分・当てはまる状況）：高さ・文字の大きさ・縦に並べるときの間・セルの上下の余白。幅は列ごとに同じ（mtgPillWidth_）
  pill: {h: 12, size: 8, gap: 2, margin: 4}
};

/*
 * 半角の英数字・記号（' '〜'~'）の幅（文字の大きさに対する割合。BIZ UDPゴシックの標準と太字の大きいほうを切り上げたもの）。
 * i 番目が文字コード 32 + i。数字は 0.76（1 だけ 0.63）で、全角の 3/4 ほどある（0.6 と見積もると「2026/12」が折り返される）
 */
const MTG_ASCII_WIDTHS = [
  0.34, 0.38, 0.52, 0.94, 0.8, 1, 0.92, 0.32, 0.51, 0.51, 0.5, 0.84, 0.34, 0.48, 0.34, 0.5,                // ' ' 〜 '/'
  0.76, 0.63, 0.76, 0.76, 0.76, 0.76, 0.76, 0.76, 0.76, 0.76, 0.34, 0.34, 0.84, 0.84, 0.84, 0.8,          // '0' 〜 '?'
  0.98, 0.78, 0.77, 0.82, 0.86, 0.71, 0.67, 0.86, 0.84, 0.44, 0.57, 0.76, 0.61, 1, 0.84, 0.88,            // '@' 〜 'O'
  0.76, 0.9, 0.76, 0.73, 0.71, 0.86, 0.76, 1, 0.82, 0.75, 0.75, 0.55, 0.5, 0.55, 0.5, 0.31,               // 'P' 〜 '_'
  0.5, 0.67, 0.73, 0.65, 0.73, 0.69, 0.51, 0.67, 0.73, 0.32, 0.4, 0.67, 0.36, 0.96, 0.73, 0.71,           // '`' 〜 'o'
  0.73, 0.73, 0.5, 0.59, 0.5, 0.73, 0.61, 0.9, 0.63, 0.65, 0.59, 0.54, 0.44, 0.54, 0.5                    // 'p' 〜 '~'
];

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
  // スプレッドシートの画面に出せないとき（エディタをスプレッドシートから開いていないなど）は、聞かずにすべての追跡シートに記録し、実行ログに書く
  const ui = mtgUiOrNull_();
  const names = ui ? mtgAskSheets_(ss, ui, '会議サマリの基準を記録します（会議日は記録しません）。')
    : Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  if (!names) return;
  if (!names.length) throw new Error('追跡しているシート（test.gs の DIFF_RULES）が見つかりません。');
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const at = Utilities.formatDate(now, tz, 'MM/dd HH:mm');
  const answer = !ui ? null : ui.alert(CRED_OPTIONS.menuTitle,
    '「' + names.join('」「') + '」に、今（' + at + '）を基準として記録します。\n' +
    '次に作る会議サマリは、基準より後の変化だけを出します（今までに貼り付けた案件を「新規」に数えません）。\n' +
    '今の値は「' + MTG_OPTIONS.copyName + '」（別のスプレッドシート）に保存し、次の資料はそれと比べます。\n' +
    '移行の貼り付けがすべて終わってから実行してください。記録してよろしいですか？', ui.ButtonSet.YES_NO);
  if (ui && answer !== ui.Button.YES) return;
  const props = PropertiesService.getDocumentProperties();
  const date = Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  const copy = mtgSaveCopy_(ss, names, now, '（基準）');
  names.forEach(name => {
    const list = mtgMeetings_(props, name).filter(m => !m.baseline)
      .concat([{date, cutoff: now.getTime(), baseline: true, copy: {id: copy.id, sheet: name}}])
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).slice(-MTG_OPTIONS.keepMeetings);
    props.setProperty(MTG_OPTIONS.meetingsKeyPrefix + name, JSON.stringify(list));
  });
  const message = '「' + names.join('」「') + '」に、基準（' + at + '）を記録しました。次に作る会議サマリは、基準より後の変化を出します。\n' +
    '今の値を「' + copy.title + '」に保存しました：' + copy.url +
    (ui ? '' : '（スプレッドシートの画面に出せないため、確かめずにすべての追跡シートに記録しました）');
  console.log(message);
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/**
 * 記録した会議日を取り消す（会議日を打ち間違えたとき・本番を作り直す前に戻すとき。エディタから実行する）。
 * シートと会議日を聞き、その会議日の記録を消す。基準の記録と、保存した資料のファイルは消さない
 */
function removeMeetingRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const names = mtgAskSheets_(ss, ui, '会議サマリの会議日の記録を取り消します（資料のファイル・会議時点のコピーは消しません）。');
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
    const dummy = smpCountEvents_(SpreadsheetApp.getActiveSpreadsheet(), SMP_OPTIONS.dummy.eventPrefix + '-', name).changes;
    const previous = base && (!held || base.cutoff > held.cutoff) ? base : held;
    return '・' + name + '：' + from + (previous && previous.copy ? '（その時点のコピーと比べる）' : '') +
      (same ? '\n　※ この会議日の資料は ' + fmt(same.cutoff) + ' に作成済み。作り直すと、次の会議の起点が今になります。' : '') +
      (dummy ? '\n　※ ダミーの変更の記録（' + dummy + '件）があります。資料に【ダミーデータを含む】と出ます。配る資料なら、先に removeMeetingDummyHistory() で消してください。' : '');
  });
  const answer = ui.alert(CRED_OPTIONS.menuTitle,
    '本番の資料を作り、会議日を記録します。\n\n会議日：' + p[0] + '/' + p[1] + '/' + p[2] + '（' + week + '）\n対象期間\n' + lines.join('\n') +
    '\n\n今の値は「' + MTG_OPTIONS.copyName + '」（別のスプレッドシート。タイトルは今の日時）に保存し、次の会議の資料はそれと比べます。' +
    '\n記録した会議日は、次の会議の資料の起点になります。配布したあとは作り直さないでください。\n作成してよろしいですか？', ui.ButtonSet.YES_NO);
  return answer === ui.Button.YES;
}

function mtgRun_(final) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const names = mtgAskSheets_(ss, ui, final ? '会議サマリを作ります。' : '会議サマリを試しに作ります（会議日は記録しません）。');
  if (!names) return;
  const now = new Date();
  const meeting = mtgAskMeeting_(ui, names, now);
  if (!meeting) return;
  if (final && !mtgConfirmFinal_(ui, names, meeting)) return;
  const results = mtgCreate_(ss, names, now, final, meeting);
  // 作れなかったところ（データなし）があれば、シートごとに理由を出す（資料にも、その場所に灰色の枠で書いてある）
  const lines = results.map(r => {
    const gaps = mtgGapList_(r.summary);
    // 変化が無く、サンプル（addSampleHistory()。会議サマリは数えない）の記録だけがあるときは、資料の確認用の関数を案内する
    const samples = !r.summary.kpi.changed && !r.summary.dummy ? smpCountEvents_(ss, SMP_OPTIONS.eventPrefix + '-', r.summary.name).changes : 0;
    return '「' + r.summary.name + '」：' + r.file.getName() + '（' + mtgCountsText_(r.summary) + '）\n　' + r.file.getUrl() +
      (r.summary.copy ? '\n　前回からの変化は「' + r.summary.copy.title + '」と比べました。' : '') +
      (r.summary.copyError ? '\n　※ ' + r.summary.copyError : '') +
      (r.summary.unlogged ? '\n　※ 変更の記録が無い変更 ' + r.summary.unlogged + '件（差分追跡が止まっていた間などの変更。資料には出しています）。' +
        'diagnoseDiffState() で差分追跡が動いているか確かめてください。' : '') +
      (gaps.length ? '\n　' + MTG_OPTIONS.gapText + '：' + gaps.join('\n　　') : '') +
      (r.summary.dummy ? '\n　※ ダミーの変更の記録（' + r.summary.dummy + '件）を含みます（資料に【ダミーデータを含む】）。消すときは removeMeetingDummyHistory()。' : '') +
      (samples ? '\n　※ サンプル（addSampleHistory()）の変更の記録 ' + samples + '件は、会議サマリには出ません（基準を変えても出ません）。' +
        '資料の形を確かめるときは、addMeetingDummyHistory() でダミーの変更の記録を足してください。' : '');
  });
  lines.push('', '保存先：' + results[0].folder.getUrl());
  if (results[0].copy) lines.push('今の値を「' + results[0].copy.title + '」に保存しました（次の会議の資料はこれと比べます）：' + results[0].copy.url);
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
  const ui = mtgUi_();
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
    '送る資料の格納先の URL を貼り付けてください（ドライブのフォルダかファイル。参加者が開けるように共有しておく）。\n' +
    'フォルダなら、送るときに、その中でいちばん新しく更新した PDF のリンクを送ります。' +
    (current && current.url ? '\n空欄なら今の URL のまま：' + current.url : ''), ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const url = String(answer.getResponseText() || '').trim() || (current && current.url) || '';
  if (!/^https:\/\/\S+$/.test(url)) throw new Error('資料の格納先の URL を読めません。https:// で始まる URL を貼り付けてください。');
  props.setProperty(o.webhookKey, webhook);
  props.setProperty(o.urlKey, url);
  const removed = mtgRemoveLegacyTrigger_();
  const lines = ['Google チャットへの通知を設定しました。', '・送り先：' + mtgChatSpace_(webhook), '・送る URL：' + url,
    '・notifyMeetingSummaries() を実行すると、資料のファイルのリンクをすぐに送ります（フォルダなら、その中でいちばん新しい PDF）。']
    .concat(removed ? ['・前の版の自動通知（2日前）のトリガーを外しました。'] : []);
  const test = ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n') + '\n\n確かめのメッセージを、いまチャットに送りますか？', ui.ButtonSet.YES_NO);
  if (test === ui.Button.YES) {
    mtgPostChat_(webhook, '【テスト送信】会議サマリの通知の設定を確認するためのメッセージです。\n今後、' + MTG_OPTIONS.meetingName +
      ' の事前資料のご案内をこちらにお送りいたします。');
    lines.push('確かめのメッセージを送りました。チャットに届いたか確かめてください。');
  }
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/**
 * 会議日程と、資料のファイルのリンクを、すぐに Google チャットに送る（エディタから実行する。「はい」のときだけ）。
 * 会議日程（日付。時刻は省けば MTG_OPTIONS.chat.meetingTime）を聞く。格納先がフォルダなら、その中でいちばん新しく更新した PDF（mtgChatFile_）
 */
function notifyMeetingSummaries() {
  const chat = mtgChatSettings_();
  if (!chat || !chat.url) throw new Error('Google チャットへの通知が設定されていません。先に setMeetingChat() で webhook の URL と資料の格納先の URL を登録してください。');
  mtgRemoveLegacyTrigger_();
  const ui = mtgUi_();
  const file = mtgChatFile_(chat.url);   // 先に確かめる（資料が無ければ、会議日程を聞かずに知らせる）
  const time = MTG_OPTIONS.chat.meetingTime;
  const response = ui.prompt(CRED_OPTIONS.menuTitle,
    '会議日程を入力してください（例：10/23）。\n' +
    (time ? '時刻を省くと「' + time + '」にします。ほかの時刻のときは「10/23 13:00〜14:30」のように入力してください。' : '時刻を書くときは「10/23 13:00〜14:30」のように入力してください。'),
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const schedule = mtgChatSchedule_(response.getResponseText(), new Date());
  const answer = ui.alert(CRED_OPTIONS.menuTitle,
    'Google チャット（' + mtgChatSpace_(chat.webhook) + '）に、次のメッセージをすぐに送ります。よろしいですか？\n\n' + mtgChatText_(file, schedule, true),
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  mtgPostChat_(chat.webhook, mtgChatText_(file, schedule));
  const message = 'Google チャットに送りました：' + file.name + '\n' + file.url;
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/**
 * 送る資料のファイル {name, url}。格納先がドライブのフォルダなら、その中でいちばん新しく更新した PDF（PDF が無ければいちばん新しいファイル。ゴミ箱のものは除く）。
 * ドライブのファイルの URL ならそのファイル。ドライブ以外の URL なら、名前は URL のまま
 */
function mtgChatFile_(url) {
  const folderId = /\/folders\/([\w-]{10,})/.exec(url);
  if (folderId) {
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId[1]);
    } catch (error) {
      throw new Error('資料の格納先のフォルダ（ID：' + folderId[1] + '）を開けません。このアカウントで開けるか確かめ、setMeetingChat() で登録し直してください。');
    }
    const newest = files => {
      let best = null;
      while (files.hasNext()) {
        const f = files.next();
        if (!f.isTrashed() && (!best || f.getLastUpdated().getTime() > best.getLastUpdated().getTime())) best = f;
      }
      return best;
    };
    const file = newest(folder.getFilesByType(MimeType.PDF)) || newest(folder.getFiles());
    if (!file) throw new Error('資料の格納先のフォルダ（' + folder.getName() + '）にファイルがありません。資料を置いてから送ってください。');
    return {name: file.getName(), url: file.getUrl()};
  }
  const fileId = /\/d\/([\w-]{10,})/.exec(url) || /[?&]id=([\w-]{10,})/.exec(url);
  if (fileId) {
    try {
      const file = DriveApp.getFileById(fileId[1]);
      return {name: file.getName(), url: file.getUrl()};
    } catch (error) {
      throw new Error('資料のファイル（ID：' + fileId[1] + '）を開けません。このアカウントで開けるか確かめ、setMeetingChat() で登録し直してください。');
    }
  }
  return {name: url, url};
}

/**
 * 入力した会議日程（「10/23」「2026/10/23」、時刻つきなら「10/23 13:00〜14:30」）を「2026年10月23日（金）9:00〜10:30」にする。
 * 時刻を省けば MTG_OPTIONS.chat.meetingTime。読めなければ止める
 */
function mtgChatSchedule_(text, now) {
  const t = String(text || '').normalize('NFKC').trim();
  const at = /\s*(\d{1,2}:\d{2}.*)$/.exec(t);
  const date = mtgParseDate_(at ? t.slice(0, at.index) : t, now);
  if (!date) {
    throw new Error('「' + t + '」を会議日程として読めません。10/23 か 2026/10/23（時刻を書くときは 10/23 13:00〜14:30）のように入力してください。');
  }
  return mtgDateLabel_(date) + (at ? at[1].replace(/~/g, '〜') : MTG_OPTIONS.chat.meetingTime || '');
}

/** 日付（yyyy/MM/dd）を「2026年10月23日（金）」に（表紙の会議日・チャットの会議日程） */
function mtgDateLabel_(date) {
  const p = date.split('/').map(Number);
  return p[0] + '年' + p[1] + '月' + p[2] + '日（' + '日月火水木金土'.charAt(new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay()) + '）';
}

/**
 * チャットに送る文（MTG_OPTIONS.chat.message の {meeting} を会議の名前に、{date} を会議日程に、{file} をファイルのリンク（<URL|ファイル名>）に、
 * {url} をファイルの URL に置き換える）。plain なら、確かめの画面に出す形（リンクは「ファイル名」と URL の2行）
 */
function mtgChatText_(file, schedule, plain) {
  const name = String(file.name).replace(/[|<>]/g, ' ');
  const link = plain ? name + '\n' + file.url : '<' + file.url + '|' + name + '>';
  return MTG_OPTIONS.chat.message.join('\n').replace(/\{meeting\}/g, MTG_OPTIONS.meetingName).replace(/\{date\}/g, schedule)
    .replace(/\{file\}/g, link).replace(/\{url\}/g, file.url);
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

/**
 * 入力・確かめの画面を出す先（スプレッドシートの画面）。出せなければ、どうすればよいかを書いて止める
 * （エディタをスプレッドシートの「拡張機能」→「Apps Script」から開いていない、スプレッドシートを閉じているなど）
 */
function mtgUi_() {
  const ui = mtgUiOrNull_();
  if (!ui) {
    throw new Error('スプレッドシートの画面に、入力・確かめの画面を出せません。スプレッドシートを開いたまま、そのスプレッドシートの' +
      '「拡張機能」→「Apps Script」から開いたエディタで実行してください（入力・確かめの画面は、スプレッドシートのタブに出ます）。');
  }
  return ui;
}

/** スプレッドシートの画面。出せなければ null */
function mtgUiOrNull_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (error) {
    return null;
  }
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
    statusTable: mtgReadStatusTable_(ss),
    credentials: hsCredentialIndex_(ss)
  };
  // 前回の会議時点のコピーがあれば、前回からの変化はそれと比べる（開けなければ、変更の記録から出して知らせる）
  const copies = plans.map(p => (p.previous && p.previous.copy ? mtgOpenCopy_(p.previous.copy, p.name) : null));
  const summaries = plans.map((p, i) => mtgSummarize_(ss, p.name, Object.assign({
    since: p.since, previous: p.previous, legacy: p.legacy, baseline: p.baseline, draft: p.draft, final,
    copy: copies[i] && copies[i].table ? copies[i] : null, copyError: copies[i] && copies[i].error || ''
  }, context)));
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  const results = summaries.map(summary => {
    const out = mtgExportDeck_(mtgFileName_(summary.name, meeting, final), mtgSlides_(summary), folder);
    return {summary, file: out.pdf, deck: out.deck, folder};
  });
  // 本番：今の値を会議時点のコピーに保存する（次の会議の資料はこれと比べる）。資料を作れてから保存し、会議日の記録に結び付ける
  const copy = final ? mtgSaveCopy_(ss, names, now, '（会議日 ' + meeting.replace(/\//g, '-') + '）') : null;
  results.forEach(r => { r.copy = copy; });
  plans.forEach((p, i) => {
    if (final) {
      const list = p.meetings.filter(m => m.baseline || m.date !== meeting)
        .concat([{date: meeting, cutoff: now.getTime(), copy: {id: copy.id, sheet: p.name}}])
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

/* ---------------- 会議時点のコピー ---------------- */

/**
 * 追跡シート（names）の今の値を、新しいスプレッドシート（タイトルは「会議時点のコピー_実行日時」＋ note）に保存する。シートごとに同じ名前のシートにし、
 * 同じ位置に値だけを写す（数式は計算した値に。文字は先頭に ' を付けて数式にならないように。日付・数値の書式も写す）。
 * 保存先は mtgCopyFolder_。途中で止まったら作りかけのファイルはゴミ箱に移す。戻り値 {id, title, url}
 */
function mtgSaveCopy_(ss, names, now, note) {
  const o = MTG_OPTIONS;
  const title = o.copyName + '_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + (note || '');
  const folder = mtgCopyFolder_(ss);
  const book = SpreadsheetApp.create(title);
  try {
    const blank = book.getSheets()[0];
    names.forEach((name, i) => {
      const source = ss.getSheetByName(name);
      if (!source) throw new Error('「' + name + '」シートが見つかりません。');
      const target = i === 0 ? blank.setName(name) : book.insertSheet(name);
      const rows = Math.max(source.getLastRow(), 1);
      const cols = Math.max(source.getLastColumn(), 1);
      diffEnsureGrid_(target, rows, cols);
      for (let start = 1; start <= rows; start += o.copyChunkRows) {
        const count = Math.min(o.copyChunkRows, rows - start + 1);
        const from = source.getRange(start, 1, count, cols);
        target.getRange(start, 1, count, cols)
          .setValues(from.getValues().map(line => line.map(v => (typeof v === 'string' ? credText_(v) : v))))
          .setNumberFormats(from.getNumberFormats());
      }
      diffTrimGrid_(target, cols, rows);
      const header = DIFF_RULES[name] ? DIFF_RULES[name].headerRow : 1;
      if (target.getMaxRows() > header) target.setFrozenRows(header);
    });
    DriveApp.getFileById(book.getId()).moveTo(folder);
  } catch (error) {
    try { DriveApp.getFileById(book.getId()).setTrashed(true); } catch (_) {}
    throw new Error('会議時点のコピーを保存できませんでした（' + (error && error.message || error) + '）。もう一度お試しください。');
  }
  return {id: book.getId(), title, url: book.getUrl()};
}

/** 会議時点のコピーの保存先（会議サマリのフォルダの中の copyFolderName。作ったフォルダの ID をドキュメントのプロパティに残す） */
function mtgCopyFolder_(ss) {
  const props = PropertiesService.getDocumentProperties();
  const saved = props.getProperty(MTG_OPTIONS.copyFolderKey);
  if (saved) {
    try {
      const folder = DriveApp.getFolderById(saved);
      if (!folder.isTrashed()) return folder;
    } catch (_) {}
  }
  const folder = mtgFolder_(ss).createFolder(MTG_OPTIONS.copyFolderName);
  props.setProperty(MTG_OPTIONS.copyFolderKey, folder.getId());
  return folder;
}

/**
 * 前回の会議時点のコピー（copy = {id, sheet}）の、そのシートの案件を読む。{table: mtgReadDeals_ の形, title, url}。
 * 開けない（消した・権限が無い）・シートや見出しが無いときは {error: 理由}（資料は変更の記録から作る）
 */
function mtgOpenCopy_(copy, name) {
  let book;
  try {
    book = SpreadsheetApp.openById(copy.id);
  } catch (_) {
    return {error: '前回の「' + MTG_OPTIONS.copyName + '」を開けないため、前回からの変化は変更の記録から出しました（ファイルを消した・権限が無いなど）。'};
  }
  const sheet = book.getSheetByName(copy.sheet || name);
  if (!sheet) return {error: '前回の「' + book.getName() + '」に「' + (copy.sheet || name) + '」シートが無いため、前回からの変化は変更の記録から出しました。'};
  try {
    return {table: mtgReadDeals_(sheet, DIFF_RULES[name]), title: book.getName(), url: book.getUrl()};
  } catch (error) {
    return {error: '前回の「' + book.getName() + '」を読めないため、前回からの変化は変更の記録から出しました（' + error.message + '）。'};
  }
}

/** 追跡範囲（DIFF_RULES の ranges）の列の見出し [{key: 正規化した見出し, header}]（最終更新日時は除く）。コピーと今の行を比べる列 */
function mtgTrackedLabels_(sheet, rule, headers) {
  const stamp = diffNormalizeStampHeader_(DIFF_OPTIONS.stampHeader);
  const out = [];
  rule.ranges.forEach(a1 => {
    const range = sheet.getRange(a1);
    for (let c = range.getColumn(); c <= range.getLastColumn(); c++) {
      const header = String(headers[c - 1] || '');
      const key = credNormalize_(header);
      if (!key || diffNormalizeStampHeader_(header) === stamp || out.some(x => x.key === key)) continue;
      out.push({key, header});
    }
  });
  return out;
}

/* ---------------- 読む ---------------- */

/**
 * 記録のシート（変更履歴_差分・変更時点スナップショット）を下から読み、since 以降の行を返す（古い順）。
 * 記録は追記なので、サンプル（イベントID が SMP-）・ダミー（DMY-。addMeetingDummyHistory()）以外で since より前の行に届いたら止める
 * （どちらも過去の日時で足されるため。サンプルは読まず、ダミーは since 以降のものを読む）。
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
      const id = idCol != null ? String(rows[i][idCol]) : '';
      const at = mtgTime_(rows[i][atCol]);
      if (/^SMP-/.test(id)) continue;
      if (at != null && at < limit) {
        if (/^DMY-/.test(id)) continue;
        return Object.assign(out, {rows: out.rows.reverse()});
      }
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
      eventId: mtgCellText_(r[i['イベントID']]),
      at: i['記録日時'] != null ? mtgTime_(r[i['記録日時']]) : null,
      sheet: i['シート'] != null ? mtgCellText_(r[i['シート']]) : '',
      customer: pick(cols.customer) || (i['得意先'] != null ? mtgCellText_(r[i['得意先']]).trim() : ''),
      project: pick(cols.project),
      status: pick(cols.status),
      // その時点の行の値（変更のあとの値。数式の列も計算した値）。前回の時点の値に使う
      values: {status: pick(cols.status), amount: mtgNumber_(pick(cols.amount)), weighted: mtgNumber_(pick(cols.weighted)),
        monthText: pick(cols.orderMonth), activity: pick(cols.activity)}
    });
  });
  return map;
}

/**
 * 追跡シートの案件（得意先か案件名のある行）を読む。
 * {headers, has: {見出しのキー: シートにあるか}, statusOrder: ステータスのプルダウンの選択肢（無ければ空）, weightedFormula: 期待値の列が数式か,
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

  const out = {headers, has, hasStamp: stampCol >= 0, statusOrder: [], deals: [], weightedFormula: false};
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return out;
  // 期待値の列が数式（想定売上 × ステータスの率など）なら、変更の記録に残らないので、前回の値はスナップショットから取る
  if (has.weighted) out.weightedFormula = !!String(sheet.getRange(headerRow + 1, cols.weighted + 1).getFormula() || '').trim();
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
      missing: required.filter(r => !String(line[r.col]).trim()).map(r => r.label),
      // 見出し（正規化）→ 表示の値（会議時点のコピーと今の行を比べる用）
      cells: headers.reduce((map, h, c) => {
        const key = credNormalize_(h);
        if (key && !(key in map)) map[key] = String(line[c]).trim();
        return map;
      }, {})
    });
  });
  return out;
}

/* ---------------- まとめる ---------------- */

/**
 * 1つのシートの会議サマリの中身。
 * context = {since, previous, legacy, baseline: 基準を記録した時刻|null, draft, final, now, meeting, windowStart, trackedSince, log, snapshots, credentials}
 * 戻り値 {name, meeting, previous, legacy, since, first, final, now, windowStart, kpi, recordColumn: 商談記録の列があるか,
 *         stallPending: 停滞をまだ数えられないとき {at: 記録の始まりか基準の時刻, baseline},
 *         gaps: 作れなかった部分 → 理由（mtgGaps_）, alertGaps: チェックできなかった入力内容のチェックの種類 → 理由, bridge, stages, quarters,
 *         changes, stalled, focus: 期待値の大きい進行中の案件, overdue: 受注月を過ぎた進行中の案件, alerts: 入力内容のチェック,
 *         resolved, unmatched, dummy: 使ったダミーの変更の記録（DMY-）の数}
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
  const statusTable = 'statusTable' in context ? context.statusTable : mtgReadStatusTable_(ss);
  const rank = mtgStatusRank_(table.statusOrder, statusTable);
  const copy = context.copy || null;
  const rules = {rank, statusTable, weightedFormula: table.weightedFormula, copy: !!copy,
    tracked: copy ? mtgTrackedLabels_(sheet, DIFF_RULES[name], table.headers) : []};

  // 前回からの変化（案件ごとに1つの区分）。行を削除した案件は「削除」
  const changes = [];
  const changed = new Map();
  // 削除した行の記録（前回より後）。会議時点のコピーと比べるときは、削除した時刻に使う
  const removedLines = history.removed.filter(line => line.at > since).map(line => {
    const r = Object.assign(mtgParseRemoved_(line.text, table.headers), {at: line.at});
    r.key = mtgDealKey_(r.customer, r.project);
    return r;
  });
  if (copy) {
    // 前回の会議時点のコピーの行と今の行を、得意先＋案件名で突き合わせる（名前を書き換えた案件は、変更の記録で前の名前とつなげる）。
    // コピーに無い案件は新規、今のシートに無い案件は削除
    const before = new Map();
    copy.table.deals.forEach(d => {
      const key = history.root(d.key);
      if (!before.has(key)) before.set(key, d);
    });
    byKey.forEach((deal, key) => {
      const item = mtgChange_(key, deal, null, history.events.get(key) || [], [], since, labels, rules, before.get(key) || null);
      if (item) { changes.push(item); changed.set(key, item); }
    });
    before.forEach((prev, key) => {
      if (byKey.has(key)) return;
      const line = removedLines.filter(r => history.root(r.key) === key).pop();
      const r = Object.assign({}, prev, {key, at: line ? line.at : 0, logged: !!line});
      const item = mtgChange_(key, null, r, history.events.get(key) || [], [], since, labels, rules, prev);
      if (item) { changes.push(item); changed.set(key, item); }
    });
  } else {
    byKey.forEach((deal, key) => {
      const item = mtgChange_(key, deal, null, history.events.get(key) || [], history.states.get(key) || [], since, labels, rules);
      if (item) { changes.push(item); changed.set(key, item); }
    });
    removedLines.forEach(r => {
      if (!(r.customer || r.project) || byKey.has(r.key) || changed.has(r.key)) return;
      const item = mtgChange_(r.key, null, r, history.events.get(r.key) || [], history.states.get(r.key) || [], since, labels, rules);
      if (item) { changes.push(item); changed.set(r.key, item); }
    });
  }
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
  const alerts = mtgAlerts_(table, context, {slipsOf, changes, tracked});
  // 作れなかったところ（シートに列が無い・記録が無いなど。資料の型は変えずに、その場所に理由を書く）
  const gaps = mtgGaps_(ss, table, context, stallPending);
  // 会議時点のコピーでは変わっているのに、変更の記録が無い案件（差分追跡が止まっていた間・2,000セルを超える貼り付けなど）
  const unlogged = copy && context.trackedSince != null ? changes.filter(c => !c.logged).length : 0;
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
  const owners = new Set(alerts.map(a => a.owner || o.blankText));
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
      // 期待値・想定売上が空（合計に含めない）。列が無いときは空欄ではなく「データなし」（gaps）なので数えない
      missingWeighted: table.has.weighted ? active.filter(d => d.weighted == null).length : 0,
      missingAmount: table.has.amount ? active.filter(d => d.amount == null).length : 0,
      missingStatus: active.filter(d => !d.status).length,              // ステータスが空（進行中として数える。資料に件数を書く）
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
    gaps: gaps.parts,
    alertGaps: gaps.alerts,
    bridge,
    stages: mtgStages_(deals, table.statusOrder),
    quarters: mtgQuarters_(active, context.now),
    changes,
    stalled,
    focus,
    overdue: items.filter(d => d.overdue).sort((a, b) => (b.amount || 0) - (a.amount || 0)),
    alerts,
    resolved,
    unmatched: copy ? 0 : history.unmatched,   // コピーと比べるときは、変化は記録に頼らない
    copy: copy ? {title: copy.title, url: copy.url} : null,
    copyError: context.copyError || '',
    unlogged,
    dummy: history.dummy || 0
  };
}

/**
 * 作れなかったところ {parts: 部分 → 理由, alerts: 入力内容のチェックの種類 → 理由}（資料にそのまま書く。推定で補わない）。
 * parts のキー：changes 前回からの変化（変更の記録が無い）・stalled 停滞（記録の始まり・基準から stallDays 日たっていない）・
 * amount・weighted・orderMonth・activity・issue・owner・record（シートにその列が無い。キーは MTG_OPTIONS.columns）
 */
function mtgGaps_(ss, table, context, stallPending) {
  const o = MTG_OPTIONS;
  const tz = Session.getScriptTimeZone();
  const column = key => 'シートに「' + String(o.columns[key]).replace(/\s+/g, '') + '」の列が無い';
  const parts = {};
  const noLog = context.trackedSince == null;
  // 前回の会議時点のコピーと比べるときは、変更の記録が無くても前回からの変化を出せる
  if (noLog && !context.copy) parts.changes = '「' + DIFF_OPTIONS.logSheet + '」シートに変更の記録が無い';
  ['amount', 'weighted', 'orderMonth', 'activity', 'issue', 'owner', 'record'].forEach(key => {
    if (!table.has[key]) parts[key] = column(key);
  });
  if (stallPending) {
    const day = t => Utilities.formatDate(new Date(t), tz, 'M/d');
    parts.stalled = (stallPending.baseline ? '基準' : '記録の開始') + '（' + day(stallPending.at) + '）から ' + o.stallDays + '日未満のため、まだ数えられない（' +
      day(stallPending.at + o.stallDays * 86400000) + ' から数える）';
  }
  // 変更の記録も最終更新日時の列も無ければ、ステータスが同じ日数が分からない（停滞を数えられない）
  if (!parts.stalled && noLog && !table.hasStamp) parts.stalled = '変更の記録も「最終更新日時」の列も無く、ステータスが同じ日数が分からない';
  const alerts = {};
  if (!ss.getSheetByName(CRED_OPTIONS.logSheet)) alerts['予定日超過'] = '「' + CRED_OPTIONS.logSheet + '」シートが無い';
  if (parts.orderMonth) alerts['受注月超過'] = parts.orderMonth;
  if (parts.record) alerts['資料なし'] = parts.record;
  if (!table.hasStamp && noLog) alerts['更新なし'] = 'シートに「最終更新日時」の列が無く、変更の記録も無い';
  if (parts.amount || parts.weighted) alerts['金額の不整合'] = parts.amount || parts.weighted;
  if (parts.orderMonth || !table.has.startMonth) alerts['月の不整合'] = parts.orderMonth || column('startMonth');
  if (!table.statusOrder.length) alerts['選択肢外のステータス'] = '「' + o.columns.status + '」の列にプルダウンが無い';
  if (!table.has.delivery) alerts['デリバリー担当なし'] = column('delivery');
  else if (noLog) alerts['デリバリー担当なし'] = '「' + DIFF_OPTIONS.logSheet + '」シートに変更の記録が無い';
  return {parts, alerts};
}

/**
 * 作れなかったところの一覧 ['停滞：基準（10/16）から…', …]（作り終えたときの画面とログに出す）。
 * 同じ理由の部分はまとめる（「受注月超過・入力内容のチェック「受注月超過」：シートに「受注月」の列が無い」）
 */
function mtgGapList_(s) {
  const names = {changes: '前回からの変化', stalled: '停滞', amount: '想定売上', weighted: '期待値', orderMonth: '受注月・受注月超過',
    activity: '活動状況', issue: '活動における課題', owner: '責任者', record: '商談記録未登録'};
  const byReason = new Map();
  const add = (label, reason) => {
    if (!byReason.has(reason)) byReason.set(reason, []);
    byReason.get(reason).push(label);
  };
  Object.keys(s.gaps || {}).forEach(key => add(names[key] || key, s.gaps[key]));
  MTG_ALERT_TYPES.filter(type => (s.alertGaps || {})[type]).forEach(type => add('チェック「' + (MTG_ALERT_LABELS[type] || type) + '」', s.alertGaps[type]));
  return Array.from(byReason.keys()).map(reason => byReason.get(reason).join('・') + '：' + reason);
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
  let dummy = 0;   // ダミーの記録（DMY-。addMeetingDummyHistory()）の数。あれば資料に【ダミーデータを含む】と出す
  context.log.forEach(e => {
    if (resolve(e.sheet) !== name || e.at == null || /^(SMP|CRD)-/.test(e.eventId)) return;
    if (/^DMY-/.test(e.eventId)) dummy++;
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
  // 案件ごとの、変更の記録の時点の行の値（スナップショット。古い順）。前回の時点の値は、前回より前でいちばん新しいものから取る
  const states = new Map();
  context.snapshots.forEach(snap => {
    // サンプル（SMP-）・クレデンシャル（CRD-）の記録は数えない（変更の記録と同じ。サンプルは過去の日時で足される）
    if (resolve(snap.sheet) !== name || snap.at == null || !(snap.customer || snap.project) || /^(SMP|CRD)-/.test(snap.eventId)) return;
    const key = root(mtgDealKey_(snap.customer, snap.project));
    if (!states.has(key)) states.set(key, []);
    states.get(key).push({at: snap.at, values: snap.values});
  });
  states.forEach(list => list.sort((a, b) => a.at - b.at));
  return {events, removed, unmatched, states, dummy, root};
}

/**
 * ステータスの定義のシート（MTG_OPTIONS.statusTable）を読む。[{status, ratio: 期待値の率（無ければ null）}]（定義の上から）。シートが無ければ null
 */
function mtgReadStatusTable_(ss) {
  const o = MTG_OPTIONS.statusTable;
  const sheet = o && ss.getSheetByName(o.sheet);
  if (!sheet || sheet.getLastRow() <= o.headerRow) return null;
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(o.headerRow, 1, 1, width).getDisplayValues()[0].map(credNormalize_);
  const sc = headers.indexOf(credNormalize_(o.status));
  const rc = headers.indexOf(credNormalize_(o.ratio));
  if (sc < 0) return null;
  return sheet.getRange(o.headerRow + 1, 1, sheet.getLastRow() - o.headerRow, width).getValues()
    .map(line => ({status: String(line[sc]).trim(), ratio: rc >= 0 ? mtgNumber_(line[rc]) : null}))
    .filter(item => item.status);
}

/** ステータスの期待値の率（定義のシートにあるときだけ。無ければ null） */
function mtgStatusRatio_(statusTable, status) {
  const item = (statusTable || []).find(x => credNormalize_(x.status) === credNormalize_(status));
  return item ? item.ratio : null;
}

/**
 * ステータスの進み具合（大きいほど進んでいる。分からないステータスは -1）を返す関数。
 * 定義のシートがあれば、期待値の率の大きい順、率が同じなら定義の下の行ほど進んでいる（A-受注済み(100%) > B > C-稟議対応中 > C-選定対応中 > C-提案中 > D …）。
 * 無ければ、プルダウンの順（下ほど進んでいる）
 */
function mtgStatusRank_(statusOrder, statusTable) {
  const table = (statusTable || []).filter(x => x.ratio != null);
  if (table.length) {
    return status => {
      const i = table.findIndex(x => credNormalize_(x.status) === credNormalize_(status));
      return i < 0 ? -1 : table[i].ratio * 1000 + i;
    };
  }
  const order = (statusOrder || []).map(credNormalize_);
  return status => order.indexOf(credNormalize_(status));
}

/**
 * 1つの案件の前回からの変化（変化が無ければ null）。deal は今の行、removed は削除した行（どちらか一方）。
 * 前回の時点の値（推定で補わない）：
 * - 前回（since）より前の変更の記録があれば、そのうちいちばん新しいスナップショット（変更のあとの行の値。数式の列も計算した値）
 * - 無ければ、前回以降で最初の変更の「変更前」（その項目の変更が無ければ今の値）。期待値の列が数式なら変更前は記録に残らないので、
 *   ステータス・想定売上が変わっていなければ今の値、変わっていれば、シートの数式と同じく前回の想定売上 × 前回のステータスの率（statusTable）
 * rules = {rank: ステータスの進み具合, statusTable, weightedFormula: 期待値の列が数式か}。states は案件のスナップショット（古い順）
 */
function mtgChange_(key, deal, removed, events, states, since, labels, rules, prev) {
  const recent = events.filter(e => e.at > since);   // 前回の資料を作った時刻ちょうどの記録は、前回の資料に入っている
  // 会議時点のコピーと比べるとき（rules.copy）：prev はコピーの行（無ければ null：前回は無かった案件）。値が同じなら変化なし
  const byCopy = !!rules.copy;
  const differ = byCopy && deal && prev
    ? rules.tracked.filter(t => (prev.cells[t.key] || '') !== (deal.cells[t.key] || '')).map(t => t.header) : [];
  if (byCopy ? deal && prev && !differ.length : !recent.length && !removed) return null;
  const born = !byCopy && recent.some(e => e.born);
  if (born && removed) return null;   // 前回より後に足して、消した行
  const cur = deal || removed;
  const pick = (label, value, parse) => {
    const e = recent.find(x => x.label === label);
    return e ? parse(e.before) : value;
  };
  const prior = states.filter(st => st.at <= since).pop();
  let base = null;
  if (byCopy) {
    base = prev ? {source: 'copy', status: prev.status, amount: prev.amount, weighted: prev.weighted, monthText: prev.orderMonthText,
      activity: prev.activity || ''} : null;
  } else if (born) {
    base = null;
  } else if (prior) {
    base = Object.assign({source: 'snapshot'}, prior.values);
  } else {
    base = {
      source: 'log',
      status: pick(labels.status, cur.status, String),
      amount: pick(labels.amount, cur.amount, mtgNumber_),
      weighted: pick(labels.weighted, cur.weighted, mtgNumber_),
      monthText: pick(labels.orderMonth, cur.orderMonthText, String),
      activity: pick(labels.activity, cur.activity || '', String)
    };
    if (rules.weightedFormula) {
      const touched = recent.some(e => e.label === labels.status || e.label === labels.amount);
      const ratio = touched ? mtgStatusRatio_(rules.statusTable, base.status) : null;
      base.weighted = !touched ? cur.weighted : ratio != null && base.amount != null ? Math.round(base.amount * ratio * 1e6) / 1e6 : null;
    }
  }
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
      } else if (deal.kind === 'active' && rules.rank(base.status) >= 0 && rules.rank(deal.status) >= 0) {
        flags.push(rules.rank(deal.status) > rules.rank(base.status) ? 'forward' : 'backward');
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
  const columns = byCopy ? differ.slice() : [];
  if (!byCopy) {
    recent.forEach(e => {
      if (e.column && e.column !== '（行全体）' && columns.indexOf(e.column) < 0) columns.push(e.column);
    });
  }
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
    // 変更の記録があるか（コピーと比べたとき、記録の無い変化を知らせる）
    logged: recent.length > 0 || !!(removed && removed.logged),
    at: Math.max.apply(null, recent.map(e => e.at).concat([removed ? removed.at || 0 : 0, !recent.length && deal ? deal.updatedAt || 0 : 0]))
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
 * info = {slipsOf: キー → 受注月の延期の回数, changes, tracked: 変更の記録がそろっている始まり（ms）}
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
  // （サンプル（名称の先頭が【サンプル】。addSampleHistory() で足したもの）は数えない）
  const customers = new Map();
  active.forEach(d => {
    const key = credNormalize_(d.customer);
    if (d.customer && !customers.has(key)) customers.set(key, d.customer);
  });
  const overdue = [];
  customers.forEach((customer, key) => {
    (context.credentials[key] || []).forEach(c => {
      if (String(c.values.name || '').indexOf(SMP_OPTIONS.namePrefix) === 0) return;
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
    if (table.statusOrder.length && d.status && table.statusOrder.map(credNormalize_).indexOf(credNormalize_(d.status)) < 0) {
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
    const label = d.status || MTG_OPTIONS.blankText;
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

/** 金額（百万）の表示：「1,234.5」（整数なら小数なし）。空（null）は「未入力」（blankText。0 とは書かない）。 */
function mtgAmount_(value) {
  if (value == null) return MTG_OPTIONS.blankText;
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
  return (item.customer || '（得意先未入力）') + '／' + (item.project || '（案件名未入力）');
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
  // Appendix はいつも ①〜③ を出す（中身が無ければ、表の見出しの行の下に「データなし」と理由・「該当なし」）
  const appendix = [].concat(mtgAppendixChanges_(summary), mtgAppendixStalled_(summary), mtgAppendixOwners_(summary),
    mtgAppendixAlerts_(summary));
  return mtgNumberSlides_(mtgSheetFrame_(summary), [cover].concat(main, [mtgDividerSlide_(appendix, 1 + main.length + 1 + 1)], appendix));
}

/**
 * 表紙（ひな形の1枚目）：ひな形の決まった文字を置き換える（template。見出しにサブインダストリー、その下に会議の名前と「事前資料」
 * （試作は【試作】）、日付に会議日、名前の欄は空）。その下にデータの時点・期間を小さく足す（elements）
 */
function mtgCoverSlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const line = (y, text) => mtgBox_(S.x, y, S.w, 16, mtgWrap_(text, S.w, 9, 1)[0], {size: 9, color: c.white, align: 'CENTER'});
  return {
    title: MTG_TITLES.cover,
    cover: true,
    template: {title: s.name, subtitle: o.meetingName + '　事前資料' + (s.final ? '' : '【試作】'),
      date: mtgDateLabel_(s.meeting), name: ''},
    elements: [
      line(350, '出典：「' + s.name + '」シート' + (s.copy ? '・前回の会議時点のコピー' : '') + 'および変更履歴（' + (s.dummy ? 'ダミーデータを含む。' : '') +
        Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/M/d HH:mm') + ' 時点）'),
      line(365, '対象期間：' + mtgPeriodText_(s))
    ]
  };
}

/**
 * 扉（Appendix。ひな形の1枚目を複製）：見出しを「Appendix」にし（ほかの決まった文字は空）、その下に Appendix の見出し・始まりのページの一覧
 * （白い字と白い線。中身が無いものは「（データなし）」「（該当なし）」を付ける）。appendix は Appendix のスライド、firstPage はその1枚目のページ番号
 */
function mtgDividerSlide_(appendix, firstPage) {
  const c = MTG_OPTIONS.colors;
  const items = [];
  appendix.forEach((slide, i) => {
    const m = /^(\S+ \S+)　(.*)$/.exec(slide.title) || [slide.title, slide.title, ''];
    if (items.some(it => it.label === m[1])) return;
    // 中身が無い Appendix も一覧に出し、「（データなし）」（作れなかった）・「（該当なし）」を付ける
    const state = slide.empty === 'gap' ? '（' + MTG_OPTIONS.gapText + '）' : slide.empty === 'none' ? '（' + MTG_OPTIONS.noneText + '）' : '';
    items.push({label: m[1], title: m[2].replace(/（[^（）]*）$/, '') + state, page: firstPage + i});
  });
  const el = [];
  const w = 400;   // 一覧の幅（見出し 98・名前・ページ番号 40）。「（データなし）」が付いても1行に入る幅
  const x = (MTG_SLIDE.width - w) / 2;
  const rowH = 26;
  items.forEach((it, i) => {
    const y = 228 + i * rowH;
    el.push(mtgBox_(x, y, 104, 22, it.label, {size: 12, bold: true, color: c.white, role: 'toc'}));
    el.push(mtgBox_(x + 98, y, w - 138, 22, mtgWrap_(it.title, w - 138, 12, 1)[0], {size: 12, color: c.white, role: 'toc'}));
    el.push(mtgBox_(x + w - 40, y, 40, 22, String(it.page), {size: 12, color: c.white, align: 'END', role: 'toc'}));
    el.push({type: 'box', x, y: y + 23, w, h: 0.75, fill: c.white, alpha: 0.5, shape: 'RECTANGLE'});
  });
  return {title: MTG_TITLES.appendix, divider: true, template: {title: MTG_TITLES.appendix, subtitle: '', date: '', name: ''}, elements: el};
}

/** スライドに、どのスライドにも付ける部品（mtgFrame_）とページ番号を付ける（表紙・扉の印と、ひな形の置き換える文字は残す） */
function mtgNumberSlides_(frame, slides) {
  return slides.map((slide, i) => {
    const out = {title: slide.title, elements: mtgFrame_(frame, slide, i + 1).concat(slide.elements)};
    ['cover', 'divider', 'template', 'empty'].forEach(key => { if (slide[key]) out[key] = slide[key]; });
    return out;
  });
}

/** シートの資料のいちばん下の行（どのサブインダストリー・会議日の資料か。試作は先頭に【試作】）{source}。データの時点・期間は表紙に書く */
function mtgSheetFrame_(s) {
  return {source: (s.final ? '' : '【試作】') + (s.dummy ? '【ダミーデータを含む】' : '') + s.name + '（' + mtgShortDate_(s.meeting) + ' 会議資料）　金額の単位：百万円'};
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
  // 箱の高さは、上の余白（padY の半分）と文字の行の分（文字は箱の上にそろえるので、下の余白は数えない）
  const el = [mtgBox_(S.x, S.leadY, S.w, step(first.length) + S.padY / 2, first.join('\n'), {size, color: c.ink, valign: 'TOP', role: 'lead'})];
  if (rest) {
    const second = mtgWrap_(rest, S.w, size, 3 - first.length);
    el.push(mtgBox_(S.x, S.leadY + step(first.length), S.w, step(second.length) + S.padY / 2, second.join('\n'),
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
  const lack = s.gaps || {};
  const G = o.gapText;
  // 前回比は変更の記録から出す（記録が無ければ、前回比は「データなし」と書き、±0 と書かない）
  const lead = '進行中案件 ' + k.active + '件' + (lack.changes ? '' : '（' + vs + mtgSigned_(k.active - b.startCount) + '件）') + '、' +
    (lack.weighted ? '期待値合計は' + G + '（' + lack.weighted + '）' : '期待値合計 ' + mtgAmount_(b.end) + '百万円' +
      (lack.changes ? '' : '（' + vs + mtgSigned_(b.end - b.start) + '）')) + '。' +
    (lack.changes ? (s.first ? mtgPeriodShort_(s) + 'の増減' : vs.trim()) + 'は' + G + '（' + lack.changes + '）。' : '') + mtgBlankNote_(k);
  // 中身はリード文の下から（リード文が2行になれば下げる。箱の高さには上の余白が入っている）
  const top = Math.max(S.summaryTop, Math.max.apply(null, mtgLead_(lead).map(e => e.y + e.h)) + 2);
  const el = [];

  // 表：案件の名前は1行から始め、ページに入る限り上の行から2行にする。入る分だけ並べ、残りは「他 N件」。件数は大きな緑の字。
  // 区分の名前のあとに小さな字の補足（停滞の定義）を付けられる（同じ行に入らなければ2行目に）
  const columns = [136, 60, 62, S.w - 258];
  const subSize = 7.5;
  const rowOf = (line, lines, lay) => {
    const body = lay.body;
    const dash = (align) => ({text: '—', size: body, color: c.muted, fill: c.white, align});
    const inline = line.sub && mtgTextWidth_(line.label, body) + mtgTextWidth_(line.sub, subSize) <= (columns[0] - S.padX) * S.fit;
    const label = !line.sub ? {text: line.label, size: body, lines: 2, bold: true, color: c.ink, fill: c.greenSoft}
      : inline ? {text: line.label + line.sub, size: subSize, color: c.muted, fill: c.greenSoft,
        spans: [{start: 0, end: line.label.length, size: body, bold: true, color: c.ink}]}
        : {text: line.label + '\n' + line.sub, size: subSize, lines: 2, color: c.muted, fill: c.greenSoft, parts: [{bold: true, size: body, color: c.ink}, {}]};
    // 作れなかった行（記録・列が無い）：件数・想定売上は「—」、案件の欄は灰色の「データなし」と理由（理由を見出しに書く行は「データなし」だけ）
    if (line.gap) {
      return mtgRow_(columns, [label, dash('CENTER'), dash('END'), mtgGapCell_({size: body, lines}, G + (line.gapShort ? '' : '：' + line.gap))], 0);
    }
    return mtgRow_(columns, [
      label,
      {text: line.items.length + '件', size: lay.count, bold: true, color: c.greenDark, fill: c.white, align: 'CENTER'},
      {text: line.items.length ? mtgAmount_(mtgSum_(line.items, 'amount')) : '—', size: body, fill: c.white, align: 'END'},
      {text: line.items.length || line.prefix ? mtgNames_(line.shown || line.items, line.note, columns[3], lines, line.prefix, body)
        : line.empty || MTG_OPTIONS.noneText, size: body, lines, fill: c.white}
    ], 0);
  };
  const groups = mtgBreakdown_(s);
  const all = groups.moves.concat(groups.states);
  const height = rows => rows.reduce((sum, r) => sum + r.h, 0);
  // 高さ：数字の箱 th・箱の下 after・表の見出し headH・2つの表の間 gap、文字の大きさ：件数 count・表 body。
  // 入りきらなければ（リード文が長いときなど）、件数の文字 → 箱の下 → 表の間 → 表の見出し → 数字の箱 → 表の文字 の順に、決めた下限まで小さくする
  const lay = {th: 34, after: 6, headH: 16, gap: 5, count: 10, body: S.bodySize};
  const build = () => {
    const head = mtgHeadRow_(columns, ['区分', '件数', '想定売上', '主な案件（想定売上順）'], lay.body);
    const rows = all.map(line => rowOf(line, 1, lay));
    return {head, rows, fixed: top + lay.th + lay.after + 2 * lay.headH + lay.gap + head.h};
  };
  let plan = build();
  [['count', 9], ['after', 4], ['gap', 3], ['headH', 14], ['th', 30], ['body', 8], ['count', 8]].forEach(([key, min]) => {
    const short = plan.fixed + height(plan.rows) - S.bottom;
    if (short <= 0) return;
    lay[key] = key === 'count' || key === 'body' ? Math.min(lay[key], min) : Math.max(min, lay[key] - short);
    plan = build();
  });
  const head = plan.head;
  let rows = plan.rows;
  const headH = lay.headH;
  const gap = lay.gap;
  const tablesTop = top + lay.th + lay.after;
  const room = S.bottom - plan.fixed;
  all.forEach((line, i) => {
    const two = rowOf(line, 2, lay);
    if (height(rows) - rows[i].h + two.h <= room) rows[i] = two;
  });
  // 下にあいた分は、表の行に分けて高くする（1行あたり 10pt まで。ページの下が空かないように）
  const extra = Math.floor(Math.max(0, Math.min(10, (room - height(rows)) / rows.length)) * 10) / 10;
  if (extra > 0) rows = rows.map(r => Object.assign({}, r, {h: r.h + extra}));
  const moves = [head].concat(rows.slice(0, groups.moves.length));
  const states = rows.slice(groups.moves.length);
  const heading = (y, text) => mtgBox_(S.x, y, S.w, headH, text, {size: 10.5, bold: true, color: c.ink, role: 'heading'});
  const statesY = tablesTop + headH + height(moves) + gap;

  // 合計の数字：緑の帯の見出しと、大きな数字・小さな前回比（1行）
  // 列が無い金額は「データなし」（灰色の箱）。前回比は、変更の記録が無ければ「前回比 データなし」
  const versus = delta => vs + (lack.changes ? G : delta);
  const tiles = [
    {label: '進行中案件', value: k.active + '件', sub: versus(mtgSigned_(k.active - b.startCount) + '件')},
    lack.amount ? {label: '想定売上合計（進行中・百万円）', value: G, gap: true}
      : {label: '想定売上合計（進行中・百万円）', value: mtgAmount_(b.amountEnd), sub: versus(mtgSigned_(b.amountEnd - b.amountStart))},
    lack.weighted ? {label: '期待値合計（進行中・百万円）', value: G, gap: true}
      : {label: '期待値合計（進行中・百万円）', value: mtgAmount_(b.end), sub: versus(mtgSigned_(b.end - b.start))},
    {label: '入力内容の確認事項', value: k.alerts + '項目', sub: k.alerts ? k.alertDeals + '案件・責任者 ' + k.owners + '名' : ''}
  ];
  const tw = (S.w - 3 * 8) / 4;
  const th = lay.th;
  tiles.forEach((t, i) => {
    const x = S.x + i * (tw + 8);
    const text = t.value + (t.sub ? '　' + t.sub : '');
    el.push({type: 'box', x, y: top, w: tw, h: th, fill: t.gap ? c.gap : c.white, line: t.gap ? c.gapLine : c.frame, lineWeight: 0.75,
      dash: !!t.gap, shape: 'RECTANGLE', role: t.gap ? 'gap' : undefined});
    el.push(mtgBox_(x, top, tw, 14, mtgWrap_(t.label, tw, 8.5, 1)[0],
      {size: 8.5, bold: true, color: c.white, fill: c.green, shape: 'RECTANGLE', align: 'CENTER', role: 'tile'}));
    el.push(mtgBox_(x, top + 14, tw, th - 14, mtgWrap_(text, tw, 8, 1)[0], {size: 8, color: c.muted, align: 'CENTER', role: 'tile-value',
      spans: [t.gap ? {start: 0, end: t.value.length, size: 11, bold: true, color: c.dark} : {start: 0, end: t.value.length, size: 15, bold: true, color: c.greenDark}]}));
  });
  el.push(heading(tablesTop, period + 'の動き（' + (lack.changes ? G + '：' + lack.changes : '変化のある案件 ' + k.changed + '件') + '）'),
    {type: 'table', x: S.x, y: tablesTop + headH, columns, rows: moves},
    heading(statesY, '現在の状態（進行中案件）'),
    {type: 'table', x: S.x, y: statesY + headH, columns, rows: states});
  return {title: MTG_TITLES.summary, lead, elements: el};
}

/**
 * サマリーの表の行 {moves, states}。どの行も {label, sub: 区分の名前の下の小さな補足, items: 数える案件（想定売上の大きい順）,
 * shown: 名前を並べる案件（無ければ items）, prefix: 名前の前に書く文, note: 案件の名前の後ろの（）の中,
 * gap: 作れない理由（灰色の「データなし」にする）, gapShort: 理由を表の見出しに書くので、セルは「データなし」だけ}。
 * moves：前回以降の動き。案件ごとの区分（MTG_CATEGORIES）で1回だけ数え、すべての区分（その他更新を除く）をどれかの行に入れる
 * （受注／失注・保留・後退・削除／新規・前進・復帰／活動更新／受注月・期待値などの変更。P4 の並び（MTG_MOVE_ORDER）と同じ順）。
 * 活動更新は、最新の活動状況を（）に書く。
 * states：現在の状態（受注月超過・停滞・商談記録未登録）。停滞は、前回以降に更新の無い案件を「うち前回以降の更新なし N件」として名前を並べる
 * （初回は「直近N日間」、基準の後は「基準以降」。mtgQuietLabel_）。
 * どの行も、どのシートでも出す（資料の型を変えない）。作れない行（変更の記録・列が無い、停滞をまだ数えられない）は gap に理由を入れる
 */
function mtgBreakdown_(s) {
  const o = MTG_OPTIONS;
  const byAmount = list => list.slice().sort((a, b) => (b.amount || 0) - (a.amount || 0));
  const of = keys => byAmount(s.changes.filter(ch => keys.indexOf(ch.category) >= 0));
  const label = ch => mtgCategory_(ch.category).label;
  const status = ch => (ch.statusFrom || '（未入力）') + ' ▶ ' + (ch.statusTo || '（未入力）');
  const month = ch => mtgMonthLabel_(ch.monthFrom) + ' ▶ ' + mtgMonthLabel_(ch.monthTo);
  const value = ch => mtgAmount_(ch.weightedFrom) + ' ▶ ' + mtgAmount_(ch.weightedTo);
  const quiet = s.stalled.filter(d => !d.touched);
  const gap = s.gaps || {};
  // gap：作れない理由（変更の記録・列が無い、停滞をまだ数えられない）。どの行も必ず出す（資料の型を変えない）
  const states = [
    {label: '受注月超過', items: byAmount(s.overdue), note: d => '受注月 ' + mtgMonthLabel_(d.orderMonthText), gap: gap.orderMonth},
    {label: '停滞', sub: '（同一ステータス' + o.stallDays + '日以上）', items: s.stalled, shown: quiet,
      prefix: s.stalled.length ? mtgQuietLabel_(s) + ' ' + quiet.length + '件' + (quiet.length ? '：' : '') : '',
      note: d => d.days + '日' + (d.atLeast ? '以上' : ''), gap: s.stalled.length ? null : gap.stalled},
    {label: '商談記録未登録', items: byAmount(s.alerts.filter(a => a.type === '資料なし')), gap: gap.record}
  ];
  const moves = [
    {label: '受注', items: of(['won'])},
    {label: '失注・保留・後退・削除', items: of(['lost', 'hold', 'backward', 'removed']), note: ch => (ch.category === 'backward' ? status(ch) : label(ch))},
    {label: '新規・前進・復帰', items: of(['new', 'forward', 'revived']), note: ch => (ch.category === 'forward' ? status(ch) : label(ch))},
    {label: '活動更新', items: of(['activity']), note: ch => mtgClip_(ch.activity || o.blankText, 20), gap: gap.activity},
    {label: '受注月・期待値などの変更', items: of(['slipped', 'pulled', 'up', 'down', 'status']),
      note: ch => label(ch) + ({slipped: ' ' + month(ch), pulled: ' ' + month(ch), up: ' ' + value(ch), down: ' ' + value(ch), status: ' ' + status(ch)}[ch.category] || '')}
  ];
  // 変更の記録が無ければ、動きの行はどれも「データなし」（理由は表の見出しに書く）
  if (gap.changes) moves.forEach(line => Object.assign(line, {gap: gap.changes, gapShort: true}));
  return {moves, states};
}

/**
 * サマリーのリード文の補足：空欄をどう扱ったか（推定で補わないので、件数を書く）。空欄が無ければ空。
 * 「\nステータス未入力 1件は進行中に含め、想定売上未入力 1件・期待値未入力 1件は金額の合計に含めていない。」
 */
function mtgBlankNote_(k) {
  const amounts = [k.missingAmount ? '想定売上未入力 ' + k.missingAmount + '件' : '', k.missingWeighted ? '期待値未入力 ' + k.missingWeighted + '件' : '']
    .filter(Boolean).join('・');
  const status = k.missingStatus ? 'ステータス未入力 ' + k.missingStatus + '件は進行中に含め' : '';
  if (!status && !amounts) return '';
  return '\n' + (status && amounts ? status + '、' + amounts + 'は金額の合計に含めていない。'
    : status ? status + 'ている。' : amounts + 'は金額の合計に含めていない。');
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
 * width・lines を渡すと、その幅・行数に切れずに入る件数まで減らす（「他 N件」が「…」で切れないように）。prefix は名前の前に書く文、
 * size は文字の大きさ（既定は bodySize）
 */
function mtgNames_(items, note, width, lines, prefix, size) {
  const head = prefix || '';
  if (!items.length) return head || MTG_OPTIONS.noneText;
  const text = n => head + items.slice(0, n).map(d => mtgDealLabel_(d) + (note ? '（' + note(d) + '）' : '')).join('、') +
    (items.length > n ? '　他 ' + (items.length - n) + '件' : '');
  let n = Math.min(MTG_OPTIONS.mainRows.deals, items.length);
  while (width && n > 1 && mtgWrap_(text(n), width, size || MTG_SLIDE.bodySize, 99).length > lines) n--;
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
  const lack = s.gaps || {};
  const G = MTG_OPTIONS.gapText;
  // 期待値の列・変更の記録が無ければ、滝の図の場所に「データなし」と理由（見出しは残す）
  const flowGap = lack.weighted || lack.changes;
  const el = flowGap ? [mtgBox_(S.x, S.top - 2, 190, 20, '期待値の増減（進行中・百万円）', {size: 10.5, bold: true, color: MTG_OPTIONS.colors.ink}),
    mtgEmptyBox_(S.x, S.top + 26, 360, 64, 'gap', flowGap)] : mtgWaterfall_(s, S.x, S.top, 360, S.bottom - S.top);
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
  const labelW = mtgBarLabelWidth_(statusRows.concat(monthRows).map(r => r.label));
  const byStatus = mtgBarChart_('ステータス別期待値（進行中・百万円）', statusRows, right, S.top, width, max, rowH, '進行中案件なし', labelW,
    lack.weighted);
  const byMonth = mtgBarChart_('受注月別期待値（進行中・百万円）', monthRows, right, byStatus.bottom + gap, width, max, rowH, '進行中案件なし', labelW,
    lack.weighted || lack.orderMonth);
  el.push.apply(el, byStatus.elements.concat(byMonth.elements));
  const first = lack.weighted ? '進行中の期待値は' + G + '（' + lack.weighted + '）。'
    : lack.changes ? '進行中の期待値は ' + mtgAmount_(b.end) + '百万円。' + (s.first ? mtgPeriodShort_(s) + 'の増減' : mtgVersus_(s).trim()) + 'は' + G + '（' + lack.changes + '）。'
      : '進行中の期待値は ' + mtgAmount_(b.start) + ' → ' + mtgAmount_(b.end) + '百万円（' + mtgSigned_(delta) + '）。増加 ' + mtgSigned_(plus) +
        '・減少 ' + mtgSigned_(minus) + (won && won.value ? '（うち受注 ' + mtgSigned_(won.value) + '）' : '') +
        (biggest && biggest.value ? '、内訳の最大は' + biggest.label + '（' + biggest.count + '件・' + mtgSigned_(biggest.value) + '）' : '') + '。';
  const second = lack.weighted ? '' : lack.orderMonth ? '受注月別の期待値は' + G + '（' + lack.orderMonth + '）。'
    : '受注月が今期の期待値は ' + mtgAmount_(mtgSum_(s.quarters.filter(q => /Q/.test(q.label)), 'value')) + '百万円' +
      (peak ? '（最大は' + peak.label.replace(/（.*/, '') + 'の ' + mtgAmount_(peak.value) + '）' : '') +
      (over ? '、受注月超過は ' + mtgAmount_(over.value) + '百万円' : '') +
      (s.kpi.missingWeighted ? '。期待値未入力の ' + s.kpi.missingWeighted + '件は金額に含めていない' : '') + '。';
  return {title: MTG_TITLES.pipeline, lead: first + (second ? '\n' + second : ''), elements: el};
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
  const startLabel = s.first ? MTG_OPTIONS.firstDays + '日前' : s.previous && s.previous.baseline ? '基準' : '前回';
  const bars = [{label: startLabel, value: b.start, count: b.startCount, total: true}]
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

/** 横棒のグラフの名前の箱の幅：いちばん長い名前が1行に入る幅（92〜150pt。ステータスの名前「B-クロージング中(70%)」などを「…」で切らない） */
function mtgBarLabelWidth_(labels) {
  let w = 92;
  while (w < 150 && labels.some(label => mtgWrap_(label, w, 8.5, 99).length > 1)) w += 2;
  return w;
}

/**
 * 横棒のグラフ（パイプラインの変化の右。ステータス別・受注月別の期待値）。見出しは滝の図と同じ形で、行ごとに 名前・件数・棒・値。
 * 棒は期待値の色（greenDark。滝の図の「期待値」と同じ）。値は棒の先に書くので、目盛りは置かない（0 の線だけ）。
 * rows = [{label, count, value}]（この順に上から）。max は棒の長さの基準（2つのグラフで同じにする）。左上 (x, y)・幅 w・行の高さ rowH。
 * gap があれば（列が無いなど）、グラフの代わりに「データなし」と gap（理由）を置く。戻り値 {elements, bottom}
 */
function mtgBarChart_(title, rows, x, y, w, max, rowH, empty, labelW, gap) {
  const c = MTG_OPTIONS.colors;
  labelW = labelW || 92;                     // 名前の箱（2つのグラフで同じ幅にして、棒の始まりをそろえる。mtgBarLabelWidth_）
  const countW = 40;                         // 件数の箱（右寄せ。「123件」が折り返さない幅。名前の箱に 10pt 重ねる）
  const valueW = 51;                         // 値の箱（「1,234.5」が折り返さない幅）
  const base = x + labelW + countW - 6;      // 棒の始まり（0 の線）
  const room = x + w - base - (valueW - 3);  // 棒のいちばん長い長さ（右に値を書く分をあける。値は棒の先から 4pt あける）
  const top = y + 22;
  const el = [mtgBox_(x, y - 2, w, 20, title, {size: 10.5, bold: true, color: c.ink})];
  if (gap) {   // 作れない（列が無い）：見出しの下に「データなし」と理由
    el.push(mtgEmptyBox_(x, top, w, 52, 'gap', gap));
    return {elements: el, bottom: top + 52};
  }
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
  const lack = s.gaps || {};
  const moves = s.changes.filter(ch => ch.category !== 'updated');
  const columns = mtgColumns_([138, 66, 126, 66, 76, 124, 64], 5);   // ステータスの列は「C-稟議対応中(50%) ▶」「B-クロージング中(70%)」が1行ずつ入る幅
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '区分', 'ステータス\n前回▶今回', '期待値\n前回▶今回', '受注月\n前回▶今回', '活動状況（最新）', '責任者'])];
  // 上から mainRows.changes 件まで、ページに入る分だけ（行の高さは中身で変わる）
  const top = [];
  moves.slice(0, MTG_OPTIONS.mainRows.changes).some(ch => {
    const row = mtgRow_(columns, [
      mtgDealCell_(ch, columns[0]),
      {text: '', fill: c.white},
      {text: mtgStatusMove_(ch), lines: 3, fill: c.white},   // 長いステータスの名前（「C-稟議対応中(50%) ▶ B-クロージング中(70%)」）は3行まで
      mtgFieldCell_(s, 'weighted', mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'orderMonth', mtgMonthMove_(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'activity', ch.activity, {lines: 2, fill: c.white}),
      mtgFieldCell_(s, 'owner', ch.owner, {lines: 2, fill: c.white, align: 'CENTER'})
    ]);
    if (S.top + rows.reduce((sum, r) => sum + r.h, 0) + row.h > S.bottom) return true;
    rows.push(row);
    top.push(ch);
    return false;
  });
  const el = [];
  if (top.length) {
    const table = {type: 'table', x: S.x, y: S.top, columns, rows};
    el.push(table);
    top.forEach((ch, i) => el.push.apply(el, mtgPills_(table, i + 1, 1, [mtgCategoryPill_(ch.category)], mtgCategoryPillWidth_(columns[1]))));
  } else {
    // 表の見出しの行の下に、変更の記録が無ければ「データなし」と理由、変化が無ければ「該当なし」（状態はリード文に書くので、同じ文を2回書かない）
    el.push.apply(el, mtgEmptyTable_(columns, rows[0], lack.changes ? 'gap' : 'none', lack.changes || ''));
  }
  return {
    title: MTG_TITLES.progress,
    lead: lack.changes ? mtgChangedLabel_(s) + 'は' + MTG_OPTIONS.gapText + '（' + lack.changes + '）。'
      : (moves.length ? mtgChangedLabel_(s) + 'は ' + moves.length + '件' + mtgUpdatedNote_(s) + '。'
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
  const columns = mtgColumns_([140, 62, 104, 62, 98, 122, 72], 4);
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', '期待値／\n想定売上', 'ステータス\n（滞留日数）', '受注月', '活動状況', '活動における課題',
    '該当事項'])];
  const G = MTG_OPTIONS.gapText;
  const lack = s.gaps || {};
  s.focus.forEach(d => rows.push(mtgRow_(columns, [
    mtgDealCell_(d, columns[0]),
    // 期待値は大きな緑の字。空なら「未入力」を本文の大きさ・グレーで（数のように見せない）。列が無ければ灰色のセルに「データなし」
    lack.weighted && lack.amount ? mtgGapCell_({lines: 2, align: 'CENTER'})
      : {text: (lack.weighted ? G : mtgAmount_(d.weighted)) + '\n' + (lack.amount ? G : mtgAmount_(d.amount)),
        parts: [lack.weighted || d.weighted == null ? {color: c.muted} : {bold: true, color: c.greenDark, size: 11}, lack.amount ? {color: c.muted} : {}],
        lines: 2, fill: lack.weighted || lack.amount ? c.gap : c.white, align: 'CENTER'},
    {text: mtgStatusLines_(d.status || MTG_OPTIONS.blankText, columns[2]) + '\n' + (d.days == null ? '記録なし' : d.days + '日' + (d.atLeast ? '以上' : '')),
      parts: mtgStatusLines_(d.status || MTG_OPTIONS.blankText, columns[2]).split('\n').map(() => ({})).concat([{color: c.muted}]),
      lines: 3, fill: c.white, align: 'CENTER'},
    mtgFieldCell_(s, 'orderMonth', mtgMonthLabel_(d.orderMonthText), {fill: c.white, align: 'CENTER'}),
    mtgFieldCell_(s, 'activity', d.activity, {lines: 3, fill: c.white}),
    mtgFieldCell_(s, 'issue', d.issue, {lines: 3, fill: c.white}),
    {text: mtgFlags_(d).length ? '' : MTG_OPTIONS.noneText, fill: c.white, align: 'CENTER'}
  ], Math.max(44, mtgPillsHeight_(mtgFlags_(d).length)))));
  const el = [];
  if (s.focus.length) {
    const table = {type: 'table', x: S.x, y: S.top, columns, rows};
    el.push(table);
    const width = mtgPillWidth_(['停滞', '受注月超過', '延期9回'], columns[6]);
    s.focus.forEach((d, i) => el.push.apply(el, mtgPills_(table, i + 1, 6, mtgFlags_(d).map(flag => ({text: flag, tone: flag === '停滞' ? 'stall' : 'note'})),
      width)));
  } else {
    el.push.apply(el, mtgEmptyTable_(columns, rows[0], 'none', ''));
  }
  const share = k.weighted > 0 ? Math.round(mtgSum_(s.focus, 'weighted') / k.weighted * 100) : 0;
  return {
    title: MTG_TITLES.focus,
    // 期待値の列が無ければ、想定売上の大きい順（並べ方と、期待値はデータなしであることを書く）
    lead: !s.focus.length ? '進行中案件なし。'
      : lack.weighted ? '期待値は' + G + '（' + lack.weighted + '）のため、進行中案件のうち想定売上上位 ' + s.focus.length + '件。'
        : '進行中案件のうち期待値上位 ' + s.focus.length + '件（進行中案件の期待値合計の ' + share + '%）。',
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
  const lack = s.gaps || {};
  const columns = mtgColumns_([120, 66, 118, 62, 72, 86, 74, 62], 5);
  return mtgPagedTable_({
    title: MTG_TITLES.changes,
    lead: lack.changes ? mtgChangedLabel_(s) + 'は' + MTG_OPTIONS.gapText + '（' + lack.changes + '）。'
      : mtgChangedLabel_(s) + 'は ' + s.kpi.changed + '件' + mtgUpdatedNote_(s) + '。' +
        (s.unmatched ? '\n案件と照合できなかった変更記録 ' + s.unmatched + '件は集計対象外。' : '') +
        (s.unlogged ? '\nうち ' + s.unlogged + '件は変更の記録が無い（前回の会議時点のコピーとの比較で検出）。' : '') +
        (s.copyError ? '\n前回の会議時点のコピーを開けないため、変更の記録から集計。' : ''),
    columns,
    head: ['得意先／案件名', '区分', 'ステータス\n前回▶今回', '期待値\n前回▶今回', '受注月\n前回▶今回', '活動状況（最新）', '更新項目', '責任者'],
    rows: s.changes.map(ch => [
      mtgDealCell_(ch, columns[0], MTG_SLIDE.appendixSize),
      {text: '', fill: c.white},
      {text: mtgStatusMove_(ch), lines: 3, fill: c.white},
      mtgFieldCell_(s, 'weighted', mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'orderMonth', mtgMonthMove_(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      ch.removed ? {text: '（削除）', lines: 2, fill: c.white} : mtgFieldCell_(s, 'activity', ch.activity, {lines: 2, fill: c.white}),
      {text: ch.category === 'new' ? '（新規追加）' : ch.columns.map(mtgShortHeader_).join('・') || (ch.removed ? '（削除）' : MTG_OPTIONS.noneText), lines: 2,
        fill: c.white},
      mtgFieldCell_(s, 'owner', ch.owner, {lines: 2, fill: c.white, align: 'CENTER'})
    ]),
    // 変化が無い・記録が無いときも、表の見出しの行を置き、その下に「該当なし」・「データなし」と理由
    empty: lack.changes ? {kind: 'gap', text: lack.changes} : {kind: 'none'},
    pills: {col: 1, labels: i => [mtgCategoryPill_(s.changes[i].category)], width: mtgCategoryPillWidth_(columns[1])}
  });
}

/** Appendix ② 停滞案件（一覧）。同じステータスの日数と、最新の活動状況（前回以降に更新があるかが分かるように） */
function mtgAppendixStalled_(s) {
  const c = MTG_OPTIONS.colors;
  const k = s.kpi;
  const columns = mtgColumns_([120, 44, 44, 88, 62, 60, 34, 60, 84, 64], 9);
  const lack = s.gaps || {};
  const G = MTG_OPTIONS.gapText;
  const blankAmount = lack.amount ? 0 : s.stalled.filter(d => d.amount == null).length;   // 想定売上が空（合計に含めない）
  // 停滞をまだ数えられない（記録の始まり・基準から stallDays 日たっていない）ときは、その旨と、数えられるようになる日
  const pending = lack.stalled ? lack.stalled + '。' : '';
  return mtgPagedTable_({
    title: MTG_TITLES.stalled,
    lead: pending && !s.stalled.length ? '停滞案件は、' + pending
      : '進行中案件のうち、同一ステータスに ' + MTG_OPTIONS.stallDays + '日以上滞留している案件は ' + s.stalled.length + '件（想定売上合計 ' +
        (lack.amount ? G : mtgAmount_(k.stalledAmount) + '百万円') + (blankAmount ? '。想定売上未入力の ' + blankAmount + '件を除く' : '') + '）。\n' +
        mtgQuietLabel_(s) + ' ' + k.stalledQuiet + '件。' + (pending ? 'ほかの案件は、' + pending : ''),
    columns,
    head: ['得意先／案件名', '想定\n売上', '期待値', 'ステータス', '滞留日数', '受注月', '延期\n回数', '責任者', '活動における課題', '活動状況（最新）'],
    rows: s.stalled.map(d => [
      mtgDealCell_(d, columns[0], MTG_SLIDE.appendixSize),
      mtgFieldCell_(s, 'amount', mtgAmount_(d.amount), {fill: c.white, align: 'END'}),
      mtgFieldCell_(s, 'weighted', mtgAmount_(d.weighted), {fill: c.white, align: 'END'}),
      {text: mtgStatusLines_(d.status || MTG_OPTIONS.blankText, columns[3], MTG_SLIDE.appendixSize), lines: 2, fill: c.white},
      {text: d.days + '日' + (d.atLeast ? '以上' : ''), fill: c.white, align: 'CENTER'},
      mtgFieldCell_(s, 'orderMonth', mtgMonthLabel_(d.orderMonthText), {fill: c.white, align: 'CENTER'}),
      {text: d.slips ? d.slips + '回' : '', fill: c.white, align: 'CENTER'},
      mtgFieldCell_(s, 'owner', d.owner, {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'issue', d.issue, {lines: 2, fill: c.white}),
      mtgFieldCell_(s, 'activity', d.activity, {lines: 2, fill: c.white})
    ]),
    empty: lack.stalled ? {kind: 'gap', text: lack.stalled} : {kind: 'none'}
  });
}

/**
 * Appendix ③ 入力内容のチェック（責任者別）：責任者 × 種類の件数・試作から解消した件数。種類はいつもすべてを列に出し（資料の型を変えない）、
 * チェックできなかった種類（列・記録が無い。summary.alertGaps）は灰色の列にして、表の下に「データなし」と理由を書く。確認事項が無ければ合計の行だけ。
 * 種類の説明は置かない（案件ごとの一覧に内容を書く）
 */
function mtgAppendixOwners_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const lack = s.alertGaps || {};
  const types = MTG_ALERT_TYPES.slice();
  const groups = new Map();
  s.alerts.forEach(a => {
    const owner = a.owner || MTG_OPTIONS.blankText;
    if (!groups.has(owner)) groups.set(owner, []);
    groups.get(owner).push(a);
  });
  const skipped = types.filter(t => lack[t]);
  const notes = [];
  if (!s.alerts.length) notes.push({kind: 'none', text: '', lines: 0});
  if (skipped.length) {
    notes.push({kind: 'gap', text: '灰色の列は、チェックできなかった項目\n' + skipped.map(t => (MTG_ALERT_LABELS[t] || t) + '：' + lack[t]).join('\n'),
      lines: skipped.length + 1});
  }
  const noteH = note => mtgTextHeight_(1, 10.5) + (note.lines || 0) * 9 * S.lineFactor;   // 「データなし」の行・理由の行・上下の余白
  const notesH = notes.reduce((sum, note) => sum + noteH(note) + 6, 0);

  const ownerW = 112;
  const totalW = 40;
  const typeW = (S.w - ownerW - totalW) / types.length;
  const columns = [ownerW].concat(types.map(() => typeW), [totalW]);
  const head = mtgHeadRow_(columns, ['責任者'].concat(types.map(t => MTG_ALERT_SHORT[t] || t), ['計']), 7.5);
  head.cells[head.cells.length - 1].fill = c.greenDeep;
  types.forEach((t, i) => { if (lack[t]) head.cells[i + 1].fill = c.gapLine; });
  // 責任者は名前の順（件数の順に並べると順位表に見えるため）。表と下の枠がページに入らないときは、件数の少ない責任者を「その他 N名」にまとめる
  const byCount = Array.from(groups.keys()).sort((a, b) => groups.get(b).length - groups.get(a).length);
  const build = limit => {
    const shown = byCount.slice(0, byCount.length > limit ? limit - 1 : limit);
    const lines = mtgOwnerOrder_(shown).map(owner => ({label: owner, list: groups.get(owner)}));
    if (byCount.length > limit) {
      const rest = byCount.slice(limit - 1);
      lines.push({label: 'その他 ' + rest.length + '名', list: [].concat.apply([], rest.map(owner => groups.get(owner)))});
    }
    lines.push({label: '合計', list: s.alerts, total: true});
    return [head].concat(lines.map(line => {
      const fill = line.total ? c.greenSoft : c.white;
      return mtgRow_(columns, [{text: line.label, lines: 2, bold: true, color: c.ink, fill: c.greenSoft, align: 'CENTER'}]
        .concat(types.map(t => {
          if (lack[t]) return mtgGapCell_({align: 'CENTER'}, '—');
          const n = line.list.filter(a => a.type === t).length;
          return {text: n || line.total ? String(n) : '', fill, align: 'CENTER', bold: !!line.total};
        }), [{text: String(line.list.length), fill, align: 'CENTER', bold: true, color: c.ink}]), 20);
    }));
  };
  let limit = MTG_OPTIONS.mainRows.owners;
  let rows = build(limit);
  const height = list => list.reduce((sum, r) => sum + r.h, 0);
  while (limit > 2 && S.top + height(rows) + notesH > S.bottom) rows = build(--limit);
  const el = [{type: 'table', x: S.x, y: S.top, columns, rows}];
  let y = S.top + height(rows) + 6;
  notes.forEach(note => {
    el.push(mtgEmptyBox_(S.x, y, S.w, noteH(note), note.kind, note.text, note.lines));
    y += noteH(note) + 6;
  });
  const r = s.resolved;
  return [{
    title: MTG_TITLES.owners,
    lead: '入力内容の確認事項は ' + s.alerts.length + '項目' + (s.alerts.length ? '（' + k.alertDeals + '案件・責任者 ' + k.owners + '名）' : '') + '。' +
      (r ? '試作時（' + Utilities.formatDate(new Date(r.at), Session.getScriptTimeZone(), 'M/d') + '）から ' + r.fixed + '項目解消、' + r.added + '項目追加。' : '') +
      (skipped.length ? '\nチェックできなかった項目は ' + skipped.length + '種類（灰色の列）。' : ''),
    empty: s.alerts.length ? undefined : 'none',
    elements: el
  }];
}

/** 責任者の並び：名前の順（localeCompare の日本語の順。漢字は JIS の順で、読みの50音順ではない。責任者が空は最後） */
function mtgOwnerOrder_(owners) {
  const blank = MTG_OPTIONS.blankText;   // 責任者が空（未入力）は最後
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
    const owner = a.owner || MTG_OPTIONS.blankText;
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
        const label = MTG_ALERT_LABELS[a.type] || a.type;
        // 1行に入らない名前は、言葉の切れ目で2行に（「デリバリー担当未設／定」にしない）
        const t = mtgWrap_(label, columns[2], size, 99).length > 1 && MTG_ALERT_SHORT[a.type] ? MTG_ALERT_SHORT[a.type].split('\n') : mtgWrap_(label, columns[2], size, 2);
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
        {text: list[0].dealKey === '' ? '—' : mtgAmount_(list[0].amount), fill: c.white, align: 'END'}   // 得意先の予定は案件の金額が無い（対象外）
      ]);
    });
  });
  return mtgPagedTable_({
    title: MTG_TITLES.alerts,
    lead: '入力内容の確認事項は ' + s.alerts.length + '項目' + (s.alerts.length ? '（' + s.kpi.alertDeals + '案件）' : '') + '。',
    columns,
    head: ['責任者', '得意先／案件名', '項目', '内容', '想定\n売上'],
    rows,
    empty: {kind: 'none'}
  });
}

/**
 * Appendix の表を、入りきる行ずつのページに分ける（見出しの行は各ページに）。
 * 行が無くてもページは作り（資料の型を変えない）、表の見出しの行の下に spec.empty（{kind: 'gap'|'none', text}。mtgEmptyBox_）を置く。
 * spec = {title, lead, columns, head, rows: [[セル]], empty, pills: {col, labels: 行の番号 → [{text, tone}]}}（pills はセルに重ねるラベル。mtgPills_）。
 * 行が無いページには empty: 'gap'|'none' を付ける（扉の一覧に出す）
 */
function mtgPagedTable_(spec) {
  const S = MTG_SLIDE;
  const head = mtgHeadRow_(spec.columns, spec.head, S.appendixSize);
  if (!spec.rows.length) {
    const empty = spec.empty || {kind: 'none'};
    return [{title: spec.title, lead: spec.lead, empty: empty.kind, elements: mtgEmptyTable_(spec.columns, head, empty.kind, empty.text || '')}];
  }
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

/**
 * ステータスの名前を表のセルに入れる形に：1行に入らなければ、率の「(70%)」の前で改行する（「B-クロージン／グ中(70%)」のように言葉の途中で切らない）
 */
function mtgStatusLines_(status, width, size) {
  const text = String(status || '');
  if (mtgWrap_(text, width, size || MTG_SLIDE.bodySize, 99).length <= 1) return text;
  const at = text.search(/[（(]/);
  return at > 0 ? text.slice(0, at) + '\n' + text.slice(at) : text;
}

/** ステータスの動き（「提案中 ▶ 受注」。変わっていなければ今の値。空は「未入力」、前回に無かった案件は「（新規）」から） */
function mtgStatusMove_(ch) {
  const blank = v => v || MTG_OPTIONS.blankText;
  if (ch.removed) return blank(ch.statusFrom) + ' ▶ （削除）';
  if (!ch.base) return '（新規） ▶ ' + blank(ch.statusTo);
  if (credNormalize_(ch.statusFrom) === credNormalize_(ch.statusTo)) return blank(ch.statusTo);
  return blank(ch.statusFrom) + ' ▶ ' + blank(ch.statusTo);
}

/** 金額の動き（「10 ▶ 15」「未入力 ▶ 15」。変わっていなければ今の値。空と 0 は別の値として比べる） */
function mtgValueMove_(from, to, ch) {
  if (ch.removed) return mtgAmount_(from) + ' ▶ （削除）';
  if (!ch.base || (from == null ? null : from) === (to == null ? null : to)) return mtgAmount_(to);
  return mtgAmount_(from) + ' ▶ ' + mtgAmount_(to);
}

/** 受注月の動き（「2026/11 ▶ 2026/12」）。月として読めれば月で比べる（スナップショットとシートで書き方が違っても、同じ月なら動きにしない） */
function mtgMonthMove_(ch) {
  if (ch.removed) return mtgMonthLabel_(ch.monthFrom);
  if (!ch.base || mtgMonthLabel_(ch.monthFrom) === mtgMonthLabel_(ch.monthTo)) return mtgMonthLabel_(ch.monthTo);
  return mtgMonthLabel_(ch.monthFrom) + ' ▶ ' + mtgMonthLabel_(ch.monthTo);
}

/** 受注月の表示（「2026年11月」「2026/11/01」→「2026/11」。読めなければそのまま。空は「未入力」） */
function mtgMonthLabel_(text) {
  const m = mtgMonth_(text);
  return m ? m.y + '/' + m.m : (String(text || '').trim() || MTG_OPTIONS.blankText);
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

/**
 * 表の中身の代わりに置く枠。kind：'gap' 作れなかった（記録・列が無いなど。灰色の地に点線の枠で「データなし」と理由）・
 * 'none' 該当する案件が無い（白の地に細い線で「該当なし」と text）。text は改行で行を分けられる（合わせて max 行まで。既定は3行）
 */
function mtgEmptyBox_(x, y, w, h, kind, text, max) {
  const c = MTG_OPTIONS.colors;
  const gap = kind === 'gap';
  const head = gap ? MTG_OPTIONS.gapText : MTG_OPTIONS.noneText;
  const lines = text ? mtgWrap_(text, w, 9, max || 3) : [];
  const all = [head].concat(lines).join('\n');
  return mtgBox_(x, y, w, h, all, {size: 9, color: gap ? c.dark : c.muted, align: 'CENTER', fill: gap ? c.gap : c.white,
    line: gap ? c.gapLine : c.line, lineWeight: 0.75, dash: gap, shape: 'RECTANGLE', role: gap ? 'gap' : 'empty',
    spans: [{start: 0, end: head.length, size: 10.5, bold: true}]});
}

/** 表の見出しの行だけを置き、その下に mtgEmptyBox_ を置く（Appendix・本編の表が空のときも、表の形を残す）。head は見出しの行（mtgHeadRow_） */
function mtgEmptyTable_(columns, head, kind, text) {
  const S = MTG_SLIDE;
  const w = columns.reduce((a, b) => a + b, 0);
  const lines = text ? mtgWrap_(text, w, 9, 3).length : 0;
  return [
    {type: 'table', x: S.x, y: S.top, columns, rows: [head]},
    mtgEmptyBox_(S.x, S.top + head.h + 8, w, Math.max(48, mtgTextHeight_(1, 10.5) + lines * 9 * S.lineFactor + 8), kind, text)
  ];
}

/**
 * 作れなかった項目の表のセル（灰色の地に「データなし」。text を渡せばその文）。cell の大きさなどはそのまま。
 * 「データなし」だけのときは、狭い列でも切れないように2行まで（「データ」「なし」）
 */
function mtgGapCell_(cell, text) {
  const c = MTG_OPTIONS.colors;
  const out = Object.assign({}, cell, {text: text || MTG_OPTIONS.gapText, fill: c.gap, color: c.muted, bold: false},
    text ? {} : {lines: Math.max(2, cell.lines || 1)});
  delete out.parts;
  delete out.spans;
  return out;
}

/** 表のセル：シートに列が無い項目（summary.gaps[key]）は灰色の「データなし」、そうでなければ text（空なら「未入力」） */
function mtgFieldCell_(s, key, text, cell) {
  if (s.gaps && s.gaps[key]) return mtgGapCell_(cell);
  return Object.assign({}, cell, {text: text || MTG_OPTIONS.blankText});
}

/** 得意先（1行目・太字）と案件名（2行目）のセル（1列目なので薄い緑）。size は文字の大きさ（既定は bodySize） */
function mtgDealCell_(item, width, size) {
  const c = MTG_OPTIONS.colors;
  const pt = size || MTG_SLIDE.bodySize;
  // 2行目：案件名。得意先の予定（予定日超過。案件に結び付かない）は「（得意先の予定）」、案件名が空なら「（案件名未入力）」
  const project = item.project || (item.dealKey === '' ? '（得意先の予定）' : '（案件名未入力）');
  return {text: mtgWrap_(item.customer || '（得意先未入力）', width, pt, 1)[0] + '\n' + mtgWrap_(project, width, pt, 1)[0],
    size: pt, parts: [{bold: true}, {}], lines: 2, color: c.ink, fill: c.greenSoft};
}

/**
 * 表の1行。cells: [{text, lines, size, bold, color, fill, align, parts}]（文字は列の幅で折り返し、lines 行（既定は1）を超えたら「…」）。
 * parts は改行で分けた部分ごとの書式 [{bold, color, size}]（スライドでは spans にする）。parts の代わりに spans を直接渡してもよい（1行の中で大きさを変えるとき）。
 * 行の高さは、いちばん高いセルに合わせる（min 以上。mtgTextHeight_）。Google スライドでは、中身がこの高さより高ければ行が伸びるので、少し大きめに見積もる
 */
function mtgRow_(columns, cells, min) {
  let h = min || 0;
  const out = cells.map((cell, i) => {
    const size = cell.size || MTG_SLIDE.bodySize;
    const lines = mtgBalance_(cell.text, columns[i], size, cell.lines || 1);
    const text = lines.join('\n');
    const spans = cell.parts ? mtgPartSpans_(String(cell.text), cell.parts, columns[i], size, lines) : cell.spans;
    h = Math.max(h, mtgTextHeight_(lines, size, spans));
    return Object.assign({}, cell, {text, size}, spans ? {spans} : {});
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
 * 文字1つの幅の見積もり（文字の大きさに対する割合）。BIZ UDPゴシック（プロポーショナル）の幅を、標準と太字の大きいほうで：
 * 半角の英数字・記号は1文字ずつ（MTG_ASCII_WIDTHS。全角の英数字も同じ幅）、句読点 0.6、かっこ・中黒・コロン・半角カナ 0.5、
 * 全角の空白 0.67、小さいかな 0.94、カタカナ 0.96（長音を含む。0.78〜1、平均 0.93）、漢字・ひらがな・そのほか 1
 */
function mtgCharWidth_(ch) {
  const code = ch.codePointAt(0);
  if (code >= 32 && code < 127) return MTG_ASCII_WIDTHS[code - 32];
  if (/[０-９Ａ-Ｚａ-ｚ]/.test(ch)) return MTG_ASCII_WIDTHS[code - 0xFEE0 - 32];
  if (/[、。，．]/.test(ch)) return 0.6;
  if (/[・：；（）「」『』【】〔〕〈〉《》［］｛｝\uff61-\uff9f]/.test(ch)) return 0.5;
  if (ch === '　') return 0.67;
  if (/[±×÷]/.test(ch)) return 0.84;
  if (/[ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ]/.test(ch)) return 0.94;
  if (/[\u30a0-\u30ff]/.test(ch)) return 0.96;
  return 1;
}

/** 文の幅（pt）の見積もり：文字の大きさ size（改行は無いものとして数える） */
function mtgTextWidth_(text, size) {
  return Array.from(String(text == null ? '' : text)).reduce((sum, ch) => sum + mtgCharWidth_(ch), 0) * (size || MTG_SLIDE.bodySize);
}

/**
 * 文を、幅 widthPt の箱（セル）の中に収まる行に分ける（改行はそのまま行の区切り。pt で数える）。max 行を超えたら最後の行を「…」の分をあけて切る。
 * 文字の幅は見積もり（mtgCharWidth_。箱の左右の余白 MTG_SLIDE.padX を引き、さらに MTG_SLIDE.fit の余裕をとる）なので、本物の折り返しより早めに分ける。
 * 行の頭に来てはいけない文字（、。）」など）は、本当の幅に収まるなら前の行の終わりにぶら下げ、収まらなければ前の文字と一緒に次の行へ送る。
 * 半角の語（「2026/12」など）の途中では切らず、語の前の空白で折り返す。「前 ▶ 今」は矢印のあとで折り返す。開きかっこは行の終わりに残さない。
 */
function mtgWrap_(text, widthPt, size, max) {
  const room = (widthPt - MTG_SLIDE.padX) / (size || MTG_SLIDE.bodySize);
  const limit = room * MTG_SLIDE.fit;
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
        out.push(keep.replace(/\s+$/, ''));   // 行の終わりの空白は残さない
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
 * 区切りは、近くの言葉の切れ目に動かす（mtgBreakAt_。「シ／ナリオ」「サポー／ト」のように語の途中で切らない）。
 * 「…」で切れるときは、行の長さはそろえず、区切りだけを動かす
 */
function mtgBalance_(text, widthPt, size, max) {
  const lines = mtgWrap_(text, widthPt, size, max);
  if (lines.length < 2) return lines;
  if (/…$/.test(lines[lines.length - 1])) {
    // 「…」で切れるとき：切る前の行の区切りを言葉の切れ目に動かしてから、最後の行を切る
    if (String(text).indexOf('\n') >= 0) return lines;
    const full = mtgWrap_(text, widthPt, size, 99);
    const moved = mtgBreakAt_(text, full, widthPt, size);
    if (moved === full) return lines;
    return moved.slice(0, lines.length - 1).concat(mtgWrap_(moved.slice(lines.length - 1).join(''), widthPt, size, 1));
  }
  const units = line => Array.from(line).reduce((sum, ch) => sum + mtgCharWidth_(ch), 0);
  const even = () => {
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
  };
  const out = even();
  return String(text).indexOf('\n') < 0 ? mtgBreakAt_(text, out, widthPt, size) : out;
}

/**
 * 折り返した文（lines）の区切りを、それぞれ前後4文字の中の言葉の切れ目に動かす（上の区切りから順に。動かしても両側の行が1行に入るときだけ。
 * 空白で分けた区切りがあれば、そのまま）。
 * 切れ目の良さ：句読点・閉じかっこのあと ＞ カタカナ・英数字の語の前後、ひらがなから漢字に変わるところ ＞ 漢字・ひらがなの続き ＞ 語の途中・送りがなの前。
 * 行の頭に来ない文字の前、開きかっこのあとでは切らない
 */
function mtgBreakAt_(text, lines, widthPt, size) {
  const chars = Array.from(String(text).replace(/[ \t]+/g, ' ').trim());
  const cuts = [];
  let pos = 0;
  for (let k = 0; k < lines.length; k++) {
    const part = Array.from(lines[k]);
    if (chars.slice(pos, pos + part.length).join('') !== lines[k]) return lines;
    pos += part.length;
    if (k < lines.length - 1) {
      if (chars[pos] === ' ') return lines;
      cuts.push(pos);
    }
  }
  if (pos !== chars.length) return lines;
  const kind = ch => (/[0-9A-Za-z.,:\/%+\-０-９Ａ-Ｚａ-ｚ]/.test(ch) ? 'alnum' : /[\u30a0-\u30ff]/.test(ch) ? 'kata'
    : /[\u3040-\u309f]/.test(ch) ? 'hira' : /[\u4e00-\u9fff々]/.test(ch) ? 'kanji' : 'other');
  const score = i => {   // chars[i - 1] と chars[i] の間
    const a = chars[i - 1];
    const b = chars[i];
    if (/[、。，．,.)）」』】・ー！？!?%％:：;；ぁぃぅぇぉっゃゅょァィゥェォッャュョ]/.test(b) || /[（「『【(]/.test(a) || a === ' ' || b === ' ') return -1;
    if (/[、。，．）」』】・]/.test(a)) return 3;
    const ka = kind(a);
    const kb = kind(b);
    if (ka === kb) return ka === 'kanji' || ka === 'hira' ? 1 : 0;
    return ka === 'kanji' && kb === 'hira' ? 0 : 2;
  };
  const fits = (from, to) => to > from && mtgWrap_(chars.slice(from, to).join(''), widthPt, size, 99).length === 1;
  // 2回まわす（下の区切りを動かしてから、上の区切りを動かせることがある）
  [0, 1].forEach(() => cuts.forEach((cut, k) => {
    const from = k ? cuts[k - 1] : 0;
    const to = k < cuts.length - 1 ? cuts[k + 1] : chars.length;
    let best = {q: cut, score: score(cut)};
    for (let q = Math.max(from + 1, cut - 4); q <= Math.min(to - 1, cut + 4); q++) {
      const sc = score(q);
      const better = sc > best.score || (sc === best.score && Math.abs(q - cut) < Math.abs(best.q - cut));
      if (q !== cut && better && fits(from, q) && fits(q, to)) best = {q, score: sc};
    }
    cuts[k] = best.q;
  }));
  return cuts.concat([chars.length]).map((cut, k) => chars.slice(k ? cuts[k - 1] : 0, cut).join(''));
}

/**
 * セル・箱の高さ（pt）の見積もり：折り返したあとの行 lines（または行の数）・文字の大きさ size（上下の余白を含む）。
 * 行の送りは、その行のいちばん大きな文字に合わせる（spans [{start, end, size}] で大きさを変えた文字。位置は lines を改行でつないだ文字で数える）
 */
function mtgTextHeight_(lines, size, spans) {
  const S = MTG_SLIDE;
  const base = size || S.bodySize;
  const list = typeof lines === 'number' ? Array.from({length: lines}, () => '') : lines;
  let offset = 0;
  const total = list.reduce((sum, line) => {
    const end = offset + line.length;
    const big = (spans || []).filter(sp => sp.size && sp.start <= end && sp.end > offset).reduce((m, sp) => Math.max(m, sp.size), base);
    offset = end + 1;
    return sum + big * S.lineFactor;
  }, 0);
  return total + S.padY;
}

/* ---------------- スライド（16:9）に書く ---------------- */

/*
 * スライドの部品（位置・大きさは pt。スライドは MTG_SLIDE.width × height）
 * - 箱：{type: 'box', x, y, w, h, text, size, bold, color, font（既定は MTG_SLIDE.font）, fill, alpha（塗りの濃さ 0〜1）, line, lineWeight, dash（枠を点線に）,
 *        align: 'START'|'CENTER'|'END',
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
    if (slide.template) {
      // ひな形が無いときの表紙・扉：背景は緑一色にしておき、あとで画像にする（mtgCoverBackground_）
      requests.push({updatePageProperties: {objectId: pageId, pageProperties: {pageBackgroundFill: {solidFill: {color: {rgbColor:
        mtgRgb_(MTG_OPTIONS.colors.cover)}}}}, fields: 'pageBackgroundFill.solidFill.color'}});
      elements = mtgTemplateFallback_(slide.template).concat(elements);
    }
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
        ? {outlineFill: {solidFill: {color: {rgbColor: mtgRgb_(el.line)}}}, weight: pt(el.lineWeight || 2.25), dashStyle: el.dash ? 'DASH' : 'SOLID'}
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
 * ひな形が無いときの表紙・扉の文字：ひな形と同じ位置（ひな形の 960 × 540pt を 0.75 倍）に、見出し・その下の文字・日付・名前を白で描く。
 * 背景（緑のグラデーションと大きな「RX」）は、ページの背景に画像で付ける（mtgCoverBackground_。付けるまでは緑一色。mtgSlideRequests_）
 */
function mtgTemplateFallback_(template) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const line = (y, h, text, size, bold) => mtgBox_(S.x, y, S.w, h, mtgWrap_(text, S.w, size, 1)[0], {size, bold: !!bold, color: c.white, align: 'CENTER'});
  return [
    line(170, 46, template.title, 30, true),
    line(226, 26, template.subtitle, 15, true),
    line(326, 20, template.date, 12),
    line(344, 20, template.name, 12)
  ].filter(el => el.text);
}

/* ---------------- 表紙・扉の背景（ひな形が無いとき） ---------------- */

/*
 * 表紙・扉の背景の画像：左上から右下へ明るくなる緑（いただいた表紙のひな形と同じ色）に、大きな「RX」を濃い緑で重ねる（はみ出した部分は切る）。
 * 画像はスクリプトの中で描いて PNG にする（ファイルを別に置かない）。
 */
const MTG_COVER = {
  width: 1280,                 // 画像の大きさ（px。16:9）
  height: 720,
  from: '#3A7D22',             // 左上の色（0〜stop は from のまま）
  to: '#8CD872',               // 右下の色
  stop: 0.26,
  ink: '#275316',              // RX の色
  // RX の字形（Liberation Sans Bold。SIL Open Font License）：文字の大きさ size（px）・左端 x・ベースライン baseline（画像の px）。
  // 輪郭はフォントの単位（upm = 文字の大きさ）で、点は「x,y」、点の間は空白、輪郭の間は「|」。X は R の送り幅（1479）だけ右にずらしてある
  glyph: {size: 1142.67, x: -109.4, baseline: 756.2, upm: 2048},
  contours: '1105,0 778,535 432,535 432,0 137,0 137,1409 841,1409 902,1407 960,1402 1014,1394 1064,1382 1111,1367 1154,1348 1194,1326 1230,1300 1266,1268 1297,1231 1322,1190 1342,1146 1356,1097 1364,1045 1367,989 1365,941 1358,895 1346,851 1330,810 1309,770 1283,734 1253,700 1220,670 1184,644 1145,623 1102,605 1056,592 1437,0|1070,977 1065,1031 1049,1076 1022,1114 985,1143 937,1163 879,1176 810,1180 432,1180 432,764 818,764 865,766 908,773 945,784 978,800 1006,820 1034,851 1054,887 1066,929|2517,0 2163,561 1809,0 1497,0 1985,741 1538,1409 1850,1409 2163,911 2476,1409 2786,1409 2358,741 2827,0'
};

/**
 * 表紙・扉（pageIds）のページの背景に、MTG_COVER の画像を付ける（SlidesApp。Slides API ではファイルの画像を背景にできないため）。
 * 付けられなければ、緑一色の背景のまま続ける（資料は作る）
 */
function mtgCoverBackground_(presentationId, pageIds) {
  if (!pageIds.length) return;
  try {
    const deck = SlidesApp.openById(presentationId);
    const image = mtgCoverImage_();
    pageIds.forEach(id => {
      const page = deck.getSlideById(id);
      if (page) page.getBackground().setPictureFill(image);
    });
    deck.saveAndClose();
  } catch (error) {
    console.warn('表紙・扉の背景の画像を付けられませんでした（緑一色のままにします）：' + (error && error.message || error));
  }
}

/** 表紙・扉の背景の画像（PNG の Blob）。描くのは実行ごとに1回（いくつかのシートの資料を作るときも同じ画像を使う） */
function mtgCoverImage_() {
  if (!MTG_COVER_CACHE.bytes) MTG_COVER_CACHE.bytes = Array.from(mtgCoverPng_(), b => (b > 127 ? b - 256 : b));
  return Utilities.newBlob(MTG_COVER_CACHE.bytes, 'image/png', 'meeting-cover.png');
}

/* 描いた背景の画像（実行の間だけ覚えておく） */
const MTG_COVER_CACHE = {bytes: null};

/**
 * MTG_COVER の画像を描いて PNG のバイト列（Uint8Array）にする。RX は輪郭を横の線で塗る（1px を縦に4本・横は面積で数えて、縁をなめらかに）
 */
function mtgCoverPng_() {
  const C = MTG_COVER;
  const W = C.width;
  const H = C.height;
  const k = C.glyph.size / C.glyph.upm;
  const rgb = hex => [1, 3, 5].map(i => parseInt(String(hex).slice(i, i + 2), 16));
  // 輪郭の辺 [上の x, 上の y, 下の x, 下の y]（横の辺は塗りに関係しないので除く）
  const edges = [];
  C.contours.split('|').forEach(contour => {
    const points = contour.trim().split(/\s+/).map(p => p.split(',').map(Number)).map(([x, y]) => [C.glyph.x + x * k, C.glyph.baseline - y * k]);
    points.forEach((p, i) => {
      const q = points[(i + 1) % points.length];
      if (p[1] !== q[1]) edges.push(p[1] < q[1] ? [p[0], p[1], q[0], q[1]] : [q[0], q[1], p[0], p[1]]);
    });
  });
  const from = rgb(C.from);
  const to = rgb(C.to);
  const ink = rgb(C.ink);
  const SUB = 4;
  const stride = W * 3 + 1;
  const raw = new Uint8Array(stride * H);
  const cover = new Float32Array(W + 1);
  for (let y = 0; y < H; y++) {
    cover.fill(0);
    for (let s = 0; s < SUB; s++) {
      const sy = y + (s + 0.5) / SUB;
      const xs = [];
      edges.forEach(e => { if (sy >= e[1] && sy < e[3]) xs.push(e[0] + (sy - e[1]) * (e[2] - e[0]) / (e[3] - e[1])); });
      xs.sort((a, b) => a - b);
      for (let i = 0; i + 1 < xs.length; i += 2) {   // 内と外を交互に（R の穴は外側の輪郭と逆）
        const a = Math.max(0, Math.min(W, xs[i]));
        const b = Math.max(0, Math.min(W, xs[i + 1]));
        if (b <= a) continue;
        const ia = Math.floor(a);
        const ib = Math.floor(b);
        if (ia === ib) {
          cover[ia] += (b - a) / SUB;
          continue;
        }
        cover[ia] += (ia + 1 - a) / SUB;
        for (let x = ia + 1; x < ib; x++) cover[x] += 1 / SUB;
        cover[ib] += (b - ib) / SUB;
      }
    }
    let o = y * stride;
    raw[o++] = 0;   // 行の頭（PNG のフィルタ：なし）
    for (let x = 0; x < W; x++) {
      const t = Math.max(0, ((x + y + 1) / (W + H) - C.stop) / (1 - C.stop));   // 左上 0 → 右下 1（CSS の 135deg と同じ向き）
      const a = Math.min(1, cover[x]);
      for (let ch = 0; ch < 3; ch++) {
        const base = from[ch] + (to[ch] - from[ch]) * t;
        raw[o++] = Math.round(base + (ink[ch] - base) * a);
      }
    }
  }
  return mtgPng_(W, H, raw);
}

/** PNG（24 ビットのカラー）のバイト列。raw は行ごとに先頭にフィルタの種類（0）を付けた画素（R・G・B） */
function mtgPng_(width, height, raw) {
  const crcTable = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    crcTable.push(c >>> 0);
  }
  const u32 = v => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
  const chunk = (type, data) => {
    const body = Array.from(type, ch => ch.charCodeAt(0)).concat(Array.from(data));
    let crc = 0xFFFFFFFF;
    body.forEach(b => { crc = crcTable[(crc ^ b) & 255] ^ (crc >>> 8); });
    return u32(data.length).concat(body, u32((crc ^ 0xFFFFFFFF) >>> 0));
  };
  const header = u32(width).concat(u32(height), [8, 2, 0, 0, 0]);   // 8 ビット・RGB
  const bytes = [137, 80, 78, 71, 13, 10, 26, 10].concat(chunk('IHDR', header), chunk('IDAT', mtgZlib_(raw, width * 3 + 1)), chunk('IEND', []));
  return Uint8Array.from(bytes);
}

/**
 * zlib の形式で圧縮する（deflate の固定ハフマン符号。同じ並びは、1画素前・1行上・斜め上だけを探す。表紙の画像のように、
 * 同じ色や行をずらした並びが続く画像なら十分に小さくなる）。stride は1行のバイト数
 */
function mtgZlib_(data, stride) {
  const out = [0x78, 0x01];
  let bitBuf = 0;
  let bitLen = 0;
  const bits = (value, count) => {
    bitBuf |= value << bitLen;
    bitLen += count;
    while (bitLen >= 8) {
      out.push(bitBuf & 255);
      bitBuf >>>= 8;
      bitLen -= 8;
    }
  };
  const reversed = (code, length) => {
    let r = 0;
    for (let i = 0; i < length; i++) r = (r << 1) | ((code >> i) & 1);
    return r;
  };
  const symbol = n => {   // 固定ハフマン符号の文字・長さの記号
    if (n < 144) bits(reversed(0x30 + n, 8), 8);
    else if (n < 256) bits(reversed(0x190 + n - 144, 9), 9);
    else if (n < 280) bits(reversed(n - 256, 7), 7);
    else bits(reversed(0xC0 + n - 280, 8), 8);
  };
  const lenBase = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  const lenExtra = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  const distBase = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193,
    12289, 16385, 24577];
  const distExtra = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  const code = (base, value) => {
    let i = base.length - 1;
    while (base[i] > value) i--;
    return i;
  };
  const distances = [3, stride - 3, stride, stride + 3].filter(d => d > 0 && d <= 32768);
  bits(1, 1);   // 最後のブロック
  bits(1, 2);   // 固定ハフマン符号
  const n = data.length;
  let i = 0;
  while (i < n) {
    let best = 0;
    let dist = 0;
    distances.forEach(d => {
      if (d > i) return;
      const max = Math.min(258, n - i);
      let len = 0;
      while (len < max && data[i + len] === data[i + len - d]) len++;
      if (len > best) {
        best = len;
        dist = d;
      }
    });
    if (best >= 3) {
      const lc = code(lenBase, best);
      symbol(257 + lc);
      if (lenExtra[lc]) bits(best - lenBase[lc], lenExtra[lc]);
      const dc = code(distBase, dist);
      bits(reversed(dc, 5), 5);
      if (distExtra[dc]) bits(dist - distBase[dc], distExtra[dc]);
      i += best;
    } else {
      symbol(data[i]);
      i++;
    }
  }
  symbol(256);   // ブロックの終わり
  if (bitLen > 0) out.push(bitBuf & 255);
  let a = 1;
  let b = 0;
  for (let j = 0; j < n; j++) {
    a = (a + data[j]) % 65521;
    b = (b + a) % 65521;
  }
  return out.concat([(b >>> 8) & 255, b & 255, (a >>> 8) & 255, a & 255]);
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
  const ui = mtgUi_();
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
    // ひな形が無ければ、表紙・扉の背景をいただいたひな形と同じ絵柄（緑のグラデーションと RX）の画像にする
    if (!templateId) mtgCoverBackground_(deck.presentationId, slides.map((slide, i) => (slide.cover || slide.divider ? 'mtg_p' + (i + 1) : '')).filter(Boolean));
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
  if (MTG_OPTIONS.folderId) {
    try {
      return DriveApp.getFolderById(MTG_OPTIONS.folderId);
    } catch (error) {
      throw new Error('保存先のフォルダ（ID：' + MTG_OPTIONS.folderId + '）を開けません。実行するアカウントが、そのフォルダに書き込めるか確かめてください' +
        '（MeetingSummary.gs の MTG_OPTIONS.folderId）。');
    }
  }
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
