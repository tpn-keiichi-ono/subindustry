/**
 * 会議サマリ（追跡シートごとの会議資料。16:9 のスライドと PDF）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs・HistorySidebar.gs と同じプロジェクトに置くファイル。
 *
 * 目的
 * - RX Business Meeting（パイプライン会議。月1回）の事前資料。会議の前に配り、参加者（営業の統括・リーダー、BX、サブインダストリーのリーダー）が
 *   通読してから会議に臨む。発表は1サブインダストリー数分（報告＋フィードバック）なので、会議の目的（案件の進捗・受注見込みを確かめ、
 *   優先順位と次のアクションを決める）に向けて、状況を事前に共有しておく。
 * - 資料には、シートと変更の記録から分かる事実だけを書く。論点・決めてほしいこと・支援の要否などの判断は、各サブインダストリーのリーダーが
 *   会議で話すことなので、資料には入れない（資料が判断を先に書くと、読み手を誤った方向に導くおそれがある）。
 * - 資料はサブインダストリー（追跡シート）ごとに1つ。スライドの見出しはどのシートでも同じ（MTG_TITLES）で、状態は見出しの下のリード文に書く。
 *   すべてのシートの資料を作ったときは、それらを1つにまとめた統合版（全サブインダストリー。MTG_OPTIONS.combined）も作る。
 *   使う言葉は docs/DESIGN.md（会議サマリの数え方）で決め、資料の中で同じ意味で使う。
 *
 * 資料の構成（16:9。緑を基調にしたトンマナ：黒の見出し・リード文・緑の見出しの表・区分のラベル・区切り線と出典・ページ番号。
 *   表紙・扉は、いただいたひな形（Google スライド。URL はスクリプトのプロパティ MTG_TEMPLATE_URL。1枚目 緑＝表紙（統合版も）・2枚目 グレー＝扉）を
 *   そのまま使う。登録が無ければ、ひな形と同じ絵柄（グラデーションに大きな「RX」）の画像をスクリプトで描いて背景にし、同じ位置に文字を描く（mtgCoverPng_）
 *   表紙：会議の名前と「事前資料」・サブインダストリー・会議日・データの時点・期間
 *   1 サマリー：期待値と進行中案件の数・前回以降の動き（案件ごとに1回だけ数える区分の件数と案件）・現在の状態（受注月超過・停滞）
 *   2 計画に対する積み上げ：今年度の売上開始月ごとの期待値を確度で塗り分けて積み上げ、計画（MTG_OPTIONS.plan）と比べる
 *   3 パイプラインの変化：前回からの期待値の増減（新規・増額・復帰・減額・受注・失注・保留・削除）・ステータス別・受注月（四半期）別の期待値
 *   4 前回から変化のある案件：ステータス・期待値・受注月・活動状況が変わった案件（今のステータスの順に上位）
 *   5 期待値上位案件の状況・課題：進行中の案件の期待値の大きい順（ステータスの滞留日数・受注月・活動状況・課題・該当事項）
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
 * - 管理者がエディタで createBusinessMtgSummaries() を実行し、作るシートと今回の会議日を入力する（試すときは previewBusinessMtgSummaries()）。
 *   シートごとに Google スライドと PDF をドライブのフォルダに保存する（MTG_OPTIONS.folderId。空なら、このスプレッドシートと
 *   同じフォルダの中に MTG_OPTIONS.folderName のフォルダを作る）。配る（共有する）のは手動。
 * - 「前回から」は、前回の会議（今回の会議日より前で、いちばん新しい会議日）の資料を作った日時から。会議日と作った日時は
 *   シートごとにドキュメントのプロパティ（MTG_MEETINGS_シート名）に残す（同じ会議日で作り直したら、作った日時を上書きする）。
 *   前回の会議の記録が無ければ、MTG_OPTIONS.firstDays 日前から。previewBusinessMtgSummaries() は会議日を記録しないので、何度でも試せる
 *   （ファイル名に「試作」が付く）。
 * - 旧フォーマットからの移行など、案件をまとめて貼り付けたあとは setBusinessMtgBaseline() で今を「基準」として記録する（会議の記録と同じ
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
  meetingName: 'RX Business Meeting',    // 表紙に出す会議の名前（会議のフォルダの名前・チャットの文にも使う）
  formerMeetingNames: ['RX Business MTG'],   // 前の名前（2026/10/17 まで）。この名前の会議のフォルダも、この会議のフォルダとして使う
  // 空欄・該当なしの書き方。資料には推定で値を補わず、空欄は blankText と書く（「—」で済ませない）。該当する案件が無いときは noneText。
  // シートに列が無い・記録が無いなどで作れない（数えられない）ところは gapText と理由を、灰色の枠・セルで書く（資料の型は変えない）
  blankText: '未入力',
  noneText: '該当なし',
  gapText: 'データなし',
  // ファイル名（スライド・PDF）。{date} は会議日（yyyyMMdd）、{name} はサブインダストリー（シート名）。
  // 例：20261023_スーパー・GMS_RX Business Meeting用資料.pdf（試作は末尾に「_試作」）
  fileName: '{date}_{name}_RX Business Meeting用資料',
  folderId: '1A0rkoFLICOK4KV25dUTeRdgoOaM0vJEK',   // 保存先のフォルダの ID（空なら、このスプレッドシートと同じフォルダの中に folderName を作る）
  folderName: '会議サマリ',
  folderKey: 'MTG_FOLDER_ID',         // 作ったフォルダの ID（ドキュメントのプロパティ）
  // 出力先は、スクリプトのプロパティでも指定できる（あればコードの設定より優先。値はフォルダの URL か ID。今の出力先は showOutputFolders() で確かめる）。
  //   save：保存先（folderId の代わり。資料を作る場所・試作の置き場所）
  //   meeting：事前資料の格納先（この中に会議日のフォルダを作る。meetingFolder の代わり。全サブインダストリー・統合版で共通）
  //   copy：時点コピー（会議時点・日次のスプレッドシートのコピー）の保存先（copyFolderId の代わり）
  folderProps: {save: 'MTG_SAVE_FOLDER', meeting: 'MTG_MEETING_FOLDER', copy: 'MTG_COPY_FOLDER'},
  // 前の版のプロパティ（サブインダストリーごとの格納先「MTG_MEETING_FOLDER_<シート名>」）。使わない（showOutputFolders() に、消してよいと出す）
  formerFolderPropPrefix: 'MTG_MEETING_FOLDER_',
  // 会議のフォルダへの格納（本番だけ）：事前資料の格納先（meetingFolder。全サブインダストリー共通の1つのフォルダ）の中の、名前が会議日の
  // meetingFolderDate（yyMMdd）＋「_」で始まり、会議の名前（RX Business Meeting・RX Service Meeting）を含むフォルダ（例：261023_RX Business Meeting_第5回）に、
  // その会議日の全サブインダストリーの事前資料と統合版の PDF を置く（2026/10/26 から。前の版はサブインダストリーごとのフォルダの中に会議日のフォルダを作っていた）。
  // RX Service Meeting の資料（ServiceMeeting.gs）も同じ格納先の中の、自分の会議の名前のフォルダに置く。
  // 会議の名前を含まない前の版のフォルダ（261023_第5回会議）は、RX Business Meeting のフォルダとして使い続ける。
  // 見つからない・格納できないときは保存先（folderId）にだけ残して知らせ、フォルダを用意してから storeBusinessMtgSummaries() で格納する
  meetingFolder: 'https://drive.google.com/drive/folders/1qL4qzi2NrefekmByCDz5RAAJmJRdGWZi',   // 事前資料の格納先（URL か ID）
  meetingFolderDate: 'yyMMdd',
  // 会議日のフォルダが無ければ作る（false なら作らずに知らせる）。名前は meetingFolderName の {date}（meetingFolderDate）・{meeting}（会議の名前）・
  // {n}（第何回）を置き換えたもの。{n} は、格納先の中の、この会議日より前の同じ会議のフォルダ（「yyMMdd_…第N回」）のいちばん大きい N ＋ 1（無ければ 1）
  createMeetingFolder: true,
  meetingFolderName: '{date}_{meeting}_第{n}回',
  // 統合版（全サブインダストリーの資料を1つにまとめた資料）。すべてのサブインダストリーの資料を作ったあとに、createBusinessMtgCombinedSummary() で作る
  // （資料を作る関数とは分ける。1回の実行の6分の上限に収めるため）。
  // 表紙・全サブインダストリーの状況（一覧）のあとに、各サブインダストリーの資料（表紙から Appendix まで）を DIFF_RULES の順に入れる（ページ番号は通し）。
  // 本番は、各サブインダストリーの資料と同じ会議のフォルダ（meetingFolder の中の yyMMdd_RX Business Meeting_第N回）に格納し、インデックスの name の列に書く
  combined: {
    enabled: true,
    name: '全サブインダストリー',     // 表紙の見出し・ファイル名の {name}・インデックスの列の見出し
    coverTone: 'cover'                // ひな形が無いときの表紙の色（'cover' 緑・'combined' 紫。MTG_COVER）
  },
  // Apps Script の1回の実行は6分まで（入力・確かめの画面を出している時間も含む）。作り始めてから limitMs を過ぎそうなら、そこで区切り、
  // 残りは continueAfterMs あとに自動で続ける（時間主導のトリガー continueBusinessMtgSummaries。1回きり）。次の仕事にかかる時間の見積もり：
  // summaryMs 中身を作る（シートと記録を読む）・deckMs 資料1つ・copyMs 時点コピーと記録・storeMs 会議のフォルダへの格納とインデックス1つ
  run: {limitMs: 320000, summaryMs: 30000, deckMs: 35000, copyMs: 60000, storeMs: 10000, continueAfterMs: 60000,
    handler: 'continueBusinessMtgSummaries', jobKey: 'MTG_JOB', resultKey: 'MTG_JOB_RESULT',
    // 前の版の続きのトリガーの関数（2026/10/17 に関数名を改めた）。残っていれば、区切る・やめる・続けるときに一緒に外す
    legacyHandlers: ['continueMeetingSummaries']},
  // 会議資料のインデックス（このスプレッドシートのシート）。会議のフォルダに格納したら、行＝会議日・列＝サブインダストリー（1行目の見出しがシート名）の
  // セルに、PDF のファイルをスマートチップ（chip。Sheets API の拡張サービスが要る）で書く。チップにできないときは、ファイル名のリンクで書く。
  // link を 'folder' にすると会議のフォルダ。シートが無ければ作る。書けなかったときは storeBusinessMtgSummaries() で書き直す。
  // 前の形式（行＝サブインダストリー・列＝会議日）のシートや、書き漏れは rebuildBusinessMtgIndex() で作り直す
  index: {
    sheet: 'INDEX_会議資料',
    dateHeader: '日付',              // A列：会議日（日付の値。order の順）
    order: 'desc',                   // 会議日の行の並び：'desc' 新しい会議日が上（2026/10/17 から）・'asc' 古い会議日が上
    kindHeader: '会議名',            // B列：会議の名前（2026/10/17 に「種別」から改めた）
    formerKindHeaders: ['種別'],     // 前の版の見出し（見つかったら、次に書くときに kindHeader に書き換える）
    kind: 'RX Business Meeting',     // 会議名の列に書く値（同じ日付でも会議名が違う行には書かない）
    formerKinds: ['RX Business MTG'], // 前の版で書いた会議名（同じ会議の行として使い、書くときに今の会議名に書き換える）
    // C列は統合版（combined.name の見出し）、その後ろにサブインダストリーの列（DIFF_RULES の順）
    link: 'file',
    chip: true,
    dateFormat: 'yyyy/m/d',
    fontSize: 9,                     // 字（太字・黒・上下中央）
    rowHeight: 38,                   // 会議日の行の高さ（px）
    tabColor: '#000000',
    widths: {date: 83, kind: 119, combined: 158, sheet: 138}   // 列の幅（px）
  },
  // 時点コピー：スプレッドシート全体を別のスプレッドシート（タイトルは保存した日時と、いつ保存したか）にコピーする。保存するのは、本番の資料を作ったとき
  // （資料作成時）・基準を記録したとき（基準）・毎日（日次。installDailySheetCopy() で始める）。次の資料の「前回からの変化」は、前回の資料作成時の
  // コピーと今のシートを比べて出す（差分追跡が止まっていた間・2,000セルを超える貼り付けの変更も出る）。資料を作る画面では、保存先のフォルダの
  // どの時点コピーと比べるかも選べる。保存先は、会議サマリのフォルダの中の copyFolderName のフォルダ
  copyName: '時点コピー',             // タイトル（時点コピー_2026-10-16 18:30（資料作成時・会議日 2026-10-23）・時点コピー_2026-10-17 05:00（日次））
  formerCopyNames: ['会議時点のコピー'], // 前の名前（2026/10/17 まで）。保存先のフォルダのこの名前のファイルも、時点コピーとして選べる
  copyKinds: {meeting: '資料作成時', daily: '日次', baseline: '基準'},   // タイトルの（）に書く、いつ保存したか
  copyFolderId: '',                   // 時点コピーの保存先のフォルダ（URL か ID）。空なら、会議サマリのフォルダの中の copyFolderName
  copyFolderName: '時点コピー',       // 前の版で作ったフォルダ（会議時点のコピー）は、そのまま使う（copyFolderKey に ID がある）
  // 日次の時点コピー：毎日 hour 時台に保存する（時間主導のトリガー handler。installDailySheetCopy() で始め、uninstallDailySheetCopy() でやめる）。
  // keepDays 日より古い日次のコピーはゴミ箱に移す（資料作成時・基準のコピーは消さない）
  dailyCopy: {hour: 5, keepDays: 90, handler: 'saveDailySheetCopy'},
  copyFolderKey: 'MTG_COPY_FOLDER_ID', // 作ったフォルダの ID（ドキュメントのプロパティ）
  indexLockWaitMs: 30000,             // インデックスに書くときのロック（ドキュメントロック）の待ち時間
  // 表紙・扉のひな形（Google スライド。いただいた .pptx を Google スライドにしたもの）。資料はひな形のコピーに作る（ページの大きさもひな形のまま）。
  // ひな形の URL は、スクリプトのプロパティ templateUrlKey（MTG_TEMPLATE_URL）に入れる（エディタの「プロジェクトの設定」→「スクリプト プロパティ」。
  // setMeetingTemplate() でも入れられる）。.pptx のままでもよい（Google スライドに変換したもの（templateConvertedKey に覚える）を使う。.pptx を
  // 更新したら変換し直す）。スクリプトのプロパティが無ければ templateId（コードの設定）。どれも無ければ、ひな形なしで作る（表紙・扉はコードで描く）
  templateId: '',
  templateUrlKey: 'MTG_TEMPLATE_URL',             // スクリプトのプロパティ：ひな形の URL（Google スライドか .pptx）
  templateConvertedKey: 'MTG_TEMPLATE_CONVERTED', // スクリプトのプロパティ：.pptx を変換した Google スライド {source, updated, id}
  templateKey: 'MTG_TEMPLATE_ID',                 // 前の版の登録（ドキュメントのプロパティ。スクリプトのプロパティが空のときだけ読む）
  // ひな形の何枚目を使うか（2026/10/17 にいただいたひな形：1枚目 緑＝表紙、2枚目 グレー＝扉（Appendix）、3枚目 紫）。
  // 0 か無いページは、扉なら描き（mtgTemplateFallback_ と、同じ絵柄のグレーの画像）、統合版の表紙なら1枚目（緑）を使う。
  // 統合版の表紙は緑（1枚目。2026/10/17 から）。紫の3枚目にするときは combined を 3、MTG_OPTIONS.combined.coverTone を 'combined' にする。
  // service は RX Service Meeting の事前資料（ServiceMeeting.gs）の表紙（紫の3枚目）。無ければ、同じ絵柄の紫の画像を背景に描く
  templatePages: {cover: 1, divider: 2, combined: 0, service: 3},
  // ひな形のページの、置き換える文字。文字の枠ごとに書き換える（見出しとその下の文字が同じ「XXXXX」なので、文字の大きい順・上から順に
  // 見出し → その下の文字に当てる。5文字以上の X は同じとみなす。mtgTemplateTextRequests_）
  templateTexts: {title: 'XXXXX', subtitle: 'XXXXX', date: '2026年 4月 17日', name: 'XXXX'},
  meetingsKeyPrefix: 'MTG_MEETINGS_', // シートごとの会議日と資料を作った日時（ドキュメントのプロパティ。後ろにシート名）
  draftKeyPrefix: 'MTG_DRAFT_',       // シートごとの、会議日ごとの最初の試作の入力内容のチェックの結果（解消した件数を出す用）
  lastKeyPrefix: 'MTG_LAST_',         // 前の版の「前回の作成日時」（会議日の記録が無いときだけ使う）
  keepMeetings: 24,                  // 会議日の記録をいくつ残すか
  keepSlides: false,                 // Google スライドもフォルダに残すか（true なら推進リーダーが書き足せる。false なら PDF だけ。2026/10 から PDF だけ）
  // PDF と同じ名前の PowerPoint（.pptx）も作り、PDF と同じフォルダに置く（会議のフォルダへの格納も一緒に移す。2026/10/29 から。
  // Drive API の書き出し（files.export）で作る。作れなくても PDF は作る）。インデックス・チャットの通知・統合版の確認は PDF を使う
  exportPptx: true,
  // Google チャットへの通知は、サブインダストリーごとのスペースへの2つだけ（下の subChat。2026/10/19 に全体のスペースへの通知はやめた）
  chat: {
    meetingTime: '9:00〜10:30',      // 会議の時刻（会議日程の日付のあとに書く。送るときに時刻を入力すれば、そちら。空なら日付だけ）
    oldKeys: ['MTG_CHAT_WEBHOOK', 'MTG_CHAT_URL'],   // 前の版の全体のスペースの設定（スクリプトのプロパティ）。setBusinessMtgSubChats() で消す
    legacyHandler: 'mtgNotifyScheduled' // 前の版の自動通知（2日前）のトリガーの関数。残っていれば setBusinessMtgSubChats() で外す
  },
  // サブインダストリーごとの Google チャットのスペースへの通知（webhook）。setBusinessMtgSubChats() でシートごとに webhook を登録し、
  // remindBusinessMtgUpdate()（管理フォーマットの更新のお願い）・notifySubBusinessMtgSummaries()（そのサブインダストリーの事前資料のご案内）で送る
  // （どちらもエディタから実行。確かめてから送る。自動では送らない）。
  // 文の {meeting} は会議の名前、{name} はサブインダストリー（シート名）、{date} は会議日程、{deadline} は更新の締め切り、{now} は送る日時、
  // {stats} は今の入力状況（行ごと）、{sheet} は管理フォーマットのシートのリンク、{file} は事前資料のリンク。*〜* は太字（題と見出し）。
  // 2つの文は同じ形（題「【会議】 …（サブインダストリー）」・1行目に会議日程と会議の名前・太字の「▼ 見出し」）
  subChat: {
    webhookKeyPrefix: 'MTG_SUBCHAT_',  // スクリプトのプロパティ：後ろにシート名。webhook の URL（鍵を含むので、コードやシートに書かない）
    deadlineDays: 2,                   // 更新の締め切り：会議日の何日前か（事前資料を作る前）
    deadlineTime: '17:00',             // 更新の締め切りの時刻（空なら日付だけ）
    remind: [
      '*【{meeting}】 管理フォーマットの更新のお願い（{name}）*',
      '{date} の 「{meeting}」 に向けて、事前資料（{name}）を作成します。',
      'お手数ですが、更新期限までに、管理フォーマットの案件の情報を最新にしてください。',
      '',
      '*▼ 更新期限*',
      '{deadline}',
      '',
      '*▼ 更新していただきたい項目*',
      '・案件ステータス・想定売上・期待値・受注月・売上開始月',
      '・活動状況・活動における課題・次回アクション',
      '',
      '*▼ 今の入力状況（{now} 時点）*',
      '{stats}',
      '',
      '*▼ 会議日程*',
      '{date}',
      '',
      '*▼ 管理フォーマット*',
      '{sheet}'
    ],
    summary: [
      '*【{meeting}】 事前資料のご案内（{name}）*',
      '{date} の 「{meeting}」 の事前資料（{name}）が作成されました。',
      '',
      '*▼ レビュアーの方々*',
      'お手数ですが、会議の前にご一読いただきレビュー観点を整理ください。',
      '',
      '*▼ 発表者の方々*',
      '資料内容をご確認の上、発表内容のご準備をお願いいたします。',
      '',
      '*▼ 会議日程*',
      '{date}',
      '',
      '*▼ 事前資料*',
      '{file}'
    ]
  },
  firstDays: 14,                     // 初回（前回の会議の記録が無い）は、何日前からの変化を出すか
  lookbackDays: 180,                 // ステータスの滞留・受注月の延期の回数を数えるために、何日前までの記録を読むか
  stallDays: 45,                     // 停滞：ステータスがこの日数以上同じ進行中の案件
  recentMonths: 2,                   // 期待値上位案件の直近の動き：何か月前（同じ日）からの変更の記録を出すか
  recentMax: 8,                      // 期待値上位案件の直近の動き：1案件に出す動きの数（新しい順。残りは「ほか N件」）
  // RXシナリオ構成別の期待値（createBusinessMtgScenarioSummary()。構成が未確定のあいだの別の資料）。複数のシナリオを選んだ案件は、
  // 期待値を選んだ数で等分して配分する（按分。合計は期待値合計と一致）。あわせて、そのシナリオを含む案件の期待値をそのまま足した「関与」（重複あり）も出す
  scenario: {
    title: 'RXシナリオ構成別の期待値',
    fileName: '{date}_RXシナリオ構成別の期待値_試作',
    unsetLabel: '未設定',
    combos: 15                       // 組み合わせの表に出す数（期待値の大きい順）
  },
  noUpdateDays: 30,                  // 入力内容のチェック「更新なし」：最終更新日時がこの日数より前の進行中の案件
  overdueDays: 90,                   // 入力内容のチェック「予定日超過」：予定日を過ぎてこの日数以内のもの（古いものは出さない）
  fiscalStartMonth: 4,               // 期の始まりの月（受注月別の期待値の四半期・計画に対する積み上げの今年度）
  // 計画に対する積み上げ（本編2枚目。統合版は全サブインダストリーの状況の次）：今年度の売上開始月ごとの期待値を、確度で塗り分けて積み上げ、計画と比べる。
  // amount はサブインダストリーごとの計画（百万円。33シナリオ ダッシュボードの計画と同じ）。plans でシートごとに変えられる（{'スーパー・GMS': 500}）。
  // 統合版の計画は、収録したサブインダストリーの計画の合計。確度はステータスの「-」の前（「C-提案中(50%)」なら C。ダッシュボードと同じ）。
  // ranks は確度の高い順（積み上げの下から。色は濃い順）。どれでもなければ otherLabel
  plan: {
    enabled: true,
    amount: 400,
    plans: {},
    ranks: [{key: 'A', color: '#2E6A1A'}, {key: 'B', color: '#5A9E3A'}, {key: 'C', color: '#8FC46B'}, {key: 'D', color: '#B9D9A0'},
      {key: '創注', color: '#DCEBCF'}],
    otherLabel: '未設定',
    otherColor: '#C9C9C9'
  },
  // 読む列の見出し（改行・空白、全角・半角の違いは無視。customer・project・status はどのシートにも必要。ほかは無ければ使わない。空なら読まない）
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
    nextAction: '次回アクション',
    scenario: 'RXシナリオ構成',           // I列。複数選んだときは「A, B」の形（プルダウンの複数選択）。無ければ使わない（ほかの見出しは columnAliases）
    // デリバリー担当：新FMT にはこの列が無いので空（入力内容のチェック「デリバリー担当未設定」を出さない）。列を作ったら、その見出しを書く
    delivery: '',
    record: '商談記録（ＵＲＬ）'          // 提案資料・商談の音声の格納先（入力内容のチェック「資料なし」）
  },
  // 同じ列のほかの見出し（シートによって見出しが違うもの。columns の見出しが無ければ、これを順に探す）。
  // 2026/10/26 時点で、I列はシートによって「RXシナリオ構成」と「RXシナリオ構成（サービス）」の2通り。
  // 2026/10/29 に受注月の見出しが「受注月（改行）(yy/mm)」になった（変更の記録の列名も、ここでそろえて読む。mtgCanonicalLabel_）
  columnAliases: {scenario: ['RXシナリオ構成（サービス）'], orderMonth: ['受注月(yy/mm)']},
  // 資料に含めない列（見出しの形。正規表現。全角・半角と空白の違いは無視）。その列の変更は、資料の変化・更新項目・時点コピーとの比較に出さない
  // （2026/10/29 に足した P〜U列「FY26前半期売上（百万）」〜「FY28後半期売上（百万）」。年度が変わっても同じ形なら除く）
  ignoreHeaders: ['^FY\\d+(前|後)半期売上'],
  // 必須項目：進行中の案件で空なら、入力内容のチェックの「未入力」に出す（シートに無い見出しは見ない）
  required: ['案件ステータス', '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月', '活動状況', 'アカウント責任者'],   // columns の見出しなら columnAliases も
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
  mainRows: {changes: 7, focus: 10, owners: 9, deals: 4},   // 本編の表に出す行の数（残りは Appendix）。owners は Appendix ③ の責任者別の表の行、
                                                           // deals はサマリーの内訳の1区分に名前を出す案件の数
  // 色（緑を基調にしたトンマナ）。表の見出し・帯は green、表の1列目・合計は greenSoft、大きな数字・期待値の棒は greenDark
  // Appendix のページ（扉のあと）の色：扉（グレー）に合わせて、緑の見出し・1列目・大きな数字をグレーにする（本編と見分けやすく。2026/10/17 から）。
  // 区分・該当事項のラベルは意味の色（前向きは緑）なので変えない（mtgPills_ は colors のまま）。null なら本編と同じ緑
  appendixColors: {green: '#7F7F7F', greenSoft: '#EDEDED', greenDeep: '#595959', greenDark: '#404040', greenText: '#404040'},
  colors: {
    green: '#70AD47', greenDark: '#38761D', greenDeep: '#548235', greenSoft: '#E2EFDA', greenText: '#274E13',
    up: '#8FD36B', down: '#A6A6A6', dark: '#595959',
    ink: '#1A1A1A', body: '#1A1A1A', sub: '#404040', muted: '#595959', faint: '#7F7F7F',
    line: '#D0D0D0', frame: '#BFBFBF', rule: '#D9D9D9', white: '#FFFFFF', graySoft: '#EDEDED',
    gap: '#F2F2F2', gapLine: '#A6A6A6',   // データなし（作れなかったところ）の塗りと点線の枠
    cover: '#3A7D22',                // 表紙の背景（画像を付けるまで・付けられないとき）
    divider: '#7F7F7F',              // 扉（Appendix）の背景（同じ。扉はグレー）
    combined: '#7030A0'              // 紫の表紙の背景（同じ。統合版の表紙を紫にするとき。MTG_OPTIONS.combined.coverTone）
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
  plan: '計画に対する積み上げ',
  contents: '目次',                         // 統合版の2枚目
  overview: '全サブインダストリーの状況',   // 統合版の3枚目
  appendix: 'Appendix',
  changes: 'Appendix ①　前回から変化のある案件（一覧）',
  stalled: 'Appendix ②　停滞案件（一覧）',
  owners: 'Appendix ③　入力内容の確認事項（責任者別）',
  alerts: 'Appendix ③　入力内容の確認事項（一覧）'
};

/*
 * 前回から変化のある案件は、区分（MTG_CATEGORIES）ではなく今のステータスでまとめて並べる（2026/10/19 から。区分のラベルは出さない）：
 * 確度（MTG_OPTIONS.plan.ranks。A → 創注）→ 失注・保留・削除（ステータス未入力なども）。サマリーの「前回以降の動き」の行・本編・Appendix ① の順（mtgStatusGroups_）
 */

/* 期待値の増減の内訳（増える → 減るの順） */
const MTG_BRIDGE = [
  {key: 'new', label: '新規'}, {key: 'up', label: '期待値増'}, {key: 'revived', label: '復帰'},
  {key: 'down', label: '期待値減'}, {key: 'won', label: '受注'}, {key: 'lost', label: '失注'}, {key: 'hold', label: '保留'},
  {key: 'removed', label: '削除'}
];

/* 入力内容のチェックの種類（この順に出す。意味は docs/DESIGN.md の「入力内容のチェック」） */
// 予定日超過はクレデンシャル・オファリング履歴を使っているときだけ、デリバリー担当なしは MTG_OPTIONS.columns.delivery があるときだけ（mtgAlertTypes_）
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
  minFitSize: 7,         // 長い文（活動状況・課題）を全文入れるために小さくする下限（fit のセル）
  maxGrowLines: 26,      // 行を伸ばしても入れる行の数の上限（grow のセル。1ページに入る分）
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
 * 会議サマリ（Google スライドと PDF）を作って保存する。画面（MeetingRunDialog）で、会議日程（「INDEX_会議資料」の RX Business Meeting の日程。
 * Google カレンダーから importMeetingSchedules() で読む）をプルダウンで選び、作るシートを選んで作る。
 * 会議日と作った日時を記録し、次の会議の資料は今以降の変化を出す。
 */
function createBusinessMtgSummaries() {
  mschOpenRunDialog_('business', true);
}

/** 試しに会議サマリを作る（保存するが、会議日は記録しない。ファイル名に「試作」が付く）。画面は createBusinessMtgSummaries() と同じ */
function previewBusinessMtgSummaries() {
  mschOpenRunDialog_('business', false);
}

/**
 * 今を会議サマリの「基準」として記録する（作るシートを聞く）。旧フォーマットからの移行など、案件をまとめて貼り付けたあとに実行する。
 * 次に作る資料は、基準より後の変化だけを出す（貼り付けた案件を「新規」に数えない）。会議日は記録しない。もう一度実行したら基準を今に置き換える。
 */
function setBusinessMtgBaseline() {
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
    'スプレッドシート全体を「' + MTG_OPTIONS.copyName + '」（別のスプレッドシート）にコピーし、次の資料はそれと比べます。\n' +
    '移行の貼り付けがすべて終わってから実行してください。記録してよろしいですか？', ui.ButtonSet.YES_NO);
  if (ui && answer !== ui.Button.YES) return;
  const copy = mtgRecordBaseline_(ss, names, now);
  const message = '「' + names.join('」「') + '」に、基準（' + at + '）を記録しました。次に作る会議サマリは、基準より後の変化を出します。\n' +
    'スプレッドシート全体を「' + copy.title + '」にコピーしました：' + copy.url +
    (ui ? '' : '（スプレッドシートの画面に出せないため、確かめずにすべての追跡シートに記録しました）');
  console.log(message);
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/**
 * names のシートに、now を会議サマリの基準として記録する（時点コピーに今の値を保存し、MTG_MEETINGS_<シート名> の基準を置き換える）。
 * setBusinessMtgBaseline()・旧フォーマットからの転記（migrateLegacyDeals）から呼ぶ。戻り値 時点コピー {id, title, url}
 */
function mtgRecordBaseline_(ss, names, now) {
  const props = PropertiesService.getDocumentProperties();
  const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const copy = mtgSaveCopy_(ss, now, mtgCopyNote_('baseline'));
  names.forEach(name => {
    const list = mtgMeetings_(props, name).filter(m => !m.baseline)
      .concat([{date, cutoff: now.getTime(), baseline: true, copy: {id: copy.id, sheet: name}}])
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).slice(-MTG_OPTIONS.keepMeetings);
    props.setProperty(MTG_OPTIONS.meetingsKeyPrefix + name, JSON.stringify(list));
  });
  return copy;
}

/**
 * プルダウン（入力規則）の選択肢。一覧（requireValueInList）はその値、範囲（requireValueInRange。'定義_ステータス'!H2:H9 など）は範囲の表示の値。
 * プルダウンでなければ null
 */
function mtgListOptions_(rule) {
  if (!rule) return null;
  const types = SpreadsheetApp.DataValidationCriteria;
  const type = rule.getCriteriaType();
  if (type === types.VALUE_IN_LIST) return rule.getCriteriaValues()[0].map(v => String(v).trim()).filter(Boolean);
  if (type === types.VALUE_IN_RANGE) {
    const range = rule.getCriteriaValues()[0];
    return [].concat.apply([], range.getDisplayValues()).map(v => String(v).trim()).filter(Boolean);
  }
  return null;
}

/**
 * 記録した会議日を取り消す（会議日を打ち間違えたとき・本番を作り直す前に戻すとき。エディタから実行する）。
 * シートと会議日を聞き、その会議日の記録を消す。基準の記録と、保存した資料のファイルは消さない
 */
function removeBusinessMtgRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const names = mtgAskSheets_(ss, ui, '会議サマリの会議日の記録を取り消します（資料のファイル・時点コピーは消しません）。');
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
 * シートごとの「前回」（mtgSummaries_ と同じ選び方）：基準・前回の会議の記録・前の版の記録の順。どれも無ければ、会議日程の一覧（インデックス）の
 * 前の会議日（scheduled。その日の 0時から）。scheduled は mtgScheduledPrevious_ の戻り値（無ければ null）。
 * {previous, legacy: 前の版の記録の時刻 | 0, held: この会議日より前の記録, same: この会議日の記録}
 */
function mtgPreviousOf_(props, name, meeting, scheduled) {
  const list = mtgMeetings_(props, name);
  const held = list.filter(m => !m.baseline && m.date < meeting).pop() || null;
  const base = list.filter(m => m.baseline && m.date <= meeting).pop() || null;
  const same = list.find(m => !m.baseline && m.date === meeting) || null;
  let previous = base && (!held || base.cutoff > held.cutoff) ? base : held;
  const legacy = previous ? 0 : Number(props.getProperty(MTG_OPTIONS.lastKeyPrefix + name)) || 0;
  if (!previous && !legacy && scheduled) previous = scheduled;
  return {previous, legacy, held, same};
}

/**
 * 会議日程の一覧（インデックスの RX Business Meeting の行）で、meeting より前のいちばん新しい会議日（前回の会議）。
 * {date, cutoff: その日の 0時, scheduled: true}。無ければ null（前回の会議の資料の記録が無いときの「前回」に使う）
 */
function mtgScheduledPrevious_(meeting) {
  const prev = mtgIndexMeetings_(MTG_OPTIONS.index.kind).filter(m => m.date < meeting).pop();
  if (!prev) return null;
  return {date: prev.date, cutoff: Utilities.parseDate(prev.date, Session.getScriptTimeZone(), 'yyyy/MM/dd').getTime(), scheduled: true};
}

/**
 * 比べる時点の候補（画面のプルダウン。新しい順）。時点コピー（保存先のフォルダにあるもの（前の名前も）と、シートの会議の記録にあるもの。同じコピーは1つ）と、
 * 会議日程の一覧（インデックスの RX Business Meeting）の会議日（その日の 0時から。変更の記録で比べる）。
 * [{key: 'copy:<ID>' | 'date:yyyy/MM/dd', label, group: プルダウンの見出し, day: その時点の日付（資料作成時のコピーは会議日。画面では、選んだ会議日より前の
 *  ものだけ出す）, at, date, copyId, copyKind: 'meeting' | 'daily' | 'baseline' | '', meeting}]
 */
function mtgComparePoints_(ss) {
  const props = PropertiesService.getDocumentProperties();
  const tz = Session.getScriptTimeZone();
  const copies = new Map();
  Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name)).forEach(name => {
    mtgMeetings_(props, name).forEach(m => {
      if (!m.copy || !m.copy.id) return;
      if (!copies.has(m.copy.id)) {
        copies.set(m.copy.id, {copyId: m.copy.id, at: m.cutoff, copyKind: m.baseline ? 'baseline' : 'meeting', meeting: m.baseline ? '' : m.date});
      }
    });
  });
  mtgListCopies_().forEach(f => {
    if (!copies.has(f.id)) copies.set(f.id, {copyId: f.id, at: f.at, copyKind: f.kind, meeting: f.meeting});
  });
  const points = Array.from(copies.values()).map(c => {
    const date = c.meeting || Utilities.formatDate(new Date(c.at), tz, 'yyyy/MM/dd');
    return Object.assign(c, {key: 'copy:' + c.copyId, day: date, date, group: '時点コピー（その時点のシートの値と比べる）',
      label: Utilities.formatDate(new Date(c.at), tz, 'yyyy/M/d HH:mm') + '（' + mtgCopyKindText_(c.copyKind, c.meeting) + '）'});
  });
  mtgIndexMeetings_(MTG_OPTIONS.index.kind).forEach(m => {
    points.push({key: 'date:' + m.date, day: m.date, date: m.date, at: Utilities.parseDate(m.date, tz, 'yyyy/MM/dd').getTime(),
      group: '会議日から（変更の記録で比べる）', label: mtgDateLabel_(m.date) + ' の会議日（0時）から'});
  });
  return points.sort((a, b) => b.at - a.at);
}

/** 時点コピーの、いつ保存したかの文：「資料作成時・会議日 10/23」「日次」「基準」（分からなければ「保存時」） */
function mtgCopyKindText_(kind, meeting) {
  const k = MTG_OPTIONS.copyKinds;
  if (kind === 'meeting') return k.meeting + (meeting ? '・会議日 ' + mtgShortDate_(meeting) : '');
  return k[kind] || '保存時';
}

/**
 * 画面で選んだ比べる時点（key）を、資料の「前回」の形にする。空なら null（シートごとに前回の会議から。mtgPreviousOf_）。
 * 'copy:<ID>'：その時点コピーと比べる {date, cutoff, copy: {id}, sheets, baseline, copyKind, meeting, chosen}（コピーを開いて、あるシートを確かめる）
 * 'date:yyyy/MM/dd'：その日の 0時から、変更の記録で比べる {date, cutoff, scheduled, chosen}
 * 'at:yyyy/MM/dd HH:mm'：その日時から、変更の記録で比べる {date, cutoff, at, chosen}
 * 分からない・開けない・今より後なら止める
 */
function mtgResolvePoint_(key) {
  const text = String(key || '').trim();
  if (!text) return null;
  const tz = Session.getScriptTimeZone();
  let point = null;
  if (text.indexOf('copy:') === 0) {
    const c = mtgComparePoints_(SpreadsheetApp.getActiveSpreadsheet()).find(p => p.key === text);
    if (!c) throw new Error('比べる時点（' + MTG_OPTIONS.copyName + '）が見つかりません。画面を開き直して、選び直してください。');
    // コピーにあるシート（スプレッドシート全体のコピーなら、すべての追跡シート。前の版の値だけのコピーは、そのとき選んだシート）
    let sheets;
    try {
      sheets = SpreadsheetApp.openById(c.copyId).getSheets().map(sheet => sheet.getName());
    } catch (_) {
      throw new Error('選んだ' + MTG_OPTIONS.copyName + '（' + c.label + '）を開けません（ファイルを消した・権限が無いなど）。ほかの時点を選んでください。');
    }
    point = {date: c.date, cutoff: c.at, copy: {id: c.copyId}, sheets, baseline: c.copyKind === 'baseline', copyKind: c.copyKind, meeting: c.meeting,
      chosen: true};
  } else {
    const m = /^(date|at):(\d{4}\/\d{2}\/\d{2})(?: (\d{2}:\d{2}))?$/.exec(text);
    if (!m || (m[1] === 'at') !== !!m[3]) throw new Error('比べる時点「' + text + '」が分かりません。画面を開き直して、選び直してください。');
    const cutoff = Utilities.parseDate(m[2] + ' ' + (m[3] || '00:00'), tz, 'yyyy/MM/dd HH:mm').getTime();
    point = m[1] === 'at' ? {date: m[2], cutoff, at: true, chosen: true} : {date: m[2], cutoff, scheduled: true, chosen: true};
  }
  if (point.cutoff >= Date.now()) throw new Error('比べる時点は、今より前を選んでください。');
  return point;
}

/**
 * シートごとの「前回」（mtgPreviousOf_）を、選んだ比べる時点（point。mtgResolvePoint_ の戻り値）に置き換える。point が無い・前回と同じなら前回のまま。
 * 時点コピーにこのシートが無ければ、その時刻から変更の記録で比べる（noCopy）
 */
function mtgPreviousWith_(props, name, meeting, scheduled, point) {
  const p = mtgPreviousOf_(props, name, meeting, scheduled);
  if (!point) return p;
  const auto = p.previous;
  if (auto && (point.copy ? auto.copy && auto.copy.id === point.copy.id : !auto.copy && auto.cutoff === point.cutoff)) return p;
  const previous = {date: point.date, cutoff: point.cutoff, chosen: true};
  if (point.baseline) previous.baseline = true;
  if (point.copyKind) previous.copyKind = point.copyKind;
  if (point.meeting) previous.meeting = point.meeting;
  if (point.scheduled) previous.scheduled = true;
  if (point.at) previous.at = true;
  if (point.copy) {
    if (point.sheets.indexOf(name) >= 0) previous.copy = {id: point.copy.id, sheet: name};
    else previous.noCopy = true;
  }
  return Object.assign({}, p, {previous, legacy: 0});
}

/**
 * 選んだ比べる時点の文（確かめの文・資料の対象期間）：「10/8 18:20 の時点コピー（資料作成時・会議日 10/9）以降」「10/9 の会議日（0時）以降」「10/1 09:00 以降」
 */
function mtgChosenText_(prev) {
  const fmt = at => Utilities.formatDate(new Date(at), Session.getScriptTimeZone(), 'M/d HH:mm');
  if (prev.baseline) return '基準時点（' + fmt(prev.cutoff) + '）以降';
  if (prev.scheduled) return mtgShortDate_(prev.date) + ' の会議日（0時）以降';
  if (prev.at) return fmt(prev.cutoff) + ' 以降';
  return fmt(prev.cutoff) + ' の' + MTG_OPTIONS.copyName + '（' + mtgCopyKindText_(prev.copyKind, prev.meeting) + '）以降';
}

/** 作る前の確かめの文（会議日（曜日）・シートごとの対象期間・同じ会議日の資料を作成済みか・ダミーの記録）。本番なら記録の注意も。point は選んだ比べる時点 */
function mtgRunCheckText_(final, names, meeting, point) {
  const props = PropertiesService.getDocumentProperties();
  const tz = Session.getScriptTimeZone();
  const fmt = at => Utilities.formatDate(new Date(at), tz, 'M/d HH:mm');
  const scheduled = mtgScheduledPrevious_(meeting);
  const logStart = point && !point.copy ? mtgLogStart_(SpreadsheetApp.getActiveSpreadsheet()) : null;
  const lines = names.map(name => {
    const p = mtgPreviousWith_(props, name, meeting, scheduled, point);
    const prev = p.previous;
    const from = prev && prev.chosen ? '選んだ時点：' + mtgChosenText_(prev) +
      (prev.noCopy ? '（その時点のコピーにこのシートが無いため、変更の記録から）' : prev.copy ? '' : '（変更の記録から）')
      : prev && prev.baseline ? '基準（' + fmt(prev.cutoff) + '）以降'
      : prev && prev.scheduled ? '前回会議（' + mtgShortDate_(prev.date) + '）以降（前回の資料の記録が無いため、会議日から）'
        : prev ? '前回会議（' + mtgShortDate_(prev.date) + '）の資料作成時（' + fmt(prev.cutoff) + '）以降'
          : p.legacy ? '前回の資料作成時（' + fmt(p.legacy) + '）以降'
            : '直近' + MTG_OPTIONS.firstDays + '日間（前回の会議が無い）';
    const dummy = smpCountEvents_(SpreadsheetApp.getActiveSpreadsheet(), SMP_OPTIONS.dummy.eventPrefix + '-', name).changes;
    return '・' + name + '：' + from + (prev && prev.copy ? '（その時点のコピーと比べる）' : '') +
      (prev && prev.chosen && !prev.copy && logStart && prev.cutoff < logStart ?
        '\n　※ 変更の記録は ' + fmt(logStart) + ' からです。それより前の変化は出ません。' : '') +
      (final && p.same ? '\n　※ この会議日の資料は ' + fmt(p.same.cutoff) + ' に作成済み。作り直すと、次の会議の起点が今になります。' : '') +
      (dummy ? '\n　※ ダミーの変更の記録（' + dummy + '件）があります。資料に【ダミーデータを含む】と出ます。' +
        (final ? '配る資料なら、先に removeBusinessMtgDummyHistory() で消してください。' : '') : '');
  });
  return (final ? '本番の資料を作り、会議日を記録します。' : '試しに資料を作ります（会議日は記録しません。ファイル名に「試作」が付きます）。') +
    '\n\n会議日：' + mtgDateLabel_(meeting) + '\n比べる時点：' + (point ? '選んだ時点（全シート共通）' : '前回の会議（シートごと）') +
    '\n対象期間\n' + lines.join('\n') +
    (final ? '\n\nスプレッドシート全体を「' + MTG_OPTIONS.copyName + '」（別のスプレッドシート。タイトルは今の日時）にコピーし、次の会議の資料はそれと比べます。' +
      '\n記録した会議日は、次の会議の資料の起点になります。配布したあとは作り直さないでください。' : '');
}

/**
 * 会議サマリを作り始める（画面の「作成する」から。MeetingSchedule.gs の runMeetingMaterials）。作っている途中の仕事があれば、cancelJob ならやめて作り、
 * そうでなければ止める。時間がかかれば区切って続ける（mtgJobRun_）。戻り値は終わり（区切ったとき）の知らせ
 */
function mtgRunStart_(final, meeting, names, cancelJob, point) {
  const started = mtgNow_();   // 6分の上限は、この実行を始めたときから
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (mtgJobLoad_()) {
    if (!cancelJob) throw new Error(mtgJobPendingText_());
    mtgJobClear_();
    mtgJobDeleteTriggers_();
  }
  const now = new Date();
  const job = {final, meeting, now: now.getTime(), names, items: [], stored: 0, copy: null, index: null, folderUrl: '',
    combined: !!MTG_OPTIONS.combined.enabled, point: point || null};
  const done = mtgJobRun_(ss, job, started);
  const message = mtgJobMessage_(job, done);
  console.log(message);
  return message;
}


/* ---------------- 作る仕事（6分の上限に掛からないよう、区切って続ける） ---------------- */

/**
 * 会議サマリを作る仕事を、時間の許す限り進める。job = {final, meeting, now: 作り始めた時刻（ms）, names, items: 作った資料 [{name, pdf: ID, line: 知らせる文}],
 * stored: 格納した数, copy: 時点コピー, index: 書いたインデックス, folderUrl, combined: すべてのシートを作るか（統合版の案内を出す）,
 * point: 選んだ比べる時点（mtgResolvePoint_。無ければ null）}。
 * 順に、資料（シートごと）→ 本番なら時点コピーと会議日の記録 → 会議のフォルダへの格納とインデックス（シートごと）。
 * 統合版は作らない（すべてのサブインダストリーの資料を作ったあとに、createBusinessMtgCombinedSummary() で作る）。
 * 次の仕事が MTG_OPTIONS.run の見積もりで上限（limitMs）を越えそうなら、仕事を保存して1回きりのトリガーで続きを予約し、false を返す。全部済んだら true
 */
function mtgJobRun_(ss, job, started) {
  const o = MTG_OPTIONS;
  const R = o.run;
  const fits = ms => mtgNow_() - started + ms <= R.limitMs;
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  job.folderUrl = folder.getUrl();
  // 資料：残りのシートの中身をまとめて作り、1つずつスライドと PDF にする（中身は作り始めた時刻 job.now で数える）
  if (job.items.length < job.names.length) {
    if (!fits(R.summaryMs + R.deckMs)) return mtgJobPause_(job);
    const props = PropertiesService.getDocumentProperties();
    const {plans, summaries} = mtgSummaries_(ss, job.names.slice(job.items.length), new Date(job.now), job.final, job.meeting, job.point);
    for (let i = 0; i < summaries.length; i++) {
      if (i > 0 && !fits(R.deckMs)) return mtgJobPause_(job);
      const summary = summaries[i];
      const out = mtgExportDeck_(mtgFileName_(summary.name, job.meeting, job.final), mtgSlides_(summary), folder);
      // 試作：その会議日の最初の試作の、入力内容のチェックの結果を覚える（同じ会議日の資料で、解消した件数を出す）
      if (!job.final && !plans[i].draft) {
        props.setProperty(o.draftKeyPrefix + summary.name, JSON.stringify({meeting: job.meeting, at: job.now, keys: summary.alerts.map(a => mtgHash_(a.key))}));
      }
      job.items.push({name: summary.name, pdf: out.pdf.getId(), line: mtgResultLine_(ss, summary, out.pdf)});
    }
  }
  // 本番：今の値を時点コピーに保存し（次の会議の資料はこれと比べる）、会議日を記録する。資料を全部作れてから
  if (job.final && !job.copy) {
    if (!fits(R.copyMs)) return mtgJobPause_(job);
    const copy = mtgSaveCopy_(ss, new Date(job.now), mtgCopyNote_('meeting', job.meeting));
    job.copy = {id: copy.id, title: copy.title, url: copy.url};
    const props = PropertiesService.getDocumentProperties();
    job.names.forEach(name => {
      const list = mtgMeetings_(props, name).filter(m => m.baseline || m.date !== job.meeting)
        .concat([{date: job.meeting, cutoff: job.now, copy: {id: copy.id, sheet: name}}])
        .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0)).slice(-o.keepMeetings);
      props.setProperty(o.meetingsKeyPrefix + name, JSON.stringify(list));
    });
  }
  // 本番：会議のフォルダに PDF を格納し、インデックスに書く（できなくても止めない。資料は保存先にあるので、あとで storeBusinessMtgSummaries() で格納できる）
  while (job.final && job.stored < job.items.length) {
    if (!fits(R.storeMs)) return mtgJobPause_(job);
    const item = job.items[job.stored];
    const stored = mtgStoreToMeetingFolder_(item.name, job.meeting, DriveApp.getFileById(item.pdf));
    const index = stored.error ? null : mtgWriteIndex_(item.name, job.meeting, stored);
    item.line += mtgStoredText_(stored, index);
    if (index && index.url && !job.index) job.index = index;
    job.stored++;
  }
  return true;
}

/** 今の時刻（ms。実行時間の上限を見る用。テストで差し替える） */
function mtgNow_() {
  return Date.now();
}

/** 資料1つの知らせ（ファイル名・件数・URL と、作れなかったところ（データなし）などの注意） */
function mtgResultLine_(ss, summary, file) {
  const gaps = mtgGapList_(summary);
  // 変化が無く、サンプル（addSampleHistory()。会議サマリは数えない）の記録だけがあるときは、資料の確認用の関数を案内する
  const samples = !summary.kpi.changed && !summary.dummy ? smpCountEvents_(ss, SMP_OPTIONS.eventPrefix + '-', summary.name).changes : 0;
  return '「' + summary.name + '」：' + file.getName() + '（' + mtgCountsText_(summary) + '）\n　' + file.getUrl() +
    (summary.copy ? '\n　' + (summary.previous && summary.previous.chosen ? '' : '前回からの') + '変化は「' + summary.copy.title + '」と比べました。' : '') +
    (summary.copyError ? '\n　※ ' + summary.copyError : '') +
    (summary.unlogged ? '\n　※ 変更の記録が無い変更 ' + summary.unlogged + '件（差分追跡が止まっていた間などの変更。資料には出しています）。' +
      'diagnoseDiffState() で差分追跡が動いているか確かめてください。' : '') +
    (gaps.length ? '\n　' + MTG_OPTIONS.gapText + '：' + gaps.join('\n　　') : '') +
    (summary.dummy ? '\n　※ ダミーの変更の記録（' + summary.dummy + '件）を含みます（資料に【ダミーデータを含む】）。消すときは removeBusinessMtgDummyHistory()。' : '') +
    (samples ? '\n　※ サンプル（addSampleHistory()）の変更の記録 ' + samples + '件は、会議サマリには出ません（基準を変えても出ません）。' +
      '資料の形を確かめるときは、addBusinessMtgDummyHistory() でダミーの変更の記録を足してください。' : '');
}

/** 会議のフォルダへの格納とインデックスの知らせ（資料1つの知らせの後ろに付ける） */
function mtgStoredText_(stored, index) {
  if (stored.error) {
    return '\n　※ 会議のフォルダに格納できませんでした（' + stored.error + '）。資料は保存先にあります。' +
      'フォルダを用意してから、エディタで storeBusinessMtgSummaries() を実行して格納してください。';
  }
  return '\n　会議のフォルダ「' + stored.folder.getName() + '」' + (stored.created ? 'を作って' : 'に') + '格納しました：' + stored.folder.getUrl() +
    (index && index.error ? '\n　※ ' + index.error : '');
}

/** 終わり（done）・区切ったときの知らせ */
function mtgJobMessage_(job, done) {
  const o = MTG_OPTIONS;
  const lines = job.items.map(item => item.line);
  if (!done) {
    const rest = job.names.slice(job.items.length);
    const left = [].concat(rest.length ? ['資料 ' + rest.length + '件（' + rest.join('・') + '）'] : [],
      job.final && !job.copy ? ['時点コピーと会議日の記録'] : [],
      job.final && job.stored < job.names.length ? ['会議のフォルダへの格納とインデックス'] : []);
    return ['時間がかかるため、ここで区切りました（Apps Script の1回の実行は6分まで）。残りは約' + Math.round(o.run.continueAfterMs / 60000) +
      '分後に自動で続けます：' + left.join('・') + '。',
    '続きが終わったら、エディタで showBusinessMtgSummariesResult() を実行すると結果が出ます（実行ログにも出ます）。終わるまで、会議サマリの関数は実行しないでください。',
    ''].concat(lines, ['', '保存先：' + job.folderUrl]).join('\n');
  }
  // 本番：この会議日の本番の資料がそろったか（サブインダストリーごとに作ってもよい）。そろったら統合版の作り方、まだなら残りのサブインダストリー
  if (job.final && job.combined) {
    const rest = mtgMeetingPending_(SpreadsheetApp.getActiveSpreadsheet(), job.meeting);
    lines.push('', rest.length
      ? 'この会議日の本番の資料が、まだできていないサブインダストリー（' + rest.length + '件）：' + rest.join('・') + '。' +
        'すべてできたら、createBusinessMtgSummaries() の終わりにこの画面で知らせます（統合版（' + o.combined.name + '）は、そのあと createBusinessMtgCombinedSummary() で作る）。'
      : 'この会議日の本番の資料が、すべてのサブインダストリーでそろいました。統合版（' + o.combined.name + '）は、エディタで createBusinessMtgCombinedSummary() を実行して作ってください（会議日を入力します）。');
  }
  lines.push('', '保存先：' + job.folderUrl);
  if (job.index) {
    lines.push('「' + job.index.title + '」シートに、資料のファイルを書きました：' + job.index.url);
    if (o.index.chip && !job.index.chip) {
      lines.push('（スマートチップにできなかったので、ファイル名のリンクで書きました。チップにするには、エディタの「サービス」に Google Sheets API を追加してください）');
    }
  }
  if (job.copy) lines.push('スプレッドシート全体を「' + job.copy.title + '」にコピーしました（次の会議の資料はこれと比べます）：' + job.copy.url);
  if (o.keepSlides) lines.push('同じフォルダに、Google スライド（同じ名前）も保存しました。');
  lines.push(job.final
    ? '会議日（' + job.meeting + '）を記録しました。次の会議の資料は、今以降の変化を出します。配るときは、PDF を参加者に共有してください（保存先のフォルダは共有しない）。'
    : '試作のため、会議日は記録していません。');
  return lines.join('\n');
}

/** 仕事を保存し（スクリプトのプロパティ。資料ごとの知らせは別のキー）、続きのトリガーを1つだけ作る。false を返す */
function mtgJobPause_(job) {
  const R = MTG_OPTIONS.run;
  const props = PropertiesService.getScriptProperties();
  mtgJobClear_();
  job.items.forEach((item, i) => props.setProperty(R.jobKey + '_ITEM_' + i, JSON.stringify(item)));
  props.setProperty(R.jobKey, JSON.stringify(Object.assign({}, job, {items: job.items.length, savedAt: Date.now()})));
  mtgJobDeleteTriggers_();
  ScriptApp.newTrigger(R.handler).timeBased().after(R.continueAfterMs).create();
  console.log('会議サマリ：時間がかかるため区切りました（資料 ' + job.items.length + '/' + job.names.length + '）。続きを予約しました。');
  return false;
}

/** 保存した仕事（無ければ null） */
function mtgJobLoad_() {
  const R = MTG_OPTIONS.run;
  const props = PropertiesService.getScriptProperties();
  const job = mtgReadJson_(props, R.jobKey);
  if (!job) return null;
  const items = [];
  for (let i = 0; i < job.items; i++) items.push(mtgReadJson_(props, R.jobKey + '_ITEM_' + i));
  job.items = items;
  return job;
}

/** 保存した仕事を消す（続きのトリガーは mtgJobDeleteTriggers_） */
function mtgJobClear_() {
  const R = MTG_OPTIONS.run;
  const props = PropertiesService.getScriptProperties();
  props.getKeys().filter(key => key === R.jobKey || key.indexOf(R.jobKey + '_ITEM_') === 0).forEach(key => props.deleteProperty(key));
}

/** 続きのトリガーを外す（前の版の関数名のトリガーも） */
function mtgJobDeleteTriggers_() {
  const R = MTG_OPTIONS.run;
  const handlers = [R.handler].concat(R.legacyHandlers || []);
  ScriptApp.getProjectTriggers().filter(t => handlers.indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
}

/** 結果の知らせを保存する（スクリプトのプロパティ。1つの値の大きさに上限があるので、2,000文字ずつ分ける） */
function mtgJobSaveResult_(message) {
  const R = MTG_OPTIONS.run;
  const props = PropertiesService.getScriptProperties();
  props.getKeys().filter(key => key.indexOf(R.resultKey) === 0).forEach(key => props.deleteProperty(key));
  const text = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' に終わりました。\n\n' + message;
  for (let i = 0; i * 2000 < text.length; i++) props.setProperty(R.resultKey + '_' + i, text.slice(i * 2000, (i + 1) * 2000));
}

/**
 * 区切った会議サマリの続き（時間主導のトリガーから。1回きり）。保存した仕事を続け、まだ残れば、また区切って次を予約する。
 * 終わったら（止まったら）知らせを保存する（showBusinessMtgSummariesResult() で出す。実行ログにも出す）
 */
function continueBusinessMtgSummaries() {
  const started = mtgNow_();
  mtgJobDeleteTriggers_();
  const job = mtgJobLoad_();
  if (!job) return;
  let message;
  try {
    if (!mtgJobRun_(SpreadsheetApp.getActiveSpreadsheet(), job, started)) return;   // また区切った（次を予約した）
    message = mtgJobMessage_(job, true);
  } catch (error) {
    console.error(error && error.stack || error);
    message = '途中で止まりました：' + (error && error.message || error) + '\ncreateBusinessMtgSummaries()（試作なら previewBusinessMtgSummaries()）で、作り直してください。\n\n' +
      mtgJobMessage_(job, true);
  }
  mtgJobClear_();
  mtgJobSaveResult_(message);
  console.log(message);
}

/** 会議サマリの続き（自動で続けたもの）の結果・進み具合を出す（エディタから実行） */
function showBusinessMtgSummariesResult() {
  const R = MTG_OPTIONS.run;
  const ui = mtgUi_();
  const job = mtgJobLoad_();
  let message;
  if (job) {
    const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === R.handler).length;
    message = '会議サマリを作っている途中です（会議日 ' + job.meeting + '。資料 ' + job.items.length + '/' + job.names.length + '件' +
      (job.final ? '・格納 ' + job.stored + '/' + job.names.length + '件' : '') + '）。' +
      (triggers ? '自動で続けています。少し待ってから、もう一度実行してください。'
        : '続きの予約がありません。createBusinessMtgSummaries()（試作なら previewBusinessMtgSummaries()）で作り直してください。');
  } else {
    const props = PropertiesService.getScriptProperties();
    const keys = props.getKeys().filter(key => key.indexOf(R.resultKey + '_') === 0).sort((a, b) => Number(a.split('_').pop()) - Number(b.split('_').pop()));
    message = keys.length ? keys.map(key => props.getProperty(key)).join('') : '自動で続けた会議サマリの結果はありません。';
  }
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** 作っている途中の仕事があれば、その知らせ（画面に出し、「やめて新しく作る」を選んでもらう）。無ければ空 */
function mtgJobPendingText_() {
  const job = mtgJobLoad_();
  if (!job) return '';
  const at = Utilities.formatDate(new Date(job.savedAt || job.now), Session.getScriptTimeZone(), 'M/d HH:mm');
  return '前に始めた会議サマリ（会議日 ' + job.meeting + '・' + (job.final ? '本番' : '試作') + '。資料 ' + job.items.length + '/' + job.names.length + '件。' +
    at + ' に区切り）を、自動で続けている途中です。やめて新しく作るときは「やめて新しく作る」にチェックを入れてください' +
    '（続きが終わったら showBusinessMtgSummariesResult() で結果を見られます）。';
}

/* ---------------- Google チャットへの通知（webhook）の共通 ---------------- */

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
      '\nwebhook の URL が消されていないか、setBusinessMtgSubChats() で登録し直してください。');
  }
}

/** 送り先の見せ方（「スペース AAAA…」。鍵・トークンは出さない） */
function mtgChatSpace_(webhook) {
  const m = /\/spaces\/([^/?]+)/.exec(webhook);
  return 'スペース ' + (m ? m[1] : '（不明）');
}

/* ---------------- RXシナリオ構成別の期待値（構成が未確定のあいだの別の資料） ---------------- */

/** RXシナリオ構成のセルの値を、選んだシナリオの一覧に（プルダウンの複数選択は「A, B」。改行でも区切る。選択肢の名前に読点があっても切らない。同じものは1つ） */
function mtgScenarioList_(text) {
  const out = [];
  String(text || '').split(/\s*[,，\n]\s*/).map(v => v.trim()).filter(Boolean).forEach(v => { if (out.indexOf(v) < 0) out.push(v); });
  return out;
}

/**
 * RXシナリオ構成別の期待値の資料を作る（エディタから実行する。構成が未確定のあいだは、事前資料とは別の資料）。シートを選び（空欄ならすべて）、
 * 進行中の案件の期待値を、シナリオ（「RXシナリオ構成」の列。複数選択）ごとにまとめた PDF を保存先のフォルダに作る（会議日・記録は変えない）。
 * 複数のシナリオを選んだ案件は、期待値を選んだ数で等分して配分（按分。合計は期待値合計と一致）し、あわせて関与（そのシナリオを含む案件の期待値の合計。
 * 重複あり）・件数・単独（そのシナリオだけを選んだ案件）を出す。空欄の案件は「未設定」
 */
function createBusinessMtgScenarioSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const title = CRED_OPTIONS.menuTitle;
  const o = MTG_OPTIONS.scenario;
  const names = mtgAskSheets_(ss, ui, o.title + 'の資料（試作）を作ります。');
  if (!names) return;
  const data = mtgScenarioData_(ss, names);
  if (!data.withColumn.length) {
    const names = [MTG_OPTIONS.columns.scenario].concat((MTG_OPTIONS.columnAliases || {}).scenario || []);
    throw new Error('選んだシートに「' + names.join('」か「') + '」の列がありません。2行目の見出しを確かめてください（MeetingSummary.gs の MTG_OPTIONS.columns.scenario・columnAliases）。');
  }
  const t = mtgScenarioTotals_(data);
  const answer = ui.alert(title, o.title + 'の資料（試作）を作ります。\n\n' +
    '・対象：' + data.withColumn.join('・') + '（進行中の案件 ' + t.count + '件・期待値 ' + mtgAmount_(t.total) + '百万円）\n' +
    '・シナリオ：' + t.items.filter(item => !item.unset).map(item => item.key).join('・') + '\n' +
    '・' + o.unsetLabel + '：' + t.unset.count + '件' +
    (data.withoutColumn.length ? '\n・列が無いので除くシート：' + data.withoutColumn.join('・') : '') +
    '\n\n保存先のフォルダに PDF を作ります（会議日の記録・会議のフォルダ・インデックスは変えません）。作りますか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  const now = new Date();
  const slides = mtgScenarioSlides_(data, t, now);
  const name = o.fileName.replace('{date}', Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd'));
  const out = mtgExportDeck_(name, slides, folder);
  const message = o.title + 'の資料（試作）を作りました：' + out.pdf.getName() + '（' + slides.length + 'ページ）\n' + out.pdf.getUrl();
  console.log(message);
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * シートごとの進行中の案件と、シナリオの選択肢（読むだけ）。
 * {withColumn: 列のあるシート, withoutColumn, options: プルダウンの選択肢の順（どのシートのものも。選択肢に無い値はあとに足す）,
 *  deals: [{sheet, customer, project, status, rank: 確度, weighted: 期待値（空は 0）, blank: 期待値が空か, scenarios}]}
 */
function mtgScenarioData_(ss, names) {
  const out = {withColumn: [], withoutColumn: [], options: [], deals: []};
  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const rule = DIFF_RULES[name];
    const table = mtgReadDeals_(sheet, rule);
    if (!table.has.scenario) {
      out.withoutColumn.push(name);
      return;
    }
    out.withColumn.push(name);
    const col = mtgColumnIndex_(table.headers, 'scenario');
    (mtgListOptions_(sheet.getRange(rule.headerRow + 1, col + 1).getDataValidation()) || []).forEach(v => { if (out.options.indexOf(v) < 0) out.options.push(v); });
    table.deals.filter(d => d.kind === 'active').forEach(d => out.deals.push({sheet: name, customer: d.customer, project: d.project, status: d.status,
      rank: mtgPlanRank_(d.status), weighted: d.weighted || 0, blank: d.weighted == null, scenarios: d.scenarios}));
  });
  out.deals.forEach(d => d.scenarios.forEach(v => { if (out.options.indexOf(v) < 0) out.options.push(v); }));
  return out;
}

/**
 * シナリオごとの集計。{total, count, blank: 期待値が空の件数, items: [{key, unset, share: 按分した期待値, full: 関与（含む案件の期待値の合計）, count, solo: 単独の件数,
 *  byRank: {確度: 按分}, bySheet: {シート: 按分}}]（按分の大きい順。未設定は最後）, unset: 未設定の item, combos: [{key: 「A＋B」, size, count, value, deals}]（期待値の大きい順）}
 */
function mtgScenarioTotals_(data) {
  const o = MTG_OPTIONS.scenario;
  const make = (key, unset) => ({key, unset: !!unset, share: 0, full: 0, count: 0, solo: 0, byRank: {}, bySheet: {}});
  const items = new Map(data.options.map(v => [v, make(v)]));
  const unset = make(o.unsetLabel, true);
  const combos = new Map();
  let total = 0;
  data.deals.forEach(d => {
    total += d.weighted;
    const list = d.scenarios.length ? d.scenarios.map(v => items.get(v)) : [unset];
    const part = d.weighted / list.length;
    list.forEach(item => {
      item.share += part;
      item.full += d.weighted;
      item.count++;
      if (list.length === 1) item.solo++;
      item.byRank[d.rank] = (item.byRank[d.rank] || 0) + part;
      item.bySheet[d.sheet] = (item.bySheet[d.sheet] || 0) + part;
    });
    if (d.scenarios.length) {
      const key = data.options.filter(v => d.scenarios.indexOf(v) >= 0).join('＋');
      if (!combos.has(key)) combos.set(key, {key, size: d.scenarios.length, count: 0, value: 0, deals: []});
      const c = combos.get(key);
      c.count++;
      c.value += d.weighted;
      c.deals.push(d);
    }
  });
  const sorted = Array.from(items.values()).sort((a, b) => (b.share - a.share) || (b.count - a.count));
  return {total, count: data.deals.length, blank: data.deals.filter(d => d.blank).length, items: sorted.concat([unset]), unset,
    combos: Array.from(combos.values()).sort((a, b) => (b.value - a.value) || (b.count - a.count))};
}

/** RXシナリオ構成別の期待値の資料のスライド：表紙・シナリオ別（横棒）・シナリオ × サブインダストリー・組み合わせ */
function mtgScenarioSlides_(data, t, now) {
  const o = MTG_OPTIONS.scenario;
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const tz = Session.getScriptTimeZone();
  const line = (y, text) => mtgBox_(S.x, y, S.w, 16, mtgWrap_(text, S.w, 9, 1)[0], {size: 9, color: c.white, align: 'CENTER'});
  const cover = {
    title: MTG_TITLES.cover,
    cover: true,
    combined: true,
    template: {title: o.title, subtitle: MTG_OPTIONS.meetingName + '　参考資料【試作】', date: mtgDateLabel_(Utilities.formatDate(now, tz, 'yyyy/MM/dd')), name: ''},
    elements: [
      line(350, '収録：' + data.withColumn.join('・') + '（' + data.withColumn.length + 'サブインダストリー）'),
      line(365, '出典：各サブインダストリーのシートの「' + MTG_OPTIONS.columns.scenario + '」と期待値（' + Utilities.formatDate(now, tz, 'yyyy/M/d HH:mm') + ' 時点）')
    ]
  };
  const frame = {source: '【試作】' + o.title + '（' + Utilities.formatDate(now, tz, 'M/d') + ' 時点）　金額の単位：百万円'};
  return mtgNumberSlides_(frame, [cover, mtgScenarioBarsSlide_(t), mtgScenarioMatrixSlide_(data, t)].concat(mtgScenarioCombosSlides_(t)));
}

/** シナリオ別：按分した期待値の横棒（確度別）と、按分・関与・件数・単独。未設定は最後（灰色） */
function mtgScenarioBarsSlide_(t) {
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const ranks = o.plan.ranks.concat([{key: o.plan.otherLabel, color: o.plan.otherColor}]);
  const multi = t.combos.filter(cb => cb.size > 1).reduce((sum, cb) => sum + cb.count, 0);
  const lead = '進行中案件 ' + t.count + '件・期待値 ' + mtgAmount_(t.total) + '百万円を、RXシナリオ構成ごとに集計（按分の大きい順）。' +
    '\n按分：複数選んだ案件（' + multi + '件）は選んだ数で等分（合計＝期待値）。関与：そのシナリオを含む案件の期待値の合計（重複あり）。' +
    (t.blank ? '期待値が空の ' + t.blank + '件は 0。' : '');
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 420, 20, 'RXシナリオ構成別の期待値（進行中・按分・百万円）', {size: 10.5, bold: true, color: c.ink})];
  const legend = ranks.map(r => ({label: r.key, color: r.color}));
  legend.forEach(item => { item.bw = mtgTextWidth_(item.label, 8.5) + 16; });
  let lx = x + w - legend.reduce((sum, item) => sum + 6 + item.bw, 0) - (legend.length - 1) * (8 - 7.2 - 1.6);
  legend.forEach(item => {
    el.push({type: 'box', x: lx, y: y + 5, w: 7, h: 7, fill: item.color, shape: 'RECTANGLE', role: 'legend'});
    el.push(mtgBox_(lx + 6, y - 1, item.bw, 18, item.label, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += 6 + item.bw + (8 - 7.2 - 1.6);
  });
  const nameW = 150;
  const cols = [{label: '按分', w: 54}, {label: '関与（重複あり）', w: 78}, {label: '件数', w: 40}, {label: 'うち単独', w: 46}];
  const colsW = cols.reduce((sum, col) => sum + col.w, 0);
  const barX = x + nameW + 6;
  const barW = w - nameW - 6 - colsW - 10;
  const headY = y + 22;
  const rows = t.items;
  const rowH = Math.max(12, Math.min(22, (S.bottom - headY - 16) / Math.max(rows.length, 1)));
  const max = Math.max.apply(null, [1].concat(rows.map(item => item.share)));
  const scale = barW / max;
  let cx = barX + barW + 10;
  cols.forEach(col => {
    el.push(mtgBox_(cx - S.padX / 2, headY, col.w + S.padX, 14, col.label, {size: 8, color: c.sub, align: 'CENTER', role: 'scenario-head'}));
    col.x = cx;
    cx += col.w;
  });
  const size = rowH < 16 ? 7.5 : 8.5;
  rows.forEach((item, i) => {
    const ry = headY + 16 + rowH * i;
    if (item.unset) el.push({type: 'box', x, y: ry, w, h: 0.75, fill: c.line, shape: 'RECTANGLE'});
    else if (i % 2 === 1) el.push({type: 'box', x, y: ry, w, h: rowH, fill: c.greenSoft, alpha: 0.35, shape: 'RECTANGLE', role: 'scenario-band'});
    el.push(mtgBox_(x, ry, nameW, rowH, mtgWrap_(item.key, nameW, size, 1)[0], {size, bold: true, color: item.unset ? c.muted : c.ink, role: 'scenario-name'}));
    let level = 0;
    const bh = Math.max(rowH - 8, 5);
    ranks.forEach(r => {
      const v = item.byRank[r.key] || 0;
      if (v <= 0) return;
      el.push({type: 'box', x: barX + level * scale, y: ry + (rowH - bh) / 2, w: Math.max(v * scale, 0.75), h: bh, fill: item.unset ? o.plan.otherColor : r.color,
        line: c.white, lineWeight: 0.5, shape: 'RECTANGLE', role: 'scenario-bar'});
      level += v;
    });
    [mtgAmount_(Math.round(item.share * 10) / 10), mtgAmount_(Math.round(item.full * 10) / 10), item.count + '件', item.unset ? '—' : item.solo + '件'].forEach((text, k) => {
      el.push(mtgBox_(cols[k].x - S.padX / 2, ry, cols[k].w + S.padX, rowH, text, {size, bold: k === 0, color: k === 0 ? c.ink : c.body, align: 'END', role: 'scenario-value'}));
    });
  });
  return {title: o.scenario.title + '（全サブインダストリー）', lead, elements: el};
}

/** シナリオ × サブインダストリー：按分した期待値の表（いちばん右に合計、いちばん下に合計）。0 は「—」 */
function mtgScenarioMatrixSlide_(data, t) {
  const S = MTG_SLIDE;
  const c = MTG_OPTIONS.colors;
  const keys = t.items.map(item => item.key);
  const nameW = 110;
  const each = Math.floor((S.w - nameW) / (keys.length + 1));
  const columns = [S.w - each * (keys.length + 1)].concat(keys.map(() => each), [each]);
  const size = each < 52 ? 7 : 8;
  const head = mtgHeadRow_(columns, ['サブインダストリー'].concat(keys, ['合計']), size);
  const num = v => (Math.abs(v) < 0.05 ? '—' : mtgAmount_(Math.round(v * 10) / 10));
  const row = (label, values, total) => mtgRow_(columns, [{text: label, size, bold: true, color: c.ink, fill: c.greenSoft}]
    .concat(values.map(v => ({text: num(v), size, fill: total ? c.greenSoft : c.white, align: 'END', bold: !!total}))), 0);
  const rows = data.withColumn.map(name => {
    const values = t.items.map(item => item.bySheet[name] || 0);
    return row(name, values.concat([values.reduce((a, b) => a + b, 0)]));
  }).concat([row('合計', t.items.map(item => item.share).concat([t.total]), true)]);
  const lead = 'サブインダストリーごとの RXシナリオ構成別の期待値（進行中・按分）。行の合計は各サブインダストリーの期待値合計と一致。';
  return {title: MTG_OPTIONS.scenario.title + '（サブインダストリー別）', lead, elements: [{type: 'table', x: S.x, y: S.top, columns, rows: [head].concat(rows)}]};
}

/** よく選ばれる組み合わせ（単独も）：期待値の大きい順に scenario.combos 件まで。入りきらなければ続きのページ */
function mtgScenarioCombosSlides_(t) {
  const S = MTG_SLIDE;
  const c = MTG_OPTIONS.colors;
  const o = MTG_OPTIONS.scenario;
  const columns = mtgColumns_([250, 50, 70, 290], 3);
  const head = mtgHeadRow_(columns, ['RXシナリオ構成（組み合わせ）', '件数', '期待値', '主な案件（期待値の大きい順）']);
  const list = t.combos.slice(0, o.combos);
  const title = o.title + '（組み合わせ）';
  const lead = '選ばれた RXシナリオ構成の組み合わせ（単独も含む）ごとの件数・期待値（進行中。按分しない）。期待値の大きい順に ' + list.length + '件' +
    (t.combos.length > list.length ? '（全 ' + t.combos.length + '通り）' : '') + '。';
  if (!list.length) return [{title, lead: 'RXシナリオ構成を選んだ進行中の案件なし。', elements: mtgEmptyTable_(columns, head, 'none', '')}];
  const rows = list.map(cb => {
    const top = cb.deals.slice().sort((a, b) => b.weighted - a.weighted);
    return mtgRow_(columns, [
      {text: cb.key, lines: 3, bold: true, color: c.ink, fill: c.greenSoft},
      {text: cb.count + '件', fill: c.white, align: 'CENTER'},
      {text: mtgAmount_(Math.round(cb.value * 10) / 10), fill: c.white, align: 'END'},
      {text: mtgNames_(top, d => d.sheet, columns[3], 2), lines: 2, fill: c.white}
    ]);
  });
  return mtgTablePages_(head, rows).map((p, n) => ({title: title + (n ? '（続き）' : ''), lead,
    elements: [{type: 'table', x: S.x, y: S.top, columns, rows: p.rows}]}));
}

/* ---------------- サブインダストリーごとの Google チャットへの通知 ---------------- */

/**
 * サブインダストリーごとの Google チャットのスペースの webhook を登録する（エディタから実行する。管理者が1回。スペースが変わったときも）。
 * シートを選び、シートごとに webhook の URL を聞いてスクリプトのプロパティに残す（空欄なら今のまま、「削除」と入力すると外す）。
 * 登録したスペースには、確かめのメッセージを送れる
 */
function setBusinessMtgSubChats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const title = CRED_OPTIONS.menuTitle;
  const props = PropertiesService.getScriptProperties();
  // 前の版の全体のスペースへの通知の設定（webhook。鍵を含む）と、自動通知のトリガーは、使わないので消す
  MTG_OPTIONS.chat.oldKeys.forEach(key => props.deleteProperty(key));
  mtgRemoveLegacyTrigger_();
  const names = mtgAskSheets_(ss, ui, 'サブインダストリーごとの Google チャットのスペースの webhook を登録します。\n今の設定：\n' +
    Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name)).map(name => '・' + name + '：' + mtgSubChatState_(name)).join('\n'));
  if (!names) return;
  const changed = [];
  const lines = [];
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const response = ui.prompt(title, '「' + name + '」の Google チャットのスペースの webhook の URL を貼り付けてください' +
      '（スペースの「アプリと統合」→「Webhook」で作ったもの）。\n空欄なら今のまま、「削除」と入力すると外します。\n\n今の設定：' + mtgSubChatState_(name),
      ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) break;
    const text = String(response.getResponseText() || '').trim();
    if (!text) {
      lines.push('・' + name + '：変えていません（' + mtgSubChatState_(name) + '）');
      continue;
    }
    if (text === '削除') {
      props.deleteProperty(MTG_OPTIONS.subChat.webhookKeyPrefix + name);
      lines.push('・' + name + '：外しました');
      continue;
    }
    if (!/^https:\/\/chat\.googleapis\.com\/v1\/spaces\/[^/?]+\/messages\?\S*key=\S+&token=\S+$/.test(text)) {
      throw new Error('「' + name + '」：Google チャットの webhook の URL ではありません。https://chat.googleapis.com/v1/spaces/…/messages?key=…&token=… の形の URL を貼り付けてください。' +
        (lines.length ? '\n（ここまでの登録は済んでいます：' + lines.join(' ') + '）' : ''));
    }
    props.setProperty(MTG_OPTIONS.subChat.webhookKeyPrefix + name, text);
    changed.push(name);
    lines.push('・' + name + '：' + mtgChatSpace_(text) + ' に送ります');
  }
  if (!lines.length) return;
  const message = 'サブインダストリーごとの Google チャットの設定：\n' + lines.join('\n') +
    '\n\nremindBusinessMtgUpdate() で管理フォーマットの更新のお願いを、notifySubBusinessMtgSummaries() で事前資料のご案内を送れます。';
  if (changed.length && ui.alert(title, message + '\n\n登録したスペース（' + changed.length + '件）に、確かめのメッセージをいま送りますか？', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    const results = changed.map(name => mtgSubChatPost_(name, '【テスト送信】' + name + ' の ' + MTG_OPTIONS.meetingName +
      'の通知の設定を確認するためのメッセージです。\n今後、管理フォーマットの更新のお願いと事前資料のご案内をこちらにお送りいたします。'));
    const done = message + '\n\n確かめのメッセージ：\n' + results.map(r => '・' + r.name + '：' + (r.error ? '送れませんでした（' + r.error + '）' : '送りました')).join('\n');
    console.log(done);
    ui.alert(title, done, ui.ButtonSet.OK);
    return;
  }
  console.log(message);
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * 管理フォーマットの更新のお願いを、サブインダストリーごとの Google チャットのスペースに送る（エディタから実行する。確かめてから送る）。
 * 会議日（空欄なら会議日程の一覧の次の会議）とシートを選ぶ。文は MTG_OPTIONS.subChat.remind（締め切りは会議日の deadlineDays 日前）。
 * 今の入力状況（進行中の案件・必須項目が未入力・受注月を過ぎている・最終更新から noUpdateDays 日以上）をシートごとに数えて付ける
 */
function remindBusinessMtgUpdate() {
  mtgSubChatSend_('remind');
}

/**
 * サブインダストリーの事前資料（その会議日の本番の PDF）のご案内を、サブインダストリーごとの Google チャットのスペースに送る
 * （エディタから実行する。確かめてから送る）。会議日（空欄なら会議日程の一覧の、今日以降でいちばん近い会議）とシートを選ぶ。
 * PDF は会議のフォルダ（無ければ保存先のフォルダ）の「yyyyMMdd_シート名_…用資料.pdf」。無いシートには送らない
 */
function notifySubBusinessMtgSummaries() {
  mtgSubChatSend_('summary');
}

/** remind・summary の共通：会議日とシートを聞き、送る文を作って確かめ、シートごとに送る（送れないシートがあっても、ほかは送る） */
function mtgSubChatSend_(kind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const title = CRED_OPTIONS.menuTitle;
  const what = kind === 'remind' ? '管理フォーマットの更新のお願い' : '事前資料のご案内';
  const registered = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name) && mtgSubChatWebhook_(name));
  if (!registered.length) {
    throw new Error('サブインダストリーごとの Google チャットのスペースが登録されていません。先に setBusinessMtgSubChats() で webhook の URL を登録してください。');
  }
  const meeting = mtgSubChatMeeting_(ui, what);
  if (!meeting) return;
  const names = mtgAskSheets_(ss, ui, what + 'を送ります（会議日 ' + meeting.label + '）。\nwebhook を登録したシート：' + registered.join('・'));
  if (!names) return;
  const now = new Date();
  const plans = names.map(name => {
    if (!mtgSubChatWebhook_(name)) return {name, skip: 'webhook が未登録（setBusinessMtgSubChats() で登録）'};
    if (kind === 'remind') return {name, text: mtgSubChatText_('remind', name, meeting, {now, sheet: mtgSheetLink_(ss, name), stats: mtgUpdateStats_(ss, name, now)})};
    const pdf = mtgStoredPdf_(ss, name, meeting.date);
    if (!pdf) return {name, skip: '会議日 ' + meeting.date + ' の本番の事前資料（PDF）が無い（createBusinessMtgSummaries() で作る）'};
    return {name, text: mtgSubChatText_('summary', name, meeting, {file: {name: pdf.getName(), url: pdf.getUrl()}})};
  });
  const sending = plans.filter(p => p.text);
  const skipped = plans.filter(p => p.skip);
  const list = sending.map(p => '・' + p.name + '：' + mtgChatSpace_(mtgSubChatWebhook_(p.name))).join('\n') +
    (skipped.length ? '\n\n送らないシート：\n' + skipped.map(p => '・' + p.name + '：' + p.skip).join('\n') : '');
  if (!sending.length) {
    ui.alert(title, what + 'を送れるシートがありません。\n' + list, ui.ButtonSet.OK);
    return;
  }
  const preview = mtgSubChatPlain_(sending[0].text);
  const answer = ui.alert(title, what + 'を、次のスペースにすぐに送ります。よろしいですか？\n' + list + '\n\n（例：' + sending[0].name + ' に送る文）\n' + preview,
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const results = sending.map(p => mtgSubChatPost_(p.name, p.text));
  const message = what + 'を送りました（' + results.filter(r => !r.error).length + '/' + results.length + 'スペース）。\n' +
    results.map(r => '・' + r.name + '：' + (r.error ? '送れませんでした（' + r.error + '）' : '送りました')).join('\n') +
    (skipped.length ? '\n\n送らなかったシート：\n' + skipped.map(p => '・' + p.name + '：' + p.skip).join('\n') : '');
  console.log(message);
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * 送る会議日を聞く。空欄なら会議日程の一覧（インデックスの RX Business Meeting）の、今日以降でいちばん近い会議日。
 * {date: 'yyyy/MM/dd', label: 会議日程（「2026年10月23日（金）9:00〜10:30」）}。キャンセルなら null
 */
function mtgSubChatMeeting_(ui, what) {
  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const next = mtgIndexMeetings_(MTG_OPTIONS.index.kind).map(m => m.date).filter(date => date >= today)[0] || '';
  const time = MTG_OPTIONS.chat.meetingTime;
  const response = ui.prompt(CRED_OPTIONS.menuTitle, what + 'の会議日を入力してください（例：10/23）。' +
    (next ? '\n空欄なら、会議日程の次の会議（' + mtgShortDate_(next) + '）。' : '') +
    (time ? '\n時刻を省くと「' + time + '」にします。ほかの時刻のときは「10/23 13:00〜14:30」のように入力してください。' : ''), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const text = String(response.getResponseText() || '').normalize('NFKC').trim() || (next ? mtgShortDate_(next) : '');
  if (!text) throw new Error('会議日を入力してください（会議日程の一覧に、今日以降の会議がありません）。');
  const label = mtgChatSchedule_(text, now);
  const at = /\s*(\d{1,2}:\d{2}.*)$/.exec(text);
  return {date: mtgParseDate_(at ? text.slice(0, at.index) : text, now), label};
}

/**
 * 送る文（MTG_OPTIONS.subChat の kind の行の、{meeting}・{name}・{date}・{deadline}・{now}・{stats}・{sheet}・{file} を置き換える）。
 * リンクは <URL|名前>（チャットでは名前で出る）
 */
function mtgSubChatText_(kind, name, meeting, values) {
  const o = MTG_OPTIONS.subChat;
  const tz = Session.getScriptTimeZone();
  const link = item => (item ? '<' + item.url + '|' + String(item.name).replace(/[|<>]/g, ' ') + '>' : '');
  const p = meeting.date.split('/').map(Number);
  const due = new Date(Date.UTC(p[0], p[1] - 1, p[2] - o.deadlineDays));
  const dueDate = due.getUTCFullYear() + '/' + String(due.getUTCMonth() + 1).padStart(2, '0') + '/' + String(due.getUTCDate()).padStart(2, '0');
  const deadline = mtgDateLabel_(dueDate).replace(/^\d+年/, '') + (o.deadlineTime ? ' ' + o.deadlineTime : '');
  const map = {meeting: MTG_OPTIONS.meetingName, name, date: meeting.label, deadline,
    now: values.now ? Utilities.formatDate(values.now, tz, 'M/d HH:mm') : '', stats: (values.stats || []).join('\n'),
    sheet: link(values.sheet), file: link(values.file)};
  return o[kind].join('\n').replace(/\{(\w+)\}/g, (all, key) => (key in map ? map[key] : all));
}

/** 確かめの画面に出す形（リンク <URL|名前> を「名前（URL）」に） */
function mtgSubChatPlain_(text) {
  return text.replace(/<([^|>]+)\|([^>]*)>/g, '$2（$1）');
}

/** 管理フォーマットのシートのリンク {name: シート名, url: そのシートを開く URL} */
function mtgSheetLink_(ss, name) {
  return {name: '「' + name + '」シート', url: ss.getUrl() + '#gid=' + ss.getSheetByName(name).getSheetId()};
}

/**
 * 今の入力状況（更新のお願いに付ける。読むだけ）：進行中の案件の数と、必須項目が未入力・受注月を過ぎている・最終更新から noUpdateDays 日以上
 * （最終更新日時の列があるとき）の案件の数。0件のものは書かない（どれも0件なら、そう書く）。['・進行中の案件 34件', '・必須項目（…）が未入力の案件 5件', …]
 */
function mtgUpdateStats_(ss, name, now) {
  const o = MTG_OPTIONS;
  const table = mtgReadDeals_(ss.getSheetByName(name), DIFF_RULES[name]);
  const active = table.deals.filter(d => d.kind === 'active');
  const month = mtgMonthIndex_(mtgMonth_(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM')));
  const checks = [
    ['必須項目（' + o.required.map(label => String(label).replace(/\s+/g, '')).join('・') + '）が未入力の案件', active.filter(d => d.missing.length).length],
    ['受注月を過ぎている案件', table.has.orderMonth ? active.filter(d => d.orderMonth && mtgMonthIndex_(d.orderMonth) < month).length : 0],
    ['最終更新から ' + o.noUpdateDays + '日以上たっている案件',
      table.hasStamp ? active.filter(d => d.updatedAt != null && now.getTime() - d.updatedAt >= o.noUpdateDays * 86400000).length : 0]
  ].filter(([, count]) => count > 0);
  return ['・進行中の案件 ' + active.length + '件'].concat(checks.length ? checks.map(([label, count]) => '・' + label + ' ' + count + '件')
    : ['・必須項目の未入力・受注月を過ぎている・長く更新の無い案件はありません']);
}

/** シートの webhook の URL（スクリプトのプロパティ。無ければ null） */
function mtgSubChatWebhook_(name) {
  return PropertiesService.getScriptProperties().getProperty(MTG_OPTIONS.subChat.webhookKeyPrefix + name);
}

/** 設定の見せ方（「スペース AAAA… に送る」「未登録」。鍵・トークンは出さない） */
function mtgSubChatState_(name) {
  const webhook = mtgSubChatWebhook_(name);
  return webhook ? mtgChatSpace_(webhook) + ' に送る' : '未登録';
}

/** シートのスペースに送る。{name, error}（送れなければ error に理由。止めない。URL は鍵を含むので出さない） */
function mtgSubChatPost_(name, text) {
  try {
    mtgPostChat_(mtgSubChatWebhook_(name), text);
    return {name, error: ''};
  } catch (error) {
    return {name, error: String(error && error.message || error).split('\n')[0]};
  }
}

/**
 * 前の版で、会議日の2日前に自動で送っていた時間主導のトリガーの関数。今は自動では送らない。
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
 * トリガーは持ち主にしか見えないので、前の版で通知を設定した人が setBusinessMtgSubChats() を実行したとき（またはトリガーが動いたとき）に外れる
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

/** シートを聞く（空欄ならすべて。「1,3」のように複数も）。heading は最初の行。extra はシートの後ろに並べる名前（統合版など）。やめたら null。 */
function mtgAskSheets_(ss, ui, heading, extra) {
  const sheets = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  if (!sheets.length) throw new Error('追跡しているシート（test.gs の DIFF_RULES）が見つかりません。');
  const all = sheets.concat(extra || []);
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


/** 完成した結果の件数の知らせ（「変化のある案件 12・停滞 5・確認事項 7」） */
function mtgCountsText_(summary) {
  return '変化のある案件 ' + summary.kpi.changed + '・停滞 ' + summary.stalled.length + '・確認事項 ' + summary.alerts.length;
}

/**
 * 会議日（meeting）の本番の資料（本番の記録と PDF）が、まだできていないサブインダストリー（追跡シートの順）。サブインダストリーごとに作るときの残り。
 * 空なら、すべてそろった（統合版を作れる）
 */
function mtgMeetingPending_(ss, meeting) {
  const props = PropertiesService.getDocumentProperties();
  return Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name))
    .filter(name => !mtgMeetings_(props, name).some(m => !m.baseline && m.date === meeting) || !mtgStoredPdf_(ss, name, meeting));
}

/**
 * names のシートの会議サマリの中身（mtgSummarize_）を、前回の会議（今回より前でいちばん新しい会議。そのあとに基準があれば基準）からの変化で作る。
 * 戻り値 {plans: [{name, meetings, previous, legacy, since, baseline, draft}], summaries}（読むだけ。記録しない）
 */
function mtgSummaries_(ss, names, now, final, meeting, point) {
  const o = MTG_OPTIONS;
  const props = PropertiesService.getDocumentProperties();
  const scheduled = mtgScheduledPrevious_(meeting);
  const plans = names.map(name => {
    const meetings = mtgMeetings_(props, name);
    // 前回：今回より前でいちばん新しい会議。そのあとに基準を記録していれば（今回の会議日まで）基準。どちらも無ければ会議日程の一覧の前の会議日。
    // 比べる時点を選んだ（point）なら、その時点
    const {previous, legacy} = mtgPreviousWith_(props, name, meeting, scheduled, point);
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
  // 前回の時点コピーがあれば、前回からの変化はそれと比べる（開けなければ、変更の記録から出して知らせる）
  const copies = plans.map(p => (p.previous && p.previous.copy ? mtgOpenCopy_(p.previous.copy, p.name) : null));
  const summaries = plans.map((p, i) => mtgSummarize_(ss, p.name, Object.assign({
    since: p.since, previous: p.previous, legacy: p.legacy, baseline: p.baseline, draft: p.draft, final,
    copy: copies[i] && copies[i].table ? copies[i] : null, copyError: copies[i] && copies[i].error || ''
  }, context)));
  return {plans, summaries};
}

/**
 * 統合版（全サブインダストリーの資料を1つにまとめた資料。mtgCombinedSlides_）を作り、PDF を保存先に保存する。
 * 本番は会議のフォルダ（事前資料の格納先 MTG_OPTIONS.meetingFolder の中。各サブインダストリーの資料と同じフォルダ）に格納し、インデックスの統合版の列に書く（できなくても止めない）。
 * 戻り値 {name, file: PDF, pages, stored, storeError, index}
 */
function mtgCreateCombined_(summaries, meeting, final, folder) {
  const name = MTG_OPTIONS.combined.name;
  const slides = mtgCombinedSlides_(summaries, meeting, final);
  const out = mtgExportDeck_(mtgFileName_(name, meeting, final), slides, folder);
  const result = {name, file: out.pdf, deck: out.deck, pages: slides.length};
  if (final) {
    const stored = mtgStoreToMeetingFolder_(name, meeting, out.pdf);
    if (stored.error) result.storeError = stored.error;
    else {
      result.stored = stored;
      result.index = mtgWriteIndex_(name, meeting, stored);
    }
  }
  return result;
}

/**
 * 統合版（全サブインダストリー）を作る（エディタから実行）。すべてのサブインダストリーの資料を createBusinessMtgSummaries() で作ったあとに実行する
 * （資料を作る関数とは分けてある。1回の実行の6分の上限に収めるため）。個別の資料を作り直したあとも、これで作り直す。
 * 画面（MeetingRunDialog）で会議日程をプルダウンで選び、すべての追跡シートの会議サマリを今の値で作り直して1つの資料にする（会議日・時点コピーは
 * 記録しない。前回からの変化は、個別の資料と同じく前回の会議時点と比べる）。本番として作り、会議のフォルダに格納してインデックスに書く。
 * 実行の条件：すべてのサブインダストリーに、その会議日の本番の記録と PDF（会議のフォルダか保存先）があること。満たさなければ、足りないものを知らせて
 * 何も作らずに終える（試作も作らない）。資料を作っている途中（自動で続けている途中）なら、終わるまで待ってもらう
 */
function createBusinessMtgCombinedSummary() {
  mschOpenRunDialog_('combined', true);
}

/** 統合版を作れるか（作れなければ {error: 理由}、作れれば {text: 確かめの文}） */
function mtgCombinedCheck_(meeting, point) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const o = MTG_OPTIONS.combined;
  if (!o.enabled) return {error: '統合版を使わない設定です（MeetingSummary.gs の MTG_OPTIONS.combined.enabled）。'};
  if (mtgJobLoad_()) {
    return {error: 'サブインダストリーの資料を作っている途中です（自動で続けています）。終わってから（showBusinessMtgSummariesResult() で確かめてから）実行してください。'};
  }
  const names = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  if (!names.length) return {error: '追跡しているシート（test.gs の DIFF_RULES）が見つかりません。'};
  // 実行の条件：すべてのサブインダストリーの、この会議日の本番の資料（PDF）ができている（インデックスのその会議日の行に空のセルが無い）。満たさなければ作らずに終える
  const missing = mtgMeetingPending_(ss, meeting);
  if (missing.length) {
    return {error: '統合版を作れません。会議日 ' + meeting + ' の本番の資料が、まだできていないサブインダストリーがあります（' + missing.length + '件）：\n' +
      missing.map(name => '・' + name).join('\n') + '\n\n' +
      'createBusinessMtgSummaries() で、すべてのサブインダストリーの本番の資料を作ってから（「' + MTG_OPTIONS.index.sheet + '」のその会議日の行に空のセルが無くなってから）、' +
      'もう一度実行してください。'};
  }
  return {text: '「' + o.name + '」の資料（' + names.length + 'サブインダストリーを1つにまとめた統合版）を、今のシートの値で作ります。\n\n' +
    '会議日 ' + mtgDateLabel_(meeting) + ' の本番の資料として作り、会議のフォルダに格納します（会議日の記録・時点コピーは変えません）。\n' +
    '比べる時点：' + (point ? mtgChosenText_(point) + '（全シート共通）' : '前回の会議（シートごと）')};
}

/** 統合版を作る（画面の「作成する」から。作れるかは先に mtgCombinedCheck_ で確かめる）。戻り値は終わりの知らせ */
function mtgCombinedStart_(meeting, point) {
  const check = mtgCombinedCheck_(meeting, point);
  if (check.error) throw new Error(check.error);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  const {summaries} = mtgSummaries_(ss, names, new Date(), true, meeting, point);
  mtgRequireSlides_();
  const folder = mtgFolder_(ss);
  const message = mtgCombinedLines_(mtgCombinedPlain_(mtgCreateCombined_(summaries, meeting, true, folder))).concat(['', '保存先：' + folder.getUrl()]).join('\n');
  console.log(message);
  return message;
}

/** サブインダストリー（name）の、会議日（meeting）の本番の資料（PDF）。会議のフォルダ（無ければ保存先のフォルダ）にあるもの。無ければ null */
function mtgStoredPdf_(ss, name, meeting) {
  const fileName = mtgFileName_(name, meeting, true) + '.pdf';
  const target = mtgMeetingFolder_(name, meeting, true);
  const folders = [target.folder, mtgFolder_(ss)].filter(Boolean);
  for (let i = 0; i < folders.length; i++) {
    const files = folders[i].getFilesByName(fileName);
    if (files.hasNext()) return files.next();
  }
  return null;
}

/** 統合版の結果の知らせ（終わりの画面の行） */
function mtgCombinedLines_(combined) {
  const name = MTG_OPTIONS.combined.name;
  if (!combined) return [];
  if (!combined.fileName) {
    return ['※ 統合版（' + name + '）は作りませんでした（理由：' + combined.error + '）。' +
      'エディタで createBusinessMtgCombinedSummary() を実行すると、統合版だけを作ります。'];
  }
  return ['「' + name + '」（統合版）：' + combined.fileName + '（' + combined.pages + 'ページ）\n　' + combined.fileUrl +
    (combined.folderName ? '\n　会議のフォルダ「' + combined.folderName + '」' + (combined.created ? 'を作って' : 'に') + '格納しました：' + combined.folderUrl : '') +
    (combined.indexError ? '\n　※ ' + combined.indexError : '') +
    (combined.storeError ? '\n　※ 会議のフォルダに格納できませんでした（' + combined.storeError + '）。資料は保存先にあります。' +
      'フォルダを用意してから、エディタで storeBusinessMtgSummaries() を実行し「' + name + '」を選んで格納してください。' : '')];
}

/** 統合版の結果（mtgCreateCombined_）を、保存できる形（文字だけ）に */
function mtgCombinedPlain_(c) {
  return {fileName: c.file.getName(), fileUrl: c.file.getUrl(), pages: c.pages,
    folderName: c.stored ? c.stored.folder.getName() : '', folderUrl: c.stored ? c.stored.folder.getUrl() : '', created: !!(c.stored && c.stored.created),
    indexError: c.index && c.index.error || '', storeError: c.storeError || ''};
}

/**
 * シートの会議の記録 [{date: 'yyyy/MM/dd', cutoff: 資料を作った時刻（ms）}]（会議日の古い順）。
 * setBusinessMtgBaseline() の基準も同じ形で {date: 記録した日, cutoff: 記録した時刻, baseline: true}（1つだけ）
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

/* ---------------- 会議のフォルダへの格納 ---------------- */

/**
 * 作成済みの本番の事前資料（PDF）を、会議のフォルダに格納する（エディタから実行）。作ったときに格納できなかった
 * （会議のフォルダがまだ無かった・権限が無かったなど）ときに、フォルダを用意してから実行する。シートと会議日を聞き、
 * 保存先（MTG_OPTIONS.folderId）にある本番の PDF（mtgFileName_。20261023_スーパー・GMS_RX Business Meeting用資料.pdf）を、会議のフォルダに移す（同じ名前の古い版はゴミ箱に移す）。
 * もう会議のフォルダにあれば移さずに、インデックスだけ書き直す。
 * 会議資料のインデックス（MTG_OPTIONS.index）にも書き直す（インデックスに書けなかったときもこれで直す）
 */
function storeBusinessMtgSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const combined = MTG_OPTIONS.combined;
  const names = mtgAskSheets_(ss, ui, '作成済みの事前資料（PDF）を、会議のフォルダに格納します。', combined.enabled ? [combined.name] : []);
  if (!names) return;
  const response = ui.prompt(CRED_OPTIONS.menuTitle, '格納する資料の会議日を入力してください（例：10/23）。', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const text = String(response.getResponseText() || '').trim();
  const meeting = text && mtgParseDate_(text, new Date());
  if (!meeting) throw new Error('「' + text + '」を会議日として読めません。10/23 か 2026/10/23 のように入力してください。');
  const folder = mtgFolder_(ss);
  const lines = names.map(name => {
    const fileName = mtgFileName_(name, meeting, true) + '.pdf';
    const files = folder.getFilesByName(fileName);
    const pdf = files.hasNext() ? files.next() : null;
    let stored;
    if (pdf) {
      stored = mtgStoreToMeetingFolder_(name, meeting, pdf);
    } else {
      // 保存先に無ければ、もう会議のフォルダにある（格納済み）か見る。あればインデックスだけ書き直す
      const target = mtgMeetingFolder_(name, meeting, true);
      const inside = target.folder ? target.folder.getFilesByName(fileName) : null;
      if (!inside || !inside.hasNext()) {
        return '「' + name + '」：本番の資料（' + fileName + '）が保存先・会議のフォルダにありません。createBusinessMtgSummaries() で作ってください。';
      }
      stored = {folder: target.folder, file: DriveApp.getFileById(inside.next().getId()), already: true};
    }
    if (stored.error) return '「' + name + '」：格納できませんでした（' + stored.error + '）。';
    const index = mtgWriteIndex_(name, meeting, stored);
    return '「' + name + '」：会議のフォルダ「' + stored.folder.getName() + '」' + (stored.already ? 'に格納済みです' : (stored.created ? 'を作って' : 'に') + '格納しました') +
      '：' + stored.folder.getUrl() +
      (index.error ? '\n　※ ' + index.error : '\n　「' + index.title + '」シートに書きました：' + index.url);
  });
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/**
 * 事前資料の PDF（pdf）を、そのシートの会議日のフォルダに移す（保存先と会議のフォルダに同じ PDF が2つできないように。同じ名前の古い版はゴミ箱に移す）。
 * 戻り値 {folder, file, created: 会議のフォルダを作ったか} か {error: 理由}（止めない。移せなければ PDF は保存先に残る）
 */
function mtgStoreToMeetingFolder_(name, meeting, pdf, kind) {
  const target = mtgMeetingFolder_(name, meeting, false, kind);
  if (target.error) return target;
  try {
    const fileName = pdf.getName();
    const file = DriveApp.getFileById(pdf.getId());
    file.moveTo(target.folder);
    mtgTrashOld_(target.folder, [fileName], [pdf.getId()]);
    mtgMoveSameNamePptx_(fileName, target.folder);
    return {folder: target.folder, file, created: !!target.created};
  } catch (error) {
    return {error: '「' + target.folder.getName() + '」に移せない：' + (error && error.message || error)};
  }
}

/**
 * PDF（fileName）と同じ名前の PowerPoint（.pptx。mtgExportPptx_）が保存先にあれば、会議のフォルダ（folder）に移す（同じ名前の古い版はゴミ箱へ）。
 * 移せなくても止めない（PowerPoint は保存先に残る）
 */
function mtgMoveSameNamePptx_(fileName, folder) {
  const pptxName = String(fileName).replace(/\.pdf$/i, '') + '.pptx';
  try {
    const files = mtgFolder_(SpreadsheetApp.getActiveSpreadsheet()).getFilesByName(pptxName);
    if (!files.hasNext()) return;
    const id = files.next().getId();
    DriveApp.getFileById(id).moveTo(folder);
    mtgTrashOld_(folder, [pptxName], [id]);
  } catch (error) {
    console.warn('PowerPoint（' + pptxName + '）を会議のフォルダに移せませんでした：' + (error && error.message || error));
  }
}

/**
 * シート（name）の会議のフォルダ：事前資料の格納先（mtgMeetingFolderSetting_。全サブインダストリー・統合版で共通）の中の、
 * 名前が会議日の meetingFolderDate（yyMMdd）で始まり、そのすぐあとが数字でなく、会議 kind（会議の名前。省けば MTG_OPTIONS.meetingName）のフォルダ
 * （mtgFolderIsMeeting_。261023_RX Business Meeting_第5回）。{folder} か {error: 理由}（設定が無い・開けない・見つからない・2つ以上ある）。
 * 見つからなければ作る（createMeetingFolder。noCreate なら作らない。名前は meetingFolderName）
 */
function mtgMeetingFolder_(name, meeting, noCreate, kind) {
  const meetingName = kind || MTG_OPTIONS.meetingName;
  const setting = mtgMeetingFolderSetting_();
  const id = mtgFolderIdOf_(setting.value);
  if (setting.value && !id) return {error: '事前資料の格納先の設定（' + setting.where + '）をフォルダの URL か ID として読めない'};
  if (!id) return {error: '事前資料の格納先が設定されていない（' + setting.where + '）'};
  let parent;
  try {
    parent = DriveApp.getFolderById(id);
  } catch (_) {
    return {error: '事前資料の格納先のフォルダ（ID：' + id + '）を開けない（実行するアカウントに編集の権限があるか確かめてください）'};
  }
  const p = meeting.split('/');
  const prefix = MTG_OPTIONS.meetingFolderDate.replace('yyyy', p[0]).replace('yy', p[0].slice(2)).replace('MM', p[1]).replace('dd', p[2]);
  const found = [];
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = String(folder.getName());
    if (folderName.indexOf(prefix) === 0 && !/^\d/.test(folderName.slice(prefix.length)) && mtgFolderIsMeeting_(folderName, meetingName) &&
      !folder.isTrashed()) found.push(folder);
  }
  if (!found.length) {
    if (!MTG_OPTIONS.createMeetingFolder || noCreate) {
      return {error: '「' + parent.getName() + '」の中に「' + prefix + '_」で始まる ' + meetingName + ' のフォルダが無い'};
    }
    const folderName = MTG_OPTIONS.meetingFolderName.replace('{date}', prefix).replace('{meeting}', meetingName)
      .replace('{n}', String(mtgNextMeetingNumber_(parent, prefix, meetingName)));
    try {
      return {folder: parent.createFolder(folderName), created: true};
    } catch (error) {
      return {error: '「' + parent.getName() + '」の中に会議のフォルダ「' + folderName + '」を作れない：' + (error && error.message || error)};
    }
  }
  if (found.length > 1) {
    return {error: '「' + parent.getName() + '」の中に「' + prefix + '_」で始まる ' + meetingName + ' のフォルダが ' + found.length + '個ある（「' +
      found.map(f => f.getName()).join('」「') + '」。1つにしてください）'};
  }
  return {folder: found[0]};
}

/**
 * 会議のフォルダを探す・作るフォルダ（事前資料の格納先）の設定 {value: URL か ID, where: 設定の場所}。
 * 全サブインダストリー・統合版で共通（2026/10/26 から）：スクリプトのプロパティ「MTG_MEETING_FOLDER」→ MTG_OPTIONS.meetingFolder の順
 */
function mtgMeetingFolderSetting_() {
  return mtgOutputSetting_([MTG_OPTIONS.folderProps.meeting], MTG_OPTIONS.meetingFolder, 'MeetingSummary.gs の MTG_OPTIONS.meetingFolder');
}

/**
 * 出力先の設定 {value, where: 設定の場所, prop: 使っているスクリプトのプロパティの名前（コードの設定なら空）, keys: 見たプロパティの名前, code: コードの設定の場所}。
 * keys のスクリプトのプロパティを順に見て、値があればそれ、どれも無ければコードの設定（codeValue）
 */
function mtgOutputSetting_(keys, codeValue, codeWhere) {
  let props = null;
  try { props = PropertiesService.getScriptProperties(); } catch (_) {}
  for (const key of keys) {
    const value = props ? String(props.getProperty(key) || '').trim() : '';
    if (value) return {value, where: 'スクリプトのプロパティ「' + key + '」', prop: key, keys, code: codeWhere};
  }
  return {value: String(codeValue || '').trim(), where: codeWhere, prop: '', keys, code: codeWhere};
}

/**
 * 今の出力先と素材（ひな形）を一覧にする（エディタから実行する。読むだけで、フォルダは作らない）。項目ごとに、使っているフォルダ・ファイル（○ 開ける・✗ 開けない）と、
 * スクリプトのプロパティに値を入れているか（入れていなければコードの設定を使う）を出す。スクリプトのプロパティを変えたあとの確かめに使う。画面に出せなければ実行ログだけ
 */
function showOutputFolders() {
  const o = MTG_OPTIONS;
  const open = (id, kind) => {
    try {
      return '○ ' + (kind === 'file' ? DriveApp.getFileById(id) : DriveApp.getFolderById(id)).getName() + '（ID：' + id + '）';
    } catch (_) {
      return '✗ 開けません（ID：' + id + '。実行するアカウントに権限があるか確かめてください）';
    }
  };
  // 項目：1行目に使っているもの、2行目にスクリプトのプロパティの状態
  const item = (label, setting, empty, kind) => {
    const id = kind === 'file' ? mtgFileIdOf_(setting.value) : mtgFolderIdOf_(setting.value);
    const used = !setting.value ? empty : !id ? '✗ URL か ID として読めません：' + setting.value : open(id, kind);
    const prop = setting.prop
      ? 'スクリプトのプロパティ「' + setting.prop + '」：設定あり（これを使用）'
      : 'スクリプトのプロパティ「' + setting.keys[0] + '」：未設定 → コードの設定（' + setting.code + '）を使用';
    return '・' + label + '\n　使っているもの：' + used + '\n　' + prop;
  };
  const lines = ['出力先・素材の一覧（○＝開ける　✗＝開けない。スクリプトのプロパティが未設定なら、コードの設定を使います）', '', '■ 出力先'];
  lines.push(item('事前資料などの保存先（作る場所・試作）', mtgSaveFolderSetting_(), '（無し。このスプレッドシートと同じフォルダの中に「' + o.folderName + '」を作る）'));
  lines.push(item('時点コピーの保存先', mtgCopyFolderSetting_(), '（無し。保存先の中の「' + o.copyFolderName + '」）'));
  if (typeof apDestFolderSetting_ === 'function') lines.push(item('アカウントプランシートの保存先', apDestFolderSetting_(), '✗ 無し'));
  lines.push(item('事前資料の格納先（この中に会議日のフォルダ「yyMMdd_RX Business Meeting_第N回」を作り、全サブインダストリーと統合版の資料を入れる）',
    mtgMeetingFolderSetting_(), '✗ 無し（本番の資料を会議のフォルダに格納できません）'));
  // 素材（ひな形）：表紙・扉のひな形は、プロパティ → コードの templateId → 前の版の登録の順
  lines.push('', '■ 素材（ひな形）');
  const legacy = PropertiesService.getDocumentProperties().getProperty(o.templateKey) || '';
  const template = mtgOutputSetting_([o.templateUrlKey], o.templateId || legacy,
    o.templateId ? 'MeetingSummary.gs の MTG_OPTIONS.templateId' : legacy ? '前の版の登録（setMeetingTemplate()）' : 'MeetingSummary.gs の MTG_OPTIONS.templateId');
  lines.push(item('事前資料の表紙・扉のひな形', template, '（無し。表紙・扉はコードで描くので、素材が無くても作れる）', 'file'));
  if (typeof apTemplateSetting_ === 'function') lines.push(item('アカウントプランシートのひな形', apTemplateSetting_(), '✗ 無し', 'file'));
  // 前の版のプロパティ（サブインダストリーごとの格納先）は使わないので、消してよいと知らせる
  const former = PropertiesService.getScriptProperties().getKeys().filter(k => k.indexOf(o.formerFolderPropPrefix) === 0).sort();
  if (former.length) {
    lines.push('', '■ 使わなくなったスクリプトのプロパティ（' + former.length + '件。消してかまいません。removeFormerMeetingFolderProps() でまとめて消せます）',
      former.map(k => '・' + k).join('\n'));
  }
  const text = lines.join('\n');
  console.log(text);
  const ui = diffUiOrNull_();
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, text, ui.ButtonSet.OK);
  return lines;
}

/**
 * 前の版のスクリプトのプロパティ（サブインダストリーごとの事前資料の格納先「MTG_MEETING_FOLDER_<シート名>」）を消す（エディタから実行する。確かめてから）。
 * 2026/10/26 から格納先は全サブインダストリー共通の1つ（MTG_MEETING_FOLDER）になり、使わない
 */
function removeFormerMeetingFolderProps() {
  const ui = mtgUi_();
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys().filter(k => k.indexOf(MTG_OPTIONS.formerFolderPropPrefix) === 0).sort();
  if (!keys.length) {
    ui.alert(CRED_OPTIONS.menuTitle, '使わなくなったプロパティ（' + MTG_OPTIONS.formerFolderPropPrefix + '<シート名>）はありません。', ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert(CRED_OPTIONS.menuTitle, '次のスクリプトのプロパティを消します（今は使っていません。フォルダや資料は消えません）。よろしいですか？\n\n' +
    keys.map(k => '・' + k).join('\n'), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  keys.forEach(k => props.deleteProperty(k));
  ui.alert(CRED_OPTIONS.menuTitle, keys.length + '件のプロパティを消しました。', ui.ButtonSet.OK);
}

/** 出力先の設定のフォルダの ID（設定が無ければ空）。値があるのにフォルダの URL・ID として読めなければ止める */
function mtgOutputFolderId_(setting, label) {
  const id = mtgFolderIdOf_(setting.value);
  if (setting.value && !id) {
    throw new Error(label + 'の設定（' + setting.where + '）の値「' + setting.value + '」を、フォルダの URL か ID として読めません。' +
      'ドライブでフォルダを開いたときのアドレス欄の URL を入れてください。');
  }
  return id;
}

/**
 * 会議資料のインデックス（このスプレッドシートの MTG_OPTIONS.index.sheet）の、行＝会議日（meeting）・列＝サブインダストリー（name。統合版は combined.name）のセルに、
 * PDF（stored.file。index.link が 'folder' なら会議のフォルダ）をスマートチップで書く（できなければファイル名のリンク）。
 * シートが無ければ作り、1行目に見出し（日付・種別・統合版・追跡シートの名前（DIFF_RULES の順））を書く。
 * kind は種別の列の値（省けば MTG_OPTIONS.index.kind。RX Service Meeting の資料は、その種別の行の統合版の列に書く）。
 * 会議日の行（日付が同じで、種別が kind か空の行）が無ければ、index.order の順（既定は新しい会議日が上）になる位置に差し込み（並びが違えば先に並べ直す）、
 * サブインダストリーの列が無ければ右に足す。
 * 同じセルは上書き（作り直したとき）。書き込みはドキュメントロックの中（差分追跡の対象外のシートだが、資料作成はエディタから実行するので待ってよい）。
 * 戻り値 {url: シートの URL, title: シート名} か {error: 知らせる文}（止めない）
 */
function mtgWriteIndex_(name, meeting, stored, kind) {
  const o = MTG_OPTIONS.index;
  const rowKind = kind || o.kind;   // 会議名の列の値（RX Service Meeting の資料は SMTG_OPTIONS.index.kind）
  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(MTG_OPTIONS.indexLockWaitMs)) throw new Error('ほかの処理が実行中');
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const {sheet, layout} = mtgIndexPrepare_(ss);
      // サブインダストリーの列（前の名前で書いてあれば今の名前とみなす。無ければ右に足す）
      let col = layout.columns[name];
      if (!col) {
        col = Math.max(sheet.getLastColumn(), 1) + 1;
        diffEnsureGrid_(sheet, 1, col);
        sheet.getRange(1, col).setValue(credText_(name));
        mtgIndexFormat_(sheet.getRange(1, col, Math.max(sheet.getLastRow(), 1), 1));
        sheet.setColumnWidth(col, name === MTG_OPTIONS.combined.name ? o.widths.combined : o.widths.sheet);
      }
      const row = mtgIndexRow_(sheet, layout, meeting, rowKind).row;
      const target = o.link === 'folder' || !stored.file ? stored.folder : stored.file;
      const url = target.getUrl();
      const chip = !!o.chip && mtgWriteChip_(ss, sheet, row, col, url);
      if (!chip) {
        const fileName = target.getName();
        sheet.getRange(row, col).setRichTextValues([[SpreadsheetApp.newRichTextValue().setText(fileName).setLinkUrl(url)
          .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setUnderline(false).build()).build()]]);   // 黒・下線なし
      }
      return {url: ss.getUrl() + '#gid=' + sheet.getSheetId(), title: o.sheet, chip};
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return {error: '「' + o.sheet + '」シートに書き込めませんでした（' + (error && error.message || error) + '）。' +
      (error && error.hint || (kind ? '' : 'storeBusinessMtgSummaries() を実行すると書き直します。'))};
  }
}

/**
 * インデックスのシートと見出しの位置を用意する（ロックの中で呼ぶ）。シートが無ければ作る。前の版の見出し（種別）は今の見出し（会議名）に書き換える。
 * 前の形式（行＝サブインダストリー）なら、rebuildBusinessMtgIndex() の案内を付けて止める。{sheet, layout}
 */
function mtgIndexPrepare_(ss) {
  const o = MTG_OPTIONS.index;
  const sheet = mtgIndexSheet_(ss);
  const layout = mtgIndexLayout_(sheet);
  if (!layout) {
    const error = new Error('1行目に「' + o.dateHeader + '」の見出しがありません');
    error.hint = '前の形式（行がサブインダストリー）のシートは、エディタで rebuildBusinessMtgIndex() を実行すると、今の形式で作り直します。';
    throw error;
  }
  if (layout.kind && layout.kindStale) sheet.getRange(1, layout.kind).setValue(credText_(o.kindHeader));
  return {sheet, layout};
}

/**
 * インデックスの、会議日 meeting・会議名 kind の行（ロックの中で呼ぶ）。会議名が同じ（前の版の会議名も）か空の行を使い、前の版の会議名・空なら今の会議名に書き換える。
 * 無ければ、index.order の順（既定は新しい会議日が上）になる位置に差し込む（並びが違えば、先に並べ直す）。{row, added: 足したか}
 */
function mtgIndexRow_(sheet, layout, meeting, kind) {
  const o = MTG_OPTIONS.index;
  const kinds = mtgIndexKinds_(kind);   // 同じ会議とみなす会議名（前の版の会議名も）
  const desc = o.order !== 'asc';
  const lastRow = sheet.getLastRow();
  const read = () => (lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, Math.max(layout.date, layout.kind || 1)).getValues() : []);
  const kindOf = line => (layout.kind ? String(line[layout.kind - 1] == null ? '' : line[layout.kind - 1]).trim() : '');
  const dateOf = line => mtgIndexDate_(line[layout.date - 1]);
  const before = (a, b) => (desc ? a > b : a < b);   // a の行が b の行より上
  let lines = read();
  const dates = lines.map(dateOf).filter(Boolean);
  if (dates.some((d, i) => i && before(d, dates[i - 1]))) {
    sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 1)).sort({column: layout.date, ascending: !desc});
    lines = read();
  }
  let at = lines.findIndex(line => dateOf(line) === meeting && (!kindOf(line) || kinds.indexOf(kindOf(line)) >= 0));
  if (at >= 0) {
    const row = at + 2;
    if (layout.kind && kindOf(lines[at]) !== kind) sheet.getRange(row, layout.kind).setValue(credText_(kind));
    return {row, added: false};
  }
  let row;
  at = lines.findIndex(line => dateOf(line) && before(meeting, dateOf(line)));
  if (at >= 0) {
    row = at + 2;
    sheet.insertRowBefore(row);
  } else {
    row = lastRow + 1;
    diffEnsureGrid_(sheet, row, 1);
  }
  sheet.getRange(row, layout.date).setValue(Utilities.parseDate(meeting, Session.getScriptTimeZone(), 'yyyy/MM/dd')).setNumberFormat(o.dateFormat);
  if (layout.kind) sheet.getRange(row, layout.kind).setValue(credText_(kind));
  mtgIndexFormat_(sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 1)));
  sheet.setRowHeight(row, o.rowHeight);
  return {row, added: true};
}

/**
 * 会議の日程（会議日・会議名）の行を、インデックスに先に足す（Google カレンダーから読んだ日程。MeetingSchedule.gs）。ある行はそのまま。
 * entries = [{date: 'yyyy/MM/dd', kind: 会議名}]。{added, existing, url} か {error}
 */
function mtgIndexEnsureRows_(entries) {
  const o = MTG_OPTIONS.index;
  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(MTG_OPTIONS.indexLockWaitMs)) throw new Error('ほかの処理が実行中です。少し待ってからもう一度お試しください');
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const {sheet, layout} = mtgIndexPrepare_(ss);
      let added = 0;
      entries.forEach(e => { if (mtgIndexRow_(sheet, layout, e.date, e.kind).added) added++; });
      return {added, existing: entries.length - added, url: ss.getUrl() + '#gid=' + sheet.getSheetId()};
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return {error: '「' + o.sheet + '」シートに書き込めませんでした（' + (error && error.message || error) + '）。' + (error && error.hint || '')};
  }
}

/**
 * インデックスにある、会議名 kind の会議日（読むだけ）。[{date: 'yyyy/MM/dd', filled: その行に資料が書いてあるか}]（会議日の古い順。同じ日は1つ）。
 * 会議名が空の行（会議名の列ができる前の行）は RX Business Meeting の行とみなす。シートが無い・前の形式なら空
 */
function mtgIndexMeetings_(kind) {
  const o = MTG_OPTIONS.index;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(o.sheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const layout = mtgIndexLayout_(sheet);
  if (!layout) return [];
  const kinds = mtgIndexKinds_(kind);
  const width = Math.max(sheet.getLastColumn(), layout.date, layout.kind || 1);
  const lines = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  const shown = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getDisplayValues();
  const byDate = new Map();
  lines.forEach((line, i) => {
    const date = mtgIndexDate_(line[layout.date - 1]);
    const name = layout.kind ? String(line[layout.kind - 1] == null ? '' : line[layout.kind - 1]).trim() : '';
    if (!date || (name ? kinds.indexOf(name) < 0 : kind !== o.kind)) return;
    const filled = shown[i].some((v, c) => c + 1 !== layout.date && c + 1 !== layout.kind && String(v).trim() !== '');
    byDate.set(date, (byDate.get(date) || false) || filled);
  });
  return Array.from(byDate.keys()).sort().map(date => ({date, filled: byDate.get(date)}));
}

/**
 * インデックスのシート（無ければ作る。白紙なら見出しの行を用意する。ロックの中で呼ぶ）。
 * 見出し：日付・種別・統合版（MTG_OPTIONS.combined.name）・追跡シートの名前（DIFF_RULES の順）。太字・黒・上下中央、1行目を固定、タブは黒。使わない行・列は持たない
 */
function mtgIndexSheet_(ss) {
  const o = MTG_OPTIONS.index;
  let sheet = ss.getSheetByName(o.sheet);
  if (!sheet) {
    const active = ss.getActiveSheet();
    sheet = ss.insertSheet(o.sheet, ss.getSheets().length);
    if (active) active.activate();   // insertSheet は作ったシートを開くので、開いていたシートに戻す
  }
  if (sheet.getLastRow() > 0) return sheet;
  const headers = [o.dateHeader, o.kindHeader, MTG_OPTIONS.combined.name].concat(Object.keys(DIFF_RULES));
  diffEnsureGrid_(sheet, 2, headers.length);
  mtgIndexFormat_(sheet.getRange(1, 1, 1, headers.length).setValues([headers.map(credText_)]));
  const widths = [o.widths.date, o.widths.kind, o.widths.combined];
  headers.forEach((_, i) => sheet.setColumnWidth(i + 1, i < widths.length ? widths[i] : o.widths.sheet));
  sheet.setFrozenRows(1);
  if (o.tabColor) sheet.setTabColor(o.tabColor);
  diffTrimGrid_(sheet, headers.length, 2);
  return sheet;
}

/** インデックスの字の書式（太字・黒・上下中央・fontSize） */
function mtgIndexFormat_(range) {
  return range.setFontWeight('bold').setFontSize(MTG_OPTIONS.index.fontSize).setFontColor('#000000').setVerticalAlignment('middle');
}

/**
 * インデックスの見出し（1行目）から列の位置を読む。{date, kind, columns: {シート名（統合版は combined.name）: 列}}（列は 1 から。前の名前の見出しは今の名前に）。
 * 「日付」の見出しが無ければ null（前の形式（行＝サブインダストリー）など）
 */
function mtgIndexLayout_(sheet) {
  const o = MTG_OPTIONS.index;
  const width = Math.max(sheet.getLastColumn(), 1);
  const heads = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(h => String(h).trim());
  const date = heads.indexOf(o.dateHeader) + 1;
  if (!date) return null;
  const resolve = diffSheetNameResolver_();
  const columns = {};
  heads.forEach((h, i) => {
    const name = resolve(h);
    if ((DIFF_RULES[name] || name === MTG_OPTIONS.combined.name) && !columns[name]) columns[name] = i + 1;
  });
  // 会議名の列（前の版の見出し「種別」なら kindStale。mtgIndexPrepare_ が書き換える）
  let kind = heads.indexOf(o.kindHeader) + 1;
  const former = kind ? 0 : (o.formerKindHeaders || []).map(h => heads.indexOf(h) + 1).find(Boolean) || 0;
  return {date, kind: kind || former, kindStale: !kind && !!former, columns};
}

/**
 * 会議資料のインデックスを、各サブインダストリー・統合版の会議のフォルダ（DIFF_RULES・MTG_OPTIONS.combined の meetingFolder の中の「yyMMdd_RX Business Meeting_第N回」・前の版の「yyMMdd_第N回会議」）にある
 * 本番の事前資料（PDF。mtgFileName_）から書き直す（エディタから実行）。前の形式（行＝サブインダストリー）のシートは、中身を消して今の形式で作り直す。
 * 今の形式なら、ある行・列はそのまま（手で入れた値も残す）で、資料のセルだけ書き直す。何度実行してもよい
 */
function rebuildBusinessMtgIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const o = MTG_OPTIONS.index;
  const current = ss.getSheetByName(o.sheet);
  const old = !!current && current.getLastRow() > 0 && !mtgIndexLayout_(current);
  const answer = ui.alert(CRED_OPTIONS.menuTitle, '各サブインダストリーの会議のフォルダにある事前資料（PDF）から、「' + o.sheet + '」シートを書き直します。' +
    (old ? '\n今のシートは前の形式（行がサブインダストリー）なので、中身を消して、行＝会議日・列＝サブインダストリーの形式で作り直します。' : '') +
    '\nよろしいですか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  if (old) {
    const lock = LockService.getDocumentLock();
    lock.waitLock(MTG_OPTIONS.indexLockWaitMs);
    try {
      current.clear();
    } finally {
      lock.releaseLock();
    }
  }
  let written = 0;
  let index = null;
  const combined = MTG_OPTIONS.combined;
  const lines = Object.keys(DIFF_RULES).concat(combined.enabled ? [combined.name] : []).map(name => {
    const found = mtgStoredDecks_(name);
    if (found.error) return '・' + name + '：' + found.error;
    const errors = [];
    found.decks.forEach(d => {
      const r = mtgWriteIndex_(name, d.meeting, {folder: d.folder, file: d.file});
      if (r.error) errors.push(d.meeting + '：' + r.error);
      else {
        written++;
        index = r;
      }
    });
    return '・' + name + '：' + (found.decks.length ? found.decks.length + '件（' + found.decks.map(d => d.meeting).join('・') + '）' : '資料なし') +
      (errors.length ? '\n　※ ' + errors.join('\n　※ ') : '');
  });
  const message = '「' + o.sheet + '」シートに、資料のファイルを ' + written + '件書きました。' + (index ? index.url : '') + '\n\n' + lines.join('\n');
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/**
 * サブインダストリー（name。統合版は combined.name）の会議のフォルダにある本番の事前資料。{decks: [{meeting: 'yyyy/MM/dd', folder, file}]}（会議日の順）か {error: 理由}。
 * 会議のフォルダは、名前が meetingFolderDate（yyMMdd）の日付と「_」で始まるもの。PDF の名前は mtgFileName_（本番）
 */
function mtgStoredDecks_(name) {
  const setting = mtgMeetingFolderSetting_();
  const id = mtgFolderIdOf_(setting.value);
  if (setting.value && !id) return {error: '会議のフォルダの設定（' + setting.where + '）をフォルダの URL か ID として読めない'};
  if (!id) return {error: '会議のフォルダが設定されていない（' + setting.where + '）'};
  let parent;
  try {
    parent = DriveApp.getFolderById(id);
  } catch (_) {
    return {error: 'フォルダ（ID：' + id + '）を開けない'};
  }
  const decks = [];
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.isTrashed()) continue;
    const meeting = mtgMeetingOfFolder_(String(folder.getName()));
    if (!meeting) continue;
    const files = folder.getFilesByName(mtgFileName_(name, meeting, true) + '.pdf');
    if (files.hasNext()) decks.push({meeting, folder, file: DriveApp.getFileById(files.next().getId())});
  }
  decks.sort((a, b) => (a.meeting < b.meeting ? -1 : a.meeting > b.meeting ? 1 : 0));
  return {decks};
}

/** 会議のフォルダの名前（261023_RX Business Meeting_第5回・261023_第5回会議）から会議日 'yyyy/MM/dd'（MTG_OPTIONS.meetingFolderDate の形の日付と「_」で始まらなければ空） */
function mtgMeetingOfFolder_(folderName) {
  const keys = [];
  const pattern = MTG_OPTIONS.meetingFolderDate.replace(/yyyy|yy|MM|dd|[^A-Za-z]/g, t => {
    if (/^[yMd]/.test(t)) {
      keys.push(t);
      return '(\\d{' + t.length + '})';
    }
    return t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  });
  const m = new RegExp('^' + pattern + '_').exec(folderName.normalize('NFKC'));
  if (!m) return '';
  const part = {};
  keys.forEach((k, i) => { part[k] = m[i + 1]; });
  const year = part.yyyy || '20' + part.yy;
  const month = Number(part.MM);
  const day = Number(part.dd);
  if (!(month >= 1 && month <= 12 && day >= 1 && day <= 31)) return '';
  return year + '/' + ('0' + month).slice(-2) + '/' + ('0' + day).slice(-2);
}

/**
 * セル（row, col）に、Drive のファイル（url）のスマートチップ（リッチリンクのチップ）を書く。Sheets API（拡張サービス「Google Sheets API」）で書く
 * （SpreadsheetApp ではチップを書けないため）。拡張サービスが無い・書けないときは false（呼び出し元がリンクで書く）
 */
function mtgWriteChip_(ss, sheet, row, col, url) {
  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets) return false;
  try {
    SpreadsheetApp.flush();   // 列の差し込みなどを先に反映する（Sheets API は別の経路で書くため）
    Sheets.Spreadsheets.batchUpdate({requests: [{updateCells: {
      start: {sheetId: sheet.getSheetId(), rowIndex: row - 1, columnIndex: col - 1},
      // 文字は黒・下線なし（リンクの青・下線の書式が残っていると、チップの名前も青の下線になるため）
      rows: [{values: [{userEnteredValue: {stringValue: '@'}, chipRuns: [{startIndex: 0, chip: {richLinkProperties: {uri: url}}}],
        userEnteredFormat: {textFormat: {underline: false, foregroundColorStyle: {rgbColor: {red: 0, green: 0, blue: 0}}}}}]}],
      fields: 'userEnteredValue,chipRuns,textFormatRuns,userEnteredFormat.textFormat.underline,userEnteredFormat.textFormat.foregroundColorStyle'
    }}]}, ss.getId());
    return true;
  } catch (error) {
    console.warn('インデックスにチップを書けなかったので、リンクで書きます：' + (error && error.message || error));
    return false;
  }
}

/** インデックスの1行目の値を会議日 'yyyy/MM/dd' に（日付の値・「2026/10/23」の文字。読めなければ空） */
function mtgIndexDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const text = String(value == null ? '' : value).trim();
  return /^\d{4}\s*[\/\-.年]/.test(text) ? mtgParseDate_(text, new Date()) : '';
}

/** 種別 kind と、同じ会議とみなす前の版の種別（MTG_OPTIONS.index.formerKinds・SMTG_OPTIONS.index.formerKinds） */
function mtgIndexKinds_(kind) {
  const all = [MTG_OPTIONS.index].concat(typeof SMTG_OPTIONS !== 'undefined' ? [SMTG_OPTIONS.index] : []);
  const own = all.find(x => x.kind === kind);
  return [kind].concat(own && own.formerKinds || []);
}

/**
 * 次の会議のフォルダの回数：parent の中の、会議日が prefix（yyMMdd）より前の、同じ会議（kind。mtgFolderIsMeeting_）のフォルダ
 * （「yyMMdd_RX Business Meeting_第N回」「yyMMdd_第N回会議」）のいちばん大きい N ＋ 1（無ければ 1）
 */
function mtgNextMeetingNumber_(parent, prefix, kind) {
  let max = 0;
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folderName = String(folders.next().getName());
    const m = /^(\d{6})_.*?第\s*(\d+)\s*回/.exec(folderName.normalize('NFKC'));
    if (m && m[1] < prefix && mtgFolderIsMeeting_(folderName, kind || MTG_OPTIONS.meetingName)) max = Math.max(max, Number(m[2]));
  }
  return max + 1;
}

/**
 * 会議のフォルダの名前が、会議 meetingName のものか：名前に会議の名前か前の名前（全角・半角・大文字小文字・空白の違いは無視）を含む。どの会議の名前も含まない
 * 前の版のフォルダ（261023_第5回会議）は、RX Business Meeting（MTG_OPTIONS.meetingName）のものとみなす
 */
function mtgFolderIsMeeting_(folderName, meetingName) {
  const norm = v => String(v || '').normalize('NFKC').toLowerCase().replace(/\s/g, '');
  // 会議ごとの名前（今の名前と前の名前。formerMeetingNames）
  const meetings = [MTG_OPTIONS].concat(typeof SMTG_OPTIONS !== 'undefined' ? [SMTG_OPTIONS] : [])
    .map(o => ({name: o.meetingName, names: [o.meetingName].concat(o.formerMeetingNames || [])}));
  const own = meetings.filter(m => m.names.some(n => norm(folderName).indexOf(norm(n)) >= 0)).map(m => m.name);
  return own.length ? own.indexOf(meetingName) >= 0 : meetingName === MTG_OPTIONS.meetingName;
}

/** フォルダの URL（…/folders/ID、?id=ID）か ID から ID を取り出す（読めなければ空）。 */
function mtgFolderIdOf_(text) {
  const t = String(text || '').trim();
  const m = /\/folders\/([A-Za-z0-9_-]{10,})/.exec(t) || /[?&]id=([A-Za-z0-9_-]{10,})/.exec(t) || /^([A-Za-z0-9_-]{10,})$/.exec(t);
  return m ? m[1] : '';
}

/* ---------------- 時点コピー ---------------- */

/**
 * このスプレッドシート全体（すべてのシート・数式・書式・入力規則）を、新しいスプレッドシート（タイトルは「時点コピー_実行日時」＋ note）にコピーする
 * （Drive のファイルのコピー。値を1セルずつ写すより速い）。保存先は mtgCopyFolder_。比べるときは、コピーの追跡シートの値を読む
 * （追跡シートの数式は同じファイルの中のシート（定義_ステータス など）だけを参照するので、コピーでも保存した時点の値になる）。
 * 復旧の控えにも使える。コピーにはスクリプトも入る（トリガーは入らない）。戻り値 {id, title, url}
 */
function mtgSaveCopy_(ss, now, note) {
  const title = MTG_OPTIONS.copyName + '_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + (note || '');
  const folder = mtgCopyFolder_(ss);
  try {
    const file = DriveApp.getFileById(ss.getId()).makeCopy(title, folder);
    return {id: file.getId(), title, url: file.getUrl()};
  } catch (error) {
    throw new Error('時点コピーを保存できませんでした（' + (error && error.message || error) + '）。もう一度お試しください。');
  }
}

/**
 * 時点コピーの保存先（MTG_OPTIONS.copyFolderId。空なら会議サマリのフォルダの中の copyFolderName。作ったフォルダの ID をドキュメントのプロパティに残す）。
 * noCreate なら作らず、無い・開けなければ null（読むだけのとき）
 */
function mtgCopyFolder_(ss, noCreate) {
  const setting = mtgCopyFolderSetting_();
  const id = mtgOutputFolderId_(setting, '時点コピーの保存先');
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (error) {
      if (noCreate) return null;
      throw new Error('時点コピーの保存先のフォルダ（ID：' + id + '）を開けません。実行するアカウントが、そのフォルダに書き込めるか確かめてください' +
        '（' + setting.where + '）。');
    }
  }
  const props = PropertiesService.getDocumentProperties();
  const saved = props.getProperty(MTG_OPTIONS.copyFolderKey);
  if (saved) {
    try {
      const folder = DriveApp.getFolderById(saved);
      if (!folder.isTrashed()) return folder;
    } catch (_) {}
  }
  if (noCreate) return null;
  const folder = mtgFolder_(ss).createFolder(MTG_OPTIONS.copyFolderName);
  props.setProperty(MTG_OPTIONS.copyFolderKey, folder.getId());
  return folder;
}

/** 時点コピーのタイトルの（）：「（資料作成時・会議日 2026-10-23）」「（日次）」「（基準）」。kind は MTG_OPTIONS.copyKinds のキー */
function mtgCopyNote_(kind, meeting) {
  return '（' + MTG_OPTIONS.copyKinds[kind] + (meeting ? '・会議日 ' + meeting.replace(/\//g, '-') : '') + '）';
}

/**
 * 時点コピーのタイトル（今の名前・前の名前）を読む。{at: 保存した日時（分まで。ms）, kind: 'meeting' | 'daily' | 'baseline' | ''（分からない）,
 * meeting: 資料作成時の会議日（yyyy/MM/dd）}。時点コピーのタイトルでなければ null。前の版の「（会議日 2026-10-23）」は資料作成時
 */
function mtgParseCopyTitle_(title) {
  const o = MTG_OPTIONS;
  const text = String(title || '');
  const name = [o.copyName].concat(o.formerCopyNames || []).find(n => text.indexOf(n + '_') === 0);
  if (!name) return null;
  const m = /^(\d{4}-\d{2}-\d{2} \d{2}:\d{2})(?:（(.*)）)?$/.exec(text.slice(name.length + 1));
  if (!m) return null;
  const note = m[2] || '';
  const day = /会議日 (\d{4})-(\d{2})-(\d{2})/.exec(note);
  const kind = note.indexOf(o.copyKinds.daily) === 0 ? 'daily' : note.indexOf(o.copyKinds.baseline) === 0 ? 'baseline' : day ? 'meeting' : '';
  return {at: Utilities.parseDate(m[1], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm').getTime(), kind,
    meeting: kind === 'meeting' && day ? day[1] + '/' + day[2] + '/' + day[3] : ''};
}

/** 保存先のフォルダにある時点コピー [{id, title, at, kind, meeting}]（新しい順。ゴミ箱のもの・タイトルの形が違うものは除く。フォルダが無ければ空） */
function mtgListCopies_() {
  const folder = mtgCopyFolder_(SpreadsheetApp.getActiveSpreadsheet(), true);
  if (!folder) return [];
  const out = [];
  try {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const file = files.next();
      if (file.isTrashed()) continue;
      const parsed = mtgParseCopyTitle_(file.getName());
      if (parsed) out.push(Object.assign({id: file.getId(), title: file.getName()}, parsed));
    }
  } catch (error) {
    console.warn('時点コピーの保存先のフォルダを読めませんでした：' + (error && error.message || error));
  }
  return out.sort((a, b) => b.at - a.at);
}

/**
 * 日次の時点コピーを始める（エディタから実行する。管理者のアカウントで）。毎日 MTG_OPTIONS.dailyCopy.hour 時台に、スプレッドシート全体を時点コピー（日次）にコピーする
 * トリガーを1つだけ作る（もう一度実行しても増えない）。資料を作る画面の「比べる時点」に出る
 */
function installDailySheetCopy() {
  const o = MTG_OPTIONS.dailyCopy;
  const message = mtgInstallDailyCopyTrigger_() ?
    '日次の' + MTG_OPTIONS.copyName + 'を始めました。毎日 ' + o.hour + '時台に、スプレッドシート全体を「' + MTG_OPTIONS.copyName + '_日時（' + MTG_OPTIONS.copyKinds.daily + '）」にコピーします' +
      '（' + o.keepDays + '日より古い日次のコピーはゴミ箱に移します）。今すぐ1つ保存するときは saveDailySheetCopy() を実行してください。'
    : '日次の' + MTG_OPTIONS.copyName + 'はすでに始めています（毎日 ' + o.hour + '時台）。';
  console.log(message);
  const ui = mtgUiOrNull_();
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** 日次の時点コピーをやめる（エディタから実行する。トリガーを外す。保存したコピーは消さない） */
function uninstallDailySheetCopy() {
  const handler = MTG_OPTIONS.dailyCopy.handler;
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === handler);
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  const message = triggers.length ? '日次の' + MTG_OPTIONS.copyName + 'をやめました（保存したコピーは残しています）。'
    : '日次の' + MTG_OPTIONS.copyName + 'は始めていません。';
  console.log(message);
  const ui = mtgUiOrNull_();
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** 日次の時点コピーのトリガーを1つだけにする（無ければ作る。2つ以上あれば1つを残す）。作ったら true */
function mtgInstallDailyCopyTrigger_() {
  const o = MTG_OPTIONS.dailyCopy;
  const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === o.handler);
  existing.slice(1).forEach(t => ScriptApp.deleteTrigger(t));
  if (existing.length) return false;
  ScriptApp.newTrigger(o.handler).timeBased().everyDays(1).atHour(o.hour).inTimezone(Session.getScriptTimeZone()).create();
  return true;
}

/**
 * 日次の時点コピーを保存する（時間主導のトリガーから。エディタから実行してもよい）。スプレッドシート全体を「時点コピー_日時（日次）」にコピーし、
 * keepDays 日より古い日次のコピーをゴミ箱に移す（資料作成時・基準のコピーは消さない）。シートは読むだけなのでロックは取らない
 */
function saveDailySheetCopy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const copy = mtgSaveCopy_(ss, now, mtgCopyNote_('daily'));
  const limit = now.getTime() - MTG_OPTIONS.dailyCopy.keepDays * 86400000;
  let removed = 0;
  mtgListCopies_().filter(c => c.kind === 'daily' && c.at < limit).forEach(c => {
    try {
      DriveApp.getFileById(c.id).setTrashed(true);
      removed++;
    } catch (error) {
      console.warn('古い日次の' + MTG_OPTIONS.copyName + '（' + c.title + '）をゴミ箱に移せませんでした：' + (error && error.message || error));
    }
  });
  console.log('スプレッドシート全体を「' + copy.title + '」にコピーしました：' + copy.url + (removed ? '（古い日次のコピー ' + removed + '件をゴミ箱に移しました）' : ''));
  return copy;
}

/**
 * 前回の時点コピー（copy = {id, sheet}）の、そのシートの案件を読む。{table: mtgReadDeals_ の形, title, url}。
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
    const range = diffRuleRange_(sheet, a1);   // シートより広い範囲（列を消したあと）はシートの中まで
    if (!range) return;
    for (let c = range.getColumn(); c <= range.getLastColumn(); c++) {
      const header = String(headers[c - 1] || '');
      const key = credNormalize_(header);
      if (!key || diffNormalizeStampHeader_(header) === stamp || mtgIgnoredHeader_(header) || out.some(x => x.key === key)) continue;
      out.push({key, header});
    }
  });
  return out;
}

/* ---------------- 読む ---------------- */

/**
 * 記録のシート（変更履歴_差分・変更時点スナップショット）を下から読み、since 以降の行を返す（古い順）。
 * 記録は追記なので、サンプル（イベントID が SMP-）・ダミー（DMY-。addBusinessMtgDummyHistory()）以外で since より前の行に届いたら止める
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
  const find = label => (label ? headers.findIndex(h => credNormalize_(h) === credNormalize_(label)) : -1);
  const cols = {};
  const has = {};
  Object.keys(MTG_OPTIONS.columns).forEach(key => {
    cols[key] = mtgColumnIndex_(headers, key);
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
  const required = MTG_OPTIONS.required.map(label => ({label, col: mtgHeaderIndex_(headers, label)})).filter(r => r.col >= 0);

  const out = {headers, has, hasStamp: stampCol >= 0, statusOrder: [], deals: [], weightedFormula: false};
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return out;
  // 期待値の列が数式（想定売上 × ステータスの率など）なら、変更の記録に残らないので、前回の値はスナップショットから取る
  if (has.weighted) out.weightedFormula = !!String(sheet.getRange(headerRow + 1, cols.weighted + 1).getFormula() || '').trim();
  // ステータスのプルダウンの選択肢（一覧でも、定義のシートの範囲でも）
  out.statusOrder = mtgListOptions_(sheet.getRange(headerRow + 1, cols.status + 1).getDataValidation()) || [];
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
      planMonth: mtgPlanMonth_(value('startMonth')),
      activity: text('activity'),
      issue: text('issue'),
      owner: text('owner'),
      delivery: text('delivery'),
      nextAction: text('nextAction'),
      scenarios: mtgScenarioList_(text('scenario')),
      record: text('record'),
      updatedAt: stampCol >= 0 ? mtgTime_(values[n][stampCol]) : null,
      missing: required.filter(r => !String(line[r.col]).trim()).map(r => r.label),
      // 見出し（正規化）→ 表示の値（時点コピーと今の行を比べる用）
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
 *         changes, stalled, focus: 期待値の大きい進行中の案件, recent: focus の直近の動き {from, logFrom: 記録の始まり|null, moves: [[動き]]（focus の順）}
 *         （mtgRecentMoves_）, month: 今月の番号, overdue: 受注月を過ぎた進行中の案件, alerts: 入力内容のチェック,
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
  // 削除した行の記録（前回より後）。時点コピーと比べるときは、削除した時刻に使う
  const removedLines = history.removed.filter(line => line.at > since).map(line => {
    const r = Object.assign(mtgParseRemoved_(line.text, table.headers), {at: line.at});
    r.key = mtgDealKey_(r.customer, r.project);
    return r;
  });
  if (copy) {
    // 前回の時点コピーの行と今の行を、得意先＋案件名で突き合わせる（名前を書き換えた案件は、変更の記録で前の名前とつなげる）。
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
  const statusGroups = mtgStatusGroups_(table.statusOrder);
  const groupIndex = ch => statusGroups.findIndex(gr => gr.key === mtgStatusGroupOf_(ch));
  changes.sort((a, b) => (groupIndex(a) - groupIndex(b)) || ((a.category === 'updated' ? 1 : 0) - (b.category === 'updated' ? 1 : 0)) ||
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
  // 期待値上位案件の直近の動き（recentMonths か月前の同じ日から。変更の記録だけから。記録の始まりより前は分からない）
  const recentFrom = mtgMonthsAgo_(context.now, o.recentMonths);
  const recent = {from: recentFrom, logFrom: context.trackedSince == null ? null : context.trackedSince,
    moves: focus.map(d => mtgRecentMoves_(history.events.get(d.key) || [], recentFrom))};

  // 入力内容のチェックと、試作から解消した件数
  const alerts = mtgAlerts_(table, context, {slipsOf, changes, tracked});
  // 作れなかったところ（シートに列が無い・記録が無いなど。資料の型は変えずに、その場所に理由を書く）
  const gaps = mtgGaps_(ss, table, context, stallPending);
  // 時点コピーでは変わっているのに、変更の記録が無い案件（差分追跡が止まっていた間・2,000セルを超える貼り付けなど）
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
    scenarioColumn: !!table.has.scenario,
    stallPending,
    gaps: gaps.parts,
    alertGaps: gaps.alerts,
    bridge,
    stages: mtgStages_(deals, table.statusOrder),
    quarters: mtgQuarters_(active, context.now),
    plan: mtgPlanOf_(deals, name, context.now, table.has),
    changes,
    stalled,
    focus,
    recent,
    month,
    statusGroups,
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
 * amount・weighted・orderMonth・activity・issue・owner・nextAction・record（シートにその列が無い。キーは MTG_OPTIONS.columns）
 */
function mtgGaps_(ss, table, context, stallPending) {
  const o = MTG_OPTIONS;
  const tz = Session.getScriptTimeZone();
  const column = key => 'シートに「' + String(o.columns[key]).replace(/\s+/g, '') + '」の列が無い';
  const parts = {};
  const noLog = context.trackedSince == null;
  // 前回の時点コピーと比べるときは、変更の記録が無くても前回からの変化を出せる
  if (noLog && !context.copy) parts.changes = '「' + DIFF_OPTIONS.logSheet + '」シートに変更の記録が無い';
  ['amount', 'weighted', 'orderMonth', 'activity', 'issue', 'owner', 'nextAction', 'record'].forEach(key => {
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
  if (credEnabled_() && !ss.getSheetByName(CRED_OPTIONS.logSheet)) alerts['予定日超過'] = '「' + CRED_OPTIONS.logSheet + '」シートが無い';
  if (parts.orderMonth) alerts['受注月超過'] = parts.orderMonth;
  if (parts.record) alerts['資料なし'] = parts.record;
  if (!table.hasStamp && noLog) alerts['更新なし'] = 'シートに「最終更新日時」の列が無く、変更の記録も無い';
  if (parts.amount || parts.weighted) alerts['金額の不整合'] = parts.amount || parts.weighted;
  if (parts.orderMonth || !table.has.startMonth) alerts['月の不整合'] = parts.orderMonth || column('startMonth');
  if (!table.statusOrder.length) alerts['選択肢外のステータス'] = '「' + o.columns.status + '」の列にプルダウンが無い';
  // デリバリー担当の列を使わない設定（columns.delivery が空）なら、チェックしないので「データなし」にも出さない
  if (o.columns.delivery) {
    if (!table.has.delivery) alerts['デリバリー担当なし'] = column('delivery');
    else if (noLog) alerts['デリバリー担当なし'] = '「' + DIFF_OPTIONS.logSheet + '」シートに変更の記録が無い';
  }
  return {parts, alerts};
}

/**
 * 作れなかったところの一覧 ['停滞：基準（10/16）から…', …]（作り終えたときの画面とログに出す）。
 * 同じ理由の部分はまとめる（「受注月超過・入力内容のチェック「受注月超過」：シートに「受注月」の列が無い」）
 */
function mtgGapList_(s) {
  const names = {changes: '前回からの変化', stalled: '停滞', amount: '想定売上', weighted: '期待値', orderMonth: '受注月・受注月超過',
    activity: '活動状況', issue: '活動における課題', owner: '責任者', nextAction: '次回アクション', record: '商談記録未登録'};
  const byReason = new Map();
  const add = (label, reason) => {
    if (!byReason.has(reason)) byReason.set(reason, []);
    byReason.get(reason).push(label);
  };
  Object.keys(s.gaps || {}).forEach(key => add(names[key] || key, s.gaps[key]));
  mtgAlertTypes_().filter(type => (s.alertGaps || {})[type]).forEach(type => add('チェック「' + (MTG_ALERT_LABELS[type] || type) + '」', s.alertGaps[type]));
  return Array.from(byReason.keys()).map(reason => byReason.get(reason).join('・') + '：' + reason);
}

/** 正規化した見出し（変更の記録の列名と比べる） */
/** columns の key の列（0 から。無ければ -1）。columns の見出しが無ければ columnAliases の見出しを順に探す */
function mtgColumnIndex_(headers, key) {
  const names = [MTG_OPTIONS.columns[key]].concat((MTG_OPTIONS.columnAliases || {})[key] || []).filter(Boolean).map(credNormalize_);
  for (const name of names) {
    const i = headers.findIndex(h => credNormalize_(h) === name);
    if (i >= 0) return i;
  }
  return -1;
}

/** 見出し label の列（0 から。無ければ -1）。columns の見出しなら columnAliases の見出しも探す（必須項目など） */
function mtgHeaderIndex_(headers, label) {
  const cols = MTG_OPTIONS.columns;
  const key = Object.keys(cols).find(k => cols[k] && credNormalize_(cols[k]) === credNormalize_(label));
  return key ? mtgColumnIndex_(headers, key) : headers.findIndex(h => credNormalize_(h) === credNormalize_(label));
}

/** 変更の記録などの列名を、資料で使う見出し（mtgLabels_ の値）にそろえる（columnAliases の見出しは columns の見出しに）。ほかは正規化しただけ */
function mtgCanonicalLabel_(column) {
  const cols = MTG_OPTIONS.columns;
  const key = Object.keys(cols).find(k => cols[k] && mtgColumnIndex_([column], k) === 0);
  return credNormalize_(key ? cols[key] : column);
}

/** 資料に含めない列か（MTG_OPTIONS.ignoreHeaders。FY別の売上の列など） */
function mtgIgnoredHeader_(header) {
  const text = credNormalize_(header);
  return (MTG_OPTIONS.ignoreHeaders || []).some(pattern => new RegExp(pattern).test(text));
}

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
    const label = mtgCanonicalLabel_(e.column);
    if (isKey(label)) keyBefore.set(e.eventId + '#' + mtgCellRow_(e.cell) + '#' + label, empty(e.before));
  });
  const alias = new Map();
  const items = [];
  const removed = [];
  let unmatched = 0;
  let dummy = 0;   // ダミーの記録（DMY-。addBusinessMtgDummyHistory()）の数。あれば資料に【ダミーデータを含む】と出す
  context.log.forEach(e => {
    if (resolve(e.sheet) !== name || e.at == null || /^(SMP|CRD)-/.test(e.eventId)) return;
    if (mtgIgnoredHeader_(e.column)) return;   // 資料に含めない列（FY別の売上など）の変更は、変化にも更新項目にも出さない
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
    const label = mtgCanonicalLabel_(e.column);
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
  // 時点コピーと比べるとき（rules.copy）：prev はコピーの行（無ければ null：前回は無かった案件）。値が同じなら変化なし
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
 * 入力内容のチェックの種類（MTG_ALERT_TYPES の順）。クレデンシャル・オファリング履歴を使っていない間は「予定日超過」を、
 * デリバリー担当の列を使わない設定（MTG_OPTIONS.columns.delivery が空）なら「デリバリー担当なし」を除く
 */
function mtgAlertTypes_() {
  return MTG_ALERT_TYPES.filter(type => (type !== '予定日超過' || credEnabled_()) && (type !== 'デリバリー担当なし' || !!MTG_OPTIONS.columns.delivery));
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
function mtgSlides_(summary, start, options) {
  const offset = start || 0;   // 前に入るページの数（統合版。ページ番号と扉の目次のページを通しにする）
  // 統合版は Appendix（扉と ①〜③）を入れない（options.appendix: false。2026/10/19 から。全件の一覧は各サブインダストリーの資料にある）
  const withAppendix = !(options && options.appendix === false);
  const cover = mtgCoverSlide_(summary);
  const main = [mtgSummarySlide_(summary)].concat(MTG_OPTIONS.plan.enabled ? [mtgPlanSlide_(summary.plan, summary.now)] : [],
    [mtgBridgeSlide_(summary), mtgMovesSlide_(summary, withAppendix)], mtgFocusSlide_(summary), mtgRecentSlide_(summary));
  if (!withAppendix) return mtgNumberSlides_(mtgSheetFrame_(summary), [cover].concat(main), offset);
  // Appendix はいつも ①〜③ を出す（中身が無ければ、表の見出しの行の下に「データなし」と理由・「該当なし」）。色は appendixColors（グレー）
  const appendix = mtgWithColors_(MTG_OPTIONS.appendixColors, () => [].concat(mtgAppendixChanges_(summary), mtgAppendixStalled_(summary),
    mtgAppendixOwners_(summary), mtgAppendixAlerts_(summary)));
  return mtgNumberSlides_(mtgSheetFrame_(summary), [cover].concat(main, [mtgDividerSlide_(appendix, offset + 1 + main.length + 1 + 1)], appendix), offset);
}

/**
 * 統合版のスライドの並び：表紙（全サブインダストリー）・目次（mtgContentsSlide_）・全サブインダストリーの状況（mtgOverviewSlide_）のあとに、各サブインダストリーの資料
 * （mtgSlides_。表紙から本編まで。Appendix は入れない。summaries の順）をそのまま続ける。ページ番号は資料全体の通し番号にする
 */
function mtgCombinedSlides_(summaries, meeting, final) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const name = o.combined.name;
  const dummy = summaries.some(s => s.dummy);
  const planOn = !!o.plan.enabled;
  let page = planOn ? 5 : 3;   // 前に表紙・目次・全サブインダストリーの状況（・計画に対する積み上げの全体とサブインダストリー別）
  const sections = summaries.map(s => {
    const slides = mtgSlides_(s, page, {appendix: false});
    const section = {summary: s, start: page + 1, slides};
    page += slides.length;
    return section;
  });
  const line = (y, text) => mtgBox_(S.x, y, S.w, 16, mtgWrap_(text, S.w, 9, 1)[0], {size: 9, color: c.white, align: 'CENTER'});
  const now = summaries.length ? summaries[0].now : new Date();
  const cover = {
    title: MTG_TITLES.cover,
    cover: true,
    combined: true,   // 統合版の表紙（緑。MTG_OPTIONS.combined.coverTone・templatePages.combined）
    template: {title: name, subtitle: o.meetingName + '　事前資料' + (final ? '' : '【試作】'), date: mtgDateLabel_(meeting), name: ''},
    elements: [
      line(350, '収録：' + summaries.map(s => s.name).join('・') + '（' + summaries.length + 'サブインダストリー）'),
      line(365, '出典：各サブインダストリーのシートおよび変更履歴（' + (dummy ? 'ダミーデータを含む。' : '') +
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/M/d HH:mm') + ' 時点）')
    ]
  };
  const frame = {source: (final ? '' : '【試作】') + (dummy ? '【ダミーデータを含む】' : '') + name + '（' + mtgShortDate_(meeting) + ' 会議資料）　金額の単位：百万円'};
  const plan = planOn ? [mtgPlanSlide_(mtgMergePlans_(summaries.map(s => s.plan)), now, name), mtgPlanBySubSlide_(summaries, now)] : [];
  const first = mtgNumberSlides_(frame, [cover, mtgContentsSlide_(sections), mtgOverviewSlide_(sections)].concat(plan));
  return first.concat.apply(first, sections.map(x => x.slides));
}

/**
 * 統合版の2枚目「目次」：全サブインダストリーの状況と、各サブインダストリーの資料（表紙）の、この資料の中のページ。名前とページ番号は、そのページへのリンク
 * （部品の link にスライドの番号（0 から）を入れ、mtgExportDeck_ がすべてのスライドを作ったあとにリンクにする）。9件以上なら2列。
 * ブラウザで開いた Google ドライブのプレビュー画面は PDF の中のページへのリンクに移らないので、リード文の補足にそう書く
 */
function mtgContentsSlide_(sections) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const entries = [{label: MTG_TITLES.overview, page: 3}].concat(MTG_OPTIONS.plan.enabled ? [{label: MTG_TITLES.plan, page: 4}] : [],
    sections.map((x, k) => ({label: x.summary.name, page: x.start, n: k + 1})));
  const cols = entries.length > 8 ? 2 : 1;
  const per = Math.ceil(entries.length / cols);
  const gap = 16;
  const colW = (S.w - gap * (cols - 1)) / cols;
  const rowH = Math.min(34, Math.floor((S.bottom - S.top) / per));
  const el = [];
  entries.forEach((entry, i) => {
    const x = S.x + Math.floor(i / per) * (colW + gap);
    const y = S.top + (i % per) * rowH;
    const h = rowH - 6;
    const link = entry.page - 1;
    el.push(mtgBox_(x, y, colW - 56, h, mtgWrap_((entry.n ? entry.n + '.　' : '') + entry.label, colW - 56, 12, 1)[0],
      {size: 12, bold: true, color: c.ink, fill: c.greenSoft, shape: 'RECTANGLE', link, role: 'toc'}));
    el.push(mtgBox_(x + colW - 56, y, 56, h, String(entry.page), {size: 12, bold: true, color: c.greenDark, fill: c.greenSoft, shape: 'RECTANGLE',
      align: 'END', link, role: 'toc'}));
  });
  const lead = '名前かページ番号を押すと、そのページに移ります。\n' +
    '※ ブラウザの Google ドライブのプレビュー画面ではリンクは動きません（ダウンロードして Acrobat などで開いてください）。';
  return {title: MTG_TITLES.contents, lead, elements: el};
}

/**
 * 統合版の3枚目「全サブインダストリーの状況」：サブインダストリーごとの進行中案件・期待値合計・前回比・変化のある案件・停滞・確認事項と、
 * その資料の表紙のページ（この資料の中のページ。各資料のサマリーと同じ数え方。作れない値は灰色の「データなし」）。いちばん下に合計（作れたシートの分だけ）。
 * sections = [{summary, start: 始まりのページ}]
 */
function mtgOverviewSlide_(sections) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const S = MTG_SLIDE;
  const G = o.gapText;
  const columns = mtgColumns_([142, 70, 84, 84, 92, 68, 70, 50], 0);
  // 増減の列の見出しは、いちばん多い比べ方（前回比・10/17比（基準）・直近14日間の増減）。ほかの比べ方のシートは、セルの1行目に何との比かを書く
  const versus = s => mtgVersus_(s).trim().replace(/で$/, '');   // 前回比・10/17比（基準）・直近14日間
  const counts = {};
  sections.forEach(x => { counts[versus(x.summary)] = (counts[versus(x.summary)] || 0) + 1; });
  const common = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || '前回比';
  const vs = s => (versus(s) === common ? '' : versus(s) + '\n');
  const labels = ['サブインダストリー', '進行中案件', '期待値合計', /比$/.test(common) ? common : common + 'の増減', '変化のある案件', '停滞', '確認事項', 'ページ'];
  const total = {active: 0, weighted: 0, delta: 0, changed: 0, stalled: 0, alerts: 0, lack: {}};
  const lines = sections.map(({summary: s, start}) => {
    const k = s.kpi;
    const lack = s.gaps || {};
    const b = s.bridge;
    total.active += k.active;
    total.alerts += k.alerts;
    if (lack.weighted) total.lack.weighted = true;
    else total.weighted += b.end;
    if (lack.weighted || lack.changes) total.lack.delta = true;
    else total.delta += b.end - b.start;
    if (lack.changes) total.lack.changed = true;
    else total.changed += k.changed;
    if (lack.stalled) total.lack.stalled = true;
    else total.stalled += k.stalled;
    return {name: s.name, page: String(start), values: [
      k.active + '件',
      lack.weighted ? null : mtgAmount_(b.end),
      lack.weighted || lack.changes ? null : vs(s) + mtgSigned_(b.end - b.start),
      lack.changes ? null : k.changed + '件',
      lack.stalled ? null : k.stalled + '件',
      k.alerts + '項目'
    ]};
  });
  // 合計：作れたシートの分だけ（作れないシートがあれば ※。どのシートも作れなければ「データなし」）
  const of = (i, text, key) => (lines.every(line => line.values[i] == null) ? null : text + (total.lack[key] ? '※' : ''));
  const sum = {name: '合計', page: '', total: true, values: [total.active + '件', of(1, mtgAmount_(total.weighted), 'weighted'),
    of(2, mtgSigned_(total.delta), 'delta'), of(3, total.changed + '件', 'changed'), of(4, total.stalled + '件', 'stalled'), total.alerts + '項目']};
  // リード文は1行（表を大きく。11サブインダストリーと合計が1ページに入るように）。合計に ※ があるときは、出典の行の後ろに意味を書く
  const lead = sections.length + 'サブインダストリーの進行中案件 ' + total.active + '件、期待値合計 ' + mtgAmount_(total.weighted) + '百万円' +
    (total.lack.weighted ? '（期待値の列が無いシートを除く）' : '') + '。';
  const marked = sum.values.some(v => /※$/.test(v || ''));
  const top = Math.max(S.summaryTop, Math.max.apply(null, mtgLead_(lead).map(e => e.y + e.h)) + 4);
  // 表の文字は、ページに入る大きさ（9pt から 0.5pt ずつ 7pt まで）
  const build = size => {
    const cell = (text, extra) => Object.assign({text, size, lines: 2, align: 'CENTER', fill: c.white, color: c.body}, extra || {});
    const rows = lines.concat([sum]).map(line => mtgRow_(columns, [
      {text: line.name, size, bold: true, color: c.ink, fill: c.greenSoft}
    ].concat(line.values.map((v, i) => (v == null ? mtgGapCell_(cell('')) : cell(v, i === 1 || i === 2 ? {align: 'END'} : {}))))
      .concat([cell(line.page, {bold: true, color: c.greenDark})])
      .map(x => (line.total && x.fill !== c.gap ? Object.assign({}, x, {bold: true, fill: c.greenSoft}) : x)), 0));
    return [mtgHeadRow_(columns, labels, size)].concat(rows);
  };
  let size = S.bodySize;
  let rows = build(size);
  const height = list => list.reduce((h, r) => h + r.h, 0);
  while (top + height(rows) > S.bottom && size - 0.5 >= S.minFitSize) {
    size = Math.round((size - 0.5) * 10) / 10;
    rows = build(size);
  }
  return Object.assign({title: MTG_TITLES.overview, lead, elements: [{type: 'table', x: S.x, y: top, columns, rows}]},
    marked ? {note: '※ は「' + G + '」のシートを除いた合計'} : {});
}

/**
 * fn を、MTG_OPTIONS.colors の一部を overrides に置き換えた色で作る（Appendix をグレーにする。終わったら元に戻す）。
 * 区分などのラベル（mtgPills_）は意味の色なので、元の色（MTG_OPTIONS.baseColors）で描く
 */
function mtgWithColors_(overrides, fn) {
  if (!overrides) return fn();
  const o = MTG_OPTIONS;
  const base = o.colors;
  o.baseColors = base;
  o.colors = Object.assign({}, base, overrides);
  try {
    return fn();
  } finally {
    o.colors = base;
    delete o.baseColors;
  }
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
      line(350, '出典：「' + s.name + '」シート' + (s.copy ? '・' + mtgCopyWord_(s) : '') + 'および変更履歴（' + (s.dummy ? 'ダミーデータを含む。' : '') +
        Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/M/d HH:mm') + ' 時点）'),
      line(365, '対象期間：' + mtgPeriodText_(s))
    ]
  };
}

/**
 * 扉（Appendix。ひな形の2枚目（グレー）。無ければ同じ絵柄のグレーの背景に描く）：見出しを「Appendix」にし（ほかの決まった文字は空）、その下に Appendix の見出し・始まりのページの一覧
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

/** スライドに、どのスライドにも付ける部品（mtgFrame_）とページ番号（offset ＋ 1 から）を付ける（表紙・扉の印と、ひな形の置き換える文字は残す） */
function mtgNumberSlides_(frame, slides, offset) {
  return slides.map((slide, i) => {
    const out = {title: slide.title, elements: mtgFrame_(frame, slide, (offset || 0) + i + 1).concat(slide.elements)};
    ['cover', 'divider', 'combined', 'service', 'template', 'empty'].forEach(key => { if (slide[key]) out[key] = slide[key]; });
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

/**
 * 比べる時点の短い名前：「前回」、選んだ時点なら「10/9」、基準なら基準を記録した月日「10/17」（「基準比」では前回比と見分けにくいため。2026/10/19 から。
 * 年月（26/10）だったのを、ほかの日付の書き方（M/d）にそろえた）
 */
function mtgFromWord_(s) {
  const prev = s.previous;
  if (prev && prev.baseline) return Utilities.formatDate(new Date(prev.cutoff), Session.getScriptTimeZone(), 'M/d');
  return prev && prev.chosen ? mtgShortDate_(prev.date) : '前回';
}

/** 「前回の時点コピー」「10/8 18:20 の時点コピー」（選んだ時点）「基準の時点コピー」（比べたコピーの呼び方） */
function mtgCopyWord_(s) {
  const prev = s.previous;
  const name = MTG_OPTIONS.copyName;
  if (!prev || !prev.chosen) return '前回の' + name;
  return prev.baseline ? '基準の' + name : Utilities.formatDate(new Date(prev.cutoff), Session.getScriptTimeZone(), 'M/d HH:mm') + ' の' + name;
}

/** 「前回以降」「10/17以降」（基準）「10/9以降」（選んだ時点）「直近14日間」（数字の見出し・リード文の期間） */
function mtgPeriodShort_(s) {
  if (s.first) return '直近' + MTG_OPTIONS.firstDays + '日間';
  return mtgFromWord_(s) + '以降';
}

/** 「うち前回以降の更新なし」「うち10/17以降の更新なし」「うち直近14日間の更新なし」（停滞のうち、対象期間に変化の無い案件） */
function mtgQuietLabel_(s) {
  return 'うち' + mtgPeriodShort_(s) + 'の更新なし';
}

/** 「前回比 」「10/17比 」（基準）「10/9比 」（選んだ時点）「直近14日間で 」（前回比の書き出し） */
function mtgVersus_(s) {
  if (s.first) return mtgPeriodShort_(s) + 'で ';
  return mtgFromWord_(s) + '比 ';
}

/** 「（ほかに、その他の項目のみの更新 N件）」（その他更新は「変化のある案件」に数えないので、別に書く。無ければ空） */
function mtgUpdatedNote_(s) {
  return s.kpi.updated ? '（ほかに、その他の項目のみの更新 ' + s.kpi.updated + '件）' : '';
}

/** 「前回から変化のある案件」「10/17以降に変化のある案件」（基準）「10/9以降に…」（選んだ時点）「直近14日間に変化のある案件」 */
function mtgChangedLabel_(s) {
  if (s.first) return '直近' + MTG_OPTIONS.firstDays + '日間に変化のある案件';
  return s.previous && (s.previous.chosen || s.previous.baseline) ? mtgFromWord_(s) + '以降に変化のある案件' : '前回から変化のある案件';
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
  // 1列目は、今回のステータスのまとまりの名前（「C 提案中・選定対応中・稟議対応中」）が1行に入る幅
  const columns = [176, 52, 58, S.w - 286];
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
  // 入りきらなければ（リード文が長いときなど）、件数の文字 → 箱の下 → 表の間 → 表の見出し → 表の文字 → … → 数字の箱 の順に、決めた下限まで小さくする
  const lay = {th: 34, after: 6, headH: 16, gap: 5, count: 10, body: S.bodySize};
  const build = () => {
    const head = mtgHeadRow_(columns, ['今回のステータス', '件数', '想定売上', '主な案件（想定売上順。（）は変わったこと）'], lay.body);
    const rows = all.map(line => rowOf(line, 1, lay));
    return {head, rows, fixed: top + lay.th + lay.after + 2 * lay.headH + lay.gap + head.h};
  };
  let plan = build();
  // （動きの行はステータスのまとまりの6行。リード文が2行でも入るよう、表の文字は 7.5pt まで）
  // 数字の箱は最後に詰める（大きな数字が箱からはみ出さないように 30pt まで）
  [['count', 9], ['after', 4], ['gap', 3], ['headH', 14], ['body', 8], ['count', 8], ['headH', 13], ['body', 7.5], ['count', 7.5], ['after', 2], ['gap', 2],
    ['headH', 12], ['th', 32], ['body', 7], ['count', 7], ['th', 30]].forEach(([key, min]) => {
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
 * moves：前回以降の動き。変化のある案件（その他更新を除く）を、今のステータスのまとまり（mtgStatusGroups_。確度 A → 創注・失注・保留・削除）の行に
 * 1回だけ数える（区分は使わない。2026/10/19 から）。名前の後ろの（）は何が変わったか（ステータスが変われば「前回 C-提案中」、新規、受注月・期待値の変更、
 * 活動更新。失注・保留・削除の行は今のステータスか「削除」も）。
 * states：現在の状態（受注月超過・停滞・商談記録未登録）。停滞は、前回以降に更新の無い案件を「うち前回以降の更新なし N件」として名前を並べる
 * （初回は「直近N日間」、基準の後は「基準以降」。mtgQuietLabel_）。
 * どの行も、どのシートでも出す（資料の型を変えない）。作れない行（変更の記録・列が無い、停滞をまだ数えられない）は gap に理由を入れる
 */
function mtgBreakdown_(s) {
  const o = MTG_OPTIONS;
  const byAmount = list => list.slice().sort((a, b) => (b.amount || 0) - (a.amount || 0));
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
  // 動き：今のステータスのまとまりごと（確度 A → 創注・失注・保留・削除。ステータス未入力・その他は案件があるときだけ）。その他更新は数えない。
  // 名前の後ろには、何が変わったか（前回のステータス・新規・受注月・期待値・活動更新）
  const short = st => (st ? String(st).normalize('NFKC').replace(/\(.*$/, '').trim() : '未入力');
  const what = ch => {
    // 失注・保留・削除の行は、ステータスがまちまちなので今のステータス（削除）も書く
    const closed = mtgStatusGroupOf_(ch) === 'closed';
    const now = ch.removed || ch.category === 'removed' ? '削除' : short(ch.statusTo);
    if (ch.category === 'new') return closed ? now + '・新規' : '新規';
    if (closed) return now + ((ch.statusFrom || '') !== (ch.statusTo || '') && ch.statusFrom ? '・前回 ' + short(ch.statusFrom) : '');
    if ((ch.statusFrom || '') !== (ch.statusTo || '')) return '前回 ' + short(ch.statusFrom);
    if (ch.category === 'slipped' || ch.category === 'pulled') return '受注月 ' + month(ch);
    if (ch.category === 'up' || ch.category === 'down') return '期待値 ' + value(ch);
    if (ch.category === 'activity') return '活動更新';
    return '';
  };
  const changed = s.changes.filter(ch => ch.category !== 'updated');
  const moves = (s.statusGroups || mtgStatusGroups_([])).map(gr => ({label: gr.label, key: gr.key,
    items: byAmount(changed.filter(ch => mtgStatusGroupOf_(ch) === gr.key)), note: what}));
  // 確度の無いステータス（未入力・プルダウンに無いものなど）の案件も最後の行に入れ、そのときは名前に「その他」を足す
  const last = moves[moves.length - 1];
  if (last.items.some(ch => !(ch.removed || ch.category === 'removed') && ['lost', 'hold'].indexOf(mtgStatusKind_(ch.statusTo || '')) < 0)) {
    last.label += '・その他';
  }
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
  const text = n => head + items.slice(0, n).map(d => mtgDealLabel_(d) + (note && note(d) ? '（' + note(d) + '）' : '')).join('、') +
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

/* ---------------- 計画に対する積み上げ ---------------- */

/**
 * 計画に対する積み上げの月（売上開始月）。日付か「yyyy/MM」の文字（2000年以降）のときだけ月とみなす（33シナリオ ダッシュボード・パイプライン管理シートと同じ。
 * 「26/09」「2026年度後半期」などは時期未定）。{y, m} か null
 */
function mtgPlanMonth_(value) {
  if (value && typeof value.getTime === 'function') return mtgMonth_(value);
  const m = /^(\d{4})\/(\d{2})$/.exec(String(value == null ? '' : value).trim());
  return m && Number(m[1]) >= 2000 && Number(m[2]) >= 1 && Number(m[2]) <= 12 ? {y: Number(m[1]), m: Number(m[2])} : null;
}

/** 案件の確度（ステータスの「-」の前。MTG_OPTIONS.plan.ranks のどれでもなければ otherLabel） */
function mtgPlanRank_(status) {
  const o = MTG_OPTIONS.plan;
  const head = String(status || '').normalize('NFKC').split('-')[0].trim();
  return o.ranks.some(r => r.key === head) ? head : o.otherLabel;
}

/** シート（name）の計画（百万円。MTG_OPTIONS.plan.plans にあればその値、無ければ amount） */
function mtgPlanAmount_(name) {
  const o = MTG_OPTIONS.plan;
  return o.plans && o.plans[name] != null ? Number(o.plans[name]) : o.amount;
}

/**
 * 計画に対する積み上げの中身。今年度（MTG_OPTIONS.fiscalStartMonth から12か月）の売上開始月ごとの期待値を、確度（mtgPlanRank_）ごとに足す。
 * 進行中・受注の案件が対象。積み上げに入らない案件（時期未定・今年度より前・来年度以降・失注・保留）は、それぞれの件数と期待値（確度別も）を残す
 * （資料では積み上げの右に参考として出す）。期待値が空の案件は 0 として数え、件数を残す。has はシートの列があるか（期待値・売上開始月の列が無ければ gap に理由）。
 * 戻り値 {plan, start: 今年度の最初の月の番号, months: [{y, m, values: {確度: 期待値}, total, count}], total, byRank: {確度: 期待値},
 *  undated: 売上開始月が空・読めない, before: 今年度より前, after: 来年度以降, excluded: 失注・保留（どれも {count, value, values: {確度: 期待値}}）,
 *  missing: 期待値が空の件数, gap}
 */
function mtgPlanOf_(deals, name, now, has) {
  const o = MTG_OPTIONS.plan;
  const start = mtgFiscalStart_(now);
  const columns = MTG_OPTIONS.columns;
  const lack = ['weighted', 'startMonth'].filter(key => has && !has[key]).map(key => '「' + String(columns[key]).replace(/\s+/g, '') + '」');
  const bucket = () => ({count: 0, value: 0, values: {}});
  const out = {plan: mtgPlanAmount_(name), start, months: [], total: 0, byRank: {}, undated: bucket(), before: bucket(), after: bucket(),
    excluded: bucket(), missing: 0, gap: lack.length ? 'シートに' + lack.join('・') + 'の列が無い' : ''};
  for (let i = 0; i < 12; i++) {
    const mo = mtgMonthOf_(start + i);
    out.months.push({y: mo.y, m: mo.m, values: {}, total: 0, count: 0});
  }
  o.ranks.map(r => r.key).concat([o.otherLabel]).forEach(k => { out.byRank[k] = 0; });
  deals.forEach(d => {
    const v = d.weighted || 0;
    const rank = mtgPlanRank_(d.status);
    const i = d.planMonth ? mtgMonthIndex_(d.planMonth) - start : null;
    const lost = d.kind === 'lost' || d.kind === 'hold';
    if (!lost && d.weighted == null) out.missing++;
    const aside = lost ? out.excluded : i == null ? out.undated : i < 0 ? out.before : i >= 12 ? out.after : null;
    if (aside) {
      aside.count++;
      aside.value += v;
      aside.values[rank] = (aside.values[rank] || 0) + v;
      return;
    }
    const mo = out.months[i];
    mo.values[rank] = (mo.values[rank] || 0) + v;
    mo.total += v;
    mo.count++;
    out.byRank[rank] += v;
    out.total += v;
  });
  return out;
}

/** 統合版の計画に対する積み上げ：各サブインダストリーの mtgPlanOf_ の戻り値を足したもの（計画も合計。どれかに gap があれば最初の gap） */
function mtgMergePlans_(plans) {
  const out = JSON.parse(JSON.stringify(plans[0] || mtgPlanOf_([], '', new Date(), null)));
  out.plan = 0;
  out.total = 0;
  out.missing = 0;
  const asides = ['undated', 'before', 'after', 'excluded'];
  asides.forEach(k => { out[k] = {count: 0, value: 0, values: {}}; });
  Object.keys(out.byRank).forEach(k => { out.byRank[k] = 0; });
  out.months.forEach(mo => { mo.values = {}; mo.total = 0; mo.count = 0; });
  plans.forEach(p => {
    out.plan += p.plan;
    out.total += p.total;
    out.missing += p.missing;
    asides.forEach(k => {
      out[k].count += p[k].count;
      out[k].value += p[k].value;
      Object.keys(p[k].values).forEach(r => { out[k].values[r] = (out[k].values[r] || 0) + p[k].values[r]; });
    });
    Object.keys(p.byRank).forEach(k => { out.byRank[k] = (out.byRank[k] || 0) + p.byRank[k]; });
    p.months.forEach((mo, i) => {
      const to = out.months[i];
      Object.keys(mo.values).forEach(k => { to.values[k] = (to.values[k] || 0) + mo.values[k]; });
      to.total += mo.total;
      to.count += mo.count;
    });
  });
  out.gap = (plans.find(p => p.gap) || {}).gap || '';
  return out;
}

/**
 * 計画に対する積み上げ（本編2枚目。統合版は全サブインダストリーの状況の次）：今年度の売上開始月ごとの期待値を左から積み上げる滝の図。
 * 月の段は確度（濃い緑 A → 淡い緑 創注・灰色 未設定）で塗り分け、最後に 合計（確度別）・計画（白地に枠）・差（計画に届かない分 ▲／超えた分 +）。
 * 計画の高さに点線の横線を引く（どの月で計画に届くかが分かる）。右に区切って、積み上げに含まない案件（時期未定・今年度より前・来年度以降・失注・保留）を
 * 参考として 0 からの確度別の棒で出す（期待値と件数）。p は mtgPlanOf_（統合版は mtgMergePlans_）の戻り値。name は統合版のときだけ（見出しに付ける）
 */
function mtgPlanSlide_(p, now, name) {
  const S = MTG_SLIDE;
  const c = MTG_OPTIONS.colors;
  const o = MTG_OPTIONS.plan;
  const title = MTG_TITLES.plan + (name ? '（' + name + '）' : '');
  const first = mtgMonthOf_(p.start);
  const last = mtgMonthOf_(p.start + 11);
  const period = String(first.y).slice(2) + '/' + first.m + '〜' + String(last.y).slice(2) + '/' + last.m;
  const heading = '今年度（' + period + '）の売上開始月別 期待値の積み上げ（百万円）';
  if (p.gap) {
    return {title, lead: '計画に対する積み上げは' + MTG_OPTIONS.gapText + '（' + p.gap + '）。', elements: [
      mtgBox_(S.x, S.top - 2, 420, 20, heading, {size: 10.5, bold: true, color: c.ink}), mtgEmptyBox_(S.x, S.top + 26, 360, 64, 'gap', p.gap)]};
  }
  const ranks = o.ranks.concat([{key: o.otherLabel, color: o.otherColor}]);
  const diff = p.total - p.plan;
  const rate = p.plan > 0 ? Math.round(p.total / p.plan * 100) : null;
  const peak = p.months.slice().sort((a, b) => b.total - a.total)[0];
  const sure = (p.byRank.A || 0) + (p.byRank.B || 0);
  const share = v => (p.total > 0 ? Math.round(v / p.total * 100) + '%' : '—');
  // 積み上げに入らない案件（右に参考として出す）：時期未定（売上開始月が空・読めない）・今年度より前（〜26/3）・来年度以降（27/4〜）・失注・保留
  const prevEnd = mtgMonthOf_(p.start - 1);
  const asides = [
    {key: 'undated', label: '時期未定'},
    {key: 'before', label: '〜' + String(prevEnd.y).slice(2) + '/' + prevEnd.m},
    {key: 'after', label: String(last.y + (last.m === 12 ? 1 : 0)).slice(2) + '/' + (last.m % 12 + 1) + '〜'},
    {key: 'excluded', label: '失注・保留'}
  ];
  const asideCount = asides.reduce((sum, a) => sum + p[a.key].count, 0);
  const asideValue = asides.reduce((sum, a) => sum + p[a.key].value, 0);
  const lead = '今年度（' + period + '）の期待値の積み上げは ' + mtgAmount_(p.total) + '百万円。計画 ' + mtgAmount_(p.plan) + '百万円に対して ' +
    mtgSigned_(diff) + (rate == null ? '' : '（達成率 ' + rate + '%）') + '。\n' +
    (peak && peak.total > 0 ? '最大は' + peak.m + '月の ' + mtgAmount_(peak.total) + '（積み上げの ' + share(peak.total) + '）、確度A・B は ' +
      mtgAmount_(sure) + '（' + share(sure) + '）。' : '今年度に売上開始月がある案件の期待値はなし。') +
    (asideCount ? '積み上げに含まない ' + asideCount + '件（' + mtgAmount_(asideValue) + '）は右に表示。' : '') +
    (p.missing ? '期待値が空の ' + p.missing + '件は 0 として数えた。' : '');

  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 420, 20, heading, {size: 10.5, bold: true, color: c.ink})];
  // 凡例：確度の色（右端を図の右端にそろえる）と、計画の枠
  const legend = ranks.map(r => ({label: r.key, color: r.color})).concat([{label: '計画', color: c.white, line: c.dark}]);
  legend.forEach(item => { item.bw = mtgTextWidth_(item.label, 8.5) + 16; });
  let lx = x + w - legend.reduce((sum, item) => sum + 6 + item.bw, 0) - (legend.length - 1) * (8 - 7.2 - 1.6);
  legend.forEach(item => {
    el.push({type: 'box', x: lx, y: y + 5, w: 7, h: 7, fill: item.color, line: item.line, lineWeight: item.line ? 1 : undefined, shape: 'RECTANGLE', role: 'legend'});
    el.push(mtgBox_(lx + 6, y - 1, item.bw, 18, item.label, {size: 8.5, color: c.sub, role: 'legend'}));
    lx += 6 + item.bw + 8 - 7.2 - 1.6;
  });
  // 棒：12か月（前の月までの積み上げから浮かせる）・合計・計画・差（15本。幅 unit）。少しあけて、積み上げに含まない案件（参考。0 から。確度別。
  // 名前が長いので幅は unit × 1.3）
  const top = y + 40;
  const axis = S.bottom - 30;
  const unit = w / (15 + 0.5 + asides.length * 1.3);
  const lefts = [];
  const widths = [];
  for (let i = 0; i < 15 + asides.length; i++) {
    lefts.push(i < 15 ? x + unit * i : x + unit * (15.5 + (i - 15) * 1.3));
    widths.push(i < 15 ? unit : unit * 1.3);
  }
  const bw = Math.min(30, unit * 0.66);
  const bx = i => lefts[i] + (widths[i] - bw) / 2;
  const max = Math.max.apply(null, [p.total, p.plan, 1].concat(asides.map(a => p[a.key].value)));
  const scale = (axis - top) / max;
  const ypos = v => axis - v * scale;
  // 棒の上の数：100 以上は整数に丸める（隣の棒の数と重ならないように。リード文は丸めない）
  const num = (v, signed) => {
    const r = Math.abs(v) >= 100 ? Math.round(v) : v;
    return signed ? mtgSigned_(r) : mtgAmount_(r);
  };
  // 名前・数の箱は、文字の左右の余白（padX）の分だけ棒の幅より広く（折り返さないように）
  const pad = S.padX + 4;
  const label = (i, text, size) => el.push(mtgBox_(lefts[i] - pad / 2, axis + 3, widths[i] + pad, String(text).split('\n').length * 11 + 4, text,
    {size: size || 8.5, bold: true, color: c.ink, align: 'CENTER', valign: 'TOP', role: 'plan-label'}));
  // 隣の列の数と同じくらいの高さで横に重なるとき（3月と合計など）は、1行上げる
  let prev = null;
  const value = (i, v, text) => {
    let top = ypos(v) - 15;
    const half = String(text).length * 3.2;
    if (prev && prev.i === i - 1 && Math.abs(prev.top - top) < 13 &&
      prev.half + half > (lefts[i] + widths[i] / 2) - (lefts[prev.i] + widths[prev.i] / 2)) top = prev.top - 13;
    prev = {i, top, half};
    el.push(mtgBox_(lefts[i] - pad / 2, top, widths[i] + pad, 14, text, {size: 8.5, bold: true, color: c.ink, align: 'CENTER', role: 'plan-value'}));
  };
  const stack = (i, base, values, role) => {
    let level = base;
    ranks.forEach(r => {
      const v = values[r.key] || 0;
      if (v <= 0) return;
      const h = Math.max(v * scale, 0.75);
      el.push({type: 'box', x: bx(i), y: ypos(level) - h, w: bw, h, fill: r.color, line: c.white, lineWeight: 0.5, shape: 'RECTANGLE', role});
      level += v;
    });
  };
  el.push({type: 'box', x, y: axis, w, h: 0.75, fill: c.faint, shape: 'RECTANGLE'});
  // 計画の高さの横線（点線。12か月と合計の上に重ねる。どの月で計画に届くかが分かる）
  el.push({type: 'line', x, y: ypos(p.plan), w: unit * 13, h: 0, color: c.dark, weight: 1, dash: true, role: 'plan-line'});
  let level = 0;
  p.months.forEach((mo, i) => {
    stack(i, level, mo.values, 'plan-month');
    if (mo.total) value(i, level + mo.total, num(mo.total, true));
    level += mo.total;
    label(i, (i === 0 || mo.m === 1 ? String(mo.y).slice(2) + '/' : '') + mo.m + '月');
  });
  stack(12, 0, p.byRank, 'plan-total');
  value(12, p.total, num(p.total));
  label(12, '合計');
  el.push({type: 'box', x: bx(13), y: ypos(p.plan), w: bw, h: Math.max(p.plan * scale, 0.75), fill: c.white, line: c.dark, lineWeight: 1,
    shape: 'RECTANGLE', role: 'plan-plan'});
  value(13, p.plan, num(p.plan));
  label(13, '計画');
  const lo = Math.min(p.total, p.plan);
  const hi = Math.max(p.total, p.plan);
  el.push({type: 'box', x: bx(14), y: ypos(hi), w: bw, h: Math.max((hi - lo) * scale, 0.75), fill: diff < 0 ? c.graySoft : c.up,
    line: diff < 0 ? c.down : undefined, lineWeight: diff < 0 ? 0.75 : undefined, dash: diff < 0, shape: 'RECTANGLE', role: 'plan-gap'});
  value(14, hi, num(diff, true));
  label(14, diff < 0 ? '不足' : '超過');
  // 積み上げに含まない案件（参考）：区切りの縦線と見出し、0 からの確度別の棒、名前と件数（2行。小さめの字）
  el.push({type: 'box', x: x + unit * 15.25, y: top - 6, w: 0.75, h: axis - top + 6, fill: c.line, shape: 'RECTANGLE', role: 'plan-divider'});
  el.push(mtgBox_(lefts[15] - pad / 2, y + 21, x + w - lefts[15] + pad / 2, 14, '積み上げに含まない（参考）',
    {size: 8, color: c.muted, align: 'CENTER', role: 'plan-aside-title'}));
  asides.forEach((a, k) => {
    const i = 15 + k;
    const b = p[a.key];
    stack(i, 0, b.values, 'plan-aside');
    value(i, b.value, num(b.value));
    label(i, a.label + '\n' + b.count + '件', 8);
  });
  return {title, lead, elements: el};
}

/**
 * 統合版の計画に対する積み上げの2ページ目（同じセクション）「サブインダストリー別」：サブインダストリーごとに、今年度の期待値の積み上げ（確度別の横棒。
 * 前のページと同じ数え方）と計画（点線）。右に積み上げ・達成率・計画との差・確度A・B の割合・時期未定（積み上げに含まない）。達成率の高い順で、
 * いちばん下に合計。作れないシート（期待値・売上開始月の列が無い）は下に「データなし」と理由
 */
function mtgPlanBySubSlide_(summaries, now) {
  const S = MTG_SLIDE;
  const o = MTG_OPTIONS;
  const c = o.colors;
  const ranks = o.plan.ranks.concat([{key: o.plan.otherLabel, color: o.plan.otherColor}]);
  const title = MTG_TITLES.plan + '（サブインダストリー別）';
  const rate = p => (p.plan > 0 ? Math.round(p.total / p.plan * 100) : null);
  const sure = p => (p.byRank.A || 0) + (p.byRank.B || 0);
  const pct = (v, total) => (total > 0 ? Math.round(v / total * 100) + '%' : '—');
  const ok = summaries.filter(s => !s.plan.gap).map(s => ({name: s.name, p: s.plan}))
    .sort((a, b) => ((rate(b.p) || 0) - (rate(a.p) || 0)) || (b.p.total - a.p.total));
  const lacking = summaries.filter(s => s.plan.gap);
  const merged = mtgMergePlans_(ok.map(x => x.p));
  const first = mtgMonthOf_(merged.start);
  const last = mtgMonthOf_(merged.start + 11);
  const period = String(first.y).slice(2) + '/' + first.m + '〜' + String(last.y).slice(2) + '/' + last.m;
  const reached = ok.filter(x => x.p.total >= x.p.plan && x.p.plan > 0);
  const low = ok[ok.length - 1];
  const lead = ok.length
    ? '達成率の高い順。計画に届いているのは ' + ok.length + 'のうち ' + reached.length + '。' +
      (low && rate(low.p) != null ? '最も低いのは ' + low.name + '（' + rate(low.p) + '%）。' : '') + '\n時期未定：売上開始月が空・読めない案件の期待値（積み上げに含まない）。'
    : '計画に対する積み上げは' + o.gapText + '（どのシートにも期待値・売上開始月の列が無い）。';
  const x = S.x;
  const w = S.w;
  const y = S.top;
  const el = [mtgBox_(x, y - 2, 420, 20, '今年度（' + period + '）の期待値の積み上げ（百万円）', {size: 10.5, bold: true, color: c.ink})];
  // 凡例：確度の色と、計画の点線（右端を図の右端にそろえる）
  const legend = ranks.map(r => ({label: r.key, color: r.color})).concat([{label: '計画', plan: true}]);
  legend.forEach(item => { item.bw = mtgTextWidth_(item.label, 8.5) + 16 + (item.plan ? 10 : 0); });
  let lx = x + w - legend.reduce((sum, item) => sum + 6 + item.bw, 0) - (legend.length - 1) * (8 - 7.2 - 1.6);
  legend.forEach(item => {
    if (item.plan) {
      el.push({type: 'line', x: lx + 2, y: y + 8.5, w: 12, h: 0, color: c.dark, weight: 1, dash: true, role: 'legend'});
      el.push(mtgBox_(lx + 16, y - 1, item.bw - 10, 18, item.label, {size: 8.5, color: c.sub, role: 'legend'}));
    } else {
      el.push({type: 'box', x: lx, y: y + 5, w: 7, h: 7, fill: item.color, shape: 'RECTANGLE', role: 'legend'});
      el.push(mtgBox_(lx + 6, y - 1, item.bw, 18, item.label, {size: 8.5, color: c.sub, role: 'legend'}));
    }
    lx += 6 + item.bw + (8 - 7.2 - 1.6);
  });
  if (!ok.length) {
    el.push(mtgEmptyBox_(x, y + 26, 360, 64, 'gap', lacking.map(s => s.name + '：' + s.plan.gap).join('\n')));
    return {title, lead, elements: el};
  }
  // 列：名前・横棒・積み上げ・達成率・計画との差・確度A・B・時期未定
  const nameW = 112;
  const cols = [{label: '積み上げ', w: 52}, {label: '達成率', w: 44}, {label: '計画との差', w: 58}, {label: '確度A・B', w: 50}, {label: '時期未定', w: 54}];
  const colsW = cols.reduce((sum, col) => sum + col.w, 0);
  const barX = x + nameW + 6;
  const barW = w - nameW - 6 - colsW - 10;
  const headY = y + 22;
  const rows = ok.map(item => Object.assign({total: false}, item)).concat([{name: '合計', p: merged, total: true}]);
  const lines = lacking.length ? 1 : 0;
  const rowH = Math.min(22, (S.bottom - headY - 16 - lines * 14) / rows.length);
  const max = Math.max.apply(null, [1].concat(ok.map(item => Math.max(item.p.total, item.p.plan))));
  const scale = barW / max;
  let cx = barX + barW + 10;
  cols.forEach(col => {
    el.push(mtgBox_(cx - S.padX / 2, headY, col.w + S.padX, 14, col.label, {size: 8, color: c.sub, align: 'CENTER', role: 'plan-sub-head'}));
    col.x = cx;
    cx += col.w;
  });
  rows.forEach((item, i) => {
    const ry = headY + 16 + rowH * i;
    const p = item.p;
    if (item.total) el.push({type: 'box', x, y: ry, w, h: 0.75, fill: c.line, shape: 'RECTANGLE'});
    else if (i % 2 === 1) el.push({type: 'box', x, y: ry, w, h: rowH, fill: c.greenSoft, alpha: 0.35, shape: 'RECTANGLE', role: 'plan-sub-band'});
    el.push(mtgBox_(x, ry, nameW, rowH, item.name, {size: 8.5, bold: true, color: c.ink, role: 'plan-sub-name'}));
    // 合計の行は棒を描かない（目盛りが違う。全体の図は前のページ）
    if (!item.total) {
      let level = 0;
      const bh = Math.max(rowH - 8, 6);
      ranks.forEach(r => {
        const v = p.byRank[r.key] || 0;
        if (v <= 0) return;
        el.push({type: 'box', x: barX + level * scale, y: ry + (rowH - bh) / 2, w: Math.max(v * scale, 0.75), h: bh, fill: r.color, line: c.white,
          lineWeight: 0.5, shape: 'RECTANGLE', role: 'plan-sub-bar'});
        level += v;
      });
      el.push({type: 'line', x: barX + p.plan * scale, y: ry + 1, w: 0, h: rowH - 2, color: c.dark, weight: 1, dash: true, role: 'plan-sub-plan'});
    }
    const r = rate(p);
    const diff = p.total - p.plan;
    const undated = p.undated || {count: 0, value: 0};
    const values = [mtgAmount_(Math.round(p.total)), r == null ? '—' : r + '%', mtgSigned_(Math.round(diff)), pct(sure(p), p.total),
      undated.count ? mtgAmount_(Math.round(undated.value)) + '（' + undated.count + '件）' : '—'];
    values.forEach((text, k) => {
      const short = k === 2 && diff < 0;
      el.push(mtgBox_(cols[k].x - S.padX / 2, ry, cols[k].w + S.padX, rowH, text, {size: k === 4 ? 7.5 : 8.5, bold: k <= 2,
        color: short ? c.dark : k === 1 && r != null && r >= 100 ? c.greenDark : c.ink, align: 'END', role: 'plan-sub-value'}));
    });
  });
  if (lacking.length) {
    el.push(mtgBox_(x, S.bottom - 14, w, 14, mtgWrap_(o.gapText + '：' + lacking.map(s => s.name + '（' + s.plan.gap + '）').join('・'), w, 8, 1)[0],
      {size: 8, color: c.muted, role: 'plan-sub-gap'}));
  }
  return {title, lead, elements: el};
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
  const startLabel = s.first ? MTG_OPTIONS.firstDays + '日前' : mtgFromWord_(s);
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

/** 3 前回から変化のある案件：ステータス・期待値・受注月・活動状況が変わった案件（今のステータスの順（mtgStatusGroups_）、その中は想定売上の大きい順） */
function mtgMovesSlide_(s, withAppendix) {
  // 全件の一覧の場所：この資料の Appendix ①（統合版は Appendix を入れないので、各サブインダストリーの資料の Appendix ①）
  const all = withAppendix === false ? '全件は各サブインダストリーの資料の Appendix ① を参照。' : '全件は Appendix ① を参照。';
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  const lack = s.gaps || {};
  const moves = s.changes.filter(ch => ch.category !== 'updated');
  const columns = mtgColumns_([126, 126, 60, 76, 214, 58], 4);   // ステータスの列は「C-稟議対応中(50%) ▶」「B-クロージング中(70%)」が1行ずつ入る幅。活動状況は広く
  const rows = [mtgHeadRow_(columns, ['得意先／案件名', 'ステータス\n' + mtgFromWord_(s) + '▶今回', '期待値\n' + mtgFromWord_(s) + '▶今回', '受注月\n' + mtgFromWord_(s) + '▶今回', '活動状況（最新）', '責任者'])];
  // 上から mainRows.changes 件まで、ページに入る分だけ（行の高さは中身で変わる）
  const top = [];
  moves.slice(0, MTG_OPTIONS.mainRows.changes).some(ch => {
    const row = mtgRow_(columns, [
      mtgDealCell_(ch, columns[0]),
      {text: mtgStatusMove_(ch), lines: 3, fill: c.white},   // 長いステータスの名前（「C-稟議対応中(50%) ▶ B-クロージング中(70%)」）は3行まで
      mtgFieldCell_(s, 'weighted', mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'orderMonth', mtgMonthMove_(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      // 長い文は小さくして入れる（本編は件数を優先して行は伸ばさない。全文は Appendix ①）
      mtgFieldCell_(s, 'activity', ch.activity, {lines: 3, fill: c.white, fit: true}),
      mtgFieldCell_(s, 'owner', ch.owner, {lines: 2, fill: c.white, align: 'CENTER'})
    ]);
    if (S.top + rows.reduce((sum, r) => sum + r.h, 0) + row.h > S.bottom) return true;
    rows.push(row);
    top.push(ch);
    return false;
  });
  const el = [];
  if (top.length) {
    el.push({type: 'table', x: S.x, y: S.top, columns, rows});
  } else {
    // 表の見出しの行の下に、変更の記録が無ければ「データなし」と理由、変化が無ければ「該当なし」（状態はリード文に書くので、同じ文を2回書かない）
    el.push.apply(el, mtgEmptyTable_(columns, rows[0], lack.changes ? 'gap' : 'none', lack.changes || ''));
  }
  return {
    title: MTG_TITLES.progress,
    lead: lack.changes ? mtgChangedLabel_(s) + 'は' + MTG_OPTIONS.gapText + '（' + lack.changes + '）。'
      : (moves.length ? mtgChangedLabel_(s) + 'は ' + moves.length + '件' + mtgUpdatedNote_(s) + '。'
      : mtgPeriodShort_(s) + '、ステータス・期待値・受注月・活動状況に変化のある案件なし' + mtgUpdatedNote_(s) + '。') +
      (moves.length > top.length ? '\n今回のステータスの順（確度 A → 創注 → 失注・保留・削除）に ' + top.length + '件を掲載。' + all
        : k.updated ? '\n' + all : ''),
    elements: el
  };
}

/** 4 期待値上位案件の状況・課題：進行中の案件の期待値の大きい順（summary.focus）。当てはまる状況（停滞・受注月超過・延期）は判定の中身ごと文で（mtgFlagTexts_） */
function mtgFocusSlide_(s) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const k = s.kpi;
  // 活動状況の列を広く（長い文は小さくし、それでも入らなければ行を伸ばして次のページへ。全文を載せる）
  // 該当事項は判定の中身を文で書く（単語のラベルだけでは何を見て判定したか分からないため。2026/10/19 から）ので広く
  const columns = mtgColumns_([106, 62, 96, 62, 120, 104, 120], 4);
  const head = mtgHeadRow_(columns, ['得意先／案件名', '期待値／\n想定売上', 'ステータス\n（滞留日数）', '受注月', '活動状況', '活動における課題',
    '該当事項']);
  const rows = [head];
  const G = MTG_OPTIONS.gapText;
  const lack = s.gaps || {};
  // RXシナリオ構成の列があれば、案件名の下に小さな字で（複数は「・」でつなぐ）
  const dealCell = d => {
    const cell = mtgDealCell_(d, columns[0]);
    if (!s.scenarioColumn) return cell;
    const line = d.scenarios && d.scenarios.length ? d.scenarios.join('・') : 'RXシナリオ：' + MTG_OPTIONS.blankText;
    return Object.assign(cell, {text: cell.text + '\n' + mtgWrap_(line, columns[0], 7.5, 1)[0], lines: 3,
      parts: cell.parts.concat([{size: 7.5, color: c.muted}])});
  };
  s.focus.forEach(d => rows.push(mtgRow_(columns, [
    dealCell(d),
    // 期待値は大きな緑の字。空なら「未入力」を本文の大きさ・グレーで（数のように見せない）。列が無ければ灰色のセルに「データなし」
    lack.weighted && lack.amount ? mtgGapCell_({lines: 2, align: 'CENTER'})
      : {text: (lack.weighted ? G : mtgAmount_(d.weighted)) + '\n' + (lack.amount ? G : mtgAmount_(d.amount)),
        parts: [lack.weighted || d.weighted == null ? {color: c.muted} : {bold: true, color: c.greenDark, size: 11}, lack.amount ? {color: c.muted} : {}],
        lines: 2, fill: lack.weighted || lack.amount ? c.gap : c.white, align: 'CENTER'},
    {text: mtgStatusLines_(d.status || MTG_OPTIONS.blankText, columns[2]) + '\n' + (d.days == null ? '記録なし' : d.days + '日' + (d.atLeast ? '以上' : '')),
      parts: mtgStatusLines_(d.status || MTG_OPTIONS.blankText, columns[2]).split('\n').map(() => ({})).concat([{color: c.muted}]),
      lines: 3, fill: c.white, align: 'CENTER'},
    mtgFieldCell_(s, 'orderMonth', mtgMonthLabel_(d.orderMonthText), {fill: c.white, align: 'CENTER'}),
    mtgFieldCell_(s, 'activity', d.activity, {lines: 3, fill: c.white, fit: true, grow: true}),
    mtgFieldCell_(s, 'issue', d.issue, {lines: 3, fill: c.white, fit: true, grow: true}),
    mtgFlagTexts_(d).length ? {text: mtgFlagTexts_(d).join('\n'), size: 8, lines: 99, fill: c.white}
      : {text: MTG_OPTIONS.noneText, fill: c.white, align: 'CENTER'}
  ], 44)));
  const share = k.weighted > 0 ? Math.round(mtgSum_(s.focus, 'weighted') / k.weighted * 100) : 0;
  // 期待値の列が無ければ、想定売上の大きい順（並べ方と、期待値はデータなしであることを書く）
  const lead = !s.focus.length ? '進行中案件なし。'
    : lack.weighted ? '期待値は' + G + '（' + lack.weighted + '）のため、進行中案件のうち想定売上上位 ' + s.focus.length + '件。'
      : '進行中案件のうち期待値上位 ' + s.focus.length + '件（進行中案件の期待値合計の ' + share + '%）。';
  if (!s.focus.length) return [{title: MTG_TITLES.focus, lead, elements: mtgEmptyTable_(columns, head, 'none', '')}];
  const pages = mtgTablePages_(head, rows.slice(1));
  return pages.map((p, n) => {
    const table = {type: 'table', x: S.x, y: S.top, columns, rows: p.rows};
    const el = [table];
    return {title: MTG_TITLES.focus + (n ? '（続き）' : ''), lead, elements: el};
  });
}

/**
 * 表の行をページに分ける：行が伸びてページに入らなければ、次のページ（「（続き）」）に続ける。各ページの先頭に見出しの行（head）。
 * 戻り値 [{rows: [head, …], index: [rows の何番目か]}]
 */
function mtgTablePages_(head, rows) {
  const S = MTG_SLIDE;
  const pages = [];
  rows.forEach((row, i) => {
    let page = pages[pages.length - 1];
    if (!page || page.y + row.h > S.bottom) {
      page = {rows: [head], y: S.top + head.h, index: []};
      pages.push(page);
    }
    page.rows.push(row);
    page.index.push(i);
    page.y += row.h;
  });
  return pages;
}

/**
 * 期待値上位案件の状況・課題の続き「直近2か月の動き」（同じセクション。案件は前のページと同じ並び）：案件ごとに、直近 recentMonths か月の
 * 変更の記録（summary.recent。ステータス・受注月・期待値・活動状況・次回アクションの変更。新しい順）と、次のアクション（今の次回アクションと受注月。
 * 受注月が今月・来月なら「今月受注予定」「来月受注予定」、過ぎていれば「受注月超過」）。記録の始まりより前の動きは分からないので、リード文に書く
 */
function mtgRecentSlide_(s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const G = o.gapText;
  const lack = s.gaps || {};
  const recent = s.recent;
  if (!s.focus.length || !recent) return [];
  const tz = Session.getScriptTimeZone();
  const day = t => Utilities.formatDate(new Date(t), tz, 'M/d');
  const months = '直近' + o.recentMonths + 'か月';
  const title = MTG_TITLES.focus + '（' + months + 'の動き）';
  const columns = mtgColumns_([130, 360, 170], 1);
  const head = mtgHeadRow_(columns, ['得意先／案件名', months + 'の動き（' + day(recent.from) + '〜。新しい順）', '次回アクション／受注月']);
  const noLog = recent.logFrom == null;
  // 記録の始まり：記録のシートの最初の行の時刻。それより前の日時の記録（ダミー）を出すときは、その日から
  const earliest = Math.min.apply(null, [Infinity].concat([].concat.apply([], recent.moves).map(m => m.at)));
  const logFrom = noLog ? null : Math.min(recent.logFrom, earliest);
  const rows = s.focus.map((d, i) => {
    const moves = recent.moves[i] || [];
    // 動きは同じ大きさの字で全文（案件ごとに字の大きさが変わると読みにくい）。多ければ新しいほうから recentMax 件と「ほか N件」
    const shown = moves.slice(0, o.recentMax).map(m => m.text);
    if (moves.length > shown.length) shown.push('ほか ' + (moves.length - shown.length) + '件（' + day(moves[moves.length - 1].at) + '〜）');
    const history = noLog ? mtgGapCell_({lines: 2}, G + '（変更の記録が無い）')
      : moves.length ? {text: shown.join('\n'), size: 8, lines: 99, fill: c.white,
        parts: shown.map((t, k) => (k === o.recentMax ? {color: c.muted} : {}))}
        : {text: '記録された動きなし', color: c.muted, fill: c.white};
    // 次のアクション：次回アクション（列が無ければデータなし）と、受注月（今月・来月の受注予定・受注月超過）
    const index = mtgMonthIndex_(d.orderMonth);
    const when = index == null ? '' : index < s.month ? '受注月超過' : index === s.month ? '今月受注予定' : index === s.month + 1 ? '来月受注予定' : '';
    const order = lack.orderMonth ? '受注月：' + G : '受注月：' + mtgMonthLabel_(d.orderMonthText) + (when ? '（' + when + '）' : '');
    const action = lack.nextAction ? G : d.nextAction || o.blankText;
    const next = {text: action + '\n' + order, size: 8, lines: 99, fill: c.white,
      parts: [lack.nextAction || !d.nextAction ? {color: c.muted} : {}, when ? {bold: true, color: when === '受注月超過' ? c.dark : c.greenDark} : {color: c.muted}]};
    return mtgRow_(columns, [mtgDealCell_(d, columns[0]), history, next], 40);
  });
  const lead = noLog ? months + 'の動きは' + G + '（「' + DIFF_OPTIONS.logSheet + '」シートに変更の記録が無い）。次のアクションは今のシートから。'
    : '期待値上位 ' + s.focus.length + '件の、' + months + '（' + day(recent.from) + '〜）に記録された動き（ステータス・受注月・期待値・活動状況・次回アクション）。' +
      (logFrom > recent.from ? '\n変更の記録は ' + day(logFrom) + ' からのため、それより前の動きは出ない。' : '');
  return mtgTablePages_(head, rows).map((p, n) => ({title: title + (n ? '（続き）' : ''), lead,
    elements: [{type: 'table', x: MTG_SLIDE.x, y: MTG_SLIDE.top, columns, rows: p.rows}]}));
}

/** 列の幅（合わせて 660pt で決めた幅）を中身の幅（MTG_SLIDE.w）に合わせる。差は flex 番目の列（文の列）で調整する */
function mtgColumns_(columns, flex) {
  const out = columns.slice();
  out[flex] += MTG_SLIDE.w - columns.reduce((a, b) => a + b, 0);
  return out;
}

/**
 * 当てはまる状況を、判定の中身ごと文で書く（期待値上位案件の「該当事項」。単語のラベルだけでは、入力した人に何を見て判定したかが分からないため）：
 * 停滞「ステータスが 52日 変わっていない（45日以上で停滞）」・受注月超過「受注月 2026/9 を過ぎても受注になっていない」・
 * 延期「受注月を後ろにずらした記録 2回（過去180日）」
 */
function mtgFlagTexts_(d) {
  const o = MTG_OPTIONS;
  return [
    d.stalled && 'ステータスが ' + d.days + '日' + (d.atLeast ? '以上' : '') + ' 変わっていない（' + o.stallDays + '日以上で停滞）',
    d.overdue && '受注月 ' + mtgMonthLabel_(d.orderMonthText) + ' を過ぎても受注になっていない',
    d.slips && '受注月を後ろにずらした記録 ' + d.slips + '回（過去' + o.lookbackDays + '日）'
  ].filter(Boolean);
}

/**
 * 変化のある案件をまとめる今のステータスのまとまり [{key, label}]：確度ごと（MTG_OPTIONS.plan.ranks。label は「C 提案中・選定対応中・稟議対応中」のように、
 * ステータスのプルダウン（statusOrder）のその確度の名前。プルダウンが無ければ確度だけ）と、最後に失注・保留・削除（ステータス未入力・その他も）。
 * サマリーの表がページに入るよう、失注・保留・削除は1つにまとめる（行は6つ）
 */
function mtgStatusGroups_(statusOrder) {
  const names = key => (statusOrder || []).filter(st => mtgStatusKind_(st) !== 'lost' && mtgStatusKind_(st) !== 'hold' && mtgPlanRank_(st) === key)
    .map(st => String(st).normalize('NFKC').replace(/^[^-]*-/, '').replace(/\(.*$/, '').trim()).filter(Boolean);
  return MTG_OPTIONS.plan.ranks.map(r => ({key: r.key, label: r.key + (names(r.key).length ? ' ' + names(r.key).join('・') : '')}))
    .concat([{key: 'closed', label: '失注・保留・削除'}]);
}

/** 変化のある案件の、今のステータスのまとまり（mtgStatusGroups_ の key。確度の無いステータス（失注・保留・削除・未入力など）は 'closed'） */
function mtgStatusGroupOf_(ch) {
  if (ch.removed || ch.category === 'removed') return 'closed';
  const kind = mtgStatusKind_(ch.statusTo || '');
  if (kind === 'lost' || kind === 'hold') return 'closed';
  const rank = ch.statusTo ? mtgPlanRank_(ch.statusTo) : '';
  return MTG_OPTIONS.plan.ranks.some(r => r.key === rank) ? rank : 'closed';
}

/** months か月前の同じ日の 0時（時刻 ms。その月に同じ日が無ければ月末） */
function mtgMonthsAgo_(now, months) {
  const tz = Session.getScriptTimeZone();
  const p = Utilities.formatDate(now, tz, 'yyyy-MM-dd').split('-').map(Number);
  const index = p[0] * 12 + p[1] - months;
  const mo = mtgMonthOf_(index);
  const last = new Date(Date.UTC(mo.y, mo.m, 0)).getUTCDate();
  const text = mo.y + '/' + String(mo.m).padStart(2, '0') + '/' + String(Math.min(p[2], last)).padStart(2, '0');
  return Utilities.parseDate(text, tz, 'yyyy/MM/dd').getTime();
}

/**
 * 期待値上位案件の直近の動き：案件の変更の記録（mtgHistory_ の events。古い順）のうち from 以降の、ステータス・受注月・期待値・活動状況・
 * 次回アクションの変更と、案件の登録。同じ日の同じ項目の変更は1つにまとめる（その日の最初の変更前 → 最後の変更後。同じ値に戻ったものは出さない）。
 * [{at, key: 項目（MTG_OPTIONS.columns のキー。登録は 'born'）, text: 資料に出す文}]（新しい順）
 */
function mtgRecentMoves_(events, from) {
  const labels = mtgLabels_();
  const keys = ['status', 'orderMonth', 'weighted', 'activity', 'nextAction'];
  const tz = Session.getScriptTimeZone();
  const groups = new Map();
  // 記録のシートの行の順は時刻の順とは限らない（ダミーの記録は過去の日時で足す）ので、時刻の順に並べてからまとめる
  events.slice().sort((a, b) => a.at - b.at).forEach(e => {
    if (e.at < from) return;
    const key = e.born ? 'born' : keys.find(k => labels[k] && labels[k] === e.label);
    if (!key) return;
    const id = Utilities.formatDate(new Date(e.at), tz, 'yyyy/MM/dd') + '|' + key;
    const g = groups.get(id);
    if (g) {
      g.after = e.after;
      g.at = e.at;
    } else {
      groups.set(id, {at: e.at, key, before: e.before, after: e.after});
    }
  });
  const arrow = (a, b) => a + ' ▶ ' + b;
  const blank = v => v || MTG_OPTIONS.blankText;
  const amount = v => (mtgNumber_(v) == null ? blank(v) : mtgAmount_(mtgNumber_(v)));
  // 受注月：値の貼り付けで日付がシリアル値（「46357」）のまま記録されたものは、日付に直す
  const month = v => {
    if (!v) return blank(v);
    if (/^\d{5}(\.\d+)?$/.test(v)) return Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(v)) * 86400000), 'UTC', 'yyyy/M');
    return mtgMonthLabel_(v);
  };
  const body = g => {
    if (g.key === 'born') return '案件を登録';
    if (g.key === 'status') return 'ステータス：' + arrow(blank(g.before), blank(g.after));
    if (g.key === 'orderMonth') return '受注月：' + arrow(month(g.before), month(g.after));
    if (g.key === 'weighted') return '期待値：' + arrow(amount(g.before), amount(g.after));
    return mtgShortHeader_(MTG_OPTIONS.columns[g.key]) + '：' + mtgAddedText_(g.before, g.after);
  };
  return Array.from(groups.values())
    .filter(g => g.key === 'born' || g.before !== g.after)
    .sort((a, b) => b.at - a.at)
    .map(g => ({at: g.at, key: g.key, text: Utilities.formatDate(new Date(g.at), tz, 'M/d') + '　' + body(g)}));
}

/**
 * 文の変更のうち、書き足した・書き換えた部分（前と後で同じ頭・おしりを除いた、後の文）。改行は「／」に。
 * 頭・おしりは文の切れ目（改行・句読点・空白など）でだけ除く（「提案中」→「提案書を提出」を「書を提出」にしない。切れ目が無ければ後の文の全文）。
 * 消しただけなら「（一部を削除）」、全部消したなら「（削除）」
 */
function mtgAddedText_(before, after) {
  const a = String(before || '');
  const b = String(after || '');
  if (!b.trim()) return '（削除）';
  let head = 0;
  while (head < a.length && head < b.length && a[head] === b[head]) head++;
  let tail = 0;
  while (tail < a.length - head && tail < b.length - head && a[a.length - 1 - tail] === b[b.length - 1 - tail]) tail++;
  if (!b.slice(head, b.length - tail).trim()) return '（一部を削除）';
  const cut = ch => /[\n。．.！!？?、，,；;／\s]/.test(ch);
  // 切れ目：その位置の前か後ろの字が区切り（「…ヒアリング」＋「\n提案書」のように、書き足しが改行から始まるときも）
  const at = i => i <= 0 || i >= b.length || cut(b[i - 1]) || cut(b[i]);
  while (!at(head)) head--;
  let end = b.length - tail;
  while (!at(end)) end++;
  const added = b.slice(head, end).replace(/\s*\n\s*/g, '／').replace(/^[／\s]+|[／\s]+$/g, '');
  return added || '（一部を削除）';
}

/**
 * 表のセルの上に重ねるラベル（角の丸い四角。区分・当てはまる状況）。table は表の部品（x・y・columns・rows）、row・col はセルの位置
 * （見出しの行が 0）、labels は [{text, tone}]（tone：up 前向き＝濃い緑・down 後ろ向き＝濃いグレー・note そのほか＝薄いグレー・stall 停滞＝薄い緑。
 * どれも枠なし）。width はラベルの幅（列の中で同じにする。mtgPillWidth_）。複数は縦に並べ、セルの上下の中央に置く。セルの文字は空にしておく
 */
function mtgPills_(table, row, col, labels, width) {
  const c = MTG_OPTIONS.baseColors || MTG_OPTIONS.colors;   // Appendix をグレーにしていても、ラベルは意味の色のまま
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

/** Appendix ① 前回から変化のある案件（一覧。その他更新・削除も）。本編と同じ今のステータスの順 */
function mtgAppendixChanges_(s) {
  const c = MTG_OPTIONS.colors;
  const lack = s.gaps || {};
  const columns = mtgColumns_([118, 120, 60, 72, 156, 74, 60], 4);   // 活動状況は広く（長い文は小さく、それでも入らなければ行を伸ばして全文）
  return mtgPagedTable_({
    title: MTG_TITLES.changes,
    lead: lack.changes ? mtgChangedLabel_(s) + 'は' + MTG_OPTIONS.gapText + '（' + lack.changes + '）。'
      : mtgChangedLabel_(s) + 'は ' + s.kpi.changed + '件' + mtgUpdatedNote_(s) + '。' +
        (s.unmatched ? '\n案件と照合できなかった変更記録 ' + s.unmatched + '件は集計対象外。' : '') +
        (s.unlogged ? '\nうち ' + s.unlogged + '件は変更の記録が無い（' + mtgCopyWord_(s) + 'との比較で検出）。' : '') +
        (s.copyError ? '\n' + mtgCopyWord_(s) + 'を開けないため、変更の記録から集計。' : ''),
    columns,
    head: ['得意先／案件名', 'ステータス\n' + mtgFromWord_(s) + '▶今回', '期待値\n' + mtgFromWord_(s) + '▶今回', '受注月\n' + mtgFromWord_(s) + '▶今回', '活動状況（最新）', '更新項目', '責任者'],
    rows: s.changes.map(ch => [
      mtgDealCell_(ch, columns[0], MTG_SLIDE.appendixSize),
      {text: mtgStatusMove_(ch), lines: 3, fill: c.white},
      mtgFieldCell_(s, 'weighted', mtgValueMove_(ch.weightedFrom, ch.weightedTo, ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      mtgFieldCell_(s, 'orderMonth', mtgMonthMove_(ch), {lines: 2, fill: c.white, align: 'CENTER'}),
      ch.removed ? {text: '（削除）', lines: 2, fill: c.white} : mtgFieldCell_(s, 'activity', ch.activity, {lines: 2, fill: c.white, fit: true, grow: true}),
      {text: ch.category === 'new' ? '（新規追加）' : ch.columns.map(mtgShortHeader_).join('・') || (ch.removed ? '（削除）' : MTG_OPTIONS.noneText), lines: 2,
        fill: c.white},
      mtgFieldCell_(s, 'owner', ch.owner, {lines: 2, fill: c.white, align: 'CENTER'})
    ]),
    // 変化が無い・記録が無いときも、表の見出しの行を置き、その下に「該当なし」・「データなし」と理由
    empty: lack.changes ? {kind: 'gap', text: lack.changes} : {kind: 'none'}
  });
}

/** Appendix ② 停滞案件（一覧）。同じステータスの日数と、最新の活動状況（前回以降に更新があるかが分かるように） */
function mtgAppendixStalled_(s) {
  const c = MTG_OPTIONS.colors;
  const k = s.kpi;
  const columns = mtgColumns_([100, 44, 44, 86, 60, 60, 34, 60, 76, 96], 9);   // 課題・活動状況は広く（長い文は小さく、それでも入らなければ行を伸ばして全文）
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
      mtgFieldCell_(s, 'issue', d.issue, {lines: 2, fill: c.white, fit: true, grow: true}),
      mtgFieldCell_(s, 'activity', d.activity, {lines: 2, fill: c.white, fit: true, grow: true})
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
  const types = mtgAlertTypes_();
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
    startMonth: '売上開始月', activity: '活動状況', issue: '課題', owner: '責任者', delivery: 'デリバリー担当', nextAction: '次回アクション',
    scenario: 'RXシナリオ構成', record: '商談記録'};
  const key = Object.keys(names).find(k => mtgColumnIndex_([header], k) === 0);
  return key ? names[key] : String(header).replace(/\s+/g, '');
}

/** 「前回会議（10/9）の資料作成時（10/07 18:00）以降」「直近14日間（09/30〜。初回のため）」 */
function mtgPeriodText_(s) {
  const tz = Session.getScriptTimeZone();
  if (s.previous && s.previous.chosen) return mtgChosenText_(s.previous) + '（選んだ時点）';
  if (s.previous && s.previous.baseline) {
    return '基準時点（' + Utilities.formatDate(new Date(s.previous.cutoff), tz, 'M/d HH:mm') + '）以降';
  }
  if (s.previous && s.previous.scheduled) return '前回会議（' + mtgShortDate_(s.previous.date) + '）以降';
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
 * 表の1行。cells: [{text, lines, size, bold, color, fill, align, parts, fit, grow}]（文字は列の幅で折り返し、lines 行（既定は1）を超えたら「…」）。
 * fit：lines 行に入らなければ、入るまで文字を小さくする（minFitSize まで）。grow：それでも入らなければ、行を伸ばして全文を入れる（maxGrowLines まで）。
 * parts は改行で分けた部分ごとの書式 [{bold, color, size}]（スライドでは spans にする）。parts の代わりに spans を直接渡してもよい（1行の中で大きさを変えるとき）。
 * 行の高さは、いちばん高いセルに合わせる（min 以上。mtgTextHeight_）。Google スライドでは、中身がこの高さより高ければ行が伸びるので、少し大きめに見積もる
 */
function mtgRow_(columns, cells, min) {
  let h = min || 0;
  const out = cells.map((cell, i) => {
    const {size, lines} = mtgFitLines_(cell, columns[i]);
    const text = lines.join('\n');
    const spans = cell.parts ? mtgPartSpans_(String(cell.text), cell.parts, columns[i], size, lines) : cell.spans;
    h = Math.max(h, mtgTextHeight_(lines, size, spans));
    return Object.assign({}, cell, {text, size}, spans ? {spans} : {});
  });
  return {h: Math.ceil(h * 10) / 10, cells: out};
}

/**
 * セルの文字の大きさと、折り返した行（mtgRow_）。fit なら、lines 行に全文が入るまで 0.5pt ずつ小さくし（MTG_SLIDE.minFitSize まで）、
 * それでも入らなければ grow なら行を増やして全文（maxGrowLines まで）、grow でなければ最後を「…」で切る
 */
function mtgFitLines_(cell, width) {
  const base = cell.size || MTG_SLIDE.bodySize;
  const max = cell.lines || 1;
  if (!cell.fit || !cell.text) return {size: base, lines: mtgBalance_(cell.text, width, base, max)};
  const count = size => mtgWrap_(cell.text, width, size, 999).length;
  let size = base;
  while (count(size) > max && size - 0.5 >= MTG_SLIDE.minFitSize) size = Math.round((size - 0.5) * 10) / 10;
  const need = count(size);
  const limit = need > max && cell.grow ? Math.min(need, MTG_SLIDE.maxGrowLines) : max;
  return {size, lines: mtgBalance_(cell.text, width, size, limit)};
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
 * 部品（slide.elements）は type：'box'（四角・文字）・'table'（表）・'line'（直線。{x, y, w, h, color, weight, dash}）。
 * options = {scale: 位置・大きさ・文字の倍率（ひな形のページの大きさ ÷ MTG_SLIDE.width。既定 1）, existing: true ならスライドは作らない（ひな形のページ）,
 *            idPrefix: 部品の ID の先頭（既定は pageId）}。
 * slide.template（表紙・扉）は、ひな形のページなら何もしない（決まった文字の置き換えは mtgTemplateTextRequests_）。ひな形が無ければ同じ位置に描く（mtgTemplateFallback_）
 */
function mtgSlideRequests_(slide, pageId, index, options) {
  const opt = options || {};
  const k = opt.scale || 1;
  const pt = v => mtgPt_(v * k);
  const requests = [];
  let elements = slide.elements;
  if (!opt.existing) {
    requests.push({createSlide: {objectId: pageId, insertionIndex: index, slideLayoutReference: {predefinedLayout: 'BLANK'}}});
    if (slide.template) {
      // ひな形が無いときの表紙・扉：背景は緑一色にしておき、あとで画像にする（mtgCoverBackground_）
      requests.push({updatePageProperties: {objectId: pageId, pageProperties: {pageBackgroundFill: {solidFill: {color: {rgbColor:
        mtgRgb_(MTG_OPTIONS.colors[mtgCoverTone_(slide)])}}}}, fields: 'pageBackgroundFill.solidFill.color'}});
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
    if (el.type === 'line') {
      // 直線（color・weight・dash で点線）。幅か高さが 0 の線は作れないことがあるので、ごくわずかな大きさにする（見た目は水平・垂直）
      elementProperties.size.width = pt(Math.max(width, 0.01));
      elementProperties.size.height = pt(Math.max(height, 0.01));
      requests.push({createLine: {objectId: id, lineCategory: 'STRAIGHT', elementProperties}});
      requests.push({updateLineProperties: {objectId: id, lineProperties: {lineFill: {solidFill: {color: {rgbColor: mtgRgb_(el.color)}}},
        weight: pt(el.weight || 0.75), dashStyle: el.dash ? 'DASH' : 'SOLID'}, fields: 'lineFill,weight,dashStyle'}});
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
 * 表紙・扉の色（'cover' 表紙＝緑・'divider' 扉＝グレー・統合版の表紙は MTG_OPTIONS.combined.coverTone（緑。'combined' なら紫）・
 * RX Service Meeting の表紙（service）は 'combined'（紫）。MTG_COVER・MTG_OPTIONS.colors のキー）
 */
function mtgCoverTone_(slide) {
  if (slide.service) return 'combined';
  return slide.divider ? 'divider' : slide.combined ? MTG_OPTIONS.combined.coverTone || 'cover' : 'cover';
}

/**
 * ひな形のページの使い方。slides の表紙・扉ごとに、ひな形のどのページを使うか（MTG_OPTIONS.templatePages）を決め、2回目からは複製する。
 * pages はひな形のコピーのページの ID（順に）。戻り値 {pageOf: {スライドの番号（0 から）: ページの ID}, requests: 使わないページを消し・複製するリクエスト}。
 * ひな形に扉のページが無ければ、扉は pageOf に入れない（描く）。統合版の表紙のページが無ければ、表紙のページを使う
 */
function mtgTemplatePlan_(slides, pages) {
  const P = MTG_OPTIONS.templatePages;
  const page = n => (n ? pages[n - 1] || null : null);
  // RX Service Meeting の表紙（service）は紫のページだけ（無ければ緑の表紙は使わずに、紫で描く）
  const source = slide => (slide.service ? page(P.service) : slide.cover ? (slide.combined && page(P.combined)) || page(P.cover)
    : slide.divider ? page(P.divider) : null);
  const used = {};
  const pageOf = {};
  const copies = [];
  slides.forEach((slide, i) => {
    const from = source(slide);
    if (!from) return;
    if (!used[from]) {
      used[from] = true;
      pageOf[i] = from;
      return;
    }
    pageOf[i] = 'mtg_t' + (i + 1);
    copies.push({duplicateObject: {objectId: from, objectIds: {[from]: pageOf[i]}}});
  });
  return {pageOf, requests: pages.filter(id => !used[id]).map(id => ({deleteObject: {objectId: id}})).concat(copies)};
}

/**
 * ひな形のページ（page。Slides API の Page）の決まった文字（MTG_OPTIONS.templateTexts）を、template = {title, subtitle, date, name} に書き換えるリクエスト。
 * 文字の枠ごとに書き換える（同じ文字の枠が2つあるため、replaceAllText は使わない）：枠の文字が決まった文字と同じ（X は5文字以上なら同じとみなす）枠を探し、
 * 同じ文字のキー（title・subtitle）には、文字の大きい順・上から順に当てる。書き換えは、新しい文字を先頭に入れてから元の文字を消す
 * （元の文字の書式（大きさ・太さ・色・字体）を引き継ぐ）。空なら文字を消す
 */
function mtgTemplateTextRequests_(page, template) {
  if (!page || !template) return [];
  const texts = MTG_OPTIONS.templateTexts;
  const same = (a, b) => a === b || (/^X{5,}$/.test(a) && /^X{5,}$/.test(b));
  const shapes = (page.pageElements || []).filter(el => el.shape && el.shape.text).map(el => {
    const runs = (el.shape.text.textElements || []).filter(t => t.textRun);
    const style = runs.length && runs[0].textRun.style || {};
    return {id: el.objectId, text: runs.map(t => t.textRun.content).join('').replace(/\n$/, ''),
      size: style.fontSize ? style.fontSize.magnitude : 0, y: el.transform && el.transform.translateY || 0};
  }).sort((a, b) => (b.size - a.size) || (a.y - b.y));
  const taken = {};
  const requests = [];
  Object.keys(texts).forEach(key => {
    if (!texts[key] || !(key in template)) return;
    const shape = shapes.find(sh => !taken[sh.id] && same(sh.text, texts[key]));
    if (!shape) return;
    taken[shape.id] = true;
    const value = String(template[key] == null ? '' : template[key]);
    if (!value) {
      requests.push({deleteText: {objectId: shape.id, textRange: {type: 'ALL'}}});
      return;
    }
    requests.push({insertText: {objectId: shape.id, insertionIndex: 0, text: value}},
      {deleteText: {objectId: shape.id, textRange: {type: 'FIXED_RANGE', startIndex: value.length, endIndex: value.length + shape.text.length}}});
  });
  return requests;
}

/**
 * ひな形が無いときの表紙・扉の文字：ひな形と同じ位置（ひな形の 960 × 540pt を 0.75 倍）に、見出し・その下の文字・日付・名前を白で描く。
 * 背景（緑のグラデーションと大きな「RX」）は、ページの背景に画像で付ける（mtgCoverBackground_。付けるまでは緑一色。mtgSlideRequests_）
 */
function mtgTemplateFallback_(template) {
  const c = MTG_OPTIONS.colors;
  const S = MTG_SLIDE;
  const line = (y, h, text, size, bold) => mtgBox_(S.x, y, S.w, h, mtgWrap_(text, S.w, size, 1)[0], {size, bold: !!bold, color: c.white, align: 'CENTER'});
  // 文字の中心を、ひな形の枠の文字の中心にそろえる（見出し 194.8・その下 240.3・日付 340.1・名前 357.4pt）
  return [
    line(171.8, 46, template.title, 30, true),
    line(227.3, 26, template.subtitle, 15, true),
    line(330.1, 20, template.date, 12),
    line(347.4, 20, template.name, 12)
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
  // 扉（Appendix）は、いただいたひな形の2枚目（グレー）と同じ色で。紫（3枚目と同じ）は、統合版の表紙を紫にするとき（from・to・ink・RX の位置を置き換える）。
  // グレー：背景の 50%（#7F7F7F）→ 85%（#D9D9D9）、RX は 65%（#A6A6A6）。紫：#7030A0 → #A02B93、RX は #421C5E で、33pt（960pt の幅で）左
  divider: {from: '#7F7F7F', to: '#D9D9D9', ink: '#A6A6A6'},
  combined: {from: '#7030A0', to: '#A02B93', ink: '#421C5E', shift: -44.3},
  // RX の字形（Liberation Sans Bold。SIL Open Font License）：文字の大きさ size（px）・左端 x・ベースライン baseline（画像の px）。
  // 輪郭はフォントの単位（upm = 文字の大きさ）で、点は「x,y」、点の間は空白、輪郭の間は「|」。X は R の送り幅（1479）だけ右にずらしてある
  glyph: {size: 1142.67, x: -109.4, baseline: 756.2, upm: 2048},
  contours: '1105,0 778,535 432,535 432,0 137,0 137,1409 841,1409 902,1407 960,1402 1014,1394 1064,1382 1111,1367 1154,1348 1194,1326 1230,1300 1266,1268 1297,1231 1322,1190 1342,1146 1356,1097 1364,1045 1367,989 1365,941 1358,895 1346,851 1330,810 1309,770 1283,734 1253,700 1220,670 1184,644 1145,623 1102,605 1056,592 1437,0|1070,977 1065,1031 1049,1076 1022,1114 985,1143 937,1163 879,1176 810,1180 432,1180 432,764 818,764 865,766 908,773 945,784 978,800 1006,820 1034,851 1054,887 1066,929|2517,0 2163,561 1809,0 1497,0 1985,741 1538,1409 1850,1409 2163,911 2476,1409 2786,1409 2358,741 2827,0'
};

/**
 * 表紙・扉のページ（pages = [{id, tone: mtgCoverTone_}]）の背景に、MTG_COVER の画像を付ける
 * （SlidesApp。Slides API ではファイルの画像を背景にできないため）。付けられなければ、一色の背景のまま続ける（資料は作る）
 */
function mtgCoverBackground_(presentationId, pages) {
  if (!pages.length) return;
  try {
    const deck = SlidesApp.openById(presentationId);
    pages.forEach(({id, tone}) => {
      const page = deck.getSlideById(id);
      if (page) page.getBackground().setPictureFill(mtgCoverImage_(tone));
    });
    deck.saveAndClose();
  } catch (error) {
    console.warn('表紙・扉の背景の画像を付けられませんでした（一色のままにします）：' + (error && error.message || error));
  }
}

/** 表紙・扉の背景の画像（PNG の Blob。tone は mtgCoverTone_）。描くのは実行ごと・色ごとに1回（いくつかの資料を作るときも同じ画像を使う） */
function mtgCoverImage_(tone) {
  const key = tone === 'divider' || tone === 'combined' ? tone : 'cover';
  if (!MTG_COVER_CACHE[key]) MTG_COVER_CACHE[key] = Array.from(mtgCoverPng_(key), b => (b > 127 ? b - 256 : b));
  return Utilities.newBlob(MTG_COVER_CACHE[key], 'image/png', 'meeting-' + key + '.png');
}

/* 描いた背景の画像（実行の間だけ覚えておく。色ごと） */
const MTG_COVER_CACHE = {cover: null, divider: null, combined: null};

/**
 * MTG_COVER の画像を描いて PNG のバイト列（Uint8Array）にする。RX は輪郭を横の線で塗る（1px を縦に4本・横は面積で数えて、縁をなめらかに）。
 * tone が 'divider' なら扉の色（MTG_COVER.divider。グレー）、'combined' なら紫（RX の位置も。統合版の表紙を紫にするとき）
 */
function mtgCoverPng_(tone) {
  const C = Object.assign({}, MTG_COVER, tone === 'divider' || tone === 'combined' ? MTG_COVER[tone] : {});
  C.glyph = Object.assign({}, MTG_COVER.glyph, {x: MTG_COVER.glyph.x + (C.shift || 0)});
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

/**
 * ひな形のファイルの ID（スクリプトのプロパティ MTG_TEMPLATE_URL・MTG_OPTIONS.templateId・前の版の登録の順。ほかの設定と同じく、スクリプトのプロパティを優先する。
 * Google スライドか .pptx。無ければ空）
 */
function mtgTemplateSourceId_() {
  const o = MTG_OPTIONS;
  const url = PropertiesService.getScriptProperties().getProperty(o.templateUrlKey);
  if (url) {
    const id = mtgFileIdOf_(url);
    if (!id) throw new Error('スクリプトのプロパティ ' + o.templateUrlKey + ' の「' + url + '」から、ひな形のファイルの ID を読めません。Google スライド（か .pptx）を開いたときの URL を入れてください。');
    return id;
  }
  if (o.templateId) return mtgFileIdOf_(o.templateId);
  return PropertiesService.getDocumentProperties().getProperty(o.templateKey) || '';
}

/**
 * 資料を作るときにコピーする、ひな形の Google スライドの ID（無ければ空）。ひな形が .pptx なら、Google スライドに変換したものを使う
 * （mtgTemplateSlides_。変換は .pptx を更新したときだけ）
 */
function mtgTemplateId_() {
  const id = mtgTemplateSourceId_();
  return id ? mtgTemplateSlides_(id).getId() : '';
}

/** 「https://docs.google.com/presentation/d/ID/edit」や ID だけの文字から ID を取り出す（読めなければ空） */
function mtgFileIdOf_(text) {
  const t = String(text || '').trim();
  const m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(t) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(t) || /^([A-Za-z0-9_-]{20,})$/.exec(t);
  return m ? m[1] : '';
}

/**
 * 表紙・扉のひな形の URL を、スクリプトのプロパティ（MTG_TEMPLATE_URL）に入れる・外す（エディタから実行。URL を聞く。空欄なら外し、ひな形なしで作る）。
 * エディタの「プロジェクトの設定」→「スクリプト プロパティ」で直接入れてもよい。Google スライドでも .pptx（ドライブに置いたもの）でもよい
 */
function setMeetingTemplate() {
  const ui = mtgUi_();
  const o = MTG_OPTIONS;
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty(o.templateUrlKey);
  const response = ui.prompt(CRED_OPTIONS.menuTitle,
    '会議サマリの表紙・扉のひな形（Google スライドか .pptx）の URL を入力してください。\n空欄なら登録を外します（ひな形なしで作ります）。\n\n' +
    '今の登録（スクリプトのプロパティ ' + o.templateUrlKey + '）：' + (current || 'なし'), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const text = String(response.getResponseText() || '').trim();
  PropertiesService.getDocumentProperties().deleteProperty(o.templateKey);   // 前の版の登録は使わない
  if (!text) {
    props.deleteProperty(o.templateUrlKey);
    ui.alert(CRED_OPTIONS.menuTitle, 'ひな形の登録を外しました。会議サマリは、ひな形なしで作ります。', ui.ButtonSet.OK);
    return;
  }
  const id = mtgFileIdOf_(text);
  if (!id) throw new Error('「' + text + '」から、ひな形のファイルの ID を読めません。Google スライド（か .pptx）を開いたときのアドレス（URL）を貼り付けてください。');
  mtgRequireSlides_();
  const slides = mtgTemplateSlides_(id);
  props.setProperty(o.templateUrlKey, text);
  const message = mtgTemplateMessage_(slides, id);
  console.log(message);
  ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
}

/** 登録したひな形の知らせ（何枚目を何に使うか。.pptx なら変換した Google スライドの URL も） */
function mtgTemplateMessage_(slides, sourceId) {
  const count = (Slides.Presentations.get(slides.getId(), {fields: MTG_SLIDES_FIELDS.count}).slides || []).length;
  const P = MTG_OPTIONS.templatePages;
  const use = [[P.cover, '表紙'], [P.divider, '扉（Appendix）'], [P.combined, '統合版（全サブインダストリー）の表紙'], [P.service, 'RX Service Meeting の表紙']]
    .filter(([n]) => n && n <= count).map(([n, label]) => n + '枚目を' + label);
  return 'ひな形を登録しました（' + count + '枚。' + use.join('、') + 'に使います' +
    (count < P.divider ? '。扉はグレーの絵柄で描きます' : '') + (P.combined && P.combined <= count ? '' : '。統合版の表紙も1枚目を使います') +
    (P.service && P.service <= count ? '' : '。RX Service Meeting の表紙は紫の絵柄で描きます') + '）。次に作る会議サマリから使います。' +
    (slides.getId() !== sourceId ? '\n.pptx を Google スライドに変換して使います（.pptx を更新したら、次に作るときに変換し直します）：' + slides.getUrl() : '');
}

/**
 * ひな形のファイル（id）を Google スライドとして返す（DriveApp の File）。.pptx などなら、Google スライドに変換したもの（保存先のフォルダの
 * 「会議サマリのひな形（…から変換）」。スクリプトのプロパティ templateConvertedKey に元の ID・更新日時と一緒に覚え、.pptx を更新したら変換し直す）
 */
function mtgTemplateSlides_(id) {
  let file;
  try {
    file = DriveApp.getFileById(id);
  } catch (e) {
    throw new Error('ひな形（ID：' + id + '）を開けません。ファイルがあるか、この管理者のアカウントで開けるかを確かめてください（' + e.message + '）。');
  }
  if (file.getMimeType() === MimeType.GOOGLE_SLIDES) return file;
  if (!/presentation|powerpoint/i.test(String(file.getMimeType()))) {
    throw new Error('ひな形（' + file.getName() + '）が Google スライドでも .pptx でもありません。スクリプトのプロパティ ' + MTG_OPTIONS.templateUrlKey + ' の URL を確かめてください。');
  }
  const o = MTG_OPTIONS;
  const props = PropertiesService.getScriptProperties();
  const updated = file.getLastUpdated().getTime();
  const saved = mtgReadJson_(props, o.templateConvertedKey);
  if (saved && saved.source === id && saved.updated === updated) {
    try {
      const converted = DriveApp.getFileById(saved.id);
      if (!converted.isTrashed()) return converted;
    } catch (_) {}
  }
  const folder = mtgFolder_(SpreadsheetApp.getActiveSpreadsheet());
  const response = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/copy?supportsAllDrives=true&fields=id', {
    method: 'post',
    contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: JSON.stringify({name: '会議サマリのひな形（' + file.getName() + ' から変換）', mimeType: MimeType.GOOGLE_SLIDES, parents: [folder.getId()]}),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('ひな形（' + file.getName() + '）を Google スライドに変換できませんでした（' + code + '：' + String(response.getContentText()).slice(0, 200) + '）。' +
      'ドライブで開いて「ファイル」→「Google スライドとして保存」し、できた Google スライドの URL をスクリプトのプロパティ ' + o.templateUrlKey + ' に入れてください。');
  }
  const converted = DriveApp.getFileById(JSON.parse(response.getContentText()).id);
  if (saved && saved.id && saved.id !== converted.getId()) {
    try { DriveApp.getFileById(saved.id).setTrashed(true); } catch (_) {}   // 前に変換したもの（古い .pptx から）
  }
  props.setProperty(o.templateConvertedKey, JSON.stringify({source: id, updated, id: converted.getId()}));
  console.log('ひな形（' + file.getName() + '）を Google スライドに変換しました：' + converted.getUrl());
  return converted;
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

/** ファイル名（20261023_スーパー・GMS_RX Business Meeting用資料。MTG_OPTIONS.fileName。試作は「_試作」つき） */
function mtgFileName_(name, meeting, final) {
  return MTG_OPTIONS.fileName.replace('{date}', meeting.replace(/\//g, '')).replace('{name}', name) + (final ? '' : '_試作');
}

/**
 * スライドの並び（mtgSlides_）を、Google スライド（16:9）と PDF にしてフォルダに保存する。
 * 表紙・扉のひな形（mtgTemplateId_）があれば、ひな形をコピーした資料に作る：表紙・扉・統合版の表紙は、ひな形のページ（MTG_OPTIONS.templatePages。
 * 2回目からは複製）の決まった文字を書き換えて使う（mtgTemplatePlan_・mtgTemplateTextRequests_）。ほかのスライドは白紙のスライドを足して描き、位置・大きさ・文字はひな形のページの大きさに合わせて
 * 拡大する（MTG_SLIDE の 720 × 405pt を基準に描いているため）。ひな形が無ければ、新しい資料（720 × 405pt）に作る。
 * スライドは Slides API（拡張サービス「Google Slides API」）で作り、できたスライドを PDF に書き出す。
 * MTG_OPTIONS.keepSlides なら Google スライドもフォルダに残す（推進リーダーが書き足せる）。うまくいかなかったときはゴミ箱に移す。
 * 戻り値 {pdf, deck}（deck は残さないときは null）
 */
function mtgExportDeck_(name, slides, folder) {
  const F = MTG_SLIDES_FIELDS;
  const templateId = mtgTemplateId_();
  let file;
  let deck;
  if (templateId) {
    file = DriveApp.getFileById(templateId).makeCopy(name, folder);
    deck = Slides.Presentations.get(file.getId(), {fields: F.deck});
  } else {
    deck = Slides.Presentations.create({title: name, pageSize: {width: mtgPt_(MTG_SLIDE.width), height: mtgPt_(MTG_SLIDE.height)}});
    file = DriveApp.getFileById(deck.presentationId);
  }
  try {
    const scale = mtgPageWidth_(deck) / MTG_SLIDE.width;
    const pages = (deck.slides || []).map(sl => sl.objectId);
    // ひな形：表紙・扉・統合版の表紙に、ひな形のページ（MTG_OPTIONS.templatePages）を使う（2回目からは複製。使わないページは消す）。
    // 先に消す・複製してから読み直し、ひな形のページの決まった文字を枠ごとに書き換える
    const plan = templateId ? mtgTemplatePlan_(slides, pages) : {pageOf: {}, requests: []};
    const byId = {};
    if (templateId) {
      if (plan.requests.length) Slides.Presentations.batchUpdate({requests: plan.requests}, deck.presentationId);
      (Slides.Presentations.get(deck.presentationId, {fields: F.texts}).slides || []).forEach(sl => { byId[sl.objectId] = sl; });
    }
    let requests = [];
    let first = true;
    const flush = () => {
      if (first && !templateId) pages.forEach(id => requests.push({deleteObject: {objectId: id}}));   // 作ったときの白紙のスライド
      Slides.Presentations.batchUpdate({requests}, deck.presentationId);
      requests = [];
      first = false;
    };
    slides.forEach((slide, i) => {
      const own = 'mtg_p' + (i + 1);
      const page = plan.pageOf[i];
      const add = page ? mtgTemplateTextRequests_(byId[page], slide.template).concat(mtgSlideRequests_(slide, page, i, {scale, existing: true, idPrefix: own}))
        : mtgSlideRequests_(slide, own, i, {scale});
      if (requests.length && requests.length + add.length > MTG_OPTIONS.batchRequests) flush();
      requests = requests.concat(add);
    });
    flush();
    const extra = mtgPlaceholderRequests_(Slides.Presentations.get(deck.presentationId, {fields: F.placeholders}), slides.map((slide, i) => 'mtg_p' + (i + 1)));
    // ページへのリンク（目次）：リンク先のスライドができてから付ける（部品の ID は「mtg_p{ページ}_{部品の番号}」）
    const pageId = i => plan.pageOf[i] || 'mtg_p' + (i + 1);
    slides.forEach((slide, i) => slide.elements.forEach((el, n) => {
      if (el.link == null || !slides[el.link] || el.text == null || el.text === '') return;
      // リンクを付けると青の下線になるので、字の色を戻し、下線を付けない
      extra.push({updateTextStyle: {objectId: 'mtg_p' + (i + 1) + '_' + n, textRange: {type: 'ALL'}, style: {link: {pageObjectId: pageId(el.link)},
        underline: false, foregroundColor: {opaqueColor: {rgbColor: mtgRgb_(el.color || MTG_OPTIONS.colors.body)}}}, fields: 'link,underline,foregroundColor'}});
    }));
    // ひな形のページ（複製は元のページのすぐ次に入る）と足したスライドを、1枚目から順に決まった位置に動かす（前のスライドの位置は決まっているので、順に動かせばそろう）
    if (templateId) slides.forEach((slide, i) => extra.push({updateSlidesPosition: {slideObjectIds: [pageId(i)], insertionIndex: i}}));
    if (extra.length) Slides.Presentations.batchUpdate({requests: extra}, deck.presentationId);
    // ひな形のページを使わない表紙・扉の背景を、いただいたひな形と同じ絵柄（グラデーションと RX）の画像にする（表紙は緑・扉はグレー。mtgCoverTone_）
    mtgCoverBackground_(deck.presentationId, slides.map((slide, i) => ((slide.cover || slide.divider) && !plan.pageOf[i]
      ? {id: 'mtg_p' + (i + 1), tone: mtgCoverTone_(slide)} : null)).filter(Boolean));
    if (!templateId && MTG_OPTIONS.keepSlides) file.moveTo(folder);   // 残さないときは動かさずにゴミ箱へ（呼び出しを減らす）
    const blob = file.getAs('application/pdf').setName(name + '.pdf');
    const pdf = folder.createFile(slides.some(slide => slide.elements.some(el => el.link != null)) ? mtgPdfFitLinks_(blob) : blob);
    const pptx = MTG_OPTIONS.exportPptx ? mtgExportPptx_(file.getId(), name, folder) : {};
    if (pptx.error) console.warn(name + '：' + pptx.error + '（PDF は作りました）');
    mtgTrashOld_(folder, [name, name + '.pdf', name + '.pptx'], [file.getId(), pdf.getId()].concat(pptx.file ? [pptx.file.getId()] : []));
    if (!MTG_OPTIONS.keepSlides) file.setTrashed(true);
    return {pdf, pptx: pptx.file || null, pptxError: pptx.error || '', deck: MTG_OPTIONS.keepSlides ? file : null};
  } catch (e) {
    try { file.setTrashed(true); } catch (_) {}
    throw e;
  }
}

/*
 * Slides API で資料を読むときの項目（fields）。読む項目を絞る：統合版（130ページほど）を丸ごと読むと、応答が大きすぎて失敗する
 * （「レスポンス コード: 413。メッセージ: response too large」）。deck：ページの大きさとページの ID、texts：ひな形のページの文字の枠
 * （mtgTemplateTextRequests_）、placeholders：レイアウトの枠（mtgPlaceholderRequests_）、count：ページの数
 */
const MTG_SLIDES_FIELDS = {
  deck: 'presentationId,pageSize,slides/objectId',
  texts: 'slides(objectId,pageElements(objectId,transform/translateY,shape/text/textElements/textRun(content,style/fontSize)))',
  placeholders: 'slides(objectId,pageElements(objectId,shape/placeholder/type))',
  count: 'slides/objectId'
};

/**
 * PDF のページへのリンク（目次）の飛び先を、ページ全体を出す形（/Fit）にする。Google スライドの PDF は「/XYZ null null 0」（縦の位置を変えない）で書き出し、
 * 見る画面によってはページの頭に移らないため。同じ長さの文字に置き換える（後ろを空白で埋める）ので、ほかの場所（相互参照表の位置）は変わらない。
 * 置き換えるものが無ければ blob をそのまま返す
 */
function mtgPdfFitLinks_(blob) {
  const bytes = blob.getBytes();
  let text = '';
  for (let i = 0; i < bytes.length; i += 8192) text += String.fromCharCode.apply(null, bytes.slice(i, i + 8192).map(b => b & 255));
  let count = 0;
  const fixed = text.replace(/\/XYZ\s+null\s+null\s+0(?=[\s\]])/g, m => {
    count++;
    return '/Fit' + ' '.repeat(m.length - 4);
  });
  if (!count) return blob;
  const out = new Array(fixed.length);
  for (let i = 0; i < fixed.length; i++) {
    const code = fixed.charCodeAt(i);
    out[i] = code > 127 ? code - 256 : code;
  }
  return Utilities.newBlob(out, 'application/pdf', blob.getName());
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
/**
 * 作った Google スライド（id）を PowerPoint（name.pptx）に書き出して folder に置く（Drive API の files.export）。
 * {file} か {error: 理由}（止めない。PDF は作ってあるので、PowerPoint だけ作れなかったと知らせる）
 */
function mtgExportPptx_(id, name, folder) {
  const type = MimeType.MICROSOFT_POWERPOINT;
  try {
    const response = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=' + encodeURIComponent(type), {
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    if (code !== 200) return {error: 'PowerPoint に書き出せませんでした（' + code + '：' + String(response.getContentText()).slice(0, 200) + '）'};
    return {file: folder.createFile(response.getBlob().setName(name + '.pptx'))};
  } catch (error) {
    return {error: 'PowerPoint に書き出せませんでした（' + (error && error.message || error) + '）'};
  }
}

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

/** 時点コピーの保存先の設定（スクリプトのプロパティ MTG_COPY_FOLDER → MTG_OPTIONS.copyFolderId） */
function mtgCopyFolderSetting_() {
  return mtgOutputSetting_([MTG_OPTIONS.folderProps.copy], MTG_OPTIONS.copyFolderId, 'MeetingSummary.gs の MTG_OPTIONS.copyFolderId');
}

/** 保存先の設定（スクリプトのプロパティ MTG_SAVE_FOLDER → MTG_OPTIONS.folderId） */
function mtgSaveFolderSetting_() {
  return mtgOutputSetting_([MTG_OPTIONS.folderProps.save], MTG_OPTIONS.folderId, 'MeetingSummary.gs の MTG_OPTIONS.folderId');
}

/**
 * 保存先のフォルダ（スクリプトのプロパティ MTG_SAVE_FOLDER、無ければ MTG_OPTIONS.folderId。どちらも空なら、
 * このスプレッドシートと同じフォルダの中に folderName を作り、次からも使う）。
 */
function mtgFolder_(ss) {
  const setting = mtgSaveFolderSetting_();
  const id = mtgOutputFolderId_(setting, '保存先');
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (error) {
      throw new Error('保存先のフォルダ（ID：' + id + '）を開けません。実行するアカウントが、そのフォルダに書き込めるか確かめてください' +
        '（' + setting.where + '）。');
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
