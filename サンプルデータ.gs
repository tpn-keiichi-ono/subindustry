/**
 * サンプルの履歴データ（画面の確認・説明用）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs と同じプロジェクトに置くファイル。
 *
 * - addSampleHistory()（エディタから実行）：選んだ行の得意先に、次のサンプルを追加する
 *     クレデンシャル 5件・オファリング 5件（クレデンシャルに紐づける）… クレデンシャル履歴_記録
 *     変更の記録 5件（過去2週間ほど。列はその行の追跡範囲から）… 変更履歴_差分・変更時点スナップショット
 *   追跡シートのセルの値は変えない（記録だけを足す）。
 * - サンプルは見分けられるようにしてある：名称の先頭に「【サンプル】」、登録元・登録者・編集者は「サンプルデータ」、
 *   変更の記録のイベントIDは「SMP-」で始まる。
 * - removeSampleHistory()（エディタから実行）：上の印が付いた行だけを削除する
 *   （記録は追記のみが原則だが、サンプルの行に限って削除する）。
 * - addBusinessMtgDummyHistory()（エディタから実行。シートを選ぶ）：会議サマリ（MeetingSummary.gs）の確認用に、追跡シートの案件へ
 *   ダミーの変更の記録（変更履歴_差分・変更時点スナップショット）を足す。シートのセルの値は変えず、今の値になるまでの記録
 *   （前回の資料・基準の時点の値と、そのあとの変更）を足すので、資料の「前回から変化のある案件」（前進・後退・延期・増額・減額・活動更新・新規・削除）
 *   と停滞に出る。サンプル（SMP-）は会議サマリが数えないので、イベントIDは「DMY-」にする（会議サマリは数え、資料の出典の行に【ダミーデータを含む】と出す）。
 * - removeBusinessMtgDummyHistory()（エディタから実行）：イベントIDが「DMY-」の行だけを削除する（サンプルと同じく、ダミーの行に限って削除する）。
 * - addServiceMtgSampleData()（エディタから実行）：RX Service MTG の事前資料（ServiceMeeting.gs）の確認用に、サービスリクエスト・
 *   サービス案タスク管理のシートの最後の行の下に、サンプルのリクエスト・タスクを書く（SMP_OPTIONS.service）。リクエストの先頭に「【サンプル】」を付ける
 *   （タスクの「サービス案」にも入る）。得意先・案件名は、リクエストを読むシートの案件（進行中で期待値の大きいもの）を使い、案件が無いシートでは架空の名前にする。
 *   追跡シート・取り込み元のシートには書かない。
 * - removeServiceMtgSampleData()（エディタから実行）：リクエストが「【サンプル】」で始まる行と、そのタスクの行だけを削除する。
 */

const SMP_OPTIONS = {
  marker: 'サンプルデータ',      // 登録元・登録者・編集者に入れる印（削除するときの目印）
  eventPrefix: 'SMP',            // 変更の記録のイベントIDの接頭辞（削除するときの目印）
  namePrefix: '【サンプル】',     // 名称・変更後の値の先頭に付ける
  kindSuffix: ' / サンプル',      // 変更履歴_差分 の「種類」に付ける
  maxRows: 20,                   // 一度に追加できる行（得意先）の数
  owner: 'サンプル 担当',
  ownerDept: '第一営業部',
  // クレデンシャル：done = 実施日（何日前）、plan = 予定日（何日後。負の数は過ぎた予定）
  credentials: [
    {name: 'ブランド刷新の実績紹介', person: '山田 太郎', title: '営業部 部長', done: 60},
    {name: '販促DXの導入事例紹介', person: '佐藤 花子', title: '販促企画部 課長', done: 45},
    {name: '店舗什器リニューアルの事例紹介', person: '鈴木 一郎', title: '店舗開発部 部長', done: 30},
    {name: 'EC立ち上げ支援の実績紹介', person: '高橋 美咲', title: 'EC推進室 室長', done: 14},
    {name: '物流改善の事例紹介', person: '田中 健', title: '物流部 課長', plan: 10}
  ],
  // オファリング：上のクレデンシャルに順に紐づける
  offerings: [
    {name: '新ブランド立ち上げのご提案', person: '山田 太郎', title: '営業部 部長', done: 50},
    {name: '販促DXの試験導入のご提案', person: '佐藤 花子', title: '販促企画部 課長', plan: -3},
    {name: '店舗什器入れ替えのご提案', person: '鈴木 一郎', title: '店舗開発部 部長', plan: 7},
    {name: 'EC運用代行のご提案', person: '高橋 美咲', title: 'EC推進室 室長', done: 7},
    {name: '物流センター見直しのご提案', person: '田中 健', title: '物流部 課長', plan: 20}
  ],
  note: 'サンプルデータです。エディタで removeSampleHistory を実行すると消せます。',
  // 変更の記録：何日前・時刻・変更後の値（列は追跡範囲の見出しのある列から順に使う）
  changes: [
    {daysAgo: 13, time: '10:15', value: '初回訪問'},
    {daysAgo: 9, time: '14:30', value: '課題ヒアリング'},
    {daysAgo: 6, time: '09:05', value: '提案書提出'},
    {daysAgo: 3, time: '16:40', value: '見積提出'},
    {daysAgo: 1, time: '11:20', value: '受注見込み'}
  ],
  maxChangeColumns: 3,           // 変更の記録に使う列の数
  // 会議サマリの確認用のダミーの変更の記録（addBusinessMtgDummyHistory()）
  dummy: {
    eventPrefix: 'DMY',          // イベントIDの接頭辞（会議サマリは数える。removeBusinessMtgDummyHistory() で消す目印）
    marker: 'ダミーデータ',       // 編集者メールアドレスの列に入れる印
    kindSuffix: ' / ダミー',      // 種類に付ける（行の追加・削除の記録には付けない。会議サマリが種類で見分けるため）
    activity: '（ダミー）前回までの活動状況',   // 前回の時点の活動状況
    // 案件に付ける区分（上から、付けられる案件に1つずつ。案件が足りなければ下は付けない）。stall は停滞（ステータスを最後に変えた日を stallDays 日前にする）、
    // stallActivity は停滞＋前回以降の活動更新。forward 前進・backward 後退・slipped 延期・up 増額・down 減額・activity 活動更新・new 新規
    roles: ['forward', 'stall', 'slipped', 'activity', 'new', 'stallActivity', 'up', 'backward', 'down'],
    stallDays: {stall: 100, stallActivity: 70},
    // 削除した案件（シートに無い架空の案件）の、削除前の行の内容（ステータスは進行中のいちばん前のもの、期待値は想定売上 × その率、受注月は来月）
    removed: {customer: 'ダミー商事', project: '（ダミー）削除した案件', amount: 20, owner: 'ダミー 担当'}
  },
  // RX Service MTG の確認用のサンプル（addServiceMtgSampleData()）。sheet はリクエストを読むシート（その案件を使う。無ければ順に別のシート。
  // 案件が無ければ customer・project の架空の名前）。decision は判断（空なら未判断）。リクエストの先頭に namePrefix を付けて書く
  service: {
    requests: [
      {sheet: 'スーパー・GMS', customer: 'サンプルスーパー', project: '会員アプリ刷新', request: '会員アプリの購買データから、優良顧客の離反の兆しを早めに知りたい',
        decision: 'サービス化検討', service: '購買データ分析', department: 'DX事業本部', owner: '山本',
        feedback: '既存の ID-POS 分析基盤で対応できる見込み。データ連携の範囲を確かめたい'},
      {sheet: 'コンビニ', customer: 'サンプルマート', project: '売場改善', request: '店舗ごとの売れ筋を、メーカーと共有する仕組みがほしい',
        decision: 'サービス化検討', service: '購買データ分析', department: 'DX事業本部', owner: '山本',
        feedback: '共有の範囲（個店・エリア）で、必要な加工が変わる'},
      {sheet: '化粧品・トイレタリー', customer: 'サンプル化粧品', project: 'EC連携', request: 'EC と店舗の購買をつないで、顧客ごとに提案したい',
        decision: 'サービス化検討', service: '購買データ分析', department: 'DX事業本部', owner: '山本',
        feedback: '会員 ID で EC と店舗の購買をつなげるかを、先方の情報システム部門と確かめる'},
      {sheet: 'スーパー・GMS', customer: 'サンプルスーパー', project: '店舗DX', request: '紙の POP の差し替えの手間を減らしたい（電子棚札）',
        decision: 'サービス化検討', service: 'デジタル販促（電子棚札・サイネージ）', department: '情報コミュニケーション事業本部', owner: '田中',
        feedback: '電源・通信の工事の要否で、費用が大きく変わる'},
      {sheet: '百貨店', customer: 'サンプル百貨店', project: '売場サイネージ', request: '売場のサイネージの内容を、本部からまとめて配信したい',
        decision: 'サービス化検討', service: 'デジタル販促（電子棚札・サイネージ）', department: '情報コミュニケーション事業本部', owner: '田中',
        feedback: '既存の配信の仕組みで、売場ごとの出し分けまでできる見込み'},
      {sheet: '生協', customer: 'サンプル生協', project: '宅配チラシ', request: '宅配のチラシを、アプリでも見られるようにしたい',
        decision: 'サービス化検討', service: 'デジタル販促（電子棚札・サイネージ）', department: '', owner: '',   // 主管本部・担当者未設定の確認用
        feedback: '紙とアプリの両方に出す運用の手間を、先方と確かめる'},
      {sheet: '飲料・食品', customer: 'サンプル飲料', project: '店頭販促', request: '店頭什器を季節ごとに入れ替えたいが、買い切りは負担が大きい',
        decision: 'サービス化検討', service: '店頭什器のサブスクリプション', department: '生活・産業事業本部', owner: '中村',
        feedback: '回収・再利用の物流の設計が課題'},
      {sheet: '住設', customer: 'サンプル住設', project: 'ショールーム改装', request: 'ショールームの展示什器を、短いサイクルで変えたい',
        decision: 'サービス化検討', service: '店頭什器のサブスクリプション', department: '生活・産業事業本部', owner: '中村',
        feedback: 'ショールームは入れ替えの頻度が高く、月額にしやすい'},
      {sheet: '自動車販売', customer: 'サンプル自動車販売', project: '来店予約', request: '来店予約から商談までを、アプリで完結させたい',
        decision: 'サービス化検討', service: '来店予約アプリ', department: 'DX事業本部', owner: '小林', feedback: '既存のアプリの型で作れるか確かめる'},
      {sheet: 'コンビニ', customer: 'サンプルマート', project: '会員施策', request: 'カード会員のアンケートを、店頭の端末で取りたい',
        decision: 'サービス化検討', service: '', department: '', owner: '',   // サービス案未設定の確認用
        feedback: '店頭端末の改修が要るため、先方の端末の更新時期を確かめる'},
      {sheet: '地方銀行', customer: 'サンプル銀行', project: 'アプリ利用促進', request: 'アプリの利用を店頭で案内する販促物を作りたい',
        decision: '', service: '', department: '', owner: '', feedback: ''},
      {sheet: '電鉄・不動産', customer: 'サンプル電鉄', project: '駅ナカ活性化', request: '駅ナカ店舗の購買データを、テナントと共有したい',
        decision: '', service: '', department: '', owner: '', feedback: ''},
      {sheet: 'エネルギー', customer: 'サンプルエネルギー', project: '検針票', request: '検針票の紙をなくしたいが、高齢の顧客への案内が課題',
        decision: '棄却', service: '', department: '', owner: '', feedback: 'ほかの部署の通知サービスで対応済みのため、ここでは扱わない'}
    ],
    // 案件が無い（架空の名前の）リクエストの営業の列
    sales: {accountDept: '第一営業本部', accountOwner: 'サンプル 担当', bxDept: 'BX推進部', bxOwner: 'サンプル BX'},
    // タスク（request は上の requests の番号（0 から））。owner が空のタスクは、資料の確認事項（タスクの担当者未入力）の確認用。
    // progress は進捗状況、planStart・planEnd・actualStart・actualEnd は開始日・終了日（予定・実績）の今日からの日数（負は前。null は空欄）
    tasks: [
      {request: 0, task: '既存の ID-POS 分析基盤で、離反の兆しの指標を試算する', department: 'DX事業本部', owner: '山本', status: '対応中',
        progress: '60%：指標案3つのうち2つを試算済み', planStart: -21, planEnd: 14, actualStart: -18, actualEnd: null,
        note: '1社のデータで試算中。11月上旬に結果を出す'},
      {request: 1, task: 'メーカーへの共有の範囲（個店・エリア）の要件を整理する', department: 'DX事業本部', owner: '山本', status: '未着手',
        progress: '0%：先方の窓口を確認中', planStart: 7, planEnd: 35, actualStart: null, actualEnd: null,
        note: '先方のマーチャンダイジング部門と打ち合わせを設定する'},
      {request: 0, task: '料金の考え方（分析の回数・データ量）の案を作る', department: '事業企画部', owner: '佐々木', status: '未着手',
        progress: '0%：試算の結果を待って着手', planStart: 14, planEnd: 42, actualStart: null, actualEnd: null,
        note: '他社の分析サービスの料金を調べておく'},
      {request: 3, task: '電子棚札の工事の要否を、店舗の型ごとに整理する', department: '情報コミュニケーション事業本部', owner: '田中', status: '完了',
        progress: '100%：3つの型に分類済み', planStart: -45, planEnd: -14, actualStart: -42, actualEnd: -10,
        note: '型ごとの概算費用は別紙にまとめた'},
      {request: 4, task: 'サイネージの配信を、既存の CMS で動かせるか確かめる', department: '情報コミュニケーション事業本部', owner: '田中', status: '対応中',
        progress: '40%：検証環境を準備済み', planStart: -14, planEnd: 21, actualStart: -10, actualEnd: null,
        note: '売場ごとの出し分けの動作を確かめる'},
      {request: 3, task: '導入事例の資料をまとめる', department: '情報コミュニケーション事業本部', owner: '', status: '未着手',
        progress: '0%：担当者を決める', planStart: 7, planEnd: 28, actualStart: null, actualEnd: null,
        note: '社内の導入事例を3件集める'},
      {request: 6, task: '什器の回収・再利用の物流の費用を試算する', department: '生活・産業事業本部', owner: '中村', status: '未着手',
        progress: '0%：物流会社に見積を依頼する', planStart: 10, planEnd: 40, actualStart: null, actualEnd: null,
        note: '回収の頻度は季節ごと（年4回）で試算する'}
    ],
    // タスク管理の、TASK_OPTIONS.columns に無い列（シートに足した列）に書くサンプル：見出し → 上のタスクの項目。見出しが無ければ書かない。
    // date は日付の列（今日からの日数を日付にして書く）
    taskExtraColumns: [
      {label: '進捗状況', key: 'progress'},
      {label: '開始日（予定）', key: 'planStart', date: true},
      {label: '終了日（予定）', key: 'planEnd', date: true},
      {label: '開始日（実績）', key: 'actualStart', date: true},
      {label: '終了日（実績）', key: 'actualEnd', date: true}
    ]
  }
};

/* 区分の名前（確かめの画面に出す） */
const SMP_DUMMY_LABELS = {forward: '前進', backward: '後退', slipped: '延期', up: '増額', down: '減額', activity: '活動更新', new: '新規',
  stall: '停滞', stallActivity: '停滞＋活動更新'};

/* ---------------- エディタから実行 ---------------- */

/** 選んだ行の得意先に、サンプルの履歴を追加する。 */
function addSampleHistory() {
  const ui = SpreadsheetApp.getUi();
  const title = CRED_OPTIONS.menuTitle;   // 確認画面のタイトルはメニューの名前にそろえる
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targets = smpSelectedTargets_(SpreadsheetApp.getActiveSheet());
    const names = targets.map(t => t.customer);
    const answer = ui.alert(title,
      '次の得意先に、サンプルの履歴を追加します（' + names.length + '件）。\n' + names.join('、') + '\n\n' +
      (credEnabled_() ? '・クレデンシャル ' + SMP_OPTIONS.credentials.length + '件\n' +
        '・オファリング ' + SMP_OPTIONS.offerings.length + '件\n' : '') +
      '・変更の記録 ' + SMP_OPTIONS.changes.length + '件\n\n' +
      'シートのセルの値は変わりません。サンプルは、エディタで removeSampleHistory を実行すると消せます。\n' +
      '※ サンプルは会議サマリには出ません（会議サマリの確認用のダミーデータは addBusinessMtgDummyHistory()）。追加しますか？',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;

    const total = smpWithLock_(() => {
      const now = new Date();
      const sum = {credentials: 0, offerings: 0, changes: 0};
      targets.forEach(target => {
        const result = smpAddForTarget_(ss, target, now);
        Object.keys(sum).forEach(k => { sum[k] += result[k]; });
      });
      return sum;
    });
    ui.alert(title,
      'サンプルを追加しました：' + (credEnabled_() ? 'クレデンシャル ' + total.credentials + '件・オファリング ' + total.offerings + '件・' : '') +
      '変更の記録 ' + total.changes + '件。\n履歴サイドバーは「再読み込み」で反映されます。',
      ui.ButtonSet.OK);
  } catch (error) {
    console.error(error.stack || String(error));
    ui.alert(title, error.message, ui.ButtonSet.OK);
  }
}

/** サンプルの印が付いた行だけを、記録のシートから削除する。 */
function removeSampleHistory() {
  const ui = SpreadsheetApp.getUi();
  const title = CRED_OPTIONS.menuTitle;
  try {
    const answer = ui.alert(title,
      'サンプルの履歴をすべて削除します。\n' +
      '（登録元が「' + SMP_OPTIONS.marker + '」のクレデンシャル・オファリングと、イベントIDが「' +
      SMP_OPTIONS.eventPrefix + '-」の変更の記録）\n\nほかの履歴は消えません。削除しますか？',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;

    const result = smpWithLock_(() => smpRemoveAll_(SpreadsheetApp.getActiveSpreadsheet()));
    ui.alert(title,
      '削除しました：クレデンシャル・オファリング ' + result.credentials + '件・変更の記録 ' + result.changes +
      '件（スナップショット ' + result.snapshots + '件）。',
      ui.ButtonSet.OK);
  } catch (error) {
    console.error(error.stack || String(error));
    ui.alert(title, error.message, ui.ButtonSet.OK);
  }
}

/* ---------------- 追加 ---------------- */

/**
 * 選んでいる行から、得意先ごとに1つずつ対象を返す。得意先が空の行（案件の無い行）は飛ばし、1行も無ければ、選ぶ行と
 * 会議サマリ用のダミーデータの関数を案内して止める。
 */
function smpSelectedTargets_(sheet) {
  const range = sheet.getActiveRange();
  if (!range) throw new Error('対象行のセルを選択してください。');
  const rule = DIFF_RULES[sheet.getName()];
  const targets = [];
  const seen = new Set();
  for (let row = range.getRow(); row <= range.getLastRow(); row++) {
    if (rule && row > rule.headerRow && row <= sheet.getMaxRows()) {
      const customerColumn = credColumnByHeader_(sheet, rule, CRED_OPTIONS.customerHeader);
      if (!sheet.getRange(row, customerColumn).getDisplayValue().trim()) continue;   // 得意先が空の行は飛ばす
    }
    const target = credResolveTarget_(sheet, row);
    const key = credNormalize_(target.customer);
    if (seen.has(key)) continue;
    seen.add(key);
    targets.push(target);
    if (targets.length > SMP_OPTIONS.maxRows) {
      throw new Error('一度に追加できるのは ' + SMP_OPTIONS.maxRows + '件の得意先までです。選ぶ行を減らしてください。');
    }
  }
  if (!targets.length) {
    throw new Error('選んだ行（' + range.getRow() + (range.getLastRow() > range.getRow() ? '〜' + range.getLastRow() : '') + '行目）に、' +
      '「' + CRED_OPTIONS.customerHeader + '」が入力された行がありません。案件のある行を選んでから実行してください。\n' +
      '会議サマリの確認用のダミーの変更の記録を足すときは、addBusinessMtgDummyHistory() を実行してください（行を選ぶ必要はありません）。');
  }
  return targets;
}

/** ロック取得中に呼ぶこと。戻り値 {credentials, offerings, changes} */
function smpAddForTarget_(ss, target, now) {
  // クレデンシャル・オファリング履歴を使っていない間は、サンプルのクレデンシャルは足さない（CRED_OPTIONS.enabled）
  const creds = credEnabled_() ? smpAddCredentials_(ss, target, now) : {credentials: 0, offerings: 0};
  return {credentials: creds.credentials, offerings: creds.offerings, changes: smpAddChanges_(ss, target, now)};
}

function smpAddCredentials_(ss, target, now) {
  const L = CRED_OPTIONS.linking;
  const S = CRED_OPTIONS.schedule;
  const {log, cols} = credLogSheet_(ss, false);
  const timezone = ss.getSpreadsheetTimeZone();
  const day = offset => Utilities.formatDate(new Date(now.getTime() + offset * 86400000), timezone, 'yyyy/MM/dd');
  const fieldsOf = (item, kind) => {
    const values = {
      kind, title: item.title, person: item.person, name: SMP_OPTIONS.namePrefix + item.name,
      note: SMP_OPTIONS.note, owner: SMP_OPTIONS.owner, ownerDept: SMP_OPTIONS.ownerDept
    };
    values[S.planKey] = item.plan != null ? day(item.plan) : '';
    values[S.doneKey] = item.done != null ? day(-item.done) : '';
    return credValidate_(values, null);   // 入力欄と同じ決まりでそろえる
  };
  const entryOf = (fields, link) => ({
    id: Utilities.getUuid(), at: now, customer: target.customer, fields, link: link || '',
    editor: SMP_OPTIONS.marker, source: SMP_OPTIONS.marker
  });

  const parents = SMP_OPTIONS.credentials.map(item => entryOf(fieldsOf(item, L.parentKind)));
  if (parents.length) credAppendLog_(log, cols, parents);
  const children = SMP_OPTIONS.offerings.map((item, i) =>
    entryOf(fieldsOf(item, L.childKind), parents.length ? parents[i % parents.length].id : ''));
  if (children.length) credAppendLog_(log, cols, children);
  return {credentials: parents.length, offerings: children.length};
}

/** 変更の記録（変更履歴_差分 と 変更時点スナップショット）を足す。戻り値は件数。 */
function smpAddChanges_(ss, target, now) {
  const {sheet, rule, row, customerColumn} = target;
  const stampColumn = diffStampColumn_(sheet, rule);
  const columns = [];
  diffBlocks_(sheet, rule, sheet.getRange(row, 1, 1, sheet.getMaxColumns()), stampColumn, false)
    .forEach(([, col, , count]) => {
      diffColumnNames_(sheet, rule, col, count).forEach((name, i) => {
        // 得意先の列は変えない（記録を得意先で探すため）。見出しの無い列も使わない
        if (col + i === customerColumn || /列（見出しなし）$/.test(name)) return;
        columns.push({column: col + i, name});
      });
    });
  const use = columns.slice(0, SMP_OPTIONS.maxChangeColumns);
  if (!use.length) return 0;

  const timezone = ss.getSpreadsheetTimeZone();
  const schema = diffSnapshotSchema_(sheet, rule);
  const budget = {remaining: DIFF_OPTIONS.maxLcsCellsPerEdit};
  const lastValue = {};
  SMP_OPTIONS.changes.forEach((change, i) => {
    const target = use[i % use.length];
    const date = Utilities.formatDate(new Date(now.getTime() - change.daysAgo * 86400000), timezone, 'yyyy/MM/dd');
    const at = Utilities.parseDate(date + ' ' + change.time, timezone, 'yyyy/MM/dd HH:mm');
    const before = lastValue[target.column] || '';
    const after = SMP_OPTIONS.namePrefix + change.value;
    lastValue[target.column] = after;
    const eventId = diffCreateEventId_(at, timezone, SMP_OPTIONS.eventPrefix);
    const cell = diffA1_(row, target.column);
    diffAppend_(ss, [{
      eventId,
      editorEmail: SMP_OPTIONS.marker,
      sheet: sheet.getName(),
      cell,
      columnName: target.name,
      kind: (before ? '文字列' : '空欄') + ' -> 文字列' + SMP_OPTIONS.kindSuffix,
      before,
      after,
      parts: diffCharacters_(before, after, budget).parts
    }], at);
    diffAppendRowSnapshots_(ss, sheet, rule, [row], new Map([[row, new Set([cell])]]),
      eventId, SMP_OPTIONS.marker, at, schema);
  });
  return SMP_OPTIONS.changes.length;
}

/* ---------------- 削除 ---------------- */

/** ロック取得中に呼ぶこと。戻り値 {credentials, changes, snapshots} */
function smpRemoveAll_(ss) {
  const result = {credentials: 0, changes: 0, snapshots: 0};

  const credLog = ss.getSheetByName(CRED_OPTIONS.logSheet);
  if (credLog && credLog.getLastRow() >= 2) {
    const {cols} = credLogSheet_(ss, false);
    const values = credLog.getRange(2, 1, credLog.getLastRow() - 1, credWidth_(cols)).getDisplayValues();
    const rows = [];
    values.forEach((v, i) => {
      if (v[cols.source - 1] === SMP_OPTIONS.marker && v[cols.editor - 1] === SMP_OPTIONS.marker) rows.push(i + 2);
    });
    smpDeleteRows_(credLog, rows);
    result.credentials = rows.length;
  }

  const events = smpRemoveEvents_(ss, SMP_OPTIONS.eventPrefix + '-', '');
  result.changes = events.changes;
  result.snapshots = events.snapshots;
  return result;
}

/**
 * 変更履歴_差分・変更時点スナップショットから、イベントIDが prefix で始まる行を削除する（sheetName があれば、そのシートの記録だけ）。
 * ロック取得中に呼ぶこと。戻り値 {changes, snapshots}
 */
function smpRemoveEvents_(ss, prefix, sheetName) {
  const result = {changes: 0, snapshots: 0};
  smpEventRows_(ss, prefix, sheetName).forEach(({sheet, key, rows}) => {
    smpDeleteRows_(sheet, rows);
    result[key] = rows.length;
  });
  return result;
}

/** イベントIDが prefix で始まる行の数 {changes, snapshots}（sheetName があれば、そのシートの記録だけ）。読むだけ */
function smpCountEvents_(ss, prefix, sheetName) {
  const result = {changes: 0, snapshots: 0};
  smpEventRows_(ss, prefix, sheetName).forEach(({key, rows}) => { result[key] = rows.length; });
  return result;
}

/** 変更履歴_差分・変更時点スナップショット とも、B列がイベントID・D列がシート。[{sheet, key, rows: 行番号}] */
function smpEventRows_(ss, prefix, sheetName) {
  const out = [];
  [[DIFF_OPTIONS.logSheet, 'changes'], [DIFF_OPTIONS.rowSnapshotSheet, 'snapshots']].forEach(([name, key]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const rows = [];
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getDisplayValues().forEach((v, i) => {
      if (String(v[0]).indexOf(prefix) === 0 && (!sheetName || String(v[2]).trim() === sheetName)) rows.push(i + 2);
    });
    out.push({sheet, key, rows});
  });
  return out;
}

/** 行を下から削除する（続いている行はまとめて）。固定行以外がすべて無くなる場合は、先に空の行を足す。 */
function smpDeleteRows_(sheet, rows) {
  if (!rows.length) return;
  if (sheet.getMaxRows() - rows.length <= sheet.getFrozenRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }
  const sorted = rows.slice().sort((a, b) => b - a);
  for (let i = 0; i < sorted.length;) {
    let j = i;
    while (j + 1 < sorted.length && sorted[j + 1] === sorted[j] - 1) j++;
    sheet.deleteRows(sorted[j], j - i + 1);
    i = j + 1;
  }
}

/* ---------------- 会議サマリの確認用のダミーの変更の記録 ---------------- */

/**
 * 会議サマリの確認用に、追跡シートの案件へダミーの変更の記録を足す（エディタから実行。シートを選ぶ）。
 * シートのセルの値は変えない。進行中の案件に区分（SMP_OPTIONS.dummy.roles）を1つずつ付け、今の値になるまでの記録を足す：
 * - 前回の時点の値：前回の資料（または基準）を作った時刻の1秒前に、前回の値のスナップショットと活動状況の記録
 * - 前回より後の変更：前回の値 → 今の値の記録（今の行のスナップショット）。新規は行の追加、削除は架空の案件の行の削除
 * - 停滞：ステータスを stallDays 日前に変えた記録（それより新しいステータスの変更・案件ができた記録が無い案件だけ。あると停滞にならないため）
 * 同じシートに前に足したダミーは、消してから入れ直す。消すときは removeBusinessMtgDummyHistory()
 */
function addBusinessMtgDummyHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const title = CRED_OPTIONS.menuTitle;
  const D = SMP_OPTIONS.dummy;
  const names = mtgAskSheets_(ss, ui, '会議サマリの確認用に、ダミーの変更の記録を足します。');
  if (!names) return;
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const plans = names.map(name => smpDummyPlan_(ss, name, now));
  const lines = plans.map(p => {
    const roles = p.deals.map(d => d.customer + '：' + SMP_DUMMY_LABELS[d.role]).join('、');
    return '・' + p.name + '（' + p.since.label + ' ' + Utilities.formatDate(new Date(p.since.at), tz, 'M/d HH:mm') + ' 以降の変化として）\n　' +
      (roles ? roles + '、' : '') + '削除（架空の案件「' + D.removed.customer + '」）' +
      (p.skippedStall ? '\n　停滞は付けません（' + p.skippedStall + '）' : '') +
      (p.since.copy ? '\n　※ ' + p.since.label + 'の「' + MTG_OPTIONS.copyName + '」があるため、前回からの変化はコピーと比べます。' +
        'ダミーの記録は前回からの変化には出ません（停滞の日数などにだけ使われます）' : '') +
      (p.existing ? '\n　前に足したダミーの記録 ' + p.existing + '件は消して入れ直します' : '');
  });
  const answer = ui.alert(title, '会議サマリの確認用に、次のダミーの変更の記録を足します。\n\n' + lines.join('\n') + '\n\n' +
    'シートのセルの値は変わりません。記録のイベントIDは「' + D.eventPrefix + '-」で始まり、資料の出典の行に【ダミーデータを含む】と出ます。\n' +
    '消すときは、エディタで removeBusinessMtgDummyHistory() を実行してください。足しますか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const results = smpWithLock_(() => plans.map(plan => smpAddDummy_(ss, plan, now)));
  const message = 'ダミーの変更の記録を足しました。\n' +
    results.map(r => '・' + r.name + '：変更の記録 ' + r.changes + '件・スナップショット ' + r.snapshots + '件').join('\n') +
    '\n\npreviewBusinessMtgSummaries() で資料を作って確かめてください（会議日は、記録の時刻より後の日にする）。';
  console.log(message);
  ui.alert(title, message, ui.ButtonSet.OK);
}

/** ダミーの変更の記録（イベントIDが DMY-）を、すべてのシートの分、記録のシートから削除する（エディタから実行）。 */
function removeBusinessMtgDummyHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const title = CRED_OPTIONS.menuTitle;
  const prefix = SMP_OPTIONS.dummy.eventPrefix + '-';
  const count = smpCountEvents_(ss, prefix, '');
  if (!count.changes && !count.snapshots) {
    ui.alert(title, 'ダミーの変更の記録（イベントIDが「' + prefix + '」で始まるもの）はありません。', ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert(title, 'ダミーの変更の記録を削除します（イベントIDが「' + prefix + '」で始まる、変更の記録 ' + count.changes +
    '件・スナップショット ' + count.snapshots + '件）。\nほかの記録は消えません。削除しますか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const result = smpWithLock_(() => smpRemoveEvents_(ss, prefix, ''));
  const message = 'ダミーの変更の記録を削除しました：変更の記録 ' + result.changes + '件・スナップショット ' + result.snapshots + '件。';
  console.log(message);
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * 1つのシートに足すダミーの記録の計画（読むだけ）。
 * {name, sheet, rule, table, since: {at, label}, deals: [案件 + role], prev: 区分ごとの前回の値を作る関数, skippedStall: 停滞を付けない理由, existing}
 */
function smpDummyPlan_(ss, name, now) {
  const o = MTG_OPTIONS;
  const D = SMP_OPTIONS.dummy;
  const sheet = ss.getSheetByName(name);
  const rule = DIFF_RULES[name];
  const table = mtgReadDeals_(sheet, rule);
  const statusTable = mtgReadStatusTable_(ss);
  const rank = mtgStatusRank_(table.statusOrder, statusTable);
  // 進行中のステータス（進み具合の順）。前進・後退の前回のステータスは、1つ前・1つあとのもの
  const choices = (table.statusOrder.length ? table.statusOrder : (statusTable || []).map(st => st.status))
    .filter(st => mtgStatusKind_(st) === 'active' && rank(st) >= 0).sort((a, b) => rank(a) - rank(b));
  const step = (status, dir) => {
    const i = choices.findIndex(st => credNormalize_(st) === credNormalize_(status));
    return i >= 0 && i + dir >= 0 && i + dir < choices.length ? choices[i + dir] : null;
  };
  const since = smpDummySince_(name, now);
  // 停滞を付けられるか：ステータスの変更・案件ができた記録（ダミーを除く）があれば、それより前の日付にしても停滞にならない
  const windowStart = new Date(now.getTime() - o.lookbackDays * 86400000);
  const notDummy = id => id.indexOf(D.eventPrefix + '-') !== 0;
  const snapshots = new Map();
  mtgReadSnapshots_(ss, windowStart).forEach((snap, key) => { if (notDummy(snap.eventId)) snapshots.set(key, snap); });
  const history = mtgHistory_(name, {since: new Date(since.at), log: mtgReadLog_(ss, windowStart).filter(e => notDummy(e.eventId)), snapshots});
  const labels = mtgLabels_();
  const moved = d => (history.events.get(d.key) || []).some(e => e.label === labels.status || e.born);
  // 前回より後の本物の変更。前回の時点の値は、その変更前の値（ダミーの前回の時点のスナップショットに使い、本物の変化を隠さない）
  const recentReal = d => (history.events.get(d.key) || []).filter(e => e.at > since.at).sort((a, b) => a.at - b.at);
  const baseOf = d => {
    const out = {};
    ['status', 'amount', 'weighted', 'orderMonth', 'activity'].forEach(key => {
      const e = recentReal(d).find(x => x.label === labels[key]);
      if (e) out[key] = e.before;
    });
    return out;
  };
  const ok = {
    forward: d => !!step(d.status, -1),
    backward: d => !!step(d.status, 1),
    slipped: d => table.has.orderMonth && !!d.orderMonth,
    up: d => table.has.amount && d.amount > 0,
    down: d => table.has.amount && d.amount > 0,
    activity: d => table.has.activity && !!d.activity,
    new: () => true,
    stall: d => !moved(d),
    stallActivity: d => table.has.activity && !moved(d)
  };
  // 前回より後に本物の変更が無い案件から選ぶ（本物の変化と重ねない）。その中は期待値の大きい順
  const active = table.deals.filter(d => d.kind === 'active')
    .sort((a, b) => (recentReal(a).length ? 1 : 0) - (recentReal(b).length ? 1 : 0) ||
      ((b.weighted || 0) - (a.weighted || 0)) || ((b.amount || 0) - (a.amount || 0)) || (a.row - b.row));
  const used = new Set();
  const deals = [];
  D.roles.forEach(role => {
    const deal = active.find(d => !used.has(d.row) && ok[role](d));
    if (!deal) return;
    used.add(deal.row);
    deals.push(Object.assign({}, deal, {role}));
  });
  const stallRoles = D.roles.filter(role => /^stall/.test(role));
  const skippedStall = deals.some(d => /^stall/.test(d.role)) || !stallRoles.length ? ''
    : 'どの案件にも、ステータスの変更か案件ができた記録があり、それより前の日付にしても停滞にならないため';
  const existing = smpCountEvents_(ss, D.eventPrefix + '-', name).changes;
  return {name, sheet, rule, table, statusTable, step, lowest: choices[0] || '', since, deals, skippedStall, existing, baseOf};
}

/**
 * ダミーの変化の起点：会議サマリの前回（いちばん新しい会議の資料・基準の時刻）。無ければ前の版の作成日時、それも無ければ firstDays 日前。{at, label}
 */
function smpDummySince_(name, now) {
  const o = MTG_OPTIONS;
  const props = PropertiesService.getDocumentProperties();
  const list = mtgMeetings_(props, name).filter(m => m.cutoff <= now.getTime());
  if (list.length) {
    const last = list.reduce((a, b) => (b.cutoff > a.cutoff ? b : a));
    // copy：その時点の会議時点のコピーがある（会議サマリの前回からの変化はコピーと比べるので、ダミーの記録は出ない）
    return {at: last.cutoff, label: last.baseline ? '基準' : '前回の資料', copy: !!last.copy};
  }
  const legacy = Number(props.getProperty(o.lastKeyPrefix + name)) || 0;
  if (legacy) return {at: legacy, label: '前回の作成'};
  return {at: now.getTime() - o.firstDays * 86400000, label: '直近' + o.firstDays + '日間の始まり'};
}

/** 計画（smpDummyPlan_）どおりにダミーの記録を足す。ロック取得中に呼ぶこと。戻り値 {name, changes, snapshots} */
function smpAddDummy_(ss, plan, now) {
  const o = MTG_OPTIONS;
  const D = SMP_OPTIONS.dummy;
  const {sheet, rule, table} = plan;
  smpRemoveEvents_(ss, D.eventPrefix + '-', plan.name);
  const tz = ss.getSpreadsheetTimeZone();
  const columnOf = key => table.headers.findIndex(h => credNormalize_(h) === credNormalize_(o.columns[key])) + 1;   // 無ければ 0
  const names = diffColumnNames_(sheet, rule, 1, table.headers.length);
  const schema = diffSnapshotSchema_(sheet, rule);
  const number = n => (n == null ? '' : String(Math.round(n * 10) / 10));
  const ratio = status => mtgStatusRatio_(plan.statusTable, status);
  // 月を delta か月ずらした表示（シートの表示の形のまま：「2026/9」→「2026/8」、「2026/09」→「2026/08」、「2026年9月」→「2026年8月」）
  const shiftMonth = (text, delta) => {
    const m = mtgMonth_(text);
    if (!m) return '';
    const index = mtgMonthIndex_(m) + delta;
    const y = Math.floor((index - 1) / 12);
    const month = index - y * 12;
    if (/年/.test(text)) return y + '年' + month + '月';
    return y + '/' + (/^\s*\d{4}\s*[\/\-.]\s*0\d/.test(String(text).normalize('NFKC')) ? String(month).padStart(2, '0') : month);
  };
  // 変更の記録の値と種類（差分追跡と同じ形：日時は DIFF_OPTIONS.dateFormat の文字、数は数の文字、空欄は「空欄」（記録のシートには「（空欄）」と出る））
  const token = value => {
    if (value === '' || value == null) return {type: '空欄', text: ''};
    if (typeof value.getTime === 'function') return {type: '日時', text: Utilities.formatDate(value, tz, DIFF_OPTIONS.dateFormat)};
    if (typeof value === 'number') return {type: '数値', text: String(value)};
    if (typeof value === 'boolean') return {type: '真偽値', text: String(value)};
    return {type: '文字列', text: String(value)};
  };
  const valueOf = (row, key) => sheet.getRange(row, columnOf(key)).getValue();
  // 日付を delta か月ずらす（スクリプトのタイムゾーンで。日は 28日まで）
  const addMonths = (date, delta) => {
    const p = Utilities.formatDate(date, tz, 'yyyy-M-d-HH-mm-ss').split('-').map(Number);
    const index = p[0] * 12 + (p[1] - 1) + delta;
    const pad = n => String(n).padStart(2, '0');
    return Utilities.parseDate(Math.floor(index / 12) + '/' + pad(index % 12 + 1) + '/' + pad(Math.min(p[2], 28)) + ' ' +
      pad(p[3]) + ':' + pad(p[4]) + ':' + pad(p[5]), tz, 'yyyy/MM/dd HH:mm:ss');
  };
  // 前回の値の、記録に書く値：受注月は今の値が日付なら1か月前の日付、想定売上・期待値は数（スナップショットには表示の形 prev[key] を書く）
  const typedPrev = (row, key, display) => {
    const cur = valueOf(row, key);
    if (key === 'orderMonth' && cur && typeof cur.getTime === 'function') return addMonths(cur, -1);
    if ((key === 'amount' || key === 'weighted') && display !== '' && !isNaN(Number(display))) return Number(display);
    return display;
  };

  // 区分ごとの前回の値 {見出しのキー: 表示の値}（今の値と違う項目だけ）
  const prevOf = d => {
    const prev = {};
    const status = d.role === 'forward' ? plan.step(d.status, -1) : d.role === 'backward' ? plan.step(d.status, 1) : null;
    if (status) prev.status = status;
    if (d.role === 'slipped') prev.orderMonth = shiftMonth(d.orderMonthText, -1);
    let amount = d.amount;
    if (d.role === 'up' || d.role === 'down') {
      const diff = Math.max(1, Math.round(d.amount * 0.2));
      amount = d.role === 'up' ? d.amount - diff : d.amount + diff;
      prev.amount = number(amount);
    }
    // 期待値：想定売上 × 前回のステータスの率（率が分からなければ、想定売上の比で）
    if (table.has.weighted && d.weighted != null && (prev.status || prev.amount)) {
      const r = ratio(prev.status || d.status);
      const w = r != null ? amount * r : d.weighted * (d.amount ? amount / d.amount : 1);
      if (Math.abs(w - d.weighted) > 1e-9) prev.weighted = number(w);
    }
    if (table.has.activity && d.role !== 'new' && !/^stall$/.test(d.role)) prev.activity = D.activity;
    return prev;
  };
  // 前回より前の日付のスナップショットの値：前回の時点の本物の値（前回より後の本物の変更の、変更前の値）に、ダミーの前回の値を重ねる。
  // 期待値の列が数式なら、ステータス・想定売上から計算し直す
  const pastOf = (d, prev) => {
    const out = Object.assign({}, plan.baseOf(d), prev);
    if (table.has.weighted && table.weightedFormula && prev.weighted == null && (out.status != null || out.amount != null)) {
      const r = ratio(out.status != null ? out.status : d.status);
      const a = out.amount != null ? mtgNumber_(out.amount) : d.amount;
      if (r != null && a != null) out.weighted = number(a * r);
    }
    return out;
  };

  const changes = [];     // [{at, eventId, records}]
  const snapshots = [];   // [{at, eventId, row, cells, overrides}]
  const event = (at, records, row, overrides) => {
    const eventId = diffCreateEventId_(new Date(at), tz, D.eventPrefix);
    const list = records.map(r => Object.assign({eventId, editorEmail: D.marker, sheet: sheet.getName()}, r));
    changes.push({at, records: list});
    if (row) snapshots.push({at, eventId, row, cells: list.map(r => r.cell).join(', '), overrides: overrides || {}});
  };
  const edit = (row, key, before, after) => {
    const b = token(before);
    const a = token(after);
    return {cell: diffA1_(row, columnOf(key)), columnName: names[columnOf(key) - 1], kind: b.type + ' -> ' + a.type + D.kindSuffix,
      before: b.text, after: a.text};
  };
  const keys = ['status', 'amount', 'weighted', 'orderMonth', 'activity'];

  // 前回より後の変更の時刻：前回の1分後から今の1分前まで、等しい間隔で
  const recent = plan.deals.filter(d => !/^stall$/.test(d.role)).length + 1;
  const lo = plan.since.at + Math.min(60000, (now.getTime() - plan.since.at) / 4);
  const hi = now.getTime() - Math.min(60000, (now.getTime() - plan.since.at) / 4);
  let k = 0;
  const nextAt = () => lo + (hi - lo) * (++k) / (recent + 1);

  plan.deals.forEach(d => {
    const row = d.row;
    if (d.role === 'new') {
      event(nextAt(), [{cell: row + '行', columnName: '（行全体）', kind: DIFF_ROW_OPTIONS.kindInsert, before: '',
        after: smpRowSummary_(sheet, rule, row, token)}], row);
      return;
    }
    const prev = prevOf(d);
    if (/^stall/.test(d.role)) {
      // ステータスを stallDays 日前に変えた記録（停滞の日数の起点）
      event(now.getTime() - D.stallDays[d.role] * 86400000, [edit(row, 'status', plan.step(d.status, -1) || '', valueOf(row, 'status'))], row,
        pastOf(d, prev));
      if (d.role === 'stall') return;
    }
    // 前回の時点（前回の1秒前。前回より前でいちばん新しいスナップショットになる）：前回の値の行。
    // 記録は、そのときの活動状況の書き込み（列が無ければ、最初に変わる項目。ステータスの記録にしないのは、停滞の日数を変えないため）
    const first = prev.activity != null ? 'activity' : keys.find(key => prev[key] != null && columnOf(key));
    event(plan.since.at - 1000, [edit(row, first, '', typedPrev(row, first, prev[first]))], row, pastOf(d, prev));
    // 前回より後：前回の値 → 今の値（期待値の列が数式なら、期待値の記録は無い。スナップショットに出る）
    const records = keys.filter(key => prev[key] != null && columnOf(key) && !(key === 'weighted' && table.weightedFormula))
      .map(key => edit(row, key, typedPrev(row, key, prev[key]), valueOf(row, key)));
    event(nextAt(), records, row);
  });

  // 削除：シートに無い架空の案件の行を消した記録（差分追跡の行の削除の記録と同じ形。列の順に「見出し: 値」）
  const R = D.removed;
  const status = plan.lowest;
  const r = ratio(status);
  const next = addMonths(Utilities.parseDate(Utilities.formatDate(now, tz, 'yyyy/MM') + '/01 00:00:00', tz, 'yyyy/MM/dd HH:mm:ss'), 1);   // 受注月は来月の1日
  const removedText = [['customer', R.customer], ['project', R.project], ['status', status], ['amount', R.amount],
    ['weighted', r != null ? Math.round(R.amount * r * 10) / 10 : ''], ['orderMonth', next], ['owner', R.owner]]
    .filter(([key, value]) => value !== '' && columnOf(key)).sort((a, b) => columnOf(a[0]) - columnOf(b[0]))
    .map(([key, value]) => names[columnOf(key) - 1] + ': ' + token(value).text).join('\n');
  event(nextAt(), [{cell: (sheet.getLastRow() + 1) + '行（削除前の行番号）', columnName: '（行全体）', kind: DIFF_ROW_OPTIONS.kindDelete,
    before: removedText, after: ''}], 0);

  changes.forEach(c => diffAppend_(ss, c.records, new Date(c.at)));
  smpAppendSnapshots_(ss, sheet, rule, schema, snapshots, D.marker);
  return {name: plan.name, changes: changes.reduce((sum, c) => sum + c.records.length, 0), snapshots: snapshots.length};
}

/**
 * 行の追加の記録の、行の内容（差分追跡の行の追加の記録と同じ形：追跡範囲の列（最終更新日時の列を除く）の「見出し: 値」を改行でつなぐ。
 * 値は token（日時は DIFF_OPTIONS.dateFormat の文字、数は数の文字）。空欄の列は書かない）
 */
function smpRowSummary_(sheet, rule, row, token) {
  const stamp = diffStampColumn_(sheet, rule);
  const lines = [];
  diffBlocks_(sheet, rule, sheet.getRange(row, 1, 1, sheet.getMaxColumns()), stamp, false).forEach(([, col, , count]) => {
    const values = sheet.getRange(row, col, 1, count).getValues()[0];
    diffColumnNames_(sheet, rule, col, count).forEach((name, i) => {
      const value = token(values[i]);
      if (value.text !== '') lines.push(name + ': ' + value.text);
    });
  });
  return lines.join('\n');
}

/**
 * 変更時点スナップショットに行を足す（差分追跡と同じ形。行の内容は、今の行の値に overrides（見出しのキー → 値）を重ねたもの）。
 * items = [{at, eventId, row, cells, overrides}]
 */
function smpAppendSnapshots_(ss, sheet, rule, schema, items, editor) {
  if (!items.length) return;
  const {sheet: snap, customerColumn, contentColumn} = diffRowSnapshotSheet_(ss);
  const start = snap.getLastRow() + 1;
  diffEnsureGrid_(snap, start + items.length - 1, Math.max(customerColumn, contentColumn));
  snap.getRange(start, 1, items.length, ROW_SNAPSHOT_FIXED_HEADERS.length)
    .setValues(items.map(it => [new Date(it.at), it.eventId, editor, sheet.getName(), it.row, it.cells]));
  snap.getRange(start, 1, items.length, 1).setNumberFormat(DIFF_OPTIONS.dateFormat);
  const customerKey = diffSnapshotCustomerHeader_();
  const skipped = diffSnapshotSkippedHeaders_();
  const labels = Object.keys(MTG_OPTIONS.columns).map(key => ({key, label: credNormalize_(MTG_OPTIONS.columns[key])}));
  const tail = items.map(it => {
    const values = sheet.getRange(it.row, schema.firstColumn, 1, schema.columnCount).getDisplayValues()[0];
    const pairs = [];
    let customer = '';
    schema.items.forEach((item, i) => {
      const hit = labels.find(l => l.label === credNormalize_(item.key));
      const value = String(hit && it.overrides[hit.key] != null ? it.overrides[hit.key] : values[i] == null ? '' : values[i]);
      if (item.key === customerKey) customer = value;
      if (value !== '' && !skipped.has(item.key)) pairs.push([item.key, value]);
    });
    return [diffSnapshotText_(customer), diffSnapshotText_(diffSnapshotContent_(pairs))];
  });
  if (contentColumn === customerColumn + 1) {
    snap.getRange(start, customerColumn, items.length, 2).setRichTextValues(tail);
  } else {
    snap.getRange(start, customerColumn, items.length, 1).setRichTextValues(tail.map(t => [t[0]]));
    snap.getRange(start, contentColumn, items.length, 1).setRichTextValues(tail.map(t => [t[1]]));
  }
  snap.getRange(start, 1, items.length, ROW_SNAPSHOT_FIXED_HEADERS.length).setVerticalAlignment('top');
}

/* ---------------- RX Service MTG のサンプル（サービスリクエスト・タスク） ---------------- */

/**
 * RX Service MTG の事前資料の確認用に、サンプルのリクエストとタスクを書く（エディタから実行）。
 * サービスリクエスト・サービス案タスク管理のシートの、最後の行の下に書く（リクエストには ID を振り、タスクの「サービス案」のプルダウンの候補も合わせる）。
 * すでにサンプルがあれば書かない（removeServiceMtgSampleData() で消してから）
 */
function addServiceMtgSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUiOrNull_();
  const title = SVC_OPTIONS.menuTitle;
  const sheets = smpServiceSheets_(ss);
  const existing = smpServiceSampleRows_(sheets);
  if (existing.requests.length || existing.tasks.length) {
    const text = 'サンプルのリクエスト ' + existing.requests.length + '件・タスク ' + existing.tasks.length + '件が、すでにあります。' +
      '入れ直すときは、先に removeServiceMtgSampleData() で消してください。';
    if (ui) ui.alert(title, text, ui.ButtonSet.OK);
    console.log(text);
    return;
  }
  const S = SMP_OPTIONS.service;
  const plan = smpServicePlan_(ss);
  if (ui) {
    const answer = ui.alert(title, 'RX Service MTG の事前資料の確認用に、サンプルを書きます。\n\n' +
      '・「' + SVC_OPTIONS.requestSheet + '」：リクエスト ' + S.requests.length + '件（リクエストの先頭に「' + SMP_OPTIONS.namePrefix + '」。' +
      '得意先・案件名は、シートの案件 ' + plan.filter(p => p.deal).length + '件と架空の名前 ' + plan.filter(p => !p.deal).length + '件）\n' +
      '・「' + TASK_OPTIONS.sheet + '」：タスク ' + S.tasks.length + '件\n\n' +
      'どちらも最後の行の下に書きます（今の行は変えません。リクエストを読むシートには書きません）。\n' +
      (plan.some(p => !p.deal) ? '架空の名前のリクエストは、資料の Appendix ③ に「元の案件が見つからない」と出ます。\n' : '') +
      '消すときは、エディタで removeServiceMtgSampleData() を実行してください。書きますか？', ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }
  const result = taskWithLock_(() => smpAddServiceSample_(ss, sheets, plan), SVC_OPTIONS.lockWaitMs, 'addServiceMtgSampleData()');
  const message = 'サンプルを書きました：リクエスト ' + result.requests + '件（' + result.ids[0] + '〜' + result.ids[result.ids.length - 1] + '）・タスク ' +
    result.tasks + '件' + (result.extras.length ? '（「' + result.extras.join('」「') + '」の列にも）' : '') + '。' +
    (result.skipped.length ? '\n「' + TASK_OPTIONS.sheet + '」の「' + result.skipped.join('」「') + '」の列には、サンプルを入れていません（SMP_OPTIONS.service.taskExtraColumns に足すと入ります）。' : '') +
    '\npreviewServiceMtgSummaries() で資料を作って確かめてください（資料に【サンプルデータを含む】と出ます）。';
  console.log(message);
  if (ui) ui.alert(title, message, ui.ButtonSet.OK);
}

/** サンプルのリクエスト（リクエストが「【サンプル】」で始まる行）と、そのタスクの行を削除する（エディタから実行）。ほかの行は消さない */
function removeServiceMtgSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUiOrNull_();
  const title = SVC_OPTIONS.menuTitle;
  const sheets = smpServiceSheets_(ss);
  const found = smpServiceSampleRows_(sheets);
  if (!found.requests.length && !found.tasks.length) {
    const text = 'サンプルのリクエスト・タスク（「' + SMP_OPTIONS.namePrefix + '」の付いたもの）はありません。';
    if (ui) ui.alert(title, text, ui.ButtonSet.OK);
    console.log(text);
    return;
  }
  if (ui) {
    const answer = ui.alert(title, 'サンプルを削除します：「' + SVC_OPTIONS.requestSheet + '」のリクエスト ' + found.requests.length + '行・「' +
      TASK_OPTIONS.sheet + '」のタスク ' + found.tasks.length + '行（「' + SMP_OPTIONS.namePrefix + '」の付いた行だけ）。\nほかの行は消えません。削除しますか？',
    ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }
  const result = taskWithLock_(() => {
    const rows = smpServiceSampleRows_(sheets);   // ロックの中で読み直す
    smpDeleteRows_(sheets.task, rows.tasks);
    smpDeleteRows_(sheets.request, rows.requests);
    const warning = svcRefreshTaskChoices_(ss);
    return {requests: rows.requests.length, tasks: rows.tasks.length, warning};
  }, SVC_OPTIONS.lockWaitMs, 'removeServiceMtgSampleData()');
  const message = 'サンプルを削除しました：リクエスト ' + result.requests + '行・タスク ' + result.tasks + '行。' + (result.warning ? '\n' + result.warning : '');
  console.log(message);
  if (ui) ui.alert(title, message, ui.ButtonSet.OK);
}

/** サービスリクエスト・タスク管理のシート {request, task}（無ければ、用意する関数を知らせて止める） */
function smpServiceSheets_(ss) {
  const request = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!request) throw new Error('「' + SVC_OPTIONS.requestSheet + '」シートがありません。先に setupRequestSheet() を実行してください。');
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  if (!task) throw new Error('「' + TASK_OPTIONS.sheet + '」シートがありません。先に setupTaskSheet() を実行してください。');
  return {request, task};
}

/** サンプルの行 {requests: 行番号, tasks: 行番号}（リクエストが「【サンプル】」で始まる行・サービス案に「【サンプル】」を含むタスクの行） */
function smpServiceSampleRows_(sheets) {
  const marker = SMP_OPTIONS.namePrefix;
  const requests = svcRegistered_(sheets.request).rows.filter(r => r.values.request.indexOf(marker) === 0).map(r => r.row);
  const tasks = [];
  const headerRow = TASK_OPTIONS.headerRow;
  const count = sheets.task.getLastRow() - headerRow;
  const col = taskColumnMap_(sheets.task, TASK_OPTIONS.columns, headerRow).service;
  if (count > 0 && col) {
    sheets.task.getRange(headerRow + 1, col, count, 1).getDisplayValues().forEach((line, i) => {
      if (String(line[0]).indexOf(marker) >= 0) tasks.push(headerRow + 1 + i);
    });
  }
  return {requests, tasks};
}

/**
 * サンプルのリクエストごとに使う案件（読むだけ）。[{spec, sheet: シート名, deal: 案件（無ければ null）}]。
 * spec.sheet（無ければリクエストを読むシートを順に）の、進行中で期待値の大きい案件から、まだ使っていないもの（同じ得意先・案件名のリクエストは同じ案件）
 */
function smpServicePlan_(ss) {
  const sources = svcSourceSheets_().filter(name => ss.getSheetByName(name));
  const pools = {};
  const poolOf = name => {
    if (!(name in pools)) {
      let deals = [];
      try {
        deals = mtgReadDeals_(ss.getSheetByName(name), DIFF_RULES[name]).deals.filter(d => d.customer && d.project && d.kind === 'active');
      } catch (_) {}
      pools[name] = deals.sort((a, b) => (b.weighted || 0) - (a.weighted || 0));
    }
    return pools[name];
  };
  const used = new Set();
  const byName = new Map();
  return SMP_OPTIONS.service.requests.map((spec, i) => {
    const sheet = sources.indexOf(spec.sheet) >= 0 ? spec.sheet : sources.length ? sources[i % sources.length] : spec.sheet;
    const named = spec.customer + '\u0001' + spec.project;
    if (byName.has(named)) return {spec, sheet: byName.get(named).sheet, deal: byName.get(named).deal};
    const deal = sources.indexOf(sheet) >= 0 ? poolOf(sheet).find(d => !used.has(d.key)) || null : null;
    if (deal) used.add(deal.key);
    byName.set(named, {sheet, deal});
    return {spec, sheet, deal};
  });
}

/**
 * サンプルを書く（ロックの中で呼ぶ）：リクエストを最後の行の下に書いて ID を振り、タスクの候補を合わせてから、タスクを最後の行の下に書く。
 * {requests, tasks, ids}
 */
function smpAddServiceSample_(ss, sheets, plan) {
  const S = SMP_OPTIONS.service;
  const marker = SMP_OPTIONS.namePrefix;
  const request = sheets.request;
  const cols = taskColumnMap_(request, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
  const start = svcRegistered_(request).lastRow + 1;
  const numbers = cols.id ? svcTakeIds_(plan.length, svcMaxIdInSheet_(request, cols.id)) : [];
  const ids = numbers.map(svcFormatId_);
  const width = Math.max.apply(null, Object.keys(cols).map(k => cols[k]));
  diffEnsureGrid_(request, start + plan.length - 1, width);
  const values = plan.map((p, i) => {
    const d = p.deal;
    const cell = label => (d && d.cells[credNormalize_(label)]) || '';
    const row = new Array(width).fill('');
    const put = (key, value) => { if (cols[key]) row[cols[key] - 1] = credText_(value || ''); };
    put('id', ids[i]);
    put('subIndustry', cell('サブインダストリー') || p.sheet);
    put('customer', d ? d.customer : p.spec.customer);
    put('project', d ? d.project : p.spec.project);
    // 営業の列：案件の値（案件が無い・空なら SMP_OPTIONS.service.sales）
    put('accountDept', cell('アカウント責任者部署') || S.sales.accountDept);
    put('accountOwner', (d && d.owner) || S.sales.accountOwner);
    put('bxDept', cell('BX部署') || S.sales.bxDept);
    put('bxOwner', cell('BX担当') || S.sales.bxOwner);
    put('request', marker + p.spec.request);
    put('decision', p.spec.decision || SVC_OPTIONS.defaultDecision);
    put('feedback', p.spec.feedback);
    put('service', p.spec.service);
    put('department', p.spec.department);
    put('owner', p.spec.owner);
    return row;
  });
  request.getRange(start, 1, values.length, width).setValues(values);

  // タスクの「サービス案」は、今書いたリクエストの候補（ID とリクエスト）から選ぶので、先に候補を合わせる
  svcLinkTaskServiceColumn_(ss, true);
  const task = sheets.task;
  const tcols = taskColumnMap_(task, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow);
  // シートに足した列（進捗状況・開始日・終了日など。SMP_OPTIONS.service.taskExtraColumns の見出しがある列）
  const theaders = task.getRange(TASK_OPTIONS.headerRow, 1, 1, Math.max(task.getLastColumn(), 1)).getDisplayValues()[0];
  const extras = S.taskExtraColumns.map(x => Object.assign({col: taskFindColumn_(theaders, x.label, '', false)}, x)).filter(x => x.col);
  const tlast = task.getLastRow();
  const twidth = Math.max.apply(null, Object.keys(tcols).map(k => tcols[k]).concat(extras.map(x => x.col)));
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayOf = days => new Date(today.getTime() + days * 86400000);
  diffEnsureGrid_(task, tlast + S.tasks.length, twidth);
  const tvalues = S.tasks.map(t => {
    const row = new Array(twidth).fill('');
    const put = (key, value) => { if (tcols[key]) row[tcols[key] - 1] = credText_(value || ''); };
    put('service', svcTaskChoiceLabel_(ids[t.request] || '', marker + S.requests[t.request].request));
    put('task', t.task);
    put('department', t.department);
    put('owner', t.owner);
    put('status', t.status);
    put('note', t.note);
    extras.forEach(x => {
      const value = t[x.key];
      if (value == null || value === '') return;
      row[x.col - 1] = x.date ? dayOf(Number(value)) : credText_(value);
    });
    return row;
  });
  task.getRange(tlast + 1, 1, tvalues.length, twidth).setValues(tvalues);
  // 日付の列は、日付の表示形式にそろえる（書いた行だけ）
  extras.filter(x => x.date).forEach(x => task.getRange(tlast + 1, x.col, tvalues.length, 1).setNumberFormat('yyyy/MM/dd'));
  const filled = Object.keys(tcols).map(k => tcols[k]).concat(extras.map(x => x.col));
  const skipped = theaders.filter((h, i) => String(h).trim() && filled.indexOf(i + 1) < 0);
  return {requests: values.length, tasks: tvalues.length, ids, extras: extras.map(x => x.label), skipped};
}

/* ---------------- 共通 ---------------- */

function smpWithLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CRED_OPTIONS.lockWaitMs)) {
    throw new Error('他の処理が実行中です。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}
