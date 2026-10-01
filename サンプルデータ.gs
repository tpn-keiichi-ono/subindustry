/**
 * サンプルの履歴データ（画面の確認・説明用）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs と同じプロジェクトに置くファイル。
 *
 * - メニュー「（管理者）サンプル履歴を選択行に追加」：選んだ行の得意先に、次のサンプルを追加する
 *     クレデンシャル 5件・オファリング 5件（クレデンシャルに紐づける）… クレデンシャル履歴_記録
 *     変更の記録 5件（過去2週間ほど。列はその行の追跡範囲から）… 変更履歴_差分・変更時点スナップショット
 *   追跡シートのセルの値は変えない（記録だけを足す）。
 * - サンプルは見分けられるようにしてある：名称の先頭に「【サンプル】」、登録元・登録者・編集者は「サンプルデータ」、
 *   変更の記録のイベントIDは「SMP-」で始まる。
 * - メニュー「（管理者）サンプル履歴をすべて削除」：上の印が付いた行だけを削除する
 *   （記録は追記のみが原則だが、サンプルの行に限って削除する）。
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
  note: 'サンプルデータです。メニューの「（管理者）サンプル履歴をすべて削除」で消せます。',
  // 変更の記録：何日前・時刻・変更後の値（列は追跡範囲の見出しのある列から順に使う）
  changes: [
    {daysAgo: 13, time: '10:15', value: '初回訪問'},
    {daysAgo: 9, time: '14:30', value: '課題ヒアリング'},
    {daysAgo: 6, time: '09:05', value: '提案書提出'},
    {daysAgo: 3, time: '16:40', value: '見積提出'},
    {daysAgo: 1, time: '11:20', value: '受注見込み'}
  ],
  maxChangeColumns: 3            // 変更の記録に使う列の数
};

/* ---------------- メニューから ---------------- */

/** 選んだ行の得意先に、サンプルの履歴を追加する。 */
function addSampleHistory() {
  const ui = SpreadsheetApp.getUi();
  const title = 'サンプル履歴の追加';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targets = smpSelectedTargets_(SpreadsheetApp.getActiveSheet());
    const names = targets.map(t => t.customer);
    const answer = ui.alert(title,
      '次の得意先に、サンプルの履歴を追加します（' + names.length + '件）。\n' + names.join('、') + '\n\n' +
      '・クレデンシャル ' + SMP_OPTIONS.credentials.length + '件\n' +
      '・オファリング ' + SMP_OPTIONS.offerings.length + '件\n' +
      '・変更の記録 ' + SMP_OPTIONS.changes.length + '件\n\n' +
      'シートのセルの値は変わりません。サンプルは「（管理者）サンプル履歴をすべて削除」で消せます。追加しますか？',
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
      'サンプルを追加しました：クレデンシャル ' + total.credentials + '件・オファリング ' + total.offerings +
      '件・変更の記録 ' + total.changes + '件。\n履歴サイドバーは「再読み込み」で反映されます。',
      ui.ButtonSet.OK);
  } catch (error) {
    console.error(error.stack || String(error));
    ui.alert(title, error.message, ui.ButtonSet.OK);
  }
}

/** サンプルの印が付いた行だけを、記録のシートから削除する。 */
function removeSampleHistory() {
  const ui = SpreadsheetApp.getUi();
  const title = 'サンプル履歴の削除';
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

/** 選んでいる行から、得意先ごとに1つずつ対象を返す。 */
function smpSelectedTargets_(sheet) {
  const range = sheet.getActiveRange();
  if (!range) throw new Error('対象行のセルを選択してください。');
  const targets = [];
  const seen = new Set();
  for (let row = range.getRow(); row <= range.getLastRow(); row++) {
    const target = credResolveTarget_(sheet, row);
    const key = credNormalize_(target.customer);
    if (seen.has(key)) continue;
    seen.add(key);
    targets.push(target);
    if (targets.length > SMP_OPTIONS.maxRows) {
      throw new Error('一度に追加できるのは ' + SMP_OPTIONS.maxRows + '件の得意先までです。選ぶ行を減らしてください。');
    }
  }
  return targets;
}

/** ロック取得中に呼ぶこと。戻り値 {credentials, offerings, changes} */
function smpAddForTarget_(ss, target, now) {
  const creds = smpAddCredentials_(ss, target, now);
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

  // 変更履歴_差分・変更時点スナップショット とも、B列がイベントID
  const prefix = SMP_OPTIONS.eventPrefix + '-';
  [[DIFF_OPTIONS.logSheet, 'changes'], [DIFF_OPTIONS.rowSnapshotSheet, 'snapshots']].forEach(([name, key]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const rows = [];
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().forEach((v, i) => {
      if (String(v[0]).indexOf(prefix) === 0) rows.push(i + 2);
    });
    smpDeleteRows_(sheet, rows);
    result[key] = rows.length;
  });
  return result;
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
