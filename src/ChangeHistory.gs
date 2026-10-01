/**
 * 変更履歴モーダル（行ごとの変更の確認）
 * 差分追跡スクリプト（v5）・クレデンシャル履歴と同じプロジェクトに置くファイル。
 *
 * 仕組み
 * - 追跡シートの「変更履歴」列を、各行の起動ボタン（チェックボックス）にする
 *   （チェックの検知は CredentialHistory.gs の onCredentialLauncherEdit が兼ねる）
 * - 「変更時点スナップショット」から、同じシート・同じ得意先の記録を新しい順に読み、
 *   同じイベントIDの「変更履歴_差分」から、変わった項目の変更前・変更後を取り出して表示する
 * - 行番号は行の削除・並べ替えで変わるため、「シート名＋得意先」で記録を探す
 *
 * セットアップ（トリガー所有者のアカウントで）
 * 1. 追跡シートに「変更履歴」列を追加し、DIFF_RULES の ranges から外して setupDiffTracking() を実行
 * 2. setupChangeHistoryLauncher() を実行（チェックボックスを入れ、起動トリガーを確認）
 */

const CHG_OPTIONS = {
  launcherHeader: '変更履歴',          // ボタン列の見出し
  dialogTemplate: 'ChangeHistoryDialog',
  maxEvents: 300,                      // 表示する記録の上限（新しい順）
  emptyMark: '（空欄）'                // 変更履歴_差分 で空欄を表す文字
};

/* ---------------- 起動 ---------------- */

/** CredentialHistory.gs の onCredentialLauncherEdit / credIsLauncherEdit_ から呼ばれる。 */
function chgIsLauncherEdit_(e, rule) {
  try {
    if (e.range.getNumColumns() !== 1) return false;
    return e.range.getColumn() ===
      credColumnByHeader_(e.range.getSheet(), rule, CHG_OPTIONS.launcherHeader);
  } catch (_) {
    return false;
  }
}

/** メニューから開く場合（ボタンが使えないときの代替・初回の権限承認用）。 */
function openChangeHistoryDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    if (!range) throw new Error('対象行のセルを選択してください。');
    if (range.getNumRows() > 1) throw new Error('1行だけ選択してください。');
    chgShowDialog_(sheet, range.getRow(), credActiveEmail_());
  } catch (error) {
    ui.alert('変更履歴', error.message, ui.ButtonSet.OK);
  }
}

/** options.focusEventId を渡すと、その変更の位置まで移動し、全項目を開いた状態で開く。 */
function chgShowDialog_(sheet, row, email, options) {
  const target = credResolveTarget_(sheet, row);
  const size = credDialogSize_(email);   // クレデンシャル履歴と同じ大きさ（ブラウザ幅の95%）

  const template = HtmlService.createTemplateFromFile(CHG_OPTIONS.dialogTemplate);
  template.sheetName = target.sheet.getName();
  template.row = target.row;
  template.customer = target.customer;
  template.focusEventId = (options && options.focusEventId) || '';
  template.openedWidth = size.width;
  template.openedHeight = size.height;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(size.width).setHeight(size.height),
    ' '   // 見出しはモーダルの中に表示する
  );
}

/* ---------------- 管理者用セットアップ ---------------- */

function setupChangeHistoryLauncher() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(CRED_OPTIONS.lockWaitMs);

  try {
    const plans = Object.entries(DIFF_RULES).map(([name, rule]) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error('シートが見つかりません: ' + name);

      const launcher = credColumnByHeader_(sheet, rule, CHG_OPTIONS.launcherHeader);
      credColumnByHeader_(sheet, rule, CRED_OPTIONS.customerHeader);   // 得意先列があるか確認
      const stamp = diffStampColumn_(sheet, rule);
      const firstRow = rule.headerRow + 1;
      const rowCount = Math.max(sheet.getMaxRows() - rule.headerRow, 1);
      const column = sheet.getRange(firstRow, launcher, rowCount, 1);

      if (diffBlocks_(sheet, rule, column, stamp, false).length) {
        throw new Error(
          name + ' の「' + CHG_OPTIONS.launcherHeader + '」列（' + diffColumnLetter_(launcher) +
          '列）が DIFF_RULES の ranges に含まれています。範囲から外して setupDiffTracking() を実行してから、もう一度実行してください。'
        );
      }
      return {sheet, launcher, firstRow};
    });

    let buttons = 0;
    for (const p of plans) {
      const lastRow = p.sheet.getLastRow();
      if (lastRow < p.firstRow) continue;
      const count = lastRow - p.firstRow + 1;
      const range = p.sheet.getRange(p.firstRow, p.launcher, count, 1);
      range.clearContent();
      range.clearDataValidations();
      range.insertCheckboxes();
      buttons += count;
    }

    credInstallTrigger_(ss);   // チェックを検知するトリガー（クレデンシャル履歴と共通）
    SpreadsheetApp.flush();
    ss.toast('変更履歴のボタンを ' + buttons + ' 行に設定しました。', '変更履歴', 8);
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

/* ---------------- ダイアログから呼ばれる関数 ---------------- */

/**
 * 対象行（シート＋得意先）の変更履歴を新しい順で返す。
 * {sheetName, row, customer, events: [...], truncated, customerMissing}
 */
function getChangeHistoryData(sheetName, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(String(sheetName));
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  const target = credResolveTarget_(sheet, Number(row));
  const name = target.sheet.getName();

  const result = {
    sheetName: name,
    row: target.row,
    customer: target.customer,
    events: [],
    truncated: false,
    customerMissing: false
  };

  // ---- 変更時点スナップショット
  const snap = ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet);
  if (!snap || snap.getLastRow() < 2) return result;

  const values = snap.getRange(1, 1, snap.getLastRow(), snap.getLastColumn()).getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const col = label => headers.indexOf(label);
  const iAt = col('記録日時'), iEvent = col('イベントID'), iEditor = col('編集者メールアドレス');
  const iSheet = col('シート'), iRow = col('行番号'), iCells = col('変更セル');
  const iCustomer = col(diffCleanHeaderText_(CRED_OPTIONS.customerHeader));
  if (iCustomer < 0) { result.customerMissing = true; return result; }

  // 全項目の表示から外す列（ボタン列）
  const hidden = new Set([CRED_OPTIONS.launcherHeader, CHG_OPTIONS.launcherHeader]
    .map(h => diffCleanHeaderText_(h)));
  const fixed = ROW_SNAPSHOT_FIXED_HEADERS.length;
  // スナップショット同士を比べるときに無視する列（自動で入る最終更新日時）
  const stampKey = diffCleanHeaderText_(DIFF_OPTIONS.stampHeader);
  const ignoreInCompare = key => diffNormalizeStampHeader_(key) === diffNormalizeStampHeader_(stampKey);

  const key = credNormalize_(target.customer);
  let matches = [];
  for (let r = 1; r < values.length; r++) {
    const v = values[r];
    if (v[iSheet] === name && credNormalize_(v[iCustomer]) === key) matches.push(v);
  }
  if (matches.length > CHG_OPTIONS.maxEvents) {
    matches = matches.slice(-CHG_OPTIONS.maxEvents);
    result.truncated = true;
  }

  const snapshotOf = v => headers.slice(fixed)
    .map((h, j) => [h, v[fixed + j]])
    .filter(([h, value]) => h && !hidden.has(h) && value !== '');

  // ---- 変更履歴_差分（同じイベントID・同じシート・同じ行のセル）
  const eventIds = new Set(matches.map(v => v[iEvent]));
  const recordsByEvent = new Map();
  const log = ss.getSheetByName(DIFF_OPTIONS.logSheet);
  if (log && log.getLastRow() >= 2) {
    // B:I = イベントID, 編集者メールアドレス, シート, セル, 列名, 種類, 変更前, 変更後
    log.getRange(2, 2, log.getLastRow() - 1, 8).getDisplayValues().forEach(r => {
      if (!eventIds.has(r[0]) || r[2] !== name) return;
      if (!recordsByEvent.has(r[0])) recordsByEvent.set(r[0], []);
      recordsByEvent.get(r[0]).push({cell: r[3], column: r[4], kind: r[5], before: r[6], after: r[7]});
    });
  }

  const unmark = text => text === CHG_OPTIONS.emptyMark ? '' : text;
  const budget = {remaining: DIFF_OPTIONS.maxLcsCellsPerEdit};
  const withParts = change => Object.assign(change, {
    parts: diffCharacters_(change.before, change.after, budget).parts
  });

  // 古い順に見て、差分ログが無い記録は1つ前のスナップショットとの違いで補う
  let previous = null;
  const events = matches.map(v => {
    const rowNumber = Number(v[iRow]);
    const snapshot = snapshotOf(v);
    let changes = (recordsByEvent.get(v[iEvent]) || [])
      .filter(rec => {
        const m = String(rec.cell).match(/^[A-Z]+(\d+)$/);
        return !m || Number(m[1]) === rowNumber;   // 「5行」などの行全体の記録も含める
      })
      .map(rec => withParts({
        column: rec.column, cell: rec.cell, kind: rec.kind,
        before: unmark(rec.before), after: unmark(rec.after)
      }));

    if (!changes.length && previous) {
      const before = new Map(previous);
      const after = new Map(snapshot);
      const keys = new Set([...before.keys(), ...after.keys()]);
      changes = [...keys]
        .filter(k => !ignoreInCompare(k) && (before.get(k) || '') !== (after.get(k) || ''))
        .map(k => withParts({
          column: k, cell: '', kind: 'スナップショットの比較',
          before: before.get(k) || '', after: after.get(k) || ''
        }));
    }
    previous = snapshot;

    return {
      eventId: v[iEvent],
      type: String(v[iEvent]).split('-')[0],
      at: v[iAt],
      editor: v[iEditor],
      rowNumber,
      cells: v[iCells],
      changes,
      snapshot
    };
  });

  result.events = events.reverse();
  return result;
}
