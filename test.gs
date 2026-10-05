/**
 * Character-level differences for Google Sheets (container-bound Apps Script).
 * Revision: event IDs + editor email + row snapshots + source column names
 *           + lock-busy deferral + catch-up scan + automatic row resync
 *           + row shift detection.
 *
 * CHANGES IN THIS REVISION (v6)：行のずれで履歴が崩れないようにする
 * - 行数が変わらない「ずれ」にも気づく。行の削除と追加がトリガーより先に両方起きて行数が元に戻った場合・
 *   並べ替え・ドラッグでの移動のあとは、比較基準の行と今の行が食い違い、「変更前」に別の行の値が出ていた。
 *   見分け列（DIFF_ROW_OPTIONS.keyHeaders：得意先・案件名）の値が、比較基準では別の行にあった値なら
 *   「ずれた」とみなし、行の突き合わせ（diffResyncRows_）をしてから記録する。
 *   調べるのは、編集トリガー（編集したシート）・変更トリガー（行数が変わったシートと、操作したシート）・
 *   取りこぼし回収（印の付いたシート）だけ。追跡シートが多くても、全シートは読まない。
 * - 突き合わせで、見分け列が同じ行は同じ行として組にする。行がずれたあとの編集は「行削除＋行追加」ではなく、
 *   その行の変更として記録する。移動しただけの行は何も記録しない。
 * - 編集トリガーから突き合わせたとき、編集した範囲の変更はふつうの編集（EVT・編集者つき）として記録し、
 *   それ以外の未記録の変更は編集者「不明」で記録する（編集した人の変更にしない）。
 * - 変更トリガーがロックを取れなかったとき、全シートを「全セル比較」の印にせず、
 *   行数・列数が変わったシートと操作したシートだけに「行の確かめ」の印を付ける（取りこぼし回収で軽く処理する）。
 * - 行の突き合わせ・取りこぼし回収は、データのある最後の行までだけを読む。
 *
 * CHANGES IN v5:
 * - 行の追加・削除で追跡を止めない。行数だけが変わった場合は、自動で
 *   ①比較基準と現在の行を突き合わせ、②削除された行（行全体の内容）・追加された行・
 *   その間に変わったセルを 変更履歴_差分 に記録し、③比較基準を取り直す。
 *   （＝ setupDiffTracking() 相当の処理を、そのシートだけ自動で行う）
 *   イベントIDは「ROW-」、種類は「行削除」「行追加」。
 * - 列の追加・削除は、これまでどおり追跡を停止する（DIFF_RULES の列指定が
 *   ずれるため、確認のうえ setupDiffTracking() を手動で実行する）。
 * - recordDiffEdit() / catchUpDiffTracking() も、行数の変化に気づいたら同じ処理を行う
 *   （変更トリガーが混み合って取りこぼした場合の保険）。
 *
 * CHANGES IN v4:
 * - recordDiffEdit(): if the document lock cannot be obtained, the edit is
 *   NOT treated as an error. Tracking is not paused and no toast is shown.
 *   The sheet is marked "dirty" instead.
 * - catchUpDiffTracking() (time-driven, every few minutes): for dirty sheets,
 *   compares the whole tracked range with the baseline and records any
 *   missed differences (event ID prefix "REC-", editor "不明（取りこぼし回収）").
 * - pauseDiffOnStructureChange(): no longer fails on lock contention.
 * - Unchanged blocks no longer rewrite the baseline.
 *
 * AFTER REPLACING THIS FILE:
 * Run setupDiffTracking() once. It resets the baseline, clears dirty flags,
 * and installs all three triggers (edit / change / time-driven).
 *
 * Setup (new installation):
 * 1. Make sure no other .gs file in the project declares the same names.
 * 2. Edit DIFF_RULES below and add one timestamp header to each source sheet.
 * 3. Run setupDiffTracking() manually once and authorize it.
 *    Only ONE designated account should install these triggers.
 *
 * Re-running setupDiffTracking() takes a NEW baseline. Existing logs remain.
 * Run it after changing tracked ranges or after column changes.
 * Do not delete/edit the hidden __CHAR_DIFF_* sheets while tracking.
 * Sorting/moving rows is detected at the next edit (v6), but rows whose
 * 得意先 and 案件名 are both the same cannot be told apart; avoid sorting.
 * Source cell CONTENT/FORMATTING is never rewritten; only the
 * dedicated timestamp column is written. Differences appear in the log.
 *
 * If the script's required permissions change (e.g. after a code update),
 * re-authorize and run reinstallDiffTriggers().
 *
 * Catch-up records:
 * - The editor and exact edit time of a missed edit are unknown.
 *   The timestamp column receives the catch-up time, not the edit time.
 * - Changes made by other scripts/APIs in the tracked range are also
 *   picked up by a catch-up scan (they never fire the edit trigger).
 *
 * Comparison = last captured cell value, NOT last approved/confirmed value.
 * This is NOT a complete audit trail: intermediate states between captures
 * may be lost. Formula recalculation does not trigger this workflow.
 * Source formulas are compared as expressions, not as calculated results.
 */

// 追跡するシート（キーがシート名）。ほかの機能もこのシート名と headerRow を使う：
// - requests: true を付けたシートは、リクエスト シート（ServiceManagement.gs）がサービスのリクエストを読む
// - formerNames は前のシート名。差分追跡の記録はシート名で残るので、名前を変えたら前の名前をここに足す
//   （前の名前の記録も、このシートの記録としてサイドバー・変更履歴に出す。記録は書き換えない）
// シートを足した・名前を変えたら、setupAfterSheetChange() を実行する（docs/OPERATIONS.md「シート名の一覧」）。
const DIFF_RULES = {
  'スーパー・GMS': {
    headerRow: 2,
    ranges: ['D3:AF'],
    requests: true,          // サービスリクエストの選択パネルに出す（「サービスのリクエスト」の列を読む）
    formerNames: ['新FMT']
    // Optional: capture an explicit row span in the row snapshot.
    // , snapshotColumns: 'A:Z'
  },
  'コンビニ': {headerRow: 2, ranges: ['D3:AF'], requests: true, formerNames: ['新FMT2']},
  '百貨店': {headerRow: 2, ranges: ['D3:AF'], requests: true}
};

/**
 * 記録に残っているシート名を、今の DIFF_RULES のシート名にそろえる関数を返す
 * （formerNames に書いた前の名前なら今の名前、どちらでもなければそのまま）。
 */
function diffSheetNameResolver_() {
  const names = new Map();
  Object.keys(DIFF_RULES).forEach(name => {
    names.set(name, name);
    (DIFF_RULES[name].formerNames || []).forEach(former => { if (!names.has(former)) names.set(String(former), name); });
  });
  return value => {
    const text = String(value == null ? '' : value);
    return names.get(text) || text;
  };
}

const DIFF_OPTIONS = {
  stampHeader: '最終更新日時',
  logSheet: '変更履歴_差分',
  rowSnapshotSheet: '変更時点スナップショット',
  snapshotPrefix: '__CHAR_DIFF_',
  stateKey: 'CHAR_DIFF_STATE_V1',
  dirtyPrefix: 'CHAR_DIFF_DIRTY_V1_',
  editLockWaitMs: 60000,       // edit trigger: wait this long, then defer to catch-up
  catchUpLockWaitMs: 10000,    // catch-up: wait this long, then retry next run
  catchUpMinutes: 5,           // allowed: 1, 5, 10, 15, 30
  catchUpChunkCells: 5000,
  catchUpEditor: '不明（取りこぼし回収）',
  catchUpKindSuffix: ' / 取りこぼし回収',
  maxCellsPerEdit: 2000,
  maxSnapshotCellsPerEdit: 50000,
  maxTokenLength: 45000,
  maxLcsCells: 1000000,
  maxLcsCellsPerEdit: 4000000,
  dateFormat: 'yyyy/MM/dd HH:mm',
  // 変更時点スナップショットの「行の内容」：その行の見出しと値（空欄の項目は持たない）を1つのセルに JSON（[[見出し, 値], …]）で入れる
  snapshotContentHeader: '行の内容',
  snapshotContentLimit: 45000   // 1つのセルに入れる文字数の上限（超えるときは長い値を省略する）
};

// 行の追加・削除の自動処理（v5）
const DIFF_ROW_OPTIONS = {
  eventPrefix: 'ROW',
  unknownEditor: '不明（行の追加・削除）',
  kindDelete: '行削除',
  kindInsert: '行追加',
  kindSuffix: ' / 行の追加・削除時に回収',
  structureLockWaitMs: 30000,
  maxAlignCells: 4000000,      // 行の突き合わせ（LCS）の上限。超えたら変化した範囲をまとめて扱う
  readChunkCells: 5000,
  maxSummaryCharacters: 40000,
  // 行を見分ける列の見出し（追跡範囲の中の列）。行のずれの検知と、突き合わせで同じ行を見つけるのに使う。
  // 見つからない見出しは使わない（1つも見つからないシートでは、ずれの検知はしない）
  keyHeaders: ['得意先', '案件名'],
  // 「行の確かめ」だけが必要なシートの印（取りこぼし回収で、全セルは比べずに行数・ずれだけを確かめる）
  checkPrefix: 'CHAR_DIFF_ROWCHECK_V1_'
};

// セル数を節約するため9列（前の版の「差分」「確認済み」はやめた。文字単位の差分は変更履歴の画面がその場で作る）。
// Changehistory.gs・HistorySidebar.gs は B〜I 列を位置で読む
const DIFF_HEADERS = [
  '記録日時', 'イベントID', '編集者メールアドレス', 'シート', 'セル', '列名',
  '種類', '変更前', '変更後'
];

// 前の版の11列（「差分」「確認済み」がある）。compactDiffRecords() で2列を削除する
const DIFF_LEGACY_TAIL_HEADERS = ['差分', '確認済み'];

// Previous 10-column schema: editor email does not exist.
const DIFF_LEGACY_HEADERS_10 = [
  '記録日時', 'イベントID', 'シート', 'セル', '列名',
  '種類', '変更前', '変更後',
  '差分', '確認済み'
];

// Previous 9-column schema: column name exists, event ID does not.
const DIFF_LEGACY_HEADERS_9 = [
  '記録日時', 'シート', 'セル', '列名',
  '種類', '変更前', '変更後',
  '差分', '確認済み'
];

// Oldest 8-column schema: neither event ID nor column name exists.
const DIFF_LEGACY_HEADERS_8 = [
  '記録日時', 'シート', 'セル',
  '種類', '変更前', '変更後',
  '差分', '確認済み'
];

const ROW_SNAPSHOT_FIXED_HEADERS = [
  '記録日時', 'イベントID', '編集者メールアドレス', 'シート', '行番号', '変更セル'
];

const ROW_SNAPSHOT_LEGACY_HEADERS_5 = [
  '記録日時', 'イベントID', 'シート', '行番号', '変更セル'
];

const DIFF_TRIGGER_HANDLERS = [
  'recordDiffEdit',
  'pauseDiffOnStructureChange',
  'catchUpDiffTracking'
];

/**
 * Run once when upgrading the log layout from an older version.
 * It does NOT reset the comparison baseline.
 * If DIFF_RULES itself changed, run setupDiffTracking() instead.
 */
function upgradeDiffTrackingV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    diffLogSheet_(ss);
    diffRowSnapshotSheet_(ss);
    diffInstallTriggers_(ss);
    ss.toast(
      '差分履歴・行スナップショットに編集者メールアドレス列を追加しました。',
      'Diff tracking',
      8
    );
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

/** Backward-compatible alias. */
function upgradeDiffTrackingV2() {
  upgradeDiffTrackingV3();
}

/** Backward-compatible alias. */
function upgradeDiffLogColumns() {
  upgradeDiffTrackingV3();
}

/** Run manually. Resets the comparison baseline; never deletes log history. */
function setupDiffTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const previous = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
    const plans = Object.entries(DIFF_RULES).map(([name, rule]) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error('Source sheet not found: ' + name);
      if (name === DIFF_OPTIONS.logSheet ||
          name === DIFF_OPTIONS.rowSnapshotSheet ||
          name.startsWith(DIFF_OPTIONS.snapshotPrefix)) {
        throw new Error('Reserved sheet name: ' + name);
      }
      const column = diffStampColumn_(sheet, rule);
      const blocks = diffBlocks_(sheet, rule, null, column, true);
      for (const block of blocks) {
        const range = sheet.getRange(...block);
        if (range.getMergedRanges().length) {
          throw new Error('Merged cells are not supported: ' + name + '!' + range.getA1Notation());
        }
      }
      const snapshotName = DIFF_OPTIONS.snapshotPrefix + sheet.getSheetId();
      const existing = ss.getSheetByName(snapshotName);
      const owned = previous[sheet.getSheetId()];
      if (existing && owned && owned.snapshotId !== existing.getSheetId()) {
        throw new Error('Reserved snapshot name already exists: ' + snapshotName);
      }
      if (existing && !owned) {
        console.log('Adopting existing snapshot sheet (copied file?): ' + snapshotName);
      }
      return {sheet, rule, column, blocks, existing, snapshotName};
    });
    diffLogSheet_(ss);
    diffRowSnapshotSheet_(ss);

    // Clear dirty flags BEFORE reading the baseline. Any edit deferred while
    // this setup holds the lock sets a new flag and is caught up afterwards.
    plans.forEach(plan => {
      diffClearDirty_(plan.sheet.getSheetId());
      diffClearRowCheck_(plan.sheet.getSheetId());
    });

    const state = {};
    // Mark tracking as paused until the complete baseline is saved.
    Object.keys(previous).forEach(id => previous[id].paused = true);
    props.setProperty(DIFF_OPTIONS.stateKey, JSON.stringify(previous));
    for (const plan of plans) {
      const {sheet, rule, column, blocks, snapshotName} = plan;
      const snapshot = plan.existing || ss.insertSheet(snapshotName);
      // Record ownership immediately so a failed setup can be safely retried.
      previous[sheet.getSheetId()] = {snapshotId: snapshot.getSheetId(), paused: true};
      props.setProperty(DIFF_OPTIONS.stateKey, JSON.stringify(previous));
      snapshot.clearContents();
      for (const [row, col, rows, cols] of blocks) {
        diffEnsureGrid_(snapshot, row + rows - 1, col + cols - 1);
        // Bound each baseline read/write to about 2,000 cells.
        const step = Math.max(1, Math.floor(2000 / cols));
        for (let offset = 0; offset < rows; offset += step) {
          const count = Math.min(step, rows - offset);
          const source = sheet.getRange(row + offset, col, count, cols);
          snapshot.getRange(row + offset, col, count, cols)
            .setValues(diffTokens_(source));
        }
      }
      // 比較基準は追跡範囲の右端の列・追跡シートの行数までにする（空のセルも上限に数えられるため）
      diffTrimGrid_(snapshot, Math.max(...blocks.map(([, col, , cols]) => col + cols - 1), 1), sheet.getMaxRows());
      snapshot.hideSheet();
      state[sheet.getSheetId()] = {
        snapshotId: snapshot.getSheetId(),
        name: sheet.getName(),
        signature: JSON.stringify(rule),
        stampColumn: column,
        gridRows: sheet.getMaxRows(),
        gridColumns: sheet.getMaxColumns(),
        paused: false
      };
    }
    SpreadsheetApp.flush();
    props.setProperty(DIFF_OPTIONS.stateKey, JSON.stringify(state));
    diffInstallTriggers_(ss);
    ss.toast('現在の値を比較基準にしました。', 'Diff tracking', 8);
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/* ---------------- シートを足した・名前を変えたあと（管理者がエディタから実行する） ---------------- */

/**
 * DIFF_RULES のシートを足した・名前を変えたあとに必要な設定を、まとめて実行する。何度実行してもよい。
 * 1. 差分追跡（setupDiffTracking）  2. クレデンシャルのボタン列（setupCredentialLauncher）
 * 3. 変更履歴のボタン列（setupChangeHistoryLauncher）  4. サービスリクエスト（setupRequestSheet）
 * ボタン列は、見出しのあるシートにだけ付ける（どのシートにも無ければ 2・3 は省く）。
 * 途中で止まって設定が半分だけにならないよう、先にすべてのシートの見出しを確かめ、足りなければ何も変えずに知らせる。
 * トリガーを設置するので、管理者アカウント（トリガーの所有者）で実行すること。
 */
function setupAfterSheetChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const {problems, buttons} = diffCheckSheets_(ss);
  if (problems.length) {
    ui.alert('シートの設定', '次の点を直してから、もう一度実行してください（まだ何も変えていません）。\n\n・' + problems.join('\n・'),
      ui.ButtonSet.OK);
    return;
  }

  let notes = [];
  const steps = [{label: '差分追跡', run: () => setupDiffTracking()}];
  [
    {label: 'クレデンシャルのボタン列', header: CRED_OPTIONS.launcherHeader, sheets: buttons.credential, run: () => setupCredentialLauncher()},
    {label: '変更履歴のボタン列', header: CHG_OPTIONS.launcherHeader, sheets: buttons.change, run: () => setupChangeHistoryLauncher()}
  ].forEach(step => {
    if (step.sheets.length) steps.push(step);
    else notes.push('「' + credHeaderText_(step.header) + '」の列のあるシートが無いので、' + step.label + 'の設定は省きました。');
  });
  steps.push({label: 'サービスリクエスト', run: () => svcSetupRequestSheet_(ss)});
  const done = [];
  steps.forEach(step => {
    try {
      const result = step.run();
      if (Array.isArray(result)) notes = notes.concat(result);
      else if (typeof result === 'string') notes.push(result);
    } catch (error) {
      throw new Error('「' + step.label + '」で止まりました（' + error.message + '）。' +
        (done.length ? '「' + done.join('」「') + '」は済んでいます。' : '') +
        '原因を直してから、もう一度 setupAfterSheetChange() を実行してください（何度実行しても大丈夫です）。');
    }
    done.push(step.label);
  });

  ui.alert('シートの設定',
    '対象のシート：「' + Object.keys(DIFF_RULES).join('」「') + '」\n' +
    '「' + done.join('」「') + '」を設定しました。\n\n' + notes.join('\n'),
    ui.ButtonSet.OK);
}

/**
 * DIFF_RULES のシートが setupAfterSheetChange() の設定に足りているかを確かめる（読むだけ）。
 * シートがあるか・見出し（最終更新日時・得意先。requests: true ならリクエストの4つの見出し）があるか・
 * ボタン列（使うシートだけに置く。無くてもよい）があれば、1つだけで追跡範囲（ranges）の外にあるか。
 * {problems: 直す点, buttons: {credential: ボタン列のあるシート, change: 変更履歴のボタン列のあるシート}}
 */
function diffCheckSheets_(ss) {
  const problems = [];
  const buttons = {credential: [], change: []};
  Object.entries(DIFF_RULES).forEach(([name, rule]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      problems.push('「' + name + '」シートがありません（DIFF_RULES のシート名と、シートの名前を同じにしてください）。');
      return;
    }
    const check = fn => {
      try { return fn(); } catch (error) { problems.push(error.message); return 0; }
    };
    if (!Number.isInteger(rule.headerRow) || rule.headerRow < 1 || !Array.isArray(rule.ranges) || !rule.ranges.length) {
      problems.push('DIFF_RULES の「' + name + '」の headerRow（見出しの行）と ranges（追跡範囲）を確かめてください。');
      return;
    }
    let stamp = 0;
    try {
      stamp = diffStampColumn_(sheet, rule);
    } catch (_) {
      problems.push(name + ' の ' + rule.headerRow + '行目に「' + DIFF_OPTIONS.stampHeader + '」の見出しがちょうど1つ必要です。');
    }
    check(() => credColumnByHeader_(sheet, rule, CRED_OPTIONS.customerHeader));
    [['credential', CRED_OPTIONS.launcherHeader], ['change', CHG_OPTIONS.launcherHeader]].forEach(([kind, header]) => {
      const launcher = check(() => credColumnByHeader_(sheet, rule, header, true));
      if (!launcher) return;
      buttons[kind].push(name);
      if (!stamp) return;
      const column = sheet.getRange(rule.headerRow + 1, launcher, Math.max(sheet.getMaxRows() - rule.headerRow, 1), 1);
      if (diffBlocks_(sheet, rule, column, stamp, false).length) {
        problems.push(name + ' の「' + credHeaderText_(header) + '」列（' + diffColumnLetter_(launcher) +
          '列）が DIFF_RULES の ranges に含まれています。ボタン列は追跡範囲の外にしてください。');
      }
    });
    if (rule.requests) {
      const headers = sheet.getRange(rule.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
      SVC_OPTIONS.sourceHeaders.forEach(header => {
        check(() => taskFindColumn_(headers, header, name + ' の ' + rule.headerRow + '行目', true));
      });
    }
  });
  return {problems, buttons};
}

/**
 * Installed edit trigger. Do not run manually or add an onEdit wrapper.
 *
 * - differences are recorded per actually changed cell
 * - row snapshot = one record per event ID x changed row
 * - if the lock is busy, the edit is deferred to catchUpDiffTracking()
 * - if only the number of rows changed (row insert/delete), or the rows are
 *   shifted (v6), the sheet is resynchronized first. The edited cells are
 *   still recorded as a normal edit (EVT) with the editor.
 */
function recordDiffEdit(e) {
  console.log('recordDiffEdit fired: ' + (e && e.range
    ? e.range.getSheet().getName() + '!' + e.range.getA1Notation()
    : 'no event'));

  if (!e || !e.range || !e.source) return;

  const sheet = e.range.getSheet();
  const rule = DIFF_RULES[sheet.getName()];
  if (!rule || e.range.getLastRow() <= rule.headerRow) return;
  // 追加：履歴ボタン列のチェック操作は差分追跡の対象外
  if (typeof credIsLauncherEdit_ === 'function' && credIsLauncherEdit_(e, rule)) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DIFF_OPTIONS.editLockWaitMs)) {
    // Not an error: another execution is busy. Leave the baseline untouched
    // and let the catch-up scan record this change later.
    diffMarkDirty_(sheet.getSheetId());
    console.warn('Lock busy; deferred to catch-up: ' +
      sheet.getName() + '!' + e.range.getA1Notation());
    return;
  }

  try {
    const edit = {
      r1: e.range.getRow(), r2: e.range.getLastRow(),
      c1: e.range.getColumn(), c2: e.range.getLastColumn()
    };
    const resync = () => diffResyncRows_(e.source, sheet, rule, {
      structure: DIFF_ROW_OPTIONS.unknownEditor,
      changes: diffEditorEmail_(e),
      others: DIFF_ROW_OPTIONS.unknownEditor,
      edit
    });

    // 行の追加・削除の直後で、変更トリガーより先にこの編集が届いた場合
    if (diffRowCountOnlyChanged_(sheet, rule)) {
      resync();
      return;
    }

    const ctx = diffContext_(e.source, sheet, rule);

    // 行数は同じでも行がずれている（削除と追加が続いた・並べ替え・移動）なら、突き合わせ直してから記録する
    // （そのまま比べると、別の行の値が「変更前」になる）
    if (diffRowsShifted_(sheet, rule, ctx.baseline, ctx.stampColumn, edit)) {
      resync();
      return;
    }

    const blocks = diffBlocks_(sheet, rule, e.range, ctx.stampColumn, false);
    if (!blocks.length) return;

    const total = blocks.reduce((n, b) => n + b[2] * b[3], 0);
    if (total > DIFF_OPTIONS.maxCellsPerEdit) {
      throw new Error('Edit exceeds ' + DIFF_OPTIONS.maxCellsPerEdit + ' tracked cells.');
    }

    const now = new Date();
    diffProcessBlocks_(ctx, blocks, {
      now,
      eventId: diffCreateEventId_(now, ctx.timezone, 'EVT'),
      editorEmail: diffEditorEmail_(e),
      kindSuffix: ''
    });

  } catch (error) {
    try {
      diffPause_(e.source, sheet.getSheetId());
    } catch (pauseError) {
      console.error(pauseError.stack || String(pauseError));
    }

    console.error(error.stack || String(error));

    try {
      e.source.toast(
        '差分記録を停止しました。Apps Scriptの実行履歴を確認してください。',
        'Diff error',
        15
      );
    } catch (_) {}

    throw error;

  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

/**
 * Installed time-driven trigger (every DIFF_OPTIONS.catchUpMinutes minutes).
 * Does nothing unless a sheet was marked by a deferred edit (dirty) or by a
 * deferred structure change (row check).
 * - 行数の変化・行のずれがあれば、行の突き合わせ（diffResyncRows_）で処理する（変わったセルもここで記録される）
 * - 「行の確かめ」の印だけなら、全セルは比べない
 * - dirty の印なら、データのある最後の行までを比較基準と比べ、見つかった違いをすべて記録する
 * 印は全シートぶんを1回で読むので、追跡シートが多くても、印が無ければすぐに終わる。
 */
function catchUpDiffTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  if (!diffPendingSheetIds_(props.getProperties()).length) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DIFF_OPTIONS.catchUpLockWaitMs)) {
    console.warn('Lock busy; catch-up will retry on the next run.');
    return;
  }

  try {
    const all = props.getProperties();   // ロックを取ったあとで読み直す
    const state = JSON.parse(all[DIFF_OPTIONS.stateKey] || '{}');

    for (const id of diffPendingSheetIds_(all)) {
      const saved = state[id];
      const rule = DIFF_RULES[saved.name];
      const sheet = ss.getSheetByName(saved.name);
      if (!sheet || String(sheet.getSheetId()) !== id) {
        console.warn('Source sheet not found; catch-up skipped: ' + saved.name);
        continue;
      }
      const dirtyKey = DIFF_OPTIONS.dirtyPrefix + id;
      const checkKey = DIFF_ROW_OPTIONS.checkPrefix + id;
      const dirtyToken = all[dirtyKey];
      const checkToken = all[checkKey];

      try {
        // 列の追加・削除など（変更トリガーがロックを取れずに処理できなかった場合）は停止する
        if (!diffSameLayout_(sheet, rule, saved)) {
          throw new Error('Structure/configuration changed. Reinitialize the baseline before editing.');
        }

        // 行の追加・削除・ずれを変更トリガーが処理しきれていなければ、ここで処理する
        const baseline = diffBaseline_(ss, saved);
        if (saved.gridRows !== sheet.getMaxRows() ||
            diffRowsShifted_(sheet, rule, baseline, saved.stampColumn, null)) {
          diffResyncRows_(ss, sheet, rule, {
            structure: DIFF_ROW_OPTIONS.unknownEditor,
            changes: DIFF_OPTIONS.catchUpEditor
          });
          continue;
        }

        if (!dirtyToken) {
          // 行の確かめだけだった（行数もずれも変わっていない）
          if (props.getProperty(checkKey) === checkToken) props.deleteProperty(checkKey);
          continue;
        }

        const ctx = diffContext_(ss, sheet, rule);
        const lastRow = Math.max(sheet.getLastRow(), ctx.baseline.getLastRow());
        const blocks = diffSplitBlocks_(
          diffLimitBlocks_(diffBlocks_(sheet, rule, null, ctx.stampColumn, false), lastRow),
          DIFF_OPTIONS.catchUpChunkCells
        );
        const now = new Date();
        const count = diffProcessBlocks_(ctx, blocks, {
          now,
          eventId: diffCreateEventId_(now, ctx.timezone, 'REC'),
          editorEmail: DIFF_OPTIONS.catchUpEditor,
          kindSuffix: DIFF_OPTIONS.catchUpKindSuffix
        });
        console.log('Catch-up done: ' + saved.name + ' / recovered cells=' + count);

        // Clear only if no newer deferral happened during this scan.
        if (props.getProperty(dirtyKey) === dirtyToken) props.deleteProperty(dirtyKey);
        if (checkToken && props.getProperty(checkKey) === checkToken) props.deleteProperty(checkKey);

      } catch (error) {
        diffPause_(ss, id);
        console.error(error.stack || String(error));
        throw error;
      }
    }
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

/**
 * 取りこぼし回収が必要なシートの ID（追跡中で、dirty か「行の確かめ」の印があるもの）。
 * all = ドキュメントのプロパティすべて（getProperties() を1回だけ呼んで渡す）
 */
function diffPendingSheetIds_(all) {
  const state = JSON.parse(all[DIFF_OPTIONS.stateKey] || '{}');
  return Object.keys(state).filter(id => {
    const saved = state[id];
    return saved && !saved.paused && DIFF_RULES[saved.name] &&
      !!(all[DIFF_OPTIONS.dirtyPrefix + id] || all[DIFF_ROW_OPTIONS.checkPrefix + id]);
  });
}

/** 範囲を lastRow 行目までに切り詰める（その下はシートも比較基準も空なので比べなくてよい）。 */
function diffLimitBlocks_(blocks, lastRow) {
  return blocks
    .map(([row, col, rows, cols]) => [row, col, Math.min(rows, lastRow - row + 1), cols])
    .filter(block => block[2] > 0);
}

/**
 * Installed change trigger.
 *
 * - 行の追加・削除（行数だけが変わった場合）：追跡は止めずに diffResyncRows_() で
 *   削除・追加された行を記録し、そのシートの比較基準を取り直す。
 * - 行の追加・削除の通知なのに、操作したシートの行数が変わっていない（削除と追加が続いて行数が戻った）：
 *   そのシートだけ見分け列で行のずれを確かめ、ずれていれば同じように突き合わせ直す（v6）。
 *   通知にはシートの情報が無いため、操作したシートはアクティブなシートで判断する。ほかのシートのずれは、
 *   次の編集・取りこぼし回収のときに気づく。
 * - 列の追加・削除、シートの削除：これまでどおり、そのシートの追跡を停止する。
 *   DIFF_RULES の列指定がずれるため、確認のうえ setupDiffTracking() を手動で実行する。
 * - ロックが取れないときは、行数・列数が変わったシートと操作したシートだけに「行の確かめ」の印を付け、
 *   catchUpDiffTracking() で処理する（全シートの全セル比較はしない）。
 */
function pauseDiffOnStructureChange(e) {
  if (!e || !e.source) return;

  const type = String(e.changeType || '');
  const watched = [
    'INSERT_ROW',
    'REMOVE_ROW',
    'INSERT_COLUMN',
    'REMOVE_COLUMN',
    'REMOVE_GRID'
  ];
  if (!watched.includes(type)) return;
  const rowEvent = type === 'INSERT_ROW' || type === 'REMOVE_ROW';

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DIFF_ROW_OPTIONS.structureLockWaitMs)) {
    // 取りこぼし回収に任せる（行数の変化・行のずれは catchUpDiffTracking() でも処理される）
    diffFlagStructureChange_(e.source, rowEvent);
    console.warn('Lock busy; structure change deferred to catch-up (' + type + ').');
    return;
  }

  try {
    const props = PropertiesService.getDocumentProperties();
    const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
    const sheetsById = new Map(
      e.source.getSheets().map(s => [String(s.getSheetId()), s])
    );
    const pausedNames = [];
    const rowChanged = [];

    for (const [id, saved] of Object.entries(state)) {
      if (saved.paused) continue;
      const sheet = sheetsById.get(id);
      if (!sheet || saved.gridColumns !== sheet.getMaxColumns()) {
        saved.paused = true;
        pausedNames.push(saved.name || id);
      } else if (saved.gridRows !== sheet.getMaxRows()) {
        rowChanged.push(sheet);
      }
    }

    if (pausedNames.length) {
      props.setProperty(DIFF_OPTIONS.stateKey, JSON.stringify(state));
      e.source.toast(
        pausedNames.join(', ') +
        ' の列構成が変わったため差分追跡を停止しました。DIFF_RULES を確認し、setupDiffTracking()を再実行してください。',
        'Diff paused',
        15
      );
    }

    // 行数が変わっていなくても、操作したシートの行がずれていれば突き合わせ直す
    const shifted = [];
    if (rowEvent) {
      const active = diffActiveTrackedSheet_(e.source);
      const id = active ? String(active.getSheetId()) : '';
      const saved = state[id];
      try {
        if (active && saved && !saved.paused && !rowChanged.some(s => String(s.getSheetId()) === id) &&
            diffSameLayout_(active, DIFF_RULES[active.getName()], saved) &&
            diffRowsShifted_(active, DIFF_RULES[active.getName()], diffBaseline_(e.source, saved),
              saved.stampColumn, null)) {
          shifted.push(active);
        }
      } catch (error) {
        // 確かめられなかったときは、次の編集・取りこぼし回収に任せる
        console.error(error.stack || String(error));
        diffMarkRowCheck_(id);
      }
    }

    const editor = diffEditorEmail_(e);
    const structureEditor = editor === '取得不可' ? DIFF_ROW_OPTIONS.unknownEditor : editor;

    for (const sheet of rowChanged.concat(shifted)) {
      const rule = DIFF_RULES[sheet.getName()];
      const saved = state[String(sheet.getSheetId())];
      if (!rule || !diffSameLayout_(sheet, rule, saved)) {
        diffPause_(e.source, sheet.getSheetId());
        e.source.toast(
          sheet.getName() + ' の構成が変わったため差分追跡を停止しました。setupDiffTracking()を再実行してください。',
          'Diff paused',
          15
        );
        continue;
      }
      try {
        const result = diffResyncRows_(e.source, sheet, rule, {
          structure: structureEditor,
          changes: DIFF_ROW_OPTIONS.unknownEditor
        });
        const parts = [];
        if (result.deleted) parts.push('削除 ' + result.deleted + '行');
        if (result.inserted) parts.push('追加 ' + result.inserted + '行');
        if (result.changedCells) parts.push('変更 ' + result.changedCells + 'セル');
        e.source.toast(
          sheet.getName() + ' の行の追加・削除を反映しました' +
            (parts.length ? '（' + parts.join('・') + 'を記録）' : '') + '。',
          'Diff tracking',
          6
        );
      } catch (error) {
        diffPause_(e.source, sheet.getSheetId());
        console.error(error.stack || String(error));
        e.source.toast(
          sheet.getName() + ' の行の追加・削除を反映できず、差分追跡を停止しました。setupDiffTracking()を再実行してください。',
          'Diff error',
          15
        );
      }
    }
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

/**
 * 変更トリガーがロックを取れなかったときの印。行数・列数が変わったシートと、
 * 行の追加・削除の通知なら操作したシート（アクティブなシート）だけに「行の確かめ」の印を付ける。
 * ロックの外で読むだけ（印はプロパティに書くだけで、シートには書き込まない）。
 */
function diffFlagStructureChange_(ss, rowEvent) {
  const state = JSON.parse(
    PropertiesService.getDocumentProperties().getProperty(DIFF_OPTIONS.stateKey) || '{}');
  const sheetsById = new Map(ss.getSheets().map(s => [String(s.getSheetId()), s]));
  const active = rowEvent ? diffActiveTrackedSheet_(ss) : null;
  const activeId = active ? String(active.getSheetId()) : '';

  Object.entries(state).forEach(([id, saved]) => {
    if (!saved || saved.paused || !DIFF_RULES[saved.name]) return;
    const sheet = sheetsById.get(id);
    if (!sheet) return;
    if (id === activeId ||
        saved.gridRows !== sheet.getMaxRows() ||
        saved.gridColumns !== sheet.getMaxColumns()) {
      diffMarkRowCheck_(id);
    }
  });
}

/** 操作したシート（アクティブなシート）が追跡シートならそれを、違えば null。 */
function diffActiveTrackedSheet_(ss) {
  try {
    const sheet = ss.getActiveSheet();
    return sheet && DIFF_RULES[sheet.getName()] ? sheet : null;
  } catch (_) {
    return null;
  }
}

/* ---------------- Row insert/delete: record and rebaseline (v5) ---------------- */

/** 追跡の状態（setupDiffTracking() が保存したもの）。 */
function diffSavedState_(sheet) {
  const state = JSON.parse(
    PropertiesService.getDocumentProperties().getProperty(DIFF_OPTIONS.stateKey) || '{}');
  return state[sheet.getSheetId()];
}

/** 列・設定・シート名が比較基準と同じ（行数は問わない）で、追跡中なら true。 */
function diffSameLayout_(sheet, rule, saved) {
  if (!rule || !saved || saved.paused) return false;

  let stampColumn;
  try { stampColumn = diffStampColumn_(sheet, rule); } catch (_) { return false; }

  return saved.name === sheet.getName() &&
    saved.signature === JSON.stringify(rule) &&
    saved.stampColumn === stampColumn &&
    saved.gridColumns === sheet.getMaxColumns();
}

/**
 * 行数だけが比較基準と違う（列・設定・シート名は同じ）なら true。
 * このときは停止せず diffResyncRows_() で処理できる。
 */
function diffRowCountOnlyChanged_(sheet, rule) {
  const saved = diffSavedState_(sheet);
  return !!saved && saved.gridRows !== sheet.getMaxRows() && diffSameLayout_(sheet, rule, saved);
}

/** 比較基準のシート（非表示の __CHAR_DIFF_*）。 */
function diffBaseline_(ss, saved) {
  const baseline = ss.getSheets().find(s => s.getSheetId() === saved.snapshotId);
  if (!baseline) throw new Error('Comparison snapshot is missing.');
  return baseline;
}

/**
 * 見分け列（DIFF_ROW_OPTIONS.keyHeaders のうち、追跡範囲にある列）の列番号。
 * 見出しの改行・空白は無視して照合する。
 */
function diffKeyColumns_(sheet, rule, blocks) {
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  const headers = sheet
    .getRange(rule.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getDisplayValues()[0]
    .map(normalize);
  const tracked = column => blocks.some(b => column >= b[1] && column < b[1] + b[3]);

  const columns = [];
  DIFF_ROW_OPTIONS.keyHeaders.forEach(header => {
    const index = headers.indexOf(normalize(header));
    if (index >= 0 && tracked(index + 1) && !columns.includes(index + 1)) columns.push(index + 1);
  });
  return columns;
}

/**
 * 行のずれを調べる（ロック取得中に呼ぶこと）。比較基準の行と今の行が食い違っていれば true。
 * ある行の見分け列の値が、比較基準では別の行にあった値なら「ずれた」とみなす
 * （行数が変わらない行の削除と追加・並べ替え・ドラッグでの移動）。
 * 比較基準のどこにも無い値（新しく入力した・書き換えた値）は、ふつうの変更として扱う。
 * edit（今回の編集の範囲 {r1, r2, c1, c2}）で見分け列が変わった行は、手がかりにしない。
 * 読むのは、見分け列をまたぐ1つの範囲（シート・比較基準）を、データのある最後の行までだけ。
 */
function diffRowsShifted_(sheet, rule, baseline, stampColumn, edit) {
  const blocks = diffBlocks_(sheet, rule, null, stampColumn, false);
  if (!blocks.length) return false;
  const keys = diffKeyColumns_(sheet, rule, blocks);
  if (!keys.length) return false;

  const first = Math.min(...blocks.map(b => b[0]));
  const last = Math.max(sheet.getLastRow(), baseline.getLastRow());
  if (last < first) return false;
  const count = last - first + 1;
  const left = Math.min(...keys);
  const width = Math.max(...keys) - left + 1;

  // 1行を1つの文字列にする（見分け列がすべて空なら ''）
  const read = (target, isBaseline) => {
    const rows = Math.max(0, Math.min(count, target.getMaxRows() - first + 1));
    let values = [];
    if (rows && left + width - 1 <= target.getMaxColumns()) {
      const range = target.getRange(first, left, rows, width);
      values = isBaseline
        ? range.getValues().map(row => row.map(v => String(v || '')))
        : diffTokens_(range);
    }
    const list = values.map(row => keys.some(c => row[c - left])
      ? keys.map(c => row[c - left]).join('\u0001')
      : '');
    while (list.length < count) list.push('');
    return list;
  };
  const before = read(baseline, true);
  const after = read(sheet, false);

  const known = new Set(before.filter(Boolean));
  const keyEdited = row => !!edit && row >= edit.r1 && row <= edit.r2 &&
    keys.some(c => c >= edit.c1 && c <= edit.c2);
  for (let i = 0; i < count; i++) {
    if (after[i] === before[i] || !after[i] || keyEdited(first + i)) continue;
    if (known.has(after[i])) return true;   // 比較基準では別の行にあった値が、この行に来ている
  }
  return false;
}

/**
 * ロック取得中に呼ぶこと。比較基準と現在の行を突き合わせて、
 * 削除された行・追加された行・その間に変わったセルを記録し、比較基準を取り直す。
 * 行数が変わった場合のほか、行数が同じまま行がずれた場合（v6）にも使う。
 * who = {structure: 行の追加・削除をした人, changes: セルを変えた人,
 *        edit: 編集トリガーから呼ぶときの編集範囲 {r1, r2, c1, c2}, others: 編集範囲の外の変更の編集者}
 * edit があれば、編集範囲の変更はふつうの編集（EVT・changes）として、それ以外とは別のイベントで記録する。
 * 戻り値 {deleted, inserted, changedCells}
 */
function diffResyncRows_(ss, sheet, rule, who) {
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
  const sheetId = sheet.getSheetId();
  const saved = state[sheetId];
  if (!saved) throw new Error('Tracking is not initialized. Run setupDiffTracking().');

  const baseline = diffBaseline_(ss, saved);
  diffLogSheet_(ss);
  diffRowSnapshotSheet_(ss);

  const timezone = ss.getSpreadsheetTimeZone();
  const stampColumn = saved.stampColumn;
  const blocks = diffBlocks_(sheet, rule, null, stampColumn, false);
  const result = {deleted: 0, inserted: 0, changedCells: 0};

  // 読み込みより前に印を消す（このあと届いた編集は印が付き、取りこぼし回収される）
  diffClearDirty_(sheetId);
  diffClearRowCheck_(sheetId);

  if (blocks.length) {
    for (const block of blocks) {
      if (sheet.getRange(...block).getMergedRanges().length) {
        throw new Error('Merged cells are not supported: ' + sheet.getRange(...block).getA1Notation());
      }
    }

    // データのある最後の行までを読む（その下は、シートも比較基準も空なので比べなくてよい）
    const rowStart = Math.min(...blocks.map(b => b[0]));
    const newEnd = Math.min(sheet.getMaxRows(), sheet.getLastRow());
    const oldEnd = baseline.getLastRow();
    const newCount = Math.max(0, newEnd - rowStart + 1);
    const oldCount = Math.max(0, oldEnd - rowStart + 1);
    const specs = blocks.map(b => ({r1: b[0], col: b[1], cols: b[3]}));
    const maxColumn = Math.max(...specs.map(s => s.col + s.cols - 1));

    diffEnsureGrid_(baseline, Math.max(oldEnd, 1), maxColumn);
    const oldRows = diffReadRowTokens_(baseline, true, specs, rowStart, oldCount);
    const newRows = diffReadRowTokens_(sheet, false, specs, rowStart, newCount);

    const columnNames = [];
    const columnNumbers = [];
    specs.forEach(spec => {
      diffColumnNames_(sheet, rule, spec.col, spec.cols).forEach(n => columnNames.push(n));
      for (let c = 0; c < spec.cols; c++) columnNumbers.push(spec.col + c);
    });
    const keyIndexes = diffKeyColumns_(sheet, rule, blocks)
      .map(column => columnNumbers.indexOf(column))
      .filter(index => index >= 0);

    const matched = diffPairRows_(
      diffAlignRows_(
        oldRows.map(r => r.join('\u0001')),
        newRows.map(r => r.join('\u0001'))
      ),
      oldRows,
      newRows,
      keyIndexes
    );

    const now = new Date();
    const rowEventId = diffCreateEventId_(now, timezone, DIFF_ROW_OPTIONS.eventPrefix);
    // 編集トリガーから呼ばれたときは、編集した範囲の変更を、ふつうの編集として別のイベントにする
    const editEventId = who.edit ? diffCreateEventId_(now, timezone, 'EVT') : '';
    const otherEditor = who.edit ? (who.others || DIFF_ROW_OPTIONS.unknownEditor) : who.changes;
    const inEdit = (rowNumber, column) => !!who.edit &&
      rowNumber >= who.edit.r1 && rowNumber <= who.edit.r2 &&
      (column == null || (column >= who.edit.c1 && column <= who.edit.c2));

    const budget = {remaining: DIFF_OPTIONS.maxLcsCellsPerEdit};
    const records = [];
    const groups = new Map();   // スナップショットを取る行（イベントID・編集者ごと）
    const isEmpty = row => row.every(token => !token);
    const summary = row => {
      const text = row
        .map((token, c) => token ? columnNames[c] + ': ' + diffDecode_(token, timezone).text : '')
        .filter(Boolean).join('\n');
      return text.length > DIFF_ROW_OPTIONS.maxSummaryCharacters
        ? text.slice(0, DIFF_ROW_OPTIONS.maxSummaryCharacters) + '…（省略）'
        : text;
    };
    const markRow = (eventId, editor, rowNumber, cell) => {
      const key = eventId + '\u0001' + editor;
      if (!groups.has(key)) groups.set(key, {eventId, editor, cellsByRow: new Map()});
      const cells = groups.get(key).cellsByRow;
      if (!cells.has(rowNumber)) cells.set(rowNumber, new Set());
      cells.get(rowNumber).add(cell);
    };

    matched.deleted.forEach(oldIndex => {
      const row = oldRows[oldIndex];
      if (isEmpty(row)) return;
      const text = summary(row);
      records.push({
        eventId: rowEventId,
        editorEmail: who.structure,
        sheet: sheet.getName(),
        cell: (rowStart + oldIndex) + '行（削除前の行番号）',
        columnName: '（行全体）',
        kind: DIFF_ROW_OPTIONS.kindDelete,
        before: text,
        after: '',
        parts: [{kind: 'del', text}]
      });
      result.deleted++;
    });

    // 組になった行 = 同じ行（移動していても）。変わったセルだけを記録する
    matched.pairs.forEach(([oldIndex, newIndex]) => {
      const before = oldRows[oldIndex];
      const after = newRows[newIndex];
      const rowNumber = rowStart + newIndex;
      after.forEach((token, c) => {
        if (token === before[c]) return;
        const b = diffDecode_(before[c], timezone);
        const a = diffDecode_(token, timezone);
        const diff = diffCharacters_(b.text, a.text, budget);
        const cell = diffA1_(rowNumber, columnNumbers[c]);
        const edited = inEdit(rowNumber, columnNumbers[c]);
        const eventId = edited ? editEventId : rowEventId;
        const editor = edited ? who.changes : otherEditor;
        records.push({
          eventId,
          editorEmail: editor,
          sheet: sheet.getName(),
          cell,
          columnName: columnNames[c],
          kind: b.type + ' -> ' + a.type + (diff.coarse ? ' / ブロック差分' : '') +
            (edited ? '' : DIFF_ROW_OPTIONS.kindSuffix),
          before: b.text,
          after: a.text,
          parts: diff.parts
        });
        markRow(eventId, editor, rowNumber, cell);
        result.changedCells++;
      });
    });

    matched.inserted.forEach(newIndex => {
      const row = newRows[newIndex];
      if (isEmpty(row)) return;
      const text = summary(row);
      const rowNumber = rowStart + newIndex;
      const editor = inEdit(rowNumber) ? who.changes : otherEditor;
      records.push({
        eventId: rowEventId,
        editorEmail: editor,
        sheet: sheet.getName(),
        cell: rowNumber + '行',
        columnName: '（行全体）',
        kind: DIFF_ROW_OPTIONS.kindInsert,
        before: '',
        after: text,
        parts: [{kind: 'add', text}]
      });
      markRow(rowEventId, editor, rowNumber, rowNumber + '行');
      result.inserted++;
    });

    if (records.length) {
      const changedRows = new Set();
      groups.forEach(group => group.cellsByRow.forEach((_, rowNumber) => changedRows.add(rowNumber)));
      const sortedRows = Array.from(changedRows).sort((a, b) => a - b);
      let schema = null;
      if (sortedRows.length) {
        schema = diffSnapshotSchema_(sheet, rule);
        diffWriteTimestamps_(sheet, sortedRows, stampColumn, now);
        SpreadsheetApp.flush();
      }
      diffAppend_(ss, records, now);
      if (sortedRows.length) {
        if (sortedRows.length * schema.items.length <= DIFF_OPTIONS.maxSnapshotCellsPerEdit) {
          groups.forEach(group => {
            const rows = Array.from(group.cellsByRow.keys()).sort((a, b) => a - b);
            diffAppendRowSnapshots_(
              ss, sheet, rule, rows, group.cellsByRow, group.eventId, group.editor, now, schema);
          });
        } else {
          console.warn('Row snapshot skipped (too many rows): ' + sortedRows.length);
        }
      }
    }

    // 比較基準を今の行で取り直す（読み込んだ値をそのまま使う）
    baseline.clearContents();
    diffEnsureGrid_(baseline, Math.max(newEnd, 1), maxColumn);
    let offset = 0;
    specs.forEach(spec => {
      const skip = spec.r1 - rowStart;
      const count = Math.max(0, newCount - skip);
      const step = Math.max(1, Math.floor(2000 / spec.cols));
      for (let done = 0; done < count; done += step) {
        const n = Math.min(step, count - done);
        const values = newRows.slice(skip + done, skip + done + n)
          .map(row => row.slice(offset, offset + spec.cols));
        baseline.getRange(spec.r1 + done, spec.col, n, spec.cols).setValues(values);
      }
      offset += spec.cols;
    });
  }

  const fresh = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
  if (fresh[sheetId]) {
    fresh[sheetId].gridRows = sheet.getMaxRows();
    fresh[sheetId].paused = false;
    props.setProperty(DIFF_OPTIONS.stateKey, JSON.stringify(fresh));
  }
  SpreadsheetApp.flush();
  console.log('Row resync done: ' + sheet.getName() + ' / ' + JSON.stringify(result));
  return result;
}

/**
 * 行の突き合わせ（diffAlignRows_）で一致しなかった行を組にする。
 * 1. 見分け列の値が同じ行（その値の行が、削除側・追加側とも1行だけ）は、離れていても同じ行とみなす
 *    （移動・並べ替え・行がずれたあとの編集）
 * 2. 残りは、すき間ごとに同じ数なら位置で組にする（行はそのままで中身が変わった）。数が違えば削除・追加
 * 戻り値 {pairs: [[旧の番号, 新の番号]]（新の番号の順）, deleted: [旧の番号], inserted: [新の番号]}
 */
function diffPairRows_(gaps, oldRows, newRows, keyIndexes) {
  const keyOf = row => keyIndexes.length && keyIndexes.some(i => row[i])
    ? keyIndexes.map(i => row[i]).join('\u0001')
    : '';
  const tally = (indexes, rows) => {
    const counts = new Map();
    indexes.forEach(i => {
      const key = keyOf(rows[i]);
      if (key) counts.set(key, (counts.get(key) || 0) + 1);
    });
    return counts;
  };
  const oldAll = [].concat(...gaps.map(g => g.old));
  const newAll = [].concat(...gaps.map(g => g.new));
  const oldCounts = tally(oldAll, oldRows);
  const newCounts = tally(newAll, newRows);

  const oldByKey = new Map();
  oldAll.forEach(i => {
    const key = keyOf(oldRows[i]);
    if (key && oldCounts.get(key) === 1 && newCounts.get(key) === 1) oldByKey.set(key, i);
  });

  const pairs = [];
  const pairedOld = new Set();
  const pairedNew = new Set();
  newAll.forEach(j => {
    const i = oldByKey.get(keyOf(newRows[j]));
    if (i === undefined) return;
    pairs.push([i, j]);
    pairedOld.add(i);
    pairedNew.add(j);
  });

  const deleted = [];
  const inserted = [];
  gaps.forEach(gap => {
    const o = gap.old.filter(i => !pairedOld.has(i));
    const n = gap.new.filter(j => !pairedNew.has(j));
    if (o.length === n.length) {
      o.forEach((i, k) => pairs.push([i, n[k]]));
    } else {
      o.forEach(i => deleted.push(i));
      n.forEach(j => inserted.push(j));
    }
  });

  pairs.sort((a, b) => a[1] - b[1]);
  return {pairs, deleted, inserted};
}

/**
 * 行ごとの比較用トークンを読む。specs は追跡範囲の列のまとまり。
 * 戻り値は [行][追跡している列を左から順に] のトークン。
 */
function diffReadRowTokens_(target, isBaseline, specs, rowStart, count) {
  const rows = Array.from({length: count}, () => []);
  specs.forEach(spec => {
    const skip = Math.max(0, spec.r1 - rowStart);
    for (let i = 0; i < Math.min(skip, count); i++) {
      for (let c = 0; c < spec.cols; c++) rows[i].push('');
    }
    const n = Math.max(0, count - skip);
    const step = Math.max(1, Math.floor(DIFF_ROW_OPTIONS.readChunkCells / spec.cols));
    for (let done = 0; done < n; done += step) {
      const k = Math.min(step, n - done);
      const range = target.getRange(spec.r1 + done, spec.col, k, spec.cols);
      const values = isBaseline
        ? range.getValues().map(row => row.map(v => String(v || '')))
        : diffTokens_(range);
      values.forEach((row, j) => row.forEach(token => rows[skip + done + j].push(token)));
    }
  });
  return rows;
}

/**
 * 行の突き合わせ。前後の一致部分を除いた範囲で LCS を取り、
 * 一致しなかった行を「すき間」ごとに返す：[{old: [旧行の番号], new: [新行の番号]}]
 */
function diffAlignRows_(a, b) {
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let s = 0;
  while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;

  const midA = a.length - p - s;
  const midB = b.length - p - s;
  const gaps = [];
  if (!midA && !midB) return gaps;

  const range = (from, count) => Array.from({length: count}, (_, i) => from + i);
  if ((midA + 1) * (midB + 1) > DIFF_ROW_OPTIONS.maxAlignCells) {
    gaps.push({old: range(p, midA), new: range(p, midB)});
    return gaps;
  }

  const table = Array.from({length: midA + 1}, () => new Uint32Array(midB + 1));
  for (let i = midA - 1; i >= 0; i--) {
    for (let j = midB - 1; j >= 0; j--) {
      table[i][j] = a[p + i] === b[p + j] ? table[i + 1][j + 1] + 1 :
        Math.max(table[i + 1][j], table[i][j + 1]);
    }
  }

  let i = 0, j = 0;
  let current = {old: [], new: []};
  const flush = () => {
    if (current.old.length || current.new.length) gaps.push(current);
    current = {old: [], new: []};
  };
  while (i < midA || j < midB) {
    if (i < midA && j < midB && a[p + i] === b[p + j]) {
      flush(); i++; j++;
    } else if (i < midA && (j === midB || table[i + 1][j] >= table[i][j + 1])) {
      current.old.push(p + i++);
    } else {
      current.new.push(p + j++);
    }
  }
  flush();
  return gaps;
}


/* ---------------- Core processing ---------------- */

/**
 * Validates state/configuration for a tracked sheet. Call while holding the lock.
 * Throws if tracking is paused or the structure/configuration changed.
 */
function diffContext_(ss, sheet, rule) {
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
  const saved = state[sheet.getSheetId()];

  if (!saved || saved.paused) {
    throw new Error('Tracking is not initialized or is paused. Run setupDiffTracking().');
  }

  const stampColumn = diffStampColumn_(sheet, rule);

  if (saved.name !== sheet.getName() ||
      saved.signature !== JSON.stringify(rule) ||
      saved.stampColumn !== stampColumn ||
      saved.gridRows !== sheet.getMaxRows() ||
      saved.gridColumns !== sheet.getMaxColumns()) {
    throw new Error('Structure/configuration changed. Reinitialize the baseline before editing.');
  }

  const baseline = ss.getSheets().find(s => s.getSheetId() === saved.snapshotId);
  if (!baseline) throw new Error('Comparison snapshot is missing.');

  // Validate both output sheets before changing the source timestamp.
  diffLogSheet_(ss);
  diffRowSnapshotSheet_(ss);

  return {
    ss,
    sheet,
    rule,
    stampColumn,
    baseline,
    timezone: ss.getSpreadsheetTimeZone()
  };
}

/**
 * Compares blocks with the baseline, writes timestamps, log records and row
 * snapshots, then advances the baseline. Returns the number of changed cells.
 * meta = {now, eventId, editorEmail, kindSuffix}
 */
function diffProcessBlocks_(ctx, blocks, meta) {
  const {ss, sheet, rule, stampColumn, baseline, timezone} = ctx;

  const changedRows = new Set();
  const changedCellsByRow = new Map();
  const records = [];
  const writes = [];
  const seen = new Set();
  const budget = {remaining: DIFF_OPTIONS.maxLcsCellsPerEdit};

  for (const block of blocks) {
    const [row, col, countRows, countCols] = block;
    const source = sheet.getRange(...block);

    if (source.getMergedRanges().length) {
      throw new Error('Merged cells are not supported: ' + source.getA1Notation());
    }

    const current = diffTokens_(source);
    const columnNames = diffColumnNames_(sheet, rule, col, countCols);

    diffEnsureGrid_(baseline, row + countRows - 1, col + countCols - 1);
    const previous = baseline.getRange(...block).getValues();
    let blockChanged = false;

    for (let r = 0; r < countRows; r++) {
      for (let c = 0; c < countCols; c++) {
        const absoluteRow = row + r;
        const cell = diffA1_(absoluteRow, col + c);

        if (seen.has(cell)) continue;
        seen.add(cell);

        const beforeToken = String(previous[r][c] || '');
        const afterToken = current[r][c];

        // Same value: no diff, no timestamp, no row snapshot.
        if (beforeToken === afterToken) continue;
        blockChanged = true;

        const before = diffDecode_(beforeToken, timezone);
        const after = diffDecode_(afterToken, timezone);
        const result = diffCharacters_(before.text, after.text, budget);

        records.push({
          eventId: meta.eventId,
          editorEmail: meta.editorEmail,
          sheet: sheet.getName(),
          cell,
          columnName: columnNames[c],
          kind: before.type + ' -> ' + after.type +
            (result.coarse ? ' / ブロック差分' : '') +
            (meta.kindSuffix || ''),
          before: before.text,
          after: after.text,
          parts: result.parts
        });

        changedRows.add(absoluteRow);

        if (!changedCellsByRow.has(absoluteRow)) {
          changedCellsByRow.set(absoluteRow, new Set());
        }
        changedCellsByRow.get(absoluteRow).add(cell);
      }
    }

    if (blockChanged) writes.push({block, values: current});
  }

  if (records.length) {
    const sortedRows = Array.from(changedRows).sort((a, b) => a - b);

    // Check expected row-snapshot volume before writing any logs/timestamps.
    const schema = diffSnapshotSchema_(sheet, rule);
    const snapshotCells = sortedRows.length * schema.items.length;
    if (snapshotCells > DIFF_OPTIONS.maxSnapshotCellsPerEdit) {
      throw new Error(
        'Row snapshot exceeds ' +
        DIFF_OPTIONS.maxSnapshotCellsPerEdit +
        ' cells: ' + snapshotCells
      );
    }

    // Timestamp first so the row snapshot contains this timestamp.
    diffWriteTimestamps_(sheet, sortedRows, stampColumn, meta.now);
    SpreadsheetApp.flush();

    diffAppend_(ss, records, meta.now);
    diffAppendRowSnapshots_(
      ss,
      sheet,
      rule,
      sortedRows,
      changedCellsByRow,
      meta.eventId,
      meta.editorEmail,
      meta.now,
      schema
    );
  }

  // Advance baseline only after logs/snapshots have been written.
  for (const write of writes) {
    baseline.getRange(...write.block).setValues(write.values);
  }

  return records.length;
}

/** Splits blocks into row chunks of at most about maxCells cells. */
function diffSplitBlocks_(blocks, maxCells) {
  const out = [];
  for (const [row, col, rows, cols] of blocks) {
    const step = Math.max(1, Math.floor(maxCells / cols));
    for (let offset = 0; offset < rows; offset += step) {
      out.push([row + offset, col, Math.min(step, rows - offset), cols]);
    }
  }
  return out;
}

function diffMarkDirty_(sheetId) {
  PropertiesService.getDocumentProperties()
    .setProperty(DIFF_OPTIONS.dirtyPrefix + sheetId, Utilities.getUuid());
}

function diffClearDirty_(sheetId) {
  PropertiesService.getDocumentProperties()
    .deleteProperty(DIFF_OPTIONS.dirtyPrefix + sheetId);
}

/** 「行の確かめ」の印（行数・ずれだけを取りこぼし回収で確かめる。全セルは比べない）。 */
function diffMarkRowCheck_(sheetId) {
  PropertiesService.getDocumentProperties()
    .setProperty(DIFF_ROW_OPTIONS.checkPrefix + sheetId, Utilities.getUuid());
}

function diffClearRowCheck_(sheetId) {
  PropertiesService.getDocumentProperties()
    .deleteProperty(DIFF_ROW_OPTIONS.checkPrefix + sheetId);
}

function diffPause_(ss, sheetId) {
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
  for (const id of Object.keys(state)) {
    if (sheetId == null || id === String(sheetId)) state[id].paused = true;
  }
  props.setProperty(DIFF_OPTIONS.stateKey, JSON.stringify(state));
}

function diffStampColumn_(sheet, rule) {
  if (!Number.isInteger(rule.headerRow) || rule.headerRow < 1 ||
      !Array.isArray(rule.ranges) || !rule.ranges.length) {
    throw new Error('Invalid rule: ' + sheet.getName());
  }

  const headers = sheet
    .getRange(rule.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getDisplayValues()[0];

  const target = diffNormalizeStampHeader_(DIFF_OPTIONS.stampHeader);

  const matches = headers
    .map((value, index) =>
      diffNormalizeStampHeader_(value) === target ? index + 1 : null)
    .filter(value => value !== null);

  if (matches.length !== 1) {
    throw new Error(
      'Exactly one timestamp header is required: ' +
      sheet.getName() +
      ' / found=' + matches.length +
      ' / headerRow=' + rule.headerRow
    );
  }

  return matches[0];
}

function diffNormalizeStampHeader_(value) {
  return String(value == null ? '' : value)
    .replace(/\s/g, '')
    .replace(/（自動入力）/g, '')
    .replace(/\(自動入力\)/g, '')
    .trim();
}

/** Read labels once per block; avoid one header read per changed cell. */
function diffColumnNames_(sheet, rule, firstColumn, columnCount) {
  const labels = sheet.getRange(rule.headerRow, firstColumn, 1, columnCount)
    .getDisplayValues()[0];
  return labels.map((label, offset) => {
    const text = label.trim();
    const letter = diffColumnLetter_(firstColumn + offset);
    return text || letter + '列（見出しなし）';
  });
}

/** Return watched rectangles, excluding headers and timestamp column. */
function diffBlocks_(sheet, rule, edit, stampColumn, usedOnly) {
  const blocks = [];
  for (const a1 of rule.ranges) {
    const target = sheet.getRange(a1);
    const r1 = Math.max(target.getRow(), rule.headerRow + 1, edit ? edit.getRow() : 1);
    const r2 = Math.min(target.getLastRow(), edit ? edit.getLastRow() : sheet.getMaxRows(),
      usedOnly ? sheet.getLastRow() : sheet.getMaxRows());
    const c1 = Math.max(target.getColumn(), edit ? edit.getColumn() : 1);
    const c2 = Math.min(target.getLastColumn(), edit ? edit.getLastColumn() : sheet.getMaxColumns());
    if (r1 > r2 || c1 > c2) continue;
    for (const [left, right] of [[c1, Math.min(c2, stampColumn - 1)],
      [Math.max(c1, stampColumn + 1), c2]]) {
      if (left <= right) blocks.push([r1, left, r2 - r1 + 1, right - left + 1]);
    }
  }
  return blocks;
}

function diffTokens_(range) {
  const values = range.getValues();
  const formulas = range.getFormulas();
  return values.map((row, r) => row.map((value, c) => {
    let token;
    if (formulas[r][c]) token = 'F' + formulas[r][c];
    else if (value === '') token = '';
    else if (value instanceof Date) token = 'D' + value.getTime();
    else if (typeof value === 'number') token = 'N' + String(value);
    else if (typeof value === 'boolean') token = 'B' + (value ? 'TRUE' : 'FALSE');
    else token = 'S' + String(value);
    if (token.length > DIFF_OPTIONS.maxTokenLength) {
      throw new Error('Cell text exceeds the snapshot limit at ' + range.getA1Notation());
    }
    return token; // Nonempty tokens NEVER begin with '=': no formula execution.
  }));
}

function diffDecode_(token, timezone) {
  if (!token) return {type: '空欄', text: ''};
  const types = {S: '文字列', N: '数値', B: '真偽値', D: '日時', F: '数式'};
  const kind = token[0];
  const value = token.slice(1);
  return {
    type: types[kind] || '不明',
    text: kind === 'D' ? Utilities.formatDate(new Date(Number(value)), timezone,
      DIFF_OPTIONS.dateFormat) : value
  };
}

/** LCS over Unicode code points; style offsets below use UTF-16 string lengths. */
function diffCharacters_(before, after, budget) {
  const a = Array.from(before), b = Array.from(after);
  let prefix = 0, suffix = 0;
  while (prefix < a.length && prefix < b.length && a[prefix] === b[prefix]) prefix++;
  while (suffix < a.length - prefix && suffix < b.length - prefix &&
    a[a.length - suffix - 1] === b[b.length - suffix - 1]) suffix++;
  const left = a.slice(prefix, a.length - suffix);
  const right = b.slice(prefix, b.length - suffix);
  const parts = [];
  const push = (kind, text) => {
    if (!text) return;
    const last = parts[parts.length - 1];
    if (last && last.kind === kind) last.text += text;
    else parts.push({kind, text});
  };
  push('same', a.slice(0, prefix).join(''));
  const cost = (left.length + 1) * (right.length + 1);
  const coarse = left.length > 0 && right.length > 0 &&
    (cost > DIFF_OPTIONS.maxLcsCells || cost > budget.remaining);
  if (!left.length || !right.length || coarse) {
    push('del', left.join(''));
    push('add', right.join(''));
  } else {
    budget.remaining -= cost;
    const table = Array.from({length: left.length + 1}, () => new Uint32Array(right.length + 1));
    for (let i = left.length - 1; i >= 0; i--) {
      for (let j = right.length - 1; j >= 0; j--) {
        table[i][j] = left[i] === right[j] ? table[i + 1][j + 1] + 1 :
          Math.max(table[i + 1][j], table[i][j + 1]);
      }
    }
    let i = 0, j = 0;
    while (i < left.length || j < right.length) {
      if (i < left.length && j < right.length && left[i] === right[j]) {
        push('same', left[i]); i++; j++;
      } else if (i < left.length && (j === right.length || table[i + 1][j] >= table[i][j + 1])) {
        push('del', left[i++]);
      } else {
        push('add', right[j++]);
      }
    }
  }
  push('same', suffix ? a.slice(a.length - suffix).join('') : '');
  return {parts, coarse};
}

function diffPlain_(text) {
  return SpreadsheetApp.newRichTextValue().setText(text || '（空欄）')
    .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('#000000')
      .setBold(false).setStrikethrough(false).build()).build();
}


/* ---------------- Difference log ---------------- */

function diffLogSheet_(ss) {
  let log = ss.getSheetByName(DIFF_OPTIONS.logSheet);

  if (!log) {
    log = ss.insertSheet(DIFF_OPTIONS.logSheet);
    diffEnsureGrid_(log, 1, DIFF_HEADERS.length);
    log.getRange(1, 1, 1, DIFF_HEADERS.length).setValues([DIFF_HEADERS]);
    diffFormatDiffHeader_(log);
    diffTrimGrid_(log, DIFF_HEADERS.length);   // 使わない右の列・下の行を持たない（セル数の節約）
    return log;
  }

  diffEnsureGrid_(log, 1, DIFF_HEADERS.length);

  // 前の版の列構成（8・9・10列）も読めるよう、右の列まで読む
  const width = Math.min(log.getMaxColumns(), Math.max(log.getLastColumn(), DIFF_LEGACY_HEADERS_10.length));
  const actual = log.getRange(1, 1, 1, width).getDisplayValues()[0];

  // 今の9列（前の版の11列で「差分」「確認済み」が右に残っているときも、左の9列は同じ）
  if (diffHeadersEqual_(actual, DIFF_HEADERS)) {
    return log;
  }

  // Previous 10-column version: insert editor email after Event ID.
  if (diffHeadersEqual_(actual, DIFF_LEGACY_HEADERS_10)) {
    log.insertColumnAfter(2);
    log.getRange(1, 1, 1, DIFF_HEADERS.length).setValues([DIFF_HEADERS]);
    diffFormatDiffHeader_(log);
    return log;
  }

  // Previous 9-column version: insert Event ID and editor email.
  if (diffHeadersEqual_(actual, DIFF_LEGACY_HEADERS_9)) {
    log.insertColumnAfter(1);
    log.insertColumnAfter(2);
    log.getRange(1, 1, 1, DIFF_HEADERS.length).setValues([DIFF_HEADERS]);
    diffFormatDiffHeader_(log);
    return log;
  }

  // Oldest 8-column version: insert column name, Event ID, and editor email.
  if (diffHeadersEqual_(actual, DIFF_LEGACY_HEADERS_8)) {
    log.insertColumnAfter(3);
    log.insertColumnAfter(1);
    log.insertColumnAfter(2);
    log.getRange(1, 1, 1, DIFF_HEADERS.length).setValues([DIFF_HEADERS]);
    diffFormatDiffHeader_(log);
    return log;
  }

  throw new Error(
    'Log sheet has unexpected headers. Back it up and restore the expected layout before retrying.'
  );
}

function diffFormatDiffHeader_(log) {
  log.getRange(1, 1, 1, DIFF_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#e8eef5');

  log.setFrozenRows(1);
  log.setColumnWidth(1, 160);
  log.setColumnWidth(2, 210);
  log.setColumnWidth(3, 240);
  log.setColumnWidth(4, 130);
  log.setColumnWidth(5, 80);
  log.setColumnWidth(6, 180);
  log.setColumnWidth(7, 180);
  log.setColumnWidths(8, 2, 300);
}

function diffAppend_(ss, records, now) {
  const log = diffLogSheet_(ss);
  const row = log.getLastRow() + 1;

  diffEnsureGrid_(log, row + records.length - 1, DIFF_HEADERS.length);

  log.getRange(row, 1, records.length, 1)
    .setValues(records.map(() => [now]))
    .setNumberFormat(DIFF_OPTIONS.dateFormat);

  // Rich text prevents strings beginning with '=' from becoming executable formulas.
  log.getRange(row, 2, records.length, DIFF_HEADERS.length - 1).setRichTextValues(records.map(rec => [
    diffPlain_(rec.eventId),
    diffPlain_(rec.editorEmail),
    diffPlain_(rec.sheet),
    diffPlain_(rec.cell),
    diffPlain_(rec.columnName),
    diffPlain_(rec.kind),
    diffPlain_(rec.before),
    diffPlain_(rec.after)
  ]));

  log.getRange(row, 1, records.length, DIFF_HEADERS.length)
    .setVerticalAlignment('top')
    .setWrap(true);
}


/* ---------------- Row snapshot log ---------------- */

/**
 * 変更時点スナップショットのシート。列は、固定の6列（ROW_SNAPSHOT_FIXED_HEADERS）・得意先・行の内容（JSON）の8列。
 * 前の版（固定の列のあとに見出しごとの列が並ぶ）のシートには、得意先・行の内容の列が無ければ右端に足し、
 * 新しい行はこの2列だけに書く（compactDiffRecords() で前の版の行も移し替え、見出しごとの列を削除する）。
 * 戻り値 {sheet, customerColumn, contentColumn}
 */
function diffRowSnapshotSheet_(ss) {
  let sheet = ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet);
  const tail = [diffSnapshotCustomerHeader_(), DIFF_OPTIONS.snapshotContentHeader];

  if (!sheet) {
    const headers = ROW_SNAPSHOT_FIXED_HEADERS.concat(tail);
    sheet = ss.insertSheet(DIFF_OPTIONS.rowSnapshotSheet);
    diffEnsureGrid_(sheet, 1, headers.length);
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#e8eef5');

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 210);
    sheet.setColumnWidth(3, 240);
    sheet.setColumnWidth(4, 130);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 180);
    sheet.setColumnWidth(8, 600);
    diffTrimGrid_(sheet, headers.length);   // 使わない右の列・下の行を持たない（セル数の節約）
    return {sheet, customerColumn: 7, contentColumn: 8};
  }

  diffEnsureGrid_(sheet, 1, ROW_SNAPSHOT_FIXED_HEADERS.length);

  const actual = sheet
    .getRange(1, 1, 1, ROW_SNAPSHOT_FIXED_HEADERS.length)
    .getDisplayValues()[0];

  if (!diffHeadersEqual_(actual, ROW_SNAPSHOT_FIXED_HEADERS)) {
    // Previous 5-column fixed layout: insert editor email after Event ID.
    if (diffHeadersEqual_(actual.slice(0, ROW_SNAPSHOT_LEGACY_HEADERS_5.length),
        ROW_SNAPSHOT_LEGACY_HEADERS_5)) {
      sheet.insertColumnAfter(2);
      sheet.getRange(1, 1, 1, ROW_SNAPSHOT_FIXED_HEADERS.length)
        .setValues([ROW_SNAPSHOT_FIXED_HEADERS])
        .setFontWeight('bold')
        .setBackground('#e8eef5');
      sheet.setColumnWidth(3, 240);
    } else {
      throw new Error(
        'Row snapshot sheet has unexpected fixed headers: ' +
        DIFF_OPTIONS.rowSnapshotSheet
      );
    }
  }

  // 得意先・行の内容の列（無ければ右端に足す）
  const fixed = ROW_SNAPSHOT_FIXED_HEADERS.length;
  const last = Math.max(sheet.getLastColumn(), fixed);
  const headers = sheet.getRange(1, 1, 1, last).getDisplayValues()[0].map(diffCleanHeaderText_);
  const columns = tail.map(header => {
    const index = headers.indexOf(header, fixed);
    return index >= 0 ? index + 1 : 0;
  });
  const missing = tail.filter((header, i) => !columns[i]);
  if (missing.length) {
    const start = sheet.getLastColumn() + 1;
    diffEnsureGrid_(sheet, 1, start + missing.length - 1);
    sheet.getRange(1, start, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setBackground('#e8eef5');
    missing.forEach((header, i) => { columns[tail.indexOf(header)] = start + i; });
  }
  return {sheet, customerColumn: columns[0], contentColumn: columns[1]};
}

/** スナップショットの「得意先」の列の見出し（クレデンシャル・変更履歴・サイドバーが得意先で記録を探すのに使う）。 */
function diffSnapshotCustomerHeader_() {
  return diffCleanHeaderText_(typeof CRED_OPTIONS !== 'undefined' ? CRED_OPTIONS.customerHeader : '得意先');
}

/** スナップショットに入れない列（ボタン列。チェックの ON/OFF は行の内容ではない）。 */
function diffSnapshotSkippedHeaders_() {
  const headers = [];
  if (typeof CRED_OPTIONS !== 'undefined') headers.push(CRED_OPTIONS.launcherHeader);
  if (typeof CHG_OPTIONS !== 'undefined') headers.push(CHG_OPTIONS.launcherHeader);
  return new Set(headers.map(diffCleanHeaderText_));
}

/**
 * 「行の内容」のセルに入れる文字（[[見出し, 値], …] の JSON。空欄の項目は持たない）。
 * 1つのセルの上限を超えるときは、長い値から省略する。
 */
function diffSnapshotContent_(pairs) {
  let json = JSON.stringify(pairs);
  for (const limit of [2000, 500, 100]) {
    if (json.length <= DIFF_OPTIONS.snapshotContentLimit) break;
    json = JSON.stringify(pairs.map(([key, value]) =>
      [key, value.length > limit ? value.slice(0, limit) + '…（省略）' : value]));
  }
  return json;
}

/** スナップショットの行の「行の内容」を [[見出し, 値], …] にする（読めないときは空）。 */
function diffParseSnapshotContent_(text) {
  if (!text) return [];
  try {
    const pairs = JSON.parse(String(text));
    return Array.isArray(pairs) ? pairs.filter(p => Array.isArray(p) && p.length === 2).map(([k, v]) => [String(k), String(v)]) : [];
  } catch (_) {
    return [];
  }
}

function diffAppendRowSnapshots_(
    ss, sourceSheet, rule, rows, changedCellsByRow, eventId, editorEmail, now, schema) {
  if (!rows.length) return;

  const {sheet: snapshotSheet, customerColumn, contentColumn} = diffRowSnapshotSheet_(ss);
  schema = schema || diffSnapshotSchema_(sourceSheet, rule);

  const snapshotCells = rows.length * schema.items.length;
  if (snapshotCells > DIFF_OPTIONS.maxSnapshotCellsPerEdit) {
    throw new Error(
      'Row snapshot exceeds ' +
      DIFF_OPTIONS.maxSnapshotCellsPerEdit +
      ' cells: ' + snapshotCells
    );
  }

  const startRow = snapshotSheet.getLastRow() + 1;
  diffEnsureGrid_(snapshotSheet, startRow + rows.length - 1, Math.max(customerColumn, contentColumn));

  const metadata = rows.map(rowNumber => [
    now,
    eventId,
    editorEmail,
    sourceSheet.getName(),
    rowNumber,
    Array.from(changedCellsByRow.get(rowNumber) || []).join(', ')
  ]);

  snapshotSheet
    .getRange(startRow, 1, rows.length, ROW_SNAPSHOT_FIXED_HEADERS.length)
    .setValues(metadata);

  snapshotSheet
    .getRange(startRow, 1, rows.length, 1)
    .setNumberFormat(DIFF_OPTIONS.dateFormat);

  const firstRow = rows[0];
  const contiguous = rows.every((value, index) => value === firstRow + index);
  const rowValues = new Map();

  if (contiguous) {
    const block = sourceSheet
      .getRange(firstRow, schema.firstColumn, rows.length, schema.columnCount)
      .getDisplayValues();

    rows.forEach((rowNumber, i) => rowValues.set(rowNumber, block[i]));
  } else {
    for (const rowNumber of rows) {
      rowValues.set(
        rowNumber,
        sourceSheet
          .getRange(rowNumber, schema.firstColumn, 1, schema.columnCount)
          .getDisplayValues()[0]
      );
    }
  }

  // 得意先と、行の内容（見出しと値。空欄・ボタン列は持たない）。リッチテキストで書くので数式にならない
  const customerKey = diffSnapshotCustomerHeader_();
  const skipped = diffSnapshotSkippedHeaders_();
  const tail = rows.map(rowNumber => {
    const values = rowValues.get(rowNumber);
    const pairs = [];
    let customer = '';
    schema.items.forEach((item, index) => {
      const value = String(values[index] == null ? '' : values[index]);
      if (item.key === customerKey) customer = value;
      if (value !== '' && !skipped.has(item.key)) pairs.push([item.key, value]);
    });
    return {customer, content: diffSnapshotContent_(pairs)};
  });

  if (contentColumn === customerColumn + 1) {
    snapshotSheet.getRange(startRow, customerColumn, rows.length, 2)
      .setRichTextValues(tail.map(t => [diffSnapshotText_(t.customer), diffSnapshotText_(t.content)]));
  } else {
    snapshotSheet.getRange(startRow, customerColumn, rows.length, 1)
      .setRichTextValues(tail.map(t => [diffSnapshotText_(t.customer)]));
    snapshotSheet.getRange(startRow, contentColumn, rows.length, 1)
      .setRichTextValues(tail.map(t => [diffSnapshotText_(t.content)]));
  }

  snapshotSheet
    .getRange(startRow, 1, rows.length, ROW_SNAPSHOT_FIXED_HEADERS.length)
    .setVerticalAlignment('top');
}

/**
 * Snapshot columns:
 * - rule.snapshotColumns = 'A:Z' can explicitly define them.
 * - if omitted, A through the last non-empty header cell are captured.
 */
function diffSnapshotSchema_(sheet, rule) {
  const bounds = diffSnapshotColumnBounds_(sheet, rule);
  const headers = sheet
    .getRange(rule.headerRow, bounds.firstColumn, 1, bounds.columnCount)
    .getDisplayValues()[0];

  const cleaned = headers.map(diffCleanHeaderText_);
  const counts = {};

  cleaned.forEach(header => {
    if (header) counts[header] = (counts[header] || 0) + 1;
  });

  const items = cleaned.map((header, index) => {
    const column = bounds.firstColumn + index;
    const letter = diffColumnLetter_(column);
    let key;

    if (!header) {
      key = letter + '列（見出しなし）';
    } else if (counts[header] > 1) {
      key = header + ' [' + letter + ']';
    } else {
      key = header;
    }

    return {key, column};
  });

  return {
    firstColumn: bounds.firstColumn,
    columnCount: bounds.columnCount,
    items
  };
}

function diffSnapshotColumnBounds_(sheet, rule) {
  if (rule.snapshotColumns) {
    const range = sheet.getRange(rule.snapshotColumns);
    return {
      firstColumn: range.getColumn(),
      columnCount: range.getNumColumns()
    };
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet
    .getRange(rule.headerRow, 1, 1, lastColumn)
    .getDisplayValues()[0];

  let finalColumn = 1;

  for (let i = headers.length - 1; i >= 0; i--) {
    if (diffCleanHeaderText_(headers[i])) {
      finalColumn = i + 1;
      break;
    }
  }

  return {firstColumn: 1, columnCount: finalColumn};
}



/* ---------------- Event/timestamp helpers ---------------- */

function diffWriteTimestamps_(sheet, sortedRows, stampColumn, now) {
  for (let i = 0; i < sortedRows.length;) {
    let j = i;

    while (j + 1 < sortedRows.length &&
           sortedRows[j + 1] === sortedRows[j] + 1) {
      j++;
    }

    sheet
      .getRange(sortedRows[i], stampColumn, j - i + 1, 1)
      .setValue(now)
      .setNumberFormat(DIFF_OPTIONS.dateFormat);

    i = j + 1;
  }
}

/**
 * Returns the editing user's email address when Google exposes it to the trigger.
 * Never uses Session.getEffectiveUser(), because that would return the trigger owner.
 */
function diffEditorEmail_(e) {
  try {
    if (e && e.user) {
      const email = e.user.getEmail();
      if (email) return email;
    }
  } catch (_) {}

  try {
    const email = Session.getActiveUser().getEmail();
    if (email) return email;
  } catch (_) {}

  return '取得不可';
}

/** prefix: 'EVT' for edit events, 'REC' for catch-up scans, 'ROW' for row insert/delete. */
function diffCreateEventId_(now, timezone, prefix) {
  const timestamp = Utilities.formatDate(now, timezone, 'yyyyMMdd-HHmmss');
  const shortUuid = Utilities.getUuid()
    .replace(/-/g, '')
    .slice(0, 8)
    .toUpperCase();

  return (prefix || 'EVT') + '-' + timestamp + '-' + shortUuid;
}

/** Ensures exactly one of each diff trigger exists for this account. */
function diffInstallTriggers_(ss) {
  const specs = [
    {
      handler: 'recordDiffEdit',
      match: t => t.getTriggerSourceId() === ss.getId(),
      create: () => ScriptApp.newTrigger('recordDiffEdit')
        .forSpreadsheet(ss).onEdit().create()
    },
    {
      handler: 'pauseDiffOnStructureChange',
      match: t => t.getTriggerSourceId() === ss.getId(),
      create: () => ScriptApp.newTrigger('pauseDiffOnStructureChange')
        .forSpreadsheet(ss).onChange().create()
    },
    {
      handler: 'catchUpDiffTracking',
      match: t => t.getEventType() === ScriptApp.EventType.CLOCK,
      create: () => ScriptApp.newTrigger('catchUpDiffTracking')
        .timeBased().everyMinutes(DIFF_OPTIONS.catchUpMinutes).create()
    }
  ];

  const triggers = ScriptApp.getProjectTriggers();

  for (const spec of specs) {
    const existing = triggers.filter(t =>
      t.getHandlerFunction() === spec.handler && spec.match(t));

    existing.slice(1).forEach(t => ScriptApp.deleteTrigger(t));

    if (!existing.length) spec.create();
  }
}

function diffCleanHeaderText_(value) {
  return String(value == null ? '' : value)
    .replace(/\s+/g, ' ')
    .trim();
}

function diffHeadersEqual_(actual, expected) {
  if (actual.length < expected.length) return false;

  for (let i = 0; i < expected.length; i++) {
    if (actual[i] !== expected[i]) return false;
  }

  return true;
}

function diffColumnLetter_(column) {
  let letters = '';

  for (let c = column; c > 0; c = Math.floor((c - 1) / 26)) {
    letters =
      String.fromCharCode(65 + ((c - 1) % 26)) +
      letters;
  }

  return letters;
}

function diffSnapshotText_(text) {
  return SpreadsheetApp.newRichTextValue()
    .setText(text == null ? '' : String(text))
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setForegroundColor('#000000')
        .setBold(false)
        .setStrikethrough(false)
        .build()
    )
    .build();
}

/**
 * 使わない右の列・下の行を削除する（空のセルも、スプレッドシートの1,000万セルの上限に数えられるため）。
 * columns より右の列と、rows（省略したときはデータのある最後の行）より下の行。データのある列・行は消さない。
 */
function diffTrimGrid_(sheet, columns, rows) {
  const keepColumns = Math.max(columns || 1, sheet.getLastColumn(), 1);
  if (sheet.getMaxColumns() > keepColumns) sheet.deleteColumns(keepColumns + 1, sheet.getMaxColumns() - keepColumns);
  const keepRows = Math.max(rows || 0, sheet.getLastRow(), sheet.getFrozenRows() + 1, 1);
  if (sheet.getMaxRows() > keepRows) sheet.deleteRows(keepRows + 1, sheet.getMaxRows() - keepRows);
}

/** スプレッドシートのセル数（空のセルも含む。シートの行数×列数の合計）。 */
function diffCountCells_(ss) {
  return ss.getSheets().reduce((n, sheet) => n + sheet.getMaxRows() * sheet.getMaxColumns(), 0);
}

/**
 * シートごとのセル数（空のセルも含む）を、多い順に画面とログに出す（読むだけ）。
 * スプレッドシートは1ファイル1,000万セルまで。
 */
function diagnoseCellUsage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const limit = 10000000;
  const sheets = ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    cells: sheet.getMaxRows() * sheet.getMaxColumns(),
    rows: sheet.getMaxRows(),
    columns: sheet.getMaxColumns()
  })).sort((a, b) => b.cells - a.cells);
  const total = sheets.reduce((n, s) => n + s.cells, 0);
  const lines = ['合計 ' + total.toLocaleString() + ' セル（上限 1,000万セルの ' + (total / limit * 100).toFixed(1) + '%）', '']
    .concat(sheets.map(s => s.name + '：' + s.cells.toLocaleString() + ' セル（' + s.rows + '行 × ' + s.columns + '列）'));
  console.log(lines.join('\n'));
  const ui = SpreadsheetApp.getUi();
  ui.alert('セル数', lines.join('\n'), ui.ButtonSet.OK);
}

/**
 * 差分追跡の記録を、セル数の少ない形に移し替える（管理者がエディタから実行する。何度実行してもよい）。
 * - 変更履歴_差分：前の版の「差分」「確認済み」の列を削除する（文字単位の差分は変更履歴の画面がその場で作る）
 * - 変更時点スナップショット：見出しごとの列の値を「行の内容」（1つのセル）にまとめてから、見出しごとの列を削除する（値は変えない）
 * - 記録のシート・比較基準（__CHAR_DIFF_*）・クレデンシャル履歴_記録：使っていない右の列・下の行を削除する
 * 実行する前に「ファイル」→「コピーを作成」で控えを取ること。行が多くて時間内に終わらないときは途中で止め、
 * もう一度実行すると続きから行う（見出しごとの列は、すべての行を移し替えてから削除する）。
 */
function compactDiffRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('記録の移し替え',
    '「' + DIFF_OPTIONS.logSheet + '」「' + DIFF_OPTIONS.rowSnapshotSheet + '」を、セル数の少ない形に移し替えます（記録の値は変わりません）。\n' +
    '実行する前に「ファイル」→「コピーを作成」で控えを取ってください。\n' +
    '移し替えの間は差分の記録を待たせるので、編集の少ない時間に実行してください。\n\n実行しますか？',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const started = Date.now();
  const before = diffCountCells_(ss);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  let result;
  try {
    const log = diffCompactLog_(ss);
    const snapshot = diffCompactSnapshots_(ss, started);
    if (snapshot.done) diffTrimRecordSheets_(ss);
    result = {log, snapshot};
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
  const after = diffCountCells_(ss);

  const lines = [];
  if (result.log) lines.push('「' + DIFF_OPTIONS.logSheet + '」の「' + DIFF_LEGACY_TAIL_HEADERS.join('」「') + '」の列を削除しました。');
  if (result.snapshot.converted) lines.push('「' + DIFF_OPTIONS.rowSnapshotSheet + '」の ' + result.snapshot.converted + '行を「' + DIFF_OPTIONS.snapshotContentHeader + '」にまとめました。');
  if (result.snapshot.removed) lines.push('見出しごとの列 ' + result.snapshot.removed + '列を削除しました。');
  if (!result.snapshot.done) {
    lines.push('行が多いため途中で止めました。もう一度 compactDiffRecords() を実行してください（続きから行います）。');
  }
  lines.push('セル数：' + before.toLocaleString() + ' → ' + after.toLocaleString() + '（' + (before - after).toLocaleString() + ' 減）');
  ui.alert('記録の移し替え', lines.join('\n'), ui.ButtonSet.OK);
}

/** 変更履歴_差分 の前の版の列（差分・確認済み）を削除する（ロック取得中に呼ぶこと）。削除したら true。 */
function diffCompactLog_(ss) {
  if (!ss.getSheetByName(DIFF_OPTIONS.logSheet)) return false;
  const log = diffLogSheet_(ss);   // 前の版（8・9・10列）なら、先に今の列にそろえる
  const headers = log.getRange(1, 1, 1, Math.max(log.getLastColumn(), 1)).getDisplayValues()[0].map(diffCleanHeaderText_);
  let removed = false;
  for (let c = headers.length; c > DIFF_HEADERS.length; c--) {
    if (DIFF_LEGACY_TAIL_HEADERS.indexOf(headers[c - 1]) < 0) continue;
    log.deleteColumn(c);
    removed = true;
  }
  return removed;
}

/**
 * 変更時点スナップショットの前の版の行（見出しごとの列）を「行の内容」にまとめ、見出しごとの列を削除する（ロック取得中に呼ぶこと）。
 * 時間が足りなくなったら途中で止める（もう一度呼ぶと、「行の内容」が空の行から続ける）。{converted, removed, done}
 */
function diffCompactSnapshots_(ss, started) {
  const result = {converted: 0, removed: 0, done: true};
  if (!ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet)) return result;
  const {sheet, customerColumn, contentColumn} = diffRowSnapshotSheet_(ss);
  const fixed = ROW_SNAPSHOT_FIXED_HEADERS.length;
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(diffCleanHeaderText_);
  const wide = [];   // 見出しごとの列（得意先・行の内容は残す）
  for (let c = fixed + 1; c <= width; c++) {
    if (c !== customerColumn && c !== contentColumn) wide.push(c);
  }
  if (!wide.length) return result;

  // 前の版の行を「行の内容」にまとめる（見出しの順。空欄・ボタン列は持たない。得意先も含める）
  const skipped = diffSnapshotSkippedHeaders_();
  const columns = [];
  for (let c = fixed + 1; c <= width; c++) {
    if (c !== contentColumn && headers[c - 1] && !skipped.has(headers[c - 1])) columns.push(c);
  }
  const chunk = 500;
  for (let row = 2; row <= sheet.getLastRow(); row += chunk) {
    if (Date.now() - started > 270000) {   // 4分半を過ぎたら止める（エディタからの実行は6分まで）
      result.done = false;
      return result;
    }
    const count = Math.min(chunk, sheet.getLastRow() - row + 1);
    const values = sheet.getRange(row, 1, count, width).getDisplayValues();
    let changed = false;
    const contents = values.map(v => {
      const current = v[contentColumn - 1];
      if (current || !v[1]) return current;   // まとめ済み・イベントIDの無い行はそのまま
      changed = true;
      result.converted++;
      return diffSnapshotContent_(columns.filter(c => v[c - 1] !== '').map(c => [headers[c - 1], v[c - 1]]));
    });
    if (changed) {
      sheet.getRange(row, contentColumn, count, 1).setRichTextValues(contents.map(text => [diffSnapshotText_(text)]));
    }
  }

  // すべての行をまとめたので、見出しごとの列を右から削除する（続いている列はまとめて）
  for (let i = wide.length - 1; i >= 0;) {
    let j = i;
    while (j > 0 && wide[j - 1] === wide[j] - 1) j--;
    sheet.deleteColumns(wide[j], i - j + 1);
    result.removed += i - j + 1;
    i = j - 1;
  }
  sheet.setColumnWidth(fixed + 1, 180);
  sheet.setColumnWidth(fixed + 2, 600);
  return result;
}

/** 記録のシート・比較基準・クレデンシャル履歴_記録の、使っていない右の列・下の行を削除する（ロック取得中に呼ぶこと）。 */
function diffTrimRecordSheets_(ss) {
  [DIFF_OPTIONS.logSheet, DIFF_OPTIONS.rowSnapshotSheet,
    typeof CRED_OPTIONS !== 'undefined' ? CRED_OPTIONS.logSheet : ''].forEach(name => {
    const sheet = name && ss.getSheetByName(name);
    if (sheet) diffTrimGrid_(sheet, sheet.getLastColumn(), sheet.getLastRow());   // 記録を足すときは、必要な行だけ足す
  });
  // 比較基準：追跡範囲の右端の列・追跡シートの行数まで
  const state = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DIFF_OPTIONS.stateKey) || '{}');
  ss.getSheets().forEach(sheet => {
    const saved = state[sheet.getSheetId()];
    const rule = saved && DIFF_RULES[sheet.getName()];
    if (!rule) return;
    const baseline = ss.getSheets().find(s => s.getSheetId() === saved.snapshotId);
    if (!baseline) return;
    const right = Math.max(...rule.ranges.map(a1 => sheet.getRange(a1).getLastColumn()));
    diffTrimGrid_(baseline, right, sheet.getMaxRows());
  });
}

function diffEnsureGrid_(sheet, rows, columns) {
  if (rows > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (columns > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
}

function diffA1_(row, column) {
  return diffColumnLetter_(column) + row;
}


/* ---------------- Optional protection ---------------- */

/**
 * Optional: run manually as the administrator.
 * Protects the timestamp column on all configured sheets.
 */
function protectTimestampColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const me = Session.getEffectiveUser();
  const myEmail = me.getEmail();
  const description = 'DIFF_TIMESTAMP_PROTECTION';

  for (const [sheetName, rule] of Object.entries(DIFF_RULES)) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const stampColumn = diffStampColumn_(sheet, rule);
    const startRow = rule.headerRow + 1;
    const rowCount = sheet.getMaxRows() - startRow + 1;
    if (rowCount <= 0) continue;

    sheet
      .getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .filter(p => p.getDescription() === description)
      .forEach(p => p.remove());

    const protection = sheet
      .getRange(startRow, stampColumn, rowCount, 1)
      .protect()
      .setDescription(description);

    protection.addEditor(me);

    const others = protection.getEditors()
      .filter(user => user.getEmail() !== myEmail);

    if (others.length) protection.removeEditors(others);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  }

  ss.toast('最終更新日時列を保護しました。', 'Diff tracking', 8);
}


/* ---------------- Maintenance / diagnostics ---------------- */

/** Deletes and recreates this account's diff triggers (use after re-authorization). */
function reinstallDiffTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (DIFF_TRIGGER_HANDLERS.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  diffInstallTriggers_(ss);

  ss.toast('差分追跡トリガーを再作成しました。', 'Diff tracking', 8);
}

/** Lists this account's project triggers. */
function diagnoseDiffTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  console.log('対象スプレッドシート');
  console.log('名前: ' + ss.getName());
  console.log('ID: ' + ss.getId());
  console.log('実行ユーザー: ' + Session.getEffectiveUser().getEmail());

  const triggers = ScriptApp.getProjectTriggers();

  console.log('トリガー数: ' + triggers.length);

  triggers.forEach((t, i) => {
    console.log(
      [
        'No.' + (i + 1),
        'function=' + t.getHandlerFunction(),
        'event=' + t.getEventType(),
        'source=' + t.getTriggerSource(),
        'sourceId=' + t.getTriggerSourceId()
      ].join(' / ')
    );
  });

  DIFF_TRIGGER_HANDLERS.forEach(handler => {
    const found = triggers.some(t => t.getHandlerFunction() === handler);
    console.log((found ? '○ ' : '✗ ') + handler);
  });
}

/** Checks saved state against the current sheets and DIFF_RULES. */
function diagnoseDiffState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');

  console.log('実行ユーザー: ' + Session.getEffectiveUser().getEmail());
  console.log('保存状態: ' + JSON.stringify(state, null, 2));
  console.log('シート名一覧: ' + ss.getSheets().map(s => '[' + s.getName() + ']').join(', '));

  for (const [name, rule] of Object.entries(DIFF_RULES)) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { console.log('✗ シートが見つかりません: [' + name + ']'); continue; }

    const saved = state[sheet.getSheetId()];
    if (!saved) { console.log('✗ 未初期化: ' + name + ' → setupDiffTracking() を実行'); continue; }

    let stamp;
    try { stamp = diffStampColumn_(sheet, rule); }
    catch (err) { console.log('✗ ' + name + ': ' + err.message); continue; }

    const dirty = props.getProperty(DIFF_OPTIONS.dirtyPrefix + sheet.getSheetId());
    const rowCheck = props.getProperty(DIFF_ROW_OPTIONS.checkPrefix + sheet.getSheetId());

    const checks = [
      ['停止していない', saved.paused === false, saved.paused],
      ['シート名', saved.name === sheet.getName(), saved.name + ' / ' + sheet.getName()],
      ['DIFF_RULES', saved.signature === JSON.stringify(rule), saved.signature + ' / ' + JSON.stringify(rule)],
      ['タイムスタンプ列', saved.stampColumn === stamp, saved.stampColumn + ' / ' + stamp],
      ['行数（違う場合は次の編集・回収で自動反映）', saved.gridRows === sheet.getMaxRows(), saved.gridRows + ' / ' + sheet.getMaxRows()],
      ['列数', saved.gridColumns === sheet.getMaxColumns(), saved.gridColumns + ' / ' + sheet.getMaxColumns()],
      ['比較用シート', ss.getSheets().some(s => s.getSheetId() === saved.snapshotId), saved.snapshotId],
      ['未回収の取りこぼしなし', !dirty, '回収待ち（次回の catchUpDiffTracking で処理）'],
      ['行の確かめ待ちなし', !rowCheck, '確かめ待ち（次回の catchUpDiffTracking で行数・ずれを確かめる）']
    ];
    checks.forEach(([label, ok, detail]) =>
      console.log((ok ? '○ ' : '✗ ') + name + ' / ' + label +
        (ok ? '' : ' （' + detail + '）')));
  }
}

/** Optional test handler: attach an installed onEdit trigger to it manually. */
function debugInstalledEdit(e) {
  if (!e || !e.range) return;

  e.range.setNote(
    'installed edit fired: ' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy/MM/dd HH:mm:ss'
    )
  );
}