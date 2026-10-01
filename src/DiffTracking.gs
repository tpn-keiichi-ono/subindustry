/**
 * Character-level differences for Google Sheets (container-bound Apps Script).
 * Revision: event IDs + editor email + row snapshots + source column names
 *           + lock-busy deferral + catch-up scan + automatic row resync.
 *
 * CHANGES IN THIS REVISION (v5):
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
 * Run it after changing tracked ranges or after column changes/sorting.
 * Do not delete/edit the hidden __CHAR_DIFF_* sheets or sort the source while
 * tracking. Source cell CONTENT/FORMATTING is never rewritten; only the
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

const DIFF_RULES = {
  '新FMT': {
    headerRow: 2,
    ranges: ['A3:AF']
    // Optional: capture an explicit row span in the row snapshot.
    // , snapshotColumns: 'A:Z'
  }
  // Add other existing tabs, for example:
  , '新FMT2': { headerRow: 2, ranges: ['A3:AF']}
};

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
  maxDiffCharacters: 45000,
  maxLcsCells: 1000000,
  maxLcsCellsPerEdit: 4000000,
  dateFormat: 'yyyy/MM/dd HH:mm'
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
  maxSummaryCharacters: 40000
};

const DIFF_HEADERS = [
  '記録日時', 'イベントID', '編集者メールアドレス', 'シート', 'セル', '列名',
  '種類', '変更前', '変更後',
  '差分', '確認済み'
];

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
    plans.forEach(plan => diffClearDirty_(plan.sheet.getSheetId()));

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

/**
 * Installed edit trigger. Do not run manually or add an onEdit wrapper.
 *
 * - differences are recorded per actually changed cell
 * - row snapshot = one record per event ID x changed row
 * - if the lock is busy, the edit is deferred to catchUpDiffTracking()
 * - if only the number of rows changed (row insert/delete), the sheet is
 *   resynchronized first (the edit is recorded as part of that)
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
    // 行の追加・削除の直後で、変更トリガーより先にこの編集が届いた場合
    if (diffRowCountOnlyChanged_(sheet, rule)) {
      diffResyncRows_(e.source, sheet, rule, {
        structure: DIFF_ROW_OPTIONS.unknownEditor,
        changes: diffEditorEmail_(e)
      });
      return;
    }

    const ctx = diffContext_(e.source, sheet, rule);

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
 * Does nothing unless a sheet was marked dirty by a deferred edit.
 * For each dirty sheet, compares the entire tracked range with the baseline
 * and records every difference found.
 */
function catchUpDiffTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  const pending = Object.entries(DIFF_RULES)
    .map(([name, rule]) => ({name, rule, sheet: ss.getSheetByName(name)}))
    .filter(p => p.sheet &&
      props.getProperty(DIFF_OPTIONS.dirtyPrefix + p.sheet.getSheetId()));

  if (!pending.length) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DIFF_OPTIONS.catchUpLockWaitMs)) {
    console.warn('Lock busy; catch-up will retry on the next run.');
    return;
  }

  try {
    const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');

    for (const p of pending) {
      const sheetId = p.sheet.getSheetId();
      const key = DIFF_OPTIONS.dirtyPrefix + sheetId;
      const token = props.getProperty(key);
      if (!token) continue;

      const saved = state[sheetId];
      if (!saved || saved.paused) {
        // setupDiffTracking() will take a fresh baseline and clear the flag.
        console.warn('Tracking paused; catch-up skipped: ' + p.name);
        continue;
      }

      try {
        // 行の追加・削除を変更トリガーが処理しきれていなければ、ここで処理する
        if (diffRowCountOnlyChanged_(p.sheet, p.rule)) {
          diffResyncRows_(ss, p.sheet, p.rule, {
            structure: DIFF_ROW_OPTIONS.unknownEditor,
            changes: DIFF_OPTIONS.catchUpEditor
          });
          continue;
        }

        const ctx = diffContext_(ss, p.sheet, p.rule);
        const blocks = diffSplitBlocks_(
          diffBlocks_(p.sheet, p.rule, null, ctx.stampColumn, false),
          DIFF_OPTIONS.catchUpChunkCells
        );
        const now = new Date();
        const count = diffProcessBlocks_(ctx, blocks, {
          now,
          eventId: diffCreateEventId_(now, ctx.timezone, 'REC'),
          editorEmail: DIFF_OPTIONS.catchUpEditor,
          kindSuffix: DIFF_OPTIONS.catchUpKindSuffix
        });
        console.log('Catch-up done: ' + p.name + ' / recovered cells=' + count);

        // Clear only if no newer deferral happened during this scan.
        if (props.getProperty(key) === token) props.deleteProperty(key);

      } catch (error) {
        diffPause_(ss, sheetId);
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
 * Installed change trigger.
 *
 * - 行の追加・削除（行数だけが変わった場合）：追跡は止めずに diffResyncRows_() で
 *   削除・追加された行を記録し、そのシートの比較基準を取り直す。
 * - 列の追加・削除、シートの削除：これまでどおり、そのシートの追跡を停止する。
 *   DIFF_RULES の列指定がずれるため、確認のうえ setupDiffTracking() を手動で実行する。
 * - ロックが取れないときは「取りこぼし」として印を付け、catchUpDiffTracking() で処理する。
 *
 * Note: row sorting or drag-moving rows does not change the grid size and
 * cannot be detected here.
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

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DIFF_ROW_OPTIONS.structureLockWaitMs)) {
    // 取りこぼし回収に任せる（行数の変化は catchUpDiffTracking() でも処理される）
    Object.keys(DIFF_RULES).forEach(name => {
      const sheet = e.source.getSheetByName(name);
      if (sheet) diffMarkDirty_(sheet.getSheetId());
    });
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

    const editor = diffEditorEmail_(e);
    const structureEditor = editor === '取得不可' ? DIFF_ROW_OPTIONS.unknownEditor : editor;

    for (const sheet of rowChanged) {
      const rule = DIFF_RULES[sheet.getName()];
      if (!rule || !diffRowCountOnlyChanged_(sheet, rule)) {
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


/* ---------------- Row insert/delete: record and rebaseline (v5) ---------------- */

/**
 * 行数だけが比較基準と違う（列・設定・シート名は同じ）なら true。
 * このときは停止せず diffResyncRows_() で処理できる。
 */
function diffRowCountOnlyChanged_(sheet, rule) {
  const state = JSON.parse(
    PropertiesService.getDocumentProperties().getProperty(DIFF_OPTIONS.stateKey) || '{}');
  const saved = state[sheet.getSheetId()];
  if (!saved || saved.paused || saved.gridRows === sheet.getMaxRows()) return false;

  let stampColumn;
  try { stampColumn = diffStampColumn_(sheet, rule); } catch (_) { return false; }

  return saved.name === sheet.getName() &&
    saved.signature === JSON.stringify(rule) &&
    saved.stampColumn === stampColumn &&
    saved.gridColumns === sheet.getMaxColumns();
}

/**
 * ロック取得中に呼ぶこと。比較基準（行が増減する前）と現在の行を突き合わせて、
 * 削除された行・追加された行・その間に変わったセルを記録し、比較基準を取り直す。
 * who = {structure: 行の追加・削除をした人, changes: セルを変えた人}
 * 戻り値 {deleted, inserted, changedCells}
 */
function diffResyncRows_(ss, sheet, rule, who) {
  const props = PropertiesService.getDocumentProperties();
  const state = JSON.parse(props.getProperty(DIFF_OPTIONS.stateKey) || '{}');
  const sheetId = sheet.getSheetId();
  const saved = state[sheetId];
  if (!saved) throw new Error('Tracking is not initialized. Run setupDiffTracking().');

  const baseline = ss.getSheets().find(s => s.getSheetId() === saved.snapshotId);
  if (!baseline) throw new Error('Comparison snapshot is missing.');
  diffLogSheet_(ss);
  diffRowSnapshotSheet_(ss);

  const timezone = ss.getSpreadsheetTimeZone();
  const stampColumn = saved.stampColumn;
  const blocks = diffBlocks_(sheet, rule, null, stampColumn, false);
  const result = {deleted: 0, inserted: 0, changedCells: 0};

  // 読み込みより前に印を消す（このあと届いた編集は印が付き、取りこぼし回収される）
  diffClearDirty_(sheetId);

  if (blocks.length) {
    for (const block of blocks) {
      if (sheet.getRange(...block).getMergedRanges().length) {
        throw new Error('Merged cells are not supported: ' + sheet.getRange(...block).getA1Notation());
      }
    }

    const rowStart = Math.min(...blocks.map(b => b[0]));
    const newEnd = Math.max(...blocks.map(b => b[0] + b[2] - 1));
    const oldEnd = newEnd + (saved.gridRows - sheet.getMaxRows());
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

    const gaps = diffAlignRows_(
      oldRows.map(r => r.join('\u0001')),
      newRows.map(r => r.join('\u0001'))
    );

    const now = new Date();
    const eventId = diffCreateEventId_(now, timezone, DIFF_ROW_OPTIONS.eventPrefix);
    const budget = {remaining: DIFF_OPTIONS.maxLcsCellsPerEdit};
    const records = [];
    const changedRows = new Set();
    const changedCellsByRow = new Map();
    const isEmpty = row => row.every(token => !token);
    const summary = row => {
      const text = row
        .map((token, c) => token ? columnNames[c] + ': ' + diffDecode_(token, timezone).text : '')
        .filter(Boolean).join('\n');
      return text.length > DIFF_ROW_OPTIONS.maxSummaryCharacters
        ? text.slice(0, DIFF_ROW_OPTIONS.maxSummaryCharacters) + '…（省略）'
        : text;
    };
    const markRow = (rowNumber, cell) => {
      changedRows.add(rowNumber);
      if (!changedCellsByRow.has(rowNumber)) changedCellsByRow.set(rowNumber, new Set());
      changedCellsByRow.get(rowNumber).add(cell);
    };

    for (const gap of gaps) {
      if (gap.old.length === gap.new.length) {
        // 同じ数だけ入れ替わった行 = 行はそのままで中身が変わった（取りこぼしていた編集）
        gap.old.forEach((oldIndex, k) => {
          const newIndex = gap.new[k];
          const before = oldRows[oldIndex];
          const after = newRows[newIndex];
          const rowNumber = rowStart + newIndex;
          after.forEach((token, c) => {
            if (token === before[c]) return;
            const b = diffDecode_(before[c], timezone);
            const a = diffDecode_(token, timezone);
            const diff = diffCharacters_(b.text, a.text, budget);
            const cell = diffA1_(rowNumber, columnNumbers[c]);
            records.push({
              eventId,
              editorEmail: who.changes,
              sheet: sheet.getName(),
              cell,
              columnName: columnNames[c],
              kind: b.type + ' -> ' + a.type + (diff.coarse ? ' / ブロック差分' : '') +
                DIFF_ROW_OPTIONS.kindSuffix,
              before: b.text,
              after: a.text,
              parts: diff.parts
            });
            markRow(rowNumber, cell);
            result.changedCells++;
          });
        });
        continue;
      }

      gap.old.forEach(oldIndex => {
        const row = oldRows[oldIndex];
        if (isEmpty(row)) return;
        const text = summary(row);
        records.push({
          eventId,
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

      gap.new.forEach(newIndex => {
        const row = newRows[newIndex];
        if (isEmpty(row)) return;
        const text = summary(row);
        const rowNumber = rowStart + newIndex;
        records.push({
          eventId,
          editorEmail: who.changes,
          sheet: sheet.getName(),
          cell: rowNumber + '行',
          columnName: '（行全体）',
          kind: DIFF_ROW_OPTIONS.kindInsert,
          before: '',
          after: text,
          parts: [{kind: 'add', text}]
        });
        markRow(rowNumber, rowNumber + '行');
        result.inserted++;
      });
    }

    if (records.length) {
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
          diffAppendRowSnapshots_(
            ss, sheet, rule, sortedRows, changedCellsByRow, eventId, who.changes, now, schema);
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

function diffRich_(parts) {
  const normal = SpreadsheetApp.newTextStyle().setForegroundColor('#000000')
    .setBold(false).setStrikethrough(false).build();
  const styles = {
    add: SpreadsheetApp.newTextStyle().setForegroundColor('#ff0000').setBold(true)
      .setStrikethrough(false).build(),
    del: SpreadsheetApp.newTextStyle().setForegroundColor('#ff0000').setBold(false)
      .setStrikethrough(true).build()
  };
  const full = parts.map(p => p.text).join('');
  if (full.length > DIFF_OPTIONS.maxDiffCharacters) {
    return diffPlain_('差分が長いため省略。変更前・変更後列を確認。');
  }
  const builder = SpreadsheetApp.newRichTextValue().setText(full || '空欄').setTextStyle(normal);
  let offset = 0;
  for (const part of parts) {
    if (styles[part.kind]) builder.setTextStyle(offset, offset + part.text.length, styles[part.kind]);
    offset += part.text.length;
  }
  return builder.build();
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
    return log;
  }

  diffEnsureGrid_(log, 1, DIFF_HEADERS.length);

  const actual = log.getRange(1, 1, 1, DIFF_HEADERS.length)
    .getDisplayValues()[0];

  if (diffHeadersEqual_(actual, DIFF_HEADERS)) {
    return log;
  }

  // Previous 10-column version: insert editor email after Event ID.
  if (diffHeadersEqual_(actual.slice(0, DIFF_LEGACY_HEADERS_10.length),
      DIFF_LEGACY_HEADERS_10)) {
    log.insertColumnAfter(2);
    log.getRange(1, 1, 1, DIFF_HEADERS.length).setValues([DIFF_HEADERS]);
    diffFormatDiffHeader_(log);
    return log;
  }

  // Previous 9-column version: insert Event ID and editor email.
  if (diffHeadersEqual_(actual.slice(0, DIFF_LEGACY_HEADERS_9.length),
      DIFF_LEGACY_HEADERS_9)) {
    log.insertColumnAfter(1);
    log.insertColumnAfter(2);
    log.getRange(1, 1, 1, DIFF_HEADERS.length).setValues([DIFF_HEADERS]);
    diffFormatDiffHeader_(log);
    return log;
  }

  // Oldest 8-column version: insert column name, Event ID, and editor email.
  if (diffHeadersEqual_(actual.slice(0, DIFF_LEGACY_HEADERS_8.length),
      DIFF_LEGACY_HEADERS_8)) {
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
  log.setColumnWidth(10, 420);
  log.setColumnWidth(11, 90);
}

function diffAppend_(ss, records, now) {
  const log = diffLogSheet_(ss);
  const row = log.getLastRow() + 1;

  diffEnsureGrid_(log, row + records.length - 1, DIFF_HEADERS.length);

  log.getRange(row, 1, records.length, 1)
    .setValues(records.map(() => [now]))
    .setNumberFormat(DIFF_OPTIONS.dateFormat);

  // Rich text prevents strings beginning with '=' from becoming executable formulas.
  log.getRange(row, 2, records.length, 9).setRichTextValues(records.map(rec => [
    diffPlain_(rec.eventId),
    diffPlain_(rec.editorEmail),
    diffPlain_(rec.sheet),
    diffPlain_(rec.cell),
    diffPlain_(rec.columnName),
    diffPlain_(rec.kind),
    diffPlain_(rec.before),
    diffPlain_(rec.after),
    diffRich_(rec.parts)
  ]));

  log.getRange(row, 11, records.length, 1).insertCheckboxes();

  log.getRange(row, 1, records.length, DIFF_HEADERS.length)
    .setVerticalAlignment('top')
    .setWrap(true);
}


/* ---------------- Row snapshot log ---------------- */

function diffRowSnapshotSheet_(ss) {
  let sheet = ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet);

  if (!sheet) {
    sheet = ss.insertSheet(DIFF_OPTIONS.rowSnapshotSheet);
    diffEnsureGrid_(sheet, 1, ROW_SNAPSHOT_FIXED_HEADERS.length);
    sheet.getRange(1, 1, 1, ROW_SNAPSHOT_FIXED_HEADERS.length)
      .setValues([ROW_SNAPSHOT_FIXED_HEADERS])
      .setFontWeight('bold')
      .setBackground('#e8eef5');

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 210);
    sheet.setColumnWidth(3, 240);
    sheet.setColumnWidth(4, 130);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 200);
    return sheet;
  }

  diffEnsureGrid_(sheet, 1, ROW_SNAPSHOT_FIXED_HEADERS.length);

  const actual = sheet
    .getRange(1, 1, 1, ROW_SNAPSHOT_FIXED_HEADERS.length)
    .getDisplayValues()[0];

  if (diffHeadersEqual_(actual, ROW_SNAPSHOT_FIXED_HEADERS)) {
    return sheet;
  }

  // Previous 5-column fixed layout: insert editor email after Event ID.
  if (diffHeadersEqual_(actual.slice(0, ROW_SNAPSHOT_LEGACY_HEADERS_5.length),
      ROW_SNAPSHOT_LEGACY_HEADERS_5)) {
    sheet.insertColumnAfter(2);
    sheet.getRange(1, 1, 1, ROW_SNAPSHOT_FIXED_HEADERS.length)
      .setValues([ROW_SNAPSHOT_FIXED_HEADERS])
      .setFontWeight('bold')
      .setBackground('#e8eef5');
    sheet.setColumnWidth(3, 240);
    return sheet;
  }

  throw new Error(
    'Row snapshot sheet has unexpected fixed headers: ' +
    DIFF_OPTIONS.rowSnapshotSheet
  );
}

function diffAppendRowSnapshots_(
    ss, sourceSheet, rule, rows, changedCellsByRow, eventId, editorEmail, now, schema) {
  if (!rows.length) return;

  const snapshotSheet = diffRowSnapshotSheet_(ss);
  schema = schema || diffSnapshotSchema_(sourceSheet, rule);

  const snapshotCells = rows.length * schema.items.length;
  if (snapshotCells > DIFF_OPTIONS.maxSnapshotCellsPerEdit) {
    throw new Error(
      'Row snapshot exceeds ' +
      DIFF_OPTIONS.maxSnapshotCellsPerEdit +
      ' cells: ' + snapshotCells
    );
  }

  const keyToSheetColumn =
    diffEnsureRowSnapshotColumns_(snapshotSheet, schema.items);

  const startRow = snapshotSheet.getLastRow() + 1;
  const finalColumn = snapshotSheet.getLastColumn();

  diffEnsureGrid_(
    snapshotSheet,
    startRow + rows.length - 1,
    finalColumn
  );

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

  const dynamicWidth = Math.max(
    0,
    finalColumn - ROW_SNAPSHOT_FIXED_HEADERS.length
  );

  if (dynamicWidth) {
    const matrix = rows.map(rowNumber => {
      const output = Array.from(
        {length: dynamicWidth},
        () => diffSnapshotText_('')
      );

      const values = rowValues.get(rowNumber);

      schema.items.forEach((item, index) => {
        const targetColumn = keyToSheetColumn[item.key];
        const dynamicIndex =
          targetColumn - ROW_SNAPSHOT_FIXED_HEADERS.length - 1;

        if (dynamicIndex >= 0 && dynamicIndex < dynamicWidth) {
          output[dynamicIndex] = diffSnapshotText_(values[index]);
        }
      });

      return output;
    });

    snapshotSheet
      .getRange(
        startRow,
        ROW_SNAPSHOT_FIXED_HEADERS.length + 1,
        rows.length,
        dynamicWidth
      )
      .setRichTextValues(matrix);
  }

  snapshotSheet
    .getRange(startRow, 1, rows.length, finalColumn)
    .setVerticalAlignment('top')
    .setWrap(true);
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

function diffEnsureRowSnapshotColumns_(snapshotSheet, schemaItems) {
  const fixed = ROW_SNAPSHOT_FIXED_HEADERS.length;
  const currentLast = Math.max(snapshotSheet.getLastColumn(), fixed);

  const existingHeaders = currentLast > fixed
    ? snapshotSheet
        .getRange(1, fixed + 1, 1, currentLast - fixed)
        .getDisplayValues()[0]
    : [];

  const map = {};

  existingHeaders.forEach((header, index) => {
    if (header) map[header] = fixed + index + 1;
  });

  const missing = schemaItems
    .map(item => item.key)
    .filter(key => map[key] == null);

  if (missing.length) {
    const startColumn = snapshotSheet.getLastColumn() + 1;

    diffEnsureGrid_(
      snapshotSheet,
      1,
      startColumn + missing.length - 1
    );

    snapshotSheet
      .getRange(1, startColumn, 1, missing.length)
      .setRichTextValues([missing.map(value => diffSnapshotText_(value))])
      .setFontWeight('bold')
      .setBackground('#e8eef5');

    snapshotSheet.setColumnWidths(startColumn, missing.length, 180);

    missing.forEach((header, index) => {
      map[header] = startColumn + index;
    });
  }

  return map;
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

    const checks = [
      ['停止していない', saved.paused === false, saved.paused],
      ['シート名', saved.name === sheet.getName(), saved.name + ' / ' + sheet.getName()],
      ['DIFF_RULES', saved.signature === JSON.stringify(rule), saved.signature + ' / ' + JSON.stringify(rule)],
      ['タイムスタンプ列', saved.stampColumn === stamp, saved.stampColumn + ' / ' + stamp],
      ['行数（違う場合は次の編集・回収で自動反映）', saved.gridRows === sheet.getMaxRows(), saved.gridRows + ' / ' + sheet.getMaxRows()],
      ['列数', saved.gridColumns === sheet.getMaxColumns(), saved.gridColumns + ' / ' + sheet.getMaxColumns()],
      ['比較用シート', ss.getSheets().some(s => s.getSheetId() === saved.snapshotId), saved.snapshotId],
      ['未回収の取りこぼしなし', !dirty, '回収待ち（次回の catchUpDiffTracking で処理）']
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
