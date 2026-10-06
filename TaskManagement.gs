/**
 * タスク管理
 * リクエスト シート（ServiceManagement.gs）でリクエストをもとに考えたサービス案ごとに、検討の進捗をタスクで管理するシート「タスク管理」。
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs と同じプロジェクトに置くファイル。
 *
 * 仕組み
 * - 今は手で行を追加する。1行が1件のタスクで、「サービス案」に対応付け、タスク・担当者・期限・状況・メモを入れる
 * - サービス案のプルダウンは、リクエスト シートのサービス案の列を範囲で参照する（リクエストにサービス案を付けると選べるようになる）
 * - 管理者が setupTaskSheet() をエディタから実行して、シート・見出し・プルダウンを用意する。
 *   前の版のシート（案件の連動プルダウンの「サブインダストリー」「得意先」「案件名」の列がある）は、
 *   確認してからその列を削除し、サービス案を左端に置く（ほかの列の値はそのまま残す）
 * - 見出しの用意・ロックなどの共通の処理は、ServiceManagement.gs からも使う
 */

const TASK_OPTIONS = {
  sheet: 'サービス案タスク管理',
  headerRow: 1,
  // aliases は前の版の見出し（見つかったら新しい見出しに書き換えて使い続ける）
  columns: [
    {key: 'service', label: 'サービス案', width: 220, type: 'service', aliases: ['サービス']},
    {key: 'task', label: 'タスク', width: 280},
    // 担当者とセットで入れる本部（前からあるシートに足すときは「タスク」のすぐ右＝担当者の左に差し込む）
    {key: 'department', label: '本部', width: 140, insertAfter: 'タスク'},
    {key: 'owner', label: '担当者', width: 120},
    {key: 'due', label: '期限', width: 100, type: 'date'},
    {key: 'status', label: '状況', width: 90, options: ['未着手', '対応中', '完了']},
    {key: 'note', label: 'メモ', width: 280}
  ],
  // 前の版の列（案件の連動プルダウン）。setupTaskSheet() で確認してから削除する
  removedColumns: ['サブインダストリー', '得意先', '案件名'],
  headerBackground: '#E9F4F1',
  lockWaitMs: 10000     // taskWithLock_ の既定の待ち時間（サービス管理は SVC_OPTIONS の待ち時間を渡す）
};

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * タスク管理シートを用意する（無ければ作る）。見出しが空なら書き、足りない列は右端に足し、全行にプルダウンを付ける。
 * 前の版の列（サブインダストリー・得意先・案件名）があれば、確認してから削除し、サービス案を左端に動かす。
 * 何度実行してもよい（入力済みのタスクは消さない）。
 */
function setupTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 前の版の列を消すかは、ロックを取る前に聞く（答えるまでほかの処理を待たせないため）
  const existing = ss.getSheetByName(TASK_OPTIONS.sheet);
  const oldColumns = existing ? taskFoundLabels_(existing, TASK_OPTIONS.removedColumns, TASK_OPTIONS.headerRow) : [];
  let removeOld = false;
  if (oldColumns.length) {
    const ui = SpreadsheetApp.getUi();
    removeOld = ui.alert(SVC_OPTIONS.menuTitle,
      'タスク管理は、案件ではなく、リクエストのサービス案に対応付けるようになりました。\n' +
      '「' + oldColumns.join('」「') + '」の列を削除し、「サービス案」の列を左端に置きます。\n' +
      'ほかの列（タスク・担当者など）の値はそのまま残ります。削除する列に入っている値は消えます。\n\n削除しますか？',
      ui.ButtonSet.YES_NO) === ui.Button.YES;
  }

  const result = taskWithLock_(() => {
    const {sheet, created, added} = taskEnsureSheet_(ss, TASK_OPTIONS.sheet, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow);
    let removed = [];
    if (removeOld) {
      removed = taskDeleteColumns_(sheet, TASK_OPTIONS.removedColumns, TASK_OPTIONS.headerRow);
      const col = taskColumnMap_(sheet, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow).service;
      if (col > 1) sheet.moveColumns(sheet.getRange(diffColumnLetter_(col) + ':' + diffColumnLetter_(col)), 1);
    }

    const rows = sheet.getMaxRows() - TASK_OPTIONS.headerRow;
    if (rows < 1) throw new Error(TASK_OPTIONS.sheet + ' にデータの行がありません。行を追加してから実行してください。');
    const cols = taskColumnMap_(sheet, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow);
    const warnings = taskApplyColumnRules_(sheet, TASK_OPTIONS.columns, cols, TASK_OPTIONS.headerRow + 1, rows);
    return {created, added, removed, warnings};
  });

  const lines = [result.created ? TASK_OPTIONS.sheet + ' シートを作りました。' : TASK_OPTIONS.sheet + ' シートのプルダウンを付け直しました。'];
  if (result.removed.length) lines.push('「' + result.removed.join('」「') + '」の列を削除し、「サービス案」を左端に置きました。');
  else if (oldColumns.length) lines.push('「' + oldColumns.join('」「') + '」の列は残しました（使いません。不要なら削除してください）。');
  if (result.added.length) lines.push('「' + result.added.join('」「') + '」の列を足しました。');
  ss.toast(lines.concat(result.warnings).join('\n'), SVC_OPTIONS.menuTitle, 10);
}

/* ---------------- シートの形（サービス管理と共通） ---------------- */

/** シートが無ければ作り、見出しを用意する。{sheet, created, added: 足した列の見出し} */
function taskEnsureSheet_(ss, name, columns, headerRow) {
  let sheet = ss.getSheetByName(name);
  const created = !sheet;
  if (!sheet) sheet = ss.insertSheet(name);
  return {sheet, created, added: taskEnsureHeader_(sheet, columns, headerRow)};
}

/**
 * 見出しの行が空なら見出しを書き、形を整える。
 * 見出しが入っているシートは書き換えない：前の版の見出し（aliases）は新しい見出しに書き換え、
 * 必須の列（required）が無ければ止め、足りない列だけを足す。insertAfter のある列は、その見出しの列のすぐ右に差し込み
 * （見出しが無ければ右端）、ほかは右端に足す。戻り値は足した列の見出し。
 */
function taskEnsureHeader_(sheet, columns, headerRow) {
  const where = sheet.getName() + ' の ' + headerRow + '行目';
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  let last = headers.length;
  while (last > 0 && !String(headers[last - 1]).trim()) last--;

  let start = 1;
  let targets = columns;
  if (last > 0) {
    columns.forEach(c => {
      if (taskHeaderCount_(headers, c.label) > 0) return;
      const alias = (c.aliases || []).find(a => taskHeaderCount_(headers, a) === 1);
      if (!alias) return;
      const col = taskFindColumn_(headers, alias, where, true);
      sheet.getRange(headerRow, col).setValue(c.label);
      headers[col - 1] = c.label;
    });
    columns.filter(c => c.required).forEach(c => taskFindColumn_(headers, c.label, where, true));
    targets = columns.filter(c => !c.required && taskHeaderCount_(headers, c.label) === 0);
  }

  // insertAfter の見出しがある列はそのすぐ右に、insertBefore の見出しがある列はそのすぐ左に差し込む
  // （右の列はずれる。差し込んだ列は隣の列の入力規則を引き継がないようにする）
  const inserted = [];
  if (last > 0) {
    targets = targets.filter(c => {
      const after = c.insertAfter ? taskFindColumn_(headers, c.insertAfter, where, false) : 0;
      const before = !after && c.insertBefore ? taskFindColumn_(headers, c.insertBefore, where, false) : 0;
      if (!after && !before) return true;
      const at = after ? after + 1 : before;
      if (after) sheet.insertColumnAfter(after); else sheet.insertColumnBefore(before);
      sheet.getRange(headerRow, at).setValue(c.label).setFontWeight('bold').setBackground(TASK_OPTIONS.headerBackground);
      if (sheet.getMaxRows() > headerRow) sheet.getRange(headerRow + 1, at, sheet.getMaxRows() - headerRow, 1).clearDataValidations();
      if (c.width) sheet.setColumnWidth(at, c.width);
      headers.splice(at - 1, 0, c.label);
      last++;
      inserted.push(c.label);
      return false;
    });
    start = last + 1;
  }
  if (!targets.length) return inserted;

  const needed = start + targets.length - 1;
  if (sheet.getMaxColumns() < needed) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  sheet.getRange(headerRow, start, 1, targets.length).setValues([targets.map(c => c.label)])
    .setFontWeight('bold').setBackground(TASK_OPTIONS.headerBackground);
  targets.forEach((c, i) => { if (c.width) sheet.setColumnWidth(start + i, c.width); });
  if (last === 0) sheet.setFrozenRows(headerRow);
  return last > 0 ? inserted.concat(targets.map(c => c.label)) : [];
}

/** 見出しから {key: 列番号} を作る（見つからない列は 0。required の列はちょうど1つ必要）。 */
function taskColumnMap_(sheet, columns, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const where = sheet.getName() + ' の ' + headerRow + '行目';
  const out = {};
  columns.forEach(c => { out[c.key] = taskFindColumn_(headers, c.label, where, !!c.required); });
  return out;
}

/** labels のうち、見出しの行にあるもの。 */
function taskFoundLabels_(sheet, labels, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  return labels.filter(label => taskHeaderCount_(headers, label) > 0);
}

/** 見出しが labels の列を削除する（右の列から消す）。戻り値は削除した列の見出し。 */
function taskDeleteColumns_(sheet, labels, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  const targets = labels.map(normalize);
  const removed = [];
  for (let col = headers.length; col >= 1; col--) {
    const index = targets.indexOf(normalize(headers[col - 1]));
    if (index < 0) continue;
    sheet.deleteColumn(col);
    removed.unshift(labels[index]);
  }
  return removed;
}

/**
 * 日付・一覧（options）・サービス案（type: 'service'）の列に、first 行目から rows 行の入力規則を付ける。
 * 戻り値は利用者に知らせる文（リクエスト シートが無いときなど）。
 */
function taskApplyColumnRules_(sheet, columns, cols, first, rows) {
  const warnings = [];
  columns.forEach(column => {
    const col = cols[column.key];
    if (!col) return;
    let rule = null;
    if (column.type === 'date') {
      rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('日付を入力してください。').build();
    } else if (column.type === 'service') {
      rule = typeof svcServiceRule_ === 'function' ? svcServiceRule_(sheet.getParent()) : null;
      if (!rule) {
        warnings.push('リクエスト シートが無いため、「' + column.label + '」の列にプルダウンを付けていません。setupRequestSheet() を実行してください。');
        return;
      }
    } else if (column.options) {
      rule = SpreadsheetApp.newDataValidation().requireValueInList(column.options, true).setAllowInvalid(false)
        .setHelpText(column.label + 'は一覧から選んでください。').build();
    } else {
      return;
    }
    sheet.getRange(first, col, rows, 1).setDataValidation(rule);
  });
  return warnings;
}

/** 見出しの列番号（改行・空白の違いは無視）。required なら、ちょうど1つでないときに止める。 */
function taskFindColumn_(headers, label, where, required) {
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  const target = normalize(label);
  const matches = [];
  headers.forEach((h, i) => { if (normalize(h) === target) matches.push(i + 1); });
  if (matches.length === 1) return matches[0];
  if (!required) return 0;
  throw new Error(where + 'に「' + label + '」の見出しがちょうど1つ必要です（見つかった数: ' + matches.length + '）。');
}

function taskHeaderCount_(headers, label) {
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  return headers.filter(h => normalize(h) === normalize(label)).length;
}

/**
 * サービス管理・タスク管理の書き込みのロックの中で fn を実行する（waitMs を省くと TASK_OPTIONS.lockWaitMs）。
 * サービス管理・タスク管理は、サービスリクエスト・タスク管理のシート（差分追跡の対象外）だけに書き、取り込み元は読むだけなので、
 * 差分追跡・クレデンシャル履歴などのドキュメントロックとは別の、スクリプトロックを使う
 * （ドキュメントロックだと、取り込み元の編集のたびに動く差分追跡の後ろに並ばされ、選択パネルの「追加」が待ちきれずに止まっていた）。
 * サービス管理・タスク管理の処理どうしは、このロックで1つずつ動く。
 */
function taskWithLock_(fn, waitMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(waitMs || TASK_OPTIONS.lockWaitMs)) {
    throw new Error('他の処理が実行中です。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}
