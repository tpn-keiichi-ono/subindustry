/**
 * サービス管理
 * 新FMT の「サービスのリクエスト」（Z列）を一覧にして、いくつかをまとめてサービス化を検討する・棄却するためのシート。
 * TaskManagement.gs と同じプロジェクトに置くファイル（見出しの用意・ロック・onEdit は TaskManagement.gs と共通）。
 *
 * 仕組み
 * - 「サービス」シート：1行が1つのサービス（サービス名・状況・概要・担当者・メモ）。サービス名は重複させない
 * - 「リクエスト」シート：1行が1件のリクエスト。importServiceRequests() をエディタから実行すると、
 *   新FMT のリクエストのうち、まだ一覧に無いもの（得意先・案件名・リクエストの組で見分ける）だけを追記する。
 *   「サービス」でまとめる先のサービスを選び、「判断」（未判断・サービス化検討・棄却）を付ける
 * - タスク管理シートの「サービス」の列で、タスクをサービス（の開発）に紐づける
 * - サービスを選ぶ列のプルダウンは、サービス シートのサービス名の列を範囲で参照する（サービスを足すと自動で選べる）
 * - 単純トリガーの onEdit（TaskManagement.gs）から：
 *   - リクエストでサービスを選んだとき、判断が未判断なら「サービス化検討」にする
 *   - サービス名を変えたとき、リクエスト・タスク管理の同じ名前も変える（同じ名前が2つあるときは変えずに知らせる）
 * - 取り込み元から消えた・書き換えられたリクエストは、行を消さずに「リクエスト」のセルに注を付ける
 */

const SVC_OPTIONS = {
  serviceSheet: 'サービス',
  requestSheet: 'リクエスト',
  headerRow: 1,
  // 取り込み元（TASK_OPTIONS.sourceSheets。新FMT）でリクエストが書かれた列の見出し（Z列）。1つのセルに1件
  requestHeader: 'サービスのリクエスト',
  serviceColumns: [
    {key: 'name', label: 'サービス名', width: 220, required: true},
    {key: 'status', label: '状況', width: 100, options: ['検討中', '開発中', '提供中', '見送り']},
    {key: 'summary', label: '概要', width: 320},
    {key: 'owner', label: '担当者', width: 120},
    {key: 'note', label: 'メモ', width: 280}
  ],
  // 左の4列は取り込みで入れる（手で変えない）。サービス・判断・メモを手で入れる
  requestColumns: [
    {key: 'subIndustry', label: 'サブインダストリー', width: 160, required: true},
    {key: 'customer', label: '得意先', width: 200, required: true},
    {key: 'project', label: '案件名', width: 240, required: true},
    {key: 'request', label: 'リクエスト', width: 360, required: true},
    {key: 'service', label: 'サービス', width: 200, type: 'service'},
    {key: 'decision', label: '判断', width: 130, options: ['未判断', 'サービス化検討', '棄却']},
    {key: 'note', label: 'メモ', width: 280},
    {key: 'importedAt', label: '取り込み日', width: 100}
  ],
  defaultDecision: '未判断',             // 取り込んだときの判断
  decisionWithService: 'サービス化検討',  // サービスを選んだとき、判断が未判断（または空）ならこれにする
  missingNote: '取り込み元に見つかりません',   // 取り込み元から消えた・書き換えられたリクエストに付ける注の先頭
  lockWaitMs: 30000                       // エディタから実行する処理のロックの待ち時間
};

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * サービス・リクエストのシートを用意する（無ければ作る）。見出しとプルダウンを付ける。
 * タスク管理シートがあれば「サービス」の列（無ければ右端に足す）にもプルダウンを付ける。何度実行してもよい。
 */
function setupServiceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const sheets = svcEnsureSheets_(ss);
    return Object.assign(sheets, svcApplyRules_(ss, sheets));
  }, SVC_OPTIONS.lockWaitMs);
  ss.toast(svcSetupMessage_(result), 'サービス管理', 10);
}

/**
 * 新FMT のリクエスト（Z列）のうち、まだリクエスト シートに無いものを追記する。
 * シートが無ければ作る。追記したあと、全行のプルダウンを付け直す。
 */
function importServiceRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const sheets = svcEnsureSheets_(ss);
    const imported = svcImport_(ss, sheets.request);
    return Object.assign(sheets, imported, svcApplyRules_(ss, sheets));
  }, SVC_OPTIONS.lockWaitMs);

  ss.toast(
    'リクエストを ' + result.added + '件追加しました（取り込み元のリクエスト ' + result.total + '件）。' +
    (result.missing ? '\n取り込み元に見つからないリクエストが ' + result.missing + '件あります（「リクエスト」のセルに注を付けました）。' : '') +
    (result.warnings.length ? '\n' + result.warnings.join('\n') : ''),
    'サービス管理', 10);
}

/* ---------------- シートの用意 ---------------- */

function svcEnsureSheets_(ss) {
  // サービスを選ぶ列はサービス シートを参照するので、先にサービス シートを用意する
  const service = taskEnsureSheet_(ss, SVC_OPTIONS.serviceSheet, SVC_OPTIONS.serviceColumns, SVC_OPTIONS.headerRow);
  const request = taskEnsureSheet_(ss, SVC_OPTIONS.requestSheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
  return {
    service: service.sheet,
    request: request.sheet,
    created: [service, request].filter(s => s.created).map(s => s.sheet.getName()),
    added: [service, request].filter(s => s.added.length).map(s => s.sheet.getName() + '：' + s.added.join('・'))
  };
}

/** 全行のプルダウン（状況・サービス・判断）を付ける。タスク管理シートがあれば、その「サービス」の列にも付ける。 */
function svcApplyRules_(ss, sheets) {
  const warnings = [];
  [[sheets.service, SVC_OPTIONS.serviceColumns], [sheets.request, SVC_OPTIONS.requestColumns]].forEach(([sheet, columns]) => {
    const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
    if (rows < 1) return;
    const cols = taskColumnMap_(sheet, columns, SVC_OPTIONS.headerRow);
    warnings.push.apply(warnings, taskApplyColumnRules_(sheet, columns, cols, SVC_OPTIONS.headerRow + 1, rows));
    if (cols.importedAt) sheet.getRange(SVC_OPTIONS.headerRow + 1, cols.importedAt, rows, 1).setNumberFormat('yyyy/MM/dd');
  });

  let taskAdded = [];
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  if (task) {
    taskAdded = taskEnsureHeader_(task, taskAllColumns_(), TASK_OPTIONS.headerRow);
    const rows = task.getMaxRows() - TASK_OPTIONS.headerRow;
    const serviceColumns = TASK_OPTIONS.columns.filter(c => c.type === 'service');
    if (rows > 0) taskApplyColumnRules_(task, serviceColumns, taskColumns_(task).others, TASK_OPTIONS.headerRow + 1, rows);
  }
  return {warnings, taskAdded};
}

function svcSetupMessage_(result) {
  const lines = [result.created.length
    ? '「' + result.created.join('」「') + '」シートを作りました。'
    : 'サービス・リクエストのシートのプルダウンを付け直しました。'];
  if (result.added.length) lines.push('右端に列を足しました（' + result.added.join('、') + '）。');
  if (result.taskAdded.length) lines.push(TASK_OPTIONS.sheet + ' の右端に「' + result.taskAdded.join('」「') + '」の列を足しました。');
  return lines.concat(result.warnings).join('\n');
}

/**
 * サービスを選ぶ列の入力規則（サービス シートのサービス名の列を範囲で参照する）。サービス シートが無ければ null。
 * タスク管理シートの「サービス」の列（TaskManagement.gs）でも使う。
 */
function svcServiceRule_(ss) {
  const sheet = ss.getSheetByName(SVC_OPTIONS.serviceSheet);
  if (!sheet) return null;
  const nameColumn = SVC_OPTIONS.serviceColumns.filter(c => c.key === 'name');
  const letter = diffColumnLetter_(taskColumnMap_(sheet, nameColumn, SVC_OPTIONS.headerRow).name);
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange(letter + (SVC_OPTIONS.headerRow + 1) + ':' + letter), true)
    .setAllowInvalid(false)
    .setHelpText('「' + SVC_OPTIONS.serviceSheet + '」シートのサービス名から選んでください（空欄でもかまいません）。')
    .build();
}

/* ---------------- リクエストの取り込み ---------------- */

/** リクエストを見分けるキー（得意先・案件名・リクエスト。全角・半角と空白の違いは無視）。 */
function svcRequestKey_(customer, project, request) {
  return [customer, project, request].map(credNormalize_).join('\u0001');
}

/**
 * 取り込み元のリクエストのうち、まだシートに無いものを末尾に追記する。
 * シートにあって取り込み元に無いリクエストには注を付ける（取り込み元に戻ったら注を外す。利用者が書いた注は変えない）。
 * {added, missing, total}
 */
function svcImport_(ss, sheet) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);

  const sourceHeaders = TASK_OPTIONS.cascade.map(c => c.source).concat([SVC_OPTIONS.requestHeader]);
  const source = [];
  const sourceKeys = new Set();
  taskReadSource_(ss, sourceHeaders).forEach(values => {
    if (!values[3]) return;
    const key = svcRequestKey_(values[1], values[2], values[3]);
    if (sourceKeys.has(key)) return;
    sourceKeys.add(key);
    source.push({values, key});
  });

  const lastRow = Math.max(sheet.getLastRow(), headerRow);
  const count = lastRow - headerRow;
  const existing = new Set();
  let missing = 0;
  if (count > 0) {
    const read = col => sheet.getRange(headerRow + 1, col, count, 1).getDisplayValues().map(r => r[0]);
    const customers = read(cols.customer);
    const projects = read(cols.project);
    const requests = read(cols.request);
    const noteRange = sheet.getRange(headerRow + 1, cols.request, count, 1);
    const notes = noteRange.getNotes().map(r => String(r[0] || ''));
    const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd');
    let changed = false;

    const next = requests.map((request, i) => {
      if (!String(request).trim()) return notes[i];
      const key = svcRequestKey_(customers[i], projects[i], request);
      existing.add(key);
      const ours = notes[i].indexOf(SVC_OPTIONS.missingNote) === 0;
      if (sourceKeys.has(key)) {
        if (!ours) return notes[i];
        changed = true;
        return '';
      }
      missing++;
      if (notes[i]) return notes[i];   // 前回から見つからないまま、または利用者が書いた注
      changed = true;
      return SVC_OPTIONS.missingNote + '（' + today + ' の取り込みで気づきました）。' +
        TASK_OPTIONS.sourceSheets.join('・') + ' で書き換えか削除された可能性があります。';
    });
    if (changed) noteRange.setNotes(next.map(n => [n]));
  }

  const fresh = source.filter(r => !existing.has(r.key));
  if (fresh.length) {
    const start = lastRow + 1;
    const end = start + fresh.length - 1;
    if (end > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), end - sheet.getMaxRows());
    const width = Math.max.apply(null, Object.keys(cols).map(k => cols[k]));
    const now = new Date();
    const rows = fresh.map(r => {
      const row = new Array(width).fill('');
      // 取り込み元の文字列は数式にならないよう、先頭に ' を付けて書く
      ['subIndustry', 'customer', 'project', 'request'].forEach((key, i) => { row[cols[key] - 1] = credText_(r.values[i]); });
      if (cols.decision) row[cols.decision - 1] = SVC_OPTIONS.defaultDecision;
      if (cols.importedAt) row[cols.importedAt - 1] = now;
      return row;
    });
    sheet.getRange(start, 1, rows.length, width).setValues(rows);
  }
  return {added: fresh.length, missing, total: source.length};
}

/* ---------------- 編集したとき（TaskManagement.gs の onEdit から） ---------------- */

/** リクエストでサービスを選んだ行は、判断が未判断（または空）なら「サービス化検討」にする。 */
function svcHandleRequestEdit_(sheet, range) {
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
  if (!cols.service || !cols.decision) return;
  if (cols.service < range.getColumn() || cols.service > range.getLastColumn()) return;

  const first = Math.max(range.getRow(), SVC_OPTIONS.headerRow + 1);
  const last = Math.min(range.getLastRow(), first + TASK_OPTIONS.maxEditRows - 1);
  const rows = last - first + 1;
  taskWithLock_(() => {
    const services = sheet.getRange(first, cols.service, rows, 1).getDisplayValues();
    const decisions = sheet.getRange(first, cols.decision, rows, 1).getDisplayValues();
    services.forEach((service, i) => {
      const decision = String(decisions[i][0]).trim();
      if (String(service[0]).trim() && (!decision || decision === SVC_OPTIONS.defaultDecision)) {
        sheet.getRange(first + i, cols.decision).setValue(SVC_OPTIONS.decisionWithService);
      }
    });
  });
}

/**
 * サービス名を変えたとき：同じ名前が2つ以上あれば知らせる。
 * 1つのセルの名前を変えたのなら、リクエスト・タスク管理の「サービス」の列の古い名前も新しい名前にする。
 */
function svcHandleServiceEdit_(sheet, range, e) {
  const nameColumn = SVC_OPTIONS.serviceColumns.filter(c => c.key === 'name');
  const col = taskColumnMap_(sheet, nameColumn, SVC_OPTIONS.headerRow).name;
  if (col < range.getColumn() || col > range.getLastColumn()) return;

  const ss = sheet.getParent();
  const count = sheet.getLastRow() - SVC_OPTIONS.headerRow;
  const names = count > 0
    ? sheet.getRange(SVC_OPTIONS.headerRow + 1, col, count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : [];
  const seen = new Set();
  const duplicates = [];
  names.forEach(name => {
    if (!name) return;
    const key = credNormalize_(name);
    if (seen.has(key) && duplicates.indexOf(name) < 0) duplicates.push(name);
    seen.add(key);
  });
  if (duplicates.length) {
    ss.toast('同じ名前のサービスが2つ以上あります（' + duplicates.join('、') + '）。サービス名は重複しないようにしてください。',
      'サービス管理', 10);
    return;
  }

  const single = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const oldName = single && e && e.oldValue != null ? String(e.oldValue).trim() : '';
  const newName = single && e && e.value != null ? String(e.value).trim() : '';
  if (!oldName || !newName || oldName === newName || names.indexOf(oldName) >= 0) return;

  const changed = taskWithLock_(() => svcRenameService_(ss, oldName, newName));
  if (changed.requests || changed.tasks) {
    ss.toast('サービス名の変更を、リクエスト ' + changed.requests + '件・タスク ' + changed.tasks + '件に反映しました。',
      'サービス管理', 8);
  }
}

/** リクエスト・タスク管理の「サービス」の列で、oldName のセルを newName にする。{requests, tasks} */
function svcRenameService_(ss, oldName, newName) {
  const targets = {
    requests: [SVC_OPTIONS.requestSheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow],
    tasks: [TASK_OPTIONS.sheet, taskAllColumns_(), TASK_OPTIONS.headerRow]
  };
  const out = {requests: 0, tasks: 0};
  Object.keys(targets).forEach(kind => {
    const [name, columns, headerRow] = targets[kind];
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const col = taskColumnMap_(sheet, columns.filter(c => c.type === 'service'), headerRow).service;
    const count = sheet.getLastRow() - headerRow;
    if (!col || count < 1) return;
    sheet.getRange(headerRow + 1, col, count, 1).getDisplayValues().forEach((value, i) => {
      if (String(value[0]).trim() !== oldName) return;
      sheet.getRange(headerRow + 1 + i, col).setValue(credText_(newName));
      out[kind]++;
    });
  });
  return out;
}
