/**
 * サービス管理（リクエストのとりまとめ）
 * 新FMT の「サービスのリクエスト」（Z列）を「リクエスト」シートに一覧にし、リクエストをもとにサービスを考える。
 * TaskManagement.gs と同じプロジェクトに置くファイル（見出しの用意・ロック・onEdit は TaskManagement.gs と共通）。
 *
 * 仕組み
 * - サービスは最初から決まっているものではなく、リクエストをもとに考える。サービスの一覧のシートは持たず、
 *   リクエストの「サービス」の列に付けた名前がサービスになる（同じ名前を付けたリクエストが1つのサービスにまとまる）
 * - importServiceRequests() をエディタから実行すると、新FMT のリクエストのうち、まだ一覧に無いもの
 *   （得意先・案件名・リクエストの組で見分ける）だけを追記する
 * - 「サービス」の列のプルダウンは、同じ列（ほかのリクエストに付けた名前）を範囲で参照する。新しい名前も入力できる
 * - サービスの検討の進捗は、タスク管理シート（TaskManagement.gs）でタスクとして管理する。
 *   タスクの「サービス」はリクエストに付けたサービス名から選び、選ぶと案件の選択肢がそのサービスにまとめたリクエストの案件に絞られる
 * - 単純トリガーの onEdit（TaskManagement.gs）から：リクエストでサービスを付けたとき、判断が未判断なら「サービス化検討」にする。
 *   リクエストのサービス・判断を変えたら、タスク管理の案件の選択肢も作り直す
 * - 取り込み元から消えた・書き換えられたリクエストは、行を消さずに「リクエスト」のセルに注を付ける
 */

const SVC_OPTIONS = {
  requestSheet: 'リクエスト',
  headerRow: 1,
  // 取り込み元（TASK_OPTIONS.sourceSheets。新FMT）でリクエストが書かれた列の見出し（Z列）。1つのセルに1件
  requestHeader: 'サービスのリクエスト',
  // 左の4列は取り込みで入れる（手で変えない）。サービス・判断・メモを手で入れる
  requestColumns: [
    {key: 'subIndustry', label: 'サブインダストリー', width: 160, required: true},
    {key: 'customer', label: '得意先', width: 200, required: true},
    {key: 'project', label: '案件名', width: 240, required: true},
    {key: 'request', label: 'リクエスト', width: 360, required: true},
    // まとめる先のサービス名。新しい名前を入力するか、ほかのリクエストに付けた名前から選ぶ
    {key: 'service', label: 'サービス', width: 200, type: 'service', allowNew: true},
    {key: 'decision', label: '判断', width: 130, options: ['未判断', 'サービス化検討', '棄却']},
    {key: 'note', label: 'メモ', width: 280},
    {key: 'importedAt', label: '取り込み日', width: 100}
  ],
  defaultDecision: '未判断',             // 取り込んだときの判断
  rejectedDecision: '棄却',              // この判断のリクエストは、タスクの案件の選択肢に入れない
  decisionWithService: 'サービス化検討',  // サービスを付けたとき、判断が未判断（または空）ならこれにする
  missingNote: '取り込み元に見つかりません',   // 取り込み元から消えた・書き換えられたリクエストに付ける注の先頭
  lockWaitMs: 30000                       // エディタから実行する処理のロックの待ち時間
};

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * リクエスト シートを用意する（無ければ作る）。見出しとプルダウンを付ける。
 * タスク管理シートがあれば「サービス」の列（無ければ右端に足す）にもプルダウンを付ける。何度実行してもよい。
 */
function setupRequestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const ensured = svcEnsureSheet_(ss);
    return Object.assign(ensured, svcApplyRules_(ss, ensured.sheet));
  }, SVC_OPTIONS.lockWaitMs);
  ss.toast(svcSetupMessage_(result).join('\n'), 'サービス管理', 10);
}

/**
 * 新FMT のリクエスト（Z列）のうち、まだリクエスト シートに無いものを追記する。
 * シートが無ければ作る。追記したあと、全行のプルダウンを付け直す。
 */
function importServiceRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const ensured = svcEnsureSheet_(ss);
    const imported = svcImport_(ss, ensured.sheet);
    return Object.assign(ensured, imported, svcApplyRules_(ss, ensured.sheet));
  }, SVC_OPTIONS.lockWaitMs);

  const lines = ['リクエストを ' + result.added + '件追加しました（取り込み元のリクエスト ' + result.total + '件）。'];
  if (result.missing) lines.push('取り込み元に見つからないリクエストが ' + result.missing + '件あります（「リクエスト」のセルに注を付けました）。');
  ss.toast(lines.concat(svcSetupMessage_(result).slice(1)).join('\n'), 'サービス管理', 10);
}

/* ---------------- シートの用意 ---------------- */

/** {sheet, created, added} */
function svcEnsureSheet_(ss) {
  return taskEnsureSheet_(ss, SVC_OPTIONS.requestSheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
}

/** 全行のプルダウン（サービス・判断）を付ける。タスク管理シートがあれば、その「サービス」の列にも付ける。 */
function svcApplyRules_(ss, sheet) {
  const warnings = [];
  const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
  if (rows > 0) {
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    warnings.push.apply(warnings, taskApplyColumnRules_(sheet, SVC_OPTIONS.requestColumns, cols, SVC_OPTIONS.headerRow + 1, rows));
    if (cols.importedAt) sheet.getRange(SVC_OPTIONS.headerRow + 1, cols.importedAt, rows, 1).setNumberFormat('yyyy/MM/dd');
  }

  let taskAdded = [];
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  if (task) {
    taskAdded = taskEnsureHeader_(task, taskAllColumns_(), TASK_OPTIONS.headerRow);
    const taskRows = task.getMaxRows() - TASK_OPTIONS.headerRow;
    if (taskRows > 0) {
      taskApplyColumnRules_(task, [TASK_OPTIONS.service], {service: taskColumns_(task).service}, TASK_OPTIONS.headerRow + 1, taskRows);
    }
  }
  return {warnings, taskAdded};
}

/** 用意した結果の知らせ（1行目はシートを作った・付け直した）。 */
function svcSetupMessage_(result) {
  const name = SVC_OPTIONS.requestSheet;
  const lines = [result.created ? '「' + name + '」シートを作りました。' : '「' + name + '」シートのプルダウンを付け直しました。'];
  if (result.added.length) lines.push(name + ' の右端に「' + result.added.join('」「') + '」の列を足しました。');
  if (result.taskAdded.length) lines.push(TASK_OPTIONS.sheet + ' の右端に「' + result.taskAdded.join('」「') + '」の列を足しました。');
  return lines.concat(result.warnings);
}

/**
 * サービスを選ぶ列の入力規則。リクエスト シートの「サービス」の列（リクエストに付けたサービス名）を範囲で参照する。
 * allowNew：リクエスト シート自身の列。同じ列を参照するので新しい名前もその場で範囲に入るが、念のため範囲に無い値も入力できるようにする。
 * タスク管理（TaskManagement.gs）では、リクエストに付けた名前からだけ選べるようにする。リクエスト シートが無ければ null。
 */
function svcServiceRule_(ss, allowNew) {
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) return null;
  const col = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns.filter(c => c.key === 'service'), SVC_OPTIONS.headerRow).service;
  if (!col) return null;
  const letter = diffColumnLetter_(col);
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange(letter + (SVC_OPTIONS.headerRow + 1) + ':' + letter), true)
    .setAllowInvalid(!!allowNew)
    .setHelpText(allowNew
      ? 'サービス名を入力するか、ほかのリクエストに付けたサービス名から選んでください。同じ名前を付けたリクエストが1つのサービスにまとまります。'
      : '「' + SVC_OPTIONS.requestSheet + '」シートでリクエストに付けたサービス名から選んでください。')
    .build();
}

/**
 * サービス名（credNormalize_）ごとに、そのサービスにまとめたリクエストの案件 [サブインダストリー, 得意先, 案件名] を返す
 * （棄却したリクエストは除く）。タスク管理の連動プルダウン（TaskManagement.gs の taskSources_）で使う。
 */
function svcRequestRecordsByService_(ss) {
  const out = new Map();
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) return out;
  const count = sheet.getLastRow() - SVC_OPTIONS.headerRow;
  if (count < 1) return out;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
  if (!cols.service) return out;

  const keys = ['subIndustry', 'customer', 'project', 'service', 'decision'].filter(k => cols[k]);
  const left = Math.min.apply(null, keys.map(k => cols[k]));
  const right = Math.max.apply(null, keys.map(k => cols[k]));
  sheet.getRange(SVC_OPTIONS.headerRow + 1, left, count, right - left + 1).getDisplayValues().forEach(line => {
    const get = key => (cols[key] ? String(line[cols[key] - left]).trim() : '');
    const service = get('service');
    if (!service || get('decision') === SVC_OPTIONS.rejectedDecision) return;
    const values = [get('subIndustry'), get('customer'), get('project')];
    const key = credNormalize_(service);
    if (!out.has(key)) out.set(key, []);
    out.get(key).push({values, keys: values.map(credNormalize_)});
  });
  return out;
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

/**
 * リクエストのサービス・判断を変えたとき：
 * サービスを付けた行は、判断が未判断（または空）なら「サービス化検討」にする。
 * まとめ方が変わるので、タスク管理の案件の選択肢も作り直す（値は空にしない）。
 */
function svcHandleRequestEdit_(sheet, range) {
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
  const edited = col => col > 0 && col >= range.getColumn() && col <= range.getLastColumn();
  if (!edited(cols.service) && !edited(cols.decision)) return;

  const first = Math.max(range.getRow(), SVC_OPTIONS.headerRow + 1);
  const last = Math.min(range.getLastRow(), first + TASK_OPTIONS.maxEditRows - 1);
  const rows = last - first + 1;
  taskWithLock_(() => {
    if (edited(cols.service) && cols.decision) {
      const services = sheet.getRange(first, cols.service, rows, 1).getDisplayValues();
      const decisions = sheet.getRange(first, cols.decision, rows, 1).getDisplayValues();
      services.forEach((service, i) => {
        const decision = String(decisions[i][0]).trim();
        if (String(service[0]).trim() && (!decision || decision === SVC_OPTIONS.defaultDecision)) {
          sheet.getRange(first + i, cols.decision).setValue(SVC_OPTIONS.decisionWithService);
        }
      });
    }
    svcRefreshTaskRows_(sheet.getParent());
  });
}

/** タスク管理の入力済みの行（最後の行まで）の連動プルダウンを作り直す。値は空にしない。 */
function svcRefreshTaskRows_(ss) {
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  if (!task) return;
  const rows = task.getLastRow() - TASK_OPTIONS.headerRow;
  if (rows < 1) return;
  taskRefreshRows_(task, taskColumns_(task), taskSources_(ss), TASK_OPTIONS.headerRow + 1, rows, Infinity);
}
