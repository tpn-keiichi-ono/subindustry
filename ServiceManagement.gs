/**
 * サービス管理（リクエストのとりまとめ）
 * 新FMT の「サービスのリクエスト」（Z列）を「リクエスト」シートに転記し、リクエストをもとにサービスを考える。
 * TaskManagement.gs と同じプロジェクトに置くファイル（見出しの用意・ロック・onEdit は TaskManagement.gs と共通）。
 *
 * 仕組み
 * - サービスは最初から決まっているものではなく、リクエストをもとに考える。サービスの一覧のシートは持たず、
 *   リクエストの「サービス」の列に入力した名前がサービスになる（同じ名前を付けたリクエストが1つのサービスにまとまる）
 * - リクエスト シートの左の4列（サブインダストリー・得意先・案件名・リクエスト）は、新FMT から自動で転記する：
 *   新FMT のこの4列を編集すると、単純トリガーの onEdit（TaskManagement.gs）が svcHandleSourceEdit_ で転記する。
 *   まだ無いリクエスト（得意先・案件名・リクエストの組で見分ける）は末尾に足し、サブインダストリーは新FMT に合わせる。
 *   1つのセルで得意先・案件名・リクエストを書き換えたときは、新しい行を足さずに同じ行を直す
 * - 管理者が importServiceRequests() をエディタから実行すると、シートを用意し、まとめて転記し直す（最初の1回・漏れたとき）
 * - 転記する4列は、手で変えると警告が出るように保護する（警告だけ。転記は止まらない）
 * - サービスの検討の進捗は、タスク管理シート（TaskManagement.gs）でタスクとして管理する。
 *   タスクの「サービス」はリクエストに付けたサービス名から選び、選ぶと案件の選択肢がそのサービスにまとめたリクエストの案件に絞られる
 * - リクエストでサービスを付けたとき、判断が未判断なら「サービス化検討」にする。
 *   リクエストのサービス・判断を変えたら、タスク管理の案件の選択肢も作り直す
 * - 新FMT から消えた・書き換えられたリクエストは、行を消さずに「リクエスト」のセルに注を付ける
 */

const SVC_OPTIONS = {
  requestSheet: 'リクエスト',
  headerRow: 1,
  // 取り込み元（TASK_OPTIONS.sourceSheets。新FMT）でリクエストが書かれた列の見出し（Z列）。1つのセルに1件
  requestHeader: 'サービスのリクエスト',
  // 左の4列（copied）は新FMT から転記する。サービス・判断・メモを手で入れる
  requestColumns: [
    {key: 'subIndustry', label: 'サブインダストリー', width: 160, required: true, copied: true},
    {key: 'customer', label: '得意先', width: 200, required: true, copied: true},
    {key: 'project', label: '案件名', width: 240, required: true, copied: true},
    {key: 'request', label: 'リクエスト', width: 360, required: true, copied: true},
    {key: 'service', label: 'サービス', width: 200},   // まとめる先のサービス名（自由に入力する）
    {key: 'decision', label: '判断', width: 130, options: ['未判断', 'サービス化検討', '棄却']},
    {key: 'note', label: 'メモ', width: 280},
    {key: 'importedAt', label: '取り込み日', width: 100}
  ],
  defaultDecision: '未判断',             // 取り込んだときの判断
  rejectedDecision: '棄却',              // この判断のリクエストは、タスクの案件の選択肢に入れない
  decisionWithService: 'サービス化検討',  // サービスを付けたとき、判断が未判断（または空）ならこれにする
  missingNote: '取り込み元に見つかりません',   // 新FMT から消えた・書き換えられたリクエストに付ける注の先頭
  changedNote: '新FMT でリクエストが書き換えられました',   // サービス・判断を付けたあとで書き換えられたときの注
  protectDescription: 'リクエスト：新FMT から自動で転記する列',
  lockWaitMs: 30000                       // エディタから実行する処理のロックの待ち時間
};

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * リクエスト シートを用意する（無ければ作る）。見出し・プルダウン・転記する列の保護を付ける。
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
 * リクエスト シートを用意し、新FMT のリクエストをまとめて転記する（最初の1回・転記が漏れたとき）。
 * ふだんは新FMT を編集したときに自動で転記される。
 */
function importServiceRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const ensured = svcEnsureSheet_(ss);
    const synced = svcSync_(ss, ensured.sheet, null);
    return Object.assign(ensured, synced, svcApplyRules_(ss, ensured.sheet));
  }, SVC_OPTIONS.lockWaitMs);

  const lines = ['リクエストを ' + result.added + '件追加し、' + result.updated + '件を新FMT に合わせました（新FMT のリクエスト ' + result.total + '件）。'];
  if (result.missing) lines.push('新FMT に見つからないリクエストが ' + result.missing + '件あります（「リクエスト」のセルに注を付けました）。');
  ss.toast(lines.concat(svcSetupMessage_(result).slice(1)).join('\n'), 'サービス管理', 10);
}

/* ---------------- シートの用意 ---------------- */

/** {sheet, created, added} */
function svcEnsureSheet_(ss) {
  return taskEnsureSheet_(ss, SVC_OPTIONS.requestSheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
}

/**
 * 全行のプルダウン（判断）を付け、転記する列を保護する。サービスの列は自由に入力するので、プルダウンを付けない（前の版で付けたものは外す）。
 * タスク管理シートがあれば、その「サービス」の列（リクエストに付けたサービス名から選ぶ）にもプルダウンを付ける。
 */
function svcApplyRules_(ss, sheet) {
  const warnings = [];
  const first = SVC_OPTIONS.headerRow + 1;
  const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
  if (rows > 0) {
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    warnings.push.apply(warnings, taskApplyColumnRules_(sheet, SVC_OPTIONS.requestColumns, cols, first, rows));
    if (cols.service) sheet.getRange(first, cols.service, rows, 1).clearDataValidations();
    if (cols.importedAt) sheet.getRange(first, cols.importedAt, rows, 1).setNumberFormat('yyyy/MM/dd');
    svcProtectCopiedColumns_(sheet, cols);
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

/** 転記する列を、手で変えると警告が出るように保護する（すでに保護していれば何もしない）。 */
function svcProtectCopiedColumns_(sheet, cols) {
  const protectedAlready = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).some(p => {
    try { return p.getDescription() === SVC_OPTIONS.protectDescription; } catch (_) { return false; }
  });
  if (protectedAlready) return;
  SVC_OPTIONS.requestColumns.filter(c => c.copied).forEach(c => {
    const letter = diffColumnLetter_(cols[c.key]);
    sheet.getRange(letter + (SVC_OPTIONS.headerRow + 1) + ':' + letter).protect()
      .setDescription(SVC_OPTIONS.protectDescription).setWarningOnly(true);
  });
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
 * タスク管理の「サービス」の列の入力規則。リクエスト シートの「サービス」の列（リクエストに付けたサービス名）を範囲で参照し、
 * そこに無い名前は入力できないようにする。リクエスト シートが無ければ null。
 */
function svcServiceRule_(ss) {
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) return null;
  const col = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns.filter(c => c.key === 'service'), SVC_OPTIONS.headerRow).service;
  if (!col) return null;
  const letter = diffColumnLetter_(col);
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange(letter + (SVC_OPTIONS.headerRow + 1) + ':' + letter), true)
    .setAllowInvalid(false)
    .setHelpText('「' + SVC_OPTIONS.requestSheet + '」シートでリクエストに付けたサービス名から選んでください。')
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

/* ---------------- 新FMT からの転記 ---------------- */

/** リクエストを見分けるキー（得意先・案件名・リクエスト。全角・半角と空白の違いは無視）。 */
function svcRequestKey_(customer, project, request) {
  return [customer, project, request].map(credNormalize_).join('\u0001');
}

/** 取り込み元で読む列の見出し [サブインダストリー, 得意先, 案件名, リクエスト]。 */
function svcSourceHeaders_() {
  return TASK_OPTIONS.cascade.map(c => c.source).concat([SVC_OPTIONS.requestHeader]);
}

/**
 * 新FMT のリクエストをリクエスト シートに転記する。
 * - renamed（1つのセルで得意先・案件名・リクエストを書き換えた）があれば、古いキーの行をその場で新しい値に直す
 *   （新しいキーの行がまだ無いときだけ）。リクエストの書き換えで、サービス・判断を付けたあとなら注を付ける
 * - まだ無いリクエストは末尾に足す。あるものはサブインダストリーを新FMT に合わせる
 * - 新FMT に無い行には注を付ける（新FMT に戻ったら外す。利用者が書いた注は変えない）
 * {added, updated, missing, total}
 */
function svcSync_(ss, sheet, renamed) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const copied = ['subIndustry', 'customer', 'project', 'request'];
  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd');

  const source = new Map();
  taskReadSource_(ss, svcSourceHeaders_()).forEach(values => {
    if (!values[3]) return;
    const key = svcRequestKey_(values[1], values[2], values[3]);
    if (!source.has(key)) source.set(key, values);
  });

  const lastRow = Math.max(sheet.getLastRow(), headerRow);
  const count = lastRow - headerRow;
  const existing = new Set();
  let updated = 0;
  let missing = 0;
  if (count > 0) {
    const read = key => (cols[key]
      ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
      : new Array(count).fill(''));
    const current = copied.map(read);
    const services = read('service');
    const decisions = read('decision');
    const keyOf = i => (current[3][i] ? svcRequestKey_(current[1][i], current[2][i], current[3][i]) : '');
    const keys = current[3].map((_, i) => keyOf(i));
    const changedRows = new Set();

    // 1つのセルの書き換え：古いキーの行を新しい値に直す
    if (renamed && keys.indexOf(renamed.newKey) < 0) {
      const i = keys.indexOf(renamed.oldKey);
      if (i >= 0) {
        sheet.getRange(headerRow + 1 + i, cols[copied[renamed.level]]).setValue(credText_(renamed.value));
        current[renamed.level][i] = renamed.value;
        keys[i] = keyOf(i);
        updated++;
        const judged = services[i] || (decisions[i] && decisions[i] !== SVC_OPTIONS.defaultDecision);
        if (renamed.level === 3 && judged) changedRows.add(i);
      }
    }

    // サブインダストリーを新FMT に合わせる
    keys.forEach((key, i) => {
      const values = key && source.get(key);
      if (!values || values[0] === current[0][i]) return;
      sheet.getRange(headerRow + 1 + i, cols.subIndustry).setValue(credText_(values[0]));
      updated++;
    });

    const noteRange = sheet.getRange(headerRow + 1, cols.request, count, 1);
    const notes = noteRange.getNotes().map(r => String(r[0] || ''));
    let notesChanged = false;
    const next = keys.map((key, i) => {
      if (!key) return notes[i];
      existing.add(key);
      const missingNote = notes[i].indexOf(SVC_OPTIONS.missingNote) === 0;
      if (source.has(key)) {
        if (changedRows.has(i) && (!notes[i] || missingNote)) {
          notesChanged = true;
          return SVC_OPTIONS.changedNote + '（' + today + '）。サービス・判断を見直してください。';
        }
        if (!missingNote) return notes[i];
        notesChanged = true;
        return '';
      }
      missing++;
      if (notes[i]) return notes[i];   // 前回から見つからないまま、または利用者が書いた注
      notesChanged = true;
      return SVC_OPTIONS.missingNote + '（' + today + ' に気づきました）。' +
        TASK_OPTIONS.sourceSheets.join('・') + ' で書き換えか削除された可能性があります。';
    });
    if (notesChanged) noteRange.setNotes(next.map(n => [n]));
  }

  const fresh = Array.from(source.entries()).filter(([key]) => !existing.has(key)).map(([, values]) => values);
  if (fresh.length) {
    const start = lastRow + 1;
    const end = start + fresh.length - 1;
    if (end > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), end - sheet.getMaxRows());
    const width = Math.max.apply(null, Object.keys(cols).map(k => cols[k]));
    const now = new Date();
    const rows = fresh.map(values => {
      const row = new Array(width).fill('');
      // 新FMT の文字列は数式にならないよう、先頭に ' を付けて書く
      copied.forEach((key, i) => { row[cols[key] - 1] = credText_(values[i]); });
      if (cols.decision) row[cols.decision - 1] = SVC_OPTIONS.defaultDecision;
      if (cols.importedAt) row[cols.importedAt - 1] = now;
      return row;
    });
    sheet.getRange(start, 1, rows.length, width).setValues(rows);
    if (cols.importedAt) sheet.getRange(start, cols.importedAt, rows.length, 1).setNumberFormat('yyyy/MM/dd');
  }
  return {added: fresh.length, updated, missing, total: source.size};
}

/* ---------------- 編集したとき（TaskManagement.gs の onEdit から） ---------------- */

/**
 * 新FMT（取り込み元）を編集したとき：サブインダストリー・得意先・案件名・リクエストの列を変えたら、リクエスト シートに転記する。
 * リクエスト シートが無い・見出しが足りないときは何もしない（importServiceRequests() を実行すると知らせる）。
 */
function svcHandleSourceEdit_(sheet, range, e) {
  const ss = sheet.getParent();
  const request = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!request) return;
  const headerRow = taskSourceHeaderRow_(sheet.getName());
  if (range.getLastRow() <= headerRow) return;

  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const cols = svcSourceHeaders_().map(h => (taskHeaderCount_(headers, h) === 1 ? taskFindColumn_(headers, h, '', true) : 0));
  if (cols.some(col => !col)) return;
  if (!cols.some(col => col >= range.getColumn() && col <= range.getLastColumn())) return;

  // 1つのセルで得意先・案件名・リクエストを書き換えたときは、リクエスト シートの同じ行を直す（新しい行を足さない）
  let renamed = null;
  const level = cols.indexOf(range.getColumn());
  if (range.getNumRows() === 1 && range.getNumColumns() === 1 && level >= 1 && e && e.oldValue != null) {
    const left = Math.min.apply(null, cols);
    const line = sheet.getRange(range.getRow(), left, 1, Math.max.apply(null, cols) - left + 1).getDisplayValues()[0];
    const after = cols.map(col => String(line[col - left]).trim());
    const before = after.slice();
    before[level] = String(e.oldValue).trim();
    if (after[3] && before[3]) {
      renamed = {
        level,
        value: after[level],
        oldKey: svcRequestKey_(before[1], before[2], before[3]),
        newKey: svcRequestKey_(after[1], after[2], after[3])
      };
    }
  }

  try {
    taskWithLock_(() => {
      const result = svcSync_(ss, request, renamed);
      // 得意先・案件名などを直したときは、タスク管理の案件の選択肢も作り直す（値は空にしない）
      if (result.updated) svcRefreshTaskRows_(ss);
    });
  } catch (error) {
    throw new Error('リクエスト シートへの転記ができませんでした（' + error.message + '）。' +
      'あとで importServiceRequests() を実行するか、もう一度編集してください。');
  }
}

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
