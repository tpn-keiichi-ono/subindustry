/**
 * タスク管理
 * サービスの検討（ServiceManagement.gs。リクエストをもとに考え、リクエストに名前を付けてまとめたサービス）の進捗を、
 * タスクで管理するシート「タスク管理」。
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs と同じプロジェクトに置くファイル。
 *
 * 仕組み
 * - 今は手で行を追加する。1行が1件のタスクで、左から「サービス」（リクエストに付けたサービス名）を選び、
 *   続けて「サブインダストリー」「得意先」「案件名」を
 *   連動するプルダウンで選ぶ（得意先はサブインダストリーで、案件名はサブインダストリーと得意先で絞り込む）
 * - サービスを選んだ行の連動プルダウンは、そのサービスにまとめたリクエスト（リクエスト シート。棄却したものは除く）の案件に絞る。
 *   サービスが空、またはリクエストに無いサービス名の行は、TASK_OPTIONS.sourceSheets（新FMT）のすべての案件から選ぶ
 * - タスク管理シートの行を編集するたびに、その行のプルダウンを作り直す
 *   （単純トリガーの onEdit。承認していない人の編集でも動き、トリガーの設置は要らない）
 * - サービスや左の列を選び直して、右の列の値が選択肢から外れたときは、その値を空にする
 * - サービスだけのタスク（案件を選ばない）でもよい
 * - 管理者が setupTaskSheet() をエディタから実行して、シート・見出し・プルダウンを用意する
 *   （新FMT にサブインダストリー・得意先・案件名が増えたときも、もう一度実行すると全行の選択肢を作り直す。
 *   見出しが入っているシートに足りない列があれば、右端に足す）
 * - onEdit・見出しの用意・ロックなどの共通の処理は、ServiceManagement.gs からも使う
 */

const TASK_OPTIONS = {
  sheet: 'タスク管理',
  headerRow: 1,
  // 選択肢を作るシート。見出しの行は DIFF_RULES の headerRow（DIFF_RULES に無いシートは sourceHeaderRow）
  sourceSheets: ['新FMT'],
  sourceHeaderRow: 2,
  // タスクを紐づけるサービス（リクエスト シートでリクエストに付けたサービス名から選ぶ）。
  // 選ぶと、右の連動プルダウンがそのサービスにまとめたリクエストの案件に絞られる
  service: {key: 'service', label: 'サービス', width: 200, type: 'service'},
  // 連動するプルダウン（左から順に絞り込む）。label はタスク管理シートの見出し、source は選択肢を作るシートの見出し
  cascade: [
    {key: 'subIndustry', label: 'サブインダストリー', source: 'サブインダストリー', width: 160},
    {key: 'customer', label: '得意先', source: '得意先', width: 200},
    {key: 'project', label: '案件名', source: '案件名', width: 240}
  ],
  // そのほかの列（手で入力する）
  columns: [
    {key: 'task', label: 'タスク', width: 280},
    {key: 'owner', label: '担当者', width: 120},
    {key: 'due', label: '期限', width: 100, type: 'date'},
    {key: 'status', label: '状況', width: 90, options: ['未着手', '対応中', '完了']},
    {key: 'note', label: 'メモ', width: 280}
  ],
  headerBackground: '#E9F4F1',
  maxListItems: 500,    // プルダウンの選択肢の上限（超えたときは選択肢を付けずに知らせる）
  maxEditRows: 500,     // 1回の編集（貼り付けなど）で選択肢を作り直す行数の上限
  lockWaitMs: 10000     // 単純トリガーは30秒で止まるため短めにする
};

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * タスク管理シートを用意する（無ければ作る）。見出しが空なら書き込み、全行のプルダウンを作り直す。
 * 何度実行してもよい（入力済みのタスクは消さない。見出しが違うときは書き換えずに止める）。
 */
function setupTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const {sheet, created, added} = taskEnsureSheet_(ss, TASK_OPTIONS.sheet, taskAllColumns_(), TASK_OPTIONS.headerRow);
    const cols = taskColumns_(sheet);
    const first = TASK_OPTIONS.headerRow + 1;
    const rows = sheet.getMaxRows() - TASK_OPTIONS.headerRow;
    if (rows < 1) throw new Error(TASK_OPTIONS.sheet + ' にデータの行がありません。行を追加してから実行してください。');

    const warnings = taskApplyColumnRules_(sheet, [TASK_OPTIONS.service].concat(TASK_OPTIONS.columns),
      Object.assign({service: cols.service}, cols.others), first, rows);
    const sources = taskSources_(ss);
    const cascadeWarnings = taskRefreshRows_(sheet, cols, sources, first, rows, Infinity);
    const counts = TASK_OPTIONS.cascade.map((c, i) => c.label + ' ' + taskOptionsFor_(sources.all, i, []).length + '件');
    return {created, added, counts, warnings: warnings.concat(cascadeWarnings)};
  });

  ss.toast(
    (result.created ? TASK_OPTIONS.sheet + ' シートを作りました。' : TASK_OPTIONS.sheet + ' シートのプルダウンを作り直しました。') +
    (result.added.length ? '右端に「' + result.added.join('」「') + '」の列を足しました。' : '') +
    '（選択肢：' + result.counts.join('・') + '）' + (result.warnings.length ? '\n' + result.warnings.join('\n') : ''),
    'タスク管理', 10);
}

/* ---------------- 編集したとき（単純トリガー） ---------------- */

/**
 * 単純トリガー。onEdit() はプロジェクト全体でこの1つだけにすること。
 * タスク管理・リクエストのシートの編集だけを扱い、ほかのシートではすぐ戻る
 * （差分追跡はインストール型の recordDiffEdit が別に動く）。
 * 編集した本人として動くので、まだ承認していない人の編集でもプルダウンが切り替わる。
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const handlers = [{sheet: TASK_OPTIONS.sheet, title: 'タスク管理', fn: taskHandleEdit_}];
  if (typeof SVC_OPTIONS !== 'undefined') {
    handlers.push({sheet: SVC_OPTIONS.requestSheet, title: 'サービス管理', fn: svcHandleRequestEdit_});
  }
  const handler = handlers.find(h => h.sheet === name);
  if (!handler || e.range.getLastRow() <= TASK_OPTIONS.headerRow) return;
  try {
    handler.fn(sheet, e.range, e);
  } catch (error) {
    console.warn(name + ' edit skipped: ' + (error && error.stack || error));
    try { sheet.getParent().toast(error.message, handler.title, 10); } catch (_) {}
  }
}

/** 編集した行のプルダウンを作り直し、選び直した列より右の、選択肢から外れた値を空にする。 */
function taskHandleEdit_(sheet, range) {
  const first = Math.max(range.getRow(), TASK_OPTIONS.headerRow + 1);
  const last = Math.min(range.getLastRow(), first + TASK_OPTIONS.maxEditRows - 1);
  const cols = taskColumns_(sheet);
  const edited = col => col > 0 && col >= range.getColumn() && col <= range.getLastColumn();

  // 選択肢から外れた値を空にし始める連動の列：サービスを選び直したら連動の3列すべて、
  // 連動の列を選び直したらその右から（それ以外の編集では空にせず、選択肢だけ作り直す）
  let clearFrom = Infinity;
  if (edited(cols.service)) {
    clearFrom = 0;
  } else {
    const level = cols.cascade.findIndex(edited);
    if (level >= 0) clearFrom = level + 1;
  }

  taskWithLock_(() => {
    const sources = taskSources_(sheet.getParent());
    const warnings = taskRefreshRows_(sheet, cols, sources, first, last - first + 1, clearFrom);
    if (warnings.length) sheet.getParent().toast(warnings.join('\n'), 'タスク管理', 10);
  });
}

/* ---------------- 選択肢 ---------------- */

/**
 * 選択肢を作るシートから、[サブインダストリー, 得意先, 案件名] の組を行の順に返す（すべて空の行は除く）。
 * 見比べ用に、全角・半角と空白の違いを無くしたキー（credNormalize_）も持たせる。
 */
function taskSourceRecords_(ss) {
  return taskReadSource_(ss, TASK_OPTIONS.cascade.map(c => c.source))
    .map(values => ({values, keys: values.map(credNormalize_)}));
}

/**
 * 連動プルダウンの選択肢の元。
 * all：新FMT のすべての案件。byService：サービス名（credNormalize_）ごとの、そのサービスにまとめたリクエストの案件。
 */
function taskSources_(ss) {
  return {
    all: taskSourceRecords_(ss),
    byService: typeof svcRequestRecordsByService_ === 'function' ? svcRequestRecordsByService_(ss) : new Map()
  };
}

/**
 * TASK_OPTIONS.sourceSheets（新FMT）から、指定した見出しの列の値を行の順に返す（前後の空白は除く。すべて空の行は除く）。
 * サービスのリクエストの取り込み（ServiceManagement.gs）でも使う。
 */
function taskReadSource_(ss, sourceHeaders) {
  const rows = [];
  TASK_OPTIONS.sourceSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('読み込むシート「' + name + '」が見つかりません。');
    const headerRow = DIFF_RULES[name] ? DIFF_RULES[name].headerRow : TASK_OPTIONS.sourceHeaderRow;
    const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    const cols = sourceHeaders.map(h => taskFindColumn_(headers, h, name + ' の ' + headerRow + '行目', true));
    const count = sheet.getLastRow() - headerRow;
    if (count < 1) return;

    const left = Math.min.apply(null, cols);
    const right = Math.max.apply(null, cols);
    sheet.getRange(headerRow + 1, left, count, right - left + 1).getDisplayValues().forEach(line => {
      const values = cols.map(col => String(line[col - left] == null ? '' : line[col - left]).trim());
      if (values.some(Boolean)) rows.push(values);
    });
  });
  return rows;
}

/**
 * level 番目の列の選択肢。左の列で選んでいる値（空の列では絞り込まない）に合う値を、出てきた順に重複なく返す。
 * memo を渡すと、同じ絞り込みの結果を使い回す。
 */
function taskOptionsFor_(records, level, selected, memo) {
  const keys = [];
  for (let i = 0; i < level; i++) keys.push(credNormalize_(selected[i] || ''));
  const memoKey = level + '\u0001' + keys.join('\u0001');
  if (memo && memo.has(memoKey)) return memo.get(memoKey);

  const seen = new Set();
  const out = [];
  records.forEach(record => {
    for (let i = 0; i < level; i++) {
      if (keys[i] && record.keys[i] !== keys[i]) return;
    }
    const value = record.values[level];
    if (!value || seen.has(record.keys[level])) return;
    seen.add(record.keys[level]);
    out.push(value);
  });
  if (memo) memo.set(memoKey, out);
  return out;
}

/**
 * startRow から numRows 行の連動プルダウンを作り直す（cols は taskColumns_、sources は taskSources_）。
 * 連動の列のうち clearFrom 番目（0 から）より右の列で、値が新しい選択肢に無いものは空にする（Infinity なら空にしない）。
 * 戻り値は利用者に知らせる文（選択肢が多すぎて付けられなかった列など）。
 */
function taskRefreshRows_(sheet, cols, sources, startRow, numRows, clearFrom) {
  const cascadeCols = cols.cascade;
  const levels = cascadeCols.length;
  const read = col => sheet.getRange(startRow, col, numRows, 1).getDisplayValues().map(r => String(r[0]).trim());
  const current = cascadeCols.map(read);
  const services = cols.service ? read(cols.service) : [];
  const rules = cascadeCols.map(() => []);
  const cleared = cascadeCols.map(() => []);
  const optionsMemos = new Map();   // 選択肢の元（records）ごとの、絞り込みの結果
  const ruleMemo = new Map();
  const tooMany = new Set();

  for (let r = 0; r < numRows; r++) {
    const selected = current.map(column => column[r]);
    // サービスを選んでいて、そのサービスにまとめたリクエストがあれば、その案件だけから選ぶ
    const service = credNormalize_(services[r] || '');
    const records = service && sources.byService.has(service) ? sources.byService.get(service) : sources.all;
    if (!optionsMemos.has(records)) optionsMemos.set(records, new Map());
    const optionsMemo = optionsMemos.get(records);
    for (let level = 0; level < levels; level++) {
      let options = taskOptionsFor_(records, level, selected, optionsMemo);
      // 全角・半角と空白の違いは同じ値とみなす。入っている値の書き方が選択肢と違えば、その行の選択肢を入っている書き方にする
      const key = credNormalize_(selected[level]);
      const index = key ? options.findIndex(option => credNormalize_(option) === key) : -1;
      if (index >= 0 && options[index] !== selected[level]) {
        options = options.slice();
        options[index] = selected[level];
      }
      if (level >= clearFrom && selected[level] && index < 0) {
        selected[level] = '';
        cleared[level].push(startRow + r);
      }
      if (options.length > TASK_OPTIONS.maxListItems) {
        tooMany.add(level);
        rules[level].push([null]);
      } else {
        rules[level].push([options.length ? taskListRule_(options, level, ruleMemo) : null]);
      }
    }
  }

  cleared.forEach((rows, level) => rows.forEach(row => sheet.getRange(row, cascadeCols[level]).clearContent()));
  cascadeCols.forEach((col, level) => sheet.getRange(startRow, col, numRows, 1).setDataValidations(rules[level]));

  return Array.from(tooMany).map(level => {
    const label = TASK_OPTIONS.cascade[level].label;
    return '「' + label + '」の選択肢が ' + TASK_OPTIONS.maxListItems + ' 件を超えるため、プルダウンを付けられない行があります。' +
      '左の列を先に選ぶと絞り込まれます。';
  });
}

/** 一覧から選ぶ入力規則。同じ選択肢の規則は使い回す。 */
function taskListRule_(options, level, memo) {
  const key = level + '\u0001' + options.join('\u0001');
  if (memo.has(key)) return memo.get(key);
  const label = TASK_OPTIONS.cascade[level].label;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .setHelpText(label + 'は一覧から選んでください（' + TASK_OPTIONS.sourceSheets.join('・') + ' の値です）。')
    .build();
  memo.set(key, rule);
  return rule;
}

/* ---------------- シートの形（サービス管理と共通） ---------------- */

/** タスク管理シートの列（サービス・連動の3列・そのほか。連動の3列は必須）。 */
function taskAllColumns_() {
  return [TASK_OPTIONS.service]
    .concat(TASK_OPTIONS.cascade.map(c => Object.assign({required: true}, c)))
    .concat(TASK_OPTIONS.columns);
}

/** シートが無ければ作り、見出しを用意する。{sheet, created, added: 右端に足した列の見出し} */
function taskEnsureSheet_(ss, name, columns, headerRow) {
  let sheet = ss.getSheetByName(name);
  const created = !sheet;
  if (!sheet) sheet = ss.insertSheet(name);
  return {sheet, created, added: taskEnsureHeader_(sheet, columns, headerRow)};
}

/**
 * 見出しの行が空なら見出しを書き、形を整える。
 * 見出しが入っているシートは書き換えない：必須の列（required）が無ければ止め、足りない列だけを右端に足す。
 * 戻り値は右端に足した列の見出し。
 */
function taskEnsureHeader_(sheet, columns, headerRow) {
  const where = sheet.getName() + ' の ' + headerRow + '行目';
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  let last = headers.length;
  while (last > 0 && !String(headers[last - 1]).trim()) last--;

  let start = 1;
  let targets = columns;
  if (last > 0) {
    columns.filter(c => c.required).forEach(c => taskFindColumn_(headers, c.label, where, true));
    targets = columns.filter(c => !c.required && taskHeaderCount_(headers, c.label) === 0);
    start = last + 1;
  }
  if (!targets.length) return [];

  const needed = start + targets.length - 1;
  if (sheet.getMaxColumns() < needed) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  sheet.getRange(headerRow, start, 1, targets.length).setValues([targets.map(c => c.label)])
    .setFontWeight('bold').setBackground(TASK_OPTIONS.headerBackground);
  targets.forEach((c, i) => { if (c.width) sheet.setColumnWidth(start + i, c.width); });
  if (last === 0) sheet.setFrozenRows(headerRow);
  return last > 0 ? targets.map(c => c.label) : [];
}

/** 見出しから {key: 列番号} を作る（見つからない列は 0。required の列はちょうど1つ必要）。 */
function taskColumnMap_(sheet, columns, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const where = sheet.getName() + ' の ' + headerRow + '行目';
  const out = {};
  columns.forEach(c => { out[c.key] = taskFindColumn_(headers, c.label, where, !!c.required); });
  return out;
}

/** タスク管理シートの列の位置 {service, cascade: [3列], others}。連動する3列は見出しがちょうど1つずつ必要。 */
function taskColumns_(sheet) {
  const map = taskColumnMap_(sheet, taskAllColumns_(), TASK_OPTIONS.headerRow);
  const others = {};
  TASK_OPTIONS.columns.forEach(c => { others[c.key] = map[c.key]; });
  return {service: map[TASK_OPTIONS.service.key], cascade: TASK_OPTIONS.cascade.map(c => map[c.key]), others};
}

/**
 * 日付・一覧（options）・サービス（type: 'service'）の列に、first 行目から rows 行の入力規則を付ける。
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
      rule = typeof svcServiceRule_ === 'function' ? svcServiceRule_(sheet.getParent(), !!column.allowNew) : null;
      if (!rule) {
        warnings.push('リクエスト シートが無いため、「' + column.label + '」の列にプルダウンを付けていません。importServiceRequests() を実行してください。');
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

function taskHeaderCount_(headers, label) {
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  return headers.filter(h => normalize(h) === normalize(label)).length;
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

/** ドキュメントロックの中で fn を実行する。waitMs を省くと単純トリガー向けの短い待ち時間。 */
function taskWithLock_(fn, waitMs) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(waitMs || TASK_OPTIONS.lockWaitMs)) {
    throw new Error('他の処理が実行中です。少し待ってからもう一度お試しください（選び直した値は反映されていません）。');
  }
  try {
    return fn();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}
