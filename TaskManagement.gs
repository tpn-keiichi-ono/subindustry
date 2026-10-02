/**
 * タスク管理
 * 追跡シート（新FMT）の行（案件）ごとのタスクを、別のシート「タスク管理」で管理する。
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs と同じプロジェクトに置くファイル。
 *
 * 仕組み
 * - 今は手で行を追加する。1行が1件のタスクで、左から「サブインダストリー」「得意先」「案件名」を
 *   連動するプルダウンで選ぶ（得意先はサブインダストリーで、案件名はサブインダストリーと得意先で絞り込む）
 * - 選択肢は TASK_OPTIONS.sourceSheets（新FMT）の値から作る。タスク管理シートの行を編集するたびに、
 *   その行のプルダウンを作り直す（単純トリガーの onEdit。承認していない人の編集でも動き、トリガーの設置は要らない）
 * - 左の列を選び直して、右の列の値が選択肢から外れたときは、右の列を空にする
 * - 管理者が setupTaskSheet() をエディタから実行して、シート・見出し・プルダウンを用意する
 *   （新FMT にサブインダストリー・得意先・案件名が増えたときも、もう一度実行すると全行の選択肢を作り直す）
 */

const TASK_OPTIONS = {
  sheet: 'タスク管理',
  headerRow: 1,
  // 選択肢を作るシート。見出しの行は DIFF_RULES の headerRow（DIFF_RULES に無いシートは sourceHeaderRow）
  sourceSheets: ['新FMT'],
  sourceHeaderRow: 2,
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
    let sheet = ss.getSheetByName(TASK_OPTIONS.sheet);
    const created = !sheet;
    if (!sheet) sheet = ss.insertSheet(TASK_OPTIONS.sheet);
    taskEnsureHeader_(sheet);

    const cols = taskColumns_(sheet);
    const first = TASK_OPTIONS.headerRow + 1;
    const rows = sheet.getMaxRows() - TASK_OPTIONS.headerRow;
    if (rows < 1) throw new Error(TASK_OPTIONS.sheet + ' にデータの行がありません。行を追加してから実行してください。');

    TASK_OPTIONS.columns.forEach(column => {
      const col = cols.others[column.key];
      if (!col) return;
      const range = sheet.getRange(first, col, rows, 1);
      if (column.type === 'date') {
        range.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireDate().setAllowInvalid(false).setHelpText('日付を入力してください。').build());
      } else if (column.options) {
        range.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInList(column.options, true).setAllowInvalid(false)
          .setHelpText(column.label + 'は一覧から選んでください。').build());
      }
    });

    const records = taskSourceRecords_(ss);
    const warnings = taskRefreshRows_(sheet, cols.cascade, records, first, rows, -1);
    const counts = TASK_OPTIONS.cascade.map((c, i) => c.label + ' ' + taskOptionsFor_(records, i, []).length + '件');
    return {created, counts, warnings};
  });

  ss.toast(
    (result.created ? TASK_OPTIONS.sheet + ' シートを作りました。' : TASK_OPTIONS.sheet + ' シートのプルダウンを作り直しました。') +
    '（選択肢：' + result.counts.join('・') + '）' + (result.warnings.length ? '\n' + result.warnings.join('\n') : ''),
    'タスク管理', 10);
}

/* ---------------- 編集したとき（単純トリガー） ---------------- */

/**
 * 単純トリガー。onEdit() はプロジェクト全体でこの1つだけにすること。
 * タスク管理シートの編集だけを扱い、ほかのシートではすぐ戻る（差分追跡はインストール型の recordDiffEdit が別に動く）。
 * 編集した本人として動くので、まだ承認していない人の編集でもプルダウンが切り替わる。
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TASK_OPTIONS.sheet) return;
  if (e.range.getLastRow() <= TASK_OPTIONS.headerRow) return;
  try {
    taskHandleEdit_(sheet, e.range);
  } catch (error) {
    console.warn('Task sheet edit skipped: ' + (error && error.stack || error));
    try { sheet.getParent().toast(error.message, 'タスク管理', 10); } catch (_) {}
  }
}

/** 編集した行のプルダウンを作り直し、選び直した列より右の、選択肢から外れた値を空にする。 */
function taskHandleEdit_(sheet, range) {
  const first = Math.max(range.getRow(), TASK_OPTIONS.headerRow + 1);
  const last = Math.min(range.getLastRow(), first + TASK_OPTIONS.maxEditRows - 1);
  const cols = taskColumns_(sheet);

  // 選び直した連動列のうち、いちばん左のもの（連動列以外の編集なら -1：空にはせず、選択肢だけ作り直す）
  let changed = -1;
  cols.cascade.forEach((col, i) => {
    if (changed < 0 && col >= range.getColumn() && col <= range.getLastColumn()) changed = i;
  });

  taskWithLock_(() => {
    const records = taskSourceRecords_(sheet.getParent());
    const warnings = taskRefreshRows_(sheet, cols.cascade, records, first, last - first + 1, changed);
    if (warnings.length) sheet.getParent().toast(warnings.join('\n'), 'タスク管理', 10);
  });
}

/* ---------------- 選択肢 ---------------- */

/**
 * 選択肢を作るシートから、[サブインダストリー, 得意先, 案件名] の組を行の順に返す（すべて空の行は除く）。
 * 見比べ用に、全角・半角と空白の違いを無くしたキー（credNormalize_）も持たせる。
 */
function taskSourceRecords_(ss) {
  const records = [];
  TASK_OPTIONS.sourceSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('選択肢を作るシート「' + name + '」が見つかりません。');
    const headerRow = DIFF_RULES[name] ? DIFF_RULES[name].headerRow : TASK_OPTIONS.sourceHeaderRow;
    const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    const cols = TASK_OPTIONS.cascade.map(c => taskFindColumn_(headers, c.source, name + ' の ' + headerRow + '行目', true));
    const count = sheet.getLastRow() - headerRow;
    if (count < 1) return;

    const left = Math.min.apply(null, cols);
    const right = Math.max.apply(null, cols);
    sheet.getRange(headerRow + 1, left, count, right - left + 1).getDisplayValues().forEach(line => {
      const values = cols.map(col => String(line[col - left] == null ? '' : line[col - left]).trim());
      if (values.some(Boolean)) records.push({values, keys: values.map(credNormalize_)});
    });
  });
  return records;
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
 * startRow から numRows 行の連動プルダウンを作り直す。
 * changedLevel より右の列で、値が新しい選択肢に無いものは空にする（-1 なら空にしない）。
 * 戻り値は利用者に知らせる文（選択肢が多すぎて付けられなかった列など）。
 */
function taskRefreshRows_(sheet, cascadeCols, records, startRow, numRows, changedLevel) {
  const levels = cascadeCols.length;
  const current = cascadeCols.map(col => sheet.getRange(startRow, col, numRows, 1).getDisplayValues().map(r => String(r[0]).trim()));
  const rules = cascadeCols.map(() => []);
  const cleared = cascadeCols.map(() => []);
  const optionsMemo = new Map();
  const ruleMemo = new Map();
  const tooMany = new Set();

  for (let r = 0; r < numRows; r++) {
    const selected = current.map(column => column[r]);
    for (let level = 0; level < levels; level++) {
      let options = taskOptionsFor_(records, level, selected, optionsMemo);
      // 全角・半角と空白の違いは同じ値とみなす。入っている値の書き方が選択肢と違えば、その行の選択肢を入っている書き方にする
      const key = credNormalize_(selected[level]);
      const index = key ? options.findIndex(option => credNormalize_(option) === key) : -1;
      if (index >= 0 && options[index] !== selected[level]) {
        options = options.slice();
        options[index] = selected[level];
      }
      if (level > changedLevel && changedLevel >= 0 && selected[level] && index < 0) {
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

/* ---------------- シートの形 ---------------- */

/** 見出しの行が空なら見出しを書き、形を整える。空でなければ書き換えない（違う見出しは taskColumns_ で止める）。 */
function taskEnsureHeader_(sheet) {
  const labels = TASK_OPTIONS.cascade.concat(TASK_OPTIONS.columns).map(c => c.label);
  const widths = TASK_OPTIONS.cascade.concat(TASK_OPTIONS.columns).map(c => c.width);
  if (sheet.getMaxColumns() < labels.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), labels.length - sheet.getMaxColumns());

  const header = sheet.getRange(TASK_OPTIONS.headerRow, 1, 1, Math.max(sheet.getLastColumn(), labels.length));
  if (header.getDisplayValues()[0].some(v => String(v).trim())) return;

  const range = sheet.getRange(TASK_OPTIONS.headerRow, 1, 1, labels.length);
  range.setValues([labels]);
  range.setFontWeight('bold').setBackground(TASK_OPTIONS.headerBackground);
  sheet.setFrozenRows(TASK_OPTIONS.headerRow);
  widths.forEach((width, i) => { if (width) sheet.setColumnWidth(i + 1, width); });
}

/** タスク管理シートの列の位置。連動する3列は見出しがちょうど1つずつ必要。 */
function taskColumns_(sheet) {
  const headers = sheet.getRange(TASK_OPTIONS.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const where = TASK_OPTIONS.sheet + ' の ' + TASK_OPTIONS.headerRow + '行目';
  const others = {};
  TASK_OPTIONS.columns.forEach(c => { others[c.key] = taskFindColumn_(headers, c.label, where, false); });
  return {
    cascade: TASK_OPTIONS.cascade.map(c => taskFindColumn_(headers, c.label, where, true)),
    others
  };
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

function taskWithLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(TASK_OPTIONS.lockWaitMs)) {
    throw new Error('他の処理が実行中のため、プルダウンを更新できませんでした。少し待ってから選び直してください。');
  }
  try {
    return fn();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}
