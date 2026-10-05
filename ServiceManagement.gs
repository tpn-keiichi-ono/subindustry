/**
 * サービス管理（リクエストのとりまとめ）
 * 新FMT の「サービスのリクエスト」（Z列）から、検討するリクエストを「リクエスト」シートに選び、リクエストをもとにサービス案を考える。
 * TaskManagement.gs と同じプロジェクトに置くファイル（見出しの用意・ロックは TaskManagement.gs と共通）。
 *
 * 仕組み
 * - リクエスト シートの1行が1件のリクエスト。A〜D（サブインダストリー → 得意先 → 案件名 → リクエスト）はすべて連動プルダウンで、
 *   候補は新FMT から自動で作る。左の列から順に選び、左の列が空の列のプルダウンには「（先に〜を選んでください）」だけが出る。
 *   左から選んだ値で1つに決まる次の列（B・C）は自動で入る（例：サブインダストリーを選んで得意先が1つなら得意先）。
 *   D（リクエスト）とその右の列には、スクリプトは値を入れない（人が選ぶ・入力する）。
 *   候補は「新FMT でリクエストが入っていて、ほかの行でまだ選んでいない案件」だけ。
 *   候補が無くなると、プルダウンには「（すべて選択済み）」だけが出る（リクエストをすべて選んだことが分かる）
 * - サービスは最初から決まっているものではなく、リクエストをもとに考える。リクエストの「サービス案」に入力した名前がサービス案になる
 *   （同じ名前を付けたリクエストが1つのサービス案にまとまる）
 * - 選んだ行は、新FMT の変更に合わせる（単純トリガーの onEdit）：サブインダストリーは同じ案件の値にそろえ、
 *   1つのセルで得意先・案件名・リクエストを書き換えたときは同じ行の値を直す。新FMT から消えたリクエストには注を付ける。
 *   新FMT にリクエストが増えたり減ったりしたら、候補も作り直す
 * - 管理者が setupRequestSheet() をエディタから実行して、シート・見出し・プルダウンを用意する（候補をまとめて作り直すときも）。
 *   前の版で全件を自動で取り込んでいた行のうち、手を付けていない行は removeUntouchedRequests() で削除して候補に戻せる
 * - サービス案ごとの検討の進捗は、タスク管理シート（TaskManagement.gs）でタスクとして管理する
 */

const SVC_OPTIONS = {
  requestSheet: 'リクエスト',
  headerRow: 1,
  // リクエストを読むシート。見出しの行は DIFF_RULES の headerRow（DIFF_RULES に無いシートは sourceHeaderRow）。新FMT2 も使うときは足す
  sourceSheets: ['新FMT'],
  sourceHeaderRow: 2,
  // 読む列の見出し（リクエスト シートの A〜D の順）。リクエストは Z列。1つのセルに1件
  sourceHeaders: ['サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'],
  // cascade：連動プルダウンの順（0 から）。aliases は前の版の見出し
  requestColumns: [
    {key: 'subIndustry', label: 'サブインダストリー', width: 160, required: true, cascade: 0},
    {key: 'customer', label: '得意先', width: 200, required: true, cascade: 1},
    {key: 'project', label: '案件名', width: 240, required: true, cascade: 2},
    {key: 'request', label: 'リクエスト', width: 360, required: true, cascade: 3},
    {key: 'service', label: 'サービス案', width: 220, aliases: ['サービス']},   // まとめる先のサービス案の名前（自由に入力する）
    {key: 'decision', label: '判断', width: 130, options: ['未判断', 'サービス化検討', '棄却']},
    {key: 'note', label: 'メモ', width: 280},
    {key: 'addedAt', label: '追加日', width: 100, aliases: ['取り込み日']}
  ],
  allSelected: '（すべて選択済み）',      // 候補をすべて選んだときにプルダウンに出す文字
  noRequests: '（新FMT にリクエストがありません）',   // 新FMT にリクエストが1件も無いときにプルダウンに出す文字
  chooseLeftFirst: '（先に{label}を選んでください）',  // 左の列が空のときに、右の列のプルダウンに出す文字（{label} は空の列の見出し）
  maxListItems: 500,                     // プルダウンの候補の上限（超えたときは候補を付けずに知らせる）
  defaultDecision: '未判断',             // まだ判断していないことを表す判断（空と同じに扱う）
  missingNote: '取り込み元に見つかりません',   // 新FMT から消えた・書き換えられたリクエストに付ける注の先頭
  changedNote: '新FMT でリクエストが書き換えられました',   // サービス案・判断を付けたあとで書き換えられたときの注
  // 前の版で付けた保護（外す）
  legacyProtectDescriptions: ['リクエスト：新FMT から自動で転記する列', 'リクエスト：案件を選ぶと自動で入る列'],
  lockWaitMs: 30000,                      // エディタから実行する処理のロックの待ち時間
  // リクエスト シートの編集（onEdit）でロックを待つ時間。続けて選ぶと前の選択の処理を待つので長めにする（単純トリガーは30秒で止まる）
  editLockWaitMs: 20000
};

/** 連動プルダウンの列の key（左から）。 */
function svcCascadeKeys_() {
  return SVC_OPTIONS.requestColumns.filter(c => c.cascade != null).sort((a, b) => a.cascade - b.cascade).map(c => c.key);
}

/** 連動プルダウンの列の見出し（左から）。 */
function svcCascadeLabels_() {
  return svcCascadeKeys_().map(key => SVC_OPTIONS.requestColumns.find(c => c.key === key).label);
}

/** level 番目の連動列が空のときに、その右の列のプルダウンに出す文字。 */
function svcChooseFirstText_(level) {
  return SVC_OPTIONS.chooseLeftFirst.replace('{label}', svcCascadeLabels_()[level]);
}

/** プルダウンの案内の文字（選んでも入らない値）なら、選んだ人に知らせる文を返す。案内の文字でなければ空。 */
function svcPlaceholderMessage_(value, hasRecords) {
  if (value === SVC_OPTIONS.allSelected || value === SVC_OPTIONS.noRequests) {
    return hasRecords ? 'リクエストのある案件は、すべて選んでいます。' : SVC_OPTIONS.noRequests.replace(/[（）]/g, '') + '。';
  }
  const labels = svcCascadeLabels_();
  const level = labels.findIndex((label, i) => value === svcChooseFirstText_(i));
  return level >= 0 ? '先に' + labels[level] + 'を選んでください（左の列から順に選びます）。' : '';
}

/* ---------------- 編集したとき（単純トリガー） ---------------- */

/**
 * 単純トリガー。onEdit() はプロジェクト全体でこの1つだけにすること。
 * 新FMT とリクエスト シートの編集だけを扱い、ほかのシートではすぐ戻る
 * （差分追跡はインストール型の recordDiffEdit が別に動く）。編集した本人として、承認なしで動く。
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  let handler = null;
  if (name === SVC_OPTIONS.requestSheet) handler = svcHandleRequestEdit_;
  else if (SVC_OPTIONS.sourceSheets.indexOf(name) >= 0) handler = svcHandleSourceEdit_;
  if (!handler || e.range.getLastRow() <= SVC_OPTIONS.headerRow) return;
  try {
    handler(sheet, e.range, e);
  } catch (error) {
    console.warn(name + ' edit skipped: ' + (error && error.stack || error));
    try { sheet.getParent().toast(error.message, 'サービス管理', 10); } catch (_) {}
  }
}

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * リクエスト シートを用意する（無ければ作る）。見出し・プルダウンを付け、選んだ行を新FMT に合わせてから、全行の候補を作り直す。
 * タスク管理シートがあれば「サービス案」の列（無ければ右端に足す）にもプルダウンを付ける。何度実行してもよい。
 */
function setupRequestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = taskWithLock_(() => {
    const ensured = svcEnsureSheet_(ss);
    const records = svcSourceRecords_(ss);
    const synced = svcSync_(ss, ensured.sheet, null, records);
    const rules = svcApplyRules_(ss, ensured.sheet);
    const refreshed = svcRefreshRequestRows_(ensured.sheet, records, null);
    return Object.assign(ensured, synced, rules, refreshed, {warnings: rules.warnings.concat(refreshed.warnings)});
  }, SVC_OPTIONS.lockWaitMs);

  const name = SVC_OPTIONS.requestSheet;
  const lines = [result.created ? '「' + name + '」シートを作りました。' : '「' + name + '」シートのプルダウンを作り直しました。'];
  lines.push(svcCountMessage_(result));
  if (result.missing) lines.push('新FMT に見つからないリクエストが ' + result.missing + '件あります（「リクエスト」のセルに注を付けました）。');
  if (result.added.length) lines.push(name + ' の右端に「' + result.added.join('」「') + '」の列を足しました。');
  if (result.taskAdded.length) lines.push(TASK_OPTIONS.sheet + ' の右端に「' + result.taskAdded.join('」「') + '」の列を足しました。');
  ss.toast(lines.concat(result.warnings, result.messages).join('\n'), 'サービス管理', 10);
}

/**
 * 前の版で全件を自動で取り込んだまま手を付けていない行（サービス案・メモが空で、判断が未判断か空）を削除し、その案件をまた候補に戻す。
 * 削除する前に件数を見せて確かめる。サービス案・判断・メモのどれかが入っている行は残す。
 */
function removeUntouchedRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) throw new Error('「' + SVC_OPTIONS.requestSheet + '」シートがありません。setupRequestSheet() を実行してください。');
  const ui = SpreadsheetApp.getUi();
  const count = svcUntouchedRows_(sheet).length;
  if (!count) {
    ui.alert('サービス管理', '削除できる行はありません（どの行も、サービス案・判断・メモのどれかが入っています）。', ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert('サービス管理',
    'サービス案・メモが空で、判断が「' + SVC_OPTIONS.defaultDecision + '」の行が ' + count + '件あります。\n' +
    'これらの行を削除して、その案件をまた候補に戻しますか？（サービス案・判断・メモのどれかが入っている行は残します）',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const result = taskWithLock_(() => {
    const rows = svcUntouchedRows_(sheet);   // ロックの中で数え直す
    // すべての行を消すことはできないので、足りなければ空の行を足しておく
    if (sheet.getMaxRows() - rows.length <= Math.max(sheet.getFrozenRows(), SVC_OPTIONS.headerRow)) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    }
    // 下の行から、続いている行はまとめて削除する
    for (let end = rows.length - 1; end >= 0;) {
      let start = end;
      while (start > 0 && rows[start - 1] === rows[start] - 1) start--;
      sheet.deleteRows(rows[start], end - start + 1);
      end = start - 1;
    }
    return Object.assign({removed: rows.length}, svcRefreshRequestRows_(sheet, svcSourceRecords_(ss), null));
  }, SVC_OPTIONS.lockWaitMs);
  ss.toast(result.removed + '行を削除しました。' + svcCountMessage_(result), 'サービス管理', 10);
}

/** 手を付けていない行（リクエストがあり、サービス案・メモが空で、判断が未判断か空）の行番号（上から）。 */
function svcUntouchedRows_(sheet) {
  const headerRow = SVC_OPTIONS.headerRow;
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return [];
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const read = key => (cols[key]
    ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : new Array(count).fill(''));
  const requests = read('request');
  const services = read('service');
  const notes = read('note');
  const decisions = read('decision');
  const rows = [];
  requests.forEach((request, i) => {
    if (!request || services[i] || notes[i]) return;
    if (decisions[i] && decisions[i] !== SVC_OPTIONS.defaultDecision) return;
    rows.push(headerRow + 1 + i);
  });
  return rows;
}

/** 選んだ件数の知らせ（{total, selected, remaining, unselectable, untouched}）。 */
function svcCountMessage_(result) {
  if (!result.total) {
    return '新FMT にリクエストが見つかりません。新FMT の「' + SVC_OPTIONS.sourceHeaders[3] + '」の列にリクエストが入っているか確かめてください。';
  }
  const lines = ['新FMT のリクエスト ' + result.total + '件のうち ' + result.selected + '件を選んでいます（まだ選んでいないもの ' +
    result.remaining + '件）。'];
  if (result.unselectable) {
    lines.push('新FMT でサブインダストリー・得意先・案件名のどれかが空のため、選べないリクエストが ' + result.unselectable + '件あります' +
      '（左の列から順に選ぶため。新FMT で入力すると候補に出ます）。');
  }
  if (!result.remaining && result.untouched) {
    lines.push('サービス案・判断・メモが空の行が ' + result.untouched + '件あります。前の版で自動で取り込んだ行なら、' +
      'エディタで removeUntouchedRequests() を実行すると、削除して候補に戻せます。');
  }
  return lines.join('\n');
}

/* ---------------- シートの用意 ---------------- */

/** {sheet, created, added} */
function svcEnsureSheet_(ss) {
  return taskEnsureSheet_(ss, SVC_OPTIONS.requestSheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
}

/**
 * 判断のプルダウン・追加日の書式を付け、前の版の保護を外す。サービス案の列は自由に入力するので、プルダウンを付けない（前の版で付けたものは外す）。
 * タスク管理シートがあれば、その「サービス案」の列（リクエストに付けたサービス案から選ぶ）にもプルダウンを付ける。
 */
function svcApplyRules_(ss, sheet) {
  const warnings = [];
  const first = SVC_OPTIONS.headerRow + 1;
  const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
  if (rows > 0) {
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    warnings.push.apply(warnings, taskApplyColumnRules_(sheet, SVC_OPTIONS.requestColumns, cols, first, rows));
    if (cols.service) sheet.getRange(first, cols.service, rows, 1).clearDataValidations();
    if (cols.addedAt) sheet.getRange(first, cols.addedAt, rows, 1).setNumberFormat('yyyy/MM/dd');
    svcRemoveLegacyProtections_(sheet);
  }

  let taskAdded = [];
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  if (task) {
    taskAdded = taskEnsureHeader_(task, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow);
    const taskRows = task.getMaxRows() - TASK_OPTIONS.headerRow;
    if (taskRows > 0) {
      taskApplyColumnRules_(task, TASK_OPTIONS.columns.filter(c => c.type === 'service'),
        taskColumnMap_(task, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow), TASK_OPTIONS.headerRow + 1, taskRows);
    }
  }
  return {warnings, taskAdded};
}

/** 前の版で付けた保護（A〜D を手で変えると警告が出る）を外す。今は A〜D をプルダウンで選ぶため。 */
function svcRemoveLegacyProtections_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    let description = '';
    try { description = p.getDescription(); } catch (_) { return; }
    if (SVC_OPTIONS.legacyProtectDescriptions.indexOf(description) >= 0) p.remove();
  });
}

/**
 * タスク管理の「サービス案」の列の入力規則。リクエスト シートの「サービス案」の列を範囲で参照し、
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
    .setHelpText('「' + SVC_OPTIONS.requestSheet + '」シートでリクエストに付けたサービス案から選んでください。')
    .build();
}

/* ---------------- 新FMT のリクエスト ---------------- */

/** リクエストを見分けるキー（得意先・案件名・リクエスト。全角・半角と空白の違いは無視）。 */
function svcRequestKey_(customer, project, request) {
  return [customer, project, request].map(credNormalize_).join('\u0001');
}

/** 読むシートの見出しの行（DIFF_RULES の headerRow。DIFF_RULES に無いシートは SVC_OPTIONS.sourceHeaderRow）。 */
function svcSourceHeaderRow_(name) {
  return DIFF_RULES[name] ? DIFF_RULES[name].headerRow : SVC_OPTIONS.sourceHeaderRow;
}

/**
 * 新FMT（SVC_OPTIONS.sourceSheets）でリクエストが入っている行を、行の順に返す（同じキーは1つ）。
 * [{values: [サブインダストリー, 得意先, 案件名, リクエスト], keys: 見比べ用（credNormalize_）, key}]。見出しが足りないときは止める。
 */
function svcSourceRecords_(ss) {
  const records = [];
  const seen = new Set();
  SVC_OPTIONS.sourceSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('リクエストを読むシート「' + name + '」が見つかりません。');
    const headerRow = svcSourceHeaderRow_(name);
    const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    const cols = SVC_OPTIONS.sourceHeaders.map(h => taskFindColumn_(headers, h, name + ' の ' + headerRow + '行目', true));
    const count = sheet.getLastRow() - headerRow;
    if (count < 1) return;

    const left = Math.min.apply(null, cols);
    const right = Math.max.apply(null, cols);
    sheet.getRange(headerRow + 1, left, count, right - left + 1).getDisplayValues().forEach(line => {
      const values = cols.map(col => String(line[col - left] == null ? '' : line[col - left]).trim());
      if (!values[3]) return;
      const key = svcRequestKey_(values[1], values[2], values[3]);
      if (seen.has(key)) return;
      seen.add(key);
      records.push({values, keys: values.map(credNormalize_), key});
    });
  });
  return records;
}

/**
 * 選んだ行を新FMT に合わせる（行は足さない）。
 * - renamed（1つのセルで得意先・案件名・リクエストを書き換えた）があれば、古いキーの行をその場で新しい値に直す
 *   （新しいキーの行がまだ無いときだけ）。リクエストの書き換えで、サービス案・判断を付けたあとなら注を付ける
 * - サブインダストリーを新FMT に合わせる
 * - 新FMT に無い行には注を付ける（新FMT に戻ったら外す。利用者が書いた注は変えない）
 * {updated, missing}
 */
function svcSync_(ss, sheet, renamed, records) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const fields = ['subIndustry', 'customer', 'project', 'request'];
  const source = new Map(records.map(r => [r.key, r.values]));
  const count = sheet.getLastRow() - headerRow;
  let updated = 0;
  let missing = 0;
  if (count < 1) return {updated, missing};

  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd');
  const read = key => (cols[key]
    ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : new Array(count).fill(''));
  const current = fields.map(read);
  const services = read('service');
  const decisions = read('decision');
  const keyOf = i => (current[3][i] ? svcRequestKey_(current[1][i], current[2][i], current[3][i]) : '');
  const keys = current[3].map((_, i) => keyOf(i));
  const changedRows = new Set();

  // 1つのセルの書き換え：古いキーの行を新しい値に直す
  if (renamed && keys.indexOf(renamed.newKey) < 0) {
    const i = keys.indexOf(renamed.oldKey);
    if (i >= 0) {
      sheet.getRange(headerRow + 1 + i, cols[fields[renamed.level]]).setValue(credText_(renamed.value));
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
    const missingNote = notes[i].indexOf(SVC_OPTIONS.missingNote) === 0;
    const ours = missingNote || notes[i].indexOf(SVC_OPTIONS.changedNote) === 0;   // 利用者が書いた注は変えない
    if (source.has(key)) {
      if (changedRows.has(i) && (!notes[i] || ours)) {
        notesChanged = true;
        return SVC_OPTIONS.changedNote + '（' + today + '）。サービス案・判断を見直してください。';
      }
      if (!missingNote) return notes[i];
      notesChanged = true;
      return '';
    }
    missing++;
    if (missingNote || (notes[i] && !ours)) return notes[i];   // 前回から見つからないまま、または利用者が書いた注
    notesChanged = true;
    return SVC_OPTIONS.missingNote + '（' + today + ' に気づきました）。' +
      SVC_OPTIONS.sourceSheets.join('・') + ' で書き換えか削除された可能性があります。';
  });
  if (notesChanged) noteRange.setNotes(next.map(n => [n]));
  return {updated, missing};
}

/* ---------------- 候補（連動プルダウン） ---------------- */

/**
 * level 番目の連動列の候補。左の列で選んでいる値（空の列では絞り込まない）に合う値を、出てくる順に重複なく返す。
 * memo を渡すと、同じ絞り込みの結果を使い回す（キーに scope を含める）。
 */
function svcOptionsFor_(records, level, selected, memo, scope) {
  const keys = [];
  for (let i = 0; i < level; i++) keys.push(credNormalize_(selected[i] || ''));
  const memoKey = (scope || '') + '\u0002' + level + '\u0001' + keys.join('\u0001');
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
 * リクエスト シートの候補を作り直す。候補は、新FMT でリクエストが入っていて、ほかの行でまだ選んでいない案件だけ。
 * 連動列は左から順に選ぶ。左の列が空の列のプルダウンは「（先に〜を選んでください）」だけにする。
 * edited（{first, last, from, levels: 編集した連動列}）の行は、from 番目から右の連動列で、左の列が空の値と候補に無い値を空にし
 * （編集した列の値なら知らせる）、左から選んだ値に合う案件で1つに決まる次の列（B・C）を埋める（リクエスト・判断・追加日などは入れない）。
 * 入力済みの行は行ごとに、最後の行より下の空いている行はまとめて、プルダウンを付ける。
 * {total: 新FMT のリクエストの数, selected: 選んだ行の数, remaining: まだ選んでいない（選べる）リクエストの数,
 *  unselectable: 新FMT でサブインダストリー〜案件名のどれかが空のため選べないリクエストの数,
 *  untouched: 手を付けていない行の数, messages: 編集した人に知らせる文, warnings: 管理者に知らせる文}
 */
function svcRefreshRequestRows_(sheet, records, edited) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const cascadeKeys = svcCascadeKeys_();
  const fields = cascadeKeys;
  const first = headerRow + 1;
  const lastRow = Math.max(sheet.getLastRow(), headerRow);
  const count = lastRow - headerRow;
  const read = key => (count > 0 && cols[key]
    ? sheet.getRange(first, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : new Array(Math.max(count, 0)).fill(''));
  const values = {};
  fields.forEach(key => { values[key] = read(key); });
  const decisions = read('decision');
  const messages = [];

  const keyOf = i => (values.request[i] ? svcRequestKey_(values.customer[i], values.project[i], values.request[i]) : '');
  const keys = [];
  const used = new Map();
  for (let i = 0; i < count; i++) {
    keys[i] = keyOf(i);
    if (keys[i]) used.set(keys[i], (used.get(keys[i]) || 0) + 1);
  }
  const isUsed = key => (used.get(key) || 0) > 0;
  // 左から順に選ぶので、新FMT でサブインダストリー〜案件名のどれかが空の案件は選べない（候補に出さない）
  const selectable = records.filter(rec => rec.values.slice(0, cascadeKeys.length - 1).every(Boolean));
  const labels = svcCascadeLabels_();
  const isEdited = level => edited.levels.indexOf(level) >= 0;

  // 編集した行：左の列が空の値と候補に無い値を空にし、左から選んだ値で1つに決まる列を埋める
  if (edited) {
    const last = Math.min(edited.last, lastRow);
    for (let row = Math.max(edited.first, first); row <= last; row++) {
      const i = row - first;
      if (keys[i]) used.set(keys[i], used.get(keys[i]) - 1);   // この行の選択はいったん外して考える
      const available = selectable.filter(rec => !isUsed(rec.key));
      const selected = cascadeKeys.map(key => values[key][i]);
      let rejected = false;   // 入力した値を空にした行は、ほかの候補で勝手に埋めない
      const skipped = [];     // 左の列が空なのに入力した値（空にした）
      const emptied = [];     // 消した左の列に続けて空にした列の見出し

      for (let level = Math.max(edited.from, 0); level < cascadeKeys.length; level++) {
        if (!selected[level]) continue;
        const placeholder = svcPlaceholderMessage_(selected[level], records.length > 0);
        if (placeholder) {
          messages.push(placeholder);
          selected[level] = '';
          continue;
        }
        // 左の列が空なら空にする（左の列から順に選ぶ）
        const gap = selected.slice(0, level).findIndex(value => !value);
        if (gap >= 0) {
          if (isEdited(level)) {
            skipped.push(selected[level]);
            rejected = true;
          } else if (isEdited(gap)) {
            emptied.push(labels[level]);
          }
          selected[level] = '';
          continue;
        }
        const options = svcOptionsFor_(available, level, selected);
        if (options.some(option => credNormalize_(option) === credNormalize_(selected[level]))) continue;
        if (isEdited(level)) {
          messages.push(row + '行目の「' + selected[level] + '」は候補に無いため、空にしました（ほかの行で選んでいるか、リクエストがありません）。');
          rejected = true;
        }
        selected[level] = '';
      }
      const gap = selected.findIndex(value => !value);
      if (skipped.length) messages.push(row + '行目の「' + skipped.join('」「') + '」を空にしました（先に' + labels[gap] + 'を選んでください）。');
      if (emptied.length) messages.push(row + '行目の' + labels[gap] + 'が空になったため、' + emptied.join('・') + 'も空にしました。');

      // 左から選んだ値に合う案件で、どれも同じ値になる次の列を埋める（例：サブインダストリーを選んで得意先が1つなら得意先）。
      // リクエスト（D）は、案件が1つに決まっても入れない（人がプルダウンで選ぶ）。人が消した列も埋め直さない
      if (!rejected && selected[0]) {
        const matches = available.filter(rec => selected.every((value, level) => !value || rec.keys[level] === credNormalize_(value)));
        for (let level = 1; level < cascadeKeys.length && matches.length; level++) {
          if (selected[level]) continue;
          if (cascadeKeys[level] === 'request' || isEdited(level)) break;
          const same = matches[0].keys[level];
          if (!matches.every(rec => rec.keys[level] === same)) break;
          selected[level] = matches[0].values[level];
        }
      }
      const next = selected;

      fields.forEach((key, n) => {
        if (next[n] === values[key][i]) return;
        const cell = sheet.getRange(row, cols[key]);
        if (next[n]) cell.setValue(credText_(next[n])); else cell.clearContent();
        values[key][i] = next[n];
      });
      keys[i] = keyOf(i);
      if (keys[i]) used.set(keys[i], (used.get(keys[i]) || 0) + 1);
    }
  }

  // プルダウン：候補は、どの行でもまだ選んでいない案件（入力済みの行は、その行で選んでいる案件も入れる）
  const base = selectable.filter(rec => !isUsed(rec.key));
  const memo = new Map();
  const ruleMemo = new Map();
  const tooMany = new Set();
  const warnings = [];
  const ruleFor = (options, level) => {
    if (options.length > SVC_OPTIONS.maxListItems) {
      tooMany.add(level);
      return null;
    }
    const empty = records.length ? SVC_OPTIONS.allSelected : SVC_OPTIONS.noRequests;
    const list = options.length ? options : [empty];
    const key = level + '\u0001' + list.join('\u0001');
    if (!ruleMemo.has(key)) {
      const label = SVC_OPTIONS.requestColumns.find(c => c.key === cascadeKeys[level]).label;
      ruleMemo.set(key, SpreadsheetApp.newDataValidation()
        .requireValueInList(list, true)
        .setAllowInvalid(false)
        .setHelpText(options.length
          ? label + 'は一覧から選んでください（リクエストがあり、まだ選んでいない案件だけが出ます）。'
          : records.length ? 'リクエストのある案件は、すべて選んでいます。' : '新FMT にリクエストがありません。')
        .build());
    }
    return ruleMemo.get(key);
  };
  // 左の列（gap 番目）が空の列のプルダウン：案内の文字だけにして、左の列から順に選んでもらう
  const waitRule = gap => {
    const key = 'wait\u0001' + gap;
    if (!ruleMemo.has(key)) {
      ruleMemo.set(key, SpreadsheetApp.newDataValidation()
        .requireValueInList([svcChooseFirstText_(gap)], true)
        .setAllowInvalid(false)
        .setHelpText('先に' + labels[gap] + 'を選んでください（左の列から順に選びます）。')
        .build());
    }
    return ruleMemo.get(key);
  };

  const rules = cascadeKeys.map(() => []);
  for (let i = 0; i < count; i++) {
    const own = keys[i] && used.get(keys[i]) === 1 ? keys[i] : '';
    const available = own ? selectable.filter(rec => !isUsed(rec.key) || rec.key === own) : base;
    const selected = cascadeKeys.map(key => values[key][i]);
    const gap = selected.findIndex(value => !value);
    cascadeKeys.forEach((key, level) => {
      // 左に空の列があれば案内の文字だけ（選べる案件が残っていないときは「（すべて選択済み）」）
      if (gap >= 0 && level > gap) {
        rules[level].push([available.length ? waitRule(gap) : ruleFor([], level)]);
        return;
      }
      let options = svcOptionsFor_(available, level, selected, memo, own);
      // 全角・半角と空白の違いは同じ値とみなす。入っている値の書き方が候補と違えば、その行の候補を入っている書き方にする
      const target = credNormalize_(selected[level]);
      const index = target ? options.findIndex(option => credNormalize_(option) === target) : -1;
      if (index >= 0 && options[index] !== selected[level]) {
        options = options.slice();
        options[index] = selected[level];
      }
      rules[level].push([ruleFor(options, level)]);
    });
  }
  // 編集した行のプルダウンを先に反映する（全行の付け直しを待たずに、続けて選べるように）
  if (edited && count > 0) {
    const from = Math.max(edited.first, first) - first;
    const to = Math.min(edited.last, lastRow) - first;
    if (from <= to) {
      cascadeKeys.forEach((key, level) => {
        sheet.getRange(first + from, cols[key], to - from + 1, 1).setDataValidations(rules[level].slice(from, to + 1));
      });
      SpreadsheetApp.flush();
    }
  }
  if (count > 0) cascadeKeys.forEach((key, level) => sheet.getRange(first, cols[key], count, 1).setDataValidations(rules[level]));

  const rest = sheet.getMaxRows() - lastRow;
  if (rest > 0) {
    cascadeKeys.forEach((key, level) => {
      const rule = level === 0 ? ruleFor(svcOptionsFor_(base, 0, [], memo, ''), 0) : base.length ? waitRule(0) : ruleFor([], level);
      sheet.getRange(lastRow + 1, cols[key], rest, 1).setDataValidation(rule);
    });
  }

  tooMany.forEach(level => {
    const label = SVC_OPTIONS.requestColumns.find(c => c.key === cascadeKeys[level]).label;
    warnings.push('「' + label + '」の候補が ' + SVC_OPTIONS.maxListItems + ' 件を超えるため、プルダウンを付けられない行があります。' +
      '左の列を先に選ぶと絞り込まれます。');
  });
  const services = read('service');
  const notes = read('note');
  const untouched = keys.filter((key, i) => key && !services[i] && !notes[i] &&
    (!decisions[i] || decisions[i] === SVC_OPTIONS.defaultDecision)).length;
  const unselectable = records.filter(rec => !isUsed(rec.key)).length - base.length;
  return {total: records.length, selected: keys.filter(Boolean).length, remaining: base.length, unselectable, untouched,
    messages: Array.from(new Set(messages)), warnings};
}

/* ---------------- 編集したとき（onEdit から） ---------------- */

/**
 * 新FMT を編集したとき：サブインダストリー・得意先・案件名・リクエストの列を変えたら、選んだ行を合わせ、候補を作り直す。
 * リクエスト シートが無い・見出しが足りないときは何もしない（setupRequestSheet() を実行すると知らせる）。
 */
function svcHandleSourceEdit_(sheet, range, e) {
  const ss = sheet.getParent();
  const request = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!request) return;
  const headerRow = svcSourceHeaderRow_(sheet.getName());
  if (range.getLastRow() <= headerRow) return;

  const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  const cols = SVC_OPTIONS.sourceHeaders.map(h => (taskHeaderCount_(headers, h) === 1 ? taskFindColumn_(headers, h, '', true) : 0));
  if (cols.some(col => !col)) return;
  if (!cols.some(col => col >= range.getColumn() && col <= range.getLastColumn())) return;

  // 1つのセルで得意先・案件名・リクエストを書き換えたときは、リクエスト シートの同じ行を直す
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
      const records = svcSourceRecords_(ss);
      svcSync_(ss, request, renamed, records);
      svcRefreshRequestRows_(request, records, null);
    });
  } catch (error) {
    throw new Error('リクエスト シートを新FMT に合わせられませんでした（' + error.message + '）。' +
      'あとで setupRequestSheet() を実行するか、もう一度編集してください。');
  }
}

/**
 * リクエスト シートを編集したとき：A〜D（サブインダストリー・得意先・案件名・リクエスト）を変えたら、
 * その行を整え（左の列が空の値・候補に無い値を空にし、1つに決まる次の列（B・C）を埋める）、全行の候補を作り直す。
 * リクエスト（D）とその右の列には値を入れない。
 */
function svcHandleRequestEdit_(sheet, range) {
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
  const edited = col => col > 0 && col >= range.getColumn() && col <= range.getLastColumn();
  const levels = svcCascadeKeys_().map((key, level) => (edited(cols[key]) ? level : -1)).filter(level => level >= 0);
  const level = levels.length ? levels[0] : -1;
  if (level < 0) return;

  const ss = sheet.getParent();
  const first = Math.max(range.getRow(), SVC_OPTIONS.headerRow + 1);
  const last = range.getLastRow();
  taskWithLock_(() => {
    const result = svcRefreshRequestRows_(sheet, svcSourceRecords_(ss), {first, last, from: level, levels});
    if (result.messages.length) ss.toast(result.messages.join('\n'), 'サービス管理', 10);
  }, SVC_OPTIONS.editLockWaitMs);
}
