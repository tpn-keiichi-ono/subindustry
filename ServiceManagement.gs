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
 * - 候補は非表示の「__REQUEST_LISTS」シートの数式で作る（左の列を選ぶとすぐに変わる）。リクエスト シートの onEdit は、
 *   編集した行を整える（候補に無い値を空にする・1つに決まる列を埋める）だけで、プルダウンは付け直さない
 * - サービスは最初から決まっているものではなく、リクエストをもとに考える。リクエストの「サービス案」に入力した名前がサービス案になる
 *   （同じ名前を付けたリクエストが1つのサービス案にまとまる）
 * - 選んだ行は、新FMT の変更に合わせる（単純トリガーの onEdit）：サブインダストリーは同じ案件の値にそろえ、
 *   1つのセルで得意先・案件名・リクエストを書き換えたときは同じ行の値を直す。新FMT から消えたリクエストには注を付ける。
 *   新FMT にリクエストが増えたり減ったりしたら、__REQUEST_LISTS の案件も書き直す
 * - 管理者が setupRequestSheet() をエディタから実行して、シート・見出し・プルダウンを用意する（候補をまとめて作り直すときも）。
 *   前の版で全件を自動で取り込んでいた行のうち、手を付けていない行は removeUntouchedRequests() で削除して候補に戻せる
 * - サービス案ごとの検討の進捗は、タスク管理シート（TaskManagement.gs）でタスクとして管理する
 */

const SVC_OPTIONS = {
  requestSheet: 'リクエスト',
  headerRow: 1,
  // リクエストを読むシートは、test.gs の DIFF_RULES で requests: true を付けたシート（見出しの行も DIFF_RULES の headerRow）。
  // 新FMT2 も読むときは、DIFF_RULES の新FMT2 に requests: true を足す（svcSourceSheets_）
  // 読む列の見出し（リクエスト シートの A〜D の順）。リクエストは Z列。1つのセルに1件
  sourceHeaders: ['サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'],
  // cascade：連動プルダウンの順（0 から）。aliases は前の版の見出し
  requestColumns: [
    {key: 'subIndustry', label: 'サブインダストリー', width: 160, required: true, cascade: 0},
    {key: 'customer', label: '得意先', width: 200, required: true, cascade: 1},
    {key: 'project', label: '案件名', width: 240, required: true, cascade: 2},
    {key: 'request', label: 'リクエスト', width: 360, required: true, cascade: 3},
    {key: 'decision', label: '判断', width: 130, options: ['未判断', 'サービス化検討', '棄却']},
    // insertAfter：前からあるシートに足すときは、その見出しの列のすぐ右に差し込む（無ければ右端）
    {key: 'feedback', label: 'サービス部門からのフィードバック', width: 300, insertAfter: '判断'},   // 自由に入力する
    {key: 'service', label: 'サービス案', width: 220, aliases: ['サービス']},   // まとめる先のサービス案の名前（自由に入力する）
    {key: 'owner', label: '担当者', width: 120, insertAfter: 'サービス案'}       // 自由に入力する
  ],
  // 使わなくなった列（setupRequestSheet() で、確認してから削除する。取り込み日は追加日の前の版の見出し）
  removedColumns: ['メモ', '追加日', '取り込み日'],
  allSelected: '（すべて選択済み）',      // 候補をすべて選んだときにプルダウンに出す文字
  noRequests: '（新FMT にリクエストがありません）',   // 新FMT にリクエストが1件も無いときにプルダウンに出す文字
  chooseLeftFirst: '（先に{label}を選んでください）',  // 左の列が空のときに、右の列のプルダウンに出す文字（{label} は空の列の見出し）
  maxListItems: 500,                     // プルダウンの候補の上限（超えたときは知らせる）
  listSheet: '__REQUEST_LISTS',          // 候補を数式で作る非表示のシート（リクエスト シートの A〜D のプルダウンが参照する）
  listMargin: 5,                         // 候補の欄に足しておく余白（新FMT で候補が増えても、作り直さずに済むように）
  defaultDecision: '未判断',             // まだ判断していないことを表す判断（空と同じに扱う）
  missingNote: '取り込み元に見つかりません',   // 新FMT から消えた・書き換えられたリクエストに付ける注の先頭
  changedNote: '新FMT でリクエストが書き換えられました',   // サービス案・判断を付けたあとで書き換えられたときの注（セルの注。列ではない）
  // 前の版で付けた保護（外す）
  legacyProtectDescriptions: ['リクエスト：新FMT から自動で転記する列', 'リクエスト：案件を選ぶと自動で入る列'],
  lockWaitMs: 30000,                      // エディタから実行する処理のロックの待ち時間
  // リクエスト シートの編集（onEdit）でロックを待つ時間。続けて選ぶと前の選択の処理を待つので長めにする（単純トリガーは30秒で止まる）
  editLockWaitMs: 20000
};

/** リクエストを読むシートの名前（test.gs の DIFF_RULES で requests: true を付けたシート。DIFF_RULES の順）。 */
function svcSourceSheets_() {
  return Object.keys(DIFF_RULES).filter(name => DIFF_RULES[name] && DIFF_RULES[name].requests);
}

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
  else if (svcSourceSheets_().indexOf(name) >= 0) handler = svcHandleSourceEdit_;
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
 * リクエスト シートを用意する（無ければ作る）。見出し・プルダウンを付け、選んだ行を新FMT に合わせてから、
 * 候補のシート（__REQUEST_LISTS）と全行のプルダウンを作り直す。行を足した・並べ替えたあとも実行してよい。
 * 使わなくなった列（SVC_OPTIONS.removedColumns。メモ・追加日）があれば、確認してから削除する。
 * タスク管理シートがあれば「サービス案」の列（無ければ右端に足す）にもプルダウンを付ける。何度実行してもよい。
 */
function setupRequestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 使わなくなった列を消すかは、ロックを取る前に聞く（答えるまでほかの処理を待たせないため）
  const existing = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  const oldColumns = existing ? taskFoundLabels_(existing, SVC_OPTIONS.removedColumns, SVC_OPTIONS.headerRow) : [];
  let removeOld = false;
  if (oldColumns.length) {
    const ui = SpreadsheetApp.getUi();
    removeOld = ui.alert('サービス管理',
      '「' + SVC_OPTIONS.requestSheet + '」シートの「' + oldColumns.join('」「') + '」の列は使わなくなりました。\n' +
      'これらの列を削除します。ほかの列の値はそのまま残ります。削除する列に入っている値は消えます。\n\n削除しますか？',
      ui.ButtonSet.YES_NO) === ui.Button.YES;
  }

  const result = taskWithLock_(() => {
    const removed = removeOld && existing ? taskDeleteColumns_(existing, SVC_OPTIONS.removedColumns, SVC_OPTIONS.headerRow) : [];
    const ensured = Object.assign(svcEnsureSheet_(ss), {removed});
    const records = svcSourceRecords_(ss);
    const synced = svcSync_(ss, ensured.sheet, null, records);
    const rules = svcApplyRules_(ss, ensured.sheet);
    const lists = svcBuildLists_(ss, ensured.sheet, records, true);
    return Object.assign(ensured, synced, rules, svcCountRequests_(ensured.sheet, records),
      {warnings: rules.warnings.concat(lists.warnings)});
  }, SVC_OPTIONS.lockWaitMs);
  result.warnings = result.warnings.concat(svcCheckLists_(ss));   // 数式の計算を待つため、ロックの外で読む

  const name = SVC_OPTIONS.requestSheet;
  const lines = [result.created ? '「' + name + '」シートを作りました。' : '「' + name + '」シートのプルダウンを作り直しました。'];
  lines.push(svcCountMessage_(result));
  if (result.missing) lines.push('新FMT に見つからないリクエストが ' + result.missing + '件あります（「リクエスト」のセルに注を付けました）。');
  if (result.removed.length) lines.push('「' + result.removed.join('」「') + '」の列を削除しました。');
  else if (oldColumns.length) lines.push('「' + oldColumns.join('」「') + '」の列は残しました（使いません。不要なら削除してください）。');
  if (result.added.length) lines.push('「' + result.added.join('」「') + '」の列を足しました。');
  if (result.taskAdded.length) lines.push(TASK_OPTIONS.sheet + ' の右端に「' + result.taskAdded.join('」「') + '」の列を足しました。');
  ss.toast(lines.concat(result.warnings).join('\n'), 'サービス管理', 10);
}

/**
 * 前の版で全件を自動で取り込んだまま手を付けていない行（判断が未判断か空で、フィードバック・サービス案・担当者が空）を削除し、
 * その案件をまた候補に戻す。削除する前に件数を見せて確かめる。判断・フィードバック・サービス案・担当者のどれかが入っている行は残す。
 */
function removeUntouchedRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) throw new Error('「' + SVC_OPTIONS.requestSheet + '」シートがありません。setupRequestSheet() を実行してください。');
  const ui = SpreadsheetApp.getUi();
  const count = svcUntouchedRows_(sheet).length;
  if (!count) {
    ui.alert('サービス管理', '削除できる行はありません（どの行も、判断・フィードバック・サービス案・担当者のどれかが入っています）。', ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert('サービス管理',
    '判断が「' + SVC_OPTIONS.defaultDecision + '」か空で、フィードバック・サービス案・担当者が空の行が ' + count + '件あります。\n' +
    'これらの行を削除して、その案件をまた候補に戻しますか？（判断・フィードバック・サービス案・担当者のどれかが入っている行は残します）',
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
    const records = svcSourceRecords_(ss);
    svcBuildLists_(ss, sheet, records, true);   // 行を消したので、プルダウンの参照を作り直す
    return Object.assign({removed: rows.length}, svcCountRequests_(sheet, records));
  }, SVC_OPTIONS.lockWaitMs);
  ss.toast(result.removed + '行を削除しました。' + svcCountMessage_(result), 'サービス管理', 10);
}

/** 手を付けていない行（リクエストがあり、判断が未判断か空で、フィードバック・サービス案・担当者が空）の行番号（上から）。 */
function svcUntouchedRows_(sheet) {
  const headerRow = SVC_OPTIONS.headerRow;
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return [];
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const read = key => (cols[key]
    ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : new Array(count).fill(''));
  const decisions = read('decision');
  const filled = ['feedback', 'service', 'owner'].map(read);
  const rows = [];
  read('request').forEach((request, i) => {
    if (!svcUntouched_(request, decisions[i], filled.map(values => values[i]))) return;
    rows.push(headerRow + 1 + i);
  });
  return rows;
}

/** 手を付けていない行か（リクエストがあり、判断が未判断か空で、others（フィードバック・サービス案・担当者）がすべて空）。 */
function svcUntouched_(request, decision, others) {
  return !!request && (!decision || decision === SVC_OPTIONS.defaultDecision) && others.every(value => !value);
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
    lines.push('判断・フィードバック・サービス案・担当者が空の行が ' + result.untouched + '件あります。前の版で自動で取り込んだ行なら、' +
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
 * 判断のプルダウンを付け、前の版の保護を外す。フィードバック・サービス案・担当者の列は自由に入力するので、
 * プルダウンを付けない（前の版で付けたもの・列を差し込んだときに左の列から引き継いだものは外す）。
 * タスク管理シートがあれば、その「サービス案」の列（リクエストに付けたサービス案から選ぶ）にもプルダウンを付ける。
 */
function svcApplyRules_(ss, sheet) {
  const warnings = [];
  const first = SVC_OPTIONS.headerRow + 1;
  const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
  if (rows > 0) {
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    warnings.push.apply(warnings, taskApplyColumnRules_(sheet, SVC_OPTIONS.requestColumns, cols, first, rows));
    SVC_OPTIONS.requestColumns.filter(c => c.cascade == null && !c.options && cols[c.key]).forEach(c => {
      sheet.getRange(first, cols[c.key], rows, 1).clearDataValidations();
    });
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

/** 読むシートの見出しの行（DIFF_RULES の headerRow）。 */
function svcSourceHeaderRow_(name) {
  return DIFF_RULES[name].headerRow;
}

/**
 * 新FMT（読むシート。svcSourceSheets_）でリクエストが入っている行を、シート・行の順に返す（同じキーは1つ）。
 * [{values: [サブインダストリー, 得意先, 案件名, リクエスト], keys: 見比べ用（credNormalize_）, key}]。見出しが足りないときは止める。
 */
function svcSourceRecords_(ss) {
  const records = [];
  const seen = new Set();
  const names = svcSourceSheets_();
  if (!names.length) {
    throw new Error('リクエストを読むシートがありません。test.gs の DIFF_RULES で、読むシート（新FMT など）に requests: true を付けてください。');
  }
  names.forEach(name => {
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
    if (!values || credNormalize_(values[0]) === credNormalize_(current[0][i])) return;
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
      svcSourceSheets_().join('・') + ' で書き換えか削除された可能性があります。';
  });
  if (notesChanged) noteRange.setNotes(next.map(n => [n]));
  return {updated, missing};
}

/* ---------------- 候補（連動プルダウン） ---------------- */

/*
 * 候補はスクリプトではなく、非表示の「__REQUEST_LISTS」シート（SVC_OPTIONS.listSheet）の数式で作る。
 * 選ぶたびにスクリプトでプルダウンを付け直すと、onEdit が終わるまで（数秒）次の列を選べないため。
 * 数式はシートが計算するので、左の列を選ぶとすぐに右の列の候補が変わる。
 * - A〜D 列：新FMT の案件（候補にできるものだけ。svcListRecords_）。新FMT を編集するたびにスクリプトが書き直す
 * - E 列：その案件を選んだ行の数（SUMPRODUCT。リクエスト シートの得意先・案件名・リクエストと見比べる）
 * - F1：レイアウト（行数・候補の幅・読む列）。行の追加・削除・並べ替えで合わなくなったら作り直す
 * - G 列から右：リクエスト シートの行ごとの候補。列ごとに幅（widths）ぶんの欄を取り、TRANSPOSE で横に並べる。
 *   リクエスト シートの n 行目の A〜D のプルダウンは、このシートの n 行目の欄を範囲で参照する
 */

/** 候補のシートの列（A〜D：案件、E：選んだ行の数、F：レイアウト、G から右：行ごとの候補）。 */
function svcListColumns_() {
  const levels = svcCascadeKeys_().length;
  return {used: levels + 1, meta: levels + 2, start: levels + 3};
}

/**
 * 候補にする案件。左から順に選ぶので、新FMT でサブインダストリー〜案件名のどれかが空の案件は選べない（外す）。
 * 全角・半角と空白だけが違う値は、列ごとに最初に出てきた書き方にそろえる（プルダウンに同じ値が2つ出ないように）。
 */
function svcListRecords_(records) {
  const levels = svcCascadeKeys_().length;
  const spellings = svcCascadeKeys_().map(() => new Map());
  return records.filter(rec => rec.values.slice(0, levels - 1).every(Boolean)).map(rec => {
    const values = rec.values.map((value, level) => {
      if (!spellings[level].has(rec.keys[level])) spellings[level].set(rec.keys[level], value);
      return spellings[level].get(rec.keys[level]);
    });
    return {values, keys: rec.keys, key: rec.key};
  });
}

/** 列ごとの候補の数の最大（左の列で選んだ値の組ごとに数える。大文字・小文字の違いはシートと同じく同じ値とみなす）。 */
function svcListCounts_(list) {
  return svcCascadeKeys_().map((key, level) => {
    const groups = new Map();
    list.forEach(rec => {
      const prefix = rec.values.slice(0, level).map(v => String(v).toLowerCase()).join('\u0001');
      if (!groups.has(prefix)) groups.set(prefix, new Set());
      groups.get(prefix).add(rec.values[level]);
    });
    let max = 0;
    groups.forEach(values => { max = Math.max(max, values.size); });
    return max;
  });
}

/** 候補のシートのレイアウト（F1 の JSON）。無い・読めないときは null。 */
function svcListMeta_(text) {
  try {
    const meta = JSON.parse(String(text || ''));
    return meta && Array.isArray(meta.widths) ? meta : null;
  } catch (_) {
    return null;
  }
}

/**
 * 候補のシートを読む（1回の読み取り）。{meta, list: 候補にする案件}。シートが無ければ null。
 * リクエスト シートを編集するたびに使うので、新FMT は読まない。
 */
function svcReadLists_(ss) {
  const helper = ss.getSheetByName(SVC_OPTIONS.listSheet);
  if (!helper) return null;
  const levels = svcCascadeKeys_().length;
  const cols = svcListColumns_();
  const lines = helper.getRange(1, 1, helper.getMaxRows(), cols.meta).getDisplayValues();
  const list = [];
  for (let i = 1; i < lines.length && lines[i][levels - 1]; i++) {
    const values = lines[i].slice(0, levels).map(v => String(v).trim());
    const keys = values.map(credNormalize_);
    list.push({values, keys, key: svcRequestKey_(values[1], values[2], values[3])});
  }
  return {meta: svcListMeta_(lines[0][cols.meta - 1]), list};
}

/**
 * 候補のシートを用意する。案件（A〜E 列）はいつも書き直す。
 * relayout が true のとき・レイアウトが合わないとき（リクエスト シートの行数・読む列が変わった、候補が欄に入りきらない）は、
 * 行ごとの候補の数式と、リクエスト シートの A〜D のプルダウン（範囲で参照する）も作り直す。{relaid, warnings}
 */
function svcBuildLists_(ss, sheet, records, relayout) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cascadeKeys = svcCascadeKeys_();
  const levels = cascadeKeys.length;
  const labels = svcCascadeLabels_();
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const layout = svcListColumns_();
  const list = svcListRecords_(records);
  const counts = svcListCounts_(list);
  const warnings = [];
  counts.forEach((count, level) => {
    if (count <= SVC_OPTIONS.maxListItems) return;
    warnings.push('「' + labels[level] + '」の候補が ' + SVC_OPTIONS.maxListItems + ' 件を超えるため、プルダウンにすべては出せません（SVC_OPTIONS.maxListItems）。');
  });

  let helper = ss.getSheetByName(SVC_OPTIONS.listSheet);
  if (!helper) {
    helper = ss.insertSheet(SVC_OPTIONS.listSheet);
    helper.hideSheet();
    relayout = true;
  }
  const rows = sheet.getMaxRows();
  const columns = cascadeKeys.map(key => cols[key]).join(',');
  const old = svcListMeta_(helper.getRange(1, layout.meta).getDisplayValue());
  const fit = counts.map(count => Math.min(count, SVC_OPTIONS.maxListItems));
  if (!old || old.rows !== rows || old.columns !== columns || old.widths.length !== levels ||
      fit.some((count, level) => count > old.widths[level])) {
    relayout = true;
  }
  // 欄の幅は、新FMT で候補が少し増えても作り直さずに済むように余白を足す（案内の文字のために少なくとも1）
  const widths = relayout
    ? fit.map(count => Math.min(Math.max(count, 1) + Math.max(SVC_OPTIONS.listMargin, Math.ceil(count * 0.2)), SVC_OPTIONS.maxListItems))
    : old.widths;
  const starts = [];
  widths.reduce((col, width) => { starts.push(col); return col + width; }, layout.start);

  // 大きさ：行はリクエスト シートの行と案件の数の多いほう、列は候補の欄の右端まで
  const needRows = Math.max(rows, list.length + 1);
  const needCols = layout.start + widths.reduce((a, b) => a + b, 0) - 1;
  if (helper.getMaxRows() < needRows) helper.insertRowsAfter(helper.getMaxRows(), needRows - helper.getMaxRows());
  if (helper.getMaxColumns() < needCols) helper.insertColumnsAfter(helper.getMaxColumns(), needCols - helper.getMaxColumns());

  // 案件と、それを選んだ行の数
  const requestRef = "'" + sheet.getName().replace(/'/g, "''") + "'!";
  const column = key => requestRef + '$' + diffColumnLetter_(cols[key]) + '$' + (headerRow + 1) + ':$' + diffColumnLetter_(cols[key]);
  if (relayout) helper.clearContents();
  else helper.getRange(2, 1, helper.getMaxRows() - 1, layout.used).clearContent();
  if (list.length) {
    helper.getRange(2, 1, list.length, levels).setValues(list.map(rec => rec.values.map(credText_)));
    helper.getRange(2, layout.used, list.length, 1).setFormulas(list.map((rec, i) => {
      const row = i + 2;
      return ['=SUMPRODUCT(' + ['customer', 'project', 'request'].map(key => {
        const level = cascadeKeys.indexOf(key);
        return '(' + column(key) + '=$' + diffColumnLetter_(level + 1) + row + ')';
      }).join('*') + ')'];
    }));
  }
  const meta = {rows, columns, widths, sourceRows: list.length};
  helper.getRange(1, 1, 1, layout.meta).setValues([labels.concat(['選んだ行の数', credText_(JSON.stringify(meta))])]);
  if (!relayout) return {relaid: false, warnings};

  // 行ごとの候補の数式と、リクエスト シートのプルダウン
  const first = headerRow + 1;
  const count = rows - headerRow;
  if (count > 0) {
    const requestCells = cascadeKeys.map(key => requestRef + '$' + diffColumnLetter_(cols[key]));
    cascadeKeys.forEach((key, level) => {
      helper.getRange(1, starts[level]).setValue(labels[level] + 'の候補');
      const formulas = [];
      const rules = [];
      const help = labels[level] + 'は一覧から選んでください（' + (level ? '左の列から順に選びます。' : '') +
        'リクエストがあり、まだ選んでいない案件だけが出ます）。';
      for (let row = first; row < first + count; row++) {
        formulas.push([svcListFormula_(level, row, requestCells)]);
        rules.push([SpreadsheetApp.newDataValidation()
          .requireValueInRange(helper.getRange(row, starts[level], 1, widths[level]), true)
          .setAllowInvalid(false)
          .setHelpText(help)
          .build()]);
      }
      helper.getRange(first, starts[level], count, 1).setFormulas(formulas);
      sheet.getRange(first, cols[key], count, 1).setDataValidations(rules);
    });
  }
  return {relaid: true, warnings};
}

/**
 * リクエスト シートの row 行目の、level 番目の連動列の候補を横に並べる数式。
 * 左に空の列があれば「（先に〇〇を選んでください）」、選べる案件が無ければ「（すべて選択済み）」だけにする。
 * 候補は、左の列で選んだ値に合う案件のうち、どの行でも選んでいないもの（その行で入っている値は残す）。
 */
function svcListFormula_(level, row, requestCells) {
  const levels = svcCascadeKeys_().length;
  const quote = text => '"' + String(text).replace(/"/g, '""') + '"';
  const source = k => '$' + diffColumnLetter_(k + 1) + '$2:$' + diffColumnLetter_(k + 1);
  const used = '$' + diffColumnLetter_(svcListColumns_().used) + '$2:$' + diffColumnLetter_(svcListColumns_().used);
  const cell = k => requestCells[k] + row;
  const conditions = [source(levels - 1) + '<>""'];
  for (let k = 0; k < level; k++) conditions.push(source(k) + '=' + cell(k));
  conditions.push('(' + used + '=0)+(' + source(level) + '=' + cell(level) + ')');
  let formula = 'IFERROR(TRANSPOSE(UNIQUE(FILTER(' + source(level) + ',' + conditions.join(',') + '))),' +
    quote(SVC_OPTIONS.allSelected) + ')';
  for (let k = level - 1; k >= 0; k--) formula = 'IF(' + cell(k) + '="",' + quote(svcChooseFirstText_(k)) + ',' + formula + ')';
  if (level === 0) formula = 'IF(COUNTA(' + source(levels - 1) + ')=0,' + quote(SVC_OPTIONS.noRequests) + ',' + formula + ')';
  return '=' + formula;
}

/**
 * 候補の数式がエラーになっていないかを確かめる（setupRequestSheet() の最後に、ロックの外で）。
 * リクエスト シートの最初のデータの行の候補を読み、「#」で始まる値（#NAME? など）があれば知らせる文を返す。
 */
function svcCheckLists_(ss) {
  const helper = ss.getSheetByName(SVC_OPTIONS.listSheet);
  if (!helper) return [];
  const layout = svcListColumns_();
  const meta = svcListMeta_(helper.getRange(1, layout.meta).getDisplayValue());
  const row = SVC_OPTIONS.headerRow + 1;
  if (!meta || meta.rows < row) return [];
  const width = meta.widths.reduce((a, b) => a + b, 0);
  const values = helper.getRange(row, layout.start, 1, width).getDisplayValues()[0];
  const errors = Array.from(new Set(values.filter(v => /^#/.test(String(v)))));
  return errors.length
    ? ['候補を作る数式でエラーが出ています（' + errors.join('・') + '）。「' + SVC_OPTIONS.listSheet + '」シートを表示して確かめてください。']
    : [];
}

/**
 * level 番目の連動列の候補。左の列で選んでいる値（空の列では絞り込まない）に合う値を、出てくる順に重複なく返す。
 * 編集した行を整えるとき（svcFixEditedRows_）に使う。
 */
function svcOptionsFor_(records, level, selected) {
  const keys = [];
  for (let i = 0; i < level; i++) keys.push(credNormalize_(selected[i] || ''));
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
  return out;
}

/**
 * 編集した行（edited = {first, last, from, levels: 編集した連動列}）を整える。list は候補にする案件（svcListRecords_）、
 * cols はリクエスト シートの列（taskColumnMap_）。
 * from 番目から右の連動列で、左の列が空の値と候補に無い値（ほかの行で選んでいる案件など）を空にし（編集した列の値なら知らせる）、
 * 左から選んだ値に合う案件で1つに決まる次の列（得意先・案件名）を埋める。リクエスト・判断などは入れない。
 * プルダウンは数式が作るので、ここでは付け直さない。戻り値は編集した人に知らせる文。
 */
function svcFixEditedRows_(sheet, list, edited, cols) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cascadeKeys = svcCascadeKeys_();
  const labels = svcCascadeLabels_();
  const first = headerRow + 1;
  const lastRow = Math.max(sheet.getLastRow(), headerRow);
  const count = lastRow - headerRow;
  const messages = [];
  if (count < 1) return messages;

  // 連動列をまとめて1回で読む
  const left = Math.min.apply(null, cascadeKeys.map(key => cols[key]));
  const right = Math.max.apply(null, cascadeKeys.map(key => cols[key]));
  const lines = sheet.getRange(first, left, count, right - left + 1).getDisplayValues();
  const values = {};
  cascadeKeys.forEach(key => { values[key] = lines.map(line => String(line[cols[key] - left]).trim()); });

  const keyOf = i => (values.request[i] ? svcRequestKey_(values.customer[i], values.project[i], values.request[i]) : '');
  const keys = [];
  const used = new Map();
  for (let i = 0; i < count; i++) {
    keys[i] = keyOf(i);
    if (keys[i]) used.set(keys[i], (used.get(keys[i]) || 0) + 1);
  }
  const isUsed = key => (used.get(key) || 0) > 0;
  const isEdited = level => edited.levels.indexOf(level) >= 0;

  const last = Math.min(edited.last, lastRow);
  for (let row = Math.max(edited.first, first); row <= last; row++) {
    const i = row - first;
    if (keys[i]) used.set(keys[i], used.get(keys[i]) - 1);   // この行の選択はいったん外して考える
    const available = list.filter(rec => !isUsed(rec.key));
    const selected = cascadeKeys.map(key => values[key][i]);
    let rejected = false;   // 入力した値を空にした行は、ほかの候補で勝手に埋めない
    const skipped = [];     // 左の列が空なのに入力した値（空にした）
    const emptied = [];     // 消した左の列に続けて空にした列の見出し

    for (let level = Math.max(edited.from, 0); level < cascadeKeys.length; level++) {
      if (!selected[level]) continue;
      const placeholder = svcPlaceholderMessage_(selected[level], list.length > 0);
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

    cascadeKeys.forEach((key, n) => {
      if (selected[n] === values[key][i]) return;
      const target = sheet.getRange(row, cols[key]);
      if (selected[n]) target.setValue(credText_(selected[n])); else target.clearContent();
      values[key][i] = selected[n];
    });
    keys[i] = keyOf(i);
    if (keys[i]) used.set(keys[i], (used.get(keys[i]) || 0) + 1);
  }
  return Array.from(new Set(messages));
}

/** 選んだ件数を数える。{total, selected, remaining, unselectable, untouched}（svcCountMessage_ で知らせる）。 */
function svcCountRequests_(sheet, records) {
  const headerRow = SVC_OPTIONS.headerRow;
  const count = sheet.getLastRow() - headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const read = key => (count > 0 && cols[key]
    ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : []);
  const customers = read('customer');
  const projects = read('project');
  const decisions = read('decision');
  const filled = ['feedback', 'service', 'owner'].map(read);
  const used = new Set();
  let selected = 0;
  let untouched = 0;
  read('request').forEach((request, i) => {
    if (!request) return;
    selected++;
    used.add(svcRequestKey_(customers[i], projects[i], request));
    if (svcUntouched_(request, decisions[i], filled.map(values => values[i]))) untouched++;
  });
  const list = svcListRecords_(records);
  const listed = new Set(list.map(rec => rec.key));
  return {
    total: records.length,
    selected,
    remaining: list.filter(rec => !used.has(rec.key)).length,
    unselectable: records.filter(rec => !listed.has(rec.key) && !used.has(rec.key)).length,
    untouched
  };
}

/**
 * 編集した行のプルダウンが、候補のシートの同じ行を参照しているか。
 * 行の追加・削除・並べ替えでずれたとき（またはリクエスト シートの行数・読む列が変わったとき）は false。
 */
function svcListsInPlace_(sheet, meta, cols, first, last) {
  const columns = svcCascadeKeys_().map(key => cols[key]).join(',');
  if (!meta || meta.rows !== sheet.getMaxRows() || meta.columns !== columns) return false;
  const rules = sheet.getRange(first, cols[svcCascadeKeys_()[0]], last - first + 1, 1).getDataValidations();
  return rules.every((line, i) => {
    const rule = line[0];
    if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return false;
    const range = rule.getCriteriaValues()[0];
    return range.getSheet().getName() === SVC_OPTIONS.listSheet && range.getRow() === first + i;
  });
}

/* ---------------- 編集したとき（onEdit から） ---------------- */

/**
 * 新FMT を編集したとき：サブインダストリー・得意先・案件名・リクエストの列を変えたら、選んだ行を合わせ、候補のシートの案件を書き直す。
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
      svcBuildLists_(ss, request, records, false);
    });
  } catch (error) {
    throw new Error('リクエスト シートを新FMT に合わせられませんでした（' + error.message + '）。' +
      'あとで setupRequestSheet() を実行するか、もう一度編集してください。');
  }
}

/**
 * リクエスト シートを編集したとき：A〜D（サブインダストリー・得意先・案件名・リクエスト）を変えたら、その行を整える
 * （左の列が空の値・候補に無い値を空にし、1つに決まる次の列（得意先・案件名）を埋める）。リクエスト（D）とその右の列には値を入れない。
 * 候補は数式が作るので、プルダウンは付け直さない（速く終わるように、新FMT も読まない）。
 * 行の追加・削除・並べ替えでプルダウンの参照がずれていたときだけ、候補のシートとプルダウンを作り直す。
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
    const lists = svcReadLists_(ss);
    const messages = svcFixEditedRows_(sheet, lists ? lists.list : svcListRecords_(svcSourceRecords_(ss)),
      {first, last, from: level, levels}, cols);
    if (messages.length) ss.toast(messages.join('\n'), 'サービス管理', 10);
    if (!lists || !svcListsInPlace_(sheet, lists.meta, cols, first, last)) {
      svcBuildLists_(ss, sheet, lists && lists.meta ? lists.list : svcSourceRecords_(ss), true);
    }
  }, SVC_OPTIONS.editLockWaitMs);
}
