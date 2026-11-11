/**
 * RX Business Meeting の会議サマリ「報告案件版」（MeetingSummary.gs の会議サマリと同じ作り方で、案件のページだけを変えたもの）。
 *
 * - 数字のページ（表紙・サマリー・計画に対する積み上げ・パイプラインの変化・RXシナリオのポートフォリオ）は、会議サマリと同じ。
 * - 案件のページのうち「期待値上位案件の状況・課題」と Appendix（扉と ①〜③）は入れない。
 *   代わりに、シートの「報告フラグ」（MTG_OPTIONS.columns.reportFlag）にチェックを入れた案件を、1件1枚の「報告案件」のスライド
 *   （見出しは MTG_OPTIONS.report.title「RXシナリオ進捗における報告案件」と番号）にする。
 * - 報告案件のあとに「前回から変化のある案件」（会議サマリと同じ表。Appendix が無いので、上から7件までにせず全件を載せる。mtgMovesSlide_ の all）。
 * - 報告案件のスライドは枠の位置を固定（ひな形として同じ形）にし、中の文章の分量で字の大きさを変える
 *   （右の文章の枠は、枠ごとに文の量と枠の大きさで入るいちばん大きい字（MTG_OPTIONS.report の maxSize（見出しと同じ）から minSize まで）。どの枠も文は省略しない）。
 *
 * 作り方（エディタから実行）：createBusinessMtgReportSummaries()（本番）・previewBusinessMtgReportSummaries()（試作）。
 * 前の関数（createBusinessMtgSummaries()・previewBusinessMtgSummaries()）はそのまま使える。本番は、会議日ごとにどちらか一方で作る
 * （時点コピー・会議日の記録・会議のフォルダ・インデックスは同じものを使うため）。ファイル名には MTG_OPTIONS.report.fileSuffix（「_報告案件版」）が付く。
 */

/** 報告案件版の会議サマリを作る（本番。資料を作る画面を開く。会議日を記録し、時点コピーを保存する） */
function createBusinessMtgReportSummaries() {
  mschOpenRunDialog_('report', true);
}

/** 報告案件版の会議サマリを試しに作る（会議日は記録しない。ファイル名に「_試作」が付く） */
function previewBusinessMtgReportSummaries() {
  mschOpenRunDialog_('report', false);
}

/**
 * 各サブインダストリーのシート（DIFF_RULES の追跡シート）の MTG_OPTIONS.report.flagColumn（AG）列に、「報告フラグ」の列を足す（エディタから実行。管理者の作業）。
 * 見出しの行（headerRow）に「報告フラグ」と書き（書式は左の列の見出しと同じ）、その下の行をすべてチェックボックスにする（入力規則だけ。値は書かないので、
 * シートの最後の行は変わらない）。何度実行してもよい：もう「報告フラグ」の列があるシートは、チェックボックスを付け直すだけ（AG 以外の列にあれば、そのまま）。
 * 足さないシート：AG列に何か書いてある（上書きしない）・AG列が追跡範囲（DIFF_RULES の ranges）の中・AF列まで無い（列を増やすと追跡範囲が変わるため）。
 * 追跡範囲の外なので、チェックしても変更の記録（差分追跡）には残らない。列を増やしたシート（AF列までのシート）は、差分追跡が列の数の変化を
 * 「構成が変わった」として止めるので、足したあとにその場で比較基準を取り直す（diffRebaselineSheet_。列を足すだけなので値は変わらない）。
 * 取り直す前に、まだ記録していない変更（取りこぼし回収の印）があるシートは足さない（取り直すと、その変更が記録されなくなるため）
 */
function addBusinessMtgReportFlagColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const o = MTG_OPTIONS;
  const names = mtgTrackedNames_(ss);
  if (!names.length) throw new Error('追跡しているシート（test.gs の DIFF_RULES）が見つかりません。');
  const ui = mtgUiOrNull_();
  const column = o.report.flagColumn;
  if (ui) {
    const answer = ui.alert(CRED_OPTIONS.menuTitle, '「' + names.join('」「') + '」の ' + column + '列に「' + o.columns.reportFlag +
      '」の列（チェックボックス）を足します。\n' + column + '列に何か書いてあるシートには足しません。よろしいですか？', ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('ほかの処理（差分追跡など）が実行中です。少し待ってからもう一度お試しください。');
  let lines;
  try {
    lines = names.map(name => '「' + name + '」：' + mtgReportAddFlagColumn_(ss, ss.getSheetByName(name), DIFF_RULES[name], column));
  } finally {
    lock.releaseLock();
  }
  const message = '報告フラグの列\n' + lines.join('\n') + '\n\n営業リーダーが、会議で報告する案件にチェックを入れます。' +
    '報告案件版の資料は createBusinessMtgReportSummaries()（試作は previewBusinessMtgReportSummaries()）で作ります。';
  console.log(message);
  if (ui) ui.alert(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK);
  return lines;
}

/** 1つのシートに報告フラグの列を足す（addBusinessMtgReportFlagColumn。ドキュメントロックの中で呼ぶ）。結果の文（足した・付け直した・足さなかった理由） */
function mtgReportAddFlagColumn_(ss, sheet, rule, letter) {
  const o = MTG_OPTIONS;
  const label = o.columns.reportFlag;
  const col = mtgColumnNumber_(letter);
  const headerRow = rule.headerRow;
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  const checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const boxes = c => (maxRows > headerRow ? sheet.getRange(headerRow + 1, c, maxRows - headerRow, 1) : null);
  // もう列があれば、チェックボックスを付け直すだけ
  const headers = sheet.getRange(headerRow, 1, 1, maxCols).getDisplayValues()[0];
  const found = mtgColumnIndex_(headers, 'reportFlag');
  if (found >= 0) {
    if (found + 1 !== col) return '「' + label + '」の列は ' + mtgColumnLetter_(found + 1) + '列にあります（そのまま）。';
    const range = boxes(col);
    if (range) range.setDataValidation(checkbox);
    return '「' + label + '」の列は ' + letter + '列にあります（チェックボックスを付け直しました）。';
  }
  // 追跡範囲の中には置かない（チェックのたびに変更の記録が増え、会議サマリの「その他更新」にも出るため）
  const tracked = (rule.ranges || []).filter(a1 => {
    const m = /^\$?([A-Z]*)\$?\d*(?::\$?([A-Z]*)\$?\d*)?$/.exec(String(a1).toUpperCase().trim());
    if (!m) return false;
    const from = m[1] ? mtgColumnNumber_(m[1]) : 1;
    const to = m[2] === undefined ? from : m[2] ? mtgColumnNumber_(m[2]) : Infinity;
    return from <= col && col <= to;
  });
  if (tracked.length) return letter + '列が追跡範囲（' + tracked.join('・') + '）の中のため、足しませんでした（MTG_OPTIONS.report.flagColumn を範囲の外の列にしてください）。';
  // 左の列（AF）まで無いシートは、列を増やすと追跡範囲（D3:AF など）が広がって変更の記録の比べ方が変わるので、足さない
  if (maxCols < col - 1) return mtgColumnLetter_(col - 1) + '列まで無いため、足しませんでした（今は ' + mtgColumnLetter_(maxCols) + '列まで）。';
  let rebuilt = false;
  if (maxCols < col) {
    // 差分追跡は列の数（gridColumns）を覚えていて、変わると次の編集で止まるので、列を足したら比較基準を取り直す。
    // まだ記録していない変更があれば、取り直すと記録されなくなるので足さない（数分ごとの取りこぼし回収のあとで、もう一度実行する）
    const saved = diffSavedState_(sheet);
    const tracking = !!(saved && !saved.paused);
    if (tracking && diffPendingSheetIds_(PropertiesService.getDocumentProperties().getProperties()).indexOf(String(sheet.getSheetId())) >= 0) {
      return 'まだ変更の記録に書いていない編集があるため、足しませんでした（数分ごとの取りこぼし回収のあとで、もう一度実行してください）。';
    }
    sheet.insertColumnsAfter(maxCols, col - maxCols);
    if (tracking) {
      diffRebaselineSheet_(ss, sheet.getName());
      rebuilt = true;
    }
  } else {
    // AG列に何か書いてあれば上書きしない
    const values = sheet.getRange(1, col, maxRows, 1).getDisplayValues();
    const used = values.findIndex(v => String(v[0]).trim() !== '');
    if (used >= 0) {
      return letter + '列に値があるため、足しませんでした（' + letter + (used + 1) + '：「' + mtgWrap_(values[used][0], 200, 9, 1)[0] + '」）。' +
        'MTG_OPTIONS.report.flagColumn を空いている列にしてください。';
    }
  }
  const head = sheet.getRange(headerRow, col);
  if (col > 1) sheet.getRange(headerRow, col - 1).copyTo(head, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  head.setValue(label);
  const range = boxes(col);
  if (range) range.setDataValidation(checkbox).setHorizontalAlignment('center');
  sheet.setColumnWidth(col, o.report.flagWidth);
  return letter + '列に「' + label + '」の列を足しました（' + (headerRow + 1) + '行目から下をチェックボックスに' +
    (rebuilt ? '。列が増えたので、変更履歴の比較基準を取り直しました' : '') + '）。';
}

/** 列の文字（AG）→ 番号（33） */
function mtgColumnNumber_(letters) {
  return String(letters).toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

/** 列の番号（33）→ 文字（AG） */
function mtgColumnLetter_(n) {
  let out = '';
  for (let v = n; v > 0; v = Math.floor((v - 1) / 26)) out = String.fromCharCode(65 + (v - 1) % 26) + out;
  return out;
}

/* ---------------- 前回の会議の宿題事項 ---------------- */

/**
 * 前回の会議の宿題事項（MTG_OPTIONS.homework のシートから読む。読むだけ）。meeting は今の会議日（yyyy/MM/dd）。
 * {date: 前回の会議日|null, rows: [{name: サブインダストリー（シートの A列の順。前の名前は今の名前に）, items: [宿題]}], gap: 読めない理由（シート・見出しが無い）}
 */
function mtgHomeworkOf_(ss, name, meeting) {
  const H = MTG_OPTIONS.homework;
  // 読むシート：homework.sheet、無ければ名前が index.sheet（INDEX_会議資料）で始まるシートのうち、「宿題」の列が読めるもの（シートの並びの順）。
  // 今の INDEX_会議資料（行＝会議日の形）は「サブインダストリー」の見出しの行が無いので読まずに次へ
  // シートの名前の一覧（Sheets API で読んだ入れ物は持っている一覧。SpreadsheetApp ならシートから。API の呼び出しを増やさない）
  const titles = typeof ss.getSheetNames === 'function' ? ss.getSheetNames() : ss.getSheets().map(sh => sh.getName());
  const names = [H.sheet].concat(titles.filter(t => t !== H.sheet && String(t).indexOf(MTG_OPTIONS.index.sheet) === 0));
  let first = null;
  for (const title of names) {
    const sheet = ss.getSheetByName(title);
    if (!sheet) continue;
    const got = mtgHomeworkRead_(sheet, meeting);
    if (!got.gap) return got;
    if (!first) first = got;
  }
  return first || {date: null, rows: [], gap: '「' + H.sheet + '」シート（または名前が「' + MTG_OPTIONS.index.sheet + '」で始まり、「' + H.column + '」の列があるシート）が無い'};
}

/** 宿題のシート（sheet）を読む（mtgHomeworkOf_）。{date, rows, gap}（見出しの行が無ければ gap） */
function mtgHomeworkRead_(sheet, meeting) {
  const H = MTG_OPTIONS.homework;
  const out = {date: null, rows: [], gap: '', sheet: sheet.getName()};
  const rows = sheet.getLastRow();
  const cols = sheet.getLastColumn();
  if (!rows || !cols) return out;
  const range = sheet.getRange(1, 1, rows, cols);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  const label = r => String(shown[r][0] == null ? '' : shown[r][0]).trim();
  const find = text => shown.findIndex((line, r) => label(r) === text);
  const dateRow = find(H.dateLabel);
  const kindRow = find(H.kindLabel);
  const subRow = find(H.nameLabel);
  if (dateRow < 0 || subRow < 0) {
    return Object.assign(out, {gap: '「' + sheet.getName() + '」シートの A列に「' + H.dateLabel + '」「' + H.nameLabel + '」の見出しが無い'});
  }
  // 列ごとの会議日・会議名（結合したセルは左の列だけに値があるので、右へ引き継ぐ）
  const kinds = mtgIndexKinds_(MTG_OPTIONS.index.kind);
  let date = '';
  let kind = '';
  let best = null;
  for (let c = 1; c < cols; c++) {
    const d = mtgIndexDate_(values[dateRow][c]);
    if (d) {
      date = d;
      kind = '';
    }
    const k = kindRow >= 0 ? String(shown[kindRow][c] || '').trim() : '';
    if (k) kind = k;
    const sub = String(shown[subRow][c] || '').trim();
    if (sub !== H.column || !date || date >= meeting) continue;
    if (kindRow >= 0 && kind && kinds.indexOf(kind) < 0) continue;   // RX Service Meeting などの宿題は使わない
    if (!best || date > best.date) best = {date, col: c};
  }
  if (!best) return out;
  out.date = best.date;
  const resolve = diffSheetNameResolver_();
  // 宿題のセルは改行で1件ずつ（先頭の「・」「①」などは取る。資料では ①② と番号を振る）
  const items = text => String(text == null ? '' : text).split(/\r?\n/)
    .map(line => line.replace(/^\s*(?:[・●○■□\-－*＊]|[\u2460-\u2473]|\(?\d+[).．、）])\s*/, '').trim()).filter(Boolean);
  for (let r = subRow + 1; r < rows; r++) {
    if (!label(r)) continue;
    out.rows.push({name: resolve(label(r)), items: items(shown[r][best.col])});
  }
  return out;
}

/**
 * 前回の宿題事項のスライド（報告案件版。表紙のあと・サマリーの前）：この資料のサブインダストリーの宿題だけを、ページの幅いっぱいの1つの表に
 * （「No.」「指摘事項」。件数が少ないので、字は読みやすい大きさ（リード文と同じ 10.5pt）。1枚に入らなければ 8pt まで小さくし、それでも入らなければ
 * 続きのページに分けて全件を出す）。シートが無ければ「データなし」、宿題が無ければ（シートにこのサブインダストリーの行が無いときも）「該当なし」
 */
function mtgHomeworkSlides_(s) {
  const S = MTG_SLIDE;
  const c = MTG_OPTIONS.colors;
  const H = MTG_OPTIONS.homework;
  const h = s.homework || {date: null, rows: [], gap: ''};
  const title = MTG_TITLES.homework;
  const columns = mtgColumns_([48, S.w - 48], 1);
  const labels = ['No.', H.itemHead];
  const when = h.date ? '前回会議（' + mtgShortDate_(h.date) + '）' : '前回会議';
  if (h.gap) {
    return [{title, lead: '前回の会議の宿題事項は' + MTG_OPTIONS.gapText + '（' + h.gap + '）。', elements: mtgEmptyTable_(columns, mtgHeadRow_(columns, labels), 'gap', h.gap)}];
  }
  const own = h.rows.filter(r => r.name === s.name);
  const items = own.reduce((list, r) => list.concat(r.items), []);
  if (!items.length) {
    const none = !h.date ? '前の RX Business Meeting の宿題の列が無い'
      : !own.length ? '「' + (h.sheet || H.sheet) + '」シートに「' + s.name + '」の行が無い' : when + 'の宿題は登録されていない';
    return [{title, lead: none + '。', elements: mtgEmptyTable_(columns, mtgHeadRow_(columns, labels), 'none', none)}];
  }
  const lead = when + 'で出た宿題 ' + items.length + '件。';
  // 指摘事項のセルには、見積もりで折り返した改行を入れず、シートの文をそのまま入れる（スライド・PDF からコピーして貼り付けたときに、文の途中で改行しない。
  // 折り返しは Google スライドに任せる。行の高さは見積もりの行の数で決める（見積もりは本当の幅より早めに折り返すので、高さは足りる））
  const build = size => items.map((text, i) => {
    const row = mtgRow_(columns, [
      {text: String(i + 1), size: size + 1.5, bold: true, color: c.greenDark, fill: c.greenSoft, align: 'CENTER'},
      {text, size, lines: 6, fill: c.white}
    ], Math.round(size * 3));
    row.cells[1] = Object.assign({}, row.cells[1], {text});
    return row;
  });
  // 1枚に入るいちばん大きい字（10.5pt から 0.5pt ずつ 8pt まで）
  for (let size = S.leadSize; size >= 8 - 1e-9; size -= 0.5) {
    const head = mtgHeadRow_(columns, labels);
    const rows = build(size);
    if (S.top + head.h + rows.reduce((n, r) => n + r.h, 0) <= S.bottom) {
      return [{title, lead, elements: [{type: 'table', x: S.x, y: S.top, columns, rows: [head].concat(rows)}]}];
    }
  }
  return mtgTablePages_(mtgHeadRow_(columns, labels), build(8)).map((p, n) => ({title: title + (n ? '（続き）' : ''), lead,
    elements: [{type: 'table', x: S.x, y: S.top, columns, rows: p.rows}]}));
}

/** 作る前の確かめ（資料を作る画面）の、報告案件版の説明（MeetingSchedule.gs の checkMeetingRun が、会議サマリの確かめの前に付ける） */
function mtgReportCheckText_() {
  const o = MTG_OPTIONS;
  return '【報告案件版】数字のページ（サマリー〜RXシナリオ）のあとに、シートの「' + o.columns.reportFlag + '」にチェックを入れた案件を1件1枚で載せ、' +
    'そのあとに「' + MTG_TITLES.progress + '」（報告案件を除く全件）を載せます（期待値上位案件・Appendix は入れません）。ファイル名の後ろに「' + o.report.fileSuffix + '」が付きます。';
}

/**
 * 報告案件版のスライドの並び：表紙 → 前回の宿題事項 → 数字のページ（mtgSlides_ と同じ部品）→ 報告案件（1件1枚。無ければ「報告案件なし」の1枚）
 * → 前回から変化のある案件（会議サマリと同じ表を全件。入らなければ「（続き）」のページ）。Appendix は入れない
 */
function mtgReportSlides_(s) {
  const o = MTG_OPTIONS;
  // サマリーは「現在の状態（進行中案件）」（受注月超過・停滞・商談記録未登録）の表を出さない（報告案件版は Appendix の確認事項も無いため）。
  // 表の金額の列は期待値、4つ目の数字はシナリオ見込額（mtgSummarySlide_ の options）
  const main = [mtgSummarySlide_(s, {states: false, weighted: true, scenarioTile: true})].concat(o.plan.enabled ? [mtgPlanSlide_(s.plan, s.now)] : [], [mtgBridgeSlide_(s)],
    o.portfolio.enabled ? mtgPortfolioSlides_(s.portfolio) : []);
  const reported = s.reported || [];
  const reports = reported.length ? reported.map((item, i) => mtgReportSlide_(s, item, i + 1)) : [mtgReportEmptySlide_(s)];
  // 報告案件のあとに、前回から変化のある案件（Appendix ① が無いので全件を載せ、「全件は Appendix ① を参照」と書かない）。
  // 報告フラグのある案件は、それぞれのページがあるので表に載せない（2026/11/11）
  const moves = mtgMovesSlide_(s, false, {all: true, exclude: new Set(reported.map(item => item.key))});
  // 表紙のあと、サマリーの前に前回の会議の宿題事項
  return mtgNumberSlides_(mtgSheetFrame_(s), [mtgCoverSlide_(s)].concat(mtgHomeworkSlides_(s), main, reports, moves));
}

/**
 * 報告案件の中身（mtgSummarize_ の reported）：報告フラグにチェックがある案件（シートの順）に、前回の値（前回からの変化 changed の base。
 * 変化が無ければ今の値と同じ。前回に無かった案件は isNew）を付けたもの
 */
function mtgReportItems_(table, changed) {
  if (!table.has.reportFlag) return [];
  return table.deals.filter(d => d.reportFlag).map(d => {
    const ch = changed.get(d.key);
    const base = ch ? ch.base : {status: d.status, weighted: d.weighted, monthText: d.orderMonthText, activity: d.activity, issue: d.issue,
      nextAction: d.nextAction};
    // 課題・活動状況・次回アクションの前回：記録（スナップショット・時点コピー・変更の記録）に無ければ undefined（帯は「データなし」）
    const text = key => (base && base[key] != null ? String(base[key]) : undefined);
    return Object.assign({}, d, {hasOwner: !!table.has.owner, hasBx: !!table.has.bx, isNew: !base, prevStatus: base ? base.status || '' : '', prevWeighted: base ? base.weighted : null,
      prevMonthText: base ? base.monthText || '' : '', prevActivity: text('activity'), prevIssue: text('issue'), prevNextAction: text('nextAction')});
  });
}

/** 報告フラグのセルの値がチェック（TRUE）か（チェックボックスの true・文字の「TRUE」） */
function mtgReportFlag_(value, text) {
  return value === true || /^true$/i.test(String(text == null ? '' : text).trim());
}

/** ステータスの確度（「A-受注済み(100%)」→ A・「創注-仮検討(10%)」→ 創注。分からなければ空） */
function mtgReportRank_(status) {
  const m = /^\s*(創注|[A-D])/.exec(String(status || '').normalize('NFKC'));
  return m ? m[1] : '';
}

/** 報告案件のスライドの位置（pt。720 × 405）。左の列に数字の5つの枠、右の列に文章の枠（戦略・価値創出・下に3つ） */
function mtgReportLayout_() {
  const S = MTG_SLIDE;
  const top = 48;
  const bottom = S.bottom - 1;
  const gap = 6;
  const left = {x: S.x, w: 150};   // 左の列（数字）は狭く、右の列（文章）を広く
  const right = {x: S.x + left.w + 8, w: S.w - left.w - 8};
  const strategyH = 100;
  const valueH = 56;
  const rowY = top + strategyH + gap + valueH + gap;
  // 左の5つの枠の高さは右の列の段にそろえる：案件・得意先情報 = 戦略の枠、ステータス = 価値創出の枠、
  // 期待値・受注月（右半分は売上）・RXシナリオ = 下の段（課題・活動状況・次回アクション）を3つに分けた高さ
  const lowH = (bottom - rowY - 2 * gap) / 3;
  const leftH = [strategyH, valueH, lowH, lowH, lowH];
  return {top, bottom, gap, left, right, leftH, strategyH, valueH, rowY, rowH: bottom - rowY, rowW: (right.w - 2 * gap) / 3,
    head: 12 + MTG_OPTIONS.report.headTop,    // head：枠の見出しの分（枠の上端から本文の箱の上端まで。見出しの上の余白 headTop を含む）
    prevH: 58};  // prevH：活動状況・次回アクションの枠の下の、前回の内容の帯の分（帯の上下のすき間を含む）
}

/** 報告案件の1枚（item は mtgReportItems_ の1件、no は番号（1 から））。ふつうの見出し・リード文の代わりに「RXシナリオ進捗における報告案件①」（mtgReportHeader_） */
function mtgReportSlide_(s, item, no) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const R = o.report;
  const L = mtgReportLayout_();
  const el = mtgReportHeader_(item, mtgReportStaff_(item), no);
  const blank = o.blankText;
  // 右の列：文章の枠（字の大きさは枠ごとに、入るいちばん大きい字。見出しより大きくしない）
  const texts = [
    {key: 'strategy', label: o.columns.strategy, x: L.right.x, y: L.top, w: L.right.w, h: L.strategyH, fill: c.greenSoft, line: null},
    {key: 'valueCase', label: o.columns.valueCase, x: L.right.x, y: L.top + L.strategyH + L.gap, w: L.right.w, h: L.valueH},
    // 下の段は 活動における課題 → 活動状況 → 次回アクション（活動状況と次回アクションを隣に並べる）
    {key: 'issue', label: o.columns.issue, x: L.right.x, y: L.rowY, w: L.rowW, h: L.rowH},
    {key: 'activity', label: o.columns.activity, x: L.right.x + L.rowW + L.gap, y: L.rowY, w: L.rowW, h: L.rowH},
    {key: 'nextAction', label: o.columns.nextAction, x: L.right.x + 2 * (L.rowW + L.gap), y: L.rowY, w: L.rowW, h: L.rowH}
  ].map(b => Object.assign(b, {text: String(item[b.key] || '').trim()}));
  // 課題・活動状況・次回アクションは、枠の下に前回の内容を灰色の帯で並べる（今回の文はその上。mtgReportPrevious_）
  const withPrev = {issue: 'prevIssue', activity: 'prevActivity', nextAction: 'prevNextAction'};
  texts.forEach(b => { b.bodyH = b.h - L.head - (withPrev[b.key] ? L.prevH : 0); });
  const fit = mtgReportFit_(texts.map(b => ({text: b.text, w: b.w, h: b.bodyH})));
  texts.forEach((b, i) => {
    el.push(mtgReportFrame_(b.x, b.y, b.w, b.h, String(b.label).replace(/\s+/g, ''), {fill: b.fill, line: b.line}));
    el.push(mtgBox_(b.x, b.y + L.head, b.w, b.bodyH, b.text ? fit.texts[i] : blank,
      {size: fit.sizes[i], color: b.text ? c.ink : c.faint, valign: 'TOP', role: 'report-text'}));
    if (withPrev[b.key]) el.push.apply(el, mtgReportPrevious_(b, item, item[withPrev[b.key]], fit.sizes[i], mtgReportPrevLabel_(s)));
  });
  // 左の列：先方・ステータス・期待値・受注月・RXシナリオ
  const box = k => ({x: L.left.x, y: L.top + L.leftH.slice(0, k).reduce((sum, h) => sum + h + L.gap, 0), w: L.left.w, h: L.leftH[k]});
  el.push.apply(el, mtgReportClient_(box(0), item));
  el.push.apply(el, mtgReportStatus_(box(1), item));
  el.push.apply(el, mtgReportValue_(box(2), item));
  // 受注月の枠は左右に分け、右に年度別売上（FY26〜28。2026/11/11）。右の枠は、見出し（「年度別売上（百万）」）が1行に入る幅（半分より狭くしない）
  const month = box(3);
  const salesW = Math.min(month.w - L.gap - 50, Math.max((month.w - L.gap) / 2, mtgReportHeadWidth_(o.report.salesLabel) + MTG_SLIDE.padX + 2));
  const monthW = month.w - L.gap - salesW;
  el.push.apply(el, mtgReportMonth_(Object.assign({}, month, {w: monthW}), item));
  el.push.apply(el, mtgReportSales_(Object.assign({}, month, {x: month.x + monthW + L.gap, w: salesW}), item, s));
  el.push.apply(el, mtgReportScenario_(box(4), item, s.scenarioColumn));
  // 出典の行は省略しない：1行に入る大きさまで小さくする（footSize から 0.5pt ずつ minSize まで）
  const footRight = mtgReportRecord_(item.record, s.recordColumn, item.recordLinks);
  const source = mtgReportSource_(s, item);
  const sourceW = MTG_SLIDE.w - 60 - footRight.w;
  let footSize = R.footSize;
  while (footSize > R.minSize && mtgWrap_(source, sourceW, footSize, 2).length > 1) footSize -= 0.5;
  return {title: mtgReportTitle_(mtgDealLabel_(item), no), noTitle: true, source, footSize, footRight, cut: fit.cut, elements: el};
}

/** 「前回」の帯の見出し：前回の会議の資料と比べるなら「前回（10/30 会議）」、ほか（初回・選んだ時点・基準）は「前回」（いつかは出典の行と表紙） */
function mtgReportPrevLabel_(s) {
  const p = s.previous;
  return p && p.date && !p.chosen && !p.baseline ? '前回（' + mtgShortDate_(p.date) + ' 会議）' : '前回';
}

/**
 * 課題・活動状況・次回アクションの枠の下の「前回」の帯（灰色の地に、前回の内容を灰色の字で）。b は枠 {x, y, w, h, text: 今回の文}、prev は前回の文
 * （undefined なら分からない）。今回と同じなら、帯に「前回から更新なし」と書き、枠の右上に「更新なし」の印を付ける。前回に無かった案件は「（新規）」。
 * 字は MTG_OPTIONS.report.prevSize（今回の文より小さく）。省略しない：入らなければ帯に入る大きさまで小さくする（minSize まで）
 */
function mtgReportPrevious_(b, item, prev, size, label) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const R = o.report;
  const L = mtgReportLayout_();
  const S = MTG_SLIDE;
  const before = prev == null ? null : String(prev).trim();
  const same = !item.isNew && before != null && before !== '' && credNormalize_(before) === credNormalize_(b.text);
  const text = item.isNew ? R.newText.replace(/[（）]/g, '') + 'の案件のため、前回はありません'
    : before == null ? o.gapText : same ? '今回と同じ（前回から更新なし）' : before || o.blankText;
  const muted = item.isNew || before == null || same || !before;
  const band = {x: b.x + 4, y: b.y + b.h - L.prevH + 2, w: b.w - 8, h: L.prevH - 6};
  const room = band.h - 8.5;   // 「前回」の見出しの分を除く
  let small = Math.min(R.prevSize, size);
  while (small > R.minSize && mtgTextHeight_(mtgWrap_(text, band.w, small, 999).length, small) > room + 0.01) small -= 0.5;
  const body = text;
  const el = [
    mtgBox_(band.x, band.y, band.w, band.h, label || '前回', {size: 7, bold: true, color: c.muted, fill: c.graySoft, shape: 'RECTANGLE', valign: 'TOP', role: 'report-prev-frame'}),
    mtgBox_(band.x, band.y + 8.5, band.w, band.h - 8.5, body, {size: small, color: muted ? c.faint : c.sub, valign: 'TOP', role: 'report-prev'})
  ];
  if (same) {
    const badge = mtgReportBadge_(0, b.y + 8.5, '更新なし', 'stale', 'report-stale', 7);
    el.push(Object.assign(badge, {x: b.x + b.w - badge.w - 4}));
  }
  return el;
}

/**
 * 報告案件の出典の行：どのシートの何行目の値を、いつ読んだか（資料を作った時刻）だけ。
 * 例：「出典：「スーパー・GMS」シート 12行目（2026/11/06 10:00 時点）」。試作の印・前回の時点は書かない（前回の時点は表紙と、枠の「前回（10/30 会議）」にある）。
 * ダミーの記録を使ったときだけ、先頭に【ダミーデータを含む】（配る資料に紛れないように）
 */
function mtgReportSource_(s, item) {
  const at = Utilities.formatDate(s.now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  return (s.dummy ? '【ダミーデータを含む】' : '') + '出典：「' + s.name + '」シート ' + item.row + '行目（' + at + ' 時点）';
}

/** 報告フラグにチェックのある案件が無いとき（列が無いときは「データなし」と理由） */
function mtgReportEmptySlide_(s) {
  const o = MTG_OPTIONS;
  const R = o.report;
  const S = MTG_SLIDE;
  const L = mtgReportLayout_();
  const gap = !s.reportColumn;
  const text = gap ? 'シートに「' + o.columns.reportFlag + '」の列が無い（見出しの行に列を足し、チェックボックスにしてください）'
    : R.noneText;
  return {title: R.title, noTitle: true, empty: gap ? 'gap' : 'none',
    elements: mtgReportHeader_(null).concat([mtgEmptyBox_(S.x, L.top + 40, S.w, 80, gap ? 'gap' : 'none', text)])};
}

/**
 * いちばん上の見出し「RXシナリオ進捗における報告案件①」（MTG_OPTIONS.report.title と番号。案件名・得意先名は左上の案件・得意先情報の枠に書くので、
 * 見出しには入れない）。字の大きさはほかのスライドの見出しと同じ（MTG_SLIDE.titleSize。入らなければ小さくする。省略はしない）。
 * item が無ければ番号なし（報告案件が無いときの1枚）。no は報告案件の番号（1 から）。
 * staff（担当営業・BX担当の2行の文）があれば右上に2行（見出しの半分の大きさで、2行の高さが見出しの行と同じ。灰色の右寄せ）で書き、その分だけ見出しの幅を狭める
 */
function mtgReportHeader_(item, staff, no) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const S = MTG_SLIDE;
  const small = S.titleSize / 2;   // 2行（行の送り 1.2）の高さ = 見出し1行の高さ
  const staffLines = item && staff ? String(staff).split('\n') : [];
  // 右上の担当は省略しない：幅（staffWidth まで）に入らなければ字を小さくする
  const need = size => Math.ceil(Math.max.apply(null, staffLines.map(t => mtgTextWidth_(t, size))) / S.fit + S.padX + 2);
  let staffSize = small;
  while (staffLines.length && staffSize > R.minSize && need(staffSize) > R.staffWidth) staffSize -= 0.5;
  const staffW = staffLines.length ? Math.min(R.staffWidth, need(staffSize)) : 0;
  // 案件名は左上の案件・得意先情報の枠に書く（見出しに入れると長い案件名が入らないため）
  const full = item ? mtgReportTitle_('', no).replace(/：$/, '') : R.title;
  const w = S.w - staffW;
  let size = S.titleSize;
  while (size > 10 && mtgWrap_(full, w, size, 2).length > 1) size -= 1;
  const el = [mtgBox_(S.x, S.titleY, w, 30, full, {size, bold: true, color: c.ink, role: 'report-title'})];
  if (staffLines.length) {
    el.push(mtgBox_(S.x + w, S.titleY, staffW, 30, staffLines.join('\n'), {size: staffSize, color: c.muted, align: 'END', role: 'report-staff'}));
  }
  return el;
}

/**
 * 報告案件の見出し「RXシナリオ進捗における報告案件①：得意先／案件名」（スライドの名前。前の部分は MTG_OPTIONS.report.title。label が空なら「…報告案件①：」。
 * 番号は ①〜⑳、21 からは「(21)」。no が無ければ「…報告案件：」）
 */
function mtgReportTitle_(label, no) {
  const R = MTG_OPTIONS.report;
  const mark = !no ? '' : no <= 20 ? String.fromCharCode(0x2460 + no - 1) : '(' + no + ')';
  return R.title + mark + '：' + label;
}

/** 右上の担当：1行目「担当営業：佐藤」・2行目「BX担当：田中」（改行でつなぐ。空なら「未入力」。列が無い項目は出さない） */
function mtgReportStaff_(item) {
  const o = MTG_OPTIONS;
  const parts = [];
  if (item.hasOwner !== false) parts.push('担当営業：' + (item.owner || o.blankText));
  if (item.hasBx !== false) parts.push('BX担当：' + (item.bx || o.blankText));
  return parts.join('\n');
}

/** 枠（四角）と、その左上の見出し（緑の太字）。opt = {fill, line}（line: null なら枠の線なし） */
function mtgReportFrame_(x, y, w, h, heading, opt) {
  const c = MTG_OPTIONS.colors;
  const o = opt || {};
  const line = 'line' in o && o.line === null ? null : o.line || c.line;
  // 見出しの（）書き（「期待値（百万円）」の「（百万円）」など）は小さく（report.headNoteSize）
  const note = String(heading || '').indexOf('（');
  return mtgBox_(x, y, w, h, heading, Object.assign({size: 8, bold: true, color: c.greenDark, fill: o.fill || c.white,
    line, lineWeight: 0.75, shape: 'RECTANGLE', valign: 'TOP', spaceAbove: MTG_OPTIONS.report.headTop, role: 'report-frame'},
    note > 0 ? {spans: [{start: note, end: String(heading).length, size: MTG_OPTIONS.report.headNoteSize}]} : {}));
}

/** 枠の見出しの幅（pt。（）書きは report.headNoteSize の大きさで数える。mtgReportFrame_ と同じ） */
function mtgReportHeadWidth_(heading) {
  const text = String(heading || '');
  const note = text.indexOf('（');
  return note > 0 ? mtgTextWidth_(text.slice(0, note), 8) + mtgTextWidth_(text.slice(note), MTG_OPTIONS.report.headNoteSize) : mtgTextWidth_(text, 8);
}

/**
 * 文章の枠の字の大きさ：どの枠にも収まるいちばん大きい大きさ（MTG_OPTIONS.report の maxSize から 0.5pt ずつ minSize まで）。省略はしない
 * （minSize でも入らなければ minSize のまま全文を入れる。cut: true で知らせる）。boxes = [{text, w, h: 本文の箱の高さ}]。{size, texts: 箱に入れる文（全文）, cut}
 */
function mtgReportFit_(boxes) {
  const R = MTG_OPTIONS.report;
  const lines = (b, size) => mtgWrap_(b.text, b.w, size, 999).length;
  const fits = (b, size) => !b.text || mtgTextHeight_(lines(b, size), size) <= b.h;
  // 枠ごとに、入るいちばん大きい字（maxSize = 見出しの大きさから 0.5pt ずつ。2026/11/11 から。前の版はスライドの中で同じ大きさにしていたので、
  // 文の多い枠に合わせて、大きな枠（戦略など）の字まで小さくなっていた）
  const sizes = boxes.map(b => {
    for (let size = R.maxSize; size >= R.minSize - 1e-9; size -= 0.5) if (fits(b, size)) return size;
    return R.minSize;
  });
  const cut = boxes.some((b, i) => !fits(b, sizes[i]));
  return {size: Math.min.apply(null, sizes.concat([R.maxSize])), sizes, texts: boxes.map(b => b.text), cut};
}

/**
 * 字を小さくして箱に収める倍率（1 から 0.05 ずつ。いちばん大きい字が MTG_OPTIONS.report.minSize になるまで）。
 * lines = [{text, size}]、w・h は箱の幅と高さ（上下左右の余白を含む）。省略しないための、左の列の枠の中身用
 */
function mtgReportShrink_(lines, w, h) {
  const S = MTG_SLIDE;
  const R = MTG_OPTIONS.report;
  // 高さは上の余白だけを数える（下の余白には文字がかかってもよい。箱の下の線・確度の段とは重ならない）
  const height = k => lines.reduce((sum, l) => sum + mtgWrap_(l.text, w, l.size * k, 999).length * l.size * k * S.lineFactor, 0) + S.padY / 2;
  const min = Math.min(1, R.minSize / Math.max.apply(null, lines.map(l => l.size).concat([1])));
  for (let k = 1; k >= min - 1e-9; k -= 0.05) {
    if (height(k) <= h + 0.01) return k;
  }
  return min;
}

/**
 * 本文の箱（枠の見出しのすぐ下から。左の列の枠の中身）。lines = [{text, size, bold, color}]（空の行は除く）。
 * 省略しない：全文を入れ、箱（下に bottomSpace をあける）に入らなければ、すべての行の字を同じ割合で小さくする（mtgReportShrink_）。戻り値の scale はその倍率
 */
function mtgReportBody_(b, lines, bottomSpace) {
  const L = mtgReportLayout_();
  const list = lines.filter(line => line && line.text);
  const h = b.h - L.head - (bottomSpace || 0);
  const k = mtgReportShrink_(list, b.w, h);
  const size = v => Math.round(v * k * 10) / 10;
  const text = list.map(line => line.text).join('\n');
  let at = 0;
  const spans = list.map(line => {
    const span = {start: at, end: at + line.text.length, size: size(line.size), bold: !!line.bold, color: line.color};
    at += line.text.length + 1;
    return span;
  });
  return Object.assign(mtgBox_(b.x, b.y + L.head, b.w, h, text, {size: list.length ? size(list[0].size) : MTG_OPTIONS.report.boxSize,
    color: MTG_OPTIONS.colors.ink, spans, valign: 'TOP', role: 'report-value'}), {scale: k});
}

/**
 * 左の列の1つ目：案件・得意先情報（案件名（太字）、1行あけて得意先名・先方部門／先方担当役職。どれも全文を書く（省略しない。
 * 枠に入らなければ字を小さくする）。名前（先方担当氏名）は出さない。空の項目は灰色の「未入力」）
 */
function mtgReportClient_(b, item) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const blank = MTG_OPTIONS.blankText;
  const dept = [item.clientDept, item.clientTitle].filter(Boolean).join('／');
  const line = (text, bold) => ({text: text || blank, size: R.boxSize, bold: !!bold && !!text, color: text ? c.ink : c.faint});
  return [mtgReportFrame_(b.x, b.y, b.w, b.h, R.clientLabel), mtgReportBody_(b, [
    line(item.project, true),
    {text: ' ', size: 4},   // 案件名と得意先の情報の間をあける
    line(item.customer),
    line(dept)
  ])];
}

/**
 * 受注月の枠の右半分：売上（年度別の売上の列。今年度から report.salesYears 年度分を、年度ごとに前半期＋後半期で。FY26・FY27・FY28 を縦に並べ、
 * 左に年度・右に値）。空は「−」、シートに今年度の列が無ければ「データなし」
 */
function mtgReportSales_(b, item, s) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const R = o.report;
  const S = MTG_SLIDE;
  const L = mtgReportLayout_();
  const year = mtgFiscalYearKey_(s.now);
  const years = Array.from({length: R.salesYears}, (_, k) => ('0' + ((Number(year) + k) % 100)).slice(-2));
  const el = [mtgReportFrame_(b.x, b.y, b.w, b.h, R.salesLabel)];
  const has = (s.salesYears || []).indexOf(year) >= 0;
  if (!has) {
    el.push(mtgBox_(b.x, b.y + L.head, b.w, b.h - L.head, o.gapText, {size: R.subSize, color: c.faint, valign: 'MIDDLE', role: 'report-sales'}));
    return el;
  }
  // 年度ごとの行（左に「FY26」、右に値）。見出しの下を年度の数で等分する（行の高さは 12pt まで）
  const top = b.y + L.head + 1;
  const rowH = Math.min(12, (b.y + b.h - 2 - top) / years.length);
  years.forEach((y, k) => {
    const v = (item.sales || {})[y];
    const ry = top + rowH * k;
    el.push(mtgBox_(b.x, ry, b.w / 2, rowH, 'FY' + y, {size: R.subSize, color: c.muted, valign: 'MIDDLE', role: 'report-sales-head'}));
    el.push(mtgBox_(b.x + b.w / 2 - S.padX, ry, b.w / 2 + S.padX, rowH, v == null ? '−' : mtgAmount_(v),
      {size: R.boxSize, bold: v != null, color: v == null ? c.faint : c.ink, align: 'END', valign: 'MIDDLE', role: 'report-sales'}));
  });
  return el;
}

/** 左の列の2つ目：ステータス（今のステータス・前回・確度の段（創注・D・C・B・A。今は緑・前回は薄い緑）） */
function mtgReportStatus_(b, item) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const blank = MTG_OPTIONS.blankText;
  const prev = item.isNew ? R.newText : item.prevStatus || blank;
  const stepH = 10;
  const el = [mtgReportFrame_(b.x, b.y, b.w, b.h, 'ステータス'),
    mtgReportBody_(b, [{text: item.status || blank, size: R.boxSize, bold: true, color: item.status ? c.ink : c.faint},
      {text: '前回：' + prev, size: R.subSize, color: c.muted}], stepH + 4)];
  const now = mtgReportRank_(item.status);
  const before = item.isNew ? '' : mtgReportRank_(item.prevStatus);
  const steps = R.steps;
  // 段の幅は文字の幅に合わせる（「創注」は2文字なので広く。箱の左右の余白 MTG_SLIDE.padX の分も要る）
  const weight = step => Math.max(1, mtgTextWidth_(step, 7) / 6);
  const total = steps.reduce((sum, step) => sum + weight(step), 0);
  const unit = (b.w - 8 - (steps.length - 1) * 2) / total;
  let x = b.x + 4;
  steps.forEach(step => {
    const on = step === now;
    const was = !on && step === before;
    const sw = unit * weight(step);
    const at = x;
    x += sw + 2;
    el.push(mtgBox_(at, b.y + b.h - stepH - 4, sw, stepH, step, {size: 7, bold: on, align: 'CENTER', shape: 'RECTANGLE',
      fill: on ? c.green : was ? c.greenSoft : c.graySoft, color: on ? c.white : was ? c.greenDark : c.faint, role: 'report-step'}));
  });
  return el;
}

/** 左の列の3つ目：期待値（百万円。1行目に今の値と、その右に前回からの増減（色付きの字）・2行目に前回の値と想定売上） */
function mtgReportValue_(b, item) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const S = MTG_SLIDE;
  const L = mtgReportLayout_();
  const now = item.weighted == null ? MTG_OPTIONS.blankText : mtgAmount_(item.weighted);
  const delta = item.isNew || item.weighted == null || item.prevWeighted == null ? null : item.weighted - item.prevWeighted;
  const rounded = delta == null ? null : Math.round(delta * 10) / 10;
  const second = '前回：' + (item.isNew ? R.newText : mtgAmount_(item.prevWeighted)) + '　／　想定売上：' + mtgAmount_(item.amount);
  // 行ごとに高さを決めた箱に分け、上下の中央に置く（数字の行と、その下の「前回」の行）
  const rowH = 13;
  const y1 = b.y + L.head + 5;   // 枠の見出し（下端は枠の上から約 17pt）にバッジがかからないように下げる
  const y2 = y1 + rowH;
  let sub = R.subSize;   // 省略しない：1行に入らなければ小さくする
  while (sub > R.minSize && mtgWrap_(second, b.w, sub, 2).length > 1) sub -= 0.5;
  // 前回からの増減は、数字の右に色付きの字で（増 = 青「+20」・減 = 赤「▲12」・同じ = 灰色「±0」。数字と同じ大きさ）
  const deltaText = rounded == null ? '' : '　' + mtgSigned_(rounded);
  const first = now + deltaText;
  const spans = [{start: 0, end: now.length, bold: true, color: item.weighted == null ? c.faint : c.ink}];
  if (deltaText) spans.push({start: now.length, end: first.length, bold: true, color: rounded > 0 ? c.plus : rounded < 0 ? c.minus : c.faint});
  return [mtgReportFrame_(b.x, b.y, b.w, b.h, '期待値（百万円）'),
    mtgBox_(b.x, y1, b.w, rowH, first, {size: R.boxSize, color: c.ink, spans, valign: 'MIDDLE', role: 'report-value'}),
    mtgBox_(b.x, y2, b.w, rowH, second, {size: sub, color: c.muted, valign: 'MIDDLE', role: 'report-value'})];
}

/**
 * 平らなバッジ（角の無い四角・枠の線なし・薄い地に色の字。確度の段と同じつくり。「更新なし」の印など）。x は左端、cy は上下の中央、
 * tone：'up' 緑・'down' 濃いグレー・'flat' 灰色・'stale' 黄色（更新なし）・'chip' 緑（シナリオの番号）。size は字の大きさ（省けば MTG_OPTIONS.report.subSize）、height は高さ（省けば字の大きさ + 4）
 */
function mtgReportBadge_(x, cy, text, tone, role, size, height) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const pt = size || R.subSize;
  const h = height || pt + 4;
  const w = Math.ceil(mtgTextWidth_(text, pt) / 0.92 + MTG_SLIDE.padX + 2);
  const tones = {
    up: {fill: c.greenSoft, color: c.greenDark},
    down: {fill: c.graySoft, color: c.dark},
    flat: {fill: c.graySoft, color: c.faint},
    stale: {fill: R.staleFill, color: c.ink},
    chip: {fill: c.greenSoft, color: c.greenDark}
  };
  return mtgBox_(x, cy - h / 2, w, h, text, Object.assign({size: pt, bold: tone !== 'chip', align: 'CENTER', valign: 'MIDDLE', shape: 'RECTANGLE', role},
    tones[tone] || tones.flat));
}

/** 左の列の4つ目：受注月（今の月・前回の月（遅くなったら「（延期）」・早くなったら「（前倒し）」）） */
function mtgReportMonth_(b, item) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const now = mtgMonth_(item.orderMonthText);
  const before = item.isNew ? null : mtgMonth_(item.prevMonthText);
  const move = now && before ? mtgMonthIndex_(now) - mtgMonthIndex_(before) : 0;
  const label = m => m.y + '/' + String(m.m).padStart(2, '0');
  const nowText = now ? label(now) : String(item.orderMonthText || '').trim() || MTG_OPTIONS.blankText;
  const prevText = item.isNew ? R.newText : before ? label(before) : String(item.prevMonthText || '').trim() || MTG_OPTIONS.blankText;
  return [mtgReportFrame_(b.x, b.y, b.w, b.h, '受注月'), mtgReportBody_(b, [
    {text: nowText, size: R.boxSize, bold: true, color: now || item.orderMonthText ? c.ink : c.faint},
    {text: '前回：' + prevText + (move > 0 ? '（延期）' : move < 0 ? '（前倒し）' : ''), size: R.subSize, color: c.muted}
  ])];
}

/**
 * 左の列の5つ目：RXシナリオ（1行目に見込金額、その下に選んだシナリオの番号を角の丸い枠（バッジ）で並べる）。
 * 省略しない：すべてのバッジを並べ、枠に入らなければバッジの字を小さくする（chipSize から minSize まで）
 */
function mtgReportScenario_(b, item, hasColumn) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const R = o.report;
  const S = MTG_SLIDE;
  const L = mtgReportLayout_();
  const el = [mtgReportFrame_(b.x, b.y, b.w, b.h, 'RXシナリオ')];
  const amount = '見込金額：' + (item.scenarioAmount == null ? o.blankText : mtgAmount_(item.scenarioAmount) + '百万円');
  el.push(mtgBox_(b.x, b.y + L.head, b.w, R.subSize * S.lineFactor + S.padY, amount, {size: R.subSize, color: c.muted, valign: 'TOP', role: 'report-value'}));
  // シナリオは番号だけ（「01-05」など。名前は長いので。番号が無ければ名前のまま。mtgScenarioNumber_）
  const names = (item.scenarios || []).map(mtgScenarioNumber_);
  const top = b.y + L.head + S.padY / 2 + R.subSize * S.lineFactor + 3;
  if (!hasColumn || !names.length) {
    el.push(mtgBox_(b.x, top - S.padY / 2, b.w, b.y + b.h - top + S.padY / 2, hasColumn ? o.blankText : o.gapText,
      {size: R.subSize, color: c.faint, valign: 'TOP', role: 'report-value'}));
    return el;
  }
  const x0 = b.x + 7;
  const right = b.x + b.w - 7;
  const gap = 3;
  const layout = size => {
    const h = size + 4.5;
    // 幅は文字の幅に箱の左右の余白（MTG_SLIDE.padX）を足す（足りないと「01-05」が2行に折り返す）
    const width = text => Math.min(right - x0, Math.ceil(mtgTextWidth_(text, size) / 0.92 + S.padX + 4));
    const placed = [];
    let x = x0;
    let row = 0;
    names.forEach(name => {
      const w = width(name);
      if (x + w > right && x > x0) {
        row++;
        x = x0;
      }
      placed.push({text: name, x, row, w});
      x += w + gap;
    });
    return {size, h, placed, height: (row + 1) * (h + gap) - gap};
  };
  let fit = layout(R.chipSize);
  while (fit.size > R.minSize && top + fit.height > b.y + b.h - 4) fit = layout(fit.size - 0.5);
  fit.placed.forEach(p => el.push(mtgBox_(p.x, top + p.row * (fit.h + gap), p.w, fit.h, p.text, {size: fit.size, align: 'CENTER', valign: 'MIDDLE',
    shape: 'RECTANGLE', fill: c.greenSoft, color: c.greenDark, role: 'report-chip'})));
  return el;
}

/**
 * 右下の「商談記録：格納先リンク」（「格納先リンク」を商談記録のセルのリンクへのハイパーリンクに。URL が2つ以上なら「格納先リンク①　②」と
 * すべて）。セルにリンクが無く文字だけなら、その文字（リンクなし）。空なら「未入力」、列が無ければ「データなし」。
 * mtgFrame_ が出典の行の右に置く
 */
function mtgReportRecord_(record, hasColumn, links) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const head = '商談記録：';
  const value = String(record || '').trim();
  const urls = links || [];
  // 幅は文字に合わせる（R.recordWidth まで。残りは出典の行に使う）。省略しない：入らなければ字を小さくする
  const box = (text, spans) => {
    const need = size => Math.ceil(mtgTextWidth_(text, size) / MTG_SLIDE.fit + MTG_SLIDE.padX + 2);
    let size = R.footSize;
    while (size > R.minSize && need(size) > R.recordWidth) size -= 0.5;
    return {w: Math.min(R.recordWidth, need(size)), text, size, color: c.faint, spans: spans || []};
  };
  if (!hasColumn || (!value && !urls.length)) return box(head + (hasColumn ? MTG_OPTIONS.blankText : MTG_OPTIONS.gapText));
  if (!urls.length) return box(head + value);
  let text = head;
  const spans = [];
  urls.forEach((url, i) => {
    const label = R.linkLabel + (urls.length > 1 ? (i < 20 ? String.fromCharCode(0x2460 + i) : '(' + (i + 1) + ')') : '');
    if (i) text += '　';
    spans.push({start: text.length, end: text.length + label.length, url, color: c.link, underline: true});
    text += label;
  });
  return box(text, spans);
}
