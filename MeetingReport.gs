/**
 * RX Business Meeting の会議サマリ「報告案件版」（MeetingSummary.gs の会議サマリと同じ作り方で、案件のページだけを変えたもの）。
 *
 * - 数字のページ（表紙・サマリー・計画に対する積み上げ・パイプラインの変化・RXシナリオのポートフォリオ）は、会議サマリと同じ。
 * - 案件のページ（前回から変化のある案件・期待値上位案件の状況・課題）と Appendix（扉と ①〜③）は入れない。
 *   代わりに、シートの「報告フラグ」（MTG_OPTIONS.columns.reportFlag）にチェックを入れた案件を、1件1枚の「報告案件」のスライドにする。
 * - 報告案件のスライドは枠の位置を固定（ひな形として同じ形）にし、中の文章の分量で字の大きさを変える
 *   （右の文章の枠は、スライドの中で同じ大きさ。MTG_OPTIONS.report の maxSize から minSize まで。入らなければ minSize で「…」で切る）。
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

/** 作る前の確かめ（資料を作る画面）の、報告案件版の説明（MeetingSchedule.gs の checkMeetingRun が、会議サマリの確かめの前に付ける） */
function mtgReportCheckText_() {
  const o = MTG_OPTIONS;
  return '【報告案件版】数字のページ（サマリー〜RXシナリオ）のあとに、シートの「' + o.columns.reportFlag + '」にチェックを入れた案件を1件1枚で載せます' +
    '（前回から変化のある案件・期待値上位案件・Appendix は入れません）。ファイル名の後ろに「' + o.report.fileSuffix + '」が付きます。';
}

/**
 * 報告案件版のスライドの並び：表紙 → 数字のページ（mtgSlides_ と同じ部品）→ 報告案件（1件1枚。無ければ「報告案件なし」の1枚）。Appendix は入れない
 */
function mtgReportSlides_(s) {
  const o = MTG_OPTIONS;
  const main = [mtgSummarySlide_(s)].concat(o.plan.enabled ? [mtgPlanSlide_(s.plan, s.now)] : [], [mtgBridgeSlide_(s)],
    o.portfolio.enabled ? mtgPortfolioSlides_(s.portfolio) : []);
  const reported = s.reported || [];
  const reports = reported.length ? reported.map(item => mtgReportSlide_(s, item)) : [mtgReportEmptySlide_(s)];
  return mtgNumberSlides_(mtgSheetFrame_(s), [mtgCoverSlide_(s)].concat(main, reports));
}

/**
 * 報告案件の中身（mtgSummarize_ の reported）：報告フラグにチェックがある案件（シートの順）に、前回の値（前回からの変化 changed の base。
 * 変化が無ければ今の値と同じ。前回に無かった案件は isNew）を付けたもの
 */
function mtgReportItems_(table, changed) {
  if (!table.has.reportFlag) return [];
  return table.deals.filter(d => d.reportFlag).map(d => {
    const ch = changed.get(d.key);
    const base = ch ? ch.base : {status: d.status, weighted: d.weighted, monthText: d.orderMonthText};
    return Object.assign({}, d, {isNew: !base, prevStatus: base ? base.status || '' : '', prevWeighted: base ? base.weighted : null,
      prevMonthText: base ? base.monthText || '' : ''});
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
  const left = {x: S.x, w: 180};
  const right = {x: S.x + left.w + 8, w: S.w - left.w - 8};
  const leftH = (bottom - top - 4 * gap) / 5;
  const strategyH = 128;
  const valueH = 66;
  const rowY = top + strategyH + gap + valueH + gap;
  return {top, bottom, gap, left, right, leftH, strategyH, valueH, rowY, rowH: bottom - rowY, rowW: (right.w - 2 * gap) / 3,
    head: 12};   // head：枠の見出しの分（枠の上端から本文の箱の上端まで）
}

/** 報告案件の1枚（item は mtgReportItems_ の1件）。見出し・リード文は付けず、いちばん上に「報告案件：得意先／案件名」 */
function mtgReportSlide_(s, item) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const R = o.report;
  const L = mtgReportLayout_();
  const el = mtgReportHeader_(mtgDealLabel_(item));
  const blank = o.blankText;
  // 右の列：文章の枠（字の大きさはスライドの中で同じ）
  const texts = [
    {key: 'strategy', label: o.columns.strategy, x: L.right.x, y: L.top, w: L.right.w, h: L.strategyH, fill: c.greenSoft, line: null},
    {key: 'valueCase', label: o.columns.valueCase, x: L.right.x, y: L.top + L.strategyH + L.gap, w: L.right.w, h: L.valueH},
    {key: 'activity', label: o.columns.activity, x: L.right.x, y: L.rowY, w: L.rowW, h: L.rowH},
    {key: 'issue', label: o.columns.issue, x: L.right.x + L.rowW + L.gap, y: L.rowY, w: L.rowW, h: L.rowH},
    {key: 'nextAction', label: o.columns.nextAction, x: L.right.x + 2 * (L.rowW + L.gap), y: L.rowY, w: L.rowW, h: L.rowH}
  ].map(b => Object.assign(b, {text: String(item[b.key] || '').trim()}));
  const fit = mtgReportFit_(texts.map(b => ({text: b.text, w: b.w, h: b.h - L.head})));
  texts.forEach((b, i) => {
    el.push(mtgReportFrame_(b.x, b.y, b.w, b.h, String(b.label).replace(/\s+/g, ''), {fill: b.fill, line: b.line}));
    el.push(mtgBox_(b.x, b.y + L.head, b.w, b.h - L.head, b.text ? fit.texts[i] : blank,
      {size: fit.size, color: b.text ? c.ink : c.faint, valign: 'TOP', role: 'report-text'}));
  });
  // 左の列：先方・ステータス・期待値・受注月・RXシナリオ
  const box = k => ({x: L.left.x, y: L.top + k * (L.leftH + L.gap), w: L.left.w, h: L.leftH});
  el.push.apply(el, mtgReportClient_(box(0), item));
  el.push.apply(el, mtgReportStatus_(box(1), item));
  el.push.apply(el, mtgReportValue_(box(2), item));
  el.push.apply(el, mtgReportMonth_(box(3), item));
  el.push.apply(el, mtgReportScenario_(box(4), item, s.scenarioColumn));
  return {title: R.title + '：' + mtgDealLabel_(item), noTitle: true, note: R.note, footRight: mtgReportRecord_(item.record, s.recordColumn),
    cut: fit.cut, elements: el};
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
    elements: mtgReportHeader_('').concat([mtgEmptyBox_(S.x, L.top + 40, S.w, 80, gap ? 'gap' : 'none', text)])};
}

/** いちばん上の行「報告案件：得意先／案件名」（1行に入る大きさ。入らなければ「…」で切る）と、その下の緑の線 */
function mtgReportHeader_(label) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const S = MTG_SLIDE;
  const prefix = R.prefix;
  const full = label ? prefix + label : R.title;
  const sizes = [15, 14, 13, 12, 11];
  const size = sizes.find(pt => mtgWrap_(full, S.w, pt, 2).length === 1) || sizes[sizes.length - 1];
  const text = mtgWrap_(full, S.w, size, 1)[0];
  const spans = label ? [{start: 0, end: prefix.length, color: c.greenDark}] : [];
  return [
    mtgBox_(S.x, 11, S.w, 28, text, {size, bold: true, color: c.ink, spans, role: 'report-title'}),
    {type: 'box', x: S.x, y: 40, w: S.w, h: 1.2, fill: c.green, shape: 'RECTANGLE', role: 'report-rule'}
  ];
}

/** 枠（四角）と、その左上の見出し（緑の太字）。opt = {fill, line}（line: null なら枠の線なし） */
function mtgReportFrame_(x, y, w, h, heading, opt) {
  const c = MTG_OPTIONS.colors;
  const o = opt || {};
  const line = 'line' in o && o.line === null ? null : o.line || c.line;
  return mtgBox_(x, y, w, h, mtgWrap_(heading, w, 8, 1)[0], {size: 8, bold: true, color: c.greenDark, fill: o.fill || c.white,
    line, lineWeight: 0.75, shape: 'RECTANGLE', valign: 'TOP', role: 'report-frame'});
}

/**
 * 文章の枠の字の大きさ：どの枠にも収まるいちばん大きい大きさ（MTG_OPTIONS.report の maxSize から 0.5pt ずつ minSize まで）。
 * boxes = [{text, w, h: 本文の箱の高さ}]。{size, texts: 箱に入れる文（入らなければ minSize で入る行まで。最後に「…」）, cut: 切ったか}
 */
function mtgReportFit_(boxes) {
  const R = MTG_OPTIONS.report;
  const S = MTG_SLIDE;
  const lines = (b, size) => mtgWrap_(b.text, b.w, size, 999).length;
  const fits = (b, size) => !b.text || mtgTextHeight_(lines(b, size), size) <= b.h;
  for (let size = R.maxSize; size >= R.minSize - 1e-9; size -= 0.5) {
    if (boxes.every(b => fits(b, size))) return {size, texts: boxes.map(b => b.text), cut: false};
  }
  const size = R.minSize;
  let cut = false;
  const texts = boxes.map(b => {
    if (fits(b, size)) return b.text;
    cut = true;
    const max = Math.max(1, Math.floor((b.h - S.padY) / (size * S.lineFactor)));
    return mtgWrap_(b.text, b.w, size, max).join('\n');
  });
  return {size, texts, cut};
}

/** 本文の箱（枠の見出しの下。上下の中央）。lines = [{text, size, bold, color}]（空の行は除く） */
function mtgReportBody_(b, lines, bottomSpace) {
  const L = mtgReportLayout_();
  const list = lines.filter(line => line && line.text);
  const text = list.map(line => line.text).join('\n');
  let at = 0;
  const spans = list.map(line => {
    const span = {start: at, end: at + line.text.length, size: line.size, bold: !!line.bold, color: line.color};
    at += line.text.length + 1;
    return span;
  });
  return mtgBox_(b.x, b.y + L.head, b.w, b.h - L.head - (bottomSpace || 0), text, {size: list.length ? list[0].size : 9,
    color: MTG_OPTIONS.colors.ink, spans, valign: 'MIDDLE', role: 'report-value'});
}

/** 左の列の1つ目：先方（部署は太字・役職は普通の字。名前（先方担当氏名）は出さない） */
function mtgReportClient_(b, item) {
  const c = MTG_OPTIONS.colors;
  const fit = (text, size) => mtgWrap_(text, b.w, size, 2).join('\n');
  const lines = !item.clientDept && !item.clientTitle ? [{text: MTG_OPTIONS.blankText, size: 9.5, color: c.faint}]
    : [{text: item.clientDept ? fit(item.clientDept, 10.5) : '', size: 10.5, bold: true, color: c.ink},
      {text: item.clientTitle ? fit(item.clientTitle, 9.5) : '', size: 9.5, color: c.ink}];
  return [mtgReportFrame_(b.x, b.y, b.w, b.h, '先方'), mtgReportBody_(b, lines)];
}

/** 左の列の2つ目：ステータス（今のステータス・前回・確度の段（創注・D・C・B・A。今は緑・前回は薄い緑）） */
function mtgReportStatus_(b, item) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const blank = MTG_OPTIONS.blankText;
  const prev = item.isNew ? R.newText : item.prevStatus || blank;
  const stepH = 11;
  const el = [mtgReportFrame_(b.x, b.y, b.w, b.h, 'ステータス'),
    mtgReportBody_(b, [{text: mtgWrap_(item.status || blank, b.w, 10.5, 1)[0], size: 10.5, bold: true, color: item.status ? c.ink : c.faint},
      {text: mtgWrap_('前回：' + prev, b.w, 8, 1)[0], size: 8, color: c.muted}], stepH + 4)];
  const now = mtgReportRank_(item.status);
  const before = item.isNew ? '' : mtgReportRank_(item.prevStatus);
  const steps = R.steps;
  const sw = (b.w - 14 - (steps.length - 1) * 2) / steps.length;
  steps.forEach((step, i) => {
    const on = step === now;
    const was = !on && step === before;
    el.push(mtgBox_(b.x + 7 + i * (sw + 2), b.y + b.h - stepH - 5, sw, stepH, step, {size: 7, bold: on, align: 'CENTER', shape: 'RECTANGLE',
      fill: on ? c.green : was ? c.greenSoft : c.graySoft, color: on ? c.white : was ? c.greenDark : c.faint, role: 'report-step'}));
  });
  return el;
}

/** 左の列の3つ目：期待値（百万円。今の値・前回からの増減・前回の値と想定売上） */
function mtgReportValue_(b, item) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const now = item.weighted == null ? MTG_OPTIONS.blankText : mtgAmount_(item.weighted);
  const delta = item.isNew || item.weighted == null || item.prevWeighted == null ? null : item.weighted - item.prevWeighted;
  const rounded = delta == null ? null : Math.round(delta * 10) / 10;
  const first = now + (rounded == null ? '' : '　' + mtgSigned_(rounded));
  const spans = [{start: 0, end: now.length, size: item.weighted == null ? 10.5 : 18, bold: true, color: item.weighted == null ? c.faint : c.ink}];
  if (rounded != null) {
    spans.push({start: now.length, end: first.length, size: 9.5, bold: true, color: rounded > 0 ? c.greenDark : rounded < 0 ? c.dark : c.faint});
  }
  const second = mtgWrap_('前回：' + (item.isNew ? R.newText : mtgAmount_(item.prevWeighted)) + '　／　想定売上：' + mtgAmount_(item.amount), b.w, 8, 1)[0];
  const L = mtgReportLayout_();
  const text = first + '\n' + second;
  spans.push({start: first.length + 1, end: text.length, size: 8, color: c.muted});
  return [mtgReportFrame_(b.x, b.y, b.w, b.h, '期待値（百万円）'),
    mtgBox_(b.x, b.y + L.head, b.w, b.h - L.head, text, {size: 9.5, color: c.ink, spans, valign: 'MIDDLE', role: 'report-value'})];
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
    {text: nowText, size: now ? 13.5 : 10.5, bold: true, color: now || item.orderMonthText ? c.ink : c.faint},
    {text: '前回：' + prevText + (move > 0 ? '（延期）' : move < 0 ? '（前倒し）' : ''), size: 8, color: c.muted}
  ])];
}

/** 左の列の5つ目：RXシナリオ（見出しに見込金額。選んだシナリオを角の丸い枠で並べる。入らない分は「ほか N件」） */
function mtgReportScenario_(b, item, hasColumn) {
  const o = MTG_OPTIONS;
  const c = o.colors;
  const L = mtgReportLayout_();
  const heading = 'RXシナリオ（見込金額 ' + (item.scenarioAmount == null ? o.blankText : mtgAmount_(item.scenarioAmount) + '百万円') + '）';
  const el = [mtgReportFrame_(b.x, b.y, b.w, b.h, heading)];
  const names = item.scenarios || [];
  if (!hasColumn || !names.length) {
    el.push(mtgBox_(b.x, b.y + L.head, b.w, b.h - L.head, hasColumn ? o.blankText : o.gapText, {size: 9.5, color: c.faint, valign: 'MIDDLE', role: 'report-value'}));
    return el;
  }
  const size = 7.5;
  const h = 12;
  const gap = 3;
  const x0 = b.x + 7;
  const right = b.x + b.w - 7;
  const width = text => Math.min(right - x0, Math.ceil(mtgTextWidth_(text, size) / 0.92 + 10));
  const rows = Math.max(1, Math.floor((b.h - L.head - 8 + gap) / (h + gap)));
  const placed = [];
  let x = x0;
  let row = 0;
  for (let i = 0; i < names.length; i++) {
    const w = width(names[i]);
    if (x + w > right && x > x0) {
      row++;
      x = x0;
    }
    if (row >= rows) break;
    placed.push({text: mtgWrap_(names[i], w + MTG_SLIDE.padX - 10, size, 1)[0], x, row, w});
    x += w + gap;
  }
  // 入らない分は、最後の枠の代わりに「ほか N件」
  if (placed.length < names.length) {
    const rest = names.length - placed.length + 1;
    const last = placed.pop();
    if (last) placed.push({text: 'ほか ' + rest + '件', x: last.x, row: last.row, w: width('ほか ' + rest + '件'), more: true});
  }
  const used = (placed.length ? placed[placed.length - 1].row + 1 : 1) * (h + gap) - gap;
  const y0 = b.y + L.head + (b.h - L.head - used) / 2;
  placed.forEach(p => el.push(mtgBox_(p.x, y0 + p.row * (h + gap), p.w, h, p.text, {size, align: 'CENTER', shape: 'ROUND_RECTANGLE',
    fill: c.white, line: p.more ? c.line : c.green, lineWeight: 0.75, color: p.more ? c.muted : c.greenDark, role: 'report-chip'})));
  return el;
}

/** 右下の「商談記録：」（URL ならリンク。無ければ「未入力」。列が無ければ「データなし」）。mtgFrame_ が出典の行の右に置く */
function mtgReportRecord_(record, hasColumn) {
  const c = MTG_OPTIONS.colors;
  const R = MTG_OPTIONS.report;
  const head = '商談記録：';
  const value = String(record || '').trim();
  if (!hasColumn || !value) {
    const text = head + (hasColumn ? MTG_OPTIONS.blankText : MTG_OPTIONS.gapText);
    return {w: R.recordWidth, text, size: 8, color: c.faint};
  }
  const url = (/https?:\/\/[^\s"'<>]+/.exec(value) || [''])[0];
  const shown = mtgWrap_(head + (url || value), R.recordWidth, 8, 1)[0];
  const spans = url ? [{start: head.length, end: shown.length, url, color: c.link, underline: true}] : [];
  return {w: R.recordWidth, text: shown, size: 8, color: c.faint, spans};
}
