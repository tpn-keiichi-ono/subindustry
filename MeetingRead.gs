/**
 * 会議資料の中身を、このスプレッドシートから Sheets API（拡張サービス）で読む（2026/11/04 から）
 * 会議サマリ（MeetingSummary.gs）と同じプロジェクトに置くファイル。接頭辞は mtg（会議サマリ）。
 *
 * なぜ Sheets API か
 * - 大勢が編集している間は、SpreadsheetApp でこのスプレッドシートに触れるたびに数式の計算を待たされ、225行のシートを読むだけで約4分かかった。
 *   Sheets API は同じ時間帯に約30秒だった（diagnoseMeetingRead()。2026/11/04 20時ごろ）。
 * - スプレッドシートをコピーして読む方法は、作ったばかりのコピーを初めて開くと数式の計算し直しで数分かかり（274秒）、
 *   ほかのファイルを読む数式（IMPORTRANGE）が 0 になった（シナリオによる見込金額）。Sheets API なら、元のファイルで計算済みの値をそのまま読む。
 *
 * 仕組み
 * - mtgApiBook_ が、要るシート（追跡シート・ステータスの定義）を1回の呼び出しでまとめて読み（値・表示の文字・数式・書式の種類・入力規則）、
 *   記録のシート（変更履歴_差分・変更時点スナップショット）をもう1回で読む（値だけ）。
 * - 読んだ値を、SpreadsheetApp のスプレッドシート・シート・範囲と同じ使い方ができる入れ物（読むだけ）にして返す。
 *   資料の中身を作る処理（mtgSummaries_ など）は、そのまま使える。読んでいないシートを開こうとしたら、そのときに読む。
 * - Sheets API が使えない（拡張サービスを外した）・読めないときは、前の版と同じく SpreadsheetApp で読む（呼ぶ側が決める）。
 */

/** Sheets API（拡張サービス）が使えるか */
function mtgApiAvailable_() {
  return typeof Sheets !== 'undefined' && !!Sheets.Spreadsheets && typeof Sheets.Spreadsheets.get === 'function' &&
    !!Sheets.Spreadsheets.Values && typeof Sheets.Spreadsheets.Values.batchGet === 'function';
}

/** シート名を API の範囲の書き方に（'シート名'） */
function mtgApiQuote_(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}

/**
 * このスプレッドシートのシートの名前（並びの順）。Sheets API で読む（SpreadsheetApp でシートを探すと、大勢が編集している間は待たされるため）。
 * Sheets API が使えなければ SpreadsheetApp
 */
function mtgSheetTitles_(ss) {
  if (MTG_OPTIONS.run.readMode === 'api' && mtgApiAvailable_()) {
    try {
      const meta = Sheets.Spreadsheets.get(ss.getId(), {fields: 'sheets.properties(title,index)'});
      return (meta.sheets || []).map(s => s.properties).sort((a, b) => (a.index || 0) - (b.index || 0)).map(p => p.title);
    } catch (error) {
      console.warn('Sheets API でシートの一覧を読めませんでした（SpreadsheetApp で読みます）：' + (error && error.message || error));
    }
  }
  return ss.getSheets().map(sheet => sheet.getName());
}

/** このスプレッドシートにある追跡シート（DIFF_RULES の順） */
function mtgTrackedNames_(ss) {
  const titles = mtgSheetTitles_(ss);
  return Object.keys(DIFF_RULES).filter(name => titles.indexOf(name) >= 0);
}

/**
 * このスプレッドシートを Sheets API で読んだ入れ物（読むだけ。SpreadsheetApp のスプレッドシートと同じ使い方）。
 * names の追跡シートとステータスの定義は、値・表示の文字・数式・書式の種類・入力規則まで1回で読む。options.records が false でなければ、
 * 記録のシート（変更履歴_差分・変更時点スナップショット。クレデンシャルを使っていれば クレデンシャル履歴_記録）を値だけもう1回で読む。
 * options.grid が false なら、追跡シートは読まない（記録のシートだけ）。手順ごとの時間を実行ログに書く
 */
function mtgApiBook_(ss, names, options) {
  const opt = options || {};
  const timer = mschTimer_('Sheets API で読む', true);
  const id = ss.getId();
  const meta = Sheets.Spreadsheets.get(id, {fields: 'properties(title,timeZone),sheets.properties(sheetId,title,index,gridProperties(rowCount,columnCount))'});
  const book = {
    id,
    title: (meta.properties || {}).title || '',
    tz: (meta.properties || {}).timeZone || Session.getScriptTimeZone(),
    props: {},
    sheets: {},
    timer
  };
  (meta.sheets || []).forEach(s => { book.props[s.properties.title] = s.properties; });
  timer.lap('シートの一覧');
  const exists = name => !!book.props[name];
  const grid = opt.grid === false ? [] : (names || []).concat([MTG_OPTIONS.statusTable && MTG_OPTIONS.statusTable.sheet])
    .filter((name, i, list) => name && exists(name) && list.indexOf(name) === i);
  if (grid.length) mtgApiLoadGrid_(book, grid);
  if (opt.records !== false) {
    const records = [DIFF_OPTIONS.logSheet, DIFF_OPTIONS.rowSnapshotSheet].concat(credEnabled_() ? [CRED_OPTIONS.logSheet] : []).filter(exists);
    if (records.length) mtgApiLoadValues_(book, records);
  }
  return mtgApiBookObject_(book);
}

/** シートを、値・表示の文字・数式・書式の種類・入力規則まで読む（1回の呼び出し） */
function mtgApiLoadGrid_(book, names) {
  const fields = 'sheets(properties(title),data(startRow,startColumn,rowData(values(formattedValue,effectiveValue,' +
    'userEnteredValue,effectiveFormat(numberFormat(type)),dataValidation(condition)))))';
  const res = Sheets.Spreadsheets.get(book.id, {ranges: names.map(mtgApiQuote_), includeGridData: true, fields});
  (res.sheets || []).forEach(s => {
    const rows = [];
    (s.data || []).forEach(data => {
      const r0 = data.startRow || 0;
      const c0 = data.startColumn || 0;
      (data.rowData || []).forEach((row, i) => {
        const line = rows[r0 + i] = rows[r0 + i] || [];
        (row.values || []).forEach((cell, j) => { line[c0 + j] = mtgApiCell_(cell, book.tz); });
      });
    });
    book.sheets[s.properties.title] = mtgApiTrim_(rows);
  });
  book.timer.lap(names.map(n => '「' + n + '」').join('') + 'を読む');
}

/** 記録のシートを値だけ読む（1回の呼び出し）。見出しに「日時」を含む列の値は日時にする */
function mtgApiLoadValues_(book, names) {
  const res = Sheets.Spreadsheets.Values.batchGet(book.id, {ranges: names.map(mtgApiQuote_), valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER', majorDimension: 'ROWS'});
  (res.valueRanges || []).forEach((range, k) => {
    const values = range.values || [];
    const heads = (values[0] || []).map(h => String(h == null ? '' : h));
    const dates = heads.map(h => /日時/.test(h));
    const rows = values.map((line, r) => line.map((v, c) => {
      let value = v == null ? '' : v;
      if (r > 0 && dates[c] && typeof value === 'number') value = mtgApiDate_(value, book.tz);
      return {v: value, d: mtgApiDisplay_(value, book.tz), f: '', u: true};
    }));
    book.sheets[names[k]] = mtgApiTrim_(rows);
  });
  book.timer.lap(names.map(n => '「' + n + '」').join('') + 'を読む（' +
    names.map(n => Math.max(((book.sheets[n] || {}).lastRow || 1) - 1, 0) + '件').join('・') + '）');
}

/** API のセル → {v: 値（getValues と同じ型）, d: 表示の文字, f: 数式, u: 入力した値・数式があるか（無ければ配列の数式が広げた値）, dv: 入力規則の条件} */
function mtgApiCell_(cell, tz) {
  const ev = cell.effectiveValue || {};
  const type = ((cell.effectiveFormat || {}).numberFormat || {}).type || '';
  let v = '';
  if (ev.numberValue != null) v = /^(DATE|DATE_TIME|TIME)$/.test(type) ? mtgApiDate_(ev.numberValue, tz) : ev.numberValue;
  else if (ev.stringValue != null) v = ev.stringValue;
  else if (ev.boolValue != null) v = ev.boolValue;
  else if (ev.errorValue) v = cell.formattedValue || '#ERROR!';
  return {v, d: cell.formattedValue == null ? '' : String(cell.formattedValue), f: ((cell.userEnteredValue || {}).formulaValue) || '',
    u: !!cell.userEnteredValue, dv: cell.dataValidation ? cell.dataValidation.condition || null : null};
}

/** シリアル値（スプレッドシートのタイムゾーンの日時）→ Date */
function mtgApiDate_(serial, tz) {
  const ms = Math.round((serial - 25569) * 86400000);
  const base = Math.floor(ms / 1000) * 1000;   // 秒までを日時の文字にしてタイムゾーンで読み、ミリ秒を足す
  const text = Utilities.formatDate(new Date(base), 'UTC', 'yyyy/MM/dd HH:mm:ss');
  return new Date(Utilities.parseDate(text, tz, 'yyyy/MM/dd HH:mm:ss').getTime() + (ms - base));
}

/** 値だけ読んだセルの表示の文字（日時は「yyyy/MM/dd HH:mm:ss」、真偽は TRUE/FALSE） */
function mtgApiDisplay_(value, tz) {
  if (value === '' || value == null) return '';
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy/MM/dd HH:mm:ss');
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

/** 最後の行・列（値か数式のあるセルまで）に合わせて、後ろの空の行を除く。{rows, lastRow, lastColumn} */
function mtgApiTrim_(rows) {
  let lastRow = 0;
  let lastColumn = 0;
  rows.forEach((line, r) => {
    (line || []).forEach((cell, c) => {
      if (cell && (cell.f || (cell.v !== '' && cell.v != null))) {
        lastRow = Math.max(lastRow, r + 1);
        lastColumn = Math.max(lastColumn, c + 1);
      }
    });
  });
  return {rows: rows.slice(0, lastRow), lastRow, lastColumn};
}

/** 読んだ値の入れ物を、SpreadsheetApp のスプレッドシートと同じ使い方ができる形に */
function mtgApiBookObject_(book) {
  const cache = {};
  const object = {
    getId: () => book.id,
    getName: () => book.title,
    getUrl: () => 'https://docs.google.com/spreadsheets/d/' + book.id + '/edit',
    getSpreadsheetTimeZone: () => book.tz,
    /** 読んだシート（読んでいないシートは、そのときに読む。無ければ null） */
    getSheetByName: name => {
      if (cache[name]) return cache[name];
      if (!book.props[name]) return null;
      if (!book.sheets[name]) mtgApiLoadGrid_(book, [name]);
      cache[name] = mtgApiSheetObject_(object, book, name);
      return cache[name];
    },
    /** 読んだシートだけ（並びの順） */
    getSheets: () => Object.keys(book.sheets).sort((a, b) => (book.props[a].index || 0) - (book.props[b].index || 0)).map(object.getSheetByName),
    isApiBook: true
  };
  return object;
}

/** シートの入れ物（読むだけ） */
function mtgApiSheetObject_(parent, book, name) {
  const data = book.sheets[name];
  const grid = (book.props[name] || {}).gridProperties || {};
  const sheet = {
    getName: () => name,
    getSheetId: () => (book.props[name] || {}).sheetId,
    getParent: () => parent,
    getLastRow: () => data.lastRow,
    getLastColumn: () => data.lastColumn,
    getMaxRows: () => grid.rowCount || data.lastRow,
    getMaxColumns: () => grid.columnCount || data.lastColumn,
    getDataRange: () => mtgApiRangeObject_(sheet, data, book, 1, 1, Math.max(data.lastRow, 1), Math.max(data.lastColumn, 1)),
    getRange: (a, b, c, d) => {
      if (typeof a === 'string') {
        const r = mtgApiParseA1_(a, data);
        return mtgApiRangeObject_(sheet, data, book, r.row, r.col, r.rows, r.cols);
      }
      return mtgApiRangeObject_(sheet, data, book, a, b, c == null ? 1 : c, d == null ? 1 : d);
    },
    /** 配列の数式が広げた値のある列（0 から）：row 行目から rows 行・width 列のうち、入力した値も数式も無いのに値があるセルがある列 */
    spilled: (row, rows, width) => {
      const out = [];
      for (let c = 0; c < width; c++) {
        for (let r = row; r < row + rows; r++) {
          const x = (data.rows[r - 1] || [])[c];
          if (x && !x.u && !x.f && x.v !== '' && x.v != null) {
            out.push(c);
            break;
          }
        }
      }
      return out;
    }
  };
  return sheet;
}

/** 「A1」「F3:F20」「A:C」「3:5」 → {row, col, rows, cols}（行・列を省いたときは、データのある最後まで） */
function mtgApiParseA1_(a1, data) {
  const col = letters => letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const part = text => {
    const m = /^\$?([A-Z]*)\$?(\d*)$/.exec(String(text).toUpperCase());
    if (!m) throw new Error('範囲「' + a1 + '」を読めません。');
    return {c: m[1] ? col(m[1]) : null, r: m[2] ? Number(m[2]) : null};
  };
  const pieces = String(a1).replace(/^.*!/, '').split(':');
  const p = part(pieces[0]);
  const q = pieces[1] != null ? part(pieces[1]) : p;
  const r1 = p.r || 1;
  const c1 = p.c || 1;
  const r2 = q.r || Math.max(data.lastRow, r1);
  const c2 = q.c || Math.max(data.lastColumn, c1);
  return {row: r1, col: c1, rows: r2 - r1 + 1, cols: c2 - c1 + 1};
}

/** 範囲の入れ物（読むだけ。getValues・getDisplayValues・getFormulas・getDataValidation など） */
function mtgApiRangeObject_(sheet, data, book, row, col, rows, cols) {
  const cell = (r, c) => ((data.rows[r - 1] || [])[c - 1]) || null;
  const map = fn => {
    const out = [];
    for (let r = 0; r < rows; r++) {
      const line = [];
      for (let c = 0; c < cols; c++) line.push(fn(cell(row + r, col + c)));
      out.push(line);
    }
    return out;
  };
  const rule = x => (x && x.dv ? mtgApiRule_(x.dv, sheet.getParent()) : null);
  const range = {
    getSheet: () => sheet,
    getRow: () => row,
    getColumn: () => col,
    getNumRows: () => rows,
    getNumColumns: () => cols,
    getLastRow: () => row + rows - 1,
    getLastColumn: () => col + cols - 1,
    getA1Notation: () => diffColumnLetter_(col) + row + (rows > 1 || cols > 1 ? ':' + diffColumnLetter_(col + cols - 1) + (row + rows - 1) : ''),
    getValues: () => map(x => (x ? x.v : '')),
    getDisplayValues: () => map(x => (x ? x.d : '')),
    getFormulas: () => map(x => (x ? x.f : '')),
    getValue: () => range.getValues()[0][0],
    getDisplayValue: () => range.getDisplayValues()[0][0],
    getFormula: () => range.getFormulas()[0][0],
    getDataValidations: () => map(rule),
    getDataValidation: () => rule(cell(row, col))
  };
  return range;
}

/**
 * 入力規則の条件 → SpreadsheetApp の入力規則と同じ使い方（getCriteriaType・getCriteriaValues）。一覧と範囲（同じスプレッドシートのシート）だけ。
 * 範囲のシートを読んでいなければ、そのときに読む
 */
function mtgApiRule_(condition, parent) {
  const types = SpreadsheetApp.DataValidationCriteria;
  const values = (condition.values || []).map(v => String(v.userEnteredValue == null ? '' : v.userEnteredValue));
  if (condition.type === 'ONE_OF_LIST') {
    return {getCriteriaType: () => types.VALUE_IN_LIST, getCriteriaValues: () => [values, true]};
  }
  if (condition.type === 'ONE_OF_RANGE' && values[0]) {
    const m = /^=?\s*(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(values[0].trim());
    const sheet = m && parent.getSheetByName(m[1] ? m[1].replace(/''/g, "'") : m[2].trim());
    if (!sheet) return null;
    const target = sheet.getRange(m[3].replace(/\$/g, ''));
    return {getCriteriaType: () => types.VALUE_IN_RANGE, getCriteriaValues: () => [target, true]};
  }
  return null;
}

/**
 * 時点コピー（copyId）の追跡シート（names）の、数式で出している値（数式のセルと、配列の数式が広げたセル）を、元のスプレッドシートの値（book。
 * Sheets API で読んだもの）にする（Sheets API で書く）。コピーは新しいファイルなので、ほかのファイルを読む数式（IMPORTRANGE など）は計算し直されて
 * 0・エラーになるため（シナリオによる見込金額）。期待値の列は数式のまま残す（同じファイルの中だけで計算する。数式かどうかで前回の値の取り方を変えるため）。
 * 列ごとに、見出しの下の行をすべて書く（配列の数式は先頭のセルにしか数式が無いので、列ごと値にする）。戻り値 {cells, columns, error}
 */
function mtgApiFixCopy_(copyId, book, names) {
  const out = {cells: 0, columns: [], error: ''};
  try {
    const data = [];
    names.forEach(name => {
      const sheet = book.getSheetByName(name);
      const rule = DIFF_RULES[name];
      if (!sheet || !rule) return;
      const rows = sheet.getLastRow() - rule.headerRow;
      const width = sheet.getLastColumn();
      if (rows < 1 || width < 1) return;
      const headers = sheet.getRange(rule.headerRow, 1, 1, width).getDisplayValues()[0];
      const skip = mtgColumnIndex_(headers, 'weighted');
      const formulas = sheet.getRange(rule.headerRow + 1, 1, rows, width).getFormulas();
      const values = sheet.getRange(rule.headerRow + 1, 1, rows, width).getValues();
      // 数式で出している列：数式のセルがある列と、配列の数式が広げた列（入力した値が無いのに値がある列。API の読み取りで分かる）
      const spilled = mtgApiSpilledColumns_(book, name, rule.headerRow, rows, width);
      for (let c = 0; c < width; c++) {
        if (c === skip) continue;
        if (!formulas.some(line => String(line[c] || '').trim()) && spilled.indexOf(c) < 0) continue;
        const letter = diffColumnLetter_(c + 1);
        data.push({range: mtgApiQuote_(name) + '!' + letter + (rule.headerRow + 1) + ':' + letter + (rule.headerRow + rows), majorDimension: 'COLUMNS',
          values: [values.map(line => mtgApiRaw_(line[c], book.getSpreadsheetTimeZone()))]});
        out.cells += rows;
        out.columns.push(name + ' の' + letter + '列（' + String(headers[c] || '').replace(/\s+/g, '') + '）');
      }
    });
    if (data.length) Sheets.Spreadsheets.Values.batchUpdate({valueInputOption: 'RAW', data}, copyId);
  } catch (error) {
    out.error = '時点コピーの、数式で出している値を元の値にできませんでした（' + (error && error.message || error) + '）';
    console.warn(out.error);
  }
  return out;
}

/** 配列の数式が広げた列（0 から）：数式も入力した値も無いのに値があるセルがある列（Sheets API で読んだシートだけ。ほかは空） */
function mtgApiSpilledColumns_(book, name, headerRow, rows, width) {
  const sheet = book.getSheetByName(name);
  if (!sheet || !sheet.spilled) return [];
  return sheet.spilled(headerRow + 1, rows, width);
}

/** 書く値（日時はシリアル値。セルの書式で日時に見える） */
function mtgApiRaw_(value, tz) {
  if (value instanceof Date) {
    const local = Utilities.formatDate(value, tz, 'yyyy/MM/dd HH:mm:ss').split(/[\/ :]/).map(Number);
    return (Date.UTC(local[0], local[1] - 1, local[2], local[3], local[4], local[5]) + value.getTime() % 1000) / 86400000 + 25569;
  }
  return value == null ? '' : value;
}
