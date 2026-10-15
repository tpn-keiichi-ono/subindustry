/**
 * Apps Script のモック（テスト用）
 *
 * - リポジトリ直下の *.gs を1つのグローバルスコープ（vm のコンテキスト）に読み込む。
 *   Apps Script と同じく、ファイルをまたいで関数・定数を参照でき、同じ名前の const があると読み込みで失敗する。
 * - SpreadsheetApp などは、このプロジェクトが使う範囲だけを再現している。
 *   使う API が無ければここに足すこと（無いメソッドを呼ぶと TypeError になる）。
 *
 * 再現している主な動き
 * - setValue / setValues の文字列は、手入力と同じく変換する：
 *   先頭が ' なら文字列のまま（' は値に含めない）、= なら数式、数字・TRUE/FALSE・日付（2026/10/01）は数値・真偽値・日時。
 * - setRichTextValues は変換しない（文字列のまま入る）。
 * - getRange はシートの大きさを超えるとエラーになる（行・列の追加を忘れたときに気づけるように）。
 * - スクリプトの書き込みは gas.writes に「ロックを持っていたか」と一緒に残る（利用者の入力 gas.edit() は残さない）。
 * - 数式は計算しない（表示は空）。
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const ROOT = path.join(__dirname, '..', '..');   // Apps Script のファイルはリポジトリ直下

const isDate = v => Object.prototype.toString.call(v) === '[object Date]';

function colToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

/** 'A1' / 'A1:B2' / 'D3:AF'（最終行まで）/ 'A:Z'（全行）/ '2:2'（全列） */
function parseA1(text, maxRows, maxCols) {
  const parts = String(text).toUpperCase().replace(/\$/g, '').split(':');
  const parse = p => {
    const m = /^([A-Z]*)(\d*)$/.exec(p);
    if (!m || (!m[1] && !m[2])) throw new Error('Range not found: ' + text);
    return {col: m[1] ? colToNum(m[1]) : 0, row: m[2] ? Number(m[2]) : 0};
  };
  const s = parse(parts[0]);
  const e = parts.length > 1 ? parse(parts[1]) : s;
  if (parts.length === 1 && (!s.col || !s.row)) throw new Error('Range not found: ' + text);
  const row = s.row || 1;
  const col = s.col || 1;
  return {
    row,
    col,
    numRows: (e.row || maxRows) - row + 1,
    numCols: (e.col || maxCols) - col + 1
  };
}

function formatDate(date, timeZone, format) {
  const parts = {};
  new Intl.DateTimeFormat('en-US', {
    timeZone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit'
  }).formatToParts(date).forEach(p => { parts[p.type] = p.value; });
  const map = {yyyy: parts.year, MM: parts.month, M: String(Number(parts.month)), dd: parts.day, d: String(Number(parts.day)),
    HH: parts.hour, mm: parts.minute, ss: parts.second};
  return String(format).replace(/yyyy|MM|M|dd|d|HH|mm|ss/g, t => map[t]);
}

/* ---------------- RichTextValue / TextStyle ---------------- */

function newTextStyle() {
  const style = {};
  const builder = {
    setForegroundColor(v) { style.color = v; return builder; },
    setBold(v) { style.bold = v; return builder; },
    setItalic(v) { style.italic = v; return builder; },
    setStrikethrough(v) { style.strikethrough = v; return builder; },
    setUnderline(v) { style.underline = v; return builder; },
    setFontSize(v) { style.fontSize = v; return builder; },
    build() { return Object.assign({}, style); }
  };
  return builder;
}

function newRichTextValue() {
  let text = '';
  let link = null;
  const runs = [];
  const builder = {
    setText(v) { text = String(v); return builder; },
    /** setLinkUrl(url) で全体、setLinkUrl(start, end, url) で一部にリンク（getLinkUrl は全体のリンクだけ返す） */
    setLinkUrl(a, b, c) { if (b === undefined) link = a; return builder; },
    setTextStyle(a, b, c) {
      if (c === undefined) runs.push({start: 0, end: text.length, style: a});
      else runs.push({start: a, end: b, style: c});
      return builder;
    },
    build() {
      const value = {
        getText: () => text,
        getLinkUrl: () => link,
        runs: runs.slice(),
        /** 位置ごとの最後に指定された書式（テストで「赤の取り消し線」などを確かめる用） */
        styleAt(index) {
          let found = {};
          runs.forEach(r => { if (index >= r.start && index < r.end) found = Object.assign({}, found, r.style); });
          return found;
        }
      };
      return value;
    }
  };
  return builder;
}

/* ---------------- 入力規則（プルダウン・日付） ---------------- */

const DataValidationCriteria = {VALUE_IN_LIST: 'VALUE_IN_LIST', VALUE_IN_RANGE: 'VALUE_IN_RANGE', DATE_IS_VALID_DATE: 'DATE_IS_VALID_DATE'};

function newDataValidation() {
  const rule = {type: null, values: [], showDropdown: true, allowInvalid: true, helpText: ''};
  const builder = {
    requireValueInList(values, showDropdown) {
      if (!Array.isArray(values) || !values.length) throw new Error('The list of values must contain at least one item.');
      rule.type = DataValidationCriteria.VALUE_IN_LIST;
      rule.values = values.map(String);
      rule.showDropdown = showDropdown !== false;
      return builder;
    },
    requireValueInRange(range, showDropdown) {
      rule.type = DataValidationCriteria.VALUE_IN_RANGE;
      rule.values = [range];
      rule.showDropdown = showDropdown !== false;
      return builder;
    },
    requireDate() { rule.type = DataValidationCriteria.DATE_IS_VALID_DATE; rule.values = []; return builder; },
    setAllowInvalid(allow) { rule.allowInvalid = !!allow; return builder; },
    setHelpText(text) { rule.helpText = String(text); return builder; },
    build() {
      if (!rule.type) throw new Error('A data validation rule needs criteria.');
      const built = Object.assign({}, rule, {values: rule.values.slice()});
      return {
        getCriteriaType: () => built.type,
        getCriteriaValues: () => {
          if (built.type === DataValidationCriteria.VALUE_IN_LIST) return [built.values.slice(), built.showDropdown];
          if (built.type === DataValidationCriteria.VALUE_IN_RANGE) return [built.values[0], built.showDropdown];
          return [];
        },
        getAllowInvalid: () => built.allowInvalid,
        getHelpText: () => built.helpText
      };
    }
  };
  return builder;
}

/* ---------------- Range ---------------- */

class Range {
  constructor(sheet, row, col, numRows, numCols) {
    if (!(numRows >= 1)) throw new Error('The number of rows in the range must be at least 1.');
    if (!(numCols >= 1)) throw new Error('The number of columns in the range must be at least 1.');
    if (row < 1 || row + numRows - 1 > sheet.maxRows) throw new Error('Those rows are out of bounds.');
    if (col < 1 || col + numCols - 1 > sheet.maxCols) throw new Error('Those columns are out of bounds.');
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getLastRow() { return this.row + this.numRows - 1; }
  getLastColumn() { return this.col + this.numCols - 1; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getA1Notation() {
    const start = numToCol(this.col) + this.row;
    return this.numRows === 1 && this.numCols === 1
      ? start : start + ':' + numToCol(this.getLastColumn()) + this.getLastRow();
  }

  map(fn) {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) line.push(fn(this.sheet.cell(this.row + r, this.col + c, false), r, c));
      out.push(line);
    }
    return out;
  }

  each(fn) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) fn(this.sheet.cell(this.row + r, this.col + c, true), r, c);
    }
  }

  write(kind) { this.sheet.gas.recordWrite(this, kind); }

  read(kind) { this.sheet.gas.recordRead(this, kind); }

  getValues() { this.read('getValues'); return this.map(cell => (cell ? cell.v : '')); }
  getDisplayValues() {
    this.read('getDisplayValues');
    return this.map(cell => this.sheet.gas.display(cell));
  }
  getFormulas() { this.read('getFormulas'); return this.map(cell => (cell ? cell.f : '')); }
  getRichTextValues() {
    return this.map(cell => (cell && cell.rich) ||
      newRichTextValue().setText(this.sheet.gas.display(cell)).build());
  }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getFormula() { return this.getFormulas()[0][0]; }
  getNote() { const cell = this.sheet.cell(this.row, this.col, false); return (cell && cell.note) || ''; }

  setValue(value) {
    this.checkValidation(() => value);
    this.write('setValue');
    this.each(cell => this.sheet.gas.store(cell, value));
    return this;
  }

  setValues(matrix) {
    this.checkShape(matrix);
    this.checkValidation((r, c) => matrix[r][c]);
    this.write('setValues');
    this.each((cell, r, c) => this.sheet.gas.store(cell, matrix[r][c]));
    return this;
  }

  /**
   * スクリプトの書き込みを、セルの入力規則で確かめる。本物と同じく、無効な値を拒否する一覧・範囲の規則に合わない値は、
   * 規則のヘルプの文で止める（空にするのは止めない）。利用者の入力（gas.edit）は確かめない。
   */
  checkValidation(valueAt) {
    const gas = this.sheet.gas;
    if (gas.userInput) return;
    this.map((cell, r, c) => {
      const rule = cell && cell.validation;
      if (!rule || rule.getAllowInvalid()) return;
      const type = rule.getCriteriaType();
      if (type !== DataValidationCriteria.VALUE_IN_LIST && type !== DataValidationCriteria.VALUE_IN_RANGE) return;
      const probe = {};
      gas.store(probe, valueAt(r, c));
      const text = gas.display(probe);
      if (text === '') return;
      const criteria = rule.getCriteriaValues()[0];
      const allowed = type === DataValidationCriteria.VALUE_IN_LIST
        ? criteria : [].concat(...criteria.map(item => gas.display(item)));
      if (allowed.indexOf(text) < 0) {
        throw new Error(rule.getHelpText() || 'The data you entered in cell ' + numToCol(this.col + c) + (this.row + r) +
          ' violates the data validation rules set on this cell.');
      }
    });
  }

  setFormula(text) { return this.setFormulas(this.map(() => text)); }

  setFormulas(matrix) {
    this.checkShape(matrix);
    this.write('setFormulas');
    this.each((cell, r, c) => {
      const text = String(matrix[r][c] == null ? '' : matrix[r][c]);
      cell.rich = null;
      cell.v = '';
      cell.f = text;
    });
    return this;
  }

  setRichTextValues(matrix) {
    this.checkShape(matrix);
    this.write('setRichTextValues');
    this.each((cell, r, c) => {
      cell.v = matrix[r][c].getText();
      cell.f = '';
      cell.rich = matrix[r][c];
    });
    return this;
  }

  checkShape(matrix) {
    if (!Array.isArray(matrix) || matrix.length !== this.numRows) {
      throw new Error('The number of rows in the data does not match the number of rows in the range. The data has ' +
        (matrix && matrix.length) + ' but the range has ' + this.numRows + '.');
    }
    matrix.forEach(line => {
      if (!Array.isArray(line) || line.length !== this.numCols) {
        throw new Error('The number of columns in the data does not match the number of columns in the range. The data has ' +
          (line && line.length) + ' but the range has ' + this.numCols + '.');
      }
    });
  }

  setNumberFormat(format) { this.each(cell => { cell.nf = format; }); return this; }
  setNote(note) { this.each(cell => { cell.note = String(note); }); return this; }
  getNotes() { return this.map(cell => (cell && cell.note) || ''); }
  setNotes(matrix) {
    this.checkShape(matrix);
    this.write('setNotes');
    this.each((cell, r, c) => { cell.note = String(matrix[r][c] == null ? '' : matrix[r][c]); });
    return this;
  }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setVerticalAlignment() { return this; }
  setWrap() { return this; }

  /** 書式（まとめて付ける形）。セルの fmt に残す（テストと見た目の確認用） */
  setFormats(key, matrix) {
    this.checkShape(matrix);
    this.each((cell, r, c) => { cell.fmt = Object.assign(cell.fmt || {}, {[key]: matrix[r][c]}); });
    return this;
  }
  setBackgrounds(matrix) { return this.setFormats('bg', matrix); }
  setFontColors(matrix) { return this.setFormats('fc', matrix); }
  setFontWeights(matrix) { return this.setFormats('fw', matrix); }
  setFontSizes(matrix) { return this.setFormats('fs', matrix); }
  setHorizontalAlignments(matrix) { return this.setFormats('ha', matrix); }
  setWrapStrategies(matrix) { return this.setFormats('wrap', matrix); }
  setBorder(top, left, bottom, right, vertical, horizontal, color, style) {
    this.sheet.borders.push({a1: this.getA1Notation(), top, left, bottom, right, color, style});
    return this;
  }
  getFormat(key) { const cell = this.sheet.cell(this.row, this.col, false); return cell && cell.fmt ? cell.fmt[key] : undefined; }
  merge() { this.sheet.merges.push(this.getA1Notation()); return this; }

  insertCheckboxes() {
    this.write('insertCheckboxes');
    this.each(cell => {
      cell.checkbox = true;
      if (cell.v === '' && !cell.f) cell.v = false;
    });
    return this;
  }

  clearContent() {
    this.write('clearContent');
    this.each(cell => { cell.v = ''; cell.f = ''; cell.rich = null; });
    return this;
  }

  clearDataValidations() { this.each(cell => { cell.checkbox = false; cell.validation = null; }); return this; }

  setDataValidation(rule) {
    this.write('setDataValidation');
    this.each(cell => { cell.validation = rule || null; });
    return this;
  }

  setDataValidations(matrix) {
    this.checkShape(matrix);
    this.write('setDataValidations');
    this.each((cell, r, c) => { cell.validation = matrix[r][c] || null; });
    return this;
  }

  getDataValidation() { return this.getDataValidations()[0][0]; }
  getDataValidations() { return this.map(cell => (cell && cell.validation) || null); }
  getMergedRanges() { return []; }

  protect() {
    const protection = new Protection(this);
    this.sheet.protections.push(protection);
    return protection;
  }
}

class Protection {
  constructor(range, type) {
    this.range = range;
    this.type = type || 'RANGE';
    this.description = '';
    this.editors = [];
    this.domainEdit = true;
    this.removed = false;
  }
  setDescription(text) { this.description = text; return this; }
  getDescription() { return this.description; }
  getProtectionType() { return this.type; }
  getRange() { return this.range; }
  setWarningOnly(warningOnly) { this.warningOnly = !!warningOnly; return this; }
  isWarningOnly() { return !!this.warningOnly; }
  addEditor(user) { this.editors.push(typeof user === 'string' ? user : user.getEmail()); return this; }
  removeEditors(users) {
    const emails = users.map(u => (typeof u === 'string' ? u : u.getEmail()));
    this.editors = this.editors.filter(e => emails.indexOf(e) < 0);
    return this;
  }
  getEditors() { return this.editors.map(email => ({getEmail: () => email})); }
  canDomainEdit() { return this.domainEdit; }
  setDomainEdit(value) { this.domainEdit = value; return this; }
  remove() { this.removed = true; }
}

/* ---------------- Sheet ---------------- */

class Sheet {
  constructor(gas, ss, name, id, maxRows, maxCols) {
    this.gas = gas;
    this.ss = ss;
    this.name = name;
    this.id = id;
    this.maxRows = maxRows;
    this.maxCols = maxCols;
    this.data = [];          // data[行-1][列-1] = {v, f, rich, nf, checkbox, note, fmt}
    this.widths = {};        // 列の幅（setColumnWidth で付けたもの）
    this.heights = {};       // 行の高さ（setRowHeight(s) で付けたもの）
    this.borders = [];       // setBorder の記録
    this.merges = [];        // merge したセルの範囲
    this.hidden = false;
    this.frozenRows = 0;
    this.protections = [];
  }

  cell(row, col, create) {
    let line = this.data[row - 1];
    if (!line) {
      if (!create) return null;
      line = this.data[row - 1] = [];
    }
    let cell = line[col - 1];
    if (!cell && create) cell = line[col - 1] = {v: '', f: ''};
    return cell || null;
  }

  getName() { return this.name; }
  setName(name) { this.name = name; return this; }
  getSheetId() { return this.id; }
  getParent() { return this.ss; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }

  getLastRow() {
    for (let r = this.data.length; r >= 1; r--) {
      if ((this.data[r - 1] || []).some(hasContent)) return r;
    }
    return 0;
  }

  getLastColumn() {
    let last = 0;
    this.data.forEach(line => (line || []).forEach((cell, i) => { if (hasContent(cell)) last = Math.max(last, i + 1); }));
    return last;
  }

  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const p = parseA1(a, this.maxRows, this.maxCols);
      return new Range(this, p.row, p.col, p.numRows, p.numCols);
    }
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }

  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }

  insertRowsAfter(after, count) {
    if (this.data.length > after) this.data.splice(after, 0, ...new Array(count));
    this.maxRows += count;
    return this;
  }
  insertRowAfter(after) { return this.insertRowsAfter(after, 1); }
  insertRowsBefore(before, count) { return this.insertRowsAfter(before - 1, count); }

  insertColumnsAfter(after, count) {
    this.data.forEach(line => { if (line && line.length > after) line.splice(after, 0, ...new Array(count)); });
    this.maxCols += count;
    return this;
  }
  insertColumnAfter(after) { return this.insertColumnsAfter(after, 1); }
  insertColumnsBefore(before, count) { return this.insertColumnsAfter(before - 1, count); }
  insertColumnBefore(before) { return this.insertColumnsAfter(before - 1, 1); }

  deleteRows(row, count) {
    // Apps Script と同じく、固定行以外をすべて消すことはできない
    if (this.maxRows - count <= this.frozenRows) throw new Error('Sorry, it is not possible to delete all non-frozen rows.');
    if (this.data.length >= row) this.data.splice(row - 1, count);
    this.maxRows -= count;
    return this;
  }
  deleteRow(row) { return this.deleteRows(row, 1); }

  deleteColumns(col, count) {
    this.data.forEach(line => { if (line && line.length >= col) line.splice(col - 1, count); });
    this.maxCols -= count;
    return this;
  }
  deleteColumn(col) { return this.deleteColumns(col, 1); }

  /** destinationIndex は動かす前の列番号で数える（Apps Script と同じ）。 */
  moveColumns(columnSpec, destinationIndex) {
    const from = columnSpec.getColumn();
    const count = columnSpec.getNumColumns();
    const at = destinationIndex > from ? destinationIndex - 1 - count : destinationIndex - 1;
    this.data.forEach((line, i) => {
      if (!line) return;
      while (line.length < from - 1 + count) line.push(undefined);
      const moved = line.splice(from - 1, count);
      while (line.length < at) line.push(undefined);
      line.splice(at, 0, ...moved);
      this.data[i] = line;
    });
  }

  appendRow(values) {
    const row = this.getLastRow() + 1;
    if (row > this.maxRows) this.maxRows = row;
    if (values.length > this.maxCols) this.maxCols = values.length;
    this.getRange(row, 1, 1, values.length).setValues([values]);
    return this;
  }

  clearContents() {
    this.gas.recordWrite(this.getRange(1, 1, this.maxRows, this.maxCols), 'clearContents');
    this.data.forEach(line => (line || []).forEach(cell => {
      if (cell) { cell.v = ''; cell.f = ''; cell.rich = null; }
    }));
    return this;
  }

  hideSheet() { this.hidden = true; return this; }
  activate() { this.gas.selection = {sheet: this, range: this.getRange('A1')}; return this; }
  isSheetHidden() { return this.hidden; }
  setFrozenRows(n) { this.frozenRows = n; return this; }
  getFrozenRows() { return this.frozenRows; }
  setColumnWidth(col, width) { this.widths[col] = width; return this; }
  setColumnWidths(start, count, width) {
    for (let i = 0; i < count; i++) this.widths[start + i] = width;
    return this;
  }
  setRowHeight(row, height) { this.heights[row] = height; return this; }
  setRowHeights(start, count, height) {
    for (let i = 0; i < count; i++) this.heights[start + i] = height;
    return this;
  }
  /** type（RANGE・SHEET）を渡すとその種類だけ（本物と同じ）。 */
  getProtections(type) { return this.protections.filter(p => !p.removed && (!type || p.getProtectionType() === type)); }
  protect() {
    const protection = new Protection(this.getRange(1, 1, this.maxRows, this.maxCols), 'SHEET');
    this.protections.push(protection);
    return protection;
  }

  getActiveRange() {
    const sel = this.gas.selection;
    return sel && sel.sheet === this ? sel.range : null;
  }
  getCurrentCell() {
    const range = this.getActiveRange();
    return range ? this.getRange(range.getRow(), range.getColumn()) : null;
  }
}

function hasContent(cell) {
  return !!cell && ((cell.v !== '' && cell.v != null) || !!cell.f);
}

/* ---------------- Spreadsheet ---------------- */

class Spreadsheet {
  constructor(gas, options) {
    this.gas = gas;
    this.id = options.id;
    this.name = options.name;
    this.timeZone = options.timeZone;
    this.sheets = [];
    this.nextSheetId = 1001;
  }
  getId() { return this.id; }
  getName() { return this.name; }
  getSpreadsheetTimeZone() { return this.timeZone; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }

  insertSheet(name, options) {
    name = name || 'シート' + (this.sheets.length + 1);
    if (this.getSheetByName(name)) {
      throw new Error('A sheet with the name "' + name + '" already exists. Please enter another name.');
    }
    const o = options || {};
    const sheet = new Sheet(this.gas, this, name, this.nextSheetId++, o.rows || 1000, o.columns || 26);
    this.sheets.push(sheet);
    // 本物と同じく、作ったシートを開く（何も選んでいないテスト・別のスプレッドシートでは選ばない）
    if (this.gas.selection && this === this.gas.ss) sheet.activate();
    return sheet;
  }

  deleteSheet(sheet) { this.sheets = this.sheets.filter(s => s !== sheet); }
  getActiveSheet() { return (this.gas.selection && this.gas.selection.sheet) || this.sheets[0] || null; }
  toast(message, title) { this.gas.toasts.push({message: String(message), title: title == null ? '' : String(title)}); }
}

/* ---------------- 全体 ---------------- */

/** テストで使う追跡シート（tests/lib/fixture.js の 新FMT・新FMT2 と同じ形）。本物の DIFF_RULES の代わりに使う。 */
const TEST_DIFF_RULES = {
  '新FMT': {headerRow: 2, ranges: ['D3:AF'], requests: true},
  '新FMT2': {headerRow: 2, ranges: ['D3:AF']}
};

/**
 * テスト1件ぶんの Apps Script 環境を作り、リポジトリ直下の *.gs を読み込む。
 * options = {user, owner, timeZone, files, realSheetNames}
 */
function createGas(options) {
  options = options || {};
  const gas = {
    activeUser: options.user === undefined ? 'user@example.com' : options.user,
    effectiveUser: options.owner || 'owner@example.com',
    lockBusy: false,          // true にすると、ほかの実行がロックを持っている状態になる
    confirmAnswer: 'YES',     // ui.alert で「はい／いいえ」を聞かれたときの答え
    lockHolders: 0,
    userInput: 0,
    writes: [],
    reads: [],                // 値の読み取り（getValues・getDisplayValues・getFormulas）。どのシートをどれだけ読んだか
    toasts: [],
    menus: [],
    dialogs: [],
    sidebars: [],
    alerts: [],
    logs: [],
    triggers: [],
    scopeChecks: [],          // ScriptApp.requireAllScopes に渡された AuthMode
    authRequired: false,      // true にすると、getAuthorizationInfo が「承認が足りない」を返す
    driveCopies: [],
    files: {},                // DriveApp のファイルの種類 {ID: {mimeType}}（無ければスプレッドシート）
    slideTemplates: {},       // コピーすると Slides の資料になる Google スライド {ID: {pageSize, slides: [{objectId}]}}
    layoutPlaceholders: [],   // 白紙のスライドに付くレイアウトのプレースホルダの種類（例 ['SLIDE_NUMBER']）
    created: [],              // SpreadsheetApp.create で作ったスプレッドシート
    prompts: [],
    promptAnswer: '',
    promptAnswers: [],        // 続けて聞かれるときの答え（先頭から使う。空になったら promptAnswer）
    promptButton: 'OK',
    fetches: [],              // UrlFetchApp.fetch の記録
    fetchStatus: 200,
    fetchStatuses: [],        // 続けて呼ばれるときの応答のコード（先頭から使う。空になったら fetchStatus）
    fetchText: '{}',          // 応答の本文
    driveFiles: [],           // フォルダに作ったファイル
    driveFolders: [],         // 作ったフォルダ
    trashed: [],              // ゴミ箱に移したファイルの ID
    decks: [],                // Slides（拡張サービス）で作ったスライド {id, title, pageSize, requests, batches: batchUpdate ごとのリクエストの数}
    moved: [],                // moveTo で動かしたファイル {id, folder}
    selection: null,
    props: {document: new Map(), user: new Map(), script: new Map()},
    cache: new Map()
  };

  const ss = new Spreadsheet(gas, {
    id: 'spreadsheet-id',
    name: 'シナリオ攻略先リスト',
    timeZone: options.timeZone || 'Asia/Tokyo'
  });
  gas.ss = ss;

  /* ----- 値の扱い ----- */

  gas.store = (cell, value) => {
    cell.rich = null;
    cell.f = '';
    if (value == null || value === '') { cell.v = ''; return; }
    if (typeof value !== 'string') { cell.v = value; return; }
    if (value[0] === "'") { cell.v = value.slice(1); return; }
    if (value[0] === '=') { cell.v = ''; cell.f = value; return; }
    if (/^-?\d+(\.\d+)?$/.test(value)) { cell.v = Number(value); return; }
    if (/^(true|false)$/i.test(value)) { cell.v = /^true$/i.test(value); return; }
    const m = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/.exec(value);
    if (m) { cell.v = new gas.Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); return; }
    cell.v = value;
  };

  gas.display = cell => {
    const v = cell ? cell.v : '';
    if (v === '' || v == null) return '';
    if (isDate(v)) {
      if (cell.nf && /[yd]/.test(cell.nf)) return formatDate(v, ss.timeZone, cell.nf);
      return formatDate(v, ss.timeZone, formatDate(v, ss.timeZone, 'HH:mm:ss') === '00:00:00'
        ? 'yyyy/MM/dd' : 'yyyy/MM/dd HH:mm:ss');
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  };

  gas.recordRead = (range, kind) => {
    gas.reads.push({
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      kind,
      cells: range.getNumRows() * range.getNumColumns()
    });
  };

  gas.recordWrite = (range, kind) => {
    if (gas.userInput) return;
    gas.writes.push({
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      kind,
      locked: gas.lockHolders > 0,
      ss: range.getSheet().getParent().getId()   // 書いたスプレッドシート（SpreadsheetApp.create で作ったものと見分ける）
    });
  };

  /* ----- サービス ----- */

  const ui = {
    ButtonSet: {OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL'},
    Button: {OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE'},
    createMenu(title) {
      const menu = {title, items: []};
      const builder = {
        addItem(label, fn) { menu.items.push({label, fn}); return builder; },
        addSeparator() { menu.items.push({separator: true}); return builder; },
        addSubMenu(sub) { menu.items.push({subMenu: sub}); return builder; },
        addToUi() { gas.menus.push(menu); }
      };
      return builder;
    },
    showModalDialog(html, title) { gas.dialogs.push({html, title}); },
    /** 入力を聞かれたときは gas.promptAnswers の先頭（無ければ gas.promptAnswer。既定は空）と gas.promptButton（既定 'OK'）を返す */
    prompt(title, message, buttons) {
      gas.prompts.push({title, message, buttons});
      const answer = gas.promptAnswers.length ? gas.promptAnswers.shift() : gas.promptAnswer;
      return {getResponseText: () => answer, getSelectedButton: () => gas.promptButton};
    },
    showModelessDialog(html, title) { gas.dialogs.push({html, title, modeless: true}); },
    showSidebar(html) { gas.sidebars.push({html}); },
    /** YES_NO などで聞かれたときは gas.confirmAnswer（既定 'YES'）を返す */
    alert(a, b, c) {
      gas.alerts.push(b === undefined ? {message: a} : {title: a, message: b, buttons: c});
      return c && c !== 'OK' ? gas.confirmAnswer : 'OK';
    }
  };

  let createdSeq = 0;
  const SpreadsheetApp = {
    ProtectionType: {RANGE: 'RANGE', SHEET: 'SHEET'},
    WrapStrategy: {WRAP: 'WRAP', OVERFLOW: 'OVERFLOW', CLIP: 'CLIP'},
    BorderStyle: {SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK', DOTTED: 'DOTTED', DASHED: 'DASHED'},
    /** 新しいスプレッドシート（gas.created に残す。最初のシートは「シート1」） */
    create(name) {
      const created = new Spreadsheet(gas, {id: 'created-' + (++createdSeq), name: String(name), timeZone: ss.timeZone});
      created.insertSheet('シート1');
      gas.created.push(created);
      return created;
    },
    DataValidationCriteria,
    newDataValidation,
    getActiveSpreadsheet: () => ss,
    getActive: () => ss,
    getActiveSheet: () => ss.getActiveSheet(),
    getUi: () => ui,
    flush() {},
    newRichTextValue,
    newTextStyle
  };

  class HtmlOutput {
    constructor(file, data) { this.file = file; this.data = data; this.width = null; this.height = null; this.title = ''; }
    setWidth(w) { this.width = w; return this; }
    setHeight(h) { this.height = h; return this; }
    setTitle(t) { this.title = t; return this; }
    getContent() { return fs.readFileSync(path.join(ROOT, this.file + '.html'), 'utf8'); }
  }

  const htmlExists = name => {
    if (!fs.existsSync(path.join(ROOT, name + '.html'))) throw new Error('No HTML file named ' + name + ' was found.');
  };

  const HtmlService = {
    createTemplateFromFile(name) {
      htmlExists(name);
      const template = {
        evaluate() {
          const data = {};
          Object.keys(template).forEach(k => { if (typeof template[k] !== 'function') data[k] = template[k]; });
          return new HtmlOutput(name, data);
        }
      };
      return template;
    },
    createHtmlOutputFromFile(name) {
      htmlExists(name);
      return new HtmlOutput(name, {});
    }
  };

  const store = map => ({
    getProperty: key => (map.has(key) ? map.get(key) : null),
    setProperty(key, value) { map.set(key, String(value)); return this; },
    deleteProperty(key) { map.delete(key); return this; },
    getProperties: () => Object.fromEntries(map),
    getKeys: () => Array.from(map.keys())
  });
  const PropertiesService = {
    getDocumentProperties: () => store(gas.props.document),
    getUserProperties: () => store(gas.props.user),
    getScriptProperties: () => store(gas.props.script)
  };

  /** gas.lockBusy が true ならどのロックも、'document' / 'script' / 'user' ならその種類のロックだけが、ほかの処理に取られている */
  class Lock {
    constructor(kind) { this.kind = kind; this.mine = false; }
    tryLock() {
      if (this.mine) return true;
      if (gas.lockBusy === true || gas.lockBusy === this.kind) return false;
      this.mine = true;
      gas.lockHolders++;
      return true;
    }
    waitLock(ms) {
      if (!this.tryLock(ms)) throw new Error('Lock timeout: another process was holding the lock for too long.');
    }
    releaseLock() {
      if (this.mine) { this.mine = false; gas.lockHolders--; }
    }
    hasLock() { return this.mine; }
  }
  const LockService = {
    getDocumentLock: () => new Lock('document'),
    getScriptLock: () => new Lock('script'),
    getUserLock: () => new Lock('user')
  };

  const Utilities = {
    formatDate,
    getUuid: () => crypto.randomUUID(),
    parseDate(text, timeZone, format) {
      const tokens = [];
      const pattern = String(format).replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/yyyy|MM|dd|HH|mm|ss/g, t => {
        tokens.push(t);
        return '(\\d{1,4})';
      });
      const m = new RegExp('^' + pattern + '$').exec(String(text));
      if (!m) throw new Error('Unparseable date: "' + text + '"');
      const p = {yyyy: 1970, MM: 1, dd: 1, HH: 0, mm: 0, ss: 0};
      tokens.forEach((t, i) => { p[t] = Number(m[i + 1]); });
      const utc = Date.UTC(p.yyyy, p.MM - 1, p.dd, p.HH, p.mm, p.ss);
      const shown = formatDate(new Date(utc), timeZone, 'yyyy-MM-dd-HH-mm-ss').split('-').map(Number);
      const shownUtc = Date.UTC(shown[0], shown[1] - 1, shown[2], shown[3], shown[4], shown[5]);
      return new gas.Date(utc - (shownUtc - utc));
    },
    sleep() {}
  };

  const Session = {
    getActiveUser: () => ({getEmail: () => gas.activeUser}),
    getEffectiveUser: () => ({getEmail: () => gas.effectiveUser}),
    getScriptTimeZone: () => ss.timeZone
  };

  class Trigger {
    constructor(handler, eventType, source, sourceId) {
      this.handler = handler;
      this.eventType = eventType;
      this.source = source;
      this.sourceId = sourceId;
      this.uid = crypto.randomUUID();
    }
    getHandlerFunction() { return this.handler; }
    getEventType() { return this.eventType; }
    getTriggerSource() { return this.source; }
    getTriggerSourceId() { return this.sourceId; }
    getUniqueId() { return this.uid; }
  }

  const ScriptApp = {
    EventType: {CLOCK: 'CLOCK', ON_EDIT: 'ON_EDIT', ON_CHANGE: 'ON_CHANGE', ON_OPEN: 'ON_OPEN', ON_FORM_SUBMIT: 'ON_FORM_SUBMIT'},
    TriggerSource: {SPREADSHEETS: 'SPREADSHEETS', CLOCK: 'CLOCK'},
    AuthMode: {NONE: 'NONE', CUSTOM_FUNCTION: 'CUSTOM_FUNCTION', LIMITED: 'LIMITED', FULL: 'FULL'},
    // 本物は、承認が足りなければ実行を止めて承認の画面を出す。モックは呼ばれたことだけを残す
    requireAllScopes(authMode) { gas.scopeChecks.push(authMode); },
    getOAuthToken: () => 'oauth-token',
    AuthorizationStatus: {REQUIRED: 'REQUIRED', NOT_REQUIRED: 'NOT_REQUIRED'},
    getAuthorizationInfo: () => ({
      getAuthorizationStatus: () => (gas.authRequired ? 'REQUIRED' : 'NOT_REQUIRED'),
      getAuthorizationUrl: () => (gas.authRequired ? 'https://script.google.com/macros/d/script-id/authorize' : null)
    }),
    getProjectTriggers: () => gas.triggers.slice(),
    deleteTrigger(trigger) { gas.triggers = gas.triggers.filter(t => t !== trigger); },
    newTrigger(handler) {
      const make = (eventType, source, sourceId, options) => ({
        create() {
          const trigger = new Trigger(handler, eventType, source, sourceId);
          trigger.options = options || {};   // 時間主導の設定（everyDays・atHour・inTimezone など）。テストで確かめる用
          gas.triggers.push(trigger);
          return trigger;
        }
      });
      return {
        forSpreadsheet(target) {
          const id = typeof target === 'string' ? target : target.getId();
          return {
            onEdit: () => make('ON_EDIT', 'SPREADSHEETS', id),
            onChange: () => make('ON_CHANGE', 'SPREADSHEETS', id),
            onOpen: () => make('ON_OPEN', 'SPREADSHEETS', id),
            onFormSubmit: () => make('ON_FORM_SUBMIT', 'SPREADSHEETS', id)
          };
        },
        timeBased() {
          const options = {};
          const set = key => value => { options[key] = value; return builder; };
          const builder = Object.assign(make('CLOCK', 'CLOCK', null, options), {
            everyMinutes: set('everyMinutes'),
            everyHours: set('everyHours'),
            everyDays: set('everyDays'),
            atHour: set('atHour'),
            nearMinute: set('nearMinute'),
            inTimezone: set('inTimezone')
          });
          return builder;
        }
      };
    }
  };

  const cacheStore = () => ({
    get: key => (gas.cache.has(key) ? gas.cache.get(key) : null),
    put(key, value) { gas.cache.set(key, String(value)); },
    remove(key) { gas.cache.delete(key); }
  });
  const CacheService = {getUserCache: cacheStore, getDocumentCache: cacheStore, getScriptCache: cacheStore};

  let copySeq = 0;
  let folderSeq = 0;
  const folderOf = (id, name) => ({
    getId: () => id,
    getName: () => name || id,
    getUrl: () => 'https://drive.google.com/drive/folders/' + id,
    isTrashed: () => false,
    createFolder(childName) {
      const child = folderOf('folder-' + (++folderSeq), childName);
      gas.driveFolders.push({id: child.getId(), name: childName, parent: id});
      return child;
    },
    createFile(blob) {
      const fileId = 'file-' + (gas.driveFiles.length + 1);
      gas.driveFiles.push({id: fileId, folder: id, name: blob.getName(), blob});
      return {getId: () => fileId, getName: () => blob.getName(), getUrl: () => 'https://drive.google.com/file/d/' + fileId + '/view'};
    },
    /** フォルダの中の、その名前のファイル（作った PDF と、フォルダに置いた Google スライド。ゴミ箱のものは除く） */
    getFilesByName(fileName) {
      const files = gas.driveFiles.filter(f => f.folder === id && f.name === fileName).map(f => f.id)
        .concat(gas.decks.filter(d => d.folder === id && d.title === fileName).map(d => d.id))
        .filter(fileId => gas.trashed.indexOf(fileId) < 0)
        .map(fileId => ({getId: () => fileId, getName: () => fileName, setTrashed(flag) { if (flag) gas.trashed.push(fileId); return this; }}));
      return {hasNext: () => files.length > 0, next: () => files.shift()};
    }
  });
  const DriveApp = {
    getRootFolder: () => folderOf('root', 'マイドライブ'),
    getFileById(id) {
      return {
        getId: () => id,
        getParents() {
          const parents = [folderOf('parent-of-' + id, 'スプレッドシートのフォルダ')];
          return {hasNext: () => parents.length > 0, next: () => parents.shift()};
        },
        setTrashed(flag) { if (flag) gas.trashed.push(id); return this; },
        isTrashed: () => gas.trashed.indexOf(id) >= 0,
        getName: () => (gas.driveFiles.find(f => f.id === id) || {name: id}).name,
        getMimeType: () => (gas.files[id] || {}).mimeType ||
          (gas.slideTemplates[id] ? 'application/vnd.google-apps.presentation' : 'application/vnd.google-apps.spreadsheet'),
        getUrl: () => 'https://docs.google.com/presentation/d/' + id + '/edit',
        moveTo(folder) {
          gas.moved.push({id, folder: folder.getId()});
          const deck = gas.decks.find(d => d.id === id);
          if (deck) deck.folder = folder.getId();
          return this;
        },
        /** PDF などに書き出した中身（名前を付けられる入れ物だけ） */
        getAs(mimeType) {
          return {name: '', from: id, setName(n) { this.name = String(n); return this; }, getName() { return this.name; },
            getContentType: () => mimeType};
        },
        makeCopy(name, folder) {
          const copyId = 'copy-' + (++copySeq);
          gas.driveCopies.push({from: id, name, folder: folder ? folder.getId() : null, locked: gas.lockHolders > 0});
          const template = gas.slideTemplates[id];
          if (template) {
            // Google スライドのコピー：Slides（拡張サービス）で開ける資料にする
            gas.decks.push({id: copyId, title: name, template: id, folder: folder ? folder.getId() : null, pageSize: template.pageSize,
              slides: template.slides.slice(), requests: [], batches: []});
            return DriveApp.getFileById(copyId);
          }
          const source = gas.driveFiles.find(f => f.id === id);
          if (source) {
            // フォルダに作ったファイル（PDF）のコピー：同じフォルダの一覧（getFilesByName）にも入れる
            gas.driveFiles.push({id: copyId, folder: folder ? folder.getId() : null, name, blob: source.blob, copyOf: id});
            return {getId: () => copyId, getName: () => name, getUrl: () => 'https://drive.google.com/file/d/' + copyId + '/view'};
          }
          return {
            getId: () => copyId,
            getName: () => name,
            getUrl: () => 'https://docs.google.com/spreadsheets/d/' + copyId + '/edit'
          };
        }
      };
    },
    getFolderById: id => folderOf(id)
  };

  /** Google Slides API（拡張サービス）。作ったスライドと batchUpdate のリクエストを gas.decks に残す */
  let deckSeq = 0;
  const Slides = {
    Presentations: {
      create(resource) {
        const id = 'deck-' + (++deckSeq);
        gas.decks.push({id, title: resource.title, pageSize: resource.pageSize, slides: [{objectId: 'p'}], requests: [], batches: []});
        return {presentationId: id, pageSize: resource.pageSize, slides: [{objectId: 'p'}]};
      },
      /** 資料（作ったもの・ひな形のコピー）の大きさとスライドの ID */
      get(id) {
        const deck = gas.decks.find(d => d.id === id) || (gas.slideTemplates[id] && Object.assign({id}, gas.slideTemplates[id]));
        if (!deck) throw new Error('Requested entity was not found.');
        return {presentationId: id, pageSize: deck.pageSize, slides: deck.slides.map(sl => ({objectId: sl.objectId, pageElements: (sl.pageElements || []).slice()}))};
      },
      batchUpdate(resource, id) {
        const deck = gas.decks.find(d => d.id === id);
        if (!deck) throw new Error('Requested entity was not found.');
        if (gas.slidesError) throw new Error(gas.slidesError);
        deck.batches.push(resource.requests.length);
        deck.requests.push(...resource.requests);
        // 白紙のスライドには、レイアウトのプレースホルダ（gas.layoutPlaceholders の種類）が付く。消したものは除く
        resource.requests.forEach(q => {
          if (q.createSlide) deck.slides.push({objectId: q.createSlide.objectId,
            pageElements: (gas.layoutPlaceholders || []).map((type, n) => ({objectId: q.createSlide.objectId + '_ph' + n, shape: {placeholder: {type}}}))});
          if (q.deleteObject) {
            const gone = q.deleteObject.objectId;
            deck.slides = deck.slides.filter(sl => sl.objectId !== gone)
              .map(sl => (sl.pageElements ? Object.assign({}, sl, {pageElements: sl.pageElements.filter(el => el.objectId !== gone)}) : sl));
          }
        });
        return {replies: []};
      }
    }
  };

  /** 外への送信（Google チャットの webhook など）。gas.fetchStatuses の先頭（無ければ gas.fetchStatus）と gas.fetchText を返す */
  const UrlFetchApp = {
    fetch(url, params) {
      gas.fetches.push({url, params});
      const code = gas.fetchStatuses.length ? gas.fetchStatuses.shift() : gas.fetchStatus;
      const blob = {name: '', setName(n) { this.name = String(n); return this; }, getName() { return this.name; },
        getContentType: () => 'application/pdf'};
      return {getResponseCode: () => code, getContentText: () => gas.fetchText, getBlob: () => blob};
    }
  };

  const Browser = {msgBox(message) { gas.alerts.push({message: String(message)}); return 'ok'; }};

  const capture = level => (...args) => {
    const line = args.map(a => (typeof a === 'string' ? a : (a && a.stack) || String(a))).join(' ');
    gas.logs.push({level, line});
    if (process.env.GAS_VERBOSE) console.log('[' + level + '] ' + line);
  };
  const logger = {log: capture('log'), info: capture('info'), warn: capture('warn'), error: capture('error')};

  /* ----- プロジェクトの読み込み ----- */

  // 拡張サービス（Slides）は、有効にしていないプロジェクトを試すときは options.noSlides で外す
  const MimeType = {GOOGLE_SLIDES: 'application/vnd.google-apps.presentation', GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
    PDF: 'application/pdf'};
  const globals = {
    SpreadsheetApp, HtmlService, PropertiesService, LockService, Utilities, Session,
    ScriptApp, CacheService, DriveApp, MimeType, UrlFetchApp, Browser, console: logger
  };
  if (!options.noSlides) globals.Slides = Slides;
  const context = vm.createContext(globals);
  gas.Date = vm.runInContext('Date', context);
  gas.global = context;

  const files = options.files ||
    fs.readdirSync(ROOT).filter(name => name.endsWith('.gs')).sort();
  files.forEach(name => {
    vm.runInContext(fs.readFileSync(path.join(ROOT, name), 'utf8'), context, {filename: name});
  });

  // シート名の設定（管理者が実際のシートに合わせて変える）にテストが左右されないよう、テスト用の名前に置き換える。
  // 本物の設定のまま読むときは createGas({realSheetNames: true})（tests/project.test.js で設定の形を確かめる）
  if (!options.realSheetNames) {
    vm.runInContext('(' + function (rules) {
      if (typeof DIFF_RULES !== 'undefined') {
        Object.keys(DIFF_RULES).forEach(name => { delete DIFF_RULES[name]; });
        Object.assign(DIFF_RULES, rules);
      }
      if (typeof SVC_OPTIONS !== 'undefined') SVC_OPTIONS.requestSheet = 'リクエスト';
      if (typeof TASK_OPTIONS !== 'undefined') TASK_OPTIONS.sheet = 'タスク管理';
    } + ')(' + JSON.stringify(TEST_DIFF_RULES) + ')', context);
  }

  /* ----- テスト用の補助 ----- */

  /** グローバルの値（const も含む）を返す。 */
  gas.get = name => vm.runInContext(name, context);

  /** シートを作り、values（1行目から）を手入力したのと同じように入れる。 */
  gas.addSheet = (name, values, size) => {
    const s = size || {};
    const sheet = ss.insertSheet(name, {rows: s.rows || 20, columns: s.columns || 26});
    if (values && values.length) {
      const width = Math.max(...values.map(r => r.length));
      gas.asUser(() => sheet.getRange(1, 1, values.length, width)
        .setValues(values.map(r => r.concat(new Array(width - r.length).fill('')))));
    }
    return sheet;
  };

  /** 利用者の操作として実行する（gas.writes に残さない）。 */
  gas.asUser = fn => {
    gas.userInput++;
    try { return fn(); } finally { gas.userInput--; }
  };

  /** セルを選ぶ（getActiveSheet / getActiveRange / getCurrentCell に反映）。 */
  gas.select = (sheet, a1) => {
    gas.selection = {sheet, range: sheet.getRange(a1)};
  };

  /** 利用者がセルを編集したことにして、インストール型 onEdit のイベントを返す。 */
  gas.edit = (sheet, a1, value, opt) => {
    const user = (opt && opt.user) || gas.activeUser;
    const range = sheet.getRange(a1);
    const single = range.getNumRows() === 1 && range.getNumColumns() === 1;
    const oldValue = single ? range.getDisplayValue() : undefined;
    gas.asUser(() => (Array.isArray(value) ? range.setValues(value) : range.setValue(value)));
    const e = {source: ss, range, user: {getEmail: () => user}, authMode: 'FULL'};
    if (single) {
      e.value = range.getDisplayValue();
      e.oldValue = oldValue;
    }
    return e;
  };

  /** 変更トリガーのイベント（INSERT_ROW など）。 */
  gas.change = (changeType, opt) => {
    const user = (opt && opt.user) || gas.activeUser;
    return {source: ss, changeType, user: {getEmail: () => user}, authMode: 'FULL'};
  };

  /** シートの値を2次元配列で返す（getDisplayValues）。 */
  gas.dump = (sheet, a1) => {
    const target = typeof sheet === 'string' ? ss.getSheetByName(sheet) : sheet;
    if (!target) return [];
    const range = a1 ? target.getRange(a1) : target.getDataRange();
    return range.getDisplayValues();
  };

  /** 1行目を見出しとして、各行を {見出し: 値} にする。 */
  gas.records = sheetName => {
    const rows = gas.dump(sheetName);
    if (rows.length < 2) return [];
    const headers = rows[0];
    return rows.slice(1).map(row => {
      const out = {};
      headers.forEach((h, i) => { if (h) out[h] = row[i]; });
      return out;
    });
  };

  return gas;
}

/** vm の中で作られた値を、比較しやすい普通のオブジェクトにする。 */
function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

module.exports = {createGas, plain, colToNum, numToCol, ROOT};
