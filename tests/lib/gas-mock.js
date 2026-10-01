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
  const map = {yyyy: parts.year, MM: parts.month, dd: parts.day, HH: parts.hour, mm: parts.minute, ss: parts.second};
  return String(format).replace(/yyyy|MM|dd|HH|mm|ss/g, t => map[t]);
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

  getValues() { return this.map(cell => (cell ? cell.v : '')); }
  getDisplayValues() { return this.map(cell => this.sheet.gas.display(cell)); }
  getFormulas() { return this.map(cell => (cell ? cell.f : '')); }
  getRichTextValues() {
    return this.map(cell => (cell && cell.rich) ||
      newRichTextValue().setText(this.sheet.gas.display(cell)).build());
  }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getFormula() { return this.getFormulas()[0][0]; }
  getNote() { const cell = this.sheet.cell(this.row, this.col, false); return (cell && cell.note) || ''; }

  setValue(value) {
    this.write('setValue');
    this.each(cell => this.sheet.gas.store(cell, value));
    return this;
  }

  setValues(matrix) {
    this.checkShape(matrix);
    this.write('setValues');
    this.each((cell, r, c) => this.sheet.gas.store(cell, matrix[r][c]));
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
  setFontWeight() { return this; }
  setBackground() { return this; }
  setVerticalAlignment() { return this; }
  setWrap() { return this; }

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

  clearDataValidations() { this.each(cell => { cell.checkbox = false; }); return this; }
  getMergedRanges() { return []; }

  protect() {
    const protection = new Protection(this);
    this.sheet.protections.push(protection);
    return protection;
  }
}

class Protection {
  constructor(range) {
    this.range = range;
    this.description = '';
    this.editors = [];
    this.domainEdit = true;
    this.removed = false;
  }
  setDescription(text) { this.description = text; return this; }
  getDescription() { return this.description; }
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
    this.data = [];          // data[行-1][列-1] = {v, f, rich, nf, checkbox, note}
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
  isSheetHidden() { return this.hidden; }
  setFrozenRows(n) { this.frozenRows = n; return this; }
  getFrozenRows() { return this.frozenRows; }
  setColumnWidth() { return this; }
  setColumnWidths() { return this; }
  getProtections() { return this.protections.filter(p => !p.removed); }

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
    return sheet;
  }

  deleteSheet(sheet) { this.sheets = this.sheets.filter(s => s !== sheet); }
  getActiveSheet() { return (this.gas.selection && this.gas.selection.sheet) || this.sheets[0] || null; }
  toast(message, title) { this.gas.toasts.push({message: String(message), title: title == null ? '' : String(title)}); }
}

/* ---------------- 全体 ---------------- */

/**
 * テスト1件ぶんの Apps Script 環境を作り、リポジトリ直下の *.gs を読み込む。
 * options = {user, owner, timeZone, files}
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
    toasts: [],
    menus: [],
    dialogs: [],
    sidebars: [],
    alerts: [],
    logs: [],
    triggers: [],
    driveCopies: [],
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

  gas.recordWrite = (range, kind) => {
    if (gas.userInput) return;
    gas.writes.push({
      sheet: range.getSheet().getName(),
      a1: range.getA1Notation(),
      kind,
      locked: gas.lockHolders > 0
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
    showModelessDialog(html, title) { gas.dialogs.push({html, title, modeless: true}); },
    showSidebar(html) { gas.sidebars.push({html}); },
    /** YES_NO などで聞かれたときは gas.confirmAnswer（既定 'YES'）を返す */
    alert(a, b, c) {
      gas.alerts.push(b === undefined ? {message: a} : {title: a, message: b, buttons: c});
      return c && c !== 'OK' ? gas.confirmAnswer : 'OK';
    }
  };

  const SpreadsheetApp = {
    ProtectionType: {RANGE: 'RANGE', SHEET: 'SHEET'},
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

  class Lock {
    constructor() { this.mine = false; }
    tryLock() {
      if (this.mine) return true;
      if (gas.lockBusy) return false;
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
    getDocumentLock: () => new Lock(),
    getScriptLock: () => new Lock(),
    getUserLock: () => new Lock()
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
    getProjectTriggers: () => gas.triggers.slice(),
    deleteTrigger(trigger) { gas.triggers = gas.triggers.filter(t => t !== trigger); },
    newTrigger(handler) {
      const make = (eventType, source, sourceId) => ({
        create() {
          const trigger = new Trigger(handler, eventType, source, sourceId);
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
          const builder = Object.assign(make('CLOCK', 'CLOCK', null), {
            everyMinutes: () => builder,
            everyHours: () => builder,
            everyDays: () => builder
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
  const DriveApp = {
    getFileById(id) {
      return {
        getId: () => id,
        makeCopy(name, folder) {
          const copyId = 'copy-' + (++copySeq);
          gas.driveCopies.push({from: id, name, folder: folder ? folder.getId() : null, locked: gas.lockHolders > 0});
          return {
            getId: () => copyId,
            getName: () => name,
            getUrl: () => 'https://docs.google.com/spreadsheets/d/' + copyId + '/edit'
          };
        }
      };
    },
    getFolderById: id => ({getId: () => id})
  };

  const Browser = {msgBox(message) { gas.alerts.push({message: String(message)}); return 'ok'; }};

  const capture = level => (...args) => {
    const line = args.map(a => (typeof a === 'string' ? a : (a && a.stack) || String(a))).join(' ');
    gas.logs.push({level, line});
    if (process.env.GAS_VERBOSE) console.log('[' + level + '] ' + line);
  };
  const logger = {log: capture('log'), info: capture('info'), warn: capture('warn'), error: capture('error')};

  /* ----- プロジェクトの読み込み ----- */

  const context = vm.createContext({
    SpreadsheetApp, HtmlService, PropertiesService, LockService, Utilities, Session,
    ScriptApp, CacheService, DriveApp, Browser, console: logger
  });
  gas.Date = vm.runInContext('Date', context);
  gas.global = context;

  const files = options.files ||
    fs.readdirSync(ROOT).filter(name => name.endsWith('.gs')).sort();
  files.forEach(name => {
    vm.runInContext(fs.readFileSync(path.join(ROOT, name), 'utf8'), context, {filename: name});
  });

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
