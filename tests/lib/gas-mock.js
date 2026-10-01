/**
 * Apps Script の最小限のモック。Node.js だけで src/*.gs のロジックを動かして確かめるためのもの。
 * 実際の SpreadsheetApp の挙動（書式・保護・トリガーなど）は再現していない。
 */
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, '..', '..', 'src');

/** 2次元配列を持つシートのモック。rows × cols のグリッド。 */
function makeSheet(name, id, rows, cols) {
  const data = Array.from({length: rows}, () => Array(cols).fill(''));
  const colNumber = s => s.split('').reduce((x, ch) => x * 26 + ch.charCodeAt(0) - 64, 0);
  const sheet = {
    name, id, data,
    getName: () => sheet.name,
    getSheetId: () => sheet.id,
    getMaxRows: () => sheet.data.length,
    getMaxColumns: () => sheet.data[0].length,
    getLastRow: () => {
      for (let r = sheet.data.length - 1; r >= 0; r--) if (sheet.data[r].some(v => v !== '')) return r + 1;
      return 0;
    },
    getLastColumn: () => {
      let m = 0;
      sheet.data.forEach(r => r.forEach((v, c) => { if (v !== '') m = Math.max(m, c + 1); }));
      return m;
    },
    insertRowsAfter: (after, n) => { for (let i = 0; i < n; i++) sheet.data.splice(after, 0, Array(sheet.data[0].length).fill('')); },
    insertColumnsAfter: (after, n) => sheet.data.forEach(r => { for (let i = 0; i < n; i++) r.splice(after, 0, ''); }),
    deleteRow: r => sheet.data.splice(r - 1, 1),
    clearContents: () => sheet.data.forEach(r => r.fill('')),
    hideSheet() {}, setFrozenRows() {}, setColumnWidth() {}, setColumnWidths() {}, insertColumnAfter() {},
    getCurrentCell: () => sheet.currentCell || null,
    getActiveRange: () => sheet.currentCell || null,
    getRange: (a, b, c, d) => {
      let r1, c1, nr, nc;
      if (typeof a === 'string') {
        let m = a.match(/^([A-Z]+)(\d+):([A-Z]+)(\d*)$/);
        if (m) {
          r1 = +m[2]; c1 = colNumber(m[1]); nc = colNumber(m[3]) - c1 + 1;
          nr = (m[4] ? +m[4] : sheet.data.length) - r1 + 1;
        } else if ((m = a.match(/^(\d+):(\d+)$/))) {           // '2:2' のような行指定
          r1 = +m[1]; nr = +m[2] - r1 + 1; c1 = 1; nc = sheet.data[0].length;
        } else throw new Error('unsupported A1: ' + a);
      } else { r1 = a; c1 = b; nr = c || 1; nc = d || 1; }
      if (r1 + nr - 1 > sheet.data.length || c1 + nc - 1 > sheet.data[0].length) {
        throw new Error('out of range ' + [r1, c1, nr, nc]);
      }
      const get = () => sheet.data.slice(r1 - 1, r1 - 1 + nr).map(r => r.slice(c1 - 1, c1 - 1 + nc));
      const range = {
        getSheet: () => sheet,
        getRow: () => r1, getColumn: () => c1,
        getLastRow: () => r1 + nr - 1, getLastColumn: () => c1 + nc - 1,
        getNumRows: () => nr, getNumColumns: () => nc,
        getValues: get,
        getValue: () => get()[0][0],
        getDisplayValues: () => get().map(r => r.map(v => String(v))),
        getDisplayValue: () => String(get()[0][0]),
        getFormulas: () => get().map(r => r.map(() => '')),
        getFormula: () => '',
        getMergedRanges: () => [],
        getA1Notation: () => 'R' + r1 + 'C' + c1,
        setValues: vals => { vals.forEach((row, i) => row.forEach((v, j) => { sheet.data[r1 - 1 + i][c1 - 1 + j] = v; })); return range; },
        setValue: v => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) sheet.data[r1 - 1 + i][c1 - 1 + j] = v; return range; },
        setRichTextValues: vals => { vals.forEach((row, i) => row.forEach((v, j) => { sheet.data[r1 - 1 + i][c1 - 1 + j] = v.text; })); return range; },
        setRichTextValue: v => range.setValue(v.text),
        setNumberFormat: () => range, insertCheckboxes: () => range, setVerticalAlignment: () => range,
        setWrap: () => range, setFontWeight: () => range, setBackground: () => range,
        clearContent: () => range.setValue(''), clearDataValidations: () => range
      };
      return range;
    }
  };
  return sheet;
}

/** スプレッドシートと各サービスのモックをまとめて作る。 */
function makeEnvironment() {
  const props = {};
  const cache = {};
  const sheets = [];
  const ss = {
    sheets,
    getSheets: () => sheets,
    getSheetByName: n => sheets.find(s => s.name === n) || null,
    insertSheet: n => { const s = makeSheet(n, 1000 + sheets.length, 1, 12); sheets.push(s); return s; },
    getSpreadsheetTimeZone: () => 'Asia/Tokyo',
    getActiveSheet: () => ss.activeSheet || sheets[0],
    toast() {},
    getId: () => 'SPREADSHEET_ID'
  };
  const richText = () => {
    const o = {t: '', setText(x) { o.t = x; return o; }, setTextStyle() { return o; }, build() { return {text: o.t}; }};
    return o;
  };
  const textStyle = () => {
    const o = {setForegroundColor() { return o; }, setBold() { return o; }, setStrikethrough() { return o; }, build() { return {}; }};
    return o;
  };
  const globals = {
    PropertiesService: {
      getDocumentProperties: () => ({
        getProperty: k => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = v; },
        deleteProperty: k => { delete props[k]; }
      })
    },
    CacheService: { getUserCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getActiveSheet: () => ss.getActiveSheet(),
      flush() {},
      newRichTextValue: richText,
      newTextStyle: textStyle
    },
    LockService: { getDocumentLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }) },
    Utilities: {
      getUuid: (() => { let n = 0; return () => 'uuid-' + (++n).toString().padStart(8, '0') + '-0000'; })(),
      formatDate: () => '2026/10/01 10:00'
    },
    ScriptApp: {
      getProjectTriggers: () => [],
      newTrigger: () => ({ forSpreadsheet() { return this; }, onEdit() { return this; }, onChange() { return this; },
        onOpen() { return this; }, timeBased() { return this; }, everyMinutes() { return this; }, create() {} }),
      EventType: { CLOCK: 'CLOCK' },
      deleteTrigger() {}
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => 'user@example.com' }),
      getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' })
    },
    HtmlService: {},
    console: { log() {}, warn() {}, error: console.error }
  };
  return {ss, props, cache, globals, makeSheet};
}

/** src のファイルを1つの関数スコープで読み込み、指定した関数を返す（Apps Script と同じく全ファイルが同じスコープ）。 */
function loadProject(env, files, exportNames) {
  const source = files.map(f => fs.readFileSync(path.join(SRC, f), 'utf8')).join('\n');
  const factory = new Function(...Object.keys(env.globals),
    source + '\nreturn {' + exportNames.join(', ') + '};');
  return factory(...Object.values(env.globals));
}

module.exports = {makeEnvironment, loadProject, makeSheet};
