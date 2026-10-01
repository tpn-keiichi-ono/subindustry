/** 行の追加・削除の自動反映（DiffTracking.gs v5）を確かめる。 */
const assert = require('assert');
const {makeEnvironment, loadProject, makeSheet} = require('./lib/gas-mock');

const env = makeEnvironment();
const api = loadProject(env, ['DiffTracking.gs', 'CredentialHistory.gs', 'ChangeHistory.gs'],
  ['setupDiffTracking', 'pauseDiffOnStructureChange', 'recordDiffEdit', 'DIFF_OPTIONS']);

// 追跡シート：2行目が見出し、A3:AF を追跡、AG 列に「最終更新日時」
function addSource(name, id) {
  const s = makeSheet(name, id, 12, 34);
  s.data[1] = Array.from({length: 34}, (_, i) => i === 32 ? '最終更新日時' : '列' + (i + 1));
  s.data[1][0] = '得意先';
  for (let r = 3; r <= 8; r++) { s.data[r - 1][0] = '得意先' + r; s.data[r - 1][1] = '値' + r; }
  env.ss.sheets.push(s);
  return s;
}
const sheet = addSource('新FMT', 1);
addSource('新FMT2', 2);
api.setupDiffTracking();

const logRows = () => env.ss.getSheetByName('変更履歴_差分').data.slice(1).filter(r => r[1]);
const state = () => JSON.parse(env.props[api.DIFF_OPTIONS.stateKey]);

// 1) 5行目を削除し、反映前に3行目も編集されていた
sheet.deleteRow(5);
sheet.data[2][1] = '値3-変更';
api.pauseDiffOnStructureChange({source: env.ss, changeType: 'REMOVE_ROW', user: {getEmail: () => 'deleter@example.com'}});

let logs = logRows();
const deleted = logs.find(r => r[6] === '行削除');
assert(deleted, '行削除が記録される');
assert.strictEqual(deleted[2], 'deleter@example.com', '削除した人が記録される');
assert(String(deleted[7]).includes('得意先5'), '削除された行の中身が記録される');
assert(logs.some(r => String(r[1]).startsWith('ROW-') && r[4] === 'B3'), '反映前の未記録の編集も記録される');
assert.strictEqual(state()['1'].gridRows, sheet.getMaxRows(), '行数が更新される');
assert.strictEqual(state()['1'].paused, false, '追跡は止まらない');

// 2) 行を挿入して入力。変更トリガーより先に編集トリガーが届く
sheet.insertRowsAfter(4, 1);
sheet.data[4][0] = '新しい得意先';
api.recordDiffEdit({source: env.ss, range: sheet.getRange(5, 1, 1, 1), user: {getEmail: () => 'typist@example.com'}});
logs = logRows();
const inserted = logs.find(r => r[6] === '行追加');
assert(inserted && inserted[2] === 'typist@example.com', '入力した行は「行追加」として入力者付きで記録される');

const countAfterInsert = logs.length;
api.pauseDiffOnStructureChange({source: env.ss, changeType: 'INSERT_ROW', user: {getEmail: () => 'x@example.com'}});
assert.strictEqual(logRows().length, countAfterInsert, '遅れて届いた変更トリガーで二重に記録しない');

// 3) 反映後の通常の編集
sheet.data[6][1] = '値7-更新';
api.recordDiffEdit({source: env.ss, range: sheet.getRange(7, 2, 1, 1), user: {getEmail: () => 'u@example.com'}});
const lastLog = logRows().slice(-1)[0];
assert(String(lastLog[1]).startsWith('EVT-') && lastLog[4] === 'B7', '通常の編集は EVT として記録される');

const before = logRows().length;
api.recordDiffEdit({source: env.ss, range: sheet.getRange(3, 1, 1, 1), user: {getEmail: () => 'u@example.com'}});
assert.strictEqual(logRows().length, before, '値が変わっていない編集は記録しない');

console.log('row-resync: ok');
