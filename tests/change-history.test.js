'use strict';

const {test, assert} = require('./lib/harness');
const {plain} = require('./lib/gas-mock');
const {setupProject} = require('./lib/fixture');

test('getChangeHistoryData：その行の得意先の変更を新しい順に、変更前・変更後つきで返す', () => {
  const {gas, g, sheet} = setupProject();
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注', {user: 'a@example.com'}));
  g.recordDiffEdit(gas.edit(sheet, 'E3', '高橋', {user: 'b@example.com'}));
  g.recordDiffEdit(gas.edit(sheet, 'F4', '失注'));   // 別の得意先（B社）

  const data = plain(g.getChangeHistoryData('新FMT', 3));
  assert.strictEqual(data.customer, 'A社');
  assert.strictEqual(data.events.length, 2);
  const [latest, first] = data.events;
  assert.strictEqual(latest.editor, 'b@example.com');
  assert.deepStrictEqual(latest.changes.map(c => [c.column, c.before, c.after]), [['担当', '佐藤', '高橋']]);
  assert.strictEqual(first.type, 'EVT');
  assert.deepStrictEqual(first.changes.map(c => [c.column, c.before, c.after]), [['状況', '提案中', '受注']]);

  // 全項目の表示にボタン列は出さない
  const keys = latest.snapshot.map(([key]) => key);
  assert.ok(keys.includes('得意先'));
  assert.ok(!keys.includes('変更履歴'));
  assert.ok(!keys.some(k => /クレデンシャル/.test(k)));
});

test('getChangeHistoryData：行が並べ替わっても、得意先で記録を探す', () => {
  const {gas, g, sheet} = setupProject();
  g.recordDiffEdit(gas.edit(sheet, 'F4', '保留'));   // B社（4行目）
  gas.asUser(() => sheet.deleteRow(3));            // A社の行を削除 → B社は3行目へ
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));

  const data = plain(g.getChangeHistoryData('新FMT', 3));
  assert.strictEqual(data.customer, 'B社');
  assert.ok(data.events.some(ev => ev.changes.some(c => c.column === '状況' && c.after === '保留')));
});

test('getChangeHistoryData：記録が無ければ空で返す', () => {
  const {g} = setupProject();
  const data = plain(g.getChangeHistoryData('新FMT', 3));
  assert.deepStrictEqual(data.events, []);
  assert.strictEqual(data.customerMissing, false);
});

test('openChangeHistoryDialog：選んだ行の変更履歴モーダルを開く', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'E4');
  g.openChangeHistoryDialog();
  assert.strictEqual(gas.dialogs.length, 1);
  assert.strictEqual(gas.dialogs[0].html.file, 'ChangeHistoryDialog');
  assert.strictEqual(gas.dialogs[0].html.data.customer, 'B社');
});

test('openChangeHistoryDialog：見出し行を選んでいたら知らせる', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'E2');
  g.openChangeHistoryDialog();
  assert.strictEqual(gas.dialogs.length, 0);
  assert.match(gas.alerts[0].message, /データ行（3行目以降）を選んでください/);
});
