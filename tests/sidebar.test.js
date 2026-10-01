'use strict';

const {test, assert} = require('./lib/harness');
const {plain} = require('./lib/gas-mock');
const {setupProject} = require('./lib/fixture');

test('getSidebarSelection：選んだ行の得意先を返し、同じ行なら読み込みを省く', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'F5');
  const first = plain(g.getSidebarSelection(null));
  assert.deepStrictEqual(first, {sheet: '新FMT', row: 5, customer: 'Ｃ社', key: 'C社'});
  assert.deepStrictEqual(plain(g.getSidebarSelection({sheet: '新FMT', row: 5})), {sheet: '新FMT', row: 5, same: true});
});

test('getSidebarSelection：見出し行・対象外のシートでは理由を返す', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'A2');
  assert.strictEqual(g.getSidebarSelection(null).reason, 'header');
  gas.select(gas.ss.getSheetByName('変更履歴_差分'), 'A2');
  assert.strictEqual(g.getSidebarSelection(null).reason, 'sheet');
});

test('getSidebarBundle：得意先ごとの履歴と、シート・得意先ごとの最近の変更をまとめて返す', () => {
  const {gas, g, sheet} = setupProject();
  const data = plain(g.getCredentialDialogData('新FMT', 5));
  g.saveCredentialEntry({
    sheetName: '新FMT', row: 5, customer: data.customer, customerColumn: data.customerColumn, linkId: '',
    values: {plan: '2026/12/01', date: '', kind: 'オファリング', title: '', person: '田中', name: '提案', note: '',
      owner: '佐藤', ownerDept: ''}
  });
  g.recordDiffEdit(gas.edit(sheet, 'F5', '提案中'));

  const bundle = plain(g.getSidebarBundle());
  assert.deepStrictEqual(bundle.sheets, ['新FMT', '新FMT2']);
  assert.strictEqual(bundle.credentials['C社'].length, 1);
  assert.strictEqual(bundle.credentials['C社'][0].values.name, '提案');
  const changes = bundle.changes['新FMT']['C社'];
  assert.strictEqual(changes.length, 1);
  assert.deepStrictEqual(changes[0].changes, [{column: '状況', before: '', after: '提案中'}]);
  assert.strictEqual(changes[0].editor, 'user');
});

test('getSidebarBundle：読み取りだけで、シートに書き込まない', () => {
  const {gas, g} = setupProject();
  gas.writes.length = 0;
  g.getSidebarBundle();
  assert.deepStrictEqual(gas.writes, []);
});

test('自動表示：トリガーは1つだけ。オフにした人には開かない', () => {
  const {gas, g} = setupProject();
  g.setupHistorySidebarAutoOpen();
  g.setupHistorySidebarAutoOpen();
  assert.strictEqual(gas.triggers.filter(t => t.getHandlerFunction() === 'autoOpenHistorySidebar').length, 1);

  g.autoOpenHistorySidebar({user: {getEmail: () => 'user@example.com'}});
  assert.strictEqual(gas.sidebars.length, 1);

  g.toggleHistorySidebarAutoOpen();   // user@example.com がオフにする
  g.autoOpenHistorySidebar({user: {getEmail: () => 'user@example.com'}});
  assert.strictEqual(gas.sidebars.length, 1);
  g.autoOpenHistorySidebar({user: {getEmail: () => 'other@example.com'}});
  assert.strictEqual(gas.sidebars.length, 2);
});

test('getSidebarBundle：変更は行（シート・得意先）ごとに新しい方から上限まで。超えた行には印を付ける', () => {
  const {gas, g, sheet} = setupProject();
  const limit = gas.get('HS_OPTIONS').eventsPerCustomer;
  for (let i = 0; i <= limit; i++) g.recordDiffEdit(gas.edit(sheet, 'F3', '状況' + i));   // A社に limit + 1 回
  g.recordDiffEdit(gas.edit(sheet, 'F4', '保留'));                                        // B社に1回

  const bundle = plain(g.getSidebarBundle());
  const events = bundle.changes['新FMT']['A社'];
  assert.strictEqual(events.length, limit);
  assert.strictEqual(events[0].changes[0].after, '状況' + limit, '新しい順になっていません');
  assert.strictEqual(events[limit - 1].changes[0].after, '状況1', '最も古い1回が残っています');
  assert.deepStrictEqual(bundle.changesMore, {'新FMT': {'A社': true}});
  assert.strictEqual(bundle.eventsPerCustomer, limit);
  assert.strictEqual(bundle.changes['新FMT']['B社'].length, 1);
});

test('openHistorySidebar：サイドバーの名称は「得意先別の履歴情報」', () => {
  const {gas, g} = setupProject();
  g.openHistorySidebar();
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'HistorySidebarView');
  assert.strictEqual(gas.sidebars[0].html.title, '得意先別の履歴情報');
});
