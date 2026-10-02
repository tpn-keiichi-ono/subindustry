'use strict';

const {test, assert} = require('./lib/harness');
const {createGas} = require('./lib/gas-mock');

test('リポジトリ直下の *.gs を1つのグローバルスコープに読み込める（同じ名前の const がない）', () => {
  const gas = createGas();
  ['DIFF_RULES', 'DIFF_OPTIONS', 'CRED_OPTIONS', 'CHG_OPTIONS', 'HS_OPTIONS'].forEach(name => {
    assert.strictEqual(typeof gas.get(name), 'object', name);
  });
});

test('onOpen で「履歴機能」と「アカウントプランシート作成」の両方のメニューが出る', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  assert.deepStrictEqual(gas.menus.map(m => m.title), ['履歴機能', 'アカウントプランシート作成']);
});

test('「履歴機能」のメニューは「権限を承認する（初回のみ）」だけ（ほかの処理はエディタから実行する）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  const menu = gas.menus.find(m => m.title === '履歴機能');
  assert.deepStrictEqual(menu.items, [{label: '権限を承認する（初回のみ）', fn: 'authorizeHistoryFeatures'}]);
  // メニューから外した関数も、エディタから実行できるよう残っている
  ['openHistorySidebar', 'openCredentialDialog', 'openChangeHistoryDialog', 'openOpportunityCanvas',
    'setupCredentialLauncher', 'setupChangeHistoryLauncher', 'setupHistorySidebarAutoOpen',
    'addSampleHistory', 'removeSampleHistory', 'toggleHistorySidebarAutoOpen'].forEach(fn => {
    assert.strictEqual(typeof gas.global[fn], 'function', fn + ' がありません');
  });
});

test('権限を承認する：すべての権限の承認を確かめてから、サイドバーを開き直す', () => {
  const gas = createGas();
  gas.global.authorizeHistoryFeatures();
  assert.deepStrictEqual(gas.scopeChecks, ['FULL']);
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'HistorySidebarView');
  assert.strictEqual(gas.toasts.length, 1);
  assert.strictEqual(gas.toasts[0].title, '履歴機能');
  assert.strictEqual(gas.writes.length, 0, 'シートに書き込まない');
});

test('メニュー項目の関数がすべて存在し、内部用（末尾 _）ではない', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  const items = [].concat(...gas.menus.map(m => m.items)).filter(item => item.fn);
  assert.ok(items.length > 0);
  items.forEach(item => {
    assert.strictEqual(typeof gas.global[item.fn], 'function', item.fn + ' がありません');
    assert.ok(!item.fn.endsWith('_'), item.fn + ' は内部用の名前です');
  });
});

test('onOpen が単純トリガーとして動く（承認前でも失敗しない範囲の処理だけ）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  assert.strictEqual(gas.writes.length, 0, 'onOpen ではシートに書き込まない');
  assert.strictEqual(gas.triggers.length, 0, 'onOpen ではトリガーを作らない');
});
