'use strict';

const {test, assert} = require('./lib/harness');
const {createGas} = require('./lib/gas-mock');

test('リポジトリ直下の *.gs を1つのグローバルスコープに読み込める（同じ名前の const がない）', () => {
  const gas = createGas();
  ['DIFF_RULES', 'DIFF_OPTIONS', 'CRED_OPTIONS', 'CHG_OPTIONS', 'HS_OPTIONS'].forEach(name => {
    assert.strictEqual(typeof gas.get(name), 'object', name);
  });
});

test('onOpen で「履歴機能」「アカウントプランシート作成」「サービス管理」のメニューが出る', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  assert.deepStrictEqual(gas.menus.map(m => m.title), ['履歴機能', 'アカウントプランシート作成', 'サービス管理']);
  assert.deepStrictEqual(gas.menus[2].items, [{label: 'リクエストを追加', fn: 'openRequestPicker'}]);
});

test('「履歴機能」のメニューは「履歴サイドバーを開く」「権限を承認する（初回のみ）」だけ（ほかの処理はエディタから実行する）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  const menu = gas.menus.find(m => m.title === '履歴機能');
  assert.deepStrictEqual(menu.items, [
    {label: '履歴サイドバーを開く', fn: 'openHistorySidebar'},
    {label: '権限を承認する（初回のみ）', fn: 'authorizeHistoryFeatures'}
  ]);
  // メニューから外した関数も、エディタから実行できるよう残っている
  ['openHistorySidebar', 'openCredentialDialog', 'openChangeHistoryDialog', 'openOpportunityCanvas',
    'setupCredentialLauncher', 'setupChangeHistoryLauncher', 'removeHistorySidebarAutoOpen',
    'addSampleHistory', 'removeSampleHistory'].forEach(fn => {
    assert.strictEqual(typeof gas.global[fn], 'function', fn + ' がありません');
  });
});

test('権限を承認する：すべての権限の承認を確かめてから、対象のシートならサイドバーを開き直す', () => {
  const gas = createGas();
  gas.select(gas.addSheet('新FMT', [['新FMT']], {rows: 5, columns: 3}), 'A2');
  gas.global.authorizeHistoryFeatures();
  assert.deepStrictEqual(gas.scopeChecks, ['FULL']);
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'HistorySidebarView');
  assert.strictEqual(gas.toasts.length, 1);
  assert.strictEqual(gas.toasts[0].title, '履歴機能');
  assert.match(gas.toasts[0].message, /履歴サイドバーを開きました/);
  assert.strictEqual(gas.writes.length, 0, 'シートに書き込まない');
});

test('権限を承認する：対象外のシートでは、承認だけしてサイドバーは開かない', () => {
  const gas = createGas();
  gas.select(gas.addSheet('メモ用', [['']], {rows: 5, columns: 3}), 'A1');
  gas.global.authorizeHistoryFeatures();
  assert.deepStrictEqual(gas.scopeChecks, ['FULL']);
  assert.strictEqual(gas.sidebars.length, 0);
  assert.strictEqual(gas.alerts.length, 0, '承認のときは画面を止めない（トーストで知らせる）');
  assert.strictEqual(gas.toasts[0].message, '権限は承認済みです。履歴サイドバーは「新FMT」「新FMT2」のシートで開いてください。');
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

test('単純トリガーの onEdit は置かない（リクエストは選択パネルで登録する。編集のたびに動く処理を作らない）', () => {
  const gas = createGas();
  assert.strictEqual(typeof gas.global.onEdit, 'undefined');
  assert.strictEqual(typeof gas.global.doGet, 'function', 'doGet は1つ（HistorySidebar.gs）');
});

test('onOpen が単純トリガーとして動く（承認前でも失敗しない範囲の処理だけ）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  assert.strictEqual(gas.writes.length, 0, 'onOpen ではシートに書き込まない');
  assert.strictEqual(gas.triggers.length, 0, 'onOpen ではトリガーを作らない');
  assert.strictEqual(gas.sidebars.length, 0, 'onOpen ではサイドバーを開かない（利用者がメニューから開く）');
});

test('本物のシート名の設定：DIFF_RULES・前の名前（formerNames）・リクエスト・タスク管理のシート名が重ならない', () => {
  const gas = createGas({realSheetNames: true});
  const rules = gas.get('DIFF_RULES');
  const diff = gas.get('DIFF_OPTIONS');
  const svc = gas.get('SVC_OPTIONS');
  const names = Object.keys(rules);
  assert.ok(names.length > 0);
  names.forEach(name => {
    const rule = rules[name];
    assert.ok(Number.isInteger(rule.headerRow) && rule.headerRow >= 1, name + ' の headerRow');
    assert.ok(Array.isArray(rule.ranges) && rule.ranges.length > 0, name + ' の ranges');
    assert.ok(!rule.formerNames || Array.isArray(rule.formerNames), name + ' の formerNames は配列');
  });
  assert.ok(names.some(name => rules[name].requests), 'リクエストを読むシート（requests: true）が1つ以上ある');

  const former = [].concat(...names.map(name => rules[name].formerNames || []));
  former.forEach(f => assert.ok(!names.includes(f), '前の名前「' + f + '」が今のシート名と同じ'));
  assert.strictEqual(new Set(former).size, former.length, '同じ前の名前が2つのシートにある');

  // 機能ごとのシート名は、追跡するシートやほかの機能のシートと重ならない
  const others = [diff.logSheet, diff.rowSnapshotSheet, gas.get('CRED_OPTIONS').logSheet,
    svc.requestSheet, svc.legacyListSheet, gas.get('TASK_OPTIONS').sheet];
  assert.strictEqual(new Set(others).size, others.length, '機能ごとのシート名が重なっている');
  others.forEach(name => {
    assert.ok(!names.includes(name) && !former.includes(name), '「' + name + '」が追跡するシートと同じ名前');
    assert.ok(!name.startsWith(diff.snapshotPrefix) || name === diff.snapshotPrefix, name);
  });
});
