'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');
const {COL, setupProject} = require('./lib/fixture');

const LOG = '変更履歴_差分';
const SNAP = '変更時点スナップショット';

/* ---------------- 文字の差分・行の突き合わせ ---------------- */

test('diffCharacters_：増えた文字・消えた文字を取り出す', () => {
  const g = createGas().global;
  const result = plain(g.diffCharacters_('東京本社', '大阪本社', {remaining: 1e6}));
  assert.deepStrictEqual(result.parts, [
    {kind: 'del', text: '東京'},
    {kind: 'add', text: '大阪'},
    {kind: 'same', text: '本社'}
  ]);
  assert.strictEqual(result.coarse, false);
});

test('diffCharacters_：上限を超えるとまとめて「削除＋追加」にする', () => {
  const g = createGas().global;
  const result = plain(g.diffCharacters_('abc', 'xyz', {remaining: 1}));
  assert.strictEqual(result.coarse, true);
  assert.deepStrictEqual(result.parts, [{kind: 'del', text: 'abc'}, {kind: 'add', text: 'xyz'}]);
});

test('diffAlignRows_：削除された行と追加された行を見つける', () => {
  const g = createGas().global;
  assert.deepStrictEqual(plain(g.diffAlignRows_(['a', 'b', 'c'], ['a', 'c'])), [{old: [1], new: []}]);
  assert.deepStrictEqual(plain(g.diffAlignRows_(['a', 'c'], ['a', 'b', 'c'])), [{old: [], new: [1]}]);
  assert.deepStrictEqual(plain(g.diffAlignRows_(['a', 'b'], ['a', 'b'])), []);
});

/* ---------------- セットアップ ---------------- */

test('setupDiffTracking：比較基準を作り、トリガーを1つずつだけ設置する', () => {
  const {gas, g} = setupProject({launchers: false});
  g.setupDiffTracking();   // 2回目
  const handlers = gas.triggers.map(t => t.getHandlerFunction()).sort();
  assert.deepStrictEqual(handlers, ['catchUpDiffTracking', 'pauseDiffOnStructureChange', 'recordDiffEdit']);
  const state = JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'));
  assert.strictEqual(Object.keys(state).length, 2);
  Object.values(state).forEach(s => assert.strictEqual(s.paused, false));
  const baseline = gas.ss.getSheets().filter(s => s.getName().startsWith('__CHAR_DIFF_'));
  assert.strictEqual(baseline.length, 2);
  baseline.forEach(s => assert.ok(s.isSheetHidden()));
});

/* ---------------- 編集の記録 ---------------- */

test('recordDiffEdit：変更したセルを差分・編集者つきで記録し、最終更新日時を入れる', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注', {user: 'editor@example.com'}));

  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  const rec = log[0];
  assert.match(rec['イベントID'], /^EVT-\d{8}-\d{6}-[0-9A-F]{8}$/);
  assert.strictEqual(rec['編集者メールアドレス'], 'editor@example.com');
  assert.strictEqual(rec['シート'], '新FMT');
  assert.strictEqual(rec['セル'], 'F3');
  assert.strictEqual(rec['列名'], '状況');
  assert.strictEqual(rec['種類'], '文字列 -> 文字列');
  assert.strictEqual(rec['変更前'], '提案中');
  assert.strictEqual(rec['変更後'], '受注');

  // 最終更新日時（AG列）
  assert.ok(sheet.getRange(3, COL.stamp).getValue() instanceof gas.Date);

  // 行全体のスナップショット
  const snap = gas.records(SNAP);
  assert.strictEqual(snap.length, 1);
  assert.strictEqual(snap[0]['イベントID'], rec['イベントID']);
  assert.strictEqual(snap[0]['行番号'], '3');
  assert.strictEqual(snap[0]['変更セル'], 'F3');
  assert.strictEqual(snap[0]['得意先'], 'A社');
  assert.strictEqual(snap[0]['状況'], '受注');
});

test('recordDiffEdit：値が変わっていなければ何も記録しない', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));
  assert.deepStrictEqual(gas.records(LOG), []);
});

test('recordDiffEdit：「=」で始まる入力も、差分ログでは数式にならない', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '=IMPORTXML("https://example.com","//a")'));
  const log = gas.ss.getSheetByName(LOG);
  const formulas = log.getRange(2, 1, 1, 11).getFormulas()[0];
  assert.ok(formulas.every(f => f === ''), '差分ログに数式が入っています: ' + formulas.join(' / '));
  assert.strictEqual(gas.records(LOG)[0]['変更後'], '=IMPORTXML("https://example.com","//a")');
  assert.strictEqual(gas.records(LOG)[0]['種類'], '文字列 -> 数式');
});

test('recordDiffEdit：書き込みはすべてドキュメントロックの中で行う', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'E4', '高橋'));
  assert.ok(gas.writes.length > 0);
  gas.writes.forEach(w => assert.ok(w.locked, 'ロックの外で書き込み: ' + w.sheet + '!' + w.a1));
});

test('recordDiffEdit：ボタン列（チェックボックス）の操作は記録しない', () => {
  const {gas, g, sheet} = setupProject();
  const rule = gas.get('DIFF_RULES')['新FMT'];
  ['AH3', 'AI3'].forEach(a1 => {
    const e = gas.edit(sheet, a1, true);
    assert.strictEqual(g.credIsLauncherEdit_(e, rule), true, a1);
    g.recordDiffEdit(e);
  });
  assert.strictEqual(g.credIsLauncherEdit_(gas.edit(sheet, 'F3', '受注'), rule), false);
  assert.deepStrictEqual(gas.records(LOG), []);
});

test('recordDiffEdit：ロックが取れないときは取りこぼしとして印を付け、回収で記録する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.lockBusy = true;
  g.recordDiffEdit(gas.edit(sheet, 'F3', '失注'));
  assert.deepStrictEqual(gas.records(LOG), []);
  assert.ok(gas.props.document.get('CHAR_DIFF_DIRTY_V1_' + sheet.getSheetId()));

  gas.lockBusy = false;
  g.catchUpDiffTracking();
  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  assert.match(log[0]['イベントID'], /^REC-/);
  assert.strictEqual(log[0]['編集者メールアドレス'], '不明（取りこぼし回収）');
  assert.strictEqual(log[0]['種類'], '文字列 -> 文字列 / 取りこぼし回収');
  assert.strictEqual(gas.props.document.get('CHAR_DIFF_DIRTY_V1_' + sheet.getSheetId()), undefined);
});

test('catchUpDiffTracking：印が無ければ何もしない', () => {
  const {gas, g} = setupProject({launchers: false});
  g.catchUpDiffTracking();
  assert.deepStrictEqual(gas.records(LOG), []);
  assert.strictEqual(gas.writes.length, 0);
});

/* ---------------- 行・列の追加・削除 ---------------- */

test('行の追加：追加された行を「行追加」として記録し、追跡を続ける', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => {
    sheet.insertRowsAfter(3, 1);
    sheet.getRange(4, COL.customer, 1, 3).setValues([['Ｄ社', '伊藤', '新規']]);
  });
  g.pauseDiffOnStructureChange(gas.change('INSERT_ROW', {user: 'editor@example.com'}));

  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  assert.match(log[0]['イベントID'], /^ROW-/);
  assert.strictEqual(log[0]['種類'], '行追加');
  assert.strictEqual(log[0]['セル'], '4行');
  assert.match(log[0]['変更後'], /得意先: Ｄ社/);

  // 比較基準が取り直され、次の編集もふつうに記録される
  g.recordDiffEdit(gas.edit(sheet, 'F5', '保留'));   // もとの B社 の行
  const next = gas.records(LOG)[1];
  assert.match(next['イベントID'], /^EVT-/);
  assert.strictEqual(next['セル'], 'F5');
  assert.strictEqual(next['変更前'], '受注');
});

test('行の削除：削除された行の内容を「行削除」として記録する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => sheet.deleteRow(4));   // B社
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW', {user: 'editor@example.com'}));

  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  assert.strictEqual(log[0]['種類'], '行削除');
  assert.strictEqual(log[0]['編集者メールアドレス'], 'editor@example.com');
  assert.strictEqual(log[0]['セル'], '4行（削除前の行番号）');
  assert.match(log[0]['変更前'], /得意先: B社/);
});

test('列の追加：追跡を停止して知らせる', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => sheet.insertColumnAfter(2));
  g.pauseDiffOnStructureChange(gas.change('INSERT_COLUMN'));
  const state = JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'));
  assert.strictEqual(state[sheet.getSheetId()].paused, true);
  assert.ok(gas.toasts.some(t => /列構成が変わったため差分追跡を停止しました/.test(t.message)));
});
