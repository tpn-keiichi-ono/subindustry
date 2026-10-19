'use strict';

const {test, assert} = require('./lib/harness');
const {createGas} = require('./lib/gas-mock');

const SHEET = 'アカウントプラン';

/** 1行目が見出し。E列（5）が得意先、M列（13）が作成したファイルの URL。 */
function addPlanSheet(gas, rows) {
  const header = new Array(13).fill('');
  header[4] = '得意先';
  header[12] = 'アカウントプラン';
  return gas.addSheet(SHEET, [header].concat(rows.map(([e, m]) => {
    const row = new Array(13).fill('');
    row[4] = e;
    row[12] = m || '';
    return row;
  })), {rows: 10, columns: 13});
}

test('checkTargetCount：E列が入力済みで M列が空の行を数える', () => {
  const gas = createGas();
  addPlanSheet(gas, [['A社', ''], ['B社', 'https://example.com/b'], ['', ''], ['Ｃ社', '']]);
  assert.strictEqual(gas.global.checkTargetCount(SHEET), 2);
});

test('checkTargetCount：シートが無ければ知らせる', () => {
  const gas = createGas();
  assert.throws(() => gas.global.checkTargetCount('ないシート'), /シート「ないシート」が見つかりません/);
});

test('executeCopyProcess：対象の行ごとにテンプレートをコピーし、M列に URL を入れる', () => {
  const gas = createGas();
  const sheet = addPlanSheet(gas, [['A社', ''], ['B社', 'https://example.com/b'], ['Ｃ社', '']]);
  const message = gas.global.executeCopyProcess(SHEET);
  assert.strictEqual(message, '2 件の処理が完了しました。');
  assert.deepStrictEqual(gas.driveCopies.map(c => c.name),
    ['33シナリオ起点_アカウントプラン_A社', '33シナリオ起点_アカウントプラン_Ｃ社']);
  assert.match(sheet.getRange('M2').getValue(), /^https:\/\/docs\.google\.com\//);
  assert.strictEqual(sheet.getRange('M3').getValue(), 'https://example.com/b');
  assert.match(sheet.getRange('M4').getValue(), /^https:\/\/docs\.google\.com\//);
});

test('protectSpecificColumnsWithLock：「案件_」のシートの列だけを保護し、2回目は増やさない', () => {
  const gas = createGas();
  const target = gas.addSheet('案件_A社', [['見出し']], {rows: 5, columns: 30});
  const other = gas.addSheet('一覧', [['見出し']], {rows: 5, columns: 30});
  gas.global.protectSpecificColumnsWithLock();
  gas.global.protectSpecificColumnsWithLock();
  assert.deepStrictEqual(target.getProtections().map(p => p.getDescription()), [
    '【編集不可】案件_A社_16列目', '【編集不可】案件_A社_18列目', '【編集不可】案件_A社_20列目',
    '【編集不可】案件_A社_24列目', '【編集不可】案件_A社_27列目', '【編集不可】案件_A社_28列目'
  ]);
  assert.strictEqual(other.getProtections().length, 0);
  target.getProtections().forEach(p => assert.strictEqual(p.canDomainEdit(), false));
});

test('executeCopyProcess：コピーと書き込みはドキュメントロックの中で行う', () => {
  const gas = createGas();
  addPlanSheet(gas, [['A社', ''], ['B社', '']]);
  gas.global.executeCopyProcess(SHEET);
  assert.strictEqual(gas.driveCopies.length, 2);
  gas.driveCopies.forEach(c => assert.ok(c.locked, 'ロックの外でコピー: ' + c.name));
  const writes = gas.writes.filter(w => w.sheet === SHEET);
  assert.strictEqual(writes.length, 2);
  writes.forEach(w => assert.ok(w.locked, 'ロックの外で書き込み: ' + w.a1));
  assert.strictEqual(gas.lockHolders, 0, 'ロックが外れていません');
});

test('executeCopyProcess：ロックが取れないときはコピーせずに止め、そう伝える', () => {
  const gas = createGas();
  const sheet = addPlanSheet(gas, [['A社', '']]);
  gas.lockBusy = true;
  const message = gas.global.executeCopyProcess(SHEET);
  assert.strictEqual(message, '【中断】他の処理が実行中のため0件で停止しました。少し待ってから再度実行してください。');
  assert.strictEqual(gas.driveCopies.length, 0);
  assert.strictEqual(sheet.getRange('M2').getValue(), '');
});

test('executeCopyProcess：読み込んだあとで行がずれたら、違う行に URL を入れない', () => {
  const gas = createGas();
  const sheet = addPlanSheet(gas, [['A社', ''], ['B社', ''], ['Ｃ社', '']]);
  // 1件目のロックを取る直前に、だれかが2行目に行を挿入した
  const LockService = gas.global.LockService;
  const original = LockService.getDocumentLock;
  let inserted = false;
  LockService.getDocumentLock = () => {
    if (!inserted) {
      inserted = true;
      gas.asUser(() => sheet.insertRowsBefore(2, 1));
    }
    return original();
  };

  gas.global.executeCopyProcess(SHEET);
  assert.strictEqual(gas.driveCopies.length, 0);
  assert.deepStrictEqual(sheet.getRange('M2:M5').getValues(), [[''], [''], [''], ['']]);

  // もう一度実行すれば、正しい行に入る
  gas.global.executeCopyProcess(SHEET);
  const rows = sheet.getRange('E3:M5').getValues().map(r => [r[0], r[8] !== '']);
  assert.deepStrictEqual(rows, [['A社', true], ['B社', true], ['Ｃ社', true]]);
  assert.deepStrictEqual(gas.driveCopies.map(c => c.name.replace('33シナリオ起点_アカウントプラン_', '')),
    ['A社', 'B社', 'Ｃ社']);
});

test('executeCopyProcess：差分追跡の対象シートなら、書き込んだ URL を取りこぼし回収で記録する', () => {
  const {setupProject} = require('./lib/fixture');
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => sheet.getRange('M2').setValue('アカウントプラン'));   // 見出し（2行目は処理しない）

  const message = g.executeCopyProcess('新FMT');
  assert.strictEqual(message, '3 件の処理が完了しました。');
  assert.ok(gas.props.document.get('CHAR_DIFF_DIRTY_V1_' + sheet.getSheetId()), '取りこぼしの印が付いていません');

  g.catchUpDiffTracking();
  const log = gas.records('変更履歴_差分');
  assert.deepStrictEqual(log.map(r => [r['セル'], r['列名']]),
    [['M3', 'アカウントプラン'], ['M4', 'アカウントプラン'], ['M5', 'アカウントプラン']]);
  log.forEach(r => {
    assert.match(r['イベントID'], /^REC-/);
    assert.match(r['変更後'], /^https:\/\/docs\.google\.com\//);
  });
});

test('executeCopyProcess：差分追跡の対象でないシートには印を付けない', () => {
  const gas = createGas();
  addPlanSheet(gas, [['A社', '']]);
  gas.global.executeCopyProcess(SHEET);
  assert.deepStrictEqual([...gas.props.document.keys()].filter(k => k.startsWith('CHAR_DIFF_DIRTY')), []);
});
