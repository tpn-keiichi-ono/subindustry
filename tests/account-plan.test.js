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
