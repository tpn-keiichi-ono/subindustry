'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/*
 * 追跡シートに案件の行を足す（test.gs の addDealRows）
 * 本物と同じ形：2行目が見出し、3行目の B（No.）・O（期待値）に、最後の行までの範囲で計算する数式。プルダウンも最後の行まで。
 */
const HEADERS = ['最終更新日時\n（自動入力）', 'No.', 'サブインダストリー', '得意先', '先方部門', '先方担当役職', '先方担当氏名',
  'クライアントとのリレーション\nクライアント内のガバナンス', '案件名', '案件のスコープ', 'スコープに対する競合他社', 'プロジェクト開始-終了', '案件ステータス',
  '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）'];
const WIDTH = 32;

function setup(options) {
  const o = options || {};
  const rows = o.rows || 12;          // シートの行数
  const end = o.end || rows;          // 数式・プルダウンの範囲の最後の行
  const gas = createGas();
  const g = gas.global;
  const pad = row => row.concat(new Array(WIDTH - row.length).fill(''));
  ['新FMT', '新FMT2'].forEach(name => {
    const deals = (o.deals || ['A社', 'B社', 'C社']).map(c => pad(['', '', name, c, '', '', '', '', c + 'の案件', '', '', '', '提案中', '10']));
    const sheet = gas.addSheet(name, [pad(['']), pad(HEADERS)].concat(deals), {rows, columns: WIDTH});
    gas.asUser(() => {
      sheet.getRange(3, 2).setFormula('=ROW(B3:B' + end + ')-2');
      sheet.getRange(3, 15).setFormula('=IF(M3:M' + end + '="",,N3:N' + end + '*0.5)');
    });
    sheet.getRange('M3:M' + end).setDataValidation(g.SpreadsheetApp.newDataValidation().requireValueInList(['提案中', '受注'], true).build());
    sheet.getRange('N3:N' + end).setNumberFormat('#,##0');
  });
  g.setupDiffTracking();
  gas.writes.length = 0;
  return {gas, g, sheet: gas.ss.getSheetByName('新FMT')};
}

const LOG = '変更履歴_差分';
const state = (gas, sheet) => JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'))[sheet.getSheetId()];

test('addDealRows：数式・プルダウンの範囲の最後の行の上に足し（範囲が伸びる）、書式・プルダウンを上の行とそろえる。差分追跡には何も記録しない', () => {
  const {gas, g, sheet} = setup();
  gas.promptAnswers = ['1', '3'];
  g.addDealRows();
  assert.strictEqual(sheet.getMaxRows(), 15);
  assert.deepStrictEqual(plain(sheet.getRange('D3:D15').getValues().map(r => r[0])),
    ['A社', 'B社', 'C社', '', '', '', '', '', '', '', '', '', ''], '案件はそのまま');
  [12, 13, 14].forEach(row => {
    assert.ok(sheet.getRange(row, 13).getDataValidation(), row + '行目にもプルダウン');
    assert.strictEqual(sheet.getRange(row, 14).getNumberFormats()[0][0], '#,##0', '書式も上の行と同じ');
  });
  assert.ok(gas.writes.filter(w => w.sheet === '新FMT').every(w => w.locked), '書き込みはロックの中');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：3行足しました（12〜14行目。案件の行は 3〜15行目）/);
  // 差分追跡：行をそろえるだけで記録しない。そのあとの編集はふつうに記録する
  assert.ok(!gas.ss.getSheetByName(LOG) || gas.ss.getSheetByName(LOG).getLastRow() <= 1);
  assert.deepStrictEqual([state(gas, sheet).paused, state(gas, sheet).gridRows], [false, 15]);
  g.recordDiffEdit(gas.edit(sheet, 'D12', 'D社'));
  assert.strictEqual(gas.ss.getSheetByName(LOG).getLastRow(), 2);
  assert.strictEqual(gas.ss.getSheetByName('新FMT2').getMaxRows(), 12, '選ばなかったシートはそのまま');
});

test('addDealRows：範囲の最後の行に案件があれば、その案件を足した行の上に動かす（空の行を案件の間に入れない）', () => {
  const deals = ['A社', 'B社', 'C社', 'D社', 'E社', 'F社', 'G社', 'H社', 'I社', 'J社'];   // 3〜12行目（最後の行まで）
  const {gas, g, sheet} = setup({deals});
  gas.promptAnswers = ['1', '2'];
  g.addDealRows();
  assert.deepStrictEqual(plain(sheet.getRange('D3:D14').getValues().map(r => r[0])), deals.concat(['', '']));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /2行足しました（13〜14行目。案件の行は 3〜14行目。最後の行の案件は、足した行の上に動かしました）/);
  assert.ok(!gas.ss.getSheetByName(LOG) || gas.ss.getSheetByName(LOG).getLastRow() <= 1, '動かしただけの行は記録しない');
});

test('addDealRows：数式の範囲がシートの最後の行より上で終わっていれば、その範囲の最後の行の上に足す。行の数が正しくなければ止める', () => {
  const {gas, g, sheet} = setup({rows: 14, end: 12});
  gas.promptAnswers = ['1', '１'];   // 全角の数字も読む
  g.addDealRows();
  assert.strictEqual(sheet.getMaxRows(), 15);
  assert.ok(sheet.getRange(12, 13).getDataValidation(), '12行目（範囲の最後の行の上）に足す');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /1行足しました（12〜12行目。案件の行は 3〜13行目）/);
  gas.promptAnswers = ['1', '0'];
  assert.throws(() => g.addDealRows(), /足す行の数は 1〜500 の数字で入力してください/);
  gas.promptAnswers = ['1', 'たくさん'];
  assert.throws(() => g.addDealRows(), /足す行の数は/);
});
