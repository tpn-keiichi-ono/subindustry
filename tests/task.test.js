'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/** 新FMT（2行目が見出し。C: サブインダストリー / D: 得意先 / E: 案件名 / F: サービスのリクエスト） */
function setup() {
  const gas = createGas();
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'];
  gas.addSheet('新FMT', [['新FMT'], header,
    ['1', '', '食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['2', '', 'ドラッグストア', 'C社', 'EC立ち上げ', '在庫を店舗と共有したい']], {rows: 20, columns: 6});
  return {gas, g: gas.global};
}

const listOf = cell => {
  const rule = cell.getDataValidation();
  return rule && rule.getCriteriaType() === 'VALUE_IN_LIST' ? plain(rule.getCriteriaValues()[0]) : null;
};
const rangeOf = cell => {
  const rule = cell.getDataValidation();
  if (!rule || rule.getCriteriaType() !== 'VALUE_IN_RANGE' || rule.getAllowInvalid()) return null;
  const range = rule.getCriteriaValues()[0];
  return range.getSheet().getName() + '!' + range.getA1Notation();
};

test('setupTaskSheet：サービス案に対応付けるタスク管理シートを作る', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  assert.deepStrictEqual(gas.dump(sheet, 'A1:F1')[0], ['サービス案', 'タスク', '担当者', '期限', '状況', 'メモ']);
  assert.strictEqual(sheet.getFrozenRows(), 1);
  assert.strictEqual(rangeOf(sheet.getRange('A2')), 'リクエスト!G2:G1000', 'リクエストに付けたサービス案から選ぶ');
  assert.strictEqual(rangeOf(sheet.getRange('A' + sheet.getMaxRows())), 'リクエスト!G2:G1000', '最後の行まで付ける');
  assert.strictEqual(sheet.getRange('B2').getDataValidation(), null, 'タスクの列は自由に入力できる');
  assert.strictEqual(sheet.getRange('D2').getDataValidation().getCriteriaType(), 'DATE_IS_VALID_DATE');
  assert.deepStrictEqual(listOf(sheet.getRange('E2')), ['未着手', '対応中', '完了']);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /作りました/);
  assert.strictEqual(gas.alerts.length, 0, '前の版の列が無ければ確認しない');
});

test('setupTaskSheet：リクエスト シートが無いうちは、サービス案にプルダウンを付けずに知らせる', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  assert.strictEqual(gas.ss.getSheetByName('タスク管理').getRange('A2').getDataValidation(), null);
  assert.match(gas.toasts[0].message, /setupRequestSheet\(\) を実行してください/);
});

/** 前の版のタスク管理シート（A〜C が案件の連動プルダウン。サービスは右端に足されていた） */
function addOldTaskSheet(gas, withService) {
  const header = ['サブインダストリー', '得意先', '案件名', 'タスク', '担当者', '期限', '状況', 'メモ'];
  const row = ['食品スーパー', 'A社', 'アプリ刷新', '提案書を送る', '佐藤', '', '対応中', '先方に確認'];
  if (withService) { header.push('サービス'); row.push('会員分析基盤'); }
  return gas.addSheet('タスク管理', [header, row], {rows: 10, columns: 10});
}

test('setupTaskSheet：前の版のシートは、確認してから A〜C を削除し、サービス案を左端に置く（D列以降の値は残す）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const sheet = addOldTaskSheet(gas, true);
  gas.confirmAnswer = 'YES';
  g.setupTaskSheet();

  assert.strictEqual(gas.alerts.length, 1);
  assert.match(gas.alerts[0].message, /「サブインダストリー」「得意先」「案件名」の列を削除し/);
  assert.deepStrictEqual(gas.dump(sheet, 'A1:F2'), [
    ['サービス案', 'タスク', '担当者', '期限', '状況', 'メモ'],
    ['会員分析基盤', '提案書を送る', '佐藤', '', '対応中', '先方に確認']
  ]);
  assert.strictEqual(sheet.getLastColumn(), 6);
  assert.strictEqual(rangeOf(sheet.getRange('A2')), 'リクエスト!G2:G1000');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /列を削除し、「サービス案」を左端に置きました/);
});

test('setupTaskSheet：サービスの列が無い前の版でも、サービス案を左端に足す', () => {
  const {gas, g} = setup();
  const sheet = addOldTaskSheet(gas, false);
  g.setupTaskSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A1:F2'), [
    ['サービス案', 'タスク', '担当者', '期限', '状況', 'メモ'],
    ['', '提案書を送る', '佐藤', '', '対応中', '先方に確認']
  ]);
});

test('setupTaskSheet：削除しないと答えたら A〜C は残し、サービス案だけを用意する', () => {
  const {gas, g} = setup();
  const sheet = addOldTaskSheet(gas, true);
  gas.confirmAnswer = 'NO';
  g.setupTaskSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A1:I1')[0],
    ['サブインダストリー', '得意先', '案件名', 'タスク', '担当者', '期限', '状況', 'メモ', 'サービス案']);
  assert.strictEqual(gas.dump(sheet, 'I2')[0][0], '会員分析基盤');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /列は残しました/);
});

test('setupTaskSheet：何度実行しても入力済みのタスクは消えない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  gas.asUser(() => sheet.getRange('A2:E2').setValues([['会員分析基盤', '要件を聞く', '佐藤', '2026/11/01', '未着手']]));
  g.setupTaskSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A2:E2')[0], ['会員分析基盤', '要件を聞く', '佐藤', '2026/11/01', '未着手']);
  assert.strictEqual(gas.alerts.length, 0);
});

test('タスク管理シートの編集で動く処理は無い（単純トリガーの onEdit を置かない。プルダウンはリクエストのサービス案を参照するだけ）', () => {
  const {g} = setup();
  assert.strictEqual(typeof g.onEdit, 'undefined');
});
