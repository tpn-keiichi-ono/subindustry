'use strict';

const {test, assert} = require('./lib/harness');
const {plain} = require('./lib/gas-mock');
const {COL, WIDTH, headerRow, dataRow, setupProject} = require('./lib/fixture');

/** リクエストの見出し（G：サブインダストリー、H：案件名、I：サービスのリクエスト）を足した見出しの行 */
function requestHeaderRow() {
  const row = headerRow();
  row[6] = 'サブインダストリー';
  row[7] = '案件名';
  row[8] = 'サービスのリクエスト';
  return row;
}

/** 追跡シートの1行に、サブインダストリー・案件名・リクエストを入れる */
function requestRow(no, customer, subIndustry, project, request) {
  const row = dataRow(no, customer, '佐藤', '提案中');
  row[6] = subIndustry;
  row[7] = project;
  row[8] = request;
  return row;
}

/** 新FMT にリクエストの見出しと1件を足し、DIFF_RULES に「百貨店」を足す（シートを足したあとの状態） */
function addDepartmentStore(gas, header) {
  const fmt = gas.ss.getSheetByName('新FMT');
  gas.asUser(() => {
    fmt.getRange(2, 1, 1, WIDTH).setValues([requestHeaderRow()]);
    fmt.getRange(3, 7, 1, 3).setValues([['食品スーパー', 'アプリ刷新', '会員の購買分析をしたい']]);
  });
  const title = new Array(WIDTH).fill('');
  title[0] = '百貨店';
  const sheet = gas.addSheet('百貨店', [title, header || requestHeaderRow(),
    requestRow(1, 'D百貨店', '百貨店', '外商DX', '顧客の来店を予測したい')], {rows: 10, columns: WIDTH});
  gas.get('DIFF_RULES')['百貨店'] = {headerRow: 2, ranges: ['D3:AF'], requests: true};
  return sheet;
}

test('setupAfterSheetChange：差分追跡・2つのボタン列・サービスリクエストの候補をまとめて設定する（トリガーは増やさない）', () => {
  const {gas, g} = setupProject();
  const triggers = gas.triggers.map(t => t.getHandlerFunction()).sort();
  const sheet = addDepartmentStore(gas);
  gas.evaluateFormulas = true;

  g.setupAfterSheetChange();
  const message = gas.alerts[gas.alerts.length - 1].message;
  assert.match(message, /対象のシート：「新FMT」「新FMT2」「百貨店」/);
  assert.match(message, /「差分追跡」「クレデンシャルのボタン列」「変更履歴のボタン列」「サービスリクエストの候補」を設定しました/);
  assert.match(message, /「新FMT」「百貨店」のリクエスト 2件のうち 0件を選んでいます/, 'サービスリクエストの結果も知らせる');

  // ボタン列（チェックボックス）と差分追跡
  assert.strictEqual(sheet.getRange(3, COL.credButton).getValue(), false);
  assert.strictEqual(sheet.getRange(3, COL.changeButton).getValue(), false);
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注'));
  const log = gas.ss.getSheetByName('変更履歴_差分');
  assert.ok(gas.dump(log, 'D2:D' + log.getLastRow()).some(r => r[0] === '百貨店'), '足したシートの変更を記録する');

  // サービスリクエストの候補
  const listOf = cell => {
    const range = cell.getDataValidation().getCriteriaValues()[0];
    return plain(gas.evaluate(range.getSheet(), range.getRow(), range.getColumn()))[0];
  };
  assert.deepStrictEqual(listOf(gas.ss.getSheetByName('リクエスト').getRange('A2')), ['食品スーパー', '百貨店']);

  assert.deepStrictEqual(gas.triggers.map(t => t.getHandlerFunction()).sort(), triggers, 'トリガーは1つずつのまま');
  g.setupAfterSheetChange();
  assert.deepStrictEqual(gas.triggers.map(t => t.getHandlerFunction()).sort(), triggers, '何度実行してもよい');
});

test('setupAfterSheetChange：見出しが足りないシートがあれば、何も変えずに直す点をまとめて知らせる', () => {
  const {gas, g} = setupProject();
  const header = new Array(WIDTH).fill('');
  header[COL.customer - 1] = '得意先';   // 最終更新日時・ボタン列・リクエストの見出しが無い
  addDepartmentStore(gas, header);
  gas.get('DIFF_RULES')['専門店'] = {headerRow: 2, ranges: ['D3:AF']};   // シートが無い

  g.setupAfterSheetChange();
  assert.strictEqual(gas.alerts.length, 1);
  const message = gas.alerts[0].message;
  assert.match(message, /まだ何も変えていません/);
  assert.match(message, /百貨店 の 2行目に「最終更新日時」の見出しがちょうど1つ必要です/);
  assert.match(message, /百貨店 の 2行目に「クレデンシャル/);
  assert.match(message, /百貨店 の 2行目に「変更履歴」/);
  assert.match(message, /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
  assert.match(message, /「専門店」シートがありません/);
  assert.deepStrictEqual(gas.writes, [], '何も書き込まない');
  assert.strictEqual(gas.ss.getSheetByName('リクエスト'), null);
});

test('setupAfterSheetChange：ボタン列が追跡範囲に入っていたら、何も変えずに知らせる', () => {
  const {gas, g} = setupProject();
  addDepartmentStore(gas);
  gas.get('DIFF_RULES')['百貨店'].ranges = ['D3:AI'];   // ボタン列（AH・AI）まで含めてしまった
  g.setupAfterSheetChange();
  assert.match(gas.alerts[0].message, /百貨店 の「クレデンシャル.*」列（AH列）が DIFF_RULES の ranges に含まれています/);
  assert.deepStrictEqual(gas.writes, []);
});
