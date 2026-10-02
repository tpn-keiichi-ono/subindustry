'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/** 新FMT（2行目が見出し。C: サブインダストリー / D: 得意先 / E: 案件名 / F: サービスのリクエスト） */
const SOURCE = [
  ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
  ['食品スーパー', 'B社', '店舗什器', ''],
  ['ドラッグストア', 'C社', 'EC立ち上げ', "'=在庫を店舗と共有したい"],   // 文字列の「=…」（取り込みでも数式にしない）
  ['ドラッグストア', 'Ｃ社', 'EC立ち上げ', "'=在庫を店舗と共有したい"]    // 全角の得意先でも同じリクエスト
];

function setup(rows) {
  const gas = createGas();
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'];
  gas.addSheet('新FMT', [['新FMT'], header].concat((rows || SOURCE).map((r, i) => [String(i + 1), ''].concat(r))),
    {rows: 30, columns: 6});
  return {gas, g: gas.global, source: gas.ss.getSheetByName('新FMT')};
}

const ruleOf = cell => cell.getDataValidation();
const listOf = cell => {
  const rule = ruleOf(cell);
  return rule && rule.getCriteriaType() === 'VALUE_IN_LIST' ? plain(rule.getCriteriaValues()[0]) : null;
};
const rangeOf = cell => {
  const rule = ruleOf(cell);
  if (!rule || rule.getCriteriaType() !== 'VALUE_IN_RANGE') return null;
  const range = rule.getCriteriaValues()[0];
  return range.getSheet().getName() + '!' + range.getA1Notation();
};

/** シートでセルを編集したことにして、単純トリガーの onEdit を呼ぶ */
function edit(gas, sheetName, a1, value) {
  const sheet = gas.ss.getSheetByName(sheetName);
  gas.global.onEdit(gas.edit(sheet, a1, value));
  return sheet;
}

test('importServiceRequests：サービス・リクエストのシートを作り、新FMT のリクエストを追記する', () => {
  const {gas, g} = setup();
  g.importServiceRequests();

  const service = gas.ss.getSheetByName('サービス');
  assert.deepStrictEqual(gas.dump(service, 'A1:E1')[0], ['サービス名', '状況', '概要', '担当者', 'メモ']);
  assert.deepStrictEqual(listOf(service.getRange('B2')), ['検討中', '開発中', '提供中', '見送り']);

  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(gas.dump(request, 'A1:H1')[0],
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス', '判断', 'メモ', '取り込み日']);
  const today = gas.global.Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  assert.deepStrictEqual(gas.dump(request, 'A2:H4'), [
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '未判断', '', today],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい', '', '未判断', '', today],
    ['', '', '', '', '', '', '', '']
  ], 'リクエストが空の行は取り込まない。全角・半角だけが違う同じリクエストは1件');
  assert.strictEqual(request.getRange('D3').getFormula(), '', '数式にしない');
  assert.strictEqual(rangeOf(request.getRange('E2')), 'サービス!A2:A1000', 'サービスはサービス シートのサービス名から選ぶ');
  assert.deepStrictEqual(listOf(request.getRange('F2')), ['未判断', 'サービス化検討', '棄却']);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /2件追加しました/);
});

test('importServiceRequests：もう一度実行しても重複させず、新しいリクエストだけを足す（判断・サービスは残す）', () => {
  const {gas, g, source} = setup();
  g.importServiceRequests();
  edit(gas, 'サービス', 'A2', '会員分析基盤');
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  gas.asUser(() => source.getRange('F4').setValue('什器の在庫を見える化したい'));

  g.importServiceRequests();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(gas.dump(request, 'A2:F4'), [
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', 'サービス化検討'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい', '', '未判断'],
    ['食品スーパー', 'B社', '店舗什器', '什器の在庫を見える化したい', '', '未判断']
  ]);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /1件追加しました/);
});

test('importServiceRequests：取り込み元から消えた・書き換えられたリクエストは行を残して注を付け、戻ったら外す', () => {
  const {gas, g, source} = setup();
  g.importServiceRequests();
  const request = gas.ss.getSheetByName('リクエスト');
  gas.asUser(() => request.getRange('D3').setNote('担当者に確認中'));   // 利用者が書いた注
  gas.asUser(() => source.getRange('F3').setValue('会員の購買データを分析したい'));   // 書き換え
  gas.asUser(() => source.getRange('F5:F6').setValues([[''], ['']]));               // 削除

  g.importServiceRequests();
  assert.strictEqual(gas.dump(request, 'D4')[0][0], '会員の購買データを分析したい', '書き換えたものは新しい行として足す');
  assert.match(request.getRange('D2').getNote(), /^取り込み元に見つかりません/);
  assert.strictEqual(request.getRange('D3').getNote(), '担当者に確認中', '利用者の注は変えない');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /見つからないリクエストが 2件/);

  gas.asUser(() => source.getRange('F3').setValue('会員の購買分析をしたい'));
  g.importServiceRequests();
  assert.strictEqual(request.getRange('D2').getNote(), '', '取り込み元に戻ったら注を外す');
});

test('リクエストでサービスを選ぶと、判断が未判断なら「サービス化検討」にする（棄却などはそのまま）', () => {
  const {gas, g} = setup();
  g.importServiceRequests();
  edit(gas, 'サービス', 'A2', '会員分析基盤');
  edit(gas, 'リクエスト', 'F3', '棄却');
  const request = edit(gas, 'リクエスト', 'E2:E3', [['会員分析基盤'], ['会員分析基盤']]);
  assert.deepStrictEqual(gas.dump(request, 'F2:F3'), [['サービス化検討'], ['棄却']]);
});

test('タスク管理シートにも「サービス」の列ができ、サービス名から選べる（古い形のシートは右端に足す）', () => {
  const {gas, g} = setup();
  // サービス管理より前に作った、サービスの列が無いタスク管理シート
  gas.addSheet('タスク管理', [['サブインダストリー', '得意先', '案件名', 'タスク'], ['食品スーパー', 'A社', 'アプリ刷新', '提案書']],
    {rows: 10, columns: 8});
  g.setupServiceSheets();
  const task = gas.ss.getSheetByName('タスク管理');
  assert.deepStrictEqual(gas.dump(task, 'A1:I2'), [
    ['サブインダストリー', '得意先', '案件名', 'タスク', 'サービス', '担当者', '期限', '状況', 'メモ'],
    ['食品スーパー', 'A社', 'アプリ刷新', '提案書', '', '', '', '', '']
  ]);
  assert.strictEqual(rangeOf(task.getRange('E2')), 'サービス!A2:A1000');
  assert.match(gas.toasts[0].message, /タスク管理 の右端に「サービス」「担当者」「期限」「状況」「メモ」の列を足しました/);

  // サービス シートがあれば、setupTaskSheet もサービスの列にプルダウンを付ける
  gas.asUser(() => task.getRange('E2:E10').clearDataValidations());
  g.setupTaskSheet();
  assert.strictEqual(rangeOf(task.getRange('E2')), 'サービス!A2:A1000');
  assert.doesNotMatch(gas.toasts[gas.toasts.length - 1].message, /setupServiceSheets/);
});

test('サービス名を変えると、リクエスト・タスク管理の同じ名前も変わる', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  g.importServiceRequests();
  edit(gas, 'サービス', 'A2', '会員分析基盤');
  edit(gas, 'サービス', 'A3', '在庫共有');
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  edit(gas, 'リクエスト', 'E3', '在庫共有');
  edit(gas, 'タスク管理', 'D2', '会員分析基盤');

  edit(gas, 'サービス', 'A2', '会員データ分析');
  assert.deepStrictEqual(gas.dump('リクエスト', 'E2:E3'), [['会員データ分析'], ['在庫共有']]);
  assert.strictEqual(gas.dump('タスク管理', 'D2')[0][0], '会員データ分析');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /リクエスト 1件・タスク 1件に反映しました/);
});

test('同じ名前のサービスを作ると知らせ、名前の変更は反映しない', () => {
  const {gas, g} = setup();
  g.importServiceRequests();
  edit(gas, 'サービス', 'A2', '会員分析基盤');
  edit(gas, 'サービス', 'A3', '在庫共有');
  edit(gas, 'リクエスト', 'E3', '在庫共有');
  edit(gas, 'サービス', 'A3', '会員分析基盤');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /同じ名前のサービスが2つ以上あります（会員分析基盤）/);
  assert.strictEqual(gas.dump('リクエスト', 'E3')[0][0], '在庫共有');
});

test('リクエストの列の見出しが無ければ、取り込まずに知らせる', () => {
  const {gas, g, source} = setup();
  gas.asUser(() => source.getRange('F2').setValue('要望'));
  assert.throws(() => g.importServiceRequests(), /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
});
