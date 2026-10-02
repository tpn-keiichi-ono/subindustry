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
/** 範囲を参照し、範囲に無い値は入力できない入力規則なら「シート!範囲」 */
const rangeOf = cell => {
  const rule = ruleOf(cell);
  if (!rule || rule.getCriteriaType() !== 'VALUE_IN_RANGE' || rule.getAllowInvalid()) return null;
  const range = rule.getCriteriaValues()[0];
  return range.getSheet().getName() + '!' + range.getA1Notation();
};

/** シートでセルを編集したことにして、単純トリガーの onEdit を呼ぶ */
function edit(gas, sheetName, a1, value) {
  const sheet = gas.ss.getSheetByName(sheetName);
  gas.global.onEdit(gas.edit(sheet, a1, value));
  return sheet;
}

test('importServiceRequests：リクエスト シートを作り、新FMT のリクエストを追記する（サービスの一覧のシートは作らない）', () => {
  const {gas, g} = setup();
  g.importServiceRequests();
  assert.strictEqual(gas.ss.getSheetByName('サービス'), null);

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
  assert.strictEqual(request.getRange('E2').getDataValidation(), null, 'サービスは自由に入力する（プルダウンにしない）');
  assert.deepStrictEqual(listOf(request.getRange('F2')), ['未判断', 'サービス化検討', '棄却']);
  // 転記する4列は、手で変えると警告が出る（警告だけ）
  const protections = request.getProtections().filter(p => p.getDescription() === 'リクエスト：新FMT から自動で転記する列');
  assert.deepStrictEqual(protections.map(p => p.range.getA1Notation()), ['A2:A1000', 'B2:B1000', 'C2:C1000', 'D2:D1000']);
  assert.ok(protections.every(p => p.isWarningOnly()));
  g.importServiceRequests();
  assert.strictEqual(request.getProtections().length, 4, '2回目は保護を増やさない');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /2件追加し/);
});

test('importServiceRequests：もう一度実行しても重複させず、新しいリクエストだけを足す（判断・サービスは残す）', () => {
  const {gas, g, source} = setup();
  g.importServiceRequests();
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  gas.asUser(() => source.getRange('F4').setValue('什器の在庫を見える化したい'));

  g.importServiceRequests();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(gas.dump(request, 'A2:F4'), [
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', 'サービス化検討'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい', '', '未判断'],
    ['食品スーパー', 'B社', '店舗什器', '什器の在庫を見える化したい', '', '未判断']
  ]);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /1件追加し/);
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

test('新FMT でリクエストなどを編集すると、リクエスト シートに自動で転記する', () => {
  const {gas, g, source} = setup();
  const sourceEdit = (a1, value) => g.onEdit(gas.edit(source, a1, value));

  sourceEdit('C3', '食品スーパー');
  assert.strictEqual(gas.ss.getSheetByName('リクエスト'), null, 'リクエスト シートを用意するまでは何もしない');

  g.importServiceRequests();
  const request = gas.ss.getSheetByName('リクエスト');
  gas.writes.length = 0;
  sourceEdit('B3', 'メモを書く');
  assert.deepStrictEqual(gas.writes, [], '転記する列以外の編集では何もしない');

  // リクエストを書いた行は、末尾に足す
  sourceEdit('F4', '店舗の在庫を見たい');
  assert.deepStrictEqual(gas.dump(request, 'A4:F4')[0], ['食品スーパー', 'B社', '店舗什器', '店舗の在庫を見たい', '', '未判断']);

  // リクエストを書き換えると、同じ行を直す（サービス・判断を付けたあとなら注を付ける）
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  sourceEdit('F3', '会員の購買データを分析したい');
  assert.deepStrictEqual(gas.dump(request, 'D2:F2')[0], ['会員の購買データを分析したい', '会員分析基盤', 'サービス化検討']);
  assert.match(request.getRange('D2').getNote(), /^新FMT でリクエストが書き換えられました/);
  sourceEdit('F4', '店舗ごとの在庫を見たい');
  assert.strictEqual(gas.dump(request, 'D4')[0][0], '店舗ごとの在庫を見たい');
  assert.strictEqual(request.getRange('D4').getNote(), '', 'サービス・判断を付ける前なら注は付けない');

  // 得意先・案件名・サブインダストリーを変えても同じ行を直す
  sourceEdit('D4', 'B社（本社）');
  sourceEdit('E4', '店舗什器の入れ替え');
  sourceEdit('C4', 'ホームセンター');
  assert.deepStrictEqual(gas.dump(request, 'A4:D5'), [
    ['ホームセンター', 'B社（本社）', '店舗什器の入れ替え', '店舗ごとの在庫を見たい'],
    ['', '', '', '']
  ]);

  // リクエストを消すと、行は残して注を付ける
  sourceEdit('F4', '');
  assert.match(request.getRange('D4').getNote(), /^取り込み元に見つかりません/);

  // 貼り付けで何行も変えても転記する
  g.onEdit(gas.edit(source, 'C7:F7', [['ドラッグストア', 'E社', '店頭サイネージ', '売場の案内を変えたい']]));
  assert.deepStrictEqual(gas.dump(request, 'A5:D5')[0], ['ドラッグストア', 'E社', '店頭サイネージ', '売場の案内を変えたい']);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('新FMT で得意先を変えると、そのサービスのタスクの案件の選択肢も変わる', () => {
  const {gas, g, source} = setup();
  g.importServiceRequests();
  g.setupTaskSheet();
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  const task = edit(gas, 'タスク管理', 'A2', '会員分析基盤');
  const listOfTask = () => listOf(task.getRange('C2'));
  assert.deepStrictEqual(listOfTask(), ['A社']);
  g.onEdit(gas.edit(source, 'D3', 'A社（本社）'));
  assert.deepStrictEqual(gas.dump('リクエスト', 'B2')[0][0], 'A社（本社）');
  assert.deepStrictEqual(listOfTask(), ['A社（本社）']);
});

test('リクエストにサービス名を付けると、判断が未判断なら「サービス化検討」にする（棄却などはそのまま）', () => {
  const {gas, g} = setup();
  g.importServiceRequests();
  edit(gas, 'リクエスト', 'F3', '棄却');
  const request = edit(gas, 'リクエスト', 'E2:E3', [['会員分析基盤'], ['会員分析基盤']]);
  assert.deepStrictEqual(gas.dump(request, 'F2:F3'), [['サービス化検討'], ['棄却']]);
});

test('タスク管理の「サービス」は、リクエストに付けたサービス名から選ぶ（古い形のシートは右端に列を足す）', () => {
  const {gas, g} = setup();
  // 前の版で作った、サービスの列が無いタスク管理シート
  gas.addSheet('タスク管理', [['サブインダストリー', '得意先', '案件名', 'タスク'], ['食品スーパー', 'A社', 'アプリ刷新', '提案書']],
    {rows: 10, columns: 8});
  g.setupRequestSheet();
  const task = gas.ss.getSheetByName('タスク管理');
  assert.deepStrictEqual(gas.dump(task, 'A1:I2'), [
    ['サブインダストリー', '得意先', '案件名', 'タスク', 'サービス', '担当者', '期限', '状況', 'メモ'],
    ['食品スーパー', 'A社', 'アプリ刷新', '提案書', '', '', '', '', '']
  ]);
  assert.strictEqual(rangeOf(task.getRange('E2')), 'リクエスト!E2:E1000', 'リクエストに無い名前は入れられない');
  assert.match(gas.toasts[0].message, /タスク管理 の右端に「サービス」「担当者」「期限」「状況」「メモ」の列を足しました/);

  // リクエスト シートがあれば、setupTaskSheet もサービスの列にプルダウンを付ける
  gas.asUser(() => task.getRange('E2:E10').clearDataValidations());
  g.setupTaskSheet();
  assert.strictEqual(rangeOf(task.getRange('E2')), 'リクエスト!E2:E1000');
  assert.doesNotMatch(gas.toasts[gas.toasts.length - 1].message, /importServiceRequests/);
});

test('リクエスト シートが無いうちは、タスク管理のサービスの列にプルダウンを付けずに知らせる', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  assert.strictEqual(gas.ss.getSheetByName('タスク管理').getRange('A2').getDataValidation(), null);
  assert.match(gas.toasts[0].message, /importServiceRequests\(\) を実行してください/);
});

test('タスクでサービスを選ぶと、案件の選択肢がそのサービスにまとめたリクエストの案件に絞られる（棄却は除く）', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['食品スーパー', 'B社', '店舗什器', '什器の在庫を見える化したい'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '会員データを活用したい'],
    ['ドラッグストア', 'D社', '物流見直し', '在庫を店舗と共有したい']
  ]);
  g.importServiceRequests();
  g.setupTaskSheet();
  // リクエストをもとにサービスを考え、同じ名前を付けて複数のリクエストをまとめる
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  edit(gas, 'リクエスト', 'E4', '会員分析基盤');
  edit(gas, 'リクエスト', 'E3', '在庫共有');
  edit(gas, 'リクエスト', 'E5', '在庫共有');
  edit(gas, 'リクエスト', 'F5', '棄却');

  let task = edit(gas, 'タスク管理', 'A2', '会員分析基盤');
  assert.deepStrictEqual(listOf(task.getRange('B2')), ['食品スーパー', 'ドラッグストア']);
  assert.deepStrictEqual(listOf(task.getRange('C2')), ['A社', 'C社']);
  assert.deepStrictEqual(listOf(task.getRange('D2')), ['アプリ刷新', 'EC立ち上げ']);
  assert.strictEqual(rangeOf(task.getRange('A2')), 'リクエスト!E2:E1000');

  task = edit(gas, 'タスク管理', 'A3', '在庫共有');
  assert.deepStrictEqual(listOf(task.getRange('C3')), ['B社'], '棄却したリクエストの案件は出さない');

  task = edit(gas, 'タスク管理', 'A4', 'まとめ直して無くなったサービス');
  assert.deepStrictEqual(listOf(task.getRange('C4')), ['A社', 'B社', 'C社', 'D社'], 'リクエストに無いサービス名なら新FMT のすべての案件');

  // サービスを選び直すと、合わなくなった案件は空にする
  edit(gas, 'タスク管理', 'B2', 'ドラッグストア');
  edit(gas, 'タスク管理', 'C2', 'C社');
  edit(gas, 'タスク管理', 'D2', 'EC立ち上げ');
  edit(gas, 'タスク管理', 'E2', '要件を聞く');
  task = edit(gas, 'タスク管理', 'A2', '在庫共有');
  assert.deepStrictEqual(gas.dump(task, 'A2:E2')[0], ['在庫共有', '', '', '', '要件を聞く']);
  assert.deepStrictEqual(listOf(task.getRange('C2')), ['B社']);

  // リクエストのまとめ方を変えると、タスクの行を触らなくても選択肢が変わる
  edit(gas, 'リクエスト', 'E4', '在庫共有');
  assert.deepStrictEqual(listOf(task.getRange('C2')), ['B社', 'C社']);
  edit(gas, 'リクエスト', 'F5', 'サービス化検討');
  assert.deepStrictEqual(listOf(task.getRange('C3')), ['B社', 'C社', 'D社'], '棄却をやめたリクエストの案件が戻る');
});

test('リクエストの列の見出しが無ければ、取り込まずに知らせる', () => {
  const {gas, g, source} = setup();
  gas.asUser(() => source.getRange('F2').setValue('要望'));
  assert.throws(() => g.importServiceRequests(), /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
});
