'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/** 新FMT（2行目が見出し。C: サブインダストリー / D: 得意先 / E: 案件名） */
const SOURCE = [
  ['食品スーパー', 'A社', 'アプリ刷新'],
  ['食品スーパー', 'A社', 'チラシのデジタル化'],
  ['食品スーパー', 'B社', '店舗什器'],
  ['ドラッグストア', 'C社', 'EC立ち上げ'],
  ['ドラッグストア', 'Ａ社', '物流見直し'],   // 全角の「Ａ社」（別のサブインダストリー）
  ['食品スーパー', 'Ａ社', '会員分析'],       // 全角の「Ａ社」は A社 と同じ得意先として扱う
  ['', '', '']
];

function setup(rows) {
  const gas = createGas();
  const width = 6;
  const title = ['新FMT', '', '', '', '', ''];
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', '状況'];
  gas.addSheet('新FMT', [title, header].concat((rows || SOURCE).map((r, i) => [String(i + 1), '', r[0], r[1], r[2], ''])),
    {rows: Math.max(20, (rows || SOURCE).length + 5), columns: width});
  const g = gas.global;
  return {gas, g};
}

const listOf = cell => {
  const rule = cell.getDataValidation();
  return rule && rule.getCriteriaType() === 'VALUE_IN_LIST' ? plain(rule.getCriteriaValues()[0]) : null;
};

/** タスク管理シートでセルを選んだことにして、単純トリガーの onEdit を呼ぶ */
function pick(gas, a1, value) {
  const sheet = gas.ss.getSheetByName('タスク管理');
  gas.global.onEdit(gas.edit(sheet, a1, value));
  return sheet;
}

test('setupTaskSheet：シート・見出しを作り、全行に連動プルダウン（何も選んでいなければすべての値）を付ける', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  assert.ok(sheet, 'タスク管理シートがありません');
  assert.deepStrictEqual(gas.dump(sheet, 'A1:H1')[0], ['サブインダストリー', '得意先', '案件名', 'タスク', '担当者', '期限', '状況', 'メモ']);
  assert.strictEqual(sheet.getFrozenRows(), 1);

  assert.deepStrictEqual(listOf(sheet.getRange('A2')), ['食品スーパー', 'ドラッグストア']);
  assert.deepStrictEqual(listOf(sheet.getRange('B2')), ['A社', 'B社', 'C社'], '全角・半角の違いは1つにまとめる');
  assert.deepStrictEqual(listOf(sheet.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化', '店舗什器', 'EC立ち上げ', '物流見直し', '会員分析']);
  assert.deepStrictEqual(listOf(sheet.getRange('C' + sheet.getMaxRows())), listOf(sheet.getRange('C2')), '最後の行まで付ける');
  assert.deepStrictEqual(listOf(sheet.getRange('G2')), ['未着手', '対応中', '完了']);
  assert.strictEqual(sheet.getRange('F2').getDataValidation().getCriteriaType(), 'DATE_IS_VALID_DATE');
  assert.strictEqual(sheet.getRange('D2').getDataValidation(), null, 'タスクの列は自由に入力できる');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /作りました/);
});

test('サブインダストリーを選ぶと、その行の得意先・案件名がそのサブインダストリーのものに絞られる', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  const sheet = pick(gas, 'A3', 'ドラッグストア');
  assert.deepStrictEqual(listOf(sheet.getRange('B3')), ['C社', 'Ａ社']);
  assert.deepStrictEqual(listOf(sheet.getRange('C3')), ['EC立ち上げ', '物流見直し']);
  assert.deepStrictEqual(listOf(sheet.getRange('B2')), ['A社', 'B社', 'C社'], 'ほかの行は変わらない');
});

test('得意先を選ぶと、案件名がサブインダストリーと得意先に合うものに絞られる（全角・半角の違いは無視）', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  pick(gas, 'A2', '食品スーパー');
  const sheet = pick(gas, 'B2', 'A社');
  assert.deepStrictEqual(listOf(sheet.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化', '会員分析']);
  // サブインダストリーが空なら、得意先だけで絞る（ドラッグストアの「Ａ社」の案件も入る）
  const other = pick(gas, 'B4', 'A社');
  assert.deepStrictEqual(listOf(other.getRange('C4')), ['アプリ刷新', 'チラシのデジタル化', '物流見直し', '会員分析']);
});

test('サブインダストリーを選び直すと、合わなくなった得意先・案件名を空にする（合うものは残す）', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  pick(gas, 'A2', '食品スーパー');
  pick(gas, 'B2', 'B社');
  pick(gas, 'C2', '店舗什器');
  pick(gas, 'D2', '提案書を送る');

  let sheet = pick(gas, 'A2', 'ドラッグストア');
  assert.deepStrictEqual(gas.dump(sheet, 'A2:D2')[0], ['ドラッグストア', '', '', '提案書を送る'], 'タスクの列は消さない');
  assert.deepStrictEqual(listOf(sheet.getRange('B2')), ['C社', 'Ａ社']);

  pick(gas, 'B2', 'C社');
  pick(gas, 'C2', 'EC立ち上げ');
  sheet = pick(gas, 'B2', 'Ａ社');
  assert.deepStrictEqual(gas.dump(sheet, 'A2:C2')[0], ['ドラッグストア', 'Ａ社', ''], '得意先を選び直すと案件名だけを空にする');
  assert.deepStrictEqual(listOf(sheet.getRange('C2')), ['物流見直し']);

  pick(gas, 'C2', '物流見直し');
  sheet = pick(gas, 'A2', '');
  assert.deepStrictEqual(gas.dump(sheet, 'A2:C2')[0], ['', 'Ａ社', '物流見直し'], 'サブインダストリーを消しても、合う値は残す');
  assert.deepStrictEqual(listOf(sheet.getRange('B2')), ['Ａ社', 'B社', 'C社'], '選択肢は入っている値の書き方に合わせる');
});

test('タスクの列だけを編集した行にもプルダウンを付ける（新しく足した行など）。値は空にしない', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  gas.asUser(() => sheet.getRange('A5:C5').setValues([['食品スーパー', 'C社', 'EC立ち上げ']]));   // 合わない値が入っている
  gas.asUser(() => sheet.getRange('A5:C5').clearDataValidations());
  pick(gas, 'D5', '電話する');
  assert.deepStrictEqual(gas.dump(sheet, 'A5:C5')[0], ['食品スーパー', 'C社', 'EC立ち上げ']);
  assert.deepStrictEqual(listOf(sheet.getRange('B5')), ['A社', 'B社']);
});

test('貼り付けなどで複数行を一度に変えたときも、行ごとに絞り込む', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  const sheet = pick(gas, 'A2:B3', [['食品スーパー', 'C社'], ['ドラッグストア', 'C社']]);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:B3'), [['食品スーパー', ''], ['ドラッグストア', 'C社']]);
  assert.deepStrictEqual(listOf(sheet.getRange('C3')), ['EC立ち上げ']);
});

test('ほかのシートの編集では何もしない', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  gas.writes.length = 0;
  const source = gas.ss.getSheetByName('新FMT');
  g.onEdit(gas.edit(source, 'D3', 'Z社'));
  assert.deepStrictEqual(gas.writes, []);
  g.onEdit(gas.edit(gas.ss.getSheetByName('タスク管理'), 'A1', 'サブインダストリー'));   // 見出しの行
  assert.deepStrictEqual(gas.writes, []);
});

test('新FMT に値が増えたら、次に編集した行・setupTaskSheet でプルダウンに入る。入力済みのタスクは消さない', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  pick(gas, 'D2', '既存のタスク');
  const source = gas.ss.getSheetByName('新FMT');
  gas.asUser(() => source.getRange('C9:E9').setValues([['ホームセンター', 'D社', '園芸売場']]));
  const sheet = pick(gas, 'A2', 'ホームセンター');
  assert.deepStrictEqual(listOf(sheet.getRange('A2')), ['食品スーパー', 'ドラッグストア', 'ホームセンター']);
  assert.deepStrictEqual(listOf(sheet.getRange('B2')), ['D社']);

  g.setupTaskSheet();
  assert.deepStrictEqual(listOf(sheet.getRange('A10')), ['食品スーパー', 'ドラッグストア', 'ホームセンター']);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:D2')[0], ['ホームセンター', '', '', '既存のタスク']);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /作り直しました/);
});

test('setupTaskSheet：見出しが違うシートは書き換えずに止める', () => {
  const {gas, g} = setup();
  gas.addSheet('タスク管理', [['項目', '内容']], {rows: 10, columns: 8});
  assert.throws(() => g.setupTaskSheet(), /「サブインダストリー」の見出しがちょうど1つ必要です/);
  assert.deepStrictEqual(gas.dump(gas.ss.getSheetByName('タスク管理'), 'A1:B1')[0], ['項目', '内容']);
});

test('ほかの処理がロックを持っているときは、変えずにトーストで知らせる', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  pick(gas, 'A2', '食品スーパー');
  pick(gas, 'B2', 'B社');
  gas.lockBusy = true;
  gas.writes.length = 0;
  gas.toasts.length = 0;
  const sheet = pick(gas, 'A2', 'ドラッグストア');
  assert.deepStrictEqual(gas.writes, []);
  assert.strictEqual(gas.dump(sheet, 'B2')[0][0], 'B社');
  assert.match(gas.toasts[0].message, /他の処理が実行中/);
});

test('選択肢が上限を超える列はプルダウンを付けずに知らせる', () => {
  const rows = [];
  for (let i = 0; i < 501; i++) rows.push(['食品スーパー', '得意先' + i, '案件' + i]);
  const {gas, g} = setup(rows);
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  assert.strictEqual(sheet.getRange('B2').getDataValidation(), null);
  assert.deepStrictEqual(listOf(sheet.getRange('A2')), ['食品スーパー']);
  assert.match(gas.toasts[0].message, /「得意先」の選択肢が 500 件を超える/);
});
