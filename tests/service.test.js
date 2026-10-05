'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/** 新FMT（2行目が見出し。C: サブインダストリー / D: 得意先 / E: 案件名 / F: サービスのリクエスト） */
const SOURCE = [
  ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
  ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい'],
  ['食品スーパー', 'B社', '店舗什器', ''],                                   // リクエストが無い → 候補に出さない
  ['ドラッグストア', 'C社', 'EC立ち上げ', "'=在庫を店舗と共有したい"],      // 文字列の「=…」（数式にしない）
  ['ドラッグストア', 'Ｃ社', 'EC立ち上げ', "'=在庫を店舗と共有したい"]       // 全角だけが違う同じリクエスト
];

function setup(rows) {
  const gas = createGas();
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'];
  gas.addSheet('新FMT', [['新FMT'], header].concat((rows || SOURCE).map((r, i) => [String(i + 1), ''].concat(r))),
    {rows: 30, columns: 6});
  return {gas, g: gas.global, source: gas.ss.getSheetByName('新FMT')};
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

/** シートでセルを編集したことにして、単純トリガーの onEdit を呼ぶ */
function edit(gas, sheetName, a1, value) {
  const sheet = gas.ss.getSheetByName(sheetName);
  gas.global.onEdit(gas.edit(sheet, a1, value));
  return sheet;
}

const today = gas => gas.global.Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

test('setupRequestSheet：リクエスト シートを作る。行は足さず、A〜D のプルダウンにリクエストのある案件だけを自動で候補にする', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(gas.dump(request, 'A1:H1')[0],
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス案', '判断', 'メモ', '追加日']);
  assert.strictEqual(request.getLastRow(), 1, '行は自動では足さない');

  assert.deepStrictEqual(listOf(request.getRange('A2')), ['食品スーパー', 'ドラッグストア']);
  assert.deepStrictEqual(listOf(request.getRange('B2')), ['A社', 'C社'], 'リクエストの無い B社 は出さない。全角・半角の違いは1つ');
  assert.deepStrictEqual(listOf(request.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化', 'EC立ち上げ']);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい', 'チラシの効果を測りたい', '=在庫を店舗と共有したい']);
  assert.deepStrictEqual(listOf(request.getRange('D1000')), listOf(request.getRange('D2')), '最後の行まで付ける');
  assert.strictEqual(request.getRange('E2').getDataValidation(), null, 'サービス案は自由に入力する');
  assert.deepStrictEqual(listOf(request.getRange('F2')), ['未判断', 'サービス化検討', '棄却']);

  assert.strictEqual(request.getProtections().length, 0, 'A〜D はプルダウンで選ぶので保護しない');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /まだ選んでいないリクエストが 3件あります/);
});

test('A から順に選ぶと、1つに決まる列は自動で入り、選んだ案件はほかの行の候補から消える', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'A2', '食品スーパー');
  assert.strictEqual(gas.dump(request, 'B2')[0][0], 'A社', '食品スーパーの得意先は A社 だけなので入る');
  assert.deepStrictEqual(listOf(request.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化']);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい', 'チラシの効果を測りたい']);
  request = edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  assert.deepStrictEqual(gas.dump(request, 'A2:H2')[0],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '未判断', '', today(gas)]);

  assert.deepStrictEqual(listOf(request.getRange('C3')), ['チラシのデジタル化', 'EC立ち上げ'], 'ほかの行の候補から消える');
  assert.deepStrictEqual(listOf(request.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化'], '自分の行には残る');
});

test('案件名やリクエストを先に選ぶと、ほかの列も入る。すべて選ぶと候補は「（すべて選択済み）」だけになる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'C2', 'EC立ち上げ');
  assert.deepStrictEqual(gas.dump(request, 'A2:D2')[0], ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい']);
  assert.strictEqual(request.getRange('D2').getFormula(), '', '数式にしない');

  request = edit(gas, 'リクエスト', 'D3', 'チラシの効果を測りたい');
  assert.deepStrictEqual(gas.dump(request, 'A3:D3')[0], ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  request = edit(gas, 'リクエスト', 'C4', 'アプリ刷新');
  ['A5', 'B5', 'C5', 'D5', 'A20'].forEach(a1 => assert.deepStrictEqual(listOf(request.getRange(a1)), ['（すべて選択済み）'], a1));

  gas.toasts.length = 0;
  request = edit(gas, 'リクエスト', 'A5', '（すべて選択済み）');
  assert.strictEqual(gas.dump(request, 'A5')[0][0], '', '選んでも入らない');
  assert.match(gas.toasts[0].message, /すべて選んでいます/);
  g.setupRequestSheet();
  assert.match(gas.toasts[gas.toasts.length - 1].message, /リクエストはすべて選んでいます/);
});

test('左の列を選び直すと、合わなくなった右の列とリクエストを空にし、その案件はまた候補に戻る', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  edit(gas, 'リクエスト', 'C3', 'チラシのデジタル化');
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  assert.deepStrictEqual(listOf(gas.ss.getSheetByName('リクエスト').getRange('A2')), ['食品スーパー', 'ドラッグストア']);

  const request = edit(gas, 'リクエスト', 'A2', 'ドラッグストア');
  assert.deepStrictEqual(gas.dump(request, 'A2:F2')[0],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい', '会員分析基盤', 'サービス化検討'],
    '合わなくなった得意先・案件名・リクエストを空にし、ドラッグストアの候補は1つなので入れる。サービス案・判断は消さない');
  assert.deepStrictEqual(listOf(request.getRange('C4')), ['アプリ刷新'], 'アプリ刷新が候補に戻る');
});

test('貼り付けなどで同じ案件を2つの行に入れたら、あとの行は空にして知らせる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  gas.toasts.length = 0;
  const request = edit(gas, 'リクエスト', 'A3:C3', [['食品スーパー', 'A社', 'アプリ刷新']]);
  assert.deepStrictEqual(gas.dump(request, 'A3:D3')[0], ['食品スーパー', 'A社', '', '']);
  assert.match(gas.toasts[0].message, /3行目の「アプリ刷新」は候補に無いため、空にしました/);
});

test('候補に無いリクエストを入力したら空にして知らせる（ほかの候補で勝手に埋めない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  gas.toasts.length = 0;
  let request = edit(gas, 'リクエスト', 'D2', '書き換え');
  assert.deepStrictEqual(gas.dump(request, 'C2:D2')[0], ['アプリ刷新', '']);
  assert.match(gas.toasts[0].message, /2行目の「書き換え」は候補に無いため、空にしました/);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい']);
  request = edit(gas, 'リクエスト', 'D2', '会員の購買分析をしたい');
  assert.strictEqual(gas.dump(request, 'D2')[0][0], '会員の購買分析をしたい');
});

test('同じ案件にリクエストが2つあれば、案件名を選んでもリクエストは入れず、プルダウンで選ぶ', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['食品スーパー', 'A社', 'アプリ刷新', 'クーポンを配りたい']
  ]);
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  assert.deepStrictEqual(gas.dump(request, 'A2:D2')[0], ['食品スーパー', 'A社', 'アプリ刷新', '']);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい', 'クーポンを配りたい']);
  request = edit(gas, 'リクエスト', 'D2', 'クーポンを配りたい');
  assert.deepStrictEqual(listOf(request.getRange('D3')), ['会員の購買分析をしたい']);
  request = edit(gas, 'リクエスト', 'C3', 'アプリ刷新');
  assert.strictEqual(gas.dump(request, 'D3')[0][0], '会員の購買分析をしたい', '残りが1つなら入る');
});

test('リクエストにサービス案を付けると、判断が未判断なら「サービス化検討」にする（棄却などはそのまま）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  edit(gas, 'リクエスト', 'C3', 'EC立ち上げ');
  edit(gas, 'リクエスト', 'F3', '棄却');
  const request = edit(gas, 'リクエスト', 'E2:E3', [['会員分析基盤'], ['会員分析基盤']]);
  assert.deepStrictEqual(gas.dump(request, 'F2:F3'), [['サービス化検討'], ['棄却']]);
});

test('新FMT の変更：選んだ行は同じ行を直し、候補も作り直す（行は足さない）', () => {
  const {gas, g, source} = setup();
  const sourceEdit = (a1, value) => g.onEdit(gas.edit(source, a1, value));
  sourceEdit('C3', '食品スーパー');
  assert.strictEqual(gas.ss.getSheetByName('リクエスト'), null, 'リクエスト シートを用意するまでは何もしない');

  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  gas.writes.length = 0;
  sourceEdit('B3', 'メモを書く');
  assert.deepStrictEqual(gas.writes, [], '読む列以外の編集では何もしない');

  // リクエストを書いた行は、行は足さずに候補に入る
  sourceEdit('F5', '店舗の在庫を見たい');
  assert.strictEqual(request.getLastRow(), 2);
  assert.deepStrictEqual(listOf(request.getRange('B3')), ['A社', 'B社', 'C社']);

  // 選んだ行のリクエスト・得意先・サブインダストリーを変えると、同じ行を直す（サービス案を付けたあとなら注を付ける）
  sourceEdit('F3', '会員の購買データを分析したい');
  assert.deepStrictEqual(gas.dump(request, 'D2:E2')[0], ['会員の購買データを分析したい', '会員分析基盤']);
  assert.match(request.getRange('D2').getNote(), /^新FMT でリクエストが書き換えられました/);
  sourceEdit('D3', 'A社（本社）');
  sourceEdit('C3', '大型スーパー');
  assert.deepStrictEqual(gas.dump(request, 'A2:C2')[0], ['大型スーパー', 'A社（本社）', 'アプリ刷新']);

  // リクエストを消すと、行は残して注を付ける
  sourceEdit('F3', '');
  assert.match(request.getRange('D2').getNote(), /^取り込み元に見つかりません/);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('前の版のシート：見出し（サービス・取り込み日）を書き換え、選んであった行と値はそのまま使う。前の版の保護は外す', () => {
  const {gas, g} = setup();
  const old = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス', '判断', 'メモ', '取り込み日'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', 'サービス化検討', '', '2026/10/01']
  ], {rows: 20, columns: 8});
  old.getRange('A2:A').protect().setDescription('リクエスト：新FMT から自動で転記する列').setWarningOnly(true);
  gas.addSheet('タスク管理', [['サービス案', 'タスク', '担当者', '期限', '状況', 'メモ']], {rows: 10, columns: 6});

  g.setupRequestSheet();
  assert.deepStrictEqual(gas.dump(old, 'A1:H2'), [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス案', '判断', 'メモ', '追加日'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', 'サービス化検討', '', '2026/10/01']
  ]);
  assert.strictEqual(old.getProtections().length, 0);
  assert.deepStrictEqual(listOf(old.getRange('C3')), ['チラシのデジタル化', 'EC立ち上げ'], '選んであった案件は候補から外す');
  assert.strictEqual(rangeOf(gas.ss.getSheetByName('タスク管理').getRange('A2')), 'リクエスト!E2:E20');
});

test('新FMT のリクエストの列の見出しが無ければ、止めて知らせる', () => {
  const {gas, g, source} = setup();
  gas.asUser(() => source.getRange('F2').setValue('要望'));
  assert.throws(() => g.setupRequestSheet(), /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
});
