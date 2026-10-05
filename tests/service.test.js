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
  gas.evaluateFormulas = true;   // 候補のシート（__REQUEST_LISTS）の数式を計算する
  return {gas, g: gas.global, source: gas.ss.getSheetByName('新FMT')};
}

/**
 * セルのプルダウンの候補。範囲を参照するプルダウン（候補のシートの欄）は、欄の先頭の数式を計算して横に並んだ値を返す
 * （欄に入りきらなければ、シートと同じく #REF!）。
 */
const listOf = cell => {
  const rule = cell.getDataValidation();
  if (!rule) return null;
  if (rule.getCriteriaType() === 'VALUE_IN_LIST') return plain(rule.getCriteriaValues()[0]);
  if (rule.getCriteriaType() !== 'VALUE_IN_RANGE') return null;
  const range = rule.getCriteriaValues()[0];
  const sheet = range.getSheet();
  const result = sheet.gas.evaluate(sheet, range.getRow(), range.getColumn());
  if (result && result.error) return [result.error];
  const values = Array.isArray(result) ? [].concat(...result) : [result];
  if (values.length > range.getNumColumns()) return ['#REF!'];
  return values.filter(v => v !== '' && v != null).map(String);
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

/** 行の A〜D を左から順にプルダウンで選ぶ（自動で入った列は選ばない） */
function choose(gas, row, values) {
  const sheet = gas.ss.getSheetByName('リクエスト');
  values.forEach((value, i) => {
    const a1 = 'ABCD'[i] + row;
    if (gas.dump(sheet, a1)[0][0] !== value) edit(gas, 'リクエスト', a1, value);
  });
  return sheet;
}

const WAIT_SUB = ['（先にサブインダストリーを選んでください）'];

test('setupRequestSheet：リクエスト シートを作る。行は足さず、A のプルダウンにリクエストのある案件だけを候補にする（B〜D は左の列を選んでから）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(gas.dump(request, 'A1:H1')[0],
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者']);
  assert.strictEqual(request.getLastRow(), 1, '行は自動では足さない');
  assert.strictEqual(gas.alerts.length, 0, '使わなくなった列が無ければ確認しない');

  assert.deepStrictEqual(listOf(request.getRange('A2')), ['食品スーパー', 'ドラッグストア']);
  ['B2', 'C2', 'D2', 'D1000'].forEach(a1 => assert.deepStrictEqual(listOf(request.getRange(a1)), WAIT_SUB, a1 + '：左の列が空なら選べない'));
  assert.match(request.getRange('C2').getDataValidation().getHelpText(), /案件名は一覧から選んでください（左の列から順に選びます/);
  assert.ok(gas.ss.getSheetByName('__REQUEST_LISTS').isSheetHidden(), '候補を作るシートは隠す');
  assert.deepStrictEqual(listOf(request.getRange('E2')), ['未判断', 'サービス化検討', '棄却']);
  ['F2', 'G2', 'H2'].forEach(a1 => assert.strictEqual(request.getRange(a1).getDataValidation(), null, a1 + '：フィードバック・サービス案・担当者は自由に入力する'));

  assert.strictEqual(request.getProtections().length, 0, 'A〜D はプルダウンで選ぶので保護しない');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /「新FMT」のリクエスト 3件のうち 0件を選んでいます（まだ選んでいないもの 3件）/);
});

test('A から順に選ぶと、1つに決まる得意先・案件名は自動で入り、選んだ案件はほかの行の候補から消える（D列以降は自動で入れない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'A2', '食品スーパー');
  assert.strictEqual(gas.dump(request, 'B2')[0][0], 'A社', '食品スーパーの得意先は A社 だけなので入る（リクエストの無い B社 は出さない）');
  assert.deepStrictEqual(listOf(request.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化']);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['（先に案件名を選んでください）']);
  request = edit(gas, 'リクエスト', 'C2', 'アプリ刷新');
  assert.deepStrictEqual(gas.dump(request, 'A2:H2')[0], ['食品スーパー', 'A社', 'アプリ刷新', '', '', '', '', ''],
    'リクエストが1つでも入れない（プルダウンで選ぶ）');
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい']);

  request = edit(gas, 'リクエスト', 'D2', '会員の購買分析をしたい');
  assert.deepStrictEqual(gas.dump(request, 'A2:H2')[0],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '', '', ''], '判断・追加日は入れない');

  request = edit(gas, 'リクエスト', 'A3', '食品スーパー');
  assert.deepStrictEqual(gas.dump(request, 'B3:D3')[0], ['A社', 'チラシのデジタル化', ''], 'アプリ刷新はほかの行の候補から消え、残りの案件名が1つなので入る');
  assert.deepStrictEqual(listOf(request.getRange('C2')), ['アプリ刷新', 'チラシのデジタル化'], '自分の行には残る');
});

test('A〜C を選ぶたびに、その行の右の列の候補が絞られる（候補は数式で作るので、選ぶたびにプルダウンを付け直さない）', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['食品スーパー', 'A社', 'アプリ刷新', 'クーポンを配りたい'],
    ['食品スーパー', 'B社', '店舗什器', '什器の在庫を見たい'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '在庫を店舗と共有したい']
  ]);
  g.setupRequestSheet();
  let request = gas.ss.getSheetByName('リクエスト');
  const event = gas.edit(request, 'A3', '食品スーパー');
  gas.writes.length = 0;
  gas.reads.length = 0;
  g.onEdit(event);
  assert.deepStrictEqual(gas.writes.filter(w => /Validation|Formula/.test(w.kind)), [], 'プルダウン・数式は付け直さない');
  assert.ok(!gas.reads.some(r => r.sheet === '新FMT'), '新FMT は読まない（速く終わるように）');
  assert.deepStrictEqual(gas.reads.map(r => r.sheet), ['リクエスト', '__REQUEST_LISTS', 'リクエスト'],
    '読むのは見出し・候補のシート・連動列の3回だけ');
  assert.deepStrictEqual(listOf(request.getRange('B3')), ['A社', 'B社']);
  assert.deepStrictEqual(listOf(request.getRange('C3')), ['（先に得意先を選んでください）']);
  assert.deepStrictEqual(listOf(request.getRange('D3')), ['（先に得意先を選んでください）']);
  request = edit(gas, 'リクエスト', 'B3', 'A社');
  assert.deepStrictEqual(gas.dump(request, 'C3:D3')[0], ['アプリ刷新', ''], '案件名は1つなので入る。リクエストは選ぶ');
  assert.deepStrictEqual(listOf(request.getRange('D3')), ['会員の購買分析をしたい', 'クーポンを配りたい']);
  assert.deepStrictEqual(listOf(request.getRange('A2')), ['食品スーパー', 'ドラッグストア'], 'ほかの行は絞らない');
  assert.deepStrictEqual(listOf(request.getRange('B2')), WAIT_SUB);
});

test('左の列が空のまま右の列に入力・貼り付けすると、空にして知らせる（案内の文字を選んでも入らない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'C3', WAIT_SUB[0]);
  assert.strictEqual(gas.dump(request, 'C3')[0][0], '');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /先にサブインダストリーを選んでください（左の列から順に選びます）/);

  request = edit(gas, 'リクエスト', 'B4:D4', [['A社', 'アプリ刷新', '会員の購買分析をしたい']]);
  assert.deepStrictEqual(gas.dump(request, 'A4:D4')[0], ['', '', '', ''], '左の列から埋めることもしない');
  assert.strictEqual(gas.toasts[gas.toasts.length - 1].message,
    '4行目の「A社」「アプリ刷新」「会員の購買分析をしたい」を空にしました（先にサブインダストリーを選んでください）。');

  choose(gas, 2, ['食品スーパー', 'A社', 'チラシのデジタル化']);
  request = edit(gas, 'リクエスト', 'D2', 'チラシの効果を測りたい');
  request = edit(gas, 'リクエスト', 'A4:C4', [['食品スーパー', '', 'アプリ刷新']]);
  assert.deepStrictEqual(gas.dump(request, 'A4:C4')[0], ['食品スーパー', '', ''], '消した列・空にした列は自動で埋めない');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /4行目の「アプリ刷新」を空にしました（先に得意先を選んでください）/);
});

test('左の列を消すと、右の列（リクエストまで）も空にして知らせる。消した列は自動で埋め直さない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');

  let request = edit(gas, 'リクエスト', 'B2', '');
  assert.deepStrictEqual(gas.dump(request, 'A2:E2')[0], ['食品スーパー', '', '', '', '会員分析基盤'], '得意先は1つでも埋め直さない。サービス案は残す');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /2行目の得意先が空になったため、案件名・リクエストも空にしました/);
  assert.deepStrictEqual(listOf(request.getRange('B2')), ['A社']);
  assert.deepStrictEqual(listOf(request.getRange('C2')), ['（先に得意先を選んでください）']);

  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  request = edit(gas, 'リクエスト', 'A2', '');
  assert.deepStrictEqual(gas.dump(request, 'A2:D2')[0], ['', '', '', '']);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /2行目のサブインダストリーが空になったため、得意先・案件名・リクエストも空にしました/);
  gas.toasts.length = 0;
  request = edit(gas, 'リクエスト', 'A2:D2', [['', '', '', '']]);
  assert.strictEqual(gas.toasts.length, 0, 'まとめて消したときは知らせない');
});

test('すべて選ぶと候補は「（すべて選択済み）」だけになる。リクエストの「=…」は数式にしない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'A2', 'ドラッグストア');
  assert.deepStrictEqual(gas.dump(request, 'A2:D2')[0], ['ドラッグストア', 'C社', 'EC立ち上げ', ''], '全角・半角だけが違う得意先は1つ');
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['=在庫を店舗と共有したい']);
  request = edit(gas, 'リクエスト', 'D2', "'=在庫を店舗と共有したい");
  assert.strictEqual(gas.dump(request, 'D2')[0][0], '=在庫を店舗と共有したい');
  assert.strictEqual(request.getRange('D2').getFormula(), '', '数式にしない');

  choose(gas, 3, ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  choose(gas, 4, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  ['A5', 'A20'].forEach(a1 => assert.deepStrictEqual(listOf(request.getRange(a1)), ['（すべて選択済み）'], a1));
  assert.deepStrictEqual(listOf(request.getRange('B5')), WAIT_SUB);

  gas.toasts.length = 0;
  request = edit(gas, 'リクエスト', 'A5', '（すべて選択済み）');
  assert.strictEqual(gas.dump(request, 'A5')[0][0], '', '選んでも入らない');
  assert.match(gas.toasts[0].message, /すべて選んでいます/);
  g.setupRequestSheet();
  assert.match(gas.toasts[gas.toasts.length - 1].message, /3件のうち 3件を選んでいます（まだ選んでいないもの 0件）/);
});

test('左の列を選び直すと、合わなくなった右の列とリクエストを空にし、その案件はまた候補に戻る', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  choose(gas, 3, ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  edit(gas, 'リクエスト', 'G2', '会員分析基盤');
  edit(gas, 'リクエスト', 'E2', 'サービス化検討');
  assert.deepStrictEqual(listOf(gas.ss.getSheetByName('リクエスト').getRange('A2')), ['食品スーパー', 'ドラッグストア']);

  let request = edit(gas, 'リクエスト', 'A2', 'ドラッグストア');
  assert.deepStrictEqual(gas.dump(request, 'A2:G2')[0],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '', 'サービス化検討', '', '会員分析基盤'],
    '合わなくなった得意先・案件名・リクエストを空にし、1つに決まる得意先・案件名だけを入れる。サービス案・判断は消さない');
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['=在庫を店舗と共有したい']);
  request = edit(gas, 'リクエスト', 'A4', '食品スーパー');
  assert.deepStrictEqual(gas.dump(request, 'A4:C4')[0], ['食品スーパー', 'A社', 'アプリ刷新'], 'アプリ刷新が候補に戻る');
});

test('貼り付けなどで同じ案件を2つの行に入れたら、あとの行は空にして知らせる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  gas.toasts.length = 0;
  const request = edit(gas, 'リクエスト', 'A3:C3', [['食品スーパー', 'A社', 'アプリ刷新']]);
  assert.deepStrictEqual(gas.dump(request, 'A3:D3')[0], ['食品スーパー', 'A社', '', ''], 'ほかの案件名で勝手に埋めない');
  assert.match(gas.toasts[0].message, /3行目の「アプリ刷新」は候補に無いため、空にしました/);
});

test('候補に無いリクエストを入力したら空にして知らせる（ほかの候補で勝手に埋めない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新']);
  gas.toasts.length = 0;
  let request = edit(gas, 'リクエスト', 'D2', '書き換え');
  assert.deepStrictEqual(gas.dump(request, 'C2:D2')[0], ['アプリ刷新', '']);
  assert.match(gas.toasts[0].message, /2行目の「書き換え」は候補に無いため、空にしました/);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい']);
  request = edit(gas, 'リクエスト', 'D2', '会員の購買分析をしたい');
  assert.strictEqual(gas.dump(request, 'D2')[0][0], '会員の購買分析をしたい');
});

test('案件名まで決まっても、リクエストは入れずプルダウンで選ぶ（残りが1つでも）', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['食品スーパー', 'A社', 'アプリ刷新', 'クーポンを配りたい']
  ]);
  g.setupRequestSheet();
  let request = edit(gas, 'リクエスト', 'A2', '食品スーパー');
  assert.deepStrictEqual(gas.dump(request, 'A2:D2')[0], ['食品スーパー', 'A社', 'アプリ刷新', '']);
  assert.deepStrictEqual(listOf(request.getRange('D2')), ['会員の購買分析をしたい', 'クーポンを配りたい']);
  request = edit(gas, 'リクエスト', 'D2', 'クーポンを配りたい');
  request = edit(gas, 'リクエスト', 'A3', '食品スーパー');
  assert.deepStrictEqual(gas.dump(request, 'A3:D3')[0], ['食品スーパー', 'A社', 'アプリ刷新', ''], '残りが1つでも入れない');
  assert.deepStrictEqual(listOf(request.getRange('D3')), ['会員の購買分析をしたい']);
});

test('判断・フィードバック・サービス案・担当者を入力しても、ほかの列は変えない（判断も自動で入れない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  choose(gas, 3, ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  edit(gas, 'リクエスト', 'E3', '棄却');
  gas.writes.length = 0;
  const request = edit(gas, 'リクエスト', 'G2:G3', [['会員分析基盤'], ['会員分析基盤']]);
  edit(gas, 'リクエスト', 'F2', '既存の分析基盤で対応できそう');
  edit(gas, 'リクエスト', 'H2', '佐藤');
  assert.deepStrictEqual(gas.writes, [], '何も書き込まない');
  assert.deepStrictEqual(gas.dump(request, 'E2:H3'), [
    ['', '既存の分析基盤で対応できそう', '会員分析基盤', '佐藤'],
    ['棄却', '', '会員分析基盤', '']
  ]);
});

test('新FMT でサブインダストリー〜案件名のどれかが空のリクエストは、選べないので候補に出さずに知らせる', () => {
  const {gas, g} = setup([
    ['', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['食品スーパー', 'B社', '店舗什器', '什器の在庫を見たい']
  ]);
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(listOf(request.getRange('A2')), ['食品スーパー']);
  assert.match(gas.toasts[0].message, /2件のうち 0件を選んでいます（まだ選んでいないもの 1件）/);
  assert.match(gas.toasts[0].message, /どれかが空のため、選べないリクエストが 1件あります/);
});

test('リクエスト シートの行を足す・消すなどでプルダウンの参照がずれたら、次に選んだときに作り直す', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  const rowOf = a1 => request.getRange(a1).getDataValidation().getCriteriaValues()[0].getRow();
  assert.strictEqual(rowOf('A2'), 2, 'n 行目のプルダウンは、候補のシートの n 行目を参照する');

  // 行を消して足す（行数は同じ）：消した行より下のプルダウンは、1つ下の行の候補を参照したままになる
  gas.asUser(() => { request.deleteRows(3, 1); request.insertRowsAfter(request.getMaxRows(), 1); });
  assert.strictEqual(rowOf('A3'), 4);
  edit(gas, 'リクエスト', 'A3', '食品スーパー');
  assert.strictEqual(rowOf('A3'), 3);
  assert.strictEqual(rowOf('A10'), 10);
  assert.deepStrictEqual(listOf(request.getRange('C3')), ['アプリ刷新', 'チラシのデジタル化']);

  // 行を足す：足した行にもプルダウンを付ける
  const max = request.getMaxRows();
  gas.asUser(() => request.insertRowsAfter(max, 5));
  assert.strictEqual(request.getRange('A' + (max + 3)).getDataValidation(), null);
  edit(gas, 'リクエスト', 'A2', 'ドラッグストア');
  assert.deepStrictEqual(listOf(request.getRange('A' + (max + 3))), ['食品スーパー', 'ドラッグストア']);
});

test('新FMT で候補が増えて欄に入りきらなくなったら、プルダウンと候補のシートを作り直す（入るうちは案件だけを書き直す）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  gas.writes.length = 0;
  edit(gas, '新FMT', 'C8:F8', [['食品スーパー', 'D社', '新店', '売場を見直したい']]);
  assert.deepStrictEqual(gas.writes.filter(w => w.kind === 'setDataValidations'), [], '欄に入るうちは作り直さない');

  const more = ['E社', 'F社', 'G社', 'H社', 'I社', 'J社'].map(c => ['食品スーパー', c, '新店', '売場を見直したい']);
  edit(gas, '新FMT', 'C9:F14', more);
  assert.ok(gas.writes.some(w => w.kind === 'setDataValidations'), '欄を広げて作り直す');
  edit(gas, 'リクエスト', 'A2', '食品スーパー');
  assert.deepStrictEqual(listOf(request.getRange('B2')), ['A社', 'D社', 'E社', 'F社', 'G社', 'H社', 'I社', 'J社']);
});

test('候補を作る数式がエラーになっていたら、setupRequestSheet() のあとに知らせる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  assert.strictEqual(g.svcCheckLists_(gas.ss).length, 0, 'エラーが無ければ何も言わない');
  gas.asUser(() => gas.ss.getSheetByName('__REQUEST_LISTS').getRange('G2').setValue('=NOSUCHFUNCTION()'));
  assert.match(g.svcCheckLists_(gas.ss)[0], /候補を作る数式でエラーが出ています（#NAME\?）/);
});

test('新FMT の変更：選んだ行は同じ行を直し、候補も作り直す（行は足さない）', () => {
  const {gas, g, source} = setup();
  const sourceEdit = (a1, value) => g.onEdit(gas.edit(source, a1, value));
  sourceEdit('C3', '食品スーパー');
  assert.strictEqual(gas.ss.getSheetByName('リクエスト'), null, 'リクエスト シートを用意するまでは何もしない');

  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  gas.writes.length = 0;
  sourceEdit('B3', 'メモを書く');
  assert.deepStrictEqual(gas.writes, [], '読む列以外の編集では何もしない');

  // リクエストを書いた行は、行は足さずに候補に入る
  sourceEdit('F5', '店舗の在庫を見たい');
  assert.strictEqual(request.getLastRow(), 2);
  edit(gas, 'リクエスト', 'A3', '食品スーパー');
  assert.deepStrictEqual(listOf(request.getRange('B3')), ['A社', 'B社']);

  // 選んだ行のリクエスト・得意先・サブインダストリーを変えると、同じ行を直す（サービス案を付けたあとなら注を付ける）
  sourceEdit('F3', '会員の購買データを分析したい');
  assert.deepStrictEqual(gas.dump(request, 'D2:E2')[0], ['会員の購買データを分析したい', '会員分析基盤']);
  assert.match(request.getRange('D2').getNote(), /^取り込み元でリクエストが書き換えられました/);
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
  assert.match(gas.alerts[0].message, /「メモ」「取り込み日」の列は使わなくなりました/);
  assert.deepStrictEqual(gas.dump(old, 'A1:H2'), [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス案', '担当者', '判断', 'サービス部門からのフィードバック'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', '', 'サービス化検討', '']
  ], 'メモ・取り込み日を削除し、担当者はサービス案の、フィードバックは判断のすぐ右に差し込む');
  assert.strictEqual(old.getProtections().length, 0);
  edit(gas, 'リクエスト', 'A3', '食品スーパー');
  assert.deepStrictEqual(gas.dump(old, 'B3:C3')[0], ['A社', 'チラシのデジタル化'], '選んであった案件は候補から外す（残りの1つが入る）');
  assert.strictEqual(rangeOf(gas.ss.getSheetByName('タスク管理').getRange('A2')), 'リクエスト!E2:E20');
});

/** 今のシートの並び（判断の右にサービス案。メモ・追加日がある） */
function addCurrentRequestSheet(gas) {
  return gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス案', 'メモ', '追加日'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '会員分析基盤', '先方に確認', '2026/10/01']
  ], {rows: 20, columns: 8});
}

test('setupRequestSheet：確認してからメモ・追加日を削除し、判断の右にフィードバック、サービス案の右に担当者を足す（ほかの値は残す）', () => {
  const {gas, g} = setup();
  const sheet = addCurrentRequestSheet(gas);
  sheet.getRange('E2:E20').setDataValidation(gas.global.SpreadsheetApp.newDataValidation()
    .requireValueInList(['未判断', 'サービス化検討', '棄却'], true).build());
  gas.addSheet('タスク管理', [['サービス案', 'タスク', '担当者', '期限', '状況', 'メモ']], {rows: 10, columns: 6});
  gas.confirmAnswer = 'YES';
  g.setupRequestSheet();

  assert.strictEqual(gas.alerts.length, 1);
  assert.match(gas.alerts[0].message, /「メモ」「追加日」の列は使わなくなりました。\n.*削除する列に入っている値は消えます/);
  assert.deepStrictEqual(gas.dump(sheet, 'A1:H2'), [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '']
  ]);
  assert.strictEqual(sheet.getLastColumn(), 8);
  assert.strictEqual(sheet.getRange('F2').getDataValidation(), null, '差し込んだ列は判断のプルダウンを引き継がない');
  assert.deepStrictEqual(listOf(sheet.getRange('E2')), ['未判断', 'サービス化検討', '棄却']);
  assert.strictEqual(rangeOf(gas.ss.getSheetByName('タスク管理').getRange('A2')), 'リクエスト!G2:G20', 'タスク管理はずれたサービス案の列を参照する');
  const toast = gas.toasts[gas.toasts.length - 1].message;
  assert.match(toast, /「メモ」「追加日」の列を削除しました。/);
  assert.match(toast, /「サービス部門からのフィードバック」「担当者」の列を足しました。/);

  gas.alerts.length = 0;
  g.setupRequestSheet();
  assert.strictEqual(gas.alerts.length, 0, '2回目は確認しない');
  assert.strictEqual(sheet.getLastColumn(), 8, '2回目は列を足さない');
});

test('setupRequestSheet：削除しないと答えたら、メモ・追加日は残してフィードバック・担当者だけを足す', () => {
  const {gas, g} = setup();
  const sheet = addCurrentRequestSheet(gas);
  gas.confirmAnswer = 'NO';
  g.setupRequestSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A1:J2'), [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者', 'メモ', '追加日'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '', '先方に確認', '2026/10/01']
  ]);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /「メモ」「追加日」の列は残しました/);
});

test('新FMT にリクエストが1件も無ければ、候補は「（取り込み元にリクエストがありません）」にして知らせる（すべて選択済みと区別する）', () => {
  const {gas, g} = setup([['食品スーパー', 'A社', 'アプリ刷新', '']]);
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(listOf(request.getRange('A2')), ['（取り込み元にリクエストがありません）']);
  assert.deepStrictEqual(listOf(request.getRange('D2')), WAIT_SUB);
  assert.match(gas.toasts[0].message, /「新FMT」にリクエストが見つかりません/);
});

test('removeUntouchedRequests：前の版で全件を取り込んだまま手を付けていない行を、確かめてから削除して候補に戻す', () => {
  const {gas, g} = setup();
  const header = ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス', '判断', 'メモ', '取り込み日'];
  const sheet = gas.addSheet('リクエスト', [header,
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', 'サービス化検討', '', '2026/10/01'],
    ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', '', '未判断', '', '2026/10/01'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい', '', '未判断', '', '2026/10/01']
  ].map((r, i) => (i && r[3][0] === '=' ? r.map((v, j) => (j === 3 ? "'" + v : v)) : r)), {rows: 10, columns: 8});
  g.setupRequestSheet();
  assert.deepStrictEqual(listOf(sheet.getRange('A5')), ['（すべて選択済み）'], '前の版の行はすべて選んだ行になっている');
  assert.match(gas.toasts[0].message, /判断・フィードバック・サービス案・担当者が空の行が 2件あります。.*removeUntouchedRequests\(\)/);

  gas.confirmAnswer = 'NO';
  g.removeUntouchedRequests();
  assert.strictEqual(sheet.getLastRow(), 4, '「いいえ」なら消さない');

  gas.confirmAnswer = 'YES';
  g.removeUntouchedRequests();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /判断が「未判断」か空で、フィードバック・サービス案・担当者が空の行が 2件あります/);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:E3'), [
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤'],
    ['', '', '', '', '']
  ], 'サービス案を付けた行は残す');
  assert.deepStrictEqual(listOf(sheet.getRange('A3')), ['食品スーパー', 'ドラッグストア'], '候補に戻る');
  assert.deepStrictEqual(listOf(sheet.getRange('C3')), WAIT_SUB);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /2行を削除しました。「新FMT」のリクエスト 3件のうち 1件を選んでいます/);
});

test('removeUntouchedRequests：データの行をすべて消すときも止まらない（空の行を1つ残す）', () => {
  const {gas, g} = setup();
  const header = ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス案', '判断', 'メモ', '追加日'];
  const sheet = gas.addSheet('リクエスト', [header,
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '未判断', '', ''],
    ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', '', '', '', '']
  ], {rows: 3, columns: 8});
  gas.confirmAnswer = 'YES';
  g.removeUntouchedRequests();
  assert.strictEqual(sheet.getLastRow(), 1);
  assert.strictEqual(sheet.getMaxRows(), 2);
  assert.deepStrictEqual(listOf(sheet.getRange('A2')), ['食品スーパー', 'ドラッグストア']);
});

test('読むシートは DIFF_RULES で requests: true を付けたシート：新FMT2 に付けると、新FMT2 のリクエストも候補にする', () => {
  const {gas, g} = setup();
  gas.get('DIFF_RULES')['新FMT2'].requests = true;
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'];
  const fmt2 = gas.addSheet('新FMT2', [['新FMT2'], header,
    ['1', '', 'ホームセンター', 'D社', '会員アプリ', 'ポイントを統合したい'],
    ['2', '', '食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']   // 新FMT と同じリクエストは1件とみなす
  ], {rows: 20, columns: 6});
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(listOf(request.getRange('A2')), ['食品スーパー', 'ドラッグストア', 'ホームセンター']);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /「新FMT」「新FMT2」のリクエスト 4件のうち 0件/);

  // 新FMT2 の編集でも候補を作り直す
  g.onEdit(gas.edit(fmt2, 'C5:F5', [['ホームセンター', 'E社', '店舗DX', '在庫を見える化したい']]));
  edit(gas, 'リクエスト', 'A2', 'ホームセンター');
  assert.deepStrictEqual(listOf(request.getRange('B2')), ['D社', 'E社']);
});

test('requests: true を付けたシートが DIFF_RULES に無ければ、止めて知らせる', () => {
  const {gas, g} = setup();
  delete gas.get('DIFF_RULES')['新FMT'].requests;
  assert.throws(() => g.setupRequestSheet(), /DIFF_RULES で、読むシートに requests: true を付けてください/);
});

test('前の版の注（「新FMT でリクエストが書き換えられました」）も、このスクリプトの注として付け替える', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  gas.asUser(() => request.getRange('D2').setNote('新FMT でリクエストが書き換えられました（2026/10/01）。サービス案・判断を見直してください。'));
  g.onEdit(gas.edit(source, 'F3', ''));
  assert.match(request.getRange('D2').getNote(), /^取り込み元に見つかりません/);
});

test('新FMT のリクエストの列の見出しが無ければ、止めて知らせる', () => {
  const {gas, g, source} = setup();
  gas.asUser(() => source.getRange('F2').setValue('要望'));
  assert.throws(() => g.setupRequestSheet(), /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
});

test('diagnoseRequestSources：シートごとに、リクエストの件数と候補に出ない理由を出す（読むだけ）', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['', 'B社', '店舗什器', '什器の在庫を見たい'],                     // サブインダストリーが空
    ['食品スーパー', 'C社', '', '配送を早めたい'],                       // 案件名が空
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']      // 同じリクエスト
  ]);
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  gas.writes.length = 0;
  gas.alerts.length = 0;
  g.diagnoseRequestSources();
  assert.deepStrictEqual(gas.writes, [], '何も書き換えない');
  const lines = gas.alerts[0].message.split('\n');
  assert.deepStrictEqual(lines, [
    '「新FMT」：リクエストのある行 4件 → 候補に出る 0件（出ないもの：選択済み 1件、サブインダストリー・得意先・案件名のどれかが空 2件' +
      '（サブインダストリーが空 1件、案件名が空 1件）、ほかの行・シートと同じリクエスト 1件）。',
    '「新FMT2」：読みません（test.gs の DIFF_RULES に requests: true がありません）。',
    '候補のシートとプルダウンは、今の内容で作られています。'
  ]);
});

test('diagnoseRequestSources：候補のシートが古い・プルダウンが前の版のままなら、setupRequestSheet() を案内する', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  // スクリプトを通さずに新FMT にリクエストを足す（onEdit が動かなかったときと同じ）
  gas.asUser(() => source.getRange('C8:F8').setValues([['食品スーパー', 'D社', '新店', '売場を見直したい']]));
  const request = gas.ss.getSheetByName('リクエスト');
  gas.asUser(() => request.getRange('A2').setDataValidation(g.SpreadsheetApp.newDataValidation().requireValueInList(['食品スーパー'], true).build()));
  g.diagnoseRequestSources();
  const message = gas.alerts[gas.alerts.length - 1].message;
  assert.match(message, /「新FMT」：リクエストのある行 5件 → 候補に出る 4件/);
  assert.match(message, /候補のシート（__REQUEST_LISTS）が古くなっています（今の案件 4件、候補のシート 3件）。setupRequestSheet\(\) を実行してください。/);
  assert.match(message, /「リクエスト」のプルダウンが前の版のままです。setupRequestSheet\(\) を実行してください。/);
});
