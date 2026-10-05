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
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス案', '判断', 'メモ', '追加日']);
  assert.strictEqual(request.getLastRow(), 1, '行は自動では足さない');

  assert.deepStrictEqual(listOf(request.getRange('A2')), ['食品スーパー', 'ドラッグストア']);
  ['B2', 'C2', 'D2', 'D1000'].forEach(a1 => assert.deepStrictEqual(listOf(request.getRange(a1)), WAIT_SUB, a1 + '：左の列が空なら選べない'));
  assert.match(request.getRange('C2').getDataValidation().getHelpText(), /先にサブインダストリーを選んでください/);
  assert.strictEqual(request.getRange('E2').getDataValidation(), null, 'サービス案は自由に入力する');
  assert.deepStrictEqual(listOf(request.getRange('F2')), ['未判断', 'サービス化検討', '棄却']);

  assert.strictEqual(request.getProtections().length, 0, 'A〜D はプルダウンで選ぶので保護しない');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /新FMT のリクエスト 3件のうち 0件を選んでいます（まだ選んでいないもの 3件）/);
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

test('A〜C を選ぶたびに、その行の右の列の候補が絞られる（編集した行のプルダウンを先に反映する）', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['食品スーパー', 'A社', 'アプリ刷新', 'クーポンを配りたい'],
    ['食品スーパー', 'B社', '店舗什器', '什器の在庫を見たい'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '在庫を店舗と共有したい']
  ]);
  g.setupRequestSheet();
  gas.writes.length = 0;
  let request = edit(gas, 'リクエスト', 'A3', '食品スーパー');
  const validations = gas.writes.filter(w => w.kind === 'setDataValidations').map(w => w.a1);
  assert.deepStrictEqual(validations.slice(0, 4), ['A3', 'B3', 'C3', 'D3'], '編集した行を先に');
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
  ['A5', 'B5', 'C5', 'D5', 'A20', 'D20'].forEach(a1 => assert.deepStrictEqual(listOf(request.getRange(a1)), ['（すべて選択済み）'], a1));

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
  edit(gas, 'リクエスト', 'E2', '会員分析基盤');
  edit(gas, 'リクエスト', 'F2', 'サービス化検討');
  assert.deepStrictEqual(listOf(gas.ss.getSheetByName('リクエスト').getRange('A2')), ['食品スーパー', 'ドラッグストア']);

  let request = edit(gas, 'リクエスト', 'A2', 'ドラッグストア');
  assert.deepStrictEqual(gas.dump(request, 'A2:F2')[0],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '', '会員分析基盤', 'サービス化検討'],
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

test('サービス案・判断・メモ・追加日を入力しても、ほかの列は変えない（判断も自動で入れない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  choose(gas, 2, ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  choose(gas, 3, ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  edit(gas, 'リクエスト', 'F3', '棄却');
  gas.writes.length = 0;
  const request = edit(gas, 'リクエスト', 'E2:E3', [['会員分析基盤'], ['会員分析基盤']]);
  edit(gas, 'リクエスト', 'G2:H2', [['先方に確認', '2026/10/05']]);
  assert.deepStrictEqual(gas.writes, [], '何も書き込まない');
  assert.deepStrictEqual(gas.dump(request, 'E2:H3'), [
    ['会員分析基盤', '', '先方に確認', '2026/10/05'],
    ['会員分析基盤', '棄却', '', '']
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
  edit(gas, 'リクエスト', 'A3', '食品スーパー');
  assert.deepStrictEqual(gas.dump(old, 'B3:C3')[0], ['A社', 'チラシのデジタル化'], '選んであった案件は候補から外す（残りの1つが入る）');
  assert.strictEqual(rangeOf(gas.ss.getSheetByName('タスク管理').getRange('A2')), 'リクエスト!E2:E20');
});

test('新FMT にリクエストが1件も無ければ、候補は「（新FMT にリクエストがありません）」にして知らせる（すべて選択済みと区別する）', () => {
  const {gas, g} = setup([['食品スーパー', 'A社', 'アプリ刷新', '']]);
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  ['A2', 'D2'].forEach(a1 => assert.deepStrictEqual(listOf(request.getRange(a1)), ['（新FMT にリクエストがありません）'], a1));
  assert.match(gas.toasts[0].message, /新FMT にリクエストが見つかりません/);
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
  assert.match(gas.toasts[0].message, /サービス案・判断・メモが空の行が 2件あります。.*removeUntouchedRequests\(\)/);

  gas.confirmAnswer = 'NO';
  g.removeUntouchedRequests();
  assert.strictEqual(sheet.getLastRow(), 4, '「いいえ」なら消さない');

  gas.confirmAnswer = 'YES';
  g.removeUntouchedRequests();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /判断が「未判断」の行が 2件あります/);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:E3'), [
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤'],
    ['', '', '', '', '']
  ], 'サービス案を付けた行は残す');
  assert.deepStrictEqual(listOf(sheet.getRange('A3')), ['食品スーパー', 'ドラッグストア'], '候補に戻る');
  assert.deepStrictEqual(listOf(sheet.getRange('C3')), WAIT_SUB);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /2行を削除しました。新FMT のリクエスト 3件のうち 1件を選んでいます/);
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

test('新FMT のリクエストの列の見出しが無ければ、止めて知らせる', () => {
  const {gas, g, source} = setup();
  gas.asUser(() => source.getRange('F2').setValue('要望'));
  assert.throws(() => g.setupRequestSheet(), /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
});
