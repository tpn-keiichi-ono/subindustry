'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/** 新FMT（2行目が見出し。C: サブインダストリー / D: 得意先 / E: 案件名 / F: サービスのリクエスト） */
const SOURCE = [
  ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
  ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい'],
  ['食品スーパー', 'B社', '店舗什器', ''],                                   // リクエストが無い → パネルに出さない
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

/** パネルに出ているリクエスト（A〜D をつないだ文字） */
const shown = data => plain(data.items.map(item => item.values.join(' / ')));

/** 選択パネルで、案件名とリクエストが合うものを「追加」する */
function add(g, project, request) {
  const item = g.getRequestPickerData().items.find(i => i.values[2] === project && i.values[3] === request);
  assert.ok(item, project + '／' + request + ' がパネルに出ていません');
  return g.addServiceRequest(item.key);
}

/* ---------------- シートの用意 ---------------- */

test('setupRequestSheet：リクエスト シートを作る。行は足さず、A〜D にプルダウンを付けない（選択パネルで登録する）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.deepStrictEqual(gas.dump(request, 'A1:H1')[0],
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者']);
  assert.strictEqual(request.getLastRow(), 1, '行は自動では足さない');
  assert.strictEqual(gas.alerts.length, 0, '使わなくなった列が無ければ確認しない');

  ['A2', 'B2', 'C2', 'D2', 'D1000'].forEach(a1 => assert.strictEqual(request.getRange(a1).getDataValidation(), null, a1 + '：プルダウンを付けない'));
  assert.deepStrictEqual(listOf(request.getRange('E2')), ['未判断', 'サービス化検討', '棄却']);
  ['F2', 'G2', 'H2'].forEach(a1 => assert.strictEqual(request.getRange(a1).getDataValidation(), null, a1 + '：フィードバック・サービス案・担当者は自由に入力する'));
  assert.strictEqual(gas.ss.getSheetByName('__REQUEST_LISTS'), null, '候補のシートは作らない');

  assert.strictEqual(request.getProtections().length, 0);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[0].message, /「新FMT」のリクエスト 3件のうち 0件を登録しています（まだ登録していないもの 3件）/);
  assert.doesNotMatch(gas.toasts[0].message, /前の版/, '前の版のプルダウンが無ければ知らせない');
});

test('setupRequestSheet：前の版の A〜D のプルダウンと候補のシート（__REQUEST_LISTS）を外す（判断のプルダウン・値は残す）', () => {
  const {gas, g} = setup();
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '佐藤']
  ], {rows: 20, columns: 8});
  const lists = gas.addSheet('__REQUEST_LISTS', [['（前の版の候補）']], {rows: 20, columns: 10});
  lists.hideSheet();
  const SpreadsheetApp = g.SpreadsheetApp;
  gas.asUser(() => {
    sheet.getRange('A2:D20').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(lists.getRange('A2:J2'), true).build());
  });

  g.setupRequestSheet();
  assert.strictEqual(gas.ss.getSheetByName('__REQUEST_LISTS'), null);
  ['A2', 'B2', 'C2', 'D2', 'D20'].forEach(a1 => assert.strictEqual(sheet.getRange(a1).getDataValidation(), null, a1));
  assert.deepStrictEqual(listOf(sheet.getRange('E2')), ['未判断', 'サービス化検討', '棄却']);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:H2')[0],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '佐藤'], '値はそのまま');
  const toast = gas.toasts[gas.toasts.length - 1].message;
  assert.match(toast, /前の版の A〜D のプルダウンと「__REQUEST_LISTS」シートを外しました。リクエストは、メニュー「サービス管理」→「リクエストを追加」から登録します。/);
  assert.match(toast, /「新FMT」のリクエスト 3件のうち 1件を登録しています（まだ登録していないもの 2件）/);

  g.setupRequestSheet();
  assert.doesNotMatch(gas.toasts[gas.toasts.length - 1].message, /前の版/, '2回目は外すものが無い');
});

/* ---------------- 選択パネル ---------------- */

test('メニュー「サービス管理」→「リクエストを追加」で、選択パネル（サイドバー）を開く', () => {
  const {gas, g} = setup();
  g.svcAddMenu_();
  assert.deepStrictEqual(plain(gas.menus), [{title: 'サービス管理', items: [{label: 'リクエストを追加', fn: 'openRequestPicker'}]}]);
  g.openRequestPicker();
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'RequestPickerView');
  assert.strictEqual(gas.sidebars[0].html.title, 'リクエストを追加');
  assert.deepStrictEqual(gas.writes, [], '開くだけでは書き込まない');

  // 画面から呼ぶサーバーの関数がある（末尾 _ の内部用ではない）
  const html = gas.sidebars[0].html.getContent();
  ['getRequestPickerData', 'addServiceRequest'].forEach(name => {
    assert.match(html, new RegExp('\\.' + name + '\\('), name + ' を呼んでいない');
    assert.strictEqual(typeof g[name], 'function', name);
  });
});

test('getRequestPickerData：まだ登録していないリクエストを、取り込み元の順に返す（リクエストの無い行・同じリクエストは1つ）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  gas.writes.length = 0;
  const data = g.getRequestPickerData();
  assert.deepStrictEqual(plain(data.items.map(item => [item.sheet].concat(item.values))), [
    ['新FMT', '食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['新FMT', '食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい'],
    ['新FMT', 'ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい']
  ]);
  assert.strictEqual(new Set(data.items.map(item => item.key)).size, 3, 'キーは1件ずつ違う');
  assert.deepStrictEqual(plain(Object.assign({}, data, {items: null, loadedAt: null})), {
    items: null, total: 3, registered: 0, remaining: 3, missing: 0, sources: ['新FMT'], requestSheet: 'リクエスト', loadedAt: null, synced: true
  });
  assert.match(data.loadedAt, /^\d{2}:\d{2}$/);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(data)), plain(data), 'google.script.run で返せる値だけ（Set・Date を含まない）');
  assert.deepStrictEqual(gas.writes, [], '登録した行が無ければ書き込まない');
});

test('addServiceRequest：選んだリクエストを最後の行の下に追加し、パネルの新しい内容を返す（「=…」は数式にしない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  const before = g.getRequestPickerData();

  let result = add(g, 'アプリ刷新', '会員の購買分析をしたい');
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.row, 2);
  assert.strictEqual(result.message, '2行目に追加しました。');
  assert.deepStrictEqual(gas.dump(request, 'A2:H2')[0],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '', '', ''], '判断などは入れない');
  assert.deepStrictEqual(shown(result.data), ['食品スーパー / A社 / チラシのデジタル化 / チラシの効果を測りたい',
    'ドラッグストア / C社 / EC立ち上げ / =在庫を店舗と共有したい'], '追加したものはパネルから消える');
  assert.strictEqual(result.data.registered, 1);
  assert.strictEqual(result.data.remaining, 2);

  result = add(g, 'EC立ち上げ', '=在庫を店舗と共有したい');
  assert.strictEqual(result.row, 3);
  assert.strictEqual(request.getRange('D3').getDisplayValue(), '=在庫を店舗と共有したい');
  assert.strictEqual(request.getRange('D3').getFormula(), '', '数式にしない');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');

  // 古い一覧から同じものを選んだとき（ほかの人がすでに追加した）
  result = g.addServiceRequest(before.items[0].key);
  assert.strictEqual(result.ok, false);
  assert.strictEqual(result.message, 'このリクエストは、すでに登録されています。');
  assert.strictEqual(request.getLastRow(), 3, '同じリクエストは2回書かない');
  assert.strictEqual(result.data.remaining, 1, '新しい一覧を返す');

  result = g.addServiceRequest('どこにもないキー');
  assert.strictEqual(result.ok, false);
  assert.match(result.message, /^取り込み元にこのリクエストが見つかりません/);
  assert.strictEqual(request.getLastRow(), 3);
});

test('addServiceRequest：判断だけを書いた行には書き足さない。行が足りなければ足して、判断のプルダウンも付ける', () => {
  const {gas, g} = setup();
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '', '', ''],
    ['', '', '', '', '棄却', '', '', '']
  ], {rows: 3, columns: 8});
  g.setupRequestSheet();
  const result = add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  assert.strictEqual(result.row, 4, '3行目（判断だけの行）の下に足す');
  assert.deepStrictEqual(gas.dump(sheet, 'A3:E4'), [
    ['', '', '', '', '棄却'],
    ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', '']
  ]);
  assert.strictEqual(sheet.getMaxRows(), 4);
  assert.deepStrictEqual(listOf(sheet.getRange('E4')), ['未判断', 'サービス化検討', '棄却'], '足した行にも判断のプルダウン');
});

test('サブインダストリー・得意先・案件名が空のリクエストも、パネルに出して登録できる', () => {
  const {gas, g} = setup([
    ['', 'B社', '店舗什器', '什器の在庫を見たい'],
    ['食品スーパー', 'C社', '', '配送を早めたい']
  ]);
  g.setupRequestSheet();
  assert.deepStrictEqual(shown(g.getRequestPickerData()), [' / B社 / 店舗什器 / 什器の在庫を見たい', '食品スーパー / C社 /  / 配送を早めたい']);
  const result = add(g, '店舗什器', '什器の在庫を見たい');
  assert.strictEqual(result.row, 2);
  assert.deepStrictEqual(gas.dump(gas.ss.getSheetByName('リクエスト'), 'A2:D2')[0], ['', 'B社', '店舗什器', '什器の在庫を見たい']);
  assert.deepStrictEqual(shown(result.data), ['食品スーパー / C社 /  / 配送を早めたい']);
  assert.strictEqual(add(g, '', '配送を早めたい').row, 3, 'A が空の行の下に足す');
});

test('ほかの処理がロックを持っているとき：パネルは開ける（取り込み元に合わせない）。追加は止めて知らせる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const item = g.getRequestPickerData().items[0];
  gas.lockBusy = true;
  const data = g.getRequestPickerData();
  assert.strictEqual(data.synced, false);
  assert.strictEqual(data.items.length, 3);
  assert.throws(() => g.addServiceRequest(item.key), /他の処理が実行中です/);
  assert.strictEqual(gas.ss.getSheetByName('リクエスト').getLastRow(), 1);
});

test('getRequestPickerData：リクエスト シートが無ければ、管理者に setupRequestSheet() を頼むよう知らせる', () => {
  const {g} = setup();
  assert.throws(() => g.getRequestPickerData(), /「リクエスト」シートがありません。管理者に setupRequestSheet\(\) の実行を頼んでください/);
  assert.throws(() => g.addServiceRequest('x'), /「リクエスト」シートがありません/);
});

/* ---------------- 取り込み元の変更 ---------------- */

test('パネルはいつも取り込み元の最新を読む。開いたときに、登録した行を取り込み元に合わせる（行は足さない）', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');

  // スクリプトを通さずに取り込み元を書き換える（編集で動く処理は無い）
  gas.asUser(() => source.getRange('C8:F8').setValues([['食品スーパー', 'D社', '新店', '売場を見直したい']]));
  assert.ok(shown(g.getRequestPickerData()).includes('食品スーパー / D社 / 新店 / 売場を見直したい'), '書き足したリクエストがすぐ出る');
  assert.strictEqual(request.getLastRow(), 2, '行は足さない');

  // サブインダストリーは同じ案件の値にそろえる
  gas.asUser(() => source.getRange('C3').setValue('大型スーパー'));
  g.getRequestPickerData();
  assert.strictEqual(gas.dump(request, 'A2')[0][0], '大型スーパー');

  // リクエストが書き換わったら、行は残して注を付け（新しいリクエストはパネルに出る）、戻したら外す
  gas.asUser(() => source.getRange('F3').setValue('会員の購買データを分析したい'));
  let data = g.getRequestPickerData();
  assert.match(request.getRange('D2').getNote(), /^取り込み元に見つかりません（\d{4}\/\d{2}\/\d{2} に気づきました）。「新FMT」で書き換えか削除された可能性があります。/);
  assert.strictEqual(data.missing, 1);
  assert.ok(shown(data).includes('大型スーパー / A社 / アプリ刷新 / 会員の購買データを分析したい'));
  gas.asUser(() => source.getRange('F3').setValue('会員の購買分析をしたい'));
  data = g.getRequestPickerData();
  assert.strictEqual(request.getRange('D2').getNote(), '');
  assert.strictEqual(data.missing, 0);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('取り込み元に見つからない行の注：前の版の注は付け替え、利用者が書いた注は変えない', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  gas.asUser(() => {
    request.getRange('D2').setNote('新FMT でリクエストが書き換えられました（2026/10/01）。サービス案・判断を見直してください。');
    request.getRange('D3').setNote('先方に確認中');
    source.getRange('F3:F4').setValues([[''], ['']]);
  });
  g.getRequestPickerData();
  assert.match(request.getRange('D2').getNote(), /^取り込み元に見つかりません/);
  assert.strictEqual(request.getRange('D3').getNote(), '先方に確認中');
});

/* ---------------- 前の版のシート・列 ---------------- */

test('前の版のシート：見出し（サービス・取り込み日）を書き換え、登録してあった行と値はそのまま使う。前の版の保護は外す', () => {
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
  assert.deepStrictEqual(shown(g.getRequestPickerData()), ['食品スーパー / A社 / チラシのデジタル化 / チラシの効果を測りたい',
    'ドラッグストア / C社 / EC立ち上げ / =在庫を店舗と共有したい'], '登録してあったリクエストはパネルに出さない');
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

test('取り込み元にリクエストが1件も無ければ、そう知らせる（すべて登録済みと区別する）', () => {
  const {gas, g} = setup([['食品スーパー', 'A社', 'アプリ刷新', '']]);
  g.setupRequestSheet();
  assert.match(gas.toasts[0].message, /「新FMT」にリクエストが見つかりません/);
  const data = g.getRequestPickerData();
  assert.strictEqual(data.total, 0);
  assert.strictEqual(data.items.length, 0);
});

test('removeUntouchedRequests：前の版で全件を取り込んだまま手を付けていない行を、確かめてから削除してパネルに戻す', () => {
  const {gas, g} = setup();
  const header = ['サブインダストリー', '得意先', '案件名', 'リクエスト', 'サービス', '判断', 'メモ', '取り込み日'];
  const sheet = gas.addSheet('リクエスト', [header,
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '会員分析基盤', 'サービス化検討', '', '2026/10/01'],
    ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', '', '未判断', '', '2026/10/01'],
    ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい', '', '未判断', '', '2026/10/01']
  ].map((r, i) => (i && r[3][0] === '=' ? r.map((v, j) => (j === 3 ? "'" + v : v)) : r)), {rows: 10, columns: 8});
  g.setupRequestSheet();
  assert.strictEqual(g.getRequestPickerData().items.length, 0, '前の版の行はすべて登録した行になっている');
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
  assert.deepStrictEqual(shown(g.getRequestPickerData()), ['食品スーパー / A社 / チラシのデジタル化 / チラシの効果を測りたい',
    'ドラッグストア / C社 / EC立ち上げ / =在庫を店舗と共有したい'], 'パネルに戻る');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /2行を削除しました。「新FMT」のリクエスト 3件のうち 1件を登録しています/);
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
});

/* ---------------- 読むシート ---------------- */

test('読むシートは DIFF_RULES で requests: true を付けたシート：新FMT2 に付けると、新FMT2 のリクエストもパネルに出す', () => {
  const {gas, g} = setup();
  gas.get('DIFF_RULES')['新FMT2'].requests = true;
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'];
  const fmt2 = gas.addSheet('新FMT2', [['新FMT2'], header,
    ['1', '', 'ホームセンター', 'D社', '会員アプリ', 'ポイントを統合したい'],
    ['2', '', '食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']   // 新FMT と同じリクエストは1件とみなす
  ], {rows: 20, columns: 6});
  g.setupRequestSheet();
  assert.match(gas.toasts[gas.toasts.length - 1].message, /「新FMT」「新FMT2」のリクエスト 4件のうち 0件/);
  let data = g.getRequestPickerData();
  assert.deepStrictEqual(plain(data.sources), ['新FMT', '新FMT2']);
  assert.deepStrictEqual(plain(data.items.map(item => item.sheet)), ['新FMT', '新FMT', '新FMT', '新FMT2']);

  gas.asUser(() => fmt2.getRange('C5:F5').setValues([['ホームセンター', 'E社', '店舗DX', '在庫を見える化したい']]));
  data = g.getRequestPickerData();
  assert.ok(shown(data).includes('ホームセンター / E社 / 店舗DX / 在庫を見える化したい'));
});

test('requests: true を付けたシートが DIFF_RULES に無ければ、止めて知らせる', () => {
  const {gas, g} = setup();
  delete gas.get('DIFF_RULES')['新FMT'].requests;
  assert.throws(() => g.setupRequestSheet(), /DIFF_RULES で、読むシートに requests: true を付けてください/);
});

test('新FMT のリクエストの列の見出しが無ければ、止めて知らせる', () => {
  const {gas, g, source} = setup();
  gas.asUser(() => source.getRange('F2').setValue('要望'));
  assert.throws(() => g.setupRequestSheet(), /「サービスのリクエスト」の見出しがちょうど1つ必要です/);
});

/* ---------------- 確認 ---------------- */

test('diagnoseRequestSources：シートごとに、リクエストの件数とパネルに出ない理由を出す（読むだけ）', () => {
  const {gas, g} = setup([
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['', 'B社', '店舗什器', '什器の在庫を見たい'],                     // サブインダストリーが空でも出す
    ['食品スーパー', 'C社', '', '配送を早めたい'],                       // 案件名が空でも出す
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']      // 同じリクエスト
  ]);
  g.setupRequestSheet();
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  gas.writes.length = 0;
  gas.alerts.length = 0;
  g.diagnoseRequestSources();
  assert.deepStrictEqual(gas.writes, [], '何も書き換えない');
  assert.deepStrictEqual(gas.alerts[0].message.split('\n'), [
    '「新FMT」：リクエストのある行 4件 → パネルに出る 2件（出ないもの：登録済み 1件、ほかの行・シートと同じリクエスト 1件）。',
    '「新FMT2」：読みません（test.gs の DIFF_RULES に requests: true がありません）。'
  ]);
});

test('diagnoseRequestSources：前の版の候補のシートが残っていれば、setupRequestSheet() を案内する', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  gas.addSheet('__REQUEST_LISTS', [['']], {rows: 2, columns: 2});
  g.diagnoseRequestSources();
  const message = gas.alerts[gas.alerts.length - 1].message;
  assert.match(message, /「新FMT」：リクエストのある行 4件 → パネルに出る 3件（出ないもの：ほかの行・シートと同じリクエスト 1件）。/);
  assert.match(message, /前の版の「__REQUEST_LISTS」シートが残っています。setupRequestSheet\(\) を実行すると外します。/);
});
