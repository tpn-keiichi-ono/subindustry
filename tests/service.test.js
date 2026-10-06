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

/** パネルに出ているリクエスト（サブインダストリー・得意先・案件名・リクエストをつないだ文字） */
const shown = data => plain(data.items.map(item => item.values.slice(0, 4).join(' / ')));

/** リクエスト シートの見出し（setupRequestSheet() で作ったときの並び） */
const REQUEST_HEADERS = ['ID', 'サブインダストリー', '得意先', '案件名', 'アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当',
  'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '主管本部', '担当者'];
/** 取り込み元から書く列のうち、リクエストを見分ける4つ */
const ABCD = ['サブインダストリー', '得意先', '案件名', 'リクエスト'];
const SALES = ['アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当'];

/** 1行目の見出しが label の列の、row 行目のセル（列の並びに左右されないように見出しで探す） */
function cellOf(sheet, label, row) {
  const headers = Array.from(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
  const col = headers.indexOf(label) + 1;
  assert.ok(col > 0, '「' + label + '」の列がありません');
  return sheet.getRange(row, col);
}
/** row 行目の、labels の見出しの列の値（表示） */
const valuesOf = (sheet, row, labels) => labels.map(label => cellOf(sheet, label, row).getDisplayValue());
/** 1行目の見出し */
const headersOf = sheet => Array.from(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);

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
  assert.deepStrictEqual(headersOf(request), REQUEST_HEADERS, 'ID は左端、リクエストした営業の4列は案件名とリクエストの間、主管本部は担当者の左');
  assert.strictEqual(request.getLastRow(), 1, '行は自動では足さない');
  assert.strictEqual(gas.alerts.length, 0, '使わなくなった列が無ければ確認しない');

  ABCD.concat(SALES).forEach(label => {
    assert.strictEqual(cellOf(request, label, 2).getDataValidation(), null, label + '：プルダウンを付けない');
    assert.strictEqual(cellOf(request, label, 1000).getDataValidation(), null, label);
  });
  assert.deepStrictEqual(listOf(cellOf(request, '判断', 2)), ['未判断', 'サービス化検討', '棄却']);
  ['サービス部門からのフィードバック', 'サービス案', '主管本部', '担当者'].forEach(label =>
    assert.strictEqual(cellOf(request, label, 2).getDataValidation(), null, label + '：自由に入力する'));
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
  ABCD.concat(SALES).forEach(label => {
    assert.strictEqual(cellOf(sheet, label, 2).getDataValidation(), null, label);
    assert.strictEqual(cellOf(sheet, label, 20).getDataValidation(), null, label);
  });
  assert.deepStrictEqual(listOf(cellOf(sheet, '判断', 2)), ['未判断', 'サービス化検討', '棄却']);
  assert.deepStrictEqual(valuesOf(sheet, 2, ABCD.concat(['判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'])),
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '佐藤'], '値はそのまま');
  const toast = gas.toasts[gas.toasts.length - 1].message;
  assert.match(toast, /前の版の A〜D のプルダウンと「__REQUEST_LISTS」シートを外しました。リクエストは、メニュー「🟪RXサービスMTG用」→「リクエストを追加」から登録します。/);
  assert.match(toast, /「新FMT」のリクエスト 3件のうち 1件を登録しています（まだ登録していないもの 2件）/);

  g.setupRequestSheet();
  assert.doesNotMatch(gas.toasts[gas.toasts.length - 1].message, /前の版/, '2回目は外すものが無い');
});

/* ---------------- 選択パネル ---------------- */

/** 前の版のシート（A〜D に連動プルダウン。候補に無い値は拒否。setupRequestSheet() をまだ実行していない） */
function addLegacyDropdownSheet(gas) {
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者']
  ], {rows: 20, columns: 8});
  const lists = gas.addSheet('__REQUEST_LISTS', [['（前の版の候補）']], {rows: 20, columns: 10});
  lists.hideSheet();
  const SpreadsheetApp = gas.global.SpreadsheetApp;
  gas.asUser(() => {
    sheet.getRange('A2:A20').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(lists.getRange('G2:J2'), true)
      .setAllowInvalid(false).setHelpText('サブインダストリーは一覧から選んでください（リクエストがあり、まだ選んでいない案件だけが出ます）。').build());
    sheet.getRange('B2:D20').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['（先にサブインダストリーを選んでください）'], true)
      .setAllowInvalid(false).build());
    sheet.getRange('E2:E20').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['未判断', 'サービス化検討', '棄却'], true)
      .setAllowInvalid(false).build());
  });
  return sheet;
}

test('前の版の A〜D のプルダウンが残っていても追加できる：パネルを開いたときに外す（setupRequestSheet() を待たない）', () => {
  const {gas, g} = setup();
  const sheet = addLegacyDropdownSheet(gas);
  // 本物と同じく、プルダウンが残っているとスクリプトの書き込みも止まる
  assert.throws(() => sheet.getRange('A2').setValue('食品スーパー'), /サブインダストリーは一覧から選んでください/);

  const result = add(g, 'アプリ刷新', '会員の購買分析をしたい');
  assert.strictEqual(result.ok, true, result.message);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:D2')[0], ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  ['A2', 'B2', 'C2', 'D20'].forEach(a1 => assert.strictEqual(sheet.getRange(a1).getDataValidation(), null, a1 + '：A〜D のプルダウンは外す'));
  assert.deepStrictEqual(listOf(sheet.getRange('E2')), ['未判断', 'サービス化検討', '棄却'], '判断のプルダウンは残す');
  assert.strictEqual(gas.ss.getSheetByName('__REQUEST_LISTS'), null, '前の版の候補のシートも外す');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('パネルを開いたときにロックが取れず前の版のプルダウンを外せなくても、「追加」のときに外して書く', () => {
  const {gas, g} = setup();
  const sheet = addLegacyDropdownSheet(gas);
  gas.lockBusy = true;
  const data = g.getRequestPickerData();
  assert.strictEqual(data.synced, false);
  assert.ok(sheet.getRange('A2').getDataValidation(), 'ロックが取れなければ何も変えない');
  gas.lockBusy = false;
  const result = g.addServiceRequest(data.items[1].key);
  assert.strictEqual(result.ok, true, result.message);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:D2')[0], ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  assert.strictEqual(sheet.getRange('A3').getDataValidation(), null);
});


test('メニュー「🟪RXサービスMTG用」→「リクエストを追加」で、選択パネル（サイドバー）を開く（リクエスト シートのときだけ）', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  gas.writes.length = 0;
  g.svcAddMenu_();
  assert.deepStrictEqual(plain(gas.menus), [{title: '🟪RXサービスMTG用', items: [{label: 'リクエストを追加', fn: 'openRequestPicker'}]}]);

  // リクエスト シート以外では開かず、開けるシートを知らせる
  gas.select(source, 'D3');
  assert.strictEqual(g.openRequestPicker(), false);
  assert.strictEqual(gas.sidebars.length, 0);
  assert.strictEqual(gas.alerts[0].title, '🟪RXサービスMTG用');
  assert.strictEqual(gas.alerts[0].message, 'このシートでは開けません。「リクエスト」のシートで開いてください。');

  gas.select(gas.ss.getSheetByName('リクエスト'), 'A2');
  assert.strictEqual(g.openRequestPicker(), true);
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
  assert.deepStrictEqual(plain(data.items.map(item => [item.sheet].concat(item.values.slice(0, 4)))), [
    ['新FMT', '食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['新FMT', '食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい'],
    ['新FMT', 'ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい']
  ]);
  assert.strictEqual(new Set(data.items.map(item => item.key)).size, 3, 'キーは1件ずつ違う');
  assert.deepStrictEqual(plain(Object.assign({}, data, {items: null, loadedAt: null})), {
    items: null, total: 3, registered: 0, remaining: 3, sources: ['新FMT'], requestSheet: 'リクエスト', loadedAt: null, synced: true
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
  assert.strictEqual(result.message, 'SR-0001 として2行目に追加しました。');
  assert.strictEqual(result.id, 'SR-0001');
  assert.deepStrictEqual(valuesOf(request, 2, REQUEST_HEADERS),
    ['SR-0001', '食品スーパー', 'A社', 'アプリ刷新', '', '', '', '', '会員の購買分析をしたい', '', '', '', '', ''], '判断などは入れない');
  assert.deepStrictEqual(shown(result.data), ['食品スーパー / A社 / チラシのデジタル化 / チラシの効果を測りたい',
    'ドラッグストア / C社 / EC立ち上げ / =在庫を店舗と共有したい'], '追加したものはパネルから消える');
  assert.strictEqual(result.data.registered, 1);
  assert.strictEqual(result.data.remaining, 2);

  result = add(g, 'EC立ち上げ', '=在庫を店舗と共有したい');
  assert.strictEqual(result.row, 3);
  assert.strictEqual(cellOf(request, 'リクエスト', 3).getDisplayValue(), '=在庫を店舗と共有したい');
  assert.strictEqual(cellOf(request, 'リクエスト', 3).getFormula(), '', '数式にしない');
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
  assert.deepStrictEqual([3, 4].map(row => valuesOf(sheet, row, ABCD.concat(['判断']))), [
    ['', '', '', '', '棄却'],
    ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', '']
  ]);
  assert.strictEqual(sheet.getMaxRows(), 4);
  assert.deepStrictEqual(listOf(cellOf(sheet, '判断', 4)), ['未判断', 'サービス化検討', '棄却'], '足した行にも判断のプルダウン');
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
  assert.deepStrictEqual(valuesOf(gas.ss.getSheetByName('リクエスト'), 2, ABCD), ['', 'B社', '店舗什器', '什器の在庫を見たい']);
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

test('差分追跡などがドキュメントロックを使っていても、パネルの追加は待たされない（サービス管理は別のロック）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  gas.lockBusy = 'document';   // 取り込み元の編集で、差分追跡がドキュメントロックを持っている
  const data = g.getRequestPickerData();
  assert.strictEqual(data.synced, true, '登録した行を取り込み元に合わせられる');
  const result = g.addServiceRequest(data.items[0].key);
  assert.strictEqual(result.ok, true, result.message);
  assert.strictEqual(result.id, 'SR-0001');

  // サービス管理どうしは同じロックで1つずつ動く（ほかの人が追加している間は待つ）
  gas.lockBusy = 'script';
  assert.throws(() => g.addServiceRequest(data.items[1].key), /他の処理が実行中です/);
  assert.strictEqual(g.getRequestPickerData().synced, false);
  gas.lockBusy = false;
});

test('ロックを待ちきれなかったときは、ロックを持っている処理（何を・誰が・いつから）を知らせる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const key = g.getRequestPickerData().items[0].key;
  const cache = g.CacheService.getScriptCache();
  assert.strictEqual(cache.get('TASK_LOCK_HOLDER'), null, '終わった処理の記録は残さない');
  cache.put('TASK_LOCK_HOLDER', JSON.stringify({label: 'リクエストの追加', user: 'sato@example.com', at: Date.now() - 12000}), 600);
  gas.lockBusy = 'script';
  assert.throws(() => g.addServiceRequest(key),
    /^Error: 他の処理（リクエストの追加・sato@example\.com・1[23]秒前から）が実行中です。少し待ってからもう一度お試しください。$/);
  gas.lockBusy = false;
});

test('diagnoseServiceLock：サービス管理のロックの種類と、今使われているロックを出す（読むだけ）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  gas.writes.length = 0;
  g.diagnoseServiceLock();
  assert.deepStrictEqual(gas.alerts[0].message.split('\n'), [
    'サービス管理のロック：差分追跡とは別のロック（新しい版）を使っています。',
    'サービス管理のロック：空いています。',
    '差分追跡などのロック：空いています。'
  ]);
  assert.strictEqual(gas.alerts[0].title, '🟪RXサービスMTG用');
  assert.deepStrictEqual(gas.writes, [], '何も書き換えない');

  gas.lockBusy = 'document';
  g.diagnoseServiceLock();
  assert.match(gas.alerts[1].message, /差分追跡などのロック：今、使われています/);
  assert.match(gas.alerts[1].message, /サービス管理のロック：空いています/);
  gas.lockBusy = false;
});

test('getRequestPickerData：リクエスト シートが無ければ、管理者に setupRequestSheet() を頼むよう知らせる', () => {
  const {g} = setup();
  assert.throws(() => g.getRequestPickerData(), /「リクエスト」シートがありません。管理者に setupRequestSheet\(\) の実行を頼んでください/);
  assert.throws(() => g.addServiceRequest('x'), /「リクエスト」シートがありません/);
});

/* ---------------- 一覧のモーダル（パネルの件数のバッジ） ---------------- */

test('openRequestListDialog：バッジから一覧のモーダルを開く（押したバッジのタブで開く。見出しは画面の中）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.openRequestListDialog('registered');
  g.openRequestListDialog('なにか');
  assert.strictEqual(gas.dialogs.length, 2);
  assert.strictEqual(gas.dialogs[0].html.file, 'RequestListDialog');
  assert.strictEqual(gas.dialogs[0].title, ' ');
  assert.strictEqual(gas.dialogs[0].html.data.kind, 'registered');
  assert.strictEqual(gas.dialogs[1].html.data.kind, 'unregistered', '知らない値なら未登録のタブ');
  assert.ok(gas.dialogs[0].html.width >= 600 && gas.dialogs[0].html.height >= 400);
  const html = gas.dialogs[0].html.getContent();
  ['getRequestListData', 'saveCredentialDialogSize', 'addServiceRequestFromList', 'deleteServiceRequest'].forEach(name => {
    assert.match(html, new RegExp('\\.' + name + '\\('), name + ' を呼んでいない');
    assert.strictEqual(typeof g[name], 'function', name);
  });
  assert.match(gas.global.HtmlService.createTemplateFromFile('RequestPickerView').evaluate().getContent(), /\.openRequestListDialog\(/,
    'パネルのバッジから開く');
});

test('getRequestListData：未登録と登録済みの一覧を返す（登録済みは行番号・判断など。取り込み元に無いものに印。読むだけ）', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  gas.asUser(() => {
    [['判断', 'サービス化検討'], ['サービス部門からのフィードバック', '前向きに検討'], ['サービス案', '会員分析基盤'],
      ['主管本部', 'サービス本部'], ['担当者', '佐藤']].forEach(([label, value]) => cellOf(request, label, 2).setValue(value));
    source.getRange('F4').setValue('チラシの配布数を減らしたい');   // 登録済みの3行目は取り込み元に無くなる
  });
  gas.writes.length = 0;
  const data = g.getRequestListData();
  assert.deepStrictEqual(gas.writes, [], '何も書き込まない');
  assert.ok(data.unregistered.concat(data.registered).every(item => item.key), '追加・削除に使うキーを付ける');
  const withoutKey = list => plain(list).map(item => { delete item.key; item.values = item.values.slice(0, 4); return item; });
  assert.deepStrictEqual(withoutKey(data.unregistered), [
    {sheet: '新FMT', values: ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの配布数を減らしたい']},
    {sheet: '新FMT', values: ['ドラッグストア', 'C社', 'EC立ち上げ', '=在庫を店舗と共有したい']}
  ]);
  assert.deepStrictEqual(withoutKey(data.registered), [
    {row: 2, id: 'SR-0001', sheet: '新FMT', missing: false, values: ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
      decision: 'サービス化検討', feedback: '前向きに検討', service: '会員分析基盤', department: 'サービス本部', owner: '佐藤'},
    {row: 3, id: 'SR-0002', sheet: '', missing: true, values: ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい'],
      decision: '', feedback: '', service: '', department: '', owner: ''}
  ]);
  assert.strictEqual(data.requestSheet, 'リクエスト');
  assert.match(data.loadedAt, /^\d{2}:\d{2}$/);
});

test('addServiceRequestFromList：一覧のモーダルから追加し、一覧の新しい内容を返す', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  const item = g.getRequestListData().unregistered.find(i => i.values[2] === 'チラシのデジタル化');
  const result = g.addServiceRequestFromList(item.key);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.row, 2);
  assert.strictEqual(result.message, 'SR-0001 として2行目に追加しました。');
  assert.deepStrictEqual(valuesOf(request, 2, ABCD), ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい']);
  assert.strictEqual(result.list.unregistered.length, 2);
  assert.deepStrictEqual(plain(result.list.registered.map(r => [r.row, r.values[2]])), [[2, 'チラシのデジタル化']]);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');

  const again = g.addServiceRequestFromList(item.key);
  assert.strictEqual(again.ok, false);
  assert.strictEqual(again.message, 'このリクエストは、すでに登録されています。');
  assert.strictEqual(request.getLastRow(), 2);
});

test('deleteServiceRequest：登録済みの行を削除し、そのリクエストを未登録に戻す（判断などの入力も行ごと消える）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  gas.asUser(() => [['判断', 'サービス化検討'], ['サービス案', '会員分析基盤'], ['担当者', '佐藤']]
    .forEach(([label, value]) => cellOf(request, label, 2).setValue(value)));
  const target = g.getRequestListData().registered.find(r => r.row === 2);

  gas.writes.length = 0;
  const result = g.deleteServiceRequest(target.row, target.key);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.message, 'SR-0001（2行目）を削除しました。');
  assert.deepStrictEqual(valuesOf(request, 2, ABCD.concat(['判断'])), ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', ''], '下の行が上がる');
  assert.deepStrictEqual(plain(result.list.registered.map(r => r.row)), [2]);
  assert.ok(result.list.unregistered.some(i => i.values[2] === 'アプリ刷新'), '未登録に戻る');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');

  // 同じ行をもう一度削除しようとしても（古い一覧から）、ほかの行は消さない
  const stale = g.deleteServiceRequest(target.row, target.key);
  assert.strictEqual(stale.ok, false);
  assert.match(stale.message, /^この行が見つかりません/);
  assert.strictEqual(request.getLastRow(), 2);
});

test('deleteServiceRequest：並べ替えなどで行がずれていても、同じリクエストの行を探して削除する（1つに決まらなければ削除しない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  const target = g.getRequestListData().registered.find(r => r.values[2] === 'チラシのデジタル化');   // 3行目
  gas.asUser(() => request.insertRowsAfter(1, 1));   // 一覧を開いたあとに、ほかの人が上に行を足した → 4行目にずれる

  const result = g.deleteServiceRequest(target.row, target.key);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.row, 4);
  assert.deepStrictEqual([2, 3, 4].map(row => valuesOf(request, row, ['案件名'])), [[''], ['アプリ刷新'], ['']]);

  // 同じリクエストの行が2つあり、どちらも指定の行でなければ削除しない
  gas.asUser(() => [5, 6].forEach(row => ABCD.forEach((label, i) =>
    cellOf(request, label, row).setValue(['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'][i]))));
  const dup = g.getRequestListData().registered.find(r => r.row === 3);
  const before = request.getLastRow();
  const ambiguous = g.deleteServiceRequest(99, dup.key);
  assert.strictEqual(ambiguous.ok, false);
  assert.match(ambiguous.message, /^同じリクエストの行が複数ある/);
  assert.strictEqual(request.getLastRow(), before);
  assert.strictEqual(g.deleteServiceRequest(5, dup.key).row, 5, '行が合えば、その行を削除する');
});

test('deleteServiceRequest：データの行をすべて消すときも止まらない（空の行を1つ残す）', () => {
  const {gas, g} = setup();
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '', '', '']
  ], {rows: 2, columns: 8});
  const target = g.getRequestListData().registered[0];
  assert.strictEqual(g.deleteServiceRequest(target.row, target.key).ok, true);
  assert.strictEqual(sheet.getLastRow(), 1);
  assert.strictEqual(sheet.getMaxRows(), 2);
});

test('getRequestListData：ほかの処理がロックを持っていても読める', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  gas.lockBusy = true;
  assert.strictEqual(g.getRequestListData().unregistered.length, 3);
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
  assert.strictEqual(cellOf(request, 'サブインダストリー', 2).getDisplayValue(), '大型スーパー');

  // リクエストが書き換わっても、行はそのまま残し、メモは付けない（新しいリクエストはパネルに出る）
  gas.asUser(() => source.getRange('F3').setValue('会員の購買データを分析したい'));
  const data = g.getRequestPickerData();
  assert.strictEqual(cellOf(request, 'リクエスト', 2).getNote(), '', 'メモは付けない');
  assert.deepStrictEqual(valuesOf(request, 2, ABCD), ['大型スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい']);
  assert.ok(shown(data).includes('大型スーパー / A社 / アプリ刷新 / 会員の購買データを分析したい'));
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('前の版で「リクエスト」のセルに付けたメモは外し、利用者が書いたメモは変えない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  add(g, 'EC立ち上げ', '=在庫を店舗と共有したい');
  gas.asUser(() => {
    cellOf(request, 'リクエスト', 2).setNote('取り込み元に見つかりません（2026/10/05 に気づきました）。「新FMT」で書き換えか削除された可能性があります。');
    cellOf(request, 'リクエスト', 3).setNote('新FMT でリクエストが書き換えられました（2026/10/01）。サービス案・判断を見直してください。');
    cellOf(request, 'リクエスト', 4).setNote('先方に確認中');
  });
  g.setupRequestSheet();
  assert.deepStrictEqual([2, 3, 4].map(row => cellOf(request, 'リクエスト', row).getNote()), ['', '', '先方に確認中']);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /前の版で「リクエスト」のセルに付けたメモを 2件外しました。/);

  // 前の版の後片付けが済む前（前の版から貼り替えた直後）なら、パネルを開いたときにも外す
  gas.asUser(() => cellOf(request, 'リクエスト', 2).setNote('取り込み元でリクエストが書き換えられました（2026/10/01）。'));
  g.PropertiesService.getDocumentProperties().deleteProperty('SVC_LEGACY_CLEANED');
  g.getRequestPickerData();
  assert.strictEqual(cellOf(request, 'リクエスト', 2).getNote(), '');

  // 後片付けが済んだら、パネルを開いてもメモ・入力規則は読まない（読むのに時間がかかるため）
  gas.reads.length = 0;
  g.getRequestPickerData();
  assert.deepStrictEqual(gas.reads.filter(r => r.sheet === 'リクエスト' && /Notes|Validations/.test(r.kind)), []);
});

/* ---------------- リクエストした営業・主管本部 ---------------- */

/** 新FMT に、リクエストした営業の4列（アカウント責任者部署・アカウント責任者・BX部署・BX担当）もある */
function setupWithSales() {
  const gas = createGas();
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当',
    'サービスのリクエスト'];
  gas.addSheet('新FMT', [['新FMT'], header,
    ['1', '', '食品スーパー', 'A社', 'アプリ刷新', '流通第一本部', '佐藤', 'BX推進部', '鈴木', '会員の購買分析をしたい'],
    ['2', '', '食品スーパー', 'B社', '店舗什器', '流通第二本部', '田中', '', '', '什器の在庫を見たい']
  ], {rows: 20, columns: 10});
  return {gas, g: gas.global, source: gas.ss.getSheetByName('新FMT')};
}

test('リクエストした営業の4列も取り込み元から書き、パネルを開いたときに取り込み元に合わせる（列ごとにまとめて書く）', () => {
  const {gas, g, source} = setupWithSales();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, '店舗什器', '什器の在庫を見たい');
  assert.deepStrictEqual(valuesOf(request, 2, SALES), ['流通第一本部', '佐藤', 'BX推進部', '鈴木']);
  assert.deepStrictEqual(valuesOf(request, 3, SALES), ['流通第二本部', '田中', '', '']);

  gas.asUser(() => { source.getRange('G3').setValue('高橋'); source.getRange('I4').setValue('山本'); });
  gas.writes.length = 0;
  g.getRequestPickerData();
  assert.deepStrictEqual(valuesOf(request, 2, SALES), ['流通第一本部', '高橋', 'BX推進部', '鈴木']);
  assert.deepStrictEqual(valuesOf(request, 3, SALES), ['流通第二本部', '田中', '', '山本']);
  assert.strictEqual(gas.writes.length, 2, '変わった列（アカウント責任者・BX担当）だけを、列ごとに1回で書く');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');

  gas.writes.length = 0;
  g.getRequestPickerData();
  assert.deepStrictEqual(gas.writes, [], '変わっていなければ書かない');
  const item = g.getRequestListData().registered[0];
  assert.deepStrictEqual(plain(item.values.slice(4)), ['流通第一本部', '高橋', 'BX推進部', '鈴木'], '一覧のモーダルにも渡す');
});

test('setupRequestSheet：前からある登録済みの行に、営業の4列を取り込み元から入れる', () => {
  const {gas, g} = setupWithSales();
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '佐藤']
  ], {rows: 10, columns: 8});
  g.setupRequestSheet();
  assert.deepStrictEqual(valuesOf(sheet, 2, SALES.concat(['判断', 'サービス案', '主管本部', '担当者'])),
    ['流通第一本部', '佐藤', 'BX推進部', '鈴木', 'サービス化検討', '会員分析基盤', '', '佐藤'], '担当者の値はそのまま、主管本部は空で足す');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /登録済みの行を取り込み元に合わせました（4セル）。/);
});

test('setupRequestSheet の前（営業の列がまだ無いシート）でも、パネルから追加できる（ある列だけ書く）', () => {
  const {gas, g} = setupWithSales();
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者']
  ], {rows: 10, columns: 8});
  const result = add(g, 'アプリ刷新', '会員の購買分析をしたい');
  assert.strictEqual(result.ok, true, result.message);
  assert.deepStrictEqual(gas.dump(sheet, 'A2:H2')[0], ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '', '', '']);
  assert.strictEqual(sheet.getLastColumn(), 8, '列は足さない（setupRequestSheet() で足す）');
  assert.strictEqual(g.getRequestPickerData().registered, 1);
});

test('removeUntouchedRequests：主管本部だけが入っている行も「手を付けた行」として残す', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  gas.asUser(() => cellOf(request, '主管本部', 2).setValue('サービス本部'));
  gas.confirmAnswer = 'YES';
  g.removeUntouchedRequests();
  assert.deepStrictEqual(valuesOf(request, 2, ['案件名', '主管本部']), ['アプリ刷新', 'サービス本部']);
  assert.strictEqual(request.getLastRow(), 2, '手を付けていない3行目だけを削除する');
});

/* ---------------- ID ---------------- */

test('ID：追加するたびに SR-0001 から順に振り、削除した番号は使い回さない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  assert.strictEqual(add(g, 'アプリ刷新', '会員の購買分析をしたい').id, 'SR-0001');
  assert.strictEqual(add(g, 'チラシのデジタル化', 'チラシの効果を測りたい').id, 'SR-0002');
  const last = g.getRequestListData().registered.find(r => r.id === 'SR-0002');
  assert.strictEqual(g.deleteServiceRequest(last.row, last.key, last.id).message, 'SR-0002（3行目）を削除しました。');
  assert.strictEqual(add(g, 'EC立ち上げ', '=在庫を店舗と共有したい').id, 'SR-0003', '最後の番号を削除しても、次は続きの番号');
  assert.deepStrictEqual([2, 3].map(row => cellOf(request, 'ID', row).getDisplayValue()), ['SR-0001', 'SR-0003']);
});

test('ID：振る前にシートの ID を読み、手で書いた番号（SR-0050・12 など）より大きい番号を振る', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  gas.asUser(() => cellOf(request, 'ID', 2).setValue('ＳＲ－００５０'));   // 全角で手入力
  assert.strictEqual(add(g, 'チラシのデジタル化', 'チラシの効果を測りたい').id, 'SR-0051');
  gas.asUser(() => cellOf(request, 'ID', 3).setValue('120'));             // 番号だけ
  assert.strictEqual(add(g, 'EC立ち上げ', '=在庫を店舗と共有したい').id, 'SR-0121');
});

test('ID：行のコピーなどで重複した ID・読めない ID は、パネルを開いたときに下の行へ新しい番号を振り直して知らせる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  add(g, 'EC立ち上げ', '=在庫を店舗と共有したい');
  gas.asUser(() => {
    cellOf(request, 'ID', 3).setValue('SR-0001');   // 2行目と同じ（行をコピーした）
    cellOf(request, 'ID', 4).setValue('あとで振る');
  });
  gas.toasts.length = 0;
  g.getRequestPickerData();
  assert.deepStrictEqual([2, 3, 4].map(row => cellOf(request, 'ID', row).getDisplayValue()), ['SR-0001', 'SR-0004', 'SR-0005'],
    '上の行の ID は残し、下の行に続きの番号を振る');
  assert.match(gas.toasts[0].message, /ほかの行と同じ ID（または読めない ID）だった 2行に、新しい ID を振りました（3行目：SR-0001 → SR-0004、4行目：あとで振る → SR-0005）。/);
  assert.strictEqual(gas.toasts[0].title, '🟪RXサービスMTG用');

  gas.toasts.length = 0;
  gas.writes.length = 0;
  g.getRequestPickerData();
  assert.strictEqual(gas.toasts.length, 0, '直っていれば知らせない');
  assert.deepStrictEqual(gas.writes.filter(w => w.sheet === 'リクエスト'), [], '直っていれば書かない');
});

test('ID：setupRequestSheet() で、ID の列を左端に足して、前からある行に上から順に振る', () => {
  const {gas, g} = setup();
  const sheet = gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス部門からのフィードバック', 'サービス案', '担当者'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', '', '', '', ''],
    ['', '', '', '', '棄却', '', '', ''],   // リクエストの無い行には振らない
    ['食品スーパー', 'A社', 'チラシのデジタル化', 'チラシの効果を測りたい', '', '', '', '']
  ], {rows: 10, columns: 8});
  g.setupRequestSheet();
  assert.strictEqual(headersOf(sheet)[0], 'ID');
  assert.deepStrictEqual([2, 3, 4].map(row => cellOf(sheet, 'ID', row).getDisplayValue()), ['SR-0001', '', 'SR-0002']);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /ID の無い 2行に ID を振りました。/);
  assert.strictEqual(add(g, 'EC立ち上げ', '=在庫を店舗と共有したい').id, 'SR-0003');
});

test('ID：削除は ID で行を決める（同じリクエストの行が複数あっても、指定した ID の行だけを削除する）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  gas.asUser(() => ABCD.forEach((label, i) =>
    cellOf(request, label, 3).setValue(['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'][i])));
  g.getRequestPickerData();   // 3行目に SR-0002 が振られる
  const target = g.getRequestListData().registered.find(r => r.id === 'SR-0002');
  gas.asUser(() => request.insertRowsAfter(1, 1));   // ほかの人が上に行を足した（行がずれる）
  const result = g.deleteServiceRequest(target.row, target.key, target.id);
  assert.strictEqual(result.ok, true, result.message);
  assert.strictEqual(result.message, 'SR-0002（4行目）を削除しました。');
  assert.deepStrictEqual(plain(result.list.registered.map(r => r.id)), ['SR-0001']);
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
  assert.deepStrictEqual(gas.dump(old, 'A1:N2'), [
    ['ID', 'サブインダストリー', '得意先', '案件名', 'アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当', 'リクエスト',
      'サービス案', '主管本部', '担当者', '判断', 'サービス部門からのフィードバック'],
    ['SR-0001', '食品スーパー', 'A社', 'アプリ刷新', '', '', '', '', '会員の購買分析をしたい', '会員分析基盤', '', '', 'サービス化検討', '']
  ], 'メモ・取り込み日を削除し、ID は左端に、営業の4列は案件名の、主管本部・担当者はサービス案の、フィードバックは判断のすぐ右に差し込む');
  assert.strictEqual(old.getProtections().length, 0);
  assert.deepStrictEqual(shown(g.getRequestPickerData()), ['食品スーパー / A社 / チラシのデジタル化 / チラシの効果を測りたい',
    'ドラッグストア / C社 / EC立ち上げ / =在庫を店舗と共有したい'], '登録してあったリクエストはパネルに出さない');
  assert.strictEqual(rangeOf(gas.ss.getSheetByName('タスク管理').getRange('A2')), '__TASK_CHOICES!A1', 'タスク管理は登録したリクエストから選ぶ');
  assert.strictEqual(gas.dump(gas.ss.getSheetByName('__TASK_CHOICES'), 'A1')[0][0], 'SR-0001 会員の購買分析をしたい');
});

/** 今のシートの並び（判断の右にサービス案。メモ・追加日がある） */
function addCurrentRequestSheet(gas) {
  return gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス案', 'メモ', '追加日'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '会員分析基盤', '先方に確認', '2026/10/01']
  ], {rows: 20, columns: 8});
}

test('setupRequestSheet：確認してからメモ・追加日を削除し、案件名の右に営業の4列、判断の右にフィードバック、サービス案の右に主管本部・担当者を足す（ほかの値は残す）', () => {
  const {gas, g} = setup();
  const sheet = addCurrentRequestSheet(gas);
  sheet.getRange('E2:E20').setDataValidation(gas.global.SpreadsheetApp.newDataValidation()
    .requireValueInList(['未判断', 'サービス化検討', '棄却'], true).build());
  gas.addSheet('タスク管理', [['サービス案', 'タスク', '担当者', '期限', '状況', 'メモ']], {rows: 10, columns: 6});
  gas.confirmAnswer = 'YES';
  g.setupRequestSheet();

  assert.strictEqual(gas.alerts.length, 1);
  assert.match(gas.alerts[0].message, /「メモ」「追加日」の列は使わなくなりました。\n.*削除する列に入っている値は消えます/);
  assert.deepStrictEqual(headersOf(sheet), REQUEST_HEADERS, '新しく作ったときと同じ並びになる');
  assert.deepStrictEqual(gas.dump(sheet, 'A2:N2')[0],
    ['SR-0001', '食品スーパー', 'A社', 'アプリ刷新', '', '', '', '', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '', '']);
  ['サービス部門からのフィードバック', '主管本部'].concat(SALES).forEach(label =>
    assert.strictEqual(cellOf(sheet, label, 2).getDataValidation(), null, label + '：差し込んだ列は左の列のプルダウンを引き継がない'));
  assert.deepStrictEqual(listOf(cellOf(sheet, '判断', 2)), ['未判断', 'サービス化検討', '棄却']);
  assert.strictEqual(rangeOf(gas.ss.getSheetByName('タスク管理').getRange('A2')), '__TASK_CHOICES!A1', 'タスク管理は登録したリクエストから選ぶ');
  const toast = gas.toasts[gas.toasts.length - 1].message;
  assert.match(toast, /「メモ」「追加日」の列を削除しました。/);
  assert.match(toast, /「ID」「アカウント責任者部署」「アカウント責任者」「BX部署」「BX担当」「サービス部門からのフィードバック」「主管本部」「担当者」の列を足しました。/);
  assert.match(toast, /ID の無い 1行に ID を振りました。/);
  assert.match(toast, /タスク管理 に「本部」の列を足しました。/);

  gas.alerts.length = 0;
  g.setupRequestSheet();
  assert.strictEqual(gas.alerts.length, 0, '2回目は確認しない');
  assert.strictEqual(sheet.getLastColumn(), 14, '2回目は列を足さない');
});

test('setupRequestSheet：削除しないと答えたら、メモ・追加日は残して足りない列だけを足す', () => {
  const {gas, g} = setup();
  const sheet = addCurrentRequestSheet(gas);
  gas.confirmAnswer = 'NO';
  g.setupRequestSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A1:P2'), [
    REQUEST_HEADERS.concat(['メモ', '追加日']),
    ['SR-0001', '食品スーパー', 'A社', 'アプリ刷新', '', '', '', '', '会員の購買分析をしたい', 'サービス化検討', '', '会員分析基盤', '', '', '先方に確認', '2026/10/01']
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
  assert.match(gas.toasts[0].message, /判断・フィードバック・サービス案・主管本部・担当者が空の行が 2件あります。.*removeUntouchedRequests\(\)/);

  gas.confirmAnswer = 'NO';
  g.removeUntouchedRequests();
  assert.strictEqual(sheet.getLastRow(), 4, '「いいえ」なら消さない');

  gas.confirmAnswer = 'YES';
  g.removeUntouchedRequests();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /判断が「未判断」か空で、フィードバック・サービス案・主管本部・担当者が空の行が 2件あります/);
  assert.deepStrictEqual([2, 3].map(row => valuesOf(sheet, row, ABCD.concat(['サービス案']))), [
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
    '「新FMT」：「アカウント責任者部署」「アカウント責任者」「BX部署」「BX担当」の見出しがちょうど1つではないため、その列は空のまま登録します。',
    '「新FMT2」：読みません（test.gs の DIFF_RULES に requests: true がありません）。'
  ]);
});

test('diagnoseRequestSources：登録済みで取り込み元に見つからない行の行番号を出す（メモの代わり）', () => {
  const {gas, g, source} = setup();
  g.setupRequestSheet();
  add(g, 'アプリ刷新', '会員の購買分析をしたい');
  add(g, 'チラシのデジタル化', 'チラシの効果を測りたい');
  gas.asUser(() => source.getRange('F4').setValue('チラシの配布数を減らしたい'));
  gas.alerts.length = 0;
  g.diagnoseRequestSources();
  assert.match(gas.alerts[0].message, /「リクエスト」で取り込み元に見つからない行：3行目（取り込み元で書き換えか削除された可能性があります）。/);

  // 読めないシートがあるときは、見つからない行を出さない（読めなかったシートにあるかもしれないため）
  gas.get('DIFF_RULES')['新FMT2'].requests = true;
  gas.alerts.length = 0;
  g.diagnoseRequestSources();
  assert.doesNotMatch(gas.alerts[0].message, /見つからない行/);
  assert.match(gas.alerts[0].message, /「新FMT2」：/);
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

test('パネルを開いたときの読み取りの回数は、行の数によらない（行ごとに読まない。遅くならないように）', () => {
  const readsFor = count => {
    const rows = [];
    for (let i = 0; i < count; i++) rows.push(['食品スーパー', '得意先' + i, '案件' + i, 'リクエスト' + i]);
    const {gas, g} = setup(rows);
    g.setupRequestSheet();
    g.setupTaskSheet();
    const items = g.getRequestPickerData().items;
    items.slice(0, Math.min(count, 5)).forEach(item => g.addServiceRequest(item.key));
    g.getRequestPickerData();   // 1回目（確かめた記録を作る）
    gas.reads.length = 0;
    g.getRequestPickerData();
    return gas.reads.length;
  };
  const small = readsFor(5);
  const large = readsFor(25);
  assert.strictEqual(large, small, '読み取りの回数が行の数で増えている');
  assert.ok(small <= 20, '読み取りが多すぎます（' + small + '回）');
});
