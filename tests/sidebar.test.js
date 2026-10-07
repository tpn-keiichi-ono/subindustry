'use strict';

const {test, assert} = require('./lib/harness');
const {plain} = require('./lib/gas-mock');
const {setupProject} = require('./lib/fixture');

test('getSidebarSelection：選んだ行の得意先を返し、同じ行なら読み込みを省く', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'F5');
  const first = plain(g.getSidebarSelection(null));
  assert.deepStrictEqual(first, {sheet: '新FMT', row: 5, customer: 'Ｃ社', key: 'C社'});
  assert.deepStrictEqual(plain(g.getSidebarSelection({sheet: '新FMT', row: 5})), {sheet: '新FMT', row: 5, same: true});
});

test('getSidebarSelection：見出し行・対象外のシートでは理由を返す', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'A2');
  assert.strictEqual(g.getSidebarSelection(null).reason, 'header');
  gas.select(gas.ss.getSheetByName('変更履歴_差分'), 'A2');
  assert.strictEqual(g.getSidebarSelection(null).reason, 'sheet');
});

test('getSidebarBundle：得意先ごとの履歴と、シート・得意先ごとの最近の変更をまとめて返す', () => {
  const {gas, g, sheet} = setupProject();
  const data = plain(g.getCredentialDialogData('新FMT', 5));
  g.saveCredentialEntry({
    sheetName: '新FMT', row: 5, customer: data.customer, customerColumn: data.customerColumn, linkId: '',
    values: {plan: '2026/12/01', date: '', kind: 'オファリング', title: '', person: '田中', name: '提案', note: '',
      owner: '佐藤', ownerDept: ''}
  });
  g.recordDiffEdit(gas.edit(sheet, 'F5', '提案中'));

  const bundle = plain(g.getSidebarBundle());
  assert.deepStrictEqual(bundle.sheets, ['新FMT', '新FMT2']);
  assert.strictEqual(bundle.credentials['C社'].length, 1);
  assert.strictEqual(bundle.credentials['C社'][0].values.name, '提案');
  const changes = bundle.changes['新FMT']['C社'];
  assert.strictEqual(changes.length, 1);
  assert.deepStrictEqual(changes[0].changes, [{column: '状況', before: '', after: '提案中'}]);
  assert.strictEqual(changes[0].editor, 'user');
});

test('getSidebarBundle：読み取りだけで、シートに書き込まない', () => {
  const {gas, g} = setupProject();
  gas.writes.length = 0;
  g.getSidebarBundle();
  assert.deepStrictEqual(gas.writes, []);
});

test('開いたときの自動表示はやめた：前の版のトリガーが残っていても開かず、そのトリガーと設定を外す', () => {
  const {gas, g} = setupProject();
  const before = gas.triggers.length;
  g.ScriptApp.newTrigger('autoOpenHistorySidebar').forSpreadsheet(gas.ss).onOpen().create();
  g.PropertiesService.getDocumentProperties().setProperty('HS_AUTO_OFF_user@example.com', '1');
  assert.strictEqual(typeof g.setupHistorySidebarAutoOpen, 'undefined', '自動表示を設置する関数は無い');
  assert.strictEqual(typeof g.toggleHistorySidebarAutoOpen, 'undefined');

  g.autoOpenHistorySidebar({user: {getEmail: () => 'user@example.com'}});
  assert.strictEqual(gas.sidebars.length, 0, 'サイドバーは開かない');
  assert.strictEqual(gas.triggers.length, before, '残っていたトリガーを外す（ほかのトリガーはそのまま）');
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('HS_AUTO_OFF_user@example.com'), null);

  g.ScriptApp.newTrigger('autoOpenHistorySidebar').forSpreadsheet(gas.ss).onOpen().create();
  g.removeHistorySidebarAutoOpen();
  assert.strictEqual(gas.triggers.length, before);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /前の版の自動表示のトリガーを外しました（1件）/);
  g.removeHistorySidebarAutoOpen();
  assert.match(gas.toasts[gas.toasts.length - 1].message, /自動表示のトリガーはありません/);
});

test('メニュー「🟩RXビジネスMTG用」＞「履歴サイドバーを開く」で開く（対象のシートのときだけ）', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'E3');
  assert.strictEqual(g.openHistorySidebar(), true);
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'HistorySidebarView');
  assert.strictEqual(gas.alerts.length, 0);

  // 対象外のシート（DIFF_RULES に無い）では開かず、対象のシートを知らせる
  gas.select(gas.addSheet('集計', [['']], {rows: 5, columns: 3}), 'A1');
  assert.strictEqual(g.openHistorySidebar(), false);
  assert.strictEqual(gas.sidebars.length, 1, '開かない');
  assert.strictEqual(gas.alerts[0].message, 'このシートでは開けません。「新FMT」「新FMT2」のシートで開いてください。');
  assert.strictEqual(gas.writes.length, 0);
});

test('getSidebarBundle：変更は行（シート・得意先）ごとに新しい方から上限まで。超えた行には印を付ける', () => {
  const {gas, g, sheet} = setupProject();
  const limit = gas.get('HS_OPTIONS').eventsPerCustomer;
  for (let i = 0; i <= limit; i++) g.recordDiffEdit(gas.edit(sheet, 'F3', '状況' + i));   // A社に limit + 1 回
  g.recordDiffEdit(gas.edit(sheet, 'F4', '保留'));                                        // B社に1回

  const bundle = plain(g.getSidebarBundle());
  const events = bundle.changes['新FMT']['A社'];
  assert.strictEqual(events.length, limit);
  assert.strictEqual(events[0].changes[0].after, '状況' + limit, '新しい順になっていません');
  assert.strictEqual(events[limit - 1].changes[0].after, '状況1', '最も古い1回が残っています');
  assert.deepStrictEqual(bundle.changesMore, {'新FMT': {'A社': true}});
  assert.strictEqual(bundle.eventsPerCustomer, limit);
  assert.strictEqual(bundle.changes['新FMT']['B社'].length, 1);
});

test('openHistorySidebar：サイドバーの名称は「得意先別の履歴情報」', () => {
  const {gas, g} = setupProject();
  g.openHistorySidebar();
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'HistorySidebarView');
  assert.strictEqual(gas.sidebars[0].html.title, '得意先別の履歴情報');
});

/** 33シナリオ攻略先リスト（1行目が見出し、E列が得意先、M列がアカウントプランの URL） */
function addPlanList(gas, rows) {
  const header = new Array(13).fill('');
  header[4] = '得意先';
  header[12] = 'アカウントプラン';
  const sheet = gas.addSheet('33シナリオ攻略先リスト', [header].concat(rows.map(([customer, url]) => {
    const row = new Array(13).fill('');
    row[4] = customer;
    row[12] = url || '';
    return row;
  })), {rows: 20, columns: 13});
  return sheet;
}

test('getSidebarBundle：33シナリオ攻略先リストの E列（得意先）と M列（URL）から、得意先ごとのアカウントプランを返す', () => {
  const {gas, g} = setupProject();
  const sheet = addPlanList(gas, [
    ['A社', 'https://docs.google.com/spreadsheets/d/a/edit'],
    ['B 社', ' https://docs.google.com/spreadsheets/d/b/edit '],   // 空白・全角半角の違いは無視
    ['Ｃ社', ''],                                               // まだ作っていない
    ['D社', 'javascript:alert(1)'],                             // http(s) 以外はリンクにしない
    ['E社', '=HYPERLINK("https://docs.google.com/spreadsheets/d/e/edit","開く")'],
    ['F社', '開く'],                                            // セルの文字にリンクを付けたもの
    ['A社', 'https://docs.google.com/spreadsheets/d/a2/edit']   // 同じ得意先は上の行を使う
  ]);
  gas.asUser(() => sheet.getRange('M7').setRichTextValues([[
    gas.global.SpreadsheetApp.newRichTextValue().setText('開く')
      .setLinkUrl('https://docs.google.com/spreadsheets/d/f/edit').build()
  ]]));

  const bundle = plain(g.getSidebarBundle());
  assert.deepStrictEqual(bundle.accountPlans, {
    'A社': 'https://docs.google.com/spreadsheets/d/a/edit',
    'B社': 'https://docs.google.com/spreadsheets/d/b/edit',
    'E社': 'https://docs.google.com/spreadsheets/d/e/edit',
    'F社': 'https://docs.google.com/spreadsheets/d/f/edit'
  });
  assert.strictEqual(bundle.accountPlanSheet, '33シナリオ攻略先リスト');
  // サイドバーの得意先キー（新FMT の得意先）と同じ形で照合できる
  gas.select(gas.ss.getSheetByName('新FMT'), 'D3');
  assert.ok(bundle.accountPlans[g.getSidebarSelection(null).key]);
});

test('getSidebarBundle：33シナリオ攻略先リストが無ければ、アカウントプランは空', () => {
  const {g} = setupProject();
  assert.deepStrictEqual(plain(g.getSidebarBundle()).accountPlans, {});
});

test('画面：まだ承認していない人への案内は、実際のメニュー名・項目名と一致する（メニューの名前は CRED_OPTIONS.menuTitle を渡す）', () => {
  const fs = require('fs');
  const path = require('path');
  const html = fs.readFileSync(path.join(__dirname, '..', 'HistorySidebarView.html'), 'utf8');
  const page = fs.readFileSync(path.join(__dirname, '..', 'AuthorizeView.html'), 'utf8');
  const {gas, g, sheet} = setupProject();
  g.onOpen({source: gas.ss});
  const title = gas.get('CRED_OPTIONS').menuTitle;
  const menu = gas.menus.find(m => m.title === title);
  const item = menu.items.find(i => i.fn === 'authorizeHistoryFeatures');
  assert.ok(html.includes('メニュー「<?= menuTitle ?>」＞「' + item.label + '」'), 'サイドバーの案内と項目の名前が違います');
  assert.ok(html.includes('」＞「' + item.label + '」を実行し'), 'サイドバーの案内（ボタンが無いとき）と項目の名前が違います');
  assert.ok(page.includes('メニュー「<?= menuTitle ?>」＞「' + item.label + '」'), '承認のページの案内と項目の名前が違います');
  assert.ok(!/履歴機能」/.test(html + page), '前のメニューの名前が残っています');

  gas.select(sheet, 'E3');
  g.openHistorySidebar();
  assert.strictEqual(gas.sidebars[0].html.data.menuTitle, title, 'サイドバーにメニューの名前を渡す');
  assert.strictEqual(g.doGet({}).data.menuTitle, title, '承認のページにメニューの名前を渡す');
});

test('openHistorySidebar：承認用のウェブアプリの URL（…/exec）だけをサイドバーに渡す', () => {
  const {gas, g} = setupProject();
  const options = gas.get('HS_OPTIONS');
  const open = url => {
    options.authorizeUrl = url;
    gas.sidebars.length = 0;
    g.openHistorySidebar();
    return gas.sidebars[0].html.data.authorizeUrl;
  };
  assert.strictEqual(open(''), '', '未設定なら空（サイドバーはメニューでの承認を案内する）');
  const url = 'https://script.google.com/a/macros/example.com/s/AKfycbx123/exec';
  assert.strictEqual(open('  ' + url + '  '), url);
  assert.strictEqual(open('https://script.google.com/macros/s/AKfycbx123/dev'), '', 'テスト用の …/dev は使わない');
  assert.strictEqual(open('http://script.google.com/macros/s/AKfycbx123/exec'), '');
  assert.strictEqual(open('https://example.com/exec'), '');
  options.authorizeUrl = '';
});

test('doGet：承認が済んでいれば完了のページ、足りなければもう一度承認するリンクを出す（シートは読み書きしない）', () => {
  const {gas, g} = setupProject();
  gas.writes.length = 0;
  const done = g.doGet({});
  assert.strictEqual(done.file, 'AuthorizeView');
  assert.deepStrictEqual(plain(done.data), {status: 'done', retryUrl: '', menuTitle: gas.get('CRED_OPTIONS').menuTitle});
  assert.strictEqual(done.title, gas.get('CRED_OPTIONS').menuTitle + 'の権限の承認');

  gas.authRequired = true;
  const missing = g.doGet({});
  assert.strictEqual(missing.data.status, 'missing');
  assert.match(missing.data.retryUrl, /^https:\/\/script\.google\.com\//);
  assert.deepStrictEqual(gas.writes, []);
});

test('getAuthorizationLink：承認が足りなければ Google の承認の画面の URL、取れなければ承認用のウェブアプリの URL を返す（読むだけ）', () => {
  const {gas, g} = setupProject();
  const options = gas.get('HS_OPTIONS');
  gas.writes.length = 0;
  assert.deepStrictEqual(plain(g.getAuthorizationLink()), {required: false, url: ''}, '承認が済んでいれば行き先は要らない');

  gas.authRequired = true;
  assert.deepStrictEqual(plain(g.getAuthorizationLink()), {required: true, url: 'https://script.google.com/macros/d/script-id/authorize'},
    'ウェブアプリを用意していなくてもボタンを出せる');

  const webApp = 'https://script.google.com/a/macros/example.com/s/AKfycbx123/exec';
  options.authorizeUrl = webApp;
  const original = g.ScriptApp.getAuthorizationInfo;
  g.ScriptApp.getAuthorizationInfo = () => ({getAuthorizationStatus: () => 'REQUIRED', getAuthorizationUrl: () => null});
  try {
    assert.deepStrictEqual(plain(g.getAuthorizationLink()), {required: true, url: webApp}, '承認の画面の URL が取れなければウェブアプリ');
  } finally {
    g.ScriptApp.getAuthorizationInfo = original;
    options.authorizeUrl = '';
  }
  assert.deepStrictEqual(gas.writes, []);
});

test('画面：サイドバーは承認用の URL を data-authorize-url で受け取り、承認のページは状態を data-status で受け取る', () => {
  const fs = require('fs');
  const path = require('path');
  const sidebar = fs.readFileSync(path.join(__dirname, '..', 'HistorySidebarView.html'), 'utf8');
  assert.ok(sidebar.includes('<body data-authorize-url="<?= authorizeUrl ?>" data-menu-title="<?= menuTitle ?>">'));
  ['auth', 'authText', 'authBtn', 'authNote'].forEach(id => assert.ok(sidebar.includes('id="' + id + '"'), id + ' がありません'));
  assert.ok(sidebar.includes('.getAuthorizationLink()'), '承認の案内を出すときに、ボタンの行き先を問い合わせる');
  const page = fs.readFileSync(path.join(__dirname, '..', 'AuthorizeView.html'), 'utf8');
  assert.ok(page.includes('data-status="<?= status ?>"') && page.includes('data-retry-url="<?= retryUrl ?>"'));
});
