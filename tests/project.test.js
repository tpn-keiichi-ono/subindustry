'use strict';

const {test, assert} = require('./lib/harness');
const {createGas} = require('./lib/gas-mock');

test('リポジトリ直下の *.gs を1つのグローバルスコープに読み込める（同じ名前の const がない）', () => {
  const gas = createGas();
  ['DIFF_RULES', 'DIFF_OPTIONS', 'CRED_OPTIONS', 'CHG_OPTIONS', 'HS_OPTIONS'].forEach(name => {
    assert.strictEqual(typeof gas.get(name), 'object', name);
  });
});

test('onOpen のメニューは左から「🟩RXビジネスMTG用」「🟪RXサービスMTG用」（アカウントプランシート作成は出さない）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  assert.deepStrictEqual(gas.menus.map(m => m.title), ['🟩RXビジネスMTG用', '🟪RXサービスMTG用']);
  assert.strictEqual(typeof gas.global.showLoadingDialog, 'function', 'アカウントプランシート作成はエディタから実行できる');
});

test('CRED_OPTIONS.accountPlanMenu を true にすると、アカウントプランシート作成のメニューを右端に出す', () => {
  const gas = createGas();
  gas.get('CRED_OPTIONS').accountPlanMenu = true;
  gas.global.onOpen({source: gas.ss});
  assert.deepStrictEqual(gas.menus.map(m => m.title), ['🟩RXビジネスMTG用', '🟪RXサービスMTG用', 'アカウントプランシート作成']);
  assert.deepStrictEqual(gas.menus[2].items, [{label: '作成開始', fn: 'showLoadingDialog'}]);
});

test('「🟩RXビジネスMTG用」のメニューは「履歴サイドバーを開く」「権限を承認する（初回のみ）」だけ（ほかの処理はエディタから実行する）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  const menu = gas.menus.find(m => m.title === '🟩RXビジネスMTG用');
  assert.deepStrictEqual(menu.items, [
    {label: '履歴サイドバーを開く', fn: 'openHistorySidebar'},
    {label: '権限を承認する（初回のみ）', fn: 'authorizeHistoryFeatures'}
  ]);
  // メニューから外した関数も、エディタから実行できるよう残っている
  ['openHistorySidebar', 'openCredentialDialog', 'openChangeHistoryDialog',
    'setupCredentialLauncher', 'setupChangeHistoryLauncher', 'removeHistorySidebarAutoOpen',
    'addSampleHistory', 'removeSampleHistory'].forEach(fn => {
    assert.strictEqual(typeof gas.global[fn], 'function', fn + ' がありません');
  });
});

test('権限を承認する：すべての権限の承認を確かめてから、対象のシートならサイドバーを開き直す', () => {
  const gas = createGas();
  gas.select(gas.addSheet('新FMT', [['新FMT']], {rows: 5, columns: 3}), 'A2');
  gas.global.authorizeHistoryFeatures();
  assert.deepStrictEqual(gas.scopeChecks, ['FULL']);
  assert.strictEqual(gas.sidebars.length, 1);
  assert.strictEqual(gas.sidebars[0].html.file, 'HistorySidebarView');
  assert.strictEqual(gas.toasts.length, 1);
  assert.strictEqual(gas.toasts[0].title, '🟩RXビジネスMTG用');
  assert.match(gas.toasts[0].message, /履歴サイドバーを開きました/);
  assert.strictEqual(gas.writes.length, 0, 'シートに書き込まない');
});

test('権限を承認する：対象外のシートでは、承認だけしてサイドバーは開かない', () => {
  const gas = createGas();
  gas.select(gas.addSheet('メモ用', [['']], {rows: 5, columns: 3}), 'A1');
  gas.global.authorizeHistoryFeatures();
  assert.deepStrictEqual(gas.scopeChecks, ['FULL']);
  assert.strictEqual(gas.sidebars.length, 0);
  assert.strictEqual(gas.alerts.length, 0, '承認のときは画面を止めない（トーストで知らせる）');
  assert.strictEqual(gas.toasts[0].message, '権限は承認済みです。履歴サイドバーは「新FMT」「新FMT2」のシートで開いてください。');
});

test('メニュー項目の関数がすべて存在し、内部用（末尾 _）ではない', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  const items = [].concat(...gas.menus.map(m => m.items)).filter(item => item.fn);
  assert.ok(items.length > 0);
  items.forEach(item => {
    assert.strictEqual(typeof gas.global[item.fn], 'function', item.fn + ' がありません');
    assert.ok(!item.fn.endsWith('_'), item.fn + ' は内部用の名前です');
  });
});

test('単純トリガーの onEdit は置かない（リクエストは選択パネルで登録する。編集のたびに動く処理を作らない）', () => {
  const gas = createGas();
  assert.strictEqual(typeof gas.global.onEdit, 'undefined');
  assert.strictEqual(typeof gas.global.doGet, 'function', 'doGet は1つ（HistorySidebar.gs）');
});

test('onOpen が単純トリガーとして動く（承認前でも失敗しない範囲の処理だけ）', () => {
  const gas = createGas();
  gas.global.onOpen({source: gas.ss});
  assert.strictEqual(gas.writes.length, 0, 'onOpen ではシートに書き込まない');
  assert.strictEqual(gas.triggers.length, 0, 'onOpen ではトリガーを作らない');
  assert.strictEqual(gas.sidebars.length, 0, 'onOpen ではサイドバーを開かない（利用者がメニューから開く）');
});

test('本物のシート名の設定：DIFF_RULES・前の名前（formerNames）・リクエスト・タスク管理のシート名が重ならない', () => {
  const gas = createGas({realSheetNames: true});
  const rules = gas.get('DIFF_RULES');
  const diff = gas.get('DIFF_OPTIONS');
  const svc = gas.get('SVC_OPTIONS');
  const names = Object.keys(rules);
  assert.ok(names.length > 0);
  names.forEach(name => {
    const rule = rules[name];
    assert.ok(Number.isInteger(rule.headerRow) && rule.headerRow >= 1, name + ' の headerRow');
    assert.ok(Array.isArray(rule.ranges) && rule.ranges.length > 0, name + ' の ranges');
    assert.ok(!rule.formerNames || Array.isArray(rule.formerNames), name + ' の formerNames は配列');
  });
  assert.ok(names.some(name => rules[name].requests), 'リクエストを読むシート（requests: true）が1つ以上ある');

  const former = [].concat(...names.map(name => rules[name].formerNames || []));
  former.forEach(f => assert.ok(!names.includes(f), '前の名前「' + f + '」が今のシート名と同じ'));
  assert.strictEqual(new Set(former).size, former.length, '同じ前の名前が2つのシートにある');

  // 機能ごとのシート名は、追跡するシートやほかの機能のシートと重ならない
  const others = [diff.logSheet, diff.rowSnapshotSheet, gas.get('CRED_OPTIONS').logSheet,
    svc.requestSheet, svc.legacyListSheet, svc.taskChoiceSheet, gas.get('TASK_OPTIONS').sheet, gas.get('MTG_OPTIONS').statusTable.sheet,
    gas.get('MTG_OPTIONS').index.sheet];
  assert.strictEqual(new Set(others).size, others.length, '機能ごとのシート名が重なっている');
  others.forEach(name => {
    assert.ok(!names.includes(name) && !former.includes(name), '「' + name + '」が追跡するシートと同じ名前');
    assert.ok(!name.startsWith(diff.snapshotPrefix) || name === diff.snapshotPrefix, name);
  });
});

test('本物のシートの見出し（2026/10/29 のスーパー・GMS の2行目。P〜U列に FY別の売上、受注月は「受注月(yy/mm)」、売上開始月は無い。A〜AF の32列）に、会議サマリ・リクエストの取り込みが読む列がそろっている', () => {
  const gas = createGas({realSheetNames: true});
  const mtg = gas.get('MTG_OPTIONS');
  const g = gas.global;
  const headers = ['最終更新日時\n（自動入力）', 'No.', 'サブインダストリー', '得意先', '先方部門', '先方担当役職', '先方担当氏名',
    '案件名', 'RXシナリオ構成', 'サービスパッケージ合算（百万円）', '案件ステータス',
    '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月\n(yy/mm)', '提案開始月\n(yy/mm)',
    'FY26前半期売上（百万）', 'FY26後半期売上（百万）', 'FY27前半期売上（百万）', 'FY27後半期売上（百万）', 'FY28前半期売上（百万）', 'FY28後半期売上（百万）',
    '活動状況', '活動における課題', '次回アクション',
    '商談記録（ＵＲＬ）', 'アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当', '想定される価値創出のケース／期待される導入効果',
    '提案を勝ち取るための戦略・差異化要素', 'サービスのリクエスト'];
  const at = label => headers.findIndex(h => g.credNormalize_(h) === g.credNormalize_(label));
  // 売上開始月は 2026/10/29 にシートから無くなった（計画に対する積み上げ・売上開始月の確認は「データなし」）
  Object.keys(mtg.columns).filter(key => key !== 'delivery' && key !== 'startMonth').forEach(key =>
    assert.ok(g.mtgColumnIndex_(headers, key) >= 0, '見出し：' + mtg.columns[key]));
  assert.strictEqual(g.mtgColumnIndex_(headers, 'startMonth'), -1);
  assert.strictEqual(g.mtgColumnIndex_(headers, 'orderMonth'), 13, '受注月は「受注月(yy/mm)」の N列（columnAliases）');
  // P〜U列（FY別の売上）は資料に含めない
  assert.deepStrictEqual(headers.map((h, i) => (g.mtgIgnoredHeader_(h) ? i : -1)).filter(i => i >= 0), [15, 16, 17, 18, 19, 20], 'P〜U列だけ');
  assert.ok(at('デリバリー担当') < 0 && mtg.columns.delivery === '', 'デリバリー担当の列は無いので使わない設定（入力内容のチェック「デリバリー担当なし」は出さない）');
  mtg.required.forEach(label => assert.ok(g.mtgHeaderIndex_(headers, label) >= 0, '必須項目：' + label));
  assert.strictEqual(at(mtg.columns.owner), headers.indexOf('アカウント責任者'), '「アカウント責任者部署」と取り違えない');
  const stamp = g.diffNormalizeStampHeader_(gas.get('DIFF_OPTIONS').stampHeader);
  assert.deepStrictEqual(headers.map((h, i) => (g.diffNormalizeStampHeader_(h) === stamp ? i : -1)).filter(i => i >= 0), [0], '最終更新日時はA列の1つだけ');
  // 差分追跡の範囲（D3:AF）に、会議サマリが前回の値を見る列（得意先〜サービスのリクエスト）が入っている
  const rule = gas.get('DIFF_RULES')['スーパー・GMS'];
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rule.headerRow, rule.ranges])), [2, ['D3:AF']]);
  ['customer', 'project', 'status', 'amount', 'weighted', 'orderMonth', 'activity', 'scenario'].forEach(key => assert.ok(g.mtgColumnIndex_(headers, key) + 1 >= 4, key));
  assert.strictEqual(g.mtgColumnIndex_(headers, 'scenario'), 8, 'RXシナリオ構成は I列（案件名の右）');
  // コンビニ・地方銀行・生協は「RXシナリオ構成（サービス）」（columnAliases で読む）
  assert.strictEqual(g.mtgColumnIndex_(headers.map(h => (h === 'RXシナリオ構成' ? 'RXシナリオ構成（サービス）' : h)), 'scenario'), 8);
  assert.ok(headers.length <= 32, 'AF 列まで（追跡範囲 D3:AF。シートが狭くても、追跡範囲はシートの中まで読む。diffRuleRange_）');
  // サービスリクエストの取り込み（ServiceManagement.gs）が読む列も、見出しがちょうど1つずつある（列を足しても見出しで探すのでずれない）
  const svc = gas.get('SVC_OPTIONS');
  svc.sourceHeaders.concat(svc.sourceOptionalHeaders).forEach(label => assert.strictEqual(g.taskHeaderCount_(headers, label), 1, 'リクエストの取り込み：' + label));
  assert.strictEqual(headers.indexOf('サービスのリクエスト') + 1, 32, 'サービスのリクエストは AF列（追跡範囲 D3:AF の中）');
});

test('本物の設定：11のサブインダストリーのシート（2026/10/16 にスーパー・GMS と同じ見出しにそろえた）を、同じ見出しの行・追跡範囲で追跡し、リクエストも読む', () => {
  const gas = createGas({realSheetNames: true});
  const rules = gas.get('DIFF_RULES');
  const subs = ['スーパー・GMS', 'コンビニ', '百貨店', '化粧品・トイレタリー', '飲料・食品', '自動車販売', '住設', '地方銀行', '電鉄・不動産', 'エネルギー', '生協'];
  assert.deepStrictEqual(Object.keys(rules), subs);
  subs.forEach(name => assert.deepStrictEqual(JSON.parse(JSON.stringify([rules[name].headerRow, rules[name].ranges, rules[name].requests])),
    [2, ['D3:AF'], true], name));
});

test('本物の設定：クレデンシャル・オファリング履歴の管理・登録は一旦使わない（CRED_OPTIONS.enabled: false）', () => {
  const gas = createGas({realSheetNames: true});
  assert.strictEqual(gas.get('CRED_OPTIONS').enabled, false);
  assert.strictEqual(gas.global.credEnabled_(), false);
});

test('本物の設定：事前資料の格納先は全サブインダストリー共通の1つ（MTG_OPTIONS.meetingFolder。シートごとの meetingFolder は無い）', () => {
  const gas = createGas({realSheetNames: true});
  const rules = gas.get('DIFF_RULES');
  assert.ok(gas.global.mtgFolderIdOf_(gas.get('MTG_OPTIONS').meetingFolder), 'MTG_OPTIONS.meetingFolder');
  Object.keys(rules).forEach(name => assert.strictEqual(rules[name].meetingFolder, undefined, name + ' に meetingFolder が残っている'));
  // 統合版（全サブインダストリー）も作り、会議のフォルダに格納する。名前はシート名・インデックスの見出しと重ならない
  const combined = gas.get('MTG_OPTIONS').combined;
  assert.strictEqual(combined.enabled, true);
  assert.strictEqual(combined.meetingFolder, undefined, '統合版も同じ格納先');
  const index = gas.get('MTG_OPTIONS').index;
  assert.ok(!rules[combined.name] && [index.dateHeader, index.kindHeader].indexOf(combined.name) < 0);
});

test('本物の設定：RX Service Meeting の事前資料は、ひな形の3枚目（紫）を表紙に使い、インデックスには RX Business Meeting と別の会議名で書く', () => {
  const gas = createGas({realSheetNames: true});
  const mtg = gas.get('MTG_OPTIONS');
  const smtg = gas.get('SMTG_OPTIONS');
  assert.strictEqual(mtg.templatePages.service, 3);
  assert.ok(smtg.index.kind && smtg.index.kind !== mtg.index.kind, '会議名で RX Business Meeting の行と分ける');
  assert.notStrictEqual(smtg.fileName, mtg.fileName);
  // 会議のフォルダは同じ親フォルダに並ぶので、名前に会議の名前を入れて見分ける（第N回も会議ごと）
  assert.ok(/\{meeting\}/.test(mtg.meetingFolderName));
  assert.ok(gas.global.mtgFolderIsMeeting_('261113_' + smtg.meetingName + '_第3回', smtg.meetingName));
  assert.ok(!gas.global.mtgFolderIsMeeting_('261113_' + smtg.meetingName + '_第3回', mtg.meetingName));
  // 判断・状況の選択肢（前回からの変化の記録は選択肢の番号で残す）
  assert.deepStrictEqual(Array.from(gas.global.smtgDecisionOptions_()), ['未判断', 'サービス化検討', '棄却']);
  assert.deepStrictEqual(Array.from(gas.global.smtgStatusOptions_()), ['未着手', '対応中', '完了']);
});

test('本物の設定：カレンダーの予定の名前で、RX Business Meeting・RX Service Meeting を見分け、インデックスの会議名と同じ名前で書く', () => {
  const gas = createGas({realSheetNames: true});
  const g = gas.global;
  assert.strictEqual(gas.get('MTG_OPTIONS').index.kindHeader, '会議名');
  assert.strictEqual(g.mschKindName_('business'), gas.get('MTG_OPTIONS').index.kind);
  assert.strictEqual(g.mschKindName_('service'), gas.get('SMTG_OPTIONS').index.kind);
  assert.strictEqual(g.mschKeyOfTitle_('RX Service Meeting'), 'service');
  assert.strictEqual(g.mschKeyOfTitle_('RX Business MTG（パイプライン会議）'), 'business');
  assert.strictEqual(g.mschKeyOfTitle_('RX Business Meeting'), 'business');
  assert.strictEqual(g.mschKeyOfTitle_('部会'), '');
});

test('オポチュニティキャンバスは廃止（2026/10/26）：サーバーの関数・モーダル・サイドバーのボタンが無い', () => {
  const fs = require('fs');
  const path = require('path');
  const root = path.join(__dirname, '..');
  ['OpportunityCanvas.gs', 'OpportunityCanvasDialog.html'].forEach(f => assert.ok(!fs.existsSync(path.join(root, f)), f + ' が残っています'));
  const gas = createGas();
  ['openOpportunityCanvas', 'openOpportunityCanvasFromSidebar', 'getOpportunityCanvasData'].forEach(fn =>
    assert.strictEqual(typeof gas.global[fn], 'undefined', fn + ' が残っています'));
  const sidebar = fs.readFileSync(path.join(root, 'HistorySidebarView.html'), 'utf8');
  assert.ok(!/キャンバス|Canvas/.test(sidebar), 'サイドバーにキャンバスが残っています');
});
