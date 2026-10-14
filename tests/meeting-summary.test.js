'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/*
 * 会議サマリ（MeetingSummary.gs）
 * 追跡シート（2行目が見出し。追跡範囲は D3:AF）：
 *   D 得意先 / E 案件名 / F 案件ステータス / G 想定売上規模（百万） / H 期待値調整済（改行）想定売上規模（百万） / I 受注月 /
 *   J 活動状況 / K 活動における課題 / L アカウント責任者 / AG 最終更新日時 / AH クレデンシャル（改行）オファリング登録 / AI 変更履歴
 */
const WIDTH = 35;
const COL = {customer: 4, project: 5, status: 6, amount: 7, weighted: 8, orderMonth: 9, activity: 10, issue: 11, owner: 12,
  stamp: 33, credButton: 34, changeButton: 35};
const HEADERS = {customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', orderMonth: '受注月', activity: '活動状況', issue: '活動における課題',
  owner: 'アカウント責任者', stamp: '最終更新日時', credButton: 'クレデンシャル\nオファリング登録', changeButton: '変更履歴'};
const STATUSES = ['初回訪問', '提案中', '受注見込み', '受注', '失注'];

/** 今月（日本時間）から offset か月ずらした受注月 {y, m, text: '2026/09'} */
function month(offset) {
  const now = new Date(Date.now() + 9 * 3600000);
  const d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + offset, 1));
  return {y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, text: d.getUTCFullYear() + '/' + String(d.getUTCMonth() + 1).padStart(2, '0')};
}

/** 得意先・案件名・ステータス・想定・期待値・受注月・活動状況・課題・責任者 */
const DEALS = [
  ['A社', 'アプリ刷新', '初回訪問', '30', '10', month(-1).text, '初回訪問', '', '佐藤'],     // 3行目：ステータスが動く（受注月は先月）
  ['B社', 'EC立ち上げ', '受注', '50', '50', month(-2).text, '契約済み', '', '鈴木'],         // 4行目：受注（終わった案件）
  ['C社', '店舗什器', '提案中', '20', '5', month(2).text, '見積提出', '', '田中'],           // 5行目：失注になる
  ['D社', '物流改善', '', '', '', '', '', '', ''],                                         // 6行目：必須項目が空
  ['E社', '販促DX', '提案中', '80', '40', month(1).text, '提案書を送付', '決裁者に会えていない', '伊藤'],   // 7行目：停滞
  ['F社', 'ブランド刷新', '提案中', '10', '3', month(2).text, '', '', '高橋']               // 8行目：削除する
];

function row(values) {
  const line = new Array(WIDTH).fill('');
  Object.keys(values).forEach(key => { line[COL[key] - 1] = values[key]; });
  return line;
}

function setup(options) {
  const o = options || {};
  const gas = createGas();
  const g = gas.global;
  ['新FMT', '新FMT2'].forEach(name => {
    const deals = name === '新FMT' ? DEALS : DEALS.slice(0, 2);
    const lines = [row({customer: name}), row(HEADERS)].concat(deals.map(d => row({
      customer: d[0], project: d[1], status: d[2], amount: d[3], weighted: d[4], orderMonth: d[5], activity: d[6], issue: d[7], owner: d[8]
    })));
    const sheet = gas.addSheet(name, lines, {rows: 20, columns: WIDTH});
    sheet.getRange('F3:F20').setDataValidation(g.SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  });
  g.setupDiffTracking();
  const sheet = gas.ss.getSheetByName('新FMT');
  if (o.edits !== false) {
    g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));                 // A社：初回訪問 → 提案中
    g.recordDiffEdit(gas.edit(sheet, 'J3', '10/10 提案書を提出'));       // A社：活動状況
    g.recordDiffEdit(gas.edit(sheet, 'F5', '失注'));                   // C社：提案中 → 失注
    g.recordDiffEdit(gas.edit(sheet, 'J6', '10/12 打ち合わせ'));         // D社：活動状況
    gas.asUser(() => sheet.deleteRow(8));                             // F社の行を削除
    g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));
  }
  // E社は40日前から更新なし（最終更新日時を手で古くする。追跡範囲の外の列）
  gas.asUser(() => sheet.getRange(7, COL.stamp).setValue(new gas.Date(Date.now() - 40 * 86400000)));
  gas.writes.length = 0;
  gas.toasts.length = 0;
  return {gas, g, sheet};
}

/** 1つのシートの中身（PDF にする前の形）を作る（直近14日の変化。初回と同じ）。trackedSince：変更の記録を始めた時刻 */
function summarize(g, name, options) {
  const o = options || {};
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  const since = new Date(Date.now() - 14 * 86400000);
  return g.mtgSummarize_(ss, name, {
    since, first: true, final: true, now: new Date(),
    trackedSince: 'trackedSince' in o ? o.trackedSince : g.mtgLogStart_(ss),
    log: g.mtgReadLog_(ss, since),
    snapshots: g.mtgReadSnapshots_(ss, since),
    credentials: g.hsCredentialIndex_(ss)
  });
}

const label = item => [item.customer, item.project].filter(Boolean).join('／');

/* ---------------- 中身 ---------------- */

test('① トータルサマリー：ステータスごとの件数・金額をプルダウンの順に（進行中 → 受注 → 失注）。進行中の合計', () => {
  const {g} = setup();
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.pipeline.map(p => [p.status, p.kind, p.count, p.amount]), [
    ['提案中', 'active', 2, 110],
    ['（未設定）', 'active', 1, 0],
    ['受注', 'won', 1, 50],
    ['失注', 'lost', 1, 20]
  ]);
  assert.deepStrictEqual([s.kpi.deals, s.kpi.active, s.kpi.won, s.kpi.lost, s.kpi.amount, s.kpi.weighted], [5, 3, 1, 1, 110, 50]);
});

test('② 前回からの変化：変更の記録を得意先＋案件名ごとにまとめ、ステータスの動き（受注・失注を先に）と更新した項目を出す', () => {
  const {g} = setup();
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.statusChanged, c.from, c.to, c.toKind, c.columns.join('・')]), [
    ['C社／店舗什器', true, '提案中', '失注', 'lost', ''],
    ['A社／アプリ刷新', true, '初回訪問', '提案中', 'active', '活動状況'],
    ['D社／物流改善', false, null, null, 'active', '活動状況']
  ]);
  assert.strictEqual(s.changes[1].activity, '10/10 提案書を提出', '活動状況は今の値');
  assert.deepStrictEqual([s.kpi.changed, s.kpi.statusChanged], [3, 2]);
});

test('③ 停滞・失注：最終更新が30日以上前の進行中の案件（想定売上の大きい順）と、前回から失注になった案件・削除した行', () => {
  const {g} = setup();
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.stale.map(d => [label(d), d.days, d.issue]), [['E社／販促DX', 40, '決裁者に会えていない']],
    '動いた A社・D社 と、終わった B社・C社 は出さない');
  assert.deepStrictEqual(s.closed.map(d => [label(d), d.removed, d.to]).sort(), [
    ['C社／店舗什器', false, '失注'],
    ['F社／ブランド刷新', true, '']
  ]);
});

test('③ 停滞：最終更新日時が空の行は、変更の記録を始めた日から動いていないとみなす（30日たっていなければ停滞にしない）', () => {
  const {g} = setup({edits: false});
  const started = new Date(Date.now() - 50 * 86400000);
  const s = plain(summarize(g, '新FMT', {trackedSince: started}));
  assert.deepStrictEqual(s.stale.map(d => [label(d), d.days, d.noRecord]),
    [['E社／販促DX', 40, false], ['A社／アプリ刷新', 50, true], ['C社／店舗什器', 50, true], ['F社／ブランド刷新', 50, true],
      ['D社／物流改善', 50, true]], '想定売上の大きい順');
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: new Date(Date.now() - 10 * 86400000)})).stale.map(label),
    ['E社／販促DX'], '記録を始めて10日なら、空の行は判断しない');
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: null})).stale.map(label), ['E社／販促DX'], '記録が無ければ判断しない');
  assert.strictEqual(g.mtgLayout_(summarize(g, '新FMT', {trackedSince: started})).values[25][2], '50日以上');
});

test('④ アラート：予定日を過ぎても実施日が空・受注月を過ぎても進行中・必須項目が空（サンプルの変更の記録は変化に入れない）', () => {
  const {gas, g, sheet} = setup();
  gas.select(sheet, 'D3');
  g.addSampleHistory();   // A社に、3日前が予定日で未実施のオファリングなどを足す（変更の記録は SMP-）
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.alerts.map(a => [a.type, label(a)]), [
    ['予定超過', 'A社'],
    ['受注月超過', 'A社／アプリ刷新'],
    ['未入力', 'D社／物流改善']
  ]);
  assert.match(s.alerts[0].text, /^オファリング「【サンプル】販促DXの試験導入のご提案」 予定 \d{2}\/\d{2}（3日超過） 担当 サンプル 担当$/);
  assert.strictEqual(s.alerts[1].text, '受注月 ' + month(-1).y + '/' + month(-1).m + ' を過ぎても「提案中」');
  assert.strictEqual(s.alerts[2].text, '空：想定売上規模（百万）・受注月・アカウント責任者');
  assert.deepStrictEqual(s.changes.map(label), ['C社／店舗什器', 'A社／アプリ刷新', 'D社／物流改善'], 'サンプルの記録は数えない');
});

test('ステータスの分け方・受注月・金額の読み方', () => {
  const {g} = setup({edits: false});
  assert.deepStrictEqual(['受注', '受注済み', '受注見込み', '受注前', '失注', '見送り', '提案中', ''].map(g.mtgStatusKind_),
    ['won', 'won', 'active', 'active', 'lost', 'lost', 'active', 'active']);
  assert.deepStrictEqual(plain(['2026/9', '2026年10月', '２０２６－１１', '10月', ''].map(g.mtgMonth_)),
    [{y: 2026, m: 9}, {y: 2026, m: 10}, {y: 2026, m: 11}, null, null]);
  assert.deepStrictEqual(plain(g.mtgMonth_(new Date(Date.UTC(2026, 8, 30, 20)))), {y: 2026, m: 10}, '日付の値は日本時間の月');
  assert.deepStrictEqual(['1,234.5', '12百万', '', 'TBD', 7].map(g.mtgNumber_), [1234.5, 12, null, null, 7]);
  assert.strictEqual(g.mtgAmount_(1234.56), '1,234.6');
  assert.strictEqual(g.mtgFit_('あいうえおかきくけこさしすせそ', 100, 9), 'あいうえおか…', '「…」も1文字ぶんとして収める');
  assert.strictEqual(g.mtgFit_('あいうえお', 100, 9), 'あいうえお');
});

test('記録は下から読み、前回より前の記録に届いたら止める（サンプルの過去の記録は飛ばす）', () => {
  const gas = createGas();
  const g = gas.global;
  const log = gas.addSheet('変更履歴_差分', [['記録日時', 'イベントID', 'シート']], {rows: 20, columns: 3});
  const day = n => new gas.Date(Date.now() - n * 86400000);
  const rows = [[day(30), 'EVT-1', 'x'], [day(20), 'EVT-2', 'x'], [day(10), 'EVT-3', 'x'], [day(2), 'EVT-4', 'x'],
    [day(1), 'EVT-5', 'x'], [day(40), 'SMP-1', 'x'], [day(40), 'SMP-2', 'x']];
  gas.asUser(() => log.getRange(2, 1, rows.length, 3).setValues(rows));
  gas.get('MTG_OPTIONS').logChunkRows = 2;
  gas.reads.length = 0;
  const read = g.mtgReadRecent_(log, day(14));
  assert.deepStrictEqual(plain(read.rows.map(r => r[1])), ['EVT-3', 'EVT-4', 'EVT-5'], '14日以内の記録を古い順に');
  assert.deepStrictEqual(gas.reads.filter(r => r.kind === 'getValues').map(r => r.a1), ['A7:C8', 'A5:C6', 'A3:C4'],
    '20日前の記録（3行目）に届いたところで止める（2行目は読まない）');
});

/* ---------------- 1枚の表と PDF ---------------- */

test('1枚の表：見出し・数字のタイル・4つの段を決まった行に置き、多いときは最後の行を「ほか N件」にする', () => {
  const {gas, g} = setup();
  const grid = gas.get('MTG_GRID');
  const s = summarize(g, '新FMT');
  const layout = g.mtgLayout_(s);
  const v = (r, c) => layout.values[r - 1][c - 1];
  assert.strictEqual(v(1, 1), 'RXビジネスMTG サマリ｜新FMT');
  assert.match(v(2, 1), /^作成 \d{4}\/\d{2}\/\d{2} \d{2}:\d{2}　・　初回のため、直近14日（\d{4}\/\d{2}\/\d{2}〜）の変化$/);
  assert.deepStrictEqual([v(4, 1), v(5, 1), v(5, 8), v(5, 11)], ['進行中の案件', '3 件', '3 件', '2 件']);
  assert.deepStrictEqual([v(7, 1), v(7, 8), v(23, 1), v(23, 8)], ['① トータルサマリー', '② 前回からの進捗・変化', '③ 停滞・失注', '④ アラート']);
  assert.deepStrictEqual([v(9, 8), v(9, 9)], ['C社／店舗什器', '提案中 → 失注']);
  assert.deepStrictEqual([v(25, 1), v(25, 3), v(25, 6)], ['E社／販促DX', '40日', '決裁者に会えていない']);
  assert.strictEqual(layout.styles[8][8].fc, gas.get('MTG_OPTIONS').colors.danger, '失注への動きは赤');
  assert.strictEqual(layout.values.length, grid.height);

  // 変化が多いときは、13行目を「ほか N件」にする
  const many = Object.assign({}, s, {changes: Array.from({length: 20}, (_, i) => Object.assign({}, s.changes[2], {project: '案件' + i}))});
  const big = g.mtgLayout_(many);
  assert.strictEqual(big.values[9 + 12 - 1][7], 'ほか 8 件（変更履歴で確認できます）');
  assert.ok(big.values.every(line => line.length === grid.widths.length));
});

test('createMeetingSummaries：シートごとに A4 横・1ページの PDF をフォルダに保存し、一時的なスプレッドシートはゴミ箱へ。次は今以降の変化を出す', () => {
  const {gas, g} = setup();
  gas.promptAnswer = '';   // 空欄 → すべてのシート
  g.createMeetingSummaries();

  assert.match(gas.prompts[0].message, /1\. 新FMT\n2\. 新FMT2/);
  assert.deepStrictEqual(gas.driveFiles.map(f => f.name.replace(/\d{4}-\d{2}-\d{2}/, 'DATE')),
    ['会議サマリ_新FMT_DATE.pdf', '会議サマリ_新FMT2_DATE.pdf']);
  assert.deepStrictEqual(plain(gas.driveFolders), [{id: 'folder-1', name: '会議サマリ', parent: 'parent-of-spreadsheet-id'}],
    'スプレッドシートと同じフォルダの中に作る');
  assert.ok(gas.driveFiles.every(f => f.folder === 'folder-1'));
  const temp = gas.created[0];
  assert.deepStrictEqual(gas.trashed, [temp.getId()], '一時的なスプレッドシートはゴミ箱へ');
  assert.strictEqual(gas.fetches.length, 2);
  const url = gas.fetches[0].url;
  assert.ok(url.startsWith('https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?'));
  ['format=pdf', 'size=A4', 'portrait=false', 'scale=4', 'gridlines=false', 'gid=' + temp.getSheetByName('新FMT').getSheetId(),
    'r2=39', 'c2=13'].forEach(p => assert.ok(url.indexOf(p) >= 0, p));
  assert.strictEqual(gas.fetches[0].params.headers.Authorization, 'Bearer oauth-token');
  assert.deepStrictEqual(gas.writes.filter(w => w.ss === gas.ss.getId()), [], 'スプレッドシートには書き込まない');
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /「新FMT」：会議サマリ_新FMT_.*\.pdf（動いた 3・停滞 1・アラート 2）/);
  assert.match(done, /保存先：https:\/\/drive\.google\.com\/drive\/folders\/folder-1/);

  // 2回目：前回（今）以降の変化だけ。フォルダは同じものを使う
  const sheet = gas.ss.getSheetByName('新FMT');
  g.recordDiffEdit(gas.edit(sheet, 'F7', '受注見込み'));   // E社
  gas.promptAnswer = '1';
  g.createMeetingSummaries();
  assert.strictEqual(gas.driveFolders.length, 1);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：.*（動いた 1・停滞 0・アラート 2）/);
  assert.strictEqual(gas.driveFiles.length, 3, '選んだシートだけ作る');
});

test('previewMeetingSummaries：試作は「前回から」の基準を進めない（ファイル名に「試作」）。番号の間違い・やめたときは作らない', () => {
  const {gas, g} = setup();
  gas.promptAnswer = '1';
  g.previewMeetingSummaries();
  g.previewMeetingSummaries();
  assert.ok(gas.driveFiles.every(f => /_試作\.pdf$/.test(f.name)));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（動いた 3・停滞 1・アラート 2）[\s\S]*試作のため/, '2回目も同じ変化を出す');
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_LAST_新FMT'), null);

  gas.promptAnswer = '3';
  assert.throws(() => g.previewMeetingSummaries(), /「3」に当たるシートがありません。1〜2 の番号で入力してください。/);
  gas.promptButton = 'CANCEL';
  const before = gas.driveFiles.length;
  g.createMeetingSummaries();
  assert.strictEqual(gas.driveFiles.length, before);
});

test('PDF に書き出せなかったときは止め、一時的なスプレッドシートはゴミ箱に移す（基準も進めない）', () => {
  const {gas, g} = setup();
  gas.fetchStatus = 403;
  gas.promptAnswer = '1';
  assert.throws(() => g.createMeetingSummaries(), /PDF に書き出せませんでした（403）/);
  assert.strictEqual(gas.trashed.length, 1);
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_LAST_新FMT'), null);
});
