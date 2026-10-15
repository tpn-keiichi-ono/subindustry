'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/*
 * 会議サマリ（MeetingSummary.gs）
 * 追跡シート（2行目が見出し。追跡範囲は D3:AF）：
 *   D 得意先 / E 案件名 / F 案件ステータス / G 想定売上規模（百万） / H 期待値調整済（改行）想定売上規模（百万） / I 受注月 /
 *   J 活動状況 / K 活動における課題 / L アカウント責任者 / M 売上開始月 / N デリバリー担当 /
 *   AG 最終更新日時 / AH クレデンシャル（改行）オファリング登録 / AI 変更履歴
 */
const WIDTH = 35;
const COL = {customer: 4, project: 5, status: 6, amount: 7, weighted: 8, orderMonth: 9, activity: 10, issue: 11, owner: 12,
  startMonth: 13, delivery: 14, stamp: 33, credButton: 34, changeButton: 35};
const HEADERS = {customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', orderMonth: '受注月', activity: '活動状況', issue: '活動における課題',
  owner: 'アカウント責任者', startMonth: '売上開始月', delivery: 'デリバリー担当', stamp: '最終更新日時',
  credButton: 'クレデンシャル\nオファリング登録', changeButton: '変更履歴'};
const STATUSES = ['初回訪問', '提案中', '受注見込み', '受注', '失注', '保留'];

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
  const gas = createGas(o.gas);
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
    g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));                 // A社：初回訪問 → 提案中（前進）
    g.recordDiffEdit(gas.edit(sheet, 'J3', '10/10 提案書を提出'));       // A社：活動状況
    g.recordDiffEdit(gas.edit(sheet, 'F5', '失注'));                   // C社：提案中 → 失注
    g.recordDiffEdit(gas.edit(sheet, 'J6', '10/12 打ち合わせ'));         // D社：活動状況（更新のみ）
    gas.asUser(() => sheet.deleteRow(8));                             // F社の行を削除
    g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));
  }
  // E社は40日前から更新なし（最終更新日時を手で古くする。追跡範囲の外の列）
  gas.asUser(() => sheet.getRange(7, COL.stamp).setValue(new gas.Date(Date.now() - 40 * 86400000)));
  gas.writes.length = 0;
  gas.toasts.length = 0;
  return {gas, g, sheet};
}

/** fn の中で足された変更の記録（変更履歴_差分・変更時点スナップショット）の日時を days 日前にする */
function aged(gas, days, fn) {
  const sheets = ['変更履歴_差分', '変更時点スナップショット'].map(name => gas.ss.getSheetByName(name));
  const before = sheets.map(s => (s ? s.getLastRow() : 1));
  fn();
  sheets.forEach((s, i) => {
    const sheet = s || gas.ss.getSheetByName(['変更履歴_差分', '変更時点スナップショット'][i]);
    if (!sheet) return;
    for (let r = Math.max(before[i], 1) + 1; r <= sheet.getLastRow(); r++) {
      gas.asUser(() => sheet.getRange(r, 1).setValue(new gas.Date(Date.now() - days * 86400000)));
    }
  });
}

/**
 * 1つのシートの中身（スライドにする前の形）を作る（直近14日の変化。初回と同じ）。
 * options.trackedSince：変更の記録を始めた時刻（既定は変更履歴_差分 の最初の行）。options.draft：試作の入力内容のチェックの結果
 */
function summarize(g, name, options) {
  const o = options || {};
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const since = new Date(Date.now() - 14 * 86400000);
  const windowStart = new Date(Date.now() - 180 * 86400000);
  return g.mtgSummarize_(ss, name, {
    since, previous: null, legacy: null, draft: o.draft || null, final: true, now, meeting: '2026/10/23', windowStart,
    trackedSince: 'trackedSince' in o ? o.trackedSince : g.mtgLogStart_(ss),
    log: g.mtgReadLog_(ss, windowStart),
    snapshots: g.mtgReadSnapshots_(ss, windowStart),
    credentials: g.hsCredentialIndex_(ss)
  });
}

const label = item => [item.customer, item.project].filter(Boolean).join('／');

/* ---------------- 読み方 ---------------- */

test('ステータスの分類（進行中・受注・失注・保留。statusKinds で個別に）・受注月・金額・会議日の読み方', () => {
  const {gas, g} = setup({edits: false});
  assert.deepStrictEqual(['受注', '受注済み', '受注見込み', '受注前', '失注', '見送り', '保留', 'ペンディング中', '提案中', ''].map(g.mtgStatusKind_),
    ['won', 'won', 'active', 'active', 'lost', 'lost', 'hold', 'hold', 'active', 'active']);
  gas.get('MTG_OPTIONS').statusKinds = {'検討中（休止）': 'hold', '提案中': 'lost'};
  assert.deepStrictEqual(['検討中（休止）', '提案中'].map(g.mtgStatusKind_), ['hold', 'lost'], 'statusKinds が先');
  gas.get('MTG_OPTIONS').statusKinds = {};

  assert.deepStrictEqual(plain(['2026/9', '2026年10月', '２０２６－１１', '10月', ''].map(g.mtgMonth_)),
    [{y: 2026, m: 9}, {y: 2026, m: 10}, {y: 2026, m: 11}, null, null]);
  assert.deepStrictEqual(plain(g.mtgMonth_(new Date(Date.UTC(2026, 8, 30, 20)))), {y: 2026, m: 10}, '日付の値は日本時間の月');
  assert.deepStrictEqual(['2026年11月', '2026/11/01', '未定', ''].map(g.mtgMonthLabel_), ['2026/11', '2026/11', '未定', '—']);
  assert.deepStrictEqual(['1,234.5', '12百万', '', 'TBD', 7].map(g.mtgNumber_), [1234.5, 12, null, null, 7]);
  assert.strictEqual(g.mtgAmount_(1234.56), '1,234.6');
  assert.deepStrictEqual([12.34, -3, 0, 0.04].map(g.mtgSigned_), ['+12.3', '▲3', '±0', '±0']);

  const today = new Date(Date.UTC(2026, 9, 14, 3));   // 日本時間 2026/10/14 12:00
  assert.deepStrictEqual(['10/23', '10月23日', '2026/9/30', '２０２６年１０月２３日', '1/15', '4/20', '2/30', '13/1', 'abc', '']
    .map(t => g.mtgParseDate_(t, today)),
  ['2026/10/23', '2026/10/23', '2026/09/30', '2026/10/23', '2027/01/15', '2026/04/20', '', '', '', ''],
  '年が無ければ今年（半年より前になるなら来年）');
  assert.strictEqual(g.mtgShortDate_('2026/10/09'), '10/9');
});

test('記録は下から読み、読む期間より前の記録に届いたら止める（サンプルの過去の記録は飛ばす）', () => {
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

test('削除した行の「見出し: 値」は、シートの見出しで区切って読む（見出し・値に改行があっても）', () => {
  const {g} = setup({edits: false});
  const headers = Object.values(HEADERS);
  const text = '得意先: F社\n案件名: ブランド刷新\n期待値調整済\n想定売上規模（百万）: 3\n想定売上規模（百万）: 10\n活動状況: 1行目\n2行目\nアカウント責任者: 高橋';
  assert.deepStrictEqual(plain(g.mtgParseRemoved_(text, headers)), {customer: 'F社', project: 'ブランド刷新', status: '', amount: 10,
    weighted: 3, orderMonth: null, orderMonthText: '', owner: '高橋'});
});

/* ---------------- 中身 ---------------- */

test('前回からの変化：案件ごとに1つの区分（失注・延期・前進・新規・削除・更新のみ）。区分の順', () => {
  const {gas, g, sheet} = setup();
  g.recordDiffEdit(gas.edit(sheet, 'I7', month(3).text));                          // E社：受注月を延ばす
  g.recordDiffEdit(gas.edit(sheet, 'D8:F8', [['G社', 'POS連携', '初回訪問']]));      // G社：空の行に新しい案件
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.category]), [
    ['C社／店舗什器', 'lost'],
    ['E社／販促DX', 'slipped'],
    ['A社／アプリ刷新', 'forward'],
    ['G社／POS連携', 'new'],
    ['F社／ブランド刷新', 'removed'],
    ['D社／物流改善', 'updated']
  ]);
  const a = s.changes[2];
  assert.deepStrictEqual([a.statusFrom, a.statusTo, a.activity, a.columns.join('・')], ['初回訪問', '提案中', '10/10 提案書を提出', '案件ステータス・活動状況']);
  const e = s.changes[1];
  assert.deepStrictEqual([e.monthFrom, e.monthTo, e.weightedFrom, e.weightedTo], [month(1).text, month(3).text, 40, 40]);
  assert.deepStrictEqual([s.changes[4].removed, s.changes[4].statusFrom, s.changes[4].weightedFrom], [true, '提案中', 3], '削除した行は失注に数えない');
  const k = s.kpi;
  assert.deepStrictEqual([k.changed, k.updated, k.lost, k.forward, k.new, k.removed, k.slips, k.slipped], [5, 1, 1, 1, 1, 1, 1, 1]);
});

test('前回からの変化：得意先・案件名を書き換えた案件・あとから案件名を入れた案件は、前の名前の記録をつなげる', () => {
  const {gas, g, sheet} = setup({edits: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));                 // A社：前進
  g.recordDiffEdit(gas.edit(sheet, 'E3', 'アプリ刷新（第2期）'));     // 案件名を書き換える
  g.recordDiffEdit(gas.edit(sheet, 'D9', 'H社'));                    // 得意先だけ先に入れて
  g.recordDiffEdit(gas.edit(sheet, 'E9', '店舗分析'));               // あとから案件名
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.category, c.columns.join('・')]), [
    ['A社／アプリ刷新（第2期）', 'forward', '案件ステータス・案件名'],
    ['H社／店舗分析', 'new', '得意先・案件名']
  ]);
  assert.strictEqual(s.unmatched, 0);
});

test('期待値の増減：前回の時点で進行中だった期待値 + 内訳（増額・減額・受注・保留・復帰）= 今の期待値', () => {
  const {gas, g, sheet} = setup({edits: false});
  [['F3', '提案中'], ['H3', '15'],          // A社：前進。期待値 10 → 15（増額）
    ['F5', '初回訪問'], ['H5', '2'],         // C社：後退。期待値 5 → 2（減額）
    ['F7', '受注'],                         // E社：受注（期待値 40 が進行中から外れる）
    ['F4', '提案中'],                       // B社：受注 → 提案中（復帰。期待値 50 が入る）
    ['F8', '保留']                          // F社：保留（期待値 3 が外れる）
  ].forEach(([a1, value]) => g.recordDiffEdit(gas.edit(sheet, a1, value)));
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.category]), [
    ['E社／販促DX', 'won'], ['C社／店舗什器', 'backward'], ['A社／アプリ刷新', 'forward'], ['B社／EC立ち上げ', 'revived'], ['F社／ブランド刷新', 'hold']
  ]);
  const b = s.bridge;
  assert.deepStrictEqual([b.start, b.end, b.startCount, b.endCount], [58, 67, 5, 4], '前回：A10・C5・D空・E40・F3／今：A15・B50・C2・D空');
  assert.deepStrictEqual(b.steps.filter(st => st.count).map(st => [st.label, st.value, st.count]),
    [['増額', 5, 1], ['復帰', 50, 1], ['減額', -3, 1], ['受注', -40, 1], ['保留', -3, 1]]);
  assert.strictEqual(b.start + b.steps.reduce((sum, st) => sum + st.value, 0), b.end);
  assert.deepStrictEqual([s.kpi.won, s.kpi.wonAmount, s.kpi.backward, s.kpi.hold, s.kpi.revived, s.kpi.active, s.kpi.weighted], [1, 80, 1, 1, 1, 4, 67]);
});

test('停滞：ステータスが最後に変わってからの日数（変わった記録が無ければ「N日以上」）。期待値の大きい案件と、当てはまる状況（停滞・受注月超過・延期）', () => {
  const {gas, g, sheet} = setup({edits: false});
  aged(gas, 50, () => g.recordDiffEdit(gas.edit(sheet, 'F5', '受注見込み')));   // C社：50日前にステータスが変わった
  g.recordDiffEdit(gas.edit(sheet, 'F8', '受注見込み'));                       // F社：今日変わった
  const started = new Date(Date.now() - 100 * 86400000).getTime();
  const s = plain(summarize(g, '新FMT', {trackedSince: started}));
  assert.deepStrictEqual(s.stalled.map(d => [label(d), d.days, d.atLeast]), [
    ['E社／販促DX', 100, true],
    ['A社／アプリ刷新', 100, true],
    ['C社／店舗什器', 50, false],
    ['D社／物流改善', 100, true]
  ], '想定売上の大きい順（D社は想定売上が空）');

  // 期待値の大きい進行中の案件（5件まで）。判断（決めること・支援の要否など）は持たず、当てはまる状況だけ
  assert.deepStrictEqual(s.focus.map(d => [label(d), d.weighted, g.mtgFlags_(d).join('・')]), [
    ['E社／販促DX', 40, '停滞'],
    ['A社／アプリ刷新', 10, '停滞・受注月超過'],
    ['C社／店舗什器', 5, '停滞'],
    ['F社／ブランド刷新', 3, ''],
    ['D社／物流改善', null, '停滞']
  ]);
  assert.deepStrictEqual([s.focus[3].days, s.focus[3].atLeast], [0, false], '今日ステータスが変わった');
  assert.ok(s.focus.every(d => !('decide' in d) && !('star' in d) && !('support' in d)));
  assert.deepStrictEqual(s.overdue.map(label), ['A社／アプリ刷新']);
  assert.strictEqual(s.kpi.overdue, 1);
  gas.get('MTG_OPTIONS').mainRows.focus = 2;
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: started})).focus.map(label), ['E社／販促DX', 'A社／アプリ刷新']);
  gas.get('MTG_OPTIONS').mainRows.focus = 5;

  // 記録を始めて10日：最終更新日時が40日前の E社 も、45日に届かない
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: Date.now() - 10 * 86400000})).stalled.map(label), ['C社／店舗什器']);
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: null})).stalled.map(label), ['C社／店舗什器'], '記録が無ければ最終更新日時だけで見る');
});

test('入力内容のチェック：予定日超過・受注月超過（延期の回数）・未入力・更新なし・金額の不整合・月の不整合・選択肢外のステータス・デリバリー担当なし', () => {
  const {gas, g, sheet} = setup({edits: false});
  [['I3', month(-2).text], ['I3', month(-1).text],   // A社：受注月を前倒し → 延期（1回）。先月のまま進行中
    ['H3', '40'],                                     // A社：期待値 40 > 想定 30
    ['M3', month(-3).text],                           // A社：売上開始月が受注月より前
    ['F6', '検討中'],                                  // D社：プルダウンにないステータス
    ['F7', '受注']                                     // E社：受注したがデリバリー担当が空
  ].forEach(([a1, value]) => g.recordDiffEdit(gas.edit(sheet, a1, value)));
  gas.asUser(() => sheet.getRange(5, COL.stamp).setValue(new gas.Date(Date.now() - 35 * 86400000)));   // C社：35日前から更新なし
  gas.select(sheet, 'D3');
  g.addSampleHistory();   // A社に、3日前が予定日で未実施のオファリングなどを足す（変更の記録は SMP-。数えない）
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.alerts.map(a => [a.type, label(a)]), [
    ['予定日超過', 'A社'],
    ['受注月超過', 'A社／アプリ刷新'],
    ['未入力', 'F社／ブランド刷新'],
    ['未入力', 'D社／物流改善'],
    ['更新なし', 'C社／店舗什器'],
    ['金額の不整合', 'A社／アプリ刷新'],
    ['月の不整合', 'A社／アプリ刷新'],
    ['選択肢外のステータス', 'D社／物流改善'],
    ['デリバリー担当なし', 'E社／販促DX']
  ]);
  const text = (type, customer) => s.alerts.find(a => a.type === type && (!customer || a.customer === customer)).text;
  assert.match(text('予定日超過'), /^オファリング「【サンプル】販促DXの試験導入のご提案」の予定日（\d{2}\/\d{2}）から 3日、実施日が空$/);
  assert.strictEqual(s.alerts[0].owner, 'サンプル 担当', '責任者の列には当社の担当者');
  assert.strictEqual(text('受注月超過'), '受注月 ' + month(-1).y + '/' + month(-1).m + ' を過ぎても「初回訪問」（延期 1回）');
  assert.strictEqual(text('未入力', 'D社'), '空：想定売上規模（百万）・期待値調整済想定売上規模（百万）・受注月・活動状況・アカウント責任者');
  assert.strictEqual(text('未入力', 'F社'), '空：活動状況', '想定の大きい順');
  assert.strictEqual(text('更新なし'), '最終更新から 35日');
  assert.strictEqual(text('金額の不整合'), '期待値 40 が想定売上 30 より大きい');
  assert.strictEqual(text('月の不整合'), '売上開始月 ' + month(-3).y + '/' + month(-3).m + ' が受注月 ' + month(-1).y + '/' + month(-1).m + ' より前');
  assert.strictEqual(text('選択肢外のステータス'), '「検討中」はプルダウンの選択肢にない（進行中として数える）');
  assert.strictEqual(text('デリバリー担当なし'), '受注したが「デリバリー担当」が空');
  assert.ok(s.changes.every(c => !/^【サンプル】/.test(c.project)), 'サンプルの記録は変化に入れない');
  assert.strictEqual(s.kpi.owners, 6, 'サンプル 担当・佐藤・高橋・空・田中・伊藤');
});

test('今期の受注と着地見込み・受注月（四半期）別の期待値・ステータス別（プルダウンの順。進行中 → 保留 → 受注 → 失注）', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').targets['新FMT'] = 200;
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.stages.map(p => [p.status, p.kind, p.count, p.amount, p.weighted]), [
    ['提案中', 'active', 2, 110, 50],
    ['（未設定）', 'active', 1, 0, 0],
    ['受注', 'won', 1, 50, 50],
    ['失注', 'lost', 1, 20, 5]
  ]);
  const inYear = m => {
    const now = month(0);
    const start = (now.m >= 4 ? now.y : now.y - 1) * 12 + 4;
    return m.y * 12 + m.m >= start && m.y * 12 + m.m < start + 12;
  };
  const won = inYear(month(-2)) ? 50 : 0;
  const planned = [[month(-1), 10], [month(1), 40]].filter(p => inYear(p[0])).reduce((sum, p) => sum + p[1], 0);
  assert.deepStrictEqual([s.fiscal.wonAmount, s.fiscal.planned, s.fiscal.landing, s.fiscal.target], [won, planned, won + planned, 200]);
  assert.strictEqual(s.quarters.reduce((sum, q) => sum + q.count, 0), 3, '進行中の3件をどこかに入れる');
  assert.ok(s.quarters.some(q => q.label === '受注月なし' && q.count === 1));
  assert.deepStrictEqual(s.quarters.filter(q => /Q/.test(q.label)).map(q => q.label), ['1Q（4〜6月）', '2Q（7〜9月）', '3Q（10〜12月）', '4Q（1〜3月）']);
});

/* ---------------- スライド ---------------- */

/** 部品の文字（表はセルをつないだもの） */
const textOf = slide => slide.elements.map(el => (el.type === 'table'
  ? el.rows.map(r => r.cells.map(c => c.text).join(' | ')).join('\n') : String(el.text || ''))).join('\n');

test('スライド：本編4枚（サマリー → パイプラインの変化 → 前回から変わった案件 → 期待値の大きい案件の状況・課題）、扉、Appendix A〜E。見出しはどのシートでも同じ、状態はリード文に書く', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const T = gas.get('MTG_TITLES');
  const c = gas.get('MTG_OPTIONS').colors;
  const s = summarize(g, '新FMT', {trackedSince: Date.now() - 100 * 86400000});
  const slides = plain(g.mtgSlides_(s));
  const titles = slides.map(sl => sl.title);
  assert.deepStrictEqual(titles.slice(0, 11),
    [T.summary, T.pipeline, T.progress, T.focus, T.appendix, T.changes, T.stalled, T.owners, T.alerts, T.status, T.glossary]);
  assert.deepStrictEqual(titles.slice(0, 11), ['サマリー', 'パイプラインの変化', '前回から変わった案件', '期待値の大きい案件の状況・課題', 'Appendix',
    'Appendix A　前回から変わった案件（一覧）', 'Appendix B　停滞している案件（一覧）', 'Appendix C　入力内容のチェック（責任者別）',
    'Appendix C　入力内容のチェック（一覧）', 'Appendix D　ステータスの分類', 'Appendix E　用語の定義']);
  assert.ok(titles.slice(11).every(t => t === 'Appendix E　用語の定義（続き）'));

  // 扉：大きな「Appendix」と、Appendix の見出し・始まりのページ（「（続き）」のページは出さない）
  const divider = slides[4];
  assert.ok(!divider.elements.some(el => el.y === 26), '扉には、ほかのスライドの見出しの位置の見出しを置かない');
  assert.ok(divider.elements.some(el => el.text === 'Appendix' && el.size >= 30));
  const contents = divider.elements.find(el => el.type === 'table').rows.map(r => r.cells.map(cell => cell.text));
  assert.deepStrictEqual(contents, titles.map((t, i) => [t, i + 1]).filter(([t]) => /^Appendix \S/.test(t) && !/（続き）$/.test(t))
    .map(([t, page]) => [t.split('　')[0], t.split('　')[1], String(page)]));
  assert.deepStrictEqual(contents[0], ['Appendix A', '前回から変わった案件（一覧）', '6']);

  // 左上のラベルは置かない。どのサブインダストリー・会議日の資料かは、いちばん下の出典の行に書く
  slides.forEach((slide, i) => {
    assert.ok(!slide.elements.some(el => el.x === 0 && el.y === 0), (i + 1) + '枚目：左上のラベル');
    assert.match(slide.elements.find(el => el.y === 383).text, /^新FMT（10\/23 会議）　データ：シートと変更履歴_差分（\d{4}\/\d{2}\/\d{2} \d{2}:\d{2} 時点）　期間：/);
  });
  assert.match(plain(g.mtgSlides_(Object.assign({}, s, {final: false})))[0].elements.find(el => el.y === 383).text, /^【試作】新FMT（10\/23 会議）/);
  const other = plain(g.mtgSlides_(summarize(g, '新FMT2', {trackedSince: Date.now() - 100 * 86400000}))).map(sl => sl.title);
  assert.deepStrictEqual(other.slice(0, 4), titles.slice(0, 4), '本編の見出しは、中身が違うシートでも同じ');

  // 事前資料には事実だけを書く（論点・決めてほしいこと・支援の要否などの判断は、会議でリーダーが話す）
  const all = slides.map(textOf).join('\n');
  ['決めて', '論点', '★', '支援', '候補', '確認したい', '要対応', '悪い', '打開', '重点', '付録'].forEach(word => assert.ok(all.indexOf(word) < 0, '使わない言葉：' + word));
  slides.forEach(slide => slide.elements.filter(el => el.type === 'table').forEach(t => assert.ok(!(t.borders || []).length, '赤枠で強調しない')));

  // リード文（タイトルの下）に状態を書く
  const lead = slide => slide.elements.find(el => el.y === 62).text.replace(/\n/g, '');
  assert.strictEqual(lead(slides[0]), '期待値は 50百万円（直近14日で ▲8）、進行中 3件。直近14日で受注 0件・失注 1件・前進 1件・後退 0件・新規 0件。' +
    'ステータスが45日以上同じ案件は 2件（想定売上 80百万円）、受注月を過ぎた案件は 1件。');
  assert.match(lead(slides[1]), /^期待値は 58 → 50百万円（▲8）。増えた分 ±0・減った分 ▲8、内訳で最も大きいのは失注（1件・▲5）。今期に受注予定の期待値は /);
  assert.strictEqual(lead(slides[2]), '直近14日でステータス・期待値・受注月が変わった案件は 3件（受注 0・失注 1・前進 1・後退 0・延期 0・新規 0）。' +
    '全件と、ほかの項目だけ更新した 1件は Appendix A。');
  assert.strictEqual(lead(slides[3]), '進行中の案件のうち、期待値の大きい 3件（進行中の期待値の 100%）。このうち停滞 2件・受注月超過 1件・受注月の延期あり 0件。停滞の全 2件は Appendix B。');
  slides.forEach((slide, i) => {
    const text = (slide.elements.find(el => el.y === 62) || {text: ''}).text;
    assert.ok(!/…$/.test(text), (i + 1) + '枚目：リード文が切れている ' + text);
  });

  slides.forEach((slide, i) => {
    const page = slide.elements.find(el => el.text === String(i + 1));
    assert.ok(page && page.x > 600 && page.y > 370, (i + 1) + '枚目のページ番号');
    slide.elements.forEach(el => {
      const w = el.type === 'table' ? el.columns.reduce((a, b) => a + b, 0) : el.w;
      const h = el.type === 'table' ? el.rows.reduce((a, r) => a + r.h, 0) : el.h;
      assert.ok(el.x >= 0 && el.y >= 0 && el.x + w <= S.width + 0.01 && el.y + h <= S.height, (i + 1) + '枚目：はみ出す部品 ' + JSON.stringify(el).slice(0, 80));
      if (el.type === 'table') {
        assert.ok(el.y + h <= S.bottom, (i + 1) + '枚目：表が出典・ページ番号にかかる');
        assert.ok(el.columns.every(cw => cw >= 32), (i + 1) + '枚目：Slides API は 32pt より狭い列をエラーにする ' + el.columns.join(','));
        el.rows.forEach(r => r.cells.forEach((cell, col) => {
          const lines = String(cell.text).split('\n');
          assert.strictEqual(g.mtgWrap_(cell.text, el.columns[col], cell.size, 99).length, lines.length, '折り返し済み：' + cell.text);
          assert.ok(g.mtgLineHeight_(lines.length, cell.size) <= r.h + 0.01, '行の高さ：' + cell.text);
        }));
      }
    });
  });

  // 1枚目：今期の帯・4つの数字・前回からの状況の内訳（どのシートでも同じ区分。該当が無ければ0件）
  const first = textOf(slides[0]);
  assert.match(first, /今期（\d{4}\/4〜\d{4}\/3）　受注 /);
  assert.match(first, /進行中の期待値（百万円）\n50\n直近14日で ▲8・3件（想定売上 110）/);
  const breakdown = slides[0].elements.find(el => el.type === 'table').rows.map(r => r.cells.map(cell => cell.text.replace(/\n/g, '')));
  assert.deepStrictEqual(breakdown, [
    ['区分', '件数', '想定売上', '案件（想定売上の大きい順）'],
    ['新規・前進', '1件', '30', 'A社／アプリ刷新（初回訪問→提案中）'],
    ['受注', '0件', '—', '—'],
    ['失注・後退・削除', '2件', '30', 'C社／店舗什器（失注）、F社／ブランド刷新（削除）'],
    ['受注月の延期・超過', '1件', '30', 'A社／アプリ刷新（' + month(-1).y + '/' + month(-1).m + 'を過ぎて進行中）'],
    ['停滞（45日以上）', '2件', '80', 'E社／販促DX（100日以上）、D社／物流改善（100日以上）']
  ]);
  // 3枚目：失注・後退・延期の区分は赤字
  const moves = slides[2].elements.find(el => el.type === 'table');
  assert.match(moves.rows[1].cells.map(cell => cell.text).join(' | '), /^C社\n店舗什器 \| 失注 \| 提案中 ▶ 失注 \| 5 \|/);
  assert.strictEqual(moves.rows[1].cells[1].color, c.red);
  // 4枚目：期待値の大きい順。当てはまる状況は事実だけ
  const focus = slides[3].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(focus.rows[0].cells.map(cell => cell.text.replace(/\n/g, '')),
    ['得意先／案件名', '期待値／想定売上', 'ステータス（同じ日数）', '受注月', '活動状況', '活動における課題', '当てはまる状況']);
  assert.deepStrictEqual(focus.rows.slice(1).map(r => [r.cells[0].text.split('\n')[0], r.cells[6].text]), [['E社', '停滞'], ['A社', '受注月超過'], ['D社', '停滞']]);
  assert.match(focus.rows[2].cells.map(cell => cell.text).join(' | '), /^A社\nアプリ刷新 \| 10\n30 \| 提案中\n0日 \| \d{4}\/\d+ \|/);
  // Appendix E：用語の定義（設定の日数・区分の順を入れる）
  const glossary = slides.filter(sl => /^Appendix E/.test(sl.title)).map(textOf).join('\n');
  assert.match(glossary, /区分 \| 前回からの変化の分け方。案件ごとに1つで、当てはまるものが複数あれば 新規 > 受注 > 失注 > 保留 > 復帰 > 前進 > 後退 > 延期 >/);
  assert.match(glossary, /停滞 \| 進行中で、ステータスが45日以上同じ案件/);
  assert.match(glossary, /前回 \| 前回の会議の資料を作った日時（前回の会議の記録が無ければ、直近14日）。/);
});

test('スライド：Appendix は入りきらなければ続きのページに分け、見出しの行を繰り返す（全件を出す）。中身の無い Appendix は作らない', () => {
  const {g} = setup();
  const s = summarize(g, '新FMT');
  const many = Object.assign({}, s, {changes: Array.from({length: 40}, (_, i) => Object.assign({}, s.changes[1], {project: '案件' + i}))});
  const pages = plain(g.mtgSlides_(many)).filter(sl => /^Appendix A/.test(sl.title));
  assert.ok(pages.length >= 3, pages.length + 'ページ');
  assert.strictEqual(pages[1].title, 'Appendix A　前回から変わった案件（一覧）（続き）');
  const tables = pages.map(p => p.elements.find(el => el.type === 'table'));
  tables.forEach(t => assert.strictEqual(t.rows[0].cells[0].text, '得意先／案件名', '見出しの行を繰り返す'));
  assert.strictEqual(tables.reduce((sum, t) => sum + t.rows.length - 1, 0), 40, '全件を出す（省かない）');

  const empty = Object.assign({}, s, {changes: [], stalled: [], overdue: [], alerts: []});
  const slides = plain(g.mtgSlides_(empty));
  assert.deepStrictEqual(slides.map(sl => sl.title).slice(4).filter(t => !/（続き）$/.test(t)), ['Appendix', 'Appendix D　ステータスの分類', 'Appendix E　用語の定義']);
  assert.deepStrictEqual(slides[4].elements.find(el => el.type === 'table').rows.map(r => r.cells[0].text), ['Appendix D', 'Appendix E'],
    '扉には、作った Appendix だけを出す');
  const breakdown = slides[0].elements.find(el => el.type === 'table').rows.slice(1).map(r => r.cells.map(cell => cell.text));
  assert.ok(breakdown.every(r => r[1] === '0件' && r[3] === '—'), '区分は同じまま、0件で出す');
});

test('文字の折り返し：見積もりの幅で分け、最後の行は「…」。行の頭に来ない文字はぶら下げ、半角の語は途中で切らない', () => {
  const {g} = setup({edits: false});
  assert.deepStrictEqual(plain(g.mtgWrap_('あいうえおかきくけこ', 100, 9, 3)), ['あいうえおかきく', 'けこ']);
  assert.deepStrictEqual(plain(g.mtgWrap_('あいうえおかきくけこさしすせそたちつてと', 100, 9, 2)), ['あいうえおかきく', 'けこさしすせそ…'],
    '最後の行は「…」の分をあけて切る');
  assert.deepStrictEqual(plain(g.mtgWrap_('あいうえおかきく。', 100, 9, 3)), ['あいうえおかきく。'], '本当の幅に収まる「。」はぶら下げる');
  assert.deepStrictEqual(plain(g.mtgWrap_('期待値（前▶今）', 72, 7.5, 3)), ['期待値（前▶', '今）'], '収まらなければ前の文字と一緒に送る');
  assert.deepStrictEqual(plain(g.mtgWrap_('2026/11 ▶ 2026/12', 80, 7.5, 2)), ['2026/11 ▶', '2026/12']);
  assert.deepStrictEqual(plain(g.mtgWrap_('一行目\n二行目', 300, 9, 2)), ['一行目', '二行目']);
});

test('mtgSlideRequests_：スライド・箱・表を Slides API のリクエストにする（空のセルにも空白を入れる・赤枠）', () => {
  const {g} = setup({edits: false});
  const slide = {elements: [
    {type: 'box', x: 0, y: 0, w: 100, h: 22, text: 'タグ', size: 10, bold: true, color: '#FFFFFF', fill: '#7030A0', shape: 'RECTANGLE'},
    {type: 'table', x: 24, y: 98, columns: [100, 50], rows: [{h: 18, cells: [{text: '見出し'}, {text: ''}]}, {h: 20, cells: [{text: 'A'}, {text: 'B', fill: '#F2F2F2'}]}],
      borders: [{row: 1, col: 0, rows: 1, cols: 2, color: '#E60012'}]}
  ]};
  const requests = plain(g.mtgSlideRequests_(slide, 'mtg_p1', 0));
  assert.deepStrictEqual(requests[0], {createSlide: {objectId: 'mtg_p1', insertionIndex: 0, slideLayoutReference: {predefinedLayout: 'BLANK'}}});
  const shape = requests.find(r => r.createShape).createShape;
  assert.deepStrictEqual([shape.objectId, shape.shapeType, shape.elementProperties.size.width.magnitude], ['mtg_p1_0', 'RECTANGLE', 100]);
  const table = requests.find(r => r.createTable).createTable;
  assert.deepStrictEqual([table.objectId, table.rows, table.columns, table.elementProperties.size.height.magnitude], ['mtg_p1_1', 2, 2, 38]);
  const texts = requests.filter(r => r.insertText).map(r => r.insertText.text);
  assert.deepStrictEqual(texts, ['タグ', '見出し', ' ', 'A', 'B'], '空のセルには空白（既定の大きな文字で行が伸びないように）');
  const style = requests.find(r => r.updateTextStyle && r.updateTextStyle.objectId === 'mtg_p1_0').updateTextStyle.style;
  assert.deepStrictEqual([style.fontFamily, style.fontSize.magnitude, style.bold], ['Noto Sans JP', 10, true]);
  const borders = requests.filter(r => r.updateTableBorderProperties).map(r => r.updateTableBorderProperties);
  assert.deepStrictEqual(borders.map(b => [b.borderPosition, b.tableRange.location.rowIndex, b.tableBorderProperties.weight.magnitude]),
    [['ALL', 0, 1.5], ['OUTER', 1, 2.25]]);
});

/* ---------------- 作る・保存する ---------------- */

test('createMeetingSummaries：シートと会議日を聞き、シートごとに Google スライド（16:9）と PDF を保存する。会議日を記録し、次の会議は今以降の変化を出す', () => {
  const {gas, g, sheet} = setup();
  const meeting = g.mtgParseDate_('10/23', new Date());
  const file = meeting.replace(/\//g, '-');
  gas.get('MTG_OPTIONS').batchRequests = 500;   // まとめて送る数を小さくして、分けて送るところも確かめる
  gas.promptAnswers = ['', '10/23'];   // 空欄 → すべてのシート
  g.createMeetingSummaries();

  assert.match(gas.prompts[0].message, /1\. 新FMT\n2\. 新FMT2/);
  assert.match(gas.prompts[1].message, /今回の会議日を入力してください（例：10\/23）。[\s\S]*・新FMT：記録なし\n・新FMT2：記録なし$/);
  assert.deepStrictEqual(gas.decks.map(d => [d.title, d.pageSize.width.magnitude, d.pageSize.height.magnitude]),
    [['会議サマリ_新FMT_' + file, 720, 405], ['会議サマリ_新FMT2_' + file, 720, 405]], 'サブインダストリー（シート）ごとに1つ。まとめた資料は作らない');
  const requests = gas.decks[0].requests;
  assert.strictEqual(requests.filter(r => r.createSlide).length, 11, '本編4枚・扉・Appendix A・C（責任者別・一覧）・D・E（2ページ）（停滞が無いので B は無い）');
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[1]), /^サマリー\n/);
  assert.match(inserted(gas.decks[1]), /\n新FMT2（10\/23 会議）　データ：/);
  assert.deepStrictEqual(requests.filter(r => r.deleteObject).map(r => r.deleteObject.objectId), ['p'], '作ったときの白紙のスライドを消す');
  const batches = gas.decks[0].batches;
  assert.ok(batches.length >= 2 && batches.length < 8, '何枚かずつまとめて送る（' + batches.length + '回）');
  assert.ok(requests.findIndex(r => r.deleteObject) < batches[0], '白紙のスライドは1回目で消す');
  assert.deepStrictEqual(plain(gas.driveFolders), [{id: 'folder-1', name: '会議サマリ', parent: 'parent-of-spreadsheet-id'}],
    'スプレッドシートと同じフォルダの中に作る');
  assert.deepStrictEqual(gas.driveFiles.map(f => [f.name, f.folder, f.blob.from, f.blob.getContentType()]), [
    ['会議サマリ_新FMT_' + file + '.pdf', 'folder-1', 'deck-1', 'application/pdf'],
    ['会議サマリ_新FMT2_' + file + '.pdf', 'folder-1', 'deck-2', 'application/pdf']
  ]);
  assert.deepStrictEqual(plain(gas.moved), [{id: 'deck-1', folder: 'folder-1'}, {id: 'deck-2', folder: 'folder-1'}], 'スライドも同じフォルダに残す');
  assert.deepStrictEqual(gas.trashed, []);
  assert.deepStrictEqual(gas.writes.filter(w => w.ss === gas.ss.getId()), [], 'スプレッドシートには書き込まない');
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /「新FMT」：会議サマリ_新FMT_.*\.pdf（変わった案件 3・停滞 0・チェック 3）/);
  assert.doesNotMatch(done, /全体/);
  assert.match(done, /保存先：https:\/\/drive\.google\.com\/drive\/folders\/folder-1/);
  assert.match(done, /会議日（\d{4}\/10\/23）を記録しました。/);
  const props = g.PropertiesService.getDocumentProperties();
  const saved = JSON.parse(props.getProperty('MTG_MEETINGS_新FMT'));
  assert.deepStrictEqual(saved.map(m => m.date), [meeting]);
  assert.ok(Math.abs(saved[0].cutoff - Date.now()) < 60000);

  // 次の会議：前回（10/23 の資料を作った時）以降の変化だけ。フォルダは同じものを使う
  g.recordDiffEdit(gas.edit(sheet, 'F7', '受注見込み'));   // E社：前進
  const next = g.mtgParseDate_('10/30', new Date());
  gas.promptAnswers = ['1', '10/30'];
  g.createMeetingSummaries();
  assert.match(gas.prompts[3].message, new RegExp('・新FMT：' + meeting + '$'));
  assert.strictEqual(gas.driveFolders.length, 1);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：.*（変わった案件 1・停滞 0・チェック 2）/);
  assert.deepStrictEqual(gas.driveFiles.slice(2).map(f => f.name.replace(/_\d{4}-\d{2}-\d{2}/, '')), ['会議サマリ_新FMT.pdf'], '選んだシートだけ作る');
  // 同じ会議日で作り直しても、前回（10/23）からの変化（作った日時だけ上書き）
  gas.promptAnswers = ['1', '10/30'];
  g.createMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（変わった案件 1・/);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).map(m => m.date), [meeting, next]);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT2')).map(m => m.date), [meeting], '選ばなかったシートの会議日は記録しない');
});

test('previewMeetingSummaries：試作は会議日を記録しない（ファイル名に「試作」）。最初の試作の入力内容のチェックの結果を覚え、同じ会議日の資料で解消した件数を出す', () => {
  const {gas, g, sheet} = setup();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.ok(gas.driveFiles.every(f => /_試作\.pdf$/.test(f.name)));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（変わった案件 3・停滞 0・チェック 3）[\s\S]*試作のため、会議日は記録していません。/, '2回目も同じ変化を出す');
  const props = g.PropertiesService.getDocumentProperties();
  assert.strictEqual(props.getProperty('MTG_MEETINGS_新FMT'), null);
  const draft = JSON.parse(props.getProperty('MTG_DRAFT_新FMT'));
  assert.deepStrictEqual([draft.meeting, draft.keys.length], [g.mtgParseDate_('10/23', new Date()), 3]);
  assert.deepStrictEqual(gas.decks.map(d => d.title.replace(/_\d{4}-\d{2}-\d{2}/, '')), ['会議サマリ_新FMT_試作', '会議サマリ_新FMT_試作']);
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[1]), /入力内容のチェックに当てはまったのは 3件（責任者 3人）。\n?試作（\d{2}\/\d{2}）の時点から 0件が解消、\n?新たに 0件。/,
    '2回目の試作：まだ解消していない');

  // E社の活動状況を書いて「更新なし」を解消してから、本番を作る
  g.recordDiffEdit(gas.edit(sheet, 'J7', '10/20 決裁者と面談を設定'));
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(inserted(gas.decks[2]), /入力内容のチェックに当てはまったのは 2件（責任者 2人）。\n?試作（\d{2}\/\d{2}）の時点から 1件が解消、\n?新たに 0件。/);
});

test('会議日・シートの番号の間違い、やめたときは作らない', () => {
  const {gas, g} = setup();
  gas.promptAnswers = ['3'];
  assert.throws(() => g.previewMeetingSummaries(), /「3」に当たるシートがありません。1〜2 の番号で入力してください。/);
  gas.promptAnswers = ['1', '来週'];
  assert.throws(() => g.previewMeetingSummaries(), /「来週」を会議日として読めません。10\/23 か 2026\/10\/23 のように入力してください。/);
  gas.promptAnswers = ['1', ''];
  assert.throws(() => g.previewMeetingSummaries(), /会議日を入力してください（例：10\/23）。/);
  gas.promptButton = 'CANCEL';
  g.createMeetingSummaries();
  assert.deepStrictEqual([gas.decks.length, gas.driveFiles.length], [0, 0]);
});

test('Slides API が無い・失敗したときは止め、作りかけのスライドはゴミ箱に移す（会議日も記録しない）', () => {
  const off = setup({gas: {noSlides: true}});
  off.gas.promptAnswers = ['1', '10/23'];
  assert.throws(() => off.g.createMeetingSummaries(), /Google Slides API（拡張サービス）が有効になっていません。エディタの左の「サービス」の＋から「Google Slides API」を追加/);

  const {gas, g} = setup();
  gas.slidesError = 'Invalid requests[3]';
  gas.promptAnswers = ['1', '10/23'];
  assert.throws(() => g.createMeetingSummaries(), /Invalid requests\[3\]/);
  assert.deepStrictEqual(gas.trashed, ['deck-1']);
  assert.strictEqual(gas.driveFiles.length, 0);
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT'), null);
});
