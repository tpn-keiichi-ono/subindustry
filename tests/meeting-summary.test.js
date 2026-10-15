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
 * options.trackedSince：変更の記録を始めた時刻（既定は変更履歴_差分 の最初の行）。options.draft：試作の要対応
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

test('前回からの変化：案件ごとに1つの区分（失注・延期・前進・新規・削除・更新のみ）。大きな動きの順', () => {
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
  assert.deepStrictEqual([k.changed, k.updated, k.lost, k.forward, k.new, k.removed, k.slips, k.bad], [5, 1, 1, 1, 1, 1, 1, 2]);
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

test('停滞：ステータスが最後に変わってからの日数（変わった記録が無ければ「N日以上」）。★は当てはまる案件の想定の大きい3件まで', () => {
  const {gas, g, sheet} = setup({edits: false});
  aged(gas, 50, () => g.recordDiffEdit(gas.edit(sheet, 'F5', '受注見込み')));   // C社：50日前にステータスが変わった
  g.recordDiffEdit(gas.edit(sheet, 'F8', '受注見込み'));                       // F社：今日変わった
  const started = new Date(Date.now() - 100 * 86400000).getTime();
  const s = plain(summarize(g, '新FMT', {trackedSince: started}));
  assert.deepStrictEqual(s.stalled.map(d => [label(d), d.days, d.atLeast, d.star]), [
    ['E社／販促DX', 100, true, true],
    ['A社／アプリ刷新', 100, true, true],
    ['C社／店舗什器', 50, false, false],
    ['D社／物流改善', 100, true, true]
  ], '想定の大きい順。C社は60日未満で★の候補ではない（D社は想定が空）');
  assert.deepStrictEqual(s.stalled.map(d => d.decide), ['上席の同行の要否', '受注月の見直しか、\nクローズか', '次の打ち手の確認', '継続か、クローズか'],
    'A社は受注月（先月）を過ぎている');

  // 重点案件：★の案件を先に入れて5件まで（並びは期待値の大きい順）。支援の要否を決めたい案件（停滞・受注月超過・課題に決裁者など）に印
  assert.deepStrictEqual(s.focus.map(d => [label(d), d.weighted, d.star, d.support, d.reasons.join('・')]), [
    ['E社／販促DX', 40, true, true, '停滞・課題に決裁者'],
    ['A社／アプリ刷新', 10, true, true, '停滞・受注月超過'],
    ['C社／店舗什器', 5, false, true, '停滞'],
    ['F社／ブランド刷新', 3, false, false, ''],
    ['D社／物流改善', null, true, true, '停滞']
  ]);
  assert.deepStrictEqual([s.focus[3].days, s.focus[3].atLeast, s.focus[3].decide], [0, false, '受注に向けた\n次のアクション'], '今日ステータスが変わった');
  assert.strictEqual(s.kpi.support, 4);
  gas.get('MTG_OPTIONS').mainRows.focus = 2;
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: started})).focus.map(label), ['E社／販促DX', 'A社／アプリ刷新'], '★を先に入れる');
  gas.get('MTG_OPTIONS').mainRows.focus = 5;
  gas.get('MTG_OPTIONS').canvas.max = 1;
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: started})).stalled.filter(d => d.star).map(label), ['E社／販促DX']);

  // 記録を始めて10日：最終更新日時が40日前の E社 も、45日に届かない
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: Date.now() - 10 * 86400000})).stalled.map(label), ['C社／店舗什器']);
  assert.deepStrictEqual(plain(summarize(g, '新FMT', {trackedSince: null})).stalled.map(label), ['C社／店舗什器'], '記録が無ければ最終更新日時だけで見る');
});

test('要対応：予定超過・受注月超過（延期の回数）・未入力・更新なし・金額の矛盾・日付の矛盾・未定義ステータス・体制未登録', () => {
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
    ['予定超過', 'A社'],
    ['受注月超過', 'A社／アプリ刷新'],
    ['未入力', 'F社／ブランド刷新'],
    ['未入力', 'D社／物流改善'],
    ['更新なし', 'C社／店舗什器'],
    ['金額の矛盾', 'A社／アプリ刷新'],
    ['日付の矛盾', 'A社／アプリ刷新'],
    ['未定義ステータス', 'D社／物流改善'],
    ['体制未登録', 'E社／販促DX']
  ]);
  const text = (type, customer) => s.alerts.find(a => a.type === type && (!customer || a.customer === customer)).text;
  assert.match(text('予定超過'), /^オファリング「【サンプル】販促DXの試験導入のご提案」の予定日（\d{2}\/\d{2}）から 3日、実施日が空$/);
  assert.strictEqual(s.alerts[0].owner, 'サンプル 担当', '責任者の列には当社の担当者');
  assert.strictEqual(text('受注月超過'), '受注月 ' + month(-1).y + '/' + month(-1).m + ' を過ぎても「初回訪問」（延期 1回）');
  assert.strictEqual(text('未入力', 'D社'), '空：想定売上規模（百万）・期待値調整済想定売上規模（百万）・受注月・活動状況・アカウント責任者');
  assert.strictEqual(text('未入力', 'F社'), '空：活動状況', '想定の大きい順');
  assert.strictEqual(text('更新なし'), '最終更新から 35日');
  assert.strictEqual(text('金額の矛盾'), '期待値 40 が想定 30 より大きい');
  assert.strictEqual(text('未定義ステータス'), '「検討中」はプルダウンの選択肢にない（進行中として数える）');
  assert.strictEqual(text('体制未登録'), '受注したが「デリバリー担当」が空');
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

test('スライド：本編4枚（サマリー → パイプラインの変化 → 前回からの進捗 → 重点案件の状況・課題）と付録 A〜D。見出しはどのシートでも同じ、状態はリード文に書く', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const T = gas.get('MTG_TITLES');
  const s = summarize(g, '新FMT', {trackedSince: Date.now() - 100 * 86400000});
  const slides = plain(g.mtgSlides_(s));
  const titles = slides.map(sl => sl.title);
  assert.deepStrictEqual(titles, [T.summary, T.pipeline, T.progress, T.focus, T.changes, T.stalled, T.owners, T.alerts, T.status]);
  assert.deepStrictEqual(titles, ['サマリー', 'パイプラインの変化', '前回からの進捗', '重点案件の状況・課題', '付録A　前回からの変化（一覧）',
    '付録B　停滞している案件（一覧）', '付録C　要対応（責任者別）', '付録C　要対応（一覧）', '付録D　ステータスの分類']);
  const other = plain(g.mtgSlides_(summarize(g, '新FMT2', {trackedSince: Date.now() - 100 * 86400000}))).map(sl => sl.title);
  assert.deepStrictEqual(other.slice(0, 4), titles.slice(0, 4), '本編の見出しは、中身が違うシートでも同じ');

  // リード文（タイトルの下）に状態を書く
  const lead = slide => slide.elements.find(el => el.y === 62).text.replace(/\n/g, '');
  assert.strictEqual(lead(slides[0]), '期待値は 50百万円（直近14日で ▲8）。直近14日で受注 0件・失注 1件・前進 1件・後退 0件・新規 0件。' +
    '最優先の論点は、ステータスが45日以上動いていない案件 2件（想定 80百万円）の打開。');
  assert.match(lead(slides[1]), /^期待値は 58 → 50百万円（▲8）。増えた分 ±0・減った分 ▲8、最大の要因は失注（1件・▲5）。今期に受注予定の期待値は /);
  assert.strictEqual(lead(slides[2]), '直近14日で動いた案件は 3件（受注 0・失注 1・前進 1・後退 0・延期 0・新規 0）。悪い動き（失注・後退・延期）の 1件は赤枠。' +
    '全件と、ほかの項目だけ更新した 1件は付録A。');
  assert.match(lead(slides[3]), /^期待値の大きい進行中の案件と★の案件の 3件（進行中の期待値の 100%）。停滞の全 2件は付録B。赤枠の 3件は支援の要否を決めたい案件/);

  const band = gas.get('MTG_OPTIONS').colors.highlight;
  slides.forEach((slide, i) => {
    const page = slide.elements.find(el => el.text === String(i + 1));
    assert.ok(page && page.x > 600 && page.y > 370, (i + 1) + '枚目のページ番号');
    assert.strictEqual(slide.elements.some(el => el.fill === band), i < 4, (i + 1) + '枚目：結論の帯は本編だけ');
    slide.elements.forEach(el => {
      const w = el.type === 'table' ? el.columns.reduce((a, b) => a + b, 0) : el.w;
      const h = el.type === 'table' ? el.rows.reduce((a, r) => a + r.h, 0) : el.h;
      assert.ok(el.x >= 0 && el.y >= 0 && el.x + w <= S.width + 0.01 && el.y + h <= S.height, (i + 1) + '枚目：はみ出す部品 ' + JSON.stringify(el).slice(0, 80));
      if (el.type === 'table') {
        assert.ok(el.y + h <= (i < 4 ? S.bottom + 6 : S.appendixBottom), (i + 1) + '枚目：表が帯・下の余白にかかる');
        assert.ok(el.columns.every(cw => cw >= 32), (i + 1) + '枚目：Slides API は 32pt より狭い列をエラーにする ' + el.columns.join(','));
        el.rows.forEach(r => r.cells.forEach((cell, c) => {
          const lines = String(cell.text).split('\n');
          assert.strictEqual(g.mtgWrap_(cell.text, el.columns[c], cell.size, 99).length, lines.length, '折り返し済み：' + cell.text);
          assert.ok(g.mtgLineHeight_(lines.length, cell.size) <= r.h + 0.01, '行の高さ：' + cell.text);
        }));
      }
    });
  });

  // 1枚目：今期の帯・4つの数字・本日の論点（いちばん上は赤枠）
  const first = textOf(slides[0]);
  assert.match(first, /今期（\d{4}\/4〜\d{4}\/3）　受注 /);
  assert.match(first, /進行中の期待値（百万円）\n50\n直近14日で ▲8・3件（想定 110）/);
  assert.match(first, /本日の論点 \|  \| 状況 \|  \| 決めてほしいこと\n停滞している\n大型案件\n2件 \| ▶ \| 想定 80百万円が、ステータス45日以上そのまま/);
  const issues = slides[0].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(issues.borders.map(b => [b.row, b.color]), [[1, gas.get('MTG_OPTIONS').colors.red]]);
  // 3枚目：悪い動き（失注）は赤枠
  const moves = slides[2].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(moves.borders.map(b => b.row), [1]);
  assert.match(moves.rows[1].cells.map(c => c.text).join(' | '), /^C社\n店舗什器 \| 失注 \| 提案中 ▶ 失注 \| 5 \|/);
  // 4枚目：★の案件を先に入れた重点案件（期待値の大きい順）。支援の要否を決めたい案件は赤枠
  const focus = slides[3].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(focus.rows.slice(1).map(r => r.cells[0].text.split('\n')[0]), ['★ E社', 'A社', '★ D社']);
  assert.deepStrictEqual(focus.borders.map(b => b.row), [1, 2, 3]);
  assert.match(focus.rows[2].cells.map(c => c.text).join(' | '), /^A社\nアプリ刷新 \| 10\n30 \| 提案中\n0日 \| \d{4}\/\d+\n超過 \|/);
});

test('スライド：付録は入りきらなければ続きのページに分け、見出しの行を繰り返す（全件を出す）。中身の無い付録は作らない', () => {
  const {g} = setup();
  const s = summarize(g, '新FMT');
  const many = Object.assign({}, s, {changes: Array.from({length: 40}, (_, i) => Object.assign({}, s.changes[1], {project: '案件' + i}))});
  const pages = plain(g.mtgSlides_(many)).filter(sl => /^付録A/.test(sl.title));
  assert.ok(pages.length >= 3, pages.length + 'ページ');
  assert.strictEqual(pages[1].title, '付録A　前回からの変化（一覧）（続き）');
  const tables = pages.map(p => p.elements.find(el => el.type === 'table'));
  tables.forEach(t => assert.strictEqual(t.rows[0].cells[0].text, '得意先／案件名', '見出しの行を繰り返す'));
  assert.strictEqual(tables.reduce((sum, t) => sum + t.rows.length - 1, 0), 40, '全件を出す（省かない）');

  const empty = Object.assign({}, s, {changes: [], stalled: [], alerts: []});
  const slides = plain(g.mtgSlides_(empty));
  assert.deepStrictEqual(slides.map(sl => sl.title).slice(4), ['付録D　ステータスの分類']);
  assert.match(textOf(slides[0]), /直近14日で大きな動きは無く、各案件の次のアクションを確認したい。/);
});

test('全体サマリー：すべての追跡シートの数字と入力状況（前回から更新した案件・最終更新日時）、シートごとの本日の論点', () => {
  const {gas, g} = setup();
  const red = gas.get('MTG_OPTIONS').colors.red;
  const slides = plain(g.mtgOverallSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], {meeting: '2026/10/23', final: true, now: new Date()}));
  assert.deepStrictEqual(slides.map(sl => sl.title), ['全体サマリー', '本日の論点（サブインダストリー別）']);
  assert.strictEqual(slides[0].elements[0].text, 'RXビジネスMTG｜全体（10/23 会議）');
  const table = slides[0].elements.find(el => el.type === 'table');
  const rows = table.rows.map(r => r.cells.map(c => c.text.replace(/\n/g, '')));
  assert.deepStrictEqual(rows[0], ['サブインダストリー', '進行中', '期待値', '前回比', '今期の受注', '着地見込み', '受注／失注', '前進／後退', '停滞', '要対応',
    '更新した案件', '最終更新']);
  assert.deepStrictEqual(rows.slice(1).map(r => [r[0], r[1], r[2], r[3], r[6], r[7], r[10]]), [
    ['新FMT', '3', '50', '▲8', '0／1', '1／0', '4'],
    ['新FMT2', '1', '10', '±0', '0／0', '0／0', '0'],
    ['合計', '4', '60', '▲8', '0／1', '1／0', '4']
  ]);
  assert.strictEqual(table.rows[2].cells[10].color, red, '前回から更新の無いシートは赤');
  assert.strictEqual(table.rows[1].cells[3].color, red, '前回比がマイナスは赤');
  const lead = slides[0].elements.find(el => el.y === 62).text.replace(/\n/g, '');
  assert.match(lead, /^全体（2シート）の期待値は 60百万円（前回比 ▲8）。今期の受注 \d+百万円・着地見込み \d+百万円。受注 0件・失注 1件・停滞 0件・要対応 4件。前回から更新の無いシート：新FMT2。$/);
  assert.ok(slides[0].elements.some(el => el.fill === gas.get('MTG_OPTIONS').colors.highlight), '全体サマリーは結論の帯つき');

  const issues = slides[1].elements.find(el => el.type === 'table').rows.slice(1).map(r => r.cells.map(c => c.text.replace(/\n/g, '')));
  assert.deepStrictEqual(issues.map(r => [r[0], r[1], r[3]]), [
    ['新FMT', '受注月の遅れ 1件', '受注月の見通しと、今期の着地への影響'],
    ['新FMT2', '受注月の遅れ 1件', '受注月の見通しと、今期の着地への影響']
  ], '新FMT2 も A社 の受注月（先月）を過ぎている');
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
    [['会議サマリ_新FMT_' + file, 720, 405], ['会議サマリ_新FMT2_' + file, 720, 405], ['会議サマリ_全体_' + file, 720, 405]]);
  const requests = gas.decks[0].requests;
  assert.strictEqual(requests.filter(r => r.createSlide).length, 8, '本編4枚と付録A・C（責任者別・一覧）・D（停滞が無いので付録B は無い）');
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[2]), /^RXビジネスMTG｜全体（10\/23 会議）\n全体サマリー\n/);
  assert.strictEqual(gas.decks[2].requests.filter(r => r.createSlide).length, 2);
  assert.deepStrictEqual(requests.filter(r => r.deleteObject).map(r => r.deleteObject.objectId), ['p'], '作ったときの白紙のスライドを消す');
  const batches = gas.decks[0].batches;
  assert.ok(batches.length >= 2 && batches.length < 8, '何枚かずつまとめて送る（' + batches.length + '回）');
  assert.ok(requests.findIndex(r => r.deleteObject) < batches[0], '白紙のスライドは1回目で消す');
  assert.deepStrictEqual(plain(gas.driveFolders), [{id: 'folder-1', name: '会議サマリ', parent: 'parent-of-spreadsheet-id'}],
    'スプレッドシートと同じフォルダの中に作る');
  assert.deepStrictEqual(gas.driveFiles.map(f => [f.name, f.folder, f.blob.from, f.blob.getContentType()]), [
    ['会議サマリ_新FMT_' + file + '.pdf', 'folder-1', 'deck-1', 'application/pdf'],
    ['会議サマリ_新FMT2_' + file + '.pdf', 'folder-1', 'deck-2', 'application/pdf'],
    ['会議サマリ_全体_' + file + '.pdf', 'folder-1', 'deck-3', 'application/pdf']
  ]);
  assert.deepStrictEqual(plain(gas.moved).map(m => m.folder), ['folder-1', 'folder-1', 'folder-1'], 'スライドも同じフォルダに残す');
  assert.deepStrictEqual(gas.trashed, []);
  assert.deepStrictEqual(gas.writes.filter(w => w.ss === gas.ss.getId()), [], 'スプレッドシートには書き込まない');
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /「新FMT」：会議サマリ_新FMT_.*\.pdf（動いた 3・停滞 0・要対応 3）/);
  assert.match(done, /全体サマリー：会議サマリ_全体_.*\.pdf（追跡シート 2枚）/);
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
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：.*（動いた 1・停滞 0・要対応 2）/);
  assert.deepStrictEqual(gas.driveFiles.slice(3).map(f => f.name.replace(/_\d{4}-\d{2}-\d{2}/, '')), ['会議サマリ_新FMT.pdf', '会議サマリ_全体.pdf'],
    '選んだシートだけ作る（全体サマリーはいつも作る）');
  assert.match(inserted(gas.decks[4]), /\n新FMT2\n/, '全体サマリーには、選ばなかったシートも入れる');
  // 同じ会議日で作り直しても、前回（10/23）からの変化（作った日時だけ上書き）
  gas.promptAnswers = ['1', '10/30'];
  g.createMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（動いた 1・/);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).map(m => m.date), [meeting, next]);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT2')).map(m => m.date), [meeting], '選ばなかったシートの会議日は記録しない');
});

test('previewMeetingSummaries：試作は会議日を記録しない（ファイル名に「試作」）。最初の試作の要対応を覚え、同じ会議日の資料で解消した件数を出す', () => {
  const {gas, g, sheet} = setup();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.ok(gas.driveFiles.every(f => /_試作\.pdf$/.test(f.name)));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（動いた 3・停滞 0・要対応 3）[\s\S]*試作のため、会議日は記録していません。/, '2回目も同じ変化を出す');
  const props = g.PropertiesService.getDocumentProperties();
  assert.strictEqual(props.getProperty('MTG_MEETINGS_新FMT'), null);
  const draft = JSON.parse(props.getProperty('MTG_DRAFT_新FMT'));
  assert.deepStrictEqual([draft.meeting, draft.keys.length], [g.mtgParseDate_('10/23', new Date()), 3]);
  assert.deepStrictEqual(gas.decks.map(d => d.title.replace(/_\d{4}-\d{2}-\d{2}/, '')),
    ['会議サマリ_新FMT_試作', '会議サマリ_全体_試作', '会議サマリ_新FMT_試作', '会議サマリ_全体_試作']);
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[2]), /要対応は 3件（責任者 3人）。試作（\d{2}\/\d{2}）の時点から 0件を解消・新たに 0件。/, '2回目の試作：まだ解消していない');

  // E社の活動状況を書いて「更新なし」を解消してから、本番を作る
  g.recordDiffEdit(gas.edit(sheet, 'J7', '10/20 決裁者と面談を設定'));
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(inserted(gas.decks[4]), /要対応は 2件（責任者 2人）。試作（\d{2}\/\d{2}）の時点から 1件を解消・新たに 0件。/);
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
