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

/** 今日（日本時間）から days 日ずらした日 'yyyy/MM/dd' */
function day(days) {
  const d = new Date(Date.now() + 9 * 3600000 + days * 86400000);
  return d.getUTCFullYear() + '/' + String(d.getUTCMonth() + 1).padStart(2, '0') + '/' + String(d.getUTCDate()).padStart(2, '0');
}

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
    g.recordDiffEdit(gas.edit(sheet, 'J6', '10/12 打ち合わせ'));         // D社：活動状況（活動の更新）
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
 * options.trackedSince：変更の記録を始めた時刻（既定は変更履歴_差分 の最初の行）。options.draft：試作の入力内容のチェックの結果。
 * options.baseline：基準を記録した時刻
 */
function summarize(g, name, options) {
  const o = options || {};
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const since = new Date(Date.now() - 14 * 86400000);
  const windowStart = new Date(Date.now() - 180 * 86400000);
  return g.mtgSummarize_(ss, name, {
    since, previous: null, legacy: null, baseline: o.baseline || null, draft: o.draft || null, final: true, now, meeting: '2026/10/23', windowStart,
    copy: o.copy || null, copyError: '',
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
  assert.deepStrictEqual(['2026年11月', '2026/11/01', '未定', ''].map(g.mtgMonthLabel_), ['2026/11', '2026/11', '未定', '未入力'], '空は「未入力」');
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

test('前回からの変化：案件ごとに1つの区分（失注・延期・前進・新規・削除・活動の更新・更新のみ）。区分の順（受注・失注・保留・後退・削除 → 新規・前進・復帰 → 活動更新 → 受注月・期待値などの変更）', () => {
  const {gas, g, sheet} = setup();
  g.recordDiffEdit(gas.edit(sheet, 'I7', month(3).text));                          // E社：受注月を延ばす
  g.recordDiffEdit(gas.edit(sheet, 'D8:F8', [['G社', 'POS連携', '初回訪問']]));      // G社：空の行に新しい案件
  g.recordDiffEdit(gas.edit(sheet, 'K4', '請求の時期を確認中'));                    // B社：活動における課題だけ（更新のみ）
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.category]), [
    ['C社／店舗什器', 'lost'],
    ['F社／ブランド刷新', 'removed'],
    ['G社／POS連携', 'new'],
    ['A社／アプリ刷新', 'forward'],
    ['D社／物流改善', 'activity'],
    ['E社／販促DX', 'slipped'],
    ['B社／EC立ち上げ', 'updated']
  ]);
  assert.ok(s.changes[3].flags.indexOf('activity') >= 0, 'A社は活動状況も変わったが、区分は前進（上の区分が先）');
  const a = s.changes[3];
  assert.deepStrictEqual([a.statusFrom, a.statusTo, a.activity, a.columns.join('・')], ['初回訪問', '提案中', '10/10 提案書を提出', '案件ステータス・活動状況']);
  const e = s.changes[5];
  assert.deepStrictEqual([e.monthFrom, e.monthTo, e.weightedFrom, e.weightedTo], [month(1).text, month(3).text, 40, 40]);
  assert.deepStrictEqual([s.changes[1].removed, s.changes[1].statusFrom, s.changes[1].weightedFrom], [true, '提案中', 3], '削除した行は失注に数えない');
  const k = s.kpi;
  assert.deepStrictEqual([k.changed, k.activity, k.updated, k.lost, k.forward, k.new, k.removed, k.slips, k.slipped], [6, 1, 1, 1, 1, 1, 1, 1, 1]);

  // 活動状況を書いて元に戻した案件は、活動の更新に数えない（前回の値と今の値で比べる）
  g.recordDiffEdit(gas.edit(sheet, 'J7', '仮の書き込み'));
  g.recordDiffEdit(gas.edit(sheet, 'J7', '提案書を送付'));
  assert.strictEqual(plain(summarize(g, '新FMT')).changes.find(c => c.customer === 'E社').flags.indexOf('activity'), -1);
});

test('前回からの変化：得意先・案件名を書き換えた案件・あとから案件名を入れた案件は、前の名前の記録をつなげる', () => {
  const {gas, g, sheet} = setup({edits: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));                 // A社：前進
  g.recordDiffEdit(gas.edit(sheet, 'E3', 'アプリ刷新（第2期）'));     // 案件名を書き換える
  g.recordDiffEdit(gas.edit(sheet, 'D9', 'H社'));                    // 得意先だけ先に入れて
  g.recordDiffEdit(gas.edit(sheet, 'E9', '店舗分析'));               // あとから案件名
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.category, c.columns.join('・')]), [
    ['H社／店舗分析', 'new', '得意先・案件名'],
    ['A社／アプリ刷新（第2期）', 'forward', '案件ステータス・案件名']
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
    ['E社／販促DX', 'won'], ['F社／ブランド刷新', 'hold'], ['C社／店舗什器', 'backward'], ['A社／アプリ刷新', 'forward'], ['B社／EC立ち上げ', 'revived']
  ]);
  const b = s.bridge;
  assert.deepStrictEqual([b.start, b.end, b.startCount, b.endCount], [58, 67, 5, 4], '前回：A10・C5・D空・E40・F3／今：A15・B50・C2・D空');
  assert.deepStrictEqual(b.steps.filter(st => st.count).map(st => [st.label, st.value, st.count]),
    [['期待値増', 5, 1], ['復帰', 50, 1], ['期待値減', -3, 1], ['受注', -40, 1], ['保留', -3, 1]]);
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

test('前回の時点の値：前回より前の変更の記録があれば、そのスナップショット（数式の期待値も、その時点の計算した値）から取る', () => {
  const {gas, g, sheet} = setup({edits: false});
  aged(gas, 20, () => g.recordDiffEdit(gas.edit(sheet, 'K3', '常務に説明済み')));   // A社：20日前の変更（スナップショット：初回訪問・期待値 10）
  sheet.cell(3, COL.weighted, true).v = 15;                                      // 期待値が数式で変わった（変更の記録には残らない）
  g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));                              // 今日：ステータスを変えた
  const s = plain(summarize(g, '新FMT'));
  const a = s.changes.find(c => c.customer === 'A社');
  assert.deepStrictEqual([a.category, a.base.source, a.statusFrom, a.weightedFrom, a.weightedTo], ['forward', 'snapshot', '初回訪問', 10, 15],
    '前回（14日前）の値は、その前でいちばん新しいスナップショット（20日前）から');
  assert.deepStrictEqual(s.bridge.steps.filter(st => st.count).map(st => [st.key, st.value]), [['up', 5]], '期待値の増減に、数式で変わった分も入る');
});

test('期待値の列が数式で、前回より前のスナップショットが無いとき：ステータス・想定売上が変わった案件は、定義の率で前回の期待値を計算する', () => {
  const {gas, g, sheet} = setup({edits: false});
  gas.addSheet('定義_ステータス', [['確度', 'ステータス', '期待値'], ['A', '受注', 1], ['B', '受注見込み', 0.7], ['C', '提案中', 0.5], ['D', '初回訪問', 0.2]],
    {rows: 10, columns: 3});
  sheet.cell(3, COL.weighted, true).f = '=IF(F3:F="", , G3:G * XLOOKUP(F3:F, 定義_ステータス!B2:B, 定義_ステータス!C2:C))';
  g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));     // A社：初回訪問 → 提案中（想定 30）
  g.recordDiffEdit(gas.edit(sheet, 'J7', '提案書を再送'));   // E社：活動状況だけ
  const s = plain(summarize(g, '新FMT'));
  const a = s.changes.find(c => c.customer === 'A社');
  assert.deepStrictEqual([a.base.source, a.weightedFrom], ['log', 6], '前回の期待値 = 前回の想定売上 30 × 初回訪問の率 0.2（シートの数式と同じ）');
  const e = s.changes.find(c => c.customer === 'E社');
  assert.strictEqual(e.weightedFrom, e.weightedTo, 'ステータス・想定売上が変わっていなければ、期待値は今の値のまま');
  assert.strictEqual(e.category, 'activity');
});

test('ステータスの進み具合：定義のシートの率の大きい順、率が同じなら定義の下の行ほど進んでいる（本物の定義_ステータス）。無ければプルダウンの順', () => {
  const {g} = setup({edits: false});
  const statuses = ['A-受注済み(100%)', 'B-クロージング中(70%)', 'C-提案中(50%)', 'C-選定対応中(50%)', 'C-稟議対応中(50%)', 'D-ヒアリング中(20%)',
    'D-課題整理中(20%)', '創注-仮検討(10%)'];
  const ratios = [1, 0.7, 0.5, 0.5, 0.5, 0.2, 0.2, 0.1];
  const rank = g.mtgStatusRank_(statuses, statuses.map((status, i) => ({status, ratio: ratios[i]})));
  const ascending = ['創注-仮検討(10%)', 'D-ヒアリング中(20%)', 'D-課題整理中(20%)', 'C-提案中(50%)', 'C-選定対応中(50%)', 'C-稟議対応中(50%)',
    'B-クロージング中(70%)', 'A-受注済み(100%)'];
  ascending.slice(1).forEach((status, i) => assert.ok(rank(status) > rank(ascending[i]), status + ' は ' + ascending[i] + ' より進んでいる'));
  assert.strictEqual(rank('検討中'), -1);
  assert.deepStrictEqual(statuses.map(g.mtgStatusKind_), ['won', 'active', 'active', 'active', 'active', 'active', 'active', 'active']);
  const byOrder = g.mtgStatusRank_(['初回訪問', '提案中', '受注'], null);
  assert.ok(byOrder('提案中') > byOrder('初回訪問'), '定義のシートが無ければプルダウンの順');
  assert.strictEqual(g.mtgStatusRatio_(statuses.map((status, i) => ({status, ratio: ratios[i]})), 'C-提案中(50%)'), 0.5);
  assert.strictEqual(g.mtgStatusLines_('B-クロージング中(70%)', 90, 9), 'B-クロージング中\n(70%)', '1行に入らなければ率の前で改行（言葉の途中で切らない）');
  assert.strictEqual(g.mtgStatusLines_('C-提案中(50%)', 120, 9), 'C-提案中(50%)', '入れば1行');
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
  g.addSampleHistory();   // A社に、3日前が予定日で未実施のオファリングなどを足す（サンプル。予定日超過にも変化にも数えない）
  // 当社の担当者が登録した、3日前が予定日で未実施のオファリング
  const cred = gas.get('CRED_OPTIONS');
  const {log, cols} = g.credLogSheet_(gas.ss, false);
  const fields = {kind: cred.linking.childKind, name: '販促DXの試験導入のご提案', person: '佐藤 花子', title: '販促企画部 課長', owner: '鈴木'};
  fields[cred.schedule.planKey] = day(-3);
  fields[cred.schedule.doneKey] = '';
  g.credAppendLog_(log, cols, [{id: 'cred-real-0001', at: new gas.Date(), customer: 'A社', fields: g.credValidate_(fields, null), link: '',
    editor: 'suzuki@example.com', source: 'テスト'}]);
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
  assert.match(text('予定日超過'), /^オファリング「販促DXの試験導入のご提案」：予定日（\d{1,2}\/\d{1,2}）から 3日経過、実施日未入力$/);
  assert.strictEqual(s.alerts.filter(a => a.type === '予定日超過').length, 1, 'サンプル（【サンプル】）の予定は数えない');
  assert.strictEqual(s.alerts[0].owner, '鈴木', '責任者の列には当社の担当者');
  assert.strictEqual(text('受注月超過'), '受注月（' + month(-1).y + '/' + month(-1).m + '）経過、ステータス「初回訪問」（延期 1回）');
  assert.strictEqual(text('未入力', 'D社'), '未入力：想定売上規模（百万）・期待値調整済想定売上規模（百万）・受注月・活動状況・アカウント責任者');
  assert.strictEqual(text('未入力', 'F社'), '未入力：活動状況', '想定の大きい順');
  assert.strictEqual(text('更新なし'), '最終更新から 35日経過');
  assert.strictEqual(text('金額の不整合'), '期待値（40）が想定売上（30）を上回る');
  assert.strictEqual(text('月の不整合'), '売上開始月（' + month(-3).y + '/' + month(-3).m + '）が受注月（' + month(-1).y + '/' + month(-1).m + '）より前');
  assert.strictEqual(text('選択肢外のステータス'), '「検討中」は選択肢外の値（進行中として集計）');
  assert.strictEqual(text('デリバリー担当なし'), '受注済み、「デリバリー担当」未入力');
  assert.ok(s.changes.every(c => !/^【サンプル】/.test(c.project)), 'サンプルの記録は変化に入れない');
  assert.strictEqual(s.kpi.owners, 6, '鈴木・佐藤・高橋・空・田中・伊藤');
});

test('入力内容のチェック「資料なし」：商談記録（ＵＲＬ）の見出しがあるシートだけ、進行中の案件で空のもの（見出しの全角・半角は問わない）', () => {
  const {gas, g, sheet} = setup({edits: false});
  gas.asUser(() => sheet.getRange(2, 15).setValue('商談記録（URL）'));                                // O列。設定は全角の「ＵＲＬ」
  gas.asUser(() => sheet.getRange(3, 15).setValue('https://drive.google.com/drive/folders/abc'));   // A社だけ格納先がある
  const s = plain(summarize(g, '新FMT'));
  const empty = '「商談記録（ＵＲＬ）」未入力';
  assert.deepStrictEqual(s.alerts.filter(a => a.type === '資料なし').map(a => [label(a), a.text, a.owner]), [
    ['E社／販促DX', empty, '伊藤'],
    ['C社／店舗什器', empty, '田中'],
    ['F社／ブランド刷新', empty, '高橋'],
    ['D社／物流改善', empty, '']
  ], '進行中だけ（受注の B社は見ない）。想定売上の大きい順');
  const types = s.alerts.map(a => a.type);
  assert.ok(types.lastIndexOf('未入力') < types.indexOf('資料なし') && types.lastIndexOf('資料なし') < types.indexOf('更新なし'), '未入力の次');

  const slides = plain(g.mtgSlides_(summarize(g, '新FMT')));
  const states = slides[1].elements.filter(el => el.type === 'table')[1];
  assert.deepStrictEqual(states.rows[2].cells.slice(0, 3).map(cell => cell.text), ['商談記録未登録', '4件', '110'],
    '商談記録の見出しがあるシートでは、サマリーの現在の状態に商談記録未登録の行を出す');
  const owners = slides.find(sl => sl.title === 'Appendix ③　入力内容の確認事項（責任者別）');
  assert.ok(owners.elements.find(el => el.type === 'table').rows[0].cells.some(cell => cell.text === '商談記録\n未登録'));
  const list = slides.find(sl => sl.title === 'Appendix ③　入力内容の確認事項（一覧）').elements.find(el => el.type === 'table');
  assert.deepStrictEqual(list.rows[0].cells.map(cell => cell.text), ['責任者', '得意先／案件名', '項目', '内容', '想定\n売上']);
  const all = gas.get('MTG_ALERT_TYPES').map((type, i) => ({type, key: type + '|' + i, dealKey: 'k' + i, customer: 'X社', project: '案件' + i, text: '内容', amount: 1, owner: '担当'}));
  const wide = plain(g.mtgSlides_(Object.assign({}, summarize(g, '新FMT'), {alerts: all})))
    .filter(sl => /入力内容の確認事項（一覧）/.test(sl.title)).map(sl => sl.elements.find(el => el.type === 'table'));
  const labels = wide.map(t => t.rows.slice(1).map(r => r.cells[2].text).join('\n')).join('\n');
  assert.ok(/(^|\n)(デリバリー担当未設定|デリバリー\n担当未設定)(\n|$)/.test(labels) && labels.indexOf('未設\n定') < 0,
    '長い種類の名前は1行に入れ、入らなければ言葉の切れ目で2行に：' + labels);
  assert.ok(list.rows.some(r => r.cells[0].text === '伊藤' && r.cells[2].text.split('\n').indexOf('商談記録未登録') >= 0 &&
    r.cells[3].text.split('\n')[r.cells[2].text.split('\n').indexOf('商談記録未登録')] === empty), '内容は案件ごとの一覧に書く（項目は資料に出す名前。項目と内容の行をそろえる）');
});

test('受注月（四半期）別の期待値（受注月超過 → 今の四半期〜4Q → 来期以降 → 未入力）・ステータス別（プルダウンの順。進行中 → 保留 → 受注 → 失注）', () => {
  const {g} = setup();
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.stages.map(p => [p.status, p.kind, p.count, p.amount, p.weighted]), [
    ['提案中', 'active', 2, 110, 50],
    ['未入力', 'active', 1, 0, 0],
    ['受注', 'won', 1, 50, 50],
    ['失注', 'lost', 1, 20, 5]
  ]);
  const inYear = m => {
    const now = month(0);
    const start = (now.m >= 4 ? now.y : now.y - 1) * 12 + 4;
    return m.y * 12 + m.m >= start && m.y * 12 + m.m < start + 12;
  };
  const planned = inYear(month(1)) ? 40 : 0;   // A社（先月・10）は受注月超過。E社（来月・40）は今期か来期以降
  const lead = leadOf(plain(g.mtgSlides_(summarize(g, '新FMT')))[2]);
  assert.ok(lead.indexOf('受注月が今期の期待値は ' + planned + '百万円') >= 0, '受注月別のグラフ（今の四半期〜4Q）の合計：' + lead);
  assert.ok(lead.indexOf('受注月超過は 10百万円') >= 0, '受注月を過ぎた進行中の期待値は四半期に入れない');
  assert.strictEqual(s.fiscal, undefined, '今期の受注・着地見込み（サマリーの帯）は作らない');
  assert.strictEqual(s.quarters.reduce((sum, q) => sum + q.count, 0), 3, '進行中の3件をどこかに入れる');
  assert.deepStrictEqual(s.quarters[0], {label: '受注月超過', value: 10, count: 1}, '受注月超過は先頭');
  assert.deepStrictEqual(s.quarters[s.quarters.length - 1], {label: '受注月未入力', value: 0, count: 1}, '受注月未入力は最後');
  const all = ['1Q（4〜6月）', '2Q（7〜9月）', '3Q（10〜12月）', '4Q（1〜3月）'];
  const current = Math.floor(((month(0).m + 8) % 12) / 3);   // 4月 → 0（1Q）
  assert.deepStrictEqual(s.quarters.filter(q => /Q/.test(q.label)).map(q => q.label), all.slice(current), '過ぎた四半期は出さない（今の四半期から）');
  assert.strictEqual(s.quarters.some(q => q.label === '来期以降'), !inYear(month(1)), '来期以降は当てはまる案件があるときだけ');
});

/* ---------------- スライド ---------------- */

/** 部品の文字（表はセルをつないだもの） */
const textOf = slide => slide.elements.map(el => (el.type === 'table'
  ? el.rows.map(r => r.cells.map(c => c.text).join(' | ')).join('\n') : String(el.text || ''))).join('\n');
/** 役割（role）の部品 */
const roleOf = (slide, role) => slide.elements.filter(el => el.role === role);
/** リード文（1つ目の文と補足をつなげ、改行を除く） */
const leadOf = slide => roleOf(slide, 'lead').map(el => el.text).join('').replace(/\n/g, '');

test('スライド：表紙、本編4枚（サマリー → パイプラインの変化 → 前回から変化のある案件 → 期待値上位案件の状況・課題）、扉、Appendix ①〜③。見出しはどのシートでも同じ、状態はリード文に書く', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const T = gas.get('MTG_TITLES');
  const c = gas.get('MTG_OPTIONS').colors;
  const s = summarize(g, '新FMT', {trackedSince: Date.now() - 100 * 86400000});
  const slides = plain(g.mtgSlides_(s));
  const titles = slides.map(sl => sl.title);
  assert.deepStrictEqual(titles, [T.cover, T.summary, T.pipeline, T.progress, T.focus, T.appendix, T.changes, T.stalled, T.owners, T.alerts]);
  assert.deepStrictEqual(titles, ['表紙', 'サマリー', 'パイプラインの変化', '前回から変化のある案件', '期待値上位案件の状況・課題', 'Appendix',
    'Appendix ①　前回から変化のある案件（一覧）', 'Appendix ②　停滞案件（一覧）', 'Appendix ③　入力内容の確認事項（責任者別）',
    'Appendix ③　入力内容の確認事項（一覧）'], 'ステータスの分類・用語の定義の Appendix は作らない');
  assert.ok(!('status' in T) && !('glossary' in T));

  // 表紙（ひな形の1枚目）：決まった文字を、サブインダストリー・会議の名前と事前資料・会議日・（名前の欄は空）に置き換え、
  // その下にデータの時点・期間を白い字で足す。出典の行・ページ番号は置かない
  assert.deepStrictEqual(slides[0].template, {title: '新FMT', subtitle: 'RX Business MTG　事前資料',
    date: slides[0].template.date, name: ''});
  assert.ok(slides[0].cover && !slides[0].divider);
  assert.match(slides[0].template.date, /^\d{4}年10月23日（[日月火水木金土]）$/);
  assert.deepStrictEqual(slides[0].elements.map(el => el.text.replace(/\d{4}\/\d{1,2}\/\d{1,2} \d{2}:\d{2}|\d{1,2}\/\d{1,2}/g, '#')), [
    '出典：「新FMT」シートおよび変更履歴（# 時点）', '対象期間：直近14日間（#〜。初回のため）']);
  assert.ok(slides[0].elements.every(el => el.color === c.white && el.align === 'CENTER'));
  assert.ok(!['title', 'footer', 'page', 'lead'].some(role => roleOf(slides[0], role).length), '表紙には見出し・出典の行・ページ番号を置かない');
  assert.strictEqual(plain(g.mtgSlides_(Object.assign({}, s, {final: false})))[0].template.subtitle, 'RX Business MTG　事前資料【試作】');
  assert.strictEqual(plain(g.mtgSlides_(Object.assign({}, s, {meeting: '2026/11/06'})))[0].template.date, '2026年11月6日（金）');

  // 扉（ひな形の1枚目を複製）：見出しを「Appendix」に（ほかの決まった文字は空）、その下に Appendix の見出し・始まりのページ（「（続き）」のページは出さない）
  const divider = slides[5];
  assert.ok(divider.divider && !divider.cover);
  assert.ok(!['title', 'footer', 'page'].some(role => roleOf(divider, role).length), '扉には見出し・出典の行・ページ番号を置かない');
  assert.deepStrictEqual(divider.template, {title: 'Appendix', subtitle: '', date: '', name: ''});
  const toc = roleOf(divider, 'toc').map(el => el.text);
  const contents = toc.filter((_, i) => i % 3 === 0).map((label, i) => [label, toc[i * 3 + 1], toc[i * 3 + 2]]);
  assert.deepStrictEqual(contents, [['Appendix ①', '前回から変化のある案件', '7'], ['Appendix ②', '停滞案件', '8'], ['Appendix ③', '入力内容の確認事項', '9']],
    'Appendix ごとに1行（③の責任者別・一覧も1行。見出しの（一覧）などは付けない）');

  // 左上のラベルは置かない。どのサブインダストリー・会議日の資料かは、表紙といちばん下の出典の行（区切り線の下）に書く
  slides.forEach((slide, i) => {
    assert.ok(!slide.elements.some(el => el.type === 'box' && el.x === 0 && el.y === 0 && el.text), (i + 1) + '枚目：左上のラベル');
    if (i === 0 || i === 5) return;
    assert.deepStrictEqual(roleOf(slide, 'footer').map(el => [el.text, el.y, el.color]), [['新FMT（10/23 会議資料）　金額の単位：百万円', S.footY, c.faint]],
      (i + 1) + '枚目の出典の行（金額の単位も）');
    assert.ok(slide.elements.some(el => el.type === 'box' && el.y === S.ruleY && el.h < 1 && el.fill === c.rule), (i + 1) + '枚目の区切り線');
    assert.deepStrictEqual(roleOf(slide, 'title').map(el => [el.y, el.bold, el.color]), [[S.titleY, true, c.ink]], (i + 1) + '枚目の見出し（黒の太字）');
  });
  assert.strictEqual(roleOf(plain(g.mtgSlides_(Object.assign({}, s, {final: false})))[1], 'footer')[0].text, '【試作】新FMT（10/23 会議資料）　金額の単位：百万円');
  const other = plain(g.mtgSlides_(summarize(g, '新FMT2', {trackedSince: Date.now() - 100 * 86400000}))).map(sl => sl.title);
  assert.deepStrictEqual(other.slice(0, 5), titles.slice(0, 5), '本編の見出しは、中身が違うシートでも同じ');

  // 事前資料には事実だけを書く（論点・決めてほしいこと・支援の要否などの判断は、会議でリーダーが話す）
  const all = slides.map(textOf).join('\n');
  ['決めて', '論点', '★', '支援', '候補', '確認したい', '要対応', '悪い', '打開', '重点', '付録'].forEach(word => assert.ok(all.indexOf(word) < 0, '使わない言葉：' + word));
  slides.forEach(slide => slide.elements.filter(el => el.type === 'table').forEach(t => assert.ok(!(t.borders || []).length, '赤枠で強調しない')));

  // リード文（タイトルの下）に状態を書く。1つ目の文と補足（改行のあと）は行を分ける。大きさ・色はどのページも同じ
  const lead = leadOf;
  assert.strictEqual(lead(slides[1]), '進行中案件 3件（直近14日間で ▲2件）、期待値合計 50百万円（直近14日間で ▲8）。' +
    'ステータス未入力 1件は進行中に含め、想定売上未入力 1件・期待値未入力 1件は金額の合計に含めていない。');
  assert.deepStrictEqual(roleOf(slides[1], 'lead').map(el => el.text.split('\n').length), [1, 1],
    '空欄の扱い（D社：ステータス・想定売上・期待値が空）は補足の行に書く（推定で補わない）');
  assert.deepStrictEqual(roleOf(slides[2], 'lead').map(el => [el.color, el.size]), [[c.ink, S.leadSize], [c.ink, S.leadSize]]);
  slides.forEach((slide, i) => roleOf(slide, 'lead').forEach(el => assert.deepStrictEqual([el.size, el.color], [S.leadSize, c.ink], (i + 1) + '枚目：リード文の大きさ・色')));
  assert.match(lead(slides[2]), /^進行中の期待値は 58 → 50百万円（▲8）。増加 ±0・減少 ▲8、内訳の最大は失注（1件・▲5）。受注月が今期の期待値は /);
  assert.strictEqual(lead(slides[3]), '直近14日間に変化のある案件は 4件。', '区分ごとの件数（サマリーにある）・0件のものは書かない');
  assert.strictEqual(lead(slides[4]), '進行中案件のうち期待値上位 3件（進行中案件の期待値合計の 100%）。', '該当事項の数は書かない（表のラベルで分かる）');
  slides.forEach((slide, i) => roleOf(slide, 'lead').forEach(el => {
    assert.ok(!/…$/.test(el.text), (i + 1) + '枚目：リード文が切れている ' + el.text);
    assert.ok(el.y + el.h <= S.top + 0.01, (i + 1) + '枚目：リード文が中身にかかる');
  }));

  slides.forEach((slide, i) => {
    const page = roleOf(slide, 'page')[0];
    assert.ok(i === 0 || i === 5 ? !page : page && page.text === String(i + 1) && page.x > 600 && page.y > 370, (i + 1) + '枚目のページ番号（表紙・扉には出さない）');
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
          assert.ok(g.mtgTextHeight_(lines, cell.size, cell.spans) <= r.h + 0.01, '行の高さ：' + cell.text);
        }));
      }
    });
  });

  // 見出しとリード文の間はあける。表はどれも中身の幅いっぱい（列の幅を中身の幅に合わせる）
  slides.forEach((slide, i) => {
    const t = roleOf(slide, 'title')[0];
    const l = roleOf(slide, 'lead')[0];
    if (t && l) assert.ok(l.y - (t.y + 30) >= 5, (i + 1) + '枚目：見出しとリード文の間');
    slide.elements.filter(el => el.type === 'table' && el.columns.length > 4 && !/責任者別/.test(slide.title))
      .forEach(el => assert.ok(Math.abs(el.x + el.columns.reduce((a, b) => a + b, 0) - (S.x + S.w)) < 0.01, (i + 1) + '枚目：表が中身の幅いっぱい'));
  });
  // 1枚目：「動き」と「現在の状態」の2つの表（どのシートでも同じ区分。該当が無ければ0件）。数字の箱・今期の帯は置かない
  const first = textOf(slides[1]);
  assert.ok(!/今期|着地/.test(first), 'サマリーに今期の受注・着地見込みは置かない');
  // 上の数字は進行中案件全体の合計（下の表の区分とは重ならない）。前回比つき
  assert.deepStrictEqual(roleOf(slides[1], 'tile').map(el => el.text), ['進行中案件', '想定売上合計（進行中・百万円）', '期待値合計（進行中・百万円）', '入力内容の確認事項']);
  const sum = plain(summarize(g, '新FMT'));
  assert.strictEqual(sum.bridge.amountEnd, sum.kpi.amount, '想定売上合計は進行中案件の想定売上の合計');
  assert.deepStrictEqual(roleOf(slides[1], 'tile-value').map(el => el.text), [
    '3件　直近14日間で ▲2件',
    g.mtgAmount_(sum.bridge.amountEnd) + '　直近14日間で ' + g.mtgSigned_(sum.bridge.amountEnd - sum.bridge.amountStart),
    '50　直近14日間で ▲8',
    '3項目　3案件・責任者 3名'
  ], '確認事項は項目の数（1案件に複数あれば別に数える）と、案件・責任者の数');
  assert.ok(roleOf(slides[1], 'tile-value').every(el => el.spans[0].size === 15 && el.spans[0].color === c.greenDark), '数字は大きく、前回比は小さく');
  assert.deepStrictEqual(roleOf(slides[1], 'heading').map(el => el.text), ['直近14日間の動き（変化のある案件 4件）', '現在の状態（進行中案件）']);
  const [movesTable, statesTable] = slides[1].elements.filter(el => el.type === 'table');
  const textsOf = t => t.rows.map(r => r.cells.map(cell => cell.text.replace(/\n/g, '')));
  assert.deepStrictEqual(textsOf(movesTable), [
    ['区分', '件数', '想定売上', '主な案件（想定売上順）'],
    ['受注', '0件', '—', '該当なし'],
    ['失注・保留・後退・削除', '2件', '30', 'C社／店舗什器（失注）、F社／ブランド刷新（削除）'],
    ['新規・前進・復帰', '1件', '30', 'A社／アプリ刷新（初回訪問 ▶ 提案中）'],
    ['活動更新', '1件', '0', 'D社／物流改善（10/12 打ち合わせ）'],
    ['受注月・期待値などの変更', '0件', '—', '該当なし']
  ], '行の順は P4 の区分の順と同じ。活動更新は独立した行にし、最新の活動状況を書く');
  assert.strictEqual(movesTable.rows.slice(1).reduce((sum, r) => sum + Number(r.cells[1].text.replace('件', '')), 0), 4,
    '動きの行の件数を合わせると、見出しの「変化のある案件」の件数になる（案件ごとに1回だけ数える）');
  assert.deepStrictEqual(textsOf(statesTable), [
    ['受注月超過', '1件', '30', 'A社／アプリ刷新（受注月 ' + month(-1).y + '/' + month(-1).m + '）'],
    ['停滞（同一ステータス45日以上）', '2件', '80', 'うち直近14日間の更新なし 1件：E社／販促DX（100日以上）'],
    ['商談記録未登録', '—', '—', 'データなし：シートに「商談記録（ＵＲＬ）」の列が無い']
  ], '停滞は定義を添え、名前は対象期間に更新の無い案件だけ（D社は活動を更新した）。商談記録の見出しが無いシートでも行は出し（型を変えない）、' +
    '数えられない理由を灰色のセルに書く');
  assert.deepStrictEqual([statesTable.rows[2].cells[3].fill, statesTable.rows[2].cells[3].color], [c.gap, c.muted], 'データなしは灰色のセル');
  const stall = statesTable.rows[1].cells[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify([stall.text, stall.size, stall.color, stall.spans])),
    ['停滞（同一ステータス45日以上）', 7.5, c.muted, [{start: 0, end: 2, size: S.bodySize, bold: true, color: c.ink}]],
    '定義は区分の名前のあとに小さな字で（同じ行に入るので、行を高くしない）');
  assert.strictEqual(statesTable.rows[1].h, statesTable.rows[0].h, '停滞の行も、ほかの行と同じ高さ');
  const counts = movesTable.rows.slice(1).concat(statesTable.rows).map(r => r.cells[1]).filter(cell => cell.text !== '—');   // —：数えられない行
  assert.ok(counts.every(cell => cell.size >= S.bodySize && cell.size === counts[0].size && cell.bold && cell.color === c.greenDark), '件数は緑の太字（どの行も同じ大きさ）');
  assert.deepStrictEqual(statesTable.columns, movesTable.columns, '2つの表の列はそろえる');
  // 表は緑の見出し・1列目は薄い緑（得意先は太字）・ほかは白
  const moves = slides[3].elements.find(el => el.type === 'table');
  assert.ok(moves.rows[0].cells.every(cell => cell.fill === c.green && cell.color === c.white && cell.bold));
  assert.deepStrictEqual([moves.rows[1].cells[0].fill, moves.rows[1].cells[0].spans, moves.rows[1].cells[2].fill],
    [c.greenSoft, [{start: 0, end: 2, bold: true}], c.white]);
  // 3枚目：区分はセルに重ねたラベル（前向きは濃い緑・後ろ向きは濃いグレー・そのほかは薄いグレー。枠なし）。活動状況だけ変わった案件（活動の更新）も、最後に活動状況と一緒に出す
  assert.match(moves.rows[1].cells.map(cell => cell.text).join(' | '), /^C社\n店舗什器 \|  \| 提案中 ▶ 失注 \| 5 \|/);
  const pills = roleOf(slides[3], 'pill');
  assert.deepStrictEqual(pills.map(el => [el.text, el.fill, el.color, el.line || null, el.shape]), [
    ['失注', c.dark, c.white, null, 'ROUND_RECTANGLE'],
    ['削除', c.dark, c.white, null, 'ROUND_RECTANGLE'],
    ['前進', c.greenDeep, c.white, null, 'ROUND_RECTANGLE'],
    ['活動更新', c.graySoft, c.sub, null, 'ROUND_RECTANGLE']
  ], 'サマリーの動きの表と同じ順');
  assert.strictEqual(new Set(pills.map(el => el.w)).size, 1, '区分のラベルの幅は同じ');
  assert.ok(pills.every(el => el.h === S.pill.h && el.size === S.pill.size));
  pills.forEach((el, i) => {
    const x0 = moves.x + moves.columns[0];
    const y0 = moves.y + moves.rows.slice(0, i + 1).reduce((a, r) => a + r.h, 0);
    assert.ok(el.x >= x0 + 4 && el.x + el.w <= x0 + moves.columns[1] - 4 && el.y >= y0 + 4 && el.y + el.h <= y0 + moves.rows[i + 1].h - 4,
      'ラベルがセルに余白をあけて入る：' + el.text);
    assert.deepStrictEqual(plain(g.mtgWrap_(el.text, el.w, el.size, 2)), [el.text], 'ラベルの文字が1行に入る：' + el.text);
  });
  assert.deepStrictEqual([moves.rows[4].cells[0].text.split('\n')[0], moves.rows[4].cells[5].text], ['D社', '10/12 打ち合わせ']);
  // 4枚目：期待値の大きい順。当てはまる状況は事実だけ
  const focus = slides[4].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(focus.rows[0].cells.map(cell => cell.text.replace(/\n/g, '')),
    ['得意先／案件名', '期待値／想定売上', 'ステータス（滞留日数）', '受注月', '活動状況', '活動における課題', '該当事項']);
  assert.deepStrictEqual(roleOf(slides[4], 'pill').map(el => [el.text, el.fill, el.line || null]),
    [['停滞', c.greenSoft, null], ['受注月超過', c.graySoft, null], ['停滞', c.greenSoft, null]], '当てはまる状況のラベル（E社・A社・D社の順。停滞は薄い緑、ほかは薄いグレー）');
  assert.strictEqual(new Set(roleOf(slides[4], 'pill').map(el => el.w)).size, 1, '当てはまる状況のラベルの幅は同じ');
  assert.match(focus.rows[2].cells.map(cell => cell.text).join(' | '), /^A社\nアプリ刷新 \| 10\n30 \| 提案中\n0日 \| \d{4}\/\d+ \|/);
  assert.deepStrictEqual([focus.rows[2].cells[1].spans, focus.rows[2].cells[2].spans],
    [[{start: 0, end: 2, bold: true, color: c.greenDark, size: 11}], [{start: 4, end: 6, color: c.muted}]], '期待値は大きく緑、日数はグレー');
  // Appendix：件数だけ（区分の決め方・並びの説明は書かない）
  assert.strictEqual(lead(slides[6]), '直近14日間に変化のある案件は 4件。');
  const appendixRows = slides[6].elements.find(el => el.type === 'table').rows.slice(1);
  assert.strictEqual(appendixRows.length, 4, 'Appendix ① の行の数は、サマリー・P4 の「変化のある案件」の件数と同じ（その他の項目のみの更新は入れない）');
  assert.deepStrictEqual(appendixRows.map(r => r.cells[0].text.split('\n')[0]), ['C社', 'F社', 'A社', 'D社'], 'P4 と同じ順');
  assert.strictEqual(lead(slides[7]), '進行中案件のうち、同一ステータスに 45日以上滞留している案件は 2件（想定売上合計 80百万円。想定売上未入力の 1件を除く）。' +
    'うち直近14日間の更新なし 1件。', '合計に含めていない空欄の件数を書く');
  assert.deepStrictEqual(roleOf(slides[7], 'lead').map(el => el.text.split('\n').length).slice(1), [1], '1つ目の文と、更新なしの件数は行を分ける');
  assert.ok(roleOf(slides[7], 'lead')[1].text.indexOf('うち直近14日間の更新なし') === 0);
  assert.strictEqual(lead(slides[9]), '入力内容の確認事項は 3項目（3案件）。');
  assert.ok(all.indexOf('Appendix ④') < 0 && all.indexOf('Appendix ⑤') < 0, '無い Appendix を案内しない');
  // Appendix ③：責任者別は表だけ（種類の説明は置かない。内容は一覧に書く）。見出しの無いシートでは資料なしを見ない
  assert.strictEqual(lead(slides[8]), '入力内容の確認事項は 3項目（3案件・責任者 3名）。チェックできなかった項目は 2種類（灰色の列）。',
    'クレデンシャルの記録のシートと商談記録の列が無いので、予定日超過・商談記録未登録はチェックできない');
  assert.deepStrictEqual(slides[8].elements.filter(el => !el.role && el.type === 'table').length, 1);
  assert.ok(all.indexOf('チェックの種類') < 0 && all.indexOf('資料なし') < 0);
});

test('空欄は推定で補わず「未入力」と書く（「—」や 0 にしない）。該当する案件が無いときは「該当なし」', () => {
  const {gas, g} = setup();   // D社：ステータス・想定売上・期待値・受注月・責任者・活動における課題が空（活動状況は書いた）
  const s = summarize(g, '新FMT', {trackedSince: Date.now() - 100 * 86400000});
  const slides = plain(g.mtgSlides_(s));
  const rowOf = (slide, customer) => slide.elements.filter(el => el.type === 'table').map(t => t.rows.find(r => r.cells[0].text.split('\n')[0] === customer))
    .find(Boolean).cells.map(cell => cell.text);
  assert.deepStrictEqual(rowOf(slides[3], 'D社'), ['D社\n物流改善', '', '未入力', '未入力', '未入力', '10/12 打ち合わせ', '未入力'],
    'P3：ステータス・期待値・受注月・責任者');
  assert.deepStrictEqual(rowOf(slides[4], 'D社').slice(1, 6), ['未入力\n未入力', '未入力\n100日以上', '未入力', '10/12 打ち合わせ', '未入力'],
    '期待値上位：期待値・想定売上・ステータス・受注月・課題');
  const focusCell = slides[4].elements.find(el => el.type === 'table').rows.find(r => /^D社/.test(r.cells[0].text)).cells[1];
  assert.ok(!focusCell.spans.some(sp => sp.size === 11 && sp.start === 0), '空の期待値は、数のような大きな緑の字にしない');
  assert.deepStrictEqual(rowOf(slides[7], 'D社').slice(1, 4).concat(rowOf(slides[7], 'D社').slice(7)).map(t => t.replace(/\n/g, ' ')),
    ['未入力', '未入力', '未入力', '未入力', '未入力', '10/12 打ち合わせ'], 'Appendix ②：想定売上・期待値・ステータス・責任者・課題');

  slides.forEach((slide, i) => slide.elements.filter(el => el.type === 'table').forEach(t => t.rows.forEach(r => r.cells.forEach(cell =>
    assert.ok(!/未…|未入\n力|未\n入力/.test(cell.text), (i + 1) + '枚目：「未入力」が切れる・折り返す ' + cell.text)))));

  // 金額の動き：空と 0 は別の値（空 → 10 は「未入力 ▶ 10」、空のままなら「未入力」）
  const ch = {base: true};
  assert.deepStrictEqual([g.mtgValueMove_(null, 10, ch), g.mtgValueMove_(null, null, ch), g.mtgValueMove_(0, null, ch), g.mtgValueMove_(3, 3, {removed: true})],
    ['未入力 ▶ 10', '未入力', '0 ▶ 未入力', '3 ▶ （削除）']);
  assert.strictEqual(g.mtgStatusMove_({statusTo: '', base: false}), '（新規） ▶ 未入力', '前回に無かった案件は「（新規）」から');
  assert.deepStrictEqual([g.mtgMonthMove_({base: true, monthFrom: '2026/09/01', monthTo: '2026年9月'}),
    g.mtgMonthMove_({base: true, monthFrom: '2026/09/01', monthTo: '2026年10月'})], ['2026/9', '2026/9 ▶ 2026/10'],
    '受注月は月で比べる（スナップショットとシートで書き方が違っても、同じ月なら動きにしない）');
  assert.deepStrictEqual(plain(g.mtgDealLabel_({customer: 'H社', project: ''})), 'H社／（案件名未入力）');
  assert.match(g.mtgDealCell_({customer: 'A社', project: '', dealKey: ''}, 200).text, /\n（得意先の予定）$/, '案件に結び付かない得意先の予定');
  assert.match(g.mtgDealCell_({customer: '', project: ''}, 200).text, /^（得意先未入力）\n（案件名未入力）$/);

  // ステータスが空なら、入力内容の確認事項（必須項目未入力）にも出す
  assert.strictEqual(plain(s.alerts).find(a => a.type === '未入力' && a.customer === 'D社').text,
    '未入力：案件ステータス・想定売上規模（百万）・期待値調整済想定売上規模（百万）・受注月・アカウント責任者');

  // 空欄が多くリード文の補足が長くても、サマリーの表はページに入る（数字の箱・余白を詰める）
  const crowded = plain(g.mtgSlides_(Object.assign({}, s, {kpi: Object.assign({}, s.kpi, {missingStatus: 12, missingAmount: 15, missingWeighted: 18})})))[1];
  assert.match(leadOf(crowded), /ステータス未入力 12件は進行中に含め、想定売上未入力 15件・期待値未入力 18件は金額の合計に含めていない。$/);
  crowded.elements.filter(el => el.type === 'table').forEach(t =>
    assert.ok(t.y + t.rows.reduce((a, r) => a + r.h, 0) <= gas.get('MTG_SLIDE').bottom + 0.01, 'サマリーの表がページの下にはみ出す'));
  assert.ok(crowded.elements.filter(el => el.role === 'tile-value').every(el => el.h >= 32 - 14), '数字の箱は下限より小さくしない');

  // 停滞の数を数えられないときは書く。期待値上位の該当事項が無ければ「該当なし」
  const fresh = plain(g.mtgSlides_(summarize(g, '新FMT', {trackedSince: Date.now() - 3 * 86400000})));
  const states = fresh[1].elements.filter(el => el.type === 'table')[1];
  assert.match(states.rows[1].cells[3].text, /^データなし：記録の開始（\d{1,2}\/\d{1,2}）から 45日未満のため、まだ数えられない（\d{1,2}\/\d{1,2} から数える）$/);
  assert.deepStrictEqual(states.rows[1].cells.slice(1, 3).map(cell => cell.text), ['—', '—'], '数えられないので 0件と書かない');
  assert.ok(fresh[4].elements.find(el => el.type === 'table').rows.slice(1).some(r => r.cells[6].text === '該当なし'), '該当事項が無い案件');
});

test('長い文（活動状況・課題）：入らなければ文字を小さくし（7pt まで）、それでも入らなければ grow のセルは行を伸ばして全文（「…」で切らない）', () => {
  const gas = createGas();
  const g = gas.global;
  const S = plain(gas.get('MTG_SLIDE'));
  const mid = '【提案進行中】★\n・9/3に電通からSEJ岡崎役員へ提案完了。次回は10月中旬に再提案の予定';
  const fit = plain(g.mtgFitLines_({text: mid, lines: 3, fit: true}, 160));
  assert.ok(fit.size < S.bodySize && fit.size >= S.minFitSize, '小さくして入れる（' + fit.size + 'pt）');
  assert.ok(fit.lines.length <= 3 && fit.lines.join('').indexOf('…') < 0, '全文');
  const short = plain(g.mtgFitLines_({text: '見積提出', lines: 3, fit: true}, 160));
  assert.deepStrictEqual([short.size, short.lines], [S.bodySize, ['見積提出']], '入る文はそのまま');
  const long = '・' + 'あいうえおかきくけこ'.repeat(30);
  const grown = plain(g.mtgFitLines_({text: long, lines: 3, fit: true, grow: true}, 160));
  assert.strictEqual(grown.size, S.minFitSize);
  assert.ok(grown.lines.length > 3 && grown.lines.join('').indexOf('…') < 0, '行を伸ばして全文（' + grown.lines.length + '行）');
  assert.strictEqual(grown.lines.join('').replace(/\s/g, ''), long, '文を落とさない');
  const clipped = plain(g.mtgFitLines_({text: long, lines: 3, fit: true}, 160));
  assert.ok(clipped.lines.length === 3 && /…$/.test(clipped.lines[2]), 'grow でなければ最後を「…」');
  const row = plain(g.mtgRow_([160], [{text: long, lines: 3, fit: true, grow: true}]));
  assert.strictEqual(row.cells[0].size, S.minFitSize, '表のセルの文字の大きさも小さくする');
  assert.ok(row.h > g.mtgTextHeight_(3, S.bodySize), '行が伸びる');
});

test('期待値上位案件：活動状況が長くて行が伸び、ページに入らなければ「（続き）」のページに続ける（全文・全件を出す）。Appendix の始まりのページもずれる', () => {
  const {gas, g} = setup();
  const s = summarize(g, '新FMT');
  const long = '【提案進行中】★\n' + Array.from({length: 6}, (_, i) => '・' + (i + 1) + '回目の打ち合わせで要件を確認し、見積と体制案を提示。先方の稟議の進み具合を確認する').join('\n');
  const focus = Array.from({length: 5}, (_, i) => Object.assign({}, s.focus[0], {customer: '得意先' + i, activity: long}));
  const slides = plain(g.mtgSlides_(Object.assign({}, s, {focus})));
  const pages = slides.filter(sl => /^期待値上位案件の状況・課題/.test(sl.title));
  assert.ok(pages.length >= 2, pages.length + 'ページ');
  assert.strictEqual(pages[1].title, '期待値上位案件の状況・課題（続き）');
  const tables = pages.map(p => p.elements.find(el => el.type === 'table'));
  tables.forEach(t => assert.strictEqual(t.rows[0].cells[0].text, '得意先／案件名', '見出しの行を繰り返す'));
  const rows = [].concat.apply([], tables.map(t => t.rows.slice(1)));
  assert.deepStrictEqual(rows.map(r => r.cells[0].text.split('\n')[0]), ['得意先0', '得意先1', '得意先2', '得意先3', '得意先4'], '全件・順のまま');
  rows.forEach(r => assert.strictEqual(r.cells[4].text.replace(/\s/g, ''), long.replace(/\s/g, ''), '活動状況は全文'));
  const S = gas.get('MTG_SLIDE');
  tables.forEach(t => assert.ok(S.top + t.rows.reduce((sum, r) => sum + r.h, 0) <= S.bottom + 0.1, 'ページに入る'));
  // 扉の目次の Appendix ① のページ番号も、ページが増えた分だけずれる
  const divider = slides.find(sl => sl.divider);
  const first = slides.findIndex(sl => /^Appendix ①/.test(sl.title)) + 1;
  const toc = divider.elements.filter(el => el.role === 'toc' && /^\d+$/.test(el.text)).map(el => Number(el.text));
  assert.strictEqual(toc[0], first, '目次の Appendix ① のページ番号（' + first + '）');
});

test('スライド：Appendix は入りきらなければ続きのページに分け、見出しの行を繰り返す（全件を出す）。中身が無くても ①〜③ を出し、表の見出しの行の下に理由を書く', () => {
  const {gas, g} = setup();
  const s = summarize(g, '新FMT');
  const many = Object.assign({}, s, {changes: Array.from({length: 40}, (_, i) => Object.assign({}, s.changes[1], {project: '案件' + i}))});
  const pages = plain(g.mtgSlides_(many)).filter(sl => /^Appendix ①/.test(sl.title));
  assert.ok(pages.length >= 3, pages.length + 'ページ');
  assert.strictEqual(pages[1].title, 'Appendix ①　前回から変化のある案件（一覧）（続き）');
  const tables = pages.map(p => p.elements.find(el => el.type === 'table'));
  tables.forEach(t => assert.strictEqual(t.rows[0].cells[0].text, '得意先／案件名', '見出しの行を繰り返す'));
  assert.strictEqual(tables.reduce((sum, t) => sum + t.rows.length - 1, 0), 40, '全件を出す（省かない）');

  const empty = Object.assign({}, s, {changes: [], stalled: [], overdue: [], alerts: []});
  const slides = plain(g.mtgSlides_(empty));
  assert.deepStrictEqual(slides.map(sl => sl.title), ['表紙', 'サマリー', 'パイプラインの変化', '前回から変化のある案件', '期待値上位案件の状況・課題',
    'Appendix', 'Appendix ①　前回から変化のある案件（一覧）', 'Appendix ②　停滞案件（一覧）', 'Appendix ③　入力内容の確認事項（責任者別）',
    'Appendix ③　入力内容の確認事項（一覧）'], '中身が無くても資料の型（Appendix ①〜③）は同じ');
  // 中身の無い Appendix：表の見出しの行を置き、その下に「該当なし」、作れないときは「データなし」と理由（記録の開始から45日たっていないので、停滞はまだ数えられない）
  const bodyOf = slide => [slide.elements.filter(el => el.type === 'table').map(t => t.rows.length), roleOf(slide, 'empty').concat(roleOf(slide, 'gap'))
    .map(el => [el.role, el.text.split('\n')[0]])];
  assert.deepStrictEqual(bodyOf(slides[6]), [[1], [['empty', '該当なし']]]);
  assert.deepStrictEqual(bodyOf(slides[7]), [[1], [['gap', 'データなし']]]);
  assert.match(roleOf(slides[7], 'gap')[0].text, /\n記録の開始（\d{1,2}\/\d{1,2}）から 45日未満のため、まだ数えられない（\d{1,2}\/\d{1,2} から数える）$/);
  assert.match(leadOf(slides[7]), /^停滞案件は、記録の開始（\d{1,2}\/\d{1,2}）から 45日未満のため、まだ数えられない（\d{1,2}\/\d{1,2} から数える）。$/);
  assert.deepStrictEqual(bodyOf(slides[9]), [[1], [['empty', '該当なし']]]);
  const gapBox = roleOf(slides[7], 'gap')[0];
  assert.deepStrictEqual([gapBox.fill, gapBox.line, gapBox.dash], [gas.get('MTG_OPTIONS').colors.gap, gas.get('MTG_OPTIONS').colors.gapLine, true], 'データなしは灰色の地に点線の枠');
  assert.deepStrictEqual(slides[8].elements.find(el => el.type === 'table').rows.map(r => r.cells[0].text), ['責任者', '合計'], '確認事項が無ければ合計の行だけ');
  assert.strictEqual(leadOf(slides[8]).indexOf('入力内容の確認事項は 0項目。'), 0);
  // 扉の一覧：中身の無い Appendix には「（該当なし）」、作れないものには「（データなし）」
  const toc = roleOf(slides[5], 'toc').map(el => el.text);
  assert.deepStrictEqual(toc.filter((_, i) => i % 3 === 1), ['前回から変化のある案件（該当なし）', '停滞案件（データなし）', '入力内容の確認事項（該当なし）']);
  const onlyAlerts = plain(g.mtgSlides_(Object.assign({}, empty, {alerts: s.alerts})));
  assert.deepStrictEqual(roleOf(onlyAlerts[5], 'toc').filter((_, i) => i % 3 === 1).map(el => el.text),
    ['前回から変化のある案件（該当なし）', '停滞案件（データなし）', '入力内容の確認事項'], '中身のある Appendix には付けない');
  const breakdown = slides[1].elements.find(el => el.type === 'table').rows.slice(1).map(r => r.cells.map(cell => cell.text));
  assert.ok(breakdown.every(r => r[1] === '0件' && r[3] === '該当なし'), '区分は同じまま、0件・該当なしで出す');
});

test('作れないところ（列・記録が無い）は、資料の型を変えずに、その場所に灰色の「データなし」と理由を書く。作り終えた画面にも一覧を出す', () => {
  const gas = createGas();
  const g = gas.global;
  // 得意先・案件名・案件ステータスだけのシート。差分追跡を設定していない（変更の記録が無い）・クレデンシャルの記録も無い
  gas.addSheet('新FMT', [['新FMT', '', ''], ['得意先', '案件名', '案件ステータス'], ['A社', 'アプリ刷新', '提案中'], ['B社', '店舗什器', '初回訪問']],
    {rows: 10, columns: 3});
  const S = gas.get('MTG_SLIDE');
  const c = gas.get('MTG_OPTIONS').colors;
  const s = summarize(g, '新FMT');
  assert.deepStrictEqual(Object.keys(plain(s.gaps)), ['changes', 'amount', 'weighted', 'orderMonth', 'activity', 'issue', 'owner', 'record', 'stalled']);
  assert.strictEqual(s.gaps.changes, '「変更履歴_差分」シートに変更の記録が無い');
  assert.strictEqual(s.gaps.amount, 'シートに「想定売上規模（百万）」の列が無い');
  assert.strictEqual(s.gaps.stalled, '変更の記録も「最終更新日時」の列も無く、ステータスが同じ日数が分からない');
  assert.strictEqual(s.kpi.missingAmount, 0, '列が無いのは空欄（未入力）ではない');
  assert.deepStrictEqual(Object.keys(plain(s.alertGaps)), plain(gas.get('MTG_ALERT_TYPES').filter(t => t !== '未入力')), '必須項目未入力のほかは、どれもチェックできない');
  const slides = plain(g.mtgSlides_(s));
  assert.strictEqual(slides.length, 10, '資料の型（表紙・本編4枚・扉・Appendix ①〜③）は同じ');

  // サマリー：前回比は「データなし」（±0 と書かない）、金額の箱は灰色、動きの行はどれも「データなし」（理由は見出し）、状態の行は理由つき
  const summary = slides[1];
  assert.match(leadOf(summary), /^進行中案件 2件、期待値合計はデータなし（シートに「期待値調整済想定売上規模（百万）」の列が無い）。直近14日間の増減はデータなし（「変更履歴_差分」シートに変更の記録が無い）。$/);
  assert.ok(!/±0/.test(textOf(summary)), '分からない増減を ±0 と書かない');
  assert.deepStrictEqual(roleOf(summary, 'tile-value').map(el => el.text), ['2件　直近14日間で データなし', 'データなし', 'データなし', '0項目']);
  assert.strictEqual(roleOf(summary, 'gap').length, 2, '想定売上・期待値の箱は灰色の点線');
  assert.deepStrictEqual(roleOf(summary, 'heading').map(el => el.text),
    ['直近14日間の動き（データなし：「変更履歴_差分」シートに変更の記録が無い）', '現在の状態（進行中案件）']);
  const [moves, states] = summary.elements.filter(el => el.type === 'table');
  assert.ok(moves.rows.slice(1).every(r => r.cells[1].text === '—' && r.cells[3].text === 'データなし' && r.cells[3].fill === c.gap));
  assert.deepStrictEqual(states.rows.map(r => r.cells[3].text), ['データなし：シートに「受注月」の列が無い',
    'データなし：変更の記録も「最終更新日時」の列も無く、ステータスが同じ日数が分からない', 'データなし：シートに「商談記録（ＵＲＬ）」の列が無い']);

  // パイプラインの変化：図・グラフの見出しは残し、その下に「データなし」と理由
  const pipeline = slides[2];
  assert.deepStrictEqual(roleOf(pipeline, 'gap').map(el => el.text.split('\n')), [
    ['データなし', 'シートに「期待値調整済想定売上規模（百万）」の列が無い'],
    ['データなし', 'シートに「期待値調整済想定売上規模（百万）」の列が無い'],
    ['データなし', 'シートに「期待値調整済想定売上規模（百万）」の列が無い']]);
  assert.ok(!pipeline.elements.some(el => /^bar/.test(el.role || '')), '棒は描かない');
  assert.strictEqual(leadOf(pipeline), '進行中の期待値はデータなし（シートに「期待値調整済想定売上規模（百万）」の列が無い）。');

  // 前回から変化のある案件：表の見出しの行と「データなし」と理由
  assert.deepStrictEqual(slides[3].elements.filter(el => el.type === 'table').map(t => t.rows.length), [1]);
  assert.deepStrictEqual(roleOf(slides[3], 'gap').map(el => el.text), ['データなし\n「変更履歴_差分」シートに変更の記録が無い']);
  assert.strictEqual(leadOf(slides[3]), '直近14日間に変化のある案件はデータなし（「変更履歴_差分」シートに変更の記録が無い）。');

  // 期待値上位：列の無い項目のセルは灰色の「データなし」（未入力と書かない）。並びは想定売上も無いので、シートの順
  const focus = slides[4].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(focus.rows[1].cells.map(cell => [cell.text, cell.fill === c.gap]), [
    ['A社\nアプリ刷新', false], ['データなし', true], ['提案中\n記録なし', false], ['データなし', true], ['データなし', true], ['データなし', true], ['該当なし', false]]);
  assert.match(leadOf(slides[4]), /^期待値はデータなし（シートに「期待値調整済想定売上規模（百万）」の列が無い）のため、進行中案件のうち想定売上上位 2件。$/);

  // 扉・Appendix：①・②は「データなし」、③は該当なし（チェックできなかった種類は灰色の列と理由）
  assert.deepStrictEqual(roleOf(slides[5], 'toc').filter((_, i) => i % 3 === 1).map(el => el.text),
    ['前回から変化のある案件（データなし）', '停滞案件（データなし）', '入力内容の確認事項（該当なし）']);
  assert.ok(roleOf(slides[8], 'gap')[0].text.indexOf('灰色の列は、チェックできなかった項目\n予定日超過：「クレデンシャル履歴_記録」シートが無い') > 0);

  // どの部品もページに入る
  slides.forEach((slide, i) => slide.elements.forEach(el => {
    const h = el.type === 'table' ? el.rows.reduce((a, r) => a + r.h, 0) : el.h;
    if (!slide.cover && !slide.divider && ['footer', 'page'].indexOf(el.role) < 0 && h > 1) assert.ok(el.y + h <= S.bottom + 0.01, (i + 1) + '枚目：' + (el.role || el.type));
  }));

  // 作れなかったところの一覧（同じ理由はまとめる）。作り終えた画面にも出す
  const list = plain(g.mtgGapList_(s));
  assert.ok(list.indexOf('前回からの変化：「変更履歴_差分」シートに変更の記録が無い') >= 0);
  assert.ok(list.indexOf('受注月・受注月超過・チェック「受注月超過」・チェック「売上開始月不整合」：シートに「受注月」の列が無い') >= 0);
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /\n　データなし：前回からの変化：「変更履歴_差分」シートに変更の記録が無い\n　　想定売上・チェック「金額不整合」：/);
});

test('スライド：入力内容のチェック（責任者別）は責任者 × 種類の件数の表だけ（種類の説明は置かない）。種類が多くてもページの幅に入る', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const s = summarize(g, '新FMT');
  const ownersOf = summary => plain(g.mtgSlides_(summary)).find(sl => sl.title === 'Appendix ③　入力内容の確認事項（責任者別）');
  const lead = leadOf;

  const few = ownersOf(s);
  assert.ok(!few.elements.some(el => el.text === 'チェックの種類'));
  assert.strictEqual(lead(few), '入力内容の確認事項は 3項目（3案件・責任者 3名）。チェックできなかった項目は 2種類（灰色の列）。');
  assert.deepStrictEqual(few.elements.find(el => el.type === 'table').rows.map(r => r.cells[0].text), ['責任者', '伊藤', '佐藤', '未入力', '合計'],
    '責任者は名前の順（日本語の順。件数の順ではない。責任者が空（未入力）は最後）');
  // 種類はいつもすべてを列に出す（資料の型を変えない）。チェックできなかった種類は灰色の列で、表の下に理由
  const c = gas.get('MTG_OPTIONS').colors;
  const types = gas.get('MTG_ALERT_TYPES');
  const fixed = few.elements.find(el => el.type === 'table');
  assert.strictEqual(fixed.columns.length, types.length + 2);
  assert.ok(Math.abs(fixed.columns.reduce((a, b) => a + b, 0) - S.w) < 0.01, '表は中身の幅いっぱい');
  const grayed = fixed.rows[0].cells.map((cell, i) => (cell.fill === c.gapLine ? cell.text.replace(/\n/g, '') : null)).filter(Boolean);
  assert.deepStrictEqual(grayed, ['予定日超過', '商談記録未登録']);
  assert.ok(fixed.rows.slice(1).every(r => r.cells[1].text === '—' && r.cells[1].fill === c.gap), 'チェックできなかった列は「—」');
  assert.deepStrictEqual(fixed.rows[fixed.rows.length - 1].cells.map(cell => cell.text),
    ['合計', '—', '1', '1', '—', '1', '0', '0', '0', '0', '3'], '合計の行は、チェックした種類の 0 も書く');
  const note = roleOf(few, 'gap')[0];
  assert.strictEqual(note.text, 'データなし\n灰色の列は、チェックできなかった項目\n予定日超過：「クレデンシャル履歴_記録」シートが無い\n' +
    '商談記録未登録：シートに「商談記録（ＵＲＬ）」の列が無い');
  assert.ok(note.y + note.h <= S.bottom, '理由の枠もページに入る');

  const alerts = types.map((type, i) => ({type, key: type + '|' + i, customer: 'X社', project: '案件' + i, text: '内容', amount: 1, owner: '担当' + (i % 3)}));
  const many = ownersOf(Object.assign({}, s, {alerts}));
  assert.ok(!many.elements.some(el => el.text === 'チェックの種類'));
  const table = many.elements.find(el => el.type === 'table');
  assert.strictEqual(table.columns.length, types.length + 2, '責任者・種類ごと・計');
  assert.ok(S.x + table.columns.reduce((a, b) => a + b, 0) <= S.x + S.w, '表がスライドの幅に入る');
  assert.ok(table.columns.every(w => w >= 32));
  table.rows[0].cells.forEach(cell => assert.ok(!/…/.test(cell.text), '見出しが切れない：' + cell.text));
});

test('パイプラインの変化：ステータス別・受注月別の期待値は横棒のグラフ（同じ目盛り。値は棒の先、件数は名前の右）', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const c = gas.get('MTG_OPTIONS').colors;
  const s = summarize(g, '新FMT');
  const slide = plain(g.mtgSlides_(s))[2];
  assert.ok(!slide.elements.some(el => el.type === 'table'), '表ではなくグラフ');
  const titles = slide.elements.filter(el => el.size === 10.5 && el.bold).map(el => el.text);
  assert.deepStrictEqual(titles, ['期待値の増減（進行中・百万円）', 'ステータス別期待値（進行中・百万円）', '受注月別期待値（進行中・百万円）'], '見出しは滝の図と同じ形');
  // 滝の図：期待値（前回・今）・増加・減少の色の凡例。受注は白抜き（進行中から外れたが、悪い減少ではない）
  assert.deepStrictEqual(roleOf(slide, 'legend').filter(el => el.text).map(el => el.text), ['期待値', '増加', '減少', '受注']);
  assert.deepStrictEqual(roleOf(slide, 'bar-total').map(el => el.fill), [c.greenDark, c.greenDark], '前回・今の期待値の棒');
  assert.deepStrictEqual(roleOf(slide, 'bar-lost').map(el => el.fill), [c.down], '減少（失注）の棒');
  roleOf(slide, 'legend').forEach(el => assert.ok(el.x + el.w <= S.x + S.w + 0.01, '凡例が右にはみ出さない：' + (el.text || '')));
  const wonSetup = setup({edits: false});
  wonSetup.g.recordDiffEdit(wonSetup.gas.edit(wonSetup.sheet, 'F7', '受注'));   // E社：受注（進行中の期待値 40 が外れる）
  const wonSlide = plain(wonSetup.g.mtgSlides_(summarize(wonSetup.g, '新FMT')))[2];
  const won = roleOf(wonSlide, 'bar-won');
  assert.deepStrictEqual(won.map(el => [el.fill, el.line]), [[c.white, c.greenDark]], '受注の棒は白抜きで緑の枠（減少の色にしない）');
  assert.match(leadOf(wonSlide), /減少 ▲40（うち受注 ▲40）/, '減少のうち受注の分をリード文に書く');
  const labels = roleOf(slide, 'bar-label').map(el => el.text);
  const values = roleOf(slide, 'bar-value');
  const bars = roleOf(slide, 'bar');
  const stages = plain(s.stages).filter(st => st.kind === 'active');
  assert.deepStrictEqual(labels, stages.map(st => st.status).concat(plain(s.quarters).map(q => q.label)), 'ステータスはプルダウンの順、受注月は時間の順');
  assert.deepStrictEqual(values.map(el => el.text), stages.map(st => g.mtgAmount_(st.weighted)).concat(plain(s.quarters).map(q => g.mtgAmount_(q.value))));
  assert.ok(bars.every(el => el.fill === c.greenDark), '棒は期待値の色（滝の図と同じ）');
  assert.strictEqual(new Set(bars.map(el => el.x)).size, 1, '棒は同じ線から');
  const unit = bars.map(el => el.w / Number(values.find(v => Math.abs(v.y + v.h / 2 - (el.y + el.h / 2)) < 0.01).text));
  assert.ok(unit.every(u => Math.abs(u - unit[0]) < 0.01), '2つのグラフは同じ目盛り');
  bars.forEach(el => {
    const value = values.find(v => Math.abs(v.y + v.h / 2 - (el.y + el.h / 2)) < 0.01);
    assert.ok(Math.abs(value.x + 3 - (el.x + el.w)) < 0.01, '値は棒の先');
  });
  slide.elements.filter(el => el.x >= S.x + 370 && el.y >= S.top - 2 && el.y < S.ruleY)
    .forEach(el => assert.ok(el.y + el.h <= S.bottom + 0.01 && el.x + el.w <= S.x + S.w + 0.01, 'はみ出さない：' + (el.text || el.role)));
});

test('サマリーの内訳：案件の名前は入る件数まで並べ、残りは「他 N件」（「…」で切らない）', () => {
  const {g} = setup({edits: false});
  const items = Array.from({length: 12}, (_, i) => ({customer: '得意先' + i, project: 'ネットスーパーの立ち上げ支援' + i, amount: 1}));
  const note = () => '100日以上';
  assert.strictEqual(g.mtgNames_(items, note).split('、').length, 4, '幅を渡さなければ mainRows.deals 件');
  const fit = g.mtgNames_(items, note, 300, 2);
  const lines = g.mtgWrap_(fit, 300, 9, 99);
  assert.ok(lines.length <= 2 && !/…$/.test(lines[lines.length - 1]), fit);
  const shown = fit.split('　他 ')[0].split('、').length;
  assert.ok(shown < 4 && new RegExp('　他 ' + (12 - shown) + '件$').test(fit), fit);
  assert.strictEqual(g.mtgNames_(items.slice(0, 1), note, 300, 2), '得意先0／ネットスーパーの立ち上げ支援0（100日以上）');
  assert.strictEqual(g.mtgNames_([], note, 300, 2), '該当なし');
});

test('ラベル：当てはまる状況が3つでも、行を高くしてセルに余白をあけて入れる（幅は列の中で同じ）', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const s = summarize(g, '新FMT');
  const focus = s.focus.map((d, i) => (i === 0 ? Object.assign({}, d, {stalled: true, overdue: true, slips: 2}) : d));
  const slide = plain(g.mtgSlides_(Object.assign({}, s, {focus})))[4];
  const table = slide.elements.find(el => el.type === 'table');
  const pills = roleOf(slide, 'pill');
  const row1 = pills.filter(el => el.y < table.y + table.rows[0].h + table.rows[1].h);
  assert.deepStrictEqual(row1.map(el => el.text), ['停滞', '受注月超過', '延期2回']);
  assert.ok(table.rows[1].h >= g.mtgPillsHeight_(3), '3つ並べても行に入る');
  const y0 = table.y + table.rows[0].h;
  assert.ok(row1[0].y >= y0 + S.pill.margin - 0.01 && row1[2].y + row1[2].h <= y0 + table.rows[1].h - S.pill.margin + 0.01, '上下に余白');
  assert.ok(row1.every((el, i) => !i || el.y - (row1[i - 1].y + row1[i - 1].h) === S.pill.gap), '間は同じ');
  assert.strictEqual(new Set(pills.map(el => el.w)).size, 1, '幅は列の中で同じ');
  assert.strictEqual(new Set(pills.map(el => el.x)).size, 1, '左右の中央にそろえる');
  assert.ok(table.y + table.rows.reduce((sum, r) => sum + r.h, 0) <= S.bottom, '表が下にはみ出さない');
});

test('文字の折り返し：「前 ▶ 今」は矢印のあとで、開きかっこは行の終わりに残さない。表のセルは行の長さをそろえる', () => {
  const {g} = setup({edits: false});
  assert.deepStrictEqual(plain(g.mtgWrap_('提案中 ▶ 課題ヒアリング', 118, 9, 2)), ['提案中 ▶', '課題ヒアリング']);
  g.mtgWrap_('青葉マート／需要予測による発注最適化（95日）、若葉フーズ／ネットスーパーの立ち上げ支援（52日）', 200, 9, 9)
    .forEach(line => assert.ok(!/^\d*日/.test(line) && !/（\d*$/.test(line), '数字の並びの途中・かっこの直後で切らない：' + line));
  for (let w = 120; w <= 300; w += 4) {
    g.mtgWrap_('港町GMS／PB商品のパッケージ刷新（100日以上）、さくらスーパー／店舗DXの全体構想（75日）', w, 9, 9)
      .forEach(line => assert.ok(!/[（「]$/.test(line), w + 'pt：' + line));
  }
  assert.deepStrictEqual(plain(g.mtgWrap_('決裁者（社長）に会えていない', 134, 9, 3)).map(t => t.length), [13, 1]);
  assert.deepStrictEqual(plain(g.mtgBalance_('決裁者（社長）に会えていない', 134, 9, 3)), ['決裁者（社長）', 'に会えていない'],
    '最後の行に1文字だけ残さない（区切りは言葉の切れ目に）');
  assert.deepStrictEqual(plain(g.mtgBalance_('期待値・受注月', 70, 8.5, 2)), ['期待値・', '受注月']);
  assert.deepStrictEqual(plain(g.mtgBalance_('5/29離反シナリオ紹介。', 114, 9, 3)), ['5/29離反', 'シナリオ紹介。'], 'カタカナの語の途中で切らない');
  assert.deepStrictEqual(plain(g.mtgBalance_('ユニーの販促全体サポート提案中', 114, 9, 3)), ['ユニーの販促全体', 'サポート提案中']);
  assert.deepStrictEqual(plain(g.mtgBalance_('9/1初回訪問。9/15課題ヒアリング。9/29提案書の方向性確認', 114, 9, 3)),
    ['9/1初回訪問。9/15', '課題ヒアリング。9/29', '提案書の方向性確認'], '3行でも、それぞれの区切りを言葉の切れ目に');
  assert.deepStrictEqual(plain(g.mtgBalance_('短い', 142, 9, 3)), ['短い'], '1行に入るものはそのまま');
  const dated = '10/13 価格の再提示';
  assert.strictEqual(g.mtgWrap_(dated, 98, 8.5, 2).length, 2);
  assert.deepStrictEqual(plain(g.mtgBalance_(dated, 98, 8.5, 2)), ['10/13', '価格の再提示'], '空白で分ければ1行ずつに入るなら、言葉の途中で切らずに空白で分ける');
  const para = '東西マート／物流センターの改善（失注）、富士ストア／店舗アプリ改修（提案中→課題ヒアリング）、海風マート／来店分析（削除）';
  assert.deepStrictEqual(plain(g.mtgBalance_(para, 360, 9, 3)), plain(g.mtgWrap_(para, 360, 9, 3)), '最後の行が十分に長く、区切りが言葉の切れ目なら、そのまま（行をいっぱいに使う）');
  assert.deepStrictEqual(plain(g.mtgBalance_(para, 400, 9, 3)).map(line => line.slice(-4)), ['中→課題', '（削除）'], '「ヒアリ／ング」ではなく「課題／ヒアリング」');
  const long = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも';
  assert.deepStrictEqual(plain(g.mtgBalance_(long, 100, 9, 2)), plain(g.mtgWrap_(long, 100, 9, 2)), '「…」で切れるものはそのまま');
});

test('文字の折り返し：見積もりの幅で分け、最後の行は「…」。行の頭に来ない文字はぶら下げ、半角の語は途中で切らない', () => {
  const {g} = setup({edits: false});
  assert.deepStrictEqual(plain(g.mtgWrap_('あいうえおかきくけこ', 92, 9, 3)), ['あいうえおかきく', 'けこ']);
  assert.deepStrictEqual(plain(g.mtgWrap_('あいうえおかきくけこさしすせそたちつてと', 92, 9, 2)), ['あいうえおかきく', 'けこさしすせそ…'],
    '最後の行は「…」の分をあけて切る');
  assert.deepStrictEqual(plain(g.mtgWrap_('あいうえおかきく。', 92, 9, 3)), ['あいうえおかきく。'], '本当の幅に収まる「。」はぶら下げる');
  assert.deepStrictEqual(plain(g.mtgWrap_('期待値（前▶今）', 66, 7.5, 3)), ['期待値（前▶', '今）'], '収まらなければ前の文字と一緒に送る');
  assert.deepStrictEqual(plain(g.mtgWrap_('2026/11 ▶ 2026/12', 80, 7.5, 2)), ['2026/11 ▶', '2026/12']);
  assert.deepStrictEqual(plain(g.mtgWrap_('一行目\n二行目', 300, 9, 2)), ['一行目', '二行目']);
  assert.strictEqual(g.mtgWrap_('ABCDEFGHIJKLMNOPQRSTUVWXYZ', 150, 8, 3).length, 2, '英大文字は幅を広めに見積もる（BIZ UDPゴシックでは 0.44〜1）');
  // Google スライドで書き出した PDF で測った幅（9pt）：「2026/9」38.7pt・「B-クロージング中」70.6pt・「シナリオ提案」50.3pt。見積もりはそれ以上
  [['2026/9', 38.7], ['B-クロージング中', 70.6], ['シナリオ提案', 50.3]].forEach(([text, real]) => {
    const est = g.mtgTextWidth_(text, 9);
    assert.ok(est >= real - 0.1 && est <= real * 1.06, text + '：' + est);
  });
  assert.strictEqual(g.mtgWrap_('2026/12', 62, 9, 2).length, 1, '受注月の列（62pt）に「2026/12」が1行で入る');
  assert.strictEqual(g.mtgWrap_('2026/12', 56, 9, 2).length, 2, '56pt では入らない（Google スライドでも「2026/1」「2」に折り返された）');
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
  assert.deepStrictEqual([style.fontFamily, style.fontSize.magnitude, style.bold], ['BIZ UDPGothic', 10, true], 'BIZ UDPゴシック');
  const borders = requests.filter(r => r.updateTableBorderProperties).map(r => r.updateTableBorderProperties);
  assert.deepStrictEqual(borders.map(b => [b.borderPosition, b.tableRange.location.rowIndex, b.tableBorderProperties.weight.magnitude]),
    [['ALL', 0, 0.75], ['OUTER', 1, 2.25]], '区切り線は薄いグレー（0.75pt）');
});

test('mtgSlideRequests_：一部の文字の書式（spans）・塗りの濃さ（alpha）・倍率（scale）・ひな形のページ（existing）とひな形が無いときの表紙', () => {
  const {g} = setup({edits: false});
  const slide = {elements: [
    {type: 'box', x: 10, y: 10, w: 200, h: 20, text: '今期（2026/4〜）　受注', size: 9, spans: [{start: 0, end: 10, bold: true}, {start: 12, end: 40, color: '#595959', size: 8}]},
    {type: 'box', x: 300, y: 300, w: 360, h: 0.75, fill: '#FFFFFF', alpha: 0.5, shape: 'RECTANGLE'}
  ]};
  const requests = plain(g.mtgSlideRequests_(slide, 'mtg_p1', 0));
  const spans = requests.filter(r => r.updateTextStyle && r.updateTextStyle.textRange.type === 'FIXED_RANGE').map(r => r.updateTextStyle);
  assert.deepStrictEqual(spans.map(r => [r.textRange.startIndex, r.textRange.endIndex, r.fields]), [[0, 10, 'bold'], [12, 14, 'foregroundColor,fontSize']],
    '文字数を超える範囲は、文字の終わりまで');
  const alpha = requests.filter(r => r.updateShapeProperties).pop().updateShapeProperties.shapeProperties.shapeBackgroundFill.solidFill;
  assert.strictEqual(alpha.alpha, 0.5);

  // 改行で分けた部分の書式は、折り返したあとの位置になる
  const row = plain(g.mtgRow_([60], [{text: 'とても長い得意先の名前\n案件', parts: [{bold: true}, {color: '#595959'}], lines: 4, size: 9}]));
  const text = row.cells[0].text;
  assert.ok(text.split('\n').length >= 3, text);
  assert.deepStrictEqual(row.cells[0].spans.map(sp => text.slice(sp.start, sp.end).replace(/\n/g, '')), ['とても長い得意先の名前', '案件']);

  // 倍率：位置・大きさ・文字・線の太さを掛ける（ひな形のページ 960 × 540pt なら 4/3）
  const big = plain(g.mtgSlideRequests_(slide, 'mtg_p1', 0, {scale: 4 / 3}));
  const shape = big.find(r => r.createShape).createShape.elementProperties;
  assert.deepStrictEqual([shape.size.width.magnitude, shape.size.height.magnitude, shape.transform.translateX, shape.transform.translateY],
    [266.67, 26.67, 13.33, 13.33]);
  const fonts = big.filter(r => r.updateTextStyle).map(r => r.updateTextStyle.style.fontSize && r.updateTextStyle.style.fontSize.magnitude);
  assert.deepStrictEqual(fonts, [12, undefined, 10.67]);

  // ひな形のページ：スライドは作らず、足す部品だけ（決まった文字の書き換えは mtgTemplateTextRequests_）。部品の ID は idPrefix から
  const cover = {template: {title: 'スーパー・GMS', subtitle: 'RX Business MTG　事前資料', date: '2026年11月6日（金）', name: ''}, elements: [slide.elements[0]]};
  const onTemplate = plain(g.mtgSlideRequests_(cover, 'g1a2b3', 0, {scale: 4 / 3, existing: true, idPrefix: 'mtg_p1'}));
  assert.ok(!onTemplate.some(r => r.createSlide || r.replaceAllText || r.updatePageProperties));
  const added = onTemplate.find(r => r.createShape).createShape;
  assert.deepStrictEqual([added.objectId, added.elementProperties.pageObjectId], ['mtg_p1_0', 'g1a2b3']);
  // ひな形が無いとき：ページの背景を緑一色にし（あとで画像にする）、ひな形と同じ位置で白い字（名前の欄は空なので描かない）
  const plainCover = plain(g.mtgSlideRequests_(cover, 'mtg_p1', 0));
  assert.ok(plainCover[0].createSlide);
  assert.deepStrictEqual(plainCover[1], {updatePageProperties: {objectId: 'mtg_p1', pageProperties: {pageBackgroundFill: {solidFill: {color: {rgbColor:
    plain(g.mtgRgb_('#3A7D22'))}}}}, fields: 'pageBackgroundFill.solidFill.color'}});
  assert.ok(!plainCover.some(r => r.createShape && r.createShape.elementProperties.size.width.magnitude >= 720), 'ページ全体を覆う四角は置かない（背景の画像が隠れる）');
  assert.deepStrictEqual(plainCover.filter(r => r.insertText).map(r => r.insertText.text),
    ['スーパー・GMS', 'RX Business MTG　事前資料', '2026年11月6日（金）', '今期（2026/4〜）　受注']);
});

test('表紙・扉の背景の画像：1280 × 720 の PNG（表紙は緑のグラデーションに濃い緑の RX、扉はグレー）。正しく展開でき、小さい', () => {
  const zlib = require('zlib');
  const {g} = setup({edits: false});
  const png = Buffer.from(g.mtgCoverPng_());
  assert.ok(png.length < 120000, '大きさ：' + png.length);
  assert.deepStrictEqual(Array.from(png.slice(0, 8)), [137, 80, 78, 71, 13, 10, 26, 10]);
  // チャンクの CRC を確かめながら読む
  const crc = buf => {
    let c = 0xFFFFFFFF;
    for (const b of buf) {
      c ^= b;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  };
  const chunks = {};
  for (let at = 8; at < png.length;) {
    const length = png.readUInt32BE(at);
    const type = png.slice(at + 4, at + 8).toString('latin1');
    const data = png.slice(at + 8, at + 8 + length);
    assert.strictEqual(png.readUInt32BE(at + 8 + length), crc(png.slice(at + 4, at + 8 + length)), type + ' の CRC');
    chunks[type] = data;
    at += 12 + length;
  }
  assert.deepStrictEqual(Object.keys(chunks), ['IHDR', 'IDAT', 'IEND']);
  assert.deepStrictEqual([chunks.IHDR.readUInt32BE(0), chunks.IHDR.readUInt32BE(4), chunks.IHDR[8], chunks.IHDR[9]], [1280, 720, 8, 2]);
  const raw = zlib.inflateSync(chunks.IDAT);   // zlib の形式（Adler-32 も確かめる）
  const stride = 1280 * 3 + 1;
  assert.strictEqual(raw.length, stride * 720);
  const pixel = (x, y) => '#' + Array.from(raw.slice(y * stride + 1 + x * 3, y * stride + 4 + x * 3)).map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  assert.strictEqual(pixel(60, 400), '#275316', 'R の縦棒は濃い緑');
  assert.strictEqual(pixel(300, 200), '#3A7D22', 'R の穴は左上の緑（グラデーションの始まり）');
  assert.strictEqual(pixel(1279, 719), '#275316', '右下の角は X の脚');
  assert.ok(/^#[0-9A-F]{6}$/.test(pixel(1000, 400)) && pixel(1000, 400) !== '#3A7D22', '右のほうは明るくなる');
  for (let y = 0; y < 720; y++) assert.strictEqual(raw[y * stride], 0, '行の頭はフィルタなし');
  // 扉：同じ絵柄のグレー
  const gray = zlib.inflateSync(Buffer.from(g.mtgCoverPng_('divider')).slice(41, -12));
  const grayPixel = (x, y) => '#' + Array.from(gray.slice(y * stride + 1 + x * 3, y * stride + 4 + x * 3)).map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  assert.deepStrictEqual([grayPixel(60, 400), grayPixel(300, 200), grayPixel(1279, 719)], ['#A6A6A6', '#7F7F7F', '#A6A6A6'],
    'いただいたひな形の2枚目と同じ色（RX は背景より明るいグレー）');
  let colored = 0;
  let light = 0;
  for (let y = 0; y < 720; y++) {
    for (let x = 0; x < 1280; x++) {
      const at = y * stride + 1 + x * 3;
      if (gray[at] !== gray[at + 1] || gray[at + 1] !== gray[at + 2]) colored++;
      light = Math.max(light, gray[at]);
    }
  }
  assert.deepStrictEqual([colored, light >= 0xC0], [0, true], 'どの画素もグレーで、右下へ明るくなる（いちばん明るい ' + light.toString(16) + '）');
  // 統合版の表紙：紫（ひな形の3枚目）。RX は 33pt 左（R の縦棒が左に寄る）
  const purple = zlib.inflateSync(Buffer.from(g.mtgCoverPng_('combined')).slice(41, -12));
  const purplePixel = (x, y) => '#' + Array.from(purple.slice(y * stride + 1 + x * 3, y * stride + 4 + x * 3)).map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  assert.deepStrictEqual([purplePixel(20, 400), purplePixel(300, 200)], ['#421C5E', '#7030A0']);
  const stemEnd = at => { let x = 0; while (at(x, 400) === at(0, 400)) x++; return x; };   // R の縦棒の右端（y = 400）
  const shift = stemEnd(pixel) - stemEnd(purplePixel);
  assert.ok(shift >= 43 && shift <= 46, 'RX は 33pt（1280px の画像で 44px）左：' + shift);
});

test('表紙・扉の背景の画像を付けられなくても、資料は作る（一色の背景のまま。ログに残す）', () => {
  const {gas, g} = setup();
  gas.slidesAppError = 'SlidesApp が使えません';
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.strictEqual(gas.slideBackgrounds.length, 0);
  assert.strictEqual(gas.driveFiles.filter(f => /\.pdf$/.test(f.name)).length, 1, 'PDF は作る');
  assert.ok(gas.logs.some(l => l.level === 'warn' && /表紙・扉の背景の画像を付けられませんでした/.test(l.line)));
  assert.ok(gas.decks[0].requests.some(r => r.updatePageProperties), '背景は緑一色');
});

test('クレデンシャル・オファリング履歴を使っていない間（CRED_OPTIONS.enabled: false）は、入力内容のチェックに「予定日超過」を出さない（列も作らない）', () => {
  const {gas, g} = setup({edits: false});
  const cred = gas.get('CRED_OPTIONS');
  const {log, cols} = g.credLogSheet_(gas.ss, false);
  const fields = {kind: cred.linking.childKind, name: '販促DXの試験導入のご提案', person: '佐藤 花子', title: '販促企画部 課長', owner: '鈴木'};
  fields[cred.schedule.planKey] = day(-3);
  fields[cred.schedule.doneKey] = '';
  g.credAppendLog_(log, cols, [{id: 'cred-real-0001', at: new gas.Date(), customer: 'A社', fields: g.credValidate_(fields, null), link: '',
    editor: 'suzuki@example.com', source: 'テスト'}]);
  assert.ok(summarize(g, '新FMT').alerts.some(a => a.type === '予定日超過'), '使っているときは出る');
  cred.enabled = false;
  gas.ss.deleteSheet(gas.ss.getSheetByName(cred.logSheet));
  const s = summarize(g, '新FMT');
  assert.ok(!s.alerts.some(a => a.type === '予定日超過'));
  assert.ok(!('予定日超過' in s.alertGaps), 'クレデンシャルのシートが無くても「データなし」にしない');
  assert.deepStrictEqual(plain(g.mtgAlertTypes_()), plain(gas.get('MTG_ALERT_TYPES').filter(t => t !== '予定日超過')));
  const owners = g.mtgAppendixOwners_(s)[0];
  const table = owners.elements.find(el => el.type === 'table');
  assert.strictEqual(table.columns.length, gas.get('MTG_ALERT_TYPES').length - 1 + 2);
  assert.ok(!table.rows[0].cells.some(cell => /予定日/.test(cell.text)));
});

/* ---------------- 会議のフォルダへの格納 ---------------- */

/** 新FMT のサブインダストリーのフォルダ（会議のフォルダ folders を中に置く）を用意し、DIFF_RULES の meetingFolder に設定する */
function meetingFolders(gas, folders) {
  const parent = 'SubIndustryFolder0001';
  gas.driveFolders.push({id: parent, name: 'スーパー・GMS', parent: 'root'});
  folders.forEach((name, i) => gas.driveFolders.push({id: 'meeting-folder-' + i, name, parent}));
  gas.get('DIFF_RULES')['新FMT'].meetingFolder = 'https://drive.google.com/drive/folders/' + parent + '?usp=sharing';
  return parent;
}

/** 会議日 'yyyy/MM/dd' の yyMMdd */
const yymmdd = date => date.slice(2, 4) + date.slice(5, 7) + date.slice(8, 10);

test('本番：会議日のフォルダ（yyMMdd_第N回会議）に事前資料の PDF を移して格納する（保存先と二重にしない）。作り直したら古い版はゴミ箱へ。試作は格納しない', () => {
  const {gas, g} = setup();
  const meeting = g.mtgParseDate_('10/23', new Date());
  const name = meeting.replace(/\//g, '') + '_新FMT_RX Business Meeting用資料.pdf';
  meetingFolders(gas, [yymmdd(meeting) + '_第5回会議', '260819_第4回会議', yymmdd(meeting) + '0_別の会議']);

  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.ok(!gas.driveFiles.some(f => f.folder === 'meeting-folder-0'), '試作は格納しない');

  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  const stored = gas.driveFiles.filter(f => f.folder === 'meeting-folder-0');
  assert.deepStrictEqual(stored.map(f => f.name), [name]);
  assert.ok(!gas.driveFiles.some(f => f.folder === 'meeting-folder-1' || f.folder === 'meeting-folder-2'), 'ほかの会議のフォルダには置かない');
  assert.match(gas.alerts[gas.alerts.length - 1].message, new RegExp('会議のフォルダ「' + yymmdd(meeting) + '_第5回会議」に格納しました：https://drive.google.com/drive/folders/meeting-folder-0'));
  assert.strictEqual(gas.driveFiles.filter(f => f.name === name && gas.trashed.indexOf(f.id) < 0).length, 1, 'PDF は会議のフォルダの1つだけ（保存先には残さない）');

  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  const again = gas.driveFiles.filter(f => f.folder === 'meeting-folder-0');
  assert.strictEqual(again.length, 2);
  assert.ok(gas.trashed.indexOf(stored[0].id) >= 0, '会議のフォルダの古い版はゴミ箱へ');
  assert.ok(gas.trashed.indexOf(again[1].id) < 0);
});

test('会議のフォルダを作らない設定（createMeetingFolder: false）で無いときは、保存先にだけ残して知らせ、フォルダを用意してから storeMeetingSummaries() で格納する', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').createMeetingFolder = false;
  const meeting = g.mtgParseDate_('10/23', new Date());
  meetingFolders(gas, ['260819_第4回会議']);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, new RegExp('※ 会議のフォルダに格納できませんでした（「スーパー・GMS」の中に「' + yymmdd(meeting) +
    '_」で始まる会議のフォルダが無い）。資料は保存先にあります。フォルダを用意してから、エディタで storeMeetingSummaries\\(\\) を実行して格納してください。'));
  assert.ok(JSON.parse(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT')).some(m => m.date === meeting), '会議日は記録する');

  gas.driveFolders.push({id: 'meeting-folder-new', name: yymmdd(meeting) + '_第5回会議', parent: 'SubIndustryFolder0001'});
  gas.promptAnswers = ['1', '10/23'];
  g.storeMeetingSummaries();
  assert.deepStrictEqual(gas.driveFiles.filter(f => f.folder === 'meeting-folder-new').map(f => f.name), [meeting.replace(/\//g, '') + '_新FMT_RX Business Meeting用資料.pdf']);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：会議のフォルダ「\d{6}_第5回会議」に格納しました/);

  gas.promptAnswers = ['1', '10/30'];
  g.storeMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：本番の資料（\d{4}1030_新FMT_RX Business Meeting用資料\.pdf）が保存先・会議のフォルダにありません。createMeetingSummaries\(\) で作ってください。/);
});

test('会議のフォルダが無ければ「yyMMdd_第N回会議」を作って格納する（N は前の会議のフォルダの回数＋1。後の会議のフォルダは数えない）', () => {
  const {gas, g} = setup();
  const meeting = g.mtgParseDate_('10/23', new Date());
  const later = String(Number(meeting.slice(2, 4)) + 1) + '0101_第9回会議';   // 来年の会議（数えない）
  meetingFolders(gas, ['260420_第1回会議', '260819_第４回会議', later, '資料置き場']);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  const created = gas.driveFolders.filter(f => f.parent === 'SubIndustryFolder0001' && f.name.indexOf(yymmdd(meeting)) === 0);
  assert.deepStrictEqual(created.map(f => f.name), [yymmdd(meeting) + '_第5回会議'], '全角の数字も読む');
  assert.ok(gas.driveFiles.some(f => f.folder === created[0].id && /_新FMT_RX Business Meeting用資料\.pdf$/.test(f.name)));
  assert.match(gas.alerts[gas.alerts.length - 1].message, new RegExp('会議のフォルダ「' + yymmdd(meeting) + '_第5回会議」を作って格納しました'));
  // 作り直しても、同じフォルダを使う（もう1つ作らない）
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.strictEqual(gas.driveFolders.filter(f => f.name.indexOf(yymmdd(meeting)) === 0).length, 1);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /会議のフォルダ「\d{6}_第5回会議」に格納しました/);
  assert.strictEqual(g.mtgNextMeetingNumber_(g.DriveApp.getFolderById('SubIndustryFolder0001'), '260101'), 1, '前の会議が無ければ第1回');
});

test('mtgMeetingFolder_：設定が無い・開けない・2つある・コピーできないときは理由を返す（止めない）。URL・ID のどちらでも設定できる', () => {
  const {gas, g} = setup({edits: false});
  assert.match(g.mtgMeetingFolder_('新FMT', '2026/10/23').error, /「新FMT」の会議のフォルダが設定されていない（test\.gs の DIFF_RULES の meetingFolder）/);
  meetingFolders(gas, ['261023_第5回会議', '261023_第5回会議（予備）']);
  assert.match(g.mtgMeetingFolder_('新FMT', '2026/10/23').error, /「261023_」で始まるフォルダが 2個ある（「261023_第5回会議」「261023_第5回会議（予備）」。1つにしてください）/);
  gas.driveFolders.pop();
  assert.strictEqual(g.mtgMeetingFolder_('新FMT', '2026/10/23').folder.getId(), 'meeting-folder-0');
  gas.get('DIFF_RULES')['新FMT'].meetingFolder = 'SubIndustryFolder0001';
  assert.strictEqual(g.mtgMeetingFolder_('新FMT', '2026/10/23').folder.getId(), 'meeting-folder-0', 'ID だけでもよい');
  gas.missingFolders = ['SubIndustryFolder0001'];
  assert.match(g.mtgMeetingFolder_('新FMT', '2026/10/23').error, /フォルダ（ID：SubIndustryFolder0001）を開けない/);
  gas.missingFolders = [];
  gas.copyErrors = ['meeting-folder-0'];
  const pdf = g.DriveApp.getFolderById('x').createFile({getName: () => '20261023_新FMT_RX Business Meeting用資料.pdf'});
  assert.match(g.mtgStoreToMeetingFolder_('新FMT', '2026/10/23', pdf).error, /「261023_第5回会議」に移せない：Access denied/);
  assert.deepStrictEqual(['https://drive.google.com/drive/folders/AbCdEf123456?usp=drive_link', 'https://drive.google.com/open?id=AbCdEf123456',
    'AbCdEf123456', ''].map(g.mtgFolderIdOf_), ['AbCdEf123456', 'AbCdEf123456', 'AbCdEf123456', '']);
});

test('インデックス：会議のフォルダに格納したら、「INDEX_会議資料」シートの行＝会議日・列＝サブインダストリーのセルに PDF のファイルを書く（Sheets API が無ければファイル名のリンク）', () => {
  const {gas, g, sheet: source} = setup();
  gas.select(source, 'D3');
  const meeting = g.mtgParseDate_('10/23', new Date());
  meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();

  const sheet = gas.ss.getSheetByName('INDEX_会議資料');
  assert.ok(sheet, 'シートが無ければ作る');
  assert.strictEqual(gas.ss.getActiveSheet().getName(), '新FMT', '開いていたシートに戻す');
  const pdf = gas.driveFiles.find(f => f.folder === 'meeting-folder-0');
  const day = d => d.replace(/\/0?/g, '/');   // 2026/10/09 → 2026/10/9（表示形式 yyyy/m/d）
  assert.deepStrictEqual(plain(sheet.getRange(1, 1, 2, 5).getDisplayValues()), [
    ['日付', '種別', '全サブインダストリー', '新FMT', '新FMT2'],
    [day(meeting), 'RX Business MTG', '', meeting.replace(/\//g, '') + '_新FMT_RX Business Meeting用資料.pdf', '']
  ]);
  assert.deepStrictEqual([sheet.getMaxRows(), sheet.getMaxColumns()], [2, 5], '使わない行・列は持たない');
  assert.strictEqual(sheet.getRange(2, 4).getRichTextValues()[0][0].getLinkUrl(), 'https://drive.google.com/file/d/' + pdf.id + '/view', 'ファイルへの直リンク');
  assert.strictEqual(Object.prototype.toString.call(sheet.getRange(2, 1).getValue()), '[object Date]', '会議日は日付の値');
  assert.deepStrictEqual([sheet.getFrozenRows(), sheet.tabColor, sheet.widths[1], sheet.widths[3], sheet.widths[4], sheet.heights[2]], [1, '#000000', 83, 158, 138, 38]);
  assert.deepStrictEqual(plain(sheet.cell(2, 4).fmt), {fw: 'bold', fs: 9, fc: '#000000', va: 'middle'}, '太字・9pt・黒・上下中央');
  assert.ok(gas.writes.filter(w => w.sheet === 'INDEX_会議資料').every(w => w.locked), '書き込みはロックの中');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「INDEX_会議資料」シートに、資料のファイルを書きました：https:\/\/docs\.google\.com\/spreadsheets\/d\/spreadsheet-id\/edit#gid=\d+\n（スマートチップにできなかったので、ファイル名のリンクで書きました。チップにするには、エディタの「サービス」に Google Sheets API を追加してください）/);

  // 会議日は日付の順（前の会議日はその前に差し込む）。同じセルは上書き。ほかのサブインダストリーは右に足す。日付の値で入力した会議日も読む
  const at = url => ({folder: {getUrl: () => 'folder', getName: () => 'folder'}, file: {getUrl: () => url, getName: () => url}});
  const year = meeting.slice(0, 4);
  assert.ok(g.mtgWriteIndex_('新FMT', year + '/10/30', at('u-1030')).url);
  assert.ok(g.mtgWriteIndex_('新FMT2', year + '/10/16', at('u-1016')).url);
  assert.ok(g.mtgWriteIndex_('新FMT', meeting, at('u-1023b')).url);
  assert.ok(g.mtgWriteIndex_('専門店', year + '/10/30', at('u-x')).url);
  assert.deepStrictEqual(plain(sheet.getRange(1, 1, 4, 6).getDisplayValues()), [
    ['日付', '種別', '全サブインダストリー', '新FMT', '新FMT2', '専門店'],
    [year + '/10/16', 'RX Business MTG', '', '', 'u-1016', ''],
    [day(meeting), 'RX Business MTG', '', 'u-1023b', '', ''],
    [year + '/10/30', 'RX Business MTG', '', 'u-1030', '', 'u-x']
  ]);
  assert.deepStrictEqual(plain(sheet.cell(2, 1).fmt), {fw: 'bold', fs: 9, fc: '#000000', va: 'middle'}, '差し込んだ行も同じ書式');
  // 手で入れた行：日付の値・種別が空なら同じ行に書く。種別が違う行には書かない（新しい行にする）
  gas.asUser(() => sheet.getRange(4, 1, 1, 3).setValues([[new gas.Date(Number(year), 9, 30), '', '全体の資料']]));
  g.mtgWriteIndex_('新FMT2', year + '/10/30', at('u-1030b'));
  assert.deepStrictEqual(plain(sheet.getRange(4, 3, 1, 3).getDisplayValues()[0]), ['全体の資料', 'u-1030', 'u-1030b'], '手で入れた値は残す');
  gas.asUser(() => sheet.getRange(4, 2).setValue('別の会議'));
  g.mtgWriteIndex_('新FMT2', year + '/10/30', at('u-1030c'));
  assert.deepStrictEqual(plain(sheet.getRange(4, 1, 2, 5).getDisplayValues()), [
    [year + '/10/30', '別の会議', '全体の資料', 'u-1030', 'u-1030b'],
    [year + '/10/30', 'RX Business MTG', '', '', 'u-1030c']
  ]);
  assert.strictEqual(sheet.getLastColumn(), 6, '列は増やさない');
});

test('インデックス：前の形式（行＝サブインダストリー）のシートには書かずに知らせ、rebuildMeetingIndex() で会議のフォルダの資料から今の形式に作り直す', () => {
  const {gas, g} = setup({edits: false});
  const year = g.mtgParseDate_('10/23', new Date()).slice(0, 4);
  const old = gas.addSheet('INDEX_会議資料', [['サブインダストリー', year + '/10/23'], ['新FMT', '前のチップ']], {rows: 2, columns: 2});
  const at = {folder: {getUrl: () => 'f', getName: () => 'f'}, file: {getUrl: () => 'u', getName: () => 'u'}};
  const r = g.mtgWriteIndex_('新FMT', year + '/10/23', at);
  assert.match(r.error, /「INDEX_会議資料」シートに書き込めませんでした（1行目に「日付」の見出しがありません）。前の形式（行がサブインダストリー）のシートは、エディタで rebuildMeetingIndex\(\) を実行すると、今の形式で作り直します。/);
  assert.strictEqual(old.getRange(2, 2).getValue(), '前のチップ', '前の形式のシートは書き換えない');

  // 会議のフォルダ：新FMT は 10/16・10/23（10/30 は資料なし）、新FMT2 は設定なし
  gas.get('DIFF_RULES')['新FMT'].meetingFolder = 'https://drive.google.com/drive/folders/SubIndustryFolder0001';
  gas.driveFolders.push({id: 'm1023', name: year.slice(2) + '1023_第2回会議', parent: 'SubIndustryFolder0001'},
    {id: 'm1016', name: year.slice(2) + '1016_第1回会議', parent: 'SubIndustryFolder0001'},
    {id: 'm1030', name: year.slice(2) + '1030_第3回会議', parent: 'SubIndustryFolder0001'},
    {id: 'other', name: '資料置き場', parent: 'SubIndustryFolder0001'});
  const pdf = (folder, date) => gas.driveFiles.push({id: 'pdf-' + folder, folder, name: date + '_新FMT_RX Business Meeting用資料.pdf', blob: {}});
  pdf('m1023', year + '1023');
  pdf('m1016', year + '1016');
  gas.driveFiles.push({id: 'draft', folder: 'm1030', name: year + '1030_新FMT_RX Business Meeting用資料_試作.pdf', blob: {}});

  gas.confirmAnswer = 'NO';
  g.rebuildMeetingIndex();
  assert.match(gas.alerts[0].message, /今のシートは前の形式（行がサブインダストリー）なので、中身を消して、行＝会議日・列＝サブインダストリーの形式で作り直します。/);
  assert.strictEqual(old.getRange(2, 2).getValue(), '前のチップ', '「いいえ」なら何もしない');

  gas.confirmAnswer = 'YES';
  g.rebuildMeetingIndex();
  const sheet = gas.ss.getSheetByName('INDEX_会議資料');
  assert.deepStrictEqual(plain(sheet.getRange(1, 1, 3, 5).getDisplayValues()), [
    ['日付', '種別', '全サブインダストリー', '新FMT', '新FMT2'],
    [year + '/10/16', 'RX Business MTG', '', year + '1016_新FMT_RX Business Meeting用資料.pdf', ''],
    [year + '/10/23', 'RX Business MTG', '', year + '1023_新FMT_RX Business Meeting用資料.pdf', '']
  ]);
  assert.strictEqual(sheet.getLastRow(), 3, '試作の資料・会議日でないフォルダは書かない');
  assert.ok(gas.writes.filter(w => w.sheet === 'INDEX_会議資料').every(w => w.locked), '書き込みはロックの中');
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /「INDEX_会議資料」シートに、資料のファイルを 2件書きました。/);
  assert.match(done, new RegExp('・新FMT：2件（' + year + '/10/16・' + year + '/10/23）\n・新FMT2：会議のフォルダが設定されていない'));

  // 今の形式なら消さずに書き直す（何度実行してもよい）
  gas.asUser(() => sheet.getRange(2, 3).setValue('全体の資料'));
  g.rebuildMeetingIndex();
  assert.strictEqual(gas.alerts[gas.alerts.length - 2].message.indexOf('前の形式'), -1);
  assert.deepStrictEqual([sheet.getLastRow(), sheet.getRange(2, 3).getValue()], [3, '全体の資料']);
  assert.deepStrictEqual(['2026/10/23', '20261023', '261023_第5回会議', '261023第5回', '261332_第5回'].map(g.mtgMeetingOfFolder_), ['', '', '2026/10/23', '', '']);
});

test('インデックス：Sheets API（拡張サービス）があれば、PDF をスマートチップ（リッチリンク）で書く。書けなければリンクにする', () => {
  const {gas, g} = setup();
  const calls = [];
  g.Sheets = {Spreadsheets: {batchUpdate: (resource, id) => { calls.push({resource, id}); return {}; }}};
  const meeting = g.mtgParseDate_('10/23', new Date());
  meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  const sheet = gas.ss.getSheetByName('INDEX_会議資料');
  const pdf = gas.driveFiles.find(f => f.folder === 'meeting-folder-0');
  assert.strictEqual(calls.length, 1);
  assert.strictEqual(calls[0].id, gas.ss.getId());
  assert.deepStrictEqual(plain(calls[0].resource), {requests: [{updateCells: {
    start: {sheetId: sheet.getSheetId(), rowIndex: 1, columnIndex: 3},
    rows: [{values: [{userEnteredValue: {stringValue: '@'}, chipRuns: [{startIndex: 0, chip: {richLinkProperties: {uri: 'https://drive.google.com/file/d/' + pdf.id + '/view'}}}],
      userEnteredFormat: {textFormat: {underline: false, foregroundColorStyle: {rgbColor: {red: 0, green: 0, blue: 0}}}}}]}],
    fields: 'userEnteredValue,chipRuns,textFormatRuns,userEnteredFormat.textFormat.underline,userEnteredFormat.textFormat.foregroundColorStyle'
  }}]});
  assert.doesNotMatch(gas.alerts[gas.alerts.length - 1].message, /スマートチップにできなかった/);

  g.Sheets = {Spreadsheets: {batchUpdate: () => { throw new Error('Invalid requests[0].updateCells'); }}};
  const r = g.mtgWriteIndex_('新FMT2', meeting, {folder: {getUrl: () => 'f', getName: () => 'f'}, file: {getUrl: () => 'https://drive.google.com/file/d/X/view', getName: () => 'X.pdf'}});
  assert.strictEqual(r.chip, false);
  assert.strictEqual(sheet.getRange(2, 5).getRichTextValues()[0][0].getLinkUrl(), 'https://drive.google.com/file/d/X/view', '書けなければファイル名のリンク');
  assert.deepStrictEqual(plain(sheet.getRange(2, 5).getRichTextValues()[0][0].styleAt(0)), {color: '#000000', underline: false}, '黒・下線なし');
  assert.ok(gas.logs.some(l => l.level === 'warn' && /チップを書けなかった/.test(l.line)));
});

test('インデックス：書けないとき（ほかの処理が実行中など）は格納は済ませて知らせ、storeMeetingSummaries() で書き直す', () => {
  const {gas, g} = setup();
  const meeting = g.mtgParseDate_('10/23', new Date());
  meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  gas.lockBusy = 'document';   // 差分追跡などがドキュメントロックを持っている
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  gas.lockBusy = false;
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /会議のフォルダ「\d{6}_第5回会議」に格納しました/);
  assert.match(done, /※ 「INDEX_会議資料」シートに書き込めませんでした（ほかの処理が実行中）。storeMeetingSummaries\(\) を実行すると書き直します。/);
  assert.strictEqual(gas.ss.getSheetByName('INDEX_会議資料'), null);

  gas.promptAnswers = ['1', '10/23'];
  g.storeMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「INDEX_会議資料」シートに書きました：https:\/\/docs\.google\.com\/spreadsheets\//);
  assert.match(gas.ss.getSheetByName('INDEX_会議資料').getRange(2, 4).getDisplayValue(), /_新FMT_RX Business Meeting用資料\.pdf$/);
});

test('会議時点のコピー：copyFolderId を設定すれば、そのフォルダに保存する（開けなければ止める）', () => {
  const {gas, g} = setup({edits: false});
  gas.get('MTG_OPTIONS').copyFolderId = 'https://drive.google.com/drive/folders/CopyFolder000001';
  g.mtgSaveCopy_(gas.ss, ['新FMT'], new Date(), '');
  assert.deepStrictEqual(plain(gas.moved), [{id: gas.created[0].getId(), folder: 'CopyFolder000001'}]);
  assert.strictEqual(gas.driveFolders.length, 0, 'フォルダは作らない');
  gas.missingFolders = ['CopyFolder000001'];
  assert.throws(() => g.mtgSaveCopy_(gas.ss, ['新FMT'], new Date(), ''), /会議時点のコピーの保存先のフォルダ（ID：CopyFolder000001）を開けません/);
});

/* ---------------- 会議時点のコピー ---------------- */

test('mtgSaveCopy_：追跡シートの値を、実行日時のタイトルの別のスプレッドシートに写す（文字は数式にならない。書式・見出しの固定も）', () => {
  const {gas, g, sheet} = setup({edits: false});
  gas.asUser(() => sheet.getRange(3, COL.issue).setValue("'=SUM(1)"));   // 「=」で始まる文字
  sheet.getRange(7, COL.stamp).setNumberFormat('yyyy/MM/dd HH:mm');
  const now = new Date(Date.UTC(2026, 9, 16, 9, 30));   // 日本時間 2026/10/16 18:30
  const copy = g.mtgSaveCopy_(gas.ss, ['新FMT', '新FMT2'], now, '（会議日 2026-10-23）');
  assert.strictEqual(copy.title, '会議時点のコピー_2026-10-16 18:30（会議日 2026-10-23）');
  const book = gas.created[0];
  assert.deepStrictEqual([book.getId(), book.getName()], [copy.id, copy.title]);
  assert.deepStrictEqual(book.getSheets().map(t => t.getName()), ['新FMT', '新FMT2'], '追跡シートごとに同じ名前のシート');
  const target = book.getSheetByName('新FMT');
  const rows = sheet.getLastRow();
  const cols = sheet.getLastColumn();
  assert.deepStrictEqual(plain(target.getRange(1, 1, rows, cols).getDisplayValues()), plain(sheet.getRange(1, 1, rows, cols).getDisplayValues()),
    '同じ位置に、同じ表示の値');
  assert.deepStrictEqual([target.getRange(3, COL.issue).getFormula(), target.getRange(3, COL.issue).getValue()], ['', '=SUM(1)'], '文字のまま（数式にしない）');
  assert.ok(target.getRange(7, COL.stamp).getValue() instanceof gas.Date, '日付は日付のまま');
  assert.strictEqual(target.getFrozenRows(), 2, '見出しの行まで固定');
  assert.deepStrictEqual([target.getMaxRows(), target.getMaxColumns()], [rows, cols], '使わない行・列は持たない');
  assert.deepStrictEqual(plain(gas.driveFolders), [{id: 'folder-1', name: '会議時点のコピー', parent: gas.get('MTG_OPTIONS').folderId}], '会議サマリのフォルダの中のフォルダ');
  assert.deepStrictEqual(plain(gas.moved), [{id: copy.id, folder: 'folder-1'}]);
  g.mtgSaveCopy_(gas.ss, ['新FMT'], now, '');
  assert.strictEqual(gas.driveFolders.length, 1, '2回目からは同じフォルダ');
  assert.deepStrictEqual(gas.writes.filter(w => w.ss === gas.ss.getId()), [], '元のスプレッドシートには書き込まない');
});

test('前回からの変化を会議時点のコピーと比べる：差分追跡を通らなかった変更（止まっていた間など）も出し、記録の無い変更として数える', () => {
  const {gas, g, sheet} = setup({edits: false});
  const copy = g.mtgSaveCopy_(gas.ss, ['新FMT'], new Date(), '');
  g.recordDiffEdit(gas.edit(sheet, 'J3', '10/10 提案書を提出'));   // A社：活動状況（記録あり）
  // ここからは差分追跡を通らない変更（止まっていた間の変更の代わり）
  gas.asUser(() => {
    sheet.getRange(7, COL.status).setValue('受注見込み');        // E社：提案中 → 受注見込み（前進）
    sheet.getRange(5, COL.status).setValue('失注');              // C社：失注
    sheet.getRange(4, COL.issue).setValue('請求の確認');          // B社：ほかの項目だけ
    sheet.deleteRow(8);                                        // F社：削除
    sheet.getRange(8, COL.customer, 1, 3).setValues([['G社', '新規案件', '初回訪問']]);   // G社：新規
  });
  const opened = g.mtgOpenCopy_({id: copy.id, sheet: '新FMT'}, '新FMT');
  assert.strictEqual(opened.title, copy.title);
  const s = summarize(g, '新FMT', {copy: opened});
  const byLabel = {};
  s.changes.forEach(c => { byLabel[label(c)] = c; });
  assert.deepStrictEqual(Object.keys(byLabel).sort().map(k => [k, byLabel[k].category]), [
    ['A社／アプリ刷新', 'activity'], ['B社／EC立ち上げ', 'updated'], ['C社／店舗什器', 'lost'], ['E社／販促DX', 'forward'],
    ['F社／ブランド刷新', 'removed'], ['G社／新規案件', 'new']]);
  assert.deepStrictEqual(plain(byLabel['B社／EC立ち上げ'].columns), ['活動における課題'], '更新項目はコピーと違う列');
  assert.deepStrictEqual([byLabel['F社／ブランド刷新'].statusFrom, byLabel['F社／ブランド刷新'].amount], ['提案中', 10], '削除した案件はコピーの値');
  assert.strictEqual(s.unlogged, 5, 'A社のほかは変更の記録が無い');
  assert.ok(byLabel['A社／アプリ刷新'].logged);
  assert.strictEqual(s.gaps.changes, undefined);
  assert.deepStrictEqual(plain(s.copy), {title: copy.title, url: opened.url});
  assert.match(g.mtgAppendixChanges_(s)[0].lead, /うち 5件は変更の記録が無い（前回の会議時点のコピーとの比較で検出）。/);
  assert.match(JSON.stringify(g.mtgCoverSlide_(s).elements), /出典：「新FMT」シート・前回の会議時点のコピーおよび変更履歴/);
  // 期待値の増減もコピーの値から：C社（5）が失注、F社（3）が削除、G社は期待値なし
  assert.strictEqual(s.bridge.start, 10 + 5 + 40 + 3);
  assert.strictEqual(s.bridge.end, 10 + 40);
});

test('会議時点のコピーと比べるとき：得意先・案件名を書き換えた案件は、変更の記録で前の名前とつなげる（新規・削除にしない）。値が戻っていれば変化なし', () => {
  const {gas, g, sheet} = setup({edits: false});
  const copy = g.mtgSaveCopy_(gas.ss, ['新FMT'], new Date(), '');
  g.recordDiffEdit(gas.edit(sheet, 'E3', 'アプリ刷新（第2期）'));
  g.recordDiffEdit(gas.edit(sheet, 'J5', '一時的な変更'));
  g.recordDiffEdit(gas.edit(sheet, 'J5', '見積提出'));   // 元に戻した
  const s = summarize(g, '新FMT', {copy: g.mtgOpenCopy_({id: copy.id, sheet: '新FMT'}, '新FMT')});
  assert.deepStrictEqual(plain(s.changes.map(c => [label(c), c.category, c.columns.join('・')])),
    [['A社／アプリ刷新（第2期）', 'updated', '案件名']]);
});

test('前回の会議時点のコピーを開けないときは、変更の記録から出して知らせる。試作はコピーを保存せず、基準は保存する', () => {
  const {gas, g, sheet} = setup();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.strictEqual(gas.created.length, 0, '試作はコピーを保存しない');
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.strictEqual(gas.created.length, 1);
  const props = g.PropertiesService.getDocumentProperties();
  assert.deepStrictEqual(plain(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT'))[0].copy), {id: gas.created[0].getId(), sheet: '新FMT'});
  assert.match(gas.alerts[gas.alerts.length - 1].message, /今の値を「会議時点のコピー_.*（会議日 \d{4}-10-23）」に保存しました（次の会議の資料はこれと比べます）：https:\/\/docs\.google\.com\/spreadsheets\/d\/created-1\/edit/);

  g.recordDiffEdit(gas.edit(sheet, 'F7', '受注見込み'));
  gas.unopenable = [gas.created[0].getId()];
  gas.promptAnswers = ['1', '10/30'];
  g.previewMeetingSummaries();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /前回の「会議時点のコピー」を開けないため、前回からの変化は変更の記録から出しました/);
  assert.match(done, /（変化のある案件 1・/, '変更の記録から出す');

  gas.promptAnswers = ['1'];
  g.setMeetingBaseline();
  assert.match(gas.created[gas.created.length - 1].getName(), /^会議時点のコピー_\d{4}-\d{2}-\d{2} \d{2}:\d{2}（基準）$/);
  const base = JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).find(m => m.baseline);
  assert.deepStrictEqual(plain(base.copy), {id: gas.created[gas.created.length - 1].getId(), sheet: '新FMT'});
});

/* ---------------- 作る・保存する ---------------- */

test('createMeetingSummaries：シートと会議日を聞き、シートごとに Google スライド（16:9）と PDF を保存する。会議日を記録し、次の会議は今以降の変化を出す', () => {
  const {gas, g, sheet} = setup();
  gas.get('MTG_OPTIONS').keepSlides = true;   // Google スライドも残す設定で確かめる
  gas.get('MTG_OPTIONS').folderId = '';   // 保存先を決めていないとき：スプレッドシートと同じフォルダの中に作る
  const meeting = g.mtgParseDate_('10/23', new Date());
  const file = meeting.replace(/\//g, '');   // ファイル名の会議日（yyyyMMdd）
  gas.get('MTG_OPTIONS').batchRequests = 500;   // まとめて送る数を小さくして、分けて送るところも確かめる
  gas.promptAnswers = ['', '10/23'];   // 空欄 → すべてのシート
  g.createMeetingSummaries();

  assert.match(gas.prompts[0].message, /1\. 新FMT\n2\. 新FMT2/);
  assert.match(gas.prompts[1].message, /今回の会議日を入力してください（例：10\/23）。[\s\S]*・新FMT：記録なし\n・新FMT2：記録なし$/);
  assert.deepStrictEqual(gas.decks.map(d => [d.title, d.pageSize.width.magnitude, d.pageSize.height.magnitude]),
    [[file + '_新FMT_RX Business Meeting用資料', 720, 405], [file + '_新FMT2_RX Business Meeting用資料', 720, 405]], 'サブインダストリー（シート）ごとに1つ（統合版を使わない設定のとき）');
  const requests = gas.decks[0].requests;
  assert.strictEqual(requests.filter(r => r.createSlide).length, 10, '表紙・本編4枚・扉・Appendix ①・②・③（責任者別・一覧）（停滞が無くても ② を出す）');
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[1]), /^新FMT2\nRX Business MTG　事前資料\n/);
  assert.match(inserted(gas.decks[1]), /\nサマリー\n/);
  assert.match(inserted(gas.decks[1]), /\n新FMT2（10\/23 会議資料）　金額の単位：百万円\n/);
  assert.deepStrictEqual(requests.filter(r => r.deleteObject).map(r => r.deleteObject.objectId), ['p'], '作ったときの白紙のスライドを消す');
  const batches = gas.decks[0].batches;
  assert.ok(batches.length >= 2 && batches.length < 8, '何枚かずつまとめて送る（' + batches.length + '回）');
  assert.ok(requests.findIndex(r => r.deleteObject) < batches[0], '白紙のスライドは1回目で消す');
  assert.deepStrictEqual(plain(gas.driveFolders), [{id: 'folder-1', name: '会議サマリ', parent: 'parent-of-spreadsheet-id'},
    {id: 'folder-2', name: '会議時点のコピー', parent: 'folder-1'}], 'スプレッドシートと同じフォルダの中に作る（会議時点のコピーはその中のフォルダ）');
  assert.deepStrictEqual(gas.driveFiles.map(f => [f.name, f.folder, f.blob.from, f.blob.getContentType()]), [
    [file + '_新FMT_RX Business Meeting用資料.pdf', 'folder-1', 'deck-1', 'application/pdf'],
    [file + '_新FMT2_RX Business Meeting用資料.pdf', 'folder-1', 'deck-2', 'application/pdf']
  ]);
  assert.deepStrictEqual(plain(gas.moved), [{id: 'deck-1', folder: 'folder-1'}, {id: 'deck-2', folder: 'folder-1'}, {id: 'created-1', folder: 'folder-2'}],
    'スライドも同じフォルダに残す。会議時点のコピーは、その中のフォルダ');
  // ひな形が無いので、表紙（1枚目）と扉（6枚目）の背景を、緑のグラデーションと RX の画像にする（PDF にする前に保存する）
  assert.deepStrictEqual(gas.slideBackgrounds.map(b => [b.deck, b.page, b.blob.getContentType(), b.blob.getName()]), [
    ['deck-1', 'mtg_p1', 'image/png', 'meeting-cover.png'], ['deck-1', 'mtg_p6', 'image/png', 'meeting-divider.png'],
    ['deck-2', 'mtg_p1', 'image/png', 'meeting-cover.png'], ['deck-2', 'mtg_p6', 'image/png', 'meeting-divider.png']], '表紙は緑、扉はグレーの絵柄');
  assert.deepStrictEqual(gas.slideBackgrounds[0].blob.getBytes().slice(0, 4), [-119, 80, 78, 71], 'PNG（バイトは -128〜127 で渡す）');
  assert.deepStrictEqual(gas.decks.map(d => d.savedAndClosed), [1, 1]);
  assert.deepStrictEqual(gas.trashed, []);
  assert.deepStrictEqual(gas.writes.filter(w => w.ss === gas.ss.getId()), [], 'スプレッドシートには書き込まない');
  // 本番は、作る前に会議日（曜日）・シートごとの対象期間を出して確かめる（配布したあとは作り直さない）
  const confirm = gas.alerts.find(a => /本番の資料を作り、会議日を記録します。/.test(a.message)).message;
  assert.match(confirm, /会議日：\d{4}\/10\/23（[日月火水木金土]）/);
  assert.match(confirm, /・新FMT：直近14日間（前回の会議の記録なし）\n・新FMT2：直近14日間（前回の会議の記録なし）/);
  assert.match(confirm, /配布したあとは作り直さないでください。/);
  assert.doesNotMatch(confirm, /作成済み/);
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /「新FMT」：\d{4}1023_新FMT_RX Business Meeting用資料\.pdf（変化のある案件 4・停滞 0・確認事項 3）\n　https:\/\/drive\.google\.com\/file\/d\/file-1\/view/,
    '配る PDF の URL を出す');
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
  assert.strictEqual(gas.driveFolders.length, 2);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /前回からの変化は「会議時点のコピー_\d{4}-\d{2}-\d{2} \d{2}:\d{2}（会議日 \d{4}-10-23）」と比べました。/);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：.*（変化のある案件 1・停滞 0・確認事項 2）/);
  assert.deepStrictEqual(gas.driveFiles.slice(2).map(f => f.name.replace(/^\d{8}_/, '')), ['新FMT_RX Business Meeting用資料.pdf'], '選んだシートだけ作る');
  // 同じ会議日で作り直しても、前回（10/23）からの変化（作った日時だけ上書き）。作り直す前に、作成済みであることを知らせる
  const before = {deck: gas.decks[gas.decks.length - 1].id, pdf: gas.driveFiles[gas.driveFiles.length - 1].id};
  gas.promptAnswers = ['1', '10/30'];
  g.createMeetingSummaries();
  assert.match(gas.alerts.filter(a => /本番の資料を作り/.test(a.message)).pop().message,
    /・新FMT：前回会議（10\/23）の資料作成時（\d{1,2}\/\d{1,2} \d{2}:\d{2}）以降（その時点のコピーと比べる）\n　※ この会議日の資料は \d{1,2}\/\d{1,2} \d{2}:\d{2} に作成済み。作り直すと、次の会議の起点が今になります。/);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（変化のある案件 1・/);
  assert.deepStrictEqual(gas.trashed.slice().sort(), [before.deck, before.pdf].sort(), '同じ名前の古い版（スライド・PDF）はゴミ箱に移す（古い版を配らないように）');
  assert.ok(gas.trashed.indexOf(gas.decks[gas.decks.length - 1].id) < 0 && gas.trashed.indexOf(gas.driveFiles[gas.driveFiles.length - 1].id) < 0, '新しい版は残す');
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).map(m => m.date), [meeting, next]);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT2')).map(m => m.date), [meeting], '選ばなかったシートの会議日は記録しない');
});

test('Google スライドを残さない（keepSlides: false。今の設定）：PDF だけをフォルダに置き、スライドはフォルダに動かさずにゴミ箱へ', () => {
  const {gas, g} = setup();
  assert.strictEqual(gas.get('MTG_OPTIONS').keepSlides, false);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.deepStrictEqual(gas.decks.map(d => gas.trashed.indexOf(d.id) >= 0), [true], 'スライドはゴミ箱へ');
  assert.ok(!gas.moved.some(m => m.id === gas.decks[0].id), 'スライドはフォルダに動かさない');
  assert.strictEqual(gas.driveFiles.filter(f => /\.pdf$/.test(f.name) && gas.trashed.indexOf(f.id) < 0).length, 1, 'PDF は残す');
  assert.doesNotMatch(gas.alerts[gas.alerts.length - 1].message, /Google スライド（同じ名前）も保存しました/);
});

test('setMeetingTemplate・ひな形：登録した Google スライドをコピーして資料を作る（表紙は決まった文字の置き換え、扉はグレーの絵柄の背景に描く、ほかはページの大きさに合わせて拡大）', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').folderId = '';   // 保存先を決めていないとき：スプレッドシートと同じフォルダの中に作る
  const props = g.PropertiesService.getScriptProperties();
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789';   // .pptx を変換した Google スライド
  gas.slideTemplates[id] = pptxTemplate();
  const pptx = '1xnBht3M2bJRMtD3ptzYp1glYyhfXnxJN';
  gas.files[pptx] = {mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', name: '表紙.pptx', updated: 1000};
  gas.promptAnswers = ['ひな形'];
  assert.throws(() => g.setMeetingTemplate(), /「ひな形」から、ひな形のファイルの ID を読めません/);
  gas.files['1SheetFileIdAbCdEfGhIjKlMnOpQrStUv'] = {mimeType: 'application/vnd.google-apps.spreadsheet', name: '表'};
  gas.promptAnswers = ['https://docs.google.com/spreadsheets/d/1SheetFileIdAbCdEfGhIjKlMnOpQrStUv/edit'];
  assert.throws(() => g.setMeetingTemplate(), /ひな形（表）が Google スライドでも \.pptx でもありません/);
  // .pptx の変換に失敗したら、Google スライドとして保存する手順を案内する
  gas.fetchStatuses = [403];
  gas.fetchText = '{"error": "insufficient permissions"}';
  gas.promptAnswers = ['https://docs.google.com/presentation/d/' + pptx + '/edit?rtpof=true'];
  assert.throws(() => g.setMeetingTemplate(), /ひな形（表紙\.pptx）を Google スライドに変換できませんでした（403：.*）。ドライブで開いて「ファイル」→「Google スライドとして保存」/);
  assert.strictEqual(props.getProperty('MTG_TEMPLATE_URL'), null);

  // .pptx の URL（Google スライドで開いたときの ?rtpof=true つき）を、スクリプトのプロパティに入れる。Google スライドに変換して使う
  gas.fetchText = JSON.stringify({id});
  gas.promptAnswers = ['https://docs.google.com/presentation/d/' + pptx + '/edit?rtpof=true'];
  g.setMeetingTemplate();
  assert.strictEqual(props.getProperty('MTG_TEMPLATE_URL'), 'https://docs.google.com/presentation/d/' + pptx + '/edit?rtpof=true');
  const convert = gas.fetches[gas.fetches.length - 1];
  assert.strictEqual(convert.url, 'https://www.googleapis.com/drive/v3/files/' + pptx + '/copy?supportsAllDrives=true&fields=id');
  assert.deepStrictEqual([convert.params.method, convert.params.headers.Authorization, JSON.parse(convert.params.payload)],
    ['post', 'Bearer oauth-token', {name: '会議サマリのひな形（表紙.pptx から変換）', mimeType: 'application/vnd.google-apps.presentation', parents: ['folder-1']}]);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_TEMPLATE_CONVERTED')), {source: pptx, updated: 1000, id});
  assert.match(gas.alerts[gas.alerts.length - 1].message,
    /ひな形を登録しました（3枚。1枚目を表紙、2枚目を扉（Appendix）に使います。統合版の表紙も1枚目を使います）。[\s\S]*\.pptx を Google スライドに変換して使います/);
  const fetched = gas.fetches.length;

  gas.layoutPlaceholders = ['SLIDE_NUMBER', 'FOOTER'];   // ひな形の白紙のレイアウトにはページ番号・フッターの枠がある
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.deepStrictEqual(gas.driveCopies.map(cp => [cp.from, cp.folder]), [[id, 'folder-1']], '変換したひな形を保存先のフォルダにコピーする');
  assert.strictEqual(gas.fetches.length, fetched, '.pptx を更新していなければ、変換し直さない');
  const deck = gas.decks.find(d => d.template === id);
  const r = plain(deck.requests);
  assert.deepStrictEqual(r.filter(q => q.deleteObject && /_ph\d$/.test(q.deleteObject.objectId)).map(q => q.deleteObject.objectId),
    ['mtg_p2_ph0', 'mtg_p2_ph1', 'mtg_p3_ph0', 'mtg_p3_ph1', 'mtg_p4_ph0', 'mtg_p4_ph1', 'mtg_p5_ph0', 'mtg_p5_ph1',
      'mtg_p7_ph0', 'mtg_p7_ph1', 'mtg_p8_ph0', 'mtg_p8_ph1', 'mtg_p9_ph0', 'mtg_p9_ph1', 'mtg_p10_ph0', 'mtg_p10_ph1'],
    '足したスライドのレイアウトの枠（ページ番号など）は消す（ページ番号が2つ出ないように）');
  assert.ok(deck.slides.every(sl => !(sl.pageElements || []).some(el => el.shape && el.shape.placeholder)));
  assert.deepStrictEqual(r[0], {deleteObject: {objectId: 'g_combined'}}, '使わないページ（統合版の表紙）は消す');
  assert.deepStrictEqual(deck.batches[0], 1, '先に消して（複製して）から、ひな形のページを読み直す');
  assert.ok(!r.some(q => q.duplicateObject || q.replaceAllText), '1回ずつなので複製しない。同じ文字の枠が2つあるので replaceAllText は使わない');
  // 決まった文字は枠ごとに書き換える：見出し（40pt）→ サブインダストリー、その下（20pt）→ 会議の名前、日付、名前の欄は空
  const texts = page => r.filter(q => (q.insertText || q.deleteText) && (q.insertText || q.deleteText).objectId.indexOf(page + '_') === 0)
    .map(q => (q.insertText ? ['入れる', q.insertText.objectId.slice(page.length + 1), q.insertText.insertionIndex, q.insertText.text]
      : ['消す', q.deleteText.objectId.slice(page.length + 1), q.deleteText.textRange.type, q.deleteText.textRange.startIndex, q.deleteText.textRange.endIndex]));
  const date = r.find(q => q.insertText && q.insertText.objectId === 'g_cover_date').insertText.text;
  assert.match(date, /^\d{4}年10月23日（[日月火水木金土]）$/);
  assert.deepStrictEqual(texts('g_cover').filter(t => t[1] !== 'extra'), [
    ['入れる', 'title', 0, '新FMT'], ['消す', 'title', 'FIXED_RANGE', 4, 9],
    ['入れる', 'subtitle', 0, 'RX Business MTG　事前資料'], ['消す', 'subtitle', 'FIXED_RANGE', 20, 25],
    ['入れる', 'date', 0, date], ['消す', 'date', 'FIXED_RANGE', date.length, date.length + '2026年 4月 17日'.length],
    ['消す', 'name', 'ALL', undefined, undefined]
  ], '新しい文字を先頭に入れてから元の文字を消す（元の書式を引き継ぐ）');
  // 扉：ひな形の2枚目（グレー）。見出しは「Appendix」、ほかの文字は消す。一覧はその上に描く
  assert.deepStrictEqual(texts('g_divider'), [['入れる', 'title', 0, 'Appendix'], ['消す', 'title', 'FIXED_RANGE', 8, 13],
    ['消す', 'subtitle', 'ALL', undefined, undefined], ['消す', 'date', 'ALL', undefined, undefined], ['消す', 'name', 'ALL', undefined, undefined]]);
  assert.ok(r.some(q => q.createShape && q.createShape.elementProperties.pageObjectId === 'g_divider'), '扉に一覧を足す');
  assert.ok(!r.some(q => q.createSlide && q.createSlide.objectId === 'mtg_p6'), '扉はひな形のページ（作らない）');
  assert.ok(!r.some(q => q.deleteObject && /^g_(cover|divider)$/.test(q.deleteObject.objectId)), '表紙・扉（ひな形の1・2枚目）は消さない');
  assert.deepStrictEqual(r.filter(q => q.createSlide).map(q => q.createSlide.insertionIndex), [1, 2, 3, 4, 6, 7, 8, 9]);
  // 最後に、1枚目から順に決まった位置に動かす（ひな形の扉のページは表紙の次にあるため）
  const moves = r.filter(q => q.updateSlidesPosition).map(q => [q.updateSlidesPosition.slideObjectIds[0], q.updateSlidesPosition.insertionIndex]);
  assert.deepStrictEqual(moves, ['g_cover', 'mtg_p2', 'mtg_p3', 'mtg_p4', 'mtg_p5', 'g_divider', 'mtg_p7', 'mtg_p8', 'mtg_p9', 'mtg_p10'].map((id, i) => [id, i]));
  // 本編はページの大きさ（960 × 540pt）に合わせて 4/3 倍
  const title = r.find(q => q.createShape && q.createShape.elementProperties.pageObjectId === 'mtg_p2').createShape;
  assert.deepStrictEqual([title.elementProperties.size.width.magnitude, title.elementProperties.transform.translateX], [901.33, 29.33]);
  const titleStyle = r.find(q => q.updateTextStyle && q.updateTextStyle.objectId === title.objectId).updateTextStyle.style;
  assert.strictEqual(titleStyle.fontSize.magnitude, 28);
  assert.ok(r.some(q => q.createShape && q.createShape.elementProperties.pageObjectId === 'g_cover'), '表紙にデータの時点・期間を足す');
  assert.ok(!gas.moved.some(m => m.id === deck.id), 'コピーは保存先のフォルダにできるので動かさない');
  assert.strictEqual(gas.driveFiles[0].blob.from, deck.id, 'ひな形のコピーを PDF にする');
  assert.deepStrictEqual(gas.slideBackgrounds, [], 'ひな形があれば、表紙・扉はひな形のまま（背景の画像は付けない）');

  // .pptx を更新したら変換し直す（前に変換したものはゴミ箱へ）。1枚だけのひな形なら、扉はグレーの絵柄で描く
  const id2 = '1ConvertedAgainAbCdEfGhIjKlMnOpQrStUv';
  gas.slideTemplates[id2] = pptxTemplate(1);
  gas.files[pptx].updated = 2000;
  gas.fetchText = JSON.stringify({id: id2});
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.strictEqual(gas.fetches.length, fetched + 1);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_TEMPLATE_CONVERTED')), {source: pptx, updated: 2000, id: id2});
  assert.ok(gas.trashed.indexOf(id) >= 0, '前に変換したものはゴミ箱へ');
  const one = gas.decks[gas.decks.length - 1];
  assert.ok(one.requests.some(q => q.createSlide && q.createSlide.objectId === 'mtg_p6'));
  assert.deepStrictEqual(plain(gas.slideBackgrounds.map(b => [b.deck, b.page, b.blob.getName()])), [[one.id, 'mtg_p6', 'meeting-divider.png']]);

  // 空欄で登録を外すと、ひな形なしで作る
  gas.promptAnswers = [''];
  g.setMeetingTemplate();
  assert.strictEqual(props.getProperty('MTG_TEMPLATE_URL'), null);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.strictEqual(gas.driveCopies.length, 2);

  // スクリプトのプロパティに Google スライドの URL を直接入れてもよい（変換しない）。前の版の登録（ドキュメントのプロパティ）は、空のときだけ読む
  props.setProperty('MTG_TEMPLATE_URL', 'https://docs.google.com/presentation/d/' + id + '/edit');
  g.PropertiesService.getDocumentProperties().setProperty('MTG_TEMPLATE_ID', id2);
  assert.strictEqual(g.mtgTemplateId_(), id);
  props.deleteProperty('MTG_TEMPLATE_URL');
  assert.strictEqual(g.mtgTemplateId_(), id2);
  assert.strictEqual(gas.fetches.length, fetched + 1);
});

/**
 * いただいたひな形（2026/10/17。3枚：緑の表紙・グレーの扉・紫の統合版の表紙）を Google スライドにしたものと同じ形（文字の枠は見出し 40pt・その下 20pt が
 * どちらも「XXXXX」、日付・名前 16pt）。count 枚まで
 */
function pptxTemplate(count) {
  const box = (page, key, text, size, y) => ({objectId: page + '_' + key, transform: {translateY: y * 12700, unit: 'EMU'},
    shape: {shapeType: 'TEXT_BOX', text: {textElements: [{paragraphMarker: {}}, {textRun: {content: text + '\n', style: {fontSize: {magnitude: size, unit: 'PT'}}}}]}}});
  const page = id => ({objectId: id, pageElements: [{objectId: id + '_bg', shape: {shapeType: 'RECTANGLE'}},
    box(id, 'rx', 'RX', 857, -252.9), box(id, 'subtitle', 'XXXXX', 20, 305.6), box(id, 'title', 'XXXXX', 40, 231.8),
    box(id, 'name', 'XXXX', 16, 463.2), box(id, 'date', '2026年 4月 17日', 16, 443.8)]});
  return {pageSize: {width: {magnitude: 12192000, unit: 'EMU'}, height: {magnitude: 6858000, unit: 'EMU'}},
    slides: ['g_cover', 'g_divider', 'g_combined'].slice(0, count || 3).map(page)};
}

test('previewMeetingSummaries：試作は会議日を記録しない（ファイル名に「試作」）。最初の試作の入力内容のチェックの結果を覚え、同じ会議日の資料で解消した件数を出す', () => {
  const {gas, g, sheet} = setup();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.ok(gas.driveFiles.every(f => /_試作\.pdf$/.test(f.name)));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（変化のある案件 4・停滞 0・確認事項 3）[\s\S]*試作のため、会議日は記録していません。/, '2回目も同じ変化を出す');
  const props = g.PropertiesService.getDocumentProperties();
  assert.strictEqual(props.getProperty('MTG_MEETINGS_新FMT'), null);
  const draft = JSON.parse(props.getProperty('MTG_DRAFT_新FMT'));
  assert.deepStrictEqual([draft.meeting, draft.keys.length], [g.mtgParseDate_('10/23', new Date()), 3]);
  assert.deepStrictEqual(gas.decks.map(d => d.title.replace(/^\d{8}_/, '')), ['新FMT_RX Business Meeting用資料_試作', '新FMT_RX Business Meeting用資料_試作']);
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[1]), /入力内容の確認事項は 3項目（3案件・責任者 3名）。\n?試作時（\d{1,2}\/\d{1,2}）から 0項目解消、\n?0項目追加。/,
    '2回目の試作：まだ解消していない');

  // E社の活動状況を書いて「更新なし」を解消してから、本番を作る
  g.recordDiffEdit(gas.edit(sheet, 'J7', '10/20 決裁者と面談を設定'));
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(inserted(gas.decks[2]), /入力内容の確認事項は 2項目（2案件・責任者 2名）。\n?試作時（\d{1,2}\/\d{1,2}）から 1項目解消、\n?0項目追加。/);
});

test('スプレッドシートの画面に出せないとき（エディタをスプレッドシートから開いていない）：基準はすべての追跡シートに記録してログに書く。ほかは、どうすればよいかを書いて止める', () => {
  const {gas, g} = setup({edits: false});
  gas.noUi = true;
  g.setMeetingBaseline();
  const props = g.PropertiesService.getDocumentProperties();
  assert.deepStrictEqual(['新FMT', '新FMT2'].map(name => JSON.parse(props.getProperty('MTG_MEETINGS_' + name)).map(m => !!m.baseline)), [[true], [true]]);
  assert.ok(gas.logs.some(l => /「新FMT」「新FMT2」に、基準（.*）を記録しました。[\s\S]*確かめずにすべての追跡シートに記録しました/.test(l.line)), '実行ログに書く');
  assert.deepStrictEqual([gas.prompts.length, gas.alerts.length], [0, 0]);
  ['createMeetingSummaries', 'previewMeetingSummaries', 'removeMeetingRecord', 'setMeetingChat', 'notifyMeetingSummaries', 'setMeetingTemplate'].forEach(fn => {
    if (fn === 'notifyMeetingSummaries') {
      g.PropertiesService.getScriptProperties().setProperty('MTG_CHAT_WEBHOOK', 'https://chat.googleapis.com/v1/spaces/A/messages?key=k&token=t');
      g.PropertiesService.getScriptProperties().setProperty('MTG_CHAT_URL', 'https://drive.google.com/file/d/abcdefghijk/view');
    }
    assert.throws(() => g[fn](), /スプレッドシートの画面に、入力・確かめの画面を出せません。スプレッドシートを開いたまま、そのスプレッドシートの「拡張機能」→「Apps Script」から開いたエディタで実行してください/, fn);
  });
});

test('removeMeetingRecord：会議日の記録を取り消す（はいのときだけ）。基準の記録・資料のファイルは残す', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').keepSlides = true;   // Google スライドも残す設定で確かめる
  const props = g.PropertiesService.getDocumentProperties();
  const lastAlert = () => gas.alerts[gas.alerts.length - 1];
  gas.promptAnswers = ['1'];
  g.setMeetingBaseline();
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  const meeting = g.mtgParseDate_('10/23', new Date());
  const files = gas.driveFiles.length;

  gas.promptAnswers = ['1', '10/30'];
  g.removeMeetingRecord();
  assert.match(gas.prompts[gas.prompts.length - 1].message, /取り消す会議日を入力してください（例：10\/23）。\n\n記録してある会議日\n・新FMT：10\/23$/);
  assert.match(lastAlert().message, /会議日 \d{4}\/10\/30 の記録はありません。/);

  gas.confirmAnswer = 'NO';
  gas.promptAnswers = ['1', '10/23'];
  g.removeMeetingRecord();
  assert.match(lastAlert().message, /「新FMT」の会議日（\d{4}\/10\/23）の記録を取り消します。/);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).map(m => m.date).slice(-1), [meeting], '「いいえ」なら消さない');

  gas.confirmAnswer = 'YES';
  gas.promptAnswers = ['1', '10/23'];
  g.removeMeetingRecord();
  assert.match(lastAlert().message, /「新FMT」の会議日（\d{4}\/10\/23）の記録を取り消しました。/);
  const list = JSON.parse(props.getProperty('MTG_MEETINGS_新FMT'));
  assert.deepStrictEqual(list.map(m => !!m.baseline), [true], '基準の記録は残す');
  assert.strictEqual(gas.driveFiles.length, files);
  assert.deepStrictEqual(gas.trashed, [], '資料のファイルは消さない');
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(gas.alerts.filter(a => /本番の資料を作り/.test(a.message)).pop().message, /・新FMT：基準（/, '取り消したあとは、基準からの変化に戻る');
  assert.throws(() => { gas.promptAnswers = ['1', 'あした']; g.removeMeetingRecord(); }, /「あした」を会議日として読めません。/);
});

test('本番の確認で「いいえ」なら作らない（会議日も記録しない）', () => {
  const {gas, g} = setup();
  gas.confirmAnswer = 'NO';
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.deepStrictEqual([gas.decks.length, gas.driveFiles.length], [0, 0]);
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT'), null);
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.strictEqual(gas.decks.length, 1, '試作は確かめずに作る（会議日を記録しないため）');

  // 前の版の記録（MTG_LAST_）しか無いシートは、その時からの変化（資料と同じ起点を出す）
  g.PropertiesService.getDocumentProperties().setProperty('MTG_LAST_新FMT', String(Date.now() - 3 * 86400000));
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(gas.alerts.filter(a => /本番の資料を作り/.test(a.message)).pop().message, /・新FMT：前回の資料作成時（\d{1,2}\/\d{1,2} \d{2}:\d{2}）以降\n/);
});

test('mtgTrashOld_：同じ名前の古い版をゴミ箱に移す（新しい版は残す）。移せないファイルがあっても止めない', () => {
  const {g} = setup({edits: false});
  const trashed = [];
  const file = (id, locked) => ({getId: () => id, setTrashed() { if (locked) throw new Error('権限がありません'); trashed.push(id); }});
  const files = {'資料': [file('old-1'), file('new-1'), file('other', true)], '資料.pdf': [file('old-2'), file('new-2')]};
  const folder = {getFilesByName: name => { const list = (files[name] || []).slice(); return {hasNext: () => list.length > 0, next: () => list.shift()}; }};
  assert.strictEqual(g.mtgTrashOld_(folder, ['資料', '資料.pdf'], ['new-1', 'new-2']), 2);
  assert.deepStrictEqual(trashed, ['old-1', 'old-2']);
});

/* ---------------- Google チャットへの通知 ---------------- */

const WEBHOOK = 'https://chat.googleapis.com/v1/spaces/AAAAtest123/messages?key=KEY123&token=TOKEN456';
const DEMO_URL = 'https://drive.google.com/drive/folders/demo-folder-0001';
/** 送った文（webhook に送った JSON の text） */
const sentText = fetch => JSON.parse(fetch.params.payload).text;

test('setMeetingChat：webhook の URL と資料の格納先の URL を登録する（トリガーは作らない。前の版の自動通知のトリガーは外す）。確かめのメッセージを送れる。空欄で外す', () => {
  const {gas, g} = setup({edits: false});
  const props = g.PropertiesService.getScriptProperties();
  g.ScriptApp.newTrigger('mtgNotifyScheduled').timeBased().everyDays(1).create();   // 前の版で作った自動通知のトリガー
  gas.promptAnswers = [WEBHOOK, DEMO_URL];
  g.setMeetingChat();
  assert.match(gas.prompts[0].message, /webhook の URL を貼り付けてください[\s\S]*今の設定：未設定$/);
  assert.match(gas.prompts[1].message, /^送る資料の格納先の URL を貼り付けてください/);
  assert.deepStrictEqual([props.getProperty('MTG_CHAT_WEBHOOK'), props.getProperty('MTG_CHAT_URL')], [WEBHOOK, DEMO_URL],
    'URL はスクリプトのプロパティに残す（コードに書かない）');
  assert.deepStrictEqual(gas.triggers.filter(t => t.getHandlerFunction() === 'mtgNotifyScheduled').length, 0, '前の版の自動通知のトリガーは外す');
  assert.strictEqual(gas.fetches.length, 1, '確かめのメッセージ（「はい」のとき）');
  assert.strictEqual(gas.fetches[0].url, WEBHOOK);
  assert.deepStrictEqual([gas.fetches[0].params.method, gas.fetches[0].params.contentType], ['post', 'application/json; charset=UTF-8']);
  assert.strictEqual(sentText(gas.fetches[0]), '【テスト送信】会議サマリの通知の設定を確認するためのメッセージです。\n' +
    '今後、RX Business MTG の事前資料のご案内をこちらにお送りいたします。');
  assert.match(gas.alerts[gas.alerts.length - 1].message,
    new RegExp('・送り先：スペース AAAAtest123\n・送る URL：' + DEMO_URL + '\n・notifyMeetingSummaries\\(\\) を実行すると、資料のファイルのリンクをすぐに送ります（フォルダなら、その中でいちばん新しい PDF）。\n・前の版の自動通知（2日前）のトリガーを外しました。'));

  // 格納先の URL は空欄なら今のまま。確かめのメッセージは「いいえ」なら送らない
  gas.confirmAnswer = 'NO';
  gas.promptAnswers = [WEBHOOK, ''];
  g.setMeetingChat();
  assert.match(gas.prompts[2].message, /今の設定：スペース AAAAtest123 に送る$/);
  assert.match(gas.prompts[3].message, new RegExp('空欄なら今の URL のまま：' + DEMO_URL + '$'));
  assert.strictEqual(props.getProperty('MTG_CHAT_URL'), DEMO_URL);
  assert.strictEqual(gas.fetches.length, 1);
  assert.ok(gas.alerts.concat(gas.prompts).every(a => !/KEY123|TOKEN456/.test(a.message)) &&
    gas.logs.every(l => !/KEY123|TOKEN456/.test(l.line)), '鍵・トークンは画面・ログに出さない');

  gas.promptAnswers = ['https://example.com/hook'];
  assert.throws(() => g.setMeetingChat(), /Google チャットの webhook の URL ではありません。/);
  gas.promptAnswers = [WEBHOOK, 'デモのフォルダ'];
  assert.throws(() => g.setMeetingChat(), /資料の格納先の URL を読めません。/);
  assert.strictEqual(props.getProperty('MTG_CHAT_URL'), DEMO_URL, '間違えたときは今の設定のまま');

  gas.promptAnswers = [''];
  g.setMeetingChat();
  assert.deepStrictEqual([props.getProperty('MTG_CHAT_WEBHOOK'), props.getProperty('MTG_CHAT_URL')], [null, null]);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /通知の設定を外しました。/);
});

test('notifyMeetingSummaries：会議日程を聞き、格納先のフォルダの、いちばん新しい PDF のリンク（ファイル名で表示）と一緒に、確かめのあとすぐに送る。失敗は1回だけやり直し、だめなら止める', () => {
  const {gas, g} = setup({edits: false});
  gas.promptAnswer = '2026/10/23';   // 会議日程（ほかの答えを決めていなければこれ）
  assert.throws(() => g.notifyMeetingSummaries(), /先に setMeetingChat\(\) で webhook の URL と資料の格納先の URL を登録してください。/);
  const props = g.PropertiesService.getScriptProperties();
  props.setProperty('MTG_CHAT_WEBHOOK', WEBHOOK);
  props.setProperty('MTG_CHAT_URL', DEMO_URL);
  assert.throws(() => g.notifyMeetingSummaries(), /資料の格納先のフォルダ（demo-folder-0001）にファイルがありません。資料を置いてから送ってください。/);
  assert.strictEqual(gas.prompts.length, 0, '資料が無ければ、会議日程を聞かずに知らせる');

  // 格納先のフォルダ：古い PDF・新しい PDF・もっと新しいスライド（PDF を選ぶ）・ほかのフォルダの PDF
  const file = (id, folder, name, type, daysAgo) => gas.driveFiles.push({id, folder, name, blob: {getContentType: () => type},
    updated: Date.now() - daysAgo * 86400000});
  file('pdf-old', 'demo-folder-0001', '会議サマリ_デモ_旧.pdf', 'application/pdf', 5);
  file('pdf-new', 'demo-folder-0001', '会議サマリ_デモ.pdf', 'application/pdf', 1);
  file('slide-1', 'demo-folder-0001', '会議サマリ_デモ（スライド）', 'application/vnd.google-apps.presentation', 0);
  file('pdf-other-0001', 'other-folder-0001', '別の資料.pdf', 'application/pdf', 0);
  const lines = (link, schedule) => [
    '*【RX Business MTG】事前資料のご案内*',
    'RX Business MTG の事前資料を、下記に格納いたしました。',
    'お手数ですが、会議の前にご一読いただけますと幸いです。',
    '',
    '▼ 会議日程',
    schedule || '2026年10月23日（金）9:00〜10:30',
    '',
    '▼ 目的',
    '案件の進捗・受注見込みを確認し、優先順位と次のアクションを決める',
    '',
    '▼ 議題',
    '・更新シートで進捗報告',
    '・重点案件の状況・課題',
    '・質疑・フィードバック',
    '・サービス連携テーマ',
    '',
    '▼ 事前資料',
    link
  ].join('\n');
  const url = 'https://drive.google.com/file/d/pdf-new/view';

  gas.confirmAnswer = 'NO';
  g.notifyMeetingSummaries();
  assert.strictEqual(gas.alerts[gas.alerts.length - 1].message,
    'Google チャット（スペース AAAAtest123）に、次のメッセージをすぐに送ります。よろしいですか？\n\n' + lines('会議サマリ_デモ.pdf\n' + url),
    '確かめの画面には、ファイル名と URL を出す');
  assert.strictEqual(gas.fetches.length, 0, '「いいえ」なら送らない');
  assert.deepStrictEqual(gas.prompts.map(pr => pr.message), [
    '会議日程を入力してください（例：10/23）。\n時刻を省くと「9:00〜10:30」にします。ほかの時刻のときは「10/23 13:00〜14:30」のように入力してください。'
  ], '聞くのは会議日程だけ');

  gas.confirmAnswer = 'YES';
  g.notifyMeetingSummaries();
  assert.deepStrictEqual(gas.fetches.map(f => [f.url, sentText(f)]), [[WEBHOOK, lines('<' + url + '|会議サマリ_デモ.pdf>')]],
    'フォルダの中のいちばん新しい PDF のリンク（ファイル名で表示）');
  assert.strictEqual(gas.alerts[gas.alerts.length - 1].message, 'Google チャットに送りました：会議サマリ_デモ.pdf\n' + url);

  // 会議日程：時刻を書けばその時刻（全角でもよい）。やめたら送らない。読めなければ止める
  gas.promptAnswers = ['１０／２３　１３：００～１４：３０'];
  gas.confirmAnswer = 'NO';
  g.notifyMeetingSummaries();
  assert.ok(gas.alerts[gas.alerts.length - 1].message.indexOf('▼ 会議日程\n' + g.mtgParseDate_('10/23', new Date()).slice(0, 4) +
    '年10月23日（') >= 0 && /）13:00〜14:30\n/.test(gas.alerts[gas.alerts.length - 1].message), gas.alerts[gas.alerts.length - 1].message);
  gas.confirmAnswer = 'YES';
  gas.promptButton = 'CANCEL';
  g.notifyMeetingSummaries();
  assert.strictEqual(gas.fetches.length, 1, '会議日程を聞かれてやめたら送らない');
  gas.promptButton = 'OK';
  gas.promptAnswers = ['あした'];
  assert.throws(() => g.notifyMeetingSummaries(), /「あした」を会議日程として読めません。/);
  assert.strictEqual(gas.fetches.length, 1);
  gas.get('MTG_OPTIONS').chat.meetingTime = '';
  gas.promptAnswers = ['2026/11/6 15:00〜16:00'];
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[1]), /\n▼ 会議日程\n2026年11月6日（金）15:00〜16:00\n/);
  gas.fetches.pop();
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[1]), /\n▼ 会議日程\n2026年10月23日（金）\n/, '時刻の設定が空なら日付だけ');
  gas.fetches.pop();
  gas.get('MTG_OPTIONS').chat.meetingTime = '9:00〜10:30';

  // ゴミ箱に移した PDF は選ばない。PDF が無ければ、いちばん新しいファイル
  gas.trashed.push('pdf-new');
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[1]), /<https:\/\/drive\.google\.com\/file\/d\/pdf-old\/view\|会議サマリ_デモ_旧\.pdf>/);
  gas.trashed.push('pdf-old');
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[2]), /<https:\/\/drive\.google\.com\/file\/d\/slide-1\/view\|会議サマリ_デモ（スライド）>/);

  // ファイルの URL を登録したときは、そのファイル。ドライブ以外の URL はそのまま
  props.setProperty('MTG_CHAT_URL', 'https://drive.google.com/file/d/pdf-other-0001/view?usp=sharing');
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[3]), /\n<https:\/\/drive\.google\.com\/file\/d\/pdf-other-0001\/view\|別の資料\.pdf>$/);
  props.setProperty('MTG_CHAT_URL', 'https://example.com/demo.pdf');
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[4]), /\n<https:\/\/example\.com\/demo\.pdf\|https:\/\/example\.com\/demo\.pdf>$/);

  // 文は MTG_OPTIONS.chat.message で変えられる（{meeting}・{date}・{file}・{url} を置き換える）
  props.setProperty('MTG_CHAT_URL', 'https://drive.google.com/file/d/pdf-other-0001/view');
  gas.get('MTG_OPTIONS').chat.message = ['{meeting}（{date}）：{file}（{url}）'];
  gas.fetchStatuses = [503, 200];
  g.notifyMeetingSummaries();
  assert.strictEqual(gas.fetches.length, 7, '一時的な失敗（5xx）は1回だけやり直す');
  assert.strictEqual(sentText(gas.fetches[6]),
    'RX Business MTG（2026年10月23日（金）9:00〜10:30）：<https://drive.google.com/file/d/pdf-other-0001/view|別の資料.pdf>（https://drive.google.com/file/d/pdf-other-0001/view）');

  gas.fetchStatuses = [404];
  gas.fetchText = '{"error": {"message": "Requested entity was not found."}}';
  assert.throws(() => g.notifyMeetingSummaries(), error => /Google チャットに送れませんでした（HTTP 404）：.*Requested entity/.test(error.message) &&
    !/KEY123|TOKEN456/.test(error.message));
});

test('mtgNotifyScheduled（前の版の自動通知）：トリガーが残っていても何も送らず、そのトリガーを外す', () => {
  const {gas, g} = setup({edits: false});
  const props = g.PropertiesService.getScriptProperties();
  props.setProperty('MTG_CHAT_WEBHOOK', WEBHOOK);
  props.setProperty('MTG_CHAT_URL', DEMO_URL);
  g.ScriptApp.newTrigger('mtgNotifyScheduled').timeBased().everyDays(1).create();
  g.mtgNotifyScheduled();
  assert.deepStrictEqual([gas.fetches.length, gas.triggers.filter(t => t.getHandlerFunction() === 'mtgNotifyScheduled').length], [0, 0]);
});

test('保存先のフォルダ（MTG_OPTIONS.folderId）：決めてあればそのフォルダに保存する（フォルダは作らない）。開けなければ止める', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').keepSlides = true;   // Google スライドも残す設定で確かめる
  const id = gas.get('MTG_OPTIONS').folderId;
  assert.strictEqual(id, '1A0rkoFLICOK4KV25dUTeRdgoOaM0vJEK', '本物の保存先');
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.deepStrictEqual([gas.driveFolders.length, gas.driveFiles.map(f => f.folder), plain(gas.moved).map(m => m.folder)], [0, [id], [id]]);
  assert.match(gas.alerts[gas.alerts.length - 1].message, new RegExp('保存先：https://drive\\.google\\.com/drive/folders/' + id));
  gas.missingFolders = [id];
  gas.promptAnswers = ['1', '10/23'];
  assert.throws(() => g.previewMeetingSummaries(), /保存先のフォルダ（ID：1A0rkoFLICOK4KV25dUTeRdgoOaM0vJEK）を開けません。/);
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

test('addMeetingDummyHistory：シートの値は変えずに、ダミーの変更の記録（DMY-）を足す。資料の変化・停滞に出て【ダミーデータを含む】と出る。入れ直せて、removeMeetingDummyHistory で消せる', () => {
  const {gas, g, sheet} = setup({edits: false});
  const values = JSON.stringify(sheet.getDataRange().getValues());
  gas.promptAnswers = ['1'];
  g.addMeetingDummyHistory();
  assert.strictEqual(JSON.stringify(sheet.getDataRange().getValues()), values, 'シートのセルの値は変えない');
  assert.match(gas.alerts[0].message, /・新FMT（直近14日間の始まり \d{1,2}\/\d{1,2} \d{2}:\d{2} 以降の変化として）\n　E社：前進、A社：停滞、C社：延期、F社：新規、D社：停滞＋活動更新、削除（架空の案件「ダミー商事」）/);
  assert.match(gas.alerts[1].message, /・新FMT：変更の記録 12件・スナップショット 9件/);

  const dummyRows = name => gas.ss.getSheetByName(name).getDataRange().getDisplayValues().slice(1).filter(r => /^DMY-/.test(r[1]));
  assert.ok(dummyRows('変更履歴_差分').every(r => r[2] === 'ダミーデータ' && r[3] === '新FMT'), '編集者は「ダミーデータ」');
  assert.strictEqual(dummyRows('変更時点スナップショット').length, 9);

  // 会議サマリはダミーを数える（サンプル SMP- は数えない）：前回より後の変化・停滞
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.map(c => [label(c), c.category]), [['ダミー商事／（ダミー）削除した案件', 'removed'], ['F社／ブランド刷新', 'new'],
    ['E社／販促DX', 'forward'], ['D社／物流改善', 'activity'], ['C社／店舗什器', 'slipped']]);
  assert.deepStrictEqual(s.changes.filter(c => c.category === 'slipped').map(c => [c.monthFrom, c.monthTo]), [[month(1).text, month(2).text]]);
  assert.deepStrictEqual(s.stalled.map(d => [label(d), d.days, d.touched]), [['A社／アプリ刷新', 100, false], ['D社／物流改善', 70, true]],
    'ステータスを最後に変えた日を 100日・70日前に（D社は前回以降に活動を更新した）');
  assert.strictEqual(s.unmatched, 0, 'どの記録にもスナップショットがある（案件と照合できる）');
  assert.strictEqual(s.dummy, 12);
  const slides = plain(g.mtgSlides_(summarize(g, '新FMT')));
  assert.strictEqual(roleOf(slides[1], 'footer')[0].text, '【ダミーデータを含む】新FMT（10/23 会議資料）　金額の単位：百万円');
  assert.match(slides[0].elements[0].text, /^出典：「新FMT」シートおよび変更履歴（ダミーデータを含む。/);

  // 入れ直し：前に足した分を消してから足す（増えない）
  gas.promptAnswers = ['1'];
  g.addMeetingDummyHistory();
  assert.match(gas.alerts[2].message, /前に足したダミーの記録 12件は消して入れ直します/);
  assert.strictEqual(dummyRows('変更履歴_差分').length, 12);

  // 本番の確かめの画面で知らせる（配る資料に入れないように）
  gas.confirmAnswer = 'NO';
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /※ ダミーの変更の記録（12件）があります。資料に【ダミーデータを含む】と出ます。配る資料なら、先に removeMeetingDummyHistory\(\) で消してください。/);

  // 消す：ダミーの行だけ（ほかの記録は残す）
  gas.confirmAnswer = 'YES';
  g.recordDiffEdit(gas.edit(sheet, 'J5', '10/14 見積を再提出'));   // 本物の変更
  g.removeMeetingDummyHistory();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^ダミーの変更の記録を削除しました：変更の記録 12件・スナップショット 9件。$/);
  assert.deepStrictEqual([dummyRows('変更履歴_差分').length, dummyRows('変更時点スナップショット').length], [0, 0]);
  const after = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual([after.dummy, after.changes.map(c => [label(c), c.category])], [0, [['C社／店舗什器', 'activity']]]);
  g.removeMeetingDummyHistory();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /ダミーの変更の記録（イベントIDが「DMY-」で始まるもの）はありません/);
});

test('サンプル（addSampleHistory()）だけを入れて試作したときは、会議サマリには出ないことと addMeetingDummyHistory() を案内する（基準を変えても出ない）', () => {
  const {gas, g, sheet} = setup({edits: false});
  gas.select(sheet, 'D3');
  g.addSampleHistory();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /（変化のある案件 0・/);
  assert.match(done, /※ サンプル（addSampleHistory\(\)）の変更の記録 5件は、会議サマリには出ません（基準を変えても出ません）。資料の形を確かめるときは、addMeetingDummyHistory\(\) でダミーの変更の記録を足してください。/);
  // ダミーを足せば、変化が出る（案内は出ない）
  gas.promptAnswers = ['1'];
  g.addMeetingDummyHistory();
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  const again = gas.alerts[gas.alerts.length - 1].message;
  assert.doesNotMatch(again, /サンプル（addSampleHistory\(\)）の変更の記録/);
  assert.match(again, /※ ダミーの変更の記録（\d+件）を含みます/);
});

test('addMeetingDummyHistory：記録は差分追跡と同じ形（受注月が日付なら「日時」で dateFormat の文字、想定売上は「数値」、空欄は（空欄）。行の追加・削除は「見出し: 値」）', () => {
  const {gas, g, sheet} = setup({edits: false});
  // 受注月を日付（表示は 2026/9 の形）にする（本物のスーパー・GMS と同じ）
  for (let r = 3; r <= 8; r++) {
    const text = sheet.getRange(r, COL.orderMonth).getDisplayValue();
    if (!text) continue;
    const [y, m] = text.split('/').map(Number);
    sheet.getRange(r, COL.orderMonth).setValue(new gas.Date(Date.UTC(y, m - 1, 1) - 9 * 3600000)).setNumberFormat('yyyy/M');
  }
  gas.promptAnswers = ['1'];
  g.addMeetingDummyHistory();
  const log = gas.ss.getSheetByName('変更履歴_差分').getDataRange().getDisplayValues().slice(1).filter(r => /^DMY-/.test(r[1]));
  const slip = log.find(r => r[5] === '受注月');
  const c = month(2);
  const p = month(1);
  assert.deepStrictEqual(slip.slice(3, 9), ['新FMT', 'I5', '受注月', '日時 -> 日時 / ダミー',
    p.y + '/' + String(p.m).padStart(2, '0') + '/01 00:00', c.y + '/' + String(c.m).padStart(2, '0') + '/01 00:00'], '差分追跡と同じ dateFormat の文字');
  assert.ok(log.some(r => r[5] === '活動状況' && r[6] === '空欄 -> 文字列 / ダミー' && r[7] === '（空欄）'), '空欄は（空欄）');
  const born = log.find(r => r[6] === '行追加');
  assert.deepStrictEqual([born[4], born[5], born[7]], ['8行', '（行全体）', '（空欄）']);
  assert.match(born[8], /^得意先: F社\n案件名: ブランド刷新\n案件ステータス: 提案中\n想定売上規模（百万）: 10\n/, '追跡範囲の列だけを「見出し: 値」で');
  assert.ok(!/最終更新日時/.test(born[8]), '最終更新日時の列は書かない');
  const removed = log.find(r => r[6] === '行削除');
  assert.match(removed[7], /^得意先: ダミー商事\n案件名: （ダミー）削除した案件\n案件ステータス: 初回訪問\n想定売上規模（百万）: 20\n受注月: \d{4}\/\d{2}\/01 00:00\nアカウント責任者: ダミー 担当$/);
  // スナップショットの受注月は、シートの表示の形（2026/9 → 前回は 2026/8。0 を付けない）
  const snaps = gas.ss.getSheetByName('変更時点スナップショット').getDataRange().getDisplayValues().slice(1)
    .filter(r => /^DMY-/.test(r[1]) && r[4] === '5');
  const months = snaps.map(r => JSON.parse(r[7]).find(([k]) => k === '受注月')[1]);
  assert.deepStrictEqual(months, [p.y + '/' + p.m, c.y + '/' + c.m], '前回の時点（1か月前）→ 今');
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.changes.filter(ch => label(ch) === 'C社／店舗什器').map(ch => [ch.category, ch.monthFrom, ch.monthTo]),
    [['slipped', p.y + '/' + p.m, c.y + '/' + c.m]]);
});

test('addMeetingDummyHistory：前回より後の本物の変化は隠さない（ダミーの前回の時点の値は、本物の変更前の値）。日付の古いダミーの行があっても、本物の記録を読む', () => {
  const {gas, g} = setup();   // A社 前進・C社 失注・D社 活動の更新・F社 削除（本物の記録）
  gas.promptAnswers = ['1'];
  g.addMeetingDummyHistory();
  const s = plain(summarize(g, '新FMT'));
  const of = name => s.changes.filter(c => label(c) === name).map(c => c.category);
  assert.deepStrictEqual([of('A社／アプリ刷新'), of('C社／店舗什器'), of('D社／物流改善'), of('F社／ブランド刷新')], [['forward'], ['lost'], ['activity'], ['removed']]);
  assert.deepStrictEqual(s.stalled.map(d => [label(d), d.days, d.touched]), [['D社／物流改善', 100, true]],
    '本物の記録でステータスを変えた案件（A社）は停滞にしない（それより前の日付にしても停滞にならない）。D社は前回以降に活動を更新した');
});

test('setMeetingBaseline：今を基準として記録し（はいのときだけ）、次の資料は基準より後の変化だけを出す。基準より前の記録の停滞は「N日以上」', () => {
  const {gas, g, sheet} = setup();   // A社 前進・C社 失注・D社 活動の更新・F社 削除（どれも基準より前）
  g.recordDiffEdit(gas.edit(sheet, 'D8:I8', [['G社', 'POS連携', '提案中', '20', '6', month(2).text]]));   // 移行で貼り付けた案件
  const props = g.PropertiesService.getDocumentProperties();
  const lastPrompt = () => gas.prompts[gas.prompts.length - 1].message;
  const lastAlert = () => gas.alerts[gas.alerts.length - 1];

  gas.confirmAnswer = 'NO';
  gas.promptAnswers = ['1'];
  g.setMeetingBaseline();
  assert.match(lastPrompt(), /^会議サマリの基準を記録します（会議日は記録しません）。\nシートの番号を入力してください/);
  assert.match(lastAlert().message, /^「新FMT」に、今（\d{2}\/\d{2} \d{2}:\d{2}）を基準として記録します。/);
  assert.strictEqual(lastAlert().buttons, 'YES_NO');
  assert.strictEqual(props.getProperty('MTG_MEETINGS_新FMT'), null, '「いいえ」なら記録しない');

  gas.confirmAnswer = 'YES';
  gas.promptAnswers = ['1'];
  g.setMeetingBaseline();
  assert.match(lastAlert().message, /「新FMT」に、基準（\d{2}\/\d{2} \d{2}:\d{2}）を記録しました。/);
  const saved = JSON.parse(props.getProperty('MTG_MEETINGS_新FMT'));
  assert.deepStrictEqual(saved.map(m => [m.date, m.baseline]), [[day(0), true]]);
  assert.strictEqual(props.getProperty('MTG_MEETINGS_新FMT2'), null, '選んだシートだけ');
  const cutoff = saved[0].cutoff;

  // 停滞の日数：基準より前に変わったステータス（貼り付けた行も）は「N日以上」
  const focus = options => plain(summarize(g, '新FMT', options)).focus.find(d => d.customer === 'A社');
  assert.deepStrictEqual([focus({}).days, focus({}).atLeast], [0, false]);
  assert.deepStrictEqual([focus({baseline: cutoff}).days, focus({baseline: cutoff}).atLeast], [0, true]);

  while (Date.now() <= cutoff) { /* 基準と同じ時刻の記録にしない */ }
  g.recordDiffEdit(gas.edit(sheet, 'F7', '受注見込み'));   // E社：基準のあとに前進
  const meeting = day(7);
  gas.promptAnswers = ['1', meeting];
  g.createMeetingSummaries();
  assert.match(gas.prompts[gas.prompts.length - 1].message, /・新FMT：記録なし（\d{2}\/\d{2} \d{2}:\d{2} に基準を記録）$/);
  assert.match(lastAlert().message, /「新FMT」：.*（変化のある案件 1・/, '基準より前の変化（貼り付けた G社 など）は数えない');
  const inserted = gas.decks[0].requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted, /\n対象期間：基準時点（\d{2}\/\d{2} \d{2}:\d{2}）以降\n/, '表紙の期間');
  assert.match(gas.alerts.filter(a => /本番の資料を作り/.test(a.message)).pop().message, /・新FMT：基準（\d{1,2}\/\d{1,2} \d{2}:\d{2}）以降/,
    '本番の確認：基準からの変化を出すこと');
  assert.match(inserted, /期待値合計 \d+百万円（基準比 /, '前回比ではなく基準比');
  assert.doesNotMatch(inserted, /前回比|前回以降/, '前回の会議の記録が無いので「前回」とは書かない');
  assert.match(inserted, /\n停滞（同一ステータス45日以上）\n—\n—\nデータなし：基準（\d{1,2}\/\d{1,2}）から 45日未満のため、まだ数えられない（\d{1,2}\/\d{1,2} から数える）\n/,
    '基準から45日たつまでは、停滞を数えられないことを書く（0件を「停滞なし」と読ませない）');
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).map(m => [m.date, !!m.baseline]), [[day(0), true], [meeting, false]],
    '会議を記録しても基準は残す');

  // 次の会議は、前回の会議の資料のほうが基準より新しいので、前回の会議から
  const made = JSON.parse(props.getProperty('MTG_MEETINGS_新FMT'))[1].cutoff;
  while (Date.now() <= made) { /* 同じ時刻の記録にしない */ }
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注見込み'));   // A社：前進
  gas.promptAnswers = ['1', day(14)];
  g.createMeetingSummaries();
  assert.match(gas.prompts[gas.prompts.length - 1].message, new RegExp('・新FMT：' + meeting + '$'), '前回の会議より前の基準は出さない');
  assert.match(lastAlert().message, /（変化のある案件 1・/);
  const second = gas.decks[1].requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(second, /対象期間：前回会議（\d+\/\d+）の資料作成時（/);

  // もう一度記録すると、基準を今に置き換える（基準は1つ）
  gas.promptAnswers = ['1'];
  g.setMeetingBaseline();
  const list = JSON.parse(props.getProperty('MTG_MEETINGS_新FMT'));
  assert.strictEqual(list.filter(m => m.baseline).length, 1);
  assert.ok(list.find(m => m.baseline).cutoff > cutoff);
});

/* ---------------- 統合版（全サブインダストリー） ---------------- */

test('統合版のスライド：表紙・目次・全サブインダストリーの状況のあとに、各サブインダストリーの資料をそのまま続ける（ページ番号・扉の目次は通し）', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const T = gas.get('MTG_TITLES');
  const summaries = [summarize(g, '新FMT'), summarize(g, '新FMT2')];
  const one = plain(g.mtgSlides_(summaries[0]));
  const slides = plain(g.mtgCombinedSlides_(summaries, '2026/10/23', true));
  assert.strictEqual(slides.length, 3 + one.length + plain(g.mtgSlides_(summaries[1])).length);
  assert.deepStrictEqual(slides[0].template, {title: '全サブインダストリー', subtitle: 'RX Business MTG　事前資料', date: '2026年10月23日（金）', name: ''});
  assert.match(slides[0].elements[0].text, /^収録：新FMT・新FMT2（2サブインダストリー）$/);
  assert.deepStrictEqual([slides[0].combined, slides[0].cover], [true, true]);
  assert.deepStrictEqual(slides.slice(3, 3 + one.length).map(sl => sl.title), one.map(sl => sl.title), '各サブインダストリーの資料はそのまま');
  const start2 = 4 + one.length;
  assert.deepStrictEqual([slides[3].template.title, slides[start2 - 1].template.title], ['新FMT', '新FMT2'], '各資料の表紙が区切り');
  // 目次：全サブインダストリーの状況（3ページ）と各資料の表紙のページ。名前とページ番号は、そのスライドへのリンク（スライドの番号は 0 から）
  const contents = slides[1];
  assert.deepStrictEqual([contents.title, leadOf(contents)], ['目次', '名前かページ番号を押すと、そのページに移ります。' +
    '※ ブラウザの Google ドライブのプレビュー画面ではリンクは動きません（ダウンロードして Acrobat などで開いてください）。']);
  assert.ok(roleOf(contents, 'lead').every(el => el.text.indexOf('\n') < 0) && roleOf(contents, 'lead').length === 2, 'リード文は2行（補足は1行に入る）');
  assert.deepStrictEqual(roleOf(contents, 'toc').map(el => [el.text, el.link]), [
    ['全サブインダストリーの状況', 2], ['3', 2], ['1.　新FMT', 3], ['4', 3], ['2.　新FMT2', start2 - 1], [String(start2), start2 - 1]]);
  const eleven = Array.from({length: 11}, (_, i) => ({summary: Object.assign({}, summaries[0], {name: 'サブインダストリー' + (i + 1)}), start: 4 + i * 10}));
  const twelve = plain(g.mtgContentsSlide_(eleven)).elements;
  assert.strictEqual(new Set(twelve.map(el => el.x)).size, 4, '9件以上なら2列（名前とページ番号で4つの位置）');
  assert.ok(twelve.every(el => el.y >= S.top && el.y + el.h <= S.bottom), '12件がページに入る');
  // ページ番号は通し（表紙・扉には付けない）
  slides.forEach((sl, i) => {
    const page = roleOf(sl, 'page');
    if (sl.cover || sl.divider) assert.strictEqual(page.length, 0);
    else assert.strictEqual(page[0].text, String(i + 1), (i + 1) + '枚目');
  });
  assert.strictEqual(roleOf(slides[2], 'footer')[0].text, '全サブインダストリー（10/23 会議資料）　金額の単位：百万円');
  assert.strictEqual(roleOf(slides[start2], 'footer')[0].text, '新FMT2（10/23 会議資料）　金額の単位：百万円');
  // 2つ目の資料の扉の目次は、この資料の中のページ
  const divider = slides.findIndex((sl, i) => i >= start2 && sl.divider);
  const toc = roleOf(slides[divider], 'toc').map(el => el.text);
  assert.deepStrictEqual([toc[0], toc[2]], ['Appendix ①', String(divider + 2)]);

  // 全サブインダストリーの状況：サブインダストリーごとの数と、その資料の表紙のページ。合計
  const overview = slides[2];
  assert.strictEqual(overview.title, T.overview);
  const table = overview.elements.find(el => el.type === 'table');
  const text = table.rows.map(r => r.cells.map(cell => cell.text.replace(/\n/g, ' ')));
  assert.deepStrictEqual(table.rows[0].cells.map(cell => cell.text.replace(/\n/g, '')), ['サブインダストリー', '進行中案件', '期待値合計', '直近14日間の増減', '変化のある案件', '停滞', '確認事項', 'ページ'],
    '増減の見出しは、いちばん多い比べ方（どれも初回なので直近14日間）');
  const k = summaries.map(s => plain(s.kpi));
  const b = summaries.map(s => plain(s.bridge));
  const signed = v => g.mtgSigned_(v);
  // 停滞をまだ数えられないシートは「データなし」（どのシートも数えられなければ合計も）
  assert.ok(summaries.every(s => s.first && s.gaps.stalled));
  assert.deepStrictEqual(text.slice(1), [
    ['新FMT', k[0].active + '件', g.mtgAmount_(b[0].end), signed(b[0].end - b[0].start), k[0].changed + '件', 'データなし', k[0].alerts + '項目', '4'],
    ['新FMT2', k[1].active + '件', g.mtgAmount_(b[1].end), signed(b[1].end - b[1].start), k[1].changed + '件', 'データなし', k[1].alerts + '項目', String(start2)],
    ['合計', (k[0].active + k[1].active) + '件', g.mtgAmount_(b[0].end + b[1].end), signed(b[0].end - b[0].start + b[1].end - b[1].start),
      (k[0].changed + k[1].changed) + '件', 'データなし', (k[0].alerts + k[1].alerts) + '項目', '']
  ]);
  assert.match(leadOf(overview), /^2サブインダストリーの進行中案件 \d+件、期待値合計 [\d,.]+百万円。$/);
  assert.strictEqual(roleOf(overview, 'footer')[0].text, '全サブインダストリー（10/23 会議資料）　金額の単位：百万円', '※ が無ければ補足なし');
  // 比べ方が違うシートは、セルの1行目に何との比か
  const held = Object.assign({}, summaries[1], {first: false, previous: {date: '2026/10/09', cutoff: 1}});
  const mixed = plain(g.mtgOverviewSlide_([{summary: summaries[0], start: 3}, {summary: held, start: 13}, {summary: summaries[0], start: 23}]));
  assert.strictEqual(mixed.elements[0].rows[2].cells[3].text, '前回比\n' + signed(b[1].end - b[1].start));
  assert.ok(table.y + table.rows.reduce((h, r) => h + r.h, 0) <= S.bottom, 'ページに入る');

  // 作れない値（列・記録が無い）は灰色の「データなし」、合計は作れたシートの分だけ（※）
  const lacking = Object.assign({}, summaries[1], {gaps: Object.assign({}, summaries[1].gaps, {weighted: '列が無い', changes: '記録が無い'})});
  const t2 = plain(g.mtgCombinedSlides_([summaries[0], lacking], '2026/10/23', false))[2];
  const rows2 = t2.elements.find(el => el.type === 'table').rows;
  assert.deepStrictEqual(rows2[2].cells.slice(2, 5).map(cell => [cell.text, cell.fill]), [['データなし', '#F2F2F2'], ['データなし', '#F2F2F2'], ['データなし', '#F2F2F2']]);
  assert.deepStrictEqual(rows2[3].cells.slice(2, 5).map(cell => cell.text), [g.mtgAmount_(b[0].end) + '※', signed(b[0].end - b[0].start) + '※', k[0].changed + '件※']);
  assert.match(leadOf(t2), /（期待値の列が無いシートを除く）。$/);
  assert.strictEqual(roleOf(t2, 'footer')[0].text, '【試作】全サブインダストリー（10/23 会議資料）　金額の単位：百万円　※ は「データなし」のシートを除いた合計');
  assert.strictEqual(rows2[3].cells[5].text, 'データなし', '停滞はどのシートも数えられないので、合計も「データなし」');
  assert.strictEqual(plain(g.mtgCombinedSlides_(summaries, '2026/10/23', false))[0].template.subtitle, 'RX Business MTG　事前資料【試作】');

  // 11のサブインダストリーでも1ページに入る（文字を小さくする）
  const many = Array.from({length: 11}, (_, i) => Object.assign({}, summaries[0], {name: 'サブインダストリー' + (i + 1)}));
  many[0] = Object.assign({}, many[0], {gaps: {weighted: '列が無い'}});   // リード文が長く、合計に ※ があっても
  const big = plain(g.mtgOverviewSlide_(many.map((s, i) => ({summary: s, start: 3 + i * 10}))));
  const bigTable = big.elements.find(el => el.type === 'table');
  assert.strictEqual(bigTable.rows.length, 13);
  assert.ok(bigTable.y + bigTable.rows.reduce((h, r) => h + r.h, 0) <= S.bottom, '11行と合計がページに入る');
});

test('統合版：createMeetingSummaries() では作らず、すべてのシートを作ったら createCombinedMeetingSummary() を案内する。統合版は会議のフォルダに格納してインデックスの「全サブインダストリー」の列に書く', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const meeting = g.mtgParseDate_('10/23', new Date());
  const parent = meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  Object.assign(o.combined, {enabled: true, meetingFolder: parent});
  gas.promptAnswers = ['1', '10/23'];
  g.previewMeetingSummaries();
  assert.doesNotMatch(gas.alerts[gas.alerts.length - 1].message, /統合版/, '試作では案内しない');
  gas.promptAnswers = ['', '10/23'];
  g.createMeetingSummaries();
  const file = meeting.replace(/\//g, '');
  assert.deepStrictEqual(gas.decks.slice(1).map(d => d.title), [file + '_新FMT_RX Business Meeting用資料', file + '_新FMT2_RX Business Meeting用資料'],
    '統合版は作らない（関数を分ける）');
  const run = gas.alerts[gas.alerts.length - 1].message;
  assert.match(run, /この会議日の本番の資料が、すべてのサブインダストリーでそろいました。統合版（全サブインダストリー）は、エディタで createCombinedMeetingSummary\(\) を実行して作ってください（会議日を入力します）。/);
  assert.ok(run.indexOf('統合版') < run.indexOf('保存先：'));

  gas.promptAnswers = ['10/23'];
  g.createCombinedMeetingSummary();
  const combined = gas.decks[gas.decks.length - 1];
  assert.strictEqual(combined.title, file + '_全サブインダストリー_RX Business Meeting用資料');
  assert.strictEqual(combined.requests.filter(r => r.createSlide).length, 3 + 10 + 10, '表紙・目次・状況と、2つの資料（10枚ずつ）');
  // 目次のリンク：すべてのスライドを作ったあとに、リンク先のスライドの ID で付ける
  // 目次のリンク：名前もページ番号も、この資料のそのページ
  const links = combined.requests.filter(r => r.updateTextStyle && r.updateTextStyle.fields === 'link,underline,foregroundColor')
    .map(r => [r.updateTextStyle.objectId, r.updateTextStyle.style.link.pageObjectId]);
  assert.deepStrictEqual(plain(links).map(l => l[1]), ['mtg_p3', 'mtg_p3', 'mtg_p4', 'mtg_p4', 'mtg_p14', 'mtg_p14']);
  assert.ok(plain(links).every(l => /^mtg_p2_\d+$/.test(l[0])), '目次（2枚目）の部品');
  const linkStyle = combined.requests.find(r => r.updateTextStyle && r.updateTextStyle.style.link).updateTextStyle.style;
  assert.deepStrictEqual(plain([linkStyle.underline, linkStyle.foregroundColor]), [false, {opaqueColor: {rgbColor: plain(g.mtgRgb_('#1A1A1A'))}}],
    'リンクにしても、字の色はそのまま・下線なし');
  const lastCreate = combined.requests.map(r => !!r.createSlide).lastIndexOf(true);
  assert.ok(combined.requests.findIndex(r => r.updateTextStyle && r.updateTextStyle.fields === 'link,underline,foregroundColor') > lastCreate, 'リンク先ができてから');
  assert.ok(combined.batches.length >= 2, '何枚かずつまとめて送る');
  assert.strictEqual(gas.slideBackgrounds.filter(b => b.deck === combined.id).length, 5, '表紙と、各資料の表紙・扉（ひな形なし）');
  const pdf = gas.driveFiles.find(f => f.name === file + '_全サブインダストリー_RX Business Meeting用資料.pdf');
  assert.strictEqual(pdf.folder, 'meeting-folder-0', '会議のフォルダに格納');
  const index = gas.ss.getSheetByName('INDEX_会議資料');
  assert.deepStrictEqual(plain(index.getRange(1, 3, 2, 2).getDisplayValues()), [['全サブインダストリー', '新FMT'],
    [file + '_全サブインダストリー_RX Business Meeting用資料.pdf', file + '_新FMT_RX Business Meeting用資料.pdf']]);
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, new RegExp('「全サブインダストリー」（統合版）：' + file + '_全サブインダストリー_RX Business Meeting用資料\\.pdf（23ページ）\\n　https://drive\\.google\\.com/file/d/' +
    pdf.id + '/view\\n　会議のフォルダ「\\d{6}_第5回会議」に格納しました'));

  // rebuildMeetingIndex は統合版も書き直す
  index.clear();
  g.rebuildMeetingIndex();
  assert.strictEqual(index.getRange(2, 3).getDisplayValue(), file + '_全サブインダストリー_RX Business Meeting用資料.pdf');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /・全サブインダストリー：1件/);
});

test('createCombinedMeetingSummary：すべてのシートに会議日の本番の資料（記録と PDF）があるときだけ作る（無ければ何も作らずに終える）。会議日・コピーは記録しない', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const meeting = g.mtgParseDate_('10/23', new Date());
  const parent = meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  Object.assign(o.combined, {enabled: true, meetingFolder: parent});
  const props = g.PropertiesService.getDocumentProperties();

  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();   // 新FMT だけ本番
  const decks0 = gas.decks.length;
  const alerts0 = gas.alerts.length;
  gas.promptAnswers = ['10/23'];
  g.createCombinedMeetingSummary();
  assert.strictEqual(gas.alerts.length, alerts0 + 1, '確かめの画面も出さずに終える');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^統合版を作れません。会議日 \d{4}\/10\/23 の本番の資料が、まだできていないサブインダストリーがあります（1件）：\n・新FMT2\n\ncreateMeetingSummaries\(\) で、すべてのサブインダストリーの本番の資料を作ってから（「INDEX_会議資料」のその会議日の行に空のセルが無くなってから）、もう一度実行してください。$/);
  assert.strictEqual(gas.decks.length, decks0, '試作も作らない');
  // 本番の記録があっても、PDF が無ければ作らない（消した・名前を変えたなど）
  gas.promptAnswers = ['2', '10/23'];
  g.createMeetingSummaries();   // 新FMT2 も本番
  const pdf = gas.driveFiles.find(f => /_新FMT2_RX Business Meeting用資料\.pdf$/.test(f.name));
  gas.trashed.push(pdf.id);
  gas.promptAnswers = ['10/23'];
  g.createCombinedMeetingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^統合版を作れません。[\s\S]*（1件）：\n・新FMT2\n/);
  gas.trashed.pop();

  const before = [props.getProperty('MTG_MEETINGS_新FMT'), props.getProperty('MTG_MEETINGS_新FMT2')];
  const copies = gas.created.length;
  gas.promptAnswers = ['10/23'];
  g.createCombinedMeetingSummary();
  assert.match(gas.alerts[gas.alerts.length - 2].message, /会議日 \d{4}\/10\/23 の本番の資料として作り、会議のフォルダに格納します/);
  assert.match(gas.decks[gas.decks.length - 1].title, /\d{8}_全サブインダストリー_RX Business Meeting用資料$/);
  assert.ok(gas.driveFiles.some(f => /_全サブインダストリー_RX Business Meeting用資料\.pdf$/.test(f.name) && f.folder === 'meeting-folder-0'));
  assert.deepStrictEqual([props.getProperty('MTG_MEETINGS_新FMT'), props.getProperty('MTG_MEETINGS_新FMT2')], before, '会議日の記録は変えない');
  assert.strictEqual(gas.created.length, copies, '会議時点のコピーは作らない');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「全サブインダストリー」（統合版）：.*（23ページ）/);

  // 「いいえ」なら作らない。使わない設定なら止める
  const decks = gas.decks.length;
  gas.promptAnswers = ['10/23'];
  gas.confirmAnswer = 'NO';
  g.createCombinedMeetingSummary();
  assert.strictEqual(gas.decks.length, decks);
  o.combined.enabled = false;
  g.createCombinedMeetingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /統合版を使わない設定です/);
  assert.strictEqual(gas.decks.length, decks);
});

test('統合版・ひな形：統合版の表紙も各資料の表紙もひな形の1枚目（緑）、扉は2枚目を使い、2回目からは複製して、最後に決まった順に並べ直す（3枚目は使わない）', () => {
  const {gas, g} = setup();
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789';
  gas.slideTemplates[id] = pptxTemplate();
  g.PropertiesService.getScriptProperties().setProperty('MTG_TEMPLATE_URL', 'https://docs.google.com/presentation/d/' + id + '/edit');
  const summaries = [summarize(g, '新FMT'), summarize(g, '新FMT2')];
  const slides = g.mtgCombinedSlides_(summaries, '2026/10/23', true);
  g.mtgExportDeck_('統合', slides, g.DriveApp.getFolderById('folder-x'));
  const deck = gas.decks[gas.decks.length - 1];
  const r = plain(deck.requests);
  const kinds = slides.map(sl => (sl.cover ? 'cover' : sl.divider ? 'divider' : ''));
  const covers = kinds.map((k, i) => (k === 'cover' ? i : -1)).filter(i => i >= 0);
  const dividers = kinds.map((k, i) => (k === 'divider' ? i : -1)).filter(i => i >= 0);
  assert.deepStrictEqual(covers.length, 3, '統合版・新FMT・新FMT2 の表紙');
  assert.deepStrictEqual(r.filter(q => q.duplicateObject).map(q => [q.duplicateObject.objectId, q.duplicateObject.objectIds[q.duplicateObject.objectId]]),
    [['g_cover', 'mtg_t' + (covers[1] + 1)], ['g_cover', 'mtg_t' + (covers[2] + 1)], ['g_divider', 'mtg_t' + (dividers[1] + 1)]],
    '統合版の表紙が1枚目そのもの。各資料の表紙と2つ目の扉は複製');
  assert.deepStrictEqual(r.filter(q => q.deleteObject && /^g_/.test(q.deleteObject.objectId)).map(q => q.deleteObject.objectId), ['g_combined'], '紫の3枚目は使わない');
  const expected = slides.map((sl, i) => (i === 0 ? 'g_cover' : i === dividers[0] ? 'g_divider' : covers.indexOf(i) > 0 || i === dividers[1] ? 'mtg_t' + (i + 1)
    : 'mtg_p' + (i + 1)));
  const moves = r.filter(q => q.updateSlidesPosition).map(q => [q.updateSlidesPosition.slideObjectIds[0], q.updateSlidesPosition.insertionIndex]);
  assert.deepStrictEqual(moves, plain(expected).map((page, i) => [page, i]));
  // 複製したページも、枠ごとに書き換える（複製の部品の ID はページの ID から）
  const title = page => r.filter(q => q.insertText && /_e\d+$|_title$/.test(q.insertText.objectId) && q.insertText.objectId.indexOf(page + '_') === 0)
    .map(q => q.insertText.text)[0];
  assert.deepStrictEqual(['g_cover', 'mtg_t' + (covers[1] + 1), 'mtg_t' + (covers[2] + 1), 'g_divider', 'mtg_t' + (dividers[1] + 1)].map(title),
    ['全サブインダストリー', '新FMT', '新FMT2', 'Appendix', 'Appendix']);
  assert.deepStrictEqual(gas.slideBackgrounds, [], '表紙・扉はどれもひな形のページ');
  // 目次のリンクは、ひな形のページ（複製）の ID に
  const links = r.filter(q => q.updateTextStyle && q.updateTextStyle.fields === 'link,underline,foregroundColor').map(q => q.updateTextStyle.style.link.pageObjectId);
  assert.deepStrictEqual(links, ['mtg_p3', 'mtg_p3', 'mtg_t' + (covers[1] + 1), 'mtg_t' + (covers[1] + 1), 'mtg_t' + (covers[2] + 1), 'mtg_t' + (covers[2] + 1)]);
  // 統合版は 130ページほどになるので、資料を丸ごと読まない（応答が大きすぎて 413 で失敗した）。どの読み方も項目を絞る
  assert.ok(gas.slidesGets.length >= 3 && gas.slidesGets.every(r => r.fields && /^(presentationId,pageSize,slides\/objectId|slides[(\/])/.test(r.fields)),
    JSON.stringify(gas.slidesGets.map(r => r.fields)));
});

/* ---------------- 6分の上限：区切って続ける ---------------- */

/** 資料1つに 100秒かかる時計（mtgNow_ を差し替える） */
function slowClock(g) {
  const clock = {t: 0};
  g.mtgNow_ = () => clock.t;
  const exportDeck = g.mtgExportDeck_;
  g.mtgExportDeck_ = (...args) => { clock.t += 100000; return exportDeck(...args); };
  return clock;
}

test('6分の上限：次の仕事が上限を越えそうなら区切り、続きを1回きりのトリガーで自動で続ける（continueMeetingSummaries）。結果は showMeetingSummariesResult() で出す', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const meeting = g.mtgParseDate_('10/23', new Date());
  meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  o.run.limitMs = 205000;   // 資料2つ（200秒）まで。会議時点のコピー（見積もり 60秒）からは次へ
  slowClock(g);
  const handlers = () => gas.triggers.filter(t => t.getHandlerFunction() === 'continueMeetingSummaries');
  gas.promptAnswers = ['', '10/23'];
  g.createMeetingSummaries();

  const first = gas.alerts[gas.alerts.length - 1].message;
  assert.match(first, /^時間がかかるため、ここで区切りました（Apps Script の1回の実行は6分まで）。残りは約1分後に自動で続けます：会議時点のコピーと会議日の記録・会議のフォルダへの格納とインデックス。/);
  assert.match(first, /「新FMT」：\d{8}_新FMT_RX Business Meeting用資料\.pdf/);
  assert.deepStrictEqual(handlers().map(t => t.options.after), [60000], '続きのトリガーは1つ（1回きり）');
  const sprops = g.PropertiesService.getScriptProperties();
  assert.ok(sprops.getProperty('MTG_JOB'));
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT'), null, '会議日はまだ記録しない（資料を全部作れてから）');
  assert.ok(!gas.driveFiles.some(f => f.folder === 'meeting-folder-0'), 'まだ格納していない');

  // 途中で新しく作ろうとしたら聞く（いいえなら何もしない）。統合版も、終わるまで作らない
  gas.confirmAnswer = 'NO';
  g.createMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /自動で続けている途中です。[\s\S]*それをやめて、新しく作りますか？/);
  assert.strictEqual(gas.decks.length, 2);
  o.combined.enabled = true;
  g.createCombinedMeetingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^サブインダストリーの資料を作っている途中です/);
  g.showMeetingSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /作っている途中です（会議日 \d{4}\/10\/23。資料 2\/2件・格納 0\/2件）。自動で続けています。/);
  gas.confirmAnswer = 'YES';

  // 続き（トリガーから）：会議時点のコピー・会議日の記録（作り始めた時刻で）・格納・インデックス。知らせを保存して、仕事とトリガーを消す
  g.continueMeetingSummaries();
  assert.ok(gas.driveFiles.some(f => f.folder === 'meeting-folder-0'), '格納した');
  const saved = JSON.parse(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT'));
  assert.deepStrictEqual(saved.map(m => m.date), [meeting]);
  assert.ok(saved[0].copy && saved[0].copy.id);
  assert.deepStrictEqual([sprops.getProperty('MTG_JOB'), handlers().length], [null, 0]);
  g.showMeetingSummariesResult();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /^\d{4}\/\d{2}\/\d{2} \d{2}:\d{2} に終わりました。\n\n「新FMT」：/);
  assert.match(done, /会議のフォルダ「\d{6}_第5回会議」に格納しました/);
  assert.match(done, /「INDEX_会議資料」シートに、資料のファイルを書きました/);
  assert.match(done, /会議日（\d{4}\/10\/23）を記録しました。/);
  assert.ok(gas.logs.some(l => /に終わりました/.test(l.line)), '実行ログにも出す');
});

test('6分の上限：資料の途中でも区切る（作った資料は作り直さない）。続きで止まったら、知らせを残して仕事を消す', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  o.run.limitMs = 130000;   // 資料1つ（100秒）＋次の資料の見積もり（35秒）は入らない
  slowClock(g);
  gas.promptAnswers = ['', '10/23'];
  g.previewMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /残りは約1分後に自動で続けます：資料 1件（新FMT2）。/);
  assert.strictEqual(gas.decks.length, 1);
  g.continueMeetingSummaries();
  assert.deepStrictEqual(gas.decks.map(d => d.title.replace(/^\d{8}_/, '')), ['新FMT_RX Business Meeting用資料_試作', '新FMT2_RX Business Meeting用資料_試作']);
  g.showMeetingSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：[\s\S]*「新FMT2」：[\s\S]*試作のため、会議日は記録していません。/);

  // 続きで止まった：知らせを残して仕事を消す（作り直しを案内する）
  gas.promptAnswers = ['', '10/23'];
  g.previewMeetingSummaries();
  g.mtgExportDeck_ = () => { throw new Error('Slides API の失敗'); };
  g.continueMeetingSummaries();
  assert.strictEqual(g.PropertiesService.getScriptProperties().getProperty('MTG_JOB'), null);
  g.showMeetingSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /途中で止まりました：Slides API の失敗\ncreateMeetingSummaries\(\)（試作なら previewMeetingSummaries\(\)）で、作り直してください。/);
});

test('mtgPdfFitLinks_：PDF のページへのリンクの飛び先「/XYZ null null 0」を、同じ長さの「/Fit」に置き換える（ほかの位置は変えない）。無ければそのまま', () => {
  const {gas, g} = setup({edits: false});
  const pdf = '%PDF-1.4\n23\n0\nobj\n[\n<<\n/Type\n/Annot\n/Subtype\n/Link\n/Dest\n[\n25\n0\nR\n/XYZ\nnull\nnull\n0\n]\n>>\n]\nendobj\nxref\n0000000009 00000 n \n\xe3\x81\x82%%EOF';
  const bytes = Array.from(Buffer.from(pdf, 'latin1'), b => (b > 127 ? b - 256 : b));
  const blob = {getBytes: () => bytes.slice(), getName: () => '統合.pdf'};
  const out = g.mtgPdfFitLinks_(blob);
  const text = Buffer.from(out.getBytes().map(b => b & 255)).toString('latin1');
  assert.strictEqual(text.length, pdf.length, '長さは同じ');
  assert.ok(text.indexOf('/Dest\n[\n25\n0\nR\n/Fit            \n]') > 0, JSON.stringify(text));
  assert.strictEqual(text.replace('/Fit            ', '/XYZ\nnull\nnull\n0'), pdf, 'ほかは変えない（日本語のバイトも）');
  assert.deepStrictEqual([out.getName(), out.getContentType()], ['統合.pdf', 'application/pdf']);
  const plainPdf = {getBytes: () => Array.from(Buffer.from('%PDF-1.4\n%%EOF', 'latin1')), getName: () => 'x.pdf'};
  assert.strictEqual(g.mtgPdfFitLinks_(plainPdf), plainPdf);
  // 統合版を作るとき（リンクのある資料）だけ置き換える
  gas.get('MTG_OPTIONS').combined.enabled = true;
  gas.pdfBytes = bytes;
  const slides = g.mtgCombinedSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/10/23', true);
  const file = g.mtgExportDeck_('統合', slides, g.DriveApp.getFolderById('folder-x')).pdf;
  const saved = gas.driveFiles.find(f => f.id === file.getId()).blob;
  assert.ok(Buffer.from(saved.getBytes().map(b => b & 255)).toString('latin1').indexOf('/Fit') > 0);
});

test('Appendix のページはグレー（扉に合わせる）：表の見出し・1列目・大きな数字をグレーに。区分・該当事項のラベルは意味の色のまま。本編は緑', () => {
  const {gas, g} = setup();
  const T = gas.get('MTG_TITLES');
  const c = gas.get('MTG_OPTIONS').colors;
  const slides = plain(g.mtgSlides_(summarize(g, '新FMT')));
  const fills = slide => [].concat(...slide.elements.filter(el => el.type === 'table').map(t => [].concat(...t.rows.map(r => r.cells.map(cell => cell.fill)))));
  const appendix = slides.filter(sl => /^Appendix /.test(sl.title));
  assert.ok(appendix.length >= 3);
  appendix.forEach(sl => {
    const f = fills(sl);
    assert.ok(f.indexOf('#7F7F7F') >= 0 && f.indexOf(c.green) < 0 && f.indexOf(c.greenSoft) < 0, sl.title + '：' + JSON.stringify(Array.from(new Set(f))));
  });
  const changes = appendix.find(sl => sl.title === T.changes);
  const pills = changes.elements.filter(el => el.role === 'pill');
  assert.ok(pills.some(p => p.fill === c.greenDeep), '前向きの区分は緑のまま');
  // 本編は緑のまま（色を戻している）
  const moves = slides.find(sl => sl.title === T.progress);
  assert.ok(fills(moves).indexOf(c.green) >= 0 && fills(moves).indexOf('#7F7F7F') < 0);
  assert.strictEqual(gas.get('MTG_OPTIONS').colors.green, '#70AD47');
  assert.strictEqual(gas.get('MTG_OPTIONS').baseColors, undefined);
  // appendixColors を null にすると、本編と同じ緑
  gas.get('MTG_OPTIONS').appendixColors = null;
  const green = plain(g.mtgSlides_(summarize(g, '新FMT'))).find(sl => sl.title === T.changes);
  assert.ok(fills(green).indexOf(c.green) >= 0);
});

test('サブインダストリーごとに作る：本番を1つずつ作ると、終わりの画面に、この会議日の本番の資料がまだできていないサブインダストリーを出し、そろったら統合版を案内する', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const meeting = g.mtgParseDate_('10/23', new Date());
  const parent = meetingFolders(gas, [yymmdd(meeting) + '_第5回会議']);
  Object.assign(o.combined, {enabled: true, meetingFolder: parent});
  gas.promptAnswers = ['2', '10/23'];
  g.createMeetingSummaries();
  const first = gas.alerts[gas.alerts.length - 1].message;
  assert.match(first, /^「新FMT2」：/);
  assert.doesNotMatch(first, /「新FMT」：/, '選んだシートだけ作る');
  assert.match(first, /この会議日の本番の資料が、まだできていないサブインダストリー（1件）：新FMT。すべてできたら、createMeetingSummaries\(\) の終わりにこの画面で知らせます/);
  assert.deepStrictEqual(JSON.parse(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT2')).map(m => m.date), [meeting]);
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT'), null, '選ばなかったシートは記録しない');
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /この会議日の本番の資料が、すべてのサブインダストリーでそろいました。統合版（全サブインダストリー）は、エディタで createCombinedMeetingSummary\(\) を実行して作ってください/);
  gas.promptAnswers = ['10/23'];
  g.createCombinedMeetingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「全サブインダストリー」（統合版）：/);
});
