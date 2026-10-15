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
  assert.match(text('予定日超過'), /^オファリング「【サンプル】販促DXの試験導入のご提案」：予定日（\d{1,2}\/\d{1,2}）から 3日経過、実施日未入力$/);
  assert.strictEqual(s.alerts[0].owner, 'サンプル 担当', '責任者の列には当社の担当者');
  assert.strictEqual(text('受注月超過'), '受注月（' + month(-1).y + '/' + month(-1).m + '）経過、ステータス「初回訪問」（延期 1回）');
  assert.strictEqual(text('未入力', 'D社'), '未入力：想定売上規模（百万）・期待値調整済想定売上規模（百万）・受注月・活動状況・アカウント責任者');
  assert.strictEqual(text('未入力', 'F社'), '未入力：活動状況', '想定の大きい順');
  assert.strictEqual(text('更新なし'), '最終更新から 35日経過');
  assert.strictEqual(text('金額の不整合'), '期待値（40）が想定売上（30）を上回る');
  assert.strictEqual(text('月の不整合'), '売上開始月（' + month(-3).y + '/' + month(-3).m + '）が受注月（' + month(-1).y + '/' + month(-1).m + '）より前');
  assert.strictEqual(text('選択肢外のステータス'), '「検討中」は選択肢外の値（進行中として集計）');
  assert.strictEqual(text('デリバリー担当なし'), '受注済み、「デリバリー担当」未入力');
  assert.ok(s.changes.every(c => !/^【サンプル】/.test(c.project)), 'サンプルの記録は変化に入れない');
  assert.strictEqual(s.kpi.owners, 6, 'サンプル 担当・佐藤・高橋・空・田中・伊藤');
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
  assert.ok(list.rows.some(r => r.cells[0].text === '伊藤' && r.cells[2].text.split('\n').indexOf('商談記録未登録') >= 0 &&
    r.cells[3].text.split('\n')[r.cells[2].text.split('\n').indexOf('商談記録未登録')] === empty), '内容は案件ごとの一覧に書く（項目は資料に出す名前。項目と内容の行をそろえる）');
});

test('受注月（四半期）別の期待値（受注月超過 → 今の四半期〜4Q → 来期以降 → 未入力）・ステータス別（プルダウンの順。進行中 → 保留 → 受注 → 失注）', () => {
  const {g} = setup();
  const s = plain(summarize(g, '新FMT'));
  assert.deepStrictEqual(s.stages.map(p => [p.status, p.kind, p.count, p.amount, p.weighted]), [
    ['提案中', 'active', 2, 110, 50],
    ['（未入力）', 'active', 1, 0, 0],
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
  assert.strictEqual(lead(slides[1]), '進行中案件 3件（直近14日間で ▲2件）、期待値合計 50百万円（直近14日間で ▲8）。');
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
          assert.ok(g.mtgLineHeight_(lines.length, cell.size) <= r.h + 0.01, '行の高さ：' + cell.text);
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
    ['受注', '0件', '—', '—'],
    ['失注・保留・後退・削除', '2件', '30', 'C社／店舗什器（失注）、F社／ブランド刷新（削除）'],
    ['新規・前進・復帰', '1件', '30', 'A社／アプリ刷新（初回訪問 ▶ 提案中）'],
    ['活動更新', '1件', '0', 'D社／物流改善（10/12 打ち合わせ）'],
    ['受注月・期待値などの変更', '0件', '—', '—']
  ], '行の順は P4 の区分の順と同じ。活動更新は独立した行にし、最新の活動状況を書く');
  assert.strictEqual(movesTable.rows.slice(1).reduce((sum, r) => sum + Number(r.cells[1].text.replace('件', '')), 0), 4,
    '動きの行の件数を合わせると、見出しの「変化のある案件」の件数になる（案件ごとに1回だけ数える）');
  assert.deepStrictEqual(textsOf(statesTable), [
    ['受注月超過', '1件', '30', 'A社／アプリ刷新（受注月 ' + month(-1).y + '/' + month(-1).m + '）'],
    ['停滞（同一ステータス45日以上）', '2件', '80', 'うち直近14日間の更新なし 1件：E社／販促DX（100日以上）']
  ], '停滞は定義を添え、名前は対象期間に更新の無い案件だけ（D社は活動を更新した）。商談記録の見出しが無いシートでは商談記録未登録の行を出さない');
  const stall = statesTable.rows[1].cells[0];
  assert.deepStrictEqual([stall.text, stall.size, stall.parts && stall.parts[0].bold], ['停滞\n（同一ステータス45日以上）', 7.5, true], '定義は小さな字で2行目に');
  assert.ok(movesTable.rows.slice(1).concat(statesTable.rows).every(r => r.cells[1].size === 11 && r.cells[1].color === c.greenDark), '件数は大きな緑の字');
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
  assert.strictEqual(lead(slides[7]), '進行中案件のうち、同一ステータスに 45日以上滞留している案件は 2件（想定売上合計 80百万円）。うち直近14日間の更新なし 1件。');
  assert.deepStrictEqual(roleOf(slides[7], 'lead').map(el => el.text.split('\n').length), [1, 1], '1つ目の文と、更新なしの件数は行を分ける');
  assert.strictEqual(lead(slides[9]), '入力内容の確認事項は 3項目（3案件）。');
  assert.ok(all.indexOf('Appendix ④') < 0 && all.indexOf('Appendix ⑤') < 0, '無い Appendix を案内しない');
  // Appendix ③：責任者別は表だけ（種類の説明は置かない。内容は一覧に書く）。見出しの無いシートでは資料なしを見ない
  assert.strictEqual(lead(slides[8]), '入力内容の確認事項は 3項目（3案件・責任者 3名）。');
  assert.deepStrictEqual(slides[8].elements.filter(el => !el.role && el.type === 'table').length, 1);
  assert.ok(all.indexOf('チェックの種類') < 0 && all.indexOf('資料なし') < 0);
});

test('スライド：Appendix は入りきらなければ続きのページに分け、見出しの行を繰り返す（全件を出す）。中身の無い Appendix は作らない', () => {
  const {g} = setup();
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
  assert.deepStrictEqual(slides.map(sl => sl.title), ['表紙', 'サマリー', 'パイプラインの変化', '前回から変化のある案件', '期待値上位案件の状況・課題'],
    'Appendix が1つも無ければ扉も作らない');
  const onlyAlerts = plain(g.mtgSlides_(Object.assign({}, empty, {alerts: s.alerts})));
  assert.deepStrictEqual(roleOf(onlyAlerts[5], 'toc').filter((_, i) => i % 3 === 0).map(el => el.text), ['Appendix ③'],
    '扉には、作った Appendix だけを出す');
  const breakdown = slides[1].elements.find(el => el.type === 'table').rows.slice(1).map(r => r.cells.map(cell => cell.text));
  assert.ok(breakdown.every(r => r[1] === '0件' && r[3] === '—'), '区分は同じまま、0件で出す');
});

test('スライド：入力内容のチェック（責任者別）は責任者 × 種類の件数の表だけ（種類の説明は置かない）。種類が多くてもページの幅に入る', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const s = summarize(g, '新FMT');
  const ownersOf = summary => plain(g.mtgSlides_(summary)).find(sl => sl.title === 'Appendix ③　入力内容の確認事項（責任者別）');
  const lead = leadOf;

  const few = ownersOf(s);
  assert.ok(!few.elements.some(el => el.text === 'チェックの種類'));
  assert.strictEqual(lead(few), '入力内容の確認事項は 3項目（3案件・責任者 3名）。');
  assert.deepStrictEqual(few.elements.find(el => el.type === 'table').rows.map(r => r.cells[0].text), ['責任者', '伊藤', '佐藤', '（責任者未設定）', '合計'],
    '責任者は名前の順（日本語の順。件数の順ではない。責任者未設定は最後）');
  const narrow = few.elements.find(el => el.type === 'table');
  assert.deepStrictEqual(narrow.columns, [120, 80, 80, 80, 40], '種類が少なければ、種類の列は 80pt まで（左に寄せる）');

  const types = gas.get('MTG_ALERT_TYPES');
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
  assert.strictEqual(g.mtgNames_([], note, 300, 2), '—');
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
  assert.deepStrictEqual(plain(g.mtgWrap_('決裁者（社長）に会えていない', 142, 9, 3)).map(t => t.length), [13, 1]);
  assert.deepStrictEqual(plain(g.mtgBalance_('決裁者（社長）に会えていない', 142, 9, 3)).map(t => t.length), [7, 7], '最後の行に1文字だけ残さない');
  assert.deepStrictEqual(plain(g.mtgBalance_('期待値・受注月', 74, 8.5, 2)), ['期待値・', '受注月']);
  assert.deepStrictEqual(plain(g.mtgBalance_('短い', 142, 9, 3)), ['短い'], '1行に入るものはそのまま');
  const dated = '10/13 価格の再提示';
  assert.strictEqual(g.mtgWrap_(dated, 98, 8.5, 2).length, 2);
  assert.deepStrictEqual(plain(g.mtgBalance_(dated, 98, 8.5, 2)), ['10/13', '価格の再提示'], '空白で分ければ1行ずつに入るなら、言葉の途中で切らずに空白で分ける');
  const para = '東西マート／物流センターの改善（失注）、富士ストア／店舗アプリ改修（提案中→課題ヒアリング）、海風マート／来店分析（削除）';
  assert.deepStrictEqual(plain(g.mtgBalance_(para, 438, 9, 3)), plain(g.mtgWrap_(para, 438, 9, 3)), '最後の行が十分に長ければ、そのまま（行をいっぱいに使う）');
  const long = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも';
  assert.deepStrictEqual(plain(g.mtgBalance_(long, 100, 9, 2)), plain(g.mtgWrap_(long, 100, 9, 2)), '「…」で切れるものはそのまま');
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
  assert.strictEqual(g.mtgWrap_('ABCDEFGHIJKLMNOPQRSTUVWXYZ', 150, 8, 3).length, 2, '英大文字は幅を広めに見積もる（BIZ UDPゴシックでは約 0.72）');
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

  // ひな形のページ：スライドは作らず、決まった文字を長い順に置き換える（そのページだけ）。部品の ID は idPrefix から
  const cover = {template: {title: 'スーパー・GMS', subtitle: 'RX Business MTG　事前資料', date: '2026年11月6日（金）', name: ''}, elements: [slide.elements[0]]};
  const onTemplate = plain(g.mtgSlideRequests_(cover, 'g1a2b3', 0, {scale: 4 / 3, existing: true, idPrefix: 'mtg_p1'}));
  assert.ok(!onTemplate.some(r => r.createSlide));
  assert.deepStrictEqual(onTemplate.filter(r => r.replaceAllText).map(r => [r.replaceAllText.containsText.text, r.replaceAllText.replaceText,
    r.replaceAllText.pageObjectIds.join(), r.replaceAllText.containsText.matchCase]), [
    ['2026年 4月 17日', '2026年11月6日（金）', 'g1a2b3', true],
    ['XXXXXX', 'スーパー・GMS', 'g1a2b3', true],
    ['XXXXX', 'RX Business MTG　事前資料', 'g1a2b3', true],
    ['XXXX', '', 'g1a2b3', true]
  ], '「XXXXXX」は「XXXXX」より先に置き換える');
  const added = onTemplate.find(r => r.createShape).createShape;
  assert.deepStrictEqual([added.objectId, added.elementProperties.pageObjectId], ['mtg_p1_0', 'g1a2b3']);
  // ひな形が無いとき：緑の背景に、ひな形と同じ位置で白い字（名前の欄は空なので描かない）
  const plainCover = plain(g.mtgSlideRequests_(cover, 'mtg_p1', 0));
  assert.ok(plainCover[0].createSlide && !plainCover.some(r => r.replaceAllText));
  assert.deepStrictEqual(plainCover.filter(r => r.insertText).map(r => r.insertText.text),
    ['スーパー・GMS', 'RX Business MTG　事前資料', '2026年11月6日（金）', '今期（2026/4〜）　受注']);
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
  assert.strictEqual(requests.filter(r => r.createSlide).length, 9, '表紙・本編4枚・扉・Appendix ①・③（責任者別・一覧）（停滞が無いので ② は無い）');
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[1]), /^新FMT2\nRX Business MTG　事前資料\n/);
  assert.match(inserted(gas.decks[1]), /\nサマリー\n/);
  assert.match(inserted(gas.decks[1]), /\n新FMT2（10\/23 会議資料）　金額の単位：百万円\n/);
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
  // 本番は、作る前に会議日（曜日）・シートごとの対象期間を出して確かめる（配布したあとは作り直さない）
  const confirm = gas.alerts.find(a => /本番の資料を作り、会議日を記録します。/.test(a.message)).message;
  assert.match(confirm, /会議日：\d{4}\/10\/23（[日月火水木金土]）/);
  assert.match(confirm, /・新FMT：直近14日間（前回の会議の記録なし）\n・新FMT2：直近14日間（前回の会議の記録なし）/);
  assert.match(confirm, /配布したあとは作り直さないでください。/);
  assert.doesNotMatch(confirm, /作成済み/);
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /「新FMT」：会議サマリ_新FMT_.*\.pdf（変化のある案件 4・停滞 0・確認事項 3）\n　https:\/\/drive\.google\.com\/file\/d\/file-1\/view/,
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
  assert.strictEqual(gas.driveFolders.length, 1);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「新FMT」：.*（変化のある案件 1・停滞 0・確認事項 2）/);
  assert.deepStrictEqual(gas.driveFiles.slice(2).map(f => f.name.replace(/_\d{4}-\d{2}-\d{2}/, '')), ['会議サマリ_新FMT.pdf'], '選んだシートだけ作る');
  // 同じ会議日で作り直しても、前回（10/23）からの変化（作った日時だけ上書き）。作り直す前に、作成済みであることを知らせる
  const before = {deck: gas.decks[gas.decks.length - 1].id, pdf: gas.driveFiles[gas.driveFiles.length - 1].id};
  gas.promptAnswers = ['1', '10/30'];
  g.createMeetingSummaries();
  assert.match(gas.alerts.filter(a => /本番の資料を作り/.test(a.message)).pop().message,
    /・新FMT：前回会議（10\/23）の資料作成時（\d{1,2}\/\d{1,2} \d{2}:\d{2}）以降\n　※ この会議日の資料は \d{1,2}\/\d{1,2} \d{2}:\d{2} に作成済み。作り直すと、次の会議の起点が今になります。/);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /（変化のある案件 1・/);
  assert.deepStrictEqual(gas.trashed.slice().sort(), [before.deck, before.pdf].sort(), '同じ名前の古い版（スライド・PDF）はゴミ箱に移す（古い版を配らないように）');
  assert.ok(gas.trashed.indexOf(gas.decks[gas.decks.length - 1].id) < 0 && gas.trashed.indexOf(gas.driveFiles[gas.driveFiles.length - 1].id) < 0, '新しい版は残す');
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT')).map(m => m.date), [meeting, next]);
  assert.deepStrictEqual(JSON.parse(props.getProperty('MTG_MEETINGS_新FMT2')).map(m => m.date), [meeting], '選ばなかったシートの会議日は記録しない');
});

test('setMeetingTemplate・ひな形：登録した Google スライドをコピーして資料を作る（表紙は決まった文字の置き換え、扉は表紙の複製、ほかはページの大きさに合わせて拡大）', () => {
  const {gas, g} = setup();
  const props = g.PropertiesService.getDocumentProperties();
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789';
  gas.slideTemplates[id] = {pageSize: {width: {magnitude: 12192000, unit: 'EMU'}, height: {magnitude: 6858000, unit: 'EMU'}},
    slides: [{objectId: 'g_cover'}, {objectId: 'g_extra'}]};
  const pptx = '1PptxFileIdAbCdEfGhIjKlMnOpQrStUv';
  gas.files[pptx] = {mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'};
  gas.promptAnswers = ['https://docs.google.com/presentation/d/' + pptx + '/edit'];
  assert.throws(() => g.setMeetingTemplate(), /ひな形が Google スライドではありません（\.pptx のままなど）。ドライブで開いて「ファイル」→「Google スライドとして保存」/);
  gas.promptAnswers = ['ひな形'];
  assert.throws(() => g.setMeetingTemplate(), /「ひな形」から Google スライドの ID を読めません/);
  assert.strictEqual(props.getProperty('MTG_TEMPLATE_ID'), null);
  gas.promptAnswers = ['https://docs.google.com/presentation/d/' + id + '/edit#slide=id.p'];
  g.setMeetingTemplate();
  assert.strictEqual(props.getProperty('MTG_TEMPLATE_ID'), id);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /ひな形を登録しました（2枚。1枚目を表紙・扉に使います）/);

  gas.layoutPlaceholders = ['SLIDE_NUMBER', 'FOOTER'];   // ひな形の白紙のレイアウトにはページ番号・フッターの枠がある
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.deepStrictEqual(gas.driveCopies.map(cp => [cp.from, cp.folder]), [[id, 'folder-1']], 'ひな形を保存先のフォルダにコピーする');
  const deck = gas.decks.find(d => d.template === id);
  const r = plain(deck.requests);
  assert.deepStrictEqual(r.filter(q => q.deleteObject && /_ph\d$/.test(q.deleteObject.objectId)).map(q => q.deleteObject.objectId),
    ['mtg_p2_ph0', 'mtg_p2_ph1', 'mtg_p3_ph0', 'mtg_p3_ph1', 'mtg_p4_ph0', 'mtg_p4_ph1', 'mtg_p5_ph0', 'mtg_p5_ph1',
      'mtg_p7_ph0', 'mtg_p7_ph1', 'mtg_p8_ph0', 'mtg_p8_ph1', 'mtg_p9_ph0', 'mtg_p9_ph1'],
    '足したスライドのレイアウトの枠（ページ番号など）は消す（ページ番号が2つ出ないように）');
  assert.ok(deck.slides.every(sl => !(sl.pageElements || []).length));
  assert.strictEqual(deck.batches[deck.batches.length - 1], 14, '枠の削除は最後にまとめて送る');
  assert.deepStrictEqual(r.slice(0, 2), [{deleteObject: {objectId: 'g_extra'}}, {duplicateObject: {objectId: 'g_cover', objectIds: {g_cover: 'mtg_divider'}}}],
    'ひな形の2枚目以降は消し、1枚目を複製して扉にする');
  const replaced = page => r.filter(q => q.replaceAllText && q.replaceAllText.pageObjectIds[0] === page)
    .map(q => [q.replaceAllText.containsText.text, q.replaceAllText.replaceText]);
  assert.deepStrictEqual(replaced('g_cover').slice(1), [['XXXXXX', '新FMT'], ['XXXXX', 'RX Business MTG　事前資料'], ['XXXX', '']]);
  assert.match(replaced('g_cover')[0][1], /^\d{4}年10月23日（[日月火水木金土]）$/);
  assert.deepStrictEqual(replaced('mtg_divider'), [['2026年 4月 17日', ''], ['XXXXXX', 'Appendix'], ['XXXXX', ''], ['XXXX', '']]);
  assert.ok(!r.some(q => q.deleteObject && q.deleteObject.objectId === 'g_cover'), '表紙（ひな形の1枚目）は消さない');
  assert.deepStrictEqual(r.filter(q => q.createSlide).map(q => q.createSlide.insertionIndex), [1, 2, 3, 4, 6, 7, 8],
    '表紙の次から本編を入れると、複製した扉は本編のあとになり、Appendix はそのあと');
  // 本編はページの大きさ（960 × 540pt）に合わせて 4/3 倍
  const title = r.find(q => q.createShape && q.createShape.elementProperties.pageObjectId === 'mtg_p2').createShape;
  assert.deepStrictEqual([title.elementProperties.size.width.magnitude, title.elementProperties.transform.translateX], [901.33, 29.33]);
  const titleStyle = r.find(q => q.updateTextStyle && q.updateTextStyle.objectId === title.objectId).updateTextStyle.style;
  assert.strictEqual(titleStyle.fontSize.magnitude, 28);
  assert.ok(r.some(q => q.createShape && q.createShape.elementProperties.pageObjectId === 'g_cover'), '表紙にデータの時点・期間を足す');
  assert.ok(!gas.moved.some(m => m.id === deck.id), 'コピーは保存先のフォルダにできるので動かさない');
  assert.strictEqual(gas.driveFiles[0].blob.from, deck.id, 'ひな形のコピーを PDF にする');

  // 空欄で登録を外すと、ひな形なしで作る
  gas.promptAnswers = [''];
  g.setMeetingTemplate();
  assert.strictEqual(props.getProperty('MTG_TEMPLATE_ID'), null);
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.strictEqual(gas.driveCopies.length, 1);
});

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
  assert.deepStrictEqual(gas.decks.map(d => d.title.replace(/_\d{4}-\d{2}-\d{2}/, '')), ['会議サマリ_新FMT_試作', '会議サマリ_新FMT_試作']);
  const inserted = deck => deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.match(inserted(gas.decks[1]), /入力内容の確認事項は 3項目（3案件・責任者 3名）。\n?試作時（\d{1,2}\/\d{1,2}）から 0項目解消、\n?0項目追加。/,
    '2回目の試作：まだ解消していない');

  // E社の活動状況を書いて「更新なし」を解消してから、本番を作る
  g.recordDiffEdit(gas.edit(sheet, 'J7', '10/20 決裁者と面談を設定'));
  gas.promptAnswers = ['1', '10/23'];
  g.createMeetingSummaries();
  assert.match(inserted(gas.decks[2]), /入力内容の確認事項は 2項目（2案件・責任者 2名）。\n?試作時（\d{1,2}\/\d{1,2}）から 1項目解消、\n?0項目追加。/);
});

test('removeMeetingRecord：会議日の記録を取り消す（はいのときだけ）。基準の記録・資料のファイルは残す', () => {
  const {gas, g} = setup();
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
    '今後、RX Business MTG の事前資料のご案内をこちらにお送りいたします。どうぞよろしくお願いいたします。');
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

test('notifyMeetingSummaries：格納先のフォルダの、いちばん新しい PDF のリンク（ファイル名で表示）を、確かめのあとすぐに送る。失敗は1回だけやり直し、だめなら止める', () => {
  const {gas, g} = setup({edits: false});
  assert.throws(() => g.notifyMeetingSummaries(), /先に setMeetingChat\(\) で webhook の URL と資料の格納先の URL を登録してください。/);
  const props = g.PropertiesService.getScriptProperties();
  props.setProperty('MTG_CHAT_WEBHOOK', WEBHOOK);
  props.setProperty('MTG_CHAT_URL', DEMO_URL);
  assert.throws(() => g.notifyMeetingSummaries(), /資料の格納先のフォルダ（demo-folder-0001）にファイルがありません。資料を置いてから送ってください。/);

  // 格納先のフォルダ：古い PDF・新しい PDF・もっと新しいスライド（PDF を選ぶ）・ほかのフォルダの PDF
  const file = (id, folder, name, type, daysAgo) => gas.driveFiles.push({id, folder, name, blob: {getContentType: () => type},
    updated: Date.now() - daysAgo * 86400000});
  file('pdf-old', 'demo-folder-0001', '会議サマリ_デモ_旧.pdf', 'application/pdf', 5);
  file('pdf-new', 'demo-folder-0001', '会議サマリ_デモ.pdf', 'application/pdf', 1);
  file('slide-1', 'demo-folder-0001', '会議サマリ_デモ（スライド）', 'application/vnd.google-apps.presentation', 0);
  file('pdf-other-0001', 'other-folder-0001', '別の資料.pdf', 'application/pdf', 0);
  const lines = link => [
    '*【RX Business MTG】事前資料のご案内*',
    '皆さま、お疲れさまです。',
    'RX Business MTG の事前資料を、下記に格納いたしました。',
    'お手数ですが、会議の前にご一読いただけますと幸いです。',
    '',
    '▼ 事前資料',
    link,
    '',
    'ご不明な点がございましたら、お気軽にお知らせください。',
    'どうぞよろしくお願いいたします。'
  ].join('\n');
  const url = 'https://drive.google.com/file/d/pdf-new/view';

  gas.confirmAnswer = 'NO';
  g.notifyMeetingSummaries();
  assert.strictEqual(gas.alerts[gas.alerts.length - 1].message,
    'Google チャット（スペース AAAAtest123）に、次のメッセージをすぐに送ります。よろしいですか？\n\n' + lines('会議サマリ_デモ.pdf\n' + url),
    '確かめの画面には、ファイル名と URL を出す');
  assert.strictEqual(gas.fetches.length, 0, '「いいえ」なら送らない');
  assert.strictEqual(gas.prompts.length, 0, 'ほかには何も聞かない');

  gas.confirmAnswer = 'YES';
  g.notifyMeetingSummaries();
  assert.deepStrictEqual(gas.fetches.map(f => [f.url, sentText(f)]), [[WEBHOOK, lines('<' + url + '|会議サマリ_デモ.pdf>')]],
    'フォルダの中のいちばん新しい PDF のリンク（ファイル名で表示）');
  assert.strictEqual(gas.alerts[gas.alerts.length - 1].message, 'Google チャットに送りました：会議サマリ_デモ.pdf\n' + url);

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
  assert.match(sentText(gas.fetches[3]), /\n<https:\/\/drive\.google\.com\/file\/d\/pdf-other-0001\/view\|別の資料\.pdf>\n/);
  props.setProperty('MTG_CHAT_URL', 'https://example.com/demo.pdf');
  g.notifyMeetingSummaries();
  assert.match(sentText(gas.fetches[4]), /\n<https:\/\/example\.com\/demo\.pdf\|https:\/\/example\.com\/demo\.pdf>\n/);

  // 文は MTG_OPTIONS.chat.message で変えられる（{meeting}・{file}・{url} を置き換える）
  props.setProperty('MTG_CHAT_URL', 'https://drive.google.com/file/d/pdf-other-0001/view');
  gas.get('MTG_OPTIONS').chat.message = ['{meeting}：{file}（{url}）'];
  gas.fetchStatuses = [503, 200];
  g.notifyMeetingSummaries();
  assert.strictEqual(gas.fetches.length, 7, '一時的な失敗（5xx）は1回だけやり直す');
  assert.strictEqual(sentText(gas.fetches[6]),
    'RX Business MTG：<https://drive.google.com/file/d/pdf-other-0001/view|別の資料.pdf>（https://drive.google.com/file/d/pdf-other-0001/view）');

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
  assert.match(inserted, /\n停滞\n（同一ステータス45日以上）\n0件\n—\n基準（\d{1,2}\/\d{1,2}）から 45日未満のため、まだ数えられない\n/,
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
