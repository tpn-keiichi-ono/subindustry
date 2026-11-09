'use strict';

const {test, assert} = require('./lib/harness');
const {createGas: createMockGas, plain} = require('./lib/gas-mock');
const {withRunDialog} = require('./lib/meeting-run');
const createGas = options => withRunDialog(createMockGas(options));

/*
 * 統合版（報告用。MeetingCompact.gs）：表紙 → 全サブインダストリーの状況・計画進捗 → 計画に対する積み上げ（全体）→ 変化のあった案件（サブインダストリーごと）
 * 追跡シート（2行目が見出し）：D 得意先 / E 案件名 / F 案件ステータス / G 想定売上 / H 期待値 / I 受注月 / J 活動状況 / K 課題 / L 責任者 / M 売上開始月
 * （テストでは計画に対する積み上げを売上開始月で数える設定。tests/lib/gas-mock.js）
 */
const WIDTH = 35;
const COL = {customer: 4, project: 5, status: 6, amount: 7, weighted: 8, orderMonth: 9, activity: 10, issue: 11, owner: 12, startMonth: 13, stamp: 33};
const HEADERS = {customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', orderMonth: '受注月', activity: '活動状況', issue: '活動における課題',
  owner: 'アカウント責任者', startMonth: '売上開始月', stamp: '最終更新日時'};
const STATUSES = ['創注-ニーズ探索(10%)', 'D-初回提案(30%)', 'C-提案中(50%)', 'B-クロージング中(70%)', 'A-内示(90%)', '受注', '失注'];

/** 今年度（4月始まり）の n か月目（0 から）の月 'yyyy/MM'。テストを実行する日に左右されないように、今年度の中の月にする */
function fiscalMonth(n) {
  const now = new Date(Date.now() + 9 * 3600000);
  const y = now.getUTCMonth() + 1 >= 4 ? now.getUTCFullYear() : now.getUTCFullYear() - 1;
  const d = new Date(Date.UTC(y, 3 + n, 1));
  return d.getUTCFullYear() + '/' + String(d.getUTCMonth() + 1).padStart(2, '0');
}

function row(values) {
  const line = new Array(WIDTH).fill('');
  Object.keys(values).forEach(key => { line[COL[key] - 1] = values[key]; });
  return line;
}

/** 得意先・案件名・ステータス・想定・期待値・売上開始月（今年度の何か月目）・活動状況・課題 */
const DEALS = {
  新FMT: [
    ['A社', '店舗DX', 'A-内示(90%)', 100, 90, 8, '内示を受けた', '契約条件の調整'],
    ['B社', '需要予測', 'B-クロージング中(70%)', 50, 35, 9, '見積を提出', ''],
    ['C社', '販促', 'C-提案中(50%)', 40, 20, 10, '提案した', '決裁者に会えていない'],
    ['D社', '物流', 'C-提案中(50%)', 20, 10, 11, '初回の提案', ''],
    ['E社', '会員', '失注', 30, 0, 9, '他社に決まった', '']
  ],
  新FMT2: [
    ['F社', 'EC', 'B-クロージング中(70%)', 20, 14, 8, '交渉中', ''],
    ['G社', '店舗', 'D-初回提案(30%)', 10, 3, 10, '初回訪問', '']
  ]
};

function setup() {
  const gas = createGas();
  const g = gas.global;
  const o = gas.get('MTG_OPTIONS');
  o.plan.plans = {新FMT: 200, 新FMT2: 100};
  Object.keys(DEALS).forEach(name => {
    const lines = [row({customer: name}), row(HEADERS)].concat(DEALS[name].map(d => row({customer: d[0], project: d[1], status: d[2], amount: d[3],
      weighted: d[4], orderMonth: fiscalMonth(d[5]), startMonth: fiscalMonth(d[5]), activity: d[6], issue: d[7], owner: '佐藤'})));
    const sheet = gas.addSheet(name, lines, {rows: 30, columns: WIDTH});
    sheet.getRange('F3:F30').setDataValidation(g.SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  });
  g.setupDiffTracking();
  const sheet = gas.ss.getSheetByName('新FMT');
  g.recordDiffEdit(gas.edit(sheet, 'F5', 'B-クロージング中(70%)'));   // C社：C → B（前進）
  g.recordDiffEdit(gas.edit(sheet, 'H5', 28));                        // C社：期待値 20 → 28
  g.recordDiffEdit(gas.edit(sheet, 'H6', 4));                         // D社：期待値 10 → 4（減額）
  g.recordDiffEdit(gas.edit(sheet, 'J6', '11/2 提案を見直し'));          // D社：活動状況
  gas.writes.length = 0;
  return {gas, g, o, sheet};
}

/** 1つのシートの中身（前回の会議 = 14日前の資料作成時。資料の見出しは「前回比（M/D）」） */
function summarize(g, name, previous) {
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const since = new Date(Date.now() - 14 * 86400000);
  const windowStart = new Date(Date.now() - 180 * 86400000);
  return g.mtgSummarize_(ss, name, {
    since, previous: previous === undefined ? {date: '2026/10/30', cutoff: since.getTime()} : previous, legacy: null, baseline: null, draft: null,
    final: true, now, meeting: '2026/11/13', windowStart, copy: null, copyError: '', trackedSince: g.mtgLogStart_(ss),
    log: g.mtgReadLog_(ss, windowStart), snapshots: g.mtgReadSnapshots_(ss, windowStart), credentials: g.hsCredentialIndex_(ss)
  });
}

const texts = slide => slide.elements.filter(e => e.text != null).map(e => String(e.text));
const tableOf = slide => slide.elements.find(e => e.type === 'table');
const leadOf = slide => slide.elements.filter(e => e.role === 'lead').map(e => e.text).join('');
const cellTexts = table => table.rows.map(r => r.cells.map(c => String(c.text)));

test('統合版（報告用）のスライド：表紙（統合版と同じひな形）→ 状況・計画進捗 → 計画に対する積み上げ → 変化のあった案件（サブインダストリーごと）。目次・各資料のページは無い', () => {
  const {gas, g} = setup();
  const T = gas.get('MTG_TITLES');
  const slides = plain(g.mtgCompactSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/11/13', true));
  assert.deepStrictEqual(slides.map(s => s.title), [T.cover, '全サブインダストリーの状況・計画進捗', T.plan + '（全サブインダストリー）',
    '変化のあった案件：新FMT', '変化のあった案件：新FMT2']);
  // 表紙は統合版と同じ（ひな形の統合版の表紙・同じ文字）
  const combined = plain(g.mtgCombinedSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/11/13', true));
  assert.deepStrictEqual(slides[0].template, combined[0].template);
  assert.ok(slides[0].cover && slides[0].combined);
  assert.ok(!slides.some(s => s.title === T.contents || s.title === T.overview), '目次・統合版の状況のページは無い');
  // ページ番号は表紙から通し（表紙には出さない）
  assert.deepStrictEqual(slides.slice(1).map(s => s.elements.find(e => e.role === 'page').text), ['2', '3', '4', '5']);
  assert.match(texts(slides[1]).join('\n'), /全サブインダストリー（11\/13 会議資料）　金額の単位：百万円/);
});

test('変化のあった案件：サブインダストリーごとのページ。No. は資料の通し番号（緑の網掛け）、サブインダストリーの列・該当事項の列は無い。変化が無ければ「該当なし」', () => {
  const {gas, g} = setup();
  const c = gas.get('MTG_OPTIONS').colors;
  g.recordDiffEdit(gas.edit(gas.ss.getSheetByName('新FMT2'), 'J3', '11/3 条件を提示'));   // 新FMT2 の F社：活動状況
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const t1 = tableOf(slides[3]);
  const head = cellTexts(t1)[0];
  assert.deepStrictEqual(head, ['No.', '得意先／案件名', 'ステータス\n前回▶今回', '期待値\n前回▶今回', '受注月\n前回▶今回', '活動状況（最新）', '活動における課題']);
  const rows1 = cellTexts(t1).slice(1);
  assert.deepStrictEqual(rows1.map(r => r[0]), ['1', '2'], '新FMT は 1 から');
  assert.deepStrictEqual(rows1.map(r => r[1].split('\n')[0]), ['C社', 'D社'], '今回のステータスの順');
  assert.strictEqual(rows1[0][2], 'C-提案中(50%)\n▶ B-クロージング中(70%)', '1行に入らなければ「▶」の前で改行');
  assert.strictEqual(rows1[1][3], '10 ▶ 4');
  assert.strictEqual(rows1[0][6], '決裁者に会えていない', '活動における課題まで');
  assert.ok(t1.rows.slice(1).every(r => r.cells[0].fill === c.greenSoft), 'No. は緑の網掛け');
  const rows2 = cellTexts(tableOf(slides[4])).slice(1);
  assert.deepStrictEqual(rows2.map(r => r[0]), ['3'], '次のサブインダストリーは続きの番号');
  assert.match(leadOf(slides[3]), /^前回から変化のある案件は 2件。/);

  // 変化が無いサブインダストリー：表の見出しの下に「該当なし」
  const {g: g2} = setup();
  const empty = plain(g2.mtgCompactSlides_([summarize(g2, '新FMT'), summarize(g2, '新FMT2')], '2026/11/13', true))[4];
  assert.strictEqual(empty.title, '変化のあった案件：新FMT2');
  assert.match(texts(empty).join('\n'), /該当なし/);
  assert.match(leadOf(empty), /変化のある案件なし/);
});

test('変化のあった案件：行が入らなければ「（続き）」のページに分ける。表は出典の行の上の区切り線より tableMargin 上までに収める', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const o = gas.get('MTG_OPTIONS');
  const sheet = gas.ss.getSheetByName('新FMT');
  const long = '先方の部長・担当者と打ち合わせ、導入範囲とスケジュールを確認した。次回は見積の前提をすり合わせる。'.repeat(3);
  [3, 4, 5, 6].forEach(r => g.recordDiffEdit(gas.edit(sheet, 'J' + r, r + '日 ' + long)));
  const s = summarize(g, '新FMT');
  const part = plain(g.mtgCompactChangeSlides_(s, 10));
  assert.ok(part.slides.length >= 2, '続きのページ');
  assert.strictEqual(part.count, 4);
  assert.deepStrictEqual(part.slides.map(sl => sl.title), ['変化のあった案件：新FMT'].concat(part.slides.slice(1).map(() => '変化のあった案件：新FMT（続き）')));
  part.slides.forEach(sl => {
    const t = tableOf(sl);
    const bottom = t.y + t.rows.reduce((h, r) => h + r.h, 0);
    assert.ok(bottom <= S.bottom - o.compact.tableMargin + 0.01, '表の下端 ' + bottom + ' が出典の行にかからない');
    assert.strictEqual(cellTexts(t)[0][0], 'No.', '各ページに見出しの行');
  });
  const numbers = [].concat.apply([], part.slides.map(sl => cellTexts(tableOf(sl)).slice(1).map(r => r[0])));
  assert.deepStrictEqual(numbers, ['11', '12', '13', '14'], '前のサブインダストリーの続きの番号から');
});

test('状況・計画進捗：確定（A＋B）の横棒（A・B だけ・データラベル）と計画線（透明度のあるグレー・見出し「計画」）、進捗率（A＋B）の高い順、前回比は日付・増は青・減は赤', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const c = o.colors;
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const slide = plain(g.mtgCompactOverviewSlide_([s1, s2]));
  assert.strictEqual(slide.title, '全サブインダストリーの状況・計画進捗');
  const all = texts(slide);
  // 列の見出し（まとまりごと）
  ['今年度の計画進捗', '案件の状況（失注除く）', '計画', '進捗率\n（A＋B）', '確定（A＋B）\n/件', '見込（C）\n/件', '期待値\n合計', '前回比\n（10/30）', '案件数',
    '変化の\nある案件'].forEach(label => assert.ok(all.indexOf(label) >= 0, label));
  assert.ok(all.indexOf('停滞') < 0 && all.indexOf('確認\n事項') < 0, '停滞・確認事項の列は無い');
  // 進捗率（A＋B）＝ 確定 ÷ 計画。新FMT：A 90 + B（35 + C社 28）= 153 / 200 = 77%、新FMT2：B 14 / 100 = 14%。高い順
  const names = slide.elements.filter(e => e.role === 'compact-name').map(e => e.text);
  assert.deepStrictEqual(names, ['新FMT', '新FMT2', '合計']);
  const values = slide.elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual(values.slice(0, 8), ['200', '77%', '153/3件', '4/1件', '157', '+2', '4件', '2件'], '新FMT の行');
  assert.deepStrictEqual(values.slice(8, 16), ['100', '14%', '14/1件', '0/0件', '17', '±0', '2件', '0件'], '新FMT2 の行');
  assert.deepStrictEqual(values.slice(16), ['300', '56%', '167/4件', '4/1件', '174', '+2', '6件', '2件'], '合計の行');
  // 前回比の色：増は青、減は赤、±0 は黒
  const delta = text => slide.elements.find(e => e.role === 'compact-value' && e.text === text);
  assert.strictEqual(delta('+2').color, c.plus);
  assert.strictEqual(delta('±0').color, c.ink);
  const minus = plain(g.mtgCompactOverviewSlide_([Object.assign({}, s1, {bridge: Object.assign({}, s1.bridge, {start: s1.bridge.end + 5})})]));
  assert.strictEqual(minus.elements.find(e => e.role === 'compact-value' && /^▲/.test(e.text)).color, c.minus);
  // 横棒は A・B だけ（C は描かない）。データラベルは棒の中（白）か、入らなければ棒の右に「B 14」のように確度を付けて
  const bars = slide.elements.filter(e => e.role === 'compact-bar');
  const colorOf = key => o.plan.ranks.find(r => r.key === key).color;
  assert.deepStrictEqual(bars.map(b => b.fill), [colorOf('A'), colorOf('B'), colorOf('B')]);
  const labels = slide.elements.filter(e => e.role === 'compact-bar-label').map(e => e.text);
  assert.ok(labels.indexOf('90') >= 0 || labels.indexOf('A 90') >= 0);
  assert.ok(labels.some(t => /^(B )?14$/.test(t)));
  // 凡例は A・B だけ
  assert.deepStrictEqual(slide.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text), ['A', 'B']);
  // 計画線：透明度のあるグレーの点線。上に見出し「計画」（計画の位置ごと）
  const lines = slide.elements.filter(e => e.role === 'compact-plan');
  assert.ok(lines.length >= 2 && lines.every(l => l.color === o.compact.planLine.color && l.alpha === o.compact.planLine.alpha && l.dash));
  assert.deepStrictEqual(slide.elements.filter(e => e.role === 'compact-plan-label').map(e => e.text), ['計画', '計画']);
  // いちばん大きい計画（200）は、棒の幅の 1/planScale の位置
  const barX = Math.min.apply(null, bars.map(b => b.x));
  const head = lines.filter(l => l.y < slide.elements.find(e => e.role === 'compact-name').y);
  const far = Math.max.apply(null, head.map(l => l.x));
  const right = Math.max.apply(null, slide.elements.filter(e => e.role === 'compact-head').map(e => e.x));
  assert.ok(far > barX && far < right, '計画線は棒の場所の中');
  assert.match(slide.lead, /^進捗率（A＋B）は 56%（計画に届いているのは 2のうち 0）。案件 6件・期待値合計 174百万円（失注除く）。\n進捗率（A＋B）の高い順。/);
});

test('状況・計画進捗：前回の会議でなく基準・直近14日間と比べるときは、前回比の見出しに日付を付けない。期待値・月の列が無いシートは「データなし」', () => {
  const {g} = setup();
  const first = plain(g.mtgCompactOverviewSlide_([summarize(g, '新FMT', null)]));
  assert.ok(texts(first).indexOf('直近14日間の増減') >= 0);
  const s = summarize(g, '新FMT');
  const gap = Object.assign({}, s, {plan: Object.assign({}, s.plan, {gap: 'シートに「期待値」の列が無い'})});
  const slide = plain(g.mtgCompactOverviewSlide_([gap, summarize(g, '新FMT2')]));
  assert.ok(texts(slide).some(t => /^データなし（シートに「期待値」の列が無い）/.test(t)));
  assert.deepStrictEqual(slide.elements.filter(e => e.role === 'compact-name').map(e => e.text), ['新FMT2', '新FMT', '合計'], '作れないシートは下');
});

test('計画線の透明度は Slides API の線の色に入る（alpha）。透明度の無い線は今までどおり', () => {
  const {g} = setup();
  const requests = plain(g.mtgSlideRequests_({title: 't', elements: [
    {type: 'line', x: 10, y: 10, w: 0, h: 20, color: '#808080', alpha: 0.6, weight: 1.25, dash: true},
    {type: 'line', x: 20, y: 10, w: 0, h: 20, color: '#000000'}
  ]}, 'p1', 1));
  const props = requests.filter(r => r.updateLineProperties).map(r => r.updateLineProperties.lineProperties.lineFill.solidFill);
  assert.deepStrictEqual(props.map(p => p.alpha), [0.6, undefined]);
});

test('createBusinessMtgReportingSummary：会議日を選んで作り、ファイル名に「_報告用」。会議のフォルダに格納し、インデックスには書かない。統合版とは別の資料', () => {
  const {gas, g, o} = setup();
  const meeting = g.mtgParseDate_('11/13', new Date());
  const parent = 'SubIndustryFolder0001';
  gas.driveFolders.push({id: parent, name: 'スーパー・GMS', parent: 'root'});
  gas.driveFolders.push({id: 'meeting-folder-0', name: meeting.slice(2, 4) + meeting.slice(5, 7) + meeting.slice(8, 10) + '_第6回会議', parent});
  o.meetingFolder = 'https://drive.google.com/drive/folders/' + parent;
  gas.confirmAnswer = 'YES';
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingSummary();
  const check = gas.alerts[gas.alerts.length - 2].message;
  assert.match(check, /^「全サブインダストリー」の統合版（報告用）（2サブインダストリーの状況・計画進捗と、変化のあった案件）を、今のシートの値で作ります。/);
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /^「統合版（報告用）（全サブインダストリー）」：\d{8}_全サブインダストリー_RX Business Meeting用資料_報告用\.pdf（5ページ）/);
  assert.match(done, /会議のフォルダ「\d{6}_第6回会議」に格納しました/);
  assert.ok(gas.driveFiles.some(f => f.folder === 'meeting-folder-0' && /_報告用\.pdf$/.test(f.name)));
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), ['mtg_p1', 'mtg_p2', 'mtg_p3', 'mtg_p4', 'mtg_p5'], '前のページは最後に足して、番号の順に並べる');
  const index = gas.ss.getSheetByName(o.index.sheet);
  assert.ok(!index || !/_報告用/.test(JSON.stringify(index.getDataRange().getDisplayValues())), 'インデックスには書かない');
  assert.strictEqual(g.PropertiesService.getScriptProperties().getProperty('MTG_JOB'), null);
});

test('createBusinessMtgReportingSummary：時間がかかれば区切り、続き（continueBusinessMtgSummaries）で残りと仕上げを作る。できあがりは区切らずに作ったときと同じ', () => {
  const once = (() => {
    const {gas, g} = setup();
    gas.confirmAnswer = 'YES';
    gas.promptAnswers = ['11/13'];
    g.createBusinessMtgReportingSummary();
    const deck = gas.decks[gas.decks.length - 1];
    return {slides: deck.slides.map(sl => sl.objectId), texts: deck.requests.filter(r => r.insertText).map(r => r.insertText.text)
      .filter(t => !/\d{1,2}:\d{2}/.test(t)).sort()};
  })();
  const {gas, g, o} = setup();
  const sprops = g.PropertiesService.getScriptProperties();
  const clock = {t: 0};
  g.mtgNow_ = () => clock.t;
  const add = g.mtgDeckAdd_;
  g.mtgDeckAdd_ = (...args) => { clock.t += 100000; return add(...args); };
  o.run.limitMs = 150000;
  gas.confirmAnswer = 'YES';
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^時間がかかるため、ここで区切りました（Apps Script の1回の実行は6分まで）。残りは約1分後に自動で続けます：変化のあった案件のページ 1件（新FMT2）・/);
  const saved = JSON.parse(sprops.getProperty('MTG_JOB'));
  assert.deepStrictEqual([saved.kind, saved.page, saved.no], ['compact', 4, 2]);
  // 途中：ほかの資料・統合版はすぐには作らない。結果の画面に途中と出す
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^統合版（報告用）（全サブインダストリー）を作っている途中です（会議日 \d{4}\/11\/13。1\/2サブインダストリーのページ）。/);
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^統合版（報告用）（全サブインダストリー）を作っている途中です（自動で続けています）。/);
  o.run.limitMs = 400000;
  clock.t = 0;
  g.continueBusinessMtgSummaries();
  assert.strictEqual(sprops.getProperty('MTG_JOB'), null);
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), once.slides);
  assert.deepStrictEqual(deck.requests.filter(r => r.insertText).map(r => r.insertText.text).filter(t => !/\d{1,2}:\d{2}/.test(t)).sort(), once.texts,
    '区切らずに作ったときと同じ文字');
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「統合版（報告用）（全サブインダストリー）」：.*_報告用\.pdf（5ページ）/);
});

test('統合版（報告用）：作っている途中で止まったら、作りかけの資料をゴミ箱に移し、createBusinessMtgReportingSummary() での作り直しを案内する', () => {
  const {gas, g, o} = setup();
  const clock = {t: 0};
  g.mtgNow_ = () => clock.t;
  const add = g.mtgDeckAdd_;
  g.mtgDeckAdd_ = (...args) => { clock.t += 100000; return add(...args); };
  o.run.limitMs = 150000;
  gas.confirmAnswer = 'YES';
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingSummary();
  const id = gas.decks[gas.decks.length - 1].id;
  g.mtgDeckAdd_ = () => { throw new Error('Slides API の失敗'); };
  clock.t = 0;
  o.run.limitMs = 400000;
  g.continueBusinessMtgSummaries();
  assert.ok(gas.trashed.indexOf(id) >= 0);
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message,
    /統合版（報告用）（全サブインダストリー）を作る途中で止まりました：Slides API の失敗\ncreateBusinessMtgReportingSummary\(\) で、作り直してください。/);
});
