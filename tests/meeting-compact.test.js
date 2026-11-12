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
const COL = {customer: 4, project: 5, status: 6, amount: 7, weighted: 8, orderMonth: 9, activity: 10, issue: 11, owner: 12, startMonth: 13,
  salesH1: 16, salesH2: 17, stamp: 33};
const HEADERS = {customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', orderMonth: '受注月', activity: '活動状況', issue: '活動における課題',
  owner: 'アカウント責任者', startMonth: '売上開始月', stamp: '最終更新日時'};

/** 今年度の年度の下2桁（4月始まり。'26'） */
function fiscalYear() {
  const now = new Date(Date.now() + 9 * 3600000);
  return String(now.getUTCMonth() + 1 >= 4 ? now.getUTCFullYear() : now.getUTCFullYear() - 1).slice(-2);
}
// 今年度の売上の列（本物の P・Q列「FY26前半期売上（百万）」「FY26後半期売上（百万）」と同じ見出し）
HEADERS.salesH1 = 'FY' + fiscalYear() + '前半期売上（百万）';
HEADERS.salesH2 = 'FY' + fiscalYear() + '後半期売上（百万）';
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

/** 得意先・案件名・ステータス・想定・期待値・売上開始月（今年度の何か月目）・活動状況・課題・今年度の売上（前半期・後半期） */
const DEALS = {
  新FMT: [
    ['A社', '店舗DX', 'A-内示(90%)', 100, 90, 8, '内示を受けた', '契約条件の調整', 30, 40],
    ['B社', '需要予測', 'B-クロージング中(70%)', 50, 35, 9, '見積を提出', '', 20, 10],
    ['C社', '販促', 'C-提案中(50%)', 40, 20, 10, '提案した', '決裁者に会えていない', 10, ''],
    ['D社', '物流', 'C-提案中(50%)', 20, 10, 11, '初回の提案', '', '', ''],
    ['E社', '会員', '失注', 30, 0, 9, '他社に決まった', '', 50, '']
  ],
  新FMT2: [
    ['F社', 'EC', 'B-クロージング中(70%)', 20, 14, 8, '交渉中', '', 5, 5],
    ['G社', '店舗', 'D-初回提案(30%)', 10, 3, 10, '初回訪問', '', '', '']
  ]
};

function setup() {
  const gas = createGas();
  const g = gas.global;
  const o = gas.get('MTG_OPTIONS');
  o.plan.plans = {新FMT: 200, 新FMT2: 100};
  Object.keys(DEALS).forEach(name => {
    const lines = [row({customer: name}), row(HEADERS)].concat(DEALS[name].map(d => row({customer: d[0], project: d[1], status: d[2], amount: d[3],
      weighted: d[4], orderMonth: fiscalMonth(d[5]), startMonth: fiscalMonth(d[5]), activity: d[6], issue: d[7], owner: '佐藤', salesH1: d[8], salesH2: d[9]})));
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
/** 想定売上だけで集計する設定（compact.salesPage: false。2026/11/10 までの既定。想定売上のページの決まりを確かめるテスト） */
const amountOnly = gas => { gas.get('MTG_OPTIONS').compact.salesPage = false; };
const tableOf = slide => slide.elements.find(e => e.type === 'table');
const leadOf = slide => slide.elements.filter(e => e.role === 'lead').map(e => e.text).join('');
const cellTexts = table => table.rows.map(r => r.cells.map(c => String(c.text)));

test('統合版（報告用）のスライド：表紙（統合版と同じひな形）→ 状況・計画進捗 → 計画に対する積み上げ → 変化のあった案件（サブインダストリーごと）。目次・各資料のページは無い', () => {
  const {gas, g} = setup();
  const T = gas.get('MTG_TITLES');
  const slides = plain(g.mtgCompactSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/11/13', true));
  const fy = 'FY' + g.mtgFiscalYearKey_(new Date());
  // 集計の元は今年度の売上（2026/11/10 から）。想定売上のページは入れない（compact.amountPages）
  const span = fy + '〜' + (Number(g.mtgFiscalYearKey_(new Date())) + 2);
  // 2026/11/12 から、売上のページは FY26〜28 の版（状況・計画進捗1枚・計画に対する積み上げ2枚）。前の FY26売上の2枚は入れない（compact.outlookPage）
  assert.deepStrictEqual(slides.map(s => s.title), [T.cover, '全サブインダストリーの状況・計画進捗（' + span + '売上）',
    T.plan + '（' + span + '売上）', T.plan + '（' + span + '売上）：確定（A＋B）',
    '変化のあった案件：新FMT', '変化のあった案件：新FMT2']);
  // 想定売上のページも入れる（amountPages）：状況・計画進捗、計画に対する積み上げのそれぞれ、今年度の売上のページのあと
  gas.get('MTG_OPTIONS').compact.amountPages = true;
  assert.deepStrictEqual(plain(g.mtgCompactSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/11/13', true)).map(s => s.title),
    [T.cover, '全サブインダストリーの状況・計画進捗（' + span + '売上）', '全サブインダストリーの状況・計画進捗', '全サブインダストリーの状況・計画進捗（見込（C）を含む）',
      T.plan + '（' + span + '売上）', T.plan + '（' + span + '売上）：確定（A＋B）',
      T.plan + '：想定売上', T.plan + '：想定売上・確定（A＋B）', '変化のあった案件：新FMT', '変化のあった案件：新FMT2']);
  assert.strictEqual(g.mtgCompactFrontCount_(), 8);
  assert.strictEqual(g.mtgCompactFrontCount_('changes'), 1, '変化のあった案件だけは表紙の1枚');
  gas.get('MTG_OPTIONS').compact.amountPages = false;
  // 表紙は統合版と同じひな形。見出しの下は「事前資料」ではなく「報告資料」
  const combined = plain(g.mtgCombinedSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/11/13', true));
  assert.deepStrictEqual(slides[0].template, Object.assign({}, combined[0].template, {subtitle: 'RX Business Meeting　報告資料'}));
  assert.strictEqual(combined[0].template.subtitle, 'RX Business Meeting　事前資料', '統合版はそのまま');
  assert.ok(slides[0].cover && slides[0].combined);
  assert.ok(!slides.some(s => s.title === T.contents || s.title === T.overview), '目次・統合版の状況のページは無い');
  // ページ番号は表紙から通し（表紙には出さない）
  assert.deepStrictEqual(slides.slice(1).map(s => s.elements.find(e => e.role === 'page').text), ['2', '3', '4', '5', '6']);
  assert.match(texts(slides[1]).join('\n'), /全サブインダストリー（11\/13 会議資料）　金額の単位：百万円/);
});

test('変化のあった案件：サブインダストリーごとのページ。No. は資料の通し番号（緑の網掛け）、サブインダストリーの列・該当事項の列は無い。変化が無ければ「該当なし」', () => {
  const {gas, g} = setup();
  amountOnly(gas);
  const c = gas.get('MTG_OPTIONS').colors;
  g.recordDiffEdit(gas.edit(gas.ss.getSheetByName('新FMT2'), 'J3', '11/3 条件を提示'));   // 新FMT2 の F社：活動状況
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const t1 = tableOf(slides[5]);
  const head = cellTexts(t1)[0];
  assert.deepStrictEqual(head, ['No', '得意先／案件名', 'ステータス\n前回▶今回', '想定売上\n前回▶今回', '受注月\n前回▶今回', '活動状況（最新）', '活動における課題']);
  const rows1 = cellTexts(t1).slice(1);
  assert.deepStrictEqual(rows1.map(r => r[0]), ['1', '2'], '新FMT は 1 から');
  assert.deepStrictEqual(rows1.map(r => r[1].split('\n')[0]), ['C社', 'D社'], '確度 B → C の順');
  assert.strictEqual(rows1[0][2], 'C-提案中(50%)\n▶ B-クロージング中(70%)', '1行に入らなければ「▶」の前で改行');
  assert.strictEqual(rows1[1][3], '20', '想定売上（期待値は 10 ▶ 4 と変わったが、想定売上は変わっていない）');
  assert.strictEqual(rows1[0][6], '決裁者に会えていない', '活動における課題まで');
  assert.ok(t1.rows.slice(1).every(r => r.cells[0].fill === c.greenSoft), 'No. は緑の網掛け');
  assert.ok(t1.rows.slice(1).every(r => r.cells[0].size <= 8), 'No. は小さめの字');
  const rows2 = cellTexts(tableOf(slides[6])).slice(1);
  assert.deepStrictEqual(rows2.map(r => r[0]), ['3'], '次のサブインダストリーは続きの番号');
  assert.match(leadOf(slides[5]), /^前回から変化のある案件は 2件。/);

  // 変化が無いサブインダストリー：表の見出しの下に「該当なし」
  const {g: g2} = setup();
  const empty = plain(g2.mtgCompactSlides_([summarize(g2, '新FMT'), summarize(g2, '新FMT2')], '2026/11/13', true)).find(sl => sl.title === '変化のあった案件：新FMT2');
  assert.strictEqual(empty.title, '変化のあった案件：新FMT2');
  assert.match(texts(empty).join('\n'), /該当なし/);
  assert.match(leadOf(empty), /変化のある案件なし/);
});

test('変化のあった案件：行が入らなければ次のページに分け、タイトルに「（1/2ページ）」。表は出典の行の上の区切り線より tableMargin 上までに収める', () => {
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
  assert.deepStrictEqual(part.slides.map(sl => sl.title), part.slides.map((_, n) => '変化のあった案件：新FMT（' + (n + 1) + '/' + part.slides.length + 'ページ）'));
  part.slides.forEach(sl => {
    const t = tableOf(sl);
    const bottom = t.y + t.rows.reduce((h, r) => h + r.h, 0);
    assert.ok(bottom <= S.bottom - o.compact.tableMargin + 0.01, '表の下端 ' + bottom + ' が出典の行にかからない');
    assert.strictEqual(cellTexts(t)[0][0], 'No', '各ページに見出しの行');
  });
  const numbers = [].concat.apply([], part.slides.map(sl => cellTexts(tableOf(sl)).slice(1).map(r => r[0])));
  assert.deepStrictEqual(numbers, ['11', '12', '13', '14'], '前のサブインダストリーの続きの番号から');
});

test('状況・計画進捗：確定（A＋B）の横棒（A・B だけ・データラベル）と計画線（透明度のあるグレー・見出し「計画」）、シートの順（並べ替えない）、前回比は日付・増は青・減は赤', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const c = o.colors;
  amountOnly(gas);
  g.recordDiffEdit(gas.edit(gas.ss.getSheetByName('新FMT'), 'G5', 50));   // C社：想定売上 40 → 50
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const slide = plain(g.mtgCompactOverviewSlide_([s1, s2]));
  assert.strictEqual(slide.title, '全サブインダストリーの状況・計画進捗');
  const all = texts(slide);
  // 列の見出し
  assert.ok(all.indexOf('今年度の計画進捗') < 0 && all.indexOf('案件の状況（失注除く）') < 0, 'まとまりの見出しは無い');
  ['計画', '進捗率（A＋B）', '確定（A＋B）/件', '見込（C）/件', '想定売上合計', '前回比（10/30）', '案件数', '変化のある案件']
    .forEach(label => assert.ok(all.indexOf(label) >= 0, label));
  // 見出しは折り返さずに1行（列の幅の 9割に入る大きさ）
  slide.elements.filter(e => e.role === 'compact-head').forEach(e => {
    assert.ok(e.text.indexOf('\n') < 0 && g.mtgTextWidth_(e.text, e.size) <= (e.w - 14.4) * 0.9 + 0.01, e.text + ' は1行');
  });
  // 対象の期間：計画・進捗率・確定・見込は「（今年度）」、右の4列は「（全期間）」。どの列も同じ高さの行に置き、見出しは下にそろえる
  const periods = slide.elements.filter(e => e.role === 'compact-head-period');
  assert.deepStrictEqual(periods.map(e => e.text), ['（今年度）', '（今年度）', '（今年度）', '（今年度）', '（全期間）', '（全期間）', '（全期間）', '（全期間）']);
  assert.strictEqual(new Set(periods.map(e => e.y)).size, 1, '期間の行は同じ高さ');
  const heads = slide.elements.filter(e => e.role === 'compact-head');
  assert.ok(heads.every(e => e.valign === 'BOTTOM' && e.y === heads[0].y && e.h === heads[0].h), '見出しは同じ箱で下にそろえる');
  assert.ok(all.indexOf('停滞') < 0 && all.indexOf('確認\n事項') < 0, '停滞・確認事項の列は無い');
  // 集計はどれも想定売上（期待値ではない）。進捗率（A＋B）＝ 確定 ÷ 計画。新FMT：A 100 + B（B社 50 + C社 50）= 200 / 200 = 100%、
  // 新FMT2：B 20 / 100 = 20%。高い順。想定売上合計（失注除く）：新FMT 100 + 50 + 50 + 20 = 220（前回比 +10）、新FMT2 30
  const names = slide.elements.filter(e => e.role === 'compact-name').map(e => e.text);
  assert.deepStrictEqual(names, ['新FMT', '新FMT2', '合計']);
  const values = slide.elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual(values.slice(0, 8), ['200', '100%', '200/3件', '20/1件', '220', '+10', '4件', '2件'], '新FMT の行');
  assert.deepStrictEqual(values.slice(8, 16), ['100', '20%', '20/1件', '0/0件', '30', '±0', '2件', '0件'], '新FMT2 の行');
  assert.deepStrictEqual(values.slice(16), ['300', '73%', '220/4件', '20/1件', '250', '+10', '6件', '2件'], '合計の行');
  // 前回比の色：増は青、減は赤、±0 は黒
  const delta = text => slide.elements.find(e => e.role === 'compact-value' && e.text === text);
  assert.strictEqual(delta('+10').color, c.plus);
  assert.strictEqual(delta('±0').color, c.ink);
  const minus = plain(g.mtgCompactOverviewSlide_([Object.assign({}, s1, {bridge: Object.assign({}, s1.bridge, {amountStart: s1.bridge.amountEnd + 5})})]));
  assert.strictEqual(minus.elements.find(e => e.role === 'compact-value' && /^▲/.test(e.text)).color, c.minus);
  // 横棒は A・B だけ（C は描かない）。データラベルは棒の中（白）か、入らなければ棒の右に「B 20」のように確度を付けて
  const bars = slide.elements.filter(e => e.role === 'compact-bar');
  const colorOf = key => o.plan.ranks.find(r => r.key === key).color;
  assert.deepStrictEqual(bars.map(b => b.fill), [colorOf('A'), colorOf('B'), colorOf('B')]);
  const labels = slide.elements.filter(e => e.role === 'compact-bar-label').map(e => e.text);
  assert.ok(labels.filter(t => /^(A |B )?100$/.test(t)).length === 2, 'A 100・B 100');
  assert.ok(labels.some(t => /^(B )?20$/.test(t)));
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
  assert.match(slide.lead, /^進捗率（A＋B）は 73%（計画に届いているのは 2のうち 1）。案件 6件・想定売上合計 250百万円（全期間・失注除く）。\nサブインダストリーはシートの順。.*の案件の想定売上/);
  assert.ok(slide.elements.some(e => /の確定（A＋B）の想定売上と計画（百万円）$/.test(e.text || '')));
});

test('状況・計画進捗：前回の会議でなく基準・直近14日間と比べるときは、前回比の見出しに日付を付けない。想定売上・月の列が無いシートは「データなし」', () => {
  const {gas, g} = setup();
  amountOnly(gas);
  const first = plain(g.mtgCompactOverviewSlide_([summarize(g, '新FMT', null)]));
  assert.ok(texts(first).indexOf('直近14日間の増減') >= 0);
  const s = summarize(g, '新FMT');
  const gap = Object.assign({}, s, {planAmount: Object.assign({}, s.planAmount, {gap: 'シートに「想定売上規模（百万）」の列が無い'})});
  const slide = plain(g.mtgCompactOverviewSlide_([gap, summarize(g, '新FMT2')]));
  assert.ok(texts(slide).some(t => /^データなし（シートに「想定売上/.test(t)));
  assert.deepStrictEqual(slide.elements.filter(e => e.role === 'compact-name').map(e => e.text), ['新FMT', '新FMT2', '合計'], '作れないシートもシートの順');
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
  assert.match(done, /^「統合版（報告用）（全サブインダストリー）」：\d{8}_全サブインダストリー_RX Business Meeting用資料_報告用\.pdf（6ページ）/);
  assert.match(done, /会議のフォルダ「\d{6}_第6回会議」に格納しました/);
  assert.ok(gas.driveFiles.some(f => f.folder === 'meeting-folder-0' && /_報告用\.pdf$/.test(f.name)));
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), ['mtg_p1', 'mtg_p2', 'mtg_p3', 'mtg_p4', 'mtg_p5', 'mtg_p6'], '前のページは最後に足して、番号の順に並べる');
  const index = gas.ss.getSheetByName(o.index.sheet);
  assert.ok(!index || !/_報告用/.test(JSON.stringify(index.getDataRange().getDisplayValues())), 'インデックスには書かない');
  assert.strictEqual(g.PropertiesService.getScriptProperties().getProperty('MTG_JOB'), null);
});

test('createBusinessMtgReportingOverview：全体サマリだけ（表紙・状況と計画進捗・計画のページ）を作り、変化のあった案件のページは作らない。保存先に「_報告用_全体サマリ」で保存し、会議のフォルダには格納しない', () => {
  const {gas, g, o} = setup();
  const meeting = g.mtgParseDate_('11/13', new Date());
  const parent = 'SubIndustryFolder0001';
  gas.driveFolders.push({id: parent, name: 'スーパー・GMS', parent: 'root'});
  gas.driveFolders.push({id: 'meeting-folder-0', name: meeting.slice(2, 4) + meeting.slice(5, 7) + meeting.slice(8, 10) + '_第6回会議', parent});
  o.meetingFolder = 'https://drive.google.com/drive/folders/' + parent;
  const data = g.getMeetingRunData('compact', false);
  assert.strictEqual(data.title, '統合版（報告用）の全体サマリ（全サブインダストリー）を作る');
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingOverview();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /^「統合版（報告用）の全体サマリ（全サブインダストリー）」：\d{8}_全サブインダストリー_RX Business Meeting用資料_報告用_全体サマリ\.pdf（4ページ）/);
  assert.ok(!/会議のフォルダ/.test(done), '会議のフォルダには格納しない');
  assert.ok(!gas.driveFiles.some(f => f.folder === 'meeting-folder-0'));
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), ['mtg_p1', 'mtg_p2', 'mtg_p3', 'mtg_p4']);
  const texts = deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.ok(texts.indexOf('全サブインダストリーの状況・計画進捗（FY') >= 0);
  assert.ok(texts.indexOf('変化のあった案件：') < 0, '変化のあった案件のページは作らない');
  assert.strictEqual(g.PropertiesService.getScriptProperties().getProperty('MTG_JOB'), null);
  // 確かめの文
  assert.match(g.checkMeetingRun('compact', false, '11/13', [], '').text, /の全体サマリ（2サブインダストリーの状況・計画進捗と計画に対する積み上げ。変化のあった案件のページは作りません）/);
});

test('createBusinessMtgReportingChanges：表紙と変化のあった案件のページだけを作る（状況・計画進捗・計画のページは作らない）。保存先に「_報告用_変化のあった案件」で保存し、会議のフォルダには格納しない', () => {
  const {gas, g, o} = setup();
  const meeting = g.mtgParseDate_('11/13', new Date());
  const parent = 'SubIndustryFolder0001';
  gas.driveFolders.push({id: parent, name: 'スーパー・GMS', parent: 'root'});
  gas.driveFolders.push({id: 'meeting-folder-0', name: meeting.slice(2, 4) + meeting.slice(5, 7) + meeting.slice(8, 10) + '_第6回会議', parent});
  o.meetingFolder = 'https://drive.google.com/drive/folders/' + parent;
  const data = g.getMeetingRunData('compact', 'changes');
  assert.strictEqual(data.title, '統合版（報告用）の変化のあった案件（全サブインダストリー）を作る');
  assert.strictEqual(data.final, false);
  // 確かめの文
  assert.match(g.checkMeetingRun('compact', 'changes', '11/13', [], '').text,
    /の変化のあった案件（表紙と、2サブインダストリーの変化のあった案件のページ。状況・計画進捗と計画に対する積み上げのページは作りません）/);
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingChanges();
  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /^「統合版（報告用）の変化のあった案件（全サブインダストリー）」：\d{8}_全サブインダストリー_RX Business Meeting用資料_報告用_変化のあった案件\.pdf（3ページ）/);
  assert.ok(!/会議のフォルダ/.test(done), '会議のフォルダには格納しない');
  assert.ok(!gas.driveFiles.some(f => f.folder === 'meeting-folder-0'));
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), ['mtg_p1', 'mtg_p2', 'mtg_p3'], '表紙の次から変化のあった案件');
  const texts = deck.requests.filter(r => r.insertText).map(r => r.insertText.text).join('\n');
  assert.ok(texts.indexOf('変化のあった案件：新FMT') >= 0 && texts.indexOf('変化のあった案件：新FMT2') >= 0);
  assert.ok(texts.indexOf('全サブインダストリーの状況・計画進捗') < 0, '状況・計画進捗のページは作らない');
  assert.ok(texts.indexOf(gas.get('MTG_TITLES').plan) < 0, '計画に対する積み上げのページは作らない');
  assert.strictEqual(g.PropertiesService.getScriptProperties().getProperty('MTG_JOB'), null);
  // 結果の控えは試作・本番の区別に関わらない（本番ではない）
  assert.strictEqual(g.getMeetingRunState('compact').final, false);
});

test('createBusinessMtgReportingChanges：時間がかかれば区切り、続きで残りと表紙を作る（ページ番号は表紙の次から通し）', () => {
  const {gas, g, o} = setup();
  const sprops = g.PropertiesService.getScriptProperties();
  const clock = {t: 0};
  g.mtgNow_ = () => clock.t;
  const add = g.mtgDeckAdd_;
  g.mtgDeckAdd_ = (...args) => { clock.t += 100000; return add(...args); };
  o.run.limitMs = 150000;
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingChanges();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /残りは約1分後に自動で続けます：変化のあった案件のページ 1件（新FMT2）・表紙と PDF。/);
  const saved = JSON.parse(sprops.getProperty('MTG_JOB'));
  assert.deepStrictEqual([saved.kind, saved.part, saved.page], ['compact', 'changes', 2]);
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^統合版（報告用）の変化のあった案件（全サブインダストリー）を作っている途中です/);
  o.run.limitMs = 400000;
  clock.t = 0;
  g.continueBusinessMtgSummaries();
  assert.strictEqual(sprops.getProperty('MTG_JOB'), null);
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), ['mtg_p1', 'mtg_p2', 'mtg_p3']);
  const pages = deck.requests.filter(r => r.insertText && /^\d+$/.test(r.insertText.text)).map(r => r.insertText.text);
  assert.ok(pages.indexOf('2') >= 0 && pages.indexOf('3') >= 0, 'ページ番号 2・3');
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /_報告用_変化のあった案件\.pdf（3ページ）/);
});

test('状況・計画進捗（FY26〜28売上）の中身：今年度は受注月が今年度の案件の売上を確度 A〜D 別に。来年度以降は受注月・確度によらない売上（半期ごとも）。全体合計は今年度（A〜D）＋来年度以降。失注は入れない', () => {
  const {g, o} = setup();
  const y0 = fiscalYear();
  const y1 = String(Number(y0) + 1);
  const y2 = String(Number(y0) + 2);
  const month = n => g.mtgMonth_(fiscalMonth(n));
  const deal = (status, n, y0Sales, halves) => ({status, planMonth: n == null ? null : month(n), sales: {[y0]: y0Sales},
    salesHalves: Object.keys(halves || {}).reduce((m, k) => (m[k.replace('Y1', y1).replace('Y2', y2)] = halves[k], m), {})});
  const deals = [
    deal('A-内示(90%)', 2, 30, {Y1前: 12, Y1後: 8, Y2前: 10}),        // 今年度 A 30。来年度以降はそのまま
    deal('B-クロージング中(70%)', 14, 5, {Y1前: 40}),                 // 受注月が来年度：今年度の分は入れない（来年度以降は入れる）
    deal('C-提案中(50%)', null, 15, {Y1後: 15}),                      // 受注月が空：同じ
    deal('D-初回提案(30%)', 0, 8),                                    // 今年度 D 8
    deal('創注-ニーズ探索(10%)', 0, 100, {Y1前: 6}),                   // 今年度は確度 A〜D の外。来年度以降は確度によらず入れる
    deal('失注', 0, 50, {Y1前: 99}),                                  // 失注はどれにも入れない
    deal('A-内示(90%)', -3, 7),                                       // 受注月が今年度より前：今年度の分は入れない
    deal('B-クロージング中(70%)', 30, 0, {Y1後: 9, Y2前: 9})           // 受注月が再来年度でも、来年度以降は入れる
  ];
  const table = {salesYears: [y0, y1, y2], has: {[o.plan.monthColumn]: true}};
  const out = plain(g.mtgSalesOutlookOf_(deals, '新FMT', new Date(), table));
  assert.deepStrictEqual(out.byYear, {[y0]: {A: 30, B: 0, C: 0, D: 8}});
  assert.deepStrictEqual(out.later, {[y1]: 90, [y2]: 19});
  assert.deepStrictEqual(out.halves.map(h => [h.year + h.half, h.value]), [[y1 + '前', 58], [y1 + '後', 32], [y2 + '前', 19], [y2 + '後', 0]]);
  assert.strictEqual(out.all, 38 + 90 + 19);
  assert.deepStrictEqual([out.year, out.years, out.plan, out.gap], [y0, [y0, y1, y2], 200, '']);
  // 今年度の売上の列・受注月の列が無ければ「データなし」の理由
  assert.match(g.mtgSalesOutlookOf_(deals, '新FMT', new Date(), {salesYears: [y1], has: {[o.plan.monthColumn]: true}}).gap, new RegExp('FY' + y0 + '前半期売上'));
  assert.match(g.mtgSalesOutlookOf_(deals, '新FMT', new Date(), {salesYears: [y0], has: {}}).gap,
    new RegExp('「' + String(o.columns[o.plan.monthColumn]).replace(/\s+/g, '') + '」の列が無い'));
  // 会議サマリの中身と、区切って作るときの中身の一部に入る
  const s1 = summarize(g, '新FMT');
  assert.deepStrictEqual(plain(s1.salesOutlook).byYear[y0], {A: 70, B: 40, C: 0, D: 0}, 'テストのシート：A社 70・B社 30＋C社 10（E社は失注）');
  assert.deepStrictEqual(plain(g.mtgCombinedLite_(s1)).salesOutlook, plain(s1.salesOutlook));
});

test('状況・計画進捗（FY26〜28売上）のページ：棒は今年度の A〜D、右に計画・進捗率・確定（A＋B）・見込（C）（FY26）｜売上（FY27・FY28。受注月・確度によらない）｜全体合計。件数は出さない。列の無いシートは「データなし」、合計に ※', () => {
  const {gas, g} = setup();
  const y0 = fiscalYear();
  const y1 = String(Number(y0) + 1);
  const y2 = String(Number(y0) + 2);
  const span = 'FY' + y0 + '〜' + y2;
  const now = new Date();
  const outlook = (plan, a0, b0, c0, d0, l1, l2) => ({year: y0, years: [y0, y1, y2], plan, all: a0 + b0 + c0 + d0 + l1 + l2, gap: '',
    byYear: {[y0]: {A: a0, B: b0, C: c0, D: d0}}, later: {[y1]: l1, [y2]: l2}});
  const summaries = [
    {name: '新FMT', now, salesOutlook: outlook(200, 70, 40, 10, 5, 50, 20)},
    {name: '新FMT2', now, salesOutlook: outlook(100, 0, 10, 0, 0, 5, 0)},
    {name: '新FMT3', now, salesOutlook: {year: y0, years: [y0, y1, y2], plan: 50, all: 0, gap: 'シートに「FY' + y0 + '前半期売上」「FY' + y0 + '後半期売上」の列が無い',
      byYear: {}, later: {}}}
  ];
  const slide = plain(g.mtgCompactOutlookSlide_(summaries));
  assert.strictEqual(slide.title, '全サブインダストリーの状況・計画進捗（' + span + '売上）');
  const heads = slide.elements.filter(e => e.role === 'compact-head').map(e => e.text);
  assert.deepStrictEqual(heads, ['計画', '進捗率（A＋B）', '確定（A＋B）', '見込（C）', '売上', '売上', '全体合計']);
  const periods = slide.elements.filter(e => e.role === 'compact-head-period').map(e => e.text);
  assert.deepStrictEqual(periods, ['（FY' + y0 + '）', '（FY' + y0 + '）', '（FY' + y0 + '）', '（FY' + y0 + '）', '（FY' + y1 + '）', '（FY' + y2 + '）', '（' + span + '）']);
  slide.elements.filter(e => e.role === 'compact-head' || e.role === 'compact-head-period').forEach(e => {
    assert.ok(g.mtgTextWidth_(e.text, e.size) <= e.w - 14.4 + 0.01, e.text + ' は1行（' + e.size + 'pt）');
  });
  const values = slide.elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual(values.slice(0, 7), ['200', '55%', '110', '10', '50', '20', '195'], '新FMT の行（件数は出さない）');
  assert.deepStrictEqual(values.slice(7, 14), ['100', '10%', '10', '0', '5', '0', '15'], '新FMT2 の行');
  const G = gas.get('MTG_OPTIONS').gapText;
  assert.deepStrictEqual(values.slice(14, 21), [G, '—', G, G, G, G, G], '列の無いシート');
  assert.deepStrictEqual(values.slice(21), ['300※', '40%', '120※', '10※', '55※', '20※', '210※'], '合計は作れたシートの分だけ（※）');
  assert.strictEqual(slide.note, '※ 全体合計 ＝ FY' + y0 + '（A〜D）＋FY' + y1 + '＋FY' + y2 + '。※ のある合計は「' + G + '」のシートを除く');
  // 棒：今年度の A〜D を積み上げ（新FMT は A・B・C・D の4本、新FMT2 は B の1本）。棒の右に A〜D の合計。計画線は2本
  assert.strictEqual(slide.elements.filter(e => e.role === 'compact-bar').length, 5);
  assert.deepStrictEqual(slide.elements.filter(e => e.role === 'compact-bar-total').map(e => e.text), ['125', '10']);
  assert.strictEqual(slide.elements.filter(e => e.role === 'compact-plan').length, 2);
  assert.strictEqual(slide.elements.filter(e => e.role === 'compact-gap').length, 1);
  assert.deepStrictEqual(slide.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text), ['A', 'B', 'C', 'D']);
  assert.match(slide.lead, new RegExp('^FY' + y0 + ' の確定（A＋B）は 120百万円（計画比 40%）、A〜D の合計は 135百万円。FY' + y1 + ' 55・FY' + y2 + ' 20百万円（失注除く）。\n'));
  assert.match(slide.lead, new RegExp('FY' + y1 + '・FY' + y2 + ' は' + g.mtgPlanMonthName_() + '・確度によらない。$'));
  // 列の無いシートが無ければ、補足は全体合計の数え方だけ
  assert.strictEqual(plain(g.mtgCompactOutlookSlide_(summaries.slice(0, 2))).note, '※ 全体合計 ＝ FY' + y0 + '（A〜D）＋FY' + y1 + '＋FY' + y2);
});

test('計画に対する積み上げ（FY26〜28売上）：今年度は受注月の四半期ごとに確度 A〜D を続けて積み上げ（累計を状況・計画進捗の全体合計とそろえる）、合計・計画・不足。来年度以降は半期ごと（受注月・確度によらない1色）に今年度の合計から続けて積み上げ、累計。確定（A＋B）だけのページも', () => {
  const {gas, g} = setup();
  const T = gas.get('MTG_TITLES');
  const y0 = fiscalYear();
  const y1 = String(Number(y0) + 1);
  const y2 = String(Number(y0) + 2);
  const span = 'FY' + y0 + '〜' + y2;
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  // 来年度以降の半期（テストのシートには列が無いので、中身に入れる）
  s1.salesOutlook.halves = [{year: y1, half: '前', value: 25}, {year: y1, half: '後', value: 10}, {year: y2, half: '前', value: 4}, {year: y2, half: '後', value: 0}];
  const now = new Date();
  const sales = g.mtgMergePlans_([s1.planSalesByMonth, s2.planSalesByMonth]);
  sales.year = y0;
  // 創注・未設定の今年度の売上（状況・計画進捗の全体合計と同じく、確度 A〜D だけを積み上げるので入れない）
  const other = gas.get('MTG_OPTIONS').plan.otherLabel;
  sales.months[8].values['創注'] = 7;
  sales.months[8].values[other] = 3;
  sales.months[8].total += 10;
  sales.total += 10;
  const valuesOf = slide => slide.elements.filter(e => e.role === 'plan-years-value').map(e => e.text);
  const axisOf = slide => slide.elements.filter(e => e.role === 'plan-years-axis').map(e => e.text.split('\n')[0]);
  // 新FMT：A社 70（12月。Q3）・B社 30（1月。Q4）・C社 10（2月。Q4。B に前進）、新FMT2：F社 10（12月。Q3）。計画 300
  const all = plain(g.mtgCompactPlanYearsSlide_(sales, [s1, s2], now));
  assert.strictEqual(all.title, T.plan + '（' + span + '売上）');
  assert.deepStrictEqual(axisOf(all), ['Q1', 'Q2', 'Q3', 'Q4', 'FY' + y0, '計画', '不足', 'FY' + y1, 'FY' + y1, 'FY' + y2, 'FY' + y2, '累計']);
  assert.deepStrictEqual(valuesOf(all), ['+80', '+40', '120', '300', '▲180', '+25', '+10', '+4', '159']);
  assert.strictEqual(all.lead, 'FY' + y0 + '売上（A〜D。創注・' + other + 'を除く）の積み上げは 120百万円で、計画 300百万円に対して ▲180（達成率 40%）。\nFY' + y1 +
    ' 35・FY' + y2 + ' 4を続けた ' + span + 'の累計は 159百万円（状況・計画進捗の全体合計と同じ。失注除く）。');
  assert.ok(!all.note, '受注月が空・今年度の外の案件の売上が無ければ補足なし');
  assert.deepStrictEqual(all.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text).slice(0, 4), ['A', 'B', 'C', 'D'], '創注・未設定は凡例にも出さない');
  // 凡例は確度 A〜D・来年度以降の年度・半期・計画。来年度以降の棒は確度によらず、年度ごとの色味の濃淡（compact.laterColors）
  const halfKeys = ['FY' + y1 + '前', 'FY' + y1 + '後', 'FY' + y2 + '前', 'FY' + y2 + '後'];
  assert.deepStrictEqual(all.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text).slice(-5), halfKeys.concat(['計画']));
  const colors = gas.get('MTG_OPTIONS').compact.laterColors;
  const bars = all.elements.filter(e => e.role === 'plan-years-bar');
  [colors[0][0], colors[0][1], colors[1][0]].forEach((color, n) => {
    assert.strictEqual(bars.filter(e => e.fill === color).length, 2, halfKeys[n] + '：その半期の棒と、累計の上の1つ');
  });
  assert.ok(colors[0][0] !== colors[1][0] && colors[0][0] !== colors[0][1], '年度・半期で色が違う');
  // 確定（A＋B）だけ：今年度は A・B の分だけ。来年度以降は確度によらないので同じ値
  const sure = plain(g.mtgCompactPlanYearsSlide_(sales, [s1, s2], now, {only: ['A', 'B'], label: '確定（A＋B）'}));
  assert.strictEqual(sure.title, T.plan + '（' + span + '売上）：確定（A＋B）');
  assert.deepStrictEqual(valuesOf(sure), ['+80', '+40', '120', '300', '▲180', '+25', '+10', '+4', '159']);
  assert.deepStrictEqual(sure.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text), ['A', 'B'].concat(halfKeys, ['計画']));
  assert.match(sure.lead, /^確定（A＋B）の FY\d+売上の積み上げは 120百万円で、計画 300百万円に対して ▲180（進捗率 40%）。\n.*の累計は 159百万円（失注除く）。$/);
  // 計画を超えたら「超過」
  const over = plain(g.mtgCompactPlanYearsSlide_(Object.assign({}, sales, {plan: 100}), [s1, s2], now));
  assert.deepStrictEqual(axisOf(over)[6], '超過');
  assert.ok(valuesOf(over).indexOf('+20') >= 0 && !over.elements.some(e => e.role === 'plan-years-gap'));
  // 受注月が空の案件の今年度の売上は入れず、補足に金額
  const monthCol = String.fromCharCode(64 + COL[gas.get('MTG_OPTIONS').plan.monthColumn]);   // 積み上げの月の列（テストの設定）
  g.recordDiffEdit(gas.edit(gas.ss.getSheetByName('新FMT2'), monthCol + '3', ''));   // F社：月を空に
  const t2 = summarize(g, '新FMT2');
  const sales2 = g.mtgMergePlans_([s1.planSalesByMonth, t2.planSalesByMonth]);
  sales2.year = y0;
  const noted = plain(g.mtgCompactPlanYearsSlide_(sales2, [s1, t2], now));
  assert.deepStrictEqual(valuesOf(noted).slice(0, 3), ['+70', '+40', '110']);
  assert.match(noted.note, new RegExp('^※ ' + g.mtgPlanMonthName_() + 'が空・FY' + y0 + '外の FY' + y0 + '売上 10百万円は除く$'));
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
  assert.deepStrictEqual([saved.kind, saved.page, saved.no], ['compact', 5, 2]);
  // 途中：ほかの資料・統合版はすぐには作らない。結果の画面に途中と出す
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^統合版（報告用）（全サブインダストリー）を作っている途中です（会議日 \d{4}\/11\/13。1\/2サブインダストリーのページ）。/);
  // 途中にもう一度：画面に作っている途中の知らせ（「やめて新しく作る」を選ばなければ作らない）
  gas.confirmAnswer = 'NO';
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingSummary();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /^前に始めた統合版（報告用）（全サブインダストリー）（会議日 \d{4}\/11\/13。1\/2サブインダストリーのページ。.*「やめて新しく作る」にチェックを入れてください/);
  assert.ok(sprops.getProperty('MTG_JOB'), '仕事はそのまま');
  gas.confirmAnswer = 'YES';
  o.run.limitMs = 400000;
  clock.t = 0;
  g.continueBusinessMtgSummaries();
  assert.strictEqual(sprops.getProperty('MTG_JOB'), null);
  const deck = gas.decks[gas.decks.length - 1];
  assert.deepStrictEqual(deck.slides.map(sl => sl.objectId), once.slides);
  assert.deepStrictEqual(deck.requests.filter(r => r.insertText).map(r => r.insertText.text).filter(t => !/\d{1,2}:\d{2}/.test(t)).sort(), once.texts,
    '区切らずに作ったときと同じ文字');
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /「統合版（報告用）（全サブインダストリー）」：.*_報告用\.pdf（6ページ）/);
});

test('シートと記録を読むのに時間がかかっても、1回の実行で必ず1シート以上進める（読んだあとに区切ると、続きも同じだけ読んで区切り、終わらなかった）', () => {
  const {gas, g, o} = setup();
  const sprops = g.PropertiesService.getScriptProperties();
  const clock = {t: 0};
  g.mtgNow_ = () => clock.t;
  const read = g.mtgJobSource_;
  g.mtgJobSource_ = (...args) => { clock.t += 250000; return read(...args); };   // 読むのに約4分（見積もりの残りが足りない）
  o.run.limitMs = 320000;
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingOverview();
  const runs = [];
  for (let k = 0; k < 6 && sprops.getProperty('MTG_JOB'); k++) {
    const before = JSON.parse(sprops.getProperty('MTG_JOB')).items;
    runs.push(before);
    clock.t = 0;
    g.continueBusinessMtgSummaries();
  }
  assert.strictEqual(sprops.getProperty('MTG_JOB'), null, '終わった（続き ' + runs.length + '回）');
  assert.ok(runs.length <= 2, '実行ごとに進む：' + JSON.stringify(runs));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /全体サマリ/);
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /_報告用_全体サマリ\.pdf（4ページ）/);
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

test('統合版（報告用）：続きの実行をキャンセルして止まったままの仕事（続きの予約なし）も、画面の「やめて新しく作る」で片付けて作り直せる（作りかけの資料はゴミ箱へ）', () => {
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
  const first = gas.decks[gas.decks.length - 1].id;
  // 続きの実行をキャンセルした：仕事は「実行中」のまま、続きのトリガーも無い
  const job = JSON.parse(sprops.getProperty('MTG_JOB'));
  job.running = true;
  sprops.setProperty('MTG_JOB', JSON.stringify(job));
  g.ScriptApp.getProjectTriggers().forEach(t => g.ScriptApp.deleteTrigger(t));
  g.showBusinessMtgSummariesResult();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /続きの予約がありません（続きの実行をキャンセルした など）。createBusinessMtgReportingSummary\(\) を実行し、画面で「やめて新しく作る」にチェックを入れて作り直してください/);
  // 作り直す：画面に作っている途中の知らせ →「やめて新しく作る」を選んで作る
  g.mtgDeckAdd_ = add;
  clock.t = 0;
  o.run.limitMs = 400000;
  gas.promptAnswers = ['11/13'];
  g.createBusinessMtgReportingSummary();
  assert.ok(gas.alerts.some(a => /^前に始めた統合版（報告用）/.test(a.message)), '作っている途中の知らせ');
  assert.ok(gas.trashed.indexOf(first) >= 0, '作りかけの資料はゴミ箱へ');
  assert.strictEqual(sprops.getProperty('MTG_JOB'), null, '新しく作り終えた');
  assert.match(gas.alerts[gas.alerts.length - 1].message, /_報告用\.pdf（6ページ）/);
  assert.notStrictEqual(gas.decks[gas.decks.length - 1].id, first);
});

test('No. は3桁でも折り返さない（1行に入るまで字を小さくする）', () => {
  const {g} = setup();
  const s = summarize(g, '新FMT');
  const t = tableOf(plain(g.mtgCompactChangeSlides_(s, 378)).slides[0]);
  const cells = t.rows.slice(1).map(r => r.cells[0]);
  assert.deepStrictEqual(cells.map(c => c.text), ['379', '380'], '改行は入らない');
  cells.forEach(c => assert.ok(g.mtgWrap_(c.text, g.mtgCellWidth_(t.columns[0]), c.size, 9).length === 1, c.text + ' は1行'));
});

test('計画に対する積み上げ（確定（A＋B））：A・B の案件だけを月ごとに積み上げ、凡例も A・B。積み上げに含まない案件も A・B の分だけ', () => {
  const {gas, g} = setup();
  amountOnly(gas);
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const merged = g.mtgMergePlans_([s1.planAmount, s2.planAmount]);
  const only = plain(g.mtgPlanRanksOnly_(merged, ['A', 'B']));
  assert.strictEqual(only.total, 100 + 50 + 40 + 20, '想定売上：A社 100・B社 50・C社（B に前進）40・F社 20');
  assert.deepStrictEqual(Object.keys(only.byRank).sort(), ['A', 'B']);
  assert.strictEqual(only.months.reduce((t, mo) => t + mo.total, 0), only.total);
  assert.strictEqual(only.byRankCount.B, 3);
  assert.strictEqual(only.excluded.count, 0, '失注は確度が無いので入らない');
  const slide = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true))[4];
  const legend = slide.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text);
  assert.deepStrictEqual(legend, ['A', 'B', '計画']);
  const lead = leadOf(slide);
  assert.match(lead, /^今年度（[^）]+）の確定（A＋B）の想定売上の積み上げは 210百万円。計画 300百万円に対して ▲90（進捗率 70%）。/);
  assert.ok(slide.elements.some(e => /月別 確定（A＋B）の想定売上の積み上げ（百万円）$/.test(e.text || '')));
  // 全確度のページも想定売上（統合版と同じ形）
  const all = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true))[3];
  assert.match(leadOf(all), /^今年度（[^）]+）の想定売上の積み上げは 240百万円。/);
  const colors = gas.get('MTG_OPTIONS').plan.ranks.filter(r => r.key === 'A' || r.key === 'B').map(r => r.color);
  assert.ok(slide.elements.filter(e => e.role === 'plan-month' || e.role === 'plan-total').every(e => colors.indexOf(e.fill) >= 0), '棒は A・B の色だけ');
});

test('統合版（報告用）の金額は小数点以下を繰り上げる（12.1 → 13）。ほかの資料の金額の出し方は変えない', () => {
  const {gas, g} = setup();
  amountOnly(gas);
  const sheet = gas.ss.getSheetByName('新FMT');
  g.recordDiffEdit(gas.edit(sheet, 'G3', 100.2));   // A社：想定売上 100 → 100.2
  g.recordDiffEdit(gas.edit(sheet, 'H3', 90.2));    // 期待値も（想定売上だけの変更は「変化のあった案件」に入らないため）
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const values = slides[1].elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.strictEqual(values[2], '191/3件', '確定 190.2 → 191');
  assert.ok(!slides.some(sl => sl.elements.some(e => /\d\.\d/.test(String(e.text || '')) && /▶|百万円|件/.test(String(e.text || '')))), '小数は出さない');
  const rows = cellTexts(tableOf(slides[5])).slice(1);
  assert.ok(rows.some(r => r[3] === '100 ▶ 101'), '変化のあった案件の想定売上も繰り上げ：' + JSON.stringify(rows.map(r => r[3])));
  // 計画に対する積み上げの棒の上の数も、リード文と同じ繰り上げ（先に丸めない）
  const plan = slides[4];
  const gapText = plan.elements.filter(e => e.role === 'plan-value').map(e => e.text).find(t => /^▲/.test(t));
  assert.ok(leadOf(plan).indexOf(gapText) >= 0, '不足の数 ' + gapText + ' はリード文と同じ');
  // 終わったら元に戻す（ほかの資料は小数1桁）
  assert.strictEqual(g.mtgAmount_(90.2), '90.2');
  assert.strictEqual(gas.get('MTG_OPTIONS').amountCeil, undefined);
});

test('変化のあった案件は、確度 A から今回のステータスの順、その中は想定売上の大きい順。No. もその順', () => {
  const {gas, g} = setup();
  amountOnly(gas);
  const sheet = gas.ss.getSheetByName('新FMT');
  g.recordDiffEdit(gas.edit(sheet, 'G6', 500));   // D社（C）：想定売上 500（いちばん大きいが、確度 C なので B の案件のあと）
  g.recordDiffEdit(gas.edit(sheet, 'J4', '11/4 価格を再提示'));   // B社（B）：想定売上 50
  g.recordDiffEdit(gas.edit(sheet, 'J3', '11/4 契約書を確認'));   // A社（A）：想定売上 100
  const part = plain(g.mtgCompactChangeSlides_(summarize(g, '新FMT'), 0));
  const rows = cellTexts(tableOf(part.slides[0])).slice(1);
  assert.deepStrictEqual(rows.map(r => [r[0], r[1].split('\n')[0], r[3]]),
    [['1', 'A社', '100'], ['2', 'B社', '50'], ['3', 'C社', '40'], ['4', 'D社', '20 ▶ 500']]);
  assert.match(part.slides[0].lead, /確度 A から順に、同じ確度は想定売上の大きい順/);
});

test('変化のあった案件（既定）：金額は今年度の売上の列（前半期＋後半期）。確度 A から今回のステータスの順、その中は今年度の売上の大きい順', () => {
  const {gas, g} = setup();
  const fy = fiscalYear();
  const sheet = gas.ss.getSheetByName('新FMT');
  g.recordDiffEdit(gas.edit(sheet, 'G6', 500));   // D社（C）：想定売上は大きいが、今年度の売上は無い
  g.recordDiffEdit(gas.edit(sheet, 'J4', '11/4 価格を再提示'));   // B社（B）：売上 30
  g.recordDiffEdit(gas.edit(sheet, 'J3', '11/4 契約書を確認'));   // A社（A）：売上 70
  const part = plain(g.mtgCompactChangeSlides_(summarize(g, '新FMT'), 0));
  const table = cellTexts(tableOf(part.slides[0]));
  assert.strictEqual(table[0][3], 'FY' + fy + '売上\n（今回）');
  const blank = gas.get('MTG_OPTIONS').blankText;
  // C社は B に前進（売上 10）で B社（30）のあと。D社は C で売上なし
  assert.deepStrictEqual(table.slice(1).map(r => [r[0], r[1].split('\n')[0], r[3]]),
    [['1', 'A社', '70'], ['2', 'B社', '30'], ['3', 'C社', '10'], ['4', 'D社', blank]]);
  assert.match(part.slides[0].lead, new RegExp('確度 A から順に、同じ確度はFY' + fy + '売上の大きい順'));
});

test('状況・計画進捗の次のページ（見込（C）を含む）：棒の確定（A・B）の右に、見込（C）を薄い地と点線の枠で足す。並び・右の数字は前のページと同じ', () => {
  const {gas, g} = setup();
  amountOnly(gas);
  const o = gas.get('MTG_OPTIONS');
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const [base, withC] = [slides[1], slides[2]];
  assert.strictEqual(withC.title, '全サブインダストリーの状況・計画進捗（見込（C）を含む）');
  const values = sl => sl.elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual(values(withC), values(base), '右の数字は同じ');
  const names = sl => sl.elements.filter(e => e.role === 'compact-name').map(e => e.text);
  assert.deepStrictEqual(names(withC), names(base), '並びは同じ');
  const cColor = o.plan.ranks.find(r => r.key === 'C').color;
  const dashed = withC.elements.filter(e => e.role === 'compact-bar' && e.prospect);
  assert.strictEqual(dashed.length, 1, '新FMT の見込（C）D社 20');
  assert.deepStrictEqual([dashed[0].fill, dashed[0].alpha], [cColor, o.compact.prospect.alpha], 'C の色の薄い地');
  assert.ok(withC.elements.some(e => e.role === 'compact-bar-edge' && e.dashStyle === 'DOT'), '細かい点線');
  assert.ok(!base.elements.some(e => e.role === 'compact-bar' && e.prospect), '前のページには見込（C）は無い');
  assert.deepStrictEqual(withC.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text), ['A', 'B', 'C']);
  assert.match(leadOf(withC), /^確定（A＋B）に見込（C）を足すと 230百万円で、計画の 77%（進捗率（A＋B）は 70%）。/);
});

test('見込（C）を含むページ：見込の棒は確定の棒に続け、薄い地と上・下・右の細かい点線（DOT）。凡例の見本は A・B と同じ大きさ。入らない値は、いちばん右の確度だけ棒のすぐ右に出す', () => {
  const {gas, g} = setup();
  const o = gas.get('MTG_OPTIONS');
  const P = o.compact.prospect;
  const s1 = summarize(g, '新FMT');
  const withC = plain(g.mtgCompactOverviewSlide_([s1], {withC: true}));
  const bars = withC.elements.filter(e => e.role === 'compact-bar');
  const cBar = bars.find(e => e.prospect);
  const bBar = bars[bars.indexOf(cBar) - 1];
  assert.deepStrictEqual([cBar.alpha, cBar.line], [P.alpha, undefined], '見込の棒は薄い地だけ（枠なし）');
  assert.ok(Math.abs(cBar.x - (bBar.x + bBar.w)) < 0.01, 'B の端に続ける（あけない）');
  // 点線は上・下・右の3辺（左の辺は B との境目なので引かない）
  const edges = withC.elements.filter(e => e.role === 'compact-bar-edge');
  assert.strictEqual(edges.length, 3);
  assert.ok(edges.every(e => e.dashStyle === P.dashStyle && e.color === P.line));
  assert.ok(!edges.some(e => e.w === 0 && Math.abs(e.x - cBar.x) < 0.01), '左の辺は無い');
  assert.ok(edges.some(e => e.w === 0 && Math.abs(e.x - (cBar.x + cBar.w)) < 0.01), '右の辺');
  const swatch = withC.elements.find(e => e.role === 'legend' && e.dashStyle);
  const swatches = withC.elements.filter(e => e.role === 'legend' && !e.text);
  assert.ok(swatches.every(e => e.w === swatch.w && e.h === swatch.h), '凡例の見本は A・B・C とも同じ大きさ');
  const legendText = withC.elements.filter(e => e.role === 'legend' && e.text);
  swatches.forEach((sw, i) => assert.ok(legendText[i].x + 7.2 - (sw.x + sw.w) >= 5 - 0.01, '見本と文字の間は 5pt 以上'));
  // 間の確度の値が入らなくても、離れた場所（棒の右）には出さない
  const narrow = Object.assign({}, s1, {planAmount: Object.assign({}, s1.planAmount, {byRank: {A: 1, B: 150, C: 0}})});
  const labels = plain(g.mtgCompactOverviewSlide_([narrow])).elements.filter(e => e.role === 'compact-bar-label').map(e => e.text);
  assert.ok(labels.indexOf('150') >= 0);
  assert.ok(!labels.some(t => /^A /.test(t)), '間の確度（A）の値は棒の右に出さない');
  const last = Object.assign({}, s1, {planAmount: Object.assign({}, s1.planAmount, {byRank: {A: 150, B: 1, C: 0}})});
  assert.ok(plain(g.mtgCompactOverviewSlide_([last])).elements.some(e => e.role === 'compact-bar-label' && e.text === 'B 1'), 'いちばん右の確度は棒のすぐ右に');
  // Slides API：細かい点線（線の種類 DOT）
  const req = plain(g.mtgSlideRequests_({title: 't', elements: [edges[0]]}, 'p1', 1)).find(r => r.updateLineProperties);
  assert.strictEqual(req.updateLineProperties.lineProperties.dashStyle, 'DOT');
  const swatchReq = plain(g.mtgSlideRequests_({title: 't', elements: [swatch]}, 'p1', 1)).find(r => r.updateShapeProperties);
  assert.strictEqual(swatchReq.updateShapeProperties.shapeProperties.outline.dashStyle, 'DOT', '凡例の見本は点線の枠');
});

test('リード文は1つのテキストボックス（1つ目の文と補足を行で分け、行の送りは文字の大きさ × leadLine）', () => {
  const {gas, g} = setup();
  const S = gas.get('MTG_SLIDE');
  const slides = plain(g.mtgCompactSlides_([summarize(g, '新FMT'), summarize(g, '新FMT2')], '2026/11/13', true));
  slides.filter(sl => !sl.cover).forEach(sl => {
    const leads = sl.elements.filter(e => e.role === 'lead');
    assert.strictEqual(leads.length, 1, sl.title + ' のリード文は1つの箱');
    assert.ok(Math.abs(leads[0].lineSpacing - S.leadLine / S.lineFactor * 100) < 0.1, '行間');
  });
  // Slides API：段落の行間に入る
  const lead = slides[1].elements.find(e => e.role === 'lead');
  const req = plain(g.mtgSlideRequests_({title: 't', elements: [lead]}, 'p1', 1)).find(r => r.updateParagraphStyle);
  assert.strictEqual(req.updateParagraphStyle.style.lineSpacing, lead.lineSpacing);
});

test('状況のページの「変化のある案件」は、今ある案件（失注を除く。案件数と同じ）だけを数える。削除した行・失注になった案件は数えない', () => {
  const {gas, g} = setup();
  const sheet = gas.ss.getSheetByName('新FMT');
  g.recordDiffEdit(gas.edit(sheet, 'F4', '失注'));   // B社：失注になった（変化はあるが、案件数に入らない）
  const s1 = summarize(g, '新FMT');
  assert.ok(s1.kpi.changed >= 3, '資料の「変化のある案件」には失注も入る');
  assert.strictEqual(g.mtgCompactChangedNow_(s1), 2, 'C社・D社だけ');
  const values = plain(g.mtgCompactOverviewSlide_([s1])).elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual(values.slice(6, 8), ['3件', '2件'], '案件数 3件（A社・C社・D社）・変化のある案件 2件');
  // 区切って作るときの中身の一部にも入れる
  assert.strictEqual(plain(g.mtgCombinedLite_(s1)).changedNow, 2);
  assert.strictEqual(g.mtgCompactChangedNow_(plain(g.mtgCombinedLite_(s1))), 2);
  // 変化のあった案件のページには失注も載せ、リード文にその件数を書き添える（状況・計画進捗の件数と見比べられるように）
  const part = plain(g.mtgCompactChangeSlides_(s1, 0));
  assert.match(part.slides[0].lead, /^前回から変化のある案件は 3件（うち失注・削除 1件。状況・計画進捗の件数には含めない）/);
});

test('状況・計画進捗（FY売上）：今年度の売上の列（前半期＋後半期。本物の P・Q列）で、棒・計画・進捗率・確定・見込を集計したページ（受注月が今年度の案件だけ。失注は除く）', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').compact.outlookPage = false;   // 前の FY26売上のページ（FY26〜28 の版に差し替える前）
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  const fy = 'FY' + fiscalYear();
  // 新FMT：A社 70・B（B社 30 + C社（B に前進）10）= 40 → 確定 110/3件、C は D社（売上なし）で 0/0件。E社（失注）の 50 は入れない。計画 200 → 55%
  assert.deepStrictEqual([s1.planSales.byRank.A, s1.planSales.byRank.B, s1.planSales.byRank.C, s1.planSales.total], [70, 40, 0, 110]);
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const slide = slides.find(sl => sl.title === '全サブインダストリーの状況・計画進捗（' + fy + '売上）');
  assert.ok(slide && slides.indexOf(slide) === 1, '表紙の次（集計の元は今年度の売上）');
  const values = slide.elements.filter(e => e.role === 'compact-value').map(e => e.text);
  // 列：計画・進捗率・確定（A＋B）/件・前回比（確定（A＋B）の売上の増減。確定の右）・見込（C）/件｜売上合計・案件数・変化のある案件
  // 新FMT の前回比 +10：C社（売上 10）が C → B に前進して確定に入った
  assert.deepStrictEqual(values.slice(0, 5), ['200', '55%', '110/3件', '+10', '0/0件'], '新FMT の行（左の5列は売上）');
  assert.deepStrictEqual(values.slice(8, 13), ['100', '10%', '10/1件', '±0', '0/0件'], '新FMT2 の行');
  // 売上合計は3年度分（今年度から）の売上（テストのシートは今年度の列だけ。新FMT は E社（失注）を除く 110）。
  // 案件数・変化のある案件は、想定売上のページ（amountPages）と同じ
  assert.deepStrictEqual([values[5], values[13]], ['110', '10']);
  const heads = slide.elements.filter(e => e.role === 'compact-head').map(e => e.text);
  const periods = slide.elements.filter(e => e.role === 'compact-head-period').map(e => e.text);
  assert.deepStrictEqual(heads, ['計画', '進捗率（A＋B）', '確定（A＋B）/件', heads[3], '見込（C）/件', '売上合計', '案件数', '変化のある案件']);
  assert.match(heads[3], /^前回比（\d+\/\d+）$/);
  // 期間の字は折り返さない（「（FY26〜28）」も1行。案件数の列は広げてある）
  slide.elements.filter(e => e.role === 'compact-head-period').forEach(e => {
    assert.ok(g.mtgTextWidth_(e.text, e.size) <= e.w - 14.4 + 0.01, e.text + ' は1行（' + e.size + 'pt）');
  });
  // 今年度の列は「（FY26）」（想定売上のページは「（今年度）」）
  assert.deepStrictEqual(periods, ['（' + fy + '）', '（' + fy + '）', '（' + fy + '）', '（' + fy + '）', '（' + fy + '）',
    '（' + fy + '〜' + (Number(fiscalYear()) + 2) + '）', '（' + fy + '〜' + (Number(fiscalYear()) + 2) + '）', '（' + fy + '〜' + (Number(fiscalYear()) + 2) + '）']);
  assert.match(leadOf(slide), new RegExp('案件 6件・' + fy + '〜' + (Number(fiscalYear()) + 2) + '売上合計 120百万円（失注除く）'));
  gas.get('MTG_OPTIONS').compact.amountPages = true;
  const both = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  gas.get('MTG_OPTIONS').compact.amountPages = false;
  const amountPage = both.find(sl => sl.title === '全サブインダストリーの状況・計画進捗');
  assert.strictEqual(both.indexOf(amountPage), 3, '想定売上のページは、売上の2枚のあと');
  const base = amountPage.elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual([values.slice(6, 8), values.slice(14, 16)], [base.slice(6, 8), base.slice(14, 16)]);
  assert.deepStrictEqual([base[4], base[12]], ['210', '30'], '想定売上のページは全期間の想定売上合計');
  // 想定売上のページは前と同じ列（前回比は全期間の想定売上の増減で、想定売上合計の右）
  assert.deepStrictEqual(amountPage.elements.filter(e => e.role === 'compact-head').map(e => e.text).slice(2, 6),
    ['確定（A＋B）/件', '見込（C）/件', '想定売上合計', heads[3]]);
  assert.strictEqual(amountPage.elements.find(e => e.role === 'compact-head-period').text, '（今年度）');
  assert.deepStrictEqual(amountPage.elements.filter(e => e.role === 'compact-head-period').map(e => e.text).slice(4), ['（全期間）', '（全期間）', '（全期間）', '（全期間）']);
  assert.ok(slide.elements.some(e => e.text === fy + ' の売上（前半期＋後半期）の確定（A＋B）と計画（百万円）'));
  assert.match(leadOf(slide), new RegExp('確定・見込は.+が今年度（[^）]+）の案件の ' + fy + '売上。'));
  // 受注月が時期未定・今年度より前・来年度以降の案件は、確定・見込に入れない（計画に対する積み上げと同じ）
  const sheet = gas.ss.getSheetByName('新FMT');
  const col = gas.get('MTG_OPTIONS').plan.monthColumn === 'startMonth' ? 'M' : 'I';
  g.recordDiffEdit(gas.edit(sheet, col + '3', ''));            // A社（70）：時期未定
  g.recordDiffEdit(gas.edit(sheet, col + '4', fiscalMonth(-1)));   // B社（30）：今年度より前
  const moved = summarize(g, '新FMT');
  assert.deepStrictEqual([moved.planSales.byRank.A, moved.planSales.byRank.B, moved.planSales.undated.value, moved.planSales.before.value], [0, 10, 70, 30]);
  assert.deepStrictEqual([moved.salesBridge.sureEnd, moved.salesBridge.sureStart], [10, 100], '前回比の今の確定も同じく絞る（前回は前回の受注月）');
  // 状況・計画進捗の確定・見込・参考は、計画に対する積み上げ（月別）と同じ数（ページどうしで食い違わない）
  const inMonths = (p, keys) => p.months.reduce((t, m) => t + keys.reduce((u, k) => u + (m.values[k] || 0), 0), 0);
  const ps = moved.planSales;
  const pm = moved.planSalesByMonth;
  assert.deepStrictEqual([ps.byRank.A + ps.byRank.B, ps.byRank.C, ps.total, ps.undated.value, ps.before.value, ps.excluded.value],
    [inMonths(pm, ['A', 'B']), inMonths(pm, ['C']), pm.total, pm.undated.value, pm.before.value, pm.excluded.value]);
  assert.strictEqual(slide.elements.find(e => e.role === 'lead').text.split('\n').length, 2, 'リード文は2行');
  // 区切って作るときの中身の一部にも入れる
  assert.deepStrictEqual(plain(g.mtgCombinedLite_(s1)).planSales.byRank.A, 70);
  // 並びはシートの順（DIFF_RULES の順。進捗率で並べ替えない）。想定売上のページ（amountPages）も同じ並び
  const names = sl => sl.elements.filter(e => e.role === 'compact-name').map(e => e.text);
  assert.deepStrictEqual(names(amountPage), names(slide));
  assert.match(leadOf(slide), /サブインダストリーはシートの順。/);
  assert.match(leadOf(amountPage), /サブインダストリーはシートの順。/);
  const big = Object.assign({}, s2, {planSales: Object.assign({}, s2.planSales, {byRank: Object.assign({}, s2.planSales.byRank, {A: 500})})});
  assert.deepStrictEqual(names(plain(g.mtgCompactOverviewSlide_([s1, big], {sales: true}))), ['新FMT', '新FMT2', '合計'],
    '売上の進捗率が高くても、並びはシートの順');
  assert.deepStrictEqual(names(plain(g.mtgCompactOverviewSlide_([s1, big]))), ['新FMT', '新FMT2', '合計']);
  // 見込（C）を含む売上のページ（次のページ）：C を点線で足す。並び・右の数字は同じ
  const withC = slides[2];
  assert.strictEqual(withC.title, '全サブインダストリーの状況・計画進捗（' + fy + '売上・見込（C）を含む）');
  assert.deepStrictEqual(names(withC), names(slide));
  assert.deepStrictEqual(withC.elements.filter(e => e.role === 'compact-value').map(e => e.text), values);
  assert.match(leadOf(withC), new RegExp('^確定（A＋B）に見込（C）を足すと 120百万円で、計画の 40%（進捗率（A＋B）は 40%）。'));
  // 出さない設定（compact.salesPage: false）
  gas.get('MTG_OPTIONS').compact.salesPage = false;
  assert.ok(!plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true)).some(sl => /売上）$/.test(sl.title)));
  assert.strictEqual(g.mtgCompactFrontCount_(), 5);
});

test('状況・計画進捗（FY売上）の前回比：確定（A＋B）の今年度の売上の増減（前回は前回のステータスで A・B の案件）。前回の値は、変更の記録（売上の列の変更前）から。失注になった案件は前回の分だけ減る。売上合計は3年度分', () => {
  const {gas, g, sheet} = setup();
  const fy = fiscalYear();
  const next = String(Number(fy) + 1);
  // 来年度の列（FY27前半期）を足す（3年度分の売上合計に入る。前回比には入らない）
  sheet.getRange('R2').setValue('FY' + next + '前半期売上（百万）');
  sheet.getRange('R3').setValue(100);   // A社
  g.recordDiffEdit(gas.edit(sheet, 'P3', 50));   // A社：今年度の前半期 30 → 50（+20）
  g.recordDiffEdit(gas.edit(sheet, 'Q5', 15));   // C社：今年度の後半期 空 → 15（+15）
  g.recordDiffEdit(gas.edit(sheet, 'F4', '失注'));   // B社：失注（前回の 30 の分だけ減る）
  g.recordDiffEdit(gas.edit(sheet, 'P7', 80));   // E社（失注のまま）：入れない
  const s1 = summarize(g, '新FMT');
  // 確定（A＋B）の分：前回 A社 70・B社 30 = 100（C社は前回 C）→ 今 A社 90・C社 25 = 115（B社は失注）
  assert.deepStrictEqual(plain(s1.salesBridge), {year: fy, years: [fy, next, String(Number(fy) + 2)], start: 110, end: 115, total: 215,
    sureStart: 100, sureEnd: 115});
  assert.strictEqual(s1.salesBridge.sureEnd, s1.planSales.byRank.A + s1.planSales.byRank.B, '今の分は棒の確定（A＋B）と同じ');
  // 売上の列の変更は、資料の「変化」には出さない（ignoreHeaders）。B社（失注）・C社・D社は前からの変化
  assert.ok(!s1.changes.some(c => c.customer === 'A社'));
  const values = plain(g.mtgCompactOverviewSlide_([s1], {sales: true})).elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual([values[3], values[5]], ['+15', '215'], '前回比は確定（A＋B）の増減（+15）、売上合計は3年度分');
  // 区切って作るときの中身の一部にも入れる
  assert.deepStrictEqual(plain(g.mtgCombinedLite_(s1)).salesBridge.start, 110);
  // 想定売上のページは、全期間の想定売上合計と、その前回比のまま
  const amount = plain(g.mtgCompactOverviewSlide_([s1])).elements.filter(e => e.role === 'compact-value').map(e => e.text);
  assert.deepStrictEqual(amount.slice(4, 6), ['160', '▲50']);
});

test('状況・計画進捗（FY売上）の前回比：前回の時点コピーと比べるときは、コピーの売上の列の値と比べる（コピーに無い案件は 0）', () => {
  const {gas, g, sheet} = setup();
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  // 前回の時点コピー：今のシートの写し（売上を書き換える前）
  const copySheet = gas.addSheet('前回コピー', sheet.getRange(1, 1, sheet.getLastRow(), WIDTH).getValues(), {rows: 30, columns: WIDTH});
  const copy = {table: g.mtgReadDeals_(copySheet, gas.get('DIFF_RULES')['新FMT']), title: '前回', url: ''};
  sheet.getRange('P3').setValue(60);   // A社：30 → 60（変更の記録が無くても、コピーと比べる）
  // H社：前回の後に足した案件（コピーに無い。前回は 0）。今年度の前半期（P列）7
  sheet.getRange('D8:Q8').setValues([['H社', '新規', 'C-提案中(50%)', 10, 5, fiscalMonth(9), '', '', '', '', '', '', 7, '']]);
  const now = new Date();
  const since = new Date(Date.now() - 14 * 86400000);
  const windowStart = new Date(Date.now() - 180 * 86400000);
  const s1 = g.mtgSummarize_(ss, '新FMT', {since, previous: {date: '2026/10/30', cutoff: since.getTime()}, legacy: null, baseline: null, draft: null,
    final: true, now, meeting: '2026/11/13', windowStart, copy, copyError: '', trackedSince: g.mtgLogStart_(ss),
    log: g.mtgReadLog_(ss, windowStart), snapshots: g.mtgReadSnapshots_(ss, windowStart), credentials: g.hsCredentialIndex_(ss)});
  assert.deepStrictEqual([s1.salesBridge.start, s1.salesBridge.end], [110, 147]);
  // 確定（A＋B）：コピーのステータスで A・B の案件（A社 70・B社 30・C社 10）→ 今（A社 100・B社 30・C社 10。H社は C）
  assert.deepStrictEqual([s1.salesBridge.sureStart, s1.salesBridge.sureEnd], [110, 140]);
});

test('diagnoseBusinessMtgTotals：今年度（P・Q列）と3年度分の売上を、列の合計 − 失注の行 ＝ 資料で突き合わせ、確度別・数値でない値・案件として数えない行の売上も出す', () => {
  const {gas, g, sheet} = setup();
  const fy = fiscalYear();
  sheet.getRange('Q4').setValue('約5');    // B社：後半期が文字（足さない）
  sheet.getRange('P10').setValue(9);       // 得意先などが空の行の売上（案件として数えないので足さない）
  const s1 = summarize(g, '新FMT');
  gas.promptAnswers = ['1'];
  g.diagnoseBusinessMtgTotals();
  const text = gas.alerts[gas.alerts.length - 1].message;
  // 列の合計：A社 70・B社 20・C社 10・E社（失注）50 = 150 − 50 = 100（資料の確定（A＋B）＋見込（C）＋そのほか）
  assert.ok(text.indexOf('　FY' + fy + 'の売上（P・Q列。状況・計画進捗の棒・確定・見込）：列の合計 150 − 失注の行 50 ＝ 資料 100') >= 0, text);
  assert.ok(text.indexOf('　FY' + fy + '〜' + (Number(fy) + 2) + 'の売上（P・Q列。状況・計画進捗の「売上合計」）：列の合計 150 − 失注の行 50 ＝ 資料 100') >= 0);
  assert.ok(text.indexOf('　　うち確定（A＋B）100・見込（C）0・そのほか（状況・計画進捗には出ない）0') >= 0);
  // 読んでいる列の見出しと、確定（A＋B）に入れたステータスごとの金額・行
  assert.ok(text.indexOf('　FY' + fy + 'の売上として読む列：P列「FY' + fy + '前半期売上（百万）」・Q列「FY' + fy + '後半期売上（百万）」') >= 0, text);
  assert.ok(text.indexOf('　　　A-内示(90%)：70（1件：3行目 70）') >= 0);
  assert.ok(text.indexOf('　　　B-クロージング中(70%)：30（2件：4行目 20・5行目 10）') >= 0);
  assert.strictEqual(s1.planSales.byRank.A + s1.planSales.byRank.B, 100, '資料の確定（A＋B）と同じ');
  assert.strictEqual(s1.salesBridge.total, 100, '資料の売上合計と同じ');
  assert.ok(text.indexOf('　　数値でない値（文字・エラー。足していません）：Q4「約5」') >= 0);
  assert.ok(text.indexOf('　　案件として数えない行（得意先・案件名・ステータス・想定売上・期待値が空）にある売上（足していません）：10行目 9') >= 0);
});

test('状況・計画進捗（FY売上）：今年度の売上の列が無いシートは、棒の場所に「データなし」と理由', () => {
  const {g} = setup();
  const s1 = summarize(g, '新FMT');
  const lacking = Object.assign({}, s1, {planSales: g.mtgSalesPlanOf_([], '新FMT', new Date(), {salesYears: []})});
  const slide = plain(g.mtgCompactOverviewSlide_([lacking], {sales: true}));
  assert.ok(slide.elements.some(e => /^データなし（シートに「FY\d+/.test(e.text || "")));
});

test('計画に対する積み上げ（FY売上・salesPlanBy: half）：今年度の売上の列を前半期 → 後半期と積み上げ、合計・計画・不足。確定（A＋B）だけのページも。失注は参考', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').compact.outlookPage = false;   // 前の FY26売上のページ（FY26〜28 の版に差し替える前）
  gas.get('MTG_OPTIONS').compact.salesPlanBy = 'half';
  const fy = 'FY' + fiscalYear();
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  // 半期ごと：新FMT 前半期 A社 30・B社 20・C社 10 = 60、後半期 A社 40・B社 10 = 50。新FMT2 F社 5・5。E社（失注）の 50 は参考
  assert.deepStrictEqual(plain(s1.planSales.halves).map(h => h.total), [60, 50]);
  const merged = plain(g.mtgMergeSalesPlans_([s1.planSales, s2.planSales]));
  assert.deepStrictEqual([merged.halves.map(h => h.total), merged.total, merged.plan, merged.excluded.value], [[65, 55], 120, 300, 50]);
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const all = slides.find(sl => sl.title === gas.get('MTG_TITLES').plan);
  const sure = slides.find(sl => sl.title === gas.get('MTG_TITLES').plan + '：確定（A＋B）');
  assert.ok(all && sure && slides.indexOf(sure) === slides.indexOf(all) + 1 && slides.indexOf(all) === 3, '状況・計画進捗の2枚の次');
  assert.deepStrictEqual(all.elements.filter(e => e.role === 'plan-label').map(e => e.text), ['前半期', '後半期', '合計', '計画', '不足', '失注\n1件']);
  assert.deepStrictEqual(all.elements.filter(e => e.role === 'plan-value').map(e => e.text), ['+65', '+55', '120', '300', '▲180', '50']);
  assert.match(leadOf(all), new RegExp('^' + fy + ' の売上（前半期＋後半期の列。失注除く）は 120百万円。計画 300百万円に対して ▲180（達成率 40%）。'));
  // 確定（A＋B）：A 70・B（B社 30・C社 10・F社 10）= 120（新FMT の C は売上なし）。凡例も A・B
  assert.deepStrictEqual(sure.elements.filter(e => e.role === 'legend' && e.text).map(e => e.text), ['A', 'B', '計画']);
  assert.match(leadOf(sure), new RegExp('^確定（A＋B）の' + fy + ' の売上（前半期＋後半期の列。失注除く）は 120百万円。.*（進捗率 40%）'));
  // 出さない設定（compact.salesPage: false）なら、この2枚も入れない
  gas.get('MTG_OPTIONS').compact.salesPage = false;
  assert.ok(!plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true)).some(sl => sl.title.indexOf(fy + '売上') >= 0));
});

test('計画に対する積み上げ（FY売上・salesPlanBy: orderMonth。既定）：案件の今年度の売上の合計を受注月の月に置き、想定売上の積み上げと同じ月別の図にする', () => {
  const {gas, g} = setup();
  gas.get('MTG_OPTIONS').compact.outlookPage = false;   // 前の FY26売上のページ（FY26〜28 の版に差し替える前）
  const T = gas.get('MTG_TITLES');
  const fy = 'FY' + fiscalYear();
  const s1 = summarize(g, '新FMT');
  const s2 = summarize(g, '新FMT2');
  // 月（テストでは売上開始月）ごと：8か月目 A社 70 + F社 10 = 80、9か月目 B社 30、10か月目 C社 10。E社（失注）の 50 は参考の「失注」
  const merged = plain(g.mtgMergePlans_([s1.planSalesByMonth, s2.planSalesByMonth]));
  assert.deepStrictEqual(merged.months.map(m => m.total).slice(7, 11), [0, 80, 30, 10]);
  assert.deepStrictEqual([merged.total, merged.excluded.value, merged.plan], [120, 50, 300]);
  const slides = plain(g.mtgCompactSlides_([s1, s2], '2026/11/13', true));
  const all = slides.find(sl => sl.title === T.plan);
  const sure = slides.find(sl => sl.title === T.plan + '：確定（A＋B）');
  assert.ok(all && sure && slides.indexOf(all) === 3 && slides.indexOf(sure) === 4);
  assert.ok(all.elements.some(e => /月別 FY\d+売上の積み上げ（百万円）$/.test(e.text || '')), '月別の図');
  assert.match(leadOf(all), new RegExp('^今年度（[^）]+）の' + fy + '売上の積み上げは 120百万円。計画 300百万円に対して ▲180（達成率 40%）。'));
  assert.match(leadOf(sure), new RegExp('^今年度（[^）]+）の確定（A＋B）の' + fy + '売上の積み上げは 120百万円。.*（進捗率 40%）'));
  assert.ok(!all.elements.some(e => /空・数値でない/.test(e.text || '')), '売上が空の案件の件数は書かない');
  // 今年度の売上が無い案件（D社・G社）は数えない（積み上げに含まない案件は、失注の E社 1件だけ）
  assert.strictEqual(merged.undated.count + merged.before.count + merged.after.count + merged.excluded.count, 1);
});
