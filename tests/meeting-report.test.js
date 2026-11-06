'use strict';

const {test, assert} = require('./lib/harness');
const {createGas: createMockGas, plain} = require('./lib/gas-mock');
const {withRunDialog} = require('./lib/meeting-run');
const createGas = options => withRunDialog(createMockGas(options));

/*
 * 会議サマリの報告案件版（MeetingReport.gs）
 * 追跡シート（2行目が見出し。追跡範囲は D3:AF）：D 得意先 / E 案件名 / F 案件ステータス / G 想定売上 / H 期待値 / I 受注月 /
 *   J 活動状況 / K 活動における課題 / L アカウント責任者 / M 先方部門 / N 先方担当役職 / O 次回アクション / P RXシナリオ構成 /
 *   Q サービスパッケージ合算 / R 商談記録 / S 戦略 / T 価値創出 / AG 最終更新日時 / AJ 報告フラグ（追跡範囲の外）
 */
const WIDTH = 36;
const COL = {customer: 4, project: 5, status: 6, amount: 7, weighted: 8, orderMonth: 9, activity: 10, issue: 11, owner: 12,
  clientDept: 13, clientTitle: 14, nextAction: 15, scenario: 16, scenarioAmount: 17, record: 18, strategy: 19, valueCase: 20,
  stamp: 33, reportFlag: 36};
const HEADERS = {customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', orderMonth: '受注月', activity: '活動状況', issue: '活動における課題',
  owner: 'アカウント責任者', clientDept: '先方部門', clientTitle: '先方担当役職', nextAction: '次回アクション', scenario: 'RXシナリオ構成',
  scenarioAmount: 'サービスパッケージ合算（百万円）', record: '商談記録（ＵＲＬ）', strategy: '提案を勝ち取るための戦略・差異化要素',
  valueCase: '想定される価値創出のケース／期待される導入効果', stamp: '最終更新日時', reportFlag: '報告フラグ'};
const STATUSES = ['創注-仮検討(10%)', 'D-ヒアリング中(20%)', 'C-提案中(50%)', 'B-クロージング中(70%)', 'A-受注済み(100%)', '失注'];

function row(values) {
  const line = new Array(WIDTH).fill('');
  Object.keys(values).forEach(key => { line[COL[key] - 1] = values[key]; });
  return line;
}

const DEALS = [
  {customer: 'A社', project: '店舗DX基盤の刷新', status: 'C-提案中(50%)', amount: 100, weighted: 50, orderMonth: '2026/12', activity: '11/2 役員向けに最終提案',
    issue: '情報システム部門の稼働が逼迫', owner: '佐藤', clientDept: '経営企画部', clientTitle: '部長', nextAction: '11/15 稟議資料の事前確認',
    scenario: '店舗オペレーション効率化, 需要予測', scenarioAmount: 80, record: 'https://drive.google.com/drive/folders/abc',
    strategy: '業務設計から運用定着まで一括で提案する', valueCase: '', reportFlag: true},
  {customer: 'B社', project: '需要予測', status: 'D-ヒアリング中(20%)', amount: 30, weighted: 6, orderMonth: '2027/03', owner: '鈴木', reportFlag: false},
  {customer: 'C社', project: '販促最適化', status: 'C-提案中(50%)', amount: 40, weighted: 20, orderMonth: '2027/01', owner: '田中', reportFlag: true}
];

function setup(options) {
  const o = options || {};
  const gas = createGas();
  const g = gas.global;
  const headers = Object.assign({}, HEADERS);
  if (o.noFlagColumn) delete headers.reportFlag;
  const deals = o.deals || DEALS;
  const lines = [row({customer: '新FMT'}), row(headers)].concat(deals.map(d => row(d)));
  const sheet = gas.addSheet('新FMT', lines, {rows: 20, columns: WIDTH});
  sheet.getRange('F3:F20').setDataValidation(g.SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  gas.addSheet('新FMT2', [row({customer: '新FMT2'}), row(HEADERS)], {rows: 10, columns: WIDTH});
  g.setupDiffTracking();
  if (o.edits !== false) {
    g.recordDiffEdit(gas.edit(sheet, 'F3', 'B-クロージング中(70%)'));   // A社：C → B
    g.recordDiffEdit(gas.edit(sheet, 'H3', 70));                       // A社：期待値 50 → 70
    g.recordDiffEdit(gas.edit(sheet, 'I3', '2027/02'));                // A社：受注月 2026/12 → 2027/02（延期）
  }
  gas.writes.length = 0;
  return {gas, g, sheet};
}

function summarize(g, name) {
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  const windowStart = new Date(Date.now() - 180 * 86400000);
  return g.mtgSummarize_(ss, name, {
    since: new Date(Date.now() - 14 * 86400000), previous: null, legacy: null, baseline: null, draft: null, final: true, now: new Date(),
    meeting: '2026/10/23', windowStart, copy: null, copyError: '', trackedSince: g.mtgLogStart_(ss),
    log: g.mtgReadLog_(ss, windowStart), snapshots: g.mtgReadSnapshots_(ss, windowStart), credentials: g.hsCredentialIndex_(ss)
  });
}

const texts = slide => slide.elements.filter(el => el.text != null).map(el => String(el.text));
const byRole = (slide, role) => slide.elements.filter(el => el.role === role);

test('報告案件：報告フラグにチェックのある案件だけを、前回の値（変化が無ければ今の値）と一緒に取り出す（シートの順）', () => {
  const {g} = setup();
  const s = summarize(g, '新FMT');
  assert.strictEqual(s.reportColumn, true);
  const items = plain(s.reported).map(d => [d.customer, d.status, d.prevStatus, d.weighted, d.prevWeighted, d.orderMonthText, d.prevMonthText, d.isNew]);
  assert.deepStrictEqual(items, [
    ['A社', 'B-クロージング中(70%)', 'C-提案中(50%)', 70, 50, '2027/02', '2026/12', false],
    ['C社', 'C-提案中(50%)', 'C-提案中(50%)', 20, 20, '2027/01', '2027/01', false]
  ]);
  assert.strictEqual(s.reported[0].strategy, '業務設計から運用定着まで一括で提案する');
  assert.strictEqual(s.reported[0].valueCase, '');
  assert.strictEqual(g.mtgReportFlag_('TRUE', 'TRUE'), true, '文字の TRUE もチェックとみなす');
  assert.strictEqual(g.mtgReportFlag_(false, 'FALSE'), false);
  assert.deepStrictEqual(['A-受注済み(100%)', 'Ｂ-クロージング中(70%)', '創注-仮検討(10%)', '失注', ''].map(g.mtgReportRank_), ['A', 'B', '創注', '', '']);
});

test('報告案件版のスライド：表紙と数字のページは会議サマリと同じで、案件のページと Appendix の代わりに報告案件を1件1枚', () => {
  const {gas, g} = setup();
  const s = summarize(g, '新FMT');
  const normal = plain(g.mtgSlides_(s));
  const slides = plain(g.mtgReportSlides_(s));
  const titles = slides.map(sl => sl.title);
  const numbers = normal.slice(0, normal.findIndex(sl => sl.title === gas.get('MTG_TITLES').progress)).map(sl => sl.title);
  assert.deepStrictEqual(titles.slice(0, numbers.length), numbers, '表紙から RXシナリオまでは同じ');
  assert.deepStrictEqual(titles.slice(numbers.length), ['報告案件：A社／店舗DX基盤の刷新', '報告案件：C社／販促最適化']);
  assert.ok(!slides.some(sl => sl.divider || /Appendix|前回から変化のある案件|期待値上位案件/.test(sl.title)), 'Appendix・案件のページは入れない');

  const a = slides[numbers.length];
  assert.ok(!byRole(a, 'title').length, '見出し（タイトル）は付けない');
  const head = byRole(a, 'report-title')[0];
  assert.strictEqual(head.text, '報告案件：A社／店舗DX基盤の刷新');
  assert.deepStrictEqual(head.spans.map(sp => [sp.start, sp.end]), [[0, 5]], '「報告案件：」だけ緑');
  const all = texts(a).join('\n');
  ['先方', '経営企画部', '部長', 'B-クロージング中(70%)', '前回：C-提案中(50%)', '70　+20', '前回：50　／　想定売上：100', '2027/02', '前回：2026/12（延期）',
    'RXシナリオ（見込金額 80百万円）', '店舗オペレーション効率化', '需要予測', '提案を勝ち取るための戦略・差異化要素', '業務設計から運用定着まで一括で提案する',
    '想定される価値創出のケース／期待される導入効果', '11/2 役員向けに最終提案', '情報システム部門の稼働が逼迫', '11/15 稟議資料の事前確認']
    .forEach(t => assert.ok(all.indexOf(t) >= 0, t + ' が無い：\n' + all));
  assert.ok(all.indexOf('佐藤') < 0, '責任者は出さない');
  const bodies = byRole(a, 'report-text');
  assert.strictEqual(bodies.length, 5);
  assert.strictEqual(new Set(bodies.map(b => b.size)).size, 1, '文章の枠の字の大きさはスライドの中で同じ');
  assert.strictEqual(bodies[0].size, 11, '短ければいちばん大きい大きさ');
  assert.strictEqual(bodies[1].text, '未入力', '空の項目は「未入力」');
  // 確度の段：今（B）は緑・前回（C）は薄い緑
  const steps = byRole(a, 'report-step');
  assert.deepStrictEqual(steps.map(st => st.text), ['創注', 'D', 'C', 'B', 'A']);
  assert.deepStrictEqual(steps.map(st => st.fill), ['#EDEDED', '#EDEDED', '#E2EFDA', '#70AD47', '#EDEDED']);
  // 右下の商談記録はリンク（出典の行は、その分だけ狭い）
  const record = byRole(a, 'footer-right')[0];
  assert.match(record.text, /^商談記録：https:\/\/drive\.google\.com\//, '長い URL は「…」で切って見せる（リンクは全部）');
  assert.deepStrictEqual(record.spans.map(sp => sp.url), ['https://drive.google.com/drive/folders/abc']);
  assert.ok(byRole(a, 'footer')[0].w < 676 - 60, '出典の行は商談記録の分だけ狭い');
  assert.match(byRole(a, 'footer')[0].text, /前回＝表紙の対象期間の始まり/);
  assert.ok(byRole(a, 'page').length === 1, 'ページ番号は付ける');

  const cSlide = slides[numbers.length + 1];
  assert.ok(texts(cSlide).join('\n').indexOf('前回：C-提案中(50%)') >= 0, '変化が無ければ前回も今と同じ');
  assert.strictEqual(byRole(cSlide, 'footer-right')[0].text, '商談記録：未入力');
  assert.ok(texts(cSlide).some(t => /^20　±0\n/.test(t)), '期待値が同じなら ±0');
});

test('報告案件版のスライド：文章が多いと、どの枠にも入る大きさまで全部の枠を小さくし、最小でも入らなければ「…」で切る', () => {
  const long = 'あいうえおかきくけこ'.repeat(30);
  const huge = 'あいうえおかきくけこ'.repeat(200);
  const {g} = setup({edits: false, deals: [Object.assign({}, DEALS[0], {strategy: long}), Object.assign({}, DEALS[2], {activity: huge, reportFlag: true})]});
  const s = summarize(g, '新FMT');
  const slides = plain(g.mtgReportSlides_(s)).filter(sl => /^報告案件：/.test(sl.title));
  const sizeOf = sl => byRole(sl, 'report-text')[0].size;
  assert.ok(sizeOf(slides[0]) < 11 && sizeOf(slides[0]) >= 7, '長い戦略の分だけ小さく：' + sizeOf(slides[0]));
  assert.strictEqual(new Set(byRole(slides[0], 'report-text').map(b => b.size)).size, 1);
  assert.ok(!slides[0].cut);
  assert.ok(byRole(slides[0], 'report-text')[0].text === long, '入れば全文');
  assert.strictEqual(sizeOf(slides[1]), 7, '入らなければ最小の大きさ');
  assert.strictEqual(slides[1].cut, true);
  const activity = byRole(slides[1], 'report-text')[2];
  assert.match(activity.text, /…$/, '最小でも入らなければ「…」で切る');
  const box = activity;
  assert.ok(g.mtgTextHeight_(activity.text.split('\n').length, 7) <= box.h + 0.01, '箱の高さに収める');
});

test('報告案件版のスライド：チェックのある案件が無ければ「該当なし」、報告フラグの列が無ければ「データなし」の1枚', () => {
  const none = setup({edits: false, deals: DEALS.map(d => Object.assign({}, d, {reportFlag: false}))});
  const s1 = summarize(none.g, '新FMT');
  const last1 = plain(none.g.mtgReportSlides_(s1)).pop();
  assert.strictEqual(last1.empty, 'none');
  assert.ok(texts(last1).some(t => /報告フラグにチェックの入った案件がありません/.test(t)));

  const gap = setup({edits: false, noFlagColumn: true});
  const s2 = summarize(gap.g, '新FMT');
  assert.strictEqual(s2.reportColumn, false);
  assert.deepStrictEqual(plain(s2.reported), []);
  const last2 = plain(gap.g.mtgReportSlides_(s2)).pop();
  assert.strictEqual(last2.empty, 'gap');
  assert.ok(texts(last2).some(t => /シートに「報告フラグ」の列が無い/.test(t)));
});

test('報告案件版を作る関数：前の関数はそのままで、ファイル名に「_報告案件版」が付く（試作は「_報告案件版_試作」）', () => {
  const {gas, g} = setup();
  assert.strictEqual(g.mtgFileName_('新FMT', '2026/10/23', true, 'report'), '20261023_新FMT_RX Business Meeting用資料_報告案件版');
  assert.strictEqual(g.mtgFileName_('新FMT', '2026/10/23', false, 'report'), '20261023_新FMT_RX Business Meeting用資料_報告案件版_試作');
  assert.strictEqual(g.mtgFileName_('新FMT', '2026/10/23', false), '20261023_新FMT_RX Business Meeting用資料_試作');

  gas.promptAnswers = ['1', '10/23'];
  g.previewBusinessMtgReportSummaries();
  const meeting = g.mtgParseDate_('10/23', new Date()).replace(/\//g, '');
  assert.ok(gas.driveFiles.some(f => f.name === meeting + '_新FMT_RX Business Meeting用資料_報告案件版_試作.pdf'),
    gas.driveFiles.map(f => f.name).join('\n'));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /_報告案件版_試作/);

  gas.promptAnswers = ['1', '10/23'];
  g.previewBusinessMtgSummaries();
  assert.ok(gas.driveFiles.some(f => f.name === meeting + '_新FMT_RX Business Meeting用資料_試作.pdf'), '前の関数は前のファイル名');

  // 本番の確かめの文に、報告案件版の説明が付く
  const check = g.checkMeetingRun('report', true, '10/23', ['新FMT'], '');
  assert.match(check.text, /^【報告案件版】/);
  assert.doesNotMatch(g.checkMeetingRun('business', true, '10/23', ['新FMT'], '').text, /報告案件版/);
  assert.match(g.getMeetingRunData('report', true).title, /報告案件版/);
});
