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
  clientDept: 13, clientTitle: 14, nextAction: 15, scenario: 16, scenarioAmount: 17, record: 18, strategy: 19, valueCase: 20, bx: 21,
  stamp: 33, reportFlag: 36};
const HEADERS = {customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', orderMonth: '受注月', activity: '活動状況', issue: '活動における課題',
  owner: 'アカウント責任者', clientDept: '先方部門', clientTitle: '先方担当役職', nextAction: '次回アクション', scenario: 'RXシナリオ構成',
  scenarioAmount: 'サービスパッケージ合算（百万円）', record: '商談記録（ＵＲＬ）', strategy: '提案を勝ち取るための戦略・差異化要素',
  valueCase: '想定される価値創出のケース／期待される導入効果', bx: 'BX担当', stamp: '最終更新日時', reportFlag: '報告フラグ'};
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
    strategy: '業務設計から運用定着まで一括で提案する', valueCase: '', bx: '山田', reportFlag: true},
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
  // 報告案件版の資料を作るときと同じく、前回の会議の宿題事項も読む（mtgJobRun_）
  return Object.assign(summarizeOnly(g, ss, name, windowStart), {homework: g.mtgHomeworkOf_(ss, name, '2026/10/23')});
}

function summarizeOnly(g, ss, name, windowStart) {
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

test('報告案件版のスライド：表紙と数字のページは会議サマリと同じで、そのあとに報告案件を1件1枚、最後に前回から変化のある案件（期待値上位案件・Appendix は入れない）', () => {
  const {gas, g} = setup();
  const s = summarize(g, '新FMT');
  const normal = plain(g.mtgSlides_(s));
  const slides = plain(g.mtgReportSlides_(s));
  const titles = slides.map(sl => sl.title);
  const numbers = normal.slice(0, normal.findIndex(sl => sl.title === gas.get('MTG_TITLES').progress)).map(sl => sl.title);
  // 表紙のあと・サマリーの前に「前回の宿題事項」（宿題事項のシートが無いので「データなし」の1枚）
  assert.deepStrictEqual(titles.slice(0, numbers.length + 1), [numbers[0], '前回の宿題事項'].concat(numbers.slice(1)), '表紙・宿題・数字のページ');
  assert.deepStrictEqual(titles.slice(numbers.length + 1), ['RXシナリオ進捗における報告案件①：A社／店舗DX基盤の刷新', 'RXシナリオ進捗における報告案件②：C社／販促最適化',
    '前回から変化のある案件'], '報告案件のあとに、前回から変化のある案件');
  assert.ok(!slides.some(sl => sl.divider || /Appendix|期待値上位案件/.test(sl.title)), 'Appendix・期待値上位案件のページは入れない');
  // サマリーは「現在の状態（進行中案件）」の表を出さない（前の版のサマリーには出す）
  const summaryOf = list => list.find(sl => sl.title === gas.get('MTG_TITLES').summary);
  assert.ok(!texts(summaryOf(slides)).some(t => /現在の状態/.test(t)), '報告案件版のサマリーに「現在の状態」は無い');
  assert.strictEqual(summaryOf(slides).elements.filter(el => el.type === 'table').length, 1, '表は「前回以降の動き」だけ');
  assert.ok(texts(summaryOf(normal)).some(t => /現在の状態/.test(t)), '前の版のサマリーには残す');
  // 表の金額の列は期待値（A社 70・C社 20）。4つ目の数字はシナリオ見込額（失注除く。A社 80）
  const table = summaryOf(slides).elements.find(el => el.type === 'table');
  assert.strictEqual(table.rows[0].cells[2].text, '期待値');
  assert.strictEqual(table.rows[0].cells[3].text, '主な案件（期待値順。（）は変わったこと）');
  const rowOf = label => table.rows.find(r => r.cells[0].text.indexOf(label) === 0);
  assert.strictEqual(rowOf('B').cells[2].text, '70', 'B の行（A社）は期待値 70（想定売上は 100）');
  assert.ok(texts(summaryOf(slides)).some(t => t === 'シナリオ見込額（失注除く・百万円）'));
  assert.ok(texts(summaryOf(slides)).some(t => /^80　/.test(t)), 'シナリオ見込額 80');
  assert.ok(!texts(summaryOf(slides)).some(t => /入力内容の確認事項/.test(t)));
  const oldTable = summaryOf(normal).elements.find(el => el.type === 'table');
  assert.strictEqual(oldTable.rows[0].cells[2].text, '想定売上', '前の版は想定売上のまま');
  assert.ok(texts(summaryOf(normal)).some(t => /入力内容の確認事項/.test(t)), '前の版は確認事項のまま');

  const a = slides[numbers.length + 1];
  assert.ok(!byRole(a, 'title').length, '見出し（タイトル）は付けない');
  const head = byRole(a, 'report-title')[0];
  assert.strictEqual(head.text, 'RXシナリオ進捗における報告案件①', '見出しは番号だけ（案件名・得意先名は左上の案件・得意先情報の枠）');
  assert.strictEqual(head.size, 21, '見出しはほかのスライドと同じ大きさ（長くなっても1行に入る）');
  assert.strictEqual(gas.get('MTG_OPTIONS').report.title, 'RXシナリオ進捗における報告案件');
  assert.ok(!(head.spans || []).length && head.color === '#1A1A1A', '色を付けない');

  assert.ok(!byRole(a, 'report-rule').length, '見出しの下の線は無い');
  assert.strictEqual(g.mtgReportTitle_('X', 21), 'RXシナリオ進捗における報告案件(21)：X');
  const all = texts(a).join('\n');
  ['案件・得意先情報', '店舗DX基盤の刷新\n \nA社\n経営企画部／部長', 'B-クロージング中(70%)', '前回：C-提案中(50%)', '前回：50　／　想定売上：100', '2027/02', '前回：2026/12（延期）',
    'RXシナリオ', '見込金額：80百万円', '店舗オペレーション効率化', '需要予測', '提案を勝ち取るための戦略・差異化要素', '業務設計から運用定着まで一括で提案する',
    '想定される価値創出のケース／期待される導入効果', '11/2 役員向けに最終提案', '情報システム部門の稼働が逼迫', '11/15 稟議資料の事前確認']
    .forEach(t => assert.ok(all.indexOf(t) >= 0, t + ' が無い：\n' + all));
  const staff = byRole(a, 'report-staff')[0];
  assert.strictEqual(staff.text, '担当営業：佐藤\nBX担当：山田', '右上に担当営業と BX担当の2行');
  assert.ok(staff.size * 2 === head.size && staff.y === head.y && staff.h === head.h && staff.align === 'END', '2行の高さが見出しの行と同じ（右寄せ）');
  assert.ok(head.x + head.w <= staff.x + 0.01, '見出しは担当の左まで');
  // 枠の中の字は、枠の見出し（8pt）より大きくしない（左の列の値・前回の帯・確度の段・シナリオの枠も）
  const heading = Math.max.apply(null, byRole(a, 'report-frame').map(el => el.size));
  ['report-value', 'report-text', 'report-prev', 'report-step', 'report-chip'].forEach(role => byRole(a, role).forEach(el => {
    [el.size].concat((el.spans || []).map(sp => sp.size).filter(Boolean)).forEach(size => assert.ok(size <= heading, role + '：' + size + 'pt'));
  }));
  const bodies = byRole(a, 'report-text');
  assert.strictEqual(bodies.length, 5);
  assert.strictEqual(new Set(bodies.map(b => b.size)).size, 1, '文章の枠の字の大きさはスライドの中で同じ');
  assert.strictEqual(bodies[0].size, gas.get('MTG_OPTIONS').report.maxSize, '短ければいちばん大きい大きさ（枠の見出しと同じ 8pt）');
  assert.strictEqual(bodies[1].text, '未入力', '空の項目は「未入力」');
  // 確度の段：今（B）は緑・前回（C）は薄い緑
  const steps = byRole(a, 'report-step');
  assert.deepStrictEqual(steps.map(st => st.text), ['創注', 'D', 'C', 'B', 'A']);
  assert.deepStrictEqual(steps.map(st => st.fill), ['#EDEDED', '#EDEDED', '#E2EFDA', '#70AD47', '#EDEDED']);
  // 右下の商談記録はリンク（出典の行は、その分だけ狭い）
  const record = byRole(a, 'footer-right')[0];
  assert.strictEqual(record.text, '商談記録：格納先リンク', 'URL は見せずに「格納先リンク」をハイパーリンクに');
  assert.deepStrictEqual(record.spans.map(sp => [sp.url, record.text.slice(sp.start, sp.end)]), [['https://drive.google.com/drive/folders/abc', '格納先リンク']]);
  assert.ok(byRole(a, 'footer')[0].w < 676 - 60, '出典の行は商談記録の分だけ狭い');
  const foot = byRole(a, 'footer')[0];
  assert.match(foot.text, /^出典：「新FMT」シート 3行目（\d{4}\/\d{2}\/\d{2} \d{2}:\d{2} 時点）$/,
    'どのシートの何行目を、いつ読んだか：' + foot.text);
  assert.strictEqual(foot.size, 8);
  assert.match(byRole(slides[numbers.length + 2], 'footer')[0].text, /「新FMT」シート 5行目/, 'C社は5行目');
  assert.ok(byRole(a, 'page').length === 1, 'ページ番号は付ける');

  const cSlide = slides[numbers.length + 2];
  assert.ok(texts(cSlide).join('\n').indexOf('前回：C-提案中(50%)') >= 0, '変化が無ければ前回も今と同じ');
  assert.strictEqual(byRole(cSlide, 'footer-right')[0].text, '商談記録：未入力');
  // 期待値の増減は、数字の右の色付きの字（バッジにしない）
  const valueRow = sl => sl.elements.find(el => el.role === 'report-value' && /^\d+　/.test(el.text || ''));
  assert.strictEqual(valueRow(cSlide).text, '20　±0', '期待値が同じなら ±0');
  const up = valueRow(a);
  assert.strictEqual(up.text, '70　+20');
  assert.deepStrictEqual(up.spans.map(sp => [up.text.slice(sp.start, sp.end), sp.color]), [['70', '#1A1A1A'], ['　+20', '#0070C0']], '増えたら青の字（減ったら赤）');
  assert.ok(!byRole(a, 'report-delta').length, 'バッジにしない');
  assert.ok(a.elements.every(el => el.shape !== 'ROUND_RECTANGLE'), '角の丸い四角は使わない');
});

test('報告案件のスライド：受注月の枠を左右に分け、右半分に売上（今年度から3年度分。年度ごとの前半期＋後半期）。列が無いシートは「データなし」', () => {
  const {gas, g} = setup();
  const s = summarize(g, '新FMT');
  const year = g.mtgFiscalYearKey_(new Date());
  const fy = k => ('0' + ((Number(year) + k) % 100)).slice(-2);
  const item = Object.assign({}, s.reported[0], {sales: {[fy(0)]: 30, [fy(1)]: 55}});
  const slide = plain(g.mtgReportSlide_(Object.assign({}, s, {salesYears: [fy(0), fy(1), fy(2)]}), item, 1));
  assert.deepStrictEqual(byRole(slide, 'report-sales-head').map(el => el.text), ['FY' + fy(0), 'FY' + fy(1), 'FY' + fy(2)]);
  assert.deepStrictEqual(byRole(slide, 'report-sales').map(el => el.text), ['30', '55', '−'], '空の年度は「−」');
  // 受注月と年度別売上は同じ段で、左右に並ぶ。右の枠は見出しが1行に入る幅。見出しの（）書きは小さく
  const frames = slide.elements.filter(el => el.text === '受注月' || el.text === '年度別売上（百万）');
  assert.strictEqual(frames.length, 2);
  assert.ok(frames[0].y === frames[1].y && frames[1].x > frames[0].x, '左に受注月・右に年度別売上');
  assert.ok(g.mtgReportHeadWidth_(frames[1].text) <= frames[1].w - 14.4, '見出しは1行');
  const R = gas.get('MTG_OPTIONS').report;
  assert.deepStrictEqual(plain(frames[1].spans), [{start: 5, end: 9, size: R.headNoteSize}]);
  const value = slide.elements.find(el => el.text === '期待値（百万円）');
  assert.deepStrictEqual(plain(value.spans), [{start: 3, end: 8, size: R.headNoteSize}], '期待値の（）書きも小さく');
  // 年度別の売上の列が無いシート
  const none = plain(g.mtgReportSlide_(Object.assign({}, s, {salesYears: []}), item, 1));
  assert.deepStrictEqual(byRole(none, 'report-sales').map(el => el.text), [gas.get('MTG_OPTIONS').gapText]);
});

test('報告案件版のスライド：報告案件のあとの「前回から変化のある案件」は、報告フラグのある案件を除く全件（Appendix が無いので7件までにしない・「Appendix ① を参照」と書かない）。前の版は今のまま', () => {
  const {gas, g} = setup();
  const s = summarize(g, '新FMT');
  const movesOf = list => plain(list).filter(sl => /^前回から変化のある案件/.test(sl.title));
  const rowsOf = list => [].concat.apply([], list.map(sl => sl.elements.find(el => el.type === 'table').rows.slice(1)));
  const leadOf = sl => byRole(sl, 'lead').map(el => el.text).join('\n');   // リード文（見出しの下。番号を付けたスライドでは部品の中）
  // 変化のあった案件は A社（ステータス・期待値・受注月）だけ。A社は報告案件（報告フラグあり）なので、こちらの表には出さない（2026/11/11）
  const one = movesOf(g.mtgReportSlides_(s));
  assert.strictEqual(one.length, 1);
  assert.strictEqual(leadOf(one[0]), '直近14日間、ステータス・期待値・受注月・活動状況に変化のある案件なし（報告案件 1件を除く）。');
  assert.ok(byRole(one[0], 'title').length === 1 && byRole(one[0], 'page').length === 1, 'ふつうの見出しとページ番号を付ける');
  // 報告フラグの無い案件は載せる
  const unflagged = movesOf(g.mtgReportSlides_(Object.assign({}, s, {reported: []})));
  assert.deepStrictEqual(rowsOf(unflagged).map(r => r.cells[0].text.replace(/\n/g, '／')), ['A社／店舗DX基盤の刷新']);
  assert.strictEqual(leadOf(unflagged[0]), '直近14日間に変化のある案件は 1件。');
  // 前の版（会議サマリ）は、報告フラグがあっても載せる
  assert.deepStrictEqual(rowsOf(movesOf(g.mtgSlides_(s))).map(r => r.cells[0].text.replace(/\n/g, '／')), ['A社／店舗DX基盤の刷新']);
  const slides = plain(g.mtgReportSlides_(s));
  assert.strictEqual(slides[slides.length - 1].title, '前回から変化のある案件', '資料の最後（報告案件のあと）');
  assert.match(slides[slides.length - 2].title, /報告案件②：/);

  // 変化のあった案件が多いとき（10件・その他の項目のみの更新 2件）：前の版は7件までで残りは Appendix ①、報告案件版は10件とも
  const many = Array.from({length: 10}, (_, i) => Object.assign({}, s.changes[0], {key: 'k' + i, project: '案件' + i}));
  const big = Object.assign({}, s, {changes: many, kpi: Object.assign({}, s.kpi, {updated: 2})});
  const report = movesOf(g.mtgReportSlides_(big));
  assert.deepStrictEqual(rowsOf(report).map(r => r.cells[0].text.split('\n').pop()), many.map(ch => ch.project), '10件とも（シートの値の順のまま）');
  assert.ok(report.length > 1, '入らなければ続きのページ');
  assert.deepStrictEqual(report.map(sl => sl.title).slice(0, 2), ['前回から変化のある案件', '前回から変化のある案件（続き）']);
  report.forEach(sl => {
    assert.strictEqual(leadOf(sl), '直近14日間に変化のある案件は 10件（ほかに、その他の項目のみの更新 2件）。', 'Appendix の案内・空の2行目は無い');
    const t = sl.elements.find(el => el.type === 'table');
    assert.ok(t.y + t.rows.reduce((h, r) => h + r.h, 0) <= gas.get('MTG_SLIDE').bottom + 0.01, 'ページに入る');
  });
  assert.ok(!plain(g.mtgReportSlides_(big)).some(sl => texts(sl).some(t => /Appendix/.test(t))), '報告案件版のどのページにも「Appendix」と書かない');
  const normal = movesOf(g.mtgSlides_(big));
  assert.strictEqual(rowsOf(normal).length, gas.get('MTG_OPTIONS').mainRows.changes, '前の版は7件まで');
  assert.match(leadOf(normal[0]), /に 7件を掲載。全件は Appendix ① を参照。$/);
  // ページ番号は表紙から通し（報告案件・変化のある案件にも続けて付く）
  const pages = plain(g.mtgReportSlides_(big)).map(sl => (byRole(sl, 'page')[0] || {}).text);
  assert.deepStrictEqual(pages.slice(1), pages.slice(1).map((_, i) => String(i + 2)));
});

test('報告案件版のスライド：文章の字は枠ごとに、入るいちばん大きい大きさ（見出しより大きくしない）。文の多い枠だけ小さくし、最小でも入らなくても省略しない（全文）', () => {
  const long = 'あいうえおかきくけこ'.repeat(60);
  const huge = 'あいうえおかきくけこ'.repeat(200);
  const {gas, g} = setup({edits: false, deals: [Object.assign({}, DEALS[0], {strategy: long}), Object.assign({}, DEALS[2], {activity: huge, reportFlag: true})]});
  const s = summarize(g, '新FMT');
  const slides = plain(g.mtgReportSlides_(s)).filter(sl => /報告案件[①-⑳]：/.test(sl.title));
  const sizeOf = sl => byRole(sl, 'report-text')[0].size;
  const R = gas.get('MTG_OPTIONS').report;
  assert.ok(sizeOf(slides[0]) < R.maxSize && sizeOf(slides[0]) >= R.minSize, '長い戦略の分だけ小さく：' + sizeOf(slides[0]));
  // ほかの枠（短い文）は見出しと同じ大きさのまま（文の多い枠に合わせて小さくしない。2026/11/11 から）
  assert.ok(byRole(slides[0], 'report-text').slice(1).every(b => b.size === R.maxSize), '短い文の枠は小さくしない');
  // 活動状況が入りきらない2枚目も、戦略の枠は見出しと同じ大きさ
  assert.strictEqual(sizeOf(slides[1]), R.maxSize);
  assert.ok(!slides[0].cut);
  assert.ok(byRole(slides[0], 'report-text')[0].text === long, '入れば全文');
  assert.strictEqual(slides[1].cut, true);
  const activity = byRole(slides[1], 'report-text')[3];   // 戦略・価値創出・課題・活動状況・次回アクションの順
  assert.strictEqual(activity.size, R.minSize, '入らなければ最小の大きさ');
  assert.strictEqual(activity.text, huge, '最小でも入らなくても省略しない');
  // 枠の見出しは、枠の線から離す（段落の上の余白）
  assert.ok(byRole(slides[0], 'report-frame').every(b => b.spaceAbove === R.headTop && R.headTop > 0));
});

test('報告案件版のスライド：チェックのある案件が無ければ「該当なし」、報告フラグの列が無ければ「データなし」の1枚', () => {
  const none = setup({edits: false, deals: DEALS.map(d => Object.assign({}, d, {reportFlag: false}))});
  const s1 = summarize(none.g, '新FMT');
  // 報告案件のページ（番号の無い見出し。そのあとは前回から変化のある案件）
  const emptyOf = (g, s) => plain(g.mtgReportSlides_(s)).find(sl => sl.title === 'RXシナリオ進捗における報告案件');
  const last1 = emptyOf(none.g, s1);
  assert.strictEqual(last1.empty, 'none');
  assert.strictEqual(byRole(last1, 'report-title')[0].text, 'RXシナリオ進捗における報告案件', '報告案件が無いときの見出しは番号なし');
  assert.ok(texts(last1).some(t => /報告フラグにチェックの入った案件がありません/.test(t)));

  const gap = setup({edits: false, noFlagColumn: true});
  const s2 = summarize(gap.g, '新FMT');
  assert.strictEqual(s2.reportColumn, false);
  assert.deepStrictEqual(plain(s2.reported), []);
  const last2 = emptyOf(gap.g, s2);
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
  assert.match(check.text, /そのあとに「前回から変化のある案件」（報告案件を除く全件）を載せます（期待値上位案件・Appendix は入れません）/);
  assert.doesNotMatch(g.checkMeetingRun('business', true, '10/23', ['新FMT'], '').text, /報告案件版/);
  assert.match(g.getMeetingRunData('report', true).title, /報告案件版/);
});

test('diagnoseBusinessMtgScenarioAmount：シナリオ見込額の前回比を、前回の時点コピーと案件ごとに突き合わせる（合計・違う行と値・列が数式か。得意先名は出さない）', () => {
  const {gas, g, sheet} = setup();
  gas.confirmAnswer = 'YES';
  gas.promptAnswers = ['1', '10/23'];
  g.createBusinessMtgSummaries();   // 本番：10/23 の会議日と時点コピーを記録する
  const ss = g.SpreadsheetApp.getActiveSpreadsheet();
  sheet.getRange('Q3').setValue(60);           // A社：80 → 60
  sheet.getRange('Q5').setValue('#N/A');       // C社：数値でない値
  const lines = g.mtgScenarioAmountLines_(ss, ['新FMT'], g.mtgParseDate_('10/30', new Date())).join('\n');
  assert.match(lines, /前回：10\/23 の時点コピー「時点コピー_/);
  assert.match(lines, /合計（失注除く）：前回 80 → 今 60（▲20）/);
  assert.match(lines, /前回 3行目 80 → 今 3行目 60/);
  assert.match(lines, /前回 5行目 空 → 今 5行目 「#N\/A」（数値でない）/);
  assert.match(lines, /列の中身：今のシート Q列は数式 0行・時点コピーは数式 0行/);
  assert.ok(lines.indexOf('A社') < 0 && lines.indexOf('C社') < 0, '得意先名は出さない');
  // 前回が時点コピーでなければ、そう書く
  const none = g.mtgScenarioAmountLines_(ss, ['新FMT'], g.mtgParseDate_('10/20', new Date())).join('\n');
  assert.match(none, /時点コピーと比べないので、シナリオ見込額の前回比は出ません/);
});

test('addBusinessMtgReportFlagColumn：各シートの AG列に「報告フラグ」の見出しとチェックボックスを足す（値は書かない）。2回目は付け直すだけ', () => {
  const gas = createGas();
  const g = gas.global;
  // 本物と同じく、最終更新日時は A列（AG列は空ける）
  const plainHeaders = Object.assign({}, HEADERS);
  delete plainHeaders.reportFlag;
  delete plainHeaders.stamp;
  const headerLine = () => Object.assign(row(plainHeaders), {0: HEADERS.stamp});
  const lines = [row({customer: '新FMT'}), headerLine(), row({customer: 'A社', project: 'X', status: 'C-提案中(50%)'})].map(l => l.slice(0, 32));
  // 新FMT：A〜AF の32列（本物と同じ）。新FMT2：AG列に値がある
  const sheet = gas.addSheet('新FMT', lines, {rows: 10, columns: 32});
  gas.addSheet('新FMT2', [row({customer: '新FMT2'}), headerLine()].concat([Object.assign(row({}), {32: 'メモ'})]), {rows: 10, columns: WIDTH});
  g.setupDiffTracking();
  gas.writes.length = 0;
  const result = g.addBusinessMtgReportFlagColumn();
  assert.deepStrictEqual(plain(result), [
    '「新FMT」：AG列に「報告フラグ」の列を足しました（3行目から下をチェックボックスに。列が増えたので、変更履歴の比較基準を取り直しました）。',
    '「新FMT2」：AG列に値があるため、足しませんでした（AG3：「メモ」）。MTG_OPTIONS.report.flagColumn を空いている列にしてください。'
  ]);
  assert.strictEqual(sheet.getMaxColumns(), 33);
  assert.strictEqual(sheet.getRange('AG2').getValue(), '報告フラグ');
  assert.strictEqual(sheet.getRange('AG3').getDataValidation().getCriteriaType(), 'CHECKBOX');
  assert.strictEqual(sheet.getRange('AG10').getDataValidation().getCriteriaType(), 'CHECKBOX');
  assert.strictEqual(sheet.getRange('AG3').getValue(), '', '値は書かない（最後の行を変えない）');
  assert.strictEqual(sheet.getLastRow(), 3);
  assert.strictEqual(gas.ss.getSheetByName('新FMT2').getRange('AG3').getValue(), 'メモ', '上書きしない');

  // 足した列を資料が読む（チェックした案件が報告案件になる）
  sheet.getRange('AG3').setValue(true);
  assert.deepStrictEqual(plain(summarize(g, '新FMT').reported).map(d => d.customer), ['A社']);

  // 2回目：もうある列は付け直すだけ
  assert.strictEqual(g.addBusinessMtgReportFlagColumn()[0], '「新FMT」：「報告フラグ」の列は AG列にあります（チェックボックスを付け直しました）。');
  assert.strictEqual(sheet.getMaxColumns(), 33);

  // いいえなら何もしない
  gas.confirmAnswer = 'NO';
  assert.strictEqual(g.addBusinessMtgReportFlagColumn(), undefined);
});

test('addBusinessMtgReportFlagColumn：AF列まで無いシート・追跡範囲の中の列には足さない', () => {
  const gas = createGas();
  const g = gas.global;
  // 本物と同じく、最終更新日時は A列（AG列は空ける）
  const plainHeaders = Object.assign({}, HEADERS);
  delete plainHeaders.reportFlag;
  delete plainHeaders.stamp;
  const headerLine = () => Object.assign(row(plainHeaders), {0: HEADERS.stamp});
  gas.addSheet('新FMT', [row({customer: '新FMT'}), headerLine()].map(l => l.slice(0, 27)), {rows: 10, columns: 27});
  gas.addSheet('新FMT2', [row({customer: '新FMT2'}), headerLine()], {rows: 10, columns: WIDTH});
  g.setupDiffTracking();
  gas.get('MTG_OPTIONS').report.flagColumn = 'AE';
  assert.match(g.addBusinessMtgReportFlagColumn()[1], /^「新FMT2」：AE列が追跡範囲（D3:AF）の中のため、足しませんでした/);
  gas.get('MTG_OPTIONS').report.flagColumn = 'AG';
  const result = g.addBusinessMtgReportFlagColumn();
  assert.strictEqual(result[0], '「新FMT」：AF列まで無いため、足しませんでした（今は AA列まで）。');
  assert.strictEqual(gas.ss.getSheetByName('新FMT').getMaxColumns(), 27);
  assert.deepStrictEqual([g.mtgColumnNumber_('AG'), g.mtgColumnLetter_(33), g.mtgColumnLetter_(27), g.mtgColumnLetter_(26)], [33, 'AG', 'AA', 'Z']);
});

/** 本物と同じ形（最終更新日時が A列・A〜AF の32列）の追跡シートを2つ作り、差分追跡を始める */
function realLayout() {
  const gas = createGas();
  const g = gas.global;
  const plainHeaders = Object.assign({}, HEADERS);
  delete plainHeaders.reportFlag;
  delete plainHeaders.stamp;
  const header = Object.assign(row(plainHeaders), {0: HEADERS.stamp}).slice(0, 32);
  const deal = row({customer: 'A社', project: 'X', status: 'C-提案中(50%)', activity: '提案'}).slice(0, 32);
  const sheet = gas.addSheet('新FMT', [row({}).slice(0, 32), header, deal], {rows: 10, columns: 32});
  gas.addSheet('新FMT2', [row({}).slice(0, 32), header], {rows: 10, columns: 32});
  g.setupDiffTracking();
  return {gas, g, sheet};
}

const logRows = gas => gas.ss.getSheetByName('変更履歴_差分').getLastRow();
const diffState = (gas, g) => JSON.parse(g.PropertiesService.getDocumentProperties().getProperty(gas.get('DIFF_OPTIONS').stateKey));

test('報告フラグの列を足しても差分追跡は止まらない（列を増やしたシートは比較基準を取り直す）。チェックは変更の記録・最終更新日時・行の内容に残さない', () => {
  const {gas, g, sheet} = realLayout();
  const lines = g.addBusinessMtgReportFlagColumn();
  assert.match(lines[0], /列が増えたので、変更履歴の比較基準を取り直しました/);
  assert.ok(Object.values(diffState(gas, g)).every(s => !s.paused && s.gridColumns === 33), '止まっていない');

  // 追跡範囲の編集は今までどおり記録される
  const before = logRows(gas);
  g.recordDiffEdit(gas.edit(sheet, 'J3', '11/6 再提案'));
  assert.strictEqual(logRows(gas), before + 1, '活動状況の変更を記録する');
  assert.ok(Object.values(diffState(gas, g)).every(s => !s.paused));
  const stamp = sheet.getRange('A3').getValue();

  // 報告フラグのチェックは記録しない・最終更新日時を変えない
  g.recordDiffEdit(gas.edit(sheet, 'AG3', true));
  g.onCredentialLauncherEdit(gas.edit(sheet, 'AG3', true));
  assert.strictEqual(logRows(gas), before + 1, 'チェックは変更の記録に残らない');
  assert.strictEqual(sheet.getRange('A3').getValue(), stamp, '最終更新日時は変わらない');
  assert.strictEqual(sheet.getRange('AG3').getValue(), true, 'ボタン列ではないので、チェックは戻さない');

  // チェックの入った行を編集しても、行の内容（スナップショット）に報告フラグは入れない
  g.recordDiffEdit(gas.edit(sheet, 'K3', '課題あり'));
  const snap = gas.ss.getSheetByName('変更時点スナップショット');
  const content = snap.getRange(snap.getLastRow(), 1, 1, snap.getLastColumn()).getDisplayValues()[0].join('\n');
  assert.match(content, /課題あり/);
  assert.doesNotMatch(content, /報告フラグ/);
});

test('報告フラグの列を足すとき、まだ記録していない編集（取りこぼし回収の印）があるシートには足さない', () => {
  const {gas, g, sheet} = realLayout();
  const props = g.PropertiesService.getDocumentProperties();
  props.setProperty(gas.get('DIFF_OPTIONS').dirtyPrefix + sheet.getSheetId(), 'x');
  const lines = g.addBusinessMtgReportFlagColumn();
  assert.strictEqual(lines[0], '「新FMT」：まだ変更の記録に書いていない編集があるため、足しませんでした（数分ごとの取りこぼし回収のあとで、もう一度実行してください）。');
  assert.strictEqual(sheet.getMaxColumns(), 32);
  assert.match(lines[1], /^「新FMT2」：AG列に「報告フラグ」の列を足しました/);
});

test('報告案件：活動状況・次回アクションの下に前回の内容。前回から変わっていなければ「今回と同じ（前回から更新なし）」と「更新なし」の印', () => {
  const {gas, g, sheet} = setup();
  g.recordDiffEdit(gas.edit(sheet, 'J3', '11/9 稟議の事前説明'));   // A社：活動状況を書き換える（次回アクションはそのまま）
  const s = summarize(g, '新FMT');
  const a = plain(s.reported[0]);
  assert.deepStrictEqual([a.activity, a.prevActivity, a.nextAction, a.prevNextAction],
    ['11/9 稟議の事前説明', '11/2 役員向けに最終提案', '11/15 稟議資料の事前確認', '11/15 稟議資料の事前確認']);
  const slide = plain(g.mtgReportSlides_(s)).find(sl => sl.title === 'RXシナリオ進捗における報告案件①：A社／店舗DX基盤の刷新');
  const prev = byRole(slide, 'report-prev').map(el => el.text);
  assert.deepStrictEqual(prev, ['今回と同じ（前回から更新なし）', '11/2 役員向けに最終提案', '今回と同じ（前回から更新なし）'],
    '課題・活動状況・次回アクションの順。活動状況は前回の文、課題・次回アクションは更新なし');
  assert.deepStrictEqual(byRole(slide, 'report-prev-frame').map(el => el.text), ['前回', '前回', '前回'], '前回の会議が無いときは見出しは「前回」だけ');
  const stale = byRole(slide, 'report-stale');
  assert.strictEqual(stale.length, 2, '「更新なし」の印は課題と次回アクションの枠');
  assert.ok(stale[1].x > byRole(slide, 'report-text')[4].x, '次回アクションの枠の右上');
  assert.strictEqual(a.prevIssue, '情報システム部門の稼働が逼迫');
  // 前回の会議の資料と比べるときは、見出しに会議日
  assert.strictEqual(g.mtgReportPrevLabel_({previous: {date: '2026/10/30', cutoff: 0}}), '前回（10/30 会議）');
  assert.strictEqual(g.mtgReportPrevLabel_({previous: {date: '2026/10/30', cutoff: 0, chosen: true}}), '前回');
  // 前回に無かった案件・前回の値が分からない案件
  const fresh = plain(g.mtgReportPrevious_({x: 0, y: 0, w: 300, h: 160, text: 'a'}, {isNew: true}, undefined, 10, '前回'));
  assert.strictEqual(fresh.find(el => el.role === 'report-prev').text, '新規の案件のため、前回はありません');
  const unknown = plain(g.mtgReportPrevious_({x: 0, y: 0, w: 300, h: 160, text: 'a'}, {isNew: false}, undefined, 10, '前回'));
  assert.strictEqual(unknown.find(el => el.role === 'report-prev').text, 'データなし');
});

test('報告案件：商談記録の「格納先リンク」は、セルのリンク（ラベル付きのリンク・HYPERLINK の数式・文字の中の URL）を開く。Sheets API で読んでも同じ', () => {
  const run = api => {
    const {gas, g, sheet} = setup({edits: false});
    // A社：「提案資料」という文字にリンク。C社：HYPERLINK の数式と、文字の中の URL の2つ
    sheet.getRange('R3').setRichTextValues([[g.SpreadsheetApp.newRichTextValue().setText('提案資料').setLinkUrl('https://drive.google.com/x').build()]]);
    sheet.getRange('R5').setFormula('=HYPERLINK("https://drive.google.com/y","音声")');
    if (api) gas.enableSheetsApi();
    const ss = api ? g.mtgApiBook_(g.SpreadsheetApp.getActiveSpreadsheet(), ['新FMT'], {}) : g.SpreadsheetApp.getActiveSpreadsheet();
    const table = g.mtgReadDeals_(ss.getSheetByName('新FMT'), {headerRow: 2});
    return plain(table.deals.map(d => d.recordLinks));
  };
  assert.deepStrictEqual(run(false), [['https://drive.google.com/x'], [], ['https://drive.google.com/y']]);
  assert.deepStrictEqual(run(true), [['https://drive.google.com/x'], [], ['https://drive.google.com/y']], 'Sheets API（hyperlink）でも読む');

  const {gas, g} = setup({edits: false});
  const record = (text, links) => plain(g.mtgReportRecord_(text, true, links));
  const two = record('a', ['https://a', 'https://b']);
  assert.strictEqual(two.text, '商談記録：格納先リンク①　格納先リンク②');
  assert.deepStrictEqual(two.spans.map(sp => [sp.url, two.text.slice(sp.start, sp.end)]), [['https://a', '格納先リンク①'], ['https://b', '格納先リンク②']]);
  assert.strictEqual(record('共有ドライブの「提案」フォルダ', []).text, '商談記録：共有ドライブの「提案」フォルダ', 'リンクが無ければ文字のまま');
  assert.strictEqual(record('', []).text, '商談記録：未入力');
  assert.deepStrictEqual(g.mtgCellLinks_(gas.ss.getSheetByName('新FMT').getRange('R3')).map(x => plain(x)),
    [['https://drive.google.com/drive/folders/abc']], '文字の中の URL');
});

test('報告案件：RXシナリオは番号だけ。案件・得意先情報は、長い得意先名・部門／役職も全文（入らなければ字を小さくする。省略しない）', () => {
  const long = Object.assign({}, DEALS[0], {customer: '株式会社サンプルリテールホールディングスジャパン', clientDept: 'デジタルマーケティング本部 本部長室',
    clientTitle: '本部長', scenario: '01-05 店舗オペレーション効率化, 02-03 需要予測'});
  const {gas, g} = setup({edits: false, deals: [long]});
  const slide = plain(g.mtgReportSlides_(summarize(g, '新FMT'))).find(sl => /報告案件①/.test(sl.title));
  assert.deepStrictEqual(byRole(slide, 'report-chip').map(el => el.text), ['01-05', '02-03'], 'シナリオは番号だけ');
  const client = byRole(slide, 'report-value')[0].text;
  assert.ok(client.indexOf('…') < 0, '切らない：' + client);
  assert.strictEqual(client.replace(/\n/g, ''), '店舗DX基盤の刷新 株式会社サンプルリテールホールディングスジャパンデジタルマーケティング本部 本部長室／本部長');
  // 枠に入らなければ、字を小さくして全文を入れる
  const huge = Object.assign({}, long, {project: 'とても長い案件名'.repeat(20), scenario: Array.from({length: 12}, (_, i) => '0' + (i % 9 + 1) + '-0' + (i % 7 + 1) + ' シナリオ' + i).join(', ')});
  const s2 = setup({edits: false, deals: [huge]});
  const slide2 = plain(s2.g.mtgReportSlides_(summarize(s2.g, '新FMT'))).find(sl => /報告案件①/.test(sl.title));
  const box = slide2.elements.find(el => el.role === 'report-value');
  assert.ok(box.text.indexOf('とても長い案件名'.repeat(20)) === 0, '案件名は全文');
  assert.ok(box.spans[0].size < 8, '字を小さくして入れる：' + box.spans[0].size);
  assert.strictEqual(byRole(slide2, 'report-chip').length, 12, 'シナリオのバッジもすべて（「ほか N件」にしない）');
  // どの部品にも「…」で省略した文が無い
  slide2.elements.concat(slide.elements).forEach(el => assert.ok(!/…/.test(String(el.text || '')), '省略しない：' + el.role + ' ' + el.text));
});

test('前回の宿題事項：宿題のシート（列＝会議日の「会議資料」「宿題」・行＝サブインダストリー）から、今の会議日より前のいちばん新しい RX Business Meeting の、この資料のサブインダストリーの宿題だけを出す', () => {
  const {gas, g} = setup({edits: false});
  // 1行目 日付（結合したセルは左の列だけ）・2行目 会議名・3行目 サブインダストリー（会議資料／宿題）
  gas.addSheet('宿題事項', [
    ['日付', '2026/10/09', '', '2026/10/16', '', '2026/10/23', '', '2026/10/30', ''],
    ['会議名', 'RX Business Meeting', '', 'RX Business Meeting', '', 'RX Service Meeting', '', 'RX Business Meeting', ''],
    ['サブインダストリー', '会議資料', '宿題', '会議資料', '宿題', '会議資料', '宿題', '会議資料', '宿題'],
    ['新FMT', '', '古い宿題', '', '・A社の決裁者を確認する\n①B社の見積もりを再提出する', '', 'サービス会議の宿題', '', '先の宿題'],
    ['新FMT2', '', '', '', '1. ドラッグストアの攻略方針を整理する', '', '', '', '']
  ], {rows: 10, columns: 9});
  const s = summarize(g, '新FMT');   // 会議日 2026/10/23（10/23 の RX Service Meeting と 10/30 は使わない）
  assert.deepStrictEqual(plain(s.homework), {date: '2026/10/16', gap: '', sheet: '宿題事項', rows: [
    {name: '新FMT', items: ['A社の決裁者を確認する', 'B社の見積もりを再提出する']},
    {name: '新FMT2', items: ['ドラッグストアの攻略方針を整理する']}
  ]});
  const slides = plain(g.mtgReportSlides_(s));
  const hw = slides[1];
  assert.strictEqual(hw.title, '前回の宿題事項');
  assert.match(texts(hw).join('\n'), /前回会議（10\/16）で出た宿題 2件。/);
  const tables = hw.elements.filter(el => el.type === 'table');
  assert.strictEqual(tables.length, 1, 'ページの幅いっぱいの1つの表');
  assert.deepStrictEqual(tables[0].rows.map(r => r.cells.map(cell => cell.text.replace(/\n/g, ''))),
    [['No.', '指摘事項'], ['1', 'A社の決裁者を確認する'], ['2', 'B社の見積もりを再提出する']], 'この資料のサブインダストリーの宿題だけ');
  assert.strictEqual(tables[0].rows[1].cells[1].size, 10.5, '件数が少なければ読みやすい大きさ');
  assert.strictEqual(slides[2].title, gas.get('MTG_TITLES').summary, '宿題のあとにサマリー');
  // 前回の宿題が空なら「該当なし」。シートにこのサブインダストリーの行が無いときも
  gas.ss.getSheetByName('宿題事項').getRange('E4').setValue('');
  assert.match(texts(plain(g.mtgReportSlides_(summarize(g, '新FMT')))[1]).join('\n'), /前回会議（10\/16）の宿題は登録されていない/);
  assert.match(texts(plain(g.mtgReportSlides_(summarize(g, '新FMT2')))[1]).join('\n'), /前回会議（10\/16）で出た宿題 1件。/);
  gas.ss.getSheetByName('宿題事項').getRange('A4').setValue('百貨店');
  assert.match(texts(plain(g.mtgReportSlides_(summarize(g, '新FMT')))[1]).join('\n'), /「宿題事項」シートに「新FMT」の行が無い/);
});

test('前回の宿題事項：多ければ字を小さくし（8pt まで）、それでも入らなければ続きのページに分けて全件を出す。宿題のシートが無ければ「データなし」', () => {
  const {gas, g} = setup({edits: false});
  const gap = plain(g.mtgReportSlides_(summarize(g, '新FMT')))[1];
  assert.match(texts(gap).join('\n'), /「宿題事項」シート（または名前が「INDEX_会議資料」で始まり、「宿題」の列があるシート）が無い/);
  const list = n => Array.from({length: n}, (_, i) => '宿題その' + (i + 1) + '：' + 'ていねいに確認して報告する'.repeat(3)).join('\n');
  gas.addSheet('宿題事項', [['日付', '2026/10/16', ''], ['会議名', 'RX Business Meeting', ''], ['サブインダストリー', '会議資料', '宿題'], ['新FMT', '', list(8)]],
    {rows: 6, columns: 3});
  const mid = plain(g.mtgReportSlides_(summarize(g, '新FMT'))).filter(sl => /^前回の宿題事項/.test(sl.title));
  assert.strictEqual(mid.length, 1, '8件なら1枚');
  const size = mid[0].elements.find(el => el.type === 'table').rows[1].cells[1].size;
  assert.ok(size < 10.5 && size >= 8, '字を小さくして入れる：' + size);
  gas.ss.getSheetByName('宿題事項').getRange('C4').setValue(list(30));
  const slides = plain(g.mtgReportSlides_(summarize(g, '新FMT'))).filter(sl => /^前回の宿題事項/.test(sl.title));
  assert.ok(slides.length >= 2, '続きのページ：' + slides.length);
  assert.strictEqual(slides[1].title, '前回の宿題事項（続き）');
  assert.strictEqual(slides.reduce((n, sl) => n + sl.elements.find(el => el.type === 'table').rows.length - 1, 0), 30, '全件');
  // コピーして貼り付けられるように、指摘事項のセルには見積もりの改行を入れない（シートの文のまま）
  slides.forEach(sl => sl.elements.find(el => el.type === 'table').rows.slice(1).forEach(r =>
    assert.ok(r.cells[1].text.indexOf('\n') < 0 && /^宿題その\d+：(ていねいに確認して報告する){3}$/.test(r.cells[1].text), r.cells[1].text)));
  slides.forEach(sl => { const t = sl.elements.find(el => el.type === 'table'); assert.ok(t.y + t.rows.reduce((a, r) => a + r.h, 0) <= gas.get('MTG_SLIDE').bottom + 0.01); });
});

test('前回の宿題事項：「宿題事項」シートが無ければ、名前が INDEX_会議資料 で始まるシート（「INDEX_会議資料_」など）の「宿題」の列から読む。今の INDEX_会議資料（行＝会議日）は読まない', () => {
  const {gas, g} = setup({edits: false});
  // 今の INDEX_会議資料（行＝会議日・列＝サブインダストリー）
  gas.addSheet('INDEX_会議資料', [['日付', '会議名', '新FMT'], ['2026/10/16', 'RX Business Meeting', '20261016_新FMT.pdf']], {rows: 5, columns: 3});
  // 宿題の列がある INDEX（列＝会議日。10/30 の RX Service Meeting は会議資料だけ）
  gas.addSheet('INDEX_会議資料_', [
    ['日付', '2026/10/09', '', '2026/10/12', '2026/10/16', ''],
    ['会議名', 'RX Business Meeting', '', 'RX Service Meeting', 'RX Business Meeting', ''],
    ['サブインダストリー', '会議資料', '宿題', '会議資料', '会議資料', '宿題'],
    ['全サブインダストリー', '20261009_全サブインダストリー.pdf', '', '', '20261016_全サブインダストリー.pdf', ''],
    ['新FMT', '20261009_新FMT.pdf', '古い宿題', '', '20261016_新FMT.pdf', '①ドラッグへの入り込みに関して中部薬品、ツルハへの提案を通じて検討し報告\n②PB商品の販促やデザインを含めた販売構成案']
  ], {rows: 8, columns: 6});
  const s = summarize(g, '新FMT');
  assert.strictEqual(s.homework.sheet, 'INDEX_会議資料_');
  assert.deepStrictEqual(plain(s.homework.rows.find(r => r.name === '新FMT').items),
    ['ドラッグへの入り込みに関して中部薬品、ツルハへの提案を通じて検討し報告', 'PB商品の販促やデザインを含めた販売構成案']);
  const hw = plain(g.mtgReportSlides_(s))[1];
  assert.match(texts(hw).join('\n'), /前回会議（10\/16）で出た宿題 2件。/);
  // どちらにも宿題の列が無ければ「データなし」
  const only = setup({edits: false});
  only.gas.addSheet('INDEX_会議資料', [['日付', '会議名', '新FMT'], ['2026/10/16', 'RX Business Meeting', 'x']], {rows: 5, columns: 3});
  assert.match(texts(plain(only.g.mtgReportSlides_(summarize(only.g, '新FMT')))[1]).join('\n'), /前回の会議の宿題事項はデータなし/);
});
