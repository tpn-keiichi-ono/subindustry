'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/*
 * RX Service MTG の事前資料（ServiceMeeting.gs）とサンプル（サンプルデータ.gs の addServiceMtgSampleData）
 * リクエストを読むシート 新FMT（requests: true。2行目が見出し）：C サブインダストリー / D 得意先 / E 案件名 / F 案件ステータス /
 *   G 想定売上規模（百万） / H 期待値調整済（改行）想定売上規模（百万） / L アカウント責任者 / O サービスのリクエスト
 */
const WIDTH = 35;
const COL = {sub: 3, customer: 4, project: 5, status: 6, amount: 7, weighted: 8, owner: 12, request: 15};
const HEADERS = {sub: 'サブインダストリー', customer: '得意先', project: '案件名', status: '案件ステータス', amount: '想定売上規模（百万）',
  weighted: '期待値調整済\n想定売上規模（百万）', owner: 'アカウント責任者', request: 'サービスのリクエスト'};
/** 得意先・案件名・ステータス・想定・期待値・責任者 */
const DEALS = [
  ['A社', 'アプリ刷新', '提案中', '30', '10', '佐藤'],
  ['B社', 'EC立ち上げ', '受注見込み', '50', '40', '鈴木'],
  ['C社', '店舗什器', '失注', '20', '5', '田中']
];

function row(values) {
  const line = new Array(WIDTH).fill('');
  Object.keys(values).forEach(key => { line[COL[key] - 1] = values[key]; });
  return line;
}

/** 新FMT（案件）・リクエスト・タスク管理のシートを用意する（リクエスト・タスクはまだ無い） */
function setup() {
  const gas = createGas();
  const g = gas.global;
  ['新FMT', '新FMT2'].forEach(name => {
    gas.addSheet(name, [row({customer: name}), row(HEADERS)].concat(DEALS.map(d => row({sub: '食品スーパー', customer: d[0], project: d[1],
      status: d[2], amount: d[3], weighted: d[4], owner: d[5]}))), {rows: 20, columns: WIDTH});
  });
  g.setupRequestSheet();
  g.setupTaskSheet();
  gas.writes.length = 0;
  gas.alerts.length = 0;
  gas.toasts.length = 0;
  return {gas, g, request: gas.ss.getSheetByName('リクエスト'), task: gas.ss.getSheetByName('タスク管理')};
}

const headersOf = sheet => Array.from(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
/** 1行目の見出しで列を探して、2行目から最後の行までの表示の値（labels の列だけ） */
function rowsOf(sheet, labels) {
  const headers = headersOf(sheet);
  if (sheet.getLastRow() < 2) return [];
  return plain(sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues())
    .map(line => labels.map(label => line[headers.indexOf(label)]));
}

/** リクエストのシートに1行書く（利用者の入力として。ID は書いたとおり） */
function addRequest(gas, sheet, values) {
  const headers = headersOf(sheet);
  const line = headers.map(h => (h in values ? values[h] : ''));
  gas.asUser(() => sheet.getRange(sheet.getLastRow() + 1, 1, 1, line.length).setValues([line]));
}

function addTask(gas, sheet, values) {
  const headers = headersOf(sheet);
  const line = headers.map(h => (h in values ? values[h] : ''));
  gas.asUser(() => sheet.getRange(sheet.getLastRow() + 1, 1, 1, line.length).setValues([line]));
}

/* ---------------- サンプル ---------------- */

test('addServiceMtgSampleData：リクエスト・タスクのシートの最後の行の下にサンプルを書く（ID を振り、リクエストに【サンプル】。案件はシートの進行中の案件を使う）', () => {
  const {gas, g, request, task} = setup();
  addRequest(gas, request, {'ID': 'SR-0001', 'サブインダストリー': '食品スーパー', '得意先': 'A社', '案件名': 'アプリ刷新', 'リクエスト': '本物のリクエスト'});
  g.addServiceMtgSampleData();
  assert.match(gas.alerts[0].message, /リクエスト 13件（リクエストの先頭に「【サンプル】」。得意先・案件名は、シートの案件 2件と架空の名前 11件）[\s\S]*タスク 7件/);
  const S = gas.get('SMP_OPTIONS').service;
  const requests = rowsOf(request, ['ID', '得意先', '案件名', 'リクエスト', '判断', 'サービス案']);
  assert.strictEqual(requests.length, 1 + S.requests.length, '今の行の下に足す');
  assert.deepStrictEqual(requests[0], ['SR-0001', 'A社', 'アプリ刷新', '本物のリクエスト', '', ''], '今の行は変えない');
  assert.deepStrictEqual(requests.slice(1).map(r => r[0]), plain(S.requests.map((_, i) => 'SR-' + String(i + 2).padStart(4, '0'))), 'ID は続きから');
  assert.ok(requests.slice(1).every(r => r[3].indexOf('【サンプル】') === 0));
  // 案件：期待値の大きい進行中の案件から（失注の C社 は使わない）。案件が足りなければ架空の名前
  assert.deepStrictEqual(requests.slice(1, 4).map(r => r[1] + '/' + r[2]), ['B社/EC立ち上げ', 'A社/アプリ刷新', 'サンプル化粧品/EC連携']);
  assert.deepStrictEqual(requests[1].slice(4), ['サービス化検討', '購買データ分析']);
  assert.strictEqual(requests[11][4], '未判断', '判断が空のサンプルは「未判断」');

  const tasks = rowsOf(task, ['サービス案', 'タスク', '担当者', '状況']);
  assert.strictEqual(tasks.length, S.tasks.length);
  assert.match(tasks[0][0], /^SR-0002 【サンプル】会員アプリの購買データから/, 'サービス案はリクエストの候補（ID とリクエスト）');
  const choiceSheet = gas.ss.getSheetByName('__TASK_CHOICES');
  const choices = plain(choiceSheet.getRange(1, 1, choiceSheet.getMaxRows(), 1).getDisplayValues()).map(r => r[0]);
  assert.ok(tasks.every(t => choices.indexOf(t[0]) >= 0), 'プルダウンの候補にある値');
  assert.ok(gas.writes.filter(w => w.sheet === 'リクエスト' || w.sheet === 'タスク管理').every(w => w.locked), '書き込みはロックの中');
  assert.ok(!gas.writes.some(w => w.sheet === '新FMT'), '案件のシートには書かない');

  // もう一度実行しても、二重には書かない
  g.addServiceMtgSampleData();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /サンプルのリクエスト 13件・タスク 7件が、すでにあります。/);
  assert.strictEqual(rowsOf(request, ['ID']).length, 14);
});

test('removeServiceMtgSampleData：【サンプル】の付いたリクエスト・タスクの行だけを消す', () => {
  const {gas, g, request, task} = setup();
  addRequest(gas, request, {'ID': 'SR-0001', '得意先': 'A社', '案件名': 'アプリ刷新', 'リクエスト': '本物のリクエスト'});
  g.addServiceMtgSampleData();
  g.svcLinkTaskServiceColumn_(gas.ss, true);
  addTask(gas, task, {'サービス案': 'SR-0001 本物のリクエスト', 'タスク': '本物のタスク'});
  g.removeServiceMtgSampleData();
  assert.match(gas.alerts[gas.alerts.length - 2].message, /リクエスト 13行・「タスク管理」のタスク 7行/);
  assert.deepStrictEqual(rowsOf(request, ['ID', 'リクエスト']), [['SR-0001', '本物のリクエスト']]);
  assert.deepStrictEqual(rowsOf(task, ['サービス案', 'タスク']), [['SR-0001 本物のリクエスト', '本物のタスク']]);
  g.removeServiceMtgSampleData();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /サンプルのリクエスト・タスク（「【サンプル】」の付いたもの）はありません。/);
});

/* ---------------- まとめる ---------------- */

/** リクエスト4件（2つのサービス案・棄却・判断待ち）とタスク3件 */
function fill(gas, request, task) {
  addRequest(gas, request, {'ID': 'SR-0001', 'サブインダストリー': '食品スーパー', '得意先': 'A社', '案件名': 'アプリ刷新', 'リクエスト': '会員の購買分析をしたい',
    '判断': 'サービス化検討', 'サービス部門からのフィードバック': '分析基盤で対応できる', 'サービス案': '購買分析', '主管本部': 'DX事業本部', '担当者': '山本'});
  addRequest(gas, request, {'ID': 'SR-0002', 'サブインダストリー': 'ドラッグストア', '得意先': 'B社', '案件名': 'EC立ち上げ', 'リクエスト': '売れ筋を共有したい',
    '判断': 'サービス化検討', 'サービス案': '購買 分析'});   // 空白の違いは同じサービス案
  addRequest(gas, request, {'ID': 'SR-0003', 'サブインダストリー': '食品スーパー', '得意先': 'C社', '案件名': '店舗什器', 'リクエスト': '什器を借りたい',
    '判断': '棄却', 'サービス案': '什器サブスク'});
  addRequest(gas, request, {'ID': 'SR-0004', 'サブインダストリー': '食品スーパー', '得意先': 'Z社', '案件名': '架空', 'リクエスト': 'チラシをなくしたい'});
  gas.global.svcLinkTaskServiceColumn_(gas.ss, true);
  addTask(gas, task, {'サービス案': 'SR-0001 会員の購買分析をしたい', 'タスク': '指標を試算する', '本部': 'DX事業本部', '担当者': '山本', '状況': '対応中'});
  addTask(gas, task, {'サービス案': 'SR-0002 売れ筋を共有したい', 'タスク': '要件を整理する', '状況': '完了'});
  addTask(gas, task, {'サービス案': 'SR-0009 消したリクエスト', 'タスク': '迷子のタスク', '担当者': '佐藤', '状況': '未着手'});
}

test('smtgCollect_：同じ名前のサービス案のリクエストをまとめ（棄却は除く）、タスクを ID で結び付け、元の案件の期待値（進行中だけ）・確認事項を出す', () => {
  const {gas, g, request, task} = setup();
  fill(gas, request, task);
  const data = g.smtgCollect_(gas.ss, new Date(), '2026/10/30');
  assert.deepStrictEqual(plain(data.kpi), {requests: 4, undecided: 1, considering: 2, rejected: 1, services: 1, serviceRequests: 2, customers: 2,
    tasks: 3, done: 1});
  const s = data.services[0];
  assert.deepStrictEqual(plain([s.name, s.requests.map(r => r.id), s.tasks.map(t => t.task), s.customers, s.subIndustries, s.owners]),
    ['購買分析', ['SR-0001', 'SR-0002'], ['指標を試算する', '要件を整理する'], ['A社', 'B社'], ['食品スーパー', 'ドラッグストア'], ['DX事業本部 山本']]);
  assert.deepStrictEqual([s.weighted, s.active.length, plain(s.counts)], [50, 2, {total: 2, todo: 0, doing: 1, done: 1, blank: 0, other: 0}],
    'A社（提案中 10）と B社（受注見込み 40）');
  assert.deepStrictEqual(plain(data.checks.map(c => [c.type, c.target])), [
    ['判断とサービス案の不整合', 'SR-0003 什器を借りたい'],
    ['タスクの担当者未入力', '要件を整理する（タスク管理 3行目）'],
    ['リクエストが見つからない', '迷子のタスク（タスク管理 4行目）'],
    ['元の案件が見つからない', 'SR-0004 チラシをなくしたい']
  ]);
  assert.deepStrictEqual(plain(data.fit.columns), ['食品スーパー', 'ドラッグストア']);
  assert.deepStrictEqual(plain(data.fit.rows.map(r => r.counts)), [[1, 1]]);
  assert.ok(data.moves.first, '前回の会議の記録が無い');
  assert.ok(!data.sample);
});

/* ---------------- スライド ---------------- */

test('スライド：紫の表紙（ひな形の3枚目）・サマリー・一覧・広がり・サービス案・判断待ち・扉・Appendix ①〜③。表の見出しは紫、Appendix はグレー。どの部品もページに入る', () => {
  const {gas, g, request, task} = setup();
  g.addServiceMtgSampleData();
  const data = g.smtgCollect_(gas.ss, new Date(), '2026/10/30');
  const slides = plain(g.smtgSlides_(data, '2026/10/30', false));
  const titles = slides.map(sl => sl.title.replace(/（続き）$/, ''));
  const order = ['表紙', 'サマリー', 'サービス案の一覧', '顧客・業界の広がり'];
  assert.deepStrictEqual(titles.slice(0, 4), order);
  assert.ok(titles.indexOf('サービス案：購買データ分析') === 4);
  const divider = titles.indexOf('Appendix');
  assert.strictEqual(titles[divider - 1], '判断待ちのリクエスト');
  assert.deepStrictEqual(Array.from(new Set(titles.slice(divider + 1))),
    ['Appendix ①　リクエスト（一覧）', 'Appendix ②　タスク（一覧）', 'Appendix ③　入力内容の確認事項（一覧）']);
  // 表紙：紫（service）。ひな形の決まった文字
  assert.deepStrictEqual([slides[0].cover, slides[0].service, slides[0].template.title, slides[0].template.subtitle],
    [true, true, 'サービス案の検討状況', 'RX Service MTG　事前資料【試作】']);
  assert.strictEqual(g.mtgCoverTone_(slides[0]), 'combined', '絵柄は紫');
  // 色：本編の表の見出しは紫、Appendix はグレー。判断のラベル（サービス化検討）は Appendix でも紫
  const head = sl => sl.elements.find(el => el.type === 'table').rows[0].cells[0].fill;
  assert.strictEqual(head(slides[2]), '#7030A0');
  assert.strictEqual(head(slides[divider + 1]), '#7F7F7F');
  const pill = slides[divider + 1].elements.find(el => el.role === 'pill' && el.text === 'サービス化検討');
  assert.strictEqual(pill.fill, '#7030A0');
  assert.strictEqual(gas.get('MTG_OPTIONS').colors.green, '#70AD47', '作り終えたら会議サマリの色に戻す');
  // 出典の行・ページ番号
  const footer = slides[1].elements.find(el => el.role === 'footer').text;
  assert.strictEqual(footer, '【試作】【サンプルデータを含む】RX Service MTG（10/30 会議資料）　金額の単位：百万円');
  assert.deepStrictEqual(slides.slice(1).filter(sl => !sl.divider).map(sl => Number(sl.elements.find(el => el.role === 'page').text)),
    slides.map((_, i) => i + 1).slice(1).filter(i => i !== divider + 1));
  // サービス案の一覧のページの列は、そのサービス案のページ
  const list = slides[2].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(list.rows.slice(1).map(r => slides[Number(r.cells[7].text) - 1].title.replace(/（続き）$/, '')),
    plain(data.services.map(s => 'サービス案：' + s.name)));
  // 議題と資料のページ
  const agenda = slides[1].elements.filter(el => el.role === 'agenda').map(el => el.text.split('\n')[0]);
  const first = titles.indexOf('サービス案：購買データ分析') + 1;
  const pending = titles.indexOf('判断待ちのリクエスト') + 1;
  assert.deepStrictEqual(agenda[0], 'P.3・P.' + first + '〜' + (pending - 1));
  assert.deepStrictEqual(agenda[1], 'P.4・P.' + first + '〜' + (pending - 1));
  // 扉の一覧のページ
  const toc = slides[divider].elements.filter(el => el.role === 'toc');
  assert.deepStrictEqual(toc.filter((_, i) => i % 3 === 2).map(el => Number(el.text)),
    ['Appendix ①', 'Appendix ②', 'Appendix ③'].map(label => titles.findIndex(t => t.indexOf(label) === 0) + 1));
  // どの部品もページに入る。文字は切らない（表の見出し・ラベル）
  const S = gas.get('MTG_SLIDE');
  slides.forEach((sl, i) => sl.elements.forEach(el => {
    const h = el.type === 'table' ? el.rows.reduce((a, r) => a + r.h, 0) : el.h;
    if (!sl.cover && !sl.divider && ['footer', 'page'].indexOf(el.role) < 0 && h > 1) assert.ok(el.y + h <= S.bottom + 0.01, (i + 1) + '枚目：' + (el.role || el.type));
    if (el.role === 'pill' || el.role === 'title') assert.ok(!/…$/.test(el.text), (i + 1) + '枚目：' + el.text);
    if (el.type === 'table') assert.ok(!el.rows[0].cells.some(cell => /…/.test(cell.text)), (i + 1) + '枚目の表の見出し');
  }));
});

test('スライド：データが無いときも同じ型（表の見出しの下に「該当なし」）。前回の記録が無ければ、前回以降の動きは灰色の「データなし」', () => {
  const {gas, g} = setup();
  const data = g.smtgCollect_(gas.ss, new Date(), '2026/10/30');
  const slides = plain(g.smtgSlides_(data, '2026/10/30', true));
  assert.deepStrictEqual(slides.map(sl => sl.title), ['表紙', 'サマリー', 'サービス案の一覧', '顧客・業界の広がり', '判断待ちのリクエスト', 'Appendix',
    'Appendix ①　リクエスト（一覧）', 'Appendix ②　タスク（一覧）', 'Appendix ③　入力内容の確認事項（一覧）']);
  assert.deepStrictEqual(slides[2].elements.filter(el => el.role === 'empty').map(el => el.text.split('\n')[0]), ['該当なし']);
  const moves = slides[1].elements.find(el => el.type === 'table');
  assert.ok(moves.rows.slice(1).every(r => r.cells[2].text === 'データなし：前回の会議の記録なし'));
  assert.deepStrictEqual(slides[5].elements.filter(el => el.role === 'toc').filter((_, i) => i % 3 === 1).map(el => el.text),
    ['リクエスト（該当なし）', 'タスク（該当なし）', '入力内容の確認事項（該当なし）']);
});

/* ---------------- 作る（試作・本番） ---------------- */

test('previewServiceMtgSummary：試作は PDF を保存し（ファイル名に「試作」）、会議日は記録しない', () => {
  const {gas, g, request, task} = setup();
  fill(gas, request, task);
  gas.writes.length = 0;
  gas.promptAnswers = ['10/30'];
  g.previewServiceMtgSummary();
  const year = g.mtgParseDate_('10/30', new Date()).slice(0, 4);
  assert.deepStrictEqual(gas.decks.map(d => d.title), [year + '1030_RX Service Meeting用資料_試作']);
  assert.ok(gas.driveFiles.some(f => f.name === year + '1030_RX Service Meeting用資料_試作.pdf'));
  const done = gas.alerts[gas.alerts.length - 1];
  assert.strictEqual(done.title, '🟪RXサービスMTG用');
  assert.match(done.message, /RX Service MTG の事前資料を作りました（試作）：[\s\S]*リクエスト 4件（判断待ち 1件）・サービス案 1件・タスク 3件・確認事項 4件[\s\S]*試作のため、会議日は記録していません。/);
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('SMTG_MEETINGS'), null);
  assert.strictEqual(gas.ss.getSheetByName('INDEX_会議資料'), null, '試作はインデックスに書かない');
  assert.ok(!gas.writes.some(w => ['リクエスト', 'タスク管理', '新FMT'].indexOf(w.sheet) >= 0), 'シートには書かない');
});

test('createServiceMtgSummary：本番は会議日と判断・状況を記録し、インデックスの「RX Service MTG」の行に書く。次の会議の資料に前回以降の動きと前回比を出す', () => {
  const {gas, g, request, task} = setup();
  fill(gas, request, task);
  const meeting = g.mtgParseDate_('10/09', new Date());
  gas.promptAnswers = ['10/09'];
  g.createServiceMtgSummary();
  assert.match(gas.alerts[gas.alerts.length - 2].message, /本番の資料を作り、会議日を記録します。[\s\S]*前回の会議：記録なし/);
  const props = g.PropertiesService.getDocumentProperties();
  const list = JSON.parse(props.getProperty('SMTG_MEETINGS'));
  assert.deepStrictEqual(list.map(m => [m.date, m.lastId, m.counts.services]), [[meeting, 4, 1]]);
  const detail = JSON.parse(props.getProperty('SMTG_MEETING_' + meeting.replace(/\//g, '')));
  assert.strictEqual(detail.decisions, '1:1,2:1,3:2,4:0');
  // インデックス：種別 RX Service MTG の行の「全サブインダストリー」の列
  const index = gas.ss.getSheetByName('INDEX_会議資料');
  const day = d => d.replace(/\/0?/g, '/');
  assert.deepStrictEqual(plain(index.getRange(1, 1, 2, 3).getDisplayValues()),
    [['日付', '種別', '全サブインダストリー'], [day(meeting), 'RX Service MTG', meeting.replace(/\//g, '') + '_RX Service Meeting用資料.pdf']]);
  // 同じ日付でも、RX Business MTG の資料は別の行
  const at = url => ({file: {getUrl: () => url, getName: () => url}});
  g.mtgWriteIndex_('新FMT', meeting, at('business'));
  assert.deepStrictEqual(plain(index.getRange(2, 2, 2, 1).getDisplayValues()).map(r => r[0]).sort(), ['RX Business MTG', 'RX Service MTG']);

  // 次の会議までに：リクエストを足し、判断を決め、タスクを完了・追加する
  addRequest(gas, request, {'ID': 'SR-0005', '得意先': 'A社', '案件名': 'アプリ刷新', 'リクエスト': '新しいリクエスト'});
  const headers = headersOf(request);
  gas.asUser(() => request.getRange(5, headers.indexOf('判断') + 1).setValue('サービス化検討'));   // SR-0004：未判断 → サービス化検討
  const theaders = headersOf(task);
  gas.asUser(() => task.getRange(2, theaders.indexOf('状況') + 1).setValue('完了'));
  g.svcLinkTaskServiceColumn_(gas.ss, true);
  addTask(gas, task, {'サービス案': 'SR-0001 会員の購買分析をしたい', 'タスク': '料金を考える', '担当者': '山本', '状況': '未着手'});
  const data = g.smtgCollect_(gas.ss, new Date(), g.mtgParseDate_('11/13', new Date()));
  assert.strictEqual(data.previous.date, meeting);
  assert.deepStrictEqual(plain([data.moves.newRequests.map(r => r.id), data.moves.decided.map(d => [d.request.id, d.from, d.request.decision]),
    data.moves.completed.map(t => t.task), data.moves.newTasks.map(t => t.task)]),
  [['SR-0005'], [['SR-0004', '未判断', 'サービス化検討']], ['指標を試算する'], ['料金を考える']]);
  const slides = plain(g.smtgSlides_(data, data.meeting, true));
  assert.match(slides[1].elements.find(el => el.role === 'lead').text + slides[1].elements.filter(el => el.role === 'lead')[1].text,
    /前回の会議（\d+\/\d+）以降：新規リクエスト 1件・判断の変更 1件・完了したタスク 1件・新しいタスク 1件。/);
  const tiles = slides[1].elements.filter(el => el.role === 'tile-value').map(el => el.text);
  assert.deepStrictEqual(tiles, ['5件　前回比 +1', '1件　前回比 ±0', '3件　前回比 +1', '1件　前回比 ±0', '2／4件　完了 前回比 +1']);
  const moves = slides[1].elements.find(el => el.type === 'table');
  assert.deepStrictEqual(moves.rows.slice(1).map(r => r.cells[2].text),
    ['SR-0005 新しいリクエスト', 'SR-0004 未判断 ▶ サービス化検討', '指標を試算する（購買分析）', '料金を考える（購買分析）']);
  // 判断待ちのページ：前回以降に登録したリクエストは ID の下に「新規」
  const pending = slides.find(sl => sl.title === '判断待ちのリクエスト').elements.find(el => el.type === 'table');
  assert.deepStrictEqual(pending.rows.slice(1).map(r => r.cells[0].text), ['SR-0005\n新規']);
});

test('ひな形：RX Service MTG の表紙はひな形の3枚目（紫）を使い、1枚目（緑）は使わない。3枚目が無ければ紫の絵柄で描く', () => {
  const {gas, g, request, task} = setup();
  fill(gas, request, task);
  const page = id => ({objectId: id, pageElements: [{objectId: id + '_title', transform: {translateY: 231.8 * 12700},
    shape: {text: {textElements: [{textRun: {content: 'XXXXX\n', style: {fontSize: {magnitude: 40}}}}]}}}]});
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789';
  gas.slideTemplates[id] = {pageSize: {width: {magnitude: 12192000, unit: 'EMU'}}, slides: ['g_cover', 'g_divider', 'g_purple'].map(page)};
  g.PropertiesService.getScriptProperties().setProperty('MTG_TEMPLATE_URL', 'https://docs.google.com/presentation/d/' + id + '/edit');
  gas.promptAnswers = ['10/30'];
  g.previewServiceMtgSummary();
  const r = plain(gas.decks[gas.decks.length - 1].requests);
  assert.deepStrictEqual(r.filter(q => q.deleteObject && /^g_/.test(q.deleteObject.objectId)).map(q => q.deleteObject.objectId), ['g_cover'], '緑の表紙は使わない');
  assert.deepStrictEqual(r.filter(q => q.insertText && q.insertText.objectId === 'g_purple_title').map(q => q.insertText.text), ['サービス案の検討状況']);
  assert.deepStrictEqual(gas.slideBackgrounds, [], '表紙・扉はどちらもひな形のページ');

  // 2枚しかないひな形：表紙は紫の絵柄の画像を背景に描く（緑の1枚目は使わない）
  gas.slideTemplates[id] = {pageSize: {width: {magnitude: 12192000, unit: 'EMU'}}, slides: ['g_cover', 'g_divider'].map(page)};
  gas.promptAnswers = ['10/30'];
  g.previewServiceMtgSummary();
  assert.deepStrictEqual(gas.slideBackgrounds.map(b => b.page), ['mtg_p1']);
});
