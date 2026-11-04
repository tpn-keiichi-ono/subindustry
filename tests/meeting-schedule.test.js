'use strict';

const {test, assert} = require('./lib/harness');
const {createGas: createMockGas, plain} = require('./lib/gas-mock');
const {withRunDialog} = require('./lib/meeting-run');
const createGas = options => withRunDialog(createMockGas(options));

/*
 * 会議の日程（MeetingSchedule.gs）：Google カレンダーから RX Business Meeting・RX Service Meeting の日程を読んで「INDEX_会議資料」に先に出し、
 * 資料を作る画面（MeetingRunDialog）で、その一覧から会議日程を選んで作る
 */

const HEAD = ['', '', '', '得意先', '案件名', '案件ステータス'];

function setup() {
  const gas = createGas();
  const g = gas.global;
  ['新FMT', '新FMT2'].forEach(name => {
    gas.addSheet(name, [[name], HEAD, ['', '', '', 'A社', 'アプリ刷新', '提案中']], {rows: 10, columns: 6});
  });
  return {gas, g};
}

/** 今年の月・日の Date（日本時間の時刻） */
function at(g, month, day, hour) {
  const year = Number(g.mtgParseDate_('1/1', new Date()).slice(0, 4));
  return new Date(Date.UTC(year, month - 1, day, (hour || 9) - 9));
}

const indexRows = gas => {
  const sheet = gas.ss.getSheetByName('INDEX_会議資料');
  return plain(sheet.getRange(1, 1, sheet.getLastRow(), 3).getDisplayValues());
};

test('importMeetingSchedules：カレンダーの予定の名前で RX Business Meeting・RX Service Meeting を見分け、「INDEX_会議資料」に会議日と会議名の行を足す（新しい会議日が上。ある行はそのまま）', () => {
  const {gas, g} = setup();
  const y = g.mtgParseDate_('1/1', new Date()).slice(0, 4);
  const now = new Date();
  // 今日の前後の会議（読む範囲：pastDays・aheadDays）
  const shift = days => new Date(now.getTime() + days * 86400000);
  const ymd = d => g.Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  gas.calendarEvents = [
    {title: '【定例】RX Business MTG（パイプライン会議）', start: shift(-21), end: shift(-21)},
    {title: 'ＲＸ　Business　Meeting', start: shift(7), end: shift(7)},              // 全角・空白の違いは同じ
    {title: 'RX Service MTG（サービス改善）', start: shift(14), end: shift(14)},
    {title: 'RX Service Meeting', start: shift(14), end: shift(14)},                  // 同じ日の同じ会議は1つ
    {title: 'RX Business Meeting', start: shift(35), end: shift(35), status: 'NO'},   // 断った予定は読まない
    {title: '【中止】RX Business Meeting', start: shift(42), end: shift(42)},         // 中止は読まない
    {title: '部会', start: shift(1), end: shift(1)},
    {title: 'RX Business Meeting', start: shift(500), end: shift(500)}                // 読む範囲の外
  ];
  g.importMeetingSchedules();
  const message = gas.alerts[gas.alerts.length - 1].message;
  assert.match(message, /RX Business Meeting 2件・RX Service Meeting 1件の日程を読みました。\n「INDEX_会議資料」に 3行を足しました（すでにある 0行はそのまま）/);
  assert.match(message, new RegExp('・RX Business Meeting：次回 ' + y + '年|・RX Business Meeting：次回 \\d{4}年'));
  const day = d => ymd(d).replace(/\/0?/g, '/');
  assert.deepStrictEqual(indexRows(gas), [
    ['日付', '会議名', '全サブインダストリー'],
    [day(shift(14)), 'RX Service Meeting', ''],
    [day(shift(7)), 'RX Business Meeting', ''],
    [day(shift(-21)), 'RX Business Meeting', '']
  ]);
  assert.ok(gas.writes.filter(w => w.sheet === 'INDEX_会議資料').every(w => w.locked), '書き込みはロックの中');
  // もう一度読んでも行は増えない
  g.importMeetingSchedules();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /0行を足しました（すでにある 3行はそのまま）/);
  // 読んだ範囲は今日の pastDays 日前〜aheadDays 日後
  const read = gas.calendarReads[0];
  assert.deepStrictEqual([Math.round((now - read.start) / 86400000), Math.round((read.end - now) / 86400000)], [365, 400]);
});

test('importMeetingSchedules：前の版の見出し「種別」は「会議名」に書き換え、種別が空・前の名前の行は同じ会議の行として使う。カレンダーが開けない・予定が無いときは知らせる', () => {
  const {gas, g} = setup();
  const y = g.mtgParseDate_('1/1', new Date()).slice(0, 4);
  gas.addSheet('INDEX_会議資料', [['日付', '種別', '全サブインダストリー'], [y + '/10/23', 'RX Business MTG', '資料A'], [y + '/10/09', '', '']]);
  gas.calendarEvents = [{title: 'RX Business Meeting', start: at(g, 10, 23), end: at(g, 10, 23)},
    {title: 'RX Business Meeting', start: at(g, 10, 9), end: at(g, 10, 9)}];
  gas.get('MSCH_OPTIONS').pastDays = 4000;
  g.importMeetingSchedules();
  assert.deepStrictEqual(indexRows(gas), [['日付', '会議名', '全サブインダストリー'], [y + '/10/23', 'RX Business Meeting', '資料A'],
    [y + '/10/09', 'RX Business Meeting', '']]);
  gas.calendarEvents = [];
  g.importMeetingSchedules();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /RX Business Meeting 0件・RX Service Meeting 0件の日程を読みました。\n予定の名前に、MeetingSchedule\.gs の MSCH_OPTIONS\.titles の言葉/);
  gas.get('MSCH_OPTIONS').calendarId = 'team@example.com';
  assert.throws(() => g.importMeetingSchedules(), /Google カレンダー（ID：team@example\.com）を開けません。/);
});

test('資料を作る関数は画面（MeetingRunDialog）を開く。画面の会議日程は「INDEX_会議資料」のその会議の日程（新しい順に並べ、初めは今日以降でいちばん近い会議日）', () => {
  const {gas, g} = setup();
  const now = new Date();
  const ymd = days => g.Utilities.formatDate(new Date(now.getTime() + days * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  gas.addSheet('INDEX_会議資料', [['日付', '会議名', '全サブインダストリー', '新FMT'],
    [ymd(21), 'RX Business Meeting', '', ''], [ymd(14), 'RX Service Meeting', '', ''], [ymd(7), 'RX Business Meeting', '', ''],
    [ymd(-21), 'RX Business Meeting', '', '資料'], [ymd(-28), 'ほかの会議', '', '']]);
  gas.realRun.createBusinessMtgSummaries();
  const opened = () => {
    const d = gas.dialogs[gas.dialogs.length - 1];
    return [d.html.file, d.html.data.kind, d.html.data.final];
  };
  assert.deepStrictEqual(opened(), ['MeetingRunDialog', 'business', true]);
  gas.realRun.previewServiceMtgSummaries();
  assert.deepStrictEqual(opened(), ['MeetingRunDialog', 'service', false]);
  gas.realRun.createBusinessMtgCombinedSummary();
  assert.deepStrictEqual(opened(), ['MeetingRunDialog', 'combined', true]);
  assert.match(gas.dialogs[0].html.getContent(), /data-kind="<\?= kind \?>" data-final="<\?= final \?>"/);

  const data = plain(g.getMeetingRunData('business', true));
  assert.deepStrictEqual(data.meetings.map(m => [m.date, m.filled, m.past]), [[ymd(-21), true, true], [ymd(7), false, false], [ymd(21), false, false]]);
  assert.strictEqual(data.selected, ymd(7));
  assert.deepStrictEqual([data.meetingName, data.title], ['RX Business Meeting', '会議サマリ（RX Business Meeting）を作る']);
  assert.deepStrictEqual(plain(g.mtgIndexMeetings_('RX Service Meeting')).map(m => m.date), [ymd(14)], 'RX Service Meeting の画面には、その会議の日程');
  assert.strictEqual(g.getMeetingRunData('combined', true).sheets.length, 0, '統合版はシートを選ばない');
  // 画面の「カレンダーから読み込む」
  gas.calendarEvents = [{title: 'RX Service Meeting', start: new Date(now.getTime() + 42 * 86400000), end: new Date(now.getTime() + 42 * 86400000)}];
  const imported = g.importMeetingSchedulesFromDialog('combined', true);
  assert.match(imported.message, /1行を足しました/);
  assert.deepStrictEqual(plain(g.mtgIndexMeetings_('RX Service Meeting')).map(m => m.date), [ymd(14), ymd(42)]);
  assert.deepStrictEqual(plain(imported.data.meetings.map(m => m.date)), [ymd(-21), ymd(7), ymd(21)], '画面の一覧（統合版は RX Business Meeting の日程）');
});

test('前回の資料の記録が無ければ、会議日程の一覧の前の会議日（その日の 0時）からの変化を出す（RX Business Meeting）', () => {
  const {gas, g} = setup();
  const y = g.mtgParseDate_('1/1', new Date()).slice(0, 4);
  gas.addSheet('INDEX_会議資料', [['日付', '会議名', '全サブインダストリー'], [y + '/10/23', 'RX Business Meeting', ''], [y + '/10/09', 'RX Business Meeting', ''],
    [y + '/10/16', 'RX Service Meeting', '']]);
  const check = g.checkMeetingRun('business', true, y + '/10/23', ['新FMT']);
  assert.match(check.text, /・新FMT：前回会議（10\/9）以降（前回の資料の記録が無いため、会議日から）/, 'RX Service Meeting の日程は数えない');
  const {summaries} = g.mtgSummaries_(gas.ss, ['新FMT'], new Date(), false, y + '/10/23');
  const s = summaries[0];
  assert.deepStrictEqual([s.previous.date, s.previous.scheduled, s.since.getTime(), s.first],
    [y + '/10/09', true, g.Utilities.parseDate(y + '/10/09', 'Asia/Tokyo', 'yyyy/MM/dd').getTime(), false]);
  assert.strictEqual(g.mtgPeriodText_(s), '前回会議（10/9）以降');
  // 一覧の最初の会議なら、前回は無い（直近14日間）
  assert.match(g.checkMeetingRun('business', false, y + '/10/09', ['新FMT']).text, /・新FMT：直近14日間（前回の会議が無い）/);
});

test('資料を作る画面の読み込みは、手順ごとの時間を実行ログに残す。diagnoseMeetingRunDialog() は同じ読み込みの時間と時点コピーの数を出す（読むだけ）', () => {
  const {gas, g} = setup();
  const logs = [];
  const original = g.console.log;
  g.console.log = (...args) => { logs.push(args.join(' ')); };
  try {
    const data = g.getMeetingRunData('business', false);
    assert.ok(data.sheets.length >= 1);
    const line = logs.find(text => /^資料を作る画面の読み込み：/.test(text));
    assert.match(line, /会議日程（INDEX_会議資料） [\d.]+秒・シートと前回の資料 [\d.]+秒・比べる時点（時点コピー \d+件） [\d.]+秒（合計 [\d.]+秒）$/);
    logs.length = 0;
    const writes = gas.writes.length;
    g.diagnoseMeetingRunDialog();
    const text = logs.join('\n');
    assert.match(text, /資料を作る画面の読み込みの確かめ/);
    assert.match(text, /読み込み：会議日程/);
    assert.match(text, /会議日程：\d+件・シート：\d+・比べる時点：\d+件/);
    assert.match(text, /時点コピーの保存先：/);
    assert.strictEqual(gas.writes.length, writes, 'シートには書き込まない');
  } finally {
    g.console.log = original;
  }
});

test('会議日程は「INDEX_会議資料」の控え（ドキュメントのプロパティ）から読む。スクリプトが書いたら控えも直し、手で直したら refreshMeetingIndexCache() で読み直す', () => {
  const {gas, g} = setup();
  const y = g.mtgParseDate_('1/1', new Date()).slice(0, 4);
  const sheet = gas.addSheet('INDEX_会議資料', [['日付', '会議名', '全サブインダストリー', '新FMT'], [y + '/10/23', 'RX Business Meeting', '', ''],
    [y + '/10/09', 'RX Business Meeting', '', '資料']]);
  const dates = () => plain(g.mtgIndexMeetings_('RX Business Meeting')).map(m => [m.date, m.filled]);
  assert.deepStrictEqual(dates(), [[y + '/10/09', true], [y + '/10/23', false]], '控えが無ければシートから読んで控えを作る');
  const key = gas.get('MTG_OPTIONS').index.cacheKey;
  const props = g.PropertiesService.getDocumentProperties();
  assert.ok(JSON.parse(props.getProperty(key)).rows.length === 2);
  // 控えがあればシートのセルは読まない（手で足した行は、読み直すまで出ない）
  gas.asUser(() => sheet.appendRow([y + '/11/06', 'RX Business Meeting', '', '']));
  assert.deepStrictEqual(dates().map(d => d[0]), [y + '/10/09', y + '/10/23']);
  g.refreshMeetingIndexCache();
  assert.deepStrictEqual(dates().map(d => d[0]), [y + '/10/09', y + '/10/23', y + '/11/06'], 'refreshMeetingIndexCache() で読み直す');
  // 終わりは待たない通知（toast）で知らせる。ui.alert は OK を押すまで実行が終わらない（エディタからの実行が「実行中」のままになる）
  assert.strictEqual(gas.alerts.length, 0, 'refreshMeetingIndexCache() は確かめの画面を出さない');
  assert.ok(gas.toasts.some(t => /3行）を読み直しました（.*秒/.test(t.message)), '読んだ行の数とかかった時間を通知する');
  // スクリプトが書いたら控えも直す（カレンダーから足した行・資料を書いたセル）。シートは読み直さない
  assert.ok(!g.mtgIndexEnsureRows_([{date: y + '/11/20', kind: 'RX Business Meeting'}]).error);
  const stored = {file: {getUrl: () => 'https://drive.google.com/file/d/PDF00001/view', getName: () => '20261023_新FMT_RX Business Meeting用資料.pdf',
    getId: () => 'PDF00001'}};
  g.mtgWriteIndex_('新FMT', y + '/10/23', stored);
  assert.deepStrictEqual(dates(), [[y + '/10/09', true], [y + '/10/23', true], [y + '/11/06', false], [y + '/11/20', false]]);
  // 読み直しても同じ（控えとシートが合っている）
  g.refreshMeetingIndexCache();
  assert.deepStrictEqual(dates(), [[y + '/10/09', true], [y + '/10/23', true], [y + '/11/06', false], [y + '/11/20', false]]);
});

test('「作成する」の途中・結果を控え、画面が読み込み直されても待って結果を出す（同じ人は作っている途中にもう一度作らない。止まったもの・古い結果は出さない）', () => {
  const {gas, g} = setup();
  const y = g.mtgParseDate_('1/1', new Date()).slice(0, 4);
  gas.addSheet('INDEX_会議資料', [['日付', '会議名', '全サブインダストリー', '新FMT'], [y + '/10/23', 'RX Business Meeting', '', '']]);
  const key = gas.get('MSCH_OPTIONS').runStateKey;
  const props = g.PropertiesService.getUserProperties();
  assert.strictEqual(g.getMeetingRunData('business', false).run, null, '作っていなければ何も出さない');

  // 作り終えたら結果を控え、画面を開き直すと出す（ほかの種類の画面には出さない）
  const result = g.runMeetingMaterials('business', false, y + '/10/23', ['新FMT'], false, '');
  const run = plain(g.getMeetingRunData('business', false).run);
  assert.deepStrictEqual([run.status, run.final, run.meeting, run.message], ['done', false, y + '/10/23', result.message]);
  assert.match(run.started, /^\d{1,2}:\d{2}$/);
  assert.strictEqual(g.getMeetingRunState('service'), null);

  // 作っている途中（ほかの画面で押した・画面が読み込み直された）：もう一度押しても作らない。画面は途中の知らせを受け取って待つ
  const now = Date.now();
  props.setProperty(key, JSON.stringify({status: 'running', kind: 'business', final: false, meeting: y + '/10/23', at: now - 60000}));
  assert.strictEqual(g.getMeetingRunData('business', false).run.status, 'running');
  assert.throws(() => g.runMeetingMaterials('business', false, y + '/10/23', ['新FMT'], false, ''), /から資料を作っている途中です。終わるまでお待ちください/);
  assert.strictEqual(JSON.parse(props.getProperty(key)).at, now - 60000, '途中の控えはそのまま');

  // 途中のまま7分を過ぎたら、止まったとみなす（また作れる）。済んだ手順と時間も出す
  props.setProperty(key, JSON.stringify({status: 'running', kind: 'business', final: false, meeting: y + '/10/23', at: now - 8 * 60000}));
  const timer = g.mschTimer_('資料の中身を作る（新FMT）', true);
  timer.lap('変更の記録（3件）');
  timer.lap('「新FMT」のシートを読んでまとめる');
  const stopped = g.getMeetingRunState('business');
  assert.strictEqual(stopped.status, 'error');
  assert.match(stopped.message, /途中で止まったようです/);
  assert.match(stopped.message, /止まる前に済んだ手順と時間：\n・資料の中身を作る（新FMT）：変更の記録（3件） [\d.]+秒・「新FMT」のシートを読んでまとめる [\d.]+秒/);
  assert.ok(!g.runMeetingMaterials('business', false, y + '/10/23', ['新FMT'], false, '').error);

  // 作れなかったときも控え、エラーはそのまま画面に返す
  const real = g.mtgRunStart_;
  g.mtgRunStart_ = () => { throw new Error('テストの失敗'); };
  try {
    assert.throws(() => g.runMeetingMaterials('business', false, y + '/10/23', ['新FMT'], false, ''), /テストの失敗/);
  } finally {
    g.mtgRunStart_ = real;
  }
  const failed = g.getMeetingRunState('business');
  assert.strictEqual(failed.status, 'error');
  assert.match(failed.message, /^テストの失敗/);

  // 終えてから30分を過ぎた結果は出さない
  props.setProperty(key, JSON.stringify({status: 'done', kind: 'business', final: false, meeting: y + '/10/23', at: now - 40 * 60000, finishedAt: now - 31 * 60000, message: '済み'}));
  assert.strictEqual(g.getMeetingRunState('business'), null);
});

test('diagnoseMeetingRead：コピーを作らずに読む時間を、Sheets API と SpreadsheetApp で比べて出す（読むだけ。Sheets API が無ければ ✗ と理由）', () => {
  const {gas, g} = setup();
  gas.addSheet('変更履歴_差分', [['記録日時', 'イベントID', '編集者メールアドレス', 'シート', 'セル', '列名', '種類', '変更前', '変更後']]);
  const writes = gas.writes.length;
  g.diagnoseMeetingRead();
  const message = gas.alerts[gas.alerts.length - 1].message;
  assert.match(message, /① Sheets API：✗ Sheets API（拡張サービス）が使えません/);
  assert.match(message, /② SpreadsheetApp：「新FMT」\d+行 [\d.]+秒・「変更履歴_差分」1行 [\d.]+秒（合計 [\d.]+秒）/);
  assert.match(message, /③ 値だけ（CSV）：/);
  assert.strictEqual(gas.writes.length, writes, '読むだけ');
});

test('diagnoseSpreadsheetLoad：計算を重くしている数式（ほかのファイルを読む・編集のたびに計算し直す・範囲の広い計算・列全体の参照・記録のシートを参照）をシートごとに数える（読むだけ）', () => {
  const {gas, g} = setup();
  const sheet = gas.addSheet('集計', [['a', 'b', 'c']], {rows: 10, columns: 5});
  gas.asUser(() => {
    sheet.getRange('A2').setFormula('=IMPORTRANGE("abc","x!A1")');
    sheet.getRange('A3').setFormula('=IMPORTRANGE("abc","x!A2")');
    sheet.getRange('B2').setFormula('=TODAY()');
    sheet.getRange('C2').setFormula('=COUNTIF(D:D,"x")');
    sheet.getRange('C3').setFormula("=COUNTA('変更履歴_差分'!A2:A100)");
  });
  const writes = gas.writes.length;
  g.diagnoseSpreadsheetLoad();
  const message = gas.alerts[gas.alerts.length - 1].message;
  assert.match(message, /調べたファイル：このスプレッドシート（時点コピーが無いため/);
  assert.match(message, /・集計：数式 5（形は 4種類）/);
  assert.match(message, /〔ほかのファイルを読む（コピーでは 0・エラーになる）〕\n・IMPORTRANGE：2件（集計 2）　例 集計!A2/);
  assert.match(message, /〔編集のたびに計算し直す〕\n・TODAY：1件（集計 1）　例 集計!B2/);
  assert.match(message, /・COUNTIF：1件（集計 1）　例 集計!C2/);
  assert.match(message, /・列全体の参照（A:A など）：1件（集計 1）　例 集計!C2/);
  assert.match(message, /・記録のシートを参照：1件（集計 1）　例 集計!C3/);
  assert.strictEqual(gas.writes.length, writes, '読むだけ');
});
