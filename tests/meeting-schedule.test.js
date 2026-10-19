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
