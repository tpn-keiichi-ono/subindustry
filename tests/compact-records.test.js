'use strict';

const {test, assert} = require('./lib/harness');
const {plain} = require('./lib/gas-mock');
const {COL, setupProject} = require('./lib/fixture');

const LOG = '変更履歴_差分';
const SNAP = '変更時点スナップショット';

test('記録の形：変更履歴_差分は9列、変更時点スナップショットは8列（得意先・行の内容）。使わない列・行は持たない', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注'));
  g.recordDiffEdit(gas.edit(sheet, 'E4', '高橋'));

  const log = gas.ss.getSheetByName(LOG);
  assert.deepStrictEqual(gas.dump(log, 'A1:I1')[0],
    ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', 'セル', '列名', '種類', '変更前', '変更後']);
  assert.strictEqual(log.getMaxColumns(), 9);
  assert.strictEqual(log.getMaxRows(), 3, '記録の行だけ');

  const snap = gas.ss.getSheetByName(SNAP);
  assert.deepStrictEqual(gas.dump(snap, 'A1:H1')[0],
    ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', '行番号', '変更セル', '得意先', '行の内容']);
  assert.strictEqual(snap.getMaxColumns(), 8);
  assert.strictEqual(snap.getMaxRows(), 3);
  assert.strictEqual(snap.getRange('H2').getFormula(), '', '数式にしない');

  // 変更履歴の画面は、行の内容から全項目を出す
  const data = plain(g.getChangeHistoryData('新FMT', 3));
  const snapshot = new Map(data.events[0].snapshot);
  assert.strictEqual(snapshot.get('状況'), '受注');
  assert.strictEqual(snapshot.get('得意先'), 'A社');
  assert.deepStrictEqual(data.events[0].changes.map(c => [c.column, c.before, c.after]), [['状況', '提案中', '受注']]);
  assert.ok(data.events[0].changes[0].parts.length > 0, '文字単位の差分はその場で作る');
});

test('比較基準（__CHAR_DIFF_*）は、追跡範囲の右端の列・追跡シートの行数まで', () => {
  const {gas, sheet} = setupProject({launchers: false});
  const baseline = gas.ss.getSheets().find(s => s.getName().indexOf('__CHAR_DIFF_') === 0);
  assert.strictEqual(baseline.getMaxColumns(), 32, 'D3:AF の右端（AF列）まで');
  assert.strictEqual(baseline.getMaxRows(), sheet.getMaxRows());
});

/** 前の版の記録を作る（差分ログは11列、スナップショットは見出しごとの列） */
function addLegacyRecords(gas) {
  [LOG, SNAP].forEach(name => { const old = gas.ss.getSheetByName(name); if (old) gas.ss.deleteSheet(old); });
  gas.addSheet(LOG, [
    ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', 'セル', '列名', '種類', '変更前', '変更後', '差分', '確認済み'],
    ['2026/09/01 10:00', 'EVT-20260901-100000-AAAA0001', 'old@example.com', '新FMT', 'E3', '担当', '文字列 -> 文字列', '山田', '佐藤', '佐藤', 'FALSE']
  ], {rows: 1000, columns: 26});
  gas.addSheet(SNAP, [
    ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', '行番号', '変更セル', 'No', 'メモ', '得意先', '担当', '状況', '変更履歴'],
    ['2026/09/01 10:00', 'EVT-20260901-100000-AAAA0001', 'old@example.com', '新FMT', '3', 'E3', '1', '', 'A社', '佐藤', '提案中', 'FALSE']
  ], {rows: 1000, columns: 26});
}

test('前の版の記録が残っていても、新しい記録は新しい形で足し、変更履歴は両方を読む', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  addLegacyRecords(gas);
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注'));

  const snap = gas.records(SNAP);
  assert.strictEqual(snap.length, 2);
  assert.strictEqual(snap[1]['得意先'], 'A社');
  assert.strictEqual(new Map(JSON.parse(snap[1]['行の内容'])).get('状況'), '受注');
  assert.strictEqual(snap[1]['状況'], '', '新しい行は見出しごとの列に書かない');

  const data = plain(g.getChangeHistoryData('新FMT', 3));
  assert.strictEqual(data.events.length, 2);
  assert.strictEqual(new Map(data.events[1].snapshot).get('担当'), '佐藤', '前の版の行は見出しごとの列から読む');
  assert.ok(!new Map(data.events[1].snapshot).has('変更履歴'), 'ボタン列は出さない');
  assert.strictEqual(new Map(data.events[0].snapshot).get('状況'), '受注');
});

test('compactDiffRecords：前の版の記録を、値を変えずに新しい形へ移し替え、使わない列・行を削除する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  addLegacyRecords(gas);
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注'));
  const historyBefore = plain(g.getChangeHistoryData('新FMT', 3));
  const sidebarBefore = plain(g.getSidebarBundle()).changes;
  const cellsBefore = g.diffCountCells_(gas.ss);

  gas.confirmAnswer = 'YES';
  g.compactDiffRecords();

  const log = gas.ss.getSheetByName(LOG);
  assert.strictEqual(log.getMaxColumns(), 9, '差分・確認済みの列を削除する');
  assert.strictEqual(log.getMaxRows(), 3);
  assert.deepStrictEqual(gas.records(LOG).map(r => r['変更後']), ['佐藤', '受注'], '記録の値は変えない');

  const snap = gas.ss.getSheetByName(SNAP);
  assert.deepStrictEqual(gas.dump(snap, 'A1:H1')[0],
    ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', '行番号', '変更セル', '得意先', '行の内容']);
  assert.strictEqual(snap.getMaxColumns(), 8);
  assert.strictEqual(snap.getMaxRows(), 3);
  const legacy = new Map(JSON.parse(gas.records(SNAP)[0]['行の内容']));
  assert.deepStrictEqual([...legacy.keys()], ['No', '得意先', '担当', '状況'], '見出しの順。空欄・ボタン列は持たない');

  assert.deepStrictEqual(plain(g.getChangeHistoryData('新FMT', 3)), historyBefore, '変更履歴の画面の中身は同じ');
  assert.deepStrictEqual(plain(g.getSidebarBundle()).changes, sidebarBefore, 'サイドバーの中身は同じ');
  assert.ok(g.diffCountCells_(gas.ss) < cellsBefore);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /1行を「行の内容」にまとめました。[\s\S]*セル数：/);

  // もう一度実行しても変わらない
  g.compactDiffRecords();
  assert.strictEqual(snap.getMaxColumns(), 8);
  assert.deepStrictEqual(plain(g.getChangeHistoryData('新FMT', 3)), historyBefore);
});

test('compactDiffRecords：「いいえ」なら何もしない', () => {
  const {gas, g} = setupProject({launchers: false});
  addLegacyRecords(gas);
  gas.writes.length = 0;
  gas.confirmAnswer = 'NO';
  g.compactDiffRecords();
  assert.deepStrictEqual(gas.writes, []);
  assert.strictEqual(gas.ss.getSheetByName(LOG).getMaxColumns(), 26);
});

test('diagnoseCellUsage：シートごとのセル数（空のセルも含む）を多い順に出す', () => {
  const {gas, g} = setupProject({launchers: false});
  g.diagnoseCellUsage();
  const lines = gas.alerts[0].message.split('\n');
  assert.match(lines[0], /^合計 [\d,]+ セル（上限 1,000万セルの \d+\.\d%）$/);
  assert.match(lines[2], /セル（\d+行 × \d+列）$/);
});
