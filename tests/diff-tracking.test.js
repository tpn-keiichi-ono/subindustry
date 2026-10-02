'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');
const {COL, WIDTH, headerRow, dataRow, addTrackedSheets, setupProject} = require('./lib/fixture');

const LOG = '変更履歴_差分';
const SNAP = '変更時点スナップショット';

/* ---------------- 文字の差分・行の突き合わせ ---------------- */

test('diffCharacters_：増えた文字・消えた文字を取り出す', () => {
  const g = createGas().global;
  const result = plain(g.diffCharacters_('東京本社', '大阪本社', {remaining: 1e6}));
  assert.deepStrictEqual(result.parts, [
    {kind: 'del', text: '東京'},
    {kind: 'add', text: '大阪'},
    {kind: 'same', text: '本社'}
  ]);
  assert.strictEqual(result.coarse, false);
});

test('diffCharacters_：上限を超えるとまとめて「削除＋追加」にする', () => {
  const g = createGas().global;
  const result = plain(g.diffCharacters_('abc', 'xyz', {remaining: 1}));
  assert.strictEqual(result.coarse, true);
  assert.deepStrictEqual(result.parts, [{kind: 'del', text: 'abc'}, {kind: 'add', text: 'xyz'}]);
});

test('diffAlignRows_：削除された行と追加された行を見つける', () => {
  const g = createGas().global;
  assert.deepStrictEqual(plain(g.diffAlignRows_(['a', 'b', 'c'], ['a', 'c'])), [{old: [1], new: []}]);
  assert.deepStrictEqual(plain(g.diffAlignRows_(['a', 'c'], ['a', 'b', 'c'])), [{old: [], new: [1]}]);
  assert.deepStrictEqual(plain(g.diffAlignRows_(['a', 'b'], ['a', 'b'])), []);
});

/* ---------------- セットアップ ---------------- */

test('setupDiffTracking：比較基準を作り、トリガーを1つずつだけ設置する', () => {
  const {gas, g} = setupProject({launchers: false});
  g.setupDiffTracking();   // 2回目
  const handlers = gas.triggers.map(t => t.getHandlerFunction()).sort();
  assert.deepStrictEqual(handlers, ['catchUpDiffTracking', 'pauseDiffOnStructureChange', 'recordDiffEdit']);
  const state = JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'));
  assert.strictEqual(Object.keys(state).length, 2);
  Object.values(state).forEach(s => assert.strictEqual(s.paused, false));
  const baseline = gas.ss.getSheets().filter(s => s.getName().startsWith('__CHAR_DIFF_'));
  assert.strictEqual(baseline.length, 2);
  baseline.forEach(s => assert.ok(s.isSheetHidden()));
});

/* ---------------- 編集の記録 ---------------- */

test('recordDiffEdit：変更したセルを差分・編集者つきで記録し、最終更新日時を入れる', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注', {user: 'editor@example.com'}));

  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  const rec = log[0];
  assert.match(rec['イベントID'], /^EVT-\d{8}-\d{6}-[0-9A-F]{8}$/);
  assert.strictEqual(rec['編集者メールアドレス'], 'editor@example.com');
  assert.strictEqual(rec['シート'], '新FMT');
  assert.strictEqual(rec['セル'], 'F3');
  assert.strictEqual(rec['列名'], '状況');
  assert.strictEqual(rec['種類'], '文字列 -> 文字列');
  assert.strictEqual(rec['変更前'], '提案中');
  assert.strictEqual(rec['変更後'], '受注');

  // 最終更新日時（AG列）
  assert.ok(sheet.getRange(3, COL.stamp).getValue() instanceof gas.Date);

  // 行全体のスナップショット
  const snap = gas.records(SNAP);
  assert.strictEqual(snap.length, 1);
  assert.strictEqual(snap[0]['イベントID'], rec['イベントID']);
  assert.strictEqual(snap[0]['行番号'], '3');
  assert.strictEqual(snap[0]['変更セル'], 'F3');
  assert.strictEqual(snap[0]['得意先'], 'A社');
  assert.strictEqual(snap[0]['状況'], '受注');
});

test('recordDiffEdit：値が変わっていなければ何も記録しない', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '提案中'));
  assert.deepStrictEqual(gas.records(LOG), []);
});

test('recordDiffEdit：「=」で始まる入力も、差分ログでは数式にならない', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'F3', '=IMPORTXML("https://example.com","//a")'));
  const log = gas.ss.getSheetByName(LOG);
  const formulas = log.getRange(2, 1, 1, 11).getFormulas()[0];
  assert.ok(formulas.every(f => f === ''), '差分ログに数式が入っています: ' + formulas.join(' / '));
  assert.strictEqual(gas.records(LOG)[0]['変更後'], '=IMPORTXML("https://example.com","//a")');
  assert.strictEqual(gas.records(LOG)[0]['種類'], '文字列 -> 数式');
});

test('recordDiffEdit：書き込みはすべてドキュメントロックの中で行う', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'E4', '高橋'));
  assert.ok(gas.writes.length > 0);
  gas.writes.forEach(w => assert.ok(w.locked, 'ロックの外で書き込み: ' + w.sheet + '!' + w.a1));
});

test('recordDiffEdit：ボタン列（チェックボックス）の操作は記録しない', () => {
  const {gas, g, sheet} = setupProject();
  const rule = gas.get('DIFF_RULES')['新FMT'];
  ['AH3', 'AI3'].forEach(a1 => {
    const e = gas.edit(sheet, a1, true);
    assert.strictEqual(g.credIsLauncherEdit_(e, rule), true, a1);
    g.recordDiffEdit(e);
  });
  assert.strictEqual(g.credIsLauncherEdit_(gas.edit(sheet, 'F3', '受注'), rule), false);
  assert.deepStrictEqual(gas.records(LOG), []);
});

test('recordDiffEdit：ロックが取れないときは取りこぼしとして印を付け、回収で記録する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.lockBusy = true;
  g.recordDiffEdit(gas.edit(sheet, 'F3', '失注'));
  assert.deepStrictEqual(gas.records(LOG), []);
  assert.ok(gas.props.document.get('CHAR_DIFF_DIRTY_V1_' + sheet.getSheetId()));

  gas.lockBusy = false;
  g.catchUpDiffTracking();
  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  assert.match(log[0]['イベントID'], /^REC-/);
  assert.strictEqual(log[0]['編集者メールアドレス'], '不明（取りこぼし回収）');
  assert.strictEqual(log[0]['種類'], '文字列 -> 文字列 / 取りこぼし回収');
  assert.strictEqual(gas.props.document.get('CHAR_DIFF_DIRTY_V1_' + sheet.getSheetId()), undefined);
});

test('catchUpDiffTracking：印が無ければ何もしない', () => {
  const {gas, g} = setupProject({launchers: false});
  g.catchUpDiffTracking();
  assert.deepStrictEqual(gas.records(LOG), []);
  assert.strictEqual(gas.writes.length, 0);
});

/* ---------------- 行・列の追加・削除 ---------------- */

test('行の追加：追加された行を「行追加」として記録し、追跡を続ける', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => {
    sheet.insertRowsAfter(3, 1);
    sheet.getRange(4, COL.customer, 1, 3).setValues([['Ｄ社', '伊藤', '新規']]);
  });
  g.pauseDiffOnStructureChange(gas.change('INSERT_ROW', {user: 'editor@example.com'}));

  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  assert.match(log[0]['イベントID'], /^ROW-/);
  assert.strictEqual(log[0]['種類'], '行追加');
  assert.strictEqual(log[0]['セル'], '4行');
  assert.match(log[0]['変更後'], /得意先: Ｄ社/);

  // 比較基準が取り直され、次の編集もふつうに記録される
  g.recordDiffEdit(gas.edit(sheet, 'F5', '保留'));   // もとの B社 の行
  const next = gas.records(LOG)[1];
  assert.match(next['イベントID'], /^EVT-/);
  assert.strictEqual(next['セル'], 'F5');
  assert.strictEqual(next['変更前'], '受注');
});

test('行の削除：削除された行の内容を「行削除」として記録する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => sheet.deleteRow(4));   // B社
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW', {user: 'editor@example.com'}));

  const log = gas.records(LOG);
  assert.strictEqual(log.length, 1);
  assert.strictEqual(log[0]['種類'], '行削除');
  assert.strictEqual(log[0]['編集者メールアドレス'], 'editor@example.com');
  assert.strictEqual(log[0]['セル'], '4行（削除前の行番号）');
  assert.match(log[0]['変更前'], /得意先: B社/);
});

/* ---------------- 行のずれ（行数が変わらない場合・v6） ---------------- */

const brief = r => [r['イベントID'].split('-')[0], r['編集者メールアドレス'], r['セル'], r['種類'], r['変更前'], r['変更後']];
const CHECK = 'CHAR_DIFF_ROWCHECK_V1_';
const DIRTY = 'CHAR_DIFF_DIRTY_V1_';

test('行の削除と追加が変更トリガーより先に両方起きても（行数が同じ）、操作したシートを突き合わせ直して行削除を記録する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.select(sheet, 'A1');
  gas.asUser(() => { sheet.deleteRow(3); sheet.insertRowsAfter(4, 1); });   // A社 を削除、Ｃ社の下に空行
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW', {user: 'editor@example.com'}));
  g.pauseDiffOnStructureChange(gas.change('INSERT_ROW', {user: 'editor@example.com'}));

  assert.deepStrictEqual(gas.records(LOG).map(brief), [
    ['ROW', 'editor@example.com', '3行（削除前の行番号）', '行削除', '得意先: A社\n担当: 佐藤\n状況: 提案中', '（空欄）']
  ]);
  // 比較基準が取り直され、次の編集の「変更前」は同じ行（B社）の値
  g.recordDiffEdit(gas.edit(sheet, 'E3', '高橋', {user: 'b@example.com'}));
  assert.deepStrictEqual(brief(gas.records(LOG)[1]), ['EVT', 'b@example.com', 'E3', '文字列 -> 文字列', '鈴木', '高橋']);
});

test('行数が同じままずれた状態で編集が届いたら、突き合わせ直してから記録する（変更前は同じ行の値）', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.select(gas.ss.getSheetByName('新FMT2'), 'A1');   // 変更トリガーからは、操作したシートが分からなかった場合
  gas.asUser(() => { sheet.deleteRow(3); sheet.insertRowsAfter(4, 1); });
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));
  g.pauseDiffOnStructureChange(gas.change('INSERT_ROW'));
  assert.deepStrictEqual(gas.records(LOG), []);

  g.recordDiffEdit(gas.edit(sheet, 'E3', '高橋', {user: 'b@example.com'}));   // B社（3行目に繰り上がった）
  assert.deepStrictEqual(gas.records(LOG).map(brief), [
    ['ROW', '不明（行の追加・削除）', '3行（削除前の行番号）', '行削除', '得意先: A社\n担当: 佐藤\n状況: 提案中', '（空欄）'],
    ['EVT', 'b@example.com', 'E3', '文字列 -> 文字列', '鈴木', '高橋']
  ]);
  // 変更履歴：B社の行には B社の変更だけ
  const data = plain(g.getChangeHistoryData('新FMT', 3));
  assert.deepStrictEqual(data.events.map(ev => ev.changes.map(c => [c.column, c.before, c.after])), [[['担当', '鈴木', '高橋']]]);
});

test('削除の直後、変更トリガーより先に下の行の編集が届いても、その行の変更（EVT・編集者つき）として記録する', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => sheet.deleteRow(3));
  gas.asUser(() => sheet.getRange('F4').setValue('保留'));   // Ｃ社：トリガーがまだ動いていない別の人の編集
  g.recordDiffEdit(gas.edit(sheet, 'E3', '高橋', {user: 'b@example.com'}));   // B社
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));   // あとから届いた変更トリガー

  assert.deepStrictEqual(gas.records(LOG).map(brief), [
    ['ROW', '不明（行の追加・削除）', '3行（削除前の行番号）', '行削除', '得意先: A社\n担当: 佐藤\n状況: 提案中', '（空欄）'],
    ['EVT', 'b@example.com', 'E3', '文字列 -> 文字列', '鈴木', '高橋'],
    // 編集した範囲の外の変更は、編集した人のものにしない
    ['ROW', '不明（行の追加・削除）', 'F4', '空欄 -> 文字列 / 行の追加・削除時に回収', '（空欄）', '保留']
  ]);
  const snap = gas.records(SNAP);
  assert.deepStrictEqual(snap.map(r => [r['イベントID'].split('-')[0], r['編集者メールアドレス'], r['行番号'], r['得意先']]), [
    ['EVT', 'b@example.com', '3', 'B社'], ['ROW', '不明（行の追加・削除）', '4', 'Ｃ社']
  ]);
  // Ｃ社の変更履歴に、削除した A社の行の記録は出さない
  const data = plain(g.getChangeHistoryData('新FMT', 4));
  assert.deepStrictEqual(data.events.map(ev => ev.changes.map(c => [c.column, c.after])), [[['状況', '保留']]]);
});

test('取りこぼし回収：ずれた行を「変わった」と記録しない（突き合わせ直してから比べる）', () => {
  const rows = [['A社', '佐藤', '提案中'], ['B社', '鈴木', '受注'], ['Ｃ社', '田中', '見積'], ['Ｄ社', '伊藤', '失注']];
  const {gas, g, sheet} = setupProject({launchers: false, rows});
  gas.select(gas.ss.getSheetByName('新FMT2'), 'A1');
  gas.asUser(() => { sheet.deleteRow(3); sheet.insertRowsAfter(5, 1); });
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));
  g.pauseDiffOnStructureChange(gas.change('INSERT_ROW'));
  gas.lockBusy = true;
  g.recordDiffEdit(gas.edit(sheet, 'F5', '保留'));   // Ｄ社（5行目）。ロック待ちで取りこぼし
  gas.lockBusy = false;
  g.catchUpDiffTracking();

  assert.deepStrictEqual(gas.records(LOG).map(r => [r['セル'], r['種類'], r['変更前'], r['変更後']]), [
    ['3行（削除前の行番号）', '行削除', '得意先: A社\n担当: 佐藤\n状況: 提案中', '（空欄）'],
    ['F5', '文字列 -> 文字列 / 行の追加・削除時に回収', '失注', '保留']
  ]);
  assert.strictEqual(gas.props.document.get(DIRTY + sheet.getSheetId()), undefined);
});

test('並べ替え（行の入れ替え）だけなら何も記録せず、次の編集の変更前も正しい', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => {
    sheet.getRange(3, COL.customer, 2, 3).setValues([['B社', '鈴木', '受注'], ['A社', '佐藤', '提案中']]);
  });
  g.recordDiffEdit(gas.edit(sheet, 'F4', '受注', {user: 'a@example.com'}));   // A社（4行目に移った）
  assert.deepStrictEqual(gas.records(LOG).map(brief), [
    ['EVT', 'a@example.com', 'F4', '文字列 -> 文字列', '提案中', '受注']
  ]);
  // 比較基準も並べ替えたあとの順になっている
  g.recordDiffEdit(gas.edit(sheet, 'E3', '高橋'));   // B社
  assert.deepStrictEqual(gas.records(LOG)[1]['変更前'], '鈴木');
});

test('新しく入力した得意先・書き換えた得意先は、ずれとみなさない（ふつうの編集として記録する）', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  g.recordDiffEdit(gas.edit(sheet, 'D6:F6', [['Ｄ社', '伊藤', '新規']]));
  g.recordDiffEdit(gas.edit(sheet, 'D3', 'Ａ社'));
  gas.lockBusy = true;
  g.recordDiffEdit(gas.edit(sheet, 'D7', 'Ｅ社'));   // 取りこぼし
  gas.lockBusy = false;
  g.catchUpDiffTracking();
  assert.deepStrictEqual(gas.records(LOG).map(r => r['イベントID'].split('-')[0]), ['EVT', 'EVT', 'EVT', 'EVT', 'REC']);
});

/* ---------------- 追跡シートが11個ある場合 ---------------- */

/** 追跡シートを11個にする（新FMT・新FMT2 に、同じ形のシートを9個足す）。 */
function setupEleven() {
  const gas = createGas();
  addTrackedSheets(gas);
  const rules = gas.get('DIFF_RULES');
  for (let i = 3; i <= 11; i++) {
    const name = '新FMT' + i;
    rules[name] = {headerRow: 2, ranges: ['D3:AF']};
    const title = new Array(WIDTH).fill('');
    title[0] = name;
    gas.addSheet(name, [title, headerRow(), dataRow(1, 'X' + i + '社', '担当' + i, ''), dataRow(2, 'Y' + i + '社', '担当' + i, '')],
      {rows: 10, columns: WIDTH});
  }
  gas.global.setupDiffTracking();
  gas.writes.length = 0;
  gas.reads.length = 0;
  const sheetOf = name => gas.ss.getSheetByName(name);
  const baselineOf = name => '__CHAR_DIFF_' + sheetOf(name).getSheetId();
  /** 読んだシートの名前（記録用のシートを除く） */
  const readSheets = () => Array.from(new Set(gas.reads.map(r => r.sheet)))
    .filter(n => n !== LOG && n !== SNAP).sort();
  return {gas, g: gas.global, sheetOf, baselineOf, readSheets};
}

test('11シート：編集・取りこぼし回収で読むのは、そのシート（と比較基準）だけ', () => {
  const {gas, g, sheetOf, baselineOf, readSheets} = setupEleven();
  assert.strictEqual(Object.keys(JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'))).length, 11);

  g.recordDiffEdit(gas.edit(sheetOf('新FMT7'), 'F3', '受注'));
  assert.deepStrictEqual(readSheets(), [baselineOf('新FMT7'), '新FMT7'].sort());

  // 印が無ければ、回収はどのシートも読まない
  gas.reads.length = 0;
  g.catchUpDiffTracking();
  assert.deepStrictEqual(gas.reads, []);

  // 取りこぼした編集は、そのシートだけを比べる
  gas.lockBusy = true;
  g.recordDiffEdit(gas.edit(sheetOf('新FMT9'), 'F4', '保留'));
  gas.lockBusy = false;
  gas.reads.length = 0;
  g.catchUpDiffTracking();
  assert.deepStrictEqual(readSheets(), [baselineOf('新FMT9'), '新FMT9'].sort());
  // データのある最後の行（4行目）までしか読まない
  assert.ok(gas.reads.filter(r => r.sheet === '新FMT9').every(r => !/(\d+)$/.test(r.a1) || Number(r.a1.match(/(\d+)$/)[1]) <= 4),
    JSON.stringify(gas.reads.filter(r => r.sheet === '新FMT9').map(r => r.a1)));
  assert.deepStrictEqual(gas.records(LOG).map(r => [r['シート'], r['セル'], r['変更後']]), [['新FMT7', 'F3', '受注'], ['新FMT9', 'F4', '保留']]);
});

test('11シート：変更トリガーがロックを取れないときは、行数が変わったシートと操作したシートだけに「行の確かめ」の印を付ける', () => {
  const {gas, g, sheetOf, baselineOf, readSheets} = setupEleven();
  const ids = name => String(sheetOf(name).getSheetId());
  gas.select(sheetOf('新FMT8'), 'A1');                 // 操作したシート
  gas.asUser(() => sheetOf('新FMT5').deleteRow(3));     // 行数が変わったシート
  gas.lockBusy = true;
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));
  gas.lockBusy = false;

  const keys = Array.from(gas.props.document.keys());
  assert.deepStrictEqual(keys.filter(k => k.startsWith(CHECK)).sort(), [CHECK + ids('新FMT5'), CHECK + ids('新FMT8')].sort());
  assert.deepStrictEqual(keys.filter(k => k.startsWith(DIRTY)), [], '全セル比較の印は付けない');

  gas.reads.length = 0;
  g.catchUpDiffTracking();
  assert.deepStrictEqual(readSheets(), [baselineOf('新FMT5'), '新FMT5', baselineOf('新FMT8'), '新FMT8'].sort());
  // 新FMT8 は行数もずれも変わっていないので、見分け列だけを読んで終わる（全セルは読まない）
  const wide = gas.reads.filter(r => r.sheet === '新FMT8' && r.kind !== 'getDisplayValues');
  assert.ok(wide.every(r => /^D\d+:D\d+$/.test(r.a1)), JSON.stringify(wide.map(r => r.a1)));
  assert.deepStrictEqual(gas.records(LOG).map(r => [r['シート'], r['種類']]), [['新FMT5', '行削除']]);
  assert.deepStrictEqual(Array.from(gas.props.document.keys()).filter(k => k.startsWith(CHECK)), []);
});

test('11シート：行数が同じ行の追加・削除の通知では、操作したシートだけを確かめる', () => {
  const {gas, g, sheetOf, baselineOf, readSheets} = setupEleven();
  gas.select(sheetOf('新FMT6'), 'A1');
  gas.asUser(() => { sheetOf('新FMT6').deleteRow(3); sheetOf('新FMT6').insertRowsAfter(4, 1); });
  gas.reads.length = 0;
  g.pauseDiffOnStructureChange(gas.change('REMOVE_ROW'));
  assert.deepStrictEqual(readSheets(), [baselineOf('新FMT6'), '新FMT6'].sort());
  assert.deepStrictEqual(gas.records(LOG).map(r => [r['シート'], r['セル'], r['種類']]), [['新FMT6', '3行（削除前の行番号）', '行削除']]);
});

test('列の追加：追跡を停止して知らせる', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.asUser(() => sheet.insertColumnAfter(2));
  g.pauseDiffOnStructureChange(gas.change('INSERT_COLUMN'));
  const state = JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'));
  assert.strictEqual(state[sheet.getSheetId()].paused, true);
  assert.ok(gas.toasts.some(t => /列構成が変わったため差分追跡を停止しました/.test(t.message)));
});
