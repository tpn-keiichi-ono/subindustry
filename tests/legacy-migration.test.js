'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/*
 * 旧フォーマットからの転記（LegacyMigration.gs）
 * 新：本物と同じ見出し（2行目。A〜AC）。B（No.）と O（期待値）は3行目に数式。
 * 旧：「案件_{シート名}」の1〜2行目が見出し、3行目が入力の種類、4行目からデータ。
 */
const NEW_HEADERS = ['最終更新日時\n（自動入力）', 'No.', 'サブインダストリー', '得意先', '先方部門', '先方担当役職', '先方担当氏名',
  'クライアントとのリレーション\nクライアント内のガバナンス', '案件名', '案件のスコープ', 'スコープに対する競合他社', 'プロジェクト開始-終了', '案件ステータス',
  '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月', '売上開始月', '提案開始日', '活動状況', '活動における課題', '次回アクション',
  '商談記録（ＵＲＬ）', 'アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当', '想定される価値創出のケース／期待される導入効果',
  '提案を勝ち取るための戦略・差異化要素', 'サービスのリクエスト'];
const WIDTH = 32;   // AF まで
const col = header => NEW_HEADERS.indexOf(header) + 1;

const OLD_HEADERS = ['No.', 'サブインダストリー', '得意先', '案件名', '提案シナリオ', '先方部門', '先方担当役職', '先方担当氏名', '営業担当', '企画担当',
  '提案開始日', 'ステータス', '活動状況', '次回アクション', '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）',
  '配分', '売上規模', '配分', '売上規模', '配分', '売上規模', '配分', '売上規模', '受注月', '売上開始月', 'サービス配分\nチェック', '', '備考'];
const OLD_GROUPS = {16: 'マーケ', 18: 'コマース', 20: 'カスタマー', 22: 'その他'};   // 1行目（0 から数えた列）

function setup(oldRows, options) {
  const o = options || {};
  const gas = createGas();
  const g = gas.global;
  const pad = row => row.concat(new Array(WIDTH - row.length).fill(''));
  gas.addSheet('得意先マスタ', [['No', '得意先'], ['1', 'A社'], ['2', 'B社'], ['3', 'C社']], {rows: 10, columns: 2});
  gas.addSheet('定義_ステータス', [['ステータス'], ['C-提案中(50%)'], ['B-クロージング中(70%)'], ['A-受注済み(100%)']], {rows: 10, columns: 1});
  ['新FMT', '新FMT2'].forEach(name => {
    const sheet = gas.addSheet(name, [pad(['', '', '機会詳細']), pad(NEW_HEADERS),
      pad(['', '', name, 'B社', '前の部門', '', '', '', '前の案件', '', '', '', 'C-提案中(50%)', '10'])], {rows: o.rows || 12, columns: WIDTH});
    gas.asUser(() => {
      sheet.getRange(3, col('No.')).setFormula('=ROW(B3:B12)-2');
      sheet.getRange(3, col('期待値調整済\n想定売上規模（百万）')).setFormula('=IF(M3:M12="",,N3:N12*0.5)');
    });
    const V = g.SpreadsheetApp.newDataValidation;
    const last = o.rows || 12;
    sheet.getRange('C3:C' + last).setDataValidation(V().requireValueInList([name], true).build());
    sheet.getRange('D3:D' + last).setDataValidation(V().requireValueInRange(gas.ss.getSheetByName('得意先マスタ').getRange('B2:B10'), true).build());
    sheet.getRange('M3:M' + last).setDataValidation(V().requireValueInRange(gas.ss.getSheetByName('定義_ステータス').getRange('A2:A10'), true).build());
  });
  g.setupDiffTracking();

  const book = g.SpreadsheetApp.create('旧フォーマット');
  book.id = '1OldFormatSpreadsheet0000000000000000';
  const old = book.getSheets()[0].setName('案件_新FMT');
  old.insertColumnsAfter(old.getMaxColumns(), OLD_HEADERS.length - old.getMaxColumns());
  const group = new Array(OLD_HEADERS.length).fill('');
  Object.keys(OLD_GROUPS).forEach(i => { group[i] = OLD_GROUPS[i]; });
  const kinds = OLD_HEADERS.map(() => '手動入力');
  gas.asUser(() => old.getRange(1, 1, 3 + oldRows.length, OLD_HEADERS.length).setValues([group, OLD_HEADERS, kinds].concat(oldRows)));
  gas.writes.length = 0;
  gas.alerts.length = 0;
  return {gas, g, book};
}

/** 旧の1行（見出しの名前 → 値） */
function oldRow(values) {
  return OLD_HEADERS.map((h, i) => {
    if (h === '配分' || h === '売上規模') return '';
    return values[h] != null ? values[h] : '';
  });
}

const URL = 'https://docs.google.com/spreadsheets/d/1OldFormatSpreadsheet0000000000000000/edit#gid=0';

test('migrateLegacyDeals：旧の「案件_{シート名}」の案件を、見出しで対応する列に転記する（今の行は上書き。数式の列・対応しない列は書かない）', () => {
  const date = new Date(Date.UTC(2026, 3, 15));
  // 活動状況は文字の「=初回訪問」（先頭の ' は旧のシートで文字として入力した印）
  const {gas, g} = setup([
    oldRow({'No.': '1', 'サブインダストリー': '食品', '得意先': 'A社', '案件名': '販促DX', '提案シナリオ': 'シナリオ1', '先方部門': '営業企画部',
      '先方担当役職': '部長', '先方担当氏名': '山田', '営業担当': '後藤', '企画担当': '横井', '提案開始日': date, 'ステータス': 'C-提案中(50%)',
      '活動状況': "'=初回訪問", '次回アクション': '見積提示', '想定売上規模（百万）': 10, '期待値調整済\n想定売上規模（百万）': 5,
      '受注月': date, '売上開始月': date, '備考': '備考の文'}),
    oldRow({'得意先': '', '案件名': ''}),   // 空の行は飛ばす
    oldRow({'得意先': 'Z社', '案件名': '新規開拓', 'ステータス': '失注', '想定売上規模（百万）': 3})   // 選択肢に無い得意先・ステータス
  ]);
  gas.promptAnswers = [URL];
  g.migrateLegacyDeals();

  assert.match(gas.alerts[0].message, /「旧フォーマット」から、次のとおり転記します。[\s\S]*・新FMT：2件（選択肢に無い値 2セルは空欄にします）\n・新FMT2：旧のスプレッドシートに「案件_新FMT2」シートがありません（転記しません）/);
  const sheet = gas.ss.getSheetByName('新FMT');
  const at = (row, header) => sheet.getRange(row, col(header)).getValue();
  assert.deepStrictEqual(['サブインダストリー', '得意先', '先方部門', '先方担当役職', '先方担当氏名', '案件名', '案件ステータス', '想定売上規模（百万）',
    '活動状況', '次回アクション', 'アカウント責任者'].map(h => at(3, h)),
  ['新FMT', 'A社', '営業企画部', '部長', '山田', '販促DX', 'C-提案中(50%)', 10, '=初回訪問', '見積提示', '後藤']);
  assert.strictEqual(sheet.getRange(3, col('活動状況')).getFormula(), '', '「=」で始まる文字も数式にしない');
  assert.ok(['提案開始日', '受注月', '売上開始月'].every(h => Object.prototype.toString.call(at(3, h)) === '[object Date]'), '日付は日付のまま');
  assert.deepStrictEqual(['BX担当', '案件のスコープ', '活動における課題', 'サービスのリクエスト'].map(h => at(3, h)), ['', '', '', ''],
    '企画担当・提案シナリオ・備考は転記しない');
  assert.deepStrictEqual([sheet.getRange('B3').getFormula(), sheet.getRange('O3').getFormula()], ['=ROW(B3:B12)-2', '=IF(M3:M12="",,N3:N12*0.5)'],
    '数式の列（No.・期待値）は書かない・消さない');
  assert.deepStrictEqual(['得意先', '案件名', '案件ステータス', '想定売上規模（百万）'].map(h => at(4, h)), ['', '新規開拓', '', 3],
    '選択肢に無い値は空欄。ほかの列は書く');
  assert.deepStrictEqual(['得意先', '案件名', '先方部門'].map(h => at(5, h)), ['', '', ''], '前の行は上書き（消す）');
  assert.ok(gas.writes.filter(w => w.ss === gas.ss.getId() && w.sheet === '新FMT').every(w => w.locked), '書き込みはロックの中（時点コピーは別のファイル）');

  const done = gas.alerts[gas.alerts.length - 1].message;
  assert.match(done, /選択肢に無いため空欄にしたセル（2セル）：\n新FMT!D4（得意先）：Z社\n新FMT!M4（案件ステータス）：失注\n得意先マスタ・ステータスの定義に足してから、もう一度実行すると入ります/);
  // 転記は変更として記録しない（比較基準を取り直す）。会議サマリの基準を記録する
  const state = JSON.parse(gas.props.document.get('CHAR_DIFF_STATE_V1'));
  Object.values(state).forEach(s => assert.strictEqual(s.paused, false));
  const log = gas.ss.getSheetByName('変更履歴_差分');
  assert.ok(!log || log.getLastRow() <= 1, '変更履歴に転記の記録を残さない');
  g.recordDiffEdit(gas.edit(sheet, 'S3', '2回目の訪問'));
  assert.strictEqual(gas.ss.getSheetByName('変更履歴_差分').getLastRow(), 2, 'そのあとの編集は記録する');
  const base = JSON.parse(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT'));
  assert.ok(base.some(m => m.baseline && m.copy));
  assert.strictEqual(g.PropertiesService.getDocumentProperties().getProperty('MTG_MEETINGS_新FMT2'), null, '転記しなかったシートは基準を記録しない');
  assert.strictEqual(gas.ss.getSheetByName('新FMT2').getRange(3, col('案件名')).getValue(), '前の案件', '転記しなかったシートはそのまま');
});

test('migrateLegacyDeals：「いいえ」なら何も書かない。行が足りないシートは書かずに知らせる', () => {
  const rows = [1, 2, 3].map(n => oldRow({'得意先': 'A社', '案件名': '案件' + n}));
  let {gas, g} = setup(rows);
  gas.promptAnswers = [URL];
  gas.confirmAnswer = 'NO';
  g.migrateLegacyDeals();
  assert.deepStrictEqual(gas.writes.filter(w => w.sheet === '新FMT'), []);

  ({gas, g} = setup(rows, {rows: 4}));   // データの行は 3〜4 の2行だけ
  gas.promptAnswers = [URL];
  g.migrateLegacyDeals();
  assert.strictEqual(gas.alerts.length, 1);
  assert.match(gas.alerts[0].message, /転記できるシートがありません。[\s\S]*・新FMT：行が足りません（3件に対して 2行）。エディタで addDealRows\(\) を実行して行を足してから、もう一度実行してください/);
  assert.deepStrictEqual(gas.writes.filter(w => w.sheet === '新FMT'), []);
});

test('migrateLegacyDeals：URL でないとき・開けないときは止める', () => {
  const {gas, g} = setup([]);
  gas.promptAnswers = ['旧のファイル'];
  assert.throws(() => g.migrateLegacyDeals(), /スプレッドシートの URL として読めません/);
  gas.promptAnswers = ['https://docs.google.com/spreadsheets/d/1NotSharedSpreadsheet000000000000000/edit'];
  assert.throws(() => g.migrateLegacyDeals(), /旧フォーマットのスプレッドシートを開けません/);
  assert.ok(plain(gas.writes).every(w => w.sheet !== '新FMT'));
});
