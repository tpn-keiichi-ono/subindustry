/** 履歴サイドバーのデータ（HistorySidebar.gs）を確かめる。 */
const assert = require('assert');
const {makeEnvironment, loadProject, makeSheet} = require('./lib/gas-mock');

const env = makeEnvironment();
const api = loadProject(env, ['DiffTracking.gs', 'CredentialHistory.gs', 'ChangeHistory.gs', 'HistorySidebar.gs'],
  ['getSidebarBundle', 'getSidebarSelection']);

const fromRows = (name, id, rows) => {
  const width = Math.max(...rows.map(r => r.length));
  const s = makeSheet(name, id, rows.length, width);
  rows.forEach((r, i) => r.forEach((v, j) => { s.data[i][j] = v; }));
  env.ss.sheets.push(s);
  return s;
};
const source = fromRows('新FMT', 1, [[''], ['得意先', '案件名'], ['株式会社サンプル商事', 'A']]);
source.currentCell = source.getRange(3, 2);
env.ss.activeSheet = source;

fromRows('変更時点スナップショット', 10, [
  ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', '行番号', '変更セル', '得意先', '案件名'],
  ['2026/09/29 11:00', 'EVT-2', 'b@x.com', '新FMT', '3', 'B3', '株式会社サンプル商事', 'A案件（改）'],
  ['2026/09/30 09:00', 'EVT-4', 'c@x.com', '新FMT', '3', 'B3', '株式会社 サンプル商事', 'B']]);
fromRows('変更履歴_差分', 11, [
  ['記録日時', 'イベントID', '編集者', 'シート', 'セル', '列名', '種類', '変更前', '変更後', '差分', '確認済み'],
  ['t', 'EVT-2', 'b', '新FMT', 'B3', '案件名', 'x', 'A案件', 'A案件（改）', '', ''],
  ['t', 'EVT-4', 'c', '新FMT', 'B3', '案件名', 'x', 'A案件（改）', '（空欄）', '', '']]);
fromRows('クレデンシャル履歴_記録', 12, [
  ['ID', '登録日時', '得意先', '予定日', '実施日', '種別', '役職', '氏名', '名称', 'フリーコメント', '担当者名', '担当部署', '紐づくクレデンシャルID', '登録者', '登録元（登録時の行）', '更新者', '更新日時', '削除者', '削除日時'],
  ['c1', '2026/09/30 10:00', '株式会社サンプル商事', '', '2026/09/30', 'クレデンシャル', '執行役員', 'テスト花子', '実績集', '', '佐藤', '', '', 'a', '', '', '', '', ''],
  ['o1', '2026/09/30 10:05', '株式会社サンプル商事', '2026/10/20', '', 'オファリング', '', 'テスト太郎', '提案', '', '', '', 'c1', 'a', '', '', '', '', ''],
  ['x1', '2026/09/30 10:05', '株式会社サンプル商事', '', '2026/09/01', 'オファリング', '', '', '削除済み', '', '', '', '', 'a', '', '', '', 'b', '2026/09/30']]);

const first = api.getSidebarSelection({sheet: null, row: null});
assert.deepStrictEqual([first.sheet, first.row, first.customer], ['新FMT', 3, '株式会社サンプル商事'], '選択中の行と得意先を返す');
assert.strictEqual(api.getSidebarSelection({sheet: '新FMT', row: 3}).same, true, '同じ行なら得意先を読まずに same を返す');

const bundle = api.getSidebarBundle();
const credentials = bundle.credentials['株式会社サンプル商事'];
assert.deepStrictEqual(credentials.map(e => e.id), ['c1', 'o1'], '削除済みの履歴は含めない');
assert.strictEqual(credentials[0].values.owner, '佐藤', '担当者も返す');

const changes = bundle.changes['新FMT']['株式会社サンプル商事'];
assert.deepStrictEqual(changes.map(e => e.eventId), ['EVT-4', 'EVT-2'], '得意先名の空白の違いを無視してまとめ、新しい順に返す');
assert.strictEqual(changes[0].changes[0].after, '', '「（空欄）」は空として扱う');

console.log('sidebar: ok');
