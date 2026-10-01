/** 変更履歴モーダルのデータ（ChangeHistory.gs getChangeHistoryData）を確かめる。 */
const assert = require('assert');
const {makeEnvironment, loadProject, makeSheet} = require('./lib/gas-mock');

const env = makeEnvironment();
const api = loadProject(env, ['DiffTracking.gs', 'CredentialHistory.gs', 'ChangeHistory.gs'], ['getChangeHistoryData']);

const fromRows = (name, id, rows) => {
  const width = Math.max(...rows.map(r => r.length));
  const s = makeSheet(name, id, rows.length, width);
  rows.forEach((r, i) => r.forEach((v, j) => { s.data[i][j] = v; }));
  env.ss.sheets.push(s);
  return s;
};
fromRows('新FMT', 1, [[''], ['得意先', '案件名', '金額', 'クレデンシャル\nオファリング登録', '変更履歴', '最終更新日時'],
  ['株式会社サンプル商事', 'A案件', '100', 'FALSE', 'FALSE', '']]);
fromRows('変更時点スナップショット', 10, [
  ['記録日時', 'イベントID', '編集者メールアドレス', 'シート', '行番号', '変更セル', '得意先', '案件名', '金額', 'クレデンシャル オファリング登録', '変更履歴', '最終更新日時'],
  ['2026/09/28 10:00', 'EVT-1', 'a@x.com', '新FMT', '3', 'C3', '株式会社サンプル商事', 'A案件', '80', 'FALSE', 'FALSE', '2026/09/28 10:00'],
  ['2026/09/29 11:00', 'EVT-2', 'b@x.com', '新FMT', '3', 'B3, C3', '株式会社サンプル商事', 'A案件（改）', '100', 'FALSE', 'FALSE', '2026/09/29 11:00'],
  ['2026/09/29 12:00', 'EVT-3', 'c@x.com', '新FMT', '4', 'B4', '別会社', 'X', '1', '', '', ''],
  ['2026/09/30 09:00', 'REC-4', '不明', '新FMT', '3', 'C3', '株式会社サンプル商事', 'A案件（改）', '120', 'FALSE', 'FALSE', '2026/09/30 09:00']]);
fromRows('変更履歴_差分', 11, [
  ['記録日時', 'イベントID', '編集者', 'シート', 'セル', '列名', '種類', '変更前', '変更後', '差分', '確認済み'],
  ['t', 'EVT-1', 'a@x.com', '新FMT', 'C3', '金額', '数値 -> 数値', '（空欄）', '80', '', ''],
  ['t', 'EVT-2', 'b@x.com', '新FMT', 'B3', '案件名', '文字列 -> 文字列', 'A案件', 'A案件（改）', '', ''],
  ['t', 'EVT-2', 'b@x.com', '新FMT', 'C3', '金額', '数値 -> 数値', '80', '100', '', ''],
  ['t', 'EVT-2', 'b@x.com', '新FMT', 'B9', '案件名', 'x', 'q', 'w', '', '']]);

const result = api.getChangeHistoryData('新FMT', 3);
assert.deepStrictEqual(result.events.map(e => e.eventId), ['REC-4', 'EVT-2', 'EVT-1'], '同じ得意先の記録だけを新しい順に返す');

const evt2 = result.events[1];
assert.deepStrictEqual(evt2.changes.map(c => c.column), ['案件名', '金額'], '別の行（B9）の記録は含めない');
assert.strictEqual(result.events[2].changes[0].before, '', '「（空欄）」は空として扱う');

const rec = result.events[0];
assert.deepStrictEqual(rec.changes.map(c => [c.column, c.before, c.after]), [['金額', '100', '120']],
  '差分ログが無い記録は前のスナップショットと比べて補う（最終更新日時は無視）');
assert(!rec.snapshot.some(([k]) => k.includes('変更履歴')), 'ボタン列は全項目の表示から外す');

console.log('change-history: ok');
