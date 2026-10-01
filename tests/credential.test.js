'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');
const {COL, setupProject} = require('./lib/fixture');

const LOG = 'クレデンシャル履歴_記録';
const DIFF_LOG = '変更履歴_差分';

function values(overrides) {
  return Object.assign({
    plan: '', date: '2026/09/30', kind: 'クレデンシャル', title: '購買部 部長', person: '山田 太郎',
    name: 'ブランド刷新の実績紹介', note: '', owner: '佐藤 一郎', ownerDept: '第一営業部'
  }, overrides);
}

/** ダイアログを開いたときと同じ順で、対象行の情報を取ってから登録する。 */
function open(g, row) {
  return plain(g.getCredentialDialogData('新FMT', row || 3));
}

function request(data, extra) {
  return Object.assign({
    sheetName: data.sheetName, row: data.row, customer: data.customer, customerColumn: data.customerColumn
  }, extra);
}

/* ---------------- 入力の検証 ---------------- */

test('credValidate_：予定日か実施日のどちらかが必須', () => {
  const g = createGas().global;
  assert.throws(() => g.credValidate_(values({plan: '', date: ''}), null), /「予定日」か「実施日」のどちらか/);
  assert.strictEqual(g.credValidate_(values({plan: '2026-10-1', date: ''}), null).plan, '2026/10/01');
});

test('credValidate_：必須・文字数・選択肢・日付を確かめる', () => {
  const g = createGas().global;
  assert.throws(() => g.credValidate_(values({person: ''}), null), /「氏名」は必須です/);
  assert.throws(() => g.credValidate_(values({name: 'あ'.repeat(201)}), null), /200文字以内/);
  assert.throws(() => g.credValidate_(values({kind: 'その他'}), null), /選択肢が正しくありません/);
  assert.strictEqual(g.credValidate_(values({kind: '旧い種別'}), {kind: '旧い種別'}).kind, '旧い種別');
  assert.throws(() => g.credValidate_(values({date: '2026/02/30'}), null), /日付が正しくありません/);
  assert.strictEqual(g.credValidate_(values({name: '  改行\nを含む\t名称 '}), null).name, '改行 を含む 名称');
});

test('credNormalize_：全角・半角と空白の違いを無視する', () => {
  const g = createGas().global;
  assert.strictEqual(g.credNormalize_('Ｃ社 本社'), g.credNormalize_('C社本社'));
});

/* ---------------- セットアップ ---------------- */

test('setupCredentialLauncher：ボタン列にチェックボックスを入れ、トリガーは1つだけ', () => {
  const {gas, g, sheet} = setupProject();
  g.setupCredentialLauncher();
  g.setupChangeHistoryLauncher();
  assert.deepStrictEqual(sheet.getRange(3, COL.credButton, 3, 1).getValues(), [[false], [false], [false]]);
  assert.deepStrictEqual(sheet.getRange(3, COL.changeButton, 3, 1).getValues(), [[false], [false], [false]]);
  const launchers = gas.triggers.filter(t => t.getHandlerFunction() === 'onCredentialLauncherEdit');
  assert.strictEqual(launchers.length, 1);
});

test('setupCredentialLauncher：ボタン列が追跡範囲に含まれていれば止める', () => {
  const gas = createGas();
  require('./lib/fixture').addTrackedSheets(gas);
  const rules = gas.get('DIFF_RULES');
  rules['新FMT'].ranges = ['D3:AI'];   // ボタン列（AH・AI）まで含めてしまった場合
  rules['新FMT2'].ranges = ['D3:AI'];
  gas.global.setupDiffTracking();
  assert.throws(() => gas.global.setupCredentialLauncher(), /DIFF_RULES の ranges に含まれています/);
});

/* ---------------- 登録・編集・削除 ---------------- */

test('登録：記録シートに文字列のまま書き込み（数式・日付にしない）、変更履歴_差分にも残す', () => {
  const {gas, g} = setupProject();
  const data = open(g);
  assert.strictEqual(data.customer, 'A社');
  const result = plain(g.saveCredentialEntry(request(data, {values: values({name: '=SUM(1,2)'}), linkId: ''})));
  assert.strictEqual(result.message, '');
  assert.ok(result.entry.id);

  const log = gas.ss.getSheetByName(LOG);
  const headers = log.getRange(1, 1, 1, log.getLastColumn()).getDisplayValues()[0];
  const row = log.getRange(2, 1, 1, headers.length);
  assert.ok(row.getFormulas()[0].every(f => f === ''), '記録シートに数式が入っています');
  const byHeader = h => row.getValues()[0][headers.indexOf(h)];
  assert.strictEqual(byHeader('名称'), '=SUM(1,2)');
  assert.strictEqual(byHeader('実施日'), '2026/09/30', '日付が日時の値に変わっています');
  assert.strictEqual(byHeader('得意先'), 'A社');
  assert.strictEqual(byHeader('登録者'), 'user@example.com');

  const audit = gas.records(DIFF_LOG);
  assert.strictEqual(audit.length, 1);
  assert.match(audit[0]['イベントID'], /^CRD-/);
  assert.strictEqual(audit[0]['種類'], '履歴登録');
  assert.strictEqual(audit[0]['シート'], LOG);

  const again = open(g);
  assert.strictEqual(again.entries.length, 1);
  assert.strictEqual(again.entries[0].values.name, '=SUM(1,2)');
});

test('登録：書き込みはすべてドキュメントロックの中で行う', () => {
  const {gas, g} = setupProject();
  const data = open(g);
  gas.writes.length = 0;
  g.saveCredentialEntry(request(data, {values: values(), linkId: ''}));
  assert.ok(gas.writes.length > 0);
  gas.writes.forEach(w => assert.ok(w.locked, 'ロックの外で書き込み: ' + w.sheet + '!' + w.a1));
});

test('登録：ロックが取れないときはエラーにして書き込まない', () => {
  const {gas, g} = setupProject();
  const data = open(g);
  gas.lockBusy = true;
  assert.throws(() => g.saveCredentialEntry(request(data, {values: values(), linkId: ''})), /他の処理が実行中です/);
  gas.lockBusy = false;
  assert.strictEqual(open(g).entries.length, 0);
});

test('登録：開いたあとで行の得意先が変わっていたら止める', () => {
  const {gas, g, sheet} = setupProject();
  const data = open(g);
  gas.edit(sheet, 'D3', 'Z社');
  assert.throws(() => g.saveCredentialEntry(request(data, {values: values(), linkId: ''})), /得意先が変わっています/);
});

test('編集：変わった項目だけを書き換えて記録し、他の人の先の変更は上書きしない', () => {
  const {gas, g} = setupProject();
  const data = open(g);
  const saved = plain(g.saveCredentialEntry(request(data, {values: values(), linkId: ''}))).entry;

  const updated = plain(g.updateCredentialEntry(request(data, {
    id: saved.id, original: saved.values, originalLinkId: '', values: values({person: '山田 花子'}), linkId: ''
  })));
  assert.strictEqual(updated.entry.values.person, '山田 花子');
  assert.strictEqual(updated.entry.updatedBy, 'user@example.com');
  const audit = gas.records(DIFF_LOG).filter(r => r['種類'] === '履歴編集');
  assert.strictEqual(audit.length, 1);
  assert.strictEqual(audit[0]['列名'], '氏名（A社）');
  assert.strictEqual(audit[0]['変更前'], '山田 太郎');
  assert.strictEqual(audit[0]['変更後'], '山田 花子');

  // 古い内容（original）のまま送ってきた別の人の編集は止める
  assert.throws(() => g.updateCredentialEntry(request(data, {
    id: saved.id, original: saved.values, originalLinkId: '', values: values({person: '別の人'}), linkId: ''
  })), /他の人が先に変更しています/);
});

test('削除（soft）：削除者・削除日時を残し、一覧から消える', () => {
  const {gas, g} = setupProject();
  const data = open(g);
  const saved = plain(g.saveCredentialEntry(request(data, {values: values(), linkId: ''}))).entry;
  g.deleteCredentialEntry(request(data, {id: saved.id, original: saved.values, originalLinkId: ''}));

  assert.strictEqual(open(g).entries.length, 0);
  const rec = gas.records(LOG)[0];
  assert.strictEqual(rec['削除者'], 'user@example.com');
  assert.ok(rec['削除日時']);
  assert.ok(gas.records(DIFF_LOG).some(r => r['種類'] === '履歴削除'));
});

test('紐づけ：オファリングをクレデンシャルに紐づけ、子があるクレデンシャルは種別を変えられない', () => {
  const {g} = setupProject();
  const data = open(g);
  const parent = plain(g.saveCredentialEntry(request(data, {values: values(), linkId: ''}))).entry;
  const child = plain(g.saveCredentialEntry(request(data, {
    values: values({kind: 'オファリング', name: '新ブランドのご提案', date: '', plan: '2026/11/01'}), linkId: parent.id
  }))).entry;
  assert.strictEqual(child.linkId, parent.id);

  assert.throws(() => g.updateCredentialEntry(request(data, {
    id: parent.id, original: parent.values, originalLinkId: '', values: values({kind: 'オファリング'}), linkId: ''
  })), /紐づいているオファリングがあるため、種別を変更できません/);

  // クレデンシャル以外には紐づけられない
  assert.throws(() => g.saveCredentialEntry(request(data, {
    values: values({kind: 'オファリング'}), linkId: child.id
  })), /種別が「クレデンシャル」ではありません/);
});

test('別の得意先の行から開いた場合は、その得意先の履歴だけを返す', () => {
  const {g} = setupProject();
  g.saveCredentialEntry(request(open(g, 3), {values: values(), linkId: ''}));
  assert.strictEqual(open(g, 3).entries.length, 1);
  assert.strictEqual(open(g, 4).entries.length, 0);
});

/* ---------------- ボタン（インストール型 onEdit） ---------------- */

test('onCredentialLauncherEdit：ボタン列をチェックするとモーダルを開き、チェックを戻す', () => {
  const {gas, g, sheet} = setupProject();
  g.onCredentialLauncherEdit(gas.edit(sheet, 'AH4', true));
  g.onCredentialLauncherEdit(gas.edit(sheet, 'AI5', true));
  assert.deepStrictEqual(gas.dialogs.map(d => d.html.file), ['CredentialDialog', 'ChangeHistoryDialog']);
  assert.deepStrictEqual(gas.dialogs.map(d => d.html.data.row), [4, 5]);
  assert.strictEqual(sheet.getRange('AH4').getValue(), false);
  assert.strictEqual(sheet.getRange('AI5').getValue(), false);
});

test('onCredentialLauncherEdit：ボタン列以外の編集・チェックを外す操作では何もしない', () => {
  const {gas, g, sheet} = setupProject();
  g.onCredentialLauncherEdit(gas.edit(sheet, 'F3', 'TRUE'));
  g.onCredentialLauncherEdit(gas.edit(sheet, 'AH3', false));
  assert.strictEqual(gas.dialogs.length, 0);
});
