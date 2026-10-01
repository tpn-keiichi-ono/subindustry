'use strict';

const {test, assert} = require('./lib/harness');
const {plain} = require('./lib/gas-mock');
const {setupProject} = require('./lib/fixture');

const CRED_LOG = 'クレデンシャル履歴_記録';
const DIFF_LOG = '変更履歴_差分';
const SNAP = '変更時点スナップショット';

test('addSampleHistory：選んだ行の得意先に、クレデンシャル・オファリング・変更の記録を5件ずつ追加する', () => {
  const {gas, g, sheet} = setupProject();
  const before = sheet.getRange('A3:AI5').getValues();
  gas.select(sheet, 'E3');
  g.addSampleHistory();

  // クレデンシャル 5件・オファリング 5件（オファリングはクレデンシャルに紐づく）
  const creds = plain(g.getCredentialDialogData('新FMT', 3)).entries;
  const parents = creds.filter(e => e.values.kind === 'クレデンシャル');
  const children = creds.filter(e => e.values.kind === 'オファリング');
  assert.strictEqual(parents.length, 5);
  assert.strictEqual(children.length, 5);
  const parentIds = new Set(parents.map(e => e.id));
  children.forEach(c => assert.ok(parentIds.has(c.linkId), '紐づけ先がサンプルのクレデンシャルではありません'));
  creds.forEach(e => {
    assert.ok(e.values.name.startsWith('【サンプル】'), e.values.name);
    assert.strictEqual(e.source, 'サンプルデータ');
    assert.ok(e.values.plan || e.values.date, '予定日・実施日のどちらも空です');
  });
  assert.ok(children.some(c => !c.values.date && c.values.plan), '予定のオファリングがありません');

  // 変更の記録 5件（変更履歴_差分 と 変更時点スナップショット）
  const log = gas.records(DIFF_LOG).filter(r => r['イベントID'].startsWith('SMP-'));
  assert.strictEqual(log.length, 5);
  log.forEach(r => {
    assert.strictEqual(r['シート'], '新FMT');
    assert.match(r['セル'], /^[A-Z]+3$/);
    assert.notStrictEqual(r['列名'], '得意先', '得意先の列は変えない');
    assert.strictEqual(r['編集者メールアドレス'], 'サンプルデータ');
  });
  assert.strictEqual(gas.records(SNAP).filter(r => r['イベントID'].startsWith('SMP-')).length, 5);
  // 監査の記録（CRD）は作らない
  assert.strictEqual(gas.records(DIFF_LOG).filter(r => r['イベントID'].startsWith('CRD-')).length, 0);

  // 追跡シートのセルは変わらない
  assert.deepStrictEqual(sheet.getRange('A3:AI5').getValues(), before);
  // 書き込みはロックの中
  gas.writes.forEach(w => assert.ok(w.locked, 'ロックの外で書き込み: ' + w.sheet + '!' + w.a1));
});

test('addSampleHistory：サイドバーにも表示され、変更の記録は日時の新しい順になる', () => {
  const {gas, g, sheet} = setupProject();
  g.recordDiffEdit(gas.edit(sheet, 'F3', '受注'));   // 今日の本物の編集（行としては先に記録される）
  gas.select(sheet, 'E3');
  g.addSampleHistory();                             // 過去の日時のサンプル（行としてはあとに足される）

  const bundle = plain(g.getSidebarBundle());
  assert.strictEqual(bundle.credentials['A社'].length, 10);
  const events = bundle.changes['新FMT']['A社'];
  assert.strictEqual(events.length, 6);
  assert.match(events[0].eventId, /^EVT-/, '今日の本物の編集がいちばん新しいはずです');
  const times = events.map(e => e.at);
  assert.deepStrictEqual(times, times.slice().sort().reverse());
  assert.strictEqual(new Set(events.map(e => e.at.slice(0, 10))).size, 6, '日付が分かれていません');

  // 変更履歴モーダルも同じ順
  const modal = plain(g.getChangeHistoryData('新FMT', 3));
  assert.deepStrictEqual(modal.events.map(e => e.at), times);
});

test('addSampleHistory：複数行を選ぶと得意先ごとに追加し、「いいえ」なら何もしない', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'E3:E5');
  gas.confirmAnswer = 'NO';
  g.addSampleHistory();
  assert.strictEqual(gas.records(CRED_LOG).length, 0);

  gas.confirmAnswer = 'YES';
  g.addSampleHistory();
  const customers = gas.records(CRED_LOG).map(r => r['得意先']);
  assert.deepStrictEqual([...new Set(customers)], ['A社', 'B社', 'Ｃ社']);
  assert.strictEqual(customers.length, 30);
});

test('addSampleHistory：見出し行を選んでいたら知らせて、何も追加しない', () => {
  const {gas, g, sheet} = setupProject();
  gas.select(sheet, 'E2');
  g.addSampleHistory();
  assert.match(gas.alerts[gas.alerts.length - 1].message, /データ行（3行目以降）を選んでください/);
  assert.strictEqual(gas.records(CRED_LOG).length, 0);
});

test('removeSampleHistory：サンプルの行だけを削除し、本物の履歴は残す', () => {
  const {gas, g, sheet} = setupProject();
  // 本物の履歴
  const data = plain(g.getCredentialDialogData('新FMT', 4));
  g.saveCredentialEntry({sheetName: '新FMT', row: 4, customer: data.customer, customerColumn: data.customerColumn, linkId: '',
    values: {plan: '', date: '2026/09/01', kind: 'クレデンシャル', title: '', person: '本物 太郎', name: '本物の実績', note: '',
      owner: '佐藤', ownerDept: ''}});
  g.recordDiffEdit(gas.edit(sheet, 'F4', '保留'));
  const realLog = gas.records(DIFF_LOG).length;
  const realSnap = gas.records(SNAP).length;

  gas.select(sheet, 'E3');
  g.addSampleHistory();
  g.removeSampleHistory();

  assert.deepStrictEqual(gas.records(CRED_LOG).map(r => r['名称']), ['本物の実績']);
  assert.strictEqual(gas.records(DIFF_LOG).length, realLog);
  assert.ok(gas.records(DIFF_LOG).every(r => !r['イベントID'].startsWith('SMP-')));
  assert.strictEqual(gas.records(SNAP).length, realSnap);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /クレデンシャル・オファリング 10件・変更の記録 5件（スナップショット 5件）/);
});

test('removeSampleHistory：記録のシートにサンプルしか無くても削除できる（固定行以外を全部は消せないため）', () => {
  const {gas, g, sheet} = setupProject({launchers: false});
  gas.select(sheet, 'E3');
  g.addSampleHistory();
  // 記録のシートを、ちょうどサンプルの行で終わる大きさにする
  [CRED_LOG, DIFF_LOG, SNAP].forEach(name => {
    const s = gas.ss.getSheetByName(name);
    gas.asUser(() => s.deleteRows(s.getLastRow() + 1, s.getMaxRows() - s.getLastRow()));
  });
  g.removeSampleHistory();
  [CRED_LOG, DIFF_LOG, SNAP].forEach(name => assert.strictEqual(gas.records(name).length, 0, name));
  assert.match(gas.alerts[gas.alerts.length - 1].message, /削除しました/);
});
