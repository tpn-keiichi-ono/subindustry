'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

/** 新FMT（2行目が見出し。C: サブインダストリー / D: 得意先 / E: 案件名 / F: サービスのリクエスト） */
function setup() {
  const gas = createGas();
  const header = ['No', 'メモ', 'サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'];
  gas.addSheet('新FMT', [['新FMT'], header,
    ['1', '', '食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい'],
    ['2', '', 'ドラッグストア', 'C社', 'EC立ち上げ', '在庫を店舗と共有したい']], {rows: 20, columns: 6});
  return {gas, g: gas.global};
}

const listOf = cell => {
  const rule = cell.getDataValidation();
  return rule && rule.getCriteriaType() === 'VALUE_IN_LIST' ? plain(rule.getCriteriaValues()[0]) : null;
};
const rangeOf = cell => {
  const rule = cell.getDataValidation();
  if (!rule || rule.getCriteriaType() !== 'VALUE_IN_RANGE' || rule.getAllowInvalid()) return null;
  const range = rule.getCriteriaValues()[0];
  return range.getSheet().getName() + '!' + range.getA1Notation();
};

test('setupTaskSheet：サービス案に対応付けるタスク管理シートを作る', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  assert.deepStrictEqual(gas.dump(sheet, 'A1:G1')[0], ['サービス案', 'タスク', '本部', '担当者', '期限', '状況', 'メモ'], '本部は担当者の左');
  assert.strictEqual(sheet.getFrozenRows(), 1);
  assert.strictEqual(rangeOf(sheet.getRange('A2')), '__TASK_CHOICES!A1', '登録したリクエストの候補から選ぶ（まだ1件も無い）');
  assert.strictEqual(rangeOf(sheet.getRange('A' + sheet.getMaxRows())), '__TASK_CHOICES!A1', '最後の行まで付ける');
  const choices = gas.ss.getSheetByName('__TASK_CHOICES');
  assert.ok(choices.isSheetHidden(), '候補のシートは非表示');
  assert.deepStrictEqual([choices.getMaxRows(), choices.getMaxColumns()], [1, 1], '使わない列・行を持たせない');
  ['B2', 'C2', 'D2'].forEach(a1 => assert.strictEqual(sheet.getRange(a1).getDataValidation(), null, a1 + '：タスク・本部・担当者は自由に入力できる'));
  assert.strictEqual(sheet.getRange('E2').getDataValidation().getCriteriaType(), 'DATE_IS_VALID_DATE');
  assert.deepStrictEqual(listOf(sheet.getRange('F2')), ['未着手', '対応中', '完了']);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /作りました/);
  assert.strictEqual(gas.alerts.length, 0, '前の版の列が無ければ確認しない');
});

test('setupTaskSheet：リクエスト シートが無いうちは、サービス案にプルダウンを付けずに知らせる', () => {
  const {gas, g} = setup();
  g.setupTaskSheet();
  assert.strictEqual(gas.ss.getSheetByName('タスク管理').getRange('A2').getDataValidation(), null);
  assert.match(gas.toasts[0].message, /setupRequestSheet\(\) を実行してください/);
});

/** 前の版のタスク管理シート（A〜C が案件の連動プルダウン。サービスは右端に足されていた） */
function addOldTaskSheet(gas, withService) {
  const header = ['サブインダストリー', '得意先', '案件名', 'タスク', '担当者', '期限', '状況', 'メモ'];
  const row = ['食品スーパー', 'A社', 'アプリ刷新', '提案書を送る', '佐藤', '', '対応中', '先方に確認'];
  if (withService) { header.push('サービス'); row.push('会員分析基盤'); }
  return gas.addSheet('タスク管理', [header, row], {rows: 10, columns: 10});
}

test('setupTaskSheet：前の版のシートは、確認してから A〜C を削除し、サービス案を左端に置く（D列以降の値は残す）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const sheet = addOldTaskSheet(gas, true);
  gas.confirmAnswer = 'YES';
  g.setupTaskSheet();

  assert.strictEqual(gas.alerts.length, 1);
  assert.match(gas.alerts[0].message, /「サブインダストリー」「得意先」「案件名」の列を削除し/);
  assert.deepStrictEqual(gas.dump(sheet, 'A1:G2'), [
    ['サービス案', 'タスク', '本部', '担当者', '期限', '状況', 'メモ'],
    ['会員分析基盤', '提案書を送る', '', '佐藤', '', '対応中', '先方に確認']
  ]);
  assert.strictEqual(sheet.getLastColumn(), 7);
  assert.strictEqual(rangeOf(sheet.getRange('A2')), '__TASK_CHOICES!A1');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /リクエストを選び直してほしいタスクが 1件あります/, '前の版のサービス案は、合うリクエストが無ければ残して知らせる');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /列を削除し、「サービス案」を左端に置きました/);
});

test('setupTaskSheet：サービスの列が無い前の版でも、サービス案を左端に足す', () => {
  const {gas, g} = setup();
  const sheet = addOldTaskSheet(gas, false);
  g.setupTaskSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A1:G2'), [
    ['サービス案', 'タスク', '本部', '担当者', '期限', '状況', 'メモ'],
    ['', '提案書を送る', '', '佐藤', '', '対応中', '先方に確認']
  ]);
});

test('setupTaskSheet：削除しないと答えたら A〜C は残し、サービス案だけを用意する', () => {
  const {gas, g} = setup();
  const sheet = addOldTaskSheet(gas, true);
  gas.confirmAnswer = 'NO';
  g.setupTaskSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A1:J1')[0],
    ['サブインダストリー', '得意先', '案件名', 'タスク', '本部', '担当者', '期限', '状況', 'メモ', 'サービス案']);
  assert.strictEqual(gas.dump(sheet, 'J2')[0][0], '会員分析基盤');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /列は残しました/);
});

test('setupTaskSheet：何度実行しても入力済みのタスクは消えない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  gas.asUser(() => sheet.getRange('A2:F2').setValues([['会員分析基盤', '要件を聞く', 'サービス本部', '佐藤', '2026/11/01', '未着手']]));
  g.setupTaskSheet();
  assert.deepStrictEqual(gas.dump(sheet, 'A2:F2')[0], ['会員分析基盤', '要件を聞く', 'サービス本部', '佐藤', '2026/11/01', '未着手']);
  assert.strictEqual(gas.alerts.length, 0);
});

/** 取り込み元のリクエストを、選択パネルからすべて登録する（上から SR-0001, SR-0002 …） */
function addAll(g) {
  g.getRequestPickerData().items.forEach(item => assert.strictEqual(g.addServiceRequest(item.key).ok, true));
}
/** 候補のシート（__TASK_CHOICES）の A 列 */
const choicesOf = gas => gas.dump(gas.ss.getSheetByName('__TASK_CHOICES'), 'A1:A' + gas.ss.getSheetByName('__TASK_CHOICES').getMaxRows()).map(r => r[0]);

test('タスク管理のサービス案は、登録したリクエストを「ID リクエスト」で選ぶ。追加・削除するとすぐ候補が変わる', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const sheet = gas.ss.getSheetByName('タスク管理');
  addAll(g);
  assert.deepStrictEqual(choicesOf(gas), ['SR-0001 会員の購買分析をしたい', 'SR-0002 在庫を店舗と共有したい']);
  assert.strictEqual(rangeOf(sheet.getRange('A2')), '__TASK_CHOICES!A1:A2', '追加したら、パネルを開き直さなくても選べる');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
  sheet.getRange('A2').setValue('SR-0002 在庫を店舗と共有したい');
  assert.throws(() => sheet.getRange('A3').setValue('会員分析基盤'), /登録したリクエスト（ID とリクエスト）から選んでください/, '候補に無い値は入れられない');
  assert.strictEqual(sheet.getRange('A2').getDataValidation().getHelpText(), '「リクエスト」シートに登録したリクエスト（ID とリクエスト）から選んでください。');

  gas.writes.length = 0;
  const list = g.getRequestListData();
  const target = list.registered.find(r => r.id === 'SR-0001');
  assert.strictEqual(g.deleteServiceRequest(target.row, target.key, target.id).ok, true);
  assert.deepStrictEqual(choicesOf(gas), ['SR-0002 在庫を店舗と共有したい'], '削除したリクエストは候補から外す');
  assert.strictEqual(rangeOf(sheet.getRange('A2')), '__TASK_CHOICES!A1');
  assert.strictEqual(gas.dump(sheet, 'A2')[0][0], 'SR-0002 在庫を店舗と共有したい');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('候補の文字：改行・続いた空白は1つの空白に、長いリクエストは「…」で切る。ID の順に並べ、同じ ID は上の行だけ', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  const long = 'あ'.repeat(45);
  gas.asUser(() => request.getRange('A2:I4').setValues([
    ['SR-0003', '食品スーパー', 'A社', 'アプリ刷新', '', '', '', '', '会員の\n購買分析を  したい'],
    ['SR-0001', 'ドラッグストア', 'C社', 'EC立ち上げ', '', '', '', '', long],
    ['SR-0003', 'ドラッグストア', 'D社', '店舗', '', '', '', '', 'コピーした行']
  ]));
  g.getRequestPickerData();
  assert.deepStrictEqual(choicesOf(gas), ['SR-0001 ' + 'あ'.repeat(40) + '…', 'SR-0003 会員の 購買分析を したい', 'SR-0004 コピーした行'],
    '重なった ID は振り直してから候補にする');
});

test('入力済みのタスクは、今の候補に合わせる：同じ ID の候補へ、前の版のサービス案の名前はそのリクエストが1件だけなら移す（決まらないものは残して知らせる）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  addAll(g);
  const request = gas.ss.getSheetByName('リクエスト');
  gas.asUser(() => request.getRange('L2:L3').setValues([['会員分析基盤'], ['共通基盤']]));
  gas.addSheet('タスク管理', [
    ['サービス案', 'タスク', '本部', '担当者', '期限', '状況', 'メモ'],
    ['会員分析基盤', '要件を聞く', '', '', '', '', ''],          // サービス案の名前（リクエストが1件だけ）→ SR-0001
    ['SR-0002 前のリクエストの文', '見積もり', '', '', '', '', ''],   // ID が同じ → 今の文に
    ['名前の無いサービス', '調べる', '', '', '', '', ''],            // 合うものが無い → 残す
    ['', '', '', '', '', '', ''],
    ['SR-0002 在庫を店舗と共有したい', 'そのまま', '', '', '', '', '']
  ], {rows: 10, columns: 7});
  gas.toasts.length = 0;
  g.getRequestPickerData();
  const sheet = gas.ss.getSheetByName('タスク管理');
  assert.deepStrictEqual(gas.dump(sheet, 'A2:B6'), [
    ['SR-0001 会員の購買分析をしたい', '要件を聞く'],
    ['SR-0002 在庫を店舗と共有したい', '見積もり'],
    ['名前の無いサービス', '調べる'],
    ['', ''],
    ['SR-0002 在庫を店舗と共有したい', 'そのまま']
  ]);
  const toast = gas.toasts[gas.toasts.length - 1].message;
  assert.match(toast, /タスク管理 の「サービス案」を、今のリクエスト（ID とリクエスト）に合わせました（2件）。/);
  assert.match(toast, /リクエストを選び直してほしいタスクが 1件あります（セルの右上が赤いもの）。/);

  // 同じサービス案のリクエストが2件あれば、どちらか決められないので残す
  gas.asUser(() => {
    request.getRange('L3').setValue('会員分析基盤');
    sheet.getRange('A2').setValue('会員分析基盤');
  });
  g.setupTaskSheet();
  assert.strictEqual(gas.dump(sheet, 'A2')[0][0], '会員分析基盤');
  assert.match(gas.toasts[gas.toasts.length - 1].message, /選び直してほしいタスクが 2件/);
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('リクエスト シートに ID の列がまだ無い（setupRequestSheet() の前）うちは、候補を作らず、入力済みのタスクも変えない', () => {
  const {gas, g} = setup();
  gas.addSheet('リクエスト', [
    ['サブインダストリー', '得意先', '案件名', 'リクエスト', '判断', 'サービス案'],
    ['食品スーパー', 'A社', 'アプリ刷新', '会員の購買分析をしたい', 'サービス化検討', '会員分析基盤']
  ], {rows: 20, columns: 6});
  const sheet = gas.addSheet('タスク管理', [['サービス案', 'タスク', '本部', '担当者', '期限', '状況', 'メモ'],
    ['会員分析基盤', '要件を聞く', '', '', '', '', '']], {rows: 10, columns: 7});
  g.setupTaskSheet();
  assert.strictEqual(gas.dump(sheet, 'A2')[0][0], '会員分析基盤');
  assert.strictEqual(sheet.getRange('A2').getDataValidation(), null);
  assert.strictEqual(gas.ss.getSheetByName('__TASK_CHOICES'), null);
  assert.match(gas.toasts[gas.toasts.length - 1].message, /ID の列）が無いため、「サービス案」の列にプルダウンを付けていません。setupRequestSheet\(\) を実行してください/);
});

test('パネルを開くたびには確かめない（候補・列・行数が同じなら30分に1回）。プルダウンが消えた・候補のシートが消えたら付け直す。開いているシートは変えない', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  const request = gas.ss.getSheetByName('リクエスト');
  gas.select(request, 'A1');
  addAll(g);
  g.setupTaskSheet();
  gas.select(request, 'A1');
  const sheet = gas.ss.getSheetByName('タスク管理');
  gas.ss.deleteSheet(gas.ss.getSheetByName('__TASK_CHOICES'));
  g.getRequestPickerData();
  assert.strictEqual(gas.ss.getActiveSheet().getName(), 'リクエスト', '候補のシートを作っても、開いているシートは変えない');
  assert.strictEqual(rangeOf(sheet.getRange('A2')), '__TASK_CHOICES!A1:A2', '候補のシートが消えたら作り直して付け直す');

  gas.writes.length = 0;
  gas.reads.length = 0;
  g.getRequestPickerData();
  assert.deepStrictEqual(gas.writes.filter(w => w.sheet === 'タスク管理' || w.sheet === '__TASK_CHOICES'), []);
  assert.deepStrictEqual(gas.reads.filter(w => (w.sheet === 'タスク管理' && w.a1 !== 'A1:G1') || w.sheet === '__TASK_CHOICES'), [],
    '候補もタスクも読まない（タスク管理は見出しだけ）');

  // 貼り付けでプルダウンが消えた行・ほかを見ているプルダウンがあれば、確かめる間隔を過ぎてから開いたときに列ごと付け直す
  const last = 'A' + sheet.getMaxRows();
  gas.asUser(() => {
    sheet.getRange('A3').clearDataValidations();
    sheet.getRange(last).setDataValidation(g.SpreadsheetApp.newDataValidation()
      .requireValueInRange(request.getRange('L2:L'), true).setAllowInvalid(false).build());
  });
  const props = g.PropertiesService.getDocumentProperties();
  const saved = JSON.parse(props.getProperty('SVC_TASK_LINK'));
  props.setProperty('SVC_TASK_LINK', JSON.stringify(Object.assign(saved, {at: Date.now() - 31 * 60 * 1000})));   // 31分前に確かめた
  g.getRequestPickerData();
  ['A2', 'A3', last].forEach(a1 => assert.strictEqual(rangeOf(sheet.getRange(a1)), '__TASK_CHOICES!A1:A2', a1));

  // 前の版の記録（サービス案の列を参照していたとき）が残っていても付け直す
  props.setProperty('SVC_TASK_LINK', JSON.stringify({state: request.getSheetId() + ':12:' + sheet.getSheetId() + ':1:9', at: Date.now()}));
  gas.asUser(() => sheet.getRange('A2:A' + sheet.getMaxRows()).setDataValidation(g.SpreadsheetApp.newDataValidation()
    .requireValueInRange(request.getRange('L2:L'), true).setAllowInvalid(false).build()));
  g.getRequestPickerData();
  assert.strictEqual(rangeOf(sheet.getRange('A2')), '__TASK_CHOICES!A1:A2');
  assert.ok(gas.writes.every(w => w.locked), '書き込みはロックの中で行う');
});

test('候補を更新できなかったときは黙らずに知らせる：パネルの読み込み（ほかは済ませる）・追加・削除の結果に warning（ほかの処理が実行中のときは知らせない）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  const original = g.svcLinkTaskServiceColumn_;
  g.svcLinkTaskServiceColumn_ = () => { throw new Error('Exception: 保護されているセルやオブジェクトを編集しようとしています。'); };
  try {
    const data = g.getRequestPickerData();
    assert.strictEqual(data.synced, true, 'ID・取り込み元に合わせるのは済ませる');
    assert.strictEqual(data.warning, 'タスク管理 の「サービス案」の候補を更新できませんでした：保護されているセルやオブジェクトを編集しようとしています。' +
      '（管理者に diagnoseTaskChoices() の実行を頼んでください）');
    const added = g.addServiceRequest(data.items[0].key);
    assert.strictEqual(added.ok, true, '追加はできる');
    assert.match(added.warning, /候補を更新できませんでした/);
    const fromList = g.addServiceRequestFromList(g.getRequestListData().unregistered[0].key);
    assert.match(fromList.warning, /候補を更新できませんでした/);
    const target = g.getRequestListData().registered[0];
    assert.match(g.deleteServiceRequest(target.row, target.key, target.id).warning, /候補を更新できませんでした/);
  } finally {
    g.svcLinkTaskServiceColumn_ = original;
  }
  assert.strictEqual(g.getRequestPickerData().warning, '', '更新できれば知らせない');
  assert.strictEqual(g.addServiceRequest(g.getRequestPickerData().items[0].key).warning, '');

  gas.lockBusy = 'script';
  const busy = g.getRequestPickerData();
  gas.lockBusy = false;
  assert.deepStrictEqual([busy.synced, busy.warning], [false, ''], 'ほかの処理が実行中のときは、読むだけにして知らせない');
});

test('diagnoseTaskChoices：コードの版・候補の数・候補のシート・プルダウンの参照先・最後に合わせた時刻・保護を画面に出す（読むだけ）', () => {
  const {gas, g} = setup();
  g.setupRequestSheet();
  g.setupTaskSheet();
  addAll(g);
  const sheet = gas.ss.getSheetByName('タスク管理');
  gas.writes.length = 0;
  g.diagnoseTaskChoices();
  assert.deepStrictEqual(gas.writes, [], '書き込まない');
  const ok = gas.alerts[gas.alerts.length - 1].message;
  assert.match(ok, /コードの版：最新です。/);
  assert.match(ok, /候補にするリクエスト（「リクエスト」の ID のある行）：2件（SR-0001〜SR-0002）/);
  assert.match(ok, /候補のシート（__TASK_CHOICES）：2件（今のリクエストと同じ）。/);
  assert.match(ok, /「タスク管理」の「サービス案」：A 列/);
  assert.match(ok, /A2：__TASK_CHOICES!A1:A2 を参照（今の候補）。/);
  assert.match(ok, /最後に候補を合わせた時刻：\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}/);
  assert.match(ok, /保護：ありません。/);

  // 前の版の関数が動いている・候補が古い・前の版のプルダウン・保護がある
  const original = g.svcServiceRule_;
  g.svcServiceRule_ = function svcServiceRule_(ss) { return null; };
  gas.asUser(() => {
    gas.ss.getSheetByName('__TASK_CHOICES').getRange('A2').setValue('SR-0002 古い文');
    sheet.getRange('A2').setDataValidation(g.SpreadsheetApp.newDataValidation()
      .requireValueInRange(gas.ss.getSheetByName('リクエスト').getRange('L2:L'), true).setAllowInvalid(false).build());
    sheet.protect().addEditor('admin@example.com');
  });
  try {
    g.diagnoseTaskChoices();
  } finally {
    g.svcServiceRule_ = original;
  }
  const bad = gas.alerts[gas.alerts.length - 1].message;
  assert.match(bad, /コードの版：前の版の処理が動いています（svcServiceRule_）。/);
  assert.match(bad, /今のリクエストと合っていません（足りない 1件・余分 1件）。/);
  assert.match(bad, /A2：リクエスト!L2:L\d+ を参照（前の版：サービス案の名前の列）。/);
  assert.match(bad, /保護：「タスク管理」のシート全体（編集できる人：admin@example.com）。/);
});

test('タスク管理シートの編集で動く処理は無い（単純トリガーの onEdit を置かない。プルダウンは候補のシートを参照するだけ）', () => {
  const {g} = setup();
  assert.strictEqual(typeof g.onEdit, 'undefined');
});
