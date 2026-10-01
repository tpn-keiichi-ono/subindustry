'use strict';

const {test, assert} = require('./lib/harness');
const {createGas, plain} = require('./lib/gas-mock');

// 新FMT の見出し（2行目）。実際のシートと同じ並び（C列から）
const HEADERS = [
  'サブインダストリー', '得意先', '先方部門', '先方担当役職', '先方担当氏名',
  'クライアントとのリレーション\nクライアント内のガバナンス', '案件名', '案件のスコープ', 'スコープに対する競合他社',
  'プロジェクト開始-終了', '案件ステータス', '想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月',
  '売上開始月', '提案開始日', '活動状況', '活動における課題', 'アカウント責任者', 'BX担当', '品質責任者', 'デリバリー担当',
  'クライアントが置かれている状況・課題・現在の解決策', 'アカウントプラン（リンク）', 'サブインシナリオ',
  '本案件におけるターゲットユーザー', 'マーケットに出すソリューションか', '想定される価値創出のケース／期待される導入効果',
  '提案を勝ち取るための戦略・差異化要素'
];
const FIRST_COLUMN = 3;   // C列

/** 見出しと1行ぶんの値で、新FMT・新FMT2 を作る。values = {見出し: 値} */
function setup(values, options) {
  const o = options || {};
  const gas = createGas();
  const headers = (o.headers || HEADERS).concat(['最終更新日時']);
  const width = FIRST_COLUMN - 1 + headers.length;
  ['新FMT', '新FMT2'].forEach(name => {
    const header = new Array(width).fill('');
    const row = new Array(width).fill('');
    headers.forEach((h, i) => {
      header[FIRST_COLUMN - 1 + i] = h;
      row[FIRST_COLUMN - 1 + i] = values[h] || '';
    });
    gas.addSheet(name, [new Array(width).fill(''), header, row], {rows: 10, columns: width});
  });
  return {gas, g: gas.global, sheet: gas.ss.getSheetByName('新FMT')};
}

const VALUES = {
  'サブインダストリー': '食品小売',
  '得意先': '株式会社万代',
  '先方部門': '営業企画部',
  '先方担当役職': '部長',
  '先方担当氏名': '原田 明博',
  'クライアントとのリレーション\nクライアント内のガバナンス': 'テストABC',
  '案件名': 'シナリオ提案',
  '案件のスコープ': '販促DXの全社展開',
  'スコープに対する競合他社': 'A社・B社',
  'プロジェクト開始-終了': '2026/11〜2027/03',
  '案件ステータス': '提案中',
  '想定売上規模（百万）': '30',
  '期待値調整済\n想定売上規模（百万）': '15',
  '受注月': '2026/12',
  '活動状況': '4/15初回訪問。5/27離反シナリオ紹介。',
  '活動における課題': '決裁者に会えていない',
  'アカウント責任者': '佐藤',
  'BX担当': '鈴木',
  'クライアントが置かれている状況・課題・現在の解決策': '来店客数の減少。チラシ中心の販促。',
  'アカウントプラン（リンク）': 'https://docs.google.com/spreadsheets/d/plan/edit',
  'サブインシナリオ': '離反防止',
  '本案件におけるターゲットユーザー': '店舗の販促担当',
  '想定される価値創出のケース／期待される導入効果': '来店頻度の向上',
  '提案を勝ち取るための戦略・差異化要素': '他社事例の提示'
};

const section = (data, key) => data.sections.find(s => s.key === key);
const pairs = s => s.items.map(i => [i.label, i.value]);

test('getOpportunityCanvasData：見出し名で列を探し、キャンバスの枠に分けて返す', () => {
  const {gas, g} = setup(VALUES);
  const data = plain(g.getOpportunityCanvasData('新FMT', 3));

  assert.deepStrictEqual(data.header, {
    title: 'シナリオ提案', customer: '株式会社万代', status: '提案中',
    planUrl: 'https://docs.google.com/spreadsheets/d/plan/edit', updatedAt: ''
  });
  // 当てはまる列が無い枠（利用の指標・予算・ビジネス上の課題）は出さない。枠の名前は日本語
  assert.deepStrictEqual(data.sections.map(s => [s.key, s.no, s.title]), [
    ['users', 2, '顧客・ユーザー'], ['problems', 1, '課題'], ['today', 3, '現在の解決策'], ['ideas', 1, '解決策のアイデア'],
    ['use', 4, '使われ方・導入効果'], ['adoption', 5, '導入戦略'], ['benefits', 6, 'ビジネス上の効果・指標']
  ]);
  data.sections.forEach(s => assert.ok(!/[A-Za-z]/.test(s.title), s.title));
  assert.deepStrictEqual(pairs(section(data, 'users')), [
    ['サブインダストリー', '食品小売'], ['先方部門', '営業企画部'], ['先方担当役職', '部長'], ['先方担当氏名', '原田 明博'],
    ['本案件におけるターゲットユーザー', '店舗の販促担当']
  ]);
  assert.deepStrictEqual(pairs(section(data, 'adoption')),
    [['提案を勝ち取るための戦略・差異化要素', '他社事例の提示'], ['活動状況', '4/15初回訪問。5/27離反シナリオ紹介。']]);
  assert.deepStrictEqual(pairs(section(data, 'problems')),
    [['クライアントが置かれている状況・課題・現在の解決策', '来店客数の減少。チラシ中心の販促。']]);
  // 改行入りの見出しも、セルの見出しと照合できる（表示は改行を空白にする）
  assert.deepStrictEqual(pairs(section(data, 'benefits')), [
    ['想定売上規模（百万）', '30'], ['期待値調整済 想定売上規模（百万）', '15'], ['受注月', '2026/12'], ['売上開始月', '']
  ]);
  // 枠と合わない列・体制は出さない
  const text = JSON.stringify(data);
  ['テストABC', '決裁者に会えていない', '2026/11〜2027/03', '佐藤', '鈴木'].forEach(v =>
    assert.ok(!text.includes(v), '枠と合わない列の値が入っています: ' + v));
  assert.strictEqual(data.team, undefined);
  assert.deepStrictEqual(data.missing, []);
  // 読み取りだけ
  assert.deepStrictEqual(gas.writes, []);
});

test('getOpportunityCanvasData：クレデンシャル・オファリングの履歴は含めない（行の値だけ）', () => {
  const {g} = setup(VALUES);
  const base = {sheetName: '新FMT', row: 3, customer: '株式会社万代', customerColumn: FIRST_COLUMN + 1, linkId: ''};
  const values = extra => Object.assign({plan: '', date: '2026/09/01', kind: 'クレデンシャル', title: '', person: '原田 明博',
    name: '実績紹介', note: '', owner: '佐藤', ownerDept: ''}, extra);
  const parent = plain(g.saveCredentialEntry(Object.assign({}, base, {values: values({name: '販促DXの事例紹介'})}))).entry;
  g.saveCredentialEntry(Object.assign({}, base, {linkId: parent.id,
    values: values({kind: 'オファリング', name: '離反防止施策のご提案', date: '', plan: '2026/11/01'})}));

  const data = plain(g.getOpportunityCanvasData('新FMT', 3));
  const text = JSON.stringify(data);
  assert.ok(!text.includes('販促DXの事例紹介') && !text.includes('離反防止施策のご提案'), '履歴の内容が入っています');
  data.sections.forEach(s => assert.deepStrictEqual(Object.keys(s).sort(),
    ['configured', 'items', 'key', 'no', 'question', 'title']));
  assert.deepStrictEqual(pairs(section(data, 'ideas')), [
    ['案件のスコープ', '販促DXの全社展開'], ['サブインシナリオ', '離反防止']
  ]);
});

test('getOpportunityCanvasData：見つからない列は missing で知らせる', () => {
  const headers = HEADERS.filter(h => h !== '先方部門' && h !== '受注月');
  const {g} = setup(VALUES, {headers});
  const data = plain(g.getOpportunityCanvasData('新FMT', 3));
  assert.deepStrictEqual(data.missing, ['先方部門', '受注月']);
  assert.ok(!section(data, 'users').items.some(i => i.label === '先方部門'));
});

test('getOpportunityCanvasData：行にアカウントプランの URL が無ければ、33シナリオ攻略先リストから探す', () => {
  const values = Object.assign({}, VALUES, {'アカウントプラン（リンク）': ''});
  const {gas, g} = setup(values);
  const header = new Array(13).fill('');
  const row = new Array(13).fill('');
  row[4] = '株式会社 万代';
  row[12] = 'https://docs.google.com/spreadsheets/d/from-list/edit';
  gas.addSheet('33シナリオ攻略先リスト', [header, row], {rows: 5, columns: 13});
  assert.strictEqual(plain(g.getOpportunityCanvasData('新FMT', 3)).header.planUrl,
    'https://docs.google.com/spreadsheets/d/from-list/edit');
});

test('openOpportunityCanvas：選んだ行のキャンバスをモーダルで開く。見出し行なら知らせる', () => {
  const {gas, g, sheet} = setup(VALUES);
  gas.select(sheet, 'H3');
  g.openOpportunityCanvas();
  assert.strictEqual(gas.dialogs.length, 1);
  assert.strictEqual(gas.dialogs[0].html.file, 'OpportunityCanvasDialog');
  assert.deepStrictEqual([gas.dialogs[0].html.data.row, gas.dialogs[0].html.data.customer], [3, '株式会社万代']);

  gas.select(sheet, 'H2');
  g.openOpportunityCanvas();
  assert.strictEqual(gas.dialogs.length, 1);
  assert.match(gas.alerts[gas.alerts.length - 1].message, /データ行（3行目以降）を選んでください/);
});

test('サンプルデータ（samples/新FMT_サンプルデータ.tsv）：見出しの並びどおりで、どの行もキャンバスの全部の枠に値が入る', () => {
  const file = require('path').join(__dirname, '..', 'samples', '新FMT_サンプルデータ.tsv');
  const rows = require('fs').readFileSync(file, 'utf8').split('\n').filter(Boolean).map(line => line.split('\t'));
  assert.strictEqual(rows.length, 5);
  rows.forEach(row => {
    assert.strictEqual(row.length, HEADERS.length, '列の数が見出しと合いません: ' + row[1]);
    row.forEach(v => assert.ok(!/^[=+\-@]/.test(v), '数式になる値があります: ' + v));
    const values = {};
    HEADERS.forEach((h, i) => { values[h] = row[i]; });
    const {g} = setup(values);
    const data = plain(g.getOpportunityCanvasData('新FMT', 3));
    assert.strictEqual(data.header.customer, row[1]);
    assert.strictEqual(data.header.title, row[6]);
    assert.deepStrictEqual(data.missing, []);
    data.sections.forEach(s => assert.ok(s.items.some(i => i.value), row[1] + ' の「' + s.title + '」が空です'));
  });
});
