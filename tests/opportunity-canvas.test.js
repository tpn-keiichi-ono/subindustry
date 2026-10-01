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
  assert.deepStrictEqual(data.sections.map(s => s.key),
    ['users', 'problems', 'today', 'ideas', 'use', 'metrics', 'adoption', 'challenges', 'budget', 'benefits']);
  // 改行入りの見出しも、セルの見出しと照合できる（表示は改行を空白にする）
  assert.deepStrictEqual(pairs(section(data, 'users')), [
    ['サブインダストリー', '食品小売'], ['先方部門', '営業企画部'], ['先方担当役職', '部長'], ['先方担当氏名', '原田 明博'],
    ['本案件におけるターゲットユーザー', '店舗の販促担当'], ['クライアントとのリレーション クライアント内のガバナンス', 'テストABC']
  ]);
  assert.deepStrictEqual(pairs(section(data, 'problems')),
    [['クライアントが置かれている状況・課題・現在の解決策', '来店客数の減少。チラシ中心の販促。']]);
  assert.deepStrictEqual(pairs(section(data, 'benefits')), [
    ['想定売上規模（百万）', '30'], ['期待値調整済 想定売上規模（百万）', '15'], ['受注月', '2026/12'], ['売上開始月', '']
  ]);
  // 当てはまる列が無い枠
  assert.strictEqual(section(data, 'metrics').configured, false);
  assert.deepStrictEqual(section(data, 'metrics').items, []);
  // 体制
  assert.deepStrictEqual(data.team.map(t => [t.label, t.value]),
    [['アカウント責任者', '佐藤'], ['BX担当', '鈴木'], ['品質責任者', ''], ['デリバリー担当', '']]);
  assert.deepStrictEqual(data.missing, []);
  // 読み取りだけ
  assert.deepStrictEqual(gas.writes, []);
});

test('getOpportunityCanvasData：クレデンシャル履歴のオファリングを Solution ideas、クレデンシャルを Adoption Strategy に出す', () => {
  const {g} = setup(VALUES);
  const base = {sheetName: '新FMT', row: 3, customer: '株式会社万代', customerColumn: FIRST_COLUMN + 1, linkId: ''};
  const values = extra => Object.assign({plan: '', date: '2026/09/01', kind: 'クレデンシャル', title: '', person: '原田 明博',
    name: '実績紹介', note: '', owner: '佐藤', ownerDept: ''}, extra);
  const parent = plain(g.saveCredentialEntry(Object.assign({}, base, {values: values({name: '販促DXの事例紹介'})}))).entry;
  g.saveCredentialEntry(Object.assign({}, base, {linkId: parent.id,
    values: values({kind: 'オファリング', name: '離反防止施策のご提案', date: '', plan: '2026/11/01'})}));

  const data = plain(g.getOpportunityCanvasData('新FMT', 3));
  assert.deepStrictEqual(section(data, 'ideas').offerings.map(e => [e.name, e.date, e.plan]),
    [['離反防止施策のご提案', '', '2026/11/01']]);
  assert.deepStrictEqual(section(data, 'adoption').credentials.map(e => [e.name, e.date]),
    [['販促DXの事例紹介', '2026/09/01']]);
});

test('getOpportunityCanvasData：見つからない列は missing で知らせる', () => {
  const headers = HEADERS.filter(h => h !== 'BX担当' && h !== '活動における課題');
  const {g} = setup(VALUES, {headers});
  const data = plain(g.getOpportunityCanvasData('新FMT', 3));
  assert.deepStrictEqual(data.missing, ['活動における課題', 'BX担当']);
  assert.deepStrictEqual(section(data, 'challenges').items, []);
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
