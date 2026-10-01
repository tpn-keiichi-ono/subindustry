/**
 * テスト用のスプレッドシート（新FMT・新FMT2 と同じ形）
 *
 * 2行目が見出し。追跡範囲は D3:AF（DIFF_RULES）。
 *   A: No / B: メモ / D: 得意先 / E: 担当 / F: 状況
 *   AG: 最終更新日時 / AH: クレデンシャル（改行）オファリング登録 / AI: 変更履歴（ボタン列は追跡範囲の外）
 */
'use strict';

const {createGas} = require('./gas-mock');

const COL = {no: 1, memo: 2, customer: 4, owner: 5, status: 6, stamp: 33, credButton: 34, changeButton: 35};
const WIDTH = 35;

function headerRow() {
  const row = new Array(WIDTH).fill('');
  row[COL.no - 1] = 'No';
  row[COL.memo - 1] = 'メモ';
  row[COL.customer - 1] = '得意先';
  row[COL.owner - 1] = '担当';
  row[COL.status - 1] = '状況';
  row[COL.stamp - 1] = '最終更新日時';
  row[COL.credButton - 1] = 'クレデンシャル\nオファリング登録';
  row[COL.changeButton - 1] = '変更履歴';
  return row;
}

function dataRow(no, customer, owner, status) {
  const row = new Array(WIDTH).fill('');
  row[COL.no - 1] = String(no);
  row[COL.customer - 1] = customer;
  row[COL.owner - 1] = owner;
  row[COL.status - 1] = status;
  return row;
}

const DEFAULT_ROWS = [
  ['A社', '佐藤', '提案中'],
  ['B社', '鈴木', '受注'],
  ['Ｃ社', '田中', '']
];

/** 追跡シートを2つ作る。rows = [[得意先, 担当, 状況], ...] */
function addTrackedSheets(gas, rows) {
  rows = rows || DEFAULT_ROWS;
  ['新FMT', '新FMT2'].forEach(name => {
    const title = new Array(WIDTH).fill('');
    title[0] = name;
    gas.addSheet(name, [title, headerRow()].concat(rows.map((r, i) => dataRow(i + 1, r[0], r[1], r[2]))),
      {rows: 10, columns: WIDTH});
  });
  return gas.ss.getSheetByName('新FMT');
}

/**
 * 管理者のセットアップまで済ませた環境を返す。
 * options = {rows, launchers: false でボタン列の設定を省く, user}
 */
function setupProject(options) {
  const o = options || {};
  const gas = createGas({user: o.user});
  const sheet = addTrackedSheets(gas, o.rows);
  const g = gas.global;
  g.setupDiffTracking();
  if (o.launchers !== false) {
    g.setupCredentialLauncher();
    g.setupChangeHistoryLauncher();
  }
  gas.writes.length = 0;
  gas.toasts.length = 0;
  return {gas, g, sheet};
}

module.exports = {COL, WIDTH, addTrackedSheets, setupProject};
