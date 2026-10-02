/**
 * オポチュニティキャンバス（案件ごとのキャンバス）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs・HistorySidebar.gs と同じプロジェクトに置くファイル。
 * 画面は HTMLファイル「OpportunityCanvasDialog」。
 *
 * 仕組み
 * - 追跡シート（新FMT など）の1行を1件の案件として、その行の値をキャンバスの枠に並べてモーダルで表示する
 * - どの列をどの枠に出すかは OC_OPTIONS.sections の見出し名で決める（列の位置ではなく見出しで探すので、列を動かしても使える）
 *   見出しの改行・空白、全角・半角の違いは無視して照合する
 * - 出すのは追跡シートの行に入力されている値だけ（クレデンシャル・オファリングの履歴は含めない）
 * - 読み取りだけ（シートには書き込まない）。値を自動で考えて埋めることはせず、入っている内容だけを出す
 */

const OC_OPTIONS = {
  template: 'OpportunityCanvasDialog',   // HTMLファイル名（.gs と同じ名前は付けられないため別名）
  // 見出しの上に出す項目
  header: {
    title: '案件名',
    customer: '得意先',
    status: '案件ステータス',
    planLink: 'アカウントプラン（リンク）'
  },
  // キャンバスの枠（no は記入順の番号、question は枠の問い）
  // 新FMT に当てはまる列が無い枠（利用の指標・予算・ビジネス上の課題）は出さない。枠の内容と合わない列も入れない
  sections: [
    {key: 'users', no: 2, title: '顧客・ユーザー',
      question: 'この課題を抱えているのは、どんな顧客・ユーザーか',
      headers: ['サブインダストリー', '先方部門', '先方担当役職', '先方担当氏名', '本案件におけるターゲットユーザー']},
    {key: 'problems', no: 1, title: '課題',
      question: '顧客・ユーザーが今抱えている課題・ニーズは何か',
      headers: ['クライアントが置かれている状況・課題・現在の解決策']},
    {key: 'today', no: 3, title: '現在の解決策',
      question: '顧客は今その課題にどう対処しているか（競合・代替手段）',
      headers: ['スコープに対する競合他社']},
    {key: 'ideas', no: 1, title: '解決策のアイデア',
      question: '提供する商品・サービス・提案の内容',
      headers: ['案件のスコープ', 'サブインシナリオ']},
    {key: 'use', no: 4, title: '使われ方・導入効果',
      question: '解決策によって、顧客の行動や成果はどう変わるか',
      headers: ['想定される価値創出のケース／期待される導入効果']},
    {key: 'adoption', no: 5, title: '導入戦略',
      question: '顧客はどうやって解決策を知り、採用するか',
      headers: ['提案を勝ち取るための戦略・差異化要素', '活動状況']},
    {key: 'benefits', no: 6, title: 'ビジネス上の効果・指標',
      question: '受注によって、自社の業績指標はどう変わるか',
      headers: ['想定売上規模（百万）', '期待値調整済\n想定売上規模（百万）', '受注月', '売上開始月']}
  ]
};

/* ---------------- 開く ---------------- */

/** エディタから実行：選んでいる行のオポチュニティキャンバスを開く（サイドバーの「キャンバス」と同じ）。 */
function openOpportunityCanvas() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    if (!range) throw new Error('対象行のセルを選択してください。');
    if (range.getNumRows() > 1) throw new Error('1行だけ選択してください。');
    ocShowDialog_(sheet, range.getRow(), credActiveEmail_());
  } catch (error) {
    ui.alert('オポチュニティキャンバス', error.message, ui.ButtonSet.OK);
  }
}

/** サイドバーのボタンから。 */
function openOpportunityCanvasFromSidebar(sheetName, row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(sheetName));
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  ocShowDialog_(sheet, Number(row), credActiveEmail_());
}

function ocShowDialog_(sheet, row, email) {
  const target = credResolveTarget_(sheet, row);
  const size = credDialogSize_(email);   // ほかのモーダルと同じ大きさ（ブラウザ幅の95%）

  const template = HtmlService.createTemplateFromFile(OC_OPTIONS.template);
  template.sheetName = target.sheet.getName();
  template.row = target.row;
  template.customer = target.customer;
  template.openedWidth = size.width;
  template.openedHeight = size.height;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(size.width).setHeight(size.height),
    ' '   // 見出しはモーダルの中に表示する
  );
}

/* ---------------- モーダルから呼ばれる関数（読み取りのみ） ---------------- */

/**
 * 対象行の値をキャンバスの枠に分けて返す。
 * {sheetName, row, customer, header, sections: [{key, no, title, question, configured, items}],
 *  missing: [見つからなかった見出し], today}
 */
function getOpportunityCanvasData(sheetName, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(String(sheetName));
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  const target = credResolveTarget_(sheet, Number(row));
  const rule = target.rule;
  const timezone = ss.getSpreadsheetTimeZone();

  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(rule.headerRow, 1, 1, width).getDisplayValues()[0];
  const values = sheet.getRange(target.row, 1, 1, width).getDisplayValues()[0];
  const byHeader = new Map();
  headers.forEach((h, i) => {
    const key = ocHeaderKey_(h);
    if (key && !byHeader.has(key)) byHeader.set(key, {label: diffCleanHeaderText_(h), value: String(values[i] || '').trim()});
  });

  const missing = [];
  const pick = header => {
    const found = byHeader.get(ocHeaderKey_(header));
    if (!found) missing.push(diffCleanHeaderText_(header));
    return found || null;
  };
  const valueOf = header => { const found = pick(header); return found ? found.value : ''; };

  const sections = OC_OPTIONS.sections.map(section => ({
    key: section.key, no: section.no, title: section.title, question: section.question,
    configured: section.headers.length > 0,
    items: section.headers.map(pick).filter(Boolean)
  }));

  // アカウントプラン：行の「アカウントプラン（リンク）」、無ければ 33シナリオ攻略先リスト から
  let planUrl = valueOf(OC_OPTIONS.header.planLink);
  if (!hsIsUrl_(planUrl)) planUrl = hsAccountPlanIndex_(ss)[credNormalize_(target.customer)] || '';

  let updatedAt = '';
  try {
    updatedAt = sheet.getRange(target.row, diffStampColumn_(sheet, rule)).getDisplayValue();
  } catch (_) {}

  return {
    sheetName: sheet.getName(),
    row: target.row,
    customer: target.customer,
    header: {
      title: valueOf(OC_OPTIONS.header.title),
      customer: target.customer,
      status: valueOf(OC_OPTIONS.header.status),
      planUrl: hsIsUrl_(planUrl) ? planUrl.trim() : '',
      updatedAt
    },
    sections,
    missing: Array.from(new Set(missing)),
    today: Utilities.formatDate(new Date(), timezone, 'yyyy/MM/dd')
  };
}

/** 見出しの照合用（改行・空白、全角・半角の違いを無視する）。 */
function ocHeaderKey_(header) {
  return credNormalize_(header);
}
