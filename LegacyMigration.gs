/**
 * 旧フォーマットからの転記（管理者がエディタから1回実行する）
 * 旧フォーマットのスプレッドシートの「案件_{サブインダストリー名}」シートの案件を、このスプレッドシートの同じサブインダストリーのシート
 * （test.gs の DIFF_RULES のシート）に転記する。旧のスプレッドシートは URL を聞いて開く（中身は実行した人の環境だけで読む）。
 *
 * - 列は見出しで対応させる（MIG_OPTIONS.columns。旧の見出し → 新の見出し）。対応しない列は空欄にする
 *   （企画担当・提案シナリオ・マーケなどの配分・サービス配分チェック・備考は転記しない）。
 * - 新のシートの数式の列（3行目に数式がある No.・期待値）は書かない・消さない。
 * - 新のシートの今の行は上書きする（3行目から下を消してから書く）。
 * - プルダウンの選択肢に無い値（得意先マスタに無い得意先・定義に無いステータス）も、プルダウンが「警告を表示」（無効な値も入れられる）なら
 *   そのまま書き、どのセルかを一覧で知らせる（MIG_OPTIONS.writeOutOfList。2026/10/26 から）。「入力を拒否」のプルダウンは書くと止まるので空欄にする
 *   （得意先マスタなどに足すか、プルダウンを「警告を表示」にしてから、もう一度実行すると入る。何度実行してもよい）。
 * - 書いたあと、差分追跡の比較基準を取り直し（setupDiffTracking。転記を変更として記録しない）、会議サマリの基準を記録する
 *   （転記した案件を次の資料で「新規」に数えない）。
 */

const MIG_OPTIONS = {
  sheetPrefix: '案件_',        // 旧のシートの名前（案件_スーパー・GMS）
  headerRows: [1, 2],          // 旧の見出しの行（2行に分かれている・結合している。下の行を優先）
  dataStartRow: 4,             // 旧のデータの始まり（3行目は「手動入力」「プルダウンより選択」の行）
  // 旧の見出し → 新の見出し（見出しは改行・空白・全角半角の違いを無視して比べる）
  columns: [
    ['得意先', '得意先'],
    ['案件名', '案件名'],
    ['先方部門', '先方部門'],
    ['先方担当役職', '先方担当役職'],
    ['先方担当氏名', '先方担当氏名'],
    ['営業担当', 'アカウント責任者'],
    ['提案開始日', '提案開始日'],
    ['ステータス', '案件ステータス'],
    ['活動状況', '活動状況'],
    ['次回アクション', '次回アクション'],
    ['想定売上規模（百万）', '想定売上規模（百万）'],
    ['受注月', '受注月'],
    ['売上開始月', '売上開始月']
  ],
  // プルダウンの選択肢に無い値も書く（プルダウンが「警告を表示」のとき。「入力を拒否」なら空欄）。false なら選択肢に無い値はいつも空欄
  writeOutOfList: true,
  subIndustryHeader: 'サブインダストリー',   // 新のこの列には、転記した行にシートの名前を入れる（プルダウンの選択肢がシートの名前だけのため）
  maxListed: 30                // 終わりの画面に出す、空欄にしたセルの数（残りはログ）
};

/** 旧フォーマットのスプレッドシートの URL を聞いて、案件を新フォーマットのシートに転記する（エディタから実行）。 */
function migrateLegacyDeals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = mtgUi_();
  const title = CRED_OPTIONS.menuTitle;
  const response = ui.prompt(title, '旧フォーマットのスプレッドシートの URL を入力してください。\n' +
    '「' + MIG_OPTIONS.sheetPrefix + '{サブインダストリー名}」シートの案件を、このスプレッドシートの同じ名前のシートに転記します（今の行は上書きします）。',
  ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const text = String(response.getResponseText() || '').trim();
  const id = mtgFileIdOf_(text);
  if (!id) throw new Error('スプレッドシートの URL として読めません。ブラウザのアドレス欄の URL をそのまま貼り付けてください。');
  let source;
  try {
    source = SpreadsheetApp.openById(id);
  } catch (_) {
    throw new Error('旧フォーマットのスプレッドシートを開けません。実行するアカウントが閲覧できるか確かめてください。');
  }

  // 先にすべて読んで確かめる（まだ何も書かない）
  const plans = Object.keys(DIFF_RULES).map(name => migPlan_(ss, source, name));
  const ready = plans.filter(p => p.rows);
  const notes = p => [p.outOfList.length ? '選択肢に無い値 ' + p.outOfList.length + 'セルもそのまま書きます' : '',
    p.blanked.length ? '「入力を拒否」のプルダウンの選択肢に無い値 ' + p.blanked.length + 'セルは空欄にします' : ''].filter(Boolean);
  const lines = plans.map(p => '・' + p.name + '：' + (p.error || p.rows.length + '件' + (notes(p).length ? '（' + notes(p).join('。') + '）' : '')));
  if (!ready.length) {
    ui.alert(title, '転記できるシートがありません。\n\n' + lines.join('\n'), ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert(title, '「' + source.getName() + '」から、次のとおり転記します。転記するシートの今の行は消えます（上書き）。\n\n' +
    lines.join('\n') + '\n\n転記したあと、差分追跡の比較基準と会議サマリの基準を記録し直します。よろしいですか？', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    ready.forEach(p => migWrite_(p));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  // 転記を変更として記録しないよう、比較基準を今の値に取り直し、会議サマリの基準を記録する
  setupDiffTracking();
  const names = ready.map(p => p.name);
  const copy = mtgRecordBaseline_(ss, names, new Date());

  const cells = key => [].concat.apply([], ready.map(p => p[key].map(b => p.name + '!' + b.a1 + '（' + b.header + '）：' + b.value)));
  const listed = list => list.slice(0, MIG_OPTIONS.maxListed).join('\n') +
    (list.length > MIG_OPTIONS.maxListed ? '\n…ほか ' + (list.length - MIG_OPTIONS.maxListed) + 'セル（実行ログに全部あります）' : '');
  const outOfList = cells('outOfList');
  const blanked = cells('blanked');
  outOfList.forEach(line => console.log('選択肢に無いがそのまま書いたセル ' + line));
  blanked.forEach(line => console.log('空欄にしたセル ' + line));
  const message = '転記しました。\n\n' + lines.join('\n') + '\n\n' +
    '差分追跡の比較基準と、会議サマリの基準（「' + copy.title + '」）を記録しました。' +
    (outOfList.length ? '\n\nプルダウンの選択肢に無いが、そのまま書いたセル（' + outOfList.length + 'セル。シートでは警告の印が付きます）：\n' + listed(outOfList) +
      '\n必要なら、得意先マスタ・ステータスの定義に足すか、シートで選択肢の値に直してください。' : '') +
    (blanked.length ? '\n\n「入力を拒否」のプルダウンの選択肢に無いため空欄にしたセル（' + blanked.length + 'セル）：\n' + listed(blanked) +
      '\n得意先マスタ・ステータスの定義に足すか、その列のプルダウンを「警告を表示」にしてから、もう一度実行すると入ります（何度実行しても大丈夫です）。' : '');
  console.log(message);
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * 1つのシートの転記の計画（読むだけ）。
 * {name, error} か {name, sheet, rule, width, formulaCols: 書かない列（1から）, rows: 書く値（新の1列目から width 列）,
 *  outOfList: 選択肢に無いがそのまま書くセル [{a1, header, value}], blanked: 「入力を拒否」のプルダウンの選択肢に無いので空欄にするセル}
 */
function migPlan_(ss, source, name) {
  const o = MIG_OPTIONS;
  const rule = DIFF_RULES[name];
  const sheet = ss.getSheetByName(name);
  if (!sheet) return {name, error: 'このスプレッドシートにシートがありません'};
  const old = source.getSheetByName(o.sheetPrefix + name);
  if (!old) return {name, error: '旧のスプレッドシートに「' + o.sheetPrefix + name + '」シートがありません（転記しません）'};

  // 新：見出しの列と、数式の列（データの1行目に数式があれば、その列は書かない・消さない）
  const firstRow = rule.headerRow + 1;
  // 書く・消す列：追跡範囲（シートの中まで）と、値のある最後の列の広いほう
  const width = Math.max(1, sheet.getLastColumn(), ...rule.ranges.map(a1 => diffRuleRange_(sheet, a1)).filter(Boolean).map(range => range.getLastColumn()));
  const headers = sheet.getRange(rule.headerRow, 1, 1, width).getDisplayValues()[0].map(credNormalize_);
  const formulas = sheet.getRange(firstRow, 1, 1, width).getFormulas()[0];
  const formulaCols = [];
  formulas.forEach((f, i) => { if (String(f).trim()) formulaCols.push(i + 1); });
  const colOf = label => headers.indexOf(credNormalize_(label)) + 1;

  // 旧：見出し（2行。下の行を優先）
  const oldWidth = Math.max(old.getLastColumn(), 1);
  const oldHeads = o.headerRows.map(r => old.getRange(r, 1, 1, oldWidth).getDisplayValues()[0]);
  const oldCol = label => {
    for (let k = oldHeads.length - 1; k >= 0; k--) {
      const i = oldHeads[k].map(credNormalize_).indexOf(credNormalize_(label));
      if (i >= 0) return i + 1;
    }
    return 0;
  };
  const pairs = o.columns.map(([from, to]) => ({from: oldCol(from), to: colOf(to), header: to}))
    .filter(p => p.from && p.to && formulaCols.indexOf(p.to) < 0);
  const customerFrom = oldCol('得意先');
  const projectFrom = oldCol('案件名');
  if (!customerFrom && !projectFrom) return {name, error: '旧のシートに「得意先」「案件名」の見出しがありません'};

  // 旧のデータ（得意先か案件名のある行）
  const count = old.getLastRow() - o.dataStartRow + 1;
  const values = count > 0 ? old.getRange(o.dataStartRow, 1, count, oldWidth).getValues() : [];
  const filled = v => String(v == null ? '' : v).trim() !== '';
  const deals = values.filter(line => (customerFrom && filled(line[customerFrom - 1])) || (projectFrom && filled(line[projectFrom - 1])));
  const capacity = sheet.getMaxRows() - rule.headerRow;
  if (deals.length > capacity) {
    return {name, error: '行が足りません（' + deals.length + '件に対して ' + capacity + '行）。エディタで addDealRows() を実行して行を足してから、もう一度実行してください'};
  }

  // プルダウンの選択肢と、無効な値を拒否するか（「入力を拒否」なら選択肢に無い値は書けないので空欄。「警告を表示」なら writeOutOfList のときそのまま書く）
  const options = {};
  const strict = {};
  pairs.forEach(p => {
    const validation = sheet.getRange(firstRow, p.to).getDataValidation();
    options[p.to] = mtgListOptions_(validation);
    strict[p.to] = !!validation && !validation.getAllowInvalid();
  });
  const subCol = colOf(o.subIndustryHeader);
  const outOfList = [];
  const blanked = [];
  const rows = deals.map((line, n) => {
    const out = new Array(width).fill('');
    if (subCol && formulaCols.indexOf(subCol) < 0) out[subCol - 1] = credText_(name);
    pairs.forEach(p => {
      let v = line[p.from - 1];
      if (v == null || v === '') return;
      const list = options[p.to];
      if (list && list.map(credNormalize_).indexOf(credNormalize_(String(v))) < 0) {
        const cell = {a1: diffA1_(firstRow + n, p.to), header: p.header, value: String(v)};
        if (strict[p.to] || !o.writeOutOfList) {
          blanked.push(cell);
          return;
        }
        outOfList.push(cell);
      }
      out[p.to - 1] = typeof v === 'string' ? credText_(v) : v;   // 文字は数式にならないように
    });
    return out;
  });
  return {name, sheet, rule, width, firstRow, formulaCols, rows, outOfList, blanked};
}

/** 計画どおりに書く（ドキュメントロックの中で呼ぶ）。数式の列を除いて、データの行（firstRow から最後の行まで）を消してから書く。 */
function migWrite_(plan) {
  const {sheet, width, firstRow, formulaCols, rows} = plan;
  const total = sheet.getMaxRows() - firstRow + 1;
  // 数式の列をよけて、続いた列ごとに書く
  const segments = [];
  for (let c = 1; c <= width; c++) {
    if (formulaCols.indexOf(c) >= 0) continue;
    const last = segments[segments.length - 1];
    if (last && last.end === c - 1) last.end = c;
    else segments.push({start: c, end: c});
  }
  segments.forEach(({start, end}) => {
    const cols = end - start + 1;
    sheet.getRange(firstRow, start, total, cols).clearContent();
    if (rows.length) sheet.getRange(firstRow, start, rows.length, cols).setValues(rows.map(line => line.slice(start - 1, end)));
  });
}
