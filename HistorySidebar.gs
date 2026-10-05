/**
 * 履歴サイドバー
 * CredentialHistory.gs・Changehistory.gs・差分追跡スクリプト（test.gs）と同じプロジェクトに置くファイル。
 * 画面は HTMLファイル「HistorySidebarView」。
 *
 * 仕組み
 * - 利用者がメニュー「🟩RXビジネスMTG用」（CRED_OPTIONS.menuTitle）＞「履歴サイドバーを開く」で開く（閉じるまで出たまま）。
 *   開いているシートが対象（DIFF_RULES のシート）でなければ開かずに知らせる。
 *   開いたときの自動表示（前の版の autoOpenHistorySidebar）はやめた。トリガーが残っていれば removeHistorySidebarAutoOpen() で外す
 * - サイドバーから呼ぶ関数は、開いた本人の権限で動く。まだ承認していない人はサイドバーに「承認する」ボタンを出す。
 *   ボタンは承認用のウェブアプリ（doGet。HS_OPTIONS.authorizeUrl）を新しいタブで開き、そこで Google の承認画面が出る
 *   （サイドバーの中からは承認画面を出せないため）。URL が未設定のときはメニューでの承認を案内する
 * - サイドバーは短い間隔で「今選ばれている行」を問い合わせ、行が変わったら
 *   その行の得意先のクレデンシャル・オファリングと変更履歴に表示を切り替える
 *   （画面はタブで「クレデンシャル」「変更履歴」を切り替える。変更履歴は期間・項目で絞り込める）
 * - 履歴は開いたときに全得意先ぶんをまとめて読み込んでおく（切り替えのたびに読まない）。
 *   その後は1分ごと、または「更新」ボタンで読み直す
 * - 登録・編集や詳しい差分は、サイドバーのボタンから各モーダルを開いて行う
 * - 得意先名の横の「アカウントプラン」は、HS_OPTIONS.accountPlan のシート（E列：得意先、M列：URL）から
 *   同じ得意先の URL を探して開くリンク（全角・半角と空白の違いは無視して照合する）
 */

const HS_OPTIONS = {
  template: 'HistorySidebarView',   // HTMLファイル名（.gs と同じ名前は付けられないため別名）
  title: '得意先別の履歴情報',   // サイドバー上部に出る名称
  eventsPerCustomer: 30,   // サイドバーに読み込む変更の記録の件数（シート・得意先ごと・新しい順）。それより古いものはモーダルで見る
  textLimit: 200,          // 変更前・変更後の表示文字数（それ以上はモーダルで確認）
  cacheSeconds: 600,       // 得意先列の位置を覚えておく秒数
  // 前の版の自動表示（やめた）。残っているトリガー・設定を外すときに使う
  legacyAutoOpenHandler: 'autoOpenHistorySidebar',   // 開いたときに自動表示していたトリガーの関数名
  legacyAutoOffPrefix: 'HS_AUTO_OFF_',               // 自動表示をオフにした人（メールアドレスごと）
  // 承認用のウェブアプリの URL（「デプロイ」→「ウェブアプリ」で作った …/exec。docs/OPERATIONS.md）。
  // 空のときは、サイドバーにボタンを出さず、メニュー（CRED_OPTIONS.menuTitle）での承認を案内する
  authorizeUrl: '',
  // 得意先名の横の「アカウントプラン」ボタンが開く URL の読み込み元
  accountPlan: {
    sheet: '33シナリオ攻略先リスト',
    customerColumn: 'E',   // 得意先名
    urlColumn: 'M',        // アカウントプランの URL（コード.gs の作成処理が入れる列）
    firstRow: 2            // 1行目は見出し
  }
};

/* ---------------- 開く ---------------- */

/**
 * メニュー「🟩RXビジネスMTG用」＞「履歴サイドバーを開く」。
 * 開いているシートが対象（DIFF_RULES のシート）でなければ開かず、対象のシートを知らせる。開いたら true。
 */
function openHistorySidebar() {
  if (!hsIsTargetSheet_(SpreadsheetApp.getActiveSheet())) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(CRED_OPTIONS.menuTitle, hsTargetSheetMessage_(), ui.ButtonSet.OK);
    return false;
  }
  hsShowSidebar_();
  return true;
}

/** 開いているシートが、履歴を出すシート（DIFF_RULES のシート）か。 */
function hsIsTargetSheet_(sheet) {
  return !!sheet && Object.prototype.hasOwnProperty.call(DIFF_RULES, sheet.getName());
}

/** 対象外のシートで開こうとしたときの知らせ。 */
function hsTargetSheetMessage_() {
  return 'このシートでは開けません。' + hsTargetSheetsLabel_() + 'のシートで開いてください。';
}

/** 履歴を出すシートの名前を知らせる文に入れる形（「スーパー・GMS」「コンビニ」）。 */
function hsTargetSheetsLabel_() {
  return '「' + Object.keys(DIFF_RULES).join('」「') + '」';
}

function hsShowSidebar_() {
  const template = HtmlService.createTemplateFromFile(HS_OPTIONS.template);
  template.authorizeUrl = hsAuthorizeUrl_();
  template.menuTitle = CRED_OPTIONS.menuTitle;   // 承認の案内に出すメニューの名前
  const html = template.evaluate().setTitle(HS_OPTIONS.title);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ---------------- 権限の承認（ウェブアプリ） ---------------- */

/** 承認用のウェブアプリの URL。https://script.google.com/…/exec の形でなければ使わない（空を返す）。 */
function hsAuthorizeUrl_() {
  const url = String(HS_OPTIONS.authorizeUrl || '').trim();
  return /^https:\/\/script\.google\.com\/\S+\/exec$/.test(url) ? url : '';
}

/**
 * 承認用のウェブアプリ（サイドバーの「承認する」ボタンが開く）。
 * 「ウェブアプリにアクセスしているユーザー」として実行するようにデプロイしておくと、
 * まだ承認していない人には Google の承認画面が出て、承認が済むとこのページが表示される。
 * 承認の状態を表示するだけで、シートは読み書きしない。
 */
function doGet() {
  const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  // 承認画面で一部の権限のチェックを外した人には、もう一度承認するためのリンクを出す
  const missing = info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED;
  const template = HtmlService.createTemplateFromFile('AuthorizeView');
  template.status = missing ? 'missing' : 'done';
  template.retryUrl = missing ? (info.getAuthorizationUrl() || '') : '';
  template.menuTitle = CRED_OPTIONS.menuTitle;
  return template.evaluate().setTitle(CRED_OPTIONS.menuTitle + 'の権限の承認');
}

/* ---------------- 前の版の自動表示（やめた） ---------------- */

/**
 * 前の版で、スプレッドシートを開いたときにサイドバーを自動で表示していた「起動時」トリガーの関数。
 * 今はサイドバーを開かない（利用者がメニュー「🟩RXビジネスMTG用」＞「履歴サイドバーを開く」で開く）。
 * トリガーが残っていても「関数が見つからない」エラーにならないよう名前だけを残し、動いたらそのトリガーを外す。
 */
function autoOpenHistorySidebar() {
  try {
    hsRemoveAutoOpen_();
  } catch (error) {
    console.warn('Legacy auto-open cleanup skipped: ' + (error && error.message));
  }
}

/** 管理者がエディタから実行する：前の版の自動表示のトリガーと、各自のオン・オフの設定を外す。 */
function removeHistorySidebarAutoOpen() {
  const removed = hsRemoveAutoOpen_();
  SpreadsheetApp.getActiveSpreadsheet().toast(removed
    ? '前の版の自動表示のトリガーを外しました（' + removed + '件）。'
    : '自動表示のトリガーはありません。', '履歴サイドバー', 8);
}

/**
 * 前の版の自動表示のトリガー（実行した人が所有するもの）と、各自のオフの設定（ドキュメントのプロパティ）を外す。
 * 外したトリガーの数を返す。
 */
function hsRemoveAutoOpen_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers().filter(t =>
    t.getHandlerFunction() === HS_OPTIONS.legacyAutoOpenHandler && t.getTriggerSourceId() === ss.getId());
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  const props = PropertiesService.getDocumentProperties();
  props.getKeys().filter(key => key.indexOf(HS_OPTIONS.legacyAutoOffPrefix) === 0).forEach(key => props.deleteProperty(key));
  return triggers.length;
}

/** サイドバーのボタンから各モーダルを開く。 */
/** focusId：選んだクレデンシャル・オファリングを編集する状態で開く（省略可）。 */
function openCredentialFromSidebar(sheetName, row, focusId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(sheetName));
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  credShowDialog_(sheet, Number(row), credActiveEmail_(), {focusId: String(focusId || '')});
}

/** focusEventId：選んだ変更の位置で開く（省略可）。 */
function openChangeHistoryFromSidebar(sheetName, row, focusEventId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(sheetName));
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  chgShowDialog_(sheet, Number(row), credActiveEmail_(), {focusEventId: String(focusEventId || '')});
}

/* ---------------- 選択行（短い間隔で呼ばれる） ---------------- */

/**
 * 今選ばれている行を返す。last と同じ行なら得意先を読まずに {same: true} を返す（速くするため）。
 * 戻り値 {sheet, row, customer, key} / {sheet, row, reason: 'sheet' | 'header'} / {same: true}
 */
function getSidebarSelection(last) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const name = sheet.getName();
  const cell = sheet.getCurrentCell() || sheet.getActiveRange();
  const row = cell ? cell.getRow() : 0;

  if (last && last.sheet === name && Number(last.row) === row && !last.force) {
    return {sheet: name, row, same: true};
  }

  const rule = DIFF_RULES[name];
  if (!rule) return {sheet: name, row, reason: 'sheet'};
  if (row <= rule.headerRow) return {sheet: name, row, reason: 'header'};

  const customer = sheet.getRange(row, hsCustomerColumn_(sheet, rule)).getDisplayValue().trim();
  return {sheet: name, row, customer, key: credNormalize_(customer)};
}

/** 得意先列の位置。見出しの読み込みを毎回しないよう、しばらく覚えておく。 */
function hsCustomerColumn_(sheet, rule) {
  const cache = CacheService.getUserCache();
  const key = 'HS_CUSTCOL_' + sheet.getSheetId();
  const hit = cache.get(key);
  if (hit) return Number(hit);
  const column = credColumnByHeader_(sheet, rule, CRED_OPTIONS.customerHeader);
  cache.put(key, String(column), HS_OPTIONS.cacheSeconds);
  return column;
}

/* ---------------- 履歴の一括読み込み ---------------- */

/**
 * 全得意先ぶんの履歴をまとめて返す（読み取りのみ）。
 * credentials: {得意先キー: [記録]}, changes: {シート名: {得意先キー: [変更]}}
 */
function getSidebarBundle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = ss.getSpreadsheetTimeZone();
  const kindField = CRED_OPTIONS.fields.find(f => f.key === CRED_OPTIONS.linking.kindKey);
  const changeIndex = hsChangeIndex_(ss);
  return {
    sheets: Object.keys(DIFF_RULES),
    credentials: hsCredentialIndex_(ss),
    accountPlans: hsAccountPlanIndex_(ss),   // {得意先キー: URL}
    accountPlanSheet: HS_OPTIONS.accountPlan.sheet,
    changes: changeIndex.changes,
    changesMore: changeIndex.more,   // {シート名: {得意先キー: true}}：上限を超えた古い記録がある
    eventsPerCustomer: HS_OPTIONS.eventsPerCustomer,
    linking: CRED_OPTIONS.linking,
    schedule: CRED_OPTIONS.schedule,
    kindOptions: kindField ? kindField.options : [],
    today: Utilities.formatDate(new Date(), timezone, 'yyyy/MM/dd'),
    loadedAt: Utilities.formatDate(new Date(), timezone, 'HH:mm')
  };
}

/**
 * アカウントプランの URL を得意先ごとに返す {得意先キー: URL}。シートが無ければ空。
 * URL がそのまま入っていない行は、HYPERLINK の数式とセルのリンクも見る。
 * http(s) の URL だけを使う（それ以外はリンクにしない）。同じ得意先が複数あれば上の行を使う。
 */
function hsAccountPlanIndex_(ss) {
  const o = HS_OPTIONS.accountPlan;
  const out = {};
  const sheet = ss.getSheetByName(o.sheet);
  if (!sheet) return out;
  const last = sheet.getLastRow();
  if (last < o.firstRow) return out;

  const customers = sheet.getRange(o.customerColumn + o.firstRow + ':' + o.customerColumn + last).getDisplayValues();
  const urlRange = sheet.getRange(o.urlColumn + o.firstRow + ':' + o.urlColumn + last);
  const shown = urlRange.getDisplayValues();
  let formulas = null;
  let links = null;
  customers.forEach((row, i) => {
    const key = credNormalize_(row[0]);
    if (!key || out[key]) return;
    let url = shown[i][0].trim();
    if (!hsIsUrl_(url)) {
      if (!formulas) {   // 表示が URL でない行があるときだけ、数式とリンクを1回で読む
        formulas = urlRange.getFormulas();
        links = urlRange.getRichTextValues();
      }
      const m = /^=HYPERLINK\(\s*"([^"]+)"/i.exec(formulas[i][0] || '');
      url = m ? m[1] : ((links[i][0] && links[i][0].getLinkUrl()) || '');
    }
    if (hsIsUrl_(url)) out[key] = url.trim();
  });
  return out;
}

function hsIsUrl_(text) {
  return /^https?:\/\/\S+$/i.test(String(text || '').trim());
}

function hsCredentialIndex_(ss) {
  const out = {};
  const log = ss.getSheetByName(CRED_OPTIONS.logSheet);
  if (!log || log.getLastRow() < 2) return out;

  const values = log.getRange(1, 1, log.getLastRow(), log.getLastColumn()).getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const layout = credLogLayout_();
  const labels = layout.map(c => c.label);
  const cols = {};
  layout.forEach(c => { const i = headers.indexOf(c.label); if (i >= 0) cols[c.key] = i + 1; });
  Object.entries(CRED_OPTIONS.legacyColumns || {}).forEach(([label, fieldKey]) => {
    const i = headers.indexOf(label);
    if (i >= 0 && labels.indexOf(label) < 0) cols['legacy.' + fieldKey] = i + 1;
  });
  if (!cols.id || !cols.customer) return out;

  const keep = ['kind', 'name', 'plan', 'date', 'person', 'title', 'owner', 'ownerDept', 'note'];
  for (let r = 1; r < values.length; r++) {
    const e = credRowToEntry_(values[r], cols);
    if (e.deleted || !e.id || !e.customer) continue;
    const values2 = {};
    keep.forEach(k => { values2[k] = e.fields[k] || ''; });
    const key = credNormalize_(e.customer);
    (out[key] = out[key] || []).push({id: e.id, linkId: e.linkId, at: e.at, values: values2});
  }
  return out;
}

/** 戻り値 {changes: {シート名: {得意先キー: [変更]}}, more: {シート名: {得意先キー: true}}} */
function hsChangeIndex_(ss) {
  const out = {};
  const more = {};
  const snap = ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet);
  if (!snap || snap.getLastRow() < 2) return {changes: out, more};

  const values = snap.getRange(1, 1, snap.getLastRow(), snap.getLastColumn()).getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const col = label => headers.indexOf(label);
  const iAt = col('記録日時'), iEvent = col('イベントID'), iEditor = col('編集者メールアドレス');
  const iSheet = col('シート'), iRow = col('行番号');
  const iCustomer = col(diffCleanHeaderText_(CRED_OPTIONS.customerHeader));
  if (iCustomer < 0) return {changes: out, more};

  // シート＋得意先ごとに、新しい方から決まった件数だけ残す（シート名を変える前の記録は、今のシート名にそろえる）
  const current = diffSheetNameResolver_();
  const groups = new Map();
  for (let r = 1; r < values.length; r++) {
    const v = values[r];
    const sheetName = current(v[iSheet]);
    if (!sheetName || !DIFF_RULES[sheetName]) continue;
    const key = sheetName + '\u0001' + credNormalize_(v[iCustomer]);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(v);
  }
  // 行の順ではなく記録日時の順にそろえる（過去の日時の記録があとから足された場合も、新しい方から残るように）
  groups.forEach(list => list.sort((a, b) => (a[iAt] < b[iAt] ? -1 : a[iAt] > b[iAt] ? 1 : 0)));
  const needed = new Set();
  groups.forEach((list, key) => {
    const extra = list.length - HS_OPTIONS.eventsPerCustomer;
    if (extra > 0) {
      list.splice(0, extra);
      const [sheetName, customerKey] = key.split('\u0001');
      (more[sheetName] = more[sheetName] || {})[customerKey] = true;
    }
    list.forEach(v => needed.add(v[iEvent]));
  });

  // 変更前・変更後（変更履歴_差分 の B:I）
  const records = new Map();
  const log = ss.getSheetByName(DIFF_OPTIONS.logSheet);
  if (log && log.getLastRow() >= 2) {
    log.getRange(2, 2, log.getLastRow() - 1, 8).getDisplayValues().forEach(r => {
      if (!needed.has(r[0])) return;
      if (!records.has(r[0])) records.set(r[0], []);
      records.get(r[0]).push({sheet: current(r[2]), cell: r[3], column: r[4], before: r[6], after: r[7]});
    });
  }

  const cut = text => {
    text = text === CHG_OPTIONS.emptyMark ? '' : String(text || '').replace(/\s+/g, ' ');
    return text.length > HS_OPTIONS.textLimit ? text.slice(0, HS_OPTIONS.textLimit) + '…' : text;
  };

  groups.forEach((list, key) => {
    const [sheetName, customerKey] = key.split('\u0001');
    out[sheetName] = out[sheetName] || {};
    out[sheetName][customerKey] = list.slice().reverse().map(v => {
      const rowNumber = Number(v[iRow]);
      const changes = (records.get(v[iEvent]) || [])
        .filter(rec => rec.sheet === sheetName && chgRecordOfRow_(rec.cell, rowNumber))
        .map(rec => ({column: rec.column, before: cut(rec.before), after: cut(rec.after)}));
      return {
        eventId: v[iEvent],
        at: v[iAt],
        editor: String(v[iEditor] || '').split('@')[0],
        type: String(v[iEvent]).split('-')[0],
        changes
      };
    });
  });
  return {changes: out, more};
}