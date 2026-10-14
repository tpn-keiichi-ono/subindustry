/**
 * 会議サマリ（追跡シートごとの会議資料の PDF）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs・HistorySidebar.gs と同じプロジェクトに置くファイル。
 *
 * 目的
 * - 追跡シート（スーパー・GMS など。DIFF_RULES のシート）ごとの推進リーダーが会議で説明に使い、本部長クラスがレビューするサマリ資料。
 *   参加者が全リストをスクロールしなくても要点（全体感・前回からの動き・停滞と失注・不備）をつかめるようにし、会議の2〜3日前に配る（事前に読む運用）。
 * - 「動いた案件・止まった案件・不備のある案件」のふるい分けに絞る。個別の案件を深く見るのは、オポチュニティキャンバスなど別の画面で行う。
 *
 * 資料の構成（A4 横。1ページずつ、社内資料のトンマナ：左上の紫のタグ・紫の大きなタイトル・リード文・紫の見出しの表・結論の帯・ページ番号）
 *   1 全体サマリー：案件ステータス別の件数・想定売上・期待値調整済（進行中・受注・失注）
 *   2 今回の論点：論点 ▶ 状況 ▶ 確認したいこと（受注・失注・ステータスが動いた案件・停滞している大型案件・要対応）
 *   3〜 前回からの進捗・変化：変更履歴_差分 から、前回以降に値が変わった案件（ステータスの動き・更新した項目・今の活動状況）
 *   〜 停滞・失注案件：最終更新日時が MTG_OPTIONS.staleDays 日以上前の進行中の案件と、前回以降に失注・クローズになった案件・削除された行
 *      （最終更新日時が空の行は、変更の記録を始めた日（変更履歴_差分 の最初の行）から動いていないとみなし、「N日以上」と出す）
 *   〜 要対応（アラート）：クレデンシャル・オファリングの予定日を過ぎても実施日が空・受注月を過ぎても進行中・必須項目が空
 *   詳しいページは、入りきらなければ続きのページに分ける（中身が無ければ作らない）。
 *
 * 仕組み
 * - 管理者がエディタで createMeetingSummaries() を実行し、作るシートを選ぶ（試すときは previewMeetingSummaries()）。
 *   シートごとに PDF をドライブのフォルダに保存する（MTG_OPTIONS.folderId。空なら、このスプレッドシートと同じフォルダの中に
 *   MTG_OPTIONS.folderName のフォルダを作る）。配る（共有する）のは手動。
 * - 「前回から」は、そのシートで前に createMeetingSummaries() で作った日時から（初回は MTG_OPTIONS.firstDays 日前から）。
 *   previewMeetingSummaries() は前回の日時を進めないので、何度でも試せる（ファイル名に「試作」が付く）。
 * - スプレッドシートは読むだけで、書き込まない。変更の記録は、変更時点スナップショット（その時の行の内容）で得意先＋案件名に結び付ける
 *   （行の追加・削除で行番号がずれても合う）。どちらの記録も追記なので、下（新しいほう）から、前回より前の記録に届くまでだけ読む。
 * - PDF は、一時的なスプレッドシートに1ページを1枚のシートとして書き、すべてのシートを「幅に合わせる」で A4 横に書き出す
 *   （どのページも同じ幅なので文字の大きさがそろい、1枚のシートが1ページになる）。書き出したら一時的なスプレッドシートはゴミ箱に移す。
 *   書き出しに UrlFetchApp を使う（外部サービスへの接続の権限）。
 */

const MTG_OPTIONS = {
  title: 'RXビジネスMTG',             // 左上のタグ（「タグ｜シート名」と出す）
  fileName: '会議サマリ',             // PDF のファイル名の先頭（会議サマリ_スーパー・GMS_2026-10-20.pdf）
  folderId: '',                      // 保存先のフォルダの ID（空なら、このスプレッドシートと同じフォルダの中に folderName を作る）
  folderName: '会議サマリ',
  folderKey: 'MTG_FOLDER_ID',         // 作ったフォルダの ID（ドキュメントのプロパティ）
  lastKeyPrefix: 'MTG_LAST_',         // シートごとの前回の作成日時（ドキュメントのプロパティ。後ろにシート名）
  firstDays: 14,                     // 初回（前回の記録が無い）は、何日前からの変化を出すか
  staleDays: 30,                     // 停滞：最終更新日時がこの日数より前の進行中の案件
  staleAlertDays: 60,                // 停滞の日数を、この日数から赤で出す
  // 読む列の見出し（改行・空白、全角・半角の違いは無視。customer・project・status はどのシートにも必要）
  columns: {
    customer: '得意先',
    project: '案件名',
    status: '案件ステータス',
    amount: '想定売上規模（百万）',
    weighted: '期待値調整済\n想定売上規模（百万）',
    orderMonth: '受注月',
    activity: '活動状況',
    issue: '活動における課題',
    owner: 'アカウント責任者'
  },
  // 必須項目：進行中の案件で空ならアラートに出す（シートに無い見出しは見ない）
  required: ['想定売上規模（百万）', '受注月', '活動状況', 'アカウント責任者'],
  // ステータスの分け方（その語を含むか）。失注を先に見る。受注は notWonWords を含むもの（受注見込み など）を除く
  lostWords: ['失注', '辞退', '見送', '中止', 'クローズ', 'ロスト', '消滅'],
  wonWords: ['受注', '成約', '契約済'],
  notWonWords: ['見込', '予定', '前', '待'],
  logChunkRows: 2000,                // 変更の記録を下から読むときの1回の行数
  tempName: '会議サマリ（作成中。終わるとゴミ箱に移します）',
  staleFocus: 3,                     // 停滞している案件のうち、想定売上の大きい何件を論点にするか（赤枠）
  topicItems: 3,                     // 今回の論点に出す案件の数（論点ごと。残りは詳しいページを案内する）
  // 今回の論点（「論点 ▶ 状況 ▶ 確認したいこと」の行）。focus は中身があれば赤枠で囲む
  topics: {
    won: {title: '①受注', ask: '受注後の立ち上げ体制と、ほかの得意先への横展開の可否'},
    lost: {title: '②失注・クローズ', ask: '失注の要因と、再アプローチの要否'},
    moved: {title: '③ステータスが\n動いた案件', ask: '次のアクションと、受注月の見通し'},
    stale: {title: '④停滞している\n大型案件', ask: '上席の同席・支援の要否', focus: true},
    alerts: {title: '⑤要対応\n（アラート）', ask: '期限を過ぎた予定の実施・受注月の見直し・未入力の項目の入力を、各責任者に依頼'}
  },
  // 色（社内資料のトンマナ：紫を基調に、薄い紫・薄いグレーの面、強調は赤枠、結論はピンクの帯）
  colors: {
    purple: '#7030A0', purpleSoft: '#EEE3F5', gray: '#F2F2F2', grayHead: '#A6A6A6', arrow: '#A6A6A6',
    red: '#E60012', highlight: '#F6D5F7', ink: '#1b1b1c', body: '#262626', muted: '#595959', faint: '#8C8C8C',
    lost: '#8C8C8C', white: '#ffffff'
  }
};

/*
 * ページの形。どのページも幅を MTG_PAGE.width にそろえ、高さは maxHeight 以内にする（A4 横を幅に合わせて書き出すと、
 * 幅 1200 に対して高さ約 830 まで1ページに入る）。表の列の幅は、ページの種類ごとに足して width になるように決める。
 */
const MTG_PAGE = {
  width: 1200,
  maxHeight: 800,
  fontSize: 11,          // 本文（幅に合わせて書き出すと、約10ポイントで印刷される）
  lineHeight: 19,        // 複数行のセルの1行ぶんの高さ
  rowHeight: 28,         // 表の1行の高さ（1行の文）
  leadLines: 2,          // リード文の行数
  tagWidth: 330,         // 左上のタグの幅（この幅以上になるまで左の列を紫にする。どのページも列の境目がこの近くに来るように幅を決める）
  barWidth: 32,          // 全体サマリーの棒の最大の長さ（文字数）
  widths: {
    overview: [220, 70, 90, 90, 90, 640],                    // 案件ステータス・件数・想定・期待値・構成比・棒
    topics: [200, 40, 100, 330, 40, 490],                    // 論点・▶・件数・状況・▶・確認したいこと
    changes: [250, 70, 70, 200, 160, 90, 360],               // 得意先／案件名・想定・期待値・ステータスの動き・更新した項目・責任者・活動状況
    stale: [230, 110, 80, 70, 70, 80, 90, 235, 235],         // 得意先／案件名・ステータス・停滞・想定・期待値・受注月・責任者・課題・活動状況
    alerts: [250, 120, 600, 80, 150]                         // 得意先／案件名・種類・内容・想定・責任者
  }
};

/* ---------------- エディタから実行する ---------------- */

/**
 * 会議サマリの PDF を作って保存する（作るシートを選ぶ）。
 * 作ったシートは、「前回から」の基準を今の日時に進める（次は今以降の変化を出す）。
 */
function createMeetingSummaries() {
  mtgRun_(true);
}

/** 試しに会議サマリの PDF を作る（保存するが、「前回から」の基準は進めない。ファイル名に「試作」が付く）。 */
function previewMeetingSummaries() {
  mtgRun_(false);
}

function mtgRun_(final) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const names = mtgAskSheets_(ss, ui, final);
  if (!names) return;
  const results = mtgCreate_(ss, names, new Date(), final);
  const lines = results.map(r => '「' + r.summary.name + '」：' + r.file.getName() + '（' + mtgCountsText_(r.summary) + '）');
  lines.push('', '保存先：' + results[0].folder.getUrl());
  lines.push(final
    ? '次に作るときは、今以降の変化を出します。配るときは、保存先のフォルダから共有してください。'
    : '試作のため、「前回から」の基準は進めていません。');
  console.log(lines.join('\n'));
  ui.alert(CRED_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/** 作るシートを聞く（空欄ならすべて。「1,3」のように複数も）。やめたら null。 */
function mtgAskSheets_(ss, ui, final) {
  const all = Object.keys(DIFF_RULES).filter(name => ss.getSheetByName(name));
  if (!all.length) throw new Error('追跡しているシート（test.gs の DIFF_RULES）が見つかりません。');
  const message = (final ? '会議サマリの PDF を作ります。' : '会議サマリの PDF を試しに作ります（「前回から」の基準は進めません）。') +
    '\n作るシートの番号を入力してください（空欄ならすべて。複数は「1,3」のように）。\n\n' +
    all.map((name, i) => (i + 1) + '. ' + name).join('\n');
  const response = ui.prompt(CRED_OPTIONS.menuTitle, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const text = credNormalize_(response.getResponseText());
  if (!text) return all;
  const picked = [];
  text.split(/[,、・]/).filter(Boolean).forEach(part => {
    const index = Number(part) - 1;
    if (!(index >= 0 && index < all.length)) throw new Error('「' + part + '」に当たるシートがありません。1〜' + all.length + ' の番号で入力してください。');
    if (picked.indexOf(all[index]) < 0) picked.push(all[index]);
  });
  return picked;
}

/** 完成した結果の件数の知らせ（「動いた 12・停滞 5・アラート 7」）。 */
function mtgCountsText_(summary) {
  return '動いた ' + summary.changes.length + '・停滞 ' + summary.stale.length + '・アラート ' + summary.alerts.length;
}

/**
 * names のシートの会議サマリを作り、PDF を保存する。final なら「前回から」の基準を now に進める。
 * 戻り値 [{summary, file, folder}]
 */
function mtgCreate_(ss, names, now, final) {
  const props = PropertiesService.getDocumentProperties();
  const plans = names.map(name => {
    const last = Number(props.getProperty(MTG_OPTIONS.lastKeyPrefix + name)) || 0;
    return {name, first: !last, since: new Date(last || now.getTime() - MTG_OPTIONS.firstDays * 86400000)};
  });
  const oldest = new Date(Math.min.apply(null, plans.map(p => p.since.getTime())));
  const context = {
    now,
    trackedSince: mtgLogStart_(ss),
    log: mtgReadLog_(ss, oldest),
    snapshots: mtgReadSnapshots_(ss, oldest),
    credentials: hsCredentialIndex_(ss)
  };
  const summaries = plans.map(p => mtgSummarize_(ss, p.name, Object.assign({since: p.since, first: p.first, final}, context)));
  const folder = mtgFolder_(ss);
  const files = mtgExportPdfs_(summaries, folder, now, final);
  if (final) plans.forEach(p => props.setProperty(MTG_OPTIONS.lastKeyPrefix + p.name, String(now.getTime())));
  return summaries.map((summary, i) => ({summary, file: files[i], folder}));
}

/* ---------------- 読む ---------------- */

/**
 * 記録のシート（変更履歴_差分・変更時点スナップショット）を下から読み、since 以降の行を返す（古い順）。
 * 記録は追記なので、サンプル（イベントID が SMP-。過去の日時で足される）以外で since より前の行に届いたら止める。
 * {index: 見出し → 列の位置（0 から）, rows}
 */
function mtgReadRecent_(sheet, since) {
  const out = {index: {}, rows: []};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const width = sheet.getLastColumn();
  sheet.getRange(1, 1, 1, width).getDisplayValues()[0].forEach((h, i) => { out.index[String(h).trim()] = i; });
  const atCol = out.index['記録日時'];
  const idCol = out.index['イベントID'];
  if (atCol == null) return out;
  const limit = since.getTime();
  const step = MTG_OPTIONS.logChunkRows;
  for (let end = sheet.getLastRow(); end >= 2; end -= step) {
    const start = Math.max(2, end - step + 1);
    const rows = sheet.getRange(start, 1, end - start + 1, width).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      const sample = idCol != null && /^SMP-/.test(String(rows[i][idCol]));
      const at = mtgTime_(rows[i][atCol]);
      if (sample) continue;
      if (at != null && at < limit) return Object.assign(out, {rows: out.rows.reverse()});
      if (at != null) out.rows.push(rows[i]);
    }
  }
  out.rows.reverse();
  return out;
}

/** 変更履歴_差分 の since 以降の記録 [{at, eventId, sheet, cell, column, type, before, after}]（古い順） */
function mtgReadLog_(ss, since) {
  const read = mtgReadRecent_(ss.getSheetByName(DIFF_OPTIONS.logSheet), since);
  const i = read.index;
  const need = ['記録日時', 'イベントID', 'シート', 'セル', '列名', '種類', '変更前', '変更後'];
  if (read.rows.length && need.some(h => i[h] == null)) {
    throw new Error('「' + DIFF_OPTIONS.logSheet + '」の見出しが今の形ではありません。エディタで compactDiffRecords() を実行してから、もう一度お試しください。');
  }
  const text = v => mtgCellText_(v);
  return read.rows.map(r => ({
    at: mtgTime_(r[i['記録日時']]),
    eventId: text(r[i['イベントID']]),
    sheet: text(r[i['シート']]),
    cell: text(r[i['セル']]),
    column: text(r[i['列名']]),
    type: text(r[i['種類']]),
    before: text(r[i['変更前']]),
    after: text(r[i['変更後']])
  }));
}

/** 変更の記録を始めた時刻（変更履歴_差分 の最初の行。記録が無ければ null）。 */
function mtgLogStart_(ss) {
  const log = ss.getSheetByName(DIFF_OPTIONS.logSheet);
  if (!log || log.getLastRow() < 2) return null;
  const headers = log.getRange(1, 1, 1, log.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
  const col = headers.indexOf('記録日時');
  return col < 0 ? null : mtgTime_(log.getRange(2, col + 1).getValue());
}

/**
 * 変更時点スナップショットの since 以降の行を、「イベントID#行番号」→ {customer, project, status} にまとめる
 * （その時の行の内容。変更の記録を得意先＋案件名に結び付けるのに使う）。
 */
function mtgReadSnapshots_(ss, since) {
  const read = mtgReadRecent_(ss.getSheetByName(DIFF_OPTIONS.rowSnapshotSheet), since);
  const i = read.index;
  const map = new Map();
  if (i['イベントID'] == null || i['行番号'] == null) return map;
  const cols = MTG_OPTIONS.columns;
  read.rows.forEach(r => {
    const content = new Map();
    if (i['行の内容'] != null) {
      diffParseSnapshotContent_(r[i['行の内容']]).forEach(([k, v]) => content.set(credNormalize_(k), v));
    }
    const pick = label => String(content.get(credNormalize_(label)) || '').trim();
    map.set(mtgCellText_(r[i['イベントID']]) + '#' + Number(r[i['行番号']]), {
      customer: pick(cols.customer) || (i['得意先'] != null ? mtgCellText_(r[i['得意先']]).trim() : ''),
      project: pick(cols.project),
      status: pick(cols.status)
    });
  });
  return map;
}

/**
 * 追跡シートの案件（得意先か案件名のある行）を読む。
 * {headers, statusOrder: ステータスのプルダウンの選択肢（無ければ空）, deals: [{row, key, customer, project, status, kind,
 *  amount, weighted, orderMonth: {y, m}|null, orderMonthText, activity, issue, owner, updatedAt: 時刻（ms）|null, missing: [見出し]}]}
 */
function mtgReadDeals_(sheet, rule) {
  const headerRow = rule.headerRow;
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, width).getDisplayValues()[0];
  const find = label => headers.findIndex(h => credNormalize_(h) === credNormalize_(label));
  const cols = {};
  Object.keys(MTG_OPTIONS.columns).forEach(key => { cols[key] = find(MTG_OPTIONS.columns[key]); });
  ['customer', 'project', 'status'].forEach(key => {
    if (cols[key] < 0) {
      throw new Error('「' + sheet.getName() + '」の ' + headerRow + '行目に「' + MTG_OPTIONS.columns[key] +
        '」の見出しがありません（MeetingSummary.gs の MTG_OPTIONS.columns を確かめてください）。');
    }
  });
  const stamp = diffNormalizeStampHeader_(DIFF_OPTIONS.stampHeader);
  const stampCol = headers.findIndex(h => diffNormalizeStampHeader_(h) === stamp);
  const required = MTG_OPTIONS.required.map(label => ({label, col: find(label)})).filter(r => r.col >= 0);

  const out = {headers, statusOrder: [], deals: []};
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return out;
  const rule0 = sheet.getRange(headerRow + 1, cols.status + 1).getDataValidation();
  if (rule0 && rule0.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    out.statusOrder = rule0.getCriteriaValues()[0].map(v => String(v).trim()).filter(Boolean);
  }
  const range = sheet.getRange(headerRow + 1, 1, count, width);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  shown.forEach((line, n) => {
    const text = key => (cols[key] >= 0 ? String(line[cols[key]]).trim() : '');
    const customer = text('customer');
    const project = text('project');
    if (!customer && !project) return;
    const status = text('status');
    out.deals.push({
      row: headerRow + 1 + n,
      key: mtgDealKey_(customer, project),
      customer, project, status,
      kind: mtgStatusKind_(status),
      amount: cols.amount >= 0 ? mtgNumber_(values[n][cols.amount]) : null,
      weighted: cols.weighted >= 0 ? mtgNumber_(values[n][cols.weighted]) : null,
      orderMonth: cols.orderMonth >= 0 ? mtgMonth_(values[n][cols.orderMonth]) : null,
      orderMonthText: text('orderMonth'),
      activity: text('activity'),
      issue: text('issue'),
      owner: text('owner'),
      updatedAt: stampCol >= 0 ? mtgTime_(values[n][stampCol]) : null,
      missing: required.filter(r => !String(line[r.col]).trim()).map(r => r.label)
    });
  });
  return out;
}

/* ---------------- まとめる ---------------- */

/**
 * 1つのシートの会議サマリの中身。context = {since, first, final, now, trackedSince, log, snapshots, credentials}
 * {name, since, first, final, now, kpi, pipeline, changes, stale, closed, alerts, unmatched}
 */
function mtgSummarize_(ss, name, context) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('「' + name + '」シートが見つかりません。');
  const table = mtgReadDeals_(sheet, DIFF_RULES[name]);
  const deals = table.deals;
  const byKey = new Map();
  deals.forEach(d => { if (!byKey.has(d.key)) byKey.set(d.key, d); });
  const active = deals.filter(d => d.kind === 'active');
  const now = context.now.getTime();

  // ② 前回からの変化（変更の記録を得意先＋案件名ごとにまとめる）
  const resolve = diffSheetNameResolver_();
  const statusLabel = credNormalize_(MTG_OPTIONS.columns.status);
  const moved = new Map();
  const removed = [];
  let unmatched = 0;
  const empty = v => (v === CHG_OPTIONS.emptyMark ? '' : String(v || '').trim());
  context.log.forEach(e => {
    if (resolve(e.sheet) !== name || e.at == null || e.at < context.since.getTime() || /^(SMP|CRD)-/.test(e.eventId)) return;
    if (e.type === '行削除') {
      const gone = mtgParseRemoved_(e.before);
      if (gone.customer || gone.project) removed.push(Object.assign(gone, {at: e.at}));
      return;
    }
    const snap = context.snapshots.get(e.eventId + '#' + mtgCellRow_(e.cell));
    if (!snap || !(snap.customer || snap.project)) { unmatched++; return; }
    const key = mtgDealKey_(snap.customer, snap.project);
    const item = moved.get(key) || {key, customer: snap.customer, project: snap.project, from: null, to: null, isNew: false, columns: [], at: 0};
    if (e.type === '行追加') {
      item.isNew = true;
    } else if (e.column && e.column !== '（行全体）') {
      if (credNormalize_(e.column) === statusLabel) {
        if (item.from === null) item.from = empty(e.before);
        item.to = empty(e.after);
      } else if (item.columns.indexOf(e.column) < 0) {
        item.columns.push(e.column);
      }
    }
    item.at = Math.max(item.at, e.at);
    moved.set(key, item);
  });
  const changes = [];
  moved.forEach(item => {
    const deal = byKey.get(item.key);
    const statusChanged = item.from !== null && credNormalize_(item.from) !== credNormalize_(item.to);
    if (!statusChanged && !item.isNew && !item.columns.length) return;
    changes.push(Object.assign(item, {
      deal,
      statusChanged,
      status: deal ? deal.status : (item.to || ''),
      toKind: mtgStatusKind_(statusChanged ? item.to : (deal ? deal.status : '')),
      amount: deal ? deal.amount : null,
      weighted: deal ? deal.weighted : null,
      owner: deal ? deal.owner : '',
      orderMonthText: deal ? deal.orderMonthText : '',
      issue: deal ? deal.issue : '',
      activity: deal ? deal.activity : ''
    }));
  });
  const kindOrder = {won: 0, lost: 1, active: 2};
  changes.sort((a, b) => (b.statusChanged - a.statusChanged) || (kindOrder[a.toKind] - kindOrder[b.toKind]) ||
    ((b.amount || 0) - (a.amount || 0)) || (b.at - a.at));

  // ③ 停滞（進行中で最終更新が古い）と、前回から失注・クローズになった案件・削除された行。
  // 最終更新日時が空の行は、変更の記録を始めた時から動いていない（日数は「それ以上」）。記録が無ければ判断しない
  const start = context.trackedSince;
  const stale = active.map(d => {
    if (d.updatedAt != null) return Object.assign({}, d, {days: Math.floor((now - d.updatedAt) / 86400000), noRecord: false});
    return start == null ? null : Object.assign({}, d, {days: Math.floor((now - start) / 86400000), noRecord: true});
  }).filter(d => d && d.days >= MTG_OPTIONS.staleDays)
    .sort((a, b) => ((b.amount || 0) - (a.amount || 0)) || (b.days - a.days));
  const closed = changes.filter(c => c.statusChanged && c.toKind === 'lost')
    .map(c => ({customer: c.customer, project: c.project, from: c.from, to: c.to, at: c.at, amount: c.amount, weighted: c.weighted,
      owner: c.owner, orderMonthText: c.orderMonthText, issue: c.issue, activity: c.activity, removed: false}))
    .concat(removed.map(r => ({customer: r.customer, project: r.project, from: r.status, to: '', at: r.at,
      amount: r.amount, weighted: null, owner: r.owner, orderMonthText: '', issue: '', activity: '', removed: true})))
    .sort((a, b) => b.at - a.at);

  // ④ アラート
  const alerts = mtgAlerts_(deals, active, context);

  return {
    name,
    since: context.since,
    first: context.first,
    final: context.final,
    now: context.now,
    kpi: {
      deals: deals.length,
      active: active.length,
      won: deals.filter(d => d.kind === 'won').length,
      lost: deals.filter(d => d.kind === 'lost').length,
      amount: mtgSum_(active, 'amount'),
      weighted: mtgSum_(active, 'weighted'),
      changed: changes.length,
      statusChanged: changes.filter(c => c.statusChanged).length
    },
    pipeline: mtgPipeline_(deals, table.statusOrder),
    changes,
    stale,
    closed,
    alerts,
    unmatched
  };
}

/**
 * アラート [{type, customer, project, text, amount, owner}]：予定日超過（クレデンシャル・オファリング）→ 受注月超過 → 必須項目が空。
 * 予定日超過は得意先ごと（このシートに案件のある得意先）。それ以外は進行中の案件ごと。
 */
function mtgAlerts_(deals, active, context) {
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(context.now, tz, 'yyyy/MM/dd');
  const out = [];
  const customers = new Map();
  deals.forEach(d => { if (d.customer && !customers.has(credNormalize_(d.customer))) customers.set(credNormalize_(d.customer), d.customer); });
  customers.forEach((customer, key) => {
    (context.credentials[key] || []).forEach(c => {
      const plan = mtgDateText_(c.values.plan);
      if (!plan || mtgDateText_(c.values.date) || plan >= today) return;
      const days = Math.round((mtgDay_(today) - mtgDay_(plan)) / 86400000);
      out.push({type: '予定超過', customer, project: '', owner: c.values.owner || '',
        text: (c.values.kind || '予定') + '「' + c.values.name + '」の予定日（' + plan.slice(5) + '）から ' + days + '日、実施日が空',
        amount: null, days});
    });
  });
  const month = Utilities.formatDate(context.now, tz, 'yyyy-MM').split('-').map(Number);
  const current = month[0] * 12 + month[1];
  const order = [];
  active.forEach(d => {
    if (d.orderMonth && d.orderMonth.y * 12 + d.orderMonth.m < current) {
      order.push({type: '受注月超過', customer: d.customer, project: d.project, amount: d.amount, owner: d.owner,
        text: '受注月 ' + d.orderMonth.y + '/' + d.orderMonth.m + ' を過ぎても「' + (d.status || '未設定') + '」'});
    }
  });
  const missing = active.filter(d => d.missing.length).map(d => ({type: '未入力', customer: d.customer, project: d.project,
    amount: d.amount, owner: d.owner,
    text: '空：' + d.missing.map(label => label.replace(/\s+/g, '')).join('・')}));
  const byAmount = (a, b) => (b.amount || 0) - (a.amount || 0);
  out.sort((a, b) => b.days - a.days);
  return out.concat(order.sort(byAmount), missing.sort(byAmount));
}

/** ステータスごとの件数・金額 [{status, kind, count, amount, weighted}]（プルダウンの順。進行中 → 受注 → 失注） */
function mtgPipeline_(deals, statusOrder) {
  const map = new Map();
  const order = statusOrder.slice();
  deals.forEach(d => {
    const label = d.status || '（未設定）';
    if (order.indexOf(label) < 0) order.push(label);
    const item = map.get(label) || {status: label, kind: d.kind, count: 0, amount: 0, weighted: 0};
    item.count++;
    item.amount += d.amount || 0;
    item.weighted += d.weighted || 0;
    map.set(label, item);
  });
  const rank = {active: 0, won: 1, lost: 2};
  return order.filter(label => map.has(label)).map(label => map.get(label))
    .sort((a, b) => rank[a.kind] - rank[b.kind]);   // 同じ種類の中はプルダウンの順（安定ソート）
}

/** 削除した行の「変更前」（「見出し: 値」の行）から {customer, project, status, amount, owner} */
function mtgParseRemoved_(text) {
  const values = new Map();
  String(text || '').split('\n').forEach(line => {
    const m = /^([^:：]+)[:：]\s?(.*)$/.exec(line);
    if (m) values.set(credNormalize_(m[1]), m[2].trim());
  });
  const cols = MTG_OPTIONS.columns;
  const get = label => values.get(credNormalize_(label)) || '';
  return {customer: get(cols.customer), project: get(cols.project), status: get(cols.status),
    amount: mtgNumber_(get(cols.amount)), owner: get(cols.owner)};
}

/* ---------------- 小さな補助 ---------------- */

/** 得意先＋案件名のキー（全角・半角と空白の違いは無視）。 */
function mtgDealKey_(customer, project) {
  return credNormalize_(customer) + '\u0001' + credNormalize_(project);
}

/** ステータスの種類：'won'（受注）・'lost'（失注・クローズなど）・'active'（進行中。空も） */
function mtgStatusKind_(status) {
  const text = credNormalize_(status);
  if (!text) return 'active';
  const has = words => words.some(w => text.indexOf(credNormalize_(w)) >= 0);
  if (has(MTG_OPTIONS.lostWords)) return 'lost';
  if (has(MTG_OPTIONS.wonWords) && !has(MTG_OPTIONS.notWonWords)) return 'won';
  return 'active';
}

/** セルの値の時刻（ms）。日時の値か「yyyy/MM/dd HH:mm」の文字。読めなければ null。 */
function mtgTime_(value) {
  if (value && typeof value.getTime === 'function') return value.getTime();
  const m = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(String(value || '').trim());
  if (!m) return null;
  const text = m[1] + '/' + m[2].padStart(2, '0') + '/' + m[3].padStart(2, '0') + ' ' +
    (m[4] || '0').padStart(2, '0') + ':' + (m[5] || '00') + ':' + (m[6] || '00');
  return Utilities.parseDate(text, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss').getTime();
}

/** 日付の文字（「2026/10/01」の形。読めなければ空）。 */
function mtgDateText_(value) {
  const m = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/.exec(String(value || '').trim());
  return m ? m[1] + '/' + m[2].padStart(2, '0') + '/' + m[3].padStart(2, '0') : '';
}

/** 「2026/10/01」の日の ms（日数の差を出す用。時差の影響を受けないよう UTC で数える）。 */
function mtgDay_(text) {
  const p = text.split('/').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2]);
}

/** 数（「1,234.5」「12百万」も読む）。空・読めなければ null。 */
function mtgNumber_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : null;
  const text = String(value == null ? '' : value).normalize('NFKC').replace(/[,\s]|百万|円/g, '');
  return /^-?\d+(\.\d+)?$/.test(text) ? Number(text) : null;
}

/** 受注月 {y, m}（日付の値・「2026/10」「2026年10月」「2026-10-01」）。読めなければ null。 */
function mtgMonth_(value) {
  if (value && typeof value.getTime === 'function') {
    const p = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM').split('-').map(Number);
    return {y: p[0], m: p[1]};
  }
  const m = /(\d{4})\s*[\/\-年.]\s*(\d{1,2})/.exec(String(value || '').normalize('NFKC'));
  return m && Number(m[2]) >= 1 && Number(m[2]) <= 12 ? {y: Number(m[1]), m: Number(m[2])} : null;
}

/** 「F12」「12行」「12行（削除前の行番号）」の行番号（読めなければ 0）。 */
function mtgCellRow_(cell) {
  const m = /^[A-Z]+(\d+)$/.exec(String(cell)) || /^(\d+)行/.exec(String(cell));
  return m ? Number(m[1]) : 0;
}

/** セルの値を文字に（日時は記録の形）。 */
function mtgCellText_(value) {
  if (value && typeof value.getTime === 'function') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), DIFF_OPTIONS.dateFormat);
  }
  return String(value == null ? '' : value);
}

function mtgSum_(list, key) {
  return list.reduce((sum, d) => sum + (d[key] || 0), 0);
}

/** 金額（百万）の表示：「1,234.5」（整数なら小数なし）。null は「—」。 */
function mtgAmount_(value) {
  if (value == null) return '—';
  const rounded = Math.round(value * 10) / 10;
  const parts = String(Math.abs(rounded)).split('.');
  return (rounded < 0 ? '-' : '') + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] ? '.' + parts[1] : '');
}

/** 得意先／案件名 */
function mtgDealLabel_(item) {
  return [item.customer, item.project].filter(Boolean).join('／');
}

/**
 * 列の幅に収まるように文字を切る（収まらなければ末尾を「…」に）。全角は1文字、半角は0.55文字として数える。
 * px：列の幅、pt：文字の大きさ、lines：何行まで（折り返すセル。改行で分けた行ごとに数える）。
 */
function mtgFit_(text, px, pt, lines) {
  const value = String(text == null ? '' : text);
  if ((lines || 1) > 1) {
    return mtgLines_(value, px, pt, lines).join('\n');
  }
  const flat = value.replace(/[ \t\r\n]+/g, ' ').trim();   // 全角の空白（区切りに使う）は残す
  const limit = (px - 10) / ((pt || MTG_PAGE.fontSize) * 96 / 72);
  let used = 0;
  const chars = Array.from(flat);
  for (let i = 0; i < chars.length; i++) {
    used += /[\u0000-ÿ｡-ﾟ]/.test(chars[i]) ? 0.55 : 1;
    if (used > limit) return chars.slice(0, Math.max(i - 1, 0)).join('') + '…';
  }
  return flat;
}

/**
 * 文を、幅 px に収まる行に分ける（改行はそのまま行の区切り）。max 行を超えたら、最後の行を「…」の分をあけて切る。
 * 折り返しの位置は文字の幅の見積もりなので、本物の折り返しより早めに分ける（余った行が隠れないように）。
 */
function mtgLines_(text, px, pt, max) {
  const limit = (px - 12) / ((pt || MTG_PAGE.fontSize) * 96 / 72) * 0.88;
  const width = ch => (/[\u0000-\u00ff\uff61-\uff9f]/.test(ch) ? 0.6 : 1);
  const out = [];
  String(text == null ? '' : text).split('\n').forEach(part => {
    let line = '';
    let used = 0;
    Array.from(part.replace(/[ \t]+/g, ' ').trim()).forEach(ch => {
      if (used + width(ch) > limit && line) {
        out.push(line);
        line = '';
        used = 0;
      }
      line += ch;
      used += width(ch);
    });
    out.push(line);
  });
  if (out.length <= max) return out;
  const kept = out.slice(0, max);
  let last = '';
  let used = 0;
  Array.from(kept[max - 1]).some(ch => {
    if (used + width(ch) > limit - 1) return true;
    last += ch;
    used += width(ch);
    return false;
  });
  kept[max - 1] = last + '…';
  return kept;
}

/* ---------------- 資料（ページ）にする ---------------- */

/**
 * 会議サマリをページの並びにする（シートに書く前の形。テストと見た目の確認にも使う）。
 * 全体サマリー → 今回の論点 → 前回からの進捗・変化 → 停滞・失注案件 → 要対応（アラート）。
 * 詳しいページは、入りきらなければ続きのページに分け、中身が無ければ作らない（今回の論点に「なし」と出る）。
 * 戻り値 [{widths, values, styles, heights, boxes, merges}]（ページ番号は下の行の右）
 */
function mtgPages_(summary) {
  const changes = summary.changes.length ? mtgChangePages_(summary) : [];
  const stale = summary.stale.length || summary.closed.length ? mtgStalePages_(summary) : [];
  const alerts = summary.alerts.length ? mtgAlertPages_(summary) : [];
  const refs = {
    changes: changes.length ? 3 : 0,
    stale: stale.length ? 3 + changes.length : 0,
    alerts: alerts.length ? 3 + changes.length + stale.length : 0
  };
  const pages = [mtgOverviewPage_(summary), mtgTopicsPage_(summary, refs)].concat(changes, stale, alerts);
  pages.forEach((page, i) => mtgPageFoot_(page, summary, i + 1));
  return pages;
}

/** 新しいページ（左上のタグ・タイトル（日付つき）・リード文まで）。 */
function mtgNewPage_(summary, widths, title, lead) {
  const c = MTG_OPTIONS.colors;
  const page = {widths, values: [], styles: [], heights: [], boxes: [], merges: []};
  const tag = mtgRow_(page, 28);
  let n = 0;
  for (let w = 0; n < widths.length && w < MTG_PAGE.tagWidth; n++) w += widths[n];
  mtgPaint_(page, tag, 1, n, {bg: c.purple});
  mtgPut_(page, tag, 1, MTG_OPTIONS.title + '｜' + summary.name + (summary.final ? '' : '（試作）'),
    {fc: c.white, fw: 'bold', fs: 12, wrap: 'overflow'}, 0);
  mtgRow_(page, 8);
  const head = mtgRow_(page, 50);
  mtgPut_(page, head, 1, title + '（' + mtgDateLabel_(summary.now) + '）', {fc: c.purple, fw: 'bold', fs: 24, wrap: 'overflow'}, 0);
  mtgLines_(lead, MTG_PAGE.width, 12, MTG_PAGE.leadLines).forEach(line => {
    const row = mtgRow_(page, 24);
    mtgPut_(page, row, 1, line, {fc: c.ink, fw: 'bold', fs: 12, wrap: 'overflow'}, 0);
  });
  mtgRow_(page, 12);
  return page;
}

/** 行を足す（白地・本文の文字）。足した行の番号（1 から）。 */
function mtgRow_(page, height) {
  const c = MTG_OPTIONS.colors;
  page.values.push(page.widths.map(() => ''));
  page.styles.push(page.widths.map(() => ({bg: c.white, fc: c.body, fw: 'normal', fs: MTG_PAGE.fontSize, ha: 'left', wrap: 'clip'})));
  page.heights.push(height);
  return page.values.length;
}

/** セルに書く。fit：0 なら切らない、数ならその幅（px）、省けばその列の幅に収める（wrap: 'wrap' のセルは lines 行まで）。 */
function mtgPut_(page, row, col, text, style, fit, lines) {
  const s = Object.assign(page.styles[row - 1][col - 1], style || {});
  const px = fit === 0 ? 0 : (fit || page.widths[col - 1]);
  page.values[row - 1][col - 1] = px ? mtgFit_(text, px, s.fs, s.wrap === 'wrap' ? (lines || 2) : 1) : String(text == null ? '' : text);
}

function mtgPaint_(page, row, col, count, style) {
  for (let i = 0; i < count; i++) Object.assign(page.styles[row - 1][col - 1 + i], style);
}

function mtgPageHeight_(page) {
  return page.heights.reduce((sum, h) => sum + h, 0);
}

/** 結論の帯（ピンクの帯に紫の太字。ページの下。帯の行は1つのセルにまとめて真ん中に書く）。 */
function mtgConclusion_(page, text) {
  const c = MTG_OPTIONS.colors;
  mtgRow_(page, 10);
  const row = mtgRow_(page, 34);
  mtgPaint_(page, row, 1, page.widths.length, {bg: c.highlight});
  page.merges.push({row, col: 1, cols: page.widths.length});
  mtgPut_(page, row, 1, text, {fc: c.purple, fw: 'bold', fs: 14, ha: 'center'}, MTG_PAGE.width);
}

/** ページの下の行：出典と注記（左）・ページ番号（右）。 */
function mtgPageFoot_(page, summary, number) {
  const c = MTG_OPTIONS.colors;
  mtgRow_(page, 6);
  const row = mtgRow_(page, 18);
  mtgPut_(page, row, 1, '出典：' + summary.name + '・' + DIFF_OPTIONS.logSheet + '・' + CRED_OPTIONS.logSheet + '　／　金額の単位：百万円　／　' +
    '停滞：最終更新日時が' + MTG_OPTIONS.staleDays + '日以上前の進行中の案件（「日以上」は変更の記録を始めた日から）' +
    (summary.unmatched ? '　／　案件を特定できない変更 ' + summary.unmatched + ' 件は含めていません' : ''),
  {fc: c.faint, fs: 8, wrap: 'overflow'}, 0);
  mtgPut_(page, row, page.widths.length, String(number), {fc: c.faint, fs: 9, ha: 'right'});
}

/** 結論の帯と下の行のぶんの高さ（表を続きのページに分けるときに残しておく）。 */
const MTG_PAGE_BOTTOM = 10 + 34 + 6 + 18;

/**
 * 表をページに書く（紫の見出し・薄い紫の1列目・薄いグレーのセル・白の区切り線）。入りきらなければ続きのページを作り、見出しの行を繰り返す。
 * section = {widths, title, lead, conclusion, tables: [{caption, columns: [{label, ha, fill, head}], rows: [{cells, lines, focus}], empty}]}
 *   columns の fill：'label'（薄い紫）・'gray'（薄いグレー。既定）、head：'gray' ならグレーの見出し、wrap：折り返す列（文の列。数の列は1行）
 *   rows の cells：[{text, style}]（null は空）、lines：何行まで（1 か 2）、focus：赤枠で囲む行
 * 戻り値はページの並び。
 */
function mtgSectionPages_(summary, section) {
  const c = MTG_OPTIONS.colors;
  const pages = [];
  const limit = MTG_PAGE.maxHeight - MTG_PAGE_BOTTOM;
  let page = null;
  const newPage = cont => {
    if (page) mtgConclusion_(page, section.conclusion);
    page = mtgNewPage_(summary, section.widths, section.title + (cont ? '（続き）' : ''), section.lead);
    pages.push(page);
  };
  newPage(false);
  section.tables.forEach((table, t) => {
    const first = table.rows[0];
    const rowHeight = r => (r && r.lines > 1 ? MTG_PAGE.lineHeight * r.lines + 8 : MTG_PAGE.rowHeight);
    // 2つ目からの表は、見出しと最初の行が入らなければ次のページから
    if (t > 0) {
      if (mtgPageHeight_(page) + 14 + 30 + 28 + rowHeight(first) > limit) newPage(true);
      else mtgRow_(page, 14);
    }
    let start = 0;
    const head = () => {
      if (table.caption) {
        const row = mtgRow_(page, 30);
        mtgPut_(page, row, 1, table.caption, {fc: c.purple, fw: 'bold', fs: 14, wrap: 'overflow'}, 0);
      }
      const row = mtgRow_(page, 28);
      table.columns.forEach((col, i) => {
        mtgPaint_(page, row, i + 1, 1, {bg: col.head === 'gray' ? c.grayHead : c.purple});
        mtgPut_(page, row, i + 1, col.label, {fc: c.white, fw: 'bold', fs: 11, ha: 'center'});
      });
      start = row;
    };
    const close = () => {
      // 白い区切り線（セルどうしの間をあける）。赤枠は区切り線のあとに付ける
      const end = page.values.length;
      page.boxes.unshift({row: start, col: 1, rows: end - start + 1, cols: table.columns.length, all: true, color: c.white, style: 'medium'});
    };
    head();
    // 赤枠：続いている focus の行をまとめて囲む（ページをまたぐときはページごと）
    let focusFrom = 0;
    let focusTo = 0;
    const closeFocus = () => {
      if (!focusFrom) return;
      page.boxes.push({row: focusFrom, col: 1, rows: focusTo - focusFrom + 1, cols: table.columns.length, outer: true, color: c.red, style: 'thick'});
      focusFrom = 0;
      focusTo = 0;
    };
    table.rows.forEach(item => {
      if (mtgPageHeight_(page) + rowHeight(item) > limit) {
        closeFocus();
        close();
        newPage(true);
        head();
      }
      const row = mtgRow_(page, rowHeight(item));
      table.columns.forEach((col, i) => {
        mtgPaint_(page, row, i + 1, 1, {bg: col.fill === 'label' ? c.purpleSoft : col.fill === 'none' ? c.white : c.gray, ha: col.ha || 'left'});
        const cell = item.cells[i];
        if (!cell) return;
        const style = Object.assign({wrap: item.lines > 1 && col.wrap ? 'wrap' : 'clip'}, cell.style || {});
        mtgPut_(page, row, i + 1, cell.text, style, cell.fit, item.lines);
      });
      if (item.focus) {
        if (!focusFrom) focusFrom = row;
        focusTo = row;
      } else {
        closeFocus();
      }
    });
    closeFocus();
    if (!table.rows.length) {
      const row = mtgRow_(page, MTG_PAGE.rowHeight);
      mtgPut_(page, row, 1, table.empty || 'ありません。', {fc: c.faint, wrap: 'overflow'}, 0);
    }
    close();
  });
  mtgConclusion_(page, section.conclusion);
  return pages;
}

/** 全体サマリー（1ページ）：紫の帯に全体の数字、その下に案件ステータス別の表。 */
function mtgOverviewPage_(summary) {
  const c = MTG_OPTIONS.colors;
  const k = summary.kpi;
  const widths = MTG_PAGE.widths.overview;
  const page = mtgNewPage_(summary, widths, '全体サマリー', mtgPeriodText_(summary) + 'の変化をもとに、パイプライン全体と今回の論点をまとめました。');
  const banner = mtgRow_(page, 38);
  mtgPaint_(page, banner, 1, widths.length, {bg: c.purple});
  page.merges.push({row: banner, col: 1, cols: widths.length});
  mtgPut_(page, banner, 1, '進行中 ' + k.active + '件・想定 ' + mtgAmount_(k.amount) + '百万円（期待値調整済 ' + mtgAmount_(k.weighted) + '）　｜　' +
    '前回から動いた案件 ' + k.changed + '件　｜　停滞 ' + summary.stale.length + '件　｜　要対応 ' + summary.alerts.length + '件',
  {fc: c.white, fw: 'bold', fs: 15, ha: 'center'}, MTG_PAGE.width);
  mtgRow_(page, 16);

  const limit = MTG_PAGE.maxHeight - MTG_PAGE_BOTTOM;
  const head = mtgRow_(page, 28);
  ['案件ステータス', '件数', '想定', '期待値', '構成比', '件数の分布'].forEach((label, i) => {
    mtgPaint_(page, head, i + 1, 1, {bg: c.purple});
    mtgPut_(page, head, i + 1, label, {fc: c.white, fw: 'bold', fs: 11, ha: 'center'});
  });
  const pipeline = summary.pipeline;
  const room = Math.floor((limit - mtgPageHeight_(page) - MTG_PAGE.rowHeight * 2) / MTG_PAGE.rowHeight);
  const shown = pipeline.length > room ? pipeline.slice(0, room - 1) : pipeline;
  const maxCount = Math.max.apply(null, pipeline.map(p => p.count).concat([1]));
  const line = (label, count, amount, weighted, style, bar, barColor) => {
    const row = mtgRow_(page, MTG_PAGE.rowHeight);
    mtgPaint_(page, row, 1, 1, {bg: c.purpleSoft});
    mtgPaint_(page, row, 2, widths.length - 1, {bg: c.gray});
    mtgPut_(page, row, 1, label, style);
    mtgPut_(page, row, 2, String(count), Object.assign({}, style, {ha: 'right'}));
    mtgPut_(page, row, 3, mtgAmount_(amount), Object.assign({}, style, {ha: 'right'}));
    mtgPut_(page, row, 4, mtgAmount_(weighted), Object.assign({}, style, {ha: 'right'}));
    mtgPut_(page, row, 5, k.deals ? Math.round(count / k.deals * 100) + '%' : '—', Object.assign({}, style, {ha: 'right'}));
    if (bar) mtgPut_(page, row, 6, bar, {fc: barColor, fs: 10});
    return row;
  };
  shown.forEach(p => {
    const style = p.kind === 'won' ? {fc: c.purple, fw: 'bold'} : p.kind === 'lost' ? {fc: c.lost} : {fc: c.ink};
    const tag = p.kind === 'won' && p.status.indexOf('受注') < 0 ? '（受注）' : p.kind === 'lost' && p.status.indexOf('失注') < 0 ? '（終了）' : '';
    line(p.status + tag, p.count, p.amount, p.weighted, style,
      '█'.repeat(Math.max(1, Math.round(p.count / maxCount * MTG_PAGE.barWidth))), p.kind === 'lost' ? c.grayHead : c.purple);
  });
  if (pipeline.length > shown.length) {
    const rest = pipeline.slice(shown.length);
    line('ほか ' + rest.length + ' ステータス', mtgSum_(rest, 'count'), mtgSum_(rest, 'amount'), mtgSum_(rest, 'weighted'), {fc: c.muted});
  }
  if (!pipeline.length) line('案件がありません', 0, 0, 0, {fc: c.faint});
  const total = line('進行中の合計', k.active, k.amount, k.weighted, {fc: c.ink, fw: 'bold'});
  page.boxes.push({row: total, col: 1, rows: 1, cols: widths.length, top: true, color: c.purple, style: 'medium'});
  page.boxes.unshift({row: head, col: 1, rows: total - head + 1, cols: widths.length, all: true, color: c.white, style: 'medium'});

  mtgConclusion_(page, '前回から ' + k.changed + '件が動き、受注 ' + mtgWon_(summary).length + '件・失注など ' + summary.closed.length + '件。' +
    '停滞 ' + summary.stale.length + '件と要対応 ' + summary.alerts.length + '件を、次のページの論点で確認したい');
  return page;
}

/** 今回の論点（1ページ）：論点 ▶ 状況 ▶ 確認したいこと。refs は詳しいページの番号（無ければ 0）。 */
function mtgTopicsPage_(summary, refs) {
  const c = MTG_OPTIONS.colors;
  const o = MTG_OPTIONS;
  const k = summary.kpi;
  const widths = MTG_PAGE.widths.topics;
  const staleAmount = mtgSum_(summary.stale, 'amount');
  const page = mtgNewPage_(summary, widths, '今回の論点',
    '前回から ' + k.changed + '件が動き、受注 ' + mtgWon_(summary).length + '件・失注など ' + summary.closed.length + '件。' +
    o.staleDays + '日以上更新のない進行中の案件は ' + summary.stale.length + '件（想定 ' + mtgAmount_(staleAmount) + '百万円）、要対応は ' +
    summary.alerts.length + '件です。');
  const items = (list, text) => list.slice(0, o.topicItems).map(item => '・' + text(item));
  const more = (list, ref) => (list.length > o.topicItems ? ['ほか ' + (list.length - o.topicItems) + '件' + (ref ? '（' + ref + 'ページ）' : '')] : []);
  const won = mtgWon_(summary);
  const moved = summary.changes.filter(ch => ch.statusChanged && ch.toKind === 'active');
  const focusStale = summary.stale.slice(0, o.staleFocus);
  const counts = ['予定超過', '受注月超過', '未入力'].map(t => [t, summary.alerts.filter(a => a.type === t).length]).filter(([, n]) => n);
  const topics = [
    {key: 'won', count: won.length, lines: won.length
      ? ['想定 ' + mtgAmount_(mtgSum_(won, 'amount')) + '百万円'].concat(
        items(won, ch => mtgDealLabel_(ch) + '（' + mtgAmount_(ch.amount) + '）'), more(won, refs.changes))
      : []},
    {key: 'lost', count: summary.closed.length, lines: items(summary.closed,
      d => mtgDealLabel_(d) + '（' + (d.removed ? '行を削除' : (d.from || '未設定') + ' ▶ ' + d.to) + '）').concat(more(summary.closed, refs.stale))},
    {key: 'moved', count: moved.length, lines: items(moved,
      ch => mtgDealLabel_(ch) + '：' + (ch.from || '未設定') + ' ▶ ' + (ch.to || '未設定')).concat(more(moved, refs.changes))},
    {key: 'stale', count: summary.stale.length, lines: summary.stale.length
      ? ['想定 ' + mtgAmount_(staleAmount) + '百万円（想定の大きい順）'].concat(
        items(focusStale, d => mtgDealLabel_(d) + '（' + mtgAmount_(d.amount) + '・' + d.days + (d.noRecord ? '日以上' : '日') + '）'),
        summary.stale.length > focusStale.length ? ['ほか ' + (summary.stale.length - focusStale.length) + '件' + (refs.stale ? '（' + refs.stale + 'ページ）' : '')] : [])
      : []},
    {key: 'alerts', count: summary.alerts.length, lines: summary.alerts.length
      ? [counts.map(([t, n]) => t + ' ' + n + '件').join('・')].concat(
        items(summary.alerts, a => mtgDealLabel_(a) + '：' + a.type), more(summary.alerts, refs.alerts))
      : []}
  ];
  const rows = topics.map(t => {
    // 折り返すセルなので、1行ずつ本物の折り返しより早めに切る（1行が2行に折り返されて、下の行が隠れないように）
    const lines = t.count ? t.lines.map(text => mtgLines_(text, widths[3], MTG_PAGE.fontSize, 1)[0]) : ['なし'];
    const ask = t.count ? mtgLines_(o.topics[t.key].ask, widths[5], 12, 3) : ['—'];
    const height = Math.max(lines.length, ask.length, 2) * MTG_PAGE.lineHeight + 14;
    return {topic: o.topics[t.key], count: t.count, lines, ask, height};
  });
  const limit = MTG_PAGE.maxHeight - MTG_PAGE_BOTTOM;
  const spare = limit - mtgPageHeight_(page) - 28 - rows.reduce((sum, r) => sum + r.height, 0);
  if (spare < 0) {
    // 入りきらないときは、論点ごとの案件の行を減らす（最初の行と「ほか」は残す）
    rows.forEach(r => { if (r.lines.length > 3) r.lines = r.lines.slice(0, 2).concat(r.lines.slice(-1)); });
    rows.forEach(r => { r.height = Math.max(r.lines.length, r.ask.length, 2) * MTG_PAGE.lineHeight + 14; });
  }
  const head = mtgRow_(page, 28);
  [['論点', c.purple], ['', c.white], ['件数', c.grayHead], ['状況', c.grayHead], ['', c.white], ['確認したいこと', c.purple]].forEach(([label, bg], i) => {
    mtgPaint_(page, head, i + 1, 1, {bg});
    mtgPut_(page, head, i + 1, label, {fc: c.white, fw: 'bold', fs: 12, ha: 'center'});
  });
  rows.forEach(r => {
    const row = mtgRow_(page, r.height);
    mtgPaint_(page, row, 1, 1, {bg: c.purple});
    mtgPut_(page, row, 1, r.topic.title, {fc: c.white, fw: 'bold', fs: 13, ha: 'center', wrap: 'wrap'}, 0);
    [2, 5].forEach(col => mtgPut_(page, row, col, '▶', {fc: c.arrow, fs: 18, ha: 'center'}, 0));
    mtgPaint_(page, row, 3, 2, {bg: c.gray});
    mtgPut_(page, row, 3, r.count + '件', {fc: r.count ? c.ink : c.faint, fw: 'bold', fs: 16, ha: 'center'});
    mtgPut_(page, row, 4, r.lines.join('\n'), {fc: r.count ? c.ink : c.faint, wrap: 'wrap'}, 0);
    mtgPaint_(page, row, 6, 1, {bg: c.purpleSoft});
    mtgPut_(page, row, 6, r.ask.join('\n'), {fc: r.count ? c.ink : c.faint, fw: r.count ? 'bold' : 'normal', fs: 12, wrap: 'wrap'}, 0);
    if (r.topic.focus && r.count) page.boxes.push({row, col: 1, rows: 1, cols: widths.length, outer: true, color: c.red, style: 'thick'});
  });
  page.boxes.unshift({row: head, col: 1, rows: page.values.length - head + 1, cols: widths.length, all: true, color: c.white, style: 'medium'});
  mtgConclusion_(page, mtgFocusText_(summary));
  return page;
}

/** 前回からの進捗・変化（詳しいページ）。ステータスが動いた案件を赤枠で囲む。 */
function mtgChangePages_(summary) {
  const c = MTG_OPTIONS.colors;
  const k = summary.kpi;
  const won = mtgWon_(summary).length;
  const lostCount = summary.changes.filter(ch => ch.statusChanged && ch.toKind === 'lost').length;
  return mtgSectionPages_(summary, {
    widths: MTG_PAGE.widths.changes,
    title: '前回からの進捗・変化',
    lead: mtgPeriodText_(summary) + 'に値が変わった案件は ' + k.changed + '件。ステータスが動いたのは ' + k.statusChanged + '件' +
      (k.statusChanged ? '（受注 ' + won + '・失注など ' + lostCount + '）' : '') + 'です。',
    conclusion: k.statusChanged
      ? 'ステータスが動いた ' + k.statusChanged + '件（赤枠）を中心に、次のアクションと受注月の見通しを確認したい'
      : '活動の更新内容を確認し、次のアクションを決めたい',
    tables: [{
      columns: [{label: '得意先／案件名', fill: 'label', wrap: true}, {label: '想定', ha: 'right'}, {label: '期待値', ha: 'right'},
        {label: 'ステータスの動き', wrap: true}, {label: '更新した項目', wrap: true}, {label: '責任者', ha: 'center'}, {label: '活動状況（今）', wrap: true}],
      rows: summary.changes.map(ch => ({
        lines: 2,
        focus: ch.statusChanged,
        cells: [
          {text: mtgDealLabel_(ch), style: {fc: c.ink, fw: 'bold'}},
          {text: mtgAmount_(ch.amount)},
          {text: mtgAmount_(ch.weighted)},
          {text: ch.statusChanged ? (ch.from || '未設定') + ' ▶ ' + (ch.to || '未設定') : (ch.isNew ? '新規・' : '') + (ch.status || '—'),
            style: !ch.statusChanged ? {fc: c.muted}
              : ch.toKind === 'won' ? {fc: c.purple, fw: 'bold'} : ch.toKind === 'lost' ? {fc: c.red, fw: 'bold'} : {fc: c.ink, fw: 'bold'}},
          {text: ch.columns.map(col => col.replace(/\s+/g, '')).join('・') || '—', style: {fc: c.muted}},
          {text: ch.owner || '—'},
          {text: ch.activity || '—'}
        ]
      }))
    }]
  });
}

/** 停滞・失注案件（詳しいページ）。停滞している案件のうち想定の大きい MTG_OPTIONS.staleFocus 件を赤枠で囲む。 */
function mtgStalePages_(summary) {
  const c = MTG_OPTIONS.colors;
  const o = MTG_OPTIONS;
  const columns = labels => labels.map((label, i) => ({label, fill: i === 0 ? 'label' : 'gray', wrap: [0, 1, 7, 8].indexOf(i) >= 0,
    ha: [2, 3, 4].indexOf(i) >= 0 ? 'right' : [5, 6].indexOf(i) >= 0 ? 'center' : 'left'}));
  return mtgSectionPages_(summary, {
    widths: MTG_PAGE.widths.stale,
    title: '停滞・失注案件',
    lead: o.staleDays + '日以上更新のない進行中の案件は ' + summary.stale.length + '件（想定 ' + mtgAmount_(mtgSum_(summary.stale, 'amount')) +
      '百万円）。前回から失注・クローズ・削除になった案件は ' + summary.closed.length + '件です。',
    conclusion: mtgFocusText_(summary),
    tables: [{
      caption: summary.closed.length ? '停滞している案件（想定の大きい順）' : '',
      columns: columns(['得意先／案件名', 'ステータス', '停滞', '想定', '期待値', '受注月', '責任者', '活動における課題', '活動状況']),
      empty: '停滞している案件はありません。',
      rows: summary.stale.map((d, i) => ({
        lines: 2,
        focus: i < o.staleFocus,
        cells: [
          {text: mtgDealLabel_(d), style: {fc: c.ink, fw: 'bold'}},
          {text: d.status || '未設定'},
          {text: d.days + (d.noRecord ? '日以上' : '日'), style: d.days >= o.staleAlertDays ? {fc: c.red, fw: 'bold'} : {}},
          {text: mtgAmount_(d.amount)},
          {text: mtgAmount_(d.weighted)},
          {text: d.orderMonthText || '—'},
          {text: d.owner || '—'},
          {text: d.issue || '—'},
          {text: d.activity || '—', style: {fc: c.muted}}
        ]
      }))
    }].concat(summary.closed.length ? [{
      caption: '前回から失注・クローズ・削除になった案件',
      columns: columns(['得意先／案件名', '動き', '日付', '想定', '期待値', '受注月', '責任者', '活動における課題', '活動状況']),
      rows: summary.closed.map(d => ({
        lines: 2,
        cells: [
          {text: mtgDealLabel_(d), style: {fc: c.ink, fw: 'bold'}},
          {text: d.removed ? '行を削除' : (d.from || '未設定') + ' ▶ ' + d.to, style: {fc: c.red, fw: 'bold'}},
          {text: Utilities.formatDate(new Date(d.at), Session.getScriptTimeZone(), 'MM/dd')},
          {text: mtgAmount_(d.amount)},
          {text: mtgAmount_(d.weighted)},
          {text: d.orderMonthText || '—'},
          {text: d.owner || '—'},
          {text: d.issue || '—'},
          {text: d.activity || '—', style: {fc: c.muted}}
        ]
      }))
    }] : [])
  });
}

/** 要対応（アラート。詳しいページ）。 */
function mtgAlertPages_(summary) {
  const c = MTG_OPTIONS.colors;
  const counts = ['予定超過', '受注月超過', '未入力'].map(t => [t, summary.alerts.filter(a => a.type === t).length]);
  return mtgSectionPages_(summary, {
    widths: MTG_PAGE.widths.alerts,
    title: '要対応（アラート）',
    lead: '予定日を過ぎても実施日が空 ' + counts[0][1] + '件・受注月を過ぎても進行中 ' + counts[1][1] + '件・必須項目（' +
      MTG_OPTIONS.required.map(label => label.replace(/\s+/g, '')).join('・') + '）が空 ' + counts[2][1] + '件です。',
    conclusion: '期限を過ぎた予定の実施・受注月の見直し・未入力の項目の入力を、各責任者に依頼したい',
    tables: [{
      columns: [{label: '得意先／案件名', fill: 'label', wrap: true}, {label: '種類', ha: 'center'}, {label: '内容', wrap: true}, {label: '想定', ha: 'right'},
        {label: '責任者・担当', ha: 'center'}],
      rows: summary.alerts.map(a => ({
        lines: 1,
        cells: [
          {text: mtgDealLabel_(a), style: {fc: c.ink, fw: 'bold'}},
          {text: a.type, style: {fc: c.red, fw: 'bold'}},
          {text: a.text},
          {text: mtgAmount_(a.amount)},
          {text: a.owner || '—'}
        ]
      }))
    }]
  });
}

/** 前回から受注になった案件（ステータスの変化のうち受注へ）。 */
function mtgWon_(summary) {
  return summary.changes.filter(ch => ch.statusChanged && ch.toKind === 'won');
}

/** 今回いちばん確かめたいことの結論の文（停滞の大型案件 → 要対応 → ステータスの動き の順に、あるものを出す）。 */
function mtgFocusText_(summary) {
  const o = MTG_OPTIONS;
  const focus = summary.stale.slice(0, o.staleFocus);
  if (focus.length) {
    return '特に想定の大きい停滞案件 ' + focus.length + '件（想定 ' + mtgAmount_(mtgSum_(focus, 'amount')) + '百万円）は、' + o.topics.stale.ask + 'を確認したい';
  }
  if (summary.alerts.length) return '要対応 ' + summary.alerts.length + '件の解消を、各責任者に依頼したい';
  if (summary.kpi.statusChanged) return 'ステータスが動いた ' + summary.kpi.statusChanged + '件の次のアクションを確認したい';
  return '前回からの大きな動きはありません。各案件の次のアクションを確認したい';
}

/** 「前回（2026/09/30 18:04）以降」「初回のため直近14日（09/30〜）」 */
function mtgPeriodText_(summary) {
  const tz = Session.getScriptTimeZone();
  return summary.first
    ? '初回のため直近' + MTG_OPTIONS.firstDays + '日（' + Utilities.formatDate(summary.since, tz, 'MM/dd') + '〜）'
    : '前回（' + Utilities.formatDate(summary.since, tz, 'yyyy/MM/dd HH:mm') + '）以降';
}

/** タイトルに付ける日付（「10/14」。先頭の 0 は付けない） */
function mtgDateLabel_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd').replace(/^0/, '').replace(/\/0/, '/');
}

/* ---------------- 書き出す ---------------- */

/** 1ページをシートに書く（列の幅・行の高さ・値・書式・枠線・セルのまとめ。書式はページごとにまとめて1回ずつ）。 */
function mtgRender_(sheet, page) {
  const rows = page.values.length;
  const cols = page.widths.length;
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxRows() > rows) sheet.deleteRows(rows + 1, sheet.getMaxRows() - rows);
  if (sheet.getMaxColumns() > cols) sheet.deleteColumns(cols + 1, sheet.getMaxColumns() - cols);
  page.widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  // 同じ高さの行はまとめて付ける
  for (let start = 0; start < rows;) {
    let end = start;
    while (end + 1 < rows && page.heights[end + 1] === page.heights[start]) end++;
    sheet.setRowHeights(start + 1, end - start + 1, page.heights[start]);
    start = end + 1;
  }
  const range = sheet.getRange(1, 1, rows, cols);
  const map = fn => page.styles.map(line => line.map(fn));
  // 文字は先頭に ' を付けて書く（「10/06」などが日付・数にならないように）
  range.setValues(page.values.map(line => line.map(v => credText_(v))));
  range.setBackgrounds(map(s => s.bg));
  range.setFontColors(map(s => s.fc));
  range.setFontWeights(map(s => s.fw));
  range.setFontSizes(map(s => s.fs));
  range.setHorizontalAlignments(map(s => s.ha));
  range.setVerticalAlignment('middle');
  const W = SpreadsheetApp.WrapStrategy;
  range.setWrapStrategies(map(s => (s.wrap === 'overflow' ? W.OVERFLOW : s.wrap === 'wrap' ? W.WRAP : W.CLIP)));
  page.merges.forEach(m => sheet.getRange(m.row, m.col, 1, m.cols).merge());
  const B = SpreadsheetApp.BorderStyle;
  page.boxes.forEach(b => {
    const style = b.style === 'thick' ? B.SOLID_THICK : b.style === 'medium' ? B.SOLID_MEDIUM : B.SOLID;
    const edge = b.all || b.outer;
    sheet.getRange(b.row, b.col, b.rows, b.cols).setBorder(edge || b.top || null, edge || b.left || null, edge || b.bottom || null,
      edge || b.right || null, b.all || null, b.all || null, b.color, style);
  });
}

/**
 * 会議サマリを PDF にしてフォルダに保存する。シートごとに一時的なスプレッドシートを作り、1ページを1枚のシートに書いて、
 * すべてのシートを A4 横・幅に合わせて書き出す。一時的なスプレッドシートは、うまくいかなかったときもゴミ箱に移す。
 * 戻り値は保存したファイル（summaries と同じ順）。
 */
function mtgExportPdfs_(summaries, folder, now, final) {
  return summaries.map(summary => {
    const temp = SpreadsheetApp.create(MTG_OPTIONS.tempName);
    try {
      const blank = temp.getSheets()[0];
      mtgPages_(summary).forEach((page, i) => mtgRender_(temp.insertSheet('P' + (i + 1)), page));
      temp.deleteSheet(blank);
      SpreadsheetApp.flush();
      const blob = mtgFetchPdf_(temp.getId());
      const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return folder.createFile(blob.setName(MTG_OPTIONS.fileName + '_' + summary.name + '_' + date + (final ? '' : '_試作') + '.pdf'));
    } finally {
      try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch (e) { console.warn(e); }
    }
  });
}

/** スプレッドシートのすべてのシートを PDF で書き出す（A4 横・幅に合わせる・枠線なし）。混んでいるときは1回だけやり直す。 */
function mtgFetchPdf_(spreadsheetId) {
  const params = {
    format: 'pdf', size: 'A4', portrait: 'false', scale: '2', fitw: 'true',   // 2 = 幅に合わせる（どのページも同じ幅なので、文字の大きさがそろう）
    top_margin: '0.3', bottom_margin: '0.3', left_margin: '0.3', right_margin: '0.3',
    gridlines: 'false', printnotes: 'false', printtitle: 'false', sheetnames: 'false', pagenum: 'UNDEFINED',
    horizontal_alignment: 'CENTER', vertical_alignment: 'TOP', fzr: 'false', attachment: 'false'
  };
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
    Object.keys(params).map(key => key + '=' + encodeURIComponent(params[key])).join('&');
  const options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true};
  let response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 429 || response.getResponseCode() >= 500) {
    Utilities.sleep(5000);
    response = UrlFetchApp.fetch(url, options);
  }
  if (response.getResponseCode() !== 200) {
    throw new Error('PDF に書き出せませんでした（' + response.getResponseCode() + '）。少し待ってからもう一度お試しください。');
  }
  return response.getBlob();
}

/** 保存先のフォルダ（MTG_OPTIONS.folderId。空なら、このスプレッドシートと同じフォルダの中に folderName を作り、次からも使う）。 */
function mtgFolder_(ss) {
  if (MTG_OPTIONS.folderId) return DriveApp.getFolderById(MTG_OPTIONS.folderId);
  const props = PropertiesService.getDocumentProperties();
  const saved = props.getProperty(MTG_OPTIONS.folderKey);
  if (saved) {
    try {
      const folder = DriveApp.getFolderById(saved);
      if (!folder.isTrashed()) return folder;
    } catch (_) {}
  }
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder(MTG_OPTIONS.folderName);
  props.setProperty(MTG_OPTIONS.folderKey, folder.getId());
  return folder;
}
