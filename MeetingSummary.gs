/**
 * 会議サマリ（追跡シートごとのワンペーパー PDF）
 * 差分追跡スクリプト（test.gs）・CredentialHistory.gs・HistorySidebar.gs と同じプロジェクトに置くファイル。
 *
 * 目的
 * - 会議の参加者が全リストをスクロールしなくても、数分で要点（全体感・前回からの動き・停滞と失注・不備）をつかめるようにする。
 *   追跡シート（スーパー・GMS など。DIFF_RULES のシート）ごとに1枚の PDF を作り、会議の2〜3日前に配る（事前に読む運用）。
 * - このサマリは「動いた案件・止まった案件・不備のある案件」のふるい分けに絞る。個別の案件を深く見るのは、
 *   オポチュニティキャンバスなど別の画面で行う。
 *
 * 仕組み
 * - 管理者がエディタで createMeetingSummaries() を実行し、作るシートを選ぶ（試すときは previewMeetingSummaries()）。
 *   シートごとに PDF をドライブのフォルダに保存する（MTG_OPTIONS.folderId。空なら、このスプレッドシートと同じフォルダの中に
 *   MTG_OPTIONS.folderName のフォルダを作る）。配る（共有する）のは手動。
 * - 「前回から」は、そのシートで前に createMeetingSummaries() で作った日時から（初回は MTG_OPTIONS.firstDays 日前から）。
 *   previewMeetingSummaries() は前回の日時を進めないので、何度でも試せる（ファイル名に「試作」が付く）。
 * - 載せるもの（スプレッドシートは読むだけで、書き込まない）
 *   ① トータルサマリー：案件ステータス別の件数・想定売上・期待値調整済（進行中・受注・失注に分ける）
 *   ② 前回からの進捗・変化：変更履歴_差分 から、前回以降に値が変わった案件（ステータスの動き・更新した項目・今の活動状況）
 *   ③ 停滞・失注：最終更新日時が MTG_OPTIONS.staleDays 日以上前の進行中の案件と、前回以降に失注・クローズになった案件・削除された行
 *      （最終更新日時が空の行は、変更の記録を始めた日（変更履歴_差分 の最初の行）から動いていないとみなし、「N日以上」と出す）
 *   ④ アラート：クレデンシャル・オファリングの予定日を過ぎても実施日が空・受注月を過ぎても進行中・必須項目が空
 * - 変更の記録は、変更時点スナップショット（その時の行の内容）で得意先＋案件名に結び付ける（行の追加・削除で行番号がずれても合う）。
 *   どちらの記録も追記なので、下（新しいほう）から MTG_OPTIONS.logChunkRows 行ずつ、前回より前の記録に届くまでだけ読む。
 * - PDF は、一時的なスプレッドシートに1枚の表として書き、「ページに合わせる」で A4 横の1ページに書き出す
 *   （書き出したら、一時的なスプレッドシートはゴミ箱に移す）。書き出しに UrlFetchApp を使う（外部サービスへの接続の権限）。
 */

const MTG_OPTIONS = {
  title: 'RXビジネスMTG サマリ',      // PDF の見出し（「見出し｜シート名」と出す）
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
  // 色（docs/DESIGN.md の「画面のデザイン」と同じ）
  colors: {
    ink: '#1b1b1c', body: '#3c4642', muted: '#5a6660', faint: '#8a968a',
    line: '#EFF2EF', border: '#DDE2DD', soft: '#F4F6F4', panel: '#FAFBFA',
    teal: '#36A191', tealText: '#297A6E', tealSoft: '#E9F4F1', cred: '#1F6E5E',
    danger: '#B42318', dangerSoft: '#FCEDEC', white: '#ffffff'
  }
};

/*
 * 1枚の表の形（行・列は 1 から）。左右2段：左（1〜6列）に ①③、右（8〜13列）に ②④、7列は間のすき間。
 * 行の高さを決めておき、どの段も同じ行数に収める（多いときは最後の行を「ほか N件」にする）。
 */
const MTG_GRID = {
  widths: [150, 78, 62, 58, 66, 150, 14, 150, 150, 100, 52, 66, 170],
  left: 1,
  right: 8,
  rows: {title: 1, sub: 2, kpiLabel: 4, kpiValue: 5, sec1: 7, head1: 8, items1: 9, sec2: 23, head2: 24, items2: 25, foot: 39},
  slots: 13,          // 段ごとの中身の行数
  height: 39,
  heights: {1: 34, 2: 20, 3: 6, 4: 18, 5: 32, 6: 10, 7: 24, 8: 18, 22: 10, 23: 24, 24: 18, 38: 6, 39: 18},
  rowHeight: 22,      // 中身の行の高さ
  fontSize: 9,
  barWidth: 14        // ①の棒の最大の長さ（文字数）
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
      owner: deal ? deal.owner : '',
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
    .map(c => ({customer: c.customer, project: c.project, from: c.from, to: c.to, at: c.at, amount: c.amount, owner: c.owner, removed: false}))
    .concat(removed.map(r => ({customer: r.customer, project: r.project, from: r.status, to: '', at: r.at,
      amount: r.amount, owner: r.owner, removed: true})))
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
 * アラート [{type, customer, project, text, amount}]：予定日超過（クレデンシャル・オファリング）→ 受注月超過 → 必須項目が空。
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
      out.push({type: '予定超過', customer, project: '',
        text: (c.values.kind || '予定') + '「' + c.values.name + '」 予定 ' + plan.slice(5) + '（' + days + '日超過）' +
          (c.values.owner ? ' 担当 ' + c.values.owner : ''),
        amount: null, days});
    });
  });
  const month = Utilities.formatDate(context.now, tz, 'yyyy-MM').split('-').map(Number);
  const current = month[0] * 12 + month[1];
  const order = [];
  active.forEach(d => {
    if (d.orderMonth && d.orderMonth.y * 12 + d.orderMonth.m < current) {
      order.push({type: '受注月超過', customer: d.customer, project: d.project, amount: d.amount,
        text: '受注月 ' + d.orderMonth.y + '/' + d.orderMonth.m + ' を過ぎても「' + (d.status || '未設定') + '」'});
    }
  });
  const missing = active.filter(d => d.missing.length).map(d => ({type: '未入力', customer: d.customer, project: d.project, amount: d.amount,
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
 * px：列の幅、pt：文字の大きさ。
 */
function mtgFit_(text, px, pt) {
  const value = String(text == null ? '' : text).replace(/\s+/g, ' ').trim();
  const limit = (px - 8) / ((pt || MTG_GRID.fontSize) * 96 / 72);
  let used = 0;
  const chars = Array.from(value);
  for (let i = 0; i < chars.length; i++) {
    used += /[\u0000-ÿ｡-ﾟ]/.test(chars[i]) ? 0.55 : 1;
    if (used > limit) return chars.slice(0, Math.max(i - 1, 0)).join('') + '…';
  }
  return value;
}

/* ---------------- 1枚の表にする ---------------- */

/**
 * 会議サマリを MTG_GRID の形の表にする（シートに書く前の形。テストと見た目の確認にも使う）。
 * {values, styles: [行][列] = {bg, fc, fw, fs, ha, wrap}, boxes: [{row, col, rows, cols, ...枠線}]}
 */
function mtgLayout_(summary) {
  const c = MTG_OPTIONS.colors;
  const g = MTG_GRID;
  const rows = g.height;
  const cols = g.widths.length;
  const values = Array.from({length: rows}, () => new Array(cols).fill(''));
  const styles = Array.from({length: rows}, () => Array.from({length: cols}, () =>
    ({bg: c.white, fc: c.body, fw: 'normal', fs: g.fontSize, ha: 'left', wrap: 'clip'})));
  const boxes = [];
  const width = col => g.widths[col - 1];
  const span = (col, n) => g.widths.slice(col - 1, col - 1 + n).reduce((a, b) => a + b, 0);
  const put = (row, col, text, style, fitPx) => {
    const s = Object.assign(styles[row - 1][col - 1], style || {});
    values[row - 1][col - 1] = fitPx === 0 ? String(text) : mtgFit_(text, fitPx || width(col), s.fs);
  };
  const paint = (row, col, n, style) => {
    for (let i = 0; i < n; i++) Object.assign(styles[row - 1][col - 1 + i], style);
  };
  const tz = Session.getScriptTimeZone();
  const at = (date, format) => Utilities.formatDate(date, tz, format || 'yyyy/MM/dd HH:mm');
  const L = g.left;
  const R = g.right;
  const r = g.rows;

  // 見出し
  put(r.title, L, MTG_OPTIONS.title + '｜' + summary.name, {fc: c.ink, fw: 'bold', fs: 16, wrap: 'overflow'}, 0);
  put(r.sub, L, '作成 ' + at(summary.now) + '　・　' + (summary.first
    ? '初回のため、直近' + MTG_OPTIONS.firstDays + '日（' + at(summary.since, 'yyyy/MM/dd') + '〜）の変化'
    : '前回（' + at(summary.since) + '）以降の変化') + (summary.final ? '' : '　・　試作'), {fc: c.muted, fs: 9, wrap: 'overflow'}, 0);
  put(r.sub, R + 5, '金額の単位：百万円', {fc: c.faint, fs: 8, ha: 'right'});

  // 数字のタイル（左：全体、右：前回から）
  const k = summary.kpi;
  const tiles = [
    {col: L, n: 1, label: '進行中の案件', value: k.active + ' 件'},
    {col: L + 1, n: 3, label: '想定売上（進行中）', value: mtgAmount_(k.amount)},
    {col: L + 4, n: 2, label: '期待値調整済（進行中）', value: mtgAmount_(k.weighted)},
    {col: R, n: 1, label: '前回から動いた案件', value: k.changed + ' 件'},
    {col: R + 1, n: 2, label: '停滞案件（' + MTG_OPTIONS.staleDays + '日以上）', value: summary.stale.length + ' 件'},
    {col: R + 3, n: 3, label: 'アラート', value: summary.alerts.length + ' 件', alert: summary.alerts.length > 0}
  ];
  tiles.forEach(t => {
    const bg = t.alert ? c.dangerSoft : c.soft;
    paint(r.kpiLabel, t.col, t.n, {bg});
    paint(r.kpiValue, t.col, t.n, {bg});
    put(r.kpiLabel, t.col, t.label, {fc: t.alert ? c.danger : c.muted, fs: 8, wrap: 'overflow'}, 0);
    put(r.kpiValue, t.col, t.value, {fc: t.alert ? c.danger : c.ink, fw: 'bold', fs: 18, wrap: 'overflow'}, 0);
    boxes.push({row: r.kpiLabel, col: t.col, rows: 2, cols: t.n, left: true, right: true, color: c.white, thick: true});
  });

  // 段の見出し（左右）
  const section = (row, col, title, count, tone) => {
    const bg = tone === 'strong' ? c.cred : tone === 'alert' ? c.dangerSoft : c.tealSoft;
    const fc = tone === 'strong' ? c.white : tone === 'alert' ? c.danger : c.cred;
    paint(row, col, 6, {bg});
    put(row, col, title, {fc, fw: 'bold', fs: 11, wrap: 'overflow'}, 0);
    // 件数は右寄せ（長いときは左の空いたセルにはみ出す）
    if (count) put(row, col + 5, count, {fc, fw: 'bold', fs: 9, ha: 'right', wrap: 'overflow'}, 0);
  };
  const header = (row, col, labels) => {
    labels.forEach((label, i) => {
      if (label === null) return;
      put(row, col + i, label, {fc: c.faint, fs: 8, ha: i > 0 && /件数|想定|期待値|停滞/.test(label) ? 'right' : 'left'});
    });
    boxes.push({row, col, rows: 1, cols: 6, bottom: true, color: c.border});
  };
  const zebra = (first, col, count) => {
    for (let i = 0; i < count; i++) if (i % 2 === 1) paint(first + i, col, 6, {bg: c.panel});
  };
  const more = (row, col, rest, text) => put(row, col, 'ほか ' + rest + ' 件（' + text + '）', {fc: c.muted, wrap: 'overflow'}, 0);
  const none = (row, col, text) => put(row, col, text, {fc: c.faint, wrap: 'overflow'}, 0);

  // ① トータルサマリー（ステータスごと）
  section(r.sec1, L, '① トータルサマリー', '全 ' + k.deals + ' 件（受注 ' + k.won + '・失注 ' + k.lost + '）');
  header(r.head1, L, ['案件ステータス', '件数', '想定', '期待値', '構成', null]);
  const pipeline = summary.pipeline;
  zebra(r.items1, L, Math.min(pipeline.length, g.slots));
  const shownPipeline = pipeline.length > g.slots ? pipeline.slice(0, g.slots - 1) : pipeline;
  const maxCount = Math.max.apply(null, pipeline.map(p => p.count).concat([1]));
  shownPipeline.forEach((p, i) => {
    const row = r.items1 + i;
    const tone = p.kind === 'won' ? {fc: c.tealText, fw: 'bold'} : p.kind === 'lost' ? {fc: c.faint} : {fc: c.ink};
    // 受注・失注のステータスは、名前に「受注」「失注」が無ければ（成約・見送り など）種類を添える
    const tag = p.kind === 'won' && p.status.indexOf('受注') < 0 ? '（受注）' : p.kind === 'lost' && p.status.indexOf('失注') < 0 ? '（終了）' : '';
    put(row, L, p.status + tag, tone);
    put(row, L + 1, String(p.count), Object.assign({ha: 'right'}, tone));
    put(row, L + 2, mtgAmount_(p.amount), {ha: 'right'});
    put(row, L + 3, mtgAmount_(p.weighted), {ha: 'right'});
    put(row, L + 4, '█'.repeat(Math.max(1, Math.round(p.count / maxCount * g.barWidth))),
      {fc: p.kind === 'won' ? c.tealText : p.kind === 'lost' ? c.border : c.teal, fs: 7, wrap: 'overflow'}, 0);
  });
  if (pipeline.length > g.slots) more(r.items1 + g.slots - 1, L, pipeline.length - shownPipeline.length, 'ステータス');
  if (!pipeline.length) none(r.items1, L, '案件がありません。');

  // ② 前回からの進捗・変化（いちばん大事なので、見出しを濃い色にする）
  const changes = summary.changes;
  section(r.sec1, R, '② 前回からの進捗・変化', k.changed + ' 件（ステータス変更 ' + k.statusChanged + '）', 'strong');
  header(r.head1, R, ['得意先／案件名', 'ステータス', '更新した項目', '想定', '責任者', '活動状況（今）']);
  zebra(r.items1, R, Math.min(changes.length, g.slots));
  const shownChanges = changes.length > g.slots ? changes.slice(0, g.slots - 1) : changes;
  shownChanges.forEach((ch, i) => {
    const row = r.items1 + i;
    put(row, R, mtgDealLabel_(ch), {fc: c.ink, fw: ch.statusChanged ? 'bold' : 'normal'});
    const status = ch.statusChanged ? (ch.from || '未設定') + ' → ' + (ch.to || '未設定') : (ch.isNew ? '新規・' : '') + (ch.status || '');
    const tone = !ch.statusChanged ? {fc: ch.isNew ? c.tealText : c.faint}
      : ch.toKind === 'won' ? {fc: c.tealText, fw: 'bold'} : ch.toKind === 'lost' ? {fc: c.danger, fw: 'bold'} : {fc: c.ink, fw: 'bold'};
    put(row, R + 1, status, tone);
    put(row, R + 2, ch.columns.map(col => col.replace(/\s+/g, '')).join('・') || '—', {fc: c.muted});
    put(row, R + 3, mtgAmount_(ch.amount), {ha: 'right'});
    put(row, R + 4, ch.owner || '—');
    put(row, R + 5, ch.activity || '—');
  });
  if (changes.length > g.slots) more(r.items1 + g.slots - 1, R, changes.length - shownChanges.length, '変更履歴で確認できます');
  if (!changes.length) none(r.items1, R, '前回以降に動いた案件はありません。');

  // ③ 停滞・失注（停滞を上に、前回から失注・クローズ・削除を下に）
  const stale = summary.stale;
  const closed = summary.closed;
  section(r.sec2, L, '③ 停滞・失注', '停滞 ' + stale.length + '・失注など ' + closed.length);
  header(r.head2, L, ['得意先／案件名', 'ステータス', '停滞', '想定', '責任者', '活動における課題']);
  const closedSlots = Math.min(Math.max(closed.length, 1), 4);
  const staleSlots = g.slots - 1 - closedSlots;
  zebra(r.items2, L, Math.min(stale.length, staleSlots));
  const shownStale = stale.length > staleSlots ? stale.slice(0, staleSlots - 1) : stale;
  shownStale.forEach((d, i) => {
    const row = r.items2 + i;
    put(row, L, mtgDealLabel_(d), {fc: c.ink});
    put(row, L + 1, d.status || '未設定');
    put(row, L + 2, d.days + (d.noRecord ? '日以上' : '日'), {ha: 'right', fc: d.days >= MTG_OPTIONS.staleAlertDays ? c.danger : c.body});
    put(row, L + 3, mtgAmount_(d.amount), {ha: 'right'});
    put(row, L + 4, d.owner || '—');
    put(row, L + 5, d.issue || d.activity || '—', {fc: c.muted});
  });
  if (stale.length > staleSlots) more(r.items2 + staleSlots - 1, L, stale.length - shownStale.length, '想定売上の大きい順に表示');
  if (!stale.length) none(r.items2, L, '停滞している案件はありません。');
  const closedRow = r.items2 + staleSlots;
  put(closedRow, L, '前回から失注・クローズ・削除', {fc: c.danger, fw: 'bold', fs: 8, wrap: 'overflow'}, 0);
  boxes.push({row: closedRow, col: L, rows: 1, cols: 6, top: true, color: c.border});
  const shownClosed = closed.length > closedSlots ? closed.slice(0, closedSlots - 1) : closed;
  shownClosed.forEach((d, i) => {
    const row = closedRow + 1 + i;
    put(row, L, mtgDealLabel_(d), {fc: c.ink});
    put(row, L + 1, d.removed ? '行を削除' : (d.from || '未設定') + '→' + d.to, {fc: c.danger});
    put(row, L + 2, at(new Date(d.at), 'MM/dd'), {ha: 'right', fc: c.muted});
    put(row, L + 3, mtgAmount_(d.amount), {ha: 'right'});
    put(row, L + 4, d.owner || '—');
  });
  if (closed.length > closedSlots) more(closedRow + closedSlots, L, closed.length - shownClosed.length, '変更履歴で確認できます');
  if (!closed.length) none(closedRow + 1, L, 'ありません。');

  // ④ アラート
  const alerts = summary.alerts;
  const counts = ['予定超過', '受注月超過', '未入力'].map(t => [t, alerts.filter(a => a.type === t).length]).filter(([, n]) => n);
  section(r.sec2, R, '④ アラート', counts.map(([t, n]) => t + ' ' + n).join('・'), 'alert');
  header(r.head2, R, ['得意先／案件名', '種類', '内容', null, null, null]);
  zebra(r.items2, R, Math.min(alerts.length, g.slots));
  const shownAlerts = alerts.length > g.slots ? alerts.slice(0, g.slots - 1) : alerts;
  shownAlerts.forEach((a, i) => {
    const row = r.items2 + i;
    put(row, R, mtgDealLabel_(a), {fc: c.ink});
    put(row, R + 1, a.type, {fc: c.danger, fw: 'bold'});
    put(row, R + 2, a.text, {wrap: 'overflow'}, span(R + 2, 4));
  });
  if (alerts.length > g.slots) more(r.items2 + g.slots - 1, R, alerts.length - shownAlerts.length, 'シートで確認してください');
  if (!alerts.length) none(r.items2, R, 'アラートはありません。');

  // 注記
  put(r.foot, L, '停滞：最終更新日時が' + MTG_OPTIONS.staleDays + '日以上前の進行中の案件（「日以上」は変更の記録を始めた日から数えた日数）　／　' +
    '受注・失注：案件ステータスの語で判断　／　出典：' + summary.name + '・' + DIFF_OPTIONS.logSheet + '・' + CRED_OPTIONS.logSheet +
    (summary.unmatched ? '　／　案件を特定できない変更 ' + summary.unmatched + ' 件は含めていません' : ''),
    {fc: c.faint, fs: 8, wrap: 'overflow'}, 0);

  return {values, styles, boxes, rows, cols};
}

/* ---------------- 書き出す ---------------- */

/** 表をシートに書く（列の幅・行の高さ・値・書式。書式は表ごとにまとめて1回ずつ）。 */
function mtgRender_(sheet, layout) {
  const g = MTG_GRID;
  if (sheet.getMaxRows() > layout.rows) sheet.deleteRows(layout.rows + 1, sheet.getMaxRows() - layout.rows);
  if (sheet.getMaxColumns() > layout.cols) sheet.deleteColumns(layout.cols + 1, sheet.getMaxColumns() - layout.cols);
  g.widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeights(1, layout.rows, g.rowHeight);
  Object.keys(g.heights).forEach(row => sheet.setRowHeight(Number(row), g.heights[row]));
  const range = sheet.getRange(1, 1, layout.rows, layout.cols);
  const map = fn => layout.styles.map(line => line.map(fn));
  // 文字は先頭に ' を付けて書く（「10/06」などが日付・数にならないように）
  range.setValues(layout.values.map(line => line.map(v => credText_(v))));
  range.setBackgrounds(map(s => s.bg));
  range.setFontColors(map(s => s.fc));
  range.setFontWeights(map(s => s.fw));
  range.setFontSizes(map(s => s.fs));
  range.setHorizontalAlignments(map(s => s.ha));
  range.setVerticalAlignment('middle');
  const W = SpreadsheetApp.WrapStrategy;
  range.setWrapStrategies(map(s => (s.wrap === 'overflow' ? W.OVERFLOW : W.CLIP)));
  layout.boxes.forEach(b => {
    sheet.getRange(b.row, b.col, b.rows, b.cols).setBorder(b.top || null, b.left || null, b.bottom || null, b.right || null, null, null,
      b.color, b.thick ? SpreadsheetApp.BorderStyle.SOLID_THICK : SpreadsheetApp.BorderStyle.SOLID);
  });
}

/**
 * 会議サマリを PDF にしてフォルダに保存する。一時的なスプレッドシートにシートごとに書き、A4 横・1ページに合わせて書き出す。
 * 一時的なスプレッドシートは、うまくいかなかったときもゴミ箱に移す。戻り値は保存したファイル（summaries と同じ順）。
 */
function mtgExportPdfs_(summaries, folder, now, final) {
  const temp = SpreadsheetApp.create(MTG_OPTIONS.tempName);
  try {
    return summaries.map(summary => {
      const sheet = temp.insertSheet(summary.name);
      const layout = mtgLayout_(summary);
      mtgRender_(sheet, layout);
      SpreadsheetApp.flush();
      const blob = mtgFetchPdf_(temp.getId(), sheet.getSheetId(), layout);
      const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return folder.createFile(blob.setName(MTG_OPTIONS.fileName + '_' + summary.name + '_' + date + (final ? '' : '_試作') + '.pdf'));
    });
  } finally {
    try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch (e) { console.warn(e); }
  }
}

/** スプレッドシートの1枚を PDF で書き出す（A4 横・ページに合わせる・枠線なし）。混んでいるときは1回だけやり直す。 */
function mtgFetchPdf_(spreadsheetId, sheetId, layout) {
  const params = {
    format: 'pdf', size: 'A4', portrait: 'false', scale: '4',   // 4 = ページに合わせる（必ず1ページ）
    top_margin: '0.3', bottom_margin: '0.3', left_margin: '0.3', right_margin: '0.3',
    gridlines: 'false', printnotes: 'false', printtitle: 'false', sheetnames: 'false', pagenum: 'UNDEFINED',
    horizontal_alignment: 'CENTER', vertical_alignment: 'TOP', fzr: 'false', attachment: 'false',
    gid: String(sheetId), r1: '0', c1: '0', r2: String(layout.rows), c2: String(layout.cols)
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
