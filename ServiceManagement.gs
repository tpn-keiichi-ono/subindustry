/**
 * サービス管理（リクエストのとりまとめ）
 * 取り込み元のシート（test.gs の DIFF_RULES で requests: true を付けたシート）の「サービスのリクエスト」から、
 * 検討するリクエストを「サービスリクエスト」シートに登録し、リクエストをもとにサービス案を考える。
 * TaskManagement.gs と同じプロジェクトに置くファイル（見出しの用意・ロックは TaskManagement.gs と共通）。
 *
 * 仕組み
 * - リクエストは、選択パネル（RequestPickerView。メニュー「🟪RXサービスMTG用」→「リクエストを追加」）で選んで登録する。
 *   パネルはサービスリクエストのシートを開いているときだけ開く。
 *   パネルは開いたとき・追加したときに取り込み元の最新を読み、まだ登録していないリクエストだけを出す
 *   （サブインダストリー → 得意先 → 案件名で絞り込み、文字で探せる）。選ぶと、サービスリクエストの最後の行の下に A〜D を書く。
 * - パネルの件数のバッジ（未登録・登録済み）を押すと、一覧をモーダル（RequestListDialog）で確かめられる。
 *   モーダルからも未登録のリクエストを追加でき、登録済みのリクエストの行を削除できる（削除は確かめてから）。
 * - 前の版は、セルの連動プルダウン（候補を非表示の __REQUEST_LISTS の数式で作る）と onEdit で登録していたが、
 *   行の並べ替えでプルダウンがずれる・onEdit が30秒で止まる・差分追跡とロックを取り合う・取り込み元の変更を取りこぼす、
 *   などで安定しなかったのでやめた。setupRequestSheet() が、前の版の A〜D のプルダウンと __REQUEST_LISTS を外す。
 * - 登録したリクエストには ID（SR-0001 の形）を振る。追加したとき・パネルを開いたとき・setupRequestSheet() のときに、
 *   ID の無い行と、ほかの行と同じ ID の行（行のコピーなど）に新しい番号を振る（上の行の ID は残す。削除した番号も使い回さない）。
 * - サービスは最初から決まっているものではなく、リクエストをもとに考える。「サービス案」に入力した名前がサービス案になる
 *   （同じ名前を付けたリクエストが1つのサービス案にまとまる）。判断はプルダウン、ほかの列は自由に入力する
 * - 登録した行は、パネルを開いたとき・setupRequestSheet() のときに取り込み元に合わせる（svcSync_）：
 *   サブインダストリーを同じ案件の値にそろえる。取り込み元に無くなったリクエストにメモは付けない
 *   （行番号は diagnoseRequestSources() で確かめる。前の版で付けたメモは外す）
 * - サービス案ごとの検討の進捗は、タスク管理シート（TaskManagement.gs）でタスクとして管理する
 */

const SVC_OPTIONS = {
  requestSheet: 'サービスリクエスト',
  headerRow: 1,
  // リクエストを読むシートは、test.gs の DIFF_RULES で requests: true を付けたシート（見出しの行も DIFF_RULES の headerRow）。
  // 読む列の見出し（サブインダストリー・得意先・案件名・リクエスト）。1つのセルに1件。どのシートにもちょうど1つずつ必要
  sourceHeaders: ['サブインダストリー', '得意先', '案件名', 'サービスのリクエスト'],
  // 一緒に読む列（リクエストした営業）。見出しが無いシートでは空として扱う。番号は sourceHeaders に続く（4〜7）
  sourceOptionalHeaders: ['アカウント責任者部署', 'アカウント責任者', 'BX部署', 'BX担当'],
  // 登録したリクエストの ID（SR-0001 の形。追加したときに自動で振る。重複させず、削除した番号も使い回さない）
  idPrefix: 'SR-',
  idDigits: 4,
  idCounterKey: 'SVC_LAST_REQUEST_ID',   // これまでに振った最大の番号（ドキュメントのプロパティ）
  // source：取り込み元の列（sourceHeaders・sourceOptionalHeaders を続けた何番目か）。取り込み元から書き、パネルを開いたときに合わせる
  // aliases は前の版の見出し
  requestColumns: [
    // insertBefore：前からあるシートに足すときは、その見出しの列のすぐ左に差し込む
    {key: 'id', label: 'ID', width: 90, insertBefore: 'サブインダストリー'},
    {key: 'subIndustry', label: 'サブインダストリー', width: 160, required: true, source: 0},
    {key: 'customer', label: '得意先', width: 200, required: true, source: 1},
    {key: 'project', label: '案件名', width: 240, required: true, source: 2},
    // リクエストした営業（取り込み元の列と同じ並び）。insertAfter：前からあるシートに足すときは、その見出しの列のすぐ右に差し込む
    {key: 'accountDept', label: 'アカウント責任者部署', width: 160, source: 4, insertAfter: '案件名'},
    {key: 'accountOwner', label: 'アカウント責任者', width: 130, source: 5, insertAfter: 'アカウント責任者部署'},
    {key: 'bxDept', label: 'BX部署', width: 140, source: 6, insertAfter: 'アカウント責任者'},
    {key: 'bxOwner', label: 'BX担当', width: 120, source: 7, insertAfter: 'BX部署'},
    {key: 'request', label: 'リクエスト', width: 360, required: true, source: 3},
    {key: 'decision', label: '判断', width: 130, options: ['未判断', 'サービス化検討', '棄却']},
    // ここから下は人が入力する（insertAfter の見出しが無ければ右端に足す）
    {key: 'feedback', label: 'サービス部門からのフィードバック', width: 300, insertAfter: '判断'},
    {key: 'service', label: 'サービス案', width: 220, aliases: ['サービス']},   // まとめる先のサービス案の名前
    // サービス部門として対応する主管本部と担当者
    {key: 'department', label: '主管本部', width: 140, insertAfter: 'サービス案'},
    {key: 'owner', label: '担当者', width: 120, insertAfter: '主管本部'}
  ],
  // 人が入力する列（どれかが入っていれば「手を付けた行」。判断は別に見る）
  inputKeys: ['feedback', 'service', 'department', 'owner'],
  // 使わなくなった列（setupRequestSheet() で、確認してから削除する。取り込み日は追加日の前の版の見出し）
  removedColumns: ['メモ', '追加日', '取り込み日'],
  menuTitle: '🟪RXサービスMTG用',   // メニューの名前（画面の案内・管理者向けの確認画面のタイトルにもこの名前を出す）
  pickerTemplate: 'RequestPickerView',
  pickerTitle: 'リクエストを追加',
  listTemplate: 'RequestListDialog',   // パネルの件数のバッジから開く一覧のモーダル
  defaultDecision: '未判断',             // まだ判断していないことを表す判断（空と同じに扱う）
  // 前の版で「リクエスト」のセルに付けたメモの先頭（今はメモを付けない。パネルを開いたとき・setupRequestSheet() で外す）
  legacyNotes: ['取り込み元に見つかりません', '取り込み元でリクエストが書き換えられました', '新FMT でリクエストが書き換えられました'],
  // 前の版で付けた保護・候補のシート（setupRequestSheet() で外す）
  legacyProtectDescriptions: ['リクエスト：新FMT から自動で転記する列', 'リクエスト：案件を選ぶと自動で入る列'],
  legacyListSheet: '__REQUEST_LISTS',
  // 前の版の後片付け（A〜D のプルダウン・候補のシート・メモ）が済んだ印（ドキュメントのプロパティ）。
  // 済んでいれば、パネルを開くたび・追加するたびには確かめない（setupRequestSheet() では毎回確かめる）
  legacyDoneKey: 'SVC_LEGACY_CLEANED',
  // タスク管理のサービス案のプルダウンを確かめた記録（ドキュメントのプロパティ）。サービス案の列・行数が変わらなければ、
  // パネルを開いても taskLinkCheckMinutes 分に1回しか確かめない（毎回すべての行を確かめると、パネルの読み込みが遅くなる）
  taskLinkKey: 'SVC_TASK_LINK',
  taskLinkCheckMinutes: 30,
  lockWaitMs: 30000,        // エディタから実行する処理のロックの待ち時間
  pickerLockWaitMs: 20000   // パネルから追加するときのロックの待ち時間
};

/** リクエストを読むシートの名前（test.gs の DIFF_RULES で requests: true を付けたシート。DIFF_RULES の順）。 */
function svcSourceSheets_() {
  return Object.keys(DIFF_RULES).filter(name => DIFF_RULES[name] && DIFF_RULES[name].requests);
}

/** 読むシートの名前を知らせる文に入れる形（「スーパー・GMS」「コンビニ」）。 */
function svcSourceLabel_() {
  return '「' + svcSourceSheets_().join('」「') + '」';
}

/** 取り込み元で読む列の見出し（sourceHeaders に sourceOptionalHeaders を続けたもの。requestColumns の source の番号）。 */
function svcAllSourceHeaders_() {
  return SVC_OPTIONS.sourceHeaders.concat(SVC_OPTIONS.sourceOptionalHeaders || []);
}

/** 取り込み元から書く列の key（source の順：サブインダストリー・得意先・案件名・リクエスト・営業の4列）。 */
function svcSourceKeys_() {
  return SVC_OPTIONS.requestColumns.filter(c => c.source != null).sort((a, b) => a.source - b.source).map(c => c.key);
}

/* ---------------- メニューと選択パネル ---------------- */

/** メニュー「🟪RXサービスMTG用」（SVC_OPTIONS.menuTitle）。onOpen（CredentialHistory.gs）から呼ぶ。 */
function svcAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(SVC_OPTIONS.menuTitle)
    .addItem('リクエストを追加', 'openRequestPicker')
    .addToUi();
}

/**
 * メニュー「リクエストを追加」：選択パネル（サイドバー）を開く。
 * 開いているシートがサービスリクエスト（SVC_OPTIONS.requestSheet）でなければ開かずに知らせる。開いたら true。
 */
function openRequestPicker() {
  const ui = SpreadsheetApp.getUi();
  const active = SpreadsheetApp.getActiveSheet();
  if (!active || active.getName() !== SVC_OPTIONS.requestSheet) {
    ui.alert(SVC_OPTIONS.menuTitle, 'このシートでは開けません。「' + SVC_OPTIONS.requestSheet + '」のシートで開いてください。', ui.ButtonSet.OK);
    return false;
  }
  const html = HtmlService.createTemplateFromFile(SVC_OPTIONS.pickerTemplate).evaluate().setTitle(SVC_OPTIONS.pickerTitle);
  ui.showSidebar(html);
  return true;
}

/**
 * 選択パネルが読む内容（google.script.run から呼ぶ）。前の版の A〜D のプルダウンが残っていれば外し、
 * 登録した行を取り込み元に合わせてから、まだ登録していないリクエストを返す。
 * ほかの処理が実行中で合わせられないときは、合わせずに読むだけにする（パネルは開けるように）。
 */
function getRequestPickerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = svcRequestSheetOrThrow_(ss);
  const records = svcSourceRecords_(ss);
  let synced = true;
  try {
    taskWithLock_(() => {
      const legacy = !PropertiesService.getDocumentProperties().getProperty(SVC_OPTIONS.legacyDoneKey);
      svcRemoveLegacyLists_(ss, sheet);   // 前の版のプルダウンが残っていると、スクリプトの書き込みも止まるため（済んでいれば何もしない）
      svcSync_(ss, sheet, records, {notes: legacy});
      svcNotifyIds_(ss, svcEnsureIds_(sheet));   // ID の無い行・重複した ID（行のコピーなど）を直す
      svcLinkTaskServiceColumn_(ss);      // タスク管理のサービス案のプルダウンが、今のサービス案の列を見ているか
    }, 5000, 'パネルの読み込み');
  } catch (_) {
    synced = false;
  }
  return Object.assign(svcPickerData_(sheet, records), {synced});
}

/**
 * 選択パネルで選んだリクエストを登録する（google.script.run から呼ぶ）。key は getRequestPickerData() の items[].key。
 * 取り込み元とサービスリクエストを読み直し、まだ登録していなければ、最後の行（svcRegistered_）の下の A〜D に書く
 * （前の版の A〜D のプルダウンが残っていれば、先に外す）。
 * {ok, row, message, data: パネルの新しい内容}
 */
function addServiceRequest(key) {
  const result = svcAddRequest_(key);
  return {ok: result.ok, row: result.row, id: result.id || '', message: result.message, data: svcPickerData_(result.sheet, result.records)};
}

/**
 * 一覧のモーダルの「追加」（google.script.run から呼ぶ）。addServiceRequest と同じく追加し、一覧の新しい内容を返す。
 * {ok, row, message, list: getRequestListData() と同じ形}
 */
function addServiceRequestFromList(key) {
  const result = svcAddRequest_(key);
  return {ok: result.ok, row: result.row, id: result.id || '', message: result.message, list: svcListData_(result.sheet, result.records)};
}

/**
 * リクエストを1件登録する（ロックの中で読み直してから書く）。{ok, row, message, sheet, records}
 * 取り込み元に無い・すでに登録されているときは書かずに ok: false。
 */
function svcAddRequest_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return taskWithLock_(() => {
    const sheet = svcRequestSheetOrThrow_(ss);
    const records = svcSourceRecords_(ss);
    const record = records.find(rec => rec.key === String(key));
    if (!record) {
      return {ok: false, row: 0, sheet, records,
        message: '取り込み元にこのリクエストが見つかりません（書き換えか削除された可能性があります）。一覧を読み込み直しました。'};
    }
    const registered = svcRegistered_(sheet);
    if (registered.keys.has(record.key)) {
      return {ok: false, row: 0, sheet, records, message: 'このリクエストは、すでに登録されています。'};
    }

    svcEnsureIds_(sheet);   // 先に ID の重複を直してから、次の番号を決める
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    const row = registered.lastRow + 1;
    if (row > sheet.getMaxRows()) {
      // 行が足りなければ足し、判断のプルダウンも付ける
      const maxRows = sheet.getMaxRows();
      sheet.insertRowsAfter(maxRows, row - maxRows);
      taskApplyColumnRules_(sheet, SVC_OPTIONS.requestColumns.filter(c => c.options), cols, maxRows + 1, row - maxRows);
    }
    svcRemoveLegacyLists_(ss, sheet);
    svcSourceKeys_().forEach((k, i) => {
      if (!cols[k]) return;   // 列がまだ無い（setupRequestSheet() の前）なら書かない
      // 先頭に ' を付けて書く（「=…」のリクエストも数式にならない）
      sheet.getRange(row, cols[k]).setValue(credText_(record.values[i]));
    });
    let id = '';
    if (cols.id) {
      id = svcFormatId_(svcTakeIds_(1, svcMaxIdInSheet_(sheet, cols.id))[0]);
      sheet.getRange(row, cols.id).setValue(credText_(id));
    }
    return {ok: true, row, id, sheet, records, message: (id ? id + ' として' : '') + row + '行目に追加しました。'};
  }, SVC_OPTIONS.pickerLockWaitMs, 'リクエストの追加');
}

/**
 * 一覧のモーダルの「削除」（google.script.run から呼ぶ）：登録済みのリクエストの行を、サービスリクエストのシートから削除する。
 * row は一覧を読んだときの行番号、key はその行のリクエストのキー、id はその行の ID（あれば）。
 * 同じ ID で同じリクエストの行があればその行を、無ければ同じキーの行を探して削除する
 * （並べ替え・行の追加などで行がずれていてもよい。1つに決まらないとき・見つからないときは削除しない）。
 * 判断・サービス案などの入力も行ごと消える（確認はモーダルで行う）。{ok, row, message, list}
 */
function deleteServiceRequest(row, key, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return taskWithLock_(() => {
    const sheet = svcRequestSheetOrThrow_(ss);
    const matches = svcRegistered_(sheet).rows.filter(r => r.key === String(key));
    const number = svcIdNumber_(id);
    const byId = number ? matches.filter(r => svcIdNumber_(r.values.id) === number) : [];
    const target = (byId.length === 1 ? byId[0] : null) ||
      matches.find(r => r.row === Number(row)) || (matches.length === 1 ? matches[0] : null);
    if (!target) {
      return {ok: false, row: 0, list: svcListData_(sheet, svcSourceRecords_(ss)), message: matches.length
        ? '同じリクエストの行が複数あるため、どの行か決められませんでした。一覧を読み込み直しました。'
        : 'この行が見つかりません（ほかの人が削除・変更した可能性があります）。一覧を読み込み直しました。'};
    }
    // すべての行を消すことはできないので、足りなければ空の行を足しておく
    if (sheet.getMaxRows() - 1 <= Math.max(sheet.getFrozenRows(), SVC_OPTIONS.headerRow)) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    }
    sheet.deleteRow(target.row);
    const label = target.values.id ? target.values.id + '（' + target.row + '行目）' : target.row + '行目';
    return {ok: true, row: target.row, list: svcListData_(sheet, svcSourceRecords_(ss)), message: label + 'を削除しました。'};
  }, SVC_OPTIONS.pickerLockWaitMs, 'リクエストの削除');
}

/** サービスリクエストのシート。無ければ、管理者に setupRequestSheet() を頼むよう知らせて止める。 */
function svcRequestSheetOrThrow_(ss) {
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) throw new Error('「' + SVC_OPTIONS.requestSheet + '」シートがありません。管理者に setupRequestSheet() の実行を頼んでください。');
  return sheet;
}

/**
 * パネルに出す内容。{items: まだ登録していないリクエスト [{key, sheet, values: [サブインダストリー, 得意先, 案件名, リクエスト]}],
 *  total, registered, remaining, sources: 読むシート, requestSheet, loadedAt}
 */
function svcPickerData_(sheet, records) {
  const registered = svcRegistered_(sheet);
  const items = records.filter(rec => !registered.keys.has(rec.key))
    .map(rec => ({key: rec.key, sheet: rec.sheet, values: rec.values.slice()}));
  return {
    items,
    total: records.length,
    registered: records.length - items.length,
    remaining: items.length,
    sources: svcSourceSheets_(),
    requestSheet: SVC_OPTIONS.requestSheet,
    loadedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm')
  };
}

/* ---------------- 一覧のモーダル（パネルの件数のバッジから開く） ---------------- */

/**
 * 未登録・登録済みのリクエストの一覧をモーダルで開く（選択パネルのバッジから google.script.run で呼ぶ）。
 * kind：'unregistered'（未登録）か 'registered'（登録済み）。開いたときのタブになる。大きさはクレデンシャル履歴と同じ。
 */
function openRequestListDialog(kind) {
  const size = credDialogSize_(credActiveEmail_());
  const template = HtmlService.createTemplateFromFile(SVC_OPTIONS.listTemplate);
  template.kind = kind === 'registered' ? 'registered' : 'unregistered';
  template.openedWidth = size.width;
  template.openedHeight = size.height;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(size.width).setHeight(size.height),
    ' '   // 見出しはモーダルの中に表示する
  );
}

/**
 * 一覧のモーダルが読む内容（google.script.run から呼ぶ。読むだけなのでロックは取らない）。
 * {unregistered: [{key, sheet, values: [サブインダストリー, 得意先, 案件名, リクエスト]}],
 *  registered: [{row, key, sheet: 取り込み元のシート（見つからなければ空）, missing: 取り込み元に見つからないか,
 *                values: [サブインダストリー, 得意先, 案件名, リクエスト, 営業の4列], decision, feedback, service, department, owner}],
 *  requestSheet, loadedAt}
 */
function getRequestListData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return svcListData_(svcRequestSheetOrThrow_(ss), svcSourceRecords_(ss));
}

/** getRequestListData() の中身（追加・削除のあとの一覧にも使う）。 */
function svcListData_(sheet, records) {
  const sourceSheet = new Map(records.map(rec => [rec.key, rec.sheet]));
  const registered = svcRegistered_(sheet);
  const sourceKeys = svcSourceKeys_();
  return {
    unregistered: records.filter(rec => !registered.keys.has(rec.key))
      .map(rec => ({key: rec.key, sheet: rec.sheet, values: rec.values.slice()})),
    registered: registered.rows.map(r => ({
      row: r.row,
      key: r.key,
      id: r.values.id || '',
      sheet: sourceSheet.get(r.key) || '',
      missing: !sourceSheet.has(r.key),
      values: sourceKeys.map(k => r.values[k] || ''),
      decision: r.values.decision || '',
      feedback: r.values.feedback || '',
      service: r.values.service || '',
      department: r.values.department || '',
      owner: r.values.owner || ''
    })),
    requestSheet: SVC_OPTIONS.requestSheet,
    loadedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm')
  };
}

/* ---------------- ID（SR-0001） ---------------- */

/** ID の文字（12 → SR-0012）。 */
function svcFormatId_(number) {
  return SVC_OPTIONS.idPrefix + String(number).padStart(SVC_OPTIONS.idDigits, '0');
}

/** ID の番号（SR-0012・SR12・sr-12・12 → 12。全角・半角と空白の違いは無視。読めなければ 0）。 */
function svcIdNumber_(value) {
  const text = credNormalize_(value).toUpperCase();
  const prefix = credNormalize_(SVC_OPTIONS.idPrefix).toUpperCase().replace(/-+$/, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = new RegExp('^(?:' + prefix + '-?)?0*(\\d+)$').exec(text);
  return match ? Number(match[1]) : 0;
}

/** シートの ID の列にある、いちばん大きい番号（リクエストの無い行も数える）。 */
function svcMaxIdInSheet_(sheet, idCol) {
  const count = sheet.getLastRow() - SVC_OPTIONS.headerRow;
  if (count < 1) return 0;
  return sheet.getRange(SVC_OPTIONS.headerRow + 1, idCol, count, 1).getDisplayValues()
    .reduce((max, line) => Math.max(max, svcIdNumber_(line[0])), 0);
}

/**
 * 新しい ID の番号を count 個とって、これまでに振った最大の番号として覚える（ロック取得中に呼ぶこと）。
 * シートにある最大の番号と、覚えている番号の大きいほうの次から（手で書いた ID とも、削除した行の ID とも重ならない）。
 */
function svcTakeIds_(count, maxInSheet) {
  const props = PropertiesService.getDocumentProperties();
  const last = Math.max(Number(props.getProperty(SVC_OPTIONS.idCounterKey)) || 0, maxInSheet || 0);
  props.setProperty(SVC_OPTIONS.idCounterKey, String(last + count));
  return Array.from({length: count}, (_, i) => last + 1 + i);
}

/**
 * リクエストのある行のうち、ID の無い行と、上の行と同じ ID の行（行のコピーなど）に新しい ID を振る（ロック取得中に呼ぶこと）。
 * 上の行の ID は残す。読めない ID（SR-0001 の形でないもの）も振り直す。ID の列が無ければ何もしない。
 * {assigned: ID の無かった行の数, renumbered: [{row, from, to}]（振り直した行）}
 */
function svcEnsureIds_(sheet) {
  const headerRow = SVC_OPTIONS.headerRow;
  const result = {assigned: 0, renumbered: []};
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns.filter(c => c.key === 'id' || c.key === 'request'), headerRow);
  const count = sheet.getLastRow() - headerRow;
  if (!cols.id || !cols.request || count < 1) return result;
  const read = col => sheet.getRange(headerRow + 1, col, count, 1).getDisplayValues().map(line => String(line[0]).trim());
  const ids = read(cols.id);
  const requests = read(cols.request);
  const seen = new Set();
  const targets = [];
  ids.forEach((id, i) => {
    if (!requests[i]) return;
    const number = svcIdNumber_(id);
    if (number && !seen.has(number)) { seen.add(number); return; }
    targets.push(i);
  });
  if (!targets.length) return result;

  const numbers = svcTakeIds_(targets.length, ids.reduce((max, id) => Math.max(max, svcIdNumber_(id)), 0));
  const next = ids.slice();
  targets.forEach((i, n) => {
    next[i] = svcFormatId_(numbers[n]);
    if (ids[i]) result.renumbered.push({row: headerRow + 1 + i, from: ids[i], to: next[i]});
    else result.assigned++;
  });
  const first = targets[0];
  const last = targets[targets.length - 1];
  sheet.getRange(headerRow + 1 + first, cols.id, last - first + 1, 1)
    .setValues(next.slice(first, last + 1).map(v => [credText_(v)]));
  return result;
}

/** 振り直した ID の知らせ（多いときは先頭の5件だけ）。 */
function svcRenumberedMessage_(renumbered) {
  const shown = renumbered.slice(0, 5).map(r => r.row + '行目：' + r.from + ' → ' + r.to);
  return 'ほかの行と同じ ID（または読めない ID）だった ' + renumbered.length + '行に、新しい ID を振りました（' +
    shown.join('、') + (renumbered.length > shown.length ? ' ほか' : '') + '）。';
}

/** パネルを開いたときに ID を振り直したら、トーストで知らせる（ID の無かった行に振っただけなら知らせない）。 */
function svcNotifyIds_(ss, ids) {
  if (ids.renumbered.length) ss.toast(svcRenumberedMessage_(ids.renumbered), SVC_OPTIONS.menuTitle, 10);
}

/**
 * サービスリクエストに登録したリクエストのキーと、SVC_OPTIONS.requestColumns の列（A〜H）のどれかに値のある最後の行。
 * {keys: Set, lastRow, rows: [{row, key, values: {requestColumns の key: 表示の値}}]（リクエストのある行）}
 * 判断だけを書いた行などにリクエストを書き足さないよう、最後の行はすべての列で見る（ほかの列は見ない）。
 */
function svcRegistered_(sheet) {
  const headerRow = SVC_OPTIONS.headerRow;
  const keys = new Set();
  const count = sheet.getLastRow() - headerRow;
  const rows = [];
  if (count < 1) return {keys, lastRow: headerRow, rows};
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const sourceKeys = svcSourceKeys_();
  const used = Object.keys(cols).map(k => cols[k]).filter(Boolean);
  const left = Math.min.apply(null, used);
  const right = Math.max.apply(null, used);
  const lines = sheet.getRange(headerRow + 1, left, count, right - left + 1).getDisplayValues();
  let lastRow = headerRow;
  lines.forEach((line, i) => {
    if (used.some(col => String(line[col - left]).trim())) lastRow = headerRow + 1 + i;
    const values = sourceKeys.map(k => (cols[k] ? String(line[cols[k] - left]).trim() : ''));
    if (!values[3]) return;
    const key = svcRequestKey_(values[1], values[2], values[3]);
    keys.add(key);
    const all = {};
    Object.keys(cols).forEach(k => { all[k] = cols[k] ? String(line[cols[k] - left]).trim() : ''; });
    rows.push({row: headerRow + 1 + i, key, values: all});
  });
  return {keys, lastRow, rows};
}

/* ---------------- 管理者がエディタから実行する ---------------- */

/**
 * サービスリクエストのシートを用意する（無ければ作る）。見出し・判断のプルダウンを付け、登録した行を取り込み元に合わせる。
 * 前の版の A〜D の連動プルダウンと、候補のシート（__REQUEST_LISTS）は外す。
 * 使わなくなった列（SVC_OPTIONS.removedColumns。メモ・追加日）があれば、確認してから削除する。
 * タスク管理シートがあれば「サービス案」の列（無ければ右端に足す）にもプルダウンを付ける。何度実行してもよい。
 */
function setupRequestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(svcSetupRequestSheet_(ss).join('\n'), SVC_OPTIONS.menuTitle, 10);
}

/** setupRequestSheet() の中身。知らせる文（行ごと）を返す（setupAfterSheetChange() からも使う）。 */
function svcSetupRequestSheet_(ss) {
  // 使わなくなった列を消すかは、ロックを取る前に聞く（答えるまでほかの処理を待たせないため）
  const existing = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  const oldColumns = existing ? taskFoundLabels_(existing, SVC_OPTIONS.removedColumns, SVC_OPTIONS.headerRow) : [];
  let removeOld = false;
  if (oldColumns.length) {
    const ui = SpreadsheetApp.getUi();
    removeOld = ui.alert(SVC_OPTIONS.menuTitle,
      '「' + SVC_OPTIONS.requestSheet + '」シートの「' + oldColumns.join('」「') + '」の列は使わなくなりました。\n' +
      'これらの列を削除します。ほかの列の値はそのまま残ります。削除する列に入っている値は消えます。\n\n削除しますか？',
      ui.ButtonSet.YES_NO) === ui.Button.YES;
  }

  const result = taskWithLock_(() => {
    const removed = removeOld && existing ? taskDeleteColumns_(existing, SVC_OPTIONS.removedColumns, SVC_OPTIONS.headerRow) : [];
    const ensured = Object.assign(svcEnsureSheet_(ss), {removed});
    const records = svcSourceRecords_(ss);
    const synced = svcSync_(ss, ensured.sheet, records);
    const ids = svcEnsureIds_(ensured.sheet);
    const rules = svcApplyRules_(ss, ensured.sheet);
    const legacy = svcRemoveLegacyLists_(ss, ensured.sheet, true);
    return Object.assign(ensured, synced, rules, {legacy, ids}, svcCountRequests_(ensured.sheet, records));
  }, SVC_OPTIONS.lockWaitMs, 'setupRequestSheet()');

  const name = SVC_OPTIONS.requestSheet;
  const lines = [result.created ? '「' + name + '」シートを作りました。' : '「' + name + '」シートを整えました。'];
  lines.push(svcCountMessage_(result));
  if (result.ids.assigned) lines.push('ID の無い ' + result.ids.assigned + '行に ID を振りました。');
  if (result.ids.renumbered.length) lines.push(svcRenumberedMessage_(result.ids.renumbered));
  if (result.updated) lines.push('登録済みの行を取り込み元に合わせました（' + result.updated + 'セル）。');
  if (result.notesRemoved) lines.push('前の版で「リクエスト」のセルに付けたメモを ' + result.notesRemoved + '件外しました。');
  if (result.legacy) {
    lines.push('前の版の A〜D のプルダウンと「' + SVC_OPTIONS.legacyListSheet + '」シートを外しました。' +
      'リクエストは、メニュー「' + SVC_OPTIONS.menuTitle + '」→「リクエストを追加」から登録します。');
  }
  if (result.removed.length) lines.push('「' + result.removed.join('」「') + '」の列を削除しました。');
  else if (oldColumns.length) lines.push('「' + oldColumns.join('」「') + '」の列は残しました（使いません。不要なら削除してください）。');
  if (result.added.length) lines.push('「' + result.added.join('」「') + '」の列を足しました。');
  if (result.taskAdded.length) lines.push(TASK_OPTIONS.sheet + ' に「' + result.taskAdded.join('」「') + '」の列を足しました。');
  return lines.concat(result.warnings);
}

/**
 * 前の版で全件を自動で取り込んだまま手を付けていない行（判断が未判断か空で、フィードバック・サービス案・主管本部・担当者が空）を削除し、
 * そのリクエストをまた選択パネルに出す。削除する前に件数を見せて確かめる。判断・フィードバック・サービス案・主管本部・担当者のどれかが入っている行は残す。
 */
function removeUntouchedRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) throw new Error('「' + SVC_OPTIONS.requestSheet + '」シートがありません。setupRequestSheet() を実行してください。');
  const ui = SpreadsheetApp.getUi();
  const count = svcUntouchedRows_(sheet).length;
  if (!count) {
    ui.alert(SVC_OPTIONS.menuTitle, '削除できる行はありません（どの行も、判断・フィードバック・サービス案・主管本部・担当者のどれかが入っています）。', ui.ButtonSet.OK);
    return;
  }
  const answer = ui.alert(SVC_OPTIONS.menuTitle,
    '判断が「' + SVC_OPTIONS.defaultDecision + '」か空で、フィードバック・サービス案・主管本部・担当者が空の行が ' + count + '件あります。\n' +
    'これらの行を削除して、そのリクエストをまた選択パネルに出しますか？（判断・フィードバック・サービス案・主管本部・担当者のどれかが入っている行は残します）',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const result = taskWithLock_(() => {
    const rows = svcUntouchedRows_(sheet);   // ロックの中で数え直す
    // すべての行を消すことはできないので、足りなければ空の行を足しておく
    if (sheet.getMaxRows() - rows.length <= Math.max(sheet.getFrozenRows(), SVC_OPTIONS.headerRow)) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    }
    // 下の行から、続いている行はまとめて削除する
    for (let end = rows.length - 1; end >= 0;) {
      let start = end;
      while (start > 0 && rows[start - 1] === rows[start] - 1) start--;
      sheet.deleteRows(rows[start], end - start + 1);
      end = start - 1;
    }
    return Object.assign({removed: rows.length}, svcCountRequests_(sheet, svcSourceRecords_(ss)));
  }, SVC_OPTIONS.lockWaitMs, 'removeUntouchedRequests()');
  ss.toast(result.removed + '行を削除しました。' + svcCountMessage_(result), SVC_OPTIONS.menuTitle, 10);
}

/** 手を付けていない行（リクエストがあり、判断が未判断か空で、フィードバック・サービス案・主管本部・担当者が空）の行番号（上から）。 */
function svcUntouchedRows_(sheet) {
  const headerRow = SVC_OPTIONS.headerRow;
  const count = sheet.getLastRow() - headerRow;
  if (count < 1) return [];
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const read = key => (cols[key]
    ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : new Array(count).fill(''));
  const decisions = read('decision');
  const filled = SVC_OPTIONS.inputKeys.map(read);
  const rows = [];
  read('request').forEach((request, i) => {
    if (!svcUntouched_(request, decisions[i], filled.map(values => values[i]))) return;
    rows.push(headerRow + 1 + i);
  });
  return rows;
}

/** 手を付けていない行か（リクエストがあり、判断が未判断か空で、others（フィードバック・サービス案・主管本部・担当者）がすべて空）。 */
function svcUntouched_(request, decision, others) {
  return !!request && (!decision || decision === SVC_OPTIONS.defaultDecision) && others.every(value => !value);
}

/** 登録した件数の知らせ（{total, selected, remaining, untouched}）。 */
function svcCountMessage_(result) {
  if (!result.total) {
    return svcSourceLabel_() + 'にリクエストが見つかりません。「' + SVC_OPTIONS.sourceHeaders[3] + '」の列にリクエストが入っているか確かめてください。';
  }
  const lines = [svcSourceLabel_() + 'のリクエスト ' + result.total + '件のうち ' + result.selected + '件を登録しています（まだ登録していないもの ' +
    result.remaining + '件）。'];
  if (!result.remaining && result.untouched) {
    lines.push('判断・フィードバック・サービス案・主管本部・担当者が空の行が ' + result.untouched + '件あります。前の版で自動で取り込んだ行なら、' +
      'エディタで removeUntouchedRequests() を実行すると、削除して選択パネルに戻せます。');
  }
  return lines.join('\n');
}

/* ---------------- シートの用意 ---------------- */

/** {sheet, created, added} */
function svcEnsureSheet_(ss) {
  return taskEnsureSheet_(ss, SVC_OPTIONS.requestSheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
}

/**
 * 判断のプルダウンを付け、前の版の保護を外す。フィードバック・サービス案・主管本部・担当者の列は自由に入力するので、
 * プルダウンを付けない（前の版で付けたもの・列を差し込んだときに左の列から引き継いだものは外す）。
 * タスク管理シートがあれば、その「サービス案」の列（リクエストに付けたサービス案から選ぶ）にもプルダウンを付ける。
 */
function svcApplyRules_(ss, sheet) {
  const warnings = [];
  const first = SVC_OPTIONS.headerRow + 1;
  const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
  if (rows > 0) {
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    warnings.push.apply(warnings, taskApplyColumnRules_(sheet, SVC_OPTIONS.requestColumns, cols, first, rows));
    SVC_OPTIONS.requestColumns.filter(c => c.source == null && !c.options && cols[c.key]).forEach(c => {
      sheet.getRange(first, cols[c.key], rows, 1).clearDataValidations();
    });
    svcRemoveLegacyProtections_(sheet);
  }

  let taskAdded = [];
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  if (task) {
    taskAdded = taskEnsureHeader_(task, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow);
    const taskRows = task.getMaxRows() - TASK_OPTIONS.headerRow;
    if (taskRows > 0) {
      taskApplyColumnRules_(task, TASK_OPTIONS.columns.filter(c => c.type === 'service'),
        taskColumnMap_(task, TASK_OPTIONS.columns, TASK_OPTIONS.headerRow), TASK_OPTIONS.headerRow + 1, taskRows);
    }
  }
  return {warnings, taskAdded};
}

/**
 * 前の版の連動プルダウン（A〜D の入力規則）と、候補を作っていた非表示のシート（__REQUEST_LISTS）を外す。
 * A〜D には入力規則を付けないので、残っているものは列ごと外す（候補に無い値を拒否する規則が残っていると、パネルから書けない）。
 * setupRequestSheet() のほか、選択パネルを開いたとき・追加するときにも呼ぶ（ロック取得中に呼ぶこと）。外したものがあれば true。
 * 一度確かめたら印（SVC_OPTIONS.legacyDoneKey）を付け、force でなければ次からは確かめない（入力規則を読むのは時間がかかるため）。
 */
function svcRemoveLegacyLists_(ss, sheet, force) {
  const props = PropertiesService.getDocumentProperties();
  if (!force && props.getProperty(SVC_OPTIONS.legacyDoneKey)) return false;
  let removed = false;
  const lists = ss.getSheetByName(SVC_OPTIONS.legacyListSheet);
  if (lists) {
    ss.deleteSheet(lists);
    removed = true;
  }
  const rows = sheet.getMaxRows() - SVC_OPTIONS.headerRow;
  if (rows > 0) {
    const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, SVC_OPTIONS.headerRow);
    // 前の版でプルダウンを付けていたのは A〜D（サブインダストリー・得意先・案件名・リクエスト）だけ
    svcSourceKeys_().slice(0, SVC_OPTIONS.sourceHeaders.length).forEach(key => {
      if (!cols[key]) return;
      const range = sheet.getRange(SVC_OPTIONS.headerRow + 1, cols[key], rows, 1);
      if (range.getDataValidations().some(line => line[0])) {
        range.clearDataValidations();
        removed = true;
      }
    });
  }
  props.setProperty(SVC_OPTIONS.legacyDoneKey, '1');
  return removed;
}

/** 前の版で付けた保護（A〜D を手で変えると警告が出る）を外す。 */
function svcRemoveLegacyProtections_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    let description = '';
    try { description = p.getDescription(); } catch (_) { return; }
    if (SVC_OPTIONS.legacyProtectDescriptions.indexOf(description) >= 0) p.remove();
  });
}

/**
 * タスク管理の「サービス案」の列の入力規則。サービスリクエストの「サービス案」の列を範囲で参照し、
 * そこに無い名前は入力できないようにする。サービスリクエストのシートが無ければ null。
 */
function svcServiceRule_(ss) {
  const sheet = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!sheet) return null;
  const col = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns.filter(c => c.key === 'service'), SVC_OPTIONS.headerRow).service;
  if (!col) return null;
  const letter = diffColumnLetter_(col);
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange(letter + (SVC_OPTIONS.headerRow + 1) + ':' + letter), true)
    .setAllowInvalid(false)
    .setHelpText('「' + SVC_OPTIONS.requestSheet + '」シートでリクエストに付けたサービス案から選んでください。')
    .build();
}

/**
 * タスク管理の「サービス案」のプルダウンが、サービスリクエストの今の「サービス案」の列を参照しているかを確かめ、
 * 参照していなければ列ごと付け直す（ロック取得中に呼ぶこと）。付け直したら true。
 * - サービス案の列（列の追加・削除で動く）か、タスク管理の行数が前に確かめたときと違えば、確かめずに付け直す（書くのは1回）
 * - 同じなら SVC_OPTIONS.taskLinkCheckMinutes 分に1回だけ、入力規則を1回で読んで確かめる（貼り付けでプルダウンが消えた行など）。
 *   行ごとにシート名などを問い合わせると、行の数だけ時間がかかる（1,000行で数十秒）ので、種類と参照する列だけを見る
 */
function svcLinkTaskServiceColumn_(ss) {
  const task = ss.getSheetByName(TASK_OPTIONS.sheet);
  const request = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  if (!task || !request) return false;
  const rows = task.getMaxRows() - TASK_OPTIONS.headerRow;
  if (rows < 1) return false;
  const col = taskColumnMap_(task, TASK_OPTIONS.columns.filter(c => c.type === 'service'), TASK_OPTIONS.headerRow).service;
  const serviceCol = taskColumnMap_(request, SVC_OPTIONS.requestColumns.filter(c => c.key === 'service'), SVC_OPTIONS.headerRow).service;
  if (!col || !serviceCol) return false;

  const props = PropertiesService.getDocumentProperties();
  const state = [request.getSheetId(), serviceCol, task.getSheetId(), col, rows].join(':');
  let saved = {};
  try { saved = JSON.parse(props.getProperty(SVC_OPTIONS.taskLinkKey) || '{}') || {}; } catch (_) {}
  const recent = Date.now() - (Number(saved.at) || 0) < SVC_OPTIONS.taskLinkCheckMinutes * 60 * 1000;
  if (saved.state === state && recent) return false;   // 列も行数も変わらず、少し前に確かめた

  const range = task.getRange(TASK_OPTIONS.headerRow + 1, col, rows, 1);
  let linked = false;
  if (saved.state === state) {
    const rules = range.getDataValidations().map(line => line[0]);
    const inRange = SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE;
    linked = rules.every(rule => rule && rule.getCriteriaType() === inRange) &&
      [0, Math.floor(rules.length / 2), rules.length - 1].every(i => rules[i].getCriteriaValues()[0].getColumn() === serviceCol);
  }
  if (!linked) range.setDataValidation(svcServiceRule_(ss));
  props.setProperty(SVC_OPTIONS.taskLinkKey, JSON.stringify({state, at: Date.now()}));
  return !linked;
}

/* ---------------- 「追加」が止まるときの確認（管理者がエディタから実行する） ---------------- */

/**
 * 選択パネルの「追加」が「他の処理が実行中です」で止まるときに実行する（読むだけ。何も書き換えない）。
 * サービス管理が使うロックの種類（貼り替えが済んでいるか）、今ロックを持っている処理、差分追跡などのドキュメントロックが使われているかを画面に出す。
 */
function diagnoseServiceLock() {
  const lines = svcLockLines_();
  console.log(lines.join('\n'));
  const ui = SpreadsheetApp.getUi();
  ui.alert(SVC_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/** diagnoseServiceLock() の中身。画面に出す文（行ごと）を返す。 */
function svcLockLines_() {
  const scriptLock = String(taskWithLock_).indexOf('getScriptLock') >= 0;
  const busy = lock => {
    if (!lock.tryLock(0)) return true;
    lock.releaseLock();
    return false;
  };
  return [
    scriptLock
      ? 'サービス管理のロック：差分追跡とは別のロック（新しい版）を使っています。'
      : 'サービス管理のロック：差分追跡と同じロック（前の版）を使っています。TaskManagement.gs を最新のものに貼り替えてください' +
        '（差分追跡と順番待ちになり、「追加」が止まります）。',
    'サービス管理のロック：' + (busy(LockService.getScriptLock()) ? '今、使われています' + taskLockHolderText_() : '空いています') + '。',
    '差分追跡などのロック：' + (busy(LockService.getDocumentLock())
      ? '今、使われています（取り込み元の編集の記録・取りこぼしの回収などが動いています）'
      : '空いています') + '。'
  ];
}

/* ---------------- パネルに出ないときの確認（管理者がエディタから実行する） ---------------- */

/**
 * 選択パネルにリクエストが出ない理由を調べて、画面に出す（読むだけ。何も書き換えない）。
 * DIFF_RULES の各シートについて、読む対象か・見出しがあるか・リクエストのある行の数と、パネルに出ないもの（登録済み・ほかと同じ）の数。
 */
function diagnoseRequestSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = svcDiagnoseLines_(ss);
  console.log(lines.join('\n'));
  const ui = SpreadsheetApp.getUi();
  ui.alert(SVC_OPTIONS.menuTitle, lines.join('\n'), ui.ButtonSet.OK);
}

/** diagnoseRequestSources() の中身。画面に出す文（行ごと）を返す。 */
function svcDiagnoseLines_(ss) {
  const lines = [];
  const request = ss.getSheetByName(SVC_OPTIONS.requestSheet);
  let registered = {keys: new Set()};
  if (!request) {
    lines.push('「' + SVC_OPTIONS.requestSheet + '」シートがありません。setupRequestSheet() を実行してください。');
  } else {
    try {
      registered = svcRegistered_(request);
    } catch (error) {
      lines.push(error.message);
    }
  }

  const seen = new Set();
  let complete = !!request;   // 読むシートをすべて読めたか（読めないシートがあると、見つからない行を数えられない）
  Object.keys(DIFF_RULES).forEach(name => {
    const rule = DIFF_RULES[name];
    if (!rule.requests) {
      lines.push('「' + name + '」：読みません（test.gs の DIFF_RULES に requests: true がありません）。');
      return;
    }
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      lines.push('「' + name + '」：シートがありません。');
      complete = false;
      return;
    }
    const headers = sheet.getRange(rule.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    const cols = SVC_OPTIONS.sourceHeaders.map(h => (taskHeaderCount_(headers, h) === 1 ? taskFindColumn_(headers, h, '', true) : 0));
    const missing = SVC_OPTIONS.sourceHeaders.filter((h, i) => !cols[i]);
    if (missing.length) {
      lines.push('「' + name + '」：' + rule.headerRow + '行目に「' + missing.join('」「') + '」の見出しがちょうど1つずつ必要です。');
      complete = false;
      return;
    }
    const count = sheet.getLastRow() - rule.headerRow;
    const values = count > 0
      ? sheet.getRange(rule.headerRow + 1, 1, count, Math.max.apply(null, cols)).getDisplayValues()
        .map(line => cols.map(col => String(line[col - 1]).trim()))
      : [];
    const withRequest = values.filter(v => v[3]);
    let duplicate = 0;
    let selected = 0;
    let candidates = 0;
    // svcSourceRecords_ と同じ順に見る（同じリクエストは最初の行だけ）
    withRequest.forEach(v => {
      const key = svcRequestKey_(v[1], v[2], v[3]);
      if (seen.has(key)) { duplicate++; return; }
      seen.add(key);
      if (registered.keys.has(key)) selected++; else candidates++;
    });
    const reasons = [];
    if (selected) reasons.push('登録済み ' + selected + '件');
    if (duplicate) reasons.push('ほかの行・シートと同じリクエスト ' + duplicate + '件');
    lines.push('「' + name + '」：リクエストのある行 ' + withRequest.length + '件 → パネルに出る ' + candidates + '件' +
      (reasons.length ? '（出ないもの：' + reasons.join('、') + '）' : '') + '。');
    const optional = (SVC_OPTIONS.sourceOptionalHeaders || []).filter(h => taskHeaderCount_(headers, h) !== 1);
    if (optional.length) {
      lines.push('「' + name + '」：「' + optional.join('」「') + '」の見出しがちょうど1つではないため、その列は空のまま登録します。');
    }
  });
  // 登録済みで、取り込み元に見つからない行（取り込み元で書き換えか削除された）
  if (complete && registered.rows) {
    const rows = registered.rows.filter(r => !seen.has(r.key)).map(r => r.row);
    if (rows.length) {
      const limit = 30;
      lines.push('「' + SVC_OPTIONS.requestSheet + '」で取り込み元に見つからない行：' + rows.slice(0, limit).join('・') + '行目' +
        (rows.length > limit ? ' ほか ' + (rows.length - limit) + '行' : '') + '（取り込み元で書き換えか削除された可能性があります）。');
    }
  }
  if (ss.getSheetByName(SVC_OPTIONS.legacyListSheet)) {
    lines.push('前の版の「' + SVC_OPTIONS.legacyListSheet + '」シートが残っています。setupRequestSheet() を実行すると外します。');
  }
  return lines;
}

/* ---------------- 取り込み元のリクエスト ---------------- */

/** リクエストを見分けるキー（得意先・案件名・リクエスト。全角・半角と空白の違いは無視）。 */
function svcRequestKey_(customer, project, request) {
  return [customer, project, request].map(credNormalize_).join('\u0001');
}

/** 読むシートの見出しの行（DIFF_RULES の headerRow）。 */
function svcSourceHeaderRow_(name) {
  return DIFF_RULES[name].headerRow;
}

/**
 * 取り込み元（読むシート。svcSourceSheets_）でリクエストが入っている行を、シート・行の順に返す（同じキーは1つ）。
 * [{sheet, values: [サブインダストリー, 得意先, 案件名, リクエスト, アカウント責任者部署, アカウント責任者, BX部署, BX担当], key}]。
 * 必須の見出し（sourceHeaders）が足りないときは止める。
 */
function svcSourceRecords_(ss) {
  const records = [];
  const seen = new Set();
  const names = svcSourceSheets_();
  if (!names.length) {
    throw new Error('リクエストを読むシートがありません。test.gs の DIFF_RULES で、読むシートに requests: true を付けてください。');
  }
  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('リクエストを読むシート「' + name + '」が見つかりません。');
    const headerRow = svcSourceHeaderRow_(name);
    const headers = sheet.getRange(headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    const required = SVC_OPTIONS.sourceHeaders.length;
    // 必須の列は見出しがちょうど1つでなければ止める。一緒に読む列（営業）は無ければ 0（空として扱う）
    const cols = svcAllSourceHeaders_().map((h, i) => taskFindColumn_(headers, h, name + ' の ' + headerRow + '行目', i < required));
    const count = sheet.getLastRow() - headerRow;
    if (count < 1) return;

    const found = cols.filter(Boolean);
    const left = Math.min.apply(null, found);
    const right = Math.max.apply(null, found);
    sheet.getRange(headerRow + 1, left, count, right - left + 1).getDisplayValues().forEach(line => {
      const values = cols.map(col => (col ? String(line[col - left] == null ? '' : line[col - left]).trim() : ''));
      if (!values[3]) return;
      const key = svcRequestKey_(values[1], values[2], values[3]);
      if (seen.has(key)) return;
      seen.add(key);
      records.push({sheet: name, values, key});
    });
  });
  return records;
}

/**
 * 登録した行を取り込み元に合わせる（行は足さない。ロック取得中に呼ぶこと）。
 * - 取り込み元から書く列のうち、リクエストを見分ける列（得意先・案件名・リクエスト）のほか
 *   （サブインダストリー・アカウント責任者部署・アカウント責任者・BX部署・BX担当）を、取り込み元の同じリクエストの値にそろえる
 *   （全角・半角と空白だけの違いなら変えない。取り込み元に無い行は変えない）
 * - options.notes が false でなければ、前の版で「リクエスト」のセルに付けたメモ（SVC_OPTIONS.legacyNotes で始まるもの）を外す
 *   （利用者が書いたメモは変えない。パネルを開いたときは、前の版の後片付けが済むまでだけ）
 * 取り込み元に無い行には何もしない（メモは付けない。行番号は diagnoseRequestSources() で確かめる）。
 * 読み取りは取り込み元から書く列をまとめて1回（列ごとに読むと遅いため）。{updated: 書き直したセルの数, notesRemoved}
 */
function svcSync_(ss, sheet, records, options) {
  const headerRow = SVC_OPTIONS.headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const source = new Map(records.map(r => [r.key, r.values]));
  const count = sheet.getLastRow() - headerRow;
  let updated = 0;
  let notesRemoved = 0;
  if (count < 1) return {updated, notesRemoved};

  const sourceColumns = SVC_OPTIONS.requestColumns.filter(c => c.source != null && cols[c.key]);
  const used = sourceColumns.map(c => cols[c.key]);
  const left = Math.min.apply(null, used);
  const right = Math.max.apply(null, used);
  const block = sheet.getRange(headerRow + 1, left, count, right - left + 1).getDisplayValues();
  const read = key => block.map(line => String(line[cols[key] - left]).trim());
  const [customers, projects, requests] = ['customer', 'project', 'request'].map(read);
  const keys = requests.map((request, i) => (request ? svcRequestKey_(customers[i], projects[i], request) : ''));

  // 見分ける列のほかを、取り込み元に合わせる。書き直すのは変わった行の範囲だけ（列ごとにまとめて書く）
  sourceColumns
    .filter(c => ['customer', 'project', 'request'].indexOf(c.key) < 0)
    .forEach(c => {
      const current = read(c.key);
      const next = current.slice();
      const changed = [];
      keys.forEach((key, i) => {
        const values = key && source.get(key);
        if (!values || credNormalize_(values[c.source]) === credNormalize_(current[i])) return;
        next[i] = values[c.source];
        changed.push(i);
      });
      if (!changed.length) return;
      const first = changed[0];
      const last = changed[changed.length - 1];
      sheet.getRange(headerRow + 1 + first, cols[c.key], last - first + 1, 1)
        .setValues(next.slice(first, last + 1).map(v => [credText_(v)]));
      updated += changed.length;
    });

  // 前の版で付けたメモを外す
  if (options && options.notes === false) return {updated, notesRemoved};
  const noteRange = sheet.getRange(headerRow + 1, cols.request, count, 1);
  const notes = noteRange.getNotes().map(r => String(r[0] || ''));
  const nextNotes = notes.map(note => {
    if (!note || !SVC_OPTIONS.legacyNotes.some(prefix => note.indexOf(prefix) === 0)) return note;
    notesRemoved++;
    return '';
  });
  if (notesRemoved) noteRange.setNotes(nextNotes.map(n => [n]));
  return {updated, notesRemoved};
}

/** 登録した件数を数える。{total, selected, remaining, untouched}（svcCountMessage_ で知らせる）。 */
function svcCountRequests_(sheet, records) {
  const headerRow = SVC_OPTIONS.headerRow;
  const count = sheet.getLastRow() - headerRow;
  const cols = taskColumnMap_(sheet, SVC_OPTIONS.requestColumns, headerRow);
  const read = key => (count > 0 && cols[key]
    ? sheet.getRange(headerRow + 1, cols[key], count, 1).getDisplayValues().map(r => String(r[0]).trim())
    : []);
  const customers = read('customer');
  const projects = read('project');
  const decisions = read('decision');
  const filled = SVC_OPTIONS.inputKeys.map(read);
  const used = new Set();
  let selected = 0;
  let untouched = 0;
  read('request').forEach((request, i) => {
    if (!request) return;
    selected++;
    used.add(svcRequestKey_(customers[i], projects[i], request));
    if (svcUntouched_(request, decisions[i], filled.map(values => values[i]))) untouched++;
  });
  return {
    total: records.length,
    selected,
    remaining: records.filter(rec => !used.has(rec.key)).length,
    untouched
  };
}
