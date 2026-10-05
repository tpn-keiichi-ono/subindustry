/**
 * クレデンシャル / オファリング履歴（得意先別）
 * 差分追跡スクリプト（v5、test.gs）と同じ Apps Script プロジェクトに置くファイル。
 *
 * 仕組み
 * - 追跡シートの「クレデンシャル（改行）オファリング登録」列は各行の起動ボタン（チェックボックス）
 * - チェックすると onCredentialLauncherEdit()（インストール型 onEdit）が動き、
 *   その行の「得意先」に合致する履歴をモーダルで表示する
 * - 履歴は専用シート「クレデンシャル履歴_記録」に1件1行で蓄積する
 * - モーダルから 登録・編集・削除 ができる。操作はすべて 変更履歴_差分 にも
 *   イベントID「CRD-…」で記録される（auditToDiffLog）
 *
 * 記録シートの列は「見出し名」で識別する。列の並べ替えは自由だが、見出し名は変えないこと。
 * 足りない列（ID・宛先・更新者 など）は、次に開いたとき右端に自動で追加される。
 *
 * セットアップ（トリガー所有者のアカウントで）
 * 1. DIFF_RULES の ranges から「クレデンシャル（改行）オファリング登録」列を外し、setupDiffTracking() を実行
 * 2. setupCredentialLauncher() を実行
 * 行を増やしたら setupCredentialLauncher() を再実行すると、ボタンが最終行まで補充される。
 *
 * 注意：パスワード・APIキー・トークンなどの秘密情報そのものは入力しないこと。
 *
 * 記録シートへの文字列は先頭に ' を付けて書き込む（数式・日付への自動変換を防ぐ）。
 */

const CRED_OPTIONS = {
  // ボタン列の見出し。セル内の改行・空白は照合時に無視する
  // （「クレデンシャル」改行「オファリング登録」と入力したセルに一致する）
  launcherHeader: 'クレデンシャル\nオファリング登録',
  customerHeader: '得意先',               // 照合に使う列の見出し
  logSheet: 'クレデンシャル履歴_記録',
  // メニューの名前（画面の案内にもこの名前を出す）。項目は「履歴サイドバーを開く」「権限を承認する（初回のみ）」だけ
  menuTitle: '🟩RXビジネスMTG用',
  // アカウントプランシート作成（コード.gs）のメニューを出すか。出さないときは、エディタから showLoadingDialog() を実行する
  accountPlanMenu: false,
  dialogTitle: 'クレデンシャル・オファリング履歴',   // モーダル上部に「タイトル：得意先名」で表示
  dialogWidth: 1280,        // 開いた直後の大きさ（表示後、ブラウザの広さに合わせて自動で広がる）
  dialogHeight: 760,
  lockWaitMs: 30000,
  columnWidthsKey: 'CRED_COLUMN_WIDTHS_V1',   // 利用者ごとの列幅（ユーザープロパティ）
  dialogSizePrefix: 'CRED_DIALOG_SIZE_V1_',   // 利用者ごとのモーダルの大きさ（ブラウザ幅の95%）
  triggerHandler: 'onCredentialLauncherEdit',
  deleteMode: 'soft',        // 'soft' = 削除者・削除日時を残して非表示 / 'hard' = 行ごと削除
  auditToDiffLog: true,      // 登録・編集・削除を 変更履歴_差分 にも記録する
  auditEventPrefix: 'CRD',
  // 予定と実績：実施日が空なら「予定」、入っていれば「実績」（どちらか一方は必須）
  schedule: {
    planKey: 'plan',
    doneKey: 'date'
  },
  // オファリングをクレデンシャルに紐づける設定（1つのクレデンシャルに複数のオファリング可）
  linking: {
    kindKey: 'kind',                     // 種別の項目 key
    parentKind: 'クレデンシャル',
    childKind: 'オファリング',
    header: '紐づくクレデンシャルID',   // 記録シートの見出し
    label: '紐づけるクレデンシャル',     // フォームの表示名
    copyKeys: ['title', 'person', 'owner', 'ownerDept'],   // 「オファリングを追加」で親から引き継ぐ項目
    labelKeys: ['date|plan', 'name', 'person'] // 紐づけ先の表示（a|b は a が空なら b）
  },
  // 以前の版の列。新しい項目が空なら、この列の値を表示・編集に使う
  legacyColumns: {'宛先': 'person'},
  // 見出しを変えた列。古い見出しがあれば新しい見出しに書き換えて使い続ける
  renamedColumns: {'日付': '実施日', '備考': 'フリーコメント'},
  legacySeparator: ' | ',                              // 初版（セル内追記）の区切り
  legacyFieldKeys: ['date', 'kind', 'name', 'note'],   // 初版の項目並び
  // 入力項目（label が記録シートの見出しになる）
  fields: [
    {key: 'plan', label: '予定日', type: 'date', required: false},
    {key: 'date', label: '実施日', type: 'date', required: false},
    {key: 'kind', label: '種別', type: 'select', required: true,
      options: ['クレデンシャル', 'オファリング']},
    {key: 'title', label: '役職', type: 'text', required: false, maxLength: 100,
      placeholder: '例：購買部 部長'},
    {key: 'person', label: '氏名', type: 'text', required: true, maxLength: 100,
      placeholder: '例：山田 太郎'},
    {key: 'name', label: '名称', type: 'text', required: true, maxLength: 200},
    {key: 'note', label: 'フリーコメント', type: 'textarea', maxLength: 2000},
    // 当社の担当者
    {key: 'owner', label: '担当者名', type: 'text', required: true, maxLength: 100,
      placeholder: '例：佐藤 一郎'},
    {key: 'ownerDept', label: '担当部署', type: 'text', required: false, maxLength: 100,
      placeholder: '例：第一営業部'}
  ]
};

/* ---------------- メニュー ---------------- */

/**
 * 単純トリガー。onOpen() はプロジェクト全体でこの1つだけにすること
 * （同じ名前の関数が2つあると、どちらか一方しか動かない）。
 * 別のメニューを増やすときは、そのメニューを作る関数をここから呼ぶ。
 */
function onOpen() {
  // メニューは追加した順に左から並ぶ：🟩RXビジネスMTG用 → 🟪RXサービスMTG用（→ アカウントプランシート作成）
  credAddMenu_();
  if (typeof svcAddMenu_ === 'function') svcAddMenu_();   // ServiceManagement.gs（サービスリクエストの選択パネル）
  // コード.gs（アカウントプランシート作成）。CRED_OPTIONS.accountPlanMenu が true のときだけ出す
  if (CRED_OPTIONS.accountPlanMenu && typeof apAddMenu_ === 'function') apAddMenu_();
}

/**
 * 利用者向けのメニューは、履歴サイドバーを開く・権限の承認の2つだけにする。
 * 管理者の作業（ボタン列の設定・サンプル履歴など）や、選んだ行のモーダルを開く処理は、
 * Apps Script エディタから関数を直接実行する（docs/OPERATIONS.md）。
 */
function credAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(CRED_OPTIONS.menuTitle)
    .addItem('履歴サイドバーを開く', 'openHistorySidebar')
    .addItem('権限を承認する（初回のみ）', 'authorizeHistoryFeatures')
    .addToUi();
}

/**
 * メニュー「権限を承認する（初回のみ）」。
 * まだ承認していない人がメニューから実行すると、Apps Script が承認の画面を出す。
 * サイドバーからは承認の画面を出せないため、ここ（とメニューの「履歴サイドバーを開く」）が入口になる。
 * 承認が済んだら、対象のシートならサイドバーを開き直して履歴を読み込ませる（対象外のシートでは開かない）。
 */
function authorizeHistoryFeatures() {
  // 承認の画面で一部の権限のチェックを外した人には、足りない権限の承認の画面をもう一度出す
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (hsIsTargetSheet_(ss.getActiveSheet())) {
    hsShowSidebar_();
    ss.toast('権限は承認済みです。履歴サイドバーを開きました。', CRED_OPTIONS.menuTitle, 8);
  } else {
    ss.toast('権限は承認済みです。履歴サイドバーは' + hsTargetSheetsLabel_() + 'のシートで開いてください。', CRED_OPTIONS.menuTitle, 8);
  }
}

/** エディタから実行して、選んでいる行のモーダルを開く（ボタンが使えないときの代替）。 */
function openCredentialDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    if (!range) throw new Error('対象行のセルを選択してください。');
    if (range.getNumRows() > 1) throw new Error('1行だけ選択してください。');
    credShowDialog_(sheet, range.getRow(), credActiveEmail_());
  } catch (error) {
    ui.alert(CRED_OPTIONS.dialogTitle, error.message, ui.ButtonSet.OK);
  }
}

/* ---------------- 起動ボタン（インストール型 onEdit） ---------------- */

function onCredentialLauncherEdit(e) {
  // チェックを入れた単一セルの編集だけを扱う（範囲貼り付けでは e.value が無い）
  if (!e || !e.range || e.value !== 'TRUE') return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const sheet = e.range.getSheet();
  const rule = DIFF_RULES[sheet.getName()];
  const row = e.range.getRow();
  if (!rule || row <= rule.headerRow) return;

  // 起動を速くするため、ここでは見出し行を1回読むだけにして、すぐモーダルを開く。
  // 得意先の確認や履歴の読み込みは、開いたモーダルの中で行う。
  const kind = credLauncherKind_(sheet, rule, e.range.getColumn());
  if (!kind) return;

  let email = '';
  try { email = e.user ? e.user.getEmail() : ''; } catch (_) {}

  try {
    credOpenLauncherDialog_(kind, sheet.getName(), row, email);
  } catch (error) {
    try { e.source.toast(error.message, kind === 'change' ? '変更履歴' : CRED_OPTIONS.dialogTitle, 10); } catch (_) {}
  } finally {
    e.range.setValue(false);   // ボタンとして使うので戻す（モーダルを開いたあとに行う）
  }
}

/**
 * 見出し行を1回だけ読み、その列がどのボタン列かを返す。
 * 'credential'（クレデンシャル・オファリング履歴）／'change'（変更履歴）／''（ボタン列ではない）
 */
function credLauncherKind_(sheet, rule, column) {
  const headers = sheet.getRange(rule.headerRow + ':' + rule.headerRow).getDisplayValues()[0];
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  const label = normalize(headers[column - 1]);
  if (!label) return '';
  if (label === normalize(CRED_OPTIONS.launcherHeader)) return 'credential';
  if (typeof CHG_OPTIONS !== 'undefined' && label === normalize(CHG_OPTIONS.launcherHeader)) return 'change';
  return '';
}

/** ボタンから開くときの最短の処理（読み込みはモーダル側）。 */
function credOpenLauncherDialog_(kind, sheetName, row, email) {
  const size = credDialogSize_(email);
  const template = HtmlService.createTemplateFromFile(
    kind === 'change' ? CHG_OPTIONS.dialogTemplate : 'CredentialDialog');
  template.sheetName = sheetName;
  template.row = row;
  template.customer = '';      // 得意先名はモーダルが読み込んで表示する
  template.focusId = '';
  template.focusEventId = '';
  template.initial = '';
  template.openedWidth = size.width;
  template.openedHeight = size.height;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(size.width).setHeight(size.height),
    ' '
  );
}

/** recordDiffEdit() からも呼ばれる。いずれかのボタン列の単一列編集なら true（差分追跡の対象外）。 */
function credIsLauncherEdit_(e, rule) {
  try {
    if (e.range.getNumColumns() !== 1) return false;
    return !!credLauncherKind_(e.range.getSheet(), rule, e.range.getColumn());
  } catch (_) {
    return false;
  }
}

/** options.focusId を渡すと、その履歴を選んだ状態（編集パネルを開いた状態）で開く。 */
function credShowDialog_(sheet, row, email, options) {
  const target = credResolveTarget_(sheet, row);

  // 前回その人のブラウザで測った大きさ（幅の95%）で開く。初回は既定の大きさ
  const size = credDialogSize_(email);

  // モーダルはすぐに開き、「読み込み中」を表示しながら履歴を読み込む
  const template = HtmlService.createTemplateFromFile('CredentialDialog');
  template.sheetName = target.sheet.getName();
  template.row = target.row;
  template.initial = '';
  template.customer = target.customer;
  template.focusId = (options && options.focusId) || '';
  template.openedWidth = size.width;
  template.openedHeight = size.height;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate()
      .setWidth(size.width)
      .setHeight(size.height),
    ' '   // 見出し（タイトル・得意先名・件数）はモーダルの中に表示する
  );
}

/* ---------------- 管理者用セットアップ ---------------- */

function setupCredentialLauncher() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(CRED_OPTIONS.lockWaitMs);

  try {
    // 先に全シートを検証してから変更する。ボタン列の無いシートは飛ばす（ボタン列は使うシートだけに置く）
    const skipped = [];
    const plans = Object.entries(DIFF_RULES).map(([name, rule]) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error('シートが見つかりません: ' + name);

      const launcher = credColumnByHeader_(sheet, rule, CRED_OPTIONS.launcherHeader, true);
      if (!launcher) {
        skipped.push(name);
        return null;
      }
      const customer = credColumnByHeader_(sheet, rule, CRED_OPTIONS.customerHeader);
      const stamp = diffStampColumn_(sheet, rule);
      const firstRow = rule.headerRow + 1;
      const rowCount = Math.max(sheet.getMaxRows() - rule.headerRow, 1);
      const column = sheet.getRange(firstRow, launcher, rowCount, 1);

      // チェックの ON/OFF が差分ログに残らないよう、追跡範囲外であることを必須にする
      if (diffBlocks_(sheet, rule, column, stamp, false).length) {
        throw new Error(
          name + ' の「' + credHeaderText_(CRED_OPTIONS.launcherHeader) + '」列（' +
          diffColumnLetter_(launcher) + '列）が DIFF_RULES の ranges に含まれています。' +
          '範囲から外して setupDiffTracking() を実行してから、もう一度実行してください。'
        );
      }
      return {sheet, launcher, customer, firstRow};
    }).filter(Boolean);

    const header = '「' + credHeaderText_(CRED_OPTIONS.launcherHeader) + '」';
    if (!plans.length) {
      const message = 'ボタン列（' + header + '）のあるシートが無いので、ボタンは付けていません。';
      ss.toast(message, CRED_OPTIONS.dialogTitle, 10);
      return message;
    }

    const {log, cols} = credLogSheet_(ss, true);
    const timezone = ss.getSpreadsheetTimeZone();
    let migrated = 0;
    let buttons = 0;

    for (const p of plans) {
      const lastRow = p.sheet.getLastRow();
      if (lastRow < p.firstRow) continue;

      const count = lastRow - p.firstRow + 1;
      const range = p.sheet.getRange(p.firstRow, p.launcher, count, 1);
      const values = range.getValues();
      const customers = p.sheet
        .getRange(p.firstRow, p.customer, count, 1)
        .getDisplayValues();

      // 初版でセルに書かれていた履歴を記録シートへ移す
      const entries = [];
      values.forEach((v, i) => {
        const value = v[0];
        if (value === '' || typeof value === 'boolean') return;
        const source = p.sheet.getName() + '!' + diffA1_(p.firstRow + i, p.launcher) + '（移行）';
        credParseLegacy_(String(value), timezone).forEach(entry => {
          entries.push(Object.assign(entry, {
            id: Utilities.getUuid(), customer: customers[i][0], source
          }));
        });
      });

      if (entries.length) {
        credAppendLog_(log, cols, entries);
        SpreadsheetApp.flush();
        migrated += entries.length;
      }

      range.clearContent();
      range.clearDataValidations();
      range.insertCheckboxes();
      buttons += count;
    }

    credInstallTrigger_(ss);
    SpreadsheetApp.flush();
    const message = header + 'のボタンを ' + buttons + ' 行に設定しました。' +
      (migrated ? '既存の履歴 ' + migrated + ' 件を「' + CRED_OPTIONS.logSheet + '」へ移しました。' : '') +
      (skipped.length ? '（ボタン列の無い「' + skipped.join('」「') + '」には付けていません）' : '');
    ss.toast(message, CRED_OPTIONS.dialogTitle, 10);
    return message;
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function credInstallTrigger_(ss) {
  const existing = ScriptApp.getProjectTriggers().filter(t =>
    t.getHandlerFunction() === CRED_OPTIONS.triggerHandler &&
    t.getTriggerSourceId() === ss.getId());

  existing.slice(1).forEach(t => ScriptApp.deleteTrigger(t));
  if (!existing.length) {
    ScriptApp.newTrigger(CRED_OPTIONS.triggerHandler)
      .forSpreadsheet(ss).onEdit().create();
  }
}

/* ---------------- ダイアログから呼ばれる関数 ---------------- */

function getCredentialDialogData(sheetName, row) {
  return credWithLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return credDialogData_(ss, credTarget_(ss, {sheetName, row}));
  });
}

/** ロック取得中に呼ぶこと。 */
function credDialogData_(ss, target) {
  const {log, cols} = credLogSheet_(ss, true);
  const timezone = ss.getSpreadsheetTimeZone();
  return {
    sheetName: target.sheet.getName(),
    row: target.row,
    customer: target.customer,
    customerColumn: target.customerColumn,
    fields: CRED_OPTIONS.fields,
    linking: CRED_OPTIONS.linking,
    schedule: CRED_OPTIONS.schedule,
    entries: credReadEntries_(log, cols, target.customer),
    columnWidths: credColumnWidths_(),
    currentUser: credActiveEmail_(),
    today: Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd')
  };
}

/** ダイアログで調整した列幅を、操作した人ごとに保存する。 */
function saveCredentialColumnWidths(widths) {
  const clean = {};
  Object.keys(widths || {}).slice(0, 40).forEach(key => {
    const width = Math.round(Number(widths[key]));
    if (/^[A-Za-z0-9_-]{1,40}$/.test(key) && width >= 40 && width <= 2000) clean[key] = width;
  });
  PropertiesService.getUserProperties()
    .setProperty(CRED_OPTIONS.columnWidthsKey, JSON.stringify(clean));
  return true;
}

/** ダイアログが測った「ブラウザ幅の95%」を、利用者ごとに覚えておく。 */
function saveCredentialDialogSize(width, height) {
  const email = credActiveEmail_();
  width = Math.round(Number(width));
  height = Math.round(Number(height));
  if (!email || email === '取得不可') return false;
  if (!(width >= 600 && width <= 4000 && height >= 400 && height <= 3000)) return false;
  PropertiesService.getDocumentProperties().setProperty(
    CRED_OPTIONS.dialogSizePrefix + email, JSON.stringify({width, height}));
  return true;
}

function credDialogSize_(email) {
  const fallback = {width: CRED_OPTIONS.dialogWidth, height: CRED_OPTIONS.dialogHeight};
  if (!email) return fallback;
  try {
    const saved = JSON.parse(PropertiesService.getDocumentProperties()
      .getProperty(CRED_OPTIONS.dialogSizePrefix + email) || 'null');
    return saved && saved.width && saved.height ? saved : fallback;
  } catch (_) {
    return fallback;
  }
}

function credColumnWidths_() {
  try {
    return JSON.parse(
      PropertiesService.getUserProperties().getProperty(CRED_OPTIONS.columnWidthsKey) || '{}');
  } catch (_) {
    return {};
  }
}

/** 新規登録。req = {sheetName, row, customer, customerColumn, values} → {entry, message} */
function saveCredentialEntry(req) {
  return credWithLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const target = credTarget_(ss, req);
    const fields = credValidate_((req && req.values) || {}, null);
    const {log, cols} = credLogSheet_(ss, false);
    const ids = req.linkId ? credIds_(log, cols) : [];
    const link = credResolveLink_(log, cols, ids, target.customer, fields, req.linkId, '');

    const now = new Date();
    const email = credActiveEmail_();
    const entry = {
      id: Utilities.getUuid(),
      at: now,
      customer: target.customer,
      fields,
      link: link.id,
      editor: email,
      source: target.sheet.getName() + '!' + target.row + '行'
    };
    const rowNumber = credAppendLog_(log, cols, [entry]);

    const message = credAudit_(ss, now, email, '履歴登録', [{
      cell: credRowA1_(rowNumber, cols),
      columnName: '（行全体）',
      before: '',
      after: credSummary_(target.customer, fields, link.label)
    }]);

    return {entry: credClientEntry_(entry, ss.getSpreadsheetTimeZone()), message};
  });
}

/** 編集。req = {sheetName, row, customer, customerColumn, id, original, values} → {entry, message} */
function updateCredentialEntry(req) {
  return credWithLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = ss.getSpreadsheetTimeZone();
    const {target, log, cols, rowNumber, current, ids} = credExisting_(ss, req);
    const fields = credValidate_((req && req.values) || {}, current.fields);
    const link = credResolveLink_(log, cols, ids, target.customer, fields, req.linkId, current.id);
    const L = CRED_OPTIONS.linking;

    // 紐づくオファリングがあるクレデンシャルは、種別を変えられない
    if (current.fields[L.kindKey] === L.parentKind && fields[L.kindKey] !== L.parentKind &&
        credHasChildren_(log, cols, current.id)) {
      throw new Error('紐づいている' + L.childKind + 'があるため、種別を変更できません。先に' +
        L.childKind + 'の紐づけを外してください。');
    }

    const changes = CRED_OPTIONS.fields
      .filter(f => fields[f.key] !== current.fields[f.key])
      .map(f => ({field: f, before: current.fields[f.key], after: fields[f.key]}));
    const linkChanged = link.id !== current.linkId;

    if (!changes.length && !linkChanged) {
      return {entry: credClientEntry_(current, timezone), message: '変更はありませんでした。'};
    }

    const now = new Date();
    const email = credActiveEmail_();
    changes.forEach(ch => {
      log.getRange(rowNumber, cols['f.' + ch.field.key]).setValue(credText_(ch.after));
    });
    if (linkChanged) log.getRange(rowNumber, cols.link).setValue(credText_(link.id));
    log.getRange(rowNumber, cols.updatedBy).setValue(credText_(email));
    log.getRange(rowNumber, cols.updatedAt).setValue(now).setNumberFormat(DIFF_OPTIONS.dateFormat);

    const records = changes.map(ch => ({
      cell: diffA1_(rowNumber, cols['f.' + ch.field.key]),
      columnName: ch.field.label + '（' + target.customer + '）',
      before: ch.before,
      after: ch.after
    }));
    if (linkChanged) {
      records.push({
        cell: diffA1_(rowNumber, cols.link),
        columnName: L.label + '（' + target.customer + '）',
        before: credLinkLabelById_(log, cols, ids, current.linkId),
        after: link.label || 'なし'
      });
    }
    const message = credAudit_(ss, now, email, '履歴編集', records);

    return {
      entry: credClientEntry_(
        Object.assign({}, current, {fields, link: link.id, updatedBy: email, updatedAt: now}), timezone),
      message
    };
  });
}

/** 削除。req = {sheetName, row, customer, customerColumn, id, original} → {id, message} */
function deleteCredentialEntry(req) {
  return credWithLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const {target, log, cols, rowNumber, current} = credExisting_(ss, req);

    const now = new Date();
    const email = credActiveEmail_();
    const cell = credRowA1_(rowNumber, cols);
    const before = credSummary_(target.customer, current.fields,
      current.linkId ? credLinkLabelById_(log, cols, credIds_(log, cols), current.linkId) : '');

    if (CRED_OPTIONS.deleteMode === 'hard') {
      log.deleteRow(rowNumber);
    } else {
      log.getRange(rowNumber, cols.deletedBy).setValue(credText_(email));
      log.getRange(rowNumber, cols.deletedAt).setValue(now).setNumberFormat(DIFF_OPTIONS.dateFormat);
    }

    const message = credAudit_(ss, now, email, '履歴削除', [{
      cell, columnName: '（行全体）', before, after: '（削除）'
    }]);

    return {id: current.id, message};
  });
}

/* ---------------- 記録シート ---------------- */

/** 記録シートの列構成（見出し名で識別）。 */
function credLogLayout_() {
  const col = (key, label, isDate) => ({key, label, isDate: !!isDate});
  return [col('id', 'ID'), col('at', '登録日時', true), col('customer', '得意先')]
    .concat(CRED_OPTIONS.fields.map(f =>
      ({key: 'f.' + f.key, label: f.label, fieldKey: f.key, isDate: false})))
    .concat([
      col('link', CRED_OPTIONS.linking.header),
      col('editor', '登録者'),
      col('source', '登録元（登録時の行）'),
      col('updatedBy', '更新者'),
      col('updatedAt', '更新日時', true),
      col('deletedBy', '削除者'),
      col('deletedAt', '削除日時', true)
    ]);
}

/**
 * ロック取得中に呼ぶこと。シートと不足列を整えて {log, cols} を返す。
 * fillIds = true のときだけ、以前の版で作られた行に ID を振る（ダイアログを開くとき）。
 */
function credLogSheet_(ss, fillIds) {
  const name = CRED_OPTIONS.logSheet;
  if (DIFF_RULES[name] || name === DIFF_OPTIONS.logSheet ||
      name === DIFF_OPTIONS.rowSnapshotSheet || name.startsWith(DIFF_OPTIONS.snapshotPrefix)) {
    throw new Error('記録シート名が差分追跡のシート名と重複しています: ' + name);
  }

  let log = ss.getSheetByName(name);
  if (!log) {
    log = ss.insertSheet(name);
    log.setFrozenRows(1);
  }

  const cols = credEnsureLogColumns_(log, credLogLayout_());
  if (fillIds) credFillIds_(log, cols);
  return {log, cols};
}

function credEnsureLogColumns_(log, layout) {
  const labels = layout.map(c => c.label);
  if (new Set(labels).size !== labels.length) {
    throw new Error('入力項目の見出しが固定列の見出しと重複しています: ' + labels.join(' / '));
  }

  const lastColumn = Math.max(log.getLastColumn(), 1);
  const headers = log.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(h => h.trim());
  const byLabel = {};
  headers.forEach((header, i) => {
    if (!header) return;
    if (byLabel[header]) throw new Error('「' + log.getName() + '」に同じ見出しが2つあります: ' + header);
    byLabel[header] = i + 1;
  });

  Object.entries(CRED_OPTIONS.renamedColumns || {}).forEach(([oldLabel, newLabel]) => {
    if (byLabel[oldLabel] && !byLabel[newLabel] && labels.indexOf(newLabel) >= 0) {
      log.getRange(1, byLabel[oldLabel]).setValue(newLabel);
      byLabel[newLabel] = byLabel[oldLabel];
      delete byLabel[oldLabel];
    }
  });

  const missing = layout.filter(c => !byLabel[c.label]);
  if (missing.length) {
    const start = headers.some(h => h) ? lastColumn + 1 : 1;
    diffEnsureGrid_(log, 1, start + missing.length - 1);
    log.getRange(1, start, 1, missing.length)
      .setValues([missing.map(c => c.label)])
      .setFontWeight('bold')
      .setBackground('#e8eef5');
    missing.forEach((c, i) => {
      byLabel[c.label] = start + i;
      if (c.isDate && log.getMaxRows() > 1) {
        log.getRange(2, start + i, log.getMaxRows() - 1, 1).setNumberFormat(DIFF_OPTIONS.dateFormat);
      }
    });
  }

  const cols = {};
  layout.forEach(c => { cols[c.key] = byLabel[c.label]; });

  // 以前の版の列（例：宛先）は作らず、あれば読み取りにだけ使う
  Object.entries(CRED_OPTIONS.legacyColumns || {}).forEach(([label, fieldKey]) => {
    if (byLabel[label] && labels.indexOf(label) < 0) cols['legacy.' + fieldKey] = byLabel[label];
  });
  return cols;
}

/** 以前の版で作られた行に ID を振る。 */
function credFillIds_(log, cols) {
  const last = log.getLastRow();
  if (last < 2) return;

  const idRange = log.getRange(2, cols.id, last - 1, 1);
  const ids = idRange.getDisplayValues();
  const customers = log.getRange(2, cols.customer, last - 1, 1).getDisplayValues();
  let changed = false;

  ids.forEach((row, i) => {
    if (!row[0].trim() && customers[i][0].trim()) {
      row[0] = Utilities.getUuid();
      changed = true;
    }
  });
  if (changed) idRange.setValues(ids);
}

/**
 * entries = [{id, at, customer, fields:{}, editor, source}]。先頭行の行番号を返す。
 * 1件なら appendRow の1回で書き込む（登録を速くするため）。
 */
function credAppendLog_(log, cols, entries) {
  const layout = credLogLayout_();
  const width = credWidth_(cols);
  const rows = entries.map(entry => {
    const out = new Array(width).fill('');
    layout.forEach(c => {
      const value = credEntryValue_(entry, c);
      out[cols[c.key] - 1] = c.isDate ? (value || '') : credText_(value);
    });
    return out;
  });

  let start;
  if (rows.length === 1) {
    log.appendRow(rows[0]);
    start = log.getLastRow();
  } else {
    start = log.getLastRow() + 1;
    diffEnsureGrid_(log, start + rows.length - 1, width);
    log.getRange(start, 1, rows.length, width).setValues(rows);
  }
  log.getRange(start, cols.at, rows.length, 1).setNumberFormat(DIFF_OPTIONS.dateFormat);
  return start;
}

/** 先頭に ' を付けて文字列として書き込む（'=' で始まっても数式・日付・数値にならない）。 */
function credText_(value) {
  const text = String(value == null ? '' : value);
  return text ? "'" + text : '';
}

function credWidth_(cols) {
  return Math.max.apply(null, Object.values(cols));
}

function credEntryValue_(entry, column) {
  const value = column.fieldKey ? (entry.fields || {})[column.fieldKey] : entry[column.key];
  return value == null ? '' : value;
}

function credRowToEntry_(row, cols) {
  const get = key => String(row[cols[key] - 1] == null ? '' : row[cols[key] - 1]);
  const fields = {};
  CRED_OPTIONS.fields.forEach(f => { fields[f.key] = get('f.' + f.key); });
  Object.keys(cols).filter(k => k.indexOf('legacy.') === 0).forEach(k => {
    const fieldKey = k.slice('legacy.'.length);
    if (!fields[fieldKey]) fields[fieldKey] = get(k);
  });
  return {
    id: get('id'), at: get('at'), customer: get('customer'), fields,
    linkId: get('link'), link: get('link'),
    editor: get('editor'), source: get('source'),
    updatedBy: get('updatedBy'), updatedAt: get('updatedAt'),
    deleted: !!(get('deletedAt') || get('deletedBy'))
  };
}

/** 得意先が合致する、削除されていない記録を新しい順で返す。 */
function credReadEntries_(log, cols, customer) {
  const last = log.getLastRow();
  if (last < 2) return [];

  const key = credNormalize_(customer);
  return log.getRange(2, 1, last - 1, credWidth_(cols))
    .getDisplayValues()
    .map(row => credRowToEntry_(row, cols))
    .filter(e => !e.deleted && e.id && e.customer && credNormalize_(e.customer) === key)
    .reverse()
    .map(e => ({
      id: e.id, at: e.at, values: e.fields, linkId: e.linkId, editor: e.editor, source: e.source,
      updatedBy: e.updatedBy, updatedAt: e.updatedAt
    }));
}

function credRowA1_(rowNumber, cols) {
  return diffA1_(rowNumber, 1) + ':' + diffA1_(rowNumber, credWidth_(cols));
}

/** ダイアログへ返す形。Date は文字列に整形する。 */
function credClientEntry_(entry, timezone) {
  const format = v => v instanceof Date
    ? Utilities.formatDate(v, timezone, DIFF_OPTIONS.dateFormat)
    : String(v == null ? '' : v);
  return {
    id: entry.id,
    at: format(entry.at),
    values: entry.fields,
    linkId: entry.link || '',
    editor: entry.editor || '',
    source: entry.source || '',
    updatedBy: entry.updatedBy || '',
    updatedAt: format(entry.updatedAt)
  };
}

/* ---------------- 共通処理 ---------------- */

function credWithLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CRED_OPTIONS.lockWaitMs)) {
    throw new Error('他の処理が実行中です。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } catch (error) {
    console.error(error.stack || String(error));
    throw error;
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

/**
 * ダイアログの対象行を解決し、得意先が開いたときと同じか確認する。
 * customerColumn が分かっていれば見出しの読み込みを省く（列がずれていれば値が合わず止まる）。
 */
function credTarget_(ss, req) {
  const sheet = ss.getSheetByName(String(req && req.sheetName));
  if (!sheet) throw new Error('シートが見つかりません: ' + (req && req.sheetName));

  const rule = DIFF_RULES[sheet.getName()];
  const row = Number(req.row);
  const column = Number(req.customerColumn);
  let target;

  if (rule && req.customer != null && Number.isInteger(column) && column > 0 &&
      Number.isInteger(row) && row > rule.headerRow) {
    const customer = sheet.getRange(row, column).getDisplayValue().trim();
    target = {sheet, rule, row, customerColumn: column, customer};
  } else {
    target = credResolveTarget_(sheet, row);
  }

  if (req.customer != null && credNormalize_(target.customer) !== credNormalize_(req.customer)) {
    throw new Error(
      target.row + '行目の得意先が変わっています（並べ替え・列の追加など）。ダイアログを閉じて開き直してください。'
    );
  }
  return target;
}

/** 編集・削除の対象を探し、ダイアログ表示時から変わっていないか確認する。 */
function credExisting_(ss, req) {
  const target = credTarget_(ss, req);
  const {log, cols} = credLogSheet_(ss, false);

  const id = String((req && req.id) || '');
  const ids = credIds_(log, cols);
  const index = id ? ids.indexOf(id) : -1;
  if (index < 0) throw new Error('対象の履歴が見つかりません。ダイアログを開き直してください。');

  const rowNumber = index + 2;
  const current = credRowToEntry_(
    log.getRange(rowNumber, 1, 1, credWidth_(cols)).getDisplayValues()[0], cols);

  if (current.deleted) throw new Error('この履歴はすでに削除されています。ダイアログを開き直してください。');
  if (credNormalize_(current.customer) !== credNormalize_(target.customer)) {
    throw new Error('この履歴は別の得意先のものです。ダイアログを開き直してください。');
  }

  const original = (req && req.original) || {};
  const changedByOthers = CRED_OPTIONS.fields.some(f =>
    String(original[f.key] == null ? '' : original[f.key]) !== current.fields[f.key]) ||
    (req.originalLinkId != null && String(req.originalLinkId) !== current.linkId);
  if (changedByOthers) {
    throw new Error('この履歴は他の人が先に変更しています。ダイアログを開き直して最新の内容を確認してください。');
  }

  return {target, log, cols, rowNumber, current, ids};
}

/* ---------------- 紐づけ ---------------- */

function credIds_(log, cols) {
  const last = log.getLastRow();
  return last < 2 ? [] : log.getRange(2, cols.id, last - 1, 1).getDisplayValues().map(r => r[0]);
}

function credReadRow_(log, cols, rowNumber) {
  return credRowToEntry_(
    log.getRange(rowNumber, 1, 1, credWidth_(cols)).getDisplayValues()[0], cols);
}

/**
 * オファリングの紐づけ先を検証して {id, label} を返す。
 * オファリング以外、または未指定なら紐づけなし。
 */
function credResolveLink_(log, cols, ids, customer, fields, linkId, selfId) {
  const L = CRED_OPTIONS.linking;
  linkId = String(linkId || '').trim();
  if (!linkId || fields[L.kindKey] !== L.childKind) return {id: '', label: ''};
  if (linkId === selfId) throw new Error('自分自身には紐づけられません。');

  const index = ids.indexOf(linkId);
  if (index < 0) throw new Error('紐づけ先の' + L.parentKind + 'が見つかりません。ダイアログを開き直してください。');

  const parent = credReadRow_(log, cols, index + 2);
  if (parent.deleted) throw new Error('紐づけ先の' + L.parentKind + 'は削除されています。');
  if (credNormalize_(parent.customer) !== credNormalize_(customer)) {
    throw new Error('紐づけ先は別の得意先の履歴です。');
  }
  if (parent.fields[L.kindKey] !== L.parentKind) {
    throw new Error('紐づけ先の種別が「' + L.parentKind + '」ではありません。');
  }
  return {id: linkId, label: credLinkText_(parent)};
}

function credLinkText_(entry) {
  const L = CRED_OPTIONS.linking;
  return L.labelKeys
    .map(key => key.split('|').map(k => entry.fields[k]).find(Boolean))
    .filter(Boolean).join(' / ');
}

function credLinkLabelById_(log, cols, ids, linkId) {
  if (!linkId) return 'なし';
  const index = ids.indexOf(linkId);
  if (index < 0) return '（見つかりません）';
  const parent = credReadRow_(log, cols, index + 2);
  return credLinkText_(parent) + (parent.deleted ? '（削除済み）' : '');
}

function credHasChildren_(log, cols, id) {
  const last = log.getLastRow();
  if (last < 2) return false;
  return log.getRange(2, 1, last - 1, credWidth_(cols)).getDisplayValues()
    .map(row => credRowToEntry_(row, cols))
    .some(e => !e.deleted && e.linkId === id);
}

/** 変更履歴_差分 に記録する。失敗しても本体の操作は取り消さず、警告文を返す。 */
function credAudit_(ss, now, email, kind, items) {
  if (!CRED_OPTIONS.auditToDiffLog || !items.length) return '';
  try {
    const eventId = diffCreateEventId_(now, ss.getSpreadsheetTimeZone(), CRED_OPTIONS.auditEventPrefix);
    const budget = {remaining: DIFF_OPTIONS.maxLcsCellsPerEdit};
    diffAppend_(ss, items.map(item => ({
      eventId,
      editorEmail: email,
      sheet: CRED_OPTIONS.logSheet,
      cell: item.cell,
      columnName: item.columnName,
      kind,
      before: item.before,
      after: item.after,
      parts: diffCharacters_(item.before, item.after, budget).parts
    })), now);
    return '';
  } catch (error) {
    console.error(error.stack || String(error));
    return '保存しましたが、変更履歴_差分への記録に失敗しました。管理者に実行ログの確認を依頼してください。';
  }
}

function credSummary_(customer, fields, linkLabel) {
  const L = CRED_OPTIONS.linking;
  const lines = ['得意先: ' + customer]
    .concat(CRED_OPTIONS.fields.map(f => f.label + ': ' + (fields[f.key] || '')));
  if (fields[L.kindKey] === L.childKind) lines.push(L.label + ': ' + (linkLabel || 'なし'));
  return lines.join('\n');
}

function credResolveTarget_(sheet, row) {
  const rule = DIFF_RULES[sheet.getName()];
  if (!rule) throw new Error('このシートは対象外です: ' + sheet.getName());
  if (!Number.isInteger(row) || row <= rule.headerRow || row > sheet.getMaxRows()) {
    throw new Error('データ行（' + (rule.headerRow + 1) + '行目以降）を選んでください。');
  }

  const customerColumn = credColumnByHeader_(sheet, rule, CRED_OPTIONS.customerHeader);
  const customer = sheet.getRange(row, customerColumn).getDisplayValue().trim();
  if (!customer) {
    throw new Error(row + '行目の「' + CRED_OPTIONS.customerHeader + '」が空欄です。');
  }
  return {sheet, rule, row, customerColumn, customer};
}

/** 見出しの列番号。ちょうど1つでなければ止める（optional なら、見出しが無いときは 0 を返す。ボタン列はシートに置かなくてもよいため）。 */
function credColumnByHeader_(sheet, rule, header, optional) {
  const normalize = v => String(v == null ? '' : v).replace(/\s/g, '');
  const target = normalize(header);
  const headers = sheet
    .getRange(rule.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getDisplayValues()[0];

  const matches = [];
  headers.forEach((h, i) => { if (normalize(h) === target) matches.push(i + 1); });

  if (optional && !matches.length) return 0;
  if (matches.length !== 1) {
    throw new Error(
      sheet.getName() + ' の ' + rule.headerRow + '行目に「' + credHeaderText_(header) +
      '」見出しがちょうど1つ必要です（見つかった数: ' + matches.length + '）。'
    );
  }
  return matches[0];
}

/** メッセージ表示用：見出しの改行を空白にする。 */
function credHeaderText_(header) {
  return String(header).replace(/\s+/g, ' ').trim();
}

/** 全角・半角と空白の違いを無視して照合する。 */
function credNormalize_(value) {
  return String(value == null ? '' : value).normalize('NFKC').replace(/\s+/g, '');
}

/**
 * 入力を検証して {key: value} を返す。
 * current を渡すと、選択肢から外れた旧い値（例：以前の種別）でも、変更しなければ通す。
 */
function credValidate_(input, current) {
  const out = {};
  CRED_OPTIONS.fields.forEach(field => {
    let value = input[field.key] == null ? '' : String(input[field.key]);
    value = field.type === 'textarea'
      ? value.replace(/\r\n?/g, '\n').trim()
      : value.replace(/[\r\n\t]+/g, ' ').trim();

    if (field.required && !value) throw new Error('「' + field.label + '」は必須です。');
    if (field.maxLength && value.length > field.maxLength) {
      throw new Error('「' + field.label + '」は' + field.maxLength + '文字以内で入力してください。');
    }
    if (field.type === 'date' && value) value = credNormalizeDate_(value, field.label);
    if (field.type === 'select' && value && field.options.indexOf(value) < 0 &&
        !(current && current[field.key] === value)) {
      throw new Error('「' + field.label + '」の選択肢が正しくありません。');
    }
    out[field.key] = value;
  });

  const S = CRED_OPTIONS.schedule;
  if (S && !out[S.planKey] && !out[S.doneKey]) {
    const label = key => (CRED_OPTIONS.fields.find(f => f.key === key) || {}).label || key;
    throw new Error('「' + label(S.planKey) + '」か「' + label(S.doneKey) + '」のどちらかを入力してください。');
  }
  return out;
}

/** 2026-10-01 / 2026/10/01 / 2026/10/1 などを受け付け、yyyy/MM/dd にそろえる。 */
function credNormalizeDate_(value, label) {
  const m = String(value).trim().match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  const y = m && Number(m[1]), mo = m && Number(m[2]), d = m && Number(m[3]);
  const date = m && new Date(y, mo - 1, d);
  if (!m || date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) {
    throw new Error('「' + label + '」の日付が正しくありません（例：2026/10/01）。');
  }
  const pad = n => String(n).padStart(2, '0');
  return y + '/' + pad(mo) + '/' + pad(d);
}

/** 初版の「項目 | … | 登録者 | 登録日時」形式を読み取る。合わない行は備考へ入れる。 */
function credParseLegacy_(text, timezone) {
  const keys = CRED_OPTIONS.legacyFieldKeys;
  const noteField = CRED_OPTIONS.fields.find(f => f.type === 'textarea') ||
    CRED_OPTIONS.fields[CRED_OPTIONS.fields.length - 1];

  return text.split(/\r?\n/).filter(line => line.trim()).map(line => {
    const parts = line.split(CRED_OPTIONS.legacySeparator);
    const fields = {};
    if (parts.length === keys.length + 2) {
      keys.forEach((key, i) => { fields[key] = parts[i]; });
      let at;
      try { at = Utilities.parseDate(parts[keys.length + 1].trim(), timezone, DIFF_OPTIONS.dateFormat); }
      catch (_) { at = new Date(); }
      return {at, fields, editor: parts[keys.length]};
    }
    fields[noteField.key] = line;
    return {at: new Date(), fields, editor: '不明（移行）'};
  });
}

function credActiveEmail_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) return email;
  } catch (_) {}
  return '取得不可';
}