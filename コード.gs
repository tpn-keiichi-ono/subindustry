/**
 * メニューを追加（CredentialHistory.gs の onOpen() から呼ばれる）
 * onOpen() はプロジェクトに1つだけにするため、ここには置かない。
 */
function apAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('アカウントプランシート作成')
    .addItem('作成開始', 'showLoadingDialog')
    .addToUi();
}

/**
 * 1. ローディング画面を表示
 */
function showLoadingDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Loading')
    .setWidth(350)
    .setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '処理を実行中...');
}

/**
 * 2. 実行前に「E列入力済・M列空」の件数を確認する
 */
function checkTargetCount(targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error(`シート「${targetSheetName}」が見つかりません。`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // E列(5)〜M列(13)のデータを取得（計9列分）
  const data = sheet.getRange(2, 5, lastRow - 1, 9).getValues();

  // E列(row[0])が空でなく、かつM列(row[8])が空のものを抽出してカウント
  return data.filter(row => row[0] !== "" && row[8] === "").length;
}

/**
 * 3. 実際のコピー処理（制限時間5分で安全停止）
 */
function executeCopyProcess(targetSheetName) {
  const startTime = Date.now();
  const MAX_RUNTIME_MS = 300000; // 5分で停止

  // --- 設定項目 ---
  const TEMPLATE_FILE_ID = '17Fn7jUW7gdEtX4_Hn5PtOGX4oGLHrqeU-CI6mm49x9k';
  const DEST_FOLDER_ID = '1B4-lxlvtGwroP_5iHi9erMZT7c7sW76P';
  // ----------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  const lastRow = sheet.getLastRow();
  
  // E列(5)からM列(13)までの9列分を取得
  const range = sheet.getRange(2, 5, lastRow - 1, 9);
  const values = range.getValues();
  
  const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
  const destFolder = DriveApp.getFolderById(DEST_FOLDER_ID);

  let count = 0;
  let isTimeout = false;

  for (let i = 0; i < values.length; i++) {
    // 5分経過したら中断
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      isTimeout = true;
      break; 
    }

    const eValue = values[i][0]; // E列 (index 0)
    const mValue = values[i][8]; // M列 (index 8)
    const currentRow = i + 2;    // スプレッドシート上の行番号

    // E列に値があり、かつM列が空の場合のみ処理を実行
    if (eValue !== "" && mValue === "") {
      try {
        const newFileName = "33シナリオ起点_アカウントプラン_" + eValue;
        const newFile = templateFile.makeCopy(newFileName, destFolder);
        
        // 作成したファイルのURLをM列(13列目)に書き込む
        sheet.getRange(currentRow, 13).setValue(newFile.getUrl());
        
        count++;
        SpreadsheetApp.flush(); // 1件ごとにスプレッドシートへ反映
      } catch (e) {
        console.error(`Row ${currentRow}: ${e.message}`);
      }
    }
  }

  if (isTimeout) {
    return `【中断】時間が経過したため${count}件で停止しました。残りは再度実行してください。`;
  } else {
    return `${count} 件の処理が完了しました。`;
  }
}