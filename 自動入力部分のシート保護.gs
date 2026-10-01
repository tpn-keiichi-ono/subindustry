function protectSpecificColumnsWithLock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 対象列：P=16, R=18, T=20, X=24, AA=27, AB=28
  const targetCols = [16, 18, 20, 24, 27, 28];
  const prefix = "案件_";

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.indexOf(prefix) !== 0) return;

    // 既存の保護設定を安全に取得
    let existingProtections = [];
    try {
      existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    } catch (e) {
      console.warn(`${sheetName} の保護情報取得中にエラーが発生しました。スキップします。`);
      return;
    }
    
    targetCols.forEach(col => {
      // 説明文も「警告」から「編集不可」に変更
      const description = `【編集不可】${sheetName}_${col}列目`;
      
      // すでに同じ説明の保護があるかチェック（エラー個体はスキップ）
      let alreadyExists = false;
      try {
        alreadyExists = existingProtections.some(p => {
          try {
            return p.getDescription() === description;
          } catch (e) {
            return false; // 削除済みの保護だった場合は無視
          }
        });
      } catch (e) {
        alreadyExists = false;
      }
      
      if (!alreadyExists) {
        try {
          const range = sheet.getRange(1, col, sheet.getMaxRows(), 1);
          const protection = range.protect().setDescription(description);
          
          // --- ここから修正部分 ---
          // 1. 自分（スクリプト実行者）を編集者として明示的に追加
          const me = Session.getEffectiveUser();
          protection.addEditor(me);
          
          // 2. 自分以外の編集者をすべて削除
          protection.removeEditors(protection.getEditors());
          
          // 3. Google Workspace環境の場合、ドメイン全体の編集権限もオフにする
          if (protection.canDomainEdit()) {
            protection.setDomainEdit(false);
          }
          // --- ここまで修正部分 ---
          
          console.log(`${sheetName} の ${getColumnLetter(col)}列 を編集不可にしました。`);
        } catch (e) {
          console.error(`${sheetName} の ${col}列目の保護作成に失敗しました: ${e.message}`);
        }
      }
    });
  });
  
  Browser.msgBox("処理が完了しました。");
}

function getColumnLetter(col) {
  let letter = "";
  while (col > 0) {
    let t = (col - 1) % 26;
    letter = String.fromCharCode(65 + t) + letter;
    col = (col - t - 1) / 26;
  }
  return letter;
}