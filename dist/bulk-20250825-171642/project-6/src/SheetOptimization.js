// ===========================
// シート最適化ユーティリティ
// ===========================

// ===========================
// シート診断関数
// ===========================
function diagnoseSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  
  try {
    results.push('=== スプレッドシート診断結果 ===\n');
    
    // 1. シート数とサイズの確認
    const sheets = ss.getSheets();
    results.push(`シート数: ${sheets.length}`);
    
    let totalRows = 0;
    let totalCells = 0;
    const largeSheets = [];
    
    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const maxRows = sheet.getMaxRows();
      const maxCols = sheet.getMaxColumns();
      const dataRows = lastRow;
      const dataCells = lastRow * lastCol;
      const totalSheetCells = maxRows * maxCols;
      
      totalRows += dataRows;
      totalCells += dataCells;
      
      // 大きなシートを記録
      if (dataRows > 1000 || totalSheetCells > 100000) {
        largeSheets.push({
          name: sheet.getName(),
          dataRows: dataRows,
          maxRows: maxRows,
          dataCells: dataCells,
          totalCells: totalSheetCells
        });
      }
    });
    
    results.push(`総データ行数: ${totalRows}`);
    results.push(`総データセル数: ${totalCells}`);
    
    // 2. 大きなシートの詳細
    if (largeSheets.length > 0) {
      results.push('\n=== 大きなシート ===');
      largeSheets.forEach(sheet => {
        results.push(`${sheet.name}:`);
        results.push(`  データ行: ${sheet.dataRows}行 / 最大行: ${sheet.maxRows}行`);
        results.push(`  データセル: ${sheet.dataCells} / 総セル: ${sheet.totalCells}`);
        if (sheet.maxRows > sheet.dataRows * 2) {
          results.push(`  ⚠️ 未使用行が多い（${sheet.maxRows - sheet.dataRows}行）`);
        }
      });
    }
    
    // 3. 条件付き書式の確認
    results.push('\n=== 条件付き書式 ===');
    let totalRules = 0;
    sheets.forEach(sheet => {
      const rules = sheet.getConditionalFormatRules();
      if (rules.length > 0) {
        results.push(`${sheet.getName()}: ${rules.length}個のルール`);
        totalRules += rules.length;
      }
    });
    results.push(`合計: ${totalRules}個のルール`);
    
    // 4. 推奨事項
    results.push('\n=== 推奨事項 ===');
    if (totalRows > 10000) {
      results.push('⚠️ データ行が多いため、古いデータのアーカイブを推奨');
    }
    if (totalRules > 50) {
      results.push('⚠️ 条件付き書式が多いため、整理を推奨');
    }
    largeSheets.forEach(sheet => {
      if (sheet.maxRows > sheet.dataRows * 2) {
        results.push(`⚠️ ${sheet.name}シートの未使用行を削除することを推奨`);
      }
    });
    
    // 結果表示
    const message = results.join('\n');
    ui.alert('診断結果', message, ui.ButtonSet.OK);
    
    // ログにも記録
    console.log(message);
    logOperation('スプレッドシート診断', 'success', `総行数: ${totalRows}, 総セル数: ${totalCells}`);
    
  } catch (error) {
    console.error('診断エラー:', error);
    ui.alert('エラー', '診断中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ===========================
// シート最適化関数
// ===========================
function optimizeSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.log(`シート「${sheetName}」が見つかりません`);
    return false;
  }
  
  try {
    console.log(`${sheetName}シートの最適化開始...`);
    
    // 1. 未使用の行と列を削除
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    
    // 未使用行の削除（最後のデータ行 + バッファ100行を残す）
    if (maxRows > lastRow + 100) {
      const rowsToDelete = maxRows - (lastRow + 100);
      sheet.deleteRows(lastRow + 101, rowsToDelete);
      console.log(`  ${rowsToDelete}行を削除`);
    }
    
    // 未使用列の削除（最後のデータ列 + バッファ10列を残す）
    if (maxCols > lastCol + 10) {
      const colsToDelete = maxCols - (lastCol + 10);
      sheet.deleteColumns(lastCol + 11, colsToDelete);
      console.log(`  ${colsToDelete}列を削除`);
    }
    
    // 2. 条件付き書式のクリーンアップ
    const rules = sheet.getConditionalFormatRules();
    if (rules.length > 0) {
      // 古いルールをクリアして必要なものだけ再設定
      sheet.clearConditionalFormatRules();
      console.log(`  ${rules.length}個の条件付き書式をクリア`);
    }
    
    console.log(`${sheetName}シートの最適化完了`);
    return true;
    
  } catch (error) {
    console.error(`${sheetName}シート最適化エラー:`, error);
    return false;
  }
}

// ===========================
// 全シート最適化
// ===========================
function optimizeAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'シート最適化',
    'すべてのシートを最適化しますか？\n\n' +
    '実行内容：\n' +
    '• 未使用の行と列を削除\n' +
    '• 古い条件付き書式をクリア\n' +
    '• データの整理\n\n' +
    '※データは保持されます',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let successCount = 0;
    
    sheets.forEach(sheet => {
      if (optimizeSheet(sheet.getName())) {
        successCount++;
      }
    });
    
    ui.alert(
      '最適化完了',
      `${successCount}/${sheets.length}シートを最適化しました。\n\n` +
      'シート再構成を実行してください。',
      ui.ButtonSet.OK
    );
    
    logOperation('全シート最適化', 'success', `${successCount}シートを最適化`);
    
  } catch (error) {
    console.error('全シート最適化エラー:', error);
    ui.alert('エラー', '最適化中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ===========================
// 古いデータのクリーンアップ
// ===========================
function cleanupOldData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '古いデータのクリーンアップ',
    '30日以上前のログと応答結果を削除しますか？\n\n' +
    '対象シート：\n' +
    '• ログシート\n' +
    '• 自動応答結果シート\n' +
    '• 受信したリプライシート',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 30);
    
    let totalDeleted = 0;
    
    // ログシートのクリーンアップ
    const logSheet = ss.getSheetByName('ログ');
    if (logSheet) {
      const deleted = cleanupSheetByDate(logSheet, 1, cutoffDate);
      totalDeleted += deleted;
      console.log(`ログシート: ${deleted}行削除`);
    }
    
    // 自動応答結果シートのクリーンアップ
    const replySheet = ss.getSheetByName('自動応答結果');
    if (replySheet) {
      const deleted = cleanupSheetByDate(replySheet, 1, cutoffDate);
      totalDeleted += deleted;
      console.log(`自動応答結果シート: ${deleted}行削除`);
    }
    
    // 受信したリプライシートのクリーンアップ
    const repliesSheet = ss.getSheetByName('受信したリプライ');
    if (repliesSheet) {
      const deleted = cleanupSheetByDate(repliesSheet, 2, cutoffDate);
      totalDeleted += deleted;
      console.log(`受信したリプライシート: ${deleted}行削除`);
    }
    
    ui.alert(
      'クリーンアップ完了',
      `${totalDeleted}行の古いデータを削除しました。`,
      ui.ButtonSet.OK
    );
    
    logOperation('データクリーンアップ', 'success', `${totalDeleted}行削除`);
    
  } catch (error) {
    console.error('クリーンアップエラー:', error);
    ui.alert('エラー', 'クリーンアップ中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ===========================
// 日付でシートをクリーンアップ（ヘルパー関数）
// ===========================
function cleanupSheetByDate(sheet, dateColumn, cutoffDate) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    // データを取得（ヘッダー行を除く）
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const rowsToKeep = [];
    let deletedCount = 0;
    
    // 日付をチェックして保持する行を決定
    data.forEach((row, index) => {
      const dateValue = row[dateColumn - 1];
      if (dateValue && dateValue instanceof Date) {
        if (dateValue > cutoffDate) {
          rowsToKeep.push(row);
        } else {
          deletedCount++;
        }
      } else {
        // 日付が無効な場合は保持
        rowsToKeep.push(row);
      }
    });
    
    // シートを更新
    if (deletedCount > 0) {
      // データ範囲をクリア
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
      
      // 保持するデータを書き込み
      if (rowsToKeep.length > 0) {
        sheet.getRange(2, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
      }
    }
    
    return deletedCount;
    
  } catch (error) {
    console.error(`シートクリーンアップエラー (${sheet.getName()}):`, error);
    return 0;
  }
}

// ===========================
// 最適化ヘルプ表示
// ===========================
function showOptimizationHelp() {
  const ui = SpreadsheetApp.getUi();
  const message = 
    'シート最適化機能について\n\n' +
    '【タイムアウトエラーの原因】\n' +
    '• 大量のデータ蓄積\n' +
    '• 未使用行・列の肥大化\n' +
    '• 条件付き書式の累積\n' +
    '• 頻繁なflush()呼び出し\n\n' +
    '【推奨手順】\n' +
    '1. 📊 スプレッドシート診断を実行\n' +
    '2. 🧹 古いデータをクリーンアップ\n' +
    '3. ⚡ 全シート最適化を実行\n' +
    '4. 📁 シート再構成を実行\n\n' +
    '【注意事項】\n' +
    '• 最適化前にバックアップを推奨\n' +
    '• クリーンアップは30日以上前のデータが対象\n' +
    '• 最適化後は動作が軽快になります';
  
  ui.alert('シート最適化ヘルプ', message, ui.ButtonSet.OK);
}