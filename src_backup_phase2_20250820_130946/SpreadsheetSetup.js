// ===========================
// スプレッドシート初期設定
// ===========================

/**
 * 予約投稿シートのヘッダーに画像URL使用に関する注意事項を追加
 */
function addImageUrlNotes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  if (!sheet) {
    console.error('予約投稿シートが見つかりません');
    return;
  }
  
  // ヘッダー行を取得
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  
  // 画像URL列を探す
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes('画像URL')) {
      // 列にメモを追加
      const noteText = '画像URLの使用方法:\n\n' +
                      '• Google Drive（※共有設定を「リンクを知っている全員」に変更）\n' +
                      '• Imgur、Cloudinary等の画像ホスティングサービス\n' +
                      '• 直接アクセス可能な公開URL\n' +
                      '• CDNサービスのURL\n\n' +
                      '⚠️ Google Driveの注意点:\n' +
                      '1. ファイルを右クリック→「共有」\n' +
                      '2. 「リンクを知っている全員」に変更\n' +
                      '3. リンクをコピーして使用\n\n' +
                      '例: https://drive.google.com/file/d/xxxxx/view';
      
      sheet.getRange(1, i + 1).setNote(noteText);
    }
  }
  
  // 動画URL列も同様に処理
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === '動画URL') {
      const noteText = '動画URLの使用方法:\n\n' +
                      '• Google Drive\n' +
                      '• 直接アクセス可能な公開URL\n' +
                      '• CDNサービスのURL';
      
      sheet.getRange(1, i + 1).setNote(noteText);
    }
  }
  
  console.log('画像URL使用に関する注意事項を追加しました');
}

/**
 * スプレッドシートの初期設定を実行
 */
function initializeSpreadsheet() {
  try {
    addImageUrlNotes();
    SpreadsheetApp.getActiveSpreadsheet().toast('初期設定が完了しました', '成功', 3);
  } catch (error) {
    console.error('初期設定エラー:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('初期設定中にエラーが発生しました', 'エラー', 5);
  }
}

/**
 * 予約投稿シートのURLをチェックしてGoogle DriveのURLがある場合に警告
 */
function checkScheduledPostUrls() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', '予約投稿シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 画像URL列のインデックスを取得
  const imageUrlIndices = [];
  const videoUrlIndex = headers.indexOf('動画URL');
  
  headers.forEach((header, index) => {
    if (header.includes('画像URL')) {
      imageUrlIndices.push(index);
    }
  });
  
  let emptyUrls = [];
  let validUrls = 0;
  
  // データを確認
  for (let row = 1; row < data.length; row++) {
    const rowData = data[row];
    const status = rowData[headers.indexOf('ステータス')];
    
    // pending状態の投稿のみチェック
    if (status === 'pending' || status === '投稿予約中') {
      // 画像URLチェック
      let hasMedia = false;
      imageUrlIndices.forEach(colIndex => {
        const url = rowData[colIndex];
        if (url && typeof url === 'string' && url.trim() !== '') {
          hasMedia = true;
          validUrls++;
        }
      });
      
      // 動画URLチェック
      if (videoUrlIndex !== -1) {
        const url = rowData[videoUrlIndex];
        if (url && typeof url === 'string' && url.trim() !== '') {
          hasMedia = true;
          validUrls++;
        }
      }
      
      // メディアがない投稿を記録
      if (!hasMedia) {
        const content = rowData[headers.indexOf('投稿内容')];
        if (content && content.trim() !== '') {
          emptyUrls.push({
            row: row + 1,
            content: content.substring(0, 50) + '...'
          });
        }
      }
    }
  }
  
  // 結果を表示
  const ui = SpreadsheetApp.getUi();
  let message = `URL確認結果:\n\n`;
  message += `✅ 有効なメディアURL: ${validUrls}件\n`;
  
  if (emptyUrls.length > 0) {
    message += `\n⚠️ メディアなしの投稿: ${emptyUrls.length}件\n`;
    emptyUrls.forEach(item => {
      message += `  行${item.row}: ${item.content}\n`;
    });
  }
  
  ui.alert('URLチェック結果', message, ui.ButtonSet.OK);
}