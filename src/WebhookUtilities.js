// WebhookUtilities.js - Webhook関連ユーティリティ関数

// ===========================
// 現在のWeb App URLを取得
// ===========================
function getCurrentWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  console.log('現在のWeb App URL:', url);
  return url;
}

// ===========================
// Webhook URL更新手順を表示
// ===========================
function showWebhookUpdateInstructions() {
  const ui = SpreadsheetApp.getUi();
  const currentUrl = getCurrentWebAppUrl();
  
  const message = 
    '古いWebhook URLが使用されているため、エラーが発生しています。\n\n' +
    '【解決方法】\n' +
    '1. Meta for Developersにログイン\n' +
    '2. アプリのWebhook設定ページを開く\n' +
    '3. Callback URLを以下に更新:\n\n' +
    currentUrl + '\n\n' +
    '4. 保存して、Webhook検証が成功することを確認\n\n' +
    '※ このURLをコピーして使用してください。';
  
  ui.alert('Webhook URL更新が必要', message, ui.ButtonSet.OK);
}

// ===========================
// Webhookログ表示
// ===========================
function showWebhookLogs() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Webhookログ');
  
  if (!sheet) {
    ui.alert(
      'Webhookログなし',
      'Webhookログシートがまだ作成されていません。\n' +
      'Webhookイベントを受信すると自動的に作成されます。',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    ui.alert(
      'Webhookログなし',
      'まだWebhookイベントを受信していません。',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // 最新10件のログを表示
  const recentLogs = data.slice(-11, -1).reverse(); // ヘッダーを除く最新10件
  
  let message = '最新のWebhookイベント（最大10件）:\n\n';
  
  recentLogs.forEach((log, index) => {
    const [date, eventType, id, username, text, keyword, status] = log;
    message += `${index + 1}. ${new Date(date).toLocaleString('ja-JP')}\n`;
    message += `   タイプ: ${eventType}\n`;
    message += `   ユーザー: @${username}\n`;
    message += `   テキスト: ${text ? text.substring(0, 50) + '...' : 'なし'}\n`;
    message += `   ステータス: ${status}\n`;
    if (keyword) {
      message += `   マッチキーワード: ${keyword}\n`;
    }
    message += '\n';
  });
  
  ui.alert('Webhookログ', message, ui.ButtonSet.OK);
}

// ===========================
// Webhookステータス確認
// ===========================
function checkWebhookStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // APP_SECRETの設定状態を確認
    const appSecret = getConfig('APP_SECRET');
    const hasAppSecret = appSecret && appSecret.length > 0;
    
    // デプロイメント状態を確認
    const webhookUrl = getWebhookUrl();
    const isDeployed = webhookUrl.includes('/exec');
    
    // Webhookログの存在確認
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Webhookログ');
    const hasLogs = logSheet && logSheet.getLastRow() > 1;
    
    let status = '📊 Webhook設定ステータス\n\n';
    
    // 設定状態
    status += '1. 基本設定:\n';
    status += `   - APP_SECRET: ${hasAppSecret ? '✅ 設定済み' : '❌ 未設定'}\n`;
    status += `   - Webアプリ: ${isDeployed ? '✅ デプロイ済み' : '⚠️ 未デプロイ'}\n\n`;
    
    // URL情報
    status += '2. Webhook URL:\n';
    status += `   ${webhookUrl}\n\n`;
    
    // 受信状況
    status += '3. 受信状況:\n';
    if (hasLogs) {
      const lastRow = logSheet.getRange(logSheet.getLastRow(), 1, 1, 8).getValues()[0];
      const lastReceived = new Date(lastRow[0]);
      status += `   ✅ 最終受信: ${lastReceived.toLocaleString('ja-JP')}\n`;
      status += `   イベント数: ${logSheet.getLastRow() - 1}件\n`;
    } else {
      status += '   ⚠️ まだイベントを受信していません\n';
    }
    
    // 次のステップ
    status += '\n4. 次のステップ:\n';
    if (!hasAppSecret) {
      status += '   1. 基本設定シートにAPP_SECRETを入力\n';
    }
    if (!isDeployed) {
      status += '   2. 拡張機能 > Apps Script > デプロイ > 新しいデプロイ\n';
      status += '      - 種類: ウェブアプリ\n';
      status += '      - アクセスできるユーザー: 全員\n';
    }
    status += '   3. Meta for DevelopersでWebhookを設定\n';
    
    ui.alert('Webhookステータス', status, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Webhookステータス確認エラー:', error);
    ui.alert('エラー', 'ステータス確認中にエラーが発生しました。', ui.ButtonSet.OK);
  }
}

// ===========================
// Webhook統計情報
// ===========================
function showWebhookStats() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Webhookログ');
  
  if (!sheet || sheet.getLastRow() <= 1) {
    ui.alert('統計情報なし', 'Webhookイベントのデータがありません。', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const logs = data.slice(1); // ヘッダーを除く
  
  // 統計を計算
  const stats = {
    total: logs.length,
    byType: {},
    byStatus: {},
    byKeyword: {},
    recentDays: 7,
    recent: 0
  };
  
  const recentDate = new Date();
  recentDate.setDate(recentDate.getDate() - stats.recentDays);
  
  logs.forEach(log => {
    const [date, eventType, , , , keyword, status] = log;
    
    // イベントタイプ別
    stats.byType[eventType] = (stats.byType[eventType] || 0) + 1;
    
    // ステータス別
    stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
    
    // キーワード別
    if (keyword) {
      stats.byKeyword[keyword] = (stats.byKeyword[keyword] || 0) + 1;
    }
    
    // 最近のイベント
    if (new Date(date) >= recentDate) {
      stats.recent++;
    }
  });
  
  // レポート作成
  let report = '📊 Webhook統計情報\n\n';
  report += `総イベント数: ${stats.total}件\n`;
  report += `過去${stats.recentDays}日間: ${stats.recent}件\n\n`;
  
  report += '📌 イベントタイプ別:\n';
  Object.entries(stats.byType).forEach(([type, count]) => {
    report += `  ${type}: ${count}件\n`;
  });
  
  report += '\n📌 ステータス別:\n';
  Object.entries(stats.byStatus).forEach(([status, count]) => {
    report += `  ${status}: ${count}件\n`;
  });
  
  if (Object.keys(stats.byKeyword).length > 0) {
    report += '\n📌 キーワード別（上位5件）:\n';
    const sortedKeywords = Object.entries(stats.byKeyword)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    
    sortedKeywords.forEach(([keyword, count]) => {
      report += `  ${keyword}: ${count}件\n`;
    });
  }
  
  ui.alert('Webhook統計', report, ui.ButtonSet.OK);
}

// ===========================
// Webhookログクリア
// ===========================
function clearWebhookLogs() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Webhookログクリア',
    'Webhookログをすべて削除しますか？\nこの操作は取り消せません。',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Webhookログ');
  
  if (!sheet) {
    ui.alert('エラー', 'Webhookログシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  // ヘッダー以外を削除
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  ui.alert('完了', 'Webhookログをクリアしました。', ui.ButtonSet.OK);
}