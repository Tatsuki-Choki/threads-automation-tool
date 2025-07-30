// Threads自動化ツール - メインコード
// Code.gs

// ===========================
// グローバル設定
// ===========================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const THREADS_API_BASE = 'https://graph.threads.net';

// ===========================
// メニューを再読み込み
// ===========================
function refreshMenu() {
  onOpen();
  SpreadsheetApp.getUi().alert('メニューを再読み込みしました', 'Threads自動化メニューが表示されているか確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===========================
// テスト関数（動作確認用）
// ===========================
function testFunction() {
  SpreadsheetApp.getUi().alert('テスト', 'スクリプトは正常に動作しています。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===========================
// 初期設定関数
// ===========================
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 管理者用メニュー
    ui.createMenu('🔒 管理者用')
      .addItem('基本設定を表示', 'showSettingsSheet')
      .addItem('基本設定を非表示', 'hideSettingsSheet')
      .addSeparator()
      .addItem('現在のトリガーの所有者を確認', 'checkTriggerOwners')
      .addItem('API呼び出し回数確認', 'showUrlFetchCountWithAuth')
      .addToUi();
    
    // Threads自動化メニュー
    ui.createMenu('Threads自動化')
    .addItem('🚀 クイックセットアップ', 'quickSetupWithExistingToken')
    .addSeparator()
    .addItem('⏰ トリガーを再設定', 'showRepliesTrackingTriggerDialog')
    .addSeparator()
    .addItem('📤 手動投稿実行', 'manualPostExecution')
    .addItem('💬 リプライ＋自動返信（統合実行）', 'fetchAndAutoReply')
    .addItem('🔄 自動返信のみ', 'manualAutoReply')
    .addSeparator()
    .addItem('🧪 自動返信テスト', 'simulateAutoReply')
    .addSeparator()
    .addSubMenu(ui.createMenu('📁 シート再構成')
      .addItem('💬 受信したリプライシート初期化', 'initializeRepliesSheet')
      .addItem('🔍 キーワード設定シート再構成', 'resetAutoReplyKeywordsSheet')
      .addItem('📅 予約投稿シート再構成', 'resetScheduledPostsSheet')
      .addItem('✅ 自動応答結果シート再構成', 'resetReplyHistorySheet')
      .addItem('⚙️ 基本設定シート再構成', 'resetSettingsSheet')
      .addItem('📝 ログシート再構成', 'resetLogsSheet')
      .addSeparator()
      .addItem('🔄 すべてのシートを再構成', 'resetAllSheets'))
    .addSeparator()
    .addItem('🗑️ ログクリア', 'clearLogs')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 予約投稿デバッグ')
      .addItem('🔍 トリガー状態確認', 'checkScheduledPostTriggers')
      .addItem('📋 データ確認', 'checkScheduledPostsData')
      .addItem('🐛 予約投稿デバッグ実行', 'debugScheduledPosts')
      .addItem('💪 強制実行（過去含む）', 'forceProcessScheduledPosts'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 API管理')
      .addItem('📈 API使用状況確認', 'checkAPIUsageStatus')
      .addItem('🔄 API使用回数リセット（緊急用）', 'resetAPIQuotaManually'))
    .addToUi();
  
    // 初回起動時の設定チェック
    checkInitialSetup();
    
    // 既存シートのヘッダー行を固定
    freezeExistingSheetHeaders();
  } catch (error) {
    console.error('メニュー作成エラー:', error);
    // エラーが発生してもスプレッドシートは使えるようにする
  }
}

// ===========================
// 設定管理
// ===========================
function getConfig(key) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
    if (!sheet) {
      console.error('getConfig: Settingsシートが見つかりません');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }
    
    console.log(`getConfig: キー「${key}」が見つかりません`);
    return null;
  } catch (error) {
    console.error(`getConfig エラー: ${error.toString()}`);
    return null;
  }
}

function setConfig(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  
  // 新しい設定項目の追加
  sheet.appendRow([key, value, '']);
}

// ===========================
// 初期設定チェック
// ===========================
function checkInitialSetup() {
  const requiredConfigs = ['CLIENT_ID', 'CLIENT_SECRET', 'REDIRECT_URI'];
  const missingConfigs = [];
  
  requiredConfigs.forEach(config => {
    const value = getConfig(config);
    if (!value || value === '（後で入力）') {
      missingConfigs.push(config);
    }
  });
  
  if (missingConfigs.length > 0) {
    SpreadsheetApp.getUi().alert(
      '初期設定が必要です',
      `以下の設定を「基本設定」シートに入力してください：\n${missingConfigs.join(', ')}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===========================
// OAuth認証
// ===========================
function doGet(e) {
  // OAuthコールバック処理
  if (e.parameter.code) {
    try {
      const code = e.parameter.code;
      
      // 認証コードを一時的に保存
      PropertiesService.getScriptProperties().setProperty('temp_auth_code', code);
      
      // 成功ページを表示
      return HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 40px; text-align: center;">
          <h2 style="color: #1DA1F2;">認証成功！</h2>
          <p>認証が成功しました。このタブを閉じて、スプレッドシートに戻ってください。</p>
          <p>スプレッドシートで「認証コード処理」を実行してください。</p>
          <br>
          <button onclick="window.close()" style="padding: 10px 20px; background-color: #1DA1F2; color: white; border: none; border-radius: 5px; cursor: pointer;">
            このタブを閉じる
          </button>
        </div>
      `);
    } catch (error) {
      return HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 40px; text-align: center;">
          <h2 style="color: #E1306C;">エラー</h2>
          <p>認証処理中にエラーが発生しました：</p>
          <p style="color: red;">${error.toString()}</p>
        </div>
      `);
    }
  }
  
  // エラーパラメータがある場合
  if (e.parameter.error) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 40px; text-align: center;">
        <h2 style="color: #E1306C;">認証エラー</h2>
        <p>認証が拒否されました：</p>
        <p style="color: red;">${e.parameter.error_description || e.parameter.error}</p>
      </div>
    `);
  }
  
  // デフォルトページ
  return HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 40px; text-align: center;">
      <h2>Threads自動化ツール</h2>
      <p>このページは認証コールバック用です。</p>
    </div>
  `);
}

function startOAuth() {
  const clientId = getConfig('CLIENT_ID');
  const redirectUri = getConfig('REDIRECT_URI');
  
  if (!clientId || !redirectUri) {
    SpreadsheetApp.getUi().alert('CLIENT_IDとREDIRECT_URIを設定してください');
    return;
  }
  
  const authUrl = `https://threads.net/oauth/authorize?` +
    `client_id=${clientId}` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&scope=threads_basic,threads_content_publish,threads_read_replies,threads_manage_insights` +
    `&response_type=code`;
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>Threads認証</h3>
      <p>以下のURLにアクセスして認証を完了してください：</p>
      <p><a href="${authUrl}" target="_blank">認証ページを開く</a></p>
      <br>
      <p>認証後に表示されるコードを下記に入力してください：</p>
      <input type="text" id="authCode" style="width: 300px; padding: 5px;">
      <button onclick="submitCode()">送信</button>
      <div id="result"></div>
    </div>
    <script>
      function submitCode() {
        const code = document.getElementById('authCode').value;
        if (code) {
          google.script.run.withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .exchangeCodeForToken(code);
        }
      }
      function showResult(result) {
        document.getElementById('result').innerHTML = '<p style="color: green;">認証成功！</p>';
      }
      function showError(error) {
        document.getElementById('result').innerHTML = '<p style="color: red;">エラー: ' + error + '</p>';
      }
    </script>
  `).setWidth(400).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Threads認証');
}

function processAuthCode() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const code = PropertiesService.getScriptProperties().getProperty('temp_auth_code');
    
    if (!code) {
      // 手動で認証コードを入力
      const response = ui.prompt(
        '認証コード入力',
        '認証後に表示されたコードを入力してください：',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() === ui.Button.OK && response.getResponseText()) {
        const manualCode = response.getResponseText();
        const result = exchangeCodeForToken(manualCode);
        
        if (result === 'success') {
          ui.alert('成功', 'アクセストークンの取得に成功しました！', ui.ButtonSet.OK);
        }
      }
      return;
    }
    
    // 自動取得された認証コード
    const result = exchangeCodeForToken(code);
    
    if (result === 'success') {
      PropertiesService.getScriptProperties().deleteProperty('temp_auth_code');
      ui.alert('成功', 'アクセストークンの取得に成功しました！', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('エラー', `認証処理中にエラーが発生しました：${error.toString()}`, ui.ButtonSet.OK);
    logError('processAuthCode', error);
  }
}

function exchangeCodeForToken(authCode) {
  const clientId = getConfig('CLIENT_ID');
  const clientSecret = getConfig('CLIENT_SECRET');
  const redirectUri = getConfig('REDIRECT_URI');
  
  const tokenUrl = 'https://graph.threads.net/oauth/access_token';
  
  const payload = {
    'client_id': clientId,
    'client_secret': clientSecret,
    'grant_type': 'authorization_code',
    'redirect_uri': redirectUri,
    'code': authCode
  };
  
  try {
    const response = fetchWithTracking(tokenUrl, {
      'method': 'POST',
      'payload': payload,
      'muteHttpExceptions': true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.access_token) {
      // 短期トークンを長期トークンに交換
      exchangeForLongLivedToken(result.access_token);
      return 'success';
    } else {
      throw new Error(result.error_message || 'トークン取得失敗');
    }
  } catch (error) {
    logError('exchangeCodeForToken', error);
    throw error;
  }
}

function exchangeForLongLivedToken(shortLivedToken) {
  const clientSecret = getConfig('CLIENT_SECRET');
  const exchangeUrl = 'https://graph.threads.net/oauth/access_token';
  
  const payload = {
    'grant_type': 'th_exchange_token',
    'client_secret': clientSecret,
    'access_token': shortLivedToken
  };
  
  try {
    const response = fetchWithTracking(exchangeUrl, {
      'method': 'POST',
      'payload': payload
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.access_token) {
      setConfig('ACCESS_TOKEN', result.access_token);
      setConfig('TOKEN_EXPIRES', new Date(Date.now() + 60 * 24 * 60 * 60 * 1000).toISOString());
      
      // ユーザー情報の取得
      getUserInfo(result.access_token);
      
      logOperation('OAuth認証', 'success', 'トークン取得成功');
    }
  } catch (error) {
    logError('exchangeForLongLivedToken', error);
  }
}

// ===========================
// トリガー設定
// ===========================
function setupTriggers() {
  const html = HtmlService.createHtmlOutputFromFile('TriggerDialog')
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'トリガー設定');
}

// トリガー設定処理
function processTriggerSettings(postInterval, replyInterval, tokenHour) {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });
    
    // 予約投稿用トリガー
    if (postInterval <= 30) {
      ScriptApp.newTrigger('processScheduledPosts')
        .timeBased()
        .everyMinutes(postInterval)
        .create();
    } else {
      ScriptApp.newTrigger('processScheduledPosts')
        .timeBased()
        .everyHours(1)
        .create();
    }
    
    // リプライ取得＋自動返信の統合トリガー
    if (replyInterval <= 30) {
      ScriptApp.newTrigger('fetchAndAutoReply')
        .timeBased()
        .everyMinutes(replyInterval)
        .create();
    } else if (replyInterval <= 60) {
      ScriptApp.newTrigger('fetchAndAutoReply')
        .timeBased()
        .everyHours(1)
        .create();
    } else {
      // 60分を超える場合は、時間単位で設定
      const hours = Math.floor(replyInterval / 60);
      ScriptApp.newTrigger('fetchAndAutoReply')
        .timeBased()
        .everyHours(hours)
        .create();
    }
    
    // トークンリフレッシュ用トリガー（毎日）
    ScriptApp.newTrigger('refreshAccessToken')
      .timeBased()
      .everyDays(1)
      .atHour(tokenHour)
      .create();
    
    SpreadsheetApp.getUi().alert(
      'トリガー設定完了',
      `以下の設定でトリガーを作成しました:\n\n` +
      `• 予約投稿: ${postInterval}分ごと\n` +
      `• リプライ取得: ${replyInterval}分ごと${replyInterval >= 60 ? ' (' + (replyInterval / 60) + '時間)' : ''}\n` +
      `• トークン更新: 毎日${tokenHour}時`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    logOperation('トリガー設定', 'success', 
      `投稿:${postInterval}分, リプライ:${replyInterval}分, トークン:${tokenHour}時`);
    
  } catch (error) {
    logError('processTriggerSettings', error);
    throw error;
  }
}

// ===========================
// ユーティリティ関数
// ===========================
function logOperation(operation, status, details) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');
    if (!logSheet) return;

    const timestamp = new Date();
    details = details || '';
    
    // 2行目に新しい行を挿入
    logSheet.insertRowAfter(1);
    
    // 新しく挿入した行（2行目）のRangeを取得
    const newRow = logSheet.getRange(2, 1, 1, 4); // 日時, 操作, ステータス, 詳細 の4列
    
    // 値を設定
    newRow.setValues([[timestamp, operation, status, details]]);
    
    // 書式を標準にリセット
    newRow.setBackground(null).setFontColor('#000000').setFontWeight('normal');
    
    // 日時フォーマットを設定（年月日時分秒）
    logSheet.getRange(2, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');

    // ログが2000行を超えたら古いもの（2001行目以降）を削除
    if (logSheet.getLastRow() > 2000) {
      logSheet.deleteRows(2001, logSheet.getLastRow() - 2000);
    }
  } catch (error) {
    console.error('logOperation エラー:', error);
  }
}

function logError(functionName, error) {
  console.error(`Error in ${functionName}:`, error);
  logOperation(functionName, 'error', error.toString());
  
  // エラー通知
  const email = getConfig('NOTIFICATION_EMAIL');
  if (email && getConfig('ENVIRONMENT') === 'production') {
    MailApp.sendEmail({
      to: email,
      subject: 'Threads自動化ツール エラー通知',
      body: `関数: ${functionName}\nエラー: ${error.toString()}\n時刻: ${new Date()}`
    });
  }
}

function clearLogs() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'ログクリア',
    'すべてのログを削除しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // クリア実行のログを追加（新しい方式で）
    logOperation('ログクリア', 'info', 'ログをクリアしました');
    
    ui.alert('ログをクリアしました');
  }
}

// ===========================
// ログ統計表示
// ===========================
function showLogStats() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    ui.alert('ログ統計', 'ログがありません。', ui.ButtonSet.OK);
    return;
  }
  
  // 統計を計算
  const stats = {
    total: data.length - 1,
    today: 0,
    statusCount: {
      success: 0,
      info: 0,
      warning: 0,
      error: 0
    },
    operations: {}
  };
  
  const today = new Date().toDateString();
  
  for (let i = 1; i < data.length; i++) {
    const [timestamp, operation, status] = data[i];
    const date = new Date(timestamp);
    
    if (date.toDateString() === today) stats.today++;
    
    if (stats.statusCount.hasOwnProperty(status)) {
      stats.statusCount[status]++;
    }
    
    stats.operations[operation] = (stats.operations[operation] || 0) + 1;
  }
  
  // レポート作成
  let report = '📊 ログ統計\n\n';
  report += `総ログ数: ${stats.total}件\n`;
  report += `今日のログ: ${stats.today}件\n\n`;
  
  report += '📈 ステータス別:\n';
  report += `  ✅ 成功: ${stats.statusCount.success}件\n`;
  report += `  ℹ️ 情報: ${stats.statusCount.info}件\n`;
  report += `  ⚠️ 警告: ${stats.statusCount.warning}件\n`;
  report += `  ❌ エラー: ${stats.statusCount.error}件\n\n`;
  
  report += '🔝 頻度の高い操作 (上位5件):\n';
  const sortedOps = Object.entries(stats.operations)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  sortedOps.forEach(([op, count], index) => {
    report += `  ${index + 1}. ${op}: ${count}件\n`;
  });
  
  ui.alert('ログ統計', report, ui.ButtonSet.OK);
}

// ===========================
// ユーザー情報取得
// ===========================
function getUserInfo(accessToken) {
  try {
    const response = fetchWithTracking(`${THREADS_API_BASE}/v1.0/me?fields=id,username,threads_profile_picture_url`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    const userInfo = JSON.parse(response.getContentText());
    setConfig('USER_ID', userInfo.id);
    setConfig('USERNAME', userInfo.username);
    
    logOperation('ユーザー情報取得', 'success', `@${userInfo.username}`);
  } catch (error) {
    logError('getUserInfo', error);
  }
}


// ===========================
// 基本設定シート再構成
// ===========================
function resetSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // パスワード確認
  if (!verifyPassword('基本設定シート再構成')) {
    return;
  }
  
  const response = ui.alert(
    '基本設定シート再構成',
    '基本設定シートを削除して再作成しますか？\n\n' +
    '⚠️ 既存の設定（トークン等）はすべて削除されます。\n' +
    '再度認証が必要になります。',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      initializeSettingsSheet();
      ui.alert('基本設定シートを再構成しました。\n\n' +
        '必要な情報を入力してから、認証を実行してください。');
    } catch (error) {
      console.error('基本設定シート再構成エラー:', error);
      ui.alert('エラーが発生しました: ' + error.message);
    }
  }
}

// ===========================
// 基本設定シート初期化
// ===========================
function initializeSettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを削除
  let existingSheet = spreadsheet.getSheetByName('基本設定');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet('基本設定');
  
  // ヘッダー行を設定
  const headers = ['設定項目', '値', '説明'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 設定項目のデータ
  const settings = [
    ['CLIENT_ID', '（後で入力）', 'Meta開発者ダッシュボードのThreadsアプリID【必須】'],
    ['CLIENT_SECRET', '（後で入力）', 'Meta開発者ダッシュボードのThreadsアプリシークレット【必須】'],
    ['ACCESS_TOKEN', '', '長期アクセストークン【必須】'],
    ['USER_ID', '', '（自動入力）ThreadsユーザーID'],
    ['USERNAME', '', '（自動入力）Threadsユーザー名'],
    ['TOKEN_EXPIRES', '', '（自動入力）トークン有効期限'],
    ['REDIRECT_URI', 'https://script.google.com/a/macros/tsukichiyo.jp/s/AKfycbwZQCRvj97_y_fAUTlWKvC3EsDCoyDRaQT0tALUKK2ZvQXSNr-fFimDPnkFD_N6yimi/exec', 'OAuth認証のリダイレクトURI（Google Apps ScriptのURL）この値は固定で自動的に入ります'],
    ['SCRIPT_ID', '', '（自動入力）このスプレッドシートのスクリプトID']
  ];
  
  sheet.getRange(2, 1, settings.length, settings[0].length).setValues(settings);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150);  // 設定項目
  sheet.setColumnWidth(2, 300);  // 値
  sheet.setColumnWidth(3, 400);  // 説明
  
  // 説明コメントを追加
  sheet.getRange('A1').setNote(
    '基本設定シート\n\n' +
    '必須項目:\n' +
    '1. CLIENT_ID - Meta開発者ダッシュボードのThreadsアプリから取得\n' +
    '2. CLIENT_SECRET - Meta開発者ダッシュボードのThreadsアプリから取得\n' +
    '3. REDIRECT_URI - Google Apps ScriptのウェブアプリケーションURL\n' +
    '4. ACCESS_TOKEN - 長期アクセストークン\n\n' +
    '自動入力項目:\n' +
    '- ACCESS_TOKEN, USER_ID, USERNAME, TOKEN_EXPIRES は認証時に自動設定されます'
  );
  
  // 必須項目の背景色を設定
  sheet.getRange(2, 1, 1, 3).setBackground('#FFF3CD');  // CLIENT_ID
  sheet.getRange(3, 1, 1, 3).setBackground('#FFF3CD');  // CLIENT_SECRET
  sheet.getRange(4, 1, 1, 3).setBackground('#FFF3CD');  // ACCESS_TOKEN
  
  logOperation('基本設定シート初期化', 'success', 'シートを再構成しました');
}

// ===========================
// ログシート再構成
// ===========================
function resetLogsSheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'ログシート再構成',
    'ログシートを削除して再作成しますか？\n\n' +
    '⚠️ 既存のログはすべて削除されます。',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      initializeLogsSheet();
      ui.alert('ログシートを再構成しました。');
    } catch (error) {
      console.error('ログシート再構成エラー:', error);
      ui.alert('エラーが発生しました: ' + error.message);
    }
  }
}

// ===========================
// ログシート初期化
// ===========================
function initializeLogsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを削除
  let existingSheet = spreadsheet.getSheetByName('ログ');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet('ログ');
  
  // ヘッダー行を設定
  const headers = [
    '日時',
    '操作',
    'ステータス',
    '詳細'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 1行目を固定
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // 日時
  sheet.setColumnWidth(2, 200); // 操作
  sheet.setColumnWidth(3, 100); // ステータス
  sheet.setColumnWidth(4, 500); // 詳細
  
  // 日付列のフォーマット
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  
  // 条件付き書式を追加（ステータスの視覚化）
  const statusRange = sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1);
  
  const successRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('success')
    .setFontColor('#155724')
    .setRanges([statusRange])
    .build();
  
  const infoRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('info')
    .setFontColor('#0C5460')
    .setRanges([statusRange])
    .build();
    
  const warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('warning')
    .setFontColor('#856404')
    .setRanges([statusRange])
    .build();
    
  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('error')
    .setFontColor('#721C24')
    .setRanges([statusRange])
    .build();
  
  sheet.setConditionalFormatRules([successRule, infoRule, warningRule, errorRule]);
  
  // 説明コメントを追加
  sheet.getRange('A1').setNote(
    'システムログシート\n\n' +
    '日時: 操作が実行された日時\n' +
    '操作: 実行された機能名\n' +
    'ステータス: success(成功), info(情報), warning(警告), error(エラー)\n' +
    '詳細: 操作の詳細情報やエラーメッセージ'
  );
  
  // 初期ログエントリ（直接2行目に追加）
  sheet.getRange(2, 1, 1, 4).setValues([[
    new Date(),
    'ログシート初期化',
    'success',
    'ログシートを再構成しました'
  ]]);
  
  // 初期ログのフォントウェイトを標準に
  sheet.getRange(2, 1, 1, 4).setFontWeight('normal');
}

// ===========================
// すべてのシートを再構成
// ===========================
function resetAllSheets() {
  const ui = SpreadsheetApp.getUi();
  
  // パスワード確認
  if (!verifyPassword('すべてのシート再構成')) {
    return;
  }
  
  const response = ui.alert(
    'すべてのシートを再構成',
    'すべてのシートを削除して再作成しますか？\n\n' +
    '⚠️ 警告:\n' +
    '・すべてのデータが削除されます\n' +
    '・設定情報も初期化されます\n' +
    '・この操作は取り消せません',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      // 各シートを順番に再構成
      initializeSettingsSheet();
      SpreadsheetApp.flush();
      
      initializeScheduledPostsSheet();
      SpreadsheetApp.flush();
      
      initializeRepliesSheet();
      SpreadsheetApp.flush();
      
      initializeキーワード設定Sheet();
      SpreadsheetApp.flush();
      
      initializeReplyHistorySheet();
      SpreadsheetApp.flush();
      
      initializeLogsSheet();
      SpreadsheetApp.flush();
      
      // 完了メッセージ
      ui.alert(
        '再構成完了',
        'すべてのシートを再構成しました。\n\n' +
        '次のステップ:\n' +
        '1. 基本設定シートに必要な情報を入力\n' +
        '2. 初期設定を実行\n' +
        '3. トリガーを設定',
        ui.ButtonSet.OK
      );
      
      logOperation('全シート再構成', 'success', 'すべてのシートを再構成しました');
      
    } catch (error) {
      console.error('全シート再構成エラー:', error);
      ui.alert('エラー', 'シートの再構成中にエラーが発生しました:\n' + error.message, ui.ButtonSet.OK);
      logError('resetAllSheets', error);
    }
  }
}

// ===========================
// 既存シートの行固定設定
// ===========================
function freezeExistingSheetHeaders() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToFreeze = [
    '基本設定',
    '予約投稿',
    'リプライ追跡',
    '自動返信キーワード設定',
    '自動応答結果',
    'ログ'
  ];
  
  sheetsToFreeze.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet && sheet.getFrozenRows() === 0) {
      sheet.setFrozenRows(1);
    }
  });
}

// ===========================
// シート保護機能
// ===========================
// パスワード定数
const SHEET_PROTECTION_PASSWORD = 'tsukichiyo.inc@gmail.com';

// 共通のパスワード確認関数
function verifyPassword(promptTitle) {
  const ui = SpreadsheetApp.getUi();
  const passwordPrompt = ui.prompt(
    promptTitle || 'パスワード入力',
    'パスワードを入力してください',
    ui.ButtonSet.OK_CANCEL
  );

  if (passwordPrompt.getSelectedButton() !== ui.Button.OK) {
    return false;
  }

  const input = passwordPrompt.getResponseText();
  
  if (input !== SHEET_PROTECTION_PASSWORD) {
    ui.alert('エラー', 'パスワードが違います。', ui.ButtonSet.OK);
    return false;
  }
  
  return true;
}

function showSettingsSheet() {
  if (!verifyPassword('基本設定シート表示')) {
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
  if (!sheet) {
    ui.alert('エラー', '「基本設定」シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  sheet.showSheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  ui.alert('成功', '「基本設定」シートを表示しました。', ui.ButtonSet.OK);
  logOperation('基本設定シート表示', 'success', 'パスワード認証成功');
}

function hideSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
  
  if (!sheet) {
    ui.alert('エラー', '「基本設定」シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  const result = ui.alert(
    '確認',
    '「基本設定」シートを非表示にしますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    sheet.hideSheet();
    ui.alert('完了', '「基本設定」シートを非表示にしました。', ui.ButtonSet.OK);
    logOperation('基本設定シート非表示', 'success', 'シートを非表示にしました');
  }
}

function initializeSheetProtection() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
  if (sheet) {
    sheet.hideSheet();
    console.log('基本設定シートを初期状態で非表示に設定しました。');
    logOperation('基本設定シート保護初期化', 'success', '初期状態で非表示に設定');
  }
}

// ===========================
// トリガー所有者確認機能
// ===========================
function checkTriggerOwners() {
  // パスワード確認
  if (!verifyPassword('トリガー所有者確認')) {
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      ui.alert('トリガー情報', '現在設定されているトリガーはありません。', ui.ButtonSet.OK);
      return;
    }
    
    // トリガー情報を整理
    let triggerInfo = '📋 現在のトリガー所有者情報\n\n';
    const triggersByOwner = {};
    
    triggers.forEach((trigger, index) => {
      const handlerFunction = trigger.getHandlerFunction();
      const triggerSource = trigger.getTriggerSource();
      const eventType = trigger.getEventType();
      
      // トリガー所有者の取得（メールアドレス）
      let ownerEmail = 'unknown';
      try {
        // トリガーのユニークIDから所有者を推定
        ownerEmail = trigger.getUniqueId();
        // より詳細な情報が必要な場合は、Session.getActiveUser()等を検討
      } catch (e) {
        ownerEmail = '取得不可';
      }
      
      // トリガータイプの判定
      let triggerType = '';
      if (triggerSource.toString() === 'SPREADSHEETS') {
        triggerType = 'スプレッドシート';
      } else if (triggerSource.toString() === 'CLOCK') {
        triggerType = '時間ベース';
      }
      
      triggerInfo += `【トリガー ${index + 1}】\n`;
      triggerInfo += `  関数: ${handlerFunction}\n`;
      triggerInfo += `  タイプ: ${triggerType}\n`;
      triggerInfo += `  イベント: ${eventType}\n`;
      triggerInfo += `  ユニークID: ${trigger.getUniqueId()}\n`;
      triggerInfo += '\n';
    });
    
    // 現在のユーザー情報も追加
    triggerInfo += '━━━━━━━━━━━━━━━━━━\n';
    triggerInfo += '💡 トリガー所有者について:\n';
    triggerInfo += '• トリガーは作成したユーザーが所有者となります\n';
    triggerInfo += '• UrlFetchAppの実行はトリガー所有者のクォータを消費します\n';
    triggerInfo += '• 「トリガーを再設定」で所有者を変更できます\n';
    
    // アラートで表示
    ui.alert('トリガー所有者情報', triggerInfo, ui.ButtonSet.OK);
    
    // ログに記録
    logOperation('トリガー所有者確認', 'success', `${triggers.length}個のトリガーを確認`);
    
  } catch (error) {
    ui.alert('エラー', 'トリガー情報の取得中にエラーが発生しました:\n' + error.toString(), ui.ButtonSet.OK);
    logError('checkTriggerOwners', error);
  }
}

