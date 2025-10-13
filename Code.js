// Threads自動化ツール - メインコード
// Code.gs

// ===========================
// グローバル設定
// ===========================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const THREADS_API_BASE = 'https://graph.threads.net';
const LOG_MAX_ENTRIES = 150; // ログシートの最大保持件数（ヘッダー除く）
const REPLY_HISTORY_MAX_ENTRIES = 200; // 自動応答結果の最大件数
const REPLY_HISTORY_MAX_DAYS = 30; // 自動応答結果の保持日数
const REPLIES_MAX_ENTRIES = 200; // 受信したリプライの最大件数
const REPLIES_MAX_DAYS = 150; // 受信したリプライの保持日数
const REDIRECT_URI = 'https://script.google.com/macros/s/1MenZk9WxpZAv0U79xESYwz3aLENP_toLQuanvx_f7imC5uPijTOUBj5k/exec'; // OAuth認証のリダイレクトURI

// シート名の定数（全体で統一）
const KEYWORD_SHEET_NAME = 'キーワード設定';
const REPLY_HISTORY_SHEET_NAME = '自動応答結果';
const REPLIES_SHEET_NAME = '受信したリプライ';

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
// 設定テスト関数
// ===========================
function testConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const results = [];

  try {
    // 基本設定のテスト
    const clientId = getConfig('CLIENT_ID');
    const clientSecret = getConfig('CLIENT_SECRET');
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');

    results.push(`CLIENT_ID: ${clientId ? '設定済み' : '未設定'}`);
    results.push(`CLIENT_SECRET: ${clientSecret ? '設定済み' : '未設定'}`);
    results.push(`ACCESS_TOKEN: ${accessToken ? '設定済み' : '未設定'}`);
    results.push(`USER_ID: ${userId ? '設定済み' : '未設定'}`);

    // API接続テスト（アクセス可能な場合）
    if (accessToken && userId) {
      try {
        const response = fetchWithTracking(
          `${THREADS_API_BASE}/v1.0/${userId}?fields=id,username`,
          { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        if (response.getResponseCode() === 200) {
          results.push('API接続: ✅ 正常');
        } else {
          results.push(`API接続: ❌ ステータスコード ${response.getResponseCode()}`);
        }
      } catch (apiError) {
        results.push(`API接続: ❌ ${apiError.toString()}`);
      }
    } else {
      results.push('API接続: ⚠️ 認証情報が不完全');
    }

    // 結果表示
    const message = '設定テスト結果:\n\n' + results.join('\n');
    ui.alert('設定テスト', message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('テストエラー', `設定テスト中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
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
      .addToUi();
    
    // Threads自動化メニュー
    ui.createMenu('Threads自動化')
    .addItem('🚀 クイックセットアップ', 'quickSetupWithExistingToken')
    .addSeparator()
    .addItem('⏰ トリガーを再設定', 'resetAutomationTriggers')
    .addItem('⚠️ 全トリガーを削除', 'deleteAllTriggers')
    .addSeparator()
    .addItem('📤 手動投稿実行', 'manualPostExecution')
    .addItem('🧵 最新投稿50件を取得', 'fetchLatestThreadsPosts')
    .addItem('💬 リプライ＋自動返信（統合実行）', 'fetchAndAutoReply')
    .addItem('🔄 自動返信のみ', 'manualAutoReply')
    .addItem('⏪ 過去6時間を再処理', 'manualBackfill6Hours')
    .addSeparator()
    .addItem('🔑 トークン手動更新', 'manualRefreshToken')
    .addItem('🧹 データクリーンアップ', 'manualCleanupData')
    .addSeparator()
    .addSubMenu(ui.createMenu('🐛 デバッグ')
      .addItem('🔍 認証情報デバッグ', 'debugConfigForSpreadsheet')
      .addItem('🧪 自動返信テスト実行', 'testAutoReplyForCurrentSheet')
      .addItem('🧪 返信シミュレーション', 'simulateAutoReply')
      .addItem('⚙️ 設定テスト', 'testConfiguration')
      .addItem('🔎 自動返信システムデバッグ', 'debugAutoReplySystem')
      .addItem('📊 自動返信統計', 'showAutoReplyStats'))
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
    // まず基本設定シートから取得を試みる
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
    if (!sheet) {
      console.error('getConfig: 基本設定シートが見つかりません');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        const value = data[i][1];
        // 値が存在し、空でない場合は返す
        if (value && value !== '' && value !== '（後で入力）') {
          // 「設定済み」と表示されている場合は、Script Propertiesから取得を試みる
          if (value === '設定済み') {
            const scriptValue = PropertiesService.getScriptProperties().getProperty(key);
            if (scriptValue) {
              console.log(`機密情報 ${key} を Script Properties から取得しました`);
              return scriptValue;
            }
            console.log(`getConfig: キー「${key}」は設定済みですが、値を取得できません`);
            return null;
          }
          return value;
        }
      }
    }

    // シートに見つからない場合、機密情報の場合はScript Propertiesも確認
    const sensitiveKeys = ['ACCESS_TOKEN', 'CLIENT_SECRET', 'APP_SECRET', 'WEBHOOK_VERIFY_TOKEN'];
    if (sensitiveKeys.includes(key)) {
      const value = PropertiesService.getScriptProperties().getProperty(key);
      if (value) {
        console.log(`機密情報 ${key} を Script Properties から取得しました`);
        return value;
      }
    }

    console.log(`getConfig: キー「${key}」が見つかりません`);
    return null;
  } catch (error) {
    console.error(`getConfig エラー: ${error.toString()}`);
    logError('getConfig', error);
    return null;
  }
}

function setConfig(key, value) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
    if (!sheet) {
      console.error('setConfig: 基本設定シートが見つかりません');
      return;
    }

    const data = sheet.getDataRange().getValues();

    // まず基本設定シートに値を保存
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }

    // 新しい設定項目の追加
    if (!found) {
      sheet.appendRow([key, value, '']);
    }

    // 機密情報は追加でScript Propertiesにも保存（二重保存）
    const sensitiveKeys = ['ACCESS_TOKEN', 'CLIENT_SECRET', 'APP_SECRET'];
    if (sensitiveKeys.includes(key) && value) {
      PropertiesService.getScriptProperties().setProperty(key, value);
      console.log(`機密情報 ${key} を Script Properties にも保存しました`);
    }

  } catch (error) {
    console.error(`setConfig エラー: ${error.toString()}`);
    logError('setConfig', error);
  }
}

// ===========================
// 初期設定チェック
// ===========================
function checkInitialSetup() {
  const requiredConfigs = ['CLIENT_ID', 'CLIENT_SECRET'];
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
// OAuth認証（現在未使用 - 長期トークン直接入力方式を採用）
// ===========================
// 注意: 以下のOAuth認証関連の関数は現在使用されていませんが、
// 将来的にOAuth認証フローが必要になった場合のために保持しています。
// 現在のクイックセットアップでは長期トークンを直接入力する方式を採用しています。

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
  
  if (!clientId) {
    SpreadsheetApp.getUi().alert('CLIENT_IDを設定してください');
    return;
  }
  
  const authUrl = `https://threads.net/oauth/authorize?` +
    `client_id=${clientId}` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&scope=threads_basic,threads_publish,threads_manage_replies,threads_read_replies,threads_manage_insights` +
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
  
  const tokenUrl = 'https://graph.threads.net/oauth/access_token';
  
  const payload = {
    'client_id': clientId,
    'client_secret': clientSecret,
    'grant_type': 'authorization_code',
    'redirect_uri': REDIRECT_URI,
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
      
      // 更新日時を保存（現在時刻）
      const now = new Date();
      const formattedDate = Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss');
      setConfig('TOKEN_EXPIRES', formattedDate);

      // ユーザー情報の取得
      getUserInfo(result.access_token);
      
      logOperation('OAuth認証', 'success', `トークン取得成功 (更新日時: ${formattedDate})`);
    }
  } catch (error) {
    logError('exchangeForLongLivedToken', error);
  }
}


// ===========================
// トリガーを再設定（デフォルト値で作成）
// ===========================
function resetAutomationTriggers() {
  // ドロップダウンで間隔を選択するダイアログを表示
  const html = HtmlService.createHtmlOutputFromFile('TriggerResetDialog')
    .setWidth(420)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'トリガーを再設定');
}

// ユーザー選択値をもとにトリガー再設定を適用
function applyAutomationTriggerSettings(postIntervalMinutes, replyIntervalMinutes) {
  const ui = SpreadsheetApp.getUi();
  const deleted = [];
  const created = [];

  try {
    const targets = new Set(['processScheduledPosts', 'fetchAndAutoReply', 'refreshLongLivedToken', 'fetchAndSaveReplies', 'autoCleanupData']);
    const triggers = ScriptApp.getProjectTriggers();

    // 対象トリガー削除
    triggers.forEach(tr => {
      const handler = tr.getHandlerFunction();
      if (targets.has(handler)) {
        try {
          ScriptApp.deleteTrigger(tr);
          deleted.push(handler);
        } catch (e) {
          console.error('トリガー削除エラー:', handler, e);
        }
      }
    });

    // 予約投稿トリガー: 5〜60分（5分刻み）。60は毎時に変換
    if (postIntervalMinutes >= 60) {
      ScriptApp.newTrigger('processScheduledPosts')
        .timeBased()
        .everyHours(1)
        .create();
      created.push('processScheduledPosts(60分≒毎時)');
    } else {
      ScriptApp.newTrigger('processScheduledPosts')
        .timeBased()
        .everyMinutes(postIntervalMinutes)
        .create();
      created.push(`processScheduledPosts(${postIntervalMinutes}分)`);
    }

    // リプライ取得＋自動返信: 30〜150分（30分刻み）+ 10分(非推奨)
    if (replyIntervalMinutes <= 59) {
      ScriptApp.newTrigger('fetchAndAutoReply')
        .timeBased()
        .everyMinutes(replyIntervalMinutes)
        .create();
      created.push(`fetchAndAutoReply(${replyIntervalMinutes}分)`);
    } else if (replyIntervalMinutes >= 60) {
      // Apps Scriptの制約により、90/150分等は時間単位に丸めます
      const hours = Math.max(1, Math.floor(replyIntervalMinutes / 60));
      ScriptApp.newTrigger('fetchAndAutoReply')
        .timeBased()
        .everyHours(hours)
        .create();
      created.push(`fetchAndAutoReply(約${hours}時間)`);
    }

    // 長期トークン更新は固定: 2週間ごとに午前3時
    ScriptApp.newTrigger('refreshLongLivedToken')
      .timeBased()
      .everyDays(14)
      .atHour(3)
      .create();
    created.push('refreshLongLivedToken(2週間ごと)');

    // データクリーンアップは固定: 毎週日曜日の午前3時
    ScriptApp.newTrigger('autoCleanupData')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(3)
      .create();
    created.push('autoCleanupData(毎週日曜日)');

    const msg = `削除: ${deleted.length ? deleted.join(', ') : 'なし'}\n作成: ${created.join(', ')}`;
    logOperation('トリガー再設定', 'success', msg);
    return { success: true, message: msg };

  } catch (error) {
    logError('applyAutomationTriggerSettings', error);
    return { success: false, error: error.toString() };
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

    // ログ行数制限の処理
    // ヘッダー行を含めた最大行数
    const maxTotalRows = 1 + LOG_MAX_ENTRIES;
    const currentLastRow = logSheet.getLastRow();
    
    // 現在の行数が最大行数を超えている場合、古い行を削除
    if (currentLastRow > maxTotalRows) {
      const rowsToDelete = currentLastRow - maxTotalRows;
      // 最古の行から削除（maxTotalRows + 1行目から最終行まで）
      logSheet.deleteRows(maxTotalRows + 1, rowsToDelete);
    }
  } catch (error) {
    console.error('logOperation エラー:', error);
  }
}

function logError(functionName, error, context = {}) {
  try {
    // LoggingUtilities.jsのlogError関数を使用（より高度なログ機能）
    if (typeof logError === 'function' && logError !== arguments.callee) {
      logError(functionName, error, context);
    } else {
      // フォールバック：シンプルなログ
      console.error(`Error in ${functionName}:`, error);
      logOperation(functionName, 'error', error.toString());
    }

    // エラー通知（本番環境のみ）
    const email = getConfig('NOTIFICATION_EMAIL');
    const environment = getConfig('ENVIRONMENT');

    if (email && environment === 'production') {
      const subject = 'Threads自動化ツール エラー通知';
      const body = `関数: ${functionName}\nエラー: ${error.toString()}\n時刻: ${new Date()}\n\n詳細:\n${JSON.stringify(context, null, 2)}`;

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body
        });
        console.log('エラー通知メールを送信しました');
      } catch (mailError) {
        console.error('エラー通知メール送信失敗:', mailError);
      }
    }
  } catch (logError) {
    // ログ記録自体が失敗した場合の最終フォールバック
    console.error('logError function failed:', logError);
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
// データクリーンアップ関連
// ===========================
/**
 * 汎用シートクリーンアップ関数
 * @param {string} sheetName - シート名
 * @param {number} maxEntries - 最大保持件数
 * @param {number} maxDays - 最大保持日数
 * @param {number} dateColumnIndex - 日時列のインデックス（0始まり）
 * @return {number} 削除した行数
 */
function cleanupSheetData(sheetName, maxEntries, maxDays, dateColumnIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      console.log(`シート「${sheetName}」が見つかりません`);
      return 0;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return 0; // ヘッダーのみの場合
    }
    
    let totalDeleted = 0;
    
    // 1. 日数制限による削除
    if (maxDays > 0) {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - maxDays);
      
      const data = sheet.getRange(2, dateColumnIndex + 1, lastRow - 1, 1).getValues();
      const rowsToDelete = [];
      
      for (let i = 0; i < data.length; i++) {
        const cellDate = new Date(data[i][0]);
        if (!isNaN(cellDate.getTime()) && cellDate < cutoffDate) {
          rowsToDelete.push(i + 2); // 実際の行番号（ヘッダー+1）
        }
      }
      
      // 後ろから削除（行番号がずれないように）
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
        totalDeleted++;
      }
    }
    
    // 2. 件数制限による削除
    const currentLastRow = sheet.getLastRow();
    const currentDataRows = currentLastRow - 1; // ヘッダーを除く
    
    if (maxEntries > 0 && currentDataRows > maxEntries) {
      const rowsToDelete = currentDataRows - maxEntries;
      // 古いデータから削除（2行目から削除）
      sheet.deleteRows(2, rowsToDelete);
      totalDeleted += rowsToDelete;
    }
    
    return totalDeleted;
    
  } catch (error) {
    console.error(`クリーンアップエラー (${sheetName}):`, error);
    return 0;
  }
}

/**
 * 手動データクリーンアップ（メニューから呼び出し）
 */
function manualCleanupData() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'データクリーンアップ',
    '以下のシートから古いデータを削除します：\n\n' +
    `• ログシート: ${LOG_MAX_ENTRIES}件まで保持\n` +
    `• 自動応答結果: ${REPLY_HISTORY_MAX_ENTRIES}件 & ${REPLY_HISTORY_MAX_DAYS}日以上前を削除\n` +
    `• 受信したリプライ: ${REPLIES_MAX_ENTRIES}件 & ${REPLIES_MAX_DAYS}日以上前を削除\n\n` +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    let totalDeleted = 0;
    
    // ログシートのクリーンアップ（日数制限なし、件数のみ）
    const logDeleted = cleanupSheetData('ログ', LOG_MAX_ENTRIES, 0, 0);
    totalDeleted += logDeleted;
    
    // 自動応答結果シートのクリーンアップ
    const replyHistoryDeleted = cleanupSheetData(
      REPLY_HISTORY_SHEET_NAME,
      REPLY_HISTORY_MAX_ENTRIES,
      REPLY_HISTORY_MAX_DAYS,
      0
    );
    totalDeleted += replyHistoryDeleted;
    
    // 受信したリプライシートのクリーンアップ
    const repliesDeleted = cleanupSheetData(
      REPLIES_SHEET_NAME,
      REPLIES_MAX_ENTRIES,
      REPLIES_MAX_DAYS,
      1
    );
    totalDeleted += repliesDeleted;
    
    ui.alert(
      'クリーンアップ完了',
      `${totalDeleted}行の古いデータを削除しました。\n\n` +
      `• ログ: ${logDeleted}行\n` +
      `• 自動応答結果: ${replyHistoryDeleted}行\n` +
      `• 受信したリプライ: ${repliesDeleted}行`,
      ui.ButtonSet.OK
    );
    
    logOperation('データクリーンアップ', 'success', `合計${totalDeleted}行削除`);
    
  } catch (error) {
    ui.alert('エラー', `クリーンアップ中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
    logError('manualCleanupData', error);
  }
}

/**
 * 自動データクリーンアップ（トリガーから呼び出し）
 */
function autoCleanupData() {
  try {
    let totalDeleted = 0;
    
    // ログシートのクリーンアップ（日数制限なし、件数のみ）
    const logDeleted = cleanupSheetData('ログ', LOG_MAX_ENTRIES, 0, 0);
    totalDeleted += logDeleted;
    
    // 自動応答結果シートのクリーンアップ
    const replyHistoryDeleted = cleanupSheetData(
      REPLY_HISTORY_SHEET_NAME,
      REPLY_HISTORY_MAX_ENTRIES,
      REPLY_HISTORY_MAX_DAYS,
      0
    );
    totalDeleted += replyHistoryDeleted;
    
    // 受信したリプライシートのクリーンアップ
    const repliesDeleted = cleanupSheetData(
      REPLIES_SHEET_NAME,
      REPLIES_MAX_ENTRIES,
      REPLIES_MAX_DAYS,
      1
    );
    totalDeleted += repliesDeleted;
    
    logOperation(
      '自動データクリーンアップ',
      'success',
      `合計${totalDeleted}行削除 (ログ:${logDeleted}, 応答結果:${replyHistoryDeleted}, リプライ:${repliesDeleted})`
    );
    
  } catch (error) {
    logError('autoCleanupData', error);
  }
}

// ===========================
// トークン手動更新
// ===========================
function manualRefreshToken() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークン更新を実行
    refreshLongLivedToken();
    
    // 結果を確認
    const tokenExpires = getConfig('TOKEN_EXPIRES');
    if (tokenExpires) {
      ui.alert(
        '更新完了',
        `トークンの更新が完了しました。\n\n更新日時: ${tokenExpires}\n\nログシートで詳細を確認できます。`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '完了',
        'トークン更新処理が完了しました。\nログシートで詳細を確認してください。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert('エラー', `トークン更新中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
    logError('manualRefreshToken', error);
  }
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
// 最新投稿50件取得（シート出力）
// ===========================
function fetchLatestThreadsPosts() {
  const ui = SpreadsheetApp.getUi();
  try {
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    const username = getConfig('USERNAME');
    if (!accessToken || !userId) {
      ui.alert('エラー', '基本設定の ACCESS_TOKEN / USER_ID が未設定です。', ui.ButtonSet.OK);
      return;
    }

    // まずシートを用意
    const sheet = resetLatestPostsSheet_();

    // 取得: 最大50件（paging対応: 念のため2ページ目まで）
    const perPage = 50; // APIのlimitを50に
    let collected = [];
    let url = `${THREADS_API_BASE}/v1.0/${userId}/threads?fields=id,text,timestamp,media_type,media_url,permalink&limit=${perPage}`;

    for (let page = 0; page < 2 && collected.length < 50; page++) {
      const resp = fetchWithTracking(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` },
        muteHttpExceptions: true
      });
      const json = JSON.parse(resp.getContentText());
      if (json.error) {
        throw new Error(json.error.message || 'APIエラー');
      }
      const data = Array.isArray(json.data) ? json.data : [];
      collected = collected.concat(data);
      if (collected.length >= 50 || !json.paging || !json.paging.next) break;
      url = json.paging.next;
    }
    collected = collected.slice(0, 50);

    // 表示用データ整形
    const rows = collected.map(item => {
      const id = item.id || '';
      const text = item.text || '';
      const mediaType = item.media_type || '';
      const mediaUrl = item.media_url || '';
      const permalink = item.permalink || (username ? `https://www.threads.net/@${username}/post/${id}` : '');
      const ts = item.timestamp ? new Date(item.timestamp) : new Date();
      const tsJst = Utilities.formatDate(ts, 'JST', 'yyyy/MM/dd HH:mm:ss');
      const fetchedAt = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
      return [id, text, mediaType, mediaUrl, permalink, tsJst, fetchedAt];
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);
      // 自動列幅調整（必要列のみ）
      sheet.setColumnWidth(1, 220); // ID
      sheet.setColumnWidth(2, 500); // 本文
      sheet.setColumnWidth(3, 120); // メディアタイプ
      sheet.setColumnWidth(4, 320); // メディアURL
      sheet.setColumnWidth(5, 320); // 投稿URL
      sheet.setColumnWidth(6, 160); // 投稿日時
      sheet.setColumnWidth(7, 160); // 取得時刻
    }

    ui.alert('完了', `最新投稿を ${rows.length} 件取得しました。`, ui.ButtonSet.OK);
    logOperation('最新投稿取得', 'success', `${rows.length}件取得`);
  } catch (error) {
    logError('fetchLatestThreadsPosts', error);
    SpreadsheetApp.getUi().alert('エラー', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function resetLatestPostsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '最新投稿';
  const existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(name);

  const headers = ['ID', '本文', 'メディアタイプ', 'メディアURL', '投稿URL', '投稿日時(JST)', '取得時刻(JST)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 説明ノート
  sheet.getRange('A1').setNote(
    '直近の投稿一覧\n' +
    '・本文やURLは取得時の状態です\n' +
    '・投稿日時はJSTで表示'
  );

  return sheet;
}


// ===========================
// 基本設定シート再構成
// ===========================
function resetSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('基本設定シート再構成', 
    '基本設定シートを再構成しますか？\n既存の設定値は保持されます。', 
    ui.ButtonSet.YES_NO);
  
  if (!verifyPassword('基本設定シート再構成')) {
    return;
  }
  
  try {
    console.log('基本設定シート再構成開始');
    
    // 既存の値を保存（最小限のアクセス）
    const existingValues = {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('基本設定');
    
    if (sheet) {
      // 必要な範囲のみ取得
      const lastRow = Math.min(sheet.getLastRow(), 20); // 最大20行まで
      if (lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (let i = 0; i < data.length; i++) {
          if (data[i][0] && data[i][1]) {
            existingValues[data[i][0]] = data[i][1];
          }
        }
      }
      
      // 処理を分散
      Utilities.sleep(500);
      
      // シートを削除
      ss.deleteSheet(sheet);
      
      Utilities.sleep(500);
    }
    
    // 新しいシートを作成（シンプルに）
    const newSheet = ss.insertSheet('基本設定');
    
    // ヘッダーのみ設定
    newSheet.getRange(1, 1, 1, 3).setValues([['設定項目', '値', '説明']]);
    newSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    // 基本データの設定（最小限）
    const basicSettings = [
      ['CLIENT_ID', existingValues['CLIENT_ID'] || '', 'Threads App ID'],
      ['CLIENT_SECRET', existingValues['CLIENT_SECRET'] || '', 'Threads App Secret'],
      ['ACCESS_TOKEN', existingValues['ACCESS_TOKEN'] || '', 'アクセストークン'],
      ['USER_ID', existingValues['USER_ID'] || '', 'ユーザーID'],
      ['USERNAME', existingValues['USERNAME'] || '', 'ユーザー名'],
      ['SPREADSHEET_LINK', existingValues['SPREADSHEET_LINK'] || '', 'スプレッドシート共有リンク'],
      ['SCRIPT_ID', existingValues['SCRIPT_ID'] || '', 'スクリプトID'],
      ['TOKEN_EXPIRES', existingValues['TOKEN_EXPIRES'] || '', 'トークン更新日時']
    ];
    
    if (basicSettings.length > 0) {
      newSheet.getRange(2, 1, basicSettings.length, 3).setValues(basicSettings);
      // CLIENT_ID（B2セル）を書式なしテキストに設定
      newSheet.getRange(2, 2).setNumberFormat('@');
    }
    
    // 列幅のみ設定（装飾は最小限）
    newSheet.setColumnWidth(1, 150);
    newSheet.setColumnWidth(2, 300);
    newSheet.setColumnWidth(3, 400);
    
    ui.alert('基本設定シートを再構成しました。');
    console.log('基本設定シート再構成完了');
    
  } catch (error) {
    console.error('基本設定シート再構成エラー:', error);
    ui.alert('エラーが発生しました: ' + error.message);
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
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  // 設定項目のデータ
  const settings = [
    ['CLIENT_ID', '（後で入力）', 'Meta開発者ダッシュボードのThreadsアプリID【必須】'],
    ['CLIENT_SECRET', '（後で入力）', 'Meta開発者ダッシュボードのThreadsアプリシークレット【必須】'],
    ['ACCESS_TOKEN', '', '長期アクセストークン【必須】'],
    ['USER_ID', '', '（自動入力）ThreadsユーザーID'],
    ['USERNAME', '', '（自動入力）Threadsユーザー名'],
    ['SPREADSHEET_LINK', '', '（自動入力）このスプレッドシートの共有リンク'],
    ['SCRIPT_ID', '', '（自動入力）このスプレッドシートのスクリプトID'],
    ['TOKEN_EXPIRES', '', '（自動入力）トークン更新日時']
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
    '3. ACCESS_TOKEN - 長期アクセストークン\n\n' +
    '自動入力項目:\n' +
    '- USER_ID, USERNAME, SPREADSHEET_LINK, SCRIPT_ID, TOKEN_EXPIRES（トークン更新日時）は認証時に自動設定されます'
  );
  
  // 必須項目の背景色を設定
  sheet.getRange(2, 1, 1, 3).setBackground('#F5F5F5');  // CLIENT_ID
  sheet.getRange(3, 1, 1, 3).setBackground('#F5F5F5');  // CLIENT_SECRET
  sheet.getRange(4, 1, 1, 3).setBackground('#F5F5F5');  // ACCESS_TOKEN
  
  logOperation('基本設定シート初期化', 'success', 'シートを再構成しました');
}

// ===========================
// ログシート再構成
// ===========================
function resetLogsSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('ログシート再構成', 
    'ログシートを再構成しますか？\n既存のログは削除されます。', 
    ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    logOperation(
      'ログシート再構成',
      'info',
      '再構成開始'
    );
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ログ');
    
    // 既存シートを削除
    if (sheet) {
      ss.deleteSheet(sheet);
      Utilities.sleep(500); // 処理を分散
    }
    
    // 新しいシートを作成（シンプルに）
    const newSheet = ss.insertSheet('ログ');
    
    // ヘッダーのみ設定
    const headers = ['日時', '操作', 'ステータス', '詳細'];
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    newSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    newSheet.setFrozenRows(1);
    
    // 最小限の列幅設定
    newSheet.autoResizeColumns(1, headers.length);
    
    ui.alert('ログシートを再構成しました。');
    console.log('ログシート再構成完了');
    
  } catch (error) {
    console.error('ログシート再構成エラー:', error);
    ui.alert('エラーが発生しました: ' + error.message);
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
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
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
    .setFontColor('#333333')
    .setRanges([statusRange])
    .build();
  
  const infoRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('info')
    .setFontColor('#666666')
    .setRanges([statusRange])
    .build();
    
  const warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('warning')
    .setFontColor('#333333')
    .setRanges([statusRange])
    .build();
    
  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('error')
    .setFontColor('#FF0000')
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
      
      initializeKeywordSettingsSheet();
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
    KEYWORD_SHEET_NAME,
    REPLY_HISTORY_SHEET_NAME,
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
// パスワード定数（ハードコーディング）
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
  // CLIENT_ID（B2セル）を書式なしテキストに設定（念のため）
  sheet.getRange(2, 2).setNumberFormat('@');
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
// 全トリガー削除機能
// ===========================
function deleteAllTriggers() {
  const ui = SpreadsheetApp.getUi();

  // 確認ダイアログ
  const response = ui.alert(
    '全トリガー削除',
    '本当に全てのトリガーを削除しますか？\n\n' +
    '⚠️ 警告:\n' +
    '・予約投稿の自動実行が停止します\n' +
    '・リプライの自動取得・返信が停止します\n' +
    '・アクセストークンの自動更新が停止します\n\n' +
    '削除後は「トリガーを再設定」から再度設定できます。',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const triggers = ScriptApp.getProjectTriggers();

    if (triggers.length === 0) {
      ui.alert('情報', '削除するトリガーがありません。', ui.ButtonSet.OK);
      logOperation('全トリガー削除', 'info', '削除対象のトリガーなし');
      return;
    }

    const deletedFunctions = [];
    let errorCount = 0;

    triggers.forEach(trigger => {
      try {
        const handlerFunction = trigger.getHandlerFunction();
        ScriptApp.deleteTrigger(trigger);
        deletedFunctions.push(handlerFunction);
      } catch (e) {
        console.error('トリガー削除エラー:', e);
        errorCount++;
      }
    });

    // 結果メッセージ
    let message = `${deletedFunctions.length}件のトリガーを削除しました。`;
    if (errorCount > 0) {
      message += `\n\n${errorCount}件のトリガーで削除エラーが発生しました。`;
    }

    ui.alert('完了', message, ui.ButtonSet.OK);

    // ログ記録
    const logDetails = `削除: ${deletedFunctions.join(', ')}${errorCount > 0 ? ` (エラー: ${errorCount}件)` : ''}`;
    logOperation('全トリガー削除', 'success', logDetails);

  } catch (error) {
    ui.alert('エラー', `トリガー削除中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
    logError('deleteAllTriggers', error);
  }
}

