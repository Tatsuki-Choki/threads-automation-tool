// WebApp.js - Threads Webhook受信エンドポイント

// ===========================
// Webhook設定
// ===========================
// Webhook設定（Script Propertiesから取得）
function getWebhookConfig() {
  const scriptProps = PropertiesService.getScriptProperties();
  
  return {
    // Meta for Developersで設定する検証トークン
    VERIFY_TOKEN: scriptProps.getProperty('WEBHOOK_VERIFY_TOKEN'),
    
    // アプリシークレット（署名検証用）
    APP_SECRET: scriptProps.getProperty('APP_SECRET') || getConfig('APP_SECRET'),
    
    // サポートするイベントタイプ
    SUPPORTED_EVENTS: ['threads_replies', 'threads_mentions']
  };
}

// ===========================
// メインWebhookエンドポイント
// ===========================
function doPost(e) {
  const startTime = Date.now();
  
  try {
    logInfo('Webhook受信開始');
    
    // リクエストヘッダーとボディを取得
    const headers = e.parameter;
    const postData = e.postData;
    
    // ペイロードの検証
    if (!postData || !postData.contents) {
      logWarn('リクエストボディが空です');
      // 400 Bad Requestを返す
      return ContentService
        .createTextOutput(JSON.stringify({ 
          error: 'Bad Request',
          message: 'Request body is required'
        }))
        .setMimeType(ContentService.MimeType.JSON)
        .setResponseCode(400);
    }
    
    // 署名検証
    const signature = headers['x-hub-signature-256'] || '';
    if (!verifyWebhookSignature(postData.contents, signature)) {
      logWarn('署名検証失敗', {
        hasSignature: !!signature,
        ip: e.parameter['REMOTE_ADDR'] || 'unknown'
      });
      // 401 Unauthorizedを返す
      return ContentService
        .createTextOutput(JSON.stringify({ 
          error: 'Unauthorized',
          message: 'Invalid signature'
        }))
        .setMimeType(ContentService.MimeType.JSON)
        .setResponseCode(401);
    }
    
    // JSONペイロードをパース
    let payload;
    try {
      payload = JSON.parse(postData.contents);
    } catch (parseError) {
      logError('doPost', parseError, {
        contentType: postData.type,
        contentLength: postData.contents?.length
      });
      return ContentService
        .createTextOutput(JSON.stringify({ 
          error: 'Bad Request',
          message: 'Invalid JSON payload'
        }))
        .setMimeType(ContentService.MimeType.JSON)
        .setResponseCode(400);
    }
    
    // デバッグモードの場合のみペイロードをログ出力
    const config = getLogConfig();
    if (config.enableDebugMode) {
      logDebug('受信ペイロード', {
        payload: maskSensitiveData(payload)
      });
    }
    
    // イベントタイプを確認
    if (!payload.entry || !Array.isArray(payload.entry) || payload.entry.length === 0) {
      logInfo('処理対象のイベントがありません');
      return ContentService
        .createTextOutput(JSON.stringify({ 
          status: 'ok',
          message: 'No events to process'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 処理結果を記録
    const results = {
      processed: 0,
      failed: 0,
      errors: []
    };
    
    // 各エントリーを処理
    for (const entry of payload.entry) {
      try {
        processWebhookEntry(entry);
        results.processed++;
      } catch (entryError) {
        results.failed++;
        results.errors.push({
          entryId: entry.id,
          error: entryError.toString()
        });
        logError('processWebhookEntry', entryError, {
          entryId: entry.id
        });
      }
    }
    
    // 処理時間を記録
    const duration = Date.now() - startTime;
    
    // 処理結果をログ
    logInfo('Webhook処理完了', {
      duration: `${duration}ms`,
      processed: results.processed,
      failed: results.failed,
      total: payload.entry.length
    });
    
    // 部分的な失敗がある場合は警告
    if (results.failed > 0) {
      logWarn('一部のエントリー処理に失敗', {
        failed: results.failed,
        errors: results.errors
      });
    }
    
    // 成功レスポンスを返す（Metaへの再送信を防ぐため常に200）
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'ok',
        processed: results.processed,
        failed: results.failed
      }))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // 予期しないエラー
    const duration = Date.now() - startTime;
    
    logError('doPost', error, {
      duration: `${duration}ms`,
      type: 'unexpected_error'
    });
    
    // エラーの詳細は返さない（セキュリティ上の理由）
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error',
        message: 'Internal server error'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Webhook検証は今後doPostで処理するため、doGetは削除（Code.jsに統合）

// ===========================
// 署名検証
// ===========================
function verifyWebhookSignature(payload, signature) {
  try {
    // アプリシークレットを取得（Script Properties優先）
    const scriptProps = PropertiesService.getScriptProperties();
    const appSecret = scriptProps.getProperty('APP_SECRET') || getConfig('APP_SECRET');
    
    if (!appSecret) {
      console.error('❌ APP_SECRETが設定されていません。Webhook署名検証が無効です。');
      logOperation('Webhook署名検証', 'error', 'APP_SECRETが未設定のため検証をスキップ');
      // セキュリティ上、本番環境では必ずfalseを返すべき
      return false;
    }
    
    // 署名がない場合は拒否
    if (!signature) {
      console.error('❌ Webhook署名が提供されていません');
      return false;
    }
    
    // 期待される署名を計算
    const expectedSignature = 'sha256=' + Utilities.computeHmacSha256Signature(payload, appSecret)
      .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2))
      .join('');
    
    // タイミング攻撃を防ぐため、固定時間での比較を行う
    const isValid = secureCompare(signature, expectedSignature);
    
    if (!isValid) {
      console.error('❌ Webhook署名が一致しません');
      logOperation('Webhook署名検証', 'error', '署名不一致');
    }
    
    return isValid;
    
  } catch (error) {
    console.error('❌ 署名検証エラー:', error);
    logError('verifyWebhookSignature', error);
    return false;
  }
}

// タイミング攻撃を防ぐための安全な文字列比較
function secureCompare(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') {
    return false;
  }
  
  const lengthA = a.length;
  const lengthB = b.length;
  const maxLength = Math.max(lengthA, lengthB);
  
  let result = lengthA === lengthB ? 0 : 1;
  
  for (let i = 0; i < maxLength; i++) {
    const charA = i < lengthA ? a.charCodeAt(i) : 0;
    const charB = i < lengthB ? b.charCodeAt(i) : 0;
    result |= charA ^ charB;
  }
  
  return result === 0;
}

// ===========================
// Webhookエントリー処理
// ===========================
function processWebhookEntry(entry) {
  try {
    console.log(`エントリー処理開始: ID=${entry.id}, Time=${entry.time}`);
    
    // 変更内容を処理
    if (entry.changes && Array.isArray(entry.changes)) {
      entry.changes.forEach(change => {
        processWebhookChange(change);
      });
    }
    
    // メッセージングイベントを処理
    if (entry.messaging && Array.isArray(entry.messaging)) {
      entry.messaging.forEach(message => {
        processWebhookMessage(message);
      });
    }
    
  } catch (error) {
    console.error('エントリー処理エラー:', error);
    logError('processWebhookEntry', error);
  }
}

// ===========================
// 変更イベント処理
// ===========================
function processWebhookChange(change) {
  try {
    const { field, value } = change;
    
    console.log(`変更イベント: Field=${field}`);
    
    switch (field) {
      case 'replies':
        processReplyEvent(value);
        break;
        
      case 'mentions':
        processMentionEvent(value);
        break;
        
      default:
        console.log(`未対応のフィールド: ${field}`);
    }
    
  } catch (error) {
    console.error('変更イベント処理エラー:', error);
    logError('processWebhookChange', error);
  }
}

// ===========================
// リプライイベント処理
// ===========================
function processReplyEvent(value) {
  try {
    console.log('===== リプライイベント処理 =====');
    
    // リプライデータの検証
    if (!value || !value.id) {
      console.error('無効なリプライデータ');
      return;
    }
    
    const reply = {
      id: value.id,
      text: value.text || '',
      username: value.from?.username || 'unknown',
      userId: value.from?.id || null,
      timestamp: value.created_time || new Date().toISOString(),
      parentId: value.parent_id || null,
      mediaType: value.media_type || 'TEXT'
    };
    
    console.log('リプライ情報:', JSON.stringify(reply, null, 2));
    
    // 設定を検証
    const config = validateConfig();
    if (!config) {
      console.error('設定が無効です');
      return;
    }
    
    // 自分の返信は除外
    if (reply.username === config.username) {
      console.log('自分の返信のため処理をスキップ');
      return;
    }
    
    // 既に返信済みかチェック
    if (hasAlreadyRepliedToday(reply.id, reply.userId || reply.username)) {
      console.log('本日既に返信済みのため処理をスキップ');
      return;
    }
    
    // キーワードマッチング
    const matchedKeyword = findMatchingKeyword(reply.text);
    if (matchedKeyword) {
      console.log(`キーワードマッチ: "${matchedKeyword.keyword}"`);
      
      // 親投稿IDを取得（必要に応じてAPIで取得）
      const postId = reply.parentId || getParentPostId(reply.id, config);
      
      if (postId) {
        // 自動返信を送信
        const success = sendAutoReply(postId, reply, matchedKeyword, config);
        
        if (success) {
          console.log('自動返信送信成功');
          
          // Webhook処理履歴を記録
          recordWebhookEvent('reply', reply, matchedKeyword.keyword, 'success');
        } else {
          console.error('自動返信送信失敗');
          recordWebhookEvent('reply', reply, matchedKeyword.keyword, 'failed');
        }
      }
    } else {
      console.log('マッチするキーワードなし');
      recordWebhookEvent('reply', reply, null, 'no_match');
    }
    
  } catch (error) {
    console.error('リプライイベント処理エラー:', error);
    logError('processReplyEvent', error);
  }
}

// ===========================
// メンションイベント処理
// ===========================
function processMentionEvent(value) {
  try {
    console.log('===== メンションイベント処理 =====');
    
    // メンション処理のロジックを実装
    // 現在は基本的にリプライと同じ処理を行う
    processReplyEvent(value);
    
  } catch (error) {
    console.error('メンションイベント処理エラー:', error);
    logError('processMentionEvent', error);
  }
}

// ===========================
// 親投稿ID取得
// ===========================
function getParentPostId(replyId, config) {
  try {
    // Threads APIで親投稿IDを取得
    const url = `${THREADS_API_BASE}/v1.0/${replyId}`;
    const params = {
      fields: 'parent_id,root_post,replied_to'
    };
    
    const response = fetchWithTracking(url + '?' + buildQueryString(params), {
      headers: {
        'Authorization': `Bearer ${config.accessToken}`
      },
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.error) {
      console.error('親投稿ID取得エラー:', result.error.message);
      return null;
    }
    
    return result.parent_id || result.root_post || result.replied_to || null;
    
  } catch (error) {
    console.error('親投稿ID取得エラー:', error);
    return null;
  }
}

// ===========================
// Webhookイベント記録
// ===========================
function recordWebhookEvent(eventType, data, matchedKeyword, status) {
  // Webhookログシートは削除されたため、通常のログに記録
  try {
    const details = `Event: ${eventType}, User: ${data.username || 'unknown'}, Text: ${data.text || ''}, Matched: ${matchedKeyword || 'none'}`;
    logOperation('Webhook受信', status || 'success', details);
  } catch (error) {
    console.error('Webhookイベント記録エラー:', error);
  }
}



// ===========================
// メッセージイベント処理（将来の拡張用）
// ===========================
function processWebhookMessage(message) {
  try {
    console.log('メッセージイベント:', JSON.stringify(message, null, 2));
    
    // 将来的にDMなどのメッセージング機能が追加された場合の処理
    
  } catch (error) {
    console.error('メッセージイベント処理エラー:', error);
    logError('processWebhookMessage', error);
  }
}

// ===========================
// テスト用関数
// ===========================
function testWebhookEndpoint() {
  const ui = SpreadsheetApp.getUi();
  
  // テスト用のペイロード
  const testPayload = {
    entry: [{
      id: 'test_entry_001',
      time: Date.now(),
      changes: [{
        field: 'replies',
        value: {
          id: 'test_reply_001',
          text: 'これは質問のテストです',
          from: {
            id: 'test_user_001',
            username: 'test_user'
          },
          created_time: new Date().toISOString(),
          parent_id: 'test_post_001',
          media_type: 'TEXT'
        }
      }]
    }]
  };
  
  // doPost関数をシミュレート
  const e = {
    parameter: {
      'x-hub-signature-256': 'test_signature'
    },
    postData: {
      contents: JSON.stringify(testPayload),
      type: 'application/json'
    }
  };
  
  console.log('===== Webhookテスト開始 =====');
  
  try {
    const response = doPost(e);
    console.log('レスポンス:', response.getContent());
    
    ui.alert(
      'Webhookテスト完了',
      'テストが完了しました。\nコンソールログを確認してください。',
      ui.ButtonSet.OK
    );
  } catch (error) {
    console.error('テストエラー:', error);
    ui.alert(
      'エラー',
      'テスト中にエラーが発生しました:\n' + error.toString(),
      ui.ButtonSet.OK
    );
  }
  
  console.log('===== Webhookテスト終了 =====');
}

// ===========================
// Webhook URL取得ヘルパー
// ===========================
function getWebhookUrl() {
  const scriptId = ScriptApp.getScriptId();
  const deploymentId = getLatestDeploymentId();
  
  if (deploymentId) {
    return `https://script.google.com/macros/s/${deploymentId}/exec`;
  } else {
    return `https://script.google.com/macros/s/${scriptId}/exec`;
  }
}

// ===========================
// 最新デプロイメントID取得
// ===========================
function getLatestDeploymentId() {
  try {
    // 注意: この機能はGoogle Apps Script APIを有効にする必要があります
    const scriptId = ScriptApp.getScriptId();
    const token = ScriptApp.getOAuthToken();
    
    const url = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    const response = fetchWithTracking(url, {
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.deployments && result.deployments.length > 0) {
      // 最新のWebアプリデプロイメントを探す
      const webAppDeployment = result.deployments.find(d => 
        d.deploymentConfig && d.deploymentConfig.type === 'WEB_APP'
      );
      
      return webAppDeployment ? webAppDeployment.deploymentId : null;
    }
    
  } catch (error) {
    console.error('デプロイメントID取得エラー:', error);
  }
  
  return null;
}

// ===========================
// Webhook設定表示
// ===========================
function showWebhookSettings() {
  const ui = SpreadsheetApp.getUi();
  
  const webhookUrl = getWebhookUrl();
  const config = getWebhookConfig();
  const verifyToken = config.VERIFY_TOKEN;
  
  const message = `Webhook設定情報\n\n` +
    `1. Webhook URL:\n${webhookUrl}\n\n` +
    `2. 検証トークン:\n${verifyToken}\n\n` +
    `3. Meta for Developersでの設定手順:\n` +
    `- アプリダッシュボードで「Webhooks」を選択\n` +
    `- 「コールバックURL」に上記URLを入力\n` +
    `- 「検証トークン」に上記トークンを入力\n` +
    `- サブスクリプションフィールドで「threads」を選択\n\n` +
    `※ APP_SECRETは設定シートに保存してください`;
  
  ui.alert('Webhook設定情報', message, ui.ButtonSet.OK);
}

// ===========================
// UrlFetchApp呼び出しトラッキング
// ===========================

/**
 * UrlFetchApp.fetchのラッパー関数。呼び出し回数をカウントする。
 * @param {string} url 取得するURL
 * @param {object} params UrlFetchApp.fetchに渡すパラメータ
 * @return {GoogleAppsScript.URL_Fetch.HTTPResponse} UrlFetchApp.fetchのレスポンス
 */
function fetchWithTracking(url, params) {
  const properties = PropertiesService.getScriptProperties();
  const today = new Date().toLocaleDateString('ja-JP');
  
  const lastCallDate = properties.getProperty('URL_FETCH_LAST_CALL_DATE');
  let count = parseInt(properties.getProperty('URL_FETCH_COUNT') || '0', 10);
  
  if (lastCallDate !== today) {
    // 日付が変わっていれば、カウントをリセット
    count = 1;
  } else {
    // 同じ日なら、カウントをインクリメント
    count++;
  }
  
  properties.setProperty('URL_FETCH_COUNT', count.toString());
  properties.setProperty('URL_FETCH_LAST_CALL_DATE', today);
  
  console.log(`UrlFetchApp call #${count} for today.`);
  console.log(`Request URL: ${url}`);
  if (params && params.payload) {
    console.log(`Request payload: ${params.payload}`);
  }
  
  try {
    // 元のUrlFetchApp.fetchを実行
    const response = UrlFetchApp.fetch(url, params);
    const responseCode = response.getResponseCode();
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode >= 400) {
      const responseText = response.getContentText();
      console.error(`HTTP Error ${responseCode}: ${responseText}`);
    }
    
    return response;
  } catch (error) {
    console.error(`fetchWithTracking error for URL: ${url}`);
    console.error(`Error details: ${error.toString()}`);
    console.error(`Error stack: ${error.stack}`);
    throw error;
  }
}

/**
 * 本日のUrlFetchAppの呼び出し回数を取得する
 * @return {number} 本日の呼び出し回数
 */
function getTodaysUrlFetchCount() {
  const properties = PropertiesService.getScriptProperties();
  const today = new Date().toLocaleDateString('ja-JP');
  
  const lastCallDate = properties.getProperty('URL_FETCH_LAST_CALL_DATE');
  
  if (lastCallDate === today) {
    return parseInt(properties.getProperty('URL_FETCH_COUNT') || '0', 10);
  } else {
    // 今日はまだ呼び出しがない
    return 0;
  }
}

/**
 * （任意）エディタのメニューから手動で回数を確認するための関数
 */
function showUrlFetchCount() {
  const count = getTodaysUrlFetchCount();
  const ui = SpreadsheetApp.getUi();
  ui.alert('API呼び出し回数', `本日のUrlFetchApp呼び出し回数: ${count} 回`, ui.ButtonSet.OK);
}

/**
 * パスワード認証付きでAPI呼び出し回数を確認する関数
 */
function showUrlFetchCountWithAuth() {
  const ui = SpreadsheetApp.getUi();
  
  // パスワード入力を求める
  const response = ui.prompt(
    'パスワード確認',
    '基本設定を表示するためのパスワードを入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const inputPassword = response.getResponseText();
  const correctPassword = getConfig('ADMIN_PASSWORD') || 'tsukichiyo.inc@gmail.com';
  
  if (inputPassword !== correctPassword) {
    ui.alert('エラー', 'パスワードが正しくありません。', ui.ButtonSet.OK);
    return;
  }
  
  // パスワードが正しい場合、API呼び出し回数を表示
  const count = getTodaysUrlFetchCount();
  const properties = PropertiesService.getScriptProperties();
  const lastCallDate = properties.getProperty('URL_FETCH_LAST_CALL_DATE') || '未使用';
  
  const message = `📊 API呼び出し統計\n\n` +
    `本日のAPI呼び出し回数: ${count} 回\n` +
    `最終呼び出し日: ${lastCallDate}\n\n` +
    `※ Google Apps Scriptの制限:\n` +
    `- 1日あたり20,000回まで（無料アカウント）\n` +
    `- 1日あたり100,000回まで（Workspaceアカウント）`;
  
  ui.alert('API呼び出し回数', message, ui.ButtonSet.OK);
}