// WebApp.js - Threads Webhook受信エンドポイント

// ===========================
// Webhook設定
// ===========================
const WEBHOOK_CONFIG = {
  // Meta for Developersで設定する検証トークン（環境変数として設定推奨）
  VERIFY_TOKEN: 'threads-gas-webhook-very-secret-token-12345',
  
  // アプリシークレット（署名検証用）
  APP_SECRET: null, // getConfig('APP_SECRET')で取得
  
  // サポートするイベントタイプ
  SUPPORTED_EVENTS: ['threads_replies', 'threads_mentions']
};

// ===========================
// メインWebhookエンドポイント
// ===========================
function doPost(e) {
  try {
    console.log('===== Webhook受信 =====');
    
    // リクエストヘッダーとボディを取得
    const headers = e.parameter;
    const postData = e.postData;
    
    if (!postData || !postData.contents) {
      console.error('リクエストボディが空です');
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Bad Request' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 署名検証
    const signature = headers['x-hub-signature-256'] || '';
    if (!verifyWebhookSignature(postData.contents, signature)) {
      console.error('署名検証に失敗しました');
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // JSONペイロードをパース
    const payload = JSON.parse(postData.contents);
    console.log('受信したペイロード:', JSON.stringify(payload, null, 2));
    
    // イベントタイプを確認
    if (!payload.entry || payload.entry.length === 0) {
      console.log('処理対象のイベントがありません');
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 各エントリーを処理
    payload.entry.forEach(entry => {
      processWebhookEntry(entry);
    });
    
    // 成功レスポンスを返す
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('Webhook処理エラー:', error);
    logError('doPost', error);
    
    // エラーでも200を返す（Metaへの再送信を防ぐため）
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================
// Webhook検証エンドポイント（初回設定時）
// ===========================
function doGet(e) {
  try {
    // Webhook検証リクエストかどうかを確認
    if (e.parameter['hub.mode'] && e.parameter['hub.verify_token'] && e.parameter['hub.challenge']) {
      console.log('===== Webhook検証リクエスト =====');
      
      const mode = e.parameter['hub.mode'];
      const token = e.parameter['hub.verify_token'];
      const challenge = e.parameter['hub.challenge'];
      
      console.log(`Mode: ${mode}, Token: ${token}, Challenge: ${challenge}`);
      
      // 検証モードかつトークンが一致する場合
      if (mode === 'subscribe' && token === WEBHOOK_CONFIG.VERIFY_TOKEN) {
        console.log('Webhook検証成功');
        
        // challengeをそのまま返す（プレーンテキストとして）
        return ContentService
          .createTextOutput(challenge)
          .setMimeType(ContentService.MimeType.TEXT);
      }
      
      console.error('Webhook検証失敗: トークンが一致しません');
      return ContentService
        .createTextOutput('Forbidden')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    
    // Webhook検証以外のリクエストの場合
    return ContentService
      .createTextOutput('Threads Webhook Endpoint is active')
      .setMimeType(ContentService.MimeType.TEXT);
    
  } catch (error) {
    console.error('Webhook検証エラー:', error);
    return ContentService
      .createTextOutput('Internal Server Error')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ===========================
// 署名検証
// ===========================
function verifyWebhookSignature(payload, signature) {
  try {
    // アプリシークレットを取得
    const appSecret = getConfig('APP_SECRET');
    if (!appSecret) {
      console.warn('APP_SECRETが設定されていません。署名検証をスキップします。');
      return true; // 開発環境では一時的にtrueを返す
    }
    
    // 期待される署名を計算
    const expectedSignature = 'sha256=' + Utilities.computeHmacSha256Signature(payload, appSecret)
      .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2))
      .join('');
    
    // 署名を比較
    const isValid = signature === expectedSignature;
    
    if (!isValid) {
      console.error('署名が一致しません');
      console.error(`Expected: ${expectedSignature}`);
      console.error(`Received: ${signature}`);
    }
    
    return isValid;
    
  } catch (error) {
    console.error('署名検証エラー:', error);
    return false;
  }
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
    
    const response = UrlFetchApp.fetch(url + '?' + buildQueryString(params), {
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
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Webhookログ');
    
    if (!sheet) {
      // Webhookログシートがない場合は作成
      createWebhookLogSheet();
      return recordWebhookEvent(eventType, data, matchedKeyword, status);
    }
    
    sheet.appendRow([
      new Date(),                    // 受信日時
      eventType,                     // イベントタイプ
      data.id,                       // ID
      data.username || 'unknown',    // ユーザー名
      data.text || '',               // テキスト
      matchedKeyword || '',          // マッチしたキーワード
      status,                        // ステータス
      JSON.stringify(data)           // 生データ
    ]);
    
  } catch (error) {
    console.error('Webhookイベント記録エラー:', error);
  }
}

// ===========================
// Webhookログシート作成
// ===========================
function createWebhookLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet('Webhookログ');
  
  // ヘッダー設定
  const headers = [
    '受信日時',
    'イベントタイプ',
    'ID',
    'ユーザー名',
    'テキスト',
    'マッチキーワード',
    'ステータス',
    '生データ'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅調整
  sheet.setColumnWidth(1, 150); // 受信日時
  sheet.setColumnWidth(2, 120); // イベントタイプ
  sheet.setColumnWidth(3, 200); // ID
  sheet.setColumnWidth(4, 150); // ユーザー名
  sheet.setColumnWidth(5, 300); // テキスト
  sheet.setColumnWidth(6, 150); // マッチキーワード
  sheet.setColumnWidth(7, 100); // ステータス
  sheet.setColumnWidth(8, 400); // 生データ
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
    const response = UrlFetchApp.fetch(url, {
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
  const verifyToken = WEBHOOK_CONFIG.VERIFY_TOKEN;
  
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