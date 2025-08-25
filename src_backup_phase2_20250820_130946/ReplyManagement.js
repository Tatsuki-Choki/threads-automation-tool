// ReplyManagement.js - Threads自動返信管理（新実装）

// ===========================
// メイン自動返信処理
// ===========================
function autoReplyToThreads() {
  try {
    const config = validateConfig();
    if (!config) return;
    
    console.log('===== 自動返信処理開始 =====');
    
    // 最終チェック時刻を取得
    const lastCheckTime = getLastCheckTime();
    const currentTime = Date.now();
    
    // 自分の投稿を取得
    const myPosts = getMyThreadsPosts(config);
    if (!myPosts || myPosts.length === 0) {
      console.log('投稿が見つかりません');
      return;
    }
    
    console.log(`${myPosts.length}件の投稿を確認中...`);
    
    let totalProcessed = 0;
    let totalReplied = 0;
    
    // 各投稿の返信を確認
    for (const post of myPosts) {
      const result = processPostReplies(post, config, lastCheckTime);
      totalProcessed += result.processed;
      totalReplied += result.replied;
    }
    
    // 最終チェック時刻を更新
    updateLastCheckTime(currentTime);
    
    console.log(`===== 自動返信処理完了: ${totalProcessed}件確認, ${totalReplied}件返信 =====`);
    logOperation('自動返信処理', 'success', `確認: ${totalProcessed}件, 返信: ${totalReplied}件`);
    
  } catch (error) {
    console.error('自動返信処理エラー:', error);
    logError('autoReplyToThreads', error);
  }
}

// ===========================
// 設定検証
// ===========================
function validateConfig() {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  const username = getConfig('USERNAME');
  
  if (!accessToken || !userId) {
    console.error('認証情報が設定されていません');
    logOperation('自動返信処理', 'error', '認証情報未設定');
    return null;
  }
  
  return { accessToken, userId, username };
}

// ===========================
// 自分の投稿を取得
// ===========================
function getMyThreadsPosts(config, limit = 25) {
  try {
    const url = `${THREADS_API_BASE}/v1.0/${config.userId}/threads`;
    const params = {
      fields: 'id,text,timestamp,media_type,reply_audience',
      limit: limit
    };
    
    console.log(`投稿を取得中: ${url}`);
    
    const response = fetchWithTracking(url + '?' + buildQueryString(params), {
      headers: {
        'Authorization': `Bearer ${config.accessToken}`
      },
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.error) {
      console.error('投稿取得エラー:', result.error);
      throw new Error(result.error.message || 'API error');
    }
    
    console.log(`${result.data ? result.data.length : 0}件の投稿を取得`);
    return result.data || [];
    
  } catch (error) {
    logError('getMyThreadsPosts', error);
    return [];
  }
}

// ===========================
// 投稿の返信を処理
// ===========================
function processPostReplies(post, config, lastCheckTime) {
  let processed = 0;
  let replied = 0;
  
  try {
    // 返信を取得（ページネーション対応）
    const replies = getAllReplies(post.id, config);
    
    for (const reply of replies) {
      // 新しい返信のみ処理
      const replyTime = new Date(reply.timestamp).getTime();
      if (replyTime <= lastCheckTime) continue;
      
      processed++;
      
      // 自分の返信は除外
      if (reply.username === config.username) continue;
      
      // 既に返信済みかチェック
      if (hasAlreadyRepliedToday(reply.id, reply.from?.id || reply.username)) continue;
      
      // キーワードマッチング
      const matchedKeyword = findMatchingKeyword(reply.text);
      if (matchedKeyword) {
        const success = sendAutoReply(reply.id, reply, matchedKeyword, config);
        if (success) replied++;
      }
    }
    
  } catch (error) {
    console.error(`投稿 ${post.id} の返信処理エラー:`, error);
  }
  
  return { processed, replied };
}

// ===========================
// 全返信を取得（ページネーション対応）
// ===========================
function getAllReplies(postId, config) {
  const allReplies = [];
  let hasNext = true;
  let after = null;
  
  console.log(`投稿 ${postId} の返信を取得中...`);
  
  while (hasNext) {
    try {
      // Threads APIのドキュメントに基づいた正しいエンドポイント
      const url = `${THREADS_API_BASE}/v1.0/${postId}/replies`;
      const params = {
        fields: 'id,text,username,timestamp,from,media_type,media_url,has_replies',
        limit: 25
      };
      
      if (after) params.after = after;
      
      console.log(`リクエストURL: ${url}?${buildQueryString(params)}`);
      
      const response = fetchWithTracking(url + '?' + buildQueryString(params), {
        headers: {
          'Authorization': `Bearer ${config.accessToken}`
        },
        muteHttpExceptions: true
      });
      
      const responseText = response.getContentText();
      console.log(`レスポンスステータス: ${response.getResponseCode()}`);
      
      const result = JSON.parse(responseText);
      
      if (result.error) {
        console.error('返信取得エラー詳細:', JSON.stringify(result.error, null, 2));
        // 権限エラーの場合は空配列を返す
        if (result.error.code === 190 || result.error.code === 100) {
          console.log('権限エラーのため返信取得をスキップします');
          return [];
        }
        break;
      }
      
      if (result.data && result.data.length > 0) {
        console.log(`${result.data.length}件の返信を取得`);
        allReplies.push(...result.data);
      } else {
        console.log('返信データなし');
      }
      
      // ページネーション確認
      if (result.paging && result.paging.next) {
        after = result.paging.cursors.after;
        console.log('次のページがあります');
      } else {
        hasNext = false;
      }
      
    } catch (error) {
      console.error('返信取得例外エラー:', error.toString());
      hasNext = false;
    }
  }
  
  console.log(`合計 ${allReplies.length} 件の返信を取得しました`);
  return allReplies;
}

// ===========================
// キーワードマッチング
// ===========================
function matchesKeyword(text, keyword) {
  const keywordText = keyword.keyword.toLowerCase();
  
  switch (keyword.matchType) {
    case '完全一致':
    case 'exact':
      return text === keywordText;
      
    case '部分一致':
    case 'partial':
      return text.includes(keywordText);
      
    case '正規表現':
    case 'regex':
      try {
        const regex = new RegExp(keyword.keyword, 'i');
        return regex.test(text);
      } catch (error) {
        logError('matchesKeyword', `無効な正規表現: ${keyword.keyword}`);
        return false;
      }
      
    default:
      return false;
  }
}

// ===========================
// アクティブなキーワード取得
// ===========================
function getActiveKeywords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('キーワード設定');
  if (!sheet) {
    console.error('キーワード設定シートが見つかりません');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const keywords = [];
  
  console.log('キーワード設定シートのデータ行数:', data.length);
  
  for (let i = 1; i < data.length; i++) {
    const [id, keyword, matchType, replyContent, enabled, priority, probability] = data[i];
    
    console.log(`行${i}: キーワード="${keyword}", 有効="${enabled}"`);
    
    if (enabled === true || enabled === 'TRUE' || enabled === 'はい') {
      keywords.push({
        id: id,
        keyword: keyword,
        matchType: matchType,
        replyContent: replyContent,
        priority: priority || 999,
        probability: probability || 100
      });
    }
  }
  
  // 優先度でソート
  keywords.sort((a, b) => a.priority - b.priority);
  
  console.log(`アクティブなキーワード数: ${keywords.length}`);
  return keywords;
}

// ===========================
// 確率に基づいてキーワードを選択
// ===========================
function selectKeywordByProbability(keywords) {
  // 各キーワードの確率を取得（デフォルトは100）
  let totalProbability = 0;
  const keywordProbabilities = keywords.map(keyword => {
    const probability = keyword.probability || 100;
    totalProbability += probability;
    return { keyword, probability };
  });
  
  // 0からtotalProbabilityまでのランダムな数値を生成
  const random = Math.random() * totalProbability;
  
  // 累積確率でキーワードを選択
  let cumulativeProbability = 0;
  for (const { keyword, probability } of keywordProbabilities) {
    cumulativeProbability += probability;
    if (random < cumulativeProbability) {
      console.log(`確率選択: ${keyword.keyword} (確率: ${probability}/${totalProbability})`);
      return keyword;
    }
  }
  
  // 念のため最後のキーワードを返す（通常はここには来ない）
  return keywords[keywords.length - 1];
}

// ===========================
// キーワードマッチング
// ===========================
function findMatchingKeyword(text) {
  const keywords = getActiveKeywords();
  const lowerText = text.toLowerCase();
  
  console.log(`キーワードマッチング開始: "${text}"`);
  console.log(`アクティブキーワード数: ${keywords.length}`);
  
  // マッチするキーワードを収集
  const matchedKeywords = [];
  
  for (const keyword of keywords) {
    console.log(`キーワード "${keyword.keyword}" (${keyword.matchType}) をチェック中...`);
    if (matchesKeyword(lowerText, keyword)) {
      console.log(`✓ マッチしました！`);
      matchedKeywords.push(keyword);
    } else {
      console.log(`✗ マッチしませんでした`);
    }
  }
  
  if (matchedKeywords.length === 0) {
    return null;
  }
  
  // 1つしかマッチしない場合はそれを返す
  if (matchedKeywords.length === 1) {
    return matchedKeywords[0];
  }
  
  // 複数マッチする場合は、優先度でグループ化
  const groupedByPriority = {};
  matchedKeywords.forEach(keyword => {
    const priority = keyword.priority || 999;
    if (!groupedByPriority[priority]) {
      groupedByPriority[priority] = [];
    }
    groupedByPriority[priority].push(keyword);
  });
  
  // 最も高い優先度（数値が小さい）を取得
  const highestPriority = Math.min(...Object.keys(groupedByPriority).map(Number));
  const highestPriorityKeywords = groupedByPriority[highestPriority];
  
  // 同じ優先度のキーワードが1つの場合
  if (highestPriorityKeywords.length === 1) {
    return highestPriorityKeywords[0];
  }
  
  // 同じ優先度のキーワードが複数ある場合は確率で選択
  return selectKeywordByProbability(highestPriorityKeywords);
}

// ===========================
// 自動返信送信
// ===========================
function sendAutoReply(replyId, reply, keyword, config) {
  try {
    // 返信テキストを生成
    let replyText = keyword.replyContent;
    
    // 変数置換
    const username = reply.from?.username || reply.username || 'ユーザー';
    replyText = replyText.replace(/{username}/g, `@${username}`);
    replyText = replyText.replace(/{time}/g, new Date().toLocaleTimeString('ja-JP'));
    replyText = replyText.replace(/{date}/g, new Date().toLocaleDateString('ja-JP'));
    
    // 自動返信プレフィックス
    // replyText = '[自動返信] ' + replyText;
    
    console.log(`返信を作成中: @${username} への返信「${replyText}」`);
    
    // 返信作成
    const createUrl = `${THREADS_API_BASE}/v1.0/${config.userId}/threads`;
    const createResponse = fetchWithTracking(createUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        media_type: 'TEXT',
        text: replyText,
        reply_to_id: replyId  // リプライのIDに対して返信
      }),
      muteHttpExceptions: true
    });
    
    const createResult = JSON.parse(createResponse.getContentText());
    
    if (createResult.error) {
      throw new Error(createResult.error.message || 'Create error');
    }
    
    if (!createResult.id) {
      throw new Error('作成IDが返されませんでした');
    }
    
    // 返信公開
    const publishUrl = `${THREADS_API_BASE}/v1.0/${config.userId}/threads_publish`;
    const publishResponse = fetchWithTracking(publishUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        creation_id: createResult.id
      }),
      muteHttpExceptions: true
    });
    
    const publishResult = JSON.parse(publishResponse.getContentText());
    
    if (publishResult.error) {
      throw new Error(publishResult.error.message || 'Publish error');
    }
    
    // 返信履歴を記録
    recordAutoReply(reply, replyText, keyword.keyword);
    
    console.log(`自動返信送信成功: @${username} <- "${keyword.keyword}"`);
    return true;
    
  } catch (error) {
    console.error('自動返信送信エラー:', error);
    return false;
  }
}

// ===========================
// 返信履歴記録
// ===========================
function recordAutoReply(reply, replyText, matchedKeyword) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
    const username = reply.from?.username || reply.username || 'unknown';
    
    sheet.appendRow([
      new Date(),                    // 送信日時
      reply.id,                      // コメントID
      reply.from?.id || username,    // ユーザーID
      reply.text,                    // 元のコメント
      replyText,                     // 返信内容
      matchedKeyword                 // マッチしたキーワード
    ]);
    
    // 新しい行の背景色をクリアして文字色を黒に設定
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, 6);
    range.setBackground(null);
    range.setFontColor('#000000');
    
    // キャッシュも更新
    updateReplyCache(reply.id, reply.from?.id || username);
    
  } catch (error) {
    console.error('返信履歴記録エラー:', error);
  }
}

// ===========================
// 重複返信チェック
// ===========================
function hasAlreadyRepliedToday(replyId, userId) {
  const today = new Date().toDateString();
  const cacheKey = `replied_${userId}_${today}`;
  
  // キャッシュチェック
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  
  if (cachedData) {
    const repliedIds = JSON.parse(cachedData);
    return repliedIds.includes(replyId);
  }
  
  // シートチェック
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i > 0; i--) { // 最新から確認
    const [timestamp, , recordedUserId] = data[i];
    
    if (new Date(timestamp).toDateString() !== today) break;
    if (recordedUserId === userId) return true;
  }
  
  return false;
}

// ===========================
// 返信キャッシュ更新
// ===========================
function updateReplyCache(replyId, userId) {
  const today = new Date().toDateString();
  const cacheKey = `replied_${userId}_${today}`;
  
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  
  let repliedIds = [];
  if (cachedData) {
    repliedIds = JSON.parse(cachedData);
  }
  
  if (!repliedIds.includes(replyId)) {
    repliedIds.push(replyId);
  }
  
  // 24時間キャッシュ
  cache.put(cacheKey, JSON.stringify(repliedIds), 86400);
}

// ===========================
// 最終チェック時刻管理
// ===========================
function getLastCheckTime() {
  const stored = PropertiesService.getScriptProperties().getProperty('lastReplyCheck');
  return stored ? parseInt(stored) : Date.now() - (24 * 60 * 60 * 1000); // デフォルト24時間前
}

function updateLastCheckTime(timestamp) {
  PropertiesService.getScriptProperties().setProperty('lastReplyCheck', timestamp.toString());
}

// ===========================
// ユーティリティ関数
// ===========================
function buildQueryString(params) {
  return Object.keys(params)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
    .join('&');
}

// ===========================
// 手動実行用関数
// ===========================
function manualAutoReply() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '自動返信実行',
    '直近24時間以内の未処理リプライに対して自動返信を実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      console.log('===== 手動自動返信処理開始 =====');
      
      // 設定検証
      const config = validateConfig();
      if (!config) {
        ui.alert('エラー', '認証情報が設定されていません。', ui.ButtonSet.OK);
        return;
      }
      
      // 24時間前の日時を計算
      const oneDayAgo = new Date();
      oneDayAgo.setTime(oneDayAgo.getTime() - (24 * 60 * 60 * 1000));
      
      console.log(`対象期間: ${oneDayAgo.toLocaleString('ja-JP')} 〜 現在`);
      
      // 受信したリプライシートから未処理リプライを取得
      const result = processUnprocessedRepliesFromSheet(config, oneDayAgo);
      
      console.log(`===== 手動自動返信処理完了: ${result.processed}件処理, ${result.replied}件返信 =====`);
      
      ui.alert(
        '自動返信処理完了',
        `処理結果:\n\n` +
        `対象リプライ: ${result.processed}件\n` +
        `自動返信送信: ${result.replied}件\n` +
        `スキップ: ${result.skipped}件\n\n` +
        `詳細はログをご確認ください。`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      console.error('手動自動返信エラー:', error);
      ui.alert('エラー', `自動返信処理中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}

// ===========================
// 受信したリプライシートから未処理リプライを処理
// ===========================
function processUnprocessedRepliesFromSheet(config, startDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
  
  if (!sheet) {
    throw new Error('受信したリプライシートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const historyData = historySheet ? historySheet.getDataRange().getValues() : [];
  
  // 処理済みリプライIDのセットを作成
  const processedReplyIds = new Set();
  for (let i = 1; i < historyData.length; i++) {
    const replyId = historyData[i][1]; // コメントID列
    if (replyId) {
      // IDから先頭のシングルクォートを除去
      const cleanId = replyId.toString().replace(/^'/, '');
      processedReplyIds.add(cleanId);
    }
  }
  
  console.log(`処理済みリプライ数: ${processedReplyIds.size}`);
  console.log(`受信したリプライシートのデータ行数: ${data.length - 1}`);
  
  let processed = 0;
  let replied = 0;
  let skipped = 0;
  let skipReasons = {
    alreadyProcessed: 0,
    ownReply: 0,
    noKeywordMatch: 0,
    oldReply: 0
  };
  
  // データをリプライ日時で降順ソート（新しいものから処理）
  const sortedData = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // 空行でない場合
      sortedData.push(data[i]);
    }
  }
  console.log(`有効なリプライデータ数: ${sortedData.length}`);
  sortedData.sort((a, b) => new Date(b[3]) - new Date(a[3]));
  
  console.log(`24時間前の日時: ${startDate.toLocaleString('ja-JP')}`);
  console.log(`現在の日時: ${new Date()}`);
  
  // 最初の5件のデータをログ出力（デバッグ用）
  if (sortedData.length > 0) {
    console.log('最新のリプライデータ（最大5件）:');
    for (let i = 0; i < Math.min(5, sortedData.length); i++) {
      const [fetchTime, replyId, postId, replyTime, username, replyText] = sortedData[i];
      console.log(`  ${i+1}. @${username}: "${replyText?.substring(0, 30)}..." (${new Date(replyTime)})`);
    }
  }
  
  for (const row of sortedData) {
    const [fetchTime, replyId, postId, replyTime, username, replyText] = row;
    
    // リプライIDのクリーニング
    const cleanReplyId = replyId.toString().replace(/^'/, '');
    
    // 日付チェック
    const replyDate = new Date(replyTime);
    if (replyDate < startDate) {
      console.log(`古いリプライをスキップ: ${replyDate.toISOString()} < ${startDate.toISOString()}`);
      skipReasons.oldReply++;
      continue;
    }
    
    // 既に処理済みかチェック
    if (processedReplyIds.has(cleanReplyId)) {
      console.log(`既に処理済みのリプライをスキップ: ID=${cleanReplyId}`);
      skipped++;
      skipReasons.alreadyProcessed++;
      continue;
    }
    
    processed++;
    
    // 自分のリプライは除外
    if (username === config.username) {
      console.log(`自分のリプライをスキップ: @${username}`);
      skipped++;
      skipReasons.ownReply++;
      continue;
    }
    
    // リプライオブジェクトを構築
    const reply = {
      id: cleanReplyId,
      text: replyText,
      username: username,
      from: { username: username, id: cleanReplyId },
      timestamp: replyTime.toISOString()
    };
    
    console.log(`処理中: @${username} - "${replyText}"`);
    
    // キーワードマッチング
    const matchedKeyword = findMatchingKeyword(replyText);
    if (matchedKeyword) {
      console.log(`キーワード「${matchedKeyword.keyword}」にマッチ`);
      
      // 自動返信送信
      const success = sendAutoReply(cleanReplyId, reply, matchedKeyword, config);
      if (success) {
        replied++;
        console.log(`自動返信送信成功: @${username}`);
        // 処理済みセットに追加
        processedReplyIds.add(cleanReplyId);
      } else {
        console.log(`自動返信送信失敗: @${username}`);
      }
    } else {
      console.log(`マッチするキーワードなし: @${username}`);
      skipped++;
      skipReasons.noKeywordMatch++;
    }
    
    // API制限対策
    if (processed % 10 === 0) {
      Utilities.sleep(1000);
    }
  }
  
  console.log('\n===== スキップ理由の内訳 =====');
  console.log(`古いリプライ: ${skipReasons.oldReply}件`);
  console.log(`処理済み: ${skipReasons.alreadyProcessed}件`);
  console.log(`自分のリプライ: ${skipReasons.ownReply}件`);
  console.log(`キーワード不一致: ${skipReasons.noKeywordMatch}件`);
  console.log('================================\n');
  
  return { processed, replied, skipped };
}

// ===========================
// 統合実行用（リプライ取得＋自動返信）
// ===========================
function fetchAndAutoReply() {
  console.log('===== 統合処理開始 =====');
  
  // API制限チェック
  try {
    const quota = checkDailyAPIQuota();
    if (quota && quota.remaining < 1000) {
      console.log('API残り回数が少ないため、リプライ取得をスキップします');
      logOperation('リプライ取得', 'warning', `API残り回数: ${quota.remaining}`);
      return;
    }
  } catch (error) {
    console.error('API制限チェックエラー:', error);
  }
  
  // リプライ取得
  fetchAndSaveReplies();
  
  // 自動返信
  autoReplyToThreads();
  
  console.log('===== 統合処理完了 =====');
}

// ===========================
// デバッグ機能
// ===========================
function debugAutoReplySystem() {
  const ui = SpreadsheetApp.getUi();
  
  console.log('===== 自動返信システムデバッグ =====');
  
  // 1. 設定確認
  const config = validateConfig();
  if (!config) {
    console.error('❌ 設定検証失敗');
    ui.alert('エラー', '認証情報が設定されていません', ui.ButtonSet.OK);
    return;
  }
  console.log('✅ 設定検証成功');
  console.log(`- USER_ID: ${config.userId}`);
  console.log(`- USERNAME: ${config.username}`);
  
  // 2. API接続テスト
  try {
    const posts = getMyThreadsPosts(config, 1);
    console.log(`✅ API接続成功: ${posts.length}件の投稿を取得`);
    
    if (posts.length > 0) {
      console.log(`最新投稿: "${posts[0].text?.substring(0, 50)}..."`);
    }
  } catch (error) {
    console.error('❌ API接続失敗:', error);
    ui.alert('エラー', 'Threads APIへの接続に失敗しました', ui.ButtonSet.OK);
    return;
  }
  
  // 3. キーワード設定確認
  const keywords = getActiveKeywords();
  console.log(`✅ アクティブキーワード: ${keywords.length}件`);
  keywords.forEach((kw, index) => {
    console.log(`  ${index + 1}. "${kw.keyword}" (${kw.matchType})`);
  });
  
  // 4. 最終チェック時刻
  const lastCheck = getLastCheckTime();
  console.log(`✅ 最終チェック時刻: ${new Date(lastCheck)}`);
  
  // 5. 返信テスト
  const testReply = {
    id: 'test_' + Date.now(),
    text: 'これは質問のテストです',
    username: 'test_user',
    from: { id: 'test_user_id', username: 'test_user' },
    timestamp: new Date().toISOString()
  };
  
  const matchedKeyword = findMatchingKeyword(testReply.text);
  if (matchedKeyword) {
    console.log(`✅ キーワードマッチ: "${matchedKeyword.keyword}"`);
    console.log(`  返信内容: "${matchedKeyword.replyContent}"`);
  } else {
    console.log('⚠️ マッチするキーワードなし');
  }
  
  console.log('===== デバッグ完了 =====');
  
  ui.alert('デバッグ完了', 'コンソールログを確認してください', ui.ButtonSet.OK);
}

// ===========================
// 返信シミュレーション
// ===========================
function simulateAutoReply() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '返信シミュレーション',
    'テストするコメントを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const testText = response.getResponseText();
  if (!testText) return;
  
  console.log('===== 返信シミュレーション =====');
  console.log(`テストコメント: "${testText}"`);
  
  const matchedKeyword = findMatchingKeyword(testText);
  
  if (matchedKeyword) {
    console.log(`✅ マッチしたキーワード: "${matchedKeyword.keyword}"`);
    
    // 返信テキストを生成
    let replyText = matchedKeyword.replyContent;
    replyText = replyText.replace(/{username}/g, '@test_user');
    replyText = replyText.replace(/{time}/g, new Date().toLocaleTimeString('ja-JP'));
    replyText = replyText.replace(/{date}/g, new Date().toLocaleDateString('ja-JP'));
    // replyText = '[自動返信] ' + replyText;
    
    console.log(`生成された返信: "${replyText}"`);
    
    ui.alert(
      'シミュレーション結果',
      `マッチしたキーワード: ${matchedKeyword.keyword}\n` +
      `優先度: ${matchedKeyword.priority || 999}\n` +
      `確率: ${matchedKeyword.probability || 100}%\n\n` +
      `返信内容:\n${replyText}`,
      ui.ButtonSet.OK
    );
  } else {
    console.log('❌ マッチするキーワードがありません');
    
    // アクティブなキーワードを表示
    const keywords = getActiveKeywords();
    const keywordList = keywords.map(kw => `・${kw.keyword} (${kw.matchType})`).join('\n');
    
    ui.alert(
      'マッチなし',
      'マッチするキーワードがありません。\n\n' +
      '現在のアクティブキーワード:\n' + keywordList,
      ui.ButtonSet.OK
    );
  }
}

// ===========================
// 統計情報表示
// ===========================
function showAutoReplyStats() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    ui.alert('統計情報', '自動返信の履歴がありません。', ui.ButtonSet.OK);
    return;
  }
  
  // 統計を計算
  const stats = {
    total: data.length - 1,
    today: 0,
    week: 0,
    keywords: {},
    users: new Set()
  };
  
  const today = new Date().toDateString();
  const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  
  for (let i = 1; i < data.length; i++) {
    const [timestamp, , userId, , , keyword] = data[i];
    const date = new Date(timestamp);
    const username = userId; // ユーザーIDを使用
    
    if (date.toDateString() === today) stats.today++;
    if (date >= weekAgo) stats.week++;
    
    if (keyword) {
      stats.keywords[keyword] = (stats.keywords[keyword] || 0) + 1;
    }
    
    if (username) stats.users.add(username);
  }
  
  // レポート作成
  let report = '📊 自動返信統計\n\n';
  report += `総返信数: ${stats.total}件\n`;
  report += `今日の返信: ${stats.today}件\n`;
  report += `過去7日間: ${stats.week}件\n`;
  report += `ユニークユーザー: ${stats.users.size}人\n\n`;
  
  report += '🏷️ キーワード別返信数:\n';
  const sortedKeywords = Object.entries(stats.keywords)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  sortedKeywords.forEach(([keyword, count]) => {
    report += `  ${keyword}: ${count}件\n`;
  });
  
  ui.alert('自動返信統計', report, ui.ButtonSet.OK);
}