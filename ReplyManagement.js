// ReplyManagement.js - Threads自動返信管理（新実装）

// ===========================
// メイン自動返信処理
// ===========================
function autoReplyToThreads() {
  const functionStartTime = new Date();
  console.log(`===== [${functionStartTime.toISOString()}] 自動返信処理開始 =====`);
  console.log(`スプレッドシートID: ${SpreadsheetApp.getActiveSpreadsheet().getId()}`);
  console.log(`スプレッドシート名: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`);
  
  try {
    // 設定検証
    console.log('Step 1: 認証情報を検証中...');
    const config = RM_validateConfig();
    if (!config) {
      const errorMsg = '認証情報の検証に失敗しました';
      console.error(errorMsg);
      logOperation('自動返信処理', 'error', errorMsg);
      return;
    }
    console.log('✓ 認証情報の検証成功');
    
    console.log('===== 自動返信処理開始 =====');
    
    // 最終チェック時刻を取得
    console.log('Step 2: 最終チェック時刻を取得中...');
    const lastCheckTime = RM_getLastCheckTime();
    const currentTime = Date.now();
    console.log(`最終チェック時刻: ${new Date(lastCheckTime).toISOString()}`);
    console.log(`現在時刻: ${new Date(currentTime).toISOString()}`);
    
    // 自分の投稿を取得
    console.log('Step 3: 自分の投稿を取得中...');
    const myPosts = RM_getMyThreadsPosts(config);
    if (!myPosts || myPosts.length === 0) {
      const msg = '投稿が見つかりません';
      console.log(msg);
      logOperation('自動返信処理', 'info', msg);
      return;
    }
    
    console.log(`✓ ${myPosts.length}件の投稿を取得しました`);
    console.log(`Step 4: 各投稿の返信を確認中...`);
    
    let totalProcessed = 0;
    let totalReplied = 0;
    
    // 各投稿の返信を確認
    for (let i = 0; i < myPosts.length; i++) {
      const post = myPosts[i];
      console.log(`  [${i + 1}/${myPosts.length}] 投稿ID: ${post.id} を処理中...`);
      const result = RM_processPostReplies(post, config, lastCheckTime);
      totalProcessed += result.processed;
      totalReplied += result.replied;
      console.log(`  → 確認: ${result.processed}件, 返信: ${result.replied}件`);
    }
    
    // 最終チェック時刻を更新
    console.log('Step 5: 最終チェック時刻を更新中...');
    RM_updateLastCheckTime(currentTime);
    console.log('✓ 最終チェック時刻を更新しました');
    
    const functionEndTime = new Date();
    const elapsedSeconds = ((functionEndTime - functionStartTime) / 1000).toFixed(2);
    console.log(`===== 自動返信処理完了: ${totalProcessed}件確認, ${totalReplied}件返信 (処理時間: ${elapsedSeconds}秒) =====`);
    logOperation('自動返信処理', 'success', `確認: ${totalProcessed}件, 返信: ${totalReplied}件, 処理時間: ${elapsedSeconds}秒`);
    
  } catch (error) {
    console.error('自動返信処理エラー:', error);
    console.error('エラースタック:', error.stack);
    logError('autoReplyToThreads', error);
    logOperation('自動返信処理', 'error', `エラー: ${error.message}`);
  }
}

// ===========================
// バックフィル（過去N時間）自動返信
// ===========================
function RM_backfillAutoReply(hours = 6, updateBaseline = true) {
  try {
    const config = RM_validateConfig();
    if (!config) return { processed: 0, replied: 0 };

    const cutoffTime = Date.now() - (hours * 60 * 60 * 1000);
    const posts = RM_getMyThreadsPosts(config);

    console.log(`===== バックフィル開始: 過去${hours}時間 =====`);
    console.log(`${posts.length}件の投稿を確認中...`);

    let totalProcessed = 0;
    let totalReplied = 0;

    for (const post of posts) {
      const result = RM_processPostReplies(post, config, cutoffTime);
      totalProcessed += result.processed;
      totalReplied += result.replied;
    }

    if (updateBaseline) {
      RM_updateLastCheckTime(Date.now());
    }

    console.log(`===== バックフィル完了: 確認 ${totalProcessed}件, 返信 ${totalReplied}件 =====`);
    logOperation('バックフィル自動返信', 'success', `対象: 過去${hours}時間, 確認: ${totalProcessed}件, 返信: ${totalReplied}件`);

    return { processed: totalProcessed, replied: totalReplied };

  } catch (error) {
    console.error('バックフィル処理エラー:', error);
    logError('RM_backfillAutoReply', error);
    return { processed: 0, replied: 0 };
  }
}

// ===========================
// 設定検証
// ===========================
function RM_validateConfig() {
  console.log('  → 認証情報を取得中...');
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  const username = getConfig('USERNAME');
  
  console.log(`  → ACCESS_TOKEN: ${accessToken ? '設定済み (長さ: ' + accessToken.length + ')' : '未設定'}`);
  console.log(`  → USER_ID: ${userId || '未設定'}`);
  console.log(`  → USERNAME: ${username || '未設定'}`);
  
  const errors = [];
  if (!accessToken) errors.push('ACCESS_TOKEN');
  if (!userId) errors.push('USER_ID');
  
  if (errors.length > 0) {
    const errorMsg = `認証情報が設定されていません: ${errors.join(', ')}`;
    console.error(`  ✗ ${errorMsg}`);
    logOperation('自動返信処理', 'error', errorMsg);
    
    // スクリプトプロパティの全キーを確認
    try {
      const allProps = PropertiesService.getScriptProperties().getProperties();
      const propKeys = Object.keys(allProps);
      console.log(`  スクリプトプロパティに設定されているキー: ${propKeys.join(', ')}`);
    } catch (e) {
      console.error('  スクリプトプロパティの取得に失敗:', e.message);
    }
    
    return null;
  }
  
  console.log('  ✓ すべての認証情報が設定されています');
  return { accessToken, userId, username };
}

// ===========================
// 自分の投稿を取得
// ===========================
function RM_getMyThreadsPosts(config, limit = 25) {
  try {
    const url = `${THREADS_API_BASE}/v1.0/${config.userId}/threads`;
    const params = {
      fields: 'id,text,timestamp,media_type,reply_audience',
      limit: limit
    };
    
    console.log(`投稿を取得中: ${url}`);
    
    const response = fetchWithTracking(url + '?' + RM_buildQueryString(params), {
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
    logError('RM_getMyThreadsPosts', error);
    return [];
  }
}

// ===========================
// 投稿の返信を処理
// ===========================
function RM_processPostReplies(post, config, lastCheckTime) {
  let processed = 0;
  let replied = 0;
  
  try {
    // 返信を取得（ページネーション対応）
    const replies = RM_getAllReplies(post.id, config);
    
    for (const reply of replies) {
      // 新しい返信のみ処理
      const replyTime = new Date(reply.timestamp).getTime();
      if (replyTime <= lastCheckTime) continue;
      
      processed++;
      
      // 自分の返信は除外
      if (reply.username === config.username) continue;
      
      // 既に返信済みかチェック
      if (RM_hasAlreadyRepliedToday(reply.id, reply.from?.id || reply.username)) continue;
      
      // キーワードマッチング
      const matchedKeyword = RM_findMatchingKeyword(reply.text);
      if (matchedKeyword) {
        const success = RM_sendAutoReply(reply.id, reply, matchedKeyword, config);
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
function RM_getAllReplies(postId, config) {
  const allReplies = [];
  let hasNext = true;
  let after = null;
  let pageCount = 0;
  
  console.log(`  → 投稿 ${postId} の返信を取得中...`);
  
  while (hasNext) {
    pageCount++;
    try {
      // Threads APIのドキュメントに基づいた正しいエンドポイント
      const url = `${THREADS_API_BASE}/v1.0/${postId}/replies`;
      const params = {
        fields: 'id,text,username,timestamp,from,media_type,media_url,has_replies',
        limit: 25
      };
      
      if (after) params.after = after;
      
      console.log(`    → ページ${pageCount}: API呼び出し中...`);
      console.log(`    URL: ${url}?${RM_buildQueryString(params)}`);
      
      const response = fetchWithTracking(url + '?' + RM_buildQueryString(params), {
        headers: {
          'Authorization': `Bearer ${config.accessToken}`
        },
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      console.log(`    → HTTPステータス: ${responseCode}`);
      
      if (responseCode !== 200) {
        console.error(`    ✗ API呼び出し失敗 (HTTP ${responseCode})`);
        console.error(`    レスポンス: ${responseText.substring(0, 500)}`);
        logOperation('リプライ取得', 'error', `投稿${postId}: HTTP ${responseCode}`);
        hasNext = false;
        break;
      }
      
      let result;
      try {
        result = JSON.parse(responseText);
      } catch (parseError) {
        console.error(`    ✗ JSONパースエラー: ${parseError.message}`);
        console.error(`    レスポンステキスト: ${responseText.substring(0, 500)}`);
        logOperation('リプライ取得', 'error', `投稿${postId}: JSONパースエラー`);
        hasNext = false;
        break;
      }
      
      if (result.error) {
        const errorCode = result.error.code;
        const errorMsg = result.error.message;
        console.error(`    ✗ APIエラー (code: ${errorCode}): ${errorMsg}`);
        console.error(`    エラー詳細:`, JSON.stringify(result.error, null, 2));
        logOperation('リプライ取得', 'error', `投稿${postId}: ${errorMsg} (code: ${errorCode})`);
        
        // 権限エラーの場合は空配列を返す
        if (errorCode === 190 || errorCode === 100) {
          console.log('    権限エラーのため返信取得をスキップします');
          return [];
        }
        
        // その他のエラーはループを抜ける
        hasNext = false;
        break;
      }
      
      if (result.data && result.data.length > 0) {
        console.log(`    ✓ ${result.data.length}件の返信を取得`);
        allReplies.push(...result.data);
      } else {
        console.log(`    返信データなし (空配列)`);
      }
      
      // ページネーション確認
      if (result.paging && result.paging.next) {
        after = result.paging.cursors?.after;
        console.log(`    次のページあり (cursor: ${after})`);
      } else {
        console.log(`    これが最終ページです`);
        hasNext = false;
      }
      
    } catch (error) {
      console.error(`    ✗ 返信取得例外エラー (ページ${pageCount}):`, error.message);
      console.error(`    スタックトレース:`, error.stack);
      logError('RM_getAllReplies', error);
      logOperation('リプライ取得', 'error', `投稿${postId} ページ${pageCount}: ${error.message}`);
      hasNext = false;
    }
  }
  
  console.log(`  ✓ 合計 ${allReplies.length} 件の返信を取得しました (${pageCount}ページ)`);
  return allReplies;
}

// ===========================
// キーワードマッチング
// ===========================
function RM_matchesKeyword(text, keyword) {
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
        logError('RM_matchesKeyword', `無効な正規表現: ${keyword.keyword}`);
        return false;
      }
      
    default:
      return false;
  }
}

// ===========================
// アクティブなキーワード取得
// ===========================
function RM_getActiveKeywords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYWORD_SHEET_NAME);
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
function RM_selectKeywordByProbability(keywords) {
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
function RM_findMatchingKeyword(text) {
  const keywords = RM_getActiveKeywords();
  const lowerText = text.toLowerCase();
  
  console.log(`キーワードマッチング開始: "${text}"`);
  console.log(`アクティブキーワード数: ${keywords.length}`);
  
  // マッチするキーワードを収集
  const matchedKeywords = [];
  
  for (const keyword of keywords) {
    console.log(`キーワード "${keyword.keyword}" (${keyword.matchType}) をチェック中...`);
    if (RM_matchesKeyword(lowerText, keyword)) {
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
  return RM_selectKeywordByProbability(highestPriorityKeywords);
}

// ===========================
// 自動返信送信
// ===========================
function RM_sendAutoReply(replyId, reply, keyword, config) {
  const username = reply.from?.username || reply.username || 'ユーザー';
  
  try {
    console.log(`    → 自動返信を送信中: @${username} (リプライID: ${replyId})`);
    console.log(`    → マッチキーワード: "${keyword.keyword}"`);
    
    // 返信テキストを生成
    let replyText = keyword.replyContent;
    
    // 変数置換
    replyText = replyText.replace(/{username}/g, `@${username}`);
    replyText = replyText.replace(/{time}/g, new Date().toLocaleTimeString('ja-JP'));
    replyText = replyText.replace(/{date}/g, new Date().toLocaleDateString('ja-JP'));
    
    // 自動返信プレフィックス
    // replyText = '[自動返信] ' + replyText;
    
    console.log(`    → 返信内容: "${replyText}"`);
    
    // 返信作成
    console.log(`    → Step 1/2: Threads作成中...`);
    const createUrl = `${THREADS_API_BASE}/v1.0/${config.userId}/threads`;
    const createPayload = {
      media_type: 'TEXT',
      text: replyText,
      reply_to_id: replyId
    };
    
    const createResponse = fetchWithTracking(createUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(createPayload),
      muteHttpExceptions: true
    });
    
    const createResponseCode = createResponse.getResponseCode();
    const createResponseText = createResponse.getContentText();
    console.log(`    → 作成APIレスポンスコード: ${createResponseCode}`);
    
    if (createResponseCode !== 200) {
      console.error(`    ✗ 作成API失敗 (HTTP ${createResponseCode}): ${createResponseText}`);
      logOperation('自動返信送信', 'error', `作成失敗 @${username}: HTTP ${createResponseCode}`);
      return false;
    }
    
    const createResult = JSON.parse(createResponseText);
    
    if (createResult.error) {
      const error = createResult.error;
      const errorCode = error.code;
      const errorSubcode = error.error_subcode;
      const errorMsg = error.message || 'Unknown error';
      const errorUserMsg = error.error_user_msg || '';
      
      console.error(`    ✗ API作成エラー:`);
      console.error(`       Code: ${errorCode}, Subcode: ${errorSubcode}`);
      console.error(`       Message: ${errorMsg}`);
      console.error(`       User Message: ${errorUserMsg}`);
      
      // エラーの種類に応じて処理を分岐
      if (errorSubcode === 2207051) {
        // アクションがブロックされた（スパム判定）
        const blockMsg = `⚠️ アカウントがブロックされました: ${errorUserMsg}`;
        console.error(`    ${blockMsg}`);
        logOperation('自動返信送信', 'warning', `@${username}: アクションブロック (Subcode: 2207051)`);
        // この場合は処理を中断せず、次の返信に進む
        return false;
      } else if (errorSubcode === 4279009) {
        // メディアが見つからない（投稿が削除された可能性）
        const notFoundMsg = `⚠️ 返信先が見つかりません: ${errorUserMsg}`;
        console.error(`    ${notFoundMsg}`);
        logOperation('自動返信送信', 'warning', `@${username}: メディア不在 (Subcode: 4279009)`);
        // この場合も処理を中断せず、次の返信に進む
        return false;
      } else {
        // その他のエラー
        const generalErrorMsg = `作成エラー (Code: ${errorCode}, Subcode: ${errorSubcode}): ${errorMsg}`;
        console.error(`    ✗ ${generalErrorMsg}`);
        logOperation('自動返信送信', 'error', `@${username}: ${generalErrorMsg}`);
        throw new Error(generalErrorMsg);
      }
    }
    
    if (!createResult.id) {
      console.error('    ✗ 作成IDが返されませんでした');
      console.error('    レスポンス内容:', JSON.stringify(createResult));
      logOperation('自動返信送信', 'error', `@${username}: 作成IDなし`);
      throw new Error('作成IDが返されませんでした');
    }
    
    console.log(`    ✓ Threads作成成功 (creation_id: ${createResult.id})`);
    
    // 返信公開
    console.log(`    → Step 2/2: Threads公開中...`);
    const publishUrl = `${THREADS_API_BASE}/v1.0/${config.userId}/threads_publish`;
    const publishPayload = {
      creation_id: createResult.id
    };
    
    const publishResponse = fetchWithTracking(publishUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.accessToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(publishPayload),
      muteHttpExceptions: true
    });
    
    const publishResponseCode = publishResponse.getResponseCode();
    const publishResponseText = publishResponse.getContentText();
    console.log(`    → 公開APIレスポンスコード: ${publishResponseCode}`);
    
    if (publishResponseCode !== 200) {
      console.error(`    ✗ 公開API失敗 (HTTP ${publishResponseCode}): ${publishResponseText}`);
      logOperation('自動返信送信', 'error', `公開失敗 @${username}: HTTP ${publishResponseCode}`);
      return false;
    }
    
    const publishResult = JSON.parse(publishResponseText);
    
    if (publishResult.error) {
      const error = publishResult.error;
      const errorCode = error.code;
      const errorSubcode = error.error_subcode;
      const errorMsg = error.message || 'Unknown error';
      const errorUserMsg = error.error_user_msg || '';
      
      console.error(`    ✗ API公開エラー:`);
      console.error(`       Code: ${errorCode}, Subcode: ${errorSubcode}`);
      console.error(`       Message: ${errorMsg}`);
      console.error(`       User Message: ${errorUserMsg}`);
      
      // エラーの種類に応じて処理を分岐
      if (errorSubcode === 2207051) {
        // アクションがブロックされた（スパム判定）
        const blockMsg = `⚠️ アカウントがブロックされました: ${errorUserMsg}`;
        console.error(`    ${blockMsg}`);
        logOperation('自動返信送信', 'warning', `@${username}: アクションブロック (Subcode: 2207051)`);
        return false;
      } else if (errorSubcode === 4279009) {
        // メディアが見つからない
        const notFoundMsg = `⚠️ 返信先が見つかりません: ${errorUserMsg}`;
        console.error(`    ${notFoundMsg}`);
        logOperation('自動返信送信', 'warning', `@${username}: メディア不在 (Subcode: 4279009)`);
        return false;
      } else {
        // その他のエラー
        const generalErrorMsg = `公開エラー (Code: ${errorCode}, Subcode: ${errorSubcode}): ${errorMsg}`;
        console.error(`    ✗ ${generalErrorMsg}`);
        logOperation('自動返信送信', 'error', `@${username}: ${generalErrorMsg}`);
        throw new Error(generalErrorMsg);
      }
    }
    
    console.log(`    ✓ Threads公開成功 (投稿ID: ${publishResult.id})`);
    
    // 返信履歴を記録
    console.log(`    → 返信履歴を記録中...`);
    RM_recordAutoReply(reply, replyText, keyword.keyword);
    console.log(`    ✓ 返信履歴を記録しました`);
    
    console.log(`    ✓✓ 自動返信送信成功: @${username} <- "${keyword.keyword}"`);
    logOperation('自動返信送信', 'success', `@${username}: "${keyword.keyword}"`);
    return true;
    
  } catch (error) {
    console.error(`    ✗✗ 自動返信送信エラー (@${username}):`, error.message);
    console.error('    エラー詳細:', error.stack);
    logError('RM_sendAutoReply', error);
    logOperation('自動返信送信', 'error', `@${username}: ${error.message}`);
    return false;
  }
}

// ===========================
// 返信履歴記録
// ===========================
function RM_recordAutoReply(reply, replyText, matchedKeyword) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPLY_HISTORY_SHEET_NAME);
    if (!sheet) {
      // シートがない場合は自動作成
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(REPLY_HISTORY_SHEET_NAME);
      sheet.appendRow(['送信日時','コメントID','ユーザーID','元のコメント','返信内容','マッチキーワード']);
      sheet.setFrozenRows(1);
    }
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
    RM_updateReplyCache(reply.id, reply.from?.id || username);
    
    // 自動クリーンアップ: 件数制限チェック
    const currentDataRows = sheet.getLastRow() - 1; // ヘッダーを除く
    if (currentDataRows > REPLY_HISTORY_MAX_ENTRIES) {
      const rowsToDelete = currentDataRows - REPLY_HISTORY_MAX_ENTRIES;
      sheet.deleteRows(2, rowsToDelete); // 古いデータから削除
      console.log(`自動応答結果: ${rowsToDelete}行を件数制限により削除`);
    }
    
  } catch (error) {
    console.error('返信履歴記録エラー:', error);
  }
}

// ===========================
// 重複返信チェック
// ===========================
function RM_hasAlreadyRepliedToday(replyId, userId) {
  try {
    const today = new Date().toDateString();
    const cacheKey = `replied_${userId}_${today}`;
    
    console.log(`      → 重複チェック: ユーザーID=${userId}, リプライID=${replyId}`);
    
    // キャッシュチェック
    try {
      const cache = CacheService.getScriptCache();
      const cachedData = cache.get(cacheKey);
      
      if (cachedData) {
        const repliedIds = JSON.parse(cachedData);
        const isInCache = repliedIds.includes(replyId);
        console.log(`      → キャッシュ: ${isInCache ? '既に返信済み' : '未返信'} (キャッシュ内件数: ${repliedIds.length})`);
        if (isInCache) {
          return true;
        }
      } else {
        console.log(`      → キャッシュ: なし`);
      }
    } catch (cacheError) {
      console.error(`      ✗ キャッシュチェックエラー: ${cacheError.message}`);
      // キャッシュエラーでも処理を続行
    }
    
    // シートチェック
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPLY_HISTORY_SHEET_NAME);
      if (!sheet) {
        console.log(`      → シート「${REPLY_HISTORY_SHEET_NAME}」が存在しません`);
        return false;
      }
      
      const data = sheet.getDataRange().getValues();
      console.log(`      → シートチェック: 全${data.length - 1}行を確認中...`);
      
      let todayCount = 0;
      let matchFound = false;
      
      // 最新から確認
      for (let i = data.length - 1; i > 0; i--) {
        const [timestamp, recordedReplyId, recordedUserId] = data[i];
        
        // 日付チェック
        const recordDate = new Date(timestamp).toDateString();
        if (recordDate !== today) {
          console.log(`      → 今日以前のデータに到達 (${data.length - i - 1}行確認)`);
          break;
        }
        
        todayCount++;
        
        // ユーザーIDまたはリプライIDで確認
        if (recordedUserId === userId || recordedReplyId === replyId) {
          console.log(`      ✓ 既に返信済み (行${i}: ユーザーID=${recordedUserId}, リプライID=${recordedReplyId})`);
          matchFound = true;
          break;
        }
      }
      
      if (!matchFound) {
        console.log(`      → シート: 未返信 (今日の記録: ${todayCount}件)`);
      }
      
      return matchFound;
      
    } catch (sheetError) {
      console.error(`      ✗ シートチェックエラー: ${sheetError.message}`);
      // エラー時は安全側に倒して「返信済み」として扱わない（二重返信の可能性はあるが、未返信よりはマシ）
      return false;
    }
    
  } catch (error) {
    console.error(`      ✗✗ 重複チェック全体エラー: ${error.message}`);
    console.error(`      スタック: ${error.stack}`);
    // 安全側に倒して未返信として扱う
    return false;
  }
}

// ===========================
// 返信キャッシュ更新
// ===========================
function RM_updateReplyCache(replyId, userId) {
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
function RM_getLastCheckTime() {
  const stored = PropertiesService.getScriptProperties().getProperty('lastReplyCheck');
  return stored ? parseInt(stored) : Date.now() - (24 * 60 * 60 * 1000); // デフォルト24時間前
}

function RM_updateLastCheckTime(timestamp) {
  PropertiesService.getScriptProperties().setProperty('lastReplyCheck', timestamp.toString());
}

// ===========================
// ユーティリティ関数
// ===========================
function RM_buildQueryString(params) {
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
    '自動返信処理を実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    autoReplyToThreads();
    ui.alert('自動返信処理が完了しました。\n詳細はログをご確認ください。');
  }
}

// ===========================
// 手動実行用（過去6時間のバックフィル）
// ===========================
function manualBackfill6Hours() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '過去6時間を再処理',
    '過去6時間の未返信リプライに自動返信を行いますか？',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    const result = RM_backfillAutoReply(6, true);
    ui.alert('バックフィル完了', `確認: ${result.processed}件\n返信: ${result.replied}件`, ui.ButtonSet.OK);
  }
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
  const config = RM_validateConfig();
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
    const posts = RM_getMyThreadsPosts(config, 1);
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
  const keywords = RM_getActiveKeywords();
  console.log(`✅ アクティブキーワード: ${keywords.length}件`);
  keywords.forEach((kw, index) => {
    console.log(`  ${index + 1}. "${kw.keyword}" (${kw.matchType})`);
  });
  
  // 4. 最終チェック時刻
  const lastCheck = RM_getLastCheckTime();
  console.log(`✅ 最終チェック時刻: ${new Date(lastCheck)}`);
  
  // 5. 返信テスト
  const testReply = {
    id: 'test_' + Date.now(),
    text: 'これは質問のテストです',
    username: 'test_user',
    from: { id: 'test_user_id', username: 'test_user' },
    timestamp: new Date().toISOString()
  };
  
  const matchedKeyword = RM_findMatchingKeyword(testReply.text);
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
  
  const matchedKeyword = RM_findMatchingKeyword(testText);
  
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
    const keywords = RM_getActiveKeywords();
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPLY_HISTORY_SHEET_NAME);
  if (!sheet) {
    ui.alert('統計情報', '自動応答結果シートがありません。', ui.ButtonSet.OK);
    return;
  }
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

// ===========================
// 認証情報デバッグ（現在のスプレッドシート用）
// ===========================
function debugConfigForSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  
  console.log('===== 認証情報デバッグ開始 =====');
  console.log(`実行日時: ${new Date().toISOString()}`);
  console.log(`スプレッドシートID: ${SpreadsheetApp.getActiveSpreadsheet().getId()}`);
  console.log(`スプレッドシート名: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`);
  console.log(`スプレッドシートURL: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
  
  let report = '認証情報デバッグレポート\n\n';
  report += `スプレッドシート名: ${SpreadsheetApp.getActiveSpreadsheet().getName()}\n`;
  report += `スプレッドシートID: ${SpreadsheetApp.getActiveSpreadsheet().getId()}\n\n`;
  
  // スクリプトプロパティを確認
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    const propKeys = Object.keys(allProps);
    
    console.log(`スクリプトプロパティ数: ${propKeys.length}`);
    console.log(`キー一覧: ${propKeys.join(', ')}`);
    
    report += '【スクリプトプロパティ】\n';
    report += `設定されているキー数: ${propKeys.length}\n`;
    report += `キー: ${propKeys.join(', ')}\n\n`;
    
    // 主要な設定値を確認
    const accessToken = allProps['ACCESS_TOKEN'];
    const userId = allProps['USER_ID'];
    const username = allProps['USERNAME'];
    const clientId = allProps['CLIENT_ID'];
    const clientSecret = allProps['CLIENT_SECRET'];
    
    report += '【主要設定値】\n';
    report += `ACCESS_TOKEN: ${accessToken ? '設定済み (長さ: ' + accessToken.length + ')' : '未設定'}\n`;
    report += `USER_ID: ${userId || '未設定'}\n`;
    report += `USERNAME: ${username || '未設定'}\n`;
    report += `CLIENT_ID: ${clientId ? '設定済み' : '未設定'}\n`;
    report += `CLIENT_SECRET: ${clientSecret ? '設定済み' : '未設定'}\n\n`;
    
    console.log(`ACCESS_TOKEN: ${accessToken ? '設定済み (長さ: ' + accessToken.length + ')' : '未設定'}`);
    console.log(`USER_ID: ${userId || '未設定'}`);
    console.log(`USERNAME: ${username || '未設定'}`);
    
  } catch (error) {
    console.error('スクリプトプロパティ取得エラー:', error);
    report += `エラー: スクリプトプロパティの取得に失敗\n${error.message}\n\n`;
  }
  
  // シート構成を確認
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(s => s.getName());
    
    console.log(`シート数: ${sheets.length}`);
    console.log(`シート名: ${sheetNames.join(', ')}`);
    
    report += '【シート構成】\n';
    report += `シート数: ${sheets.length}\n`;
    report += `シート名:\n`;
    sheetNames.forEach(name => {
      report += `  - ${name}\n`;
    });
    report += '\n';
    
    // 必須シートの存在確認
    const requiredSheets = ['キーワード設定', '自動応答結果', '受信したリプライ', 'ログ'];
    report += '【必須シート確認】\n';
    requiredSheets.forEach(name => {
      const exists = sheetNames.includes(name);
      const status = exists ? '✓ 存在' : '✗ 不在';
      console.log(`${status}: ${name}`);
      report += `${status}: ${name}\n`;
    });
    report += '\n';
    
  } catch (error) {
    console.error('シート構成確認エラー:', error);
    report += `エラー: シート構成の確認に失敗\n${error.message}\n\n`;
  }
  
  // キーワード設定を確認
  try {
    const keywordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYWORD_SHEET_NAME);
    if (keywordSheet) {
      const data = keywordSheet.getDataRange().getValues();
      const keywords = [];
      
      for (let i = 1; i < data.length; i++) {
        const [id, keyword, matchType, replyContent, enabled] = data[i];
        if (enabled === true || enabled === 'TRUE') {
          keywords.push({ keyword, matchType });
        }
      }
      
      console.log(`アクティブキーワード数: ${keywords.length}`);
      report += '【キーワード設定】\n';
      report += `全キーワード数: ${data.length - 1}\n`;
      report += `アクティブキーワード数: ${keywords.length}\n`;
      
      if (keywords.length > 0) {
        report += 'アクティブキーワード:\n';
        keywords.forEach(kw => {
          report += `  - "${kw.keyword}" (${kw.matchType})\n`;
        });
      } else {
        report += '⚠️ アクティブなキーワードがありません\n';
      }
      report += '\n';
      
    } else {
      report += '【キーワード設定】\n';
      report += '✗ キーワード設定シートが存在しません\n\n';
    }
  } catch (error) {
    console.error('キーワード設定確認エラー:', error);
    report += `エラー: キーワード設定の確認に失敗\n${error.message}\n\n`;
  }
  
  // トリガー情報を確認
  try {
    const triggers = ScriptApp.getProjectTriggers();
    console.log(`トリガー数: ${triggers.length}`);
    
    report += '【トリガー設定】\n';
    report += `トリガー数: ${triggers.length}\n`;
    
    if (triggers.length > 0) {
      triggers.forEach(trigger => {
        const handlerName = trigger.getHandlerFunction();
        const eventType = trigger.getEventType();
        report += `  - ${handlerName} (${eventType})\n`;
        console.log(`トリガー: ${handlerName} (${eventType})`);
      });
    } else {
      report += '⚠️ トリガーが設定されていません\n';
    }
    report += '\n';
    
  } catch (error) {
    console.error('トリガー情報確認エラー:', error);
    report += `エラー: トリガー情報の確認に失敗\n${error.message}\n\n`;
  }
  
  console.log('===== 認証情報デバッグ完了 =====');
  
  ui.alert('認証情報デバッグ', report, ui.ButtonSet.OK);
}

// ===========================
// 自動返信テスト実行（現在のスプレッドシート用）
// ===========================
function testAutoReplyForCurrentSheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '自動返信テスト',
    'このスプレッドシートで自動返信機能をテスト実行しますか？\n\n' +
    '※実際にThreads APIを呼び出しますが、返信は送信しません。',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  console.log('===== 自動返信テスト開始 =====');
  console.log(`スプレッドシート: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`);
  
  let report = '自動返信テスト結果\n\n';
  let allSuccess = true;
  
  // Step 1: 認証情報確認
  console.log('Step 1: 認証情報を確認中...');
  report += '【Step 1: 認証情報確認】\n';
  
  const config = RM_validateConfig();
  if (!config) {
    report += '✗ 失敗: 認証情報が設定されていません\n';
    report += '→ クイックセットアップを実行してください\n\n';
    allSuccess = false;
    ui.alert('テスト失敗', report, ui.ButtonSet.OK);
    return;
  }
  
  report += '✓ 認証情報の確認成功\n';
  report += `  USER_ID: ${config.userId}\n`;
  report += `  USERNAME: ${config.username}\n\n`;
  
  // Step 2: API接続テスト
  console.log('Step 2: Threads API接続テスト中...');
  report += '【Step 2: API接続テスト】\n';
  
  try {
    const posts = RM_getMyThreadsPosts(config, 1);
    if (posts && posts.length > 0) {
      report += `✓ API接続成功: ${posts.length}件の投稿を取得\n`;
      report += `  最新投稿ID: ${posts[0].id}\n\n`;
    } else {
      report += '✓ API接続成功（投稿なし）\n\n';
    }
  } catch (error) {
    report += `✗ API接続失敗: ${error.message}\n`;
    report += '→ ACCESS_TOKENが正しいか確認してください\n\n';
    allSuccess = false;
  }
  
  // Step 3: キーワード設定確認
  console.log('Step 3: キーワード設定を確認中...');
  report += '【Step 3: キーワード設定確認】\n';
  
  const keywords = RM_getActiveKeywords();
  if (keywords.length > 0) {
    report += `✓ アクティブキーワード: ${keywords.length}件\n`;
    keywords.slice(0, 3).forEach(kw => {
      report += `  - "${kw.keyword}" (${kw.matchType})\n`;
    });
    if (keywords.length > 3) {
      report += `  ... 他${keywords.length - 3}件\n`;
    }
    report += '\n';
  } else {
    report += '✗ アクティブなキーワードがありません\n';
    report += '→ キーワード設定シートでキーワードを有効化してください\n\n';
    allSuccess = false;
  }
  
  // Step 4: シート構成確認
  console.log('Step 4: シート構成を確認中...');
  report += '【Step 4: シート構成確認】\n';
  
  const requiredSheets = [
    { name: 'キーワード設定', constant: KEYWORD_SHEET_NAME },
    { name: '自動応答結果', constant: REPLY_HISTORY_SHEET_NAME }
  ];
  
  requiredSheets.forEach(({ name, constant }) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(constant);
    if (sheet) {
      report += `✓ ${name}シート: 存在\n`;
    } else {
      report += `✗ ${name}シート: 不在\n`;
      allSuccess = false;
    }
  });
  report += '\n';
  
  // 総合結果
  console.log('===== 自動返信テスト完了 =====');
  
  report += '【総合結果】\n';
  if (allSuccess) {
    report += '✓✓ すべてのテストに合格しました\n';
    report += '自動返信機能は正常に動作する準備ができています。\n';
  } else {
    report += '✗✗ 一部のテストに失敗しました\n';
    report += '上記のエラーを修正してください。\n';
  }
  
  ui.alert('テスト完了', report, ui.ButtonSet.OK);
}
