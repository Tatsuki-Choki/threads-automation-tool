// ReplyFunctions.gs - 自動返信関連機能

// ===========================
// 自動返信メイン処理
// ===========================
function checkAndReplyToComments() {
  try {
    // 認証情報の事前チェック
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    
    if (!accessToken || !userId) {
      console.error('checkAndReplyToComments: 認証情報が未設定です');
      console.error(`ACCESS_TOKEN: ${accessToken ? '設定済み' : '未設定'}`);
      console.error(`USER_ID: ${userId || '未設定'}`);
      
      // 設定が存在しない場合は早期リターン
      logOperation('自動返信チェック', 'error', '認証情報が未設定のためスキップ');
      return;
    }
    
    const lastCheckTime = PropertiesService.getScriptProperties()
      .getProperty('lastCommentCheck') || '0';
    
    const currentTime = Date.now();
    
    logOperation('自動返信チェック開始', 'info', `最終チェック: ${new Date(parseInt(lastCheckTime))}`);
    
    // 最近の投稿を取得
    const recentPosts = getRecentPosts();
    
    let totalReplies = 0;
    
    // 各投稿のコメントをチェック
    for (const post of recentPosts) {
      const replies = checkPostComments(post.id, lastCheckTime);
      totalReplies += replies;
    }
    
    // 最終チェック時刻を更新
    PropertiesService.getScriptProperties()
      .setProperty('lastCommentCheck', currentTime.toString());
    
    logOperation('自動返信チェック完了', 'success', `${totalReplies}件の返信を実行`);
    
  } catch (error) {
    logError('checkAndReplyToComments', error);
  }
}

// ===========================
// 最近の投稿取得
// ===========================
function getRecentPosts() {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  // デバッグログ
  console.log('getRecentPosts - accessToken:', accessToken ? '設定済み' : '未設定');
  console.log('getRecentPosts - userId:', userId || '未設定');
  
  if (!accessToken || !userId) {
    const errorDetails = [];
    if (!accessToken) errorDetails.push('ACCESS_TOKENが未設定');
    if (!userId) errorDetails.push('USER_IDが未設定');
    throw new Error(`認証情報が見つかりません: ${errorDetails.join(', ')}`);
  }
  
  try {
    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads?fields=id,text,timestamp&limit=25`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const result = JSON.parse(response.getContentText());
    
    if (result.data) {
      return result.data;
    } else {
      throw new Error(result.error?.message || '投稿取得失敗');
    }
    
  } catch (error) {
    logError('getRecentPosts', error);
    return [];
  }
}

// ===========================
// 投稿のコメントチェック
// ===========================
function checkPostComments(postId, lastCheckTime) {
  const accessToken = getConfig('ACCESS_TOKEN');
  let repliesCount = 0;
  
  try {
    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${postId}/replies?fields=id,text,username,timestamp,from`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const result = JSON.parse(response.getContentText());
    
    if (result.data) {
      for (const comment of result.data) {
        try {
          // コメントの必須フィールドをチェック
          if (!comment.id || !comment.text) {
            console.log('不完全なコメントデータをスキップ:', comment);
            continue;
          }
          
          // 新しいコメントのみ処理
          const commentTime = new Date(comment.timestamp).getTime();
          if (commentTime > parseInt(lastCheckTime)) {
            if (shouldReplyToComment(comment)) {
              const replyResult = sendAutoReply(comment, comment.id);
              if (replyResult) {
                repliesCount++;
              }
            }
          }
        } catch (commentError) {
          console.error(`コメント処理エラー (ID: ${comment.id || 'unknown'}):`, commentError);
          continue;
        }
      }
    }
    
  } catch (error) {
    logError('checkPostComments', error);
  }
  
  return repliesCount;
}

// ===========================
// 返信判定
// ===========================
function shouldReplyToComment(comment) {
  console.log(`shouldReplyToComment チェック開始: "${comment.text}"`);
  
  // 自分のコメントは除外
  const myUsername = getConfig('USERNAME');
  const commentUsername = comment.from?.username || comment.username;
  console.log(`コメントユーザー: ${commentUsername}, 自分: ${myUsername}`);
  
  if (commentUsername === myUsername) {
    console.log('自分のコメントのため除外');
    return false;
  }
  
  // 既に返信済みかチェック
  const commentUserId = comment.from?.id || comment.id;
  const alreadyReplied = hasAlreadyReplied(comment.id, commentUserId);
  console.log(`既に返信済み?: ${alreadyReplied}`);
  
  if (alreadyReplied) {
    console.log('既に返信済みのため除外');
    return false;
  }
  
  // キーワードマッチング
  const keywords = getActiveKeywords();
  const commentText = comment.text.toLowerCase();
  console.log(`キーワードマッチング: "${commentText}"`);
  console.log(`アクティブキーワード数: ${keywords.length}`);
  
  for (const keyword of keywords) {
    console.log(`キーワード "${keyword.keyword}" (${keyword.matchType}) をチェック中...`);
    
    if (matchesKeyword(commentText, keyword)) {
      console.log(`✓ キーワード "${keyword.keyword}" にマッチ！`);
      comment.matchedKeyword = keyword;
      return true;
    } else {
      console.log(`✗ キーワード "${keyword.keyword}" にマッチしない`);
    }
  }
  
  console.log('マッチするキーワードがないため対象外');
  return false;
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
  const data = sheet.getDataRange().getValues();
  const keywords = [];
  
  for (let i = 1; i < data.length; i++) {
    const [id, keyword, matchType, replyContent, enabled, priority, probability] = data[i];
    
    if (enabled === true || enabled === 'TRUE') {
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
  
  return keywords;
}

// ===========================
// 返信済みチェック
// ===========================
function hasAlreadyReplied(commentId, userId) {
  const today = new Date().toDateString();
  const cacheKey = `replied_${userId}_${today}`;
  
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  
  if (cachedData) {
    const repliedComments = JSON.parse(cachedData);
    return repliedComments.includes(commentId);
  }
  
  // 履歴シートもチェック
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const [timestamp, recordedCommentId, recordedUserId] = data[i];
    
    // 今日の同じユーザーへの返信をチェック
    if (recordedUserId === userId && 
        new Date(timestamp).toDateString() === today) {
      return true;
    }
  }
  
  return false;
}

// ===========================
// 自動返信送信
// ===========================
function sendAutoReply(comment, replyToId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!comment.matchedKeyword) {
    return false;
  }
  
  try {
    // 返信内容の生成
    let replyText = comment.matchedKeyword.replyContent;
    
    // 変数置換
    const username = comment.from?.username || comment.username || 'ユーザー';
    replyText = replyText.replace(/{username}/g, `@${username}`);
    replyText = replyText.replace(/{time}/g, new Date().toLocaleTimeString('ja-JP'));
    replyText = replyText.replace(/{date}/g, new Date().toLocaleDateString('ja-JP'));
    
    // 自動返信であることを示す
    // replyText = '[自動返信] ' + replyText;
    
    // 返信の作成
    const createResponse = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          media_type: 'TEXT',
          text: replyText,
          reply_to_id: replyToId
        }),
        muteHttpExceptions: true
      }
    );
    
    const createResult = JSON.parse(createResponse.getContentText());
    
    if (createResult.id) {
      // 返信の公開
      const publishResponse = fetchWithTracking(
        `${THREADS_API_BASE}/v1.0/${userId}/threads_publish`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify({
            creation_id: createResult.id
          }),
          muteHttpExceptions: true
        }
      );
      
      const publishResult = JSON.parse(publishResponse.getContentText());
      
      if (publishResult.id) {
        // 返信履歴に記録
        recordReply(comment, replyText);
        
        // キャッシュ更新
        const userId = comment.from?.id || comment.id;
        updateReplyCache(comment.id, userId);
        
        const username = comment.from?.username || comment.username || 'ユーザー';
        logOperation('自動返信送信', 'success', 
          `ユーザー: @${username}, キーワード: ${comment.matchedKeyword.keyword}`);
        
        return true;
      }
    }
    
  } catch (error) {
    logError('sendAutoReply', error);
  }
  
  return false;
}

// ===========================
// 返信履歴記録
// ===========================
function recordReply(comment, replyText) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
  
  const userId = comment.from?.id || comment.id || 'unknown';
  const matchedKeyword = comment.matchedKeyword?.keyword || '';
  
  sheet.appendRow([
    new Date(),
    comment.id,
    userId,
    comment.text,
    replyText,
    matchedKeyword
  ]);
  
  // 新しい行の背景色をクリアして文字色を黒に設定
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, 6);
  range.setBackground(null);
  range.setFontColor('#000000');
  
  // 最新100件のみ表示（ヘッダーを除くデータ行を100件に制限）
  try {
    const maxDataRows = 100;
    const totalDataRows = sheet.getLastRow() - 1; // ヘッダー除く
    if (totalDataRows > maxDataRows) {
      const rowsToDelete = totalDataRows - maxDataRows;
      // 古い行（ヘッダー直下）から削除
      sheet.deleteRows(2, rowsToDelete);
    }
  } catch (e) {
    console.warn('自動応答結果のトリミング中に警告:', e);
  }
}

// ===========================
// 返信キャッシュ更新
// ===========================
function updateReplyCache(commentId, userId) {
  const today = new Date().toDateString();
  const cacheKey = `replied_${userId}_${today}`;
  
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  
  let repliedComments = [];
  
  if (cachedData) {
    repliedComments = JSON.parse(cachedData);
  }
  
  repliedComments.push(commentId);
  
  // キャッシュに保存（24時間）
  cache.put(cacheKey, JSON.stringify(repliedComments), 86400);
}

// ===========================
// 手動実行
// ===========================
function manualReplyCheck() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '手動返信チェック',
    'コメントをチェックして自動返信を実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    checkAndReplyToComments();
    ui.alert('返信チェックを実行しました。詳細はログをご確認ください。');
  }
}

// ===========================
// キーワード設定シート再構成
// ===========================
function resetAutoReplyKeywordsSheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'キーワード設定シート再構成',
    '既存の入力値は保持したまま、レイアウトと機能を最新化します。\n\n' +
    '・列構成、検証、条件付き書式、デザインを更新\n' +
    '・B列「キーワード」は書式なしテキストに設定',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      rebuildKeywordSettingsSheetNonDestructive_();
      ui.alert('キーワード設定シートを再構成しました（既存データは保持されました）。');
    } catch (error) {
      console.error('キーワード設定シート再構成エラー:', error);
      ui.alert('エラーが発生しました: ' + error.message);
    }
  }
}

// ===========================
// キーワード設定シート初期化
// ===========================
function initializeKeywordSettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 非破壊方式へ移行
  const existingSheet = spreadsheet.getSheetByName('キーワード設定');
  if (existingSheet) {
    rebuildKeywordSettingsSheetNonDestructive_();
    return;
  }
  
  // 新しいシートを作成（初回のみサンプルを投入）
  const sheet = spreadsheet.insertSheet('キーワード設定');
  
  // ヘッダー行を設定
  const headers = [
    'ID',
    'キーワード',
    'マッチタイプ',
    '返信内容',
    '有効/無効',
    '優先度',
    '確率(%)'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  // B列（キーワード）を書式なしテキストへ
  sheet.getRange(1, 2, sheet.getMaxRows(), 1).setNumberFormat('@');
  
  // サンプルデータを追加
  const sampleData = [
    [1, '質問', '部分一致', 'ご質問ありがとうございます！{username}さん、詳細についてお答えします。', true, 1, 100],
    [2, '購入', '部分一致', '{username}さん、購入についてのお問い合わせですね。こちらからご案内します。', true, 2, 100],
    [3, '詳細', '部分一致', '{username}さん、詳細情報をお送りします。少々お待ちください。', true, 3, 100],
    [4, 'ありがとう', '部分一致', '{username}さん、こちらこそありがとうございます！', true, 4, 50],
    [5, 'ありがとう', '部分一致', '{username}さん、ありがとうございます！今日も良い一日を！', true, 4, 50],
    [6, 'おすすめ', '部分一致', '{username}さん、おすすめの商品をご紹介します。', true, 5, 100],
    [7, '^こんにちは$', '正規表現', 'こんにちは、{username}さん！何かお手伝いできることはありますか？', false, 6, 100],
    [8, 'サポート', '完全一致', '{username}さん、サポートチームがお手伝いします。', false, 7, 100],
    [9, '値段|価格|料金', '正規表現', '{username}さん、価格についてご案内します。', false, 8, 100]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 50);   // ID
  sheet.setColumnWidth(2, 150);  // キーワード
  sheet.setColumnWidth(3, 100);  // マッチタイプ
  sheet.setColumnWidth(4, 400);  // 返信内容
  sheet.setColumnWidth(5, 80);   // 有効/無効
  sheet.setColumnWidth(6, 60);   // 優先度
  sheet.setColumnWidth(7, 80);   // 確率(%)
  
  // データ検証を追加
  // マッチタイプ列
  const matchTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['完全一致', '部分一致', '正規表現'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, 1000, 1).setDataValidation(matchTypeRule);
  
  // 有効/無効列
  const enabledRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([true, false])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 1000, 1).setDataValidation(enabledRule);
  
  // 説明コメントを追加
  sheet.getRange('A1').setNote(
    '自動返信キーワードの設定シート\n\n' +
    'ID: 一意の識別番号\n' +
    'キーワード: 検索するキーワード\n' +
    'マッチタイプ: 完全一致、部分一致、正規表現\n' +
    '返信内容: 自動返信するメッセージ（{username}, {time}, {date}が使用可能）\n' +
    '有効/無効: TRUE/FALSEで有効/無効を設定\n' +
    '優先度: 数値が小さいほど優先度が高い\n' +
    '確率(%): 同じ優先度の中でこのキーワードが選ばれる確率（1-100）'
  );
  
  // 条件付き書式を追加（有効/無効の視覚化）
  const enabledRange = sheet.getRange(2, 5, 1000, 1);
  const enabledRule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TRUE')
    .setBackground('#FFFFFF')
    .setFontColor('#000000')
    .setRanges([enabledRange])
    .build();
  
  const enabledRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#FF0000')
    .setRanges([enabledRange])
    .build();
  
  sheet.setConditionalFormatRules([enabledRule1, enabledRule2]);
  
  // 1行目を固定
  sheet.setFrozenRows(1);
  
  logOperation('キーワード設定シート初期化', 'success', 'シートを再構成しました');
}

// ===========================
// キーワード設定シート: 非破壊再構成
// ===========================
function rebuildKeywordSettingsSheetNonDestructive_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'キーワード設定';
  const oldSheet = ss.getSheetByName(name);
  
  const desiredHeaders = ['ID', 'キーワード', 'マッチタイプ', '返信内容', '有効/無効', '優先度', '確率(%)'];
  
  if (!oldSheet) {
    // 既存がなければ通常初期化
    initializeKeywordSettingsSheet();
    return;
  }
  
  // 旧データ読み込み
  const oldValues = oldSheet.getDataRange().getValues();
  const oldHeaders = (oldValues.length > 0) ? oldValues[0].map(h => (h || '').toString()) : [];
  
  // 追加列（未知列）を検出
  const unknownHeaders = [];
  for (let c = 0; c < oldHeaders.length; c++) {
    const h = oldHeaders[c] || '';
    if (h && desiredHeaders.indexOf(h) === -1) {
      unknownHeaders.push(h);
    }
  }
  
  const newHeaders = desiredHeaders.concat(unknownHeaders);
  
  // TMPシートを作成
  const tmpName = `${name}_TMP`;
  const tmpSheet = ss.getSheetByName(tmpName) ? ss.getSheetByName(tmpName) : ss.insertSheet(tmpName);
  tmpSheet.clear();
  
  // ヘッダー設定
  tmpSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  tmpSheet.getRange(1, 1, 1, newHeaders.length)
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  tmpSheet.setFrozenRows(1);
  
  // 列幅（既存トンマナ）
  tmpSheet.setColumnWidth(1, 50);   // ID
  tmpSheet.setColumnWidth(2, 150);  // キーワード
  tmpSheet.setColumnWidth(3, 100);  // マッチタイプ
  tmpSheet.setColumnWidth(4, 400);  // 返信内容
  tmpSheet.setColumnWidth(5, 80);   // 有効/無効
  tmpSheet.setColumnWidth(6, 60);   // 優先度
  tmpSheet.setColumnWidth(7, 80);   // 確率(%)
  
  // B列（キーワード）を書式なしテキストへ
  tmpSheet.getRange(1, 2, tmpSheet.getMaxRows(), 1).setNumberFormat('@');
  
  // データを新レイアウトへ再配置
  const newRows = [];
  const oldHeaderIndexMap = {};
  for (let i = 0; i < oldHeaders.length; i++) {
    oldHeaderIndexMap[oldHeaders[i]] = i;
  }
  
  for (let r = 1; r < oldValues.length; r++) {
    const oldRow = oldValues[r];
    if (!oldRow || oldRow.length === 0) continue;
    const newRow = new Array(newHeaders.length).fill('');
    for (let c = 0; c < newHeaders.length; c++) {
      const header = newHeaders[c];
      if (oldHeaderIndexMap.hasOwnProperty(header)) {
        const oldIdx = oldHeaderIndexMap[header];
        newRow[c] = oldRow[oldIdx];
      }
    }
    newRows.push(newRow);
  }
  
  if (newRows.length > 0) {
    tmpSheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
  }
  
  // 検証再適用（マッチタイプ / 有効列）
  const lastRow = Math.max(2, tmpSheet.getLastRow());
  const matchTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['完全一致', '部分一致', '正規表現'])
    .setAllowInvalid(false)
    .build();
  tmpSheet.getRange(2, 3, Math.max(0, lastRow - 1), 1).setDataValidation(matchTypeRule);
  
  const enabledRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([true, false])
    .setAllowInvalid(false)
    .build();
  tmpSheet.getRange(2, 5, Math.max(0, lastRow - 1), 1).setDataValidation(enabledRule);
  
  // 条件付き書式（有効/無効）
  const enabledRange = tmpSheet.getRange(2, 5, 1000, 1);
  const enabledRule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TRUE')
    .setBackground('#FFFFFF')
    .setFontColor('#000000')
    .setRanges([enabledRange])
    .build();
  const enabledRule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#FF0000')
    .setRanges([enabledRange])
    .build();
  tmpSheet.setConditionalFormatRules([enabledRule1, enabledRule2]);
  
  // 説明ノート
  tmpSheet.getRange('A1').setNote(
    '自動返信キーワードの設定シート\n\n' +
    'ID: 一意の識別番号\n' +
    'キーワード: 検索するキーワード（B列は書式なしテキスト）\n' +
    'マッチタイプ: 完全一致、部分一致、正規表現\n' +
    '返信内容: 自動返信するメッセージ（{username}, {time}, {date}が使用可能）\n' +
    '有効/無効: TRUE/FALSEで有効/無効を設定\n' +
    '優先度: 数値が小さいほど優先度が高い\n' +
    '確率(%): 同じ優先度の中でこのキーワードが選ばれる確率（1-100）'
  );
  
  // 旧→新の切り替え
  ss.deleteSheet(oldSheet);
  tmpSheet.setName(name);
  
  logOperation('キーワード設定シート再構成（非破壊）', 'success', `行数: ${newRows.length}`);
}

// ===========================
// 詳細デバッグ機能
// ===========================
function debugAutoReplyIssue() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('===== 自動返信デバッグ開始 =====');
    
    // 1. 認証情報チェック
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    const username = getConfig('USERNAME');
    
    console.log('認証情報:');
    console.log(`- ACCESS_TOKEN: ${accessToken ? '設定済み' : '未設定'}`);
    console.log(`- USER_ID: ${userId || '未設定'}`);
    console.log(`- USERNAME: ${username || '未設定'}`);
    
    // 2. キーワード設定チェック
    const keywords = getActiveKeywords();
    console.log('\nアクティブなキーワード:');
    keywords.forEach((keyword, index) => {
      console.log(`${index + 1}. ID:${keyword.id}, キーワード:"${keyword.keyword}", タイプ:${keyword.matchType}, 優先度:${keyword.priority}`);
    });
    
    // 3. Repliesシートの最新リプライをチェック
    const repliesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
    const repliesData = repliesSheet.getDataRange().getValues();
    
    console.log('\n最新のリプライ (最大5件):');
    const recentReplies = repliesData.slice(-6, -1); // 最新5件（ヘッダー除く）
    
    recentReplies.forEach((reply, index) => {
      if (reply[0]) { // 空行でない場合
        const [fetchTime, replyId, postId, replyTime, replyUsername, replyText] = reply;
        console.log(`${index + 1}. ${replyUsername}: "${replyText}" (${replyTime})`);
        
        // このリプライが自動返信対象かチェック
        const mockComment = {
          id: replyId,
          text: replyText,
          from: { username: replyUsername, id: replyId },
          timestamp: replyTime
        };
        
        console.log(`   自分のコメント?: ${mockComment.from.username === username}`);
        
        // キーワードマッチングをテスト
        const commentText = mockComment.text.toLowerCase();
        let matched = false;
        
        for (const keyword of keywords) {
          if (matchesKeyword(commentText, keyword)) {
            console.log(`   ✓ キーワード "${keyword.keyword}" にマッチ`);
            matched = true;
            break;
          }
        }
        
        if (!matched) {
          console.log(`   ✗ マッチするキーワードなし`);
        }
        
        // 返信済みかチェック
        const alreadyReplied = hasAlreadyReplied(mockComment.id, mockComment.from.id);
        console.log(`   既に返信済み?: ${alreadyReplied}`);
        
        console.log(`   shouldReplyToComment結果: ${shouldReplyToComment(mockComment)}`);
        console.log('---');
      }
    });
    
    // 4. ReplyHistoryシートの確認
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動応答結果');
    const historyData = historySheet.getDataRange().getValues();
    
    console.log('\n最新の返信履歴 (最大3件):');
    const recentHistory = historyData.slice(-4, -1); // 最新3件（ヘッダー除く）
    
    recentHistory.forEach((history, index) => {
      if (history[0]) { // 空行でない場合
        const [timestamp, commentId, userId, username, originalText, replyText] = history;
        console.log(`${index + 1}. ${timestamp}: @${username} に返信 "${replyText}"`);
      }
    });
    
    // 5. 最終チェック時刻の確認
    const lastCheckTime = PropertiesService.getScriptProperties().getProperty('lastCommentCheck') || '0';
    console.log(`\n最終チェック時刻: ${new Date(parseInt(lastCheckTime))}`);
    
    console.log('===== 自動返信デバッグ完了 =====');
    
    ui.alert('デバッグ完了', 'デバッグ情報をコンソールに出力しました。\n詳細はスクリプトエディタのログをご確認ください。', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert('エラー', `デバッグ中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
  }
}

// ===========================
// 特定リプライの自動返信テスト
// ===========================
function testSpecificReply() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '特定リプライテスト',
    'テストするリプライの内容を入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const testText = response.getResponseText();
    
    console.log('===== 特定リプライテスト =====');
    console.log(`テスト文字列: "${testText}"`);
    
    const keywords = getActiveKeywords();
    const commentText = testText.toLowerCase();
    
    console.log(`小文字変換後: "${commentText}"`);
    
    let matched = false;
    
    for (const keyword of keywords) {
      console.log(`キーワード "${keyword.keyword}" (${keyword.matchType}) をテスト中...`);
      
      if (matchesKeyword(commentText, keyword)) {
        console.log(`✓ マッチしました！`);
        console.log(`返信内容: "${keyword.replyContent}"`);
        matched = true;
        break;
      } else {
        console.log(`✗ マッチしませんでした`);
      }
    }
    
    if (!matched) {
      console.log('⚠️ どのキーワードにもマッチしませんでした');
    }
    
    console.log('===== テスト完了 =====');
    
    ui.alert('テスト完了', 'テスト結果をコンソールに出力しました。\n詳細はスクリプトエディタのログをご確認ください。', ui.ButtonSet.OK);
  }
}

// ===========================
// 自動応答結果シート再構成
// ===========================
function resetReplyHistorySheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '自動応答結果シート再構成',
    '既存の入力値は保持したまま、レイアウトと機能を最新化します。',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '自動応答結果';
  const oldSheet = ss.getSheetByName(name);
  if (!oldSheet) {
    initializeReplyHistorySheet();
    return;
  }
  const desiredHeaders = ['送信日時','コメントID','ユーザーID','元のコメント','返信内容','マッチキーワード'];
  const oldValues = oldSheet.getDataRange().getValues();
  const oldHeaders = (oldValues.length > 0) ? oldValues[0].map(h => (h || '').toString()) : [];
  const unknownHeaders = [];
  for (let c = 0; c < oldHeaders.length; c++) {
    const h = oldHeaders[c] || '';
    if (h && desiredHeaders.indexOf(h) === -1) unknownHeaders.push(h);
  }
  const newHeaders = desiredHeaders.concat(unknownHeaders);
  
  const tmpName = `${name}_TMP`;
  const tmpSheet = ss.getSheetByName(tmpName) ? ss.getSheetByName(tmpName) : ss.insertSheet(tmpName);
  tmpSheet.clear();
  tmpSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  tmpSheet.getRange(1, 1, 1, newHeaders.length)
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  tmpSheet.setFrozenRows(1);
  tmpSheet.setColumnWidth(1, 150);
  tmpSheet.setColumnWidth(2, 150);
  tmpSheet.setColumnWidth(3, 150);
  tmpSheet.setColumnWidth(4, 350);
  tmpSheet.setColumnWidth(5, 350);
  tmpSheet.setColumnWidth(6, 120);
  tmpSheet.getRange(2, 1, tmpSheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  
  const oldHeaderIndexMap = {};
  for (let i = 0; i < oldHeaders.length; i++) oldHeaderIndexMap[oldHeaders[i]] = i;
  const newRows = [];
  for (let r = 1; r < oldValues.length; r++) {
    const oldRow = oldValues[r];
    if (!oldRow || oldRow.length === 0) continue;
    const newRow = new Array(newHeaders.length).fill('');
    for (let c = 0; c < newHeaders.length; c++) {
      const header = newHeaders[c];
      if (oldHeaderIndexMap.hasOwnProperty(header)) newRow[c] = oldRow[oldHeaderIndexMap[header]];
    }
    newRows.push(newRow);
  }
  if (newRows.length > 0) tmpSheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
  
  ss.deleteSheet(oldSheet);
  tmpSheet.setName(name);
  logOperation('自動応答結果シート再構成（非破壊）', 'success', `行数: ${newRows.length}`);
}

// ===========================
// 自動応答結果シート初期化
// ===========================
function initializeReplyHistorySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを削除
  let existingSheet = spreadsheet.getSheetByName('自動応答結果');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet('自動応答結果');
  
  // ヘッダー行を設定
  const headers = [
    '送信日時',
    'コメントID',
    'ユーザーID',
    '元のコメント',
    '返信内容',
    'マッチキーワード'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // 送信日時
  sheet.setColumnWidth(2, 150); // コメントID
  sheet.setColumnWidth(3, 150); // ユーザーID
  sheet.setColumnWidth(4, 350); // 元のコメント
  sheet.setColumnWidth(5, 350); // 返信内容
  sheet.setColumnWidth(6, 120); // マッチキーワード
  
  // 日付列のフォーマット
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  
  // 最上行を固定
  sheet.setFrozenRows(1);
  
  // 説明コメントを追加
  sheet.getRange('A1').setNote(
    '自動応答の履歴シート\n\n' +
    '送信日時: 自動返信を送信した日時\n' +
    'コメントID: 返信元のコメントID\n' +
    'ユーザーID: 返信先ユーザーのID\n' +
    '元のコメント: ユーザーが投稿したコメント内容\n' +
    '返信内容: 自動送信した返信内容\n' +
    'マッチキーワード: 自動返信をトリガーしたキーワード'
  );
  
  logOperation('自動応答結果シート初期化', 'success', 'シートを再構成しました');
}

// ===========================
// テスト機能
// ===========================
function testKeywordMatching() {
  const testCases = [
    { text: '詳細を教えてください', expectedMatch: true },
    { text: '購入したいです', expectedMatch: true },
    { text: '質問', expectedMatch: true },
    { text: 'こんにちは', expectedMatch: false }
  ];
  
  const keywords = getActiveKeywords();
  
  testCases.forEach(testCase => {
    let matched = false;
    
    for (const keyword of keywords) {
      if (matchesKeyword(testCase.text.toLowerCase(), keyword)) {
        matched = true;
        console.log(`"${testCase.text}" → キーワード "${keyword.keyword}" にマッチ`);
        break;
      }
    }
    
    if (!matched && testCase.expectedMatch) {
      console.log(`⚠️ "${testCase.text}" → マッチするキーワードがありません`);
    }
  });
}