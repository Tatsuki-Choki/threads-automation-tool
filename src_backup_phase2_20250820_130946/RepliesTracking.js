// 受信したリプライTracking.js - リプライ取得・管理機能
// processUnprocessedRepliesFromSheet関数はReplyManagement.jsから参照

// ===========================
// 統合処理：リプライ取得と自動返信
// ===========================
function fetchRepliesAndAutoReply() {
  try {
    console.log('===== 統合処理開始 =====');
    
    // 認証情報の事前チェック
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    const username = getConfig('USERNAME');
    
    if (!accessToken || !userId) {
      console.error('fetchRepliesAndAutoReply: 認証情報が未設定です');
      logOperation('統合処理', 'error', '認証情報が未設定のためスキップ');
      return;
    }
    
    const config = { accessToken, userId, username };
    
    // 24時間前の日時を計算
    const oneDayAgo = new Date();
    oneDayAgo.setTime(oneDayAgo.getTime() - (24 * 60 * 60 * 1000));
    
    console.log(`対象期間: ${oneDayAgo.toLocaleString('ja-JP')} 〜 現在`);
    console.log('=================================');
    
    // 1. まずAPIから最新のリプライを取得してシートに保存
    console.log('\n【ステップ1/2】 APIから最新のリプライを取得中...');
    const startTime = Date.now();
    fetchAndSaveReplies();
    const fetchTime = Date.now() - startTime;
    console.log(`→ リプライ取得完了（${Math.round(fetchTime / 1000)}秒）`);
    
    // 2. シートから24時間以内の未処理リプライを取得して自動返信
    console.log('\n【ステップ2/2】 未処理リプライの自動返信処理中...');
    const replyStartTime = Date.now();
    const result = processUnprocessedRepliesFromSheet(config, oneDayAgo);
    const replyTime = Date.now() - replyStartTime;
    console.log(`→ 自動返信処理完了（${Math.round(replyTime / 1000)}秒）`);
    
    console.log('\n===== 統合処理完了 =====');
    console.log(`総処理時間: ${Math.round((fetchTime + replyTime) / 1000)}秒`);
    logOperation('統合処理', 'success', 
      `処理: ${result.processed}件, 自動返信: ${result.replied}件, スキップ: ${result.skipped}件`);
    
  } catch (error) {
    console.error('統合処理エラー:', error);
    logError('fetchRepliesAndAutoReply', error);
  }
}

// ===========================
// リプライ取得と自動返信統合処理
// ===========================
function fetchRepliesAndCheckAutoReply(lastCheckTime) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  let total受信したリプライ = 0;
  let auto受信したリプライ = 0;
  
  try {
    // 24時間以内の全ての投稿を取得（キャッシュ無効）
    const recentPosts = getMyRecentPosts(null, false);
    
    if (!recentPosts || recentPosts.length === 0) {
      logOperation('統合処理', 'info', '投稿が見つかりませんでした');
      return { total受信したリプライ: 0, auto受信したリプライ: 0 };
    }
    
    // 各投稿のリプライを取得し、同時に自動返信をチェック
    for (const post of recentPosts) {
      const result = fetchRepliesForPostAndAutoReply(post, lastCheckTime);
      total受信したリプライ += result.newCount;
      auto受信したリプライ += result.auto受信したリプライ;
    }
    
    return { total受信したリプライ, auto受信したリプライ };
    
  } catch (error) {
    logError('fetchRepliesAndCheckAutoReply', error);
    return { total受信したリプライ: 0, auto受信したリプライ: 0 };
  }
}

// ===========================
// リプライ取得メイン処理
// ===========================
function fetchAndSaveReplies() {
  try {
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    
    if (!accessToken || !userId) {
      console.error('fetchAndSaveReplies: 認証情報が未設定です');
      logOperation('リプライ取得', 'error', '認証情報が未設定のためスキップ');
      return;
    }
    
    logOperation('リプライ取得開始', 'info', '最近の投稿から24時間以内のリプライを取得します');
    
    // 24時間以内の全ての投稿を取得（キャッシュ無効）
    console.log('投稿を取得中...');
    const recentPosts = getMyRecentPosts(null, false);
    
    if (!recentPosts || recentPosts.length === 0) {
      logOperation('リプライ取得', 'info', '投稿が見つかりませんでした');
      return;
    }
    
    console.log(`\n📊 取得した投稿数: ${recentPosts.length}件`);
    console.log('各投稿のリプライを取得中...\n');
    
    let totalNew受信したリプライ = 0;
    let totalUpdated受信したリプライ = 0;
    let processedPosts = 0;
    
    // 各投稿のリプライを取得
    for (const post of recentPosts) {
      processedPosts++;
      console.log(`[${processedPosts}/${recentPosts.length}] 投稿のリプライを取得中...`);
      const result = fetchRepliesForPost(post);
      totalNew受信したリプライ += result.newCount;
      totalUpdated受信したリプライ += result.updatedCount;
    }
    
    console.log('\n📈 リプライ取得結果:');
    console.log(`  - 新規リプライ: ${totalNew受信したリプライ}件`);
    console.log(`  - 更新済み: ${totalUpdated受信したリプライ}件`);
    
    logOperation('リプライ取得完了', 'success', 
      `新規: ${totalNew受信したリプライ}件, 更新: ${totalUpdated受信したリプライ}件`);
    
  } catch (error) {
    logError('fetchAndSaveReplies', error);
  }
}

// ===========================
// 自分の最近の投稿を取得（ページネーション対応）
// ===========================
function getMyRecentPosts(limit = null, useCache = false) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  // キャッシュの確認（統合処理時は無効化）
  if (useCache) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `recent_posts_${limit || 'all'}`;
    const cachedData = cache.get(cacheKey);
    
    if (cachedData) {
      console.log('getMyRecentPosts: キャッシュから取得');
      return JSON.parse(cachedData);
    }
  }
  
  try {
    const allPosts = [];
    let hasNext = true;
    let after = null;
    let pageCount = 0;
    const pageSize = 25; // APIの1ページあたりの最大件数
    
    // 24時間前の日時を計算
    const oneDayAgo = new Date();
    oneDayAgo.setTime(oneDayAgo.getTime() - (24 * 60 * 60 * 1000));
    
    console.log(`getMyRecentPosts: 24時間以内の投稿を取得中...`);
    
    while (hasNext) {
      pageCount++;
      let url = `${THREADS_API_BASE}/v1.0/${userId}/threads?fields=id,text,timestamp,media_type,media_url,permalink&limit=${pageSize}`;
      if (after) {
        url += `&after=${after}`;
      }
      
      console.log(`ページ ${pageCount} を取得中...`);
      
      const response = fetchWithTracking(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.data && result.data.length > 0) {
        // 24時間以内の投稿のみをフィルタリング
        const filteredPosts = result.data.filter(post => {
          const postDate = new Date(post.timestamp);
          return postDate >= oneDayAgo;
        });
        
        allPosts.push(...filteredPosts);
        
        // 24時間より古い投稿に到達したら終了
        const oldestPost = result.data[result.data.length - 1];
        if (new Date(oldestPost.timestamp) < oneDayAgo) {
          console.log('24時間より古い投稿に到達したため取得を終了');
          hasNext = false;
        } else if (result.paging && result.paging.next) {
          after = result.paging.cursors.after;
        } else {
          hasNext = false;
        }
        
        // 制限数に達したら終了
        if (limit && allPosts.length >= limit) {
          hasNext = false;
        }
      } else {
        hasNext = false;
      }
    }
    
    console.log(`getMyRecentPosts: 合計 ${allPosts.length} 件の投稿を取得しました`);
    
    // キャッシュに保存（useCache時のみ）
    if (useCache) {
      const cache = CacheService.getScriptCache();
      const cacheKey = `recent_posts_${limit || 'all'}`;
      cache.put(cacheKey, JSON.stringify(allPosts), 300);
    }
    
    return limit ? allPosts.slice(0, limit) : allPosts;
    
  } catch (error) {
    console.error('getMyRecentPosts エラー:', error.toString());
    logError('getMyRecentPosts', error);
    
    // URLFetch制限エラーの場合は空配列を返す
    if (error.toString().includes('urlfetch') || error.toString().includes('サービス')) {
      logOperation('API制限', 'error', 'URLFetch制限に達しました。リプライ取得をスキップします。');
      return [];
    }
    
    return [];
  }
}

// ===========================
// 特定の投稿のリプライを取得し、自動返信をチェック（ページネーション対応）
// ===========================
function fetchRepliesForPostAndAutoReply(post, lastCheckTime) {
  const accessToken = getConfig('ACCESS_TOKEN');
  let newCount = 0;
  let auto受信したリプライ = 0;
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    
    let hasNext = true;
    let after = null;
    let pageCount = 0;
    
    while (hasNext) {
      pageCount++;
      let url = `${THREADS_API_BASE}/v1.0/${post.id}/replies?fields=id,text,username,timestamp,from&limit=25`;
      if (after) {
        url += `&after=${after}`;
      }
      
      const response = fetchWithTracking(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.data && result.data.length > 0) {
        for (const reply of result.data) {
          try {
            // 必須フィールドをチェック
            if (!reply.id || !reply.text) {
              console.log('不完全なリプライデータをスキップ:', reply);
              continue;
            }
            
            const replyTime = new Date(reply.timestamp);
            
            // 24時間以内のリプライのみ処理
            if (replyTime >= oneDayAgo) {
              const saveResult = saveReplyToSheet(reply, post, sheet);
              if (saveResult === 'new') {
                newCount++;
              }
              
              // 新しいリプライかつ自動返信対象かチェック
              if (replyTime.getTime() > parseInt(lastCheckTime)) {
                console.log(`新しいリプライ発見: ${reply.from?.username || reply.username} - "${reply.text}"`);
                console.log(`リプライ時刻: ${replyTime}, 最終チェック時刻: ${new Date(parseInt(lastCheckTime))}`);
                
                if (shouldReplyToComment(reply)) {
                  console.log(`自動返信対象: ${reply.from?.username || reply.username}`);
                  const replyResult = sendAutoReply(reply, reply.id);
                  if (replyResult) {
                    auto受信したリプライ++;
                    console.log(`自動返信送信成功: ${reply.from?.username || reply.username}`);
                  } else {
                    console.log(`自動返信送信失敗: ${reply.from?.username || reply.username}`);
                  }
                } else {
                  console.log(`自動返信対象外: ${reply.from?.username || reply.username}`);
                }
              } else {
                console.log(`古いリプライ: ${reply.from?.username || reply.username} - "${reply.text}"`);
                console.log(`リプライ時刻: ${replyTime}, 最終チェック時刻: ${new Date(parseInt(lastCheckTime))}`);
              }
            }
          } catch (replyError) {
            console.error(`リプライ処理エラー (ID: ${reply.id || 'unknown'}):`, replyError);
            continue;
          }
        }
        
        // ページネーション確認
        if (result.paging && result.paging.next) {
          after = result.paging.cursors.after;
        } else {
          hasNext = false;
        }
      } else {
        hasNext = false;
      }
    }
    
  } catch (error) {
    console.error(`投稿 ${post.id} のリプライ取得エラー:`, error);
  }
  
  return { newCount, auto受信したリプライ };
}

// ===========================
// 特定の投稿のリプライを取得（ページネーション対応）
// ===========================
function fetchRepliesForPost(post) {
  const accessToken = getConfig('ACCESS_TOKEN');
  let newCount = 0;
  let updatedCount = 0;
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    
    let hasNext = true;
    let after = null;
    let pageCount = 0;
    
    console.log(`投稿 ${post.id} のリプライを取得中...`);
    
    while (hasNext) {
      pageCount++;
      let url = `${THREADS_API_BASE}/v1.0/${post.id}/replies?fields=id,text,username,timestamp,from&limit=25`;
      if (after) {
        url += `&after=${after}`;
      }
      
      const response = fetchWithTracking(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.data && result.data.length > 0) {
        console.log(`  ページ ${pageCount}: ${result.data.length} 件のリプライ`);
        
        for (const reply of result.data) {
          // 24時間以内のリプライのみ処理
          const replyTime = new Date(reply.timestamp);
          if (replyTime >= oneDayAgo) {
            const saveResult = saveReplyToSheet(reply, post, sheet);
            if (saveResult === 'new') {
              newCount++;
            }
          }
        }
        
        // ページネーション確認
        if (result.paging && result.paging.next) {
          after = result.paging.cursors.after;
        } else {
          hasNext = false;
        }
      } else {
        hasNext = false;
      }
    }
    
    if (newCount > 0) {
      console.log(`  → ${newCount} 件の新規リプライを保存`);
    }
    
  } catch (error) {
    console.error(`投稿 ${post.id} のリプライ取得エラー:`, error);
  }
  
  return { newCount, updatedCount };
}

// ===========================
// リプライをシートに保存
// ===========================
function saveReplyToSheet(reply, originalPost, sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
  }
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // 既存のリプライをチェック（重複防止） - 両方を文字列として比較
  const replyIdStr = reply.id.toString();
  for (let i = 1; i < data.length; i++) {
    const sheetIdStr = data[i][1].toString();
    if (sheetIdStr === replyIdStr) { // リプライIDが一致
      // 既に存在する場合はスキップ
      console.log(`リプライID: ${replyIdStr} は既に存在するためスキップします。`);
      return 'exists';
    }
  }
  
  // 新規リプライを追加
  const username = reply.from?.username || reply.username || 'unknown';
  
  // IDを文字列として保存するために、先頭にシングルクォートを付与
  const replyIdForSheet = "'" + reply.id;
  
  const newRow = sheet.appendRow([
    new Date(), // 取得日時
    replyIdForSheet, // リプライID (文字列として保存)
    originalPost.id, // 元投稿ID
    new Date(reply.timestamp), // リプライ日時
    username, // リプライユーザー名
    reply.text || '', // リプライ内容
    new Date(), // 最終更新日時
    '' // メモ欄
  ]);
  
  // 新しい行の背景色をクリアして文字色を黒に設定
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, 8);
  range.setBackground(null);
  range.setFontColor('#000000');
  
  return 'new';
}

// ===========================
// 特定期間のリプライ統計を取得
// ===========================
function getReplyStatistics(startDate, endDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
  const data = sheet.getDataRange().getValues();
  
  const stats = {
    total受信したリプライ: 0,
    uniqueUsers: new Set(),
    topRepliers: {},
    dailyCount: {},
    postReplyCounts: {} // 投稿ごとのリプライ数
  };
  
  const start = startDate ? new Date(startDate) : new Date(0);
  const end = endDate ? new Date(endDate) : new Date();
  
  for (let i = 1; i < data.length; i++) {
    const replyDate = new Date(data[i][3]); // リプライ日時（新しいカラム位置）
    
    if (replyDate >= start && replyDate <= end) {
      stats.total受信したリプライ++;
      
      const username = data[i][4]; // ユーザー名（新しいカラム位置）
      const postId = data[i][2]; // 元投稿ID
      
      stats.uniqueUsers.add(username);
      
      // トップリプライヤーを集計
      stats.topRepliers[username] = (stats.topRepliers[username] || 0) + 1;
      
      // 日別集計
      const dateKey = replyDate.toDateString();
      stats.dailyCount[dateKey] = (stats.dailyCount[dateKey] || 0) + 1;
      
      // 投稿ごとのリプライ数集計
      stats.postReplyCounts[postId] = (stats.postReplyCounts[postId] || 0) + 1;
    }
  }
  
  stats.uniqueUsersCount = stats.uniqueUsers.size;
  
  // トップリプライヤーをソート
  stats.topRepliersList = Object.entries(stats.topRepliers)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  // 最も多くリプライを受けた投稿
  stats.topPosts = Object.entries(stats.postReplyCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  return stats;
}

// ===========================
// リプライシートの初期化
// ===========================
function initializeRepliesSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを削除
  let existingSheet = spreadsheet.getSheetByName('受信したリプライ');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet('受信したリプライ');
  
  // ヘッダー行を設定（簡素化されたカラム）
  const headers = [
    '取得日時',
    'リプライID',
    '元投稿ID',
    'リプライ日時',
    'ユーザー名',
    'リプライ内容',
    '最終更新日時',
    'メモ'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // 取得日時
  sheet.setColumnWidth(2, 150); // リプライID
  sheet.setColumnWidth(3, 150); // 元投稿ID
  sheet.setColumnWidth(4, 150); // リプライ日時
  sheet.setColumnWidth(5, 120); // ユーザー名
  sheet.setColumnWidth(6, 400); // リプライ内容
  sheet.setColumnWidth(7, 150); // 最終更新日時
  sheet.setColumnWidth(8, 200); // メモ
  
  // 日付列のフォーマット
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange(2, 7, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  
  // 最上行を固定
  sheet.setFrozenRows(1);
  
  logOperation('受信したリプライシート初期化', 'success', 'シートを作成しました');
}

// ===========================
// 手動実行：リプライ取得
// ===========================
function manualFetchReplies() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'リプライ取得',
    '最近の投稿に対するリプライを取得しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    fetchAndSaveReplies();
    ui.alert('リプライの取得が完了しました。受信したリプライシートをご確認ください。');
  }
}

// ===========================
// 手動実行：統合処理
// ===========================
function manualFetchRepliesAndAutoReply() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '統合処理実行',
    'リプライ取得と自動返信を連続で実行しますか？\n\n' +
    '1. 最近7日間のリプライを取得\n' +
    '2. 自動返信の送信',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    fetchRepliesAndAutoReply();
    ui.alert('統合処理が完了しました。詳細はログをご確認ください。');
  }
}

// ===========================
// 統計レポート生成
// ===========================
function generateReplyReport() {
  const ui = SpreadsheetApp.getUi();
  
  // 期間を選択
  const response = ui.prompt(
    '統計期間',
    '何日前からの統計を表示しますか？（例: 7）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const days = parseInt(response.getResponseText()) || 7;
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);
  
  const stats = getReplyStatistics(startDate, new Date());
  
  // レポート作成
  let report = `📊 リプライ統計レポート（過去${days}日間）\n\n`;
  report += `総リプライ数: ${stats.total受信したリプライ}\n`;
  report += `ユニークユーザー数: ${stats.uniqueUsersCount}\n\n`;
  
  report += `🏆 トップリプライヤー:\n`;
  stats.topRepliersList.forEach((item, index) => {
    report += `${index + 1}. @${item[0]} - ${item[1]}回\n`;
  });
  
  if (stats.topPosts.length > 0) {
    report += `\n📌 最も多くリプライを受けた投稿:\n`;
    stats.topPosts.forEach((item, index) => {
      report += `${index + 1}. 投稿ID: ${item[0]} - ${item[1]}件のリプライ\n`;
    });
  }
  
  // HTMLで表示
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>リプライ統計レポート</h2>
      <pre style="background-color: #f5f5f5; padding: 15px; border-radius: 5px;">
${report}
      </pre>
      <button onclick="google.script.host.close()" style="margin-top: 20px; padding: 10px 20px;">
        閉じる
      </button>
    </div>
  `).setWidth(500).setHeight(600);
  
  ui.showModalDialog(htmlOutput, 'リプライ統計');
}

// ===========================
// 統合処理のテスト
// ===========================
function testIntegratedReplyAndAutoReply() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '統合処理テスト',
    '統合処理（リプライ取得＋自動返信）をテストしますか？\n\n' +
    'これにより実際のリプライ取得と自動返信が実行されます。',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      console.log('===== 統合処理テスト開始 =====');
      
      // 最終チェック時刻をリセット（テスト用）
      PropertiesService.getScriptProperties().setProperty('lastCommentCheck', '0');
      
      // 統合処理を実行
      fetchRepliesAndAutoReply();
      
      console.log('===== 統合処理テスト完了 =====');
      
      ui.alert('統合処理テストが完了しました。\n\n' +
        '詳細はログをご確認ください。\n' +
        '受信したリプライシートで取得したリプライと\n' +
        'ReplyHistoryシートで自動返信の履歴を確認できます。');
        
    } catch (error) {
      console.error('統合処理テストエラー:', error);
      ui.alert('統合処理テストでエラーが発生しました。\n\n' +
        'エラー: ' + error.message);
    }
  }
}

// ===========================
// リプライ取得デバッグ（改善版）
// ===========================
function debugFetchReplies() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('===== リプライ取得デバッグ開始 =====');
    
    // 1. 認証情報の確認
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    
    if (!accessToken || !userId) {
      ui.alert('エラー', '認証情報が設定されていません。基本設定を確認してください。', ui.ButtonSet.OK);
      return;
    }
    
    console.log('✅ 認証情報: OK');
    
    // 2. 24時間以内の投稿を取得してテスト
    console.log('\n📊 24時間以内の投稿を取得中...');
    const recentPosts = getMyRecentPosts(null, false);
    
    if (!recentPosts || recentPosts.length === 0) {
      ui.alert('エラー', '投稿が取得できませんでした。\n\n考えられる原因:\n- アクセストークンが無効\n- 投稿が存在しない', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`✅ 投稿取得成功: ${recentPosts.length}件`);
    
    // 3. 各投稿のリプライ数を集計
    console.log('\n📈 各投稿のリプライを確認中...');
    let totalReplies = 0;
    let postsWithReplies = 0;
    const sampleReplies = [];
    
    for (let i = 0; i < Math.min(5, recentPosts.length); i++) {
      const post = recentPosts[i];
      console.log(`\n投稿 ${i + 1}/${Math.min(5, recentPosts.length)}:`);
      console.log(`  ID: ${post.id}`);
      console.log(`  内容: ${post.text ? post.text.substring(0, 50) + '...' : '(なし)'}`);
      console.log(`  日時: ${new Date(post.timestamp).toLocaleString('ja-JP')}`);
      
      // リプライを取得
      let hasNext = true;
      let after = null;
      let postReplies = 0;
      
      while (hasNext && postReplies < 50) {  // 最大50件まで
        let url = `${THREADS_API_BASE}/v1.0/${post.id}/replies?fields=id,text,username,timestamp,from&limit=25`;
        if (after) url += `&after=${after}`;
        
        const response = fetchWithTracking(url, {
          headers: { 'Authorization': `Bearer ${accessToken}` },
          muteHttpExceptions: true
        });
        
        const result = JSON.parse(response.getContentText());
        
        if (result.data && result.data.length > 0) {
          postReplies += result.data.length;
          totalReplies += result.data.length;
          
          // サンプルとして最初の3件を保存
          if (sampleReplies.length < 3) {
            sampleReplies.push(...result.data.slice(0, 3 - sampleReplies.length));
          }
          
          if (result.paging && result.paging.next) {
            after = result.paging.cursors.after;
          } else {
            hasNext = false;
          }
        } else {
          hasNext = false;
        }
      }
      
      console.log(`  → リプライ数: ${postReplies}件`);
      if (postReplies > 0) postsWithReplies++;
    }
    
    // 4. 結果を表示
    const message = `📊 デバッグ結果:\n\n` +
      `✅ 24時間以内の投稿数: ${recentPosts.length}件\n` +
      `✅ リプライがある投稿: ${postsWithReplies}件\n` +
      `✅ 確認したリプライ総数: ${totalReplies}件\n\n`;
    
    let detailMessage = '';
    if (sampleReplies.length > 0) {
      detailMessage = '最新のリプライ例:\n' + 
        sampleReplies.map(r => 
          `- @${r.username || r.from?.username}: ${r.text ? r.text.substring(0, 30) + '...' : '(なし)'}`
        ).join('\n');
    } else {
      detailMessage = 'リプライが見つかりませんでした。';
    }
    
    // 5. 受信したリプライシートの確認
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
    if (sheet) {
      const sheetData = sheet.getDataRange().getValues();
      const sheetReplies = sheetData.length - 1;  // ヘッダー行を除く
      detailMessage += `\n\n📋 シート内の既存リプライ数: ${sheetReplies}件`;
    }
    
    ui.alert('デバッグ完了', message + detailMessage, ui.ButtonSet.OK);
    
    console.log('\n===== デバッグ完了 =====');
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert('エラー', `デバッグ中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// 注意: この関数は削除されました
// 単発処理のcheckAndReplyToPost関数を使用してください

// ===========================
// 監視用の自動返信送信
// ===========================
function sendAutoReplyForMonitoring(reply, originalPostId, matchedKeyword) {
  try {
    // キーワード設定シートから返信内容を取得
    const keywordSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('キーワード設定');
    
    if (!keywordSheet) {
      console.error('キーワード設定シートが見つかりません');
      return false;
    }
    
    const keywordData = keywordSheet.getDataRange().getValues();
    let replyContent = null;
    
    // マッチしたキーワードに対応する返信内容を検索
    for (let i = 1; i < keywordData.length; i++) {
      if (keywordData[i][0] === '有効' && 
          keywordData[i][1].toLowerCase() === matchedKeyword.toLowerCase()) {
        replyContent = keywordData[i][2];
        break;
      }
    }
    
    if (!replyContent) {
      console.log(`キーワード「${matchedKeyword}」に対応する返信内容が見つかりません`);
      return false;
    }
    
    // プレースホルダー置換
    const username = reply.username || reply.from?.username || 'ユーザー';
    const finalReplyContent = replyContent
      .replace(/{username}/g, username)
      .replace(/{date}/g, new Date().toLocaleDateString('ja-JP'))
      .replace(/{time}/g, new Date().toLocaleTimeString('ja-JP'));
    
    // 返信を送信
    const result = sendReply(reply.id, finalReplyContent);
    
    if (result) {
      // 履歴に記録
      saveAutoReplyHistory(
        reply.id,
        originalPostId,
        username,
        reply.text,
        matchedKeyword,
        finalReplyContent,
        'success'
      );
      
      console.log(`監視対象投稿への自動返信送信成功: @${username}`);
      return true;
    }
    
    return false;
    
  } catch (error) {
    console.error('監視用自動返信エラー:', error);
    logError('sendAutoReplyForMonitoring', error);
    return false;
  }
}