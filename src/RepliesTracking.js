// 受信したリプライTracking.js - リプライ取得・管理機能

// ===========================
// 統合処理：リプライ取得と自動返信
// ===========================
function fetchRepliesAndAutoReply() {
  try {
    console.log('===== 統合処理開始 =====');
    
    // 認証情報の事前チェック
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    
    if (!accessToken || !userId) {
      console.error('fetchRepliesAndAutoReply: 認証情報が未設定です');
      logOperation('統合処理', 'error', '認証情報が未設定のためスキップ');
      return;
    }
    
    // 最終チェック時刻を取得
    const lastCheckTime = PropertiesService.getScriptProperties()
      .getProperty('lastReplyCheck') || '0';
    
    const currentTime = Date.now();
    
    // 1. リプライ取得と自動返信を同時に実行
    console.log('リプライ取得と自動返信チェック開始');
    const result = fetchRepliesAndCheckAutoReply(lastCheckTime);
    console.log('リプライ取得と自動返信チェック完了');
    
    // 最終チェック時刻を更新
    PropertiesService.getScriptProperties()
      .setProperty('lastReplyCheck', currentTime.toString());
    
    console.log('===== 統合処理完了 =====');
    logOperation('統合処理', 'success', 
      `リプライ取得: ${result.total受信したリプライ}件, 自動返信: ${result.auto受信したリプライ}件`);
    
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
    // 最近の投稿を取得（件数を制限してAPI呼び出しを削減）
    const POST_LIMIT = 10; // 50件から10件に削減
    const recentPosts = getMyRecentPosts(POST_LIMIT);
    
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
  // 一時停止中ならスキップ
  try {
    const paused = PropertiesService.getScriptProperties().getProperty('GLOBAL_AUTOMATION_PAUSED') === 'true';
    if (paused) {
      console.log('グローバル一時停止中のため、リプライ取得をスキップ');
      return;
    }
  } catch (e) {}
  try {
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    
    if (!accessToken || !userId) {
      console.error('fetchAndSaveReplies: 認証情報が未設定です');
      logOperation('リプライ取得', 'error', '認証情報が未設定のためスキップ');
      return;
    }
    
    logOperation('リプライ取得開始', 'info', '最近の投稿からリプライを取得します');
    
    // 最近の投稿を取得（件数を制限してAPI呼び出しを削減）
    const POST_LIMIT = 10; // 50件から10件に削減
    const recentPosts = getMyRecentPosts(POST_LIMIT);
    
    if (!recentPosts || recentPosts.length === 0) {
      logOperation('リプライ取得', 'info', '投稿が見つかりませんでした');
      return;
    }
    
    let totalNew受信したリプライ = 0;
    let totalUpdated受信したリプライ = 0;
    
    // 各投稿のリプライを取得
    for (const post of recentPosts) {
      const result = fetchRepliesForPost(post);
      totalNew受信したリプライ += result.newCount;
      totalUpdated受信したリプライ += result.updatedCount;
    }
    
    logOperation('リプライ取得完了', 'success', 
      `新規: ${totalNew受信したリプライ}件, 更新: ${totalUpdated受信したリプライ}件`);
    
  } catch (error) {
    logError('fetchAndSaveReplies', error);
  }
}

// ===========================
// 自分の最近の投稿を取得（キャッシュ機能付き）
// ===========================
function getMyRecentPosts(limit = 25) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  // キャッシュの確認（5分間有効）
  const cache = CacheService.getScriptCache();
  const cacheKey = `recent_posts_${limit}`;
  const cachedData = cache.get(cacheKey);
  
  if (cachedData) {
    console.log('getMyRecentPosts: キャッシュから取得');
    return JSON.parse(cachedData);
  }
  
  try {
    console.log(`getMyRecentPosts: ${limit}件の投稿を取得中...`);
    
    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads?fields=id,text,timestamp,media_type,media_url,permalink&limit=${limit}`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const result = JSON.parse(response.getContentText());
    
    if (result.data) {
      console.log(`getMyRecentPosts: ${result.data.length}件の投稿を取得しました`);
      // キャッシュに保存（5分間）
      cache.put(cacheKey, JSON.stringify(result.data), 300);
      return result.data;
    } else {
      console.error('投稿取得エラー:', result.error?.message);
      return [];
    }
    
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
// 特定の投稿のリプライを取得し、自動返信をチェック
// ===========================
function fetchRepliesForPostAndAutoReply(post, lastCheckTime) {
  let newCount = 0;
  let auto受信したリプライ = 0;

  try {
    const config = RM_validateConfig();
    if (!config) return { newCount, auto受信したリプライ };

    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${post.id}/replies?fields=id,text,username,timestamp,from`,
      {
        headers: {
          'Authorization': `Bearer ${config.accessToken}`
        },
        muteHttpExceptions: true
      }
    );

    const result = JSON.parse(response.getContentText());

    if (result.data) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPLIES_SHEET_NAME);
      const twentyFourHoursAgo = new Date();
      twentyFourHoursAgo.setHours(twentyFourHoursAgo.getHours() - 24);

      for (const reply of result.data) {
        try {
          if (!reply.id || !reply.text) {
            console.log('不完全なリプライデータをスキップ:', reply);
            continue;
          }

          const replyTime = new Date(reply.timestamp);

          // 24時間以内のリプライのみ処理
          if (replyTime >= twentyFourHoursAgo) {
            const saveResult = saveReplyToSheet(reply, post, sheet);
            if (saveResult === 'new') newCount++;

            // 新しいリプライのみ判定
            if (replyTime.getTime() > parseInt(lastCheckTime)) {
              const username = reply.from?.username || reply.username;
              if (username === config.username) {
                console.log('自分のリプライのためスキップ');
                continue;
              }

              // 重複返信チェック
              const userId = reply.from?.id || username;
              if (RM_hasAlreadyRepliedToday(reply.id, userId)) {
                console.log('本日既に返信済みのためスキップ');
                continue;
              }

              const matchedKeyword = RM_findMatchingKeyword(reply.text);
              if (matchedKeyword) {
                const ok = RM_sendAutoReply(reply.id, reply, matchedKeyword, config);
                if (ok) auto受信したリプライ++;
              } else {
                console.log('マッチするキーワードなし');
              }
            }
          }
        } catch (replyError) {
          console.error(`リプライ処理エラー (ID: ${reply.id || 'unknown'}):`, replyError);
          continue;
        }
      }
    }

  } catch (error) {
    console.error(`投稿 ${post.id} のリプライ取得エラー:`, error);
  }

  return { newCount, auto受信したリプライ };
}

// ===========================
// 特定の投稿のリプライを取得
// ===========================
function fetchRepliesForPost(post) {
  const accessToken = getConfig('ACCESS_TOKEN');
  let newCount = 0;
  let updatedCount = 0;
  
  try {
    // 簡素化されたフィールドのみ取得
    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${post.id}/replies?fields=id,text,username,timestamp,from`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const result = JSON.parse(response.getContentText());
    
    if (result.data) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('受信したリプライ');
      const twentyFourHoursAgo = new Date();
      twentyFourHoursAgo.setHours(twentyFourHoursAgo.getHours() - 24);
      
      for (const reply of result.data) {
        // 24時間以内のリプライのみ処理
        const replyTime = new Date(reply.timestamp);
        if (replyTime >= twentyFourHoursAgo) {
          const saveResult = saveReplyToSheet(reply, post, sheet);
          if (saveResult === 'new') {
            newCount++;
          }
        }
      }
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
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPLIES_SHEET_NAME);
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
    console.warn('受信したリプライのトリミング中に警告:', e);
  }
  
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
  let existingSheet = spreadsheet.getSheetByName(REPLIES_SHEET_NAME);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet(REPLIES_SHEET_NAME);
  
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
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
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
    '1. 最近24時間のリプライを取得\n' +
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
      PropertiesService.getScriptProperties().setProperty('lastReplyCheck', '0');
      
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
// リプライ取得デバッグ
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
    
    console.log('認証情報: OK');
    
    // 2. 最近の投稿を1件だけ取得してテスト
    console.log('最近の投稿を取得中...');
    const recentPosts = getMyRecentPosts(1);
    
    if (!recentPosts || recentPosts.length === 0) {
      ui.alert('エラー', '投稿が取得できませんでした。\n\n考えられる原因:\n- アクセストークンが無効\n- 投稿が存在しない', ui.ButtonSet.OK);
      return;
    }
    
    const post = recentPosts[0];
    console.log(`投稿ID: ${post.id}, 投稿日時: ${post.timestamp}`);
    
    // 3. その投稿のリプライを取得
    console.log('リプライを取得中...');
    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${post.id}/replies?fields=id,text,timestamp,username,replied_to,is_reply`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    console.log(`レスポンスコード: ${responseCode}`);
    console.log(`レスポンス: ${responseText}`);
    
    const result = JSON.parse(responseText);
    
    if (result.error) {
      // エラーコード10は権限不足
      if (result.error.code === 10) {
        ui.alert(
          '権限エラー',
          'リプライを取得する権限がありません。\n\n' +
          'threads_manage_replies権限が必要です。\n' +
          'Meta開発者ダッシュボードで権限を追加してください。',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('エラー', `APIエラー: ${result.error.message}`, ui.ButtonSet.OK);
      }
      return;
    }
    
    // 4. 結果を表示
    const replies = result.data || [];
    const message = `デバッグ結果:\n\n` +
      `投稿ID: ${post.id}\n` +
      `投稿内容: ${post.text ? post.text.substring(0, 50) + '...' : '(なし)'}\n` +
      `リプライ数: ${replies.length}\n\n`;
    
    if (replies.length > 0) {
      const replyList = replies.slice(0, 3).map(r => 
        `- @${r.username}: ${r.text ? r.text.substring(0, 30) + '...' : '(なし)'}`
      ).join('\n');
      
      ui.alert('成功', message + '最新のリプライ:\n' + replyList, ui.ButtonSet.OK);
    } else {
      ui.alert('情報', message + 'この投稿にはまだリプライがありません。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert('エラー', `デバッグ中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}
