// PostingFunctions.gs - 投稿関連機能

// ===========================
// 予約投稿シート再構成
// ===========================
function resetScheduledPostsSheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '予約投稿シート再構成',
    '予約投稿シートを削除して再作成しますか？\n\n' +
    '⚠️ 既存の予約投稿データはすべて削除されます。\n' +
    '新しいシートにはサンプルデータが含まれます。',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      initializeScheduledPostsSheet();
      ui.alert('予約投稿シートを再構成しました。\n\n' +
        'サンプル投稿が追加されています。\n' +
        '必要に応じて編集してください。');
    } catch (error) {
      console.error('予約投稿シート再構成エラー:', error);
      ui.alert('エラーが発生しました: ' + error.message);
    }
  }
}

// ===========================
// 予約投稿シート初期化
// ===========================
function initializeScheduledPostsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを削除
  let existingSheet = spreadsheet.getSheetByName('予約投稿');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet('予約投稿');
  
  // ヘッダー行を設定
  const headers = [
    'ID',
    '投稿内容',
    '予定日付',
    '予定時刻',
    'ステータス',
    '投稿URL',
    'ツリーID',
    '投稿順序',
    '動画URL',
    '画像URL_1枚目',
    '画像URL_2枚目',
    '画像URL_3枚目',
    '画像URL_4枚目',
    '画像URL_5枚目',
    '画像URL_6枚目',
    '画像URL_7枚目',
    '画像URL_8枚目',
    '画像URL_9枚目',
    '画像URL_10枚目',
    'エラー'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#E0E0E0')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  // サンプルデータを追加
  const now = new Date();
  const futureDate = new Date(now.getFullYear() + 100, now.getMonth(), now.getDate());
  const futureDate2 = new Date(now.getFullYear() + 100, now.getMonth(), now.getDate() + 1);
  const futureDate3 = new Date(now.getFullYear() + 100, now.getMonth(), now.getDate() + 7);
  
  const sampleData = [
    [1, '【サンプル投稿1】\nこれはテスト投稿です。\n#threads #自動投稿', 
     Utilities.formatDate(futureDate, 'JST', 'yyyy/MM/dd'), '10:00',
     '投稿予約中', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    [2, '【サンプル投稿2】\n画像付き投稿のサンプルです。\n画像URLを右側の列に入力してください。', 
     Utilities.formatDate(futureDate2, 'JST', 'yyyy/MM/dd'), '15:00',
     '投稿予約中', '', '', '', '', 'https://example.com/image1.jpg', '', '', '', '', '', '', '', '', '', ''],
    [3, '【サンプル投稿3】\n複数画像の投稿サンプル', 
     Utilities.formatDate(futureDate3, 'JST', 'yyyy/MM/dd'), '12:00',
     '投稿予約中', '', '', '', '', 'https://example.com/image1.jpg', 'https://example.com/image2.jpg', '', '', '', '', '', '', '', '', ''],
    [4, '【ツリー投稿サンプル1-1】\nこれがツリーの最初の投稿です。',
     Utilities.formatDate(futureDate3, 'JST', 'yyyy/MM/dd'), '18:00',
     '投稿予約中', '', 'thread_A', '1', '', '', '', '', '', '', '', '', '', '', '', ''],
    [5, '【ツリー投稿サンプル1-2】\nこれが2番目の投稿。最初の投稿への返信になります。',
     Utilities.formatDate(futureDate3, 'JST', 'yyyy/MM/dd'), '18:00',
     '投稿予約中', '', 'thread_A', '2', '', '', '', '', '', '', '', '', '', '', '', ''],
    [6, '【ツリー投稿サンプル1-3】\nそして3番目。ツリーの最後です。',
     Utilities.formatDate(futureDate3, 'JST', 'yyyy/MM/dd'), '18:00',
     '投稿予約中', '', 'thread_A', '3', '', '', '', '', '', '', '', '', '', '', '', '']
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 50);   // ID
  sheet.setColumnWidth(2, 400);  // 投稿内容
  sheet.setColumnWidth(3, 100);  // 予定日付
  sheet.setColumnWidth(4, 80);   // 予定時刻
  sheet.setColumnWidth(5, 100);  // ステータス
  sheet.setColumnWidth(6, 300);  // 投稿URL
  sheet.setColumnWidth(7, 100);  // ツリーID
  sheet.setColumnWidth(8, 80);   // 投稿順序
  sheet.setColumnWidth(9, 200);  // 動画URL
  
  // 画像URL列の幅
  for (let i = 10; i <= 19; i++) {
    sheet.setColumnWidth(i, 200);
  }
  
  sheet.setColumnWidth(20, 300);  // エラー
  
  // データ検証を追加
  // ステータス列
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['投稿予約中', '投稿済', '失敗', 'キャンセル'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 1000, 1).setDataValidation(statusRule);
  
  // 日付列のフォーマット
  sheet.getRange(2, 3, 1000, 1).setNumberFormat('yyyy/mm/dd');
  
  // 時刻列のフォーマット設定
  sheet.getRange(2, 4, 1000, 1).setNumberFormat('hh:mm');
  
  // 15分単位の時刻リストを作成
  const timeOptions = [];
  for (let hour = 0; hour < 24; hour++) {
    for (let minute = 0; minute < 60; minute += 15) {
      const hourStr = hour.toString().padStart(2, '0');
      const minuteStr = minute.toString().padStart(2, '0');
      timeOptions.push(`${hourStr}:${minuteStr}`);
    }
  }
  
  // 時刻列にデータ検証（プルダウン）を追加
  const timeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(timeOptions)
    .setAllowInvalid(false)
    .setHelpText('15分単位で時刻を選択してください')
    .build();
  sheet.getRange(2, 4, 1000, 1).setDataValidation(timeRule);
  
  // 時刻列のヘルプテキスト
  sheet.getRange('D1').setNote('15分単位で時刻を選択できます（例: 13:30）');
  
  // 説明コメントを追加
  sheet.getRange('A1').setNote(
    '予約投稿の設定シート\\n\\n' +
    'ID: 一意の識別番号\\n' +
    '投稿内容: Threadsに投稿するテキスト\\n' +
    '予定日付: 投稿する日付（yyyy/mm/dd形式）\\n' +
    '予定時刻: 投稿する時刻（24時間表記、例: 13:30）\\n' +
    'ステータス: 投稿予約中、投稿済、失敗、キャンセル\\n' +
    '投稿URL: 投稿後のURL（自動入力）\\n' +
    'ツリーID: 同じツリーにする投稿に共通のIDを入力（例: thread_A）\\n' +
    '投稿順序: ツリー内での投稿順番（1, 2, 3...）\\n' +
    '動画URL: 投稿に添付する動画のURL\\n' +
    '画像URL: 投稿に添付する画像のURL（最大10枚）\\n' +
    'エラー: エラーメッセージ（自動入力）'
  );
  
  // ツリーID列のヘルプテキスト
  sheet.getRange('G1').setNote('同じツリー（スレッド）として投稿したい複数の行に、共通のIDを入力してください');
  
  // 投稿順序列のヘルプテキスト
  sheet.getRange('H1').setNote('ツリー内での投稿順序を数値で指定（1が最初の投稿）');
  
  // 動画URL列のヘルプテキスト
  sheet.getRange('I1').setNote('動画ファイルのURLを入力してください（Google DriveのURLも可）');
  
  // 条件付き書式を追加（ステータスの視覚化）
  const statusRange = sheet.getRange(2, 5, 1000, 1);
  
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('投稿予約中')
    .setBackground('#F5F5F5')
    .setFontColor('#333333')
    .setRanges([statusRange])
    .build();
  
  const publishedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('投稿済')
    .setBackground('#FFFFFF')
    .setFontColor('#000000')
    .setRanges([statusRange])
    .build();
    
  const failedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('失敗')
    .setBackground('#F5F5F5')
    .setFontColor('#FF0000')
    .setRanges([statusRange])
    .build();
    
  const cancelledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('キャンセル')
    .setBackground('#E0E0E0')
    .setFontColor('#666666')
    .setRanges([statusRange])
    .build();
  
  sheet.setConditionalFormatRules([pendingRule, publishedRule, failedRule, cancelledRule]);
  
  // 1行目を固定
  sheet.setFrozenRows(1);
  
  logOperation('予約投稿シート初期化', 'success', 'シートを再構成しました');
}

// ===========================
// 予約投稿処理
// ===========================
function processScheduledPosts() {
  // 必ず実行開始をログに記録
  console.log('===== processScheduledPosts 開始 =====');
  logOperation('予約投稿処理', 'start', `実行開始時刻: ${new Date().toLocaleString('ja-JP')}`);
  
  try {
    // API制限チェック（本日のAPI使用状況を確認）
    const dailyQuota = checkDailyAPIQuota();
    if (dailyQuota.isNearLimit) {
      console.log('API制限に近づいています。処理を制限します。');
      logOperation('API制限警告', 'warning', `本日のAPI使用数: ${dailyQuota.count}/20000`);
    }
    // 認証情報の事前チェック
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    
    if (!accessToken || !userId) {
      console.error('processScheduledPosts: 認証情報が未設定です');
      logOperation('予約投稿処理', 'error', '認証情報が未設定のためスキップ');
      return;
    }
    
    const posts = getScheduledPosts();
    
    if (posts.length === 0) {
      console.log('processScheduledPosts: 処理対象の投稿がありません');
      logOperation('予約投稿処理', 'info', '処理対象の投稿なし');
      return;
    }
    
    // 投稿をツリーごとにグループ化
    const postGroups = groupPostsByTree(posts);
    const totalGroups = Object.keys(postGroups).length;
    
    logOperation('予約投稿処理開始', 'info', `${posts.length}件の投稿を${totalGroups}グループとして処理`);
    
    // グループごとに処理
    Object.entries(postGroups).forEach(([groupId, groupPosts]) => {
      try {
        if (groupPosts.length === 1 && !groupPosts[0].treeId) {
          // 単発投稿
          processPost(groupPosts[0]);
        } else {
          // ツリー投稿
          processTreePosts(groupPosts);
        }
      } catch (error) {
        logError('processPostGroup', error);
        
        // API制限エラーの場合は、投稿を「投稿予約中」に戻す
        if (error.toString().includes('urlfetch') || error.toString().includes('制限')) {
          groupPosts.forEach(post => {
            updatePostStatus(post.row, 'pending', null, 
              `API制限により延期されました。次回の実行時に再試行されます。`);
          });
        } else {
          // その他のエラーの場合
          groupPosts.forEach(post => {
            updatePostStatus(post.row, 'failed', null, error.toString());
          });
        }
      }
    });
    
  } catch (error) {
    logError('processScheduledPosts', error);
  }
}

// ===========================
// 投稿取得
// ===========================
function getScheduledPosts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  const data = sheet.getDataRange().getValues();
  const posts = [];
  const now = new Date();
  
  console.log(`現在時刻（JST）: ${Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss')}`);
  
  for (let i = 1; i < data.length; i++) {
    // 新しい列構造: ID(1), 投稿内容(2), 予定日付(3), 予定時刻(4), ステータス(5), 投稿URL(6), ツリーID(7), 投稿順序(8), 動画URL(9), 画像URL_1枚目〜10枚目(10-19), エラー(20)
    const row = data[i];
    const id = row[0];
    const content = row[1];
    const scheduledDate = row[2];
    const scheduledTime = row[3];
    const status = row[4];
    const postedUrl = row[5];
    const treeId = row[6];
    const postOrder = row[7];
    const videoUrl = row[8];
    const imageUrls = row.slice(9, 19); // 画像URL_1枚目〜10枚目
    const error = row[19];
    
    // デバッグ: 生データを出力
    console.log(`行${i + 1} 生データ: ID=${id}, 日付=${scheduledDate}, 時刻=${scheduledTime}, ステータス="${status}"`);
    
    // 空行をスキップ
    if (!id && !content) {
      console.log(`行${i + 1}: 空行のためスキップ`);
      continue;
    }
    
    // 日付と時刻を統合
    let scheduledDateTime;
    try {
      if (scheduledDate && scheduledTime) {
        console.log(`行${i + 1}: 日付処理開始 - scheduledDate=${scheduledDate}, type=${typeof scheduledDate}`);
        console.log(`行${i + 1}: 時刻処理開始 - scheduledTime=${scheduledTime}, type=${typeof scheduledTime}`);
        
        // 日付の処理
        const dateObj = new Date(scheduledDate);
        const year = dateObj.getFullYear();
        const month = dateObj.getMonth();
        const day = dateObj.getDate();
        
        // 時刻の処理（Dateオブジェクトまたは文字列）
        let hours, minutes;
        if (scheduledTime instanceof Date) {
          // Google SheetsではTimeのみの値は1899年の日付として保存される
          hours = scheduledTime.getHours();
          minutes = scheduledTime.getMinutes();
        } else {
          // 文字列の場合
          [hours, minutes] = scheduledTime.toString().split(':').map(s => parseInt(s));
        }
        
        // JSTで日時を作成
        scheduledDateTime = new Date(year, month, day, hours, minutes, 0);
        
        console.log(`行${i + 1}: 予定日時=${Utilities.formatDate(scheduledDateTime, 'JST', 'yyyy/MM/dd HH:mm:ss')}`);
      } else {
        continue; // 日付または時刻がない場合はスキップ
      }
    } catch (e) {
      console.error(`行${i + 1}の日時解析エラー:`, e);
      continue;
    }
    
    if (status === 'pending' || status === '投稿予約中') {
      const isPastDue = scheduledDateTime <= now;
      console.log(`行${i + 1}: ステータス=${status}, 過去の予定=${isPastDue}`);
      
      if (isPastDue) {
        // 空でない画像URLのみを収集（最大10枚）
        const validImageUrls = imageUrls.filter(url => url && url.trim() !== '');
        
        posts.push({
          row: i + 1,
          id: id,
          content: content,
          imageUrls: validImageUrls,
          videoUrl: videoUrl || '',
          scheduledTime: scheduledDateTime,
          retryCount: 0, // リトライ列は削除したため常に0
          treeId: treeId || '',
          postOrder: postOrder || 0
        });
        
        console.log(`行${i + 1}: 予約投稿として処理対象に追加`);
      }
    }
  }
  
  return posts;
}

// ===========================
// ツリーごとに投稿をグループ化
// ===========================
function groupPostsByTree(posts) {
  const groups = {};
  
  posts.forEach(post => {
    let groupKey;
    if (post.treeId) {
      // ツリーIDがある場合は、それをグループキーとする
      groupKey = `tree_${post.treeId}`;
    } else {
      // ツリーIDがない場合は、個別のグループとする
      groupKey = `single_${post.row}`;
    }
    
    if (!groups[groupKey]) {
      groups[groupKey] = [];
    }
    
    groups[groupKey].push(post);
  });
  
  // 各グループ内で投稿順序でソート
  Object.keys(groups).forEach(key => {
    groups[key].sort((a, b) => {
      // まず投稿順序でソート
      if (a.postOrder && b.postOrder) {
        return parseInt(a.postOrder) - parseInt(b.postOrder);
      }
      // 投稿順序がない場合は行番号でソート
      return a.row - b.row;
    });
  });
  
  return groups;
}

// ===========================
// ツリー投稿の処理
// ===========================
function processTreePosts(posts) {
  console.log(`ツリー投稿開始: ${posts.length}件の投稿`);
  logOperation('ツリー投稿', 'start', `ツリーID: ${posts[0].treeId}, 投稿数: ${posts.length}`);
  
  let previousPostId = null;
  
  for (let i = 0; i < posts.length; i++) {
    const post = posts[i];
    
    try {
      console.log(`ツリー投稿 ${i + 1}/${posts.length}: 行${post.row}`);
      
      // 返信先IDを設定（最初の投稿以外）
      if (i > 0 && previousPostId) {
        post.replyToId = previousPostId;
      }
      
      // 投稿を実行
      const result = processPost(post);
      
      if (result && result.postId) {
        previousPostId = result.postId;
        console.log(`投稿成功: ID=${previousPostId}`);
        
        // API制限対策として、次の投稿までに待機（最後の投稿以外）
        if (i < posts.length - 1) {
          console.log('次の投稿まで5秒待機...');
          Utilities.sleep(5000); // 3秒から5秒に増やす
        }
      } else {
        // 投稿失敗時は、残りのツリー投稿をキャンセル
        console.error('投稿失敗 - result:', result);
        throw new Error(`ツリー投稿の${i + 1}番目が失敗しました`);
      }
      
    } catch (error) {
      console.error(`ツリー投稿エラー: ${error}`);
      logOperation('ツリー投稿', 'error', `${i + 1}番目の投稿でエラー: ${error.toString()}`);
      
      // API制限エラーの場合は、投稿を「投稿予約中」に戻す
      if (error.toString().includes('urlfetch') || error.toString().includes('制限')) {
        for (let j = i; j < posts.length; j++) {
          updatePostStatus(posts[j].row, 'pending', null, 
            `API制限により延期されました。次回の実行時に再試行されます。`);
        }
        logOperation('API制限', 'warning', 'URLFetch制限に達したため、残りの投稿を延期しました');
      } else {
        // その他のエラーの場合
        for (let j = i; j < posts.length; j++) {
          updatePostStatus(posts[j].row, 'failed', null, 
            `ツリー投稿の${i + 1}番目でエラーが発生したため、この投稿はキャンセルされました`);
        }
      }
      
      throw error;
    }
  }
  
  logOperation('ツリー投稿', 'success', `ツリーID: ${posts[0].treeId}, ${posts.length}件すべて成功`);
}

// ===========================
// 投稿処理
// ===========================
function processPost(post) {
  const maxRetries = 3;
  
  if (post.retryCount >= maxRetries) {
    updatePostStatus(post.row, 'failed', null, '最大リトライ回数に達しました');
    return null;
  }
  
  let result;
  
  // 返信先IDがある場合（ツリー投稿の2番目以降）
  if (post.replyToId) {
    if (post.imageUrls && post.imageUrls.length > 0) {
      // 画像付き返信
      if (post.imageUrls.length === 1) {
        result = postReplyWithImage(post.content, post.imageUrls[0], post.replyToId);
      } else {
        result = postReplyWithMultipleImages(post.content, post.imageUrls, post.replyToId);
      }
    } else {
      // テキストのみ返信
      result = postReplyTextOnly(post.content, post.replyToId);
    }
  } else {
    // 通常の投稿（返信ではない）
    if (post.videoUrl && post.videoUrl.trim() !== '') {
      // 動画付き投稿
      result = postWithVideo(post.content, post.videoUrl);
    } else if (post.imageUrls && post.imageUrls.length > 0) {
      // 画像付き投稿（複数画像対応）
      if (post.imageUrls.length === 1) {
        // 1枚の場合は従来の関数を使用
        result = postWithImage(post.content, post.imageUrls[0]);
      } else {
        // 複数枚の場合は新しい関数を使用
        result = postWithMultipleImages(post.content, post.imageUrls);
      }
    } else {
      // テキストのみ投稿
      result = postTextOnly(post.content);
    }
  }
  
  if (result.success) {
    updatePostStatus(post.row, 'posted', result.postUrl, null);
    logOperation('投稿成功', 'success', `ID: ${post.id}, URL: ${result.postUrl}`);
    return result; // 投稿IDを含む結果を返す
  } else {
    // API制限エラーの場合は、リトライカウントを増やさない
    if (result.error && (result.error.includes('urlfetch') || result.error.includes('制限'))) {
      updatePostStatus(post.row, 'pending', null, 
        `API制限により延期されました。次回の実行時に再試行されます。`, post.retryCount);
      logOperation('投稿延期', 'warning', `ID: ${post.id}, API制限により延期`);
    } else {
      const newRetryCount = post.retryCount + 1;
      updatePostStatus(post.row, 'pending', null, result.error, newRetryCount);
      logOperation('投稿失敗', 'error', `ID: ${post.id}, エラー: ${result.error}`);
    }
    return null;
  }
}

// ===========================
// テキスト投稿
// ===========================
function postTextOnly(text) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    console.log('postTextOnly 開始 - テキスト:', text.substring(0, 50) + '...');
    
    // メディアコンテナの作成
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
          text: text
        }),
        muteHttpExceptions: true
      }
    );
    
    const responseCode = createResponse.getResponseCode();
    const responseText = createResponse.getContentText();
    console.log('API レスポンスコード:', responseCode);
    console.log('API レスポンス:', responseText);
    
    const createResult = JSON.parse(responseText);
    incrementAPICallCount(1); // API使用回数を記録
    
    if (createResult.id) {
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      return { success: false, error: createResult.error?.message || 'コンテナ作成失敗' };
    }
    
  } catch (error) {
    console.error('postTextOnly エラー詳細:', error);
    console.error('エラースタック:', error.stack);
    console.error('エラーメッセージ:', error.message);
    
    // URLFetch制限の特定のエラーメッセージを確認
    const errorMessage = error.toString();
    if (errorMessage.includes('UrlFetch') || errorMessage.includes('urlfetch') || errorMessage.includes('サービス')) {
      console.error('URLFetch制限エラーを検出');
      return { success: false, error: `API制限エラー: ${errorMessage}` };
    }
    
    return { success: false, error: errorMessage };
  }
}

// ===========================
// 画像付き投稿
// ===========================
function postWithImage(text, imageUrl) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    // Google DriveのURLを公開URLに変換
    const publicImageUrl = convertToPublicUrl(imageUrl);
    
    // メディアコンテナの作成
    const createResponse = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          media_type: 'IMAGE',
          image_url: publicImageUrl,
          text: text || ''
        }),
        muteHttpExceptions: true
      }
    );
    
    const createResult = JSON.parse(createResponse.getContentText());
    
    if (createResult.id) {
      // 少し待機（画像処理のため）
      Utilities.sleep(3000);
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      return { success: false, error: createResult.error?.message || 'コンテナ作成失敗' };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===========================
// 動画付き投稿
// ===========================
// ===========================
// Drive動画URLの正規化
// ===========================
function convertVideoUrl(url) {
  // Google DriveのURLパターンをチェック
  const drivePatterns = [
    { regex: /drive\.google\.com\/file\/d\/([a-zA-Z0-9-_]+)/, type: 'file' },
    { regex: /drive\.google\.com\/open\?id=([a-zA-Z0-9-_]+)/, type: 'open' },
    { regex: /docs\.google\.com\/.*\/d\/([a-zA-Z0-9-_]+)/, type: 'docs' }
  ];
  
  for (const pattern of drivePatterns) {
    const match = url.match(pattern.regex);
    if (match) {
      const fileId = match[1];
      // Cookie/confirm不要の直リンクURLに変換
      const directUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      console.log(`Drive URL正規化: ${url} → ${directUrl}`);
      return directUrl;
    }
  }
  
  // DriveのURLでない場合はそのまま返す
  return url;
}

// ===========================
// 動画URLの事前プローブ（Threads要件チェック）
// ===========================
function probeMediaUrlForThreads(url, options = {}) {
  const { isVideo = true } = options;  // デフォルトをtrueに変更（動画用関数のため）
  
  try {
    console.log(`プローブ開始: ${url}`);
    
    // Range requestで最初の1バイトのみ取得（ファイルサイズとContent-Type確認用）
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 
        'Range': 'bytes=0-0',
        'User-Agent': 'Meta-Threads-Media-Probe'
      },
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    });
    
    const responseCode = response.getResponseCode();
    console.log(`HTTPレスポンスコード: ${responseCode}`);
    
    // 200 OK または 206 Partial Content が期待される
    if (!(responseCode === 200 || responseCode === 206)) {
      return { 
        isValid: false, 
        error: `HTTPエラー: ${responseCode}` 
      };
    }
    
    const headers = response.getHeaders();
    const contentType = String(headers['Content-Type'] || headers['content-type'] || '').toLowerCase();
    console.log(`Content-Type: ${contentType}`);
    
    // 動画の場合はContent-Typeチェック
    if (isVideo) {
      if (!contentType.startsWith('video/')) {
        // HTMLが返ってきた場合は認証/確認ページの可能性
        if (contentType.includes('html')) {
          return { 
            isValid: false, 
            error: `公開直リンクではありません（HTML応答）。Driveの共有設定を「リンクを知っている全員」にしてください。` 
          };
        }
        return { 
          isValid: false, 
          error: `Content-Typeがvideo/*ではありません: ${contentType}` 
        };
      }
      
      // サポートされる動画形式のチェック
      const supportedTypes = ['video/mp4', 'video/quicktime'];
      if (!supportedTypes.some(type => contentType.includes(type))) {
        console.warn(`警告: 推奨されない動画形式: ${contentType}`);
      }
    }
    
    // Content-Lengthチェック（サイズ上限）
    const contentLengthRaw = headers['Content-Length'] || headers['content-length'] || 
                            headers['Content-Range'] || headers['content-range'];
    
    let contentLength = null;
    if (contentLengthRaw) {
      // Content-Rangeの場合: "bytes 0-0/12345" から全体サイズを抽出
      if (String(contentLengthRaw).includes('/')) {
        const parts = String(contentLengthRaw).split('/');
        if (parts.length === 2) {
          contentLength = parseInt(parts[1], 10);
        }
      } else {
        contentLength = parseInt(contentLengthRaw, 10);
      }
    }
    
    if (contentLength) {
      console.log(`ファイルサイズ: ${Math.round(contentLength / 1024 / 1024)}MB`);
      
      // サイズ上限チェック（デフォルト50MB）
      const maxSizeMB = 50;  // 50MB固定
      const maxSizeBytes = maxSizeMB * 1024 * 1024;
      
      if (contentLength > maxSizeBytes) {
        return { 
          isValid: false, 
          error: `ファイルサイズが50MBを超えています: ${Math.round(contentLength / 1024 / 1024)}MB` 
        };
      }
    } else {
      console.warn('警告: Content-Lengthが取得できませんでした');
    }
    
    // すべてのチェックをパス
    return { 
      isValid: true, 
      contentType: contentType,
      contentLength: contentLength 
    };
    
  } catch (error) {
    console.error('プローブエラー:', error);
    return { 
      isValid: false, 
      error: `アクセスエラー: ${error.toString()}` 
    };
  }
}

// ===========================

function postWithVideo(text, videoUrl) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    // Google DriveのURLを正規化
    const normalizedUrl = convertVideoUrl(videoUrl);
    
    console.log('動画投稿作成中...');
    console.log(`  元URL: ${videoUrl}`);
    console.log(`  正規化URL: ${normalizedUrl}`);
    console.log(`  テキスト: ${text || 'なし'}`);
    
    // 動画URLの事前検証
    console.log('動画URLの検証中...');
    const probeResult = probeMediaUrlForThreads(normalizedUrl);
    
    if (!probeResult.isValid) {
      console.error('動画URL検証失敗:', probeResult.error);
      
      // ユーザーフレンドリーなエラーメッセージ
      let userMessage = '動画の投稿に失敗しました: ';
      
      // probeResult.errorが存在することを確認してから.includes()を使用
      if (probeResult.error) {
        if (probeResult.error.includes('50MB')) {
          userMessage += 'ファイルサイズが50MBを超えています。より小さいファイルを使用してください。';
        } else if (probeResult.error.includes('Content-Type')) {
          userMessage += '動画ファイルではないか、形式がサポートされていません。';
        } else if (probeResult.error.includes('404')) {
          userMessage += 'ファイルが見つかりません。URLと共有設定を確認してください。';
        } else if (probeResult.error.includes('403')) {
          userMessage += 'ファイルへのアクセスが拒否されました。共有設定を確認してください。';
        } else {
          userMessage += probeResult.error;
        }
      } else {
        userMessage += '不明なエラーが発生しました。';
      }
      
      return { success: false, error: userMessage };
    }
    
    console.log('動画URL検証成功:');
    console.log(`  Content-Type: ${probeResult.contentType}`);
    if (probeResult.contentLength) {
      console.log(`  ファイルサイズ: ${Math.round(probeResult.contentLength / 1024 / 1024 * 100) / 100} MB`);
    }
    
    // メディアコンテナの作成
    const createUrl = `${THREADS_API_BASE}/v1.0/${userId}/threads`;
    const createParams = {
      'media_type': 'VIDEO',
      'video_url': normalizedUrl,
      'text': text || ''
    };
    
    console.log('API呼び出しURL:', createUrl);
    console.log('パラメータ:', JSON.stringify(createParams));
    
    const createResponse = fetchWithTracking(createUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: Object.keys(createParams).map(key => 
        `${key}=${encodeURIComponent(createParams[key])}`
      ).join('&'),
      muteHttpExceptions: true
    });
    
    const createResult = JSON.parse(createResponse.getContentText());
    console.log('メディアコンテナ作成レスポンス:', JSON.stringify(createResult));
    
    if (createResult.id) {
      // 動画処理のため長めに待機
      console.log('動画処理待機中（15秒）...');
      Utilities.sleep(15000);
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      const errorMessage = createResult.error?.message || 
                         createResult.error?.error_user_msg || 
                         'ビデオコンテナ作成失敗';
      console.error('動画投稿エラー:', errorMessage);
      return { success: false, error: errorMessage };
    }
    
  } catch (error) {
    console.error('動画投稿例外:', error);
    return { success: false, error: error.toString() };
  }
}

// ===========================
// 複数画像付き投稿
// ===========================
function postWithMultipleImages(text, imageUrls) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    // 各画像URLを公開URLに変換
    const publicUrls = imageUrls.map(url => convertToPublicUrl(url));
    
    // メディアコンテナの作成
    const payload = {
      media_type: 'CAROUSEL',
      children: publicUrls.map(url => ({ media_url: url })),
      text: text
    };
    
    const createResponse = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    
    const createResult = JSON.parse(createResponse.getContentText());
    
    if (createResult.id) {
      // 少し待機（画像処理のため）
      Utilities.sleep(5000); // 複数画像の場合は少し長めに待機
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      return { success: false, error: createResult.error?.message || 'カルーセル作成失敗' };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===========================
// 投稿公開
// ===========================
function publishPost(containerId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  try {
    const publishResponse = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads_publish`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          creation_id: containerId
        }),
        muteHttpExceptions: true
      }
    );
    
    const publishResult = JSON.parse(publishResponse.getContentText());
    incrementAPICallCount(1); // API使用回数を記録
    
    if (publishResult.id) {
      // 投稿URLの構築
      const username = getConfig('USERNAME');
      const postUrl = `https://www.threads.net/@${username}/post/${publishResult.id}`;
      
      return { success: true, postUrl: postUrl, postId: publishResult.id };
    } else {
      return { success: false, error: publishResult.error?.message || '公開失敗' };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===========================
// 返信投稿（テキストのみ）
// ===========================
function postReplyTextOnly(text, replyToId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    // メディアコンテナの作成（返信先IDを指定）
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
          text: text,
          reply_to_id: replyToId
        }),
        muteHttpExceptions: true
      }
    );
    
    const createResult = JSON.parse(createResponse.getContentText());
    incrementAPICallCount(1); // API使用回数を記録
    
    if (createResult.id) {
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      return { success: false, error: createResult.error?.message || 'コンテナ作成失敗' };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===========================
// 返信投稿（画像付き）
// ===========================
function postReplyWithImage(text, imageUrl, replyToId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    // Google DriveのURLを公開URLに変換
    const publicImageUrl = convertToPublicUrl(imageUrl);
    
    // メディアコンテナの作成（返信先IDを指定）
    const createResponse = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          media_type: 'IMAGE',
          image_url: publicImageUrl,
          text: text || '',
          reply_to_id: replyToId
        }),
        muteHttpExceptions: true
      }
    );
    
    const createResult = JSON.parse(createResponse.getContentText());
    
    if (createResult.id) {
      // 少し待機（画像処理のため）
      Utilities.sleep(3000);
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      return { success: false, error: createResult.error?.message || 'コンテナ作成失敗' };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===========================
// 返信投稿（複数画像付き）
// ===========================
function postReplyWithMultipleImages(text, imageUrls, replyToId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    // 各画像URLを公開URLに変換
    const publicUrls = imageUrls.map(url => convertToPublicUrl(url));
    
    // メディアコンテナの作成（返信先IDを指定）
    const payload = {
      media_type: 'CAROUSEL',
      children: publicUrls.map(url => ({ media_url: url })),
      text: text,
      reply_to_id: replyToId
    };
    
    const createResponse = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${userId}/threads`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    
    const createResult = JSON.parse(createResponse.getContentText());
    
    if (createResult.id) {
      // 少し待機（画像処理のため）
      Utilities.sleep(5000);
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      return { success: false, error: createResult.error?.message || 'コンテナ作成失敗' };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===========================
// URL変換
// ===========================
function convertToPublicUrl(url) {
  // Google DriveのURLパターンをチェック
  const drivePatterns = [
    /drive\.google\.com\/file\/d\/([a-zA-Z0-9-_]+)/,
    /drive\.google\.com\/open\?id=([a-zA-Z0-9-_]+)/,
    /docs\.google\.com\/.*\/d\/([a-zA-Z0-9-_]+)/
  ];
  
  for (const pattern of drivePatterns) {
    const match = url.match(pattern);
    if (match) {
      const fileId = match[1];
      
      try {
        // ファイルの共有設定を変更
        const file = DriveApp.getFileById(fileId);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // 直接ダウンロードURLを返す
        return `https://drive.google.com/uc?export=download&id=${fileId}`;
      } catch (error) {
        logError('convertToPublicUrl', error);
        // エラーの場合は元のURLを返す
        return url;
      }
    }
  }
  
  // Google Drive以外のURLはそのまま返す
  return url;
}

// ===========================
// ステータス更新
// ===========================
function updatePostStatus(row, status, postUrl, error, retryCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  
  // 新しい列構造: ID(1), 投稿内容(2), 予定日付(3), 予定時刻(4), ステータス(5), 投稿URL(6), ツリーID(7), 投稿順序(8), 動画URL(9), 画像URL_1枚目〜10枚目(10-19), エラー(20)
  
  // ステータスを日本語に変換
  const statusMap = {
    'pending': '投稿予約中',
    'posted': '投稿済',
    'published': '投稿済',
    'failed': '失敗',
    'cancelled': 'キャンセル'
  };
  
  const japaneseStatus = statusMap[status] || status;
  
  // ステータス更新（列5）
  sheet.getRange(row, 5).setValue(japaneseStatus);
  
  // 投稿URL更新（列6）
  if (postUrl) {
    sheet.getRange(row, 6).setValue(postUrl);
  }
  
  // エラー更新（列20 - 最右端）
  if (error) {
    sheet.getRange(row, 20).setValue(error);
  } else {
    sheet.getRange(row, 20).setValue('');
  }
  
  // リトライ回数の更新は削除（リトライ列を削除したため）
}

// ===========================
// API使用状況チェック
// ===========================
function checkDailyAPIQuota() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const today = new Date().toDateString();
  const lastResetDate = scriptProperties.getProperty('API_QUOTA_RESET_DATE');
  const apiCount = parseInt(scriptProperties.getProperty('API_CALL_COUNT') || '0');
  
  // 日付が変わっていたらリセット
  if (lastResetDate !== today) {
    scriptProperties.setProperty('API_QUOTA_RESET_DATE', today);
    scriptProperties.setProperty('API_CALL_COUNT', '0');
    return { count: 0, isNearLimit: false };
  }
  
  // Google Apps ScriptのUrlFetchApp制限は1日20,000回
  const DAILY_LIMIT = 20000;
  const WARNING_THRESHOLD = 19000; // 95%で警告
  
  return {
    count: apiCount,
    isNearLimit: apiCount >= WARNING_THRESHOLD,
    remaining: DAILY_LIMIT - apiCount
  };
}

// ===========================
// API使用回数を記録
// ===========================
function incrementAPICallCount(count = 1) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentCount = parseInt(scriptProperties.getProperty('API_CALL_COUNT') || '0');
  scriptProperties.setProperty('API_CALL_COUNT', String(currentCount + count));
}

// ===========================
// 予約投稿データの確認（デバッグ用）
// ===========================
function checkScheduledPostsData() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  
  if (!sheet) {
    ui.alert('エラー', '予約投稿シートが見つかりません', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  console.log('===== 予約投稿データ確認 =====');
  console.log(`データ行数: ${data.length}`);
  
  let message = '予約投稿データ:\n\n';
  
  for (let i = 1; i < Math.min(data.length, 6); i++) { // 最初の5行まで
    const row = data[i];
    console.log(`行${i + 1}: ${JSON.stringify(row.slice(0, 5))}`); // 最初の5列まで
    
    const [id, content, scheduledDate, scheduledTime, status] = row;
    message += `行${i + 1}:\n`;
    message += `  ID: ${id}\n`;
    message += `  内容: ${content ? content.substring(0, 20) + '...' : '空'}\n`;
    message += `  日付: ${scheduledDate} (${typeof scheduledDate})\n`;
    message += `  時刻: ${scheduledTime} (${typeof scheduledTime})\n`;
    message += `  ステータス: "${status}"\n\n`;
  }
  
  ui.alert('データ確認', message, ui.ButtonSet.OK);
}

// ===========================
// 予約投稿の強制実行（過去の投稿も含む）
// ===========================
function forceProcessScheduledPosts() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '予約投稿の強制実行',
    '過去の予約投稿も含めてすべて実行します。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // 一時的に過去の投稿も処理するようにフラグを設定
  PropertiesService.getScriptProperties().setProperty('FORCE_PROCESS_OLD_POSTS', 'true');
  
  try {
    processScheduledPosts();
    ui.alert('完了', '予約投稿の処理が完了しました。ログシートを確認してください。', ui.ButtonSet.OK);
  } finally {
    // フラグをクリア
    PropertiesService.getScriptProperties().deleteProperty('FORCE_PROCESS_OLD_POSTS');
  }
}

// ===========================
// 予約投稿のデバッグ実行
// ===========================
function debugScheduledPosts() {
  const ui = SpreadsheetApp.getUi();
  console.log('===== 予約投稿デバッグ開始 =====');
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
    const data = sheet.getDataRange().getValues();
    
    console.log(`データ行数: ${data.length - 1}`);
    
    // 現在時刻を表示
    const now = new Date();
    console.log(`現在時刻: ${Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss')}`);
    
    // 各行の状態を確認
    for (let i = 1; i < data.length && i <= 5; i++) { // 最初の5行まで
      const [id, content, scheduledDate, scheduledTime, status] = data[i];
      console.log(`\n行${i + 1}:`);
      console.log(`  ID: ${id}`);
      console.log(`  内容: ${content ? content.substring(0, 30) + '...' : '空'}`);
      console.log(`  予定日付: ${scheduledDate}`);
      console.log(`  予定時刻: ${scheduledTime}`);
      console.log(`  ステータス: ${status}`);
    }
    
    // 実際の処理を実行
    processScheduledPosts();
    
    ui.alert('デバッグ完了', 'ログシートを確認してください。', ui.ButtonSet.OK);
  } catch (error) {
    console.error('デバッグエラー:', error);
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

// ===========================
// トリガー状態確認
// ===========================
function checkScheduledPostTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const scheduledPostTriggers = triggers.filter(trigger => 
      trigger.getHandlerFunction() === 'processScheduledPosts'
    );
    
    if (scheduledPostTriggers.length === 0) {
      ui.alert(
        'トリガー未設定',
        '予約投稿のトリガーが設定されていません。\n' +
        'メニューから「トリガー設定」を実行してください。',
        ui.ButtonSet.OK
      );
    } else {
      let message = `予約投稿トリガー: ${scheduledPostTriggers.length}個\n\n`;
      scheduledPostTriggers.forEach((trigger, index) => {
        message += `${index + 1}. ${trigger.getTriggerSource()} - ${trigger.getEventType()}\n`;
      });
      ui.alert('トリガー状態', message, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('エラー', `トリガー確認エラー: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ===========================
// 手動実行
// ===========================
function manualPostExecution() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '手動投稿実行',
    '予約投稿を今すぐ実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    processScheduledPosts();
    ui.alert('投稿処理を実行しました。詳細はログをご確認ください。');
  }
}

// ===========================
// アクセストークンリフレッシュ
// ===========================
function refreshAccessToken() {
  const refreshToken = getConfig('REFRESH_TOKEN');
  const clientId = getConfig('CLIENT_ID');
  const clientSecret = getConfig('CLIENT_SECRET');
  
  if (!refreshToken) {
    logError('refreshAccessToken', 'リフレッシュトークンが見つかりません');
    return;
  }
  
  try {
    const response = fetchWithTracking(
      'https://graph.threads.net/oauth/access_token',
      {
        method: 'POST',
        payload: {
          grant_type: 'refresh_token',
          client_id: clientId,
          client_secret: clientSecret,
          refresh_token: refreshToken
        },
        muteHttpExceptions: true
      }
    );
    
    const result = JSON.parse(response.getContentText());
    
    if (result.access_token) {
      setConfig('ACCESS_TOKEN', result.access_token);
      setConfig('TOKEN_EXPIRES', new Date(Date.now() + result.expires_in * 1000).toISOString());
      logOperation('トークンリフレッシュ', 'success', '新しいトークンを取得');
    } else {
      throw new Error(result.error?.message || 'トークンリフレッシュ失敗');
    }
    
  } catch (error) {
    logError('refreshAccessToken', error);
  }
}