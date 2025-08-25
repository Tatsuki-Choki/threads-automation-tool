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
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
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
    .setBackground('#FFF3CD')
    .setFontColor('#856404')
    .setRanges([statusRange])
    .build();
  
  const publishedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('投稿済')
    .setBackground('#D4EDDA')
    .setFontColor('#155724')
    .setRanges([statusRange])
    .build();
    
  const failedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('失敗')
    .setBackground('#F8D7DA')
    .setFontColor('#721C24')
    .setRanges([statusRange])
    .build();
    
  const cancelledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('キャンセル')
    .setBackground('#E7E8EA')
    .setFontColor('#383D41')
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
  console.log('予約投稿シートのヘッダー行:', data[0]);
  
  for (let i = 1; i < data.length; i++) {
    // 新しい列構造: ID(0), 投稿内容(1), 予定日付(2), 予定時刻(3), ステータス(4), 投稿URL(5), ツリーID(6), 投稿順序(7), 動画URL(8), 画像URL_1枚目～10枚目(9-18), エラー(19)
    const row = data[i];
    const id = row[0] || `auto_${Date.now()}_${i}`; // IDが空の場合は自動生成
    const content = row[1];
    const scheduledDate = row[2];
    const scheduledTime = row[3];
    const status = row[4];
    const postedUrl = row[5];
    const treeId = row[6];
    const postOrder = row[7];
    const videoUrl = row[8];
    const imageUrls = row.slice(9, 19); // 画像URL_1枚目～10枚目
    const error = row[19];
    
    // デバッグ: 生データを出力
    console.log(`行${i + 1} 生データ: ID=${row[0]||'空'}, 日付=${scheduledDate}, 時刻=${scheduledTime}, ステータス="${status}"`);
    console.log(`行${i + 1} メディアURL: 動画="${videoUrl}", 画像=${JSON.stringify(imageUrls)}`);
    
    // 空行をスキップ（コンテンツが空の場合のみ）
    if (!content) {
      console.log(`行${i + 1}: 投稿内容が空のためスキップ`);
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
    
    if (status === 'pending' || status === '投稿予約中' || !status || status === '') {
      const isPastDue = scheduledDateTime <= now;
      console.log(`行${i + 1}: ステータス=${status||'空'}, 過去の予定=${isPastDue}`);
      
      if (isPastDue) {
        // 空でない画像URLのみを収集（最大10枚）
        const validImageUrls = imageUrls.filter(url => url && url.trim() !== '');
        console.log(`行${i + 1}: 有効な画像URL数=${validImageUrls.length}, URLs=${JSON.stringify(validImageUrls)}`);
        
        // videoUrlの検証
        const cleanVideoUrl = videoUrl && videoUrl.trim() !== '' ? videoUrl.trim() : '';
        console.log(`行${i + 1}: 動画URL="${cleanVideoUrl}"`);
        
        posts.push({
          row: i + 1,
          id: id, // 自動生成されたIDまたは既存のID
          content: content,
          imageUrls: validImageUrls,
          videoUrl: cleanVideoUrl,
          scheduledTime: scheduledDateTime,
          retryCount: 0, // リトライ列は削除したため常に0
          treeId: treeId || '',
          postOrder: postOrder || 0
        });
        
        console.log(`行${i + 1}: 予約投稿として処理対象に追加（ID: ${id}）`);
      }
    }
  }
  
  console.log(`getScheduledPosts完了: ${posts.length}件の投稿を返却`);
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
// 投稿メイン関数（作業指示書タスク3.2）
// ===========================
function postToThreads(post) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  // メディア（画像・動画）の有無を判定
  const hasVideo = post.videoUrl && post.videoUrl.trim() !== '';
  const hasImage = post.imageUrls && post.imageUrls.length > 0;
  
  // 返信投稿の場合
  if (post.replyToId) {
    if (hasImage) {
      // 画像付き返信
      if (post.imageUrls.length === 1) {
        return postReplyWithImage(post.content, post.imageUrls[0], post.replyToId);
      } else {
        return postReplyWithMultipleImages(post.content, post.imageUrls, post.replyToId);
      }
    } else {
      // テキストのみ返信
      return postReplyTextOnly(post.content, post.replyToId);
    }
  }
  
  // 通常投稿の場合
  if (hasVideo) {
    // 動画投稿：新しい関数を使用
    return publishMediaPost(accessToken, userId, post.content, post.videoUrl, true);
  } else if (hasImage) {
    if (post.imageUrls.length === 1) {
      // 単一画像投稿：新しい関数を使用
      return publishMediaPost(accessToken, userId, post.content, post.imageUrls[0], false);
    } else {
      // 複数画像投稿：既存の関数を使用（カルーセル形式のため）
      return postWithMultipleImages(post.content, post.imageUrls);
    }
  } else {
    // テキストのみ投稿
    return postTextOnly(post.content);
  }
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
  
  // 投稿データの検証
  if (!post || !post.content) {
    console.error('processPost: 無効な投稿データ', post);
    const errorMsg = 'post.contentが空です';
    updatePostStatus(post.row, 'failed', null, errorMsg);
    logOperation('投稿失敗', 'error', `行: ${post.row}, エラー: ${errorMsg}`);
    return null;
  }
  
  console.log(`processPost開始: 行${post.row}, ID="${post.id}", 内容="${post.content.substring(0, 50)}..."`);
  
  // 作業指示書タスク3.2: postToThreads関数を呼び出す
  const result = postToThreads(post);
  
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
    console.error('認証情報が不足: accessToken=' + !!accessToken + ', userId=' + !!userId);
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  try {
    console.log('postTextOnly 開始 - テキスト:', text.substring(0, 50) + '...');
    console.log('使用するUSER_ID:', userId);
    
    // メディアコンテナの作成
    const createUrl = `${THREADS_API_BASE}/v1.0/${userId}/threads`;
    console.log('APIエンドポイント:', createUrl);
    
    const createResponse = fetchWithTracking(
      createUrl,
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
    console.log('Create API レスポンスコード:', responseCode);
    console.log('Create API レスポンス:', responseText);
    
    if (responseCode !== 200) {
      console.error(`API呼び出し失敗: HTTPステータス ${responseCode}`);
      const errorData = JSON.parse(responseText);
      const errorMsg = errorData.error?.message || errorData.error?.error_user_msg || 
                      errorData.error?.error_user_title || `APIエラー (${responseCode})`;
      logOperation('コンテナ作成失敗', 'error', `HTTPステータス: ${responseCode}, メッセージ: ${errorMsg}`);
      return { success: false, error: errorMsg };
    }
    
    const createResult = JSON.parse(responseText);
    incrementAPICallCount(1); // API使用回数を記録
    
    if (createResult.id) {
      console.log('コンテナ作成成功 - ID:', createResult.id);
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      const errorMsg = createResult.error?.message || createResult.error?.error_user_msg || 
                      createResult.error?.error_user_title || 'コンテナ作成失敗（詳細不明）';
      console.error('コンテナ作成失敗:', errorMsg);
      logOperation('コンテナ作成失敗', 'error', errorMsg);
      return { success: false, error: errorMsg };
    }
    
  } catch (error) {
    console.error('postTextOnly エラー詳細:', error);
    console.error('エラースタック:', error.stack);
    console.error('エラーメッセージ:', error.message);
    
    // URLFetch制限の特定のエラーメッセージを確認
    const errorMessage = error.toString();
    if (errorMessage.includes('UrlFetch') || errorMessage.includes('urlfetch') || errorMessage.includes('サービス')) {
      console.error('URLFetch制限エラーを検出');
      logOperation('API制限エラー', 'error', errorMessage);
      return { success: false, error: `API制限エラー: ${errorMessage}` };
    }
    
    logOperation('投稿例外エラー', 'error', errorMessage);
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
    
    console.log('画像投稿作成中...');
    console.log(`  画像URL: ${publicImageUrl}`);
    console.log(`  テキスト: ${text || 'なし'}`);
    
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
    console.log('画像投稿APIレスポンス:', JSON.stringify(createResult));
    
    if (createResult.id) {
      console.log('メディアコンテナ作成成功：', createResult.id);
      // 少し待機（画像処理のため）
      console.log('画像処理待機中（3秒）...');
      Utilities.sleep(3000);
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      // エラーの詳細をログに記録
      const errorMessage = createResult.error?.message || 
                         createResult.error_message || 
                         createResult.error?.error_user_msg || 
                         'コンテナ作成失敗';
      const errorCode = createResult.error?.code || createResult.error_code || 'N/A';
      console.error(`画像投稿エラー: ${errorMessage} (コード: ${errorCode})`);
      console.error('詳細エラー情報:', JSON.stringify(createResult));
      return { success: false, error: `${errorMessage} (エラーコード: ${errorCode})` };
    }
    
  } catch (error) {
    console.error('投稿エラーの詳細:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    return { success: false, error: errorMessage };
  }
}

// ===========================
// 動画付き投稿は publishMediaPost 関数に統合されました

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
    // 画像URLを公開URLに変換
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
    console.error('投稿エラーの詳細:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    return { success: false, error: errorMessage };
  }
}

// ===========================
// 投稿公開
// ===========================
function publishPost(containerId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  try {
    console.log(`publishPost開始 - containerId: ${containerId}`);
    
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
    
    const responseCode = publishResponse.getResponseCode();
    const responseText = publishResponse.getContentText();
    console.log('Publish API レスポンスコード:', responseCode);
    console.log('Publish API レスポンス:', responseText);
    
    const publishResult = JSON.parse(responseText);
    incrementAPICallCount(1); // API使用回数を記録
    
    if (publishResult.id) {
      // 投稿URL生成（USERNAMEが設定されていない場合の対応）
      let postUrl = `https://www.threads.net/post/${publishResult.id}`;
      const username = getConfig('USERNAME');
      if (username) {
        postUrl = `https://www.threads.net/@${username}/post/${publishResult.id}`;
      }
      
      console.log(`投稿成功 - ID: ${publishResult.id}, URL: ${postUrl}`);
      logOperation('投稿成功', 'success', `投稿ID: ${publishResult.id}, URL: ${postUrl}`);
      
      return { success: true, postUrl: postUrl, postId: publishResult.id };
    } else {
      const errorMsg = publishResult.error?.message || publishResult.error?.error_user_msg || 
                      publishResult.error?.error_user_title || '公開失敗（詳細不明）';
      console.error(`投稿公開失敗: ${errorMsg}`);
      console.error('エラー詳細:', JSON.stringify(publishResult.error || publishResult));
      logOperation('投稿公開失敗', 'error', errorMsg);
      return { success: false, error: errorMsg };
    }
    
  } catch (error) {
    console.error('投稿エラーの詳細:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    logOperation('投稿例外エラー', 'error', errorMessage);
    return { success: false, error: errorMessage };
  }
}

// ===========================
// 投稿のステータスを確認する関数
// ===========================
function verifyPostStatus(postId) {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!postId) {
    console.error('verifyPostStatus: postIdが指定されていません');
    return null;
  }
  
  try {
    console.log(`投稿ステータス確認 - ID: ${postId}`);
    
    // Threads APIで投稿の詳細を取得
    const response = fetchWithTracking(
      `${THREADS_API_BASE}/v1.0/${postId}?fields=id,text,timestamp,permalink`,
      {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('投稿確認APIレスポンスコード:', responseCode);
    console.log('投稿確認APIレスポンス:', responseText);
    
    if (responseCode === 200) {
      const postData = JSON.parse(responseText);
      console.log('投稿が正常に確認されました:', postData);
      return postData;
    } else {
      console.error('投稿の確認に失敗しました:', responseText);
      return null;
    }
  } catch (error) {
    console.error('投稿ステータス確認エラー:', error);
    return null;
  }
}

// ===========================
// デバッグ用：投稿処理の詳細確認
// ===========================
function debugPostExecution() {
  console.log('===== デバッグ: 投稿処理の詳細確認 =====');
  
  // 設定値の確認
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  const username = getConfig('USERNAME');
  
  console.log('ACCESS_TOKEN設定:', accessToken ? '設定済み（長さ:' + accessToken.length + '文字）' : '未設定');
  console.log('USER_ID設定:', userId || '未設定');
  console.log('USERNAME設定:', username || '未設定');
  
  // シートからテスト投稿データを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  if (!sheet) {
    console.error('予約投稿シートが見つかりません');
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  console.log('シート内のデータ行数:', values.length);
  
  // ヘッダー行をスキップして最初のペンディング投稿を探す
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[4]; // ステータス列
    
    if (status === '投稿予約中' || status === 'pending') {
      console.log(`\nテスト対象行: ${i + 1}`);
      console.log('投稿内容:', row[1]); // 投稿内容列
      console.log('予定日付:', row[2]); // 予定日付列
      console.log('予定時刻:', row[3]); // 予定時刻列
      
      // テスト投稿を実行
      const testPost = {
        row: i + 1,
        id: row[0] || `test_${Date.now()}`,
        content: row[1],
        imageUrls: [],
        videoUrl: row[8] || '', // 動画URL列
        retryCount: 0
      };
      
      // 画像URLを収集（列10-19）
      for (let j = 9; j < 19; j++) {
        if (row[j] && row[j].trim() !== '') {
          testPost.imageUrls.push(row[j]);
        }
      }
      
      console.log('\nテスト投稿データ:', JSON.stringify(testPost, null, 2));
      console.log('\n投稿処理を実行中...');
      
      const result = postToThreads(testPost);
      
      console.log('\n投稿結果:', JSON.stringify(result, null, 2));
      
      // 結果に基づいてステータスを更新
      if (result.success) {
        updatePostStatus(testPost.row, 'posted', result.postUrl, null);
        console.log('✅ 投稿成功！URL:', result.postUrl);
        
        // 投稿の確認
        if (result.postId) {
          const verification = verifyPostStatus(result.postId);
          if (verification) {
            console.log('✅ 投稿の存在が確認されました');
          } else {
            console.log('⚠️ 投稿IDは取得できましたが、投稿の確認に失敗しました');
          }
        }
      } else {
        console.error('❌ 投稿失敗:', result.error);
        updatePostStatus(testPost.row, 'failed', null, result.error);
      }
      
      break; // 最初の1件のみテスト
    }
  }
  
  console.log('\n===== デバッグ完了 =====');
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
    console.error('投稿エラーの詳細:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    return { success: false, error: errorMessage };
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
    
    console.log('画像付き返信作成中...');
    console.log(`  返信先ID: ${replyToId}`);
    console.log(`  画像URL: ${publicImageUrl}`);
    console.log(`  テキスト: ${text || 'なし'}`);
    
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
    console.log('画像付き返信APIレスポンス:', JSON.stringify(createResult));
    
    if (createResult.id) {
      console.log('返信コンテナ作成成功：', createResult.id);
      // 少し待機（画像処理のため）
      Utilities.sleep(3000);
      
      // 投稿の公開
      return publishPost(createResult.id);
    } else {
      // エラーの詳細をログに記録
      const errorMessage = createResult.error?.message || 
                         createResult.error_message || 
                         createResult.error?.error_user_msg || 
                         'コンテナ作成失敗';
      const errorCode = createResult.error?.code || createResult.error_code || 'N/A';
      console.error(`画像付き返信エラー: ${errorMessage} (コード: ${errorCode})`);
      console.error('詳細エラー情報:', JSON.stringify(createResult));
      return { success: false, error: `${errorMessage} (エラーコード: ${errorCode})` };
    }
    
  } catch (error) {
    console.error('投稿エラーの詳細:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    return { success: false, error: errorMessage };
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
    // 画像URLを公開URLに変換
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
    console.error('投稿エラーの詳細:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    return { success: false, error: errorMessage };
  }
}

// ===========================
// メディアコンテナ作成（作業指示書タスク3.3）
// ===========================
function createMediaContainer(accessToken, userId, mediaUrl, isVideo, text = '') {
  console.log(`createMediaContainer開始: mediaUrl=${mediaUrl}, isVideo=${isVideo}, text=${text}`);
  
  // URLの検証
  if (!mediaUrl || mediaUrl.trim() === '') {
    console.error('メディアURLが空です');
    return null;
  }
  
  // 動画の場合は専用のURL変換処理を行う
  let publicUrl;
  if (isVideo) {
    // 動画URLの変換（Google Driveの場合は公開URLに変換）
    publicUrl = convertVideoUrl(mediaUrl);
    console.log(`動画URL変換後: ${publicUrl}`);
  } else {
    // 画像URLの変換
    publicUrl = convertToPublicUrl(mediaUrl);
    console.log(`画像URL変換後: ${publicUrl}`);
  }
  
  // メディアタイプに応じたパラメータ設定
  const mediaType = isVideo ? 'VIDEO' : 'IMAGE';
  const urlParam = isVideo ? 'video_url' : 'image_url';
  
  console.log(`メディアコンテナ作成: ${mediaType}`);
  console.log(`URL: ${publicUrl}`);
  
  // APIペイロードの構築
  const payload = {
    media_type: mediaType,
    [urlParam]: publicUrl
  };
  
  // テキストがある場合は追加
  if (text && text.trim() !== '') {
    payload.text = text;
  }
  
  console.log('APIペイロード:', JSON.stringify(payload));
  
  try {
    // API呼び出し
    const response = fetchWithTracking(
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
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    console.log(`APIレスポンス - コード: ${responseCode}`);
    console.log(`APIレスポンス - 内容: ${responseText}`);
    
    // レスポンスコードのチェック
    if (responseCode >= 400) {
      console.error(`HTTPエラー ${responseCode}: ${responseText}`);
      
      // エラーレスポンスを解析
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.error) {
          const errorCode = errorData.error.code;
          const errorSubcode = errorData.error.error_subcode;
          const userTitle = errorData.error.error_user_title || '';
          const userMsg = errorData.error.error_user_msg || '';
          
          // 動画関連のエラーメッセージ
          let errorMsg = `メディアの処理に失敗しました。\n`;
          
          // メディアダウンロードエラーの場合
          if (errorSubcode === 2207052 || errorSubcode === 2207005) {
            if (isVideo) {
              errorMsg = '動画の処理に失敗しました。\n\n';
              errorMsg += '考えられる原因：\n';
              errorMsg += '1. 動画フォーマットが非対応（対応形式: MP4, MOV）\n';
              errorMsg += '2. 動画サイズが1GBを超えている\n';
              errorMsg += '3. 動画の長さが5分を超えている\n';
              errorMsg += '4. URLからアクセスできない\n';
              
              if (publicUrl.includes('drive.google.com')) {
                errorMsg += '5. Google Driveの共有設定が「制限付き」になっている\n\n';
                errorMsg += '解決方法：\n';
                errorMsg += '• Google Driveで動画を右クリック→「共有」→「リンクを知っている全員」に変更\n';
                errorMsg += '• または、動画を他のホスティングサービス（YouTube、Vimeo等）にアップロード';
              }
            } else {
              // 画像の場合の既存エラーメッセージ
              if (publicUrl.includes('drive.google.com')) {
                errorMsg = 'Google Driveの画像の投稿に失敗しました。\n\n';
                errorMsg += '考えられる原因：\n';
                errorMsg += '1. ファイルの共有設定が「制限付き」になっている\n';
                errorMsg += '2. 画像フォーマットが認識できない\n\n';
                errorMsg += '解決方法：\n';
                errorMsg += '• Google Driveでファイルを右クリック→「共有」→「リンクを知っている全員」に変更\n';
                errorMsg += '• または、画像を他のホスティングサービス（Imgur等）にアップロード';
              }
            }
          } else {
            // その他のエラー
            errorMsg += `エラーコード: ${errorCode}\n`;
            errorMsg += `詳細: ${userMsg || errorData.error.message}`;
          }
          
          console.error(errorMsg);
          logError('createMediaContainer', new Error(errorMsg));
          return null;
        }
      } catch (e) {
        // JSON解析エラーは無視
      }
      
      logError('createMediaContainer', new Error(`HTTP ${responseCode}: ${responseText}`));
      return null;
    }
    
    const result = JSON.parse(responseText);
    
    if (result.id) {
      console.log(`コンテナ作成成功: ${result.id}`);
      return result.id;
    } else {
      const errorMsg = result.error?.message || result.error_message || 'コンテナ作成失敗';
      console.error(`コンテナ作成エラー: ${errorMsg}`);
      console.error('エラー詳細:', JSON.stringify(result));
      logError('createMediaContainer', new Error(errorMsg));
      return null;
    }
    
  } catch (error) {
    console.error('メディアコンテナ作成例外 - タイプ:', typeof error);
    console.error('メディアコンテナ作成例外 - メッセージ:', error.toString());
    console.error('メディアコンテナ作成例外 - スタック:', error.stack);
    
    // エラーの詳細情報を取得
    if (error.message) {
      console.error('エラーメッセージ:', error.message);
    }
    if (error.name) {
      console.error('エラー名:', error.name);
    }
    
    logError('createMediaContainer', error);
    return null;
  }
}

// ===========================
// 動画ステータス確認（作業指示書タスク3.4）
// ===========================
function waitForVideoProcessing(accessToken, containerId) {
  const maxWaitTime = 5 * 60 * 1000; // 5分
  const checkInterval = 5000; // 5秒ごと
  const startTime = new Date().getTime();
  
  console.log(`動画処理待機開始: ${containerId}`);
  
  // 初回は30秒待機（API仕様に基づく）
  console.log('初期待機中（30秒）...');
  Utilities.sleep(30000);
  
  while (true) {
    try {
      // ステータス確認
      const response = fetchWithTracking(
        `${THREADS_API_BASE}/v1.0/${containerId}?fields=status_code,error_message`,
        {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`
          },
          muteHttpExceptions: true
        }
      );
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode >= 400) {
        console.error(`ステータス確認エラー: ${responseCode}`);
        return false;
      }
      
      const result = JSON.parse(responseText);
      console.log(`動画ステータス: ${result.status_code}`);
      
      if (result.status_code === 'FINISHED') {
        console.log('動画処理完了');
        return true;
      } else if (result.status_code === 'ERROR') {
        console.error('動画処理エラー:', result.error_message || 'Unknown error');
        logError('waitForVideoProcessing', new Error(`動画処理エラー: ${result.error_message}`));
        return false;
      }
      
      // タイムアウトチェック
      const elapsedTime = new Date().getTime() - startTime;
      if (elapsedTime > maxWaitTime) {
        console.error('動画処理タイムアウト');
        logError('waitForVideoProcessing', new Error('動画処理がタイムアウトしました'));
        return false;
      }
      
      console.log(`処理中... (経過時間: ${Math.floor(elapsedTime / 1000)}秒)`);
      
      // 待機
      Utilities.sleep(checkInterval);
      
    } catch (error) {
      console.error('動画ステータス確認エラー:', error);
      logError('waitForVideoProcessing', error);
      return false;
    }
  }
}

// ===========================
// メディア投稿公開（作業指示書タスク3.5）
// ===========================
function publishMediaPost(accessToken, userId, text, mediaUrl, isVideo) {
  try {
    console.log(`publishMediaPost開始: isVideo=${isVideo}, mediaUrl=${mediaUrl}`);
    
    // メディアコンテナを作成（テキストも一緒に送信）
    const containerId = createMediaContainer(accessToken, userId, mediaUrl, isVideo, text);
    
    if (!containerId) {
      console.error('メディアコンテナの作成に失敗しました');
      
      // より詳細なエラーメッセージを提供
      let errorMsg = 'メディアコンテナの作成に失敗しました。';
      
      if (mediaUrl.includes('drive.google.com')) {
        if (isVideo) {
          errorMsg = '動画の投稿に失敗しました。\n\n' +
                     '考えられる原因：\n' +
                     '1. 動画フォーマットが非対応（対応形式: MP4, MOV）\n' +
                     '2. 動画サイズが1GBを超えている\n' +
                     '3. 動画の長さが5分を超えている\n' +
                     '4. Google Driveの共有設定が「制限付き」になっている\n\n' +
                     '解決方法：\n' +
                     '• Google Driveで動画を右クリック→「共有」→「リンクを知っている全員」に変更\n' +
                     '• 動画の仕様を確認（MP4/MOV、1GB以下、5分以内）\n' +
                     '• または、動画を他のホスティングサービスにアップロード';
        } else {
          errorMsg = 'Google Driveの画像/動画の投稿に失敗しました。\n\n' +
                     '考えられる原因：\n' +
                     '1. ファイルの共有設定が「制限付き」になっている\n' +
                     '2. 画像/動画フォーマットが認識できない\n\n' +
                     '解決方法：\n' +
                     '• Google Driveでファイルを右クリック→「共有」→「リンクを知っている全員」に変更\n' +
                     '• または、Imgur、Cloudinary等の画像ホスティングサービスを使用';
        }
      } else {
        errorMsg += ' URLが有効であることを確認してください。';
      }
      
      return { success: false, error: errorMsg };
    }
    
    // 動画の場合は処理完了を待機
    if (isVideo) {
      const isReady = waitForVideoProcessing(accessToken, containerId);
      if (!isReady) {
        return { success: false, error: '動画処理が完了しませんでした。動画の仕様を確認してください。' };
      }
    } else {
      // 画像の場合も少し待機
      console.log('画像処理待機中（3秒）...');
      Utilities.sleep(3000);
    }
    
    // 投稿を公開
    console.log('投稿を公開中...');
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
    
    const publishResponseCode = publishResponse.getResponseCode();
    const publishResponseText = publishResponse.getContentText();
    
    console.log(`公開APIレスポンス - コード: ${publishResponseCode}`);
    console.log(`公開APIレスポンス - 内容: ${publishResponseText}`);
    
    if (publishResponseCode >= 400) {
      const errorData = JSON.parse(publishResponseText);
      const errorMsg = errorData.error?.message || '投稿公開失敗';
      console.error(`投稿公開エラー: ${errorMsg}`);
      return { success: false, error: errorMsg };
    }
    
    const publishResult = JSON.parse(publishResponseText);
    
    if (publishResult.id) {
      const username = getConfig('USERNAME');
      const postUrl = `https://www.threads.net/@${username}/post/${publishResult.id}`;
      console.log(`投稿公開成功: ${postUrl}`);
      return { success: true, postUrl: postUrl, postId: publishResult.id };
    } else {
      const errorMsg = publishResult.error?.message || '投稿公開失敗';
      console.error(`投稿公開エラー: ${errorMsg}`);
      return { success: false, error: errorMsg };
    }
    
  } catch (error) {
    console.error('メディア投稿公開例外:', error);
    console.error('エラースタック:', error.stack);
    const errorMessage = error.message || error.toString() || 'Unknown error';
    return { success: false, error: errorMessage };
  }
}

// ===========================
// URL変換
// ===========================
function convertToPublicUrl(url) {
  if (!url || url.trim() === '') {
    return url;
  }
  
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
      console.log(`Google DriveファイルID検出: ${fileId}`);
      
      // lh3.googleusercontent.com形式に変換
      // この形式はThreads APIで画像が確実に表示される
      const directUrl = `https://lh3.googleusercontent.com/d/${fileId}`;
      console.log(`変換後のURL: ${directUrl}`);
      return directUrl;
    }
  }
  
  // Google Drive以外のURLはそのまま返す
  return url;
}

// ===========================
// 動画ファイル判定
// ===========================
function isVideoFile(url) {
  if (!url || url.trim() === '') {
    return false;
  }
  
  // URLからパラメータを除去
  const urlWithoutParams = url.split('?')[0];
  const extension = urlWithoutParams.split('.').pop().toLowerCase();
  
  // Threads APIでサポートされている動画フォーマット
  const supportedVideoExtensions = ['mp4', 'mov'];
  
  // 拡張子で判定
  if (supportedVideoExtensions.includes(extension)) {
    console.log(`動画ファイル検出: ${extension}形式`);
    return true;
  }
  
  // Google ドライブのURLの場合、ファイル名から判定できない
  if (url.includes('drive.google.com')) {
    // ユーザーが明示的に動画として指定した場合は動画として扱う
    console.log('Google Driveファイル: 動画として処理を試みます');
    return true; // 動画として試すが、実際にはAPIでエラーになる可能性がある
  }
  
  return false;
}

// ===========================
// 動画仕様のバリデーション
// ===========================
function validateVideoSpecs(url) {
  const warnings = [];
  
  // URLからファイル名を取得
  const urlWithoutParams = url.split('?')[0];
  const fileName = urlWithoutParams.split('/').pop();
  
  // 拡張子チェック
  const extension = fileName.split('.').pop().toLowerCase();
  if (!['mp4', 'mov'].includes(extension)) {
    warnings.push('動画フォーマットはMP4またはMOVである必要があります');
  }
  
  // Google Driveの場合の警告
  if (url.includes('drive.google.com')) {
    warnings.push('Google Driveの動画は直接投稿できない可能性があります。他のホスティングサービスの使用を推奨します');
  }
  
  // YouTubeの場合の警告
  if (url.includes('youtube.com') || url.includes('youtu.be')) {
    warnings.push('YouTube URLは使用できません。動画ファイルの直接URLが必要です');
  }
  
  if (warnings.length > 0) {
    console.warn('動画仕様の警告:');
    warnings.forEach(warning => console.warn(`- ${warning}`));
  }
  
  return warnings;
}

// ===========================
// 動画投稿のテスト関数
// ===========================
function testVideoPost() {
  console.log('=== 動画投稿テスト開始 ===');
  
  const testCases = [
    {
      name: 'MP4動画（公開サーバー）',
      url: 'https://example.com/test.mp4',
      text: 'テスト投稿: MP4動画',
      expectedResult: 'success'
    },
    {
      name: 'MOV動画（公開サーバー）',
      url: 'https://example.com/test.mov',
      text: 'テスト投稿: MOV動画',
      expectedResult: 'success'
    },
    {
      name: 'Google Drive動画',
      url: 'https://drive.google.com/file/d/1234567890/view',
      text: 'テスト投稿: Google Drive動画',
      expectedResult: 'error'
    },
    {
      name: 'YouTube URL',
      url: 'https://youtube.com/watch?v=12345',
      text: 'テスト投稿: YouTube動画',
      expectedResult: 'error'
    },
    {
      name: '非対応フォーマット',
      url: 'https://example.com/test.avi',
      text: 'テスト投稿: AVI動画',
      expectedResult: 'error'
    }
  ];
  
  const results = [];
  
  testCases.forEach((testCase, index) => {
    console.log(`\nテストケース${index + 1}: ${testCase.name}`);
    console.log(`URL: ${testCase.url}`);
    console.log(`期待結果: ${testCase.expectedResult}`);
    
    // 動画仕様のバリデーション
    const warnings = validateVideoSpecs(testCase.url);
    if (warnings.length > 0) {
      console.log('警告:', warnings.join(', '));
    }
    
    // 動画ファイル判定
    const isVideo = isVideoFile(testCase.url);
    console.log(`動画ファイル判定: ${isVideo}`);
    
    results.push({
      testCase: testCase.name,
      warnings: warnings,
      isVideo: isVideo,
      expectedResult: testCase.expectedResult
    });
  });
  
  console.log('\n=== テスト結果サマリー ===');
  results.forEach(result => {
    console.log(`${result.testCase}: ${result.warnings.length > 0 ? '警告あり' : 'OK'}`);
  });
  
  return results;
}

// ===========================
// 手動動画投稿テスト
// ===========================
function manualTestVideoPost(videoUrl, text) {
  console.log('=== 手動動画投稿テスト ===');
  console.log(`動画URL: ${videoUrl}`);
  console.log(`テキスト: ${text || 'なし'}`);
  
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    console.error('認証情報が設定されていません');
    return { success: false, error: '認証情報が見つかりません' };
  }
  
  // 動画仕様の検証
  const warnings = validateVideoSpecs(videoUrl);
  if (warnings.length > 0) {
    console.warn('=== 警告 ===');
    warnings.forEach(warning => console.warn(warning));
  }
  
  // 動画ファイル判定
  const isVideo = isVideoFile(videoUrl);
  console.log(`動画ファイルとして処理: ${isVideo ? 'はい' : 'いいえ'}`);
  
  if (!isVideo) {
    console.error('指定されたURLは動画ファイルとして認識されませんでした');
    return { success: false, error: '動画ファイルではありません' };
  }
  
  // 投稿実行
  console.log('投稿を実行中...');
  const result = publishMediaPost(accessToken, userId, text || '', videoUrl, true);
  
  if (result.success) {
    console.log('=== 投稿成功 ===');
    console.log(`投稿URL: ${result.postUrl}`);
    console.log(`投稿ID: ${result.postId}`);
  } else {
    console.error('=== 投稿失敗 ===');
    console.error(`エラー: ${result.error}`);
  }
  
  return result;
}

// ===========================
// 動画URL変換
// ===========================
function convertVideoUrl(url) {
  if (!url || url.trim() === '') {
    return url;
  }
  
  const trimmedUrl = url.trim();
  
  // Google DriveのURLパターンをチェック
  const drivePatterns = [
    /drive\.google\.com\/file\/d\/([a-zA-Z0-9-_]+)/,
    /drive\.google\.com\/open\?id=([a-zA-Z0-9-_]+)/,
    /docs\.google\.com\/.*\/d\/([a-zA-Z0-9-_]+)/
  ];
  
  for (const pattern of drivePatterns) {
    const match = trimmedUrl.match(pattern);
    if (match) {
      const fileId = match[1];
      console.log(`Google Drive動画ファイルID検出: ${fileId}`);
      
      // 注意: Google Driveの動画は直接Threads APIで使用できない可能性が高い
      // ユーザーに警告を出すことを推奨
      console.warn('警告: Google Driveの動画はThreads APIで直接使用できない可能性があります。');
      console.warn('推奨: 動画を公開サーバー（YouTube、Vimeo等）にアップロードしてください。');
      
      // それでも試す場合のURL形式
      const directUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      console.log(`動画変換後のURL: ${directUrl}`);
      return directUrl;
    }
  }
  
  // YouTube URLの場合の処理
  if (trimmedUrl.includes('youtube.com') || trimmedUrl.includes('youtu.be')) {
    console.warn('警告: YouTube URLは直接使用できません。動画ファイルの直接URLが必要です。');
  }
  
  // その他の動画ホスティングサービスのチェック
  const supportedVideoHosts = [
    'cloudinary.com',
    'res.cloudinary.com',
    'vimeo.com',
    'wistia.com',
    'streamable.com'
  ];
  
  const isSupported = supportedVideoHosts.some(host => trimmedUrl.includes(host));
  if (isSupported) {
    console.log('対応動画ホスティングサービスのURLです');
  }
  
  // Google Drive以外のURLはそのまま返す
  return trimmedUrl;
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
  const now = new Date();
  console.log('===== 予約投稿データ確認 =====');
  console.log(`データ行数: ${data.length}`);
  console.log(`現在時刻: ${Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss')}`);
  
  let message = '予約投稿シートの状態:\n\n';
  let pendingCount = 0;
  let postedCount = 0;
  let failedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[4];
    
    if (status === '投稿予約中' || status === 'pending') pendingCount++;
    else if (status === '投稿済' || status === 'posted') postedCount++;
    else if (status === '失敗' || status === 'failed') failedCount++;
  }
  
  message += `総データ数: ${data.length - 1}件\n`;
  message += `投稿予約中: ${pendingCount}件\n`;
  message += `投稿済: ${postedCount}件\n`;
  message += `失敗: ${failedCount}件\n\n`;
  
  if (pendingCount > 0) {
    message += '投稿予約中の詳細:\n';
    for (let i = 1; i < Math.min(data.length, 6); i++) {
      const row = data[i];
      const status = row[4];
      if (status === '投稿予約中' || status === 'pending') {
        message += `行${i + 1}: ${row[1] ? row[1].substring(0, 30) + '...' : '内容なし'}\n`;
      }
    }
  }
  
  ui.alert('データ確認', message, ui.ButtonSet.OK);
}

// ===========================
// API使用状況確認（デバッグ用）
// ===========================
function checkAPIUsageStatus() {
  const ui = SpreadsheetApp.getUi();
  const quota = checkDailyAPIQuota();
  
  const message = `API使用状況レポート\n\n` +
    `本日の使用回数: ${quota.count.toLocaleString()}\n` +
    `残り使用可能回数: ${quota.remaining.toLocaleString()}\n` +
    `1日の制限: 20,000回\n\n` +
    `${quota.isNearLimit ? '⚠️ 警告: API制限に近づいています！' : '✅ 正常範囲内です'}`;
  
  ui.alert('API使用状況', message, ui.ButtonSet.OK);
}

// ===========================
// 手動でAPI使用回数をリセット（緊急用）
// ===========================
function resetAPIQuotaManually() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'API使用回数リセット',
    'API使用回数のカウンターをリセットしますか？\n\n' +
    '⚠️ 注意: これは緊急時のみ使用してください。',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('API_CALL_COUNT', '0');
    scriptProperties.setProperty('API_QUOTA_RESET_DATE', new Date().toDateString());
    
    ui.alert('完了', 'API使用回数をリセットしました。', ui.ButtonSet.OK);
    logOperation('API使用回数リセット', 'info', '手動リセット実行');
  }
}

// ===========================
// 予約投稿の強制実行（過去の投稿も含む）
// ===========================
function forceProcessScheduledPosts() {
  console.log('===== 手動実行: 予約投稿の処理 =====');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  if (!sheet) {
    console.error('予約投稿シートが見つかりません');
    SpreadsheetApp.getUi().alert('エラー', '予約投稿シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const posts = [];
  const now = new Date();
  
  console.log(`現在時刻: ${Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss')}`);
  
  // 投稿予約中のデータを収集（時刻チェックなし）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0] || `manual_${Date.now()}_${i}`; // IDが空の場合は自動生成
    const content = row[1];
    const status = row[4];
    
    // 内容があり、ステータスが投稿予約中または空
    if (content && (status === '投稿予約中' || status === 'pending' || !status || status === '')) {
      const post = {
        row: i + 1,
        id: id,
        content: content,
        imageUrls: [],
        videoUrl: row[8] || '',
        retryCount: 0,
        treeId: row[6] || '',
        postOrder: row[7] || 0
      };
      
      // 画像URLを収集
      for (let j = 9; j < 19; j++) {
        if (row[j] && row[j].trim() !== '') {
          post.imageUrls.push(row[j]);
        }
      }
      
      posts.push(post);
      console.log(`行${i + 1}: 処理対象として追加（ID: ${id}）`);
    }
  }
  
  if (posts.length === 0) {
    const message = '処理対象の投稿がありません。\n\n以下を確認してください:\n' +
                    '1. 投稿内容列に値がある\n' +
                    '2. ステータスが「投稿予約中」または空';
    console.log(message);
    SpreadsheetApp.getUi().alert('情報', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  console.log(`${posts.length}件の投稿を処理します`);
  logOperation('手動投稿処理', 'start', `${posts.length}件の投稿を処理開始`);
  
  let successCount = 0;
  let failCount = 0;
  
  // 投稿をツリーごとにグループ化
  const postGroups = groupPostsByTree(posts);
  
  // グループごとに処理
  Object.entries(postGroups).forEach(([groupId, groupPosts]) => {
    try {
      if (groupPosts.length === 1 && !groupPosts[0].treeId) {
        // 単発投稿
        const result = processPost(groupPosts[0]);
        if (result && result.success) {
          successCount++;
        } else {
          failCount++;
        }
      } else {
        // ツリー投稿
        processTreePosts(groupPosts);
        successCount += groupPosts.length;
      }
    } catch (error) {
      console.error(`グループ${groupId}の処理エラー:`, error);
      failCount += groupPosts.length;
    }
  });
  
  const resultMessage = `処理完了\n\n成功: ${successCount}件\n失敗: ${failCount}件`;
  console.log(resultMessage);
  logOperation('手動投稿処理', 'complete', resultMessage);
  SpreadsheetApp.getUi().alert('処理結果', resultMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===========================
// 予約投稿のデバッグ実行
// ===========================
function debugScheduledPosts() {
  console.log('===== デバッグ: 予約投稿シートの状態確認 =====');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約投稿');
  if (!sheet) {
    console.error('予約投稿シートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  console.log(`現在時刻（JST）: ${Utilities.formatDate(now, 'JST', 'yyyy/MM/dd HH:mm:ss')}`);
  console.log(`データ行数: ${data.length}`);
  console.log('\nヘッダー行:', data[0]);
  
  if (data.length <= 1) {
    console.log('データ行がありません（ヘッダーのみ）');
    return;
  }
  
  console.log('\n===== 全データ行の詳細 =====');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    console.log(`\n--- 行${i + 1} ---`);
    console.log('ID:', row[0] || '（空）');
    console.log('投稿内容:', row[1] ? row[1].substring(0, 50) + '...' : '（空）');
    console.log('予定日付:', row[2] || '（空）');
    console.log('予定時刻:', row[3] || '（空）');
    console.log('ステータス:', row[4] || '（空）');
    console.log('投稿URL:', row[5] || '（空）');
    console.log('ツリーID:', row[6] || '（空）');
    console.log('投稿順序:', row[7] || '（空）');
    console.log('動画URL:', row[8] || '（空）');
    
    // 画像URLの確認
    const imageUrls = [];
    for (let j = 9; j < 19; j++) {
      if (row[j] && row[j].trim() !== '') {
        imageUrls.push(row[j]);
      }
    }
    console.log('画像URL数:', imageUrls.length);
    
    console.log('エラー:', row[19] || '（空）');
    
    // 処理対象かどうかの判定
    const id = row[0];
    const content = row[1];
    const status = row[4];
    const scheduledDate = row[2];
    const scheduledTime = row[3];
    
    if (!id || !content) {
      console.log('→ スキップ理由: IDまたは投稿内容が空');
      continue;
    }
    
    if (!scheduledDate || !scheduledTime) {
      console.log('→ スキップ理由: 予定日付または予定時刻が空');
      continue;
    }
    
    // 日時の解析を試みる
    try {
      let scheduledDateTime;
      if (scheduledDate instanceof Date) {
        const dateObj = new Date(scheduledDate);
        const year = dateObj.getFullYear();
        const month = dateObj.getMonth();
        const day = dateObj.getDate();
        
        let hours, minutes;
        if (scheduledTime instanceof Date) {
          hours = scheduledTime.getHours();
          minutes = scheduledTime.getMinutes();
        } else {
          const timeStr = scheduledTime.toString();
          [hours, minutes] = timeStr.split(':').map(s => parseInt(s));
        }
        
        scheduledDateTime = new Date(year, month, day, hours, minutes, 0);
      } else {
        // 文字列として処理
        const dateStr = scheduledDate.toString();
        const timeStr = scheduledTime.toString();
        scheduledDateTime = new Date(`${dateStr} ${timeStr}`);
      }
      
      console.log('→ 予定日時:', Utilities.formatDate(scheduledDateTime, 'JST', 'yyyy/MM/dd HH:mm:ss'));
      console.log('→ 現在時刻との比較:', scheduledDateTime <= now ? '過去（処理対象）' : '未来（まだ処理しない）');
    } catch (e) {
      console.log('→ スキップ理由: 日時の解析エラー', e.toString());
      continue;
    }
    
    // ステータスの確認
    console.log('→ ステータス判定:');
    console.log('  - 現在のステータス:', status);
    console.log('  - "pending"と一致:', status === 'pending');
    console.log('  - "投稿予約中"と一致:', status === '投稿予約中');
    
    if (status === 'pending' || status === '投稿予約中') {
      console.log('→ ✅ 処理対象の可能性あり（ステータスが条件に一致）');
    } else {
      console.log('→ ❌ スキップ理由: ステータスが"pending"または"投稿予約中"ではない');
    }
  }
  
  // 実際に処理対象となる投稿を取得
  console.log('\n===== getScheduledPosts()の実行結果 =====');
  const posts = getScheduledPosts();
  console.log(`処理対象の投稿数: ${posts.length}`);
  
  if (posts.length > 0) {
    console.log('\n処理対象の投稿:');
    posts.forEach((post, index) => {
      console.log(`${index + 1}. 行${post.row}: ${post.content.substring(0, 30)}...`);
    });
  }
  
  console.log('\n===== デバッグ完了 =====');
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
  // 現在のThreads APIではリフレッシュトークンは不要のため、この関数は何もしない
  console.log('refreshAccessToken: 現在のAPIではリフレッシュトークンは不要です');
  logOperation('トークンリフレッシュ', 'info', '現在のAPIではリフレッシュトークンは不要のため、スキップしました');
}

// ===========================
// 不要なトリガーの削除
// ===========================
function removeRefreshTokenTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'refreshAccessToken') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
      console.log('refreshAccessTokenトリガーを削除しました');
    }
  });
  
  if (removedCount > 0) {
    logOperation('トリガー削除', 'success', `${removedCount}個のrefreshAccessTokenトリガーを削除しました`);
    SpreadsheetApp.getUi().alert(`${removedCount}個の不要なトリガーを削除しました`);
  } else {
    SpreadsheetApp.getUi().alert('削除対象のトリガーは見つかりませんでした');
  }
}
