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
    'エラー',
    'リトライ',
    '画像URL_1枚目',
    '画像URL_2枚目',
    '画像URL_3枚目',
    '画像URL_4枚目',
    '画像URL_5枚目',
    '画像URL_6枚目',
    '画像URL_7枚目',
    '画像URL_8枚目',
    '画像URL_9枚目',
    '画像URL_10枚目'
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
    [1, '【サンプル投稿1】\\nこれはテスト投稿です。\\n#threads #自動投稿', 
     Utilities.formatDate(futureDate, 'JST', 'yyyy/MM/dd'), '10:00',
     '投稿予約中', '', '', 0, '', '', '', '', '', '', '', '', '', ''],
    [2, '【サンプル投稿2】\\n画像付き投稿のサンプルです。\\n画像URLを右側の列に入力してください。', 
     Utilities.formatDate(futureDate2, 'JST', 'yyyy/MM/dd'), '15:00',
     '投稿予約中', '', '', 0, 'https://example.com/image1.jpg', '', '', '', '', '', '', '', '', ''],
    [3, '【サンプル投稿3】\\n複数画像の投稿サンプル', 
     Utilities.formatDate(futureDate3, 'JST', 'yyyy/MM/dd'), '12:00',
     '投稿予約中', '', '', 0, 'https://example.com/image1.jpg', 'https://example.com/image2.jpg', '', '', '', '', '', '', '', '']
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 50);   // ID
  sheet.setColumnWidth(2, 400);  // 投稿内容
  sheet.setColumnWidth(3, 100);  // 予定日付
  sheet.setColumnWidth(4, 80);   // 予定時刻
  sheet.setColumnWidth(5, 100);  // ステータス
  sheet.setColumnWidth(6, 300);  // 投稿URL
  sheet.setColumnWidth(7, 200);  // エラー
  sheet.setColumnWidth(8, 80);   // リトライ
  
  // 画像URL列の幅
  for (let i = 9; i <= 18; i++) {
    sheet.setColumnWidth(i, 200);
  }
  
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
    '投稿内容: Threadsに投稿するテキスト（改行は\\\\nで入力）\\n' +
    '予定日付: 投稿する日付（yyyy/mm/dd形式）\\n' +
    '予定時刻: 投稿する時刻（24時間表記、例: 13:30）\\n' +
    'ステータス: 投稿予約中、投稿済、失敗、キャンセル\\n' +
    '投稿URL: 投稿後のURL（自動入力）\\n' +
    'エラー: エラーメッセージ（自動入力）\\n' +
    'リトライ: 再試行回数（自動入力）\\n' +
    '画像URL: 投稿に添付する画像のURL（最大10枚）'
  );
  
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
    
    logOperation('予約投稿処理開始', 'info', `${posts.length}件の投稿を処理`);
    
    posts.forEach(post => {
      try {
        processPost(post);
      } catch (error) {
        logError('processPost', error);
        updatePostStatus(post.row, 'failed', null, error.toString());
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
    // 新しい列構造: ID, 投稿内容, 予定日付, 予定時刻, ステータス, 投稿URL, エラー, リトライ, 画像URL_1枚目〜10枚目
    const [id, content, scheduledDate, scheduledTime, status, postedUrl, error, retryCount, ...imageUrls] = data[i];
    
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
        const validImageUrls = imageUrls.slice(0, 10).filter(url => url && url.trim() !== '');
        
        posts.push({
          row: i + 1,
          id: id,
          content: content,
          imageUrls: validImageUrls,
          scheduledTime: scheduledDateTime,
          retryCount: retryCount || 0
        });
        
        console.log(`行${i + 1}: 予約投稿として処理対象に追加`);
      }
    }
  }
  
  return posts;
}

// ===========================
// 投稿処理
// ===========================
function processPost(post) {
  const maxRetries = 3;
  
  if (post.retryCount >= maxRetries) {
    updatePostStatus(post.row, 'failed', null, '最大リトライ回数に達しました');
    return;
  }
  
  let result;
  
  if (post.imageUrls && post.imageUrls.length > 0) {
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
  
  if (result.success) {
    updatePostStatus(post.row, 'posted', result.postUrl, null);
    logOperation('投稿成功', 'success', `ID: ${post.id}, URL: ${result.postUrl}`);
  } else {
    const newRetryCount = post.retryCount + 1;
    updatePostStatus(post.row, 'pending', null, result.error, newRetryCount);
    logOperation('投稿失敗', 'error', `ID: ${post.id}, エラー: ${result.error}`);
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
    // メディアコンテナの作成
    const createResponse = UrlFetchApp.fetch(
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
    
    const createResult = JSON.parse(createResponse.getContentText());
    
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
    const createResponse = UrlFetchApp.fetch(
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
    
    const createResponse = UrlFetchApp.fetch(
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
    const publishResponse = UrlFetchApp.fetch(
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
  
  // 新しい列構造に合わせて調整: ID(1), 投稿内容(2), 予定日付(3), 予定時刻(4), ステータス(5), 投稿URL(6), エラー(7), リトライ(8)
  
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
  
  // エラー更新（列7）
  if (error) {
    sheet.getRange(row, 7).setValue(error);
  } else {
    sheet.getRange(row, 7).setValue('');
  }
  
  // リトライ回数更新（列8）
  if (retryCount !== undefined) {
    sheet.getRange(row, 8).setValue(retryCount);
  }
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
    const response = UrlFetchApp.fetch(
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