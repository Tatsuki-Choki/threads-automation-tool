// QuickSetup.gs - 既存トークンでのクイックセットアップ

// ===========================
// クイックセットアップ関数
// ===========================
function quickSetupWithExistingToken() {
  const ui = SpreadsheetApp.getUi();
  
  // 必要な情報の確認
  const accessToken = getConfig('ACCESS_TOKEN');
  const clientSecret = getConfig('CLIENT_SECRET');
  
  if (!accessToken || !clientSecret) {
    ui.alert('エラー', '基本設定シートに必要な情報を入力してください：\n- ACCESS_TOKEN\n- CLIENT_SECRET', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // 1. ユーザー情報の取得とテスト
    ui.alert('セットアップ開始', 'Threads APIの接続をテストし、ユーザー情報を取得します...', ui.ButtonSet.OK);
    
    // アクセストークンからユーザー情報を取得
    const userInfo = fetchUserInfoFromToken(accessToken);
    
    if (userInfo) {
      // ユーザーIDとユーザー名を保存
      setConfig('USER_ID', userInfo.id);
      setConfig('USERNAME', userInfo.username);
      ui.alert('接続成功！', `ユーザーID: ${userInfo.id}\nユーザー名: @${userInfo.username}`, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', 'ユーザー情報の取得に失敗しました。ACCESS_TOKENが正しいか確認してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 2. CLIENT_IDの推測（オプション）
    const clientId = getConfig('CLIENT_ID');
    if (!clientId || clientId === '（後で入力）') {
      const response = ui.prompt(
        'CLIENT_IDの入力',
        'Meta開発者ダッシュボードからアプリID（CLIENT_ID）を入力してください：\n' +
        '（後で基本設定シートに直接入力することもできます）',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() === ui.Button.OK && response.getResponseText()) {
        setConfig('CLIENT_ID', response.getResponseText());
      }
    }
    
    // 3. トークンの有効期限設定（推定）
    setConfig('TOKEN_EXPIRES', new Date(Date.now() + 60 * 24 * 60 * 60 * 1000).toISOString());
    
    // 4. 基本的なログ記録
    logOperation('クイックセットアップ', 'success', `ユーザー: @${userInfo.username}`);
    
    // 5. テスト投稿の準備
    const testResponse = ui.alert(
      'セットアップ完了',
      'セットアップが完了しました！\n\nテスト投稿を実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (testResponse === ui.Button.YES) {
      executeTestPost();
    }
    
  } catch (error) {
    ui.alert('エラー', `セットアップ中にエラーが発生しました：\n${error.toString()}`, ui.ButtonSet.OK);
    logError('quickSetupWithExistingToken', error);
  }
}

// ===========================
// アクセストークンからユーザー情報を取得
// ===========================
function fetchUserInfoFromToken(accessToken) {
  try {
    const response = UrlFetchApp.fetch(
      `${THREADS_API_BASE}/v1.0/me?fields=id,username,threads_profile_picture_url`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      return JSON.parse(responseText);
    } else {
      console.error(`APIエラー: ${responseCode} - ${responseText}`);
      return null;
    }
    
  } catch (error) {
    console.error(`ユーザー情報取得エラー: ${error.toString()}`);
    return null;
  }
}

// ===========================
// 接続テストとユーザー情報取得
// ===========================
function testConnectionAndGetUserInfo(accessToken, userId) {
  try {
    const response = UrlFetchApp.fetch(
      `${THREADS_API_BASE}/v1.0/${userId}?fields=id,username,threads_profile_picture_url,threads_biography`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      return JSON.parse(responseText);
    } else {
      const error = JSON.parse(responseText);
      throw new Error(error.error?.message || `APIエラー: ${responseCode}`);
    }
    
  } catch (error) {
    throw new Error(`接続テスト失敗: ${error.toString()}`);
  }
}

// ===========================
// テスト投稿の実行
// ===========================
function executeTestPost() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // テスト投稿の内容
    const testContent = `Threads自動化ツールのテスト投稿です 🤖\n\n投稿時刻: ${new Date().toLocaleString('ja-JP')}\n\n#ThreadsAPI #自動化テスト`;
    
    // 投稿実行
    const result = postTextOnly(testContent);
    
    if (result.success) {
      ui.alert(
        'テスト投稿成功！',
        `投稿が成功しました！\n\nURL: ${result.postUrl}\n\nScheduledPostsシートで予約投稿の設定を開始できます。`,
        ui.ButtonSet.OK
      );
      
      // ログに記録
      logOperation('テスト投稿', 'success', result.postUrl);
    } else {
      ui.alert('エラー', `投稿失敗: ${result.error}`, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('エラー', `テスト投稿中にエラーが発生しました：\n${error.toString()}`, ui.ButtonSet.OK);
    logError('executeTestPost', error);
  }
}


// ===========================
// デバッグ用関数
// ===========================
function debugCheck基本設定() {
  console.log('=== 基本設定 シートの内容確認 ===');
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('基本設定');
    if (!sheet) {
      console.error('基本設定シートが見つかりません');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    console.log(`基本設定シートの行数: ${data.length}`);
    
    // 重要な設定値を確認
    const importantKeys = ['ACCESS_TOKEN', 'USER_ID', 'CLIENT_SECRET', 'USERNAME'];
    
    for (const key of importantKeys) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          const value = data[i][1];
          console.log(`${key}: ${value ? '設定済み' : '未設定'} (長さ: ${value ? value.toString().length : 0})`);
          found = true;
          break;
        }
      }
      if (!found) {
        console.log(`${key}: 行が存在しません`);
      }
    }
    
    // getConfig関数のテスト
    console.log('\n=== getConfig関数のテスト ===');
    const accessToken = getConfig('ACCESS_TOKEN');
    const userId = getConfig('USER_ID');
    console.log(`getConfig('ACCESS_TOKEN'): ${accessToken ? '取得成功' : '取得失敗'}`);
    console.log(`getConfig('USER_ID'): ${userId ? '取得成功' : '取得失敗'}`);
    
  } catch (error) {
    console.error('debugCheck基本設定 エラー:', error);
  }
}

function debugGetUserPosts() {
  const accessToken = getConfig('ACCESS_TOKEN');
  const userId = getConfig('USER_ID');
  
  if (!accessToken || !userId) {
    console.log('認証情報が不足しています');
    return;
  }
  
  try {
    const response = UrlFetchApp.fetch(
      `${THREADS_API_BASE}/v1.0/${userId}/threads?fields=id,text,timestamp,media_type,media_url&limit=5`,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        },
        muteHttpExceptions: true
      }
    );
    
    const result = JSON.parse(response.getContentText());
    
    if (result.data) {
      console.log(`最新${result.data.length}件の投稿:`);
      result.data.forEach((post, index) => {
        console.log(`\n投稿${index + 1}:`);
        console.log(`ID: ${post.id}`);
        console.log(`投稿時刻: ${new Date(post.timestamp)}`);
        console.log(`内容: ${post.text || '(テキストなし)'}`);
        console.log(`メディアタイプ: ${post.media_type}`);
      });
    } else {
      console.log('投稿が取得できませんでした:', result);
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}

// ===========================
// サンプルデータ生成
// ===========================
function generateSampleData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'サンプルデータ生成',
    'ScheduledPostsとAutoReplyKeywordsにサンプルデータを追加しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // 予約投稿サンプル
    const postsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ScheduledPosts');
    const now = new Date();
    
    const samplePosts = [
      {
        content: 'Threads自動化ツールから投稿しています！ 🚀 #自動投稿',
        scheduledTime: new Date(now.getTime() + 10 * 60 * 1000),
        images: []
      },
      {
        content: '定期投稿のテストです。\n\n今日も良い一日を！ ☀️',
        scheduledTime: new Date(now.getTime() + 30 * 60 * 1000),
        images: []
      },
      {
        content: '画像付き投稿のテスト 📸',
        scheduledTime: new Date(now.getTime() + 60 * 60 * 1000),
        images: ['https://example.com/image1.jpg']
      },
      {
        content: '複数画像投稿のテスト 📸📸📸',
        scheduledTime: new Date(now.getTime() + 90 * 60 * 1000),
        images: ['https://example.com/image1.jpg', 'https://example.com/image2.jpg', 'https://example.com/image3.jpg']
      }
    ];
    
    samplePosts.forEach((post, index) => {
      const rowData = [
        postsSheet.getLastRow(), // ID
        post.content, // 投稿内容
        post.scheduledTime, // 予定日時
        'pending', // ステータス
        '', // 投稿URL
        '', // エラー
        0 // リトライ
      ];
      
      // 画像URLを追加（最大10枚）
      for (let i = 0; i < 10; i++) {
        rowData.push(post.images[i] || '');
      }
      
      postsSheet.appendRow(rowData);
    });
    
    // 自動返信キーワードサンプル
    const keywordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoReplyKeywords');
    
    const sampleKeywords = [
      ['ありがとう', 'partial', 'こちらこそありがとうございます！ 😊', true, 1],
      ['詳細', 'partial', '詳細については、プロフィールのリンクをご確認ください。', true, 2],
      ['質問', 'partial', 'ご質問ありがとうございます。どのような内容でしょうか？', true, 3],
      ['^フォロー.*しました$', 'regex', 'フォローありがとうございます！これからよろしくお願いします！ 🎉', true, 4],
      ['購入', 'partial', 'ご購入をご検討いただきありがとうございます。DMにて詳細をお送りします。', false, 5]
    ];
    
    sampleKeywords.forEach((keyword, index) => {
      keywordsSheet.appendRow([
        keywordsSheet.getLastRow(), // ID
        keyword[0], // キーワード
        keyword[1], // マッチタイプ
        keyword[2], // 返信内容
        keyword[3], // 有効
        keyword[4] // 優先度
      ]);
    });
    
    ui.alert('完了', 'サンプルデータを追加しました！', ui.ButtonSet.OK);
    logOperation('サンプルデータ生成', 'success', `投稿${samplePosts.length}件、キーワード${sampleKeywords.length}件`);
    
  } catch (error) {
    ui.alert('エラー', `サンプルデータの生成中にエラーが発生しました：\n${error.toString()}`, ui.ButtonSet.OK);
    logError('generateSampleData', error);
  }
}