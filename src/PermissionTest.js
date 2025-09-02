// PermissionTest.js - Threads API権限の詳細テスト

// ===========================
// 返信管理権限の詳細テスト
// ===========================
function testRepliesPermissionDetailed() {
  const ui = SpreadsheetApp.getUi();
  
  console.log('===== 返信管理権限詳細テスト =====');
  
  const config = RM_validateConfig();
  if (!config) {
    ui.alert('エラー', '認証情報が設定されていません', ui.ButtonSet.OK);
    return;
  }
  
  // 1. まず自分の投稿を取得
  console.log('\n【1. 投稿取得テスト】');
  const threadsUrl = `${THREADS_API_BASE}/v1.0/${config.userId}/threads`;
  
  try {
    const threadsResponse = fetchWithTracking(threadsUrl + '?fields=id,text,media_type,reply_audience&limit=5', {
      headers: { 'Authorization': `Bearer ${config.accessToken}` },
      muteHttpExceptions: true
    });
    
    const threadsData = JSON.parse(threadsResponse.getContentText());
    
    if (threadsData.error) {
      console.error('投稿取得エラー:', threadsData.error);
      ui.alert('エラー', '投稿の取得に失敗しました', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`✅ ${threadsData.data.length}件の投稿を取得`);
    
    if (threadsData.data.length === 0) {
      ui.alert('エラー', '投稿が見つかりません。Threadsに投稿してから再度実行してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 2. 各投稿の返信を取得してテスト
    console.log('\n【2. 返信取得テスト】');
    let successCount = 0;
    let errorCount = 0;
    const errors = [];
    
    for (let i = 0; i < Math.min(3, threadsData.data.length); i++) {
      const post = threadsData.data[i];
      console.log(`\n投稿 ${i + 1}:`);
      console.log(`- ID: ${post.id}`);
      console.log(`- テキスト: "${post.text ? post.text.substring(0, 50) + '...' : 'なし'}"`);
      console.log(`- タイプ: ${post.media_type}`);
      console.log(`- 返信対象: ${post.reply_audience || '未設定'}`);
      
      // 返信を取得
      const repliesUrl = `${THREADS_API_BASE}/v1.0/${post.id}/replies`;
      console.log(`- API URL: ${repliesUrl}`);
      
      try {
        const repliesResponse = fetchWithTracking(repliesUrl + '?fields=id,text,username,timestamp,from', {
          headers: { 'Authorization': `Bearer ${config.accessToken}` },
          muteHttpExceptions: true
        });
        
        const responseCode = repliesResponse.getResponseCode();
        const responseText = repliesResponse.getContentText();
        console.log(`- レスポンスコード: ${responseCode}`);
        
        const repliesData = JSON.parse(responseText);
        
        if (repliesData.error) {
          errorCount++;
          console.error(`❌ 返信取得失敗:`, repliesData.error);
          errors.push({
            postId: post.id,
            error: repliesData.error
          });
          
          // エラーの詳細を分析
          if (repliesData.error.code === 190) {
            console.error('→ アクセストークンの問題または権限不足');
          } else if (repliesData.error.code === 100) {
            console.error('→ APIパラメータエラー');
          } else if (repliesData.error.code === 10) {
            console.error('→ threads_manage_replies権限が必要です');
          }
        } else {
          successCount++;
          console.log(`✅ 返信取得成功: ${repliesData.data ? repliesData.data.length : 0}件`);
          
          if (repliesData.data && repliesData.data.length > 0) {
            console.log('  最新の返信:');
            const latestReply = repliesData.data[0];
            console.log(`  - ユーザー: @${latestReply.username || 'unknown'}`);
            console.log(`  - 内容: "${latestReply.text || 'なし'}"`);
          }
        }
        
      } catch (error) {
        errorCount++;
        console.error(`❌ 例外エラー:`, error.toString());
        errors.push({
          postId: post.id,
          error: { message: error.toString() }
        });
      }
    }
    
    // 3. レポート作成
    console.log('\n【3. テスト結果サマリー】');
    console.log(`成功: ${successCount}件`);
    console.log(`失敗: ${errorCount}件`);
    
    let report = 'Threads API 返信権限テスト結果\n\n';
    report += `テストした投稿数: ${Math.min(3, threadsData.data.length)}件\n`;
    report += `成功: ${successCount}件\n`;
    report += `失敗: ${errorCount}件\n\n`;
    
    if (errorCount > 0) {
      report += '❌ エラー詳細:\n';
      errors.forEach((err, index) => {
        report += `${index + 1}. 投稿ID: ${err.postId}\n`;
        report += `   エラー: ${err.error.message || 'Unknown error'}\n`;
        report += `   コード: ${err.error.code || 'N/A'}\n\n`;
      });
      
      report += '\n対処法:\n';
      report += '1. Meta for Developersでアプリの権限を確認\n';
      report += '2. threads_manage_replies権限を有効化\n';
      report += '3. アクセストークンを再生成\n';
    } else {
      report += '✅ すべての返信取得に成功しました！\n';
      report += '権限は正しく設定されています。';
    }
    
    ui.alert('テスト結果', report, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('テストエラー:', error);
    ui.alert('エラー', 'テスト中にエラーが発生しました:\n' + error.toString(), ui.ButtonSet.OK);
  }
  
  console.log('===== テスト終了 =====');
}

// ===========================
// 返信作成権限テスト（実際には投稿しない）
// ===========================
function testReplyCreationPermission() {
  const ui = SpreadsheetApp.getUi();
  
  console.log('===== 返信作成権限テスト =====');
  
  const config = RM_validateConfig();
  if (!config) {
    ui.alert('エラー', '認証情報が設定されていません', ui.ButtonSet.OK);
    return;
  }
  
  // テスト用のペイロードを作成（実際には送信しない）
  const testPayload = {
    media_type: 'TEXT',
    text: 'テスト返信（実際には送信されません）',
    reply_to_id: 'test_post_id'
  };
  
  console.log('テストペイロード:', JSON.stringify(testPayload, null, 2));
  
  // 権限の有無を既存のデータから推測
  console.log('\n権限チェック:');
  console.log('1. threads_basic: 投稿一覧の取得が可能 → ✅');
  console.log('2. threads_publish: 投稿の作成が必要 → 要確認');
  console.log('3. threads_manage_replies: 返信の取得・作成が必要 → 要確認');
  
  const report = 
    '返信作成権限テスト\n\n' +
    'テストペイロード:\n' +
    JSON.stringify(testPayload, null, 2) + '\n\n' +
    '必要な権限:\n' +
    '- threads_basic ✅\n' +
    '- threads_publish（返信作成に必要）\n' +
    '- threads_manage_replies（返信管理に必要）\n\n' +
    '注意: 実際の返信は送信されません。\n' +
    '権限の確認はMeta for Developersで行ってください。';
  
  ui.alert('テスト情報', report, ui.ButtonSet.OK);
  
  console.log('===== テスト終了 =====');
}

// ===========================
// アクセストークンの詳細情報
// ===========================
function checkAccessTokenInfo() {
  const ui = SpreadsheetApp.getUi();
  
  console.log('===== アクセストークン情報確認 =====');
  
  const config = RM_validateConfig();
  if (!config) {
    ui.alert('エラー', '認証情報が設定されていません', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // トークンデバッグ情報を取得
    const debugUrl = `https://graph.facebook.com/debug_token?input_token=${config.accessToken}&access_token=${config.accessToken}`;
    
    const response = fetchWithTracking(debugUrl, {
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.error) {
      console.error('トークン情報取得エラー:', data.error);
      ui.alert('エラー', 'トークン情報の取得に失敗しました', ui.ButtonSet.OK);
      return;
    }
    
    console.log('トークン情報:', JSON.stringify(data, null, 2));
    
    const tokenData = data.data;
    let report = 'アクセストークン情報\n\n';
    
    if (tokenData) {
      report += `アプリID: ${tokenData.app_id || 'N/A'}\n`;
      report += `ユーザーID: ${tokenData.user_id || 'N/A'}\n`;
      report += `タイプ: ${tokenData.type || 'N/A'}\n`;
      report += `有効期限: ${tokenData.expires_at ? new Date(tokenData.expires_at * 1000).toLocaleString('ja-JP') : '無期限'}\n`;
      report += `発行日時: ${tokenData.issued_at ? new Date(tokenData.issued_at * 1000).toLocaleString('ja-JP') : 'N/A'}\n\n`;
      
      if (tokenData.scopes) {
        report += 'スコープ（権限）:\n';
        tokenData.scopes.forEach(scope => {
          report += `- ${scope}\n`;
        });
      }
      
      if (!tokenData.scopes || !tokenData.scopes.includes('threads_manage_replies')) {
        report += '\n⚠️ threads_manage_replies権限がありません！\n';
        report += 'Meta for Developersで権限を追加してください。';
      }
    }
    
    ui.alert('トークン情報', report, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('トークン確認エラー:', error);
    ui.alert('エラー', 'トークン情報の確認中にエラーが発生しました', ui.ButtonSet.OK);
  }
  
  console.log('===== 確認終了 =====');
}
