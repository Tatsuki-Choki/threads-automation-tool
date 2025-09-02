// ===========================
// 共通ユーティリティ関数
// ===========================

// ===========================
// キーワード処理の共通化
// ===========================

/**
 * アクティブなキーワードを取得
 * @return {Array} アクティブなキーワードの配列
 */
function CU_getActiveKeywords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('キーワード設定');
    
    if (!sheet) {
      logWarn('キーワード設定シートが見つかりません');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const activeKeywords = [];
    
    for (const row of data) {
      const [keyword, response, probability, isActive, conditions] = row;
      
      if (isActive === true || isActive === 'TRUE' || isActive === 'はい') {
        activeKeywords.push({
          keyword: keyword ? keyword.toString().trim() : '',
          response: response ? response.toString().trim() : '',
          probability: parseFloat(probability) || 100,
          conditions: conditions ? conditions.toString().trim() : ''
        });
      }
    }
    
    logDebug(`アクティブなキーワード数: ${activeKeywords.length}`);
    return activeKeywords;
    
  } catch (error) {
    logError('CU_getActiveKeywords', error);
    return [];
  }
}

/**
 * テキストがキーワードにマッチするか確認
 * @param {string} text - チェックするテキスト
 * @param {string} keyword - キーワード（正規表現対応）
 * @return {boolean} マッチする場合true
 */
function CU_matchesKeyword(text, keyword) {
  if (!text || !keyword) return false;
  
  try {
    // 正規表現として処理を試みる
    if (keyword.startsWith('/') && keyword.endsWith('/')) {
      const pattern = keyword.slice(1, -1);
      const regex = new RegExp(pattern, 'i');
      return regex.test(text);
    }
    
    // 通常のキーワードマッチング（大文字小文字を無視）
    return text.toLowerCase().includes(keyword.toLowerCase());
    
  } catch (error) {
    logWarn('CU_キーワードマッチングエラー', {
      keyword,
      error: error.toString()
    });
    // エラーの場合は通常の文字列として扱う
    return text.toLowerCase().includes(keyword.toLowerCase());
  }
}

/**
 * マッチするキーワードを検索
 * @param {string} text - チェックするテキスト
 * @param {Array} keywords - キーワード配列（省略時は自動取得）
 * @return {Object|null} マッチしたキーワードオブジェクトまたはnull
 */
function CU_findMatchingKeyword(text, keywords = null) {
  if (!text) return null;
  
  // キーワードが指定されていない場合は取得
  if (!keywords) {
    keywords = CU_getActiveKeywords();
  }
  
  // マッチするキーワードを検索
  const matchedKeywords = keywords.filter(kw => 
    kw.keyword && CU_matchesKeyword(text, kw.keyword)
  );
  
  if (matchedKeywords.length === 0) {
    return null;
  }
  
  // 確率に基づいて選択
  return CU_selectKeywordByProbability(matchedKeywords);
}

/**
 * 確率に基づいてキーワードを選択
 * @param {Array} keywords - キーワード配列
 * @return {Object} 選択されたキーワード
 */
function CU_selectKeywordByProbability(keywords) {
  if (keywords.length === 0) return null;
  if (keywords.length === 1) return keywords[0];
  
  // 確率の合計を計算
  const totalProbability = keywords.reduce((sum, kw) => sum + (kw.probability || 100), 0);
  
  // ランダム値を生成
  const random = Math.random() * totalProbability;
  
  // 確率に基づいて選択
  let accumulated = 0;
  for (const keyword of keywords) {
    accumulated += (keyword.probability || 100);
    if (random <= accumulated) {
      return keyword;
    }
  }
  
  // フォールバック（通常はここには到達しない）
  return keywords[0];
}

// ===========================
// 日付・時刻処理の共通化
// ===========================

/**
 * 今日の日付文字列を取得（JST）
 * @return {string} YYYY-MM-DD形式の日付
 */
function getTodayString() {
  const now = new Date();
  const jstOffset = 9 * 60; // JST = UTC+9
  const jstTime = new Date(now.getTime() + (now.getTimezoneOffset() + jstOffset) * 60000);
  
  const year = jstTime.getFullYear();
  const month = String(jstTime.getMonth() + 1).padStart(2, '0');
  const day = String(jstTime.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * 現在時刻文字列を取得（JST）
 * @return {string} HH:MM:SS形式の時刻
 */
function getCurrentTimeString() {
  const now = new Date();
  const jstOffset = 9 * 60;
  const jstTime = new Date(now.getTime() + (now.getTimezoneOffset() + jstOffset) * 60000);
  
  const hours = String(jstTime.getHours()).padStart(2, '0');
  const minutes = String(jstTime.getMinutes()).padStart(2, '0');
  const seconds = String(jstTime.getSeconds()).padStart(2, '0');
  
  return `${hours}:${minutes}:${seconds}`;
}

/**
 * タイムスタンプから経過時間を計算
 * @param {string|Date} timestamp - 比較するタイムスタンプ
 * @return {number} 経過時間（ミリ秒）
 */
function getElapsedTime(timestamp) {
  const then = timestamp instanceof Date ? timestamp : new Date(timestamp);
  return Date.now() - then.getTime();
}

// ===========================
// 設定値の取得・管理
// ===========================

/**
 * 設定値を安全に取得（Script Properties優先）
 * @param {string} key - 設定キー
 * @param {*} defaultValue - デフォルト値
 * @return {*} 設定値
 */
function getConfigSafe(key, defaultValue = null) {
  try {
    // 基本設定シートから取得を優先
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('基本設定');
    
    if (configSheet) {
      const dataRange = configSheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === key) {
          return values[i][1] || defaultValue;
        }
      }
    }
    
    // フォールバック: Script Propertiesから取得
    const scriptProps = PropertiesService.getScriptProperties();
    let value = scriptProps.getProperty(key);
    
    if (value !== null) {
      return value;
    }
    
    return defaultValue;
    
  } catch (error) {
    logError('getConfigSafe', error, { key });
    return defaultValue;
  }
}

/**
 * 設定値を保存（Script PropertiesとシートVの両方）
 * @param {string} key - 設定キー
 * @param {*} value - 設定値
 * @return {boolean} 成功時true
 */
function setConfigSafe(key, value) {
  try {
    // 基本設定シートに保存を優先
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('基本設定');
    
    if (configSheet) {
      const dataRange = configSheet.getDataRange();
      const values = dataRange.getValues();
      
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === key) {
          configSheet.getRange(i + 1, 2).setValue(value);
          found = true;
          break;
        }
      }
      
      // 新規追加
      if (!found) {
        const lastRow = configSheet.getLastRow();
        configSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
      }
    }
    
    // パスワード関連のみScript Propertiesにも保存（セキュリティ用）
    if (key.includes('PASSWORD') || key.includes('TOKEN')) {
      const scriptProps = PropertiesService.getScriptProperties();
      scriptProps.setProperty(key, String(value));
    }
    
    return true;
    
  } catch (error) {
    logError('setConfigSafe', error, { key });
    return false;
  }
}

// ===========================
// 返信済み判定の共通化
// ===========================

/**
 * 本日既に返信済みかチェック
 * @param {string} replyId - 返信ID
 * @param {string} userId - ユーザーID
 * @return {boolean} 返信済みの場合true
 */
function CU_hasAlreadyRepliedToday(replyId, userId) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `replied_${getTodayString()}_${userId}`;
    
    // キャッシュから確認
    const cachedReplies = cache.get(cacheKey);
    if (cachedReplies) {
      const replies = JSON.parse(cachedReplies);
      if (replies.includes(replyId)) {
        return true;
      }
    }
    
    // 自動返信履歴シートから確認
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName('自動返信履歴');
    
    if (!historySheet || historySheet.getLastRow() < 2) {
      return false;
    }
    
    const today = getTodayString();
    const data = historySheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateStr = row[0] instanceof Date 
        ? Utilities.formatDate(row[0], 'JST', 'yyyy-MM-dd')
        : String(row[0]);
      
      if (dateStr === today && String(row[5]) === String(userId)) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    logError('CU_hasAlreadyRepliedToday', error);
    return false;
  }
}

/**
 * 返信履歴をキャッシュに記録
 * @param {string} replyId - 返信ID
 * @param {string} userId - ユーザーID
 */
function CU_recordReplyInCache(replyId, userId) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `replied_${getTodayString()}_${userId}`;
    
    let replies = [];
    const cachedReplies = cache.get(cacheKey);
    if (cachedReplies) {
      replies = JSON.parse(cachedReplies);
    }
    
    if (!replies.includes(replyId)) {
      replies.push(replyId);
      // キャッシュは6時間保持
      cache.put(cacheKey, JSON.stringify(replies), 21600);
    }
    
  } catch (error) {
    logError('CU_recordReplyInCache', error);
  }
}

// ===========================
// クエリ文字列の構築
// ===========================

/**
 * オブジェクトからURLクエリ文字列を構築
 * @param {Object} params - パラメータオブジェクト
 * @return {string} クエリ文字列
 */
function CU_buildQueryString(params) {
  if (!params || typeof params !== 'object') {
    return '';
  }
  
  const queryParts = [];
  for (const key in params) {
    if (params.hasOwnProperty(key) && params[key] !== undefined && params[key] !== null) {
      queryParts.push(
        encodeURIComponent(key) + '=' + encodeURIComponent(params[key])
      );
    }
  }
  
  return queryParts.join('&');
}

// ===========================
// バリデーション共通化
// ===========================

/**
 * 必須設定の検証
 * @return {Object|null} 検証済み設定またはnull
 */
function CU_validateConfig() {
  const requiredConfigs = {
    accessToken: getConfigSafe('ACCESS_TOKEN'),
    userId: getConfigSafe('USER_ID'),
    username: getConfigSafe('USERNAME')
  };
  
  const missingConfigs = [];
  for (const [key, value] of Object.entries(requiredConfigs)) {
    if (!value) {
      missingConfigs.push(key);
    }
  }
  
  if (missingConfigs.length > 0) {
    logError('validateConfig', new Error('必須設定が不足'), {
      missing: missingConfigs
    });
    return null;
  }
  
  return requiredConfigs;
}

// ===========================
// テキスト処理ユーティリティ
// ===========================

/**
 * テキストを安全に切り詰め
 * @param {string} text - 切り詰めるテキスト
 * @param {number} maxLength - 最大文字数
 * @param {string} suffix - 末尾に追加する文字列
 * @return {string} 切り詰められたテキスト
 */
function truncateText(text, maxLength = 100, suffix = '...') {
  if (!text || text.length <= maxLength) {
    return text;
  }
  
  return text.substring(0, maxLength - suffix.length) + suffix;
}

/**
 * HTMLエンティティをデコード
 * @param {string} text - デコードするテキスト
 * @return {string} デコード済みテキスト
 */
function decodeHtmlEntities(text) {
  if (!text) return text;
  
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/g, "'")
    .replace(/&#x2F;/g, '/')
    .replace(/&#x5C;/g, '\\')
    .replace(/&#96;/g, '`');
}
