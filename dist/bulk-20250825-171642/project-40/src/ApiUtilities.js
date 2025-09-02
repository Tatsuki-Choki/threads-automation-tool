// ===========================
// API呼び出しユーティリティ
// ===========================

// ===========================
// レート制限設定
// ===========================
const RATE_LIMIT_CONFIG = {
  maxRetries: 3,                    // 最大再試行回数
  initialDelay: 1000,               // 初回再試行の待機時間（ミリ秒）
  maxDelay: 32000,                  // 最大待機時間（ミリ秒）
  backoffMultiplier: 2,             // 指数バックオフの倍率
  jitterMax: 1000,                  // ジッターの最大値（ミリ秒）
  rateLimitPerMinute: 60,           // 1分あたりの最大リクエスト数
  dailyLimit: 20000                 // 1日あたりの最大リクエスト数（GAS制限）
};

// ===========================
// 改善版fetchWithTracking関数
// ===========================
function fetchWithTracking(url, params = {}) {
  const startTime = Date.now();
  const properties = PropertiesService.getScriptProperties();
  const today = new Date().toLocaleDateString('ja-JP');
  
  // API呼び出し回数の管理
  const lastCallDate = properties.getProperty('URL_FETCH_LAST_CALL_DATE');
  let dailyCount = parseInt(properties.getProperty('URL_FETCH_COUNT') || '0', 10);
  
  if (lastCallDate !== today) {
    dailyCount = 0;
  }
  
  // 日次制限チェック
  if (dailyCount >= RATE_LIMIT_CONFIG.dailyLimit) {
    const error = new Error(`API日次制限（${RATE_LIMIT_CONFIG.dailyLimit}回）に達しました`);
    logError('fetchWithTracking', error);
    throw error;
  }
  
  // レート制限の実装（1分あたりの制限）
  enforceRateLimit();
  
  // 再試行ロジックの実装
  let lastError;
  for (let attempt = 0; attempt <= RATE_LIMIT_CONFIG.maxRetries; attempt++) {
    try {
      // 再試行の場合は待機
      if (attempt > 0) {
        const delay = calculateBackoffDelay(attempt);
        logInfo(`API再試行 ${attempt}/${RATE_LIMIT_CONFIG.maxRetries}`, {
          url: maskSensitiveData(url),
          delay: `${delay}ms`
        });
        Utilities.sleep(delay);
      }
      
      // API呼び出し実行
      const response = UrlFetchApp.fetch(url, params);
      const responseCode = response.getResponseCode();
      const duration = Date.now() - startTime;
      
      // 成功時の処理
      dailyCount++;
      properties.setProperty('URL_FETCH_COUNT', dailyCount.toString());
      properties.setProperty('URL_FETCH_LAST_CALL_DATE', today);
      
      // ログ記録（マスキング済み）
      logApiCall(url, params.method || 'GET', params.payload, {
        status: responseCode,
        headers: response.getHeaders()
      }, duration);
      
      // レート制限ヘッダーの確認
      checkRateLimitHeaders(response.getHeaders());
      
      // エラーステータスの処理
      if (responseCode >= 400) {
        const responseText = response.getContentText();
        
        // 429 Too Many Requests の場合は特別処理
        if (responseCode === 429) {
          const retryAfter = response.getHeaders()['retry-after'] || response.getHeaders()['Retry-After'];
          if (retryAfter && attempt < RATE_LIMIT_CONFIG.maxRetries) {
            const waitTime = parseInt(retryAfter) * 1000;
            logWarn('Rate limit encountered', {
              url: maskSensitiveData(url),
              retryAfter: `${waitTime}ms`
            });
            Utilities.sleep(Math.min(waitTime, RATE_LIMIT_CONFIG.maxDelay));
            continue;
          }
        }
        
        // 5xx エラーは再試行対象
        if (responseCode >= 500 && attempt < RATE_LIMIT_CONFIG.maxRetries) {
          lastError = new Error(`HTTP ${responseCode}: ${responseText}`);
          continue;
        }
        
        // 4xx エラーは即座に失敗
        throw new Error(`HTTP ${responseCode}: ${maskSensitiveData(responseText)}`);
      }
      
      return response;
      
    } catch (error) {
      lastError = error;
      
      // ネットワークエラーやタイムアウトの場合は再試行
      if (attempt < RATE_LIMIT_CONFIG.maxRetries && isRetriableError(error)) {
        continue;
      }
      
      // 再試行不可能なエラーまたは最大試行回数に達した
      logError('fetchWithTracking', error, {
        url: maskSensitiveData(url),
        attempt: attempt + 1,
        maxRetries: RATE_LIMIT_CONFIG.maxRetries
      });
      throw error;
    }
  }
  
  // すべての再試行が失敗
  throw lastError || new Error('API呼び出しが失敗しました');
}

// ===========================
// レート制限の実施
// ===========================
function enforceRateLimit() {
  const properties = PropertiesService.getScriptProperties();
  const now = Date.now();
  const windowStart = now - 60000; // 1分前
  
  // 直近のAPI呼び出しタイムスタンプを取得
  const recentCalls = JSON.parse(properties.getProperty('RECENT_API_CALLS') || '[]');
  
  // 1分以内の呼び出しをフィルタ
  const callsInWindow = recentCalls.filter(timestamp => timestamp > windowStart);
  
  // レート制限チェック
  if (callsInWindow.length >= RATE_LIMIT_CONFIG.rateLimitPerMinute) {
    const oldestCall = Math.min(...callsInWindow);
    const waitTime = 60000 - (now - oldestCall) + 1000; // 1秒の余裕を追加
    
    logWarn('レート制限により待機', {
      currentRate: callsInWindow.length,
      limit: RATE_LIMIT_CONFIG.rateLimitPerMinute,
      waitTime: `${waitTime}ms`
    });
    
    Utilities.sleep(waitTime);
  }
  
  // 現在の呼び出しを記録
  callsInWindow.push(now);
  
  // 直近100件のみ保持（メモリ節約）
  const recentCallsToSave = callsInWindow.slice(-100);
  properties.setProperty('RECENT_API_CALLS', JSON.stringify(recentCallsToSave));
}

// ===========================
// バックオフ遅延の計算
// ===========================
function calculateBackoffDelay(attempt) {
  // 指数バックオフ with ジッター
  const exponentialDelay = Math.min(
    RATE_LIMIT_CONFIG.initialDelay * Math.pow(RATE_LIMIT_CONFIG.backoffMultiplier, attempt - 1),
    RATE_LIMIT_CONFIG.maxDelay
  );
  
  // ランダムジッターを追加（衝突回避）
  const jitter = Math.random() * RATE_LIMIT_CONFIG.jitterMax;
  
  return Math.floor(exponentialDelay + jitter);
}

// ===========================
// 再試行可能なエラーの判定
// ===========================
function isRetriableError(error) {
  const errorMessage = error.toString().toLowerCase();
  
  // タイムアウト、ネットワークエラー、一時的なエラー
  const retriablePatterns = [
    'timeout',
    'timed out',
    'network',
    'connection',
    'enotfound',
    'econnrefused',
    'econnreset',
    'socket hang up',
    'temporary',
    'service unavailable'
  ];
  
  return retriablePatterns.some(pattern => errorMessage.includes(pattern));
}

// ===========================
// レート制限ヘッダーの確認
// ===========================
function checkRateLimitHeaders(headers) {
  // API固有のレート制限ヘッダーを確認
  const rateLimitHeaders = {
    'x-ratelimit-limit': headers['x-ratelimit-limit'] || headers['X-RateLimit-Limit'],
    'x-ratelimit-remaining': headers['x-ratelimit-remaining'] || headers['X-RateLimit-Remaining'],
    'x-ratelimit-reset': headers['x-ratelimit-reset'] || headers['X-RateLimit-Reset']
  };
  
  if (rateLimitHeaders['x-ratelimit-remaining']) {
    const remaining = parseInt(rateLimitHeaders['x-ratelimit-remaining']);
    const limit = parseInt(rateLimitHeaders['x-ratelimit-limit'] || '0');
    
    // 残り10%未満で警告
    if (limit > 0 && remaining < limit * 0.1) {
      logWarn('API rate limit warning', {
        remaining,
        limit,
        resetTime: rateLimitHeaders['x-ratelimit-reset'] 
          ? new Date(parseInt(rateLimitHeaders['x-ratelimit-reset']) * 1000).toISOString()
          : 'unknown'
      });
    }
  }
}

// ===========================
// バッチAPI呼び出し（並列処理の代替）
// ===========================
function fetchAllWithTracking(requests) {
  const results = [];
  const errors = [];
  
  logInfo(`バッチAPI呼び出し開始: ${requests.length}件`);
  
  for (let i = 0; i < requests.length; i++) {
    const request = requests[i];
    
    try {
      // 各リクエスト間に小さな遅延を入れる（過負荷防止）
      if (i > 0) {
        Utilities.sleep(100); // 100ms間隔
      }
      
      const response = fetchWithTracking(request.url, request.params);
      results.push({
        index: i,
        success: true,
        response: response
      });
      
    } catch (error) {
      logError(`Batch request ${i}`, error);
      errors.push({
        index: i,
        url: maskSensitiveData(request.url),
        error: error.toString()
      });
      
      results.push({
        index: i,
        success: false,
        error: error
      });
      
      // エラーが多い場合は中断
      if (errors.length > requests.length * 0.5) {
        logError('fetchAllWithTracking', new Error('バッチ処理で過半数のリクエストが失敗したため中断'));
        break;
      }
    }
  }
  
  logInfo(`バッチAPI呼び出し完了`, {
    total: requests.length,
    success: results.filter(r => r.success).length,
    failed: errors.length
  });
  
  return results;
}

// ===========================
// API使用状況の取得
// ===========================
function getApiUsageStats() {
  const properties = PropertiesService.getScriptProperties();
  const today = new Date().toLocaleDateString('ja-JP');
  
  const lastCallDate = properties.getProperty('URL_FETCH_LAST_CALL_DATE');
  const dailyCount = lastCallDate === today 
    ? parseInt(properties.getProperty('URL_FETCH_COUNT') || '0', 10)
    : 0;
  
  const recentCalls = JSON.parse(properties.getProperty('RECENT_API_CALLS') || '[]');
  const now = Date.now();
  const callsLastMinute = recentCalls.filter(t => t > now - 60000).length;
  
  return {
    daily: {
      used: dailyCount,
      limit: RATE_LIMIT_CONFIG.dailyLimit,
      remaining: RATE_LIMIT_CONFIG.dailyLimit - dailyCount,
      percentage: Math.round((dailyCount / RATE_LIMIT_CONFIG.dailyLimit) * 100)
    },
    perMinute: {
      used: callsLastMinute,
      limit: RATE_LIMIT_CONFIG.rateLimitPerMinute,
      remaining: Math.max(0, RATE_LIMIT_CONFIG.rateLimitPerMinute - callsLastMinute)
    },
    lastCallDate: lastCallDate || 'なし'
  };
}

// ===========================
// API使用状況のリセット（手動実行用）
// ===========================
function resetApiUsageStats() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('URL_FETCH_COUNT');
  properties.deleteProperty('URL_FETCH_LAST_CALL_DATE');
  properties.deleteProperty('RECENT_API_CALLS');
  
  logInfo('API使用状況をリセットしました');
  return 'API使用状況をリセットしました';
}