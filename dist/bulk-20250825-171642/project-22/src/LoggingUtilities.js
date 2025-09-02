// ===========================
// ログレベル定義
// ===========================
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3
};

// ===========================
// ログ設定の取得
// ===========================
function getLogConfig() {
  const scriptProps = PropertiesService.getScriptProperties();
  return {
    level: scriptProps.getProperty('LOG_LEVEL') || 'INFO',
    maskSensitiveData: scriptProps.getProperty('MASK_SENSITIVE_DATA') !== 'false', // デフォルトtrue
    enableDebugMode: scriptProps.getProperty('DEBUG_MODE') === 'true' // デフォルトfalse
  };
}

// ===========================
// 機密情報のマスキング
// ===========================
function maskSensitiveData(data) {
  if (typeof data !== 'string') {
    data = JSON.stringify(data);
  }
  
  // トークンのマスキング
  data = data.replace(/(["\']?access_token["\']?\s*[:=]\s*["\']?)([^"',\s}]+)/gi, '$1***MASKED***');
  data = data.replace(/(["\']?token["\']?\s*[:=]\s*["\']?)([^"',\s}]+)/gi, '$1***MASKED***');
  data = data.replace(/(["\']?secret["\']?\s*[:=]\s*["\']?)([^"',\s}]+)/gi, '$1***MASKED***');
  data = data.replace(/(["\']?password["\']?\s*[:=]\s*["\']?)([^"',\s}]+)/gi, '$1***MASKED***');
  data = data.replace(/(["\']?api_key["\']?\s*[:=]\s*["\']?)([^"',\s}]+)/gi, '$1***MASKED***');
  
  // Bearer トークンのマスキング
  data = data.replace(/(Bearer\s+)([^\s"']+)/gi, '$1***MASKED***');
  
  // OAuth認証コードのマスキング
  data = data.replace(/(["\']?code["\']?\s*[:=]\s*["\']?)([^"',\s}]{10,})/gi, '$1***MASKED***');
  
  // Stripe関連のマスキング
  data = data.replace(/(sk_[a-z]+_)([A-Za-z0-9]+)/g, '$1***MASKED***');
  data = data.replace(/(pk_[a-z]+_)([A-Za-z0-9]+)/g, '$1***MASKED***');
  
  // URLパラメータのマスキング
  data = data.replace(/(\?|&)(access_token|token|secret|password|api_key)=([^&\s"']+)/gi, '$1$2=***MASKED***');
  
  // メールアドレスの部分マスキング（最初の2文字と@以降は表示）
  data = data.replace(/([a-zA-Z0-9]{2})[a-zA-Z0-9._%+-]+(@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g, '$1***$2');
  
  // 長いID文字列の部分マスキング（最初と最後の4文字を表示）
  data = data.replace(/(["\']id["\']?\s*[:=]\s*["\']?)([a-zA-Z0-9]{8})([a-zA-Z0-9]{8,})([a-zA-Z0-9]{4})/g, '$1$2***$4');
  
  return data;
}

// ===========================
// 構造化ログ出力
// ===========================
function structuredLog(level, message, context = {}) {
  const config = getLogConfig();
  
  // ログレベルチェック
  if (LOG_LEVELS[level] < LOG_LEVELS[config.level]) {
    return;
  }
  
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    level,
    message,
    ...context
  };
  
  // 機密データのマスキング
  let logString = JSON.stringify(logEntry);
  if (config.maskSensitiveData) {
    logString = maskSensitiveData(logString);
  }
  
  // コンソール出力
  switch (level) {
    case 'ERROR':
      console.error(`[${timestamp}] [${level}] ${message}`, config.enableDebugMode ? context : '');
      break;
    case 'WARN':
      console.warn(`[${timestamp}] [${level}] ${message}`, config.enableDebugMode ? context : '');
      break;
    default:
      console.log(`[${timestamp}] [${level}] ${message}`, config.enableDebugMode ? context : '');
  }
  
  // スプレッドシートへのログ記録（ERROR以上のみ）
  if (LOG_LEVELS[level] >= LOG_LEVELS.ERROR) {
    try {
      logToSheet(level, message, logString);
    } catch (e) {
      console.error('ログシートへの記録に失敗:', e);
    }
  }
}

// ===========================
// スプレッドシートへのログ記録
// ===========================
function logToSheet(level, message, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('システムログ');
    
    if (!sheet) {
      return; // ログシートがなければスキップ
    }
    
    // マスキング済みの詳細を記録
    const maskedDetails = typeof details === 'string' ? details : maskSensitiveData(JSON.stringify(details));
    
    sheet.appendRow([
      new Date(),
      level,
      message,
      maskedDetails.substring(0, 500) // 最大500文字
    ]);
    
    // 古いログを削除（1000行を超えたら古い行から削除）
    const maxRows = 1000;
    const currentRows = sheet.getLastRow();
    if (currentRows > maxRows + 1) { // ヘッダー行を考慮
      sheet.deleteRows(2, currentRows - maxRows);
    }
  } catch (error) {
    console.error('ログシート書き込みエラー:', error);
  }
}

// ===========================
// 便利なログ関数
// ===========================
function logDebug(message, context = {}) {
  structuredLog('DEBUG', message, context);
}

function logInfo(message, context = {}) {
  structuredLog('INFO', message, context);
}

// ===========================
// 警告レベルのログ
// ===========================
function logWarning(message, context = {}) {
  structuredLog('WARNING', message, context);
}

function logWarn(message, context = {}) {
  structuredLog('WARN', message, context);
}

function logError(functionName, error, context = {}) {
  structuredLog('ERROR', `エラー in ${functionName}`, {
    functionName,
    error: error.toString(),
    stack: error.stack,
    ...context
  });
}

// ===========================
// API呼び出しログ
// ===========================
function logApiCall(url, method, requestData, responseData, duration) {
  const config = getLogConfig();
  
  // マスキング対象のデータを準備
  const logData = {
    url: maskSensitiveData(url),
    method,
    duration: `${duration}ms`,
    status: responseData?.status || 'unknown'
  };
  
  // デバッグモードの場合のみ詳細を記録
  if (config.enableDebugMode) {
    logData.request = maskSensitiveData(requestData || {});
    logData.response = maskSensitiveData(responseData || {});
  }
  
  structuredLog('INFO', `API Call: ${method} ${url}`, logData);
}

// ===========================
// 操作ログ（既存のlogOperation関数を改善）
// ===========================
function logOperationSafe(operation, status, details) {
  const config = getLogConfig();
  
  // 詳細情報のマスキング
  const maskedDetails = config.maskSensitiveData ? maskSensitiveData(details) : details;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('操作ログ');
    
    if (!sheet) {
      // 操作ログシートがない場合は作成
      sheet = ss.insertSheet('操作ログ');
      sheet.getRange(1, 1, 1, 5).setValues([['日時', '操作', 'ステータス', '詳細', 'ユーザー']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
    // 現在のユーザー情報
    const userEmail = Session.getActiveUser().getEmail() || 'システム';
    
    // ログ追加
    sheet.appendRow([
      new Date(),
      operation,
      status,
      maskedDetails,
      userEmail
    ]);
    
    // 1000行を超えたら古いログを削除
    const maxRows = 1000;
    const currentRows = sheet.getLastRow();
    if (currentRows > maxRows + 1) {
      sheet.deleteRows(2, currentRows - maxRows);
    }
    
    // 成功
    structuredLog('INFO', `操作ログ記録: ${operation}`, {
      operation,
      status,
      details: maskedDetails
    });
    
  } catch (error) {
    console.error('操作ログ記録エラー:', error);
    structuredLog('ERROR', '操作ログ記録失敗', {
      operation,
      status,
      error: error.toString()
    });
  }
}

// ===========================
// エクスポート
// ===========================
// GASでは通常のexportは使えないので、グローバル関数として定義
// 他のファイルから直接呼び出し可能