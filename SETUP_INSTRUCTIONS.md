# 📋 Script Properties設定手順書

## 🔐 生成された認証情報

以下の認証情報が自動生成されました：

```
ADMIN_PASSWORD: PB67sjj_7NPhsliuTVMqWw
SHEET_PROTECTION_PASSWORD: aLHjML6rYivBvyCRxz956A
WEBHOOK_VERIFY_TOKEN: efc6be49e719e9c5f9d9b1094287a9565af25634173515267c90e4a2bda58a64
```

## 📝 Step 1: Meta認証情報の取得

### 1.1 Meta for Developersにアクセス
1. [https://developers.facebook.com](https://developers.facebook.com) にログイン
2. Threadsアプリを選択

### 1.2 必要な情報を取得
1. **左側メニュー** → **設定** → **基本設定**
2. 以下の情報をコピー：
   - **アプリID** (CLIENT_ID として使用)
   - **app secret** (APP_SECRET として使用) ※「表示」をクリック
   - **クライアントトークン** (CLIENT_SECRET として使用)

### 1.3 script-properties-config.txtを編集
```bash
# ファイルを開く
open script-properties-config.txt
# またはエディタで開く
code script-properties-config.txt
```

Meta認証情報の部分を実際の値に置き換え：
```
APP_SECRET=実際のapp_secretの値
CLIENT_ID=実際のアプリIDの値
CLIENT_SECRET=実際のクライアントトークンの値
```

## 📝 Step 2: Google Apps ScriptでScript Propertiesを設定

### 2.1 GASプロジェクトを開く
1. [Google Apps Script](https://script.google.com) にアクセス
2. Threads自動化ツールのプロジェクトを開く

### 2.2 Script Propertiesに移動
1. **左側メニュー**の**プロジェクトの設定**（⚙️歯車アイコン）をクリック
2. 下にスクロールして**スクリプトのプロパティ**セクションを見つける

### 2.3 プロパティを追加
1. **「プロパティを追加」**ボタンをクリック
2. 以下の順番で追加：

| プロパティ（キー） | 値 |
|------------------|-----|
| `ADMIN_PASSWORD` | PB67sjj_7NPhsliuTVMqWw |
| `SHEET_PROTECTION_PASSWORD` | aLHjML6rYivBvyCRxz956A |
| `WEBHOOK_VERIFY_TOKEN` | efc6be49e719e9c5f9d9b1094287a9565af25634173515267c90e4a2bda58a64 |
| `APP_SECRET` | *(Metaから取得した値)* |
| `CLIENT_ID` | *(Metaから取得した値)* |
| `CLIENT_SECRET` | *(Metaから取得した値)* |
| `LOG_LEVEL` | INFO |
| `MASK_SENSITIVE_DATA` | true |
| `DEBUG_MODE` | false |

### 2.4 保存
**「スクリプトのプロパティを保存」**ボタンをクリック

## 📝 Step 3: 新しいユーティリティファイルを追加

### 3.1 GASエディタでファイルを作成
1. 左側のファイル一覧で**「＋」**ボタンをクリック
2. **「スクリプト」**を選択
3. 以下のファイルを順番に作成：
   - `LoggingUtilities`
   - `ApiUtilities`
   - `CommonUtilities`
   - `TestPhase2`

### 3.2 各ファイルの内容をコピー
各ファイルに対応するコードをコピー＆ペースト：
- `src/LoggingUtilities.js` → `LoggingUtilities.gs`
- `src/ApiUtilities.js` → `ApiUtilities.gs`
- `src/CommonUtilities.js` → `CommonUtilities.gs`
- `src/TestPhase2.js` → `TestPhase2.gs`

### 3.3 ファイルの順序を調整
ドラッグ＆ドロップで以下の順序に並べ替え：
1. LoggingUtilities
2. ApiUtilities
3. CommonUtilities
4. Code
5. その他のファイル

## 📝 Step 4: 動作確認

### 4.1 設定確認テスト
GASエディタで以下の関数を実行：

```javascript
function checkScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  
  console.log('設定されているプロパティ:');
  keys.forEach(key => {
    // 機密情報はマスク表示
    const value = key.includes('PASSWORD') || key.includes('SECRET') || key.includes('TOKEN') 
      ? '***設定済み***' 
      : props.getProperty(key);
    console.log(`${key}: ${value}`);
  });
  
  // 必須項目の確認
  const required = [
    'ADMIN_PASSWORD',
    'SHEET_PROTECTION_PASSWORD', 
    'WEBHOOK_VERIFY_TOKEN',
    'APP_SECRET',
    'CLIENT_ID',
    'CLIENT_SECRET'
  ];
  
  const missing = required.filter(key => !props.getProperty(key));
  
  if (missing.length > 0) {
    console.error('❌ 未設定の必須項目:', missing);
  } else {
    console.log('✅ すべての必須項目が設定されています');
  }
}
```

### 4.2 総合テストの実行
```javascript
// TestPhase2.gsファイルで実行
testAllPhase2Features()
```

## 📝 Step 5: デプロイ

### 5.1 コードの保存
1. **Ctrl+S** (Windows) または **Cmd+S** (Mac) ですべてのファイルを保存

### 5.2 Webアプリの再デプロイ（必要な場合）
1. **デプロイ** → **デプロイを管理**
2. **編集**（鉛筆アイコン）をクリック
3. **バージョン**を**新バージョン**に変更
4. **デプロイ**をクリック

### 5.3 トリガーの確認
1. 左側メニューの**トリガー**（時計アイコン）をクリック
2. 既存のトリガーが正常に動作しているか確認
3. 必要に応じて再設定

## ✅ チェックリスト

### 設定完了確認
- [ ] Meta認証情報を取得した
- [ ] script-properties-config.txtを編集した
- [ ] Script Propertiesにすべての値を設定した
- [ ] 新しいユーティリティファイルを追加した
- [ ] checkScriptProperties()で設定を確認した
- [ ] testAllPhase2Features()が成功した
- [ ] 既存機能が正常に動作する

### セキュリティ確認
- [ ] script-properties-config.txtを安全な場所に保管または削除した
- [ ] Gitに機密情報がコミットされていない
- [ ] ログに機密情報が表示されない

## 🚨 トラブルシューティング

### エラー: ReferenceError: logInfo is not defined
**原因**: LoggingUtilitiesファイルが読み込まれていない
**解決**: ファイルの順序を確認し、LoggingUtilitiesを最上位に配置

### エラー: Script Propertyが見つからない
**原因**: プロパティ名のタイプミス
**解決**: キー名を正確にコピー＆ペースト

### エラー: 認証に失敗
**原因**: Meta認証情報が正しくない
**解決**: Meta for Developersで情報を再確認

## 📞 サポート

問題が発生した場合：
1. エラーメッセージをコピー
2. どのステップで問題が発生したか記録
3. サポートに連絡

---

設定完了予定時刻: 約30分