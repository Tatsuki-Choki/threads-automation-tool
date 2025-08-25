# 環境変数設定ガイド

このドキュメントでは、Threads自動化ツールのセキュリティ強化後に必要な環境変数の設定方法を説明します。

## 📋 必要な設定項目

### 1. Google Apps Script (GAS) - Script Properties

以下の値をScript Propertiesに設定する必要があります：

| キー名 | 説明 | 例 |
|--------|------|-----|
| `ADMIN_PASSWORD` | 管理者機能のパスワード | 強力なパスワードを設定 |
| `SHEET_PROTECTION_PASSWORD` | シート保護用パスワード | 強力なパスワードを設定 |
| `WEBHOOK_VERIFY_TOKEN` | Webhook検証トークン | ランダムな文字列 |
| `APP_SECRET` | Meta App Secret | Meta for Developersから取得 |
| `CLIENT_ID` | OAuth Client ID | Meta for Developersから取得 |
| `CLIENT_SECRET` | OAuth Client Secret | Meta for Developersから取得 |

### 2. Vercel - 環境変数

以下の値をVercelプロジェクトの環境変数に設定する必要があります：

| 変数名 | 説明 | 例 |
|--------|------|-----|
| `STRIPE_SECRET_KEY` | Stripe秘密鍵 | `sk_live_...` or `sk_test_...` |
| `STRIPE_PRICE_ID` | Stripe価格ID | `price_...` |
| `STRIPE_PROMOTION_CODE` | プロモーションコード（オプション） | プロモーションコードID |
| `STRIPE_COUPON` | クーポンコード（オプション） | クーポンID |
| `ALLOWED_ORIGIN_1` | 追加の許可オリジン1（オプション） | `https://example.com` |
| `ALLOWED_ORIGIN_2` | 追加の許可オリジン2（オプション） | `https://staging.example.com` |

## 🔧 設定手順

### Google Apps Script の設定

1. **スプレッドシートを開く**
   - Threads自動化ツールのスプレッドシートを開きます

2. **Apps Scriptエディタを開く**
   - メニューから「拡張機能」→「Apps Script」を選択

3. **Script Propertiesを設定**
   - エディタの左側メニューから「プロジェクトの設定」（歯車アイコン）をクリック
   - 下部の「スクリプトのプロパティ」セクションを見つける
   - 「プロパティを追加」をクリック
   - 以下の値を設定：

```javascript
// 設定例
ADMIN_PASSWORD: "your-strong-admin-password-here"
SHEET_PROTECTION_PASSWORD: "your-strong-sheet-password-here"
WEBHOOK_VERIFY_TOKEN: "your-random-webhook-token-here"
APP_SECRET: "your-meta-app-secret-here"
CLIENT_ID: "your-meta-client-id-here"
CLIENT_SECRET: "your-meta-client-secret-here"
```

4. **保存**
   - 「スクリプトのプロパティを保存」をクリック

### Vercelの設定

1. **Vercelダッシュボードにログイン**
   - https://vercel.com にアクセス
   - プロジェクトを選択

2. **環境変数を設定**
   - プロジェクトの「Settings」タブをクリック
   - 左側メニューから「Environment Variables」を選択
   - 以下の変数を追加：

```bash
# 本番環境用
STRIPE_SECRET_KEY=sk_live_xxxxxxxxxxxxxxxx
STRIPE_PRICE_ID=price_xxxxxxxxxxxxxxxx

# テスト環境用（Previewブランチ用）
STRIPE_SECRET_KEY=sk_test_xxxxxxxxxxxxxxxx
STRIPE_PRICE_ID=price_test_xxxxxxxx
```

3. **環境を選択**
   - Production: 本番環境用
   - Preview: テスト環境用
   - Development: 開発環境用

4. **保存とデプロイ**
   - 「Save」をクリック
   - 新しいデプロイをトリガーするか、次回のデプロイ時に反映されます

## 🔐 セキュリティ注意事項

1. **パスワードの強度**
   - 最低12文字以上
   - 大文字・小文字・数字・記号を含む
   - 辞書に載っている単語を避ける

2. **トークンの生成**
   - 以下のコマンドでランダムなトークンを生成できます：
   ```bash
   openssl rand -hex 32
   ```

3. **定期的な更新**
   - パスワードとトークンは定期的に更新してください
   - 最低でも3ヶ月に1回の更新を推奨

4. **アクセス制限**
   - Script Propertiesへのアクセスは編集者権限が必要
   - Vercel環境変数はプロジェクトメンバーのみアクセス可能

## 🧪 動作確認

### GAS側の確認

1. スプレッドシートのメニューから任意の管理機能を実行
2. パスワード入力を求められることを確認
3. 正しいパスワードで機能が動作することを確認

### Vercel側の確認

1. 決済ページにアクセス
2. ブラウザの開発者ツールでネットワークタブを開く
3. 決済ボタンをクリック
4. CORSエラーが発生しないことを確認

## 🆘 トラブルシューティング

### エラー: パスワードが設定されていません

**原因**: Script Propertiesが正しく設定されていない

**解決方法**:
1. Apps Scriptエディタでプロパティを確認
2. キー名が正確に一致しているか確認
3. 値が空でないことを確認

### エラー: CORSエラー

**原因**: オリジンが許可リストに含まれていない

**解決方法**:
1. Vercelの環境変数に`ALLOWED_ORIGIN_1`または`ALLOWED_ORIGIN_2`を追加
2. 正しいURLを設定（プロトコルを含む完全なURL）
3. 再デプロイを実行

### エラー: Stripe API エラー

**原因**: Stripe APIキーが正しくない

**解決方法**:
1. Stripeダッシュボードで正しいAPIキーを確認
2. 本番/テスト環境を確認
3. Vercelの環境変数を更新

## 📝 チェックリスト

設定完了後、以下の項目を確認してください：

- [ ] すべてのScript Propertiesが設定されている
- [ ] パスワードは十分に強力である
- [ ] Vercel環境変数が設定されている
- [ ] 本番とテスト環境が分離されている
- [ ] CORSが正しく機能している
- [ ] 管理機能が正常に動作する
- [ ] 決済機能が正常に動作する
- [ ] ログに機密情報が表示されない

## 🔄 移行後の注意事項

1. **古いコードの削除確認**
   - ハードコードされたパスワードやトークンが残っていないか確認
   - Git履歴からも機密情報を削除済みか確認

2. **チーム内での共有**
   - 新しいパスワードは安全な方法で共有
   - パスワードマネージャーの使用を推奨

3. **定期的な監査**
   - 月1回はアクセスログを確認
   - 不審なアクセスがないか監視

---

## 📹 動画投稿に関する重要な制限事項

### Google Drive動画の制限

Google Driveからの動画投稿には以下の制限があります：

#### 成功条件（すべて満たす必要があります）
- **ファイルサイズ**: 50MB以下
- **フォーマット**: MP4またはMOV（推奨: MP4）
- **共有設定**: 「リンクを知っている全員」に設定済み
- **公開アクセス**: 認証不要で直接動画ファイルがダウンロード可能

#### 設定手順
1. Google Driveで動画ファイルを右クリック
2. 「共有」を選択
3. 「制限付き」から「リンクを知っている全員」に変更
4. ファイルサイズが50MB以下であることを確認

#### 推奨される代替方法

Google Driveで上記条件を満たせない場合は、以下のサービスを推奨します：

- **Amazon S3**: 公開バケットでの直接URL
- **Google Cloud Storage**: 公開アクセス設定したファイル
- **Azure Blob Storage**: 公開コンテナでの直接URL
- **Cloudinary**: 動画専用ホスティングサービス
- **Uploadcare**: メディア最適化サービス

#### トラブルシューティング

**エラー: Content-Typeがvideo/*ではありません**
- Google Driveの共有設定を確認
- ブラウザでURLに直接アクセスし、HTMLページではなく動画ファイルが表示されることを確認

**エラー: ファイルサイズが上限を超えています**
- 動画を圧縮するか、他のホスティングサービスを使用
- `VIDEO_MAX_SIZE_MB`設定で上限変更可能（デフォルト50MB）

**エラー: HTTP 403/404**
- 共有設定が正しくない可能性
- ファイルが削除されているか、アクセス権限がない

質問や問題がある場合は、管理者に連絡してください。