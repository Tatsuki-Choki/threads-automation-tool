# AutoThreads Payment Page

Vercelでホスティングする決済ページです。

## セットアップ手順

### 1. 依存関係のインストール
```bash
npm install
```

### 2. 環境変数の設定
`.env.example`をコピーして`.env`を作成し、Stripeの情報を入力：
```bash
cp .env.example .env
```

### 3. Vercel CLIのインストール
```bash
npm install -g vercel
```

### 4. ローカル開発
```bash
npm run dev
# または
vercel dev
```

### 5. デプロイ

#### 初回デプロイ
```bash
vercel
```

以下の質問に答えてください：
- Set up and deploy? → Y
- Which scope? → あなたのアカウントを選択
- Link to existing project? → N（新規の場合）
- Project name? → autothreads-payment（任意）
- Directory? → ./ （そのままEnter）
- Override settings? → N

#### 環境変数の設定（Vercelダッシュボード）
1. [Vercel Dashboard](https://vercel.com/dashboard)にアクセス
2. プロジェクトを選択
3. Settings → Environment Variables
4. 以下を追加：
   - `STRIPE_SECRET_KEY`: 本番用のStripeシークレットキー
   - `STRIPE_PRICE_ID`: price_1Rp1PFGfymrOs6EypMqBguhm
   - （任意）`STRIPE_COUPON_ID`: 2ヶ月目3,000円OFFクーポンID（duration: once, amount_off: 3000, currency: jpy）
   - （任意）`STRIPE_PROMOTION_CODE`: 上記クーポンに紐づくプロモコード（例: `SENCHAKU01` もしくは `promo_...` ID）
   - （任意）`TRIAL_PERIOD_DAYS`: 無料トライアル日数（既定: 30）

#### 本番デプロイ
```bash
npm run deploy
# または
vercel --prod
```

## 機能説明

### 翌月1日開始のサブスクリプション
- `billing_cycle_anchor`で翌月1日を指定
- 初月は自動的に日割り計算
- クーポンコード「SENKO01」で初月5,000円OFF

### キャンペーン: 初月無料 ＋ 2ヶ月目3,000円OFF
- 新ページ: `campaign-free-3k.html`
- 新API: `/api/create-checkout-session-campaign`
- 仕様:
  - `trial_end`で「翌月1日まで無料」を実現（初回請求は翌月1日）
  - 初回請求（翌月1日）に3,000円OFFの割引を自動適用（`STRIPE_PROMOTION_CODE` または `STRIPE_COUPON_ID`）
  - `STRIPE_PROMOTION_CODE` は"コード文字列"または`promo_...` IDのどちらでも指定可能（コード指定時はAPI側でIDを自動解決）
  - クーポン未設定時はプロモコード入力を許可
  - 成功時: `success.html?campaign=free-3k&session_id=...`
  - 取消時: `campaign-free-3k.html`へ戻る

### ファイル構成
```
Payment page/
├── index.html                     # メインページ
├── campaign-free-3k.html          # キャンペーンページ（初月無料＋2ヶ月目3,000円OFF）
├── success.html                   # 決済完了ページ
├── api/
│   ├── create-checkout-session.js           # 通常用 Stripe API
│   └── create-checkout-session-campaign.js  # キャンペーン用 Stripe API
├── package.json
├── vercel.json                    # Vercel設定
└── .env.example                   # 環境変数テンプレート
```

## 注意事項
- 本番環境では必ず本番用のStripeキーを使用
- 環境変数は絶対にGitにコミットしない
- クーポンコードはStripeダッシュボードで事前に作成が必要
  - キャンペーン用は「duration: once」「amount_off: 3000」「currency: jpy」を推奨
