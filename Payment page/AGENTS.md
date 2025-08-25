# Repository Guidelines

This repository hosts a lightweight payment site for AutoThreads, deployed on Vercel and integrated with Stripe Checkout. Use this guide to contribute safely and consistently.

## Project Structure & Module Organization
- Root HTML:
  - `index.html` (standard plan)
  - `campaign-free-3k.html` (campaign: free until next month + ¥3,000 OFF)
  - `success.html` (post‑checkout landing)
- Serverless API (Vercel Functions):
  - `api/create-checkout-session.js` (standard)
  - `api/create-checkout-session-campaign.js` (campaign with trial_end + discount)
- Config: `vercel.json` (functions/headers), `.env(.example)` (Stripe vars), `package.json` (scripts)

## Build, Test, and Development Commands
- `npm run dev`: Run locally via Vercel Dev (`vercel dev`). Serves pages and API.
- `npm run deploy`: Deploy to production (`vercel --prod`). Ensure env vars are set in Vercel.
- No build step is required; static HTML + serverless functions.

## Coding Style & Naming Conventions
- HTML/CSS/JS kept simple and self‑contained; prefer vanilla over frameworks.
- Match existing indentation and formatting in touched files; keep diffs minimal.
- Filenames: use kebab-case for pages (`campaign-free-3k.html`) and APIs (`create-checkout-session-*.js`).
- Do not introduce new dependencies without discussion.

## Testing Guidelines
- No automated tests configured. Perform manual verification:
  - Local: `vercel dev` → visit `/campaign-free-3k.html` and exercise checkout.
  - Stripe test: use test keys and card `4242 4242 4242 4242` to validate redirects and metadata.
  - Confirm success redirect includes `?session_id=...` and campaign query when applicable.

## Commit & Pull Request Guidelines
- Commits: concise, imperative, scoped messages (e.g., "feat: add campaign checkout API").
- PRs: include purpose, summary of changes, screenshots (UI), and steps to test. Link related issues/tickets.
- Keep changes focused; avoid unrelated refactors.

## Security & Configuration Tips
- Never commit secrets. Configure env vars in Vercel:
  - `STRIPE_SECRET_KEY`, `STRIPE_PRICE_ID`, optional `STRIPE_PROMOTION_CODE` (e.g., `SENCHAKU01`) or `STRIPE_COUPON_ID`.
- Campaign API uses `trial_end` (free until next month’s 1st) and optional discounts.
- If adding new pages/APIs, update `.env.example`, `README.md`, and `vercel.json` (function settings) accordingly.

## Communication Policy（日本語）
- Issues・TODO・PRの説明ややり取りは、原則日本語で記載してください。
- 例: 「TODO: キャンペーンページの文言をJST準拠で再確認」「PR: プロモコード適用ロジックの不具合修正」
- 外部サービス名・API名・コマンドは原文表記（英語）のままで問題ありません。
