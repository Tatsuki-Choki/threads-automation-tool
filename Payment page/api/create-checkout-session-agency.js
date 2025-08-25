const stripe = require('stripe')(process.env.STRIPE_SECRET_KEY);

module.exports = async (req, res) => {
  // 許可するオリジンのリスト
  const ALLOWED_ORIGINS = [
    'https://autothreads-payment.vercel.app',
    process.env.ALLOWED_ORIGIN_1,
    process.env.ALLOWED_ORIGIN_2
  ].filter(Boolean); // undefined を除外

  // CORSヘッダーの設定
  const origin = req.headers.origin || req.headers.referer?.replace(/\/$/, '');
  
  if (origin && ALLOWED_ORIGINS.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Vary', 'Origin');
  }
  
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,POST');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const priceId = process.env.STRIPE_PRICE_ID;
    if (!priceId) {
      return res.status(400).json({ error: 'Missing STRIPE_PRICE_ID' });
    }

    // 代理店用のCheckout Session作成
    // allow_promotion_codes: true で決済画面でのコード入力を許可
    const sessionParams = {
      payment_method_types: ['card'],
      mode: 'subscription',
      line_items: [
        {
          price: priceId,
          quantity: 1,
        },
      ],
      success_url: `${req.headers.origin}/success.html?campaign=agency&session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: `${req.headers.origin}/agency.html`,
      allow_promotion_codes: true, // 決済画面でプロモーションコード入力を許可
      subscription_data: {
        metadata: {
          campaign: 'agency',
          type: '代理店特別価格',
        },
      },
      metadata: {
        product_name: 'AutoThreads スタンダードプラン（代理店特別価格）',
        campaign: 'agency',
      },
    };

    const session = await stripe.checkout.sessions.create(sessionParams);
    res.status(200).json({ sessionId: session.id, url: session.url });
  } catch (error) {
    console.error('Stripe agency error:', error);
    res.status(500).json({ error: error.message });
  }
};