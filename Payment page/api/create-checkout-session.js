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
    // 翌月1日のタイムスタンプを計算
    const nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    nextMonth.setDate(1);
    nextMonth.setHours(0, 0, 0, 0);
    const billingAnchor = Math.floor(nextMonth.getTime() / 1000);

    // 現在の日付から翌月1日までの日数を計算（日割り計算用）
    const now = new Date();
    const daysUntilNextMonth = Math.ceil((nextMonth - now) / (1000 * 60 * 60 * 24));
    const totalDaysInMonth = 30; // 簡略化のため30日として計算
    

    // Checkout Sessionの作成
    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      line_items: [
        {
          price: process.env.STRIPE_PRICE_ID || 'price_1Rp1PFGfymrOs6EypMqBguhm',
          quantity: 1,
        },
      ],
      mode: 'subscription',
      success_url: `${req.headers.origin}/success.html?session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: `${req.headers.origin}/`,
      subscription_data: {
        billing_cycle_anchor: billingAnchor,
        proration_behavior: 'create_prorations',
        metadata: {
          initial_payment_note: `初月は日割り計算（${daysUntilNextMonth}日分）`,
        },
      },

      metadata: {
        product_name: 'AutoThreads スタンダードプラン',
      },
    });

    res.status(200).json({ sessionId: session.id, url: session.url });
  } catch (error) {
    console.error('Stripe error:', error);
    res.status(500).json({ error: error.message });
  }
};