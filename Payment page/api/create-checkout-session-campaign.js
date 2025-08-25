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

    // Prepare discounts: prefer promotion_code (if provided) then coupon
    let discounts;
    const promoEnv = process.env.STRIPE_PROMOTION_CODE;
    const couponEnv = process.env.STRIPE_COUPON_ID;

    if (promoEnv) {
      // Accept either promotion_code ID (starts with 'promo_') or code string like 'SENCHAKU01'
      let promoId = null;
      if (promoEnv.startsWith('promo_')) {
        promoId = promoEnv;
      } else {
        // Look up promotion code by code string
        const list = await stripe.promotionCodes.list({ code: promoEnv, active: true, limit: 1 });
        if (list && list.data && list.data.length > 0) {
          promoId = list.data[0].id;
        } else {
          return res.status(400).json({ error: `Promotion code not found or inactive: ${promoEnv}` });
        }
      }
      discounts = [{ promotion_code: promoId }];
    } else if (couponEnv) {
      discounts = [{ coupon: couponEnv }];
    }

    // Compute trial_end and billing anchor: free until the 1st of next month
    const nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    nextMonth.setDate(1);
    nextMonth.setHours(0, 0, 0, 0);
    const trialEnd = Math.floor(nextMonth.getTime() / 1000);

    const sessionParams = {
      payment_method_types: ['card'],
      mode: 'subscription',
      line_items: [
        {
          price: priceId,
          quantity: 1,
        },
      ],
      success_url: `${req.headers.origin}/success.html?campaign=free-3k&session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: `${req.headers.origin}/campaign-free-3k.html`,
      subscription_data: {
        trial_end: trialEnd,
        metadata: {
          campaign: 'free-3k',
          trial_note: '翌月1日まで無料',
        },
      },
      metadata: {
        product_name: 'AutoThreads スタンダードプラン（キャンペーン）',
        campaign: 'free-3k',
      },
    };

    if (discounts) {
      sessionParams.discounts = discounts;
    } else {
      sessionParams.allow_promotion_codes = true;
    }

    const session = await stripe.checkout.sessions.create(sessionParams);
    res.status(200).json({ sessionId: session.id, url: session.url });
  } catch (error) {
    console.error('Stripe campaign error:', error);
    res.status(500).json({ error: error.message });
  }
};
