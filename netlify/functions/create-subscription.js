// ─── CREATE RAZORPAY SUBSCRIPTION ────────────────────────────────────────────
// Called from frontend when user clicks "Subscribe".
// Creates a Razorpay subscription and returns the subscription ID + key.

const Razorpay = require('razorpay');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  try {
    const { email, name } = JSON.parse(event.body);

    if (!email) {
      return { statusCode: 400, body: JSON.stringify({ error: 'Email required' }) };
    }

    const razorpay = new Razorpay({
      key_id:     process.env.RAZORPAY_KEY_ID,
      key_secret: process.env.RAZORPAY_KEY_SECRET,
    });

    // Create subscription
    const subscription = await razorpay.subscriptions.create({
      plan_id:         process.env.RAZORPAY_PLAN_ID,
      customer_notify: 1,
      total_count:     120,  // max 10 years of monthly billing
      notes: {
        email,
        name: name || email,
      },
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        subscriptionId: subscription.id,
        keyId:          process.env.RAZORPAY_KEY_ID,
      }),
    };

  } catch (err) {
    console.error('Create subscription error:', err);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
