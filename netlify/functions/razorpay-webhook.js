// ─── RAZORPAY WEBHOOK HANDLER ─────────────────────────────────────────────────
// Receives payment events from Razorpay and updates Supabase Pro status.
// Events handled:
//   subscription.activated → set is_pro = true
//   subscription.charged   → keep is_pro = true (renewal)
//   subscription.cancelled → set is_pro = false
//   subscription.completed → set is_pro = false

const crypto = require('crypto');
const { createClient } = require('@supabase/supabase-js');

// Supabase admin client (service role — backend only, never in frontend)
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// Verify Razorpay webhook signature
function verifySignature(body, signature, secret) {
  const expected = crypto
    .createHmac('sha256', secret)
    .update(body)
    .digest('hex');
  return expected === signature;
}

// Update user Pro status in Supabase profiles table
async function updateProStatus(email, isPro, subscriptionId, status) {
  const { error } = await supabase
    .from('profiles')
    .update({
      is_pro:              isPro,
      subscription_id:     subscriptionId,
      subscription_status: status,
      pro_since:           isPro ? new Date().toISOString() : null,
    })
    .eq('email', email);

  if (error) throw new Error(`Supabase update failed: ${error.message}`);
}

exports.handler = async (event) => {
  // Only accept POST
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  try {
    // ── Verify webhook signature ──────────────────────────────────────────────
    const signature = event.headers['x-razorpay-signature'];
    if (!signature) {
      return { statusCode: 400, body: 'Missing signature' };
    }

    const isValid = verifySignature(
      event.body,
      signature,
      process.env.RAZORPAY_WEBHOOK_SECRET
    );
    if (!isValid) {
      console.error('Invalid Razorpay signature');
      return { statusCode: 400, body: 'Invalid signature' };
    }

    // ── Parse event ───────────────────────────────────────────────────────────
    const payload = JSON.parse(event.body);
    const eventType = payload.event;
    const subscription = payload.payload?.subscription?.entity;

    console.log('Razorpay event:', eventType);

    if (!subscription) {
      return { statusCode: 200, body: 'No subscription entity, skipping' };
    }

    const subscriptionId = subscription.id;
    // Get email from subscription notes (we'll pass it when creating subscription)
    const email = subscription.notes?.email;

    if (!email) {
      console.error('No email in subscription notes');
      return { statusCode: 200, body: 'No email found, skipping' };
    }

    // ── Handle events ─────────────────────────────────────────────────────────
    switch (eventType) {

      case 'subscription.activated':
      case 'subscription.charged':
        // Payment successful — activate Pro
        await updateProStatus(email, true, subscriptionId, 'active');
        console.log(`✓ Pro activated for ${email}`);
        break;

      case 'subscription.cancelled':
      case 'subscription.completed':
      case 'subscription.expired':
        // Subscription ended — revoke Pro
        await updateProStatus(email, false, subscriptionId, 'cancelled');
        console.log(`✓ Pro revoked for ${email}`);
        break;

      default:
        console.log(`Unhandled event: ${eventType}`);
    }

    return { statusCode: 200, body: JSON.stringify({ received: true }) };

  } catch (err) {
    console.error('Webhook error:', err.message);
    return { statusCode: 500, body: err.message };
  }
};
