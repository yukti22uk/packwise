// ─── RAZORPAY SUBSCRIPTION CHECKOUT ──────────────────────────────────────────
// Opens Razorpay payment modal for Pro subscription.
// On success → refreshes profile to reflect Pro status.
import { useState } from 'react';
import { useAuth } from '../contexts/AuthContext.jsx';
import { CONFIG } from '../config.js';

export default function RazorpayCheckout({ onSuccess, onClose }) {
  const { user, refreshProfile } = useAuth();
  const [loading, setLoading] = useState(false);
  const [error,   setError]   = useState('');

  const handlePayment = async () => {
    if (!user) { setError('Please log in first.'); return; }
    setLoading(true);
    setError('');

    try {
      // 1. Create subscription via Netlify Function
      const res = await fetch('/.netlify/functions/create-subscription', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          email: user.email,
          name:  user.user_metadata?.full_name || user.email,
        }),
      });

      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || 'Failed to create subscription');

      // 2. Load Razorpay script if not already loaded
      if (!window.Razorpay) {
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://checkout.razorpay.com/v1/checkout.js';
          script.onload = resolve;
          script.onerror = reject;
          document.body.appendChild(script);
        });
      }

      // 3. Open Razorpay checkout
      const rzp = new window.Razorpay({
        key:             data.keyId,
        subscription_id: data.subscriptionId,
        name:            'DensiCube',
        description:     'Pro Plan — ₹999/month',
        image:           '/favicon.svg',
        prefill: {
          name:  user.user_metadata?.full_name || '',
          email: user.email,
        },
        theme:   { color: '#be185d' },
        handler: async (response) => {
          // Payment successful — wait 2 seconds for webhook to process, then refresh
          setTimeout(async () => {
            await refreshProfile();
            if (onSuccess) onSuccess();
          }, 2000);
        },
        modal: {
          ondismiss: () => { setLoading(false); if (onClose) onClose(); },
        },
      });

      rzp.open();

    } catch (err) {
      setError(err.message || 'Payment failed. Please try again.');
      setLoading(false);
    }
  };

  return (
    <div style={{ padding:'24px', textAlign:'center' }}>
      {/* Plan summary */}
      <div style={{ background:'#f0fdf4', borderRadius:'12px', padding:'20px',
        border:'1px solid #bbf7d0', marginBottom:'20px' }}>
        <div style={{ fontSize:'28px', fontWeight:'900', color:'#0f172a', marginBottom:'4px' }}>
          ₹999
          <span style={{ fontSize:'15px', fontWeight:'500', color:'#64748b' }}>/month</span>
        </div>
        <div style={{ fontSize:'13px', color:'#166534', fontWeight:'600', marginBottom:'12px' }}>
          All Pro tools · Auto-renews monthly · Cancel anytime
        </div>
        {['Multi-SKU Planner','Shipment Planner','Order Analyser',
          'Unlimited Bulk SKU','PDF & Excel exports'].map((f,i) => (
          <div key={i} style={{ fontSize:'12px', color:'#374151',
            display:'flex', alignItems:'center', gap:'6px',
            justifyContent:'center', marginBottom:'4px' }}>
            <span style={{ color:'#059669', fontWeight:'700' }}>✓</span> {f}
          </div>
        ))}
      </div>

      {error && (
        <div style={{ background:'#fff1f2', border:'1px solid #fecaca', borderRadius:'8px',
          padding:'10px 14px', fontSize:'13px', color:'#991b1b', marginBottom:'14px' }}>
          ⚠ {error}
        </div>
      )}

      <button onClick={handlePayment} disabled={loading}
        style={{ width:'100%', padding:'13px',
          background: loading?'#e2e8f0':'linear-gradient(135deg,#be185d,#9d174d)',
          color: loading?'#9ca3af':'#fff', border:'none', borderRadius:'10px',
          fontWeight:'700', fontSize:'15px', cursor: loading?'not-allowed':'pointer',
          fontFamily:'inherit', boxShadow: loading?'none':'0 4px 14px rgba(190,24,93,0.35)' }}>
        {loading ? '⏳ Opening payment...' : '⭐ Subscribe — ₹999/month'}
      </button>

      <div style={{ fontSize:'11px', color:'#94a3b8', marginTop:'10px' }}>
        Secured by Razorpay · UPI, Cards, Net Banking accepted
      </div>
    </div>
  );
}
