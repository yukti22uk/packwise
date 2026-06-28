// ─── ACCOUNT PAGE ─────────────────────────────────────────────────────────────
// Shows user profile, Pro status, subscription details, cancel option.
import { useState } from 'react';
import { useAuth } from '../contexts/AuthContext.jsx';
import { CONFIG } from '../config.js';

export default function AccountPage({ setPage }) {
  const { user, profile, isPro, signOut, refreshProfile } = useAuth();
  const [cancelling, setCancelling] = useState(false);
  const [msg, setMsg] = useState('');

  if (!user) {
    setPage('auth');
    return null;
  }

  const handleSignOut = async () => {
    await signOut();
    setPage('home');
  };

  const handleCancelSubscription = async () => {
    if (!window.confirm('Are you sure you want to cancel your Pro subscription? You will lose access at the end of the current billing period.')) return;
    setCancelling(true);
    try {
      // Razorpay cancellations are handled via their dashboard or API
      // For now, direct user to email support
      setMsg('To cancel your subscription, please email us at ' + CONFIG.contactEmail + ' with your registered email and we will cancel it within 24 hours. You will retain Pro access until the end of your current billing period.');
    } catch (err) {
      setMsg('Error: ' + err.message);
    }
    setCancelling(false);
  };

  return (
    <div style={{ maxWidth:'640px', margin:'0 auto', padding:'40px 24px' }}>

      <h1 style={{ fontWeight:'800', fontSize:'28px', color:'#0f172a', margin:'0 0 8px' }}>
        My Account
      </h1>
      <p style={{ color:'#64748b', marginBottom:'32px' }}>
        Manage your profile and subscription
      </p>

      {/* Profile card */}
      <div style={{ background:'#fff', borderRadius:'12px', padding:'24px',
        boxShadow:'0 1px 4px rgba(0,0,0,0.07)', border:'1px solid #e2e8f0', marginBottom:'16px' }}>
        <div style={{ display:'flex', alignItems:'center', gap:'16px', marginBottom:'20px' }}>
          <div style={{ width:'56px', height:'56px', borderRadius:'50%',
            background:'linear-gradient(135deg,#be185d,#9d174d)',
            display:'flex', alignItems:'center', justifyContent:'center',
            fontSize:'22px', fontWeight:'800', color:'#fff', flexShrink:0 }}>
            {(profile?.full_name || user.email || 'U')[0].toUpperCase()}
          </div>
          <div>
            <div style={{ fontWeight:'700', fontSize:'16px', color:'#0f172a' }}>
              {profile?.full_name || 'DensiCube User'}
            </div>
            <div style={{ fontSize:'13px', color:'#64748b' }}>{user.email}</div>
          </div>
        </div>

        {/* Subscription status */}
        <div style={{ background: isPro?'#f0fdf4':'#f8fafc', borderRadius:'10px',
          padding:'16px', border:`1px solid ${isPro?'#bbf7d0':'#e2e8f0'}` }}>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
            <div>
              <div style={{ fontWeight:'700', fontSize:'15px',
                color: isPro?'#166534':'#374151' }}>
                {isPro ? '⭐ Pro Plan Active' : '🔒 Free Plan'}
              </div>
              <div style={{ fontSize:'13px', color:'#64748b', marginTop:'3px' }}>
                {isPro
                  ? `Active since ${profile?.pro_since ? new Date(profile.pro_since).toLocaleDateString('en-IN') : '—'} · Auto-renews monthly`
                  : 'Upgrade to unlock all 6 tools'}
              </div>
            </div>
            {!isPro && (
              <button onClick={() => setPage('pricing')}
                style={{ padding:'8px 18px', background:'#be185d', color:'#fff',
                  border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'13px',
                  cursor:'pointer', fontFamily:'inherit' }}>
                Upgrade →
              </button>
            )}
          </div>

          {isPro && (
            <div style={{ marginTop:'14px', paddingTop:'14px', borderTop:'1px solid #bbf7d0',
              display:'flex', justifyContent:'space-between', alignItems:'center',
              flexWrap:'wrap', gap:'10px' }}>
              <div style={{ fontSize:'12px', color:'#166534' }}>
                ✓ All 6 tools · ✓ Unlimited usage · ✓ PDF exports · ✓ Excel reports
              </div>
              <button onClick={handleCancelSubscription} disabled={cancelling}
                style={{ background:'none', border:'1px solid #fca5a5', color:'#dc2626',
                  borderRadius:'6px', padding:'5px 12px', fontSize:'12px', fontWeight:'600',
                  cursor:'pointer', fontFamily:'inherit' }}>
                {cancelling ? 'Processing...' : 'Cancel subscription'}
              </button>
            </div>
          )}
        </div>

        {msg && (
          <div style={{ marginTop:'12px', background:'#fffbeb', border:'1px solid #fde68a',
            borderRadius:'8px', padding:'12px', fontSize:'13px', color:'#92400e',
            lineHeight:'1.6' }}>
            {msg}
          </div>
        )}
      </div>

      {/* What's included */}
      {!isPro && (
        <div style={{ background:'#fff', borderRadius:'12px', padding:'24px',
          boxShadow:'0 1px 4px rgba(0,0,0,0.07)', border:'1px solid #e2e8f0', marginBottom:'16px' }}>
          <div style={{ fontWeight:'700', fontSize:'15px', color:'#0f172a', marginBottom:'12px' }}>
            Upgrade to Pro — ₹999/month
          </div>
          {['Multi-SKU Planner','Shipment Planner (multi-truck)','Order Analyser — ABC & FMS reports',
            'Unlimited Bulk SKU Calculator','PDF & Excel exports','Cost comparison across vehicles'].map((f,i) => (
            <div key={i} style={{ display:'flex', gap:'8px', marginBottom:'8px', fontSize:'13px' }}>
              <span style={{ color:'#059669', fontWeight:'700' }}>✓</span>
              <span style={{ color:'#374151' }}>{f}</span>
            </div>
          ))}
          <button onClick={() => setPage('pricing')}
            style={{ width:'100%', marginTop:'16px', padding:'11px',
              background:'linear-gradient(135deg,#be185d,#9d174d)',
              color:'#fff', border:'none', borderRadius:'8px', fontWeight:'700',
              fontSize:'14px', cursor:'pointer', fontFamily:'inherit' }}>
            ⭐ Upgrade Now →
          </button>
        </div>
      )}

      {/* Sign out */}
      <div style={{ background:'#fff', borderRadius:'12px', padding:'20px',
        boxShadow:'0 1px 4px rgba(0,0,0,0.07)', border:'1px solid #e2e8f0' }}>
        <button onClick={handleSignOut}
          style={{ width:'100%', padding:'10px', background:'#f8fafc',
            border:'1px solid #e2e8f0', borderRadius:'8px', fontWeight:'600',
            fontSize:'14px', color:'#374151', cursor:'pointer', fontFamily:'inherit' }}>
          Sign out
        </button>
      </div>
    </div>
  );
}
