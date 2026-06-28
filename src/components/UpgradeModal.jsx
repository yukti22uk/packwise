// ─── UPGRADE MODAL ────────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import { CONFIG } from '../config.js';
import { useAuth } from '../contexts/AuthContext.jsx';
import RazorpayCheckout from './RazorpayCheckout.jsx';

export default function UpgradeModal({ open, onClose, onUnlock }) {
  const { user, isPro, refreshProfile } = useAuth();
  const [tab, setTab] = useState('pay'); // 'pay' | 'code'
  const [code, setCode] = useState('');
  const [codeMsg, setCodeMsg] = useState('');
  const [email, setEmail] = useState('');
  const ref = useRef(null);

  useEffect(() => {
    if (open) { setTab('pay'); setCode(''); setCodeMsg(''); }
  }, [open]);

  useEffect(() => {
    const handler = (e) => {
      if (ref.current && !ref.current.contains(e.target)) onClose();
    };
    if (open) document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open, onClose]);

  if (!open) return null;
  if (isPro) { onClose(); return null; }

  const handleCode = async () => {
    if (!CONFIG.proCodes.map(c => c.toLowerCase()).includes(code.trim().toLowerCase())) {
      setCodeMsg('❌ Invalid code. Try again or contact us.'); return;
    }
    if (onUnlock) onUnlock();
    setCodeMsg('✅ Pro unlocked! Enjoy full access.');
    setTimeout(onClose, 1500);
  };

  const handlePaymentSuccess = async () => {
    await refreshProfile();
    if (onUnlock) onUnlock();
    setTimeout(onClose, 1500);
  };

  const handleContact = async () => {
    if (CONFIG.formspreeEndpoint) {
      await fetch(CONFIG.formspreeEndpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, message: 'Pro early access request' }),
      });
    } else {
      window.location.href = `mailto:${CONFIG.contactEmail}?subject=Pro%20access&body=Email:%20${encodeURIComponent(email)}`;
    }
    setCodeMsg('✅ Received! We\'ll be in touch shortly.');
  };

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)',
      zIndex:1000, display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
      <div ref={ref} style={{ background:'#fff', borderRadius:'16px', width:'100%',
        maxWidth:'480px', boxShadow:'0 24px 64px rgba(0,0,0,0.2)', overflow:'hidden' }}>

        {/* Header */}
        <div style={{ background:'linear-gradient(135deg,#be185d,#9d174d)',
          padding:'20px 24px', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
          <div>
            <div style={{ fontWeight:'800', fontSize:'18px', color:'#fff' }}>
              ⭐ Upgrade to DensiCube Pro
            </div>
            <div style={{ fontSize:'12px', color:'rgba(255,255,255,0.75)', marginTop:'2px' }}>
              Unlock all 6 tools · {CONFIG.priceLabel}
            </div>
          </div>
          <button onClick={onClose}
            style={{ background:'rgba(255,255,255,0.15)', border:'none', color:'#fff',
              width:'28px', height:'28px', borderRadius:'50%', cursor:'pointer',
              fontSize:'14px', fontFamily:'inherit' }}>✕</button>
        </div>

        {/* Login required notice */}
        {!user && (
          <div style={{ background:'#fffbeb', border:'none', borderBottom:'1px solid #fde68a',
            padding:'12px 24px', fontSize:'13px', color:'#92400e', textAlign:'center' }}>
            ⚠ Please <strong>log in or sign up</strong> before subscribing so we can activate your Pro access.
          </div>
        )}

        {/* Tabs */}
        <div style={{ display:'flex', borderBottom:'1px solid #e2e8f0' }}>
          {[['pay','💳 Pay Online'],['code','🎟 Have a Code?']].map(([id,label]) => (
            <button key={id} onClick={() => setTab(id)}
              style={{ flex:1, padding:'12px', border:'none', fontWeight:'600',
                fontSize:'13px', cursor:'pointer', fontFamily:'inherit',
                background: tab===id?'#fff':'#f8fafc',
                color: tab===id?'#be185d':'#6b7280',
                borderBottom: tab===id?'2px solid #be185d':'2px solid transparent' }}>
              {label}
            </button>
          ))}
        </div>

        {/* Pay tab */}
        {tab === 'pay' && (
          <RazorpayCheckout
            onSuccess={handlePaymentSuccess}
            onClose={onClose}
          />
        )}

        {/* Code tab */}
        {tab === 'code' && (
          <div style={{ padding:'24px' }}>
            <div style={{ fontSize:'13px', color:'#6b7280', marginBottom:'14px' }}>
              Enter your Pro access code below.
            </div>
            <input type="text" value={code} onChange={e => setCode(e.target.value.toUpperCase())}
              placeholder="e.g. PRO-2026"
              style={{ width:'100%', padding:'11px 14px', border:'1px solid #e2e8f0',
                borderRadius:'8px', fontSize:'14px', fontFamily:'inherit', outline:'none',
                boxSizing:'border-box', letterSpacing:'0.1em', marginBottom:'10px' }}
              onKeyDown={e => e.key==='Enter' && handleCode()}/>
            <button onClick={handleCode}
              style={{ width:'100%', padding:'11px', background:'#be185d', color:'#fff',
                border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'14px',
                cursor:'pointer', fontFamily:'inherit' }}>
              Activate Pro
            </button>
            {codeMsg && (
              <div style={{ marginTop:'10px', fontSize:'13px', textAlign:'center',
                color: codeMsg.startsWith('✅')?'#166534':'#991b1b', fontWeight:'600' }}>
                {codeMsg}
              </div>
            )}
            <div style={{ marginTop:'16px', paddingTop:'16px', borderTop:'1px solid #f1f5f9' }}>
              <div style={{ fontSize:'12px', color:'#9ca3af', marginBottom:'8px', textAlign:'center' }}>
                Don't have a code? Contact us for early access.
              </div>
              <input type="email" value={email} onChange={e => setEmail(e.target.value)}
                placeholder="your@email.com"
                style={{ width:'100%', padding:'9px 12px', border:'1px solid #e2e8f0',
                  borderRadius:'8px', fontSize:'13px', fontFamily:'inherit', outline:'none',
                  boxSizing:'border-box', marginBottom:'8px' }}/>
              <button onClick={handleContact}
                style={{ width:'100%', padding:'9px', background:'#f0fdf4',
                  border:'1px solid #bbf7d0', borderRadius:'8px', fontWeight:'600',
                  fontSize:'13px', color:'#166534', cursor:'pointer', fontFamily:'inherit' }}>
                Request Access →
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
