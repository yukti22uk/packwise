// ─── UPGRADE MODAL ───────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import { CONFIG } from '../config.js';
function UpgradeModal({open,onClose,onUnlock}){
  const[code,setCode]=useState("");
  const[codeMsg,setCodeMsg]=useState("");
  const[email,setEmail]=useState("");
  const[emailMsg,setEmailMsg]=useState("");
  if(!open) return null;

  const tryCode=()=>{
    if(CONFIG.proCodes.map(c=>c.toLowerCase()).includes(code.trim().toLowerCase())){
      onUnlock();setCodeMsg("✅ Pro unlocked! Enjoy.");
    }else setCodeMsg("❌ Invalid code. Check with us after payment.");
  };

  const submitEmail=async()=>{
    if(!email||!email.includes("@")){setEmailMsg("Enter a valid email.");return;}
    if(CONFIG.formspreeEndpoint){
      try{
        await fetch(CONFIG.formspreeEndpoint,{method:"POST",
          headers:{"Content-Type":"application/json"},
          body:JSON.stringify({email,interest:"Pro early access"})});
        setEmailMsg("✅ Thanks! We'll be in touch.");setEmail("");
      }catch(e){setEmailMsg("Could not submit — please email us directly.");}
    }else{
      window.location.href=`mailto:${CONFIG.contactEmail}?subject=Pro%20early%20access&body=Please%20add%20me:%20${encodeURIComponent(email)}`;
      setEmailMsg("Opening your email app...");
    }
  };

  const pay=()=>{
    if(CONFIG.paymentLink) window.open(CONFIG.paymentLink,"_blank");
    else window.location.href=`mailto:${CONFIG.contactEmail}?subject=Buy%20Pro&body=I%20want%20to%20upgrade%20to%20Pro.`;
  };

  return(
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.6)",
      display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:"20px"}}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:"16px",
        maxWidth:"460px",width:"100%",padding:"28px",boxShadow:"0 20px 60px rgba(0,0,0,0.3)",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"8px"}}>
          <h2 style={{margin:0,fontSize:"20px",fontWeight:"700",color:"#0f172a"}}>⭐ Upgrade to Pro</h2>
          <button onClick={onClose} style={{border:"none",background:"none",fontSize:"22px",
            cursor:"pointer",color:"#94a3b8",lineHeight:1}}>×</button>
        </div>
        <p style={{fontSize:"13px",color:"#64748b",marginTop:0}}>Unlock the features that save real freight money.</p>

        <div style={{background:"#f0fdf4",borderRadius:"10px",padding:"14px",margin:"14px 0"}}>
          <div style={{fontWeight:"700",color:"#166534",fontSize:"15px",marginBottom:"8px"}}>Pro includes:</div>
          {["🗃️ Multi-SKU Planner — pack 2–8 SKUs per container",
            "🚚 Shipment Planner — multi-container for big orders",
            "💰 Cost comparison — find the cheapest container",
            "📄 Branded PDF loading plans for your warehouse",
            "📊 Unlimited bulk SKU upload"].map(t=>(
            <div key={t} style={{fontSize:"13px",color:"#15803d",padding:"2px 0"}}>{t}</div>
          ))}
        </div>

        <div style={{textAlign:"center",fontSize:"22px",fontWeight:"800",color:"#0f172a",margin:"6px 0"}}>
          {CONFIG.priceLabel}
        </div>
        <button onClick={pay} style={{width:"100%",padding:"12px",background:"#059669",color:"#fff",
          border:"none",borderRadius:"10px",fontSize:"15px",fontWeight:"700",cursor:"pointer",marginBottom:"6px"}}>
          {CONFIG.paymentLink?"Pay & Get Pro":"Contact Us to Buy Pro"}
        </button>

        {/* Unlock with code */}
        <div style={{borderTop:"1px solid #e2e8f0",margin:"18px 0 14px",paddingTop:"16px"}}>
          <div style={{fontSize:"13px",fontWeight:"600",color:"#374151",marginBottom:"8px"}}>
            Already paid? Enter your access code:
          </div>
          <div style={{display:"flex",gap:"8px"}}>
            <input value={code} onChange={e=>setCode(e.target.value)} placeholder="Access code"
              style={{flex:1,border:"1px solid #d1d9e0",borderRadius:"8px",padding:"9px 12px",fontSize:"14px"}}/>
            <button onClick={tryCode} style={{padding:"9px 18px",background:"#0f172a",color:"#fff",
              border:"none",borderRadius:"8px",fontWeight:"600",cursor:"pointer"}}>Unlock</button>
          </div>
          {codeMsg&&<div style={{fontSize:"13px",marginTop:"8px",
            color:codeMsg.startsWith("✅")?"#166534":"#be185d"}}>{codeMsg}</div>}
        </div>

        {/* Early access email capture */}
        <div style={{borderTop:"1px solid #e2e8f0",paddingTop:"16px"}}>
          <div style={{fontSize:"13px",fontWeight:"600",color:"#374151",marginBottom:"4px"}}>
            Not ready yet? Get product updates & offers:
          </div>
          <div style={{display:"flex",gap:"8px"}}>
            <input value={email} onChange={e=>setEmail(e.target.value)} placeholder="your@email.com" type="email"
              style={{flex:1,border:"1px solid #d1d9e0",borderRadius:"8px",padding:"9px 12px",fontSize:"14px"}}/>
            <button onClick={submitEmail} style={{padding:"9px 18px",background:"#fff",color:"#059669",
              border:"1px solid #059669",borderRadius:"8px",fontWeight:"600",cursor:"pointer"}}>Notify Me</button>
          </div>
          {emailMsg&&<div style={{fontSize:"13px",marginTop:"8px",
            color:emailMsg.startsWith("✅")?"#166534":"#64748b"}}>{emailMsg}</div>}
        </div>
      </div>
    </div>
  );
}


export default UpgradeModal;
