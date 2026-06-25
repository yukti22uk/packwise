// ─── PRO GATE ────────────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
function ProGate({onUpgrade,feature}){
  return(
    <div style={{...S.card,textAlign:"center",padding:"48px 32px",
      background:"linear-gradient(135deg,#f0fdf4 0%,#eff6ff 100%)",border:"1px solid #bbf7d0"}}>
      <div style={{fontSize:"48px",marginBottom:"12px"}}>🔒</div>
      <div style={{fontSize:"20px",fontWeight:"700",color:"#0f172a",marginBottom:"8px"}}>
        {feature} is a Pro feature
      </div>
      <div style={{fontSize:"14px",color:"#64748b",maxWidth:"440px",margin:"0 auto 20px",lineHeight:"1.6"}}>
        Upgrade to plan full shipments across multiple containers, compare costs to find the
        cheapest option, and export branded PDF loading plans for your warehouse team.
      </div>
      <button onClick={onUpgrade} style={{padding:"12px 28px",background:"#059669",color:"#fff",
        border:"none",borderRadius:"10px",fontSize:"15px",fontWeight:"700",cursor:"pointer"}}>
        ⭐ Upgrade to Pro
      </button>
    </div>
  );
}


export default ProGate;
