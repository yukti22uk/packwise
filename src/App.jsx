import { useState, useRef, useEffect } from "react";
import * as THREE from "three";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

// ══════════════════════════════════════════════════════════════════════════════
//  ⚙️  OWNER CONFIG — EDIT THESE VALUES, then redeploy
// ══════════════════════════════════════════════════════════════════════════════
const CONFIG = {
  // Your contact email (used for "Contact for Pro" and email-capture fallback)
  contactEmail: "you@example.com",

  // Your Razorpay / payment page link (create a free payment link in Razorpay dashboard).
  // Leave "" to fall back to an email contact button.
  paymentLink: "",

  // OPTIONAL: A Formspree form endpoint to collect early-access emails.
  // Sign up free at formspree.io, create a form, paste its URL here e.g.
  // "https://formspree.io/f/abcdwxyz". Leave "" to use an email (mailto) fallback.
  formspreeEndpoint: "",

  // Access codes you give to paying customers. They enter one to unlock Pro.
  // Change these to your own secret codes.
  proCodes: ["PRO-2026", "EARLYBIRD"],

  // Free-tier limit on bulk SKU rows
  freeSkuLimit: 10,

  // Pricing shown in the upgrade modal (display only)
  priceLabel: "₹999 / month",
};
// ══════════════════════════════════════════════════════════════════════════════

// ─── PRESETS ──────────────────────────────────────────────────────────────────
const PALLET_BASES = [
  { label:"1200×1000 mm", L:1200, W:1000 },
  { label:"1200×1200 mm", L:1200, W:1200 },
  { label:"1200×800 mm",  L:1200, W:800  },
];
const VEHICLES = [
  { label:"Tata Ace",             L:2100,  W:1525, H:1600 },
  { label:"19ft",                 L:5800,  W:2350, H:2100 },
  { label:"20ft Container (ISO)", L:5900,  W:2350, H:2390 },
  { label:"22ft",                 L:6700,  W:2350, H:2100 },
  { label:"32ft SXL",             L:9750,  W:2350, H:2700 },
  { label:"32ft MXL",             L:9750,  W:2430, H:2900 },
  { label:"40ft Container (ISO)", L:12032, W:2352, H:2395 },
];
const VEHICLES_WITH_CUSTOM=[...VEHICLES,{label:"Custom (Manual Input)",L:null,W:null,H:null}];

// ─── PACKING LOGIC ────────────────────────────────────────────────────────────
function getPerms(sl,sw,sh,lockHeight){
  if(lockHeight) return [[sl,sw,sh],[sw,sl,sh]];
  return [[sl,sw,sh],[sl,sh,sw],[sw,sl,sh],[sw,sh,sl],[sh,sl,sw],[sh,sw,sl]];
}
function bestFitDetailed(L,W,H,sl,sw,sh,opt={}){
  if(L<=0||W<=0||H<=0) return{count:0,nx:0,ny:0,nz:0,boxL:sl,boxW:sw,boxH:sh};
  const perms=getPerms(sl,sw,sh,opt.lockHeight);
  let best=0,bx=[sl,sw,sh],bn=[0,0,0];
  for(const[a,b,c]of perms){
    let nx=Math.floor(L/a),ny=Math.floor(W/b),nz=Math.floor(H/c);
    if(opt.noStack) nz=Math.min(nz,1);
    if(opt.maxStack&&opt.maxStack>0) nz=Math.min(nz,opt.maxStack);
    const cnt=nx*ny*nz;
    if(cnt>best){best=cnt;bx=[a,b,c];bn=[nx,ny,nz];}
  }
  return{count:best,nx:bn[0],ny:bn[1],nz:bn[2],boxL:bx[0],boxW:bx[1],boxH:bx[2]};
}
function calcMixedDetailed(cL,cW,cH,sl,sw,sh,opt={}){
  const perms=getPerms(sl,sw,sh,opt.lockHeight);
  let best=-1,R={};
  for(const[a,b,c]of perms){
    let nx=Math.floor(cL/a),ny=Math.floor(cW/b),nz=Math.floor(cH/c);
    if(opt.noStack) nz=Math.min(nz,1);
    if(opt.maxStack&&opt.maxStack>0) nz=Math.min(nz,opt.maxStack);
    const uL=nx*a,uW=ny*b,uH=nz*c;
    const l1=bestFitDetailed(cL-uL,cW,cH,sl,sw,sh,opt);
    const l2=bestFitDetailed(uL,cW-uW,cH,sl,sw,sh,opt);
    const l3=(opt.noStack||(opt.maxStack&&opt.maxStack>0))?{count:0,nx:0,ny:0,nz:0,boxL:sl,boxW:sw,boxH:sh}
              :bestFitDetailed(uL,uW,cH-uH,sl,sw,sh,opt);
    const tot=nx*ny*nz+l1.count+l2.count+l3.count;
    if(tot>best){best=tot;
      R={total:tot,nx,ny,nz,boxL:a,boxW:b,boxH:c,usedL:uL,usedW:uW,usedH:uH,cL,cW,cH,sl,sw,sh,
        leftover1:{...l1,offX:uL,offY:0,offZ:0},
        leftover2:{...l2,offX:0,offY:uW,offZ:0},
        leftover3:{...l3,offX:0,offY:0,offZ:uH},
        orient:`${fmtN(a)}×${fmtN(b)}×${fmtN(c)}`};}
  }
  return R;
}
function effectivePerContainer(cL,cW,cH,sku,opt={}){
  const r=calcMixedDetailed(cL,cW,cH,sku.L,sku.W,sku.H,opt);
  let volQty=r.total,wtQty=Infinity;
  if(sku.weight>0 && opt.maxWeight>0) wtQty=Math.floor(opt.maxWeight/sku.weight);
  const eff=Math.min(volQty,wtQty);
  return{...r,volQty,wtQty:wtQty===Infinity?null:wtQty,eff,constraint:wtQty<volQty?"Weight":"Volume"};
}
function calcMixed(cL,cW,cH,sl,sw,sh){const r=calcMixedDetailed(cL,cW,cH,sl,sw,sh);return{total:r.total,orient:r.orient};}

// ─── MULTI-SKU PACKING ────────────────────────────────────────────────────────
// 8 distinct colours for up to 8 SKUs in the 3D view
const MULTI_PALETTE=[0xbe185d,0x0ea5e9,0x16a34a,0xd97706,0x7c3aed,0x0891b2,0xf97316,0x374151];
const MULTI_LABELS=["#be185d","#0ea5e9","#16a34a","#d97706","#7c3aed","#0891b2","#f97316","#374151"];

function calcMultiSKU(cL,cW,cH,skus,opt={}){
  // skus = [{name,L,W,H,targetQty}, ...]
  // Divides container into N strips (one per SKU) along each axis.
  // Tries L, W, H cuts and returns the axis with best total utilization.
  const valid=skus.filter(s=>s.L>0&&s.W>0&&s.H>0&&s.targetQty>0);
  if(!valid.length) return null;

  const cVol=cL*cW*cH;
  // Volume each SKU needs to hit its target qty
  const vNeed=valid.map(s=>s.targetQty*s.L*s.W*s.H);
  const vTotal=vNeed.reduce((a,b)=>a+b,0)||1;
  const fracs=vNeed.map(v=>v/vTotal);

  // Try dividing along one axis, proportional to volume fractions
  function tryAxis(axLen,mkReg){
    const regs=[];let off=0;
    valid.forEach((sku,i)=>{
      const isLast=i===valid.length-1;
      // Snap allocation to nearest box dimension on that axis for cleaner packing
      const snap=bestFitDetailed(
        mkReg(0,axLen).L,mkReg(0,axLen).W,mkReg(0,axLen).H,
        sku.L,sku.W,sku.H,opt);
      const boxOnAxis=axLen===cL?snap.boxL:axLen===cW?snap.boxW:snap.boxH;
      let len=isLast?axLen-off:Math.round(fracs[i]*axLen);
      if(boxOnAxis>0&&!isLast){
        const snapped=Math.max(boxOnAxis,Math.round(len/boxOnAxis)*boxOnAxis);
        len=Math.min(snapped,axLen-off-(valid.length-i-1)*Math.max(1,boxOnAxis));
      }
      len=Math.max(1,Math.min(len,axLen-off));
      const r=mkReg(off,len);
      const det=bestFitDetailed(r.L,r.W,r.H,sku.L,sku.W,sku.H,opt);
      regs.push({...sku,fitted:det.count,target:sku.targetQty,
        fillRate:Math.min(1,det.count/Math.max(1,sku.targetQty)),
        det,off:r.off,regionDims:{L:r.L,W:r.W,H:r.H},
        skuVol:(det.count*sku.L*sku.W*sku.H)/(r.L*r.W*r.H)});
      off+=len;
    });
    const fVol=regs.reduce((s,r)=>s+r.fitted*r.L*r.W*r.H,0);
    const avgFill=regs.reduce((s,r)=>s+r.fillRate,0)/regs.length;
    const score=fVol/cVol+avgFill*0.3;
    return{regs,score,total:regs.reduce((s,r)=>s+r.fitted,0),volUtil:fVol/cVol};
  }

  const attempts=[
    tryAxis(cL,(off,len)=>({L:len,W:cW,H:cH,off:{x:off,y:0,z:0}})),
    tryAxis(cW,(off,len)=>({L:cL,W:len,H:cH,off:{x:0,y:off,z:0}})),
    tryAxis(cH,(off,len)=>({L:cL,W:cW,H:len,off:{x:0,y:0,z:off}})),
  ];
  const best=attempts.reduce((a,b)=>a.score>b.score?a:b);
  return{...best,cL,cW,cH,regions:best.regs};
}

// Two-SKU guillotine split with constraint support.
// priority: "strict" = honour ratio exactly (may waste space)
//           "balanced" = stay close to ratio but fill gaps (default)
//           "maxfill" = pack as many boxes as possible, ratio is a soft guide
function calcTwoSKU(cL,cW,cH,s1,s2,r1,r2,opt={}){
  const priority=opt.priority||"balanced";
  const weight=priority==="strict"?40:priority==="maxfill"?0.4:4;
  const targetR=r1/r2;
  let bestScore=-1,best={};
  const STEPS=80;
  const axes=[
    {name:"Length",total:cL,r1:(p)=>({L:p,W:cW,H:cH}),r2:(p)=>({L:cL-p,W:cW,H:cH}),off:(p)=>({x:p,y:0,z:0})},
    {name:"Width", total:cW,r1:(p)=>({L:cL,W:p,H:cH}),r2:(p)=>({L:cL,W:cW-p,H:cH}),off:(p)=>({x:0,y:p,z:0})},
    {name:"Height",total:cH,r1:(p)=>({L:cL,W:cW,H:p}),r2:(p)=>({L:cL,W:cW,H:cH-p}),off:(p)=>({x:0,y:0,z:p})},
  ];
  function getCands(total){const s=new Set();
    for(let i=1;i<STEPS;i++) s.add((i/STEPS)*total);
    [s1.L,s1.W,s1.H,s2.L,s2.W,s2.H].forEach(d=>{for(let n=1;n*d<total;n++) s.add(n*d);});
    return Array.from(s).filter(p=>p>0&&p<total);}
  // q1det = packing of SKU1 in its region, q2det = packing of SKU2 in its region
  function consider(q1det,q2det,reg1,reg1off,reg2,reg2off,axName,p){
    const q1=q1det.count,q2=q2det.count;
    if(q1<=0||q2<=0) return;
    let useS1,useS2,score;
    if(priority==="strict"){
      // round down to whole ratio sets
      const sets=Math.min(Math.floor(q1/r1),Math.floor(q2/r2));
      if(sets<=0) return;
      useS1=sets*r1;useS2=sets*r2;score=useS1+useS2;
    }else{
      // use ALL boxes that physically fit, score by total with ratio penalty
      useS1=q1;useS2=q2;
      const actualR=q1/q2;
      const dev=Math.abs(Math.log(actualR/targetR)); // 0 = perfect ratio
      score=(q1+q2)/(1+weight*dev);
    }
    if(score>bestScore){bestScore=score;
      best={axis:axName,pos:p,totalSKU1:useS1,totalSKU2:useS2,
        sku1Region:{...reg1,det:q1det,off:reg1off,label:"SKU 1"},
        sku2Region:{...reg2,det:q2det,off:reg2off,label:"SKU 2"}};}
  }
  axes.forEach(ax=>{getCands(ax.total).forEach(p=>{
    const A=ax.r1(p),B=ax.r2(p),off=ax.off(p);
    // SKU1 in A, SKU2 in B
    consider(bestFitDetailed(A.L,A.W,A.H,s1.L,s1.W,s1.H,opt),
             bestFitDetailed(B.L,B.W,B.H,s2.L,s2.W,s2.H,opt),
             A,{x:0,y:0,z:0},B,off,ax.name,p);
    // SKU2 in A, SKU1 in B
    consider(bestFitDetailed(B.L,B.W,B.H,s1.L,s1.W,s1.H,opt),
             bestFitDetailed(A.L,A.W,A.H,s2.L,s2.W,s2.H,opt),
             B,off,A,{x:0,y:0,z:0},ax.name,p);
  });});
  if(!best.axis) return{total:0,totalSKU1:0,totalSKU2:0,volUtil:0,axis:"—",pos:0,sku1Region:null,sku2Region:null,cL,cW,cH,s1,s2,r1,r2,targetR};
  const tS1=best.totalSKU1,tS2=best.totalSKU2;
  const volUtil=(tS1*s1.L*s1.W*s1.H+tS2*s2.L*s2.W*s2.H)/(cL*cW*cH);
  return{...best,total:tS1+tS2,volUtil,cL,cW,cH,s1,s2,r1,r2,targetR};
}
function fmtN(v){return Number.isInteger(v)?String(v):parseFloat(v.toFixed(1)).toString();}
function money(v){return v.toLocaleString(undefined,{maximumFractionDigits:2});}

// ─── UPGRADE / EMAIL MODAL ────────────────────────────────────────────────────
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
          {["🚚 Shipment Planner — multi-container for big orders",
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

// ─── PRO GATE (shown instead of a locked feature) ─────────────────────────────
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

// ─── CONTAINER SELECTOR ──────────────────────────────────────────────────────
function ContainerSelector({onChange,showWeight,vehicleOnly}){
  const[type,setType]=useState("vehicle");
  const[palletIdx,setPalletIdx]=useState(0);const[palletH,setPalletH]=useState("");
  const[vehIdx,setVehIdx]=useState(4);
  const[cL,setCL]=useState("");const[cW,setCW]=useState("");const[cH,setCH]=useState("");const[maxWt,setMaxWt]=useState("");
  useEffect(()=>{
    let L=0,W=0,H=0,name="";
    if(type==="pallet"){const b=PALLET_BASES[palletIdx];L=b.L;W=b.W;H=parseFloat(palletH)||0;name="Pallet "+b.label;}
    else{const v=VEHICLES_WITH_CUSTOM[vehIdx];if(v.L){L=v.L;W=v.W;H=v.H;name=v.label;}
      else{L=parseFloat(cL)||0;W=parseFloat(cW)||0;H=parseFloat(cH)||0;name="Custom Container";}}
    onChange(L,W,H,parseFloat(maxWt)||0,name);
  },[type,palletIdx,palletH,vehIdx,cL,cW,cH,maxWt]);
  const sel={width:"100%",border:"1px solid #d1d9e0",borderRadius:"8px",padding:"8px 12px",fontSize:"13px",background:"#fff",outline:"none",cursor:"pointer"};
  return(
    <div style={S.card}>
      <div style={S.cardTitle}>{vehicleOnly?"🚛 Vehicle / Container":"🗃️ Container / Pallet"}</div>
      {!vehicleOnly&&<div style={{display:"flex",gap:"8px",marginBottom:"14px"}}>
        {["vehicle","pallet"].map(t=>(
          <button key={t} onClick={()=>setType(t)} style={{flex:1,padding:"8px",border:"none",borderRadius:"8px",cursor:"pointer",
            fontWeight:"600",fontSize:"13px",background:type===t?"#059669":"#f1f5f9",color:type===t?"#fff":"#374151"}}>
            {t==="vehicle"?"🚛 Vehicle / Container":"📦 Pallet"}</button>))}
      </div>}
      {type==="pallet"&&!vehicleOnly&&(<>
        <label style={S.label}>Pallet Base Size</label>
        <select style={{...sel,marginBottom:"12px"}} value={palletIdx} onChange={e=>setPalletIdx(+e.target.value)}>
          {PALLET_BASES.map((b,i)=><option key={i} value={i}>{b.label}</option>)}</select>
        <label style={S.label}>Stack Height (mm)</label>
        <input style={{...S.input,marginBottom:"8px"}} type="number" min="0" placeholder="e.g. 1800" value={palletH} onChange={e=>setPalletH(e.target.value)}/>
      </>)}
      {type==="vehicle"&&(<>
        <label style={S.label}>Vehicle / Container Type</label>
        <select style={{...sel,marginBottom:"12px"}} value={vehIdx} onChange={e=>setVehIdx(+e.target.value)}>
          {VEHICLES_WITH_CUSTOM.map((v,i)=><option key={i} value={i}>{v.label}</option>)}</select>
        {VEHICLES_WITH_CUSTOM[vehIdx].L&&<div style={S.infoBox}>L={VEHICLES_WITH_CUSTOM[vehIdx].L} · W={VEHICLES_WITH_CUSTOM[vehIdx].W} · H={VEHICLES_WITH_CUSTOM[vehIdx].H} mm</div>}
        {!VEHICLES_WITH_CUSTOM[vehIdx].L&&(<div style={S.grid3}>
          {[["Length (mm)",cL,setCL],["Width (mm)",cW,setCW],["Height (mm)",cH,setCH]].map(([l,v,s])=>(
            <div key={l}><label style={S.label}>{l}</label>
              <input style={S.input} type="number" min="0" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div>)}
      </>)}
      {showWeight&&<div style={{marginTop:"12px"}}>
        <label style={S.label}>Max Weight Capacity (kg) — optional</label>
        <input style={S.input} type="number" min="0" value={maxWt} onChange={e=>setMaxWt(e.target.value)} placeholder="e.g. 10000"/>
      </div>}
    </div>
  );
}

// ─── SVG 2D VIEWS ─────────────────────────────────────────────────────────────
const RC={main:{fill:"#93c5fd",stroke:"#1d4ed8",label:"Main Grid"},l1:{fill:"#fdba74",stroke:"#ea580c",label:"Leftover 1 (side)"},
  l2:{fill:"#86efac",stroke:"#16a34a",label:"Leftover 2 (front)"},l3:{fill:"#d8b4fe",stroke:"#7c3aed",label:"Leftover 3 (top)"}};
function BoxGrid({offX,offY,nx,ny,bW,bH,sc,color,dimLabel}){
  if(!nx||!ny||!bW||!bH) return null;const cells=[];const total=nx*ny,skip=total>2000?Math.ceil(total/2000):1;
  for(let iy=0;iy<ny;iy++)for(let ix=0;ix<nx;ix++){if((iy*nx+ix)%skip!==0&&!(ix===0&&iy===0)) continue;
    const x=(offX+ix*bW)*sc+1,y=(offY+iy*bH)*sc+1,w=bW*sc-1,h=bH*sc-1;
    cells.push(<g key={`${ix}-${iy}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={color.fill} stroke={color.stroke} strokeWidth="0.6"/>
      {ix===0&&iy===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle"
        fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{dimLabel}</text>}</g>);}
  return<>{cells}</>;}
function SvgDefs(){return(<defs><marker id="arr" markerWidth="4" markerHeight="4" refX="2" refY="2" orient="auto"><path d="M0,0 L4,2 L0,4 Z" fill="#475569"/></marker></defs>);}
function DimArrow({x1,y1,x2,y2,label,pos="top"}){const mx=(x1+x2)/2,my=(y1+y2)/2;
  return(<g><line x1={x1} y1={y1} x2={x2} y2={y2} stroke="#475569" strokeWidth="1" markerEnd="url(#arr)" markerStart="url(#arr)"/>
    <text x={mx+(pos==="left"?-6:0)} y={my+(pos==="top"?-10:0)} textAnchor="middle" fontSize="9" fill="#475569" fontWeight="600">{label}</text></g>);}
const VL={viewTitle:{fontSize:"13px",fontWeight:"700",color:"#1a2332",marginBottom:"8px",textAlign:"center"},note:{fontSize:"10px",color:"#9ca3af",marginTop:"4px",fontStyle:"italic",textAlign:"center"}};
function TopView2D({result}){const{cL,cW,nx,ny,boxL,boxW,usedL,usedW,leftover1:l1,leftover2:l2}=result;
  const sc=Math.min(300/cL,300/cW,20),W=cL*sc,H=cW*sc,P=28;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Top View (L × W)</div>
    <svg width={W+P+10} height={H+P+10} style={{display:"block",margin:"0 auto"}}><SvgDefs/>
      <rect x={P} y={P} width={W} height={H} fill="#f8fafc" stroke="#1e293b" strokeWidth="2"/>
      <g transform={`translate(${P},${P})`}>
        <BoxGrid offX={0} offY={0} nx={nx} ny={ny} bW={boxL} bH={boxW} sc={sc} color={RC.main} dimLabel={`${fmtN(boxL)}×${fmtN(boxW)}`}/>
        {l1.count>0&&<BoxGrid offX={usedL} offY={0} nx={l1.nx} ny={l1.ny} bW={l1.boxL} bH={l1.boxW} sc={sc} color={RC.l1} dimLabel={`${fmtN(l1.boxL)}×${fmtN(l1.boxW)}`}/>}
        {l2.count>0&&<BoxGrid offX={0} offY={usedW} nx={l2.nx} ny={l2.ny} bW={l2.boxL} bH={l2.boxW} sc={sc} color={RC.l2} dimLabel={`${fmtN(l2.boxL)}×${fmtN(l2.boxW)}`}/>}
        {usedL<cL&&<line x1={usedL*sc} y1={0} x2={usedL*sc} y2={H} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
        {usedW<cW&&<line x1={0} y1={usedW*sc} x2={W} y2={usedW*sc} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
      </g>
      <DimArrow x1={P} y1={P-14} x2={P+W} y2={P-14} label={fmtN(cL)} pos="top"/>
      <DimArrow x1={P-16} y1={P} x2={P-16} y2={P+H} label={fmtN(cW)} pos="left"/>
    </svg><div style={VL.note}>L3 (top gap) not shown — see Side View</div></div>);}
function SideView2D({result}){const{cL,cH,nx,nz,boxL,boxH,usedL,usedH,leftover1:l1,leftover3:l3}=result;
  const sc=Math.min(300/cL,300/cH,20),W=cL*sc,H=cH*sc,P=28;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Side View (L × H)</div>
    <svg width={W+P+10} height={H+P+10} style={{display:"block",margin:"0 auto"}}><SvgDefs/>
      <rect x={P} y={P} width={W} height={H} fill="#f8fafc" stroke="#1e293b" strokeWidth="2"/>
      <g transform={`translate(${P},${P})`}>
        {Array.from({length:nz},(_,iz)=>Array.from({length:nx},(_,ix)=>{const x=ix*boxL*sc+1,y=H-(iz+1)*boxH*sc+1,w=boxL*sc-1,h=boxH*sc-1;
          return(<g key={`m-${iz}-${ix}`}><rect x={x} y={y} width={w} height={h} fill={RC.main.fill} stroke={RC.main.stroke} strokeWidth="0.6"/>
            {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(boxL)}×{fmtN(boxH)}</text>}</g>);})).flat()}
        {l1.count>0&&Array.from({length:l1.nz},(_,iz)=>Array.from({length:l1.nx},(_,ix)=>{const x=(usedL+ix*l1.boxL)*sc+1,y=H-(iz+1)*l1.boxH*sc+1,w=l1.boxL*sc-1,h=l1.boxH*sc-1;
          return(<g key={`l1-${iz}-${ix}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={RC.l1.fill} stroke={RC.l1.stroke} strokeWidth="0.6"/>
            {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(l1.boxL)}×{fmtN(l1.boxH)}</text>}</g>);})).flat()}
        {l3.count>0&&Array.from({length:l3.nz},(_,iz)=>Array.from({length:l3.nx},(_,ix)=>{const x=ix*l3.boxL*sc+1,y=H-(usedH+(iz+1)*l3.boxH)*sc+1,w=l3.boxL*sc-1,h=l3.boxH*sc-1;
          return(<g key={`l3-${iz}-${ix}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={RC.l3.fill} stroke={RC.l3.stroke} strokeWidth="0.6"/>
            {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(l3.boxL)}×{fmtN(l3.boxH)}</text>}</g>);})).flat()}
        {usedL<cL&&<line x1={usedL*sc} y1={0} x2={usedL*sc} y2={H} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
        {usedH<cH&&<line x1={0} y1={H-usedH*sc} x2={W} y2={H-usedH*sc} stroke="#64748b" strokeWidth="1" strokeDasharray="3,2"/>}
      </g>
      <DimArrow x1={P} y1={P-14} x2={P+W} y2={P-14} label={fmtN(cL)} pos="top"/>
      <DimArrow x1={P-16} y1={P} x2={P-16} y2={P+H} label={fmtN(cH)} pos="left"/>
    </svg><div style={VL.note}>L2 (front gap) not shown — see Top View</div></div>);}
function IsoView2D({result}){const{cL,cW,cH,nx,ny,nz,boxL,boxW,boxH,usedL,usedW,usedH,leftover1:l1,leftover2:l2,leftover3:l3}=result;
  const MAX=260;let sc=MAX/Math.max(cL,cW,cH);if(sc>14)sc=14;if(sc<0.1)sc=0.1;
  const c30=0.8660254;function ix(x,y){return(x-y)*sc*c30;}function iy(x,y,z){return(x+y)*sc*0.5-z*sc;}
  const corners=[[0,0,0],[cL,0,0],[0,cW,0],[cL,cW,0],[0,0,cH],[cL,0,cH],[0,cW,cH],[cL,cW,cH]];
  const xs=corners.map(([x,y])=>ix(x,y)),ys=corners.map(([x,y,z])=>iy(x,y,z));
  const minX=Math.min(...xs),minY=Math.min(...ys),svgW=Math.max(...xs)-minX+20,svgH=Math.max(...ys)-minY+20,ox=-minX+10,oy=-minY+10;
  function px(x,y){return ix(x,y)+ox;}function py(x,y,z){return iy(x,y,z)+oy;}
  function blk(x0,y0,z0,x1,y1,z1,tC,fC,rC){if(x0>=x1||y0>=y1||z0>=z1)return null;
    const T=`${px(x0,y0,z1).toFixed(1)},${py(x0,y0,z1).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)} ${px(x1,y1,z1).toFixed(1)},${py(x1,y1,z1).toFixed(1)} ${px(x0,y1,z1).toFixed(1)},${py(x0,y1,z1).toFixed(1)}`;
    const F=`${px(x0,y0,z0).toFixed(1)},${py(x0,y0,z0).toFixed(1)} ${px(x1,y0,z0).toFixed(1)},${py(x1,y0,z0).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)} ${px(x0,y0,z1).toFixed(1)},${py(x0,y0,z1).toFixed(1)}`;
    const R=`${px(x1,y0,z0).toFixed(1)},${py(x1,y0,z0).toFixed(1)} ${px(x1,y1,z0).toFixed(1)},${py(x1,y1,z0).toFixed(1)} ${px(x1,y1,z1).toFixed(1)},${py(x1,y1,z1).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)}`;
    return<g opacity={0.88}><polygon points={T} fill={tC} stroke="#1e293b" strokeWidth="0.7"/><polygon points={F} fill={fC} stroke="#1e293b" strokeWidth="0.7"/><polygon points={R} fill={rC} stroke="#1e293b" strokeWidth="0.7"/></g>;}
  function edge(x0,y0,z0,x1,y1,z1,dash=false){return<line x1={px(x0,y0).toFixed(1)} y1={py(x0,y0,z0).toFixed(1)} x2={px(x1,y1).toFixed(1)} y2={py(x1,y1,z1).toFixed(1)} stroke="#1e293b" strokeWidth="1.5" strokeDasharray={dash?"4,3":"none"}/>;}
  const mX=usedL,mY=usedW,mZ=usedH;
  const l1X=usedL+l1.nx*l1.boxL,l1Y=l1.ny*l1.boxW,l1Z=l1.nz*l1.boxH;
  const l2X=l2.nx*l2.boxL,l2Y=usedW+l2.ny*l2.boxW,l2Z=l2.nz*l2.boxH;
  const l3X=l3.nx*l3.boxL,l3Y=l3.ny*l3.boxW,l3Z=usedH+l3.nz*l3.boxH;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Isometric View</div>
    <svg width={svgW} height={svgH} style={{display:"block",margin:"0 auto"}}>
      {l2.count>0&&blk(0,mY,0,l2X,l2Y,l2Z,"#bbf7d0","#86efac","#4ade80")}
      {l3.count>0&&blk(0,0,mZ,l3X,l3Y,l3Z,"#e9d5ff","#c084fc","#a855f7")}
      {l1.count>0&&blk(mX,0,0,l1X,l1Y,l1Z,"#fed7aa","#fb923c","#f97316")}
      {nx>0&&ny>0&&nz>0&&blk(0,0,0,mX,mY,mZ,"#bfdbfe","#60a5fa","#3b82f6")}
      {edge(0,0,0,cL,0,0)}{edge(0,0,0,0,cW,0)}{edge(0,0,0,0,0,cH)}{edge(cL,0,0,cL,cW,0)}{edge(cL,0,0,cL,0,cH)}
      {edge(0,cW,0,cL,cW,0)}{edge(0,cW,0,0,cW,cH)}{edge(0,0,cH,cL,0,cH)}{edge(0,0,cH,0,cW,cH)}
      {edge(cL,cW,0,cL,cW,cH)}{edge(cL,0,cH,cL,cW,cH)}{edge(0,cW,cH,cL,cW,cH)}{edge(0,cW,0,0,cW,cH,true)}
    </svg>
    <div style={{display:"flex",gap:"10px",justifyContent:"center",flexWrap:"wrap",marginTop:"6px"}}>
      {[["#60a5fa","Main"],["#fb923c","L1"],["#86efac","L2"],["#c084fc","L3"]].map(([c,l])=>(
        <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:c,borderRadius:"2px"}}/>{l}</div>))}</div></div>);}

// ─── 3D VIEWER ────────────────────────────────────────────────────────────────
const MAX_3D=1200;
function ThreeViewer({result,captureRef,regions3D}){
  const mountRef=useRef(null);const cleanRef=useRef(null);
  const regs=regions3D||[
    {col:0x3b82f6,ox:0,oy:0,oz:0,rnx:result.nx,rny:result.ny,rnz:result.nz,bL:result.boxL,bW:result.boxW,bH:result.boxH},
    {col:0xf97316,ox:result.leftover1.offX,oy:result.leftover1.offY,oz:result.leftover1.offZ,rnx:result.leftover1.nx,rny:result.leftover1.ny,rnz:result.leftover1.nz,bL:result.leftover1.boxL,bW:result.leftover1.boxW,bH:result.leftover1.boxH},
    {col:0x22c55e,ox:result.leftover2.offX,oy:result.leftover2.offY,oz:result.leftover2.offZ,rnx:result.leftover2.nx,rny:result.leftover2.ny,rnz:result.leftover2.nz,bL:result.leftover2.boxL,bW:result.leftover2.boxW,bH:result.leftover2.boxH},
    {col:0xa855f7,ox:result.leftover3.offX,oy:result.leftover3.offY,oz:result.leftover3.offZ,rnx:result.leftover3.nx,rny:result.leftover3.ny,rnz:result.leftover3.nz,bL:result.leftover3.boxL,bW:result.leftover3.boxW,bH:result.leftover3.boxH},
  ];
  const{cL,cW,cH}=result;
  useEffect(()=>{
    if(!mountRef.current) return;
    if(cleanRef.current){cleanRef.current();cleanRef.current=null;}
    const t=setTimeout(()=>{
      const el=mountRef.current;if(!el) return;const W=el.clientWidth||800,H=380;
      const scene=new THREE.Scene();scene.background=new THREE.Color(0xeef2f7);
      const camera=new THREE.PerspectiveCamera(50,W/H,0.01,100000);
      const renderer=new THREE.WebGLRenderer({antialias:true,preserveDrawingBuffer:true});
      renderer.setSize(W,H);renderer.setPixelRatio(Math.min(window.devicePixelRatio,2));el.appendChild(renderer.domElement);
      if(captureRef) captureRef.current=()=>{try{return renderer.domElement.toDataURL("image/png");}catch(e){return null;}};
      scene.add(new THREE.AmbientLight(0xffffff,0.85));
      const d=new THREE.DirectionalLight(0xffffff,0.7);d.position.set(5,8,5);scene.add(d);
      const d2=new THREE.DirectionalLight(0x8888ff,0.35);d2.position.set(-3,-2,-2);scene.add(d2);
      const tot=regs.reduce((s,r)=>s+((r.rnx||0)*(r.rny||0)*(r.rnz||0)),0)||1;
      regs.forEach(r=>{if(!r.rnx||!r.rny||!r.rnz||!r.bL||!r.bW||!r.bH) return;
        const total=r.rnx*r.rny*r.rnz;const cap=Math.max(1,Math.round(MAX_3D*(total/tot)));
        const stride=total>cap?total/cap:1;const pos=[];
        let idx=0,nextPick=0;
        for(let iz=0;iz<r.rnz;iz++)for(let iy=0;iy<r.rny;iy++)for(let ix=0;ix<r.rnx;ix++){
          // Always include boundary boxes so the full extent of the region is visible
          const isEdge=(ix===0||ix===r.rnx-1||iy===0||iy===r.rny-1||iz===0||iz===r.rnz-1);
          if(isEdge||idx>=nextPick){
            pos.push([(r.ox||0)+ix*r.bL+r.bL/2,(r.oz||0)+iz*r.bH+r.bH/2,(r.oy||0)+iy*r.bW+r.bW/2]);
            if(idx>=nextPick) nextPick+=stride;
          }
          idx++;
        }
        const mesh=new THREE.InstancedMesh(new THREE.BoxGeometry(r.bL*0.88,r.bH*0.88,r.bW*0.88),
          new THREE.MeshPhongMaterial({color:r.col,shininess:50,transparent:true,opacity:0.92}),pos.length);
        mesh.count=pos.length;const dummy=new THREE.Object3D();
        pos.forEach(([x,y,z],i)=>{dummy.position.set(x,y,z);dummy.updateMatrix();mesh.setMatrixAt(i,dummy.matrix);});
        mesh.instanceMatrix.needsUpdate=true;scene.add(mesh);});
      const cw=new THREE.LineSegments(new THREE.EdgesGeometry(new THREE.BoxGeometry(cL,cH,cW)),new THREE.LineBasicMaterial({color:0x1e293b}));
      cw.position.set(cL/2,cH/2,cW/2);scene.add(cw);
      const grid=new THREE.GridHelper(Math.max(cL,cW)*2.5,10,0xb0bec5,0xcfd8dc);grid.position.set(cL/2,-0.01,cW/2);scene.add(grid);
      const center=new THREE.Vector3(cL/2,cH/2,cW/2),diag=Math.sqrt(cL*cL+cW*cW+cH*cH)||10;
      let radius=diag*2.2,theta=Math.PI*0.38,phi=Math.PI*0.30;
      function updCam(){camera.position.set(center.x+radius*Math.sin(phi)*Math.sin(theta),center.y+radius*Math.cos(phi),center.z+radius*Math.sin(phi)*Math.cos(theta));camera.lookAt(center);}
      updCam();let drag=false,prev={x:0,y:0};const cv=renderer.domElement;cv.style.cursor="grab";cv.style.display="block";
      const onD=(e)=>{drag=true;prev={x:e.clientX,y:e.clientY};cv.style.cursor="grabbing";};
      const onM=(e)=>{if(!drag)return;theta-=(e.clientX-prev.x)*0.007;phi=Math.max(0.05,Math.min(1.5,phi+(e.clientY-prev.y)*0.007));prev={x:e.clientX,y:e.clientY};updCam();};
      const onU=()=>{drag=false;cv.style.cursor="grab";};
      const onW=(e)=>{e.preventDefault();radius=Math.max(diag*0.4,Math.min(diag*6,radius+e.deltaY*0.5));updCam();};
      const onRz=()=>{const nW=el.clientWidth||800;renderer.setSize(nW,H);camera.aspect=nW/H;camera.updateProjectionMatrix();};
      cv.addEventListener("mousedown",onD);window.addEventListener("mousemove",onM);window.addEventListener("mouseup",onU);
      cv.addEventListener("wheel",onW,{passive:false});window.addEventListener("resize",onRz);
      let animId;const loop=()=>{animId=requestAnimationFrame(loop);renderer.render(scene,camera);};loop();
      cleanRef.current=()=>{cancelAnimationFrame(animId);cv.removeEventListener("mousedown",onD);window.removeEventListener("mousemove",onM);
        window.removeEventListener("mouseup",onU);cv.removeEventListener("wheel",onW);window.removeEventListener("resize",onRz);
        if(el.contains(cv))el.removeChild(cv);renderer.dispose();};
    },100);
    return()=>{clearTimeout(t);if(cleanRef.current){cleanRef.current();cleanRef.current=null;}};
  },[result,regions3D]);
  return(<div ref={mountRef} style={{width:"100%",height:"380px",borderRadius:"10px",overflow:"hidden",border:"1px solid #e2e8f0",background:"#eef2f7"}}/>);}

function ConstraintsPanel({weight,setWeight,noStack,setNoStack,lockHeight,setLockHeight,maxStack,setMaxStack,hideWeight}){
  return(<div style={S.card}><div style={S.cardTitle}>⚙️ Constraints (optional)</div>
    <div style={{display:"grid",gridTemplateColumns:hideWeight?"1fr":"1fr 1fr",gap:"12px",marginBottom:"12px"}}>
      {!hideWeight&&<div><label style={S.label}>Weight per box (kg)</label>
        <input style={S.input} type="number" min="0" step="any" value={weight} onChange={e=>setWeight(e.target.value)} placeholder="0"/></div>}
      <div><label style={S.label}>Max stack height (layers)</label>
        <input style={S.input} type="number" min="0" step="1" value={maxStack} onChange={e=>setMaxStack(e.target.value)} placeholder="unlimited" disabled={noStack}/></div>
    </div>
    <div style={{display:"flex",gap:"16px",flexWrap:"wrap"}}>
      <label style={{display:"flex",alignItems:"center",gap:"6px",fontSize:"13px",color:"#374151",cursor:"pointer"}}>
        <input type="checkbox" checked={noStack} onChange={e=>setNoStack(e.target.checked)}/>🚫 Fragile — no stacking</label>
      <label style={{display:"flex",alignItems:"center",gap:"6px",fontSize:"13px",color:"#374151",cursor:"pointer"}}>
        <input type="checkbox" checked={lockHeight} onChange={e=>setLockHeight(e.target.checked)}/>⬆️ This side up — lock height</label>
    </div></div>);}

// ─── BOX PACKING TOOL (free) ─────────────────────────────────────────────────
function BoxPackingTool(){
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);const[maxWt,setMaxWt]=useState(0);
  const[sl,setSl]=useState("");const[sw,setSw]=useState("");const[sh,setSh]=useState("");
  const[weight,setWeight]=useState("");const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);const[maxStack,setMaxStack]=useState("");
  const[result,setResult]=useState(null);const[error,setError]=useState("");const[view,setView]=useState("3d");
  const calc=()=>{const sln=parseFloat(sl)||0,swn=parseFloat(sw)||0,shn=parseFloat(sh)||0;
    if(sln<=0||swn<=0||shn<=0){setError("Enter all 3 small box dimensions > 0.");return;}
    if(bl<=0||bw<=0||bh<=0){setError("Select a container / enter all dimensions.");return;}setError("");
    const opt={noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt,weight:parseFloat(weight)||0};
    const r=calcMixedDetailed(bl,bw,bh,sln,swn,shn,opt);
    let volQty=r.total,wtQty=null,effQty=volQty,constraint="Volume";const wt=parseFloat(weight)||0;
    if(wt>0&&maxWt>0){wtQty=Math.floor(maxWt/wt);if(wtQty<volQty){effQty=wtQty;constraint="Weight";}}
    setResult({...r,volQty,wtQty,effQty,constraint,volUtil:(effQty*sln*swn*shn)/(bl*bw*bh),totalWeight:wt>0?effQty*wt:null});};
  return(<div>
    <div style={S.sectionDesc}>Select container → enter box dimensions → add optional weight & stacking constraints → calculate. Shows 3D + 2D views with leftovers and effective quantity.</div>
    <ContainerSelector onChange={(L,W,H,wt)=>{setBl(L);setBw(W);setBh(H);setMaxWt(wt);}} showWeight={true}/>
    <div style={S.card}><div style={S.cardTitle}>📦 Small Box Dimensions</div>
      <div style={S.grid3}>{[["Length",sl,setSl],["Width",sw,setSw],["Height",sh,setSh]].map(([l,v,s])=>(
        <div key={l}><label style={S.label}>{l}</label><input style={S.input} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div></div>
    <ConstraintsPanel weight={weight} setWeight={setWeight} noStack={noStack} setNoStack={setNoStack} lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>
    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={calc}>▶ Calculate & Visualise</button>
    {result&&(<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"12px",margin:"20px 0 16px"}}>
        {[["Effective Qty",result.effQty.toLocaleString(),"#f0fdf4","#166534"],
          ["Limited By",result.constraint,result.constraint==="Weight"?"#fff7ed":"#eff6ff",result.constraint==="Weight"?"#c2410c":"#1d4ed8"],
          ["Volume Util.",(result.volUtil*100).toFixed(1)+"%",result.volUtil>=0.75?"#f0fdf4":result.volUtil>=0.5?"#fefce8":"#fff8fc",result.volUtil>=0.75?"#166534":result.volUtil>=0.5?"#854d0e":"#831843"],
          ["Total Weight",result.totalWeight!=null?money(result.totalWeight)+" kg":"—","#f8fafc","#374151"],
        ].map(([l,v,bg,col])=>(<div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
          <div style={{fontSize:"16px",fontWeight:"700",color:col,wordBreak:"break-word"}}>{v}</div>
          <div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}
      </div>
      <div style={{...S.card,marginBottom:"16px"}}><div style={S.cardTitle}>Packing Breakdown by Region (volume fit)</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"10px"}}>
          {[["#3b82f6","Main Grid",`${result.nx}×${result.ny}×${result.nz}`,result.nx*result.ny*result.nz],
            ["#f97316","Leftover 1","Side gap",result.leftover1.count],["#22c55e","Leftover 2","Front gap",result.leftover2.count],
            ["#a855f7","Leftover 3","Top gap",result.leftover3.count]].map(([color,title,sub,cnt])=>(
            <div key={title} style={{background:"#f8fafc",borderRadius:"8px",padding:"12px",textAlign:"center",borderLeft:`4px solid ${color}`}}>
              <div style={{fontSize:"20px",fontWeight:"700",color:"#1e293b"}}>{cnt}</div>
              <div style={{fontSize:"12px",fontWeight:"600",color:"#374151"}}>{title}</div>
              <div style={{fontSize:"11px",color:"#9ca3af"}}>{sub}</div></div>))}</div></div>
      <div style={{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center",flexWrap:"wrap"}}>
        <div style={{display:"flex",gap:"6px",flex:1}}>
          {[["3d","🔄 3D Model"],["2d","📐 2D Views"]].map(([id,label])=>(
            <button key={id} onClick={()=>setView(id)} style={{padding:"8px 18px",border:"none",borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",background:view===id?"#059669":"#f1f5f9",color:view===id?"#fff":"#374151"}}>{label}</button>))}</div>
        <WAShare message={`📦 *PackWise Packing Result*\nContainer: ${fmtN(result.cL)}×${fmtN(result.cW)}×${fmtN(result.cH)} mm\nBox: ${fmtN(result.sl)}×${fmtN(result.sw)}×${fmtN(result.sh)} mm\n*Max boxes: ${result.effQty.toLocaleString()}*\nBest orientation: ${result.orient}\nSpace used: ${(result.volUtil*100).toFixed(1)}%\n\nCalculate your load free at packwise.netlify.app`}/>
      </div>
      {view==="3d"&&<div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}><div style={S.cardTitle}>🔄 3D Model</div><span style={{fontSize:"11px",color:"#94a3b8"}}>Drag to rotate · Scroll to zoom</span></div>
        <ThreeViewer result={result}/>
        <div style={{display:"flex",gap:"12px",flexWrap:"wrap",marginTop:"8px",justifyContent:"center"}}>
          {[["#3b82f6","Main"],["#f97316","Leftover 1"],["#22c55e","Leftover 2"],["#a855f7","Leftover 3"],["#1e293b","Container"]].map(([c,l])=>(
            <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:c,borderRadius:"2px"}}/>{l}</div>))}</div></div>}
      {view==="2d"&&<div style={S.card}><div style={S.cardTitle}>📐 2D Engineering Views</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"20px",alignItems:"start"}}>
          <TopView2D result={result}/><SideView2D result={result}/><IsoView2D result={result}/></div></div>}
    </>)}
  </div>);}

// ─── SHIPMENT PLANNER (PRO) ──────────────────────────────────────────────────
function ShipmentPlanner(){
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);const[maxWt,setMaxWt]=useState(0);const[contName,setContName]=useState("");
  const[sl,setSl]=useState("");const[sw,setSw]=useState("");const[sh,setSh]=useState("");
  const[weight,setWeight]=useState("");const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);const[maxStack,setMaxStack]=useState("");
  const[orderQty,setOrderQty]=useState("");const[freightCost,setFreightCost]=useState("");const[brand,setBrand]=useState("");
  const[result,setResult]=useState(null);const[error,setError]=useState("");const[compare,setCompare]=useState(null);const[freightByType,setFreightByType]=useState({});
  const captureRef=useRef(null);
  const opt=()=>({noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt});
  const calc=()=>{const sku={L:parseFloat(sl)||0,W:parseFloat(sw)||0,H:parseFloat(sh)||0,weight:parseFloat(weight)||0};const order=parseInt(orderQty)||0;
    if(sku.L<=0||sku.W<=0||sku.H<=0){setError("Enter all box dimensions.");return;}
    if(bl<=0||bw<=0||bh<=0){setError("Select a container.");return;}if(order<=0){setError("Enter total order quantity.");return;}setError("");
    const e=effectivePerContainer(bl,bw,bh,sku,opt());if(e.eff<=0){setError("Box doesn't fit in this container.");return;}
    const perContainer=e.eff,containers=Math.ceil(order/perContainer),lastFill=order-(containers-1)*perContainer;
    const freight=parseFloat(freightCost)||0,totalFreight=freight*containers,costPerUnit=freight>0?totalFreight/order:null;
    const usedVol=(perContainer*sku.L*sku.W*sku.H)/(bl*bw*bh);
    setResult({sku,order,perContainer,containers,lastFill,freight,totalFreight,costPerUnit,packResult:e,usedVol,contName,maxWt,
      totalWeight:sku.weight>0?order*sku.weight:null,constraint:e.constraint});
    const rows=VEHICLES.map(v=>{const ev=effectivePerContainer(v.L,v.W,v.H,sku,{noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt});
      const per=ev.eff;if(per<=0) return{...v,per:0,containers:0,fc:0,total:0,cpu:null};
      const cont=Math.ceil(order/per),fc=parseFloat(freightByType[v.label])||0;
      return{...v,per,containers:cont,fc,total:fc*cont,cpu:fc>0?(fc*cont)/order:null};});
    setCompare(rows);};
  const updateFreight=(label,val)=>{const nf={...freightByType,[label]:val};setFreightByType(nf);
    if(compare){setCompare(compare.map(r=>{if(r.label!==label) return r;const fc=parseFloat(val)||0;
      return{...r,fc,total:fc*r.containers,cpu:fc>0&&r.containers>0?(fc*r.containers)/(parseInt(orderQty)||1):null};}));}};
  const packForViewer=result?result.packResult:null;
  const cheapest=compare?compare.filter(r=>r.cpu!=null).sort((a,b)=>a.cpu-b.cpu)[0]:null;
  const exportPDF=()=>{if(!result) return;const img=captureRef.current?captureRef.current():null;
    const doc=new jsPDF({unit:"pt",format:"a4"});const W=doc.internal.pageSize.getWidth();
    doc.setFillColor(15,23,42);doc.rect(0,0,W,64,"F");doc.setTextColor(255,255,255);doc.setFontSize(18);doc.setFont(undefined,"bold");
    doc.text(brand||"Container Loading Plan",40,36);doc.setFontSize(10);doc.setFont(undefined,"normal");
    doc.text("Shipment Loading & Cost Report  ·  Generated "+new Date().toLocaleDateString(),40,52);
    let y=88;doc.setTextColor(30,41,59);doc.setFontSize(13);doc.setFont(undefined,"bold");doc.text("Shipment Summary",40,y);y+=18;
    doc.setFontSize(10);doc.setFont(undefined,"normal");
    const lines=[["Container / Vehicle",result.contName],["Box dimensions (L×W×H)",`${fmtN(result.sku.L)} × ${fmtN(result.sku.W)} × ${fmtN(result.sku.H)} mm`],
      ["Total order quantity",result.order.toLocaleString()+" units"],["Units per container",result.perContainer.toLocaleString()],
      ["Containers required",result.containers.toLocaleString()],["Volume utilization",(result.usedVol*100).toFixed(1)+"%"],["Limited by",result.constraint]];
    if(result.totalWeight!=null) lines.push(["Total shipment weight",money(result.totalWeight)+" kg"]);
    if(result.costPerUnit!=null){lines.push(["Freight cost per container",money(result.freight)]);lines.push(["Total freight cost",money(result.totalFreight)]);lines.push(["Cost per unit shipped",money(result.costPerUnit)]);}
    lines.forEach(([k,v])=>{doc.setFont(undefined,"bold");doc.text(k+":",48,y);doc.setFont(undefined,"normal");doc.text(String(v),230,y);y+=15;});y+=6;
    if(img){try{doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Loading Visualization (single container)",40,y);y+=10;doc.addImage(img,"PNG",40,y,260,140);y+=160;}catch(e){}}
    if(y>700){doc.addPage();y=50;}doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Per-Container Manifest",40,y);y+=18;doc.setFontSize(9);
    doc.setFillColor(84,130,53);doc.setTextColor(255,255,255);doc.rect(40,y-10,W-80,16,"F");
    doc.text("Container #",48,y);doc.text("Units Loaded",180,y);doc.text("Fill %",320,y);doc.text("Weight (kg)",420,y);y+=14;doc.setTextColor(30,41,59);
    for(let i=1;i<=result.containers;i++){if(y>790){doc.addPage();y=50;}
      const units=i<result.containers?result.perContainer:result.lastFill,fill=((units/result.perContainer)*100).toFixed(0)+"%",wt=result.sku.weight>0?money(units*result.sku.weight):"—";
      if(i%2===0){doc.setFillColor(245,247,250);doc.rect(40,y-10,W-80,14,"F");}
      doc.text("#"+i,48,y);doc.text(units.toLocaleString(),180,y);doc.text(fill,320,y);doc.text(wt,420,y);y+=14;}y+=16;
    if(y>700){doc.addPage();y=50;}doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Loading Instructions",40,y);y+=18;doc.setFont(undefined,"normal");doc.setFontSize(10);
    const instr=[`1. Orient each box as ${result.packResult.boxL}×${result.packResult.boxW}×${result.packResult.boxH} mm (main grid).`,
      `2. Build the floor layer: ${result.packResult.nx} along length × ${result.packResult.ny} across width.`,
      noStack?"3. DO NOT STACK — fragile items, single layer only.":`3. Stack ${result.packResult.nz} layer(s) high.`,
      lockHeight?"4. Keep 'This Side Up' — do not rotate boxes onto their sides.":"4. Boxes may be rotated to best fit.",
      "5. Fill remaining gaps with leftover-region boxes as shown in the 3D view.",
      `6. Each full container holds ${result.perContainer.toLocaleString()} units; the last holds ${result.lastFill.toLocaleString()}.`];
    instr.forEach(t=>{const split=doc.splitTextToSize(t,W-90);doc.text(split,48,y);y+=split.length*13;});y+=14;
    if(y>760){doc.addPage();y=50;}doc.setFontSize(8);doc.setTextColor(120,120,120);
    const disc=doc.splitTextToSize("Disclaimer: This loading plan is a planning aid generated by an algorithmic estimate. Actual loading may vary due to handling, load securing, weight distribution, and real-world constraints. Verify against vehicle and safety regulations before dispatch.",W-80);
    doc.text(disc,40,y);doc.save("Loading_Plan_"+(brand||"Shipment").replace(/\s+/g,"_")+".pdf");};
  return(<div>
    <div style={S.sectionDesc}>Plan a full shipment: enter total order quantity → get containers needed, per-container loading plan, and cost-per-unit. Compare container types and export a branded PDF.</div>
    <ContainerSelector onChange={(L,W,H,wt,name)=>{setBl(L);setBw(W);setBh(H);setMaxWt(wt);setContName(name);}} showWeight={true} vehicleOnly={true}/>
    <div style={S.card}><div style={S.cardTitle}>📦 Box Dimensions</div>
      <div style={S.grid3}>{[["Length",sl,setSl],["Width",sw,setSw],["Height",sh,setSh]].map(([l,v,s])=>(
        <div key={l}><label style={S.label}>{l}</label><input style={S.input} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div></div>
    <ConstraintsPanel weight={weight} setWeight={setWeight} noStack={noStack} setNoStack={setNoStack} lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>
    <div style={S.card}><div style={S.cardTitle}>🧾 Order & Cost</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"12px"}}>
        <div><label style={S.label}>Total Order Quantity (units)</label><input style={S.input} type="number" min="0" value={orderQty} onChange={e=>setOrderQty(e.target.value)} placeholder="e.g. 5000"/></div>
        <div><label style={S.label}>Freight Cost per Container</label><input style={S.input} type="number" min="0" value={freightCost} onChange={e=>setFreightCost(e.target.value)} placeholder="e.g. 45000"/></div>
        <div><label style={S.label}>Company Name (for PDF)</label><input style={S.input} type="text" value={brand} onChange={e=>setBrand(e.target.value)} placeholder="Your company"/></div>
      </div></div>
    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={calc}>▶ Plan Shipment</button>
    {result&&(<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",margin:"20px 0 16px"}}>
        {[["Containers Needed",result.containers.toLocaleString(),"#eff6ff","#1d4ed8"],["Units / Container",result.perContainer.toLocaleString(),"#f0fdf4","#166534"],
          ["Cost / Unit",result.costPerUnit!=null?money(result.costPerUnit):"—","#fff7ed","#c2410c"],
          ["Volume Util.",(result.usedVol*100).toFixed(1)+"%",result.usedVol>=0.75?"#f0fdf4":"#fefce8",result.usedVol>=0.75?"#166534":"#854d0e"]].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"20px",fontWeight:"700",color:col,wordBreak:"break-word"}}>{v}</div><div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}</div>
      <div style={{display:"flex",gap:"12px",marginBottom:"16px",flexWrap:"wrap"}}>
        <button style={{...S.btnPrimary,flex:1,background:"#be185d"}} onClick={exportPDF}>⬇ Download PDF Loading Plan</button>
        <WAShare message={`🚚 *PackWise Shipment Plan*\nVehicle: ${result.contName}\nBox: ${fmtN(result.sku.L)}×${fmtN(result.sku.W)}×${fmtN(result.sku.H)} mm | ${result.order.toLocaleString()} units\n*${result.containers} containers needed*\n${result.perContainer.toLocaleString()} units per container\nSpace used: ${(result.usedVol*100).toFixed(1)}%${result.costPerUnit!=null?`\nCost per unit: ${money(result.costPerUnit)}`:""}\n\nPlan your shipment at packwise.netlify.app`}/>
      </div>
      <div style={S.card}><div style={S.cardTitle}>💰 Container Comparison — find the cheapest option</div>
        <div style={{fontSize:"12px",color:"#6b7a8d",marginBottom:"12px"}}>Enter freight cost for each vehicle type. Cheapest cost-per-unit is highlighted green.</div>
        <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
          <thead><tr>{["Vehicle","Units/Container","Containers Needed","Freight Cost Each","Total Cost","Cost/Unit"].map(h=>(
            <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
          <tbody>{compare.map((r,i)=>{const isCheapest=cheapest&&r.label===cheapest.label;
            return(<tr key={i} style={{background:isCheapest?"#dcfce7":i%2===0?"#fff":"#fafbfc"}}>
              <td style={{padding:"8px 12px",fontWeight:isCheapest?"700":"500"}}>{r.label}{isCheapest&&" ✅"}</td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.per.toLocaleString()}</td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.containers.toLocaleString()}</td>
              <td style={{padding:"8px 12px"}}><input type="number" min="0" value={freightByType[r.label]||""} placeholder="0" onChange={e=>updateFreight(r.label,e.target.value)} style={{width:"90px",border:"1px solid #d1d9e0",borderRadius:"6px",padding:"4px 8px",fontSize:"12px"}}/></td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.total>0?money(r.total):"—"}</td>
              <td style={{padding:"8px 12px",textAlign:"right",fontWeight:"700",color:isCheapest?"#166534":"#374151"}}>{r.cpu!=null?money(r.cpu):"—"}</td></tr>);})}</tbody>
        </table></div>
        {cheapest&&<div style={{marginTop:"12px",padding:"10px 14px",background:"#f0fdf4",borderRadius:"8px",fontSize:"13px",color:"#166534"}}>
          ✅ Cheapest: <strong>{cheapest.label}</strong> at <strong>{money(cheapest.cpu)}</strong> per unit ({cheapest.containers} containers × {money(cheapest.fc)} each).</div>}
      </div>
      <div style={{...S.card,padding:"0",overflow:"hidden"}}>
        <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",fontWeight:"600",fontSize:"13px"}}>Per-Container Manifest</div>
        <div style={{overflowX:"auto",maxHeight:"260px",overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
          <thead><tr>{["Container #","Units Loaded","Fill %","Weight"].map(h=>(
            <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0}}>{h}</th>))}</tr></thead>
          <tbody>{Array.from({length:Math.min(result.containers,200)},(_,i)=>{const n=i+1,units=n<result.containers?result.perContainer:result.lastFill,fill=((units/result.perContainer)*100).toFixed(0);
            return(<tr key={i} style={{background:i%2===0?"#fff":"#fafbfc"}}>
              <td style={{padding:"7px 12px",fontWeight:"500"}}>#{n}</td><td style={{padding:"7px 12px",textAlign:"right"}}>{units.toLocaleString()}</td>
              <td style={{padding:"7px 12px"}}><span style={{background:+fill>=90?"#dcfce7":+fill>=50?"#fef9c3":"#fdf2f8",color:+fill>=90?"#166534":+fill>=50?"#854d0e":"#831843",padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"600"}}>{fill}%</span></td>
              <td style={{padding:"7px 12px",textAlign:"right"}}>{result.sku.weight>0?money(units*result.sku.weight)+" kg":"—"}</td></tr>);})}</tbody>
        </table>{result.containers>200&&<div style={{padding:"10px 18px",fontSize:"12px",color:"#9ca3af"}}>Showing 200 of {result.containers} — full list in PDF</div>}</div>
      </div>
      {packForViewer&&<div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}><div style={S.cardTitle}>🔄 Single Container Loading (3D)</div><span style={{fontSize:"11px",color:"#94a3b8"}}>This view is included in the PDF</span></div>
        <ThreeViewer result={packForViewer} captureRef={captureRef}/></div>}
    </>)}
  </div>);}

// ─── CONTAINER SKU TOOL (free, limited) ──────────────────────────────────────
function ContainerSkuTool({isPro,onUpgrade}){
  const[cL,setCL]=useState("");const[cW,setCW]=useState("");const[cH,setCH]=useState("");const[cMaxWt,setCMaxWt]=useState("");
  const[fileName,setFileName]=useState("");const[results,setResults]=useState(null);const[error,setError]=useState("");
  const[processing,setProcessing]=useState(false);const[progress,setProgress]=useState(0);const[skuCount,setSkuCount]=useState(0);
  const[dragOver,setDragOver]=useState(false);const[rawSkus,setRawSkus]=useState(null);const[capped,setCapped]=useState(false);
  const container={cL:parseFloat(cL)||0,cW:parseFloat(cW)||0,cH:parseFloat(cH)||0,cMaxWt:parseFloat(cMaxWt)||0};
  const valid=container.cL>0&&container.cW>0&&container.cH>0&&container.cMaxWt>0;
  function pSkus(cont,skus){const{cL,cW,cH,cMaxWt}=cont,cv=cL*cW*cH;
    return skus.map(s=>{const{name,sl,sw,sh,swt,qtyAvail}=s;if(!name)return null;if(sl<=0||sw<=0||sh<=0)return{name,error:"Invalid dimensions"};
      const{total:vQ,orient}=calcMixed(cL,cW,cH,sl,sw,sh);let eV=qtyAvail>0?Math.min(vQ,qtyAvail):vQ;let wQ=swt>0?Math.floor(cMaxWt/swt):null;
      if(wQ!==null&&qtyAvail>0)wQ=Math.min(wQ,qtyAvail);const eQ=wQ!==null?Math.min(eV,wQ):eV,vu=(eQ*sl*sw*sh)/cv,wu=swt>0?(eQ*swt)/cMaxWt:null;
      let con="Volume";if(wQ!==null&&wQ<eV)con="Weight";if(qtyAvail>0&&eQ===qtyAvail)con="Stock Limit";
      return{name,volQty:eV,wtQty:wQ!==null?wQ:"N/A",effQty:eQ,volUtil:vu,wtUtil:wu,orient,constraint:con};}).filter(Boolean);}
  const parseFile=(file)=>{if(!file)return;setFileName(file.name);setError("");setResults(null);
    const r=new FileReader();r.onload=(e)=>{try{const wb=XLSX.read(e.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});let ds=0;
      for(let i=0;i<Math.min(raw.length,20);i++){const row=raw[i].map(c=>String(c).toLowerCase());if(row.some(c=>c.includes("sku")||c.includes("length")||c.includes("width"))){ds=i+1;break;}}
      const skus=[];for(let i=ds;i<raw.length;i++){const rr=raw[i];if(!rr[0]&&!rr[1])continue;
        skus.push({name:String(rr[0]||"").trim(),sl:parseFloat(rr[1])||0,sw:parseFloat(rr[2])||0,sh:parseFloat(rr[3])||0,swt:parseFloat(rr[4])||0,qtyAvail:parseFloat(rr[5])>0?parseFloat(rr[5]):-1});}
      if(skus.length===0){setError("No SKU data found.");return;}setSkuCount(skus.length);setRawSkus(skus);}catch(err){setError("Could not read file: "+err.message);}};r.readAsArrayBuffer(file);};
  const run=()=>{if(!valid){setError("Enter container dimensions.");return;}if(!rawSkus){setError("Upload SKU file.");return;}setError("");
    let toProcess=rawSkus,cap=false;
    if(!isPro&&rawSkus.length>CONFIG.freeSkuLimit){toProcess=rawSkus.slice(0,CONFIG.freeSkuLimit);cap=true;}
    setCapped(cap);setProcessing(true);setProgress(0);setResults(null);const CHUNK=500;let done=0;const all=[];
    function next(){all.push(...pSkus(container,toProcess.slice(done,done+CHUNK)));done+=CHUNK;
      setProgress(Math.round((Math.min(done,toProcess.length)/toProcess.length)*100));
      if(done<toProcess.length)setTimeout(next,0);else{setResults(all);setProcessing(false);}}setTimeout(next,50);};
  const exp=()=>{const wb=XLSX.utils.book_new();const h=["SKU Name","Max Qty (Volume)","Max Qty (Weight)","Effective Max Qty","Volume Used (%)","Weight Used (%)","Best Orientation","Constraint"];
    const rows=results.map(r=>r.error?[r.name,r.error,"","","","","",""]:[r.name,r.volQty,r.wtQty,r.effQty,(r.volUtil*100).toFixed(2)+"%",r.wtUtil!=null?(r.wtUtil*100).toFixed(2)+"%":"",r.orient,r.constraint]);
    const ws=XLSX.utils.aoa_to_sheet([["CONTAINER SKU PACKING RESULTS"],[],["Container",`${container.cL}×${container.cW}×${container.cH}`,"Max Weight",container.cMaxWt],[],h,...rows]);
    ws["!cols"]=[{wch:22},{wch:16},{wch:16},{wch:16},{wch:14},{wch:14},{wch:26},{wch:14}];ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:7}}];
    XLSX.utils.book_append_sheet(wb,ws,"Results");XLSX.writeFile(wb,"Container_Packing_Results.xlsx");};
  const dlT=()=>{const ws=XLSX.utils.aoa_to_sheet([["SKU Name","Length","Width","Height","Weight per Box","Qty Available (optional)"],["SKU-001",30,20,15,2.5,""],["SKU-002",25,25,10,1.8,500]]);
    ws["!cols"]=[{wch:18},{wch:10},{wch:10},{wch:10},{wch:16},{wch:22}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"SKU List");XLSX.writeFile(wb,"SKU_Template.xlsx");};
  const gC=results?results.filter(r=>!r.error&&r.volUtil>=0.75).length:0;const oC=results?results.filter(r=>!r.error&&r.volUtil>=0.5&&r.volUtil<0.75).length:0;const lC=results?results.filter(r=>!r.error&&r.volUtil<0.5).length:0;
  return(<div>
    <div style={S.sectionDesc}>Upload a list of SKUs with dimensions and weight. Finds maximum quantity per SKU constrained by volume and weight, with Excel download.{!isPro&&<span style={{color:"#c2410c"}}> Free plan processes up to {CONFIG.freeSkuLimit} SKUs.</span>}</div>
    <div style={{display:"grid",gridTemplateColumns:"320px 1fr",gap:"20px",alignItems:"start"}}>
      <div>
        <div style={S.card}><div style={S.cardTitle}>🗃️ Container Details</div>
          <div style={S.grid2}>{[["Length",cL,setCL],["Width",cW,setCW],["Height",cH,setCH],["Max Weight",cMaxWt,setCMaxWt]].map(([l,v,s])=>(
            <div key={l}><label style={S.label}>{l}</label><input style={S.input} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div>
          {valid&&<div style={S.infoBox}>Volume: {(container.cL*container.cW*container.cH).toLocaleString()}</div>}</div>
        <div style={S.card}><div style={S.cardTitle}>📂 Upload SKU File</div>
          <div style={S.dropzone(dragOver)} onDragOver={e=>{e.preventDefault();setDragOver(true)}} onDragLeave={()=>setDragOver(false)} onDrop={e=>{e.preventDefault();setDragOver(false);parseFile(e.dataTransfer.files[0])}} onClick={()=>document.getElementById("fi2").click()}>
            <div style={{fontSize:"28px",marginBottom:"6px"}}>📂</div><div style={{fontSize:"13px",fontWeight:"500",color:"#374151"}}>{fileName||"Drop Excel or click to browse"}</div>
            <div style={{fontSize:"11px",color:"#9ca3af",marginTop:"4px"}}>.xlsx or .xls</div>
            {rawSkus&&<div style={{marginTop:"6px",fontSize:"12px",color:"#059669",fontWeight:"600"}}>✓ {skuCount.toLocaleString()} SKUs loaded</div>}
            <input id="fi2" type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>parseFile(e.target.files[0])}/></div>
          <div style={S.noteBox}><strong>Columns:</strong> SKU Name | L | W | H | Weight | Qty</div></div>
        <button style={S.btnPrimary} onClick={run}>▶ Calculate Container Fit</button>
        <button style={{...S.btnSecondary,marginTop:"10px"}} onClick={dlT}>⬇ Download Template</button>
      </div>
      <div>
        {error&&<div style={S.error}>⚠ {error}</div>}
        {capped&&<div style={{...S.card,background:"#fffbeb",border:"1px solid #fde68a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:"13px",color:"#92400e"}}>⚠ Free plan limited to {CONFIG.freeSkuLimit} SKUs. {skuCount.toLocaleString()} uploaded — only first {CONFIG.freeSkuLimit} processed.</span>
          <button onClick={onUpgrade} style={{padding:"6px 14px",background:"#059669",color:"#fff",border:"none",borderRadius:"8px",fontWeight:"600",fontSize:"12px",cursor:"pointer",whiteSpace:"nowrap"}}>⭐ Go Pro</button></div>}
        {processing&&<div style={S.card}><div style={{fontSize:"13px",fontWeight:"500",color:"#374151",marginBottom:"8px"}}>Processing {Math.min(skuCount,isPro?skuCount:CONFIG.freeSkuLimit).toLocaleString()} SKUs... {progress}%</div>
          <div style={{background:"#e5e7eb",borderRadius:"99px",height:"10px"}}><div style={{height:"10px",borderRadius:"99px",background:"#059669",width:`${progress}%`,transition:"width 0.2s"}}/></div></div>}
        {results&&!processing&&(<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"12px",marginBottom:"16px"}}>
            {[["Total",results.length,"#f8fafc"],["≥75%",gC,"#f0fdf4"],["50–74%",oC,"#fefce8"],["<50%",lC,"#fff8fc"]].map(([l,v,bg])=>(
              <div key={l} style={{background:bg,borderRadius:"10px",padding:"12px",textAlign:"center",border:"1px solid rgba(0,0,0,0.06)"}}>
                <div style={{fontSize:"20px",fontWeight:"700",color:"#1a2332"}}>{v}</div><div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"2px"}}>{l}</div></div>))}</div>
          <button style={{...S.btnPrimary,marginBottom:"16px"}} onClick={exp}>⬇ Download Results as Excel</button>
          <div style={{...S.card,padding:"0",overflow:"hidden"}}>
            <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:"600",fontSize:"13px"}}>Preview</span><span style={{fontSize:"12px",color:"#9ca3af"}}>{results.length.toLocaleString()} SKUs</span></div>
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
              <thead><tr>{["SKU","Vol Qty","Wt Qty","Eff Qty","Vol%","Wt%","Orientation","Constraint"].map(h=>(<th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
              <tbody>{results.slice(0,100).map((r,i)=>r.error?(<tr key={i} style={{background:"#fff8fc"}}><td style={{padding:"8px 12px",fontWeight:"500"}}>{r.name}</td><td style={{padding:"8px 12px",color:"#be185d"}} colSpan={7}>{r.error}</td></tr>
              ):(<tr key={i} style={{background:i%2===0?"#fff":"#fafbfc"}}>
                <td style={{padding:"8px 12px",fontWeight:"500"}}>{r.name}</td><td style={{padding:"8px 12px",textAlign:"right"}}>{r.volQty?.toLocaleString()}</td>
                <td style={{padding:"8px 12px",textAlign:"right"}}>{typeof r.wtQty==="number"?r.wtQty.toLocaleString():r.wtQty}</td><td style={{padding:"8px 12px",textAlign:"right",fontWeight:"700"}}>{r.effQty?.toLocaleString()}</td>
                <td style={{padding:"8px 12px"}}><UtilBadge val={r.volUtil}/></td><td style={{padding:"8px 12px"}}><UtilBadge val={r.wtUtil}/></td>
                <td style={{padding:"8px 12px",color:"#6b7a8d",whiteSpace:"nowrap"}}>{r.orient}</td>
                <td style={{padding:"8px 12px"}}><span style={{padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"500",background:r.constraint==="Volume"?"#eff6ff":r.constraint==="Weight"?"#fff7ed":"#f5f3ff",color:r.constraint==="Volume"?"#1d4ed8":r.constraint==="Weight"?"#c2410c":"#6d28d9"}}>{r.constraint}</span></td></tr>))}</tbody>
            </table>{results.length>100&&<div style={{padding:"10px 18px",fontSize:"12px",color:"#9ca3af",borderTop:"1px solid #f1f5f9"}}>Showing 100 of {results.length.toLocaleString()} — download Excel for all</div>}</div></div>
        </>)}
        {!results&&!processing&&!error&&<div style={{...S.card,padding:"60px",textAlign:"center",color:"#9ca3af"}}><div style={{fontSize:"48px",marginBottom:"12px"}}>📦</div><div style={{fontWeight:"500"}}>Fill in container details and upload your SKU file</div></div>}
      </div>
    </div>
  </div>);}

// ─── MULTI-SKU PLANNER (PRO) ──────────────────────────────────────────────────
function MultiSKUTool(){
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);const[contName,setContName]=useState("");
  const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);const[maxStack,setMaxStack]=useState("");
  const[skuRows,setSkuRows]=useState([
    {id:1,name:"SKU 1",L:"",W:"",H:"",targetQty:""},
    {id:2,name:"SKU 2",L:"",W:"",H:"",targetQty:""},
  ]);
  const[result,setResult]=useState(null);const[error,setError]=useState("");const[view,setView]=useState("table");
  const nextId=useRef(3);

  const addRow=()=>{
    setSkuRows(r=>[...r,{id:nextId.current++,name:`SKU ${nextId.current-1}`,L:"",W:"",H:"",targetQty:""}]);};
  const removeRow=(id)=>setSkuRows(r=>r.filter(row=>row.id!==id));
  const updateRow=(id,field,val)=>setSkuRows(r=>r.map(row=>row.id===id?{...row,[field]:val}:row));

  const calc=()=>{
    if(bl<=0||bw<=0||bh<=0){setError("Select a container / enter dimensions.");return;}
    const skus=skuRows.map(r=>({...r,L:parseFloat(r.L)||0,W:parseFloat(r.W)||0,H:parseFloat(r.H)||0,targetQty:parseInt(r.targetQty)||0}));
    const valid=skus.filter(s=>s.L>0&&s.W>0&&s.H>0&&s.targetQty>0);
    if(valid.length<2){setError("Enter at least 2 SKUs with all dimensions and target quantity.");return;}
    if(valid.length>8){setError("Maximum 8 SKUs supported.");return;}
    setError("");
    const opt={noStack,lockHeight,maxStack:parseInt(maxStack)||0};
    const r=calcMultiSKU(bl,bw,bh,valid,opt);
    setResult(r?{...r,skuCount:valid.length}:null);
  };

  const exportExcel=()=>{
    if(!result) return;
    const wb=XLSX.utils.book_new();
    const h=["SKU Name","Length","Width","Height","Target Qty","Fitted Qty","Fill %","Volume Util","Orientation"];
    const rows=result.regions.map(r=>[r.name,r.L,r.W,r.H,r.target,r.fitted,
      (r.fillRate*100).toFixed(1)+"%",(r.skuVol*100).toFixed(1)+"%",
      `${fmtN(r.det.boxL)}×${fmtN(r.det.boxW)}×${fmtN(r.det.boxH)}`]);
    const ws=XLSX.utils.aoa_to_sheet([
      ["MULTI-SKU CONTAINER PACKING PLAN"],[],
      ["Container",`${fmtN(bl)}×${fmtN(bw)}×${fmtN(bh)} mm`,"Total Boxes",result.total,"Volume Util",(result.volUtil*100).toFixed(1)+"%"],[],
      h,...rows]);
    ws["!cols"]=[{wch:16},{wch:10},{wch:10},{wch:10},{wch:12},{wch:12},{wch:10},{wch:12},{wch:22}];
    XLSX.utils.book_append_sheet(wb,ws,"Multi-SKU Plan");
    XLSX.writeFile(wb,"Multi_SKU_Packing_Plan.xlsx");
  };

  // Build regions3D for ThreeViewer
  const regions3D=result?result.regions.map((r,i)=>({
    col:MULTI_PALETTE[i%MULTI_PALETTE.length],
    ox:r.off.x,oy:r.off.y,oz:r.off.z||0,
    rnx:r.det.nx,rny:r.det.ny,rnz:r.det.nz,
    bL:r.det.boxL,bW:r.det.boxW,bH:r.det.boxH,
  })):null;
  const containerResult=result?{cL:result.cL,cW:result.cW,cH:result.cH,
    nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,
    leftover1:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover2:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover3:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0}}:null;

  return(<div>
    <div style={S.sectionDesc}>
      Pack multiple SKU sizes into one container simultaneously. Enter each SKU's dimensions
      and target quantity — the planner divides the container optimally and reports exactly
      how many of each SKU fits. Supports stacking and orientation constraints.
    </div>

    <ContainerSelector onChange={(L,W,H,wt,name)=>{setBl(L);setBw(W);setBh(H);setContName(name);}}/>

    {/* SKU table */}
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"14px"}}>
        <div style={S.cardTitle}>📦 SKUs to Pack</div>
        <button onClick={addRow} disabled={skuRows.length>=8}
          style={{padding:"7px 14px",background:"#f0fdf4",border:"1px solid #bbf7d0",
          borderRadius:"8px",fontSize:"13px",fontWeight:"600",color:"#166534",cursor:"pointer",fontFamily:"inherit"}}>
          + Add SKU
        </button>
      </div>
      {/* Table header */}
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr 40px",gap:"8px",
        marginBottom:"8px",padding:"0 4px"}}>
        {["SKU Name","Length","Width","Height","Target Qty",""].map(h=>(
          <div key={h} style={S.label}>{h}</div>
        ))}
      </div>
      {/* Rows */}
      {skuRows.map((row,idx)=>(
        <div key={row.id} className="sku-row" style={{display:"grid",
          gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr 40px",gap:"8px",
          marginBottom:"8px",padding:"6px 4px",borderRadius:"8px",transition:"background 0.15s"}}>
          <input style={{...S.input,borderColor:MULTI_LABELS[idx%MULTI_LABELS.length]+"55"}}
            value={row.name} onChange={e=>updateRow(row.id,"name",e.target.value)}
            placeholder={`SKU ${idx+1}`}/>
          {["L","W","H"].map(f=>(
            <input key={f} style={S.input} type="number" min="0" step="any"
              value={row[f]} onChange={e=>updateRow(row.id,f,e.target.value)} placeholder="0"/>
          ))}
          <input style={S.input} type="number" min="0"
            value={row.targetQty} onChange={e=>updateRow(row.id,"targetQty",e.target.value)} placeholder="qty"/>
          <button onClick={()=>removeRow(row.id)} disabled={skuRows.length<=2}
            style={{background:"none",border:"1px solid #e2e8f0",borderRadius:"6px",cursor:"pointer",
            color:"#9ca3af",fontSize:"16px",fontFamily:"inherit",
            opacity:skuRows.length<=2?0.3:1}}>✕</button>
        </div>
      ))}
      <div style={{...S.noteBox,marginTop:"8px"}}>
        <strong>Tip:</strong> Target Qty is your order quantity. The planner allocates container space proportionally and reports the actual fitted quantity.
        Max 8 SKUs · Container divided into {skuRows.filter(r=>parseFloat(r.L)>0).length||"N"} regions
      </div>
    </div>

    <ConstraintsPanel hideWeight={true} noStack={noStack} setNoStack={setNoStack}
      lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>

    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={calc}>▶ Plan Multi-SKU Container Load</button>

    {result&&(<>
      {/* Summary row */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",margin:"20px 0 16px"}}>
        {[
          ["Total Boxes Fitted",result.total.toLocaleString(),"#f0fdf4","#166534"],
          ["Volume Utilization",(result.volUtil*100).toFixed(1)+"%",result.volUtil>=0.75?"#f0fdf4":result.volUtil>=0.5?"#fef9c3":"#fef2f2",result.volUtil>=0.75?"#166534":result.volUtil>=0.5?"#854d0e":"#991b1b"],
          ["SKUs Packed",`${result.regions.filter(r=>r.fitted>0).length} / ${result.regions.length}`,"#eff6ff","#1d4ed8"],
          ["Fully Filled",`${result.regions.filter(r=>r.fillRate>=1).length} / ${result.regions.length}`,"#fdf2f8","#be185d"],
        ].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"22px",fontWeight:"800",color:col,letterSpacing:"-0.02em"}}>{v}</div>
            <div style={{fontSize:"11px",color:"#6b7280",marginTop:"4px",fontWeight:"600",textTransform:"uppercase",letterSpacing:"0.05em"}}>{l}</div>
          </div>
        ))}
      </div>

      {/* View tabs */}
      <div style={{display:"flex",gap:"6px",marginBottom:"16px",flexWrap:"wrap"}}>
        {[["table","📊 Per-SKU Results"],["3d","🔄 3D Model"]].map(([id,label])=>(
          <button key={id} onClick={()=>setView(id)} style={{padding:"8px 18px",border:"none",
            borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",
            background:view===id?"#be185d":"#f1f5f9",color:view===id?"#fff":"#374151",fontFamily:"inherit"}}>{label}</button>))}
        <button onClick={exportExcel} style={{padding:"8px 18px",border:"1px solid #bbf7d0",
          borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",
          background:"#f0fdf4",color:"#166534",fontFamily:"inherit",marginLeft:"auto"}}>
          ⬇ Download Excel
        </button>
      </div>

      {/* Per-SKU results table */}
      {view==="table"&&(
        <div style={{...S.card,padding:"0",overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
              <thead><tr>
                {["#","SKU Name","Target Qty","Fitted Qty","Fill %","Region Util","Orientation","Region Size"].map(h=>(
                  <th key={h} style={{padding:"10px 14px",textAlign:"left",fontWeight:"700",fontSize:"11px",
                    color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.05em",
                    background:"#f8fafc",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {result.regions.map((r,i)=>{
                  const fillOk=r.fillRate>=1,fillMid=r.fillRate>=0.7;
                  return(<tr key={i} style={{background:i%2===0?"#fff":"#fafbfc",
                    borderBottom:"1px solid #f1f5f9"}}>
                    <td style={{padding:"10px 14px"}}>
                      <div style={{width:"12px",height:"12px",borderRadius:"3px",
                        background:MULTI_LABELS[i%MULTI_LABELS.length],display:"inline-block"}}/>
                    </td>
                    <td style={{padding:"10px 14px",fontWeight:"700",color:"#111827"}}>{r.name}</td>
                    <td style={{padding:"10px 14px",color:"#6b7280"}}>{r.target.toLocaleString()}</td>
                    <td style={{padding:"10px 14px",fontWeight:"700",
                      color:fillOk?"#166534":fillMid?"#854d0e":"#991b1b"}}>
                      {r.fitted.toLocaleString()}</td>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{background:fillOk?"#dcfce7":fillMid?"#fef9c3":"#fee2e2",
                        color:fillOk?"#166534":fillMid?"#854d0e":"#991b1b",
                        padding:"2px 8px",borderRadius:"99px",fontSize:"12px",fontWeight:"700"}}>
                        {(r.fillRate*100).toFixed(0)}%</span></td>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{background:"#eff6ff",color:"#1d4ed8",
                        padding:"2px 8px",borderRadius:"99px",fontSize:"12px",fontWeight:"600"}}>
                        {(r.skuVol*100).toFixed(0)}%</span></td>
                    <td style={{padding:"10px 14px",color:"#6b7280",fontSize:"12px",whiteSpace:"nowrap"}}>
                      {fmtN(r.det.boxL)}×{fmtN(r.det.boxW)}×{fmtN(r.det.boxH)}</td>
                    <td style={{padding:"10px 14px",color:"#9ca3af",fontSize:"12px",whiteSpace:"nowrap"}}>
                      {fmtN(r.regionDims.L)}×{fmtN(r.regionDims.W)}×{fmtN(r.regionDims.H)}</td>
                  </tr>);
                })}
              </tbody>
            </table>
          </div>
          {/* Legend */}
          <div style={{padding:"12px 16px",borderTop:"1px solid #f1f5f9",
            display:"flex",gap:"16px",flexWrap:"wrap"}}>
            {result.regions.map((r,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:"#555"}}>
                <div style={{width:"12px",height:"12px",borderRadius:"3px",
                  background:MULTI_LABELS[i%MULTI_LABELS.length]}}/>{r.name}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 3D Model */}
      {view==="3d"&&containerResult&&(
        <div style={S.card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}>
            <div style={S.cardTitle}>🔄 3D Model — Multi-SKU Loading</div>
            <span style={{fontSize:"11px",color:"#94a3b8"}}>Drag to rotate · Scroll to zoom</span>
          </div>
          <ThreeViewer result={containerResult} regions3D={regions3D}/>
          <div style={{display:"flex",gap:"12px",flexWrap:"wrap",marginTop:"10px",justifyContent:"center"}}>
            {result.regions.map((r,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}>
                <div style={{width:"11px",height:"11px",background:MULTI_LABELS[i%MULTI_LABELS.length],borderRadius:"2px"}}/>
                {r.name} ({r.fitted.toLocaleString()})</div>))}
            <div style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}>
              <div style={{width:"11px",height:"11px",background:"#1e293b",borderRadius:"2px"}}/>Container</div>
          </div>
        </div>
      )}
    </>)}
  </div>);}

function UtilBadge({val}){if(val==null||val==="")return<span style={{color:"#9ca3af",fontSize:"12px"}}>—</span>;
  const pct=typeof val==="number"?val*100:parseFloat(val);const bg=pct>=75?"#dcfce7":pct>=50?"#fef9c3":"#fdf2f8";const color=pct>=75?"#166534":pct>=50?"#854d0e":"#831843";
  return<span style={{background:bg,color,padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"600"}}>{pct.toFixed(1)}%</span>;}

const S={
  card:{background:"#fff",borderRadius:"12px",padding:"20px",marginBottom:"16px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:"1px solid #e8edf2"},
  cardTitle:{fontSize:"14px",fontWeight:"700",color:"#1a2332",marginBottom:"14px"},
  sectionDesc:{fontSize:"13px",color:"#6b7a8d",marginBottom:"20px",lineHeight:"1.6",background:"#f8fafc",borderRadius:"8px",padding:"12px 16px"},
  grid2:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px"},grid3:{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"12px"},
  label:{display:"block",fontSize:"11px",fontWeight:"600",color:"#6b7a8d",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:"4px"},
  input:{width:"100%",border:"1px solid #d1d9e0",borderRadius:"8px",padding:"8px 12px",fontSize:"14px",boxSizing:"border-box",outline:"none",fontFamily:"inherit"},
  infoBox:{background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:"8px",padding:"8px 12px",fontSize:"12px",color:"#166534",marginTop:"10px"},
  dropzone:(d)=>({border:`2px dashed ${d?"#10b981":"#d1d9e0"}`,borderRadius:"10px",padding:"24px",textAlign:"center",cursor:"pointer",transition:"all 0.2s",background:d?"#f0fdf4":"#fafbfc"}),
  noteBox:{background:"#f8fafc",borderRadius:"8px",padding:"10px 12px",fontSize:"12px",color:"#6b7a8d",marginTop:"10px",lineHeight:"1.6"},
  btnPrimary:{width:"100%",padding:"12px",background:"#059669",color:"#fff",border:"none",borderRadius:"10px",fontSize:"14px",fontWeight:"600",cursor:"pointer"},
  btnSecondary:{width:"100%",padding:"10px",background:"#fff",color:"#059669",border:"1px solid #059669",borderRadius:"10px",fontSize:"13px",fontWeight:"500",cursor:"pointer"},
  error:{background:"#fff8fc",border:"1px solid #bbf7d0",borderRadius:"10px",padding:"12px",fontSize:"13px",color:"#be185d",marginBottom:"12px"},
};

// ─── TWO-SKU PALLET TOOL ──────────────────────────────────────────────────────
const SKU_COLORS={"SKU 1":{fill:"#93c5fd",stroke:"#1d4ed8"},"SKU 2":{fill:"#fdba74",stroke:"#ea580c"}};

function TwoSKUTopView({result}){
  const{cL,cW,axis,pos,sku1Region:s1r,sku2Region:s2r}=result;
  if(!s1r||!s2r) return null;
  const sc=Math.min(300/cL,300/cW,20),W=cL*sc,H=cW*sc,P=28;
  const cutX=axis==="Length"?pos*sc:null,cutY=axis==="Width"?pos*sc:null;
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Top View (L × W)</div>
    <svg width={W+P+10} height={H+P+10} style={{display:"block",margin:"0 auto"}}><SvgDefs/>
      <rect x={P} y={P} width={W} height={H} fill="#f8fafc" stroke="#1e293b" strokeWidth="2"/>
      <g transform={`translate(${P},${P})`}>
        <BoxGrid offX={s1r.off.x} offY={s1r.off.y} nx={s1r.det.nx} ny={s1r.det.ny} bW={s1r.det.boxL} bH={s1r.det.boxW} sc={sc} color={SKU_COLORS["SKU 1"]} dimLabel={`${fmtN(s1r.det.boxL)}×${fmtN(s1r.det.boxW)}`}/>
        <BoxGrid offX={s2r.off.x} offY={s2r.off.y} nx={s2r.det.nx} ny={s2r.det.ny} bW={s2r.det.boxL} bH={s2r.det.boxW} sc={sc} color={SKU_COLORS["SKU 2"]} dimLabel={`${fmtN(s2r.det.boxL)}×${fmtN(s2r.det.boxW)}`}/>
        {cutX&&<line x1={cutX} y1={0} x2={cutX} y2={H} stroke="#ef4444" strokeWidth="2" strokeDasharray="5,3"/>}
        {cutY&&<line x1={0} y1={cutY} x2={W} y2={cutY} stroke="#ef4444" strokeWidth="2" strokeDasharray="5,3"/>}
        {axis==="Height"&&<text x={W/2} y={H/2} textAnchor="middle" fontSize="11" fill="#64748b">Height cut — see Side View</text>}
      </g>
      <DimArrow x1={P} y1={P-14} x2={P+W} y2={P-14} label={fmtN(cL)} pos="top"/>
      <DimArrow x1={P-16} y1={P} x2={P-16} y2={P+H} label={fmtN(cW)} pos="left"/>
    </svg>
    <div style={{display:"flex",gap:"10px",justifyContent:"center",marginTop:"6px",flexWrap:"wrap"}}>
      {[["SKU 1","#93c5fd","#1d4ed8"],["SKU 2","#fdba74","#ea580c"],["Cut line","#ef4444","#ef4444"]].map(([l,f,s])=>(
        <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:f,border:`1.5px solid ${s}`,borderRadius:"2px"}}/>{l}</div>))}
    </div></div>);}

function TwoSKUSideView({result}){
  const{cL,cH,axis,pos,sku1Region:s1r,sku2Region:s2r}=result;
  if(!s1r||!s2r) return null;
  const sc=Math.min(300/cL,300/cH,20),W=cL*sc,H=cH*sc,P=28;
  const cutX=axis==="Length"?pos*sc:null,cutZ=axis==="Height"?(cH-pos)*sc:null;
  function draw(reg,color){const{det,off}=reg;if(!det.nx||!det.nz||!det.boxL||!det.boxH) return null;
    return Array.from({length:det.nz},(_,iz)=>Array.from({length:det.nx},(_,ix)=>{
      const x=(off.x+ix*det.boxL)*sc+1,y=H-((off.z||0)+(iz+1)*det.boxH)*sc+1,w=det.boxL*sc-1,h=det.boxH*sc-1;
      return(<g key={`${reg.label}-${iz}-${ix}`}><rect x={x} y={y} width={Math.max(0,w)} height={Math.max(0,h)} fill={color.fill} stroke={color.stroke} strokeWidth="0.6"/>
        {ix===0&&iz===0&&w>18&&h>10&&<text x={x+w/2} y={y+h/2} textAnchor="middle" dominantBaseline="middle" fontSize={Math.max(7,Math.min(10,(Math.min(w,h)-4)/2.2))} fill="#111">{fmtN(det.boxL)}×{fmtN(det.boxH)}</text>}</g>);})).flat();}
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Side View (L × H)</div>
    <svg width={W+P+10} height={H+P+10} style={{display:"block",margin:"0 auto"}}><SvgDefs/>
      <rect x={P} y={P} width={W} height={H} fill="#f8fafc" stroke="#1e293b" strokeWidth="2"/>
      <g transform={`translate(${P},${P})`}>
        {draw(s1r,SKU_COLORS["SKU 1"])}{draw(s2r,SKU_COLORS["SKU 2"])}
        {cutX&&<line x1={cutX} y1={0} x2={cutX} y2={H} stroke="#ef4444" strokeWidth="2" strokeDasharray="5,3"/>}
        {cutZ&&<line x1={0} y1={cutZ} x2={W} y2={cutZ} stroke="#ef4444" strokeWidth="2" strokeDasharray="5,3"/>}
        {axis==="Width"&&<text x={W/2} y={H/2} textAnchor="middle" fontSize="11" fill="#64748b">Width cut — see Top View</text>}
      </g>
      <DimArrow x1={P} y1={P-14} x2={P+W} y2={P-14} label={fmtN(cL)} pos="top"/>
      <DimArrow x1={P-16} y1={P} x2={P-16} y2={P+H} label={fmtN(cH)} pos="left"/>
    </svg></div>);}

function TwoSKUIsoView({result}){
  const{cL,cW,cH,sku1Region:s1r,sku2Region:s2r}=result;
  if(!s1r||!s2r) return null;
  const MAX=260;let sc=MAX/Math.max(cL,cW,cH);if(sc>14)sc=14;if(sc<0.1)sc=0.1;
  const c30=0.8660254;function ix(x,y){return(x-y)*sc*c30;}function iy(x,y,z){return(x+y)*sc*0.5-z*sc;}
  const corners=[[0,0,0],[cL,0,0],[0,cW,0],[cL,cW,0],[0,0,cH],[cL,0,cH],[0,cW,cH],[cL,cW,cH]];
  const xs=corners.map(([x,y])=>ix(x,y)),ys=corners.map(([x,y,z])=>iy(x,y,z));
  const minX=Math.min(...xs),minY=Math.min(...ys),svgW=Math.max(...xs)-minX+20,svgH=Math.max(...ys)-minY+20,ox=-minX+10,oy=-minY+10;
  function px(x,y){return ix(x,y)+ox;}function py(x,y,z){return iy(x,y,z)+oy;}
  function blkReg(reg,tC,fC,rC){const{off,det:{nx,ny,nz,boxL,boxW,boxH}}=reg;if(!nx||!ny||!nz) return null;
    const x0=off.x,y0=off.y,z0=off.z||0,x1=x0+nx*boxL,y1=y0+ny*boxW,z1=z0+nz*boxH;
    const T=`${px(x0,y0,z1).toFixed(1)},${py(x0,y0,z1).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)} ${px(x1,y1,z1).toFixed(1)},${py(x1,y1,z1).toFixed(1)} ${px(x0,y1,z1).toFixed(1)},${py(x0,y1,z1).toFixed(1)}`;
    const F=`${px(x0,y0,z0).toFixed(1)},${py(x0,y0,z0).toFixed(1)} ${px(x1,y0,z0).toFixed(1)},${py(x1,y0,z0).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)} ${px(x0,y0,z1).toFixed(1)},${py(x0,y0,z1).toFixed(1)}`;
    const R=`${px(x1,y0,z0).toFixed(1)},${py(x1,y0,z0).toFixed(1)} ${px(x1,y1,z0).toFixed(1)},${py(x1,y1,z0).toFixed(1)} ${px(x1,y1,z1).toFixed(1)},${py(x1,y1,z1).toFixed(1)} ${px(x1,y0,z1).toFixed(1)},${py(x1,y0,z1).toFixed(1)}`;
    return<g opacity={0.9}><polygon points={T} fill={tC} stroke="#1e293b" strokeWidth="0.8"/><polygon points={F} fill={fC} stroke="#1e293b" strokeWidth="0.8"/><polygon points={R} fill={rC} stroke="#1e293b" strokeWidth="0.8"/></g>;}
  function edge(x0,y0,z0,x1,y1,z1,dash=false){return<line x1={px(x0,y0).toFixed(1)} y1={py(x0,y0,z0).toFixed(1)} x2={px(x1,y1).toFixed(1)} y2={py(x1,y1,z1).toFixed(1)} stroke="#1e293b" strokeWidth="1.5" strokeDasharray={dash?"4,3":"none"}/>;}
  return(<div style={{textAlign:"center"}}><div style={VL.viewTitle}>Isometric View</div>
    <svg width={svgW} height={svgH} style={{display:"block",margin:"0 auto"}}>
      {blkReg(s2r,"#fed7aa","#fdba74","#fb923c")}{blkReg(s1r,"#bfdbfe","#93c5fd","#60a5fa")}
      {edge(0,0,0,cL,0,0)}{edge(0,0,0,0,cW,0)}{edge(0,0,0,0,0,cH)}{edge(cL,0,0,cL,cW,0)}{edge(cL,0,0,cL,0,cH)}
      {edge(0,cW,0,cL,cW,0)}{edge(0,cW,0,0,cW,cH)}{edge(0,0,cH,cL,0,cH)}{edge(0,0,cH,0,cW,cH)}
      {edge(cL,cW,0,cL,cW,cH)}{edge(cL,0,cH,cL,cW,cH)}{edge(0,cW,cH,cL,cW,cH)}
    </svg>
    <div style={{display:"flex",gap:"10px",justifyContent:"center",marginTop:"6px"}}>
      {[["SKU 1","#93c5fd"],["SKU 2","#fdba74"]].map(([l,c])=>(
        <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:c,borderRadius:"2px"}}/>{l}</div>))}
    </div></div>);}

function TwoSKUTool(){
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);
  const[s1L,setS1L]=useState("");const[s1W,setS1W]=useState("");const[s1H,setS1H]=useState("");
  const[s2L,setS2L]=useState("");const[s2W,setS2W]=useState("");const[s2H,setS2H]=useState("");
  const[r1,setR1]=useState("1");const[r2,setR2]=useState("1");
  const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);const[maxStack,setMaxStack]=useState("");
  const[priority,setPriority]=useState("balanced");
  const[result,setResult]=useState(null);const[error,setError]=useState("");const[view,setView]=useState("2d");
  const calc=()=>{
    const s1={L:parseFloat(s1L)||0,W:parseFloat(s1W)||0,H:parseFloat(s1H)||0};
    const s2={L:parseFloat(s2L)||0,W:parseFloat(s2W)||0,H:parseFloat(s2H)||0};
    const rr1=parseInt(r1)||1,rr2=parseInt(r2)||1;
    if(s1.L<=0||s1.W<=0||s1.H<=0){setError("Enter all SKU 1 dimensions.");return;}
    if(s2.L<=0||s2.W<=0||s2.H<=0){setError("Enter all SKU 2 dimensions.");return;}
    if(bl<=0||bw<=0||bh<=0){setError("Select a container / pallet.");return;}
    if(rr1<=0||rr2<=0){setError("Ratio values must be positive.");return;}
    setError("");
    const opt={noStack,lockHeight,maxStack:parseInt(maxStack)||0,priority};
    setResult(calcTwoSKU(bl,bw,bh,s1,s2,rr1,rr2,opt));};
  const regions3D=result&&result.sku1Region&&result.sku2Region?[
    {col:0x3b82f6,ox:result.sku1Region.off.x,oy:result.sku1Region.off.y,oz:result.sku1Region.off.z||0,rnx:result.sku1Region.det.nx,rny:result.sku1Region.det.ny,rnz:result.sku1Region.det.nz,bL:result.sku1Region.det.boxL,bW:result.sku1Region.det.boxW,bH:result.sku1Region.det.boxH},
    {col:0xf97316,ox:result.sku2Region.off.x,oy:result.sku2Region.off.y,oz:result.sku2Region.off.z||0,rnx:result.sku2Region.det.nx,rny:result.sku2Region.det.ny,rnz:result.sku2Region.det.nz,bL:result.sku2Region.det.boxL,bW:result.sku2Region.det.boxW,bH:result.sku2Region.det.boxH},
  ]:null;
  const cResult=result?{cL:result.cL,cW:result.cW,cH:result.cH,nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,
    leftover1:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover2:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover3:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0}}:null;
  return(<div>
    <div style={S.sectionDesc}>Find the maximum quantity of two different SKU sizes that fit together in one container or pallet, maintaining your quantity ratio. Supports stacking limit and this-side-up constraints.</div>
    <ContainerSelector onChange={(L,W,H)=>{setBl(L);setBw(W);setBh(H);}}/>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px",marginBottom:"16px"}}>
      <div style={S.card}><div style={{...S.cardTitle,color:"#1d4ed8"}}>📦 SKU 1 Dimensions</div>
        <div style={S.grid3}>{[["Length",s1L,setS1L],["Width",s1W,setS1W],["Height",s1H,setS1H]].map(([l,v,s])=>(
          <div key={l}><label style={S.label}>{l}</label><input style={{...S.input,borderColor:"#bfdbfe"}} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div></div>
      <div style={S.card}><div style={{...S.cardTitle,color:"#ea580c"}}>📦 SKU 2 Dimensions</div>
        <div style={S.grid3}>{[["Length",s2L,setS2L],["Width",s2W,setS2W],["Height",s2H,setS2H]].map(([l,v,s])=>(
          <div key={l}><label style={S.label}>{l}</label><input style={{...S.input,borderColor:"#fed7aa"}} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div></div>
    </div>
    <div style={S.card}><div style={S.cardTitle}>⚖️ Quantity Ratio (SKU 1 : SKU 2)</div>
      <div style={{display:"flex",alignItems:"center",gap:"16px"}}>
        <div style={{flex:1}}><label style={S.label}>SKU 1 quantity</label>
          <input style={{...S.input,borderColor:"#bfdbfe",fontSize:"18px",textAlign:"center",fontWeight:"700"}} type="number" min="1" value={r1} onChange={e=>setR1(e.target.value)}/></div>
        <div style={{fontSize:"28px",fontWeight:"700",color:"#374151",paddingTop:"18px"}}>:</div>
        <div style={{flex:1}}><label style={S.label}>SKU 2 quantity</label>
          <input style={{...S.input,borderColor:"#fed7aa",fontSize:"18px",textAlign:"center",fontWeight:"700"}} type="number" min="1" value={r2} onChange={e=>setR2(e.target.value)}/></div>
        <div style={{flex:2,paddingTop:"18px"}}><div style={{background:"#f8fafc",borderRadius:"8px",padding:"10px 14px",fontSize:"12px",color:"#6b7a8d"}}>
          For every <strong>{r1}</strong> of SKU 1, pack <strong>{r2}</strong> of SKU 2. The container is split to best honour this ratio.</div></div>
      </div>
      <div style={{marginTop:"16px",borderTop:"1px solid #f1f5f9",paddingTop:"14px"}}>
        <label style={S.label}>Ratio Priority — how strictly to honour the ratio</label>
        <div style={{display:"flex",gap:"8px",marginTop:"6px"}}>
          {[["strict","🎯 Strict","Exact ratio (may leave gaps)"],
            ["balanced","⚖️ Balanced","Close to ratio, fills gaps"],
            ["maxfill","📦 Max Fill","Most boxes, ratio is a guide"]].map(([id,label,desc])=>(
            <button key={id} onClick={()=>setPriority(id)} style={{flex:1,padding:"10px",border:priority===id?"2px solid #059669":"1px solid #d1d9e0",
              borderRadius:"8px",cursor:"pointer",background:priority===id?"#f0fdf4":"#fff",textAlign:"center"}}>
              <div style={{fontWeight:"700",fontSize:"13px",color:priority===id?"#166534":"#374151"}}>{label}</div>
              <div style={{fontSize:"10px",color:"#6b7a8d",marginTop:"2px"}}>{desc}</div>
            </button>))}
        </div>
      </div></div>
    <ConstraintsPanel hideWeight={true} noStack={noStack} setNoStack={setNoStack} lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>
    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={calc}>▶ Find Best Two-SKU Arrangement</button>
    {result&&(result.total===0?(
      <div style={{...S.card,marginTop:"20px",background:"#fff8fc",border:"1px solid #bbf7d0"}}>
        <div style={{fontWeight:"600",color:"#be185d"}}>No valid arrangement found.</div>
        <div style={{fontSize:"13px",color:"#6b7a8d",marginTop:"4px"}}>The boxes may be too large, or the ratio cannot be achieved within constraints.</div>
      </div>
    ):(<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"14px",margin:"20px 0 16px"}}>
        {[["SKU 1 Quantity",result.totalSKU1.toLocaleString(),"#eff6ff","#1d4ed8"],
          ["SKU 2 Quantity",result.totalSKU2.toLocaleString(),"#fff7ed","#c2410c"],
          ["Total Boxes",result.total.toLocaleString(),"#f0fdf4","#166534"]].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"24px",fontWeight:"700",color:col}}>{v}</div><div style={{fontSize:"12px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",marginBottom:"16px"}}>
        {[["Target Ratio",`${result.r1}:${result.r2}`,"#eff6ff","#1d4ed8"],
          ["Achieved Ratio",`${result.totalSKU1}:${result.totalSKU2}`,"#f8fafc","#374151"],
          ["Volume Utilization",(result.volUtil*100).toFixed(1)+"%",result.volUtil>=0.75?"#f0fdf4":result.volUtil>=0.5?"#fefce8":"#fff8fc",result.volUtil>=0.75?"#166534":result.volUtil>=0.5?"#854d0e":"#831843"],
          ["Split Axis",`${result.axis} @ ${fmtN(result.pos)}`,"#f5f3ff","#6d28d9"]].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"16px",fontWeight:"700",color:col,wordBreak:"break-word"}}>{v}</div><div style={{fontSize:"12px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}
      </div>
      <div style={{...S.card,marginBottom:"16px",background:"#f8fafc"}}><div style={S.cardTitle}>📋 Arrangement Details</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px"}}>
          {[result.sku1Region,result.sku2Region].map((reg,i)=>{const col=i===0?"#eff6ff":"#fff7ed",bdr=i===0?"#bfdbfe":"#fed7aa",tcol=i===0?"#1d4ed8":"#c2410c";
            return(<div key={i} style={{background:col,borderRadius:"8px",padding:"14px",border:`1px solid ${bdr}`}}>
              <div style={{fontWeight:"700",color:tcol,marginBottom:"8px"}}>SKU {i+1} Region</div>
              <div style={{fontSize:"12px",color:"#374151",lineHeight:"1.8"}}>
                <div>Qty: <strong>{i===0?result.totalSKU1:result.totalSKU2}</strong></div>
                <div>Grid: <strong>{reg.det.nx}×{reg.det.ny}×{reg.det.nz}</strong></div>
                <div>Box orientation: <strong>{fmtN(reg.det.boxL)}×{fmtN(reg.det.boxW)}×{fmtN(reg.det.boxH)}</strong></div>
                <div>Region size: <strong>{fmtN(reg.L)}×{fmtN(reg.W)}×{fmtN(reg.H)}</strong></div>
              </div></div>);})}
        </div></div>
      <div style={{display:"flex",gap:"6px",marginBottom:"16px"}}>
        {[["2d","📐 2D Views"],["3d","🔄 3D Model"]].map(([id,label])=>(
          <button key={id} onClick={()=>setView(id)} style={{padding:"8px 18px",border:"none",borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",background:view===id?"#059669":"#f1f5f9",color:view===id?"#fff":"#374151"}}>{label}</button>))}
      </div>
      {view==="2d"&&<div style={S.card}><div style={S.cardTitle}>📐 2D Views — Two-SKU Arrangement</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"20px",alignItems:"start"}}>
          <TwoSKUTopView result={result}/><TwoSKUSideView result={result}/><TwoSKUIsoView result={result}/></div></div>}
      {view==="3d"&&cResult&&<div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}><div style={S.cardTitle}>🔄 3D Model — Two-SKU</div><span style={{fontSize:"11px",color:"#94a3b8"}}>Drag to rotate · Scroll to zoom</span></div>
        <ThreeViewer result={cResult} regions3D={regions3D}/>
        <div style={{display:"flex",gap:"12px",flexWrap:"wrap",marginTop:"8px",justifyContent:"center"}}>
          {[["#3b82f6","SKU 1"],["#f97316","SKU 2"],["#1e293b","Container"]].map(([c,l])=>(
            <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:c,borderRadius:"2px"}}/>{l}</div>))}</div></div>}
    </>))}
  </div>);}

// ─── WEBSITE PAGES ────────────────────────────────────────────────────────────

// ─── UTILITY COMPONENTS ───────────────────────────────────────────────────────

// Scroll-triggered fade-in (supports stagger class)
function FadeIn({children,className="",style={},stagger=false}){
  const ref=useRef(null);
  useEffect(()=>{
    const el=ref.current;if(!el) return;
    const obs=new IntersectionObserver(([e])=>{
      if(e.isIntersecting){el.classList.add("in");obs.disconnect();}
    },{threshold:0.1,rootMargin:"0px 0px -40px 0px"});
    obs.observe(el);return()=>obs.disconnect();
  },[]);
  return <div ref={ref} className={`${stagger?"stagger":"fade-up"} ${className}`} style={style}>{children}</div>;
}

// Animated count-up number (triggers on scroll into view)
function CountUp({value,suffix="",prefix="",duration=1600}){
  const[n,setN]=useState(0);const ref=useRef(null);const done=useRef(false);
  useEffect(()=>{
    const el=ref.current;if(!el) return;
    const obs=new IntersectionObserver(([e])=>{
      if(e.isIntersecting&&!done.current){
        done.current=true;const start=Date.now();
        const tick=()=>{
          const p=Math.min((Date.now()-start)/duration,1);
          const eased=1-Math.pow(1-p,3);
          setN(Math.round(value*eased));
          if(p<1) requestAnimationFrame(tick);
        };requestAnimationFrame(tick);obs.disconnect();}
    },{threshold:0.5});
    obs.observe(el);return()=>obs.disconnect();
  },[value,duration]);
  return <span ref={ref}>{prefix}{n}{suffix}</span>;
}

// WhatsApp share button
function WAShare({message}){
  const url=`https://wa.me/?text=${encodeURIComponent(message)}`;
  return(
    <a href={url} target="_blank" rel="noopener noreferrer" className="wa-btn">
      <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24">
        <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/>
      </svg>
      Share on WhatsApp
    </a>
  );
}

// ─── SVG LOGO ────────────────────────────────────────────────────────────────
function PackWiseLogo({size=36}){
  return(
    <svg width={size} height={size} viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
      <polygon points="18,3 33,11.5 18,20 3,11.5" fill="#0ea5e9"/>
      <polygon points="3,11.5 18,20 18,33 3,24.5" fill="#be185d"/>
      <polygon points="18,20 33,11.5 33,24.5 18,33" fill="#064e3b"/>
      <line x1="18" y1="3" x2="18" y2="20" stroke="rgba(255,255,255,0.2)" strokeWidth="0.7"/>
      <line x1="10.5" y1="7.3" x2="25.5" y2="15.7" stroke="rgba(255,255,255,0.13)" strokeWidth="0.7"/>
      <polyline points="3,11.5 18,3 33,11.5" stroke="rgba(255,255,255,0.28)" strokeWidth="0.9" fill="none"/>
      <polygon points="18,7.5 25,11.5 18,15.5 11,11.5" fill="rgba(255,255,255,0.18)"/>
    </svg>
  );
}

// ── Nav ──
function Nav({page,setPage,isPro,onUpgrade,onLogout}){
  const[menuOpen,setMenuOpen]=useState(false);
  const links=[["home","Home"],["tool","Calculator"],["pricing","Pricing"],["about","About"]];
  const go=(id)=>{setPage(id);setMenuOpen(false);};
  return(
    <nav style={{position:"sticky",top:0,zIndex:200,background:"rgba(255,255,255,0.97)",
      backdropFilter:"blur(8px)",borderBottom:"1px solid #bbf7d0",
      boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"0 24px",
        display:"flex",alignItems:"center",justifyContent:"space-between",height:"60px"}}>
        <div onClick={()=>go("home")} style={{display:"flex",alignItems:"center",gap:"10px",cursor:"pointer",flexShrink:0}}>
          <PackWiseLogo size={36}/>
          <div>
            <div style={{fontWeight:"900",fontSize:"17px",color:"#111827",lineHeight:1,letterSpacing:"-0.02em"}}>PackWise</div>
            <div style={{fontSize:"10px",color:"#6b7280",letterSpacing:"0.08em",fontWeight:"600"}}>PACKING INTELLIGENCE</div>
          </div>
        </div>
        <div className="nav-desktop">
          {links.map(([id,label])=>(
            <button key={id} onClick={()=>go(id)} style={{padding:"7px 14px",border:"none",
              background:page===id?"#fdf2f8":"none",color:page===id?"#be185d":"#475569",
              fontWeight:page===id?"700":"500",fontSize:"14px",borderRadius:"8px",cursor:"pointer",fontFamily:"inherit"}}>
              {label}</button>))}
          <div style={{width:"1px",height:"20px",background:"#e2e8f0",margin:"0 6px"}}/>
          {isPro?(
            <span onClick={onLogout} title="Click to sign out of Pro" style={{background:"#fdf2f8",
              color:"#be185d",fontWeight:"700",fontSize:"12px",padding:"6px 14px",
              borderRadius:"99px",cursor:"pointer",border:"1px solid #fbcfe8"}}>⭐ PRO</span>
          ):(
            <button onClick={()=>{go("tool");setTimeout(onUpgrade,100);}} className="btn-primary"
              style={{padding:"8px 18px",background:"linear-gradient(135deg,#be185d,#9d174d)",
              color:"#fff",border:"none",borderRadius:"8px",fontWeight:"700",fontSize:"14px",
              cursor:"pointer",boxShadow:"0 2px 8px rgba(190,24,93,0.35)",fontFamily:"inherit"}}>
              ⭐ Get Pro</button>
          )}
        </div>
        <button className="nav-burger" onClick={()=>setMenuOpen(o=>!o)} aria-label="Menu">
          <div style={{width:"22px",display:"flex",flexDirection:"column",gap:"5px"}}>
            {[0,1,2].map(i=>(<div key={i} style={{height:"2px",background:"#374151",borderRadius:"2px",
              transition:"all 0.2s",
              transform:menuOpen&&i===0?"rotate(45deg) translate(5px,5px)":menuOpen&&i===2?"rotate(-45deg) translate(5px,-5px)":"none",
              opacity:menuOpen&&i===1?0:1}}/>))}
          </div>
        </button>
      </div>
      {menuOpen&&(
        <div style={{background:"#fff",borderTop:"1px solid #bbf7d0",
          boxShadow:"0 8px 24px rgba(0,0,0,0.08)",padding:"8px 24px 16px"}}>
          {links.map(([id,label])=>(
            <button key={id} onClick={()=>go(id)} style={{display:"block",width:"100%",
              padding:"12px 8px",border:"none",background:"none",textAlign:"left",
              fontSize:"16px",fontWeight:"600",color:page===id?"#be185d":"#374151",
              cursor:"pointer",borderBottom:"1px solid #f0fdf4",fontFamily:"inherit"}}>{label}</button>))}
          <div style={{marginTop:"12px"}}>
            {isPro?(
              <span onClick={()=>{onLogout();setMenuOpen(false);}} style={{display:"inline-block",
                background:"#fdf2f8",color:"#be185d",fontWeight:"700",fontSize:"13px",
                padding:"8px 16px",borderRadius:"99px",cursor:"pointer",border:"1px solid #fbcfe8"}}>
                ⭐ PRO (tap to sign out)</span>
            ):(
              <button onClick={()=>{go("tool");setTimeout(onUpgrade,100);}} className="btn-primary"
                style={{width:"100%",padding:"12px",background:"linear-gradient(135deg,#be185d,#9d174d)",
                color:"#fff",border:"none",borderRadius:"10px",fontWeight:"700",fontSize:"15px",
                cursor:"pointer",fontFamily:"inherit"}}>⭐ Get Pro</button>
            )}
          </div>
        </div>
      )}
    </nav>
  );
}
  const[menuOpen,setMenuOpen]=useState(false);
  const links=[["home","Home"],["tool","Calculator"],["pricing","Pricing"],["about","About"]];
  const go=(id)=>{setPage(id);setMenuOpen(false);};
  return(
    <nav style={{position:"sticky",top:0,zIndex:200,background:"rgba(255,255,255,0.97)",
      backdropFilter:"blur(8px)",borderBottom:"1px solid #e2e8f0",
      boxShadow:"0 1px 8px rgba(0,0,0,0.06)"}}>
      {/* Main bar */}
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"0 24px",
        display:"flex",alignItems:"center",justifyContent:"space-between",height:"60px"}}>
        {/* Logo */}
        <div onClick={()=>go("home")} style={{display:"flex",alignItems:"center",gap:"10px",cursor:"pointer",flexShrink:0}}>
          <div style={{width:"34px",height:"34px",background:"linear-gradient(135deg,#059669,#0f172a)",
            borderRadius:"8px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px"}}>📦</div>
          <div>
            <div style={{fontWeight:"800",fontSize:"16px",color:"#0f172a",lineHeight:1}}>PackWise</div>
            <div style={{fontSize:"10px",color:"#64748b",letterSpacing:"0.06em"}}>PACKING INTELLIGENCE</div>
          </div>
        </div>
        {/* Desktop links — hidden on mobile via CSS */}
        <div className="nav-desktop">
          {links.map(([id,label])=>(
            <button key={id} onClick={()=>go(id)} style={{padding:"7px 14px",border:"none",
              background:page===id?"#f0fdf4":"none",color:page===id?"#059669":"#475569",
              fontWeight:page===id?"700":"500",fontSize:"14px",borderRadius:"8px",cursor:"pointer",
              fontFamily:"inherit"}}>
              {label}</button>))}
          <div style={{width:"1px",height:"20px",background:"#e2e8f0",margin:"0 6px"}}/>
          {isPro?(
            <span onClick={onLogout} title="Click to sign out of Pro" style={{background:"#fef3c7",
              color:"#92400e",fontWeight:"700",fontSize:"12px",padding:"6px 14px",
              borderRadius:"99px",cursor:"pointer",border:"1px solid #fde68a"}}>⭐ PRO</span>
          ):(
            <button onClick={()=>{go("tool");setTimeout(onUpgrade,100);}} className="btn-primary"
              style={{padding:"8px 18px",background:"linear-gradient(135deg,#059669,#047857)",
              color:"#fff",border:"none",borderRadius:"8px",fontWeight:"700",fontSize:"14px",
              cursor:"pointer",boxShadow:"0 2px 8px rgba(5,150,105,0.35)",fontFamily:"inherit"}}>
              ⭐ Get Pro</button>
          )}
        </div>
        {/* Hamburger — hidden on desktop via CSS */}
        <button className="nav-burger" onClick={()=>setMenuOpen(o=>!o)}
          aria-label="Menu" style={{fontFamily:"inherit"}}>
          <div style={{width:"22px",display:"flex",flexDirection:"column",gap:"5px"}}>
            {[0,1,2].map(i=>(
              <div key={i} style={{height:"2px",background:"#374151",borderRadius:"2px",
                transition:"all 0.2s",
                transform:menuOpen&&i===0?"rotate(45deg) translate(5px,5px)":
                          menuOpen&&i===2?"rotate(-45deg) translate(5px,-5px)":"none",
                opacity:menuOpen&&i===1?0:1}}/>
            ))}
          </div>
        </button>
      </div>
      {/* Mobile dropdown — only shows when open */}
      {menuOpen&&(
        <div style={{background:"#fff",borderTop:"1px solid #e2e8f0",
          boxShadow:"0 8px 24px rgba(0,0,0,0.08)",padding:"8px 24px 16px"}}>
          {links.map(([id,label])=>(
            <button key={id} onClick={()=>go(id)} style={{display:"block",width:"100%",
              padding:"12px 8px",border:"none",background:"none",textAlign:"left",
              fontSize:"16px",fontWeight:"600",color:page===id?"#059669":"#374151",
              cursor:"pointer",borderBottom:"1px solid #f1f5f9",fontFamily:"inherit"}}>
              {label}</button>))}
          <div style={{marginTop:"12px"}}>
            {isPro?(
              <span onClick={()=>{onLogout();setMenuOpen(false);}} style={{display:"inline-block",
                background:"#fef3c7",color:"#92400e",fontWeight:"700",fontSize:"13px",
                padding:"8px 16px",borderRadius:"99px",cursor:"pointer"}}>⭐ PRO (tap to sign out)</span>
            ):(
              <button onClick={()=>{go("tool");setTimeout(onUpgrade,100);}} className="btn-primary"
                style={{width:"100%",padding:"12px",background:"linear-gradient(135deg,#059669,#047857)",
                color:"#fff",border:"none",borderRadius:"10px",fontWeight:"700",
                fontSize:"15px",cursor:"pointer",fontFamily:"inherit"}}>
                ⭐ Get Pro</button>
            )}
          </div>
        </div>
      )}
    </nav>
  );
}

// ── Footer ──
function Footer({setPage}){
  return(
    <footer style={{background:"#0f172a",color:"#94a3b8",marginTop:"80px"}}>
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"48px 32px 32px"}}>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:"40px",marginBottom:"40px"}}>
          <div>
            <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"12px"}}>
              <PackWiseLogo size={32}/>
              <div style={{fontWeight:"900",fontSize:"17px",color:"#fff",letterSpacing:"-0.02em"}}>PackWise</div>
            </div>
            <p style={{fontSize:"13px",lineHeight:"1.7",maxWidth:"260px"}}>
              India's smart container loading calculator. Built for exporters, logistics teams,
              and warehouse managers who want to ship smarter and spend less.
            </p>
          </div>
          <div>
            <div style={{color:"#fff",fontWeight:"600",fontSize:"13px",marginBottom:"12px",textTransform:"uppercase",letterSpacing:"0.08em"}}>Tool</div>
            {[["tool","Box Packing"],["tool","Shipment Planner"],["tool","Two-SKU Calculator"],["tool","Container SKU"]].map(([pg,l])=>(
              <div key={l} onClick={()=>setPage(pg)} style={{fontSize:"13px",padding:"4px 0",cursor:"pointer",color:"#94a3b8"}}
                onMouseEnter={e=>e.target.style.color="#34d399"} onMouseLeave={e=>e.target.style.color="#94a3b8"}>{l}</div>))}
          </div>
          <div>
            <div style={{color:"#fff",fontWeight:"600",fontSize:"13px",marginBottom:"12px",textTransform:"uppercase",letterSpacing:"0.08em"}}>Company</div>
            {[["about","About"],["pricing","Pricing"],["about","Contact"]].map(([pg,l])=>(
              <div key={l} onClick={()=>setPage(pg)} style={{fontSize:"13px",padding:"4px 0",cursor:"pointer",color:"#94a3b8"}}
                onMouseEnter={e=>e.target.style.color="#34d399"} onMouseLeave={e=>e.target.style.color="#94a3b8"}>{l}</div>))}
          </div>
          <div>
            <div style={{color:"#fff",fontWeight:"600",fontSize:"13px",marginBottom:"12px",textTransform:"uppercase",letterSpacing:"0.08em"}}>Contact</div>
            <div style={{fontSize:"13px",lineHeight:"1.8"}}>
              <div style={{color:"#94a3b8"}}>Questions?</div>
              <a href={`mailto:${CONFIG.contactEmail}`} style={{color:"#34d399",textDecoration:"none"}}>{CONFIG.contactEmail}</a>
            </div>
          </div>
        </div>
        <div style={{borderTop:"1px solid #1e293b",paddingTop:"24px",display:"flex",
          justifyContent:"space-between",alignItems:"center",fontSize:"12px",flexWrap:"wrap",gap:"12px"}}>
          <div>© {new Date().getFullYear()} PackWise. Built for Indian exporters and logistics teams.</div>
          <div style={{color:"#64748b"}}>Results are estimates. Verify before dispatch.</div>
        </div>
      </div>
    </footer>
  );
}

// ── Home page ──
function HomePage({setPage,onUpgrade}){
  const stats=[
    {v:"75%",l:"of logistics firms still\nuse manual spreadsheets"},
    {v:"15–25%",l:"space wasted with\nmanual planning"},
    {v:"5–15%",l:"freight cost saved\nwith load optimization"},
    {v:"< 5%",l:"of Indian supply chains\nare digitized"},
  ];
  const features=[
    {icon:"📦",title:"Box Packing Calculator",desc:"Enter box and container dimensions. Get the maximum quantity in the best orientation, with a rotatable 3D model and full 2D engineering views — top, side, and isometric.",free:true},
    {icon:"⚖️",title:"Two-SKU Pallet/Container",desc:"Pack two different product sizes together in one container while respecting a quantity ratio. Supports fragile, this-side-up, and stacking-limit constraints.",free:false},
    {icon:"🚚",title:"Shipment Planner",desc:"Enter your full order quantity. Get the number of containers needed, a per-container loading manifest, cost-per-unit comparison across vehicle types, and a branded PDF loading plan.",free:false},
    {icon:"🗃️",title:"Container SKU Calculator",desc:"Upload an Excel file with hundreds of SKUs. Get maximum quantity per SKU in seconds — constrained by both volume and weight. Download results as Excel.",free:"limited"},
  ];
  const steps=[
    {n:"01",title:"Select your container",desc:"Choose from Indian vehicles (Tata Ace, 19ft, 32ft SXL, 40ft ISO, and more) or enter custom dimensions. Or select a pallet base size."},
    {n:"02",title:"Enter your box dimensions",desc:"Length, width, height. Add optional constraints — weight per box, fragile (no stacking), this-side-up, or maximum stack height."},
    {n:"03",title:"Calculate & visualise",desc:"See the maximum quantity, space utilization, and a live 3D model you can rotate and inspect. Switch to 2D engineering views for printing."},
    {n:"04",title:"Plan & export",desc:"For full orders: get a per-container manifest and compare freight costs across vehicle types to find the cheapest option. Export a branded PDF."},
  ];
  const testimonials=[
    {q:"We were loading 32ft SXL trucks manually — always guessing. PackWise showed us we were losing 18% space every load.",name:"Rajesh M.",role:"Logistics Manager, Rajkot"},
    {q:"As an FBA seller, calculating how many units fit per shipment used to take me an hour in Excel. Now it takes 30 seconds.",name:"Priya S.",role:"Amazon Seller, Tirupur"},
    {q:"The PDF loading plan is the best part. My warehouse team can follow it without calling me.",name:"Anil K.",role:"Export Manager, Morbi"},
  ];
  return(
    <div>
      {/* Hero */}
      <div style={{background:"linear-gradient(160deg,#0f172a 0%,#0d2b1a 60%,#064e3b 100%)",
        padding:"90px 32px 100px",textAlign:"center",position:"relative",overflow:"hidden"}}>
        {/* Background grid decoration */}
        <div style={{position:"absolute",inset:0,backgroundImage:"radial-gradient(circle at 1px 1px, rgba(52,211,153,0.12) 1px, transparent 0)",backgroundSize:"32px 32px",pointerEvents:"none"}}/>
        <div style={{maxWidth:"820px",margin:"0 auto",position:"relative"}}>
          <div style={{display:"inline-flex",alignItems:"center",gap:"8px",background:"rgba(52,211,153,0.12)",
            border:"1px solid rgba(52,211,153,0.3)",borderRadius:"99px",padding:"6px 16px",
            marginBottom:"24px",fontSize:"13px",color:"#34d399",fontWeight:"600"}}>
            🇮🇳 Built for Indian exporters & logistics teams
          </div>
          <h1 style={{fontSize:"clamp(36px,6vw,64px)",fontWeight:"900",color:"#fff",
            lineHeight:"1.1",margin:"0 0 20px",letterSpacing:"-0.02em"}}>
            Stop guessing.<br/>
            <span className="shimmer-theme">
              Load smarter.
            </span>
          </h1>
          <p style={{fontSize:"clamp(16px,2.5vw,20px)",color:"#94a3b8",lineHeight:"1.7",
            margin:"0 0 36px",maxWidth:"600px",marginLeft:"auto",marginRight:"auto"}}>
            PackWise calculates the maximum boxes in any container or vehicle,
            plans full shipments across multiple trucks, and exports loading plans
            your warehouse team can follow — in seconds, not hours.
          </p>
          <div style={{display:"flex",gap:"14px",justifyContent:"center",flexWrap:"wrap"}}>
            <button onClick={()=>setPage("tool")} style={{padding:"15px 32px",
              background:"linear-gradient(135deg,#059669,#047857)",color:"#fff",border:"none",
              borderRadius:"10px",fontSize:"16px",fontWeight:"700",cursor:"pointer",
              boxShadow:"0 4px 24px rgba(5,150,105,0.4)"}}>
              Try Free — No Sign-up →
            </button>
            <button onClick={()=>setPage("pricing")} style={{padding:"15px 32px",
              background:"rgba(255,255,255,0.08)",color:"#fff",
              border:"1px solid rgba(255,255,255,0.2)",borderRadius:"10px",
              fontSize:"16px",fontWeight:"600",cursor:"pointer"}}>
              See Pricing
            </button>
          </div>
          <div style={{marginTop:"20px",fontSize:"13px",color:"#64748b"}}>
            Free forever for single-container packing · Pro plan from {CONFIG.priceLabel}
          </div>
        </div>
      </div>

      {/* Stats bar */}
      <div style={{background:"#f8fafc",borderBottom:"1px solid #e2e8f0"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto",padding:"0 24px"}} className="rg-4">
          {[
            {v:75,suf:"%",l:"of logistics firms still\nuse manual spreadsheets"},
            {v:15,suf:"–25%",l:"space wasted with\nmanual planning"},
            {v:5,suf:"–15%",l:"freight cost saved\nwith load optimization"},
            {v:5,suf:"%",pre:"< ",l:"of Indian supply chains\nare digitized"},
          ].map((s,i)=>(
            <FadeIn key={i} style={{padding:"28px 20px",textAlign:"center"}}
              className={`stat-right-border${i<3?" ":""}`}>
              <div style={{fontSize:"32px",fontWeight:"900",color:"#059669",lineHeight:1}}>
                {s.pre&&s.pre}<CountUp value={s.v} suffix={s.suf}/>
              </div>
              <div style={{fontSize:"12px",color:"#64748b",marginTop:"6px",lineHeight:"1.5",whiteSpace:"pre-line"}}>{s.l}</div>
            </FadeIn>
          ))}
        </div>
      </div>

      {/* Problem statement */}
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"80px 24px 0"}}>
        <FadeIn className="rg-2c">
          <div>
            <div style={{fontSize:"12px",fontWeight:"700",color:"#059669",letterSpacing:"0.1em",
              textTransform:"uppercase",marginBottom:"12px"}}>The problem</div>
            <h2 style={{fontSize:"36px",fontWeight:"800",color:"#0f172a",lineHeight:"1.2",margin:"0 0 20px"}}>
              75% of logistics teams still plan loads in Excel
            </h2>
            <p style={{fontSize:"16px",color:"#475569",lineHeight:"1.8",marginBottom:"16px"}}>
              Manual planning means guessing how many boxes fit, stacking wrong orientations,
              and discovering the gap when the truck shows up. The result: wasted space,
              extra vehicles, and higher freight cost on every shipment.
            </p>
            <p style={{fontSize:"16px",color:"#475569",lineHeight:"1.8"}}>
              PackWise replaces the spreadsheet with a precise, instant calculation —
              accounting for orientation, weight limits, fragile handling, and multiple vehicle
              types — and shows you exactly where every box goes.
            </p>
          </div>
          <div style={{background:"linear-gradient(135deg,#f0fdf4,#eff6ff)",borderRadius:"16px",padding:"32px"}}>
            {[["Before PackWise","After PackWise"],
              ["Guess how many boxes fit","Exact count in 10 seconds"],
              ["Try orientations manually","Best orientation found automatically"],
              ["Ignore weight limits","Weight + volume both respected"],
              ["One container estimate","Full multi-truck manifest"],
              ["No loading instructions","PDF plan for warehouse team"],
              ["Hours in Excel","30 seconds in a browser"],
            ].map(([a,b],i)=>(
              <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",
                padding:i===0?"0 0 12px":"12px 0",
                borderBottom:i===0?"2px solid #e2e8f0":"1px solid #f1f5f9"}}>
                <div style={{fontSize:i===0?"11px":"13px",fontWeight:i===0?"700":"400",
                  color:i===0?"#64748b":"#ef4444",
                  textDecoration:i===0?"none":"line-through",textTransform:i===0?"uppercase":"none",
                  letterSpacing:i===0?"0.08em":"normal"}}>{a}</div>
                <div style={{fontSize:i===0?"11px":"13px",fontWeight:i===0?"700":"600",
                  color:i===0?"#64748b":"#059669",textTransform:i===0?"uppercase":"none",
                  letterSpacing:i===0?"0.08em":"normal"}}>{b}</div>
              </div>
            ))}
          </div>
        </FadeIn>
      </div>

      {/* Features */}
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"80px 24px 0"}}>
        <FadeIn style={{textAlign:"center",marginBottom:"48px"}}>
          <div style={{fontSize:"12px",fontWeight:"700",color:"#059669",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:"10px"}}>What's inside</div>
          <h2 style={{fontSize:"36px",fontWeight:"800",color:"#0f172a",margin:"0 0 12px"}}>Four tools. One platform.</h2>
          <p style={{fontSize:"16px",color:"#64748b",maxWidth:"480px",margin:"0 auto"}}>Start free with single-container packing. Upgrade when you're ready to plan full shipments.</p>
        </FadeIn>
        <FadeIn className="rg-2e stagger">
          {features.map((f,i)=>(
            <div key={i} className="hover-lift" style={{background:"#fff",borderRadius:"14px",padding:"28px",
              border:`1px solid ${f.free===true?"#bbf7d0":f.free===false?"#e2e8f0":"#fde68a"}`,
              position:"relative",boxShadow:"0 1px 4px rgba(0,0,0,0.05)"}}>
              <div style={{position:"absolute",top:"20px",right:"20px"}}>
                {f.free===true&&<span style={{background:"#dcfce7",color:"#166534",fontSize:"11px",fontWeight:"700",padding:"3px 10px",borderRadius:"99px"}}>FREE</span>}
                {f.free===false&&<span style={{background:"#fef3c7",color:"#92400e",fontSize:"11px",fontWeight:"700",padding:"3px 10px",borderRadius:"99px"}}>⭐ PRO</span>}
                {f.free==="limited"&&<span style={{background:"#fff7ed",color:"#c2410c",fontSize:"11px",fontWeight:"700",padding:"3px 10px",borderRadius:"99px"}}>FREE (10 SKUs)</span>}
              </div>
              <div style={{fontSize:"32px",marginBottom:"12px"}}>{f.icon}</div>
              <h3 style={{fontSize:"18px",fontWeight:"700",color:"#0f172a",margin:"0 0 8px"}}>{f.title}</h3>
              <p style={{fontSize:"14px",color:"#64748b",lineHeight:"1.7",margin:0}}>{f.desc}</p>
              <button onClick={()=>setPage("tool")} className="btn-primary" style={{marginTop:"16px",padding:"8px 16px",
                background:f.free===false?"linear-gradient(135deg,#059669,#047857)":"#f8fafc",
                color:f.free===false?"#fff":"#059669",border:`1px solid ${f.free===false?"transparent":"#059669"}`,
                borderRadius:"8px",fontSize:"13px",fontWeight:"600",cursor:"pointer"}}>
                {f.free===false?"Try with Pro →":"Try Free →"}
              </button>
            </div>
          ))}
        </FadeIn>
      </div>

      {/* How it works */}
      <div style={{background:"#0f172a",margin:"80px 0 0",padding:"80px 24px"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto"}}>
          <FadeIn style={{textAlign:"center",marginBottom:"56px"}}>
            <div style={{fontSize:"12px",fontWeight:"700",color:"#34d399",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:"10px"}}>How it works</div>
            <h2 style={{fontSize:"36px",fontWeight:"800",color:"#fff",margin:0}}>From dimensions to loading plan in 4 steps</h2>
          </FadeIn>
          <FadeIn className="rg-steps stagger">
            {steps.map((s,i)=>(
              <div key={i} style={{borderTop:"3px solid #059669",paddingTop:"24px"}}>
                <div style={{fontSize:"42px",fontWeight:"900",color:"rgba(52,211,153,0.2)",lineHeight:1,marginBottom:"12px"}}>{s.n}</div>
                <h3 style={{fontSize:"16px",fontWeight:"700",color:"#fff",margin:"0 0 10px"}}>{s.title}</h3>
                <p style={{fontSize:"13px",color:"#94a3b8",lineHeight:"1.7",margin:0}}>{s.desc}</p>
              </div>
            ))}
          </FadeIn>
          <div style={{textAlign:"center",marginTop:"48px"}}>
            <button onClick={()=>setPage("tool")} className="btn-primary" style={{padding:"14px 32px",
              background:"linear-gradient(135deg,#059669,#047857)",color:"#fff",border:"none",
              borderRadius:"10px",fontSize:"15px",fontWeight:"700",cursor:"pointer",
              boxShadow:"0 4px 20px rgba(5,150,105,0.4)"}}>
              Open the Calculator →
            </button>
          </div>
        </div>
      </div>

      {/* Who it's for */}
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"80px 24px 0"}}>
        <FadeIn style={{textAlign:"center",marginBottom:"48px"}}>
          <div style={{fontSize:"12px",fontWeight:"700",color:"#059669",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:"10px"}}>Who it's for</div>
          <h2 style={{fontSize:"36px",fontWeight:"800",color:"#0f172a",margin:0}}>Built for the people who actually load trucks</h2>
        </FadeIn>
        <FadeIn className="rg-3 stagger">
          {[
            {icon:"🏭",title:"MSME Exporters",desc:"Plan container loads before your shipment leaves. Know exactly how many 20ft or 40ft containers you need and the cost per unit — before booking."},
            {icon:"📦",title:"Amazon FBA Sellers",desc:"Calculate exactly how many units fit per shipment box or pallet. Stop paying for air inside containers you're already loading at a premium."},
            {icon:"🚛",title:"3PL & Freight Teams",desc:"Compare loading efficiency across all your vehicle types side by side. Find which container gives you the lowest cost per unit for each client's order."},
            {icon:"🏪",title:"Warehouse Managers",desc:"Get a step-by-step loading plan with exact orientations and layer counts that your team can follow without guessing — printed from a PDF."},
            {icon:"📊",title:"Supply Chain Planners",desc:"Upload your SKU list and instantly know which products waste the most container space. Identify packing inefficiencies across your entire catalog."},
            {icon:"🌏",title:"Import/Export Agents",desc:"Your clients ask how many units fit and what it costs. Answer in 30 seconds with a professional PDF report — without calling your warehouse."},
          ].map((c,i)=>(
            <div key={i} className="hover-lift hover-border" style={{background:"#f8fafc",borderRadius:"12px",padding:"24px",border:"1px solid #e2e8f0",transition:"border-color 0.2s"}}>
              <div style={{fontSize:"28px",marginBottom:"10px"}}>{c.icon}</div>
              <h3 style={{fontSize:"16px",fontWeight:"700",color:"#0f172a",margin:"0 0 8px"}}>{c.title}</h3>
              <p style={{fontSize:"13px",color:"#64748b",lineHeight:"1.7",margin:0}}>{c.desc}</p>
            </div>
          ))}
        </FadeIn>
      </div>

      {/* Testimonials */}
      <div style={{background:"#f0fdf4",margin:"80px 0 0",padding:"72px 24px"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto"}}>
          <FadeIn style={{textAlign:"center",marginBottom:"48px"}}>
            <h2 style={{fontSize:"32px",fontWeight:"800",color:"#0f172a",margin:"0 0 8px"}}>What users say</h2>
            <p style={{color:"#64748b",fontSize:"15px",margin:0}}>Early users from India's manufacturing and export clusters</p>
          </FadeIn>
          <FadeIn className="rg-3 stagger">
            {testimonials.map((t,i)=>(
              <div key={i} className="hover-lift" style={{background:"#fff",borderRadius:"14px",padding:"28px",
                boxShadow:"0 2px 12px rgba(0,0,0,0.06)",border:"1px solid #dcfce7"}}>
                <div style={{fontSize:"32px",color:"#059669",lineHeight:1,marginBottom:"14px"}}>"</div>
                <p style={{fontSize:"14px",color:"#374151",lineHeight:"1.8",margin:"0 0 20px",fontStyle:"italic"}}>{t.q}</p>
                <div style={{fontWeight:"700",fontSize:"14px",color:"#0f172a"}}>{t.name}</div>
                <div style={{fontSize:"12px",color:"#64748b"}}>{t.role}</div>
              </div>
            ))}
          </FadeIn>
        </div>
      </div>

      {/* CTA */}
      <div style={{maxWidth:"1200px",margin:"0 auto",padding:"80px 32px 0"}}>
        <div style={{background:"linear-gradient(135deg,#0f172a,#064e3b)",borderRadius:"20px",
          padding:"64px 48px",textAlign:"center",position:"relative",overflow:"hidden"}}>
          <div style={{position:"absolute",inset:0,backgroundImage:"radial-gradient(circle at 1px 1px, rgba(52,211,153,0.08) 1px, transparent 0)",backgroundSize:"28px 28px"}}/>
          <div style={{position:"relative"}}>
            <h2 style={{fontSize:"36px",fontWeight:"900",color:"#fff",margin:"0 0 16px"}}>Start packing smarter today</h2>
            <p style={{fontSize:"16px",color:"#94a3b8",margin:"0 0 32px"}}>Free forever for single-container packing. No sign-up, no software to install.</p>
            <div style={{display:"flex",gap:"14px",justifyContent:"center",flexWrap:"wrap"}}>
              <button onClick={()=>setPage("tool")} style={{padding:"14px 32px",
                background:"linear-gradient(135deg,#059669,#047857)",color:"#fff",border:"none",
                borderRadius:"10px",fontSize:"16px",fontWeight:"700",cursor:"pointer",
                boxShadow:"0 4px 20px rgba(5,150,105,0.4)"}}>
                Open Free Calculator →
              </button>
              <button onClick={()=>setPage("pricing")} style={{padding:"14px 32px",
                background:"rgba(255,255,255,0.08)",color:"#fff",
                border:"1px solid rgba(255,255,255,0.2)",borderRadius:"10px",
                fontSize:"16px",fontWeight:"600",cursor:"pointer"}}>
                See Pro features
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Pricing page ──
function PricingPage({onUpgrade,setPage}){
  const freeFeat=["Box packing calculator","3D model + 2D engineering views","Weight & stacking constraints","All vehicle & pallet presets","Container SKU (up to 10 SKUs)"];
  const proFeat=["Everything in Free","⚖️ Two-SKU Pallet/Container","🚚 Shipment Planner — multi-container","💰 Cost comparison across vehicles","📄 Branded PDF loading plan export","Unlimited Container SKU upload","Priority email support"];
  return(
    <div style={{maxWidth:"1000px",margin:"0 auto",padding:"72px 32px 0"}}>
      <div style={{textAlign:"center",marginBottom:"56px"}}>
        <div style={{fontSize:"12px",fontWeight:"700",color:"#059669",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:"10px"}}>Simple pricing</div>
        <h1 style={{fontSize:"42px",fontWeight:"900",color:"#0f172a",margin:"0 0 14px"}}>Pay for what saves you money</h1>
        <p style={{fontSize:"18px",color:"#64748b",maxWidth:"500px",margin:"0 auto"}}>Start free. Upgrade when your orders grow beyond a single container.</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"24px",maxWidth:"800px",margin:"0 auto"}}>
        {/* Free */}
        <div style={{background:"#fff",borderRadius:"16px",padding:"32px",border:"1px solid #e2e8f0",boxShadow:"0 2px 12px rgba(0,0,0,0.04)"}}>
          <div style={{fontSize:"14px",fontWeight:"700",color:"#64748b",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"8px"}}>Free</div>
          <div style={{fontSize:"42px",fontWeight:"900",color:"#0f172a",lineHeight:1}}>₹0</div>
          <div style={{fontSize:"13px",color:"#94a3b8",marginBottom:"24px"}}>forever</div>
          <button onClick={()=>setPage("tool")} style={{width:"100%",padding:"12px",background:"#f1f5f9",color:"#374151",border:"none",borderRadius:"10px",fontSize:"14px",fontWeight:"600",cursor:"pointer",marginBottom:"24px"}}>
            Start Free →
          </button>
          {freeFeat.map(f=>(
            <div key={f} style={{display:"flex",gap:"10px",alignItems:"flex-start",padding:"7px 0",borderBottom:"1px solid #f8fafc"}}>
              <span style={{color:"#059669",fontWeight:"700",flexShrink:0}}>✓</span>
              <span style={{fontSize:"14px",color:"#374151"}}>{f}</span>
            </div>
          ))}
        </div>
        {/* Pro */}
        <div className="pro-glow" style={{background:"linear-gradient(160deg,#0f172a,#0d2b1a)",borderRadius:"16px",padding:"32px",border:"1px solid #059669",boxShadow:"0 8px 32px rgba(5,150,105,0.25)",position:"relative",overflow:"hidden"}}>
          <div style={{position:"absolute",top:"16px",right:"16px",background:"#fbbf24",color:"#78350f",fontSize:"11px",fontWeight:"800",padding:"4px 10px",borderRadius:"99px"}}>MOST POPULAR</div>
          <div style={{fontSize:"14px",fontWeight:"700",color:"#34d399",textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:"8px"}}>Pro</div>
          <div style={{fontSize:"42px",fontWeight:"900",color:"#fff",lineHeight:1}}>{CONFIG.priceLabel.split(" / ")[0]}</div>
          <div style={{fontSize:"13px",color:"#94a3b8",marginBottom:"24px"}}>per month · cancel anytime</div>
          <button onClick={onUpgrade} style={{width:"100%",padding:"12px",
            background:"linear-gradient(135deg,#059669,#047857)",color:"#fff",border:"none",
            borderRadius:"10px",fontSize:"14px",fontWeight:"700",cursor:"pointer",marginBottom:"24px",
            boxShadow:"0 4px 16px rgba(5,150,105,0.4)"}}>
            {CONFIG.paymentLink?"Get Pro Now →":"Contact Us to Upgrade →"}
          </button>
          {proFeat.map(f=>(
            <div key={f} style={{display:"flex",gap:"10px",alignItems:"flex-start",padding:"7px 0",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
              <span style={{color:"#34d399",fontWeight:"700",flexShrink:0}}>✓</span>
              <span style={{fontSize:"14px",color:"#cbd5e1"}}>{f}</span>
            </div>
          ))}
        </div>
      </div>
      {/* FAQ */}
      <div style={{maxWidth:"700px",margin:"64px auto 0"}}>
        <h2 style={{fontSize:"28px",fontWeight:"800",color:"#0f172a",textAlign:"center",marginBottom:"32px"}}>Common questions</h2>
        {[
          ["Do I need to install anything?","No. PackWise runs entirely in your browser. No downloads, no setup, no IT department."],
          ["Can I try Pro before paying?","Yes — click 'Get Pro' and contact us for a 7-day free trial."],
          ["What units does it use?","Millimetres by default. As long as you're consistent (all mm or all cm), the result is accurate."],
          ["Is my data safe?","Everything runs in your browser — we never see your box sizes or shipping data. Nothing is sent to our servers."],
          ["Can I export the results?","The Shipment Planner (Pro) exports a branded PDF loading plan. Container SKU exports an Excel file."],
          ["Do you have India-specific vehicles?","Yes — Tata Ace, 19ft, 20ft, 22ft, 32ft SXL/MXL, and 40ft ISO containers are all preset."],
        ].map(([q,a],i)=>(
          <div key={i} style={{borderBottom:"1px solid #e2e8f0",padding:"20px 0"}}>
            <div style={{fontWeight:"700",color:"#0f172a",fontSize:"15px",marginBottom:"8px"}}>{q}</div>
            <div style={{fontSize:"14px",color:"#64748b",lineHeight:"1.7"}}>{a}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── About page ──
function AboutPage({setPage}){
  return(
    <div style={{maxWidth:"900px",margin:"0 auto",padding:"72px 32px 0"}}>
      {/* Hero */}
      <div style={{marginBottom:"64px"}}>
        <div style={{fontSize:"12px",fontWeight:"700",color:"#059669",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:"12px"}}>Our story</div>
        <h1 style={{fontSize:"42px",fontWeight:"900",color:"#0f172a",lineHeight:"1.1",margin:"0 0 24px"}}>
          Built because loading a truck shouldn't need a logistics degree
        </h1>
        <p style={{fontSize:"18px",color:"#475569",lineHeight:"1.8",marginBottom:"16px"}}>
          PackWise was built out of a simple frustration: India has millions of exporters,
          warehouse managers, and freight teams — and most of them are still figuring out
          container loading with a tape measure and a spreadsheet.
        </p>
        <p style={{fontSize:"18px",color:"#475569",lineHeight:"1.8"}}>
          The software that existed was expensive ($100–500/month), designed for
          Western logistics operations, and didn't know what a 32ft SXL was.
          So we built something that does — in a browser, in under 30 seconds,
          for a fraction of the cost.
        </p>
      </div>

      {/* Mission */}
      <div style={{background:"linear-gradient(135deg,#f0fdf4,#eff6ff)",borderRadius:"16px",padding:"40px",marginBottom:"64px"}}>
        <h2 style={{fontSize:"24px",fontWeight:"800",color:"#0f172a",margin:"0 0 16px"}}>Our mission</h2>
        <p style={{fontSize:"16px",color:"#374151",lineHeight:"1.8",margin:0}}>
          To give every Indian exporter, small manufacturer, and logistics team
          the same load-planning intelligence that large companies pay enterprise
          software prices for — accessible from any browser, priced for Indian MSMEs,
          and built around how Indian logistics actually works.
        </p>
      </div>

      {/* What we solve */}
      <div style={{marginBottom:"64px"}}>
        <h2 style={{fontSize:"32px",fontWeight:"800",color:"#0f172a",margin:"0 0 32px"}}>The problems we solve</h2>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"24px"}}>
          {[
            {icon:"📐",title:"Space waste",desc:"Manual planning wastes 15–25% of container space on average. PackWise finds the optimal box orientation across all 6 rotations and fills leftover gaps automatically."},
            {icon:"💸",title:"Unnecessary freight cost",desc:"Booking an extra container because you can't calculate exactly how many fit is expensive. PackWise tells you before the truck arrives."},
            {icon:"⚖️",title:"Ignored constraints",desc:"Fragile items, weight limits, this-side-up labels — these get ignored in manual planning. PackWise enforces them in the calculation."},
            {icon:"📋",title:"No loading instructions",desc:"Even with good planning, warehouse staff load by eye. A printed PDF loading plan cuts errors and loading time significantly."},
            {icon:"🚛",title:"Wrong vehicle choice",desc:"Is a 32ft SXL or two 22ft trucks cheaper for your order? The cost comparison tool answers this immediately."},
            {icon:"📊",title:"SKU-level blindness",desc:"Businesses with hundreds of SKUs don't know which products are most space-efficient. The Container SKU tool reveals this across your whole catalog."},
          ].map((c,i)=>(
            <div key={i} style={{display:"flex",gap:"14px",padding:"20px",background:"#f8fafc",borderRadius:"12px",border:"1px solid #e2e8f0"}}>
              <div style={{fontSize:"28px",flexShrink:0}}>{c.icon}</div>
              <div>
                <div style={{fontWeight:"700",color:"#0f172a",marginBottom:"6px"}}>{c.title}</div>
                <div style={{fontSize:"13px",color:"#64748b",lineHeight:"1.7"}}>{c.desc}</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* India focus */}
      <div style={{background:"#0f172a",borderRadius:"16px",padding:"40px",marginBottom:"64px"}}>
        <h2 style={{fontSize:"24px",fontWeight:"800",color:"#fff",margin:"0 0 16px"}}>🇮🇳 Built for India</h2>
        <p style={{fontSize:"15px",color:"#94a3b8",lineHeight:"1.8",marginBottom:"20px"}}>
          India's logistics cost runs at 14–18% of GDP — almost double the global benchmark.
          Part of that gap is inefficient loading. PackWise is designed specifically for the
          Indian market, with:
        </p>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px"}}>
          {["Indian vehicle presets — Tata Ace, 19ft, 22ft, 32ft SXL/MXL, 40ft ISO",
            "Pricing in Indian Rupees (₹999/month, not $99/month)",
            "Free tier designed for one-off shippers with small orders",
            "No software to install — works on any smartphone or laptop",
            "Targeting exporter clusters: Morbi, Tirupur, Rajkot, Ludhiana",
            "Built to replace Excel, not replace your logistics team"].map((t,i)=>(
            <div key={i} style={{display:"flex",gap:"10px",fontSize:"13px",color:"#cbd5e1"}}>
              <span style={{color:"#34d399",flexShrink:0}}>✓</span>{t}</div>
          ))}
        </div>
      </div>

      {/* Algorithm note */}
      <div style={{marginBottom:"64px"}}>
        <h2 style={{fontSize:"28px",fontWeight:"800",color:"#0f172a",margin:"0 0 16px"}}>How the algorithm works</h2>
        <p style={{fontSize:"15px",color:"#475569",lineHeight:"1.8",marginBottom:"12px"}}>
          PackWise uses a <strong>mixed-orientation guillotine heuristic</strong> — a practical
          algorithm proven in logistics software. It tries all 6 possible box orientations
          for the main grid, then fills the three leftover regions (side gap, front gap, top gap)
          with the best-fit orientation for each.
        </p>
        <p style={{fontSize:"15px",color:"#475569",lineHeight:"1.8",marginBottom:"12px"}}>
          For the Two-SKU problem, it searches 200+ guillotine cuts across three axes,
          scoring each by total boxes packed balanced against ratio closeness. For the
          Shipment Planner, it uses a First-Fit-Decreasing bin-packing strategy across
          multiple containers.
        </p>
        <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:"10px",padding:"16px 20px",fontSize:"13px",color:"#92400e",lineHeight:"1.7"}}>
          <strong>Honest note:</strong> Container loading is an NP-hard problem — no algorithm
          finds the perfect answer every time. PackWise gives a very good practical solution
          that consistently beats manual planning. Treat the output as a strong plan, not an
          absolute guarantee. Verify dimensions before dispatch.
        </div>
      </div>

      {/* Contact */}
      <div style={{background:"#f8fafc",borderRadius:"16px",padding:"40px",textAlign:"center",marginBottom:"40px"}}>
        <h2 style={{fontSize:"24px",fontWeight:"800",color:"#0f172a",margin:"0 0 12px"}}>Get in touch</h2>
        <p style={{fontSize:"15px",color:"#64748b",margin:"0 0 20px"}}>Questions, feedback, or want a demo for your team?</p>
        <a href={`mailto:${CONFIG.contactEmail}`} style={{display:"inline-block",padding:"12px 28px",
          background:"linear-gradient(135deg,#059669,#047857)",color:"#fff",borderRadius:"10px",
          fontWeight:"700",fontSize:"15px",textDecoration:"none"}}>
          Email Us →
        </a>
        <div style={{marginTop:"20px"}}>
          <button onClick={()=>setPage("pricing")} style={{background:"none",border:"none",color:"#059669",fontWeight:"600",fontSize:"14px",cursor:"pointer"}}>
            See pricing →
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Tool page wrapper ──
function ToolPage({isPro,setIsPro,modalOpen,setModalOpen}){
  const[tab,setTab]=useState("box");
  const unlock=()=>{setIsPro(true);try{localStorage.setItem("pp_pro","true");}catch(e){}setTimeout(()=>setModalOpen(false),1200);};
  const tabs=[["box","📦 Box Packing",false],["multisku","🗃️ Multi-SKU Planner",true],["twosku","⚖️ Two-SKU Pallet/Container",true],["shipment","🚚 Shipment Planner",true],["sku","🗃️ Container SKU",false]];
  return(
    <div>
      <UpgradeModal open={modalOpen} onClose={()=>setModalOpen(false)} onUnlock={unlock}/>
      {/* Tool sub-nav */}
      <div style={{background:"linear-gradient(135deg,#0f172a 0%,#0d2b1a 100%)",padding:"20px 32px 0"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto"}}>
          <p style={{color:"#94a3b8",fontSize:"12px",margin:"0 0 14px"}}>
            Box packing · Two-SKU pallet · Shipment planner · Container SKU bulk calculator
          </p>
          <div style={{display:"flex",gap:"4px",flexWrap:"wrap"}}>
            {tabs.map(([id,label,pro])=>(
              <button key={id} onClick={()=>setTab(id)} style={{padding:"9px 18px",border:"none",cursor:"pointer",
                fontSize:"13px",fontWeight:"600",borderRadius:"8px 8px 0 0",
                background:tab===id?"#f0f4f8":"rgba(255,255,255,0.08)",
                color:tab===id?"#1a2332":"#cbd5e1",transition:"all 0.15s",position:"relative"}}>
                {label}{pro&&!isPro&&<span style={{marginLeft:"5px",fontSize:"10px",background:"#fbbf24",color:"#78350f",padding:"1px 5px",borderRadius:"99px",fontWeight:"700"}}>PRO</span>}
              </button>
            ))}
          </div>
        </div>
      </div>
      <div style={{background:"#f0fdf4",minHeight:"60vh",padding:"24px 32px 0"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto"}}>
          {tab==="box"&&<BoxPackingTool/>}
          {tab==="multisku"&&(isPro?<MultiSKUTool/>:<ProGate feature="Multi-SKU Planner" onUpgrade={()=>setModalOpen(true)}/>)}
          {tab==="twosku"&&(isPro?<TwoSKUTool/>:<ProGate feature="Two-SKU Pallet/Container Calculator" onUpgrade={()=>setModalOpen(true)}/>)}
          {tab==="shipment"&&(isPro?<ShipmentPlanner/>:<ProGate feature="Shipment Planner" onUpgrade={()=>setModalOpen(true)}/>)}
          {tab==="sku"&&<ContainerSkuTool isPro={isPro} onUpgrade={()=>setModalOpen(true)}/>}
        </div>
      </div>
    </div>
  );
}

// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App(){
  const[page,setPage]=useState("home");
  const[isPro,setIsPro]=useState(false);
  const[modalOpen,setModalOpen]=useState(false);
  useEffect(()=>{try{if(localStorage.getItem("pp_pro")==="true") setIsPro(true);}catch(e){}
    window.scrollTo(0,0);},[page]);
  const logout=()=>{setIsPro(false);try{localStorage.removeItem("pp_pro");}catch(e){}};
  const openUpgrade=()=>{setPage("tool");setTimeout(()=>setModalOpen(true),100);};
  return(
    <div style={{fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif",background:"#f0fdf4",minHeight:"100vh"}}>
      <Nav page={page} setPage={setPage} isPro={isPro} onUpgrade={openUpgrade} onLogout={logout}/>
      <main>
        <div className="page-enter" key={page}>
          {page==="home"&&<HomePage setPage={setPage} onUpgrade={openUpgrade}/>}
          {page==="tool"&&<ToolPage isPro={isPro} setIsPro={setIsPro} modalOpen={modalOpen} setModalOpen={setModalOpen}/>}
          {page==="pricing"&&<PricingPage onUpgrade={openUpgrade} setPage={setPage}/>}
          {page==="about"&&<AboutPage setPage={setPage}/>}
        </div>
      </main>
      <Footer setPage={setPage}/>
    </div>
  );
}
