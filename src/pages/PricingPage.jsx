// ─── PRICING PAGE ────────────────────────────────────────────────────────────
import { CONFIG } from '../config.js';
function PricingPage({onUpgrade,setPage}){
  const freeFeat=["Single SKU Calculator — unlimited","3D model + 2D engineering views","Weight & stacking constraints","All vehicle & pallet presets","Bulk SKU Calculator (up to 10 SKUs)"];
  const proFeat=["Everything in Free","🗃️ Multi-SKU Planner — 2 to 8 SKUs per container","🚚 Shipment Planner — multi-container","💰 Cost comparison across vehicles","📄 Branded PDF loading plan export","Unlimited Bulk SKU upload","Priority email support"];
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
          ["Can I export the results?","The Shipment Planner (Pro) exports a branded PDF loading plan. The Bulk SKU Calculator exports an Excel file. The Multi-SKU Planner also exports Excel."],
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

export default PricingPage;
