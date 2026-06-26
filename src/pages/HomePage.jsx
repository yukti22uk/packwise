// ─── HOME PAGE ───────────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import { CONFIG } from '../config.js';
import FadeIn from '../components/FadeIn.jsx';
import CountUp from '../components/CountUp.jsx';
import PackWiseLogo from '../components/PackWiseLogo.jsx';
import BoxLoadingAnim, { ANIM_TOTAL } from '../components/BoxLoadingAnim.jsx';
import HomeDemoViewer from '../components/HomeDemoViewer.jsx';
function HomePage({setPage,onUpgrade}){
  const stats=[
    {v:"75%",l:"of logistics firms still\nuse manual spreadsheets"},
    {v:"15–25%",l:"space wasted with\nmanual planning"},
    {v:"5–15%",l:"freight cost saved\nwith load optimization"},
    {v:"< 5%",l:"of Indian supply chains\nare digitized"},
  ];
  const features=[
    {icon:"📦",title:"Single SKU Calculator",desc:"Enter one box size and a container. Instantly see the maximum quantity in the best orientation, with a rotatable 3D model and full 2D engineering views — top, side, and isometric.",free:true},
    {icon:"🗃️",title:"Multi-SKU Planner",desc:"Pack 2–8 different box sizes together in one container simultaneously. Each SKU gets an allocated region. Respects weight limits, fragile, this-side-up, and stacking constraints.",free:false},
    {icon:"🚚",title:"Shipment Planner",desc:"Enter your total order quantity. Get containers needed, per-container manifest, cost-per-unit comparison across all vehicle types, and a branded PDF loading plan for your warehouse team.",free:false},
    {icon:"🗃️",title:"Bulk SKU Calculator",desc:"Upload an Excel file with hundreds of SKUs. Get maximum quantity per SKU in seconds — constrained by both volume and weight. Download the full results as Excel.",free:"limited"},
    {icon:"🔀",title:"SKU Grouper",desc:"Upload 10,000+ SKUs and automatically cluster them into 2–8 representative box size groups using K-means. One click sends the groups to the Multi-SKU Planner for full container planning.",free:true},
    {icon:"📊",title:"Order Analyser",desc:"Upload Master SKU data and Order data in any format. AI maps columns automatically, flags anomalies, and produces a 6-sheet Excel report — ABC analysis, FMS classification, and ABC-FMS matrix.",free:false},
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
      {/* Hero — split layout */}
      <div style={{background:"linear-gradient(160deg,#0f172a 0%,#1a0a14 60%,#2d0b1a 100%)",
        padding:"80px 24px 90px",position:"relative",overflow:"hidden"}}>
        <div style={{position:"absolute",inset:0,backgroundImage:"radial-gradient(circle at 1px 1px, rgba(190,24,93,0.1) 1px, transparent 0)",backgroundSize:"32px 32px",pointerEvents:"none"}}/>
        <div style={{maxWidth:"1200px",margin:"0 auto",position:"relative"}}>
          <div className="rg-2c" style={{gap:"48px",alignItems:"center"}}>
            {/* Left — text */}
            <div>
              <div style={{display:"inline-flex",alignItems:"center",gap:"8px",
                background:"rgba(190,24,93,0.12)",border:"1px solid rgba(190,24,93,0.35)",
                borderRadius:"99px",padding:"6px 16px",marginBottom:"24px",
                fontSize:"13px",color:"#f9a8d4",fontWeight:"700"}}>
                🇮🇳 Built for Indian exporters & logistics teams
              </div>
              <h1 style={{fontSize:"clamp(36px,5vw,60px)",fontWeight:"900",color:"#fff",
                lineHeight:"1.08",margin:"0 0 20px",letterSpacing:"-0.03em"}}>
                Stop guessing.<br/>
                <span className="shimmer-theme">Load smarter.</span>
              </h1>
              <p style={{fontSize:"clamp(15px,2vw,18px)",color:"#94a3b8",lineHeight:"1.7",
                margin:"0 0 32px",maxWidth:"460px"}}>
                PackWise calculates the maximum boxes in any container or vehicle,
                plans full shipments, and exports loading plans your warehouse team
                can follow — in seconds, not hours.
              </p>
              <div className="hero-btns" style={{display:"flex",gap:"12px",flexWrap:"wrap"}}>
                <button onClick={()=>setPage("tool")} className="btn-primary"
                  style={{padding:"14px 28px",background:"linear-gradient(135deg,#be185d,#9d174d)",
                  color:"#fff",border:"none",borderRadius:"10px",fontSize:"15px",fontWeight:"700",
                  cursor:"pointer",boxShadow:"0 4px 24px rgba(190,24,93,0.4)",fontFamily:"inherit"}}>
                  Try Free — No Sign-up →
                </button>
                <button onClick={()=>setPage("pricing")}
                  style={{padding:"14px 28px",background:"rgba(255,255,255,0.07)",color:"#fff",
                  border:"1px solid rgba(255,255,255,0.18)",borderRadius:"10px",
                  fontSize:"15px",fontWeight:"600",cursor:"pointer",fontFamily:"inherit"}}>
                  See Pricing
                </button>
              </div>
              <div style={{marginTop:"16px",fontSize:"12px",color:"#4b5563"}}>
                Free forever for single-container packing · Pro from {CONFIG.priceLabel}
              </div>
            </div>
            {/* Right — animated box loading */}
            <div>
              <div style={{marginBottom:"14px",fontSize:"11px",fontWeight:"700",
                color:"rgba(255,255,255,0.3)",textTransform:"uppercase",
                letterSpacing:"0.1em",textAlign:"center"}}>Live packing animation</div>
              <BoxLoadingAnim/>
              <div style={{textAlign:"center",marginTop:"24px",fontSize:"12px",color:"#4b5563"}}>
                32ft SXL · 7×4 grid · {ANIM_TOTAL} boxes · 100% utilized
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Interactive 3D Demo */}
      <HomeDemoViewer/>

      {/* Stats bar */}
      <div style={{background:"#fff",borderBottom:"1px solid #f1f5f9"}}>
        <div style={{maxWidth:"1200px",margin:"0 auto",padding:"0 24px"}} className="rg-4">
          {[
            {v:75,suf:"%",l:"of logistics firms still\nuse manual spreadsheets"},
            {v:15,suf:"–25%",l:"space wasted with\nmanual planning"},
            {v:5,suf:"–15%",l:"freight cost saved\nwith load optimization"},
            {v:5,suf:"%",pre:"< ",l:"of Indian supply chains\nare digitized"},
          ].map((s,i)=>(
            <FadeIn key={i} style={{padding:"28px 20px",textAlign:"center"}}
              className={`stat-sep${i<3?" ":""}`}>
              <div style={{fontSize:"32px",fontWeight:"900",color:"#be185d",lineHeight:1}}>
                {s.pre&&s.pre}<CountUp value={s.v} suffix={s.suf}/>
              </div>
              <div style={{fontSize:"12px",color:"#6b7280",marginTop:"6px",lineHeight:"1.5",whiteSpace:"pre-line"}}>{s.l}</div>
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
          <h2 style={{fontSize:"36px",fontWeight:"800",color:"#0f172a",margin:"0 0 12px"}}>Six tools. One platform.</h2>
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

export default HomePage;
