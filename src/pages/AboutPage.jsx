// ─── ABOUT PAGE ──────────────────────────────────────────────────────────────
import { CONFIG } from '../config.js';
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
          DensiCube was built out of a simple frustration: India has millions of exporters,
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
            {icon:"📐",title:"Space waste",desc:"Manual planning wastes 15–25% of container space on average. DensiCube finds the optimal box orientation across all 6 rotations and fills leftover gaps automatically."},
            {icon:"💸",title:"Unnecessary freight cost",desc:"Booking an extra container because you can't calculate exactly how many fit is expensive. DensiCube tells you before the truck arrives."},
            {icon:"⚖️",title:"Ignored constraints",desc:"Fragile items, weight limits, this-side-up labels — these get ignored in manual planning. DensiCube enforces them in the calculation."},
            {icon:"📋",title:"No loading instructions",desc:"Even with good planning, warehouse staff load by eye. A printed PDF loading plan cuts errors and loading time significantly."},
            {icon:"🚛",title:"Wrong vehicle choice",desc:"Is a 32ft SXL or two 22ft trucks cheaper for your order? The cost comparison tool answers this immediately."},
            {icon:"📊",title:"SKU-level blindness",desc:"Businesses with hundreds of SKUs don't know which products are most space-efficient. The Bulk SKU Calculator reveals this across your whole catalog."},
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
          Part of that gap is inefficient loading. DensiCube is designed specifically for the
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
          DensiCube uses a <strong>mixed-orientation guillotine heuristic</strong> — a practical
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
          finds the perfect answer every time. DensiCube gives a very good practical solution
          that consistently beats manual planning. Treat the output as a strong plan, not an
          absolute guarantee. Verify dimensions before dispatch.
        </div>
      </div>

      {/* Contact */}
      <div style={{background:"#f8fafc",borderRadius:"16px",padding:"40px",textAlign:"center",marginBottom:"40px"}}>
        <h2 style={{fontSize:"24px",fontWeight:"800",color:"#0f172a",margin:"0 0 12px"}}>Get in touch</h2>
        <p style={{fontSize:"15px",color:"#64748b",margin:"0 0 24px"}}>Questions, feedback, or want a demo for your team?</p>
        <div style={{display:"flex",gap:"14px",justifyContent:"center",flexWrap:"wrap",marginBottom:"20px"}}>
          <a href={`mailto:${CONFIG.contactEmail}`} style={{display:"inline-flex",alignItems:"center",
            gap:"8px",padding:"12px 24px",
            background:"linear-gradient(135deg,#be185d,#9d174d)",color:"#fff",borderRadius:"10px",
            fontWeight:"700",fontSize:"14px",textDecoration:"none"}}>
            📧 {CONFIG.contactEmail}
          </a>
          <a href={`tel:${CONFIG.contactPhone}`} style={{display:"inline-flex",alignItems:"center",
            gap:"8px",padding:"12px 24px",
            background:"linear-gradient(135deg,#059669,#047857)",color:"#fff",borderRadius:"10px",
            fontWeight:"700",fontSize:"14px",textDecoration:"none"}}>
            📞 {CONFIG.contactPhone}
          </a>
        </div>
        <div style={{marginTop:"4px"}}>
          <button onClick={()=>setPage("pricing")} style={{background:"none",border:"none",color:"#059669",fontWeight:"600",fontSize:"14px",cursor:"pointer"}}>
            See pricing →
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Tool page wrapper ──

export default AboutPage;
