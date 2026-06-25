// ─── FOOTER ──────────────────────────────────────────────────────────────────
import PackWiseLogo from '../components/PackWiseLogo.jsx';
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
            {[["tool","Single SKU Calculator"],["tool","Multi-SKU Planner"],["tool","Shipment Planner"],["tool","Bulk SKU Calculator"],["tool","SKU Grouper"]].map(([pg,l])=>(
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

export default Footer;
