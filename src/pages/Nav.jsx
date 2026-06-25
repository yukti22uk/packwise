// ─── NAV ─────────────────────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import PackWiseLogo from '../components/PackWiseLogo.jsx';
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

// ── Footer ──

export default Nav;
