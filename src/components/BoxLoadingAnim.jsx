// ─── BOX LOADING ANIMATION ───────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
export const ANIM_COLS=7,ANIM_ROWS=4,ANIM_TOTAL=28;
const BOX_COLORS=["#be185d","#1a1a1a","#9d174d","#2d1020","#7c1d4e","#111111","#4a0d2a"];

function BoxLoadingAnim(){
  const[count,setCount]=useState(0);
  const[paused,setPaused]=useState(false);
  useEffect(()=>{
    if(paused){const t=setTimeout(()=>{setCount(0);setPaused(false);},2200);return()=>clearTimeout(t);}
    if(count>=ANIM_TOTAL){setPaused(true);return;}
    const t=setTimeout(()=>setCount(c=>c+1),72);
    return()=>clearTimeout(t);
  },[count,paused]);
  return(
    <div style={{position:"relative",userSelect:"none"}}>
      <div style={{border:"2px solid #be185d",borderRadius:"10px",padding:"14px",
        background:"rgba(190,24,93,0.04)",position:"relative",overflow:"visible",
        boxShadow:"0 0 40px rgba(190,24,93,0.12)"}}>
        <div style={{position:"absolute",left:0,top:0,bottom:0,width:"18px",borderRadius:"8px 0 0 8px",
          background:"linear-gradient(180deg,#be185d,#7c1d4e)",opacity:0.7}}/>
        <div style={{marginLeft:"20px",display:"grid",gridTemplateColumns:`repeat(${ANIM_COLS},1fr)`,gap:"5px"}}>
          {Array.from({length:ANIM_TOTAL},(_,i)=>(
            <div key={i} style={{height:"46px",borderRadius:"4px",
              background:BOX_COLORS[i%BOX_COLORS.length],
              opacity:i<count?1:0,
              transform:i<count?"scale(1) translateY(0)":"scale(0.4) translateY(8px)",
              transition:i<count?"opacity 0.22s ease,transform 0.22s ease":"none",
              boxShadow:"inset 0 6px 0 rgba(255,255,255,0.07),inset 0 0 0 1px rgba(255,255,255,0.04)"}}/>
          ))}
        </div>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:"10px",
          marginLeft:"20px",fontSize:"11px",fontWeight:"600"}}>
          <span style={{color:"rgba(255,255,255,0.35)"}}>📦 {Math.min(count,ANIM_TOTAL)}/{ANIM_TOTAL} boxes</span>
          <span style={{color:count>=ANIM_TOTAL?"#be185d":"rgba(255,255,255,0.2)",transition:"color 0.3s"}}>
            {count>=ANIM_TOTAL?"✓ 100% utilized":"loading..."}</span>
        </div>
      </div>
      <div style={{position:"absolute",top:"-12px",right:"-12px",background:"#be185d",color:"#fff",
        padding:"5px 12px",borderRadius:"6px",fontSize:"11px",fontWeight:"700",
        boxShadow:"0 4px 16px rgba(190,24,93,0.5)"}}>✓ Optimal layout</div>
      <div style={{position:"absolute",bottom:"-12px",left:"-12px",background:"#111",
        color:"rgba(255,255,255,0.7)",padding:"5px 12px",borderRadius:"6px",fontSize:"11px",
        fontWeight:"600",border:"1px solid #333",boxShadow:"0 4px 12px rgba(0,0,0,0.4)"}}>
        3D model included →</div>
    </div>
  );
}


export default BoxLoadingAnim;
