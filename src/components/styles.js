// ─── SHARED STYLES & UtilBadge ───────────────────────────────────────────────
function UtilBadge({val}){
  if(val==null||val==="") return <span style={{color:"#9ca3af",fontSize:"12px"}}>—</span>;
  const pct=typeof val==="number"?val*100:parseFloat(val);
  const bg=pct>=75?"#dcfce7":pct>=50?"#fef9c3":"#fdf2f8";
  const color=pct>=75?"#166534":pct>=50?"#854d0e":"#831843";
  return <span style={{background:bg,color,padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"600"}}>{pct.toFixed(1)}%</span>;
}

const S={
  card:{background:"#fff",borderRadius:"12px",padding:"20px",marginBottom:"16px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:"1px solid #e8edf2"},
  cardTitle:{fontSize:"14px",fontWeight:"700",color:"#1a2332",marginBottom:"14px"},
  sectionDesc:{fontSize:"13px",color:"#6b7a8d",marginBottom:"20px",lineHeight:"1.6",background:"#f8fafc",borderRadius:"8px",padding:"12px 16px"},
  grid2:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px"},
  grid3:{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"12px"},
  label:{display:"block",fontSize:"11px",fontWeight:"600",color:"#6b7a8d",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:"4px"},
  input:{width:"100%",border:"1px solid #d1d9e0",borderRadius:"8px",padding:"8px 12px",fontSize:"14px",boxSizing:"border-box",outline:"none",fontFamily:"inherit"},
  infoBox:{background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:"8px",padding:"8px 12px",fontSize:"12px",color:"#166534",marginTop:"10px"},
  dropzone:(d)=>({border:`2px dashed ${d?"#10b981":"#d1d9e0"}`,borderRadius:"10px",padding:"24px",textAlign:"center",cursor:"pointer",transition:"all 0.2s",background:d?"#f0fdf4":"#fafbfc"}),
  noteBox:{background:"#f8fafc",borderRadius:"8px",padding:"10px 12px",fontSize:"12px",color:"#6b7a8d",marginTop:"10px",lineHeight:"1.6"},
  btnPrimary:{width:"100%",padding:"12px",background:"#be185d",color:"#fff",border:"none",borderRadius:"10px",fontSize:"14px",fontWeight:"600",cursor:"pointer",fontFamily:"inherit"},
  btnSecondary:{width:"100%",padding:"10px",background:"#fff",color:"#be185d",border:"1px solid #be185d",borderRadius:"10px",fontSize:"13px",fontWeight:"500",cursor:"pointer",fontFamily:"inherit"},
  error:{background:"#fff8fc",border:"1px solid #fecaca",borderRadius:"10px",padding:"12px",fontSize:"13px",color:"#be185d",marginBottom:"12px"},
};

export { S, UtilBadge };
