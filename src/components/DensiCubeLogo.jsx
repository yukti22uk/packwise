// ─── DENSICUBE SVG LOGO ───────────────────────────────────────────────────────
function DensiCubeLogo({size=36}){
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


export default DensiCubeLogo;
