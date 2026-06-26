// ─── CONTAINER SELECTOR ──────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import { PALLET_BASES, VEHICLES, VEHICLES_WITH_CUSTOM } from '../data/presets.js';
import { S } from './styles.jsx';
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


export default ContainerSelector;
