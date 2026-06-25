// ─── CONSTRAINTS PANEL ───────────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import { S } from './styles.jsx';
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


export default ConstraintsPanel;
