// ─── SINGLE SKU CALCULATOR ───────────────────────────────────────────────────
import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import * as THREE from 'three';
import { CONFIG } from '../config.js';
import { VEHICLES } from '../data/presets.js';
import { bestFitDetailed, calcMixedDetailed, effectivePerContainer, calcMixed } from '../algorithms/packing.js';
import { MULTI_PALETTE, MULTI_LABELS, calcMultiSKU, calcMultiSKUShipment } from '../algorithms/multiSku.js';
import { runKMeans } from '../algorithms/kmeans.js';
import { fmtN, money } from '../algorithms/utils.js';
import { S, UtilBadge } from '../components/styles.jsx';
import ContainerSelector from '../components/ContainerSelector.jsx';
import ConstraintsPanel from '../components/ConstraintsPanel.jsx';
import ThreeViewer from '../components/ThreeViewer.jsx';
import WAShare from '../components/WAShare.jsx';
import { TopView2D, SideView2D, IsoView2D } from '../components/Views2D.jsx';
function BoxPackingTool(){
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);const[maxWt,setMaxWt]=useState(0);
  const[sl,setSl]=useState("");const[sw,setSw]=useState("");const[sh,setSh]=useState("");
  const[weight,setWeight]=useState("");const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);const[maxStack,setMaxStack]=useState("");
  const[result,setResult]=useState(null);const[error,setError]=useState("");const[view,setView]=useState("3d");
  const calc=()=>{const sln=parseFloat(sl)||0,swn=parseFloat(sw)||0,shn=parseFloat(sh)||0;
    if(sln<=0||swn<=0||shn<=0){setError("Enter all 3 small box dimensions > 0.");return;}
    if(bl<=0||bw<=0||bh<=0){setError("Select a container / enter all dimensions.");return;}setError("");
    const opt={noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt,weight:parseFloat(weight)||0};
    const r=calcMixedDetailed(bl,bw,bh,sln,swn,shn,opt);
    let volQty=r.total,wtQty=null,effQty=volQty,constraint="Volume";const wt=parseFloat(weight)||0;
    if(wt>0&&maxWt>0){wtQty=Math.floor(maxWt/wt);if(wtQty<volQty){effQty=wtQty;constraint="Weight";}}
    setResult({...r,volQty,wtQty,effQty,constraint,volUtil:(effQty*sln*swn*shn)/(bl*bw*bh),totalWeight:wt>0?effQty*wt:null});};
  return(<div>
    <div style={S.sectionDesc}>Enter one box size and select a container or pallet — the calculator finds the maximum quantity using all 6 orientations, shows a rotatable 3D model, 2D engineering views, and leftover region breakdown.</div>
    <ContainerSelector onChange={(L,W,H,wt)=>{setBl(L);setBw(W);setBh(H);setMaxWt(wt);}} showWeight={true}/>
    <div style={S.card}><div style={S.cardTitle}>📦 Small Box Dimensions</div>
      <div style={S.grid3}>{[["Length",sl,setSl],["Width",sw,setSw],["Height",sh,setSh]].map(([l,v,s])=>(
        <div key={l}><label style={S.label}>{l}</label><input style={S.input} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div></div>
    <ConstraintsPanel weight={weight} setWeight={setWeight} noStack={noStack} setNoStack={setNoStack} lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>
    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={calc}>▶ Calculate & Visualise</button>
    {result&&(<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"12px",margin:"20px 0 16px"}}>
        {[["Effective Qty",result.effQty.toLocaleString(),"#f0fdf4","#166534"],
          ["Limited By",result.constraint,result.constraint==="Weight"?"#fff7ed":"#eff6ff",result.constraint==="Weight"?"#c2410c":"#1d4ed8"],
          ["Volume Util.",(result.volUtil*100).toFixed(1)+"%",result.volUtil>=0.75?"#f0fdf4":result.volUtil>=0.5?"#fefce8":"#fff8fc",result.volUtil>=0.75?"#166534":result.volUtil>=0.5?"#854d0e":"#831843"],
          ["Total Weight",result.totalWeight!=null?money(result.totalWeight)+" kg":"—","#f8fafc","#374151"],
        ].map(([l,v,bg,col])=>(<div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
          <div style={{fontSize:"16px",fontWeight:"700",color:col,wordBreak:"break-word"}}>{v}</div>
          <div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}
      </div>
      <div style={{...S.card,marginBottom:"16px"}}><div style={S.cardTitle}>Packing Breakdown by Region (volume fit)</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"10px"}}>
          {[["#3b82f6","Main Grid",`${result.nx}×${result.ny}×${result.nz}`,result.nx*result.ny*result.nz],
            ["#f97316","Leftover 1","Side gap",result.leftover1.count],["#22c55e","Leftover 2","Front gap",result.leftover2.count],
            ["#a855f7","Leftover 3","Top gap",result.leftover3.count]].map(([color,title,sub,cnt])=>(
            <div key={title} style={{background:"#f8fafc",borderRadius:"8px",padding:"12px",textAlign:"center",borderLeft:`4px solid ${color}`}}>
              <div style={{fontSize:"20px",fontWeight:"700",color:"#1e293b"}}>{cnt}</div>
              <div style={{fontSize:"12px",fontWeight:"600",color:"#374151"}}>{title}</div>
              <div style={{fontSize:"11px",color:"#9ca3af"}}>{sub}</div></div>))}</div></div>
      <div style={{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center",flexWrap:"wrap"}}>
        <div style={{display:"flex",gap:"6px",flex:1}}>
          {[["3d","🔄 3D Model"],["2d","📐 2D Views"]].map(([id,label])=>(
            <button key={id} onClick={()=>setView(id)} style={{padding:"8px 18px",border:"none",borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",background:view===id?"#059669":"#f1f5f9",color:view===id?"#fff":"#374151"}}>{label}</button>))}</div>
        <WAShare message={`📦 *PackWise Packing Result*\nContainer: ${fmtN(result.cL)}×${fmtN(result.cW)}×${fmtN(result.cH)} mm\nBox: ${fmtN(result.sl)}×${fmtN(result.sw)}×${fmtN(result.sh)} mm\n*Max boxes: ${result.effQty.toLocaleString()}*\nBest orientation: ${result.orient}\nSpace used: ${(result.volUtil*100).toFixed(1)}%\n\nCalculate your load free at packwise.netlify.app`}/>
      </div>
      {view==="3d"&&<div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}><div style={S.cardTitle}>🔄 3D Model</div><span style={{fontSize:"11px",color:"#94a3b8"}}>Drag to rotate · Scroll to zoom</span></div>
        <ThreeViewer result={result}/>
        <div style={{display:"flex",gap:"12px",flexWrap:"wrap",marginTop:"8px",justifyContent:"center"}}>
          {[["#3b82f6","Main"],["#f97316","Leftover 1"],["#22c55e","Leftover 2"],["#a855f7","Leftover 3"],["#1e293b","Container"]].map(([c,l])=>(
            <div key={l} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}><div style={{width:"11px",height:"11px",background:c,borderRadius:"2px"}}/>{l}</div>))}</div></div>}
      {view==="2d"&&<div style={S.card}><div style={S.cardTitle}>📐 2D Engineering Views</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"20px",alignItems:"start"}}>
          <TopView2D result={result}/><SideView2D result={result}/><IsoView2D result={result}/></div></div>}
    </>)}
  </div>);}


export default BoxPackingTool;
