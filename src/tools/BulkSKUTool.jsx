// ─── BULK SKU CALCULATOR ─────────────────────────────────────────────────────
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
import TemplateDownload from '../components/TemplateDownload.jsx';
import { TopView2D, SideView2D, IsoView2D } from '../components/Views2D.jsx';
function ContainerSkuTool({isPro,onUpgrade}){
  const[cL,setCL]=useState("");const[cW,setCW]=useState("");const[cH,setCH]=useState("");const[cMaxWt,setCMaxWt]=useState("");
  const[fileName,setFileName]=useState("");const[results,setResults]=useState(null);const[error,setError]=useState("");
  const[processing,setProcessing]=useState(false);const[progress,setProgress]=useState(0);const[skuCount,setSkuCount]=useState(0);
  const[dragOver,setDragOver]=useState(false);const[rawSkus,setRawSkus]=useState(null);const[capped,setCapped]=useState(false);
  const container={cL:parseFloat(cL)||0,cW:parseFloat(cW)||0,cH:parseFloat(cH)||0,cMaxWt:parseFloat(cMaxWt)||0};
  const valid=container.cL>0&&container.cW>0&&container.cH>0&&container.cMaxWt>0;
  function pSkus(cont,skus){const{cL,cW,cH,cMaxWt}=cont,cv=cL*cW*cH;
    return skus.map(s=>{const{name,sl,sw,sh,swt,qtyAvail}=s;if(!name)return null;if(sl<=0||sw<=0||sh<=0)return{name,error:"Invalid dimensions"};
      const{total:vQ,orient}=calcMixed(cL,cW,cH,sl,sw,sh);let eV=qtyAvail>0?Math.min(vQ,qtyAvail):vQ;let wQ=swt>0?Math.floor(cMaxWt/swt):null;
      if(wQ!==null&&qtyAvail>0)wQ=Math.min(wQ,qtyAvail);const eQ=wQ!==null?Math.min(eV,wQ):eV,vu=(eQ*sl*sw*sh)/cv,wu=swt>0?(eQ*swt)/cMaxWt:null;
      let con="Volume";if(wQ!==null&&wQ<eV)con="Weight";if(qtyAvail>0&&eQ===qtyAvail)con="Stock Limit";
      return{name,volQty:eV,wtQty:wQ!==null?wQ:"N/A",effQty:eQ,volUtil:vu,wtUtil:wu,orient,constraint:con};}).filter(Boolean);}
  const parseFile=(file)=>{if(!file)return;setFileName(file.name);setError("");setResults(null);
    const r=new FileReader();r.onload=(e)=>{try{const wb=XLSX.read(e.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];
      const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});let ds=0;
      for(let i=0;i<Math.min(raw.length,20);i++){const row=raw[i].map(c=>String(c).toLowerCase());if(row.some(c=>c.includes("sku")||c.includes("length")||c.includes("width"))){ds=i+1;break;}}
      const skus=[];for(let i=ds;i<raw.length;i++){const rr=raw[i];if(!rr[0]&&!rr[1])continue;
        skus.push({name:String(rr[0]||"").trim(),sl:parseFloat(rr[1])||0,sw:parseFloat(rr[2])||0,sh:parseFloat(rr[3])||0,swt:parseFloat(rr[4])||0,qtyAvail:parseFloat(rr[5])>0?parseFloat(rr[5]):-1});}
      if(skus.length===0){setError("No SKU data found.");return;}setSkuCount(skus.length);setRawSkus(skus);}catch(err){setError("Could not read file: "+err.message);}};r.readAsArrayBuffer(file);};
  const run=()=>{if(!valid){setError("Enter container dimensions.");return;}if(!rawSkus){setError("Upload SKU file.");return;}setError("");
    let toProcess=rawSkus,cap=false;
    if(!isPro&&rawSkus.length>CONFIG.freeSkuLimit){toProcess=rawSkus.slice(0,CONFIG.freeSkuLimit);cap=true;}
    setCapped(cap);setProcessing(true);setProgress(0);setResults(null);const CHUNK=500;let done=0;const all=[];
    function next(){all.push(...pSkus(container,toProcess.slice(done,done+CHUNK)));done+=CHUNK;
      setProgress(Math.round((Math.min(done,toProcess.length)/toProcess.length)*100));
      if(done<toProcess.length)setTimeout(next,0);else{setResults(all);setProcessing(false);}}setTimeout(next,50);};
  const exp=()=>{const wb=XLSX.utils.book_new();const h=["SKU Name","Max Qty (Volume)","Max Qty (Weight)","Effective Max Qty","Volume Used (%)","Weight Used (%)","Best Orientation","Constraint"];
    const rows=results.map(r=>r.error?[r.name,r.error,"","","","","",""]:[r.name,r.volQty,r.wtQty,r.effQty,(r.volUtil*100).toFixed(2)+"%",r.wtUtil!=null?(r.wtUtil*100).toFixed(2)+"%":"",r.orient,r.constraint]);
    const ws=XLSX.utils.aoa_to_sheet([["CONTAINER SKU PACKING RESULTS"],[],["Container",`${container.cL}×${container.cW}×${container.cH}`,"Max Weight",container.cMaxWt],[],h,...rows]);
    ws["!cols"]=[{wch:22},{wch:16},{wch:16},{wch:16},{wch:14},{wch:14},{wch:26},{wch:14}];ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:7}}];
    XLSX.utils.book_append_sheet(wb,ws,"Results");XLSX.writeFile(wb,"Container_Packing_Results.xlsx");};
  const dlT=()=>{const ws=XLSX.utils.aoa_to_sheet([["SKU Name","Length","Width","Height","Weight per Box","Qty Available (optional)"],["SKU-001",30,20,15,2.5,""],["SKU-002",25,25,10,1.8,500]]);
    ws["!cols"]=[{wch:18},{wch:10},{wch:10},{wch:10},{wch:16},{wch:22}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"SKU List");XLSX.writeFile(wb,"SKU_Template.xlsx");};
  const gC=results?results.filter(r=>!r.error&&r.volUtil>=0.75).length:0;const oC=results?results.filter(r=>!r.error&&r.volUtil>=0.5&&r.volUtil<0.75).length:0;const lC=results?results.filter(r=>!r.error&&r.volUtil<0.5).length:0;
  return(<div>
    <div style={S.sectionDesc}>Upload an Excel file with multiple SKUs — each with dimensions and weight. Get the maximum quantity per SKU constrained by both volume and weight. Download full results as Excel.{!isPro&&<span style={{color:"#c2410c"}}> Free plan processes up to {CONFIG.freeSkuLimit} SKUs.</span>}</div>
    <div style={{display:"grid",gridTemplateColumns:"320px 1fr",gap:"20px",alignItems:"start"}}>
      <div>
        <div style={S.card}><div style={S.cardTitle}>🗃️ Container Details</div>
          <div style={S.grid2}>{[["Length",cL,setCL],["Width",cW,setCW],["Height",cH,setCH],["Max Weight",cMaxWt,setCMaxWt]].map(([l,v,s])=>(
            <div key={l}><label style={S.label}>{l}</label><input style={S.input} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div>
          {valid&&<div style={S.infoBox}>Volume: {(container.cL*container.cW*container.cH).toLocaleString()}</div>}</div>
        <div style={S.card}><div style={S.cardTitle}>📂 Upload SKU File</div>
          <div style={S.dropzone(dragOver)} onDragOver={e=>{e.preventDefault();setDragOver(true)}} onDragLeave={()=>setDragOver(false)} onDrop={e=>{e.preventDefault();setDragOver(false);parseFile(e.dataTransfer.files[0])}} onClick={()=>document.getElementById("fi2").click()}>
            <div style={{fontSize:"28px",marginBottom:"6px"}}>📂</div><div style={{fontSize:"13px",fontWeight:"500",color:"#374151"}}>{fileName||"Drop Excel or click to browse"}</div>
            <div style={{fontSize:"11px",color:"#9ca3af",marginTop:"4px"}}>.xlsx or .xls</div>
            {rawSkus&&<div style={{marginTop:"6px",fontSize:"12px",color:"#059669",fontWeight:"600"}}>✓ {skuCount.toLocaleString()} SKUs loaded</div>}
            <input id="fi2" type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>parseFile(e.target.files[0])}/></div>
          <div style={S.noteBox}><strong>Columns:</strong> SKU Name | L | W | H | Weight | Qty</div>
          <TemplateDownload mode="bulk"/></div>
        <button style={S.btnPrimary} onClick={run}>▶ Calculate Container Fit</button>
        <button style={{...S.btnSecondary,marginTop:"10px"}} onClick={dlT}>⬇ Download Template</button>
      </div>
      <div>
        {error&&<div style={S.error}>⚠ {error}</div>}
        {capped&&<div style={{...S.card,background:"#fffbeb",border:"1px solid #fde68a",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:"13px",color:"#92400e"}}>⚠ Free plan limited to {CONFIG.freeSkuLimit} SKUs. {skuCount.toLocaleString()} uploaded — only first {CONFIG.freeSkuLimit} processed.</span>
          <button onClick={onUpgrade} style={{padding:"6px 14px",background:"#059669",color:"#fff",border:"none",borderRadius:"8px",fontWeight:"600",fontSize:"12px",cursor:"pointer",whiteSpace:"nowrap"}}>⭐ Go Pro</button></div>}
        {processing&&<div style={S.card}><div style={{fontSize:"13px",fontWeight:"500",color:"#374151",marginBottom:"8px"}}>Processing {Math.min(skuCount,isPro?skuCount:CONFIG.freeSkuLimit).toLocaleString()} SKUs... {progress}%</div>
          <div style={{background:"#e5e7eb",borderRadius:"99px",height:"10px"}}><div style={{height:"10px",borderRadius:"99px",background:"#059669",width:`${progress}%`,transition:"width 0.2s"}}/></div></div>}
        {results&&!processing&&(<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"12px",marginBottom:"16px"}}>
            {[["Total",results.length,"#f8fafc"],["≥75%",gC,"#f0fdf4"],["50–74%",oC,"#fefce8"],["<50%",lC,"#fff8fc"]].map(([l,v,bg])=>(
              <div key={l} style={{background:bg,borderRadius:"10px",padding:"12px",textAlign:"center",border:"1px solid rgba(0,0,0,0.06)"}}>
                <div style={{fontSize:"20px",fontWeight:"700",color:"#1a2332"}}>{v}</div><div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"2px"}}>{l}</div></div>))}</div>
          <button style={{...S.btnPrimary,marginBottom:"16px"}} onClick={exp}>⬇ Download Results as Excel</button>
          <div style={{...S.card,padding:"0",overflow:"hidden"}}>
            <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:"600",fontSize:"13px"}}>Preview</span><span style={{fontSize:"12px",color:"#9ca3af"}}>{results.length.toLocaleString()} SKUs</span></div>
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
              <thead><tr>{["SKU","Vol Qty","Wt Qty","Eff Qty","Vol%","Wt%","Orientation","Constraint"].map(h=>(<th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
              <tbody>{results.slice(0,100).map((r,i)=>r.error?(<tr key={i} style={{background:"#fff8fc"}}><td style={{padding:"8px 12px",fontWeight:"500"}}>{r.name}</td><td style={{padding:"8px 12px",color:"#be185d"}} colSpan={7}>{r.error}</td></tr>
              ):(<tr key={i} style={{background:i%2===0?"#fff":"#fafbfc"}}>
                <td style={{padding:"8px 12px",fontWeight:"500"}}>{r.name}</td><td style={{padding:"8px 12px",textAlign:"right"}}>{r.volQty?.toLocaleString()}</td>
                <td style={{padding:"8px 12px",textAlign:"right"}}>{typeof r.wtQty==="number"?r.wtQty.toLocaleString():r.wtQty}</td><td style={{padding:"8px 12px",textAlign:"right",fontWeight:"700"}}>{r.effQty?.toLocaleString()}</td>
                <td style={{padding:"8px 12px"}}><UtilBadge val={r.volUtil}/></td><td style={{padding:"8px 12px"}}><UtilBadge val={r.wtUtil}/></td>
                <td style={{padding:"8px 12px",color:"#6b7a8d",whiteSpace:"nowrap"}}>{r.orient}</td>
                <td style={{padding:"8px 12px"}}><span style={{padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"500",background:r.constraint==="Volume"?"#eff6ff":r.constraint==="Weight"?"#fff7ed":"#f5f3ff",color:r.constraint==="Volume"?"#1d4ed8":r.constraint==="Weight"?"#c2410c":"#6d28d9"}}>{r.constraint}</span></td></tr>))}</tbody>
            </table>{results.length>100&&<div style={{padding:"10px 18px",fontSize:"12px",color:"#9ca3af",borderTop:"1px solid #f1f5f9"}}>Showing 100 of {results.length.toLocaleString()} — download Excel for all</div>}</div></div>
        </>)}
        {!results&&!processing&&!error&&<div style={{...S.card,padding:"60px",textAlign:"center",color:"#9ca3af"}}><div style={{fontSize:"48px",marginBottom:"12px"}}>📦</div><div style={{fontWeight:"500"}}>Fill in container details and upload your SKU file</div></div>}
      </div>
    </div>
  </div>);}


export default ContainerSkuTool;
