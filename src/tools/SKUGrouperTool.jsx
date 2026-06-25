// ─── SKU GROUPER ─────────────────────────────────────────────────────────────
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
import { S, UtilBadge } from '../components/styles.js';
import ContainerSelector from '../components/ContainerSelector.jsx';
import ConstraintsPanel from '../components/ConstraintsPanel.jsx';
import ThreeViewer from '../components/ThreeViewer.jsx';
import WAShare from '../components/WAShare.jsx';
import { TopView2D, SideView2D, IsoView2D } from '../components/Views2D.jsx';
function SKUGrouperTool({onSendToMultiSKU}){
  const[rawSkus,setRawSkus]=useState(null);
  const[fileName,setFileName]=useState("");
  const[k,setK]=useState(8);
  const[groups,setGroups]=useState(null);
  const[processing,setProcessing]=useState(false);
  const[error,setError]=useState("");
  const[dragOver,setDragOver]=useState(false);
  const[expandedGroup,setExpandedGroup]=useState(null);
  const[sent,setSent]=useState(false);

  const parseFile=(file)=>{
    if(!file)return;
    setFileName(file.name);setError("");setGroups(null);setSent(false);
    const reader=new FileReader();
    reader.onload=(e)=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        // Auto-detect header row
        let dataStart=0;
        for(let i=0;i<Math.min(raw.length,20);i++){
          const row=raw[i].map(c=>String(c).toLowerCase());
          if(row.some(c=>c.includes("sku")||c.includes("name")||c.includes("length")||c.includes("width"))){
            dataStart=i+1;break;
          }
        }
        const skus=[];
        for(let i=dataStart;i<raw.length;i++){
          const r=raw[i];
          if(!r[0]&&!r[1])continue;
          const L=parseFloat(r[1])||0,W=parseFloat(r[2])||0,H=parseFloat(r[3])||0;
          if(L<=0||W<=0||H<=0)continue; // skip invalid
          skus.push({name:String(r[0]||`Row ${i+1}`).trim(),L,W,H,
            weight:parseFloat(r[4])||0,qty:Math.max(1,parseInt(r[5])||1)});
        }
        if(skus.length<2){setError("Need at least 2 valid SKUs with dimensions.");return;}
        setRawSkus(skus);
      }catch(err){setError("Could not read file. Please check the format.");}
    };
    reader.readAsArrayBuffer(file);
  };

  const runGrouping=()=>{
    if(!rawSkus){setError("Upload a file first.");return;}
    setProcessing(true);setGroups(null);setSent(false);
    // Defer to let UI update
    setTimeout(()=>{
      try{
        const result=runKMeans(rawSkus,k);
        setGroups(result);
        setExpandedGroup(null);
      }catch(err){setError("Grouping failed: "+err.message);}
      setProcessing(false);
    },30);
  };

  const exportExcel=()=>{
    if(!groups)return;
    const wb=XLSX.utils.book_new();
    // Sheet 1: Group summary
    const summaryHeaders=["Group","Rep L (mm)","Rep W (mm)","Rep H (mm)","SKU Count","Total Qty","Avg Wt/Box (kg)","Accuracy %"];
    const summaryRows=groups.map(g=>[g.name,g.repL,g.repW,g.repH,g.skuCount,g.totalQty,g.avgWt>0?g.avgWt.toFixed(2):"",g.accuracy+"%"]);
    const ws1=XLSX.utils.aoa_to_sheet([["SKU GROUPER RESULTS"],[],["File:",fileName],["Groups:",k],[],summaryHeaders,...summaryRows]);
    ws1["!cols"]=[{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:18},{wch:12}];
    XLSX.utils.book_append_sheet(wb,ws1,"Group Summary");
    // Sheet 2: Full detail
    const detailHeaders=["SKU Name","Length","Width","Height","Weight","Qty","Assigned Group","Group Rep L","Group Rep W","Group Rep H"];
    const detailRows=groups.flatMap(g=>g.members.map(m=>[m.name,m.L,m.W,m.H,m.weight||"",m.qty||1,g.name,g.repL,g.repW,g.repH]));
    const ws2=XLSX.utils.aoa_to_sheet([detailHeaders,...detailRows]);
    ws2["!cols"]=[{wch:20},{wch:10},{wch:10},{wch:10},{wch:10},{wch:8},{wch:14},{wch:14},{wch:14},{wch:14}];
    XLSX.utils.book_append_sheet(wb,ws2,"SKU Detail");
    XLSX.writeFile(wb,"SKU_Groups.xlsx");
  };

  const sendToMultiSKU=()=>{
    if(!groups||!onSendToMultiSKU)return;
    const rows=groups.slice(0,8).map((g,i)=>({
      id:i+1,name:g.name,
      L:String(g.repL),W:String(g.repW),H:String(g.repH),
      weight:g.avgWt>0?String(Math.round(g.avgWt*10)/10):"",
      targetQty:String(g.totalQty),
    }));
    onSendToMultiSKU(rows);
    setSent(true);
  };

  const sizes=["XS","S","M","L","XL","XXL","3XL","4XL"];

  return(<div>
    <div style={S.sectionDesc}>
      Upload a list of up to 50,000 SKUs with dimensions. The grouper clusters them into
      2–8 representative box sizes using K-means — then sends the groups directly to the
      Multi-SKU Planner or Shipment Planner for full container planning.
    </div>

    {/* Upload area */}
    <div style={S.card}>
      <div style={S.cardTitle}>📂 Upload SKU List</div>
      <div onDragOver={e=>{e.preventDefault();setDragOver(true);}}
        onDragLeave={()=>setDragOver(false)}
        onDrop={e=>{e.preventDefault();setDragOver(false);parseFile(e.dataTransfer.files[0]);}}
        onClick={()=>document.getElementById("grouper-upload").click()}
        style={{border:`2px dashed ${dragOver?"#be185d":"#d1d9e0"}`,borderRadius:"10px",
          padding:"32px",textAlign:"center",cursor:"pointer",
          background:dragOver?"#fdf2f8":"#fafbfc",transition:"all 0.2s"}}>
        <input id="grouper-upload" type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}}
          onChange={e=>{parseFile(e.target.files[0]);e.target.value="";}}/>
        <div style={{fontSize:"32px",marginBottom:"8px"}}>📊</div>
        <div style={{fontWeight:"700",color:"#374151",marginBottom:"4px"}}>
          {fileName?fileName:"Drop Excel / CSV file here or click to browse"}
        </div>
        <div style={{fontSize:"12px",color:"#9ca3af"}}>
          Required columns: SKU Name · Length (mm) · Width (mm) · Height (mm) · Weight (kg) [optional] · Qty [optional]
        </div>
      </div>

      {/* Template download hint */}
      <div style={{...S.noteBox,marginTop:"12px"}}>
        <strong>Column order:</strong> A=SKU Name, B=Length, C=Width, D=Height, E=Weight per box (kg), F=Quantity.
        A header row is detected automatically. Rows with zero or missing dimensions are skipped.
        Supports up to 50,000 SKUs.
      </div>
    </div>

    {/* k selector + run */}
    <div style={S.card}>
      <div style={S.cardTitle}>⚙️ Grouping Settings</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"20px",alignItems:"end"}}>
        <div>
          <label style={S.label}>Number of Groups: <strong style={{color:"#be185d"}}>{k}</strong></label>
          <input type="range" min="2" max="8" step="1" value={k} onChange={e=>setK(+e.target.value)}
            style={{width:"100%",marginTop:"8px",accentColor:"#be185d"}}/>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:"11px",color:"#9ca3af",marginTop:"4px"}}>
            <span>2 groups</span><span>8 groups (max for Multi-SKU Planner)</span>
          </div>
          <div style={{...S.noteBox,marginTop:"8px",fontSize:"12px"}}>
            {rawSkus?`${rawSkus.length.toLocaleString()} SKUs loaded · ${k} representative sizes will be computed`:"Upload a file to see SKU count"}
          </div>
        </div>
        <div>
          <div style={{fontSize:"12px",color:"#6b7280",marginBottom:"8px"}}>
            The algorithm uses <strong>K-means clustering</strong> on box dimensions (L×W×H),
            weighted by quantity. Each group gets a representative box size — the weighted average
            of all SKUs in that cluster.
          </div>
          {error&&<div style={S.error}>⚠ {error}</div>}
          <button style={{...S.btnPrimary,width:"100%",opacity:(!rawSkus||processing)?0.6:1}}
            onClick={runGrouping} disabled={!rawSkus||processing}>
            {processing?"⏳ Grouping...":"▶ Group SKUs"}
          </button>
        </div>
      </div>
    </div>

    {/* Results */}
    {groups&&groups.length>0&&(<>
      {/* Summary */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",margin:"0 0 16px"}}>
        {[["SKUs Loaded",(rawSkus?.length||0).toLocaleString(),"#eff6ff","#1d4ed8"],
          ["Groups Created",groups.length.toLocaleString(),"#f0fdf4","#166534"],
          ["Total Qty",groups.reduce((s,g)=>s+g.totalQty,0).toLocaleString(),"#fff7ed","#c2410c"],
          ["Avg Accuracy",groups.reduce((s,g)=>s+g.accuracy,0)/groups.length|0+"%","#fdf2f8","#be185d"],
        ].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"22px",fontWeight:"800",color:col}}>{v}</div>
            <div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div>
          </div>))}
      </div>

      {/* Action buttons */}
      <div style={{display:"flex",gap:"10px",marginBottom:"16px",flexWrap:"wrap"}}>
        <button onClick={sendToMultiSKU}
          style={{...S.btnPrimary,flex:2,background:sent?"#166534":"#be185d",
          boxShadow:`0 4px 16px rgba(${sent?"22,163,74":"190,24,93"},0.35)`}}>
          {sent?"✓ Sent to Multi-SKU Planner!":"→ Use in Multi-SKU Planner"}
        </button>
        <button onClick={exportExcel}
          style={{flex:1,padding:"10px 18px",border:"1px solid #bbf7d0",borderRadius:"8px",
          cursor:"pointer",fontWeight:"600",fontSize:"13px",background:"#f0fdf4",
          color:"#166534",fontFamily:"inherit"}}>
          ⬇ Download Excel
        </button>
      </div>
      {sent&&<div style={{...S.noteBox,marginBottom:"16px",background:"#f0fdf4",borderColor:"#bbf7d0",color:"#166534"}}>
        ✓ Groups pre-filled in Multi-SKU Planner. Click the <strong>Multi-SKU Planner</strong> tab to continue.
      </div>}

      {/* Groups table */}
      <div style={{...S.card,padding:"0",overflow:"hidden"}}>
        <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",fontWeight:"700",fontSize:"13px"}}>
          {groups.length} Size Groups — click any row to see member SKUs
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
            <thead><tr>
              {["Group","Rep Box (L×W×H mm)","SKU Count","Total Qty","Avg Wt/Box","Accuracy","Volume"].map(h=>(
                <th key={h} style={{padding:"10px 14px",textAlign:"left",fontWeight:"600",fontSize:"11px",
                  color:"#6b7a8d",textTransform:"uppercase",letterSpacing:"0.05em",
                  background:"#f8fafc",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>))}
            </tr></thead>
            <tbody>
              {groups.map((g,i)=>{
                const isOpen=expandedGroup===i;
                const vol=g.repL*g.repW*g.repH;
                return(<>
                  <tr key={i} onClick={()=>setExpandedGroup(isOpen?null:i)}
                    style={{cursor:"pointer",background:isOpen?"#fdf2f8":i%2===0?"#fff":"#fafbfc",
                    borderLeft:`3px solid ${MULTI_LABELS[i%MULTI_LABELS.length]}`}}>
                    <td style={{padding:"10px 14px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                        <div style={{width:"14px",height:"14px",borderRadius:"3px",
                          background:MULTI_LABELS[i%MULTI_LABELS.length],flexShrink:0}}/>
                        <div>
                          <div style={{fontWeight:"700",color:"#111827"}}>{g.name}</div>
                          <div style={{fontSize:"10px",color:"#9ca3af"}}>{sizes[i]||"XL"} size</div>
                        </div>
                      </div>
                    </td>
                    <td style={{padding:"10px 14px",fontWeight:"600",color:"#374151",fontVariantNumeric:"tabular-nums"}}>
                      {fmtN(g.repL)} × {fmtN(g.repW)} × {fmtN(g.repH)}
                    </td>
                    <td style={{padding:"10px 14px",color:"#374151"}}>{g.skuCount.toLocaleString()}</td>
                    <td style={{padding:"10px 14px",fontWeight:"600",color:"#374151"}}>{g.totalQty.toLocaleString()}</td>
                    <td style={{padding:"10px 14px",color:"#6b7280"}}>{g.avgWt>0?g.avgWt.toFixed(2)+" kg":"—"}</td>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{background:g.accuracy>=90?"#dcfce7":g.accuracy>=70?"#fef9c3":"#fdf2f8",
                        color:g.accuracy>=90?"#166534":g.accuracy>=70?"#854d0e":"#831843",
                        padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"700"}}>
                        {g.accuracy}%
                      </span>
                    </td>
                    <td style={{padding:"10px 14px",color:"#9ca3af",fontSize:"12px"}}>
                      {(vol/1e6).toFixed(3)} L
                    </td>
                  </tr>
                  {/* Expanded member list */}
                  {isOpen&&(
                    <tr key={`exp-${i}`}>
                      <td colSpan={7} style={{padding:"0",borderBottom:"1px solid #f1f5f9"}}>
                        <div style={{background:"#fdf2f8",padding:"12px 18px"}}>
                          <div style={{fontSize:"12px",fontWeight:"700",color:"#be185d",marginBottom:"8px"}}>
                            {g.skuCount} SKUs in {g.name} — showing first 50
                          </div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:"6px",maxHeight:"120px",overflowY:"auto"}}>
                            {g.members.slice(0,50).map((m,j)=>(
                              <span key={j} style={{background:"#fff",border:"1px solid #fbcfe8",
                                borderRadius:"6px",padding:"2px 8px",fontSize:"11px",color:"#831843"}}>
                                {m.name} ({fmtN(m.L)}×{fmtN(m.W)}×{fmtN(m.H)})
                              </span>))}
                            {g.skuCount>50&&<span style={{fontSize:"11px",color:"#9ca3af",padding:"2px 8px"}}>
                              +{g.skuCount-50} more — see Excel export</span>}
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </>);})}
            </tbody>
          </table>
        </div>
        <div style={{padding:"10px 18px",borderTop:"1px solid #f1f5f9",background:"#f8fafc",
          fontSize:"12px",color:"#9ca3af",display:"flex",justifyContent:"space-between"}}>
          <span>Accuracy = how closely each SKU's actual size matches its group's representative size</span>
          <span>Higher groups = larger box volume</span>
        </div>
      </div>
    </>)}
  </div>);}


export default SKUGrouperTool;
