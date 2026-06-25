// ─── MULTI-SKU PLANNER ───────────────────────────────────────────────────────
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
import AIFillButton from '../components/AIFillButton.jsx';
import { TopView2D, SideView2D, IsoView2D } from '../components/Views2D.jsx';
function MultiSKUTool({preset,onPresetUsed}){
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);const[contName,setContName]=useState("");const[maxWt,setMaxWt]=useState(0);
  const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);const[maxStack,setMaxStack]=useState("");
  const[skuRows,setSkuRows]=useState([
    {id:1,name:"SKU 1",L:"",W:"",H:"",targetQty:"",weight:""},
    {id:2,name:"SKU 2",L:"",W:"",H:"",targetQty:"",weight:""},
  ]);
  const[result,setResult]=useState(null);const[error,setError]=useState("");const[view,setView]=useState("table");
  const nextId=useRef(3);

  // Apply preset from SKU Grouper
  useEffect(()=>{
    if(!preset||!preset.length) return;
    setSkuRows(preset.map((r,i)=>({...r,id:i+1})));
    nextId.current=preset.length+1;
    setResult(null);setError("");
    if(onPresetUsed) onPresetUsed();
  },[preset]);

  const handleAIFill=(rows)=>{
    setSkuRows(rows.slice(0,8));
    nextId.current=rows.length+1;
    setResult(null);
  };

  const addRow=()=>{
    setSkuRows(r=>[...r,{id:nextId.current++,name:`SKU ${nextId.current-1}`,L:"",W:"",H:"",targetQty:"",weight:""}]);};
  const removeRow=(id)=>setSkuRows(r=>r.filter(row=>row.id!==id));
  const updateRow=(id,field,val)=>setSkuRows(r=>r.map(row=>row.id===id?{...row,[field]:val}:row));

  const calc=()=>{
    if(bl<=0||bw<=0||bh<=0){setError("Select a container / enter dimensions.");return;}
    const skus=skuRows.map(r=>({...r,L:parseFloat(r.L)||0,W:parseFloat(r.W)||0,H:parseFloat(r.H)||0,targetQty:parseInt(r.targetQty)||0,weight:parseFloat(r.weight)||0}));
    const valid=skus.filter(s=>s.L>0&&s.W>0&&s.H>0&&s.targetQty>0);
    if(valid.length<2){setError("Enter at least 2 SKUs with all dimensions and target quantity.");return;}
    if(valid.length>8){setError("Maximum 8 SKUs supported.");return;}
    setError("");
    const opt={noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt};
    const r=calcMultiSKU(bl,bw,bh,valid,opt);
    setResult(r?{...r,skuCount:valid.length}:null);
  };

  const exportExcel=()=>{
    if(!result) return;
    const wb=XLSX.utils.book_new();
    const h=["SKU Name","Length","Width","Height","Target Qty","Fitted Qty","Fill %","Volume Util","Orientation"];
    const rows=result.regions.map(r=>[r.name,r.L,r.W,r.H,r.target,r.fitted,
      (r.fillRate*100).toFixed(1)+"%",(r.skuVol*100).toFixed(1)+"%",
      `${fmtN(r.det.boxL)}×${fmtN(r.det.boxW)}×${fmtN(r.det.boxH)}`]);
    const ws=XLSX.utils.aoa_to_sheet([
      ["MULTI-SKU CONTAINER PACKING PLAN"],[],
      ["Container",`${fmtN(bl)}×${fmtN(bw)}×${fmtN(bh)} mm`,"Total Boxes",result.total,"Volume Util",(result.volUtil*100).toFixed(1)+"%"],[],
      h,...rows]);
    ws["!cols"]=[{wch:16},{wch:10},{wch:10},{wch:10},{wch:12},{wch:12},{wch:10},{wch:12},{wch:22}];
    XLSX.utils.book_append_sheet(wb,ws,"Multi-SKU Plan");
    XLSX.writeFile(wb,"Multi_SKU_Packing_Plan.xlsx");
  };

  // Build regions3D for ThreeViewer
  const regions3D=result?result.regions.map((r,i)=>({
    col:MULTI_PALETTE[i%MULTI_PALETTE.length],
    ox:r.off.x,oy:r.off.y,oz:r.off.z||0,
    rnx:r.det.nx,rny:r.det.ny,rnz:r.det.nz,
    bL:r.det.boxL,bW:r.det.boxW,bH:r.det.boxH,
  })):null;
  const containerResult=result?{cL:result.cL,cW:result.cW,cH:result.cH,
    nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,
    leftover1:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover2:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover3:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0}}:null;

  return(<div>
    <div style={S.sectionDesc}>
      Pack multiple SKU sizes into one container simultaneously. Enter each SKU's dimensions
      and target quantity — the planner divides the container optimally and reports exactly
      how many of each SKU fits. Supports stacking and orientation constraints.
    </div>

    <ContainerSelector onChange={(L,W,H,wt,name)=>{setBl(L);setBw(W);setBh(H);setMaxWt(wt);setContName(name);}} showWeight={true}/>

    {/* SKU table */}
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"14px",gap:"8px",flexWrap:"wrap"}}>
        <div style={S.cardTitle}>📦 SKUs to Pack</div>
        <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
          <AIFillButton mode="multisku" onFill={handleAIFill}/>
          <button onClick={addRow} disabled={skuRows.length>=8}
            style={{padding:"7px 14px",background:"#f0fdf4",border:"1px solid #bbf7d0",
            borderRadius:"8px",fontSize:"13px",fontWeight:"600",color:"#166534",cursor:"pointer",fontFamily:"inherit"}}>
            + Add SKU
          </button>
        </div>
      </div>
      {/* Table header */}
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr 1fr 40px",gap:"8px",
        marginBottom:"8px",padding:"0 4px"}}>
        {["SKU Name","Length","Width","Height","Target Qty","Wt/Box (kg)",""].map(h=>(
          <div key={h} style={S.label}>{h}</div>
        ))}
      </div>
      {/* Rows */}
      {skuRows.map((row,idx)=>(
        <div key={row.id} className="sku-row" style={{display:"grid",
          gridTemplateColumns:"2fr 1fr 1fr 1fr 1.5fr 1fr 40px",gap:"8px",
          marginBottom:"8px",padding:"6px 4px",borderRadius:"8px",transition:"background 0.15s"}}>
          <input style={{...S.input,borderColor:MULTI_LABELS[idx%MULTI_LABELS.length]+"55"}}
            value={row.name} onChange={e=>updateRow(row.id,"name",e.target.value)}
            placeholder={`SKU ${idx+1}`}/>
          {["L","W","H"].map(f=>(
            <input key={f} style={S.input} type="number" min="0" step="any"
              value={row[f]} onChange={e=>updateRow(row.id,f,e.target.value)} placeholder="0"/>
          ))}
          <input style={S.input} type="number" min="0"
            value={row.targetQty} onChange={e=>updateRow(row.id,"targetQty",e.target.value)} placeholder="qty"/>
          <input style={S.input} type="number" min="0" step="any"
            value={row.weight||""} onChange={e=>updateRow(row.id,"weight",e.target.value)} placeholder="kg"/>
          <button onClick={()=>removeRow(row.id)} disabled={skuRows.length<=2}
            style={{background:"none",border:"1px solid #e2e8f0",borderRadius:"6px",cursor:"pointer",
            color:"#9ca3af",fontSize:"16px",fontFamily:"inherit",
            opacity:skuRows.length<=2?0.3:1}}>✕</button>
        </div>
      ))}
      <div style={{...S.noteBox,marginTop:"8px"}}>
        <strong>Tip:</strong> Target Qty is your order quantity. The planner allocates container space proportionally and reports the actual fitted quantity.
        Weight per box is optional — enter it to apply the container's weight limit.
        Max 8 SKUs · Container divided into {skuRows.filter(r=>parseFloat(r.L)>0).length||"N"} regions
      </div>
    </div>

    <ConstraintsPanel hideWeight={true} noStack={noStack} setNoStack={setNoStack}
      lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>

    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={calc}>▶ Plan Multi-SKU Container Load</button>

    {result&&(<>
      {/* Summary row */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",margin:"20px 0 16px"}}>
        {[
          ["Total Boxes Fitted",result.total.toLocaleString(),"#f0fdf4","#166534"],
          ["Volume Utilization",(result.volUtil*100).toFixed(1)+"%",result.volUtil>=0.75?"#f0fdf4":result.volUtil>=0.5?"#fef9c3":"#fef2f2",result.volUtil>=0.75?"#166534":result.volUtil>=0.5?"#854d0e":"#991b1b"],
          ["SKUs Packed",`${result.regions.filter(r=>r.fitted>0).length} / ${result.regions.length}`,"#eff6ff","#1d4ed8"],
          ["Limited By",result.regions.some(r=>r.weightConstrained)?"Weight":"Volume",result.regions.some(r=>r.weightConstrained)?"#fff7ed":"#f0fdf4",result.regions.some(r=>r.weightConstrained)?"#c2410c":"#166534"],
        ].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"22px",fontWeight:"800",color:col,letterSpacing:"-0.02em"}}>{v}</div>
            <div style={{fontSize:"11px",color:"#6b7280",marginTop:"4px",fontWeight:"600",textTransform:"uppercase",letterSpacing:"0.05em"}}>{l}</div>
          </div>
        ))}
      </div>

      {/* View tabs */}
      <div style={{display:"flex",gap:"6px",marginBottom:"16px",flexWrap:"wrap"}}>
        {[["table","📊 Per-SKU Results"],["3d","🔄 3D Model"]].map(([id,label])=>(
          <button key={id} onClick={()=>setView(id)} style={{padding:"8px 18px",border:"none",
            borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",
            background:view===id?"#be185d":"#f1f5f9",color:view===id?"#fff":"#374151",fontFamily:"inherit"}}>{label}</button>))}
        <button onClick={exportExcel} style={{padding:"8px 18px",border:"1px solid #bbf7d0",
          borderRadius:"8px",cursor:"pointer",fontWeight:"600",fontSize:"13px",
          background:"#f0fdf4",color:"#166534",fontFamily:"inherit",marginLeft:"auto"}}>
          ⬇ Download Excel
        </button>
      </div>

      {/* Per-SKU results table */}
      {view==="table"&&(
        <div style={{...S.card,padding:"0",overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
              <thead><tr>
                {["#","SKU Name","Target Qty","Fitted Qty","Fill %","Region Util","Orientation","Region Size"].map(h=>(
                  <th key={h} style={{padding:"10px 14px",textAlign:"left",fontWeight:"700",fontSize:"11px",
                    color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.05em",
                    background:"#f8fafc",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {result.regions.map((r,i)=>{
                  const fillOk=r.fillRate>=1,fillMid=r.fillRate>=0.7;
                  return(<tr key={i} style={{background:i%2===0?"#fff":"#fafbfc",
                    borderBottom:"1px solid #f1f5f9"}}>
                    <td style={{padding:"10px 14px"}}>
                      <div style={{width:"12px",height:"12px",borderRadius:"3px",
                        background:MULTI_LABELS[i%MULTI_LABELS.length],display:"inline-block"}}/>
                    </td>
                    <td style={{padding:"10px 14px",fontWeight:"700",color:"#111827"}}>{r.name}</td>
                    <td style={{padding:"10px 14px",color:"#6b7280"}}>{r.target.toLocaleString()}</td>
                    <td style={{padding:"10px 14px",fontWeight:"700",
                      color:fillOk?"#166534":fillMid?"#854d0e":"#991b1b"}}>
                      {r.fitted.toLocaleString()}</td>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{background:fillOk?"#dcfce7":fillMid?"#fef9c3":"#fee2e2",
                        color:fillOk?"#166534":fillMid?"#854d0e":"#991b1b",
                        padding:"2px 8px",borderRadius:"99px",fontSize:"12px",fontWeight:"700"}}>
                        {(r.fillRate*100).toFixed(0)}%</span></td>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{background:"#eff6ff",color:"#1d4ed8",
                        padding:"2px 8px",borderRadius:"99px",fontSize:"12px",fontWeight:"600"}}>
                        {(r.skuVol*100).toFixed(0)}%</span></td>
                    <td style={{padding:"10px 14px",color:"#6b7280",fontSize:"12px",whiteSpace:"nowrap"}}>
                      {fmtN(r.det.boxL)}×{fmtN(r.det.boxW)}×{fmtN(r.det.boxH)}</td>
                    <td style={{padding:"10px 14px",color:"#9ca3af",fontSize:"12px",whiteSpace:"nowrap"}}>
                      {fmtN(r.regionDims.L)}×{fmtN(r.regionDims.W)}×{fmtN(r.regionDims.H)}</td>
                  </tr>);
                })}
              </tbody>
            </table>
          </div>
          {/* Legend */}
          <div style={{padding:"12px 16px",borderTop:"1px solid #f1f5f9",
            display:"flex",gap:"16px",flexWrap:"wrap"}}>
            {result.regions.map((r,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"12px",color:"#555"}}>
                <div style={{width:"12px",height:"12px",borderRadius:"3px",
                  background:MULTI_LABELS[i%MULTI_LABELS.length]}}/>{r.name}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 3D Model */}
      {view==="3d"&&containerResult&&(
        <div style={S.card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}>
            <div style={S.cardTitle}>🔄 3D Model — Multi-SKU Loading</div>
            <span style={{fontSize:"11px",color:"#94a3b8"}}>Drag to rotate · Scroll to zoom</span>
          </div>
          <ThreeViewer result={containerResult} regions3D={regions3D}/>
          <div style={{display:"flex",gap:"12px",flexWrap:"wrap",marginTop:"10px",justifyContent:"center"}}>
            {result.regions.map((r,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}>
                <div style={{width:"11px",height:"11px",background:MULTI_LABELS[i%MULTI_LABELS.length],borderRadius:"2px"}}/>
                {r.name} ({r.fitted.toLocaleString()})</div>))}
            <div style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}>
              <div style={{width:"11px",height:"11px",background:"#1e293b",borderRadius:"2px"}}/>Container</div>
          </div>
        </div>
      )}
    </>)}
  </div>);}


export default MultiSKUTool;
