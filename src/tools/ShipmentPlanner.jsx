// ─── SHIPMENT PLANNER ────────────────────────────────────────────────────────
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
import PasteFromExcel from '../components/PasteFromExcel.jsx';
import { TopView2D, SideView2D, IsoView2D } from '../components/Views2D.jsx';
function ShipmentPlanner(){
  // ── Shared state ──
  const[mode,setMode]=useState("single"); // "single" | "multi"
  const[bl,setBl]=useState(0);const[bw,setBw]=useState(0);const[bh,setBh]=useState(0);
  const[maxWt,setMaxWt]=useState(0);const[contName,setContName]=useState("");
  const[noStack,setNoStack]=useState(false);const[lockHeight,setLockHeight]=useState(false);
  const[maxStack,setMaxStack]=useState("");
  const[freightCost,setFreightCost]=useState("");const[brand,setBrand]=useState("");
  const[error,setError]=useState("");const[freightByType,setFreightByType]=useState({});

  // ── Single SKU state ──
  const[sl,setSl]=useState("");const[sw,setSw]=useState("");const[sh,setSh]=useState("");
  const[weight,setWeight]=useState("");const[orderQty,setOrderQty]=useState("");
  const[result,setResult]=useState(null);const[compare,setCompare]=useState(null);
  const captureRef=useRef(null);

  // ── Multi-SKU state ──
  const nextId=useRef(3);
  const[skuRows,setSkuRows]=useState([
    {id:1,name:"SKU 1",L:"",W:"",H:"",weight:"",qty:""},
    {id:2,name:"SKU 2",L:"",W:"",H:"",weight:"",qty:""},
  ]);
  const[multiResult,setMultiResult]=useState(null);
  const[selectedCont,setSelectedCont]=useState(0);
  const[multiView,setMultiView]=useState("manifest");

  const addSkuRow=()=>setSkuRows(r=>[...r,{id:nextId.current++,name:`SKU ${nextId.current-1}`,L:"",W:"",H:"",weight:"",qty:""}]);
  const removeSkuRow=(id)=>setSkuRows(r=>r.filter(x=>x.id!==id));
  const updateSkuRow=(id,f,v)=>setSkuRows(r=>r.map(x=>x.id===id?{...x,[f]:v}:x));

  const opt=()=>({noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt});

  // ── Single SKU calculate ──
  const calcSingle=()=>{
    const sku={L:parseFloat(sl)||0,W:parseFloat(sw)||0,H:parseFloat(sh)||0,weight:parseFloat(weight)||0};
    const order=parseInt(orderQty)||0;
    if(sku.L<=0||sku.W<=0||sku.H<=0){setError("Enter all box dimensions.");return;}
    if(bl<=0||bw<=0||bh<=0){setError("Select a container.");return;}
    if(order<=0){setError("Enter total order quantity.");return;}
    setError("");
    const e=effectivePerContainer(bl,bw,bh,sku,opt());
    if(e.eff<=0){setError("Box doesn't fit in this container.");return;}
    const perContainer=e.eff,containers=Math.ceil(order/perContainer),lastFill=order-(containers-1)*perContainer;
    const freight=parseFloat(freightCost)||0,totalFreight=freight*containers,costPerUnit=freight>0?totalFreight/order:null;
    const usedVol=(perContainer*sku.L*sku.W*sku.H)/(bl*bw*bh);
    setResult({sku,order,perContainer,containers,lastFill,freight,totalFreight,costPerUnit,
      packResult:e,usedVol,contName,maxWt,totalWeight:sku.weight>0?order*sku.weight:null,constraint:e.constraint});
    const rows=VEHICLES.map(v=>{
      const ev=effectivePerContainer(v.L,v.W,v.H,sku,{noStack,lockHeight,maxStack:parseInt(maxStack)||0,maxWeight:maxWt});
      const per=ev.eff;if(per<=0)return{...v,per:0,containers:0,fc:0,total:0,cpu:null};
      const cont=Math.ceil(order/per),fc=parseFloat(freightByType[v.label])||0;
      return{...v,per,containers:cont,fc,total:fc*cont,cpu:fc>0?(fc*cont)/order:null};});
    setCompare(rows);
  };

  // ── Multi-SKU calculate ──
  const calcMulti=()=>{
    if(bl<=0||bw<=0||bh<=0){setError("Select a container.");return;}
    const skus=skuRows.map(r=>({name:r.name||`SKU ${r.id}`,L:parseFloat(r.L)||0,
      W:parseFloat(r.W)||0,H:parseFloat(r.H)||0,weight:parseFloat(r.weight)||0,qty:parseInt(r.qty)||0}));
    const valid=skus.filter(s=>s.L>0&&s.W>0&&s.H>0&&s.qty>0);
    if(valid.length<2){setError("Enter at least 2 SKUs with all dimensions and quantity to ship.");return;}
    setError("");
    const res=calcMultiSKUShipment(bl,bw,bh,maxWt,valid,opt());
    setMultiResult({...res,contName,freight:parseFloat(freightCost)||0,brand});
    setSelectedCont(0);setMultiView("manifest");
    // Cost comparison — estimate containers per vehicle using total volume
    const totalVol=valid.reduce((s,sk)=>s+sk.qty*sk.L*sk.W*sk.H,0);
    const rows=VEHICLES.map(v=>{
      const vVol=v.L*v.W*v.H;const perC=vVol>0?Math.floor(vVol*0.82/(totalVol/res.totalContainers)):0;
      const est=res.totalContainers>0?Math.max(1,Math.round(res.totalContainers*(bl*bw*bh)/(v.L*v.W*v.H))):0;
      const fc=parseFloat(freightByType[v.label])||0;
      return{...v,containers:est,fc,total:fc*est,cpu:fc>0&&valid.reduce((s,sk)=>s+sk.qty,0)>0?
        (fc*est)/valid.reduce((s,sk)=>s+sk.qty,0):null};});
    setCompare(rows);
  };

  const updateFreight=(label,val)=>{
    const nf={...freightByType,[label]:val};setFreightByType(nf);
    if(compare) setCompare(compare.map(r=>{if(r.label!==label)return r;
      const fc=parseFloat(val)||0,units=mode==="single"?(parseInt(orderQty)||1):
        (multiResult?multiResult.totalOrdered:1);
      return{...r,fc,total:fc*r.containers,cpu:fc>0&&r.containers>0?(fc*r.containers)/units:null};}));
  };

  // ── Single-SKU PDF ──
  const exportSinglePDF=()=>{if(!result)return;const img=captureRef.current?captureRef.current():null;
    const doc=new jsPDF({unit:"pt",format:"a4"});const W=doc.internal.pageSize.getWidth();
    doc.setFillColor(15,23,42);doc.rect(0,0,W,64,"F");doc.setTextColor(255,255,255);doc.setFontSize(18);doc.setFont(undefined,"bold");
    doc.text(brand||"Container Loading Plan",40,36);doc.setFontSize(10);doc.setFont(undefined,"normal");
    doc.text("Shipment Loading & Cost Report  ·  Generated "+new Date().toLocaleDateString(),40,52);
    let y=88;doc.setTextColor(30,41,59);doc.setFontSize(13);doc.setFont(undefined,"bold");doc.text("Shipment Summary",40,y);y+=18;
    doc.setFontSize(10);doc.setFont(undefined,"normal");
    const lines=[["Container / Vehicle",result.contName],["Box dimensions (L×W×H)",`${fmtN(result.sku.L)} × ${fmtN(result.sku.W)} × ${fmtN(result.sku.H)} mm`],
      ["Total order quantity",result.order.toLocaleString()+" units"],["Units per container",result.perContainer.toLocaleString()],
      ["Containers required",result.containers.toLocaleString()],["Volume utilization",(result.usedVol*100).toFixed(1)+"%"],["Limited by",result.constraint]];
    if(result.totalWeight!=null)lines.push(["Total shipment weight",money(result.totalWeight)+" kg"]);
    if(result.costPerUnit!=null){lines.push(["Freight cost per container",money(result.freight)]);lines.push(["Total freight cost",money(result.totalFreight)]);lines.push(["Cost per unit shipped",money(result.costPerUnit)]);}
    lines.forEach(([k,v])=>{doc.setFont(undefined,"bold");doc.text(k+":",48,y);doc.setFont(undefined,"normal");doc.text(String(v),230,y);y+=15;});y+=6;
    if(img){try{doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Loading Visualization",40,y);y+=10;doc.addImage(img,"PNG",40,y,260,140);y+=160;}catch(e){}}
    if(y>700){doc.addPage();y=50;}doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Per-Container Manifest",40,y);y+=18;doc.setFontSize(9);
    doc.setFillColor(190,24,93);doc.setTextColor(255,255,255);doc.rect(40,y-10,W-80,16,"F");
    doc.text("Container #",48,y);doc.text("Units Loaded",180,y);doc.text("Fill %",320,y);doc.text("Weight (kg)",420,y);y+=14;doc.setTextColor(30,41,59);
    for(let i=1;i<=Math.min(result.containers,200);i++){if(y>790){doc.addPage();y=50;}
      const units=i<result.containers?result.perContainer:result.lastFill,fill=((units/result.perContainer)*100).toFixed(0)+"%",wt=result.sku.weight>0?money(units*result.sku.weight):"—";
      if(i%2===0){doc.setFillColor(245,247,250);doc.rect(40,y-10,W-80,14,"F");}
      doc.text("#"+i,48,y);doc.text(units.toLocaleString(),180,y);doc.text(fill,320,y);doc.text(wt,420,y);y+=14;}y+=16;
    if(y>700){doc.addPage();y=50;}doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Loading Instructions",40,y);y+=18;doc.setFont(undefined,"normal");doc.setFontSize(10);
    const instr=[`1. Orient each box as ${result.packResult.boxL}×${result.packResult.boxW}×${result.packResult.boxH} mm.`,
      `2. Build floor layer: ${result.packResult.nx} along length × ${result.packResult.ny} across width.`,
      noStack?"3. DO NOT STACK — fragile items, single layer only.":`3. Stack ${result.packResult.nz} layer(s) high.`,
      `4. Each full container: ${result.perContainer.toLocaleString()} units. Last container: ${result.lastFill.toLocaleString()} units.`];
    instr.forEach(t=>{const s=doc.splitTextToSize(t,W-90);doc.text(s,48,y);y+=s.length*13;});y+=14;
    if(y>760){doc.addPage();y=50;}doc.setFontSize(8);doc.setTextColor(120,120,120);
    doc.text(doc.splitTextToSize("Results are algorithmic estimates. Verify before dispatch.",W-80),40,y);
    doc.save("Shipment_Plan_"+(brand||"Single_SKU").replace(/\s+/g,"_")+".pdf");};

  // ── Multi-SKU PDF ──
  const exportMultiPDF=()=>{if(!multiResult)return;
    const doc=new jsPDF({unit:"pt",format:"a4"});const W=doc.internal.pageSize.getWidth();
    doc.setFillColor(15,23,42);doc.rect(0,0,W,64,"F");doc.setTextColor(255,255,255);
    doc.setFontSize(18);doc.setFont(undefined,"bold");doc.text(brand||"Multi-SKU Shipment Plan",40,36);
    doc.setFontSize(10);doc.setFont(undefined,"normal");
    doc.text("Multi-SKU Loading Report  ·  "+new Date().toLocaleDateString(),40,52);
    let y=88;doc.setTextColor(30,41,59);doc.setFontSize(13);doc.setFont(undefined,"bold");doc.text("Shipment Summary",40,y);y+=18;
    doc.setFontSize(10);doc.setFont(undefined,"normal");
    const sumLines=[["Container / Vehicle",contName],["Total containers needed",multiResult.totalContainers.toLocaleString()],
      ["Total units ordered",multiResult.totalOrdered.toLocaleString()],["Total units shipped",multiResult.totalShipped.toLocaleString()],
      ["Average utilization",(multiResult.avgUtil*100).toFixed(1)+"%"],
      ["SKUs in this shipment",multiResult.skus.length.toLocaleString()]];
    if(multiResult.freight>0){sumLines.push(["Freight per container",money(multiResult.freight)]);
      sumLines.push(["Total freight cost",money(multiResult.freight*multiResult.totalContainers)]);
      sumLines.push(["Cost per unit",money(multiResult.freight*multiResult.totalContainers/Math.max(1,multiResult.totalShipped))]);}
    sumLines.forEach(([k,v])=>{doc.setFont(undefined,"bold");doc.text(k+":",48,y);doc.setFont(undefined,"normal");doc.text(String(v),250,y);y+=15;});y+=10;
    // Per-container manifest table
    if(y>700){doc.addPage();y=50;}
    doc.setFont(undefined,"bold");doc.setFontSize(13);doc.text("Per-Container Manifest",40,y);y+=18;
    const skuNames=multiResult.skus.map(s=>s.name);
    const colW=Math.min(80,Math.floor((W-200)/Math.max(1,skuNames.length)));
    doc.setFontSize(8);doc.setFillColor(190,24,93);doc.setTextColor(255,255,255);
    doc.rect(40,y-10,W-80,14,"F");doc.text("Container",48,y);doc.text("Total",120,y);
    skuNames.forEach((n,i)=>doc.text(n.substring(0,8),180+i*colW,y));doc.text("Util%",W-100,y);y+=12;
    doc.setTextColor(30,41,59);
    multiResult.containers.forEach((c,i)=>{if(y>790){doc.addPage();y=50;}
      if(i%2===0){doc.setFillColor(248,250,252);doc.rect(40,y-9,W-80,13,"F");}
      doc.text("#"+c.id,48,y);doc.text(c.total.toLocaleString(),120,y);
      skuNames.forEach((n,j)=>{const r=c.regions.find(x=>x.name===n);doc.text((r?r.fitted:0).toString(),180+j*colW,y);});
      doc.text((c.volUtil*100).toFixed(0)+"%",W-100,y);y+=13;});y+=16;
    if(y>760){doc.addPage();y=50;}doc.setFontSize(8);doc.setTextColor(120,120,120);
    doc.text("Results are algorithmic estimates. Verify before dispatch.",40,y);
    doc.save("MultiSKU_Shipment_"+(brand||"Plan").replace(/\s+/g,"_")+".pdf");};

  const cheapest=compare?compare.filter(r=>r.cpu!=null).sort((a,b)=>a.cpu-b.cpu)[0]:null;

  // ── 3D regions for selected container in multi mode ──
  const selContData=multiResult&&multiResult.containers[selectedCont];
  const multiRegions3D=selContData?selContData.regions.map((r,i)=>({
    col:MULTI_PALETTE[i%MULTI_PALETTE.length],
    ox:r.off.x,oy:r.off.y,oz:r.off.z||0,
    rnx:r.det.nx,rny:r.det.ny,rnz:r.det.nz,
    bL:r.det.boxL,bW:r.det.boxW,bH:r.det.boxH})):null;
  const multiContResult=selContData?{cL:selContData.cL,cW:selContData.cW,cH:selContData.cH,
    nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,
    leftover1:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover2:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0},
    leftover3:{nx:0,ny:0,nz:0,boxL:1,boxW:1,boxH:1,offX:0,offY:0,offZ:0}}:null;

  return(<div>
    <div style={S.sectionDesc}>
      Plan a full shipment across multiple containers. Single SKU: one box size, get exact containers needed and cost.
      Multi-SKU: multiple box types shipped together — per-container manifest for every truck.
    </div>

    {/* ── Mode toggle ── */}
    <div style={{display:"flex",gap:"8px",marginBottom:"20px"}}>
      {[["single","📦 Single SKU"],["multi","🗃️ Multi-SKU"]].map(([m,label])=>(
        <button key={m} onClick={()=>{setMode(m);setError("");setResult(null);setMultiResult(null);setCompare(null);}}
          style={{padding:"10px 24px",border:`2px solid ${mode===m?"#be185d":"#e2e8f0"}`,
          background:mode===m?"#be185d":"#fff",color:mode===m?"#fff":"#6b7280",
          borderRadius:"10px",fontWeight:"700",fontSize:"14px",cursor:"pointer",fontFamily:"inherit",
          transition:"all 0.15s"}}>
          {label}{m==="multi"&&<span style={{marginLeft:"6px",fontSize:"10px",background:mode===m?"rgba(255,255,255,0.25)":"#fef3c7",color:mode===m?"#fff":"#92400e",padding:"1px 6px",borderRadius:"99px"}}>PRO</span>}
        </button>))}
    </div>

    {/* ── Shared: container selector ── */}
    <ContainerSelector onChange={(L,W,H,wt,name)=>{setBl(L);setBw(W);setBh(H);setMaxWt(wt);setContName(name);}} showWeight={true} vehicleOnly={true}/>

    {/* ── Single SKU inputs ── */}
    {mode==="single"&&(<>
      <div style={S.card}><div style={S.cardTitle}>📦 Box Dimensions</div>
        <div style={S.grid3}>{[["Length",sl,setSl],["Width",sw,setSw],["Height",sh,setSh]].map(([l,v,s])=>(
          <div key={l}><label style={S.label}>{l}</label>
          <input style={S.input} type="number" min="0" step="any" value={v} onChange={e=>s(e.target.value)} placeholder="0"/></div>))}</div></div>
      <ConstraintsPanel weight={weight} setWeight={setWeight} noStack={noStack} setNoStack={setNoStack}
        lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>
    </>)}

    {/* ── Multi-SKU inputs ── */}
    {mode==="multi"&&(<>
      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"14px",gap:"8px",flexWrap:"wrap"}}>
          <div style={S.cardTitle}>📦 SKUs to Ship</div>
          <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
            <PasteFromExcel mode="shipment" onFill={(rows)=>{setSkuRows(rows.slice(0,8).map(r=>({...r,qty:r.qty||r.targetQty||''})));nextId.current=rows.length+1;setMultiResult(null);}}/>
            <button onClick={addSkuRow} disabled={skuRows.length>=8}
              style={{padding:"7px 14px",background:"#f0fdf4",border:"1px solid #bbf7d0",
              borderRadius:"8px",fontSize:"13px",fontWeight:"600",color:"#166534",cursor:"pointer",fontFamily:"inherit"}}>
              + Add SKU</button>
          </div>
        </div>
        {/* Header */}
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1fr 1.2fr 40px",gap:"8px",marginBottom:"8px",padding:"0 4px"}}>
          {["SKU Name","L (mm)","W (mm)","H (mm)","Wt/Box kg","Qty to Ship",""].map(h=>(
            <div key={h} style={S.label}>{h}</div>))}
        </div>
        {skuRows.map((row,idx)=>(
          <div key={row.id} className="sku-row" style={{display:"grid",
            gridTemplateColumns:"2fr 1fr 1fr 1fr 1fr 1.2fr 40px",gap:"8px",
            marginBottom:"8px",padding:"6px 4px",borderRadius:"8px",transition:"background 0.15s"}}>
            <input style={{...S.input,borderColor:MULTI_LABELS[idx%MULTI_LABELS.length]+"55"}}
              value={row.name} onChange={e=>updateSkuRow(row.id,"name",e.target.value)} placeholder={`SKU ${idx+1}`}/>
            {["L","W","H"].map(f=>(
              <input key={f} style={S.input} type="number" min="0" step="any"
                value={row[f]} onChange={e=>updateSkuRow(row.id,f,e.target.value)} placeholder="0"/>))}
            <input style={S.input} type="number" min="0" step="any"
              value={row.weight||""} onChange={e=>updateSkuRow(row.id,"weight",e.target.value)} placeholder="0"/>
            <input style={S.input} type="number" min="0"
              value={row.qty||""} onChange={e=>updateSkuRow(row.id,"qty",e.target.value)} placeholder="qty"/>
            <button onClick={()=>removeSkuRow(row.id)} disabled={skuRows.length<=2}
              style={{background:"none",border:"1px solid #e2e8f0",borderRadius:"6px",
              cursor:"pointer",color:"#9ca3af",fontSize:"16px",opacity:skuRows.length<=2?0.3:1}}>✕</button>
          </div>))}
        <div style={{...S.noteBox,marginTop:"8px"}}>
          Qty to Ship is your total order quantity per SKU. Each container will be packed with a mix of all SKUs.
          Max 8 SKUs. Weight per box is optional but recommended if your container has a weight limit.
        </div>
      </div>
      <ConstraintsPanel hideWeight={true} noStack={noStack} setNoStack={setNoStack}
        lockHeight={lockHeight} setLockHeight={setLockHeight} maxStack={maxStack} setMaxStack={setMaxStack}/>
    </>)}

    {/* ── Shared: order & cost ── */}
    <div style={S.card}><div style={S.cardTitle}>🧾 Cost & Branding</div>
      <div style={{display:"grid",gridTemplateColumns:mode==="single"?"1fr 1fr 1fr":"1fr 1fr",gap:"12px"}}>
        {mode==="single"&&<div><label style={S.label}>Total Order Quantity (units)</label>
          <input style={S.input} type="number" min="0" value={orderQty} onChange={e=>setOrderQty(e.target.value)} placeholder="e.g. 5000"/></div>}
        <div><label style={S.label}>Freight Cost per Container (₹)</label>
          <input style={S.input} type="number" min="0" value={freightCost} onChange={e=>setFreightCost(e.target.value)} placeholder="e.g. 45000"/></div>
        <div><label style={S.label}>Company Name (for PDF)</label>
          <input style={S.input} type="text" value={brand} onChange={e=>setBrand(e.target.value)} placeholder="Your company"/></div>
      </div></div>

    {error&&<div style={S.error}>⚠ {error}</div>}
    <button style={S.btnPrimary} onClick={mode==="single"?calcSingle:calcMulti}>
      ▶ {mode==="single"?"Plan Single-SKU Shipment":"Plan Multi-SKU Shipment"}
    </button>

    {/* ══════════════════════════════════════
        SINGLE SKU RESULTS
    ══════════════════════════════════════ */}
    {mode==="single"&&result&&(<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",margin:"20px 0 16px"}}>
        {[["Containers Needed",result.containers.toLocaleString(),"#eff6ff","#1d4ed8"],
          ["Units / Container",result.perContainer.toLocaleString(),"#f0fdf4","#166534"],
          ["Cost / Unit",result.costPerUnit!=null?money(result.costPerUnit):"—","#fff7ed","#c2410c"],
          ["Volume Util.",(result.usedVol*100).toFixed(1)+"%",result.usedVol>=0.75?"#f0fdf4":"#fefce8",result.usedVol>=0.75?"#166534":"#854d0e"]
        ].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"20px",fontWeight:"700",color:col,wordBreak:"break-word"}}>{v}</div>
            <div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}
      </div>
      <div style={{display:"flex",gap:"12px",marginBottom:"16px",flexWrap:"wrap"}}>
        <button style={{...S.btnPrimary,flex:1,background:"#be185d"}} onClick={exportSinglePDF}>⬇ Download PDF Loading Plan</button>
        <WAShare message={`🚚 *DensiCube Shipment Plan*\nVehicle: ${result.contName}\nBox: ${fmtN(result.sku.L)}×${fmtN(result.sku.W)}×${fmtN(result.sku.H)} mm | ${result.order.toLocaleString()} units\n*${result.containers} containers needed*\n${result.perContainer.toLocaleString()} units per container\nSpace used: ${(result.usedVol*100).toFixed(1)}%${result.costPerUnit!=null?`\nCost per unit: ${money(result.costPerUnit)}`:""}\n\nPlan your shipment at densicube.netlify.app`}/>
      </div>
      <div style={S.card}><div style={S.cardTitle}>💰 Container Comparison</div>
        <div style={{fontSize:"12px",color:"#6b7a8d",marginBottom:"12px"}}>Enter freight cost per vehicle type to find the cheapest option.</div>
        <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
          <thead><tr>{["Vehicle","Units/Container","Containers","Freight Each","Total Cost","Cost/Unit"].map(h=>(
            <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
          <tbody>{compare.map((r,i)=>{const isCheapest=cheapest&&r.label===cheapest.label;
            return(<tr key={i} style={{background:isCheapest?"#dcfce7":i%2===0?"#fff":"#fafbfc"}}>
              <td style={{padding:"8px 12px",fontWeight:isCheapest?"700":"500"}}>{r.label}{isCheapest&&" ✅"}</td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.per.toLocaleString()}</td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.containers.toLocaleString()}</td>
              <td style={{padding:"8px 12px"}}><input type="number" min="0" value={freightByType[r.label]||""} placeholder="0" onChange={e=>updateFreight(r.label,e.target.value)} style={{width:"90px",border:"1px solid #d1d9e0",borderRadius:"6px",padding:"4px 8px",fontSize:"12px"}}/></td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.total>0?money(r.total):"—"}</td>
              <td style={{padding:"8px 12px",textAlign:"right",fontWeight:"700",color:isCheapest?"#166534":"#374151"}}>{r.cpu!=null?money(r.cpu):"—"}</td>
            </tr>);})}
          </tbody></table></div>
        {cheapest&&<div style={{marginTop:"12px",padding:"10px 14px",background:"#f0fdf4",borderRadius:"8px",fontSize:"13px",color:"#166534"}}>
          ✅ Cheapest: <strong>{cheapest.label}</strong> at <strong>{money(cheapest.cpu)}</strong>/unit ({cheapest.containers} containers × {money(cheapest.fc)} each)</div>}
      </div>
      <div style={{...S.card,padding:"0",overflow:"hidden"}}>
        <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",fontWeight:"600",fontSize:"13px"}}>Per-Container Manifest</div>
        <div style={{overflowX:"auto",maxHeight:"260px",overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
          <thead><tr>{["Container #","Units Loaded","Fill %","Weight"].map(h=>(
            <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0}}>{h}</th>))}</tr></thead>
          <tbody>{Array.from({length:Math.min(result.containers,200)},(_,i)=>{
            const n=i+1,units=n<result.containers?result.perContainer:result.lastFill,fill=((units/result.perContainer)*100).toFixed(0);
            return(<tr key={i} style={{background:i%2===0?"#fff":"#fafbfc"}}>
              <td style={{padding:"7px 12px",fontWeight:"500"}}>#{n}</td>
              <td style={{padding:"7px 12px",textAlign:"right"}}>{units.toLocaleString()}</td>
              <td style={{padding:"7px 12px"}}><span style={{background:+fill>=90?"#dcfce7":+fill>=50?"#fef9c3":"#fdf2f8",color:+fill>=90?"#166534":+fill>=50?"#854d0e":"#831843",padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"600"}}>{fill}%</span></td>
              <td style={{padding:"7px 12px",textAlign:"right"}}>{result.sku.weight>0?money(units*result.sku.weight)+" kg":"—"}</td>
            </tr>);})}
          </tbody></table></div>
      </div>
      {result.packResult&&<div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}>
          <div style={S.cardTitle}>🔄 Single Container Loading (3D)</div>
          <span style={{fontSize:"11px",color:"#94a3b8"}}>Included in PDF</span>
        </div>
        <ThreeViewer result={result.packResult} captureRef={captureRef}/>
      </div>}
    </>)}

    {/* ══════════════════════════════════════
        MULTI-SKU RESULTS
    ══════════════════════════════════════ */}
    {mode==="multi"&&multiResult&&(<>
      {/* Summary cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",margin:"20px 0 16px"}}>
        {[["Containers Needed",multiResult.totalContainers.toLocaleString(),"#eff6ff","#1d4ed8"],
          ["Units Shipped",`${multiResult.totalShipped.toLocaleString()} / ${multiResult.totalOrdered.toLocaleString()}`,"#f0fdf4","#166534"],
          ["Avg Utilization",(multiResult.avgUtil*100).toFixed(1)+"%",multiResult.avgUtil>=0.75?"#f0fdf4":"#fefce8",multiResult.avgUtil>=0.75?"#166534":"#854d0e"],
          ["Cost / Unit",multiResult.freight>0?money((multiResult.freight*multiResult.totalContainers)/Math.max(1,multiResult.totalShipped)):"—","#fff7ed","#c2410c"],
        ].map(([l,v,bg,col])=>(
          <div key={l} style={{background:bg,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
            <div style={{fontSize:"20px",fontWeight:"700",color:col,wordBreak:"break-word"}}>{v}</div>
            <div style={{fontSize:"11px",color:"#6b7a8d",marginTop:"4px"}}>{l}</div></div>))}
      </div>

      {/* Action buttons */}
      <div style={{display:"flex",gap:"10px",marginBottom:"16px",flexWrap:"wrap"}}>
        {[["manifest","📋 Manifest"],["compare","💰 Cost Comparison"],["3d","🔄 3D Viewer"]].map(([id,label])=>(
          <button key={id} onClick={()=>setMultiView(id)}
            style={{padding:"8px 18px",border:"none",borderRadius:"8px",cursor:"pointer",
            fontWeight:"600",fontSize:"13px",fontFamily:"inherit",
            background:multiView===id?"#be185d":"#f1f5f9",color:multiView===id?"#fff":"#374151"}}>
            {label}</button>))}
        <button onClick={exportMultiPDF}
          style={{padding:"8px 18px",border:"1px solid #bbf7d0",borderRadius:"8px",cursor:"pointer",
          fontWeight:"600",fontSize:"13px",background:"#f0fdf4",color:"#166534",fontFamily:"inherit",marginLeft:"auto"}}>
          ⬇ Download PDF</button>
        <WAShare message={`🚚 *DensiCube Multi-SKU Shipment*\nVehicle: ${contName}\n*${multiResult.totalContainers} containers needed*\n${multiResult.totalShipped.toLocaleString()} units across ${multiResult.skus.length} SKUs\nAvg utilization: ${(multiResult.avgUtil*100).toFixed(1)}%\n\nPlan your shipment at densicube.netlify.app`}/>
      </div>

      {/* Per-container manifest */}
      {multiView==="manifest"&&(<>
        <div style={{...S.card,padding:"0",overflow:"hidden"}}>
          <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div style={{fontWeight:"700",fontSize:"13px"}}>Per-Container Manifest — {multiResult.totalContainers} containers</div>
            <div style={{fontSize:"12px",color:"#9ca3af"}}>Click a row to view its 3D model</div>
          </div>
          <div style={{overflowX:"auto",maxHeight:"340px",overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
              <thead><tr>
                <th style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0,whiteSpace:"nowrap"}}>#</th>
                {multiResult.skus.map((sk,i)=>(
                  <th key={i} style={{padding:"9px 12px",textAlign:"right",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0,whiteSpace:"nowrap"}}>
                    <span style={{display:"inline-block",width:"10px",height:"10px",borderRadius:"2px",background:MULTI_LABELS[i%MULTI_LABELS.length],marginRight:"5px"}}/>
                    {sk.name}</th>))}
                <th style={{padding:"9px 12px",textAlign:"right",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0}}>Total</th>
                <th style={{padding:"9px 12px",textAlign:"right",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0}}>Util%</th>
                <th style={{padding:"9px 12px",textAlign:"right",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",position:"sticky",top:0}}>Weight</th>
              </tr></thead>
              <tbody>
                {multiResult.containers.map((c,i)=>(
                  <tr key={i} onClick={()=>{setSelectedCont(i);setMultiView("3d");}}
                    style={{background:selectedCont===i&&multiView==="3d"?"#fdf2f8":i%2===0?"#fff":"#fafbfc",cursor:"pointer",
                    borderLeft:selectedCont===i&&multiView==="3d"?`3px solid #be185d`:"3px solid transparent"}}>
                    <td style={{padding:"8px 12px",fontWeight:"600",color:"#374151"}}>#{c.id}</td>
                    {multiResult.skus.map((sk,j)=>{const r=c.regions.find(x=>x.name===sk.name);const qty=r?r.fitted:0;
                      return(<td key={j} style={{padding:"8px 12px",textAlign:"right",color:qty>0?"#111827":"#d1d5db",fontWeight:qty>0?"600":"400"}}>{qty.toLocaleString()}</td>);})}
                    <td style={{padding:"8px 12px",textAlign:"right",fontWeight:"700"}}>{c.total.toLocaleString()}</td>
                    <td style={{padding:"8px 12px",textAlign:"right"}}>
                      <span style={{background:(c.volUtil)>=0.75?"#dcfce7":(c.volUtil)>=0.5?"#fef9c3":"#fdf2f8",
                        color:(c.volUtil)>=0.75?"#166534":(c.volUtil)>=0.5?"#854d0e":"#831843",
                        padding:"2px 8px",borderRadius:"99px",fontSize:"11px",fontWeight:"600"}}>
                        {(c.volUtil*100).toFixed(0)}%</span></td>
                    <td style={{padding:"8px 12px",textAlign:"right",color:"#6b7280"}}>
                      {c.totalWeight>0?money(c.totalWeight)+" kg":"—"}</td>
                  </tr>))}
              </tbody>
            </table>
          </div>
          {/* SKU totals row */}
          <div style={{padding:"10px 18px",borderTop:"1px solid #f1f5f9",background:"#f8fafc",
            display:"flex",gap:"16px",flexWrap:"wrap",fontSize:"12px"}}>
            <span style={{fontWeight:"700",color:"#374151"}}>Totals shipped:</span>
            {multiResult.skus.map((sk,i)=>(
              <span key={i} style={{display:"flex",alignItems:"center",gap:"5px"}}>
                <span style={{width:"10px",height:"10px",borderRadius:"2px",background:MULTI_LABELS[i%MULTI_LABELS.length],display:"inline-block"}}/>
                <span style={{fontWeight:"600"}}>{sk.name}:</span>
                <span style={{color:multiResult.shipped[sk.name]>=sk.qty?"#166534":"#c2410c"}}>
                  {(multiResult.shipped[sk.name]||0).toLocaleString()} / {sk.qty.toLocaleString()}</span>
              </span>))}
          </div>
        </div>
      </>)}

      {/* Cost comparison */}
      {multiView==="compare"&&compare&&(<div style={S.card}>
        <div style={S.cardTitle}>💰 Container Comparison — find the cheapest option</div>
        <div style={{fontSize:"12px",color:"#6b7a8d",marginBottom:"12px"}}>
          Estimated containers per vehicle type based on volume. Enter freight cost to compare.
        </div>
        <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
          <thead><tr>{["Vehicle","Est. Containers","Freight Each","Total Cost","Cost/Unit"].map(h=>(
            <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:"600",fontSize:"11px",color:"#6b7a8d",textTransform:"uppercase",background:"#f8fafc",borderBottom:"1px solid #e8edf2",whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
          <tbody>{compare.map((r,i)=>{const isCheapest=cheapest&&r.label===cheapest.label;
            return(<tr key={i} style={{background:isCheapest?"#dcfce7":i%2===0?"#fff":"#fafbfc"}}>
              <td style={{padding:"8px 12px",fontWeight:isCheapest?"700":"500"}}>{r.label}{isCheapest&&" ✅"}</td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.containers.toLocaleString()}</td>
              <td style={{padding:"8px 12px"}}><input type="number" min="0" value={freightByType[r.label]||""} placeholder="0" onChange={e=>updateFreight(r.label,e.target.value)} style={{width:"90px",border:"1px solid #d1d9e0",borderRadius:"6px",padding:"4px 8px",fontSize:"12px"}}/></td>
              <td style={{padding:"8px 12px",textAlign:"right"}}>{r.total>0?money(r.total):"—"}</td>
              <td style={{padding:"8px 12px",textAlign:"right",fontWeight:"700",color:isCheapest?"#166534":"#374151"}}>{r.cpu!=null?money(r.cpu):"—"}</td>
            </tr>);})}
          </tbody></table></div>
        {cheapest&&<div style={{marginTop:"12px",padding:"10px 14px",background:"#f0fdf4",borderRadius:"8px",fontSize:"13px",color:"#166534"}}>
          ✅ Cheapest: <strong>{cheapest.label}</strong> at <strong>{money(cheapest.cpu)}</strong>/unit ({cheapest.containers} containers × {money(cheapest.fc)} each)</div>}
      </div>)}

      {/* 3D viewer for selected container */}
      {multiView==="3d"&&multiContResult&&(<div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px",flexWrap:"wrap",gap:"10px"}}>
          <div style={S.cardTitle}>🔄 Container #{selectedCont+1} of {multiResult.totalContainers}</div>
          <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
            <button onClick={()=>setSelectedCont(Math.max(0,selectedCont-1))} disabled={selectedCont===0}
              style={{padding:"6px 14px",border:"1px solid #e2e8f0",borderRadius:"8px",cursor:"pointer",
              background:"#fff",fontFamily:"inherit",fontWeight:"600",opacity:selectedCont===0?0.4:1}}>← Prev</button>
            <span style={{fontSize:"12px",color:"#6b7280",minWidth:"80px",textAlign:"center"}}>
              {selectedCont+1} / {multiResult.totalContainers}</span>
            <button onClick={()=>setSelectedCont(Math.min(multiResult.totalContainers-1,selectedCont+1))} disabled={selectedCont===multiResult.totalContainers-1}
              style={{padding:"6px 14px",border:"1px solid #e2e8f0",borderRadius:"8px",cursor:"pointer",
              background:"#fff",fontFamily:"inherit",fontWeight:"600",opacity:selectedCont===multiResult.totalContainers-1?0.4:1}}>Next →</button>
          </div>
        </div>
        {/* Mini manifest for this container */}
        <div style={{display:"flex",gap:"10px",flexWrap:"wrap",marginBottom:"14px"}}>
          {selContData.regions.filter(r=>r.fitted>0).map((r,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:"6px",padding:"5px 10px",
              background:"#f8fafc",borderRadius:"8px",border:"1px solid #e2e8f0",fontSize:"12px"}}>
              <div style={{width:"10px",height:"10px",borderRadius:"2px",background:MULTI_LABELS[i%MULTI_LABELS.length]}}/>
              <span style={{fontWeight:"600"}}>{r.name}:</span>
              <span>{r.fitted.toLocaleString()} units</span>
            </div>))}
          <div style={{padding:"5px 10px",background:"#f0fdf4",borderRadius:"8px",border:"1px solid #bbf7d0",fontSize:"12px",color:"#166534",fontWeight:"600"}}>
            {(selContData.volUtil*100).toFixed(1)}% utilized
          </div>
        </div>
        <ThreeViewer result={multiContResult} regions3D={multiRegions3D}/>
        <div style={{display:"flex",gap:"10px",flexWrap:"wrap",marginTop:"10px",justifyContent:"center"}}>
          {selContData.regions.filter(r=>r.fitted>0).map((r,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}>
              <div style={{width:"11px",height:"11px",background:MULTI_LABELS[i%MULTI_LABELS.length],borderRadius:"2px"}}/>
              {r.name} ({r.fitted.toLocaleString()})</div>))}
          <div style={{display:"flex",alignItems:"center",gap:"4px",fontSize:"11px",color:"#555"}}>
            <div style={{width:"11px",height:"11px",background:"#1e293b",borderRadius:"2px"}}/>Container</div>
        </div>
      </div>)}
    </>)}
  </div>);}



export default ShipmentPlanner;
