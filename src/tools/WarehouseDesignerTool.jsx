// ─── WAREHOUSE DESIGNER TOOL ──────────────────────────────────────────────────
// Step 1: Warehouse parameters
// Step 2: Master SKU data (dimensions)
// Step 3: Order / Pick data (for velocity)
// Step 4: Inventory data (current stock)
// Outputs: SKU slotting, rack recommendations, warehouse sizing, SVG floor plan
import { useState } from 'react';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import { S } from '../components/styles.jsx';

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const BIN_CATALOG = {
  XS: { name:'Compartment Tray', dims:'300×200×100 mm', volCm3:6000,  fill:0.55, slotH:0.12 },
  S:  { name:'Small Tote/Bin',   dims:'400×300×200 mm', volCm3:24000, fill:0.55, slotH:0.22 },
  M:  { name:'Louvre/Shelf Bin', dims:'600×400×300 mm', volCm3:72000, fill:0.55, slotH:0.32 },
  L:  { name:'Stack Crate/Half-Pallet', dims:'800×600×400 mm', volCm3:192000, fill:0.55, slotH:0.45 },
  XL: { name:'Standard Pallet',  dims:'1200×1000 mm',  volCm3:1440000,fill:0.55, slotH:1.20 },
  LONG:{ name:'Long-Goods Slot', dims:'per item',       volCm3:null,  fill:0.40, slotH:0.40 },
};

const ZONE_DEFS = {
  golden: { label:'Golden Zone',   desc:'VF & F movers — nearest to dispatch', color:'#dcfce7', border:'#16a34a', textColor:'#166534', velocities:['VF','F'] },
  mid:    { label:'Mid-Level',     desc:'M movers — mid-warehouse',             color:'#fef9c3', border:'#ca8a04', textColor:'#854d0e', velocities:['M'] },
  reserve:{ label:'Reserve/Slow',  desc:'S movers — upper/back racking',        color:'#fff7ed', border:'#ea580c', textColor:'#9a3412', velocities:['S'] },
  bulk:   { label:'Bulk/Overflow', desc:'VS & NM — high-density back storage',  color:'#f1f5f9', border:'#64748b', textColor:'#374151', velocities:['VS','NM'] },
  long:   { label:'Long-Goods',    desc:'Awkward items — cantilever rack',       color:'#fdf4ff', border:'#9333ea', textColor:'#6b21a8', velocities:[] },
};

const RACK_DEFS = {
  shelving:   { name:'Shelving Rack',         bayW:1.0, bayD:0.6,  desc:'Totes/bins — manual pick' },
  selective:  { name:'Selective Pallet Rack', bayW:2.7, bayD:1.1,  desc:'Full access every pallet' },
  driveIn:    { name:'Drive-In Rack',         bayW:2.7, bayD:6.6,  desc:'6-deep, LIFO, high density' },
  doubleDeep: { name:'Double-Deep Rack',      bayW:2.7, bayD:2.4,  desc:'2-deep, reach truck' },
  cantilever: { name:'Cantilever Rack',       bayW:1.5, bayD:2.5,  desc:'Long/awkward items' },
  liveStorage:{ name:'Carton Live Storage',   bayW:1.0, bayD:3.0,  desc:'FIFO, high-turn small items' },
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function parseTSV(text) {
  return text.trim().split('\n')
    .map(r => r.split('\t').map(c => c.trim()))
    .filter(r => r.some(c => c));
}
function isHeaderRow(row) {
  const n1 = parseFloat(row[1]), n2 = parseFloat(row[2]);
  return isNaN(n1) && isNaN(n2);
}
function sizeBand(volCm3) {
  if (volCm3 <= 500)   return 'XS';
  if (volCm3 <= 3000)  return 'S';
  if (volCm3 <= 15000) return 'M';
  if (volCm3 <= 50000) return 'L';
  return 'XL';
}
function unitsPerBin(skuVolCm3, band) {
  const b = BIN_CATALOG[band];
  if (!b || !b.volCm3 || !skuVolCm3) return 1;
  return Math.max(1, Math.floor(b.volCm3 * b.fill / skuVolCm3));
}

// ─── RACK SELECTION ───────────────────────────────────────────────────────────
function selectRackType(sb, vb, isLong, clearH, forkType) {
  if (isLong) return 'cantilever';
  if (['XS','S'].includes(sb)) {
    if (['VF','F'].includes(vb)) return 'liveStorage';
    return 'shelving';
  }
  if (sb === 'M') return 'shelving';
  if (['L','XL'].includes(sb)) {
    if (['VF','F','M'].includes(vb)) return 'selective';
    if (vb === 'S') return clearH >= 9 && forkType !== 'manual' ? 'doubleDeep' : 'selective';
    return clearH >= 9 && forkType !== 'manual' ? 'driveIn' : 'selective';
  }
  return 'shelving';
}

// ─── VELOCITY CLASSIFICATION ──────────────────────────────────────────────────
function classifyVelocity(items) {
  const sorted = [...items].sort((a,b) => b.pickLines - a.pickLines);
  const totLines = sorted.reduce((s,r) => s+r.pickLines, 0);
  let cum = 0;
  const thresholds = { VF:0.50, F:0.75, M:0.90, S:0.98, VS:1.00 };
  const result = {};
  sorted.forEach(r => {
    if (r.pickLines === 0) { result[r.sku] = 'NM'; return; }
    cum += r.pickLines;
    const pct = cum / totLines;
    if (pct <= thresholds.VF)      result[r.sku] = 'VF';
    else if (pct <= thresholds.F)  result[r.sku] = 'F';
    else if (pct <= thresholds.M)  result[r.sku] = 'M';
    else if (pct <= thresholds.S)  result[r.sku] = 'S';
    else                            result[r.sku] = 'VS';
  });
  return result;
}

// ─── ZONE ASSIGNMENT ──────────────────────────────────────────────────────────
function getZone(vb, isLong) {
  if (isLong)                        return 'long';
  if (['VF','F'].includes(vb))       return 'golden';
  if (vb === 'M')                    return 'mid';
  if (vb === 'S')                    return 'reserve';
  return 'bulk';
}

// ─── MAIN ANALYSIS ────────────────────────────────────────────────────────────
function runAnalysis(masterRows, orderRows, inventoryRows, params) {
  const { clearH, forkType } = params;

  // Parse master SKU
  const master = {};
  masterRows.forEach(r => {
    const sku = r[0]; if (!sku) return;
    const L = parseFloat(r[1])||0, W = parseFloat(r[2])||0, H = parseFloat(r[3])||0;
    const volCm3 = +(L*W*H/1000).toFixed(2); // mm³ → cm³
    const maxDim = Math.max(L,W,H);
    master[sku] = { L,W,H, volCm3, maxDim, isLong: maxDim > 600,
      sb: sizeBand(volCm3) };
  });

  // Parse order data → pick lines per SKU
  const pickMap = {};
  orderRows.forEach(r => {
    const sku = r[2] || r[0]; if (!sku) return;
    pickMap[sku] = (pickMap[sku]||0) + 1;
  });

  // Parse inventory
  const invMap = {};
  inventoryRows.forEach(r => {
    const sku = r[0]; if (!sku) return;
    invMap[sku] = parseFloat(r[1])||0;
  });

  // All SKUs = union of master + pick + inventory
  const allSkus = new Set([...Object.keys(master), ...Object.keys(pickMap), ...Object.keys(invMap)]);

  // Prepare items for velocity classification
  const items = [...allSkus].map(sku => ({
    sku, pickLines: pickMap[sku]||0 }));
  const velocityMap = classifyVelocity(items);

  // Build per-SKU slotting data
  const slotted = [];
  allSkus.forEach(sku => {
    const m  = master[sku] || { L:0,W:0,H:0, volCm3:0, maxDim:0, isLong:false, sb:'S' };
    const vb = velocityMap[sku] || 'NM';
    const sb = m.sb;
    const isLong = m.isLong;
    const zone  = getZone(vb, isLong);
    const rack  = selectRackType(sb, vb, isLong, clearH, forkType);
    const bin   = isLong ? 'LONG' : sb;
    const upb   = isLong ? 1 : unitsPerBin(m.volCm3, bin);
    const stock = invMap[sku]||0;
    const locsReq = stock > 0 ? Math.max(1, Math.ceil(stock/upb)) : 0;
    const pl    = pickMap[sku]||0;
    slotted.push({ sku, ...m, vb, sb, isLong, zone, rack, bin,
      upb, stock, locsReq, pickLines:pl,
      binName: BIN_CATALOG[bin]?.name || '—',
      rackName: RACK_DEFS[rack]?.name || '—',
      zoneName: ZONE_DEFS[zone]?.label || '—' });
  });

  // Velocity × Size matrix (locations)
  const vbList = ['VF','F','M','S','VS','NM'];
  const sbList = ['XS','S','M','L','XL'];
  const matrix = {};
  vbList.forEach(v => sbList.forEach(s => { matrix[`${v}-${s}`] = 0; }));
  slotted.forEach(r => {
    const k = `${r.vb}-${r.sb}`;
    if (matrix[k] !== undefined) matrix[k] += r.locsReq;
  });

  // Zone summary
  const zoneSummary = {};
  Object.keys(ZONE_DEFS).forEach(z => {
    const rows = slotted.filter(r => r.zone === z);
    zoneSummary[z] = {
      skus: rows.length,
      locs: rows.reduce((s,r)=>s+r.locsReq,0),
      stock: rows.reduce((s,r)=>s+r.stock,0),
      pickLines: rows.reduce((s,r)=>s+r.pickLines,0),
    };
  });

  // Rack type summary
  const rackSummary = {};
  slotted.forEach(r => {
    if (!rackSummary[r.rack]) rackSummary[r.rack] = { locs:0, skus:0 };
    rackSummary[r.rack].locs += r.locsReq;
    rackSummary[r.rack].skus += 1;
  });

  // Headline metrics
  const totSKUs    = slotted.length;
  const totLocs    = slotted.reduce((s,r)=>s+r.locsReq,0);
  const totStock   = slotted.reduce((s,r)=>s+r.stock,0);
  const longCount  = slotted.filter(r=>r.isLong).length;
  const nmCount    = slotted.filter(r=>r.vb==='NM'&&r.stock>0).length;
  const nmStock    = slotted.filter(r=>r.vb==='NM').reduce((s,r)=>s+r.stock,0);

  return { slotted, matrix, zoneSummary, rackSummary,
    metrics: { totSKUs, totLocs, totStock, longCount, nmCount, nmStock } };
}

// ─── WAREHOUSE SIZING ─────────────────────────────────────────────────────────
function calcWarehouseSize(analysis, params) {
  const { clearH, forkType, dockCount, dockSide, aisleW, shifts } = params;
  const { zoneSummary, rackSummary } = analysis;

  // Pallet rack levels based on forklift & height
  const palletLevelH = 1.5; // m per level (pallet + beam)
  const maxLift = { manual:2.2, counterbalance:6.0, reach:9.0, vna:12.0 };
  const liftH = maxLift[forkType]||6.0;
  const palletLevels = Math.max(1, Math.floor((Math.min(liftH,clearH) - 0.8) / palletLevelH));

  // Shelf levels for bins
  const shelfSlotH = 0.35; // m per shelf level (avg)
  const shelfLevels = Math.max(1, Math.floor((Math.min(3.5, clearH) - 0.3) / shelfSlotH));

  // Area calculation per rack type
  const AISLE_FACTOR = 1 + (parseFloat(aisleW)||3.0) / 3.0; // aisle adds ~100% of rack depth
  const rackAreas = {};

  Object.entries(RACK_DEFS).forEach(([rk, rd]) => {
    const locs = rackSummary[rk]?.locs || 0;
    if (locs === 0) { rackAreas[rk] = 0; return; }
    let locsPerBay;
    if (['selective','driveIn','doubleDeep'].includes(rk)) {
      const depth = rk==='driveIn' ? 6 : rk==='doubleDeep' ? 2 : 1;
      locsPerBay = 2 * palletLevels * depth; // 2 pallets wide per bay
    } else if (rk === 'shelving' || rk === 'liveStorage') {
      locsPerBay = Math.floor(1.0/(shelfSlotH)) * shelfLevels; // approx bins per bay
    } else if (rk === 'cantilever') {
      locsPerBay = 8; // ~8 slots per cantilever bay
    } else {
      locsPerBay = 4;
    }
    const bays = Math.ceil(locs / Math.max(1,locsPerBay));
    const bayArea = rd.bayW * rd.bayD;
    rackAreas[rk] = +(bays * bayArea * AISLE_FACTOR).toFixed(1);
  });

  const netRackArea   = Object.values(rackAreas).reduce((s,v)=>s+v,0);
  const receivingArea = Math.max(50, netRackArea * 0.10);
  const dispatchArea  = Math.max(50, netRackArea * 0.10);
  const officeArea    = 50;
  const circulationArea = netRackArea * 0.08;
  const totalGrossArea = netRackArea + receivingArea + dispatchArea + officeArea + circulationArea;

  // Dock width requirement
  const minDockWidth = dockCount * 4.5 + 6; // each door 3.5m + spacing
  const recWidth = dockSide==='both' ? Math.sqrt(totalGrossArea * 0.8) : Math.sqrt(totalGrossArea * 0.6);
  const wW = Math.max(minDockWidth, Math.ceil(recWidth / 5) * 5);
  const wL = Math.ceil(totalGrossArea / wW / 5) * 5;

  // Zone areas (proportional to locations)
  const totalLocs = analysis.metrics.totLocs || 1;
  const zoneAreas = {};
  Object.entries(analysis.zoneSummary).forEach(([z, zs]) => {
    zoneAreas[z] = +(netRackArea * (zs.locs / totalLocs)).toFixed(1);
  });

  return { wW, wL, totalGrossArea: +totalGrossArea.toFixed(0),
    netRackArea: +netRackArea.toFixed(0), receivingArea: +receivingArea.toFixed(0),
    dispatchArea: +dispatchArea.toFixed(0), officeArea,
    rackAreas, zoneAreas, palletLevels, shelfLevels };
}

// ─── SVG FLOOR PLAN ───────────────────────────────────────────────────────────
function FloorPlanSVG({ analysis, design, params }) {
  const { wW, wL, zoneAreas, receivingArea, dispatchArea } = design;
  const { dockCount, dockSide } = params;

  const SVG_W = 780, SVG_H = 520;
  const MARGIN = 36;
  const scaleX = (SVG_W - MARGIN*2) / wW;
  const scaleY = (SVG_H - MARGIN*2 - 40) / wL;

  const toX = m => MARGIN + m * scaleX;
  const toY = m => MARGIN + m * scaleY;
  const toPx = (mW, mH) => ({ w: mW * scaleX, h: mH * scaleY });

  // Zone layout — horizontal bands from south (bottom = docks)
  const totalLocs = analysis.metrics.totLocs || 1;
  const zoneOrder = ['golden','mid','reserve','bulk','long'];
  const zoneHeights = {};
  zoneOrder.forEach(z => {
    const a = zoneAreas[z]||0;
    zoneHeights[z] = (a / (wW||1));
  });

  // From south: receiving → golden → mid → reserve → bulk/long → dispatch strip on north
  const recH  = (receivingArea||0) / (wW||1);
  const disH  = (dispatchArea||0) / (wW||1);
  const availH = Math.max(0, wL - recH - disH);

  const totZoneH = zoneOrder.reduce((s,z)=>s+(zoneHeights[z]||0),0)||1;
  let cursor = wL; // start from south (bottom SVG)
  const zoneRects = [];

  // Receiving at south
  cursor -= recH;
  zoneRects.push({ key:'receiving', y:cursor, h:recH, label:'Receiving / GRN',
    color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });

  // Zones (golden first = closest to receiving/dispatch)
  zoneOrder.forEach(z => {
    const zh = (zoneHeights[z]||0) / totZoneH * availH;
    cursor -= zh;
    zoneRects.push({ key:z, y:cursor, h:zh,
      label: ZONE_DEFS[z].label,
      color: ZONE_DEFS[z].color,
      border: ZONE_DEFS[z].border,
      text:  ZONE_DEFS[z].textColor });
  });

  // Dispatch at north
  cursor -= disH;
  zoneRects.push({ key:'dispatch', y:cursor, h:disH, label:'Dispatch / Packing',
    color:'#fef3c7', border:'#d97706', text:'#92400e' });

  // Rack rows inside zones
  const rackRows = [];
  zoneRects.filter(z=>!['receiving','dispatch'].includes(z.key)).forEach(zone => {
    const zH = zone.h;
    if (zH < 1) return;
    const aisleW = parseFloat(params.aisleW)||3.0;
    const rowDepth = 0.9; // m
    const slot = (rowDepth + aisleW);
    const numRows = Math.max(1, Math.floor(zH / slot));
    for (let i=0; i<numRows; i++) {
      const rowY = zone.y + (zH / numRows) * i + aisleW/2;
      rackRows.push({ y: rowY, zoneKey: zone.key });
    }
  });

  // Dock doors
  const dockDoors = [];
  const doorW = 3.5, doorSpacing = wW / (dockCount+1);
  for (let i=1; i<=dockCount; i++) {
    dockDoors.push({ x: doorSpacing * i - doorW/2, w: doorW });
  }

  return (
    <svg width={SVG_W} height={SVG_H} viewBox={`0 0 ${SVG_W} ${SVG_H}`}
      style={{ border:'1px solid #e2e8f0', borderRadius:'8px', background:'#fafafa', width:'100%', height:'auto' }}>

      {/* Warehouse outline */}
      <rect x={toX(0)} y={toY(0)} width={wW*scaleX} height={wL*scaleY}
        fill="white" stroke="#334155" strokeWidth="2"/>

      {/* Zone bands */}
      {zoneRects.map(z => (
        <g key={z.key}>
          <rect x={toX(0)} y={toY(z.y)} width={wW*scaleX} height={Math.max(2, z.h*scaleY)}
            fill={z.color} stroke={z.border} strokeWidth="1" opacity="0.9"/>
          {z.h * scaleY > 16 && (
            <text x={toX(wW/2)} y={toY(z.y + z.h/2)} textAnchor="middle"
              dominantBaseline="middle" fontSize="11" fontWeight="600" fill={z.text}>
              {z.label}
              {z.h > 5 && ` (${z.h.toFixed(0)}m)`}
            </text>
          )}
        </g>
      ))}

      {/* Rack rows */}
      {rackRows.map((r,i) => (
        <rect key={i} x={toX(1)} y={toY(r.y)} width={(wW-2)*scaleX} height={Math.max(2,0.6*scaleY)}
          fill="rgba(51,65,85,0.15)" rx="1"/>
      ))}

      {/* Dock doors */}
      {dockDoors.map((d,i) => (
        <g key={i}>
          <rect x={toX(d.x)} y={toY(wL)-4} width={d.w*scaleX} height={8}
            fill="#1d4ed8" rx="2"/>
          <text x={toX(d.x + d.w/2)} y={toY(wL)+14} textAnchor="middle"
            fontSize="9" fill="#1d4ed8" fontWeight="600">D{i+1}</text>
        </g>
      ))}

      {/* Compass / North arrow */}
      <text x={SVG_W-MARGIN+6} y={MARGIN+10} fontSize="11" fill="#64748b" fontWeight="700">N↑</text>

      {/* Scale bar */}
      <line x1={toX(0)} y1={SVG_H-18} x2={toX(Math.min(10,wW))} y2={SVG_H-18}
        stroke="#64748b" strokeWidth="2"/>
      <text x={toX(Math.min(10,wW)/2)} y={SVG_H-6} textAnchor="middle"
        fontSize="9" fill="#64748b">{Math.min(10,wW)}m</text>

      {/* Dimensions */}
      <text x={toX(wW/2)} y={MARGIN-10} textAnchor="middle"
        fontSize="11" fontWeight="700" fill="#0f172a">{wW}m</text>
      <text x={MARGIN-12} y={toY(wL/2)} textAnchor="middle"
        fontSize="11" fontWeight="700" fill="#0f172a"
        transform={`rotate(-90,${MARGIN-12},${toY(wL/2)})`}>{wL}m</text>

      {/* Total area label */}
      <text x={toX(wW/2)} y={SVG_H-4} textAnchor="middle"
        fontSize="10" fill="#64748b">
        Recommended warehouse: {wW}m × {wL}m = {(wW*wL).toLocaleString()} m²
      </text>
    </svg>
  );
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function exportExcel(analysis, design, params) {
  const wb   = XLSX.utils.book_new();
  const today= new Date().toLocaleDateString();
  const ws   = (data,cols) => {
    const s = XLSX.utils.aoa_to_sheet(data);
    if (cols) s['!cols'] = cols.map(w=>({wch:w}));
    return s;
  };
  const { slotted, metrics, zoneSummary, rackSummary, matrix } = analysis;
  const { wW, wL, totalGrossArea, netRackArea, rackAreas, palletLevels, shelfLevels } = design;

  // Sheet 1: Summary
  XLSX.utils.book_append_sheet(wb, ws([
    ['WAREHOUSE STORAGE DESIGN REPORT'],['Generated:',today],[],
    ['HEADLINE METRICS'],
    ['Total Active SKUs',metrics.totSKUs],
    ['Total Current Stock (units)',metrics.totStock],
    ['Total Storage Locations Required',metrics.totLocs],
    ['Long/Awkward Items',metrics.longCount],
    ['No-Movement SKUs (in stock)',metrics.nmCount],
    ['No-Movement Units',metrics.nmStock],[],
    ['WAREHOUSE SIZE RECOMMENDATION'],
    ['Recommended Width (m)',wW],
    ['Recommended Length (m)',wL],
    ['Total Gross Floor Area (m²)',wW*wL],
    ['Net Racking Area (m²)',netRackArea],
    ['Pallet Rack Levels',palletLevels],
    ['Shelf Rack Levels',shelfLevels],[],
    ['ZONE BREAKDOWN'],
    ['Zone','SKUs','Locations','Stock Units','Pick Lines'],
    ...Object.entries(zoneSummary).map(([z,v])=>[
      ZONE_DEFS[z]?.label||z, v.skus, v.locs, v.stock, v.pickLines]),
    [],[  'RACK TYPE SUMMARY'],
    ['Rack Type','SKUs','Locations Required','Est. Floor Area (m²)'],
    ...Object.entries(rackSummary).map(([rk,rv])=>[
      RACK_DEFS[rk]?.name||rk, rv.skus, rv.locs, rackAreas[rk]||0]),
  ],[28,18,18,18,16]),'1. Design Summary');

  // Sheet 2: Velocity × Size Matrix
  const vbList=['VF','F','M','S','VS','NM'], sbList=['XS','S','M','L','XL'];
  XLSX.utils.book_append_sheet(wb, ws([
    ['VELOCITY × SIZE MATRIX (Storage Locations Required)'],[],
    ['Velocity \\ Size',...sbList,'Total'],
    ...vbList.map(v=>[v,...sbList.map(s=>matrix[`${v}-${s}`]||0),
      sbList.reduce((sum,s)=>sum+(matrix[`${v}-${s}`]||0),0)]),
    ['Total',...sbList.map(s=>vbList.reduce((sum,v)=>sum+(matrix[`${v}-${s}`]||0),0)),
      metrics.totLocs],
  ],[18,10,10,10,10,10,10]),'2. Velocity×Size Matrix');

  // Sheet 3: SKU Slotting Detail
  XLSX.utils.book_append_sheet(wb, ws([
    ['SKU SLOTTING DETAIL'],[],
    ['SKU Code','L (mm)','W (mm)','H (mm)','Vol (cm³)','Max Dim','Long?',
     'Pick Lines','Velocity Band','Size Band','Combined Band',
     'Bin/Container','Location Size (mm)','Rack Type','Storage Zone','Units/Bin','Stock','Locs Required'],
    ...slotted.map(r=>[r.sku,r.L,r.W,r.H,r.volCm3.toFixed(0),r.maxDim,r.isLong?'YES':'',
      r.pickLines,r.vb,r.sb,`${r.vb}-${r.sb}`,r.binName,
      BIN_CATALOG[r.bin]?.dims||'—',
      r.rackName,r.zoneName,
      r.upb,r.stock,r.locsReq]),
  ],[22,8,8,8,10,10,8,12,14,10,14,22,20,24,34,10,10,14]),'3. SKU Slotting');

  // Sheet 4: Rack Schedule
  XLSX.utils.book_append_sheet(wb, ws([
    ['RACK SCHEDULE'],[],
    ['Rack Type','Description','Locations Needed','Est. Floor Area (m²)','Bay W (m)','Bay D (m)','No. of Bays'],
    ...Object.entries(rackSummary).map(([rk,rv])=>{
      const rd = RACK_DEFS[rk]||{};
      const area = rackAreas[rk]||0;
      const bayArea = (rd.bayW||1)*(rd.bayD||1);
      const bays = bayArea>0?Math.ceil(area/bayArea):0;
      return[rd.name||rk,rd.desc||'',rv.locs,area,rd.bayW,rd.bayD,bays];
    }),
  ],[26,34,16,18,10,10,14]),'4. Rack Schedule');

  XLSX.writeFile(wb,`Warehouse_Design_${today.replace(/\//g,'-')}.xlsx`);
}

// ─── PPT EXPORT ───────────────────────────────────────────────────────────────
function exportPPT(analysis, design, params) {
  const prs  = new PptxGenJS();
  const today= new Date().toLocaleDateString();
  const PINK ='BE185D', DARK='0F172A', WHITE='FFFFFF', GRAY='64748B',
        GREEN='166534', BLUE='1D4ED8', AMBER='D97706';

  prs.layout = 'LAYOUT_WIDE';
  const hdr = (sld, title, sub) => {
    sld.addShape(prs.ShapeType.rect,{x:0,y:0,w:'100%',h:1.1,fill:{color:PINK}});
    sld.addText('DensiCube — Warehouse Designer',{x:0.3,y:0.06,w:8,h:0.3,fontSize:9,color:WHITE,fontFace:'Calibri'});
    sld.addText(title,{x:0.3,y:0.3,w:9,h:0.55,fontSize:20,color:WHITE,bold:true,fontFace:'Calibri'});
    if(sub) sld.addText(sub,{x:0.3,y:0.82,w:9,h:0.25,fontSize:10,color:'FBCFE8',fontFace:'Calibri'});
    sld.addText(today,{x:9.1,y:0.06,w:2,h:0.3,fontSize:9,color:'FBCFE8',align:'right',fontFace:'Calibri'});
  };

  // Slide 1: Title
  const s1=prs.addSlide();
  s1.addShape(prs.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:DARK}});
  s1.addShape(prs.ShapeType.rect,{x:0,y:0,w:0.18,h:'100%',fill:{color:PINK}});
  s1.addText('DensiCube',{x:0.5,y:1.2,w:9,h:0.5,fontSize:13,color:'FBCFE8',bold:true,fontFace:'Calibri'});
  s1.addText('Warehouse Storage Design Report',{x:0.5,y:1.8,w:9,h:1.0,fontSize:32,color:WHITE,bold:true,fontFace:'Calibri'});
  s1.addText(`${analysis.metrics.totSKUs.toLocaleString()} SKUs · ${analysis.metrics.totLocs.toLocaleString()} locations · ${design.wW}×${design.wL}m recommended`,
    {x:0.5,y:3.0,w:9,h:0.4,fontSize:13,color:'94A3B8',fontFace:'Calibri'});
  s1.addText(`Generated: ${today}`,{x:0.5,y:5.8,w:9,h:0.3,fontSize:10,color:'475569',fontFace:'Calibri'});

  // Slide 2: Key Metrics
  const s2=prs.addSlide(); hdr(s2,'Key Metrics','Headline numbers from SKU slotting analysis');
  const mStats=[
    [analysis.metrics.totSKUs.toLocaleString(),'Total SKUs',PINK],
    [analysis.metrics.totLocs.toLocaleString(),'Locations Required',BLUE],
    [analysis.metrics.totStock.toLocaleString(),'Current Stock Units',GREEN],
    [analysis.metrics.longCount,'Long/Awkward Items',AMBER],
    [`${design.wW}×${design.wL}m`,'Recommended Size',DARK],
    [design.totalGrossArea.toLocaleString()+'m²','Gross Floor Area',GRAY],
  ];
  mStats.forEach(([v,l,c],i)=>{
    const x=0.4+(i%3)*3.2, y=i<3?1.3:2.8;
    s2.addShape(prs.ShapeType.roundRect,{x,y,w:3.0,h:1.1,fill:{color:'F8FAFC'},line:{color:'E2E8F0',pt:1},rectRadius:0.08});
    s2.addText(v,{x,y:y+0.08,w:3.0,h:0.55,fontSize:22,bold:true,color:c,align:'center',fontFace:'Calibri'});
    s2.addText(l,{x,y:y+0.66,w:3.0,h:0.3,fontSize:9,color:GRAY,align:'center',fontFace:'Calibri'});
  });

  // Slide 3: Velocity × Size Matrix
  const s3=prs.addSlide(); hdr(s3,'Velocity × Size Matrix','Storage locations required by movement speed and item size');
  const vbList=['VF','F','M','S','VS','NM'], sbList=['XS','S','M','L','XL'];
  const matHdr=[{text:'Velocity \\ Size',options:{bold:true,color:WHITE,fill:{color:DARK}}},
    ...sbList.map(s=>({text:s,options:{bold:true,color:WHITE,fill:{color:DARK},align:'center'}})),
    {text:'Total',options:{bold:true,color:WHITE,fill:{color:DARK},align:'center'}}];
  const matRows=vbList.map(v=>[
    {text:v,options:{bold:true,color:PINK,fill:{color:'F8FAFC'}}},
    ...sbList.map(s=>{const n=analysis.matrix[`${v}-${s}`]||0;return{text:n?String(n):'—',options:{align:'center',color:n?DARK:'9CA3AF'}};}),
    {text:String(sbList.reduce((sum,s)=>sum+(analysis.matrix[`${v}-${s}`]||0),0)),options:{align:'center',bold:true}},
  ]);
  s3.addTable([matHdr,...matRows],{x:0.5,y:1.3,w:9.1,colW:[1.8,1.2,1.2,1.2,1.2,1.2,1.3],
    fontSize:11,border:{type:'solid',color:'E2E8F0',pt:1},rowH:0.45,autoPage:false});

  // Slide 4: Zone Breakdown
  const s4=prs.addSlide(); hdr(s4,'Zone Layout Plan','Storage zones by velocity — drives warehouse layout');
  const zRows=Object.entries(analysis.zoneSummary).map(([z,v])=>[
    {text:ZONE_DEFS[z]?.label||z,options:{bold:true,color:DARK}},
    {text:ZONE_DEFS[z]?.desc||'',options:{color:GRAY,fontSize:9}},
    {text:String(v.skus),options:{align:'center'}},
    {text:v.locs.toLocaleString(),options:{align:'center',bold:true}},
    {text:v.pickLines.toLocaleString(),options:{align:'center'}},
  ]);
  s4.addTable([
    [{text:'Zone',options:{bold:true,color:WHITE,fill:{color:PINK}}},
     {text:'Description',options:{bold:true,color:WHITE,fill:{color:PINK}}},
     {text:'SKUs',options:{bold:true,color:WHITE,fill:{color:PINK},align:'center'}},
     {text:'Locations',options:{bold:true,color:WHITE,fill:{color:PINK},align:'center'}},
     {text:'Pick Lines',options:{bold:true,color:WHITE,fill:{color:PINK},align:'center'}}],
    ...zRows,
  ],{x:0.4,y:1.3,w:9.2,colW:[2.0,3.5,1.0,1.4,1.3],
    fontSize:11,border:{type:'solid',color:'E2E8F0',pt:1},rowH:0.5,autoPage:false});

  // Slide 5: Rack Schedule
  const s5=prs.addSlide(); hdr(s5,'Rack Type Recommendations','Storage media selected per SKU velocity & size combination');
  const rackRows2=Object.entries(analysis.rackSummary).map(([rk,rv])=>[
    {text:RACK_DEFS[rk]?.name||rk,options:{bold:true,color:DARK}},
    {text:RACK_DEFS[rk]?.desc||'',options:{color:GRAY,fontSize:9}},
    {text:rv.skus.toLocaleString(),options:{align:'center'}},
    {text:rv.locs.toLocaleString(),options:{align:'center',bold:true}},
    {text:String((design.rackAreas[rk]||0).toFixed(0))+'m²',options:{align:'center'}},
  ]);
  s5.addTable([
    [{text:'Rack Type',options:{bold:true,color:WHITE,fill:{color:PINK}}},
     {text:'Description',options:{bold:true,color:WHITE,fill:{color:PINK}}},
     {text:'SKUs',options:{bold:true,color:WHITE,fill:{color:PINK},align:'center'}},
     {text:'Locations',options:{bold:true,color:WHITE,fill:{color:PINK},align:'center'}},
     {text:'Floor Area',options:{bold:true,color:WHITE,fill:{color:PINK},align:'center'}}],
    ...rackRows2,
  ],{x:0.4,y:1.3,w:9.2,colW:[2.4,3.1,1.0,1.4,1.3],
    fontSize:11,border:{type:'solid',color:'E2E8F0',pt:1},rowH:0.5,autoPage:false});

  // Slide 6: Warehouse Size
  const s6=prs.addSlide(); hdr(s6,'Warehouse Size Recommendation','Based on SKU slotting, rack types and aisle requirements');
  const sizeData=[
    ['Net Racking Area',design.netRackArea+'m²'],
    ['Receiving Area',design.receivingArea+'m²'],
    ['Dispatch / Packing Area',design.dispatchArea+'m²'],
    ['Office / Welfare',design.officeArea+'m²'],
    ['Total Gross Floor Area',design.totalGrossArea+'m²'],
    ['Recommended Dimensions',`${design.wW}m × ${design.wL}m`],
    ['Clear Height Input',`${params.clearH}m`],
    ['Pallet Rack Levels',design.palletLevels],
    ['Shelf Rack Levels',design.shelfLevels],
  ];
  sizeData.forEach(([l,v],i)=>{
    const y=1.4+i*0.48;
    s6.addShape(prs.ShapeType.rect,{x:0.4,y,w:6.5,h:0.42,fill:{color:i%2===0?'F8FAFC':WHITE},line:{color:'E2E8F0',pt:0}});
    s6.addText(l,{x:0.5,y:y+0.06,w:4.5,h:0.3,fontSize:11,color:DARK,fontFace:'Calibri'});
    s6.addText(String(v),{x:5.0,y:y+0.06,w:2.0,h:0.3,fontSize:11,bold:true,color:PINK,align:'right',fontFace:'Calibri'});
  });
  s6.addText(`Recommendation: A ${design.wW}×${design.wL}m (${design.totalGrossArea}m²) warehouse with ${design.palletLevels} pallet levels`,
    {x:0.4,y:6.0,w:9.2,h:0.4,fontSize:10,color:GRAY,italic:true,fontFace:'Calibri'});

  prs.writeFile({fileName:`Warehouse_Design_${today.replace(/\//g,'-')}.pptx`});
}

// ─── COMPONENT ────────────────────────────────────────────────────────────────
export default function WarehouseDesignerTool() {
  // Params
  const [clearH,  setClearH]  = useState('9');
  const [dockCount,setDockCount]=useState('4');
  const [dockSide, setDockSide]= useState('one');
  const [forkType, setForkType]= useState('reach');
  const [aisleW,   setAisleW]  = useState('3.0');
  const [shifts,   setShifts]  = useState('1');
  const [peakOrders,setPeakOrders]=useState('');

  // Data
  const [masterText, setMasterText] = useState('');
  const [orderText,  setOrderText]  = useState('');
  const [invText,    setInvText]    = useState('');

  // Results
  const [analysis,  setAnalysis]  = useState(null);
  const [design,    setDesign]    = useState(null);
  const [loading,   setLoading]   = useState(false);
  const [error,     setError]     = useState('');

  const params = { clearH:parseFloat(clearH)||9, dockCount:parseInt(dockCount)||4,
    dockSide, forkType, aisleW:parseFloat(aisleW)||3.0, shifts:parseInt(shifts)||1,
    peakOrders:parseInt(peakOrders)||0 };

  const runAll = () => {
    setError(''); setLoading(true);
    setTimeout(() => {
      try {
        if (!masterText.trim()) throw new Error('Paste Master SKU data first.');
        const masterRows = parseTSV(masterText);
        const mData = isHeaderRow(masterRows[0]) ? masterRows.slice(1) : masterRows;
        const orderRows  = orderText.trim()  ? parseTSV(orderText).filter(r=>r[2]||r[0])  : [];
        const oData  = orderRows.length && isHeaderRow(orderRows[0]) ? orderRows.slice(1) : orderRows;
        const invRows    = invText.trim()    ? parseTSV(invText).filter(r=>r[0])           : [];
        const iData  = invRows.length && isHeaderRow(invRows[0]) ? invRows.slice(1) : invRows;
        if (!mData.length) throw new Error('No valid SKU rows found in Master SKU data.');
        const a = runAnalysis(mData, oData, iData, params);
        const d = calcWarehouseSize(a, params);
        setAnalysis(a); setDesign(d);
      } catch(e) { setError(e.message); }
      setLoading(false);
    }, 100);
  };

  const inp = {...S.input, marginBottom:'4px'};
  const lbl = {...S.label};

  const colHint = cols => (
    <div style={{display:'flex',gap:'4px',flexWrap:'wrap',marginBottom:'8px'}}>
      {cols.map((col,i)=>(
        <span key={i} style={{background:'#f1f5f9',border:'1px solid #e2e8f0',borderRadius:'6px',
          padding:'3px 8px',fontSize:'12px',fontWeight:'600',color:'#475569',
          display:'flex',alignItems:'center',gap:'4px'}}>
          <span style={{background:'#be185d',color:'#fff',borderRadius:'50%',width:'15px',height:'15px',
            display:'inline-flex',alignItems:'center',justifyContent:'center',fontSize:'9px',fontWeight:'800',flexShrink:0}}>{i+1}</span>
          {col}
        </span>))}
    </div>
  );

  const stepCircle = (n, done) => (
    <div style={{width:'32px',height:'32px',borderRadius:'50%',flexShrink:0,
      background:done?'#166534':'#7c3aed',display:'flex',alignItems:'center',
      justifyContent:'center',color:'#fff',fontWeight:'800',fontSize:'14px'}}>
      {done?'✓':n}
    </div>
  );

  const textarea = (val, onChange, ph) => (
    <textarea value={val} onChange={e=>onChange(e.target.value)} placeholder={ph}
      style={{width:'100%',height:'120px',border:'1px solid #e2e8f0',borderRadius:'8px',
        padding:'10px 12px',fontSize:'12px',fontFamily:'monospace',resize:'vertical',
        outline:'none',boxSizing:'border-box',color:'#374151',lineHeight:'1.6'}}/>
  );

  return (
    <div>
      <div style={S.sectionDesc}>
        Design your warehouse storage layout from your SKU data. Enter warehouse parameters,
        paste your SKU master, order history, and inventory — the tool classifies every SKU
        by velocity and size, recommends rack types and zones, sizes your warehouse,
        and generates a visual floor plan.
      </div>

      <div style={{display:'grid',gridTemplateColumns:'340px 1fr',gap:'20px',alignItems:'start'}}>

        {/* ── LEFT PANEL ──────────────────────────────────────────────────────── */}
        <div>

          {/* Step 1: Parameters */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'16px'}}>
              {stepCircle(1, false)}
              <div style={S.cardTitle}>Warehouse Parameters</div>
            </div>
            <div style={S.grid2}>
              <div><label style={lbl}>Clear Height (m)</label>
                <input style={inp} type="number" min="4" max="20" step="0.5" value={clearH}
                  onChange={e=>setClearH(e.target.value)} placeholder="9"/></div>
              <div><label style={lbl}>Dock Doors</label>
                <input style={inp} type="number" min="1" max="20" value={dockCount}
                  onChange={e=>setDockCount(e.target.value)} placeholder="4"/></div>
              <div><label style={lbl}>Dock Position</label>
                <select style={inp} value={dockSide} onChange={e=>setDockSide(e.target.value)}>
                  <option value="one">One side</option>
                  <option value="both">Opposite sides</option>
                  <option value="corner">Corner</option>
                </select></div>
              <div><label style={lbl}>Forklift Type</label>
                <select style={inp} value={forkType} onChange={e=>setForkType(e.target.value)}>
                  <option value="manual">Manual pallet jack</option>
                  <option value="counterbalance">Counterbalance</option>
                  <option value="reach">Reach truck</option>
                  <option value="vna">VNA (Very Narrow Aisle)</option>
                </select></div>
              <div><label style={lbl}>Aisle Width (m)</label>
                <select style={inp} value={aisleW} onChange={e=>setAisleW(e.target.value)}>
                  <option value="2.0">2.0m — Narrow</option>
                  <option value="2.4">2.4m — Compact</option>
                  <option value="3.0">3.0m — Standard</option>
                  <option value="3.5">3.5m — Wide</option>
                </select></div>
              <div><label style={lbl}>Working Shifts</label>
                <select style={inp} value={shifts} onChange={e=>setShifts(e.target.value)}>
                  <option value="1">1 shift</option>
                  <option value="2">2 shifts</option>
                  <option value="3">3 shifts</option>
                </select></div>
            </div>
            <div><label style={lbl}>Peak Daily Orders (units, optional)</label>
              <input style={inp} type="number" min="0" value={peakOrders}
                onChange={e=>setPeakOrders(e.target.value)} placeholder="e.g. 500"/></div>
          </div>

          {/* Step 2: Master SKU */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'12px'}}>
              {stepCircle(2, !!masterText.trim())}
              <div style={S.cardTitle}>Master SKU Data</div>
            </div>
            {colHint(['SKU Code','Length (mm)','Width (mm)','Height (mm)','Weight (kg)'])}
            {textarea(masterText, setMasterText,
              'Paste SKU master data (Ctrl+V)\n\nExample:\nSKU-001\t300\t200\t150\t2.5\nSKU-002\t650\t80\t80\t1.2')}
          </div>

          {/* Step 3: Order / Pick Data */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'12px'}}>
              {stepCircle(3, !!orderText.trim())}
              <div>
                <div style={S.cardTitle}>Order / Pick Data <span style={{fontSize:'12px',fontWeight:'400',color:'#059669'}}>(Optional)</span></div>
                <div style={{fontSize:'12px',color:'#6b7280'}}>Used to classify SKU velocity bands</div>
              </div>
            </div>
            {colHint(['Order No','Dispatch Location','SKU Code','Qty','Date'])}
            {textarea(orderText, setOrderText,
              'Paste order data here — SKU pick frequency drives zone assignment\n\nWithout this data, all SKUs will be treated as equal velocity')}
          </div>

          {/* Step 4: Inventory */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'12px'}}>
              {stepCircle(4, !!invText.trim())}
              <div>
                <div style={S.cardTitle}>Current Inventory <span style={{fontSize:'12px',fontWeight:'400',color:'#059669'}}>(Optional)</span></div>
                <div style={{fontSize:'12px',color:'#6b7280'}}>Used to calculate storage locations needed</div>
              </div>
            </div>
            {colHint(['SKU Code','Current Stock Qty','Warehouse Location (opt)'])}
            {textarea(invText, setInvText,
              'Paste current inventory\n\nExample:\nSKU-001\t2500\nSKU-002\t180')}
          </div>

          {error && <div style={{...S.error,marginBottom:'12px'}}>⚠ {error}</div>}

          <button onClick={runAll} disabled={loading||!masterText.trim()}
            style={{width:'100%',padding:'13px',
              background:masterText.trim()&&!loading?'linear-gradient(135deg,#7c3aed,#6d28d9)':'#e2e8f0',
              color:masterText.trim()&&!loading?'#fff':'#9ca3af',
              border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'15px',
              cursor:masterText.trim()&&!loading?'pointer':'not-allowed',fontFamily:'inherit',
              boxShadow:masterText.trim()?'0 4px 14px rgba(124,58,237,0.35)':'none'}}>
            {loading?'⏳ Analysing...':'🏭 Generate Warehouse Design'}
          </button>
        </div>

        {/* ── RIGHT PANEL ──────────────────────────────────────────────────── */}
        <div>
          {!analysis && !loading && (
            <div style={{...S.card,padding:'60px',textAlign:'center',color:'#9ca3af'}}>
              <div style={{fontSize:'48px',marginBottom:'12px'}}>🏭</div>
              <div style={{fontWeight:'600',fontSize:'15px',color:'#374151',marginBottom:'6px'}}>
                Warehouse Storage Designer
              </div>
              <div style={{fontSize:'13px'}}>
                Fill in parameters and paste your SKU data to generate a complete warehouse design
              </div>
            </div>
          )}

          {analysis && design && (<>

            {/* Headline metrics */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'12px',marginBottom:'16px'}}>
              {[
                ['Total SKUs', analysis.metrics.totSKUs.toLocaleString(), '#eff6ff','#1d4ed8'],
                ['Locations Needed', analysis.metrics.totLocs.toLocaleString(), '#f5f3ff','#7c3aed'],
                ['Recommended Size', `${design.wW}×${design.wL}m`, '#f0fdf4','#166534'],
                ['Gross Area', `${(design.wW*design.wL).toLocaleString()}m²`, '#fef9c3','#854d0e'],
                ['Long/Awkward SKUs', analysis.metrics.longCount, '#fdf4ff','#9333ea'],
                ['No-Movement SKUs', analysis.metrics.nmCount, '#fff1f2','#be185d'],
              ].map(([l,v,bg,col])=>(
                <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',textAlign:'center',border:`1px solid ${col}22`}}>
                  <div style={{fontSize:'18px',fontWeight:'800',color:col}}>{v}</div>
                  <div style={{fontSize:'10px',color:'#6b7280',marginTop:'3px',fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                </div>))}
            </div>

            {/* Floor Plan SVG */}
            <div style={S.card}>
              <div style={{fontWeight:'700',fontSize:'14px',color:'#0f172a',marginBottom:'12px'}}>
                🗺 Recommended Floor Layout
              </div>
              <FloorPlanSVG analysis={analysis} design={design} params={params}/>
              {/* Legend */}
              <div style={{display:'flex',gap:'12px',flexWrap:'wrap',marginTop:'12px'}}>
                {Object.entries(ZONE_DEFS).map(([k,z])=>(
                  <div key={k} style={{display:'flex',alignItems:'center',gap:'5px',fontSize:'11px'}}>
                    <div style={{width:'14px',height:'14px',background:z.color,border:`1px solid ${z.border}`,borderRadius:'3px'}}/>
                    <span style={{color:z.textColor,fontWeight:'600'}}>{z.label}</span>
                  </div>
                ))}
                <div style={{display:'flex',alignItems:'center',gap:'5px',fontSize:'11px'}}>
                  <div style={{width:'14px',height:'14px',background:'#e0f2fe',border:'1px solid #0284c7',borderRadius:'3px'}}/>
                  <span style={{color:'#0369a1',fontWeight:'600'}}>Receiving</span>
                </div>
              </div>
            </div>

            {/* Zone breakdown */}
            <div style={{...S.card,padding:'0',overflow:'hidden',marginBottom:'12px'}}>
              <div style={{padding:'12px 18px',borderBottom:'1px solid #f1f5f9',fontWeight:'700',fontSize:'13px'}}>
                Zone Breakdown
              </div>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                <thead><tr>
                  {['Zone','SKUs','Locations','Stock Units','Pick Lines'].map(h=>(
                    <th key={h} style={{padding:'8px 12px',textAlign:'left',fontWeight:'600',
                      fontSize:'11px',color:'#6b7a8d',textTransform:'uppercase',
                      background:'#f8fafc',borderBottom:'1px solid #e8edf2'}}>{h}</th>))}
                </tr></thead>
                <tbody>
                  {Object.entries(analysis.zoneSummary).map(([z,v],i)=>(
                    <tr key={z} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                      <td style={{padding:'8px 12px'}}>
                        <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
                          <div style={{width:'10px',height:'10px',borderRadius:'50%',
                            background:ZONE_DEFS[z]?.border||'#ccc',flexShrink:0}}/>
                          <span style={{fontWeight:'600'}}>{ZONE_DEFS[z]?.label||z}</span>
                        </div>
                      </td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{v.skus.toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right',fontWeight:'700',color:'#7c3aed'}}>{v.locs.toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{v.stock.toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{v.pickLines.toLocaleString()}</td>
                    </tr>))}
                </tbody>
              </table>
            </div>

            {/* Rack schedule */}
            <div style={{...S.card,padding:'0',overflow:'hidden',marginBottom:'12px'}}>
              <div style={{padding:'12px 18px',borderBottom:'1px solid #f1f5f9',fontWeight:'700',fontSize:'13px'}}>
                Rack Type Schedule
              </div>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                <thead><tr>
                  {['Rack Type','Locations','Est. Area (m²)','Bay W×D'].map(h=>(
                    <th key={h} style={{padding:'8px 12px',textAlign:'left',fontWeight:'600',
                      fontSize:'11px',color:'#6b7a8d',textTransform:'uppercase',
                      background:'#f8fafc',borderBottom:'1px solid #e8edf2'}}>{h}</th>))}
                </tr></thead>
                <tbody>
                  {Object.entries(analysis.rackSummary).map(([rk,rv],i)=>(
                    <tr key={rk} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                      <td style={{padding:'8px 12px'}}>
                        <div style={{fontWeight:'600'}}>{RACK_DEFS[rk]?.name||rk}</div>
                        <div style={{fontSize:'11px',color:'#9ca3af'}}>{RACK_DEFS[rk]?.desc}</div>
                      </td>
                      <td style={{padding:'8px 12px',textAlign:'right',fontWeight:'700'}}>{rv.locs.toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{(design.rackAreas[rk]||0).toFixed(0)}</td>
                      <td style={{padding:'8px 12px',color:'#6b7280',fontSize:'11px'}}>
                        {RACK_DEFS[rk]?.bayW}m × {RACK_DEFS[rk]?.bayD}m
                      </td>
                    </tr>))}
                </tbody>
              </table>
            </div>

            {/* Velocity × Size matrix */}
            <div style={{...S.card,padding:'0',overflow:'hidden',marginBottom:'16px'}}>
              <div style={{padding:'12px 18px',borderBottom:'1px solid #f1f5f9',fontWeight:'700',fontSize:'13px'}}>
                Velocity × Size Matrix (Locations Required)
              </div>
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                  <thead><tr>
                    <th style={{padding:'7px 12px',background:'#f8fafc',borderBottom:'1px solid #e8edf2',
                      fontWeight:'700',color:'#374151',textAlign:'left'}}>Velocity \ Size</th>
                    {['XS','S','M','L','XL','Total'].map(h=>(
                      <th key={h} style={{padding:'7px 12px',background:'#f8fafc',
                        borderBottom:'1px solid #e8edf2',fontWeight:'700',color:'#374151',textAlign:'center'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {['VF','F','M','S','VS','NM'].map((v,i)=>{
                      const row = ['XS','S','M','L','XL'].map(s=>analysis.matrix[`${v}-${s}`]||0);
                      const tot = row.reduce((a,b)=>a+b,0);
                      return(<tr key={v} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                        <td style={{padding:'7px 12px',fontWeight:'700',
                          color:['VF','F'].includes(v)?'#166534':v==='M'?'#854d0e':'#6b7280'}}>{v}</td>
                        {row.map((n,j)=>(
                          <td key={j} style={{padding:'7px 12px',textAlign:'center',
                            color:n>0?'#374151':'#d1d5db',background:n>500?'#eff6ff':n>100?'#fef9c3':'transparent'}}>
                            {n>0?n.toLocaleString():'—'}
                          </td>))}
                        <td style={{padding:'7px 12px',textAlign:'center',fontWeight:'700'}}>{tot.toLocaleString()}</td>
                      </tr>);
                    })}
                    <tr style={{background:'#f8fafc',fontWeight:'700'}}>
                      <td style={{padding:'7px 12px'}}>Total</td>
                      {['XS','S','M','L','XL'].map(s=>{
                        const t=['VF','F','M','S','VS','NM'].reduce((sum,v)=>sum+(analysis.matrix[`${v}-${s}`]||0),0);
                        return<td key={s} style={{padding:'7px 12px',textAlign:'center'}}>{t.toLocaleString()}</td>;
                      })}
                      <td style={{padding:'7px 12px',textAlign:'center',color:'#7c3aed'}}>
                        {analysis.metrics.totLocs.toLocaleString()}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            {/* Download buttons */}
            <div style={{display:'flex',gap:'12px'}}>
              <button onClick={()=>exportExcel(analysis,design,params)}
                style={{flex:1,padding:'12px',background:'linear-gradient(135deg,#059669,#047857)',
                  color:'#fff',border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'14px',
                  cursor:'pointer',fontFamily:'inherit'}}>
                ⬇ Download Excel Report
              </button>
              <button onClick={()=>exportPPT(analysis,design,params)}
                style={{flex:1,padding:'12px',background:'linear-gradient(135deg,#7c3aed,#6d28d9)',
                  color:'#fff',border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'14px',
                  cursor:'pointer',fontFamily:'inherit'}}>
                📊 Download PPT Report
              </button>
            </div>
          </>)}
        </div>
      </div>
    </div>
  );
}
