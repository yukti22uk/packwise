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
  XS: { name:'Compartment Tray',        dims:'300×200×100 mm',     phys:[300,200,100],      volCm3:6000,    fill:0.55, slotH:0.12 },
  S:  { name:'Small Tote/Bin',          dims:'400×300×200 mm',     phys:[400,300,200],      volCm3:24000,   fill:0.55, slotH:0.22 },
  M:  { name:'Louvre/Shelf Bin',        dims:'600×400×300 mm',     phys:[600,400,300],      volCm3:72000,   fill:0.55, slotH:0.32 },
  L:  { name:'Stack Crate/Half-Pallet', dims:'800×600×400 mm',     phys:[800,600,400],      volCm3:192000,  fill:0.55, slotH:0.45 },
  XL: { name:'Standard Pallet',         dims:'1200×1000×1200 mm',  phys:[1200,1000,1200],   volCm3:1440000, fill:0.55, slotH:1.20 },
  LONG:{ name:'Long-Goods Slot',        dims:'per item',            phys:null,               volCm3:null,    fill:0.40, slotH:0.40 },
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
// Check if SKU fits in bin using optimal orientation
// Sort both descending and compare element-wise (largest sku dim must fit in largest bin dim)
function fitsInBin(skuL, skuW, skuH, binKey) {
  const b = BIN_CATALOG[binKey];
  if (!b || !b.phys) return false;
  const sku = [skuL, skuW, skuH].sort((a,b)=>b-a);
  const bin = [...b.phys].sort((a,b)=>b-a);
  return sku[0]<=bin[0] && sku[1]<=bin[1] && sku[2]<=bin[2];
}

// Select smallest bin that physically fits the SKU
// Falls back to volume-based selection if dimensions unknown
function selectBin(skuL, skuW, skuH, volCm3, isLong) {
  if (isLong) return 'LONG';
  if (skuL > 0 && skuW > 0 && skuH > 0) {
    // Physical fit check — smallest bin that fits wins
    for (const band of ['XS','S','M','L','XL']) {
      if (fitsInBin(skuL, skuW, skuH, band)) return band;
    }
    return 'XL'; // larger than XL pallet — needs custom handling
  }
  // Fallback to volume-only when dimensions missing
  if (!volCm3)       return 'S';
  if (volCm3 <= 500)  return 'XS';
  if (volCm3 <= 3000) return 'S';
  if (volCm3 <= 15000)return 'M';
  if (volCm3 <= 50000)return 'L';
  return 'XL';
}

function unitsPerBin(skuL, skuW, skuH, skuVolCm3, band) {
  const b = BIN_CATALOG[band];
  if (!b || !b.volCm3 || !skuVolCm3) return 1;
  // Physical layout: how many fit per layer × layers
  if (b.phys && skuL > 0 && skuW > 0 && skuH > 0) {
    // Try orientation where SKU height = tallest dim (standing up)
    const skuDims = [skuL, skuW, skuH].sort((a,b)=>b-a);
    const binDims = [...b.phys].sort((a,b)=>b-a);
    const perRow = Math.floor(binDims[1] / skuDims[1]);
    const perCol = Math.floor(binDims[2] / skuDims[2]);
    const layers = Math.floor(binDims[0] / skuDims[0]);
    const byLayout = Math.max(1, perRow * perCol * layers);
    // Also volume-based estimate
    const byVolume = Math.max(1, Math.floor(b.volCm3 * b.fill / skuVolCm3));
    // Use the lower (more conservative) of the two
    return Math.min(byLayout, byVolume);
  }
  return Math.max(1, Math.floor(b.volCm3 * b.fill / skuVolCm3));
}


// ─── TRUCK TYPE CATALOGUE ────────────────────────────────────────────────────
const TRUCK_TYPES = {
  small:  { label:'Small (Tata Ace / Eicher 10ft)', stagingDepth:6,  defaultPallets:2,  dockTimeH:0.50, volM3:3.8  },
  medium: { label:'Medium (Eicher 17ft / 19ft)',    stagingDepth:8,  defaultPallets:8,  dockTimeH:0.75, volM3:23.0 },
  large:  { label:'Large (32ft / Container truck)', stagingDepth:13, defaultPallets:20, dockTimeH:1.25, volM3:51.0 },
};
const PALLET_FP     = 1.44;  // m² per pallet footprint (1.2×1.2m)
const PALLET_VOL    = 1.728; // m³ per pallet (1.2×1.2×1.2m)
const PALLET_FILL   = 0.65;  // pallet fill efficiency
const STAGING_SAFETY= 1.5;   // peak buffer factor
const DOCK_EFF      = 0.85;  // dock utilisation efficiency
const TRUCK_FILL    = 0.70;  // truck volume fill efficiency

// ─── STAGING AREA CALCULATION ─────────────────────────────────────────────────
function calcStagingParams(params, analysis) {
  const { truckMix, dockConfig, dockPitch, inboundDwellH, outboundDwellH,
    packingInDispatch, packingBenches, shifts,
    inboundMode, outboundMode,
    inbBoxSizes, inbStackH,
    outbTruckType, outbTrucksPerDay, outbStackH } = params;

  const workingH  = (parseInt(shifts)||1) * 8;
  const pitch     = parseFloat(dockPitch)||4.5;
  const idwell    = parseFloat(inboundDwellH)||4;
  const odwell    = parseFloat(outboundDwellH)||2;
  const stackInb  = parseInt(inbStackH)||3;
  const stackOut  = parseInt(outbStackH)||3;

  let inbVehicles=0, outVehicles=0, inbDockH=0, outDockH=0;
  let inbUnits=0, outUnits=0, inbStagingArea=0, outStagingArea=0;
  let inbPalletsInDwell=0, outPalletsInDwell=0;
  let trucksNeeded=0, outDailyVolM3=0;
  const inbLabel = inboundMode==='boxes' ? 'boxes' : 'pallets';
  const outLabel = outboundMode==='boxes' ? 'boxes' : 'pallets';

  // ── INBOUND ────────────────────────────────────────────────────────────────
  if (inboundMode === 'boxes') {
    // Sum all box sizes
    let totalBoxFP=0; // m² of box footprints in peak dwell
    (inbBoxSizes||[]).forEach(b => {
      const L   = parseFloat(b.L)/1000||0; // mm→m
      const W   = parseFloat(b.W)/1000||0;
      const qty = parseFloat(b.qtyPerDay)||0;
      if (!qty||!L||!W) return;
      inbUnits += qty;
      const boxesInDwell = qty * idwell / workingH;
      totalBoxFP += boxesInDwell * L * W;
    });
    inbStagingArea = totalBoxFP / stackInb * STAGING_SAFETY;
    // Estimate vehicles from total inbound volume
    const totalInbVol = (inbBoxSizes||[]).reduce((s,b)=>{
      const L=parseFloat(b.L)/1000||0, W=parseFloat(b.W)/1000||0,
            H=parseFloat(b.H)/1000||0.3, qty=parseFloat(b.qtyPerDay)||0;
      return s + L*W*H*qty;
    }, 0);
    inbVehicles = Math.max(1, Math.ceil(totalInbVol/(TRUCK_TYPES.medium.volM3*TRUCK_FILL)));
    inbDockH    = inbVehicles * TRUCK_TYPES.medium.dockTimeH;

  } else {
    // Pallets mode — from truck mix, pallets per truck editable
    (truckMix||[]).forEach(t => {
      const tt  = TRUCK_TYPES[t.type]||TRUCK_TYPES.medium;
      const inb = parseFloat(t.inboundVehicles)||0;
      const ppt = parseFloat(t.palletsPerTruck)||tt.defaultPallets;
      inbVehicles += inb;
      inbUnits    += inb * ppt;
      inbDockH    += inb * tt.dockTimeH;
    });
    inbPalletsInDwell = inbUnits * idwell / workingH;
    inbStagingArea    = inbPalletsInDwell * PALLET_FP * STAGING_SAFETY;
  }

  // ── OUTBOUND ───────────────────────────────────────────────────────────────
  if (outboundMode === 'boxes') {
    // Calculate from Order data
    const dailyBoxes  = analysis?.dailyOutboundBoxes  || 0;
    outDailyVolM3     = analysis?.dailyOutboundVolM3   || 0;
    outUnits          = dailyBoxes;
    const avgBoxFP    = analysis?.avgBoxFootprintM2    || 0.06;

    // Truck calculation
    const outTT = TRUCK_TYPES[outbTruckType]||TRUCK_TYPES.medium;
    if (parseFloat(outbTrucksPerDay) > 0) {
      trucksNeeded = parseFloat(outbTrucksPerDay);
    } else if (outDailyVolM3 > 0) {
      trucksNeeded = Math.ceil(outDailyVolM3 / (outTT.volM3 * TRUCK_FILL));
    }
    outVehicles  = trucksNeeded;
    outDockH     = outVehicles * outTT.dockTimeH;

    const outBoxesInDwell = outUnits * odwell / workingH;
    outStagingArea = outBoxesInDwell * avgBoxFP / stackOut * STAGING_SAFETY;

  } else {
    // Pallets mode — from Order data (daily volume) or truck mix
    if (analysis?.dailyOutboundVolM3 > 0) {
      outDailyVolM3 = analysis.dailyOutboundVolM3;
      outUnits      = Math.ceil(outDailyVolM3 / (PALLET_VOL * PALLET_FILL));
    }
    // Also add from truck mix outbound
    let truckOutPallets=0;
    (truckMix||[]).forEach(t => {
      const tt  = TRUCK_TYPES[t.type]||TRUCK_TYPES.medium;
      const out = parseFloat(t.outboundVehicles)||0;
      const ppt = parseFloat(t.palletsPerTruck)||tt.defaultPallets;
      outVehicles   += out;
      truckOutPallets += out * ppt;
      outDockH      += out * tt.dockTimeH;
    });
    // Use order data if available, else truck mix
    if (outUnits === 0) outUnits = truckOutPallets;
    outPalletsInDwell = outUnits * odwell / workingH;
    outStagingArea    = outPalletsInDwell * PALLET_FP * STAGING_SAFETY;
  }

  // ── DOCK COUNT ────────────────────────────────────────────────────────────
  const availDockH = workingH * DOCK_EFF;
  let inboundDocks, outboundDocks;
  if (dockConfig === 'separate') {
    inboundDocks  = Math.max(1, Math.ceil(inbDockH  / availDockH));
    outboundDocks = Math.max(1, Math.ceil(outDockH  / availDockH));
  } else {
    const totalDockH = inbDockH + outDockH;
    const shared     = Math.max(2, Math.ceil(totalDockH / availDockH));
    inboundDocks     = Math.max(1,Math.round(shared*(inbDockH/(totalDockH||1))));
    outboundDocks    = Math.max(1, shared - inboundDocks);
  }
  const totalDocks = inboundDocks + outboundDocks;

  // ── APRONS + PACKING ─────────────────────────────────────────────────────
  const grnApron      = inboundDocks  * pitch * 2;
  const dispatchApron = outboundDocks * pitch * 2;
  const packingArea   = packingInDispatch ? (parseInt(packingBenches)||0)*4 : 0;

  const receivingArea = Math.max(30, Math.ceil(inbStagingArea + grnApron));
  const dispatchArea  = Math.max(30, Math.ceil(outStagingArea + packingArea + dispatchApron));

  return {
    inboundDocks, outboundDocks, totalDocks,
    receivingArea, dispatchArea,
    inbUnits, outUnits, inbLabel, outLabel,
    inbPalletsInDwell, outPalletsInDwell,
    inbVehicles, outVehicles,
    trucksNeeded, outDailyVolM3,
    stagingBreakdown: {
      inbStorage:    Math.ceil(inbStagingArea),
      grnApron:      Math.ceil(grnApron),
      outStorage:    Math.ceil(outStagingArea),
      packingArea,
      dispatchApron: Math.ceil(dispatchApron),
    },
  };
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
    const isLong = maxDim > 600;
    master[sku] = { L,W,H, volCm3, maxDim, isLong,
      sb: selectBin(L, W, H, volCm3, isLong) };
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
    const upb   = isLong ? 1 : unitsPerBin(m.L, m.W, m.H, m.volCm3, bin);
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

  // Daily outbound metrics for staging calculation
  const qtyMap = {};
  orderRows.forEach(r => {
    const sku = r[2]||r[0]; if (!sku) return;
    qtyMap[sku] = (qtyMap[sku]||0) + (parseFloat(r[3])||1);
  });
  const periodDays = Math.max(1, new Set(orderRows.map(r=>r[4]).filter(d=>d)).size);
  let dailyOutboundVolM3=0, totalQtyInPeriod=0;
  slotted.forEach(r => {
    const qty = qtyMap[r.sku]||0;
    dailyOutboundVolM3 += (qty/periodDays) * (r.volCm3/1e6);
    totalQtyInPeriod   += qty;
  });
  const dailyOutboundBoxes  = totalQtyInPeriod / periodDays;
  const validSkus = slotted.filter(s=>s.L>0&&s.W>0);
  const avgBoxFootprintM2   = validSkus.length > 0
    ? validSkus.reduce((s,r)=>s+(r.L/1000)*(r.W/1000),0)/validSkus.length : 0.06;

  return { slotted, matrix, zoneSummary, rackSummary,
    metrics: { totSKUs, totLocs, totStock, longCount, nmCount, nmStock },
    dailyOutboundVolM3: +dailyOutboundVolM3.toFixed(3),
    dailyOutboundBoxes: +dailyOutboundBoxes.toFixed(1),
    avgBoxFootprintM2:  +avgBoxFootprintM2.toFixed(4) };
}

// ─── RACK CONFIGURATION ENGINE ────────────────────────────────────────────────
const SHELVING_AISLE_MM = 1200; // manual pick aisle — narrower than forklift aisle

// Try placing bin in bay with given orientation, return layout metrics
function tryShelfOrientation(binDims, bayW, bayD, shelfH, clearanceMm, orient) {
  const [bL, bW, bH] = orient === 'LW'
    ? [binDims[0], binDims[1], binDims[2]]
    : [binDims[1], binDims[0], binDims[2]];
  const acrossW    = Math.floor(bayW / bL);
  const acrossD    = Math.floor(bayD / bW);
  const levelH     = bH + clearanceMm;
  const levels     = levelH > 0 ? Math.floor(shelfH / levelH) : 0;
  const locsPerBay = acrossW * acrossD * levels;
  return { acrossW, acrossD, levels, locsPerBay,
    feasible: acrossW>0 && acrossD>0 && levels>0 };
}

// Recalculate a single rack config row from its parameters
function recalcCfg(cfg) {
  if (['shelving','liveStorage'].includes(cfg.rack) && cfg.binDims) {
    const { bayW, bayD, shelfH, clearance, orientation, tiers, locs } = cfg;
    // tierHeight = usable height per tier (user-editable); defaults to shelfH for 1 tier
    const tierH = parseFloat(cfg.tierHeight) || shelfH;
    const r = tryShelfOrientation(cfg.binDims, bayW, bayD, tierH, clearance, orientation);
    // Both orientations for comparison
    const o1 = tryShelfOrientation(cfg.binDims, bayW, bayD, tierH, clearance, 'LW');
    const o2 = tryShelfOrientation(cfg.binDims, bayW, bayD, tierH, clearance, 'WL');
    const t  = parseInt(tiers)||1;
    const locsPerBayTotal = r.locsPerBay * t;
    const baysNeeded = (r.feasible && locsPerBayTotal>0)
      ? Math.ceil(locs / locsPerBayTotal) : 0;
    const aisleMm = cfg.shelvingAisle || SHELVING_AISLE_MM;
    const bayFP   = (bayW/1000)*(bayD/1000);
    const aisleA  = baysNeeded*(bayW/1000)*(aisleMm/1000);
    return { ...cfg, ...r, o1, o2, locsPerBayTotal, baysNeeded,
      area: +(baysNeeded*bayFP + aisleA).toFixed(1) };
  } else {
    const baysNeeded = cfg.locsPerBay>0 ? Math.ceil(cfg.locs/cfg.locsPerBay) : 0;
    const bayFP  = (cfg.bayW/1000)*(cfg.bayD/1000);
    const aisleM = (cfg.aisleW||3000)/1000;
    const area   = +(baysNeeded*bayFP + baysNeeded*(cfg.bayW/1000)*aisleM).toFixed(1);
    return { ...cfg, baysNeeded, area };
  }
}

// Auto-generate rack config from analysis
function generateRackConfig(analysis, params) {
  const { clearH, forkType, aisleW } = params;
  const shelfMaxH = Math.min(3500, Math.floor(clearH*1000 - 300));
  const maxLift   = { manual:2200, counterbalance:6000, reach:9000, vna:12000 };
  const liftH     = maxLift[forkType]||6000;
  const palletLevelH = 1500;
  const palletLevels = Math.max(1,
    Math.floor((Math.min(liftH, clearH*1000) - 800) / palletLevelH));
  const aisleWmm = Math.floor(parseFloat(aisleW)*1000);

  // Group by rack + bin
  const groups = {};
  (analysis.slotted||[]).forEach(r => {
    const key = `${r.rack}|${r.bin}`;
    if (!groups[key]) groups[key] = {
      rack:r.rack, bin:r.bin, rackName:r.rackName,
      binName:r.binName, locs:0 };
    groups[key].locs += r.locsReq;
  });

  return Object.values(groups).filter(g=>g.locs>0).map(g => {
    const binDims = BIN_CATALOG[g.bin]?.phys || null;

    if (['shelving','liveStorage'].includes(g.rack) && binDims) {
      const [bL,bW,bH] = binDims;
      const bayW = 900;
      const bayD = Math.max(Math.max(bL,bW)+50, 400); // min depth to fit 1 bin
      const clearance = 50;
      const o1 = tryShelfOrientation(binDims,bayW,bayD,shelfMaxH,clearance,'LW');
      const o2 = tryShelfOrientation(binDims,bayW,bayD,shelfMaxH,clearance,'WL');
      const bestOrient = o1.locsPerBay>=o2.locsPerBay ? 'LW' : 'WL';
      const best = bestOrient==='LW' ? o1 : o2;
      return recalcCfg({ id:`${g.rack}|${g.bin}`, ...g, binDims,
        bayW, bayD, shelfH:shelfMaxH, clearance,
        tierHeight: shelfMaxH, // per-tier height (same as shelfH for 1 tier)
        orientation:bestOrient, tiers:1, shelvingAisle:SHELVING_AISLE_MM,
        locsPerBay:best.locsPerBay, o1, o2, aisleW:SHELVING_AISLE_MM,
        ...best });

    } else if (['selective','driveIn','doubleDeep'].includes(g.rack)) {
      const depth    = g.rack==='driveIn'?6 : g.rack==='doubleDeep'?2 : 1;
      const bayW_mm  = 2700;
      const bayD_mm  = g.rack==='driveIn'?depth*1100 : g.rack==='doubleDeep'?2400 : 1100;
      const locsPerBay = 2*palletLevels*depth;
      return recalcCfg({ id:`${g.rack}|${g.bin}`, ...g, binDims,
        bayW:bayW_mm, bayD:bayD_mm, levels:palletLevels,
        locsPerBay, tiers:1, orientation:'std',
        acrossW:2, acrossD:depth, aisleW:aisleWmm });

    } else if (g.rack==='cantilever') {
      const levels = Math.max(1, Math.floor((clearH*1000-500)/600));
      return recalcCfg({ id:`${g.rack}|${g.bin}`, ...g, binDims,
        bayW:1500, bayD:2500, levels, locsPerBay:2*levels,
        tiers:1, orientation:'std', aisleW:3000 });

    } else {
      return recalcCfg({ id:`${g.rack}|${g.bin}`, ...g, binDims,
        bayW:900, bayD:600, levels:4, locsPerBay:8,
        tiers:1, orientation:'std', aisleW:1200 });
    }
  });
}

// Compute rackAreas map from confirmed rack config
function rackAreasFromConfig(rackConfig) {
  const areas = {};
  (rackConfig||[]).forEach(cfg => {
    areas[cfg.rack] = (areas[cfg.rack]||0) + (cfg.area||0);
  });
  return areas;
}

// ─── WAREHOUSE SIZING ─────────────────────────────────────────────────────────
function calcWarehouseSize(analysis, params, customRackAreas) {
  const { clearH, forkType, dockSide, aisleW } = params;
  const { rackSummary } = analysis;

  // ── Throughput-based staging (replaces 10% rule) ──────────────────────────
  const staging = calcStagingParams(params, analysis);
  const { receivingArea, dispatchArea, totalDocks, inboundDocks, outboundDocks } = staging;

  // ── Rack levels ───────────────────────────────────────────────────────────
  const palletLevelH = 1.5;
  const maxLift = { manual:2.2, counterbalance:6.0, reach:9.0, vna:12.0 };
  const liftH = maxLift[forkType]||6.0;
  const palletLevels = Math.max(1, Math.floor((Math.min(liftH, clearH) - 0.8) / palletLevelH));
  const shelfSlotH   = 0.35;
  const shelfLevels  = Math.max(1, Math.floor((Math.min(3.5, clearH) - 0.3) / shelfSlotH));

  // ── Racking area (use confirmed config if provided, else estimate) ───────
  const AISLE_FACTOR = 1 + (parseFloat(aisleW)||3.0) / 3.0;
  const rackAreas = customRackAreas || (() => {
    const ra = {};
    Object.entries(RACK_DEFS).forEach(([rk, rd]) => {
      const locs = rackSummary[rk]?.locs || 0;
      if (locs === 0) { ra[rk] = 0; return; }
      let locsPerBay;
      if (['selective','driveIn','doubleDeep'].includes(rk)) {
        const depth = rk==='driveIn'?6 : rk==='doubleDeep'?2 : 1;
        locsPerBay = 2 * palletLevels * depth;
      } else if (rk==='shelving'||rk==='liveStorage') {
        locsPerBay = shelfLevels * 2; // rough estimate until config confirmed
      } else if (rk==='cantilever') {
        locsPerBay = 8;
      } else {
        locsPerBay = 4;
      }
      const bays = Math.ceil(locs/Math.max(1,locsPerBay));
      ra[rk] = +(bays*rd.bayW*rd.bayD*AISLE_FACTOR).toFixed(1);
    });
    return ra;
  })();

  const netRackArea     = Object.values(rackAreas).reduce((s,v)=>s+v, 0);
  // ── MHE charging area ────────────────────────────────────────────────────
  const MHE_BAY_M2 = { manual:0, counterbalance:14, reach:9, vna:6 };
  const mheBayM2   = MHE_BAY_M2[forkType] || 0;
  const nMHE       = parseInt(params.nMHE) || (forkType==='manual'?0:Math.max(1,Math.ceil(totalDocks/2)));
  const mheArea    = +(nMHE * mheBayM2 * 1.3).toFixed(0); // +30% circulation
  const officeArea      = 50;
  const circulationArea = netRackArea * 0.08;
  const totalGrossArea  = netRackArea + receivingArea + dispatchArea + mheArea + officeArea + circulationArea;

  // ── Warehouse dimensions ──────────────────────────────────────────────────
  const pitch        = parseFloat(params.dockPitch)||4.5;
  const minDockWidth = totalDocks * pitch + 6;
  const recWidth     = dockSide==='both' ? Math.sqrt(totalGrossArea*0.8) : Math.sqrt(totalGrossArea*0.6);
  const wW = Math.max(minDockWidth, Math.ceil(recWidth/5)*5);
  const wL = Math.ceil(totalGrossArea/wW/5)*5;

  // ── Zone areas ────────────────────────────────────────────────────────────
  const totalLocs = analysis.metrics.totLocs || 1;
  const zoneAreas = {};
  Object.entries(analysis.zoneSummary).forEach(([z, zs]) => {
    zoneAreas[z] = +(netRackArea * (zs.locs/totalLocs)).toFixed(1);
  });

  return { wW, wL,
    totalGrossArea: +totalGrossArea.toFixed(0),
    netRackArea:    +netRackArea.toFixed(0),
    receivingArea, dispatchArea, mheArea, officeArea, circulationArea: +circulationArea.toFixed(0),
    rackAreas, zoneAreas, palletLevels, shelfLevels,
    totalDocks, inboundDocks, outboundDocks, staging,
    nMHE, mheBayM2, forkType,
  };
}

// ─── SVG FLOOR PLAN ───────────────────────────────────────────────────────────
function FloorPlanSVG({ analysis, design, params, rackConfig }) {
  const MFT  = 3.2808;
  const M2FT = 10.7639;
  const ft   = m  => `${(m*MFT).toFixed(0)}'`;
  const sqft = m2 => `${Math.round(m2*M2FT).toLocaleString()} sq ft`;
  const m2lbl= (m2,label) => label ? `${label}\n${m2}m²\n(${sqft(m2)})` : `${m2}m²\n(${sqft(m2)})`;

  const { wW, wL, zoneAreas, receivingArea, dispatchArea, mheArea, officeArea,
    totalDocks, inboundDocks, outboundDocks, staging, netRackArea } = design;
  const { dockSide, aisleW:aisleWParam, dockPitch, forkType,
    packingBenches, inboundMode, outboundMode } = params;
  const pitch   = parseFloat(dockPitch)||4.5;
  const aisleM  = parseFloat(aisleWParam)||3.0;

  // Canvas
  const SVG_W=960, SVG_H=720;
  const ML=62, MR=70, MT=50, MB=70; // margins for labels
  const DW=SVG_W-ML-MR, DH=SVG_H-MT-MB;
  const sX=DW/wW, sY=DH/wL;
  const X=m=>ML+m*sX, Y=m=>MT+m*sY, W=m=>m*sX, H=m=>m*sY;

  // ── AREA HEIGHTS ────────────────────────────────────────────────────────────
  const recH    = Math.max(4,(receivingArea||0)/wW);
  const disH    = Math.max(4,(dispatchArea||0)/wW);
  const stagingH= Math.max(recH,disH);
  const offH    = Math.max(3,(officeArea||50)/wW);
  const mheH    = mheArea>0 ? Math.max(2,mheArea/wW) : 0;
  const supportH= offH + mheH;

  const isOne   = dockSide==='one';
  const isBoth  = dockSide==='both';
  const isCorner= dockSide==='corner';

  const zonesH  = isOne
    ? Math.max(0,wL-stagingH-supportH)
    : Math.max(0,wL-recH-disH-supportH);

  // Zone vertical allocation
  const zoneOrder=['golden','mid','reserve','bulk','long'];
  const totZA = zoneOrder.reduce((s,z)=>s+(zoneAreas[z]||0),0)||1;
  const zH    = {};
  zoneOrder.forEach(z=>{ zH[z]=((zoneAreas[z]||0)/totZA)*zonesH; });

  // Build zone rects (from north going south)
  const zoneRects=[], stagingRects=[], supportRects=[];

  // Support area at NORTH (top)
  let cur=0;
  if (officeArea>0) {
    supportRects.push({ key:'office', x:0, y:cur, w:wW/2, h:offH,
      label:'OFFICE / WELFARE', color:'#dbeafe', border:'#3b82f6', text:'#1d4ed8' });
  }
  if (mheH>0) {
    supportRects.push({ key:'mhe', x:wW/2, y:cur, w:wW/2, h:offH+mheH,
      label:'MHE CHARGING', color:'#fdf4ff', border:'#9333ea', text:'#6b21a8' });
  }
  cur+=supportH;

  if (isBoth) {
    // Dispatch at north (after support)
    stagingRects.push({ key:'dispatch', x:0, y:cur, w:wW, h:disH,
      label:'DISPATCH / PACKING', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
    cur+=disH;
  }

  // Zones (bulk→golden south)
  const zonesTopToBottom=['bulk','long','reserve','mid','golden'];
  zonesTopToBottom.forEach(z=>{
    const zh=zH[z]||0; if(zh<0.5) return;
    zoneRects.push({ key:z, x:0, y:cur, w:wW, h:zh,
      label:ZONE_DEFS[z].label, color:ZONE_DEFS[z].color,
      border:ZONE_DEFS[z].border, text:ZONE_DEFS[z].textColor,
      area:zoneAreas[z]||0 });
    cur+=zh;
  });

  // Staging at south
  if (isOne) {
    stagingRects.push({ key:'receiving', x:0, y:cur, w:wW/2, h:stagingH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
    stagingRects.push({ key:'dispatch', x:wW/2, y:cur, w:wW/2, h:stagingH,
      label:'DISPATCH / PACKING', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
  } else if (isBoth) {
    stagingRects.push({ key:'receiving', x:0, y:cur, w:wW, h:recH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
  } else {
    const eastW=Math.min(wW*0.3,14);
    stagingRects.push({ key:'receiving', x:0, y:cur, w:wW-eastW, h:stagingH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
    stagingRects.push({ key:'dispatch', x:wW-eastW, y:cur, w:eastW, h:stagingH,
      label:'DISPATCH', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
  }

  // Dock doors
  const dockDoors=[];
  const doorW=3.5;
  if (isOne) {
    const sp=wW/(totalDocks+1);
    for(let i=1;i<=totalDocks;i++) dockDoors.push({x:sp*i-doorW/2,y:wL,side:'south',label:`D${i}`});
  } else if (isBoth) {
    const ssp=wW/(inboundDocks+1);
    for(let i=1;i<=inboundDocks;i++) dockDoors.push({x:ssp*i-doorW/2,y:wL,side:'south',label:`D${i}`});
    const nsp=wW/(outboundDocks+1);
    for(let i=1;i<=outboundDocks;i++) dockDoors.push({x:nsp*i-doorW/2,y:0,side:'north',label:`D${inboundDocks+i}`});
  } else {
    const eastW=Math.min(wW*0.3,14);
    const southN=inboundDocks, eastN=outboundDocks;
    const ssp2=(wW-eastW)/(southN+1);
    for(let i=1;i<=southN;i++) dockDoors.push({x:ssp2*i-doorW/2,y:wL,side:'south',label:`D${i}`});
    const esp=wL/(eastN+1);
    for(let i=1;i<=eastN;i++) dockDoors.push({x:wW,y:esp*i,side:'east',label:`D${southN+i}`});
  }

  // ── RACK ROW HELPER ─────────────────────────────────────────────────────────
  // For each zone determine dominant rack type and draw stylized top-view rows
  const dominantRack = {};
  (rackConfig||[]).forEach(cfg=>{
    const z = ZONE_DEFS[cfg.rack]
      ? cfg.rack
      : (analysis?.slotted||[]).find(s=>s.rack===cfg.rack)?.zone || null;
    // Map rack type to zone via slotted data
  });
  // Build zone→rackType map from slotted
  const zoneRackType={};
  (analysis?.slotted||[]).forEach(r=>{
    if(!zoneRackType[r.zone]) zoneRackType[r.zone]={ shelving:0,selective:0,driveIn:0,cantilever:0,liveStorage:0 };
    zoneRackType[r.zone][r.rack]=(zoneRackType[r.zone][r.rack]||0)+r.locsReq;
  });
  const getDomRack=zone=>{
    const m=zoneRackType[zone]; if(!m) return 'shelving';
    return Object.entries(m).sort((a,b)=>b[1]-a[1])[0]?.[0]||'shelving';
  };

  // Draw rack rows within a zone rect
  const rackRowsForZone=(zone)=>{
    const dom=getDomRack(zone.key);
    const rows=[];
    const RACK_INFO={
      shelving:   {depth:0.5,  color:'#cbd5e1', stroke:'#94a3b8'},
      liveStorage:{depth:0.6,  color:'#bfdbfe', stroke:'#60a5fa'},
      selective:  {depth:1.1,  color:'#d1d5db', stroke:'#6b7280'},
      doubleDeep: {depth:2.4,  color:'#c7d2fe', stroke:'#818cf8'},
      driveIn:    {depth:5.5,  color:'#e9d5ff', stroke:'#a855f7'},
      cantilever: {depth:2.0,  color:'#fde8d8', stroke:'#f97316'},
    };
    const ri=RACK_INFO[dom]||RACK_INFO.shelving;
    const aisle=dom==='shelving'||dom==='liveStorage' ? 1.2 : aisleM;
    const slot=ri.depth+aisle;
    const nRows=Math.max(1,Math.floor(zone.h/slot));
    for(let i=0;i<nRows;i++){
      const ry=zone.y+(slot*i)+aisle/2;
      if(ry+ri.depth>zone.y+zone.h-0.3) break;
      rows.push({ x:zone.x+0.4, y:ry, w:zone.w-0.8, h:ri.depth, ...ri, dom });
    }
    return rows;
  };

  // Pallet symbols in staging area
  const palletSymbols=(rect,mode)=>{
    const palW=1.2, palH=1.2;
    const cols=Math.floor((rect.w-0.8)/palW);
    const rowsN=Math.floor((rect.h-0.8)/palH);
    const syms=[];
    for(let r=0;r<Math.min(rowsN,4);r++)
      for(let cl=0;cl<Math.min(cols,12);cl++)
        syms.push({x:rect.x+0.4+cl*palW,y:rect.y+0.4+r*palH,w:palW-0.1,h:palH-0.1,mode});
    return syms;
  };

  // Packing table symbols (2.4×0.8m each)
  const nBenches=Math.min(parseInt(packingBenches)||0,8);
  const packTables=[];
  const dispRect=stagingRects.find(s=>s.key==='dispatch');
  if(dispRect && nBenches>0){
    for(let i=0;i<nBenches;i++){
      packTables.push({x:dispRect.x+0.4+i*2.8,y:dispRect.y+dispRect.h-1.4,w:2.4,h:0.8});
    }
  }

  // Collect all rack rows to draw
  const allRackRows=[];
  zoneRects.forEach(zone=>{ allRackRows.push(...rackRowsForZone(zone)); });

  // Staging pallets
  const recRect=stagingRects.find(s=>s.key==='receiving');
  const recPallets=recRect ? palletSymbols(recRect,'inbound') : [];
  const disPallets=dispRect ? palletSymbols(dispRect,'outbound') : [];

  // MHE charging bays
  const MHE_BAY={counterbalance:{w:4,h:3.5},reach:{w:3.5,h:2.5},vna:{w:3,h:2}};
  const mheBay=MHE_BAY[forkType]||{w:3.5,h:2.5};
  const mheRect=supportRects.find(s=>s.key==='mhe');
  const mheBays=[];
  if(mheRect && design.nMHE>0){
    for(let i=0;i<Math.min(design.nMHE,6);i++){
      mheBays.push({x:mheRect.x+0.4+i*(mheBay.w+0.4),y:mheRect.y+0.4,w:mheBay.w,h:mheBay.h});
    }
  }

  // Right-side zone dimension arrows (y positions)
  const dimRight=[];
  let dimCur=0;
  supportRects.length && dimRight.push({y:dimCur,h:supportH,label:`Support\n${officeArea+mheArea}m²`});
  dimCur+=supportH;
  if(isBoth){dimRight.push({y:dimCur,h:disH,label:`Dispatch\n${dispatchArea}m²`});dimCur+=disH;}
  zoneRects.forEach(zr=>{
    if(zr.h>0.5) dimRight.push({y:dimCur,h:zr.h,label:`${zr.label}\n${(zr.area||0).toFixed(0)}m²`});
    dimCur+=zr.h;
  });
  stagingRects.filter(s=>s.key==='receiving').forEach(s=>{
    dimRight.push({y:dimCur,h:s.h,label:`Staging\n${receivingArea}m²`});dimCur+=s.h;
  });

  return (
    <svg width={SVG_W} height={SVG_H} viewBox={`0 0 ${SVG_W} ${SVG_H}`}
      style={{border:'1px solid #e2e8f0',borderRadius:'10px',background:'#ffffff',
               width:'100%',height:'auto',display:'block'}}>

      <defs>
        <pattern id="palletPat" x="0" y="0" width={W(1.2)} height={H(1.2)} patternUnits="userSpaceOnUse">
          <rect x="1" y="1" width={W(1.2)-2} height={H(1.2)-2} fill="none" stroke="#94a3b8" strokeWidth="1" strokeDasharray="3,2"/>
          <line x1={W(0.4)} y1={H(0.6)} x2={W(0.8)} y2={H(0.6)} stroke="#94a3b8" strokeWidth="0.7"/>
          <line x1={W(0.6)} y1={H(0.4)} x2={W(0.6)} y2={H(0.8)} stroke="#94a3b8" strokeWidth="0.7"/>
        </pattern>
        <marker id="arr" markerWidth="6" markerHeight="6" refX="3" refY="3" orient="auto">
          <polygon points="0 0,6 3,0 6" fill="#64748b"/>
        </marker>
        <marker id="arrR" markerWidth="6" markerHeight="6" refX="3" refY="3" orient="auto-start-reverse">
          <polygon points="0 0,6 3,0 6" fill="#64748b"/>
        </marker>
      </defs>

      {/* ── WAREHOUSE OUTLINE ─── */}
      <rect x={X(0)} y={Y(0)} width={W(wW)} height={H(wL)} fill="#f8fafc" stroke="#1e293b" strokeWidth="2.5" rx="2"/>

      {/* ── SUPPORT AREAS (office + MHE) ─── */}
      {supportRects.map((s,i)=>(
        <g key={`sup-${i}`}>
          <rect x={X(s.x)} y={Y(s.y)} width={W(s.w)} height={H(s.h)}
            fill={s.color} stroke={s.border} strokeWidth="1.5" opacity="0.95"/>
          {s.key==='office' && (<>
            {/* Desk symbols */}
            {[0.6,1.8,3.0].filter(dx=>dx<s.w-1).map((dx,di)=>(
              <g key={di}>
                <rect x={X(s.x+dx)} y={Y(s.y+0.3)} width={W(0.8)} height={H(0.5)} fill="#93c5fd" stroke="#3b82f6" strokeWidth="0.5" rx="1"/>
                <rect x={X(s.x+dx+0.3)} y={Y(s.y+0.85)} width={W(0.25)} height={H(0.25)} fill="#3b82f6" rx="1"/>
              </g>))}
            <text x={X(s.x+s.w/2)} y={Y(s.y+s.h/2)} textAnchor="middle" dominantBaseline="middle"
              fontSize="10" fontWeight="700" fill={s.text}>{s.label}</text>
          </>)}
          {s.key==='mhe' && (<>
            {mheBays.map((b,bi)=>(
              <g key={bi}>
                <rect x={X(b.x)} y={Y(b.y)} width={W(b.w)} height={H(b.h)}
                  fill="#e9d5ff" stroke="#9333ea" strokeWidth="1" rx="1"/>
                <text x={X(b.x+b.w/2)} y={Y(b.y+b.h/2)} textAnchor="middle"
                  dominantBaseline="middle" fontSize="9" fill="#7c3aed">⚡</text>
              </g>))}
            <text x={X(s.x+s.w/2)} y={Y(s.y+s.h/2)} textAnchor="middle" dominantBaseline="middle"
              fontSize="10" fontWeight="700" fill={s.text}>
              {s.label}{design.nMHE>0?` (${design.nMHE} units)`:''}
            </text>
          </>)}
        </g>
      ))}

      {/* ── STORAGE ZONES ─── */}
      {zoneRects.map(z=>(
        <g key={`z-${z.key}`}>
          <rect x={X(z.x)} y={Y(z.y)} width={W(z.w)} height={H(z.h)}
            fill={z.color} stroke={z.border} strokeWidth="1.5" opacity="0.85"/>
        </g>
      ))}

      {/* ── RACK ROWS (top view) ─── */}
      {allRackRows.map((r,i)=>{
        const isDI  = r.dom==='driveIn';
        const isCan = r.dom==='cantilever';
        const px=X(r.x), py=Y(r.y), pw=W(r.w), ph=H(r.h);
        return(
          <g key={`rr-${i}`}>
            <rect x={px} y={py} width={pw} height={Math.max(2,ph)}
              fill={r.color} stroke={r.stroke} strokeWidth="0.8" rx="1"/>
            {/* Internal division lines (bay dividers) */}
            {isDI
              ? [1,2,3].map(p=><line key={p} x1={px+pw*p/4} y1={py} x2={px+pw*p/4} y2={py+ph} stroke={r.stroke} strokeWidth="0.5" strokeDasharray="2,2"/>)
              : Array.from({length:Math.floor(r.w/(r.dom==='selective'?2.7:0.9))},(_,p)=>(
                  <line key={p} x1={px+p*(r.dom==='selective'?W(2.7):W(0.9))} y1={py}
                    x2={px+p*(r.dom==='selective'?W(2.7):W(0.9))} y2={py+Math.max(2,ph)}
                    stroke={r.stroke} strokeWidth="0.4" strokeOpacity="0.6"/>
                ))
            }
          </g>
        );
      })}

      {/* ── STAGING AREAS ─── */}
      {stagingRects.map((s,i)=>(
        <g key={`stg-${i}`}>
          <rect x={X(s.x)} y={Y(s.y)} width={W(s.w)} height={H(s.h)}
            fill={s.color} stroke={s.border} strokeWidth="2" opacity="0.92"/>
        </g>
      ))}

      {/* ── INBOUND PALLET SYMBOLS ─── */}
      {recPallets.map((p,i)=>(
        <g key={`rp-${i}`}>
          <rect x={X(p.x)} y={Y(p.y)} width={W(p.w)} height={H(p.h)}
            fill="none" stroke="#0284c7" strokeWidth="0.8" strokeDasharray="2,1.5" rx="1"/>
          <line x1={X(p.x)+W(p.w)*0.3} y1={Y(p.y+p.h/2)} x2={X(p.x)+W(p.w)*0.7} y2={Y(p.y+p.h/2)} stroke="#0284c7" strokeWidth="0.6"/>
          <line x1={X(p.x+p.w/2)} y1={Y(p.y)+H(p.h)*0.3} x2={X(p.x+p.w/2)} y2={Y(p.y)+H(p.h)*0.7} stroke="#0284c7" strokeWidth="0.6"/>
        </g>
      ))}

      {/* ── OUTBOUND PALLET SYMBOLS ─── */}
      {disPallets.map((p,i)=>(
        <g key={`dp-${i}`}>
          <rect x={X(p.x)} y={Y(p.y)} width={W(p.w)} height={H(p.h)}
            fill="none" stroke="#d97706" strokeWidth="0.8" strokeDasharray="2,1.5" rx="1"/>
          <line x1={X(p.x)+W(p.w)*0.3} y1={Y(p.y+p.h/2)} x2={X(p.x)+W(p.w)*0.7} y2={Y(p.y+p.h/2)} stroke="#d97706" strokeWidth="0.6"/>
          <line x1={X(p.x+p.w/2)} y1={Y(p.y)+H(p.h)*0.3} x2={X(p.x+p.w/2)} y2={Y(p.y)+H(p.h)*0.7} stroke="#d97706" strokeWidth="0.6"/>
        </g>
      ))}

      {/* ── PACKING TABLES ─── */}
      {packTables.map((t,i)=>(
        <g key={`pt-${i}`}>
          <rect x={X(t.x)} y={Y(t.y)} width={W(t.w)} height={H(t.h)}
            fill="#374151" stroke="#111827" strokeWidth="1" rx="1"/>
          <text x={X(t.x+t.w/2)} y={Y(t.y+t.h/2)} textAnchor="middle"
            dominantBaseline="middle" fontSize="7" fill="#fff" fontWeight="600">TABLE</text>
        </g>
      ))}

      {/* ── ZONE LABELS ─── */}
      {zoneRects.map(z=>{
        const px=X(z.x), py=Y(z.y), pw=W(z.w), ph=H(z.h);
        if(ph<16) return null;
        const lines=[z.label, `${(z.area||0).toFixed(0)}m² · ${sqft(z.area||0)}`];
        return(
          <g key={`zl-${z.key}`}>
            {lines.map((ln,li)=>(
              <text key={li} x={px+pw/2} y={py+ph/2+(li-0.5)*12}
                textAnchor="middle" dominantBaseline="middle"
                fontSize={li===0?11:9} fontWeight={li===0?'700':'400'} fill={z.text}>
                {ln}
              </text>))}
          </g>
        );
      })}

      {/* ── STAGING LABELS ─── */}
      {stagingRects.map((s,i)=>{
        const px=X(s.x), py=Y(s.y), pw=W(s.w), ph=H(s.h);
        if(ph<14) return null;
        return(
          <g key={`sl-${i}`}>
            <text x={px+pw/2} y={py+ph/2-5} textAnchor="middle" dominantBaseline="middle"
              fontSize="10" fontWeight="700" fill={s.text}>{s.label}</text>
            {s.subLabel&&ph>26&&<text x={px+pw/2} y={py+ph/2+7} textAnchor="middle"
              dominantBaseline="middle" fontSize="8" fill={s.text}>{s.subLabel}</text>}
          </g>
        );
      })}

      {/* ── DIVIDER LINE (one-side: between receiving & dispatch) ─── */}
      {isOne&&(<line x1={X(wW/2)} y1={Y(wL-stagingH)} x2={X(wW/2)} y2={Y(wL)}
        stroke="#94a3b8" strokeWidth="1.5" strokeDasharray="4,3"/>)}

      {/* ── DOCK DOORS ─── */}
      {dockDoors.filter(d=>d.side==='south').map((d,i)=>(
        <g key={`ds-${i}`}>
          <rect x={X(d.x)} y={Y(wL)-5} width={W(doorW)} height={10} fill="#1d4ed8" rx="2"/>
          <text x={X(d.x+doorW/2)} y={Y(wL)+16} textAnchor="middle"
            fontSize="8" fill="#1d4ed8" fontWeight="700">{d.label}</text>
        </g>
      ))}
      {dockDoors.filter(d=>d.side==='north').map((d,i)=>(
        <g key={`dn-${i}`}>
          <rect x={X(d.x)} y={Y(0)-5} width={W(doorW)} height={10} fill="#7c3aed" rx="2"/>
          <text x={X(d.x+doorW/2)} y={Y(0)-12} textAnchor="middle"
            fontSize="8" fill="#7c3aed" fontWeight="700">{d.label}</text>
        </g>
      ))}
      {dockDoors.filter(d=>d.side==='east').map((d,i)=>(
        <g key={`de-${i}`}>
          <rect x={X(wW)-5} y={Y(d.y)} width={10} height={H(doorW)} fill="#7c3aed" rx="2"/>
          <text x={X(wW)+14} y={Y(d.y+doorW/2)} textAnchor="start"
            fontSize="8" fill="#7c3aed" fontWeight="700">{d.label}</text>
        </g>
      ))}

      {/* ── DIMENSION LINES ─── */}
      {/* Top — overall width */}
      <line x1={X(0)} y1={MT-26} x2={X(wW)} y2={MT-26} stroke="#64748b" strokeWidth="1.2"
        markerStart="url(#arrR)" markerEnd="url(#arr)"/>
      <text x={X(wW/2)} y={MT-30} textAnchor="middle" fontSize="11" fontWeight="800" fill="#0f172a">
        {`${wW}m (${ft(wW)})`}
      </text>

      {/* Left — overall height */}
      <line x1={ML-36} y1={Y(0)} x2={ML-36} y2={Y(wL)} stroke="#64748b" strokeWidth="1.2"
        markerStart="url(#arrR)" markerEnd="url(#arr)"/>
      <text x={ML-40} y={Y(wL/2)} textAnchor="middle" fontSize="11" fontWeight="800" fill="#0f172a"
        transform={`rotate(-90,${ML-40},${Y(wL/2)})`}>
        {`${wL}m (${ft(wL)})`}
      </text>

      {/* Right — zone heights */}
      {dimRight.map((d,i)=>{
        if(d.h<1.5) return null;
        const y1=Y(d.y), y2=Y(d.y+d.h), xR=X(wW)+16;
        return(
          <g key={`dr-${i}`}>
            <line x1={xR} y1={y1} x2={xR} y2={y2} stroke="#94a3b8" strokeWidth="1"
              markerStart="url(#arrR)" markerEnd="url(#arr)"/>
            {(y2-y1)>18&&<text x={xR+6} y={(y1+y2)/2} dominantBaseline="middle"
              fontSize="8" fill="#64748b">
              {`${d.h.toFixed(0)}m`}
            </text>}
          </g>
        );
      })}

      {/* ── COMPASS ─── */}
      <circle cx={SVG_W-22} cy={MT+18} r="14" fill="white" stroke="#e2e8f0" strokeWidth="1"/>
      <text x={SVG_W-22} y={MT+14} textAnchor="middle" fontSize="9" fontWeight="800" fill="#0f172a">N</text>
      <line x1={SVG_W-22} y1={MT+4} x2={SVG_W-22} y2={MT+30} stroke="#64748b" strokeWidth="1"/>
      <polygon points={`${SVG_W-22},${MT+4} ${SVG_W-25},${MT+18} ${SVG_W-19},${MT+18}`} fill="#0f172a"/>

      {/* ── SCALE BAR ─── */}
      {[{m:0,label:'0'},{m:10,label:`10m\n(33')`}].map((mark,i)=>(
        <g key={i}>
          <line x1={X(mark.m)} y1={SVG_H-26} x2={X(mark.m)} y2={SVG_H-20} stroke="#64748b" strokeWidth="1.5"/>
          <text x={X(mark.m)} y={SVG_H-10} textAnchor="middle" fontSize="8" fill="#64748b">{mark.label.split('\n')[0]}</text>
          {mark.label.includes('\n')&&<text x={X(mark.m)} y={SVG_H-2} textAnchor="middle" fontSize="7" fill="#9ca3af">{mark.label.split('\n')[1]}</text>}
        </g>))}
      <line x1={X(0)} y1={SVG_H-23} x2={X(Math.min(10,wW))} y2={SVG_H-23} stroke="#64748b" strokeWidth="1.5"/>

      {/* ── LEGEND ─── */}
      {[
        ['#cbd5e1','#94a3b8','Shelving rack'],
        ['#d1d5db','#6b7280','Pallet rack (selective)'],
        ['#e9d5ff','#a855f7','Drive-in / high-density'],
        ['none','#0284c7','Inbound pallet (staging)'],
        ['none','#d97706','Outbound pallet (staging)'],
        ['#374151','#111827','Packing table'],
      ].map(([fill,stroke,label],i)=>{
        const lx=X(0)+4, ly=SVG_H-68+i*10;
        return(
          <g key={i}>
            <rect x={lx} y={ly} width={14} height={8}
              fill={fill} stroke={stroke} strokeWidth="1" strokeDasharray={fill==='none'?'2,1.5':'0'} rx="1"/>
            <text x={lx+18} y={ly+6} fontSize="8" fill="#374151">{label}</text>
          </g>
        );
      })}

      {/* ── TOTAL AREA FOOTER ─── */}
      <text x={X(wW/2)} y={SVG_H-4} textAnchor="middle" fontSize="10" fontWeight="700" fill="#374151">
        {`Total gross area: ${(wW*wL).toLocaleString()}m²  (${Math.round(wW*wL*10.7639).toLocaleString()} sq ft)  ·  ${dockSide==='one'?'One-side':'Opposite-side'} docks`}
      </text>
    </svg>
  );
}


// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function exportExcel(analysis, design, params, rackConfig) {
  const wb   = XLSX.utils.book_new();
  const today= new Date().toLocaleDateString();
  const ws   = (data,cols) => {
    const s = XLSX.utils.aoa_to_sheet(data);
    if (cols) s['!cols'] = cols.map(w=>({wch:w}));
    return s;
  };
  const { slotted, metrics, zoneSummary, rackSummary, matrix } = analysis;
  const { wW, wL, totalGrossArea, netRackArea, rackAreas, palletLevels, shelfLevels,
    receivingArea, dispatchArea, mheArea, officeArea, circulationArea,
    nMHE, mheBayM2, forkType: dFork, staging, zoneAreas } = design;

  // Sheet 1: Summary
  XLSX.utils.book_append_sheet(wb, ws([
    ['WAREHOUSE STORAGE DESIGN REPORT'],['Generated:',today],[],
    ['HEADLINE METRICS'],
    ['Total Active SKUs',metrics.totSKUs],
    ['Total Current Stock (units)',metrics.totStock],
    ['Total Storage Locations Required',metrics.totLocs],
    ['Long/Awkward Items',metrics.longCount],
    ['No-Movement SKUs (in stock)',metrics.nmCount],[],
    ['WAREHOUSE SIZE RECOMMENDATION'],
    ['Recommended Width (m)',wW],
    ['Recommended Length (m)',wL],
    ['Total Gross Floor Area (m²)',wW*wL],
    ['Net Racking Area (m²)',netRackArea],
    ['Pallet Rack Levels',palletLevels],
    ['Shelf Rack Levels',shelfLevels],
    ['MHE Units',nMHE||'Auto'],
    ['MHE Charging Area (m²)',mheArea||0],[],
    ['ZONE BREAKDOWN'],
    ['Zone','SKUs','Locations','Stock Units','Pick Lines'],
    ...Object.entries(zoneSummary).map(([z,v])=>[
      ZONE_DEFS[z]?.label||z, v.skus, v.locs, v.stock, v.pickLines]),
    [],['RACK TYPE SUMMARY'],
    ['Rack Type','SKUs','Locations Required','Floor Area (m²)'],
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

  // Sheet 4: Rack Schedule (detailed — from confirmed rack config if available)
  const rackSchedRows = rackConfig && rackConfig.length > 0
    ? rackConfig.map(cfg => {
        const isSh = ['shelving','liveStorage'].includes(cfg.rack);
        const tierH = cfg.tierHeight || cfg.shelfH || '—';
        return [
          cfg.rackName, cfg.binName||'—',
          cfg.bayW+' mm', cfg.bayD+' mm',
          isSh ? (tierH+' mm / tier') : '—',
          cfg.levels||'—',
          isSh&&cfg.tiers>1 ? cfg.tiers : '—',
          isSh ? (cfg.orientation==='LW'?'L along width':'W along width') : 'Standard',
          isSh ? `${cfg.acrossW}×${cfg.acrossD}` : '2×1',
          cfg.locsPerBayTotal||cfg.locsPerBay||'—',
          cfg.baysNeeded||'—',
          cfg.locs, cfg.area||'—',
        ];
      })
    : Object.entries(rackSummary).map(([rk,rv])=>{
        const rd = RACK_DEFS[rk]||{};
        const area = rackAreas[rk]||0;
        const bayArea = (rd.bayW||1)*(rd.bayD||1);
        const bays = bayArea>0?Math.ceil(area/bayArea):0;
        return[rd.name||rk,rd.desc||'','—','—','—','—','—','—','—','—',bays,rv.locs,area];
      });

  XLSX.utils.book_append_sheet(wb, ws([
    ['RACK SCHEDULE — DETAILED'],[],
    rackConfig&&rackConfig.length>0
      ? ['Rack Type','Bin/Pallet Type','Bay Width','Bay Depth','Height/Tier','Levels/Tier','Tiers','Bin Orientation','Bins per Level','Locs per Bay','No. of Bays','Total Locs','Floor Area (m²)']
      : ['Rack Type','Description','Bay W','Bay D','Ht/Tier','Levels','Tiers','Orientation','Bins/Level','Locs/Bay','Bays','Total Locs','Area (m²)'],
    ...rackSchedRows,
    [],
    ['TOTALS','','','','','','','','','',
      rackSchedRows.reduce((s,r)=>s+(parseFloat(r[10])||0),0),
      rackSchedRows.reduce((s,r)=>s+(parseFloat(r[11])||0),0),
      rackSchedRows.reduce((s,r)=>s+(parseFloat(r[12])||0),0),
    ],
  ],[24,20,12,12,16,12,8,18,14,12,12,12,14]),'4. Rack Schedule');

  // Sheet 5: Bin & Pallet Schedule
  const binGroups = {};
  (slotted||[]).forEach(r => {
    if (!r.bin) return;
    if (!binGroups[r.bin]) binGroups[r.bin] = { bin:r.bin, binName:r.binName||r.bin, locs:0, stock:0 };
    binGroups[r.bin].locs  += r.locsReq;
    binGroups[r.bin].stock += r.stock;
  });
  XLSX.utils.book_append_sheet(wb, ws([
    ['BIN & PALLET SCHEDULE'],[],
    ['Container Type','Dimensions','Locations Needed','Units (Bins/Pallets) Required','Notes'],
    ...Object.values(binGroups).map(b => {
      const bc = BIN_CATALOG[b.bin];
      return [
        b.binName,
        bc?.dims||'—',
        b.locs,
        b.locs, // 1 bin per location
        b.bin==='XL'?'Standard pallet 1.2×1.2m assumed'
          :b.bin==='LONG'?'Size per item — cantilever slot'
          :'One container per location',
      ];
    }),
    [],
    ['TOTAL CONTAINERS REQUIRED','',
      Object.values(binGroups).reduce((s,b)=>s+b.locs,0),
      Object.values(binGroups).reduce((s,b)=>s+b.locs,0),''],
  ],[28,22,18,24,38]),'5. Bin & Pallet Schedule');

  // Sheet 6: Area Summary
  const toSqFt = m2 => Math.round((m2||0)*10.7639).toLocaleString()+' sq ft';
  const sb = design.staging?.stagingBreakdown || {};
  const mheLabel = nMHE && mheBayM2
    ? `${nMHE} × ${mheBayM2}m² × 1.3 circulation`
    : 'Not applicable (manual MHE)';
  XLSX.utils.book_append_sheet(wb, ws([
    ['AREA SUMMARY'],[],['Generated:',today],[],
    ['Area Component','Sub-component','Area (m²)','Area (sq ft)','% of Gross'],
    ['STORAGE ZONES','','','',''],
    ...Object.entries(zoneSummary).map(([z,v])=>{
      const a=+(netRackArea*(v.locs/(metrics.totLocs||1))).toFixed(0);
      return[ZONE_DEFS[z]?.label||z,'Racking + aisles',a,toSqFt(a),+(a/(wW*wL)*100).toFixed(1)+'%'];
    }),
    ['Total Storage Zones','',netRackArea,toSqFt(netRackArea),+(netRackArea/(wW*wL)*100).toFixed(1)+'%'],
    [],
    ['INBOUND STAGING','','','',''],
    ['','Buffer storage (pallets/boxes in dwell)',sb.inbStorage||'—',sb.inbStorage?toSqFt(sb.inbStorage):'—',''],
    ['','GRN apron (dock face)',sb.grnApron||'—',sb.grnApron?toSqFt(sb.grnApron):'—',''],
    ['Total Receiving','',receivingArea,toSqFt(receivingArea),+(receivingArea/(wW*wL)*100).toFixed(1)+'%'],
    [],
    ['OUTBOUND STAGING','','','',''],
    ['','Buffer storage (pallets/boxes in dwell)',sb.outStorage||'—',sb.outStorage?toSqFt(sb.outStorage):'—',''],
    ['','Packing benches area',sb.packingArea||0,toSqFt(sb.packingArea||0),''],
    ['','Dispatch apron (dock face)',sb.dispatchApron||'—',sb.dispatchApron?toSqFt(sb.dispatchApron):'—',''],
    ['Total Dispatch','',dispatchArea,toSqFt(dispatchArea),+(dispatchArea/(wW*wL)*100).toFixed(1)+'%'],
    [],
    ['MHE CHARGING AREA',mheLabel,mheArea||0,toSqFt(mheArea||0),mheArea?+(mheArea/(wW*wL)*100).toFixed(1)+'%':'0%'],
    ['OFFICE / WELFARE','Staff amenities, lockers',officeArea||50,toSqFt(officeArea||50),+((officeArea||50)/(wW*wL)*100).toFixed(1)+'%'],
    ['CIRCULATION','Aisles, emergency egress, columns',circulationArea||0,toSqFt(circulationArea||0),+(circulationArea/(wW*wL)*100).toFixed(1)+'%'],
    [],
    ['TOTAL GROSS FLOOR AREA','',wW*wL,toSqFt(wW*wL),'100%'],
    ['Warehouse Dimensions','',`${wW}m × ${wL}m`,`${(wW*3.281).toFixed(0)}ft × ${(wL*3.281).toFixed(0)}ft`,''],
  ],[28,36,12,16,12]),'6. Area Summary');

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
    ['MHE Charging Area', design.forkType==='manual'?'Not required (manual)':`${design.mheArea||0}m² (${design.nMHE} units × ${design.mheBayM2}m² × 1.3)`],
    ['Office / Welfare',design.officeArea+'m²'],
    ['Circulation (8% of racking)',design.circulationArea+'m²'],
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
  // Warehouse params
  const [clearH,   setClearH]  = useState('9');
  const [nMHE,     setNMHE]    = useState(''); // blank = auto-calculate
  const [dockSide, setDockSide]= useState('one');
  const [dockConfig,setDockConfig]=useState('shared');
  const [dockPitch, setDockPitch]=useState('4.5');
  const [forkType, setForkType]= useState('reach');
  const [aisleW,   setAisleW]  = useState('3.0');
  const [shifts,   setShifts]  = useState('1');
  // Inbound / outbound mode
  const [inboundMode,  setInboundMode]  = useState('pallets'); // 'pallets' | 'boxes'
  const [outboundMode, setOutboundMode] = useState('pallets'); // 'pallets' | 'boxes'
  // Inbound box sizes (up to 3)
  const [inbBoxSizes, setInbBoxSizes] = useState([
    { L:'', W:'', H:'', qtyPerDay:'' },
    { L:'', W:'', H:'', qtyPerDay:'' },
    { L:'', W:'', H:'', qtyPerDay:'' },
  ]);
  const [inbStackH, setInbStackH] = useState('3');
  // Outbound boxes config
  const [outbTruckType,    setOutbTruckType]    = useState('medium');
  const [outbTrucksPerDay, setOutbTrucksPerDay] = useState('');
  const [outbStackH,       setOutbStackH]       = useState('3');
  // Truck mix (inbound vehicles)
  const [truckMix, setTruckMix]= useState([
    { type:'medium', stagingDepth:'8', inboundVehicles:'5', outboundVehicles:'5', palletsPerTruck:'8' },
  ]);
  // Dwell times
  const [inboundDwellH,  setInboundDwellH]  = useState('4');
  const [outboundDwellH, setOutboundDwellH] = useState('2');
  // Packing
  const [packingInDispatch, setPackingInDispatch] = useState(true);
  const [packingBenches,    setPackingBenches]    = useState('4');

  // Data
  const [masterText, setMasterText] = useState('');
  const [orderText,  setOrderText]  = useState('');
  const [invText,    setInvText]    = useState('');

  // Results
  const [analysis,  setAnalysis]  = useState(null);
  const [rackConfig,setRackConfig]= useState(null);
  const [design,    setDesign]    = useState(null);
  const [configConfirmed,setConfigConfirmed]=useState(false);
  const [loading,   setLoading]   = useState(false);
  const [error,     setError]     = useState('');

  const params = {
    clearH:parseFloat(clearH)||9,
    nMHE,
    dockSide, dockConfig, dockPitch,
    forkType,
    aisleW:parseFloat(aisleW)||3.0,
    shifts:parseInt(shifts)||1,
    truckMix, inboundDwellH, outboundDwellH,
    packingInDispatch, packingBenches:parseInt(packingBenches)||0,
    inboundMode, outboundMode,
    inbBoxSizes, inbStackH,
    outbTruckType, outbTrucksPerDay, outbStackH,
  };
  // Truck mix helpers
  const addTruck = () => setTruckMix(m=>[...m,{type:'medium',stagingDepth:'8',inboundVehicles:'2',outboundVehicles:'2',palletsPerTruck:'8'}]);
  const removeTruck = i => setTruckMix(m=>m.filter((_,idx)=>idx!==i));
  const updateTruck = (i, field, val) => setTruckMix(m=>m.map((t,idx)=>idx===i?{...t,[field]:val}:t));
  const onTruckTypeChange = (i, type) => {
    const tt = TRUCK_TYPES[type]||TRUCK_TYPES.medium;
    setTruckMix(m=>m.map((t,idx)=>idx===i?{...t,type,
      stagingDepth:String(tt.stagingDepth),
      palletsPerTruck:String(tt.defaultPallets)}:t));
  };
  const addTruckDefault = () => setTruckMix(m=>[...m,
    {type:'medium',stagingDepth:'8',inboundVehicles:'2',outboundVehicles:'2',palletsPerTruck:'8'}]);
  const updateInbBox = (i,field,val) => setInbBoxSizes(s=>s.map((b,idx)=>idx===i?{...b,[field]:val}:b));

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
        const rc = generateRackConfig(a, params);
        setAnalysis(a); setRackConfig(rc);
        setDesign(null); setConfigConfirmed(false);
        // Preliminary design with estimated areas (before config confirmed)
        const d = calcWarehouseSize(a, params);
        setDesign(d);
      } catch(e) { setError(e.message); }
      setLoading(false);
    }, 100);
  };

  const confirmConfig = () => {
    if (!analysis || !rackConfig) return;
    const customAreas = rackAreasFromConfig(rackConfig);
    const d = calcWarehouseSize(analysis, params, customAreas);
    setDesign(d); setConfigConfirmed(true);
  };

  const updateCfgField = (id, field, val) => {
    setRackConfig(prev => prev.map(cfg => {
      if (cfg.id !== id) return cfg;
      let updated = { ...cfg, [field]: field==='orientation'||field==='tiers'
        ? val : (parseFloat(val)||cfg[field]) };
      // When tiers change, suggest a sensible default tierHeight
      if (field === 'tiers') {
        const t = parseInt(val)||1;
        const suggestedTierH = t === 1
          ? cfg.shelfH                                   // 1 tier = full height
          : Math.floor(cfg.shelfH / t / 100) * 100;     // divide evenly, round to 100mm
        updated = { ...updated, tierHeight: suggestedTierH };
      }
      return recalcCfg(updated);
    }));
    setConfigConfirmed(false); // needs reconfirm after edit
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

            {/* Building params */}
            <div style={{fontSize:'11px',fontWeight:'700',color:'#7c3aed',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'8px'}}>Building</div>
            <div style={S.grid2}>
              <div><label style={lbl}>Clear Height (m)</label>
                <input style={inp} type="number" min="4" max="20" step="0.5" value={clearH}
                  onChange={e=>setClearH(e.target.value)} placeholder="9"/></div>
              <div><label style={lbl}>Working Shifts</label>
                <select style={inp} value={shifts} onChange={e=>setShifts(e.target.value)}>
                  <option value="1">1 shift (8h/day)</option>
                  <option value="2">2 shifts (16h/day)</option>
                  <option value="3">3 shifts (24h/day)</option>
                </select></div>
              <div><label style={lbl}>No. of MHE Units {forkType==='manual'?'(manual — no charging)':'(blank = auto)'}</label>
                <input style={inp} type="number" min="0" value={nMHE}
                  onChange={e=>setNMHE(e.target.value)}
                  placeholder={forkType==='manual'?'N/A':'Auto-calculate'}
                  disabled={forkType==='manual'}/>
                {forkType!=='manual'&&<div style={{fontSize:'10px',color:'#9ca3af',marginTop:'2px'}}>
                  Auto: max(docks÷2, 1). Charging bay: {forkType==='counterbalance'?'14':forkType==='reach'?'9':'6'}m² per unit + 30%
                </div>}
              </div>
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
            </div>

            {/* Dock configuration */}
            <div style={{fontSize:'11px',fontWeight:'700',color:'#7c3aed',textTransform:'uppercase',letterSpacing:'0.05em',margin:'14px 0 8px'}}>Dock Configuration</div>
            <div style={S.grid2}>
              <div><label style={lbl}>Dock Wall Position</label>
                <select style={inp} value={dockSide} onChange={e=>setDockSide(e.target.value)}>
                  <option value="one">One side (south wall)</option>
                  <option value="both">Opposite sides (cross-dock)</option>
                  <option value="corner">Corner (south + east)</option>
                </select></div>
              <div><label style={lbl}>Dock Allocation</label>
                <select style={inp} value={dockConfig} onChange={e=>setDockConfig(e.target.value)}>
                  <option value="shared">Shared inbound + outbound</option>
                  <option value="separate">Separate inbound / outbound</option>
                </select></div>
              <div><label style={lbl}>Dock Pitch (centre-to-centre)</label>
                <select style={inp} value={dockPitch} onChange={e=>setDockPitch(e.target.value)}>
                  <option value="4.0">4.0m — Compact</option>
                  <option value="4.5">4.5m — Standard</option>
                  <option value="5.0">5.0m — Wide</option>
                </select></div>
            </div>
            <div style={{fontSize:'11px',color:'#9ca3af',marginTop:'4px'}}>
              Dock count is calculated from your vehicle mix below — not entered manually.
            </div>

            {/* ── INBOUND SECTION ──────────────────────────────────────── */}
            <div style={{fontSize:'11px',fontWeight:'700',color:'#0369a1',textTransform:'uppercase',letterSpacing:'0.05em',margin:'14px 0 8px'}}>
              ⬅ Inbound
            </div>

            {/* Inbound mode toggle */}
            <div style={{display:'flex',gap:'8px',marginBottom:'10px'}}>
              {[['pallets','📦 Pallets'],['boxes','📫 Boxes / Cases']].map(([v,l])=>(
                <button key={v} onClick={()=>setInboundMode(v)}
                  style={{flex:1,padding:'7px',borderRadius:'7px',fontSize:'12px',fontWeight:'700',
                    cursor:'pointer',
                    border:`2px solid ${inboundMode===v?'#0284c7':'#e2e8f0'}`,
                    background:inboundMode===v?'#e0f2fe':'#fff',
                    color:inboundMode===v?'#0369a1':'#6b7280'}}>
                  {l}
                </button>))}
            </div>

            {inboundMode === 'pallets' ? (<>
              {/* Truck mix table for pallets */}
              <div style={{border:'1px solid #e2e8f0',borderRadius:'8px',overflow:'hidden',marginBottom:'8px'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                  <thead><tr style={{background:'#f0f9ff'}}>
                    {['Truck Type','Staging Depth','Pallets/truck','Vehicles/day',''].map(h=>(
                      <th key={h} style={{padding:'6px 8px',textAlign:'left',fontWeight:'700',
                        fontSize:'10px',color:'#0369a1',textTransform:'uppercase',
                        borderBottom:'1px solid #e2e8f0'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {truckMix.map((t,i)=>(
                      <tr key={i} style={{background:i%2===0?'#fff':'#f8fafc'}}>
                        <td style={{padding:'5px 7px'}}>
                          <select value={t.type} onChange={e=>onTruckTypeChange(i,e.target.value)}
                            style={{...inp,marginBottom:0,fontSize:'11px',padding:'3px 5px',width:'100%'}}>
                            {Object.entries(TRUCK_TYPES).map(([k,v])=>(
                              <option key={k} value={k}>{v.label}</option>))}
                          </select>
                        </td>
                        <td style={{padding:'5px 7px'}}>
                          <input type="number" min="1" max="20" value={t.stagingDepth}
                            onChange={e=>updateTruck(i,'stagingDepth',e.target.value)}
                            style={{...inp,marginBottom:0,width:'52px',fontSize:'11px',padding:'3px 5px'}}/>
                          <span style={{fontSize:'10px',color:'#9ca3af',marginLeft:'2px'}}>m</span>
                        </td>
                        <td style={{padding:'5px 7px'}}>
                          <input type="number" min="1" value={t.palletsPerTruck}
                            onChange={e=>updateTruck(i,'palletsPerTruck',e.target.value)}
                            style={{...inp,marginBottom:0,width:'48px',fontSize:'11px',padding:'3px 5px',
                              background:'#fffbeb',border:'1px solid #fde68a'}}/>
                          <span style={{fontSize:'9px',color:'#d97706',marginLeft:'2px'}}>editable</span>
                        </td>
                        <td style={{padding:'5px 7px'}}>
                          <input type="number" min="0" value={t.inboundVehicles}
                            onChange={e=>updateTruck(i,'inboundVehicles',e.target.value)}
                            style={{...inp,marginBottom:0,width:'48px',fontSize:'11px',padding:'3px 5px'}}/>
                        </td>
                        <td style={{padding:'5px 7px',textAlign:'center'}}>
                          {truckMix.length>1&&(
                            <button onClick={()=>removeTruck(i)}
                              style={{background:'none',border:'none',color:'#be185d',
                                cursor:'pointer',fontSize:'15px',lineHeight:1}}>×</button>)}
                        </td>
                      </tr>))}
                  </tbody>
                </table>
              </div>
              <button onClick={addTruck}
                style={{fontSize:'11px',fontWeight:'600',color:'#0369a1',background:'#e0f2fe',
                  border:'1px dashed #7dd3fc',borderRadius:'6px',padding:'5px 12px',
                  cursor:'pointer',width:'100%',marginBottom:'4px'}}>
                + Add Truck Type
              </button>
              <div style={{fontSize:'10px',color:'#9ca3af',marginBottom:'4px'}}>
                Pallet size assumed: 1.2×1.2m = 1.44m² footprint
              </div>
            </>) : (<>
              {/* Box sizes for inbound */}
              <div style={{fontSize:'11px',color:'#374151',marginBottom:'8px',fontWeight:'600'}}>
                Enter box sizes received. Qty/day is total boxes of that size arriving daily.
              </div>
              {inbBoxSizes.map((b,i)=>(
                <div key={i} style={{background:i===0?'#f0f9ff':'#f8fafc',border:`1px solid ${i===0?'#bae6fd':'#e2e8f0'}`,
                  borderRadius:'8px',padding:'10px',marginBottom:'8px'}}>
                  <div style={{fontSize:'11px',fontWeight:'700',color:i===0?'#0369a1':'#6b7280',marginBottom:'6px'}}>
                    Box Size {i+1} {i===0?'(required)':'(optional)'}
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr 1fr',gap:'6px'}}>
                    {[['L (mm)','L'],['W (mm)','W'],['H (mm)','H'],['Qty / day','qtyPerDay']].map(([label,field])=>(
                      <div key={field}>
                        <div style={{fontSize:'10px',color:'#6b7280',marginBottom:'2px'}}>{label}</div>
                        <input type="number" min="0" value={b[field]}
                          onChange={e=>updateInbBox(i,field,e.target.value)}
                          style={{...inp,marginBottom:0,padding:'4px 6px',fontSize:'11px'}}
                          placeholder={field==='qtyPerDay'?'0':'mm'}/>
                      </div>))}
                  </div>
                </div>))}
              <div style={S.grid2}>
                <div><label style={lbl}>Stacking height in receiving</label>
                  <select style={inp} value={inbStackH} onChange={e=>setInbStackH(e.target.value)}>
                    <option value="1">1 layer (flat)</option>
                    <option value="2">2 layers</option>
                    <option value="3">3 layers</option>
                    <option value="4">4 layers</option>
                  </select></div>
              </div>
            </>)}

            <div style={S.grid2}>
              <div><label style={lbl}>Inbound dwell (before put-away)</label>
                <select style={inp} value={inboundDwellH} onChange={e=>setInboundDwellH(e.target.value)}>
                  <option value="2">2 hours</option>
                  <option value="4">4 hours</option>
                  <option value="8">8 hours (next shift)</option>
                  <option value="16">16 hours</option>
                  <option value="24">24 hours (next day)</option>
                </select></div>
            </div>

            {/* ── OUTBOUND SECTION ─────────────────────────────────────── */}
            <div style={{fontSize:'11px',fontWeight:'700',color:'#d97706',textTransform:'uppercase',letterSpacing:'0.05em',margin:'14px 0 8px'}}>
              ➡ Outbound
            </div>

            {/* Outbound mode toggle */}
            <div style={{display:'flex',gap:'8px',marginBottom:'10px'}}>
              {[['pallets','📦 Pallets'],['boxes','📫 Boxes / Cases']].map(([v,l])=>(
                <button key={v} onClick={()=>setOutboundMode(v)}
                  style={{flex:1,padding:'7px',borderRadius:'7px',fontSize:'12px',fontWeight:'700',
                    cursor:'pointer',
                    border:`2px solid ${outboundMode===v?'#d97706':'#e2e8f0'}`,
                    background:outboundMode===v?'#fffbeb':'#fff',
                    color:outboundMode===v?'#92400e':'#6b7280'}}>
                  {l}
                </button>))}
            </div>

            {outboundMode === 'pallets' ? (
              <div style={{background:'#fffbeb',border:'1px solid #fde68a',borderRadius:'8px',padding:'10px',marginBottom:'8px',fontSize:'12px',color:'#92400e'}}>
                📊 Outbound pallets/day calculated automatically from Order data (daily volume ÷ pallet volume 1.2×1.2×1.2m × 65% fill).
                Paste Order data in Step 3 for accurate numbers. You can also set outbound vehicles in the inbound truck mix table.
              </div>
            ) : (<>
              <div style={{background:'#fffbeb',border:'1px solid #fde68a',borderRadius:'8px',padding:'10px',marginBottom:'10px',fontSize:'12px',color:'#92400e'}}>
                📊 Outbound boxes/day calculated from Order data. Trucks needed = daily volume ÷ truck capacity. Paste Order data in Step 3.
              </div>
              <div style={S.grid2}>
                <div><label style={lbl}>Outbound truck type</label>
                  <select style={inp} value={outbTruckType} onChange={e=>setOutbTruckType(e.target.value)}>
                    {Object.entries(TRUCK_TYPES).map(([k,v])=>(
                      <option key={k} value={k}>{v.label}</option>))}
                  </select></div>
                <div><label style={lbl}>Trucks/day (blank = auto-calculate)</label>
                  <input style={inp} type="number" min="0" value={outbTrucksPerDay}
                    onChange={e=>setOutbTrucksPerDay(e.target.value)}
                    placeholder="Auto from order volume"/>
                  <div style={{fontSize:'10px',color:'#9ca3af',marginTop:'2px'}}>
                    Leave blank to calculate from Order data
                  </div>
                </div>
                <div><label style={lbl}>Stacking height in dispatch</label>
                  <select style={inp} value={outbStackH} onChange={e=>setOutbStackH(e.target.value)}>
                    <option value="1">1 layer</option>
                    <option value="2">2 layers</option>
                    <option value="3">3 layers</option>
                    <option value="4">4 layers</option>
                  </select></div>
              </div>
            </>)}

            <div style={S.grid2}>
              <div><label style={lbl}>Outbound dwell (before loading)</label>
                <select style={inp} value={outboundDwellH} onChange={e=>setOutboundDwellH(e.target.value)}>
                  <option value="1">1 hour</option>
                  <option value="2">2 hours</option>
                  <option value="4">4 hours</option>
                  <option value="8">8 hours (next shift)</option>
                </select></div>
            </div>

            {/* ── PACKING ──────────────────────────────────────────────── */}
            <div style={{fontSize:'11px',fontWeight:'700',color:'#7c3aed',textTransform:'uppercase',letterSpacing:'0.05em',margin:'14px 0 8px'}}>Packing / Value-Add</div>
            <div style={{display:'flex',gap:'8px',marginBottom:'8px'}}>
              {[['true','In dispatch area'],['false','Separate packing area']].map(([v,l])=>(
                <button key={v} onClick={()=>setPackingInDispatch(v==='true')}
                  style={{flex:1,padding:'7px 10px',borderRadius:'7px',fontSize:'11px',fontWeight:'600',
                    cursor:'pointer',border:`1px solid ${String(packingInDispatch)===v?'#7c3aed':'#e2e8f0'}`,
                    background:String(packingInDispatch)===v?'#f5f3ff':'#fff',
                    color:String(packingInDispatch)===v?'#7c3aed':'#6b7280'}}>
                  {l}
                </button>))}
            </div>
            <div style={S.grid2}>
              <div><label style={lbl}>No. of packing benches</label>
                <input style={inp} type="number" min="0" value={packingBenches}
                  onChange={e=>setPackingBenches(e.target.value)} placeholder="4"/>
                <div style={{fontSize:'10px',color:'#9ca3af',marginTop:'2px'}}>Each bench = 4m² incl. access</div>
              </div>
            </div>
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
                ['Gross Area', `${(design.wW*design.wL).toLocaleString()}m²\n(${Math.round(design.wW*design.wL*10.7639).toLocaleString()} sq ft)`, '#fef9c3','#854d0e'],
                ['Dock Doors (calc.)', `${design.inboundDocks} inb + ${design.outboundDocks} out = ${design.totalDocks}`, '#e0f2fe','#0369a1'],
                ['No-Movement SKUs', analysis.metrics.nmCount, '#fff1f2','#be185d'],
              ].map(([l,v,bg,col])=>(
                <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',textAlign:'center',border:`1px solid ${col}22`}}>
                  <div style={{fontSize:'16px',fontWeight:'800',color:col,lineHeight:1.2}}>{v}</div>
                  <div style={{fontSize:'10px',color:'#6b7280',marginTop:'4px',fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                </div>))}
            </div>

            {/* Staging breakdown */}
            {design.staging && (
              <div style={{...S.card,background:'#f0f9ff',border:'1px solid #bae6fd',marginBottom:'12px',padding:'14px 18px'}}>
                <div style={{fontWeight:'700',fontSize:'13px',color:'#0369a1',marginBottom:'10px'}}>
                  📦 Staging Area Breakdown
                </div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px',fontSize:'12px'}}>
                  <div style={{background:'#fff',borderRadius:'8px',padding:'10px'}}>
                    <div style={{fontWeight:'700',color:'#0284c7',marginBottom:'6px'}}>
                      ⬅ Inbound / Receiving — {design.receivingArea}m²
                    </div>
                    <div style={{color:'#374151',lineHeight:1.8,fontSize:'12px'}}>
                      <div>{design.staging.inbLabel==='boxes'?'Boxes':'Pallets'}/day: <strong>{design.staging.inbUnits?.toFixed(0)||0}</strong></div>
                      <div>In dwell ({params.inboundDwellH}h): <strong>{design.staging.inbPalletsInDwell?.toFixed(0)||'—'} {design.staging.inbLabel}</strong></div>
                      <div>Storage buffer: {design.staging.stagingBreakdown?.inbStorage}m²</div>
                      <div>GRN apron: {design.staging.stagingBreakdown?.grnApron}m²</div>
                      <div style={{marginTop:'4px',fontWeight:'700',color:'#0369a1'}}>Inbound docks: {design.inboundDocks}</div>
                    </div>
                  </div>
                  <div style={{background:'#fff',borderRadius:'8px',padding:'10px'}}>
                    <div style={{fontWeight:'700',color:'#d97706',marginBottom:'6px'}}>
                      ➡ Outbound / Dispatch — {design.dispatchArea}m²
                    </div>
                    <div style={{color:'#374151',lineHeight:1.8,fontSize:'12px'}}>
                      <div>{design.staging.outLabel==='boxes'?'Boxes':'Pallets'}/day: <strong>{design.staging.outUnits?.toFixed(0)||0}</strong>
                        {design.staging.outDailyVolM3>0&&<span style={{color:'#9ca3af',fontSize:'10px'}}> ({design.staging.outDailyVolM3}m³)</span>}
                      </div>
                      <div>In dwell ({params.outboundDwellH}h): <strong>{design.staging.outPalletsInDwell?.toFixed(0)||'—'} {design.staging.outLabel}</strong></div>
                      <div>Storage buffer: {design.staging.stagingBreakdown?.outStorage}m²</div>
                      <div>Packing area: {design.staging.stagingBreakdown?.packingArea}m²</div>
                      <div>Dispatch apron: {design.staging.stagingBreakdown?.dispatchApron}m²</div>
                      {design.staging.trucksNeeded>0&&<div style={{color:'#d97706',fontWeight:'600'}}>Trucks needed: {design.staging.trucksNeeded}</div>}
                      <div style={{marginTop:'4px',fontWeight:'700',color:'#d97706'}}>Outbound docks: {design.outboundDocks}</div>
                    </div>
                  </div>
                </div>
                <div style={{fontSize:'10px',color:'#0369a1',marginTop:'8px',fontStyle:'italic'}}>
                  Sizing: pallets in dwell × 1.2m² footprint × 1.5 safety + dock apron ({params.dockPitch}m pitch × 2m depth × docks)
                </div>
                {design.mheArea > 0 && (
                  <div style={{marginTop:'8px',background:'#fdf4ff',border:'1px solid #e9d5ff',
                    borderRadius:'6px',padding:'8px 12px',fontSize:'12px',color:'#7c3aed',fontWeight:'600'}}>
                    ⚡ MHE Charging Area: <strong>{design.mheArea}m²</strong>
                    {' '}({design.nMHE} {params.forkType} truck{design.nMHE>1?'s':''} × {design.mheBayM2}m² × 1.3 circulation factor)
                  </div>
                )}
              </div>
            )}

            {/* ── RACK CONFIGURATION EDITOR ───────────────────────────── */}
            {rackConfig && (
              <div style={{...S.card, marginBottom:'12px'}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'4px'}}>
                  <div style={{fontWeight:'700',fontSize:'14px',color:'#0f172a'}}>
                    🗄 Rack Configuration
                  </div>
                  <div style={{fontSize:'11px',color:configConfirmed?'#166534':'#d97706',
                    fontWeight:'700',background:configConfirmed?'#f0fdf4':'#fffbeb',
                    padding:'3px 10px',borderRadius:'99px',border:`1px solid ${configConfirmed?'#86efac':'#fde68a'}`}}>
                    {configConfirmed ? '✓ Confirmed' : '⚠ Edit then confirm'}
                  </div>
                </div>
                <div style={{fontSize:'12px',color:'#6b7280',marginBottom:'14px'}}>
                  Auto-generated from bin sizes. Both orientations shown — select the better one.
                  Edit bay dims or tiers, then click <strong>Confirm</strong>.
                </div>

                {rackConfig.map(cfg => {
                  const isShelving = ['shelving','liveStorage'].includes(cfg.rack);
                  const binD = cfg.binDims;
                  const minClearH = 5500; // mm min for 2-tier mezzanine
                  const canMezzanine = parseFloat(params.clearH)*1000 >= minClearH;
                  return (
                    <div key={cfg.id} style={{border:'1px solid #e2e8f0',borderRadius:'10px',
                      marginBottom:'10px',overflow:'hidden'}}>

                      {/* Header */}
                      <div style={{background:'#f8fafc',padding:'9px 14px',
                        borderBottom:'1px solid #e2e8f0',
                        display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                        <div>
                          <span style={{fontWeight:'700',fontSize:'13px',color:'#0f172a'}}>{cfg.rackName}</span>
                          <span style={{fontSize:'11px',color:'#6b7280',marginLeft:'8px'}}>
                            {cfg.binName} {binD ? `(${binD[0]}×${binD[1]}×${binD[2]}mm)` : ''}
                          </span>
                        </div>
                        <span style={{fontSize:'12px',fontWeight:'700',color:'#7c3aed'}}>
                          {cfg.locs.toLocaleString()} locations needed
                        </span>
                      </div>

                      <div style={{padding:'12px 14px'}}>
                        {/* Editable bay params */}
                        <div style={{display:'grid',
                          gridTemplateColumns: isShelving ? 'repeat(4,1fr)' : 'repeat(3,1fr)',
                          gap:'8px',marginBottom:'10px'}}>
                          <div>
                            <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'600',marginBottom:'3px',textTransform:'uppercase'}}>Bay Width (mm)</div>
                            <input type="number" min="600" max="3000" step="100"
                              value={cfg.bayW}
                              onChange={e=>updateCfgField(cfg.id,'bayW',e.target.value)}
                              style={{...{border:'1px solid #e2e8f0',borderRadius:'6px',padding:'5px 8px',
                                fontSize:'12px',width:'100%',boxSizing:'border-box'}} }/>
                          </div>
                          <div>
                            <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'600',marginBottom:'3px',textTransform:'uppercase'}}>Bay Depth (mm)</div>
                            <input type="number" min="300" max="2000" step="50"
                              value={cfg.bayD}
                              onChange={e=>updateCfgField(cfg.id,'bayD',e.target.value)}
                              style={{...{border:'1px solid #e2e8f0',borderRadius:'6px',padding:'5px 8px',
                                fontSize:'12px',width:'100%',boxSizing:'border-box'}} }/>
                          </div>
                          {isShelving && (<>
                            <div>
                              <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'600',marginBottom:'3px',textTransform:'uppercase'}}>Shelf Height (mm)</div>
                              <input type="number" min="1000" max="4500" step="100"
                                value={cfg.shelfH}
                                onChange={e=>updateCfgField(cfg.id,'shelfH',e.target.value)}
                                style={{...{border:'1px solid #e2e8f0',borderRadius:'6px',padding:'5px 8px',
                                  fontSize:'12px',width:'100%',boxSizing:'border-box'}} }/>
                            </div>
                            <div>
                              <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'600',marginBottom:'3px',textTransform:'uppercase'}}>Shelf Clearance (mm)</div>
                              <input type="number" min="20" max="200" step="10"
                                value={cfg.clearance}
                                onChange={e=>updateCfgField(cfg.id,'clearance',e.target.value)}
                                style={{...{border:'1px solid #e2e8f0',borderRadius:'6px',padding:'5px 8px',
                                  fontSize:'12px',width:'100%',boxSizing:'border-box'}} }/>
                            </div>
                          </>)}
                          {/* Tier height input — always shown for shelving */}
                          {isShelving && (
                            <div style={{gridColumn:'1 / -1'}}>
                              <div style={{fontSize:'10px',color:'#7c3aed',fontWeight:'700',marginBottom:'3px',
                                textTransform:'uppercase',display:'flex',alignItems:'center',gap:'6px'}}>
                                Height per Tier (mm)
                                {cfg.tiers > 1 && (
                                  <span style={{background:'#f5f3ff',border:'1px solid #c4b5fd',
                                    borderRadius:'99px',padding:'1px 7px',fontSize:'10px',color:'#7c3aed'}}>
                                    {cfg.tiers} tiers × {cfg.tierHeight||cfg.shelfH}mm = {((cfg.tiers*(cfg.tierHeight||cfg.shelfH))/1000).toFixed(2)}m total
                                    {(cfg.tiers*(cfg.tierHeight||cfg.shelfH)) > parseFloat(params.clearH)*1000
                                      ? <span style={{color:'#be185d',marginLeft:'4px'}}>⚠ exceeds clear height!</span>
                                      : <span style={{color:'#166534',marginLeft:'4px'}}>✓ fits in {params.clearH}m</span>}
                                  </span>
                                )}
                              </div>
                              <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
                                <input type="number" min="1000" max="4000" step="100"
                                  value={cfg.tierHeight||cfg.shelfH}
                                  onChange={e=>updateCfgField(cfg.id,'tierHeight',e.target.value)}
                                  style={{...{border:'2px solid #c4b5fd',borderRadius:'6px',padding:'5px 8px',
                                    fontSize:'13px',fontWeight:'700',width:'120px',boxSizing:'border-box',color:'#7c3aed'}} }/>
                                <span style={{fontSize:'11px',color:'#6b7280'}}>
                                  Usable height per tier for bins (excl. mezzanine structure)
                                </span>
                              </div>
                            </div>
                          )}
                          {!isShelving && (
                            <div>
                              <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'600',marginBottom:'3px',textTransform:'uppercase'}}>Pick Aisle (mm)</div>
                              <input type="number" min="1500" max="5000" step="100"
                                value={cfg.aisleW}
                                onChange={e=>updateCfgField(cfg.id,'aisleW',e.target.value)}
                                style={{...{border:'1px solid #e2e8f0',borderRadius:'6px',padding:'5px 8px',
                                  fontSize:'12px',width:'100%',boxSizing:'border-box'}} }/>
                            </div>
                          )}
                        </div>

                        {/* Orientation selector (shelving only) */}
                        {isShelving && binD && (
                          <div style={{marginBottom:'10px'}}>
                            <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'700',
                              textTransform:'uppercase',marginBottom:'6px'}}>
                              Bin Orientation in Bay
                            </div>
                            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px'}}>
                              {[
                                ['LW', `L(${binD[0]}) along width`, cfg.o1],
                                ['WL', `W(${binD[1]}) along width`, cfg.o2],
                              ].map(([orient, label, stats]) => (
                                <button key={orient}
                                  onClick={()=>updateCfgField(cfg.id,'orientation',orient)}
                                  style={{padding:'8px 10px',borderRadius:'8px',textAlign:'left',cursor:'pointer',
                                    border:`2px solid ${cfg.orientation===orient?'#7c3aed':'#e2e8f0'}`,
                                    background:cfg.orientation===orient?'#f5f3ff':'#fff'}}>
                                  <div style={{fontSize:'11px',fontWeight:'700',
                                    color:cfg.orientation===orient?'#7c3aed':'#374151',marginBottom:'3px'}}>
                                    {cfg.orientation===orient?'✓ ':''}Orientation {orient}: {label}
                                  </div>
                                  {stats && (
                                    <div style={{fontSize:'10px',color:stats.feasible?'#166534':'#be185d'}}>
                                      {stats.feasible
                                        ? `${stats.acrossW}×${stats.acrossD}×${stats.levels} = ${stats.locsPerBay}/bay`
                                        : '✗ Bin does not fit'}
                                    </div>
                                  )}
                                </button>))}
                            </div>
                          </div>
                        )}

                        {/* Multi-tier (shelving only) */}
                        {isShelving && (
                          <div style={{marginBottom:'10px'}}>
                            <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'700',
                              textTransform:'uppercase',marginBottom:'6px'}}>
                              Storage Tiers {!canMezzanine&&<span style={{color:'#be185d'}}>(need ≥5.5m clearH for mezzanine)</span>}
                            </div>
                            <div style={{display:'flex',gap:'8px'}}>
                              {[1,2,3].map(t=>(
                                <button key={t}
                                  onClick={()=>{
                                    if(t>1&&!canMezzanine) return;
                                    updateCfgField(cfg.id,'tiers',t);
                                  }}
                                  disabled={t>1&&!canMezzanine}
                                  style={{flex:1,padding:'7px 6px',borderRadius:'7px',cursor:t>1&&!canMezzanine?'not-allowed':'pointer',
                                    border:`2px solid ${cfg.tiers===t?'#7c3aed':'#e2e8f0'}`,
                                    background:cfg.tiers===t?'#f5f3ff':t>1&&!canMezzanine?'#f8fafc':'#fff',
                                    color:cfg.tiers===t?'#7c3aed':t>1&&!canMezzanine?'#d1d5db':'#374151',
                                    fontSize:'12px',fontWeight:'700'}}>
                                  {t === 1 ? '1 tier (ground)' : t===2 ? '2 tiers (mezzanine)' : '3 tiers'}
                                </button>))}
                            </div>
                            {cfg.tiers > 1 && (
                              <div style={{fontSize:'10px',color:'#7c3aed',marginTop:'4px'}}>
                                ↑ {cfg.tiers} tiers: same footprint, {cfg.tiers}× the storage capacity
                              </div>
                            )}
                          </div>
                        )}

                        {/* Result row */}
                        <div style={{background:cfg.feasible===false?'#fff1f2':'#f0fdf4',
                          borderRadius:'8px',padding:'9px 12px',
                          display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:'6px'}}>
                          <div style={{fontSize:'12px',color:cfg.feasible===false?'#be185d':'#166534',fontWeight:'700'}}>
                            {cfg.feasible===false ? '✗ Bin does not fit — adjust bay dimensions' : (
                              isShelving
                                ? `✓ ${cfg.acrossW} wide × ${cfg.acrossD} deep × ${cfg.levels} levels${cfg.tiers>1?` × ${cfg.tiers} tiers`:''} = ${cfg.locsPerBayTotal}/bay`
                                : `✓ ${cfg.acrossW||2} wide × ${cfg.acrossD||1} deep × ${cfg.levels} levels = ${cfg.locsPerBay}/bay`
                            )}
                          </div>
                          <div style={{display:'flex',gap:'16px',fontSize:'12px'}}>
                            <span><strong style={{color:'#7c3aed'}}>{cfg.baysNeeded}</strong> bays</span>
                            <span><strong style={{color:'#0369a1'}}>{((cfg.bayW/1000)*(cfg.bayD/1000)).toFixed(2)}m²</strong>/bay footprint</span>
                            <span><strong style={{color:'#059669'}}>{cfg.area}m²</strong> total area</span>
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}

                {/* Summary and confirm */}
                <div style={{background:'#f5f3ff',borderRadius:'8px',padding:'10px 14px',marginTop:'4px',
                  display:'flex',justifyContent:'space-between',alignItems:'center',gap:'12px',flexWrap:'wrap'}}>
                  <div style={{fontSize:'13px',color:'#374151'}}>
                    <strong>Total racking area:</strong>{' '}
                    <span style={{color:'#7c3aed',fontWeight:'700',fontSize:'15px'}}>
                      {rackConfig.reduce((s,c)=>s+(c.area||0),0).toFixed(0)}m²
                    </span>
                    <span style={{fontSize:'11px',color:'#9ca3af',marginLeft:'8px'}}>
                      across {rackConfig.reduce((s,c)=>s+(c.baysNeeded||0),0)} bays
                    </span>
                  </div>
                  <button onClick={confirmConfig}
                    style={{padding:'10px 24px',background:'linear-gradient(135deg,#7c3aed,#6d28d9)',
                      color:'#fff',border:'none',borderRadius:'9px',fontWeight:'800',fontSize:'14px',
                      cursor:'pointer',fontFamily:'inherit',
                      boxShadow:'0 4px 14px rgba(124,58,237,0.4)'}}>
                    ✓ Confirm & Generate Layout →
                  </button>
                </div>
              </div>
            )}

            {/* ── FLOOR PLAN (only after confirmed) ───────────────────── */}
            {configConfirmed && design && (<>
            <div style={S.card}>
              <div style={{fontWeight:'700',fontSize:'14px',color:'#0f172a',marginBottom:'12px'}}>
                🗺 Recommended Floor Layout
              </div>
              <FloorPlanSVG analysis={analysis} design={design} params={params} rackConfig={rackConfig}/>
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
              <button onClick={()=>exportExcel(analysis,design,params,rackConfig)}
                style={{flex:1,padding:'12px',background:'linear-gradient(135deg,#059669,#047857)',
                  color:'#fff',border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'14px',
                  cursor:'pointer',fontFamily:'inherit'}}>
                ⬇ Download Excel Report
              </button>
              <button onClick={()=>exportPPT(analysis,design,params,rackConfig)}
                style={{flex:1,padding:'12px',background:'linear-gradient(135deg,#7c3aed,#6d28d9)',
                  color:'#fff',border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'14px',
                  cursor:'pointer',fontFamily:'inherit'}}>
                📊 Download PPT Report
              </button>
            </div>
          </>)}
          {/* End configConfirmed block */}
          </>)}
        </div>
      </div>
    </div>
  );
}
