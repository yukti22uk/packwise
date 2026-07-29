// ─── WAREHOUSE DESIGNER TOOL ──────────────────────────────────────────────────
// Step 1: Warehouse parameters
// Step 2: Master SKU data (dimensions)
// Step 3: Order / Pick data (for velocity)
// Step 4: Inventory data (current stock)
// Outputs: SKU slotting, rack recommendations, warehouse sizing, SVG floor plan
import { useState, useEffect, useRef, useMemo } from 'react';
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

// ─── BIN / PALLET SIZE VARIANTS ───────────────────────────────────────────────
// Users can define up to MAX_BIN_VARIANTS different sizes for one bin band, each
// with its own quantity. Each variant becomes its own rack configuration and its
// own section in the floor plan.
// Shape: { XL:[{L,W,H,locs,label}, {…}, {…}], M:[{…}] }
const MAX_BIN_VARIANTS = 3;

// Split a pseudo bin key like "XL#v2" back to its base bin band
function baseBinOf(key) {
  if (!key) return key;
  const i = key.indexOf('#');
  return i < 0 ? key : key.slice(0, i);
}

// Expand a binSummary so each size variant becomes its own pseudo bin band
// ("XL#v1", "XL#v2", …). Used by the User-Defined flow, which keys its whole
// pipeline off binSummary. Returns the originals when nothing is multi-size.
function expandBinVariants(analysis, binOverrides) {
  const bs = (analysis && analysis.binSummary) || {};
  // Rebuild when any band has multiple sizes OR has been zeroed out, so the
  // returned summary always reflects the user's edits.
  // Any explicit override makes the rebuild authoritative: quantities, extra
  // sizes and zero-exclusions all land in the returned summary, so no caller
  // has to re-derive them from scale factors.
  let needs = false;
  Object.keys(bs).forEach(function(k){
    const a = binOverrides ? binOverrides[k] : null;
    if (Array.isArray(a) && a.length) needs = true;
  });
  if (!needs) return { analysis, binOverrides };

  const bs2 = {}, ov2 = {};
  Object.keys(bs).forEach(function(k){
    const vs = binVariantsFor(k, binOverrides, bs[k].locs || 0);
    if (!vs.length) return;                       // zeroed out -> drop the band
    if (vs.length === 1) {
      const v0 = vs[0];
      bs2[k] = { ...bs[k], locs: v0.locs };       // honour the edited quantity
      ov2[k] = [{ L:v0.phys?v0.phys[0]:'', W:v0.phys?v0.phys[1]:'',
                  H:v0.phys?v0.phys[2]:'', locs:v0.locs, label:v0.label }];
      return;
    }
    vs.forEach(function(v, i){
      const key = k + '#v' + (i + 1);
      bs2[key] = { ...bs[k], locs: v.locs,
        name: (bs[k].name || k) + ' - ' + v.label };
      ov2[key] = [{ L:v.phys?v.phys[0]:'', W:v.phys?v.phys[1]:'',
                    H:v.phys?v.phys[2]:'', locs:v.locs, label:v.label }];
    });
  });
  return { analysis: { ...analysis, binSummary: bs2 }, binOverrides: ov2 };
}

// Split a layout key like "selective__v2" back to its base rack type
function baseRackOf(key) {
  if (!key) return key;
  const i = key.indexOf('__');
  return i < 0 ? key : key.slice(0, i);
}

// Resolve the list of size variants for a bin band
function binVariantsFor(binKey, binOverrides, fallbackLocs) {
  const bk = baseBinOf(binKey);
  const base = BIN_CATALOG[bk] ? BIN_CATALOG[bk].phys : null;
  const arr  = binOverrides ? binOverrides[binKey] : null;
  if (Array.isArray(arr) && arr.length) {
    const out = [];
    arr.forEach(function(ov, i){
      const L = parseFloat(ov.L) || (base ? base[0] : 0);
      const W = parseFloat(ov.W) || (base ? base[1] : 0);
      const H = parseFloat(ov.H) || (base ? base[2] : 0);
      const n = parseFloat(ov.locs) || 0;
      if (n <= 0) return;   // quantity 0 means "exclude this size"
      out.push({
        idx: i,
        label: ov.label || ('Size ' + (i + 1)),
        phys: (L && W && H) ? [L, W, H] : base,
        locs: n,
      });
    });
    // An explicit override is authoritative. If the user zeroed every size we
    // must return an EMPTY list so the band is dropped — falling back to the
    // system quantity here would silently ignore the edit.
    return out;
  }
  // No override defined for this band at all -> use the system-generated figure
  return [{ idx:0, label:'Size 1', phys:base, locs: fallbackLocs || 0 }];
}

// Primary (first) physical size for a band — used where a single size is needed
function binPhysFor(binKey, binOverrides) {
  const v = binVariantsFor(binKey, binOverrides, 0);
  const bk = baseBinOf(binKey);
  return v[0] ? v[0].phys : (BIN_CATALOG[bk] ? BIN_CATALOG[bk].phys : null);
}

// Total requested locations across all variants ÷ system-generated locations
function binLocScales(analysis, binOverrides) {
  const scales = {};
  if (!binOverrides) return scales;
  Object.keys(binOverrides).forEach(function(k){
    const arr = binOverrides[k];
    if (!Array.isArray(arr) || !arr.length) return;
    let want = 0;
    arr.forEach(function(v){ want += parseFloat(v.locs) || 0; });
    const orig = (analysis && analysis.binSummary && analysis.binSummary[k])
      ? (analysis.binSummary[k].locs || 0) : 0;
    // Record the scale even when it is 0 — that means "exclude this band".
    if (orig > 0) scales[k] = want / orig;
  });
  return scales;
}

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
  ground:     { name:'Ground Location',       bayW:1.2, bayD:1.2,  desc:'Odd-shaped/bulky items on floor' },
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
// Fast TSV parser — avoids re-allocating large arrays
function parseTSV(text) {
  if (!text) return [];
  const lines = text.trim().split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const cells = lines[i].split('\t').map(c => c.trim());
    if (cells.some(c => c)) out.push(cells);
  }
  return out;
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

// Select optimal bin considering preferred types and fill efficiency
function selectBin(skuL, skuW, skuH, volCm3, isLong, preferredBins) {
  if (isLong) return 'LONG';
  const ALL_BINS = ['XS','S','M','L','XL'];
  // Use preferred bins if specified, else all bins
  const preferred = (preferredBins && preferredBins.length > 0)
    ? ALL_BINS.filter(b => preferredBins.includes(b))
    : ALL_BINS;

  if (skuL > 0 && skuW > 0 && skuH > 0) {
    // 1st pass: smallest PREFERRED bin that physically fits
    const fitPreferred = preferred.filter(b => fitsInBin(skuL, skuW, skuH, b));
    if (fitPreferred.length > 0) return fitPreferred[0];
    // 2nd pass: any bin that fits (fallback when preferred set can't accommodate)
    const fitAny = ALL_BINS.filter(b => fitsInBin(skuL, skuW, skuH, b));
    return fitAny[0] || 'XL';
  }
  // Volume-only fallback — prefer from preferred set
  if (!volCm3) return preferred[0] || 'S';
  const VOL = { XS:500, S:3000, M:15000, L:50000, XL:Infinity };
  for (const b of preferred) {
    if (volCm3 <= (VOL[b]||Infinity)) return b;
  }
  return preferred[preferred.length-1] || 'XL';
}

// ─── RACK ELEVATION SVG ───────────────────────────────────────────────────────
function RackElevationSVG({ cfg, W: svgW=260, H: svgH=190 }) {
  const rack    = cfg.rack;
  const levels  = cfg.levels  || 4;
  const tiers   = parseInt(cfg.tiers)||1;
  const bayWmm  = cfg.bayW    || 2700;
  const tierHmm = parseFloat(cfg.tierHeight)||cfg.shelfH||2200;
  const clearMm = cfg.clearance||50;
  const binD    = cfg.binDims; // [L,W,H] mm
  const acrossW = cfg.acrossW || 2;
  const palLevH = 1500; // pallet + beam mm
  const totalHmm= rack==='selective'||rack==='driveIn'||rack==='doubleDeep'
    ? palLevH*levels + 400
    : tierHmm * tiers + 200;

  // Colours
  const STEEL='#475569', BEAM='#94a3b8', PALLET='#d97706',
        BIN='#3b82f6', DIM='#be185d', LABEL='#374151',
        GROUND='#1e293b', MEZZANINE='#7c3aed';

  // Layout margins
  const ML=40,MR=36,MT=22,MB=28;
  const DW=svgW-ML-MR, DH=svgH-MT-MB;
  const sX = DW/bayWmm;
  const sY = DH/(totalHmm||1);

  // Coordinate helpers (Y is flipped: 0=ground=bottom)
  const px  = mm => ML + mm*sX;
  const py  = mm => MT + (totalHmm-mm)*sY;
  const pw  = mm => Math.max(1,mm*sX);
  const ph  = mm => Math.max(1,mm*sY);
  const upW = Math.max(3, pw(60)); // upright width in px

  // Dimension annotation
  const dimH = (x1,y1,x2,y2,label,side='right') => {
    const mx=(x1+x2)/2, my=(y1+y2)/2;
    return (
      <g>
        <line x1={x1} y1={y1} x2={x2} y2={y2} stroke={DIM} strokeWidth="1"/>
        <line x1={x1-4} y1={y1} x2={x1+4} y2={y1} stroke={DIM} strokeWidth="1"/>
        <line x1={x2-4} y1={y2} x2={x2+4} y2={y2} stroke={DIM} strokeWidth="1"/>
        <text x={mx+(side==='right'?4:-4)} y={my} fontSize="7" fill={DIM}
          dominantBaseline="middle" textAnchor={side==='right'?'start':'end'}
          transform={`rotate(-90,${mx+(side==='right'?4:-4)},${my})`}>{label}</text>
      </g>
    );
  };
  const dimW = (x1,x2,y,label) => (
    <g>
      <line x1={x1} y1={y} x2={x2} y2={y} stroke={DIM} strokeWidth="1"/>
      <line x1={x1} y1={y-4} x2={x1} y2={y+4} stroke={DIM} strokeWidth="1"/>
      <line x1={x2} y1={y-4} x2={x2} y2={y+4} stroke={DIM} strokeWidth="1"/>
      <text x={(x1+x2)/2} y={y-5} fontSize="7" fill={DIM} textAnchor="middle">{label}</text>
    </g>
  );

  // ── SHELVING / LIVE STORAGE (front elevation) ─────────────────────────────
  if (rack==='shelving'||rack==='liveStorage') {
    const binH  = binD?binD[2]:200;
    const slotH = binH + clearMm;
    const levPT = Math.max(1,Math.floor(tierHmm/slotH)); // levels per tier
    const bW    = Math.max(4, pw((bayWmm-120)/Math.max(1,acrossW)) - 2);

    return (
      <svg width={svgW} height={svgH} viewBox={`0 0 ${svgW} ${svgH}`} style={{display:'block'}}>
        <rect width={svgW} height={svgH} fill="#f8fafc" rx="6"/>
        {/* Left + right uprights */}
        <rect x={px(0)} y={py(totalHmm-200)} width={upW} height={ph(totalHmm-200)} fill={STEEL}/>
        <rect x={px(bayWmm)-upW} y={py(totalHmm-200)} width={upW} height={ph(totalHmm-200)} fill={STEEL}/>
        {/* Tiers & levels */}
        {Array.from({length:tiers},(_,tier)=>
          Array.from({length:levPT},(_,lev)=>{
            const shY = tier*tierHmm + lev*slotH;
            return (
              <g key={`${tier}-${lev}`}>
                {/* Shelf board */}
                <rect x={px(60)} y={py(shY+slotH)-ph(28)} width={pw(bayWmm-120)} height={ph(28)} fill={BEAM}/>
                {/* Bins */}
                {Array.from({length:acrossW},(_,b)=>{
                  const bx=px(70+b*(bayWmm-140)/Math.max(1,acrossW));
                  const bHpx=Math.max(4,ph(binH*0.88));
                  return(
                    <rect key={b} x={bx} y={py(shY+slotH)-ph(28)-bHpx}
                      width={bW} height={bHpx}
                      fill={rack==='liveStorage'?'#60a5fa':'#3b82f6'}
                      rx="1" opacity="0.9"/>
                  );
                })}
              </g>
            );
          })
        )}
        {/* Mezzanine decks */}
        {tiers>1&&Array.from({length:tiers-1},(_,t)=>(
          <g key={t}>
            <rect x={px(0)} y={py((t+1)*tierHmm)-ph(60)} width={pw(bayWmm)} height={ph(60)} fill="#c4b5fd" opacity="0.7"/>
            <text x={px(bayWmm/2)} y={py((t+1)*tierHmm)-ph(30)} textAnchor="middle" fontSize="7" fill={MEZZANINE} fontWeight="700">MEZZANINE</text>
          </g>
        ))}
        {/* Live-storage: incline arrow */}
        {rack==='liveStorage'&&(
          <text x={px(bayWmm/2)} y={py(tierHmm/2)} textAnchor="middle" fontSize="9" fill="#1d4ed8">→ FIFO</text>
        )}
        {/* Ground line */}
        <line x1={px(0)} y1={py(0)} x2={px(bayWmm)} y2={py(0)} stroke={GROUND} strokeWidth="2.5"/>
        {/* Dimensions */}
        {dimW(px(0),px(bayWmm),MT-4,`${(bayWmm/1000).toFixed(1)}m`)}
        {dimH(px(bayWmm)+10,py(0),px(bayWmm)+10,py(totalHmm-200),`${(tierHmm*tiers/1000).toFixed(1)}m`)}
        <text x={svgW/2} y={svgH-6} textAnchor="middle" fontSize="8" fontWeight="700" fill={LABEL}>
          Shelving — {levPT*tiers} levels{tiers>1?` × ${tiers} tiers`:''}
        </text>
      </svg>
    );
  }

  // ── SELECTIVE PALLET RACK (front elevation) ───────────────────────────────
  if (rack==='selective') {
    const palH=1200, beamH=200, palW=1200;
    return (
      <svg width={svgW} height={svgH} viewBox={`0 0 ${svgW} ${svgH}`} style={{display:'block'}}>
        <rect width={svgW} height={svgH} fill="#f8fafc" rx="6"/>
        {/* Upright frames (double line) */}
        {[0, bayWmm].map((x,i)=>(
          <g key={i}>
            <rect x={px(x)-(i>0?upW:0)} y={py(palLevH*levels+100)} width={upW} height={ph(palLevH*levels+100)} fill={STEEL}/>
          </g>
        ))}
        <rect x={px(bayWmm/2)-upW/2} y={py(palLevH*levels+100)} width={upW} height={ph(palLevH*levels+100)} fill={STEEL} opacity="0.5"/>
        {/* Beam levels + pallets */}
        {Array.from({length:levels},(_,lev)=>{
          const beamY=lev*palLevH, palY=beamY+beamH;
          return (
            <g key={lev}>
              {/* Beam */}
              <rect x={px(0)} y={py(beamY+beamH)-ph(beamH)} width={pw(bayWmm)} height={ph(beamH)} fill={BEAM} rx="1"/>
              {/* 2 pallets per level */}
              {[0,1].map(p=>{
                const pxStart=px(40+p*(bayWmm/2-20));
                return(
                  <g key={p}>
                    {/* Pallet board */}
                    <rect x={pxStart} y={py(palY+palH)-ph(120)} width={pw(palW*0.85)} height={ph(120)} fill='#92400e' rx="1"/>
                    {/* Load on pallet */}
                    <rect x={pxStart+1} y={py(palY+palH)-ph(120)-ph(palH-200)} width={pw(palW*0.85)-2} height={ph(palH-200)} fill={PALLET} rx="2" opacity="0.85"/>
                  </g>
                );
              })}
            </g>
          );
        })}
        <line x1={px(0)} y1={py(0)} x2={px(bayWmm)} y2={py(0)} stroke={GROUND} strokeWidth="2.5"/>
        {dimW(px(0),px(bayWmm),MT-4,`${(bayWmm/1000).toFixed(1)}m`)}
        {dimH(px(bayWmm)+10,py(0),px(bayWmm)+10,py(palLevH*levels),`${(palLevH*levels/1000).toFixed(1)}m`)}
        <text x={svgW/2} y={svgH-6} textAnchor="middle" fontSize="8" fontWeight="700" fill={LABEL}>
          Selective Pallet Rack — {levels} beam levels
        </text>
      </svg>
    );
  }

  // ── DRIVE-IN RACK (side elevation — shows depth) ──────────────────────────
  if (rack==='driveIn') {
    const depthMm=cfg.bayD||6600, palW=1100, palH=1200, beamH=150;
    const nDeep=Math.max(1,Math.round(depthMm/palW));
    const sideX=DW/depthMm, totalSideH=palLevH*levels+300;
    const sideY=DH/(totalSideH||1);
    const sx=mm=>ML+mm*sideX, sy=mm=>MT+(totalSideH-mm)*sideY;

    return (
      <svg width={svgW} height={svgH} viewBox={`0 0 ${svgW} ${svgH}`} style={{display:'block'}}>
        <rect width={svgW} height={svgH} fill="#f8fafc" rx="6"/>
        {/* Side columns */}
        <rect x={sx(0)} y={sy(totalSideH-300)} width={pw(60)} height={ph(totalSideH-300)} fill={STEEL}/>
        <rect x={sx(depthMm)-pw(60)} y={sy(totalSideH-300)} width={pw(60)} height={ph(totalSideH-300)} fill={STEEL}/>
        {/* Levels */}
        {Array.from({length:levels},(_,lev)=>(
          <g key={lev}>
            {/* Rail beams (side view: horizontal lines) */}
            <rect x={sx(60)} y={sy(lev*palLevH+beamH)-ph(beamH)} width={sx(depthMm-60)-sx(60)} height={ph(beamH)} fill={BEAM} rx="1" opacity="0.7"/>
            {/* Pallets deep */}
            {Array.from({length:nDeep},(_,d)=>(
              <g key={d}>
                <rect x={sx(d*palW+70)+1} y={sy(lev*palLevH+palH+beamH)-ph(palH)} width={sx((d+1)*palW+70)-sx(d*palW+70)-2} height={ph(palH)} fill={PALLET} rx="1" opacity="0.8"/>
                <rect x={sx(d*palW+70)} y={sy(lev*palLevH+beamH)-ph(beamH/2)} width={sx((d+1)*palW+70)-sx(d*palW+70)} height={ph(beamH/2)} fill='#92400e' rx="0"/>
              </g>
            ))}
          </g>
        ))}
        {/* Loading arrow */}
        <text x={sx(depthMm/2)} y={sy(-60)} textAnchor="middle" fontSize="8" fill="#1d4ed8" fontWeight="700">← LOAD / UNLOAD</text>
        <line x1={sx(0)} y1={sy(0)} x2={sx(depthMm)} y2={sy(0)} stroke={GROUND} strokeWidth="2.5"/>
        {dimW(sx(0),sx(depthMm),MT-4,`${(depthMm/1000).toFixed(1)}m deep (${nDeep} pallets)`)}
        {dimH(sx(depthMm)+10,sy(0),sx(depthMm)+10,sy(palLevH*levels),`${(palLevH*levels/1000).toFixed(1)}m`)}
        <text x={svgW/2} y={svgH-6} textAnchor="middle" fontSize="8" fontWeight="700" fill={LABEL}>
          Drive-In Rack — {nDeep} deep × {levels} levels (LIFO)
        </text>
      </svg>
    );
  }

  // ── DOUBLE-DEEP RACK (side elevation) ────────────────────────────────────
  if (rack==='doubleDeep') {
    const depthMm=cfg.bayD||2400, palW=1100, palH=1200, beamH=150;
    const sideX=DW/depthMm, totalSideH=palLevH*levels+300;
    const sideY=DH/(totalSideH||1);
    const sx=mm=>ML+mm*sideX, sy=mm=>MT+(totalSideH-mm)*sideY;

    return (
      <svg width={svgW} height={svgH} viewBox={`0 0 ${svgW} ${svgH}`} style={{display:'block'}}>
        <rect width={svgW} height={svgH} fill="#f8fafc" rx="6"/>
        <rect x={sx(0)} y={sy(totalSideH-300)} width={pw(60)} height={ph(totalSideH-300)} fill={STEEL}/>
        <rect x={sx(depthMm)-pw(60)} y={sy(totalSideH-300)} width={pw(60)} height={ph(totalSideH-300)} fill={STEEL}/>
        {Array.from({length:levels},(_,lev)=>(
          <g key={lev}>
            <rect x={sx(60)} y={sy(lev*palLevH+beamH)-ph(beamH)} width={sx(depthMm-60)-sx(60)} height={ph(beamH)} fill={BEAM} rx="1" opacity="0.7"/>
            {[0,1].map(d=>(
              <g key={d}>
                <rect x={sx(d*palW+70)+1} y={sy(lev*palLevH+palH+beamH)-ph(palH)}
                  width={sx((d+1)*palW+70)-sx(d*palW+70)-2} height={ph(palH)}
                  fill={d===0?PALLET:'#f59e0b'} rx="1" opacity="0.85"/>
                <rect x={sx(d*palW+70)} y={sy(lev*palLevH+beamH)-ph(beamH/2)}
                  width={sx((d+1)*palW+70)-sx(d*palW+70)} height={ph(beamH/2)} fill='#92400e'/>
              </g>
            ))}
          </g>
        ))}
        {/* Reach truck arm indicator */}
        <rect x={sx(0)-20} y={sy(palLevH/2+palH)} width={20} height={ph(100)} fill="#64748b" rx="2"/>
        <text x={sx(0)-10} y={sy(palLevH/2+palH)+ph(50)} textAnchor="middle" fontSize="7" fill="#64748b"
          transform={`rotate(-90,${sx(0)-10},${sy(palLevH/2+palH)+ph(50)})`}>REACH</text>
        <line x1={sx(0)} y1={sy(0)} x2={sx(depthMm)} y2={sy(0)} stroke={GROUND} strokeWidth="2.5"/>
        {dimW(sx(0),sx(depthMm),MT-4,`${(depthMm/1000).toFixed(1)}m (2 pallets deep)`)}
        {dimH(sx(depthMm)+10,sy(0),sx(depthMm)+10,sy(palLevH*levels),`${(palLevH*levels/1000).toFixed(1)}m`)}
        <text x={svgW/2} y={svgH-6} textAnchor="middle" fontSize="8" fontWeight="700" fill={LABEL}>
          Double-Deep — 2 pallets deep × {levels} levels
        </text>
      </svg>
    );
  }

  // ── CANTILEVER RACK (front elevation) ─────────────────────────────────────
  if (rack==='cantilever') {
    const spineW=150, armLen=1000, armH=80, armStep=600;
    const totalW=spineW+armLen*2, totalHC=armStep*levels+400;
    const cX=DW/totalW, cY=DH/totalHC;
    const cx=mm=>ML+mm*cX, cy=mm=>MT+(totalHC-mm)*cY;

    return (
      <svg width={svgW} height={svgH} viewBox={`0 0 ${svgW} ${svgH}`} style={{display:'block'}}>
        <rect width={svgW} height={svgH} fill="#f8fafc" rx="6"/>
        {/* Central spine */}
        <rect x={cx(armLen)} y={cy(totalHC-400)} width={cx(armLen+spineW)-cx(armLen)} height={cy(0)-cy(totalHC-400)+2} fill={STEEL}/>
        {/* Arms per level */}
        {Array.from({length:levels},(_,lev)=>{
          const armY=200+lev*armStep;
          return (
            <g key={lev}>
              {/* Left arm */}
              <rect x={cx(0)} y={cy(armY+armH)-cy(armY)} width={cx(armLen)-cx(0)} height={Math.max(4,cy(armY)-cy(armY+armH))} fill={BEAM} rx="2"/>
              {/* Right arm */}
              <rect x={cx(armLen+spineW)} y={cy(armY+armH)-cy(armY)} width={cx(totalW)-cx(armLen+spineW)} height={Math.max(4,cy(armY)-cy(armY+armH))} fill={BEAM} rx="2"/>
              {/* Long items on arms */}
              <rect x={cx(0)+2} y={cy(armY+armH)-cy(armY)-ph(80)} width={cx(armLen)-cx(0)-4} height={ph(80)} fill="#fb923c" rx="1" opacity="0.85"/>
              <rect x={cx(armLen+spineW)+2} y={cy(armY+armH)-cy(armY)-ph(80)} width={cx(totalW)-cx(armLen+spineW)-4} height={ph(80)} fill="#fb923c" rx="1" opacity="0.85"/>
            </g>
          );
        })}
        <line x1={cx(0)} y1={cy(0)} x2={cx(totalW)} y2={cy(0)} stroke={GROUND} strokeWidth="2.5"/>
        {dimW(cx(0),cx(totalW),MT-4,`${(totalW/1000).toFixed(1)}m span`)}
        {dimH(cx(totalW)+10,cy(0),cx(totalW)+10,cy(armStep*levels),`${(armStep*levels/1000).toFixed(1)}m`)}
        <text x={svgW/2} y={svgH-6} textAnchor="middle" fontSize="8" fontWeight="700" fill={LABEL}>
          Cantilever Rack — {levels} arm levels (long goods)
        </text>
      </svg>
    );
  }

  // ── FALLBACK ─────────────────────────────────────────────────────────────
  return (
    <svg width={svgW} height={svgH} viewBox={`0 0 ${svgW} ${svgH}`} style={{display:'block'}}>
      <rect width={svgW} height={svgH} fill="#f8fafc" rx="6"/>
      <text x={svgW/2} y={svgH/2} textAnchor="middle" fontSize="11" fill={LABEL}>
        {cfg.rackName}
      </text>
    </svg>
  );
}

// ─── BIN CONSOLIDATION ────────────────────────────────────────────────────────
// Merge minority bin types into adjacent bins to reduce variety
const LOC_SHARE_MAP = { VF:1, F:1, M:1, S:2, VS:4, NM:8 };

function consolidateBins(slotted) {
  const SEQ = ['XS','S','M','L','XL'];
  const counts = {};
  slotted.filter(s=>s.stock>0&&SEQ.includes(s.bin)).forEach(s=>{
    counts[s.bin] = (counts[s.bin]||0)+1;
  });
  const total = Object.values(counts).reduce((a,b)=>a+b,0)||1;
  const used  = SEQ.filter(b=>counts[b]>0);
  const report = [];

  if (used.length <= 2) return { slotted, report }; // already minimal

  // Minority = bin used by < 5% of SKUs AND < 15 absolute SKUs
  const THRESH = Math.max(15, total * 0.05);
  const minority = used.filter(b => counts[b] <= THRESH && b !== 'XL');

  if (minority.length === 0) return { slotted, report };

  const updated = slotted.map(s => {
    if (!minority.includes(s.bin)||s.isLong||!SEQ.includes(s.bin)) return s;
    const idx = SEQ.indexOf(s.bin);
    for (let i = idx+1; i < SEQ.length; i++) {
      const target = SEQ[i];
      if (minority.includes(target) && i < SEQ.length-1) continue; // skip if target is also minority
      const fits = (s.L>0&&s.W>0&&s.H>0) ? fitsInBin(s.L,s.W,s.H,target) : true;
      if (!fits) continue;
      const newUpb  = unitsPerBin(s.L, s.W, s.H, s.volCm3, target);
      const share   = Math.min(LOC_SHARE_MAP[s.vb]||1, Math.max(1,newUpb));
      const rawLocs = s.stock > 0 ? Math.max(1, Math.ceil(s.stock/newUpb)) : 0;
      const newLocs = rawLocs > 0 ? Math.max(1, Math.ceil(rawLocs/share)) : 0;
      return { ...s, bin:target, binName:BIN_CATALOG[target]?.name||target,
        sb:target, upb:newUpb, locsReq:newLocs, consolidatedFrom:s.bin };
    }
    return s;
  });

  // Build consolidation report
  minority.forEach(from => {
    const moved = updated.filter(s=>s.consolidatedFrom===from);
    if (moved.length === 0) return;
    const byTarget = {};
    moved.forEach(s=>{ byTarget[s.bin]=(byTarget[s.bin]||0)+1; });
    report.push({ from, fromName:BIN_CATALOG[from]?.name||from,
      actions:Object.entries(byTarget).map(([to,n])=>
        ({to, toName:BIN_CATALOG[to]?.name||to, n})),
      totalMoved:moved.length });
  });

  return { slotted:updated, report };
}

function unitsPerBin(skuL, skuW, skuH, skuVolCm3, band) {
  const b = BIN_CATALOG[band];
  if (!b || !b.volCm3 || !skuVolCm3) return 1;
  if (b.phys && skuL > 0 && skuW > 0 && skuH > 0) {
    const [bL, bW, bH] = b.phys;
    const d = [skuL, skuW, skuH];
    // Try all 6 orientations of the SKU — pick best packing
    const ORIENTS = [[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];
    const byLayout = Math.max(1,
      ORIENTS.reduce((best, [x,y,z]) =>
        Math.max(best, Math.floor(bL/d[x]) * Math.floor(bW/d[y]) * Math.floor(bH/d[z]))
      , 0)
    );
    const byVolume = Math.max(1, Math.floor(b.volCm3 * b.fill / skuVolCm3));
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
      outDailyVolM3 = analysis?.dailyOutboundVolM3||0;
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

// ─── QUANTITY-AWARE BIN SELECTION ────────────────────────────────────────────
// Target: find the SMALLEST bin where all stock fits in ≤ BIN_LOC_TARGET locations.
// If nothing achieves the target, use the LARGEST fitting bin (fewest locations).
// This ensures high-qty SKUs get bigger bins; small bins only for low-qty SKUs.
const BIN_LOC_TARGET = 2; // aim for stock to sit in ≤ 2 locations per SKU

function selectBinByQty(skuL, skuW, skuH, volCm3, stock, preferredBins) {
  const ALL  = ['XS','S','M','L','XL'];
  const pref = preferredBins && preferredBins.length > 0
    ? ALL.filter(b => preferredBins.includes(b)) : ALL;

  // No stock → smallest fitting bin (dimension-based only)
  if (!stock || stock <= 0)
    return selectBin(skuL, skuW, skuH, volCm3, false, preferredBins);

  // Physically fitting candidates
  let cands;
  if (skuL > 0 && skuW > 0 && skuH > 0) {
    cands = pref.filter(b => fitsInBin(skuL, skuW, skuH, b));
    if (!cands.length) cands = ALL.filter(b => fitsInBin(skuL, skuW, skuH, b));
  } else {
    // No dims — volume fallback
    const VOL = {XS:500,S:3000,M:15000,L:50000,XL:Infinity};
    const src  = pref.length ? pref : ALL;
    if (!volCm3) return src[0]||'S';
    return src.find(b=>volCm3<=(VOL[b]||Infinity))||src[src.length-1]||'XL';
  }
  if (!cands.length) return 'XL';
  if (cands.length === 1) return cands[0];

  // Score each candidate by locations needed
  const scored = cands.map(b => {
    const upb  = unitsPerBin(skuL, skuW, skuH, volCm3, b);
    const locs = upb > 0 ? Math.ceil(stock / upb) : 9999;
    return { b, upb, locs };
  }).filter(s => s.upb > 0);
  if (!scored.length) return cands[0];

  // Pass 1 — smallest bin where locs <= BIN_LOC_TARGET
  const pass1 = scored.filter(s => s.locs <= BIN_LOC_TARGET);
  if (pass1.length) return pass1[0].b; // smallest achieving target

  // Pass 2 — no bin achieves target → minimum locations (largest useful bin)
  const minL = Math.min(...scored.map(s => s.locs));
  return scored.find(s => s.locs === minL)?.b || cands[0];
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
// Merge multiple runAnalysis results (from chunked processing) into one
function mergeAnalysisChunks(chunks) {
  if (!chunks.length) return null;
  if (chunks.length === 1) return chunks[0];
  const base = chunks[0];
  for (let i = 1; i < chunks.length; i++) {
    const ch = chunks[i];
    // Merge slotted arrays
    base.slotted.push(...ch.slotted);
    // Merge binSummary
    Object.entries(ch.binSummary||{}).forEach(([k,v])=>{
      if (!base.binSummary[k]) base.binSummary[k]={skus:0,locs:0,stock:0,capacity:0,name:v.name,upgrades:0};
      base.binSummary[k].skus     += v.skus;
      base.binSummary[k].locs     += v.locs;
      base.binSummary[k].stock    += v.stock;
      base.binSummary[k].capacity += v.capacity;
      base.binSummary[k].upgrades += v.upgrades;
    });
    // Merge zoneSummary
    Object.entries(ch.zoneSummary||{}).forEach(([k,v])=>{
      if (!base.zoneSummary[k]) base.zoneSummary[k]={skus:0,locs:0,area:0};
      base.zoneSummary[k].skus += v.skus;
      base.zoneSummary[k].locs += v.locs;
      base.zoneSummary[k].area += v.area;
    });
    // Merge rackSummary
    Object.entries(ch.rackSummary||{}).forEach(([k,v])=>{
      if (!base.rackSummary[k]) base.rackSummary[k]={skus:0,locs:0,bays:0,area:0};
      base.rackSummary[k].skus += v.skus;
      base.rackSummary[k].locs += v.locs;
      base.rackSummary[k].bays += v.bays||0;
      base.rackSummary[k].area += v.area||0;
    });
    // Merge metrics
    base.metrics.totSKUs  += ch.metrics.totSKUs;
    base.metrics.totLocs  += ch.metrics.totLocs;
    base.metrics.totStock += ch.metrics.totStock;
    base.totalQtyUpgrades += ch.totalQtyUpgrades||0;
    // Merge matrix (velocity × size counts)
    Object.entries(ch.matrix||{}).forEach(([vb,row])=>{
      if (!base.matrix[vb]) base.matrix[vb]={};
      Object.entries(row).forEach(([sb,n])=>{
        base.matrix[vb][sb]=(base.matrix[vb][sb]||0)+n;
      });
    });
  }
  // Recompute dailyOutbound metrics from merged slotted data
  base.metrics.totLocs = base.slotted.reduce((s,r)=>s+(r.locsReq||0),0);
  return base;
}

function runAnalysis(masterRows, orderRows, inventoryRows, params, preferredBins) {
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

  // SKU universe: all master + order SKUs (each SKU is processed exactly ONCE,
  // even across chunks). Location calculation is inventory-gated via locsReq below.
  const allSkus = new Set([...Object.keys(master), ...Object.keys(pickMap)]);

  // Velocity classification uses the same full set
  const allOrderSkus = allSkus;
  const items = [...allOrderSkus].map(sku => ({ sku, pickLines: pickMap[sku]||0 }));
  const velocityMap = classifyVelocity(items);

  // Build per-SKU slotting data
  const slotted = [];
  allSkus.forEach(sku => {
    const m  = master[sku] || { L:0,W:0,H:0, volCm3:0, maxDim:0, isLong:false, sb:'S' };
    const vb = velocityMap[sku] || 'NM';
    const sb = m.sb; // size-band for velocity×size matrix (volume-based, unchanged)
    const isLong = m.isLong;
    const zone  = getZone(vb, isLong);
    const rack  = selectRackType(sb, vb, isLong, clearH, forkType);
    const stock = invMap[sku]||0; // stock loaded FIRST so bin selection is qty-aware
    // Quantity-aware bin selection:
    // High-qty SKU → larger bin (fit all stock in ≤ 2 locations)
    // Low-qty SKU → smallest fitting bin
    const bin       = isLong ? 'LONG'
      : selectBinByQty(m.L, m.W, m.H, m.volCm3, stock, preferredBins);
    // Track if bin was upgraded from dim-only selection due to qty
    const dimOnlyBin = isLong ? 'LONG'
      : selectBin(m.L, m.W, m.H, m.volCm3, isLong, preferredBins);
    const qtyUpgraded = !isLong && bin !== dimOnlyBin;
    const upb  = isLong ? 1 : unitsPerBin(m.L, m.W, m.H, m.volCm3, bin);
    // Slow-mover pick-face sharing:
    const LOC_SHARE = { VF:1, F:1, M:1, S:2, VS:4, NM:8 };
    const share   = Math.min(LOC_SHARE[vb]||1, Math.max(1, upb));
    const rawLocs = stock > 0 ? Math.max(1, Math.ceil(stock/upb)) : 0;
    const locsReq = rawLocs > 0 ? Math.max(1, Math.ceil(rawLocs/share)) : 0;
    const pl    = pickMap[sku]||0;
    slotted.push({ sku, ...m, vb, sb, isLong, zone, rack, bin,
      upb, stock, locsReq, pickLines:pl, qtyUpgraded,
      dimOnlyBin: qtyUpgraded ? dimOnlyBin : null,
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
  const hasInv         = Object.keys(invMap).length > 0;
  const totSKUsMaster  = Object.keys(master).length;
  const totSKUs        = hasInv
    ? slotted.filter(r=>r.stock>0).length   // only inventory SKUs
    : slotted.length;                        // no inventory → all master SKUs
  const totLocs    = slotted.reduce((s,r)=>s+r.locsReq,0);
  const totStock   = slotted.reduce((s,r)=>s+r.stock,0);
  const longCount  = slotted.filter(r=>r.isLong).length;
  const nmCount    = slotted.filter(r=>r.vb==='NM'&&r.stock>0).length;
  const nmStock    = slotted.filter(r=>r.vb==='NM').reduce((s,r)=>s+r.stock,0);

  // ── Consolidation pass — reduce bin variety ───────────────────────────────
  const { slotted: consolidatedSlotted, report: binConsolidation } = consolidateBins(slotted);
  const finalSlotted = consolidatedSlotted;
  // Recompute zone/rack summaries on consolidated data
  const zoneSum2 = {}; Object.keys(ZONE_DEFS).forEach(z=>{
    const rows=finalSlotted.filter(r=>r.zone===z);
    zoneSum2[z]={ skus:rows.length, locs:rows.reduce((s,r)=>s+r.locsReq,0),
      stock:rows.reduce((s,r)=>s+r.stock,0), pickLines:rows.reduce((s,r)=>s+r.pickLines,0) };
  });
  const rackSum2 = {};
  finalSlotted.forEach(r=>{ if(!rackSum2[r.rack]) rackSum2[r.rack]={locs:0,skus:0};
    rackSum2[r.rack].locs+=r.locsReq; rackSum2[r.rack].skus+=1; });
  const binSummary = {};
  let totalQtyUpgrades = 0;
  finalSlotted.filter(s=>s.stock>0).forEach(s=>{
    if(!binSummary[s.bin]) binSummary[s.bin]={skus:0,locs:0,stock:0,capacity:0,name:s.binName,upgrades:0};
    binSummary[s.bin].skus++;
    binSummary[s.bin].locs     += s.locsReq;
    binSummary[s.bin].stock    += s.stock;
    binSummary[s.bin].capacity += s.locsReq * s.upb;
    if (s.qtyUpgraded) { binSummary[s.bin].upgrades++; totalQtyUpgrades++; }
  });
  Object.values(binSummary).forEach(b => {
    b.utilPct = b.capacity > 0 ? Math.round(b.stock / b.capacity * 100) : 0;
  });

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

  const fTotLocs  = finalSlotted.reduce((s,r)=>s+r.locsReq,0);
  const fTotStock = finalSlotted.reduce((s,r)=>s+r.stock,0);
  return { slotted:finalSlotted, matrix, zoneSummary:zoneSum2, rackSummary:rackSum2,
    binSummary, binConsolidation, totalQtyUpgrades,
    metrics: { totSKUs:finalSlotted.length, totLocs:fTotLocs, totStock:fTotStock,
      longCount, nmCount, nmStock,
      totSKUsMaster, hasInv },
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
    // Bays arranged back-to-back sharing aisles → each bay "owns" half an aisle
    const aisleA  = baysNeeded*(bayW/1000)*(aisleMm/1000/2);
    return { ...cfg, ...r, o1, o2, locsPerBayTotal, baysNeeded,
      area: +(baysNeeded*bayFP + aisleA).toFixed(1) };
  } else {
    // Pallet racks: keep locations-per-bay consistent with the bay width and the
    // ACTUAL pallet footprint, so edits to either side stay in sync.
    let c2 = cfg;
    if (cfg.binDims && ['selective','driveIn','doubleDeep'].includes(cfg.rack)) {
      const pw     = Math.min(cfg.binDims[0], cfg.binDims[1]);
      const across = Math.max(1, Math.floor((parseFloat(cfg.bayW)||0) / (pw + 100)));
      const lv     = Math.max(1, parseInt(cfg.levels)||1);
      const dp     = Math.max(1, parseInt(cfg.acrossD)||1);
      c2 = { ...cfg, acrossW:across, locsPerBay:across*lv*dp };
    }
    const baysNeeded = c2.locsPerBay>0 ? Math.ceil(c2.locs/c2.locsPerBay) : 0;
    const bayFP  = (c2.bayW/1000)*(c2.bayD/1000);
    const aisleM = (c2.aisleW||3000)/1000;
    // Shared aisle model: each bay owns half the aisle on one side
    const area   = +(baysNeeded*bayFP + baysNeeded*(c2.bayW/1000)*(aisleM/2)).toFixed(1);
    return { ...c2, baysNeeded, area };
  }
}

// Auto-generate rack config from analysis
function generateRackConfig(analysis, params, binOverrides) {
  const { clearH, forkType, aisleW } = params;
  const shelfMaxH = Math.min(3500, Math.floor(clearH*1000 - 300));
  const maxLift   = { manual:2200, counterbalance:6000, reach:9000, vna:12000 };
  const liftH     = maxLift[forkType]||6000;
  const aisleWmm  = Math.floor(parseFloat(aisleW)*1000);

  // Group by rack + bin
  const groups = {};
  (analysis?.slotted||[]).forEach(r => {
    const key = `${r.rack}|${r.bin}`;
    if (!groups[key]) groups[key] = {
      rack:r.rack, bin:r.bin, rackName:r.rackName,
      binName:r.binName, locs:0 };
    groups[key].locs += r.locsReq;
  });

  // Expand each rack|bin group into ONE CONFIG PER SIZE VARIANT.
  // If a bin is served by more than one rack type, each variant quantity is
  // apportioned by that rack type's share of the bin's total locations.
  const out = [];
  Object.values(groups).filter(g=>g.locs>0).forEach(g => {
    const binTotal = (analysis && analysis.binSummary && analysis.binSummary[g.bin])
      ? (analysis.binSummary[g.bin].locs || g.locs) : g.locs;
    const share = binTotal > 0 ? (g.locs / binTotal) : 1;
    const variants = binVariantsFor(g.bin, binOverrides, g.locs);
    const multi = variants.length > 1;

    variants.forEach((v, vi) => {
      const suffix = multi ? ('__v' + (vi+1)) : '';
      const meta = {
        id: g.rack + '|' + g.bin + suffix,
        rack: g.rack, bin: g.bin, rackName: g.rackName,
        binName: multi ? (g.binName + ' - ' + v.label) : g.binName,
        locs: Math.max(1, Math.round(v.locs * share)),
        layoutKey: g.rack + suffix,
        variantLabel: v.label, variantIdx: vi, variantCount: variants.length,
        binDims: v.phys,
      };
      out.push(buildOneCfg(meta, v.phys, g.rack, params, shelfMaxH, liftH, aisleWmm));
    });
  });
  return out;
}

// Build ONE rack configuration for a single bin size on a single rack type
function buildOneCfg(meta, binDims, rack, params, shelfMaxH, liftH, aisleWmm) {
  const clearH = params.clearH;

  if (['shelving','liveStorage'].includes(rack) && binDims) {
    const bL = binDims[0], bW = binDims[1];
    const stdBayWidths = [1800, 1500, 1200, 900];
    const bayW = stdBayWidths.find(bw => bw % Math.min(bL,bW) === 0) ||
                 stdBayWidths.find(bw => Math.floor(bw/Math.min(bL,bW)) >= 2) || 900;
    const bayD = Math.max(Math.max(bL,bW)+50, 400);
    const clearance = 50;
    const o1 = tryShelfOrientation(binDims,bayW,bayD,shelfMaxH,clearance,'LW');
    const o2 = tryShelfOrientation(binDims,bayW,bayD,shelfMaxH,clearance,'WL');
    const bestOrient = o1.locsPerBay>=o2.locsPerBay ? 'LW' : 'WL';
    const best = bestOrient==='LW' ? o1 : o2;
    return recalcCfg({ ...meta,
      bayW, bayD, shelfH:shelfMaxH, clearance, tierHeight:shelfMaxH,
      orientation:bestOrient, tiers:1, shelvingAisle:SHELVING_AISLE_MM,
      locsPerBay:best.locsPerBay, o1, o2, aisleW:SHELVING_AISLE_MM,
      ...best });

  } else if (['selective','driveIn','doubleDeep'].includes(rack)) {
    // Bay geometry derived from the ACTUAL pallet footprint
    const pl = binDims ? Math.max(binDims[0], binDims[1]) : 1200; // depth into rack
    const pw = binDims ? Math.min(binDims[0], binDims[1]) : 1000; // face width
    const ph = binDims ? binDims[2] : 1200;                        // load height
    const depth   = rack==='driveIn'?6 : rack==='doubleDeep'?2 : 1;
    const acrossW = 2;
    const sideGap = 100;
    const bayW_mm = acrossW*pw + (acrossW+1)*sideGap;
    const bayD_mm = depth*pl + 50;
    const levelPitch = Math.max(600, ph + 200);
    const levels = Math.max(1,
      Math.floor((Math.min(liftH, clearH*1000) - 800) / levelPitch));
    return recalcCfg({ ...meta,
      bayW:bayW_mm, bayD:bayD_mm, levels,
      locsPerBay: acrossW*levels*depth, tiers:1, orientation:'std',
      acrossW, acrossD:depth, aisleW:aisleWmm });

  } else if (rack==='cantilever') {
    const iL = binDims ? Math.max(binDims[0], binDims[1]) : 2500;
    const iH = binDims ? binDims[2] : 400;
    const levelPitch = Math.max(300, iH + 150);
    const levels  = Math.max(1, Math.floor((clearH*1000-500)/levelPitch));
    const bayW_mm = 1500;
    const bayD_mm = Math.max(1000, iL + 200);
    const faceW   = binDims ? Math.min(binDims[0], binDims[1]) : 750;
    const acrossW = Math.max(1, Math.floor(bayW_mm / (faceW + 50)));
    return recalcCfg({ ...meta,
      bayW:bayW_mm, bayD:bayD_mm, levels, locsPerBay:acrossW*levels,
      acrossW, acrossD:1, tiers:1, orientation:'std', aisleW:3000 });

  } else {
    // Ground / floor storage — footprint is the bin itself, stacked if height allows
    const gL = binDims ? Math.max(binDims[0], binDims[1]) : 1200;
    const gW = binDims ? Math.min(binDims[0], binDims[1]) : 1000;
    const gH = binDims ? binDims[2] : 1200;
    const stackLimit = Math.max(1,
      Math.floor((Math.min(liftH, clearH*1000) - 500) / Math.max(200, gH)));
    const levels = Math.min(4, stackLimit);
    return recalcCfg({ ...meta,
      bayW:gW+100, bayD:gL+100, levels, locsPerBay:levels,
      acrossW:1, acrossD:1, tiers:1, orientation:'std', aisleW:aisleWmm });
  }
}

// Keep the headline level counts on `design` consistent with the actual rack
// config (levels are derived from the real bin/pallet heights, which the user
// may have customised).
function syncDesignLevels(design, rackConfig) {
  if (!design || !rackConfig || !rackConfig.length) return design;
  const lv = (types) => {
    const set = rackConfig.filter(c => types.indexOf(c.rack) >= 0)
      .map(c => parseInt(c.levels) || 0).filter(n => n > 0);
    return set.length ? Math.max.apply(null, set) : 0;
  };
  const pl = lv(['selective','driveIn','doubleDeep']);
  const sl = lv(['shelving','liveStorage']);
  const out = { ...design };
  if (pl) out.palletLevels = pl;
  if (sl) out.shelfLevels  = sl;
  return out;
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
function calcWarehouseSize(analysis, params, customRackAreas, customZoneAreas) {
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
      // Correct aisle model: AISLE_FACTOR uses shared aisle (halved)
      const correctFactor = 1 + (parseFloat(aisleW)||3.0) / 2 / rd.bayD;
      ra[rk] = +(bays*rd.bayW*rd.bayD*correctFactor).toFixed(1);
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
  // Priority: customZoneAreas (exact, from uCfgs) > customRackAreas (rack-type mapped) > system zoneSummary
  const RACK_TO_ZONE_CWS = {
    shelving:'golden', liveStorage:'golden',
    selective:'reserve', doubleDeep:'reserve',
    driveIn:'bulk', cantilever:'long', ground:'bulk',
  };
  const zoneAreas = {};
  if (customZoneAreas) {
    // Exact zone areas from user rack config (no system zones bleed through)
    Object.entries(customZoneAreas).forEach(([z, a]) => { zoneAreas[z] = +a.toFixed(1); });
  } else if (customRackAreas) {
    Object.entries(customRackAreas).forEach(([rack, area]) => {
      const zone = RACK_TO_ZONE_CWS[rack] || 'golden';
      zoneAreas[zone] = +((zoneAreas[zone]||0) + area).toFixed(1);
    });
  } else {
    const totalLocs = analysis?.metrics?.totLocs || 1;
    Object.entries(analysis?.zoneSummary||{}).forEach(([z, zs]) => {
      zoneAreas[z] = +(netRackArea * (zs.locs/totalLocs)).toFixed(1);
    });
  }

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
// Module-level section layout calculator (outside FloorPlanSVG to avoid minifier TDZ)
// Module-level constants (outside FloorPlanSVG to avoid minifier TDZ)
const CROSS_AISLE_W_M = 3.0; // cross aisle width in metres
// Minimum bays per face per column — keeps sections visually meaningful
const MIN_BPF_BY_TYPE = {
  shelving:3, liveStorage:3, selective:7, doubleDeep:7,
  driveIn:5, cantilever:5, ground:5,
};
function computeSectionLayout(totalBays, sectionW, bayHm, colSlot, crossIntervalM, rackType) {
  var minBpf = (MIN_BPF_BY_TYPE[rackType]||1);
  // Max columns that physically fit
  var maxColsByWidth = Math.max(3, Math.floor(sectionW / colSlot));
  // With minBpf: cap nCols so baysPerFace >= minBpf
  var bpfAtMax = Math.max(minBpf, totalBays>0?Math.max(1,Math.ceil(totalBays/2/maxColsByWidth)):minBpf);
  var nColsMin  = totalBays>0?Math.max(3,Math.ceil(totalBays/2/bpfAtMax)):maxColsByWidth;
  var nCols = Math.min(maxColsByWidth, nColsMin);
  var baysPerFace = totalBays>0?Math.max(minBpf,Math.ceil(totalBays/2/nCols)):minBpf;
  var y = 0.3, yStor = 0, baysSinceLast = 0;
  var cYs = [];
  for(var b = 0; b < baysPerFace; b++){
    y += bayHm; yStor += bayHm; baysSinceLast++;
    var baysRemain = baysPerFace - b - 1;
    if(yStor >= crossIntervalM && baysSinceLast >= 2 && baysRemain >= 2){
      cYs.push(y); y += CROSS_AISLE_W_M; yStor = 0; baysSinceLast = 0;
    }
  }
  var exactH = y + 0.3;
  return { nCols, baysPerCol: baysPerFace, minBpf,
    height: Math.max(baysPerFace * bayHm + 0.6, exactH),
    area: +(exactH * sectionW).toFixed(1), crossYPositions: cYs };
}




// ── buildFloorPlanLayout: all computation extracted from FloorPlanSVG ──────────
// Module-level = separate esbuild scope, no TDZ conflicts with rendering code.
function buildFloorPlanLayout(design, params, rackConfig, analysis, fullscreen) {
  if (!design || !design.wW || !design.wL) return null;
  var MFT  = 3.2808;
  var M2FT = 10.7639;
  var ft   = m  => `${(m*MFT).toFixed(0)}'`;
  var sqft = m2 => `${Math.round(m2*M2FT).toLocaleString()} sq ft`;
  var m2lbl= (m2,label) => label ? `${label}\n${m2}m²\n(${sqft(m2)})` : `${m2}m²\n(${sqft(m2)})`;

  var { wW, wL, zoneAreas, receivingArea, dispatchArea, mheArea, officeArea,
    totalDocks, inboundDocks, outboundDocks, staging, netRackArea } = design;
  var { dockSide, aisleW:aisleWParam, dockPitch, forkType,
    packingBenches, inboundMode, outboundMode } = params;
  var pitch   = parseFloat(dockPitch)||4.5;
  var aisleM  = parseFloat(aisleWParam)||3.0;

  // ── USE LAYOUT-DERIVED DIMENSIONS ────────────────────────────────────────
  // actualWW starts as wW; overridden after sectionLayouts computed (line ~1346)
  var actualWW = wW;   // will be updated once sectionLayouts are built
  var actualWL = wL;   // will be updated after zone heights computed
  var SVG_W = fullscreen ? 1800 : 960;
  var SVG_H = fullscreen
    ? Math.max(1800, Math.round(1800 * (wL/wW)))
    : 820;  // rebuilt after actualWL known
  var ML=62, MR=70, MT=50, MB=70;
  var DW=SVG_W-ML-MR, DH=SVG_H-MT-MB;
  var sX=DW/wW, sY=DH/wL;
  var X=m=>ML+m*sX, Y=m=>MT+m*sY, W=m=>m*sX, H=m=>m*sY;

  // ── AREA HEIGHTS ────────────────────────────────────────────────────────────
  // Provisional band heights (recomputed from actualWW once the width is final)
  var recH    = Math.max(4,(receivingArea||0)/wW);
  var disH    = Math.max(4,(dispatchArea||0)/wW);
  var stagingH= Math.max(recH,disH);
  // Use non-ground rack width for support area height calculation
  // (actualWW computed later, use wW as proxy for now, updated post-compute)
  var areaW = Math.max(wW, 10);
  var offH    = Math.max(3,(officeArea||50)/areaW);
  var mheH    = mheArea>0 ? Math.max(2,mheArea/areaW) : 0;
  var supportH= offH + mheH;

  var isOne   = dockSide==='one';
  var isBoth  = dockSide==='both';
  var isCorner= dockSide==='corner';

  var zonesH  = isOne
    ? Math.max(0,wL-stagingH-supportH)
    : Math.max(0,wL-recH-disH-supportH);

  // Zone vertical allocation
  var zoneOrder=['golden','mid','reserve','bulk','long'];
  var totZA = zoneOrder.reduce((s,z)=>s+(zoneAreas[z]||0),0)||1;
  var zH    = {};
  zoneOrder.forEach(z=>{ zH[z]=((zoneAreas[z]||0)/totZA)*zonesH; });

  // Build zone rects (from north going south).
  // NOTE: the vertical stack is laid out AFTER actualWW is known, because every
  // staging/support band height is area/width and the width is rack-driven.
  var zoneRects=[], stagingRects=[], supportRects=[];
  let cur=0;

  // Rack-type sections — each rack type gets its own dedicated band
  // (replaces velocity zones in the physical 2D layout)
  var RACK_TYPE_STYLE={
    shelving:   {label:'Shelving', color:'#eff6ff', border:'#93c5fd', text:'#1d4ed8'},
    liveStorage:{label:'Flow / Live Storage', color:'#f0fdf4', border:'#86efac', text:'#166534'},
    selective:  {label:'Selective Pallet Rack', color:'#fefce8', border:'#fde047', text:'#854d0e'},
    doubleDeep: {label:'Double-Deep Rack', color:'#f5f3ff', border:'#c4b5fd', text:'#6d28d9'},
    driveIn:    {label:'Drive-In Rack', color:'#1e293b', border:'#334155', text:'#f1f5f9'},
    cantilever: {label:'Cantilever Rack', color:'#fff7ed', border:'#fdba74', text:'#c2410c'},
    ground:     {label:'Ground Storage', color:'#fef3c7', border:'#fbbf24', text:'#92400e'},
  };
  // Bay height in plan = actual bayW from rackConfig (converted from mm to m)
  // Fallback to standard dimensions if rackConfig not available
  var BAY_HEIGHT_M_LOOKUP={shelving:0.9,liveStorage:1.5,selective:2.7,
    doubleDeep:2.7,driveIn:2.7,cantilever:1.5,ground:1.2};

  // Build rack-type → bayW (m) map from actual rackConfig data
  var LK=function(cfg){ return cfg.layoutKey || cfg.rack; };
  var rackBayWidthM={};
  (rackConfig||[]).forEach(cfg=>{
    var k=LK(cfg);
    if(k && cfg.bayW && !rackBayWidthM[k]){
      rackBayWidthM[k]=parseFloat(cfg.bayW)/1000; // mm → m
    }
  });
  var RACK_INFO_2D_LOOKUP={
    shelving:   {depth:1.0},liveStorage:{depth:1.2},selective:{depth:2.2},
    doubleDeep: {depth:4.4},driveIn:    {depth:5.5},cantilever:{depth:2.0},ground:{depth:2.4},
  };
  // Order rack types sensibly: manual pick first (near dispatch), then pallet, then bulk
  var RACK_BASE_ORDER=['shelving','liveStorage','selective','doubleDeep','driveIn','cantilever','ground'];
  // Expand to one entry per size variant, keeping the base ordering
  var RACK_ORDER=(function(){
    var seen={}, ordered=[];
    RACK_BASE_ORDER.forEach(function(base){
      (rackConfig||[]).forEach(function(cfg){
        if(cfg.rack!==base) return;
        var k=LK(cfg);
        if(!seen[k]){ seen[k]=1; ordered.push(k); }
      });
      if(!seen[base]){ seen[base]=1; ordered.push(base); }
    });
    return ordered;
  })();
  // Variant label per layout key (for section headings)
  var rackVariantLabel={};
  (rackConfig||[]).forEach(function(cfg){
    var k=LK(cfg);
    if(cfg.variantCount>1 && !rackVariantLabel[k]){
      var d=cfg.binDims;
      rackVariantLabel[k]=cfg.variantLabel+(d?(' '+d[0]+'x'+d[1]+'x'+d[2]):'');
    }
  });
  // Group rackConfig by layout key
  var rackTypeAreas={};
  (rackConfig||[]).forEach(cfg=>{
    var k=LK(cfg);
    rackTypeAreas[k]=(rackTypeAreas[k]||0)+(cfg.area||0);
  });
  // Fallback from slotted data
  if(!Object.keys(rackTypeAreas).length){
    (analysis?.slotted||[]).forEach(s=>{
      rackTypeAreas[s.rack]=(rackTypeAreas[s.rack]||0)+0.5;
    });
  }
  var storageHTotal=Object.values(rackTypeAreas).reduce((s,a)=>s+a,0)/wW||10;
  // Build zone sections bay-first
  var sectionLayouts = {}; // rackType → { nCols, baysPerCol, height, area, crossYPositions }
  var sectionCrossAisles=[];  // cross aisles between rack type sections + staging boundary
  var SECTION_CA_W=3.0;        // 3m cross aisle between sections
  var firstRackSection=true;

  // ── HELPER: compute one rack section ────────────────────────────────────────
  var computeSection = function(rt, sectionW) {
    var totalBaysRt = (rackConfig||[]).filter(c=>LK(c)===rt).reduce((s,c)=>s+(c.baysNeeded||0),0);
    if(!totalBaysRt && !(rackTypeAreas?.[rt])) return null;
    var rtB     = baseRackOf(rt);
    var rtRi    = (RACK_INFO_2D_LOOKUP?.[rtB]) || {depth:2.2};
    var rtPa    = (rackAisleM?.[rt])||(DEFAULT_AISLE_M?.[rtB])||aisleM||1.2;
    var rtFaceD = (rackBayDepthM?.[rt]) || rtRi.depth/2 || 1.0;
    var rtGap   = (rtB==="shelving"||rtB==="liveStorage")?0.05:0.10;
    var rtSlot  = rtFaceD*2 + rtGap + rtPa;
    var rtBayH  = (rackBayWidthM?.[rt])||(BAY_HEIGHT_M_LOOKUP?.[rtB])||0.9;
    var rtCross = ({shelving:13,liveStorage:13,selective:27,
      doubleDeep:27,driveIn:27,cantilever:27,ground:27})[rtB] || 13;
    var rtBays  = totalBaysRt || Math.ceil(((rackTypeAreas?.[rt])||0) / (rtBayH * wW));
    var rtLayout= computeSectionLayout(rtBays, sectionW, rtBayH, rtSlot, rtCross, rtB);
    return {rt, rtLayout, rtBays, rtSlot, rtPa, rtFaceD, rtGap, rtStyle:(RACK_TYPE_STYLE?.[rtB])||{label:rtB,color:'#f8fafc',border:'#e2e8f0',text:'#374151'}};
  };

  // ── PASS 1: non-ground sections → derive actualWW ───────────────────────────
  RACK_ORDER.filter(rt=>baseRackOf(rt)!=='ground').forEach(rt=>{
    var s=computeSection(rt, wW);
    if(!s) return;
    // requiredW EXACTLY matches rackRowsForZone column positions:
    // curX = 0.3 + rtPa/2; rx[i] = curX + i*rtSlot + rtPa/2
    // rx[0] = 0.3 + rtPa (full leading aisle)
    // rx[nCols-1]+rtColDepth = 0.3 + rtPa + (nCols-1)*rtSlot + rtColDepth
    //                        = 0.3 + nCols*rtColDepth + nCols*rtPa
    // For trailing = rtPa: zone.w = 0.3 + nCols*rtColDepth + nCols*rtPa + rtPa + 0.3
    //                              = nCols*rtSlot + 2*rtPa + 0.6
    var rtColDepth = s.rtFaceD*2 + s.rtGap;  // pair depth (no aisle)
    sectionLayouts[rt]={...s.rtLayout, totalBays:s.rtBays, actualHeight:s.rtLayout.height,
      requiredW: s.rtLayout.nCols * s.rtSlot + 2*s.rtPa + 0.6};  // full pa BOTH sides + margins
  });
  var preLayoutWW = Object.entries(sectionLayouts).reduce(function(mx,kv){
    if(baseRackOf(kv[0])==='ground') return mx;
    return Math.max(mx, kv[1].requiredW||0);
  }, 10);  // requiredW already includes 0.6 margins; no extra +0.6 needed
  actualWW = Math.max(10, preLayoutWW);

  // ── PASS 2: ground storage uses actualWW for column count ─────────────────
  RACK_ORDER.filter(rt=>baseRackOf(rt)==='ground').forEach(function(rt){
    if(!((rackConfig||[]).some(c=>LK(c)===rt) || rackTypeAreas?.[rt])) return;
    var sg = computeSection(rt, actualWW);  // ← uses actualWW, not wW
    if(sg) sectionLayouts[rt]={...sg.rtLayout, totalBays:sg.rtBays,
      actualHeight:sg.rtLayout.height,
      requiredW: sg.rtLayout.nCols * sg.rtSlot + 2*sg.rtPa + 0.6};  // full pa BOTH sides + margins
  });

  // ── RE-DERIVE BAND HEIGHTS FROM THE FINAL WIDTH ───────────────────────────
  // Every staging/support band is (area / warehouse width). The width is driven
  // by the racks, so these must be recomputed once actualWW is final, otherwise
  // the bands are too short and the rack sections overrun them.
  recH     = Math.max(4,(receivingArea||0)/actualWW);
  disH     = Math.max(4,(dispatchArea||0)/actualWW);
  stagingH = Math.max(recH,disH);
  offH     = Math.max(3,(officeArea||50)/actualWW);
  mheH     = mheArea>0 ? Math.max(2,mheArea/actualWW) : 0;
  supportH = offH + mheH;

  // ── NORTH BANDS: office / MHE, then outbound staging when docks are on both ends
  if (officeArea>0) {
    supportRects.push({ key:'office', x:0, y:cur, w:actualWW/2, h:offH,
      label:'OFFICE / WELFARE', color:'#dbeafe', border:'#3b82f6', text:'#1d4ed8' });
  }
  if (mheH>0) {
    supportRects.push({ key:'mhe', x:actualWW/2, y:cur, w:actualWW/2, h:offH+mheH,
      label:'MHE CHARGING', color:'#fdf4ff', border:'#9333ea', text:'#6b21a8' });
  }
  cur+=supportH;

  if (isBoth) {
    stagingRects.push({ key:'dispatch', x:0, y:cur, w:actualWW, h:disH,
      label:'DISPATCH / PACKING', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
    cur+=disH;
  }

  // ── BUILD zoneRects in RACK_ORDER (now all sectionLayouts are final) ─────
  RACK_ORDER.forEach(rt => {
    var sl=sectionLayouts[rt];
    if(!sl) return;
    var rtBase=baseRackOf(rt);
    var rtStyle=(RACK_TYPE_STYLE?.[rtBase])||{label:rtBase,color:'#f8fafc',border:'#e2e8f0',text:'#374151'};
    if(rackVariantLabel[rt]) rtStyle={...rtStyle, label:rtStyle.label+' - '+rackVariantLabel[rt]};
    sectionCrossAisles.push({x:0, y:cur, w:actualWW, h:SECTION_CA_W,
      label: firstRackSection ? 'CROSS AISLE (Staging ↔ Storage)' : `CROSS AISLE`});
    cur += SECTION_CA_W;
    firstRackSection = false;

    zoneRects.push({key:rt, x:0, y:cur, w:actualWW, h:sl.height,
      label:rtStyle.label, color:rtStyle.color, border:rtStyle.border, text:rtStyle.text,
      area:sl.area, rackType:rt, sectionLayout:{...sl, totalBays:sl.totalBays},
      actualHeight:sl.height, isTruncated:false});
    cur += sl.height;
  });
  // Build rack layout summary (for summary table below plan)
  var layoutSummary=[];
  var RACK_LABELS={shelving:'Shelving',liveStorage:'Flow/Live Storage',
    selective:'Selective Pallet Rack',doubleDeep:'Double-Deep Rack',
    driveIn:'Drive-In Rack',cantilever:'Cantilever Rack',ground:'Ground Storage'};
  var summaryMap={};
  (rackConfig||[]).forEach(function(cfg){
    var rt=LK(cfg); if(!rt) return;
    if(!summaryMap[rt]) summaryMap[rt]={
      name:(RACK_LABELS[baseRackOf(rt)]||baseRackOf(rt))
        +(cfg.variantCount>1?(' - '+cfg.variantLabel):''),
      bays:0,locs:0};
    summaryMap[rt].bays+=(cfg.baysNeeded||0);
    summaryMap[rt].locs+=(cfg.locs||0);
  });
  Object.entries(summaryMap).forEach(function(e){ layoutSummary.push(e[1]); });
  // Add final cross aisle between last rack section and dispatch staging
  if(sectionCrossAisles.length>0){
    sectionCrossAisles.push({x:0, y:cur, w:actualWW, h:SECTION_CA_W,
      label:'CROSS AISLE (Storage ↔ Staging)'});
    cur += SECTION_CA_W;
  }
  // 1:1 scale: cur already tracks actual heights since display = actual.
  // Reserve exactly what the south staging branch below will push.
  var southH  = isOne ? stagingH : isBoth ? recH : stagingH;
  var layoutWL = cur + southH;
  actualWL = Math.max(wL, layoutWL);

  // WIDTH = strictly from non-ground rack requiredW.
  // We do NOT use max(wW, ...) here — the building may be wider than the racks need,
  // but ground storage should only get columns based on what non-ground racks require.
  // This keeps ground storage section height predictable and avoids 22-column sprawl.
  var layoutWW = Object.entries(sectionLayouts).reduce(function(mx, kv){
    if(baseRackOf(kv[0])==='ground') return mx;  // ground does NOT drive warehouse width
    return Math.max(mx, kv[1].requiredW||0);
  }, 10);  // requiredW already includes margins
  // actualWW = rack-driven width
  actualWW = Math.max(10, layoutWW);

  // Rebuild ALL scale factors with final actual dimensions
  SVG_H = fullscreen ? Math.round(1800 * (actualWL/actualWW) * 0.8 + 220) : 820;
  DH = SVG_H - MT - MB;
  sX = DW/actualWW; sY = DH/actualWL;
  X = m=>ML+m*sX; Y = m=>MT+m*sY; W = m=>m*sX; H = m=>m*sY;

  // Update all width-dependent rects to use finalised actualWW
  zoneRects.forEach(function(z){ z.w=actualWW; });
  sectionCrossAisles.forEach(function(ca){ ca.w=actualWW; });
  stagingRects.forEach(function(s){
    // Recalculate positions based on actualWW
    if(s.key==='receiving'&&isOne){ s.w=actualWW/2; }
    if(s.key==='dispatch'&&isOne){ s.x=actualWW/2; s.w=actualWW/2; }
    if(s.key==='receiving'&&isBoth&&s.x===0){ s.w=actualWW-eastW2; }
    if(s.key==='dispatch'&&isBoth){ s.x=actualWW-eastW; s.w=eastW2; }
    if(s.key==='receiving'&&!isOne&&!isBoth){ s.w=actualWW; }
    if(s.key==='dispatch'&&!isOne&&!isBoth){ s.w=actualWW; }
  });
  supportRects.forEach(function(s){
    if(s.key==='office'){ s.w=actualWW/2; }
    if(s.key==='mhe'){ s.x=actualWW/2; s.w=actualWW/2; }
  });

  // Staging at south
  if (isOne) {
    stagingRects.push({ key:'receiving', x:0, y:cur, w:actualWW/2, h:stagingH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
    stagingRects.push({ key:'dispatch', x:actualWW/2, y:cur, w:actualWW/2, h:stagingH,
      label:'DISPATCH / PACKING', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
  } else if (isBoth) {
    stagingRects.push({ key:'receiving', x:0, y:cur, w:actualWW, h:recH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
  } else {
    var eastW2=Math.min(wW*0.3,14);
    stagingRects.push({ key:'receiving', x:0, y:cur, w:actualWW-eastW2, h:stagingH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
    stagingRects.push({ key:'dispatch', x:actualWW-eastW, y:cur, w:eastW2, h:stagingH,
      label:'DISPATCH', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
  }

  // Dock doors
  var dockDoors=[];
  var doorW=3.5;
  if (isOne) {
    var sp=actualWW/(totalDocks+1);
    for(let i=1;i<=totalDocks;i++) dockDoors.push({x:sp*i-doorW/2,y:actualWL,side:'south',label:`D${i}`});
  } else if (isBoth) {
    var ssp=actualWW/(inboundDocks+1);
    for(let i=1;i<=inboundDocks;i++) dockDoors.push({x:ssp*i-doorW/2,y:actualWL,side:'south',label:`D${i}`});
    var nsp=actualWW/(outboundDocks+1);
    for(let i=1;i<=outboundDocks;i++) dockDoors.push({x:nsp*i-doorW/2,y:0,side:'north',label:`D${inboundDocks+i}`});
  } else {
    var eastW=Math.min(actualWW*0.3,14);
    var southN=inboundDocks, eastN=outboundDocks;
    var ssp2=(actualWW-eastW)/(southN+1);
    for(let i=1;i<=southN;i++) dockDoors.push({x:ssp2*i-doorW/2,y:actualWL,side:'south',label:`D${i}`});
    var esp=actualWL/(eastN+1);
    for(let i=1;i<=eastN;i++) dockDoors.push({x:actualWW,y:esp*i,side:'east',label:`D${southN+i}`});
  }  // end dock config

  // ── RACK ROW HELPER ─────────────────────────────────────────────────────────
  // Each zone section = one rack type only (1:1 mapping)
  var zone2RackTypes={};
  zoneRects.forEach(z=>{
    if(z.rackType) zone2RackTypes[z.key]=[{rack:z.rackType,cfg:null}];
  });

  // Bay heights in the N-S direction (vertical in plan) per rack type
  // BAY_HEIGHT_M now sourced from rackBayWidthM (actual from rackConfig)

  var RACK_INFO_2D={
    shelving:   {depth:1.0, color:'#dbeafe', stroke:'#3b82f6'},
    liveStorage:{depth:1.2, color:'#bfdbfe', stroke:'#60a5fa'},
    selective:  {depth:2.2, color:'#fde68a', stroke:'#d97706'},
    doubleDeep: {depth:4.4, color:'#c7d2fe', stroke:'#818cf8'},
    driveIn:    {depth:5.5, color:'#334155', stroke:'#1e293b'},
    cantilever: {depth:2.0, color:'#fde8d8', stroke:'#f97316'},
    ground:     {depth:2.4, color:'#d97706', stroke:'#92400e'},
  };

  // Cross aisle spacing (along zone height = N-S)
  var CROSS_AISLE_INTERVAL={
    shelving:13, liveStorage:13,
    selective:27, doubleDeep:27, driveIn:27, cantilever:27, ground:27,
  };

  // Rack aisle width per type: from rackConfig.aisle (user entered) or defaults
  var rackAisleM={};
  (rackConfig||[]).forEach(function(cfg){
    var kA=LK(cfg);
    if(kA && !rackAisleM[kA]){
      var a=parseFloat(cfg.aisle)||0;
      if(a>0) rackAisleM[kA]=a/1000; // mm → m
    }
  });
  // Fallback defaults (mm → m)
  var DEFAULT_AISLE_M={shelving:1.2,liveStorage:1.2,selective:3.0,
    doubleDeep:3.5,driveIn:3.5,cantilever:3.0,ground:4.0};
  var rackBayDepthM={};
  (rackConfig||[]).forEach(function(cfg){
    var kD=LK(cfg);
    if(kD && cfg.bayD && !rackBayDepthM[kD])
      rackBayDepthM[kD]=parseFloat(cfg.bayD)/1000;
  });
  // Default single-face depths (if rackConfig doesn't provide bayD)
  var DEFAULT_FACE_DEPTH={shelving:0.6,liveStorage:0.6,selective:1.1,
    doubleDeep:1.1,driveIn:1.1,cantilever:1.0,ground:1.2};
  // Build rack-type → total bays needed from rackConfig
  var rackTypeBays={};
  (rackConfig||[]).forEach(cfg=>{
    var kB=LK(cfg);
    if(kB) rackTypeBays[kB]=(rackTypeBays[kB]||0)+(cfg.baysNeeded||0);
  });

  var rackRowsForZone=(zone)=>{
    var rows=[], crossAisles=[], dimAnnotations=[];
    var dom=zone.rackType||(zone2RackTypes[zone.key]?.[0]?.rack)||'shelving';
    var domB=baseRackOf(dom); // base rack type for static lookup tables
    var ri=(RACK_INFO_2D?.[domB])||(RACK_INFO_2D?.shelving)||{depth:1.0,color:'#dbeafe',stroke:'#3b82f6'};
    // Picking aisle: user-entered per rack type, else global aisleM, else default
    var pa=(rackAisleM?.[dom])||(DEFAULT_AISLE_M?.[domB])||aisleM||1.2;
    var colSlot=ri.depth+pa;
    var bayHm=(rackBayWidthM?.[dom])||(BAY_HEIGHT_M_LOOKUP?.[domB])||0.9;

    // Column label: A,B,...,Z,AA,AB,...
    var colLabel=(i)=>i<26?String.fromCharCode(65+i)
      :String.fromCharCode(64+Math.floor(i/26))+String.fromCharCode(65+(i%26));

    var sl=zone.sectionLayout||sectionLayouts[dom]||{nCols:3,baysPerCol:5,crossYPositions:[]};
    // Always get totalBays from sectionLayouts[dom] first (most reliable source)
    var totalBays=(sectionLayouts?.[dom]?.totalBays)||(sl?.totalBays)||0;

    var faceDepth=(rackBayDepthM?.[dom])||(DEFAULT_FACE_DEPTH?.[domB])||0.6;
    var backGap=(domB==='shelving'||domB==='liveStorage')?0.05:0.10;
    var colDepth=faceDepth*2+backGap;
    var actualColSlot=colDepth+pa;
    var crossInterval=(CROSS_AISLE_INTERVAL?.[domB])||13;

    // Maximum columns that fit in warehouse width
    // For column computation: use original wW (not expanded actualWW) so shelving/SPR
    // don't get extra columns because ground storage expanded the warehouse width.
    // Ground storage uses zone.w (= actualWW) to fill the expanded space.
    var isGround=(domB==='ground');
    var effectiveW=isGround?zone.w:Math.min(zone.w,wW);
    var minBpfZone=(MIN_BPF_BY_TYPE?.[domB])||1;  // enforce minimum rows per column
    var maxNcols=Math.max(3, Math.floor(effectiveW/actualColSlot));
    // With minBpf: cap nCols so bpf >= minBpfZone
    var bpfAtMax=Math.max(minBpfZone, totalBays>0?Math.max(1,Math.ceil(totalBays/2/maxNcols)):minBpfZone);
    var nColsMin=totalBays>0?Math.max(3,Math.ceil(totalBays/2/bpfAtMax)):maxNcols;
    var nCols=Math.min(maxNcols, nColsMin);
    var bpf=totalBays>0?Math.max(minBpfZone,Math.ceil(totalBays/2/nCols)):bpfAtMax;

    // ── HONOUR THE RESERVED SECTION GEOMETRY ────────────────────────────────
    // computeSectionLayout already decided nCols/baysPerCol and the band height
    // that `cur` advanced by. Drawing a different bpf here would push rows past
    // the section into the next band (or into the staging area), so prefer the
    // reserved values whenever they exist.
    if(sl && sl.nCols>0)      nCols = sl.nCols;
    if(sl && sl.baysPerCol>0) bpf   = sl.baysPerCol;

    // Hard clamp: never draw taller than the band actually reserved.
    if(zone.h>0 && bayHm>0){
      var _avail = zone.h - 0.6;                       // top + bottom margin
      var _nCA   = Math.floor(_avail/Math.max(1e-6,crossInterval)); // cross aisles inside
      var _rowsH = Math.max(0, _avail - _nCA*CROSS_AISLE_W_M);
      var _maxBpf= Math.max(1, Math.floor(_rowsH/bayHm + 1e-9));
      if(bpf > _maxBpf) bpf = _maxBpf;
    }


    // Rebuild cross aisles from actual bpf (ignore pre-computed crossYPositions)
    var crossYs=[];
    {
      var _y=0.3,_s=0,_bc=0;
      for(var _b=0;_b<bpf;_b++){
        _y+=bayHm; _s+=bayHm; _bc++;
        if(_s>=crossInterval&&_bc>=2&&(bpf-_b-1)>=2){
          crossYs.push(_y);
          if(!(zone.h>0) || (_y+CROSS_AISLE_W_M/2)<=zone.h)
            crossAisles.push({x:zone.x,y:zone.y+_y-CROSS_AISLE_W_M/2,
              w:zone.w,h:CROSS_AISLE_W_M,isCrossAisle:true});
          _y+=CROSS_AISLE_W_M; _s=0; _bc=0;
        }
      }
    }

    // Fit bpf bays + cross aisles, but never exceed the reserved band height
    var zoneH=_y+0.3;
    if(zone.h>0 && zoneH>zone.h) zoneH=zone.h;
    var breakYs=[0,...crossYs,zoneH-0.3];
    let curX=zone.x+0.3+pa/2;   // +pa/2 shift → rx[0]=zone.x+0.3+pa (full picking aisle before col A)
    var globalBayNum=1;

    for(let i=0;i<nCols;i++){
      var rx=curX+i*actualColSlot+pa/2; // half-aisle on each side of pair
      if(rx+colDepth>zone.x+zone.w-0.3) break;
      var label=colLabel(i);

      // ── Pass 1: count actual drawn bays across all segments ──────────
      var actualBays=0;
      var zoneBottom=zone.y+(zone.h>0?zone.h:zoneH)-0.3;  // never draw past the band
      for(let j2=0;j2<breakYs.length-1;j2++){
        var sy0=zone.y+breakYs[j2]+(j2>0?CROSS_AISLE_W_M/2:0.3);
        var ey0=zone.y+breakYs[j2+1]-(j2<breakYs.length-2?CROSS_AISLE_W_M/2:0);
        if(ey0>zoneBottom) ey0=zoneBottom;
        if(ey0-sy0>=0.5) actualBays+=Math.max(1,Math.floor((ey0-sy0)/bayHm+1e-9));
      }

      // Each B2B module = ONE rect spanning full colDepth (faceDepth+gap+faceDepth)
      // Center partition line drawn in SVG. Bay number shown once per module.
      var colFrontStart=globalBayNum;
      var colBackStart=globalBayNum+actualBays;
      var frontOffset=0;

      // Draw TWO separate rects (front face + back face), NO partition line
      for(let j=0;j<breakYs.length-1;j++){
        var sy=zone.y+breakYs[j]+(j>0?CROSS_AISLE_W_M/2:0.3);
        var ey=zone.y+breakYs[j+1]-(j<breakYs.length-2?CROSS_AISLE_W_M/2:0);
        if(ey>zoneBottom) ey=zoneBottom;   // clamp to the reserved band
        if(ey-sy<0.5) continue;
        var segBays=Math.max(1,Math.floor((ey-sy)/bayHm+1e-9));
        rows.push({x:rx, y:sy, w:faceDepth, h:ey-sy, ...ri, dom, bayHm,
          colIdx:i, colLabel:label, segIdx:j,
          bayStart:colFrontStart+frontOffset,
          isHalfRack:'front', bayCount:segBays, faceDepth, pa, backGap});
        rows.push({x:rx+faceDepth+backGap, y:sy, w:faceDepth, h:ey-sy, ...ri, dom, bayHm,
          colIdx:i, colLabel:label, segIdx:j,
          bayStart:colBackStart+frontOffset,
          isHalfRack:'back', bayCount:segBays, faceDepth, pa, backGap});
        frontOffset+=segBays;
      }
      if(i===0) dimAnnotations.push({
        x:rx, y:zone.y+0.3, zoneX:zone.x,
        faceDepth, backGap, colDepth, aisle:pa, bayWidthM:bayHm, dom,
        lastColRight: rx + (nCols-1)*actualColSlot + colDepth });  // right edge of LAST column
      globalBayNum+=actualBays*2; // front+back
    }  // end for(i) columns loop
    return {rows, crossAisles, nCols, baysPerCol:sl.baysPerCol, totalBays, dimAnnotations};
  };

  // Pallet symbols in staging area
  var palletSymbols=(rect,mode)=>{
    var palW=1.2, palH=1.2;
    var cols=Math.floor((rect.w-0.8)/palW);
    var rowsN=Math.floor((rect.h-0.8)/palH);
    var syms=[];
    for(let r=0;r<Math.min(rowsN,4);r++)
      for(let cl=0;cl<Math.min(cols,12);cl++)
        syms.push({x:rect.x+0.4+cl*palW,y:rect.y+0.4+r*palH,w:palW-0.1,h:palH-0.1,mode});
    return syms;
  };

  // Packing table symbols (2.4×0.8m each)
  var nBenches=Math.min(parseInt(packingBenches)||0,8);
  var packTables=[];
  var dispRect=stagingRects.find(s=>s.key==='dispatch');
  if(dispRect && nBenches>0){
    for(let i=0;i<nBenches;i++){
      packTables.push({x:dispRect.x+0.4+i*2.8,y:dispRect.y+dispRect.h-1.4,w:2.4,h:0.8});
    }
  }

  // Collect all rack rows and cross aisles
  var allRackRows=[], allCrossAisles=[], allDimAnnotations=[];
  zoneRects.forEach(zone=>{
    const{rows,crossAisles,dimAnnotations}=rackRowsForZone(zone);
    allRackRows.push(...rows);
    allCrossAisles.push(...crossAisles);
    if(dimAnnotations) allDimAnnotations.push(...dimAnnotations);
  });

  // ── POST-DRAW TRAILING AISLE FIX ─────────────────────────────────────────────
  // Measure the actual rightmost column edge from all drawn racks (non-ground).
  // If actualWW < rightEdge + maxPa + 0.3, expand it to guarantee full trailing aisle.
  {
    var maxNonGroundPa = 0;
    Object.entries(sectionLayouts).forEach(function(kv){
      if(baseRackOf(kv[0])==='ground') return;
      var sl=kv[1];
      // Approximate pa from rtPa stored in sectionLayouts (set via computeSection)
      var approxPa=(rackAisleM?.[kv[0]])||(DEFAULT_AISLE_M?.[baseRackOf(kv[0])])||aisleM||1.2;
      if(approxPa>maxNonGroundPa) maxNonGroundPa=approxPa;
    });
    var maxRightEdge=0;
    allRackRows.forEach(function(r){
      if(baseRackOf(r.dom)==='ground') return;
      var re=r.x+r.w;
      if(re>maxRightEdge) maxRightEdge=re;
    });
    var neededWW=maxRightEdge+maxNonGroundPa+0.3;
    if(neededWW>actualWW+0.01){
      actualWW=neededWW;
      // Rebuild scale + update all widths
      SVG_H = fullscreen ? Math.max(1800, Math.round(1800*(actualWL/actualWW))) : 820;
      DH=SVG_H-MT-MB; sX=DW/actualWW; sY=DH/actualWL;
      X=m=>ML+m*sX; Y=m=>MT+m*sY; W=m=>m*sX; H=m=>m*sY;
      zoneRects.forEach(function(z){ z.w=actualWW; });
      sectionCrossAisles.forEach(function(ca){ ca.w=actualWW; });
      stagingRects.forEach(function(s){
        if(s.key==='receiving'&&isOne){ s.w=actualWW/2; }
        if(s.key==='dispatch'&&isOne){ s.x=actualWW/2; s.w=actualWW/2; }
        if(s.key==='receiving'&&!isOne){ s.w=actualWW; }
        if(s.key==='dispatch'&&!isOne){ s.w=actualWW; }
      });
      supportRects.forEach(function(s){
        if(s.key==='office'){ s.w=actualWW/2; }
        if(s.key==='mhe'){ s.x=actualWW/2; s.w=actualWW/2; }
      });
    }
  }

  // Staging pallets
  var recRect=stagingRects.find(s=>s.key==='receiving');
  var recPallets=recRect ? palletSymbols(recRect,'inbound') : [];
  var disPallets=dispRect ? palletSymbols(dispRect,'outbound') : [];

  // MHE charging bays
  var MHE_BAY={counterbalance:{w:4,h:3.5},reach:{w:3.5,h:2.5},vna:{w:3,h:2}};
  var mheBay=MHE_BAY[forkType]||{w:3.5,h:2.5};
  var mheRect=supportRects.find(s=>s.key==='mhe');
  var mheBays=[];
  if(mheRect && design.nMHE>0){
    for(let i=0;i<Math.min(design.nMHE,6);i++){
      mheBays.push({x:mheRect.x+0.4+i*(mheBay.w+0.4),y:mheRect.y+0.4,w:mheBay.w,h:mheBay.h});
    }
  }

  // Right-side zone dimension arrows (y positions)
  var dimRight=[];
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

  // Return everything needed for rendering
  return {
    SVG_W, SVG_H, ML, MR, MT, MB, DW, DH, sX, sY, actualWL, actualWW, wW: actualWW, wL: actualWL,
    X: m=>ML+m*sX, Y: m=>MT+m*sY, W: m=>m*sX, H: m=>m*sY,
    dockSide, forkType, packingBenches: params.packingBenches,
    nMHE: design.nMHE||0, inboundMode: params.inboundMode, outboundMode: params.outboundMode,
    stagingH, isBoth, isOne, recH, disH, offH, mheH, supportH,
    zoneRects, stagingRects, supportRects, dockDoors,
    allRackRows, allCrossAisles, allDimAnnotations, sectionCrossAisles, layoutSummary, recPallets, disPallets,
    packTables, mheBays, dimRight, sectionLayouts,
    doorW: 3.5,
    // Design values needed in render
    zoneAreas: design.zoneAreas||{}, receivingArea: design.receivingArea||80,
    dispatchArea: design.dispatchArea||80, mheArea: design.mheArea||0,
    officeArea: design.officeArea||50, netRackArea: design.netRackArea||0,
    totalDocks: design.totalDocks||2, inboundDocks: design.inboundDocks||1,
    outboundDocks: design.outboundDocks||1, staging: design.staging||{},
  };
}

function FloorPlanSVG({ analysis, design, params, rackConfig, fullscreen=false,
  measureOn=false, measurePts=[], measurements=[], onMeasurePoint, snapOn=true }) {
  if (!design?.wW || !design?.wL) return null;

  let fp;
  try {
    fp = buildFloorPlanLayout(design, params, rackConfig, analysis, fullscreen);
  } catch(e) {
    console.error('[FloorPlanSVG] buildFloorPlanLayout error:', e);
    return <div style={{color:'red',padding:'8px',fontSize:'11px'}}>Plan error: {e.message}</div>;
  }
  if (!fp?.SVG_W) return <div style={{padding:'8px',color:'#9ca3af',fontSize:'11px'}}>Generating plan…</div>;
  const {
    SVG_W, SVG_H, ML, MR, MT, MB, DW, DH, sX, sY, actualWL, actualWW, wW, wL,
    X, Y, W, H,
    dockSide, forkType, packingBenches, nMHE, inboundMode, outboundMode,
    stagingH, isBoth, isOne, recH, disH, offH, mheH, supportH,
    zoneRects, stagingRects, supportRects, dockDoors,
    allRackRows, allCrossAisles, allDimAnnotations, sectionCrossAisles, layoutSummary, recPallets, disPallets,
    packTables, mheBays, dimRight, sectionLayouts, doorW,
    zoneAreas, receivingArea, dispatchArea, mheArea, officeArea, netRackArea,
    totalDocks, inboundDocks, outboundDocks, staging,
  } = fp;
  const MFT=3.2808, M2FT=10.7639;
  const ft = m => `${(m*MFT).toFixed(0)}'`;
  const sqft = m2 => `${Math.round(m2*M2FT).toLocaleString()} sq ft`;
  const aisleM = parseFloat(params.aisleW)||3.0;

  // ── MEASURE TOOL ──────────────────────────────────────────────────────────
  // Convert a browser click to warehouse metres. getScreenCTM().inverse() handles
  // zoom, viewBox scaling, CSS sizing and scroll offset in one step.
  const M_INV_X = px => (px - ML) / sX;
  const M_INV_Y = py => (py - MT) / sY;
  // Snap candidates: every real edge in the plan, in metres.
  // Built on click (not per render) so a 7,000-row plan costs nothing to draw.
  const snapTargets = () => {
    const xs = new Set([0, actualWW]);
    const ys = new Set([0, actualWL]);
    const addRect = r => {
      if (!r) return;
      if (r.x != null) { xs.add(+r.x.toFixed(3)); xs.add(+(r.x + (r.w||0)).toFixed(3)); }
      if (r.y != null) { ys.add(+r.y.toFixed(3)); ys.add(+(r.y + (r.h||0)).toFixed(3)); }
    };
    (allRackRows||[]).forEach(addRect);
    (allCrossAisles||[]).forEach(addRect);
    (sectionCrossAisles||[]).forEach(addRect);
    (zoneRects||[]).forEach(addRect);
    (stagingRects||[]).forEach(addRect);
    (supportRects||[]).forEach(addRect);
    return { xs:[...xs], ys:[...ys] };
  };

  // Snap a raw point onto the nearest edge within ~6 SVG px
  const snapPoint = (mx, my) => {
    // ~6 screen px, but never more than 0.5m — on a long plan sY is small, so an
    // uncapped tolerance would grab edges metres away and read the wrong feature.
    const SNAP_MAX_M = 0.5;
    const tolX = Math.min(6 / Math.max(1e-6, sX), SNAP_MAX_M);
    const tolY = Math.min(6 / Math.max(1e-6, sY), SNAP_MAX_M);
    const t = snapTargets();
    let bx = mx, by = my, hitX = false, hitY = false, dbx = tolX, dby = tolY;
    t.xs.forEach(v => { const d = Math.abs(v - mx); if (d < dbx) { dbx = d; bx = v; hitX = true; } });
    t.ys.forEach(v => { const d = Math.abs(v - my); if (d < dby) { dby = d; by = v; hitY = true; } });
    return { x:bx, y:by, snapped: hitX || hitY };
  };

  const handleMeasureClick = (evt) => {
    if (!measureOn || !onMeasurePoint) return;
    const svgEl = evt.currentTarget;
    let ux, uy;
    try {
      const ctm = svgEl.getScreenCTM();
      if (!ctm) return;
      const pt = svgEl.createSVGPoint
        ? svgEl.createSVGPoint()
        : new DOMPoint(0,0);
      pt.x = evt.clientX; pt.y = evt.clientY;
      const loc = pt.matrixTransform(ctm.inverse());
      ux = loc.x; uy = loc.y;
    } catch (err) {
      // Fallback: bounding-box ratio against the viewBox
      const r = svgEl.getBoundingClientRect();
      ux = ((evt.clientX - r.left) / r.width)  * SVG_W;
      uy = ((evt.clientY - r.top)  / r.height) * SVG_H;
    }
    let mx = M_INV_X(ux), my = M_INV_Y(uy), snapped = false;
    if (snapOn) { const s = snapPoint(mx, my); mx = s.x; my = s.y; snapped = s.snapped; }
    onMeasurePoint({ x:+mx.toFixed(3), y:+my.toFixed(3), snapped });
  };

  // Straight-line distance between two points, in metres
  const mDist = (a,b) => Math.sqrt(Math.pow(b.x-a.x,2) + Math.pow(b.y-a.y,2));

  return (
    <svg width={SVG_W} height={SVG_H}
      viewBox={`0 0 ${SVG_W} ${SVG_H}`}
      id={fullscreen?'fs-plan-svg':undefined}
      onClick={handleMeasureClick}
      style={{border:'1px solid #e2e8f0',borderRadius:'10px',background:'#ffffff',
               width:'100%',height:'auto',display:'block',
               cursor:measureOn?'crosshair':'default'}}>

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
      <rect x={X(0)} y={Y(0)} width={W(actualWW)} height={H(actualWL)} fill="#f8fafc" stroke="#1e293b" strokeWidth="2.5" rx="2"/>

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

      {/* ── SECTION CROSS AISLES — between rack types + staging boundaries ─ */}
      {(sectionCrossAisles||[]).map((ca,i)=>(
        <g key={`sca-${i}`}>
          <rect x={X(ca.x)} y={Y(ca.y)} width={W(ca.w)} height={Math.max(4,H(ca.h))}
            fill="#fef9c3" stroke="#fcd34d" strokeWidth="0.8" opacity="0.95"/>
          <line x1={X(ca.x+1)} y1={Y(ca.y+ca.h/2)} x2={X(ca.x+ca.w-1)} y2={Y(ca.y+ca.h/2)}
            stroke="#d97706" strokeWidth="0.8" strokeDasharray="4,3"/>
          <text x={X(ca.x+ca.w/2)} y={Y(ca.y+ca.h/2)}
            textAnchor="middle" dominantBaseline="middle"
            fontSize={Math.max(7,Math.min(11,H(ca.h)*0.35))}
            fontWeight="700" fill="#92400e">
            {ca.label||'CROSS AISLE'}
          </text>
        </g>
      ))}

      {/* ── BAY-LEVEL CROSS AISLES (within rack sections) ─── */}
      {allCrossAisles.map((a,i)=>{
        const ay=Y(a.y), ah=Math.max(3,H(a.h));
        const aw=W(a.w);
        const caLabel=`${(a.h*1000).toFixed(0)}mm`;
        return(
          <g key={`ca-${i}`}>
            <rect x={X(a.x)} y={ay} width={aw} height={ah}
              fill="#fef9c3" stroke="#ca8a04" strokeWidth="0.5" opacity="0.8"/>
            <text x={X(a.x+a.w/2)} y={ay+ah/2} textAnchor="middle"
              dominantBaseline="middle" fontSize="7" fill="#92400e" fontWeight="700">
              CROSS AISLE
            </text>
            {/* Dimension: cross aisle width on right margin */}
            {ah>5&&<>
              <line x1={X(a.x+a.w)+4} y1={ay} x2={X(a.x+a.w)+4} y2={ay+ah}
                stroke="#92400e" strokeWidth="0.8"/>
              <line x1={X(a.x+a.w)+1} y1={ay} x2={X(a.x+a.w)+7} y2={ay}
                stroke="#92400e" strokeWidth="0.8"/>
              <line x1={X(a.x+a.w)+1} y1={ay+ah} x2={X(a.x+a.w)+7} y2={ay+ah}
                stroke="#92400e" strokeWidth="0.8"/>
              <text x={X(a.x+a.w)+10} y={ay+ah/2} dominantBaseline="middle"
                fontSize="6.5" fill="#92400e" fontWeight="700">{caLabel}</text>
            </>}
          </g>
        );
      })}

      {/* ── RACK ROWS (top view — type-specific symbols) ─── */}
      {allRackRows.map((r,i)=>{
        const px=X(r.x), py=Y(r.y), pw=W(r.w), ph=Math.max(3,H(r.h));
        const dom=r.dom;

        // ── SELECTIVE PALLET RACK ─────────────────────────────────────────
        // Vertical column: bays stack along column HEIGHT, back-to-back partition at center WIDTH
        if(dom==='selective'||dom==='doubleDeep'||dom==='driveIn'){
          // Each rack face: colored rect + horizontal bay beam lines only (no pallet symbols)
          const bayHpx2=H(r.bayHm||2.7);
          const nBeams=Math.max(0,Math.floor(r.h/(r.bayHm||2.7))-1);
          return(
            <g key={`rr-${i}`}>
              <rect x={px} y={py} width={Math.max(3,pw)} height={ph}
                fill={r.color} stroke={r.stroke} strokeWidth="0.8" rx="0.5"/>
              {Array.from({length:nBeams},(_,b)=>(
                <line key={b} x1={px} y1={py+(b+1)*bayHpx2}
                  x2={px+pw} y2={py+(b+1)*bayHpx2}
                  stroke={r.stroke} strokeWidth="1"/>
              ))}
            </g>
          );
        }

        // ── DRIVE-IN RACK ─────────────────────────────────────────────────
        if(dom==='driveIn'){
          const laneW   = W(2.7);  // lane width in px
          const nLanes  = Math.max(1,Math.floor(r.w/2.7));
          const palDepth= Math.max(2, Math.floor(r.h/1.2)); // pallets deep per lane
          const palHpx  = ph/Math.max(palDepth,1)-0.5;
          return(
            <g key={`rr-${i}`}>
              {/* Dark background — solid rack mass */}
              <rect x={px} y={py} width={pw} height={ph} fill="#334155" rx="1"/>
              {/* Lane dividers + pallet stacks */}
              {Array.from({length:nLanes},(_,ln)=>(
                <g key={ln}>
                  {/* Lane opening (lighter strip) */}
                  <rect x={px+ln*laneW+2} y={py} width={laneW-4} height={ph} fill="#475569"/>
                  {/* Pallet rectangles going deep */}
                  {Array.from({length:Math.min(palDepth,6)},(_,pd)=>(
                    <rect key={pd}
                      x={px+ln*laneW+3} y={py+pd*(palHpx+0.5)}
                      width={laneW-6} height={palHpx}
                      fill={pd%2===0?'#f59e0b':'#d97706'} stroke="#92400e" strokeWidth="0.4" rx="0.5"/>
                  ))}
                  {/* Lane divider posts */}
                  <rect x={px+ln*laneW} y={py} width={2} height={ph} fill="#0f172a"/>
                </g>
              ))}
              <rect x={px+nLanes*laneW} y={py} width={2} height={ph} fill="#0f172a"/>
              {/* Entry arrow */}
              {pw>30&&<text x={px+pw/2} y={py-3} textAnchor="middle" fontSize="7" fontWeight="700" fill="#1e293b">
                ← DRIVE-IN ({nLanes} lanes)
              </text>}
            </g>
          );
        }

        // ── CANTILEVER RACK ───────────────────────────────────────────────
        if(dom==='cantilever'){
          const spineSpacing = W(1.5); // spine posts every 1.5m
          const nSpines = Math.max(1, Math.floor(r.w/1.5));
          const armLen  = ph*0.42; // arms reach ~42% of depth each side
          const spineY  = py+ph/2; // spine runs along centre of row
          return(
            <g key={`rr-${i}`}>
              {/* Background */}
              <rect x={px} y={py} width={pw} height={ph} fill="#ede9fe" stroke="#7c3aed" strokeWidth="0.8" rx="1"/>
              {/* Long items on arms (orange) */}
              <rect x={px+3} y={py+2}       width={pw-6} height={armLen-2} fill="#fed7aa" stroke="#f97316" strokeWidth="0.5" rx="1"/>
              <rect x={px+3} y={spineY+2}   width={pw-6} height={armLen-2} fill="#fed7aa" stroke="#f97316" strokeWidth="0.5" rx="1"/>
              {/* Spine line */}
              <line x1={px} y1={spineY} x2={px+pw} y2={spineY} stroke="#7c3aed" strokeWidth="2.5"/>
              {/* Spine posts + arms */}
              {Array.from({length:nSpines+1},(_,s)=>{
                const sx=px+s*spineSpacing;
                return(
                  <g key={s}>
                    {/* Spine column */}
                    <rect x={sx-2} y={py} width={4} height={ph} fill="#6d28d9"/>
                    {/* Top arm */}
                    <line x1={sx} y1={spineY} x2={sx} y2={py+2} stroke="#7c3aed" strokeWidth="1.5"/>
                    {/* Bottom arm */}
                    <line x1={sx} y1={spineY} x2={sx} y2={py+ph-2} stroke="#7c3aed" strokeWidth="1.5"/>
                  </g>
                );
              })}
              {/* Label */}
              {pw>50&&<text x={px+pw/2} y={py+ph/2+1} textAnchor="middle" fontSize="7" fontWeight="700" fill="#6d28d9" dominantBaseline="middle">CANTILEVER</text>}
            </g>
          );
        }

        // ── SHELVING / LIVE STORAGE — full B2B width, center partition ────
        {
          const bayHpx = H(r.bayHm||0.9);
          const nBayDividers = Math.max(0, Math.floor(r.h/(r.bayHm||0.9)+1e-9)-1);
          const halfPx = Math.max(1, (pw-W(r.backGap||0.05))/2);
          return(
            <g key={`rr-${i}`}>
              {/* Full B2B column */}
              <rect x={px} y={py} width={Math.max(3,pw)} height={Math.max(2,ph)}
                fill={r.color} stroke={r.stroke} strokeWidth="0.8" rx="0.5"/>
              {/* Horizontal bay dividers */}
              {Array.from({length:nBayDividers},(_,b)=>(
                <line key={b} x1={px} y1={py+(b+1)*bayHpx}
                  x2={px+pw} y2={py+(b+1)*bayHpx}
                  stroke={r.stroke} strokeWidth="0.5" strokeOpacity="0.5"/>
              ))}
              {/* Center partition — back-to-back join */}
              {r.showPartition&&<line
                x1={px+halfPx} y1={py} x2={px+halfPx} y2={py+ph}
                stroke={r.stroke} strokeWidth="1.4" opacity="0.85"/>}
            </g>
          );
        }
      })}

      {/* ── BAY NUMBERING — sequential, front + back separate ───────────── */}
      {allRackRows.map((r,i)=>{
        const px=X(r.x), py=Y(r.y), pw=W(r.w), ph=H(r.h);
        const bayHpx=H(r.bayHm||0.9);
        const showBayNum=pw>4&&bayHpx>4;
        const colFontSz=Math.max(9,Math.min(16,pw*0.9));
        const bayFontSz=Math.max(6,Math.min(13,Math.min(pw*0.75,bayHpx*0.75)));
        const nBays=r.bayCount||Math.max(1,Math.floor(r.h/(r.bayHm||0.9)+1e-9));
        const showColLetter = r.isHalfRack!=='back' && r.segIdx===0 && pw>3;
        return(
          <g key={`lbl-${i}`}>
            {showColLetter&&<text
              x={px+pw/2} y={py-1.5}
              textAnchor="middle" fontSize={colFontSz}
              fontWeight="900" fill={r.stroke}>
              {r.colLabel}
            </text>}
            {showBayNum&&Array.from({length:Math.min(nBays,300)},(_,b)=>{
              const by=py+b*bayHpx;
              if(by+bayHpx>py+ph+1) return null;
              return(
                <text key={b} x={px+pw/2} y={by+bayHpx/2}
                  textAnchor="middle" dominantBaseline="middle"
                  fontSize={bayFontSz} fontWeight="700" fill={r.stroke} opacity="0.9">
                  {(r.bayStart||1)+b}
                </text>
              );
            })}
          </g>
        );
      })}

      {/* ── DIMENSION ANNOTATIONS — face depth + gap + aisle ──────────────── */}
      {allDimAnnotations.map((d,i)=>{
        const px=X(d.x), py=Y(d.y);
        const pFace=W(d.faceDepth);   // one rack face width in px
        const pGap =W(d.backGap);     // back-panel gap in px
        const pAisle=W(d.aisle);      // picking aisle in px
        const bayHpx=H(d.bayWidthM);
        if(pFace<3) return null;
        const ty=py+Math.min(16,bayHpx*0.25); // inside first bay (avoids cross aisle overlap)
        const TICK=5;
        return(
          <g key={`dim-${i}`}>
            {/* ── face depth dim (front face) ── */}
            <line x1={px} y1={ty} x2={px+pFace} y2={ty} stroke="#1d4ed8" strokeWidth="1"/>
            <line x1={px} y1={ty-TICK/2} x2={px} y2={ty+TICK/2} stroke="#1d4ed8" strokeWidth="1"/>
            <line x1={px+pFace} y1={ty-TICK/2} x2={px+pFace} y2={ty+TICK/2} stroke="#1d4ed8" strokeWidth="1"/>
            <text x={px+pFace/2} y={ty-4} textAnchor="middle" fontSize="8" fill="#1d4ed8" fontWeight="700">
              {(d.faceDepth*1000).toFixed(0)}mm
            </text>
            {/* ── back gap dim ── */}
            {pGap>2&&<>
              <line x1={px+pFace} y1={ty} x2={px+pFace+pGap} y2={ty} stroke="#dc2626" strokeWidth="0.8" strokeDasharray="2,1"/>
              <text x={px+pFace+pGap/2} y={ty-4} textAnchor="middle" fontSize="6.5" fill="#dc2626" fontWeight="700">
                {(d.backGap*1000).toFixed(0)}mm
              </text>
            </>}
            {/* ── back face depth dim ── */}
            <line x1={px+pFace+pGap} y1={ty} x2={px+2*pFace+pGap} y2={ty} stroke="#1d4ed8" strokeWidth="1"/>
            <line x1={px+pFace+pGap} y1={ty-TICK/2} x2={px+pFace+pGap} y2={ty+TICK/2} stroke="#1d4ed8" strokeWidth="1"/>
            <line x1={px+2*pFace+pGap} y1={ty-TICK/2} x2={px+2*pFace+pGap} y2={ty+TICK/2} stroke="#1d4ed8" strokeWidth="1"/>
            <text x={px+pFace+pGap+pFace/2} y={ty-4} textAnchor="middle" fontSize="8" fill="#1d4ed8" fontWeight="700">
              {(d.faceDepth*1000).toFixed(0)}mm
            </text>
            {/* ── PICKING AISLE dim — between col A and col B ── */}
            {pAisle>4&&(()=>{
              const aisleX1=px+W(d.colDepth);
              const aisleX2=aisleX1+pAisle;
              return(<>
                <line x1={aisleX1} y1={ty} x2={aisleX2} y2={ty} stroke="#7c3aed" strokeWidth="1.2"/>
                <line x1={aisleX1} y1={ty-TICK/2} x2={aisleX1} y2={ty+TICK/2} stroke="#7c3aed" strokeWidth="1.2"/>
                <line x1={aisleX2} y1={ty-TICK/2} x2={aisleX2} y2={ty+TICK/2} stroke="#7c3aed" strokeWidth="1.2"/>
                <text x={(aisleX1+aisleX2)/2} y={ty-4} textAnchor="middle" fontSize="8.5" fill="#7c3aed" fontWeight="800">
                  {(d.aisle*1000).toFixed(0)}mm
                </text>
                <text x={(aisleX1+aisleX2)/2} y={ty+12} textAnchor="middle" fontSize="6.5" fill="#7c3aed" fontWeight="600">
                  Picking Aisle
                </text>
              </>);
            })()}
            {/* ── LEADING aisle (before first column) ── */}
            {pAisle>4&&(()=>{
              const leadX1=X(d.zoneX||0)+4;
              const leadX2=px;
              if(leadX2-leadX1<4) return null;
              return(<>
                <line x1={leadX1} y1={ty} x2={leadX2} y2={ty} stroke="#7c3aed" strokeWidth="0.8" strokeDasharray="3,2"/>
                <line x1={leadX1} y1={ty-TICK/2} x2={leadX1} y2={ty+TICK/2} stroke="#7c3aed" strokeWidth="0.8"/>
                <text x={(leadX1+leadX2)/2} y={ty+12} textAnchor="middle" fontSize="6" fill="#7c3aed">
                  {(d.aisle*1000).toFixed(0)}mm
                </text>
              </>);
            })()}
            {/* ── TRAILING aisle (after last column) ── */}
            {pAisle>4&&d.lastColRight&&(()=>{
              const trailX1=X(d.lastColRight);
              const trailX2=X(actualWW)-4;
              if(trailX2-trailX1<4) return null;
              return(<>
                <line x1={trailX1} y1={ty} x2={trailX2} y2={ty} stroke="#7c3aed" strokeWidth="0.8" strokeDasharray="3,2"/>
                <line x1={trailX2} y1={ty-TICK/2} x2={trailX2} y2={ty+TICK/2} stroke="#7c3aed" strokeWidth="0.8"/>
                <text x={(trailX1+trailX2)/2} y={ty+12} textAnchor="middle" fontSize="6" fill="#7c3aed">
                  {(d.aisle*1000).toFixed(0)}mm
                </text>
              </>);
            })()}
            {/* ── bay N-S height dim ── */}
            {bayHpx>5&&<>
              <line x1={px-3} y1={py} x2={px-3} y2={py+bayHpx} stroke="#166534" strokeWidth="1"/>
              <line x1={px-6} y1={py} x2={px-1} y2={py} stroke="#166534" strokeWidth="1"/>
              <line x1={px-6} y1={py+bayHpx} x2={px-1} y2={py+bayHpx} stroke="#166534" strokeWidth="1"/>
              <text x={px-9} y={py+bayHpx/2} textAnchor="middle" dominantBaseline="middle"
                fontSize="7.5" fill="#166534" fontWeight="700"
                transform={`rotate(-90,${px-9},${py+bayHpx/2})`}>
                {(d.bayWidthM*1000).toFixed(0)}mm
              </text>
            </>}
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

      {/* ── SECTION LABELS (rack type name at left margin of each section) ─── */}
      {zoneRects.map(z=>{
        const py=Y(z.y), ph=H(z.h);
        if(ph<6) return null;
        const sl=z.sectionLayout||{};
        const isTrunc=z.isTruncated&&z.actualHeight>z.h;
        return(
          <g key={`zl-${z.key}`}>
            <text x={X(z.x)+6} y={py+Math.min(ph/2,18)}
              dominantBaseline="middle"
              fontSize={Math.min(11, ph*sY*0.3)} fontWeight="700" fill={z.text}
              opacity="0.85">
              {z.label}
            </text>
            {ph>24&&<text x={X(z.x)+6} y={py+Math.min(ph/2,18)+13}
              dominantBaseline="middle"
              fontSize="8" fontWeight="400" fill={z.text} opacity="0.7">
              {(z.area||0).toFixed(0)}m²
            </text>}
            {/* Truncation indicator: zigzag line at bottom of capped section */}
            {isTrunc&&(()=>{
              const bY=py+ph-4;
              const pts=[0,4,8,4,0].map((dy,i)=>`${X(z.x)+i*(W(z.w)/4)},${bY+dy}`).join(' ');
              return(<>
                <polyline points={pts} fill="none" stroke={z.border||'#94a3b8'}
                  strokeWidth="1.5" strokeDasharray="4,2"/>
                <text x={X(z.x+z.w/2)} y={bY+10} textAnchor="middle"
                  fontSize="7.5" fill={z.text} fontWeight="700" opacity="0.8">
                  ▼ {sl.totalBays?.toLocaleString()} bays total (layout shows partial pattern)
                </text>
              </>);
            })()}
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

      {/* center divider line removed — unnecessary aisle */}

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


            {/* Column dividers */}

      {/* ── TOTAL AREA FOOTER ─── */}
      <text x={X(wW/2)} y={SVG_H-4} textAnchor="middle" fontSize="10" fontWeight="700" fill="#374151">
        {`Total gross area: ${(actualWW*actualWL).toLocaleString()}m²  (${Math.round(actualWW*actualWL*10.7639).toLocaleString()} sq ft)  ·  ${Math.round(actualWW*10)/10}×${Math.round(actualWL)}m  ·  ${dockSide==='one'?'One-side':'Opposite-side'} docks  ·  Derived from ${Object.keys(sectionLayouts).length} rack type sections`}
      </text>
      {/* ── MEASUREMENT OVERLAY ─────────────────────────────────────────── */}
      {(measurements.length>0 || measurePts.length>0) && (
        <g id="measure-layer">
          {measurements.map((mm,mi)=>{
            const a=mm[0], b=mm[1];
            const dd=mDist(a,b);
            const x1=X(a.x), y1=Y(a.y), x2=X(b.x), y2=Y(b.y);
            const mx=(x1+x2)/2, my=(y1+y2)/2;
            const dx=b.x-a.x, dy=b.y-a.y;
            let ang=Math.atan2(y2-y1,x2-x1)*180/Math.PI;
            if(ang>90) ang-=180; else if(ang<-90) ang+=180;
            const lbl=(dd*1000).toFixed(0)+'mm';
            const sub=dd.toFixed(2)+'m / '+(dd*3.2808).toFixed(1)+'ft';
            const boxW=Math.max(58, lbl.length*7+22);
            // Axis-locked measurements draw green so it is obvious they are exact
            const axis=b.ortho||null;
            const CLR=axis?'#059669':'#be185d';
            const CLR2=axis?'#047857':'#9f1239';
            return (
              <g key={'ms'+mi}>
                <line x1={x1} y1={y1} x2={x2} y2={y2}
                  stroke={CLR} strokeWidth="1.6" strokeLinecap="round"/>
                {!axis && (<>
                  <line x1={x1} y1={y1} x2={x2} y2={y1}
                    stroke={CLR} strokeWidth="0.6" strokeDasharray="3,3" opacity="0.5"/>
                  <line x1={x2} y1={y1} x2={x2} y2={y2}
                    stroke={CLR} strokeWidth="0.6" strokeDasharray="3,3" opacity="0.5"/>
                </>)}
                {[[x1,y1,a],[x2,y2,b]].map(([cx,cy,pp],ci)=>(
                  pp&&pp.snapped
                    ? <rect key={ci} x={cx-3.6} y={cy-3.6} width="7.2" height="7.2"
                        fill="#fff" stroke={CLR} strokeWidth="1.6"/>
                    : <circle key={ci} cx={cx} cy={cy} r="4"
                        fill="#fff" stroke={CLR} strokeWidth="1.6"/>
                ))}
                <circle cx={x1} cy={y1} r="1.4" fill={CLR}/>
                <circle cx={x2} cy={y2} r="1.4" fill={CLR}/>
                {!axis && Math.abs(dx)>0.05 && (
                  <text x={(x1+x2)/2} y={y1-3} textAnchor="middle" fontSize="6"
                    fill={CLR} opacity="0.85">
                    {'dX '+(Math.abs(dx)*1000).toFixed(0)+'mm'}
                  </text>
                )}
                {!axis && Math.abs(dy)>0.05 && (
                  <text x={x2+4} y={(y1+y2)/2} fontSize="6" fill={CLR} opacity="0.85">
                    {'dY '+(Math.abs(dy)*1000).toFixed(0)+'mm'}
                  </text>
                )}
                <g transform={'translate('+mx+','+my+') rotate('+ang+')'}>
                  <rect x={-boxW/2} y={-19} width={boxW} height="24" rx="4"
                    fill="#ffffff" stroke={CLR} strokeWidth="1" opacity="0.96"/>
                  <text x="0" y="-9" textAnchor="middle" fontSize="9.5"
                    fontWeight="800" fill={CLR}>{lbl}</text>
                  <text x="0" y="0" textAnchor="middle" fontSize="6.5" fill={CLR2}>{sub}</text>
                  {axis && (
                    <text x={boxW/2-7} y={-11} textAnchor="middle" fontSize="6"
                      fontWeight="800" fill="#059669">{axis==='H'?'0°':'90°'}</text>
                  )}
                </g>
                <text x={x1+6} y={y1-6} fontSize="7" fontWeight="800" fill={CLR}>{mi+1}</text>
              </g>
            );
          })}
          {measurePts.map((pt,pi)=>(
            <g key={'mp'+pi}>
              <circle cx={X(pt.x)} cy={Y(pt.y)} r="5" fill="#fff"
                stroke="#be185d" strokeWidth="1.6"/>
              <circle cx={X(pt.x)} cy={Y(pt.y)} r="1.6" fill="#be185d"/>
              <line x1={X(pt.x)-9} y1={Y(pt.y)} x2={X(pt.x)+9} y2={Y(pt.y)}
                stroke="#be185d" strokeWidth="0.7" opacity="0.6"/>
              <line x1={X(pt.x)} y1={Y(pt.y)-9} x2={X(pt.x)} y2={Y(pt.y)+9}
                stroke="#be185d" strokeWidth="0.7" opacity="0.6"/>
              <text x={X(pt.x)+8} y={Y(pt.y)-7} fontSize="7"
                fontWeight="700" fill="#be185d">click 2nd point</text>
            </g>
          ))}
        </g>
      )}

    </svg>
  );
}


// ─── RACK LOCATION DRILL-DOWN DOWNLOAD ───────────────────────────────────────
function downloadRackLocations(cfg, analysis) {
  if (!cfg || !analysis) return;
  const wb    = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString();
  const ws    = (data, cols) => {
    const s = XLSX.utils.aoa_to_sheet(data);
    if (cols) s['!cols'] = cols.map(w => ({ wch: w }));
    return s;
  };

  const isShelving = ['shelving','liveStorage'].includes(cfg.rack);
  const tierH      = parseFloat(cfg.tierHeight) || cfg.shelfH || 0;
  const tiers      = parseInt(cfg.tiers) || 1;
  const locsPerBay = cfg.locsPerBayTotal || cfg.locsPerBay || 0;
  const binD       = cfg.binDims || [];

  // Filter slotted SKUs for this exact rack + bin combination
  const skusInCfg = (analysis.slotted || [])
    .filter(s => s.rack === cfg.rack && s.bin === cfg.bin)
    .sort((a, b) => b.locsReq - a.locsReq);

  // ── Sheet 1: Calculation Summary ─────────────────────────────────────────
  const calcRows = [
    [`LOCATION CALCULATION — ${cfg.rackName.toUpperCase()}`],
    ['Bin / Container Type:', cfg.binName],
    ['Generated:', today],
    [],
    ['RACK BAY CONFIGURATION'],
    ['Bay Width',  `${cfg.bayW} mm`],
    ['Bay Depth',  `${cfg.bayD} mm`],
    ...(isShelving ? [
      ['Height per Tier',         `${tierH} mm`],
      ['Shelf clearance per level',`${cfg.clearance} mm`],
      ['Number of tiers',          tiers],
      ['Bin orientation',          cfg.orientation==='LW'?`L (${binD[0]}mm) along bay width`:`W (${binD[1]}mm) along bay width`],
    ] : [
      ['Rack levels',  cfg.levels],
    ]),
    [],
    ['STEP-BY-STEP CALCULATION'],
    ...(isShelving && binD.length ? [
      ['Bin dimensions (L × W × H)', `${binD[0]} × ${binD[1]} × ${binD[2]} mm`],
      ['Step 1 — Bins across bay width',
        `floor(${cfg.bayW} ÷ ${cfg.orientation==='LW'?binD[0]:binD[1]}) = ${cfg.acrossW} bins`],
      ['Step 2 — Bins along bay depth',
        `floor(${cfg.bayD} ÷ ${cfg.orientation==='LW'?binD[1]:binD[0]}) = ${cfg.acrossD} bins`],
      ['Step 3 — Shelf levels per tier',
        `floor(${tierH} ÷ (${binD[2]} bin height + ${cfg.clearance} clearance)) = ${cfg.levels} levels`],
      ['Step 4 — Locations per bay (1 tier)',
        `${cfg.acrossW} across × ${cfg.acrossD} deep × ${cfg.levels} levels = ${cfg.locsPerBay} locations`],
      ...(tiers > 1 ? [
        ['Step 5 — Locations per bay (all tiers)',
          `${cfg.locsPerBay} × ${tiers} tiers = ${locsPerBay} locations`],
      ] : []),
    ] : [
      ['Locations per bay', locsPerBay],
    ]),
    ['Bays required',
      `ceil(${cfg.locs} total locations ÷ ${locsPerBay} per bay) = ${cfg.baysNeeded} bays`],
    [],
    ['RESULT SUMMARY'],
    ['Rack type',               cfg.rackName],
    ['Total SKUs in config',    skusInCfg.length],
    ['Total locations needed',  cfg.locs],
    ['Number of bays',          cfg.baysNeeded],
    ['Floor area (incl. aisles)', `${cfg.area || 0} m²`],
    [],
    ['PER-SKU FORMULA'],
    ['Locations per SKU', '= ceil( ceil(Stock Qty ÷ Units per Location) ÷ Sharing Factor )'],
    ['Units per Location', 'Physical fit check: bins × depth × levels in bay; capped by volume estimate'],
    ['Sharing Factor', 'VF/F/M = 1 (dedicated) | S = 2 | VS = 4 | NM = 8 SKUs share one location'],
    ['Aisle model', 'Shared-aisle (back-to-back rows): each bay owns half an aisle (aisleW ÷ 2)'],
    ['Total locations', 'Sum of all SKU effective location requirements after sharing'],
  ];

  XLSX.utils.book_append_sheet(wb,
    ws(calcRows, [40, 42]), '1. Calculation Summary');

  // ── Sheet 2: Per-SKU Detail ───────────────────────────────────────────────
  const skuRows = [
    [`SKU LOCATION DETAIL — ${cfg.rackName} (${cfg.binName})`],
    [`Config: ${cfg.bayW}mm W × ${cfg.bayD}mm D${isShelving?` × ${tierH}mm H/tier × ${tiers} tier(s)`:''} · ${locsPerBay} locations/bay`],
    [],
    ['SKU Code','L (mm)','W (mm)','H (mm)','Vol (cm³)','Velocity','Size Band',
     'Units / Location','Stock Qty','Locations Needed','Calculation','Zone'],
    ...skusInCfg.map(s => [
      s.sku,
      s.L  > 0 ? s.L  : '—',
      s.W  > 0 ? s.W  : '—',
      s.H  > 0 ? s.H  : '—',
      s.volCm3 > 0 ? +s.volCm3.toFixed(0) : '—',
      s.vb, s.sb,
      s.upb,
      s.stock,
      s.locsReq,
      s.stock > 0 && s.upb > 0
        ? `ceil(${s.stock} ÷ ${s.upb}) = ${s.locsReq}`
        : s.stock === 0 ? 'No stock' : '—',
      s.zoneName || '—',
    ]),
    [],
    ['TOTALS','','','','','','','',
      skusInCfg.reduce((s,r)=>s+r.stock,   0),
      skusInCfg.reduce((s,r)=>s+r.locsReq, 0),
      '', ''],
  ];

  XLSX.utils.book_append_sheet(wb,
    ws(skuRows, [22,8,8,8,10,10,10,16,12,16,28,20]),
    '2. SKU Detail');

  const fname = `Locations_${cfg.rack}_${cfg.bin}_${today.replace(/\//g,'-')}.xlsx`;
  XLSX.writeFile(wb, fname);
}

// ─── 3D ROTATABLE WAREHOUSE MODEL (Three.js) ────────────────────────────────

// ─── DXF FLOOR PLAN EXPORT ────────────────────────────────────────────────────
// Generates AutoCAD DXF (Drawing Exchange Format).
// Open in AutoCAD → File → Open → .dxf → then Save As → .dwg
// Supported by all CAD tools: AutoCAD, FreeCAD, LibreCAD, Rhino, SolidWorks.
// ─── DXF EXPORT (AutoCAD R12 / AC1009 ASCII) ─────────────────────────────────
// R12 is deliberate: it is the most permissive, most universally readable DXF
// flavour. It needs no entity handles, no AcDb subclass markers and no OBJECTS
// section, all of which R2000+ requires and whose absence makes readers report
// "invalid drawing input". Geometry comes from buildFloorPlanLayout so the CAD
// file matches the drawn plan. 1 unit = 1 metre. Y is flipped so north is up.
function exportDXF(analysis, design, params, rackConfig) {
  let fp;
  try {
    fp = buildFloorPlanLayout(design, params, rackConfig || [], analysis, false);
  } catch (err) {
    alert('Could not build the layout for DXF export: ' + err.message);
    return;
  }
  if (!fp || !(fp.actualWW > 0)) { alert('No layout available to export.'); return; }

  const WW = fp.actualWW, WL = fp.actualWL;
  const L  = [];
  const d  = (...pairs) => pairs.forEach(v => L.push(String(v)));
  const n  = v => (isFinite(v) ? v : 0).toFixed(4);
  const fy = y => WL - y;                 // CAD Y is up, plan Y is down

  // DXF R12 is a plain ASCII format — strip anything outside ASCII or it is
  // rejected. Also map the few symbols the plan labels actually use.
  const asc = (str) => String(str == null ? '' : str)
    .replace(/\u00b2/g, '2').replace(/\u00b3/g, '3')
    .replace(/\u00d7/g, 'x').replace(/\u2194/g, '<->')
    .replace(/[\u2013\u2014]/g, '-').replace(/\u00b7/g, '.')
    .replace(/[\u2018\u2019]/g, "'").replace(/[\u201c\u201d]/g, '"')
    .replace(/\u00b0/g, 'deg')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[^\x20-\x7E]/g, '')
    .trim();

  const LAYERS = [
    ['0',                7],
    ['WALLS',            7],
    ['RACKS_SHELVING',   5],
    ['RACKS_PALLET',     1],
    ['RACKS_DRIVEIN',    8],
    ['RACKS_CANTILEVER', 6],
    ['RACKS_GROUND',    30],
    ['CROSS_AISLE',     52],
    ['SECTION',         42],
    ['STAGING',          3],
    ['SUPPORT',          4],
    ['DOCKS',            2],
    ['TEXT',             7],
    ['DIMS',             9],
  ];
  const RACK_LAYER = {
    shelving:'RACKS_SHELVING', liveStorage:'RACKS_SHELVING',
    selective:'RACKS_PALLET',  doubleDeep:'RACKS_PALLET',
    driveIn:'RACKS_DRIVEIN',   cantilever:'RACKS_CANTILEVER',
    ground:'RACKS_GROUND',
  };

  // ── HEADER (R12 variables only) ───────────────────────────────────────────
  d('0','SECTION','2','HEADER');
  d('9','$ACADVER','1','AC1009');
  d('9','$INSBASE','10','0.0','20','0.0','30','0.0');
  d('9','$EXTMIN','10',n(-4),'20',n(-4),'30','0.0');
  d('9','$EXTMAX','10',n(WW+8),'20',n(WL+14),'30','0.0');
  d('9','$LIMMIN','10',n(-4),'20',n(-4));
  d('9','$LIMMAX','10',n(WW+8),'20',n(WL+14));
  d('0','ENDSEC');

  // ── TABLES: LTYPE must exist because every LAYER references CONTINUOUS,
  //    and STYLE must exist because every TEXT references STANDARD ──────────
  d('0','SECTION','2','TABLES');

  d('0','TABLE','2','LTYPE','70','1');
  d('0','LTYPE','2','CONTINUOUS','70','0','3','Solid line','72','65','73','0','40','0.0');
  d('0','ENDTAB');

  d('0','TABLE','2','LAYER','70',String(LAYERS.length));
  LAYERS.forEach(([name,col]) => {
    d('0','LAYER','2',name,'70','0','62',String(col),'6','CONTINUOUS');
  });
  d('0','ENDTAB');

  d('0','TABLE','2','STYLE','70','1');
  d('0','STYLE','2','STANDARD','70','0','40','0.0','41','1.0','50','0.0',
    '71','0','42','0.2','3','txt','4','');
  d('0','ENDTAB');

  d('0','ENDSEC');

  // ── ENTITY HELPERS (R12: POLYLINE + VERTEX + SEQEND, no LWPOLYLINE) ───────
  const poly = (layer, pts, closed=true) => {
    if (!pts || pts.length < 2) return;
    d('0','POLYLINE','8',layer,'66','1','70',closed?'1':'0',
      '10','0.0','20','0.0','30','0.0');
    pts.forEach(([px,py]) => {
      d('0','VERTEX','8',layer,'10',n(px),'20',n(fy(py)),'30','0.0');
    });
    d('0','SEQEND','8',layer);
  };
  const rect = (layer,x1,y1,x2,y2) =>
    poly(layer, [[x1,y1],[x2,y1],[x2,y2],[x1,y2]], true);
  const line = (layer,x1,y1,x2,y2) => {
    d('0','LINE','8',layer,'10',n(x1),'20',n(fy(y1)),'30','0.0',
      '11',n(x2),'21',n(fy(y2)),'31','0.0');
  };
  // hj: 0 = left, 1 = centre
  const txt = (layer,x,y,h,str,hj=0) => {
    const s2 = asc(str);
    if (!s2) return;
    d('0','TEXT','8',layer,'10',n(x),'20',n(fy(y)),'30','0.0',
      '40',n(h),'1',s2,'7','STANDARD');
    if (hj) d('72','1','11',n(x),'21',n(fy(y)),'31','0.0');
  };

  d('0','SECTION','2','ENTITIES');

  // ── SHELL ─────────────────────────────────────────────────────────────────
  rect('WALLS',0,0,WW,WL);

  // ── SECTION BANDS ─────────────────────────────────────────────────────────
  (fp.zoneRects||[]).forEach(z => {
    rect('SECTION', z.x, z.y, z.x+z.w, z.y+z.h);
    if (z.label) txt('TEXT', z.x+0.4, z.y+0.9, 0.55, z.label);
  });

  // ── RACK ROWS (exactly as drawn) ──────────────────────────────────────────
  (fp.allRackRows||[]).forEach(r => {
    rect(RACK_LAYER[baseRackOf(r.dom)] || 'RACKS_SHELVING',
         r.x, r.y, r.x+r.w, r.y+r.h);
  });

  // ── AISLES ────────────────────────────────────────────────────────────────
  (fp.allCrossAisles||[]).forEach(a => rect('CROSS_AISLE', a.x, a.y, a.x+a.w, a.y+a.h));
  (fp.sectionCrossAisles||[]).forEach(a => {
    rect('CROSS_AISLE', a.x, a.y, a.x+a.w, a.y+a.h);
    if (a.label) txt('TEXT', a.x+a.w/2, a.y+a.h/2, 0.45, a.label, 1);
  });

  // ── STAGING + SUPPORT ─────────────────────────────────────────────────────
  (fp.stagingRects||[]).forEach(s2 => {
    rect('STAGING', s2.x, s2.y, s2.x+s2.w, s2.y+s2.h);
    if (s2.label)    txt('TEXT', s2.x+s2.w/2, s2.y+s2.h/2,     0.6, s2.label, 1);
    if (s2.subLabel) txt('TEXT', s2.x+s2.w/2, s2.y+s2.h/2+0.9, 0.4, s2.subLabel, 1);
  });
  (fp.supportRects||[]).forEach(s2 => {
    rect('SUPPORT', s2.x, s2.y, s2.x+s2.w, s2.y+s2.h);
    if (s2.label) txt('TEXT', s2.x+s2.w/2, s2.y+s2.h/2, 0.55, s2.label, 1);
  });

  // ── DOCK DOORS ────────────────────────────────────────────────────────────
  const DW_ = fp.doorW || 3.5;
  (fp.dockDoors||[]).forEach(dr => {
    if (dr.side === 'south' || dr.side === 'north') {
      rect('DOCKS', dr.x, dr.y-0.25, dr.x+DW_, dr.y+0.25);
      txt('DOCKS', dr.x+DW_/2, dr.y + (dr.side==='south'?1.1:-0.7), 0.45, dr.label, 1);
    } else {
      rect('DOCKS', dr.x-0.25, dr.y, dr.x+0.25, dr.y+DW_);
      txt('DOCKS', dr.x-1.6, dr.y+DW_/2, 0.45, dr.label);
    }
  });

  // ── OVERALL DIMENSIONS ────────────────────────────────────────────────────
  const off = 1.6;
  line('DIMS', 0, -off, WW, -off);
  line('DIMS', 0, -off-0.35, 0, -off+0.35);
  line('DIMS', WW, -off-0.35, WW, -off+0.35);
  txt('DIMS', WW/2, -off-0.6, 0.75, WW.toFixed(2)+' m', 1);

  line('DIMS', WW+off, 0, WW+off, WL);
  line('DIMS', WW+off-0.35, 0, WW+off+0.35, 0);
  line('DIMS', WW+off-0.35, WL, WW+off+0.35, WL);
  txt('DIMS', WW+off+0.55, WL/2, 0.75, WL.toFixed(2)+' m');

  // ── TITLE BLOCK ───────────────────────────────────────────────────────────
  const rackArea = (rackConfig||[]).reduce((s2,cf)=>s2+(parseFloat(cf.area)||0),0);
  const info = [
    'DENSICUBE WAREHOUSE LAYOUT',
    'Envelope: ' + WW.toFixed(2) + ' m x ' + WL.toFixed(2) + ' m',
    'Gross area: ' + Math.round(WW*WL) + ' sq m',
    'Rack area: ' + rackArea.toFixed(1) + ' sq m',
    'Clear height: ' + (params.clearH||'-') + ' m',
    'Aisle width: ' + (params.aisleW||'-') + ' m',
    'Units: METRES (1 unit = 1 m)',
    'Generated: ' + new Date().toISOString().slice(0,10),
  ];
  info.forEach((t,i) => txt('TEXT', 0, WL + 4 + (info.length-i)*1.0, i===0?0.9:0.55, t));

  d('0','ENDSEC');
  d('0','EOF');

  // CRLF line endings: what AutoCAD writes, and what strict readers expect
  const blob = new Blob([L.join('\r\n') + '\r\n'],
    {type:'application/dxf;charset=us-ascii'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = 'Warehouse_Plan_' + WW.toFixed(1) + 'x' + WL.toFixed(1) + 'm.dxf';
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); }, 300);
}


function exportExcel(analysis, design, params, rackConfig, binOverrides) {
  const wb   = XLSX.utils.book_new();
  // Effective (user-edited) container dimensions + location scaling
  const dimsOf = (k) => {
    const ph = binPhysFor(k, binOverrides);
    if (ph) return ph[0]+' x '+ph[1]+' x '+ph[2]+' mm';
    return (BIN_CATALOG[k] && BIN_CATALOG[k].dims) || '-';
  };
  const xLocScales = binLocScales(analysis, binOverrides);
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

  // Sheet 3: SKU Slotting Detail — only export SKUs with stock (inventory-provided)
  // If no inventory pasted, export all master SKUs
  const hasInventory = slotted.some(r=>r.stock>0);
  const slottedForExport = hasInventory ? slotted.filter(r=>r.stock>0) : slotted;
  XLSX.utils.book_append_sheet(wb, ws([
    ['SKU SLOTTING DETAIL'],[],
    ['SKU Code','L (mm)','W (mm)','H (mm)','Vol (cm³)','Max Dim','Long?',
     'Pick Lines','Velocity Band','Size Band','Combined Band',
     'Bin/Container','Location Size (mm)','Rack Type','Storage Zone','Units/Bin','Stock','Locs Required'],
    ...slottedForExport.map(r=>[r.sku,r.L,r.W,r.H,r.volCm3.toFixed(0),r.maxDim,r.isLong?'YES':'',
      r.pickLines,r.vb,r.sb,`${r.vb}-${r.sb}`,r.binName,
      dimsOf(r.bin),
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
    ...Object.values(binGroups).flatMap(b => {
      const vs = binVariantsFor(b.bin, binOverrides, b.locs);
      return vs.map(v => [
        b.binName + (vs.length>1 ? (' - ' + v.label) : ''),
        v.phys ? (v.phys[0]+' x '+v.phys[1]+' x '+v.phys[2]+' mm') : dimsOf(b.bin),
        v.locs,
        v.locs, // 1 container per location
        b.bin==='LONG' ? 'Size per item - cantilever slot'
          : (vs.length>1 ? 'Size variant of ' + b.bin : 'One container per location'),
      ]);
    }),
    [],
    ['TOTAL CONTAINERS REQUIRED','',
      Object.values(binGroups).reduce((s,b)=>s+binVariantsFor(b.bin,binOverrides,b.locs).reduce((t,v)=>t+v.locs,0),0),
      Object.values(binGroups).reduce((s,b)=>s+binVariantsFor(b.bin,binOverrides,b.locs).reduce((t,v)=>t+v.locs,0),0),''],
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
    ...(function(){
      // Envelope actually produced by the floor plan (aisles, cross aisles, staging bands)
      try {
        const L = buildFloorPlanLayout(design, params, rackConfig||[], analysis, false);
        if (!L || !(L.actualWW>0) || !(L.actualWL>0)) return [];
        const lw = Math.round(L.actualWW*10)/10, ll = Math.round(L.actualWL*10)/10;
        return [
          [],
          ['AS-DRAWN LAYOUT ENVELOPE'],
          ['Layout Dimensions','',lw+'m × '+ll+'m',
            (lw*3.281).toFixed(0)+'ft × '+(ll*3.281).toFixed(0)+'ft',''],
          ['Layout Gross Floor Area','',Math.round(lw*ll),toSqFt(Math.round(lw*ll)),''],
          ['Rack Area (sum of configs)','',
            Math.round((rackConfig||[]).reduce((s,cf)=>s+(parseFloat(cf.area)||0),0)*10)/10,'',''],
        ];
      } catch(e) { return []; }
    })(),
  ],[28,36,12,16,12]),'6. Area Summary');

  XLSX.writeFile(wb,`Warehouse_Design_${today.replace(/\//g,'-')}.xlsx`);
}

// ─── PPT EXPORT ───────────────────────────────────────────────────────────────
function exportPPT(analysis, design, params, rackConfig) {
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
  s1.addText(`${(analysis.metrics?.totSKUs||0).toLocaleString()} SKUs · ${(analysis.metrics?.totLocs||0).toLocaleString()} locations · ${design.wW}×${design.wL}m recommended`,
    {x:0.5,y:3.0,w:9,h:0.4,fontSize:13,color:'94A3B8',fontFace:'Calibri'});
  s1.addText(`Generated: ${today}`,{x:0.5,y:5.8,w:9,h:0.3,fontSize:10,color:'475569',fontFace:'Calibri'});

  // Slide 2: Key Metrics
  const s2=prs.addSlide(); hdr(s2,'Key Metrics','Headline numbers from SKU slotting analysis');
  const mStats=[
    [(analysis.metrics?.totSKUs||0).toLocaleString(),'Total SKUs',PINK],
    [(analysis.metrics?.totLocs||0).toLocaleString(),'Locations Required',BLUE],
    [(analysis.metrics?.totStock||0).toLocaleString(),'Current Stock Units',GREEN],
    [(analysis.metrics?.longCount||0),'Long/Awkward Items',AMBER],
    [`${design.wW}×${design.wL}m`,'Recommended Size',DARK],
    [(design.totalGrossArea||0).toLocaleString()+'m²','Gross Floor Area',GRAY],
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
    ...sbList.map(s=>{const n=(analysis?.matrix||{})[`${v}-${s}`]||0;return{text:n?String(n):'—',options:{align:'center',color:n?DARK:'9CA3AF'}};}),
    {text:String(sbList.reduce((sum,s)=>sum+((analysis.matrix||{})[`${v}-${s}`]||0),0)),options:{align:'center',bold:true}},
  ]);
  s3.addTable([matHdr,...matRows],{x:0.5,y:1.3,w:9.1,colW:[1.8,1.2,1.2,1.2,1.2,1.2,1.3],
    fontSize:11,border:{type:'solid',color:'E2E8F0',pt:1},rowH:0.45,autoPage:false});

  // Slide 4: Zone Breakdown
  const s4=prs.addSlide(); hdr(s4,'Zone Layout Plan','Storage zones by velocity — drives warehouse layout');
  const zRows=Object.entries(analysis?.zoneSummary||{}).map(([z,v])=>[
    {text:ZONE_DEFS[z]?.label||z,options:{bold:true,color:DARK}},
    {text:ZONE_DEFS[z]?.desc||'',options:{color:GRAY,fontSize:9}},
    {text:String(v.skus),options:{align:'center'}},
    {text:(v.locs||0).toLocaleString(),options:{align:'center',bold:true}},
    {text:(v.pickLines||0).toLocaleString(),options:{align:'center'}},
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
  const rackRows2=Object.entries(analysis?.rackSummary||{}).map(([rk,rv])=>[
    {text:RACK_DEFS[rk]?.name||rk,options:{bold:true,color:DARK}},
    {text:RACK_DEFS[rk]?.desc||'',options:{color:GRAY,fontSize:9}},
    {text:(rv.skus||0).toLocaleString(),options:{align:'center'}},
    {text:r(v.locs||0).toLocaleString(),options:{align:'center',bold:true}},
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
    ...(function(){
      try {
        const L = buildFloorPlanLayout(design, params, rackConfig||[], analysis, false);
        if (!L || !(L.actualWW>0) || !(L.actualWL>0)) return [];
        const lw = Math.round(L.actualWW*10)/10, ll = Math.round(L.actualWL*10)/10;
        return [['As-Drawn Layout',lw+'m × '+ll+'m ('+Math.round(lw*ll).toLocaleString()+'m²)']];
      } catch(e) { return []; }
    })(),
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


// ─── USER-DEFINED STORAGE CALCULATION ────────────────────────────────────────
function calcUserDefinedStorage(slotted, userBins, userRacks, params) {
  const validBins  = userBins.filter(b=>parseFloat(b.L)>0&&parseFloat(b.W)>0&&parseFloat(b.H)>0);
  const validRacks = userRacks.filter(r=>parseFloat(r.bayW)>0&&parseFloat(r.bayD)>0);
  if (!validBins.length) return null;

  const aisleM = (parseFloat(params.aisleW)||1.2)/2; // shared aisle
  const LOC_SHARE_U = { VF:1,F:1,M:1,S:2,VS:4,NM:8 };

  // Sort user bins smallest to largest by volume
  const sortedBins = [...validBins].sort((a,b)=>
    parseFloat(a.L)*parseFloat(a.W)*parseFloat(a.H) -
    parseFloat(b.L)*parseFloat(b.W)*parseFloat(b.H));

  const ORIENTS = [[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];

  const fitInBin = (sL,sW,sH,bL,bW,bH) => {
    if (!sL||!sW||!sH) return true; // no dims → fits
    const s=[sL,sW,sH].sort((a,b)=>b-a), b=[bL,bW,bH].sort((a,b)=>b-a);
    return s[0]<=b[0]&&s[1]<=b[1]&&s[2]<=b[2];
  };

  const upbUser = (sL,sW,sH,sVol,bL,bW,bH,fill) => {
    const binVol = bL/10*bW/10*bH/10;
    if (sL>0&&sW>0&&sH>0) {
      const d=[sL,sW,sH];
      let maxL=0;
      ORIENTS.forEach(([x,y,z])=>{
        const n=Math.floor(bL/d[x])*Math.floor(bW/d[y])*Math.floor(bH/d[z]);
        if(n>maxL) maxL=n;
      });
      const byVol = sVol>0?Math.max(1,Math.floor(binVol*fill/sVol)):1;
      return Math.min(Math.max(1,maxL),byVol);
    }
    return sVol>0?Math.max(1,Math.floor(binVol*fill/sVol)):1;
  };

  // Per-SKU calculation
  const skuResults = slotted.map(s => {
    if (s.stock===0) return {...s,userBin:null,userUpb:0,userLocs:0,overflow:false,overflowQty:0};
    let chosenBin=null;
    for (const b of sortedBins) {
      const bL=parseFloat(b.L),bW=parseFloat(b.W),bH=parseFloat(b.H);
      if (fitInBin(s.L,s.W,s.H,bL,bW,bH)) { chosenBin={...b,bL,bW,bH,fill:parseFloat(b.fill)||0.55}; break; }
    }
    if (!chosenBin) return {...s,userBin:null,userUpb:0,userLocs:0,
      overflow:true,overflowQty:s.stock,overflowReason:'Exceeds all custom bin dimensions'};
    const upb   = upbUser(s.L,s.W,s.H,s.volCm3,chosenBin.bL,chosenBin.bW,chosenBin.bH,chosenBin.fill);
    const share = Math.min(LOC_SHARE_U[s.vb]||1,Math.max(1,upb));
    const raw   = Math.max(1,Math.ceil(s.stock/upb));
    const locs  = Math.max(1,Math.ceil(raw/share));
    const util  = Math.round((s.stock/(locs*upb))*100);
    return {...s,userBin:chosenBin,userBinName:chosenBin.name||`Bin ${chosenBin.id}`,
      userUpb:upb,userLocs:locs,overflow:false,utilPct:util};
  });

  // Bin group summaries
  const binGroupLocs={};
  skuResults.filter(s=>!s.overflow&&s.userLocs>0).forEach(s=>{
    const k=s.userBin?.id||0;
    binGroupLocs[k]=(binGroupLocs[k]||0)+s.userLocs;
  });

  // Rack calculations
  const rackResults = validRacks.map(rack=>{
    const rW=parseFloat(rack.bayW),rD=parseFloat(rack.bayD);
    const rH=parseFloat(rack.bayH)||2200;
    const lvl=parseInt(rack.levels)||1;
    const clr=50;
    const perBinType={};
    sortedBins.forEach(b=>{
      const bL=parseFloat(b.L),bW=parseFloat(b.W),bH=parseFloat(b.H);
      let best=0;
      ORIENTS.forEach(([x,y,z])=>{
        const d=[bL,bW,bH];
        const n=Math.floor(rW/d[x])*Math.floor(rD/d[y])*Math.floor(rH/(d[z]+clr));
        if(n>best) best=n;
      });
      const lpb=Math.max(0,best)*lvl;
      const needed=binGroupLocs[b.id]||0;
      const bays=lpb>0?Math.ceil(needed/lpb):0;
      const fp=(rW/1000)*(rD/1000);
      const aisleA=bays*(rW/1000)*aisleM;
      perBinType[b.id]={lpb,needed,bays,area:+(bays*fp+aisleA).toFixed(1)};
    });
    const tot=Object.values(perBinType);
    return{...rack,perBinType,totalBays:tot.reduce((s,v)=>s+v.bays,0),
      totalArea:+tot.reduce((s,v)=>s+v.area,0).toFixed(1)};
  });

  const overflow  = skuResults.filter(s=>s.overflow);
  const stored    = skuResults.filter(s=>!s.overflow&&s.userLocs>0);
  const totLocs   = stored.reduce((s,r)=>s+r.userLocs,0);
  const totStock  = stored.reduce((s,r)=>s+r.stock,0);
  const totCap    = stored.reduce((s,r)=>s+r.userLocs*r.userUpb,0);
  const totArea   = rackResults.reduce((s,r)=>s+r.totalArea,0);

  // Bin utilisation per bin
  const binUtil={};
  stored.forEach(s=>{
    const k=s.userBin?.id||0;
    if(!binUtil[k]) binUtil[k]={name:s.userBinName,locs:0,stock:0,cap:0};
    binUtil[k].locs+=s.userLocs;binUtil[k].stock+=s.stock;binUtil[k].cap+=s.userLocs*s.userUpb;
  });
  Object.values(binUtil).forEach(b=>{b.utilPct=b.cap>0?Math.round(b.stock/b.cap*100):0;});

  return{skuResults,rackResults,overflow,stored,totLocs,totArea,totStock,totCap,binUtil,
    overallUtil:totCap>0?Math.round(totStock/totCap*100):0};
}



// ─── FORWARD PICK + RESERVE STORAGE CALCULATION ───────────────────────────────
// Splits each SKU's stock between a forward pick face and reserve storage.
// Forward locations = ceil(dailyPicks × forwardDays / unitsPerBin)
// Reserve locations = ceil((totalStock - forwardStock) / unitsPerBin)
function calcForwardReserve(analysis, forwardRacks, reserveRacks, forwardDays, params, binOverrides) {
  const frLocScales = binLocScales(analysis, binOverrides);
  if (!analysis?.slotted?.length) return null;
  const CLEAR = 30;
  const ORIENTS = [[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];
  const aisleHalf = (parseFloat(params.aisleW)||3.0)/2;
  const fDays = Math.max(1, parseInt(forwardDays)||3);
  // Velocity → estimated daily picks if no order data
  const VEL_PICKS = {VF:25,F:12,M:5,S:2,VS:1,NM:0};
  // Zone assignment for forward (near dispatch) and reserve (back)
  const FWD_ZONE = {shelving:'golden',liveStorage:'golden',selective:'mid',
    doubleDeep:'mid',driveIn:'reserve',cantilever:'reserve',ground:'bulk'};
  const RES_ZONE = {shelving:'reserve',liveStorage:'reserve',selective:'bulk',
    doubleDeep:'bulk',driveIn:'bulk',cantilever:'long',ground:'bulk'};

  const vFwd  = forwardRacks.filter(r=>parseFloat(r.bayW)>0&&parseFloat(r.bayD)>0);
  const vRes  = reserveRacks.filter(r=>parseFloat(r.bayW)>0&&parseFloat(r.bayD)>0);

  // Helper: find best fitting rack from a list for a given bin
  const bestFit = (racks, bL, bW, bH) => {
    let bestLPB=0, bestCfg=null;
    racks.forEach(rk=>{
      const rW=parseFloat(rk.bayW), rD=parseFloat(rk.bayD);
      const rTH=parseFloat(rk.bayH)||2200;
      const lvl=Math.max(1,parseInt(rk.levels)||1);
      const shelfH=Math.floor(rTH/lvl);
      const isGround=(rk.rackType||'shelving')==='ground';
      if(isGround){
        const dims=[bL,bW,bH].slice().sort((a,b)=>a-b);
        const [d1,d2]=dims;
        const oA={aw:Math.floor(rW/d1),ad:Math.floor(rD/d2)};
        const oB={aw:Math.floor(rW/d2),ad:Math.floor(rD/d1)};
        const best=oA.aw*oA.ad>=oB.aw*oB.ad?oA:oB;
        if(!best.aw||!best.ad) return;
        const stack=Math.max(1,parseInt(rk.levels)||1);
        const lpb=best.aw*best.ad*stack;
        if(lpb>bestLPB){bestLPB=lpb;bestCfg={rk,aw:best.aw,ad:best.ad,stack,lvl:1,shelfH:dims[2],lpb,orient:'ground'};}
        return;
      }
      ORIENTS.forEach(([x,y,z])=>{
        const dim=[bL,bW,bH];
        const aw=Math.floor(rW/dim[x]),ad=Math.floor(rD/dim[y]);
        if(!aw||!ad) return;
        const stack=Math.floor((shelfH-CLEAR)/dim[z]);
        if(stack<1) return;
        const lpb=aw*ad*stack*lvl;
        if(lpb>bestLPB){bestLPB=lpb;bestCfg={rk,aw,ad,stack,lvl,shelfH,lpb,orient:x===0?'LW':'WL'};}
      });
    });
    return bestCfg;
  };

  const fwdCfgMap={}, resCfgMap={}, overflow=[];
  const fwdLocs={}, resLocs={};

  // Process each bin type from system analysis
  Object.entries(analysis?.binSummary||{}).forEach(([binKey,binInfo])=>{
    const ph=binPhysFor(binKey, binOverrides); if(!ph) return;
    const [bL,bW,bH]=ph;
    const frRaw=binInfo.locs||0;
    const frSc=frLocScales[binKey];
    const totalLocs = (frSc===undefined) ? frRaw
      : (frSc<=0 ? 0 : Math.max(1, Math.round(frRaw*frSc)));
    if(!totalLocs) return;

    // Estimate forward locations from slotted velocity data
    const skusForBin=(analysis.slotted||[]).filter(s=>s.bin===binKey);
    const totalDailyPicks=skusForBin.reduce((s,sku)=>{
      const picks=sku.dailyPicks||sku.avgPicks||(VEL_PICKS[sku.velocity]||2);
      const upb=sku.upb||1;
      return s+Math.ceil((picks*fDays)/upb);
    },0);
    const fwdLocsNeeded=Math.min(totalDailyPicks, totalLocs);
    const resLocsNeeded=Math.max(0, totalLocs-fwdLocsNeeded);

    // Find best forward rack
    if(fwdLocsNeeded>0 && vFwd.length>0){
      const fc=bestFit(vFwd,bL,bW,bH);
      if(fc){
        const rt=fc.rk.rackType||'shelving';
        const bays=Math.ceil(fwdLocsNeeded/fc.lpb);
        const area=+(bays*(parseFloat(fc.rk.bayW)/1000)*((parseFloat(fc.rk.bayD)/1000)+aisleHalf)).toFixed(1);
        if(!fwdCfgMap[`${binKey}-fwd`]){
          fwdCfgMap[`${binKey}-fwd`]={
            id:`fr-fwd-${binKey}`,rack:rt,bin:binKey,phase:'forward',
            rackName:fc.rk.name||'Forward Pick',binName:binInfo.name||binKey,
            binDims:[bL,bW,bH],bayW:parseFloat(fc.rk.bayW),bayD:parseFloat(fc.rk.bayD),
            shelfH:parseFloat(fc.rk.bayH)||2200,tierHeight:fc.shelfH,clearance:CLEAR,
            orientation:fc.orient,tiers:1,levels:fc.lvl,
            acrossW:fc.aw,acrossD:fc.ad,stackH:fc.stack,
            locsPerBay:fc.lpb,locsPerBayTotal:fc.lpb,
            locs:fwdLocsNeeded,baysNeeded:bays,area,feasible:true,
            zone:FWD_ZONE[rt]||'golden',
          };
        }
        fwdLocs[binKey]=fwdLocsNeeded;
      } else { overflow.push({binKey,phase:'forward',dims:`${bL}×${bW}×${bH}mm`,totalLocs:fwdLocsNeeded,reason:'No forward rack fits this bin'}); }
    }

    // Find best reserve rack
    if(resLocsNeeded>0 && vRes.length>0){
      const rc=bestFit(vRes,bL,bW,bH);
      if(rc){
        const rt=rc.rk.rackType||'selective';
        const bays=Math.ceil(resLocsNeeded/rc.lpb);
        const area=+(bays*(parseFloat(rc.rk.bayW)/1000)*((parseFloat(rc.rk.bayD)/1000)+aisleHalf)).toFixed(1);
        if(!resCfgMap[`${binKey}-res`]){
          resCfgMap[`${binKey}-res`]={
            id:`fr-res-${binKey}`,rack:rt,bin:binKey,phase:'reserve',
            rackName:rc.rk.name||'Reserve',binName:binInfo.name||binKey,
            binDims:[bL,bW,bH],bayW:parseFloat(rc.rk.bayW),bayD:parseFloat(rc.rk.bayD),
            shelfH:parseFloat(rc.rk.bayH)||6000,tierHeight:rc.shelfH,clearance:CLEAR,
            orientation:rc.orient,tiers:1,levels:rc.lvl,
            acrossW:rc.aw,acrossD:rc.ad,stackH:rc.stack,
            locsPerBay:rc.lpb,locsPerBayTotal:rc.lpb,
            locs:resLocsNeeded,baysNeeded:bays,area,feasible:true,
            zone:RES_ZONE[rt]||'reserve',
          };
        }
        resLocs[binKey]=resLocsNeeded;
      } else { overflow.push({binKey,phase:'reserve',dims:`${bL}×${bW}×${bH}mm`,totalLocs:resLocsNeeded,reason:'No reserve rack fits this bin'}); }
    }
  });

  const fwdCfgs=Object.values(fwdCfgMap);
  const resCfgs=Object.values(resCfgMap);
  const allCfgs=[...fwdCfgs,...resCfgs];
  const totFwdLocs=Object.values(fwdLocs).reduce((s,v)=>s+v,0);
  const totResLocs=Object.values(resLocs).reduce((s,v)=>s+v,0);
  const totFwdArea=fwdCfgs.reduce((s,c)=>s+(c.area||0),0);
  const totResArea=resCfgs.reduce((s,c)=>s+(c.area||0),0);

  return {fwdCfgs,resCfgs,allCfgs,overflow,
    totFwdLocs,totResLocs,totFwdArea,totResArea,fDays};
}

// ─── USER-DEFINED RACK CONFIG FROM SYSTEM BINS ───────────────────────────────
// TWO-PASS ASSIGNMENT:
//   Pass 1 — fit each bin in the APPROPRIATE rack type by affinity, then any rack as fallback
//   Pass 2 — bins that overflow + isLong items → automatically try ground racks
//
// BIN AFFINITY RULES (warehouse best-practice):
//   XS/S/M → manual-pick racks (shelving, liveStorage, cantilever)
//   L/XL   → pallet racks (selective, doubleDeep, driveIn)
//   LONG   → ground / cantilever
function calcUserRackConfigFromSystemBins(analysis, userRacks, params, binOverrides) {
  // Each size variant becomes its own pseudo bin band, so it gets its own config
  const _x = expandBinVariants(analysis, binOverrides);
  analysis = _x.analysis; binOverrides = _x.binOverrides;
  const uLocScales = binLocScales(analysis, binOverrides);
  if (!analysis?.binSummary || !Object.keys(analysis.binSummary).length) return null;
  const CLEAR    = 30;
  const ORIENTS  = [[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];
  const ZONE_MAP = {XS:'golden',S:'golden',M:'mid',L:'reserve',XL:'bulk',LONG:'bulk'};
  const aisleHalf = (parseFloat(params.aisleW)||3.0)/2;

  // Rack category sets
  const MANUAL_RACK_TYPES = new Set(['shelving','liveStorage','cantilever']);
  const PALLET_RACK_TYPES = new Set(['selective','doubleDeep','driveIn']);

  // Bin → preferred rack category
  // L (Stack Crate/Half-Pallet 800×600×400mm) → manual/shelving, not forklift pallet rack
  // XL (Standard Pallet 1200×1000×1200mm) → pallet rack (forklift required)
  const BIN_PREFERRED_CATEGORY = {
    XS:'manual', S:'manual', M:'manual', L:'manual',
    XL:'pallet',
    LONG:'ground',
  };

  // ── Per-rack-type beam/shelf clearance (mm per level) ──────────────────────
  // Beam clearance = space taken by shelf board/pallet beam per level
  // After dividing total bay height by levels, subtract this to get usable bin space
  const BEAM_CLR = {
    shelving:    50,   // shelf board (25mm) + label clearance
    liveStorage: 60,   // flow rail depth
    selective:  150,   // pallet beam (100mm) + top-of-pallet to beam headroom
    doubleDeep: 150,
    driveIn:    150,
    cantilever: 100,   // arm clearance
    ground:       0,   // no levels
  };

  // ── Orientation sets ──────────────────────────────────────────────────────
  // ALL standard bins (XS/S/M/L/XL) must always be upright — H is always vertical.
  // Only the L↔W footprint swap is allowed. Tipping a pallet or crate on its side
  // is never acceptable in a real warehouse.
  // Exception: LONG/ground storage uses cross-section method (handled separately).
  const UPRIGHT_ORIENTS = [[0,1,2],[1,0,2]]; // [L-W-H] and [W-L-H]; H (index 2) always vertical

  // Split user racks into ground vs regular
  const groundRacks  = userRacks.filter(r=>r.rackType==='ground'&&parseFloat(r.bayW)>0&&parseFloat(r.bayD)>0);
  const regularRacks = userRacks.filter(r=>r.rackType!=='ground'&&parseFloat(r.bayW)>0&&parseFloat(r.bayD)>0);
  if (!groundRacks.length && !regularRacks.length) return null;

  // ── Clearance per rack type ────────────────────────────────────────────────
  const RACK_BEAM_CLR = (rackType) => BEAM_CLR[rackType] ?? 75;

  // ── usable height per level ────────────────────────────────────────────────
  // levelH  = floor(totalBayH / levels)       — space allocated per level
  // usableH = levelH - beamClearance          — space available for bins to sit in
  // ── Auto-level calculation by rack type ─────────────────────────────────────
  // Returns { levels, stack, levelToLevel, firstLevelH, calcNote } from:
  //   clearH (total usable clear height user entered) and binH (bin vertical dimension)
  //
  // SHELVING:
  //   level-to-level = bin_H + shelf board (25mm) + pick headroom (30mm)
  //   stack per level = 1 (no bin stacking on shelves)
  //   levels = floor(clear_H / level_to_level)
  //
  // PALLET RACK (selective / doubleDeep / driveIn):
  //   level 1 sits on floor → no beam below: firstLevel = pallet_base(144) + load_H + MHE_clr(100)
  //   upper levels: + beam(100mm): levelToLevel = pallet_base + load_H + MHE_clr(100) + beam(100)
  //   levels = clear_H >= firstLevel ? 1 + floor((clear_H - firstLevel) / levelToLevel) : 0
  //   stack per level = 1 (one pallet high)
  //
  // CANTILEVER / LIVE STORAGE:
  //   level-to-level = bin_H + arm_clearance(50mm) + headroom(50mm)
  //   stack = 1
  const PALLET_BASE_H  = 144;  // mm — standard pallet base height
  const MHE_CLEARANCE  = 100;  // mm — forklift clearance above load to beam
  const PALLET_BEAM_H  = 100;  // mm — pallet beam height
  const SHELF_BOARD_H  =  25;  // mm — shelf board thickness
  const SHELF_HEADROOM =  30;  // mm — label/pick clearance above bin

  const calcLevels = (rackType, clearH, binH) => {
    const cH = Math.max(0, parseFloat(clearH) || 0);
    const bH = Math.max(1, binH);

    if (['shelving','liveStorage'].includes(rackType)) {
      const levelToLevel = bH + SHELF_BOARD_H + SHELF_HEADROOM;
      const levels = Math.max(0, Math.floor(cH / levelToLevel));
      return { levels, stack:1, levelToLevel, firstLevelH: levelToLevel,
        calcNote:`${bH}mm bin + ${SHELF_BOARD_H}mm shelf + ${SHELF_HEADROOM}mm headroom = ${levelToLevel}mm/level` };
    }
    if (['selective','doubleDeep','driveIn'].includes(rackType)) {
      const firstLevelH    = PALLET_BASE_H + bH + MHE_CLEARANCE;  // level 1 on floor — no beam
      const levelToLevel   = PALLET_BASE_H + bH + MHE_CLEARANCE + PALLET_BEAM_H;
      const levels = cH >= firstLevelH
        ? 1 + Math.max(0, Math.floor((cH - firstLevelH) / levelToLevel))
        : 0;
      return { levels, stack:1, levelToLevel, firstLevelH,
        calcNote:`Level 1 (floor): ${PALLET_BASE_H}mm pallet + ${bH}mm load + ${MHE_CLEARANCE}mm MHE = ${firstLevelH}mm | Upper levels: +${PALLET_BEAM_H}mm beam = ${levelToLevel}mm` };
    }
    if (rackType === 'cantilever') {
      const levelToLevel = bH + 50 + 50;  // bin + arm clr + headroom
      const levels = Math.max(0, Math.floor(cH / levelToLevel));
      return { levels, stack:1, levelToLevel, firstLevelH: levelToLevel,
        calcNote:`${bH}mm bin + 50mm arm clr + 50mm headroom = ${levelToLevel}mm/level` };
    }
    // Fallback
    return { levels:1, stack:1, levelToLevel:cH, firstLevelH:cH, calcNote:'' };
  };

  // ── Helper: best fit in a given set of racks ────────────────────────────────
  const bestFitInRacks = (racks, bL, bW, bH) => {
    let best=null, bestLPB=0;
    racks.forEach(rk=>{
      const rW    = parseFloat(rk.bayW);
      const rD    = parseFloat(rk.bayD);
      const clearH = parseFloat(rk.bayH) || 2200;

      if(rk.rackType === 'cantilever'){
        // Cantilever: item's LONG dimension goes ALONG the arm (depth direction).
        // Use c-prefixed names to avoid TDZ conflict with outer-block const declarations
        const cSorted = [bL,bW,bH].slice().sort((a,b)=>a-b);
        const cd1=cSorted[0], cd2=cSorted[1], cd3=cSorted[2];
        if(rD < cd3) return; // arm too short — item doesn't fit
        const coA = Math.floor(rW/cd1);
        const coB = Math.floor(rW/cd2);
        const caw = Math.max(coA, coB);
        if(!caw) return;
        const cad = Math.floor(rD/cd3)||1;
        const cLvlResult = calcLevels(rk.rackType, clearH, cd2>cd1?cd2:cd1);
        if(cLvlResult.levels<1) return;
        const clpb = caw * cad * cLvlResult.stack * cLvlResult.levels;
        if(clpb > bestLPB){
          bestLPB = clpb;
          best = {
            rk, aw:caw, ad:cad, stack:cLvlResult.stack, lvl:cLvlResult.levels, lpb:clpb,
            totalH:clearH, levelH:cLvlResult.levelToLevel, usableH:cd2, beamClr:0,
            levels:cLvlResult.levels, levelToLevel:cLvlResult.levelToLevel,
            firstLevelH:cLvlResult.firstLevelH, calcNote:cLvlResult.calcNote,
            orientDesc:`Frame W: ${caw} items × ${caw===coA?cd1:cd2}mm | Arm: ${cd3}mm along depth (${rD}mm arm)`,
            wDimMm:caw===coA?cd1:cd2, dDimMm:cd3, hDimMm:cd2>cd1?cd2:cd1,
            orient:'cantilever',
          };
        }
        return;
      }

      // All other racks — upright orientation only (H always vertical)
      // Use single named object (not destructuring) to avoid esbuild TDZ when
      // same names appear in the cantilever block above.
      const rl = calcLevels(rk.rackType, clearH, bH);
      if (rl.levels < 1) return;
      UPRIGHT_ORIENTS.forEach(([x,y,z])=>{
        const dim = [bL, bW, bH];
        const aw  = Math.floor(rW / dim[x]);
        const ad  = Math.floor(rD / dim[y]);
        if (!aw || !ad) return;
        const lpb = aw * ad * rl.stack * rl.levels;
        if (lpb > bestLPB) {
          bestLPB = lpb;
          const AXIS=['L','W','H'];
          const wDim=AXIS[x], dDim=AXIS[y];
          best = {
            rk, aw, ad, stack:rl.stack, lvl:rl.levels, lpb,
            totalH:clearH, levelH:rl.levelToLevel, usableH:bH, beamClr:0,
            levels:rl.levels, levelToLevel:rl.levelToLevel,
            firstLevelH:rl.firstLevelH, calcNote:rl.calcNote,
            orientDesc:`Bin ${wDim}(${dim[x]}mm)→width, ${dDim}(${dim[y]}mm)→depth, H(${bH}mm) vertical`,
            wDimMm:dim[x], dDimMm:dim[y], hDimMm:bH,
            orient: x===0&&y===1?'LW':'WL',
          };
        }
      });
    });
    return best;
  };

  // ── Helper: best fit in regular racks — preferred category first ─────────────
  const bestRegular = (bL,bW,bH,binKey) => {
    const pref = BIN_PREFERRED_CATEGORY[baseBinOf(binKey)];
    const preferredRacks = regularRacks.filter(rk=>
      pref==='manual' ? MANUAL_RACK_TYPES.has(rk.rackType) :
      pref==='pallet' ? PALLET_RACK_TYPES.has(rk.rackType) : false
    );
    if (preferredRacks.length>0) {
      const fc = bestFitInRacks(preferredRacks, bL,bW,bH);
      if (fc) return fc;
    }
    return bestFitInRacks(regularRacks, bL,bW,bH);
  };

  // Helper: best fit in ground racks (cross-section: 2 smallest dims as footprint)
  const bestGround = (bL,bW,bH) => {
    let best=null, bestLPB=0;
    groundRacks.forEach(rk=>{
      const rW=parseFloat(rk.bayW), rD=parseFloat(rk.bayD);
      const stackLayers=Math.max(1,parseInt(rk.levels)||1);
      const [d1,d2,d3]=[bL,bW,bH].slice().sort((a,b)=>a-b);
      const oA={aw:Math.floor(rW/d1),ad:Math.floor(rD/d2)};
      const oB={aw:Math.floor(rW/d2),ad:Math.floor(rD/d1)};
      const bo=oA.aw*oA.ad>=oB.aw*oB.ad?oA:oB;
      if(!bo.aw||!bo.ad) return;
      const lpb=bo.aw*bo.ad*stackLayers;
      if(lpb>bestLPB){bestLPB=lpb;
        best={rk,aw:bo.aw,ad:bo.ad,stack:stackLayers,lvl:1,
          totalH:parseFloat(rk.bayH)||500, levelH:0, usableH:0, beamClr:0,
          wDimMm:d1,dDimMm:d2,hDimMm:d3,
          orientDesc:`Cross-section ${d1}×${d2}mm footprint, ${d3}mm length protrudes`,
          lpb,orient:'ground'};}
    });
    return best;
  };

  // Bins flagged as LONG/odd in slotted data → auto-route to ground/cantilever if available
  const longBins=new Set((analysis.slotted||[]).filter(s=>s.isLong).map(s=>s.bin));

  // Declare before ALL forEach loops that use these arrays
  const uCfgs=[], overflowBins=[];

  const regularPass={}, groundPass={};

  Object.entries(analysis?.binSummary||{}).forEach(([binKey,binInfo])=>{
    // Resolve bin physical dimensions (skip LONG here — handled per-SKU in Pass 2)
    let bL, bW, bH;
    if (baseBinOf(binKey) === 'LONG') {
      // LONG has no single phys → always goes to ground pass for per-SKU fitting
      bL=0; bW=0; bH=0;
    } else {
      const ph = binPhysFor(binKey, binOverrides);
      if (!ph) return;
      [bL, bW, bH] = ph;
    }
    const rawLocs=binInfo.locs||0;
    const uSc=uLocScales[binKey];
    // uSc may legitimately be 0 (band excluded) — test for presence, not truthiness
    const totalLocs = (uSc===undefined) ? rawLocs
      : (uSc<=0 ? 0 : Math.max(1, Math.round(rawLocs*uSc)));
    if(!totalLocs) return;

    const isLongBin = longBins.has(binKey) || binKey==='LONG';

    if(isLongBin){
      // ── Priority for LONG goods: ──────────────────────────────────────
      // 1. Cantilever — purpose-built for long items, try first
      // 2. Ground storage — fallback (per-SKU fitting)
      // 3. Overflow — if neither selected or neither fits
      const cantileverRacks = regularRacks.filter(r=>r.rackType==='cantilever');
      const repL=bL||3000, repW=bW||300, repH=bH||200;

      if(cantileverRacks.length>0){
        const cantFit=bestFitInRacks(cantileverRacks, repL, repW, repH);
        if(cantFit){
          // Cantilever fits ✓
          regularPass[binKey]={fc:cantFit,bL:repL,bW:repW,bH:repH,totalLocs,binInfo};
          return;
        }
        // Cantilever selected but item doesn't fit → fall through to ground
      }

      if(groundRacks.length>0){
        // Ground storage — per-SKU cross-section fitting
        groundPass[binKey]={bL,bW,bH,totalLocs,binInfo,forceGround:true,perSku:true};
        return;
      }

      // Neither cantilever nor ground → overflow
      overflowBins.push({binKey,binName:binInfo.name||binKey,totalLocs,dims:'per item',
        reason:'No cantilever or ground storage rack selected. Add Cantilever or Ground Storage.'});
      return;
    }

    if(regularRacks.length>0){
      const regFit=bestRegular(bL,bW,bH,binKey);
      if(regFit){
        const fittedRackType = regFit.rk.rackType;
        const binPref = BIN_PREFERRED_CATEGORY[baseBinOf(binKey)];

        // ── Pallet-ization: manual bin (XS/S/M/L) falling back to pallet rack ──
        // When no shelving is selected, bins are grouped onto pallets.
        // 1 pallet = 1 SKU, bins stacked on a standard 1200×1000mm pallet.
        if(binPref==='manual' && PALLET_RACK_TYPES.has(fittedRackType)){
          const PALLET_L=1200, PALLET_W=1000, PALLET_LOAD_H=1200; // standard pallet
          // Bins per pallet — try both L↔W orientations, H always up
          const o1=Math.floor(PALLET_L/bL)*Math.floor(PALLET_W/bW)*Math.floor(PALLET_LOAD_H/bH);
          const o2=Math.floor(PALLET_L/bW)*Math.floor(PALLET_W/bL)*Math.floor(PALLET_LOAD_H/bH);
          const binsPerPallet=Math.max(1,Math.max(o1,o2));
          // Per-SKU pallet positions: 1 pallet per SKU, ceil(sku_locs / binsPerPallet)
          const skusForBin=(analysis.slotted||[]).filter(s=>s.bin===binKey&&(s.stock||0)>0);
          const totalPalletPositions=skusForBin.reduce((sum,s)=>
            sum+Math.max(1,Math.ceil((s.locsReq||1)/binsPerPallet)),0);
          // Re-fit using pallet footprint dimensions in the rack bay
          const palletRacks=regularRacks.filter(r=>PALLET_RACK_TYPES.has(r.rackType));
          const palletRackFit=bestFitInRacks(palletRacks, PALLET_L, PALLET_W, PALLET_LOAD_H);
          if(palletRackFit&&totalPalletPositions>0){
            regularPass[binKey]={
              fc:palletRackFit, bL, bW, bH,
              totalLocs:totalPalletPositions, binInfo,
              isPalletized:true, binsPerPallet,
              palletL:PALLET_L, palletW:PALLET_W, palletLoadH:PALLET_LOAD_H,
              skuCount:skusForBin.length,
              originalBinLocs:totalLocs,
            };
            return;
          }
        }

        regularPass[binKey]={fc:regFit,bL,bW,bH,totalLocs,binInfo}; return;
      }
    }
    groundPass[binKey]={bL,bW,bH,totalLocs,binInfo,forceGround:false,perSku:false};
  });

  // Pass 2: assign ground-pass bins to ground racks
  Object.entries(groundPass).forEach(([binKey,{bL,bW,bH,totalLocs,binInfo,forceGround,perSku}])=>{
    if(groundRacks.length>0){
      const rk = groundRacks[0]; // use first ground rack
      const rW=parseFloat(rk.bayW), rD=parseFloat(rk.bayD);
      const stackLayers=Math.max(1,parseInt(rk.levels)||1);

      if(perSku && (baseBinOf(binKey)==='LONG'||longBins.has(baseBinOf(binKey)))){
        // ── PER-SKU FITTING for LONG goods ─────────────────────────────────
        const longSkus=(analysis.slotted||[]).filter(s=>s.bin===binKey&&s.stock>0&&s.L>0&&s.W>0&&s.H>0);
        const fittedSkus=[], unfittedSkus=[];
        let sumLocs=0, sumBays=0;

        longSkus.forEach(s=>{
          // Cross-section: sort dims → 2 smallest = footprint, largest protrudes
          const [d1,d2,d3]=[s.L,s.W,s.H].slice().sort((a,b)=>a-b);
          const oA={aw:Math.floor(rW/d1),ad:Math.floor(rD/d2)};
          const oB={aw:Math.floor(rW/d2),ad:Math.floor(rD/d1)};
          const bo=oA.aw*oA.ad>=oB.aw*oB.ad?oA:oB;

          if(!bo.aw||!bo.ad){
            unfittedSkus.push({...s, reason:`Cross-section (${d1}×${d2}mm) exceeds bay (${rW}×${rD}mm)`});
            return;
          }
          const locsPerSlot=bo.aw*bo.ad*stackLayers; // items per bay
          const locsNeeded=s.locsReq||Math.ceil((s.stock||1)/1);
          const baysNeeded=Math.ceil(locsNeeded/locsPerSlot);
          const area=+(baysNeeded*(rW/1000)*((rD/1000)+aisleHalf)).toFixed(1);
          fittedSkus.push({
            sku:s.sku, desc:s.desc||s.name||'', vb:s.vb||s.velocity,
            L:s.L, W:s.W, H:s.H, stock:s.stock||0,
            d1,d2,d3, aw:bo.aw, ad:bo.ad,
            locsPerSlot, locsNeeded, baysNeeded, area,
          });
          sumLocs+=locsNeeded;
          sumBays+=baysNeeded;
        });

        const totalArea=+(sumBays*(rW/1000)*((rD/1000)+aisleHalf)).toFixed(1);

        if(fittedSkus.length>0||unfittedSkus.length>0){
          uCfgs.push({
            id:`u-${binKey}`,rack:'ground',bin:binKey,
            rackName:rk.name||'Ground Storage',binName:binInfo.name||binKey,
            binDims:null, bayW:rW, bayD:rD,
            shelfH:0,tierHeight:0,clearance:0,levelH:0,usableH:0,beamClr:0,
            orientation:'ground',tiers:1,levels:stackLayers,
            acrossW:0,acrossD:0,stackH:stackLayers,
            locsPerBay:0,locsPerBayTotal:0,
            locs:sumLocs,baysNeeded:sumBays,area:totalArea,feasible:true,zone:'bulk',
            autoAssigned:`Long goods → per-SKU cross-section fitting in ground storage`,
            perSkuFitted:fittedSkus,
            perSkuUnfitted:unfittedSkus,
            isPerSku:true,
            o1:{acrossW:0,acrossD:0,feasible:true,levels:stackLayers,locsPerBay:0},
            o2:{acrossW:0,acrossD:0,feasible:true,levels:stackLayers,locsPerBay:0},
          });
          // Push unfitted as overflow
          if(unfittedSkus.length>0){
            overflowBins.push({
              binKey, binName:binInfo.name||binKey,
              totalLocs:unfittedSkus.reduce((s,u)=>s+(u.locsReq||1),0),
              dims:'per SKU',
              skuList:unfittedSkus,
              reason:`${unfittedSkus.length} SKU(s) cross-section exceeds ground bay (${rW}×${rD}mm)`,
            });
          }
        } else {
          // No long items with stock — skip
        }
        return;
      }

      // Regular ground storage (non-LONG) — cross-section with representative dims
      const gc = {
        rk,
        ...(()=>{
          const [d1,d2,d3]=[bL,bW,bH].slice().sort((a,b)=>a-b);
          const oA={aw:Math.floor(rW/d1),ad:Math.floor(rD/d2)};
          const oB={aw:Math.floor(rW/d2),ad:Math.floor(rD/d1)};
          const bo=oA.aw*oA.ad>=oB.aw*oB.ad?oA:oB;
          return {aw:bo.aw,ad:bo.ad,stack:stackLayers,lpb:bo.aw*bo.ad*stackLayers,
            wDimMm:bo===oA?d1:d2,dDimMm:bo===oA?d2:d1,hDimMm:d3,
            totalH:parseFloat(rk.bayH)||500,
            orientDesc:`Cross-section ${bo===oA?d1:d2}×${bo===oA?d2:d1}mm footprint, ${d3}mm length protrudes`};
        })()
      };
      if(!gc.aw||!gc.ad){
        const allR=[...regularRacks,...groundRacks];
        const maxW=Math.max(...allR.map(r=>parseFloat(r.bayW)||0),0);
        const maxD=Math.max(...allR.map(r=>parseFloat(r.bayD)||0),0);
        overflowBins.push({binKey,binName:binInfo.name||binKey,totalLocs,dims:`${bL}×${bW}×${bH}mm`,
          reason:`Cross-section exceeds ground bay (${maxW}×${maxD}mm)`});
        return;
      }
      const bays=Math.ceil(totalLocs/gc.lpb);
      const area=+(bays*(rW/1000)*((rD/1000)+aisleHalf)).toFixed(1);
      uCfgs.push({
        id:`u-${binKey}`,rack:'ground',bin:binKey,
        rackName:rk.name||'Ground Storage',binName:binInfo.name||binKey,
        binDims:[bL,bW,bH],bayW:rW,bayD:rD,
        shelfH:gc.totalH||500,tierHeight:0,clearance:0,
        levelH:0,usableH:0,beamClr:0,
        orientDesc:gc.orientDesc,wDimMm:gc.wDimMm,dDimMm:gc.dDimMm,hDimMm:gc.hDimMm,
        orientation:'ground',tiers:1,levels:gc.stack,
        acrossW:gc.aw,acrossD:gc.ad,stackH:gc.stack,
        locsPerBay:gc.lpb,locsPerBayTotal:gc.lpb,
        locs:totalLocs,baysNeeded:bays,area,feasible:true,zone:'bulk',
        autoAssigned:forceGround?'Odd-shaped → auto-routed to ground':'Overflow from regular racks → ground',
        o1:{acrossW:gc.aw,acrossD:gc.ad,feasible:true,levels:gc.stack,locsPerBay:gc.lpb},
        o2:{acrossW:gc.aw,acrossD:gc.ad,feasible:true,levels:gc.stack,locsPerBay:gc.lpb},
      });
      return;
    }
    // No ground rack available — overflow
  });

  // Finalise regular rack assignments
  Object.entries(regularPass).forEach(([binKey,entry])=>{
    const {fc,bL,bW,bH,totalLocs,binInfo,
      isPalletized,binsPerPallet,palletL,palletW,palletLoadH,skuCount,originalBinLocs}=entry;
    const rt=fc.rk.rackType||'shelving';
    const rW=parseFloat(fc.rk.bayW),rD=parseFloat(fc.rk.bayD);
    const bays=Math.ceil(totalLocs/fc.lpb);
    const area=+(bays*(rW/1000)*((rD/1000)+aisleHalf)).toFixed(1);
    // All bins are always upright (H vertical) — only L↔W swap considered
    const {usableH:uH} = fc;
    const o1aw=Math.floor(rW/bL),o1ad=Math.floor(rD/bW),o1st=uH>0?Math.floor(uH/bH):0;
    const o2aw=Math.floor(rW/bW),o2ad=Math.floor(rD/bL),o2st=uH>0?Math.floor(uH/bH):0;
    uCfgs.push({
      id:`u-${binKey}`,rack:rt,bin:binKey,
      rackName:fc.rk.name||'Custom Rack',binName:binInfo.name||binKey,
      aisle:fc.rk.aisle||fc.rk.pickingAisle||null,
      binDims:[bL,bW,bH],bayW:rW,bayD:rD,
      shelfH:fc.totalH,tierHeight:fc.levelH,clearance:0,
      levelH:fc.levelH,usableH:fc.usableH,beamClr:0,
      firstLevelH:fc.firstLevelH, calcNote:fc.calcNote,
      orientDesc:fc.orientDesc,wDimMm:fc.wDimMm,dDimMm:fc.dDimMm,hDimMm:fc.hDimMm,
      orientation:fc.orient,tiers:1,levels:fc.lvl,
      acrossW:fc.aw,acrossD:fc.ad,stackH:fc.stack,
      locsPerBay:fc.lpb,locsPerBayTotal:fc.lpb,
      locs:totalLocs,baysNeeded:bays,area,feasible:true,zone:ZONE_MAP[baseBinOf(binKey)]||'golden',
      // Pallet-ization metadata (when manual bins fall back to pallet rack)
      isPalletized:!!isPalletized,
      binsPerPallet:binsPerPallet||null,
      palletL:palletL||null, palletW:palletW||null, palletLoadH:palletLoadH||null,
      skuCount:skuCount||null, originalBinLocs:originalBinLocs||null,
      // LW and WL orientations for display (based on pallet dims if palletized, bin dims otherwise)
      o1:{acrossW:o1aw,acrossD:o1ad,feasible:o1aw>0&&o1ad>0&&o1st>0,
          levels:o1st,locsPerBay:o1aw*o1ad*o1st*fc.lvl,
          desc:isPalletized?`${palletL}mm wide × ${palletW}mm deep × ${palletLoadH}mm tall (pallet)`
               :`${bL}mm wide × ${bW}mm deep × ${bH}mm tall`},
      o2:{acrossW:o2aw,acrossD:o2ad,feasible:o2aw>0&&o2ad>0&&o2st>0,
          levels:o2st,locsPerBay:o2aw*o2ad*o2st*fc.lvl,
          desc:isPalletized?`${palletW}mm wide × ${palletL}mm deep × ${palletLoadH}mm tall (pallet)`
               :`${bW}mm wide × ${bL}mm deep × ${bH}mm tall`},
    });
  });

  uCfgs.sort((a,b)=>(a.rack==='ground'?1:-1)-(b.rack==='ground'?1:-1));
  return {uCfgs,overflowBins};
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
  const [preferredBins, setPreferredBins] = useState(['S','M','L','XL']); // default: 4 types, no XS
  // Storage design mode
  const [storageMode, setStorageMode] = useState('system'); // 'system' | 'user'
  // User-defined racks only — bins are fixed from system analysis (read-only)
  const [userRacks, setUserRacks] = useState([
    {id:1,name:'Custom Rack 1',rackType:'shelving',bayW:'',bayD:'',bayH:'',levels:''},
  ]);
  // User-defined results
  const [userResult, setUserResult] = useState(null);
  const [userDesign, setUserDesign] = useState(null);
  const [userRackConfig, setUserRackConfig] = useState(null);
  const [userOverflowBins, setUserOverflowBins] = useState([]);
  const [userLoading, setUserLoading] = useState(false);

  // User Defined 4-step wizard state
  const [udStep,            setUdStep]            = useState(1);
  const [selectedRackTypes, setSelectedRackTypes] = useState(new Set());
  const [udRackDefs,        setUdRackDefs]        = useState({});

  // ── Forward Pick + Reserve mode state ──────────────────────────────────────
  const [forwardDays,  setForwardDays]  = useState('3');
  const [forwardRacks, setForwardRacks] = useState([
    {id:1,name:'Forward Pick',rackType:'shelving',bayW:'900',bayD:'600',bayH:'2200',levels:'4'},
  ]);
  const [reserveRacks, setReserveRacks] = useState([
    {id:1,name:'Reserve SPR',rackType:'selective',bayW:'2700',bayD:'1100',bayH:'6000',levels:'4'},
  ]);
  const [frResult,     setFrResult]     = useState(null);
  const [frDesign,     setFrDesign]     = useState(null);
  const [frRackConfig, setFrRackConfig] = useState(null);
  const [frOverflow,   setFrOverflow]   = useState([]);

  // Results
  const [analysis,  setAnalysis]  = useState(null);
  const [rackConfig,setRackConfig]= useState(null);
  const [design,    setDesign]    = useState(null);
  const [configConfirmed,setConfigConfirmed]=useState(false);
  // ── Bin / Pallet size + count editor (before rack selection) ──
  const [binOverrides,  setBinOverrides]  = useState(null); // {XS:{L,W,H,locs}, ...}
  const [binEditsOpen,  setBinEditsOpen]  = useState(true);
  const [binEditsDirty, setBinEditsDirty] = useState(false);
  const [binEditsError, setBinEditsError] = useState('');
  const [udViewMode,  setUdViewMode]  = useState('2d'); // user defined — default 2D (no WebGL risk)
  // ── Dimension measuring tool ──
  const [measureOn,    setMeasureOn]    = useState(false);
  const [measurePts,   setMeasurePts]   = useState([]);   // in-progress point(s), metres
  const [measurements, setMeasurements] = useState([]);   // completed [ptA, ptB] pairs
  const [snapOn,       setSnapOn]       = useState(true);  // ortho lock + edge snapping
  const [floorPlanFS, setFloorPlanFS] = useState(false); // fullscreen 2D plan
  const plan2DRef   = useRef(null);
  const userPlanRefObj = useRef(null); // user-defined mode plan container
  const planScrollRef = useRef(null); // plan scroll container
  const fsScrollRef   = useRef(null); // fullscreen scroll container
  const [loading,   setLoading]   = useState(false);
  const [progress,  setProgress]  = useState(0);   // 0-100
  const [progressMsg,setProgressMsg]= useState('');
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


  // Auto-recalculate user rack config whenever rack dims or analysis changes
  useEffect(()=>{
    if(storageMode!=='user') return;
    if(!analysis?.binSummary||!Object.keys(analysis.binSummary).length) return;
    const hasRacks=userRacks.some(r=>parseFloat(r.bayW)>0&&parseFloat(r.bayD)>0);
    if(!hasRacks){
      setUserRackConfig(null);setUserOverflowBins([]);
      setUserDesign(null);setUserResult(null); return;
    }
    const res=calcUserRackConfigFromSystemBins(analysis,userRacks,params,binOverrides);
    if(!res) return;
    setUserRackConfig(res.uCfgs.length>0?res.uCfgs:null);
    setUserOverflowBins(res.overflowBins||[]);
    setUserResult({stored:[],overflow:res.overflowBins||[],
      totLocs:res.uCfgs.reduce((s,x)=>s+x.locs,0),
      totArea:res.uCfgs.reduce((s,x)=>s+(x.area||0),0),
      totStock:0,totCap:0,overallUtil:0,binUtil:{},rackResults:[]});
    const ca={},cza={};
    res.uCfgs.forEach(x=>{ca[x.rack]=(ca[x.rack]||0)+(x.area||0);cza[x.zone]=(cza[x.zone]||0)+(x.area||0);});
    if(!Object.keys(ca).length) ca.shelving=50;
    try{setUserDesign(calcWarehouseSize(analysis,params,ca,cza));}catch(e){}
  },[userRacks,analysis,storageMode]); // eslint-disable-line


  // Auto-recalculate FR config when racks/days/analysis change
  useEffect(()=>{
    if(storageMode!=='fr') return;
    if(!analysis?.binSummary||!Object.keys(analysis.binSummary).length) return;
    const hasF=forwardRacks.some(r=>parseFloat(r.bayW)>0);
    const hasR=reserveRacks.some(r=>parseFloat(r.bayW)>0);
    if(!hasF&&!hasR) return;
    const res=calcForwardReserve(analysis,forwardRacks,reserveRacks,forwardDays,params,binOverrides);
    if(!res) return;
    setFrResult(res);
    setFrRackConfig(res.allCfgs);
    setFrOverflow(res.overflow);
    const cza={};
    res.allCfgs.forEach(cfg=>{cza[cfg.zone]=(cza[cfg.zone]||0)+(cfg.area||0);});
    if(!Object.keys(cza).length) cza.golden=50;
    try{setFrDesign(calcWarehouseSize(analysis,params,null,cza));}catch(e){}
  },[forwardRacks,reserveRacks,forwardDays,analysis,storageMode]); // eslint-disable-line

  // Yield to browser — keeps UI responsive during heavy processing
  const tick = () => new Promise(r => setTimeout(r, 0));

  const runAll = async () => {
    setError(''); setLoading(true); setProgress(0); setProgressMsg('Preparing data…');
    await tick();
    try {
      if (!masterText.trim()) throw new Error('Paste Master SKU data first.');

      // ── Step 1: Parse master SKU data ────────────────────────────────────
      setProgressMsg('Parsing SKU master data…'); setProgress(5); await tick();
      const masterRows = parseTSV(masterText);
      const mData = isHeaderRow(masterRows[0]) ? masterRows.slice(1) : masterRows;
      if (!mData.length) throw new Error('No valid SKU rows found in Master SKU data.');

      // ── Step 2: Parse order + inventory data ─────────────────────────────
      setProgressMsg(`Parsing order & inventory data… (${mData.length.toLocaleString()} SKUs)`);
      setProgress(15); await tick();
      const orderRows = orderText.trim() ? parseTSV(orderText).filter(r=>r[2]||r[0]) : [];
      const oData = orderRows.length && isHeaderRow(orderRows[0]) ? orderRows.slice(1) : orderRows;
      const invRows = invText.trim() ? parseTSV(invText).filter(r=>r[0]) : [];
      const iData = invRows.length && isHeaderRow(invRows[0]) ? invRows.slice(1) : invRows;

      // ── Step 3: Run analysis in chunks for large datasets ─────────────────
      setProgressMsg(`Classifying ${mData.length.toLocaleString()} SKUs by velocity & size…`);
      setProgress(20); await tick();

      let a;
      if (mData.length <= 20000) {
        // Small dataset — run synchronously
        a = runAnalysis(mData, oData, iData, params, preferredBins);
        setProgress(70);
      } else {
        // Large dataset — chunk processing with progress updates
        // Build velocity + stock maps first (fast, needed by all chunks)
        setProgressMsg('Building velocity maps…'); await tick();
        const velMap = {}, stockMap = {};
        oData.forEach(r => {
          const sku = String(r[1]||r[0]||'').trim();
          if (sku) velMap[sku] = (velMap[sku]||0) + (parseFloat(r[3]||r[2])||1);
        });
        iData.forEach(r => {
          const sku = String(r[0]||'').trim();
          if (sku) stockMap[sku] = parseFloat(r[2]||r[1])||0;
        });
        setProgress(30); await tick();

        // Process SKUs in chunks of 10,000
        const CHUNK = 10000;
        const chunks = Math.ceil(mData.length / CHUNK);
        const partials = [];
        for (let ci = 0; ci < chunks; ci++) {
          const slice = mData.slice(ci * CHUNK, (ci + 1) * CHUNK);
          setProgressMsg(`Processing SKUs ${(ci*CHUNK+1).toLocaleString()}–${Math.min((ci+1)*CHUNK,mData.length).toLocaleString()} of ${mData.length.toLocaleString()}…`);
          setProgress(30 + Math.round((ci/chunks)*40));
          await tick();
          // Run analysis on this chunk — pass pre-built maps to avoid re-scanning orders
          partials.push(runAnalysis(slice, oData, iData, params, preferredBins, velMap, stockMap));
        }
        setProgress(70); await tick();

        // Merge chunk results
        setProgressMsg('Merging results…'); await tick();
        a = mergeAnalysisChunks(partials);
      }
      setProgress(75); await tick();

      // ── Step 4: Seed editable bin/pallet overrides from the analysis ──────
      const seeded = {};
      Object.keys(a?.binSummary || {}).forEach(k => {
        const phys = BIN_CATALOG[k] ? BIN_CATALOG[k].phys : null;
        seeded[k] = [{
          L: phys ? phys[0] : '', W: phys ? phys[1] : '', H: phys ? phys[2] : '',
          locs: a.binSummary[k].locs || 0, label: 'Size 1',
        }];
      });
      setBinOverrides(seeded);
      setBinEditsDirty(false);
      setBinEditsOpen(true);  // open so the step is visible after generating

      // ── Step 5: Generate rack config ─────────────────────────────────────
      setProgressMsg('Generating rack configuration…'); await tick();
      const rc = generateRackConfig(a, params);
      setAnalysis(a); setRackConfig(rc);
      setDesign(null); setConfigConfirmed(false);
      setProgress(82); await tick();

      // ── Step 5: Size the warehouse ────────────────────────────────────────
      setProgressMsg('Calculating warehouse dimensions…'); await tick();
      const d = calcWarehouseSize(a, params);
      setDesign(d);
      setProgress(90); await tick();

      // ── Step 6: Mode-specific calcs ───────────────────────────────────────
      if (storageMode === 'fr') {
        setProgressMsg('Calculating Forward Pick + Reserve…'); await tick();
        const res=calcForwardReserve(a,forwardRacks,reserveRacks,forwardDays,params,seeded);
        if(res){
          setFrResult(res); setFrRackConfig(res.allCfgs); setFrOverflow(res.overflow);
          const cza={};
          res.allCfgs.forEach(cfg=>{cza[cfg.zone]=(cza[cfg.zone]||0)+(cfg.area||0);});
          if(!Object.keys(cza).length) cza.golden=50;
          try{setFrDesign(calcWarehouseSize(a,params,null,cza));}catch(e){}
        }
      }
      if (storageMode === 'user') {
        setProgressMsg('Calculating user-defined rack config…'); await tick();
        const res=calcUserRackConfigFromSystemBins(a,userRacks,params,seeded);
        if(res){
          setUserRackConfig(res.uCfgs.length>0?res.uCfgs:null);
          setUserOverflowBins(res.overflowBins||[]);
          setUserResult({stored:[],overflow:res.overflowBins||[],
            totLocs:res.uCfgs.reduce((s,x)=>s+x.locs,0),
            totArea:res.uCfgs.reduce((s,x)=>s+(x.area||0),0),
            totStock:0,totCap:0,overallUtil:0,binUtil:{},rackResults:[]});
          const ca2={},cza2={};
          res.uCfgs.forEach(x=>{ca2[x.rack]=(ca2[x.rack]||0)+(x.area||0);cza2[x.zone]=(cza2[x.zone]||0)+(x.area||0);});
          if(!Object.keys(ca2).length) ca2.shelving=50;
          setUserDesign(calcWarehouseSize(a,params,ca2,cza2));
        } else {
          setUserResult({stored:[],overflow:[],totLocs:0,totArea:0,totStock:0,totCap:0,overallUtil:0,binUtil:{},rackResults:[]});
        }
      }

      setProgress(100);
      setProgressMsg(`✓ ${mData.length.toLocaleString()} SKUs processed`);
    } catch(e) { setError(e.message); console.error(e); }
    setLoading(false);
    setTimeout(() => { setProgress(0); setProgressMsg(''); }, 2000);
  };

  // ── LAYOUT-DERIVED AREA SUMMARY ─────────────────────────────────────────
  // The headline figures must match the drawn plan, not the pre-layout estimate:
  // buildFloorPlanLayout resolves the real envelope (actualWW x actualWL) after
  // rack geometry, band heights and aisle clearances are applied.
  const summaryFromLayout = (dsn, rcfg) => {
    if (!dsn) return null;
    const cfgs = rcfg || [];
    const rackArea = cfgs.reduce((s,cf)=>s+(parseFloat(cf.area)||0), 0);
    let wWv = parseFloat(dsn.wW)||0, wLv = parseFloat(dsn.wL)||0, fromLayout = false;
    try {
      const L = buildFloorPlanLayout(dsn, params, cfgs, analysis, false);
      if (L && L.actualWW > 0 && L.actualWL > 0) {
        wWv = L.actualWW; wLv = L.actualWL; fromLayout = true;
      }
    } catch (err) { /* fall back to the design estimate */ }
    return {
      wW: Math.round(wWv*10)/10,
      wL: Math.round(wLv*10)/10,
      gross: Math.round(wWv*wLv),
      rackArea: Math.round((rackArea || parseFloat(dsn.netRackArea) || 0)*10)/10,
      fromLayout,
    };
  };
  // params is a fresh object each render, so key the memo on its serialised value
  // instead of its identity — otherwise the heavy layout build runs every render.
  const paramsKey = JSON.stringify(params);
  const udSummary  = useMemo(()=>summaryFromLayout(userDesign, userRackConfig),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [userDesign, userRackConfig, analysis, paramsKey]);
  const sysSummary = useMemo(()=>summaryFromLayout(design, rackConfig),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [design, rackConfig, analysis, paramsKey]);

  // ── Measuring tool handlers ─────────────────────────────────────────────
  // Two clicks make one measurement; the first click is held in measurePts.
  // Ortho lock: if the second point is within ORTHO_TOL_DEG of horizontal or
  // vertical, force it exactly onto that axis. Clicking a pixel-perfect 0/90
  // line by hand is not realistic, so this is on by default.
  const ORTHO_TOL_DEG = 8;
  const applyOrtho = (a, b) => {
    const dx = b.x - a.x, dy = b.y - a.y;
    if (!dx && !dy) return b;
    const ang   = Math.abs(Math.atan2(dy, dx) * 180 / Math.PI); // 0..180
    const fromH = Math.min(ang, 180 - ang);   // angular distance to horizontal
    const fromV = Math.abs(ang - 90);         // angular distance to vertical
    if (fromH <= ORTHO_TOL_DEG && fromH <= fromV)
      return { ...b, y:a.y, ortho:'H' };      // lock horizontal
    if (fromV <= ORTHO_TOL_DEG)
      return { ...b, x:a.x, ortho:'V' };      // lock vertical
    return b;
  };

  const onMeasurePoint = (pt) => {
    setMeasurePts(prev => {
      if (!prev.length) return [pt];
      const a = prev[0];
      const b = snapOn ? applyOrtho(a, pt) : pt;
      setMeasurements(ms => [...ms, [a, b]]);
      return [];
    });
  };
  const clearMeasurements = () => { setMeasurements([]); setMeasurePts([]); };
  const undoMeasurement = () => {
    if (measurePts.length) { setMeasurePts([]); return; }
    setMeasurements(ms => ms.slice(0, -1));
  };
  const toggleMeasure = () => {
    setMeasureOn(v => { if (v) setMeasurePts([]); return !v; });
  };

  // ── Bin / Pallet editor handlers ────────────────────────────────────────
  const updateBinField = (binKey, idx, field, val) => {
    setBinOverrides(prev => {
      const next = { ...(prev || {}) };
      const arr  = (next[binKey] || []).slice();
      arr[idx] = { ...(arr[idx] || {}), [field]: val };
      next[binKey] = arr;
      return next;
    });
    setBinEditsDirty(true);
  };

  // Add another size for this bin band, splitting the remaining quantity
  const addBinVariant = (binKey) => {
    setBinOverrides(prev => {
      const next = { ...(prev || {}) };
      const arr  = (next[binKey] || []).slice();
      if (arr.length >= MAX_BIN_VARIANTS) return prev;
      const src  = arr[arr.length - 1] || {};
      const half = Math.max(1, Math.floor((parseFloat(src.locs) || 2) / 2));
      if (arr.length) arr[arr.length - 1] = { ...src, locs: half };
      arr.push({
        L: src.L || '', W: src.W || '', H: src.H || '',
        locs: half, label: 'Size ' + (arr.length + 1),
      });
      next[binKey] = arr;
      return next;
    });
    setBinEditsDirty(true);
  };

  const removeBinVariant = (binKey, idx) => {
    setBinOverrides(prev => {
      const next = { ...(prev || {}) };
      const arr  = (next[binKey] || []).slice();
      if (arr.length <= 1) return prev;
      arr.splice(idx, 1);
      next[binKey] = arr.map((v, i) => ({ ...v, label: v.label || ('Size ' + (i + 1)) }));
      return next;
    });
    setBinEditsDirty(true);
  };

  // Re-run rack generation using the edited bin sizes / counts
  const applyBinEdits = () => {
    if (!analysis) return;
    // Guard: at least one container must have a non-zero quantity
    let totalWanted = 0;
    Object.keys(analysis?.binSummary||{}).forEach(k=>{
      binVariantsFor(k, binOverrides, analysis.binSummary[k].locs||0)
        .forEach(v=>{ totalWanted += v.locs||0; });
    });
    if (totalWanted <= 0) {
      setBinEditsError('Every container quantity is 0 — set a quantity for at least one bin or pallet size.');
      return false;
    }
    setBinEditsError('');
    const rc = generateRackConfig(analysis, params, binOverrides);
    setRackConfig(rc);
    setDesign(syncDesignLevels(calcWarehouseSize(analysis, params, rackAreasFromConfig(rc)), rc));
    setConfigConfirmed(false);
    setBinEditsDirty(false);
    return true;
  };

  // Restore every bin back to the system-generated values
  const resetBinEdits = () => {
    if (!analysis) return;
    const seeded = {};
    Object.keys(analysis?.binSummary || {}).forEach(k => {
      const phys = BIN_CATALOG[k] ? BIN_CATALOG[k].phys : null;
      seeded[k] = [{
        L: phys ? phys[0] : '', W: phys ? phys[1] : '', H: phys ? phys[2] : '',
        locs: analysis.binSummary[k].locs || 0, label: 'Size 1',
      }];
    });
    setBinOverrides(seeded);
    setBinEditsError('');
    const rc = generateRackConfig(analysis, params);
    setRackConfig(rc);
    setDesign(syncDesignLevels(calcWarehouseSize(analysis, params, rackAreasFromConfig(rc)), rc));
    setConfigConfirmed(false);
    setBinEditsDirty(false);
  };

  const confirmConfig = () => {
    if (!analysis || !rackConfig) return;
    const customAreas = rackAreasFromConfig(rackConfig);
    const d = syncDesignLevels(calcWarehouseSize(analysis, params, customAreas), rackConfig);
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

  // Copy system config to user-defined inputs
  const copyFromSystem = () => {
    if (!rackConfig || !analysis) return;
    // Bins come from system analysis — only copy rack sizes
    const uniqueRacks = rackConfig.filter((cfg,i,a)=>a.findIndex(x=>x.rack===cfg.rack)===i);
    if (uniqueRacks.length > 0) {
      setUserRacks(uniqueRacks.map((cfg,i)=>({
        id:i+1, name:cfg.rackName||cfg.rack,
        rackType:cfg.rack,
        bayW:String(cfg.bayW),
        bayD:String(cfg.bayD),
        bayH:String(cfg.tierHeight||cfg.shelfH||2200),
        levels:String(cfg.levels||1),
      })));
    }
    setStorageMode('user');
    setUserResult(null);
  };

  // Download 2D floor plan as SVG or high-res PNG
  const downloadPlan2D = (fmt='svg', ref=null, scaleArg=2) => {
    const host = (ref && ref.current) || plan2DRef.current
      || userPlanRefObj.current || document.querySelector('#fs-plan-svg');
    const svgEl = host
      ? (host.tagName === 'svg' ? host : host.querySelector('svg'))
      : null;
    if (!svgEl) { alert('Plan not ready yet.'); return; }
    // Inject explicit font/style so exported file is self-contained
    const svgStr = new XMLSerializer().serializeToString(svgEl);

    if (fmt === 'svg') {
      const blob = new Blob([svgStr], {type:'image/svg+xml;charset=utf-8'});
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url;
      a.download = `Warehouse_Plan_${design?.wW||''}x${design?.wL||''}m.svg`;
      document.body.appendChild(a); a.click();
      setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); }, 300);
    } else {
      // PNG at 2× resolution
      const vb   = svgEl.viewBox?.baseVal;
      const svgW = vb?.width  || 960;
      const svgH = vb?.height || 720;
      const scale = scaleArg||2;
      const canvas = document.createElement('canvas');
      canvas.width  = svgW * scale;
      canvas.height = svgH * scale;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(scale, scale);
      const blob = new Blob([svgStr], {type:'image/svg+xml;charset=utf-8'});
      const url  = URL.createObjectURL(blob);
      const img  = new Image();
      img.onload = () => {
        ctx.drawImage(img, 0, 0);
        URL.revokeObjectURL(url);
        canvas.toBlob(pngBlob => {
          const pngUrl = URL.createObjectURL(pngBlob);
          const a = document.createElement('a');
          a.href = pngUrl;
          a.download = `Warehouse_Plan_${design?.wW||''}x${design?.wL||''}m_${scale}x.png`;
          document.body.appendChild(a); a.click();
          setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(pngUrl); }, 300);
        }, 'image/png');
      };
      img.onerror = () => URL.revokeObjectURL(url);
      img.src = url;
    }
  };

  // Run user-defined storage calculation
  const runUserCalc = () => {
    if (!masterText.trim()) return;
    setUserLoading(true);
    setTimeout(()=>{
      try {
        // Step 1: get or run analysis
        const mData=parseTable(masterText),oData=parseTable(orderText),iData=parseTable(invText);
        const curA=analysis||runAnalysis(mData,oData,iData,params,preferredBins);
        if(!analysis) setAnalysis(curA);

        // Step 2: SKU storage calc (safe — may return null if no valid bins)
        // Bins from system analysis — no userBins
        const r=null;
        // Always set a non-null userResult so the panel renders
        setUserResult(r||{stored:[],overflow:[],totLocs:0,totArea:0,
          totStock:0,totCap:0,overallUtil:0,binUtil:{},rackResults:[]});

        // Step 3: build userRackConfig
        const ORI=[[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];
        const RZM={shelving:'golden',liveStorage:'golden',selective:'reserve',
          doubleDeep:'reserve',driveIn:'bulk',cantilever:'long'};
        const uCfgs=[];

        if(vBins.length>0&&vRacks.length>0){
          vBins.forEach((b,bi)=>{
            const bL=parseFloat(b.L),bW=parseFloat(b.W),bH=parseFloat(b.H);
            vRacks.forEach((rk,ri)=>{
              const rW=parseFloat(rk.bayW),rD=parseFloat(rk.bayD);
              const rH=parseFloat(rk.bayH)||2200,tiers=parseInt(rk.levels)||1,clr=50;
              const dim=[bL,bW,bH];
              let bestLPB=0,bestAW=0,bestAD=0,bestLvl=0,bestO='LW';
              ORI.forEach(([x,y,z])=>{
                const aw=Math.floor(rW/dim[x]),ad=Math.floor(rD/dim[y]);
                const lv=dim[z]>0?Math.floor(rH/(dim[z]+clr)):0;
                const lpb=aw*ad*lv*tiers;
                if(lpb>bestLPB){bestLPB=lpb;bestO=x===0?'LW':'WL';bestAW=aw;bestAD=ad;bestLvl=lv;}
              });
              const bu=r?.binUtil?.[b.id]||{};
              const locs=bu.locs||Math.ceil((r?.totLocs||curA.metrics.totLocs||0)/Math.max(1,vBins.length));
              const bays=bestLPB>0?Math.ceil(locs/bestLPB):0;
              const ah=(parseFloat(params.aisleW)||3.0)/2;
              const area=+(bays*(rW/1000)*((rD/1000)+ah)).toFixed(1);
              const rt=rk.rackType||'shelving';
              uCfgs.push({
                id:`u-${bi}-${ri}`,rack:rt,bin:`USER_${bi}`,
                rackName:rk.name||`Custom Rack ${ri+1}`,
                binName:b.name||`Custom Bin ${bi+1}`,
                binDims:[bL,bW,bH],bayW:rW,bayD:rD,shelfH:rH,tierHeight:rH,clearance:clr,
                orientation:bestO,tiers,acrossW:bestAW,acrossD:bestAD,levels:bestLvl,
                locsPerBay:bestLPB>0?Math.floor(bestLPB/tiers):0,locsPerBayTotal:bestLPB,
                locs,baysNeeded:bays,area,feasible:bestLPB>0,zone:RZM[rt]||'golden',
                o1:{acrossW:Math.floor(rW/bL),acrossD:Math.floor(rD/bW),feasible:Math.floor(rW/bL)>0,
                    levels:bH>0?Math.floor(rH/(bH+clr)):0,
                    locsPerBay:Math.floor(rW/bL)*Math.floor(rD/bW)*(bH>0?Math.floor(rH/(bH+clr)):0)},
                o2:{acrossW:Math.floor(rW/bW),acrossD:Math.floor(rD/bL),feasible:Math.floor(rW/bW)>0,
                    levels:bH>0?Math.floor(rH/(bH+clr)):0,
                    locsPerBay:Math.floor(rW/bW)*Math.floor(rD/bL)*(bH>0?Math.floor(rH/(bH+clr)):0)},
              });
            });
          });
        } else {
          // Fallback to system rackConfig or auto-generate
          const sys=(rackConfig&&rackConfig.length>0)?rackConfig
            :(generateRackConfig(curA,params)||[]);
          sys.forEach(cfg=>uCfgs.push({...cfg}));
        }
        setUserRackConfig(uCfgs.length>0?uCfgs:null);

        // Step 4: ALWAYS compute layout
        const ca={},cza={};
        uCfgs.forEach(cfg=>{ca[cfg.rack]=(ca[cfg.rack]||0)+(cfg.area||0);});cza[cfg.zone]=(cza[cfg.zone]||0)+(cfg.area||0);
        if(!Object.keys(ca).length) ca.shelving=r?.totArea||50;
        setUserDesign(calcWarehouseSize(curA,params,ca,cza));

      } catch(e){ console.error('User calc error:',e.message,e); }
      setUserLoading(false);
    },80);
  };

  // Helpers for user-defined bin/rack editing
  const updateUserRack = (id,field,val) => setUserRacks(prev=>prev.map(r=>r.id===id?{...r,[field]:val}:r));
  const addUserRack    = () => { const id=Date.now(); setUserRacks(prev=>[...prev,{id,name:`Custom Rack ${prev.length+1}`,rackType:'shelving',bayW:'',bayD:'',bayH:'',levels:''}]); };
  const removeUserRack = id => setUserRacks(prev=>prev.filter(r=>r.id!==id));

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

          {/* ── STORAGE DESIGN MODE TOGGLE ──────────────────────────── */}
          <div style={{...S.card,padding:'12px 16px'}}>
            <div style={{fontSize:'11px',fontWeight:'700',color:'#374151',
              textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'8px'}}>
              Storage Design Mode
            </div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px'}}>
              {[['system','⚙ System Defined','Tool selects optimal bins & racks'],
                ['user',  '✏ User Defined',  'Enter your own bin & rack sizes']
               ].map(([val,label,sub])=>(
                <button key={val} onClick={()=>{
                  setStorageMode(val);
                  // Auto-populate user fields from system config when switching to User Defined
                  if(val==='user' && rackConfig && rackConfig.length>0){
                    const hasEmptyRacks=userRacks.every(r=>!parseFloat(r.bayW));
                    if(hasEmptyRacks) copyFromSystem();
                  }
                }}
                  style={{padding:'10px 12px',borderRadius:'9px',textAlign:'left',cursor:'pointer',
                    border:`2px solid ${storageMode===val?'#7c3aed':'#e2e8f0'}`,
                    background:storageMode===val?'#f5f3ff':'#fff'}}>
                  <div style={{fontWeight:'700',fontSize:'12px',
                    color:storageMode===val?'#7c3aed':'#374151'}}>{label}</div>
                  <div style={{fontSize:'10px',color:'#9ca3af',marginTop:'2px'}}>{sub}</div>
                </button>
              ))}
            </div>
            {storageMode==='user' && analysis && rackConfig && (
              <button onClick={copyFromSystem}
                style={{marginTop:'8px',width:'100%',padding:'7px',
                  background:'#eff6ff',border:'1px dashed #93c5fd',borderRadius:'7px',
                  fontSize:'12px',fontWeight:'700',color:'#1d4ed8',cursor:'pointer',fontFamily:'inherit'}}>
                ⬇ Copy bin & rack sizes from System Defined
              </button>
            )}
          </div>

          {/* ── USER DEFINED BIN & RACK INPUTS ──────────────────────── */}
          {storageMode==='user' && (<>
            {/* ════════════════════════════════════════════════════════════
                USER DEFINED — 4-STEP WIZARD
                ════════════════════════════════════════════════════════════ */}

            {/* Step indicator */}
            <div style={{display:'flex',gap:'0',marginBottom:'14px',
              border:'1px solid #e2e8f0',borderRadius:'10px',overflow:'hidden'}}>
              {[['1','Bin Types'],['2','Rack Types'],['3','Rack Sizes'],['4','Layout']].map(([n,label],i)=>{
                const stepN=parseInt(n);
                const done=udStep>stepN, active=udStep===stepN;
                return(
                  <div key={n} style={{flex:1,padding:'8px 4px',textAlign:'center',
                    background:done?'#f0fdf4':active?'#f5f3ff':'#f8fafc',
                    borderRight:i<3?'1px solid #e2e8f0':'none',
                    cursor:done&&stepN<udStep?'pointer':'default'}}
                    onClick={()=>{ if(done||stepN<udStep) setUdStep(stepN); }}>
                    <div style={{fontSize:'16px',fontWeight:'800',
                      color:done?'#059669':active?'#7c3aed':'#d1d5db'}}>
                      {done?'✓':n}
                    </div>
                    <div style={{fontSize:'9px',fontWeight:'700',
                      color:done?'#059669':active?'#7c3aed':'#9ca3af',
                      textTransform:'uppercase',letterSpacing:'0.04em'}}>
                      {label}
                    </div>
                  </div>);
              })}
            </div>

            {/* ── STEP 1: Calculate Bin Types ─────────────────────────── */}
            {udStep===1&&(
              <div style={S.card}>
                <div style={S.cardTitle}>📦 Step 1 — Calculate Bin Types</div>
                <div style={{fontSize:'11px',color:'#6b7280',marginBottom:'12px'}}>
                  The tool will analyse your SKU dimensions and determine which bin
                  sizes are needed. Paste your SKU data in the Master SKU field below
                  first, then click Calculate.
                </div>
                {analysis?.binSummary&&Object.keys(analysis.binSummary).length>0?(
                  <div>
                    <div style={{fontSize:'11px',fontWeight:'700',color:'#059669',marginBottom:'8px'}}>
                      ✓ Bin types calculated from {(analysis.metrics?.totSKUs||0).toLocaleString()} SKUs
                    </div>

                    {/* ── EDITABLE BIN / PALLET SIZES + QUANTITIES ──────────── */}
                    {binOverrides && (
                      <div style={{border:'2px solid #c4b5fd',borderRadius:'10px',
                        background:'linear-gradient(180deg,#faf8ff,#fff)',
                        padding:'10px',marginBottom:'12px'}}>
                        <div style={{fontWeight:'800',fontSize:'12px',color:'#5b21b6',marginBottom:'2px'}}>
                          ✏️ Edit sizes &amp; quantities
                          <span style={{fontSize:'9px',fontWeight:'700',color:'#7c3aed',
                            background:'#f5f3ff',border:'1px solid #c4b5fd',
                            borderRadius:'99px',padding:'2px 7px',marginLeft:'7px'}}>OPTIONAL</span>
                        </div>
                        <div style={{fontSize:'10px',color:'#6b7280',marginBottom:'8px'}}>
                          Change any dimension or quantity. Use <strong>+ size</strong> for up to
                          3 different sizes per container type — each becomes its own rack
                          configuration and its own block in the layout.
                        </div>
                        <div style={{overflowX:'auto'}}>
                          <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                            <thead>
                              <tr style={{background:'#f8fafc'}}>
                                {['Bin','Size','L (mm)','W (mm)','H (mm)','Qty','Vol',''].map((h,hi)=>(
                                  <th key={hi} style={{padding:'5px 7px',textAlign:'left',
                                    fontSize:'9px',fontWeight:'700',color:'#6b7280',
                                    textTransform:'uppercase',borderBottom:'1px solid #e2e8f0',
                                    whiteSpace:'nowrap'}}>{h}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {Object.entries(analysis?.binSummary||{})
                                .sort((a,b)=>['XS','S','M','L','XL','LONG'].indexOf(a[0])
                                            -['XS','S','M','L','XL','LONG'].indexOf(b[0]))
                                .map(([band,info])=>{
                                  const variants=binOverrides[band]||[];
                                  const base=BIN_CATALOG[band]?BIN_CATALOG[band].phys:null;
                                  const dimLocked=!base;
                                  const COLORS={XS:['#f1f5f9','#64748b'],S:['#eff6ff','#1d4ed8'],
                                    M:['#f5f3ff','#7c3aed'],L:['#f0fdf4','#166534'],
                                    XL:['#fef9c3','#854d0e'],LONG:['#fdf4ff','#9333ea']};
                                  const [bg,col]=COLORS[band]||['#f8fafc','#374151'];
                                  const inS={border:'1px solid #e2e8f0',borderRadius:'5px',
                                    padding:'4px 6px',fontSize:'11px',width:'66px',
                                    boxSizing:'border-box',fontFamily:'inherit',outline:'none'};
                                  const bandQty=variants.reduce((s2,v)=>s2+(parseFloat(v.locs)||0),0);
                                  return variants.map((ov,vi)=>{
                                    const bL=parseFloat(ov.L)||0;
                                    const bW=parseFloat(ov.W)||0;
                                    const bH=parseFloat(ov.H)||0;
                                    const vol=(bL*bW*bH)/1e9;
                                    const excluded=(parseFloat(ov.locs)||0)<=0;
                                    return (
                                      <tr key={band+'-'+vi} style={{borderBottom:
                                        vi===variants.length-1?'1px solid #e2e8f0':'1px dashed #f1f5f9',
                                        background:excluded?'#fafafa':'transparent',
                                        opacity:excluded?0.55:1}}>
                                        <td style={{padding:'5px 7px',verticalAlign:'top'}}>
                                          {vi===0&&(<>
                                            <span style={{background:bg,color:col,fontWeight:'800',
                                              fontSize:'11px',borderRadius:'5px',padding:'2px 7px'}}>{band}</span>
                                            <div style={{fontSize:'9px',color:'#9ca3af',marginTop:'2px'}}>
                                              {info.name}
                                            </div>
                                            {bandQty<=0&&(
                                              <div style={{fontSize:'9px',fontWeight:'700',color:'#be185d',
                                                marginTop:'2px'}}>EXCLUDED</div>
                                            )}
                                          </>)}
                                        </td>
                                        <td style={{padding:'5px 7px'}}>
                                          <input type="text" value={ov.label||''}
                                            placeholder={'Size '+(vi+1)}
                                            onChange={e=>updateBinField(band,vi,'label',e.target.value)}
                                            style={{...inS,width:'74px',fontWeight:'700',color:col}}/>
                                        </td>
                                        <td style={{padding:'5px 7px'}}>
                                          <input type="number" min="50" step="10" disabled={dimLocked}
                                            value={ov.L||''}
                                            onChange={e=>updateBinField(band,vi,'L',e.target.value)}
                                            style={{...inS,background:dimLocked?'#f8fafc':'#fff'}}/>
                                        </td>
                                        <td style={{padding:'5px 7px'}}>
                                          <input type="number" min="50" step="10" disabled={dimLocked}
                                            value={ov.W||''}
                                            onChange={e=>updateBinField(band,vi,'W',e.target.value)}
                                            style={{...inS,background:dimLocked?'#f8fafc':'#fff'}}/>
                                        </td>
                                        <td style={{padding:'5px 7px'}}>
                                          <input type="number" min="50" step="10" disabled={dimLocked}
                                            value={ov.H||''}
                                            onChange={e=>updateBinField(band,vi,'H',e.target.value)}
                                            style={{...inS,background:dimLocked?'#f8fafc':'#fff'}}/>
                                        </td>
                                        <td style={{padding:'5px 7px'}}>
                                          <input type="number" min="0" step="1"
                                            value={ov.locs===0?'0':(ov.locs||'')}
                                            onChange={e=>updateBinField(band,vi,'locs',e.target.value)}
                                            title="Set 0 to exclude this size from the design"
                                            style={{...inS,width:'78px',
                                              borderColor:excluded?'#fecdd3':'#e2e8f0',
                                              background:excluded?'#fff1f2':'#fff',
                                              color:excluded?'#be185d':'inherit',
                                              fontWeight:excluded?'700':'400'}}/>
                                        </td>
                                        <td style={{padding:'5px 7px',color:'#6b7280',whiteSpace:'nowrap',fontSize:'10px'}}>
                                          {vol>0?vol.toFixed(3):'-'}
                                        </td>
                                        <td style={{padding:'5px 7px',whiteSpace:'nowrap'}}>
                                          {variants.length>1&&(
                                            <button onClick={()=>removeBinVariant(band,vi)} title="Remove this size"
                                              style={{border:'1px solid #fecdd3',background:'#fff1f2',
                                                color:'#be185d',borderRadius:'5px',padding:'2px 7px',
                                                fontSize:'10px',fontWeight:'700',cursor:'pointer',
                                                fontFamily:'inherit'}}>×</button>
                                          )}
                                          {vi===0&&variants.length<MAX_BIN_VARIANTS&&!dimLocked&&(
                                            <button onClick={()=>addBinVariant(band)} title="Add another size"
                                              style={{marginLeft:'3px',border:'1px dashed #c4b5fd',
                                                background:'#f5f3ff',color:'#7c3aed',borderRadius:'5px',
                                                padding:'2px 7px',fontSize:'10px',fontWeight:'700',
                                                cursor:'pointer',fontFamily:'inherit'}}>+ size</button>
                                          )}
                                        </td>
                                      </tr>
                                    );
                                  });
                                })}
                            </tbody>
                          </table>
                        </div>
                        <div style={{fontSize:'10px',color:'#6b7280',marginTop:'6px'}}>
                          Set a quantity to <strong>0</strong> to exclude that container size
                          from the rack design entirely.
                        </div>
                        {binEditsError&&(
                          <div style={{fontSize:'10px',fontWeight:'700',color:'#be185d',
                            background:'#fff1f2',border:'1px solid #fecdd3',borderRadius:'6px',
                            padding:'5px 8px',marginTop:'8px'}}>
                            ⚠ {binEditsError}
                          </div>
                        )}
                        {binEditsDirty&&(
                          <div style={{fontSize:'10px',fontWeight:'700',color:'#d97706',
                            background:'#fffbeb',border:'1px solid #fde68a',borderRadius:'6px',
                            padding:'5px 8px',marginTop:'8px'}}>
                            ⚠ You have unapplied edits — they will be used when you proceed.
                          </div>
                        )}
                        <button onClick={resetBinEdits}
                          style={{marginTop:'8px',width:'100%',padding:'7px',borderRadius:'7px',
                            border:'1px solid #e2e8f0',background:'#fff',color:'#6b7280',
                            fontWeight:'700',fontSize:'11px',fontFamily:'inherit',cursor:'pointer'}}>
                          ↺ Reset to system-generated sizes
                        </button>
                      </div>
                    )}

                    <button onClick={()=>{ if(applyBinEdits()!==false) setUdStep(2); }}
                      style={{width:'100%',padding:'10px',borderRadius:'9px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'13px',fontWeight:'700',border:'none',
                        background:'linear-gradient(135deg,#7c3aed,#6d28d9)',color:'#fff'}}>
                      → Proceed to Select Rack Types
                    </button>
                  </div>
                ):(
                  <div style={{fontSize:'11px',color:'#9ca3af',padding:'10px',
                    background:'#f8fafc',borderRadius:'6px',textAlign:'center',marginBottom:'10px'}}>
                    Click "Generate Warehouse Design" below after pasting SKU data
                  </div>
                )}
              </div>
            )}

            {/* ── STEP 2: Select Rack Types ───────────────────────────── */}
            {udStep===2&&(
              <div style={S.card}>
                <div style={S.cardTitle}>🏗 Step 2 — Select Rack Types</div>
                <div style={{fontSize:'11px',color:'#6b7280',marginBottom:'12px'}}>
                  Choose the rack types to include in your warehouse layout.
                  The tool will automatically assign the right bin sizes to each rack.
                </div>
                <div style={{display:'flex',flexDirection:'column',gap:'8px',marginBottom:'14px'}}>
                  {[
                    {key:'shelving',    icon:'📦', label:'Long Span Shelving',   desc:'XS/S/M bins, pick face, carton storage'},
                    {key:'liveStorage', icon:'🔄', label:'Carton Live / Flow',    desc:'VF/F items, FIFO gravity flow'},
                    {key:'selective',   icon:'🏗', label:'Selective Pallet Rack', desc:'Full pallets, all velocities'},
                    {key:'doubleDeep',  icon:'🔩', label:'Double-Deep Rack',      desc:'High-density pallet storage, reach truck'},
                    {key:'driveIn',     icon:'🚗', label:'Drive-In Rack',         desc:'High volume, same-SKU lanes'},
                    {key:'cantilever',  icon:'🪵', label:'Cantilever Rack',       desc:'Long bars, pipes, timber'},
                    {key:'ground',      icon:'🏔', label:'Ground Storage',         desc:'Odd-shaped, oversized, no rack needed'},
                  ].map(({key,icon,label,desc})=>{
                    const sel=selectedRackTypes.has(key);
                    return(
                      <div key={key}
                        onClick={()=>setSelectedRackTypes(prev=>{
                          const n=new Set(prev);
                          if(n.has(key)) n.delete(key); else n.add(key);
                          return n;
                        })}
                        style={{display:'flex',alignItems:'center',gap:'12px',padding:'10px 12px',
                          borderRadius:'9px',cursor:'pointer',border:`2px solid ${sel?'#7c3aed':'#e2e8f0'}`,
                          background:sel?'#f5f3ff':'#fff',transition:'all 0.15s'}}>
                        <div style={{fontSize:'20px'}}>{icon}</div>
                        <div style={{flex:1}}>
                          <div style={{fontWeight:'700',fontSize:'12px',
                            color:sel?'#6d28d9':'#0f172a'}}>{label}</div>
                          <div style={{fontSize:'10px',color:'#6b7280'}}>{desc}</div>
                        </div>
                        <div style={{width:'20px',height:'20px',borderRadius:'50%',
                          background:sel?'#7c3aed':'#e2e8f0',display:'flex',
                          alignItems:'center',justifyContent:'center',
                          fontSize:'12px',color:'#fff',fontWeight:'800',flexShrink:0}}>
                          {sel?'✓':''}
                        </div>
                      </div>
                    );
                  })}
                </div>
                <div style={{display:'flex',gap:'8px'}}>
                  <button onClick={()=>setUdStep(1)}
                    style={{flex:1,padding:'9px',borderRadius:'9px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      border:'1px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                    ← Back
                  </button>
                  <button onClick={()=>{
                      if(selectedRackTypes.size===0){alert('Select at least one rack type.');return;}
                      // Initialise udRackDefs with defaults for selected types
                      const defaults={
                        shelving:   {bayW:'900', bayD:'600', bayH:'2200',levels:'4',aisle:'1200'},
                        liveStorage:{bayW:'900', bayD:'1500',bayH:'2200',levels:'4',aisle:'1200'},
                        selective:  {bayW:'2700',bayD:'1100',bayH:'6000',levels:'4',aisle:'3000'},
                        doubleDeep: {bayW:'2700',bayD:'2200',bayH:'6000',levels:'4',aisle:'3500'},
                        driveIn:    {bayW:'2700',bayD:'6600',bayH:'6000',levels:'4',aisle:'3500'},
                        cantilever: {bayW:'1500',bayD:'2500',bayH:'3000',levels:'6',aisle:'3000'},
                        ground:     {bayW:'1500',bayD:'1200',bayH:'',   levels:'2',aisle:'4000'},
                      };
                      setUdRackDefs(prev=>{
                        const next={...prev};
                        selectedRackTypes.forEach(k=>{
                          if(!next[k]) next[k]={...defaults[k]};
                        });
                        return next;
                      });
                      setUdStep(3);
                    }}
                    style={{flex:2,padding:'9px',borderRadius:'9px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'13px',fontWeight:'700',
                      border:'none',background:'linear-gradient(135deg,#7c3aed,#6d28d9)',
                      color:'#fff',opacity:selectedRackTypes.size>0?1:0.5}}>
                    → Define Rack Sizes ({selectedRackTypes.size} selected)
                  </button>
                </div>
              </div>
            )}

            {/* ── STEP 3: Define Rack Sizes ───────────────────────────── */}
            {udStep===3&&(
              <div style={S.card}>
                <div style={S.cardTitle}>📐 Step 3 — Define Rack Sizes</div>
                <div style={{fontSize:'11px',color:'#6b7280',marginBottom:'12px'}}>
                  Enter bay dimensions for each rack type. Shelf Levels = number of
                  shelves (shelving) or stack layers (ground).
                </div>
                {[...selectedRackTypes].map(key=>{
                  const LABELS={shelving:'📦 Shelving',liveStorage:'🔄 Flow Rack',
                    selective:'🏗 Selective Pallet',doubleDeep:'🔩 Double-Deep',
                    driveIn:'🚗 Drive-In',cantilever:'🪵 Cantilever',ground:'🏔 Ground'};
                  const isPallet=['selective','doubleDeep','driveIn'].includes(key);
                  const isShelving=['shelving','liveStorage','cantilever'].includes(key);
                  const d=udRackDefs[key]||{};
                  const upd=(f,v)=>setUdRackDefs(prev=>({...prev,[key]:{...prev[key],[f]:v}}));
                  return(
                    <div key={key} style={{border:'1px solid #e2e8f0',borderRadius:'9px',
                      overflow:'hidden',marginBottom:'10px'}}>
                      <div style={{background:'#f8fafc',padding:'7px 12px',
                        fontWeight:'700',fontSize:'12px',color:'#0f172a',
                        borderBottom:'1px solid #e2e8f0'}}>
                        {LABELS[key]||key}
                      </div>
                        {/* Rack-specific dimension hints */}
                        {key==='cantilever'&&(
                          <div style={{background:'#fffbeb',border:'1px solid #fcd34d',
                            borderRadius:'7px',padding:'8px 10px',marginBottom:'8px',
                            fontSize:'10px',color:'#78350f',lineHeight:'1.7'}}>
                            <div style={{fontWeight:'700',marginBottom:'4px'}}>
                              📐 Cantilever Rack Dimensions (Top View):
                            </div>
                            <pre style={{fontFamily:'monospace',fontSize:'9px',
                              color:'#92400e',margin:'0 0 5px',lineHeight:'1.4',
                              background:'#fef3c7',padding:'6px',borderRadius:'5px',
                              overflowX:'auto'}}>
{`AISLE ─────────────────────────── (you walk along here)
  ║══ upright face ════════════║
  ║ →→ item →→→→→→→→→→→→→→ ║
  ║ →→ item →→→→→→→→→→→→→→ ║  ↕ Arm Length
  ║ →→ item →→→→→→→→→→→→→→ ║  ↕ (bayD)
  ║════════════════════════════║
  ←── Frame Width (bayW) ─────→`}
                            </pre>
                            <div style={{color:'#78350f'}}>
                              • <strong>Frame Width (bayW)</strong> = items <strong>side-by-side</strong> as seen from the aisle<br/>
                              • <strong>Arm Length (bayD)</strong> = items extend <strong>away from aisle</strong> — must be ≥ longest item<br/>
                              • <strong>Clear Height</strong> = total rack height → arm levels auto-calculated
                            </div>
                          </div>
                        )}
                        {key==='driveIn'&&(
                          <div style={{background:'#f0fdf4',border:'1px solid #86efac',
                            borderRadius:'7px',padding:'7px 10px',marginBottom:'8px',
                            fontSize:'10px',color:'#14532d',lineHeight:'1.5'}}>
                            <strong>Drive-In rack:</strong> Bay Depth = lane depth (multiple pallets deep). Bay Width = lane width (one pallet wide per lane).
                          </div>
                        )}
                        <div style={{padding:'0 0 2px',
                          display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px'}}>
                          {[
                            [key==='cantilever'?'Frame Width (mm)':'Bay Width (mm)','bayW'],
                            [key==='cantilever'?'Arm Length / Depth (mm)':key==='driveIn'?'Lane Depth (mm)':'Bay Depth (mm)','bayD'],
                            key==='ground'?['Stack Layers','levels']:['Clear Height (mm)','bayH'],
                            key==='ground'?['Stacked H (mm) — optional','bayH']:null,
                            ['Picking Aisle (mm)','aisle'],
                          ].filter(Boolean).map(([label,field])=>(
                            <div key={field}>
                              <div style={{fontSize:'10px',fontWeight:'600',marginBottom:'3px',
                                color:field==='aisle'?'#7c3aed':'#6b7280'}}>
                                {label}{field==='aisle'&&<span style={{fontWeight:'400',color:'#9ca3af',marginLeft:'4px'}}>(between back-to-back pairs)</span>}
                              </div>
                              <input type="number" min="1" value={d[field]||''}
                                onChange={e=>upd(field,e.target.value)}
                                placeholder={field==='bayH'&&key==='ground'?'optional':field==='aisle'?'mm':'mm'}
                                style={{...inp,marginBottom:0,width:'100%',
                                  fontSize:'12px',padding:'5px 8px',
                                  border:field==='aisle'?'1px solid #c4b5fd':'1px solid #e2e8f0',
                                  background:field==='aisle'?'#faf5ff':'#fff'}}/>
                            </div>
                          ))}
                      </div>
                      {/* Auto-level hint */}
                      {(isShelving||isPallet)&&d.bayH&&(
                        <div style={{padding:'5px 12px 8px',fontSize:'10px',
                          color:'#7c3aed',background:'#faf5ff',
                          borderTop:'1px solid #ede9fe'}}>
                          {isShelving&&`⚙ Levels auto-calculated: clear ${d.bayH}mm ÷ (bin H + 25mm shelf + 30mm headroom)`}
                          {isPallet&&`⚙ Levels auto-calculated: level 1 on floor (pallet 144mm + load H + 100mm MHE), upper levels +100mm beam`}
                        </div>
                      )}
                      </div>
                  );
                })}
                <div style={{display:'flex',gap:'8px'}}>
                  <button onClick={()=>setUdStep(2)}
                    style={{flex:1,padding:'9px',borderRadius:'9px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      border:'1px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                    ← Back
                  </button>
                  <button onClick={()=>{
                      // Convert udRackDefs → userRacks array format
                      const newRacks=[...selectedRackTypes].map((key,i)=>({
                        id:i+1, name:({shelving:'Shelving',liveStorage:'Flow Rack',
                          selective:'Selective Pallet',doubleDeep:'Double-Deep',
                          driveIn:'Drive-In',cantilever:'Cantilever',ground:'Ground Storage'}[key]||key),
                        rackType:key,
                        bayW:String(udRackDefs[key]?.bayW||''),
                        bayD:String(udRackDefs[key]?.bayD||''),
                        bayH:String(udRackDefs[key]?.bayH||''),
                        levels:String(udRackDefs[key]?.levels||'1'),
                        aisle:String(udRackDefs[key]?.aisle||'3000'),
                      }));
                      setUserRacks(newRacks);
                      setUdStep(4);
                      setUdViewMode('2d'); // start with safe 2D view
                    }}
                    style={{flex:2,padding:'9px',borderRadius:'9px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'13px',fontWeight:'700',
                      border:'none',background:'linear-gradient(135deg,#059669,#047857)',
                      color:'#fff'}}>
                    🏭 Create Layout
                  </button>
                </div>
              </div>
            )}

            {/* ── STEP 4: Layout created — show summary + edit options ─── */}
            {udStep===4&&(
              <div style={S.card}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'10px'}}>
                  <div style={S.cardTitle}>✓ Layout Created</div>
                  <button onClick={()=>setUdStep(3)}
                    style={{fontSize:'11px',fontWeight:'700',color:'#7c3aed',
                      background:'#f5f3ff',border:'1px solid #ede9fe',
                      borderRadius:'7px',padding:'4px 10px',cursor:'pointer'}}>
                    ✏ Edit Rack Sizes
                  </button>
                </div>
                {/* Bin type summary (read-only) */}
                <div style={{fontSize:'11px',fontWeight:'700',color:'#374151',marginBottom:'6px'}}>
                  Bin Types (from system analysis):
                </div>
                <div style={{display:'flex',flexWrap:'wrap',gap:'6px',marginBottom:'12px'}}>
                  {Object.entries(analysis?.binSummary||{})
                    .sort((a,b)=>['XS','S','M','L','XL','LONG'].indexOf(a[0])-['XS','S','M','L','XL','LONG'].indexOf(b[0]))
                    .map(([band,info])=>{
                      const overflow=userOverflowBins.find(o=>o.binKey===band)||
                        !new Set((userRackConfig||[]).map(c=>c.bin)).has(band);
                      const fitted=new Set((userRackConfig||[]).map(c=>c.bin)).has(band);
                      return(
                        <div key={band} style={{
                          background:fitted?'#f0fdf4':'#fff1f2',
                          border:`1px solid ${fitted?'#86efac':'#fecaca'}`,
                          borderRadius:'7px',padding:'4px 8px',fontSize:'10px'}}>
                          <span style={{fontWeight:'800',color:fitted?'#166534':'#be185d'}}>
                            {fitted?'✓':'✗'} {band}
                          </span>
                          <span style={{color:'#6b7280',marginLeft:'4px'}}>{(info.locs||0).toLocaleString()} locs</span>
                        </div>
                      );
                    })}
                </div>
                {/* Rack summary */}
                <div style={{fontSize:'11px',fontWeight:'700',color:'#374151',marginBottom:'6px'}}>
                  Racks ({selectedRackTypes.size} types selected):
                </div>
                {userRacks.map((r,i)=>(
                  <div key={i} style={{display:'flex',justifyContent:'space-between',
                    padding:'5px 8px',borderRadius:'6px',marginBottom:'4px',
                    background:'#f8fafc',border:'1px solid #e2e8f0',fontSize:'11px'}}>
                    <span style={{fontWeight:'700',color:'#0f172a'}}>{r.name}</span>
                    <span style={{color:'#6b7280'}}>
                      {r.bayW}×{r.bayD}{r.bayH?`×${r.bayH}`:''}mm · {r.levels} levels
                    </span>
                  </div>
                ))}
                <div style={{display:'flex',gap:'6px',marginTop:'10px'}}>
                  <button onClick={()=>{setUdStep(1);setSelectedRackTypes(new Set());setUdRackDefs({});setUserRacks([{id:1,name:'Custom Rack 1',rackType:'shelving',bayW:'',bayD:'',bayH:'',levels:''}]);}}
                    style={{flex:1,padding:'7px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                      border:'1px solid #e2e8f0',background:'#fff',color:'#9ca3af'}}>
                    ↺ Start Over
                  </button>
                  <button onClick={()=>setUdStep(2)}
                    style={{flex:1,padding:'7px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                      border:'1px solid #ede9fe',background:'#f5f3ff',color:'#7c3aed'}}>
                    Change Racks
                  </button>
                </div>
              </div>
            )}

            {/* Copy from System button (always visible in user mode) */}
            {analysis&&rackConfig&&udStep<=2&&(
              <button onClick={()=>{copyFromSystem();setUdStep(3);}}
                style={{width:'100%',padding:'9px',borderRadius:'9px',cursor:'pointer',
                  fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                  border:'1px solid #c4b5fd',background:'#f5f3ff',color:'#7c3aed',
                  marginBottom:'6px'}}>
                ⚡ Copy Rack Sizes from System Recommendation
              </button>
            )}

          </>)}

          {/* ── STEP 2: MASTER SKU DATA ────────────────────────────────── */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'12px'}}>
              {stepCircle(2, !!masterText.trim())}
              <div style={S.cardTitle}>② Master SKU Data</div>
            </div>
            <div style={{display:'flex',gap:'4px',flexWrap:'wrap',marginBottom:'8px'}}>
              {['SKU Code','Length (mm)','Width (mm)','Height (mm)','Weight (kg)'].map((col,i)=>(
                <span key={i} style={{background:'#f1f5f9',border:'1px solid #e2e8f0',
                  borderRadius:'6px',padding:'3px 8px',fontSize:'11px',fontWeight:'600',
                  color:'#475569',display:'flex',alignItems:'center',gap:'4px'}}>
                  <span style={{background:'#be185d',color:'#fff',borderRadius:'50%',
                    width:'14px',height:'14px',display:'inline-flex',alignItems:'center',
                    justifyContent:'center',fontSize:'9px',fontWeight:'800',flexShrink:0}}>{i+1}</span>
                  {col}
                </span>))}
            </div>
            <textarea value={masterText} onChange={e=>setMasterText(e.target.value)}
              placeholder={'Paste SKU master data (Ctrl+V)\n\nExample:\nSKU-001\t300\t200\t150\t2.5\nSKU-002\t650\t80\t80\t1.2'}
              style={{width:'100%',height:'110px',border:'1px solid #e2e8f0',borderRadius:'8px',
                padding:'9px 11px',fontSize:'12px',fontFamily:'monospace',resize:'vertical',
                outline:'none',boxSizing:'border-box',color:'#374151',lineHeight:'1.5'}}/>
            {masterText.trim() && (
              <div style={{fontSize:'11px',color:'#059669',marginTop:'4px',fontWeight:'600'}}>
                ✓ {masterText.trim().split('\n').filter(l=>l.trim()).length} rows detected
              </div>
            )}
          </div>

          {/* ── STEP 3: ORDER / PICK DATA ──────────────────────────────── */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'8px'}}>
              {stepCircle(3, !!orderText.trim())}
              <div>
                <div style={S.cardTitle}>③ Order / Pick Data
                  <span style={{fontSize:'11px',fontWeight:'400',color:'#059669',marginLeft:'6px'}}>(Optional)</span>
                </div>
                <div style={{fontSize:'11px',color:'#6b7280'}}>Drives velocity band classification (VF/F/M/S/VS/NM)</div>
              </div>
            </div>
            <div style={{display:'flex',gap:'4px',flexWrap:'wrap',marginBottom:'6px'}}>
              {['Order No','Dispatch Location','SKU Code','Qty','Date'].map((col,i)=>(
                <span key={i} style={{background:'#f1f5f9',border:'1px solid #e2e8f0',
                  borderRadius:'6px',padding:'2px 7px',fontSize:'10px',fontWeight:'600',color:'#475569'}}>
                  {col}
                </span>))}
            </div>
            <textarea value={orderText} onChange={e=>setOrderText(e.target.value)}
              placeholder={'Paste order/pick data — SKU frequency drives zone assignment\n\nWithout this, all SKUs treated as equal velocity'}
              style={{width:'100%',height:'90px',border:'1px solid #e2e8f0',borderRadius:'8px',
                padding:'9px 11px',fontSize:'12px',fontFamily:'monospace',resize:'vertical',
                outline:'none',boxSizing:'border-box',color:'#374151',lineHeight:'1.5'}}/>
          </div>

          {/* ── STEP 4: INVENTORY ──────────────────────────────────────── */}
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'8px'}}>
              {stepCircle(4, !!invText.trim())}
              <div>
                <div style={S.cardTitle}>④ Current Inventory
                  <span style={{fontSize:'11px',fontWeight:'400',color:'#059669',marginLeft:'6px'}}>(Optional)</span>
                </div>
                <div style={{fontSize:'11px',color:'#6b7280'}}>Calculates locations required per SKU</div>
              </div>
            </div>
            <div style={{display:'flex',gap:'4px',flexWrap:'wrap',marginBottom:'6px'}}>
              {['SKU Code','Current Stock Qty'].map((col,i)=>(
                <span key={i} style={{background:'#f1f5f9',border:'1px solid #e2e8f0',
                  borderRadius:'6px',padding:'2px 7px',fontSize:'10px',fontWeight:'600',color:'#475569'}}>
                  {col}
                </span>))}
            </div>
            <textarea value={invText} onChange={e=>setInvText(e.target.value)}
              placeholder={'Paste inventory\n\nSKU-001\t2500\nSKU-002\t180'}
              style={{width:'100%',height:'80px',border:'1px solid #e2e8f0',borderRadius:'8px',
                padding:'9px 11px',fontSize:'12px',fontFamily:'monospace',resize:'vertical',
                outline:'none',boxSizing:'border-box',color:'#374151',lineHeight:'1.5'}}/>
          </div>

          {error && <div style={{...S.error,marginBottom:'10px'}}>⚠ {error}</div>}

          {/* ── GENERATE BUTTON ────────────────────────────────────────── */}
          <button onClick={runAll} disabled={loading||!masterText.trim()}
            style={{width:'100%',padding:'13px',
              background:masterText.trim()&&!loading?'linear-gradient(135deg,#7c3aed,#6d28d9)':'#e2e8f0',
              color:masterText.trim()&&!loading?'#fff':'#9ca3af',
              border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'15px',
              cursor:masterText.trim()&&!loading?'pointer':'not-allowed',fontFamily:'inherit',
              boxShadow:masterText.trim()?'0 4px 14px rgba(124,58,237,0.35)':'none'}}>
            {loading?'⏳ Analysing...'
              :storageMode==='user'?'🏭 Generate Warehouse Design (User Defined)'
              :'🏭 Generate Warehouse Design'}
          </button>
          {/* Progress bar */}
          {loading&&(
            <div style={{marginTop:'8px'}}>
              <div style={{display:'flex',justifyContent:'space-between',
                fontSize:'10px',color:'#6b7280',marginBottom:'3px'}}>
                <span>{progressMsg||'Processing…'}</span>
                <span>{progress}%</span>
              </div>
              <div style={{background:'#e2e8f0',borderRadius:'99px',height:'6px',overflow:'hidden'}}>
                <div style={{height:'100%',borderRadius:'99px',
                  background:'linear-gradient(90deg,#7c3aed,#a78bfa)',
                  width:`${progress}%`,transition:'width 0.3s ease'}}/>
              </div>
            </div>
          )}

        </div>{/* END LEFT PANEL */}

        {/* ══ RIGHT PANEL ═══════════════════════════════════════════════ */}
        <div>

          {/* ── USER DEFINED RESULTS ──────────────────────────────────── */}
          {storageMode==='user' && udStep===4 && (userResult||userDesign) && (<>

            {/* ── Detailed rack config cards with bin fit + SKU list ─── */}
            {userRackConfig?.length>0&&(<>
              <div style={{...S.card,marginBottom:'12px',padding:'0',overflow:'hidden'}}>
                <div style={{padding:'10px 16px',background:'#f8fafc',
                  borderBottom:'1px solid #e2e8f0',
                  display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                  <span style={{fontWeight:'700',fontSize:'13px',color:'#0f172a'}}>
                    🗄 Rack Configuration — User Defined
                  </span>
                  <span style={{fontSize:'11px',color:'#6b7280'}}>
                    {userRackConfig.length} rack type{userRackConfig.length>1?'s':''}
                    · {userRackConfig.reduce((s,c)=>s+(c.locs||0),0).toLocaleString()} total locs
                  </span>
                </div>

                {userRackConfig.map((cfg,ci)=>{
                  const RACK_ICONS={shelving:'📦',liveStorage:'🔄',selective:'🏗',
                    doubleDeep:'🔩',driveIn:'🚗',cantilever:'🪵',ground:'🏔'};
                  const isGround = cfg.rack==='ground';
                  const aw = cfg.acrossW||1;
                  const ad = cfg.acrossD||1;
                  const stackH = cfg.stackH||cfg.levels||1;
                  const lvl   = cfg.levels||1;

                  // SKUs that use this bin type
                  const binSkus=(analysis?.slotted||[])
                    .filter(s=>s.bin===cfg.bin)
                    .sort((a,b)=>(b.locsReq||0)-(a.locsReq||0));

                  // Max cells to show in the grid (cap at 12 wide, 8 deep)
                  const showW=Math.min(aw,12), showD=Math.min(ad,8);

                  return(
                    <div key={ci} style={{borderBottom:ci<userRackConfig.length-1?'2px solid #e2e8f0':'none'}}>

                      {/* Card header */}
                      <div style={{padding:'10px 16px',background:'#fff',
                        display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                        <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
                          <span style={{fontSize:'18px'}}>{RACK_ICONS[cfg.rack]||'🏗'}</span>
                          <div>
                            <div style={{fontWeight:'700',fontSize:'13px',color:'#0f172a'}}>
                              {cfg.rackName||cfg.rack}
                            </div>
                            <div style={{fontSize:'10px',color:'#6b7280'}}>
                              {cfg.bayW}×{cfg.bayD}
                              {cfg.shelfH&&!isGround?`×${cfg.shelfH}`:''}mm
                              · {cfg.levels} {isGround?'stack level':'shelf level'}{cfg.levels>1?'s':''}
                            </div>
                          </div>
                        </div>
                        <div style={{textAlign:'right'}}>
                          <div style={{fontWeight:'800',fontSize:'14px',color:'#7c3aed'}}>
                            {(cfg.locs||0).toLocaleString()}
                          </div>
                          <div style={{fontSize:'9px',color:'#9ca3af',fontWeight:'600',
                            textTransform:'uppercase'}}>locations</div>
                        </div>
                      </div>

                      {cfg.autoAssigned&&(
                        <div style={{fontSize:'10px',color:'#7c3aed',fontWeight:'600',
                          background:'#f5f3ff',padding:'3px 16px',
                          borderTop:'1px solid #ede9fe'}}>
                          ⚡ {cfg.autoAssigned}
                        </div>
                      )}

                      {/* Pallet-ization notice */}
                      {cfg.isPalletized&&(
                        <div style={{background:'#fef9c3',borderBottom:'1px solid #fcd34d',
                          padding:'8px 14px',fontSize:'11px',color:'#78350f',
                          borderTop:'1px solid #fcd34d'}}>
                          <div style={{fontWeight:'700',marginBottom:'5px'}}>
                            📦 Bins loaded onto pallets → placed in pallet rack
                          </div>
                          <div style={{display:'flex',gap:'12px',flexWrap:'wrap',fontSize:'10px'}}>
                            <span>
                              Pallet: <strong>{cfg.palletL}×{cfg.palletW}mm</strong>
                            </span>
                            <span>
                              <strong>{cfg.binsPerPallet}</strong> {cfg.bin} bins per pallet
                            </span>
                            <span style={{fontWeight:'700',color:'#92400e'}}>
                              1 pallet = 1 SKU only
                            </span>
                            <span>
                              <strong>{(cfg.locs||0).toLocaleString()}</strong> pallet positions
                              ({cfg.skuCount?.toLocaleString()} SKUs)
                            </span>
                            <span style={{color:'#a16207'}}>
                              ↑ from {(cfg.originalBinLocs||0).toLocaleString()} bin locations
                            </span>
                          </div>
                        </div>
                      )}

                      {/* ── BIN TYPE + FIT LAYOUT ─────────────────────── */}
                      <div style={{padding:'10px 16px',background:'#fafafa',
                        borderTop:'1px solid #f1f5f9',borderBottom:'1px solid #f1f5f9'}}>

                        {/* Bin type badge */}
                        <div style={{display:'flex',alignItems:'center',gap:'8px',marginBottom:'10px'}}>
                          <span style={{background:'#eff6ff',color:'#1d4ed8',fontWeight:'800',
                            fontSize:'13px',padding:'3px 10px',borderRadius:'7px',
                            border:'1px solid #bfdbfe'}}>
                            {cfg.bin}
                          </span>
                          <span style={{fontSize:'11px',color:'#374151',fontWeight:'600'}}>{cfg.binName}</span>
                          {cfg.binDims&&(
                            <span style={{fontSize:'10px',color:'#9ca3af'}}>
                              (L{cfg.binDims[0]}×W{cfg.binDims[1]}×H{cfg.binDims[2]}mm)
                            </span>
                          )}
                        </div>

                        {/* Height breakdown per level */}
                        {!isGround&&cfg.calcNote&&(
                          <div style={{background:'#f1f5f9',borderRadius:'8px',padding:'8px 12px',
                            marginBottom:'10px'}}>
                            <div style={{fontWeight:'700',color:'#475569',marginBottom:'5px',
                              fontSize:'10px',textTransform:'uppercase',letterSpacing:'0.04em'}}>
                              Level Height Calculation (auto)
                            </div>
                            <div style={{fontSize:'10px',color:'#374151',lineHeight:'1.6'}}>
                              {cfg.calcNote}
                            </div>
                            <div style={{marginTop:'5px',fontSize:'11px',fontWeight:'700',
                              color:'#7c3aed'}}>
                              → <strong>{cfg.levels}</strong> level{(cfg.levels||1)>1?'s':''} fit in {cfg.shelfH}mm clear height
                              &nbsp;·&nbsp; 1 bin/pallet per level (no stacking)
                            </div>
                          </div>
                        )}

                        {/* Orientation comparison — LW vs WL */}
                        {cfg.o1&&cfg.o2&&!isGround&&(
                          <div style={{marginBottom:'10px'}}>
                            <div style={{fontSize:'10px',fontWeight:'700',color:'#374151',
                              textTransform:'uppercase',letterSpacing:'0.04em',marginBottom:'6px'}}>
                              Bin Orientations (best selected ✓):
                            </div>
                            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px'}}>
                              {[
                                {label:'Bin L→Width, W→Depth', data:cfg.o1,
                                 active:cfg.orientation==='LW'||cfg.wDimMm===cfg.binDims?.[0]},
                                {label:'Bin W→Width, L→Depth', data:cfg.o2,
                                 active:cfg.orientation==='WL'||cfg.wDimMm===cfg.binDims?.[1]},
                              ].map(({label,data,active},oi)=>(
                                data&&<div key={oi} style={{
                                  border:`2px solid ${active?'#7c3aed':'#e2e8f0'}`,
                                  borderRadius:'8px',padding:'8px 10px',
                                  background:active?'#f5f3ff':'#fff',
                                  opacity:data.feasible?1:0.5}}>
                                  <div style={{display:'flex',justifyContent:'space-between',
                                    alignItems:'center',marginBottom:'3px'}}>
                                    <span style={{fontSize:'10px',fontWeight:'700',
                                      color:active?'#7c3aed':'#6b7280'}}>{label}</span>
                                    {active&&<span style={{fontSize:'9px',fontWeight:'800',
                                      background:'#7c3aed',color:'#fff',
                                      padding:'1px 6px',borderRadius:'4px'}}>BEST</span>}
                                  </div>
                                  {data.feasible
                                    ? <div style={{fontSize:'11px',color:'#374151',lineHeight:'1.5'}}>
                                        <strong>{data.acrossW}</strong> across W ×{' '}
                                        <strong>{data.acrossD}</strong> along D ×{' '}
                                        <strong>{data.levels}</strong> stack{(data.levels||1)>1?'s':''}/level ×{' '}
                                        <strong>{cfg.levels}</strong> level{(cfg.levels||1)>1?'s':''}
                                        <span style={{marginLeft:'6px',fontWeight:'800',color:'#7c3aed'}}>
                                          = {(data.locsPerBay||0).toLocaleString()} locs/bay
                                        </span>
                                      </div>
                                    : <div style={{fontSize:'10px',color:'#be185d'}}>
                                        ✗ Does not fit in this bay
                                      </div>
                                  }
                                  {data.desc&&<div style={{fontSize:'9px',color:'#9ca3af',marginTop:'2px'}}>{data.desc}</div>}
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        {/* ── PER-SKU FITTING TABLE for LONG goods ───────── */}
                        {isGround&&cfg.isPerSku&&(
                          <div style={{marginBottom:'8px'}}>
                            {/* Header */}
                            <div style={{display:'flex',justifyContent:'space-between',
                              alignItems:'center',marginBottom:'6px'}}>
                              <span style={{fontSize:'10px',fontWeight:'700',color:'#374151',
                                textTransform:'uppercase',letterSpacing:'0.04em'}}>
                                Per-SKU Cross-Section Fit in {cfg.bayW}×{cfg.bayD}mm bay
                              </span>
                              <div style={{display:'flex',gap:'8px',fontSize:'10px'}}>
                                <span style={{color:'#166534',fontWeight:'700'}}>
                                  ✓ {(cfg.perSkuFitted||[]).length} fit
                                </span>
                                {(cfg.perSkuUnfitted||[]).length>0&&(
                                  <span style={{color:'#be185d',fontWeight:'700'}}>
                                    ✗ {cfg.perSkuUnfitted.length} don't fit
                                  </span>
                                )}
                              </div>
                            </div>

                            {/* Fitted SKUs */}
                            {(cfg.perSkuFitted||[]).length>0&&(
                              <div style={{border:'1px solid #e2e8f0',borderRadius:'8px',
                                overflow:'hidden',marginBottom:'8px'}}>
                                <div style={{background:'#f0fdf4',padding:'5px 10px',
                                  fontSize:'10px',fontWeight:'700',color:'#166534',
                                  borderBottom:'1px solid #dcfce7'}}>
                                  ✓ Fitted SKUs — cross-section fits in {cfg.bayW}×{cfg.bayD}mm
                                </div>
                                <div style={{overflowX:'auto',maxHeight:'220px',overflowY:'auto'}}>
                                  <table style={{width:'100%',borderCollapse:'collapse',
                                    fontSize:'10px',minWidth:'500px'}}>
                                    <thead>
                                      <tr style={{background:'#f8fafc'}}>
                                        {['SKU','L×W×H (mm)','Cross-section','← W','↔ D','×Stack','Locs','Bays'].map(h=>(
                                          <th key={h} style={{padding:'4px 8px',textAlign:'left',
                                            fontWeight:'700',color:'#64748b',fontSize:'9px',
                                            textTransform:'uppercase',borderBottom:'1px solid #e2e8f0',
                                            whiteSpace:'nowrap'}}>{h}</th>
                                        ))}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {cfg.perSkuFitted.map((s,si)=>(
                                        <tr key={si} style={{background:si%2===0?'#fff':'#f8fafc'}}>
                                          <td style={{padding:'3px 8px',fontWeight:'700',
                                            color:'#0f172a',whiteSpace:'nowrap',fontSize:'10px'}}>{s.sku}</td>
                                          <td style={{padding:'3px 8px',color:'#374151',fontSize:'10px',
                                            whiteSpace:'nowrap'}}>{s.L}×{s.W}×{s.H}</td>
                                          <td style={{padding:'3px 8px',fontSize:'10px',
                                            color:'#6b7280',whiteSpace:'nowrap'}}>
                                            {s.d1}×{s.d2}mm</td>
                                          <td style={{padding:'3px 8px',textAlign:'center',
                                            fontWeight:'700',color:'#1d4ed8',fontSize:'11px'}}>{s.aw}</td>
                                          <td style={{padding:'3px 8px',textAlign:'center',
                                            fontWeight:'700',color:'#166534',fontSize:'11px'}}>{s.ad}</td>
                                          <td style={{padding:'3px 8px',textAlign:'center',
                                            color:'#854d0e',fontSize:'11px'}}>{cfg.levels}</td>
                                          <td style={{padding:'3px 8px',textAlign:'right',
                                            fontWeight:'700',color:'#7c3aed',fontSize:'11px'}}>{s.locsNeeded}</td>
                                          <td style={{padding:'3px 8px',textAlign:'right',
                                            color:'#374151',fontSize:'11px'}}>{s.baysNeeded}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            )}

                            {/* Unfitted SKUs */}
                            {(cfg.perSkuUnfitted||[]).length>0&&(
                              <div style={{border:'1px solid #fecaca',borderRadius:'8px',
                                overflow:'hidden'}}>
                                <div style={{background:'#fff1f2',padding:'5px 10px',
                                  fontSize:'10px',fontWeight:'700',color:'#be185d',
                                  borderBottom:'1px solid #fecaca'}}>
                                  ✗ {cfg.perSkuUnfitted.length} SKU{cfg.perSkuUnfitted.length>1?'s':''} Don't Fit
                                  — cross-section exceeds {cfg.bayW}×{cfg.bayD}mm bay
                                </div>
                                <div style={{overflowX:'auto',maxHeight:'160px',overflowY:'auto'}}>
                                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:'10px'}}>
                                    <thead>
                                      <tr style={{background:'#fef2f2'}}>
                                        {['SKU','L×W×H (mm)','Cross-section','Reason'].map(h=>(
                                          <th key={h} style={{padding:'4px 8px',textAlign:'left',
                                            fontWeight:'700',color:'#9ca3af',fontSize:'9px',
                                            textTransform:'uppercase',borderBottom:'1px solid #fee2e2',
                                            whiteSpace:'nowrap'}}>{h}</th>
                                        ))}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {cfg.perSkuUnfitted.map((s,si)=>(
                                        <tr key={si} style={{background:si%2===0?'#fff':'#fff8f8'}}>
                                          <td style={{padding:'3px 8px',fontWeight:'700',
                                            color:'#be185d',whiteSpace:'nowrap'}}>{s.sku}</td>
                                          <td style={{padding:'3px 8px',color:'#374151',whiteSpace:'nowrap'}}>
                                            {s.L}×{s.W}×{s.H}mm</td>
                                          <td style={{padding:'3px 8px',color:'#be185d',
                                            fontWeight:'600',whiteSpace:'nowrap'}}>
                                            {Math.min(s.L,s.W,s.H)}×{[s.L,s.W,s.H].sort((a,b)=>a-b)[1]}mm
                                          </td>
                                          <td style={{padding:'3px 8px',color:'#6b7280',
                                            fontSize:'9px'}}>{s.reason}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              </div>
                            )}

                            {/* Summary */}
                            <div style={{fontSize:'10px',color:'#374151',marginTop:'6px',
                              fontWeight:'600'}}>
                              Total: {(cfg.perSkuFitted||[]).length} SKUs fitted
                              · <strong style={{color:'#7c3aed'}}>{cfg.locs?.toLocaleString()} locations</strong>
                              · <strong>{cfg.baysNeeded?.toLocaleString()}</strong> bays
                              · {cfg.area?.toFixed(0)}m²
                            </div>
                          </div>
                        )}

                        {/* Ground storage note (non-LONG) */}
                        {isGround&&!cfg.isPerSku&&cfg.orientDesc&&(
                          <div style={{fontSize:'10px',color:'#6b7280',background:'#f1f5f9',
                            borderRadius:'6px',padding:'6px 10px',marginBottom:'8px'}}>
                            📐 {cfg.orientDesc}
                          </div>
                        )}

                        {/* 4-tile summary (skip for per-SKU ground) */}
                        {!cfg.isPerSku&&(<>
                        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',
                          gap:'6px',marginBottom:'8px'}}>
                          {[
                            ['← Width', `${cfg.acrossW} bins`, '#eff6ff','#1d4ed8'],
                            ['↔ Depth',  `${cfg.acrossD} bins`, '#f0fdf4','#166534'],
                            ['↑ Stack/lvl',`${cfg.stackH}×`, '#fef9c3','#854d0e'],
                            ['× Levels', `${cfg.levels}`, '#fdf4ff','#9333ea'],
                          ].map(([l,v,bg,col])=>(
                            <div key={l} style={{background:bg,borderRadius:'8px',
                              padding:'6px 8px',textAlign:'center',border:`1px solid ${col}22`}}>
                              <div style={{fontSize:'9px',color:col,fontWeight:'700',
                                textTransform:'uppercase',marginBottom:'1px'}}>{l}</div>
                              <div style={{fontSize:'13px',fontWeight:'800',color:col}}>{v}</div>
                            </div>
                          ))}
                        </div>

                        {/* Top-view grid */}
                        <div style={{marginBottom:'8px'}}>
                          <div style={{fontSize:'10px',color:'#6b7280',fontWeight:'600',
                            marginBottom:'4px'}}>
                            Top view — 1 level ({aw} wide × {ad} deep){aw>12||ad>8?' — partial':''}:
                          </div>
                          <div style={{display:'inline-block',border:'2px solid #cbd5e1',
                            borderRadius:'6px',padding:'4px',background:'#f1f5f9'}}>
                            {Array.from({length:showD},(_,di)=>(
                              <div key={di} style={{display:'flex',gap:'2px',
                                marginBottom:di<showD-1?'2px':'0'}}>
                                {Array.from({length:showW},(_,wi)=>(
                                  <div key={wi} style={{width:'18px',height:'14px',
                                    borderRadius:'3px',background:'#3b82f6',
                                    border:'1px solid #2563eb',
                                    display:'flex',alignItems:'center',
                                    justifyContent:'center',fontSize:'7px',
                                    color:'#fff',fontWeight:'700'}}>
                                    {wi===0&&di===0?'B':''}
                                  </div>
                                ))}
                                {aw>12&&<span style={{fontSize:'9px',color:'#64748b',
                                  alignSelf:'center',marginLeft:'2px'}}>+{aw-12}</span>}
                              </div>
                            ))}
                            {ad>8&&<div style={{fontSize:'9px',color:'#64748b',
                              textAlign:'center',marginTop:'2px'}}>+{ad-8} rows</div>}
                          </div>
                          <div style={{fontSize:'10px',color:'#374151',marginTop:'5px',fontWeight:'600'}}>
                            {aw}×{ad} = <strong>{aw*ad}</strong> per level
                            × {stackH} stack × {lvl} level{lvl>1?'s':''} =&nbsp;
                            <strong style={{color:'#7c3aed'}}>{(aw*ad*stackH*lvl).toLocaleString()} bins/bay</strong>
                            &nbsp;·&nbsp;<strong>{cfg.baysNeeded?.toLocaleString()}</strong> bays
                          </div>
                        </div>

                        <div style={{display:'flex',gap:'16px',fontSize:'11px',color:'#6b7280',
                          flexWrap:'wrap',borderTop:'1px solid #e2e8f0',paddingTop:'8px'}}>
                          <span>🏗 {cfg.baysNeeded?.toLocaleString()} bays</span>
                          <span>📐 {+(cfg.area||0).toFixed(0)}m² floor area</span>
                          <span>📦 {(aw*ad*stackH*lvl).toLocaleString()} locs/bay</span>
                        </div>
                        </>)}  {/* end !cfg.isPerSku */}
                      </div>
                      {/* ── SKU LIST for this bin type ─────────────────── */}
                      <div style={{borderTop:'1px solid #f1f5f9'}}>
                        <div style={{padding:'7px 16px',background:'#f8fafc',
                          display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                          <span style={{fontSize:'11px',fontWeight:'700',color:'#374151'}}>
                            SKUs in {cfg.bin} bins ({binSkus.length.toLocaleString()} SKUs)
                          </span>
                          <span style={{fontSize:'10px',color:'#9ca3af'}}>
                            {(cfg.locs||0).toLocaleString()} locations total
                          </span>
                        </div>
                        <div style={{maxHeight:'200px',overflowY:'auto'}}>
                          <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                            <thead>
                              <tr style={{background:'#f1f5f9',position:'sticky',top:0}}>
                                {['SKU','Description','Velocity','Stock','Locs'].map(h=>(
                                  <th key={h} style={{padding:'5px 10px',textAlign:'left',
                                    fontWeight:'700',color:'#64748b',
                                    fontSize:'9px',textTransform:'uppercase',
                                    borderBottom:'1px solid #e2e8f0',whiteSpace:'nowrap'}}>
                                    {h}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {binSkus.slice(0,50).map((s,si)=>(
                                <tr key={si} style={{background:si%2===0?'#fff':'#f8fafc'}}>
                                  <td style={{padding:'4px 10px',fontWeight:'700',
                                    color:'#0f172a',whiteSpace:'nowrap',fontSize:'11px'}}>
                                    {s.sku}
                                  </td>
                                  <td style={{padding:'4px 10px',color:'#374151',
                                    maxWidth:'150px',overflow:'hidden',
                                    textOverflow:'ellipsis',whiteSpace:'nowrap',
                                    fontSize:'10px'}}>
                                    {s.desc||s.name||'—'}
                                  </td>
                                  <td style={{padding:'4px 10px'}}>
                                    <span style={{
                                      background:{VF:'#fef9c3',F:'#f0fdf4',M:'#eff6ff',
                                        S:'#fdf4ff',VS:'#f1f5f9',NM:'#f8fafc'}[s.vb||s.velocity]||'#f8fafc',
                                      color:{VF:'#854d0e',F:'#166534',M:'#1d4ed8',
                                        S:'#7c3aed',VS:'#64748b',NM:'#9ca3af'}[s.vb||s.velocity]||'#374151',
                                      padding:'1px 6px',borderRadius:'4px',
                                      fontWeight:'700',fontSize:'9px'}}>
                                      {s.vb||s.velocity||'?'}
                                    </span>
                                  </td>
                                  <td style={{padding:'4px 10px',color:'#374151',
                                    textAlign:'right',fontSize:'11px'}}>
                                    {(s.stock||0).toLocaleString()}
                                  </td>
                                  <td style={{padding:'4px 10px',fontWeight:'700',
                                    color:'#7c3aed',textAlign:'right',fontSize:'11px'}}>
                                    {s.locsReq||1}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                          {binSkus.length>50&&(
                            <div style={{padding:'5px 10px',fontSize:'10px',
                              color:'#9ca3af',fontStyle:'italic',textAlign:'center',
                              borderTop:'1px solid #f1f5f9'}}>
                              Showing 50 of {binSkus.length.toLocaleString()} SKUs
                              — download Excel for full list
                            </div>
                          )}
                          {binSkus.length===0&&(
                            <div style={{padding:'12px',textAlign:'center',
                              fontSize:'11px',color:'#9ca3af'}}>
                              No SKU data available for this bin type
                            </div>
                          )}
                        </div>
                      </div>

                    </div>
                  );
                })}
              </div>
            </>)}

            {/* Overflow section */}
            {(()=>{
              const fittedBins=new Set((userRackConfig||[]).map(c=>c.bin));
              const unfitted={};
              (analysis?.slotted||[]).forEach(s=>{
                if(fittedBins.has(s.bin)) return;
                if(!unfitted[s.bin]) unfitted[s.bin]={
                  binKey:s.bin,binName:s.binName||s.bin,
                  skus:[],totalLocs:0,
                };
                unfitted[s.bin].skus.push(s);
                unfitted[s.bin].totalLocs+=(s.locsReq||1);
              });
              const groups=Object.values(unfitted).sort((a,b)=>b.totalLocs-a.totalLocs);
              if(!groups.length) return null;
              return(
                <div style={{...S.card,padding:'0',overflow:'hidden',marginBottom:'12px',
                  border:'1px solid #fecaca'}}>
                  <div style={{padding:'10px 14px',background:'#fff1f2',
                    borderBottom:'1px solid #fecaca',
                    display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <span style={{fontWeight:'700',fontSize:'12px',color:'#991b1b'}}>
                      ⚠ {groups.length} Bin Type{groups.length>1?'s':''} Don't Fit — Excluded
                    </span>
                    <span style={{fontSize:'11px',color:'#be185d',fontWeight:'600'}}>
                      {groups.reduce((s,g)=>s+g.skus.length,0)} SKUs
                      · {groups.reduce((s,g)=>s+g.totalLocs,0).toLocaleString()} locs
                    </span>
                  </div>
                  {groups.map((grp,gi)=>(
                    <div key={gi} style={{padding:'7px 14px',borderBottom:'1px solid #fee2e2',
                      display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                      <span style={{fontWeight:'700',color:'#be185d',fontSize:'12px'}}>
                        {grp.binKey} — {grp.binName}
                      </span>
                      <span style={{fontSize:'11px',color:'#be185d',fontWeight:'600'}}>
                        {grp.skus.length} SKUs · {grp.totalLocs.toLocaleString()} locs
                      </span>
                    </div>
                  ))}
                </div>
              );
            })()}

            {/* ── RACKING LAYOUT SUMMARY (after rack config, before plan) ──── */}
            {userRackConfig&&userRackConfig.length>0&&(()=>{
              const sm={};
              userRackConfig.forEach(cfg=>{
                const rt=cfg.rack; if(!rt) return;
                if(!sm[rt]) sm[rt]={name:({'shelving':'Shelving','liveStorage':'Flow/Live Storage',
                  'selective':'Selective Pallet','doubleDeep':'Double-Deep','driveIn':'Drive-In',
                  'cantilever':'Cantilever','ground':'Ground Storage'})[rt]||rt, bays:0, locs:0};
                sm[rt].bays+=(cfg.baysNeeded||0); sm[rt].locs+=(cfg.locs||0);
              });
              const rows=Object.values(sm);
              const totBays=rows.reduce((s,r)=>s+r.bays,0);
              const totLocs=rows.reduce((s,r)=>s+r.locs,0);
              if(!rows.length) return null;
              return(
                <div style={{marginBottom:'12px',border:'1px solid #e2e8f0',borderRadius:'10px',
                  overflow:'hidden'}}>
                  <div style={{background:'#0f172a',padding:'8px 14px',
                    fontSize:'12px',fontWeight:'700',color:'#f1f5f9',letterSpacing:'0.03em'}}>
                    📊 Racking Layout Summary
                  </div>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                    <thead>
                      <tr style={{background:'#f1f5f9'}}>
                        {['Rack Type','Bays (B2B total)','Storage Locations'].map(h=>(
                          <th key={h} style={{padding:'6px 10px',textAlign:'left',
                            fontWeight:'700',color:'#475569',fontSize:'10px',
                            textTransform:'uppercase',borderBottom:'1px solid #e2e8f0'}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r,i)=>(
                        <tr key={i} style={{background:i%2===0?'#fff':'#f8fafc'}}>
                          <td style={{padding:'6px 10px',fontWeight:'600',color:'#1d4ed8'}}>{r.name}</td>
                          <td style={{padding:'6px 10px',textAlign:'right',fontWeight:'700',
                            color:'#0f172a'}}>{(r.bays||0).toLocaleString()}</td>
                          <td style={{padding:'6px 10px',textAlign:'right',fontWeight:'700',
                            color:'#7c3aed'}}>{(r.locs||0).toLocaleString()}</td>
                        </tr>
                      ))}
                      <tr style={{background:'#dbeafe',borderTop:'2px solid #93c5fd'}}>
                        <td style={{padding:'7px 10px',fontWeight:'800',color:'#1e40af'}}>TOTAL</td>
                        <td style={{padding:'7px 10px',textAlign:'right',fontWeight:'800',
                          color:'#1e40af'}}>{totBays.toLocaleString()}</td>
                        <td style={{padding:'7px 10px',textAlign:'right',fontWeight:'800',
                          color:'#1e40af'}}>{totLocs.toLocaleString()}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              );
            })()}

            {/* 3D / 2D Layout */}
            {userDesign && (<>
              {/* View mode toggle */}
              <div style={{display:'flex',gap:'8px',marginBottom:'10px'}}>
                {[['2d','🗺 2D Floor Plan'],['3d','🏭 3D View']].map(([m,label])=>(
                  <button key={m} onClick={()=>setUdViewMode(m)}
                    style={{flex:1,padding:'8px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      border:`2px solid ${udViewMode===m?'#7c3aed':'#e2e8f0'}`,
                      background:udViewMode===m?'#f5f3ff':'#fff',
                      color:udViewMode===m?'#7c3aed':'#6b7280'}}>
                    {label}
                  </button>))}
                {udViewMode==='2d'&&(
                  <button onClick={()=>setFloorPlanFS(true)}
                    style={{padding:'8px 12px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      border:'2px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                    ⛶
                  </button>
                )}
              </div>
              <div style={{...S.card,padding:'10px',marginBottom:'12px'}}>
                {/* Plan tools */}
                <div style={{display:'flex',gap:'5px',marginBottom:'6px',flexWrap:'wrap',alignItems:'center'}}>
                  <span style={{width:'1px',height:'16px',background:'#e2e8f0',margin:'0 3px'}}/>
                  <button onClick={toggleMeasure}
                    title="Click two points on the plan to measure the straight distance"
                    style={{padding:'2px 8px',borderRadius:'5px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'10px',fontWeight:'700',
                      border:`1px solid ${measureOn?'#be185d':'#e2e8f0'}`,
                      background:measureOn?'#fff1f2':'#fff',
                      color:measureOn?'#be185d':'#6b7280'}}>
                    📏 Measure{measureOn?' ON':''}
                  </button>
                  {measureOn&&(
                    <button onClick={()=>setSnapOn(v=>!v)}
                      title="Lock to 0°/90° and snap onto rack and wall edges"
                      style={{padding:'2px 8px',borderRadius:'5px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'10px',fontWeight:'700',
                        border:`1px solid ${snapOn?'#059669':'#e2e8f0'}`,
                        background:snapOn?'#f0fdf4':'#fff',
                        color:snapOn?'#059669':'#6b7280'}}>
                      ⊥ Snap {snapOn?'ON':'OFF'}
                    </button>
                  )}
                  {(measurements.length>0||measurePts.length>0)&&(<>
                    <button onClick={undoMeasurement}
                      style={{padding:'2px 7px',borderRadius:'5px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'10px',fontWeight:'700',
                        border:'1px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                      ↶ Undo
                    </button>
                    <button onClick={clearMeasurements}
                      style={{padding:'2px 7px',borderRadius:'5px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'10px',fontWeight:'700',
                        border:'1px solid #fecdd3',background:'#fff1f2',color:'#be185d'}}>
                      ✕ Clear ({measurements.length})
                    </button>
                  </>)}
                  {measureOn&&(
                    <span style={{fontSize:'10px',color:'#be185d',fontWeight:'600'}}>
                      {measurePts.length?'now click the second point':'click the first point'}
                    </span>
                  )}
                </div>
                <div ref={userPlanRefObj}
                  style={{overflow:'auto',maxHeight:'500px',border:'1px solid #e2e8f0',
                  borderRadius:'6px'}}>
                  <FloorPlanSVG analysis={analysis} design={userDesign} params={params}
                    rackConfig={userRackConfig||rackConfig} measureOn={measureOn} measurePts={measurePts}
                    measurements={measurements} onMeasurePoint={onMeasurePoint} snapOn={snapOn}/>
                </div>
                {/* Plan downloads */}
                <div style={{display:'flex',gap:'8px',marginTop:'10px',flexWrap:'wrap'}}>
                  <button onClick={()=>downloadPlan2D('png',userPlanRefObj,2)}
                    style={{padding:'7px 14px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      background:'#eff6ff',border:'1px solid #93c5fd',color:'#1d4ed8'}}>
                    ⬇ Download PNG
                  </button>
                  <button onClick={()=>downloadPlan2D('png',userPlanRefObj,4)}
                    style={{padding:'7px 14px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      background:'#eff6ff',border:'1px solid #93c5fd',color:'#1d4ed8'}}>
                    ⬇ PNG (4× print)
                  </button>
                  <button onClick={()=>downloadPlan2D('svg',userPlanRefObj)}
                    style={{padding:'7px 14px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      background:'#f0fdf4',border:'1px solid #86efac',color:'#166534'}}>
                    ⬇ Download SVG
                  </button>
                  <button onClick={()=>exportDXF(analysis,userDesign,params,userRackConfig||rackConfig)}
                    title="AutoCAD DXF in metres — opens in AutoCAD, BricsCAD, LibreCAD, Revit, SketchUp"
                    style={{padding:'7px 14px',borderRadius:'8px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      background:'#fdf4ff',border:'1px solid #d8b4fe',color:'#7e22ce'}}>
                    📐 Export CAD (DXF)
                  </button>
                </div>
              </div>
              {/* Warehouse size summary */}
              {analysis?.metrics?.hasInv&&(
                <div style={{background:'#f0fdf4',border:'1px solid #86efac',borderRadius:'8px',
                  padding:'8px 14px',marginBottom:'10px',fontSize:'12px',color:'#166534',
                  display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                  <span>
                    ✓ Locations based on <strong>{(analysis.metrics.totSKUs||0).toLocaleString()} SKUs in inventory</strong>
                    {analysis.metrics.totSKUsMaster>analysis.metrics.totSKUs&&(
                      <span style={{color:'#6b7280',fontWeight:'400'}}>
                        {' '}(master has {(analysis.metrics.totSKUsMaster||0).toLocaleString()} total SKUs —
                        {' '}{(analysis.metrics.totSKUsMaster-analysis.metrics.totSKUs).toLocaleString()} not in current stock)
                      </span>
                    )}
                  </span>
                </div>
              )}
              <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'10px',marginBottom:'12px'}}>
                {[
                  ['Gross Area', `${(udSummary?.gross||0).toLocaleString()}m²`, '#eff6ff','#1d4ed8'],
                  ['Dimensions', `${udSummary?.wW||0}×${udSummary?.wL||0}m`, '#f0fdf4','#166534'],
                  ['Rack Area', `${(udSummary?.rackArea||0).toLocaleString()}m²`, '#f5f3ff','#7c3aed'],
                ].map(([l,v,bg,col])=>(
                  <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',
                    textAlign:'center',border:`1px solid ${col}22`}}>
                    <div style={{fontSize:'16px',fontWeight:'800',color:col}}>{v}</div>
                    <div style={{fontSize:'10px',color:'#6b7280',marginTop:'3px',
                      fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                  </div>
                ))}
              </div>
              {udSummary && (
                <div style={{fontSize:'10px',color:'#6b7280',marginTop:'-6px',marginBottom:'12px',
                  textAlign:'center'}}>
                  {udSummary.fromLayout
                    ? 'Derived from the generated floor plan (includes aisles, cross aisles and staging bands)'
                    : 'Estimated — generate the layout for exact dimensions'}
                </div>
              )}
              {/* Download */}
              <div style={{display:'flex',gap:'10px'}}>
                <button onClick={()=>exportExcel(analysis,userDesign,params,userRackConfig,binOverrides)}
                  style={{flex:1,padding:'11px',background:'linear-gradient(135deg,#059669,#047857)',
                    color:'#fff',border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'13px',
                    cursor:'pointer',fontFamily:'inherit'}}>
                  ⬇ Download Excel Report
                </button>
                <button onClick={()=>exportPPT(analysis,userDesign,params,userRackConfig)}
                  style={{flex:1,padding:'11px',background:'linear-gradient(135deg,#7c3aed,#6d28d9)',
                    color:'#fff',border:'none',borderRadius:'10px',fontWeight:'700',fontSize:'13px',
                    cursor:'pointer',fontFamily:'inherit'}}>
                  📊 Download PPT
                </button>
              </div>
            </>)}

          </>)}

          {/* ── FR RESULTS PLACEHOLDER ─── */}
          {storageMode==='fr' && (<>
            {/* ── FORWARD DAYS ────────────────────────────────────────── */}
            <div style={S.card}>
              <div style={S.cardTitle}>📦 Forward Pick + Reserve Storage</div>
              <div style={{fontSize:'11px',color:'#6b7280',marginBottom:'10px'}}>
                Stock is split: a <strong>forward pick face</strong> holds active picking stock
                (days of cover), <strong>reserve</strong> holds the remainder and replenishes
                the forward face. Define both rack types below.
              </div>
              <div style={{display:'flex',alignItems:'center',gap:'10px',marginBottom:'12px',
                background:'#eff6ff',borderRadius:'8px',padding:'8px 12px'}}>
                <span style={{fontSize:'12px',fontWeight:'700',color:'#1d4ed8'}}>
                  📅 Days of Forward Stock Cover:
                </span>
                <input type="number" min="1" max="30" value={forwardDays}
                  onChange={e=>setForwardDays(e.target.value)}
                  style={{...inp,marginBottom:0,width:'60px',fontSize:'14px',
                    fontWeight:'800',textAlign:'center',padding:'4px 6px'}}/>
                <span style={{fontSize:'11px',color:'#6b7280'}}>days</span>
              </div>
            </div>

            {/* ── FORWARD PICK RACKS ─────────────────────────────────── */}
            <div style={S.card}>
              <div style={{fontWeight:'700',fontSize:'12px',color:'#059669',marginBottom:'6px'}}>
                🟢 Forward Pick Racks <span style={{fontSize:'10px',color:'#6b7280',fontWeight:'400'}}>
                  — ergonomic, near dispatch, daily replenished
                </span>
              </div>
              <div style={{border:'1px solid #d1fae5',borderRadius:'8px',overflow:'auto',marginBottom:'6px'}}>
                <table style={{width:'100%',minWidth:'480px',borderCollapse:'collapse',fontSize:'10px'}}>
                  <thead><tr style={{background:'#f0fdf4'}}>
                    {['Name','Type','Bay W','Bay D','Bay H','Levels*',''].map(h=>(
                      <th key={h} style={{padding:'5px 6px',textAlign:'left',fontWeight:'700',
                        color:'#059669',borderBottom:'1px solid #d1fae5',whiteSpace:'nowrap'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {forwardRacks.map((r,i)=>(
                      <tr key={r.id} style={{background:i%2===0?'#fff':'#f9fffe'}}>
                        <td style={{padding:'3px 4px'}}>
                          <input value={r.name} onChange={e=>setForwardRacks(prev=>prev.map(x=>x.id===r.id?{...x,name:e.target.value}:x))}
                            style={{...inp,marginBottom:0,fontSize:'10px',padding:'2px 4px',width:'65px'}}/>
                        </td>
                        <td style={{padding:'3px 4px'}}>
                          <select value={r.rackType||'shelving'} onChange={e=>setForwardRacks(prev=>prev.map(x=>x.id===r.id?{...x,rackType:e.target.value}:x))}
                            style={{...inp,marginBottom:0,fontSize:'10px',padding:'2px 3px',width:'80px'}}>
                            <option value="shelving">Shelving</option>
                            <option value="liveStorage">Flow Rack</option>
                            <option value="selective">Selective</option>
                            <option value="ground">Ground</option>
                          </select>
                        </td>
                        {['bayW','bayD','bayH'].map(f=>(
                          <td key={f} style={{padding:'3px 4px'}}>
                            <input type="number" min="1" value={r[f]} placeholder="mm"
                              onChange={e=>setForwardRacks(prev=>prev.map(x=>x.id===r.id?{...x,[f]:e.target.value}:x))}
                              style={{...inp,marginBottom:0,width:'54px',fontSize:'10px',padding:'2px 4px'}}/>
                          </td>))}
                        <td style={{padding:'3px 4px'}}>
                          <input type="number" min="1" value={r.levels} placeholder="4"
                            onChange={e=>setForwardRacks(prev=>prev.map(x=>x.id===r.id?{...x,levels:e.target.value}:x))}
                            style={{...inp,marginBottom:0,width:'40px',fontSize:'10px',padding:'2px 4px',
                              background:'#fffbeb',border:'1px solid #fcd34d'}}/>
                        </td>
                        <td style={{padding:'3px 4px',textAlign:'center'}}>
                          {forwardRacks.length>1&&<button onClick={()=>setForwardRacks(p=>p.filter(x=>x.id!==r.id))}
                            style={{background:'none',border:'none',color:'#be185d',cursor:'pointer',fontSize:'13px'}}>×</button>}
                        </td>
                      </tr>))}
                  </tbody>
                </table>
                <div style={{fontSize:'9px',color:'#6b7280',padding:'3px 8px',background:'#f9fffe',
                  borderTop:'1px solid #d1fae5'}}>
                  * Shelves per rack (shelving/flow) or stack layers (ground)
                </div>
              </div>
              {forwardRacks.length<3&&<button onClick={()=>setForwardRacks(p=>[...p,{id:Date.now(),name:`Forward ${p.length+1}`,rackType:'shelving',bayW:'',bayD:'',bayH:'',levels:''}])}
                style={{fontSize:'11px',fontWeight:'600',color:'#059669',background:'#f0fdf4',
                  border:'1px dashed #86efac',borderRadius:'6px',padding:'4px 10px',
                  cursor:'pointer',width:'100%'}}>+ Add Forward Rack</button>}
            </div>

            {/* ── RESERVE RACKS ──────────────────────────────────────── */}
            <div style={S.card}>
              <div style={{fontWeight:'700',fontSize:'12px',color:'#7c3aed',marginBottom:'6px'}}>
                🟣 Reserve Storage Racks <span style={{fontSize:'10px',color:'#6b7280',fontWeight:'400'}}>
                  — bulk, forklift access, replenishes forward
                </span>
              </div>
              <div style={{border:'1px solid #ede9fe',borderRadius:'8px',overflow:'auto',marginBottom:'6px'}}>
                <table style={{width:'100%',minWidth:'480px',borderCollapse:'collapse',fontSize:'10px'}}>
                  <thead><tr style={{background:'#f5f3ff'}}>
                    {['Name','Type','Bay W','Bay D','Bay H','Levels*',''].map(h=>(
                      <th key={h} style={{padding:'5px 6px',textAlign:'left',fontWeight:'700',
                        color:'#7c3aed',borderBottom:'1px solid #ede9fe',whiteSpace:'nowrap'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {reserveRacks.map((r,i)=>(
                      <tr key={r.id} style={{background:i%2===0?'#fff':'#faf9ff'}}>
                        <td style={{padding:'3px 4px'}}>
                          <input value={r.name} onChange={e=>setReserveRacks(prev=>prev.map(x=>x.id===r.id?{...x,name:e.target.value}:x))}
                            style={{...inp,marginBottom:0,fontSize:'10px',padding:'2px 4px',width:'65px'}}/>
                        </td>
                        <td style={{padding:'3px 4px'}}>
                          <select value={r.rackType||'selective'} onChange={e=>setReserveRacks(prev=>prev.map(x=>x.id===r.id?{...x,rackType:e.target.value}:x))}
                            style={{...inp,marginBottom:0,fontSize:'10px',padding:'2px 3px',width:'80px'}}>
                            <option value="selective">Selective SPR</option>
                            <option value="doubleDeep">Double-Deep</option>
                            <option value="driveIn">Drive-In</option>
                            <option value="cantilever">Cantilever</option>
                            <option value="ground">Ground</option>
                            <option value="shelving">Shelving</option>
                          </select>
                        </td>
                        {['bayW','bayD','bayH'].map(f=>(
                          <td key={f} style={{padding:'3px 4px'}}>
                            <input type="number" min="1" value={r[f]} placeholder="mm"
                              onChange={e=>setReserveRacks(prev=>prev.map(x=>x.id===r.id?{...x,[f]:e.target.value}:x))}
                              style={{...inp,marginBottom:0,width:'54px',fontSize:'10px',padding:'2px 4px'}}/>
                          </td>))}
                        <td style={{padding:'3px 4px'}}>
                          <input type="number" min="1" value={r.levels} placeholder="4"
                            onChange={e=>setReserveRacks(prev=>prev.map(x=>x.id===r.id?{...x,levels:e.target.value}:x))}
                            style={{...inp,marginBottom:0,width:'40px',fontSize:'10px',padding:'2px 4px',
                              background:'#fffbeb',border:'1px solid #fcd34d'}}/>
                        </td>
                        <td style={{padding:'3px 4px',textAlign:'center'}}>
                          {reserveRacks.length>1&&<button onClick={()=>setReserveRacks(p=>p.filter(x=>x.id!==r.id))}
                            style={{background:'none',border:'none',color:'#be185d',cursor:'pointer',fontSize:'13px'}}>×</button>}
                        </td>
                      </tr>))}
                  </tbody>
                </table>
                <div style={{fontSize:'9px',color:'#6b7280',padding:'3px 8px',background:'#faf9ff',
                  borderTop:'1px solid #ede9fe'}}>
                  * Shelf levels or stack layers
                </div>
              </div>
              {reserveRacks.length<4&&<button onClick={()=>setReserveRacks(p=>[...p,{id:Date.now(),name:`Reserve ${p.length+1}`,rackType:'selective',bayW:'',bayD:'',bayH:'',levels:''}])}
                style={{fontSize:'11px',fontWeight:'600',color:'#7c3aed',background:'#f5f3ff',
                  border:'1px dashed #c4b5fd',borderRadius:'6px',padding:'4px 10px',
                  cursor:'pointer',width:'100%'}}>+ Add Reserve Rack</button>}
            </div>

            {!analysis&&<div style={{fontSize:'11px',color:'#9ca3af',textAlign:'center',
              padding:'8px',background:'#f8fafc',borderRadius:'6px'}}>
              Run "Generate Warehouse Design" above to load SKU data first
            </div>}
          </>)}

          {storageMode==='system' && (<>
          {/* ── FORWARD PICK + RESERVE INPUT PANEL ─────────────────────── */}
          {/* End of user results — system results start */}
            {analysis?.binSummary && Object.keys(analysis.binSummary).length > 0 && (
              <div style={{...S.card,marginBottom:'12px'}}>
                <div style={{fontWeight:'700',fontSize:'13px',color:'#0f172a',marginBottom:'8px'}}>
                  📦 Bin Variety Summary
                  <span style={{fontSize:'11px',fontWeight:'400',color:'#6b7280',marginLeft:'8px'}}>
                    {Object.keys(analysis?.binSummary||{}).length} bin type{Object.keys(analysis?.binSummary||{}).length>1?'s':''} in use
                  </span>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(140px,1fr))',gap:'8px',marginBottom:'10px'}}>
                  {Object.entries(analysis?.binSummary||{})
                    .sort((a,b)=>['XS','S','M','L','XL','LONG'].indexOf(a[0])-['XS','S','M','L','XL','LONG'].indexOf(b[0]))
                    .map(([band,info])=>{
                      const pct = Math.round(info.skus/(analysis?.metrics?.totSKUs||1)*100);
                      const COLORS={XS:['#f1f5f9','#64748b'],S:['#eff6ff','#1d4ed8'],
                        M:['#f5f3ff','#7c3aed'],L:['#f0fdf4','#166534'],
                        XL:['#fef9c3','#854d0e'],LONG:['#fdf4ff','#9333ea']};
                      const [bg,col]=COLORS[band]||['#f8fafc','#374151'];
                      return(
                        <div key={band} style={{background:bg,border:`1px solid ${col}33`,
                          borderRadius:'8px',padding:'8px 10px'}}>
                          <div style={{fontWeight:'800',fontSize:'14px',color:col}}>{band}</div>
                          <div style={{fontSize:'10px',color:col,opacity:0.8,marginBottom:'3px'}}>{info.name}</div>
                          <div style={{fontSize:'12px',fontWeight:'600',color:'#0f172a'}}>{(info.skus||0).toLocaleString()} SKUs</div>
                          <div style={{fontSize:'11px',color:'#6b7280',marginBottom:'5px'}}>
                            {(info.locs||0).toLocaleString()} locs · {pct}% of SKUs
                            {info.upgrades>0&&<span style={{marginLeft:'6px',background:'#eff6ff',
                              color:'#1d4ed8',borderRadius:'99px',padding:'1px 6px',fontSize:'10px',fontWeight:'700'}}>
                              ↑{info.upgrades} qty-upgraded
                            </span>}
                          </div>
                          {/* Bin utilization bar */}
                          <div style={{marginTop:'4px'}}>
                            <div style={{display:'flex',justifyContent:'space-between',
                              fontSize:'10px',marginBottom:'2px'}}>
                              <span style={{color:col,fontWeight:'700'}}>Bin utilisation</span>
                              <span style={{color:info.utilPct>=80?'#166534':info.utilPct>=50?'#854d0e':'#be185d',
                                fontWeight:'800'}}>{info.utilPct}%</span>
                            </div>
                            <div style={{background:'#e2e8f0',borderRadius:'99px',height:'5px'}}>
                              <div style={{height:'5px',borderRadius:'99px',
                                background:info.utilPct>=80?'#16a34a':info.utilPct>=50?'#d97706':'#be185d',
                                width:`${Math.min(info.utilPct,100)}%`,transition:'width 0.3s'}}/>
                            </div>
                            <div style={{fontSize:'9px',color:'#9ca3af',marginTop:'2px'}}>
                              {(info.stock||0).toLocaleString()} units in {(info.capacity||0).toLocaleString()} capacity
                            </div>
                          </div>
                        </div>
                      );
                    })}
                </div>
                {/* Quantity upgrade note */}
                {(analysis?.totalQtyUpgrades||0) > 0 && (
                  <div style={{background:'#eff6ff',border:'1px solid #93c5fd',
                    borderRadius:'8px',padding:'8px 12px',fontSize:'12px',color:'#1d4ed8',marginBottom:'6px'}}>
                    <strong>📦 Quantity-driven upgrades:</strong>{' '}
                    {(analysis.totalQtyUpgrades||0).toLocaleString()} SKU{(analysis.totalQtyUpgrades||0)>1?'s were':' was'} assigned
                    to a larger bin than their size alone required, to fit stock in ≤{BIN_LOC_TARGET} locations.
                    This reduces total location count significantly.
                  </div>
                )}
                {/* Consolidation report */}
                {analysis.binConsolidation && analysis.binConsolidation.length > 0 && (
                  <div style={{background:'#f0fdf4',border:'1px solid #86efac',
                    borderRadius:'8px',padding:'8px 12px',fontSize:'12px',color:'#166534'}}>
                    <strong>✓ Auto-consolidated:</strong>
                    {analysis.binConsolidation.map(r=>(
                      <span key={r.from} style={{marginLeft:'8px'}}>
                        {r.totalMoved} SKU{r.totalMoved>1?'s':''} from <strong>{r.from}</strong> →{' '}
                        {r.actions.map(a=>`${a.to} (${a.n})`).join(', ')}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* ── BIN / PALLET SIZE + COUNT EDITOR ─────────────────────── */}
            {analysis?.binSummary && Object.keys(analysis.binSummary).length > 0 && binOverrides && (
              <div style={{...S.card, marginBottom:'12px',
                border:'2px solid #c4b5fd', background:'linear-gradient(180deg,#faf8ff,#fff)'}}>
                <div onClick={()=>setBinEditsOpen(o=>!o)}
                  style={{display:'flex',justifyContent:'space-between',
                    alignItems:'center',cursor:'pointer'}}>
                  <div style={{display:'flex',alignItems:'center',gap:'10px'}}>
                    <div style={{width:'26px',height:'26px',borderRadius:'50%',
                      background:'#7c3aed',color:'#fff',display:'flex',
                      alignItems:'center',justifyContent:'center',
                      fontSize:'13px',fontWeight:'800',flexShrink:0}}>1</div>
                    <div>
                      <div style={{fontWeight:'800',fontSize:'14px',color:'#5b21b6'}}>
                        ✏️ Edit Bin &amp; Pallet Sizes
                        <span style={{fontSize:'10px',fontWeight:'700',color:'#7c3aed',
                          background:'#f5f3ff',border:'1px solid #c4b5fd',
                          borderRadius:'99px',padding:'2px 8px',marginLeft:'8px'}}>
                          OPTIONAL
                        </span>
                      </div>
                      <div style={{fontSize:'11px',color:'#6b7280',marginTop:'2px'}}>
                        Change dimensions or quantities, or add up to 3 sizes per container type.
                        Each size becomes its own rack configuration and layout block.
                      </div>
                    </div>
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
                    {binEditsDirty && (
                      <span style={{fontSize:'11px',fontWeight:'700',color:'#d97706',
                        background:'#fffbeb',border:'1px solid #fde68a',
                        borderRadius:'99px',padding:'3px 10px',whiteSpace:'nowrap'}}>
                        ⚠ Unapplied edits
                      </span>
                    )}
                    <span style={{fontSize:'15px',color:'#7c3aed',fontWeight:'700'}}>
                      {binEditsOpen ? '▾' : '▸'}
                    </span>
                  </div>
                </div>

                {binEditsOpen && (
                  <div style={{marginTop:'12px'}}>
                    <div style={{overflowX:'auto'}}>
                        <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                          <thead>
                            <tr style={{background:'#f8fafc'}}>
                              {['Bin','Container Type','Size','Length (mm)','Width (mm)','Height (mm)','Qty (locations)','Volume',''].map((h,hi)=>(
                                <th key={hi} style={{padding:'7px 10px',textAlign:'left',
                                  fontSize:'10px',fontWeight:'700',color:'#6b7280',
                                  textTransform:'uppercase',borderBottom:'1px solid #e2e8f0',
                                  whiteSpace:'nowrap'}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {Object.entries(analysis?.binSummary||{})
                              .sort((a,b)=>['XS','S','M','L','XL','LONG'].indexOf(a[0])
                                          -['XS','S','M','L','XL','LONG'].indexOf(b[0]))
                              .map(([band,info])=>{
                                const variants = binOverrides[band] || [];
                                const base = BIN_CATALOG[band] ? BIN_CATALOG[band].phys : null;
                                const isPallet  = band==='XL' || band==='L';
                                const dimLocked = !base;
                                const COLORS={XS:['#f1f5f9','#64748b'],S:['#eff6ff','#1d4ed8'],
                                  M:['#f5f3ff','#7c3aed'],L:['#f0fdf4','#166534'],
                                  XL:['#fef9c3','#854d0e'],LONG:['#fdf4ff','#9333ea']};
                                const [bg,col]=COLORS[band]||['#f8fafc','#374151'];
                                const inputS={border:'1px solid #e2e8f0',borderRadius:'6px',
                                  padding:'5px 7px',fontSize:'12px',width:'82px',
                                  boxSizing:'border-box',fontFamily:'inherit',outline:'none'};
                                const totalQty = variants.reduce((s2,v)=>s2+(parseFloat(v.locs)||0),0);
                                const baseQty  = info.locs||0;
                                return variants.map((ov,vi)=>{
                                  const bL=parseFloat(ov.L)||0;
                                  const bW=parseFloat(ov.W)||0;
                                  const bH=parseFloat(ov.H)||0;
                                  const volM3=(bL*bW*bH)/1e9;
                                  const baseVol=base?(base[0]*base[1]*base[2])/1e9:0;
                                  const volPct=baseVol>0?Math.round(((volM3-baseVol)/baseVol)*100):0;
                                  const isFirst=vi===0;
                                  return (
                                    <tr key={band+'-'+vi} style={{
                                      borderBottom: vi===variants.length-1
                                        ? '1px solid #e2e8f0' : '1px dashed #f1f5f9'}}>
                                      <td style={{padding:'7px 10px',verticalAlign:'top'}}>
                                        {isFirst && (
                                          <span style={{background:bg,color:col,fontWeight:'800',
                                            fontSize:'12px',borderRadius:'6px',padding:'3px 9px'}}>
                                            {band}
                                          </span>
                                        )}
                                      </td>
                                      <td style={{padding:'7px 10px',color:'#374151',verticalAlign:'top'}}>
                                        {isFirst && (<>
                                          {info.name}
                                          <span style={{marginLeft:'6px',fontSize:'10px',fontWeight:'700',
                                            color:isPallet?'#854d0e':'#1d4ed8',
                                            background:isPallet?'#fef9c3':'#eff6ff',
                                            borderRadius:'99px',padding:'1px 7px'}}>
                                            {isPallet?'PALLET':'BIN'}
                                          </span>
                                          {totalQty!==baseQty && (
                                            <div style={{fontSize:'10px',fontWeight:'700',marginTop:'3px',
                                              color:totalQty>baseQty?'#166534':'#be185d'}}>
                                              total {totalQty.toLocaleString()} vs {baseQty.toLocaleString()} system
                                            </div>
                                          )}
                                        </>)}
                                      </td>
                                      <td style={{padding:'7px 10px'}}>
                                        <input type="text" value={ov.label||''}
                                          placeholder={'Size '+(vi+1)}
                                          onChange={e=>updateBinField(band,vi,'label',e.target.value)}
                                          style={{...inputS,width:'88px',fontWeight:'700',color:col}}/>
                                      </td>
                                      <td style={{padding:'7px 10px'}}>
                                        <input type="number" min="50" step="10" disabled={dimLocked}
                                          value={ov.L||''}
                                          onChange={e=>updateBinField(band,vi,'L',e.target.value)}
                                          style={{...inputS,background:dimLocked?'#f8fafc':'#fff'}}/>
                                      </td>
                                      <td style={{padding:'7px 10px'}}>
                                        <input type="number" min="50" step="10" disabled={dimLocked}
                                          value={ov.W||''}
                                          onChange={e=>updateBinField(band,vi,'W',e.target.value)}
                                          style={{...inputS,background:dimLocked?'#f8fafc':'#fff'}}/>
                                      </td>
                                      <td style={{padding:'7px 10px'}}>
                                        <input type="number" min="50" step="10" disabled={dimLocked}
                                          value={ov.H||''}
                                          onChange={e=>updateBinField(band,vi,'H',e.target.value)}
                                          style={{...inputS,background:dimLocked?'#f8fafc':'#fff'}}/>
                                      </td>
                                      <td style={{padding:'7px 10px'}}>
                                        <input type="number" min="1" step="1"
                                          value={ov.locs||''}
                                          onChange={e=>updateBinField(band,vi,'locs',e.target.value)}
                                          style={{...inputS,width:'96px'}}/>
                                      </td>
                                      <td style={{padding:'7px 10px',color:'#6b7280',whiteSpace:'nowrap'}}>
                                        {volM3>0?volM3.toFixed(3):'-'} m³
                                        {volPct!==0 && (
                                          <div style={{fontSize:'10px',fontWeight:'700',
                                            color:volPct>0?'#d97706':'#166534'}}>
                                            {volPct>0?'+':''}{volPct}% vs default
                                          </div>
                                        )}
                                      </td>
                                      <td style={{padding:'7px 10px',whiteSpace:'nowrap'}}>
                                        {variants.length>1 && (
                                          <button onClick={()=>removeBinVariant(band,vi)}
                                            title="Remove this size"
                                            style={{border:'1px solid #fecdd3',background:'#fff1f2',
                                              color:'#be185d',borderRadius:'6px',padding:'3px 8px',
                                              fontSize:'11px',fontWeight:'700',cursor:'pointer',
                                              fontFamily:'inherit'}}>×</button>
                                        )}
                                        {isFirst && variants.length<MAX_BIN_VARIANTS && !dimLocked && (
                                          <button onClick={()=>addBinVariant(band)}
                                            title="Add another size for this bin"
                                            style={{marginLeft:'4px',border:'1px dashed #c4b5fd',
                                              background:'#f5f3ff',color:'#7c3aed',borderRadius:'6px',
                                              padding:'3px 8px',fontSize:'11px',fontWeight:'700',
                                              cursor:'pointer',fontFamily:'inherit'}}>+ size</button>
                                        )}
                                      </td>
                                    </tr>
                                  );
                                });
                              })}
                          </tbody>
                        </table>
                    </div>

                    <div style={{background:'#eff6ff',border:'1px solid #93c5fd',
                      borderRadius:'8px',padding:'8px 12px',fontSize:'11px',
                      color:'#1d4ed8',marginTop:'10px'}}>
                      <strong>How this works:</strong> each size becomes its own rack
                      configuration and its own block in the floor plan, sized from the
                      dimensions you enter. Use <strong>+ size</strong> to define up to
                      {' '}{MAX_BIN_VARIANTS} different sizes per container type and split the
                      quantity between them. Click <strong>Apply</strong> to rebuild the rack
                      configuration and warehouse layout.
                    </div>

                    <div style={{display:'flex',gap:'8px',marginTop:'10px'}}>
                      <button onClick={applyBinEdits}
                        style={{flex:1,padding:'10px',borderRadius:'8px',border:'none',
                          fontWeight:'700',fontSize:'13px',fontFamily:'inherit',
                          cursor:'pointer',color:'#fff',
                          background:'linear-gradient(135deg,#7c3aed,#6d28d9)',
                          boxShadow:'0 3px 10px rgba(124,58,237,0.3)'}}>
                        ✓ Apply &amp; Rebuild Rack Config
                      </button>
                      <button onClick={resetBinEdits}
                        style={{padding:'10px 16px',borderRadius:'8px',
                          border:'1px solid #e2e8f0',background:'#fff',color:'#6b7280',
                          fontWeight:'700',fontSize:'13px',fontFamily:'inherit',
                          cursor:'pointer',whiteSpace:'nowrap'}}>
                        ↺ Reset to System
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* ── RACK CONFIGURATION EDITOR ───────────────────────────── */}
            {rackConfig && (
              <div style={{...S.card, marginBottom:'12px'}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'4px'}}>
                  <div style={{display:'flex',alignItems:'center',gap:'10px'}}>
                    <div style={{width:'26px',height:'26px',borderRadius:'50%',
                      background:'#0f172a',color:'#fff',display:'flex',
                      alignItems:'center',justifyContent:'center',
                      fontSize:'13px',fontWeight:'800',flexShrink:0}}>2</div>
                    <div style={{fontWeight:'800',fontSize:'14px',color:'#0f172a'}}>
                      🗄 Rack Configuration
                    </div>
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
                  <span style={{marginLeft:'8px',background:'#fffbeb',border:'1px solid #fde68a',
                    borderRadius:'6px',padding:'2px 8px',fontSize:'11px',color:'#92400e'}}>
                    ℹ S movers: 2 SKUs/loc · VS: 4/loc · NM: 8/loc (pick-face sharing applied)
                  </span>
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
                          {(cfg.locs||0).toLocaleString()} locations needed
                        </span>
                      </div>

                      {/* Elevation view + params side by side */}
                      <div style={{display:'grid',gridTemplateColumns:'260px 1fr',gap:'0'}}>
                        {/* Elevation diagram */}
                        <div style={{padding:'10px',background:'#fafafa',
                          borderRight:'1px solid #e2e8f0',display:'flex',flexDirection:'column',
                          alignItems:'center',gap:'6px'}}>
                          <div style={{fontSize:'9px',color:'#9ca3af',fontWeight:'700',
                            textTransform:'uppercase',letterSpacing:'0.05em'}}>
                            Elevation View — Front
                          </div>
                          <RackElevationSVG cfg={cfg} W={240} H={175}/>
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

                        {/* Result row + download */}
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
                          <div style={{display:'flex',alignItems:'center',gap:'12px',flexWrap:'wrap'}}>
                            <div style={{display:'flex',gap:'12px',fontSize:'12px',flexWrap:'wrap',alignItems:'center'}}>
                              <span><strong style={{color:'#7c3aed'}}>{cfg.baysNeeded}</strong> bays</span>
                              <span><strong style={{color:'#0369a1'}}>{((cfg.bayW/1000)*(cfg.bayD/1000)).toFixed(2)}m²</strong>/bay</span>
                              <span><strong style={{color:'#059669'}}>{cfg.area}m²</strong> total</span>
                              {/* Utilization from binSummary */}
                              {analysis?.binSummary?.[cfg.bin] && (() => {
                                const bi = analysis?.binSummary?.[cfg.bin];
                                const util = bi.utilPct || 0;
                                const uColor = util>=80?'#166534':util>=50?'#d97706':'#be185d';
                                return (
                                  <span style={{display:'inline-flex',alignItems:'center',
                                    gap:'5px',background:'#f8fafc',border:'1px solid #e2e8f0',
                                    borderRadius:'6px',padding:'2px 8px'}}>
                                    <div style={{width:'50px',background:'#e2e8f0',borderRadius:'99px',height:'5px'}}>
                                      <div style={{height:'5px',borderRadius:'99px',background:uColor,
                                        width:`${Math.min(util,100)}%`}}/>
                                    </div>
                                    <strong style={{color:uColor,fontSize:'12px'}}>{util}%</strong>
                                    <span style={{color:'#9ca3af',fontSize:'10px'}}>utilised</span>
                                  </span>
                                );
                              })()}
                            </div>
                            <button
                              onClick={()=>downloadRackLocations(cfg, analysis)}
                              title="Download location calculation with per-SKU detail"
                              style={{padding:'5px 12px',background:'#f0fdf4',
                                border:'1px solid #86efac',borderRadius:'7px',
                                fontSize:'11px',fontWeight:'700',color:'#166534',
                                cursor:'pointer',fontFamily:'inherit',whiteSpace:'nowrap'}}>
                              ⬇ Locations
                            </button>
                          </div>
                        </div>
                        </div>
                        {/* end params column */}
                      </div>
                      {/* end elevation+params grid */}
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
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'12px'}}>
                <div style={{fontWeight:'700',fontSize:'14px',color:'#0f172a'}}>
                  🗺 Plan View (Top)
                </div>
                <div style={{display:'flex',gap:'6px'}}>
                  <button onClick={()=>setFloorPlanFS(true)}
                    title="View full screen"
                    style={{padding:'5px 12px',borderRadius:'7px',cursor:'pointer',
                      fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                      border:'2px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                    ⛶ Full Screen
                  </button>
                </div>
              </div>

              {(<>
                    {/* Plan tools */}
                    <div style={{display:'flex',alignItems:'center',gap:'6px',
                      marginBottom:'6px',flexWrap:'wrap'}}>
                      <button onClick={toggleMeasure}
                        title="Click two points on the plan to measure the straight distance"
                        style={{padding:'3px 9px',borderRadius:'6px',cursor:'pointer',
                          fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                          border:`1px solid ${measureOn?'#be185d':'#e2e8f0'}`,
                          background:measureOn?'#fff1f2':'#fff',
                          color:measureOn?'#be185d':'#6b7280'}}>
                        📏 Measure{measureOn?' ON':''}
                      </button>
                      {measureOn&&(
                        <button onClick={()=>setSnapOn(v=>!v)}
                          title="Lock to 0°/90° and snap onto rack and wall edges"
                          style={{padding:'3px 9px',borderRadius:'6px',cursor:'pointer',
                            fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                            border:`1px solid ${snapOn?'#059669':'#e2e8f0'}`,
                            background:snapOn?'#f0fdf4':'#fff',
                            color:snapOn?'#059669':'#6b7280'}}>
                          ⊥ Snap {snapOn?'ON':'OFF'}
                        </button>
                      )}
                      {(measurements.length>0||measurePts.length>0)&&(<>
                        <button onClick={undoMeasurement}
                          style={{padding:'3px 8px',borderRadius:'6px',cursor:'pointer',
                            fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                            border:'1px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                          ↶ Undo
                        </button>
                        <button onClick={clearMeasurements}
                          style={{padding:'3px 8px',borderRadius:'6px',cursor:'pointer',
                            fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                            border:'1px solid #fecdd3',background:'#fff1f2',color:'#be185d'}}>
                          ✕ Clear ({measurements.length})
                        </button>
                      </>)}
                      <span style={{fontSize:'10px',color:measureOn?'#be185d':'#9ca3af',marginLeft:'4px',
                        fontWeight:measureOn?'600':'400'}}>
                        {measureOn
                          ? (measurePts.length?'Now click the second point':'Click the first point')
                          : 'Scroll inside the plan to pan'}
                      </span>
                    </div>
                    {/* Scrollable plan container */}
                    <div ref={r=>{plan2DRef.current=r; planScrollRef.current=r;}}
                      style={{overflow:'auto',border:'1px solid #e2e8f0',borderRadius:'8px',
                        maxHeight:'600px',background:'#f8fafc',cursor:'grab'}}>
                      <FloorPlanSVG analysis={analysis} design={design} params={params}
                        rackConfig={rackConfig} measureOn={measureOn} measurePts={measurePts}
                    measurements={measurements} onMeasurePoint={onMeasurePoint} snapOn={snapOn}/>
                    </div>
                  </>)}

              {/* Legend + Download */}
              {(
                <div style={{marginTop:'12px'}}>
                  {/* Download buttons */}
                  <div style={{display:'flex',gap:'8px',marginBottom:'10px',flexWrap:'wrap'}}>
                    <button onClick={()=>downloadPlan2D('svg')}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#f0fdf4',border:'1px solid #86efac',color:'#166534'}}>
                      ⬇ Download SVG
                    </button>
                    <button onClick={()=>downloadPlan2D('png',plan2DRef,2)}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#eff6ff',border:'1px solid #93c5fd',color:'#1d4ed8'}}>
                      ⬇ Download PNG
                    </button>
                    <button onClick={()=>downloadPlan2D('png',plan2DRef,4)}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#eff6ff',border:'1px solid #93c5fd',color:'#1d4ed8'}}>
                      ⬇ PNG (4× print)
                    </button>
                    <button onClick={()=>exportDXF(analysis,design,params,rackConfig)}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#fef9c3',border:'1px solid #fde047',color:'#854d0e'}}>
                      📐 Export CAD (DXF)
                    </button>
                    <span style={{fontSize:'11px',color:'#9ca3af',alignSelf:'center'}}>
                      DXF is in metres on named layers · opens in AutoCAD, BricsCAD,
                      LibreCAD, FreeCAD, Revit, SketchUp · save as DWG in AutoCAD
                    </span>
                  </div>
                  {/* Legend */}
                  <div style={{display:'flex',gap:'12px',flexWrap:'wrap'}}>
                    {Object.entries(ZONE_DEFS).map(([k,z])=>(
                      <div key={k} style={{display:'flex',alignItems:'center',gap:'5px',fontSize:'11px'}}>
                        <div style={{width:'14px',height:'14px',background:z.color,border:`1px solid ${z.border}`,borderRadius:'3px'}}/>
                        <span style={{color:z.textColor,fontWeight:'600'}}>{z.label}</span>
                      </div>))}
                    <div style={{display:'flex',alignItems:'center',gap:'5px',fontSize:'11px'}}>
                      <div style={{width:'14px',height:'14px',background:'#e0f2fe',border:'1px solid #0284c7',borderRadius:'3px'}}/>
                      <span style={{color:'#0369a1',fontWeight:'600'}}>Receiving</span>
                    </div>
                  </div>
                </div>
              )}
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
                  {Object.entries(analysis?.zoneSummary||{}).map(([z,v],i)=>(
                    <tr key={z} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                      <td style={{padding:'8px 12px'}}>
                        <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
                          <div style={{width:'10px',height:'10px',borderRadius:'50%',
                            background:ZONE_DEFS[z]?.border||'#ccc',flexShrink:0}}/>
                          <span style={{fontWeight:'600'}}>{ZONE_DEFS[z]?.label||z}</span>
                        </div>
                      </td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{(v.skus||0).toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right',fontWeight:'700',color:'#7c3aed'}}>{(v.locs||0).toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{(v.stock||0).toLocaleString()}</td>
                      <td style={{padding:'8px 12px',textAlign:'right'}}>{(v.pickLines||0).toLocaleString()}</td>
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
                  {Object.entries(analysis?.rackSummary||{}).map(([rk,rv],i)=>(
                    <tr key={rk} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                      <td style={{padding:'8px 12px'}}>
                        <div style={{fontWeight:'600'}}>{RACK_DEFS[rk]?.name||rk}</div>
                        <div style={{fontSize:'11px',color:'#9ca3af'}}>{RACK_DEFS[rk]?.desc}</div>
                      </td>
                      <td style={{padding:'8px 12px',textAlign:'right',fontWeight:'700'}}>{r(v.locs||0).toLocaleString()}</td>
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
                      const row = ['XS','S','M','L','XL'].map(s=>(analysis?.matrix||{})[`${v}-${s}`]||0);
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
                        const t=['VF','F','M','S','VS','NM'].reduce((sum,v)=>sum+((analysis?.matrix||{})[`${v}-${s}`]||0),0);
                        return<td key={s} style={{padding:'7px 12px',textAlign:'center'}}>{t.toLocaleString()}</td>;
                      })}
                      <td style={{padding:'7px 12px',textAlign:'center',color:'#7c3aed'}}>
                        {(analysis.metrics?.totLocs||0).toLocaleString()}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            {/* Layout-derived area summary */}
            {sysSummary && (
              <div style={{marginBottom:'12px'}}>
                <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'10px'}}>
                  {[
                    ['Gross Area', `${(sysSummary.gross||0).toLocaleString()}m²`, '#eff6ff','#1d4ed8'],
                    ['Dimensions', `${sysSummary.wW||0}×${sysSummary.wL||0}m`, '#f0fdf4','#166534'],
                    ['Rack Area', `${(sysSummary.rackArea||0).toLocaleString()}m²`, '#f5f3ff','#7c3aed'],
                  ].map(([l,v,bg,col])=>(
                    <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',
                      textAlign:'center',border:`1px solid ${col}22`}}>
                      <div style={{fontSize:'16px',fontWeight:'800',color:col}}>{v}</div>
                      <div style={{fontSize:'10px',color:'#6b7280',marginTop:'3px',
                        fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:'10px',color:'#6b7280',marginTop:'6px',textAlign:'center'}}>
                  {sysSummary.fromLayout
                    ? 'Derived from the generated floor plan (includes aisles, cross aisles and staging bands)'
                    : 'Estimated — generate the layout for exact dimensions'}
                </div>
              </div>
            )}

            {/* Download buttons */}
            <div style={{display:'flex',gap:'12px'}}>
              <button onClick={()=>exportExcel(analysis,design,params,rackConfig,binOverrides)}
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
          {/* End system mode */}
        </div>
      </div>

      {/* ── FULLSCREEN FLOOR PLAN OVERLAY ──────────────────────────── */}
      {floorPlanFS&&(
        <div style={{position:'fixed',inset:0,background:'#fff',zIndex:9999,
          display:'flex',flexDirection:'column'}}>
          {/* Toolbar */}
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',
            padding:'10px 20px',background:'#0f172a',flexShrink:0}}>
            <div style={{color:'#fff',fontWeight:'700',fontSize:'15px'}}>
              🗺 Warehouse Floor Plan — Full Screen
              {design&&<span style={{fontSize:'12px',fontWeight:'400',color:'#94a3b8',
                marginLeft:'12px'}}>{design.wW}m × {design.wL}m · {(design.wW*design.wL).toLocaleString()}m²</span>}
            </div>
            <div style={{display:'flex',gap:'6px',alignItems:'center'}}>
              <button onClick={toggleMeasure}
                title="Click two points on the plan to measure the straight distance"
                style={{padding:'4px 10px',borderRadius:'6px',cursor:'pointer',border:'none',
                  fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                  background:measureOn?'#be185d':'#1e293b',
                  color:measureOn?'#fff':'#94a3b8'}}>
                📏 Measure{measureOn?' ON':''}
              </button>
              {measureOn&&(
                <button onClick={()=>setSnapOn(v=>!v)}
                  title="Lock to 0°/90° and snap onto rack and wall edges"
                  style={{padding:'4px 10px',borderRadius:'6px',cursor:'pointer',border:'none',
                    fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                    background:snapOn?'#059669':'#1e293b',
                    color:snapOn?'#fff':'#94a3b8'}}>
                  ⊥ Snap {snapOn?'ON':'OFF'}
                </button>
              )}
              {(measurements.length>0||measurePts.length>0)&&(<>
                <button onClick={undoMeasurement}
                  style={{padding:'4px 8px',borderRadius:'6px',cursor:'pointer',border:'none',
                    fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                    background:'#1e293b',color:'#94a3b8'}}>↶</button>
                <button onClick={clearMeasurements}
                  style={{padding:'4px 8px',borderRadius:'6px',cursor:'pointer',border:'none',
                    fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                    background:'#1e293b',color:'#fb7185'}}>
                  ✕ {measurements.length}
                </button>
              </>)}
              {measureOn&&(
                <span style={{fontSize:'11px',color:'#fb7185',fontWeight:'600'}}>
                  {measurePts.length?'2nd point…':'1st point…'}
                </span>
              )}
              {/* Download in fullscreen */}
              <button onClick={()=>{
                  // Find the actual <svg> element (not the wrapper div)
                  const svgEl=document.querySelector('#fs-plan-container svg');
                  if(!svgEl) return;
                  // Add XML declaration + proper SVG namespace for standalone file
                  const svgStr='<?xml version="1.0" encoding="UTF-8"?>\n'
                    + svgEl.outerHTML.replace('<svg ','<svg xmlns="http://www.w3.org/2000/svg" ');
                  const blob=new Blob([svgStr],{type:'image/svg+xml;charset=utf-8'});
                  const a=document.createElement('a');
                  a.href=URL.createObjectURL(blob);
                  a.download='warehouse-floorplan.svg';
                  a.click();
                  URL.revokeObjectURL(a.href);
                }}
                style={{padding:'7px 16px',background:'#7c3aed',color:'#fff',border:'none',
                  borderRadius:'8px',cursor:'pointer',fontFamily:'inherit',
                  fontSize:'13px',fontWeight:'700'}}>
                ⬇ SVG
              </button>
              <button onClick={()=>{
                  const host=document.getElementById('fs-plan-container');
                  downloadPlan2D('png',{current:host},2);
                }}
                style={{padding:'7px 14px',background:'#1d4ed8',color:'#fff',border:'none',
                  borderRadius:'8px',cursor:'pointer',fontFamily:'inherit',
                  fontSize:'13px',fontWeight:'700'}}>
                ⬇ PNG
              </button>
              <button onClick={()=>exportDXF(analysis,userDesign||design,params,
                  userRackConfig||rackConfig)}
                title="AutoCAD DXF in metres"
                style={{padding:'7px 14px',background:'#7e22ce',color:'#fff',border:'none',
                  borderRadius:'8px',cursor:'pointer',fontFamily:'inherit',
                  fontSize:'13px',fontWeight:'700'}}>
                📐 DXF
              </button>
              <button onClick={()=>setFloorPlanFS(false)}
                style={{padding:'7px 16px',background:'#be185d',color:'#fff',border:'none',
                  borderRadius:'8px',cursor:'pointer',fontFamily:'inherit',
                  fontSize:'13px',fontWeight:'700'}}>
                ✕ Close
              </button>
            </div>
          </div>
          {/* Full-screen plan */}
          <div ref={fsScrollRef} style={{flex:1,overflow:'auto',padding:'0',background:'#f8fafc'}}>
            <div id="fs-plan-container">
              <FloorPlanSVG
                analysis={analysis}
                design={userDesign||design}
                params={params}
                rackConfig={userRackConfig||rackConfig}
                fullscreen={true}
                measureOn={measureOn} measurePts={measurePts}
                measurements={measurements} onMeasurePoint={onMeasurePoint} snapOn={snapOn}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
