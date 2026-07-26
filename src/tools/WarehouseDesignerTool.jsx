// ─── WAREHOUSE DESIGNER TOOL ──────────────────────────────────────────────────
// Step 1: Warehouse parameters
// Step 2: Master SKU data (dimensions)
// Step 3: Order / Pick data (for velocity)
// Step 4: Inventory data (current stock)
// Outputs: SKU slotting, rack recommendations, warehouse sizing, SVG floor plan
import { useState, useEffect, useRef } from 'react';
import * as THREE from 'three';
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
    const baysNeeded = cfg.locsPerBay>0 ? Math.ceil(cfg.locs/cfg.locsPerBay) : 0;
    const bayFP  = (cfg.bayW/1000)*(cfg.bayD/1000);
    const aisleM = (cfg.aisleW||3000)/1000;
    // Shared aisle model: each bay owns half the aisle on one side
    const area   = +(baysNeeded*bayFP + baysNeeded*(cfg.bayW/1000)*(aisleM/2)).toFixed(1);
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
  (analysis?.slotted||[]).forEach(r => {
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
      // Use widest standard bay that's a clean multiple of bin width
      // Wider bays = more bins per level = fewer bays = slightly less area
      const stdBayWidths = [1800, 1500, 1200, 900];
      const bayW = stdBayWidths.find(bw => bw % Math.min(bL,bW) === 0) ||
                   stdBayWidths.find(bw => Math.floor(bw/Math.min(bL,bW)) >= 2) || 900;
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
const CROSS_AISLE_W_M = 3.0; // cross aisle width in metres
function computeSectionLayout(totalBays, sectionW, bayHm, colSlot, crossIntervalM) {
  const nCols = Math.max(3, Math.floor(sectionW / colSlot));
  const baysPerCol = totalBays > 0 ? Math.ceil(totalBays / nCols) : 3;
  let y = 0.3; let yStor = 0;
  const cYs = [];
  for(let b = 0; b < baysPerCol; b++){
    y += bayHm; yStor += bayHm;
    if(yStor >= crossIntervalM && b < baysPerCol - 1){
      cYs.push(y); y += CROSS_AISLE_W_M; yStor = 0;
    }
  }
  return { nCols, baysPerCol, height: Math.max(3, y+0.3),
    area: +((y+0.3)*sectionW).toFixed(1), crossYPositions: cYs };
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
  // Warehouse length comes FROM the bay layout, not from pre-calculated areas.
  // This ensures the plan always accurately represents the physical bay count.
  var SVG_W = fullscreen ? 1800 : 960;
  var SVG_H = fullscreen ? Math.round(1800 * (actualWL/wW) * 0.8 + 120) : 720;
  var ML=62, MR=70, MT=50, MB=70;
  var DW=SVG_W-ML-MR, DH=SVG_H-MT-MB;
  var sX=DW/wW, sY=DH/actualWL;
  var X=m=>ML+m*sX, Y=m=>MT+m*sY, W=m=>m*sX, H=m=>m*sY;

  // ── AREA HEIGHTS ────────────────────────────────────────────────────────────
  var recH    = Math.max(4,(receivingArea||0)/wW);
  var disH    = Math.max(4,(dispatchArea||0)/wW);
  var stagingH= Math.max(recH,disH);
  var offH    = Math.max(3,(officeArea||50)/wW);
  var mheH    = mheArea>0 ? Math.max(2,mheArea/wW) : 0;
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

  // Build zone rects (from north going south)
  var zoneRects=[], stagingRects=[], supportRects=[];

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
  var rackBayWidthM={};
  (rackConfig||[]).forEach(cfg=>{
    if(cfg.rack && cfg.bayW && !rackBayWidthM[cfg.rack]){
      rackBayWidthM[cfg.rack]=parseFloat(cfg.bayW)/1000; // mm → m
    }
  });
  var RACK_INFO_2D_LOOKUP={
    shelving:   {depth:1.0},liveStorage:{depth:1.2},selective:{depth:2.2},
    doubleDeep: {depth:4.4},driveIn:    {depth:5.5},cantilever:{depth:2.0},ground:{depth:2.4},
  };
  // Order rack types sensibly: manual pick first (near dispatch), then pallet, then bulk
  var RACK_ORDER=['shelving','liveStorage','selective','doubleDeep','driveIn','cantilever','ground'];
  // Group rackConfig by rack type
  var rackTypeAreas={};
  (rackConfig||[]).forEach(cfg=>{
    rackTypeAreas[cfg.rack]=(rackTypeAreas[cfg.rack]||0)+(cfg.area||0);
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
  RACK_ORDER.forEach(rt => {
    var totalBaysRt = (rackConfig||[]).filter(c=>c.rack===rt).reduce((s,c)=>s+(c.baysNeeded||0),0);
    if(!totalBaysRt && !rackTypeAreas[rt]) return;
    var rtRi    = RACK_INFO_2D_LOOKUP[rt] || {depth:2.2};
    var rtPa    = rt==='shelving'||rt==='liveStorage' ? 1.2 : aisleM;
    var rtSlot  = rtRi.depth + rtPa;
    var rtBayH  = rackBayWidthM[rt] || BAY_HEIGHT_M_LOOKUP[rt] || 0.9;
    var rtCross = ({shelving:13,liveStorage:13,selective:27,
      doubleDeep:27,driveIn:27,cantilever:27,ground:27})[rt] || 13;
    var rtBays  = totalBaysRt || Math.ceil((rackTypeAreas[rt]||0) / (rtBayH * wW));
    var rtLayout= computeSectionLayout(rtBays, wW, rtBayH, rtSlot, rtCross);
    sectionLayouts[rt] = {...rtLayout, totalBays: rtBays};

    var rtStyle = RACK_TYPE_STYLE[rt] || {label:rt,color:'#f8fafc',border:'#e2e8f0',text:'#374151'};
    zoneRects.push({key:rt, x:0, y:cur, w:wW, h:rtLayout.height,
      label:rtStyle.label, color:rtStyle.color, border:rtStyle.border, text:rtStyle.text,
      area:rtLayout.area, rackType:rt, sectionLayout:rtLayout});
    cur += rtLayout.height;
  });
  // Actual warehouse length derived from layout (not from pre-calculated area)
  var layoutWL = cur + stagingH;
  var actualWL = Math.max(wL, layoutWL); // layout-derived warehouse length
  // Adjust scale factor to fit actual layout height
  SVG_H = fullscreen ? Math.round(1800 * (actualWL/wW) * 0.8 + 120) : 720;
  DH = SVG_H - MT - MB;
  sY = DH / actualWL;
  // Rebuild Y and H coordinate transformers with updated sY
  Y = m => MT + m*sY;
  H = m => m*sY;

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
    var eastW2=Math.min(wW*0.3,14);
    stagingRects.push({ key:'receiving', x:0, y:cur, w:wW-eastW2, h:stagingH,
      label:'RECEIVING / GRN', subLabel:`${receivingArea}m² (${sqft(receivingArea)})`,
      color:'#e0f2fe', border:'#0284c7', text:'#0369a1' });
    stagingRects.push({ key:'dispatch', x:wW-eastW, y:cur, w:eastW2, h:stagingH,
      label:'DISPATCH', subLabel:`${dispatchArea}m² (${sqft(dispatchArea)})`,
      color:'#fef3c7', border:'#d97706', text:'#92400e' });
  }

  // Dock doors
  var dockDoors=[];
  var doorW=3.5;
  if (isOne) {
    var sp=wW/(totalDocks+1);
    for(let i=1;i<=totalDocks;i++) dockDoors.push({x:sp*i-doorW/2,y:wL,side:'south',label:`D${i}`});
  } else if (isBoth) {
    var ssp=wW/(inboundDocks+1);
    for(let i=1;i<=inboundDocks;i++) dockDoors.push({x:ssp*i-doorW/2,y:wL,side:'south',label:`D${i}`});
    var nsp=wW/(outboundDocks+1);
    for(let i=1;i<=outboundDocks;i++) dockDoors.push({x:nsp*i-doorW/2,y:0,side:'north',label:`D${inboundDocks+i}`});
  } else {
    var eastW=Math.min(wW*0.3,14);
    var southN=inboundDocks, eastN=outboundDocks;
    var ssp2=(wW-eastW)/(southN+1);
    for(let i=1;i<=southN;i++) dockDoors.push({x:ssp2*i-doorW/2,y:wL,side:'south',label:`D${i}`});
    var esp=wL/(eastN+1);
    for(let i=1;i<=eastN;i++) dockDoors.push({x:wW,y:esp*i,side:'east',label:`D${southN+i}`});
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

  // Build rack section lookup: rackType → total bays from rackConfig
  var rackTypeBays={};
  (rackConfig||[]).forEach(cfg=>{
    rackTypeBays[cfg.rack]=(rackTypeBays[cfg.rack]||0)+(cfg.baysNeeded||0);
  });

  var rackRowsForZone=(zone)=>{
    var rows=[], crossAisles=[];
    var dom=zone.rackType||(zone2RackTypes[zone.key]?.[0]?.rack)||'shelving';
    var ri=RACK_INFO_2D[dom]||RACK_INFO_2D.shelving;
    var pa=dom==='shelving'||dom==='liveStorage'?1.2:aisleM;
    var colSlot=ri.depth+pa;
    var bayHm=rackBayWidthM[dom]||BAY_HEIGHT_M_LOOKUP[dom]||0.9;

    // Use pre-computed layout (bay-first) from zone definition
    var sl=zone.sectionLayout||sectionLayouts[dom]||{nCols:3,baysPerCol:5,crossYPositions:[]};
    var nCols=sl.nCols, totalBays=sl.totalBays||0;
    var crossYs=sl.crossYPositions||[];

    // Push cross aisles from pre-computed positions
    crossYs.forEach(y=>{
      crossAisles.push({x:zone.x,y:zone.y+y-CROSS_AISLE_W_M/2,
        w:zone.w,h:CROSS_AISLE_W_M,isCrossAisle:true});
    });

    // Draw columns using pre-computed nCols and cross aisle positions
    let curX=zone.x+0.3;
    var breakYs=[0,...crossYs,zone.h-0.3];
    for(let i=0;i<nCols;i++){
      var rx=curX+i*colSlot+pa/2;
      if(rx+ri.depth>zone.x+zone.w-0.3) break;
      for(let j=0;j<breakYs.length-1;j++){
        var sy=zone.y+breakYs[j]+(j>0?CROSS_AISLE_W_M/2:0.3);
        var ey=zone.y+breakYs[j+1]-(j<breakYs.length-2?CROSS_AISLE_W_M/2:0);
        if(ey-sy>0.5) rows.push({x:rx,y:sy,w:ri.depth,h:ey-sy,...ri,dom,bayHm});
      }
    }
    return {rows, crossAisles, nCols, baysPerCol:sl.baysPerCol, totalBays};
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
  var allRackRows=[], allCrossAisles=[];
  zoneRects.forEach(zone=>{
    const{rows,crossAisles}=rackRowsForZone(zone);
    allRackRows.push(...rows);
    allCrossAisles.push(...crossAisles);
  });

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
    SVG_W, SVG_H, ML, MR, MT, MB, DW, DH, sX, sY, actualWL, wW, wL: actualWL,
    X: m=>ML+m*sX, Y: m=>MT+m*sY, W: m=>m*sX, H: m=>m*sY,
    dockSide, forkType, packingBenches: params.packingBenches,
    nMHE: design.nMHE||0, inboundMode: params.inboundMode, outboundMode: params.outboundMode,
    stagingH, isBoth, isOne, recH, disH, offH, mheH, supportH,
    zoneRects, stagingRects, supportRects, dockDoors,
    allRackRows, allCrossAisles, recPallets, disPallets,
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

function FloorPlanSVG({ analysis, design, params, rackConfig, fullscreen=false }) {
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
    SVG_W, SVG_H, ML, MR, MT, MB, DW, DH, sX, sY, actualWL, wW, wL,
    X, Y, W, H,
    dockSide, forkType, packingBenches, nMHE, inboundMode, outboundMode,
    stagingH, isBoth, isOne, recH, disH, offH, mheH, supportH,
    zoneRects, stagingRects, supportRects, dockDoors,
    allRackRows, allCrossAisles, recPallets, disPallets,
    packTables, mheBays, dimRight, sectionLayouts, doorW,
    zoneAreas, receivingArea, dispatchArea, mheArea, officeArea, netRackArea,
    totalDocks, inboundDocks, outboundDocks, staging,
  } = fp;
  const MFT=3.2808, M2FT=10.7639;
  const ft = m => `${(m*MFT).toFixed(0)}'`;
  const sqft = m2 => `${Math.round(m2*M2FT).toLocaleString()} sq ft`;
  const aisleM = parseFloat(params.aisleW)||3.0;

  return (
    <svg width={fullscreen?'100%':SVG_W} height={fullscreen?'100%':SVG_H}
      viewBox={`0 0 ${SVG_W} ${SVG_H}`}
      id={fullscreen?'fs-plan-svg':undefined}
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

      {/* ── CROSS AISLES (yellow stripes perpendicular to rack rows) ─── */}
      {allCrossAisles.map((a,i)=>(
        <g key={`ca-${i}`}>
          <rect x={X(a.x)} y={Y(a.y)} width={W(a.w)} height={Math.max(3,H(a.h))}
            fill="#fef9c3" stroke="#ca8a04" strokeWidth="0.5" opacity="0.8"/>
          <text x={X(a.x+a.w/2)} y={Y(a.y+a.h/2)} textAnchor="middle"
            dominantBaseline="middle" fontSize="7" fill="#92400e" fontWeight="700">
            CROSS AISLE
          </text>
        </g>
      ))}

      {/* ── RACK ROWS (top view — type-specific symbols) ─── */}
      {allRackRows.map((r,i)=>{
        const px=X(r.x), py=Y(r.y), pw=W(r.w), ph=Math.max(3,H(r.h));
        const dom=r.dom;

        // ── SELECTIVE PALLET RACK ─────────────────────────────────────────
        // Vertical column: bays stack along column HEIGHT, back-to-back partition at center WIDTH
        if(dom==='selective'||dom==='doubleDeep'){
          const bayH  = H(2.7);  // one bay height in pixels (N-S direction)
          const nBays = Math.max(1,Math.floor(r.h/2.7));
          const halfW = Math.max(2,(pw-3)/2); // width of each pallet face
          return(
            <g key={`rr-${i}`}>
              {/* Column background */}
              <rect x={px} y={py} width={Math.max(3,pw)} height={ph}
                fill={r.color} stroke={r.stroke} strokeWidth="1" rx="1"/>
              {/* Bay rows — each 2.7m tall */}
              {Array.from({length:nBays},(_,b)=>{
                const by=py+b*bayH;
                if(by+bayH>py+ph) return null;
                return(
                  <g key={b}>
                    {/* Front face pallet (left half of column) */}
                    <rect x={px+1} y={by+1} width={halfW} height={bayH-2}
                      fill="#fde68a" stroke="#d97706" strokeWidth="0.6" rx="0.5"/>
                    {/* Back face pallet (right half of column) */}
                    <rect x={px+halfW+2} y={by+1} width={halfW} height={bayH-2}
                      fill="#fde68a" stroke="#d97706" strokeWidth="0.6" rx="0.5"/>
                    {/* Pallet cross marks */}
                    {[px+1+halfW/2, px+halfW+2+halfW/2].map((bx,bi)=>(
                      <g key={bi}>
                        <line x1={bx} y1={by+bayH*0.3} x2={bx} y2={by+bayH*0.7} stroke="#d97706" strokeWidth="0.5"/>
                        <line x1={bx-halfW*0.3} y1={by+bayH/2} x2={bx+halfW*0.3} y2={by+bayH/2} stroke="#d97706" strokeWidth="0.5"/>
                      </g>
                    ))}
                    {/* Bay beam (horizontal line at top of each bay) */}
                    {b>0&&<line x1={px} y1={by} x2={px+pw} y2={by} stroke="#374151" strokeWidth="1.5"/>}
                  </g>
                );
              })}
              {/* ── BACK-TO-BACK PARTITION — vertical centre line through full column height ── */}
              <line x1={px+pw/2} y1={py} x2={px+pw/2} y2={py+ph}
                stroke="#1e293b" strokeWidth="1.8"/>
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

        // ── SHELVING / LIVE STORAGE ────────────────────────────────────────
        {
          // Bay dividers run HORIZONTALLY at regular height intervals (every 0.9m for shelving)
          const bayHpx = H(r.bayHm||0.9); // bay height in pixels
          const nBayDividers = Math.max(0, Math.floor(r.h/(r.bayHm||0.9))-1);
          return(
            <g key={`rr-${i}`}>
              {/* Column fill */}
              <rect x={px} y={py} width={Math.max(3,pw)} height={Math.max(2,ph)}
                fill={r.color} stroke={r.stroke} strokeWidth="0.8" rx="0.5"/>
              {/* Bay dividers — horizontal lines at regular height intervals */}
              {Array.from({length:nBayDividers},(_,b)=>(
                <line key={b}
                  x1={px} y1={py+(b+1)*bayHpx}
                  x2={px+pw} y2={py+(b+1)*bayHpx}
                  stroke={r.stroke} strokeWidth="0.5" strokeOpacity="0.5"/>
              ))}
              {/* ── BACK-TO-BACK PARTITION ── vertical centre line (two shelving faces) */}
              <line
                x1={px+Math.max(3,pw)/2} y1={py}
                x2={px+Math.max(3,pw)/2} y2={py+ph}
                stroke={r.stroke} strokeWidth="1.4" strokeOpacity="0.9"/>
            </g>
          );
        }
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
        if(ph<14) return null;
        return(
          <g key={`zl-${z.key}`}>
            {/* Section name — left-aligned, not in center so it doesn't overlap racks */}
            <text x={X(z.x)+6} y={py+Math.min(ph/2,20)}
              dominantBaseline="middle"
              fontSize={Math.min(11, ph*sY*0.3)} fontWeight="700" fill={z.text}
              opacity="0.85">
              {z.label}
            </text>
            {ph>30&&<text x={X(z.x)+6} y={py+Math.min(ph/2,20)+13}
              dominantBaseline="middle"
              fontSize="8" fontWeight="400" fill={z.text} opacity="0.7">
              {(z.area||0).toFixed(0)}m²
            </text>}
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

      {/* ── TOTAL AREA FOOTER ─── */}
      <text x={X(wW/2)} y={SVG_H-4} textAnchor="middle" fontSize="10" fontWeight="700" fill="#374151">
        {`Total gross area: ${(wW*actualWL).toLocaleString()}m²  (${Math.round(wW*actualWL*10.7639).toLocaleString()} sq ft)  ·  ${wW}×${Math.round(actualWL)}m  ·  ${dockSide==='one'?'One-side':'Opposite-side'} docks  ·  Derived from ${Object.keys(sectionLayouts).length} rack type sections`}
      </text>
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
function Warehouse3DModel({ analysis, design, params, rackConfig }) {
  const mountRef    = useRef(null);
  const cleanupRef  = useRef(null);

  useEffect(() => {
    if (!mountRef.current || !design) return;
    const {
      wW, wL, zoneAreas={}, receivingArea=80, dispatchArea=80,
      mheArea=0, officeArea=50, totalDocks=4, inboundDocks=2, outboundDocks=2,
    } = design;
    const { dockSide, aisleW:aisleWP, clearH:clearHP, dockPitch } = params;
    const clearH = parseFloat(clearHP)||9;
    const aisleM = parseFloat(aisleWP)||3.0;
    const pitch  = parseFloat(dockPitch)||4.5;

    const container = mountRef.current;
    const W=container.clientWidth, H=Math.max(480,container.clientHeight||480);

    // ── SCENE ──────────────────────────────────────────────────────────────
    const scene    = new THREE.Scene();
    scene.background = new THREE.Color(0xdbeafe);
    scene.fog = new THREE.FogExp2(0xdbeafe, 0.006);

    const renderer = new THREE.WebGLRenderer({ antialias:true });
    renderer.setPixelRatio(window.devicePixelRatio||1);
    renderer.setSize(W, H);
    renderer.shadowMap.enabled = true;
    renderer.shadowMap.type = THREE.PCFSoftShadowMap;
    container.appendChild(renderer.domElement);

    const camera = new THREE.PerspectiveCamera(50, W/H, 0.1, 2000);
    const center = new THREE.Vector3(wW/2, clearH*0.3, wL/2);
    const diag   = Math.sqrt(wW*wW + wL*wL);

    // Orbit state
    const orbit = { theta:Math.PI*0.55, phi:1.05, radius:diag*1.1, down:false, lx:0, ly:0 };

    const updateCam = () => {
      const ph=Math.max(0.18,Math.min(1.45,orbit.phi));
      camera.position.set(
        center.x + orbit.radius*Math.sin(ph)*Math.sin(orbit.theta),
        center.y + orbit.radius*Math.cos(ph),
        center.z + orbit.radius*Math.sin(ph)*Math.cos(orbit.theta)
      );
      camera.lookAt(center);
    };
    updateCam();

    // ── LIGHTING ───────────────────────────────────────────────────────────
    scene.add(new THREE.AmbientLight(0xffffff, 0.55));
    const sun = new THREE.DirectionalLight(0xfffaed, 1.0);
    sun.position.set(wW*1.5, clearH*4, -wL*0.3);
    sun.castShadow = true;
    sun.shadow.mapSize.set(2048,2048);
    sun.shadow.camera.left=-diag; sun.shadow.camera.right=diag;
    sun.shadow.camera.top=diag;   sun.shadow.camera.bottom=-diag;
    scene.add(sun);
    const fill = new THREE.DirectionalLight(0xadd8e6, 0.35);
    fill.position.set(-wW, clearH*2, wL*1.5);
    scene.add(fill);

    // ── HELPERS ────────────────────────────────────────────────────────────
    const addMesh = (geo, mat, x, y, z, castShadow=true) => {
      const m = new THREE.Mesh(geo, mat);
      m.position.set(x, y, z);
      if (castShadow) { m.castShadow=true; m.receiveShadow=true; }
      scene.add(m); return m;
    };
    const box = (w,h,d,col,opacity=1,wireframe=false) => new THREE.Mesh(
      new THREE.BoxGeometry(w,h,d),
      new THREE.MeshPhongMaterial({color:col, opacity, transparent:opacity<1,
        wireframe, side:THREE.DoubleSide})
    );
    const addBox = (x,y,z,w,h,d,col,op=1,shadow=true) => {
      const m = box(w,h,d,col,op);
      m.position.set(x+w/2, y+h/2, z+d/2);
      if(shadow){m.castShadow=true;m.receiveShadow=true;}
      scene.add(m); return m;
    };
    const edges = (geo, col=0x000000, op=0.25) => new THREE.LineSegments(
      new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({color:col,opacity:op,transparent:true})
    );

    // ── FLOOR ──────────────────────────────────────────────────────────────
    const floorGeo = new THREE.PlaneGeometry(wW, wL);
    const floorM   = addMesh(floorGeo, new THREE.MeshPhongMaterial({color:0xcfd8e3,side:THREE.DoubleSide}),
      wW/2, 0, wL/2, false);
    floorM.rotation.x=-Math.PI/2; floorM.receiveShadow=true;

    const grid = new THREE.GridHelper(Math.max(wW,wL)*2, 30, 0xaaaaaa, 0xcccccc);
    grid.position.set(wW/2, 0.01, wL/2); scene.add(grid);

    // ── ZONE FLOORS ────────────────────────────────────────────────────────
    const ZCOL={golden:0xa7f3d0, mid:0xfef08a, reserve:0xfed7aa, bulk:0xe2e8f0, long:0xede9fe};
    const stagingH = Math.max(4,(receivingArea||80)/wW);
    const supportH = Math.max(2,((officeArea||50)+(mheArea||0))/wW);
    const availH   = Math.max(4,wL-stagingH-supportH);
    const ZORD=['golden','mid','reserve','bulk','long'];
    const RD_MAP={shelving:0.6,liveStorage:1.5,selective:1.1,driveIn:6.6,doubleDeep:2.4,cantilever:2.5};

    // ── Build rack-type → zone mapping from rackConfig ──────────────────────
    // This is the key fix: derive zone assignment from rackConfig, not just slotted data
    const RACK_DEFAULT_ZONE={
      shelving:'golden', liveStorage:'golden',
      selective:'reserve', doubleDeep:'reserve',
      driveIn:'bulk', cantilever:'long', ground:'bulk'
    };
    const zoneRackTypes={}; // zone → [{rack, cfg}]
    const zoneRackH={};     // zone → max rack height

    (rackConfig||[]).forEach(cfg=>{
      // Find which zone this rack type appears in via slotted data
      const slottedZone=(analysis?.slotted||[]).find(s=>s.rack===cfg.rack)?.zone
        || RACK_DEFAULT_ZONE[cfg.rack]||'golden';
      if(!zoneRackTypes[slottedZone]) zoneRackTypes[slottedZone]=[];
      if(!zoneRackTypes[slottedZone].find(r=>r.rack===cfg.rack))
        zoneRackTypes[slottedZone].push({rack:cfg.rack, cfg});
      // Track rack height per zone
      const rh=['shelving','liveStorage'].includes(cfg.rack)
        ? (parseFloat(cfg.tierHeight)||cfg.shelfH||2200)*(parseInt(cfg.tiers)||1)/1000
        : (cfg.levels||4)*1.5+0.3;
      zoneRackH[slottedZone]=Math.max(zoneRackH[slottedZone]||0, rh);
    });

    // Fill zones with no rackConfig using slotted dominant
    const getDom=(zone)=>{
      if(zoneRackTypes[zone]?.length) return zoneRackTypes[zone][0].rack;
      const m={};
      (analysis?.slotted||[]).filter(s=>s.zone===zone).forEach(r=>{m[r.rack]=(m[r.rack]||0)+1;});
      return Object.entries(m).sort((a,b)=>b[1]-a[1])[0]?.[0]||'shelving';
    };

    // Compute zone heights — two passes:
    // Pass 1: minimum height per rack type (each type needs at least 1 full row slot)
    const MIN_ROWS=1; // show at least 1 rack row per configured type
    const minZoneH={};
    ZORD.forEach(z=>{
      const racks=zoneRackTypes[z]||[];
      if(!racks.length){ minZoneH[z]=0; return; }
      // Each rack type needs its own sub-zone with >=1 row
      minZoneH[z]=racks.reduce((sum,{rack:dom})=>{
        const rd=RD_MAP[dom]||0.6;
        const slot=rd+(parseFloat(aisleWP)||3.0);
        return sum + slot*MIN_ROWS;
      },0);
    });

    // Pass 2: allocate proportionally from zoneAreas, but ALWAYS enforce minimums,
    //         then RESCALE everything to fit within availH (so zones never go out of bounds)
    const rawTotZA=Object.values(zoneAreas).reduce((s,a)=>s+a,0)||1;
    const totalMin=ZORD.reduce((s,z)=>s+minZoneH[z],0);
    const spare=Math.max(0, availH-totalMin);

    const rawH={};
    ZORD.forEach(z=>{
      rawH[z]=Math.max(minZoneH[z], ((zoneAreas[z]||0)/rawTotZA)*spare + minZoneH[z]);
    });
    // Rescale so total fits in availH (prevents zones rendering outside warehouse)
    const rawTotal=Object.values(rawH).reduce((s,v)=>s+v,0)||availH;
    const scale=availH/rawTotal;

    let zCur=stagingH; const ZP={};
    ZORD.forEach(z=>{
      const h=(rawH[z]||0)*scale;
      ZP[z]={z0:zCur, h:Math.max(0,h)};
      zCur+=h;
    });

    const addFloorZone = (x,z,w,d,col) => {
      const m=new THREE.Mesh(new THREE.PlaneGeometry(w,d), new THREE.MeshPhongMaterial({color:col,side:THREE.DoubleSide}));
      m.rotation.x=-Math.PI/2; m.position.set(x+w/2,0.015,z+d/2); m.receiveShadow=true; scene.add(m);
    };
    ZORD.forEach(z=>{const{z0,h}=ZP[z];if(h>0.3) addFloorZone(0,z0,wW,h,ZCOL[z]||0xf1f5f9);});
    addFloorZone(0,    0,    wW/2, stagingH, 0x93c5fd); // Receiving
    addFloorZone(wW/2, 0,    wW/2, stagingH, 0xfde68a); // Dispatch
    addFloorZone(0, wL-supportH, wW/2, supportH, 0xdbeafe); // Office
    if(mheArea>0) addFloorZone(wW/2, wL-supportH, wW/2, supportH, 0xede9fe); // MHE

    // Zone boundary lines
    ZORD.forEach(z=>{
      const{z0,h}=ZP[z];if(h<0.5) return;
      const pts=[new THREE.Vector3(0,0.05,z0),new THREE.Vector3(wW,0.05,z0)];
      const ln=new THREE.Line(new THREE.BufferGeometry().setFromPoints(pts),
        new THREE.LineBasicMaterial({color:0x94a3b8,opacity:0.5,transparent:true}));
      scene.add(ln);
    });

    // ── WALLS (semi-transparent) ────────────────────────────────────────────
    const wallMat = new THREE.MeshPhongMaterial({color:0xf1f5f9,opacity:0.22,transparent:true,side:THREE.DoubleSide});
    const wallEdgeMat = new THREE.LineBasicMaterial({color:0x64748b,opacity:0.5,transparent:true});
    [[wW/2,clearH/2,0,       wW,clearH,0.25],  // south  (z=0)
     [wW/2,clearH/2,wL,      wW,clearH,0.25],  // north
     [0,   clearH/2,wL/2,    0.25,clearH,wL],  // west
     [wW,  clearH/2,wL/2,    0.25,clearH,wL],  // east
    ].forEach(([cx,cy,cz,ww,hh,dd])=>{
      const geo=new THREE.BoxGeometry(ww,hh,dd);
      const m=new THREE.Mesh(geo,wallMat); m.position.set(cx,cy,cz); scene.add(m);
      const e=edges(geo,0x64748b,0.4); e.position.copy(m.position); scene.add(e);
    });

    // Roof wireframe
    [[wW/2,clearH,0],  [wW/2,clearH,wL],
     [0,clearH,wL/2],  [wW,clearH,wL/2]].forEach(([cx,cy,cz])=>{
      const pts=[new THREE.Vector3(cx-wW/2,cy,cz-wL/2),new THREE.Vector3(cx+wW/2,cy,cz-wL/2),
                 new THREE.Vector3(cx+wW/2,cy,cz+wL/2),new THREE.Vector3(cx-wW/2,cy,cz+wL/2),
                 new THREE.Vector3(cx-wW/2,cy,cz-wL/2)];
      scene.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints(pts),
        new THREE.LineBasicMaterial({color:0x94a3b8,opacity:0.4,transparent:true})));
    });

    // ── DOCK DOORS ─────────────────────────────────────────────────────────
    const dockW=3.5, dockH=Math.min(4.5,clearH*0.55);
    const doorMat=new THREE.MeshPhongMaterial({color:0x1d4ed8});
    for(let d=0;d<totalDocks;d++){
      const dx=(d+0.5)*(wW/totalDocks)-dockW/2;
      [[dx,0,-0.15,0.2,dockH,0.3],[dx+dockW-0.2,0,-0.15,0.2,dockH,0.3],  // posts
       [dx,dockH,-0.15,dockW,0.2,0.3]].forEach(([x,y,z,w,h,dd])=>{      // lintel
        const m=new THREE.Mesh(new THREE.BoxGeometry(w,h,dd),doorMat);
        m.position.set(x+w/2,y+h/2,z); scene.add(m);
      });
      // Door number
    }

    // ── RACK ROWS — type-specific 3D geometry ──────────────────────────────
    const RCOL={shelving:0x94a3b8,liveStorage:0x3b82f6,selective:0x475569,
      driveIn:0x1e293b,doubleDeep:0x334155,cantilever:0x7c3aed};
    const RD3={shelving:0.6,liveStorage:1.5,selective:1.1,driveIn:6.6,doubleDeep:2.4,cantilever:2.5,ground:1.2};

    const matCache={};
    const getMat=(col,op=1,sh=30)=>{
      const k=`${col}|${op}`;
      if(!matCache[k]) matCache[k]=new THREE.MeshPhongMaterial(
        {color:col,opacity:op,transparent:op<1,shininess:sh,side:THREE.DoubleSide});
      return matCache[k];
    };
    const mkBox=(w,h,d,mat,x,y,z,cast=true)=>{
      const m=new THREE.Mesh(new THREE.BoxGeometry(w,h,d),mat);
      m.position.set(x,y,z); if(cast){m.castShadow=true;m.receiveShadow=true;}
      scene.add(m); return m;
    };

    // Warehouse storage south/north limits (racks must stay inside)
    const storeSouth = stagingH + 0.3;
    const storeNorth = wL - supportH - 0.3;

    ZORD.forEach(zone=>{
      const{z0,h}=ZP[zone]||{z0:0,h:0}; if(h<0.3) return;

      // Rack types assigned to this zone
      const zRacks=zoneRackTypes[zone]||[];
      if(!zRacks.length){
        const m={};
        (analysis?.slotted||[]).filter(s=>s.zone===zone)
          .forEach(r=>{m[r.rack]=(m[r.rack]||0)+1;});
        const dom=Object.entries(m).sort((a,b)=>b[1]-a[1])[0]?.[0]||'shelving';
        zRacks.push({rack:dom,cfg:null});
      }

      // Divide zone into non-overlapping sub-zones per rack type
      const totalSlotW=zRacks.reduce((s,{rack:dom})=>s+(RD3[dom]||0.6)+aisleM,0)||1;
      let subZ=z0;

      zRacks.forEach(({rack:dom,cfg:zc})=>{
        const rd=RD3[dom]||0.6;
        const slot=rd+aisleM;
        const subH=Math.max(slot,(slot/totalSlotW)*h);

        // CAP rack height to clearH
        const cfgRh=zc
          ?(['shelving','liveStorage'].includes(dom)
              ?(parseFloat(zc.tierHeight)||zc.shelfH||2200)*(parseInt(zc.tiers)||1)/1000
              :(zc.levels||4)*1.5+0.3)
          :(zoneRackH[zone]||3.5);
        const rh=Math.min(clearH-0.25, cfgRh);

        // Cross-aisle: leave a gap after half the rows (standard warehouse practice)
        const maxRows=Math.max(1,Math.floor(subH/slot));
        const crossAisleAfter=Math.floor(maxRows/2); // insert cross-aisle after this row

        let rowIdx=0;
        for(;;){
          const rowZ=subZ+rowIdx*slot+(rowIdx>crossAisleAfter?aisleM:0)+aisleM*0.3;
          // STRICT boundary: never render outside warehouse or sub-zone
          if(rowZ+rd>Math.min(subZ+subH,storeNorth)-0.05) break;
          if(rowIdx>=maxRows+1) break;

          // Skip the cross-aisle slot itself (draw floor marker instead)
          if(rowIdx===crossAisleAfter&&maxRows>2){
            // Yellow cross-aisle floor stripe
            const ca=new THREE.Mesh(
              new THREE.PlaneGeometry(wW,aisleM*0.9),
              new THREE.MeshPhongMaterial({color:0xfef08a,side:THREE.DoubleSide}));
            ca.rotation.x=-Math.PI/2;
            ca.position.set(wW/2,0.03,rowZ-rd+rd/2+aisleM/2);
            scene.add(ca);
            rowIdx++; continue;
          }

          // ── SELECTIVE PALLET RACK ────────────────────────────────────
          if(dom==='selective'||dom==='doubleDeep'){
            const bayW=2.7, nBays=Math.max(1,Math.floor((wW-0.4)/bayW));
            const nLvl=Math.min(6,Math.floor(rh/1.5));
            const depth=dom==='doubleDeep'?2:1;
            const upMat=getMat(0xdc2626,1,40);
            const bmMat=getMat(0xf59e0b,1,80);
            const pMat =getMat(0xfbbf24,0.9,20);
            const pbMat=getMat(0x78350f,1,10);

            // Upright frames at bay boundaries
            for(let b=0;b<=nBays;b++){
              const fx=Math.min(b*bayW, wW-0.15); // clamp to warehouse width
              [[fx,rowZ+0.06],[fx,rowZ+rd-0.06]].forEach(([px,pz])=>{
                mkBox(0.12,rh,0.12,upMat,px,rh/2,pz);
              });
            }
            // Beams + pallets per level
            for(let lv=0;lv<nLvl;lv++){
              const by=lv*1.5+0.3;
              for(let b=0;b<nBays;b++){
                const bx=b*bayW;
                [rowZ+0.08, rowZ+rd-0.08].forEach(bz=>{
                  mkBox(bayW,0.10,0.09,bmMat,bx+bayW/2,by,bz,false);
                });
                // FIX: pallet offsets within bay (not cumulative x)
                for(let d=0;d<depth;d++){
                  [0.30, bayW-1.40].forEach(palOff=>{
                    mkBox(1.1,0.14,1.0,pbMat, bx+palOff, by+0.07, rowZ+0.12+d*1.15);
                    mkBox(1.05,0.92,0.95,pMat, bx+palOff, by+0.07+0.14+0.46, rowZ+0.12+d*1.15);
                  });
                }
              }
            }
          }

          // ── DRIVE-IN RACK ─────────────────────────────────────────────
          else if(dom==='driveIn'){
            const laneW=2.7, nLanes=Math.max(1,Math.floor((wW-0.4)/laneW));
            const nLvl=Math.min(4,Math.floor(rh/1.5));
            const palDeep=Math.max(2,Math.round(rd/1.2));
            const upMat =getMat(0xdc2626,1,20);
            const railMat=getMat(0xf59e0b,1,60);
            const pMat  =getMat(0xfbbf24,0.9,20);
            const pbMat =getMat(0x78350f,1,10);

            // Column frames at each lane boundary (front+mid+rear)
            for(let ln=0;ln<=nLanes;ln++){
              const fx=Math.min(ln*laneW, wW-0.15);
              [rowZ+0.1, rowZ+rd/2, rowZ+rd-0.1].forEach(pz=>{
                mkBox(0.15,rh,0.15,upMat,fx,rh/2,pz);
              });
            }
            // Side rails + deep pallets per level
            for(let lv=0;lv<nLvl;lv++){
              const ry=lv*1.5+0.9;
              const palH_each=rd/palDeep;
              for(let ln=0;ln<nLanes;ln++){
                const lx=ln*laneW;
                // Guide rails both sides
                mkBox(0.07,0.07,rd-0.2,railMat, lx+0.14,ry, rowZ+rd/2, false);
                mkBox(0.07,0.07,rd-0.2,railMat, lx+laneW-0.14,ry, rowZ+rd/2, false);
                // Pallets stored deep in lane
                for(let d=0;d<palDeep;d++){
                  const pz=rowZ+0.12+d*palH_each+palH_each/2;
                  mkBox(laneW-0.35,0.13,palH_each-0.08,pbMat, lx+laneW/2,ry+0.065,pz);
                  mkBox(laneW-0.45,0.9, palH_each-0.12,pMat,  lx+laneW/2,ry+0.065+0.13+0.45,pz);
                }
              }
            }
          }

          // ── GROUND LOCATION ──────────────────────────────────────────────
          else if(dom==='ground'){
            const stackLayers=zc?.stackH||zc?.levels||1;
            const itemH=(zc?.binDims?.[2]||300)/1000; // bin height in metres
            const gW=2.0, nCols=Math.max(1,Math.floor((wW-0.4)/gW));
            const gndMat=new THREE.MeshPhongMaterial({color:0x92400e,opacity:0.3,transparent:true});
            // Ground footprint base
            const gGeo=new THREE.BoxGeometry(wW-0.4,0.05,rd);
            const gM=new THREE.Mesh(gGeo,gndMat);
            gM.position.set(wW/2,0.025,rowZ+rd/2); scene.add(gM);
            // Each ground pad column — stack layers shown as boxes
            const itemMat=getMat(0x78350f,0.9,10);
            const frameMat=new THREE.LineBasicMaterial({color:0xfbbf24});
            for(let col=0;col<nCols;col++){
              const cx=0.2+col*gW;
              for(let lyr=0;lyr<stackLayers;lyr++){
                const by=lyr*(itemH+0.03)+0.04;
                const pad=new THREE.Mesh(new THREE.BoxGeometry(gW-0.15,itemH,rd-0.15),itemMat);
                pad.position.set(cx+gW/2,by+itemH/2,rowZ+rd/2);
                pad.castShadow=true; scene.add(pad);
                // Yellow boundary frame per layer
                const frame=new THREE.LineSegments(
                  new THREE.EdgesGeometry(new THREE.BoxGeometry(gW-0.1,itemH+0.02,rd-0.1)),
                  frameMat);
                frame.position.copy(pad.position); scene.add(frame);
              }
            }
          }

          // ── CANTILEVER RACK ───────────────────────────────────────────
          else if(dom==='cantilever'){
            const spineSpacing=1.5, nSpines=Math.max(1,Math.floor((wW-0.4)/spineSpacing));
            const nArms=Math.min(6,Math.floor(rh/0.7));
            const armLen=(rd-0.2)/2;
            const spMat =getMat(0x4c1d95,1,40);
            const armMat=getMat(0x7c3aed,1,60);
            const itMat =getMat(0xfb923c,0.85,20);
            for(let sp=0;sp<nSpines;sp++){
              const sx=Math.min(0.2+sp*spineSpacing, wW-0.2);
              mkBox(0.2,rh,0.2,spMat,sx,rh/2,rowZ+rd/2);
              mkBox(0.6,0.08,rd,getMat(0x4c1d95),sx,0.04,rowZ+rd/2);
              for(let a=0;a<nArms;a++){
                const ay=a*(rh/nArms)+0.4;
                mkBox(0.08,0.08,armLen,armMat,sx,ay,rowZ+rd/2-armLen/2-0.1,false);
                mkBox(0.08,0.08,armLen,armMat,sx,ay,rowZ+rd/2+armLen/2+0.1,false);
                if(sp<nSpines-1){
                  mkBox(spineSpacing,0.18,armLen*0.7,itMat,sx+spineSpacing/2,ay+0.09,rowZ+rd/2-armLen*0.35);
                  mkBox(spineSpacing,0.18,armLen*0.7,itMat,sx+spineSpacing/2,ay+0.09,rowZ+rd/2+armLen*0.35);
                }
              }
            }
          }

          // ── SHELVING / LIVE STORAGE ───────────────────────────────────
          else {
            const rGeo=new THREE.BoxGeometry(wW-0.4,rh,rd);
            const rMat2=new THREE.MeshPhongMaterial(
              {color:RCOL[dom]||0x94a3b8,opacity:0.65,transparent:true,shininess:20});
            const rM=new THREE.Mesh(rGeo,rMat2);
            rM.position.set(wW/2,rh/2,rowZ+rd/2);
            rM.castShadow=true; rM.receiveShadow=true; scene.add(rM);
            // Edge wireframe — must use position.copy(), not Object.assign
            const eL=new THREE.LineSegments(
              new THREE.EdgesGeometry(rGeo),
              new THREE.LineBasicMaterial({color:0x000000,opacity:0.18,transparent:true}));
            eL.position.copy(rM.position);
            scene.add(eL);
            // Uprights
            const ps=1.8,np=Math.floor((wW-0.4)/ps)+1,pMat=getMat(0x334155);
            for(let p=0;p<np;p++){
              const px=0.2+p*ps; if(px>wW-0.2) break;
              [[px,rowZ+0.05],[px,rowZ+rd-0.05]].forEach(([x,z])=>
                mkBox(0.08,rh,0.08,pMat,x,rh/2,z,false));
            }
            // Shelves
            const ns=Math.min(8,Math.floor(rh/0.35)),shM=getMat(0xd1d5db);
            for(let s=0;s<ns;s++)
              mkBox(wW-0.5,0.04,rd-0.06,shM,wW/2,s*(rh/ns)+0.05,rowZ+rd/2,false);
          }

          rowIdx++;
        } // end row loop

        subZ+=subH; // advance sub-zone cursor — prevents overlap
      }); // end zoneRacks.forEach
    }); // end ZORD.forEach

    // ── STAGING PALLETS ─────────────────────────────────────────────────────
    const palW=1.0, palH=0.15, stockH=0.75;
    const palMat  =new THREE.MeshPhongMaterial({color:0x78350f,shininess:20});
    const stockColors=[0xd97706,0xf59e0b];
    const pRows=Math.min(3,Math.floor(stagingH/1.4));
    const pCols=Math.min(10,Math.floor(wW/2/1.4));
    [0,1].forEach(side=>{
      const startX=side*(wW/2);
      for(let r=0;r<pRows;r++) for(let c=0;c<pCols;c++){
        const px=startX+0.3+c*1.35, pz=0.3+r*1.35;
        if(px+palW>startX+wW/2-0.2||pz>stagingH-0.4) continue;
        // Pallet base
        const pm=new THREE.Mesh(new THREE.BoxGeometry(palW,palH,palW),palMat);
        pm.position.set(px+palW/2,palH/2,pz+palW/2); pm.castShadow=true; scene.add(pm);
        // Stock
        const sm=new THREE.Mesh(new THREE.BoxGeometry(palW-0.05,stockH,palW-0.05),
          new THREE.MeshPhongMaterial({color:stockColors[side],shininess:40}));
        sm.position.set(px+palW/2,palH+stockH/2,pz+palW/2); sm.castShadow=true; scene.add(sm);
      }
    });

    // ── OFFICE / MHE ───────────────────────────────────────────────────────
    if(officeArea>0){
      const om=addBox(0.5,0,wL-supportH+0.5,wW/2-1,Math.min(3,clearH*0.35),supportH-1,0x93c5fd,0.8);
    }
    if(mheArea>0){
      const nMHE=Math.min(design.nMHE||2,4);
      for(let m=0;m<nMHE;m++){
        const mx=wW/2+0.5+m*3.5; if(mx+2.5>wW-0.5) break;
        addBox(mx,0,wL-supportH+0.5,2.5,1.8,1.5,0x818cf8,0.9);
        addBox(mx+0.2,1.8,wL-supportH+0.6,0.3,0.5,0.3,0x312e81);
      }
    }

    // ── COLUMN LABELS via canvas textures ─────────────────────────────────
    const makeLabel=(text,col='#1e293b')=>{
      const cv=document.createElement('canvas'); cv.width=256; cv.height=64;
      const ctx=cv.getContext('2d'); ctx.fillStyle='rgba(255,255,255,0.85)';
      ctx.roundRect(2,2,252,60,8); ctx.fill();
      ctx.fillStyle=col; ctx.font='bold 22px sans-serif'; ctx.textAlign='center';
      ctx.fillText(text,128,38);
      const tex=new THREE.CanvasTexture(cv);
      const m=new THREE.Mesh(new THREE.PlaneGeometry(4,1),new THREE.MeshBasicMaterial({map:tex,transparent:true,side:THREE.DoubleSide}));
      return m;
    };
    const ZONE_LABEL_COL={golden:'#166534',mid:'#854d0e',reserve:'#9a3412',bulk:'#374151',long:'#6b21a8'};
    ZORD.forEach(z=>{
      const{z0,h}=ZP[z]; if(h<2) return;
      const lbl=makeLabel(ZONE_DEFS[z]?.label||z, ZONE_LABEL_COL[z]||'#374151');
      lbl.position.set(wW/2, 0.5, z0+h/2); lbl.rotation.x=-Math.PI/2; scene.add(lbl);
    });
    const recLbl=makeLabel('RECEIVING','#1d4ed8');
    recLbl.position.set(wW/4,0.5,stagingH/2); recLbl.rotation.x=-Math.PI/2; scene.add(recLbl);
    const disLbl=makeLabel('DISPATCH','#d97706');
    disLbl.position.set(wW*3/4,0.5,stagingH/2); disLbl.rotation.x=-Math.PI/2; scene.add(disLbl);

    // ── ANIMATE ────────────────────────────────────────────────────────────
    let rafId; const animate=()=>{ rafId=requestAnimationFrame(animate); renderer.render(scene,camera); };
    animate();

    // ── CONTROLS ───────────────────────────────────────────────────────────
    const el=renderer.domElement;
    const down=e=>{orbit.down=true;orbit.lx=e.clientX;orbit.ly=e.clientY;};
    const up  =()=>{orbit.down=false;};
    const move=e=>{
      if(!orbit.down) return;
      orbit.theta-=(e.clientX-orbit.lx)*0.007;
      orbit.phi  -=(e.clientY-orbit.ly)*0.007;
      orbit.lx=e.clientX; orbit.ly=e.clientY; updateCam();
    };
    const wheel=e=>{ orbit.radius=Math.max(8,Math.min(400,orbit.radius+e.deltaY*0.08)); updateCam(); e.preventDefault(); };
    let lt0=0,lt1=0;
    const tstart=e=>{lt0=e.touches[0].clientX;lt1=e.touches[0].clientY;};
    const tmove =e=>{
      orbit.theta-=(e.touches[0].clientX-lt0)*0.007;
      orbit.phi  -=(e.touches[0].clientY-lt1)*0.007;
      lt0=e.touches[0].clientX;lt1=e.touches[0].clientY; updateCam(); e.preventDefault();
    };
    el.addEventListener('mousedown',down); el.addEventListener('mouseup',up);
    el.addEventListener('mouseleave',up);  el.addEventListener('mousemove',move);
    el.addEventListener('wheel',wheel,{passive:false});
    el.addEventListener('touchstart',tstart,{passive:true}); el.addEventListener('touchmove',tmove,{passive:false});
    const onResize=()=>{ const w=container.clientWidth,h=480; renderer.setSize(w,h); camera.aspect=w/h; camera.updateProjectionMatrix(); };
    window.addEventListener('resize',onResize);

    cleanupRef.current=()=>{
      cancelAnimationFrame(rafId);
      ['mousedown','mouseup','mouseleave','mousemove'].forEach(ev=>el.removeEventListener(ev,ev==='mousedown'?down:ev==='mousemove'?move:up));
      el.removeEventListener('wheel',wheel);
      el.removeEventListener('touchstart',tstart); el.removeEventListener('touchmove',tmove);
      window.removeEventListener('resize',onResize);
      renderer.dispose();
      if(container.contains(el)) container.removeChild(el);
    };
    return cleanupRef.current;
  }, [design, analysis, params, rackConfig]);

  if (!design) return null;
  return (
    <div style={{position:'relative'}}>
      <div ref={mountRef} style={{width:'100%',height:'480px',borderRadius:'10px',
        overflow:'hidden',cursor:'grab',background:'#0f172a'}}/>
      {/* Controls hint */}
      <div style={{position:'absolute',top:'10px',right:'10px',
        background:'rgba(255,255,255,0.88)',backdropFilter:'blur(4px)',
        borderRadius:'8px',padding:'7px 12px',fontSize:'11px',color:'#374151',
        lineHeight:'1.8',boxShadow:'0 2px 8px rgba(0,0,0,0.12)'}}>
        🖱 <strong>Drag</strong> to rotate<br/>
        ⚲ <strong>Scroll</strong> to zoom<br/>
        📱 <strong>Touch drag</strong> on mobile
      </div>
      {/* Legend */}
      <div style={{position:'absolute',bottom:'10px',left:'10px',
        background:'rgba(255,255,255,0.88)',backdropFilter:'blur(4px)',
        borderRadius:'8px',padding:'7px 12px',fontSize:'10px',
        display:'flex',flexWrap:'wrap',gap:'8px',maxWidth:'420px',
        boxShadow:'0 2px 8px rgba(0,0,0,0.12)'}}>
        {[['#94a3b8','Shelving'],['#475569','Pallet Rack'],['#1e293b','Drive-in'],
          ['#7c3aed','Cantilever'],['#93c5fd','Receiving'],['#fde68a','Dispatch'],['#818cf8','MHE']
        ].map(([col,lbl])=>(
          <span key={lbl} style={{display:'inline-flex',alignItems:'center',gap:'4px'}}>
            <span style={{width:'10px',height:'10px',borderRadius:'2px',
              background:col,display:'inline-block',flexShrink:0}}/>
            <span style={{color:'#374151'}}>{lbl}</span>
          </span>
        ))}
      </div>
    </div>
  );
}

// ─── DXF FLOOR PLAN EXPORT ────────────────────────────────────────────────────
// Generates AutoCAD DXF (Drawing Exchange Format).
// Open in AutoCAD → File → Open → .dxf → then Save As → .dwg
// Supported by all CAD tools: AutoCAD, FreeCAD, LibreCAD, Rhino, SolidWorks.
function exportDXF(analysis, design, params, rackConfig) {
  const { wW, wL, zoneAreas={}, receivingArea=80, dispatchArea=80,
    officeArea=50, mheArea=0, totalDocks=4, totalGrossArea=0 } = design;
  const { aisleW:aisleWP } = params;
  const aisleM = parseFloat(aisleWP)||3.0;
  const MFT    = 3.2808;

  // DXF is a list of (group-code, value) pairs — one per line
  const L = [];
  const d = (...pairs) => pairs.forEach(p => L.push(String(p)));
  const n = v => v.toFixed(4);

  // ── HEADER ─────────────────────────────────────────────────────────────
  d('0','SECTION','2','HEADER');
  d('9','$ACADVER','1','AC1015');      // AutoCAD 2000 compatible
  d('9','$INSUNITS','70','6');         // 6 = metres
  d('9','$LUNITS','70','2');           // decimal
  d('9','$LUPREC','70','3');           // 3 decimal places
  d('9','$MEASUREMENT','70','1');      // 1 = metric
  d('9','$EXTMIN','10',n(0),'20',n(0),'30','0.0');
  d('9','$EXTMAX','10',n(wW+6),'20',n(wL+6),'30','0.0');
  d('0','ENDSEC');

  // ── LAYER TABLE ─────────────────────────────────────────────────────────
  const LAYERS = [
    ['BOUNDARY',      7  ],   // white/black — warehouse outline
    ['ZONES',         3  ],   // green
    ['ZONE_RECV',     5  ],   // blue
    ['ZONE_DISP',     2  ],   // yellow
    ['ZONE_SUPPORT',  4  ],   // cyan
    ['RACKS_SHELVING',251],   // light gray
    ['RACKS_SELECTIVE',1 ],   // red — upright frames
    ['RACKS_DRIVEIN', 8  ],   // dark gray
    ['RACKS_CANTILEVER',6],   // magenta
    ['PALLETS',       30 ],   // orange
    ['CROSS_AISLE',   52 ],   // yellow-green
    ['DOCKS',         5  ],   // blue
    ['DIMENSIONS',    7  ],   // white
    ['TEXT',          7  ],   // white
  ];
  d('0','SECTION','2','TABLES','0','TABLE','2','LAYER','70',String(LAYERS.length));
  LAYERS.forEach(([name,col])=>{
    d('0','LAYER','2',name,'70','0','62',String(col),'6','CONTINUOUS');
  });
  d('0','ENDTABLE','0','ENDSEC');

  // ── ENTITIES ────────────────────────────────────────────────────────────
  d('0','SECTION','2','ENTITIES');

  // Helpers
  const rect2D = (layer, x1,y1,x2,y2,col=null) => {
    d('0','LWPOLYLINE','8',layer,'90','4','70','1');
    if(col) d('62',String(col));
    [[x1,y1],[x2,y1],[x2,y2],[x1,y2]].forEach(([x,y])=>d('10',n(x),'20',n(y)));
  };
  const line2D = (layer,x1,y1,x2,y2,col=null) => {
    d('0','LINE','8',layer);
    if(col) d('62',String(col));
    d('10',n(x1),'20',n(y1),'30','0.0','11',n(x2),'21',n(y2),'31','0.0');
  };
  const txt = (layer,x,y,h,str,hj=0) => {
    d('0','TEXT','8',layer,'10',n(x),'20',n(y),'30','0.0','40',n(h),'1',String(str));
    if(hj){ d('72',String(hj),'11',n(x),'21',n(y),'31','0.0'); }
  };
  const dimH = (x1,x2,y,label) => {
    const mx=(x1+x2)/2;
    line2D('DIMENSIONS',x1,y,x2,y); // dimension line
    line2D('DIMENSIONS',x1,y-0.4,x1,y+0.4); // tick
    line2D('DIMENSIONS',x2,y-0.4,x2,y+0.4); // tick
    txt('DIMENSIONS',mx,y+0.5,0.5,label,1);
  };
  const dimV = (x,y1,y2,label) => {
    line2D('DIMENSIONS',x,y1,x,y2);
    line2D('DIMENSIONS',x-0.4,y1,x+0.4,y1);
    line2D('DIMENSIONS',x-0.4,y2,x+0.4,y2);
    txt('DIMENSIONS',x+0.7,(y1+y2)/2,0.4,label);
  };

  // ── WAREHOUSE OUTLINE ──────────────────────────────────────────────────
  rect2D('BOUNDARY',0,0,wW,wL,7);

  // ── ZONE LAYOUT ────────────────────────────────────────────────────────
  const stagingH = Math.max(4,(receivingArea||80)/wW);
  const supportH = Math.max(2,((officeArea||50)+(mheArea||0))/wW);
  const totZA    = Object.values(zoneAreas).reduce((s,a)=>s+a,0)||1;
  const availH   = Math.max(4,wL-stagingH-supportH);
  const ZORD=['golden','mid','reserve','bulk','long'];
  const ZCOL={golden:83,mid:52,reserve:30,bulk:9,long:6};
  const ZLBL={golden:'Golden Zone (VF/F)',mid:'Mid-Level (M)',
    reserve:'Reserve (S)',bulk:'Bulk (VS)',long:'Long Goods'};
  let yCur=stagingH; const ZH={};
  ZORD.forEach(z=>{
    const h=((zoneAreas[z]||0)/totZA)*availH;
    ZH[z]={y0:yCur,h:Math.max(0,h)}; yCur+=h;
  });

  ZORD.forEach(z=>{
    const{y0,h}=ZH[z]; if(h<0.1) return;
    rect2D('ZONES',0,y0,wW,y0+h,ZCOL[z]);
    txt('TEXT',wW/2,y0+h/2+0.3,Math.min(1.2,h*0.25),ZLBL[z]||z,1);
    txt('TEXT',wW/2,y0+h/2-0.5,Math.min(0.7,h*0.12),
      `${(h*wW).toFixed(0)}m²  (${Math.round(h*wW*10.7639).toLocaleString()} sq ft)`,1);
  });

  // Staging
  rect2D('ZONE_RECV',0,0,wW/2,stagingH,5);
  rect2D('ZONE_DISP',wW/2,0,wW,stagingH,2);
  txt('TEXT',wW/4,stagingH/2+0.3,0.7,'RECEIVING / GRN',1);
  txt('TEXT',wW*3/4,stagingH/2+0.3,0.7,'DISPATCH',1);
  txt('TEXT',wW/4,stagingH/2-0.5,0.5,
    `${receivingArea}m²  (${Math.round(receivingArea*10.7639).toLocaleString()} sq ft)`,1);
  txt('TEXT',wW*3/4,stagingH/2-0.5,0.5,
    `${dispatchArea}m²  (${Math.round(dispatchArea*10.7639).toLocaleString()} sq ft)`,1);

  // Support (north)
  rect2D('ZONE_SUPPORT',0,wL-supportH,wW/2,wL,4);
  rect2D('ZONE_SUPPORT',wW/2,wL-supportH,wW,wL,4);
  txt('TEXT',wW/4,wL-supportH/2,0.6,'OFFICE / WELFARE',1);
  txt('TEXT',wW*3/4,wL-supportH/2,0.6,'MHE CHARGING',1);

  // ── RACK ROWS ──────────────────────────────────────────────────────────
  const RACK_DZ={shelving:'golden',liveStorage:'golden',selective:'reserve',
    doubleDeep:'reserve',driveIn:'bulk',cantilever:'long'};
  const RD2={shelving:0.6,liveStorage:1.5,selective:1.1,driveIn:6.6,doubleDeep:2.4,cantilever:2.5};
  const RACK_LAYER={shelving:'RACKS_SHELVING',liveStorage:'RACKS_SHELVING',
    selective:'RACKS_SELECTIVE',doubleDeep:'RACKS_SELECTIVE',
    driveIn:'RACKS_DRIVEIN',cantilever:'RACKS_CANTILEVER',ground:'RACKS_SHELVING'};

  const z2R={};
  (rackConfig||[]).forEach(cfg=>{
    const zone=(analysis?.slotted||[]).find(s=>s.rack===cfg.rack)?.zone
      || RACK_DZ[cfg.rack]||'golden';
    if(!z2R[zone]) z2R[zone]=[];
    if(!z2R[zone].find(r=>r.rack===cfg.rack)) z2R[zone].push({rack:cfg.rack,cfg});
  });

  ZORD.forEach(zone=>{
    const{y0,h}=ZH[zone]; if(h<0.3) return;
    const zRacks=z2R[zone]||[{rack:'shelving',cfg:null}];
    const totalSlot=zRacks.reduce((s,{rack:dom})=>s+(RD2[dom]||0.6)+aisleM,0)||1;
    let subY=y0;

    zRacks.forEach(({rack:dom})=>{
      const rd=RD2[dom]||0.6, slot=rd+aisleM;
      const subH=Math.max(slot,(slot/totalSlot)*h);
      const nRows=Math.max(1,Math.floor(subH/slot));
      const crossAfter=Math.floor(nRows/2);
      const layer=RACK_LAYER[dom]||'RACKS_SHELVING';

      for(let i=0;i<nRows;i++){
        const extra=(nRows>2&&i>=crossAfter)?aisleM:0;
        const ry=subY+i*slot+extra+aisleM/2;
        if(ry+rd>subY+subH-0.05) break;

        // Cross aisle stripe
        if(nRows>2&&i===crossAfter){
          rect2D('CROSS_AISLE',0,ry-aisleM,wW,ry,52);
          txt('TEXT',wW/2,ry-aisleM/2,0.4,'CROSS AISLE',1);
        }

        if(dom==='selective'||dom==='doubleDeep'){
          rect2D(layer,0,ry,wW,ry+rd);
          const bayW=2.7, nB=Math.floor(wW/bayW);
          for(let b=0;b<=nB;b++) line2D(layer,b*bayW,ry,b*bayW,ry+rd,1); // upright columns
          // Pallet positions — 2 per bay
          for(let b=0;b<nB;b++){
            const bx=b*bayW;
            rect2D('PALLETS',bx+0.15,ry+0.05,bx+1.30,ry+rd-0.05,30);
            rect2D('PALLETS',bx+1.40,ry+0.05,bx+bayW-0.15,ry+rd-0.05,30);
          }
          txt('TEXT',wW/2,ry+rd/2,0.3,dom==='doubleDeep'?'DOUBLE-DEEP RACK':'SELECTIVE PALLET RACK',1);

        } else if(dom==='driveIn'){
          rect2D(layer,0,ry,wW,ry+rd);
          const laneW=2.7, nL=Math.floor(wW/laneW);
          for(let ln=0;ln<=nL;ln++) line2D(layer,ln*laneW,ry,ln*laneW,ry+rd,8); // lane dividers
          const palDeep=Math.floor(rd/1.2);
          for(let ln=0;ln<nL;ln++){
            for(let p=0;p<palDeep;p++){
              const pz=ry+(p/palDeep)*(rd-0.1);
              rect2D('PALLETS',ln*laneW+0.2,pz+0.05,(ln+1)*laneW-0.2,pz+1.05,30);
            }
          }
          txt('TEXT',wW/2,ry+rd/2,0.3,'DRIVE-IN RACK',1);
          line2D(layer,wW/2-2,ry-0.5,wW/2+2,ry-0.5,8);
          txt('TEXT',wW/2,ry-0.7,0.3,'ENTRY ↔',1);

        } else if(dom==='cantilever'){
          rect2D(layer,0,ry,wW,ry+rd);
          line2D(layer,0,ry+rd/2,wW,ry+rd/2,6); // spine
          const ssp=1.5, nSp=Math.floor(wW/ssp);
          for(let sp=0;sp<=nSp;sp++){
            const sx=sp*ssp;
            line2D(layer,sx,ry,sx,ry+rd,6); // spine post
            line2D(layer,sx,ry+rd/2,sx,ry+0.15,6); // front arm
            line2D(layer,sx,ry+rd/2,sx,ry+rd-0.15,6); // rear arm
          }
          txt('TEXT',wW/2,ry+rd/2,0.3,'CANTILEVER RACK',1);

        } else if(dom==='ground'){
          // Ground locations — floor-level pads with yellow boundary
          rect2D(layer,0,ry,wW,ry+rd,30);
          const padW=2.0, nPads=Math.floor(wW/padW);
          for(let p=0;p<nPads;p++){
            const px=p*padW;
            rect2D(layer,px+0.05,ry+0.05,px+padW-0.05,ry+rd-0.05,2);
          }
          txt('TEXT',wW/2,ry+rd/2,0.3,'GROUND STORAGE',1);
        } else {
          // Shelving / live storage
          rect2D(layer,0.4,ry,wW-0.4,ry+rd);
          const bw=dom==='liveStorage'?1.5:0.9, nB2=Math.floor((wW-0.8)/bw);
          for(let b=1;b<nB2;b++) line2D(layer,0.4+b*bw,ry,0.4+b*bw,ry+rd);
        }
      }
      subY+=subH;
    });
  });

  // ── DOCK DOORS ─────────────────────────────────────────────────────────
  for(let d2=0;d2<totalDocks;d2++){
    const dx=(d2+0.5)*(wW/totalDocks)-1.75;
    line2D('DOCKS',dx,0,dx,-0.5);
    line2D('DOCKS',dx+3.5,0,dx+3.5,-0.5);
    line2D('DOCKS',dx,-0.5,dx+3.5,-0.5);
    txt('DOCKS',dx+1.75,-0.8,0.4,`D${d2+1}`,1);
  }

  // ── DIMENSION LINES ─────────────────────────────────────────────────────
  dimH(0,wW,-3.5,`${wW}m  (${(wW*MFT).toFixed(0)} ft)`);
  dimV(-4.5,0,wL,`${wL}m  (${(wL*MFT).toFixed(0)} ft)`);
  // Staging heights on right
  dimV(wW+3,0,stagingH,`Staging: ${stagingH.toFixed(1)}m`);
  let dimYc=stagingH;
  ZORD.forEach(z=>{ const{h}=ZH[z]; if(h<0.3) return; dimV(wW+3,dimYc,dimYc+h,`${h.toFixed(1)}m`); dimYc+=h; });
  dimV(wW+3,wL-supportH,wL,`Support: ${supportH.toFixed(1)}m`);

  // ── TITLE BLOCK ─────────────────────────────────────────────────────────
  line2D('TEXT',0,wL+2,wW,wL+2);
  txt('TEXT',wW/2,wL+3.0,1.0,'WAREHOUSE LAYOUT — FLOOR PLAN',1);
  txt('TEXT',wW/2,wL+2.2,0.55,
    `${wW}m × ${wL}m  |  ${totalGrossArea}m² gross  (${Math.round(totalGrossArea*10.7639).toLocaleString()} sq ft)`,1);
  txt('TEXT',0,wL+1.5,0.4,`Scale: 1:1  |  Units: Metres  |  Generated: ${new Date().toLocaleDateString()}`);
  txt('TEXT',wW,wL+1.5,0.4,'DensiCube Warehouse Designer',1);

  d('0','ENDSEC','0','EOF');

  const blob = new Blob([L.join('\n')], {type:'application/dxf'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `Warehouse_Plan_${wW}x${wL}m.dxf`;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); },300);
}


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
function calcForwardReserve(analysis, forwardRacks, reserveRacks, forwardDays, params) {
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
    const bc=BIN_CATALOG[binKey]; if(!bc?.phys) return;
    const [bL,bW,bH]=bc.phys;
    const totalLocs=binInfo.locs||0; if(!totalLocs) return;

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
function calcUserRackConfigFromSystemBins(analysis, userRacks, params) {
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
    const pref = BIN_PREFERRED_CATEGORY[binKey];
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
    if (binKey === 'LONG') {
      // LONG has no single phys → always goes to ground pass for per-SKU fitting
      bL=0; bW=0; bH=0;
    } else {
      const bc = BIN_CATALOG[binKey];
      if (!bc?.phys) return;
      [bL, bW, bH] = bc.phys;
    }
    const totalLocs=binInfo.locs||0; if(!totalLocs) return;

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
        const binPref = BIN_PREFERRED_CATEGORY[binKey];

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

      if(perSku && (binKey==='LONG'||longBins.has(binKey))){
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
      binDims:[bL,bW,bH],bayW:rW,bayD:rD,
      shelfH:fc.totalH,tierHeight:fc.levelH,clearance:0,
      levelH:fc.levelH,usableH:fc.usableH,beamClr:0,
      firstLevelH:fc.firstLevelH, calcNote:fc.calcNote,
      orientDesc:fc.orientDesc,wDimMm:fc.wDimMm,dDimMm:fc.dDimMm,hDimMm:fc.hDimMm,
      orientation:fc.orient,tiers:1,levels:fc.lvl,
      acrossW:fc.aw,acrossD:fc.ad,stackH:fc.stack,
      locsPerBay:fc.lpb,locsPerBayTotal:fc.lpb,
      locs:totalLocs,baysNeeded:bays,area,feasible:true,zone:ZONE_MAP[binKey]||'golden',
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
  const [viewMode3D, setViewMode3D] = useState('3d'); // '2d' | '3d'
  const [udViewMode,  setUdViewMode]  = useState('2d'); // user defined — default 2D (no WebGL risk)
  const [floorPlanFS, setFloorPlanFS] = useState(false); // fullscreen 2D plan
  const plan2DRef = useRef(null); // ref for 2D SVG download
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
    const res=calcUserRackConfigFromSystemBins(analysis,userRacks,params);
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
    const res=calcForwardReserve(analysis,forwardRacks,reserveRacks,forwardDays,params);
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

      // ── Step 4: Generate rack config ─────────────────────────────────────
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
        const res=calcForwardReserve(a,forwardRacks,reserveRacks,forwardDays,params);
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
        const res=calcUserRackConfigFromSystemBins(a,userRacks,params);
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
  const downloadPlan2D = (fmt='svg') => {
    const svgEl = plan2DRef.current?.querySelector('svg');
    if (!svgEl) return;
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
      const scale = 2;
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
          a.download = `Warehouse_Plan_${design?.wW||''}x${design?.wL||''}m_2x.png`;
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
                    <div style={{display:'flex',flexWrap:'wrap',gap:'7px',marginBottom:'12px'}}>
                      {Object.entries(analysis?.binSummary||{})
                        .sort((a,b)=>['XS','S','M','L','XL','LONG'].indexOf(a[0])-['XS','S','M','L','XL','LONG'].indexOf(b[0]))
                        .map(([band,info])=>{
                          const bc=BIN_CATALOG[band];
                          const COLORS={XS:['#f1f5f9','#64748b'],S:['#eff6ff','#1d4ed8'],
                            M:['#f5f3ff','#7c3aed'],L:['#f0fdf4','#166534'],
                            XL:['#fef9c3','#854d0e'],LONG:['#fdf4ff','#9333ea']};
                          const [bg,col]=COLORS[band]||['#f8fafc','#374151'];
                          return(
                            <div key={band} style={{background:bg,border:`1px solid ${col}33`,
                              borderRadius:'8px',padding:'7px 10px',minWidth:'100px'}}>
                              <div style={{fontWeight:'800',fontSize:'13px',color:col}}>{band}</div>
                              <div style={{fontSize:'9px',color:col,opacity:0.8}}>{info.name}</div>
                              {bc?.phys&&<div style={{fontSize:'9px',color:'#6b7280'}}>{bc.phys.join('×')}mm</div>}
                              <div style={{fontSize:'11px',fontWeight:'700',color:'#0f172a',marginTop:'2px'}}>
                                {(info.locs||0).toLocaleString()} locs
                              </div>
                            </div>
                          );
                        })}
                    </div>
                    <button onClick={()=>setUdStep(2)}
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
                        shelving:   {bayW:'900', bayD:'600', bayH:'2200',levels:'4'},
                        liveStorage:{bayW:'900', bayD:'1500',bayH:'2200',levels:'4'},
                        selective:  {bayW:'2700',bayD:'1100',bayH:'6000',levels:'4'},
                        doubleDeep: {bayW:'2700',bayD:'2200',bayH:'6000',levels:'4'},
                        driveIn:    {bayW:'2700',bayD:'6600',bayH:'6000',levels:'4'},
                        cantilever: {bayW:'1500',bayD:'2500',bayH:'3000',levels:'6'},
                        ground:     {bayW:'1500',bayD:'1200',bayH:'',   levels:'2'},
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
                          ].filter(Boolean).map(([label,field])=>(
                            <div key={field}>
                              <div style={{fontSize:'10px',color:'#6b7280',
                                fontWeight:'600',marginBottom:'3px'}}>{label}</div>
                              <input type="number" min="1" value={d[field]||''}
                                onChange={e=>upd(field,e.target.value)}
                                placeholder={field==='bayH'&&key==='ground'?'optional':'mm'}
                                style={{...inp,marginBottom:0,width:'100%',
                                  fontSize:'12px',padding:'5px 8px'}}/>
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
                {udViewMode==='3d'
                  ? <Warehouse3DModel analysis={analysis} design={userDesign} params={params} rackConfig={userRackConfig||rackConfig}/>
                  : <FloorPlanSVG analysis={analysis} design={userDesign} params={params} rackConfig={userRackConfig||rackConfig}/>
                }
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
                  ['Gross Area', `${(userDesign.wW*userDesign.wL).toLocaleString()}m²`, '#eff6ff','#1d4ed8'],
                  ['Dimensions', `${userDesign.wW}×${userDesign.wL}m`, '#f0fdf4','#166534'],
                  ['Rack Area', `${userDesign.netRackArea||0}m²`, '#f5f3ff','#7c3aed'],
                ].map(([l,v,bg,col])=>(
                  <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',
                    textAlign:'center',border:`1px solid ${col}22`}}>
                    <div style={{fontSize:'16px',fontWeight:'800',color:col}}>{v}</div>
                    <div style={{fontSize:'10px',color:'#6b7280',marginTop:'3px',
                      fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                  </div>
                ))}
              </div>
              {/* Download */}
              <div style={{display:'flex',gap:'10px'}}>
                <button onClick={()=>exportExcel(analysis,userDesign,params,userRackConfig)}
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
                  {viewMode3D==='3d'?'🧊 3D Isometric View':'🗺 Plan View (Top)'}
                </div>
                <div style={{display:'flex',gap:'6px'}}>
                  {[['2d','📐 Plan'],['3d','🧊 3D']].map(([m,l])=>(
                    <button key={m} onClick={()=>setViewMode3D(m)}
                      style={{padding:'5px 14px',borderRadius:'7px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        border:`2px solid ${viewMode3D===m?'#7c3aed':'#e2e8f0'}`,
                        background:viewMode3D===m?'#f5f3ff':'#fff',
                        color:viewMode3D===m?'#7c3aed':'#6b7280'}}>
                      {l}
                    </button>))}
                  {viewMode3D==='2d'&&(
                    <button onClick={()=>setFloorPlanFS(true)}
                      title="View full screen"
                      style={{padding:'5px 12px',borderRadius:'7px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        border:'2px solid #e2e8f0',background:'#fff',color:'#6b7280'}}>
                      ⛶ Full Screen
                    </button>
                  )}
                </div>
              </div>

              {viewMode3D==='3d'
                ? <Warehouse3DModel analysis={analysis} design={design} params={params} rackConfig={rackConfig}/>
                : (<div ref={plan2DRef}>
                    <FloorPlanSVG analysis={analysis} design={design} params={params} rackConfig={rackConfig}/>
                  </div>)}

              {/* Legend + Download (2D only) */}
              {viewMode3D==='2d'&&(
                <div style={{marginTop:'12px'}}>
                  {/* Download buttons */}
                  <div style={{display:'flex',gap:'8px',marginBottom:'10px',flexWrap:'wrap'}}>
                    <button onClick={()=>downloadPlan2D('svg')}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#f0fdf4',border:'1px solid #86efac',color:'#166534'}}>
                      ⬇ Download SVG
                    </button>
                    <button onClick={()=>downloadPlan2D('png')}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#eff6ff',border:'1px solid #93c5fd',color:'#1d4ed8'}}>
                      ⬇ Download PNG (2×)
                    </button>
                    <button onClick={()=>exportDXF(analysis,design,params,rackConfig)}
                      style={{padding:'7px 16px',borderRadius:'8px',cursor:'pointer',
                        fontFamily:'inherit',fontSize:'12px',fontWeight:'700',
                        background:'#fef9c3',border:'1px solid #fde047',color:'#854d0e'}}>
                      ⬇ Download DXF (AutoCAD)
                    </button>
                    <span style={{fontSize:'11px',color:'#9ca3af',alignSelf:'center'}}>
                      DXF opens in AutoCAD, FreeCAD, LibreCAD · Save as DWG in AutoCAD
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
            <div style={{display:'flex',gap:'8px',alignItems:'center'}}>
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
                ⬇ Download SVG
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
          <div style={{flex:1,overflow:'auto',padding:'0'}}>
            <div id="fs-plan-container" style={{width:'100%',height:'100%'}}>
              <FloorPlanSVG
                analysis={analysis}
                design={userDesign||design}
                params={params}
                rackConfig={userRackConfig||rackConfig}
                fullscreen={true}/>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
