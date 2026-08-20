// ─── BULK SKU CALCULATOR ─────────────────────────────────────────────────────
import { useState } from 'react';
import * as XLSX from 'xlsx';
import { CONFIG } from '../config.js';
import { calcMixed, calcMixedDetailed } from '../algorithms/packing.js';
import { S, UtilBadge } from '../components/styles.jsx';
import PasteFromExcel from '../components/PasteFromExcel.jsx';

// ─── PALLET PRESETS ───────────────────────────────────────────────────────────
const PALLET_PRESETS = {
  standard: { label: 'Standard (1200×1000×1200)', L:1200, W:1000, H:1200 },
  euro:     { label: 'Euro Pallet (1200×800×1200)',L:1200, W:800,  H:1200 },
  half:     { label: 'Half Pallet (600×800×1200)', L:600,  W:800,  H:1200 },
  custom:   { label: 'Custom',                     L:0,    W:0,    H:0    },
};


// ─── BOXES PER PALLET (LOCK HEIGHT) ──────────────────────────────────────────
// Lock height: box H is always vertical — only horizontal rotation (L↔W swap)

// ─── PALLET STABILITY ────────────────────────────────────────────────────────
// A load tips once tilted past theta = atan(base / height), so the
// height-to-base ratio decides whether it can travel upright:
//   ratio 2:1 -> 26.6 deg   stable
//   ratio 3:1 -> 18.4 deg   marginal, needs wrap or strapping
//   ratio 5:1 -> 11.3 deg   unstable
//
// Crucially the base is that of the WRAPPED BLOCK on the pallet, not of one
// carton. Eight thin panels standing side by side and stretch-wrapped form a
// stable block even though a single panel alone would topple. The block is
// limited by how many units are actually available, so a part-filled pallet is
// judged on what really sits on it.
const SLENDER_STABLE = 2.0, SLENDER_MARGINAL = 3.0, OVERHANG_TOL = 25;

function tipAngle(base, h) {
  return h > 0 ? Math.atan(base / h) * 180 / Math.PI : 90;
}
function rateSlender(s) {
  if (s <= SLENDER_STABLE) return 'stable';
  if (s <= SLENDER_MARGINAL) return 'marginal';
  return 'unstable';
}

// Smallest sensible pallet for the flattest orientation, rounded up to 100mm
function suggestPallet(cl, cw, ch) {
  const opts = [
    { a:cw, b:ch, h:cl, n:'lying on its length' },
    { a:cl, b:ch, h:cw, n:'lying on its width'  },
    { a:cl, b:cw, h:ch, n:'standing upright'    },
  ].map(d => ({ ...d, s: Math.min(d.a,d.b) > 0 ? d.h/Math.min(d.a,d.b) : Infinity }))
   .sort((x,y) => x.s - y.s);
  const b = opts[0], r = v => Math.ceil(v/100)*100;
  return { L:r(Math.max(b.a,b.b)), W:r(Math.min(b.a,b.b)), how:b.n,
    slender:+b.s.toFixed(2), tip:+tipAngle(Math.min(b.a,b.b), b.h).toFixed(1) };
}

// Assess one component on one pallet across all three orientations.
// qty (optional) caps the block to the units actually on hand.
function assessStability(cl, cw, ch, pl, pw, ph, qty, assumeWrapped) {
  assumeWrapped = assumeWrapped !== false;   // wrapping is standard practice
  cl = parseFloat(cl)||0; cw = parseFloat(cw)||0; ch = parseFloat(ch)||0;
  pl = parseFloat(pl)||0; pw = parseFloat(pw)||0; ph = parseFloat(ph)||0;
  const have = parseFloat(qty) > 0 ? parseFloat(qty) : Infinity;
  if (!(cl>0 && cw>0 && ch>0 && pl>0 && pw>0)) return null;

  const all = [
    { name:'Length vertical', a:cw, b:ch, h:cl },
    { name:'Width vertical',  a:cl, b:ch, h:cw },
    { name:'Height vertical', a:cl, b:cw, h:ch },
  ].map(d => {
    const fitsBase = (d.a <= pl+OVERHANG_TOL && d.b <= pw+OVERHANG_TOL)
                  || (d.b <= pl+OVERHANG_TOL && d.a <= pw+OVERHANG_TOL);
    const fitsH = !(ph > 0) || d.h <= ph;

    // Best arrangement of this footprint on the pallet
    const arrangements = [
      { nL:Math.floor(pl/d.a), nW:Math.floor(pw/d.b), ba:d.a, bb:d.b },
      { nL:Math.floor(pl/d.b), nW:Math.floor(pw/d.a), ba:d.b, bb:d.a },
    ].filter(x => x.nL > 0 && x.nW > 0)
     .sort((x,y) => (y.nL*y.nW) - (x.nL*x.nW));
    const arr = arrangements[0];

    let perLayer = 0, layers = 0, blockBase = Math.min(d.a, d.b), loadH = d.h;
    if (arr) {
      perLayer = arr.nL * arr.nW;
      layers   = ph > 0 && d.h > 0 ? Math.max(1, Math.floor(ph/d.h)) : 1;
      // Only count units we actually have
      const unitsOnPallet = Math.min(perLayer * layers, have);
      const unitsPerLayer = Math.min(perLayer, unitsOnPallet);
      // Fill the longer run first, so the block stays as square as it can
      const alongL = Math.min(arr.nL, Math.max(1, Math.ceil(unitsPerLayer/arr.nW)));
      const alongW = Math.min(arr.nW, Math.max(1, Math.ceil(unitsPerLayer/alongL)));
      const blockL = alongL * arr.ba, blockW = alongW * arr.bb;
      blockBase = Math.min(blockL, blockW);
      const usedLayers = Math.max(1, Math.ceil(unitsOnPallet / Math.max(1, perLayer)));
      loadH = Math.min(usedLayers, layers) * d.h;
    }

    const slenderSingle = Math.min(d.a,d.b) > 0 ? d.h/Math.min(d.a,d.b) : Infinity;
    const slender = blockBase > 0 ? loadH/blockBase : Infinity;
    return { ...d, base:blockBase, loadH,
      slender:+slender.toFixed(2), tip:+tipAngle(blockBase, loadH).toFixed(1),
      slenderSingle:+slenderSingle.toFixed(2),
      tipSingle:+tipAngle(Math.min(d.a,d.b), d.h).toFixed(1),
      rating:rateSlender(slender), fitsBase, fitsH,
      usable:fitsBase && fitsH, perLayer, layers };
  });

  const usable = all.filter(x => x.usable);
  const rank = { stable:0, marginal:1, unstable:2 };
  const pool = usable.length ? usable : all;
  const best = pool.slice().sort((x,y) =>
    rank[x.rating]-rank[y.rating] || y.perLayer-x.perLayer)[0];

  // Wrapping is a PROCESS the floor may skip, so judge both ways.
  // ratingLoose = what happens if nobody wraps the layer.
  if (best) {
    best.ratingLoose = rateSlender(best.slenderSingle);
    best.wrapCritical = assumeWrapped
      && best.usable
      && best.perLayer > 1
      && rateSlender(best.slender) === 'stable'
      && best.ratingLoose !== 'stable';
  }
  const okWrapped = usable.some(x => x.rating === 'stable');
  const okLoose   = usable.some(x => rateSlender(x.slenderSingle) === 'stable');
  // With wrapping assumed off, only the single-unit result counts
  const ok = assumeWrapped ? okWrapped : okLoose;

  let advice = '', suggested = null;
  if (!usable.length) {
    suggested = suggestPallet(cl, cw, ch);
    advice = 'Does not fit this pallet. Use ' + suggested.L + ' x ' + suggested.W
      + 'mm with the item ' + suggested.how + ' (' + suggested.slender + ':1).';
  } else if (ok) {
    advice = best.wrapCritical
      ? 'Only stable when the layer of ' + best.perLayer
        + ' is stretch-wrapped as one block. Loose, a single unit is '
        + best.slenderSingle + ':1 and will topple — wrapping is mandatory.'
      : (best.perLayer > 1
          ? 'Stable as a wrapped block of ' + best.perLayer + ' per layer.'
          : 'Stable as loaded.');
  } else if (!assumeWrapped && okWrapped) {
    advice = 'Unstable loose (' + best.slenderSingle + ':1). Stretch wrap the layer of '
      + best.perLayer + ' into one block to reach ' + best.slender + ':1.';
  } else {
    const marg = usable.filter(x => x.rating === 'marginal')[0];
    if (marg) {
      advice = 'Stretch wrap the layer into one block ('
        + marg.slender + ':1, ' + marg.tip + ' deg) or strap to the pallet.';
    } else {
      suggested = suggestPallet(cl, cw, ch);
      const fits = suggested.L <= pl+OVERHANG_TOL && suggested.W <= pw+OVERHANG_TOL;
      advice = fits
        ? 'Strap the load to the pallet or use a stillage.'
        : 'No stable orientation here. Use ' + suggested.L + ' x ' + suggested.W
          + 'mm with the item ' + suggested.how + ' (' + suggested.slender + ':1).';
    }
  }
  const shownRating = !usable.length ? 'unstable'
    : assumeWrapped ? best.rating : rateSlender(best.slenderSingle);
  return { ok, rating: shownRating, wrapCritical: best ? best.wrapCritical : false,
    assumeWrapped, best, all, advice, suggested };
}


// ─── STACKING LAYERS CALCULATOR ──────────────────────────────────────────────
// Returns how many layers high boxes are stacked on the pallet (vertical direction)
function calcStackLayers(pL, pW, pH, sl, sw, sh, isLocked) {
  if (isLocked) return sh > 0 ? Math.floor(pH / sh) : 0;
  // Try all 6 orientations; find which gives most boxes, return its layer count
  const dims = [sl, sw, sh];
  const perms = [[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];
  let bestTotal = 0, bestLayers = 1;
  perms.forEach(([x,y,z]) => {
    if (dims[z] <= 0) return;
    const across = Math.floor(pL/dims[x]) * Math.floor(pW/dims[y]);
    const layers = Math.floor(pH/dims[z]);
    const total  = across * layers;
    if (total > bestTotal) { bestTotal = total; bestLayers = layers; }
  });
  return bestLayers;
}

// ─── SIZE-BASED FITMENT CHECK FOR MIXED PALLETS ───────────────────────────────
// Checks whether boxes of multiple SKUs can physically co-exist on one pallet.
// Three checks:
//  1. Each box fits the pallet in at least one orientation
//  2. Combined height (separate layers per SKU) ≤ pallet height
//  3. Combined floor area (each SKU's strip on pallet base) ≤ pallet area
function calcMixedFitment(palletSkus, pL, pW, pH) {
  // Find best stacking orientation for each SKU
  const oriented = palletSkus.map(sku => {
    const dims = [sku.sl, sku.sw, sku.sh];
    if (!dims.every(d => d > 0)) return { ...sku, ok: false, reason: 'Missing dimensions' };

    const perms = [[0,1,2],[0,2,1],[1,0,2],[1,2,0],[2,0,1],[2,1,0]];
    let best = null;
    perms.forEach(([x,y,z]) => {
      if (dims[z] > pH) return;               // too tall
      const acL = Math.floor(pL / dims[x]);
      const acW = Math.floor(pW / dims[y]);
      if (!acL || !acW) return;               // doesn't tile pallet floor
      const perLayer = acL * acW;
      if (!best || perLayer > best.perLayer)
        best = { perLayer, boxH: dims[z], footL: dims[x], footW: dims[y], acL, acW };
    });

    if (!best) return { ...sku, ok: false,
      reason: `Box ${dims.join('×')}mm can't fit pallet ${pL}×${pW}×${pH}mm` };

    // How many boxes of this SKU on this mixed pallet
    const neededBoxes  = Math.max(1, Math.round((sku.remainder||0) * (sku.bpp||1)));
    const layersNeeded = Math.ceil(neededBoxes / best.perLayer);
    const heightNeeded = layersNeeded * best.boxH;
    // Floor area fraction this SKU occupies per layer
    const floorFrac    = (best.footL * best.footW * Math.min(neededBoxes, best.perLayer)) / (pL * pW);

    return { ...sku, ok: true, best, neededBoxes, layersNeeded, heightNeeded, floorFrac };
  });

  // 1. Any box incompatible with pallet?
  const bad = oriented.filter(o => !o.ok);
  if (bad.length) return {
    feasible: false,
    reason: bad.map(b => `${b.name}: ${b.reason}`).join('; '),
    warning: null, oriented,
  };

  // 2. Height check — each SKU gets its own layer stack
  const totalH = oriented.reduce((s, o) => s + o.heightNeeded, 0);
  if (totalH > pH) return {
    feasible: false,
    reason: `Stacked height ${totalH}mm > pallet height ${pH}mm`,
    warning: null, oriented,
  };

  // 3. Floor area check — each SKU occupies a horizontal strip
  const totalFloor = oriented.reduce((s, o) => s + o.floorFrac, 0);
  if (totalFloor > 1.05) return {
    feasible: false,
    reason: `Combined floor area ${(totalFloor*100).toFixed(0)}% exceeds pallet (${pL}×${pW}mm)`,
    warning: null, oriented,
  };

  // 4. Height compatibility warning (for stable stacking in shared layers)
  const stackHs = oriented.map(o => o.best.boxH);
  const hMin = Math.min(...stackHs), hMax = Math.max(...stackHs);
  const warning = (hMax / hMin > 1.5)
    ? `Layer-height mismatch: ${hMin}–${hMax}mm (${(hMax/hMin).toFixed(1)}× ratio) — may be unstable`
    : null;

  return {
    feasible: true, warning,
    totalH, totalFloor: +(totalFloor * 100).toFixed(1),
    oriented,
  };
}

// ─── PALLET MIXING ALGORITHM ──────────────────────────────────────────────────
function calcPalletMix(skus, pL, pW, pH, maxSkus, lockHeight=false, lockedSkus=new Set()) {
  // Step 1: calc boxes per pallet and pallet equivalents per SKU
  const items = skus
    .filter(s => s.sl>0 && s.sw>0 && s.sh>0 && s.qtyAvail>0)
    .map(s => {
      // Per-SKU lock overrides global lock
      const isLocked = lockHeight || lockedSkus.has(s.name);
      const bpp = isLocked
        ? calcMixedDetailed(pL, pW, pH, s.sl, s.sw, s.sh, { lockHeight:true }).total
        : calcMixed(pL, pW, pH, s.sl, s.sw, s.sh).total;
      const stackLayers = calcStackLayers(pL, pW, pH, s.sl, s.sw, s.sh, isLocked);
      if (!bpp || bpp === 0) return { ...s, bpp:0, palletEquiv:null, fullPallets:0, remainder:0, stackLayers:0, error:'Box too large for pallet' };
      const pe     = s.qtyAvail / bpp;
      const full   = Math.floor(pe);
      const rem    = +(pe - full).toFixed(6);
      return { ...s, bpp, palletEquiv:+pe.toFixed(4), fullPallets:full, remainder:rem, stackLayers, heightLocked:isLocked };
    });

  // Step 2: collect remainders (SKUs with fractional part > 0)
  const withRem = items
    .filter(r => r.remainder > 0 && r.palletEquiv !== null)
    .sort((a, b) => b.remainder - a.remainder);

  // Step 3: first-fit decreasing bin packing (max N SKUs per pallet, capacity=1.0)
  const mixedPallets = [];
  const placed = new Set();

  withRem.forEach(item => {
    if (placed.has(item.name)) return;
    let fit = false;
    for (const pallet of mixedPallets) {
      if (pallet.skus.length < maxSkus &&
          pallet.used + item.remainder <= 1.0001) {
        pallet.skus.push(item);
        pallet.used = +(pallet.used + item.remainder).toFixed(6);
        placed.add(item.name);
        fit = true;
        break;
      }
    }
    if (!fit) {
      mixedPallets.push({ skus:[item], used: item.remainder });
      placed.add(item.name);
    }
  });

  const totalFull   = items.reduce((s,r) => s + (r.fullPallets||0), 0);
  const totalBefore = items
    .filter(r => r.palletEquiv !== null)
    .reduce((s,r) => s + Math.ceil(r.palletEquiv), 0);
  const totalAfter  = totalFull + mixedPallets.length;
  const savings     = totalBefore - totalAfter;

  // ── Size-based fitment check for each mixed pallet ────────────────────────
  const mixedWithFitment = mixedPallets.map(p => {
    const fit = p.skus.length > 1
      ? calcMixedFitment(p.skus, pL, pW, pH)
      : { feasible: true, warning: null, totalH: null, totalFloor: null };
    return { ...p, fitment: fit };
  });

  return { items, mixedPallets: mixedWithFitment, totalFull, totalBefore, totalAfter, savings };
}

// ─── COMPONENT ────────────────────────────────────────────────────────────────
export default function ContainerSkuTool({ isPro, onUpgrade }) {
  // Container
  const [cL,setCL]=useState(''); const [cW,setCW]=useState('');
  const [cH,setCH]=useState(''); const [cMaxWt,setCMaxWt]=useState('');

  // SKU data
  const [fileName,setFileName]  = useState('');
  const [rawSkus, setRawSkus]   = useState(null);
  const [skuCount,setSkuCount]  = useState(0);
  const [dragOver,setDragOver]  = useState(false);

  // Container results
  const [results,   setResults]   = useState(null);
  const [processing,setProcessing]= useState(false);
  const [progress,  setProgress]  = useState(0);
  const [capped,    setCapped]    = useState(false);
  const [error,     setError]     = useState('');

  // Pallet mixing
  const [pPreset,     setPPreset]     = useState('standard');
  const [pL,setPL]=useState('1200'); const [pW,setPW]=useState('1000'); const [pH,setPH]=useState('1200');
  const [stabCheck,setStabCheck]=useState(false); // opt-in pallet stability analysis
  const [wrapped,setWrapped]=useState(true);      // are unit loads stretch-wrapped in practice?
  const [maxSkus,     setMaxSkus]     = useState(4);
  const [lockHeight,  setLockHeight]  = useState(false);
  const [lockedSkus,       setLockedSkus]       = useState(new Set());
  const [lockedCategories, setLockedCategories] = useState(new Set()); // category-level locks
  const [mixResult,   setMixResult]   = useState(null);
  const [mixError,    setMixError]    = useState('');

  const container = {
    cL:parseFloat(cL)||0, cW:parseFloat(cW)||0,
    cH:parseFloat(cH)||0, cMaxWt:parseFloat(cMaxWt)||0
  };
  const valid = container.cL>0 && container.cW>0 && container.cH>0 && container.cMaxWt>0;

  // ── Per-SKU container packing ──────────────────────────────────────────────
  function pSkus(cont, skus, locked=new Set()) {
    const { cL, cW, cH, cMaxWt } = cont;
    const cv = cL * cW * cH;
    return skus.map(s => {
      const { name, sl, sw, sh, swt, qtyAvail } = s;
      if (!name) return null;
      if (sl<=0||sw<=0||sh<=0) return { name, error:'Invalid dimensions' };
      const isLocked = locked.has(name) || lockedCategories.has(s.category||'Uncategorised');
      let vQ, orient;
      if (isLocked) {
        // Height stays vertical, but the leftover strips along L, W and H are
        // still filled. This previously used a plain floor() product which
        // silently threw away up to 13% of the container.
        const res = calcMixedDetailed(cL, cW, cH, sl, sw, sh, { lockHeight:true });
        vQ = res.total; orient = res.orient;
      } else {
        const res = calcMixed(cL, cW, cH, sl, sw, sh);
        vQ = res.total; orient = res.orient;
      }
      const cStackLayers = calcStackLayers(cL, cW, cH, sl, sw, sh, isLocked);
      let eV = qtyAvail>0 ? Math.min(vQ, qtyAvail) : vQ;
      let wQ = swt>0 ? Math.floor(cMaxWt/swt) : null;
      if (wQ!==null && qtyAvail>0) wQ = Math.min(wQ, qtyAvail);
      const eQ = wQ!==null ? Math.min(eV,wQ) : eV;
      const vu = (eQ*sl*sw*sh)/cv;
      const wu = swt>0 ? (eQ*swt)/cMaxWt : null;
      let con = 'Volume';
      if (wQ!==null && wQ<eV) con = 'Weight';
      if (qtyAvail>0 && eQ===qtyAvail) con = 'Stock Limit';
      const stab = stabCheck ? assessStability(sl, sw, sh, pL, pW, pH, qtyAvail, wrapped) : null;
      return { name, sl, sw, sh, category:s.category||'Uncategorised', volQty:eV, wtQty:wQ!==null?wQ:'N/A', effQty:eQ, qtyAvail, unitVol:sl*sw*sh, volUtil:vu, wtUtil:wu, orient, stackLayers:cStackLayers, constraint:con, heightLocked:isLocked, stab };
    }).filter(Boolean);
  }

  // ── File parse ─────────────────────────────────────────────────────────────
  const parseFile = file => {
    if (!file) return;
    setFileName(file.name); setError(''); setResults(null);
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(e.target.result,{type:'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
        let ds = 0;
        for (let i=0;i<Math.min(raw.length,20);i++) {
          const row = raw[i].map(c=>String(c).toLowerCase());
          if (row.some(c=>c.includes('sku')||c.includes('length')||c.includes('width'))) { ds=i+1; break; }
        }
        const skus = [];
        for (let i=ds;i<raw.length;i++) {
          const rr = raw[i]; if(!rr[0]&&!rr[1]) continue;
          skus.push({ name:String(rr[0]||'').trim(), sl:parseFloat(rr[1])||0,
            sw:parseFloat(rr[2])||0, sh:parseFloat(rr[3])||0,
            swt:parseFloat(rr[4])||0, qtyAvail:parseFloat(rr[5])>0?parseFloat(rr[5]):-1,
            category:String(rr[6]||'').trim()||'Uncategorised' });
        }
        if (!skus.length) { setError('No SKU data found.'); return; }
        setSkuCount(skus.length); setRawSkus(skus);
      } catch(err) { setError('Could not read file: '+err.message); }
    };
    r.readAsArrayBuffer(file);
  };

  // ── Run container calculation ──────────────────────────────────────────────
  const run = () => {
    if (!valid) { setError('Enter container dimensions.'); return; }
    if (!rawSkus) { setError('Upload or paste SKU data.'); return; }
    setError(''); setMixResult(null);
    let toProcess = rawSkus, cap = false;
    if (!isPro && rawSkus.length > CONFIG.freeSkuLimit) {
      toProcess = rawSkus.slice(0, CONFIG.freeSkuLimit); cap = true;
    }
    setCapped(cap); setProcessing(true); setProgress(0); setResults(null);
    const CHUNK = 500; let done = 0; const all = [];
    function next() {
      all.push(...pSkus(container, toProcess.slice(done, done+CHUNK), lockedSkus));
      done += CHUNK;
      setProgress(Math.round((Math.min(done,toProcess.length)/toProcess.length)*100));
      if (done < toProcess.length) setTimeout(next, 0);
      else { setResults(all); setProcessing(false); }
    }
    setTimeout(next, 50);
  };

  // ── Run pallet mixing ──────────────────────────────────────────────────────
  const runMix = () => {
    setMixError('');
    const PL=parseFloat(pL)||0, PW=parseFloat(pW)||0, PH=parseFloat(pH)||0;
    if (!PL||!PW||!PH) { setMixError('Enter pallet dimensions.'); return; }
    if (!rawSkus?.length) { setMixError('Load SKU data first.'); return; }
    const skusForMix = rawSkus
      .filter(s => s.qtyAvail > 0)
      .map(s => ({ ...s }));
    if (!skusForMix.length) { setMixError('No SKUs with available qty found.'); return; }
    const mr = calcPalletMix(skusForMix, PL, PW, PH, maxSkus, lockHeight, lockedSkus);
    setMixResult(mr);
  };

  // Toggle category-level lock — locks all SKUs in that category at once
  const toggleCategoryLock = (cat) => {
    setLockedCategories(prev => {
      const next = new Set(prev);
      if (next.has(cat)) next.delete(cat); else next.add(cat);
      // Re-compute all SKUs in this category
      if (rawSkus && valid) {
        const catSkus = rawSkus.filter(s=>(s.category||'Uncategorised')===cat);
        if (catSkus.length) {
          const updated = pSkus(container, catSkus, lockedSkus);
          setResults(prev2 => prev2
            ? prev2.map(r => {
                const u = updated.find(u=>u.name===r.name);
                return u || r;
              })
            : prev2);
        }
      }
      return next;
    });
  };

  // Toggle per-SKU lock in CONTAINER packing and instantly re-compute that SKU
  const toggleContainerLock = (skuName) => {
    setLockedSkus(prev => {
      const next = new Set(prev);
      if (next.has(skuName)) next.delete(skuName); else next.add(skuName);
      // Re-compute just this SKU's container result instantly
      const skuData = rawSkus?.find(s=>s.name===skuName);
      if (skuData && valid) {
        const [updated] = pSkus(container, [skuData], next);
        if (updated) {
          setResults(prev2 => prev2
            ? prev2.map(r => r.name===skuName ? updated : r)
            : prev2);
        }
      }
      // Also re-run pallet mix if it exists
      const PL=parseFloat(pL)||0,PW=parseFloat(pW)||0,PH=parseFloat(pH)||0;
      if (PL&&PW&&PH&&rawSkus?.length) {
        const skusForMix=rawSkus.filter(s=>s.qtyAvail>0).map(s=>({...s}));
        if (skusForMix.length) setMixResult(calcPalletMix(skusForMix,PL,PW,PH,maxSkus,lockHeight,next));
      }
      return next;
    });
  };

  // Toggle per-SKU height lock and re-run mix instantly
  const toggleSkuLock = (skuName) => {
    setLockedSkus(prev => {
      const next = new Set(prev);
      if (next.has(skuName)) next.delete(skuName);
      else next.add(skuName);
      // Re-run mix with updated locks
      const PL=parseFloat(pL)||0, PW=parseFloat(pW)||0, PH=parseFloat(pH)||0;
      if (PL && PW && PH && rawSkus?.length) {
        const skusForMix = rawSkus.filter(s => s.qtyAvail > 0).map(s => ({...s}));
        if (skusForMix.length) {
          const mr = calcPalletMix(skusForMix, PL, PW, PH, maxSkus, lockHeight, next);
          setMixResult(mr);
        }
      }
      return next;
    });
  };

  // ── Pallet preset change ───────────────────────────────────────────────────
  const onPresetChange = p => {
    setPPreset(p);
    if (p !== 'custom') {
      const pr = PALLET_PRESETS[p];
      setPL(String(pr.L)); setPW(String(pr.W)); setPH(String(pr.H));
    }
  };

  // ── Excel export ───────────────────────────────────────────────────────────
  const exp = () => {
    const wb = XLSX.utils.book_new();
    const h = ['SKU Name','Category','Max Qty (Volume)','Max Qty (Weight)',
      'Effective Max Qty','Volume Used (%)','Weight Used (%)',
      'L-axis (mm)','W-axis (mm)','H-axis / Stack (mm)',
      'Stacking Layers','Height Locked','Constraint',
      ...(stabCheck?['Pallet Stability','Height:Base Ratio','Stability Advice']:[])];
    const rows = results.map(r => {
      const cat = r.category||'Uncategorised';
      const [oL,oW,oH] = (r.orient||'').split('×');
      const st = r.stab;
      return r.error
        ? [r.name, cat, r.error, '', '', '', '', '', '', '', '', '', '',
           ...(stabCheck?['','','']:[])]
        : [r.name, cat,
           r.volQty,
           typeof r.wtQty==='number'?r.wtQty:'N/A',
           r.effQty,
           r.volUtil!=null?(r.volUtil*100).toFixed(2)+'%':'',
           r.wtUtil!=null?(r.wtUtil*100).toFixed(2)+'%':'',
           oL||'', oW||'', oH||'',
           r.stackLayers||'',
           r.heightLocked?'Locked':'',
           r.constraint,
           ...(stabCheck ? [
             st ? (st.wrapCritical&&st.ok ? 'Stable - wrap required'
                   : st.ok ? 'Stable'
                   : st.rating==='marginal' ? 'Needs securing' : 'Unstable') : '',
             st && st.best && st.best.usable
               ? (st.assumeWrapped?st.best.slender:st.best.slenderSingle)+':1' : '',
             st ? st.advice : '',
           ] : [])];
    });
    const ws = XLSX.utils.aoa_to_sheet([
      ['CONTAINER SKU PACKING RESULTS'],[],
      ['Container',`${container.cL}×${container.cW}×${container.cH}`,'Max Weight',container.cMaxWt],[],h,...rows]);
    ws['!cols']=[{wch:22},{wch:16},{wch:16},{wch:16},{wch:16},{wch:14},{wch:14},{wch:12},{wch:12},{wch:16},{wch:12},{wch:12},{wch:12},
      ...(stabCheck?[{wch:16},{wch:16},{wch:60}]:[])];
    XLSX.utils.book_append_sheet(wb,ws,'Container Results');

    if (mixResult) {
      const mixRows = [
        ['PALLET MIXING RESULTS'],
        [`Pallet: ${pL}×${pW}×${pH}mm | Max SKUs per pallet: ${maxSkus} | Height: ${lockHeight?'LOCKED (upright only)':'Free (all orientations)'}`],[],
        ['SKU','Qty Available','Boxes/Pallet','Pallet Equivalents','Full Pallets','Remainder'],
        ...mixResult.items.map(r=>[r.name, r.qtyAvail, r.bpp||'—',
          r.palletEquiv!=null?r.palletEquiv:'—', r.fullPallets, r.remainder>0?r.remainder:'0']),
        [],[],['MIXED PALLET GROUPS'],
        ['Pallet #','SKUs Mixed','Utilisation %','SKU List'],
        ...mixResult.mixedPallets.map((p,i)=>[
          `Mixed Pallet ${i+1}`,p.skus.length,(p.used*100).toFixed(1)+'%',
          p.skus.map(s=>s.name).join(', ')]),
        [],[],['SUMMARY'],
        ['Total pallets WITHOUT mixing (each SKU separate)',mixResult.totalBefore],
        ['Total pallets WITH mixing',mixResult.totalAfter],
        ['Pallets saved by mixing',mixResult.savings],
        ['% reduction',mixResult.totalBefore>0?((mixResult.savings/mixResult.totalBefore)*100).toFixed(1)+'%':'—'],
      ];
      const mws = XLSX.utils.aoa_to_sheet(mixRows);
      mws['!cols']=[{wch:24},{wch:14},{wch:14},{wch:20},{wch:14},{wch:14}];
      XLSX.utils.book_append_sheet(wb,mws,'Pallet Mixing');
    }
    XLSX.writeFile(wb,'Bulk_SKU_Results.xlsx');
  };

  // Each utilisation band is reported two ways: how many SKUs sit in it, and
  // how many units those SKUs account for. A band can hold few SKUs but most of
  // the volume, or the reverse, and the two tell very different stories.
  const band = (test) => {
    const rows = results ? results.filter(r => !r.error && test(r)) : [];
    return { skus: rows.length,
      qty: rows.reduce((s,r) => s + (Number(r.effQty)||0), 0) };
  };
  // Combined utilisation weights every SKU by the container space it actually
  // consumes, so a 981-unit SKU counts far more than a 6-unit one. A plain
  // average over SKUs treats them equally and flatters a catalogue with many
  // small, badly-packing lines.
  //   containers needed = stock / units per container
  //   combined %        = total goods volume / total container volume consumed
  const combined = (() => {
    if (!results) return null;
    let goods = 0, space = 0, weighted = 0, simple = 0, n = 0;
    results.forEach(r => {
      if (r.error || !(r.effQty > 0) || !(r.unitVol > 0) || !(cv > 0)) return;
      const stock = r.qtyAvail > 0 ? r.qtyAvail : r.effQty;
      const containers = stock / r.effQty;      // fractional: part-filled counts
      goods += stock * r.unitVol;
      space += containers * cv;
      simple += r.volUtil; n++;
    });
    if (!n || space <= 0) return null;
    weighted = goods / space;
    return { weighted, simple: simple / n, n,
      containers: space / cv, goodsM3: goods / 1e9, spaceM3: space / 1e9 };
  })();
  const bAll = band(() => true);
  const bG   = band(r => r.volUtil >= 0.75);
  const bO   = band(r => r.volUtil >= 0.5 && r.volUtil < 0.75);
  const bL   = band(r => r.volUtil < 0.5);
  const pct  = n => bAll.qty > 0 ? Math.round(n / bAll.qty * 100) : 0;

  const inp = { ...S.input, marginBottom:'4px' };
  const lbl = { ...S.label };

  return (
    <div>
      <div style={S.sectionDesc}>
        Upload your SKU list to calculate container packing per SKU. Then use the
        Pallet Mixing section to calculate pallet equivalents, group mixed SKUs per pallet,
        and see total pallet savings.
        {!isPro && <span style={{color:'#c2410c'}}> Free plan: up to {CONFIG.freeSkuLimit} SKUs.</span>}
      </div>

      <div style={{display:'grid', gridTemplateColumns:'320px 1fr', gap:'20px', alignItems:'start'}}>

        {/* ── LEFT PANEL ─────────────────────────────────────────────────── */}
        <div>
          {/* Container */}
          <div style={S.card}>
            <div style={S.cardTitle}>🗃️ Container Details</div>
            <div style={S.grid2}>
              {[['Length',cL,setCL],['Width',cW,setCW],['Height',cH,setCH],['Max Weight (kg)',cMaxWt,setCMaxWt]].map(([l,v,s])=>(
                <div key={l}><label style={lbl}>{l}</label>
                  <input style={inp} type="number" min="0" step="any" value={v}
                    onChange={e=>s(e.target.value)} placeholder="0"/></div>))}
            </div>
            {/* Opt-in stability analysis */}
              <label style={{display:'flex',alignItems:'flex-start',gap:'8px',
                cursor:'pointer',background:stabCheck?'#f0fdf4':'#f8fafc',
                border:'1px solid '+(stabCheck?'#86efac':'#e2e8f0'),
                borderRadius:'8px',padding:'9px 12px',marginBottom:'12px'}}>
                <input type="checkbox" checked={stabCheck}
                  onChange={e=>setStabCheck(e.target.checked)}
                  style={{marginTop:'2px',width:'15px',height:'15px',cursor:'pointer',
                    accentColor:'#16a34a'}}/>
                <span>
                  <span style={{fontWeight:'700',fontSize:'13px',
                    color:stabCheck?'#166534':'#0f172a'}}>
                    Apply pallet stability check
                  </span>
                  <span style={{display:'block',fontSize:'11px',color:'#6b7280',marginTop:'2px'}}>
                    Adds a column checking whether each SKU will stand safely on the pallet
                    (height-to-base ratio and tipping angle), and recommends a better
                    orientation or pallet size where it will not. Leave unticked for the
                    standard results.
                  </span>
                </span>
              </label>
            {valid && <div style={S.infoBox}>Volume: {(container.cL*container.cW*container.cH).toLocaleString()} mm³</div>}
          </div>

          {/* SKU upload */}
          <div style={S.card}>
            <div style={S.cardTitle}>📂 SKU Data</div>
            <div style={S.dropzone(dragOver)}
              onDragOver={e=>{e.preventDefault();setDragOver(true)}}
              onDragLeave={()=>setDragOver(false)}
              onDrop={e=>{e.preventDefault();setDragOver(false);parseFile(e.dataTransfer.files[0])}}
              onClick={()=>document.getElementById('fi2').click()}>
              <div style={{fontSize:'28px',marginBottom:'6px'}}>📂</div>
              <div style={{fontSize:'13px',fontWeight:'500',color:'#374151'}}>{fileName||'Drop Excel or click to browse'}</div>
              <div style={{fontSize:'11px',color:'#9ca3af',marginTop:'4px'}}>.xlsx or .xls</div>
              {rawSkus&&<div style={{marginTop:'6px',fontSize:'12px',color:'#059669',fontWeight:'600'}}>✓ {skuCount.toLocaleString()} SKUs loaded</div>}
              <input id="fi2" type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={e=>parseFile(e.target.files[0])}/>
            </div>
            <div style={S.noteBox}><strong>Columns:</strong> SKU Name | L (mm) | W (mm) | H (mm) | Weight (kg) | Qty | <strong>Category</strong></div>
            <PasteFromExcel mode="bulk" onFill={(rows)=>{
              const skus = rows.map(r=>({name:r.name,sl:parseFloat(r.L)||0,sw:parseFloat(r.W)||0,
                sh:parseFloat(r.H)||0,swt:parseFloat(r.weight)||0,qtyAvail:parseFloat(r.qty)||0,
                category:String(r.category||r[6]||'').trim()||'Uncategorised'}));
              setRawSkus(skus); setSkuCount(skus.length); setFileName('');
            }}/>
          </div>
          <button style={S.btnPrimary} onClick={run}>▶ Calculate Container Fit</button>

          {/* ── PALLET MIXING INPUTS ──────────────────────────────────────── */}
          <div style={{...S.card, marginTop:'16px'}}>
            <div style={S.cardTitle}>🪵 Pallet Mixing Settings</div>

            {/* Pallet preset */}
            <div style={{marginBottom:'10px'}}>
              <label style={lbl}>Pallet Type</label>
              <select value={pPreset} onChange={e=>onPresetChange(e.target.value)}
                style={{...S.input, width:'100%'}}>
                {Object.entries(PALLET_PRESETS).map(([k,v])=>(
                  <option key={k} value={k}>{v.label}</option>))}
              </select>
            </div>

            {/* Pallet dimensions */}
            <div style={S.grid2}>
              {[['Pallet L (mm)',pL,setPL],['Pallet W (mm)',pW,setPW],['Pallet H (mm)',pH,setPH]].map(([l,v,s])=>(
                <div key={l}><label style={lbl}>{l}</label>
                  <input style={inp} type="number" min="0" value={v}
                    onChange={e=>{ setPPreset('custom'); s(e.target.value); }}
                    placeholder="0"/></div>))}

              {/* Are loads stretch-wrapped on the floor? */}
              {stabCheck && (
              <div>
                <label style={lbl}>Unit Load Wrapping</label>
                <select value={wrapped?'yes':'no'} onChange={e=>setWrapped(e.target.value==='yes')}
                  style={{...S.input}}>
                  <option value="yes">Stretch-wrapped as one block</option>
                  <option value="no">Not wrapped / loose units</option>
                </select>
                <div style={{fontSize:'10px',color:'#9ca3af',marginTop:'3px'}}>
                  Wrapping binds a layer into one stable block. Turn this off to see
                  which SKUs topple if the floor skips it.
                </div>
              </div>
              )}

              {/* Max SKUs per pallet */}
              <div>
                <label style={lbl}>Max SKUs / Pallet</label>
                <select value={maxSkus} onChange={e=>setMaxSkus(Number(e.target.value))}
                  style={{...S.input}}>
                  {[1,2,3,4].map(n=>(
                    <option key={n} value={n}>{n} SKU{n>1?'s':''} per pallet</option>))}
                </select>
              </div>
            </div>

            {/* Lock Height toggle */}
            <div style={{marginTop:'10px',padding:'10px 12px',background:'#f8fafc',
              border:'1px solid #e2e8f0',borderRadius:'8px',
              display:'flex',alignItems:'center',gap:'10px',cursor:'pointer'}}
              onClick={()=>setLockHeight(h=>!h)}>
              <div style={{width:'36px',height:'20px',borderRadius:'99px',position:'relative',
                background:lockHeight?'#7c3aed':'#d1d5db',transition:'background 0.2s',flexShrink:0}}>
                <div style={{position:'absolute',top:'2px',
                  left:lockHeight?'18px':'2px',width:'16px',height:'16px',
                  background:'#fff',borderRadius:'50%',transition:'left 0.2s',
                  boxShadow:'0 1px 3px rgba(0,0,0,0.2)'}}/>
              </div>
              <div>
                <div style={{fontWeight:'700',fontSize:'12px',color:lockHeight?'#7c3aed':'#374151'}}>
                  🔒 Lock Height (Keep Boxes Upright)
                </div>
                <div style={{fontSize:'10px',color:'#9ca3af',marginTop:'1px'}}>
                  {lockHeight ? 'Box H is always vertical — no tipping allowed' : 'All 6 orientations tried (max boxes per pallet)'}
                </div>
              </div>
            </div>

            {mixError && <div style={{...S.error, marginTop:'8px'}}>⚠ {mixError}</div>}
            <button style={{...S.btnPrimary, marginTop:'10px',
              background:'linear-gradient(135deg,#7c3aed,#6d28d9)'}}
              onClick={runMix}>
              🔀 Calculate Pallet Mixing
            </button>
            <div style={{fontSize:'11px',color:'#9ca3af',marginTop:'6px',textAlign:'center'}}>
              Requires SKU qty data · Works independently of container calculation
            </div>
          </div>
        </div>

        {/* ── RIGHT PANEL ────────────────────────────────────────────────── */}
        <div>
          {error && <div style={S.error}>⚠ {error}</div>}

          {capped && (
            <div style={{...S.card,background:'#fffbeb',border:'1px solid #fde68a',
              display:'flex',justifyContent:'space-between',alignItems:'center'}}>
              <span style={{fontSize:'13px',color:'#92400e'}}>
                ⚠ Free plan limited to {CONFIG.freeSkuLimit} SKUs. {skuCount.toLocaleString()} uploaded — only first {CONFIG.freeSkuLimit} processed.
              </span>
              <button onClick={onUpgrade} style={{padding:'6px 14px',background:'#059669',color:'#fff',
                border:'none',borderRadius:'8px',fontWeight:'600',fontSize:'12px',cursor:'pointer',whiteSpace:'nowrap'}}>
                ⭐ Go Pro
              </button>
            </div>)}

          {processing && (
            <div style={S.card}>
              <div style={{fontSize:'13px',fontWeight:'500',color:'#374151',marginBottom:'8px'}}>
                Processing {Math.min(skuCount,isPro?skuCount:CONFIG.freeSkuLimit).toLocaleString()} SKUs... {progress}%
              </div>
              <div style={{background:'#e5e7eb',borderRadius:'99px',height:'10px'}}>
                <div style={{height:'10px',borderRadius:'99px',background:'#059669',
                  width:`${progress}%`,transition:'width 0.2s'}}/>
              </div>
            </div>)}

          {/* Container results */}
          {results && !processing && (<>
            {/* Combined utilisation — weighted by quantity, not per SKU */}
            {combined && (
              <div style={{background:'#0f172a',borderRadius:'12px',padding:'14px 18px',
                marginBottom:'12px',display:'flex',flexWrap:'wrap',gap:'18px',
                alignItems:'center',justifyContent:'space-between'}}>
                <div>
                  <div style={{fontSize:'10px',color:'#94a3b8',fontWeight:'700',
                    textTransform:'uppercase',letterSpacing:'0.06em'}}>
                    Combined utilisation
                  </div>
                  <div style={{fontSize:'26px',fontWeight:'800',
                    color:combined.weighted>=0.75?'#4ade80':combined.weighted>=0.5?'#fbbf24':'#fb7185',
                    lineHeight:'1.15'}}>
                    {(combined.weighted*100).toFixed(1)}%
                  </div>
                  <div style={{fontSize:'11px',color:'#94a3b8',marginTop:'1px'}}>
                    weighted by quantity across {combined.n} SKU{combined.n>1?'s':''}
                  </div>
                </div>
                <div style={{display:'flex',gap:'18px',flexWrap:'wrap'}}>
                  {[['Goods volume',combined.goodsM3.toFixed(1)+' m\u00b3'],
                    ['Container volume',combined.spaceM3.toFixed(1)+' m\u00b3'],
                    ['Containers needed',combined.containers.toFixed(2)],
                    ['Unweighted SKU avg',(combined.simple*100).toFixed(1)+'%']].map(([l,v])=>(
                    <div key={l}>
                      <div style={{fontSize:'15px',fontWeight:'700',color:'#e2e8f0'}}>{v}</div>
                      <div style={{fontSize:'10px',color:'#64748b',marginTop:'1px'}}>{l}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {combined && Math.abs(combined.weighted-combined.simple) > 0.05 && (
              <div style={{fontSize:'11px',color:'#92400e',background:'#fffbeb',
                border:'1px solid #fde68a',borderRadius:'8px',padding:'7px 12px',
                marginBottom:'12px'}}>
                {combined.weighted > combined.simple
                  ? 'Your high-volume SKUs pack better than the SKU average suggests \u2014 the badly-packing lines are small ones.'
                  : 'Your high-volume SKUs pack worse than the SKU average suggests \u2014 focus on the largest lines first.'}
              </div>
            )}

            <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'12px',marginBottom:'16px'}}>
              {[['Total',bAll,'#f8fafc','#1a2332'],
                ['≥75%',bG,'#f0fdf4','#166534'],
                ['50–74%',bO,'#fefce8','#854d0e'],
                ['<50%',bL,'#fff8fc','#be185d']].map(([l,b,bg,col])=>(
                <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',
                  textAlign:'center',border:'1px solid rgba(0,0,0,0.06)'}}>
                  <div style={{fontSize:'20px',fontWeight:'700',color:col}}>{b.skus}</div>
                  <div style={{fontSize:'11px',color:'#6b7a8d',marginTop:'2px'}}>
                    {l}{l==='Total'?' SKUs':' \u00b7 SKUs'}
                  </div>
                  <div style={{borderTop:'1px solid rgba(0,0,0,0.07)',marginTop:'7px',paddingTop:'6px'}}>
                    <div style={{fontSize:'15px',fontWeight:'700',color:col}}>
                      {b.qty.toLocaleString()}
                    </div>
                    <div style={{fontSize:'10px',color:'#6b7a8d',marginTop:'1px'}}>
                      units{l!=='Total'&&bAll.qty>0?' \u00b7 '+pct(b.qty)+'% of qty':''}
                    </div>
                  </div>
                </div>))}
            </div>
            <button style={{...S.btnPrimary,marginBottom:'16px'}} onClick={exp}>
              ⬇ Download Results as Excel {mixResult?'(incl. Pallet Mixing)':''}
            </button>
            <div style={{...S.card,padding:'0',overflow:'hidden'}}>
              <div style={{padding:'12px 18px',borderBottom:'1px solid #f1f5f9',
                display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span style={{fontWeight:'600',fontSize:'13px'}}>Container Packing Preview</span>
                <span style={{fontSize:'12px',color:'#9ca3af'}}>{results.length.toLocaleString()} SKUs</span>
              </div>
              {/* Category lock toggles */}
              {results && (() => {
                const cats = [...new Set(results.filter(r=>!r.error).map(r=>r.category||'Uncategorised'))].sort();
                if (cats.length < 2) return null;
                return (
                  <div style={{padding:'10px 16px',borderBottom:'1px solid #e8edf2',
                    background:'#f8fafc',display:'flex',flexWrap:'wrap',gap:'6px',alignItems:'center'}}>
                    <span style={{fontSize:'11px',fontWeight:'700',color:'#374151',marginRight:'4px'}}>
                      🔒 Lock by Category:
                    </span>
                    {cats.map(cat=>{
                      const isLocked=lockedCategories.has(cat);
                      const skuCount=results.filter(r=>r.category===cat).length;
                      return(
                        <button key={cat} onClick={()=>toggleCategoryLock(cat)}
                          title={isLocked?`Unlock all ${skuCount} SKUs in "${cat}"`:`Lock height for all ${skuCount} SKUs in "${cat}"`}
                          style={{
                            padding:'4px 12px',borderRadius:'99px',cursor:'pointer',
                            fontFamily:'inherit',fontSize:'11px',fontWeight:'700',border:'none',
                            background:isLocked?'#ede9fe':'#e2e8f0',
                            color:isLocked?'#6d28d9':'#6b7280',
                            transition:'all 0.15s',
                          }}>
                          {isLocked?'🔒':'🔓'} {cat}
                          <span style={{marginLeft:'4px',opacity:0.7,fontWeight:'400'}}>({skuCount})</span>
                        </button>
                      );
                    })}
                    {lockedCategories.size>0&&(
                      <button onClick={()=>{setLockedCategories(new Set());}}
                        style={{padding:'4px 10px',borderRadius:'99px',cursor:'pointer',
                          fontFamily:'inherit',fontSize:'10px',fontWeight:'600',border:'none',
                          background:'#fee2e2',color:'#be185d',marginLeft:'auto'}}>
                        ✕ Clear all category locks
                      </button>
                    )}
                  </div>
                );
              })()}
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                  <thead><tr>
                    {['SKU','Category',...(stabCheck?['Pallet Stability']:[]),'Vol Qty','Wt Qty','Eff Qty','Vol%','Wt%','→ L-axis','→ W-axis','→ H-axis (stack)','Layers','Constraint','🔒 Lock H'].map(h=>(
                      <th key={h} style={{padding:'9px 12px',textAlign:'left',fontWeight:'600',
                        fontSize:'11px',color:'#6b7a8d',textTransform:'uppercase',
                        background:'#f8fafc',borderBottom:'1px solid #e8edf2',whiteSpace:'nowrap'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {results.slice(0,100).map((r,i)=>r.error
                      ? (<tr key={i} style={{background:'#fff8fc'}}>
                          <td style={{padding:'8px 12px',fontWeight:'500'}}>{r.name}</td>
                          <td style={{padding:'8px 12px',color:'#be185d'}} colSpan={7}>{r.error}</td>
                        </tr>)
                      : (<tr key={i} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                          <td style={{padding:'8px 12px',fontWeight:'500'}}>{r.name}</td>
                          <td style={{padding:'8px 12px'}}>
                            <span style={{background:'#f1f5f9',color:'#475569',
                              padding:'2px 8px',borderRadius:'99px',fontSize:'10px',
                              fontWeight:'600',whiteSpace:'nowrap'}}>
                              {r.category||'—'}
                            </span>
                          </td>
                          {stabCheck && (
                          <td style={{padding:'8px 12px',minWidth:'176px'}}>
                            {r.stab ? (
                              <div style={{borderRadius:'7px',padding:'4px 8px',fontSize:'11px',
                                background:r.stab.wrapCritical?'#fffbeb':r.stab.ok?'#f0fdf4':r.stab.rating==='marginal'?'#fffbeb':'#fff1f2',
                                border:'1px solid '+(r.stab.wrapCritical?'#fbbf24':r.stab.ok?'#86efac':r.stab.rating==='marginal'?'#fde68a':'#fecdd3'),
                                color:r.stab.wrapCritical?'#92400e':r.stab.ok?'#166534':r.stab.rating==='marginal'?'#92400e':'#be185d'}}>
                                <div style={{fontWeight:'700'}}>
                                  {r.stab.wrapCritical&&r.stab.ok ? 'Stable \u2014 wrap required'
                                    : r.stab.ok ? 'Stable'
                                    : r.stab.rating==='marginal' ? 'Needs securing' : 'Unstable'}
                                  {r.stab.best&&r.stab.best.usable
                                    ? ' \u00b7 '+(r.stab.assumeWrapped?r.stab.best.slender:r.stab.best.slenderSingle)+':1 \u00b7 '+(r.stab.assumeWrapped?r.stab.best.tip:r.stab.best.tipSingle)+'\u00b0'
                                    : ' \u00b7 does not fit'}
                                </div>
                                <div style={{fontSize:'10px',fontWeight:'400',marginTop:'1px',color:'#6b7280'}}>
                                  {r.stab.best&&r.stab.best.usable?r.stab.best.name:''}
                                  {r.stab.best&&r.stab.best.usable&&r.stab.best.perLayer>1
                                    ? ' \u00b7 block of '+r.stab.best.perLayer
                                    : ''}
                                </div>
                                {r.stab.best&&r.stab.best.usable&&r.stab.best.slenderSingle>r.stab.best.slender+0.05&&(
                                  <div style={{fontSize:'9px',fontWeight:'400',color:'#9ca3af',marginTop:'1px'}}>
                                    single unit {r.stab.best.slenderSingle}:1 \u2014 wrap the layer together
                                  </div>
                                )}
                                {(!r.stab.ok || r.stab.wrapCritical) && (
                                  <div style={{fontSize:'10px',fontWeight:'400',marginTop:'3px',lineHeight:'1.35'}}>
                                    {r.stab.advice}
                                  </div>
                                )}
                              </div>
                            ) : <span style={{color:'#9ca3af',fontSize:'11px'}}>-</span>}
                          </td>
                          )}
                          <td style={{padding:'8px 12px',textAlign:'right'}}>{r.volQty?.toLocaleString()}</td>
                          <td style={{padding:'8px 12px',textAlign:'right'}}>{typeof r.wtQty==='number'?r.wtQty.toLocaleString():r.wtQty}</td>
                          <td style={{padding:'8px 12px',textAlign:'right',fontWeight:'700'}}>{r.effQty?.toLocaleString()}</td>
                          <td style={{padding:'8px 12px'}}><UtilBadge val={r.volUtil}/></td>
                          <td style={{padding:'8px 12px'}}><UtilBadge val={r.wtUtil}/></td>
                          {(()=>{
                            const [dL,dW,dH]=(r.orient||'').split('×');
                            const sl=String(r.name&&r.sl||''), sw2=String(r.sw||''), sh=String(r.sh||'');
                            // highlight which cell matches the SKU's original L/W/H
                            const cell=(val,orig,axis)=>(
                              <td key={axis} style={{padding:'6px 10px',textAlign:'center',whiteSpace:'nowrap'}}>
                                <span style={{
                                  display:'inline-block',
                                  background: val===orig?'#eff6ff':'#f8fafc',
                                  color: val===orig?'#1d4ed8':'#6b7280',
                                  border: `1px solid ${val===orig?'#bfdbfe':'#e2e8f0'}`,
                                  borderRadius:'6px',padding:'2px 8px',
                                  fontSize:'11px',fontWeight: val===orig?'700':'400'}}>
                                  {val||'—'}<span style={{fontSize:'9px',color:'#9ca3af'}}> mm</span>
                                </span>
                              </td>);
                            return(<>{cell(dL,sl,'L')}{cell(dW,sw2,'W')}{cell(dH,sh,'H')}</>);
                          })()}
                          <td style={{padding:'8px 12px',textAlign:'center'}}>
                            {r.stackLayers>0&&(
                              <span style={{
                                background:r.heightLocked?'#ede9fe':'#eff6ff',
                                color:r.heightLocked?'#6d28d9':'#1d4ed8',
                                padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'700'}}>
                                {r.stackLayers}{r.heightLocked?' 🔒':''}L
                              </span>
                            )}
                          </td>
                          <td style={{padding:'8px 12px'}}>
                            <span style={{padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'500',
                              background:r.constraint==='Volume'?'#eff6ff':r.constraint==='Weight'?'#fff7ed':'#f5f3ff',
                              color:r.constraint==='Volume'?'#1d4ed8':r.constraint==='Weight'?'#c2410c':'#6d28d9'}}>
                              {r.constraint}
                            </span>
                          </td>
                          <td style={{padding:'8px 12px',textAlign:'center'}}>
                            <button
                              onClick={()=>toggleContainerLock(r.name)}
                              title={lockedSkus.has(r.name)
                                ? 'Height locked — box stays upright. Click to unlock.'
                                : 'Click to lock height — box H stays vertical (no tipping)'}
                              style={{
                                padding:'3px 8px',borderRadius:'99px',cursor:'pointer',
                                fontFamily:'inherit',fontSize:'11px',fontWeight:'700',
                                border:'none',transition:'all 0.15s',
                                background:lockedSkus.has(r.name)?'#ede9fe':'#f1f5f9',
                                color:lockedSkus.has(r.name)?'#6d28d9':'#9ca3af',
                              }}>
                              {lockedSkus.has(r.name)?'🔒':'🔓'}
                            </button>
                          </td>
                        </tr>))}
                  </tbody>
                </table>
                {results.length>100&&<div style={{padding:'10px 18px',fontSize:'12px',color:'#9ca3af',borderTop:'1px solid #f1f5f9'}}>
                  Showing 100 of {results.length.toLocaleString()} — download Excel for all
                </div>}
              </div>
            </div>
          </>)}

          {/* ── PALLET MIXING RESULTS ───────────────────────────────────── */}
          {mixResult && (
            <div style={{marginTop:'20px'}}>
              {/* Summary comparison */}
              <div style={{...S.card, background:'linear-gradient(135deg,#f5f3ff,#eff6ff)'}}>
                <div style={S.cardTitle}>🪵 Pallet Mixing Summary</div>
                <div style={{fontSize:'12px',color:'#6b7280',marginBottom:'14px'}}>
                  Pallet: {pL}×{pW}×{pH}mm · Max {maxSkus} SKU{maxSkus>1?'s':''} per pallet{lockedSkus.size>0?` · ${lockedSkus.size} SKU${lockedSkus.size>1?'s':''} height-locked`:''}
                  {lockHeight ? <span style={{marginLeft:'8px',background:'#ede9fe',color:'#6d28d9',
                    padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'600'}}>
                    🔒 Height Locked</span> : ''}
                </div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'12px',marginBottom:'16px'}}>
                  {[
                    ['Without Mixing', mixResult.totalBefore, 'Each SKU on separate pallets','#fff1f2','#be185d'],
                    ['With Mixing',    mixResult.totalAfter,  `Up to ${maxSkus} SKUs share a pallet`,'#f0fdf4','#166534'],
                    ['Pallets Saved',  mixResult.savings,     mixResult.totalBefore>0?((mixResult.savings/mixResult.totalBefore)*100).toFixed(1)+'% reduction':'','#fffbeb','#d97706'],
                  ].map(([l,v,sub,bg,col])=>(
                    <div key={l} style={{background:bg,borderRadius:'10px',padding:'14px',textAlign:'center',border:`1px solid ${col}22`}}>
                      <div style={{fontSize:'26px',fontWeight:'800',color:col}}>{v}</div>
                      <div style={{fontSize:'11px',fontWeight:'700',color:'#374151',marginTop:'3px'}}>{l}</div>
                      <div style={{fontSize:'10px',color:'#9ca3af',marginTop:'2px'}}>{sub}</div>
                    </div>))}
                </div>

                {/* Breakdown */}
                <div style={{display:'flex',gap:'10px',fontSize:'12px',color:'#6b7280',flexWrap:'wrap'}}>
                  <span>📦 {mixResult.totalFull} full dedicated pallets</span>
                  <span>·</span>
                  <span>🔀 {mixResult.mixedPallets.length} mixed pallets (from fractional remainders)</span>
                </div>
              </div>

              {/* Per-SKU pallet equivalents */}
              <div style={{...S.card,padding:'0',overflow:'hidden',marginTop:'12px'}}>
                <div style={{padding:'12px 18px',borderBottom:'1px solid #f1f5f9',
                  display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                  <span style={{fontWeight:'700',fontSize:'13px'}}>Per-SKU Pallet Equivalents</span>
                  <span style={{fontSize:'12px',color:'#9ca3af'}}>{mixResult.items.length} SKUs</span>
                </div>
                <div style={{overflowX:'auto',maxHeight:'320px',overflowY:'auto'}}>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                    <thead><tr>
                      {['SKU','Qty Avail','Boxes/Pallet','Pallet Equiv','Full Pallets','Remainder','Layers','Status','🔒 Lock H'].map(h=>(
                        <th key={h} style={{padding:'8px 12px',textAlign:'left',fontWeight:'600',
                          fontSize:'11px',color:'#6b7a8d',textTransform:'uppercase',
                          background:'#f8fafc',borderBottom:'1px solid #e8edf2',
                          whiteSpace:'nowrap',position:'sticky',top:0}}>{h}</th>))}
                    </tr></thead>
                    <tbody>
                      {mixResult.items.map((r,i)=>(
                        <tr key={i} style={{background:i%2===0?'#fff':'#fafbfc'}}>
                          <td style={{padding:'7px 12px',fontWeight:'600'}}>{r.name}</td>
                          <td style={{padding:'7px 12px',textAlign:'right'}}>{r.qtyAvail.toLocaleString()}</td>
                          <td style={{padding:'7px 12px',textAlign:'right',color:'#6b7280'}}>{r.bpp||'—'}</td>
                          <td style={{padding:'7px 12px',textAlign:'right',fontWeight:'700',color:'#7c3aed'}}>
                            {r.palletEquiv!=null?r.palletEquiv:'—'}
                          </td>
                          <td style={{padding:'7px 12px',textAlign:'right'}}>{r.fullPallets}</td>
                          <td style={{padding:'7px 12px',textAlign:'right',
                            color:r.remainder>0?'#d97706':'#9ca3af',fontWeight:r.remainder>0?'600':'400'}}>
                            {r.remainder>0?r.remainder.toFixed(4):'0'}
                          </td>
                          <td style={{padding:'7px 12px',textAlign:'center'}}>
                            {r.stackLayers>0
                              ? <span style={{background:r.heightLocked?'#ede9fe':'#eff6ff',
                                  color:r.heightLocked?'#6d28d9':'#1d4ed8',
                                  padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'700'}}>
                                  {r.stackLayers}{r.heightLocked?' 🔒':''}
                                </span>
                              : <span style={{color:'#9ca3af'}}>—</span>}
                          </td>
                          <td style={{padding:'7px 12px'}}>
                            {r.error
                              ? <span style={{fontSize:'11px',color:'#be185d'}}>⚠ {r.error}</span>
                              : r.remainder>0
                                ? <span style={{background:'#fef9c3',color:'#854d0e',padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'600'}}>Mixed</span>
                                : r.fullPallets>0
                                  ? <span style={{background:'#f0fdf4',color:'#166534',padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'600'}}>Full</span>
                                  : <span style={{color:'#9ca3af',fontSize:'11px'}}>—</span>}
                          </td>
                          <td style={{padding:'7px 12px',textAlign:'center'}}>
                            <button
                              onClick={()=>toggleSkuLock(r.name)}
                              title={lockedSkus.has(r.name)||r.heightLocked
                                ? 'Height locked — click to unlock (allow tipping)'
                                : 'Click to lock height upright for this SKU'}
                              style={{
                                padding:'3px 8px',borderRadius:'99px',cursor:'pointer',
                                fontFamily:'inherit',fontSize:'11px',fontWeight:'700',border:'none',
                                background:lockedSkus.has(r.name)?'#ede9fe':'#f1f5f9',
                                color:lockedSkus.has(r.name)?'#6d28d9':'#9ca3af',
                                transition:'all 0.15s',
                              }}>
                              {lockedSkus.has(r.name)?'🔒':'🔓'}
                            </button>
                          </td>
                        </tr>))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Mixed pallet groups */}
              {mixResult.mixedPallets.length > 0 && (
                <div style={{...S.card,padding:'0',overflow:'hidden',marginTop:'12px'}}>
                  <div style={{padding:'12px 18px',borderBottom:'1px solid #f1f5f9',
                    display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                    <span style={{fontWeight:'700',fontSize:'13px'}}>Mixed Pallet Groups</span>
                    <span style={{fontSize:'12px',color:'#9ca3af'}}>{mixResult.mixedPallets.length} pallets</span>
                  </div>
                  <div style={{overflowX:'auto',maxHeight:'360px',overflowY:'auto'}}>
                    <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                      <thead><tr>
                        {['Pallet #','SKUs','Utilisation','Size Fitment','SKU Names'].map(h=>(
                          <th key={h} style={{padding:'8px 12px',textAlign:'left',fontWeight:'600',
                            fontSize:'11px',color:'#6b7a8d',textTransform:'uppercase',
                            background:'#f5f3ff',borderBottom:'1px solid #e8edf2',
                            whiteSpace:'nowrap',position:'sticky',top:0}}>{h}</th>))}
                      </tr></thead>
                      <tbody>
                        {mixResult.mixedPallets.map((p,i)=>{
                          const fit = p.fitment||{};
                          const fitOk = fit.feasible !== false;
                          const fitWarn = fitOk && !!fit.warning;
                          const border = !fitOk?'3px solid #be185d':fitWarn?'3px solid #d97706':'3px solid #16a34a';
                          const fitBg  = !fitOk?'#fff1f2':fitWarn?'#fffbeb':'#f0fdf4';
                          const fitCol = !fitOk?'#be185d':fitWarn?'#d97706':'#166534';
                          return(
                          <tr key={i} style={{background:i%2===0?'#faf8ff':'#f5f3ff',
                            borderLeft:border}}>
                            <td style={{padding:'7px 12px',fontWeight:'700',color:'#7c3aed'}}>Mixed {i+1}</td>
                            <td style={{padding:'7px 12px',textAlign:'center'}}>
                              <span style={{background:'#7c3aed',color:'#fff',borderRadius:'99px',
                                padding:'2px 10px',fontWeight:'700',fontSize:'12px'}}>
                                {p.skus.length}
                              </span>
                            </td>
                            <td style={{padding:'7px 12px'}}>
                              <div style={{display:'flex',alignItems:'center',gap:'8px'}}>
                                <div style={{flex:1,background:'#e9d5ff',borderRadius:'99px',height:'8px'}}>
                                  <div style={{height:'8px',borderRadius:'99px',
                                    background:p.used>=0.8?'#7c3aed':p.used>=0.5?'#a78bfa':'#c4b5fd',
                                    width:`${Math.min(p.used*100,100)}%`}}/>
                                </div>
                                <span style={{fontWeight:'700',fontSize:'12px',
                                  color:p.used>=0.8?'#6d28d9':'#9ca3af',minWidth:'36px'}}>
                                  {(p.used*100).toFixed(1)}%
                                </span>
                              </div>
                            </td>
                            {/* Size fitment status */}
                            <td style={{padding:'6px 10px',minWidth:'140px'}}>
                              <div style={{background:fitBg,borderRadius:'7px',
                                padding:'5px 8px',fontSize:'11px',color:fitCol,fontWeight:'600'}}>
                                <div>
                                  {!fitOk?'❌ Infeasible':fitWarn?'⚠️ Unstable':'✅ Compatible'}
                                </div>
                                {!fitOk&&fit.reason&&
                                  <div style={{fontSize:'10px',marginTop:'2px',fontWeight:'400'}}>
                                    {fit.reason}
                                  </div>}
                                {fitWarn&&fit.warning&&
                                  <div style={{fontSize:'10px',marginTop:'2px',fontWeight:'400'}}>
                                    {fit.warning}
                                  </div>}
                                {fitOk&&!fitWarn&&fit.totalH!=null&&
                                  <div style={{fontSize:'10px',marginTop:'2px',fontWeight:'400',color:'#166534'}}>
                                    H: {fit.totalH}mm · Floor: {fit.totalFloor}%
                                  </div>}
                              </div>
                            </td>
                            <td style={{padding:'7px 12px',color:'#374151'}}>
                              {p.skus.map(s=>(
                                <span key={s.name} style={{display:'inline-block',
                                  background:'#ede9fe',color:'#6d28d9',
                                  borderRadius:'6px',padding:'2px 8px',
                                  fontSize:'11px',fontWeight:'600',margin:'2px 3px 2px 0'}}>
                                  {s.name}
                                  <span style={{fontWeight:'400',color:'#9ca3af',marginLeft:'4px'}}>
                                    ({(s.remainder*100).toFixed(1)}%)
                                    {s.sl&&s.sw&&s.sh?` ${s.sl}×${s.sw}×${s.sh}`:''}
                                  </span>
                                  {(()=>{ const st=stabCheck ? assessStability(s.sl,s.sw,s.sh,pL,pW,pH,s.qtyAvail,wrapped) : null;
                                    if(!st||st.ok) return null;
                                    return <span title={st.advice} style={{marginLeft:'4px',
                                      color:st.rating==='marginal'?'#d97706':'#be185d',
                                      fontWeight:'700'}}>
                                      {st.rating==='marginal'?'\u26a0':'\u26a0\u26a0'}
                                    </span>; })()}
                                </span>))}
                            </td>
                          </tr>);})}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {!results && !processing && !error && (
            <div style={{...S.card,padding:'60px',textAlign:'center',color:'#9ca3af'}}>
              <div style={{fontSize:'48px',marginBottom:'12px'}}>📦</div>
              <div style={{fontWeight:'500'}}>Fill in container details and upload your SKU file to get started</div>
            </div>)}
        </div>
      </div>
    </div>
  );
}
