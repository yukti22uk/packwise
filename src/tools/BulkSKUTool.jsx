// ─── BULK SKU CALCULATOR ─────────────────────────────────────────────────────
import { useState } from 'react';
import * as XLSX from 'xlsx';
import { CONFIG } from '../config.js';
import { calcMixed } from '../algorithms/packing.js';
import { S, UtilBadge } from '../components/styles.jsx';
import PasteFromExcel from '../components/PasteFromExcel.jsx';

// ─── PALLET PRESETS ───────────────────────────────────────────────────────────
const PALLET_PRESETS = {
  standard: { label: 'Standard (1200×1000×1200)', L:1200, W:1000, H:1200 },
  euro:     { label: 'Euro Pallet (1200×800×1200)',L:1200, W:800,  H:1200 },
  half:     { label: 'Half Pallet (600×800×1200)', L:600,  W:800,  H:1200 },
  custom:   { label: 'Custom',                     L:0,    W:0,    H:0    },
};

// ─── PALLET MIXING ALGORITHM ──────────────────────────────────────────────────
function calcPalletMix(skus, pL, pW, pH, maxSkus) {
  // Step 1: calc boxes per pallet and pallet equivalents per SKU
  const items = skus
    .filter(s => s.sl>0 && s.sw>0 && s.sh>0 && s.qtyAvail>0)
    .map(s => {
      const { total: bpp } = calcMixed(pL, pW, pH, s.sl, s.sw, s.sh);
      if (!bpp || bpp === 0) return { ...s, bpp:0, palletEquiv:null, fullPallets:0, remainder:0, error:'Box too large for pallet' };
      const pe     = s.qtyAvail / bpp;
      const full   = Math.floor(pe);
      const rem    = +(pe - full).toFixed(6);
      return { ...s, bpp, palletEquiv:+pe.toFixed(4), fullPallets:full, remainder:rem };
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

  return { items, mixedPallets, totalFull, totalBefore, totalAfter, savings };
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
  const [maxSkus,     setMaxSkus]     = useState(4);
  const [mixResult,   setMixResult]   = useState(null);
  const [mixError,    setMixError]    = useState('');

  const container = {
    cL:parseFloat(cL)||0, cW:parseFloat(cW)||0,
    cH:parseFloat(cH)||0, cMaxWt:parseFloat(cMaxWt)||0
  };
  const valid = container.cL>0 && container.cW>0 && container.cH>0 && container.cMaxWt>0;

  // ── Per-SKU container packing ──────────────────────────────────────────────
  function pSkus(cont, skus) {
    const { cL, cW, cH, cMaxWt } = cont;
    const cv = cL * cW * cH;
    return skus.map(s => {
      const { name, sl, sw, sh, swt, qtyAvail } = s;
      if (!name) return null;
      if (sl<=0||sw<=0||sh<=0) return { name, error:'Invalid dimensions' };
      const { total:vQ, orient } = calcMixed(cL, cW, cH, sl, sw, sh);
      let eV = qtyAvail>0 ? Math.min(vQ, qtyAvail) : vQ;
      let wQ = swt>0 ? Math.floor(cMaxWt/swt) : null;
      if (wQ!==null && qtyAvail>0) wQ = Math.min(wQ, qtyAvail);
      const eQ = wQ!==null ? Math.min(eV,wQ) : eV;
      const vu = (eQ*sl*sw*sh)/cv;
      const wu = swt>0 ? (eQ*swt)/cMaxWt : null;
      let con = 'Volume';
      if (wQ!==null && wQ<eV) con = 'Weight';
      if (qtyAvail>0 && eQ===qtyAvail) con = 'Stock Limit';
      return { name, volQty:eV, wtQty:wQ!==null?wQ:'N/A', effQty:eQ, volUtil:vu, wtUtil:wu, orient, constraint:con };
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
            swt:parseFloat(rr[4])||0, qtyAvail:parseFloat(rr[5])>0?parseFloat(rr[5]):-1 });
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
      all.push(...pSkus(container, toProcess.slice(done, done+CHUNK)));
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
    const mr = calcPalletMix(skusForMix, PL, PW, PH, maxSkus);
    setMixResult(mr);
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
    const h = ['SKU Name','Max Qty (Volume)','Max Qty (Weight)','Effective Max Qty','Volume Used (%)','Weight Used (%)','Best Orientation','Constraint'];
    const rows = results.map(r => r.error
      ? [r.name,r.error,'','','','','','']
      : [r.name,r.volQty,r.wtQty,r.effQty,(r.volUtil*100).toFixed(2)+'%',
         r.wtUtil!=null?(r.wtUtil*100).toFixed(2)+'%':'',r.orient,r.constraint]);
    const ws = XLSX.utils.aoa_to_sheet([
      ['CONTAINER SKU PACKING RESULTS'],[],
      ['Container',`${container.cL}×${container.cW}×${container.cH}`,'Max Weight',container.cMaxWt],[],h,...rows]);
    ws['!cols']=[{wch:22},{wch:16},{wch:16},{wch:16},{wch:14},{wch:14},{wch:26},{wch:14}];
    XLSX.utils.book_append_sheet(wb,ws,'Container Results');

    if (mixResult) {
      const mixRows = [
        ['PALLET MIXING RESULTS'],
        [`Pallet: ${pL}×${pW}×${pH}mm | Max SKUs per pallet: ${maxSkus}`],[],
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

  const gC=results?results.filter(r=>!r.error&&r.volUtil>=0.75).length:0;
  const oC=results?results.filter(r=>!r.error&&r.volUtil>=0.5&&r.volUtil<0.75).length:0;
  const lC=results?results.filter(r=>!r.error&&r.volUtil<0.5).length:0;

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
            <div style={S.noteBox}><strong>Columns:</strong> SKU Name | L (mm) | W (mm) | H (mm) | Weight (kg) | Qty</div>
            <PasteFromExcel mode="bulk" onFill={(rows)=>{
              const skus = rows.map(r=>({name:r.name,sl:parseFloat(r.L)||0,sw:parseFloat(r.W)||0,
                sh:parseFloat(r.H)||0,swt:parseFloat(r.weight)||0,qtyAvail:parseFloat(r.qty)||0}));
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
            <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'12px',marginBottom:'16px'}}>
              {[['Total',results.length,'#f8fafc'],['≥75%',gC,'#f0fdf4'],['50–74%',oC,'#fefce8'],['<50%',lC,'#fff8fc']].map(([l,v,bg])=>(
                <div key={l} style={{background:bg,borderRadius:'10px',padding:'12px',
                  textAlign:'center',border:'1px solid rgba(0,0,0,0.06)'}}>
                  <div style={{fontSize:'20px',fontWeight:'700',color:'#1a2332'}}>{v}</div>
                  <div style={{fontSize:'11px',color:'#6b7a8d',marginTop:'2px'}}>{l}</div>
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
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                  <thead><tr>
                    {['SKU','Vol Qty','Wt Qty','Eff Qty','Vol%','Wt%','Orientation','Constraint'].map(h=>(
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
                          <td style={{padding:'8px 12px',textAlign:'right'}}>{r.volQty?.toLocaleString()}</td>
                          <td style={{padding:'8px 12px',textAlign:'right'}}>{typeof r.wtQty==='number'?r.wtQty.toLocaleString():r.wtQty}</td>
                          <td style={{padding:'8px 12px',textAlign:'right',fontWeight:'700'}}>{r.effQty?.toLocaleString()}</td>
                          <td style={{padding:'8px 12px'}}><UtilBadge val={r.volUtil}/></td>
                          <td style={{padding:'8px 12px'}}><UtilBadge val={r.wtUtil}/></td>
                          <td style={{padding:'8px 12px',color:'#6b7a8d',whiteSpace:'nowrap'}}>{r.orient}</td>
                          <td style={{padding:'8px 12px'}}>
                            <span style={{padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'500',
                              background:r.constraint==='Volume'?'#eff6ff':r.constraint==='Weight'?'#fff7ed':'#f5f3ff',
                              color:r.constraint==='Volume'?'#1d4ed8':r.constraint==='Weight'?'#c2410c':'#6d28d9'}}>
                              {r.constraint}
                            </span>
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
                  Pallet: {pL}×{pW}×{pH}mm · Max {maxSkus} SKU{maxSkus>1?'s':''} per pallet
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
                      {['SKU','Qty Available','Boxes/Pallet','Pallet Equivalents','Full Pallets','Remainder','Status'].map(h=>(
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
                          <td style={{padding:'7px 12px'}}>
                            {r.error
                              ? <span style={{fontSize:'11px',color:'#be185d'}}>⚠ {r.error}</span>
                              : r.remainder>0
                                ? <span style={{background:'#fef9c3',color:'#854d0e',padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'600'}}>Mixed</span>
                                : r.fullPallets>0
                                  ? <span style={{background:'#f0fdf4',color:'#166534',padding:'2px 8px',borderRadius:'99px',fontSize:'11px',fontWeight:'600'}}>Full</span>
                                  : <span style={{color:'#9ca3af',fontSize:'11px'}}>—</span>}
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
                  <div style={{overflowX:'auto',maxHeight:'300px',overflowY:'auto'}}>
                    <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                      <thead><tr>
                        {['Pallet #','SKUs Mixed','Utilisation','SKU Names'].map(h=>(
                          <th key={h} style={{padding:'8px 12px',textAlign:'left',fontWeight:'600',
                            fontSize:'11px',color:'#6b7a8d',textTransform:'uppercase',
                            background:'#f5f3ff',borderBottom:'1px solid #e8edf2',
                            whiteSpace:'nowrap',position:'sticky',top:0}}>{h}</th>))}
                      </tr></thead>
                      <tbody>
                        {mixResult.mixedPallets.map((p,i)=>(
                          <tr key={i} style={{background:i%2===0?'#faf8ff':'#f5f3ff'}}>
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
                                    background: p.used>=0.8?'#7c3aed':p.used>=0.5?'#a78bfa':'#c4b5fd',
                                    width:`${Math.min(p.used*100,100)}%`}}/>
                                </div>
                                <span style={{fontWeight:'700',fontSize:'12px',
                                  color:p.used>=0.8?'#6d28d9':'#9ca3af',minWidth:'36px'}}>
                                  {(p.used*100).toFixed(1)}%
                                </span>
                              </div>
                            </td>
                            <td style={{padding:'7px 12px',color:'#374151'}}>
                              {p.skus.map(s=>(
                                <span key={s.name} style={{display:'inline-block',background:'#ede9fe',
                                  color:'#6d28d9',borderRadius:'6px',padding:'2px 8px',
                                  fontSize:'11px',fontWeight:'600',margin:'2px 3px 2px 0'}}>
                                  {s.name} ({(s.remainder*100).toFixed(1)}%)
                                </span>))}
                            </td>
                          </tr>))}
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
