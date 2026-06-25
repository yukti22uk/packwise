// ─── ORDER ANALYSER TOOL ──────────────────────────────────────────────────────
// Step 1: Upload Master SKU data → AI maps columns → validate + flag anomalies
// Step 2: Upload Order data → AI maps columns → full analytics → 6-sheet Excel
import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { S } from '../components/styles.jsx';

// ─── CLAUDE API CALL ─────────────────────────────────────────────────────────
async function claudeCall(prompt, maxTokens = 1500) {
  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: maxTokens,
      messages: [{ role: 'user', content: prompt }] }),
  });
  if (!r.ok) throw new Error(`API ${r.status}`);
  const d = await r.json();
  const text = d.content?.[0]?.text || '';
  const m = text.match(/\{[\s\S]*\}/);
  if (!m) throw new Error('AI format error');
  return JSON.parse(m[0]);
}

// ─── PARSE EXCEL HELPER ───────────────────────────────────────────────────────
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        resolve({ rows, sheets: wb.SheetNames });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Map header names returned by Claude to 0-based column indices
function buildColIdx(headerRow, colMapResult) {
  const idx = {};
  headerRow.forEach((h, i) => { idx[String(h).toLowerCase().trim()] = i; });
  const resolve = key => {
    const val = colMapResult[key];
    if (val === null || val === undefined || val === 'null') return null;
    const found = idx[String(val).toLowerCase().trim()];
    return found !== undefined ? found : null;
  };
  return resolve;
}

// ─── MASTER SKU: DETECT ANOMALIES ────────────────────────────────────────────
function detectMasterAnomalies(dataRows, ci) {
  const anomalies = [], seen = new Map();
  dataRows.forEach((row, i) => {
    const rn = i + 2;
    const sku = String(row[ci.sku_name] ?? '').trim();
    if (!sku) { anomalies.push({ row: rn, sku: '—', field: 'SKU Name', issue: 'Missing SKU name/code', sev: 'High' }); return; }
    if (seen.has(sku)) anomalies.push({ row: rn, sku, field: 'SKU', issue: `Duplicate (first at row ${seen.get(sku)})`, sev: 'Medium' });
    else seen.set(sku, rn);
    ['length','width','height'].forEach(dim => {
      if (ci[dim] === null) return;
      const v = parseFloat(row[ci[dim]]);
      if (isNaN(v) || v === 0) anomalies.push({ row: rn, sku, field: dim, issue: `Missing or zero ${dim}`, sev: 'High' });
      else if (v < 0) anomalies.push({ row: rn, sku, field: dim, issue: `Negative ${dim}: ${v}`, sev: 'High' });
    });
    if (ci.weight !== null) {
      const w = parseFloat(row[ci.weight]);
      if (isNaN(w) || w === 0) anomalies.push({ row: rn, sku, field: 'Weight', issue: 'Missing or zero weight', sev: 'Medium' });
      else if (w < 0) anomalies.push({ row: rn, sku, field: 'Weight', issue: `Negative weight: ${w}`, sev: 'High' });
    }
  });
  return anomalies;
}

// Build master SKU lookup map
function buildMasterMap(dataRows, ci, unit) {
  const mult = unit === 'cm' ? 10 : unit === 'm' ? 1000 : 1;
  const map = new Map();
  dataRows.forEach(row => {
    const sku = String(row[ci.sku_name] ?? '').trim();
    if (!sku) return;
    const L = (parseFloat(row[ci.length]) || 0) * mult;
    const W = (parseFloat(row[ci.width])  || 0) * mult;
    const H = (parseFloat(row[ci.height]) || 0) * mult;
    const wt = ci.weight !== null ? parseFloat(row[ci.weight]) || 0 : 0;
    map.set(sku, { L, W, H, weight: wt, volume: L * W * H });
  });
  return map;
}

// ─── ORDER ANALYTICS ─────────────────────────────────────────────────────────
function runOrderAnalytics(dataRows, ci, masterMap) {
  // Parse rows
  const orders = dataRows.map((row, i) => ({
    rn: i + 2,
    orderNo:   String(row[ci.order_no]   ?? '').trim(),
    orderType: String(row[ci.order_type] ?? '').trim() || 'Unknown',
    sku:       String(row[ci.sku]        ?? '').trim(),
    qty:       parseFloat(row[ci.qty])   || 0,
    date:      String(row[ci.date]       ?? '').trim(),
    customer:  ci.customer  !== null ? String(row[ci.customer]  ?? '').trim() : '',
    lineNo:    ci.line_no   !== null ? row[ci.line_no] : null,
  }));

  // ── Anomalies ────────────────────────────────────────────────────────────────
  const anomalies = [];
  orders.forEach(o => {
    if (!o.orderNo) anomalies.push({ row: o.rn, sku: o.sku, field: 'Order No', issue: 'Missing order/invoice number', sev: 'High' });
    if (!o.sku)     anomalies.push({ row: o.rn, sku: '—',    field: 'SKU',      issue: 'Missing SKU code', sev: 'High' });
    if (o.qty <= 0) anomalies.push({ row: o.rn, sku: o.sku,  field: 'Qty',      issue: `Zero or negative qty: ${o.qty}`, sev: 'High' });
    if (!o.date)    anomalies.push({ row: o.rn, sku: o.sku,  field: 'Date',     issue: 'Missing date', sev: 'Medium' });
    if (o.sku && !masterMap.has(o.sku)) anomalies.push({ row: o.rn, sku: o.sku, field: 'Master Data', issue: `SKU not in master — dimensions missing`, sev: 'High' });
  });

  // ── Order Summary by type ────────────────────────────────────────────────────
  const typeMap = {};
  orders.forEach(o => {
    if (!typeMap[o.orderType]) typeMap[o.orderType] = { lines:0, orders:new Set(), skus:new Set(), dates:new Set(), qty:0 };
    const g = typeMap[o.orderType];
    g.lines++; g.qty += o.qty;
    if (o.orderNo) g.orders.add(o.orderNo);
    if (o.sku)  g.skus.add(o.sku);
    if (o.date) g.dates.add(o.date);
  });
  const orderSummary = Object.entries(typeMap).map(([type, g]) => ({
    orderType: type, lines: g.lines, uniqueOrders: g.orders.size,
    distinctSKUs: g.skus.size, distinctDates: g.dates.size, totalQty: g.qty,
    avgQtyPerLine: +(g.qty / g.lines).toFixed(1),
    avgLinesPerOrder: +(g.lines / Math.max(1, g.orders.size)).toFixed(1),
  }));

  // ── SKU Summary ──────────────────────────────────────────────────────────────
  const skuMap = {};
  orders.forEach(o => {
    if (!o.sku) return;
    if (!skuMap[o.sku]) skuMap[o.sku] = { lines:0, orders:new Set(), qty:0, dates:[] };
    skuMap[o.sku].lines++;
    skuMap[o.sku].qty += o.qty;
    if (o.orderNo) skuMap[o.sku].orders.add(o.orderNo);
    if (o.date)    skuMap[o.sku].dates.push(o.date);
  });
  const skuSummary = Object.entries(skuMap).map(([sku, g]) => {
    const m = masterMap.get(sku) || { L:0, W:0, H:0, weight:0, volume:0 };
    const ds = [...g.dates].sort();
    return { sku, lines: g.lines, uniqueOrders: g.orders.size, totalQty: g.qty,
      firstDate: ds[0]||'—', lastDate: ds[ds.length-1]||'—',
      L:m.L, W:m.W, H:m.H, weight:m.weight,
      volumePerUnit: m.volume,
      totalVolume: m.volume * g.qty / 1e9, // mm³ → m³
    };
  });

  // ── ABC by volume ─────────────────────────────────────────────────────────
  const abcData = [...skuSummary].sort((a,b) => b.totalVolume - a.totalVolume);
  const totVol = abcData.reduce((s,r) => s+r.totalVolume, 0);
  let cum = 0;
  abcData.forEach(r => {
    cum += r.totalVolume;
    r.cumVolPct = totVol > 0 ? +(cum/totVol*100).toFixed(2) : 0;
    r.abc = r.cumVolPct <= 70 ? 'A' : r.cumVolPct <= 90 ? 'B' : 'C';
  });

  // ── FMS by lines ──────────────────────────────────────────────────────────
  const fmsData = [...skuSummary].sort((a,b) => b.lines - a.lines);
  const totLines = fmsData.reduce((s,r) => s+r.lines, 0);
  let cumL = 0;
  fmsData.forEach(r => {
    cumL += r.lines;
    r.cumLinesPct = totLines > 0 ? +(cumL/totLines*100).toFixed(2) : 0;
    r.fms = r.cumLinesPct <= 33 ? 'Fast' : r.cumLinesPct <= 67 ? 'Medium' : 'Slow';
  });

  // ── ABC-FMS Matrix ────────────────────────────────────────────────────────
  const abcBysku = Object.fromEntries(abcData.map(r => [r.sku, r.abc]));
  const fmsBysku = Object.fromEntries(fmsData.map(r => [r.sku, r.fms]));
  const matrixData = skuSummary.map(r => ({
    ...r, abc: abcBysku[r.sku]||'C', fms: fmsBysku[r.sku]||'Slow',
  }));
  const matrix = {};
  ['A','B','C'].forEach(a => ['Fast','Medium','Slow'].forEach(f => {
    const key = `${a}-${f}`;
    const items = matrixData.filter(r => r.abc===a && r.fms===f);
    matrix[key] = { count:items.length, totalQty:items.reduce((s,r)=>s+r.totalQty,0),
      totalVol:items.reduce((s,r)=>s+r.totalVolume,0), skus:items.map(r=>r.sku) };
  }));

  return { anomalies, orderSummary, skuSummary, abcData, fmsData, matrixData, matrix, totVol };
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function exportReport(mAnom, analysis, narrative) {
  const wb = XLSX.utils.book_new();
  const { anomalies, orderSummary, skuSummary, abcData, fmsData, matrix } = analysis;
  const today = new Date().toLocaleDateString();

  const ws = (data, cols) => {
    const s = XLSX.utils.aoa_to_sheet(data);
    if (cols) s['!cols'] = cols.map(w => ({wch:w}));
    return s;
  };

  // ── 1. Anomaly Report ─────────────────────────────────────────────────────
  const allAnom = [...mAnom.map(a=>({...a,src:'Master'})), ...anomalies.map(a=>({...a,src:'Order'}))];
  XLSX.utils.book_append_sheet(wb, ws([
    ['PACKWISE — DATA ANOMALY REPORT'], ['Generated:', today], ['Total issues:', allAnom.length], [],
    ['Source','Row','SKU','Field','Issue','Severity'],
    ...(allAnom.length ? allAnom.map(a=>[a.src, a.row||a.rn, a.sku, a.field, a.issue, a.sev])
      : [['—','—','—','—','No issues found','—']]),
    [], ['SUMMARY'],
    ['High severity:', allAnom.filter(a=>a.sev==='High').length],
    ['Medium severity:', allAnom.filter(a=>a.sev==='Medium').length],
    ['Master issues:', mAnom.length], ['Order issues:', anomalies.length],
  ], [10,6,22,14,55,10]), '1. Anomalies');

  // ── 2. Order Summary ──────────────────────────────────────────────────────
  XLSX.utils.book_append_sheet(wb, ws([
    ['ORDER SUMMARY BY TYPE'], ['Generated:', today], [],
    ['Order Type','Lines','Unique Orders','Distinct SKUs','Distinct Dates','Total Qty','Avg Qty/Line','Avg Lines/Order'],
    ...orderSummary.map(r=>[r.orderType,r.lines,r.uniqueOrders,r.distinctSKUs,r.distinctDates,r.totalQty,r.avgQtyPerLine,r.avgLinesPerOrder]),
    [], ['TOTAL',
      orderSummary.reduce((s,r)=>s+r.lines,0),
      orderSummary.reduce((s,r)=>s+r.uniqueOrders,0),'—','—',
      orderSummary.reduce((s,r)=>s+r.totalQty,0),'—','—'],
  ], [16,8,14,14,14,12,14,16]), '2. Order Summary');

  // ── 3. SKU Summary ────────────────────────────────────────────────────────
  XLSX.utils.book_append_sheet(wb, ws([
    ['SKU SUMMARY'], [],
    ['SKU','Lines','Unique Orders','Total Qty','L mm','W mm','H mm','Wt/Box kg','Vol/Unit m³','Total Vol m³','First Date','Last Date'],
    ...[...skuSummary].sort((a,b)=>b.totalQty-a.totalQty).map(r=>[
      r.sku, r.lines, r.uniqueOrders, r.totalQty,
      r.L||'—', r.W||'—', r.H||'—', r.weight||'—',
      r.volumePerUnit>0 ? +(r.volumePerUnit/1e9).toFixed(6) : '—',
      r.totalVolume>0 ? +r.totalVolume.toFixed(4) : '—',
      r.firstDate, r.lastDate]),
  ], [22,8,14,12,8,8,8,12,14,14,12,12]), '3. SKU Summary');

  // ── 4. ABC Analysis ───────────────────────────────────────────────────────
  const abcSummary = ['A','B','C'].map(cls => {
    const items = abcData.filter(r=>r.abc===cls);
    const vol = items.reduce((s,r)=>s+r.totalVolume,0);
    const totV = abcData.reduce((s,r)=>s+r.totalVolume,0);
    return [cls, items.length, (items.length/abcData.length*100).toFixed(1)+'%',
      vol.toFixed(3), totV>0?(vol/totV*100).toFixed(1)+'%':'0%'];
  });
  XLSX.utils.book_append_sheet(wb, ws([
    ['ABC ANALYSIS — BY SHIPPING VOLUME'],
    ['A = Top 70% of total volume | B = Next 20% | C = Bottom 10%'], [],
    ['Rank','SKU','Total Qty','Vol/Unit (m³)','Total Volume (m³)','Cumulative Vol %','ABC Class'],
    ...abcData.map((r,i)=>[i+1, r.sku, r.totalQty,
      r.volumePerUnit>0?+(r.volumePerUnit/1e9).toFixed(6):'—',
      r.totalVolume>0?+r.totalVolume.toFixed(4):'—',
      r.cumVolPct+'%', r.abc]),
    [], ['SUMMARY'], ['Class','SKU Count','% SKUs','Total Volume (m³)','% Volume'],
    ...abcSummary,
  ], [6,22,12,14,18,18,10]), '4. ABC Analysis');

  // ── 5. FMS Analysis ───────────────────────────────────────────────────────
  const fmsSummary = ['Fast','Medium','Slow'].map(cls => {
    const items = fmsData.filter(r=>r.fms===cls);
    const lines = items.reduce((s,r)=>s+r.lines,0);
    const totL = fmsData.reduce((s,r)=>s+r.lines,0);
    return [cls, items.length, (items.length/fmsData.length*100).toFixed(1)+'%',
      lines, totL>0?(lines/totL*100).toFixed(1)+'%':'0%'];
  });
  XLSX.utils.book_append_sheet(wb, ws([
    ['FMS ANALYSIS — BY ORDER LINES (FREQUENCY)'],
    ['Fast = Top 33% of lines | Medium = Mid 33% | Slow = Bottom 33%'], [],
    ['Rank','SKU','Lines','Total Qty','Unique Orders','Cumulative Lines %','FMS Class'],
    ...fmsData.map((r,i)=>[i+1, r.sku, r.lines, r.totalQty, r.uniqueOrders, r.cumLinesPct+'%', r.fms]),
    [], ['SUMMARY'], ['Class','SKU Count','% SKUs','Total Lines','% Lines'],
    ...fmsSummary,
  ], [6,22,8,12,14,20,10]), '5. FMS Analysis');

  // ── 6. ABC-FMS Matrix ─────────────────────────────────────────────────────
  const matRows = [
    ['ABC-FMS MATRIX'], ['Volume (ABC) vs Frequency (FMS)'], [],
    ['','Fast (F)','Medium (M)','Slow (S)'],
    ...['A','B','C'].map(a => [a,
      ...['Fast','Medium','Slow'].map(f => {
        const c = matrix[`${a}-${f}`];
        return c?.count ? `${c.count} SKUs | ${c.totalQty.toLocaleString()} units` : '—';
      })
    ]),
    [], ['INTERPRETATION'],
    ['A-Fast',  'Star products — highest priority for packing & stock planning'],
    ['A-Medium','High value, moderate frequency — plan containers regularly'],
    ['A-Slow',  'High value, low frequency — consider batch shipping to consolidate'],
    ['B-Fast',  'Frequent movers — monitor closely'],
    ['C-Fast',  'Many small orders — consider minimum order quantity review'],
    ['C-Slow',  'Long tail — candidates for rationalisation or consolidation'],
    [], ['DETAILED SKU LIST BY CELL'], [],
    ...['A','B','C'].flatMap(a => ['Fast','Medium','Slow'].flatMap(f => {
      const c = matrix[`${a}-${f}`];
      if (!c?.count) return [];
      return [[`${a}-${f} (${c.count} SKUs):`],[c.skus.slice(0,50).join(', ')+(c.count>50?'...' : '')],[]];
    })),
  ];
  if (narrative) { matRows.push([], ['AI EXECUTIVE SUMMARY'], [narrative]); }
  XLSX.utils.book_append_sheet(wb, ws(matRows, [16,35,35,35]), '6. ABC-FMS Matrix');

  XLSX.writeFile(wb, `PackWise_Order_Analysis_${today.replace(/\//g,'-')}.xlsx`);
}

// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────
export default function OrderAnalyserTool() {
  // Master SKU state
  const [mFileName, setMFileName] = useState('');
  const [mRaw, setMRaw] = useState(null);
  const [mColMap, setMColMap] = useState(null);
  const [mAnom, setMAnom]     = useState([]);
  const [masterMap, setMasterMap] = useState(new Map());
  const [mDone, setMDone]     = useState(false);
  const [mLoading, setMLoading] = useState(false);
  const [mError, setMError]   = useState('');

  // Order state
  const [oFileName, setOFileName] = useState('');
  const [oRaw, setORaw]     = useState(null);
  const [oColMap, setOColMap] = useState(null);
  const [analysis, setAnalysis] = useState(null);
  const [narrative, setNarrative] = useState('');
  const [oLoading, setOLoading]   = useState(false);
  const [narLoading, setNarLoading] = useState(false);
  const [oError, setOError]   = useState('');

  const mRef = useRef(null);
  const oRef = useRef(null);
  const [mDrag, setMDrag] = useState(false);
  const [oDrag, setODrag] = useState(false);

  // ── STEP 1: Process master SKU file ────────────────────────────────────────
  const processMaster = async (file) => {
    setMFileName(file.name); setMError(''); setMDone(false);
    setMColMap(null); setMAnom([]); setMasterMap(new Map());
    setMLoading(true);
    try {
      const { rows, sheets } = await parseExcel(file);
      setMRaw(rows);
      const preview = rows.slice(0,30).map((r,i)=>`Row${i+1}: ${r.join('\t')}`).join('\n');

      // AI: map columns
      const mapped = await claudeCall(
`You are mapping columns in a Master SKU / Product data file.

File: ${file.name} | Sheets: ${sheets.join(', ')} | Total rows: ${rows.length}

PREVIEW (first 30 rows, tab-separated):
${preview}

Map to these standard fields. Return the EXACT header text from the file for each field.
If a field is not present, use null.

Return ONLY valid JSON:
{
  "sku_name":   "exact header text for SKU / material / product code",
  "length":     "exact header text for box length",
  "width":      "exact header text for box width",
  "height":     "exact header text for box height",
  "weight":     "exact header text for weight, or null",
  "category":   "exact header text for category/type, or null",
  "description":"exact header text for description, or null",
  "unit":       "mm or cm or m — your best guess based on values",
  "headerRow":  0
}

Rules: Return the actual column header string, not a column letter. Match case exactly.`
      );

      const hRow = rows[mapped.headerRow] || rows[0];
      const resolve = buildColIdx(hRow, mapped);
      const ci = {
        sku_name:    resolve('sku_name'),
        length:      resolve('length'),
        width:       resolve('width'),
        height:      resolve('height'),
        weight:      resolve('weight'),
        category:    resolve('category'),
        description: resolve('description'),
      };

      if (ci.sku_name === null) throw new Error('Could not find SKU name column. Add a hint in your file header.');

      const dataRows = rows.slice(mapped.headerRow + 1).filter(r => r.some(c => String(c).trim()));
      const anomalies = detectMasterAnomalies(dataRows, ci);
      const mm = buildMasterMap(dataRows, ci, mapped.unit);

      setMColMap({ ...ci, unit: mapped.unit, headerRow: mapped.headerRow });
      setMAnom(anomalies);
      setMasterMap(mm);
      setMDone(true);
    } catch (err) { setMError(err.message); }
    setMLoading(false);
  };

  // ── STEP 2: Process order file ──────────────────────────────────────────────
  const processOrder = async (file) => {
    setOFileName(file.name); setOError(''); setAnalysis(null); setNarrative('');
    setOColMap(null); setOLoading(true);
    try {
      const { rows, sheets } = await parseExcel(file);
      setORaw(rows);
      const preview = rows.slice(0,30).map((r,i)=>`Row${i+1}: ${r.join('\t')}`).join('\n');

      // AI: map order columns
      const mapped = await claudeCall(
`You are mapping columns in an Order / Sales data file for a logistics tool.

File: ${file.name} | Sheets: ${sheets.join(', ')} | Total rows: ${rows.length}

PREVIEW (first 30 rows, tab-separated):
${preview}

Map to these standard fields. Return the EXACT header text from the file.
If a field is not present, use null.

Return ONLY valid JSON:
{
  "order_no":   "exact header for order / invoice / document number",
  "order_type": "exact header for order type / document type / STO vs customer, or null",
  "sku":        "exact header for SKU / material / article / product code",
  "qty":        "exact header for quantity / ordered qty / delivered qty",
  "date":       "exact header for date (order / invoice / dispatch / GI date)",
  "customer":   "exact header for customer / sold-to / ship-to name, or null",
  "line_no":    "exact header for line item number, or null",
  "headerRow":  0
}

Return the actual column header string. Match case exactly.`
      );

      const hRow = rows[mapped.headerRow] || rows[0];
      const resolve = buildColIdx(hRow, mapped);
      const ci = {
        order_no:   resolve('order_no'),
        order_type: resolve('order_type'),
        sku:        resolve('sku'),
        qty:        resolve('qty'),
        date:       resolve('date'),
        customer:   resolve('customer'),
        line_no:    resolve('line_no'),
      };

      if (ci.sku === null) throw new Error('Could not find SKU column. Add a clear header like "SKU", "Material", or "Product Code".');
      if (ci.qty === null) throw new Error('Could not find Quantity column.');

      const dataRows = rows.slice(mapped.headerRow + 1).filter(r => r.some(c => String(c).trim()));
      const result = runOrderAnalytics(dataRows, ci, masterMap);

      setOColMap({ ...ci, headerRow: mapped.headerRow });
      setAnalysis(result);
    } catch (err) { setOError(err.message); }
    setOLoading(false);
  };

  // ── STEP 3: AI Narrative ────────────────────────────────────────────────────
  const getNarrative = async () => {
    if (!analysis) return;
    setNarLoading(true);
    try {
      const { orderSummary, skuSummary, abcData, fmsData } = analysis;
      const totalQty = orderSummary.reduce((s,r)=>s+r.totalQty,0);
      const totalLines = orderSummary.reduce((s,r)=>s+r.lines,0);
      const aSkus = abcData.filter(r=>r.abc==='A');
      const fSkus = fmsData.filter(r=>r.fms==='Fast');
      const aVolPct = aSkus.length>0 ? (aSkus.length/skuSummary.length*100).toFixed(0) : 0;

      const prompt = `Write a 3-paragraph executive summary for a logistics manager based on this order analysis:

- Total order lines: ${totalLines.toLocaleString()}
- Total qty shipped: ${totalQty.toLocaleString()}
- Order types: ${orderSummary.map(r=>`${r.orderType} (${r.lines} lines)`).join(', ')}
- Total distinct SKUs: ${skuSummary.length}
- ABC-A SKUs: ${aSkus.length} (${aVolPct}% of SKUs = top 70% of shipping volume)
- Fast-moving SKUs: ${fSkus.length}
- Anomalies found: ${[...mAnom,...analysis.anomalies].length}
- Top 5 SKUs by volume: ${abcData.slice(0,5).map(r=>r.sku).join(', ')}

Write in plain business English. Focus on: what the data shows, key packing/shipping priorities, and 2-3 specific recommendations. Keep it under 200 words.`;

      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ model:'claude-sonnet-4-6', max_tokens:500,
          messages:[{role:'user', content:prompt}] }),
      });
      const d = await r.json();
      setNarrative(d.content?.[0]?.text || '');
    } catch { setNarrative(''); }
    setNarLoading(false);
  };

  // ── STYLED HELPERS ─────────────────────────────────────────────────────────
  const stepNum = (n, done) => (
    <div style={{width:'32px',height:'32px',borderRadius:'50%',flexShrink:0,
      background:done?'#166534':n==='active'?'#be185d':'#e2e8f0',
      display:'flex',alignItems:'center',justifyContent:'center',
      color:done||n==='active'?'#fff':'#9ca3af',fontWeight:'800',fontSize:'14px'}}>
      {done?'✓':n}
    </div>
  );

  const upload = (drag, setDrag, ref, onFile, label, fileName, done) => (
    <div onDragOver={e=>{e.preventDefault();setDrag(true);}}
      onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);onFile(e.dataTransfer.files[0]);}}
      onClick={()=>ref.current?.click()}
      style={{border:`2px dashed ${drag?'#be185d':done?'#bbf7d0':'#d1d9e0'}`,
        borderRadius:'10px',padding:'24px',textAlign:'center',cursor:'pointer',
        background:drag?'#fdf2f8':done?'#f0fdf4':'#fafbfc',transition:'all 0.2s'}}>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{display:'none'}}
        onChange={e=>{if(e.target.files[0]) onFile(e.target.files[0]); e.target.value='';}}/>
      <div style={{fontSize:'28px',marginBottom:'6px'}}>{done?'✅':'📊'}</div>
      <div style={{fontWeight:'700',color:'#374151',fontSize:'13px'}}>
        {fileName || label}
      </div>
      {done && <div style={{fontSize:'12px',color:'#166534',fontWeight:'600',marginTop:'4px'}}>✓ Loaded — click to replace</div>}
    </div>
  );

  const sevBadge = sev => (
    <span style={{background:sev==='High'?'#fee2e2':'#fef9c3',
      color:sev==='High'?'#991b1b':'#854d0e',
      padding:'2px 7px',borderRadius:'99px',fontSize:'11px',fontWeight:'700'}}>{sev}</span>
  );

  const abcBadge = cls => {
    const colors = {A:['#dcfce7','#166534'],B:['#fef9c3','#854d0e'],C:['#f3f4f6','#6b7280']};
    const [bg,col] = colors[cls]||colors.C;
    return <span style={{background:bg,color:col,padding:'2px 8px',borderRadius:'99px',fontSize:'12px',fontWeight:'700'}}>{cls}</span>;
  };

  const fmsBadge = cls => {
    const colors = {Fast:['#dcfce7','#166534'],Medium:['#fef9c3','#854d0e'],Slow:['#fee2e2','#991b1b']};
    const [bg,col] = colors[cls]||colors.Slow;
    return <span style={{background:bg,color:col,padding:'2px 8px',borderRadius:'99px',fontSize:'12px',fontWeight:'700'}}>{cls}</span>;
  };

  return (
    <div>
      <div style={S.sectionDesc}>
        Upload your Master SKU data and Order data in any format. AI automatically maps columns,
        flags anomalies, and produces a 6-sheet Excel report with ABC analysis, FMS classification,
        and an ABC-FMS matrix — ready to share with your logistics or supply chain team.
      </div>

      {/* ── STEP 1: MASTER SKU ─────────────────────────────────────────────── */}
      <div style={S.card}>
        <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'16px'}}>
          {stepNum(1, mDone)}
          <div>
            <div style={S.cardTitle}>Master SKU Data</div>
            <div style={{fontSize:'12px',color:'#6b7280'}}>
              Any format — AI detects SKU name, length, width, height, weight columns automatically
            </div>
          </div>
        </div>

        {upload(mDrag, setMDrag, mRef, processMaster,
          'Drop Master SKU Excel / CSV here or click to browse', mFileName, mDone)}

        {mLoading && (
          <div style={{marginTop:'14px',padding:'14px',background:'#f5f3ff',borderRadius:'8px',
            fontSize:'13px',color:'#6d28d9',textAlign:'center'}}>
            ⏳ AI is reading your file and mapping columns... (~5 seconds)
          </div>
        )}

        {mError && <div style={{...S.error,marginTop:'12px'}}>⚠ {mError}</div>}

        {mDone && mColMap && (
          <div style={{marginTop:'14px'}}>
            {/* Detected columns */}
            <div style={{background:'#f0fdf4',border:'1px solid #bbf7d0',borderRadius:'8px',
              padding:'12px',marginBottom:'12px'}}>
              <div style={{fontWeight:'700',color:'#166534',fontSize:'13px',marginBottom:'8px'}}>
                ✓ Columns detected (unit: {mColMap.unit || 'mm'})
              </div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:'6px'}}>
                {[['SKU Name',mColMap.sku_name],['Length',mColMap.length],['Width',mColMap.width],
                  ['Height',mColMap.height],['Weight',mColMap.weight],['Category',mColMap.category]
                ].map(([label, val]) => (
                  <div key={label} style={{fontSize:'12px',display:'flex',gap:'5px'}}>
                    <span style={{color:'#9ca3af',minWidth:'58px'}}>{label}:</span>
                    <span style={{fontWeight:'600',color:val!==null?'#1d4ed8':'#d1d5db'}}>
                      {val !== null ? `Col ${val}` : '—'}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Anomaly summary */}
            <div style={{display:'flex',gap:'12px',marginBottom:'12px'}}>
              {[['Total SKUs', masterMap.size,'#eff6ff','#1d4ed8'],
                ['Issues Found', mAnom.length, mAnom.length?'#fff1f2':'#f0fdf4', mAnom.length?'#be185d':'#166534'],
                ['High Severity', mAnom.filter(a=>a.sev==='High').length,'#fff1f2','#991b1b'],
                ['Medium', mAnom.filter(a=>a.sev==='Medium').length,'#fef9c3','#854d0e'],
              ].map(([l,v,bg,col]) => (
                <div key={l} style={{flex:1,background:bg,borderRadius:'8px',padding:'10px',textAlign:'center'}}>
                  <div style={{fontSize:'20px',fontWeight:'800',color:col}}>{v}</div>
                  <div style={{fontSize:'10px',color:'#6b7280',marginTop:'2px',fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                </div>
              ))}
            </div>

            {/* Anomaly table */}
            {mAnom.length > 0 && (
              <div style={{border:'1px solid #e2e8f0',borderRadius:'8px',overflow:'hidden',maxHeight:'200px',overflowY:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                  <thead><tr>
                    {['Row','SKU','Field','Issue','Severity'].map(h=>(
                      <th key={h} style={{padding:'8px 10px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',
                        textAlign:'left',fontWeight:'700',fontSize:'11px',color:'#6b7280',
                        textTransform:'uppercase',letterSpacing:'0.04em',whiteSpace:'nowrap'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {mAnom.slice(0,50).map((a,i)=>(
                      <tr key={i} style={{background:i%2?'#fafbfc':'#fff'}}>
                        <td style={{padding:'7px 10px',color:'#6b7280'}}>{a.row}</td>
                        <td style={{padding:'7px 10px',fontWeight:'600'}}>{a.sku}</td>
                        <td style={{padding:'7px 10px',color:'#374151'}}>{a.field}</td>
                        <td style={{padding:'7px 10px',color:'#374151'}}>{a.issue}</td>
                        <td style={{padding:'7px 10px'}}>{sevBadge(a.sev)}</td>
                      </tr>))}
                  </tbody>
                </table>
                {mAnom.length>50&&<div style={{padding:'8px 12px',fontSize:'11px',color:'#9ca3af'}}>Showing 50 of {mAnom.length} — full list in Excel report</div>}
              </div>
            )}
            {mAnom.length===0 && <div style={{background:'#f0fdf4',border:'1px solid #bbf7d0',borderRadius:'8px',padding:'10px 14px',fontSize:'13px',color:'#166534'}}>✓ No anomalies found in master data</div>}
          </div>
        )}
      </div>

      {/* ── STEP 2: ORDER DATA ─────────────────────────────────────────────── */}
      <div style={{...S.card, opacity: mDone?1:0.5, pointerEvents: mDone?'auto':'none'}}>
        <div style={{display:'flex',alignItems:'center',gap:'12px',marginBottom:'16px'}}>
          {stepNum(2, !!analysis)}
          <div>
            <div style={S.cardTitle}>Order Data</div>
            <div style={{fontSize:'12px',color:'#6b7280'}}>
              {mDone ? 'Upload order/sales data — AI detects order no, SKU, qty, date, order type columns'
                : 'Complete Step 1 first'}
            </div>
          </div>
        </div>

        {mDone && (
          <>
            {upload(oDrag, setODrag, oRef, processOrder,
              'Drop Order Excel / CSV here or click to browse', oFileName, !!analysis)}

            {oLoading && (
              <div style={{marginTop:'14px',padding:'14px',background:'#f5f3ff',borderRadius:'8px',
                fontSize:'13px',color:'#6d28d9',textAlign:'center'}}>
                ⏳ AI is mapping order columns and running analytics... (~8 seconds)
              </div>
            )}
            {oError && <div style={{...S.error,marginTop:'12px'}}>⚠ {oError}</div>}
          </>
        )}

        {/* ── RESULTS ──────────────────────────────────────────────────────── */}
        {analysis && (
          <div style={{marginTop:'16px'}}>
            {/* Order anomaly count */}
            {analysis.anomalies.length > 0 && (
              <div style={{background:'#fff1f2',border:'1px solid #fecaca',borderRadius:'8px',
                padding:'10px 14px',marginBottom:'14px',fontSize:'13px',
                display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span style={{color:'#991b1b',fontWeight:'600'}}>
                  ⚠ {analysis.anomalies.length} order anomalies found
                  ({analysis.anomalies.filter(a=>a.sev==='High').length} high severity)
                </span>
                <span style={{fontSize:'12px',color:'#9ca3af'}}>Full list in Excel report</span>
              </div>
            )}

            {/* Summary cards */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'12px',marginBottom:'16px'}}>
              {[
                ['Order Lines', analysis.orderSummary.reduce((s,r)=>s+r.lines,0).toLocaleString(),'#eff6ff','#1d4ed8'],
                ['Unique Orders', analysis.orderSummary.reduce((s,r)=>s+r.uniqueOrders,0).toLocaleString(),'#f5f3ff','#6d28d9'],
                ['Distinct SKUs', analysis.skuSummary.length.toLocaleString(),'#f0fdf4','#166534'],
                ['Total Qty', analysis.orderSummary.reduce((s,r)=>s+r.totalQty,0).toLocaleString(),'#fff7ed','#c2410c'],
              ].map(([l,v,bg,col])=>(
                <div key={l} style={{background:bg,borderRadius:'8px',padding:'12px',textAlign:'center'}}>
                  <div style={{fontSize:'20px',fontWeight:'800',color:col}}>{v}</div>
                  <div style={{fontSize:'10px',color:'#6b7280',marginTop:'3px',fontWeight:'600',textTransform:'uppercase'}}>{l}</div>
                </div>
              ))}
            </div>

            {/* Order summary by type */}
            <div style={{marginBottom:'14px',border:'1px solid #e2e8f0',borderRadius:'8px',overflow:'hidden'}}>
              <div style={{padding:'8px 14px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',
                fontWeight:'700',fontSize:'12px',color:'#374151'}}>Order Summary by Type</div>
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                  <thead><tr>
                    {['Order Type','Lines','Unique Orders','Distinct SKUs','Distinct Dates','Total Qty','Avg Qty/Line','Avg Lines/Order'].map(h=>(
                      <th key={h} style={{padding:'7px 10px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',
                        textAlign:'left',fontWeight:'600',fontSize:'11px',color:'#6b7280',
                        textTransform:'uppercase',whiteSpace:'nowrap'}}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {analysis.orderSummary.map((r,i)=>(
                      <tr key={i} style={{background:i%2?'#fafbfc':'#fff'}}>
                        <td style={{padding:'7px 10px',fontWeight:'700'}}>{r.orderType}</td>
                        <td style={{padding:'7px 10px',textAlign:'right'}}>{r.lines.toLocaleString()}</td>
                        <td style={{padding:'7px 10px',textAlign:'right'}}>{r.uniqueOrders.toLocaleString()}</td>
                        <td style={{padding:'7px 10px',textAlign:'right'}}>{r.distinctSKUs.toLocaleString()}</td>
                        <td style={{padding:'7px 10px',textAlign:'right'}}>{r.distinctDates.toLocaleString()}</td>
                        <td style={{padding:'7px 10px',textAlign:'right',fontWeight:'600'}}>{r.totalQty.toLocaleString()}</td>
                        <td style={{padding:'7px 10px',textAlign:'right'}}>{r.avgQtyPerLine}</td>
                        <td style={{padding:'7px 10px',textAlign:'right'}}>{r.avgLinesPerOrder}</td>
                      </tr>))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* ABC preview */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px',marginBottom:'14px'}}>
              {/* ABC summary */}
              <div style={{border:'1px solid #e2e8f0',borderRadius:'8px',overflow:'hidden'}}>
                <div style={{padding:'8px 14px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',fontWeight:'700',fontSize:'12px'}}>ABC Summary (by shipping volume)</div>
                {['A','B','C'].map(cls=>{
                  const items = analysis.abcData.filter(r=>r.abc===cls);
                  const vol = items.reduce((s,r)=>s+r.totalVolume,0);
                  const tot = analysis.totVol;
                  return(
                    <div key={cls} style={{padding:'8px 14px',borderBottom:'1px solid #f1f5f9',
                      display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                      <div style={{display:'flex',gap:'8px',alignItems:'center'}}>
                        {abcBadge(cls)}
                        <span style={{fontSize:'12px',color:'#374151'}}>{items.length} SKUs</span>
                      </div>
                      <span style={{fontSize:'12px',color:'#6b7280'}}>
                        {tot>0?(vol/tot*100).toFixed(1):'0'}% of volume
                      </span>
                    </div>
                  );
                })}
              </div>
              {/* FMS summary */}
              <div style={{border:'1px solid #e2e8f0',borderRadius:'8px',overflow:'hidden'}}>
                <div style={{padding:'8px 14px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',fontWeight:'700',fontSize:'12px'}}>FMS Summary (by order lines)</div>
                {['Fast','Medium','Slow'].map(cls=>{
                  const items = analysis.fmsData.filter(r=>r.fms===cls);
                  const lines = items.reduce((s,r)=>s+r.lines,0);
                  const tot = analysis.fmsData.reduce((s,r)=>s+r.lines,0);
                  return(
                    <div key={cls} style={{padding:'8px 14px',borderBottom:'1px solid #f1f5f9',
                      display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                      <div style={{display:'flex',gap:'8px',alignItems:'center'}}>
                        {fmsBadge(cls)}
                        <span style={{fontSize:'12px',color:'#374151'}}>{items.length} SKUs</span>
                      </div>
                      <span style={{fontSize:'12px',color:'#6b7280'}}>
                        {tot>0?(lines/tot*100).toFixed(1):'0'}% of lines
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* ABC-FMS Matrix preview */}
            <div style={{border:'1px solid #e2e8f0',borderRadius:'8px',overflow:'hidden',marginBottom:'14px'}}>
              <div style={{padding:'8px 14px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',fontWeight:'700',fontSize:'12px'}}>ABC-FMS Matrix (SKU count)</div>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:'12px'}}>
                <thead><tr>
                  <th style={{padding:'8px 12px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',textAlign:'left',fontSize:'11px',color:'#6b7280'}}></th>
                  {['Fast','Medium','Slow'].map(f=>(
                    <th key={f} style={{padding:'8px 12px',background:'#f8fafc',borderBottom:'1px solid #e2e8f0',textAlign:'center',fontSize:'11px',color:'#6b7280',fontWeight:'700'}}>{f}</th>))}
                </tr></thead>
                <tbody>
                  {['A','B','C'].map((a,ai)=>(
                    <tr key={a}>
                      <td style={{padding:'10px 12px',fontWeight:'800',color:'#374151'}}>{abcBadge(a)}</td>
                      {['Fast','Medium','Slow'].map(f=>{
                        const c = analysis.matrix[`${a}-${f}`];
                        const n = c?.count||0;
                        return(
                          <td key={f} style={{padding:'10px 12px',textAlign:'center',
                            background: n>0?(a==='A'&&f==='Fast'?'#dcfce7':a==='C'&&f==='Slow'?'#fee2e2':'#fff'):'#f8fafc'}}>
                            {n>0 ? (
                              <><div style={{fontWeight:'700',fontSize:'14px',color:'#111'}}>{n}</div>
                              <div style={{fontSize:'10px',color:'#9ca3af'}}>{c?.totalQty?.toLocaleString()||0} units</div></>
                            ) : <span style={{color:'#d1d5db'}}>—</span>}
                          </td>);
                      })}
                    </tr>))}
                </tbody>
              </table>
            </div>

            {/* Narrative */}
            {!narrative && (
              <button onClick={getNarrative} disabled={narLoading}
                style={{...S.btnSecondary,width:'100%',marginBottom:'12px',
                  background:narLoading?'#f1f5f9':'#f5f3ff',
                  color:narLoading?'#9ca3af':'#6d28d9',border:'1px solid #ddd6fe'}}>
                {narLoading?'⏳ Generating executive summary...':'🤖 Generate AI Executive Summary'}
              </button>
            )}
            {narrative && (
              <div style={{background:'#f5f3ff',border:'1px solid #ddd6fe',borderRadius:'8px',
                padding:'14px',marginBottom:'12px',fontSize:'13px',color:'#374151',lineHeight:'1.7'}}>
                <div style={{fontWeight:'700',color:'#6d28d9',marginBottom:'6px'}}>🤖 Executive Summary</div>
                {narrative}
              </div>
            )}

            {/* Download button */}
            <button onClick={()=>exportReport(mAnom, analysis, narrative)}
              style={{...S.btnPrimary,width:'100%',fontSize:'15px',
                background:'linear-gradient(135deg,#be185d,#9d174d)'}}>
              ⬇ Download 6-Sheet Excel Report
            </button>
            <div style={{fontSize:'11px',color:'#9ca3af',textAlign:'center',marginTop:'6px'}}>
              Sheets: Anomalies · Order Summary · SKU Summary · ABC Analysis · FMS Analysis · ABC-FMS Matrix
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
