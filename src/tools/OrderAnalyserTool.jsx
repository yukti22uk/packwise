// ─── ORDER ANALYSER TOOL ──────────────────────────────────────────────────────
// Step 1: Paste Master SKU data (fixed column order) → validate + flag anomalies
// Step 2: Paste Order data (fixed column order) → full analytics → 6-sheet Excel
// No AI column mapping — user pastes in expected order shown on screen.
import { useState } from 'react';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import { S } from '../components/styles.jsx';

// ─── PARSE PASTED TSV ────────────────────────────────────────────────────────
function parseTSV(text) {
  return text.trim().split('\n')
    .map(r => r.split('\t').map(c => c.trim()))
    .filter(r => r.some(c => c));
}

function isHeaderRow(row) {
  // True if the row looks like headers (text in numeric columns)
  if (!row[1] && !row[2]) return false;
  const num1 = parseFloat(row[1]);
  const num2 = parseFloat(row[2]);
  return isNaN(num1) && isNaN(num2);
}

// ─── MASTER SKU: PARSE & VALIDATE ────────────────────────────────────────────
// Fixed order: SKU Name | Length (mm) | Width (mm) | Height (mm) | Weight/Box (kg)
// Missing dimension/weight values are filled with averages from valid rows.
function parseMasterSKU(text) {
  const rows = parseTSV(text);
  const dataRows = rows.length > 0 && isHeaderRow(rows[0]) ? rows.slice(1) : rows;
  const skus = [], anomalies = [], seen = new Set();

  // ── Pass 1: parse raw values ─────────────────────────────────────────────
  const raw = dataRows.map((r, i) => ({
    rowNum: i + (rows.length > dataRows.length ? 2 : 1),
    name:   (r[0] || '').trim(),
    L:      parseFloat(r[1]) || 0,
    W:      parseFloat(r[2]) || 0,
    H:      parseFloat(r[3]) || 0,
    weight: parseFloat(r[4]) || 0,
  }));

  // ── Compute averages from valid (non-zero) rows ───────────────────────────
  const validL = raw.filter(r => r.L > 0).map(r => r.L);
  const validW = raw.filter(r => r.W > 0).map(r => r.W);
  const validH = raw.filter(r => r.H > 0).map(r => r.H);
  const validWt= raw.filter(r => r.weight > 0).map(r => r.weight);
  const avg = (arr) => arr.length > 0 ? Math.round(arr.reduce((s,v)=>s+v,0)/arr.length) : 0;
  const avgL  = avg(validL);
  const avgW  = avg(validW);
  const avgH  = avg(validH);
  const avgWt = +(validWt.reduce((s,v)=>s+v,0) / Math.max(1, validWt.length)).toFixed(2);

  // ── Pass 2: validate, fill missing with averages, flag anomalies ─────────
  raw.forEach(r => {
    if (!r.name) {
      anomalies.push({ row:r.rowNum, sku:'—', field:'SKU Name', issue:'Missing SKU name — row skipped', sev:'High' });
      return;
    }
    if (seen.has(r.name)) anomalies.push({ row:r.rowNum, sku:r.name, field:'SKU', issue:'Duplicate SKU', sev:'Medium' });
    else seen.add(r.name);

    // Fill missing dims with averages
    let { L, W, H, weight } = r;
    if (!L) {
      anomalies.push({ row:r.rowNum, sku:r.name, field:'Length', issue:`Missing — filled with average (${avgL} mm)`, sev:'Medium' });
      L = avgL;
    }
    if (!W) {
      anomalies.push({ row:r.rowNum, sku:r.name, field:'Width', issue:`Missing — filled with average (${avgW} mm)`, sev:'Medium' });
      W = avgW;
    }
    if (!H) {
      anomalies.push({ row:r.rowNum, sku:r.name, field:'Height', issue:`Missing — filled with average (${avgH} mm)`, sev:'Medium' });
      H = avgH;
    }
    if (!weight) {
      anomalies.push({ row:r.rowNum, sku:r.name, field:'Weight', issue:`Missing — filled with average (${avgWt} kg)`, sev:'Low' });
      weight = avgWt;
    }

    skus.push({ name:r.name, L, W, H, weight, volume: L * W * H });
  });

  const masterMap = new Map(skus.map(s => [s.name, s]));
  const avgValues = { L: avgL, W: avgW, H: avgH, weight: avgWt,
    filledL:  anomalies.filter(a=>a.field==='Length').length,
    filledW:  anomalies.filter(a=>a.field==='Width').length,
    filledH:  anomalies.filter(a=>a.field==='Height').length,
    filledWt: anomalies.filter(a=>a.field==='Weight').length,
  };
  return { skus, anomalies, masterMap, avgValues };
}

// ─── ORDER DATA: PARSE & ANALYSE ─────────────────────────────────────────────
// Fixed order: Order No | Order Type | SKU Code | Qty | Date | Category* | Dispatch Location*
// * = optional columns
function parseOrderData(text, masterMap) {
  const rows = parseTSV(text);
  const dataRows = rows.length > 0 && isHeaderRow(rows[0]) ? rows.slice(1) : rows;
  const orders = [], anomalies = [];

  dataRows.forEach((r, i) => {
    const rowNum      = i + (rows.length > dataRows.length ? 2 : 1);
    const orderNo     = r[0] || '';
    const dispatchLoc = r[1] ? r[1].trim() : '';   // col 2 — Dispatch Location
    const sku         = r[2] || '';
    const qty         = parseFloat(r[3]) || 0;
    const date        = r[4] || '';
    const orderType   = r[5] ? r[5].trim() : 'Unknown'; // optional col 6
    const category    = r[6] ? r[6].trim() : '';         // optional col 7

    if (!orderNo) anomalies.push({ row: rowNum, sku, field: 'Order No',  issue: 'Missing order/invoice number', sev: 'High' });
    if (!sku)     anomalies.push({ row: rowNum, sku: '—', field: 'SKU',  issue: 'Missing SKU code', sev: 'High' });
    if (qty <= 0) anomalies.push({ row: rowNum, sku, field: 'Qty',       issue: `Zero or negative qty: ${r[3]}`, sev: 'High' });
    if (!date)    anomalies.push({ row: rowNum, sku, field: 'Date',      issue: 'Missing date', sev: 'Medium' });
    if (sku && masterMap.size > 0 && !masterMap.has(sku))
      anomalies.push({ row: rowNum, sku, field: 'Master Data', issue: `SKU not in master — dimensions missing`, sev: 'High' });

    orders.push({ orderNo, orderType, sku, qty, date, category, dispatchLoc });
  });

  // Order summary by type
  const typeMap = {};
  orders.forEach(o => {
    const locKey = o.dispatchLoc || 'Unspecified';
    if (!typeMap[locKey]) typeMap[locKey] = { lines:0, orders:new Set(), skus:new Set(), dates:new Set(), qty:0 };
    const g = typeMap[locKey];
    g.lines++; g.qty += o.qty;
    if (o.orderNo) g.orders.add(o.orderNo);
    if (o.sku)     g.skus.add(o.sku);
    if (o.date)    g.dates.add(o.date);
  });
  const orderSummary = Object.entries(typeMap).map(([loc, g]) => ({
    dispatchLoc: loc, lines: g.lines, uniqueOrders: g.orders.size,
    distinctSKUs: g.skus.size, distinctDates: g.dates.size, totalQty: g.qty,
    avgQtyPerLine: +(g.qty / g.lines).toFixed(1),
    avgLinesPerOrder: +(g.lines / Math.max(1, g.orders.size)).toFixed(1),
  }));

  // SKU summary
  const skuMap = {};
  orders.forEach(o => {
    if (!o.sku) return;
    if (!skuMap[o.sku]) skuMap[o.sku] = { lines:0, orders:new Set(), qty:0, dates:[], categories:new Set(), locations:new Set() };
    skuMap[o.sku].lines++;
    skuMap[o.sku].qty += o.qty;
    if (o.orderNo)     skuMap[o.sku].orders.add(o.orderNo);
    if (o.date)        skuMap[o.sku].dates.push(o.date);
    if (o.category)    skuMap[o.sku].categories.add(o.category);
    if (o.dispatchLoc) skuMap[o.sku].locations.add(o.dispatchLoc);
  });
  const skuSummary = Object.entries(skuMap).map(([sku, g]) => {
    const m = masterMap.get(sku) || { L:0, W:0, H:0, weight:0, volume:0 };
    const ds = [...g.dates].sort();
    return { sku, lines:g.lines, uniqueOrders:g.orders.size, totalQty:g.qty,
      firstDate:ds[0]||'—', lastDate:ds[ds.length-1]||'—',
      categories:[...g.categories].join(', ')||'—',
      locations:[...g.locations].join(', ')||'—',
      L:m.L, W:m.W, H:m.H, weight:m.weight,
      volumePerUnit:m.volume, totalVolume:m.volume*g.qty/1e9 };
  });

  // ABC by volume
  const abcData = [...skuSummary].sort((a,b) => b.totalVolume - a.totalVolume);
  const totVol = abcData.reduce((s,r) => s+r.totalVolume, 0);
  let cum = 0;
  abcData.forEach(r => {
    cum += r.totalVolume;
    r.cumVolPct = totVol > 0 ? +(cum/totVol*100).toFixed(2) : 0;
    r.abc = r.cumVolPct <= 70 ? 'A' : r.cumVolPct <= 90 ? 'B' : 'C';
  });

  // FMS by lines
  const fmsData = [...skuSummary].sort((a,b) => b.lines - a.lines);
  const totLines = fmsData.reduce((s,r) => s+r.lines, 0);
  let cumL = 0;
  fmsData.forEach(r => {
    cumL += r.lines;
    r.cumLinesPct = totLines > 0 ? +(cumL/totLines*100).toFixed(2) : 0;
    r.fms = r.cumLinesPct <= 33 ? 'Fast' : r.cumLinesPct <= 67 ? 'Medium' : 'Slow';
  });

  // ABC-FMS matrix
  const abcBysku = Object.fromEntries(abcData.map(r => [r.sku, r.abc]));
  const fmsBysku = Object.fromEntries(fmsData.map(r => [r.sku, r.fms]));
  const matrix = {};
  ['A','B','C'].forEach(a => ['Fast','Medium','Slow'].forEach(f => {
    const key = `${a}-${f}`;
    const items = skuSummary.filter(r => abcBysku[r.sku]===a && fmsBysku[r.sku]===f);
    matrix[key] = { count:items.length, totalQty:items.reduce((s,r)=>s+r.totalQty,0), skus:items.map(r=>r.sku) };
  }));

  // ── Group × Location × Category Cross-Analysis ──────────────────────────────
  const hasCat = orders.some(o => o.category);
  const hasLoc = orders.some(o => o.dispatchLoc);
  const grpLocCatMap = {};
  orders.forEach(o => {
    const grp = o.orderType   || 'Unknown';
    const loc = o.dispatchLoc || 'Unspecified';
    const cat = o.category    || 'Unspecified';
    const key = `${grp}|||${loc}|||${cat}`;
    if (!grpLocCatMap[key]) grpLocCatMap[key] = { group:grp, location:loc, category:cat,
      lines:0, orders:new Set(), skus:new Set(), qty:0, volume:0 };
    const g = grpLocCatMap[key];
    g.lines++; g.qty += o.qty;
    if (o.orderNo) g.orders.add(o.orderNo);
    if (o.sku)     g.skus.add(o.sku);
    const m = masterMap.get(o.sku);
    if (m) g.volume += m.volume * o.qty / 1e9;
  });
  const grpLocCatSummary = Object.values(grpLocCatMap)
    .map(g => ({ group:g.group, location:g.location, category:g.category,
      lines:g.lines, uniqueOrders:g.orders.size, distinctSKUs:g.skus.size,
      totalQty:g.qty, totalVolume:+g.volume.toFixed(4) }))
    .sort((a,b) =>
      a.group.localeCompare(b.group) ||
      a.location.localeCompare(b.location) ||
      b.totalQty - a.totalQty);

  // Period days — distinct order dates across all orders
  const periodDays = Math.max(new Set(orders.map(o=>o.date).filter(d=>d)).size, 1);

  return { anomalies, orderSummary, skuSummary, abcData, fmsData, matrix,
    totVol, grpLocCatSummary, hasCat, hasLoc, periodDays };
}

// ─── PPT GENERATION ───────────────────────────────────────────────────────────

// ─── INVENTORY: PARSE ─────────────────────────────────────────────────────────
// Fixed order: SKU Code | Current Stock Qty | Warehouse Location* | Last Received Date*
function parseInventoryData(text) {
  const rows = parseTSV(text);
  const dataRows = rows.length > 0 && isHeaderRow(rows[0]) ? rows.slice(1) : rows;
  return dataRows.filter(r => r[0]).map(r => ({
    sku: r[0].trim(), stockQty: parseFloat(r[1]) || 0,
    location: r[2] ? r[2].trim() : '', lastRecvd: r[3] ? r[3].trim() : '',
  }));
}

// ─── INVENTORY: ANALYSE ───────────────────────────────────────────────────────
function analyseInventory(invData, analysis) {
  const { skuSummary, abcData, fmsData, periodDays } = analysis;
  const days = Math.max(periodDays || 30, 1);
  const abcBySku  = Object.fromEntries(abcData.map(r => [r.sku, r.abc]));
  const fmsBySku  = Object.fromEntries(fmsData.map(r => [r.sku, r.fms]));
  const demandBySku = Object.fromEntries(
    skuSummary.map(r => [r.sku, { totalQty:r.totalQty, daily:r.totalQty/days }])
  );
  const volBySku = Object.fromEntries(
    skuSummary.map(r => [r.sku, r.volumePerUnit || 0])
  );

  const results = invData.map(inv => {
    const d       = demandBySku[inv.sku];
    const abc     = abcBySku[inv.sku]  || 'N/A';
    const fms     = fmsBySku[inv.sku]  || 'N/A';
    const daily   = d ? +d.daily.toFixed(2) : 0;
    const hasDemand = !!d && daily > 0;
    const daysCov = hasDemand ? Math.round(inv.stockQty / daily) : null;
    let flag, priority;
    if (!d)           { flag='Dead Stock'; priority='None'; }
    else if (!hasDemand) { flag='No Demand'; priority='None'; }
    else if (daysCov < 7)  { flag='Critical'; priority=(abc==='A'||abc==='B')?'P1 — Immediate':'P2 — High'; }
    else if (daysCov < 30) { flag='Low';      priority=abc==='A'?'P2 — High':'P3 — Medium'; }
    else if (daysCov <=90) { flag='Adequate'; priority='P4 — Monitor'; }
    else                   { flag='Overstock';priority='None — Review'; }
    const volPerUnit = volBySku[inv.sku] || 0;
    const stockVolume = volPerUnit > 0 ? +(inv.stockQty * volPerUnit / 1e9).toFixed(4) : 0;
    return { sku:inv.sku, stockQty:inv.stockQty, stockVolume, location:inv.location,
      lastRecvd:inv.lastRecvd, daily, daysCov, flag, priority,
      abc, fms, abcFms:(abc!=='N/A'&&fms!=='N/A')?`${abc}-${fms}`:'—' };
  });

  const invSkus = new Set(invData.map(r=>r.sku));
  const notInInv = skuSummary
    .filter(r=>!invSkus.has(r.sku))
    .map(r=>({sku:r.sku,totalQty:r.totalQty,abc:abcBySku[r.sku]||'—',fms:fmsBySku[r.sku]||'—'}))
    .sort((a,b)=>b.totalQty-a.totalQty);

  return {
    results, notInInv, days,
    critical:  results.filter(r=>r.flag==='Critical').length,
    low:       results.filter(r=>r.flag==='Low').length,
    adequate:  results.filter(r=>r.flag==='Adequate').length,
    overstock: results.filter(r=>r.flag==='Overstock').length,
    dead:      results.filter(r=>r.flag==='Dead Stock').length,
  };
}

function generatePPT(mAnom, analysis, avgValues, invAnalysis) {
  const { orderSummary, skuSummary, abcData, fmsData, matrix,
    grpLocCatSummary, hasCat, hasLoc } = analysis;
  const today = new Date().toLocaleDateString('en-IN');
  const prs = new PptxGenJS();

  // Theme
  const PINK   = 'BE185D';
  const DARK   = '0F172A';
  const BLUE   = '0EA5E9';
  const GREEN  = '059669';
  const GRAY   = '64748B';
  const LGRAY  = 'F1F5F9';
  const WHITE  = 'FFFFFF';
  const AMBER  = 'D97706';

  prs.layout       = 'LAYOUT_WIDE';
  prs.theme        = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };
  prs.author       = 'DensiCube';
  prs.company      = 'DensiCube';
  prs.subject      = 'Order Analysis Report';

  const hdr = (sld, title, sub) => {
    sld.addShape(prs.ShapeType.rect, { x:0, y:0, w:'100%', h:1.2,
      fill:{ color: PINK } });
    sld.addText('DensiCube', { x:0.3, y:0.08, w:2, h:0.35,
      fontSize:10, color:WHITE, bold:true, fontFace:'Calibri' });
    sld.addText(title, { x:0.3, y:0.38, w:9, h:0.55,
      fontSize:22, color:WHITE, bold:true, fontFace:'Calibri' });
    if (sub) sld.addText(sub, { x:0.3, y:0.88, w:9, h:0.28,
      fontSize:11, color:'FBCFE8', fontFace:'Calibri' });
    sld.addText(today, { x:9.2, y:0.08, w:1.8, h:0.35,
      fontSize:9, color:'FBCFE8', align:'right', fontFace:'Calibri' });
  };

  const statBox = (sld, x, y, val, label, color) => {
    sld.addShape(prs.ShapeType.roundRect, { x, y, w:2.1, h:1.0,
      fill:{ color: LGRAY }, line:{ color:'E2E8F0', pt:1 }, rectRadius:0.1 });
    sld.addText(String(val), { x, y:y+0.08, w:2.1, h:0.5,
      fontSize:26, bold:true, color, align:'center', fontFace:'Calibri' });
    sld.addText(label, { x, y:y+0.6, w:2.1, h:0.32,
      fontSize:9, color:GRAY, align:'center', fontFace:'Calibri' });
  };

  // ── SLIDE 1: Title ──────────────────────────────────────────────────────────
  const s1 = prs.addSlide();
  s1.addShape(prs.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{ color:DARK } });
  s1.addShape(prs.ShapeType.rect, { x:0, y:0, w:0.18, h:'100%', fill:{ color:PINK } });
  s1.addText('DensiCube', { x:0.5, y:1.2, w:9, h:0.6,
    fontSize:14, color:'FBCFE8', bold:true, fontFace:'Calibri' });
  s1.addText('Order Analysis Report', { x:0.5, y:1.9, w:9, h:1.2,
    fontSize:36, color:WHITE, bold:true, fontFace:'Calibri' });
  s1.addText('Container Intelligence', { x:0.5, y:3.1, w:9, h:0.5,
    fontSize:13, color:'94A3B8', fontFace:'Calibri' });
  s1.addText(`Generated: ${today}`, { x:0.5, y:5.5, w:9, h:0.4,
    fontSize:11, color:'475569', fontFace:'Calibri' });
  const totLines = orderSummary.reduce((s,r)=>s+r.lines,0);
  const totQty   = orderSummary.reduce((s,r)=>s+r.totalQty,0);
  s1.addText(`${totLines.toLocaleString()} order lines  ·  ${skuSummary.length} SKUs  ·  ${totQty.toLocaleString()} units`, {
    x:0.5, y:5.9, w:9, h:0.4, fontSize:11, color:'64748B', fontFace:'Calibri' });

  // ── SLIDE 2: Key Metrics ────────────────────────────────────────────────────
  const s2 = prs.addSlide();
  hdr(s2, 'Key Metrics', 'High-level summary of your order data');
  const totOrders = orderSummary.reduce((s,r)=>s+r.uniqueOrders,0);
  const totVol    = skuSummary.reduce((s,r)=>s+r.totalVolume,0);
  statBox(s2, 0.4,  1.5, totLines.toLocaleString(),   'Total Order Lines',    PINK);
  statBox(s2, 2.7,  1.5, totOrders.toLocaleString(),  'Unique Orders',        BLUE);
  statBox(s2, 5.0,  1.5, skuSummary.length,            'Distinct SKUs',        GREEN);
  statBox(s2, 7.3,  1.5, totQty.toLocaleString(),      'Total Qty Ordered',    AMBER);
  statBox(s2, 0.4,  2.8, mAnom.length + analysis.anomalies.length, 'Data Anomalies', mAnom.length+analysis.anomalies.length>0?'BE185D':GREEN);
  statBox(s2, 2.7,  2.8, abcData.filter(r=>r.abc==='A').length,  'ABC-A SKUs',  GREEN);
  statBox(s2, 5.0,  2.8, fmsData.filter(r=>r.fms==='Fast').length,'Fast-Moving SKUs', BLUE);
  statBox(s2, 7.3,  2.8, totVol>0?totVol.toFixed(2)+'m³':'—',   'Total Volume', GRAY);

  // ── SLIDE 3: Order Summary by Location ─────────────────────────────────────
  const s3 = prs.addSlide();
  hdr(s3, 'Order Summary by Dispatch Location', 'Lines, orders and quantities per origin location');
  const locRows = orderSummary.slice(0,10).map(r => ([
    { text: r.dispatchLoc,                        options:{ bold:true, color:DARK } },
    { text: r.lines.toLocaleString(),             options:{ align:'center' } },
    { text: r.uniqueOrders.toLocaleString(),      options:{ align:'center' } },
    { text: r.totalQty.toLocaleString(),          options:{ align:'center', bold:true, color:PINK } },
    { text: r.avgQtyPerLine.toString(),           options:{ align:'center' } },
  ]));
  s3.addTable([
    [{ text:'Location', options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
     { text:'Lines',    options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
     { text:'Orders',   options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
     { text:'Total Qty',options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
     { text:'Avg Qty/Line', options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } }],
    ...locRows,
  ], { x:0.4, y:1.3, w:9.2, colW:[3,1.5,1.5,2,2.2],
    fontSize:11, border:{ type:'solid', color:'E2E8F0', pt:1 },
    rowH:0.35, autoPage:false });

  // ── SLIDE 4: ABC Analysis ───────────────────────────────────────────────────
  const s4 = prs.addSlide();
  hdr(s4, 'ABC Analysis', 'Ranked by shipping volume — A = top 70%, B = next 20%, C = bottom 10%');
  const abcClasses = ['A','B','C'];
  const abcColors  = { A:GREEN, B:AMBER, C:GRAY };
  abcClasses.forEach((cls, i) => {
    const items = abcData.filter(r=>r.abc===cls);
    const vol   = items.reduce((s,r)=>s+r.totalVolume,0);
    const totV  = abcData.reduce((s,r)=>s+r.totalVolume,0);
    const x = 0.4 + i * 3.2;
    s4.addShape(prs.ShapeType.roundRect, { x, y:1.3, w:3.0, h:2.8,
      fill:{ color: LGRAY }, line:{ color:'E2E8F0', pt:1 }, rectRadius:0.1 });
    s4.addText(`Class ${cls}`, { x, y:1.35, w:3.0, h:0.45,
      fontSize:18, bold:true, color:abcColors[cls], align:'center', fontFace:'Calibri' });
    s4.addText(`${items.length} SKUs`, { x, y:1.8, w:3.0, h:0.4,
      fontSize:22, bold:true, color:DARK, align:'center', fontFace:'Calibri' });
    s4.addText(`${totV>0?(vol/totV*100).toFixed(1):0}% of volume`, { x, y:2.2, w:3.0, h:0.35,
      fontSize:12, color:GRAY, align:'center', fontFace:'Calibri' });
    const topSkus = items.slice(0,4).map(r=>r.sku).join('\n');
    s4.addText(`Top SKUs:\n${topSkus||'—'}`, { x:x+0.1, y:2.6, w:2.8, h:1.3,
      fontSize:9, color:GRAY, fontFace:'Calibri', valign:'top' });
  });
  s4.addText('💡 Focus container planning on Class A SKUs — they drive most of your freight volume.',
    { x:0.4, y:4.3, w:9.2, h:0.5, fontSize:10, color:PINK, italic:true, fontFace:'Calibri' });

  // ── SLIDE 5: FMS Analysis ───────────────────────────────────────────────────
  const s5 = prs.addSlide();
  hdr(s5, 'FMS Analysis', 'Ranked by order frequency — Fast = top 33%, Medium = mid 33%, Slow = bottom 33%');
  const fmsClasses = ['Fast','Medium','Slow'];
  const fmsColors  = { Fast:GREEN, Medium:AMBER, Slow:'DC2626' };
  fmsClasses.forEach((cls, i) => {
    const items  = fmsData.filter(r=>r.fms===cls);
    const lines  = items.reduce((s,r)=>s+r.lines,0);
    const totL   = fmsData.reduce((s,r)=>s+r.lines,0);
    const x = 0.4 + i * 3.2;
    s5.addShape(prs.ShapeType.roundRect, { x, y:1.3, w:3.0, h:2.8,
      fill:{ color:LGRAY }, line:{ color:'E2E8F0', pt:1 }, rectRadius:0.1 });
    s5.addText(cls, { x, y:1.35, w:3.0, h:0.45,
      fontSize:18, bold:true, color:fmsColors[cls], align:'center', fontFace:'Calibri' });
    s5.addText(`${items.length} SKUs`, { x, y:1.8, w:3.0, h:0.4,
      fontSize:22, bold:true, color:DARK, align:'center', fontFace:'Calibri' });
    s5.addText(`${totL>0?(lines/totL*100).toFixed(1):0}% of order lines`, { x, y:2.2, w:3.0, h:0.35,
      fontSize:12, color:GRAY, align:'center', fontFace:'Calibri' });
    const topSkus = items.slice(0,4).map(r=>r.sku).join('\n');
    s5.addText(`Top SKUs:\n${topSkus||'—'}`, { x:x+0.1, y:2.6, w:2.8, h:1.3,
      fontSize:9, color:GRAY, fontFace:'Calibri', valign:'top' });
  });
  s5.addText('💡 Fast-moving SKUs need reliable stock at dispatch locations. Slow movers are candidates for batch consolidation.',
    { x:0.4, y:4.3, w:9.2, h:0.5, fontSize:10, color:PINK, italic:true, fontFace:'Calibri' });

  // ── SLIDE 6: ABC-FMS Matrix ─────────────────────────────────────────────────
  const s6 = prs.addSlide();
  hdr(s6, 'ABC-FMS Matrix', 'Volume (ABC) × Frequency (FMS) — find your priority SKUs');
  const matColors = {
    'A-Fast':'D1FAE5','A-Medium':'FEF9C3','A-Slow':'FEF3C7',
    'B-Fast':'DBEAFE','B-Medium':'F3F4F6','B-Slow':'F3F4F6',
    'C-Fast':'FEE2E2','C-Medium':'F3F4F6','C-Slow':'F3F4F6',
  };
  const matHdr = [
    { text:'', options:{ fill:{ color:DARK } } },
    ...['Fast','Medium','Slow'].map(f => ({ text:f, options:{ bold:true, color:WHITE, fill:{ color:DARK }, align:'center' } })),
  ];
  const matDataRows = ['A','B','C'].map(a => ([
    { text:`Class ${a}`, options:{ bold:true, color:abcColors[a], fill:{ color:LGRAY } } },
    ...['Fast','Medium','Slow'].map(f => {
      const cell = matrix[`${a}-${f}`];
      return { text: cell?.count ? `${cell.count} SKUs\n${cell.totalQty.toLocaleString()} units` : '—',
        options:{ fill:{ color:matColors[`${a}-${f}`] }, align:'center',
          fontSize:10, color: cell?.count ? DARK : '9CA3AF' } };
    }),
  ]));
  s6.addTable([matHdr,...matDataRows], { x:1.5, y:1.4, w:7, colW:[1.5,1.8,1.8,1.9],
    fontSize:11, border:{ type:'solid', color:'E2E8F0', pt:1 }, rowH:0.7, autoPage:false });
  s6.addText('A-Fast = Priority / Star products  |  A-Slow = High value, low frequency  |  C-Fast = Review MOQ  |  C-Slow = Rationalise',
    { x:0.4, y:4.5, w:9.2, h:0.5, fontSize:9, color:GRAY, italic:true, align:'center', fontFace:'Calibri' });

  // ── SLIDE 7: Group × Location × Category ────────────────────────────────────
  if (grpLocCatSummary?.length > 0) {
    const s7 = prs.addSlide();
    hdr(s7, 'Group × Location × Category', 'Top shipping combinations by order type, origin and product category');
    const top10 = grpLocCatSummary.slice(0, 10);
    const totGLC = grpLocCatSummary.reduce((s,r)=>s+r.totalQty,0);
    const glcRows = top10.map((r,i) => ([
      { text:(i+1).toString(),           options:{ align:'center', color:GRAY } },
      { text:r.group,                    options:{ bold:true, color:PINK } },
      { text:r.location||'—',           options:{ color:DARK } },
      { text:r.category||'—',           options:{ color:DARK } },
      { text:r.totalQty.toLocaleString(),options:{ align:'center', bold:true } },
      { text:totGLC>0?+(r.totalQty/totGLC*100).toFixed(1)+'%':'—', options:{ align:'center', color:GRAY } },
    ]));
    s7.addTable([
      [{ text:'#',        options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
       { text:'Group',    options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
       { text:'Location', options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
       { text:'Category', options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
       { text:'Total Qty',options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
       { text:'% of Total',options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } }],
      ...glcRows,
    ], { x:0.4, y:1.35, w:9.2, colW:[0.4,1.8,2.0,2.2,1.9,1.0],
      fontSize:10, border:{ type:'solid', color:'E2E8F0', pt:1 }, rowH:0.32, autoPage:false });
  }

  // ── SLIDE 8: Recommendations ────────────────────────────────────────────────
  const s8 = prs.addSlide();
  hdr(s8, 'Recommendations', 'Data-driven actions from your order analysis');
  const aSkus  = abcData.filter(r=>r.abc==='A').length;
  const fastSk = fmsData.filter(r=>r.fms==='Fast').length;
  const anomCount = mAnom.length + analysis.anomalies.length;
  const recs = [
    { icon:'📦', title:'Prioritise A-Class SKUs', body:`${aSkus} SKUs drive 70% of your freight volume. Ensure these are always container-optimised using the DensiCube Multi-SKU Planner.` },
    { icon:'🚀', title:'Fast-Moving SKU Readiness', body:`${fastSk} fast-moving SKUs have high order frequency. Keep them stocked at primary dispatch locations to avoid split shipments.` },
    { icon:'📊', title:'Consolidate Slow Movers', body:`C-class slow-moving SKUs are candidates for batch consolidation. Ship them together to reduce per-unit freight cost.` },
    { icon:'⚠️', title:'Fix Data Anomalies', body:anomCount>0?`${anomCount} anomalies found in your data. Fix missing dimensions and zero-quantity rows before the next shipment plan.`:'No anomalies found. Your data is clean — good for accurate planning.' },
    { icon:'🗺️', title:'Location-Based Planning', body:`Use dispatch location breakdown to allocate container types per origin. High-volume locations may benefit from dedicated container schedules.` },
  ];
  recs.forEach((r, i) => {
    const y = 1.35 + i * 0.82;
    s8.addShape(prs.ShapeType.roundRect, { x:0.4, y, w:9.2, h:0.72,
      fill:{ color:i%2===0?LGRAY:WHITE }, line:{ color:'E2E8F0', pt:1 }, rectRadius:0.06 });
    s8.addText(`${r.icon}  ${r.title}`, { x:0.6, y:y+0.06, w:9.0, h:0.28,
      fontSize:11, bold:true, color:DARK, fontFace:'Calibri' });
    s8.addText(r.body, { x:0.6, y:y+0.34, w:9.0, h:0.3,
      fontSize:9, color:GRAY, fontFace:'Calibri' });
  });

  // ── SLIDE 9: Assumptions ───────────────────────────────────────────────────
  const s9 = prs.addSlide();
  hdr(s9, 'Assumptions & Methodology', 'All data processing rules and imputation assumptions used in this report');
  const asmData = buildAssumptions(mAnom, analysis, avgValues);
  const asmDataRows = asmData.slice(1, 12); // header + first 11 rows fit on slide
  s9.addTable([
    [{ text:'Category',   options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
     { text:'Assumption', options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
     { text:'Value / Detail', options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
     { text:'Rationale', options:{ bold:true, color:WHITE, fill:{ color:PINK } } }],
    ...asmDataRows.map((r, i) => ([
      { text: r[0], options:{ bold:true, color:PINK,  fill:{ color: i%2===0?'F8FAFC':WHITE }, fontSize:9 } },
      { text: r[1], options:{ color:DARK, fill:{ color: i%2===0?'F8FAFC':WHITE }, fontSize:9 } },
      { text: r[2], options:{ color:'059669', fill:{ color: i%2===0?'F8FAFC':WHITE }, fontSize:9, bold:true } },
      { text: r[3], options:{ color:GRAY, fill:{ color: i%2===0?'F8FAFC':WHITE }, fontSize:8, italic:true } },
    ])),
  ], { x:0.4, y:1.35, w:9.2, colW:[1.4, 2.8, 2.4, 2.6],
    border:{ type:'solid', color:'E2E8F0', pt:1 }, rowH:0.3, autoPage:false });
  s9.addText('Full assumptions list available in Sheet 8 of the Excel report.',
    { x:0.4, y:6.8, w:9.2, h:0.28, fontSize:9, color:GRAY, italic:true, align:'center', fontFace:'Calibri' });


  // ── SLIDE 10: Inventory Highlights ─────────────────────────────────────────
  if (invAnalysis) {
    const s10 = prs.addSlide();
    hdr(s10, 'Inventory Analysis', `Coverage based on ${invAnalysis.days}-day order period`);

    // Summary stat boxes
    const invStats = [
      [invAnalysis.critical,  'Critical (<7 days)', 'BE185D'],
      [invAnalysis.low,       'Low (7-30 days)',    'D97706'],
      [invAnalysis.adequate,  'Adequate (30-90d)',  '059669'],
      [invAnalysis.overstock, 'Overstock (>90d)',   '0EA5E9'],
      [invAnalysis.dead,      'Dead Stock',         '64748B'],
    ];
    invStats.forEach(([val, label, col], i) => {
      const x = 0.4 + i * 1.88;
      s10.addShape(prs.ShapeType.roundRect, { x, y:1.35, w:1.75, h:0.9,
        fill:{ color:'F8FAFC' }, line:{ color:'E2E8F0', pt:1 }, rectRadius:0.08 });
      s10.addText(String(val), { x, y:1.38, w:1.75, h:0.45,
        fontSize:22, bold:true, color:col, align:'center', fontFace:'Calibri' });
      s10.addText(label, { x, y:1.83, w:1.75, h:0.35,
        fontSize:8, color:'64748B', align:'center', fontFace:'Calibri' });
    });

    // Top 10 reorder priority
    const reorderList = invAnalysis.results
      .filter(r=>r.flag==='Critical'||r.flag==='Low')
      .sort((a,b)=>(a.priority<b.priority?-1:1)||(a.daysCov||999)-(b.daysCov||999))
      .slice(0,10);
    if (reorderList.length > 0) {
      s10.addText('Reorder Priority List', { x:0.4, y:2.4, w:5.5, h:0.3,
        fontSize:11, bold:true, color:'0F172A', fontFace:'Calibri' });
      s10.addTable([
        [{ text:'SKU',        options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
         { text:'Stock Qty',  options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
         { text:'Vol (m³)',   options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
         { text:'Days Left',  options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } },
         { text:'Priority',   options:{ bold:true, color:WHITE, fill:{ color:PINK } } },
         { text:'ABC-FMS',    options:{ bold:true, color:WHITE, fill:{ color:PINK }, align:'center' } }],
        ...reorderList.map((r,i) => ([
          { text:r.sku, options:{ color:'0F172A', fill:{ color:i%2===0?'FFF1F2':'FFFFFF' }, fontSize:9 } },
          { text:r.stockQty.toLocaleString(), options:{ align:'center', fontSize:9 } },
          { text:r.stockVolume>0?String(r.stockVolume):'—', options:{ align:'center', color:'64748B', fontSize:9 } },
          { text:r.daysCov!=null?String(r.daysCov):'—', options:{ align:'center', bold:true, color:'BE185D', fontSize:9 } },
          { text:r.priority, options:{ bold:true, color:'BE185D', fontSize:9 } },
          { text:r.abcFms, options:{ align:'center', fontSize:9 } },
        ])),
      ], { x:0.4, y:2.7, w:5.5, colW:[1.8,1.0,0.9,0.9,1.5,0.9],
        border:{ type:'solid', color:'E2E8F0', pt:1 }, rowH:0.28, autoPage:false });
    }

    // ABC-FMS vs inventory coverage (right panel)
    s10.addText('ABC-FMS Coverage Summary', { x:6.1, y:2.4, w:3.5, h:0.3,
      fontSize:11, bold:true, color:'0F172A', fontFace:'Calibri' });
    const abcFmsGroups = {};
    invAnalysis.results.forEach(r => {
      const k = r.abcFms==='—' ? 'Unclassified' : r.abcFms;
      if (!abcFmsGroups[k]) abcFmsGroups[k] = { total:0, critical:0, adequate:0, overstock:0 };
      abcFmsGroups[k].total++;
      if (r.flag==='Critical'||r.flag==='Low') abcFmsGroups[k].critical++;
      else if (r.flag==='Adequate')             abcFmsGroups[k].adequate++;
      else if (r.flag==='Overstock')            abcFmsGroups[k].overstock++;
    });
    const afRows = Object.entries(abcFmsGroups)
      .sort(([a],[b])=>a.localeCompare(b)).slice(0,8);
    if (afRows.length > 0) {
      s10.addTable([
        [{ text:'Class',    options:{ bold:true, color:WHITE, fill:{ color:'0F172A' } } },
         { text:'SKUs',     options:{ bold:true, color:WHITE, fill:{ color:'0F172A' }, align:'center' } },
         { text:'⚠ Low',   options:{ bold:true, color:WHITE, fill:{ color:'0F172A' }, align:'center' } },
         { text:'✓ OK',    options:{ bold:true, color:WHITE, fill:{ color:'0F172A' }, align:'center' } }],
        ...afRows.map(([k,v],i)=>([
          { text:k, options:{ bold:true, color:'0F172A', fill:{ color:i%2===0?'F8FAFC':'FFFFFF' }, fontSize:9 } },
          { text:String(v.total),    options:{ align:'center', fontSize:9 } },
          { text:v.critical>0?String(v.critical):'—', options:{ align:'center', bold:v.critical>0, color:v.critical>0?'BE185D':'9CA3AF', fontSize:9 } },
          { text:v.adequate>0?String(v.adequate):'—', options:{ align:'center', bold:false, color:'059669', fontSize:9 } },
        ])),
      ], { x:6.1, y:2.7, w:3.5, colW:[1.4,0.7,0.7,0.7],
        border:{ type:'solid', color:'E2E8F0', pt:1 }, rowH:0.28, autoPage:false });
    }

    if (invAnalysis.dead > 0) {
      s10.addText(`⚠ ${invAnalysis.dead} dead stock SKU(s) — items in inventory with zero orders in the period. Consider liquidation or write-off.`,
        { x:0.4, y:6.8, w:9.2, h:0.3, fontSize:9, color:'64748B', italic:true, align:'center', fontFace:'Calibri' });
    }
  }

  // ── Save ────────────────────────────────────────────────────────────────────
  prs.writeFile({ fileName: `DensiCube_Insights_${today.replace(/\//g,'-')}.pptx` });
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────

function buildAssumptions(mAnom, analysis, avgValues) {
  const av = avgValues || {};
  const { skuSummary, abcData, fmsData } = analysis;
  const rows = [
    ['Category', 'Assumption', 'Value / Detail', 'Rationale'],
    // Master SKU assumptions
    ['Master Data', 'ABC classification threshold — Class A', 'Top 70% cumulative shipping volume', 'Standard Pareto principle for inventory prioritisation'],
    ['Master Data', 'ABC classification threshold — Class B', 'Next 20% cumulative volume (70-90%)', 'Standard Pareto principle'],
    ['Master Data', 'ABC classification threshold — Class C', 'Bottom 10% cumulative volume (90-100%)', 'Standard Pareto principle'],
    ['Master Data', 'FMS classification — Fast', 'Top 33% by order line frequency', 'Equal tertile split of order frequency'],
    ['Master Data', 'FMS classification — Medium', 'Middle 33% by order line frequency', 'Equal tertile split'],
    ['Master Data', 'FMS classification — Slow', 'Bottom 33% by order line frequency', 'Equal tertile split'],
    ['Master Data', 'Volume calculation', 'L × W × H (mm³) converted to m³ ÷ 1,000,000,000', 'Standard cubic volume formula'],
  ];

  // Avg fill assumptions
  if (av.filledL > 0) rows.push(['Master Data', `Missing Length — ${av.filledL} SKU(s)`, `Filled with dataset average: ${av.L} mm`, 'Average of all SKUs with valid length values']);
  if (av.filledW > 0) rows.push(['Master Data', `Missing Width — ${av.filledW} SKU(s)`, `Filled with dataset average: ${av.W} mm`, 'Average of all SKUs with valid width values']);
  if (av.filledH > 0) rows.push(['Master Data', `Missing Height — ${av.filledH} SKU(s)`, `Filled with dataset average: ${av.H} mm`, 'Average of all SKUs with valid height values']);
  if (av.filledWt > 0) rows.push(['Master Data', `Missing Weight — ${av.filledWt} SKU(s)`, `Filled with dataset average: ${av.weight} kg`, 'Average of all SKUs with valid weight values']);
  if (!av.filledL && !av.filledW && !av.filledH && !av.filledWt) rows.push(['Master Data', 'Missing dimensions/weight', 'None — all SKUs had complete data', '—']);

  rows.push(
    // Order data
    ['Order Data', 'Missing Order Type', 'Labelled as "Unknown"', 'Optional column — absence does not affect core analytics'],
    ['Order Data', 'Missing Category', 'Labelled as "Unspecified"', 'Optional column — analytics still run on available data'],
    ['Order Data', 'Missing Dispatch Location', 'Labelled as "Unspecified"', 'Optional column — analytics still run on available data'],
    ['Order Data', 'SKUs not in Master data', 'Flagged as anomaly — volume set to 0 for that SKU', 'Cannot compute freight volume without dimensions'],
    ['Order Data', 'Zero or negative qty', 'Flagged as high-severity anomaly', 'Negative quantities indicate data entry errors'],
    // Analytics
    ['Analytics', 'Order Summary grouping', 'Grouped by Dispatch Location', 'Location-based grouping is more actionable for logistics planning'],
    ['Analytics', 'Duplicate SKU in Master', 'First occurrence used — duplicate flagged as anomaly', 'Prevents double-counting in volume calculations'],
    ['Analytics', 'Total volume per SKU', 'Vol per unit × Total Qty ordered', 'Represents total freight volume demand for each SKU'],
  );
  return rows;
}

function exportReport(mAnom, analysis, avgValues, invAnalysis) {
  const { anomalies, orderSummary, skuSummary, abcData, fmsData, matrix, grpLocCatSummary, hasCat, hasLoc } = analysis;
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString();
  const ws = (data, cols) => {
    const s = XLSX.utils.aoa_to_sheet(data);
    if (cols) s['!cols'] = cols.map(w => ({ wch: w }));
    return s;
  };
  const allAnom = [...mAnom.map(a=>({...a,src:'Master'})), ...anomalies.map(a=>({...a,src:'Order'}))];

  XLSX.utils.book_append_sheet(wb, ws([
    ['DENSICUBE — DATA ANOMALY REPORT'], ['Generated:', today], ['Total issues:', allAnom.length], [],
    ['Source','Row','SKU','Field','Issue','Severity'],
    ...(allAnom.length ? allAnom.map(a=>[a.src,a.row,a.sku,a.field,a.issue,a.sev]) : [['—','—','—','—','No issues found','—']]),
    [], ['SUMMARY'],
    ['High severity:', allAnom.filter(a=>a.sev==='High').length],
    ['Medium severity:', allAnom.filter(a=>a.sev==='Medium').length],
  ], [10,6,22,14,55,10]), '1. Anomalies');

  XLSX.utils.book_append_sheet(wb, ws([
    ['ORDER SUMMARY BY DISPATCH LOCATION'], ['Generated:', today], [],
    ['Dispatch Location','Lines','Unique Orders','Distinct SKUs','Distinct Dates','Total Qty','Avg Qty/Line','Avg Lines/Order'],
    ...orderSummary.map(r=>[r.dispatchLoc,r.lines,r.uniqueOrders,r.distinctSKUs,r.distinctDates,r.totalQty,r.avgQtyPerLine,r.avgLinesPerOrder]),
    [], ['TOTAL', orderSummary.reduce((s,r)=>s+r.lines,0), orderSummary.reduce((s,r)=>s+r.uniqueOrders,0),'—','—',orderSummary.reduce((s,r)=>s+r.totalQty,0),'—','—'],
  ], [16,8,14,14,14,12,14,16]), '2. Order Summary');

  XLSX.utils.book_append_sheet(wb, ws([
    ['SKU SUMMARY'], [],
    ['SKU','Lines','Unique Orders','Total Qty','L mm','W mm','H mm','Wt/Box kg','Vol/Unit m³','Total Vol m³','First Date','Last Date','Category','Dispatch Location'],
    ...[...skuSummary].sort((a,b)=>b.totalQty-a.totalQty).map(r=>[r.sku,r.lines,r.uniqueOrders,r.totalQty,
      r.L||'—',r.W||'—',r.H||'—',r.weight||'—',
      r.volumePerUnit>0?+(r.volumePerUnit/1e9).toFixed(6):'—',
      r.totalVolume>0?+r.totalVolume.toFixed(4):'—',r.firstDate,r.lastDate,
      r.categories||'—',r.locations||'—']),
  ], [22,8,14,12,8,8,8,12,14,14,12,12,20,22]), '3. SKU Summary');

  XLSX.utils.book_append_sheet(wb, ws([
    ['ABC ANALYSIS — BY SHIPPING VOLUME'],['A=Top 70% volume | B=Next 20% | C=Bottom 10%'],[],
    ['Rank','SKU','Total Qty','Vol/Unit (m³)','Total Volume (m³)','Cumulative Vol %','ABC Class'],
    ...abcData.map((r,i)=>[i+1,r.sku,r.totalQty,
      r.volumePerUnit>0?+(r.volumePerUnit/1e9).toFixed(6):'—',
      r.totalVolume>0?+r.totalVolume.toFixed(4):'—',r.cumVolPct+'%',r.abc]),
    [], ['SUMMARY'],['Class','SKU Count','% SKUs','Total Volume (m³)','% Volume'],
    ...['A','B','C'].map(cls=>{
      const items=abcData.filter(r=>r.abc===cls);
      const vol=items.reduce((s,r)=>s+r.totalVolume,0);
      const totV=abcData.reduce((s,r)=>s+r.totalVolume,0);
      return[cls,items.length,(items.length/Math.max(1,abcData.length)*100).toFixed(1)+'%',
        vol.toFixed(3),totV>0?(vol/totV*100).toFixed(1)+'%':'0%'];
    }),
  ], [6,22,12,14,18,18,10]), '4. ABC Analysis');

  XLSX.utils.book_append_sheet(wb, ws([
    ['FMS ANALYSIS — BY ORDER LINES'],['Fast=Top 33% | Medium=Mid 33% | Slow=Bottom 33%'],[],
    ['Rank','SKU','Lines','Total Qty','Unique Orders','Cumulative Lines %','FMS Class'],
    ...fmsData.map((r,i)=>[i+1,r.sku,r.lines,r.totalQty,r.uniqueOrders,r.cumLinesPct+'%',r.fms]),
    [], ['SUMMARY'],['Class','SKU Count','% SKUs','Total Lines','% Lines'],
    ...['Fast','Medium','Slow'].map(cls=>{
      const items=fmsData.filter(r=>r.fms===cls);
      const lines=items.reduce((s,r)=>s+r.lines,0);
      const totL=fmsData.reduce((s,r)=>s+r.lines,0);
      return[cls,items.length,(items.length/Math.max(1,fmsData.length)*100).toFixed(1)+'%',
        lines,totL>0?(lines/totL*100).toFixed(1)+'%':'0%'];
    }),
  ], [6,22,8,12,14,20,10]), '5. FMS Analysis');

  const matRows = [
    ['ABC-FMS MATRIX'],['Volume (ABC) vs Frequency (FMS)'],[],
    ['','Fast','Medium','Slow'],
    ...['A','B','C'].map(a=>[a,...['Fast','Medium','Slow'].map(f=>{
      const c=matrix[`${a}-${f}`];
      return c?.count?`${c.count} SKUs | ${c.totalQty.toLocaleString()} units`:'—';
    })]),
    [],[  'INTERPRETATION'],
    ['A-Fast','Star products — highest priority for container planning'],
    ['A-Slow','High value, low frequency — consider batch shipping'],
    ['C-Fast','Many small orders — review minimum order quantities'],
    ['C-Slow','Candidates for rationalisation or consolidation'],
  ];
  XLSX.utils.book_append_sheet(wb, ws(matRows, [16,35,35,35]), '6. ABC-FMS Matrix');

  // ── Sheet 7: Group × Location × Category ────────────────────────────────────
  const totGLC = grpLocCatSummary.reduce((s,r) => s+r.totalQty, 0);
  const totGLCVol = grpLocCatSummary.reduce((s,r) => s+r.totalVolume, 0);

  // Flat detail list
  const glcRows = [
    ['GROUP × LOCATION × CATEGORY ANALYSIS'],
    ['Group = Order Type | Location = Dispatch Origin | Category = Product Category'],
    [hasCat&&hasLoc ? '' : !hasCat&&!hasLoc ? 'Note: Order Type (col 6) and Category (col 7) not pasted — showing Location breakdown only'
      : !hasCat ? 'Note: Category (col 6) not pasted' : 'Note: Dispatch Location (col 7) not pasted'],
    ['Generated:', today], [],
    ['Group (Order Type)','Dispatch Location','Product Category','Lines','Unique Orders','Distinct SKUs','Total Qty','Total Volume (m³)','% of Total Qty'],
    ...grpLocCatSummary.map(r => [r.group, r.location, r.category, r.lines,
      r.uniqueOrders, r.distinctSKUs, r.totalQty,
      r.totalVolume>0 ? r.totalVolume : '—',
      totGLC>0 ? +(r.totalQty/totGLC*100).toFixed(1)+'%' : '—']),
    [], ['TOTAL','—','—',
      grpLocCatSummary.reduce((s,r)=>s+r.lines,0), '—', '—',
      totGLC, +totGLCVol.toFixed(4), '100%'],
  ];

  // Sub-totals by Group
  const groups = [...new Set(grpLocCatSummary.map(r=>r.group))];
  glcRows.push([], ['SUB-TOTALS BY GROUP']);
  glcRows.push(['Group','Lines','Total Qty','Total Volume (m³)','% of Total']);
  groups.forEach(grp => {
    const rows = grpLocCatSummary.filter(r=>r.group===grp);
    const qty = rows.reduce((s,r)=>s+r.totalQty,0);
    const vol = rows.reduce((s,r)=>s+r.totalVolume,0);
    glcRows.push([grp, rows.reduce((s,r)=>s+r.lines,0), qty, +vol.toFixed(4),
      totGLC>0 ? +(qty/totGLC*100).toFixed(1)+'%' : '—']);
  });

  XLSX.utils.book_append_sheet(wb, ws(glcRows,
    [20,22,22,8,14,14,12,16,14]), '7. Group x Location x Category');

  // ── Sheet 8: Assumptions ────────────────────────────────────────────────────
  const asmRows = buildAssumptions(mAnom, analysis, avgValues);
  XLSX.utils.book_append_sheet(wb, ws([
    ['ASSUMPTIONS & METHODOLOGY'],
    ['This sheet documents all assumptions made during data processing and analysis'],
    ['Generated:', today], [],
    ...asmRows,
  ], [16, 42, 36, 52]), '8. Assumptions');

  // ── Sheet 9: Inventory Analysis ─────────────────────────────────────────────
  if (invAnalysis) {
    const flagIcon = { Critical:'🔴', Low:'🟡', Adequate:'🟢', Overstock:'🔵', 'Dead Stock':'⚫', 'No Demand':'⚪' };
    const inv9Rows = [
      ['INVENTORY ANALYSIS'],
      [`Coverage period: ${invAnalysis.days} days based on distinct order dates`],
      ['Generated:', today], [],
      ['SKU Code','Current Stock','Stock Volume (m³)','Daily Demand','Days Coverage','Status','Reorder Priority','ABC','FMS','ABC-FMS','Warehouse','Last Received'],
      ...invAnalysis.results
        .sort((a,b)=>(a.priority<b.priority?-1:1)||(a.daysCov||9999)-(b.daysCov||9999))
        .map(r=>[r.sku, r.stockQty, r.stockVolume>0?r.stockVolume:'—', r.daily>0?r.daily:'—',
          r.daysCov!=null?r.daysCov:'—',
          (flagIcon[r.flag]||'')+(r.flag),
          r.priority, r.abc, r.fms, r.abcFms,
          r.location||'—', r.lastRecvd||'—']),
      [], ['SUMMARY'],
      ['Critical (<7 days)', invAnalysis.critical],
      ['Low (7-30 days)',     invAnalysis.low],
      ['Adequate (30-90 days)', invAnalysis.adequate],
      ['Overstock (>90 days)', invAnalysis.overstock],
      ['Dead Stock (no orders)', invAnalysis.dead],
    ];
    if (invAnalysis.notInInv.length > 0) {
      inv9Rows.push([], ['SKUs IN ORDERS BUT NOT IN INVENTORY (no stock recorded)'],
        ['SKU', 'Total Qty Ordered', 'ABC', 'FMS'],
        ...invAnalysis.notInInv.slice(0,20).map(r=>[r.sku, r.totalQty, r.abc, r.fms]));
    }
    XLSX.utils.book_append_sheet(wb, ws(inv9Rows,
      [22,14,16,14,14,14,18,8,10,12,18,16]), '9. Inventory Analysis');
  }

  XLSX.writeFile(wb, `DensiCube_Analysis_${today.replace(/\//g,'-')}.xlsx`);
}

// ─── COMPONENT ────────────────────────────────────────────────────────────────
export default function OrderAnalyserTool() {
  const [mText,    setMText]    = useState('');
  const [mAnom,    setMAnom]    = useState([]);
  const [masterMap,setMasterMap]= useState(new Map());
  const [mAvgValues,setMAvgValues]= useState(null);
  const [mDone,    setMDone]    = useState(false);
  const [mError,   setMError]   = useState('');
  const [mStats,   setMStats]   = useState(null);

  const [oText,     setOText]    = useState('');
  const [analysis,  setAnalysis] = useState(null);
  const [oError,    setOError]   = useState('');

  const [invText,    setInvText]   = useState('');
  const [invAnalysis,setInvAnalysis]= useState(null);
  const [invDone,    setInvDone]   = useState(false);
  const [invError,   setInvError]  = useState('');

  // ── Step 1: Process master SKU ─────────────────────────────────────────────
  const processMaster = () => {
    if (!mText.trim()) { setMError('Paste your Master SKU data first.'); return; }
    setMError(''); setMDone(false); setMAnom([]); setMasterMap(new Map()); setMStats(null);
    const { skus, anomalies, masterMap: mm, avgValues } = parseMasterSKU(mText);
    if (!skus.length && anomalies.filter(a=>a.sev==='High').length > 0) {
      setMError('No valid SKU rows found. Check that your columns are in the correct order: SKU Name | L | W | H | Weight');
      setMAnom(anomalies); return;
    }
    setMAnom(anomalies); setMasterMap(mm); setMAvgValues(avgValues);
    setMStats({ total: mm.size, withDims: skus.filter(s=>s.L&&s.W&&s.H).length,
      withWeight: skus.filter(s=>s.weight>0).length,
      avgFilled: anomalies.filter(a=>a.issue.includes('average')).length });
    setMDone(true);
  };

  // ── Step 2: Process order data ─────────────────────────────────────────────
  const processOrder = () => {
    if (!oText.trim()) { setOError('Paste your Order data first.'); return; }
    setOError(''); setAnalysis(null);
    const result = parseOrderData(oText, masterMap);
    if (!result.orderSummary.length) {
      setOError('No valid order rows found. Check that your columns are in the correct order: Order No | Order Type | SKU Code | Qty | Date');
      return;
    }
    setAnalysis(result);
  };



  // ── Step 3: Process inventory data ────────────────────────────────────────
  const processInventory = () => {
    if (!invText.trim()) { setInvError('Paste your inventory data first.'); return; }
    setInvError(''); setInvDone(false); setInvAnalysis(null);
    const invData = parseInventoryData(invText);
    if (!invData.length) {
      setInvError('No valid rows found. Check column order: SKU Code | Current Stock Qty');
      return;
    }
    const result = analyseInventory(invData, analysis);
    setInvAnalysis(result);
    setInvDone(true);
  };

  // ── Shared helpers ─────────────────────────────────────────────────────────
  const sev = s => <span style={{ background:s==='High'?'#fee2e2':s==='Medium'?'#fef9c3':'#f0fdf4',
    color:s==='High'?'#991b1b':s==='Medium'?'#854d0e':'#166534', padding:'2px 7px',
    borderRadius:'99px', fontSize:'11px', fontWeight:'700' }}>{s}</span>;

  const abcB = c => { const m={A:['#dcfce7','#166534'],B:['#fef9c3','#854d0e'],C:['#f3f4f6','#6b7280']};
    const[bg,col]=m[c]||m.C; return<span style={{background:bg,color:col,padding:'2px 8px',borderRadius:'99px',fontSize:'12px',fontWeight:'700'}}>{c}</span>;};
  const fmsB = c => { const m={Fast:['#dcfce7','#166534'],Medium:['#fef9c3','#854d0e'],Slow:['#fee2e2','#991b1b']};
    const[bg,col]=m[c]||m.Slow; return<span style={{background:bg,color:col,padding:'2px 8px',borderRadius:'99px',fontSize:'12px',fontWeight:'700'}}>{c}</span>;};

  const colHint = (cols) => (
    <div style={{ display:'flex', gap:'4px', flexWrap:'wrap', marginBottom:'8px' }}>
      {cols.map((col, i) => (
        <span key={i} style={{ background:'#f1f5f9', border:'1px solid #e2e8f0',
          borderRadius:'6px', padding:'3px 8px', fontSize:'12px', fontWeight:'600',
          color:'#475569', display:'flex', alignItems:'center', gap:'4px' }}>
          <span style={{ background:'#be185d', color:'#fff', borderRadius:'50%', width:'15px',
            height:'15px', display:'inline-flex', alignItems:'center', justifyContent:'center',
            fontSize:'9px', fontWeight:'800', flexShrink:0 }}>{i+1}</span>
          {col}
        </span>
      ))}
    </div>
  );

  const textarea = (value, onChange, placeholder) => (
    <textarea value={value} onChange={e => onChange(e.target.value)}
      placeholder={placeholder}
      style={{ width:'100%', height:'150px', border:'1px solid #e2e8f0', borderRadius:'8px',
        padding:'10px 12px', fontSize:'12px', fontFamily:'monospace', resize:'vertical',
        outline:'none', boxSizing:'border-box', color:'#374151', lineHeight:'1.6' }}/>
  );

  const stepCircle = (n, done) => (
    <div style={{ width:'32px', height:'32px', borderRadius:'50%', flexShrink:0,
      background:done?'#166534':'#be185d', display:'flex', alignItems:'center',
      justifyContent:'center', color:'#fff', fontWeight:'800', fontSize:'14px' }}>
      {done ? '✓' : n}
    </div>
  );

  return (
    <div>
      <div style={S.sectionDesc}>
        Paste your Master SKU data and Order data directly from Excel — no file upload needed.
        Columns must be in the order shown. Get a 6-sheet Excel report with anomaly detection,
        order summary, ABC analysis, FMS classification, and ABC-FMS matrix.
      </div>

      {/* ── STEP 1: MASTER SKU ──────────────────────────────────────────────── */}
      <div style={S.card}>
        <div style={{ display:'flex', alignItems:'center', gap:'12px', marginBottom:'16px' }}>
          {stepCircle(1, mDone)}
          <div>
            <div style={S.cardTitle}>Master SKU Data</div>
            <div style={{ fontSize:'12px', color:'#6b7280' }}>
              Copy your SKU master data from Excel and paste below
            </div>
          </div>
        </div>

        {colHint(['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight/Box (kg)'])}
        <div style={{ fontSize:'12px', color:'#6b7280', marginBottom:'8px' }}>
          In Excel: arrange columns in this order → select all data cells (including header) → Ctrl+C → paste below
        </div>

        {textarea(mText, setMText,
          'Paste Master SKU data here (Ctrl+V)\n\nExample:\nSKU Name\tLength\tWidth\tHeight\tWeight\nProduct A\t300\t200\t150\t2.5\nProduct B\t450\t320\t200\t4.0\nProduct C\t250\t180\t120\t1.8')}

        {mError && <div style={{ ...S.error, marginTop:'8px' }}>⚠ {mError}</div>}

        <button onClick={processMaster} disabled={!mText.trim()}
          style={{ marginTop:'10px', width:'100%', padding:'10px',
            background: mText.trim() ? '#be185d' : '#e2e8f0',
            color: mText.trim() ? '#fff' : '#9ca3af',
            border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'13px',
            cursor: mText.trim() ? 'pointer' : 'not-allowed', fontFamily:'inherit' }}>
          ▶ Validate Master SKU Data
        </button>

        {/* Results */}
        {mDone && mStats && (
          <div style={{ marginTop:'14px' }}>
            <div style={{ display:'flex', gap:'10px', marginBottom:'12px' }}>
              {[['Total SKUs', mStats.total, '#eff6ff','#1d4ed8'],
                ['With Dimensions', mStats.withDims, '#f0fdf4','#166534'],
                ['Avg Values Used', mStats.avgFilled, mStats.avgFilled?'#fffbeb':'#f0fdf4', mStats.avgFilled?'#d97706':'#166534'],
                ['Issues Found', mAnom.filter(a=>!a.issue.includes('average')).length, mAnom.filter(a=>!a.issue.includes('average')).length?'#fff1f2':'#f0fdf4', mAnom.filter(a=>!a.issue.includes('average')).length?'#be185d':'#166534'],
              ].map(([l,v,bg,col]) => (
                <div key={l} style={{ flex:1, background:bg, borderRadius:'8px', padding:'10px', textAlign:'center' }}>
                  <div style={{ fontSize:'20px', fontWeight:'800', color:col }}>{v}</div>
                  <div style={{ fontSize:'10px', color:'#6b7280', marginTop:'2px', fontWeight:'600', textTransform:'uppercase' }}>{l}</div>
                </div>))}
            </div>

            {mAnom.length > 0 && (
              <div style={{ border:'1px solid #e2e8f0', borderRadius:'8px', overflow:'hidden', maxHeight:'200px', overflowY:'auto' }}>
                <table style={{ width:'100%', borderCollapse:'collapse', fontSize:'12px' }}>
                  <thead><tr>
                    {['Row','SKU','Field','Issue','Sev'].map(h => (
                      <th key={h} style={{ padding:'7px 10px', background:'#f8fafc',
                        borderBottom:'1px solid #e2e8f0', textAlign:'left',
                        fontWeight:'700', fontSize:'11px', color:'#6b7280',
                        textTransform:'uppercase' }}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {mAnom.slice(0,50).map((a,i) => (
                      <tr key={i} style={{ background:i%2?'#fafbfc':'#fff' }}>
                        <td style={{ padding:'7px 10px', color:'#6b7280' }}>{a.row}</td>
                        <td style={{ padding:'7px 10px', fontWeight:'600' }}>{a.sku}</td>
                        <td style={{ padding:'7px 10px' }}>{a.field}</td>
                        <td style={{ padding:'7px 10px', color:'#374151' }}>{a.issue}</td>
                        <td style={{ padding:'7px 10px' }}>{sev(a.sev)}</td>
                      </tr>))}
                  </tbody>
                </table>
                {mAnom.length > 50 && <div style={{ padding:'8px 12px', fontSize:'11px', color:'#9ca3af' }}>Showing 50 of {mAnom.length} — full list in Excel report</div>}
              </div>
            )}
            {mAnom.length === 0 && <div style={{ background:'#f0fdf4', border:'1px solid #bbf7d0', borderRadius:'8px', padding:'10px 14px', fontSize:'13px', color:'#166534' }}>✓ No anomalies found</div>}
          </div>
        )}
      </div>

      {/* ── STEP 2: ORDER DATA ──────────────────────────────────────────────── */}
      <div style={{ ...S.card, opacity:mDone?1:0.5, pointerEvents:mDone?'auto':'none' }}>
        <div style={{ display:'flex', alignItems:'center', gap:'12px', marginBottom:'16px' }}>
          {stepCircle(2, !!analysis)}
          <div>
            <div style={S.cardTitle}>Order Data</div>
            <div style={{ fontSize:'12px', color:'#6b7280' }}>
              {mDone ? 'Copy your order data from Excel and paste below' : 'Complete Step 1 first'}
            </div>
          </div>
        </div>

        {mDone && (<>
          {colHint(['Order No', 'Dispatch Location', 'SKU Code', 'Qty', 'Date', 'Order Type (optional)', 'Category (optional)'])}
          <div style={{ fontSize:'12px', color:'#6b7280', marginBottom:'8px' }}>
            Order Type examples: STO, Customer, Export etc. Category examples: Refrigerator, TV, Washing Machine. Both are optional — paste 5 columns minimum.
          </div>

          {textarea(oText, setOText,
            'Paste Order data here (Ctrl+V)\n\nColumns 1-5 required | Columns 6-7 optional:\nOrder No | Order Type | SKU Code | Qty | Date | Category | Dispatch Location\n\nExample (with optional columns):\n1001\tCustomer\tSKU-001\t500\t01/06/2024\tRefrigerator\tMumbai\n1002\tSTO\tSKU-002\t200\t02/06/2024\tWashing Machine\tAhmedabad')}

          {oError && <div style={{ ...S.error, marginTop:'8px' }}>⚠ {oError}</div>}

          <button onClick={processOrder} disabled={!oText.trim()}
            style={{ marginTop:'10px', width:'100%', padding:'10px',
              background: oText.trim() ? '#be185d' : '#e2e8f0',
              color: oText.trim() ? '#fff' : '#9ca3af',
              border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'13px',
              cursor: oText.trim() ? 'pointer' : 'not-allowed', fontFamily:'inherit' }}>
            ▶ Run Analysis
          </button>
        </>)}

        {/* Results */}
        {analysis && (
          <div style={{ marginTop:'16px' }}>
            {analysis.anomalies.length > 0 && (
              <div style={{ background:'#fff1f2', border:'1px solid #fecaca', borderRadius:'8px',
                padding:'10px 14px', marginBottom:'14px', fontSize:'13px',
                display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                <span style={{ color:'#991b1b', fontWeight:'600' }}>
                  ⚠ {analysis.anomalies.length} order anomalies
                  ({analysis.anomalies.filter(a=>a.sev==='High').length} high severity)
                </span>
                <span style={{ fontSize:'12px', color:'#9ca3af' }}>Full list in Excel report</span>
              </div>
            )}

            {/* Summary cards */}
            <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:'12px', marginBottom:'16px' }}>
              {[['Order Lines', analysis.orderSummary.reduce((s,r)=>s+r.lines,0).toLocaleString(), '#eff6ff','#1d4ed8'],
                ['Unique Orders', analysis.orderSummary.reduce((s,r)=>s+r.uniqueOrders,0).toLocaleString(), '#f5f3ff','#6d28d9'],
                ['Distinct SKUs', analysis.skuSummary.length.toLocaleString(), '#f0fdf4','#166534'],
                ['Total Qty', analysis.orderSummary.reduce((s,r)=>s+r.totalQty,0).toLocaleString(), '#fff7ed','#c2410c'],
              ].map(([l,v,bg,col]) => (
                <div key={l} style={{ background:bg, borderRadius:'8px', padding:'12px', textAlign:'center' }}>
                  <div style={{ fontSize:'20px', fontWeight:'800', color:col }}>{v}</div>
                  <div style={{ fontSize:'10px', color:'#6b7280', marginTop:'3px', fontWeight:'600', textTransform:'uppercase' }}>{l}</div>
                </div>))}
            </div>

            {/* Order summary table */}
            <div style={{ border:'1px solid #e2e8f0', borderRadius:'8px', overflow:'hidden', marginBottom:'14px' }}>
              <div style={{ padding:'8px 14px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0', fontWeight:'700', fontSize:'12px' }}>Order Summary by Dispatch Location</div>
              <div style={{ overflowX:'auto' }}>
                <table style={{ width:'100%', borderCollapse:'collapse', fontSize:'12px' }}>
                  <thead><tr>
                    {['Dispatch Location','Lines','Unique Orders','Distinct SKUs','Dates','Total Qty','Avg Qty/Line','Avg Lines/Order'].map(h => (
                      <th key={h} style={{ padding:'7px 10px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0',
                        textAlign:'left', fontWeight:'600', fontSize:'11px', color:'#6b7280',
                        textTransform:'uppercase', whiteSpace:'nowrap' }}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {analysis.orderSummary.map((r,i) => (
                      <tr key={i} style={{ background:i%2?'#fafbfc':'#fff' }}>
                        <td style={{ padding:'7px 10px', fontWeight:'700' }}>{r.dispatchLoc}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right' }}>{r.lines.toLocaleString()}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right' }}>{r.uniqueOrders.toLocaleString()}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right' }}>{r.distinctSKUs.toLocaleString()}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right' }}>{r.distinctDates.toLocaleString()}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right', fontWeight:'600' }}>{r.totalQty.toLocaleString()}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right' }}>{r.avgQtyPerLine}</td>
                        <td style={{ padding:'7px 10px', textAlign:'right' }}>{r.avgLinesPerOrder}</td>
                      </tr>))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* ABC + FMS summary */}
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px', marginBottom:'14px' }}>
              <div style={{ border:'1px solid #e2e8f0', borderRadius:'8px', overflow:'hidden' }}>
                <div style={{ padding:'8px 14px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0', fontWeight:'700', fontSize:'12px' }}>ABC (by shipping volume)</div>
                {['A','B','C'].map(cls => {
                  const items = analysis.abcData.filter(r=>r.abc===cls);
                  const vol = items.reduce((s,r)=>s+r.totalVolume,0);
                  return (<div key={cls} style={{ padding:'8px 14px', borderBottom:'1px solid #f1f5f9',
                    display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                    <div style={{ display:'flex', gap:'8px', alignItems:'center' }}>
                      {abcB(cls)}<span style={{ fontSize:'12px', color:'#374151' }}>{items.length} SKUs</span>
                    </div>
                    <span style={{ fontSize:'12px', color:'#6b7280' }}>
                      {analysis.totVol>0?(vol/analysis.totVol*100).toFixed(1):'0'}% of volume
                    </span>
                  </div>);})}
              </div>
              <div style={{ border:'1px solid #e2e8f0', borderRadius:'8px', overflow:'hidden' }}>
                <div style={{ padding:'8px 14px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0', fontWeight:'700', fontSize:'12px' }}>FMS (by order lines)</div>
                {['Fast','Medium','Slow'].map(cls => {
                  const items = analysis.fmsData.filter(r=>r.fms===cls);
                  const lines = items.reduce((s,r)=>s+r.lines,0);
                  const tot = analysis.fmsData.reduce((s,r)=>s+r.lines,0);
                  return (<div key={cls} style={{ padding:'8px 14px', borderBottom:'1px solid #f1f5f9',
                    display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                    <div style={{ display:'flex', gap:'8px', alignItems:'center' }}>
                      {fmsB(cls)}<span style={{ fontSize:'12px', color:'#374151' }}>{items.length} SKUs</span>
                    </div>
                    <span style={{ fontSize:'12px', color:'#6b7280' }}>
                      {tot>0?(lines/tot*100).toFixed(1):'0'}% of lines
                    </span>
                  </div>);})}
              </div>
            </div>

            {/* ABC-FMS Matrix */}
            <div style={{ border:'1px solid #e2e8f0', borderRadius:'8px', overflow:'hidden', marginBottom:'14px' }}>
              <div style={{ padding:'8px 14px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0', fontWeight:'700', fontSize:'12px' }}>ABC-FMS Matrix</div>
              <table style={{ width:'100%', borderCollapse:'collapse', fontSize:'12px' }}>
                <thead><tr>
                  <th style={{ padding:'8px 12px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0', fontSize:'11px', color:'#6b7280' }}></th>
                  {['Fast','Medium','Slow'].map(f => (
                    <th key={f} style={{ padding:'8px 12px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0',
                      textAlign:'center', fontSize:'11px', color:'#6b7280', fontWeight:'700' }}>{f}</th>))}
                </tr></thead>
                <tbody>
                  {['A','B','C'].map(a => (
                    <tr key={a}>
                      <td style={{ padding:'10px 12px', fontWeight:'800' }}>{abcB(a)}</td>
                      {['Fast','Medium','Slow'].map(f => {
                        const c = analysis.matrix[`${a}-${f}`]; const n = c?.count||0;
                        return (<td key={f} style={{ padding:'10px 12px', textAlign:'center',
                          background:n>0?(a==='A'&&f==='Fast'?'#dcfce7':a==='C'&&f==='Slow'?'#fee2e2':'#fff'):'#f8fafc' }}>
                          {n>0?(<><div style={{ fontWeight:'700', fontSize:'14px' }}>{n}</div>
                            <div style={{ fontSize:'10px', color:'#9ca3af' }}>{c?.totalQty?.toLocaleString()||0} units</div></>)
                            :<span style={{ color:'#d1d5db' }}>—</span>}
                        </td>);})}
                    </tr>))}
                </tbody>
              </table>
            </div>

            <div style={{ display:'flex', gap:'10px', flexWrap:'wrap' }}>
              <button onClick={() => exportReport(mAnom, analysis, mAvgValues, invAnalysis)}
                style={{ ...S.btnPrimary, flex:1, background:'linear-gradient(135deg,#be185d,#9d174d)' }}>
                ⬇ Download Excel Report
              </button>
              <button onClick={() => generatePPT(mAnom, analysis, mAvgValues, invAnalysis)}
                style={{ ...S.btnPrimary, flex:1, background:'linear-gradient(135deg,#7c3aed,#6d28d9)' }}>
                📊 Download PPT Insights
              </button>
            </div>
            <div style={{ fontSize:'11px', color:'#9ca3af', textAlign:'center', marginTop:'6px' }}>
              Anomalies · Order Summary · SKU Summary · ABC · FMS · ABC-FMS Matrix · Group×Location×Category · Assumptions{invAnalysis ? ' · Inventory Analysis' : ''}
            </div>
          </div>
        )}
      </div>

      {/* ── STEP 3: INVENTORY (Optional) ──────────────────────────────────── */}
      <div style={{ ...S.card, opacity:analysis?1:0.4, pointerEvents:analysis?'auto':'none' }}>
        <div style={{ display:'flex', alignItems:'center', gap:'12px', marginBottom:'16px' }}>
          {stepCircle(3, invDone)}
          <div>
            <div style={S.cardTitle}>Inventory Data
              <span style={{ fontSize:'12px', fontWeight:'500', color:'#059669', marginLeft:'8px' }}>Optional</span>
            </div>
            <div style={{ fontSize:'12px', color:'#6b7280' }}>
              {analysis ? 'Paste current stock levels to get coverage analysis, reorder priority and ABC-FMS inventory cross' : 'Complete Step 2 first'}
            </div>
          </div>
        </div>

        {analysis && (<>
          {colHint(['SKU Code', 'Current Stock Qty', 'Warehouse Location (optional)', 'Last Received Date (optional)'])}
          <div style={{ fontSize:'12px', color:'#6b7280', marginBottom:'8px' }}>
            SKU codes must match your Master SKU data. Coverage calculated over <strong>{analysis.periodDays} order days</strong> detected from your order data.
          </div>
          <textarea value={invText} onChange={e => setInvText(e.target.value)}
            placeholder={'Paste inventory data here (Ctrl+V)\n\nExample:\nSKU Code\tCurrent Stock\tWarehouse\tLast Received\nSKU-001\t2500\tMumbai WH\t15/06/2024\nSKU-002\t180\tAhmedabad WH\t10/06/2024\nSKU-003\t0\tMumbai WH\t01/05/2024'}
            style={{ width:'100%', height:'140px', border:'1px solid #e2e8f0', borderRadius:'8px',
              padding:'10px 12px', fontSize:'12px', fontFamily:'monospace', resize:'vertical',
              outline:'none', boxSizing:'border-box', color:'#374151', lineHeight:'1.6' }}/>

          {invError && <div style={{ ...S.error, marginTop:'8px' }}>⚠ {invError}</div>}

          <button onClick={processInventory} disabled={!invText.trim()}
            style={{ marginTop:'10px', width:'100%', padding:'10px',
              background: invText.trim() ? '#059669' : '#e2e8f0',
              color: invText.trim() ? '#fff' : '#9ca3af',
              border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'13px',
              cursor: invText.trim() ? 'pointer' : 'not-allowed', fontFamily:'inherit' }}>
            ▶ Analyse Inventory
          </button>

          {invAnalysis && (
            <div style={{ marginTop:'16px' }}>
              {/* Stat boxes */}
              <div style={{ display:'grid', gridTemplateColumns:'repeat(5,1fr)', gap:'8px', marginBottom:'14px' }}>
                {[['🔴','Critical',invAnalysis.critical,'#fff1f2','#be185d'],
                  ['🟡','Low',invAnalysis.low,'#fffbeb','#d97706'],
                  ['🟢','Adequate',invAnalysis.adequate,'#f0fdf4','#166534'],
                  ['🔵','Overstock',invAnalysis.overstock,'#eff6ff','#1d4ed8'],
                  ['⚫','Dead Stock',invAnalysis.dead,'#f8fafc','#374151'],
                ].map(([icon,label,val,bg,col]) => (
                  <div key={label} style={{ background:bg, borderRadius:'8px', padding:'10px', textAlign:'center' }}>
                    <div style={{ fontSize:'18px', fontWeight:'800', color:col }}>{val}</div>
                    <div style={{ fontSize:'9px', color:'#6b7280', marginTop:'2px', fontWeight:'600', textTransform:'uppercase' }}>{icon} {label}</div>
                  </div>))}
              </div>

              {/* Reorder table */}
              {invAnalysis.results.filter(r=>r.flag==='Critical'||r.flag==='Low').length > 0 && (
                <div style={{ border:'1px solid #e2e8f0', borderRadius:'8px', overflow:'hidden', marginBottom:'10px' }}>
                  <div style={{ padding:'8px 14px', background:'#fff1f2', borderBottom:'1px solid #fecaca',
                    fontWeight:'700', fontSize:'12px', color:'#991b1b' }}>
                    ⚠ Reorder Priority ({invAnalysis.critical + invAnalysis.low} SKUs)
                  </div>
                  <div style={{ overflowX:'auto', maxHeight:'200px', overflowY:'auto' }}>
                    <table style={{ width:'100%', borderCollapse:'collapse', fontSize:'12px' }}>
                      <thead><tr>
                        {['SKU','Stock Qty','Stock Vol (m³)','Days Left','Demand/Day','Priority','ABC-FMS'].map(h => (
                          <th key={h} style={{ padding:'6px 10px', background:'#f8fafc',
                            borderBottom:'1px solid #e2e8f0', textAlign:'left',
                            fontWeight:'700', fontSize:'11px', color:'#6b7280',
                            textTransform:'uppercase', whiteSpace:'nowrap' }}>{h}</th>))}
                      </tr></thead>
                      <tbody>
                        {invAnalysis.results
                          .filter(r=>r.flag==='Critical'||r.flag==='Low')
                          .sort((a,b)=>(a.daysCov||999)-(b.daysCov||999))
                          .slice(0,15)
                          .map((r,i) => (
                          <tr key={i} style={{ background:i%2?'#fafbfc':'#fff' }}>
                            <td style={{ padding:'6px 10px', fontWeight:'600' }}>{r.sku}</td>
                            <td style={{ padding:'6px 10px', textAlign:'right' }}>{r.stockQty.toLocaleString()}</td>
                            <td style={{ padding:'6px 10px', textAlign:'right', color:'#6b7280' }}>{r.stockVolume>0?r.stockVolume:'—'}</td>
                            <td style={{ padding:'6px 10px', textAlign:'right', fontWeight:'700',
                              color:r.flag==='Critical'?'#be185d':'#d97706' }}>
                              {r.daysCov != null ? r.daysCov : '—'}
                            </td>
                            <td style={{ padding:'6px 10px', textAlign:'right', color:'#6b7280' }}>{r.daily}</td>
                            <td style={{ padding:'6px 10px', fontWeight:'600', color:'#be185d', fontSize:'11px' }}>{r.priority}</td>
                            <td style={{ padding:'6px 10px' }}>{r.abcFms}</td>
                          </tr>))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {(invAnalysis.dead > 0 || invAnalysis.notInInv.length > 0) && (
                <div style={{ background:'#f8fafc', border:'1px solid #e2e8f0',
                  borderRadius:'8px', padding:'10px 14px', fontSize:'13px', color:'#6b7280' }}>
                  {invAnalysis.dead > 0 && <div>⚫ <strong>{invAnalysis.dead} dead stock SKU(s)</strong> — in inventory with zero orders this period</div>}
                  {invAnalysis.notInInv.length > 0 && <div style={{ marginTop:'4px' }}>⚠ <strong>{invAnalysis.notInInv.length} ordered SKU(s)</strong> not found in inventory — check if stock is missing</div>}
                </div>
              )}
            </div>
          )}
        </>)}
      </div>
    </div>
  );
}
