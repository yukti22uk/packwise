// â”€â”€â”€ ORDER ANALYSER TOOL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Step 1: Paste Master SKU data (fixed column order) â†’ validate + flag anomalies
// Step 2: Paste Order data (fixed column order) â†’ full analytics â†’ 6-sheet Excel
// No AI column mapping â€” user pastes in expected order shown on screen.
import { useState } from 'react';
import * as XLSX from 'xlsx';
import { S } from '../components/styles.jsx';

// â”€â”€â”€ PARSE PASTED TSV â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

// â”€â”€â”€ MASTER SKU: PARSE & VALIDATE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Fixed order: SKU Name | Length (mm) | Width (mm) | Height (mm) | Weight/Box (kg)
function parseMasterSKU(text) {
  const rows = parseTSV(text);
  const dataRows = rows.length > 0 && isHeaderRow(rows[0]) ? rows.slice(1) : rows;
  const skus = [], anomalies = [], seen = new Set();

  dataRows.forEach((r, i) => {
    const rowNum = i + (rows.length > dataRows.length ? 2 : 1);
    const name = r[0] || '';
    const L = parseFloat(r[1]) || 0;
    const W = parseFloat(r[2]) || 0;
    const H = parseFloat(r[3]) || 0;
    const weight = parseFloat(r[4]) || 0;

    if (!name.trim()) { anomalies.push({ row: rowNum, sku: 'â€”', field: 'SKU Name', issue: 'Missing SKU name', sev: 'High' }); return; }
    if (seen.has(name)) anomalies.push({ row: rowNum, sku: name, field: 'SKU', issue: `Duplicate SKU (appears more than once)`, sev: 'Medium' });
    else seen.add(name);
    if (!L) anomalies.push({ row: rowNum, sku: name, field: 'Length', issue: 'Missing or zero length', sev: 'High' });
    if (!W) anomalies.push({ row: rowNum, sku: name, field: 'Width',  issue: 'Missing or zero width',  sev: 'High' });
    if (!H) anomalies.push({ row: rowNum, sku: name, field: 'Height', issue: 'Missing or zero height', sev: 'High' });
    if (!weight) anomalies.push({ row: rowNum, sku: name, field: 'Weight', issue: 'Missing or zero weight', sev: 'Medium' });

    skus.push({ name, L, W, H, weight, volume: L * W * H });
  });

  const masterMap = new Map(skus.map(s => [s.name, s]));
  return { skus, anomalies, masterMap };
}

// â”€â”€â”€ ORDER DATA: PARSE & ANALYSE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Fixed order: Order No | Order Type | SKU Code | Qty | Date | Category* | Dispatch Location*
// * = optional columns
function parseOrderData(text, masterMap) {
  const rows = parseTSV(text);
  const dataRows = rows.length > 0 && isHeaderRow(rows[0]) ? rows.slice(1) : rows;
  const orders = [], anomalies = [];

  dataRows.forEach((r, i) => {
    const rowNum      = i + (rows.length > dataRows.length ? 2 : 1);
    const orderNo     = r[0] || '';
    const orderType   = r[1] || 'Unknown';
    const sku         = r[2] || '';
    const qty         = parseFloat(r[3]) || 0;
    const date        = r[4] || '';
    const category    = r[5] ? r[5].trim() : '';   // optional
    const dispatchLoc = r[6] ? r[6].trim() : '';   // optional

    if (!orderNo) anomalies.push({ row: rowNum, sku, field: 'Order No',  issue: 'Missing order/invoice number', sev: 'High' });
    if (!sku)     anomalies.push({ row: rowNum, sku: 'â€”', field: 'SKU',  issue: 'Missing SKU code', sev: 'High' });
    if (qty <= 0) anomalies.push({ row: rowNum, sku, field: 'Qty',       issue: `Zero or negative qty: ${r[3]}`, sev: 'High' });
    if (!date)    anomalies.push({ row: rowNum, sku, field: 'Date',      issue: 'Missing date', sev: 'Medium' });
    if (sku && masterMap.size > 0 && !masterMap.has(sku))
      anomalies.push({ row: rowNum, sku, field: 'Master Data', issue: `SKU not in master â€” dimensions missing`, sev: 'High' });

    orders.push({ orderNo, orderType, sku, qty, date, category, dispatchLoc });
  });

  // Order summary by type
  const typeMap = {};
  orders.forEach(o => {
    if (!typeMap[o.orderType]) typeMap[o.orderType] = { lines:0, orders:new Set(), skus:new Set(), dates:new Set(), qty:0 };
    const g = typeMap[o.orderType];
    g.lines++; g.qty += o.qty;
    if (o.orderNo) g.orders.add(o.orderNo);
    if (o.sku)     g.skus.add(o.sku);
    if (o.date)    g.dates.add(o.date);
  });
  const orderSummary = Object.entries(typeMap).map(([type, g]) => ({
    orderType: type, lines: g.lines, uniqueOrders: g.orders.size,
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
      firstDate:ds[0]||'â€”', lastDate:ds[ds.length-1]||'â€”',
      categories:[...g.categories].join(', ')||'â€”',
      locations:[...g.locations].join(', ')||'â€”',
      L:m.L, W:m.W, H:m.H, weight:m.weight,
      volumePerUnit:m.volume, totalVolume:m.volume*g.qty/1e9 };
  });

  // â”€â”€ Category Analysis â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const hasCat = orders.some(o => o.category);
  const catMap = {};
  orders.forEach(o => {
    const cat = o.category || 'Unspecified';
    if (!catMap[cat]) catMap[cat] = { lines:0, orders:new Set(), skus:new Set(), qty:0, volume:0, locations:new Set() };
    const g = catMap[cat];
    g.lines++; g.qty += o.qty;
    if (o.orderNo)     g.orders.add(o.orderNo);
    if (o.sku)         g.skus.add(o.sku);
    if (o.dispatchLoc) g.locations.add(o.dispatchLoc);
    const m = masterMap.get(o.sku);
    if (m) g.volume += m.volume * o.qty / 1e9;
  });
  const categorySummary = Object.entries(catMap)
    .map(([cat, g]) => ({ category:cat, lines:g.lines, uniqueOrders:g.orders.size,
      distinctSKUs:g.skus.size, totalQty:g.qty, totalVolume:+g.volume.toFixed(4),
      locations:[...g.locations].join(', ')||'â€”' }))
    .sort((a,b) => b.totalQty - a.totalQty);
  const totCatVol = categorySummary.reduce((s,r) => s+r.totalVolume, 0);
  categorySummary.forEach(r => {
    r.volPct = totCatVol > 0 ? +(r.totalVolume/totCatVol*100).toFixed(1) : 0;
  });

  // â”€â”€ Dispatch Location Analysis â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const hasLoc = orders.some(o => o.dispatchLoc);
  const locMap = {};
  orders.forEach(o => {
    const loc = o.dispatchLoc || 'Unspecified';
    if (!locMap[loc]) locMap[loc] = { lines:0, orders:new Set(), skus:new Set(), qty:0, volume:0, categories:new Set() };
    const g = locMap[loc];
    g.lines++; g.qty += o.qty;
    if (o.orderNo)  g.orders.add(o.orderNo);
    if (o.sku)      g.skus.add(o.sku);
    if (o.category) g.categories.add(o.category);
    const m = masterMap.get(o.sku);
    if (m) g.volume += m.volume * o.qty / 1e9;
  });
  const locationSummary = Object.entries(locMap)
    .map(([loc, g]) => ({ location:loc, lines:g.lines, uniqueOrders:g.orders.size,
      distinctSKUs:g.skus.size, totalQty:g.qty, totalVolume:+g.volume.toFixed(4),
      categories:[...g.categories].join(', ')||'â€”' }))
    .sort((a,b) => b.totalQty - a.totalQty);
  const totLocVol = locationSummary.reduce((s,r) => s+r.totalVolume, 0);
  locationSummary.forEach(r => {
    r.volPct = totLocVol > 0 ? +(r.totalVolume/totLocVol*100).toFixed(1) : 0;
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

  // â”€â”€ Location Ã— Category Cross-Analysis â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const locCatMap = {};
  orders.forEach(o => {
    const loc = o.dispatchLoc || 'Unspecified';
    const cat = o.category    || 'Unspecified';
    const key = `${loc}|||${cat}`;
    if (!locCatMap[key]) locCatMap[key] = { location:loc, category:cat,
      lines:0, orders:new Set(), skus:new Set(), qty:0, volume:0 };
    const g = locCatMap[key];
    g.lines++; g.qty += o.qty;
    if (o.orderNo) g.orders.add(o.orderNo);
    if (o.sku)     g.skus.add(o.sku);
    const m = masterMap.get(o.sku);
    if (m) g.volume += m.volume * o.qty / 1e9;
  });
  const locCatSummary = Object.values(locCatMap)
    .map(g => ({ location:g.location, category:g.category,
      lines:g.lines, uniqueOrders:g.orders.size, distinctSKUs:g.skus.size,
      totalQty:g.qty, totalVolume:+g.volume.toFixed(4) }))
    .sort((a,b) => a.location.localeCompare(b.location) || b.totalQty - a.totalQty);

  // Pivot: unique locations and categories for matrix
  const allLocations = [...new Set(locCatSummary.map(r => r.location))];
  const allCategories = [...new Set(locCatSummary.map(r => r.category))];

  return { anomalies, orderSummary, skuSummary, abcData, fmsData, matrix, totVol,
    categorySummary, locationSummary, locCatSummary, allLocations, allCategories,
    hasCat, hasLoc };
}

// â”€â”€â”€ EXCEL EXPORT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function exportReport(mAnom, analysis) {
  const { anomalies, orderSummary, skuSummary, abcData, fmsData, matrix, categorySummary, locationSummary, hasCat, hasLoc } = analysis;
  const wb = XLSX.utils.book_new();
  const today = new Date().toLocaleDateString();
  const ws = (data, cols) => {
    const s = XLSX.utils.aoa_to_sheet(data);
    if (cols) s['!cols'] = cols.map(w => ({ wch: w }));
    return s;
  };
  const allAnom = [...mAnom.map(a=>({...a,src:'Master'})), ...anomalies.map(a=>({...a,src:'Order'}))];

  XLSX.utils.book_append_sheet(wb, ws([
    ['DENSICUBE â€” DATA ANOMALY REPORT'], ['Generated:', today], ['Total issues:', allAnom.length], [],
    ['Source','Row','SKU','Field','Issue','Severity'],
    ...(allAnom.length ? allAnom.map(a=>[a.src,a.row,a.sku,a.field,a.issue,a.sev]) : [['â€”','â€”','â€”','â€”','No issues found','â€”']]),
    [], ['SUMMARY'],
    ['High severity:', allAnom.filter(a=>a.sev==='High').length],
    ['Medium severity:', allAnom.filter(a=>a.sev==='Medium').length],
  ], [10,6,22,14,55,10]), '1. Anomalies');

  XLSX.utils.book_append_sheet(wb, ws([
    ['ORDER SUMMARY BY TYPE'], ['Generated:', today], [],
    ['Order Type','Lines','Unique Orders','Distinct SKUs','Distinct Dates','Total Qty','Avg Qty/Line','Avg Lines/Order'],
    ...orderSummary.map(r=>[r.orderType,r.lines,r.uniqueOrders,r.distinctSKUs,r.distinctDates,r.totalQty,r.avgQtyPerLine,r.avgLinesPerOrder]),
    [], ['TOTAL', orderSummary.reduce((s,r)=>s+r.lines,0), orderSummary.reduce((s,r)=>s+r.uniqueOrders,0),'â€”','â€”',orderSummary.reduce((s,r)=>s+r.totalQty,0),'â€”','â€”'],
  ], [16,8,14,14,14,12,14,16]), '2. Order Summary');

  XLSX.utils.book_append_sheet(wb, ws([
    ['SKU SUMMARY'], [],
    ['SKU','Lines','Unique Orders','Total Qty','L mm','W mm','H mm','Wt/Box kg','Vol/Unit mÂ³','Total Vol mÂ³','First Date','Last Date','Category','Dispatch Location'],
    ...[...skuSummary].sort((a,b)=>b.totalQty-a.totalQty).map(r=>[r.sku,r.lines,r.uniqueOrders,r.totalQty,
      r.L||'â€”',r.W||'â€”',r.H||'â€”',r.weight||'â€”',
      r.volumePerUnit>0?+(r.volumePerUnit/1e9).toFixed(6):'â€”',
      r.totalVolume>0?+r.totalVolume.toFixed(4):'â€”',r.firstDate,r.lastDate,
      r.categories||'â€”',r.locations||'â€”']),
  ], [22,8,14,12,8,8,8,12,14,14,12,12,20,22]), '3. SKU Summary');

  XLSX.utils.book_append_sheet(wb, ws([
    ['ABC ANALYSIS â€” BY SHIPPING VOLUME'],['A=Top 70% volume | B=Next 20% | C=Bottom 10%'],[],
    ['Rank','SKU','Total Qty','Vol/Unit (mÂ³)','Total Volume (mÂ³)','Cumulative Vol %','ABC Class'],
    ...abcData.map((r,i)=>[i+1,r.sku,r.totalQty,
      r.volumePerUnit>0?+(r.volumePerUnit/1e9).toFixed(6):'â€”',
      r.totalVolume>0?+r.totalVolume.toFixed(4):'â€”',r.cumVolPct+'%',r.abc]),
    [], ['SUMMARY'],['Class','SKU Count','% SKUs','Total Volume (mÂ³)','% Volume'],
    ...['A','B','C'].map(cls=>{
      const items=abcData.filter(r=>r.abc===cls);
      const vol=items.reduce((s,r)=>s+r.totalVolume,0);
      const totV=abcData.reduce((s,r)=>s+r.totalVolume,0);
      return[cls,items.length,(items.length/Math.max(1,abcData.length)*100).toFixed(1)+'%',
        vol.toFixed(3),totV>0?(vol/totV*100).toFixed(1)+'%':'0%'];
    }),
  ], [6,22,12,14,18,18,10]), '4. ABC Analysis');

  XLSX.utils.book_append_sheet(wb, ws([
    ['FMS ANALYSIS â€” BY ORDER LINES'],['Fast=Top 33% | Medium=Mid 33% | Slow=Bottom 33%'],[],
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
      return c?.count?`${c.count} SKUs | ${c.totalQty.toLocaleString()} units`:'â€”';
    })]),
    [],[  'INTERPRETATION'],
    ['A-Fast','Star products â€” highest priority for container planning'],
    ['A-Slow','High value, low frequency â€” consider batch shipping'],
    ['C-Fast','Many small orders â€” review minimum order quantities'],
    ['C-Slow','Candidates for rationalisation or consolidation'],
  ];
  XLSX.utils.book_append_sheet(wb, ws(matRows, [16,35,35,35]), '6. ABC-FMS Matrix');

  // â”€â”€ Sheet 7: Group Ã— Location Ã— Category â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const totGLC = grpLocCatSummary.reduce((s,r) => s+r.totalQty, 0);
  const totGLCVol = grpLocCatSummary.reduce((s,r) => s+r.totalVolume, 0);

  // Flat detail list
  const glcRows = [
    ['GROUP Ã— LOCATION Ã— CATEGORY ANALYSIS'],
    ['Group = Order Type | Location = Dispatch Origin | Category = Product Category'],
    [hasCat&&hasLoc ? '' : !hasCat&&!hasLoc ? 'Note: Category (col 6) and Dispatch Location (col 7) not pasted â€” showing Order Type breakdown only'
      : !hasCat ? 'Note: Category (col 6) not pasted' : 'Note: Dispatch Location (col 7) not pasted'],
    ['Generated:', today], [],
    ['Group (Order Type)','Dispatch Location','Product Category','Lines','Unique Orders','Distinct SKUs','Total Qty','Total Volume (mÂ³)','% of Total Qty'],
    ...grpLocCatSummary.map(r => [r.group, r.location, r.category, r.lines,
      r.uniqueOrders, r.distinctSKUs, r.totalQty,
      r.totalVolume>0 ? r.totalVolume : 'â€”',
      totGLC>0 ? +(r.totalQty/totGLC*100).toFixed(1)+'%' : 'â€”']),
    [], ['TOTAL','â€”','â€”',
      grpLocCatSummary.reduce((s,r)=>s+r.lines,0), 'â€”', 'â€”',
      totGLC, +totGLCVol.toFixed(4), '100%'],
  ];

  // Sub-totals by Group
  const groups = [...new Set(grpLocCatSummary.map(r=>r.group))];
  glcRows.push([], ['SUB-TOTALS BY GROUP']);
  glcRows.push(['Group','Lines','Total Qty','Total Volume (mÂ³)','% of Total']);
  groups.forEach(grp => {
    const rows = grpLocCatSummary.filter(r=>r.group===grp);
    const qty = rows.reduce((s,r)=>s+r.totalQty,0);
    const vol = rows.reduce((s,r)=>s+r.totalVolume,0);
    glcRows.push([grp, rows.reduce((s,r)=>s+r.lines,0), qty, +vol.toFixed(4),
      totGLC>0 ? +(qty/totGLC*100).toFixed(1)+'%' : 'â€”']);
  });

  XLSX.utils.book_append_sheet(wb, ws(glcRows,
    [20,22,22,8,14,14,12,16,14]), '7. Group x Location x Category');

  XLSX.writeFile(wb, `DensiCube_Analysis_${today.replace(/\//g,'-')}.xlsx`);
}

// â”€â”€â”€ COMPONENT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
export default function OrderAnalyserTool() {
  const [mText,    setMText]    = useState('');
  const [mAnom,    setMAnom]    = useState([]);
  const [masterMap,setMasterMap]= useState(new Map());
  const [mDone,    setMDone]    = useState(false);
  const [mError,   setMError]   = useState('');
  const [mStats,   setMStats]   = useState(null);

  const [oText,     setOText]    = useState('');
  const [analysis,  setAnalysis] = useState(null);
  const [oError,    setOError]   = useState('');

  // â”€â”€ Step 1: Process master SKU â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const processMaster = () => {
    if (!mText.trim()) { setMError('Paste your Master SKU data first.'); return; }
    setMError(''); setMDone(false); setMAnom([]); setMasterMap(new Map()); setMStats(null);
    const { skus, anomalies, masterMap: mm } = parseMasterSKU(mText);
    if (!skus.length && anomalies.filter(a=>a.sev==='High').length > 0) {
      setMError('No valid SKU rows found. Check that your columns are in the correct order: SKU Name | L | W | H | Weight');
      setMAnom(anomalies); return;
    }
    setMAnom(anomalies); setMasterMap(mm);
    setMStats({ total: mm.size, withDims: skus.filter(s=>s.L&&s.W&&s.H).length,
      withWeight: skus.filter(s=>s.weight>0).length });
    setMDone(true);
  };

  // â”€â”€ Step 2: Process order data â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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


  // â”€â”€ Shared helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const sev = s => <span style={{ background:s==='High'?'#fee2e2':'#fef9c3',
    color:s==='High'?'#991b1b':'#854d0e', padding:'2px 7px',
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
      {done ? 'âœ“' : n}
    </div>
  );

  return (
    <div>
      <div style={S.sectionDesc}>
        Paste your Master SKU data and Order data directly from Excel â€” no file upload needed.
        Columns must be in the order shown. Get a 6-sheet Excel report with anomaly detection,
        order summary, ABC analysis, FMS classification, and ABC-FMS matrix.
      </div>

      {/* â”€â”€ STEP 1: MASTER SKU â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */}
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
          In Excel: arrange columns in this order â†’ select all data cells (including header) â†’ Ctrl+C â†’ paste below
        </div>

        {textarea(mText, setMText,
          'Paste Master SKU data here (Ctrl+V)\n\nExample:\nSKU Name\tLength\tWidth\tHeight\tWeight\nProduct A\t300\t200\t150\t2.5\nProduct B\t450\t320\t200\t4.0\nProduct C\t250\t180\t120\t1.8')}

        {mError && <div style={{ ...S.error, marginTop:'8px' }}>âš  {mError}</div>}

        <button onClick={processMaster} disabled={!mText.trim()}
          style={{ marginTop:'10px', width:'100%', padding:'10px',
            background: mText.trim() ? '#be185d' : '#e2e8f0',
            color: mText.trim() ? '#fff' : '#9ca3af',
            border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'13px',
            cursor: mText.trim() ? 'pointer' : 'not-allowed', fontFamily:'inherit' }}>
          â–¶ Validate Master SKU Data
        </button>

        {/* Results */}
        {mDone && mStats && (
          <div style={{ marginTop:'14px' }}>
            <div style={{ display:'flex', gap:'10px', marginBottom:'12px' }}>
              {[['Total SKUs', mStats.total, '#eff6ff','#1d4ed8'],
                ['With Dimensions', mStats.withDims, '#f0fdf4','#166534'],
                ['Issues Found', mAnom.length, mAnom.length?'#fff1f2':'#f0fdf4', mAnom.length?'#be185d':'#166534'],
                ['High Severity', mAnom.filter(a=>a.sev==='High').length, '#fff1f2','#991b1b'],
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
                {mAnom.length > 50 && <div style={{ padding:'8px 12px', fontSize:'11px', color:'#9ca3af' }}>Showing 50 of {mAnom.length} â€” full list in Excel report</div>}
              </div>
            )}
            {mAnom.length === 0 && <div style={{ background:'#f0fdf4', border:'1px solid #bbf7d0', borderRadius:'8px', padding:'10px 14px', fontSize:'13px', color:'#166534' }}>âœ“ No anomalies found</div>}
          </div>
        )}
      </div>

      {/* â”€â”€ STEP 2: ORDER DATA â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */}
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
          {colHint(['Order No', 'Order Type', 'SKU Code', 'Qty', 'Date', 'Category (optional)', 'Dispatch Location (optional)'])}
          <div style={{ fontSize:'12px', color:'#6b7280', marginBottom:'8px' }}>
            Order Type examples: STO, Customer, Export etc. Category examples: Refrigerator, TV, Washing Machine. Dispatch Location is origin city/warehouse.
          </div>

          {textarea(oText, setOText,
            'Paste Order data here (Ctrl+V)\n\nColumns 1-5 required | Columns 6-7 optional:\nOrder No | Order Type | SKU Code | Qty | Date | Category | Dispatch Location\n\nExample (with optional columns):\n1001\tCustomer\tSKU-001\t500\t01/06/2024\tRefrigerator\tMumbai\n1002\tSTO\tSKU-002\t200\t02/06/2024\tWashing Machine\tAhmedabad')}

          {oError && <div style={{ ...S.error, marginTop:'8px' }}>âš  {oError}</div>}

          <button onClick={processOrder} disabled={!oText.trim()}
            style={{ marginTop:'10px', width:'100%', padding:'10px',
              background: oText.trim() ? '#be185d' : '#e2e8f0',
              color: oText.trim() ? '#fff' : '#9ca3af',
              border:'none', borderRadius:'8px', fontWeight:'700', fontSize:'13px',
              cursor: oText.trim() ? 'pointer' : 'not-allowed', fontFamily:'inherit' }}>
            â–¶ Run Analysis
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
                  âš  {analysis.anomalies.length} order anomalies
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
              <div style={{ padding:'8px 14px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0', fontWeight:'700', fontSize:'12px' }}>Order Summary by Type</div>
              <div style={{ overflowX:'auto' }}>
                <table style={{ width:'100%', borderCollapse:'collapse', fontSize:'12px' }}>
                  <thead><tr>
                    {['Order Type','Lines','Unique Orders','Distinct SKUs','Dates','Total Qty','Avg Qty/Line','Avg Lines/Order'].map(h => (
                      <th key={h} style={{ padding:'7px 10px', background:'#f8fafc', borderBottom:'1px solid #e2e8f0',
                        textAlign:'left', fontWeight:'600', fontSize:'11px', color:'#6b7280',
                        textTransform:'uppercase', whiteSpace:'nowrap' }}>{h}</th>))}
                  </tr></thead>
                  <tbody>
                    {analysis.orderSummary.map((r,i) => (
                      <tr key={i} style={{ background:i%2?'#fafbfc':'#fff' }}>
                        <td style={{ padding:'7px 10px', fontWeight:'700' }}>{r.orderType}</td>
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
                            :<span style={{ color:'#d1d5db' }}>â€”</span>}
                        </td>);})}
                    </tr>))}
                </tbody>
              </table>
            </div>

            <button onClick={() => exportReport(mAnom, analysis)}
              style={{ ...S.btnPrimary, background:'linear-gradient(135deg,#be185d,#9d174d)' }}>
              â¬‡ Download 6-Sheet Excel Report
            </button>
            <div style={{ fontSize:'11px', color:'#9ca3af', textAlign:'center', marginTop:'6px' }}>
              Anomalies Â· Order Summary Â· SKU Summary Â· ABC Â· FMS Â· ABC-FMS Matrix Â· GroupÃ—LocationÃ—Category
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
