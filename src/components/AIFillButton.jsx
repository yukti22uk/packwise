// ─── AI FILL FROM EXCEL ───────────────────────────────────────────────────────
// Reads any Excel file → Claude identifies columns & extracts SKU data → fills form.
// mode="multisku"  : extract up to 8 SKUs (name/L/W/H/weight/qty)
// mode="shipment"  : same, emphasises qty-to-ship column
// mode="bulk"      : detect column layout only (for large files, doesn't extract rows)

import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

export default function AIFillButton({ onFill, mode = 'multisku', label }) {
  const [open, setOpen]         = useState(false);
  const [file, setFile]         = useState(null);
  const [fileName, setFileName] = useState('');
  const [hint, setHint]         = useState('');
  const [loading, setLoading]   = useState(false);
  const [result, setResult]     = useState(null);
  const [error, setError]       = useState('');
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef(null);

  const reset = () => {
    setFile(null); setFileName(''); setHint('');
    setLoading(false); setResult(null); setError('');
  };
  const close = () => { setOpen(false); reset(); };
  const loadFile = (f) => {
    if (!f) return;
    setFile(f); setFileName(f.name); setResult(null); setError('');
  };

  const run = async () => {
    if (!file) { setError('Upload a file first.'); return; }
    setLoading(true); setError(''); setResult(null);
    try {
      // ── Parse Excel ──────────────────────────────────────────────────────
      const buf  = await file.arrayBuffer();
      const wb   = XLSX.read(buf, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const preview = rows.slice(0, 40)
        .map((r, i) => `Row ${i + 1}: ${Array.isArray(r) ? r.join('\t') : ''}`)
        .join('\n');

      // ── Build prompt ─────────────────────────────────────────────────────
      const modeInstr = {
        multisku: 'Extract up to 8 DISTINCT SKU types (different box sizes). If more exist, pick the 8 with largest quantities.',
        shipment: 'Extract SKUs with shipping quantities. "qty" = total units to ship for each SKU. Up to 8 SKUs.',
        bulk:     'Identify ONLY the column positions for name, length, width, height, weight, quantity. Return an empty skus array — do NOT extract rows.',
      };

      const prompt = `You are a data extraction assistant for a logistics tool.

FILE: ${fileName} | Sheets: ${wb.SheetNames.join(', ')} | Total rows: ${rows.length}

PREVIEW (first 40 rows, tab-separated):
${preview}

USER HINT: "${hint || 'none'}"

TASK: ${modeInstr[mode]}

Return ONLY valid JSON — no markdown, no explanation outside the JSON object:
{
  "skus": [
    { "name": "SKU name", "L": 300, "W": 200, "H": 150, "weight": 2.5, "qty": 1000 }
  ],
  "detected": {
    "nameCol":   "column letter or header name",
    "lCol":      "column letter or header name",
    "wCol":      "column letter or header name",
    "hCol":      "column letter or header name",
    "weightCol": "column letter or header name or null",
    "qtyCol":    "column letter or header name or null",
    "unit":      "mm",
    "headerRow": 1
  },
  "notes":    "What was found and any conversions made",
  "warnings": ["any issues"]
}

DIMENSION RULES — dimensions MUST be in MILLIMETRES (positive integers):
- If values look like cm  (e.g. 30, 20, 15)   → multiply by 10
- If values look like m   (e.g. 0.3, 0.2)      → multiply by 1000
- If values look like mm  (e.g. 300, 200, 150)  → use as-is
- weight = per-box kg (0 if not found)
- qty = total units (1 if not found)
- Skip header rows, totals, blank rows
- name must be a meaningful non-empty string`;

      // ── Call Claude API ───────────────────────────────────────────────────
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-6',
          max_tokens: 2000,
          messages: [{ role: 'user', content: prompt }],
        }),
      });
      if (!resp.ok) throw new Error(`API error ${resp.status}`);

      const data    = await resp.json();
      const rawText = data.content?.[0]?.text || '';

      // Robust JSON extraction (strips accidental markdown fences)
      const match = rawText.match(/\{[\s\S]*\}/);
      if (!match) throw new Error('AI returned unexpected format. Add a hint and try again.');
      const parsed = JSON.parse(match[0]);

      // Validate + clean SKUs
      const validSkus = (parsed.skus || [])
        .filter(s => s.name && s.L > 0 && s.W > 0 && s.H > 0)
        .map(s => ({
          name:   String(s.name).trim(),
          L:      Math.max(1, Math.round(Number(s.L))),
          W:      Math.max(1, Math.round(Number(s.W))),
          H:      Math.max(1, Math.round(Number(s.H))),
          weight: Number(s.weight) || 0,
          qty:    Number(s.qty)    || 1,
        }));

      if (mode !== 'bulk' && validSkus.length === 0)
        throw new Error('No valid SKUs found. Try adding a hint like "dimensions are in columns B, C, D in mm".');

      setResult({ ...parsed, skus: validSkus });

    } catch (err) {
      const msg = err.message || '';
      if (msg.includes('JSON') || msg.includes('parse'))
        setError('AI returned an unexpected format. Add a hint about your column layout and retry.');
      else if (msg.includes('API error 401'))
        setError('API key error. Please check the configuration.');
      else
        setError(msg || 'Failed. Please try again.');
    }
    setLoading(false);
  };

  const apply = () => {
    if (!result || !onFill) return;
    if (mode === 'bulk') {
      // bulk mode: pass column mapping only
      onFill({ columnMap: result.detected, totalRows: result.detected?.totalDataRows });
    } else {
      // multisku / shipment: pass formatted rows
      onFill(result.skus.map((s, i) => ({
        id: i + 1,
        name:      s.name,
        L:         String(s.L),
        W:         String(s.W),
        H:         String(s.H),
        weight:    s.weight > 0 ? String(s.weight) : '',
        qty:       String(s.qty),
        targetQty: String(s.qty),
      })));
    }
    close();
  };

  // ── Button label ────────────────────────────────────────────────────────────
  const btnLabel = label || (mode === 'bulk' ? '🤖 AI Detect Columns' : '🤖 AI Fill from Excel');

  return (
    <>
      {/* Trigger button */}
      <button onClick={() => setOpen(true)}
        style={{ padding: '7px 14px', background: 'linear-gradient(135deg,#6366f1,#4f46e5)',
          color: '#fff', border: 'none', borderRadius: '8px', fontWeight: '700',
          fontSize: '13px', cursor: 'pointer', fontFamily: 'inherit',
          boxShadow: '0 2px 8px rgba(99,102,241,0.35)' }}>
        {btnLabel}
      </button>

      {/* Modal */}
      {open && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)',
          zIndex: 999, display: 'flex', alignItems: 'center', justifyContent: 'center',
          padding: '20px' }} onClick={e => { if (e.target === e.currentTarget) close(); }}>
          <div style={{ background: '#fff', borderRadius: '16px', width: '100%',
            maxWidth: '560px', boxShadow: '0 24px 64px rgba(0,0,0,0.2)', overflow: 'hidden' }}>

            {/* Header */}
            <div style={{ background: 'linear-gradient(135deg,#6366f1,#4f46e5)',
              padding: '20px 24px', display: 'flex', justifyContent: 'space-between',
              alignItems: 'center' }}>
              <div>
                <div style={{ fontWeight: '800', fontSize: '17px', color: '#fff' }}>
                  🤖 AI Fill from Excel
                </div>
                <div style={{ fontSize: '12px', color: 'rgba(255,255,255,0.75)', marginTop: '2px' }}>
                  {mode === 'bulk'
                    ? 'Detects column layout for any file format'
                    : 'Reads any Excel format and auto-fills the form'}
                </div>
              </div>
              <button onClick={close} style={{ background: 'rgba(255,255,255,0.15)',
                border: 'none', color: '#fff', width: '30px', height: '30px',
                borderRadius: '50%', cursor: 'pointer', fontSize: '16px', fontFamily: 'inherit' }}>✕</button>
            </div>

            <div style={{ padding: '24px' }}>
              {/* File upload area */}
              {!result && (
                <>
                  <div
                    onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                    onDragLeave={() => setDragOver(false)}
                    onDrop={e => { e.preventDefault(); setDragOver(false); loadFile(e.dataTransfer.files[0]); }}
                    onClick={() => fileRef.current?.click()}
                    style={{ border: `2px dashed ${dragOver ? '#6366f1' : '#e2e8f0'}`,
                      borderRadius: '10px', padding: '24px', textAlign: 'center',
                      cursor: 'pointer', background: dragOver ? '#eef2ff' : '#fafbfc',
                      transition: 'all 0.2s', marginBottom: '14px' }}>
                    <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv"
                      style={{ display: 'none' }}
                      onChange={e => { loadFile(e.target.files[0]); e.target.value = ''; }}/>
                    <div style={{ fontSize: '28px', marginBottom: '6px' }}>📊</div>
                    <div style={{ fontWeight: '700', color: '#374151', fontSize: '14px' }}>
                      {fileName || 'Drop Excel / CSV here, or click to browse'}
                    </div>
                    {fileName && (
                      <div style={{ fontSize: '12px', color: '#6366f1', marginTop: '4px', fontWeight: '600' }}>
                        ✓ File loaded
                      </div>
                    )}
                  </div>

                  {/* Hint */}
                  <div style={{ marginBottom: '14px' }}>
                    <label style={{ fontSize: '12px', fontWeight: '700', color: '#374151',
                      letterSpacing: '0.06em', textTransform: 'uppercase', display: 'block',
                      marginBottom: '6px' }}>
                      Optional hint (helps AI find the right columns)
                    </label>
                    <input type="text" value={hint} onChange={e => setHint(e.target.value)}
                      placeholder='e.g. "Dimensions in cm" or "Qty is in column G" or "Use Sheet2"'
                      style={{ width: '100%', padding: '9px 12px', border: '1px solid #e2e8f0',
                        borderRadius: '8px', fontSize: '13px', fontFamily: 'inherit', outline: 'none',
                        boxSizing: 'border-box' }}/>
                  </div>

                  {/* Example hints */}
                  <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', marginBottom: '16px' }}>
                    {['Dimensions in cm', 'No header row', 'Qty is in column H', 'Weight in grams'].map(h => (
                      <button key={h} onClick={() => setHint(h)}
                        style={{ padding: '4px 10px', background: '#f1f5f9', border: '1px solid #e2e8f0',
                          borderRadius: '99px', fontSize: '11px', fontWeight: '600', cursor: 'pointer',
                          color: '#475569', fontFamily: 'inherit' }}>
                        + {h}
                      </button>
                    ))}
                  </div>

                  {error && (
                    <div style={{ background: '#fef2f2', border: '1px solid #fecaca',
                      borderRadius: '8px', padding: '10px 14px', color: '#991b1b',
                      fontSize: '13px', marginBottom: '14px' }}>
                      ⚠ {error}
                    </div>
                  )}

                  <button onClick={run} disabled={!file || loading}
                    style={{ width: '100%', padding: '12px', background: !file || loading
                      ? '#e2e8f0' : 'linear-gradient(135deg,#6366f1,#4f46e5)',
                      color: !file || loading ? '#9ca3af' : '#fff',
                      border: 'none', borderRadius: '10px', fontWeight: '700',
                      fontSize: '14px', cursor: !file || loading ? 'not-allowed' : 'pointer',
                      fontFamily: 'inherit', transition: 'all 0.15s' }}>
                    {loading ? '⏳ Analysing with AI...' : '▶ Detect & Fill'}
                  </button>

                  {loading && (
                    <div style={{ textAlign: 'center', marginTop: '12px', fontSize: '12px',
                      color: '#6b7280', lineHeight: '1.6' }}>
                      Reading your file and identifying columns...<br/>
                      <span style={{ color: '#6366f1' }}>Usually takes 3–8 seconds</span>
                    </div>
                  )}
                </>
              )}

              {/* Result preview */}
              {result && (
                <>
                  {/* Detection summary */}
                  <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd',
                    borderRadius: '10px', padding: '14px', marginBottom: '16px' }}>
                    <div style={{ fontWeight: '700', color: '#0369a1', fontSize: '13px',
                      marginBottom: '8px' }}>✓ Detected column layout</div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr',
                      gap: '4px 16px', fontSize: '12px', color: '#374151' }}>
                      {[
                        ['SKU Name', result.detected?.nameCol],
                        ['Length', result.detected?.lCol],
                        ['Width', result.detected?.wCol],
                        ['Height', result.detected?.hCol],
                        ['Weight', result.detected?.weightCol || '—'],
                        ['Quantity', result.detected?.qtyCol || '—'],
                        ['Unit', result.detected?.unit || 'mm'],
                      ].map(([k, v]) => (
                        <div key={k} style={{ display: 'flex', gap: '6px' }}>
                          <span style={{ color: '#9ca3af', minWidth: '58px' }}>{k}:</span>
                          <span style={{ fontWeight: '600', color: '#1d4ed8' }}>{v || '—'}</span>
                        </div>
                      ))}
                    </div>
                    {result.notes && (
                      <div style={{ marginTop: '8px', fontSize: '12px', color: '#0369a1',
                        borderTop: '1px solid #bae6fd', paddingTop: '8px' }}>
                        💡 {result.notes}
                      </div>
                    )}
                  </div>

                  {/* Warnings */}
                  {result.warnings?.length > 0 && (
                    <div style={{ background: '#fffbeb', border: '1px solid #fde68a',
                      borderRadius: '8px', padding: '10px 14px', marginBottom: '14px',
                      fontSize: '12px', color: '#92400e' }}>
                      ⚠ {result.warnings.join(' · ')}
                    </div>
                  )}

                  {/* SKU table preview */}
                  {result.skus.length > 0 && (
                    <div style={{ border: '1px solid #e2e8f0', borderRadius: '8px',
                      overflow: 'hidden', marginBottom: '16px' }}>
                      <div style={{ padding: '8px 12px', background: '#f8fafc',
                        borderBottom: '1px solid #e2e8f0', fontSize: '12px',
                        fontWeight: '700', color: '#374151' }}>
                        {result.skus.length} SKU{result.skus.length > 1 ? 's' : ''} ready to fill
                      </div>
                      <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse',
                          fontSize: '12px' }}>
                          <thead>
                            <tr>{['Name','L','W','H','Wt/box','Qty'].map(h => (
                              <th key={h} style={{ padding: '6px 10px', textAlign: 'left',
                                background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
                                color: '#6b7280', fontSize: '10px', textTransform: 'uppercase',
                                letterSpacing: '0.05em', fontWeight: '700', whiteSpace: 'nowrap' }}>
                                {h}</th>))}</tr>
                          </thead>
                          <tbody>
                            {result.skus.map((s, i) => (
                              <tr key={i} style={{ background: i % 2 === 0 ? '#fff' : '#fafbfc',
                                borderBottom: '1px solid #f1f5f9' }}>
                                <td style={{ padding: '7px 10px', fontWeight: '600',
                                  color: '#111827', maxWidth: '160px', overflow: 'hidden',
                                  textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.name}</td>
                                {[s.L, s.W, s.H].map((v, j) => (
                                  <td key={j} style={{ padding: '7px 10px', color: '#374151',
                                    textAlign: 'right' }}>{v}</td>))}
                                <td style={{ padding: '7px 10px', color: '#6b7280',
                                  textAlign: 'right' }}>{s.weight > 0 ? s.weight : '—'}</td>
                                <td style={{ padding: '7px 10px', fontWeight: '600',
                                  color: '#1d4ed8', textAlign: 'right' }}>
                                  {s.qty.toLocaleString()}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* Bulk mode: column map info only */}
                  {mode === 'bulk' && result.skus.length === 0 && (
                    <div style={{ background: '#f0fdf4', border: '1px solid #bbf7d0',
                      borderRadius: '8px', padding: '12px', marginBottom: '16px',
                      fontSize: '13px', color: '#166534' }}>
                      ✓ Column layout detected. Click Apply — your file will be processed
                      using these column positions.
                    </div>
                  )}

                  {/* Action buttons */}
                  <div style={{ display: 'flex', gap: '10px' }}>
                    <button onClick={reset}
                      style={{ flex: 1, padding: '10px', background: '#f1f5f9',
                        border: '1px solid #e2e8f0', borderRadius: '8px',
                        fontWeight: '600', fontSize: '13px', cursor: 'pointer',
                        color: '#374151', fontFamily: 'inherit' }}>
                      ← Try Again
                    </button>
                    <button onClick={apply}
                      style={{ flex: 2, padding: '10px',
                        background: 'linear-gradient(135deg,#6366f1,#4f46e5)',
                        color: '#fff', border: 'none', borderRadius: '8px',
                        fontWeight: '700', fontSize: '13px', cursor: 'pointer',
                        fontFamily: 'inherit',
                        boxShadow: '0 2px 8px rgba(99,102,241,0.35)' }}>
                      ✓ Apply to Form ({result.skus.length > 0 ? `${result.skus.length} SKUs` : 'column map'})
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </>
  );
}
