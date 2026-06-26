// ─── PASTE FROM EXCEL ─────────────────────────────────────────────────────────
// User copies cells from Excel → pastes here → rows fill instantly.
// Expected column order shown above paste area.
import { useState } from 'react';

export default function PasteFromExcel({ mode = 'multisku', onFill }) {
  const [open, setOpen]     = useState(false);
  const [text, setText]     = useState('');
  const [preview, setPreview] = useState([]);
  const [error, setError]   = useState('');

  const configs = {
    multisku: {
      cols: ['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight/Box (kg)', 'Target Qty'],
      hint: 'Arrange your Excel columns in this order, select the cells, copy (Ctrl+C), then paste below.',
    },
    shipment: {
      cols: ['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight/Box (kg)', 'Qty to Ship'],
      hint: 'Arrange your Excel columns in this order, select the cells, copy (Ctrl+C), then paste below.',
    },
    bulk: {
      cols: ['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight/Box (kg)', 'Available Qty'],
      hint: 'Arrange your Excel columns in this order, select the cells, copy (Ctrl+C), then paste below.',
    },
  };

  const cfg = configs[mode] || configs.multisku;

  const parse = (raw) => {
    setText(raw);
    setError('');
    setPreview([]);
    if (!raw.trim()) return;

    const rows = raw.trim().split('\n')
      .map(r => r.split('\t').map(c => c.trim()))
      .filter(r => r.some(c => c));

    if (!rows.length) { setError('No data found. Make sure you copied cells from Excel.'); return; }

    // Skip header row if first cell looks like text (not a number or SKU code)
    let dataRows = rows;
    const firstCell = rows[0][0]?.toLowerCase() || '';
    if (['sku','name','product','material','item','description','article'].some(k => firstCell.includes(k))) {
      dataRows = rows.slice(1);
    }

    if (!dataRows.length) { setError('Only a header row found. Select data rows too.'); return; }

    const parsed = dataRows.map((r, i) => ({
      id: i + 1,
      name:      r[0] || `SKU ${i + 1}`,
      L:         r[1] || '',
      W:         r[2] || '',
      H:         r[3] || '',
      weight:    r[4] || '',
      qty:       r[5] || '',
      targetQty: r[5] || '',
    })).filter(r => r.name && (r.L || r.W || r.H));

    if (!parsed.length) {
      setError('Could not read rows. Check column order: Name | L | W | H | Weight | Qty');
      return;
    }

    if (parsed.length > 8 && (mode === 'multisku' || mode === 'shipment')) {
      setError(`Found ${parsed.length} rows — only first 8 will be used (tool limit).`);
      setPreview(parsed.slice(0, 8));
    } else {
      setPreview(parsed);
    }
  };

  const apply = () => {
    if (!preview.length) return;
    onFill(preview.slice(0, mode === 'bulk' ? 9999 : 8));
    setOpen(false);
    setText('');
    setPreview([]);
    setError('');
  };

  return (
    <>
      <button onClick={() => setOpen(true)}
        style={{ padding: '7px 14px', background: '#f0f9ff', border: '1px solid #bae6fd',
          borderRadius: '8px', fontSize: '13px', fontWeight: '600', color: '#0369a1',
          cursor: 'pointer', fontFamily: 'inherit' }}>
        📋 Paste from Excel
      </button>

      {open && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
          zIndex: 999, display: 'flex', alignItems: 'center', justifyContent: 'center',
          padding: '20px' }} onClick={e => { if (e.target === e.currentTarget) setOpen(false); }}>
          <div style={{ background: '#fff', borderRadius: '16px', width: '100%',
            maxWidth: '580px', boxShadow: '0 24px 64px rgba(0,0,0,0.2)', overflow: 'hidden' }}>

            {/* Header */}
            <div style={{ background: '#0f172a', padding: '18px 24px',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <div style={{ fontWeight: '800', fontSize: '16px', color: '#fff' }}>
                  📋 Paste from Excel
                </div>
                <div style={{ fontSize: '12px', color: '#94a3b8', marginTop: '2px' }}>
                  Copy cells from Excel → paste below → done
                </div>
              </div>
              <button onClick={() => setOpen(false)}
                style={{ background: 'rgba(255,255,255,0.1)', border: 'none', color: '#fff',
                  width: '28px', height: '28px', borderRadius: '50%', cursor: 'pointer',
                  fontSize: '14px', fontFamily: 'inherit' }}>✕</button>
            </div>

            <div style={{ padding: '20px 24px' }}>
              {/* Column order hint */}
              <div style={{ marginBottom: '14px' }}>
                <div style={{ fontSize: '12px', fontWeight: '700', color: '#374151',
                  textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: '8px' }}>
                  Expected column order
                </div>
                <div style={{ display: 'flex', gap: '4px', flexWrap: 'wrap' }}>
                  {cfg.cols.map((col, i) => (
                    <span key={i} style={{ background: '#f1f5f9', border: '1px solid #e2e8f0',
                      borderRadius: '6px', padding: '4px 10px', fontSize: '12px',
                      fontWeight: '600', color: '#475569', display: 'flex',
                      alignItems: 'center', gap: '5px' }}>
                      <span style={{ background: '#be185d', color: '#fff', borderRadius: '50%',
                        width: '16px', height: '16px', display: 'inline-flex',
                        alignItems: 'center', justifyContent: 'center',
                        fontSize: '9px', fontWeight: '800', flexShrink: 0 }}>{i + 1}</span>
                      {col}
                    </span>
                  ))}
                </div>
                <div style={{ fontSize: '12px', color: '#6b7280', marginTop: '8px' }}>
                  {cfg.hint}
                </div>
              </div>

              {/* Paste area */}
              <textarea
                value={text}
                onChange={e => parse(e.target.value)}
                onPaste={e => {
                  // Small delay to let paste complete
                  setTimeout(() => parse(e.target.value), 10);
                }}
                placeholder={'Paste your Excel data here (Ctrl+V)\n\nExample:\nSKU-001\t300\t200\t150\t2.5\t1000\nSKU-002\t450\t320\t200\t4.0\t500'}
                style={{ width: '100%', height: '130px', border: '1px solid #e2e8f0',
                  borderRadius: '8px', padding: '10px 12px', fontSize: '12px',
                  fontFamily: 'monospace', resize: 'vertical', outline: 'none',
                  boxSizing: 'border-box', color: '#374151', lineHeight: '1.6' }}
                autoFocus
              />

              {error && (
                <div style={{ background: '#fff1f2', border: '1px solid #fecaca',
                  borderRadius: '8px', padding: '8px 12px', fontSize: '12px',
                  color: '#991b1b', marginTop: '8px' }}>
                  ⚠ {error}
                </div>
              )}

              {/* Preview table */}
              {preview.length > 0 && (
                <div style={{ marginTop: '12px', border: '1px solid #e2e8f0',
                  borderRadius: '8px', overflow: 'hidden' }}>
                  <div style={{ padding: '7px 12px', background: '#f0fdf4',
                    fontSize: '12px', fontWeight: '700', color: '#166534',
                    borderBottom: '1px solid #e2e8f0' }}>
                    ✓ {preview.length} row{preview.length > 1 ? 's' : ''} ready to fill
                  </div>
                  <div style={{ overflowX: 'auto', maxHeight: '150px', overflowY: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px' }}>
                      <thead>
                        <tr>{['Name', 'L', 'W', 'H', 'Wt', 'Qty'].map(h => (
                          <th key={h} style={{ padding: '5px 8px', background: '#f8fafc',
                            borderBottom: '1px solid #e2e8f0', textAlign: 'left',
                            color: '#6b7280', fontWeight: '700', whiteSpace: 'nowrap' }}>{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {preview.map((r, i) => (
                          <tr key={i} style={{ background: i % 2 ? '#fafbfc' : '#fff' }}>
                            <td style={{ padding: '5px 8px', fontWeight: '600', maxWidth: '120px',
                              overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.name}</td>
                            {[r.L, r.W, r.H, r.weight, r.qty].map((v, j) => (
                              <td key={j} style={{ padding: '5px 8px', color: v ? '#374151' : '#d1d5db',
                                textAlign: 'right' }}>{v || '—'}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Actions */}
              <div style={{ display: 'flex', gap: '10px', marginTop: '14px' }}>
                <button onClick={() => { setText(''); setPreview([]); setError(''); }}
                  style={{ flex: 1, padding: '9px', background: '#f1f5f9',
                    border: '1px solid #e2e8f0', borderRadius: '8px', fontWeight: '600',
                    fontSize: '13px', cursor: 'pointer', color: '#374151',
                    fontFamily: 'inherit' }}>
                  Clear
                </button>
                <button onClick={apply} disabled={!preview.length}
                  style={{ flex: 2, padding: '9px',
                    background: preview.length ? '#be185d' : '#e2e8f0',
                    color: preview.length ? '#fff' : '#9ca3af',
                    border: 'none', borderRadius: '8px', fontWeight: '700',
                    fontSize: '13px', cursor: preview.length ? 'pointer' : 'not-allowed',
                    fontFamily: 'inherit' }}>
                  ✓ Fill {preview.length > 0 ? `${preview.length} SKU${preview.length > 1 ? 's' : ''}` : 'Form'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
