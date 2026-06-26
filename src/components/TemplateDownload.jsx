// ─── EXCEL TEMPLATE DOWNLOAD ──────────────────────────────────────────────────
// Replaces AI Fill — gives users a pre-formatted Excel template to fill in.
import * as XLSX from 'xlsx';

export default function TemplateDownload({ mode = 'multisku' }) {
  const templates = {
    multisku: {
      label: '⬇ Download SKU Template',
      filename: 'PackWise_MultiSKU_Template.xlsx',
      headers: ['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight per Box (kg)', 'Target Qty'],
      examples: [
        ['Product A', 300, 200, 150, 2.5, 1000],
        ['Product B', 450, 320, 200, 4.0, 500],
        ['Product C', 250, 180, 120, 1.8, 2000],
      ],
      note: 'Fill in your SKU data. All dimensions must be in mm. Upload back here once done.',
    },
    shipment: {
      label: '⬇ Download Shipment Template',
      filename: 'PackWise_Shipment_Template.xlsx',
      headers: ['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight per Box (kg)', 'Qty to Ship'],
      examples: [
        ['Ceramic Tile Box', 400, 300, 200, 6.0, 5000],
        ['Packaging Box A', 300, 250, 180, 3.5, 2500],
        ['Export Carton B', 500, 400, 250, 8.0, 1200],
      ],
      note: 'Enter total quantity to ship per SKU. Upload this file to the Order Analyser for full analytics.',
    },
    bulk: {
      label: '⬇ Download Bulk SKU Template',
      filename: 'PackWise_BulkSKU_Template.xlsx',
      headers: ['SKU Name', 'Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight per Box (kg)', 'Available Qty'],
      examples: [
        ['SKU-001', 300, 200, 150, 2.5, 1000],
        ['SKU-002', 450, 320, 200, 4.0, 500],
        ['SKU-003', 250, 180, 120, 1.8, 2000],
        ['SKU-004', 600, 400, 300, 9.0, 750],
        ['SKU-005', 350, 280, 180, 3.2, 3000],
      ],
      note: 'Add as many SKUs as needed. Column order must be: Name, L, W, H, Weight, Qty.',
    },
  };

  const t = templates[mode] || templates.multisku;

  const download = () => {
    const wb = XLSX.utils.book_new();
    const data = [
      ['PackWise — ' + t.filename.replace('PackWise_','').replace('_Template.xlsx','').replace(/_/g,' ')],
      [''],
      ['Instructions: Fill in your data below. Do not change column headers. All dimensions in mm.'],
      [''],
      t.headers,
      ...t.examples,
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 20 }, { wch: 14 }];
    // Style header row (row 5, index 4)
    XLSX.utils.book_append_sheet(wb, ws, 'SKU Data');
    XLSX.writeFile(wb, t.filename);
  };

  return (
    <button onClick={download}
      style={{ padding: '7px 14px', background: '#f0fdf4', border: '1px solid #bbf7d0',
        borderRadius: '8px', fontSize: '13px', fontWeight: '600', color: '#166534',
        cursor: 'pointer', fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: '6px' }}>
      {t.label}
    </button>
  );
}
