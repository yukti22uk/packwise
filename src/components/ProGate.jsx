// ─── PRO GATE ─────────────────────────────────────────────────────────────────
import { S } from './styles.jsx';

const TOOL_DETAILS = {
  'Multi-SKU Planner': {
    icon: '🗃️',
    tagline: 'Pack multiple box sizes in one container',
    bullets: [
      'Pack 2–8 different SKU sizes together in a single container',
      'AI-optimised region allocation per SKU with colour-coded 3D view',
      'Respects weight limits, fragile, this-side-up & stacking constraints',
      'Export per-container manifest and WhatsApp share loading plan',
    ],
  },
  'Shipment Planner': {
    icon: '🚚',
    tagline: 'Plan your entire shipment across multiple trucks',
    bullets: [
      'Enter total order quantity → get exact containers needed',
      'Supports single SKU and multi-SKU mixed shipments',
      'Per-container manifest for every truck in your shipment',
      'Cost comparison across all vehicle types — find cheapest option',
      'Export branded PDF loading plan for your warehouse team',
    ],
  },
  'Order Analyser': {
    icon: '📊',
    tagline: 'Turn your order data into actionable logistics insights',
    bullets: [
      'Paste Master SKU data and Order data directly from Excel',
      'Flags anomalies — missing dimensions, zero quantities, duplicate SKUs',
      'ABC analysis by shipping volume (A = top 70% of your freight)',
      'FMS classification by order frequency (Fast / Medium / Slow movers)',
      'ABC-FMS matrix shows which SKUs to prioritise for container planning',
      'Download 6-sheet Excel report ready to share with your supply chain team',
    ],
  },
};

export default function ProGate({ onUpgrade, feature }) {
  const details = TOOL_DETAILS[feature] || {
    icon: '🔒',
    tagline: 'Unlock the full DensiCube suite',
    bullets: [
      'Multi-SKU Planner — pack 2–8 box sizes together',
      'Shipment Planner — plan full shipments across multiple trucks',
      'Order Analyser — ABC, FMS and anomaly reports from your order data',
    ],
  };

  return (
    <div style={{ ...S.card, padding: '40px 32px',
      background: 'linear-gradient(135deg,#f0fdf4 0%,#eff6ff 100%)',
      border: '1px solid #bbf7d0', textAlign: 'center' }}>

      {/* Icon + title */}
      <div style={{ fontSize: '48px', marginBottom: '10px' }}>{details.icon}</div>
      <div style={{ fontSize: '22px', fontWeight: '800', color: '#0f172a', marginBottom: '6px' }}>
        {feature}
      </div>
      <div style={{ fontSize: '15px', fontWeight: '600', color: '#059669',
        marginBottom: '20px' }}>
        {details.tagline}
      </div>

      {/* Feature bullets */}
      <div style={{ background: '#fff', borderRadius: '12px', padding: '20px 24px',
        maxWidth: '480px', margin: '0 auto 24px', textAlign: 'left',
        border: '1px solid #e2e8f0' }}>
        <div style={{ fontSize: '11px', fontWeight: '700', color: '#9ca3af',
          textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '12px' }}>
          What you get with Pro
        </div>
        {details.bullets.map((b, i) => (
          <div key={i} style={{ display: 'flex', gap: '10px', alignItems: 'flex-start',
            marginBottom: i < details.bullets.length - 1 ? '10px' : 0 }}>
            <span style={{ color: '#059669', fontWeight: '700', flexShrink: 0,
              marginTop: '1px' }}>✓</span>
            <span style={{ fontSize: '13px', color: '#374151', lineHeight: '1.5' }}>{b}</span>
          </div>
        ))}
      </div>

      {/* Price + CTA */}
      <div style={{ fontSize: '13px', color: '#6b7280', marginBottom: '14px' }}>
        All Pro tools included — <strong style={{ color: '#0f172a' }}>₹999 / month</strong>
      </div>
      <button onClick={onUpgrade}
        style={{ padding: '13px 32px', background: 'linear-gradient(135deg,#059669,#047857)',
          color: '#fff', border: 'none', borderRadius: '10px', fontSize: '15px',
          fontWeight: '700', cursor: 'pointer', fontFamily: 'inherit',
          boxShadow: '0 4px 14px rgba(5,150,105,0.35)' }}>
        ⭐ Upgrade to Pro — ₹999/month
      </button>
      <div style={{ fontSize: '12px', color: '#9ca3af', marginTop: '10px' }}>
        Instant access · Cancel anytime · All Pro tools unlocked
      </div>
    </div>
  );
}
