import { Package, Clock, AlertTriangle, ShieldCheck } from 'lucide-react';

const cards = [
  { key: 'total', label: 'PENDIENTES', icon: Package, format: 'number', hero: true },
  { key: 'avgWaitDays', label: 'PROM. ESPERA', icon: Clock, format: 'days' },
  { key: 'overSLA', label: 'FUERA DE SLA', icon: AlertTriangle, format: 'number', danger: true },
  { key: 'slaCompliance', label: 'CUMPLIMIENTO SLA', icon: ShieldCheck, format: 'percent' },
];

function getValueColor(key, val) {
  if (key === 'avgWaitDays') {
    return val <= 3 ? '#10b981' : val <= 7 ? '#f59e0b' : '#ef4444';
  }
  if (key === 'overSLA') {
    return val > 0 ? '#ef4444' : '#10b981';
  }
  if (key === 'slaCompliance') {
    return val >= 100 ? '#10b981' : val >= 80 ? '#f59e0b' : '#ef4444';
  }
  return '#0f172a';
}

function formatValue(val, fmt) {
  if (fmt === 'percent') return `${val}%`;
  if (fmt === 'days') return `${val}d`;
  if (typeof val === 'number') return val.toLocaleString();
  return val;
}

export default function KPICards({ pendingKPIs }) {
  if (!pendingKPIs) return null;

  return (
    <div className="kpi-grid-v2">
      {cards.map(({ key, label, icon: Icon, format: fmt, hero, danger }) => {
        const val = pendingKPIs[key];
        const color = getValueColor(key, val);
        return (
          <div key={key} className={`kpi-card-v2${hero ? ' kpi-hero' : ''}`}>
            <div className="kpi-card-header">
              <span className="kpi-label-v2">
                <Icon size={14} style={{ marginRight: 4, opacity: 0.7 }} />
                {label}
              </span>
            </div>
            <span
              className="kpi-value-v2"
              style={hero ? {} : { color }}
            >
              {formatValue(val, fmt)}
            </span>
            {key === 'total' && (
              <span className="kpi-desc-v2">
                {pendingKPIs.pagada} pagados · {pendingKPIs.withoutTicket} sin ticket
              </span>
            )}
            {key === 'overSLA' && val > 0 && (
              <span className="kpi-desc-v2" style={{ color: '#ef4444' }}>
                Max: {pendingKPIs.maxWaitDays} dias
              </span>
            )}
          </div>
        );
      })}
    </div>
  );
}
