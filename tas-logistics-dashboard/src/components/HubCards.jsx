export default function HubCards({ hubs }) {
  if (!hubs || hubs.length === 0) return null;

  return (
    <div className="hub-grid">
      {hubs.map((hub) => {
        const slaColor = hub.slaCompliance >= 100 ? '#10b981'
          : hub.slaCompliance >= 80 ? '#f59e0b'
          : '#ef4444';
        const waitColor = hub.avgWaitDays <= 3 ? '#10b981'
          : hub.avgWaitDays <= 7 ? '#f59e0b'
          : '#ef4444';

        return (
          <div key={hub.id} className="hub-card" style={{ borderLeftColor: hub.color }}>
            <div className="hub-card-header">
              <span className="hub-card-label">{hub.label}</span>
              <span className="hub-card-badge" style={{ color: slaColor }}>
                {hub.slaCompliance}% SLA
              </span>
            </div>
            <span className="hub-card-value">{hub.pending}</span>
            <span className="hub-card-subtitle">autos pendientes</span>
            <div className="hub-card-meta">
              <span style={{ color: waitColor }}>Prom: {hub.avgWaitDays}d</span>
              <span style={{ color: hub.overSLA > 0 ? '#ef4444' : '#64748b' }}>
                Fuera SLA: {hub.overSLA}
              </span>
              <span>Con ticket: {hub.withTicket}</span>
            </div>
            <div className="hub-sla-bar">
              <div
                className="hub-sla-fill"
                style={{ width: `${Math.min(hub.slaCompliance, 100)}%`, background: slaColor }}
              />
            </div>
          </div>
        );
      })}
    </div>
  );
}
