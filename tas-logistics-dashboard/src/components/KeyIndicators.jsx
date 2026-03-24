import { AlertTriangle, AlertCircle, CheckCircle } from 'lucide-react';

const statusIcons = {
  red: AlertCircle,
  yellow: AlertTriangle,
  green: CheckCircle,
};

const badgeLabels = {
  red: 'Critico',
  yellow: 'Atencion',
  green: 'OK',
};

export default function KeyIndicators({ alerts }) {
  if (!alerts || alerts.length === 0) {
    return (
      <div className="chart-card">
        <h3>
          <CheckCircle size={16} style={{ marginRight: 6, color: '#10b981' }} />
          Alertas Operativas
        </h3>
        <div className="indicators-list">
          <div className="indicator-card indicator-green">
            <div className="indicator-header">
              <span className="indicator-badge badge-green">OK</span>
              <strong>Sin alertas</strong>
            </div>
            <p className="indicator-desc">Todos los autos pendientes estan dentro del SLA de 7 dias.</p>
          </div>
        </div>
      </div>
    );
  }

  const criticalCount = alerts.filter((a) => a.severity === 'red' || a.severity === 'yellow').length;

  return (
    <div className="chart-card">
      <h3>
        <AlertTriangle size={16} style={{ marginRight: 6 }} />
        Alertas Operativas ({criticalCount > 0 ? criticalCount : alerts.length})
      </h3>
      <div className="indicators-list">
        {alerts.map((alert, i) => {
          const Icon = statusIcons[alert.severity];
          return (
            <div key={i} className={`indicator-card indicator-${alert.severity}`}>
              <div className="indicator-header">
                <span className={`indicator-badge badge-${alert.severity}`}>
                  {badgeLabels[alert.severity]}
                </span>
                <strong>{alert.title}</strong>
              </div>
              <p className="indicator-desc">{alert.description}</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}
