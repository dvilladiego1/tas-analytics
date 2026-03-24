import { formatMinutes } from '../utils/dataProcessor';
import { Clock } from 'lucide-react';

const metrics = [
  { key: 'poToJiraCreated', label: 'PO → Ticket Jira', desc: 'Tiempo entre la orden de compra y la creacion del ticket' },
  { key: 'createdToResolved', label: 'Ticket → Resolucion', desc: 'Tiempo entre creacion del ticket y resolucion' },
];

export default function TimeAnalysis({ data }) {
  if (!data) return null;

  return (
    <div className="chart-card chart-full">
      <h3>Analisis de Tiempos</h3>
      <div className="time-grid">
        {metrics.map(({ key, label, desc }) => {
          const m = data[key];
          if (!m) return null;
          return (
            <div key={key} className="time-card">
              <div className="time-header">
                <Clock size={16} />
                <span>{label}</span>
              </div>
              <div className="time-body">
                <div className="time-metric">
                  <span className="time-value">{formatMinutes(m.avg)}</span>
                  <span className="time-label">Promedio</span>
                </div>
                <div className="time-metric">
                  <span className="time-value">{formatMinutes(m.median)}</span>
                  <span className="time-label">Mediana</span>
                </div>
                <div className="time-metric">
                  <span className="time-value">{m.count.toLocaleString()}</span>
                  <span className="time-label">Registros</span>
                </div>
              </div>
              <p className="time-desc">{desc}</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}
