import { AlertTriangle, Clock, CheckCircle } from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from 'recharts';

const COLORS = {
  withinSLA: '#10b981',
  outsideSLA: '#ef4444',
};

export default function SLADistribution({ data }) {
  if (!data) return null;
  const { buckets, withinSLA, totalMeasured, slaRate, hasPublishData } = data;

  const chartData = buckets.map(b => ({
    name: b.label,
    count: b.count,
    withinSLA: b.max <= 10080,
  }));

  const slaNum = slaRate ? parseFloat(slaRate) : 0;
  const slaColor = slaNum >= 100 ? '#10b981' : slaNum >= 80 ? '#f59e0b' : '#ef4444';

  const title = hasPublishData
    ? 'Distribucion SLA: Compra a Publicado'
    : 'Distribucion SLA: Compra a Resolucion';

  return (
    <div className="chart-card chart-wide">
      <h3>
        <Clock size={16} style={{ marginRight: 6 }} />
        {title}
      </h3>

      {hasPublishData ? (
        <div className="sla-notice" style={{ background: '#f0fdf4', borderColor: '#86efac' }}>
          <CheckCircle size={14} style={{ color: '#16a34a' }} />
          <span style={{ color: '#166534' }}>
            Usando fecha real de publicacion (publication_created_date_local). Target: 100% en &lt; 7 dias naturales.
          </span>
        </div>
      ) : (
        <div className="sla-notice">
          <AlertTriangle size={14} />
          <span>
            Nota: Se usa PO &rarr; Resolucion Jira como proxy. Carga el CSV de Publish para ver datos reales.
          </span>
        </div>
      )}

      <div className="sla-summary">
        <div className="sla-metric">
          <span className="sla-metric-value" style={{ color: slaColor }}>
            {slaRate ?? '-'}%
          </span>
          <span className="sla-metric-label">Dentro del SLA (&lt; 7 dias)</span>
        </div>
        <div className="sla-metric">
          <span className="sla-metric-value">{withinSLA.toLocaleString()}</span>
          <span className="sla-metric-label">Dentro</span>
        </div>
        <div className="sla-metric">
          <span className="sla-metric-value" style={{ color: '#ef4444' }}>
            {(totalMeasured - withinSLA).toLocaleString()}
          </span>
          <span className="sla-metric-label">Fuera</span>
        </div>
        <div className="sla-metric">
          <span className="sla-metric-value">{totalMeasured.toLocaleString()}</span>
          <span className="sla-metric-label">Total medido</span>
        </div>
      </div>

      {totalMeasured > 0 ? (
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={chartData} margin={{ top: 10, right: 20, left: 10, bottom: 5 }}>
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis tick={{ fontSize: 12 }} />
            <Tooltip
              formatter={(v) => [v, 'Registros']}
              contentStyle={{ fontSize: 12, borderRadius: 8 }}
            />
            <Bar dataKey="count" radius={[4, 4, 0, 0]}>
              {chartData.map((entry, i) => (
                <Cell
                  key={i}
                  fill={entry.withinSLA ? COLORS.withinSLA : COLORS.outsideSLA}
                  fillOpacity={0.85}
                />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      ) : (
        <p className="no-data">Sin datos suficientes para calcular SLA</p>
      )}

      <div className="sla-legend">
        <span className="sla-legend-item">
          <span className="sla-legend-dot" style={{ background: COLORS.withinSLA }} />
          Dentro del SLA (&le; 7 dias)
        </span>
        <span className="sla-legend-item">
          <span className="sla-legend-dot" style={{ background: COLORS.outsideSLA }} />
          Fuera del SLA (&gt; 7 dias)
        </span>
      </div>
    </div>
  );
}
