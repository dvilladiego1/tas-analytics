import { Clock } from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, ReferenceLine } from 'recharts';

export default function AgingDistribution({ data }) {
  if (!data || data.total === 0) return null;

  const chartData = data.buckets.map((b) => ({
    name: b.label,
    count: b.count,
    color: b.color,
  }));

  return (
    <div className="chart-card chart-wide">
      <h3>
        <Clock size={16} style={{ marginRight: 6 }} />
        Distribucion de Antiguedad — Autos Pendientes
      </h3>
      <p className="chart-subtitle">
        Dias desde compra (PO) sin llegar a hub. Target: &lt; 7 dias naturales.
      </p>

      <div className="aging-summary">
        {data.buckets.map((b) => (
          <div key={b.label} className="aging-summary-item">
            <span className="aging-dot" style={{ background: b.color }} />
            <span className="aging-label">{b.label}</span>
            <span className="aging-count" style={{ color: b.color }}>{b.count}</span>
          </div>
        ))}
      </div>

      <ResponsiveContainer width="100%" height={280}>
        <BarChart data={chartData} margin={{ top: 10, right: 20, left: 10, bottom: 5 }}>
          <XAxis dataKey="name" tick={{ fontSize: 12 }} />
          <YAxis tick={{ fontSize: 12 }} />
          <Tooltip
            formatter={(v) => [v, 'Autos']}
            contentStyle={{ fontSize: 12, borderRadius: 8 }}
          />
          <ReferenceLine
            x="7-14 dias"
            stroke="#ef4444"
            strokeDasharray="4 4"
            strokeWidth={2}
            label={{ value: 'SLA 7d', position: 'top', fontSize: 11, fill: '#ef4444' }}
          />
          <Bar dataKey="count" radius={[6, 6, 0, 0]}>
            {chartData.map((entry, i) => (
              <Cell key={i} fill={entry.color} fillOpacity={0.9} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
