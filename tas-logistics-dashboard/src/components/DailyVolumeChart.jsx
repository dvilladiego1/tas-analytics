import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts';
import { format, parseISO } from 'date-fns';
import { TrendingUp, TrendingDown, Minus } from 'lucide-react';

export default function DailyVolumeChart({ data }) {
  if (!data || data.length === 0) return null;

  const formatTick = (val) => {
    try { return format(parseISO(val), 'dd/MM'); } catch { return val; }
  };

  // Compute trend: compare last 7 days avg vs prior 7 days avg
  const last7 = data.slice(-7);
  const prior7 = data.slice(-14, -7);
  const avg = (arr) => arr.length > 0 ? arr.reduce((a, b) => a + b.pending, 0) / arr.length : 0;
  const lastAvg = avg(last7);
  const priorAvg = avg(prior7);
  const delta = lastAvg - priorAvg;
  const trendText = delta > 1 ? 'Creciendo' : delta < -1 ? 'Disminuyendo' : 'Estable';
  const TrendIcon = delta > 1 ? TrendingUp : delta < -1 ? TrendingDown : Minus;
  const trendColor = delta < -1 ? '#10b981' : delta > 1 ? '#ef4444' : '#64748b';

  return (
    <div className="chart-card chart-wide">
      <div className="chart-header-row">
        <h3>Tendencia Inventario Pendiente (30 dias)</h3>
        <span className="trend-badge" style={{ color: trendColor }}>
          <TrendIcon size={14} />
          {trendText}
        </span>
      </div>
      <ResponsiveContainer width="100%" height={280}>
        <AreaChart data={data} margin={{ top: 5, right: 20, left: 0, bottom: 5 }}>
          <defs>
            <linearGradient id="pendingGradient" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.3} />
              <stop offset="95%" stopColor="#3b82f6" stopOpacity={0.05} />
            </linearGradient>
          </defs>
          <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
          <XAxis
            dataKey="date"
            tickFormatter={formatTick}
            tick={{ fontSize: 11 }}
            interval={Math.max(0, Math.floor(data.length / 10))}
          />
          <YAxis tick={{ fontSize: 11 }} />
          <Tooltip
            labelFormatter={(v) => {
              try { return format(parseISO(v), 'dd/MM/yyyy'); } catch { return v; }
            }}
            formatter={(v) => [v, 'Pendientes']}
          />
          <Area
            type="monotone"
            dataKey="pending"
            stroke="#3b82f6"
            strokeWidth={2}
            fill="url(#pendingGradient)"
          />
        </AreaChart>
      </ResponsiveContainer>
    </div>
  );
}
