import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend } from 'recharts';

const COLORS = {
  DONE: '#10b981',
  'TO DO': '#f59e0b',
  CANCELED: '#ef4444',
  'IN PROGRESS': '#3b82f6',
  'Sin ticket': '#cbd5e1',
};

const renderLabel = ({ percent }) =>
  percent > 0.03 ? `${(percent * 100).toFixed(0)}%` : '';

export default function StatusChart({ data, title = 'Status Jira' }) {
  if (!data || data.length === 0) return null;
  return (
    <div className="chart-card">
      <h3>{title}</h3>
      <ResponsiveContainer width="100%" height={280}>
        <PieChart>
          <Pie
            data={data}
            cx="50%"
            cy="50%"
            outerRadius={100}
            innerRadius={55}
            dataKey="value"
            label={renderLabel}
            labelLine={false}
            isAnimationActive={false}
          >
            {data.map((entry) => (
              <Cell key={entry.name} fill={COLORS[entry.name] || '#94a3b8'} />
            ))}
          </Pie>
          <Tooltip formatter={(v) => v.toLocaleString()} />
          <Legend />
        </PieChart>
      </ResponsiveContainer>
    </div>
  );
}
