import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from 'recharts';

export default function ProviderBreakdown({ data, title = 'Por Region Hub' }) {
  if (!data || data.length === 0) return null;

  const top = data.slice(0, 15);

  return (
    <div className="chart-card">
      <h3>{title}</h3>
      <ResponsiveContainer width="100%" height={300}>
        <BarChart data={top} layout="vertical" margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
          <XAxis type="number" tick={{ fontSize: 11 }} />
          <YAxis dataKey="name" type="category" width={140} tick={{ fontSize: 11 }} />
          <Tooltip />
          <Legend />
          <Bar dataKey="DONE" stackId="a" fill="#10b981" />
          <Bar dataKey="TO DO" stackId="a" fill="#f59e0b" />
          <Bar dataKey="CANCELED" stackId="a" fill="#ef4444" />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
