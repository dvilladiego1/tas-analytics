import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts';

const PALETTE = ['#6366f1', '#8b5cf6', '#a78bfa', '#c4b5fd', '#818cf8', '#6d28d9', '#7c3aed', '#5b21b6', '#4f46e5', '#4338ca'];

export default function BrandChart({ data }) {
  if (!data || data.length === 0) {
    return (
      <div className="chart-card">
        <h3>Marcas de Auto</h3>
        <p className="no-data">Sin datos de marca</p>
      </div>
    );
  }

  const top = data.slice(0, 15);

  return (
    <div className="chart-card">
      <h3>Top Marcas de Auto</h3>
      <ResponsiveContainer width="100%" height={300}>
        <BarChart data={top} layout="vertical" margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
          <XAxis type="number" tick={{ fontSize: 11 }} />
          <YAxis dataKey="name" type="category" width={110} tick={{ fontSize: 11 }} />
          <Tooltip />
          <Bar dataKey="value" fill="#6366f1" radius={[0, 4, 4, 0]} name="Compras" />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
