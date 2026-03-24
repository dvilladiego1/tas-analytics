import { useState } from 'react';
import { ChevronUp, ChevronDown, MapPin } from 'lucide-react';

export default function RoutesBreakdown({ data }) {
  const [sortKey, setSortKey] = useState('pending');
  const [sortDir, setSortDir] = useState('desc');
  const [page, setPage] = useState(0);
  const perPage = 15;

  if (!data || data.length === 0) {
    return (
      <div className="chart-card chart-wide">
        <h3><MapPin size={16} style={{ marginRight: 6 }} />Rutas con Inventario Pendiente</h3>
        <p className="no-data">Sin autos pendientes por ruta</p>
      </div>
    );
  }

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    else { setSortKey(key); setSortDir('desc'); }
    setPage(0);
  };

  const sorted = [...data].sort((a, b) => {
    let aVal = a[sortKey];
    let bVal = b[sortKey];
    if (aVal == null) aVal = -1;
    if (bVal == null) bVal = -1;
    if (typeof aVal === 'string') return sortDir === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
    return sortDir === 'asc' ? aVal - bVal : bVal - aVal;
  });

  const totalPages = Math.ceil(sorted.length / perPage);
  const pageData = sorted.slice(page * perPage, (page + 1) * perPage);
  const SortIcon = sortDir === 'asc' ? ChevronUp : ChevronDown;

  const cols = [
    { key: 'route', label: 'Ruta' },
    { key: 'pending', label: 'Pendientes' },
    { key: 'avgWaitDays', label: 'Prom. Espera' },
    { key: 'maxWaitDays', label: 'Max Espera' },
    { key: 'overSLA', label: 'Fuera SLA' },
    { key: 'slaCompliance', label: 'SLA %' },
    { key: 'withoutTicket', label: 'Sin Ticket' },
    { key: 'avgHistoricalDays', label: 'Hist. Prom.' },
  ];

  return (
    <div className="chart-card chart-wide">
      <h3><MapPin size={16} style={{ marginRight: 6 }} />Rutas con Inventario Pendiente ({data.length} rutas)</h3>
      <div className="table-wrapper">
        <table className="data-table">
          <thead>
            <tr>
              {cols.map(({ key, label }) => (
                <th key={key} onClick={() => toggleSort(key)}>
                  <span>
                    {label}
                    {sortKey === key && <SortIcon size={14} />}
                  </span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageData.map((r, i) => (
              <tr key={i}>
                <td style={{ maxWidth: 320, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  <span className="route-hub-dot" style={{ background: r.hubColor }} />
                  {r.route}
                </td>
                <td style={{ fontWeight: 700 }}>{r.pending}</td>
                <td style={{ color: r.avgWaitDays > 7 ? '#ef4444' : r.avgWaitDays > 3 ? '#f59e0b' : '#10b981', fontWeight: 600 }}>
                  {r.avgWaitDays}d
                </td>
                <td style={{ color: r.maxWaitDays > 7 ? '#ef4444' : '#64748b', fontWeight: 600 }}>
                  {r.maxWaitDays}d
                </td>
                <td style={{ color: r.overSLA > 0 ? '#ef4444' : '#64748b', fontWeight: 600 }}>
                  {r.overSLA}
                </td>
                <td>
                  <span className={`badge ${r.slaCompliance >= 100 ? 'badge-green' : r.slaCompliance >= 80 ? 'badge-yellow' : 'badge-red'}`}>
                    {r.slaCompliance}%
                  </span>
                </td>
                <td style={{ color: r.withoutTicket > 0 ? '#f59e0b' : '#64748b', fontWeight: 600 }}>
                  {r.withoutTicket}
                </td>
                <td style={{ color: '#64748b' }}>
                  {r.avgHistoricalDays != null ? `${r.avgHistoricalDays}d` : '-'}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {totalPages > 1 && (
        <div className="pagination">
          <button disabled={page === 0} onClick={() => setPage((p) => p - 1)}>Anterior</button>
          <span>{page + 1} / {totalPages}</span>
          <button disabled={page >= totalPages - 1} onClick={() => setPage((p) => p + 1)}>Siguiente</button>
        </div>
      )}
    </div>
  );
}
