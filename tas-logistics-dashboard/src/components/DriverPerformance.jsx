import { useState } from 'react';
import { ChevronUp, ChevronDown } from 'lucide-react';
import { formatMinutes } from '../utils/dataProcessor';

export default function DriverPerformance({ data }) {
  const [sortKey, setSortKey] = useState('total');
  const [sortDir, setSortDir] = useState('desc');
  const [page, setPage] = useState(0);
  const perPage = 15;

  if (!data || data.length === 0) return null;

  const handleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    else {
      setSortKey(key);
      setSortDir('desc');
    }
    setPage(0);
  };

  const sorted = [...data].sort((a, b) => {
    let av = a[sortKey], bv = b[sortKey];
    if (sortKey === 'doneRate' || sortKey === 'ticketRate' || sortKey === 'avgResolution') {
      av = av === '-' || av === null ? -1 : Number(av);
      bv = bv === '-' || bv === null ? -1 : Number(bv);
    }
    if (typeof av === 'string') return sortDir === 'asc' ? av.localeCompare(bv) : bv.localeCompare(av);
    return sortDir === 'asc' ? av - bv : bv - av;
  });

  const paged = sorted.slice(page * perPage, (page + 1) * perPage);
  const totalPages = Math.ceil(sorted.length / perPage);

  const SortIcon = ({ col }) => {
    if (sortKey !== col) return null;
    return sortDir === 'asc' ? <ChevronUp size={14} /> : <ChevronDown size={14} />;
  };

  const columns = [
    { key: 'agent', label: 'Agente' },
    { key: 'total', label: 'Compras' },
    { key: 'pagada', label: 'Pagadas' },
    { key: 'abierta', label: 'Abiertas' },
    { key: 'withTicket', label: 'Con Ticket' },
    { key: 'done', label: 'DONE' },
    { key: 'todo', label: 'TO DO' },
    { key: 'doneRate', label: 'DONE %' },
    { key: 'avgResolution', label: 'Tiempo Resol.' },
  ];

  return (
    <div className="chart-card chart-full">
      <h3>Performance por Agente ({data.length} agentes)</h3>
      <div className="table-wrapper">
        <table className="data-table">
          <thead>
            <tr>
              {columns.map((col) => (
                <th key={col.key} onClick={() => handleSort(col.key)}>
                  <span>{col.label} <SortIcon col={col.key} /></span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {paged.map((row, i) => (
              <tr key={i}>
                <td>{row.agent}</td>
                <td>{row.total}</td>
                <td>{row.pagada}</td>
                <td className={row.abierta > 0 ? 'text-amber' : ''}>{row.abierta}</td>
                <td>{row.withTicket}</td>
                <td>{row.done}</td>
                <td>{row.todo}</td>
                <td>
                  {row.doneRate !== '-' ? (
                    <span className={`badge ${Number(row.doneRate) >= 80 ? 'badge-green' : Number(row.doneRate) >= 50 ? 'badge-yellow' : 'badge-red'}`}>
                      {row.doneRate}%
                    </span>
                  ) : '-'}
                </td>
                <td>{formatMinutes(row.avgResolution)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {totalPages > 1 && (
        <div className="pagination">
          <button disabled={page === 0} onClick={() => setPage(page - 1)}>Anterior</button>
          <span>Pag {page + 1} de {totalPages}</span>
          <button disabled={page >= totalPages - 1} onClick={() => setPage(page + 1)}>Siguiente</button>
        </div>
      )}
    </div>
  );
}
