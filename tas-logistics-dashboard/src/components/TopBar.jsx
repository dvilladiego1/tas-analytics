import { Download, LayoutGrid, List } from 'lucide-react';

export default function TopBar({ title, subtitle, children, bbddCount, jiraCount, recordCount }) {
  return (
    <div className="top-bar">
      <div className="top-bar-left">
        <div className="top-bar-title-group">
          <h1 className="top-bar-title">{title}</h1>
          {subtitle && <p className="top-bar-subtitle">{subtitle}</p>}
        </div>
        <div className="top-bar-meta">
          <span className="top-bar-status">
            <span className="status-dot" />
            Datos cargados
          </span>
          {bbddCount != null && <span className="top-bar-badge">BBDD: {bbddCount.toLocaleString()}</span>}
          {jiraCount != null && <span className="top-bar-badge">Jira: {jiraCount.toLocaleString()}</span>}
          {recordCount != null && <span className="top-bar-badge accent">{recordCount.toLocaleString()} registros</span>}
        </div>
      </div>
      <div className="top-bar-right">
        {children}
      </div>
    </div>
  );
}
