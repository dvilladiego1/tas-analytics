import {
  LayoutDashboard, Users, BarChart3, UserCog,
  ClipboardCheck, Bot, FileBarChart, Settings, Truck
} from 'lucide-react';

const navItems = [
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { id: 'grupos', label: 'Grupos', icon: Users },
  { id: 'benchmark', label: 'Benchmark', icon: BarChart3 },
  { id: 'usuarios', label: 'Gestion de Usuarios', icon: UserCog },
  { id: 'checkin', label: 'Check-in', icon: ClipboardCheck },
  { id: 'agentes', label: 'Agentes IA', icon: Bot },
  { id: 'reportes', label: 'Reportes', icon: FileBarChart },
  { id: 'config', label: 'Configuracion', icon: Settings },
];

export default function Sidebar({ activeItem = 'dashboard' }) {
  return (
    <aside className="sidebar">
      <div className="sidebar-brand">
        <div className="sidebar-logo">
          <Truck size={22} />
        </div>
        <div className="sidebar-brand-text">
          <span className="sidebar-title">TAS Manager</span>
          <span className="sidebar-subtitle">Kavak B2B</span>
        </div>
      </div>

      <nav className="sidebar-nav">
        {navItems.map(({ id, label, icon: Icon }) => (
          <button
            key={id}
            className={`sidebar-nav-item ${activeItem === id ? 'active' : ''}`}
            title={id !== 'dashboard' ? 'Proximamente' : ''}
          >
            <Icon size={20} />
            <span>{label}</span>
          </button>
        ))}
      </nav>

      <div className="sidebar-footer">
        <div className="sidebar-user">
          <div className="sidebar-avatar">DV</div>
          <div className="sidebar-user-info">
            <span className="sidebar-user-name">Daniel Villadiego</span>
            <span className="sidebar-user-email">daniel@kavak.com</span>
          </div>
        </div>
      </div>
    </aside>
  );
}
