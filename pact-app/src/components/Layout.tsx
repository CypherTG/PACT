import React from 'react';
import { NavLink, Outlet, useNavigate } from 'react-router-dom';
import { LayoutDashboard, Users, AlertTriangle, FileText, Settings, ShieldAlert, Mail } from 'lucide-react';
import './Layout.css';
import { sharePointService } from '../services/SharePointService';
import { WorkbenchBridge } from './Runtime/WorkbenchBridge';

const navigation = [
  { name: 'Dashboard', href: '/', icon: LayoutDashboard },
  { name: 'Compliance Cases', href: '/cases', icon: ShieldAlert },
  { name: 'Staff Directory', href: '/staff', icon: Users },
  { name: 'Escalation Log', href: '/escalations', icon: AlertTriangle },
  { name: 'Policy Library', href: '/policies', icon: FileText },
  { name: 'Communication Log', href: '/admin/mail-log', icon: Mail },
  { name: 'Settings', href: '/settings', icon: Settings },
];

export const Layout: React.FC = () => {
  const navigate = useNavigate();

  React.useEffect(() => {
    sharePointService.initialize();
  }, []);

  const username = sharePointService.getUserName();
  const runtimeConnectionLabel = sharePointService.isStandalone()
    ? 'Workbench Demo'
    : 'SharePoint Native';
  const runtimeStyles = sharePointService.isStandalone()
    ? { borderColor: 'rgba(245, 158, 11, 0.35)', color: 'var(--status-warning)', background: 'rgba(245, 158, 11, 0.12)' }
    : { borderColor: 'rgba(16, 185, 129, 0.35)', color: 'var(--status-success)', background: 'rgba(16, 185, 129, 0.12)' };

  return (
    <div className="app-layout" data-testid="pact-app-shell">
      {/* Sidebar Navigation */}
      <aside className="sidebar glass-panel">
        <div className="sidebar-header">
          <div className="logo-container">
            <img src="./kcc-logo.png" alt="KCC" className="logo-img" style={{ width: '40px', height: '40px', borderRadius: '10px', objectFit: 'contain' }} onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; const el = document.getElementById('logo-fallback'); if (el) el.style.display = 'flex'; }} />
            <div id="logo-fallback" className="logo-icon" style={{ display: 'none' }}>P</div>
            <h1 className="logo-text">PACT</h1>
          </div>
        </div>

        <nav className="sidebar-nav scroll-container">
          <ul>
            {navigation.map((item) => (
              <li key={item.name}>
                <NavLink
                  to={item.href}
                  className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
                  data-testid={`nav-${item.href.replace(/[^\w]/g, '').toLowerCase() || 'home'}`}
                >
                  <item.icon className="nav-icon" size={20} />
                  <span>{item.name}</span>
                </NavLink>
              </li>
            ))}
          </ul>
        </nav>

        <div className="sidebar-footer">
          <div className="user-profile">
            <div className="avatar" style={{ background: 'var(--status-success)', color: 'white' }}>
              {username.charAt(0)}
            </div>
            <div className="user-info">
              <span className="user-name">{username}</span>
              <span className="user-role" style={{ color: runtimeStyles.color, display: 'flex', alignItems: 'center' }}>
                <span className="status-dot"></span>
                {runtimeConnectionLabel}
              </span>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <main className="main-content scroll-container" data-testid="pact-main">
        <header className="top-header glass-panel">
          <div className="header-left">
            <h2 className="page-title">
              Executive Overview
              <span style={{ fontSize: '10px', opacity: 0.3, marginLeft: '10px', fontWeight: 'normal' }}>v5.0-NATIVE</span>
            </h2>
          </div>

          <div className="header-actions">

            <div className="header-user-badge">
              <div className="badge-avatar">{username.charAt(0)}</div>
              <span className="badge-name">{username}</span>
            </div>

            <button className="btn btn-primary log-incident-btn" onClick={() => navigate('/cases/new')} data-testid="log-incident-button">
              <ShieldAlert size={16} />
              Log Incident
            </button>
          </div>
        </header>

        <div className="page-container" data-testid="pact-page-container">
          <Outlet />
        </div>
      </main>
      <WorkbenchBridge />
    </div>
  );
};
