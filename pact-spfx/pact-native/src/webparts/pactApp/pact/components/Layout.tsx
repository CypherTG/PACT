import React from 'react';
import { NavLink, useHistory } from 'react-router-dom';
import { LayoutDashboard, Users, AlertTriangle, FileText, Settings, FileSearch, ShieldAlert, Mail } from 'lucide-react';
// import './Layout.css';
import { sharePointService } from '../services/SharePointService';
import { WorkbenchBridge } from './Runtime/WorkbenchBridge';

interface LayoutProps {
  children?: React.ReactNode;
}

const navigation = [
  { name: 'Dashboard', href: '/', icon: LayoutDashboard },
  { name: 'Compliance Cases', href: '/cases', icon: ShieldAlert },
  { name: 'Staff Directory', href: '/staff', icon: Users },
  { name: 'Escalation Log', href: '/escalations', icon: AlertTriangle },
  { name: 'Policy Library', href: '/policies', icon: FileText },
  { name: 'Communication Log', href: '/admin/mail-log', icon: Mail },
  { name: 'Settings', href: '/settings', icon: Settings },
];

export const Layout: React.FC<LayoutProps> = ({ children }) => {
  const history = useHistory();
  const [session, setSession] = React.useState(sharePointService.getCurrentSession());

  React.useEffect(() => {
    sharePointService.initialize();
    // Refresh session after initialization
    setSession(sharePointService.getCurrentSession());
  }, []);

  const runtimeStyles = sharePointService.isStandalone()
    ? { borderColor: 'rgba(245, 158, 11, 0.35)', color: 'var(--status-warning)', background: 'rgba(245, 158, 11, 0.12)' }
    : { borderColor: 'rgba(16, 185, 129, 0.35)', color: 'var(--status-success)', background: 'rgba(16, 185, 129, 0.12)' };
  const runtimeConnectionLabel = sharePointService.isStandalone()
    ? 'Workbench Demo'
    : 'SharePoint Native';

  return (
    <div className="app-layout" data-testid="pact-app-shell">
      {/* Sidebar Navigation */}
      <aside className="sidebar glass-panel">
        <div className="sidebar-header">
          <div className="logo-container">
            <div className="logo-icon" style={{ overflow: 'hidden', background: 'transparent', boxShadow: 'none', border: 'none' }}>
              <img 
                src={require('../assets/kcc-logo.png')} 
                alt="KCC" 
                style={{ width: '100%', height: '100%', objectFit: 'contain' }}
                onError={(e) => {
                  (e.target as HTMLImageElement).style.display = 'none';
                  const parent = (e.target as HTMLImageElement).parentElement;
                  if (parent) {
                    parent.innerText = 'P';
                    parent.style.background = 'linear-gradient(135deg, var(--primary), #ff6b35)';
                    parent.style.display = 'flex';
                    parent.style.alignItems = 'center';
                    parent.style.justifyContent = 'center';
                    parent.style.color = 'white';
                    parent.style.fontWeight = 'bold';
                  }
                }}
              />
            </div>
            <h1 className="logo-text">PACT</h1>
          </div>
        </div>

        <nav className="sidebar-nav scroll-container">
          <ul>
            {navigation.map((item) => (
              <li key={item.name}>
                <NavLink
                  exact={item.href === '/'}
                  to={item.href}
                  className="nav-link"
                  activeClassName="active"
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
            <div className="avatar">
              {session.photoUrl ? (
                <img src={session.photoUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
              ) : (
                session.displayName.charAt(0)
              )}
            </div>
            <div className="user-info">
              <span className="user-name">{session.displayName}</span>
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
              <div className="badge-avatar">
                {session.photoUrl ? (
                  <img src={session.photoUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                ) : (
                  session.displayName.charAt(0)
                )}
              </div>
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <span className="badge-name">{session.displayName}</span>
                <span className="badge-email" style={{ fontSize: '10px', opacity: 0.6 }}>{session.email}</span>
              </div>
            </div>

            <button className="btn btn-primary log-incident-btn" onClick={() => history.push('/cases/new')} data-testid="log-incident-button">
              <ShieldAlert size={16} />
              Log Incident
            </button>
          </div>
        </header>

        <div className="page-container" data-testid="pact-page-container">
          {children}
        </div>
      </main>
      <WorkbenchBridge />
    </div>
  );
};
