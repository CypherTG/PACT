import React from 'react';
import { NavLink, useHistory } from 'react-router-dom';
import { LayoutDashboard, Users, AlertTriangle, FileText, Settings, FileSearch, ShieldAlert, Mail } from 'lucide-react';
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
  { name: 'Appeals', href: '/appeals', icon: FileSearch },
  { name: 'Policy Library', href: '/policies', icon: FileText },
  { name: 'Communication Log', href: '/admin/mail-log', icon: Mail },
  { name: 'Settings', href: '/settings', icon: Settings },
];

export const Layout = ({ children }: LayoutProps): React.ReactElement => {
  const history = useHistory();
  const [session, setSession] = React.useState(() => sharePointService.getCurrentSession());
  
  React.useEffect(() => {
    if (sharePointService) {
      sharePointService.initialize().then(() => {
        setSession(sharePointService.getCurrentSession());
      }).catch(() => {
        setSession(sharePointService.getCurrentSession());
      });
    }
  }, []);

  const username = session.displayName || "User";
  const userEmail = session.email || "No email available";
  const tenantLabel = session.tenantName || session.siteTitle || 'SharePoint';
  const runtimeLabel = sharePointService.getRuntimeLabel();
  const runtimeStyles = sharePointService.isStandalone()
    ? { borderColor: 'rgba(245, 158, 11, 0.35)', color: 'var(--status-warning)', background: 'rgba(245, 158, 11, 0.12)' }
    : { borderColor: 'rgba(16, 185, 129, 0.35)', color: 'var(--status-success)', background: 'rgba(16, 185, 129, 0.12)' };
  const handlePhotoError = (event: React.SyntheticEvent<HTMLImageElement>): void => {
    event.currentTarget.style.display = 'none';
    const fallback = event.currentTarget.parentElement?.querySelector('.avatar-fallback, .badge-avatar-fallback') as HTMLElement | null;
    if (fallback) {
      fallback.style.display = 'flex';
    }
  };

  const location = history.location.pathname;
  const getPageTitle = (): string => {
    if (location === '/') return 'Executive Overview';
    if (location.startsWith('/cases')) return 'Compliance Cases';
    if (location.startsWith('/staff')) return 'Staff Directory';
    if (location.startsWith('/escalations')) return 'Escalation Log';
    if (location.startsWith('/appeals')) return 'Appeals Management';
    if (location.startsWith('/policies')) return 'Policy Library';
    if (location.startsWith('/admin')) return 'System Administration';
    if (location.startsWith('/settings')) return 'System Settings';
    return 'Governance Platform';
  };

  return (
    <div className="app-layout" data-testid="pact-app-shell">
      {/* Sidebar Navigation */}
      <aside className="sidebar glass-panel">
        <div className="sidebar-header">
          <div className="logo-container">
            <img 
              src={require('../assets/kcc-logo.png')} 
              alt="KCC" 
              className="logo-img" 
              style={{ width: '40px', height: '40px', borderRadius: '10px', objectFit: 'contain' }} 
              onError={(e) => { 
                (e.target as HTMLImageElement).style.display = 'none'; 
                const el = document.getElementById('logo-fallback'); 
                if (el) el.style.display = 'flex'; 
              }} 
            />
            <div id="logo-fallback" className="logo-icon" style={{ display: 'none' }}>P</div>
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
              <div className="avatar avatar-photo">
                <img
                  src={session.photoUrl}
                  alt={username}
                  onError={handlePhotoError}
                  onLoad={(event) => {
                    event.currentTarget.style.display = 'block';
                  }}
                />
                <span className="avatar-fallback">{username.charAt(0) || 'U'}</span>
              </div>
              <div className="user-info">
                <span className="user-name">{username}</span>
                <span className="user-email">{userEmail}</span>
                <span className="user-role" style={{ color: 'var(--status-success)', display: 'flex', alignItems: 'center' }}>
                  <span className="status-dot"></span>
                  {tenantLabel}
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
              {getPageTitle()}
              <span style={{ fontSize: '10px', opacity: 0.3, marginLeft: '10px', fontWeight: 'normal' }}>v5.0-OPTIMIZED</span>
            </h2>
            <div className="header-subtitle">
              {tenantLabel} — Governance Dashboard
            </div>
          </div>

          
          <div className="header-actions">
            <div className="header-user-badge">
              <div className="badge-avatar badge-photo">
                <img
                  src={session.photoUrl}
                  alt={username}
                  onError={handlePhotoError}
                  onLoad={(event) => {
                    event.currentTarget.style.display = 'block';
                  }}
                />
                <span className="badge-avatar-fallback">{username.charAt(0) || 'U'}</span>
              </div>
              <div className="header-user-meta">
                <span className="badge-name">{username}</span>
                <span className="badge-email">{userEmail}</span>
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
