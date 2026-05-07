const STYLE_ID = 'pact-global-styles';

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Outfit:wght@300;400;500;600;700&display=swap');

:root {
  --bg-base: #f9fafb;
  --bg-surface: #ffffff;
  --bg-surface-hover: #f3f4f6;
  --bg-card: rgba(255, 255, 255, 0.95);
  --primary: #dc2626;
  --primary-hover: #ef4444;
  --secondary: #6b7280;
  --text-primary: var(--bodyText, #111827);
  --text-secondary: var(--bodyText, #4b5563);
  --text-muted: #9ca3af;
  --border-light: rgba(0, 0, 0, 0.1);
  --border-focus: rgba(220, 38, 38, 0.5);
  --status-info: #3b82f6;
  --status-success: #10b981;
  --status-warning: #f59e0b;
  --status-danger: #dc2626;
  --radii-md: 8px;
  --radii-lg: 16px;
  --radii-xl: 24px;
  --shadow-sm: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
  --shadow-md: 0 10px 15px -3px rgba(0, 0, 0, 0.2), 0 4px 6px -2px rgba(0, 0, 0, 0.1);
  --shadow-glow: 0 0 15px rgba(233, 69, 96, 0.3);
  --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
  --font-display: 'Outfit', var(--font-sans);
}

/* Theme Awareness Overrides */
.pact-workbench {
  --bg-panel: rgba(255, 255, 255, 0.95);
}

[data-is-dark-theme="true"] .app-layout {
  --text-primary: #f8fafc;
  --text-secondary: #cbd5e1;
  --bg-surface: #1e293b;
  --border-light: rgba(255, 255, 255, 0.1);
}

* { box-sizing: border-box; }

body {
  font-family: var(--font-sans);
  background-color: var(--bg-base);
  color: var(--text-primary);
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
  min-height: 100vh;
  margin: 0;
  padding: 0;
}

h1, h2, h3, h4, h5, h6 {
  font-family: var(--font-display);
  font-weight: 600;
  letter-spacing: -0.02em;
  margin-top: 0;
  margin-bottom: 1rem;
}

a {
  color: var(--primary);
  text-decoration: none;
  transition: color 0.2s ease;
}

a:hover { color: var(--primary-hover); }

.glass-panel {
  background: var(--bg-card);
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  border: 1px solid var(--border-light);
  border-radius: var(--radii-lg);
  box-shadow: var(--shadow-md);
}

.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 0.5rem 1.2rem;
  border-radius: var(--radii-md);
  font-weight: 500;
  font-size: 0.95rem;
  cursor: pointer;
  transition: all 0.2s ease;
  border: none;
  gap: 0.5rem;
}

.btn-primary {
  background: linear-gradient(135deg, var(--primary), #d2334e);
  color: white;
  box-shadow: var(--shadow-glow);
}

.btn-primary:hover {
  background: linear-gradient(135deg, var(--primary-hover), var(--primary));
  transform: translateY(-1px);
  box-shadow: 0 0 20px rgba(233, 69, 96, 0.5);
}

.btn-secondary {
  background: white;
  color: var(--text-primary);
  border: 1px solid var(--border-light);
  box-shadow: 0 1px 2px rgba(0,0,0,0.05);
}

.btn-secondary:hover { background: var(--bg-surface-hover); }

.fade-in { animation: fadeIn 0.4s ease-out forwards; }
.skeleton {
  background: linear-gradient(90deg, #f0f0f0 25%, #e0e0e0 50%, #f0f0f0 75%);
  background-size: 936px 100%;
  animation: shimmer 2s infinite linear;
  border-radius: 4px;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}

@keyframes shimmer {
  0% { background-position: -468px 0; }
  100% { background-position: 468px 0; }
}

.pact-table {
  width: 100%;
  border-collapse: collapse;
  text-align: left;
}

.pact-table th {
  padding: 1rem;
  font-weight: 600;
  font-size: 0.85rem;
  color: var(--text-muted);
  text-transform: uppercase;
  letter-spacing: 0.05em;
  border-bottom: 2px solid var(--border-light);
}

.pact-table td {
  padding: 1.25rem 1rem;
  border-bottom: 1px solid var(--border-light);
  color: var(--text-secondary);
}

.pact-table tr:hover td { background: rgba(0, 0, 0, 0.02); }

.status-badge {
  padding: 4px 10px;
  border-radius: 100px;
  font-size: 0.75rem;
  font-weight: 600;
}

.status-paid { background: rgba(16, 185, 129, 0.1); color: #10b981; }
.status-unpaid { background: rgba(220, 38, 38, 0.1); color: #dc2626; }
.status-overdue { background: rgba(245, 158, 11, 0.1); color: #f59e0b; }

.cases-container { display: flex; flex-direction: column; gap: 24px; }
.cases-header { display: flex; justify-content: space-between; align-items: center; gap: 16px; }
.search-bar { display: flex; align-items: center; gap: 12px; padding: 10px 16px; flex: 1; max-width: 400px; }
.search-icon { color: var(--text-muted); }
.search-bar input { background: transparent; border: none; color: var(--text-primary); font-family: inherit; font-size: 0.95rem; width: 100%; outline: none; }
.search-bar input::placeholder { color: var(--text-muted); }
.cases-actions { display: flex; gap: 12px; }
.cases-table-container { overflow-x: auto; border-radius: var(--radii-md); }
.font-medium { font-weight: 500; }
.text-white { color: var(--text-primary); }
.person-cell { display: flex; align-items: center; gap: 10px; color: var(--text-primary); }
.person-avatar {
  width: 28px;
  height: 28px;
  background: var(--bg-surface-hover);
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 0.8rem;
  font-weight: 600;
  color: var(--secondary);
}
.status-waived { background: rgba(157, 163, 180, 0.2); color: var(--text-secondary); }
.link-action { font-weight: 500; color: var(--secondary); }
.link-action:hover { text-decoration: underline; }
.loading-state { padding: 40px; text-align: center; color: var(--text-muted); }

.dashboard-container { display: flex; flex-direction: column; gap: 24px; }
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; }
@media (min-width: 1024px) { .kpi-grid { grid-template-columns: repeat(4, 1fr); } }
.kpi-card { display: flex; align-items: center; padding: 20px; gap: 16px; transition: transform 0.2s ease, box-shadow 0.2s ease; }
.kpi-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-glow); }
.kpi-icon { width: 54px; height: 54px; border-radius: 14px; display: flex; align-items: center; justify-content: center; color: white; }
.kpi-icon.info { background: linear-gradient(135deg, rgba(0, 120, 212, 0.8), rgba(0, 120, 212, 0.4)); }
.kpi-icon.warning { background: linear-gradient(135deg, rgba(245, 158, 11, 0.8), rgba(245, 158, 11, 0.4)); }
.kpi-icon.danger { background: linear-gradient(135deg, rgba(239, 68, 68, 0.8), rgba(239, 68, 68, 0.4)); }
.kpi-icon.primary { background: linear-gradient(135deg, rgba(233, 69, 96, 0.8), rgba(233, 69, 96, 0.4)); }
.kpi-content { display: flex; flex-direction: column; }
.kpi-label { font-size: 0.85rem; color: var(--text-secondary); font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em; }
.kpi-value { font-family: var(--font-display); font-size: 2rem; font-weight: 700; color: var(--text-primary); line-height: 1.2; }
.dashboard-charts-grid { display: grid; grid-template-columns: 3fr 2fr; gap: 24px; }
@media (max-width: 1024px) { .dashboard-charts-grid { grid-template-columns: 1fr; } }
.chart-panel, .activity-panel { padding: 24px; display: flex; flex-direction: column; }
.panel-header { margin-bottom: 20px; }
.panel-header h3 { display: flex; align-items: center; gap: 8px; margin: 0; font-size: 1.1rem; color: var(--text-primary); }
.chart-wrapper { height: 300px; width: 100%; }
.dashboard-lists-grid { display: grid; grid-template-columns: 1fr; gap: 24px; }
.activity-list { display: flex; flex-direction: column; gap: 16px; }
.activity-item { display: flex; gap: 16px; padding: 12px; border-radius: var(--radii-md); background: rgba(0, 0, 0, 0.02); border: 1px solid var(--border-light); transition: background 0.2s ease; }
.activity-item:hover { background: rgba(0, 0, 0, 0.05); }
.activity-indicator { width: 4px; border-radius: 2px; flex-shrink: 0; }
.activity-indicator.severity-low { background-color: var(--status-info); }
.activity-indicator.severity-medium { background-color: var(--status-warning); }
.activity-indicator.severity-high { background-color: var(--primary); }
.activity-indicator.severity-critical { background-color: var(--status-danger); }
.activity-content { display: flex; flex-direction: column; gap: 4px; }
.activity-title { font-weight: 600; font-size: 0.95rem; color: var(--text-primary); }
.activity-desc { font-size: 0.85rem; color: var(--text-secondary); }
.activity-time { font-size: 0.75rem; color: var(--text-muted); margin-top: 4px; }

.app-layout {
  display: flex;
  min-height: 100vh;
  width: 100%;
  overflow: hidden;
  background: radial-gradient(circle at top left, rgba(233, 69, 96, 0.15) 0%, transparent 40%),
              radial-gradient(circle at bottom right, rgba(0, 120, 212, 0.1) 0%, transparent 40%),
              var(--bg-base);
}
.sidebar { width: 280px; min-height: calc(100vh - 32px); margin: 16px 0 16px 16px; display: flex; flex-direction: column; transition: all 0.3s ease; z-index: 10; }
.sidebar-header { padding: 24px 24px 16px 24px; border-bottom: 1px solid var(--border-light); }
.logo-container { display: flex; align-items: center; gap: 12px; }
.logo-icon { width: 36px; height: 36px; background: linear-gradient(135deg, var(--primary), #ff6b35); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-family: var(--font-display); font-weight: 700; font-size: 1.2rem; color: white; box-shadow: var(--shadow-glow); }
.logo-text {
  font-family: var(--font-display);
  font-size: 1.5rem;
  font-weight: 800;
  letter-spacing: -0.025em;
  background: linear-gradient(135deg, var(--primary), #ef4444);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  margin: 0;
}
.logo-sub { font-weight: 300; color: var(--text-secondary); }
.sidebar-subtitle { font-size: 0.8rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin-top: 8px; }

.sidebar-nav { flex: 1; padding: 20px 12px; overflow-y: auto; }
.sidebar-nav ul { list-style: none; padding: 0; display: flex; flex-direction: column; gap: 6px; }

.form-group {
  margin-bottom: 24px;
}

.form-input {
  width: 100%;
  padding: 12px 16px;
  border: 1px solid var(--border-light);
  border-radius: var(--radii-md);
  background: rgba(255, 255, 255, 0.05);
  color: var(--text-primary);
  font-family: var(--font-sans);
  font-size: 0.95rem;
  transition: all 0.2s ease;
  display: block;
}

.form-input:focus {
  outline: none;
  border-color: var(--primary);
  box-shadow: 0 0 0 3px var(--border-focus);
}

.form-label {
  display: block;
  font-size: 0.875rem;
  font-weight: 600;
  color: var(--text-secondary);
  margin-bottom: 8px;
}

.nav-link { 
  display: flex; 
  align-items: center; 
  gap: 12px; 
  padding: 12px 16px; 
  border-radius: var(--radii-md); 
  color: var(--text-primary); 
  font-weight: 600; 
  text-decoration: none;
  transition: all 0.2s ease; 
}
.nav-link:hover { background: rgba(0, 0, 0, 0.05); color: var(--primary); }
.nav-link.active { background: linear-gradient(90deg, rgba(220, 38, 38, 0.1), transparent); color: var(--primary); border-left: 3px solid var(--primary); }
.nav-icon { opacity: 0.8; }

.sidebar-footer { padding: 20px; border-top: 1px solid var(--border-light); }
.user-profile { display: flex; align-items: center; gap: 12px; }
.user-info { display: flex; flex-direction: column; }
.user-name { font-weight: 700; font-size: 0.95rem; color: var(--text-primary); }
.user-role { font-size: 0.75rem; color: var(--text-secondary); display: flex; align-items: center; gap: 6px; }

/* Dashboard Header Text Visibility */
.app-header h1, .app-header h2 {
  color: var(--text-primary) !important;
  font-weight: 800;
}
.header-welcome h1 {
  color: var(--text-primary);
}
.header-welcome p {
  color: var(--text-secondary);
}
.avatar { width: 36px; height: 36px; background: var(--bg-surface-hover); border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 600; color: var(--primary); border: 1px solid var(--border-light); overflow: hidden; position: relative; }
.avatar-photo img { width: 100%; height: 100%; object-fit: cover; display: block; }
.avatar-fallback { position: absolute; inset: 0; display: none; align-items: center; justify-content: center; background: var(--status-success); color: white; font-weight: 700; }
.user-name { font-size: 0.9rem; font-weight: 600; color: var(--text-primary); }
.user-email { font-size: 0.78rem; color: var(--text-muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.user-role { font-size: 0.75rem; color: var(--text-secondary); }
.auth-btn { display: flex; align-items: center; gap: 4px; background: none; border: none; font-size: 0.75rem; font-weight: 600; cursor: pointer; padding: 4px 8px; border-radius: 4px; margin-top: 4px; transition: all 0.2s ease; width: fit-content; }
.auth-btn.login { background: rgba(0, 120, 212, 0.1); color: var(--primary); border: 1px solid rgba(0, 120, 212, 0.2); }
.auth-btn.login:hover { background: var(--primary); color: white; }
.auth-btn.logout { background: rgba(239, 68, 68, 0.05); color: var(--status-danger); border: 1px solid rgba(239, 68, 68, 0.1); }
.auth-btn.logout:hover { background: var(--status-danger); color: white; }
.sidebar {
  width: var(--sidebar-width);
  background: #ffffff;
  border-right: 1px solid var(--border-light);
  display: flex;
  flex-direction: column;
  z-index: 50;
  box-shadow: 2px 0 10px rgba(0,0,0,0.02);
}

.app-header {
  height: var(--header-height);
  background: rgba(255, 255, 255, 0.8);
  backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--border-light);
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 32px;
  position: sticky;
  top: 0;
  z-index: 40;
}
.top-header { display: flex; align-items: center; justify-content: space-between; padding: 16px 24px; margin-bottom: 24px; position: sticky; top: 0; z-index: 5; background: #ffffff; border-bottom: 1px solid var(--border-light); }
.header-welcome h1 {
  font-size: 1.5rem;
  font-weight: 800;
  color: #111827 !important;
  margin: 0;
}
.header-welcome p {
  font-size: 0.875rem;
  color: #4b5563 !important;
}
.page-title { font-size: 1.5rem; font-weight: 800; margin: 0; color: #111827 !important; }
.header-left { display: flex; flex-direction: column; gap: 4px; }
.header-subtitle { font-size: 0.76rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.08em; }
.header-actions { display: flex; align-items: center; gap: 12px; }
.login-header-btn { background: var(--primary); color: white; border: none; box-shadow: var(--shadow-glow); transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
.login-header-btn:hover { transform: translateY(-1px); filter: brightness(1.1); box-shadow: 0 4px 12px rgba(233, 69, 96, 0.4); }
.header-user-badge { display: flex; align-items: center; gap: 10px; padding: 8px 12px; background: rgba(255, 255, 255, 0.05); border: 1px solid var(--border-light); border-radius: 20px; margin-right: 8px; }
.badge-avatar { width: 28px; height: 28px; border-radius: 50%; background: var(--status-success); display: flex; align-items: center; justify-content: center; font-size: 0.75rem; font-weight: 700; color: white; overflow: hidden; position: relative; flex-shrink: 0; }
.badge-photo img { width: 100%; height: 100%; object-fit: cover; display: block; }
.badge-avatar-fallback { position: absolute; inset: 0; display: none; align-items: center; justify-content: center; }
.header-user-meta { display: flex; flex-direction: column; min-width: 0; }
.badge-name { font-size: 0.85rem; font-weight: 600; color: var(--text-primary); line-height: 1.1; }
.badge-email { font-size: 0.72rem; color: var(--text-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 220px; }
.page-container { flex: 1; max-width: 1400px; margin: 0 auto; width: 100%; }

body.pact-workbench .app-layout {
  width: 100% !important;
}

body.pact-workbench .main-content {
  min-width: 0;
}

body.pact-workbench .page-container {
  width: 100%;
}


body.pact-workbench .top-header {
  flex-wrap: nowrap;
}

body.pact-workbench .header-user-badge {
  min-width: 360px;
}

body.pact-workbench .badge-avatar {
  width: 36px;
  height: 36px;
}

body.pact-workbench .badge-email {
  max-width: 280px;
}

body.pact-workbench #workbenchPageContent,
body.pact-workbench .CanvasZone,
body.pact-workbench .CanvasSection,
body.pact-workbench .CanvasComponent,
body.pact-workbench .CanvasControl,
body.pact-workbench [data-automation-id='CanvasZone'],
body.pact-workbench [data-automation-id='CanvasControl'],
body.pact-workbench .CanvasZoneContainer,
body.pact-workbench .mainContent {
  min-width: 100% !important;
  width: 100% !important;
  max-width: none !important;
  padding: 0 !important;
  margin: 0 !important;
}

body.pact-workbench .ControlZone {
  padding: 0 !important;
  margin: 0 !important;
  max-width: none !important;
}

body.pact-workbench #pact-workbench-overlay {
  min-width: 100% !important;
  width: 100% !important;
  max-width: none !important;
  overflow: auto !important;
}

html,
body,
#root {
  min-height: 100%;
}

#root {
  min-height: 100vh;
}

body {
  overscroll-behavior: none;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

.header-runtime-badge {
  display: inline-flex;
  align-items: center;
  padding: 6px 12px;
  border: 1px solid;
  border-radius: 999px;
  font-size: 0.78rem;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  white-space: nowrap;
}

.qa-fab {
  position: fixed;
  right: 18px;
  bottom: 18px;
  z-index: 60;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 10px 14px;
  border-radius: 999px;
  font-size: 0.8rem;
  font-weight: 700;
  cursor: pointer;
  background: rgba(17, 24, 39, 0.92);
  color: white;
  border: 1px solid rgba(255, 255, 255, 0.12);
}

.qa-panel {
  position: fixed;
  right: 18px;
  bottom: 18px;
  z-index: 60;
  width: min(360px, calc(100vw - 36px));
  padding: 16px;
  background: rgba(255, 255, 255, 0.96);
}

.qa-panel-header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 14px;
}

.qa-kicker {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 0.72rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--status-success);
  margin-bottom: 4px;
}

.qa-panel h3 {
  font-size: 1rem;
  margin: 0;
}

.qa-icon-button {
  border: none;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
  padding: 4px;
}

.qa-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}

.qa-stat {
  padding: 10px;
  border-radius: 12px;
  background: rgba(0, 0, 0, 0.03);
  border: 1px solid var(--border-light);
}

.qa-stat span {
  display: block;
  font-size: 0.68rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--text-muted);
  margin-bottom: 4px;
}

.qa-stat strong {
  display: block;
  font-size: 0.9rem;
  color: var(--text-primary);
  word-break: break-word;
}

.qa-actions {
  display: flex;
  gap: 8px;
  margin-top: 14px;
}

.qa-actions .btn {
  flex: 1;
  font-size: 0.8rem;
}

.qa-footer {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 12px;
  font-size: 0.78rem;
  color: var(--text-secondary);
}

@media (max-width: 640px) {
  .qa-panel,
  .qa-fab {
    right: 12px;
    left: 12px;
    width: auto;
  }
}

.log-incident-btn {
  background: linear-gradient(135deg, var(--primary), #ff6b35) !important;
  border: none !important;
  color: white !important;
  font-weight: 600 !important;
  padding: 8px 16px !important;
  border-radius: 8px !important;
  box-shadow: 0 4px 15px rgba(233, 69, 96, 0.4) !important;
  transition: all 0.3s ease !important;
}

.log-incident-btn:hover {
  transform: translateY(-2px) !important;
  box-shadow: 0 6px 20px rgba(233, 69, 96, 0.6) !important;
}

.status-dot {
  display: inline-block;
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background-color: var(--status-success);
  margin-right: 6px;
  box-shadow: 0 0 8px var(--status-success);
  animation: pactPulse 2s infinite;
}

@keyframes pactPulse {
  0% { opacity: 1; transform: scale(1); }
  50% { opacity: 0.5; transform: scale(1.2); }
  100% { opacity: 1; transform: scale(1); }
}
`;

export function injectPactStyles(): void {
  if (typeof document === 'undefined') {
    return;
  }

  try {
    const isWorkbench = typeof window !== 'undefined' && /workbench/i.test(window.location.href);
    if (isWorkbench) {
      document.documentElement.classList.add('pact-workbench');
      if (document.body) {
        document.body.classList.add('pact-workbench');
      }
    }

    if (document.getElementById(STYLE_ID)) {
      return;
    }

    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.type = 'text/css';
    style.appendChild(document.createTextNode(CSS));

    const target = document.head || document.documentElement;
    if (target) {
      target.appendChild(style);
    }
  } catch (error) {
    // Styling must never block web part bootstrap.
    console.error('PACT style injection failed', error);
  }
}

injectPactStyles();
