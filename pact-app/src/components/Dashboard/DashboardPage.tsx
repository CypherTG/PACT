import React, { useEffect, useState } from 'react';
import { NavLink } from 'react-router-dom';
import {
  Users,
  AlertTriangle,
  ShieldAlert,
  TrendingUp,
  Plus
} from 'lucide-react';
import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  BarChart, Bar, Legend
} from 'recharts';
import { sharePointService } from '../../services/SharePointService';
import type { DashboardStats } from '../../config/types';
import { Mail, CheckCircle, Clock } from 'lucide-react';
import './Dashboard.css';

export const DashboardPage: React.FC = () => {
  const [stats, setStats] = useState<DashboardStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [mailLogs, setMailLogs] = useState<any[]>([]);

  // System Diagnostics States
  const [showDiagnostics, setShowDiagnostics] = useState(false);
  const [selectedDiagnosticList, setSelectedDiagnosticList] = useState('PACT Compliance Cases');
  const [diagnosticResult, setDiagnosticResult] = useState<string>('');
  const [diagnosticLoading, setDiagnosticLoading] = useState(false);

  const runColumnDiagnostic = async () => {
    setDiagnosticLoading(true);
    try {
      const res = await sharePointService.getListColumnsDiagnostic(selectedDiagnosticList);
      setDiagnosticResult(res.join('\n'));
    } catch (e) {
      setDiagnosticResult(String(e));
    }
    setDiagnosticLoading(false);
  };

  const runFirstItemDiagnostic = async () => {
    setDiagnosticLoading(true);
    try {
      const res = await sharePointService.getRawFirstItemDiagnostic(selectedDiagnosticList);
      setDiagnosticResult(res);
    } catch (e) {
      setDiagnosticResult(String(e));
    }
    setDiagnosticLoading(false);
  };

  const loadDashboard = () => {
    sharePointService.getDashboardStats().then(data => {
      setStats(data);
      setLoading(false);
    });
    sharePointService.getMailHistory().then(history => {
      if (Array.isArray(history)) {
        setMailLogs(history.slice(0, 5));
      }
    });
  };

  useEffect(() => {
    loadDashboard();

    const handleDataChanged = () => loadDashboard();
    window.addEventListener('pact-data-changed', handleDataChanged);

    const handleMailEvent = (e: any) => {
      if (!e.detail || !e.detail.to) return;
      const recipientStr = Array.isArray(e.detail.to) ? e.detail.to.join(', ') : String(e.detail.to);

      setMailLogs(prev => [
        { id: Date.now(), to: recipientStr, subject: e.detail.subject || 'No Subject', time: new Date().toLocaleTimeString() },
        ...prev
      ].slice(0, 5));
    };

    window.addEventListener('pact-mock-email', handleMailEvent);
    return () => {
      window.removeEventListener('pact-mock-email', handleMailEvent);
      window.removeEventListener('pact-data-changed', handleDataChanged);
    };
  }, []);

  if (loading || !stats) {
    return (
      <div className="dashboard-container">
        <div className="kpi-grid">
          {[1, 2, 3, 4].map(i => <div key={i} className="kpi-card glass-panel skeleton" style={{ height: '100px' }}></div>)}
        </div>
        <div className="dashboard-charts-grid" style={{ marginTop: '2rem' }}>
          <div className="chart-panel glass-panel skeleton" style={{ height: '400px' }}></div>
          <div className="chart-panel glass-panel skeleton" style={{ height: '400px' }}></div>
        </div>
      </div>
    );
  }

  return (
    <div className="dashboard-container fade-in" data-testid="dashboard-page">
      <div className="dashboard-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
        <h2 style={{ margin: 0, color: 'var(--text-primary)', fontWeight: 700 }}>Compliance Dashboard</h2>
        <NavLink to="/cases/new" className="btn btn-primary" data-testid="dashboard-new-case-button">
          <Plus size={16} /> Log New Case
        </NavLink>
      </div>

      {/* Top KPI Summary Cards */}
      <div className="kpi-grid">
        <div className="kpi-card glass-panel">
          <div className="kpi-icon info"><ShieldAlert size={24} /></div>
          <div className="kpi-content">
            <span className="kpi-label">Open Cases</span>
            <span className="kpi-value">{stats.totalActiveCases}</span>
            <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: 4 }}>
              {stats.paidCases} paid · {stats.appealPendingCases} appeal pending
            </span>
          </div>
        </div>

        <div className="kpi-card glass-panel">
          <div className="kpi-icon warning"><AlertTriangle size={24} /></div>
          <div className="kpi-content">
            <span className="kpi-label">Escalations This Month</span>
            <span className="kpi-value">{stats.escalationsThisMonth}</span>
          </div>
        </div>

        <div className="kpi-card glass-panel">
          <div className="kpi-icon primary"><TrendingUp size={24} /></div>
          <div className="kpi-content">
            <span className="kpi-label">Financial Recovery</span>
            <span className="kpi-value">₦{stats.totalFines.toLocaleString()}</span>
          </div>
        </div>

        <div className="kpi-card glass-panel">
          <div className="kpi-icon danger"><Users size={24} /></div>
          <div className="kpi-content">
            <span className="kpi-label">Repeat Offenders Watch</span>
            <span className="kpi-value">{stats.repeatOffenders}</span>
          </div>
        </div>
      </div>

      <div className="dashboard-charts-grid">
        {/* Compliance Trend Line Chart */}
        <div className="chart-panel glass-panel">
          <div className="panel-header">
            <h3><TrendingUp size={18} /> Compliance Breach Trends</h3>
          </div>
          <div className="chart-wrapper">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={stats.casesByMonth} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                <defs>
                  <linearGradient id="colorO" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#e94560" stopOpacity={0.3} />
                    <stop offset="95%" stopColor="#e94560" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border-light)" vertical={false} />
                <XAxis dataKey="month" stroke="var(--text-secondary)" tick={{ fill: 'var(--text-secondary)' }} />
                <YAxis stroke="var(--text-secondary)" tick={{ fill: 'var(--text-secondary)' }} />
                <RechartsTooltip
                  contentStyle={{ backgroundColor: 'var(--bg-surface)', borderColor: 'var(--border-light)', borderRadius: '8px' }}
                  itemStyle={{ color: 'var(--text-primary)' }}
                />
                <Area type="monotone" dataKey="tier1" name="Tier 1" stroke="#0078d4" fillOpacity={1} fill="url(#colorO)" />
                <Area type="monotone" dataKey="tier2" name="Tier 2" stroke="#f59e0b" fillOpacity={0.6} fill="#f59e0b" />
                <Area type="monotone" dataKey="tier3" name="Tier 3" stroke="#e94560" fillOpacity={0.6} fill="#e94560" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Risk Heatmap (Bar Chart Approximation for now) */}
        <div className="chart-panel glass-panel">
          <div className="panel-header">
            <h3>Department Risk Heatmap</h3>
          </div>
          <div className="chart-wrapper">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={stats.casesByDepartment} layout="vertical" margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border-light)" horizontal={false} />
                <XAxis type="number" stroke="var(--text-secondary)" />
                <YAxis dataKey="department" type="category" stroke="var(--text-secondary)" width={100} />
                <RechartsTooltip
                  cursor={{ fill: 'rgba(0,0,0,0.03)' }}
                  contentStyle={{ backgroundColor: 'var(--bg-surface)', borderColor: 'var(--border-light)', borderRadius: '8px' }}
                />
                <Legend />
                <Bar dataKey="count" name="Case Volume" fill="#0078d4" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="dashboard-lists-grid">
        {/* Recent Activity Feed */}
        <div className="activity-panel glass-panel" data-testid="dashboard-activity-feed">
          <div className="panel-header">
            <h3>Recent Compliance Activity</h3>
          </div>
          <div className="activity-list">
            {stats.recentActivity.map(activity => (
              <div key={activity.id} className="activity-item">
                <div className={`activity-indicator severity-${activity.severity}`}></div>
                <div className="activity-content">
                  <div className="activity-title">{activity.title}</div>
                  <div className="activity-desc">{activity.description}</div>
                  <div className="activity-time">{new Date(activity.timestamp).toLocaleString()}</div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Communication Log */}
        <div className="activity-panel glass-panel" data-testid="dashboard-mail-log">
          <div className="panel-header" style={{ display: 'flex', justifyContent: 'space-between' }}>
            <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Mail size={18} color="var(--primary)" /> Mail Notification Log
            </h3>
            <span style={{ fontSize: '0.7rem', background: 'rgba(255,255,255,0.05)', padding: '2px 8px', borderRadius: '4px' }}>Real-time</span>
          </div>
          <div className="activity-list">
            {mailLogs.length === 0 ? (
              <div className="text-center text-secondary" style={{ padding: '2rem', fontSize: '0.85rem' }}>
                <Clock size={24} style={{ display: 'block', margin: '0 auto 10px', opacity: 0.5 }} />
                No mail transmissions in this session.
              </div>
            ) : (
              mailLogs.map(log => (
                <div key={log.id} className="activity-item" style={{ borderLeft: '3px solid var(--status-success)', background: 'rgba(16, 185, 129, 0.03)' }}>
                  <div className="activity-content">
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <div className="activity-title" style={{ fontSize: '0.85rem', color: 'var(--text-primary)' }}>{log.subject}</div>
                      <CheckCircle size={14} color="var(--status-success)" />
                    </div>
                    <div className="activity-desc" style={{ fontSize: '0.75rem' }}>To: {log.to}</div>
                    <div className="activity-time">{log.time}</div>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* Collapsible Diagnostics Panel */}
      <div className="glass-panel" style={{ marginTop: '2rem', padding: '1.5rem', border: '1px dashed rgba(255,255,255,0.1)' }}>
        <div 
          onClick={() => setShowDiagnostics(!showDiagnostics)} 
          style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}
        >
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px', fontSize: '1.05rem', color: 'var(--text-secondary)' }}>
            ⚙️ System Diagnostics & Schema Inspector (UAT Debug Tool)
          </h3>
          <span style={{ fontSize: '0.85rem', color: 'var(--primary)' }}>
            {showDiagnostics ? '▲ Hide Diagnostics' : '▼ Show Diagnostics'}
          </span>
        </div>

        {showDiagnostics && (
          <div style={{ marginTop: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1.5rem', borderTop: '1px solid var(--border-light)', paddingTop: '1.5rem' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
              <div>
                <h4 style={{ margin: '0 0 0.5rem 0', fontSize: '0.9rem', color: 'var(--text-primary)' }}>1. Environment Context</h4>
                <table style={{ width: '100%', fontSize: '0.85rem', borderCollapse: 'collapse' }}>
                  <tbody>
                    <tr style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}><td style={{ padding: '6px 0', color: 'var(--text-secondary)' }}>App Runtime Mode:</td><td style={{ fontWeight: 600 }}>{sharePointService.getRuntimeLabel()}</td></tr>
                    <tr style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}><td style={{ padding: '6px 0', color: 'var(--text-secondary)' }}>Is Local Standalone:</td><td style={{ fontWeight: 600 }}>{sharePointService.isStandalone() ? 'Yes (Demo Mocks)' : 'No (SharePoint Native)'}</td></tr>
                    <tr style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}><td style={{ padding: '6px 0', color: 'var(--text-secondary)' }}>Logged In User:</td><td style={{ fontWeight: 600 }}>{sharePointService.getCurrentSession()?.email || 'N/A'}</td></tr>
                    <tr style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}><td style={{ padding: '6px 0', color: 'var(--text-secondary)' }}>Site Domain:</td><td style={{ fontWeight: 600, wordBreak: 'break-all' }}>{sharePointService.getCurrentSession()?.siteTitle || 'N/A'}</td></tr>
                  </tbody>
                </table>
              </div>

              <div>
                <h4 style={{ margin: '0 0 0.5rem 0', fontSize: '0.9rem', color: 'var(--text-primary)' }}>2. SharePoint List Selector</h4>
                <div style={{ display: 'flex', gap: '10px', marginBottom: '1rem' }}>
                  <select 
                    value={selectedDiagnosticList}
                    onChange={e => setSelectedDiagnosticList(e.target.value)}
                    style={{ 
                      flex: 1, 
                      padding: '8px', 
                      background: 'rgba(0,0,0,0.2)', 
                      border: '1px solid var(--border-light)', 
                      borderRadius: '6px', 
                      color: 'var(--text-primary)',
                      outline: 'none'
                    }}
                  >
                    <option value="PACT Compliance Cases">PACT Compliance Cases</option>
                    <option value="PACT Policy & Offence Library">PACT Policy & Offence Library</option>
                    <option value="PACT Staff Directory">PACT Staff Directory</option>
                    <option value="PACT Repeat Offence Tracker">PACT Repeat Offence Tracker</option>
                    <option value="PACT Appeals Register">PACT Appeals Register</option>
                    <option value="Escalation Log">Escalation Log</option>
                    <option value="Disciplinary Actions">Disciplinary Actions</option>
                    <option value="PACT Mail History">PACT Mail History</option>
                  </select>
                </div>
                <div style={{ display: 'flex', gap: '10px' }}>
                  <button 
                    onClick={runColumnDiagnostic} 
                    className="btn btn-secondary" 
                    style={{ padding: '8px 12px', fontSize: '0.8rem' }}
                    disabled={diagnosticLoading}
                  >
                    Inspect Columns Schema
                  </button>
                  <button 
                    onClick={runFirstItemDiagnostic} 
                    className="btn btn-secondary" 
                    style={{ padding: '8px 12px', fontSize: '0.8rem' }}
                    disabled={diagnosticLoading}
                  >
                    View Raw First Item
                  </button>
                </div>
              </div>
            </div>

            {diagnosticResult && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', fontWeight: 600 }}>Diagnostic Output for list: "{selectedDiagnosticList}"</span>
                  <button 
                    onClick={() => { navigator.clipboard.writeText(diagnosticResult); alert('Copied to clipboard!'); }}
                    className="link-action"
                    style={{ fontSize: '0.8rem', cursor: 'pointer', background: 'none', border: 'none', color: 'var(--primary)', padding: 0 }}
                  >
                    Copy Output
                  </button>
                </div>
                <pre style={{ 
                  background: 'rgba(0,0,0,0.4)', 
                  padding: '1rem', 
                  borderRadius: '6px', 
                  overflow: 'auto', 
                  maxHeight: '300px', 
                  fontSize: '0.8rem', 
                  fontFamily: 'monospace',
                  color: '#a7f3d0',
                  border: '1px solid var(--border-light)',
                  margin: 0
                }}>
                  {diagnosticLoading ? 'Running diagnostics...' : diagnosticResult}
                </pre>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};
