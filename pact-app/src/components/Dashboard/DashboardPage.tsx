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
import type { ComplianceCase, EscalationEntry, RepeatOffenceRecord, StaffMember } from '../../config/types';
import { Mail, CheckCircle, Clock } from 'lucide-react';
import './Dashboard.css';

export const DashboardPage: React.FC = () => {
  const [allCases, setAllCases] = useState<ComplianceCase[]>([]);
  const [allEscalations, setAllEscalations] = useState<EscalationEntry[]>([]);
  const [allTrackers, setAllTrackers] = useState<RepeatOffenceRecord[]>([]);
  const [allAppeals, setAllAppeals] = useState<any[]>([]);
  const [allStaff, setAllStaff] = useState<StaffMember[]>([]);

  const [selectedCompany, setSelectedCompany] = useState<string>('All');
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

  const loadDashboard = async () => {
    try {
      const [cases, escalations, trackers, appeals, staff, history] = await Promise.all([
        sharePointService.getCases(),
        sharePointService.getEscalationLog(),
        sharePointService.getRepeatTrackerRecords(),
        sharePointService.getAppeals(),
        sharePointService.getStaffDirectory(),
        sharePointService.getMailHistory()
      ]);

      setAllCases(cases);
      setAllEscalations(escalations);
      setAllTrackers(trackers);
      setAllAppeals(appeals);
      setAllStaff(staff);

      if (Array.isArray(history)) {
        setMailLogs(history.slice(0, 5));
      }
      setLoading(false);
    } catch (err) {
      console.error("Failed to load dashboard data:", err);
      setLoading(false);
    }
  };

  // Filter cases, escalations, trackers, appeals based on selected company
  const filteredData = React.useMemo(() => {
    if (selectedCompany === 'All') {
      return {
        cases: allCases,
        escalations: allEscalations,
        trackers: allTrackers,
        appeals: allAppeals
      };
    }

    const selectedUpper = selectedCompany.toUpperCase();

    const getStaffCompanyUpper = (s: StaffMember) => {
      const co = s.company?.toUpperCase();
      return co === 'KESL' ? 'PMT7' : co;
    };

    const filteredCases = allCases.filter(c => {
      const staff = allStaff.find(s => s.id === c.chargedPerson || s.fullName === c.chargedPersonName || s.email === c.staffEmail);
      return staff ? getStaffCompanyUpper(staff) === selectedUpper : false;
    });

    const filteredEscalations = allEscalations.filter(e => {
      const staff = allStaff.find(s => s.id === e.offender || s.fullName === e.offenderName);
      return staff ? getStaffCompanyUpper(staff) === selectedUpper : false;
    });

    const filteredTrackers = allTrackers.filter(t => {
      const staff = allStaff.find(s => s.id === t.offender || s.fullName === t.offenderName);
      return staff ? getStaffCompanyUpper(staff) === selectedUpper : false;
    });

    const filteredAppeals = allAppeals.filter(a => {
      const staff = allStaff.find(s => s.id === a.appellant || s.fullName === a.appellant);
      return staff ? getStaffCompanyUpper(staff) === selectedUpper : false;
    });

    return {
      cases: filteredCases,
      escalations: filteredEscalations,
      trackers: filteredTrackers,
      appeals: filteredAppeals
    };
  }, [allCases, allEscalations, allTrackers, allAppeals, allStaff, selectedCompany]);

  // Rebuild stats using filtered data
  const stats = React.useMemo(() => {
    if (!filteredData.cases.length && selectedCompany !== 'All') {
      return {
        totalActiveCases: 0,
        paidCases: 0,
        appealPendingCases: 0,
        escalationsThisMonth: 0,
        pendingAppeals: 0,
        repeatOffenders: 0,
        totalFines: 0,
        casesByTier: [{ tier: 'Tier 1', count: 0 }, { tier: 'Tier 2', count: 0 }, { tier: 'Tier 3', count: 0 }],
        casesByMonth: [],
        casesByDepartment: [],
        recentActivity: []
      };
    }
    return sharePointService.buildDashboardStats(
      filteredData.cases,
      filteredData.escalations,
      filteredData.trackers,
      filteredData.appeals
    );
  }, [filteredData, selectedCompany]);

  // Calculate Recovery Progress parameters
  const paidFines = React.useMemo(() => {
    return filteredData.cases
      .filter(c => c.status === 'Paid')
      .reduce((sum, c) => sum + c.penaltyAmount, 0);
  }, [filteredData.cases]);

  const totalFines = React.useMemo(() => {
    return stats.totalFines;
  }, [stats]);

  const recoveryRate = React.useMemo(() => {
    return totalFines > 0 ? Math.round((paidFines / totalFines) * 100) : 0;
  }, [paidFines, totalFines]);

  // Find any active Tier 3 (critical) escalation or case for the alert banner
  const criticalItem = React.useMemo(() => {
    return filteredData.cases.find(c => c.tier === 'Tier 3' && c.status !== 'Paid' && c.status !== 'Waived');
  }, [filteredData.cases]);

  useEffect(() => {
    loadDashboard();

    const handleDataChanged = () => {
      // Clear local SharePointService memory cache too to guarantee fresh fetches
      sharePointService.clearCache();
      loadDashboard();
    };
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
          {[1, 2, 3, 4].map(i => <div key={i} className="kpi-card glass-panel skeleton" style={{ height: '120px', borderRadius: '12px' }}></div>)}
        </div>
        <div className="dashboard-charts-grid" style={{ marginTop: '2rem' }}>
          <div className="chart-panel glass-panel skeleton" style={{ height: '400px', borderRadius: '12px' }}></div>
          <div className="chart-panel glass-panel skeleton" style={{ height: '400px', borderRadius: '12px' }}></div>
        </div>
      </div>
    );
  }

  return (
    <div className="dashboard-container fade-in" data-testid="dashboard-page">
      <div className="dashboard-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px', flexWrap: 'wrap', gap: '16px' }}>
        <h2 style={{ margin: 0, color: 'var(--text-primary)', fontWeight: 700 }}>Compliance Dashboard</h2>
        
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap' }}>
          {/* Company filter */}
          <div className="company-filter-wrapper">
            <label htmlFor="company-filter">View by:</label>
            <select
              id="company-filter"
              value={selectedCompany}
              onChange={e => setSelectedCompany(e.target.value)}
              className="company-select-input"
            >
              <option value="All">All Companies</option>
              <option value="KCC">KCC</option>
              <option value="INTERKONSTRUCT">INTERKONSTRUCT</option>
              <option value="PMT7">PMT7</option>
              <option value="NGNEERED">NGNEERED</option>
            </select>
          </div>

          <NavLink to="/cases/new" className="btn btn-primary" data-testid="dashboard-new-case-button" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Plus size={16} /> Log New Case
          </NavLink>
        </div>
      </div>

      {/* Critical Violation Alert Banner */}
      {criticalItem && (
        <div className="high-priority-banner">
          <div className="banner-icon">
            <ShieldAlert size={28} />
          </div>
          <div className="banner-content">
            <h4 className="banner-title">Critical Attention Required</h4>
            <p className="banner-desc">
              Case <strong>{criticalItem.title}</strong> ({criticalItem.chargedPersonName}) is flagged as a <strong>Tier 3 (Critical)</strong> breach. Immediate follow-up is recommended.
            </p>
          </div>
          <NavLink to={`/cases`} className="banner-action-btn" style={{ textDecoration: 'none' }}>
            Review Cases
          </NavLink>
        </div>
      )}

      {/* Top KPI Summary Cards */}
      <div className="kpi-grid">
        {/* Open Cases */}
        <div className="kpi-card glass-panel" style={{ borderLeft: '4px solid var(--primary)', position: 'relative' }}>
          <div className="kpi-icon info" style={{ background: 'linear-gradient(135deg, rgba(233, 69, 96, 0.2), rgba(233, 69, 96, 0.05))', color: 'var(--primary)' }}>
            <ShieldAlert size={24} />
          </div>
          <div className="kpi-content" style={{ flex: 1 }}>
            <span className="kpi-label">Open Cases</span>
            <span className="kpi-value">{stats.totalActiveCases}</span>
            <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: 4 }}>
              {stats.paidCases} paid · {stats.appealPendingCases} in appeal
            </span>
          </div>
        </div>

        {/* Escalations */}
        <div className="kpi-card glass-panel" style={{ borderLeft: '4px solid var(--status-warning)' }}>
          <div className="kpi-icon warning" style={{ background: 'linear-gradient(135deg, rgba(245, 158, 11, 0.2), rgba(245, 158, 11, 0.05))', color: 'var(--status-warning)' }}>
            <AlertTriangle size={24} />
          </div>
          <div className="kpi-content">
            <span className="kpi-label">Escalations (Month)</span>
            <span className="kpi-value">{stats.escalationsThisMonth}</span>
            <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: 4 }}>
              Requires review
            </span>
          </div>
        </div>

        {/* Financial Recovery (Includes Progress Gauge) */}
        <div className="kpi-card glass-panel" style={{ borderLeft: '4px solid #10b981', minHeight: '120px' }}>
          <div className="kpi-content" style={{ width: '100%' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', marginBottom: '4px' }}>
              <span className="kpi-label">Recovery Rate</span>
              <span style={{ fontSize: '0.85rem', fontWeight: 700, color: '#10b981' }}>{recoveryRate}%</span>
            </div>
            <span className="kpi-value">₦{stats.totalFines.toLocaleString()}</span>
            <div className="recovery-progress-container">
              <div className="recovery-progress-bar" style={{ width: `${recoveryRate}%` }}></div>
            </div>
            <span style={{ fontSize: '0.7rem', color: 'var(--text-secondary)', marginTop: 6, display: 'inline-block' }}>
              ₦{paidFines.toLocaleString()} recovered of total
            </span>
          </div>
        </div>

        {/* Repeat Offenders */}
        <div className="kpi-card glass-panel" style={{ borderLeft: '4px solid #ef4444' }}>
          <div className="kpi-icon danger" style={{ background: 'linear-gradient(135deg, rgba(239, 68, 68, 0.2), rgba(239, 68, 68, 0.05))', color: '#ef4444' }}>
            <Users size={24} />
          </div>
          <div className="kpi-content">
            <span className="kpi-label">Repeat Offenders</span>
            <span className="kpi-value">{stats.repeatOffenders}</span>
            <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: 4 }}>
              Active watch list
            </span>
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
