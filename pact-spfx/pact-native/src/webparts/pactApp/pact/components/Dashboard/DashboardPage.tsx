// @ts-nocheck
import React, { useEffect, useState } from 'react';
import { 
  Users, 
  AlertTriangle, 
  ShieldAlert, 
  FileSearch,
  TrendingUp
} from 'lucide-react';
import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  BarChart, Bar, Legend
} from 'recharts';
import { sharePointService } from '../../services/SharePointService';
import type { DashboardStats, MailLogEntry } from '../../config/types';
import { Mail, CheckCircle, Clock, RefreshCw } from 'lucide-react';
import { useSharePointCollection } from '../../hooks/useSharePointCollection';

export const DashboardPage = (): React.ReactElement => {
  const [stats, setStats] = useState<DashboardStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [backendState, setBackendState] = useState<'checking' | 'online' | 'offline'>('checking');
  const [backendDetail, setBackendDetail] = useState<string>('Checking SharePoint list connectivity...');
  const { data: mailLogs, refresh: refreshMailLogs } = useSharePointCollection<MailLogEntry>(() => sharePointService.getMailHistory());

  useEffect(() => {
    const checkBackend = async (): Promise<void> => {
      try {
        const [staff, policies, cases] = await Promise.all([
          sharePointService.getStaffDirectory(),
          sharePointService.getPolicyLibrary(),
          sharePointService.getCases(),
        ]);
        setBackendState('online');
        setBackendDetail(`Live SharePoint reads succeeded: ${staff.length} staff, ${policies.length} policies, ${cases.length} cases.`);
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Unknown SharePoint read failure';
        setBackendState('offline');
        setBackendDetail(`SharePoint read failed: ${message}`);
      }
    };

    const loadDashboard = async (): Promise<void> => {
      try {
        const data = await sharePointService.getDashboardStats();
        setStats(data);
      } finally {
        setLoading(false);
      }
    };

    checkBackend().catch(() => undefined);
    loadDashboard().catch(() => undefined);
    refreshMailLogs().catch(() => undefined);
  }, []);

  if (loading || !stats) {
    return (
      <div className="dashboard-container">
        <div className="kpi-grid">
          {[1,2,3,4].map(i => <div key={i} className="kpi-card glass-panel skeleton" style={{height: '100px'}} />)}
        </div>
        <div className="dashboard-charts-grid" style={{marginTop: '2rem'}}>
          <div className="chart-panel glass-panel skeleton" style={{height: '400px'}} />
          <div className="chart-panel glass-panel skeleton" style={{height: '400px'}} />
        </div>
      </div>
    );
  }

  return (
    <div className="dashboard-container fade-in" data-testid="dashboard-page">
      <div
        className="glass-panel"
        style={{
          marginBottom: '1rem',
          padding: '14px 18px',
          borderLeft: backendState === 'online' ? '4px solid #10b981' : backendState === 'offline' ? '4px solid #ef4444' : '4px solid #f59e0b',
          background: backendState === 'online'
            ? 'rgba(16, 185, 129, 0.06)'
            : backendState === 'offline'
              ? 'rgba(239, 68, 68, 0.06)'
              : 'rgba(245, 158, 11, 0.06)',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          gap: '12px',
          flexWrap: 'wrap'
        }}
      >
        <div>
          <div style={{ fontSize: '0.75rem', textTransform: 'uppercase', opacity: 0.7, letterSpacing: '0.08em' }}>
            SharePoint Backend Status
          </div>
          <div style={{ fontSize: '0.95rem', fontWeight: 600, color: 'var(--text-primary)' }}>
            {backendState === 'online' ? 'Online' : backendState === 'offline' ? 'Offline' : 'Checking'}
          </div>
          <div style={{ fontSize: '0.82rem', color: 'var(--text-secondary)' }}>
            {backendDetail}
          </div>
        </div>
        <button className="btn btn-secondary" onClick={() => {
          setBackendState('checking');
          setBackendDetail('Rechecking SharePoint list connectivity...');
          Promise.all([
            sharePointService.getStaffDirectory(),
            sharePointService.getPolicyLibrary(),
            sharePointService.getCases(),
          ]).then(([staff, policies, cases]) => {
            setBackendState('online');
            setBackendDetail(`Live SharePoint reads succeeded: ${staff.length} staff, ${policies.length} policies, ${cases.length} cases.`);
          }).catch((error: unknown) => {
            const message = error instanceof Error ? error.message : 'Unknown SharePoint read failure';
            setBackendState('offline');
            setBackendDetail(`SharePoint read failed: ${message}`);
          });
        }} data-testid="dashboard-backend-refresh">
          Refresh Backend Check
        </button>
      </div>

      <div className="cases-header" style={{ justifyContent: 'flex-end', marginBottom: '1rem' }}>
        <button className="btn btn-secondary" onClick={() => refreshMailLogs().catch(() => undefined)} data-testid="dashboard-mail-refresh">
          <RefreshCw size={16} /> Refresh Mail Queue
        </button>
      </div>
      {/* Top KPI Summary Cards */}
      <div className="kpi-grid">
        <div className="kpi-card glass-panel">
          <div className="kpi-icon info"><ShieldAlert size={24} /></div>
          <div className="kpi-content">
            <span className="kpi-label">Active Cases</span>
            <span className="kpi-value">{stats.totalActiveCases}</span>
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
          <div className="kpi-icon primary"><FileSearch size={24} /></div>
          <div className="kpi-content">
            <span className="kpi-label">Pending Appeals</span>
            <span className="kpi-value">{stats.pendingAppeals}</span>
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
            <h3><TrendingUp size={18}/> Compliance Breach Trends</h3>
          </div>
          <div className="chart-wrapper">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={stats.casesByMonth} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                <defs>
                  <linearGradient id="colorO" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#e94560" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#e94560" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border-light)" vertical={false} />
                <XAxis dataKey="month" stroke="var(--text-secondary)" tick={{fill: 'var(--text-secondary)'}} />
                <YAxis stroke="var(--text-secondary)" tick={{fill: 'var(--text-secondary)'}} />
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
                  cursor={{fill: 'rgba(0,0,0,0.03)'}}
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
                <div className={`activity-indicator severity-${activity.severity}`} />
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
              <Mail size={18} color="var(--primary)"/> Mail Notification Log
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
                <div
                  key={log.id}
                  className="activity-item"
                  style={{
                    borderLeft: log.status === 'Failed'
                      ? '3px solid #b91c1c'
                      : log.status === 'Processing'
                        ? '3px solid #2563eb'
                        : log.status === 'Pending'
                          ? '3px solid #b45309'
                          : '3px solid var(--status-success)',
                    background: 'rgba(16, 185, 129, 0.03)'
                  }}
                >
                  <div className="activity-content">
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <div className="activity-title" style={{ fontSize: '0.85rem', color: 'var(--text-primary)' }}>{log.subject}</div>
                      {log.status === 'Sent' ? (
                        <CheckCircle size={14} color="var(--status-success)" />
                      ) : (
                        <span
                          style={{
                            fontSize: '0.68rem',
                            padding: '2px 6px',
                            borderRadius: '999px',
                            background: log.status === 'Failed'
                              ? 'rgba(220, 38, 38, 0.12)'
                              : log.status === 'Processing'
                                ? 'rgba(59, 130, 246, 0.12)'
                                : 'rgba(245, 158, 11, 0.12)',
                            color: log.status === 'Failed'
                              ? '#b91c1c'
                              : log.status === 'Processing'
                                ? '#2563eb'
                                : '#b45309'
                          }}
                        >
                          {log.status}
                        </span>
                      )}
                    </div>
                    <div className="activity-desc" style={{ fontSize: '0.75rem' }}>To: {log.to.join(', ')}</div>
                    <div className="activity-time">{new Date(log.timestamp).toLocaleString()}</div>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
