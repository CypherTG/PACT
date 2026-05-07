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
import type { DashboardStats } from '../../config/types';
import { Mail, CheckCircle, Clock } from 'lucide-react';
import './Dashboard.css';

export const DashboardPage: React.FC = () => {
  const [stats, setStats] = useState<DashboardStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [mailLogs, setMailLogs] = useState<any[]>([]);

  useEffect(() => {
    sharePointService.getDashboardStats().then(data => {
      setStats(data);
      setLoading(false);
    });

    sharePointService.getMailHistory().then(history => {
      if (Array.isArray(history)) {
        setMailLogs(history.slice(0, 5));
      }
    });

    const handleMailEvent = (e: any) => {
      if (!e.detail || !e.detail.to) return;
      const recipientStr = Array.isArray(e.detail.to) ? e.detail.to.join(', ') : String(e.detail.to);
      
      setMailLogs(prev => [
        { id: Date.now(), to: recipientStr, subject: e.detail.subject || 'No Subject', time: new Date().toLocaleTimeString() },
        ...prev
      ].slice(0, 5));
    };

    window.addEventListener('pact-mock-email', handleMailEvent);
    return () => window.removeEventListener('pact-mock-email', handleMailEvent);
  }, []);

  if (loading || !stats) {
    return (
      <div className="dashboard-container">
        <div className="kpi-grid">
          {[1,2,3,4].map(i => <div key={i} className="kpi-card glass-panel skeleton" style={{height: '100px'}}></div>)}
        </div>
        <div className="dashboard-charts-grid" style={{marginTop: '2rem'}}>
          <div className="chart-panel glass-panel skeleton" style={{height: '400px'}}></div>
          <div className="chart-panel glass-panel skeleton" style={{height: '400px'}}></div>
        </div>
      </div>
    );
  }

  return (
    <div className="dashboard-container fade-in" data-testid="dashboard-page">
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
    </div>
  );
};
