import React, { useState } from 'react';
import { Search, Clock, ChevronDown, ChevronRight, Eye, RefreshCw } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { MailLogEntry } from '../../config/types';
import { useSharePointCollection } from '../../hooks/useSharePointCollection';

export const MailLogPage: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState('');
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const { data: logs, loading, refresh } = useSharePointCollection<MailLogEntry>(() => sharePointService.getMailHistory());

  const filteredLogs = (logs || []).filter(log => {
    if (!log) return false;
    const subject = log.subject || '';
    const recipients = Array.isArray(log.to) ? log.to.join(',') : String(log.to || '');
    return subject.toLowerCase().includes(searchTerm.toLowerCase()) || 
           recipients.toLowerCase().includes(searchTerm.toLowerCase());
  });

  const getStatusStyles = (status: MailLogEntry['status']): React.CSSProperties => {
    switch (status) {
      case 'Pending':
        return { background: 'rgba(245, 158, 11, 0.12)', color: '#b45309' };
      case 'Processing':
        return { background: 'rgba(59, 130, 246, 0.12)', color: '#2563eb' };
      case 'Failed':
        return { background: 'rgba(220, 38, 38, 0.12)', color: '#b91c1c' };
      case 'Sent':
      default:
        return { background: 'rgba(16, 185, 129, 0.1)', color: '#10b981' };
    }
  };

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px'}} />
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}} />
      </div>
    );
  }

  return (
    <div className="cases-container fade-in">
      <div className="cases-header">
        <div className="search-bar glass-panel">
          <Search size={18} className="search-icon" />
          <input 
            type="text" 
            placeholder="Search mail by recipient or subject..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        <button className="btn btn-secondary" onClick={() => refresh().catch(() => undefined)}>
          <RefreshCw size={16} /> Refresh
        </button>
      </div>

      <div className="cases-table-container glass-panel">
        <table className="pact-table">
          <thead>
            <tr>
              <th style={{ width: '40px' }} />
              <th>Recipient(s)</th>
              <th>Subject</th>
              <th>Timestamp</th>
              <th>Status</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            {filteredLogs.map(log => (
              <React.Fragment key={log.id}>
                <tr>
              <td onClick={() => setExpandedId(expandedId === log.id ? null : log.id)} style={{ cursor: 'pointer' }}>
                {expandedId === log.id ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
              </td>
                  <td style={{ fontWeight: 600 }}>{log.to.join(', ')}</td>
                  <td>{log.subject.substring(0, 50)}{log.subject.length > 50 ? '...' : ''}</td>
                  <td>{new Date(log.timestamp).toLocaleString()}</td>
                  <td>
                    <span className="status-badge" style={getStatusStyles(log.status)}>
                      {log.status}
                    </span>
                  </td>
                  <td>
                    <button 
                      className="link-action" 
                      onClick={() => setExpandedId(expandedId === log.id ? null : log.id)}
                      style={{ background: 'none', border: 'none', cursor: 'pointer' }}
                    >
                      <Eye size={16} style={{ verticalAlign: 'middle', marginRight: '4px' }}/> 
                      {expandedId === log.id ? 'Hide' : 'View Body'}
                    </button>
                  </td>
                </tr>
                {expandedId === log.id && (
                  <tr>
                    <td colSpan={6} style={{ padding: '0' }}>
                      <div style={{ 
                        padding: '1.5rem 2.5rem', 
                        background: 'rgba(255,255,255,0.02)', 
                        borderLeft: '4px solid var(--primary)',
                        animation: 'fadeIn 0.2s ease'
                      }}>
                        <div style={{ 
                          padding: '1.5rem', 
                          background: 'white', 
                          color: '#333', 
                          borderRadius: '8px',
                          fontFamily: 'sans-serif',
                          fontSize: '0.9rem',
                          boxShadow: 'inset 0 2px 4px rgba(0,0,0,0.1)'
                        }} dangerouslySetInnerHTML={{ __html: log.body }} />
                      </div>
                    </td>
                  </tr>
                )}
              </React.Fragment>
            ))}
            {filteredLogs.length === 0 && (
              <tr>
                <td colSpan={6} className="text-center" style={{ padding: '3rem' }}>
                  <Clock size={32} style={{ opacity: 0.3, marginBottom: '1rem' }} />
                  <div>No communication logs found.</div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};
