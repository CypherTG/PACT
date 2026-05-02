import React, { useEffect, useState } from 'react';
import { Search, Clock, ChevronDown, ChevronRight, Eye } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';

export const MailLogPage: React.FC = () => {
  const [logs, setLogs] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [expandedId, setExpandedId] = useState<number | null>(null);

  useEffect(() => {
    sharePointService.getMailHistory().then(data => {
      setLogs(data);
      setLoading(false);
    });
  }, []);

  const filteredLogs = (logs || []).filter(log => {
    if (!log) return false;
    const subject = log.subject || '';
    const recipients = Array.isArray(log.to) ? log.to.join(',') : String(log.to || '');
    return subject.toLowerCase().includes(searchTerm.toLowerCase()) || 
           recipients.toLowerCase().includes(searchTerm.toLowerCase());
  });

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px'}}></div>
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}}></div>
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
      </div>

      <div className="cases-table-container glass-panel">
        <table className="pact-table">
          <thead>
            <tr>
              <th style={{ width: '40px' }}></th>
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
                    <span className="status-badge" style={{ background: 'rgba(16, 185, 129, 0.1)', color: '#10b981' }}>
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
