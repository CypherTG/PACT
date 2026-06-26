import React, { useState } from 'react';
import { Search, Clock, ChevronDown, ChevronRight, Eye, RefreshCw, UploadCloud } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { MailLogEntry } from '../../config/types';
import { useSharePointCollection } from '../../hooks/useSharePointCollection';

import histData from '../../data/historicalData.json';

export const MailLogPage: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState('');
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const { data: logs, loading, refresh } = useSharePointCollection<MailLogEntry>(() => sharePointService.getMailHistory());
  const [importing, setImporting] = useState(false);
  const [confirmImport, setConfirmImport] = useState(false);

  const handleImportHistorical = async () => {
    if (!confirmImport) {
      setConfirmImport(true);
      return;
    }
    setConfirmImport(false);
    setImporting(true);
    try {
      // 1. Delete existing historical cases
      const allCases = await sharePointService.getCases();
      const existingHistCases = allCases.filter(c => c.title.startsWith('HIST-'));
      for (const hc of existingHistCases) {
        await sharePointService.deleteCase(hc.id);
      }

      // 1.5 Delete existing historical disciplinary actions
      const allActions = await sharePointService.getDisciplinaryActions();
      const existingHistActions = allActions.filter(a => a.caseReference && a.caseReference.startsWith('HIST-'));
      for (const ha of existingHistActions) {
        await sharePointService.deleteDisciplinaryAction(ha.id);
      }

      // 2. Reset Trackers
      await sharePointService.resetAllTrackers();

      // 3. Fetch dependencies
      const staffList = await sharePointService.getStaffDirectory();
      const policyList = await sharePointService.getPolicyLibrary();

      // 3. Import new data
      let currentHistId = 1;

      for (const record of histData) {
        // Find staff matching the exact name
        const matchedStaff = staffList.find(s => s.fullName === record.employee);
        
        // Find policy mapping roughly
        const recordOffence = String(record.offence || '').trim().toLowerCase();
        let matchedPolicy = policyList.find(p => String(p.offenceName || '').trim().toLowerCase() === recordOffence);
        if (!matchedPolicy) {
          if (recordOffence.includes('read presentation')) matchedPolicy = policyList.find(p => p.id === '1');
          if (recordOffence.includes('ddp')) matchedPolicy = policyList.find(p => p.id === '3');
          if (recordOffence.includes('meet deadline')) matchedPolicy = policyList.find(p => p.id === '4');
          if (recordOffence.includes('meeting early')) matchedPolicy = policyList.find(p => p.id === '4');
          if (recordOffence.includes('naming convention')) matchedPolicy = policyList.find(p => p.id === '1');
          if (recordOffence.includes('upload to dropbox')) matchedPolicy = policyList.find(p => p.id === '1');
          if (recordOffence.includes('attachment')) matchedPolicy = policyList.find(p => p.id === '1');
          if (recordOffence.includes('sleeping')) matchedPolicy = policyList.find(p => p.id === '7');
        }

        const safeName = record.employee && typeof record.employee === 'string' ? record.employee : 'staff';
        const fallbackEmail = `${safeName.replace(/\s+/g, '.').toLowerCase()}@konstructum.com`;

        const nextNumberStr = String(currentHistId).padStart(3, '0');
        const caseTitle = `HIST-${nextNumberStr}`;
        currentHistId++;

        const pData: Partial<any> = {
          title: caseTitle,
          chargedPerson: matchedStaff?.id || '',
          chargedPersonName: record.employee,
          staffEmail: matchedStaff?.email || fallbackEmail,
          offenceCategory: matchedPolicy?.id || '1',
          offenceCategoryName: record.offence,
          offenceDescription: record.offence,
          penaltyAmount: record.penalty,
          status: record.status === 'Paid' ? 'Paid' : 'Unpaid'
        };
        
        await sharePointService.createCase(pData, true);
      }
      alert('Historical data imported and mapped successfully!');
    } catch (err) {
      console.error(err);
      alert('Error importing data: ' + String(err));
    }
    setImporting(false);
  };

  const filteredLogs = (logs || []).filter(log => {
    if (!log) return false;
    const subject = log.subject || '';
    const recipients = Array.isArray(log.to) ? log.to.join(',') : String(log.to || '');
    return (subject || '').toLowerCase().includes((searchTerm || '').toLowerCase()) || 
           (recipients || '').toLowerCase().includes((searchTerm || '').toLowerCase());
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
        <div style={{display: 'flex', gap: '1rem'}}>
          {confirmImport ? (
            <>
              <button className="btn" style={{background: '#d13438', color: '#fff'}} onClick={handleImportHistorical} disabled={importing}>
                <UploadCloud size={16} /> Yes, Import Now
              </button>
              <button className="btn btn-secondary" onClick={() => setConfirmImport(false)}>
                Cancel
              </button>
            </>
          ) : (
            <button className="btn btn-secondary" onClick={handleImportHistorical} disabled={importing}>
              <UploadCloud size={16} /> {importing ? 'Importing...' : 'Import P5 History'}
            </button>
          )}
          <button className="btn btn-secondary" onClick={() => refresh().catch(() => undefined)}>
            <RefreshCw size={16} /> Refresh
          </button>
        </div>
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
