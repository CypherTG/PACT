import React, { useEffect, useState } from 'react';
import { AlertTriangle, Filter } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { EscalationEntry } from '../../config/types';

export const EscalationLog: React.FC = () => {
  const [logs, setLogs] = useState<EscalationEntry[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    sharePointService.getEscalationLog().then(data => {
      setLogs(data);
      setLoading(false);
    });
  }, []);

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px', borderRadius: 'var(--radii-md)'}}></div>
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}}></div>
      </div>
    );
  }

  return (
    <div className="cases-container fade-in">
      <div className="cases-header">
        <h2 style={{margin: 0, display: 'flex', alignItems: 'center', gap: '8px'}}>
          <AlertTriangle size={20} color="var(--primary)"/> Escalation Audit Trail
        </h2>
        <button className="btn btn-secondary"><Filter size={16} /> Filter Logs</button>
      </div>

      <div className="cases-table-container glass-panel">
        <table className="pact-table">
            <thead>
              <tr>
                <th>Date</th>
                <th>Offender</th>
                <th>Trigger Case</th>
                <th>Offence Tier</th>
                <th>Escalated Tier</th>
                <th>Trigger Type</th>
                <th>Reason</th>
              </tr>
            </thead>
            <tbody>
              {logs.map(log => {
                const prevTier = String(log.previousTier || 'Tier 1');
                const nextTier = String(log.newTier || 'Tier 2');
                const trigger = String(log.triggeredBy || 'System');
                const escDate = log.escalationDate ? new Date(log.escalationDate) : null;
                const dateStr = escDate && !isNaN(escDate.getTime()) ? escDate.toLocaleDateString() : '—';
                return (
                <tr key={log.id}>
                  <td>{dateStr}</td>
                  <td style={{fontWeight: 600}}>
                    {log.offenderName || `Staff ID: ${log.offender || '—'}`}
                  </td>
                  <td className="link-action" style={{cursor:'pointer'}}>{log.caseReference || '—'}</td>
                  <td>
                    <span className={`status-badge status-${prevTier.toLowerCase().replace(' ', '')}`}>
                      {prevTier}
                    </span>
                  </td>
                  <td>
                    <span className="text-secondary">{prevTier}</span> 
                    <span style={{margin: '0 8px', color: 'var(--primary)'}}>→</span> 
                    <span className={`status-badge status-${nextTier.toLowerCase().replace(' ', '')}`}>
                      {nextTier}
                    </span>
                  </td>
                  <td>
                    <span className={`status-badge ${trigger === 'System' ? 'status-paid' : 'status-unpaid'}`}>
                      {trigger}
                    </span>
                  </td>
                  <td className="text-secondary" style={{maxWidth: '300px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis'}} title={String(log.escalationReason || '')}>
                    {log.escalationReason || '—'}
                  </td>
                </tr>
                );
              })}
              {logs.length === 0 && (
                <tr>
                  <td colSpan={6} className="text-center">No escalations recorded yet.</td>
                </tr>
              )}
            </tbody>
          </table>
      </div>
    </div>
  );
};
