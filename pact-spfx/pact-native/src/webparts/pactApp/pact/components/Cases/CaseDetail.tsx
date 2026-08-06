import React, { useEffect, useState } from 'react';
import { NavLink, useParams, useHistory } from 'react-router-dom';
import { ShieldAlert, ArrowLeft, Clock, User, AlertTriangle } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';

export const CaseDetail: React.FC = () => {
  const { id } = useParams<{ id: string }>();
  const history = useHistory();
  const [caseData, setCaseData] = useState<ComplianceCase | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchCase = () => {
      // In a real app, we'd fetch the specific case.
      // For now, we fetch all and find it.
      sharePointService.getCases().then(cases => {
        const found = cases.find(c => c.id === id || c.title === id);
        if (found) setCaseData(found);
        setLoading(false);
      });
    };

    fetchCase();
    window.addEventListener('pact-data-changed', fetchCase);
    return () => window.removeEventListener('pact-data-changed', fetchCase);
  }, [id]);

  const handleDelete = async () => {
    if (window.confirm("Are you sure you want to permanently delete this incident?")) {
      await sharePointService.deleteCase(caseData!.id);
      history.push('/cases');
    }
  };

  const handleStatusChange = async (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newStatus = e.target.value as any;
    await sharePointService.updateCase(caseData!.id, { status: newStatus });
    setCaseData({ ...caseData!, status: newStatus });
  };

  if (loading) return <div className="glass-panel" style={{padding: '40px', textAlign: 'center'}}>Loading case details...</div>;
  if (!caseData) return <div className="glass-panel" style={{padding: '40px', textAlign: 'center'}}>Case not found.</div>;

  const hasLateFee = !!caseData.escalationApplied;
  const lateFeeAmount = caseData.escalationFeeAmount || 0;
  const totalDue = hasLateFee ? (caseData.totalAmountDue || (caseData.penaltyAmount + lateFeeAmount)) : caseData.penaltyAmount;

  const getSafeDate = (d: string) => {
    if (!d) return 'Just now';
    const parsed = new Date(d);
    return isNaN(parsed.getTime()) ? 'Just now' : parsed.toLocaleDateString();
  };
  const getSafeDateTime = (d: string) => {
    if (!d) return 'Just now';
    const parsed = new Date(d);
    return isNaN(parsed.getTime()) ? 'Just now' : parsed.toLocaleString();
  };

  return (
    <div className="cases-container" style={{ maxWidth: '1000px', margin: '0 auto' }}>
      <div className="cases-header" style={{ marginBottom: '16px' }}>
        <button className="btn btn-secondary" onClick={() => history.push('/cases')}>
          <ArrowLeft size={16} /> Back to Cases
        </button>
        <div style={{ display: 'flex', gap: '12px' }}>
          {sharePointService.getUserName() !== caseData.chargedPersonName && (
            <button className="btn btn-secondary" onClick={handleDelete} style={{color: 'var(--status-danger)', borderColor: 'rgba(220,38,38,0.2)'}}>
              Delete Incident
            </button>
          )}
          <select 
            value={caseData.status} 
            onChange={handleStatusChange}
            disabled={sharePointService.getUserName() === caseData.chargedPersonName}
            className="btn btn-primary"
            style={{ appearance: 'none', cursor: sharePointService.getUserName() === caseData.chargedPersonName ? 'not-allowed' : 'pointer', outline: 'none', opacity: sharePointService.getUserName() === caseData.chargedPersonName ? 0.7 : 1 }}
          >
            <option value="Unpaid">Status: Unpaid</option>
            <option value="Overdue">Status: Overdue</option>
            <option value="Paid">Status: Paid</option>
            <option value="Appeal Pending">Status: Appeal Pending</option>
            <option value="Waived">Status: Waived</option>
          </select>
        </div>
      </div>

      <div className="glass-panel" style={{ padding: '0', overflow: 'hidden' }}>
        {/* Header Section */}
        <div style={{ 
          padding: '32px', 
          borderBottom: '1px solid var(--border-light)',
          background: 'rgba(0,0,0,0.02)',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'flex-start'
        }}>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '8px' }}>
              <div className="kpi-icon info" style={{ width: '40px', height: '40px', borderRadius: '8px' }}>
                <ShieldAlert size={18} />
              </div>
              <h2 style={{ margin: 0, fontSize: '1.5rem' }}>{caseData.title}</h2>
              <span className={`status-badge status-${caseData.status.toLowerCase().replace(/\s+/g, '')}`}>{caseData.status}</span>
            </div>
              <p className="text-secondary" style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Clock size={14} /> Created on {getSafeDate(caseData.dateCreated)}
              </p>
            </div>

          <div style={{ textAlign: 'right' }}>
            <p className="text-secondary" style={{ margin: '0 0 4px', fontSize: '0.85rem', textTransform: 'uppercase' }}>
              {hasLateFee ? 'Total Amount Due' : 'Penalty Amount'}
            </p>
            <h3 style={{ margin: 0, fontSize: '1.8rem', color: hasLateFee ? 'var(--status-danger)' : 'var(--text-primary)' }}>
              ₦{totalDue.toLocaleString()}
            </h3>
            {hasLateFee && (
              <p style={{ margin: '4px 0 0', fontSize: '0.75rem', color: 'var(--status-warning)', fontWeight: 600 }}>
                Includes ₦{lateFeeAmount.toLocaleString()} late fee
              </p>
            )}
            <p className="text-secondary" style={{ margin: '4px 0 0', fontSize: '0.85rem' }}>
              Due: <strong style={{color: new Date(caseData.dueDate) < new Date() ? 'var(--status-danger)' : 'var(--text-primary)'}}>{getSafeDate(caseData.dueDate)}</strong>
            </p>
          </div>
        </div>

        {/* Details Section */}
        <div style={{ padding: '32px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '32px' }}>
          <div>
            <h4 style={{ color: 'var(--text-secondary)', textTransform: 'uppercase', fontSize: '0.8rem', letterSpacing: '0.05em', marginBottom: '16px' }}>
              Offender Information
            </h4>
            
            <div style={{ display: 'flex', alignItems: 'center', gap: '16px', background: 'rgba(0,0,0,0.02)', padding: '16px', borderRadius: '8px', border: '1px solid var(--border-light)' }}>
              <div className="person-avatar" style={{ width: '48px', height: '48px', fontSize: '1.2rem' }}>
                {caseData.chargedPersonName?.charAt(0) || 'U'}
              </div>
              <div>
                <h4 style={{ margin: '0 0 4px', color: 'var(--text-primary)' }}>{caseData.chargedPersonName}</h4>
                <p style={{ margin: '0 0 4px', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>{caseData.department}</p>
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--text-muted)' }}>{caseData.staffEmail}</p>
              </div>
            </div>
            
            <button className="btn btn-secondary" onClick={() => history.push(`/staff/${encodeURIComponent(caseData.chargedPerson || '')}`)} style={{ marginTop: '16px', width: '100%', justifyContent: 'center' }}>
              <User size={16} /> View Full Profile & History
            </button>
          </div>

          <div>
            <h4 style={{ color: 'var(--text-secondary)', textTransform: 'uppercase', fontSize: '0.8rem', letterSpacing: '0.05em', marginBottom: '16px' }}>
              Infraction Details
            </h4>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <div>
                <p className="text-secondary" style={{ margin: '0 0 4px', fontSize: '0.85rem' }}>Policy Category</p>
                <p style={{ margin: 0, color: 'var(--text-primary)', fontWeight: 500 }}>{caseData.offenceCategoryName}</p>
              </div>
              
              <div>
                <p className="text-secondary" style={{ margin: '0 0 4px', fontSize: '0.85rem' }}>Description of Breach</p>
                <div style={{ background: 'rgba(0,0,0,0.02)', padding: '16px', borderRadius: '8px', border: '1px solid var(--border-light)', color: 'var(--text-secondary)' }}>
                  {caseData.offenceDescription}
                </div>
              </div>

              <div>
                <p className="text-secondary" style={{ margin: '0 0 4px', fontSize: '0.85rem' }}>Issuer / Reporting Officer</p>
                <p style={{ margin: 0, color: 'var(--text-primary)' }}>{caseData.issuerName}</p>
              </div>
            </div>
          </div>
        </div>

        {/* Action Log Section */}
        <div style={{ padding: '0 32px 32px' }}>
          <h4 style={{ color: 'var(--text-secondary)', textTransform: 'uppercase', fontSize: '0.8rem', letterSpacing: '0.05em', marginBottom: '16px', borderTop: '1px solid var(--border-light)', paddingTop: '32px' }}>
            Disciplinary Actions & Escalations
          </h4>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <div style={{ display: 'flex', gap: '16px', alignItems: 'flex-start', padding: '16px', background: 'rgba(0,120,212,0.05)', border: '1px solid rgba(0,120,212,0.2)', borderRadius: '8px' }}>
              <div style={{ width: '8px', height: '8px', background: 'var(--secondary)', borderRadius: '50%', marginTop: '6px' }} />
              <div>
                <p style={{ margin: '0 0 4px', color: 'var(--text-primary)', fontWeight: 500 }}>Case Registered & Penalty Issued</p>
                <p style={{ margin: '0 0 4px', color: 'var(--text-secondary)', fontSize: '0.9rem' }}>Penalty of ₦{caseData.penaltyAmount} applied based on Tier 1 Policy Matrix.</p>
                <p style={{ margin: 0, color: 'var(--text-muted)', fontSize: '0.8rem' }}>{getSafeDateTime(caseData.dateCreated)} by System</p>
              </div>
            </div>
            
            {/* Conditional dummy escalation */}
            {caseData.id === '1' && (
              <div style={{ display: 'flex', gap: '16px', alignItems: 'flex-start', padding: '16px', background: 'rgba(233,69,96,0.05)', border: '1px solid rgba(233,69,96,0.2)', borderRadius: '8px' }}>
                <div style={{ width: '8px', height: '8px', background: 'var(--primary)', borderRadius: '50%', marginTop: '6px' }} />
                <div>
                  <p style={{ margin: '0 0 4px', color: 'var(--primary)', fontWeight: 500 }}>
                    <AlertTriangle size={14} style={{ display: 'inline', marginRight: '4px' }} /> 
                    Auto-Escalation Warning
                  </p>
                  <p style={{ margin: '0 0 4px', color: 'var(--text-secondary)', fontSize: '0.9rem' }}>This individual has 2 prior Tier 1 offences. Further breaches will result in Tier 2 escalation.</p>
                  <p style={{ margin: 0, color: 'var(--text-muted)', fontSize: '0.8rem' }}>{getSafeDateTime(caseData.dateCreated)} by PACT Escalation Engine</p>
                </div>
              </div>
            )}
          </div>
        </div>

      </div>
    </div>
  );
};
