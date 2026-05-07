import React from 'react';
import { FileSearch, Filter, Plus, X, CheckCircle, AlertCircle } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import { AppealForm } from './AppealForm';
import { APPEAL_SLA_DAYS } from '../../config/constants';

interface ReviewModal {
  appeal: any;
  decision: string;
  reviewingOfficer: string;
  decisionNotes: string;
}

export const AppealsPage: React.FC = () => {
  const [appeals, setAppeals] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [showForm, setShowForm] = React.useState(false);
  const [reviewModal, setReviewModal] = React.useState<ReviewModal | null>(null);
  const [saving, setSaving] = React.useState(false);
  const [saveSuccess, setSaveSuccess] = React.useState(false);

  const fetchAppeals = React.useCallback(async () => {
    setLoading(true);
    const data = await sharePointService.getAppeals();
    setAppeals(data);
    setLoading(false);
  }, []);

  React.useEffect(() => {
    fetchAppeals();
  }, [fetchAppeals]);

  const openReview = (appeal: any) => {
    setReviewModal({
      appeal,
      decision: appeal.decision === 'Pending' ? '' : appeal.decision,
      reviewingOfficer: appeal.reviewingOfficer || '',
      decisionNotes: appeal.decisionNotes || '',
    });
    setSaveSuccess(false);
  };

  const handleSaveDecision = async () => {
    if (!reviewModal || !reviewModal.decision) return;
    setSaving(true);
    try {
      await sharePointService.updateAppeal(reviewModal.appeal.id, {
        decision: reviewModal.decision,
        reviewingOfficer: reviewModal.reviewingOfficer,
        decisionNotes: reviewModal.decisionNotes,
      });
      setSaveSuccess(true);
      setTimeout(() => {
        setReviewModal(null);
        setSaveSuccess(false);
        fetchAppeals();
      }, 1500);
    } catch (err) {
      console.error('Failed to save decision:', err);
      alert('Failed to save decision. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  if (showForm) {
    return (
      <AppealForm 
        onClose={() => setShowForm(false)} 
        onSuccess={() => {
          setShowForm(false);
          fetchAppeals();
        }} 
      />
    );
  }

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px', borderRadius: 'var(--radii-md)'}}></div>
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}}></div>
      </div>
    );
  }

  const getDecisionBadgeClass = (decision: string) => {
    switch (decision) {
      case 'Upheld': return 'status-paid';
      case 'Reduced': return 'status-overdue';
      case 'Waived': return 'status-paid';
      case 'Rejected': return 'status-unpaid';
      default: return 'status-unpaid';
    }
  };

  return (
    <div className="cases-container fade-in">
      <div className="cases-header">
        <h2 style={{margin: 0, display: 'flex', alignItems: 'center', gap: '8px'}}><FileSearch size={20} color="var(--secondary)"/> Appeals Register</h2>
        <div className="cases-actions">
          <button className="btn btn-secondary"><Filter size={16} /> Filter</button>
          <button className="btn btn-primary" onClick={() => setShowForm(true)}>
            <Plus size={16} /> New Appeal
          </button>
        </div>
      </div>

      <div className="cases-table-container glass-panel">
        <table className="pact-table">
            <thead>
              <tr>
                <th>Case Ref</th>
                <th>Appellant</th>
                <th>Submission Date</th>
                <th>Reviewing Officer</th>
                <th>Decision Status</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {appeals.map(appeal => (
                <tr key={appeal.id}>
                  <td className="link-action" style={{cursor:'pointer'}}>{appeal.caseReference}</td>
                  <td>{appeal.appellant}</td>
                  <td>{new Date(appeal.appealDate).toLocaleDateString()}</td>
                  <td>{appeal.reviewingOfficer || 'Pending Assignment'}</td>
                  <td>
                    <span className={`status-badge ${getDecisionBadgeClass(appeal.decision)}`}>
                      {appeal.decision}
                    </span>
                  </td>
                  <td>
                    <span 
                      className="link-action" 
                      style={{cursor:'pointer', color: 'var(--primary)', fontWeight: 500}}
                      onClick={() => openReview(appeal)}
                    >
                      {appeal.decision === 'Pending' ? 'Review' : 'View Decision'}
                    </span>
                  </td>
                </tr>
              ))}
              {appeals.length === 0 && (
                <tr>
                  <td colSpan={6} className="text-center">No appeals found.</td>
                </tr>
              )}
            </tbody>
          </table>
      </div>

      {/* ─── Appeal Decision Modal ─── */}
      {reviewModal && (
        <div style={{
          position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
          background: 'rgba(0,0,0,0.75)', zIndex: 1000,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          padding: '20px', animation: 'fadeIn 0.2s ease'
        }}>
          <div className="glass-panel" style={{
            maxWidth: '560px', width: '100%', maxHeight: '90vh', overflow: 'auto',
            padding: '0', border: '1px solid rgba(255,255,255,0.1)',
            background: 'var(--bg-panel)', borderRadius: '16px'
          }}>
            {/* Modal Header */}
            <div style={{
              padding: '24px', borderBottom: '1px solid var(--border-light)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center'
            }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '1.1rem' }}>Review Appeal — {reviewModal.appeal.caseReference}</h3>
                <p className="text-secondary" style={{ margin: '4px 0 0', fontSize: '0.85rem' }}>
                  Filed by {reviewModal.appeal.appellant} on {new Date(reviewModal.appeal.appealDate).toLocaleDateString()}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setReviewModal(null)}
                style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Body */}
            {saveSuccess ? (
              <div style={{ padding: '48px 32px', textAlign: 'center' }}>
                <CheckCircle size={48} color="var(--status-success)" style={{ marginBottom: '16px' }} />
                <h3 style={{ color: 'var(--status-success)', margin: '0 0 8px' }}>Decision Recorded</h3>
                <p className="text-secondary" style={{ fontSize: '0.9rem' }}>
                  The appellant has been notified via email.
                </p>
              </div>
            ) : (
              <div style={{ padding: '24px' }}>
                {/* Grounds Display */}
                <div style={{ marginBottom: '20px' }}>
                  <label className="text-secondary" style={{ display: 'block', fontSize: '0.8rem', textTransform: 'uppercase', marginBottom: '6px', letterSpacing: '0.04em' }}>
                    Grounds for Appeal
                  </label>
                  <div style={{
                    padding: '14px', background: 'rgba(0,0,0,0.1)', borderRadius: '8px',
                    border: '1px solid var(--border-light)', fontSize: '0.9rem',
                    color: 'var(--text-secondary)', lineHeight: 1.6, maxHeight: '120px', overflow: 'auto'
                  }}>
                    {reviewModal.appeal.grounds || 'No grounds provided.'}
                  </div>
                </div>

                {/* Decision Select */}
                <div style={{ marginBottom: '20px' }}>
                  <label className="text-secondary" style={{ display: 'block', fontSize: '0.8rem', textTransform: 'uppercase', marginBottom: '6px', letterSpacing: '0.04em' }}>
                    Decision <span style={{ color: 'var(--primary)' }}>*</span>
                  </label>
                  <select
                    value={reviewModal.decision}
                    onChange={e => setReviewModal({ ...reviewModal, decision: e.target.value })}
                    style={{
                      width: '100%', padding: '12px', background: 'rgba(0,0,0,0.15)',
                      border: '1px solid var(--border-light)', borderRadius: '8px',
                      color: 'var(--text-primary)', outline: 'none', fontSize: '0.9rem'
                    }}
                  >
                    <option value="">-- Select Decision --</option>
                    <option value="Upheld">Upheld — Penalty Cancelled</option>
                    <option value="Reduced">Reduced — Penalty Lowered</option>
                    <option value="Waived">Waived — No Penalty Applied</option>
                    <option value="Rejected">Rejected — Original Penalty Stands</option>
                  </select>
                </div>

                {/* Reviewing Officer */}
                <div style={{ marginBottom: '20px' }}>
                  <label className="text-secondary" style={{ display: 'block', fontSize: '0.8rem', textTransform: 'uppercase', marginBottom: '6px', letterSpacing: '0.04em' }}>
                    Reviewing Officer
                  </label>
                  <input
                    type="text"
                    value={reviewModal.reviewingOfficer}
                    onChange={e => setReviewModal({ ...reviewModal, reviewingOfficer: e.target.value })}
                    placeholder="Name of reviewing officer"
                    style={{
                      width: '100%', padding: '12px', background: 'rgba(0,0,0,0.15)',
                      border: '1px solid var(--border-light)', borderRadius: '8px',
                      color: 'var(--text-primary)', outline: 'none', fontSize: '0.9rem',
                      boxSizing: 'border-box'
                    }}
                  />
                </div>

                {/* Decision Notes */}
                <div style={{ marginBottom: '24px' }}>
                  <label className="text-secondary" style={{ display: 'block', fontSize: '0.8rem', textTransform: 'uppercase', marginBottom: '6px', letterSpacing: '0.04em' }}>
                    Decision Notes
                  </label>
                  <textarea
                    value={reviewModal.decisionNotes}
                    onChange={e => setReviewModal({ ...reviewModal, decisionNotes: e.target.value })}
                    placeholder="Provide rationale for this decision..."
                    rows={3}
                    style={{
                      width: '100%', padding: '12px', background: 'rgba(0,0,0,0.15)',
                      border: '1px solid var(--border-light)', borderRadius: '8px',
                      color: 'var(--text-primary)', outline: 'none', fontSize: '0.9rem',
                      resize: 'vertical', fontFamily: 'inherit', boxSizing: 'border-box'
                    }}
                  />
                </div>

                {/* SLA Reminder */}
                <div style={{
                  padding: '12px 16px', background: 'rgba(245,158,11,0.06)',
                  border: '1px solid rgba(245,158,11,0.15)', borderRadius: '8px',
                  fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '24px'
                }}>
                  <AlertCircle size={14} style={{ color: '#f59e0b', display: 'inline', marginRight: '6px', verticalAlign: 'middle' }} />
                  Appeal SLA: Respond within <strong style={{ color: '#f59e0b' }}>{APPEAL_SLA_DAYS} working days</strong> of submission.
                </div>

                {/* Actions */}
                <div style={{ display: 'flex', gap: '12px' }}>
                  <button
                    className="btn btn-secondary"
                    onClick={() => setReviewModal(null)}
                    style={{ flex: 1, justifyContent: 'center' }}
                  >
                    Cancel
                  </button>
                  <button
                    className="btn btn-primary"
                    onClick={handleSaveDecision}
                    disabled={!reviewModal.decision || saving}
                    style={{ flex: 1, justifyContent: 'center' }}
                  >
                    {saving ? 'Saving...' : 'Record Decision'}
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
