import React, { useEffect, useState, useMemo } from 'react';
import { NavLink, useNavigate } from 'react-router-dom';
import { ShieldAlert, ArrowLeft, Upload, AlertCircle, Loader2 } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import { escalationEngine } from '../../services/EscalationEngine';
import type { StaffMember, PolicyOffence, RepeatOffenceRecord } from '../../config/types';

export const NewCaseForm: React.FC = () => {
  const navigate = useNavigate();
  const [success, setSuccess] = useState(false);

  // Dynamic Data
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [policies, setPolicies] = useState<PolicyOffence[]>([]);

  // Form State
  const [personId, setPersonId] = useState('');
  const [categoryId, setCategoryId] = useState('');
  const [description, setDescription] = useState('');
  const [dueDate, setDueDate] = useState('');
  const [staffHistory, setStaffHistory] = useState<RepeatOffenceRecord | null>(null);
  const [emailNotification, setEmailNotification] = useState<{to: string, subject: string} | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);

  useEffect(() => {
    // Load dropdown options
    Promise.all([
      sharePointService.getStaffDirectory(),
      sharePointService.getPolicyLibrary()
    ]).then(([staffData, policyData]) => {
      setStaff(staffData);
      setPolicies(policyData);
    });
  }, []);

  useEffect(() => {
    if (personId) {
      sharePointService.getRepeatTrackerRecord(personId).then(setStaffHistory);
    } else {
      setStaffHistory(null);
    }
  }, [personId]);

  useEffect(() => {
    const handleEmail = (e: any) => {
      setEmailNotification({ 
        to: e.detail.to.join(', '), 
        subject: e.detail.subject 
      });
      setTimeout(() => setEmailNotification(null), 5000); // Clear after 5s
    };
    window.addEventListener('pact-mock-email', handleEmail);
    return () => window.removeEventListener('pact-mock-email', handleEmail);
  }, []);

  const selectedPolicy = useMemo(() => policies.find(p => p.id === categoryId), [policies, categoryId]);
  const selectedStaff = useMemo(() => staff.find(s => s.id === personId), [staff, personId]);
  
  // Calculate dynamic action based on history
  const offenceCount = useMemo(() => {
    if (!selectedPolicy) return 1;
    return selectedPolicy.tier === 'Tier 1' 
      ? (staffHistory?.tier1Last6Months || 0) + 1 
      : 1;
  }, [selectedPolicy, staffHistory]);
  
  const isEscalated = useMemo(() => {
    if (!staffHistory) return false;
    return escalationEngine.checkEscalation(new Date().toISOString(), staffHistory);
  }, [staffHistory]);

  const recommendedAction = useMemo(() => {
    if (!selectedPolicy) return '';
    return escalationEngine.getRecommendedAction(selectedPolicy, offenceCount, isEscalated);
  }, [selectedPolicy, offenceCount, isEscalated]);

  const actionLabel = useMemo(() => {
    if (isEscalated) return "Escalated Action";
    if (offenceCount === 1) return "1st Offence Path";
    if (offenceCount === 2) return "2nd Offence Path";
    return "3rd Offence Path";
  }, [isEscalated, offenceCount]);
  const riskColor = useMemo(() => {
    if (!staffHistory) return 'var(--text-secondary)';
    const level = escalationEngine.calculateRiskLevel(staffHistory);
    if (level === 'Critical' || level === 'High') return '#d13438';
    if (level === 'Medium') return '#ca5010';
    return '#107c10';
  }, [staffHistory]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!personId || !categoryId) {
      alert("Please select a staff member and an infraction.");
      return;
    }
    
    setIsSubmitting(true);
    try {
      await sharePointService.createCase({
        chargedPerson: personId,
        offenceCategory: categoryId,
        offenceDescription: description,
        dueDate: dueDate || new Date(Date.now() + 7 * 86400000).toISOString(),
      });
      setIsSubmitting(false);
      setSuccess(true);
      setTimeout(() => navigate('/cases'), 4000);
    } catch (error) {
      console.error("Submission failed:", error);
      setIsSubmitting(false);
      alert("Failed to log incident. Please check your connection or reset the engine.");
    }
  };

  if (success) {
    return (
      <div className="glass-panel text-center" style={{ padding: '60px', maxWidth: '600px', margin: '40px auto' }}>
        <h2 style={{ color: 'var(--primary)' }}>✔ Incident Logged Successfully</h2>
        <div style={{ marginTop: '24px', padding: '20px', background: 'rgba(0,0,0,0.02)', borderRadius: '12px' }}>
          <p>The PACT engine has recorded this case and triggered notifications.</p>
        </div>
        <p className="text-secondary" style={{ marginTop: '24px', fontSize: '0.85rem' }}>
          Redirecting to case directory...
        </p>
      </div>
    );
  }

  return (
    <div className="cases-container" style={{ maxWidth: '800px', margin: '0 auto' }}>
      <div className="cases-header" style={{ marginBottom: '16px' }}>
        <NavLink to="/cases" className="btn btn-secondary">
          <ArrowLeft size={16} /> Back to Cases
        </NavLink>
      </div>

      <div className="glass-panel" style={{ padding: '32px' }} data-testid="new-case-form">
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '32px' }}>
          <div className="kpi-icon primary" style={{ width: '48px', height: '48px' }}>
            <ShieldAlert size={20} />
          </div>
          <div>
            <h2 style={{ margin: 0 }}>Log Compliance Incident</h2>
            <p className="text-secondary" style={{ margin: '4px 0 0', fontSize: '0.9rem' }}>
              Create a new case using the <b>Policy Matrix</b>. Notifications are sent automatically.
            </p>
          </div>
        </div>

        <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '24px' }} data-testid="new-case-submit-form">
          
          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem' }}>Offender</label>
            <select 
              value={personId}
              onChange={e => setPersonId(e.target.value)}
              required
              className="form-input"
              data-testid="case-offender-select"
              style={{ width: '100%', padding: '12px', background: 'rgba(0,0,0,0.02)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none' }}
            >
              <option value="">-- Select Offender --</option>
              {staff.map(member => (
                <option key={member.id} value={member.id}>{member.fullName} ({member.department})</option>
              ))}
            </select>
          </div>

          {selectedStaff && (
            <div style={{ padding: '16px', background: 'rgba(255,255,255,0.03)', border: '1px solid var(--border-light)', borderRadius: '8px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px', animation: 'fadeIn 0.2s ease' }}>
              <div>
                <div className="text-secondary" style={{ fontSize: '0.75rem', textTransform: 'uppercase' }}>Department</div>
                <div style={{ color: 'var(--text-primary)', fontWeight: 500 }}>{selectedStaff.department}</div>
              </div>
              <div>
                <div className="text-secondary" style={{ fontSize: '0.75rem', textTransform: 'uppercase' }}>Reporting To</div>
                <div style={{ color: 'var(--secondary)', fontWeight: 600 }}>{selectedStaff.lineManager || 'Not Assigned'}</div>
              </div>
            </div>
          )}

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem' }}>Infraction (Policy Engine)</label>
            <select 
              value={categoryId}
              onChange={e => setCategoryId(e.target.value)}
              required
              className="form-input"
              data-testid="case-policy-select"
              style={{ width: '100%', padding: '12px', background: 'rgba(0,0,0,0.02)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none' }}
            >
              <option value="">-- Select Infraction --</option>
              {policies.map(policy => (
                <option key={policy.id} value={policy.id}>{policy.offenceName} ({policy.tier})</option>
              ))}
            </select>
          </div>

          {selectedPolicy && (
            <div style={{ padding: '16px', background: 'rgba(0,120,212,0.05)', border: '1px solid rgba(0,120,212,0.2)', borderRadius: '8px', animation: 'fadeIn 0.3s ease' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
                <div>
                  <div className="text-secondary" style={{ fontSize: '0.75rem', textTransform: 'uppercase' }}>Disciplinary Action</div>
                  <div style={{ fontSize: '1.1rem', fontWeight: 600, color: 'var(--text-primary)', marginTop: '4px' }}>
                    {recommendedAction || 'Calculating...'}
                  </div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <div className="text-secondary" style={{ fontSize: '0.75rem', textTransform: 'uppercase' }}>Classification</div>
                  <span className={`status-badge status-${selectedPolicy.tier.toLowerCase().replace(' ', '')}`}>{selectedPolicy.tier}</span>
                </div>
              </div>
              
              <div style={{ background: 'rgba(0,0,0,0.05)', padding: '12px', borderRadius: '4px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px' }}>
                  <div className="text-secondary" style={{ fontSize: '0.75rem', textTransform: 'uppercase' }}>Incident Context</div>
                  <div style={{ fontSize: '0.75rem', fontWeight: 600, color: riskColor }}>
                    Risk: {staffHistory ? escalationEngine.calculateRiskLevel(staffHistory) : 'Low'}
                  </div>
                </div>
                
                <div style={{ fontSize: '0.9rem', color: 'var(--text-primary)', marginBottom: '10px' }}>
                  <div style={{ display: 'flex', gap: '12px', marginBottom: '8px', opacity: 0.8 }}>
                    <span title="Tier 1 offences in last 6 months">T1: <b>{staffHistory?.tier1Last6Months || 0}</b></span>
                    <span title="Total Tier 2 offences">T2: <b>{staffHistory?.tier2Offences || 0}</b></span>
                    <span>Status: <b>{actionLabel}</b></span>
                  </div>
                  {isEscalated && (
                    <div style={{ color: '#d13438', fontWeight: 600, fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '4px' }}>
                      <AlertCircle size={14} /> Automatic Policy Escalation Triggered
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem' }}>Infraction Description</label>
            <textarea 
              value={description}
              onChange={e => setDescription(e.target.value)}
              required
              rows={4}
              placeholder="Describe the specific details of the compliance breach..."
              data-testid="case-description"
              style={{ width: '100%', padding: '12px', background: 'rgba(0,0,0,0.02)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none', resize: 'vertical' }}
            />
          </div>

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem' }}>Date</label>
            <input 
              type="date"
              value={dueDate}
              onChange={e => setDueDate(e.target.value)}
              required
              data-testid="case-due-date"
              style={{ width: '100%', padding: '12px', background: 'rgba(0,0,0,0.02)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none' }}
            />
          </div>

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem' }}>Evidence Attachment</label>
            <div 
              style={{ 
                border: '2px dashed var(--border-light)', 
                borderRadius: '8px', 
                padding: '32px', 
                textAlign: 'center',
                background: 'rgba(0,0,0,0.02)',
                cursor: 'pointer'
              }}
            >
              <Upload size={24} className="text-secondary" style={{ margin: '0 auto 12px' }} />
              <div style={{color: 'var(--text-primary)'}}>Click or drag file to upload</div>
              <div className="text-muted" style={{ fontSize: '0.8rem', marginTop: '4px' }}>PNG, JPG, PDF up to 10MB</div>
            </div>
          </div>

          <div style={{ padding: '16px', background: 'rgba(0,0,0,0.02)', border: '1px solid var(--border-light)', borderRadius: '8px', marginTop: '8px' }}>
            <div className="text-secondary" style={{ fontSize: '0.75rem', textTransform: 'uppercase', marginBottom: '8px' }}>Notification Preview</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div style={{ fontSize: '0.85rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <ShieldAlert size={14} style={{ color: 'var(--status-success)' }} />
                <span>To Offender: <code style={{ color: 'var(--primary)' }}>{selectedStaff?.email || '...'}</code></span>
              </div>
              {selectedStaff?.lineManager && (() => {
                const manager = staff.find(s => s.email === selectedStaff.lineManager);
                return (
                <div style={{ fontSize: '0.85rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <ShieldAlert size={14} style={{ color: 'var(--secondary)' }} />
                  <span>CC Supervisor: <strong>{manager?.fullName || 'Unknown'}</strong> — <code style={{ color: 'var(--secondary)' }}>{manager?.email || selectedStaff.lineManager}</code></span>
                </div>
                );
              })()}

            </div>
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '16px', marginTop: '16px', paddingTop: '24px', borderTop: '1px solid var(--border-light)' }}>
            <NavLink to="/cases" className="btn btn-secondary">Cancel</NavLink>
            <button 
              type="button" 
              className="btn btn-secondary"
              onClick={() => setShowPreview(true)}
              disabled={!personId || !categoryId || isSubmitting}
              data-testid="case-preview-button"
            >
              Preview Mail
            </button>
            <button type="submit" className="btn btn-primary" disabled={isSubmitting} data-testid="case-submit-button" style={{ display: 'flex', alignItems: 'center', gap: '8px', minWidth: '160px', justifyContent: 'center' }}>
              {isSubmitting ? (
                <>
                  <Loader2 size={18} className="animate-spin" />
                  Processing...
                </>
              ) : 'Submit Incident'}
            </button>
          </div>

          {showPreview && (
            <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}>
              <div className="glass-panel" style={{ maxWidth: '650px', width: '100%', maxHeight: '90vh', overflow: 'auto', padding: '0', border: '1px solid #444', background: 'var(--bg-panel)' }}>
                <div style={{ padding: '20px', borderBottom: '1px solid var(--border-light)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <h3 style={{ margin: 0 }}>Notification Preview</h3>
                  <button type="button" onClick={() => setShowPreview(false)} style={{ background: 'none', border: 'none', color: 'white', cursor: 'pointer', fontSize: '1.2rem' }}>&times;</button>
                </div>
                <div style={{ padding: '30px', background: 'white', color: '#333' }}>
                  <div style={{ marginBottom: '20px', fontSize: '0.9rem', color: '#666', borderBottom: '1px solid #eee', paddingBottom: '10px' }}>
                    <b>Subject:</b> PACT ALERT: {selectedPolicy?.offenceName} - {selectedStaff?.fullName}
                  </div>
                  <div style={{ fontFamily: 'Arial, sans-serif' }}>
                    <h2 style={{ color: '#0078d4', marginTop: 0 }}>PACT Compliance Notification</h2>
                    <p>Formal record for <b>{selectedStaff?.fullName}</b>.</p>
                    <div style={{ background: '#f3f2f1', padding: '15px', borderRadius: '4px', margin: '20px 0' }}>
                      <table style={{ width: '100%', fontSize: '14px' }}>
                        <tr><td style={{ color: '#666', width: '100px' }}>Offence:</td><td>{selectedPolicy?.offenceName}</td></tr>
                        <tr><td style={{ color: '#666', padding: '10px 0' }}>Description:</td><td style={{ padding: '10px 0' }}>{description || '[No description provided]'}</td></tr>
                        <tr><td style={{ color: '#666' }}>Classification:</td><td><b>{actionLabel}</b></td></tr>
                      </table>
                    </div>
                    <div style={{ borderLeft: '4px solid #0078d4', paddingLeft: '15px', marginTop: '20px' }}>
                      <h3 style={{ margin: 0, fontSize: '16px', color: '#0078d4' }}>Disciplinary Action</h3>
                      <p style={{ fontSize: '18px', fontWeight: 'bold', margin: '10px 0' }}>{recommendedAction}</p>
                    </div>
                  </div>
                </div>
                <div style={{ padding: '20px', textAlign: 'right', display: 'flex', gap: '10px', justifyContent: 'flex-end', background: 'rgba(255,255,255,0.05)' }}>
                  <button type="button" className="btn btn-secondary" onClick={() => setShowPreview(false)} data-testid="case-preview-close">Close Preview</button>
                  <button type="button" className="btn btn-primary" onClick={(e) => { setShowPreview(false); handleSubmit(e as any); }} data-testid="case-preview-confirm">Confirm & Log Now</button>
                </div>
              </div>
            </div>
          )}
        </form>

        {/* Email Mock Toast */}
        {emailNotification && (
          <div style={{ 
            position: 'fixed', 
            bottom: '24px', 
            right: '24px', 
            background: 'var(--primary)', 
            color: 'white', 
            padding: '16px 24px', 
            borderRadius: '12px', 
            boxShadow: '0 8px 32px rgba(0,0,0,0.3)',
            display: 'flex',
            flexDirection: 'column',
            gap: '4px',
            animation: 'slideInRight 0.3s ease',
            zIndex: 1000
          }}>
            <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>📬 Mock Email Sent (Local Mode)</div>
            <div style={{ fontSize: '0.8rem', opacity: 0.9 }}>To: {emailNotification.to}</div>
          </div>
        )}
      </div>
    </div>
  );
};
