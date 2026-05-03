import React, { useEffect, useState, useMemo } from 'react';
import { NavLink, useHistory } from 'react-router-dom';
import { ShieldAlert, ArrowLeft, Upload, AlertCircle, Loader2, Users, FileText } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import { escalationEngine } from '../../services/EscalationEngine';
import type { StaffMember, PolicyOffence, RepeatOffenceRecord } from '../../config/types';

export const NewCaseForm: React.FC = () => {
  const history = useHistory();
  const [success, setSuccess] = useState(false);

  // Dynamic Data
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [policies, setPolicies] = useState<PolicyOffence[]>([]);

  // Form State
  const [personId, setPersonId] = useState('');
  const [categoryId, setCategoryId] = useState('');
  const [description, setDescription] = useState('');
  const [dueDate, setDueDate] = useState('');
  const [tier, setTier] = useState('');
  const [disciplinaryAction, setDisciplinaryAction] = useState('');
  const [secondaryEmail, setSecondaryEmail] = useState('');
  const [staffHistory, setStaffHistory] = useState<RepeatOffenceRecord | null>(null);
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
    }).catch(() => undefined);
  }, []);

  useEffect(() => {
    if (personId) {
      sharePointService.getRepeatTrackerRecord(personId).then(record => setStaffHistory(record ?? null)).catch(() => setStaffHistory(null));
    } else {
      setStaffHistory(null);
    }
  }, [personId]);

  const handleStaffChange = (id: string): void => {
    setPersonId(id);
    const selected = staff.find(s => s.id === id);
    if (selected) {
      setSecondaryEmail(selected.lineManager || '');
    } else {
      setSecondaryEmail('');
    }
  };

  const handlePolicyChange = (id: string): void => {
    setCategoryId(id);
    const selected = policies.find(p => p.id === id);
    if (selected) {
      setTier(selected.tier || '');
      setDisciplinaryAction((selected as any).firstOffenceAction || '');
    } else {
      setTier('');
      setDisciplinaryAction('');
    }
  };

  const selectedPolicy = useMemo(() => policies.find(p => p.id === categoryId), [policies, categoryId]);
  const selectedStaff = useMemo(() => staff.find(s => s.id === personId), [staff, personId]);
  
  const offenceCount = useMemo(() => {
    if (!selectedPolicy) return 1;
    if (!staffHistory) return 1;
    if (selectedPolicy.tier === 'Tier 1') return (staffHistory.tier1Last6Months || 0) + 1;
    if (selectedPolicy.tier === 'Tier 2') return (staffHistory.tier2Offences || 0) + 1;
    if (selectedPolicy.tier === 'Tier 3') return (staffHistory.tier3Offences || 0) + 1;
    return 1;
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
    if (isEscalated) return "Automatic Escalation (Tier 2)";
    if (offenceCount === 1) return "1st Offence Path";
    if (offenceCount === 2) return "2nd Offence Path";
    if (offenceCount === 3) return "3rd Offence Path";
    return `Incident #${offenceCount}`;
  }, [isEscalated, offenceCount]);

  const riskColor = useMemo(() => {
    if (!staffHistory) return 'var(--text-secondary)';
    const level = escalationEngine.calculateRiskLevel(staffHistory);
    if (level === 'Critical' || level === 'High') return '#d13438';
    if (level === 'Medium') return '#ca5010';
    return '#107c10';
  }, [staffHistory]);

  const submitIncident = async (): Promise<void> => {
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
        secondaryContact: secondaryEmail
      });
      setIsSubmitting(false);
      setSuccess(true);
      setTimeout(() => history.push('/cases'), 4000);
    } catch (error) {
      console.error("Submission failed:", error);
      setIsSubmitting(false);
      alert(`Critical System Failure: Unable to persist compliance record. Please ensure you are logged in or reset the PACT engine.\n\nTechnical Details: ${error instanceof Error ? error.message : 'Unknown exception'}`);
    }
  };

  const handleSubmit = (e: React.FormEvent): void => {
    e.preventDefault();
    submitIncident().catch(() => undefined);
  };

  if (success) {
    return (
      <div className="glass-panel text-center" style={{ padding: '60px', maxWidth: '600px', margin: '40px auto' }}>
        <h2 style={{ color: 'var(--status-success)' }}>✔ Incident Logged Successfully</h2>
        <div style={{ marginTop: '24px', padding: '20px', background: 'rgba(16, 185, 129, 0.05)', borderRadius: '12px', border: '1px solid rgba(16, 185, 129, 0.2)' }}>
          <p style={{ fontWeight: 500 }}>The PACT engine has recorded this case and triggered formal notifications.</p>
        </div>
        <p className="text-secondary" style={{ marginTop: '24px', fontSize: '0.85rem' }}>
          Redirecting to case directory...
        </p>
      </div>
    );
  }

  return (
    <div className="cases-container" style={{ maxWidth: '900px', margin: '0 auto' }} data-testid="new-case-form">
      <div className="cases-header" style={{ marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <NavLink to="/cases" className="btn btn-secondary">
          <ArrowLeft size={16} /> Case Directory
        </NavLink>
        <div style={{ padding: '6px 12px', background: 'rgba(255,255,255,0.05)', borderRadius: '20px', fontSize: '0.75rem', fontWeight: 600, color: 'var(--text-secondary)', border: '1px solid var(--border-light)' }}>
           v5.0 — POLICY COMPLIANCE
        </div>
      </div>

      <div className="form-layout-vertical" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
        
        {/* Section 1: Staff Identity */}
        <div className="glass-panel" style={{ padding: '24px' }}>
          <h3 style={{ marginBottom: '20px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '1rem', color: 'var(--text-primary)' }}>
             <Users size={18} style={{ color: 'var(--primary)' }} />
             Staff Identity
          </h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div className="form-group">
              <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.85rem', fontWeight: 600, textTransform: 'uppercase' }}>Target Employee</label>
              <select 
                value={personId}
                onChange={e => handleStaffChange(e.target.value)}
                required
                className="form-input"
                style={{ width: '100%', padding: '14px', background: 'var(--bg-surface-hover)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none', fontSize: '1rem' }}
              >
                <option value="">-- Select Employee to Charge --</option>
                {staff.map(member => (
                  <option key={member.id} value={member.id}>{member.fullName} — {member.department}</option>
                ))}
              </select>
            </div>

            {selectedStaff && (
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px', padding: '16px', background: 'rgba(255,255,255,0.02)', borderRadius: '8px', border: '1px solid var(--border-light)' }}>
                <div>
                  <div className="text-secondary" style={{ fontSize: '0.7rem', fontWeight: 600, textTransform: 'uppercase' }}>Secondary Contact (CC)</div>
                  <div style={{ padding: '8px 0', color: 'var(--text-muted)', fontSize: '0.9rem' }}>{secondaryEmail || 'Not Assigned'}</div>
                </div>
                <div>
                  <div className="text-secondary" style={{ fontSize: '0.7rem', fontWeight: 600, textTransform: 'uppercase' }}>Employee Email</div>
                  <div style={{ padding: '8px 0', color: 'var(--text-primary)', fontWeight: 500 }}>{selectedStaff.email}</div>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Section 2: Policy Matrix */}
        <div className="glass-panel" style={{ padding: '24px' }}>
          <h3 style={{ marginBottom: '20px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '1rem', color: 'var(--text-primary)' }}>
             <ShieldAlert size={18} style={{ color: 'var(--status-warning)' }} />
             Policy Matrix Enforcement
          </h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div className="form-group">
              <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.85rem', fontWeight: 600, textTransform: 'uppercase' }}>Select Infraction</label>
              <select 
                value={categoryId}
                onChange={e => handlePolicyChange(e.target.value)}
                required
                className="form-input"
                style={{ width: '100%', padding: '14px', background: 'var(--bg-surface-hover)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none', fontSize: '1rem' }}
              >
                <option value="">-- Select Categorized Offence --</option>
                {policies.map(policy => (
                  <option key={policy.id} value={policy.id}>{policy.offenceName} ({policy.tier})</option>
                ))}
              </select>
            </div>

            {selectedPolicy && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '16px' }}>
                  <div style={{ padding: '16px', background: 'rgba(0,0,0,0.1)', borderRadius: '8px', border: '1px solid var(--border-light)' }}>
                    <div className="text-secondary" style={{ fontSize: '0.7rem', fontWeight: 600, textTransform: 'uppercase' }}>Classification</div>
                    <div style={{ fontSize: '1.25rem', fontWeight: 700, color: tier.includes('3') ? 'var(--status-danger)' : 'var(--text-primary)', marginTop: '4px' }}>
                       {tier}
                    </div>
                  </div>
                  <div style={{ padding: '16px', background: 'rgba(0,120,212,0.1)', borderRadius: '8px', border: '1px solid rgba(0,120,212,0.2)' }}>
                    <div className="text-secondary" style={{ fontSize: '0.7rem', fontWeight: 600, textTransform: 'uppercase' }}>Recommended Sanction</div>
                    <div style={{ fontSize: '1.1rem', fontWeight: 600, color: 'var(--text-primary)', marginTop: '4px' }}>
                       {recommendedAction || disciplinaryAction || (tier.includes('3') ? 'Suspension / Termination' : 'Pending Calculation')}
                    </div>
                  </div>
                </div>

                <div style={{ padding: '16px', background: 'rgba(255,255,255,0.03)', borderRadius: '8px', border: '1px solid var(--border-light)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}>
                    <span className="text-secondary" style={{ fontSize: '0.75rem', fontWeight: 600 }}>GOVERNANCE PATH</span>
                    <span style={{ fontSize: '0.75rem', fontWeight: 700, color: isEscalated ? 'var(--status-danger)' : 'var(--status-success)' }}>
                      {actionLabel}
                    </span>
                  </div>
                  <div style={{ display: 'flex', gap: '20px', fontSize: '0.85rem', color: 'var(--text-primary)', opacity: 0.8 }}>
                    <span>Rolling T1 Count: <strong>{staffHistory?.tier1Last6Months || 0}</strong></span>
                    <span>Lifetime T2: <strong>{staffHistory?.tier2Offences || 0}</strong></span>
                    <span>Risk Level: <strong style={{ color: riskColor }}>{staffHistory ? escalationEngine.calculateRiskLevel(staffHistory) : 'Low'}</strong></span>
                  </div>
                  {isEscalated && (
                    <div style={{ marginTop: '12px', padding: '10px', background: 'rgba(209, 52, 56, 0.1)', borderLeft: '3px solid var(--status-danger)', borderRadius: '4px', fontSize: '0.8rem', color: 'var(--status-danger)', fontWeight: 600 }}>
                       POLICY ALERT: This incident has been automatically upgraded to Tier 2 based on repeat-offence thresholds.
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Section 3: Narrative */}
        <div className="glass-panel" style={{ padding: '24px' }}>
          <h3 style={{ marginBottom: '20px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '1rem', color: 'var(--text-primary)' }}>
             <FileText size={18} style={{ color: 'var(--secondary)' }} />
             Incident Narrative
          </h3>
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div className="form-group">
              <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.85rem', fontWeight: 600, textTransform: 'uppercase' }}>Description of Breach</label>
              <textarea 
                value={description}
                onChange={e => setDescription(e.target.value)}
                required
                rows={4}
                placeholder="Provide objective details of the infraction, including location, time, and specific policy violated..."
                style={{ width: '100%', padding: '14px', background: 'var(--bg-surface-hover)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none', resize: 'vertical' }}
              />
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
              <div className="form-group">
                <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.85rem', fontWeight: 600, textTransform: 'uppercase' }}>Incident Date</label>
                <input 
                  type="date"
                  value={dueDate}
                  onChange={e => setDueDate(e.target.value)}
                  required
                  style={{ width: '100%', padding: '14px', background: 'var(--bg-surface-hover)', border: '1px solid var(--border-light)', borderRadius: '8px', color: 'var(--text-primary)', outline: 'none' }}
                />
              </div>
              <div className="form-group">
                <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.85rem', fontWeight: 600, textTransform: 'uppercase' }}>Evidence Attachment</label>
                <div style={{ padding: '12px', background: 'rgba(255,255,255,0.05)', border: '1px dashed var(--border-light)', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', cursor: 'pointer', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>
                  <Upload size={16} /> Click to Upload Files
                </div>
              </div>
            </div>

            <div style={{ marginTop: '24px', paddingTop: '24px', borderTop: '1px solid var(--border-light)', display: 'flex', justifyContent: 'flex-end', gap: '16px' }}>
              <button 
                type="button" 
                className="btn btn-secondary"
                onClick={() => setShowPreview(true)}
                disabled={!personId || !categoryId || isSubmitting}
              >
                Preview Notice
              </button>
              <button 
                type="submit" 
                className="btn btn-primary" 
                disabled={isSubmitting} 
                style={{ minWidth: '200px', height: '48px' }}
              >
                {isSubmitting ? <><Loader2 size={18} className="animate-spin" /> Processing...</> : 'Log Compliance Case'}
              </button>
            </div>
          </form>
        </div>
      </div>

      {showPreview && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.85)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}>
          <div className="glass-panel" style={{ maxWidth: '650px', width: '100%', maxHeight: '90vh', overflow: 'auto', padding: '0', border: '1px solid #444' }}>
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
                    <tbody>
                      <tr><td style={{ color: '#666', width: '100px' }}>Offence:</td><td>{selectedPolicy?.offenceName}</td></tr>
                      <tr><td style={{ color: '#666', padding: '10px 0' }}>Description:</td><td style={{ padding: '10px 0' }}>{description || '[No description provided]'}</td></tr>
                      <tr><td style={{ color: '#666' }}>Classification:</td><td><b>{actionLabel}</b></td></tr>
                    </tbody>
                  </table>
                </div>
                <div style={{ borderLeft: '4px solid #0078d4', paddingLeft: '15px', marginTop: '20px' }}>
                  <h3 style={{ margin: 0, fontSize: '16px', color: '#0078d4' }}>Disciplinary Action</h3>
                  <p style={{ fontSize: '18px', fontWeight: 'bold', margin: '10px 0' }}>{recommendedAction}</p>
                </div>
              </div>
            </div>
            <div style={{ padding: '20px', textAlign: 'right', display: 'flex', gap: '10px', justifyContent: 'flex-end', background: 'rgba(255,255,255,0.05)' }}>
              <button type="button" className="btn btn-secondary" onClick={() => setShowPreview(false)}>Close</button>
              <button type="button" className="btn btn-primary" onClick={() => { setShowPreview(false); submitIncident().catch(() => undefined); }}>Confirm & Log</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
