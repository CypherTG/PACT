import React, { useEffect, useState } from 'react';
import { useParams, NavLink } from 'react-router-dom';
import {
  Mail, Building2, Briefcase,
  ShieldAlert,
  ArrowLeft, History, TrendingUp
} from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import { escalationEngine } from '../../services/EscalationEngine';
import type { StaffMember, ComplianceCase, RepeatOffenceRecord, DisciplinaryAction } from '../../config/types';

export const StaffProfile: React.FC = () => {
  const { id } = useParams<{ id: string }>();
  const [member, setMember] = useState<StaffMember | null>(null);
  const [cases, setCases] = useState<ComplianceCase[]>([]);
  const [actions, setActions] = useState<DisciplinaryAction[]>([]);
  const [tracker, setTracker] = useState<RepeatOffenceRecord | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!id) return;

    const loadProfile = async (): Promise<void> => {
      setLoading(true);
      const [staffList, caseList, actionList, trackerRecord] = await Promise.all([
        sharePointService.getStaffDirectory(),
        sharePointService.getCases(),
        sharePointService.getDisciplinaryActions(),
        sharePointService.getRepeatTrackerRecord(id || '')
      ]);

      const memberRecord = Array.isArray(staffList) ? staffList.find((s: StaffMember) => s.id === id) || null : null;
      const personCases = Array.isArray(caseList) ? caseList.filter((c: ComplianceCase) => c.chargedPerson === id) : [];
      const caseRefs = personCases.map((caseObj: ComplianceCase) => caseObj.title);
      const filteredActions = Array.isArray(actionList)
        ? actionList.filter((action: DisciplinaryAction) => !!action?.caseReference && caseRefs.indexOf(action.caseReference) !== -1)
        : [];

      setMember(memberRecord);
      setCases(personCases);
      setActions(filteredActions);
      setTracker(trackerRecord ?? null);
      setLoading(false);
    };

    loadProfile().catch(() => setLoading(false));
  }, [id]);

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px'}} />
        <div style={{display: 'grid', gridTemplateColumns: '300px 1fr', gap: '2rem', marginTop: '2rem'}}>
          <div className="glass-panel skeleton" style={{height: '400px'}} />
          <div className="glass-panel skeleton" style={{height: '600px'}} />
        </div>
      </div>
    );
  }

  if (!member) {
    return (
      <div className="glass-panel text-center" style={{padding: '4rem'}}>
        <h2>Staff Member Not Found</h2>
        <NavLink to="/staff" className="btn btn-secondary" style={{marginTop: '1rem'}}>Back to Directory</NavLink>
      </div>
    );
  }

  const riskLevel = tracker ? escalationEngine.calculateRiskLevel(tracker) : 'Low';
  const riskColor = riskLevel === 'Critical' ? 'var(--status-danger)' : riskLevel === 'High' ? 'var(--status-warning)' : 'var(--status-success)';

  return (
    <div className="cases-container fade-in">
      <div className="cases-header">
        <NavLink to="/staff" className="btn btn-secondary"><ArrowLeft size={16}/> Back to Directory</NavLink>
        <div style={{display: 'flex', gap: '12px'}}>
          <button className="btn btn-primary"><Mail size={16}/> Contact Staff</button>
        </div>
      </div>

      <div style={{display: 'grid', gridTemplateColumns: '350px 1fr', gap: '2rem', marginTop: '1rem'}}>
        <div style={{display: 'flex', flexDirection: 'column', gap: '2rem'}}>
          <div className="glass-panel" style={{padding: '2rem', textAlign: 'center'}}>
            <div style={{
              width: '100px', height: '100px',
              background: member.photoUrl ? 'none' : 'linear-gradient(135deg, var(--primary), #d2334e)',
              borderRadius: '50%', margin: '0 auto 1.5rem',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              fontSize: '2.5rem', fontWeight: 600, color: 'white',
              boxShadow: '0 8px 16px rgba(220, 38, 38, 0.2)',
              overflow: 'hidden'
            }}>
              {member.photoUrl ? (
                <img src={member.photoUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
              ) : (
                member.fullName.charAt(0)
              )}
            </div>
            <h2 style={{margin: 0}}>{member.fullName}</h2>
            <p className="text-secondary" style={{marginTop: '0.5rem'}}>{member.role}</p>

            <div style={{marginTop: '2rem', textAlign: 'left', display: 'flex', flexDirection: 'column', gap: '1rem'}}>
              <div style={{display: 'flex', alignItems: 'center', gap: '12px', fontSize: '0.9rem'}}>
                <Building2 size={16} color="var(--text-muted)"/> <span>{member.department} / {member.company}</span>
              </div>
              <div style={{display: 'flex', alignItems: 'center', gap: '12px', fontSize: '0.9rem'}}>
                <Mail size={16} color="var(--text-muted)"/> <span style={{wordBreak: 'break-all'}}>{member.email}</span>
              </div>
              <div style={{display: 'flex', alignItems: 'center', gap: '12px', fontSize: '0.9rem'}}>
                <Briefcase size={16} color="var(--text-muted)"/> <span>Line Manager: <b>{member.lineManager}</b></span>
              </div>
            </div>

            <div style={{marginTop: '2rem', paddingTop: '2rem', borderTop: '1px solid var(--border-light)'}}>
              <div className="text-secondary" style={{fontSize: '0.75rem', textTransform: 'uppercase', marginBottom: '0.5rem'}}>Compliance Status</div>
              <div style={{
                padding: '12px', borderRadius: '8px',
                background: `${riskColor}15`, border: `1px solid ${riskColor}30`,
                color: riskColor, fontWeight: 700, fontSize: '1.1rem'
              }}>
                {riskLevel} Risk Profile
              </div>
            </div>
          </div>

          <div className="glass-panel" style={{padding: '1.5rem'}}>
            <h3 style={{fontSize: '1rem', marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <TrendingUp size={18} color="var(--primary)"/> Offence Matrix
            </h3>
            <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
              <div style={{display: 'flex', justifyContent: 'space-between', padding: '8px', background: 'rgba(0,0,0,0.02)', borderRadius: '6px'}}>
                <span className="text-secondary">Tier 1 (6m)</span>
                <span style={{fontWeight: 600}}>{tracker?.tier1Last6Months || 0}</span>
              </div>
              <div style={{display: 'flex', justifyContent: 'space-between', padding: '8px', background: 'rgba(0,0,0,0.02)', borderRadius: '6px'}}>
                <span className="text-secondary">Tier 2 Total</span>
                <span style={{fontWeight: 600}}>{tracker?.tier2Offences || 0}</span>
              </div>
              <div style={{display: 'flex', justifyContent: 'space-between', padding: '8px', background: 'rgba(0,0,0,0.02)', borderRadius: '6px'}}>
                <span className="text-secondary">Tier 3 Total</span>
                <span style={{fontWeight: 600}}>{tracker?.tier3Offences || 0}</span>
              </div>
              <div style={{display: 'flex', justifyContent: 'space-between', padding: '8px', background: 'rgba(0,0,0,0.02)', borderRadius: '6px'}}>
                <span className="text-secondary">Total Breaches</span>
                <span style={{fontWeight: 600}}>{tracker?.totalOffences || 0}</span>
              </div>
            </div>
          </div>
        </div>

        <div style={{display: 'flex', flexDirection: 'column', gap: '2rem'}}>
          <div className="glass-panel" style={{padding: '2rem'}}>
            <h3 style={{marginBottom: '2rem', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <History size={20} color="var(--secondary)"/> Compliance History Timeline
            </h3>

            <div className="activity-list">
              {cases.length === 0 ? (
                <div className="text-center text-secondary" style={{padding: '2rem'}}>No compliance incidents recorded for this staff member.</div>
              ) : (
                cases.map((c) => (
                  <div key={c.id} className="activity-item" style={{opacity: 1, transform: 'none'}}>
                    <div className="activity-indicator" style={{background: c.status === 'Paid' ? 'var(--status-success)' : 'var(--primary)'}} />
                    <div className="activity-content" style={{width: '100%'}}>
                      <div style={{display: 'flex', justifyContent: 'space-between', width: '100%'}}>
                        <div className="activity-title" style={{fontSize: '1.05rem'}}>{c.offenceCategoryName}</div>
                        <span className={`status-badge status-${c.status.toLowerCase()}`}>{c.status}</span>
                      </div>
                      <div className="activity-desc" style={{margin: '0.5rem 0'}}>{c.offenceDescription}</div>
                      <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '0.5rem'}}>
                        <div className="activity-time">{new Date(c.dateCreated).toLocaleDateString()}</div>
                        <NavLink to={`/cases/${c.id}`} className="link-action" style={{fontSize: '0.85rem'}}>View Case File</NavLink>
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="glass-panel" style={{padding: '2rem'}}>
            <h3 style={{marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '8px'}}>
              <ShieldAlert size={20} color="var(--primary)"/> Disciplinary Ledger
            </h3>
            <table className="pact-table">
              <thead>
                <tr>
                  <th>Action Type</th>
                  <th>Date Issued</th>
                  <th>Penalty</th>
                  <th>Issued By</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {actions.length === 0 ? (
                  <tr><td colSpan={5} className="text-center">No disciplinary actions on record.</td></tr>
                ) : (
                  actions.map(action => (
                    <tr key={action.id}>
                      <td style={{fontWeight: 600, color: 'var(--text-primary)'}}>{action.actionType}</td>
                      <td>{new Date(action.actionDate).toLocaleDateString()}</td>
                      <td>{((action as any).penaltyAmount !== undefined && (action as any).penaltyAmount > 0) ? `₦${(action as any).penaltyAmount.toLocaleString()}` : '--'}</td>
                      <td>{action.actionedBy}</td>
                      <td>
                        <span className={`status-badge ${action.status === 'Enforced' ? 'status-paid' : 'status-unpaid'}`}>
                          {action.status}
                        </span>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};
