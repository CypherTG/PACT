import React, { useEffect, useState, useMemo } from 'react';
import { useHistory } from 'react-router-dom';
import { 
  ShieldAlert, 
  Search, 
  CheckCircle, 
  AlertCircle,
  Upload,
  Loader2,
  Users
} from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { StaffMember, PolicyOffence } from '../../config/types';

export const PublicReportPage: React.FC = () => {
  const history = useHistory();
  const [success, setSuccess] = useState(false);

  // Dynamic Data
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [policies, setPolicies] = useState<PolicyOffence[]>([]);

  // Form State
  const [personId, setPersonId] = useState('');
  const [categoryId, setCategoryId] = useState('');
  const [description, setDescription] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

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

  const selectedPolicy = useMemo(() => policies.find(p => p.id === categoryId), [policies, categoryId]);
  const selectedStaff = useMemo(() => staff.find(s => s.id === personId), [staff, personId]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!personId || !categoryId) {
      alert("Please select a staff member and an infraction.");
      return;
    }
    
    setIsSubmitting(true);
    try {
      // We use the same service, but the Power Automate trigger handles the "External" flow
      await sharePointService.createCase({
        chargedPerson: personId,
        offenceCategory: categoryId,
        offenceDescription: description,
        dueDate: new Date(Date.now() + 7 * 86400000).toISOString(),
        evidence: selectedFile ? selectedFile.name : undefined
      });
      setIsSubmitting(false);
      setSuccess(true);
    } catch (error) {
      console.error("Submission failed:", error);
      setIsSubmitting(false);
      alert("Failed to log incident. Please check your connection.");
    }
  };

  if (success) {
    return (
      <div className="glass-panel text-center fade-in" style={{ padding: '60px', maxWidth: '600px', margin: '40px auto', background: 'var(--bg-panel)' }}>
        <div style={{ 
          width: '80px', height: '80px', background: 'rgba(16, 124, 16, 0.1)', 
          borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', 
          margin: '0 auto 24px', color: '#107c10' 
        }}>
          <CheckCircle size={40} />
        </div>
        <h2 style={{ color: 'var(--text-primary)' }}>Incident Logged</h2>
        <p className="text-secondary" style={{ marginTop: '16px', fontSize: '1.1rem' }}>
          Thank you for your report. The PACT engine has received the details and will notify the relevant supervisors.
        </p>
        <button 
          onClick={() => window.location.reload()} 
          className="btn btn-secondary" 
          style={{ marginTop: '32px' }}
        >
          Submit Another Report
        </button>
      </div>
    );
  }

  return (
    <div className="cases-container fade-in" style={{ maxWidth: '800px', margin: '0 auto', padding: '20px', fontFamily: 'var(--font-sans)', color: 'var(--text-primary)' }}>
      <div className="glass-panel" style={{ padding: '40px', background: 'var(--bg-panel)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '40px', borderBottom: '1px solid var(--border-light)', paddingBottom: '24px' }}>
          <div className="logo-container" style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div className="logo-icon" style={{ width: '48px', height: '48px', overflow: 'hidden', background: 'transparent', boxShadow: 'none' }}>
              <img 
                src={require('../../assets/kcc-logo.png')} 
                alt="KCC" 
                style={{ width: '100%', height: '100%', objectFit: 'contain' }}
                onError={(e) => {
                  (e.target as HTMLImageElement).style.display = 'none';
                  const parent = (e.target as HTMLImageElement).parentElement;
                  if (parent) {
                    parent.innerText = 'P';
                    parent.style.background = 'linear-gradient(135deg, var(--primary), #ff6b35)';
                    parent.style.display = 'flex';
                    parent.style.alignItems = 'center';
                    parent.style.justifyContent = 'center';
                    parent.style.color = 'white';
                    parent.style.fontWeight = 'bold';
                  }
                }}
              />
            </div>
            <h2 style={{ margin: 0, color: 'var(--text-primary)', fontWeight: 600, fontFamily: 'var(--font-display)' }}>PACT: Report Compliance Incident</h2>
          </div>
        </div>

        <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem', fontWeight: 600 }}>Who was involved?</label>
            <div style={{ position: 'relative' }}>
              <select 
                value={personId}
                onChange={e => setPersonId(e.target.value)}
                required
                className="form-input"
                style={{ width: '100%', padding: '14px', paddingLeft: '40px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--border-light)', borderRadius: '12px', color: 'var(--text-primary)', outline: 'none', appearance: 'none' }}
              >
                <option value="" style={{ background: '#1e293b', color: 'white' }}>-- Select Staff Member --</option>
                {staff.map(member => (
                  <option key={member.id} value={member.id} style={{ background: '#1e293b', color: 'white' }}>{member.fullName} ({member.department})</option>
                ))}
              </select>
              <Users size={18} style={{ position: 'absolute', left: '14px', top: '50%', transform: 'translateY(-50%)', opacity: 0.5, color: 'var(--text-secondary)' }} />
            </div>
          </div>

          {selectedStaff && (
            <div style={{ padding: '16px', background: 'rgba(255,255,255,0.04)', border: '1px solid var(--border-light)', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '16px', animation: 'fadeIn 0.3s ease' }}>
              <div className="avatar" style={{ width: '48px', height: '48px', border: '1px solid var(--border-light)' }}>
                {selectedStaff.photoUrl ? <img src={selectedStaff.photoUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} /> : selectedStaff.fullName.charAt(0)}
              </div>
              <div>
                <div style={{ fontWeight: 600, color: 'var(--text-primary)' }}>{selectedStaff.fullName}</div>
                <div className="text-secondary" style={{ fontSize: '0.8rem' }}>{selectedStaff.department} • {selectedStaff.role}</div>
              </div>
            </div>
          )}

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem', fontWeight: 600 }}>What happened? (Policy Category)</label>
            <select 
              value={categoryId}
              onChange={e => setCategoryId(e.target.value)}
              required
              className="form-input"
              style={{ width: '100%', padding: '14px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--border-light)', borderRadius: '12px', color: 'var(--text-primary)', outline: 'none' }}
            >
              <option value="" style={{ background: '#1e293b', color: 'white' }}>-- Select Infraction Type --</option>
              {policies.map(policy => (
                <option key={policy.id} value={policy.id} style={{ background: '#1e293b', color: 'white' }}>{policy.offenceName} ({policy.tier})</option>
              ))}
            </select>
          </div>

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem', fontWeight: 600 }}>Provide Details</label>
            <textarea 
              value={description}
              onChange={e => setDescription(e.target.value)}
              required
              rows={5}
              placeholder="Describe the incident in detail. Include dates, times, and locations..."
              style={{ width: '100%', padding: '14px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--border-light)', borderRadius: '12px', color: 'var(--text-primary)', outline: 'none', resize: 'vertical' }}
            />
          </div>

          <div className="form-group">
            <label className="text-secondary" style={{ display: 'block', marginBottom: '8px', fontSize: '0.9rem', fontWeight: 600 }}>Upload Evidence (Optional)</label>
            <div 
              onClick={() => document.getElementById('public-upload')?.click()}
              style={{ 
                border: '2px dashed var(--border-light)', 
                borderRadius: '12px', 
                padding: '30px', 
                textAlign: 'center',
                background: 'rgba(255,255,255,0.02)',
                cursor: 'pointer',
                transition: 'all 0.2s ease'
              }}
              onMouseEnter={(e) => e.currentTarget.style.borderColor = 'var(--primary)'}
              onMouseLeave={(e) => e.currentTarget.style.borderColor = 'var(--border-light)'}
            >
              <Upload size={24} className="text-secondary" style={{ margin: '0 auto 12px' }} />
              <div style={{ color: selectedFile ? 'var(--status-success)' : 'inherit', fontWeight: selectedFile ? 600 : 400 }}>
                {selectedFile ? `✓ ${selectedFile.name}` : 'Click to upload photos or documents'}
              </div>
              <input 
                type="file" 
                id="public-upload" 
                style={{ display: 'none' }} 
                onChange={(e) => { 
                  const file = e.target.files?.[0];
                  if(file) setSelectedFile(file);
                }} 
              />
            </div>
          </div>

          <div style={{ marginTop: '24px', borderTop: '1px solid var(--border-light)', paddingTop: '32px' }}>
            <button 
              type="submit" 
              className="btn btn-primary" 
              disabled={isSubmitting} 
              style={{ width: '100%', height: '56px', fontSize: '1.1rem', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px' }}
            >
              {isSubmitting ? (
                <>
                  <Loader2 size={20} className="animate-spin" />
                  Submitting Report...
                </>
              ) : (
                <>
                  <ShieldAlert size={20} />
                  Submit Formal Incident Report
                </>
              )}
            </button>
            <p className="text-center text-secondary" style={{ marginTop: '16px', fontSize: '0.8rem' }}>
              Your report will be reviewed by the Compliance Department. PACT version 5.0-Native.
            </p>
          </div>
        </form>
      </div>
    </div>
  );
};
