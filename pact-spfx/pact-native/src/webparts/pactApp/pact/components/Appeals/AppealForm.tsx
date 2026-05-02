import React, { useState, useEffect } from 'react';
import { X, Send, AlertCircle, CheckCircle } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase, StaffMember } from '../../config/types';

interface AppealFormProps {
  onClose: () => void;
  onSuccess: () => void;
}

export const AppealForm = ({ onClose, onSuccess }: AppealFormProps): React.ReactElement => {
  const [cases, setCases] = useState<ComplianceCase[]>([]);
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);

  const [formData, setFormData] = useState({
    caseReference: '',
    appellantId: '',
    grounds: ''
  });
  const selectedStaff = staff.find(member => member.id === formData.appellantId);

  useEffect(() => {
    Promise.all([
      sharePointService.getCases(),
      sharePointService.getStaffDirectory()
    ]).then(([caseData, staffData]) => {
      setCases(caseData);
      setStaff(staffData);
      setLoading(false);
    }).catch(() => setLoading(false));
  }, []);

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    if (!formData.caseReference || !selectedStaff || !formData.grounds) {
      setError("Please fill in all fields.");
      return;
    }

    setSubmitting(true);
    setError(null);

    try {
      await sharePointService.createAppeal({
        caseReference: formData.caseReference,
        appellant: selectedStaff.fullName,
        appellantEmail: selectedStaff.email,
        grounds: formData.grounds
      });
      setSuccess(true);
      setTimeout(() => {
        onSuccess();
      }, 1500);
    } catch {
      setError("Failed to submit appeal. Please try again.");
      setSubmitting(false);
    }
  };

  if (success) {
    return (
      <div className="glass-panel" style={{ padding: '3rem', textAlign: 'center' }}>
        <CheckCircle size={48} color="var(--status-success)" style={{ marginBottom: '1rem' }} />
        <h3>Appeal Submitted Successfully</h3>
        <p className="text-secondary">Your appeal has been logged and will be reviewed by HR/Legal.</p>
      </div>
    );
  }

  return (
    <div className="glass-panel" style={{ padding: '2rem', position: 'relative' }} data-testid="appeal-form">
      <button 
        type="button"
        onClick={onClose} 
        style={{ position: 'absolute', top: '1rem', right: '1rem', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--text-muted)' }}
        data-testid="appeal-close-button"
      >
        <X size={20} />
      </button>

      <h2 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
        <AlertCircle size={24} color="var(--primary)" /> Submit Appeal
      </h2>

      <form onSubmit={handleSubmit} className="pact-form">
        <div className="form-group">
          <label>Select Challenged Case</label>
          <select 
            value={formData.caseReference} 
            onChange={(e) => setFormData({...formData, caseReference: e.target.value})}
            disabled={loading || submitting}
            className="form-input"
            data-testid="appeal-case-select"
          >
            <option value="">-- Choose a Case --</option>
            {cases.map(c => (
              <option key={c.id} value={c.title}>
                {c.title} - {c.offenceCategoryName} ({c.chargedPersonName})
              </option>
            ))}
          </select>
        </div>

        <div className="form-group">
          <label>Appellant Account</label>
          <select 
            value={formData.appellantId}
            onChange={(e) => setFormData({...formData, appellantId: e.target.value})}
            disabled={loading || submitting}
            className="form-input"
            data-testid="appeal-appellant-select"
          >
            <option value="">-- Select Appellant --</option>
            {staff.map(member => (
              <option key={member.id} value={member.id}>
                {member.fullName} ({member.email})
              </option>
            ))}
          </select>
        </div>

        {selectedStaff && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', border: '1px solid var(--border-light)', borderRadius: '8px', marginBottom: '1rem' }}>
            <div className="avatar avatar-photo" style={{ width: '48px', height: '48px' }}>
              <img
                src={(selectedStaff as any).photoUrl || sharePointService.getPhotoUrl(selectedStaff.email)}
                alt={selectedStaff.fullName || 'Appellant'}
              />
              <span className="avatar-fallback">{selectedStaff.fullName?.charAt(0) || 'U'}</span>
            </div>
            <div>
              <div style={{ fontWeight: 600 }}>{selectedStaff.fullName}</div>
              <div className="text-secondary" style={{ fontSize: '0.85rem' }}>{selectedStaff.email}</div>
            </div>
          </div>
        )}

        <div className="form-group">
          <label>Grounds for Appeal</label>
          <textarea 
            placeholder="Explain why this penalty should be waived or reduced..."
            value={formData.grounds}
            onChange={(e) => setFormData({...formData, grounds: e.target.value})}
            disabled={submitting}
            style={{ minHeight: '120px' }}
            className="form-input"
            data-testid="appeal-grounds-input"
          />
        </div>

        {error && (
          <div className="error-message" style={{ color: 'var(--status-danger)', marginBottom: '1rem', fontSize: '0.9rem' }}>
            {error}
          </div>
        )}

        <div style={{ marginTop: '2rem', display: 'flex', gap: '12px' }}>
          <button 
            type="submit" 
            className="btn btn-primary" 
            disabled={submitting || loading}
            style={{ flex: 1 }}
            data-testid="appeal-submit-button"
          >
            {submitting ? 'Submitting...' : <><Send size={16} /> Submit Appeal</>}
          </button>
          <button 
            type="button" 
            className="btn btn-secondary" 
            onClick={onClose}
            disabled={submitting}
            data-testid="appeal-cancel-button"
          >
            Cancel
          </button>
        </div>
      </form>
    </div>
  );
};
