import React, { useState, useEffect } from 'react';
import { X, Send, AlertCircle, CheckCircle } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';

interface AppealFormProps {
  onClose: () => void;
  onSuccess: () => void;
}

export const AppealForm: React.FC<AppealFormProps> = ({ onClose, onSuccess }) => {
  const [cases, setCases] = useState<ComplianceCase[]>([]);
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);

  const [formData, setFormData] = useState({
    caseReference: '',
    appellant: '',
    grounds: ''
  });

  useEffect(() => {
    sharePointService.getCases().then(data => {
      setCases(data);
      setLoading(false);
    });
  }, []);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.caseReference || !formData.appellant || !formData.grounds) {
      setError("Please fill in all fields.");
      return;
    }

    setSubmitting(true);
    setError(null);

    try {
      await sharePointService.createAppeal(formData);
      setSuccess(true);
      setTimeout(() => {
        onSuccess();
      }, 1500);
    } catch (err) {
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
          <label>Appellant Name</label>
          <input 
            type="text" 
            placeholder="Your full name"
            value={formData.appellant}
            onChange={(e) => setFormData({...formData, appellant: e.target.value})}
            disabled={submitting}
            className="form-input"
            data-testid="appeal-appellant-input"
          />
        </div>

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
