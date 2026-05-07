import React, { useState } from 'react';
import { AlertCircle, Send, CheckCircle, Clock } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';
import { APPEAL_SLA_DAYS } from '../../config/constants';

interface Props {
  caseData: ComplianceCase;
}

/**
 * AppealSubmissionPage — Pre-filled appeal form for employees.
 * Employee lands here after clicking "Appeal" in their notification email.
 * Case reference and appellant name are auto-filled and locked.
 */
export const AppealSubmissionPage: React.FC<Props> = ({ caseData }) => {
  const [grounds, setGrounds] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!grounds.trim()) {
      setError('Please provide your grounds for appeal.');
      return;
    }

    if (grounds.trim().length < 20) {
      setError('Please provide a more detailed explanation (minimum 20 characters).');
      return;
    }

    setSubmitting(true);
    setError(null);

    try {
      await sharePointService.createAppeal({
        caseReference: caseData.title,
        appellant: caseData.chargedPersonName,
        grounds: grounds.trim(),
      });
      setSubmitted(true);
    } catch (err) {
      console.error('Appeal submission failed:', err);
      setError('Failed to submit your appeal. Please try again or contact HR directly.');
      setSubmitting(false);
    }
  };

  // Success Screen
  if (submitted) {
    return (
      <div className="response-card">
        <div className="response-success">
          <div className="success-icon">
            <CheckCircle size={32} />
          </div>
          <h2>Appeal Submitted Successfully</h2>
          <p>
            Your appeal for case <strong>{caseData.title}</strong> has been received. 
            HR will review your appeal and respond within <strong>{APPEAL_SLA_DAYS} working days</strong>.
          </p>
          <div style={{
            marginTop: '28px',
            padding: '16px 24px',
            background: 'rgba(16, 185, 129, 0.08)',
            border: '1px solid rgba(16, 185, 129, 0.2)',
            borderRadius: '10px',
            display: 'inline-flex',
            alignItems: 'center',
            gap: '10px',
            fontSize: '0.85rem',
            color: '#10b981'
          }}>
            <Clock size={16} />
            Expected response by: <strong>
              {(() => {
                const d = new Date();
                let added = 0;
                while (added < APPEAL_SLA_DAYS) {
                  d.setDate(d.getDate() + 1);
                  if (d.getDay() !== 0 && d.getDay() !== 6) added++;
                }
                return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
              })()}
            </strong>
          </div>
          <p style={{ marginTop: '24px', fontSize: '0.8rem', color: '#64748b' }}>
            You will receive an email notification once a decision has been made.
            You may close this page.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="response-card">
      {/* Header */}
      <div className="response-card-header">
        <div className="icon-circle appeal">
          <AlertCircle size={24} />
        </div>
        <div>
          <h2>Submit an Appeal</h2>
          <p>Case Reference: <strong style={{ color: '#e2e8f0' }}>{caseData.title}</strong></p>
        </div>
      </div>

      <div className="response-card-body">
        {/* Case Summary */}
        <div className="response-info-grid" style={{ marginBottom: '32px' }}>
          <div className="response-info-item">
            <div className="label">Charged Person</div>
            <div className="value">{caseData.chargedPersonName}</div>
          </div>
          <div className="response-info-item">
            <div className="label">Department</div>
            <div className="value">{caseData.department}</div>
          </div>
          <div className="response-info-item">
            <div className="label">Offence</div>
            <div className="value">{caseData.offenceCategoryName}</div>
          </div>
          <div className="response-info-item">
            <div className="label">Penalty Amount</div>
            <div className="value highlight">₦{caseData.penaltyAmount.toLocaleString()}</div>
          </div>
          <div className="response-info-item full-width">
            <div className="label">Description</div>
            <div className="value" style={{ fontSize: '0.88rem', lineHeight: 1.5, color: '#94a3b8' }}>
              {caseData.offenceDescription}
            </div>
          </div>
        </div>

        {/* Appeal Form */}
        <form onSubmit={handleSubmit}>
          {/* Locked fields */}
          <div className="appeal-form-group">
            <label>Case Reference</label>
            <input
              type="text"
              value={caseData.title}
              disabled
              className="locked"
            />
          </div>

          <div className="appeal-form-group">
            <label>Appellant Name</label>
            <input
              type="text"
              value={caseData.chargedPersonName || ''}
              disabled
              className="locked"
            />
          </div>

          {/* Editable field */}
          <div className="appeal-form-group">
            <label>Grounds for Appeal <span style={{ color: '#e94560' }}>*</span></label>
            <textarea
              placeholder="Please explain in detail why you believe this penalty should be reviewed, reduced, or waived. Include any supporting evidence or circumstances..."
              value={grounds}
              onChange={(e) => {
                setGrounds(e.target.value);
                if (error) setError(null);
              }}
              disabled={submitting}
            />
            <div style={{ 
              display: 'flex', 
              justifyContent: 'space-between', 
              marginTop: '6px', 
              fontSize: '0.75rem', 
              color: '#64748b' 
            }}>
              <span>{grounds.length < 20 && grounds.length > 0 ? `Minimum 20 characters required` : ''}</span>
              <span>{grounds.length} characters</span>
            </div>
          </div>

          {/* Error */}
          {error && (
            <div style={{
              padding: '12px 16px',
              background: 'rgba(239, 68, 68, 0.1)',
              border: '1px solid rgba(239, 68, 68, 0.2)',
              borderRadius: '8px',
              color: '#ef4444',
              fontSize: '0.85rem',
              marginBottom: '20px',
              display: 'flex',
              alignItems: 'center',
              gap: '8px'
            }}>
              <AlertCircle size={16} />
              {error}
            </div>
          )}

          {/* SLA Info */}
          <div style={{
            padding: '14px 18px',
            background: 'rgba(245, 158, 11, 0.06)',
            border: '1px solid rgba(245, 158, 11, 0.15)',
            borderRadius: '10px',
            fontSize: '0.82rem',
            color: '#94a3b8',
            marginBottom: '24px',
            lineHeight: 1.5
          }}>
            <strong style={{ color: '#f59e0b' }}>Please note:</strong> Upon submission, your appeal will be 
            reviewed by HR within <strong style={{ color: '#f59e0b' }}>{APPEAL_SLA_DAYS} working days</strong>. 
            You will receive an email notification with the outcome. Until a decision is made, the original 
            penalty remains in effect.
          </div>

          {/* Submit */}
          <div style={{ display: 'flex', gap: '12px' }}>
            <button
              type="submit"
              className="response-btn primary"
              disabled={submitting || !grounds.trim()}
              style={{ flex: 1 }}
            >
              {submitting ? (
                <>
                  <div className="spinner" style={{ width: '18px', height: '18px', borderWidth: '2px', margin: 0 }} />
                  Submitting Appeal...
                </>
              ) : (
                <>
                  <Send size={18} />
                  Submit Appeal
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
