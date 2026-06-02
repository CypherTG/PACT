import React, { useState } from 'react';
import { AlertCircle, Send, CheckCircle, Clock } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';
import { APPEAL_SLA_DAYS } from '../../config/constants';

interface Props {
  caseData: ComplianceCase;
}

function expectedAppealResponseDate(): string {
  const d = new Date();
  let added = 0;
  while (added < APPEAL_SLA_DAYS) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) added++;
  }
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
}

/**
 * AppealSubmissionPage — Pre-filled appeal form for employees.
 * Employee lands here after clicking "Appeal" in their notification email.
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
        appellantEmail: caseData.staffEmail,
        department: caseData.department,
        offence: caseData.offenceCategoryName,
        offenceDescription: caseData.offenceDescription,
        penaltyAmount: caseData.penaltyAmount,
        grounds: grounds.trim(),
      });
      setSubmitted(true);
      setSubmitting(false);
    } catch (err) {
      console.error('Appeal submission failed:', err);
      const detail = err instanceof Error ? err.message : '';
      setError(
        detail
          ? `Failed to submit your appeal: ${detail}`
          : 'Failed to submit your appeal. Please try again or contact Admin or Executive support directly.'
      );
      setSubmitting(false);
    }
  };

  if (submitted) {
    return (
      <div className="response-page appeal-page">
        <div className="response-card">
          <div className="response-success">
            <div className="success-icon">
              <CheckCircle size={32} />
            </div>
            <h2>Appeal Submitted Successfully</h2>
            <p>
              Your appeal for case <strong>{caseData.title}</strong> has been received.
              Admin or Executive review will assess your appeal and respond within <strong>{APPEAL_SLA_DAYS} working days</strong>.
            </p>
            <div className="appeal-success-meta">
              <Clock size={16} aria-hidden />
              <span>
                Expected response by: <strong>{expectedAppealResponseDate()}</strong>
              </span>
            </div>
            <p style={{ marginTop: '24px', fontSize: '0.8rem', color: '#64748b' }}>
              You will receive an email notification once a decision has been made.
              You may close this page.
            </p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="response-page appeal-page">
      <div className="response-card">
        <div className="response-card-header">
          <div className="icon-circle appeal">
            <AlertCircle size={24} />
          </div>
          <div>
            <h2>Submit an Appeal</h2>
            <p>
              Case reference: <strong style={{ color: '#e2e8f0' }}>{caseData.title}</strong>
            </p>
          </div>
        </div>

        <div className="response-card-body">
          <section className="response-summary" aria-labelledby="appeal-case-summary">
            <h3 id="appeal-case-summary" className="response-section-title">
              Case summary
            </h3>
            <div className="response-info-grid">
              <div className="response-info-item">
                <div className="label">Charged person</div>
                <div className="value">{caseData.chargedPersonName}</div>
              </div>
              <div className="response-info-item">
                <div className="label">Department</div>
                <div className="value">{caseData.department || '—'}</div>
              </div>
              <div className="response-info-item">
                <div className="label">Offence</div>
                <div className="value">{caseData.offenceCategoryName}</div>
              </div>
              <div className="response-info-item">
                <div className="label">Penalty amount</div>
                <div className="value highlight">₦{caseData.penaltyAmount.toLocaleString()}</div>
              </div>
              {caseData.offenceDescription ? (
                <div className="response-info-item full-width">
                  <div className="label">Description</div>
                  <div className="value" style={{ fontSize: '0.88rem', lineHeight: 1.5, color: '#94a3b8' }}>
                    {caseData.offenceDescription}
                  </div>
                </div>
              ) : null}
            </div>
          </section>

          <section className="response-form-section" aria-labelledby="appeal-form-heading">
            <h3 id="appeal-form-heading" className="response-section-title">
              Your appeal
            </h3>

            <form onSubmit={handleSubmit}>
              <div className="appeal-locked-fields">
                <div className="appeal-form-group">
                  <label htmlFor="appeal-case-ref">Case reference</label>
                  <input
                    id="appeal-case-ref"
                    type="text"
                    value={caseData.title}
                    disabled
                    className="locked"
                    readOnly
                  />
                </div>
                <div className="appeal-form-group">
                  <label htmlFor="appeal-appellant">Appellant name</label>
                  <input
                    id="appeal-appellant"
                    type="text"
                    value={caseData.chargedPersonName || ''}
                    disabled
                    className="locked"
                    readOnly
                  />
                </div>
              </div>

              <div className="appeal-form-group">
                <label htmlFor="appeal-grounds">
                  Grounds for appeal <span style={{ color: '#e94560' }}>*</span>
                </label>
                <textarea
                  id="appeal-grounds"
                  placeholder="Explain in detail why you believe this penalty should be reviewed, reduced, or waived. Include any supporting evidence or circumstances."
                  value={grounds}
                  onChange={(e) => {
                    setGrounds(e.target.value);
                    if (error) setError(null);
                  }}
                  disabled={submitting}
                  rows={6}
                />
                <div className="appeal-char-hint">
                  <span className={grounds.length > 0 && grounds.length < 20 ? 'warn' : ''}>
                    {grounds.length > 0 && grounds.length < 20 ? 'Minimum 20 characters required' : ''}
                  </span>
                  <span>{grounds.length} characters</span>
                </div>
              </div>

              {error ? (
                <div className="appeal-alert" role="alert">
                  <AlertCircle size={16} aria-hidden />
                  <span>{error}</span>
                </div>
              ) : null}

              <div className="appeal-notice">
                <strong>Please note:</strong> Upon submission, your appeal will be reviewed by Admin or Executive review within{' '}
                <strong>{APPEAL_SLA_DAYS} working days</strong>. You will receive an email with the outcome.
                Until a decision is made, the original penalty remains in effect.
              </div>

              <div className="appeal-submit-row">
                <button
                  type="submit"
                  className="response-btn primary"
                  disabled={submitting || !grounds.trim()}
                >
                  {submitting ? (
                    <>
                      <div
                        className="spinner"
                        style={{ width: '18px', height: '18px', borderWidth: '2px', margin: 0 }}
                        aria-hidden
                      />
                      Submitting appeal…
                    </>
                  ) : (
                    <>
                      <Send size={18} aria-hidden />
                      Submit appeal
                    </>
                  )}
                </button>
              </div>
            </form>
          </section>
        </div>
      </div>
    </div>
  );
};
