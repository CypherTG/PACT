import React, { useEffect, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';
import { PaymentDetailsPage } from './PaymentDetailsPage';
import { AppealSubmissionPage } from './AppealSubmissionPage';
import { AlertTriangle } from 'lucide-react';

/**
 * CaseResponsePage — Router wrapper for employee email links.
 * Reads :caseId and :action from the URL and renders the appropriate view.
 * 
 * URL format: /case-response/:caseId/:action
 * - /case-response/PACT-001/accept → Payment Details
 * - /case-response/PACT-001/appeal → Appeal Form
 */
export const CaseResponsePage: React.FC = () => {
  const { caseId, action } = useParams<{ caseId: string; action: string }>();
  const navigate = useNavigate();
  const [caseData, setCaseData] = useState<ComplianceCase | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const loadCase = async () => {
      try {
        await sharePointService.initialize();

        if (!caseId) {
          setError('No case reference provided.');
          setLoading(false);
          return;
        }

        const found = await sharePointService.getCaseByReference(caseId);
        if (!found) {
          setError(`Case "${caseId}" not found. Please check the link in your email.`);
          setLoading(false);
          return;
        }

        setCaseData(found);
        setLoading(false);
      } catch (err) {
        console.error('Failed to load case:', err);
        setError('Unable to load case details. Please try again later.');
        setLoading(false);
      }
    };

    loadCase();
  }, [caseId]);

  if (loading) {
    return (
      <div className="response-card">
        <div className="response-loading">
          <div className="spinner" />
          <p style={{ color: '#94a3b8' }}>Loading case details...</p>
        </div>
      </div>
    );
  }

  if (error || !caseData) {
    return (
      <div className="response-card">
        <div className="response-error">
          <div className="error-icon">
            <AlertTriangle size={28} />
          </div>
          <h3>Case Not Found</h3>
          <p>{error || 'Unable to locate this case.'}</p>
          <p style={{ marginTop: '16px', fontSize: '0.8rem', color: '#64748b' }}>
            If you believe this is an error, please contact your line manager or HR.
          </p>
        </div>
      </div>
    );
  }

  // Route to the correct view
  if (action === 'appeal') {
    return <AppealSubmissionPage caseData={caseData} />;
  }

  // Default to payment details (accept)
  return <PaymentDetailsPage caseData={caseData} />;
};
