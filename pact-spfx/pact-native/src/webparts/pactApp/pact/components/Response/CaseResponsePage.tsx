import React, { useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';
import { PaymentDetailsPage } from './PaymentDetailsPage';
import { AppealSubmissionPage } from './AppealSubmissionPage';
import { AlertTriangle } from 'lucide-react';

function getHashQueryParams(): URLSearchParams {
  const raw = typeof window !== 'undefined' ? window.location.hash : '';
  const q = raw.indexOf('?');
  return q >= 0 ? new URLSearchParams(raw.slice(q + 1)) : new URLSearchParams();
}

function buildFallbackCase(caseId: string): ComplianceCase {
  const params = getHashQueryParams();
  const dateCreated = new Date().toISOString();
  const dueDate = params.get('due') || params.get('date') || new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
  const amount = Number(params.get('amount') || 0);

  return {
    id: 'fallback',
    title: caseId,
    chargedPerson: '0',
    chargedPersonName: params.get('name') || 'Staff Member',
    offenceCategory: '0',
    offenceCategoryName: params.get('offence') || 'Compliance notice',
    penaltyAmount: Number.isFinite(amount) ? amount : 0,
    dueDate,
    staffEmail: params.get('email') || '',
    department: params.get('department') || 'Not specified',
    offenceDescription: params.get('description') || 'Please review the case details provided in your compliance notice.',
    issuerName: '',
    secondaryContact: '',
    status: 'Unpaid',
    dateCreated
  };
}

function enrichCaseFromLink(caseData: ComplianceCase): ComplianceCase {
  const fallback = buildFallbackCase(caseData.title);
  return {
    ...caseData,
    chargedPersonName: caseData.chargedPersonName || fallback.chargedPersonName,
    staffEmail: caseData.staffEmail || fallback.staffEmail,
    department: caseData.department || fallback.department,
    offenceCategoryName: caseData.offenceCategoryName || fallback.offenceCategoryName,
    offenceDescription: caseData.offenceDescription || fallback.offenceDescription,
    penaltyAmount: caseData.penaltyAmount || fallback.penaltyAmount,
    dueDate: caseData.dueDate || fallback.dueDate,
    dateCreated: caseData.dateCreated || fallback.dateCreated
  };
}

/**
 * CaseResponsePage - Employee Accept / Appeal landing page.
 * Reads :caseId and :action from the URL and renders the correct self-service view.
 */
export const CaseResponsePage: React.FC = () => {
  const { caseId, action } = useParams<{ caseId: string; action: string }>();
  const [caseData, setCaseData] = useState<ComplianceCase | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const loadCase = async (): Promise<void> => {
      try {
        await sharePointService.initialize();

        if (!caseId) {
          setError('No case reference provided.');
          setLoading(false);
          return;
        }

        const found = await sharePointService.getCaseByReference(caseId);

        if (!found) {
          setCaseData(buildFallbackCase(caseId));
          setLoading(false);
          return;
        }

        setCaseData(enrichCaseFromLink(found));
        setLoading(false);
      } catch (err) {
        console.error('Failed to load case:', err);
        setError('Unable to load case details. Please try again later.');
        setLoading(false);
      }
    };

    void loadCase();
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

  if (action === 'appeal') {
    return <AppealSubmissionPage caseData={caseData} />;
  }

  return <PaymentDetailsPage caseData={caseData} />;
};
