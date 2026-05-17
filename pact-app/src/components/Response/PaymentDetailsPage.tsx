import React from 'react';
import { CheckCircle, Clock, AlertTriangle, Building } from 'lucide-react';
import type { ComplianceCase } from '../../config/types';
import { PAYMENT_DEADLINE_DAYS } from '../../config/constants';

interface Props {
  caseData: ComplianceCase;
}

/**
 * PaymentDetailsPage — Read-only page showing penalty details, bank info, and 7-day deadline.
 * Employee lands here after clicking "Accept" in their notification email.
 */
export const PaymentDetailsPage: React.FC<Props> = ({ caseData }) => {
  // Calculate deadline from due date
  const dueDate = new Date(caseData.dueDate);
  const today = new Date();
  const daysRemaining = Math.ceil((dueDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  const isOverdue = daysRemaining < 0;

  return (
    <>
      {/* Card Header */}
      <div className="response-card">
        <div className="response-card-header">
          <div className="icon-circle accept">
            <CheckCircle size={24} />
          </div>
          <div>
            <h2>Penalty Payment Details</h2>
            <p>Case Reference: <strong style={{ color: '#e2e8f0' }}>{caseData.title}</strong></p>
          </div>
        </div>

        <div className="response-card-body">
          {/* Penalty Amount Banner */}
          <div className="penalty-banner">
            <div className="penalty-label">Total Penalty Amount</div>
            <div className="penalty-amount">₦{caseData.penaltyAmount.toLocaleString()}</div>
            <div className="penalty-deadline">
              <Clock size={16} />
              {isOverdue
                ? `Overdue by ${Math.abs(daysRemaining)} day(s)`
                : `Payment deadline: ${PAYMENT_DEADLINE_DAYS} days from date of issue`
              }
            </div>
          </div>

          {/* Case Info Grid */}
          <div className="response-info-grid">
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
              <div className="label">Date Issued</div>
              <div className="value">{new Date(caseData.dateCreated).toLocaleDateString('en-GB', {
                day: 'numeric',
                month: 'long',
                year: 'numeric'
              })}</div>
            </div>
            <div className="response-info-item full-width">
              <div className="label">Description</div>
              <div className="value" style={{ fontSize: '0.88rem', lineHeight: 1.5, color: '#94a3b8' }}>
                {caseData.offenceDescription}
              </div>
            </div>
            <div className="response-info-item">
              <div className="label">Due Date</div>
              <div className="value" style={{ color: isOverdue ? '#ef4444' : '#f59e0b' }}>
                {dueDate.toLocaleDateString('en-GB', {
                  day: 'numeric',
                  month: 'long',
                  year: 'numeric'
                })}
              </div>
            </div>
            <div className="response-info-item">
              <div className="label">Status</div>
              <div className={`value ${isOverdue ? '' : 'highlight'}`} style={isOverdue ? { color: '#ef4444', fontWeight: 700 } : {}}>
                {isOverdue ? 'OVERDUE' : caseData.status}
              </div>
            </div>
          </div>

          {/* Bank Details */}
          <div className="bank-details-section">
            <h3>
              <Building size={18} />
              Payment Bank Details
            </h3>
            <div className="bank-details-row">
              <span className="bd-label">Bank Name</span>
              <span className="bd-value">First Bank of Nigeria</span>
            </div>
            <div className="bank-details-row">
              <span className="bd-label">Account Name</span>
              <span className="bd-value">Konstructum Construction Ltd</span>
            </div>
            <div className="bank-details-row">
              <span className="bd-label">Account Number</span>
              <span className="bd-value">0123456789</span>
            </div>
            <div className="bank-details-row">
              <span className="bd-label">Payment Reference</span>
              <span className="bd-value">{caseData.title}</span>
            </div>
            <div className="bank-details-row">
              <span className="bd-label">Amount</span>
              <span className="bd-value" style={{ color: '#e94560', fontWeight: 700 }}>
                ₦{caseData.penaltyAmount.toLocaleString()}
              </span>
            </div>
          </div>

          {/* Deadline Warning */}
          <div className="deadline-alert">
            <AlertTriangle size={20} className="alert-icon" />
            <div className="alert-text">
              <strong>Important:</strong> Payment must be made within <strong>{PAYMENT_DEADLINE_DAYS} days</strong> of the 
              date this notice was issued. Failure to pay within the deadline will result in automatic escalation 
              of this case. Please use the case reference <strong>{caseData.title}</strong> as your payment reference 
              to ensure correct allocation.
            </div>
          </div>

          {/* Contact Note */}
          <div style={{
            padding: '16px 20px',
            background: 'rgba(255, 255, 255, 0.02)',
            borderRadius: '10px',
            border: '1px solid rgba(255, 255, 255, 0.04)',
            fontSize: '0.82rem',
            color: '#64748b',
            lineHeight: 1.6
          }}>
            <p style={{ margin: 0 }}>
              If you believe this penalty has been issued in error, you may file an appeal within the payment deadline. 
              Contact your line manager or HR department for guidance. To file an appeal, use the Appeal button 
              provided in your notification email.
            </p>
          </div>
        </div>
      </div>
    </>
  );
};
