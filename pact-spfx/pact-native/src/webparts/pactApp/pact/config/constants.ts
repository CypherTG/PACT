/**
 * PACT Platform Constants
 * Central configuration for SharePoint site, list names, and column internal names
 */

import { readViteEnv } from './env';

// ─── Organisation Email Addresses ───────────────────────────────────────────
export const HR_EMAIL = 'mbello@konstructum.com'; // HR Department
export const LEGAL_EMAIL = 'abalogun@konstructum.com'; // Legal Department
export const COMPLIANCE_EMAIL = 'mbello@konstructum.com'; // Compliance (placeholder)
export const CHAIRMAN_EMAIL = readViteEnv('VITE_CHAIRMAN_EMAIL').trim() || 'mbello@konstructum.com'; // Chairman
export const MAIL_TRIGGER_URL = readViteEnv('VITE_MAIL_TRIGGER_URL').trim() || 'https://default37d4778d47da40aca3924a8c93c158.30.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a2651047ffd44146a15bdcb3d0fc110b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=MQClT4tptDeXdC7H3K4gcUmn0g1yb7YxzjJxYhBAtZA'; // HTTP Flow Webhook URL
export const ACCEPT_PAYMENT_TRIGGER_URL = readViteEnv('VITE_ACCEPT_PAYMENT_TRIGGER_URL').trim() || MAIL_TRIGGER_URL; // Accept/payment HTTP Flow URL
export const APPEAL_MAIL_TRIGGER_URL = readViteEnv('VITE_APPEAL_MAIL_TRIGGER_URL').trim() || 'https://default37d4778d47da40aca3924a8c93c158.30.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/13674850ca7a4f3d92da2893be5d9ad5/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=bcSyAnrS7dq9AQ1Dza6SDmpP0h0ybnqGPqRvPwsTgw0'; // Appeal HTTP Flow Webhook URL
export const APPEAL_SLA_DAYS = 3; // Working days for appeal review
export const PAYMENT_DEADLINE_DAYS = 1; // 24 hours to pay penalty

/** Values for PACT Compliance Cases → Status (add matching choices in SharePoint). */
export const CASE_STATUS = {
  UNPAID: 'Unpaid',
  PAID: 'Paid',
  OVERDUE: 'Overdue',
  WAIVED: 'Waived',
  APPEAL_PENDING: 'Appeal Pending',
} as const;

/**
 * Employee case-response one-pager (e.g. Vercel). Routes use HashRouter:
 * `/#/case-response/:caseRef/:action?...`
 * Set in `.env`: `VITE_RESPONSE_PORTAL_URL=https://your-app.vercel.app`
 * Leave empty to use the current page origin (SharePoint or local dev).
 */
const _envPortalUrl = readViteEnv('VITE_RESPONSE_PORTAL_URL').trim().replace(/\/$/, '');
export const RESPONSE_PORTAL_BASE_URL = _envPortalUrl.length > 0
  ? _envPortalUrl
  : 'https://netorgft13110820.sharepoint.com/sites/KONSTRUCTUM/SitePages/CollabHome.aspx';

/**
 * Accept / Appeal flows are gated: valid only when the URL includes this query (added to every email button link).
 */
export const CASE_RESPONSE_FROM_EMAIL_QUERY_KEY = 'pact_src';
export const CASE_RESPONSE_FROM_EMAIL_QUERY_VALUE = 'email';

// ─── SharePoint Site ────────────────────────────────────────────────────────
export const SHAREPOINT_SITE_URL = 'netorgft13110820.sharepoint.com';
export const SHAREPOINT_SITE_PATH = '/sites/KONSTRUCTUM';
export const SHAREPOINT_SITE_ID = ''; // Will be resolved at runtime via Graph

/** Site-relative document library for payment proof uploads (syncs via OneDrive). */
export const PAYMENT_PROOFS_LIBRARY = 'PACT Payment Proofs';

// ─── SharePoint List Names ──────────────────────────────────────────────────
export const LIST_NAMES = {
  COMPLIANCE_CASES: 'PACT Compliance Cases',
  STAFF_DIRECTORY: 'PACT Staff Directory',
  POLICY_LIBRARY: 'PACT Policy & Offence Library',
  DISCIPLINARY_ACTIONS: 'Disciplinary Actions',
  ESCALATION_LOG: 'Escalation Log',
  APPEALS_REGISTER: 'PACT Appeals Register',
  REPEAT_OFFENCE_TRACKER: 'PACT Repeat Offence Tracker',
  MAIL_HISTORY: 'PACT Mail History',
} as const;

// ─── Internal Column Names (SharePoint internal names) ──────────────────────
// Note: SharePoint stores internal names on column creation.
// Renaming a column only changes the display name.
// Documented typo: 'ChargedPersaon' is the internal name for 'Charged Person'
export const COLUMNS = {
  CASES: {
    TITLE: 'Title',
    CHARGED_PERSON: 'ChargedPerson', 
    STAFF_EMAIL: 'StaffEmail',
    CHARGED_PERSON_EMAIL: 'ChargedPersonEmail',
    DEPARTMENT: 'Department',
    OFFENCE_CATEGORY: 'OffenceCategory', 
    PENALTY_AMOUNT: 'PenaltyAmount',
    DUE_DATE: 'DueDate',
    ISSUER_NAME: 'IssuerName',
    SECONDARY_CONTACT: 'SecondaryContact', 
    STATUS: 'Status',
    TIER: 'Tier',
    DISCIPLINARY_ACTION: 'DisciplinaryAction',
    OFFENCE_COUNT: 'OffenceCount',
    ACTION_LABEL: 'ActionLabel',
    /** Optional on site — use display-name update if column exists */
    EVIDENCE: 'Evidence',
    PAYMENT_PROOF: 'Payment_x0020_Proof',
  },
  STAFF: {
    TITLE: 'Title',
    EMAIL: 'EmailAddress',
    DEPARTMENT: 'Department',
    ROLE: 'Role',
    LINE_MANAGER: 'LineManager',
    COMPANY: 'Company',
    EMPLOYEE_TYPE: 'Employee_x0020_Type',
    STATUS: 'Status',
  },
  POLICY: {
    TITLE: 'Title', // Offence Name
    TIER: 'Tier',
    CATEGORY: 'Category',
    DESCRIPTION: 'Offence_x0020_Description',
    DEFAULT_PENALTY: 'Default_x0020_Penalty_x0020_Amount',
    FIRST_ACTION: 'First_x0020_Offence_x0020_Action',
    SECOND_ACTION: 'Second_x0020_Offence_x0020_Action',
    THIRD_ACTION: 'Third_x0020_Offence_x0020_Action',
    ESCALATION_TRIGGER: 'Escalation_x0020_Trigger',
  },
  DISCIPLINARY: {
    TITLE: 'Title',
    CASE_REFERENCE: 'CaseReference',
    ACTION_TYPE: 'ActionType',
    ACTION_DATE: 'ActionDate',
    PENALTY_AMOUNT: 'PenaltyAmount',
    NOTES: 'Notes',
    STATUS: 'Status',
  },
  ESCALATION: {
    TITLE: 'Title',
    CASE_REFERENCE: 'CaseReference',
    OFFENDER: 'Offender',
    REASON: 'EscalationReason',
    PREVIOUS_TIER: 'PreviousTier',
    NEW_TIER: 'NewTier',
    TRIGGERED_BY: 'TriggeredBy',
    DATE: 'EscalationDate',
    NOTIFIED_TO: 'NotifiedTo',
  },
  APPEALS: {
    TITLE: 'Title',
    CASE_REFERENCE: 'CaseReference',
    APPELLANT: 'Appellant',
    APPEAL_DATE: 'AppealDate',
    GROUNDS: 'GroundsforAppeal',
    REVIEWING_OFFICER: 'ReviewingOfficer',
    DECISION: 'Decision',
    DECISION_DATE: 'DecisionDate',
    DECISION_NOTES: 'DecisionNotes',
  },
  REPEAT_TRACKER: {
    TITLE: 'Title',
    OFFENDER: 'Offender',
    TOTAL_OFFENCES: 'TotalOffences',
    TIER1_LAST_6M: 'Tier1Offences',
    TIER2_OFFENCES: 'Tier2Offences',
    TIER3_OFFENCES: 'Tier3Offences',
    RISK_LEVEL: 'RiskLevel',
    LAST_OFFENCE_DATE: 'LastOffenceDate',
    ESCALATION_DUE: 'EscalationDue',
  },
  MAIL: {
    TITLE: 'Title',
    TO: 'RecipientEmail',
    SUBJECT: 'Title',
    BODY: 'MailBody',
    STATUS: 'Status',
  },
} as const;

// ─── Enums ──────────────────────────────────────────────────────────────────
export const CASE_STATUS_OPTIONS = ['Unpaid', 'Paid', 'Overdue', 'Waived', 'Appeal Pending', 'Acknowledged'] as const;
export const TIERS = ['Tier 1', 'Tier 2', 'Tier 3'] as const;
export const CATEGORIES = ['Conduct', 'Project Integrity', 'Strategic', 'EHSQ'] as const;
export const COMPANIES = ['KCC', 'KESL', 'Interkonstruct'] as const;
export const EMPLOYEE_TYPES = ['Employee', 'Consultant', 'Contractor'] as const;
export const STAFF_STATUS = ['Active', 'Inactive'] as const;
export const ACTION_TYPES = ['Warning', 'Fine', 'Suspension', 'Termination'] as const;
export const DISCIPLINARY_STATUS = ['Pending', 'Enforced', 'Appealed', 'Waived'] as const;
export const ESCALATION_TRIGGERS = ['System', 'Manual'] as const;
export const APPEAL_DECISIONS = ['Upheld', 'Reduced', 'Waived', 'Rejected'] as const;
export const RISK_LEVELS = ['Low', 'Medium', 'High', 'Critical'] as const;

// ─── Escalation Rules ───────────────────────────────────────────────────────
export const ESCALATION_THRESHOLD = 3; // Number of Tier 1 offences in history to trigger escalation on the NEXT one (e.g. 3 existing + 1 current = 4 total)
export const ESCALATION_WINDOW_MONTHS = 6; // Rolling window in months
