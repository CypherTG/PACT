/**
 * PACT Platform Constants
 * Central configuration for SharePoint site, list names, and column internal names
 */

// ─── Organisation Email Addresses ───────────────────────────────────────────
export const HR_EMAIL = 'mbello@konstructum.com'; // HR Department
export const LEGAL_EMAIL = 'legal@konstructum.com'; // Legal Department
export const COMPLIANCE_EMAIL = 'mbello@konstructum.com'; // Compliance (placeholder)
export const APPEAL_SLA_DAYS = 3; // Working days for appeal review
export const PAYMENT_DEADLINE_DAYS = 7; // Days to pay penalty

/**
 * Employee case-response one-pager (e.g. Vercel). Same hash routes as pact-app.
 * Example: `https://your-deployment.vercel.app`
 * Leave empty to build links from the current SharePoint page URL.
 */
export const RESPONSE_PORTAL_BASE_URL = '';

/** Accept / Appeal: must match query on links from notices (blocks direct typing / bookmarks). */
export const CASE_RESPONSE_FROM_EMAIL_QUERY_KEY = 'pact_src';
export const CASE_RESPONSE_FROM_EMAIL_QUERY_VALUE = 'email';

// ─── SharePoint Site ────────────────────────────────────────────────────────
export const SHAREPOINT_SITE_URL = 'netorgft13110820.sharepoint.com';
export const SHAREPOINT_SITE_PATH = '/sites/KONSTRUCTUM';
export const SHAREPOINT_SITE_ID = ''; // Will be resolved at runtime via Graph

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
    CHARGED_PERSON: 'ChargedPersaon', 
    STAFF_EMAIL: 'Charged_x0020_Person_x0020_Email',
    DEPARTMENT: 'Department',
    OFFENCE_CATEGORY: 'Offence_x0020_Category', 
    OFFENCE_DESCRIPTION: 'Offence_x0020_Description',
    PENALTY_AMOUNT: 'Penalty_x0020_Amount',
    DUE_DATE: 'Due_x0020_Date',
    ISSUER_NAME: 'Issuer_x0020_Name',
    SECONDARY_CONTACT: 'Secondary_x0020_Contact_x0020_Email', 
    STATUS: 'Status',
    DATE_CREATED: 'DateCreated', 
    EVIDENCE: 'Evidence',
  },
  STAFF: {
    TITLE: 'Title', 
    EMAIL: 'Email',
    DEPARTMENT: 'Department',
    ROLE: 'Role',
    LINE_MANAGER: 'Line_x0020_Manager',
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
    CASE_REFERENCE: 'Case_x0020_Reference',
    ACTION_TYPE: 'Action_x0020_Type',
    ACTION_DATE: 'Action_x0020_Date',
    ACTIONED_BY: 'Actioned_x0020_By',
    NOTES: 'Notes',
    STATUS: 'Status',
  },
  ESCALATION: {
    TITLE: 'Title',
    CASE_REFERENCE: 'Case_x0020_Reference',
    OFFENDER: 'Offender',
    REASON: 'Escalation_x0020_Reason',
    PREVIOUS_TIER: 'Previous_x0020_Tier',
    NEW_TIER: 'New_x0020_Tier',
    TRIGGERED_BY: 'Triggered_x0020_By',
    DATE: 'Escalation_x0020_Date',
    NOTIFIED_TO: 'Notified_x0020_To',
  },
  APPEALS: {
    TITLE: 'Title',
    CASE_REFERENCE: 'Case_x0020_Reference',
    APPELLANT: 'Appellant',
    APPEAL_DATE: 'Appeal_x0020_Date',
    GROUNDS: 'Grounds_x0020_for_x0020_Appeal',
    REVIEWING_OFFICER: 'Reviewing_x0020_Officer',
    DECISION: 'Decision',
    DECISION_DATE: 'Decision_x0020_Date',
    DECISION_NOTES: 'Decision_x0020_Notes',
  },
  REPEAT_TRACKER: {
    TITLE: 'Title',
    OFFENDER: 'Offender',
    TOTAL_OFFENCES: 'Total_x0020_Offences',
    TIER1_LAST_6M: 'Tier1_x0020_Offences',
    TIER2_OFFENCES: 'Tier2_x0020_Offences',
    TIER3_OFFENCES: 'Tier3_x0020_Offences',
    RISK_LEVEL: 'RiskLevel',
    LAST_OFFENCE_DATE: 'Last_x0020_Offence_x0020_Date',
    ESCALATION_DUE: 'Escalation_x0020_Due',
  },
  MAIL: {
    TITLE: 'Subject',
    TO: 'RecipientEmail',
    SUBJECT: 'Subject',
    BODY: 'MailBody',
    STATUS: 'Status',
  },
} as const;

// ─── Enums ──────────────────────────────────────────────────────────────────
export const CASE_STATUS = ['Unpaid', 'Paid', 'Overdue', 'Waived', 'Acknowledged'] as const;
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
export const ESCALATION_THRESHOLD = 2; // Number of Tier 1 offences in history to trigger escalation on the NEXT one (e.g. 2 existing + 1 current = 3 total)
export const ESCALATION_WINDOW_MONTHS = 6; // Rolling window in months
