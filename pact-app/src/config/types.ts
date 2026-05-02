/**
 * PACT Platform — TypeScript Interfaces
 * Matches all 7 SharePoint List schemas exactly
 */

// ─── List 01: Compliance Cases ──────────────────────────────────────────────
export interface ComplianceCase {
  id: string;
  title: string;           // Penalty ID e.g. PACT-001
  chargedPerson: string;   // Lookup ID → Staff Directory
  chargedPersonName?: string;
  staffEmail: string;
  department: string;
  offenceCategory: string; // Lookup ID → Policy Library
  offenceCategoryName?: string;
  offenceDescription: string;
  penaltyAmount: number;
  dueDate: string;
  issuerName: string;
  secondaryContact: string;
  status: 'Unpaid' | 'Paid' | 'Overdue' | 'Waived';
  dateCreated: string;
  evidence?: string;
}

// ─── List 02: Staff Directory ───────────────────────────────────────────────
export interface StaffMember {
  id: string;
  fullName: string;        // Title
  email: string;
  department: string;
  role: string;
  lineManager: string;
  company: 'KCC' | 'KESL' | 'Interkonstruct';
  employeeType: 'Employee' | 'Consultant' | 'Contractor';
  status: 'Active' | 'Inactive';
  photoUrl?: string;
}

// ─── List 03: Policy & Offence Library ──────────────────────────────────────
export interface PolicyOffence {
  id: string;
  offenceName: string;     // Title
  tier: 'Tier 1' | 'Tier 2' | 'Tier 3';
  category: 'Conduct' | 'Project Integrity' | 'Strategic' | 'EHSQ';
  description: string;
  defaultPenaltyAmount: number;
  firstOffenceAction: string;
  secondOffenceAction: string;
  thirdOffenceAction: string;
  escalationTrigger: boolean;
}

// ─── List 04: Disciplinary Actions ──────────────────────────────────────────
export interface DisciplinaryAction {
  id: string;
  title: string;
  caseReference: string;
  caseReferenceName?: string;
  actionType: string;
  actionDate: string;
  actionedBy: string;
  penaltyAmount?: number;
  notes?: string;
  status: 'Pending' | 'Enforced' | 'Appealed' | 'Waived';
}

// ─── List 05: Escalation Log ────────────────────────────────────────────────
export interface EscalationEntry {
  id: string;
  title: string;
  caseReference: string;
  offender: string;
  offenderName?: string;
  escalationReason: string;
  previousTier: 'Tier 1' | 'Tier 2' | 'Tier 3';
  newTier: 'Tier 1' | 'Tier 2' | 'Tier 3';
  triggeredBy: 'System' | 'Manual';
  escalationDate: string;
  notifiedTo: string;
}

// ─── List 06: Appeals Register ──────────────────────────────────────────────
export interface Appeal {
  id: string;
  title: string;
  caseReference: string;
  appellantName: string;
  appellantDisplayName?: string;
  appealDate: string;
  groundsForAppeal: string;
  reviewingOfficer: string;
  decision: 'Upheld' | 'Reduced' | 'Waived' | 'Rejected';
  decisionDate?: string;
  decisionNotes?: string;
}

// ─── List 07: Repeat Offence Tracker ────────────────────────────────────────
export interface RepeatOffenceRecord {
  id: string;
  title: string;           // Offender name
  offender: string;        // Lookup ID → Staff Directory
  offenderName?: string;
  totalOffences: number;
  tier1Last6Months: number;
  tier2Offences: number;
  tier3Offences: number;
  riskLevel: 'Low' | 'Medium' | 'High' | 'Critical';
  lastOffenceDate: string;
  escalationDue: boolean;
}

// ─── List 08: Mail History ──────────────────────────────────────────────────
export interface MailLogEntry {
  id: string;
  to: string[];
  subject: string;
  body: string;
  timestamp: string;
  status: 'Sent' | 'Failed' | 'Pending' | 'Processing';
}

export interface UserSession {
  displayName: string;
  email: string;
  photoUrl?: string;
  tenantName?: string;
  siteTitle?: string;
}

// ─── UI Types ───────────────────────────────────────────────────────────────
export interface DashboardStats {
  totalActiveCases: number;
  escalationsThisMonth: number;
  pendingAppeals: number;
  repeatOffenders: number;
  totalFines: number;
  casesByTier: { tier: string; count: number }[];
  casesByMonth: { month: string; tier1: number; tier2: number; tier3: number }[];
  casesByDepartment: { department: string; count: number; risk: string }[];
  recentActivity: ActivityItem[];
}

export interface ActivityItem {
  id: string;
  type: 'case' | 'escalation' | 'appeal' | 'action';
  title: string;
  description: string;
  timestamp: string;
  severity: 'low' | 'medium' | 'high' | 'critical';
}

export type PageRoute =
  | 'dashboard'
  | 'cases'
  | 'new-case'
  | 'case-detail'
  | 'staff'
  | 'staff-profile'
  | 'escalations'
  | 'appeals'
  | 'policies'
  | 'settings';

export interface NavigationItem {
  route: PageRoute;
  label: string;
  icon: string;
  badge?: number;
}
