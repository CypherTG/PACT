/**
 * PACT Platform — TypeScript Interfaces
 * Matches all 7 SharePoint List schemas exactly
 */
export interface ComplianceCase {
    id: string;
    title: string;
    chargedPerson: string;
    chargedPersonName?: string;
    staffEmail: string;
    department: string;
    offenceCategory: string;
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
export interface StaffMember {
    id: string;
    fullName: string;
    email: string;
    department: string;
    role: string;
    lineManager: string;
    company: 'KCC' | 'KESL' | 'Interkonstruct';
    employeeType: 'Employee' | 'Consultant' | 'Contractor';
    status: 'Active' | 'Inactive';
    photoUrl?: string;
}
export interface PolicyOffence {
    id: string;
    offenceName: string;
    tier: 'Tier 1' | 'Tier 2' | 'Tier 3';
    category: 'Conduct' | 'Project Integrity' | 'Strategic' | 'EHSQ';
    description: string;
    defaultPenaltyAmount: number;
    firstOffenceAction: string;
    secondOffenceAction: string;
    thirdOffenceAction: string;
    escalationTrigger: boolean;
}
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
export interface RepeatOffenceRecord {
    id: string;
    title: string;
    offender: string;
    offenderName?: string;
    totalOffences: number;
    tier1Last6Months: number;
    tier2Offences: number;
    tier3Offences: number;
    riskLevel: 'Low' | 'Medium' | 'High' | 'Critical';
    lastOffenceDate: string;
    escalationDue: boolean;
}
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
export interface DashboardStats {
    totalActiveCases: number;
    escalationsThisMonth: number;
    pendingAppeals: number;
    repeatOffenders: number;
    totalFines: number;
    casesByTier: {
        tier: string;
        count: number;
    }[];
    casesByMonth: {
        month: string;
        tier1: number;
        tier2: number;
        tier3: number;
    }[];
    casesByDepartment: {
        department: string;
        count: number;
        risk: string;
    }[];
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
export type PageRoute = 'dashboard' | 'cases' | 'new-case' | 'case-detail' | 'staff' | 'staff-profile' | 'escalations' | 'appeals' | 'policies' | 'settings';
export interface NavigationItem {
    route: PageRoute;
    label: string;
    icon: string;
    badge?: number;
}
//# sourceMappingURL=types.d.ts.map