/**
 * PACT Platform Constants
 * Central configuration for SharePoint site, list names, and column internal names
 */
export declare const SHAREPOINT_SITE_URL = "netorgft13110820.sharepoint.com";
export declare const SHAREPOINT_SITE_PATH = "/sites/KONSTRUCTUM";
export declare const SHAREPOINT_SITE_ID = "";
export declare const LIST_NAMES: {
    readonly COMPLIANCE_CASES: "PACT Compliance Cases";
    readonly STAFF_DIRECTORY: "PACT Staff Directory";
    readonly POLICY_LIBRARY: "PACT Policy & Offence Library";
    readonly DISCIPLINARY_ACTIONS: "Disciplinary Actions";
    readonly ESCALATION_LOG: "Escalation Log";
    readonly APPEALS_REGISTER: "PACT Appeals Register";
    readonly REPEAT_OFFENCE_TRACKER: "PACT Repeat Offence Tracker";
    readonly MAIL_HISTORY: "PACT Mail History";
};
export declare const COLUMNS: {
    readonly CASES: {
        readonly TITLE: "Title";
        readonly CHARGED_PERSON: "ChargedPersaon";
        readonly STAFF_EMAIL: "Charged_x0020_Person_x0020_Email";
        readonly DEPARTMENT: "Department";
        readonly OFFENCE_CATEGORY: "Offence_x0020_Category";
        readonly OFFENCE_DESCRIPTION: "Offence_x0020_Description";
        readonly PENALTY_AMOUNT: "Penalty_x0020_Amount";
        readonly DUE_DATE: "Due_x0020_Date";
        readonly ISSUER_NAME: "Issuer_x0020_Name";
        readonly SECONDARY_CONTACT: "Secondary_x0020_Contact_x0020_Email";
        readonly STATUS: "Status";
        readonly DATE_CREATED: "DateCreated";
        readonly EVIDENCE: "Evidence";
    };
    readonly STAFF: {
        readonly TITLE: "Title";
        readonly EMAIL: "Email";
        readonly DEPARTMENT: "Department";
        readonly ROLE: "Role";
        readonly LINE_MANAGER: "Line_x0020_Manager";
        readonly COMPANY: "Company";
        readonly EMPLOYEE_TYPE: "Employee_x0020_Type";
        readonly STATUS: "Status";
    };
    readonly POLICY: {
        readonly TITLE: "Title";
        readonly TIER: "Tier";
        readonly CATEGORY: "Category";
        readonly DESCRIPTION: "Offence_x0020_Description";
        readonly DEFAULT_PENALTY: "Default_x0020_Penalty_x0020_Amount";
        readonly FIRST_ACTION: "First_x0020_Offence_x0020_Action";
        readonly SECOND_ACTION: "Second_x0020_Offence_x0020_Action";
        readonly THIRD_ACTION: "Third_x0020_Offence_x0020_Action";
        readonly ESCALATION_TRIGGER: "Escalation_x0020_Trigger";
    };
    readonly DISCIPLINARY: {
        readonly TITLE: "Title";
        readonly CASE_REFERENCE: "Case_x0020_Reference";
        readonly ACTION_TYPE: "Action_x0020_Type";
        readonly ACTION_DATE: "Action_x0020_Date";
        readonly ACTIONED_BY: "Actioned_x0020_By";
        readonly NOTES: "Notes";
        readonly STATUS: "Status";
    };
    readonly ESCALATION: {
        readonly TITLE: "Title";
        readonly CASE_REFERENCE: "Case_x0020_Reference";
        readonly OFFENDER: "Offender";
        readonly REASON: "Escalation_x0020_Reason";
        readonly PREVIOUS_TIER: "Previous_x0020_Tier";
        readonly NEW_TIER: "New_x0020_Tier";
        readonly TRIGGERED_BY: "Triggered_x0020_By";
        readonly DATE: "Escalation_x0020_Date";
        readonly NOTIFIED_TO: "Notified_x0020_To";
    };
    readonly APPEALS: {
        readonly TITLE: "Title";
        readonly CASE_REFERENCE: "Case_x0020_Reference";
        readonly APPELLANT: "Appellant";
        readonly APPEAL_DATE: "Appeal_x0020_Date";
        readonly GROUNDS: "Grounds_x0020_for_x0020_Appeal";
        readonly REVIEWING_OFFICER: "Reviewing_x0020_Officer";
        readonly DECISION: "Decision";
        readonly DECISION_DATE: "Decision_x0020_Date";
        readonly DECISION_NOTES: "Decision_x0020_Notes";
    };
    readonly REPEAT_TRACKER: {
        readonly TITLE: "Title";
        readonly OFFENDER: "Offender";
        readonly TOTAL_OFFENCES: "Total_x0020_Offences";
        readonly TIER1_LAST_6M: "Tier1_x0020_Offences";
        readonly TIER2_OFFENCES: "Tier2_x0020_Offences";
        readonly TIER3_OFFENCES: "Tier3_x0020_Offences";
        readonly RISK_LEVEL: "RiskLevel";
        readonly LAST_OFFENCE_DATE: "Last_x0020_Offence_x0020_Date";
        readonly ESCALATION_DUE: "Escalation_x0020_Due";
    };
    readonly MAIL: {
        readonly TITLE: "Subject";
        readonly TO: "RecipientEmail";
        readonly SUBJECT: "Subject";
        readonly BODY: "MailBody";
        readonly STATUS: "Status";
    };
};
export declare const CASE_STATUS: readonly ["Unpaid", "Paid", "Overdue", "Waived"];
export declare const TIERS: readonly ["Tier 1", "Tier 2", "Tier 3"];
export declare const CATEGORIES: readonly ["Conduct", "Project Integrity", "Strategic", "EHSQ"];
export declare const COMPANIES: readonly ["KCC", "KESL", "Interkonstruct"];
export declare const EMPLOYEE_TYPES: readonly ["Employee", "Consultant", "Contractor"];
export declare const STAFF_STATUS: readonly ["Active", "Inactive"];
export declare const ACTION_TYPES: readonly ["Warning", "Fine", "Suspension", "Termination"];
export declare const DISCIPLINARY_STATUS: readonly ["Pending", "Enforced", "Appealed", "Waived"];
export declare const ESCALATION_TRIGGERS: readonly ["System", "Manual"];
export declare const APPEAL_DECISIONS: readonly ["Upheld", "Reduced", "Waived", "Rejected"];
export declare const RISK_LEVELS: readonly ["Low", "Medium", "High", "Critical"];
export declare const ESCALATION_THRESHOLD = 2;
export declare const ESCALATION_WINDOW_MONTHS = 6;
//# sourceMappingURL=constants.d.ts.map