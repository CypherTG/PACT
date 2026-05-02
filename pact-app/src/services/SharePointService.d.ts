import type { ComplianceCase, DashboardStats, StaffMember, PolicyOffence, EscalationEntry, RepeatOffenceRecord, UserSession } from '../config/types';
export declare class SharePointService {
    private siteUrl;
    private runtimeMode;
    private isLocal;
    private STORAGE_PREFIX;
    private readonly REST_TIMEOUT_MS;
    private static _instance;
    private _spfxContext;
    static init(context: any): void;
    constructor();
    getCurrentSession(): UserSession;
    getPhotoUrl(email: string): string;
    /**
     * Helper to ensuring professional expanded terms are visible even in legacy cases
     */
    private expandAbbreviations;
    /**
     * Safe initialization of data.
     * Prevents app-wide failure during module load.
     */
    initialize(): Promise<void>;
    private fetchREST;
    private getFormDigest;
    private sendEmailNotification;
    private getFromLocal;
    private saveToLocal;
    isStandalone(): boolean;
    getRuntimeLabel(): string;
    getUserName(): string;
    getCases(): Promise<ComplianceCase[]>;
    getStaffDirectory(): Promise<StaffMember[]>;
    getPolicyLibrary(): Promise<PolicyOffence[]>;
    private parsePenalty;
    private readField;
    private formatDisciplinaryReference;
    private getNextDisciplinaryReference;
    createCase(caseData: Partial<ComplianceCase>): Promise<ComplianceCase>;
    updateCase(id: string, updates: Partial<ComplianceCase>): Promise<void>;
    deleteCase(id: string): Promise<void>;
    getDashboardStats(): Promise<DashboardStats>;
    private buildDashboardStats;
    getEscalationLog(): Promise<EscalationEntry[]>;
    getRepeatTrackerRecord(staffId: string): Promise<RepeatOffenceRecord | null>;
    private getRepeatTrackerRecords;
    updateRepeatTracker(staffId: string, updates: Partial<RepeatOffenceRecord>): Promise<void>;
    createEscalation(entry: Partial<EscalationEntry>): Promise<void>;
    getDisciplinaryActions(caseRef?: string): Promise<any[]>;
    createDisciplinaryAction(action: any): Promise<void>;
    getAppeals(): Promise<any[]>;
    getMailHistory(): Promise<any[]>;
    createAppeal(appeal: any): Promise<void>;
    updateAppeal(id: string, updates: any): Promise<void>;
    private mapSPItemToCase;
    private getInitialMockCases;
    private getInitialMockTrackers;
    private getInitialMockEscalations;
    private getInitialMockStaff;
    private getInitialMockPolicies;
    private getInitialMockAppeals;
    private getInitialMockDisciplinary;
}
export declare const sharePointService: SharePointService;
//# sourceMappingURL=SharePointService.d.ts.map