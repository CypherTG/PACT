import type { RepeatOffenceRecord, PolicyOffence } from '../config/types';
export declare class EscalationEngine {
    /**
     * Determine if a Tier 1 repeat-offence escalation is required.
     * Returns true when the staff member already has >= ESCALATION_THRESHOLD
     * Tier 1 offences in the rolling 6-month window (meaning the NEXT one is the 3rd+).
     */
    checkTier1Escalation(newCaseDate: string, repeatTrackerRecord: RepeatOffenceRecord): boolean;
    /**
     * @deprecated Use checkTier1Escalation. Kept for backward compatibility with NewCaseForm.
     */
    checkEscalation(newCaseDate: string, repeatTrackerRecord: RepeatOffenceRecord): boolean;
    /**
     * Determine if a Tier 3 offence should always trigger an escalation.
     * Every single Tier 3 incident is immediately escalated regardless of history.
     */
    checkTier3Escalation(tier: string): boolean;
    /**
     * Determine if a Tier 2 repeat-offence escalation is required.
     * Returns true on 2nd+ Tier 2 offence.
     */
    checkTier2Escalation(tracker: RepeatOffenceRecord | null): boolean;
    /**
     * Calculates the overall risk level for an employee based on updated counts.
     */
    calculateRiskLevel(tracker: RepeatOffenceRecord): 'Low' | 'Medium' | 'High' | 'Critical';
    /**
     * Gets the recommended action based on the tier and offence count.
     */
    getRecommendedAction(policy: PolicyOffence, offenceCount: number, isEscalated?: boolean): string;
}
export declare const escalationEngine: EscalationEngine;
//# sourceMappingURL=EscalationEngine.d.ts.map