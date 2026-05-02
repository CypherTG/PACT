import type { RepeatOffenceRecord, PolicyOffence } from '../config/types';
import { ESCALATION_THRESHOLD, ESCALATION_WINDOW_MONTHS } from '../config/constants';

export class EscalationEngine {
  
  /**
   * Determine if a Tier 1 repeat-offence escalation is required.
   * Returns true when the staff member already has >= ESCALATION_THRESHOLD
   * Tier 1 offences in the rolling 6-month window (meaning the NEXT one is the 3rd+).
   */
  public checkTier1Escalation(
    newCaseDate: string, 
    repeatTrackerRecord: RepeatOffenceRecord
  ): boolean {
    const caseDate = new Date(newCaseDate);
    const sixMonthsAgo = new Date(caseDate);
    sixMonthsAgo.setMonth(caseDate.getMonth() - ESCALATION_WINDOW_MONTHS);
    // Fire when existing Tier 1 count is already at threshold (2),
    // meaning this new case is the 3rd within the window.
    return repeatTrackerRecord.tier1Last6Months >= ESCALATION_THRESHOLD;
  }

  /**
   * @deprecated Use checkTier1Escalation. Kept for backward compatibility with NewCaseForm.
   */
  public checkEscalation(
    newCaseDate: string,
    repeatTrackerRecord: RepeatOffenceRecord
  ): boolean {
    return this.checkTier1Escalation(newCaseDate, repeatTrackerRecord);
  }

  /**
   * Determine if a Tier 3 offence should always trigger an escalation.
   * Every single Tier 3 incident is immediately escalated regardless of history.
   */
  public checkTier3Escalation(tier: string): boolean {
    return tier === 'Tier 3';
  }

  /**
   * Determine if a Tier 2 repeat-offence escalation is required.
   * Returns true on 2nd+ Tier 2 offence.
   */
  public checkTier2Escalation(tracker: RepeatOffenceRecord | null): boolean {
    if (!tracker) return false;
    return tracker.tier2Offences >= 1;
  }

  /**
   * Calculates the overall risk level for an employee based on updated counts.
   */
  public calculateRiskLevel(tracker: RepeatOffenceRecord): 'Low' | 'Medium' | 'High' | 'Critical' {
    if (tracker.tier3Offences > 0) return 'Critical';
    if (tracker.tier2Offences > 1) return 'High';
    if (tracker.tier2Offences === 1 || tracker.tier1Last6Months >= 2) return 'Medium';
    return 'Low';
  }

  /**
   * Gets the recommended action based on the tier and offence count.
   */
  public getRecommendedAction(
    policy: PolicyOffence, 
    offenceCount: number, 
    isEscalated: boolean = false
  ): string {
    if (policy.tier === 'Tier 3') {
      if (offenceCount === 1) return policy.firstOffenceAction;
      if (offenceCount === 2) return policy.secondOffenceAction;
      return policy.thirdOffenceAction;
    }

    if (isEscalated || policy.tier === 'Tier 2') {
      if (offenceCount === 1) return policy.firstOffenceAction;
      if (offenceCount === 2) return policy.secondOffenceAction;
      return policy.thirdOffenceAction;
    }
    
    // Default Tier 1 processing
    if (offenceCount === 1) return policy.firstOffenceAction;
    if (offenceCount === 2) return policy.secondOffenceAction;
    return policy.thirdOffenceAction;
  }
}

export const escalationEngine = new EscalationEngine();
