import type { RepeatOffenceRecord, PolicyOffence } from '../config/types';
import { ESCALATION_THRESHOLD, ESCALATION_WINDOW_MONTHS } from '../config/constants';

export class EscalationEngine {
  
  /**
   * Determine if an offence escalation is required for a specific policy.
   * Escalation triggers AFTER 3 prior occurrences of that specific offence
   * (meaning the current incident is the 4th+ incident of that offence).
   * Prior offences of other types or higher tiers do NOT trigger escalation for a Tier 1 offence.
   */
  public checkSpecificOffenceEscalation(specificOffenceCount: number): boolean {
    return specificOffenceCount >= ESCALATION_THRESHOLD;
  }

  /**
   * Determine if a Tier 1 repeat-offence escalation is required.
   * Returns true when the staff member already has >= ESCALATION_THRESHOLD (3)
   * offences for this specific offence category.
   */
  public checkTier1Escalation(
    newCaseDate: string, 
    repeatTrackerRecord: RepeatOffenceRecord,
    specificOffenceCount: number = 0
  ): boolean {
    if (specificOffenceCount > 0) {
      return specificOffenceCount >= ESCALATION_THRESHOLD;
    }
    return (repeatTrackerRecord.tier1Last6Months || 0) >= ESCALATION_THRESHOLD;
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
   */
  public checkTier3Escalation(tier: string): boolean {
    return tier === 'Tier 3';
  }

  /**
   * Determine if a Tier 2 repeat-offence escalation is required.
   */
  public checkTier2Escalation(tracker: RepeatOffenceRecord | null, specificOffenceCount: number = 0): boolean {
    if (specificOffenceCount > 0) {
      return specificOffenceCount >= ESCALATION_THRESHOLD;
    }
    if (!tracker) return false;
    return tracker.tier2Offences >= ESCALATION_THRESHOLD;
  }

  /**
   * Calculates penalty amount based on tier and offence count:
   * Tier 1: 1st = ₦5,000 | 2nd = ₦10,000 | 3rd+ = ₦20,000
   * Tier 2: 1st = ₦10,000 | 2nd = ₦10,000 | 3rd+ = ₦20,000
   * Tier 3: 1st = ₦20,000 | 2nd = ₦20,000 | 3rd+ = ₦20,000
   * Returns 0 if policy is non-monetary (defaultPenaltyAmount === 0).
   */
  public calculatePenaltyAmount(policy: PolicyOffence, offenceCount: number): number {
    if (policy.defaultPenaltyAmount === 0) return 0;

    const tier = (policy.tier || 'Tier 1').trim();
    const count = Math.max(1, offenceCount);

    if (tier === 'Tier 1') {
      if (count === 1) return 5000;
      if (count === 2) return 10000;
      return 20000;
    }
    if (tier === 'Tier 2') {
      if (count === 1) return 10000;
      if (count === 2) return 10000;
      return 20000;
    }
    if (tier === 'Tier 3') {
      return 20000;
    }
    return 5000;
  }

  public calculateRiskLevel(tracker: RepeatOffenceRecord): 'Low' | 'Medium' | 'High' | 'Critical' {
    const t1 = tracker.tier1Last6Months || 0;
    const t2 = (tracker.tier2Offences || 0) + Math.floor(t1 / 4);
    const t3 = (tracker.tier3Offences || 0) + Math.floor(t2 / 2);

    if (t3 > 0) return 'Critical';
    if (t2 > 1) return 'High';
    if (t2 === 1 || t1 >= 2) return 'Medium';
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

