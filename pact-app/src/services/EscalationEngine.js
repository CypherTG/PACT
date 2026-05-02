"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.escalationEngine = exports.EscalationEngine = void 0;
var constants_1 = require("../config/constants");
var EscalationEngine = /** @class */ (function () {
    function EscalationEngine() {
    }
    /**
     * Determine if a Tier 1 repeat-offence escalation is required.
     * Returns true when the staff member already has >= ESCALATION_THRESHOLD
     * Tier 1 offences in the rolling 6-month window (meaning the NEXT one is the 3rd+).
     */
    EscalationEngine.prototype.checkTier1Escalation = function (newCaseDate, repeatTrackerRecord) {
        var caseDate = new Date(newCaseDate);
        var sixMonthsAgo = new Date(caseDate);
        sixMonthsAgo.setMonth(caseDate.getMonth() - constants_1.ESCALATION_WINDOW_MONTHS);
        // Fire when existing Tier 1 count is already at threshold (2),
        // meaning this new case is the 3rd within the window.
        return repeatTrackerRecord.tier1Last6Months >= constants_1.ESCALATION_THRESHOLD;
    };
    /**
     * @deprecated Use checkTier1Escalation. Kept for backward compatibility with NewCaseForm.
     */
    EscalationEngine.prototype.checkEscalation = function (newCaseDate, repeatTrackerRecord) {
        return this.checkTier1Escalation(newCaseDate, repeatTrackerRecord);
    };
    /**
     * Determine if a Tier 3 offence should always trigger an escalation.
     * Every single Tier 3 incident is immediately escalated regardless of history.
     */
    EscalationEngine.prototype.checkTier3Escalation = function (tier) {
        return tier === 'Tier 3';
    };
    /**
     * Determine if a Tier 2 repeat-offence escalation is required.
     * Returns true on 2nd+ Tier 2 offence.
     */
    EscalationEngine.prototype.checkTier2Escalation = function (tracker) {
        if (!tracker)
            return false;
        return tracker.tier2Offences >= 1;
    };
    /**
     * Calculates the overall risk level for an employee based on updated counts.
     */
    EscalationEngine.prototype.calculateRiskLevel = function (tracker) {
        if (tracker.tier3Offences > 0)
            return 'Critical';
        if (tracker.tier2Offences > 1)
            return 'High';
        if (tracker.tier2Offences === 1 || tracker.tier1Last6Months >= 2)
            return 'Medium';
        return 'Low';
    };
    /**
     * Gets the recommended action based on the tier and offence count.
     */
    EscalationEngine.prototype.getRecommendedAction = function (policy, offenceCount, isEscalated) {
        if (isEscalated === void 0) { isEscalated = false; }
        if (policy.tier === 'Tier 3') {
            if (offenceCount === 1)
                return policy.firstOffenceAction;
            if (offenceCount === 2)
                return policy.secondOffenceAction;
            return policy.thirdOffenceAction;
        }
        if (isEscalated || policy.tier === 'Tier 2') {
            if (offenceCount === 1)
                return policy.firstOffenceAction;
            if (offenceCount === 2)
                return policy.secondOffenceAction;
            return policy.thirdOffenceAction;
        }
        // Default Tier 1 processing
        if (offenceCount === 1)
            return policy.firstOffenceAction;
        if (offenceCount === 2)
            return policy.secondOffenceAction;
        return policy.thirdOffenceAction;
    };
    return EscalationEngine;
}());
exports.EscalationEngine = EscalationEngine;
exports.escalationEngine = new EscalationEngine();
//# sourceMappingURL=EscalationEngine.js.map