"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.sharePointService = exports.SharePointService = void 0;
var tslib_1 = require("tslib");
/**
 * PACT Platform — SharePoint REST Service
 * Communicates directly with SharePoint Lists without Azure AD.
 * Bypasses IT/Azure admin requirements.
 */
var constants_1 = require("../config/constants");
var EscalationEngine_1 = require("./EscalationEngine");
var staffDirectory_json_1 = tslib_1.__importDefault(require("../data/staffDirectory.json"));
var policyLibrary_json_1 = tslib_1.__importDefault(require("../data/policyLibrary.json"));
// PLACEHOLDER: Power Automate HTTP Trigger URL
var POWER_AUTOMATE_URL = 'https://default37d4778d47da40aca3924a8c93c158.30.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/542b00c131884a3e8235161bb10bd625/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EMA0sCNKGWx86m-cXU5EyhZCq3lbbB-pSZVN8CSCk4E';
var SharePointService = /** @class */ (function () {
    function SharePointService() {
        this.STORAGE_PREFIX = 'pact_engine_v1_';
        this.REST_TIMEOUT_MS = 20000;
        this._spfxContext = null;
        this.siteUrl = "https://".concat(constants_1.SHAREPOINT_SITE_URL).concat(constants_1.SHAREPOINT_SITE_PATH);
        var params = new URLSearchParams(window.location.search);
        var requestedMode = (params.get('pactMode') || params.get('mode') || '').toLowerCase();
        var host = window.location.hostname.toLowerCase();
        var hostLooksNative = host.includes('sharepoint.com') || host.includes('sharepoint-df.com');
        this.runtimeMode = requestedMode === 'demo'
            ? 'demo'
            : requestedMode === 'sharepoint' || requestedMode === 'live' || hostLooksNative
                ? 'sharepoint'
                : 'demo';
        this.isLocal = this.runtimeMode !== 'sharepoint';
    }
    // Called by the SPFx WebPart to inject context before React renders
    SharePointService.init = function (context) {
        var _a, _b;
        if (!SharePointService._instance) {
            SharePointService._instance = exports.sharePointService;
        }
        SharePointService._instance._spfxContext = context;
        // If running on SharePoint, switch to live mode
        if ((_b = (_a = context === null || context === void 0 ? void 0 : context.pageContext) === null || _a === void 0 ? void 0 : _a.web) === null || _b === void 0 ? void 0 : _b.absoluteUrl) {
            SharePointService._instance.runtimeMode = 'sharepoint';
            SharePointService._instance.isLocal = false;
            SharePointService._instance.siteUrl = context.pageContext.web.absoluteUrl;
        }
    };
    // Returns current user session info from SPFx context or fallback
    SharePointService.prototype.getCurrentSession = function () {
        var _a, _b, _c;
        var ctx = this._spfxContext;
        if ((_a = ctx === null || ctx === void 0 ? void 0 : ctx.pageContext) === null || _a === void 0 ? void 0 : _a.user) {
            var user = ctx.pageContext.user;
            return {
                displayName: user.displayName || 'PACT User',
                email: user.email || user.loginName || '',
                photoUrl: user.email
                    ? "/_layouts/15/userphoto.aspx?size=S&accountname=".concat(encodeURIComponent(user.email))
                    : undefined,
                tenantName: ((_b = ctx.pageContext.web) === null || _b === void 0 ? void 0 : _b.title) || 'KONSTRUCTUM',
                siteTitle: ((_c = ctx.pageContext.site) === null || _c === void 0 ? void 0 : _c.absoluteUrl) || '',
            };
        }
        return {
            displayName: 'PACT Administrator',
            email: 'admin@konstructum.com',
            photoUrl: undefined,
            tenantName: 'KONSTRUCTUM',
            siteTitle: 'SharePoint',
        };
    };
    // Returns the photo URL for a specific staff member
    SharePointService.prototype.getPhotoUrl = function (email) {
        return "/_layouts/15/userphoto.aspx?size=S&accountname=".concat(encodeURIComponent(email));
    };
    /**
     * Helper to ensuring professional expanded terms are visible even in legacy cases
     */
    SharePointService.prototype.expandAbbreviations = function (text) {
        if (!text)
            return text;
        return text
            .replace(/\bWW\b/g, 'Written Warning')
            .replace(/\bVW\b/g, 'Verbal Warning')
            .replace(/\bT\b/g, 'Termination')
            .replace(/\bS1\b/g, 'Suspension (1 Day)')
            .replace(/\bS2\b/g, 'Suspension (2 Days)')
            .replace(/\bS4\b/g, 'Suspension (4 Days)')
            .replace(/\bT&P\b/g, 'Termination & Prosecution');
    };
    /**
     * Safe initialization of data.
     * Prevents app-wide failure during module load.
     */
    SharePointService.prototype.initialize = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var DATA_VERSION, currentVersion;
            return tslib_1.__generator(this, function (_a) {
                // We allow initialization regardless of isLocal to ensure local storage has fallbacks
                // but we only do destructive syncs if isLocal is detected or storage is totally empty.
                try {
                    console.log("[PACT] Initializing SharePoint Service (Mode: ".concat(this.runtimeMode, ")"));
                    DATA_VERSION = "4.0";
                    currentVersion = localStorage.getItem(this.STORAGE_PREFIX + 'data_version');
                    if (currentVersion !== DATA_VERSION) {
                        console.log("PACT: Initializing isolated storage namespace...");
                        this.saveToLocal('pact_policies', this.getInitialMockPolicies());
                        this.saveToLocal('pact_staff', this.getInitialMockStaff());
                        this.saveToLocal('pact_trackers', this.getInitialMockTrackers());
                        this.saveToLocal('pact_cases', this.getInitialMockCases());
                        this.saveToLocal('pact_escalations', this.getInitialMockEscalations());
                        this.saveToLocal('pact_appeals', this.getInitialMockAppeals());
                        this.saveToLocal('pact_disciplinary', this.getInitialMockDisciplinary());
                        if (!localStorage.getItem(this.STORAGE_PREFIX + 'pact_mail_history')) {
                            this.saveToLocal('pact_mail_history', []);
                        }
                        localStorage.setItem(this.STORAGE_PREFIX + 'data_version', DATA_VERSION);
                    }
                }
                catch (e) {
                    console.error("Initialization failed:", e);
                }
                return [2 /*return*/];
            });
        });
    };
    // ─── Core REST Helpers ──────────────────────────────────────────────────────
    SharePointService.prototype.fetchREST = function (endpoint_1) {
        return tslib_1.__awaiter(this, arguments, void 0, function (endpoint, options) {
            var url, headers, digest, controller, timeout, response, data;
            if (options === void 0) { options = {}; }
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.isLocal)
                            return [2 /*return*/, null];
                        url = "".concat(this.siteUrl, "/_api/").concat(endpoint);
                        headers = tslib_1.__assign({ 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }, options.headers);
                        if (!(options.method && options.method !== 'GET')) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.getFormDigest()];
                    case 1:
                        digest = _a.sent();
                        headers['X-RequestDigest'] = digest;
                        _a.label = 2;
                    case 2:
                        controller = new AbortController();
                        timeout = window.setTimeout(function () { return controller.abort(); }, this.REST_TIMEOUT_MS);
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, , 6, 7]);
                        return [4 /*yield*/, fetch(url, tslib_1.__assign(tslib_1.__assign({}, options), { headers: headers, signal: controller.signal }))];
                    case 4:
                        response = _a.sent();
                        if (!response.ok)
                            throw new Error("SharePoint REST error: ".concat(response.statusText));
                        return [4 /*yield*/, response.json()];
                    case 5:
                        data = _a.sent();
                        return [2 /*return*/, data.d];
                    case 6:
                        window.clearTimeout(timeout);
                        return [7 /*endfinally*/];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.getFormDigest = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, data;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, fetch("".concat(this.siteUrl, "/_api/contextinfo"), {
                            method: 'POST',
                            headers: { 'Accept': 'application/json;odata=verbose' }
                        })];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, data.d.GetContextWebInformation.FormDigestValue];
                }
            });
        });
    };
    SharePointService.prototype.sendEmailNotification = function (to, subject, body) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var debugEvent, history_1, endpoint, emailProperties, spData, error_1;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (this.isLocal) {
                            console.group("%c [LOCAL MOCK EMAIL] %c ".concat(subject), 'background: #0078d4; color: white; padding: 2px 5px; border-radius: 3px;', 'font-weight: bold;');
                            console.log("Recipients:", to.join(', '));
                            console.log("Content Preview:", body.replace(/<[^>]*>?/gm, ' ').substring(0, 100) + '...');
                            console.groupEnd();
                            debugEvent = new CustomEvent('pact-mock-email', { detail: { to: to, subject: subject } });
                            window.dispatchEvent(debugEvent);
                            history_1 = this.getFromLocal('pact_mail_history');
                            history_1.push({
                                id: Date.now(),
                                to: to,
                                subject: subject,
                                body: body,
                                timestamp: new Date().toISOString(),
                                status: 'Sent'
                            });
                            this.saveToLocal('pact_mail_history', history_1);
                            return [2 /*return*/];
                        }
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        endpoint = 'SP.Utilities.Utility.SendEmail';
                        emailProperties = {
                            'properties': {
                                '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                                'To': { 'results': to },
                                'Subject': subject,
                                'Body': body,
                                'AdditionalHeaders': {
                                    '__metadata': { 'type': 'Collection(SP.KeyValue)' },
                                    'results': [
                                        { '__metadata': { 'type': 'SP.KeyValue' }, 'Key': 'content-type', 'Value': 'text/html', 'ValueType': 'Edm.String' }
                                    ]
                                }
                            }
                        };
                        return [4 /*yield*/, this.fetchREST(endpoint, {
                                method: 'POST',
                                body: JSON.stringify(emailProperties)
                            })];
                    case 2:
                        _b.sent();
                        spData = (_a = {
                                '__metadata': { 'type': "SP.Data.".concat(constants_1.LIST_NAMES.MAIL_HISTORY.replace(/ /g, '_x0020_'), "ListItem") }
                            },
                            _a[constants_1.COLUMNS.MAIL.TO] = to.join(', '),
                            _a[constants_1.COLUMNS.MAIL.SUBJECT] = subject,
                            _a[constants_1.COLUMNS.MAIL.BODY] = body,
                            _a[constants_1.COLUMNS.MAIL.STATUS] = 'Sent',
                            _a);
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.MAIL_HISTORY, "')/items"), {
                                method: 'POST',
                                body: JSON.stringify(spData)
                            })];
                    case 3:
                        _b.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _b.sent();
                        console.error("Failed to send email via SharePoint REST:", error_1);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    // ─── LocalStorage Persistence ────────────────────────────────────────────────
    SharePointService.prototype.getFromLocal = function (key) {
        try {
            var data = localStorage.getItem(this.STORAGE_PREFIX + key);
            if (!data || data === 'undefined' || data === 'null')
                return [];
            var parsed = JSON.parse(data);
            return Array.isArray(parsed) ? parsed : [];
        }
        catch (e) {
            console.error("Error loading ".concat(key, " from storage:"), e);
            return [];
        }
    };
    SharePointService.prototype.saveToLocal = function (key, data) {
        try {
            localStorage.setItem(this.STORAGE_PREFIX + key, JSON.stringify(data));
        }
        catch (e) {
            console.error("Error saving ".concat(key, " to storage:"), e);
        }
    };
    // ─── Public API ─────────────────────────────────────────────────────────────
    SharePointService.prototype.isStandalone = function () {
        return this.runtimeMode === 'demo';
    };
    SharePointService.prototype.getRuntimeLabel = function () {
        return this.runtimeMode === 'sharepoint' ? 'SharePoint Native' : 'Workbench Demo';
    };
    SharePointService.prototype.getUserName = function () {
        return "PACT Administrator";
    };
    // --- Cases ---
    SharePointService.prototype.getCases = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var localCases, endpoint, data, _a, staff_1, policies_1, error_2;
            var _this = this;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        if (this.isLocal) {
                            localCases = this.getFromLocal('pact_cases');
                            return [2 /*return*/, localCases.map(function (c) { return (tslib_1.__assign(tslib_1.__assign({}, c), { offenceCategoryName: _this.expandAbbreviations(c.offenceCategoryName || ''), penaltyAmount: _this.parsePenalty(c.penaltyAmount) })); }).sort(function (a, b) { return new Date(b.dateCreated).getTime() - new Date(a.dateCreated).getTime(); })];
                        }
                        endpoint = "web/lists/getbytitle('".concat(constants_1.LIST_NAMES.COMPLIANCE_CASES, "')/items?$orderby=Created desc");
                        return [4 /*yield*/, this.fetchREST(endpoint)];
                    case 1:
                        data = _b.sent();
                        return [4 /*yield*/, Promise.all([
                                this.getStaffDirectory(),
                                this.getPolicyLibrary()
                            ])];
                    case 2:
                        _a = _b.sent(), staff_1 = _a[0], policies_1 = _a[1];
                        return [2 /*return*/, data.results.map(function (item) { return _this.mapSPItemToCase(item, staff_1, policies_1); })];
                    case 3:
                        error_2 = _b.sent();
                        console.warn("REST/Graph fetch failed, falling back to local mocks", error_2);
                        return [2 /*return*/, this.getFromLocal('pact_cases')];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // --- Staff ---
    SharePointService.prototype.getStaffDirectory = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                // ALWAYS use local JSON data for now as requested
                return [2 /*return*/, staffDirectory_json_1.default];
            });
        });
    };
    // --- Policies ---
    SharePointService.prototype.getPolicyLibrary = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                // ALWAYS use local JSON data for now as requested
                return [2 /*return*/, policyLibrary_json_1.default];
            });
        });
    };
    SharePointService.prototype.parsePenalty = function (value) {
        if (typeof value === 'number')
            return value;
        if (!value)
            return 0;
        // Handle currencies like "₦ 5,000.00" or "? 5,000.00"
        var cleaned = value.toString().replace(/[₦?, ]/g, '');
        var num = parseFloat(cleaned);
        return isNaN(num) ? 0 : num;
    };
    SharePointService.prototype.readField = function (item) {
        var keys = [];
        for (var _i = 1; _i < arguments.length; _i++) {
            keys[_i - 1] = arguments[_i];
        }
        for (var _a = 0, keys_1 = keys; _a < keys_1.length; _a++) {
            var key = keys_1[_a];
            var value = item === null || item === void 0 ? void 0 : item[key];
            if (value !== undefined && value !== null && value !== '') {
                return value;
            }
        }
        return undefined;
    };
    SharePointService.prototype.formatDisciplinaryReference = function (index) {
        return "P-".concat(String(index).padStart(3, '0'));
    };
    SharePointService.prototype.getNextDisciplinaryReference = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var existing, _a, maxNumber;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.isLocal) return [3 /*break*/, 1];
                        _a = this.getFromLocal('pact_disciplinary');
                        return [3 /*break*/, 3];
                    case 1: return [4 /*yield*/, this.getDisciplinaryActions()];
                    case 2:
                        _a = _b.sent();
                        _b.label = 3;
                    case 3:
                        existing = _a;
                        maxNumber = existing.reduce(function (max, item) {
                            var title = String((item === null || item === void 0 ? void 0 : item.title) || '');
                            var match = title.match(/^P-(\d+)$/i);
                            if (!match)
                                return max;
                            var value = Number.parseInt(match[1], 10);
                            return Number.isFinite(value) && value > max ? value : max;
                        }, 0);
                        return [2 /*return*/, this.formatDisciplinaryReference(maxNumber + 1)];
                }
            });
        });
    };
    SharePointService.prototype.createCase = function (caseData) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var newId, staff, _a, person, policies, _b, policy, newCase, tracker, policyTier, isTier3Escalation, isTier2Escalation, isTier1Escalation, isEscalated, reason, newTier, t3Count, updatedTier1, updatedTier2, updatedTier3, updatedTracker, offCount, actionPath, disciplinaryAction, cases, response, error_3, _c, manager, recipients, subject, body;
            var _d;
            return tslib_1.__generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        newId = Date.now().toString();
                        if (!this.isLocal) return [3 /*break*/, 1];
                        _a = this.getFromLocal('pact_staff');
                        return [3 /*break*/, 3];
                    case 1: return [4 /*yield*/, this.getStaffDirectory()];
                    case 2:
                        _a = _e.sent();
                        _e.label = 3;
                    case 3:
                        staff = _a;
                        person = staff.find(function (s) { return s.id === caseData.chargedPerson; });
                        if (!this.isLocal) return [3 /*break*/, 4];
                        _b = this.getFromLocal('pact_policies');
                        return [3 /*break*/, 6];
                    case 4: return [4 /*yield*/, this.getPolicyLibrary()];
                    case 5:
                        _b = _e.sent();
                        _e.label = 6;
                    case 6:
                        policies = _b;
                        policy = policies.find(function (p) { return p.id === caseData.offenceCategory; });
                        newCase = {
                            id: newId,
                            title: "PACT-".concat(newId.slice(-4)),
                            chargedPerson: caseData.chargedPerson || '999',
                            chargedPersonName: (person === null || person === void 0 ? void 0 : person.fullName) || 'Unknown Staff',
                            staffEmail: (person === null || person === void 0 ? void 0 : person.email) || 'staff@konstructum.com',
                            department: (person === null || person === void 0 ? void 0 : person.department) || 'General',
                            offenceCategory: caseData.offenceCategory || 'Unknown',
                            offenceCategoryName: this.expandAbbreviations((policy === null || policy === void 0 ? void 0 : policy.offenceName) || 'Unknown Offence'),
                            offenceDescription: caseData.offenceDescription || '',
                            penaltyAmount: (policy === null || policy === void 0 ? void 0 : policy.defaultPenaltyAmount) || 0,
                            dueDate: caseData.dueDate || new Date().toISOString(),
                            issuerName: this.getUserName(),
                            secondaryContact: (person === null || person === void 0 ? void 0 : person.lineManager) || '',
                            status: 'Unpaid',
                            dateCreated: new Date().toISOString()
                        };
                        if (!policy) {
                            console.warn("Policy not found for category:", caseData.offenceCategory);
                        }
                        return [4 /*yield*/, this.getRepeatTrackerRecord(newCase.chargedPerson)];
                    case 7:
                        tracker = _e.sent();
                        policyTier = (policy === null || policy === void 0 ? void 0 : policy.tier) || '';
                        isTier3Escalation = policyTier === 'Tier 3';
                        isTier2Escalation = policyTier === 'Tier 2' && EscalationEngine_1.escalationEngine.checkTier2Escalation(tracker);
                        isTier1Escalation = policyTier === 'Tier 1' && tracker
                            ? EscalationEngine_1.escalationEngine.checkTier1Escalation(newCase.dateCreated, tracker)
                            : false;
                        isEscalated = isTier3Escalation || isTier2Escalation || isTier1Escalation;
                        if (!(isEscalated && policy)) return [3 /*break*/, 9];
                        reason = '';
                        newTier = 'Tier 2';
                        if (isTier3Escalation) {
                            t3Count = ((tracker === null || tracker === void 0 ? void 0 : tracker.tier3Offences) || 0) + 1;
                            reason = "Tier 3 Offence (".concat(policy.offenceName, "): Automatic escalation. Occurrence #").concat(t3Count, ". Immediate HR & Chairman review required.");
                            newTier = 'Tier 3';
                        }
                        else if (isTier2Escalation) {
                            reason = "Repeat Tier 2 Offence (".concat(policy.offenceName, "): Staff member has ").concat(((tracker === null || tracker === void 0 ? void 0 : tracker.tier2Offences) || 0) + 1, " Tier 2 offences on record.");
                            newTier = 'Tier 3';
                        }
                        else {
                            reason = "Automatic Policy Trigger: Staff member reached 3+ Tier 1 offences within 6 months. Threshold exceeded on case ".concat(newCase.title, ".");
                            newTier = 'Tier 2';
                        }
                        return [4 /*yield*/, this.createEscalation({
                                caseReference: newCase.title,
                                offender: newCase.chargedPerson,
                                offenderName: newCase.chargedPersonName,
                                escalationReason: reason,
                                previousTier: policy.tier,
                                newTier: newTier
                            })];
                    case 8:
                        _e.sent();
                        _e.label = 9;
                    case 9:
                        updatedTier1 = ((tracker === null || tracker === void 0 ? void 0 : tracker.tier1Last6Months) || 0) + (policyTier === 'Tier 1' ? 1 : 0);
                        updatedTier2 = ((tracker === null || tracker === void 0 ? void 0 : tracker.tier2Offences) || 0) + (policyTier === 'Tier 2' ? 1 : 0);
                        updatedTier3 = ((tracker === null || tracker === void 0 ? void 0 : tracker.tier3Offences) || 0) + (policyTier === 'Tier 3' ? 1 : 0);
                        updatedTracker = {
                            totalOffences: ((tracker === null || tracker === void 0 ? void 0 : tracker.totalOffences) || 0) + 1,
                            tier1Last6Months: updatedTier1,
                            tier2Offences: updatedTier2,
                            tier3Offences: updatedTier3,
                            lastOffenceDate: newCase.dateCreated,
                            riskLevel: EscalationEngine_1.escalationEngine.calculateRiskLevel(tslib_1.__assign(tslib_1.__assign({}, (tracker || { id: '', title: '', offender: '', totalOffences: 0, tier1Last6Months: 0, tier2Offences: 0, tier3Offences: 0, riskLevel: 'Low', lastOffenceDate: '', escalationDue: false })), { tier1Last6Months: updatedTier1, tier2Offences: updatedTier2, tier3Offences: updatedTier3 }))
                        };
                        return [4 /*yield*/, this.updateRepeatTracker(newCase.chargedPerson, updatedTracker)];
                    case 10:
                        _e.sent();
                        offCount = ((tracker === null || tracker === void 0 ? void 0 : tracker.tier1Last6Months) || 0) + ((policy === null || policy === void 0 ? void 0 : policy.tier) === 'Tier 1' ? 1 : 0);
                        actionPath = isEscalated ? 'Automatic Escalation' :
                            ((policy === null || policy === void 0 ? void 0 : policy.tier) === 'Tier 1' ? (offCount === 1 ? '1st Offence' : offCount === 2 ? '2nd Offence' : '3rd+ Offence') : 'Standard');
                        disciplinaryAction = policy ? EscalationEngine_1.escalationEngine.getRecommendedAction(policy, offCount, isEscalated) : 'Standard Disciplinary Path';
                        if (!(this.isLocal || POWER_AUTOMATE_URL)) return [3 /*break*/, 15];
                        cases = this.getFromLocal('pact_cases');
                        cases.push(newCase);
                        this.saveToLocal('pact_cases', cases);
                        if (!POWER_AUTOMATE_URL) return [3 /*break*/, 14];
                        _e.label = 11;
                    case 11:
                        _e.trys.push([11, 13, , 14]);
                        return [4 /*yield*/, fetch(POWER_AUTOMATE_URL, {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify(tslib_1.__assign(tslib_1.__assign({}, newCase), { recommendedAction: disciplinaryAction }))
                            })];
                    case 12:
                        response = _e.sent();
                        if (!response.ok)
                            throw new Error('Power Automate submission failed');
                        return [3 /*break*/, 14];
                    case 13:
                        error_3 = _e.sent();
                        console.error("Power Automate Error:", error_3);
                        return [3 /*break*/, 14];
                    case 14: return [2 /*return*/, newCase];
                    case 15:
                        _c = this.createDisciplinaryAction;
                        _d = {};
                        return [4 /*yield*/, this.getNextDisciplinaryReference()];
                    case 16: 
                    // 5. Create Disciplinary Action Record
                    return [4 /*yield*/, _c.apply(this, [(_d.title = _e.sent(),
                                _d.caseReference = newCase.title,
                                _d.actionType = disciplinaryAction,
                                _d.penaltyAmount = newCase.penaltyAmount,
                                _d.notes = "Action Classification: ".concat(actionPath, ". Recommended by PACT Engine."),
                                _d.status = 'Pending',
                                _d)])];
                    case 17:
                        // 5. Create Disciplinary Action Record
                        _e.sent();
                        manager = staff.find(function (member) {
                            return member.fullName === (person === null || person === void 0 ? void 0 : person.lineManager) || member.email === (person === null || person === void 0 ? void 0 : person.lineManager);
                        });
                        recipients = Array.from(new Set([
                            newCase.staffEmail,
                            manager === null || manager === void 0 ? void 0 : manager.email
                        ].filter(Boolean)));
                        if (!(recipients.length > 0)) return [3 /*break*/, 19];
                        subject = "PACT NOTICE: ".concat(newCase.title, " - ").concat(newCase.offenceCategoryName);
                        body = "\n        <p>An offence has been logged against <b>".concat(newCase.chargedPersonName, "</b>.</p>\n        <p><b>Offence:</b> ").concat(newCase.offenceCategoryName, "</p>\n        <p><b>Sanction:</b> ").concat(disciplinaryAction, "</p>\n        <p><b>Reference:</b> ").concat(newCase.title, "</p>\n      ");
                        return [4 /*yield*/, this.sendEmailNotification(recipients, subject, body)];
                    case 18:
                        _e.sent();
                        _e.label = 19;
                    case 19: return [2 /*return*/, newCase];
                }
            });
        });
    };
    SharePointService.prototype.updateCase = function (id, updates) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cases, idx, spData;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.isLocal) {
                            cases = this.getFromLocal('pact_cases');
                            idx = cases.findIndex(function (c) { return c.id === id; });
                            if (idx > -1) {
                                cases[idx] = tslib_1.__assign(tslib_1.__assign({}, cases[idx]), updates);
                                this.saveToLocal('pact_cases', cases);
                            }
                            return [2 /*return*/];
                        }
                        spData = {
                            Status: updates.status,
                            // Add other field mapping as needed
                        };
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.COMPLIANCE_CASES, "')/items(").concat(id, ")"), {
                                method: 'POST', // SharePoint uses POST with MERGE header for updates
                                headers: {
                                    'X-HTTP-Method': 'MERGE',
                                    'IF-MATCH': '*'
                                },
                                body: JSON.stringify(spData)
                            })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.deleteCase = function (id) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cases;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.isLocal) {
                            cases = this.getFromLocal('pact_cases').filter(function (c) { return c.id !== id; });
                            this.saveToLocal('pact_cases', cases);
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.COMPLIANCE_CASES, "')/items(").concat(id, ")"), {
                                method: 'POST',
                                headers: {
                                    'X-HTTP-Method': 'DELETE',
                                    'IF-MATCH': '*'
                                }
                            })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.getDashboardStats = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var cases, escalations, trackers, appeals, error_4, cases, escalations, trackers, appeals;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        return [4 /*yield*/, this.getCases()];
                    case 1:
                        cases = _a.sent();
                        return [4 /*yield*/, this.getEscalationLog()];
                    case 2:
                        escalations = _a.sent();
                        return [4 /*yield*/, this.getRepeatTrackerRecords()];
                    case 3:
                        trackers = _a.sent();
                        return [4 /*yield*/, this.getAppeals()];
                    case 4:
                        appeals = _a.sent();
                        return [2 /*return*/, this.buildDashboardStats(cases, escalations, trackers, appeals)];
                    case 5:
                        error_4 = _a.sent();
                        console.warn('Dashboard stats fell back to local data', error_4);
                        cases = this.getFromLocal('pact_cases');
                        escalations = this.getFromLocal('pact_escalations');
                        trackers = this.getFromLocal('pact_trackers');
                        appeals = this.getFromLocal('pact_appeals');
                        return [2 /*return*/, this.buildDashboardStats(cases, escalations, trackers, appeals)];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.buildDashboardStats = function (cases, escalations, trackers, appeals) {
        if (appeals === void 0) { appeals = []; }
        var tiers = ['Tier 1', 'Tier 2', 'Tier 3'];
        var casesByTier = tiers.map(function (tier) { return ({
            tier: tier,
            count: cases.filter(function (c) { var _a; return (_a = c.offenceCategoryName) === null || _a === void 0 ? void 0 : _a.includes(tier); }).length
        }); });
        var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        var now = new Date();
        var last6Months = [];
        var _loop_1 = function (i) {
            var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            var m = months[d.getMonth()];
            var monthlyCases = cases.filter(function (c) {
                var cd = new Date(c.dateCreated);
                return cd.getMonth() === d.getMonth() && cd.getFullYear() === d.getFullYear();
            });
            last6Months.push({
                month: m,
                tier1: monthlyCases.filter(function (c) { var _a; return (_a = c.offenceCategoryName) === null || _a === void 0 ? void 0 : _a.includes('Tier 1'); }).length,
                tier2: monthlyCases.filter(function (c) { var _a; return (_a = c.offenceCategoryName) === null || _a === void 0 ? void 0 : _a.includes('Tier 2'); }).length,
                tier3: monthlyCases.filter(function (c) { var _a; return (_a = c.offenceCategoryName) === null || _a === void 0 ? void 0 : _a.includes('Tier 3'); }).length,
            });
        };
        for (var i = 5; i >= 0; i--) {
            _loop_1(i);
        }
        return {
            totalActiveCases: cases.length,
            escalationsThisMonth: escalations.filter(function (e) {
                var date = new Date(e.escalationDate);
                return date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
            }).length,
            pendingAppeals: appeals.filter(function (appeal) {
                var decision = String((appeal === null || appeal === void 0 ? void 0 : appeal.decision) || '').toLowerCase();
                return !decision || decision === 'pending';
            }).length,
            repeatOffenders: trackers.filter(function (t) { return t.totalOffences > 1; }).length,
            totalFines: cases.reduce(function (sum, c) { return sum + c.penaltyAmount; }, 0),
            casesByTier: casesByTier,
            casesByMonth: last6Months,
            casesByDepartment: Array.from(new Set(cases.map(function (c) { return c.department; }))).map(function (dept) { return ({
                department: dept,
                count: cases.filter(function (c) { return c.department === dept; }).length,
                risk: trackers.some(function (t) {
                    var staff = cases.find(function (c) { return c.department === dept && c.chargedPerson === t.offender; });
                    return staff && (t.riskLevel === 'High' || t.riskLevel === 'Critical');
                }) ? 'High' : 'Low'
            }); }),
            recentActivity: tslib_1.__spreadArray(tslib_1.__spreadArray([], cases.slice(0, 3).map(function (c) {
                var _a, _b;
                return ({
                    id: "c-".concat(c.id),
                    type: 'case',
                    title: "New Case: ".concat(c.title),
                    description: "".concat(c.chargedPersonName, " charged with ").concat(c.offenceCategoryName),
                    timestamp: c.dateCreated,
                    severity: (((_a = c.offenceCategoryName) === null || _a === void 0 ? void 0 : _a.includes('Tier 3')) ? 'critical' : ((_b = c.offenceCategoryName) === null || _b === void 0 ? void 0 : _b.includes('Tier 2')) ? 'high' : 'medium')
                });
            }), true), escalations.slice(0, 2).map(function (e) { return ({
                id: "e-".concat(e.id),
                type: 'escalation',
                title: "Escalation: ".concat(e.caseReference),
                description: "Tier upgrade due to ".concat(e.escalationReason),
                timestamp: e.escalationDate,
                severity: 'critical'
            }); }), true).sort(function (a, b) { return new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime(); })
        };
    };
    // --- Escalations & Tracker ---
    SharePointService.prototype.getEscalationLog = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var endpoint, data, _a;
            var _this = this;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        if (!!this.isLocal) return [3 /*break*/, 2];
                        endpoint = "web/lists/getbytitle('".concat(constants_1.LIST_NAMES.ESCALATION_LOG, "')/items?$orderby=ID desc");
                        return [4 /*yield*/, this.fetchREST(endpoint)];
                    case 1:
                        data = _b.sent();
                        return [2 /*return*/, (data.results || []).map(function (item) {
                                var _a, _b;
                                return ({
                                    id: item.ID.toString(),
                                    title: _this.readField(item, constants_1.COLUMNS.ESCALATION.TITLE, 'Title', 'Escalation ID'),
                                    caseReference: _this.readField(item, constants_1.COLUMNS.ESCALATION.CASE_REFERENCE, 'Case Reference', 'CaseReference', 'Case_x0020_Reference'),
                                    offender: ((_b = (_a = _this.readField(item, constants_1.COLUMNS.ESCALATION.OFFENDER, 'Offender', 'OffenderId')) === null || _a === void 0 ? void 0 : _a.toString) === null || _b === void 0 ? void 0 : _b.call(_a)) || String(_this.readField(item, constants_1.COLUMNS.ESCALATION.OFFENDER, 'Offender', 'OffenderId') || ''),
                                    offenderName: _this.readField(item, constants_1.COLUMNS.ESCALATION.OFFENDER, 'Offender', 'OffenderId'),
                                    escalationReason: _this.readField(item, constants_1.COLUMNS.ESCALATION.REASON, 'Escalation Reason', 'EscalationReason', 'Escalation_x0020_Reason'),
                                    previousTier: _this.readField(item, constants_1.COLUMNS.ESCALATION.PREVIOUS_TIER, 'Previous Tier', 'PreviousTier', 'Previous_x0020_Tier'),
                                    newTier: _this.readField(item, constants_1.COLUMNS.ESCALATION.NEW_TIER, 'New Tier', 'NewTier', 'New_x0020_Tier'),
                                    triggeredBy: _this.readField(item, constants_1.COLUMNS.ESCALATION.TRIGGERED_BY, 'Triggered By', 'TriggeredBy', 'Triggered_x0020_By'),
                                    escalationDate: _this.readField(item, constants_1.COLUMNS.ESCALATION.DATE, 'Escalation Date', 'EscalationDate', 'Escalation_x0020_Date'),
                                    notifiedTo: _this.readField(item, constants_1.COLUMNS.ESCALATION.NOTIFIED_TO, 'Notified To', 'NotifiedTo', 'Notified_x0020_To')
                                });
                            })];
                    case 2: return [2 /*return*/, this.getFromLocal('pact_escalations')];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, this.getFromLocal('pact_escalations')];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.getRepeatTrackerRecord = function (staffId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var records_1, records, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        if (!!this.isLocal) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.getRepeatTrackerRecords()];
                    case 1:
                        records_1 = _b.sent();
                        return [2 /*return*/, records_1.find(function (r) { return r.offender === staffId; }) || null];
                    case 2:
                        records = this.getFromLocal('pact_trackers');
                        return [2 /*return*/, records.find(function (r) { return r.offender === staffId; }) || null];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.getRepeatTrackerRecords = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var endpoint, data, error_5;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.isLocal) {
                            return [2 /*return*/, this.getFromLocal('pact_trackers')];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        endpoint = "web/lists/getbytitle('".concat(constants_1.LIST_NAMES.REPEAT_OFFENCE_TRACKER, "')/items?$orderby=ID desc");
                        return [4 /*yield*/, this.fetchREST(endpoint)];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, (data.results || []).map(function (item) {
                                var _a, _b;
                                return ({
                                    id: item.ID.toString(),
                                    title: _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.TITLE, 'Title'),
                                    offender: ((_b = (_a = _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.OFFENDER, 'Offender', 'OffenderId')) === null || _a === void 0 ? void 0 : _a.toString) === null || _b === void 0 ? void 0 : _b.call(_a)) || String(_this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.OFFENDER, 'Offender', 'OffenderId') || ''),
                                    offenderName: _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.TITLE, 'Title'),
                                    totalOffences: Number(_this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES, 'TotalOffences', 'Total Offences') || 0),
                                    tier1Last6Months: Number(_this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M, 'Tier1Offences', 'Tier1 Offences') || 0),
                                    tier2Offences: Number(_this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES, 'Tier2Offences', 'Tier2 Offences') || 0),
                                    tier3Offences: Number(_this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES, 'Tier3Offences', 'Tier3 Offences') || 0),
                                    riskLevel: _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.RISK_LEVEL, 'RiskLevel', 'Risk Level') || 'Low',
                                    lastOffenceDate: _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.LAST_OFFENCE_DATE, 'LastOffenceDate', 'Last Offence Date') || new Date().toISOString(),
                                    escalationDue: _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.ESCALATION_DUE, 'EscalationDue', 'Escalation Due') === true || _this.readField(item, constants_1.COLUMNS.REPEAT_TRACKER.ESCALATION_DUE, 'EscalationDue', 'Escalation Due') === 'True'
                                });
                            })];
                    case 3:
                        error_5 = _a.sent();
                        console.warn('Failed to load live repeat tracker records', error_5);
                        return [2 /*return*/, this.getFromLocal('pact_trackers')];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.updateRepeatTracker = function (staffId, updates) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var records, idx, staff, spData, existing, staff, person, createData, e_1;
            var _a, _b;
            var _c, _d, _e, _f, _g, _h, _j;
            return tslib_1.__generator(this, function (_k) {
                switch (_k.label) {
                    case 0:
                        if (!this.isLocal) return [3 /*break*/, 1];
                        records = this.getFromLocal('pact_trackers');
                        idx = records.findIndex(function (r) { return r.offender === staffId; });
                        if (idx > -1) {
                            records[idx] = tslib_1.__assign(tslib_1.__assign({}, records[idx]), updates);
                        }
                        else {
                            staff = this.getFromLocal('pact_staff').find(function (s) { return s.id === staffId; });
                            records.push(tslib_1.__assign({ id: Date.now().toString(), title: (staff === null || staff === void 0 ? void 0 : staff.fullName) || 'Unknown', offender: staffId, totalOffences: 1, tier1Last6Months: 1, tier2Offences: 0, tier3Offences: 0, riskLevel: 'Low', lastOffenceDate: new Date().toISOString(), escalationDue: false }, updates));
                        }
                        this.saveToLocal('pact_trackers', records);
                        return [3 /*break*/, 10];
                    case 1:
                        spData = (_a = {},
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES] = updates.totalOffences,
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M] = updates.tier1Last6Months,
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES] = updates.tier2Offences,
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES] = updates.tier3Offences,
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.RISK_LEVEL] = updates.riskLevel,
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.LAST_OFFENCE_DATE] = updates.lastOffenceDate,
                            _a[constants_1.COLUMNS.REPEAT_TRACKER.ESCALATION_DUE] = updates.escalationDue,
                            _a);
                        _k.label = 2;
                    case 2:
                        _k.trys.push([2, 9, , 10]);
                        return [4 /*yield*/, this.getRepeatTrackerRecord(staffId)];
                    case 3:
                        existing = _k.sent();
                        if (!existing) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.REPEAT_OFFENCE_TRACKER, "')/items(").concat(existing.id, ")"), {
                                method: 'POST',
                                headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
                                body: JSON.stringify(spData)
                            })];
                    case 4:
                        _k.sent();
                        return [3 /*break*/, 8];
                    case 5: return [4 /*yield*/, this.getStaffDirectory()];
                    case 6:
                        staff = _k.sent();
                        person = staff.find(function (s) { return s.id === staffId; });
                        createData = (_b = {
                                '__metadata': { 'type': "SP.Data.".concat(constants_1.LIST_NAMES.REPEAT_OFFENCE_TRACKER.replace(/ /g, '_x0020_'), "ListItem") }
                            },
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.TITLE] = (person === null || person === void 0 ? void 0 : person.fullName) || 'Unknown',
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.OFFENDER] = staffId,
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES] = (_c = updates.totalOffences) !== null && _c !== void 0 ? _c : 1,
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M] = (_d = updates.tier1Last6Months) !== null && _d !== void 0 ? _d : 0,
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES] = (_e = updates.tier2Offences) !== null && _e !== void 0 ? _e : 0,
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES] = (_f = updates.tier3Offences) !== null && _f !== void 0 ? _f : 0,
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.RISK_LEVEL] = (_g = updates.riskLevel) !== null && _g !== void 0 ? _g : 'Low',
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.LAST_OFFENCE_DATE] = (_h = updates.lastOffenceDate) !== null && _h !== void 0 ? _h : new Date().toISOString(),
                            _b[constants_1.COLUMNS.REPEAT_TRACKER.ESCALATION_DUE] = (_j = updates.escalationDue) !== null && _j !== void 0 ? _j : false,
                            _b);
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.REPEAT_OFFENCE_TRACKER, "')/items"), {
                                method: 'POST',
                                body: JSON.stringify(createData)
                            })];
                    case 7:
                        _k.sent();
                        _k.label = 8;
                    case 8: return [3 /*break*/, 10];
                    case 9:
                        e_1 = _k.sent();
                        console.error("Failed to update SP Repeat Tracker:", e_1);
                        return [3 /*break*/, 10];
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.createEscalation = function (entry) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var log, spData, subject, body;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.isLocal) return [3 /*break*/, 1];
                        log = this.getFromLocal('pact_escalations');
                        log.push({
                            id: Date.now().toString(),
                            title: "Escalation ".concat(entry.caseReference),
                            caseReference: entry.caseReference || '',
                            offender: entry.offender || '',
                            offenderName: entry.offenderName || '',
                            escalationReason: entry.escalationReason || '',
                            previousTier: entry.previousTier || 'Tier 1',
                            newTier: entry.newTier || 'Tier 2',
                            triggeredBy: 'System',
                            escalationDate: new Date().toISOString(),
                            notifiedTo: 'HR Manager'
                        });
                        this.saveToLocal('pact_escalations', log);
                        return [3 /*break*/, 3];
                    case 1:
                        spData = (_a = {
                                '__metadata': { 'type': "SP.Data.".concat(constants_1.LIST_NAMES.ESCALATION_LOG.replace(/ /g, '_x0020_'), "ListItem") }
                            },
                            _a[constants_1.COLUMNS.ESCALATION.TITLE] = entry.title || "Escalation ".concat(entry.caseReference),
                            _a[constants_1.COLUMNS.ESCALATION.CASE_REFERENCE] = entry.caseReference,
                            _a[constants_1.COLUMNS.ESCALATION.OFFENDER] = entry.offender,
                            _a[constants_1.COLUMNS.ESCALATION.REASON] = entry.escalationReason,
                            _a[constants_1.COLUMNS.ESCALATION.PREVIOUS_TIER] = entry.previousTier,
                            _a[constants_1.COLUMNS.ESCALATION.NEW_TIER] = entry.newTier,
                            _a[constants_1.COLUMNS.ESCALATION.TRIGGERED_BY] = entry.triggeredBy || 'System',
                            _a[constants_1.COLUMNS.ESCALATION.DATE] = entry.escalationDate || new Date().toISOString(),
                            _a[constants_1.COLUMNS.ESCALATION.NOTIFIED_TO] = entry.notifiedTo || 'HR Dept',
                            _a);
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.ESCALATION_LOG, "')/items"), {
                                method: 'POST',
                                body: JSON.stringify(spData)
                            })];
                    case 2:
                        _b.sent();
                        _b.label = 3;
                    case 3:
                        subject = "PACT ESCALATION: ".concat(entry.caseReference, " - TIER UPGRADE");
                        body = "\n      <div style=\"font-family: Arial, sans-serif; color: #333; max-width: 600px; padding: 20px; border: 2px solid #d13438; border-radius: 8px;\">\n        <h2 style=\"color: #d13438; margin-top: 0;\">Policy Escalation Triggered</h2>\n        <p>A compliance case has been automatically escalated due to repeat offences or severity.</p>\n        <div style=\"background: #fff4f4; padding: 15px; border-radius: 4px; margin: 20px 0;\">\n          <p><b>Case Ref:</b> ".concat(entry.caseReference, "</p>\n          <p><b>Previous Tier:</b> ").concat(entry.previousTier, "</p>\n          <p><b>New Tier:</b> ").concat(entry.newTier, "</p>\n          <p><b>Reason:</b> ").concat(entry.escalationReason, "</p>\n        </div>\n        <p>Immediate review by HR / Executive Management is required.</p>\n      </div>\n    ");
                        return [4 /*yield*/, this.sendEmailNotification(['hr@konstructum.com'], subject, body)];
                    case 4:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    // --- Disciplinary Actions ---
    SharePointService.prototype.getDisciplinaryActions = function (caseRef) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var filter, endpoint, data, actions, _a, actions;
            var _this = this;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        if (!!this.isLocal) return [3 /*break*/, 2];
                        filter = caseRef ? "&$filter=".concat(constants_1.COLUMNS.DISCIPLINARY.CASE_REFERENCE, " eq '").concat(caseRef, "'") : '';
                        endpoint = "web/lists/getbytitle('".concat(constants_1.LIST_NAMES.DISCIPLINARY_ACTIONS, "')/items?$orderby=ID desc").concat(filter);
                        return [4 /*yield*/, this.fetchREST(endpoint)];
                    case 1:
                        data = _b.sent();
                        return [2 /*return*/, (data.results || []).map(function (item) { return ({
                                id: item.ID.toString(),
                                title: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.TITLE, 'Title'),
                                caseReference: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.CASE_REFERENCE, 'Case Reference', 'CaseReference', 'Case_x0020_Reference'),
                                actionType: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.ACTION_TYPE, 'Action Type', 'ActionType', 'Action_x0020_Type'),
                                actionDate: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.ACTION_DATE, 'Action Date', 'ActionDate', 'Action_x0020_Date'),
                                actionedBy: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.ACTIONED_BY, 'Actioned By', 'ActionedBy', 'Actioned_x0020_By'),
                                penaltyAmount: _this.parsePenalty(_this.readField(item, 'Penalty Amount', 'PenaltyAmount', 'Penalty_x0020_Amount')),
                                notes: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.NOTES, 'Notes'),
                                status: _this.readField(item, constants_1.COLUMNS.DISCIPLINARY.STATUS, 'Status')
                            }); })];
                    case 2:
                        actions = this.getFromLocal('pact_disciplinary');
                        return [2 /*return*/, caseRef ? actions.filter(function (a) { return a.caseReference === caseRef; }) : actions];
                    case 3:
                        _a = _b.sent();
                        actions = this.getFromLocal('pact_disciplinary');
                        return [2 /*return*/, caseRef ? actions.filter(function (a) { return a.caseReference === caseRef; }) : actions];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.createDisciplinaryAction = function (action) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var log, spData;
            var _a;
            var _b;
            return tslib_1.__generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (this.isLocal) {
                            log = this.getFromLocal('pact_disciplinary');
                            log.push(tslib_1.__assign({ id: Date.now().toString(), title: action.title || this.formatDisciplinaryReference(log.length + 1), actionDate: new Date().toISOString(), actionedBy: this.getUserName() }, action));
                            this.saveToLocal('pact_disciplinary', log);
                            return [2 /*return*/];
                        }
                        spData = (_a = {
                                '__metadata': { 'type': "SP.Data.".concat(constants_1.LIST_NAMES.DISCIPLINARY_ACTIONS.replace(/ /g, '_x0020_'), "ListItem") }
                            },
                            _a[constants_1.COLUMNS.DISCIPLINARY.TITLE] = action.title || "P-".concat(Date.now()),
                            _a[constants_1.COLUMNS.DISCIPLINARY.CASE_REFERENCE] = action.caseReference,
                            _a[constants_1.COLUMNS.DISCIPLINARY.ACTION_TYPE] = action.actionType,
                            _a[constants_1.COLUMNS.DISCIPLINARY.ACTION_DATE] = new Date().toISOString(),
                            _a[constants_1.COLUMNS.DISCIPLINARY.ACTIONED_BY] = this.getUserName(),
                            _a['Penalty_x0020_Amount'] = (_b = action.penaltyAmount) !== null && _b !== void 0 ? _b : 0,
                            _a[constants_1.COLUMNS.DISCIPLINARY.NOTES] = action.notes,
                            _a[constants_1.COLUMNS.DISCIPLINARY.STATUS] = action.status || 'Pending',
                            _a);
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.DISCIPLINARY_ACTIONS, "')/items"), {
                                method: 'POST',
                                body: JSON.stringify(spData)
                            })];
                    case 1:
                        _c.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    // --- Appeals ---
    SharePointService.prototype.getAppeals = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var endpoint, data, _a;
            var _this = this;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        if (!!this.isLocal) return [3 /*break*/, 2];
                        endpoint = "web/lists/getbytitle('".concat(constants_1.LIST_NAMES.APPEALS_REGISTER, "')/items?$orderby=ID desc");
                        return [4 /*yield*/, this.fetchREST(endpoint)];
                    case 1:
                        data = _b.sent();
                        return [2 /*return*/, (data.results || []).map(function (item) { return ({
                                id: item.ID.toString(),
                                title: _this.readField(item, constants_1.COLUMNS.APPEALS.TITLE, 'Title', 'Appeal ID'),
                                caseReference: _this.readField(item, constants_1.COLUMNS.APPEALS.CASE_REFERENCE, 'Case Reference', 'CaseReference', 'Case_x0020_Reference'),
                                appellant: _this.readField(item, constants_1.COLUMNS.APPEALS.APPELLANT, 'Appellant Name', 'AppellantName', 'Appellant'),
                                appealDate: _this.readField(item, constants_1.COLUMNS.APPEALS.APPEAL_DATE, 'Appeal Date', 'AppealDate', 'Appeal_x0020_Date'),
                                grounds: _this.readField(item, constants_1.COLUMNS.APPEALS.GROUNDS, 'Grounds for Appeal', 'GroundsforAppeal', 'Grounds_x0020_for_x0020_Appeal'),
                                reviewingOfficer: _this.readField(item, constants_1.COLUMNS.APPEALS.REVIEWING_OFFICER, 'Reviewing Officer', 'ReviewingOfficer', 'Reviewing_x0020_Officer'),
                                decision: _this.readField(item, constants_1.COLUMNS.APPEALS.DECISION, 'Decision'),
                                decisionDate: _this.readField(item, constants_1.COLUMNS.APPEALS.DECISION_DATE, 'Decision Date', 'DecisionDate', 'Decision_x0020_Date'),
                                decisionNotes: _this.readField(item, constants_1.COLUMNS.APPEALS.DECISION_NOTES, 'Decision Notes', 'DecisionNotes', 'Decision_x0020_Notes')
                            }); })];
                    case 2: return [2 /*return*/, this.getFromLocal('pact_appeals')];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, this.getFromLocal('pact_appeals')];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.getMailHistory = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var endpoint, data, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        if (!!this.isLocal) return [3 /*break*/, 2];
                        endpoint = "web/lists/getbytitle('".concat(constants_1.LIST_NAMES.MAIL_HISTORY, "')/items?$orderby=Created desc&$top=50");
                        return [4 /*yield*/, this.fetchREST(endpoint)];
                    case 1:
                        data = _b.sent();
                        return [2 /*return*/, (data.results || []).map(function (item) { return ({
                                id: item.ID.toString(),
                                to: (item[constants_1.COLUMNS.MAIL.TO] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean),
                                subject: item[constants_1.COLUMNS.MAIL.SUBJECT] || item[constants_1.COLUMNS.MAIL.TITLE],
                                body: item[constants_1.COLUMNS.MAIL.BODY],
                                timestamp: item.Created,
                                status: item[constants_1.COLUMNS.MAIL.STATUS] || 'Sent'
                            }); })];
                    case 2: return [2 /*return*/, this.getFromLocal('pact_mail_history') || []];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/, this.getFromLocal('pact_mail_history') || []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.createAppeal = function (appeal) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var log, spData, subject, body;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.isLocal) return [3 /*break*/, 1];
                        log = this.getFromLocal('pact_appeals');
                        log.push(tslib_1.__assign({ id: Date.now().toString(), title: appeal.title || "A-".concat(String(log.length + 1).padStart(3, '0')), appealDate: new Date().toISOString(), decision: 'Pending' }, appeal));
                        this.saveToLocal('pact_appeals', log);
                        return [3 /*break*/, 3];
                    case 1:
                        spData = (_a = {
                                '__metadata': { 'type': "SP.Data.".concat(constants_1.LIST_NAMES.APPEALS_REGISTER.replace(/ /g, '_x0020_'), "ListItem") },
                                Title: appeal.title || "A-".concat(String(Date.now()).slice(-4))
                            },
                            _a[constants_1.COLUMNS.APPEALS.CASE_REFERENCE] = appeal.caseReference,
                            _a[constants_1.COLUMNS.APPEALS.APPELLANT] = appeal.appellant,
                            _a[constants_1.COLUMNS.APPEALS.APPEAL_DATE] = new Date().toISOString(),
                            _a[constants_1.COLUMNS.APPEALS.GROUNDS] = appeal.grounds,
                            _a[constants_1.COLUMNS.APPEALS.DECISION] = 'Pending',
                            _a[constants_1.COLUMNS.APPEALS.DECISION_NOTES] = appeal.decisionNotes || '',
                            _a);
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.APPEALS_REGISTER, "')/items"), {
                                method: 'POST',
                                body: JSON.stringify(spData)
                            })];
                    case 2:
                        _b.sent();
                        _b.label = 3;
                    case 3:
                        subject = "PACT APPEAL FILED: Case ".concat(appeal.caseReference);
                        body = "<p>An appeal has been filed by <b>".concat(appeal.appellant, "</b> for Case <b>").concat(appeal.caseReference, "</b>.</p><p>Grounds: ").concat(appeal.grounds, "</p>");
                        return [4 /*yield*/, this.sendEmailNotification(['legal@konstructum.com'], subject, body)];
                    case 4:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.updateAppeal = function (id, updates) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var log, idx, appeal, subject, body, spData;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.isLocal) return [3 /*break*/, 3];
                        log = this.getFromLocal('pact_appeals');
                        idx = log.findIndex(function (a) { return a.id === id; });
                        if (!(idx > -1)) return [3 /*break*/, 2];
                        log[idx] = tslib_1.__assign(tslib_1.__assign({}, log[idx]), updates);
                        this.saveToLocal('pact_appeals', log);
                        if (!updates.decision) return [3 /*break*/, 2];
                        appeal = log[idx];
                        subject = "PACT APPEAL DECISION: ".concat(appeal.caseReference);
                        body = "<p>Your appeal for <b>".concat(appeal.caseReference, "</b> has been <b>").concat(updates.decision, "</b>.</p><p>Notes: ").concat(updates.decisionNotes, "</p>");
                        return [4 /*yield*/, this.sendEmailNotification([appeal.appellantEmail || 'staff@konstructum.com'], subject, body)];
                    case 1:
                        _b.sent();
                        _b.label = 2;
                    case 2: return [2 /*return*/];
                    case 3:
                        spData = (_a = {},
                            _a[constants_1.COLUMNS.APPEALS.DECISION] = updates.decision,
                            _a[constants_1.COLUMNS.APPEALS.DECISION_DATE] = new Date().toISOString(),
                            _a[constants_1.COLUMNS.APPEALS.DECISION_NOTES] = updates.decisionNotes,
                            _a[constants_1.COLUMNS.APPEALS.REVIEWING_OFFICER] = updates.reviewingOfficer,
                            _a);
                        return [4 /*yield*/, this.fetchREST("web/lists/getbytitle('".concat(constants_1.LIST_NAMES.APPEALS_REGISTER, "')/items(").concat(id, ")"), {
                                method: 'POST',
                                headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
                                body: JSON.stringify(spData)
                            })];
                    case 4:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    SharePointService.prototype.mapSPItemToCase = function (item, staff, policies) {
        if (staff === void 0) { staff = []; }
        if (policies === void 0) { policies = []; }
        var chargedPersonId = String(this.readField(item, constants_1.COLUMNS.CASES.CHARGED_PERSON, 'ChargedPerson', 'Charged Persaon', 'Charged Person', 'Charged Person ') || '');
        var offenceCategoryId = String(this.readField(item, constants_1.COLUMNS.CASES.OFFENCE_CATEGORY, 'OffenceCategory', 'Offence Category', 'Offence_x0020_Category') || '');
        var chargedPerson = staff.find(function (s) { return s.id === chargedPersonId; });
        var policy = policies.find(function (p) { return p.id === offenceCategoryId; });
        return {
            id: item.ID.toString(),
            title: this.readField(item, constants_1.COLUMNS.CASES.TITLE, 'Penalty ID', 'Case ID', 'CaseID', 'Title'),
            chargedPerson: chargedPersonId,
            chargedPersonName: (chargedPerson === null || chargedPerson === void 0 ? void 0 : chargedPerson.fullName) || this.readField(item, constants_1.COLUMNS.CASES.CHARGED_PERSON, 'ChargedPerson', 'Charged Persaon', 'Charged Person', 'Charged Person ') || '',
            staffEmail: this.readField(item, constants_1.COLUMNS.CASES.STAFF_EMAIL, 'StaffEmail', 'Email', 'Charged Person Email', 'Staff Email'),
            department: this.readField(item, constants_1.COLUMNS.CASES.DEPARTMENT, 'Department'),
            offenceCategory: offenceCategoryId,
            offenceCategoryName: policy ? this.expandAbbreviations(policy.offenceName) : this.readField(item, constants_1.COLUMNS.CASES.OFFENCE_CATEGORY, 'OffenceCategory', 'Offence Category', 'Offence_x0020_Category'),
            offenceDescription: this.readField(item, constants_1.COLUMNS.CASES.OFFENCE_DESCRIPTION, 'OffenceDescription', 'Offence Description', 'Offence_x0020_Description'),
            penaltyAmount: this.parsePenalty(this.readField(item, constants_1.COLUMNS.CASES.PENALTY_AMOUNT, 'PenaltyAmount')),
            dueDate: this.readField(item, constants_1.COLUMNS.CASES.DUE_DATE, 'DueDate', 'Due Date', 'Due_x0020_Date'),
            issuerName: this.readField(item, constants_1.COLUMNS.CASES.ISSUER_NAME, 'IssuerName', 'Issuer Name', 'Issuer_x0020_Name'),
            secondaryContact: this.readField(item, constants_1.COLUMNS.CASES.SECONDARY_CONTACT, 'SecondaryContact', 'Secondary Contact Email'),
            status: this.readField(item, constants_1.COLUMNS.CASES.STATUS, 'Status'),
            dateCreated: this.readField(item, constants_1.COLUMNS.CASES.DATE_CREATED, 'Created', 'Date Created', 'Date Created ')
        };
    };
    SharePointService.prototype.getInitialMockCases = function () {
        var now = new Date();
        var lastMonth = new Date();
        lastMonth.setMonth(now.getMonth() - 1);
        var twoMonthsAgo = new Date();
        twoMonthsAgo.setMonth(now.getMonth() - 2);
        return [
            {
                id: '1', title: 'PACT-1001', chargedPerson: '17', chargedPersonName: 'Ugeh Collins',
                staffEmail: 'ccugeh.konstructum@outlook.com', department: 'Kadet', offenceCategory: 'p4',
                offenceCategoryName: 'Late to Work/Site', offenceDescription: 'Arrived 2 hours late without notice.',
                penaltyAmount: 5000, dueDate: new Date(now.getTime() + 5 * 86400000).toISOString(), issuerName: 'System Admin',
                secondaryContact: 'Victor Ochi', status: 'Unpaid', dateCreated: now.toISOString()
            },
            {
                id: '2', title: 'PACT-0982', chargedPerson: '16', chargedPersonName: 'Janet Afolabi',
                staffEmail: 'jafolabi@konstructum.com', department: 'Procurement', offenceCategory: 'p9',
                offenceCategoryName: 'Procurement Policy Breach', offenceDescription: 'Approved vendor without completing review process.',
                penaltyAmount: 5000, dueDate: new Date(now.getTime() - 2 * 86400000).toISOString(), issuerName: 'System Admin',
                secondaryContact: '', status: 'Overdue', dateCreated: lastMonth.toISOString()
            },
            {
                id: '3', title: 'PACT-0945', chargedPerson: '12', chargedPersonName: 'Victor Ochi',
                staffEmail: 'vochi@konstructum.com', department: 'Engineering', offenceCategory: 'p25',
                offenceCategoryName: 'Unauthorized Entry or Bringing Unauthorized Persons', offenceDescription: 'Allowed non-staff on site without PTW.',
                penaltyAmount: 5000, dueDate: new Date(now.getTime() - 15 * 86400000).toISOString(), issuerName: 'System Admin',
                secondaryContact: 'Abayomi Awobokun', status: 'Paid', dateCreated: twoMonthsAgo.toISOString()
            },
            {
                id: '4', title: 'PACT-1015', chargedPerson: '17', chargedPersonName: 'Ugeh Collins',
                staffEmail: 'ccugeh.konstructum@outlook.com', department: 'Kadet', offenceCategory: 'p2',
                offenceCategoryName: 'Dress Policy Contravention', offenceDescription: 'Not wearing proper safety gear on site.',
                penaltyAmount: 5000, dueDate: new Date(now.getTime() + 8 * 86400000).toISOString(), issuerName: 'System Admin',
                secondaryContact: 'Victor Ochi', status: 'Unpaid', dateCreated: new Date(now.getTime() - 86400000).toISOString()
            },
            {
                id: '5', title: 'PACT-1008', chargedPerson: '4', chargedPersonName: 'Ayomide Popoola',
                staffEmail: 'apopoola@konstructum.com', department: 'Procurement', offenceCategory: 'p11',
                offenceCategoryName: 'Confidentiality Breach', offenceDescription: 'Shared pricing index externally.',
                penaltyAmount: 5000, dueDate: new Date(now.getTime() + 10 * 86400000).toISOString(), issuerName: 'System Admin',
                secondaryContact: 'Mojisola Coker', status: 'Unpaid', dateCreated: new Date(now.getTime() - 2 * 86400000).toISOString()
            }
        ];
    };
    SharePointService.prototype.getInitialMockTrackers = function () {
        return []; // Start fresh with new IDs
    };
    SharePointService.prototype.getInitialMockEscalations = function () {
        var now = new Date();
        var threeDaysAgo = new Date(now.getTime() - 3 * 86400000);
        return [
            {
                id: "esc1", title: "Escalation PACT-0921", caseReference: "PACT-0921",
                offender: "8", offenderName: "Babatunde Adeleye",
                escalationReason: "Automatic Policy Trigger: Staff member reached 3 Tier 1 offences within 6 months. Threshold exceeded.",
                previousTier: "Tier 1", newTier: "Tier 2", triggeredBy: "System",
                escalationDate: threeDaysAgo.toISOString(), notifiedTo: "HR Manager"
            }
        ];
    };
    SharePointService.prototype.getInitialMockStaff = function () {
        return staffDirectory_json_1.default;
    };
    SharePointService.prototype.getInitialMockPolicies = function () {
        return policyLibrary_json_1.default;
    };
    SharePointService.prototype.getInitialMockAppeals = function () {
        return [
            { id: '1', caseReference: 'PACT-0945', appellant: 'Victor Ochi', appealDate: new Date().toISOString(), grounds: 'Penalty was applied to wrong department.', decision: 'Pending', reviewingOfficer: 'HR Dept' }
        ];
    };
    SharePointService.prototype.getInitialMockDisciplinary = function () {
        return [
            { id: '1', title: 'P-001', caseReference: 'PACT-0945', actionType: 'Written Warning + CBT', actionDate: '2026-04-10T10:00:00Z', actionedBy: 'PACT Admin', status: 'Enforced', notes: 'Legacy seeded example' }
        ];
    };
    SharePointService._instance = null;
    return SharePointService;
}());
exports.SharePointService = SharePointService;
exports.sharePointService = new SharePointService();
//# sourceMappingURL=SharePointService.js.map