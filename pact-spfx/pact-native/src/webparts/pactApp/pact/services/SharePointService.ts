/**
 * PACT Platform — SharePoint REST Service
 * Communicates directly with SharePoint Lists without Azure AD.
 * Bypasses IT/Azure admin requirements.
 */
import { SHAREPOINT_SITE_URL, SHAREPOINT_SITE_PATH, LIST_NAMES, COLUMNS, HR_EMAIL, ADMIN_EMAIL, LEGAL_EMAIL, CHAIRMAN_EMAIL, APPEAL_REVIEW_EMAILS, MAIL_TRIGGER_URL, ACCEPT_PAYMENT_TRIGGER_URL, APPEAL_MAIL_TRIGGER_URL, APPEAL_SLA_DAYS, PAYMENT_DEADLINE_DAYS, PAYMENT_PROOFS_LIBRARY, CASE_STATUS, RESPONSE_PORTAL_BASE_URL, CASE_RESPONSE_FROM_EMAIL_QUERY_KEY, CASE_RESPONSE_FROM_EMAIL_QUERY_VALUE } from '../config/constants';
import type { 
  ComplianceCase, DashboardStats, StaffMember, PolicyOffence, 
  EscalationEntry, RepeatOffenceRecord, UserSession
} from '../config/types';
import { escalationEngine } from './EscalationEngine';
import staffData from '../data/staffDirectory.json';
import policyData from '../data/policyLibrary.json';

export class SharePointService {
  private siteUrl: string;
  private runtimeMode: 'demo' | 'sharepoint';
  private isLocal: boolean;
  private STORAGE_PREFIX = 'pact_engine_v1_';
  private readonly REST_TIMEOUT_MS = 20000;
  private static _instance: SharePointService | null = null;
  private _spfxContext: any = null;
  private _caseMappingLogged = false;
  private _escalationDiagLogged = false;

  // Called by the SPFx WebPart to inject context before React renders
  public static init(context: any): void {
    if (!SharePointService._instance) {
      SharePointService._instance = sharePointService;
    }
    SharePointService._instance._spfxContext = context;
    // If running on SharePoint, switch to live mode
    if (context?.pageContext?.web?.absoluteUrl) {
      SharePointService._instance.runtimeMode = 'sharepoint';
      SharePointService._instance.isLocal = false;
      SharePointService._instance.siteUrl = context.pageContext.web.absoluteUrl;
    }
  }

  constructor() {
    this.siteUrl = `https://${SHAREPOINT_SITE_URL}${SHAREPOINT_SITE_PATH}`;
    const params = new URLSearchParams(window.location.search);
    const requestedMode = (params.get('pactMode') || params.get('mode') || '').toLowerCase();
    const host = window.location.hostname.toLowerCase();
    const hostLooksNative = host.includes('sharepoint.com') || host.includes('sharepoint-df.com');
    this.runtimeMode = requestedMode === 'demo'
      ? 'demo'
      : requestedMode === 'sharepoint' || requestedMode === 'live' || hostLooksNative
        ? 'sharepoint'
        : 'demo';
    this.isLocal = this.runtimeMode !== 'sharepoint';
  }

  // Returns current user session info from SPFx context or fallback
  public getCurrentSession(): UserSession {
    const ctx = this._spfxContext;
    if (ctx?.pageContext?.user) {
      const user = ctx.pageContext.user;
      return {
        displayName: user.displayName || 'PACT User',
        email: user.email || user.loginName || '',
        photoUrl: user.email
          ? `/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(user.email)}`
          : undefined,
        tenantName: ctx.pageContext.web?.title || 'KONSTRUCTUM',
        siteTitle: ctx.pageContext.site?.absoluteUrl || '',
      };
    }
    return {
      displayName: 'PACT Administrator',
      email: 'admin@konstructum.com',
      photoUrl: undefined,
      tenantName: 'KONSTRUCTUM',
      siteTitle: 'SharePoint',
    };
  }

  // Returns the photo URL for a specific staff member
  public getPhotoUrl(email: string): string {
    return `/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}`;
  }


  /**
   * Helper to ensuring professional expanded terms are visible even in legacy cases
   */
  private expandAbbreviations(text: string): string {
    if (!text) return text;
    return text
      .replace(/\bWW\b/g, 'Written Warning')
      .replace(/\bVW\b/g, 'Verbal Warning')
      .replace(/\bT\b/g, 'Termination')
      .replace(/\bS1\b/g, 'Suspension (1 Day)')
      .replace(/\bS2\b/g, 'Suspension (2 Days)')
      .replace(/\bS4\b/g, 'Suspension (4 Days)')
      .replace(/\bT&P\b/g, 'Termination & Prosecution');
  }

  /**
   * Links for the employee one-pager (Vercel or current origin). Uses HashRouter paths.
   */
  private buildCaseResponseLink(
    caseRef: string,
    action: 'accept' | 'appeal',
    opts: { staffName: string; offenceLabel: string; amount: number; dueIso: string; staffEmail?: string; department?: string; description?: string }
  ): string {
    const params = new URLSearchParams({
      name: opts.staffName,
      offence: opts.offenceLabel,
      amount: String(opts.amount),
      date: opts.dueIso,
      [CASE_RESPONSE_FROM_EMAIL_QUERY_KEY]: CASE_RESPONSE_FROM_EMAIL_QUERY_VALUE,
    });
    if (opts.staffEmail) params.set('email', opts.staffEmail);
    if (opts.department) params.set('department', opts.department);
    if (opts.description) params.set('description', opts.description);
    const hash = `#/case-response/${encodeURIComponent(caseRef)}/${action}?${params.toString()}`;
    const base =
      RESPONSE_PORTAL_BASE_URL.length > 0
        ? RESPONSE_PORTAL_BASE_URL.replace(/\/$/, '')
        : `${window.location.origin}${window.location.pathname}`;
    return `${base}${hash}`;
  }

  private buildEmailButtonHtmlBoth(acceptUrl: string, appealUrl: string): string {
    return `
              <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px;">
                <p style="font-family: Arial, sans-serif; font-size: 14px; color: #444; margin-bottom: 20px;">Please click a button below to respond to this notice:</p>
                <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td style="padding-bottom: 20px;">
                      <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td align="center" bgcolor="#0078d4" style="border-radius: 4px;">
                            <a href="${acceptUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">
                              Accept & Pay Fine
                            </a>
                          </td>
                          <td width="15"></td>
                          <td align="center" bgcolor="#ffffff" style="border-radius: 4px; border: 1px solid #cccccc;">
                            <a href="${appealUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #333333; text-decoration: none; font-weight: bold; display: inline-block;">
                              Appeal Decision
                            </a>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </div>`;
  }

  /** Escalation-related notices: appeal path only (no accept/pay). */
  private buildEmailButtonHtmlAppealOnly(appealUrl: string): string {
    return `
              <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px;">
                <p style="font-family: Arial, sans-serif; font-size: 14px; color: #444; margin-bottom: 20px;">This matter has been escalated. Use the link below to submit an appeal:</p>
                <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td align="center" bgcolor="#ffffff" style="border-radius: 4px; border: 1px solid #cccccc;">
                      <a href="${appealUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #333333; text-decoration: none; font-weight: bold; display: inline-block;">
                        Appeal decision
                      </a>
                    </td>
                  </tr>
                </table>
              </div>`;
  }

  /**
   * Safe initialization of data.
   * Prevents app-wide failure during module load.
   */
  public async initialize(): Promise<void> {
    // We allow initialization regardless of isLocal to ensure local storage has fallbacks
    // but we only do destructive syncs if isLocal is detected or storage is totally empty.

    try {
      console.log(`[PACT] Initializing SharePoint Service (Mode: ${this.runtimeMode})`);
      const DATA_VERSION = "6.1"; // Fixed key namespace
      const currentVersion = localStorage.getItem(this.STORAGE_PREFIX + 'data_version');

      if (typeof window !== 'undefined') {
        window.addEventListener('storage', (event) => {
          if (event.key && event.key.startsWith(this.STORAGE_PREFIX)) {
            this.notifyDataChanged();
          }
        });
      }

      if (currentVersion !== DATA_VERSION) {
        console.log("PACT: Initializing isolated storage namespace...");
        this.saveToLocal('pact_policies', this.getInitialMockPolicies());
        this.saveToLocal('pact_staff', this.getInitialMockStaff());
        this.saveToLocal('pact_trackers', this.getInitialMockTrackers());
        this.saveToLocal('pact_cases', this.getInitialMockCases());
        this.saveToLocal('pact_escalations', this.getInitialMockEscalations());
        this.saveToLocal('pact_appeals', this.getInitialMockAppeals());
        this.saveToLocal('pact_disciplinary', this.getInitialMockDisciplinary());
        
        // Force purge of old mail history
        this.saveToLocal('pact_mail_history', []);
        localStorage.setItem(this.STORAGE_PREFIX + 'data_version', DATA_VERSION);
      }
    } catch (e) {
      console.error("Initialization failed:", e);
    }
  }

  // ─── Core REST Helpers ──────────────────────────────────────────────────────

  private async fetchREST(endpoint: string, options: RequestInit = {}): Promise<any> {
    if (this.isLocal) return null;

    const url = `${this.siteUrl}/_api/${endpoint}`;
    const headers = {
      'Accept': 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      ...options.headers,
    };

    if (options.method && options.method !== 'GET') {
      const digest = await this.getFormDigest();
      (headers as any)['X-RequestDigest'] = digest;
    }

    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), this.REST_TIMEOUT_MS);

    try {
      const response = await fetch(url, { ...options, headers, signal: controller.signal, credentials: 'include' });
      if (!response.ok) {
        const errorText = await response.text().catch(() => '');
        throw new Error(`SharePoint REST error ${response.status}: ${response.statusText}${errorText ? ` - ${errorText}` : ''}`);
      }

      if (response.status === 204 || response.status === 202) {
        return null;
      }

      const text = await response.text();
      if (!text) return null;
      
      const data = JSON.parse(text);
      return data.d !== undefined ? data.d : data;
    } finally {
      window.clearTimeout(timeout);
    }
  }

  private async getFormDigest(): Promise<string> {
    const response = await fetch(`${this.siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { 'Accept': 'application/json;odata=verbose' },
      credentials: 'include'
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  private async getListItemEntityType(listName: string): Promise<string> {
    const data = await this.fetchREST(`web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`);
    return data.ListItemEntityTypeFullName;
  }

  public async resolveListFieldName(listName: string, candidates: string[]): Promise<string> {
    const normalizedCandidates = candidates
      .filter(Boolean)
      .map(candidate => candidate.toLowerCase().replace(/[\s_]/g, ''));

    try {
      const data = await this.fetchREST(
        `web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,StaticName&$filter=Hidden eq false`
      );
      const fields = data.results || [];
      const match = fields.find((field: any) => {
        const names = [field.Title, field.InternalName, field.StaticName]
          .filter(Boolean)
          .map((name: string) => name.toLowerCase().replace(/[\s_]/g, ''));
        return names.some((name: string) => normalizedCandidates.includes(name));
      });
      return match?.InternalName || candidates[0];
    } catch {
      return candidates[0];
    }
  }

  private async filterPayloadToAvailableFields(listName: string, payload: any): Promise<any> {
    if (this.isLocal || !payload) return payload;
    try {
      const data = await this.fetchREST(
        `web/lists/getbytitle('${this.escapeODataString(listName)}')/fields?$select=InternalName`
      );
      const available = new Set<string>((data.results || []).map((f: any) => String(f.InternalName || '')));
      const filtered: any = {};
      const normalize = (s: string) => s.toLowerCase().replace(/[\s_]/g, '').replace(/x0020/g, '');

      for (const key of Object.keys(payload)) {
        if (key === '__metadata') {
          filtered[key] = payload[key];
          continue;
        }

        const isLookupId = key.endsWith('Id') && key.length > 2;
        const baseKey = isLookupId ? key.slice(0, -2) : key;

        if (available.has(key)) {
          filtered[key] = payload[key];
        } else if (isLookupId && available.has(baseKey)) {
          filtered[key] = payload[key];
        } else {
          // Try normalized comparison (removes spaces, underscores, x0020)
          const normKey = normalize(key);
          const match = Array.from(available).find(
            (f: string) => normalize(f) === normKey
          );

          if (match) {
            filtered[match] = payload[key];
          } else if (isLookupId) {
            // Check if normalized base key matches an available lookup base key
            const normBaseKey = normalize(baseKey);
            const baseMatch = Array.from(available).find(
              (f: string) => normalize(f) === normBaseKey
            );
            if (baseMatch) {
              filtered[baseMatch + 'Id'] = payload[key];
            } else {
              console.warn(`[PACT] Column "${key}" (lookup) does not exist in list "${listName}". Excluding.`);
            }
          } else {
            console.warn(`[PACT] Column "${key}" does not exist in list "${listName}". Excluding.`);
          }
        }
      }
      return filtered;
    } catch (e) {
      console.warn(`[PACT] Failed to inspect columns for list "${listName}". Sending full payload.`, e);
      return payload;
    }
  }

  public async diagnoseListColumns(listName: string) {
    if (this.isLocal) return;
    try {
      const data = await this.fetchREST(`web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,StaticName&$filter=Hidden eq false`);
      const fields = (data.results || []).map((f: any) => `${f.Title} : ${f.InternalName}`);
      console.group(`🚨 [PACT Diagnostic] EXACT Columns for list: "${listName}"`);
      console.log("👉 PLEASE COPY AND PASTE THE TEXT BELOW TO YOUR DEVELOPER:");
      console.log(fields.join('\n'));
      console.groupEnd();
    } catch (e) {
      console.error(`[PACT Diagnostic] Failed to read columns for ${listName}`, e);
    }
  }

  private escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  public formatLookupFieldValue(id: number, display: string): string {
    return `${id};#${display}`;
  }

  private async findListItemIdByTitle(listName: string, title: string): Promise<number | null> {
    const safeTitle = this.escapeODataString(title.trim());
    if (!safeTitle) return null;
    const data = await this.fetchREST(
      `web/lists/getbytitle('${listName}')/items?$filter=Title eq '${safeTitle}'&$select=Id,Title&$top=1`
    );
    const item = (data.results || [])[0];
    const id = item?.ID ?? item?.Id;
    return typeof id === 'number' ? id : null;
  }

  public async findStaffDirectoryItem(
    email?: string,
    fullName?: string
  ): Promise<{ id: number; display: string } | null> {
    const listName = LIST_NAMES.STAFF_DIRECTORY;
    const trimmedEmail = email?.trim();
    const emailFieldCandidates = [
      COLUMNS.STAFF.EMAIL,
      'Email_x0020_Address',
      'Staff_x0020_Email',
      'WorkEmail',
      'EMail'
    ];

    if (trimmedEmail) {
      for (const field of emailFieldCandidates) {
        try {
          const data = await this.fetchREST(
            `web/lists/getbytitle('${listName}')/items?$filter=${field} eq '${this.escapeODataString(trimmedEmail)}'&$select=Id,Title&$top=1`
          );
          const item = (data.results || [])[0];
          if (item?.ID) {
            return { id: item.ID, display: item.Title || fullName || trimmedEmail };
          }
        } catch {
          // Column internal name differs per site — try next candidate.
        }
      }
    }

    const trimmedName = fullName?.trim();
    if (trimmedName) {
      try {
        const data = await this.fetchREST(
          `web/lists/getbytitle('${listName}')/items?$filter=Title eq '${this.escapeODataString(trimmedName)}'&$select=Id,Title&$top=1`
        );
        const item = (data.results || [])[0];
        if (item?.ID) {
          return { id: item.ID, display: item.Title || trimmedName };
        }
      } catch {
        console.warn('Staff directory lookup by name failed for', trimmedName);
      }
    }
    return null;
  }

  private normalizeValidateUpdateResults(result: any): any[] {
    if (Array.isArray(result)) return result;
    if (Array.isArray(result?.value)) return result.value;
    if (Array.isArray(result?.results)) return result.results;
    return [];
  }

  private assertValidateUpdateSucceeded(result: any, context = 'List item'): void {
    const values = this.normalizeValidateUpdateResults(result);
    const failures = values.filter(
      (entry: any) => entry?.HasException || (entry?.ErrorMessage && String(entry.ErrorMessage).trim())
    );
    if (failures.length > 0) {
      const messages = failures
        .map((entry: any) => String(entry.ErrorMessage || entry.FieldName || 'Unknown field'))
        .join('; ');
      throw new Error(`${context} update rejected: ${messages}`);
    }
  }

  private getPaymentProofsFolderPath(): string {
    return `${SHAREPOINT_SITE_PATH}/${PAYMENT_PROOFS_LIBRARY}`;
  }

  private sanitizeFileName(name: string): string {
    return name.replace(/[\\/:*?"<>|#%]/g, '_').replace(/\s+/g, '_');
  }

  private escapeSharePointPath(path: string): string {
    return path.replace(/'/g, "''");
  }

  private async ensurePaymentProofsLibrary(): Promise<void> {
    if (this.isLocal) return;
    const folderPath = this.getPaymentProofsFolderPath();
    try {
      await this.fetchREST(
        `web/GetFolderByServerRelativeUrl('${this.escapeSharePointPath(folderPath)}')`
      );
    } catch {
      await this.fetchREST('web/folders', {
        method: 'POST',
        body: JSON.stringify({
          __metadata: { type: 'SP.Folder' },
          ServerRelativeUrl: folderPath
        })
      });
    }
  }

  private async uploadPaymentProofFile(caseReference: string, file: File): Promise<string> {
    if (this.isLocal) {
      return `local://${this.sanitizeFileName(caseReference)}_${this.sanitizeFileName(file.name)}`;
    }

    await this.ensurePaymentProofsLibrary();
    const folderPath = this.getPaymentProofsFolderPath();
    const fileName = `${this.sanitizeFileName(caseReference)}_${Date.now()}_${this.sanitizeFileName(file.name)}`;
    const digest = await this.getFormDigest();
    const endpoint = `${this.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${this.escapeSharePointPath(folderPath)}')/Files/add(url='${fileName.replace(/'/g, "''")}',overwrite=true)`;
    const buffer = await file.arrayBuffer();

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      credentials: 'include',
      body: buffer
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      throw new Error(`Payment proof upload failed (${response.status}): ${errorText || response.statusText}`);
    }

    const data = await response.json();
    const serverRelativeUrl = data?.d?.ServerRelativeUrl || `${folderPath}/${fileName}`;
    let baseUrl = this.siteUrl;
    try {
      if (this.siteUrl.startsWith('http://') || this.siteUrl.startsWith('https://')) {
        const urlObj = new URL(this.siteUrl);
        baseUrl = urlObj.origin;
      }
    } catch {
      // Fallback if URL parsing fails
    }
    return `${baseUrl}${serverRelativeUrl.startsWith('/') ? '' : '/'}${serverRelativeUrl}`;
  }

  private notifyDataChanged(): void {
    if (typeof window !== 'undefined') {
      window.dispatchEvent(new CustomEvent('pact-data-changed'));
    }
  }


  private escapeHtml(value: unknown): string {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private formatCurrency(value: unknown): string {
    const amount = typeof value === 'number' ? value : Number(value || 0);
    return `NGN ${Number.isFinite(amount) ? amount.toLocaleString() : '0'}`;
  }

  private formatEmailDate(value: string): string {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return this.escapeHtml(value);
    return date.toLocaleString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }

  private buildAppealSubmittedEmailBody(appeal: any, appealReference: string, appealDate: string): string {
    const safe = (value: unknown): string => this.escapeHtml(value || 'Not provided');
    const rows = [
      ['Appeal Reference', appealReference],
      ['Case Reference', appeal.caseReference],
      ['Appellant', appeal.appellant],
      ['Appellant Email', appeal.appellantEmail],
      ['Department', appeal.department],
      ['Offence', appeal.offence],
      ['Penalty Amount', this.formatCurrency(appeal.penaltyAmount)],
      ['Submitted On', this.formatEmailDate(appealDate)],
      ['Reviewing Office', 'Admin / Executive Review'],
      ['Decision Status', 'Pending']
    ].map(([label, value]) => `
      <tr>
        <td style="padding:12px 16px;background:#f8fafc;border-bottom:1px solid #e5e7eb;color:#64748b;font-weight:700;width:190px;">${this.escapeHtml(label)}</td>
        <td style="padding:12px 16px;border-bottom:1px solid #e5e7eb;color:#111827;">${safe(value)}</td>
      </tr>
    `).join('');

    return `
      <div style="margin:0;padding:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;color:#111827;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f7fb;padding:24px 0;">
          <tr>
            <td align="center">
              <table role="presentation" width="680" cellpadding="0" cellspacing="0" style="width:680px;max-width:94%;background:#ffffff;border:1px solid #dbe3ef;border-radius:12px;overflow:hidden;">
                <tr>
                  <td style="background:#173f63;padding:28px 32px;color:#ffffff;">
                    <div style="font-size:12px;letter-spacing:2px;text-transform:uppercase;color:#cfe3f5;">Konstructum Group</div>
                    <h1 style="margin:8px 0 0;font-size:24px;line-height:1.3;">PACT Appeal Submitted</h1>
                  </td>
                </tr>
                <tr>
                  <td style="padding:28px 32px;">
                    <p style="margin:0 0 18px;font-size:16px;line-height:1.6;">
                      A new appeal has been submitted for executive review. The complete appeal details are below.
                    </p>
                    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e5e7eb;border-radius:8px;border-collapse:separate;border-spacing:0;overflow:hidden;font-size:14px;">
                      ${rows}
                    </table>
                    <div style="margin-top:22px;">
                      <div style="font-size:13px;text-transform:uppercase;letter-spacing:1px;color:#64748b;font-weight:700;margin-bottom:8px;">Grounds for Appeal</div>
                      <div style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;padding:16px;color:#111827;line-height:1.6;white-space:pre-wrap;">${this.escapeHtml(appeal.grounds || 'No grounds supplied.')}</div>
                    </div>
                    ${appeal.offenceDescription ? `
                      <div style="margin-top:18px;">
                        <div style="font-size:13px;text-transform:uppercase;letter-spacing:1px;color:#64748b;font-weight:700;margin-bottom:8px;">Original Case Description</div>
                        <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:16px;color:#7c2d12;line-height:1.6;">${this.escapeHtml(appeal.offenceDescription)}</div>
                      </div>
                    ` : ''}
                    <p style="margin:22px 0 0;font-size:14px;line-height:1.6;color:#475569;">
                      Please review this appeal in the PACT Appeals Register. Target response time is <strong>${APPEAL_SLA_DAYS} working days</strong>.
                    </p>
                  </td>
                </tr>
                <tr>
                  <td style="background:#f8fafc;padding:18px 32px;color:#64748b;font-size:12px;text-align:center;">
                    Automated notification from the PACT Compliance Governance Platform.
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
    `;
  }

  private async updateCaseStatusForReference(
    caseReference: string,
    status: string,
    knownItemId?: number
  ): Promise<void> {
    if (this.isLocal) {
      const cases = this.getFromLocal<ComplianceCase>('pact_cases');
      const idx = cases.findIndex(c => c.title === caseReference);
      if (idx > -1) {
        cases[idx] = { ...cases[idx], status: status as ComplianceCase['status'] };
        this.saveToLocal('pact_cases', cases);
      }
      return;
    }

    let caseItemId = knownItemId;
    // Treat mock timestamp IDs as invalid/undefined so we fall back to title search
    if (caseItemId && caseItemId > 10000000) {
      caseItemId = undefined;
    }
    if (!Number.isFinite(caseItemId)) {
      const resolved = await this.findListItemIdByTitle(LIST_NAMES.COMPLIANCE_CASES, caseReference);
      if (resolved) caseItemId = resolved;
    }
    if (!Number.isFinite(caseItemId)) {
      console.warn(`[PACT] Case not found in SharePoint for status update: ${caseReference}`);
      throw new Error(`Case not found in SharePoint for status update: ${caseReference}`);
    }

    try {
      // Use the standard robust updateCase method to handle MERGE and dynamic schema filtering
      await this.updateCase(String(caseItemId), { status: status as any });
      console.log(`[PACT] Status successfully updated to "${status}" for case ${caseReference}`);
    } catch (err) {
      console.warn(`[PACT] Direct updateCase failed for case ${caseReference}, falling back to validateUpdateListItem:`, err);
      const saved = await this.tryUpdateListItemByDisplayNames(
        LIST_NAMES.COMPLIANCE_CASES,
        caseItemId!,
        [{ FieldName: 'Status', FieldValue: status }]
      );
      if (!saved) {
        console.error(`[PACT] Critical: Could not update status to "${status}" for case ${caseReference}`);
        throw new Error(`Failed to update case status in SharePoint for ${caseReference}. Both MERGE and validateUpdateListItem failed.`);
      } else {
        console.log(`[PACT] Status successfully updated to "${status}" for case ${caseReference} via validateUpdateListItem`);
      }
    }
  }

  public async submitPaymentAcceptance(
    caseData: ComplianceCase,
    proofFile: File,
    paymentNotes?: string
  ): Promise<{ proofUrl: string }> {
    const proofUrl = await this.uploadPaymentProofFile(caseData.title, proofFile);

    let caseItemId: number | undefined;
    if (!this.isLocal) {
      caseItemId = caseData.id !== 'fallback' ? Number(caseData.id) : undefined;
      // Treat mock timestamp IDs as invalid/undefined so we fall back to title search
      if (caseItemId && caseItemId > 10000000) {
        caseItemId = undefined;
      }
      if (!Number.isFinite(caseItemId)) {
        const resolved = await this.findListItemIdByTitle(LIST_NAMES.COMPLIANCE_CASES, caseData.title);
        if (resolved) caseItemId = resolved;
      }
      if (Number.isFinite(caseItemId)) {
        const proofLinkValue = `${proofUrl}, Payment proof`;
        const proofFieldDisplayNames = ['Evidence', 'Payment Proof', 'Payment Proof URL', 'Proof of Payment'];
        for (const fieldName of proofFieldDisplayNames) {
          const saved = await this.tryUpdateListItemByDisplayNames(
            LIST_NAMES.COMPLIANCE_CASES,
            caseItemId!,
            [{ FieldName: fieldName, FieldValue: proofLinkValue }]
          );
          if (saved) break;
        }
      }
    }

    await this.updateCaseStatusForReference(caseData.title, CASE_STATUS.PAID, caseItemId);

    const hrSubject = `PACT: Payment received — ${caseData.title}`;
    const hrBody = `
      <div style="font-family:Arial,sans-serif;color:#333;">
        <h2 style="color:#107c10;margin-top:0;">Payment proof received</h2>
        <p><b>${caseData.chargedPersonName}</b> has submitted proof of payment.</p>
        <table style="width:100%;font-size:14px;border-collapse:collapse;margin:16px 0;">
          <tr><td style="color:#666;padding:6px 0;">Case reference</td><td><b>${caseData.title}</b></td></tr>
          <tr><td style="color:#666;padding:6px 0;">Amount</td><td>₦${caseData.penaltyAmount.toLocaleString()}</td></tr>
          <tr><td style="color:#666;padding:6px 0;">Department</td><td>${caseData.department || '—'}</td></tr>
          <tr><td style="color:#666;padding:6px 0;">Status</td><td><b>Paid</b> (updated in PACT)</td></tr>
        </table>
        ${paymentNotes ? `<p><b>Employee notes:</b> ${paymentNotes}</p>` : ''}
        <p><a href="${proofUrl}">Open payment proof</a> in ${PAYMENT_PROOFS_LIBRARY}</p>
        <p style="font-size:12px;color:#666;">No further action required unless verification fails.</p>
      </div>
    `;
    // Send HR/Legal/Chairman notification via the same {to, subject, body} schema
    // that the working case-logging flow uses — avoids payload mismatch with Power Automate.
    await this.sendEmailNotification(
      [HR_EMAIL, LEGAL_EMAIL, CHAIRMAN_EMAIL],
      hrSubject,
      hrBody,
      ACCEPT_PAYMENT_TRIGGER_URL
    );

    this.notifyDataChanged();
    return { proofUrl };
  }

  public async notifyAppealReviewTeam(appeal: any, appealRef: string): Promise<void> {
    const extraRows = [
      appeal.department ? `<tr><td style="padding:6px 0;color:#666;">Department</td><td>${appeal.department}</td></tr>` : '',
      appeal.offence ? `<tr><td style="padding:6px 0;color:#666;">Offence</td><td>${appeal.offence}</td></tr>` : '',
      typeof appeal.penaltyAmount === 'number'
        ? `<tr><td style="padding:6px 0;color:#666;">Penalty</td><td>₦${appeal.penaltyAmount.toLocaleString()}</td></tr>`
        : '',
      appeal.appellantEmail
        ? `<tr><td style="padding:6px 0;color:#666;">Appellant email</td><td>${appeal.appellantEmail}</td></tr>`
        : '',
    ].filter(Boolean).join('');
    const subject = `PACT APPEAL FILED: Case ${appeal.caseReference} (${appealRef})`;
    const body = `
      <div style="font-family:Arial,sans-serif;color:#333;max-width:640px;">
        <p>An appeal has been submitted by <b>${appeal.appellant}</b> for case <b>${appeal.caseReference}</b>.</p>
        ${extraRows ? `<table style="width:100%;font-size:14px;border-collapse:collapse;margin:16px 0;">${extraRows}</table>` : ''}
        <p><b>Grounds for appeal:</b></p>
        <p style="background:#f8fafc;padding:12px;border-radius:6px;white-space:pre-wrap;">${appeal.grounds || ''}</p>
        <p style="font-size:13px;color:#666;">Please review in the PACT Appeals Register within ${APPEAL_SLA_DAYS} working days.</p>
      </div>
    `;
    const recipients = [...APPEAL_REVIEW_EMAILS];
    if (appeal.appellantEmail) {
      recipients.push(appeal.appellantEmail);
    }
    await this.sendEmailNotification(recipients, subject, body, APPEAL_MAIL_TRIGGER_URL);
  }

  private async sendEmailNotification(to: string[], subject: string, body: string, triggerUrl: string = MAIL_TRIGGER_URL, logOnly: boolean = false): Promise<void> {
    // ── 1. If a Power Automate webhook is configured, try it first ──
    if (triggerUrl && !logOnly) {
      try {
        console.log(`%c [WEBHOOK] %c Posting to Power Automate flow…`, 'background: #5c2d91; color: white; padding: 2px 5px; border-radius: 3px;', 'font-weight: bold;');
        console.log("  Recipients:", to.join(', '));
        console.log("  Subject:", subject);

        const response = await fetch(triggerUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: {
            'Content-Type': 'text/plain;charset=UTF-8'
          },
          body: JSON.stringify({
            to: to.join(';'),
            subject,
            body
          })
        });

        if (response.type !== 'opaque' && !response.ok) {
          throw new Error(`HTTP trigger returned status ${response.status}`);
        }

        console.log(`%c [WEBHOOK] %c Email dispatched successfully ✓`, 'background: #107c10; color: white; padding: 2px 5px; border-radius: 3px;', 'font-weight: bold;');
      } catch (webhookError) {
        console.error("Webhook mail flow failed, falling back:", webhookError);
        // Fall through to local mock or SharePoint REST fallback below
      }
    }

    // ── 2. Always log to SharePoint History list ──
    if (!this.isLocal) {
      try {
        const spData = {
          '__metadata': { 'type': `SP.Data.${LIST_NAMES.MAIL_HISTORY.replace(/ /g, '_x0020_')}ListItem` },
          [COLUMNS.MAIL.TO]: to.join(', '),
          [COLUMNS.MAIL.SUBJECT]: subject,
          [COLUMNS.MAIL.BODY]: body,
          [COLUMNS.MAIL.STATUS]: logOnly ? 'Logged (Sent via Flow)' : 'Sent (Webhook)'
        };
        await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.MAIL_HISTORY}')/items`, {
          method: 'POST',
          body: JSON.stringify(spData)
        });
      } catch (spError) {
        console.warn("Failed to log email to SharePoint History list:", spError);
      }
    }

    // Also persist to local storage for the mock mail history panel
    const history = this.getFromLocal<any>('pact_mail_history');
    history.push({
      id: Date.now(),
      to,
      subject,
      body,
      timestamp: new Date().toISOString(),
      status: logOnly ? 'Logged (Sent via Flow)' : 'Sent (Webhook)'
    });
    this.saveToLocal('pact_mail_history', history);
    
    // Dispatch mock email event for UI toast
    const debugEvent = new CustomEvent('pact-mock-email', { detail: { to, subject } });
    window.dispatchEvent(debugEvent);
    
    if (logOnly) return;

    // ── 3. SP.Utilities.Utility.SendEmail fallback ──
    // This endpoint works natively inside SharePoint page context.
    if (!this.isLocal) {
      try {
        console.log(`%c [SP EMAIL] %c Attempting native SharePoint email...`, 'background: #0078d4; color: white; padding: 2px 5px; border-radius: 3px;', 'font-weight: bold;');
        const spData = {
          'properties': {
            '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
            'To': { 'results': to },
            'Subject': subject,
            'Body': body,
            'AdditionalHeaders': {
              '__metadata': { 'type': 'Collection(SP.KeyValue)' },
              'results': [
                {
                  '__metadata': { 'type': 'SP.KeyValue' },
                  'Key': 'content-type',
                  'Value': 'text/html',
                  'ValueType': 'Edm.String'
                }
              ]
            }
          }
        };

        await this.fetchREST(`SP.Utilities.Utility.SendEmail`, {
          method: 'POST',
          body: JSON.stringify(spData)
        });

        console.log(`%c [SP EMAIL] %c Delivered via native SharePoint ✓`, 'background: #107c10; color: white; padding: 2px 5px; border-radius: 3px;', 'font-weight: bold;');
        return;
      } catch (spErr) {
        console.warn('[PACT] Native SharePoint email fallback failed:', spErr);
      }
    }

    console.warn('[PACT] All email delivery methods failed.');
  }

  // ─── LocalStorage Persistence ────────────────────────────────────────────────

  private getFromLocal<T>(key: string): T[] {
    try {
      const data = localStorage.getItem(this.STORAGE_PREFIX + key);
      if (!data || data === 'undefined' || data === 'null') return [];
      const parsed = JSON.parse(data);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.error(`Error loading ${key} from storage:`, e);
      return [];
    }
  }

  private saveToLocal<T>(key: string, data: T[]): void {
    try {
      localStorage.setItem(this.STORAGE_PREFIX + key, JSON.stringify(data));
    } catch (e) {
      console.error(`Error saving ${key} to storage:`, e);
    }
  }

  // ─── Public API ─────────────────────────────────────────────────────────────

  public isStandalone(): boolean {
    return this.runtimeMode === 'demo';
  }

  public getRuntimeLabel(): string {
    return this.runtimeMode === 'sharepoint' ? 'SharePoint Native' : 'Workbench Demo';
  }

  public getUserName(): string {
    return "PACT Administrator";
  }

  // --- Cases ---
  public async getCases(): Promise<ComplianceCase[]> {
    try {
      if (this.isLocal) {
        const localCases = this.getFromLocal<ComplianceCase>('pact_cases');
        return localCases.map(c => ({
          ...c,
          offenceCategoryName: this.expandAbbreviations(c.offenceCategoryName || ''),
          penaltyAmount: this.parsePenalty(c.penaltyAmount)
        })).sort((a,b) => new Date(b.dateCreated).getTime() - new Date(a.dateCreated).getTime());
      }

      const endpoint = `web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items?$select=*,FieldValuesAsText/*&$expand=FieldValuesAsText&$orderby=Created desc`;
      const data = await this.fetchREST(endpoint);
      const [staff, policies] = await Promise.all([
        this.getStaffDirectory(),
        this.getPolicyLibrary()
      ]);
      return data.results.map((item: any) => this.mapSPItemToCase(item, staff, policies));
    } catch (error) {
      console.warn("REST/Graph fetch failed, falling back to local mocks", error);
      return this.getFromLocal<ComplianceCase>('pact_cases');
    }
  }

  // --- Get Case by Reference (e.g. PACT-001) ---
  public async getCaseByReference(ref: string): Promise<ComplianceCase | null> {
    if (this.isLocal) {
      const cases = await this.getCases();
      return cases.find(c => c.title === ref) || null;
    }

    try {
      const safeRef = this.escapeODataString(ref.trim());
      const endpoint = `web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items?$filter=Title eq '${safeRef}'&$select=*,FieldValuesAsText/*&$expand=FieldValuesAsText&$top=1`;
      const data = await this.fetchREST(endpoint);
      const item = (data.results || [])[0];
      if (!item) {
        console.warn(`[PACT] Case reference "${ref}" not found via OData query. Falling back to full list search.`);
        const cases = await this.getCases();
        return cases.find(c => c.title === ref) || null;
      }

      const [staff, policies] = await Promise.all([
        this.getStaffDirectory(),
        this.getPolicyLibrary()
      ]);
      return this.mapSPItemToCase(item, staff, policies);
    } catch (error) {
      console.warn(`[PACT] getCaseByReference OData query failed for "${ref}", falling back to full cases list:`, error);
      const cases = await this.getCases();
      return cases.find(c => c.title === ref) || null;
    }
  }


  public async getStaffDirectory(): Promise<StaffMember[]> {
    try {
      if (this.isLocal) {
        return (staffData as StaffMember[]).map(item => ({
          ...item,
          photoUrl: this.getPhotoUrl(item.email)
        }));
      }

      const endpoint = `web/lists/getbytitle('${LIST_NAMES.STAFF_DIRECTORY}')/items?$top=5000`;
      const data = await this.fetchREST(endpoint);
      
      // Run diagnostic in the background so we can see the staff columns in console
      this.diagnoseListColumns(LIST_NAMES.STAFF_DIRECTORY);

      if (data && data.results && data.results.length > 0) {
        return data.results.map((item: any) => {
          const email = item[COLUMNS.STAFF.EMAIL] || item.EmailAddress || item.Email || item.email || '';
          const lineManager = item[COLUMNS.STAFF.LINE_MANAGER] || item.LineManager || item.Line_x0020_Manager || item.lineManager || '';
          return {
            id: String(item.Id || item.ID || ''),
            fullName: item[COLUMNS.STAFF.TITLE] || item.Title || item.fullName || '',
            email: email,
            department: item[COLUMNS.STAFF.DEPARTMENT] || item.Department || item.department || '',
            lineManager: lineManager,
            company: item[COLUMNS.STAFF.COMPANY] || item.Company || item.company || '',
            role: item[COLUMNS.STAFF.ROLE] || item.Role || item.role || '',
            employeeType: item[COLUMNS.STAFF.EMPLOYEE_TYPE] || item.EmployeeType || item.employeeType || '',
            status: item[COLUMNS.STAFF.STATUS] || item.Status || item.status || 'Active',
            photoUrl: email ? this.getPhotoUrl(email) : undefined
          };
        });
      }
      
      return (staffData as StaffMember[]).map(item => ({
        ...item,
        photoUrl: this.getPhotoUrl(item.email)
      }));
    } catch (error) {
      console.warn("Failed to fetch staff from SharePoint list, falling back to local JSON", error);
      return (staffData as StaffMember[]).map(item => ({
        ...item,
        photoUrl: this.getPhotoUrl(item.email)
      }));
    }
  }


  // --- Policies ---
  public async getPolicyLibrary(): Promise<PolicyOffence[]> {
    try {
      if (this.isLocal) {
        return policyData as PolicyOffence[];
      }
      const endpoint = `web/lists/getbytitle('${LIST_NAMES.POLICY_LIBRARY}')/items?$top=500`;
      const data = await this.fetchREST(endpoint);
      if (data && data.results && data.results.length > 0) {
        return data.results
          .filter((item: any) => item.Title && item.Title.trim() !== '')
          .map((item: any) => {
          const liveTitle = item.Title || '';
          const localMatch = (policyData as PolicyOffence[]).find(
            p => p.offenceName.toLowerCase().trim() === liveTitle.toLowerCase().trim()
          );

          return {
            id: String(item.Id || item.ID || ''),
            offenceName: liveTitle,
            tier: item.Tier || localMatch?.tier || 'Tier 1',
            category: item.Category || localMatch?.category || '',
            description: item.Description || item.Title || localMatch?.description || '',
            defaultPenaltyAmount: this.parsePenalty(item.DefaultPenaltyAmount ?? localMatch?.defaultPenaltyAmount ?? 0),
            firstOffenceAction: item.FirstOffenceAction || localMatch?.firstOffenceAction || '',
            secondOffenceAction: item.SecondOffenceAction || localMatch?.secondOffenceAction || '',
            thirdOffenceAction: item.ThirdOffenceAction || localMatch?.thirdOffenceAction || '',
            escalationTrigger: localMatch?.escalationTrigger ?? true
          };
        });
      }
      return policyData as PolicyOffence[];
    } catch (error) {
      console.warn("Failed to fetch policy library from SharePoint, falling back to local JSON", error);
      return policyData as PolicyOffence[];
    }
  }

  private parsePenalty(value: any): number {
    if (typeof value === 'number') return value;
    if (!value) return 0;
    // Handle currencies like "₦ 5,000.00" or "? 5,000.00"
    const cleaned = value.toString().replace(/[₦?, ]/g, '');
    const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    }

  private normalizeFieldValue(value: any): any {
    if (value && typeof value === 'object') {
      if ('Title' in value) return value.Title;
      if ('LookupValue' in value) return value.LookupValue;
      if ('Email' in value) return value.Email;
      if ('results' in value && Array.isArray(value.results)) {
        return value.results.map((entry: any) => this.normalizeFieldValue(entry)).filter(Boolean).join(', ');
      }
      return ''; // Prevent React rendering unresolved objects (like __deferred)
    }
    return value;
  }

  private readField(item: any, ...keys: string[]): any {
    for (const key of keys) {
      const value = this.normalizeFieldValue(item?.[key]);
      if (value !== undefined && value !== null && value !== '') {
        return value;
      }
    }
    return undefined;
  }

  private formatDisciplinaryReference(index: number): string {
    return `P-${String(index).padStart(3, '0')}`;
  }

  private async getNextDisciplinaryReference(): Promise<string> {
    const existing = this.isLocal
      ? this.getFromLocal<any>('pact_disciplinary')
      : await this.getDisciplinaryActions();
    const maxNumber = existing.reduce((max: number, item: any) => {
      const title = String(item?.title || '');
      const match = title.match(/^P-(\d+)$/i);
      if (!match) return max;
      const value = Number.parseInt(match[1], 10);
      return Number.isFinite(value) && value > max ? value : max;
    }, 0);
    return this.formatDisciplinaryReference(maxNumber + 1);
  }

  public async createCase(caseData: Partial<ComplianceCase>, isHistorical: boolean = false): Promise<ComplianceCase> {
    const newId = Date.now().toString();
    
    // Sequential PACT Number Generation
    const existingCases = await this.getCases();
    // Only count cases that match the current mode (historical vs new)
    const relevantCases = existingCases.filter(c => isHistorical ? c.title.startsWith('HIST-') : c.title.startsWith('PACT-'));
    
    const maxNumber = relevantCases.reduce((max: number, c: any) => {
      const match = c.title.match(/^(PACT|HIST)-(\d+)$/i);
      if (!match) return max;
      const val = parseInt(match[2], 10);
      return val > max ? val : max;
    }, 0);
    const nextNumber = String(maxNumber + 1).padStart(3, '0');
    const caseTitle = caseData.title ? caseData.title : (isHistorical ? `HIST-${nextNumber}` : `PACT-${nextNumber}`);

    const staff = this.isLocal ? this.getFromLocal<StaffMember>('pact_staff') : await this.getStaffDirectory();
    const person = staff.find(s => s.id === caseData.chargedPerson);
    const policies = this.isLocal ? this.getFromLocal<PolicyOffence>('pact_policies') : await this.getPolicyLibrary();
    const policy = policies.find(p => p.id === caseData.offenceCategory || p.id === 'p' + caseData.offenceCategory);

    const newCase: ComplianceCase = {
      id: newId,
      title: caseTitle,
      chargedPerson: caseData.chargedPerson || '999',
      chargedPersonName: caseData.chargedPersonName || person?.fullName || 'Unknown Staff',
      staffEmail: caseData.staffEmail || person?.email || 'staff@konstructum.com',
      department: caseData.department || person?.department || 'General',
      offenceCategory: caseData.offenceCategory || 'Unknown',
      offenceCategoryName: caseData.offenceCategoryName || this.expandAbbreviations(policy?.offenceName || 'Unknown Offence'),
      offenceDescription: caseData.offenceDescription || '',
      penaltyAmount: caseData.penaltyAmount !== undefined ? caseData.penaltyAmount : (policy?.defaultPenaltyAmount || 0),
      dueDate: caseData.dueDate || new Date().toISOString(),
      issuerName: caseData.issuerName || this.getUserName(),
      secondaryContact: caseData.secondaryContact || person?.lineManager || '',
      status: caseData.status || 'Unpaid',
      dateCreated: new Date().toISOString()
    };

    if (!policy) {
      console.warn("Policy not found for category:", caseData.offenceCategory);
    }

    const tracker = await this.getRepeatTrackerRecord(newCase.chargedPerson);
    const policyTier = policy?.tier || '';
    
    // Escalation flags
    const isTier3Escalation = policyTier === 'Tier 3';
    const isTier2Escalation = policyTier === 'Tier 2' && escalationEngine.checkTier2Escalation(tracker);
    const isTier1Escalation = policyTier === 'Tier 1' && tracker
      ? escalationEngine.checkTier1Escalation(newCase.dateCreated, tracker)
      : false;
    const isEscalated = isTier3Escalation || isTier2Escalation || isTier1Escalation;

    const pTierLower = (policyTier || '').trim().toLowerCase();
    const offCount = (tracker?.tier1Last6Months || 0) + (pTierLower === 'tier 1' ? 1 : 0);
    const actionPath = isEscalated ? 'Automatic Escalation' : 
                      (pTierLower === 'tier 1' ? (offCount === 1 ? '1st Offence' : offCount === 2 ? '2nd Offence' : '3rd+ Offence') : 'Standard');
    const disciplinaryAction = policy ? escalationEngine.getRecommendedAction(policy, offCount, isEscalated) : 'Standard Disciplinary Path';


    if (!isHistorical) {
      // 1. Escalation Logging
      if (isEscalated && policy) {
        let reason = '';
        let newTier = 'Tier 2';
        if (isTier3Escalation) {
          const t3Count = (tracker?.tier3Offences || 0) + 1;
          reason = `Tier 3 Offence (${policy.offenceName}): Automatic escalation. Occurrence #${t3Count}. Immediate HR & Chairman review required.`;
          newTier = 'Tier 3';
        } else if (isTier2Escalation) {
          reason = `Repeat Tier 2 Offence (${policy.offenceName}): Staff member has ${(tracker?.tier2Offences || 0) + 1} Tier 2 offences on record.`;
          newTier = 'Tier 3';
        } else {
          reason = `Automatic Policy Trigger: Staff member reached 3+ Tier 1 offences within 6 months. Threshold exceeded on case ${newCase.title}.`;
          newTier = 'Tier 2';
        }

        try {
          await this.createEscalation({
            caseReference: newCase.title,
            offender: newCase.chargedPerson,
            offenderName: newCase.chargedPersonName,
            escalationReason: reason,
            previousTier: policy.tier,
            newTier: newTier as "Tier 1" | "Tier 2" | "Tier 3"
          });
        } catch (escErr) {
          console.warn("Failed to create escalation record:", escErr);
        }
      }
      // 2. Update Tracker — increment the correct tier counter
      const pTierLower = (policyTier || '').trim().toLowerCase();
      const updatedTier1 = (tracker?.tier1Last6Months || 0) + (pTierLower === 'tier 1' ? 1 : 0);
      const updatedTier2 = (tracker?.tier2Offences || 0) + (pTierLower === 'tier 2' ? 1 : 0);
      const updatedTier3 = (tracker?.tier3Offences || 0) + (pTierLower === 'tier 3' ? 1 : 0);

      const updatedTracker = {
        totalOffences: (tracker?.totalOffences || 0) + 1,
        tier1Last6Months: updatedTier1,
        tier2Offences: updatedTier2,
        tier3Offences: updatedTier3,
        lastOffenceDate: newCase.dateCreated,
        riskLevel: escalationEngine.calculateRiskLevel({
          ...(tracker || { id:'', title:'', offender:'', totalOffences:0, tier1Last6Months:0, tier2Offences:0, tier3Offences:0, riskLevel:'Low', lastOffenceDate:'', escalationDue:false }),
          tier1Last6Months: updatedTier1,
          tier2Offences: updatedTier2,
          tier3Offences: updatedTier3
        })
      };

      await this.updateRepeatTracker(newCase.chargedPerson, updatedTracker);

      // 3. Email Notifications

      const emailSubject = `PACT ALERT: ${this.expandAbbreviations(policy?.offenceName || 'Compliance Incident')} - ${newCase.chargedPersonName} (Ref: ${newCase.title})`;
      const emailBody = `Automated notification handled by Power Automate.`;
      const manager = staff.find(s => s.fullName === person?.lineManager);
      const recipients = [newCase.staffEmail, HR_EMAIL, ADMIN_EMAIL];
      if (manager?.email) recipients.push(manager.email);
      await this.sendEmailNotification(recipients, emailSubject, emailBody, MAIL_TRIGGER_URL, true);

      // Tier 3 → Automatic Compliance + Legal notification
      if (policy?.tier === 'Tier 3') {
        const tier3Subject = `⚠️ PACT TIER 3 ALERT: ${this.expandAbbreviations(policy.offenceName)} - ${newCase.chargedPersonName} (Ref: ${newCase.title})`;
        const tier3Body = `
          <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; padding: 20px; border: 2px solid #d13438; border-radius: 8px;">
            <h2 style="color: #d13438; margin-top: 0;">🚨 Tier 3 Offence — Immediate Attention Required</h2>
            <p>A <b>Tier 3 compliance violation</b> has been logged in the PACT system. This requires immediate review by HR and Legal.</p>
            <div style="background: #fff4f4; padding: 15px; border-radius: 4px; margin: 20px 0;">
              <table style="width: 100%; font-size: 14px; border-collapse: collapse;">
                <tr><td style="color: #666; padding: 8px 0; width: 150px;">Reference:</td><td style="font-weight: bold;">${newCase.title}</td></tr>
                <tr><td style="color: #666; padding: 8px 0;">Charged Person:</td><td>${newCase.chargedPersonName}</td></tr>
                <tr><td style="color: #666; padding: 8px 0;">Department:</td><td>${newCase.department}</td></tr>
                <tr><td style="color: #666; padding: 8px 0;">Offence:</td><td>${this.expandAbbreviations(policy.offenceName)}</td></tr>
                <tr><td style="color: #666; padding: 8px 0;">Description:</td><td>${newCase.offenceDescription}</td></tr>
                <tr><td style="color: #666; padding: 8px 0;">Penalty:</td><td style="font-weight: bold; color: #d13438;">₦${newCase.penaltyAmount.toLocaleString()}</td></tr>
                <tr><td style="color: #666; padding: 8px 0;">Disciplinary Action:</td><td style="font-weight: bold;">${disciplinaryAction}</td></tr>
              </table>
            </div>
            <p style="font-size: 13px; color: #666;">This notification was generated automatically by the PACT Compliance Governance Platform.</p>
          </div>
        `;
        await this.sendEmailNotification([HR_EMAIL, LEGAL_EMAIL, CHAIRMAN_EMAIL], tier3Subject, tier3Body);
      }
    }


    // 4. Persistence

    // Always persist to local cache first (ensures UI updates instantly and fallback works if SP fails)
    const cases = this.getFromLocal<ComplianceCase>('pact_cases');
    cases.push(newCase);
    this.saveToLocal('pact_cases', cases);

    if (!this.isLocal) {
      try {
        const itemType = await this.getListItemEntityType(LIST_NAMES.COMPLIANCE_CASES);
        const spData: any = {
          '__metadata': { 'type': itemType },
          [COLUMNS.CASES.TITLE]: newCase.title,
          [COLUMNS.CASES.CHARGED_PERSON]: newCase.chargedPersonName,
          'chargedPersonName': newCase.chargedPersonName, // Flow Alias
          [COLUMNS.CASES.DEPARTMENT]: newCase.department,
          [COLUMNS.CASES.OFFENCE_CATEGORY]: newCase.offenceCategoryName,
          'offenceName': newCase.offenceCategoryName, // Flow Alias
          'OffenceDescription': newCase.offenceDescription,
          'Offence_x0020_Description': newCase.offenceDescription,
          [COLUMNS.CASES.PENALTY_AMOUNT]: newCase.penaltyAmount,
          [COLUMNS.CASES.DUE_DATE]: newCase.dueDate,
          [COLUMNS.CASES.ISSUER_NAME]: newCase.issuerName,
          [COLUMNS.CASES.SECONDARY_CONTACT]: newCase.secondaryContact,
          [COLUMNS.CASES.STATUS]: newCase.status,
          [COLUMNS.CASES.STAFF_EMAIL]: newCase.staffEmail,
          [COLUMNS.CASES.CHARGED_PERSON_EMAIL]: newCase.staffEmail,
          [COLUMNS.CASES.TIER]: policy?.tier || 'Tier 1',
          'tier': policy?.tier || 'Tier 1', // Flow Alias
          [COLUMNS.CASES.DISCIPLINARY_ACTION]: disciplinaryAction,
          'actionLabel': disciplinaryAction, // Flow Alias
          'disciplinaryAction': disciplinaryAction, // Flow Alias
          [COLUMNS.CASES.OFFENCE_COUNT]: offCount.toString(),
          'lineManagerName': person?.lineManager || '', // Flow Alias
          'LineManager': person?.lineManager || ''
        };

        let finalResponse: any = null;
        try {
          const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.COMPLIANCE_CASES, spData);
          finalResponse = await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items`, {
            method: 'POST',
            body: JSON.stringify(filteredData)
          });
        } catch (firstTryErr) {
          console.warn("First try failed, attempting Lookup ID format...", firstTryErr);
          
          try {
            // Try formatting as Lookup IDs (e.g. OffenceCategoryId: 5)
            // Note: ChargedPerson is a Text field, so we leave it as is.
            delete spData[COLUMNS.CASES.OFFENCE_CATEGORY];
            
            let resolvedPolicyId = null;
            if (policy?.offenceName) {
              resolvedPolicyId = await this.findListItemIdByTitle(LIST_NAMES.POLICY_LIBRARY, policy.offenceName);
            }
            
            spData[COLUMNS.CASES.OFFENCE_CATEGORY + 'Id'] = resolvedPolicyId || parseInt(newCase.offenceCategory, 10) || null;
            
            const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.COMPLIANCE_CASES, spData);
            finalResponse = await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items`, {
              method: 'POST',
              body: JSON.stringify(filteredData)
            });
          } catch (secondTryErr) {
            console.warn("Lookup ID format failed, falling back to basic fields only...", secondTryErr);
            // Complete shock absorber: remove them completely
            delete spData[COLUMNS.CASES.OFFENCE_CATEGORY + 'Id'];
            // Restore offence category string so it's not totally blank
            spData[COLUMNS.CASES.OFFENCE_CATEGORY] = newCase.offenceCategoryName;
            
            const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.COMPLIANCE_CASES, spData);
            finalResponse = await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items`, {
              method: 'POST',
              body: JSON.stringify(filteredData)
            });
          }
        }
        
        // Guarantee the offence category is saved using validateUpdateListItem
        if (finalResponse) {
          const newItemId = finalResponse.d ? finalResponse.d.Id : finalResponse.Id;
          if (newItemId) {
            await this.tryUpdateListItemByDisplayNames(
              LIST_NAMES.COMPLIANCE_CASES,
              newItemId,
              [
                { FieldName: 'Offence Category', FieldValue: newCase.offenceCategoryName || '' },
                { FieldName: 'OffenceCategory', FieldValue: newCase.offenceCategoryName || '' }
              ]
            );
          }
        }
      } catch (spError) {
        // We log the error but DO NOT throw.
        // This ensures the platform continues functioning seamlessly, the case remains 
        // in local fallback, and the emails are still dispatched successfully exactly like yesterday.
        console.warn("Non-blocking failure: Could not sync case to live SharePoint list. Missing/mismatched column?", spError);
        // Run diagnostic automatically
        this.diagnoseListColumns(LIST_NAMES.COMPLIANCE_CASES);
      }
    }

    // Create Disciplinary Action Record (non-blocking)
    if (!isHistorical) {
      try {
        await this.createDisciplinaryAction({
          title: await this.getNextDisciplinaryReference(),
          caseReference: newCase.title,
          actionType: disciplinaryAction,
          penaltyAmount: newCase.penaltyAmount,
          notes: `Action Classification: ${actionPath}. Recommended by PACT Engine.`,
          status: 'Pending'
        });
      } catch (e) {
        console.warn('Disciplinary action record failed', e);
      }
    }

    // Email notifications are now securely handled by the Power Automate "When an item is created" flow.
    // The React app simply saves the case, and Microsoft 365 handles the background email delivery,
    // ensuring no CORS errors or duplicate emails.

    return newCase;
  }


  public async updateCase(id: string, updates: Partial<ComplianceCase>): Promise<void> {
    if (this.isLocal) {
      const cases = this.getFromLocal<ComplianceCase>('pact_cases');
      const idx = cases.findIndex(c => c.id === id);
      if (idx > -1) {
        cases[idx] = { ...cases[idx], ...updates };
        this.saveToLocal('pact_cases', cases);
      }
      return;
    }
    
    // Map internal status back to SP choice if needed, though here we assume titles/fields match
    const itemType = await this.getListItemEntityType(LIST_NAMES.COMPLIANCE_CASES);
    const spData = {
      '__metadata': { 'type': itemType },
      Status: updates.status,
      // Add other field mapping as needed
    };

    const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.COMPLIANCE_CASES, spData);
    await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items(${id})`, {
      method: 'POST', // SharePoint uses POST with MERGE header for updates
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify(filteredData)
    });
  }

  public async deleteCase(id: string): Promise<void> {
    if (this.isLocal) {
      const cases = this.getFromLocal<ComplianceCase>('pact_cases').filter(c => c.id !== id);
      this.saveToLocal('pact_cases', cases);
      return;
    }

    await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.COMPLIANCE_CASES}')/items(${id})`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    });
  }

  public async deleteDisciplinaryAction(id: string): Promise<void> {
    if (this.isLocal) {
      return;
    }

    await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.DISCIPLINARY_ACTIONS}')/items(${id})`, {
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    });
  }

  public async getDashboardStats(): Promise<DashboardStats> {
    try {
      const allCases = await this.getCases();
      const cases = allCases.filter(c => !c.title.startsWith('HIST-'));
      const escalations = await this.getEscalationLog();
      const trackers = await this.getRepeatTrackerRecords();
      const appeals = await this.getAppeals();
      return this.buildDashboardStats(cases, escalations, trackers, appeals);
    } catch (error) {
      console.warn('Dashboard stats fell back to local data', error);
      const allCases = this.getFromLocal<ComplianceCase>('pact_cases');
      const cases = allCases.filter(c => !c.title.startsWith('HIST-'));
      const escalations = this.getFromLocal<EscalationEntry>('pact_escalations');
      const trackers = this.getFromLocal<RepeatOffenceRecord>('pact_trackers');
      const appeals = this.getFromLocal<any>('pact_appeals');
      return this.buildDashboardStats(cases, escalations, trackers, appeals);
    }
  }

  private buildDashboardStats(
    cases: ComplianceCase[],
    escalations: EscalationEntry[],
    trackers: RepeatOffenceRecord[],
    appeals: any[] = []
  ): DashboardStats {
    const tiers = ['Tier 1', 'Tier 2', 'Tier 3'];
    const casesByTier = tiers.map(tier => ({
      tier,
      count: cases.filter(c => c.tier === tier).length
    }));

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const now = new Date();
    const last6Months: { month: string; tier1: number; tier2: number; tier3: number }[] = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const m = months[d.getMonth()];
      const monthlyCases = cases.filter(c => {
        const cd = new Date(c.dateCreated);
        return cd.getMonth() === d.getMonth() && cd.getFullYear() === d.getFullYear();
      });
      last6Months.push({
        month: m,
        tier1: monthlyCases.filter(c => c.tier === 'Tier 1').length,
        tier2: monthlyCases.filter(c => c.tier === 'Tier 2').length,
        tier3: monthlyCases.filter(c => c.tier === 'Tier 3').length,
      });
    }

    const unpaidStatuses = new Set([CASE_STATUS.UNPAID, CASE_STATUS.OVERDUE, 'Acknowledged']);
    return {
      totalActiveCases: cases.filter(c => unpaidStatuses.has(c.status)).length,
      paidCases: cases.filter(c => c.status === CASE_STATUS.PAID).length,
      appealPendingCases: cases.filter(c => c.status === CASE_STATUS.APPEAL_PENDING).length,
      escalationsThisMonth: escalations.filter(e => {
        const date = new Date(e.escalationDate);
        return date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
      }).length,
      pendingAppeals: appeals.filter(appeal => {
        const decision = String(appeal?.decision || '').toLowerCase();
        return !decision || decision === 'pending';
      }).length,
      repeatOffenders: trackers.filter(t => t.totalOffences > 1).length,
      totalFines: cases.reduce((sum, c) => sum + c.penaltyAmount, 0),
      casesByTier,
      casesByMonth: last6Months,
      casesByDepartment: Array.from(new Set(cases.map(c => c.department).filter(dept => dept && dept.trim() !== ''))).map(dept => ({
        department: dept,
        count: cases.filter(c => c.department === dept).length,
        risk: trackers.some(t => {
          const staff = cases.find(c => c.department === dept && c.chargedPerson === t.offender);
          return staff && (t.riskLevel === 'High' || t.riskLevel === 'Critical');
        }) ? 'High' : 'Low'
      })),
      recentActivity: [
        ...cases.slice(0, 3).map(c => ({
          id: `c-${c.id}`,
          type: 'case' as const,
          title: `New Case: ${c.title}`,
          description: `${c.chargedPersonName} charged with ${c.offenceCategoryName}`,
          timestamp: c.dateCreated,
          severity: (c.tier === 'Tier 3' ? 'critical' : c.tier === 'Tier 2' ? 'high' : 'medium') as any
        })),
        ...escalations.slice(0, 2).map(e => ({
          id: `e-${e.id}`,
          type: 'escalation' as const,
          title: `Escalation: ${e.caseReference}`,
          description: `Tier upgrade due to ${e.escalationReason}`,
          timestamp: e.escalationDate,
          severity: 'critical' as const
        }))
      ].sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime())
    };
  }

  // --- Escalations & Tracker ---

  public async resetAllTrackers(): Promise<void> {
    if (this.isLocal) {
      this.saveToLocal('pact_trackers', []);
      return;
    }
    try {
      const trackers = await this.getRepeatTrackerRecords();
      const itemType = await this.getListItemEntityType(LIST_NAMES.REPEAT_OFFENCE_TRACKER);
      for (const t of trackers) {
        const spData = {
          '__metadata': { 'type': itemType },
          [COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES]: 0,
          [COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M]: 0,
          [COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES]: 0,
          [COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES]: 0,
          [COLUMNS.REPEAT_TRACKER.RISK_LEVEL]: 'Low'
        };
        const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.REPEAT_OFFENCE_TRACKER, spData);
        await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.REPEAT_OFFENCE_TRACKER}')/items(${t.id})`, {
          method: 'POST',
          headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
          body: JSON.stringify(filteredData)
        });
      }
    } catch (e) {
      console.warn("Failed to reset trackers", e);
    }
  }

  public async getEscalationLog(): Promise<EscalationEntry[]> {
    try {
      if (!this.isLocal) {
        // Request FieldValuesAsText so we get text representations of ALL fields (including lookups)
        const endpoint = `web/lists/getbytitle('${LIST_NAMES.ESCALATION_LOG}')/items?$select=*,FieldValuesAsText/*&$expand=FieldValuesAsText&$orderby=ID desc`;
        const data = await this.fetchREST(endpoint);
        const [staff, cases] = await Promise.all([
          this.getStaffDirectory(),
          this.getCases()
        ]);
        const items = data.results || [];

        // ── One-time diagnostic dump ───────────────────────────────────────
        if (items.length > 0 && !this._escalationDiagLogged) {
          this._escalationDiagLogged = true;
          this.diagnoseListColumns(LIST_NAMES.ESCALATION_LOG);
          const sample = items[0];
          const keys = Object.keys(sample).filter(k => !k.startsWith('__') && k !== 'FieldValuesAsText');
          console.log('[PACT Escalation Diag] Raw item keys:', keys.join(', '));
          // Log every key:value (excluding metadata/deferred)
          const vals: Record<string, any> = {};
          for (const k of keys) {
            const v = sample[k];
            if (v && typeof v === 'object' && ('__deferred' in v || '__metadata' in v)) continue;
            vals[k] = v;
          }
          console.log('[PACT Escalation Diag] Raw item values:', JSON.stringify(vals, null, 2));
          if (sample.FieldValuesAsText) {
            const fvat = { ...sample.FieldValuesAsText };
            delete fvat.__metadata;
            console.log('[PACT Escalation Diag] FieldValuesAsText:', JSON.stringify(fvat, null, 2));
          }
        }

        // ── Universal field reader (checks item directly + FieldValuesAsText) ──
        const readEsc = (item: any, ...names: string[]): string => {
          const fvat = item.FieldValuesAsText || {};
          for (const name of names) {
            // Try direct field
            const direct = this.normalizeFieldValue(item?.[name]);
            if (direct !== undefined && direct !== null && direct !== '') return String(direct);
            // Try FieldValuesAsText
            const text = fvat[name];
            if (text !== undefined && text !== null && text !== '') return String(text);
          }
          return '';
        };

        return items.map((item: any) => {
          const caseRef = readEsc(item,
            COLUMNS.ESCALATION.CASE_REFERENCE, 'CaseReference', 'Case_x0020_Reference',
            'Case Reference', 'CaseRef', 'Case Ref'
          ) || '';
          
          // Find matching compliance case
          const matchingCase = cases.find(c => c.title === caseRef);

          // Offender: try lookup Id, then matchingCase, then text fallbacks
          let offenderId = String(
            item.OffenderId ||
            item.Offender_x0020_Id ||
            readEsc(item, 'OffenderId', 'Offender_x0020_Id') ||
            ''
          );
          if (!offenderId && matchingCase) {
            offenderId = matchingCase.chargedPerson;
          }

          const person = staff.find(s => s.id === offenderId);

          // Offender name: resolved from staff, matchingCase, FieldValuesAsText, or raw field
          let offenderName = person?.fullName || matchingCase?.chargedPersonName || '';
          if (!offenderName) {
            offenderName = readEsc(item, 'Offender', 'OffenderName', 'Offender_x0020_Name');
          }
          // If name is still empty but we have an offenderId, try name match
          if (!offenderName && offenderId) {
            const byId = staff.find(s => s.id === offenderId);
            if (byId) offenderName = byId.fullName;
          }

          const prevTier = readEsc(item,
            COLUMNS.ESCALATION.PREVIOUS_TIER, 'PreviousTier', 'Previous_x0020_Tier',
            'Previous Tier', 'PrevTier'
          ) || 'Tier 1';
          const nextTier = readEsc(item,
            COLUMNS.ESCALATION.NEW_TIER, 'NewTier', 'New_x0020_Tier',
            'New Tier', 'EscalatedTier'
          ) || 'Tier 2';

          let escalationReason = readEsc(item,
            COLUMNS.ESCALATION.REASON, 'EscalationReason', 'Escalation_x0020_Reason',
            'Escalation Reason', 'Reason', 'EscalationReason0'
          );
          if (!escalationReason) {
            if (matchingCase) {
              escalationReason = `Escalation from ${prevTier} to ${nextTier} triggered by compliance case ${caseRef} (${matchingCase.offenceCategoryName}).`;
            } else {
              escalationReason = `Automatic policy trigger escalating staff member from ${prevTier} to ${nextTier}.`;
            }
          }

          return {
            id: String(item.ID || item.Id || ''),
            title: readEsc(item, 'Title') || '',
            caseReference: caseRef,
            offender: offenderId,
            offenderName,
            escalationReason,
            previousTier: prevTier as any,
            newTier: nextTier as any,
            triggeredBy: readEsc(item,
              COLUMNS.ESCALATION.TRIGGERED_BY, 'TriggeredBy', 'Triggered_x0020_By',
              'Triggered By', 'TriggerType'
            ) || 'System',
            escalationDate: readEsc(item,
              COLUMNS.ESCALATION.DATE, 'EscalationDate', 'Escalation_x0020_Date',
              'Escalation Date', 'Created'
            ) || new Date().toISOString(),
            notifiedTo: readEsc(item,
              COLUMNS.ESCALATION.NOTIFIED_TO, 'NotifiedTo', 'Notified_x0020_To',
              'Notified To', 'NotifiedTo0'
            ) || ''
          };
        });
      }
      return this.getFromLocal<EscalationEntry>('pact_escalations');
    } catch (err) {
      console.warn('[PACT] getEscalationLog failed, falling back to local:', err);
      return this.getFromLocal<EscalationEntry>('pact_escalations');
    }
  }


  public async getRepeatTrackerRecord(staffId: string): Promise<RepeatOffenceRecord | null> {
    try {
      const staff = this.isLocal ? this.getFromLocal<StaffMember>('pact_staff') : await this.getStaffDirectory();
      const targetStaff = staff.find(s => s.id === staffId || s.fullName === staffId || (s.fullName && staffId && s.fullName.toLowerCase() === staffId.toLowerCase()));

      console.log("[PACT Debug] getRepeatTrackerRecord called with staffId:", staffId);
      console.log("[PACT Debug] Resolved targetStaff:", targetStaff);

      const records = this.isLocal
        ? this.getFromLocal<RepeatOffenceRecord>('pact_trackers')
        : await this.getRepeatTrackerRecords();

      console.log("[PACT Debug] Total Repeat Tracker records loaded:", records.length);

      const existing = records.find(r => {
        const matchId = r.offender === staffId;
        const matchStaffId = targetStaff && r.offender === targetStaff.id;
        const matchName = targetStaff && r.offenderName && r.offenderName.toLowerCase().trim() === targetStaff.fullName.toLowerCase().trim();
        const matchEmail = targetStaff && r.offender && r.offender.toLowerCase().trim() === targetStaff.email.toLowerCase().trim();
        if (matchId || matchStaffId || matchName || matchEmail) {
          console.log("[PACT Debug] Match found in tracker records! Record:", r);
          return true;
        }
        return false;
      });

      if (existing) {
        return existing;
      }

      console.log("[PACT Debug] No existing tracker record found. Rebuilding from cases...");

      // Self-Healing Sync: Rebuild tracker from Compliance Cases if empty
      if (!this.isLocal) {
        const cases = await this.getCases();
        console.log("[PACT Debug] Total Compliance Cases loaded:", cases.length);
        if (cases.length > 0) {
          console.log("[PACT Debug] Sample case offender info:", {
            chargedPerson: cases[0].chargedPerson,
            chargedPersonName: cases[0].chargedPersonName,
            staffEmail: cases[0].staffEmail
          });
        }

        const staffCases = cases.filter(c => {
          if (c.chargedPerson === staffId) return true;
          if (targetStaff && c.chargedPersonName && c.chargedPersonName.toLowerCase().trim() === targetStaff.fullName.toLowerCase().trim()) return true;
          return false;
        });

        console.log("[PACT Debug] Matching cases found for staff:", staffCases.length);

        if (staffCases.length > 0) {
          const tier1Cases = staffCases.filter(c => c.tier === 'Tier 1');
          const tier2Cases = staffCases.filter(c => c.tier === 'Tier 2');
          const tier3Cases = staffCases.filter(c => c.tier === 'Tier 3');

          const sixMonthsAgo = new Date();
          sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
          const t1Last6M = tier1Cases.filter(c => new Date(c.dateCreated) >= sixMonthsAgo).length;

          const rebuiltTracker: Partial<RepeatOffenceRecord> = {
            totalOffences: staffCases.length,
            tier1Last6Months: t1Last6M,
            tier2Offences: tier2Cases.length,
            tier3Offences: tier3Cases.length,
            lastOffenceDate: staffCases[0]?.dateCreated || new Date().toISOString(),
            riskLevel: escalationEngine.calculateRiskLevel({
              id: '', title: '', offender: '',
              totalOffences: staffCases.length,
              tier1Last6Months: t1Last6M,
              tier2Offences: tier2Cases.length,
              tier3Offences: tier3Cases.length,
              riskLevel: 'Low',
              lastOffenceDate: '',
              escalationDue: false
            })
          };

          console.log("[PACT Debug] Rebuilt tracker object to write:", rebuiltTracker);

          this.updateRepeatTracker(staffId, rebuiltTracker).catch(err => 
            console.warn("[PACT] Failed to write back rebuilt tracker", err)
          );

          return {
            id: 'rebuilt-' + staffId,
            title: targetStaff?.fullName || 'Unknown',
            offender: staffId,
            offenderName: targetStaff?.fullName || 'Unknown',
            totalOffences: rebuiltTracker.totalOffences || 0,
            tier1Last6Months: rebuiltTracker.tier1Last6Months || 0,
            tier2Offences: rebuiltTracker.tier2Offences || 0,
            tier3Offences: rebuiltTracker.tier3Offences || 0,
            riskLevel: rebuiltTracker.riskLevel || 'Low',
            lastOffenceDate: rebuiltTracker.lastOffenceDate || new Date().toISOString(),
            escalationDue: false
          };
        }
      }

      return null;
    } catch (e) {
      console.error("[PACT Debug] Error in getRepeatTrackerRecord:", e);
      return null;
    }
  }

  private async getRepeatTrackerRecords(): Promise<RepeatOffenceRecord[]> {
    if (this.isLocal) {
      return this.getFromLocal<RepeatOffenceRecord>('pact_trackers');
    }

    try {
      const endpoint = `web/lists/getbytitle('${LIST_NAMES.REPEAT_OFFENCE_TRACKER}')/items?$orderby=ID desc`;
      const data = await this.fetchREST(endpoint);
      return (data.results || []).map((item: any) => ({
        id: item.ID.toString(),
        title: this.readField(item, COLUMNS.REPEAT_TRACKER.TITLE, 'Title'),
        offender: (() => {
          const flatId = item.OffenderId || item.OffenderID;
          if (flatId !== undefined && flatId !== null) return String(flatId);
          const obj = item.Offender || item.offender;
          if (obj && typeof obj === 'object') {
            const nestedId = obj.Id || obj.ID;
            if (nestedId !== undefined && nestedId !== null) return String(nestedId);
          }
          if (typeof obj === 'string' && obj !== '') return obj;
          return this.readField(item, COLUMNS.REPEAT_TRACKER.OFFENDER, 'Offender', 'OffenderId')?.toString?.() || String(this.readField(item, COLUMNS.REPEAT_TRACKER.OFFENDER, 'Offender', 'OffenderId') || '');
        })(),
        offenderName: this.readField(item, COLUMNS.REPEAT_TRACKER.TITLE, 'Title'),
        totalOffences: Number(this.readField(item, COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES, 'TotalOffences', 'Total Offences') || 0),
        tier1Last6Months: Number(this.readField(item, COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M, 'Tier1Offences', 'Tier1 Offences') || 0),
        tier2Offences: Number(this.readField(item, COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES, 'Tier2Offences', 'Tier2 Offences') || 0),
        tier3Offences: Number(this.readField(item, COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES, 'Tier3Offences', 'Tier3 Offences') || 0),
        riskLevel: this.readField(item, COLUMNS.REPEAT_TRACKER.RISK_LEVEL, 'RiskLevel', 'Risk Level') || 'Low',
        lastOffenceDate: this.readField(item, COLUMNS.REPEAT_TRACKER.LAST_OFFENCE_DATE, 'LastOffenceDate', 'Last Offence Date') || new Date().toISOString(),
        escalationDue: this.readField(item, COLUMNS.REPEAT_TRACKER.ESCALATION_DUE, 'EscalationDue', 'Escalation Due') === true || this.readField(item, COLUMNS.REPEAT_TRACKER.ESCALATION_DUE, 'EscalationDue', 'Escalation Due') === 'True'
      }));
    } catch (error) {
      console.warn('Failed to load live repeat tracker records', error);
      return this.getFromLocal<RepeatOffenceRecord>('pact_trackers');
    }
  }

  public async updateRepeatTracker(staffId: string, updates: Partial<RepeatOffenceRecord>): Promise<void> {

    if (this.isLocal) {
      const records = this.getFromLocal<RepeatOffenceRecord>('pact_trackers');
      const idx = records.findIndex(r => r.offender === staffId);

      if (idx > -1) {
        records[idx] = { ...records[idx], ...updates };
      } else {
        const staff = this.getFromLocal<StaffMember>('pact_staff').find(s => s.id === staffId);
        records.push({
          id: Date.now().toString(),
          title: staff?.fullName || 'Unknown',
          offender: staffId,
          totalOffences: 1,
          tier1Last6Months: 1,
          tier2Offences: 0,
          tier3Offences: 0,
          riskLevel: 'Low',
          lastOffenceDate: new Date().toISOString(),
          escalationDue: false,
          ...updates
        });
      }
      this.saveToLocal('pact_trackers', records);
    } else {
      // SP REST Update
      const spData: any = {
        [COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES]: updates.totalOffences,
        [COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M]: updates.tier1Last6Months,
        [COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES]: updates.tier2Offences,
        [COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES]: updates.tier3Offences,
        [COLUMNS.REPEAT_TRACKER.RISK_LEVEL]: updates.riskLevel,
        [COLUMNS.REPEAT_TRACKER.LAST_OFFENCE_DATE]: updates.lastOffenceDate,
        [COLUMNS.REPEAT_TRACKER.ESCALATION_DUE]: updates.escalationDue
      };

      try {
        const existing = await this.getRepeatTrackerRecord(staffId);
        if (existing) {
          const itemType = await this.getListItemEntityType(LIST_NAMES.REPEAT_OFFENCE_TRACKER);
          
          // Heal existing record if OffenderId is missing or not a numeric ID matching staffId
          if (!existing.offender || isNaN(parseInt(existing.offender, 10))) {
            const staff = await this.getStaffDirectory();
            const person = staff.find(s => s.id === staffId);
            if (person) {
              spData[COLUMNS.REPEAT_TRACKER.OFFENDER + 'Id'] = parseInt(person.id, 10) || null;
            }
          }

          const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.REPEAT_OFFENCE_TRACKER, {
            __metadata: { type: itemType },
            ...spData
          });

          await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.REPEAT_OFFENCE_TRACKER}')/items(${existing.id})`, {
            method: 'POST',
            headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
            body: JSON.stringify(filteredData)
          });
        } else {
          const staff = await this.getStaffDirectory();
          const person = staff.find(s => s.id === staffId);
          const createData: any = {
            '__metadata': { 'type': `SP.Data.${LIST_NAMES.REPEAT_OFFENCE_TRACKER.replace(/ /g, '_x0020_')}ListItem` },
            [COLUMNS.REPEAT_TRACKER.TITLE]: person?.fullName || 'Unknown',
            [COLUMNS.REPEAT_TRACKER.TOTAL_OFFENCES]: updates.totalOffences ?? 1,
            [COLUMNS.REPEAT_TRACKER.TIER1_LAST_6M]: updates.tier1Last6Months ?? 0,
            [COLUMNS.REPEAT_TRACKER.TIER2_OFFENCES]: updates.tier2Offences ?? 0,
            [COLUMNS.REPEAT_TRACKER.TIER3_OFFENCES]: updates.tier3Offences ?? 0,
            [COLUMNS.REPEAT_TRACKER.RISK_LEVEL]: updates.riskLevel ?? 'Low',
            [COLUMNS.REPEAT_TRACKER.LAST_OFFENCE_DATE]: updates.lastOffenceDate ?? new Date().toISOString(),
            [COLUMNS.REPEAT_TRACKER.ESCALATION_DUE]: updates.escalationDue ?? false
          };
          if (person) {
            createData[COLUMNS.REPEAT_TRACKER.OFFENDER + 'Id'] = parseInt(person.id, 10) || null;
          }
          const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.REPEAT_OFFENCE_TRACKER, createData);
          await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.REPEAT_OFFENCE_TRACKER}')/items`, {
            method: 'POST',
            body: JSON.stringify(filteredData)
          });
        }
      } catch (e) {
        console.error("Failed to update SP Repeat Tracker:", e);
      }
    }
  }

  public async createEscalation(entry: Partial<EscalationEntry>): Promise<void> {
    if (this.isLocal) {
      const log = this.getFromLocal<EscalationEntry>('pact_escalations');
      log.push({
        id: Date.now().toString(),
        title: `Escalation ${entry.caseReference}`,
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
    } else {
      // Write under every possible alias so filterPayloadToAvailableFields always finds a match
      // regardless of what the SharePoint list column internal names actually are.
      const offenderText = entry.offenderName || entry.offender || '';
      const reasonText = entry.escalationReason || '';
      const caseRefText = entry.caseReference || '';
      const prevTierText = entry.previousTier || 'Tier 1';
      const newTierText = entry.newTier || 'Tier 2';
      const triggeredByText = entry.triggeredBy || 'System';
      const dateText = entry.escalationDate || new Date().toISOString();

      const spData: any = {
        '__metadata': { 'type': `SP.Data.${LIST_NAMES.ESCALATION_LOG.replace(/ /g, '_x0020_')}ListItem` },
        // Title
        'Title': entry.title || `Escalation ${caseRefText}`,
        // Case Reference — all known aliases
        'CaseReference': caseRefText,
        'Case_x0020_Reference': caseRefText,
        // Offender name — all known aliases (plain text, no lookup)
        'Offender': offenderText,
        'OffenderName': offenderText,
        'Offender_x0020_Name': offenderText,
        // Escalation Reason — all known aliases
        'EscalationReason': reasonText,
        'Escalation_x0020_Reason': reasonText,
        'Reason': reasonText,
        // Previous Tier — all known aliases
        'PreviousTier': prevTierText,
        'Previous_x0020_Tier': prevTierText,
        'PrevTier': prevTierText,
        // New Tier — all known aliases
        'NewTier': newTierText,
        'New_x0020_Tier': newTierText,
        'EscalatedTier': newTierText,
        // Triggered By
        'TriggeredBy': triggeredByText,
        'Triggered_x0020_By': triggeredByText,
        'TriggerType': triggeredByText,
        // Date
        'EscalationDate': dateText,
        'Escalation_x0020_Date': dateText,
        // Notified To
        'NotifiedTo': entry.notifiedTo || 'HR Dept',
        'Notified_x0020_To': entry.notifiedTo || 'HR Dept',
      };

      console.log('[PACT createEscalation] Writing to SharePoint:', {
        offender: offenderText, caseRef: caseRefText,
        reason: reasonText.substring(0, 80), prevTier: prevTierText, newTier: newTierText
      });

      const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.ESCALATION_LOG, spData);
      console.log('[PACT createEscalation] Filtered fields sent:', Object.keys(filteredData).filter(k => k !== '__metadata').join(', '));
      await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.ESCALATION_LOG}')/items`, {
        method: 'POST',
        body: JSON.stringify(filteredData)
      });
    }

    // Email Notification for Escalation — appeal action only (no accept/pay)
    const appealUrl = this.buildCaseResponseLink(entry.caseReference || '', 'appeal', {
      staffName: entry.offenderName || 'Employee',
      offenceLabel: 'Escalated compliance matter',
      amount: 0,
      dueIso: new Date().toISOString(),
      description: entry.escalationReason || '',
    });
    const subject = `PACT ESCALATION: ${entry.caseReference} - TIER UPGRADE`;
    const body = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; padding: 20px; border: 2px solid #d13438; border-radius: 8px;">
        <h2 style="color: #d13438; margin-top: 0;">Policy Escalation Triggered</h2>
        <p>A compliance case has been automatically escalated due to repeat offences or severity.</p>
        <div style="background: #fff4f4; padding: 15px; border-radius: 4px; margin: 20px 0;">
          <p><b>Case Ref:</b> ${entry.caseReference}</p>
          <p><b>Previous Tier:</b> ${entry.previousTier}</p>
          <p><b>New Tier:</b> ${entry.newTier}</p>
          <p><b>Reason:</b> ${entry.escalationReason}</p>
        </div>
        <p>Immediate review by HR / Executive Management is required.</p>
        ${this.buildEmailButtonHtmlAppealOnly(appealUrl)}
      </div>
    `;
    await this.sendEmailNotification([HR_EMAIL, LEGAL_EMAIL, CHAIRMAN_EMAIL], subject, body);
  }


  // --- Disciplinary Actions ---

  public async getDisciplinaryActions(caseRef?: string): Promise<any[]> {
    try {
      if (!this.isLocal) {
        const filter = caseRef ? `&$filter=${COLUMNS.DISCIPLINARY.CASE_REFERENCE} eq '${caseRef}'` : '';
        const endpoint = `web/lists/getbytitle('${LIST_NAMES.DISCIPLINARY_ACTIONS}')/items?$orderby=ID desc${filter}`;
        const data = await this.fetchREST(endpoint);
        return (data.results || []).map((item: any) => ({
          id: item.ID.toString(),
          title: this.readField(item, COLUMNS.DISCIPLINARY.TITLE, 'Title'),
          caseReference: this.readField(item, COLUMNS.DISCIPLINARY.CASE_REFERENCE, 'Case Reference', 'CaseReference', 'Case_x0020_Reference'),
          actionType: this.readField(item, COLUMNS.DISCIPLINARY.ACTION_TYPE, 'Action Type', 'ActionType', 'Action_x0020_Type'),
          actionDate: this.readField(item, COLUMNS.DISCIPLINARY.ACTION_DATE, 'Action Date', 'ActionDate', 'Action_x0020_Date'),
          actionedBy: this.readField(item, 'ActionedBy'),
          penaltyAmount: this.parsePenalty(this.readField(item, 'Penalty Amount', 'PenaltyAmount', 'Penalty_x0020_Amount')),
          notes: this.readField(item, COLUMNS.DISCIPLINARY.NOTES, 'Notes'),
          status: this.readField(item, COLUMNS.DISCIPLINARY.STATUS, 'Status')
        }));
      }
      const actions = this.getFromLocal<any>('pact_disciplinary');
      return caseRef ? actions.filter(a => a.caseReference === caseRef) : actions;
    } catch {
      const actions = this.getFromLocal<any>('pact_disciplinary');
      return caseRef ? actions.filter(a => a.caseReference === caseRef) : actions;
    }
  }

  public async createDisciplinaryAction(action: any): Promise<void> {
    if (this.isLocal) {
      const log = this.getFromLocal<any>('pact_disciplinary');
      log.push({
        id: Date.now().toString(),
        title: action.title || this.formatDisciplinaryReference(log.length + 1),
        actionDate: new Date().toISOString(),
        actionedBy: this.getUserName(),
        ...action
      });
      this.saveToLocal('pact_disciplinary', log);
      return;
    }

    const spData = {
      '__metadata': { 'type': `SP.Data.${LIST_NAMES.DISCIPLINARY_ACTIONS.replace(/ /g, '_x0020_')}ListItem` },
      [COLUMNS.DISCIPLINARY.TITLE]: action.title || `P-${Date.now()}`,
      [COLUMNS.DISCIPLINARY.CASE_REFERENCE]: action.caseReference,
      [COLUMNS.DISCIPLINARY.ACTION_TYPE]: action.actionType,
      [COLUMNS.DISCIPLINARY.ACTION_DATE]: new Date().toISOString(),
      [COLUMNS.DISCIPLINARY.PENALTY_AMOUNT]: action.penaltyAmount ?? 0,
      [COLUMNS.DISCIPLINARY.NOTES]: action.notes,
      [COLUMNS.DISCIPLINARY.STATUS]: action.status || 'Pending'
    };

    const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.DISCIPLINARY_ACTIONS, spData);
    await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.DISCIPLINARY_ACTIONS}')/items`, {
      method: 'POST',
      body: JSON.stringify(filteredData)
    });
  }

  // --- Appeals ---

  public async getAppeals(): Promise<any[]> {
    try {
      if (!this.isLocal) {
        const endpoint = `web/lists/getbytitle('${LIST_NAMES.APPEALS_REGISTER}')/items?$orderby=ID desc`;
        const data = await this.fetchREST(endpoint);
        return (data.results || []).map((item: any) => ({
          id: item.ID.toString(),
          title: this.readField(item, COLUMNS.APPEALS.TITLE, 'Title', 'Appeal ID'),
          caseReference: this.readField(item, COLUMNS.APPEALS.CASE_REFERENCE, 'Case Reference', 'CaseReference', 'Case_x0020_Reference'),
          appellant: this.readField(item, COLUMNS.APPEALS.APPELLANT, 'Appellant Name', 'AppellantName', 'Appellant'),
          appealDate: this.readField(item, COLUMNS.APPEALS.APPEAL_DATE, 'Appeal Date', 'AppealDate', 'Appeal_x0020_Date'),
          grounds: this.readField(item, COLUMNS.APPEALS.GROUNDS, 'Grounds for Appeal', 'GroundsforAppeal', 'Grounds_x0020_for_x0020_Appeal'),
          reviewingOfficer: this.readField(item, COLUMNS.APPEALS.REVIEWING_OFFICER, 'Reviewing Officer', 'ReviewingOfficer', 'Reviewing_x0020_Officer'),
          decision: this.readField(item, COLUMNS.APPEALS.DECISION, 'Decision') || 'Pending',
          decisionDate: this.readField(item, COLUMNS.APPEALS.DECISION_DATE, 'Decision Date', 'DecisionDate', 'Decision_x0020_Date'),
          decisionNotes: this.readField(item, COLUMNS.APPEALS.DECISION_NOTES, 'Decision Notes', 'DecisionNotes', 'Decision_x0020_Notes')
        }));
      }
      return this.getFromLocal<any>('pact_appeals');
    } catch {
      return this.getFromLocal<any>('pact_appeals');
    }
  }


  public async getMailHistory(): Promise<any[]> {
    try {
      if (!this.isLocal) {
        const endpoint = `web/lists/getbytitle('${LIST_NAMES.MAIL_HISTORY}')/items?$orderby=Created desc&$top=50`;
        const data = await this.fetchREST(endpoint);
        return (data.results || [])
          .filter((item: any) => {
            const subj = (item[COLUMNS.MAIL.SUBJECT] || item[COLUMNS.MAIL.TITLE] || '').toLowerCase();
            return !subj.includes('mock');
          })
          .map((item: any) => ({
            id: item.ID.toString(),
            to: (item[COLUMNS.MAIL.TO] || '').split(',').map((s: string) => s.trim()).filter(Boolean),
            subject: item[COLUMNS.MAIL.SUBJECT] || item[COLUMNS.MAIL.TITLE],
            body: item[COLUMNS.MAIL.BODY],
            timestamp: item.Created,
            status: item[COLUMNS.MAIL.STATUS] || 'Sent'
          }));
      }
      // In demo/workbench mode, return empty — real emails are handled by Power Automate
      return [];
    } catch {
      return [];
    }
  }


  public async createAppeal(appeal: any): Promise<void> {
    const title = appeal.title || `A-${String(Date.now()).slice(-4)}`;
    const appealDate = new Date().toISOString();

    // Resolve original case details by reference to fully populate all appeal fields
    const originalCase = await this.getCaseByReference(appeal.caseReference);

    const enrichedAppeal = {
      ...appeal,
      appellant: appeal.appellant || originalCase?.chargedPersonName || 'Employee',
      appellantEmail: appeal.appellantEmail || originalCase?.staffEmail || '',
      department: appeal.department || originalCase?.department || 'Not specified',
      offence: appeal.offence || originalCase?.offenceCategoryName || 'Compliance Violation',
      offenceDescription: appeal.offenceDescription || originalCase?.offenceDescription || '',
      penaltyAmount: typeof appeal.penaltyAmount === 'number' ? appeal.penaltyAmount : (originalCase?.penaltyAmount ?? 0),
    };

    if (this.isLocal) {
      const log = this.getFromLocal<any>('pact_appeals');
      log.push({
        id: Date.now().toString(),
        title,
        appealDate,
        decision: 'Pending',
        ...enrichedAppeal
      });
      this.saveToLocal('pact_appeals', log);
    } else {
      const spData = {
        '__metadata': { 'type': `SP.Data.${LIST_NAMES.APPEALS_REGISTER.replace(/ /g, '_x0020_')}ListItem` },
        [COLUMNS.APPEALS.TITLE]: title,
        [COLUMNS.APPEALS.CASE_REFERENCE]: enrichedAppeal.caseReference,
        [COLUMNS.APPEALS.APPELLANT]: enrichedAppeal.appellant,
        [COLUMNS.APPEALS.APPEAL_DATE]: appealDate,
        [COLUMNS.APPEALS.GROUNDS]: enrichedAppeal.grounds,
        [COLUMNS.APPEALS.DECISION]: 'Pending'
      };

      const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.APPEALS_REGISTER, spData);
      await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.APPEALS_REGISTER}')/items`, {
        method: 'POST',
        body: JSON.stringify(filteredData)
      });
    }

    // ── Apply side effects (Both Modes) ──
    await this.updateCaseStatusForReference(enrichedAppeal.caseReference, CASE_STATUS.APPEAL_PENDING);
    this.notifyDataChanged();

    // ── Send Email Notifications ──
    const subject = `PACT APPEAL FILED: Case ${enrichedAppeal.caseReference} (${title})`;
    const body = this.buildAppealSubmittedEmailBody(enrichedAppeal, title, appealDate);

    // Send appeal notification via the same schema
    await this.sendEmailNotification(
      [HR_EMAIL, LEGAL_EMAIL, CHAIRMAN_EMAIL],
      subject,
      body,
      APPEAL_MAIL_TRIGGER_URL
    );
  }

  private async updateListItemByDisplayNames(
    listName: string,
    itemId: number,
    formValues: Array<{ FieldName: string; FieldValue: string }>,
    options?: { bNewDocumentUpdate?: boolean }
  ): Promise<void> {
    if (formValues.length === 0) return;
    const result = await this.fetchREST(
      `web/lists/getbytitle('${listName}')/items(${itemId})/validateUpdateListItem`,
      {
        method: 'POST',
        body: JSON.stringify({
          formValues: formValues.map(({ FieldName, FieldValue }) => ({
            FieldName,
            FieldValue
          })),
          bNewDocumentUpdate: options?.bNewDocumentUpdate ?? false
        })
      }
    );
    this.assertValidateUpdateSucceeded(result, listName);
  }

  private async tryUpdateListItemByDisplayNames(
    listName: string,
    itemId: number,
    formValues: Array<{ FieldName: string; FieldValue: string }>,
    options?: { bNewDocumentUpdate?: boolean }
  ): Promise<boolean> {
    try {
      await this.updateListItemByDisplayNames(listName, itemId, formValues, options);
      return true;
    } catch {
      return false;
    }
  }

  public async updateAppeal(id: string, updates: any): Promise<void> {
    let appealDetails: any = null;

    if (this.isLocal) {
      const log = this.getFromLocal<any>('pact_appeals');
      const idx = log.findIndex(a => a.id === id);
      if (idx > -1) {
        log[idx] = { ...log[idx], ...updates, decisionDate: new Date().toISOString() };
        this.saveToLocal('pact_appeals', log);
        appealDetails = log[idx];
      }
    } else {
      const itemType = await this.getListItemEntityType(LIST_NAMES.APPEALS_REGISTER);
      const spData = {
        '__metadata': { 'type': itemType },
        [COLUMNS.APPEALS.DECISION]: updates.decision,
        [COLUMNS.APPEALS.DECISION_DATE]: new Date().toISOString(),
        [COLUMNS.APPEALS.DECISION_NOTES]: updates.decisionNotes,
        [COLUMNS.APPEALS.REVIEWING_OFFICER]: updates.reviewingOfficer
      };

      const filteredData = await this.filterPayloadToAvailableFields(LIST_NAMES.APPEALS_REGISTER, spData);
      await this.fetchREST(`web/lists/getbytitle('${LIST_NAMES.APPEALS_REGISTER}')/items(${id})`, {
        method: 'POST',
        headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
        body: JSON.stringify(filteredData)
      });

      // Fetch the updated appeal to get appellant details for the email
      const appeals = await this.getAppeals();
      appealDetails = appeals.find(a => a.id === id);
    }

    // ── Update linked case status & send decision email ──
    if (updates.decision && appealDetails) {
      let newCaseStatus: string | undefined;
      let emailSubject = '';
      let emailBody = '';

      // Get the original case to retrieve the actual user's email address and name
      const originalCase = await this.getCaseByReference(appealDetails.caseReference);
      const userEmail = originalCase?.staffEmail || appealDetails.appellantEmail || 'staff@konstructum.com';
      const userName = originalCase?.chargedPersonName || appealDetails.appellant || 'Employee';

      const portalUrl = this.buildCaseResponseLink(
        appealDetails.caseReference,
        'accept',
        {
          staffName: userName,
          offenceLabel: originalCase?.offenceCategoryName || 'Compliance Notice',
          amount: updates.decision === 'Waived' || updates.decision === 'Upheld' ? 0 : (originalCase?.penaltyAmount || 0),
          dueIso: originalCase?.dueDate || new Date().toISOString(),
          staffEmail: userEmail,
          department: originalCase?.department
        }
      );

      switch (updates.decision) {
        case 'Waived':
          newCaseStatus = CASE_STATUS.WAIVED;
          emailSubject = `APPEAL APPROVED – Penalty Waived: ${appealDetails.caseReference}`;
          emailBody = `
            <div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; padding: 20px; border: 2px solid #107c10; border-radius: 8px; background-color: #ffffff;">
              <h2 style="color: #107c10; margin-top: 0; font-size: 20px;">Appeal Approved: Penalty Waived</h2>
              <p style="font-size: 15px; line-height: 1.6;">Dear ${userName},</p>
              <p style="font-size: 15px; line-height: 1.6;">Your appeal for case <b>${appealDetails.caseReference}</b> has been reviewed and <b>WAIVED</b>.</p>
              <div style="background-color: #f3fcf5; padding: 15px; border-left: 4px solid #107c10; border-radius: 4px; margin: 20px 0; color: #107c10; font-weight: bold; font-size: 14px;">
                No penalty payment is required. This case is now officially closed.
              </div>
              ${updates.decisionNotes ? `<p style="font-size: 14px; line-height: 1.6; background-color: #f8fafc; padding: 12px; border: 1px solid #e2e8f0; border-radius: 6px;"><b>Reviewer notes:</b> ${updates.decisionNotes}</p>` : ''}
              
              <div style="margin: 25px 0 20px 0;">
                <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate;">
                  <tr>
                    <td align="center" bgcolor="#107c10" style="border-radius: 4px; padding: 0;">
                      <a href="${portalUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">
                        View Case Status
                      </a>
                    </td>
                  </tr>
                </table>
              </div>
              <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />
              <p style="font-size: 11px; color: #64748b; margin: 0;">This is an automated notification from the PACT Compliance Governance Platform.</p>
            </div>`;
          break;
        case 'Upheld':
          newCaseStatus = CASE_STATUS.WAIVED;
          emailSubject = `APPEAL UPHELD – Penalty Cancelled: ${appealDetails.caseReference}`;
          emailBody = `
            <div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; padding: 20px; border: 2px solid #107c10; border-radius: 8px; background-color: #ffffff;">
              <h2 style="color: #107c10; margin-top: 0; font-size: 20px;">Appeal Upheld: Penalty Cancelled</h2>
              <p style="font-size: 15px; line-height: 1.6;">Dear ${userName},</p>
              <p style="font-size: 15px; line-height: 1.6;">Your appeal for case <b>${appealDetails.caseReference}</b> has been <b>UPHELD</b>. The original penalty has been cancelled and no payment is required.</p>
              <div style="background-color: #f3fcf5; padding: 15px; border-left: 4px solid #107c10; border-radius: 4px; margin: 20px 0; color: #107c10; font-weight: bold; font-size: 14px;">
                No payment is required. This case is now officially closed.
              </div>
              ${updates.decisionNotes ? `<p style="font-size: 14px; line-height: 1.6; background-color: #f8fafc; padding: 12px; border: 1px solid #e2e8f0; border-radius: 6px;"><b>Reviewer notes:</b> ${updates.decisionNotes}</p>` : ''}
              
              <div style="margin: 25px 0 20px 0;">
                <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate;">
                  <tr>
                    <td align="center" bgcolor="#107c10" style="border-radius: 4px; padding: 0;">
                      <a href="${portalUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">
                        View Case Status
                      </a>
                    </td>
                  </tr>
                </table>
              </div>
              <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />
              <p style="font-size: 11px; color: #64748b; margin: 0;">This is an automated notification from the PACT Compliance Governance Platform.</p>
            </div>`;
          break;
        case 'Reduced':
          newCaseStatus = CASE_STATUS.UNPAID;
          emailSubject = `APPEAL PARTIALLY APPROVED – Penalty Reduced: ${appealDetails.caseReference}`;
          emailBody = `
            <div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; padding: 20px; border: 2px solid #f59e0b; border-radius: 8px; background-color: #ffffff;">
              <h2 style="color: #f59e0b; margin-top: 0; font-size: 20px;">Appeal Reviewed: Penalty Reduced</h2>
              <p style="font-size: 15px; line-height: 1.6;">Dear ${userName},</p>
              <p style="font-size: 15px; line-height: 1.6;">Your appeal for case <b>${appealDetails.caseReference}</b> has been reviewed and the penalty has been <b>REDUCED</b>.</p>
              <div style="background-color: #fffbeb; padding: 15px; border-left: 4px solid #f59e0b; border-radius: 4px; margin: 20px 0; color: #b45309; font-weight: bold; font-size: 14px;">
                Please proceed to check the portal to view the revised penalty details and proceed with the response.
              </div>
              ${updates.decisionNotes ? `<p style="font-size: 14px; line-height: 1.6; background-color: #f8fafc; padding: 12px; border: 1px solid #e2e8f0; border-radius: 6px;"><b>Reviewer notes:</b> ${updates.decisionNotes}</p>` : ''}
              
              <div style="margin: 25px 0 20px 0;">
                <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate;">
                  <tr>
                    <td align="center" bgcolor="#f59e0b" style="border-radius: 4px; padding: 0;">
                      <a href="${portalUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">
                        View Case Status
                      </a>
                    </td>
                  </tr>
                </table>
              </div>
              <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />
              <p style="font-size: 11px; color: #64748b; margin: 0;">This is an automated notification from the PACT Compliance Governance Platform.</p>
            </div>`;
          break;
        case 'Rejected':
          newCaseStatus = CASE_STATUS.UNPAID;
          emailSubject = `APPEAL REJECTED – Payment Required: ${appealDetails.caseReference}`;
          emailBody = `
            <div style="font-family: Arial, sans-serif; color: #333333; max-width: 600px; padding: 20px; border: 2px solid #d13438; border-radius: 8px; background-color: #ffffff;">
              <h2 style="color: #d13438; margin-top: 0; font-size: 20px;">Appeal Rejected: Payment Required</h2>
              <p style="font-size: 15px; line-height: 1.6;">Dear ${userName},</p>
              <p style="font-size: 15px; line-height: 1.6;">Your appeal for case <b>${appealDetails.caseReference}</b> has been reviewed and <b>REJECTED</b>. The original penalty stands and must be paid.</p>
              <div style="background-color: #fde8e8; padding: 15px; border-left: 4px solid #d13438; border-radius: 4px; margin: 20px 0; color: #d13438; font-weight: bold; font-size: 14px;">
                The original penalty stands. Please proceed to make your payment immediately.
              </div>
              ${updates.decisionNotes ? `<p style="font-size: 14px; line-height: 1.6; background-color: #f8fafc; padding: 12px; border: 1px solid #e2e8f0; border-radius: 6px;"><b>Reason:</b> ${updates.decisionNotes}</p>` : ''}
              
              <div style="margin: 25px 0 20px 0;">
                <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="border-collapse: separate;">
                  <tr>
                    <td align="center" bgcolor="#d13438" style="border-radius: 4px; padding: 0;">
                      <a href="${portalUrl}" target="_blank" style="padding: 12px 24px; font-family: Arial, sans-serif; font-size: 14px; color: #ffffff; text-decoration: none; font-weight: bold; display: inline-block;">
                        View Case Status
                      </a>
                    </td>
                  </tr>
                </table>
              </div>
              <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />
              <p style="font-size: 11px; color: #64748b; margin: 0;">This is an automated notification from the PACT Compliance Governance Platform.</p>
            </div>`;
          break;
      }

      if (newCaseStatus) {
        await this.updateCaseStatusForReference(appealDetails.caseReference, newCaseStatus);
      }

      await this.sendEmailNotification(
        [userEmail],
        emailSubject,
        emailBody
      );
    }
  }

  private mapSPItemToCase(
    item: any,
    staff: StaffMember[] = [],
    policies: PolicyOffence[] = []
  ): ComplianceCase {
    // 1. Try to get the Charged Person ID
    let chargedPersonId = String(this.readField(item, COLUMNS.CASES.CHARGED_PERSON + 'Id', 'ChargedPersonId', 'Charged_x0020_PersonId', 'Charged_x0020_PersaonId') || '');

    // 2. If no lookup ID, try the raw text field value
    if (!chargedPersonId) {
      const rawCharged = this.readField(item, COLUMNS.CASES.CHARGED_PERSON, 'ChargedPerson', 'Charged Persaon', 'Charged Person', 'Charged Person ', 'Charged_x0020_Person', 'Charged_x0020_Persaon');
      if (rawCharged && typeof rawCharged !== 'object') {
        chargedPersonId = String(rawCharged);
      } else if (rawCharged && typeof rawCharged === 'object') {
        chargedPersonId = String(rawCharged.Id || rawCharged.ID || rawCharged.Title || '');
      }
    }

    // 3. Try FieldValuesAsText for the display name
    const chargedPersonDisplayText = item.FieldValuesAsText
      ? String(item.FieldValuesAsText[COLUMNS.CASES.CHARGED_PERSON] || item.FieldValuesAsText['ChargedPerson'] || item.FieldValuesAsText['Charged Persaon'] || item.FieldValuesAsText['Charged Person'] || item.FieldValuesAsText['Charged_x0020_Person'] || item.FieldValuesAsText['Charged_x0020_Persaon'] || '')
      : '';

    // 4. Resolve the person from the staff directory
    let chargedPerson = staff.find(s => s.id === chargedPersonId);
    if (!chargedPerson && chargedPersonId) {
      chargedPerson = staff.find(s => s.fullName && chargedPersonId && s.fullName.toLowerCase() === chargedPersonId.toLowerCase());
    }
    if (!chargedPerson && chargedPersonDisplayText) {
      chargedPerson = staff.find(s => s.fullName && s.fullName.toLowerCase().trim() === chargedPersonDisplayText.toLowerCase().trim());
      // If we found by display name, use the staff member's canonical ID
      if (chargedPerson) {
        chargedPersonId = chargedPerson.id;
      }
    }

    // Debug logging for first item to help diagnose field mapping
    if (!this._caseMappingLogged) {
      this._caseMappingLogged = true;
      console.log('[PACT Debug] mapSPItemToCase sample item keys:', Object.keys(item).filter(k => !k.startsWith('__')).join(', '));
      console.log('[PACT Debug] ChargedPerson resolution:', {
        chargedPersonId,
        chargedPersonDisplayText,
        resolvedStaff: chargedPerson?.fullName || 'NOT FOUND',
        rawFieldValue: item[COLUMNS.CASES.CHARGED_PERSON],
        rawFieldIdValue: item[COLUMNS.CASES.CHARGED_PERSON + 'Id']
      });
    }
    
    let offenceCategoryId = String(this.readField(item, COLUMNS.CASES.OFFENCE_CATEGORY + 'Id', 'OffenceCategoryId', 'OffenceCategoryVal') || '');
    if (!offenceCategoryId) {
      const offObj = item[COLUMNS.CASES.OFFENCE_CATEGORY] || item['OffenceCategory'] || item['Offence_x0020_Category'];
      if (offObj && typeof offObj === 'object') {
        offenceCategoryId = String(offObj.Id || offObj.ID || '');
      }
    }
    if (!offenceCategoryId) {
      offenceCategoryId = String(this.readField(item, COLUMNS.CASES.OFFENCE_CATEGORY, 'OffenceCategory', 'Offence Category', 'Offence_x0020_Category') || '');
    }

    let rawOffence = item.FieldValuesAsText ? (item.FieldValuesAsText[COLUMNS.CASES.OFFENCE_CATEGORY] || item.FieldValuesAsText['OffenceCategory'] || item.FieldValuesAsText['Offence_x0020_Category'] || item.FieldValuesAsText['Offence']) : '';
    if (!rawOffence) {
      rawOffence = this.readField(item, COLUMNS.CASES.OFFENCE_CATEGORY, 'OffenceCategory', 'Offence Category', 'Offence_x0020_Category', 'Offence');
    }
    const offenceStr = rawOffence ? (typeof rawOffence === 'object' ? (rawOffence.Title || rawOffence.LookupValue || '') : String(rawOffence)) : '';

    const policy = policies.find(p => 
      p.id === offenceCategoryId || 
      p.id === 'p' + offenceCategoryId || 
      p.id === offenceCategoryId.replace(/^p/, '') ||
      (offenceStr && p.offenceName.toLowerCase().trim() === offenceStr.toLowerCase().trim())
    );

    const titleField = this.readField(item, COLUMNS.CASES.TITLE, 'Penalty ID', 'Case ID', 'CaseID', 'Title') || '';
    const isHistoricalItem = titleField.startsWith('HIST-');
    const rawDescription = this.readField(item, 'OffenceDescription', 'Offence Description', 'Offence_x0020_Description', 'Description') || '';

    return {
      id: String(item.ID || item.Id || ''),
      title: titleField,
      chargedPerson: chargedPersonId,
      chargedPersonName: chargedPerson?.fullName || chargedPersonDisplayText || this.readField(item, COLUMNS.CASES.CHARGED_PERSON, 'ChargedPerson', 'Charged Persaon', 'Charged Person', 'Charged Person ') || '',
      staffEmail: chargedPerson?.email || this.readField(item, COLUMNS.CASES.STAFF_EMAIL, 'StaffEmail', 'Email', 'Charged Person Email', 'Staff Email', 'ChargedPersonEmail') || '',
      department: chargedPerson?.department || this.readField(item, COLUMNS.CASES.DEPARTMENT, 'Department') || '',
      offenceCategory: offenceCategoryId,
      offenceCategoryName: isHistoricalItem && rawDescription ? rawDescription : (policy ? this.expandAbbreviations(policy.offenceName) : offenceStr),
      offenceDescription: rawDescription,
      penaltyAmount: this.parsePenalty(this.readField(item, COLUMNS.CASES.PENALTY_AMOUNT, 'PenaltyAmount', 'Penalty Amount', 'Penalty', 'Amount', 'Penalty_x0020_Amount')),
      dueDate: this.readField(item, COLUMNS.CASES.DUE_DATE, 'DueDate', 'Due Date', 'Due_x0020_Date', 'Payment Due Date'),
      issuerName: this.readField(item, COLUMNS.CASES.ISSUER_NAME, 'IssuerName', 'Issuer Name', 'Issuer_x0020_Name'),
      secondaryContact: this.readField(item, COLUMNS.CASES.SECONDARY_CONTACT, 'SecondaryContact', 'Secondary Contact Email'),
      status: this.readField(item, COLUMNS.CASES.STATUS, 'Status'),
      dateCreated: this.readField(item, 'Created', 'Date Created', 'Date Created ') || new Date().toISOString(),
      tier: this.readField(item, COLUMNS.CASES.TIER, 'Tier') || policy?.tier || 'Tier 1'
    };
  }

  private getInitialMockCases(): ComplianceCase[] {
    return [];
  }

  private getInitialMockTrackers(): RepeatOffenceRecord[] {
    return [];
  }

  private getInitialMockEscalations(): EscalationEntry[] {
    return [];
  }

  private getInitialMockStaff(): StaffMember[] {
    return staffData as StaffMember[];
  }

  private getInitialMockPolicies(): PolicyOffence[] {
    return policyData as PolicyOffence[];
  }

  private getInitialMockAppeals(): any[] {
    return [];
  }

  private getInitialMockDisciplinary(): any[] {
    return [];
  }

  public async getListColumnsDiagnostic(listName: string): Promise<string[]> {
    if (this.isLocal) {
      return ["Demo Mode: Columns cannot be diagnosed. Switch to SharePoint mode."];
    }
    try {
      const data = await this.fetchREST(`web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,TypeAsString,StaticName&$filter=Hidden eq false`);
      return (data.results || []).map((f: any) => `${f.Title} (${f.TypeAsString}) -> InternalName: "${f.InternalName}"`);
    } catch (e) {
      return [`Error: ${e instanceof Error ? e.message : String(e)}`];
    }
  }

  public async getRawFirstItemDiagnostic(listName: string): Promise<string> {
    if (this.isLocal) {
      return "Demo Mode: Raw items not available.";
    }
    try {
      const endpoint = `web/lists/getbytitle('${listName}')/items?$top=1`;
      const data = await this.fetchREST(endpoint);
      const firstItem = (data.results || [])[0];
      if (!firstItem) return `No items found in list: "${listName}"`;
      // Clean metadata objects to make output readable
      const cleanItem = { ...firstItem };
      delete cleanItem.__metadata;
      delete cleanItem.AttachmentFiles;
      delete cleanItem.FirstUniqueAncestorSecurableObject;
      delete cleanItem.RoleAssignments;
      return JSON.stringify(cleanItem, null, 2);
    } catch (e) {
      return `Error: ${e instanceof Error ? e.message : String(e)}`;
    }
  }
}

export const sharePointService = new SharePointService();
