/* Simulierte PIM-Schnittstelle (Microsoft Graph roleManagement/directory) für „Meine Rollen“ in der Vorschau.
   Jeder Aufruf steht in window.__coePimCalls. „Power Platform Administrator“ ist an den Vorschau-Schalter „Ohne Adminrolle“ gekoppelt. */
import { IMockResponse } from './odata';

interface IPimWindow {
  __coePimCalls?: { method: string; url: string; body: unknown }[];
  /** 'consent' = Berechtigung fehlt, 'forbidden' = Zugriff verweigert */
  __coePimDeny?: '' | 'consent' | 'forbidden';
  /** steuert die aktive Rolle im Vorschau-Token (Rollen-Abzeichen) */
  __coeRoleInactive?: boolean;
}
const w = window as unknown as IPimWindow;
const ME = '6d1c3b2a-0000-4000-8000-00000000c0e0';
const PP = '11648597-926c-4cf3-9c36-bcebb0ba8dcc';

interface IRoleSeed {
  id: string; name: string; description: string; eligible: boolean; viaGroup?: boolean; permanent?: boolean;
  maxHours: number; rules: string[]; approval?: boolean; authContext?: boolean;
}
const ROLES: IRoleSeed[] = [
  { id: PP, name: 'Power Platform Administrator', description: 'Kann alle Aspekte von Microsoft Dataverse, Power Apps, Power Automate und Copilot Studio verwalten.', eligible: true, maxHours: 8, rules: ['Justification', 'MultiFactorAuthentication'] },
  { id: 'f2ef992c-3afb-46b9-b7cf-a126ee74c451', name: 'Globaler Leser', description: 'Kann alles lesen, was ein globaler Administrator kann, aber nichts ändern.', eligible: true, maxHours: 4, rules: ['Justification', 'Ticketing'] },
  { id: '44367163-eba1-44c3-98af-f5787879f96a', name: 'Dynamics 365 Administrator', description: 'Kann alle Aspekte des Produkts Dynamics 365 verwalten.', eligible: true, maxHours: 8, rules: ['Justification'], approval: true },
  { id: '29232cdf-9323-42fd-ade2-1d097af3e4de', name: 'Exchange Administrator', description: 'Kann alle Aspekte des Produkts Exchange verwalten.', eligible: true, maxHours: 2, rules: ['Justification'], authContext: true },
  { id: 'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', name: 'SharePoint Administrator', description: 'Kann alle Aspekte des SharePoint-Diensts verwalten.', eligible: true, viaGroup: true, maxHours: 8, rules: ['Justification'] },
  { id: '790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b', name: 'Nachrichtencenter-Leser', description: 'Kann Nachrichten und Updates im Nachrichtencenter lesen.', eligible: false, permanent: true, maxHours: 0, rules: [] }
];

/** aktivierte Rollen: Rollen-ID → { seit, bis } */
const activated: { [id: string]: { from: number; to: number } } = {
  'f2ef992c-3afb-46b9-b7cf-a126ee74c451': { from: Date.now() - 30 * 60000, to: Date.now() + 90 * 60000 }
};
const pending: { [id: string]: true } = {};

function respond(status: number, body?: unknown): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return { ok: status >= 200 && status < 300, status, headers: { get: () => null }, json: () => Promise.resolve(text ? JSON.parse(text) : {}), text: () => Promise.resolve(text) };
}
const fail = (status: number, code: string, message: string): IMockResponse => respond(status, { error: { code, message } });
const def = (r: IRoleSeed): unknown => ({ id: r.id, displayName: r.name, description: r.description, templateId: r.id });

function isActive(r: IRoleSeed): { from: number; to: number } | undefined {
  if (r.id === PP) {
    if (w.__coeRoleInactive) {
      return undefined;
    }
    return activated[PP] || { from: Date.now() - 60 * 60000, to: Date.now() + 120 * 60000 };
  }
  const a = activated[r.id];
  return a && a.to > Date.now() ? a : undefined;
}

export function isPimUrl(url: string): boolean {
  return /\/v1\.0\/(roleManagement\/directory\/|policies\/roleManagementPolicyAssignments|me(\?|$))/.test(url);
}

export function mockPim(method: string, url: string, body?: string): Promise<IMockResponse> {
  return new Promise(resolve => setTimeout(() => {
    const parsed = body ? JSON.parse(body) as { [k: string]: unknown } : undefined;
    (w.__coePimCalls = w.__coePimCalls || []).push({ method, url, body: parsed });
    const path = decodeURIComponent(url.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, ''));
    if (/^\/me(\?|$)/.test(path)) {
      resolve(respond(200, { id: ME, displayName: 'Vorschau-Admin' }));
      return;
    }
    if (w.__coePimDeny === 'consent') {
      resolve(fail(403, 'Authorization_RequestDenied', "AADSTS65001: The user or administrator has not consented to use the application with ID '08e18876-6177-487e-b8b5-cf950c1e598c' named 'SharePoint Online Web Client Extensibility'."));
      return;
    }
    if (w.__coePimDeny === 'forbidden') {
      resolve(fail(403, 'Authorization_RequestDenied', 'Insufficient privileges to complete the operation.'));
      return;
    }
    if (method === 'GET' && /^\/roleManagement\/directory\/roleEligibilityScheduleInstances\/filterByCurrentUser/.test(path)) {
      resolve(respond(200, { value: ROLES.filter(r => r.eligible).map(r => ({
        id: `el-${r.id}`, principalId: ME, roleDefinitionId: r.id, directoryScopeId: '/', memberType: r.viaGroup ? 'Group' : 'Direct',
        startDateTime: new Date(Date.now() - 200 * 86400000).toISOString(), endDateTime: r.id === PP ? new Date(Date.now() + 160 * 86400000).toISOString() : null, roleDefinition: def(r)
      })) }));
      return;
    }
    if (method === 'GET' && /^\/roleManagement\/directory\/roleAssignmentScheduleInstances\/filterByCurrentUser/.test(path)) {
      const value: unknown[] = [];
      ROLES.forEach(r => {
        if (r.permanent) {
          value.push({ id: `as-${r.id}`, principalId: ME, roleDefinitionId: r.id, directoryScopeId: '/', memberType: 'Direct', assignmentType: 'Assigned', endDateTime: null, roleDefinition: def(r) });
          return;
        }
        const a = isActive(r);
        if (a) {
          value.push({ id: `as-${r.id}`, principalId: ME, roleDefinitionId: r.id, directoryScopeId: '/', memberType: 'Direct', assignmentType: 'Activated',
            startDateTime: new Date(a.from).toISOString(), endDateTime: new Date(a.to).toISOString(), roleDefinition: def(r) });
        }
      });
      resolve(respond(200, { value }));
      return;
    }
    if (method === 'GET' && /^\/roleManagement\/directory\/roleAssignmentScheduleRequests\/filterByCurrentUser/.test(path)) {
      resolve(respond(200, { value: Object.keys(pending).map(id => ({ id: `rq-${id}`, principalId: ME, roleDefinitionId: id, directoryScopeId: '/', status: 'PendingApproval', action: 'selfActivate' })) }));
      return;
    }
    if (method === 'GET' && /^\/policies\/roleManagementPolicyAssignments/.test(path)) {
      const id = (/roleDefinitionId eq '([^']+)'/.exec(path) || [])[1];
      const r = ROLES.filter(x => x.id === id)[0];
      resolve(respond(200, { value: r ? [{ id: `pol-${r.id}`, scopeId: '/', scopeType: 'DirectoryRole', roleDefinitionId: r.id, policy: { id: `p-${r.id}`, rules: [
        { '@odata.type': '#microsoft.graph.unifiedRoleManagementPolicyExpirationRule', id: 'Expiration_EndUser_Assignment', isExpirationRequired: true, maximumDuration: `PT${r.maxHours}H` },
        { '@odata.type': '#microsoft.graph.unifiedRoleManagementPolicyEnablementRule', id: 'Enablement_EndUser_Assignment', enabledRules: r.rules },
        { '@odata.type': '#microsoft.graph.unifiedRoleManagementPolicyAuthenticationContextRule', id: 'AuthenticationContext_EndUser_Assignment', isEnabled: !!r.authContext, claimValue: r.authContext ? 'c1' : '' },
        { '@odata.type': '#microsoft.graph.unifiedRoleManagementPolicyApprovalRule', id: 'Approval_EndUser_Assignment', setting: { isApprovalRequired: !!r.approval } }
      ] } }] : [] }));
      return;
    }
    if (method === 'POST' && /^\/roleManagement\/directory\/roleAssignmentScheduleRequests$/.test(path) && parsed) {
      const r = ROLES.filter(x => x.id === parsed.roleDefinitionId)[0];
      if (parsed.principalId !== ME) {
        resolve(fail(403, 'Authorization_RequestDenied', 'Only the signed-in user can be activated with selfActivate.'));
        return;
      }
      if (!r || !r.eligible || r.viaGroup) {
        resolve(fail(400, 'RoleAssignmentDoesNotExist', 'The role assignment does not exist.'));
        return;
      }
      if (parsed.action === 'selfDeactivate') {
        const a = isActive(r);
        if (!a) {
          resolve(fail(400, 'RoleAssignmentDoesNotExist', 'The role assignment does not exist.'));
          return;
        }
        if (Date.now() - a.from < 5 * 60000) {
          resolve(fail(400, 'ActiveDurationTooShort', 'Role assignment cannot be deactivated as it was activated less than 5 minutes ago.'));
          return;
        }
        delete activated[r.id];
        if (r.id === PP) {
          w.__coeRoleInactive = true;
        }
        resolve(respond(201, { id: `rq-${Date.now()}`, action: 'selfDeactivate', status: 'Revoked' }));
        return;
      }
      if (parsed.action !== 'selfActivate') {
        resolve(fail(403, 'Authorization_RequestDenied', 'Insufficient privileges to complete the operation.'));
        return;
      }
      if (r.authContext) {
        resolve(fail(400, 'RoleAssignmentRequestAcrsValidationFailed', 'The role assignment request requires an authentication context claim.'));
        return;
      }
      if (isActive(r)) {
        resolve(fail(400, 'RoleAssignmentExists', 'The Role assignment already exists.'));
        return;
      }
      if (pending[r.id]) {
        resolve(fail(400, 'RoleAssignmentRequestExists', 'A pending request already exists.'));
        return;
      }
      const sched = (parsed.scheduleInfo || {}) as { expiration?: { duration?: string } };
      const dur = /^PT(\d+)M$/.exec(String(sched.expiration && sched.expiration.duration || '')) ;
      const minutes = dur ? Number(dur[1]) : 0;
      const failed: string[] = [];
      if (!minutes || minutes > r.maxHours * 60) { failed.push('ExpirationRule'); }
      if (r.rules.indexOf('Justification') >= 0 && !String(parsed.justification || '').trim()) { failed.push('JustificationRule'); }
      const ticket = (parsed.ticketInfo || {}) as { ticketNumber?: string; ticketSystem?: string };
      if (r.rules.indexOf('Ticketing') >= 0 && (!ticket.ticketNumber || !ticket.ticketSystem)) { failed.push('TicketingRule'); }
      if (failed.length) {
        resolve(fail(400, 'RoleAssignmentRequestPolicyValidationFailed', `The following policy rules failed: ${JSON.stringify(failed)}`));
        return;
      }
      if (r.approval) {
        pending[r.id] = true;
        resolve(respond(201, { id: `rq-${Date.now()}`, action: 'selfActivate', status: 'PendingApproval' }));
        return;
      }
      activated[r.id] = { from: Date.now(), to: Date.now() + minutes * 60000 };
      if (r.id === PP) {
        w.__coeRoleInactive = false;
      }
      resolve(respond(201, { id: `rq-${Date.now()}`, action: 'selfActivate', status: 'Provisioned' }));
      return;
    }
    resolve(fail(404, 'ResourceNotFound', `Unsupported ${method} ${path}`));
  }, 180));
}
