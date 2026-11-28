import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * „Meine Rollen“ – eigene Entra-Rollen über Privileged Identity Management (Microsoft Graph, live).
 *
 * Bewusst nur SELBST-Aktionen der angemeldeten Person: „selfActivate“ und „selfDeactivate“ für Rollen, für die sie in PIM
 * berechtigt ist. Rollen für andere zuweisen kann man im Cockpit nicht. Microsoft prüft jede Anfrage gegen die PIM-Richtlinie
 * der Rolle (maximale Dauer, Begründung, Ticket, MFA, Genehmigung) und protokolliert sie im Entra-Audit-Log.
 *
 * Berechtigungen (delegiert, für „SharePoint Online Web Client Extensibility“, Skript -Pim):
 *   RoleEligibilitySchedule.Read.Directory  – berechtigte Rollen lesen
 *   RoleAssignmentSchedule.ReadWrite.Directory – aktive Rollen lesen, selbst aktivieren/deaktivieren
 *   RoleManagementPolicy.Read.Directory     – Regeln der Rolle lesen (Dauer, Begründung, Ticket …)
 */
export interface IRolePolicy {
  /** maximale Aktivierungsdauer in Minuten */
  maxMinutes: number;
  justification: boolean;
  ticket: boolean;
  mfa: boolean;
  approval: boolean;
  /** Conditional-Access-Authentifizierungskontext – nur im PIM-Portal möglich */
  authContext: boolean;
}

export interface IMyRole {
  key: string;
  roleDefinitionId: string;
  name: string;
  description: string;
  directoryScopeId: string;
  /** berechtigt (aktivierbar) */
  eligible: boolean;
  /** Berechtigung über eine Gruppe (dann im PIM-Portal über die Gruppe aktivieren) */
  viaGroup: boolean;
  eligibleUntil: string;
  active: boolean;
  /** dauerhaft zugewiesen (nicht über PIM aktiviert) */
  permanent: boolean;
  /** selbst über PIM aktiviert (direkt) – nur dann lässt sich die Rolle vorzeitig beenden */
  selfActivated: boolean;
  activeUntil: string;
  pending: boolean;
  policy?: IRolePolicy;
}

export interface IMyRoles {
  principalId: string;
  roles: IMyRole[];
  /** Regeln nicht lesbar (Berechtigung fehlt) */
  policyProblem?: string;
  loadedAt: Date;
}

export class PimError extends Error {
  constructor(message: string, public kind: 'consent' | 'forbidden' | 'rule' | 'failed', public detail?: string) {
    super(message);
    Object.setPrototypeOf(this, PimError.prototype);
  }
}

interface IGraphErr { statusCode?: number; code?: string; message?: string; body?: string }
interface IRoleDef { id?: string; displayName?: string; description?: string; templateId?: string }
interface IInstance {
  roleDefinitionId?: string; directoryScopeId?: string; memberType?: string; assignmentType?: string; endDateTime?: string | null;
  roleDefinition?: IRoleDef;
}

const CONSENT = 'Das Cockpit darf PIM noch nicht nutzen: Ein Administrator muss die Microsoft-Graph-Berechtigungen „RoleEligibilitySchedule.Read.Directory“, '
  + '„RoleAssignmentSchedule.ReadWrite.Directory“ und „RoleManagementPolicy.Read.Directory“ für SharePoint freigeben (scripts/Grant-CoeCockpitPermissions.ps1 -Pim).';

function errInfo(e: unknown): { status: number; text: string } {
  const err = (e || {}) as IGraphErr;
  let text = err.message || String(e);
  if (err.code && text.indexOf(err.code) < 0) {
    text = `${err.code}: ${text}`;
  }
  return { status: err.statusCode || 0, text };
}

/** Graph-Fehler in eine verständliche Meldung übersetzen. */
export function pimError(e: unknown): PimError {
  const { status, text } = errInfo(e);
  if (/AADSTS65001|consent|not been granted|has not consented/i.test(text)) {
    return new PimError(CONSENT, 'consent', text);
  }
  if (/Acrs|AuthenticationContext/i.test(text)) {
    return new PimError('Für diese Rolle verlangt PIM einen Authentifizierungskontext (Conditional Access). Das geht nur im PIM-Portal.', 'rule', text);
  }
  if (/MfaRule|multi-?factor|\bMFA\b/i.test(text)) {
    return new PimError('Für diese Rolle verlangt PIM eine Anmeldung mit MFA in dieser Sitzung. Bitte einmal ab- und mit MFA wieder anmelden oder im PIM-Portal aktivieren.', 'rule', text);
  }
  if (/ExpirationRule/i.test(text)) {
    return new PimError('Die gewählte Dauer ist länger, als die PIM-Richtlinie der Rolle erlaubt.', 'rule', text);
  }
  if (/JustificationRule/i.test(text)) {
    return new PimError('Die PIM-Richtlinie verlangt eine Begründung.', 'rule', text);
  }
  if (/TicketingRule/i.test(text)) {
    return new PimError('Die PIM-Richtlinie verlangt Ticketnummer und Ticketsystem.', 'rule', text);
  }
  if (/RoleAssignmentExists|already exists|already active/i.test(text)) {
    return new PimError('Die Rolle ist bereits aktiv.', 'rule', text);
  }
  if (/RoleAssignmentRequestExists|PendingRoleAssignmentRequest|pending request/i.test(text)) {
    return new PimError('Für diese Rolle läuft bereits eine Anfrage (z. B. Genehmigung ausstehend).', 'rule', text);
  }
  if (/ActiveDurationTooShort/i.test(text)) {
    return new PimError('Eine gerade aktivierte Rolle lässt sich erst nach 5 Minuten wieder deaktivieren.', 'rule', text);
  }
  if (/RoleAssignmentDoesNotExist|eligib.*not found|no eligible/i.test(text)) {
    return new PimError('Für diese Rolle besteht keine (passende) Berechtigung in PIM.', 'rule', text);
  }
  if (status === 401 || status === 403 || /Authorization_RequestDenied|Forbidden|AccessDenied/i.test(text)) {
    return new PimError(`Zugriff verweigert. ${CONSENT}`, 'forbidden', text);
  }
  return new PimError(`PIM: ${text}`, status === 400 ? 'rule' : 'failed', text);
}

/** ISO-8601-Dauer (PT8H, PT30M, P1D, PT1H30M) in Minuten. */
export function durationMinutes(iso: string | undefined): number {
  const m = /^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?)?$/i.exec(String(iso || ''));
  if (!m) {
    return 0;
  }
  return (Number(m[1] || 0) * 24 + Number(m[2] || 0)) * 60 + Number(m[3] || 0);
}

interface IRule {
  id?: string; maximumDuration?: string; isExpirationRequired?: boolean; enabledRules?: string[]; isEnabled?: boolean;
  setting?: { isApprovalRequired?: boolean };
}

function policyFrom(rules: IRule[]): IRolePolicy {
  const byId = (id: string): IRule | undefined => rules.filter(r => r.id === id)[0];
  const exp = byId('Expiration_EndUser_Assignment');
  const en = byId('Enablement_EndUser_Assignment');
  const ac = byId('AuthenticationContext_EndUser_Assignment');
  const ap = byId('Approval_EndUser_Assignment');
  const enabled = (en && en.enabledRules) || [];
  return {
    maxMinutes: durationMinutes(exp && exp.maximumDuration) || 8 * 60,
    justification: enabled.indexOf('Justification') >= 0,
    ticket: enabled.indexOf('Ticketing') >= 0,
    mfa: enabled.indexOf('MultiFactorAuthentication') >= 0,
    approval: !!(ap && ap.setting && ap.setting.isApprovalRequired),
    authContext: !!(ac && ac.isEnabled)
  };
}

async function client(factory: MSGraphClientFactory): Promise<MSGraphClientV3> {
  return factory.getClient('3');
}

async function getAll<T>(graph: MSGraphClientV3, path: string): Promise<T[]> {
  let out: T[] = [];
  let url: string | undefined = path;
  for (let i = 0; url && i < 20; i++) {
    const res = await graph.api(url).version('v1.0').get() as { value?: T[]; '@odata.nextLink'?: string };
    out = out.concat(res.value || []);
    url = res['@odata.nextLink'];
  }
  return out;
}

/** Eigene Rollen, für die die Person in PIM berechtigt ist: aktiv/nicht aktiv, ausstehend – mit Regeln je Rolle. */
export async function loadMyRoles(factory: MSGraphClientFactory): Promise<IMyRoles> {
  let graph: MSGraphClientV3;
  try {
    graph = await client(factory);
  } catch (e) {
    throw pimError(e);
  }
  const expand = '$expand=roleDefinition($select=id,displayName,description,templateId)';
  let me: { id?: string };
  let eligible: IInstance[];
  let active: IInstance[];
  try {
    [me, eligible, active] = await Promise.all([
      graph.api('/me').version('v1.0').select('id').get() as Promise<{ id?: string }>,
      getAll<IInstance>(graph, `/roleManagement/directory/roleEligibilityScheduleInstances/filterByCurrentUser(on='principal')?${expand}`),
      getAll<IInstance>(graph, `/roleManagement/directory/roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')?${expand}`)
    ]);
  } catch (e) {
    throw pimError(e);
  }
  const pending = await getAll<IInstance>(graph, "/roleManagement/directory/roleAssignmentScheduleRequests/filterByCurrentUser(on='principal')?$filter=status eq 'PendingApproval'")
    .catch(() => [] as IInstance[]);

  const map: { [key: string]: IMyRole } = {};
  const keyOf = (i: IInstance): string => `${String(i.roleDefinitionId || '').toLowerCase()}|${i.directoryScopeId || '/'}`;
  const base = (i: IInstance): IMyRole => {
    const def = i.roleDefinition || {};
    return {
      key: keyOf(i), roleDefinitionId: String(i.roleDefinitionId || def.id || ''), name: String(def.displayName || i.roleDefinitionId || ''),
      description: String(def.description || ''), directoryScopeId: String(i.directoryScopeId || '/'), eligible: false, viaGroup: false,
      eligibleUntil: '', active: false, permanent: false, selfActivated: false, activeUntil: '', pending: false
    };
  };
  eligible.forEach(i => {
    const r = map[keyOf(i)] = map[keyOf(i)] || base(i);
    const group = /group/i.test(String(i.memberType || ''));
    // direkte Berechtigung hat Vorrang vor einer über eine Gruppe
    if (!r.eligible || r.viaGroup) {
      r.eligible = true;
      r.viaGroup = group;
      r.eligibleUntil = String(i.endDateTime || '');
    }
  });
  active.forEach(i => {
    const r = map[keyOf(i)] = map[keyOf(i)] || base(i);
    r.active = true;
    const activated = /activated/i.test(String(i.assignmentType || ''));
    r.permanent = !activated && !i.endDateTime;
    r.selfActivated = r.selfActivated || (activated && !/group/i.test(String(i.memberType || '')));
    r.activeUntil = String(i.endDateTime || '');
  });
  pending.forEach(i => {
    const r = map[keyOf(i)];
    if (r) {
      r.pending = true;
    }
  });

  // Nur Rollen, für die die Person in PIM berechtigt ist (= die sie aktivieren kann) – dauerhafte Zuweisungen ohne PIM nicht
  const roles = Object.keys(map).map(k => map[k]).filter(r => r.eligible);
  // Regeln je Rolle (Rollen-Richtlinie)
  let policyProblem: string | undefined;
  await Promise.all(roles.filter(r => r.eligible && !r.viaGroup).map(async r => {
    try {
      const res = await graph.api(`/policies/roleManagementPolicyAssignments?$filter=scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '${r.roleDefinitionId}'`
        + '&$expand=policy($expand=rules)').version('v1.0').get() as { value?: { policy?: { rules?: IRule[] } }[] };
      const pol = (res.value || [])[0];
      if (pol && pol.policy) {
        r.policy = policyFrom(pol.policy.rules || []);
      }
    } catch (e) {
      const err = pimError(e);
      // Die Richtlinien lesen dürfen nur bestimmte Rollen (z. B. Globaler Leser) – das ist kein Zustimmungsproblem
      policyProblem = err.kind === 'forbidden' ? 'keine Leserechte auf die PIM-Richtlinien – Microsoft prüft die Regeln beim Aktivieren' : err.message;
    }
  }));
  roles.sort((a, b) => (a.eligible === b.eligible ? a.name.localeCompare(b.name) : a.eligible ? -1 : 1));
  return { principalId: String(me.id || ''), roles, policyProblem, loadedAt: new Date() };
}

export interface IActivateRequest {
  principalId: string;
  role: IMyRole;
  minutes: number;
  justification: string;
  ticketNumber?: string;
  ticketSystem?: string;
}

/** Eigene berechtigte Rolle aktivieren („selfActivate“). Liefert den Status der Anfrage (z. B. „Provisioned“, „PendingApproval“). */
export async function activateRole(factory: MSGraphClientFactory, req: IActivateRequest): Promise<string> {
  if (!req.role.eligible || req.role.viaGroup) {
    throw new PimError('Diese Rolle kann hier nicht aktiviert werden (keine direkte Berechtigung).', 'rule');
  }
  const body: { [k: string]: unknown } = {
    action: 'selfActivate',
    principalId: req.principalId,
    roleDefinitionId: req.role.roleDefinitionId,
    directoryScopeId: req.role.directoryScopeId || '/',
    justification: req.justification.trim(),
    scheduleInfo: {
      startDateTime: new Date().toISOString(),
      expiration: { type: 'AfterDuration', duration: `PT${Math.max(5, Math.round(req.minutes))}M` }
    }
  };
  if (req.ticketNumber || req.ticketSystem) {
    body.ticketInfo = { ticketNumber: (req.ticketNumber || '').trim(), ticketSystem: (req.ticketSystem || '').trim() };
  }
  try {
    const graph = await client(factory);
    const res = await graph.api('/roleManagement/directory/roleAssignmentScheduleRequests').version('v1.0').post(body) as { status?: string };
    return String((res && res.status) || 'Provisioned');
  } catch (e) {
    throw pimError(e);
  }
}

/** Eigene aktivierte Rolle vorzeitig beenden („selfDeactivate“). */
export async function deactivateRole(factory: MSGraphClientFactory, principalId: string, role: IMyRole): Promise<void> {
  try {
    const graph = await client(factory);
    await graph.api('/roleManagement/directory/roleAssignmentScheduleRequests').version('v1.0').post({
      action: 'selfDeactivate', principalId, roleDefinitionId: role.roleDefinitionId, directoryScopeId: role.directoryScopeId || '/'
    });
  } catch (e) {
    throw pimError(e);
  }
}
