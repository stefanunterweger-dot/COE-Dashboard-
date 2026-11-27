/* Simulierte DLP-Schnittstelle (api.bap.microsoft.com, PowerPlatform.Governance/v2) für die DLP-Auswirkungsanalyse.
   Richtlinien aus den CoE-Testdaten; jeder Aufruf steht in window.__coeBapCalls. */
import { IMockResponse } from './odata';
import { DLP_MOCK } from './data';

interface IRef { id: string; name: string; type: string }
interface IPolicy {
  name: string; displayName: string; defaultConnectorsClassification: string;
  connectorGroups: { classification: string; connectors: IRef[] }[];
  environmentType: string; environments: IRef[];
  createdBy: { displayName: string }; createdTime: string; lastModifiedBy: { displayName: string }; lastModifiedTime: string; etag: string; isLegacySchemaVersion: boolean;
}
interface IBapWindow {
  __coeAdminDeny?: '' | 'forbidden' | 'consent' | 'busy';
  __coeBapCalls?: { method: string; url: string; body: unknown }[];
  /** Name einer Richtlinie, die „jemand anderes“ vor dem nächsten Lesen ändert (Konflikt-Test) */
  __coeBapTouch?: string;
  /** Live-Liste schlägt fehl (z. B. ohne Rolle) – das Cockpit zeigt dann den CoE-Stand */
  __coeBapDown?: boolean;
}
const w = window as unknown as IBapWindow;
const DAY = 86400000;
const iso = (days: number): string => new Date(Date.now() - days * DAY).toISOString();
const ENV_TYPE = 'Microsoft.BusinessAppPlatform/scopes/environments';
const envRef = (name: string): IRef => ({ id: `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${name}`, name, type: ENV_TYPE });
const conRef = (name: string): IRef => ({ id: `/providers/Microsoft.PowerApps/apis/${name}`, name, type: 'Microsoft.PowerApps/apis' });

/** Connectors, die sich nicht blockieren lassen (Namen wie in den Testdaten). */
const UNBLOCKABLE = ['shared_approvals', 'shared_notifications', 'shared_sharepoint', 'shared_office365outlook', 'shared_microsoftteams', 'shared_office365users',
  'shared_microsoftdataverse', 'shared_onedriveforbusiness', 'shared_excelonlinebusiness', 'shared_planner', 'shared_powerbi', 'shared_microsoftforms'];

let store: IPolicy[] | undefined;
function policies(): IPolicy[] {
  if (store) {
    return store;
  }
  const envName: { [coe: string]: string } = {};
  DLP_MOCK.environments.forEach(e => { envName[e.coeId] = e.name; });
  const conName: { [coe: string]: string } = {};
  DLP_MOCK.connectors.forEach(c => { conName[c.coeId] = c.name; });
  const standard = DLP_MOCK.connectors.filter(c => !c.custom).map(c => c.name);
  const people = ['Anna Berger', 'Lukas Huber', 'Service Center of Excellence', 'Maria Gruber'];
  store = DLP_MOCK.policies.map((p, i) => {
    const business = ((p.__nav_admin_confidential_policy_connectors || []) as string[]).map(id => conName[id]).filter(n => !!n);
    if (i === 0) {
      business.push('shared_excelonlinebusiness', 'shared_office365users');
    }
    const blocked = ((p.__nav_admin_blocked_policy_connectors || []) as string[]).map(id => conName[id]).filter(n => !!n);
    const general = standard.filter(n => business.indexOf(n) < 0 && blocked.indexOf(n) < 0);
    const scope = p.admin_scope === 597910000 ? 'AllEnvironments' : p.admin_scope === 597910002 ? 'ExceptEnvironments' : 'OnlyEnvironments';
    return {
      name: `5a1d0000-0000-4000-8000-${String(i + 1).padStart(12, '0')}`,
      displayName: String(p.admin_name),
      defaultConnectorsClassification: 'General',
      connectorGroups: [
        { classification: 'Confidential', connectors: business.map(conRef) },
        { classification: 'General', connectors: general.map(conRef) },
        { classification: 'Blocked', connectors: blocked.map(conRef) }
      ],
      environmentType: scope,
      environments: scope === 'AllEnvironments' ? [] : ((p.__nav_admin_policy_environments || []) as string[]).map(id => envName[id]).filter(n => !!n).map(envRef),
      createdBy: { displayName: people[i % people.length] },
      createdTime: iso(400 - i * 60),
      lastModifiedBy: { displayName: people[(i + 1) % people.length] },
      lastModifiedTime: iso(40 - i * 7),
      etag: `etag-${i + 1}`,
      isLegacySchemaVersion: false
    };
  });
  return store;
}

function respond(status: number, body?: unknown): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: () => null },
    json: () => Promise.resolve(text ? JSON.parse(text) : undefined),
    text: () => Promise.resolve(text)
  };
}

const copy = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;

function validate(def: IPolicy | undefined): string {
  if (!def || !Array.isArray(def.connectorGroups) || def.connectorGroups.length !== 3 || !def.displayName || !def.environmentType) {
    return 'DLPConnectorDefinitionInvalid: The policy definition is incomplete.';
  }
  const blocked = (def.connectorGroups.filter(g => g.classification === 'Blocked')[0] || { connectors: [] }).connectors;
  const bad = blocked.filter(c => UNBLOCKABLE.indexOf(c.name) >= 0);
  if (bad.length) {
    return `DLPUnblockableConnectorBlocked: Connector(s) ${bad.map(c => c.name).join(', ')} cannot be blocked.`;
  }
  if (def.environmentType === 'OnlyEnvironments' && !(def.environments || []).length) {
    return 'DLPEnvironmentListInvalid: At least one environment is required.';
  }
  return '';
}

export function mockBapCall(method: string, url: string, body?: string): Promise<IMockResponse> {
  return new Promise(resolve => setTimeout(() => {
    const u = new URL(url);
    const parsed = body ? JSON.parse(body) as { policyDefinition?: IPolicy } : undefined;
    (w.__coeBapCalls = w.__coeBapCalls || []).push({ method, url, body: parsed });
    if (w.__coeAdminDeny === 'forbidden' || w.__coeBapDown) {
      resolve(respond(403, { error: { code: 'Forbidden', message: 'The caller does not have permission to perform this action (Power Platform Administrator role required).' } }));
      return;
    }
    const path = u.pathname;
    if (/\/PowerPlatform\.Governance\/v1\/connectors\/metadata\/unblockable$/i.test(path)) {
      resolve(respond(200, { value: UNBLOCKABLE.map(n => ({ id: `/providers/Microsoft.PowerApps/apis/${n}`, metadata: { unblockable: true } })) }));
      return;
    }
    const m = /\/PowerPlatform\.Governance\/v2\/policies(?:\/([^/]+))?$/i.exec(path);
    if (!m) {
      resolve(respond(404, { error: { code: 'NotFound', message: `Keine simulierte Route für ${path}` } }));
      return;
    }
    const list = policies();
    const name = m[1] ? decodeURIComponent(m[1]) : '';
    const idx = name ? list.map(p => p.name).indexOf(name) : -1;
    if (name && idx < 0) {
      resolve(respond(404, { error: { code: 'PolicyNotFound', message: `Policy ${name} not found.` } }));
      return;
    }
    if (method === 'GET' && !name) {
      // wie BAP: die Liste enthält die Connector-Gruppen nicht
      resolve(respond(200, { value: list.map(p => { const c = copy(p) as Partial<IPolicy>; delete c.connectorGroups; return { policyDefinition: c }; }) }));
      return;
    }
    if (method === 'GET') {
      if (w.__coeBapTouch === name) {
        w.__coeBapTouch = '';
        list[idx].lastModifiedTime = new Date().toISOString();
        list[idx].lastModifiedBy = { displayName: 'Jemand anderes' };
      }
      resolve(respond(200, { policyDefinition: copy(list[idx]) }));
      return;
    }
    if (method === 'POST' && !name) {
      const def = parsed && parsed.policyDefinition;
      const problem = validate(def);
      if (problem) {
        resolve(respond(400, { error: { code: problem.split(':')[0], message: problem } }));
        return;
      }
      const now = new Date().toISOString();
      const created: IPolicy = { ...copy(def!), name: `5a1d0000-0000-4000-8000-${Date.now().toString().slice(-12)}`, createdBy: { displayName: 'Vorschau-Benutzer' }, createdTime: now,
        lastModifiedBy: { displayName: 'Vorschau-Benutzer' }, lastModifiedTime: now, etag: `etag-${Date.now()}`, isLegacySchemaVersion: false };
      list.push(created);
      resolve(respond(201, { policyDefinition: copy(created) }));
      return;
    }
    if (method === 'PATCH' && name) {
      const def = parsed && parsed.policyDefinition;
      const problem = validate(def) || (def && def.name !== name ? 'DLPPolicyNameMismatch: The policy name does not match.' : '');
      if (problem) {
        resolve(respond(400, { error: { code: problem.split(':')[0], message: problem } }));
        return;
      }
      const now = new Date().toISOString();
      list[idx] = { ...list[idx], ...copy(def!), lastModifiedBy: { displayName: 'Vorschau-Benutzer' }, lastModifiedTime: now, etag: `etag-${Date.now()}` };
      resolve(respond(200, { policyDefinition: copy(list[idx]) }));
      return;
    }
    if (method === 'DELETE' && name) {
      list.splice(idx, 1);
      resolve(respond(200));
      return;
    }
    resolve(respond(405, { error: { code: 'MethodNotAllowed', message: `${method} ${path}` } }));
  }, 240));
}
