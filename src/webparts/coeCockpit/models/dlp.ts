import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { IRecord } from './ITableConfig';
import { DataverseService } from '../services/DataverseService';
import { adminRequest, POWERAPPS_RESOURCE } from '../services/DirectActionService';

/**
 * DLP-Auswirkungsanalyse – wie die CoE-App „Data Policy Impact Analysis“ (DLP Editor V2):
 *  - Richtlinien live über die Admin-Schnittstelle (api.bap.microsoft.com, PowerPlatform.Governance/v2 – dieselbe wie der
 *    Connector „Power Platform for Admins V2“ und die PowerShell-Befehle *-DlpPolicy), Token für „PowerApps Service“.
 *    Ohne aktive Admin-Rolle: Richtlinien aus dem CoE (Tabelle „DLP Policy“, Stand letzter Sync, nur lesen).
 *  - Auswirkung aus dem CoE-Inventar: „App Connections“ / „Flow Connections“ der Apps und Flows in den betroffenen Environments.
 *  - Entwürfe in der CoE-Tabelle „Data Policy Drafts“, Aufgabenliste in „DLP Impact Analysis“ (gemeinsam mit der CoE-App).
 */
export const DLP_KEY = 'dlp-impact';

const BAP_POLICIES = 'https://api.bap.microsoft.com/providers/PowerPlatform.Governance/v2/policies';
const BAP_UNBLOCKABLE = 'https://api.bap.microsoft.com/providers/PowerPlatform.Governance/v1/connectors/metadata/unblockable';
const API_VERSION = 'api-version=2016-11-01';
export const PPAC_DLP_URL = 'https://admin.powerplatform.microsoft.com/security/dataprotection/policies';
export function ppacPolicyUrl(name: string): string {
  return name ? `https://admin.powerplatform.microsoft.com/dlp/policy/${encodeURIComponent(name)}` : PPAC_DLP_URL;
}

export type DlpGroup = 'Confidential' | 'General' | 'Blocked';
export const DLP_GROUPS: DlpGroup[] = ['Confidential', 'General', 'Blocked'];
export const GROUP_LABEL: { [g in DlpGroup]: string } = { Confidential: 'Business', General: 'Non-Business', Blocked: 'Blockiert' };
export type DlpScope = 'AllEnvironments' | 'OnlyEnvironments' | 'ExceptEnvironments' | 'SingleEnvironment';
export const SCOPE_LABEL: { [s in DlpScope]: string } = {
  AllEnvironments: 'Alle Environments',
  ExceptEnvironments: 'Alle Environments außer den ausgewählten',
  OnlyEnvironments: 'Nur die ausgewählten Environments',
  SingleEnvironment: 'Ein Environment (Environment-Richtlinie)'
};

export interface IDlpRef { id: string; name: string; type: string }
export interface IDlpPolicy {
  name: string;
  displayName: string;
  defaultConnectorsClassification: DlpGroup;
  connectorGroups?: { classification: DlpGroup; connectors: IDlpRef[] }[];
  environmentType: DlpScope;
  environments: IDlpRef[];
  createdBy?: { displayName?: string };
  createdTime?: string;
  lastModifiedBy?: { displayName?: string };
  lastModifiedTime?: string;
  etag?: string;
  [k: string]: unknown;
}

/** Richtlinie in der Liste (live, aus dem CoE oder Entwurf). */
export interface IDlpListItem {
  key: string;
  source: 'live' | 'coe' | 'draft';
  /** ID der Richtlinie (live/CoE) bzw. des Entwurfs */
  id: string;
  displayName: string;
  /** Entwurf: Name des Entwurfs */
  draftName?: string;
  scope: DlpScope;
  envs: string[];
  defaultGroup?: DlpGroup;
  /** Connectors je Gruppe (falls bekannt) */
  groups?: { [connectorId: string]: DlpGroup };
  createdBy: string;
  created: string;
  modifiedBy: string;
  modified: string;
  policy?: IDlpPolicy;
}

/** Arbeitskopie im Editor. */
export interface IDlpWorking {
  mode: 'edit' | 'copy' | 'draft' | 'new';
  /** Live-Richtlinie, die geändert wird (mode edit) */
  policyName?: string;
  /** Entwurf in der CoE-Tabelle (mode draft) */
  draftId?: string;
  draftName: string;
  displayName: string;
  defaultGroup: DlpGroup;
  /** ausdrücklich zugeordnete Connectors (ID → Gruppe); alle anderen fallen in die Standardgruppe */
  groups: { [connectorId: string]: DlpGroup };
  /** Name/Typ je Connector-ID (für das Speichern) */
  refs: { [connectorId: string]: IDlpRef };
  scope: DlpScope;
  envs: string[];
  /** Ausgangszustand (für Vergleich und Auswirkung „neu betroffen“) */
  base?: { defaultGroup: DlpGroup; groups: { [id: string]: DlpGroup }; scope: DlpScope; envs: string[] };
  /** vollständige Live-Richtlinie (wird beim Speichern zurückgeschrieben) */
  policy?: IDlpPolicy;
}

export interface IDlpConnector {
  id: string;
  name: string;
  displayName: string;
  tier: string;
  publisher: string;
  apps: number;
  flows: number;
}

export interface IDlpEnv { name: string; coeId: string; displayName: string; type: string }

export interface IDlpBase {
  connectors: IDlpConnector[];
  envs: IDlpEnv[];
  envProblem?: string;
  connectorProblem?: string;
}

export interface IDlpResource {
  kind: 'app' | 'flow';
  /** Power-Platform-ID */
  id: string;
  /** Primärschlüssel im CoE */
  coeId: string;
  name: string;
  env: string;
  envCoeId: string;
  envName: string;
  owner: string;
  ownerId: string;
  modified: string;
  connectors: string[];
}

export interface IDlpImpact {
  resource: IDlpResource;
  business: string[];
  nonBusiness: string[];
  blocked: string[];
  /** schon mit der bisherigen Richtlinie betroffen */
  already: boolean;
}

export interface IDlpImpactResult {
  apps: IDlpImpact[];
  flows: IDlpImpact[];
  analyzedApps: number;
  analyzedFlows: number;
  envCount: number;
  truncated: boolean;
}

export interface IDlpTask {
  id: string;
  /** CoE-ID der App bzw. des Flows */
  resourceCoeId: string;
  name: string;
  kind: 'App' | 'Flow' | '';
  policyName: string;
  policyId: string;
  env: string;
  owner: string;
  business: string;
  nonBusiness: string;
  blocked: string;
  decision: string;
  status: string;
  actionBy: string;
  created: string;
}

export const connectorId = (name: string): string => `/providers/Microsoft.PowerApps/apis/${name}`;
const shortName = (id: string): string => String(id || '').split('/').pop() || '';
export const envRef = (name: string): IDlpRef => ({
  id: `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${name}`,
  name,
  type: 'Microsoft.BusinessAppPlatform/scopes/environments'
});
const norm = (env: string): string => String(env || '').toLowerCase().replace(/^default-/, '');

/** Connectors, die sich laut Microsoft nicht blockieren lassen (falls die Schnittstelle die Liste nicht liefert). */
const UNBLOCKABLE_FALLBACK = ['shared_approvals', 'shared_flowpush', 'shared_commondataserviceforapps', 'shared_commondataservice', 'shared_kaizala',
  'shared_office365groups', 'shared_office365', 'shared_office365users', 'shared_onedriveforbusiness', 'shared_onenote', 'shared_planner',
  'shared_powerappsnotification', 'shared_powerbi', 'shared_sharepointonline', 'shared_shifts', 'shared_teams', 'shared_todo',
  'shared_excelonlinebusiness', 'shared_yammer', 'shared_microsoftforms'];

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));

// ---------------------------------------------------------------------------------------------------------------
// Live-Schnittstelle
// ---------------------------------------------------------------------------------------------------------------

export class DlpApi {
  private client: Promise<AadHttpClient> | undefined;
  constructor(private factory: AadHttpClientFactory) {}

  private getClient = (): Promise<AadHttpClient> => {
    if (!this.client) {
      const p = this.factory.getClient(POWERAPPS_RESOURCE);
      this.client = p;
      p.catch(() => { this.client = undefined; });
    }
    return this.client;
  };

  private async req<T>(method: 'GET' | 'POST' | 'PATCH' | 'DELETE', url: string, body?: unknown): Promise<T> {
    const text = await adminRequest(this.getClient, POWERAPPS_RESOURCE, method, url, body);
    return (text ? JSON.parse(text) : {}) as T;
  }

  public async list(): Promise<IDlpPolicy[]> {
    const res = await this.req<{ value?: ({ policyDefinition?: IDlpPolicy } & IDlpPolicy)[] }>('GET', `${BAP_POLICIES}?${API_VERSION}`);
    return (res.value || []).map(v => (v.policyDefinition || v) as IDlpPolicy).filter(p => !!p && !!p.name);
  }

  public async get(name: string): Promise<IDlpPolicy> {
    const res = await this.req<{ policyDefinition?: IDlpPolicy } & IDlpPolicy>('GET', `${BAP_POLICIES}/${encodeURIComponent(name)}?${API_VERSION}`);
    return (res.policyDefinition || res) as IDlpPolicy;
  }

  public async create(def: IDlpPolicy): Promise<IDlpPolicy> {
    const res = await this.req<{ policyDefinition?: IDlpPolicy }>('POST', `${BAP_POLICIES}?${API_VERSION}`, { policyDefinition: def });
    return (res.policyDefinition || res) as IDlpPolicy;
  }

  /** Vollständige Richtlinie zurückschreiben (Microsoft verlangt die komplette Definition, nicht nur die Änderungen). */
  public async update(def: IDlpPolicy): Promise<IDlpPolicy> {
    const res = await this.req<{ policyDefinition?: IDlpPolicy }>('PATCH', `${BAP_POLICIES}/${encodeURIComponent(def.name)}?${API_VERSION}`, { policyDefinition: def });
    return (res.policyDefinition || res) as IDlpPolicy;
  }

  public async remove(name: string): Promise<void> {
    await this.req('DELETE', `${BAP_POLICIES}/${encodeURIComponent(name)}?${API_VERSION}`);
  }

  /** Nicht blockierbare Connectors (Namen wie „shared_sharepointonline“). */
  public async unblockable(): Promise<{ names: string[]; live: boolean }> {
    try {
      const res = await this.req<{ value?: { id?: string; name?: string }[] } | { id?: string; name?: string }[]>('GET', `${BAP_UNBLOCKABLE}?${API_VERSION}`);
      const list = Array.isArray(res) ? res : (res.value || []);
      const names = list.map(x => shortName(String(x.id || x.name || ''))).filter(n => !!n);
      if (names.length) {
        return { names, live: true };
      }
    } catch {
      // Route nicht überall verfügbar – Liste aus der Microsoft-Dokumentation
    }
    return { names: UNBLOCKABLE_FALLBACK, live: false };
  }
}

// ---------------------------------------------------------------------------------------------------------------
// Umwandlungen
// ---------------------------------------------------------------------------------------------------------------

function groupsOf(p: { connectorGroups?: { classification: DlpGroup; connectors: IDlpRef[] }[] }): { map: { [id: string]: DlpGroup }; refs: { [id: string]: IDlpRef } } {
  const map: { [id: string]: DlpGroup } = {};
  const refs: { [id: string]: IDlpRef } = {};
  (p.connectorGroups || []).forEach(g => (g.connectors || []).forEach(c => {
    const id = String(c.id || '').toLowerCase();
    if (id) {
      map[id] = g.classification;
      refs[id] = { id: c.id, name: c.name || shortName(c.id), type: c.type || 'Microsoft.PowerApps/apis' };
    }
  }));
  return { map, refs };
}

const envNames = (envs: IDlpRef[] | undefined): string[] => (envs || []).map(e => String(e.name || shortName(e.id))).filter(n => !!n);

export function itemFromLive(p: IDlpPolicy): IDlpListItem {
  return {
    key: `live:${p.name}`,
    source: 'live',
    id: p.name,
    displayName: p.displayName || p.name,
    scope: p.environmentType || 'AllEnvironments',
    envs: envNames(p.environments),
    defaultGroup: p.defaultConnectorsClassification,
    groups: p.connectorGroups ? groupsOf(p).map : undefined,
    createdBy: (p.createdBy && p.createdBy.displayName) || '',
    created: p.createdTime || '',
    modifiedBy: (p.lastModifiedBy && p.lastModifiedBy.displayName) || '',
    modified: p.lastModifiedTime || '',
    policy: p
  };
}

export function workingFrom(item: IDlpListItem, mode: IDlpWorking['mode']): IDlpWorking {
  const g = item.policy ? groupsOf(item.policy) : { map: { ...(item.groups || {}) }, refs: {} as { [id: string]: IDlpRef } };
  const defaultGroup = item.defaultGroup || 'General';
  const copy = mode === 'copy';
  return {
    mode,
    policyName: mode === 'edit' ? item.id : undefined,
    draftId: mode === 'draft' ? item.id : undefined,
    draftName: mode === 'draft' ? (item.draftName || item.displayName) : copy ? `Kopie von ${item.displayName}` : item.displayName,
    displayName: copy ? `Kopie von ${item.displayName}` : item.displayName,
    defaultGroup,
    groups: { ...g.map },
    refs: g.refs,
    scope: item.scope === 'SingleEnvironment' && copy ? 'OnlyEnvironments' : item.scope,
    envs: item.envs.slice(),
    base: mode === 'edit' ? { defaultGroup, groups: { ...g.map }, scope: item.scope, envs: item.envs.slice() } : undefined,
    policy: mode === 'edit' ? item.policy : undefined
  };
}

export function emptyWorking(): IDlpWorking {
  return { mode: 'new', draftName: 'Neue Richtlinie', displayName: 'Neue Richtlinie', defaultGroup: 'General', groups: {}, refs: {}, scope: 'OnlyEnvironments', envs: [] };
}

export function groupOf(w: { defaultGroup: DlpGroup; groups: { [id: string]: DlpGroup } }, id: string): DlpGroup {
  return w.groups[id.toLowerCase()] || w.defaultGroup;
}

/** Vollständige Definition zum Speichern (bei „ändern“ auf Basis der Live-Richtlinie, alle anderen Eigenschaften bleiben). */
export function toDefinition(w: IDlpWorking, connectors: IDlpConnector[], forCreate: boolean): IDlpPolicy {
  const byId: { [id: string]: IDlpRef } = {};
  connectors.forEach(c => { byId[c.id.toLowerCase()] = { id: c.id, name: c.name, type: 'Microsoft.PowerApps/apis' }; });
  const groups: { classification: DlpGroup; connectors: IDlpRef[] }[] = DLP_GROUPS.map(g => ({ classification: g, connectors: [] as IDlpRef[] }));
  Object.keys(w.groups).sort().forEach(id => {
    const ref = w.refs[id] || byId[id] || { id, name: shortName(id), type: 'Microsoft.PowerApps/apis' };
    groups[DLP_GROUPS.indexOf(w.groups[id])].connectors.push(ref);
  });
  const base: IDlpPolicy = !forCreate && w.policy ? { ...w.policy } : {} as IDlpPolicy;
  const def: IDlpPolicy = {
    ...base,
    displayName: w.displayName.trim(),
    defaultConnectorsClassification: w.defaultGroup,
    connectorGroups: groups,
    environmentType: w.scope,
    environments: w.scope === 'AllEnvironments' ? [] : w.envs.map(envRef)
  };
  if (forCreate) {
    delete (def as { name?: string }).name;
    delete def.etag;
    delete def.createdBy;
    delete def.createdTime;
    delete def.lastModifiedBy;
    delete def.lastModifiedTime;
  } else if (w.policyName) {
    def.name = w.policyName;
  }
  return def;
}

/** Gilt die Richtlinie für dieses Environment? */
export function appliesTo(scope: DlpScope, envs: string[], env: string): boolean {
  const n = norm(env);
  const listed = envs.some(e => norm(e) === n);
  return scope === 'AllEnvironments' ? true : scope === 'ExceptEnvironments' ? !listed : listed;
}

export function scopeText(item: { scope: DlpScope; envs: string[] }, envName: (name: string) => string): string {
  if (item.scope === 'AllEnvironments') {
    return 'Alle Environments';
  }
  if (item.scope === 'ExceptEnvironments') {
    return item.envs.length ? `Alle außer ${item.envs.length === 1 ? envName(item.envs[0]) : `${item.envs.length} Environments`}` : 'Alle Environments';
  }
  if (!item.envs.length) {
    return 'Kein Environment';
  }
  return item.envs.length === 1 ? envName(item.envs[0]) : `${item.envs.length} Environments`;
}

// ---------------------------------------------------------------------------------------------------------------
// CoE: Connectors, Environments, Richtlinien (Fallback), Entwürfe, Aufgabenliste, Inventar
// ---------------------------------------------------------------------------------------------------------------

const fmt = (r: IRecord, f: string): string => String(r[`${f}@OData.Community.Display.V1.FormattedValue`] || r[f] || '');

let baseCache: { service: DataverseService; at: number; p: Promise<IDlpBase> } | undefined;
let invCache: { service: DataverseService; at: number; max: number; p: Promise<{ apps: IDlpResource[]; flows: IDlpResource[]; truncated: boolean }> } | undefined;
export function resetDlpCache(): void {
  baseCache = undefined;
  invCache = undefined;
}

export function loadDlpBase(service: DataverseService, fresh: boolean = false): Promise<IDlpBase> {
  if (!fresh && baseCache && baseCache.service === service && Date.now() - baseCache.at < 10 * 60000) {
    return baseCache.p;
  }
  const p = (async (): Promise<IDlpBase> => {
    const out: IDlpBase = { connectors: [], envs: [] };
    await Promise.all([
      service.query('admin_connectors', ['admin_name', 'admin_displayname', 'admin_tier', 'admin_publisher', 'admin_iscustomapi', 'admin_numapps', 'admin_numflows'],
        'statecode eq 0 and admin_connectordeleted ne true', 20000).then(rows => {
        const seen: { [n: string]: boolean } = {};
        rows.forEach(r => {
          const name = String(r.admin_name || '');
          // Custom Connectors gehören je Environment; die regelt Microsoft über URL-Muster (im Admin Center)
          if (!name || r.admin_iscustomapi === true || seen[name.toLowerCase()]) {
            return;
          }
          seen[name.toLowerCase()] = true;
          out.connectors.push({
            id: connectorId(name), name, displayName: String(r.admin_displayname || name), tier: fmt(r, 'admin_tier'), publisher: String(r.admin_publisher || ''),
            apps: Number(r.admin_numapps) || 0, flows: Number(r.admin_numflows) || 0
          });
        });
        out.connectors.sort((a, b) => a.displayName.localeCompare(b.displayName));
      }, e => { out.connectorProblem = `Connectors aus dem CoE nicht lesbar: ${errText(e)}`; }),
      service.query('admin_environments', ['admin_environmentid', 'admin_environmentname', 'admin_displayname', 'admin_environmentsku', 'admin_environmentdeleted'],
        'statecode eq 0', 20000).then(rows => {
        out.envs = rows.filter(r => r.admin_environmentdeleted !== true && !!r.admin_environmentname).map(r => ({
          name: String(r.admin_environmentname), coeId: String(r.admin_environmentid || ''), displayName: String(r.admin_displayname || r.admin_environmentname), type: fmt(r, 'admin_environmentsku')
        })).sort((a, b) => a.displayName.localeCompare(b.displayName));
      }, e => { out.envProblem = `Environments aus dem CoE nicht lesbar: ${errText(e)}`; })
    ]);
    return out;
  })();
  baseCache = { service, at: Date.now(), p };
  p.catch(() => { baseCache = undefined; });
  return p;
}

/** Richtlinien aus dem CoE (Tabelle „DLP Policy“) – wenn die Live-Schnittstelle nicht geht. Standardgruppe ist dort nicht erfasst. */
export async function loadCoePolicies(service: DataverseService, envs: IDlpEnv[]): Promise<IDlpListItem[]> {
  const rows = await service.query('admin_dlppolicies', ['admin_dlppolicyid', 'admin_name', 'admin_scope', 'admin_policyid', 'createdon', 'modifiedon'], 'statecode eq 0', 2000);
  const envByCoe: { [id: string]: string } = {};
  envs.forEach(e => { envByCoe[e.coeId.toLowerCase()] = e.name; });
  const out = await Promise.all(rows.map(async (r): Promise<IDlpListItem> => {
    const id = String(r.admin_dlppolicyid || '');
    const scopeLabel = fmt(r, 'admin_scope');
    const scope: DlpScope = /exclud|except|außer/i.test(scopeLabel) ? 'ExceptEnvironments' : /includ|only|nur/i.test(scopeLabel) ? 'OnlyEnvironments' : 'AllEnvironments';
    const related = async (nav: string, idField: string, extra: string[]): Promise<IRecord[]> => {
      try {
        return await service.query(`admin_dlppolicies(${id})/${nav}`, [idField].concat(extra), undefined, 5000);
      } catch {
        return [];
      }
    };
    const [envRows, blocked, business] = await Promise.all([
      scope === 'AllEnvironments' ? Promise.resolve([] as IRecord[]) : related('admin_policy_environments', 'admin_environmentid', ['admin_environmentname']),
      related('admin_blocked_policy_connectors', 'admin_connectorid', ['admin_name']),
      related('admin_confidential_policy_connectors', 'admin_connectorid', ['admin_name'])
    ]);
    const groups: { [cid: string]: DlpGroup } = {};
    business.forEach(c => { if (c.admin_name) { groups[connectorId(String(c.admin_name)).toLowerCase()] = 'Confidential'; } });
    blocked.forEach(c => { if (c.admin_name) { groups[connectorId(String(c.admin_name)).toLowerCase()] = 'Blocked'; } });
    // Die Standardgruppe erfasst das CoE nicht – alle übrigen Connectors gelten hier als Non-Business
    return {
      key: `coe:${id}`, source: 'coe', id: String(r.admin_policyid || id), displayName: String(r.admin_name || ''), scope,
      envs: envRows.map(e => String(e.admin_environmentname || envByCoe[String(e.admin_environmentid || '').toLowerCase()] || '')).filter(n => !!n),
      defaultGroup: 'General', groups, createdBy: '', created: String(r.createdon || ''), modifiedBy: '', modified: String(r.modifiedon || '')
    };
  }));
  return out.sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export async function loadDrafts(service: DataverseService): Promise<IDlpListItem[]> {
  const rows = await service.query('admin_datapolicydraftses', ['admin_datapolicydraftsid', 'admin_name', 'admin_displayname', 'admin_scope', 'admin_defaultconnectorclassification',
    'admin_environments', 'admin_connectorgroups', '_createdby_value', 'createdon', '_modifiedby_value', 'modifiedon'], 'statecode eq 0', 2000);
  return rows.map(r => {
    let groups: { [id: string]: DlpGroup } = {};
    let envs: string[] = [];
    try {
      groups = groupsOf({ connectorGroups: JSON.parse(String(r.admin_connectorgroups || '[]')) as { classification: DlpGroup; connectors: IDlpRef[] }[] }).map;
    } catch {
      groups = {};
    }
    try {
      envs = envNames(JSON.parse(String(r.admin_environments || '[]')) as IDlpRef[]);
    } catch {
      envs = [];
    }
    const scope = (String(r.admin_scope || 'OnlyEnvironments') as DlpScope);
    return {
      key: `draft:${String(r.admin_datapolicydraftsid)}`, source: 'draft' as const, id: String(r.admin_datapolicydraftsid), draftName: String(r.admin_name || ''),
      displayName: String(r.admin_displayname || r.admin_name || ''), scope: SCOPE_LABEL[scope] ? scope : 'OnlyEnvironments', envs,
      defaultGroup: (DLP_GROUPS.indexOf(r.admin_defaultconnectorclassification as DlpGroup) >= 0 ? r.admin_defaultconnectorclassification : 'General') as DlpGroup,
      groups, createdBy: fmt(r, '_createdby_value'), created: String(r.createdon || ''), modifiedBy: fmt(r, '_modifiedby_value'), modified: String(r.modifiedon || '')
    };
  }).sort((a, b) => (b.modified || '').localeCompare(a.modified || ''));
}

/** Entwurf speichern (neu oder bestehend) – Format wie die CoE-App. */
export async function saveDraft(service: DataverseService, w: IDlpWorking, connectors: IDlpConnector[]): Promise<string> {
  const def = toDefinition(w, connectors, true);
  const values = {
    admin_name: (w.draftName.trim() || w.displayName.trim()).slice(0, 100),
    admin_apiversion: '2020-10-01',
    admin_displayname: w.displayName.trim().slice(0, 100),
    admin_defaultconnectorclassification: w.defaultGroup,
    admin_scope: w.scope,
    admin_environments: JSON.stringify((def.environments || []).map(e => ({ name: e.name, type: e.type, id: e.id }))),
    admin_connectorgroups: JSON.stringify(def.connectorGroups)
  };
  if (w.draftId) {
    await service.update('admin_datapolicydraftses', w.draftId, values);
    return w.draftId;
  }
  return service.create('admin_datapolicydraftses', values);
}

export async function loadTasks(service: DataverseService): Promise<IDlpTask[]> {
  const rows = await service.query('admin_dlpimpactanalysises', ['admin_dlpimpactanalysisid', 'admin_name', 'admin_dlppolicyname', 'admin_datapolicyid', '_admin_impactedapp_value',
    '_admin_impactedflow_value', '_admin_impactedenvironment_value', '_admin_impactowner_value', 'admin_conflictingconnectorbusiness', 'admin_conflictingconnectornonbusiness',
    'admin_conflictingconnectorblocked', 'admin_decision', 'admin_decisionstatus', 'admin_actionrequiredon', 'createdon'], 'statecode eq 0', 5000);
  return rows.map(r => ({
    id: String(r.admin_dlpimpactanalysisid || ''),
    resourceCoeId: String(r._admin_impactedapp_value || r._admin_impactedflow_value || '').toLowerCase(),
    name: String(r.admin_name || fmt(r, '_admin_impactedapp_value') || fmt(r, '_admin_impactedflow_value') || ''),
    kind: (r._admin_impactedapp_value ? 'App' : r._admin_impactedflow_value ? 'Flow' : '') as IDlpTask['kind'],
    policyName: String(r.admin_dlppolicyname || ''),
    policyId: String(r.admin_datapolicyid || ''),
    env: fmt(r, '_admin_impactedenvironment_value'),
    owner: fmt(r, '_admin_impactowner_value'),
    business: String(r.admin_conflictingconnectorbusiness || ''),
    nonBusiness: String(r.admin_conflictingconnectornonbusiness || ''),
    blocked: String(r.admin_conflictingconnectorblocked || ''),
    decision: fmt(r, 'admin_decision'),
    status: fmt(r, 'admin_decisionstatus'),
    actionBy: String(r.admin_actionrequiredon || ''),
    created: String(r.createdon || '')
  })).sort((a, b) => (b.created || '').localeCompare(a.created || ''));
}

/** Betroffene in die CoE-Aufgabenliste „DLP Impact Analysis“ eintragen (wie der CoE-Flow „DLP Editor | Parse impacted resources“). */
export async function addTask(service: DataverseService, i: IDlpImpact, policyName: string, policyId: string, label: (n: string) => string): Promise<void> {
  const r = i.resource;
  const values: { [k: string]: unknown } = {
    admin_name: r.name.slice(0, 100),
    admin_dlppolicyname: policyName.slice(0, 100),
    admin_datapolicyid: policyId.slice(0, 100),
    admin_conflictingconnectorbusiness: i.business.map(label).join(', '),
    admin_conflictingconnectornonbusiness: i.nonBusiness.map(label).join(', '),
    admin_conflictingconnectorblocked: i.blocked.map(label).join(', ')
  };
  values[r.kind === 'app' ? 'admin_ImpactedApp@odata.bind' : 'admin_ImpactedFlow@odata.bind'] = `/${r.kind === 'app' ? 'admin_apps' : 'admin_flows'}(${r.coeId})`;
  if (r.envCoeId) {
    values['admin_ImpactedEnvironment@odata.bind'] = `/admin_environments(${r.envCoeId})`;
  }
  if (r.ownerId) {
    values['admin_ImpactOwner@odata.bind'] = `/admin_makers(${r.ownerId})`;
  }
  try {
    await service.create('admin_dlpimpactanalysises', values);
  } catch (e) {
    // Besitzer evtl. nicht (mehr) als Maker im CoE – ohne Verantwortlichen erneut versuchen
    if (r.ownerId && /admin_maker|does not exist|not found/i.test(errText(e))) {
      delete values['admin_ImpactOwner@odata.bind'];
      await service.create('admin_dlpimpactanalysises', values);
      return;
    }
    throw e;
  }
}

/** Apps und Flows mit ihren Connectors aus dem CoE (für die Auswirkung). */
export function loadInventory(service: DataverseService, maxRecords: number, fresh: boolean = false): Promise<{ apps: IDlpResource[]; flows: IDlpResource[]; truncated: boolean }> {
  if (!fresh && invCache && invCache.service === service && invCache.max === maxRecords && Date.now() - invCache.at < 10 * 60000) {
    return invCache.p;
  }
  const split = (v: unknown, dropFlows: boolean): string[] => String(v || '').split(',').map(x => x.trim()).filter(x => !!x && !(dropFlows && x === 'shared_logicflows'))
    .filter((x, i, a) => a.indexOf(x) === i);
  const p = Promise.all([
    service.query('admin_apps', ['admin_appid', 'admin_displayname', 'admin_appconnections', 'admin_appenvironmentid', '_admin_appenvironment_value', '_admin_appowner_value', 'admin_appmodifiedon'],
      'admin_appdeleted ne true and admin_appconnections ne null', maxRecords + 1),
    service.query('admin_flows', ['admin_flowid', 'admin_flowidstring', 'admin_displayname', 'admin_flowconnections', 'admin_flowenvironmentid', '_admin_flowenvironment_value',
      '_admin_flowcreator_value', 'admin_flowmodifiedon'], 'admin_flowdeleted ne true and admin_flowconnections ne null', maxRecords + 1)
  ]).then(([appRows, flowRows]) => ({
    truncated: appRows.length > maxRecords || flowRows.length > maxRecords,
    apps: appRows.slice(0, maxRecords).map(r => ({
      kind: 'app' as const, id: String(r.admin_appid || ''), coeId: String(r.admin_appid || ''), name: String(r.admin_displayname || ''), env: String(r.admin_appenvironmentid || ''),
      envCoeId: String(r._admin_appenvironment_value || ''), envName: fmt(r, '_admin_appenvironment_value'), owner: fmt(r, '_admin_appowner_value'),
      ownerId: String(r._admin_appowner_value || ''), modified: String(r.admin_appmodifiedon || ''), connectors: split(r.admin_appconnections, true)
    })),
    flows: flowRows.slice(0, maxRecords).map(r => ({
      kind: 'flow' as const, id: String(r.admin_flowidstring || r.admin_flowid || ''), coeId: String(r.admin_flowid || ''), name: String(r.admin_displayname || ''),
      env: String(r.admin_flowenvironmentid || ''), envCoeId: String(r._admin_flowenvironment_value || ''), envName: fmt(r, '_admin_flowenvironment_value'),
      owner: fmt(r, '_admin_flowcreator_value'), ownerId: String(r._admin_flowcreator_value || ''), modified: String(r.admin_flowmodifiedon || ''),
      connectors: split(r.admin_flowconnections, false)
    }))
  }));
  invCache = { service, at: Date.now(), max: maxRecords, p };
  p.catch(() => { invCache = undefined; });
  return p;
}

interface IRuleSet { defaultGroup: DlpGroup; groups: { [id: string]: DlpGroup }; scope: DlpScope; envs: string[] }

function violation(rules: IRuleSet, res: IDlpResource): { business: string[]; nonBusiness: string[]; blocked: string[] } | undefined {
  if (!appliesTo(rules.scope, rules.envs, res.env)) {
    return undefined;
  }
  const business: string[] = [];
  const nonBusiness: string[] = [];
  const blocked: string[] = [];
  res.connectors.forEach(n => {
    const g = groupOf(rules, connectorId(n));
    (g === 'Blocked' ? blocked : g === 'Confidential' ? business : nonBusiness).push(n);
  });
  return blocked.length || (business.length && nonBusiness.length) ? { business, nonBusiness, blocked } : undefined;
}

/**
 * Auswirkung wie in der CoE-App: betroffen ist, was einen blockierten Connector nutzt oder Business- und Non-Business-Connectors mischt –
 * in den Environments, für die die Richtlinie gilt. „bereits betroffen“ = auch mit der bisherigen Richtlinie.
 */
export function computeImpact(w: IDlpWorking, inv: { apps: IDlpResource[]; flows: IDlpResource[]; truncated: boolean }, allEnvs: IDlpEnv[]): IDlpImpactResult {
  const rules: IRuleSet = { defaultGroup: w.defaultGroup, groups: w.groups, scope: w.scope, envs: w.envs };
  const run = (list: IDlpResource[]): { hits: IDlpImpact[]; analyzed: number } => {
    let analyzed = 0;
    const hits: IDlpImpact[] = [];
    list.forEach(r => {
      if (!appliesTo(rules.scope, rules.envs, r.env)) {
        return;
      }
      analyzed++;
      const v = violation(rules, r);
      if (v) {
        hits.push({ resource: r, ...v, already: !!w.base && !!violation(w.base, r) });
      }
    });
    return { hits, analyzed };
  };
  const a = run(inv.apps);
  const f = run(inv.flows);
  return {
    apps: a.hits, flows: f.hits, analyzedApps: a.analyzed, analyzedFlows: f.analyzed,
    envCount: allEnvs.filter(e => appliesTo(w.scope, w.envs, e.name)).length, truncated: inv.truncated
  };
}

/** Unterschiede zum Ausgangszustand (für „Überprüfen“). */
export function diffConnectors(w: IDlpWorking, connectors: IDlpConnector[]): { connector: IDlpConnector; from: DlpGroup; to: DlpGroup }[] {
  if (!w.base) {
    return [];
  }
  const base = w.base;
  return connectors.map(c => ({ connector: c, from: groupOf(base, c.id), to: groupOf(w, c.id) })).filter(x => x.from !== x.to);
}

export function isDirty(w: IDlpWorking, connectors: IDlpConnector[]): boolean {
  if (!w.base || !w.policy) {
    return true;
  }
  const envKey = (l: string[]): string => l.map(norm).sort().join('|');
  return w.displayName.trim() !== (w.policy.displayName || '') || w.defaultGroup !== w.base.defaultGroup || w.scope !== w.base.scope
    || (w.scope !== 'AllEnvironments' && envKey(w.envs) !== envKey(w.base.envs)) || diffConnectors(w, connectors).length > 0
    || Object.keys(w.groups).some(id => w.base && w.base.groups[id] !== w.groups[id]);
}

