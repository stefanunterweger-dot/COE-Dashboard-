import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { PowerPlatformApiService, PpApiError, PPAPI_VERSIONS } from './PowerPlatformApiService';

/**
 * Live-Daten ohne CoE – im Namen der angemeldeten Person.
 *
 *  - Inventar (offiziell): Power Platform API „resourcequery/resources/query“ (Azure Resource Graph, Tabelle
 *    „PowerPlatformResources“). Tenantweit Apps, Flows, Agents, Power Pages, Connectors und Environments in einem Aufruf,
 *    laut Microsoft höchstens 15 Minuten alt. Lesen dürfen Power Platform Administrator, Globaler Administrator,
 *    Dynamics 365 Administrator und Globaler Leser. Berechtigung „ResourceQuery.Resources.Read“.
 *  - Environments (offiziell): „environmentmanagement/environments“ – Zustand, Managed (Schutzstufe), Typ, URL.
 *  - Flow-Ausführungen (offiziell): „powerautomate/environments/{env}/flowRuns“ – Berechtigung „PowerAutomate.Flows.Read“.
 *  - Freigaben einer App und Zustand eines Flows: Admin-Schnittstellen von Power Apps bzw. Power Automate (nicht offiziell
 *    dokumentiert, dieselben wie die Admin-Connectors).
 */
export const INV_TYPE = {
  canvasApp: 'microsoft.powerapps/canvasapps',
  modelApp: 'microsoft.powerapps/modeldrivenapps',
  codeApp: 'microsoft.powerapps/codeapps',
  app: 'microsoft.powerapps/apps',
  cloudFlow: 'microsoft.powerautomate/cloudflows',
  agentFlow: 'microsoft.powerautomate/agentflows',
  agent: 'microsoft.copilotstudio/agents',
  pages: 'microsoft.powerpages/sites',
  connector: 'microsoft.powerplatformconnector/connectors',
  environment: 'microsoft.powerplatform/environments'
};
export const APP_TYPES = [INV_TYPE.canvasApp, INV_TYPE.modelApp, INV_TYPE.codeApp, INV_TYPE.app];
export const FLOW_TYPES = [INV_TYPE.cloudFlow, INV_TYPE.agentFlow];

export const TYPE_LABEL: { [type: string]: string } = {
  [INV_TYPE.canvasApp]: 'Canvas-Apps',
  [INV_TYPE.modelApp]: 'Model-driven Apps',
  [INV_TYPE.codeApp]: 'Code-Apps',
  [INV_TYPE.app]: 'Managed / App-Builder-Apps',
  [INV_TYPE.cloudFlow]: 'Cloud-Flows',
  [INV_TYPE.agentFlow]: 'Agent-Flows',
  'microsoft.powerautomate/m365agentflows': 'Workflow-Agent-Flows',
  [INV_TYPE.agent]: 'Copilot-Studio-Agents',
  [INV_TYPE.pages]: 'Power Pages',
  [INV_TYPE.connector]: 'Connectors',
  [INV_TYPE.environment]: 'Environments',
  'microsoft.powerplatform/environmentgroups': 'Environment-Gruppen'
};

export const INVENTORY_CONSENT_HINT = 'Das Webpart darf das Power-Platform-Inventar noch nicht abfragen. Ein Administrator muss die Berechtigung '
  + '„ResourceQuery.Resources.Read“ der „Power Platform API“ für SharePoint freigeben (scripts/Grant-CoeCockpitPermissions.ps1 -Live).';

/** Ein Eintrag des Inventars (vereinheitlicht). */
export interface IInvItem {
  /** Ressourcen-ID (App-ID, Flow-Name, Environment-ID …), klein geschrieben */
  id: string;
  type: string;
  displayName: string;
  environmentId: string;
  ownerId: string;
  createdBy: string;
  createdAt: string;
  lastModifiedAt: string;
  isQuarantined?: boolean;
  /** Environments */
  isManaged?: boolean;
  environmentType?: string;
  /** Flows */
  trigger?: string;
  triggerOperation?: string;
  workflowEntityId?: string;
  /** Detailabfrage: Connectors mit Anzahl genutzter Vorgänge */
  connectors?: { id: string; operations: string[] }[];
  /** Rohdaten (für die Diagnose) */
  raw: { [k: string]: unknown };
}

/** Environment laut „environmentmanagement/environments“. */
export interface ILiveEnvironment {
  id: string;
  displayName: string;
  type: string;
  state: string;
  protectionLevel: string;
  url: string;
  region: string;
  createdDateTime: string;
  raw: { [k: string]: unknown };
}

export interface IFlowRun {
  name: string;
  status: string;
  startTime: string;
  endTime: string;
  durationMs?: number;
  triggerType: string;
}

export interface IAppShare {
  principalId: string;
  name: string;
  email: string;
  type: string;
  role: string;
}

interface IInvResponse {
  totalRecords?: number;
  count?: number;
  resultTruncated?: unknown;
  skipToken?: string;
  $skipToken?: string;
  data?: unknown;
}

const str = (v: unknown): string => (v === undefined || v === null ? '' : String(v));
const lower = (v: unknown): string => str(v).toLowerCase();

/** „data“ kommt je nach Format als Objekt-Liste oder als Tabelle (columns/rows). */
function rowsOf(data: unknown): { [k: string]: unknown }[] {
  if (Array.isArray(data)) {
    return data as { [k: string]: unknown }[];
  }
  const t = data as { columns?: { name: string }[]; rows?: unknown[][] } | undefined;
  if (t && Array.isArray(t.columns) && Array.isArray(t.rows)) {
    return t.rows.map(row => {
      const o: { [k: string]: unknown } = {};
      t.columns!.forEach((c, i) => { o[c.name] = row[i]; });
      return o;
    });
  }
  return [];
}

function toItem(r: { [k: string]: unknown }): IInvItem {
  const p = (r.properties && typeof r.properties === 'object' ? r.properties : {}) as { [k: string]: unknown };
  const pick = (...keys: string[]): unknown => {
    for (const k of keys) {
      if (r[k] !== undefined && r[k] !== null && r[k] !== '') {
        return r[k];
      }
      if (p[k] !== undefined && p[k] !== null && p[k] !== '') {
        return p[k];
      }
    }
    return undefined;
  };
  const bool = (v: unknown): boolean | undefined => (v === true || v === 'true' ? true : v === false || v === 'false' ? false : undefined);
  const conns = pick('powerPlatformConnectors');
  return {
    id: lower(pick('name', 'id')).replace(/^.*\//, ''),
    type: lower(r.type),
    displayName: str(pick('displayName')),
    environmentId: lower(pick('environmentId')),
    ownerId: lower(pick('ownerId')),
    createdBy: lower(pick('createdBy')),
    createdAt: str(pick('createdAt')),
    lastModifiedAt: str(pick('lastModifiedAt')),
    isQuarantined: bool(pick('isQuarantined')),
    isManaged: bool(pick('isManaged', 'isManagedEnvironment')),
    environmentType: str(pick('environmentType')),
    trigger: str(pick('trigger')),
    triggerOperation: str(pick('triggerOperation')),
    workflowEntityId: lower(pick('workflowEntityId')),
    connectors: Array.isArray(conns)
      ? (conns as { connectorId?: string; operations?: { operationId?: string }[] }[]).map(c => ({
        id: str(c.connectorId),
        operations: (c.operations || []).map(o => str(o.operationId)).filter(x => !!x)
      }))
      : undefined,
    raw: r
  };
}

const quote = (v: string): string => `'${v.replace(/'/g, "\\'")}'`;

/** Felder für Listenabfragen (ohne Connector-Listen – die holt die Detailansicht einzeln). */
const LIST_FIELDS = [
  'name', 'type',
  'displayName = properties.displayName',
  'environmentId = tostring(properties.environmentId)',
  'ownerId = tostring(properties.ownerId)',
  'createdBy = tostring(properties.createdBy)',
  'createdAt = tostring(properties.createdAt)',
  'lastModifiedAt = tostring(properties.lastModifiedAt)',
  'isQuarantined = properties.isQuarantined',
  'isManaged = properties.isManaged',
  'environmentType = tostring(properties.environmentType)',
  'trigger = tostring(properties.trigger)',
  'triggerOperation = tostring(properties.triggerOperation)',
  'workflowEntityId = tostring(properties.workflowEntityId)'
];

export interface IInventoryList {
  items: IInvItem[];
  /** true = nicht alles geladen (Obergrenze erreicht) */
  truncated: boolean;
  total?: number;
  loadedAt: Date;
}

const POWERAPPS_ADMIN = 'https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments';
const FLOW_ADMIN = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments';

/** Zwischenspeicher für die Sitzung (10 Minuten) – sonst würde jeder Reiterwechsel tausende Einträge neu laden. */
const CACHE_MS = 10 * 60000;
const cache: { [key: string]: { at: number; promise: Promise<IInventoryList> } } = {};

export function resetLiveCache(): void {
  Object.keys(cache).forEach(k => { delete cache[k]; });
}

export class LiveInventoryService {
  constructor(private ppApi: PowerPlatformApiService, private factory?: AadHttpClientFactory) {}

  /** Rohe Inventar-Abfrage (eine Seite). */
  public query(clauses: unknown[], options?: { Top?: number; Skip?: number; SkipToken?: string }): Promise<IInvResponse> {
    return this.ppApi.post<IInvResponse>('/resourcequery/resources/query', {
      TableName: 'PowerPlatformResources',
      Clauses: clauses,
      Options: { Top: 1000, Skip: 0, ...(options || {}) }
    }, PPAPI_VERSIONS.inventory, INVENTORY_CONSENT_HINT);
  }

  /** Anzahl je Ressourcentyp (für Diagnose und Vergleich mit dem CoE). */
  public async countByType(): Promise<{ type: string; count: number }[]> {
    const res = await this.query([
      { $type: 'summarize', SummarizeClauseExpression: { OperatorName: 'count', OperatorFieldName: 'resourceCount', FieldList: ['type'] } }
    ]);
    return rowsOf(res.data).map(r => ({ type: lower(r.type), count: Number(r.resourceCount) || 0 }));
  }

  /** Alle Einträge der Typen (mit Paging, max. maxItems). Ergebnis 10 Minuten zwischengespeichert. */
  public list(types: string[], maxItems: number = 50000, fresh: boolean = false): Promise<IInventoryList> {
    const key = types.slice().sort().join('|') + '|' + maxItems;
    const hit = cache[key];
    if (!fresh && hit && Date.now() - hit.at < CACHE_MS) {
      return hit.promise;
    }
    const promise = this.loadAll(types, maxItems);
    cache[key] = { at: Date.now(), promise };
    promise.catch(() => { delete cache[key]; });
    return promise;
  }

  private async loadAll(types: string[], maxItems: number): Promise<IInventoryList> {
    const clauses = [
      { $type: 'where', FieldName: 'type', Operator: 'in~', Values: types.map(quote) },
      { $type: 'project', FieldList: LIST_FIELDS }
    ];
    const items: IInvItem[] = [];
    let skipToken = '';
    let total: number | undefined;
    for (let page = 0; page < 200; page++) {
      const res = await this.query(clauses, { Top: 1000, Skip: 0, SkipToken: skipToken });
      rowsOf(res.data).forEach(r => items.push(toItem(r)));
      total = typeof res.totalRecords === 'number' ? res.totalRecords : total;
      skipToken = str(res.skipToken || res.$skipToken);
      if (!skipToken || items.length >= maxItems) {
        break;
      }
    }
    return { items: items.slice(0, maxItems), truncated: !!skipToken || items.length > maxItems, total, loadedAt: new Date() };
  }

  /** Ein Eintrag mit Connectors (Detailansicht). */
  public async item(id: string, types: string[]): Promise<IInvItem | undefined> {
    const res = await this.query([
      { $type: 'where', FieldName: 'type', Operator: 'in~', Values: types.map(quote) },
      { $type: 'where', FieldName: 'name', Operator: '=~', Values: [quote(id)] },
      { $type: 'take', TakeCount: 1 }
    ]);
    const row = rowsOf(res.data)[0];
    return row ? toItem(row) : undefined;
  }

  /** Anzahl Ressourcen je Typ in einem Environment. */
  public async countInEnvironment(environmentId: string): Promise<{ type: string; count: number }[]> {
    const res = await this.query([
      { $type: 'where', FieldName: 'properties.environmentId', Operator: '=~', Values: [quote(environmentId)] },
      { $type: 'summarize', SummarizeClauseExpression: { OperatorName: 'count', OperatorFieldName: 'resourceCount', FieldList: ['type'] } }
    ]);
    return rowsOf(res.data).map(r => ({ type: lower(r.type), count: Number(r.resourceCount) || 0 }));
  }

  /** Alle Environments (offizielle Environment-Verwaltung). */
  public async environments(): Promise<ILiveEnvironment[]> {
    const list = await this.ppApi.getAll<{ [k: string]: unknown }>('/environmentmanagement/environments', PPAPI_VERSIONS.environments);
    return list.map(toEnv);
  }

  public async environment(id: string): Promise<ILiveEnvironment> {
    const e = await this.ppApi.get<{ [k: string]: unknown }>(`/environmentmanagement/environments/${encodeURIComponent(id)}`, PPAPI_VERSIONS.environments);
    return toEnv(e);
  }

  /** Letzte Ausführungen eines Flows (offiziell; workflowId = Dataverse-Workflow-ID laut Inventar). */
  public async flowRuns(environmentId: string, workflowId: string, top: number = 10): Promise<IFlowRun[]> {
    const res = await this.ppApi.get<{ value?: { [k: string]: unknown }[] }>(
      `/powerautomate/environments/${encodeURIComponent(environmentId)}/flowRuns?workflowId=${encodeURIComponent(workflowId)}`,
      PPAPI_VERSIONS.powerautomate);
    const runs = ((res && res.value) || []).map(r => {
      const p = (r.properties && typeof r.properties === 'object' ? r.properties : r) as { [k: string]: unknown };
      return {
        name: str(r.name || r.flowRunId || p.flowRunId),
        status: str(p.status),
        startTime: str(p.startTime || p.createdOn),
        endTime: str(p.endTime),
        durationMs: typeof p.durationMs === 'number' ? p.durationMs : undefined,
        triggerType: str(p.triggerType)
      };
    });
    return runs.sort((a, b) => b.startTime.localeCompare(a.startTime)).slice(0, top);
  }

  private async adminGet<T>(resource: string, url: string): Promise<T> {
    if (!this.factory) {
      throw new PpApiError('failed', 'Keine Anmeldung möglich.');
    }
    let res;
    try {
      const client = await this.factory.getClient(resource);
      res = await client.get(url, AadHttpClient.configurations.v1, { headers: { Accept: 'application/json' } });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      throw new PpApiError(/AADSTS65001|consent/i.test(msg) ? 'consent' : 'failed', msg, undefined, msg);
    }
    const text = await res.text();
    if (!res.ok) {
      let msg = text;
      try { const b = JSON.parse(text) as { error?: { message?: string } }; msg = (b.error && b.error.message) || text; } catch { /* Text */ }
      throw new PpApiError(res.status === 401 || res.status === 403 ? 'forbidden' : res.status === 404 ? 'notFound' : 'failed', msg || `HTTP ${res.status}`, res.status, msg);
    }
    return (text ? JSON.parse(text) : {}) as T;
  }

  /** Freigaben einer App (Admin-Schnittstelle von Power Apps, nicht offiziell dokumentiert). */
  public async appShares(environmentId: string, appId: string): Promise<IAppShare[]> {
    const res = await this.adminGet<{ value?: { [k: string]: unknown }[] }>('https://service.powerapps.com/',
      `${POWERAPPS_ADMIN}/${encodeURIComponent(environmentId)}/apps/${encodeURIComponent(appId)}/permissions?api-version=2016-11-01`);
    return ((res && res.value) || []).map(v => {
      const p = (v.properties || {}) as { roleName?: string; principal?: { id?: string; displayName?: string; email?: string; type?: string; tenantId?: string } };
      const pr = p.principal || {};
      return { principalId: lower(pr.id || pr.tenantId), name: str(pr.displayName), email: str(pr.email), type: str(pr.type), role: str(p.roleName) };
    });
  }

  /** Zustand eines Flows (Admin-Schnittstelle von Power Automate, nicht offiziell dokumentiert). */
  public async flowState(environmentId: string, flowName: string): Promise<{ state: string; displayName: string }> {
    const res = await this.adminGet<{ properties?: { state?: string; displayName?: string } }>('https://service.flow.microsoft.com/',
      `${FLOW_ADMIN}/${encodeURIComponent(environmentId)}/flows/${encodeURIComponent(flowName)}?api-version=2016-11-01`);
    return { state: str(res.properties && res.properties.state), displayName: str(res.properties && res.properties.displayName) };
  }
}

function toEnv(e: { [k: string]: unknown }): ILiveEnvironment {
  const p = (e.properties && typeof e.properties === 'object' ? e.properties : {}) as { [k: string]: unknown };
  const g = (k: string): string => str(e[k] !== undefined ? e[k] : p[k]);
  return {
    id: lower(e.id || e.name || e.environmentId).replace(/^.*\//, ''),
    displayName: g('displayName'),
    type: g('type') || g('environmentSku'),
    state: g('state'),
    protectionLevel: g('protectionLevel'),
    url: g('url') || g('dataverseOrganizationUrl'),
    region: g('azureRegion') || g('geo'),
    createdDateTime: g('createdDateTime'),
    raw: e
  };
}
