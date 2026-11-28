import { AadHttpClientFactory } from '@microsoft/sp-http';
import { IRecord } from './ITableConfig';
import { DataverseError, DataverseService } from '../services/DataverseService';

/**
 * Pipelines-Übersicht (Power Platform Pipelines) über alle Pipelines-Hosts – nur lesen.
 *
 * Die Daten liegen im Dataverse jedes Host-Environments (Tabellen „Deployment Pipeline“, „Deployment Stage“,
 * „Deployment Environment“, „Deployment Stage Run“ – Microsoft Learn „Pipeline table reference“), gelesen mit den Rechten der
 * angemeldeten Person. Hosts findet das Cockpit selbst: Es fragt jedes Dataverse-Environment aus dem CoE, ob es
 * Pipeline-Environments enthält (Ergebnis 24 Stunden im Browser gespeichert). Den von Microsoft verwalteten „Platform host“
 * (persönliche Pipelines) zeigt nur das Admin Center.
 */
export const PIPELINES_KEY = 'pipeline-hosts';
export const PPAC_DEPLOYMENT_URL = 'https://admin.powerplatform.microsoft.com/';

const CACHE_KEY = 'coe-cockpit:pipeline-hosts';
const CACHE_HOURS = 24;

export interface ICoeEnv {
  /** Primärschlüssel im CoE */
  coeId: string;
  /** Power-Platform-Environment-ID */
  name: string;
  displayName: string;
  type: string;
  url: string;
  /** Managed Environment laut CoE (undefined = unbekannt) */
  managed?: boolean;
}

export interface IPipelineHost { url: string; name: string; coeId: string; configured: boolean }

export interface IHostScan { hosts: IPipelineHost[]; checked: number; denied: number; failed: number; at: number }

export type EnvRole = 'dev' | 'target' | '';
export interface IPlEnv {
  id: string;
  /** Power-Platform-Environment-ID */
  envId: string;
  name: string;
  role: EnvRole;
  validation: string;
  validationOk?: boolean;
  error: string;
  coe?: ICoeEnv;
  /** Pipelines, in denen das Environment vorkommt */
  pipelines: { id: string; name: string; as: EnvRole; stage?: string }[];
}

export interface IPlStage { id: string; name: string; pipelineId: string; previousId: string; targetId: string }

export type RunState = 'ok' | 'failed' | 'running' | 'pending' | 'canceled' | '';
export interface IPlRun { id: string; stageId: string; pipelineId: string; artifact: string; version: string; state: RunState; status: string; start: string; end: string; by: string }

export interface IPipeline {
  id: string;
  name: string;
  description: string;
  owner: string;
  modified: string;
  devEnvIds: string[];
  stages: IPlStage[];
  lastRun?: IPlRun;
  runCount: number;
}

export interface IHostData {
  host: IPipelineHost;
  pipelines: IPipeline[];
  envs: IPlEnv[];
  runs: IPlRun[];
  /** Hinweis, wenn die Entwicklungs-Environments nur aus den Läufen abgeleitet werden konnten */
  devNote?: string;
  problem?: string;
  problemKind?: 'forbidden' | 'failed';
  loadedAt: Date;
}

const fmt = (r: IRecord, f: string): string => String(r[`${f}@OData.Community.Display.V1.FormattedValue`] || r[f] || '');
const str = (v: unknown): string => (v === undefined || v === null ? '' : String(v));
const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));
export const normEnv = (id: string): string => String(id || '').toLowerCase().replace(/^default-/, '');
export const normUrl = (url: string): string => String(url || '').trim().replace(/\/+$/, '').toLowerCase();

// ---------------------------------------------------------------------------------------------------------------
// CoE-Environments
// ---------------------------------------------------------------------------------------------------------------

let envCache: { service: DataverseService; at: number; p: Promise<ICoeEnv[]> } | undefined;
let dataCache: { [url: string]: { at: number; p: Promise<IHostData> } } = {};
let indexCache: { key: string; at: number; p: Promise<IPipelineIndex> } | undefined;

export function resetPipelineCache(scanToo: boolean = false): void {
  envCache = undefined;
  dataCache = {};
  indexCache = undefined;
  if (scanToo) {
    try {
      window.localStorage.removeItem(CACHE_KEY);
    } catch {
      // Browser-Speicher nicht verfügbar
    }
  }
}

export function loadCoeEnvs(service: DataverseService, fresh: boolean = false): Promise<ICoeEnv[]> {
  if (!fresh && envCache && envCache.service === service && Date.now() - envCache.at < 10 * 60000) {
    return envCache.p;
  }
  const p = service.query('admin_environments', ['admin_environmentid', 'admin_environmentname', 'admin_displayname', 'admin_environmentsku',
    'admin_environmentcdsinstanceurl', 'admin_governanceprotectionlevel', 'admin_environmentdeleted'], 'statecode eq 0', 20000).then(rows => rows
    .filter(r => r.admin_environmentdeleted !== true)
    .map(r => {
      const level = str(r.admin_governanceprotectionlevel);
      return {
        coeId: str(r.admin_environmentid), name: str(r.admin_environmentname || r.admin_environmentid), displayName: str(r.admin_displayname || r.admin_environmentname),
        type: fmt(r, 'admin_environmentsku'), url: str(r.admin_environmentcdsinstanceurl), managed: level ? !/^basic$/i.test(level) : undefined
      };
    }));
  envCache = { service, at: Date.now(), p };
  p.catch(() => { envCache = undefined; });
  return p;
}

export function envLookup(envs: ICoeEnv[]): (id: string) => ICoeEnv | undefined {
  const m: { [id: string]: ICoeEnv } = {};
  envs.forEach(e => { m[normEnv(e.name)] = e; m[normEnv(e.coeId)] = e; });
  return id => m[normEnv(id)];
}

// ---------------------------------------------------------------------------------------------------------------
// Hosts finden
// ---------------------------------------------------------------------------------------------------------------

export function readHostScan(): IHostScan | undefined {
  try {
    const raw = window.localStorage.getItem(CACHE_KEY);
    const v = raw ? JSON.parse(raw) as IHostScan : undefined;
    // mit nicht erreichbaren Environments nur 1 Stunde, mit verweigerten 6 Stunden gültig, sonst 24 Stunden
    const hours = v && v.failed ? 1 : v && v.denied ? 6 : CACHE_HOURS;
    return v && Array.isArray(v.hosts) && Date.now() - v.at < hours * 3600000 ? v : undefined;
  } catch {
    return undefined;
  }
}

function writeHostScan(v: IHostScan): void {
  try {
    window.localStorage.setItem(CACHE_KEY, JSON.stringify(v));
  } catch {
    // Browser-Speicher nicht verfügbar – dann wird beim nächsten Mal neu gesucht
  }
}

/** Environments, die ein Pipelines-Host sein können: mit Dataverse-URL, ohne Dataverse for Teams (kein API-Zugriff). */
export function hostCandidates(envs: ICoeEnv[]): ICoeEnv[] {
  return envs.filter(e => /^https:\/\//i.test(e.url) && !/teams/i.test(e.type || ''));
}

/**
 * Fragt jedes Dataverse-Environment aus dem CoE, ob es Pipeline-Environments enthält (= Pipelines-Host).
 * Environments ohne die Pipelines-Tabellen antworten mit „nicht gefunden“, ohne Zugriff mit 401/403.
 */
export async function scanHosts(factory: AadHttpClientFactory, envs: ICoeEnv[], onProgress?: (done: number, total: number) => void,
  alive: () => boolean = () => true): Promise<IHostScan> {
  const list = hostCandidates(envs);
  const found: IPipelineHost[] = [];
  let done = 0;
  let denied = 0;
  let failed = 0;
  const queue = list.slice();
  const worker = async (): Promise<void> => {
    for (let e = queue.shift(); e && alive(); e = queue.shift()) {
      try {
        const svc = new DataverseService(factory, e.url);
        const rows = await svc.query('deploymentenvironments', ['deploymentenvironmentid'], undefined, 1);
        if (rows.length) {
          found.push({ url: e.url.replace(/\/+$/, ''), name: e.displayName, coeId: e.coeId, configured: false });
        }
      } catch (err) {
        const status = err instanceof DataverseError ? err.status : 0;
        if (status === 401 || status === 403) {
          denied++;
        } else if (status !== 404 && status !== 400) {
          failed++;
        }
      }
      done++;
      if (onProgress) {
        onProgress(done, list.length);
      }
    }
  };
  await Promise.all([worker(), worker(), worker(), worker(), worker(), worker()]);
  const result: IHostScan = { hosts: found.sort((a, b) => a.name.localeCompare(b.name)), checked: done, denied, failed, at: Date.now() };
  // vollständige Suchen merken – außer die Mehrheit der Environments hat den Zugriff verweigert (dann ist meist die Rolle nicht aktiv)
  if (alive() && done === list.length && denied * 2 < Math.max(1, done)) {
    writeHostScan(result);
  }
  return result;
}

/** Hosts aus Einstellung (Pipelines-Host der Webpart-Einstellungen) und Suche, ohne Doppelte. */
export function mergeHosts(configuredUrl: string, scan: IHostScan | undefined, envs: ICoeEnv[]): IPipelineHost[] {
  const out: IPipelineHost[] = [];
  const byUrl: { [u: string]: ICoeEnv } = {};
  envs.forEach(e => { if (e.url) { byUrl[normUrl(e.url)] = e; } });
  if (configuredUrl) {
    const e = byUrl[normUrl(configuredUrl)];
    out.push({ url: configuredUrl.replace(/\/+$/, ''), name: e ? e.displayName : configuredUrl.replace(/^https:\/\//, ''), coeId: e ? e.coeId : '', configured: true });
  }
  (scan ? scan.hosts : []).forEach(h => {
    if (!out.some(o => normUrl(o.url) === normUrl(h.url))) {
      out.push({ ...h, configured: false });
    }
  });
  return out;
}

// ---------------------------------------------------------------------------------------------------------------
// Daten eines Hosts
// ---------------------------------------------------------------------------------------------------------------

function runState(label: string): RunState {
  if (/not ?started|nicht gestartet/i.test(label)) {
    return '';
  }
  if (/succe|erfolg/i.test(label)) {
    return 'ok';
  }
  if (/fail|fehl/i.test(label)) {
    return 'failed';
  }
  if (/cancel|abgebr/i.test(label)) {
    return 'canceled';
  }
  if (/pending|ausstehend|genehmig|approval|scheduled|geplant/i.test(label)) {
    return 'pending';
  }
  if (/progress|start|ausgef|gestartet|running/i.test(label)) {
    return 'running';
  }
  return '';
}

interface IManyToMany { Entity1LogicalName?: string; Entity2LogicalName?: string; Entity1NavigationPropertyName?: string; Entity2NavigationPropertyName?: string }

async function pagedRaw(svc: DataverseService, path: string, max: number): Promise<IRecord[]> {
  let out: IRecord[] = [];
  let url: string | undefined = path;
  while (url && out.length < max) {
    const page: { value?: IRecord[]; '@odata.nextLink'?: string } = await svc.raw(url, Math.min(5000, max));
    out = out.concat(page.value || []);
    url = page['@odata.nextLink'];
  }
  return out;
}

export function loadHost(factory: AadHttpClientFactory, host: IPipelineHost, coeEnvs: ICoeEnv[], fresh: boolean = false): Promise<IHostData> {
  const key = normUrl(host.url);
  const cached = dataCache[key];
  if (!fresh && cached && Date.now() - cached.at < 10 * 60000) {
    return cached.p;
  }
  const p = loadHostNow(factory, host, coeEnvs);
  dataCache[key] = { at: Date.now(), p };
  return p;
}

async function loadHostNow(factory: AadHttpClientFactory, host: IPipelineHost, coeEnvs: ICoeEnv[]): Promise<IHostData> {
  const out: IHostData = { host, pipelines: [], envs: [], runs: [], loadedAt: new Date() };
  const svc = new DataverseService(factory, host.url);
  const coe = envLookup(coeEnvs);
  let pipeRows: IRecord[];
  let stageRows: IRecord[];
  let envRows: IRecord[];
  try {
    [envRows, stageRows] = await Promise.all([
      svc.query('deploymentenvironments', ['deploymentenvironmentid', 'name', 'environmentid', 'environmenttype', 'validationstatus', 'errormessage'], 'statecode eq 0', 20000),
      svc.query('deploymentstages', ['deploymentstageid', 'name', '_deploymentpipelineid_value', '_previousdeploymentstageid_value', '_targetdeploymentenvironmentid_value'], 'statecode eq 0', 20000)
    ]);
    // Entwicklungs-Environments: N:N „Deployment Pipeline“ ↔ „Deployment Environment“ (Name der Beziehung aus den Metadaten)
    let nav = '';
    try {
      const rels = await svc.raw<{ value?: IManyToMany[] }>(
        "EntityDefinitions(LogicalName='deploymentpipeline')/ManyToManyRelationships?$select=Entity1LogicalName,Entity2LogicalName,Entity1NavigationPropertyName,Entity2NavigationPropertyName");
      const rel = (rels.value || []).filter(r => (r.Entity1LogicalName === 'deploymentpipeline' && r.Entity2LogicalName === 'deploymentenvironment')
        || (r.Entity2LogicalName === 'deploymentpipeline' && r.Entity1LogicalName === 'deploymentenvironment'))[0];
      nav = rel ? String((rel.Entity1LogicalName === 'deploymentpipeline' ? rel.Entity1NavigationPropertyName : rel.Entity2NavigationPropertyName) || '') : '';
    } catch {
      nav = '';
    }
    const base = 'deploymentpipelines?$select=deploymentpipelineid,name,description,_ownerid_value,modifiedon&$filter=statecode eq 0';
    pipeRows = [];
    if (nav) {
      try {
        pipeRows = await pagedRaw(svc, `${base}&$expand=${nav}($select=deploymentenvironmentid)`, 20000);
      } catch {
        nav = '';
      }
    }
    if (!nav) {
      pipeRows = await pagedRaw(svc, base, 20000);
      out.devNote = 'Entwicklungs-Environments aus den Bereitstellungen abgeleitet (Beziehung Pipeline ↔ Environment nicht lesbar).';
    }
    const navName = nav;
    // die letzten 1000 Bereitstellungen (für „letzte Bereitstellung“ je Pipeline)
    const runSelect = 'deploymentstagerunid,artifactname,artifactversion,_deploymentstageid_value,stagerunstatus,starttime,endtime,_createdby_value';
    const runRows = await pagedRaw(svc, `deploymentstageruns?$select=${runSelect},_devdeploymentenvironment_value&$orderby=starttime desc`, 1000)
      .catch(() => pagedRaw(svc, `deploymentstageruns?$select=${runSelect}&$orderby=starttime desc`, 1000))
      .catch(() => [] as IRecord[]);

    const stageById: { [id: string]: IPlStage } = {};
    const stages: IPlStage[] = stageRows.map(r => ({
      id: str(r.deploymentstageid), name: str(r.name), pipelineId: str(r._deploymentpipelineid_value),
      previousId: str(r._previousdeploymentstageid_value), targetId: str(r._targetdeploymentenvironmentid_value)
    }));
    stages.forEach(s => { stageById[s.id] = s; });
    out.runs = runRows.map(r => {
      const st = stageById[str(r._deploymentstageid_value)];
      const status = fmt(r, 'stagerunstatus');
      return {
        id: str(r.deploymentstagerunid), stageId: str(r._deploymentstageid_value), pipelineId: st ? st.pipelineId : '', artifact: str(r.artifactname),
        version: str(r.artifactversion), state: runState(status), status, start: str(r.starttime), end: str(r.endtime), by: fmt(r, '_createdby_value')
      };
    });
    const devFromRuns: { [pid: string]: string[] } = {};
    if (!navName) {
      runRows.forEach(r => {
        const st = stageById[str(r._deploymentstageid_value)];
        const dev = str(r._devdeploymentenvironment_value);
        if (st && dev) {
          const l = devFromRuns[st.pipelineId] = devFromRuns[st.pipelineId] || [];
          if (l.indexOf(dev) < 0) {
            l.push(dev);
          }
        }
      });
    }

    out.envs = envRows.map(r => {
      const type = str(r.environmenttype);
      const roleLabel = fmt(r, 'environmenttype');
      const role: EnvRole = type === '200000000' || /develop|entwickl/i.test(roleLabel) ? 'dev' : type === '200000001' || /target|ziel/i.test(roleLabel) ? 'target' : '';
      const v = str(r.validationstatus);
      const vLabel = fmt(r, 'validationstatus');
      return {
        id: str(r.deploymentenvironmentid), envId: str(r.environmentid), name: str(r.name), role, validation: vLabel,
        validationOk: v === '200000001' || /succe|erfolg/i.test(vLabel) ? true : v === '200000002' || /fail|fehl/i.test(vLabel) ? false : undefined,
        error: str(r.errormessage), coe: coe(str(r.environmentid)), pipelines: []
      };
    });
    const envById: { [id: string]: IPlEnv } = {};
    out.envs.forEach(e => { envById[e.id] = e; });

    out.pipelines = pipeRows.map(r => {
      const id = str(r.deploymentpipelineid);
      const linked = navName ? ((r[navName] || []) as IRecord[]).map(x => str(x.deploymentenvironmentid)) : (devFromRuns[id] || []);
      // Phasen in der richtigen Reihenfolge (über „vorherige Phase“)
      const own = stages.filter(s => s.pipelineId === id);
      const ordered: IPlStage[] = [];
      let next = own.filter(s => !s.previousId || !own.some(o => o.id === s.previousId))[0];
      while (next && ordered.indexOf(next) < 0) {
        ordered.push(next);
        const cur: IPlStage = next;
        next = own.filter(s => s.previousId === cur.id)[0];
      }
      own.filter(s => ordered.indexOf(s) < 0).forEach(s => ordered.push(s));
      const runs = out.runs.filter(x => x.pipelineId === id);
      return {
        id, name: str(r.name), description: str(r.description), owner: fmt(r, '_ownerid_value'), modified: str(r.modifiedon),
        devEnvIds: linked, stages: ordered, lastRun: runs[0], runCount: runs.length
      };
    }).sort((a, b) => a.name.localeCompare(b.name));

    out.pipelines.forEach(pl => {
      pl.devEnvIds.forEach(id => { const e = envById[id]; if (e) { e.pipelines.push({ id: pl.id, name: pl.name, as: 'dev' }); } });
      pl.stages.forEach(s => { const e = envById[s.targetId]; if (e) { e.pipelines.push({ id: pl.id, name: pl.name, as: 'target', stage: s.name }); } });
    });
  } catch (e) {
    const status = e instanceof DataverseError ? e.status : 0;
    out.problemKind = status === 401 || status === 403 ? 'forbidden' : 'failed';
    out.problem = out.problemKind === 'forbidden'
      ? 'Kein Zugriff auf dieses Host-Environment – dafür braucht es eine Sicherheitsrolle im Host (z. B. „Systemadministrator“: im Power Platform Admin Center beim Environment unter „Mitgliedschaft“ selbst zuweisen, nach jeder PIM-Aktivierung erneut, oder dauerhaft „Deployment Pipeline Administrator“).'
      : `Nicht lesbar: ${errText(e)}`;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------------------------
// Index Environment → Pipelines (für „Betroffen ab Februar 2027“)
// ---------------------------------------------------------------------------------------------------------------

export interface IEnvPipelineInfo {
  /** Ziel in diesen Pipelines – muss Managed bleiben */
  target: string[];
  /** Entwicklung in diesen Pipelines */
  dev: string[];
  /** selbst ein Pipelines-Host */
  host: boolean;
}

export interface IPipelineIndex {
  byEnv: { [envId: string]: IEnvPipelineInfo };
  hosts: number;
  unreadable: string[];
  /** Environments, die bei der Host-Suche nicht geprüft werden konnten (ohne Zugriff oder nicht erreichbar) */
  notChecked: number;
  /** Hosts wurden gesucht (sonst nur der eingestellte Host) */
  scanned: boolean;
  loadedAt: Date;
}

/** Alle Hosts laden und je Environment sammeln, in welchen Pipelines es vorkommt. Sucht Hosts, wenn noch keine gefunden wurden. */
export function loadPipelineIndex(factory: AadHttpClientFactory, service: DataverseService, configuredUrl: string,
  onProgress?: (text: string) => void): Promise<IPipelineIndex> {
  const key = `${service.baseUrl}|${configuredUrl}`;
  if (indexCache && indexCache.key === key && Date.now() - indexCache.at < 10 * 60000) {
    return indexCache.p;
  }
  const p = (async (): Promise<IPipelineIndex> => {
    const envs = await loadCoeEnvs(service);
    let scan = readHostScan();
    if (!scan) {
      scan = await scanHosts(factory, envs, (d, t) => onProgress && onProgress(`Pipelines-Hosts werden gesucht … ${d} von ${t} Environments`));
    }
    const hosts = mergeHosts(configuredUrl, scan, envs);
    const data = await Promise.all(hosts.map(h => loadHost(factory, h, envs)));
    const byEnv: { [envId: string]: IEnvPipelineInfo } = {};
    const get = (id: string): IEnvPipelineInfo => (byEnv[normEnv(id)] = byEnv[normEnv(id)] || { target: [], dev: [], host: false });
    const unreadable: string[] = [];
    data.forEach(d => {
      if (d.problem) {
        unreadable.push(d.host.name);
      }
      if (d.host.coeId) {
        get(d.host.coeId).host = true;
        const he = envs.filter(e => e.coeId === d.host.coeId)[0];
        if (he) {
          get(he.name).host = true;
        }
      }
      d.envs.forEach(e => {
        e.pipelines.forEach(pl => {
          const info = get(e.envId);
          const label = `${pl.name} (${d.host.name})`;
          const list = pl.as === 'target' ? info.target : info.dev;
          if (list.indexOf(label) < 0) {
            list.push(label);
          }
        });
      });
    });
    return { byEnv, hosts: hosts.length, unreadable, notChecked: (scan.denied || 0) + (scan.failed || 0), scanned: true, loadedAt: new Date() };
  })();
  indexCache = { key, at: Date.now(), p };
  p.catch(() => { indexCache = undefined; });
  return p;
}

export function pipelineInfo(index: IPipelineIndex | undefined, ...ids: string[]): IEnvPipelineInfo | undefined {
  if (!index) {
    return undefined;
  }
  for (const id of ids) {
    const v = id ? index.byEnv[normEnv(id)] : undefined;
    if (v) {
      return v;
    }
  }
  return undefined;
}
