import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord, IRelatedConfig } from '../models/ITableConfig';
import { DataverseError, DataverseService } from '../services/DataverseService';
import { displayValue, formatNumber, rawValue } from '../utils/format';
import { envSolutionId } from '../models/actions';

/**
 * „Inhalt der Solution“ im Detailbereich einer Solution:
 *  - Apps und Flows aus dem CoE-Inventar (N:N-Beziehungen admin_PPSolution_Apps / admin_PPSolution_Flow, anklickbar)
 *  - Bots, Desktop Flows und Power Pages live aus dem Environment der Solution (Tabelle solutioncomponents) –
 *    dafür gibt es im CoE keine Zuordnung. Ebenfalls live, falls das CoE für Apps/Flows nichts geliefert hat.
 */
export type SolutionKind = 'apps' | 'flows' | 'bots' | 'desktop' | 'pages';

export const SOLUTION_KINDS: { key: SolutionKind; label: string; icon: string }[] = [
  { key: 'apps', label: 'Apps', icon: 'AppIconDefault' },
  { key: 'flows', label: 'Flows', icon: 'Flow' },
  { key: 'bots', label: 'Bots', icon: 'Robot' },
  { key: 'desktop', label: 'Desktop Flows', icon: 'Processing' },
  { key: 'pages', label: 'Power Pages', icon: 'Globe' }
];

export interface ISolutionItem {
  id: string;
  name: string;
  sub?: string;
  /** Datensatz im CoE (dann anklickbar) */
  coe?: { logicalName: string; id: string };
}

export interface ISolutionContents {
  items: { [k in SolutionKind]: ISolutionItem[] };
  source: { [k in SolutionKind]: 'coe' | 'live' | '' };
  envName: string;
  envUrl: string;
  makerUrl: string;
  /** Live-Abfrage im Environment nicht möglich */
  liveProblem?: string;
  coeProblem?: string;
}

/** Feste Komponententypen; Bots und Power Pages haben je Environment eigene Nummern (solutioncomponentdefinitions). */
const TYPE_WORKFLOW = 29;
const TYPE_CANVAS = 300;
const TYPE_APPMODULE = 80;
const CHUNK = 20;

const APPS_REL: IRelatedConfig = {
  key: 'solapps', title: 'Apps', logicalName: 'admin_app', navigationProperty: 'admin_PPSolution_Apps',
  columns: [
    { field: 'admin_displayname', label: 'App', type: 'title' },
    { field: 'admin_powerappstype', label: 'Typ', type: 'choice' },
    { field: '_admin_appowner_value', label: 'Besitzer', type: 'person' }
  ]
};
const FLOWS_REL: IRelatedConfig = {
  key: 'solflows', title: 'Flows', logicalName: 'admin_flow', navigationProperty: 'admin_PPSolution_Flow',
  columns: [
    { field: 'admin_displayname', label: 'Flow', type: 'title' },
    { field: 'admin_flowstate', label: 'Status', type: 'text' },
    { field: '_admin_flowcreator_value', label: 'Maker', type: 'person' }
  ]
};

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function chunks<T>(a: T[], n: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < a.length; i += n) {
    out.push(a.slice(i, i + n));
  }
  return out;
}

function orIds(field: string, ids: string[]): string {
  return ids.map(id => `${field} eq ${id}`).join(' or ');
}

/** Datensätze einer Tabelle zu den IDs (in Paketen); fehlende Tabelle → leer. */
async function byIds(env: DataverseService, set: string, idField: string, select: string[], ids: string[]): Promise<IRecord[]> {
  if (!ids.length) {
    return [];
  }
  const out: IRecord[] = [];
  for (const part of chunks(ids, CHUNK)) {
    try {
      out.push(...await env.query(set, [idField].concat(select), orIds(idField, part), CHUNK * 2));
    } catch (e) {
      const st = (e as DataverseError).status;
      if (st === 404 || st === 400) {
        return out; // Tabelle gibt es in diesem Environment nicht
      }
      throw e;
    }
  }
  return out;
}

function sortItems(a: ISolutionItem[]): ISolutionItem[] {
  return a.sort((x, y) => x.name.localeCompare(y.name, 'de'));
}

const FLOW_STATE: { [k: string]: string } = { started: 'Ein', stopped: 'Aus', suspended: 'Angehalten' };

export async function loadSolutionContents(
  coe: DataverseService, serviceFor: (url: string) => DataverseService, solutionId: string, listRecord: IRecord
): Promise<ISolutionContents> {
  const empty = (): { [k in SolutionKind]: ISolutionItem[] } => ({ apps: [], flows: [], bots: [], desktop: [], pages: [] });
  const res: ISolutionContents = {
    items: empty(), source: { apps: '', flows: '', bots: '', desktop: '', pages: '' }, envName: '', envUrl: '', makerUrl: ''
  };

  const related = Promise.all([
    coe.getRelated('admin_solution', solutionId, APPS_REL, 2000).then(r => r, (e: Error) => { res.coeProblem = e.message; return undefined; }),
    coe.getRelated('admin_solution', solutionId, FLOWS_REL, 2000).then(r => r, (e: Error) => { res.coeProblem = e.message; return undefined; })
  ]);

  // Vollständiger Datensatz (Environment, Solution-ID im Environment)
  let sol: IRecord = listRecord;
  try {
    sol = await coe.getRecord('admin_solution', solutionId);
  } catch {
    // Liste reicht notfalls
  }
  const envId = rawValue(sol, '_admin_solutionenvironment_value');
  // Solution-ID im Environment (die CoE-ID ist eine eigene, zufällige ID)
  const liveSolutionId = envSolutionId(sol);
  res.envName = displayValue(sol, '_admin_solutionenvironment_value') || String(sol.admin_solutionenvironmentdisplayname || '');

  // CoE: Apps und Flows (N:N) – läuft parallel zur Ermittlung des Environments
  const [apps, flows] = await related;
  if (apps && apps.records.length) {
    res.items.apps = sortItems(apps.records.map(r => ({
      id: String(r[apps.meta.primaryId]), name: displayValue(r, 'admin_displayname') || '(ohne Namen)',
      sub: displayValue(r, 'admin_powerappstype'), coe: { logicalName: 'admin_app', id: String(r[apps.meta.primaryId]) }
    })));
    res.source.apps = 'coe';
  }
  if (flows && flows.records.length) {
    res.items.flows = sortItems(flows.records.map(r => {
      const st = rawValue(r, 'admin_flowstate');
      return {
        id: String(r[flows.meta.primaryId]), name: displayValue(r, 'admin_displayname') || '(ohne Namen)',
        sub: FLOW_STATE[st.toLowerCase()] || st, coe: { logicalName: 'admin_flow', id: String(r[flows.meta.primaryId]) }
      };
    }));
    res.source.flows = 'coe';
  }

  // Environment: Dataverse-URL und Power-Platform-ID
  let envUrl = '';
  if (envId) {
    try {
      const env = await coe.getRecord('admin_environment', envId);
      envUrl = String(env.admin_environmentcdsinstanceurl || '');
      res.envName = res.envName || String(env.admin_displayname || '');
      const pid = String(env.admin_environmentname || '');
      if (pid && liveSolutionId) {
        res.makerUrl = `https://make.powerapps.com/environments/${encodeURIComponent(pid)}/solutions/${encodeURIComponent(liveSolutionId)}`;
      }
    } catch {
      // ohne Environment keine Live-Abfrage
    }
  }
  res.envUrl = envUrl;
  if (!envUrl) {
    res.liveProblem = 'Für das Environment der Solution ist im CoE keine Dataverse-URL hinterlegt – Bots, Desktop Flows und Power Pages lassen sich nicht abfragen.';
    return res;
  }
  if (!liveSolutionId) {
    res.liveProblem = 'Die ID der Solution im Environment ist im CoE nicht erfasst – Bots, Desktop Flows und Power Pages lassen sich nicht abfragen.';
    return res;
  }

  // Live: Komponenten der Solution im Environment
  const env = serviceFor(envUrl.replace(/\/+$/, ''));
  try {
    const special: { [type: number]: string } = {};
    let typeFilter = `componenttype eq ${TYPE_WORKFLOW} or componenttype eq ${TYPE_CANVAS} or componenttype eq ${TYPE_APPMODULE}`;
    try {
      const defs = await env.query('solutioncomponentdefinitions', ['solutioncomponenttype', 'primaryentityname'],
        "primaryentityname eq 'bot' or primaryentityname eq 'powerpagesite' or primaryentityname eq 'adx_website'", 20);
      defs.forEach(d => { special[Number(d.solutioncomponenttype)] = String(d.primaryentityname); });
      Object.keys(special).forEach(t => { typeFilter += ` or componenttype eq ${t}`; });
    } catch {
      typeFilter += ' or componenttype ge 10000';
    }
    const comps = await env.query('solutioncomponents', ['componenttype', 'objectid'], `_solutionid_value eq ${liveSolutionId} and (${typeFilter})`, 5000);
    const idsOf = (pred: (t: number) => boolean): string[] => comps.filter(c => pred(Number(c.componenttype))).map(c => String(c.objectid)).filter(x => GUID.test(x));
    const isSpecial = (t: number, entity: string): boolean => special[t] === entity || (!Object.keys(special).length && t >= 10000);

    const [wfs, canvas, modules, bots, sites, websites] = await Promise.all([
      byIds(env, 'workflows', 'workflowid', ['name', 'category', 'statecode'], idsOf(t => t === TYPE_WORKFLOW)),
      res.source.apps ? Promise.resolve([] as IRecord[]) : byIds(env, 'canvasapps', 'canvasappid', ['displayname', 'name'], idsOf(t => t === TYPE_CANVAS)),
      res.source.apps ? Promise.resolve([] as IRecord[]) : byIds(env, 'appmodules', 'appmoduleid', ['name'], idsOf(t => t === TYPE_APPMODULE)),
      byIds(env, 'bots', 'botid', ['name', 'statecode'], idsOf(t => isSpecial(t, 'bot'))),
      byIds(env, 'powerpagesites', 'powerpagesiteid', ['name'], idsOf(t => isSpecial(t, 'powerpagesite'))),
      byIds(env, 'adx_websites', 'adx_websiteid', ['adx_name'], idsOf(t => isSpecial(t, 'adx_website')))
    ]);

    if (!res.source.apps && (canvas.length || modules.length)) {
      res.items.apps = sortItems(canvas.map(r => ({ id: String(r.canvasappid), name: String(r.displayname || r.name || ''), sub: 'Canvas-App' }))
        .concat(modules.map(r => ({ id: String(r.appmoduleid), name: String(r.name || ''), sub: 'Modellgesteuerte App' }))));
      res.source.apps = 'live';
    }
    const cloud = wfs.filter(w => Number(w.category) === 5);
    if (!res.source.flows && cloud.length) {
      res.items.flows = sortItems(cloud.map(w => ({ id: String(w.workflowid), name: String(w.name || ''), sub: Number(w.statecode) === 1 ? 'Ein' : 'Aus' })));
      res.source.flows = 'live';
    }
    res.items.desktop = sortItems(wfs.filter(w => Number(w.category) === 6).map(w => ({ id: String(w.workflowid), name: String(w.name || ''), sub: 'Desktop Flow' })));
    res.items.bots = sortItems(bots.map(b => ({ id: String(b.botid), name: String(b.name || ''), sub: Number(b.statecode) === 0 ? 'Aktiv' : 'Inaktiv' })));
    res.items.pages = sortItems(sites.map(p => ({ id: String(p.powerpagesiteid), name: String(p.name || ''), sub: 'Power Pages-Site' }))
      .concat(websites.map(p => ({ id: String(p.adx_websiteid), name: String(p.adx_name || ''), sub: 'Power Pages-Site (Standardmodell)' }))));
    (['bots', 'desktop', 'pages'] as SolutionKind[]).forEach(k => { res.source[k] = 'live'; });
  } catch (e) {
    const st = (e as DataverseError).status;
    res.liveProblem = st === 401 || st === 403
      ? `Kein Zugriff auf das Environment „${res.envName || envUrl}“. Für Bots, Desktop Flows und Power Pages brauchst du dort Leserechte (z. B. System-Administrator).`
      : `Das Environment „${res.envName || envUrl}“ konnte nicht abgefragt werden: ${(e as Error).message}`;
  }
  return res;
}

export interface ISolutionContentsProps {
  coe: DataverseService;
  serviceFor: (url: string) => DataverseService;
  solutionId: string;
  record: IRecord;
  canOpen?: (logicalName: string) => boolean;
  onOpenRelated?: (logicalName: string, id: string) => void;
}

const SHOW_MAX = 60;

export const SolutionContents: React.FC<ISolutionContentsProps> = props => {
  const [state, setState] = React.useState<{ status: 'loading' | 'ready' | 'error'; data?: ISolutionContents; error?: string }>({ status: 'loading' });
  const [kind, setKind] = React.useState<SolutionKind>('apps');

  React.useEffect(() => {
    let cancelled = false;
    setState({ status: 'loading' });
    loadSolutionContents(props.coe, props.serviceFor, props.solutionId, props.record).then(
      data => {
        if (cancelled) {
          return;
        }
        setState({ status: 'ready', data });
        // Ersten Reiter mit Inhalt vorauswählen
        const first = SOLUTION_KINDS.filter(k => data.items[k.key].length > 0)[0];
        setKind(first ? first.key : 'apps');
      },
      (e: Error) => { if (!cancelled) { setState({ status: 'error', error: e.message }); } }
    );
    return () => { cancelled = true; };
  }, [props.solutionId]);

  if (state.status === 'loading') {
    return (
      <div className={styles.solBox}>
        <div className={styles.solLoading}>
          <span className={styles.skeleton} style={{ width: '70%', height: 12 }} />
          <span className={styles.skeleton} style={{ width: '45%', height: 12 }} />
          <span className={styles.muted}>Inhalt wird ermittelt …</span>
        </div>
      </div>
    );
  }
  if (state.status === 'error' || !state.data) {
    return <div className={styles.ovFailed}><Icon iconName="Info" /> Inhalt nicht verfügbar – {state.error}</div>;
  }
  const d = state.data;
  const items = d.items[kind];
  const src = d.source[kind];
  const coeKinds = (['apps', 'flows'] as SolutionKind[]).filter(k => d.source[k] === 'coe').map(k => SOLUTION_KINDS.filter(x => x.key === k)[0].label);
  return (
    <div className={styles.solBox}>
      <div className={styles.solTabs} role="tablist" aria-label="Inhalt der Solution">
        {SOLUTION_KINDS.map(k => {
          const n = d.items[k.key].length;
          const unknown = !d.source[k.key] && !!d.liveProblem && n === 0;
          return (
            <button
              key={k.key}
              type="button"
              role="tab"
              aria-selected={kind === k.key}
              className={`${styles.solTab} ${kind === k.key ? styles.solTabActive : ''}`}
              onClick={() => setKind(k.key)}
            >
              {k.label} <span className={styles.solCount}>{unknown ? '?' : formatNumber(n)}</span>
            </button>
          );
        })}
      </div>
      <div className={styles.solBody} role="tabpanel">
        {items.length === 0
          ? (
            <div className={styles.solEmpty}>
              {!src && d.liveProblem
                ? <><Icon iconName="Lock" /> {d.liveProblem}</>
                : <><Icon iconName="CheckMark" /> Keine {SOLUTION_KINDS.filter(k => k.key === kind)[0].label} in dieser Solution.</>}
            </div>
          )
          : (
            <ul className={styles.solList}>
              {items.slice(0, SHOW_MAX).map(it => {
                const open = it.coe && props.onOpenRelated && (!props.canOpen || props.canOpen(it.coe.logicalName));
                const content = (
                  <>
                    <span className={styles.solIcon}><Icon iconName={SOLUTION_KINDS.filter(k => k.key === kind)[0].icon} /></span>
                    <span className={styles.solText}>
                      <span className={styles.solName}>{it.name || it.id}</span>
                      {it.sub && <span className={styles.solSub}>{it.sub}</span>}
                    </span>
                    {open && <Icon iconName="ChevronRight" className={styles.govChevron} />}
                  </>
                );
                return (
                  <li key={it.id}>
                    {open
                      ? <button type="button" className={styles.solItem} onClick={() => props.onOpenRelated!(it.coe!.logicalName, it.coe!.id)}>{content}</button>
                      : <div className={styles.solItem}>{content}</div>}
                  </li>
                );
              })}
              {items.length > SHOW_MAX && <li className={styles.muted}>… und {formatNumber(items.length - SHOW_MAX)} weitere</li>}
            </ul>
          )}
      </div>
      <div className={styles.solFoot}>
        <span>
          {src === 'coe' ? 'Quelle: CoE-Inventar' : src === 'live' ? `Quelle: live aus „${d.envName || d.envUrl}“` : ''}
          {src && coeKinds.length > 0 && src === 'live' ? ` · ${coeKinds.join(' und ')} aus dem CoE-Inventar` : ''}
        </span>
        {d.makerUrl && (
          <a className={styles.link} href={d.makerUrl} target="_blank" rel="noopener noreferrer">
            Solution in Power Apps öffnen <Icon iconName="OpenInNewWindow" />
          </a>
        )}
      </div>
    </div>
  );
};
