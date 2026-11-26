import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import styles from './CoeCockpit.module.scss';
import { copyText } from './NotifyDialog';
import { DataverseService } from '../services/DataverseService';
import { APP_TYPES, FLOW_TYPES, IInvItem, INV_TYPE, LiveInventoryService, TYPE_LABEL } from '../services/LiveInventoryService';
import { PowerPlatformApiService, PpApiError, PPAPI_VERSIONS } from '../services/PowerPlatformApiService';
import { IRoleStatus } from '../services/RoleCheckService';

/**
 * „Live-Schnittstellen“: prüft im eigenen Tenant jede Schnittstelle, die das Cockpit live nutzen kann, und zeigt
 * Ergebnis, Anzahl und eine Stichprobe. „Ergebnis kopieren“ liefert einen Text zum Weitergeben (ohne Tokens).
 */
export interface ILiveDiagnosticsProps {
  live: LiveInventoryService;
  ppApi: PowerPlatformApiService;
  coe: DataverseService | undefined;
  graphFactory: MSGraphClientFactory;
  role: IRoleStatus | undefined;
  themeVars: React.CSSProperties;
}

type ProbeState = 'idle' | 'running' | 'ok' | 'warn' | 'error' | 'skipped';
interface IProbeResult { state: ProbeState; summary: string; sample?: string; ms?: number }
interface IProbe {
  key: string;
  name: string;
  api: string;
  official: boolean;
  permission: string;
  usedFor: string;
}

const PROBES: IProbe[] = [
  { key: 'role', name: 'Admin-Rolle (Anmelde-Token)', api: 'Entra-Token, Claim „wids“', official: true, permission: '–', usedFor: 'Aktions-Schaltflächen, Live-Daten' },
  { key: 'coe', name: 'CoE (Dataverse)', api: 'Dataverse Web API – WhoAmI', official: true, permission: 'Dynamics CRM: user_impersonation', usedFor: 'alle CoE-Ansichten, Nutzung, Verlauf' },
  { key: 'envs', name: 'Environments', api: 'Power Platform API – environmentmanagement/environments', official: true, permission: 'EnvironmentManagement.Environments.Read', usedFor: 'Live-Abgleich Environments, Copilot-Credits' },
  { key: 'envDetail', name: 'Environment-Details', api: 'Power Platform API – environmentmanagement/environments/{id}', official: true, permission: 'EnvironmentManagement.Environments.Read', usedFor: 'Detailansicht Environment (Zustand, Managed)' },
  { key: 'invCount', name: 'Inventar – Anzahl je Typ', api: 'Power Platform API – resourcequery/resources/query', official: true, permission: 'ResourceQuery.Resources.Read', usedFor: 'Live-Abgleich Apps/Flows/Environments, Vergleich mit dem CoE' },
  { key: 'invApps', name: 'Inventar – Apps', api: 'resourcequery (canvasapps, modeldrivenapps …)', official: true, permission: 'ResourceQuery.Resources.Read', usedFor: 'Live-Abgleich Apps, Detailansicht (Connectors, Quarantäne)' },
  { key: 'invFlows', name: 'Inventar – Flows', api: 'resourcequery (cloudflows, agentflows)', official: true, permission: 'ResourceQuery.Resources.Read', usedFor: 'Live-Abgleich Flows, Detailansicht (Auslöser, Connectors)' },
  { key: 'invAgents', name: 'Inventar – Copilot-Studio-Agents', api: 'resourcequery (copilotstudio/agents)', official: true, permission: 'ResourceQuery.Resources.Read', usedFor: 'künftig: Agents live' },
  { key: 'invPages', name: 'Inventar – Power Pages', api: 'resourcequery (powerpages/sites, Vorschau)', official: true, permission: 'ResourceQuery.Resources.Read', usedFor: 'künftig: Power Pages live' },
  { key: 'runs', name: 'Flow-Ausführungen', api: 'Power Platform API – powerautomate/environments/{env}/flowRuns', official: true, permission: 'PowerAutomate.Flows.Read', usedFor: 'Detailansicht Flow: letzte Ausführungen' },
  { key: 'appShares', name: 'Freigaben einer App', api: 'Power-Apps-Admin-Schnittstelle (api.powerapps.com)', official: false, permission: 'PowerApps Service: User', usedFor: 'Detailansicht App: Freigaben live; App-Aktionen' },
  { key: 'flowState', name: 'Status eines Flows', api: 'Power-Automate-Admin-Schnittstelle (api.flow.microsoft.com)', official: false, permission: 'Microsoft Flow Service: User, Flows.Read.All', usedFor: 'Detailansicht Flow: Status live; Flow-Aktionen' },
  { key: 'credits', name: 'Copilot-Credits', api: 'Power Platform API – licensing', official: true, permission: 'Licensing.Allocations.Read', usedFor: 'Ansicht „Copilot-Credits“' },
  { key: 'graph', name: 'Benutzer lesen', api: 'Microsoft Graph – /users', official: true, permission: 'Microsoft Graph: User.Read.All', usedFor: '„Februar 2027“, Person per E-Mail' }
];

const STATE_TEXT: { [s in ProbeState]: string } = { idle: 'Nicht geprüft', running: 'Wird geprüft …', ok: 'Funktioniert', warn: 'Mit Hinweis', error: 'Geht nicht', skipped: 'Übersprungen' };
const STATE_ICON: { [s in ProbeState]: string } = { idle: 'Clock', running: 'Sync', ok: 'CompletedSolid', warn: 'Warning', error: 'ErrorBadge', skipped: 'Blocked2' };

function explain(e: unknown): string {
  if (e instanceof PpApiError) {
    if (e.kind === 'consent') {
      return 'Berechtigung für SharePoint fehlt (Skript ausführen). ' + (e.detail || '');
    }
    if (e.kind === 'forbidden') {
      return 'Zugriff verweigert – Rolle „Power Platform Administrator“ aktiv? ' + (e.detail || '');
    }
    if (e.kind === 'notFound') {
      return 'Nicht gefunden. ' + (e.detail || '');
    }
    return e.message;
  }
  const msg = e instanceof Error ? e.message : typeof e === 'object' && e !== null && 'message' in e ? String((e as { message: unknown }).message) : String(e);
  if (/failed to fetch|networkerror|load failed/i.test(msg)) {
    return 'Nicht erreichbar (Netzwerk oder Browser-Freigabe/CORS).';
  }
  if (/AADSTS65001|consent/i.test(msg)) {
    return 'Berechtigung für SharePoint fehlt (Skript ausführen). ' + msg;
  }
  return msg;
}

const sampleOf = (v: unknown): string => {
  try {
    const text = JSON.stringify(v, (_k, val) => (typeof val === 'string' && val.length > 300 ? val.slice(0, 300) + '…' : val), 2);
    return text.length > 4000 ? text.slice(0, 4000) + '\n…' : text;
  } catch {
    return String(v);
  }
};

const fmt = (n: number): string => n.toLocaleString('de-AT');

export const LiveDiagnostics: React.FC<ILiveDiagnosticsProps> = props => {
  const [results, setResults] = React.useState<{ [key: string]: IProbeResult }>({});
  const [running, setRunning] = React.useState<boolean>(false);
  const [openSample, setOpenSample] = React.useState<string>('');
  const [copied, setCopied] = React.useState<boolean>(false);
  const resultsRef = React.useRef<{ [key: string]: IProbeResult }>({});

  const set = (key: string, r: IProbeResult): void => {
    resultsRef.current = { ...resultsRef.current, [key]: r };
    setResults(resultsRef.current);
  };

  const probe = async (key: string, run: () => Promise<Omit<IProbeResult, 'ms'>>): Promise<void> => {
    const t0 = Date.now();
    set(key, { state: 'running', summary: '' });
    try {
      const r = await run();
      set(key, { ...r, ms: Date.now() - t0 });
    } catch (e) {
      set(key, { state: 'error', summary: explain(e), ms: Date.now() - t0 });
    }
  };

  const coeCount = async (logical: string, filter: string): Promise<number | undefined> => {
    if (!props.coe) {
      return undefined;
    }
    try {
      return (await props.coe.getCount(logical, filter)).value;
    } catch {
      return undefined;
    }
  };

  const runAll = async (): Promise<void> => {
    setRunning(true);
    resultsRef.current = {};
    setResults({});
    const { live, ppApi, coe, role } = props;

    await probe('role', async () => {
      if (!role) {
        return { state: 'warn', summary: 'Noch nicht geprüft.' };
      }
      const reader = role.roles.some(r => /Globaler Leser|Dynamics 365 Administrator/.test(r));
      return role.state === 'active'
        ? { state: 'ok', summary: `Aktiv: ${role.roles.join(', ')}` }
        : reader ? { state: 'warn', summary: `Nur ${role.roles.join(', ')} – Inventar lesbar, Aktionen nicht.` }
          : { state: role.state === 'inactive' ? 'error' : 'warn', summary: role.state === 'inactive' ? 'Rolle „Power Platform Administrator“ nicht aktiv – in PIM aktivieren, dann neu prüfen.' : `Nicht prüfbar: ${role.problem || ''}` };
    });

    const coeP = probe('coe', async () => {
      if (!coe) {
        return { state: 'skipped', summary: 'Keine CoE-Umgebung eingetragen.' };
      }
      const who = await coe.whoAmI();
      return { state: 'ok', summary: `Verbunden (${coe.baseUrl.replace(/^https:\/\//, '')})`, sample: sampleOf(who) };
    });

    let firstEnvId = '';
    const envP = probe('envs', async () => {
      const envs = await live.environments();
      firstEnvId = envs.length ? envs[0].id : '';
      const inCoe = await coeCount('admin_environment', 'statecode eq 0 and admin_environmentdeleted ne true');
      return {
        state: 'ok',
        summary: `${fmt(envs.length)} Environments live${inCoe !== undefined ? ` · CoE: ${fmt(inCoe)}` : ''}`,
        sample: sampleOf(envs.slice(0, 2).map(e => e.raw))
      };
    }).then(() => probe('envDetail', async () => {
      if (!firstEnvId) {
        return { state: 'skipped', summary: 'Kein Environment aus der Liste.' };
      }
      const e = await live.environment(firstEnvId);
      return { state: 'ok', summary: `${e.displayName}: Zustand ${e.state || '–'}, Schutzstufe ${e.protectionLevel || '–'}, Typ ${e.type || '–'}`, sample: sampleOf(e.raw) };
    }));

    let flowSample: IInvItem | undefined;
    let appSample: IInvItem | undefined;
    const invP = probe('invCount', async () => {
      const counts = await live.countByType();
      if (!counts.length) {
        return { state: 'warn', summary: 'Abfrage klappt, aber das Inventar ist leer.' };
      }
      const [apps, flows] = await Promise.all([
        coeCount('admin_app', 'statecode eq 0 and admin_appdeleted ne true'),
        coeCount('admin_flow', 'statecode eq 0 and admin_flowdeleted ne true')
      ]);
      const sum = (types: string[]): number => counts.filter(c => types.indexOf(c.type) >= 0).reduce((s, c) => s + c.count, 0);
      const cmp = `Vergleich: Apps live ${fmt(sum(APP_TYPES))}${apps !== undefined ? ` / CoE ${fmt(apps)}` : ''}, Flows live ${fmt(sum(FLOW_TYPES))}${flows !== undefined ? ` / CoE ${fmt(flows)}` : ''}`;
      return {
        state: 'ok',
        summary: counts.sort((a, b) => b.count - a.count).map(c => `${TYPE_LABEL[c.type] || c.type}: ${fmt(c.count)}`).join(' · ') + ` — ${cmp}`,
        sample: sampleOf(counts)
      };
    }).then(async () => {
      const sampleProbe = (key: string, types: string[], keep?: (it: IInvItem) => void): Promise<void> => probe(key, async () => {
        const res = await live.query([
          { $type: 'where', FieldName: 'type', Operator: 'in~', Values: types.map(t => `'${t}'`) },
          { $type: 'take', TakeCount: 3 }
        ], { Top: 3 });
        const rows = (Array.isArray(res.data) ? res.data : []) as { [k: string]: unknown }[];
        if (!rows.length) {
          return { state: 'warn', summary: 'Keine Einträge dieses Typs im Inventar.', sample: sampleOf(res) };
        }
        if (keep) {
          const items = await Promise.all(rows.map(r => live.item(String(r.name || ''), types).catch(() => undefined)));
          items.filter(x => !!x).forEach(x => keep(x!));
        }
        return { state: 'ok', summary: `${fmt(typeof res.totalRecords === 'number' ? res.totalRecords : rows.length)} Einträge · Stichprobe: ${rows.map(r => String((r.properties as { displayName?: string } || {}).displayName || r.name)).join(', ')}`, sample: sampleOf(rows[0]) };
      });
      await sampleProbe('invApps', APP_TYPES, it => { if (!appSample && it.type === INV_TYPE.canvasApp) { appSample = it; } });
      await sampleProbe('invFlows', FLOW_TYPES, it => { if (!flowSample || (!flowSample.workflowEntityId && it.workflowEntityId)) { flowSample = it; } });
      await sampleProbe('invAgents', [INV_TYPE.agent]);
      await sampleProbe('invPages', [INV_TYPE.pages]);
    });

    await Promise.all([coeP, envP, invP]);

    await probe('runs', async () => {
      if (!flowSample || !flowSample.workflowEntityId || !flowSample.environmentId) {
        return { state: 'skipped', summary: 'Kein Flow mit Workflow-ID aus dem Inventar (erst „Inventar – Flows“ muss funktionieren).' };
      }
      const runs = await live.flowRuns(flowSample.environmentId, flowSample.workflowEntityId, 5);
      return { state: 'ok', summary: `Flow „${flowSample.displayName}“: ${runs.length} Ausführungen gelesen`, sample: sampleOf(runs) };
    });
    await probe('appShares', async () => {
      if (!appSample || !appSample.environmentId) {
        return { state: 'skipped', summary: 'Keine Canvas-App aus dem Inventar.' };
      }
      const shares = await live.appShares(appSample.environmentId, appSample.id);
      return { state: 'ok', summary: `App „${appSample.displayName}“: ${shares.length} Freigaben`, sample: sampleOf(shares.slice(0, 3)) };
    });
    await probe('flowState', async () => {
      if (!flowSample || !flowSample.environmentId) {
        return { state: 'skipped', summary: 'Kein Flow aus dem Inventar.' };
      }
      const st = await live.flowState(flowSample.environmentId, flowSample.id);
      return { state: 'ok', summary: `Flow „${st.displayName || flowSample.displayName}“: Status ${st.state || '–'}`, sample: sampleOf(st) };
    });
    await probe('credits', async () => {
      const r = await ppApi.get<unknown>('/licensing/entitlements/MCSMessages', PPAPI_VERSIONS.licensing);
      return { state: 'ok', summary: 'Lizenzdaten lesbar', sample: sampleOf(r) };
    });
    await probe('graph', async () => {
      const graph = await props.graphFactory.getClient('3');
      const res = await graph.api('/users?$top=1&$select=id').version('v1.0').get() as { value?: unknown[] };
      return { state: 'ok', summary: `Benutzer lesbar (${(res.value || []).length} Stichprobe)` };
    });
    setRunning(false);
  };

  const counts = PROBES.reduce((acc, p) => {
    const st = results[p.key] ? results[p.key].state : 'idle';
    acc[st] = (acc[st] || 0) + 1;
    return acc;
  }, {} as { [s: string]: number });

  const report = (): string => {
    const lines = [`CoE Cockpit – Live-Schnittstellen, geprüft ${new Date().toLocaleString('de-AT')}`, ''];
    PROBES.forEach(p => {
      const r = results[p.key];
      lines.push(`[${r ? STATE_TEXT[r.state] : 'Nicht geprüft'}] ${p.name} (${p.official ? 'offiziell' : 'inoffiziell'}; ${p.permission})`);
      if (r && r.summary) {
        lines.push(`   ${r.summary}`);
      }
      if (r && r.sample && r.state !== 'ok') {
        lines.push('   ' + r.sample.split('\n').slice(0, 20).join('\n   '));
      }
    });
    return lines.join('\n');
  };

  return (
    <div className={styles.diag} style={props.themeVars}>
      <div className={styles.diagIntro}>
        <div>
          <b>Was geht in eurem Tenant live?</b> Diese Seite ruft jede Schnittstelle, die das Cockpit live nutzen kann, einmal mit deinen Rechten auf
          und zeigt das Ergebnis. Es wird nichts geändert. Ohne aktive Rolle „Power Platform Administrator“ (oder Globaler Leser für das Inventar)
          schlagen die meisten Prüfungen fehl.
        </div>
        <div className={styles.diagButtons}>
          <PrimaryButton text={running ? 'Wird geprüft …' : 'Alle prüfen'} iconProps={{ iconName: 'Play' }} disabled={running} onClick={() => { runAll().catch(() => setRunning(false)); }} />
          <DefaultButton
            text={copied ? 'Kopiert' : 'Ergebnis kopieren'}
            iconProps={{ iconName: copied ? 'CheckMark' : 'Copy' }}
            disabled={running || !Object.keys(results).length}
            onClick={() => copyText(report()).then(ok => { if (ok) { setCopied(true); window.setTimeout(() => setCopied(false), 2000); } })}
          />
        </div>
      </div>

      {Object.keys(results).length > 0 && (
        <div className={styles.diagSummary} role="status">
          {(['ok', 'warn', 'error', 'skipped'] as ProbeState[]).filter(s => counts[s]).map(s => (
            <span key={s} className={`${styles.diagState} ${s === 'ok' ? styles.diagOk : s === 'warn' ? styles.diagWarn : s === 'error' ? styles.diagBad : styles.diagIdle}`}>
              <Icon iconName={STATE_ICON[s]} /> {counts[s]} {STATE_TEXT[s]}
            </span>
          ))}
        </div>
      )}

      <table className={styles.diagTable}>
        <thead>
          <tr><th>Schnittstelle</th><th>Ergebnis</th><th>Berechtigung</th><th>Wofür</th></tr>
        </thead>
        <tbody>
          {PROBES.map(p => {
            const r = results[p.key];
            const st: ProbeState = r ? r.state : 'idle';
            return (
              <tr key={p.key} data-probe={p.key}>
                <td>
                  <div className={styles.diagName}>
                    {p.name}
                    <span className={`${styles.diagOfficial} ${p.official ? '' : styles.diagUnofficial}`}>{p.official ? 'offiziell' : 'inoffiziell'}</span>
                  </div>
                  <div className={styles.diagSub}>{p.api}</div>
                </td>
                <td>
                  <span className={`${styles.diagState} ${st === 'ok' ? styles.diagOk : st === 'warn' ? styles.diagWarn : st === 'error' ? styles.diagBad : styles.diagIdle}`}>
                    <Icon iconName={STATE_ICON[st]} /> {STATE_TEXT[st]}
                  </span>
                  {r && r.ms !== undefined && <span className={styles.diagSub}> {r.ms} ms</span>}
                  {r && r.summary && <div className={styles.diagSub}>{r.summary}</div>}
                  {r && r.sample && (
                    <>
                      <button type="button" className={styles.textButton} onClick={() => setOpenSample(openSample === p.key ? '' : p.key)}>
                        {openSample === p.key ? 'Beispiel ausblenden' : 'Beispiel anzeigen'}
                      </button>
                      {openSample === p.key && <pre className={styles.diagSample}>{r.sample}</pre>}
                    </>
                  )}
                </td>
                <td className={styles.diagSub}>{p.permission}</td>
                <td className={styles.diagSub}>{p.usedFor}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
};
