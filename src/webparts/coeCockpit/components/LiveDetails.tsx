import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Shimmer } from '@fluentui/react/lib/Shimmer';
import styles from './CoeCockpit.module.scss';
import { IRecord } from '../models/ITableConfig';
import { LIVE_FIELD, LIVE_NOTE_FIELD, LIVE_STATUS } from '../models/liveData';
import { APP_TYPES, FLOW_TYPES, IAppShare, IFlowRun, IInvItem, ILiveEnvironment, INV_TYPE, LiveInventoryService, TYPE_LABEL } from '../services/LiveInventoryService';
import { PpApiError } from '../services/PowerPlatformApiService';
import { displayValue, rawValue } from '../utils/format';

/**
 * Abschnitt „Live“ in der Detailansicht (Apps, Flows, Environments): aktueller Stand direkt von Microsoft –
 * zusätzlich zu den CoE-Daten darüber. Jeder Teil lädt für sich; was nicht geht, zeigt einen kurzen Grund.
 */
export interface ILiveDetailsProps {
  logicalName: string;
  record: IRecord;
  live: LiveInventoryService;
  /** Inoffizielle Admin-Schnittstellen verwenden (Freigaben, Flow-Status) */
  allowAdminApis: boolean;
}

type Part<T> = { status: 'loading' } | { status: 'ok'; value: T } | { status: 'error'; message: string };

const dtFmt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const fmtDate = (v: string): string => {
  const d = Date.parse(v);
  return isNaN(d) ? (v || '–') : dtFmt.format(new Date(d));
};
const reason = (e: unknown): string => {
  if (e instanceof PpApiError) {
    return e.kind === 'consent' ? 'Berechtigung für SharePoint fehlt'
      : e.kind === 'forbidden' ? 'Zugriff verweigert (Rolle nicht aktiv?)'
        : e.kind === 'notFound' ? 'nicht gefunden' : e.message;
  }
  return e instanceof Error ? e.message : String(e);
};
const RUN_LABEL: { [s: string]: string } = { succeeded: 'Erfolgreich', failed: 'Fehlgeschlagen', running: 'Läuft', cancelled: 'Abgebrochen', canceled: 'Abgebrochen', skipped: 'Übersprungen', aborted: 'Abgebrochen' };
const ROLE_LABEL: { [r: string]: string } = { owner: 'Besitzer', canedit: 'Mitbesitzer', canview: 'Benutzer', canviewwithshare: 'Benutzer (darf teilen)' };
const STATE_LABEL: { [s: string]: string } = { started: 'Aktiv', stopped: 'Angehalten', suspended: 'Gesperrt', enabled: 'Aktiviert', disabled: 'Deaktiviert', ready: 'Bereit' };

function usePart<T>(load: (() => Promise<T>) | undefined, deps: unknown[]): Part<T> | undefined {
  const [part, setPart] = React.useState<Part<T> | undefined>(undefined);
  React.useEffect(() => {
    if (!load) {
      setPart(undefined);
      return;
    }
    let cancelled = false;
    setPart({ status: 'loading' });
    load().then(
      value => { if (!cancelled) { setPart({ status: 'ok', value }); } },
      e => { if (!cancelled) { setPart({ status: 'error', message: reason(e) }); } }
    );
    return () => { cancelled = true; };
  }, deps);
  return part;
}

export const LiveDetails: React.FC<ILiveDetailsProps> = props => {
  const { logicalName, record, live } = props;
  const isApp = logicalName === 'admin_app';
  const isFlow = logicalName === 'admin_flow';
  const isEnv = logicalName === 'admin_environment';
  const id = (isApp ? rawValue(record, 'admin_appid') : isFlow ? (rawValue(record, 'admin_flowidstring') || rawValue(record, 'admin_flowid'))
    : (rawValue(record, 'admin_environmentname') || rawValue(record, 'admin_environmentid'))).toLowerCase();
  const envId = (isApp ? rawValue(record, 'admin_appenvironmentid') : isFlow ? rawValue(record, 'admin_flowenvironmentid') : id).toLowerCase();

  const item = usePart<IInvItem | undefined>(!isEnv && id ? () => live.item(id, isApp ? APP_TYPES : FLOW_TYPES) : undefined, [id]);
  const env = usePart<ILiveEnvironment>(isEnv && id ? () => live.environment(id) : undefined, [id]);
  const envCounts = usePart<{ type: string; count: number }[]>(isEnv && id ? () => live.countInEnvironment(id) : undefined, [id]);
  const shares = usePart<IAppShare[]>(isApp && props.allowAdminApis && id && envId ? () => live.appShares(envId, id) : undefined, [id, props.allowAdminApis]);
  const flowState = usePart<{ state: string }>(isFlow && props.allowAdminApis && id && envId ? () => live.flowState(envId, id) : undefined, [id, props.allowAdminApis]);
  const wfId = item && item.status === 'ok' && item.value ? item.value.workflowEntityId || '' : '';
  const runs = usePart<IFlowRun[]>(isFlow && envId && wfId ? () => live.flowRuns(envId, wfId, 10) : undefined, [envId, wfId]);

  const fact = (label: string, value: React.ReactNode): JSX.Element => (
    <div className={styles.liveFact}><span className={styles.liveFactLabel}>{label}</span><span className={styles.liveFactValue}>{value}</span></div>
  );
  const loadingRow = <div className={styles.solLoading}><Shimmer /><Shimmer width="70%" /></div>;
  const errorRow = (what: string, message: string): JSX.Element => (
    <div className={styles.solEmpty}><Icon iconName="Info" /><span>{what}: {message}</span></div>
  );
  const status = String(record[LIVE_FIELD] || '');

  return (
    <div className={styles.solBox}>
      <div className={styles.liveHead}>
        <span className={styles.livePill}><Icon iconName="LightningBolt" /> Live von Microsoft</span>
        {status && <span className={styles.liveHeadNote}>{status === LIVE_STATUS.added ? 'Nur live – noch nicht im CoE' : String(record[LIVE_NOTE_FIELD] || status)}</span>}
      </div>

      {/* Apps und Flows: Inventar (offiziell) */}
      {!isEnv && item && item.status === 'loading' && loadingRow}
      {!isEnv && item && item.status === 'error' && errorRow('Inventar', item.message)}
      {!isEnv && item && item.status === 'ok' && !item.value && errorRow('Inventar', 'nicht gefunden – vermutlich gelöscht oder noch nicht erfasst (bis 15 Minuten Verzögerung)')}
      {!isEnv && item && item.status === 'ok' && item.value && (
        <div className={styles.liveFacts}>
          {fact('Name', item.value.displayName || '–')}
          {fact('Typ', TYPE_LABEL[item.value.type] || item.value.type)}
          {fact('Zuletzt geändert', fmtDate(item.value.lastModifiedAt))}
          {isApp && fact('Quarantäne', item.value.isQuarantined === undefined ? '–' : item.value.isQuarantined ? 'Ja' : 'Nein')}
          {isFlow && fact('Status', flowState ? (flowState.status === 'ok' ? (STATE_LABEL[flowState.value.state.toLowerCase()] || flowState.value.state || '–')
            : flowState.status === 'loading' ? '…' : `nicht lesbar (${flowState.message})`) : displayValue(record, 'admin_flowstate') + ' (CoE)')}
          {isFlow && item.value.trigger && fact('Auslöser', `${item.value.trigger}${item.value.triggerOperation ? ` – ${item.value.triggerOperation}` : ''}`)}
          {fact('Connectors', item.value.connectors && item.value.connectors.length
            ? item.value.connectors.map(c => `${c.id.replace(/^shared_/, '')}${c.operations.length ? ` (${c.operations.length})` : ''}`).join(', ')
            : 'keine')}
        </div>
      )}

      {/* Flow-Ausführungen (offiziell) */}
      {isFlow && runs && (
        <div className={styles.liveSection}>
          <div className={styles.liveSectionTitle}><Icon iconName="History" /> Letzte Ausführungen</div>
          {runs.status === 'loading' && loadingRow}
          {runs.status === 'error' && errorRow('Ausführungen', runs.message)}
          {runs.status === 'ok' && !runs.value.length && <div className={styles.solEmpty}><Icon iconName="Info" /><span>Keine Ausführungen in den letzten 28 Tagen.</span></div>}
          {runs.status === 'ok' && runs.value.length > 0 && (
            <ul className={styles.solList}>
              {runs.value.map(r => {
                const st = r.status.toLowerCase();
                return (
                  <li key={r.name}>
                    <div className={styles.solItem}>
                      <span className={`${styles.solIcon} ${st === 'failed' ? styles.liveIconBad : st === 'succeeded' ? styles.liveIconOk : ''}`}>
                        <Icon iconName={st === 'failed' ? 'ErrorBadge' : st === 'succeeded' ? 'CompletedSolid' : 'Clock'} />
                      </span>
                      <span className={styles.solText}>
                        <span className={styles.solName}>{RUN_LABEL[st] || r.status}</span>
                        <span className={styles.solSub}>{fmtDate(r.startTime)}{r.durationMs !== undefined ? ` · ${Math.max(1, Math.round(r.durationMs / 1000))} s` : ''}{r.triggerType ? ` · ${r.triggerType}` : ''}</span>
                      </span>
                    </div>
                  </li>
                );
              })}
            </ul>
          )}
        </div>
      )}
      {isFlow && item && item.status === 'ok' && item.value && !wfId && (
        <div className={styles.solEmpty}><Icon iconName="Info" /><span>Ausführungen: Das Inventar nennt für diesen Flow keine Dataverse-Workflow-ID.</span></div>
      )}

      {/* Freigaben einer App (Admin-Schnittstelle) */}
      {isApp && shares && (
        <div className={styles.liveSection}>
          <div className={styles.liveSectionTitle}><Icon iconName="People" /> Freigaben (live)</div>
          {shares.status === 'loading' && loadingRow}
          {shares.status === 'error' && errorRow('Freigaben', shares.message)}
          {shares.status === 'ok' && (
            <ul className={styles.solList}>
              {shares.value.length === 0 && <li className="muted">Keine Freigaben.</li>}
              {shares.value.map(s => (
                <li key={s.principalId + s.role}>
                  <div className={styles.solItem}>
                    <span className={styles.solIcon}><Icon iconName={/group/i.test(s.type) ? 'Group' : /tenant/i.test(s.type) ? 'Globe' : 'Contact'} /></span>
                    <span className={styles.solText}>
                      <span className={styles.solName}>{s.name || s.email || (/tenant/i.test(s.type) ? 'Gesamte Organisation' : s.principalId)}</span>
                      <span className={styles.solSub}>{ROLE_LABEL[s.role.toLowerCase()] || s.role}{s.email ? ` · ${s.email}` : ''}</span>
                    </span>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}

      {/* Environment (offiziell) */}
      {isEnv && env && env.status === 'loading' && loadingRow}
      {isEnv && env && env.status === 'error' && errorRow('Environment', env.message)}
      {isEnv && env && env.status === 'ok' && (
        <div className={styles.liveFacts}>
          {fact('Name', env.value.displayName || '–')}
          {fact('Typ', env.value.type || '–')}
          {fact('Zustand', STATE_LABEL[env.value.state.toLowerCase()] || env.value.state || '–')}
          {fact('Managed', env.value.protectionLevel ? (/^basic$/i.test(env.value.protectionLevel) ? 'Nein' : `Ja (${env.value.protectionLevel})`) : '–')}
          {fact('Region', env.value.region || '–')}
          {fact('Erstellt', fmtDate(env.value.createdDateTime))}
        </div>
      )}
      {isEnv && envCounts && envCounts.status === 'ok' && envCounts.value.length > 0 && (
        <div className={styles.liveCounts}>
          {envCounts.value.filter(c => c.type !== INV_TYPE.environment).sort((a, b) => b.count - a.count).map(c => (
            <span key={c.type} className={styles.liveCount}><b>{c.count.toLocaleString('de-AT')}</b> {TYPE_LABEL[c.type] || c.type}</span>
          ))}
        </div>
      )}
      {isEnv && envCounts && envCounts.status === 'error' && errorRow('Inventar', envCounts.message)}

      <div className={styles.solFoot}>
        <span>Quelle: Power Platform API{(isApp && props.allowAdminApis) || (isFlow && props.allowAdminApis) ? ' und Admin-Schnittstellen' : ''} – im Namen der angemeldeten Person.</span>
      </div>
    </div>
  );
};
