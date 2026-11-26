import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord, ITableConfig } from '../models/ITableConfig';
import { DataverseService, ICountResult } from '../services/DataverseService';
import { IColumnFilter } from '../utils/filters';
import { formatNumber, formatTime } from '../utils/format';
import { translateValue } from '../utils/valueLabels';
import {
  count, daysAgoIso, distribution, environmentsOverview, IMonth, lastMonths, monthlyCounts, numberOf, textOf, topAppsByLaunches, topRecords
} from '../models/overview';
import { DEV_ENV_DEADLINES } from '../models/devEnvironments';
import { BarList, compactNumber, IBarItem, IPart, SegmentedBar, Sparkline, TrendChart } from './Charts';

export interface IOverviewProps {
  service: DataverseService;
  tables: ITableConfig[];
  /** Lädt neu, wenn sich der Wert ändert (Aktualisieren-Schaltfläche). */
  reloadKey: string;
  /** Betroffene Developer-Environments (aus dem Reiter „Betroffen ab 15.12.“), undefined = lädt. */
  devAffected: { count: number; empty: number; disabled: number } | undefined;
  showDevEnvironments: boolean;
  /** Öffnet die Ansicht „Sicherheit & Freigaben“ (falls eingeschaltet). */
  onOpenSecurity?: () => void;
  onNavigate: (tableKey: string, filters?: { [field: string]: IColumnFilter }) => void;
  onOpenRecord: (table: ITableConfig, record: IRecord, primaryId: string) => void;
}

interface ISection<T> {
  status: 'loading' | 'ready' | 'error';
  data?: T;
  error?: string;
}

const loading = <T,>(): ISection<T> => ({ status: 'loading' });

/** Zwischenspeicher der Übersicht (pro Browser-Sitzung, 10 Minuten) – „Aktualisieren“ lädt immer neu. */
const CACHE_TTL = 10 * 60 * 1000;
interface ICache { time: number; reloadKey: string; data: { [section: string]: unknown } }
function readCache(key: string): ICache | undefined {
  try {
    const raw = window.sessionStorage.getItem(key);
    return raw ? JSON.parse(raw) as ICache : undefined;
  } catch {
    return undefined;
  }
}
function writeCache(key: string, value: ICache): void {
  try {
    window.sessionStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Speicher voll oder gesperrt – dann eben ohne Zwischenspeicher
  }
}

/** Feste Farbzuordnung je Kategorie – die Farbe folgt dem Wert, nicht dem Rang. */
const APP_TYPE_SLOT: { [k: string]: number } = { 'Canvas': 1, 'Model Driven': 2, 'Custom Page': 3, 'SharePointFormApp': 4, 'Component Library': 5 };
const ENV_TYPE_SLOT: { [k: string]: number } = { 'Production': 1, 'Sandbox': 2, 'Developer': 3, 'Default': 4, 'Trial': 5, 'Teams': 6 };
const FLOW_STATE: { [k: string]: { color: string; label: string; darkText?: boolean } } = {
  'Started': { color: 'var(--coe-status-good)', label: 'Aktiv' },
  'Stopped': { color: 'var(--coe-status-neutral)', label: 'Angehalten' },
  'Suspended': { color: 'var(--coe-status-critical)', label: 'Gesperrt' }
};
/** Anzeigewert „Ja“ (CoE speichert „Yes“, das Webpart zeigt deutsch an). */
const YES = ['Ja'];

function slotColor(slot: number | undefined): string {
  return slot ? `var(--coe-series-${slot})` : 'var(--coe-status-neutral)';
}

/** Höchstens 6 Teile: Rest wird zu „Andere“. */
function toParts(
  dist: { label: string; value: number }[], color: (label: string) => string, onClick: (label: string) => void,
  display: (label: string) => string = l => l, max: number = 6
): IPart[] {
  const head = dist.slice(0, dist.length > max ? max - 1 : max);
  const rest = dist.slice(head.length);
  const parts: IPart[] = head.map(d => ({
    key: d.label, label: display(d.label), value: d.value, color: color(d.label),
    darkText: color(d.label) === 'var(--coe-series-4)', onClick: () => onClick(d.label)
  }));
  if (rest.length) {
    parts.push({ key: '_other', label: `Andere (${rest.length})`, value: rest.reduce((s, r) => s + r.value, 0), color: 'var(--coe-status-neutral)' });
  }
  return parts;
}

export const Card: React.FC<{ title: string; subtitle?: string; icon: string; action?: { label: string; onClick: () => void }; className?: string }> = props => (
  <section className={`${styles.ovCard} ${props.className || ''}`}>
    <header className={styles.ovCardHead}>
      <span className={styles.ovCardIcon}><Icon iconName={props.icon} /></span>
      <div style={{ minWidth: 0, flex: '1 1 auto' }}>
        <h3 className={styles.ovCardTitle}>{props.title}</h3>
        {props.subtitle && <div className={styles.ovCardSub}>{props.subtitle}</div>}
      </div>
      {props.action && (
        <button type="button" className={styles.ovCardAction} onClick={props.action.onClick}>
          {props.action.label} <Icon iconName="ChevronRight" />
        </button>
      )}
    </header>
    <div className={styles.ovCardBody}>{props.children}</div>
  </section>
);

export const Skeleton: React.FC<{ rows?: number; height?: number }> = ({ rows = 3, height = 14 }) => (
  <div className={styles.ovSkeleton}>
    {Array.from({ length: rows }, (_x, i) => <span key={i} className={styles.skeleton} style={{ height, width: `${90 - i * 12}%` }} />)}
  </div>
);

export const Failed: React.FC<{ text?: string }> = ({ text }) => (
  <div className={styles.ovFailed}><Icon iconName="Info" /> Nicht verfügbar{text ? ` – ${text}` : ''}</div>
);

export const OverviewView: React.FC<IOverviewProps> = props => {
  const { service, tables, onNavigate, onOpenRecord } = props;
  // 12 volle Monate + der laufende Monat (unvollständig, im Diagramm gestrichelt)
  const months = React.useMemo<IMonth[]>(() => lastMonths(13), [props.reloadKey]);
  const byLogical = (logical: string): ITableConfig | undefined => tables.filter(t => t.logicalName === logical)[0];
  const T = {
    apps: byLogical('admin_app'), flows: byLogical('admin_flow'), envs: byLogical('admin_environment'), makers: byLogical('admin_maker'),
    bots: byLogical('admin_pva'), connectors: byLogical('admin_connector'), rpa: byLogical('admin_rpa'), pages: byLogical('admin_portal')
  };

  const [counts, setCounts] = React.useState<{ [k: string]: ISection<ICountResult> }>({});
  const [trend, setTrend] = React.useState<ISection<{ apps: number[]; flows: number[] }>>(loading());
  const [appTypes, setAppTypes] = React.useState<ISection<{ label: string; value: number }[]>>(loading());
  const [flowStates, setFlowStates] = React.useState<ISection<{ label: string; value: number }[]>>(loading());
  const [envs, setEnvs] = React.useState<ISection<{ records: IRecord[]; primaryId: string; byType: { label: string; value: number }[] }>>(loading());
  const [topApps, setTopApps] = React.useState<ISection<{ records: IRecord[]; primaryId: string; valueField: string; source: 'field' | 'audit'; days: number }>>(loading());
  const [topConnectors, setTopConnectors] = React.useState<ISection<{ records: IRecord[]; primaryId: string }>>(loading());
  const [topMakers, setTopMakers] = React.useState<ISection<{ records: IRecord[]; primaryId: string }>>(loading());
  const [gov, setGov] = React.useState<{ [k: string]: ISection<ICountResult> }>({});
  const [loadedAt, setLoadedAt] = React.useState<{ time: number; cached: boolean } | undefined>(undefined);
  const cacheKey = `coe-cockpit:overview:${service.baseUrl}`;

  React.useEffect(() => {
    let cancelled = false;
    // Aus dem Zwischenspeicher, wenn er zu diesem Stand passt und jünger als 10 Minuten ist
    const cache = readCache(cacheKey);
    if (cache && cache.reloadKey === props.reloadKey && Date.now() - cache.time < CACHE_TTL) {
      const d = cache.data as { [k: string]: never };
      setCounts(d.counts || {}); setTrend(d.trend); setAppTypes(d.appTypes); setFlowStates(d.flowStates); setEnvs(d.envs);
      setTopApps(d.topApps); setTopConnectors(d.topConnectors); setTopMakers(d.topMakers); setGov(d.gov || {});
      setLoadedAt({ time: cache.time, cached: true });
      return () => { cancelled = true; };
    }
    setLoadedAt(undefined);
    const guard = <T,>(set: (s: ISection<T>) => void, p: Promise<T>): void => {
      set({ status: 'loading' });
      p.then(data => { if (!cancelled) { set({ status: 'ready', data }); } }, (e: Error) => { if (!cancelled) { set({ status: 'error', error: e.message }); } });
    };
    // Kennzahlen
    const countKeys: { [k: string]: ITableConfig | undefined } = T;
    Object.keys(countKeys).forEach(k => {
      const t = countKeys[k];
      if (t) {
        guard<ICountResult>(s => setCounts(prev => ({ ...prev, [k]: s })), count(service, t));
      }
    });
    // Verlauf
    if (T.apps && T.flows) {
      guard(setTrend, Promise.all([
        monthlyCounts(service, T.apps, 'admin_appcreatedon', months),
        monthlyCounts(service, T.flows, 'admin_flowcreatedon', months)
      ]).then(([apps, flows]) => ({ apps, flows })));
    } else {
      setTrend({ status: 'error' });
    }
    // Verteilungen
    // App-Typ (Auswahlfeld des CoE): Anzeigetexte als Rückfall, falls Dataverse sie bei Gruppierungen nicht mitliefert
    const appTypeLabels = { '597910000': 'Canvas', '597910001': 'Model Driven', '597910002': 'SharePointFormApp', '597910003': 'Custom Page', '597910004': 'Component Library' };
    if (T.apps) { guard(setAppTypes, distribution(service, T.apps, 'admin_powerappstype', appTypeLabels)); } else { setAppTypes({ status: 'error' }); }
    if (T.flows) { guard(setFlowStates, distribution(service, T.flows, 'admin_flowstate')); } else { setFlowStates({ status: 'error' }); }
    if (T.envs) { guard(setEnvs, environmentsOverview(service, T.envs)); } else { setEnvs({ status: 'error' }); }
    // Ranglisten
    if (T.apps) { guard(setTopApps, topAppsByLaunches(service, T.apps, 8)); } else { setTopApps({ status: 'error' }); }
    if (T.connectors) { guard(setTopConnectors, topRecords(service, T.connectors, 'admin_numapps', 8, ['admin_numapps', 'admin_numflows'])); } else { setTopConnectors({ status: 'error' }); }
    if (T.makers) { guard(setTopMakers, topRecords(service, T.makers, 'admin_numberofapps', 8, ['admin_numberofapps', 'admin_numberofflows'])); } else { setTopMakers({ status: 'error' }); }
    // Handlungsbedarf
    const govQueries: { [k: string]: [ITableConfig | undefined, string] } = {
      orphanApps: [T.apps, "cr5d5_appisorphaned eq 'Yes'"],
      orphanFlows: [T.flows, "cr5d5_flowisorphaned eq 'Yes'"],
      suspended: [T.flows, "admin_flowstate eq 'Suspended'"],
      unused: [T.apps, `(admin_applastlaunchedon lt ${daysAgoIso(90)} or admin_applastlaunchedon eq null)`],
      guests: [T.makers, 'admin_userisguest eq true']
    };
    Object.keys(govQueries).forEach(k => {
      const [t, f] = govQueries[k];
      if (t) {
        guard<ICountResult>(s => setGov(prev => ({ ...prev, [k]: s })), count(service, t, f));
      }
    });
    return () => { cancelled = true; };
  }, [service, props.reloadKey, tables.map(t => t.key + t.logicalName).join('|')]);

  // Fertig geladen? Dann Zeitpunkt merken und zwischenspeichern
  const sectionsDone = [trend, appTypes, flowStates, envs, topApps, topConnectors, topMakers].every(x => x.status !== 'loading')
    && Object.keys(counts).every(k => counts[k].status !== 'loading') && Object.keys(gov).every(k => gov[k].status !== 'loading')
    && Object.keys(counts).length > 0;
  React.useEffect(() => {
    if (!sectionsDone || (loadedAt && loadedAt.cached)) {
      return;
    }
    const time = Date.now();
    setLoadedAt({ time, cached: false });
    writeCache(cacheKey, { time, reloadKey: props.reloadKey, data: { counts, trend, appTypes, flowStates, envs, topApps, topConnectors, topMakers, gov } });
  }, [sectionsDone]);

  const value = (k: string): ICountResult | undefined => (counts[k] && counts[k].status === 'ready' ? counts[k].data : undefined);
  const resources = ['apps', 'flows', 'bots', 'rpa', 'pages'].map(value);
  const resourcesReady = resources.filter(r => !!r).length;
  const total = resources.reduce((s, r) => s + (r ? r.value : 0), 0);
  const anyAtLeast = resources.some(r => !!r && r.atLeast);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const daysToBlock = Math.round((DEV_ENV_DEADLINES.block.getTime() - today.getTime()) / 86400000);
  /** Letzter vollständiger Monat (für „neu im …“). */
  const monthName = months.length > 1 ? months[months.length - 2].long.split(' ')[0] : '';
  const complete = (vals: number[] | undefined): number[] | undefined => (vals ? vals.slice(0, vals.length - 1) : undefined);

  const nav = (t: ITableConfig | undefined, filters?: { [field: string]: IColumnFilter }): (() => void) | undefined =>
    (t ? () => onNavigate(t.key, filters) : undefined);
  const open = (t: ITableConfig | undefined, r: IRecord, pk: string): (() => void) | undefined =>
    (t ? () => onOpenRecord(t, r, pk) : undefined);

  // ---------- Kacheln ----------
  const devCount = envs.data ? (envs.data.byType.filter(t => t.label === 'Developer')[0] || { value: 0 }).value : undefined;
  const guestCount = gov.guests && gov.guests.data ? gov.guests.data.value : undefined;
  const tiles: { key: string; label: string; icon: string; table?: ITableConfig; trend?: number[]; hint?: string }[] = [
    { key: 'apps', label: 'Apps', icon: 'AppIconDefault', table: T.apps, trend: complete(trend.data ? trend.data.apps : undefined) },
    { key: 'flows', label: 'Flows', icon: 'Flow', table: T.flows, trend: complete(trend.data ? trend.data.flows : undefined) },
    { key: 'envs', label: 'Environments', icon: 'Globe', table: T.envs, hint: devCount !== undefined ? `davon ${formatNumber(devCount)} Developer` : undefined },
    { key: 'makers', label: 'Maker', icon: 'People', table: T.makers, hint: guestCount !== undefined ? `davon ${formatNumber(guestCount)} Gäste` : undefined },
    { key: 'bots', label: 'Bots', icon: 'Robot', table: T.bots },
    { key: 'connectors', label: 'Connectors', icon: 'PlugConnected', table: T.connectors }
  ];

  const govItems: { key: string; label: string; hint: string; icon: string; tone: string; onClick?: () => void; value?: number; atLeast?: boolean; loading: boolean; failed: boolean }[] = [];
  const govCount = (k: string): ISection<ICountResult> | undefined => gov[k];
  const pushGov = (k: string, label: string, hint: string, icon: string, tone: string, onClick?: () => void): void => {
    const s = govCount(k);
    govItems.push({
      key: k, label, hint, icon, tone, onClick,
      value: s && s.data ? s.data.value : undefined, atLeast: s && s.data ? s.data.atLeast : false,
      loading: !s || s.status === 'loading', failed: !!s && s.status === 'error'
    });
  };
  if (props.showDevEnvironments) {
    govItems.push({
      key: 'dev', label: 'Developer-Environments betroffen', hint: `Sperre am 15.12.2026${props.devAffected ? ` · ${props.devAffected.empty} leer` : ''}`,
      icon: 'Warning', tone: styles.ovToneCritical, onClick: () => onNavigate('devenv:affected'),
      value: props.devAffected ? props.devAffected.count : undefined, loading: !props.devAffected, failed: false
    });
  }
  pushGov('suspended', 'Gesperrte Flows', 'Status „Gesperrt“ (z. B. DLP, ungültige Verbindung)', 'Blocked2', styles.ovToneCritical, nav(T.flows, { admin_flowstate: { values: [translateValue('admin_flowstate', 'Suspended')] } }));
  pushGov('orphanApps', 'Verwaiste Apps', 'Besitzer nicht mehr im Unternehmen', 'UserWarning', styles.ovToneSerious, nav(T.apps, { cr5d5_appisorphaned: { values: YES } }));
  pushGov('orphanFlows', 'Verwaiste Flows', 'Besitzer nicht mehr im Unternehmen', 'UserWarning', styles.ovToneSerious, nav(T.flows, { cr5d5_flowisorphaned: { values: YES } }));
  pushGov('unused', 'Apps seit 90 Tagen nicht gestartet', 'Kandidaten zum Aufräumen', 'HourGlass', styles.ovToneWarning, nav(T.apps, { admin_applastlaunchedon: { date: 'older90' } }));
  pushGov('guests', 'Gast-Maker', 'Externe Konten, die Ressourcen erstellen', 'Contact', styles.ovToneInfo, nav(T.makers, { admin_userisguest: { values: ['Ja'] } }));

  return (
    <div className={styles.ov}>
      {/* ---------- Kopfbereich ---------- */}
      <section className={styles.ovHero}>
        <div className={styles.ovHeroGlow} aria-hidden="true" />
        <div className={styles.ovHeroMain}>
          <div className={styles.ovHeroEyebrow}>
            <Icon iconName="ViewDashboard" /> Power Platform · Übersicht des Mandanten
            {loadedAt && <span className={styles.ovHeroStamp}>Stand {formatTime(new Date(loadedAt.time))}{loadedAt.cached ? ' (zwischengespeichert)' : ''}</span>}
          </div>
          <div className={styles.ovHeroValue}>
            {resourcesReady === 0 ? <span className={`${styles.skeleton} ${styles.ovHeroSkeleton}`} /> : `${formatNumber(total)}${anyAtLeast ? '+' : ''}`}
          </div>
          <div className={styles.ovHeroLabel}>
            Ressourcen
            {value('envs') ? <> in <strong>{formatNumber(value('envs')!.value)}</strong> Environments</> : null}
            {value('makers') ? <> von <strong>{formatNumber(value('makers')!.value)}</strong> Makern</> : null}
          </div>
          <div className={styles.ovHeroChips}>
            {[['apps', 'Apps', T.apps], ['flows', 'Flows', T.flows], ['bots', 'Bots', T.bots], ['rpa', 'Desktop Flows', T.rpa], ['pages', 'Power Pages', T.pages]].map(([k, l, t]) => {
              const v = value(k as string);
              return (
                <button key={k as string} type="button" className={styles.ovHeroChip} onClick={nav(t as ITableConfig | undefined)} disabled={!t}>
                  <strong>{v ? compactNumber(v.value) : '…'}</strong> {l as string}
                </button>
              );
            })}
          </div>
        </div>
        {props.showDevEnvironments && (
          <button type="button" className={styles.ovCountdown} onClick={() => onNavigate('devenv:affected')}>
            <span className={styles.ovCountdownTop}><Icon iconName="Clock" /> Sperre nicht verwalteter Developer-Environments</span>
            <span className={styles.ovCountdownDays}>{daysToBlock > 0 ? formatNumber(daysToBlock) : '0'}</span>
            <span className={styles.ovCountdownUnit}>{daysToBlock === 1 ? 'Tag' : 'Tage'} bis 15.12.2026</span>
            <span className={styles.ovCountdownFoot}>
              {props.devAffected ? <><strong>{formatNumber(props.devAffected.count)}</strong> Environments betroffen</> : 'Betroffene werden ermittelt …'}
              <Icon iconName="ChevronRight" />
            </span>
          </button>
        )}
      </section>

      {/* ---------- Kennzahl-Kacheln ---------- */}
      <div className={styles.ovTiles}>
        {tiles.map(t => {
          const s = counts[t.key];
          const v = s && s.data;
          const newThisMonth = t.trend ? t.trend[t.trend.length - 1] : undefined;
          return (
            <button key={t.key} type="button" className={styles.ovTile} onClick={nav(t.table)} disabled={!t.table}>
              <span className={styles.ovTileTop}>
                <span className={styles.ovTileIcon}><Icon iconName={t.icon} /></span>
                <span className={styles.ovTileLabel}>{t.label}</span>
              </span>
              <span className={styles.ovTileValue}>
                {!s || s.status === 'loading' ? <span className={styles.skeleton} style={{ width: 72, height: 26 }} /> : v ? `${formatNumber(v.value)}${v.atLeast ? '+' : ''}` : '–'}
              </span>
              <span className={styles.ovTileFoot}>
                {newThisMonth !== undefined
                  ? <span className={styles.ovDelta}><Icon iconName="Up" /> {formatNumber(newThisMonth)} neu im {monthName}</span>
                  : <span className={styles.ovTileHint}>{t.hint || 'Alle anzeigen'}</span>}
              </span>
              {t.trend && <Sparkline values={t.trend} label={t.label} />}
            </button>
          );
        })}
      </div>

      {/* ---------- Verlauf + Handlungsbedarf ---------- */}
      <div className={styles.ovRow}>
        <Card className={styles.ovWide} title="Neu erstellt pro Monat" subtitle="Apps und Flows der letzten 12 Monate" icon="AreaChart">
          {trend.status === 'loading' && <Skeleton rows={5} height={18} />}
          {trend.status === 'error' && <Failed text={trend.error} />}
          {trend.status === 'ready' && trend.data && (
            <>
              <div className={styles.chartLegend}>
                <span><span className={styles.legendKey} style={{ background: 'var(--coe-series-1)' }} />Apps</span>
                <span><span className={styles.legendKey} style={{ background: 'var(--coe-series-2)' }} />Flows</span>
                <span><span className={styles.legendDash} />laufender Monat</span>
                <span className={styles.chartLegendNote}>
                  12 Monate: {formatNumber(complete(trend.data.apps)!.reduce((a, b) => a + b, 0))} Apps · {formatNumber(complete(trend.data.flows)!.reduce((a, b) => a + b, 0))} Flows
                </span>
              </div>
              <TrendChart
                height={300}
                partialLast={true}
                labels={months.map(m => m.short)}
                longLabels={months.map((m, i) => (i === months.length - 1 ? `${m.long} (bis heute)` : m.long))}
                series={[
                  { key: 'apps', label: 'Apps', color: 'var(--coe-series-1)', values: trend.data.apps },
                  { key: 'flows', label: 'Flows', color: 'var(--coe-series-2)', values: trend.data.flows }
                ]}
              />
            </>
          )}
        </Card>

        <Card title="Handlungsbedarf" subtitle="Governance auf einen Blick" icon="ShieldAlert"
          action={props.onOpenSecurity ? { label: 'Sicherheit', onClick: props.onOpenSecurity } : undefined}>
          <ul className={styles.govList}>
            {govItems.map(g => (
              <li key={g.key}>
                <button type="button" className={styles.govItem} onClick={g.onClick} disabled={!g.onClick}>
                  <span className={`${styles.govIcon} ${g.tone}`}><Icon iconName={g.icon} /></span>
                  <span className={styles.govText}>
                    <span className={styles.govLabel}>{g.label}</span>
                    <span className={styles.govHint}>{g.hint}</span>
                  </span>
                  <span className={styles.govValue}>
                    {g.loading ? <span className={styles.skeleton} style={{ width: 36, height: 18 }} /> : g.failed || g.value === undefined ? '–' : `${formatNumber(g.value)}${g.atLeast ? '+' : ''}`}
                  </span>
                  <Icon iconName="ChevronRight" className={styles.govChevron} />
                </button>
              </li>
            ))}
          </ul>
        </Card>
      </div>

      {/* ---------- Verteilungen ---------- */}
      <div className={styles.ovRow3}>
        <Card title="Apps nach Typ" icon="AppIconDefault" action={T.apps ? { label: 'Apps', onClick: () => onNavigate(T.apps!.key) } : undefined}>
          {appTypes.status === 'loading' && <Skeleton rows={2} height={28} />}
          {appTypes.status === 'error' && <Failed />}
          {appTypes.data && (
            <SegmentedBar label="Apps nach Typ" parts={toParts(
              appTypes.data,
              l => slotColor(APP_TYPE_SLOT[l]),
              l => T.apps && onNavigate(T.apps.key, { admin_powerappstype: { values: [translateValue('admin_powerappstype', l)] } }),
              l => translateValue('admin_powerappstype', l)
            )} />
          )}
        </Card>
        <Card title="Flows nach Status" icon="Flow" action={T.flows ? { label: 'Flows', onClick: () => onNavigate(T.flows!.key) } : undefined}>
          {flowStates.status === 'loading' && <Skeleton rows={2} height={28} />}
          {flowStates.status === 'error' && <Failed />}
          {flowStates.data && (
            <SegmentedBar
              label="Flows nach Status"
              parts={flowStates.data.map(d => ({
                key: d.label,
                label: FLOW_STATE[d.label] ? FLOW_STATE[d.label].label : d.label,
                value: d.value,
                color: FLOW_STATE[d.label] ? FLOW_STATE[d.label].color : 'var(--coe-status-neutral)',
                onClick: () => T.flows && onNavigate(T.flows.key, { admin_flowstate: { values: [translateValue('admin_flowstate', d.label)] } })
              }))}
            />
          )}
        </Card>
        <Card title="Environments nach Typ" icon="Globe" action={T.envs ? { label: 'Environments', onClick: () => onNavigate(T.envs!.key) } : undefined}>
          {envs.status === 'loading' && <Skeleton rows={2} height={28} />}
          {envs.status === 'error' && <Failed />}
          {envs.data && (
            <SegmentedBar label="Environments nach Typ" parts={toParts(
              envs.data.byType,
              l => slotColor(ENV_TYPE_SLOT[l]),
              l => T.envs && onNavigate(T.envs.key, { admin_environmentsku: { values: [translateValue('admin_environmentsku', l)] } }),
              l => translateValue('admin_environmentsku', l)
            )} />
          )}
        </Card>
      </div>

      {/* ---------- Ranglisten ---------- */}
      <div className={styles.ovRow2}>
        <Card
          title="Meistgenutzte Apps"
          subtitle={topApps.data
            ? `Starts in den letzten ${topApps.data.days} Tagen${topApps.data.source === 'audit' ? ' (aus dem CoE-Audit-Log)' : ''}`
            : 'Starts in den letzten 30 Tagen'}
          icon="Trophy2"
        >
          {topApps.status === 'loading' && <Skeleton rows={6} />}
          {topApps.status === 'error' && <Failed />}
          {topApps.data && topApps.data.records.filter(r => numberOf(r, topApps.data!.valueField) > 0).length === 0 && (
            <div className={styles.ovFailed}>
              <Icon iconName="Info" /> Keine Startdaten im CoE – die Zählung braucht den CoE-Audit-Log-Sync bzw. die Spalte „Launches in the past 30 days“.
            </div>
          )}
          {topApps.data && (
            <BarList unit="Starts" items={topApps.data.records.filter(r => numberOf(r, topApps.data!.valueField) > 0).map((r, i): IBarItem => ({
              key: String(r[topApps.data!.primaryId] || i),
              label: textOf(r, 'admin_displayname'),
              sub: `${textOf(r, '_admin_appowner_value')} · ${textOf(r, '_admin_appenvironment_value')}`,
              value: numberOf(r, topApps.data!.valueField),
              onClick: open(T.apps, r, topApps.data!.primaryId)
            }))} />
          )}
        </Card>
        <Card title="Größte Environments" subtitle="nach Anzahl Apps" icon="Globe">
          {envs.status === 'loading' && <Skeleton rows={6} />}
          {envs.status === 'error' && <Failed />}
          {envs.data && (
            <BarList unit="Apps" items={envs.data.records.slice().sort((a, b) => numberOf(b, 'admin_numberofapps') - numberOf(a, 'admin_numberofapps')).slice(0, 8).map((r, i): IBarItem => ({
              key: String(r[envs.data!.primaryId] || i),
              label: textOf(r, 'admin_displayname'),
              sub: `${textOf(r, 'admin_environmentsku')} · ${formatNumber(numberOf(r, 'admin_numberofflows'))} Flows`,
              value: numberOf(r, 'admin_numberofapps'),
              onClick: open(T.envs, r, envs.data!.primaryId)
            }))} />
          )}
        </Card>
        <Card title="Top-Connectors" subtitle="Verwendung in Apps" icon="PlugConnected">
          {topConnectors.status === 'loading' && <Skeleton rows={6} />}
          {topConnectors.status === 'error' && <Failed />}
          {topConnectors.data && (
            <BarList unit="Apps" items={topConnectors.data.records.map((r, i): IBarItem => ({
              key: String(r[topConnectors.data!.primaryId] || i),
              label: textOf(r, 'admin_displayname'),
              sub: `${textOf(r, 'admin_tier')} · ${formatNumber(numberOf(r, 'admin_numflows'))} Flows`,
              value: numberOf(r, 'admin_numapps'),
              onClick: open(T.connectors, r, topConnectors.data!.primaryId)
            }))} />
          )}
        </Card>
        <Card title="Aktivste Maker" subtitle="nach Anzahl Apps" icon="People">
          {topMakers.status === 'loading' && <Skeleton rows={6} />}
          {topMakers.status === 'error' && <Failed />}
          {topMakers.data && (
            <BarList unit="Apps" items={topMakers.data.records.map((r, i): IBarItem => ({
              key: String(r[topMakers.data!.primaryId] || i),
              label: textOf(r, 'admin_displayname'),
              sub: `${textOf(r, 'admin_department') || textOf(r, 'admin_userprincipalname')} · ${formatNumber(numberOf(r, 'admin_numberofflows'))} Flows`,
              value: numberOf(r, 'admin_numberofapps'),
              onClick: open(T.makers, r, topMakers.data!.primaryId)
            }))} />
          )}
        </Card>
      </div>
    </div>
  );
};
