import * as React from 'react';
import { ThemeProvider, PartialTheme } from '@fluentui/react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { IconButton, DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { Icon } from '@fluentui/react/lib/Icon';
import { AadHttpClientFactory } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './CoeCockpit.module.scss';
import { IColumnConfig, IRecord, ISortConfig, ITableConfig } from '../models/ITableConfig';
import {
  DataverseError, DataverseService, environmentLinkId, findDefaultEnvironment, ICountResult, IEntityMeta, IEnvironmentInfo, normalizeDataverseUrl
} from '../services/DataverseService';
import { buildDefaultEnvironmentTables, DEFAULT_ENV_PREFIX, isDefaultEnvKey } from '../models/defaultEnvironment';
import { buildDevEnvTable, DEV_ENV_DEADLINES, DEV_ENV_KEY, DEV_FIELDS, loadAffectedEnvironments, RECOMMENDATION } from '../models/devEnvironments';
import { OVERVIEW_KEY } from '../models/overview';
import { buildPipelineTables, GROUP_ALM, PIPELINE_KEY_PREFIX } from '../models/pipelines';
import { OverviewView } from './OverviewView';
import { SECURITY_KEY, SecurityView } from './SecurityView';
import { GlobalSearch } from './GlobalSearch';
import { copyCurrentLink, IViewState, readViewState, writeViewState } from '../utils/deepLink';
import {
  displayValue, downloadText, formatCount, formatNumber, formatTime, hostOf, isValidEnvironmentUrl, sortValue, toCsv
} from '../utils/format';
import { describeFilter, IColumnFilter, isFilterActive, matchesFilter } from '../utils/filters';
import { buildThemeVars } from './themeVars';
import { DataTable } from './DataTable';
import { IKpiState, KpiStrip } from './KpiStrip';
import { RecordPanel } from './RecordPanel';
import { TableEditor } from './TableEditor';

export interface ICoeCockpitProps {
  instanceId: string;
  title: string;
  subtitle: string;
  environmentUrl: string;
  tables: ITableConfig[];
  pageSize: number;
  maxRecords: number;
  showKpis: boolean;
  density: 'compact' | 'comfortable';
  theme: IReadonlyTheme | undefined;
  isEditMode: boolean;
  editorRequest: number;
  aadHttpClientFactory: AadHttpClientFactory;
  showDefaultEnvironment: boolean;
  /** Leer = automatisch erkennen (Umgebungstyp „Default“). */
  defaultEnvironmentId: string;
  /** Optional: Dataverse-URL des Default Environments (für Solutions), sonst aus dem CoE-Inventar. */
  defaultEnvironmentDataverseUrl: string;
  /** Reiter „Betroffen ab 15.12.“ (nicht verwaltete Developer-Environments) anzeigen. */
  showDevEnvironments: boolean;
  /** Übersicht (Dashboard) als Startseite anzeigen. */
  showOverview: boolean;
  /** Spalte mit dem Managed-Status im CoE-Inventar (leer = automatisch erkennen). */
  devEnvManagedField: string;
  /** Ansicht „Sicherheit & Freigaben“ anzeigen. */
  showSecurity: boolean;
  /** Dataverse-URL des Pipelines-Host-Environments (leer = ALM-Pipelines ausblenden). */
  pipelinesHostUrl: string;
  onTablesChanged: (tables: ITableConfig[]) => void;
  onConfigure: () => void;
}

interface IDefaultEnvState {
  status: 'idle' | 'loading' | 'ready' | 'notfound' | 'error';
  env?: IEnvironmentInfo;
  error?: string;
}

interface ITableState {
  status: 'loading' | 'ready' | 'error';
  loadKey: string;
  records: IRecord[];
  /** true = Zwischenstand: weitere Seiten werden noch geladen. */
  partial?: boolean;
  meta?: IEntityMeta;
  warnings: string[];
  truncated?: boolean;
  error?: string;
  errorHint?: string;
  loadedAt?: Date;
}

function errorHint(e: unknown, table?: ITableConfig): string {
  const status = e instanceof DataverseError ? e.status : 0;
  const msg = e instanceof Error ? e.message : String(e);
  if (/AADSTS65001|consent|AADSTS500011|AADSTS50105/i.test(msg)) {
    return 'Die API-Berechtigung „Dynamics CRM – user_impersonation“ ist noch nicht genehmigt. Ein SharePoint-Admin muss sie im SharePoint Admin Center unter Erweitert → API-Zugriff freigeben.';
  }
  if (table && table.environmentUrl && (status === 401 || status === 403 || /privilege|prvRead|not a member|does not have/i.test(msg))) {
    return `Dein Konto hat in der Umgebung ${hostOf(table.environmentUrl)} keine Leserechte für „${table.title}“. ${ADMIN_ROLE_HINT}`;
  }
  if (status === 401 || status === 403 || /privilege|prvRead|not a member|does not have/i.test(msg)) {
    return 'Dein Konto hat in der CoE-Umgebung keine ausreichenden Leserechte. Benötigt wird eine Sicherheitsrolle mit Lesezugriff auf die CoE-Tabellen (z. B. „Power Platform Admin SR“).';
  }
  if (status === 404 || /Could not find|does not exist/i.test(msg)) {
    return 'Die Tabelle wurde in dieser Umgebung nicht gefunden. Prüfe den logischen Tabellennamen im Tabellen-Editor und die Umgebungs-URL.';
  }
  return 'Prüfe die Umgebungs-URL und ob die CoE-Umgebung erreichbar ist.';
}

/** Hinweis, wenn Daten außerhalb des CoE-Environments nicht gelesen werden dürfen. */
const ADMIN_ROLE_HINT = 'Diese Daten liegen nicht im CoE, sondern im jeweiligen Environment. Als Power-Platform-Admin: Rolle „Power Platform-Administrator“ in Entra PIM aktivieren und – falls nötig – im Power Platform Admin Center beim Environment unter „Mitgliedschaft“ → „Mich hinzufügen“ die Rolle „System Administrator“ übernehmen. Danach die Seite neu laden.';

/** Ergebnis der Zugriffsprüfung für Tabellen außerhalb des CoE-Environments. */
interface IAccessState {
  status: 'checking' | 'ok' | 'denied' | 'missing' | 'error';
  message?: string;
}

function accessStatus(e: unknown): IAccessState {
  const status = e instanceof DataverseError ? e.status : 0;
  const msg = e instanceof Error ? e.message : String(e);
  if (status === 401 || status === 403 || /privilege|prvRead|not a member|does not have|AADSTS/i.test(msg)) {
    return { status: 'denied', message: msg };
  }
  if (status === 404 || /Could not find|does not exist|Resource not found/i.test(msg)) {
    return { status: 'missing', message: msg };
  }
  return { status: 'error', message: msg };
}

function readHidden(storageKey: string): { [tableKey: string]: string[] } | undefined {
  try {
    const raw = window.localStorage.getItem(storageKey);
    return raw ? JSON.parse(raw) as { [tableKey: string]: string[] } : undefined;
  } catch {
    return undefined;
  }
}

function writeHidden(storageKey: string, value: { [tableKey: string]: string[] }): void {
  try {
    window.localStorage.setItem(storageKey, JSON.stringify(value));
  } catch {
    // Browser-Speicher nicht verfügbar – Spaltenauswahl gilt dann nur für diese Sitzung.
  }
}

const PAGE_SIZES = [10, 25, 50, 100];
/** Bereich für Tabellen ohne eigene Gruppe. */
const DEFAULT_GROUP = 'Inventar';
/** Bereich für die Übersicht (Startseite). */
const START_GROUP = 'Start';
/** Bereich für berechnete Ansichten (Default Environment, Betroffen ab 15.12.). */
const ANALYSIS_GROUP = 'Auswertungen';
/** Ab so vielen Einträgen wird statt der Reiterzeile die Seitenleiste verwendet. */
const MAX_TABS = 7;
/** Unter dieser Breite (px) wird die schmale Darstellung verwendet. */
const NARROW_WIDTH = 860;
const COLLATOR = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });

const CoeCockpit: React.FC<ICoeCockpitProps> = props => {
  const { tables, theme, isEditMode, showKpis, maxRecords } = props;
  const themeVars = React.useMemo(() => buildThemeVars(theme), [theme]);
  const envValid = isValidEnvironmentUrl(props.environmentUrl);

  const service = React.useMemo(
    () => (envValid ? new DataverseService(props.aadHttpClientFactory, props.environmentUrl) : undefined),
    [envValid, props.environmentUrl, props.aadHttpClientFactory]
  );

  const servicesRef = React.useRef<{ [url: string]: DataverseService }>({});
  /** Zuletzt angeforderter Ladeschlüssel je Tabelle (verhindert veraltete Ergebnisse). */
  const requestedRef = React.useRef<{ [tableKey: string]: string }>({});
  const mountedRef = React.useRef<boolean>(true);
  /** Ansicht aus einem geteilten Link (#coe=…) – wird beim Start einmal übernommen. */
  const initialViewRef = React.useRef<IViewState | undefined>(readViewState());
  const [activeKey, setActiveKey] = React.useState<string | undefined>(
    (initialViewRef.current && initialViewRef.current.t) || (props.showOverview ? OVERVIEW_KEY : tables.length ? tables[0].key : undefined)
  );
  /** Filter/Suche, die beim nächsten Reiterwechsel gesetzt werden (Klick in der Übersicht, geteilter Link). */
  const pendingFiltersRef = React.useRef<{ [field: string]: IColumnFilter } | undefined>(initialViewRef.current && initialViewRef.current.f);
  const pendingSearchRef = React.useRef<string | undefined>(initialViewRef.current && initialViewRef.current.q);
  const [linkCopied, setLinkCopied] = React.useState<boolean>(false);
  // Breite des Webparts messen: schmale Abschnitte/Handy bekommen die Klasse „narrow“
  // (CSS-Container-Abfragen werden vom SPFx-Build entfernt, daher per Script)
  const rootRef = React.useRef<HTMLElement>(null);
  const [narrow, setNarrow] = React.useState<boolean>(false);
  React.useEffect(() => {
    const el = rootRef.current;
    if (!el) {
      return;
    }
    const update = (): void => setNarrow(el.getBoundingClientRect().width < NARROW_WIDTH);
    update();
    const RO = (window as unknown as { ResizeObserver?: new (cb: () => void) => { observe: (e: Element) => void; disconnect: () => void } }).ResizeObserver;
    if (RO) {
      const ro = new RO(update);
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);
  const [lastDefaultKey, setLastDefaultKey] = React.useState<string>(`${DEFAULT_ENV_PREFIX}apps`);
  const [defaultEnv, setDefaultEnv] = React.useState<IDefaultEnvState>({ status: 'idle' });
  const [data, setData] = React.useState<{ [key: string]: ITableState }>({});
  const [kpis, setKpis] = React.useState<{ [key: string]: IKpiState }>({});
  const [reloadTick, setReloadTick] = React.useState<number>(0);
  const [search, setSearch] = React.useState<string>('');
  const [filters, setFilters] = React.useState<{ [field: string]: IColumnFilter }>({});
  const [sort, setSort] = React.useState<ISortConfig | undefined>(undefined);
  const [page, setPage] = React.useState<number>(0);
  const [pageSize, setPageSize] = React.useState<number>(props.pageSize || 25);
  const [compact, setCompact] = React.useState<boolean>(props.density === 'compact');
  /** Geöffnete Datensätze (Detailansicht); mehrere = durchgeklickt, „Zurück“ führt zum vorherigen. */
  const [panelStack, setPanelStack] = React.useState<{ table: ITableConfig; record: IRecord; meta?: IEntityMeta }[]>([]);
  const panel = panelStack.length ? panelStack[panelStack.length - 1] : undefined;
  const [editorOpen, setEditorOpen] = React.useState<boolean>(false);
  const [collapsedGroups, setCollapsedGroups] = React.useState<{ [group: string]: boolean }>({});
  const storageKey = `coe-cockpit:${props.instanceId}:hidden`;
  const [hidden, setHidden] = React.useState<{ [tableKey: string]: string[] }>(() => readHidden(storageKey) || {});

  const tablesSig = React.useMemo(() => JSON.stringify(tables), [tables]);
  const mainKey = `${service ? service.baseUrl : ''}|${tablesSig}|${reloadTick}|${maxRecords}`;

  // Default Environment ermitteln
  React.useEffect(() => {
    if (!service || !props.showDefaultEnvironment) {
      setDefaultEnv({ status: 'idle' });
      return;
    }
    let cancelled = false;
    setDefaultEnv({ status: 'loading' });
    service.getEnvironments().then(
      envs => {
        if (!cancelled) {
          const env = findDefaultEnvironment(envs, props.defaultEnvironmentId);
          setDefaultEnv(env ? { status: 'ready', env } : { status: 'notfound' });
        }
      },
      (e: unknown) => {
        if (!cancelled) {
          setDefaultEnv({ status: 'error', error: e instanceof Error ? e.message : String(e) });
        }
      }
    );
    return () => { cancelled = true; };
  }, [service, props.showDefaultEnvironment, props.defaultEnvironmentId]);

  const defaultTables = React.useMemo(
    () => (defaultEnv.status === 'ready' && defaultEnv.env
      ? buildDefaultEnvironmentTables(tables, defaultEnv.env, props.defaultEnvironmentDataverseUrl)
      : []),
    [tablesSig, defaultEnv, props.defaultEnvironmentDataverseUrl]
  );
  const defaultSig = React.useMemo(() => JSON.stringify(defaultTables), [defaultTables]);
  const devTables = React.useMemo(() => (props.showDevEnvironments ? [buildDevEnvTable()] : []), [props.showDevEnvironments]);
  const pipelineTables = React.useMemo(
    () => (props.pipelinesHostUrl ? buildPipelineTables(props.pipelinesHostUrl) : []),
    [props.pipelinesHostUrl]
  );
  const allTables = tables.concat(defaultTables).concat(devTables).concat(pipelineTables);
  /** Tabellen für die Suche über alle Bereiche (ohne die gefilterten Kopien fürs Default Environment). */
  const searchTables = React.useMemo(() => tables.concat(pipelineTables), [tablesSig, pipelineTables]);
  const inDevView = props.showDevEnvironments && activeKey === DEV_ENV_KEY;
  const activeTable = allTables.filter(t => t.key === activeKey)[0];
  const inDefaultView = props.showDefaultEnvironment && isDefaultEnvKey(activeKey);

  /** Ladeschlüssel je Tabelle – ändert sich, wenn Daten neu geladen werden müssen. */
  const keyFor = (t: ITableConfig): string => (isDefaultEnvKey(t.key)
    ? `${mainKey}|${defaultSig}`
    : t.key === DEV_ENV_KEY ? `${mainKey}|${props.devEnvManagedField}` : mainKey);

  /** Dataverse-Service für eine Tabelle (CoE-Umgebung oder eigene Umgebung der Tabelle). */
  const serviceFor = (t: ITableConfig | undefined): DataverseService | undefined => {
    if (!service || !t || !t.environmentUrl) {
      return service;
    }
    const url = normalizeDataverseUrl(t.environmentUrl) || '';
    if (!url || url.toLowerCase() === service.baseUrl.toLowerCase()) {
      return service;
    }
    if (!servicesRef.current[url]) {
      servicesRef.current[url] = new DataverseService(props.aadHttpClientFactory, url);
    }
    return servicesRef.current[url];
  };
  const activeLoadKey = activeTable ? keyFor(activeTable) : '';

  React.useEffect(() => { setPageSize(props.pageSize || 25); }, [props.pageSize]);
  React.useEffect(() => { setCompact(props.density === 'compact'); }, [props.density]);
  React.useEffect(() => { if (props.editorRequest > 0) { setEditorOpen(true); } }, [props.editorRequest]);

  React.useEffect(() => {
    const waitingForDefault = (props.showDefaultEnvironment && isDefaultEnvKey(activeKey)) || (props.showOverview && activeKey === OVERVIEW_KEY)
      || (props.showSecurity && activeKey === SECURITY_KEY);
    if (!activeTable && tables.length && !waitingForDefault) {
      setActiveKey(tables[0].key);
    }
    // activeTable/defaultSig gehören dazu: z. B. wird beim Ausschalten von „Default Environment“
    // der aktive Reiter erst im nächsten Render ungültig.
    if (((!props.showOverview && activeKey === OVERVIEW_KEY) || (!props.showSecurity && activeKey === SECURITY_KEY)) && tables.length) {
      setActiveKey(props.showOverview ? OVERVIEW_KEY : tables[0].key);
    }
  }, [tablesSig, props.showDefaultEnvironment, props.showDevEnvironments, props.showOverview, props.showSecurity, props.pipelinesHostUrl, defaultSig, !!activeTable]);

  React.useEffect(() => {
    if (activeKey && isDefaultEnvKey(activeKey)) {
      setLastDefaultKey(activeKey);
    }
  }, [activeKey]);

  // Ansicht zurücksetzen, wenn der Reiter (oder dessen Konfiguration) wechselt.
  // Vorgemerkte Filter/Suche werden erst übernommen, wenn die Tabelle feststeht (Default Environment lädt später).
  const activeTableSig = activeTable ? JSON.stringify(activeTable) : '';
  React.useEffect(() => {
    const ready = !!activeTable;
    setSearch(ready ? pendingSearchRef.current || '' : '');
    setFilters(ready ? pendingFiltersRef.current || {} : {});
    if (ready) {
      pendingFiltersRef.current = undefined;
      pendingSearchRef.current = undefined;
    }
    setPage(0);
    setPanelStack(prev => (prev.length && prev[0].table.key === (activeTable && activeTable.key) ? prev : []));
    setSort(activeTable && activeTable.defaultSort ? { ...activeTable.defaultSort } : undefined);
  }, [activeKey, activeTableSig]);

  // Datensätze des aktiven Reiters laden
  React.useEffect(() => {
    const svc = serviceFor(activeTable);
    if (!svc || !activeTable || !activeTable.logicalName || activeTable.unavailable) {
      return;
    }
    const key = activeTable.key;
    const loadKey = activeLoadKey;
    const existing = data[key];
    // Läuft die Anfrage noch (z. B. nach kurzem Reiterwechsel), wird ihr Ergebnis abgewartet.
    if (existing && existing.loadKey === loadKey && (existing.status !== 'loading' || requestedRef.current[key] === loadKey)) {
      return;
    }
    // Nicht beim Reiterwechsel abbrechen: Das Ergebnis wird übernommen, solange es die
    // zuletzt angeforderte Version dieser Tabelle ist.
    requestedRef.current[key] = loadKey;
    const isCurrent = (): boolean => mountedRef.current && requestedRef.current[key] === loadKey;
    setData(prev => ({
      ...prev,
      [key]: { status: 'loading', loadKey, records: prev[key] ? prev[key].records : [], warnings: [] }
    }));
    const request = activeTable.key === DEV_ENV_KEY
      ? loadAffectedEnvironments(svc, props.devEnvManagedField, maxRecords)
      : svc.getRecords(activeTable, maxRecords, (partialRecords, meta) => {
        if (isCurrent()) {
          setData(prev => ({
            ...prev,
            [key]: { status: 'loading', partial: true, loadKey, records: partialRecords, meta, warnings: [] }
          }));
        }
      });
    request.then(
      res => {
        if (isCurrent()) {
          setData(prev => ({
            ...prev,
            [key]: { status: 'ready', loadKey, records: res.records, meta: res.meta, warnings: res.warnings, truncated: res.truncated, loadedAt: new Date() }
          }));
        }
      },
      (e: unknown) => {
        if (isCurrent()) {
          setData(prev => ({
            ...prev,
            [key]: { status: 'error', loadKey, records: [], warnings: [], error: e instanceof Error ? e.message : String(e), errorHint: errorHint(e, activeTable), loadedAt: new Date() }
          }));
        }
      }
    );
  }, [service, activeKey, activeLoadKey]);

  React.useEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);

  // Kennzahlen laden
  React.useEffect(() => {
    if (!service || !showKpis) {
      return;
    }
    let cancelled = false;
    const lk = mainKey;
    tables.filter(t => t.showInKpis !== false && !!t.logicalName).forEach(t => {
      setKpis(prev => ({ ...prev, [t.key]: { status: 'loading', loadKey: lk } }));
      const highlightFilter = t.kpi && t.kpi.highlightFilter ? t.kpi.highlightFilter : undefined;
      const svc = serviceFor(t) || service; // Tabellen mit eigener Umgebungs-URL dort zählen
      const totalP = svc.getCountWithBase(t);
      // Hervorhebung: Filter der Reihe nach probieren (Spaltentyp unterscheidet sich je nach CoE-Version)
      const highlightFilters = highlightFilter ? [highlightFilter].concat((t.kpi && t.kpi.highlightFilterAlternatives) || []) : [];
      const tryHighlight = (i: number): Promise<ICountResult | undefined> => (i >= highlightFilters.length
        ? Promise.resolve(undefined)
        : svc.getCountWithBase(t, highlightFilters[i]).catch(() => tryHighlight(i + 1)));
      const highlightP: Promise<ICountResult | undefined> = tryHighlight(0);
      Promise.all([totalP, highlightP]).then(
        ([total, highlight]) => {
          if (!cancelled) {
            setKpis(prev => ({
              ...prev,
              [t.key]: {
                status: 'ready', loadKey: lk,
                total: total.value, totalAtLeast: total.atLeast,
                highlight: highlight ? highlight.value : undefined, highlightAtLeast: highlight ? highlight.atLeast : undefined
              }
            }));
          }
        },
        () => {
          if (!cancelled) {
            setKpis(prev => ({ ...prev, [t.key]: { status: 'error', loadKey: lk } }));
          }
        }
      );
    });
    return () => { cancelled = true; };
  }, [service, mainKey, showKpis]);

  // Betroffene Developer-Environments vorab laden (Übersicht + Zähler im Reiter)
  React.useEffect(() => {
    if (!service || !props.showDevEnvironments || activeKey === DEV_ENV_KEY) {
      return;
    }
    const devTable = devTables[0];
    if (!devTable) {
      return;
    }
    const loadKey = keyFor(devTable);
    const existing = data[DEV_ENV_KEY];
    if (existing && existing.loadKey === loadKey) {
      return;
    }
    requestedRef.current[DEV_ENV_KEY] = loadKey;
    loadAffectedEnvironments(service, props.devEnvManagedField, maxRecords).then(
      res => {
        if (mountedRef.current && requestedRef.current[DEV_ENV_KEY] === loadKey) {
          setData(prev => ({ ...prev, [DEV_ENV_KEY]: { status: 'ready', loadKey, records: res.records, meta: res.meta, warnings: res.warnings, truncated: res.truncated, loadedAt: new Date() } }));
        }
      },
      () => undefined
    );
  }, [service, mainKey, props.showDevEnvironments, props.devEnvManagedField]);

  // Gesamtzahl einer abgeschnittenen Tabelle nachladen (für den Hinweis „… von X Einträgen“)
  React.useEffect(() => {
    const st = activeTable ? data[activeTable.key] : undefined;
    if (!activeTable || activeTable.virtual || !st || st.status !== 'ready' || !st.truncated) {
      return;
    }
    const k = kpis[activeTable.key];
    if (k && k.loadKey.indexOf(mainKey) === 0) {
      return;
    }
    const t = activeTable;
    const svc = serviceFor(t);
    if (!svc) {
      return;
    }
    const lk = keyFor(t);
    setKpis(prev => ({ ...prev, [t.key]: { status: 'loading', loadKey: lk } }));
    svc.getCountWithBase(t).then(
      total => setKpis(prev => ({ ...prev, [t.key]: { status: 'ready', loadKey: lk, total: total.value, totalAtLeast: total.atLeast } })),
      () => setKpis(prev => ({ ...prev, [t.key]: { status: 'error', loadKey: lk } }))
    );
  }, [activeKey, data, mainKey]);

  // Anzahl Apps / Flows / Solutions im Default Environment
  React.useEffect(() => {
    if (!service || defaultTables.length === 0) {
      return;
    }
    let cancelled = false;
    const lk = `${mainKey}|${defaultSig}`;
    defaultTables.forEach(t => {
      const svc = serviceFor(t);
      if (t.unavailable || !svc) {
        setKpis(prev => ({ ...prev, [t.key]: { status: 'error', loadKey: lk } }));
        return;
      }
      setKpis(prev => ({ ...prev, [t.key]: { status: 'loading', loadKey: lk } }));
      svc.getCountWithBase(t).then(
        total => { if (!cancelled) { setKpis(prev => ({ ...prev, [t.key]: { status: 'ready', loadKey: lk, total: total.value, totalAtLeast: total.atLeast } })); } },
        () => { if (!cancelled) { setKpis(prev => ({ ...prev, [t.key]: { status: 'error', loadKey: lk } })); } }
      );
    });
    return () => { cancelled = true; };
  }, [service, mainKey, defaultSig]);

  // Zugriff auf Tabellen außerhalb des CoE vorab prüfen (Default Environment, Pipelines-Host),
  // damit vor dem Öffnen klar ist, was ohne aktivierte Adminrolle nicht sichtbar ist.
  const externalTables = allTables.filter(t => !!t.environmentUrl && !t.unavailable && serviceFor(t) !== service);
  const externalSig = externalTables.map(t => `${t.key}|${t.environmentUrl}|${t.logicalName}`).join(',');
  const [access, setAccess] = React.useState<{ [tableKey: string]: IAccessState }>({});
  const [accessDismissed, setAccessDismissed] = React.useState<boolean>(false);
  React.useEffect(() => {
    if (!service || !externalTables.length) {
      return;
    }
    let cancelled = false;
    externalTables.forEach(t => {
      const svc = serviceFor(t);
      if (!svc) {
        return;
      }
      setAccess(prev => ({ ...prev, [t.key]: { status: 'checking' } }));
      svc.getRecords({ ...t, key: `_probe_${t.key}`, columns: [], links: [], defaultSort: undefined }, 1).then(
        () => { if (!cancelled) { setAccess(prev => ({ ...prev, [t.key]: { status: 'ok' } })); } },
        (e: unknown) => { if (!cancelled) { setAccess(prev => ({ ...prev, [t.key]: accessStatus(e) })); } }
      );
    });
    return () => { cancelled = true; };
  }, [service, externalSig, reloadTick]);
  const blocked = externalTables.filter(t => access[t.key] && ['denied', 'missing', 'error'].indexOf(access[t.key].status) >= 0);
  const isLocked = (key: string): boolean => !!access[key] && access[key].status === 'denied' && externalTables.some(t => t.key === key);

  const state: ITableState | undefined = activeTable ? data[activeTable.key] : undefined;
  const records = state ? state.records : [];
  const loading = !!service && !!activeTable && !activeTable.unavailable && (!state || (state.status === 'loading' && !state.partial));
  /** Erste Seiten sind da, weitere werden noch geladen (große Tabellen). */
  const partial = !!state && state.status === 'loading' && !!state.partial;
  /** Exakte Gesamtzahl laut Dataverse (falls bekannt) – für Fortschritt und Hinweis bei Begrenzung. */
  const activeKpi = activeTable ? kpis[activeTable.key] : undefined;
  const serverTotal = activeKpi && activeKpi.status === 'ready' && activeKpi.loadKey.indexOf(mainKey) === 0 ? activeKpi : undefined;

  const hiddenSet = React.useMemo(() => {
    const set: { [f: string]: boolean } = {};
    if (activeTable) {
      const stored = hidden[activeTable.key];
      if (stored) {
        stored.forEach(f => { set[f] = true; });
      } else {
        activeTable.columns.forEach(c => { if (c.hidden) { set[c.field] = true; } });
      }
    }
    return set;
  }, [activeTable, hidden]);

  const allColumns: IColumnConfig[] = activeTable ? activeTable.columns.filter(c => !!c.field) : [];
  const visibleColumns = allColumns.filter(c => !hiddenSet[c.field]);
  const columnByField: { [field: string]: IColumnConfig } = {};
  allColumns.forEach(c => { columnByField[c.field] = c; });
  const activeFilterFields = Object.keys(filters).filter(f => isFilterActive(filters[f]) && !!columnByField[f]);

  /** Suche + alle Spaltenfilter, optional ohne den Filter einer Spalte (für deren Werteliste). */
  const columnsSig = allColumns.map(c => c.field + c.type).join('|');
  /** Kleingeschriebener Suchtext je Datensatz – einmal pro Datenstand statt bei jedem Tastendruck. */
  const searchIndex = React.useMemo(() => {
    const map = new Map<IRecord, string>();
    records.forEach(r => {
      map.set(r, allColumns.map(c => displayValue(r, c.field)).join('\u0001').toLowerCase());
    });
    return map;
  }, [records, columnsSig]);

  const applyFilters = (rows: IRecord[], exceptField?: string): IRecord[] => {
    const q = search.trim().toLowerCase();
    const fields = activeFilterFields.filter(f => f !== exceptField);
    return rows.filter(r => {
      for (let i = 0; i < fields.length; i++) {
        if (!matchesFilter(r, columnByField[fields[i]], filters[fields[i]])) {
          return false;
        }
      }
      if (!q) {
        return true;
      }
      const text = searchIndex.get(r);
      return text !== undefined
        ? text.indexOf(q) >= 0
        : allColumns.some(c => displayValue(r, c.field).toLowerCase().indexOf(q) >= 0);
    });
  };

  const filtered = React.useMemo(() => {
    let rows = applyFilters(records);
    const sortCol = sort ? allColumns.filter(c => c.field === sort.field)[0] : undefined;
    if (sort && sortCol) {
      const dir = sort.desc ? -1 : 1;
      // Schlüssel einmal berechnen (statt in jedem Vergleich) – wichtig bei zehntausenden Einträgen.
      const keyed = rows.map(r => ({ r, v: sortValue(r, sortCol) }));
      keyed.sort((a, b) => {
        if (typeof a.v === 'number' && typeof b.v === 'number') {
          return (a.v - b.v) * dir;
        }
        return COLLATOR.compare(String(a.v), String(b.v)) * dir;
      });
      rows = keyed.map(k => k.r);
    }
    return rows;
  }, [records, searchIndex, search, filters, sort, columnsSig]);

  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
  const safePage = Math.min(page, pageCount - 1);
  const pageRows = filtered.slice(safePage * pageSize, safePage * pageSize + pageSize);
  const activeFilterCount = activeFilterFields.length + (search ? 1 : 0);

  const setColumnFilter = (field: string, filter: IColumnFilter | undefined): void => {
    setFilters(prev => {
      const next = { ...prev };
      if (filter && isFilterActive(filter)) {
        next[field] = filter;
      } else {
        delete next[field];
      }
      return next;
    });
    setPage(0);
  };

  const onSort = (field: string): void => {
    setSort(prev => (prev && prev.field === field ? { field, desc: !prev.desc } : { field, desc: false }));
    setPage(0);
  };

  const toggleColumn = (field: string): void => {
    if (!activeTable) {
      return;
    }
    const current = Object.keys(hiddenSet).filter(f => hiddenSet[f]);
    const next = current.indexOf(field) >= 0 ? current.filter(f => f !== field) : current.concat([field]);
    const value = { ...hidden, [activeTable.key]: next };
    setHidden(value);
    writeHidden(storageKey, value);
  };

  const exportCsv = (): void => {
    if (!activeTable) {
      return;
    }
    const stamp = new Date().toISOString().substring(0, 10);
    downloadText(`${activeTable.title.replace(/[^\w\-äöüÄÖÜß ]/g, '')}_${stamp}.csv`, toCsv(filtered, visibleColumns), 'text/csv;charset=utf-8');
  };

  const columnMenu: IContextualMenuItem[] = allColumns.map(c => ({
    key: c.field,
    text: c.label,
    canCheck: true,
    checked: !hiddenSet[c.field],
    onClick: (ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
      if (ev) { ev.preventDefault(); }
      toggleColumn(c.field);
      return false;
    }
  }));

  const countFor = (t: ITableConfig): string | undefined => {
    const k = kpis[t.key];
    if (k && k.status === 'ready' && k.total !== undefined) {
      return formatCount(k.total, k.totalAtLeast);
    }
    const d = data[t.key];
    if (d && d.status === 'ready' && d.loadKey === keyFor(t)) {
      return formatNumber(d.records.length) + (d.truncated ? '+' : '');
    }
    return undefined;
  };

  const renderPager = (): JSX.Element => {
    const pages: number[] = [];
    const start = Math.max(0, Math.min(safePage - 2, pageCount - 5));
    for (let i = start; i < Math.min(pageCount, start + 5); i++) {
      pages.push(i);
    }
    return (
      <div className={styles.pager}>
        <button type="button" className={styles.pageButton} disabled={safePage === 0} onClick={() => setPage(safePage - 1)} aria-label="Vorherige Seite">
          <Icon iconName="ChevronLeft" />
        </button>
        {pages.map(p => (
          <button key={p} type="button" className={`${styles.pageButton} ${p === safePage ? styles.pageButtonActive : ''}`} onClick={() => setPage(p)} aria-current={p === safePage ? 'page' : undefined}>
            {p + 1}
          </button>
        ))}
        <button type="button" className={styles.pageButton} disabled={safePage >= pageCount - 1} onClick={() => setPage(safePage + 1)} aria-label="Nächste Seite">
          <Icon iconName="ChevronRight" />
        </button>
      </div>
    );
  };

  /** Hinweis: Bereiche außerhalb des CoE, die mit den aktuellen Rechten nicht sichtbar sind. */
  const renderAccessNotice = (): JSX.Element | null => {
    if (!blocked.length || accessDismissed) {
      return null;
    }
    const areaOf = (t: ITableConfig): string => (isDefaultEnvKey(t.key) ? 'Default Environment' : t.key.indexOf(PIPELINE_KEY_PREFIX) === 0 ? GROUP_ALM : t.group || 'Weitere Umgebung');
    const groups: { area: string; host: string; titles: string[]; status: IAccessState['status'] }[] = [];
    blocked.forEach(t => {
      const area = areaOf(t);
      const host = hostOf(t.environmentUrl || '');
      let g = groups.filter(x => x.area === area && x.host === host)[0];
      if (!g) {
        g = { area, host, titles: [], status: access[t.key].status };
        groups.push(g);
      }
      g.titles.push(t.title);
      if (access[t.key].status === 'denied') {
        g.status = 'denied';
      }
    });
    const anyDenied = groups.some(g => g.status === 'denied');
    const reason = (g: { area: string; status: IAccessState['status'] }): string => (g.status === 'denied'
      ? 'kein Zugriff mit deinem aktuellen Konto'
      : g.status === 'missing'
        ? (g.area === GROUP_ALM ? 'Pipelines-Tabellen nicht gefunden – ist das richtige Host-Environment gewählt?' : 'Tabelle in dieser Umgebung nicht gefunden')
        : 'nicht erreichbar');
    return (
      <div className={styles.accessNotice} role="status">
        <span className={styles.accessIcon}><Icon iconName={anyDenied ? 'Lock' : 'Info'} /></span>
        <div className={styles.accessText}>
          <div className={styles.accessTitle}>
            {anyDenied ? 'Power-Platform-Adminrolle erforderlich – folgende Bereiche sind gerade nicht sichtbar' : 'Einige Bereiche sind nicht verfügbar'}
          </div>
          <ul className={styles.accessList}>
            {groups.map(g => (
              <li key={`${g.area}|${g.host}`}>
                <strong>{g.area}</strong> · {g.titles.join(', ')} <span className={styles.accessHost}>({g.host})</span> – {reason(g)}
              </li>
            ))}
          </ul>
          {anyDenied && <div className={styles.accessHint}>{ADMIN_ROLE_HINT}</div>}
        </div>
        <div className={styles.accessActions}>
          <DefaultButton text="Erneut prüfen" iconProps={{ iconName: 'Refresh' }} onClick={() => setReloadTick(t => t + 1)} />
          <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Hinweis ausblenden" title="Hinweis ausblenden" onClick={() => setAccessDismissed(true)} />
        </div>
      </div>
    );
  };

  const renderOnboarding = (): JSX.Element => (
    <div className={styles.state}>
      <div className={styles.stateIcon}><Icon iconName="PlugConnected" /></div>
      <p className={styles.stateTitle}>Mit der CoE-Umgebung verbinden</p>
      <p className={styles.stateText}>
        Das Cockpit liest die Tabellen des CoE Starter Kits direkt aus Dataverse. Trage dazu die URL deiner CoE-Umgebung in den Webpart-Eigenschaften ein.
      </p>
      <div className={styles.steps}>
        <div className={styles.step}>
          <span className={styles.stepNo}>1</span>
          <span className={styles.stepTitle}>Umgebungs-URL eintragen</span>
          z. B. https://contoso-coe.crm4.dynamics.com
        </div>
        <div className={styles.step}>
          <span className={styles.stepNo}>2</span>
          <span className={styles.stepTitle}>API-Zugriff genehmigen</span>
          SharePoint Admin Center → Erweitert → API-Zugriff → „Dynamics CRM“ genehmigen
        </div>
        <div className={styles.step}>
          <span className={styles.stepNo}>3</span>
          <span className={styles.stepTitle}>Tabellen anpassen</span>
          Apps, Flows, Environments und Maker sind vorkonfiguriert – weitere Tabellen per Editor ergänzen.
        </div>
      </div>
      {isEditMode && (
        <div className={styles.stateActions}>
          <PrimaryButton text="Eigenschaften öffnen" iconProps={{ iconName: 'Settings' }} onClick={props.onConfigure} />
        </div>
      )}
    </div>
  );

  const renderDefaultEnvironment = (): JSX.Element => {
    if (defaultEnv.status === 'idle' || defaultEnv.status === 'loading') {
      return (
        <div className={styles.envBanner}>
          <div className={styles.envBannerInfo}>
            <div className={styles.envBannerIcon}><Icon iconName="Home" /></div>
            <div style={{ flex: '1 1 auto' }}>
              <span className={styles.skeleton} style={{ width: 120, marginBottom: 8 }} />
              <span className={styles.skeleton} style={{ width: 260, height: 16 }} />
            </div>
          </div>
        </div>
      );
    }
    if (defaultEnv.status === 'error' || defaultEnv.status === 'notfound' || !defaultEnv.env) {
      const configured = !!(props.defaultEnvironmentId || '').trim();
      return (
        <div className={styles.state}>
          <div className={`${styles.stateIcon} ${defaultEnv.status === 'error' ? styles.stateIconDanger : ''}`}><Icon iconName="Home" /></div>
          <p className={styles.stateTitle}>Default Environment {defaultEnv.status === 'error' ? 'konnte nicht ermittelt werden' : 'nicht gefunden'}</p>
          <p className={styles.stateText}>
            {defaultEnv.status === 'error'
              ? 'Die Umgebungen konnten nicht aus dem CoE-Inventar gelesen werden.'
              : configured
                ? 'Die in den Webpart-Eigenschaften gewählte Umgebung ist im CoE-Inventar nicht (mehr) vorhanden.'
                : 'Im CoE-Inventar gibt es keine Umgebung vom Typ „Default“. Wähle das Default Environment in den Webpart-Eigenschaften aus.'}
          </p>
          {defaultEnv.error && <p className={`${styles.stateText} ${styles.mono}`}>{defaultEnv.error}</p>}
          {isEditMode && (
            <div className={styles.stateActions}>
              <PrimaryButton text="Eigenschaften öffnen" iconProps={{ iconName: 'Settings' }} onClick={props.onConfigure} />
            </div>
          )}
        </div>
      );
    }

    const env = defaultEnv.env;
    const linkId = environmentLinkId(env);
    const dataverseUrl = (props.defaultEnvironmentDataverseUrl || '').trim() || env.dataverseUrl;
    const autoDetected = !(props.defaultEnvironmentId || '').trim();

    return (
      <>
        <div className={styles.envBanner}>
          <div className={styles.envBannerInfo}>
            <div className={styles.envBannerIcon}><Icon iconName="Home" /></div>
            <div style={{ minWidth: 0, flex: '1 1 auto' }}>
              <div className={styles.envBannerEyebrow}>
                Default Environment{autoDetected ? ' · automatisch erkannt' : ''}
              </div>
              <div className={styles.envBannerTitle}>{env.name}</div>
              <div className={styles.envBannerMeta}>
                <span className={styles.envChip} title="Umgebungs-ID">{linkId}</span>
                {dataverseUrl && <span className={styles.envChip} title="Dataverse">{hostOf(dataverseUrl)}</span>}
              </div>
            </div>
            <a
              className={styles.envBannerLink}
              href={`https://admin.powerplatform.microsoft.com/environments/${encodeURIComponent(linkId)}/hub`}
              target="_blank"
              rel="noopener noreferrer"
            >
              Im Admin Center öffnen <Icon iconName="OpenInNewWindow" />
            </a>
          </div>
          <div className={styles.envStats} role="tablist" aria-label="Default Environment">
            {defaultTables.map(t => {
              const k = kpis[t.key];
              const active = t.key === activeKey;
              return (
                <button
                  key={t.key}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  className={`${styles.envStat} ${active ? styles.envStatActive : ''}`}
                  onClick={() => setActiveKey(t.key)}
                >
                  <span className={styles.envStatIcon}><Icon iconName={t.icon || 'Table'} /></span>
                  <span className={styles.envStatText}>
                    <span className={styles.envStatLabel}>{t.title}</span>
                    {!k || k.status === 'loading'
                      ? <span className={styles.skeleton} style={{ width: 48, height: 18, marginTop: 3 }} />
                      : <span className={styles.envStatValue}>{k.status === 'error' || k.total === undefined ? '–' : formatCount(k.total, k.totalAtLeast)}</span>}
                  </span>
                </button>
              );
            })}
          </div>
        </div>
        {renderBody()}
      </>
    );
  };

  /** Wechselt den Reiter und setzt dort Filter (z. B. aus der Übersicht). */
  const navigate = (tableKey: string, nextFilters?: { [field: string]: IColumnFilter }): void => {
    pendingSearchRef.current = undefined;
    if (tableKey === activeKey) {
      setFilters(nextFilters || {});
      setSearch('');
      setPage(0);
      return;
    }
    pendingFiltersRef.current = nextFilters;
    setActiveKey(tableKey);
  };

  /** Konfigurierte Tabelle zu einem logischen Namen (für das Durchklicken). */
  const tableForLogical = (logical: string): ITableConfig | undefined =>
    tables.filter(t => t.logicalName === logical && !t.environmentUrl)[0] || pipelineTables.filter(t => t.logicalName === logical)[0];

  /** Verknüpften Datensatz in der Detailansicht öffnen (vollständig nachladen, dann „durchklicken“). */
  const openRelated = (logical: string, id: string): void => {
    const t = tableForLogical(logical);
    const svc = serviceFor(t);
    if (!t || !svc) {
      return;
    }
    svc.getEntityMeta(t.logicalName)
      .then(meta => svc.getRecords({ ...t, key: `_open_${t.key}`, strictFilter: true, filter: `${meta.primaryId} eq ${id}` }, 1)
        .then(res => {
          if (res.records.length) {
            setPanelStack(prev => prev.concat([{ table: t, record: res.records[0], meta: res.meta }]));
          }
        }))
      .catch(() => undefined);
  };

  // Datensatz aus einem geteilten Link einmalig öffnen
  React.useEffect(() => {
    const r = initialViewRef.current && initialViewRef.current.r;
    if (!service || !r || !r.l || !r.id) {
      return;
    }
    initialViewRef.current = { ...initialViewRef.current, r: undefined };
    openRelated(r.l, r.id);
  }, [service]);

  // Aktuelle Ansicht in die Adresse schreiben (teilbarer Link)
  React.useEffect(() => {
    const f: { [field: string]: IColumnFilter } = {};
    activeFilterFields.forEach(field => { f[field] = filters[field]; });
    const top = panelStack.length ? panelStack[panelStack.length - 1] : undefined;
    const topId = top && top.meta ? String(top.record[top.meta.primaryId] || '') : '';
    writeViewState({
      t: activeKey,
      f,
      q: search,
      r: top && topId && !top.table.virtual ? { l: top.table.logicalName, id: topId } : undefined
    });
  }, [activeKey, filters, search, panelStack]);

  /** Datensatz aus Übersicht, Sicherheit oder Suche in der Detailansicht öffnen. */
  const openRecord = (t: ITableConfig, record: IRecord, primaryId: string, meta?: IEntityMeta): void => setPanelStack([{
    table: t, record,
    meta: meta || { logicalName: t.logicalName, entitySetName: '', primaryId, primaryName: (t.columns[0] && t.columns[0].field) || '', displayName: t.title }
  }]);

  /** Reiter wechseln und dort suchen (aus der Suche über alle Bereiche). */
  const navigateSearch = (tableKey: string, query: string): void => {
    if (tableKey === activeKey) {
      setFilters({});
      setSearch(query);
      setPage(0);
      return;
    }
    pendingFiltersRef.current = undefined;
    pendingSearchRef.current = query;
    setActiveKey(tableKey);
  };

  const inOverview = props.showOverview && activeKey === OVERVIEW_KEY;
  const inSecurity = props.showSecurity && activeKey === SECURITY_KEY;
  const devState = data[DEV_ENV_KEY];
  const devAffected = devState && devState.status === 'ready'
    ? {
      count: devState.records.length,
      empty: devState.records.filter(r => r[DEV_FIELDS.apps] === 0 && r[DEV_FIELDS.flows] === 0).length,
      disabled: devState.records.filter(r => r[DEV_FIELDS.state] === 'Deaktiviert').length
    }
    : undefined;

  const renderMain = (): JSX.Element => (inOverview && service
    ? (
      <OverviewView
        service={service}
        tables={tables}
        reloadKey={mainKey}
        devAffected={devAffected}
        showDevEnvironments={props.showDevEnvironments}
        onOpenSecurity={props.showSecurity ? () => setActiveKey(SECURITY_KEY) : undefined}
        onNavigate={navigate}
        onOpenRecord={openRecord}
      />
    )
    : inSecurity && service
      ? <SecurityView service={service} tables={tables} reloadKey={mainKey} onNavigate={navigate} onOpenRecord={openRecord} />
      : inDevView ? renderDevEnvironments() : inDefaultView ? renderDefaultEnvironment() : renderBody());

  /** Reiter „Betroffen ab 15.12.“: Hinweis-Banner mit Frist, Kacheln als Schnellfilter, darunter die Tabelle. */
  const renderDevEnvironments = (): JSX.Element => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const days = Math.round((DEV_ENV_DEADLINES.block.getTime() - today.getTime()) / 86400000);
    const deleteDays = Math.round((DEV_ENV_DEADLINES.delete.getTime() - today.getTime()) / 86400000);
    const countdown = days > 1 ? `noch ${formatNumber(days)} Tage` : days === 1 ? 'noch 1 Tag' : days === 0 ? 'heute'
      : deleteDays > 0 ? `gesperrt – Löschung in ${formatNumber(deleteDays)} Tagen` : 'Löschung läuft';
    const devState = data[DEV_ENV_KEY];
    const ready = !!devState && devState.status === 'ready';
    const rows = ready ? devState.records : [];
    const isFilter = (field: string, value: string): boolean => !!filters[field] && !!filters[field].values && filters[field].values!.length === 1 && filters[field].values![0] === value;
    const allActive = activeFilterFields.length === 0 && !search;
    const byRec = (rec: string, label: string, icon: string): { key: string; label: string; value: number; icon: string; active: boolean; onClick: () => void } => ({
      key: rec, label, icon,
      value: rows.filter(r => r[DEV_FIELDS.recommendation] === rec).length,
      active: isFilter(DEV_FIELDS.recommendation, rec) && activeFilterFields.length === 1,
      onClick: () => { setFilters({ [DEV_FIELDS.recommendation]: { values: [rec] } }); setSearch(''); setPage(0); }
    });
    const tiles = [
      { key: 'all', label: 'Betroffen', value: rows.length, icon: 'Warning', active: allActive,
        onClick: () => { setFilters({}); setSearch(''); setPage(0); } },
      byRec(RECOMMENDATION.inUse, 'In Nutzung', 'Rocket'),
      byRec(RECOMMENDATION.inactive, 'Inaktiv > 90 Tage', 'Clock'),
      byRec(RECOMMENDATION.empty, 'Leer', 'Inbox'),
      byRec(RECOMMENDATION.disabled, 'Deaktiviert', 'Blocked2')
    ];
    return (
      <>
        <div className={`${styles.envBanner} ${styles.devBanner}`}>
          <div className={styles.envBannerInfo}>
            <div className={`${styles.envBannerIcon} ${styles.devBannerIcon}`}><Icon iconName="Warning" /></div>
            <div style={{ minWidth: 0, flex: '1 1 auto' }}>
              <div className={styles.envBannerEyebrow}>Microsoft-Änderung · Developer-Environments</div>
              <div className={styles.envBannerTitle}>Ab 15.12.2026 werden nicht verwaltete Developer-Environments gesperrt</div>
              <div className={styles.devBannerText}>
                Ab 15.01.2027 werden sie gelöscht. Managed Developer-Environments sind nicht betroffen.
                Ein Environment bleibt erhalten, wenn es auf Managed umgestellt wird – dann braucht aber jeder,
                der dort Apps nutzt, eine Premium-Lizenz (ab Februar 2027 werden Apps für User ohne Lizenz blockiert).
              </div>
            </div>
            <span className={styles.devCountdown} title="Sperre am 15.12.2026">
              <Icon iconName="Clock" /> {countdown}
            </span>
          </div>
          <div className={`${styles.envStats} ${styles.devStats}`} role="group" aria-label="Schnellfilter">
            {tiles.map(t => (
              <button
                key={t.key}
                type="button"
                aria-pressed={t.active}
                className={`${styles.envStat} ${t.active ? styles.envStatActive : ''}`}
                onClick={t.onClick}
              >
                <span className={styles.envStatIcon}><Icon iconName={t.icon} /></span>
                <span className={styles.envStatText}>
                  <span className={styles.envStatLabel}>{t.label}</span>
                  {ready
                    ? <span className={styles.envStatValue}>{formatNumber(t.value)}</span>
                    : <span className={styles.skeleton} style={{ width: 48, height: 18, marginTop: 3 }} />}
                </span>
              </button>
            ))}
          </div>
        </div>
        {renderBody()}
      </>
    );
  };

  const renderBody = (): JSX.Element => {
    if (!activeTable) {
      return (
        <div className={styles.state}>
          <div className={styles.stateIcon}><Icon iconName="Table" /></div>
          <p className={styles.stateTitle}>Keine Tabellen konfiguriert</p>
          <p className={styles.stateText}>Öffne den Tabellen-Editor, um Dataverse-Tabellen hinzuzufügen.</p>
          {isEditMode && <div className={styles.stateActions}><PrimaryButton text="Tabellen verwalten" onClick={() => setEditorOpen(true)} /></div>}
        </div>
      );
    }
    if (activeTable.unavailable) {
      return (
        <div className={styles.state}>
          <div className={styles.stateIcon}><Icon iconName="Info" /></div>
          <p className={styles.stateTitle}>{activeTable.title} nicht verfügbar</p>
          <p className={styles.stateText}>{activeTable.unavailable}</p>
          {isEditMode && (
            <div className={styles.stateActions}>
              <DefaultButton text="Eigenschaften öffnen" iconProps={{ iconName: 'Settings' }} onClick={props.onConfigure} />
            </div>
          )}
        </div>
      );
    }
    if (state && state.status === 'error') {
      return (
        <div className={styles.state}>
          <div className={`${styles.stateIcon} ${styles.stateIconDanger}`}><Icon iconName="ErrorBadge" /></div>
          <p className={styles.stateTitle}>{activeTable.title} konnten nicht geladen werden</p>
          <p className={styles.stateText}>{state.errorHint}</p>
          <p className={`${styles.stateText} ${styles.mono}`}>{state.error}</p>
          <div className={styles.stateActions}>
            <DefaultButton text="Erneut versuchen" iconProps={{ iconName: 'Refresh' }} onClick={() => setReloadTick(t => t + 1)} />
          </div>
        </div>
      );
    }
    return (
      <>
        <div className={styles.toolbar}>
          <SearchBox
            placeholder={`${activeTable.title} durchsuchen`}
            value={search}
            onChange={(_e, v) => { setSearch(v || ''); setPage(0); }}
            styles={{ root: { width: 280, borderRadius: 6 } }}
          />
          {activeFilterFields.map(f => (
            <button
              key={f}
              type="button"
              className={styles.filterChip}
              onClick={() => setColumnFilter(f, undefined)}
              title="Filter entfernen"
            >
              <Icon iconName="FilterSolid" className={styles.filterChipIcon} />
              <span className={styles.ellipsis}>{describeFilter(columnByField[f], filters[f])}</span>
              <Icon iconName="Cancel" className={styles.filterChipClose} />
            </button>
          ))}
          {activeFilterCount > 0 && (
            <button type="button" className={styles.textButton} onClick={() => { setFilters({}); setSearch(''); setPage(0); }}>
              Alle zurücksetzen
            </button>
          )}
          <span className={styles.grow} />
          <span className={styles.resultInfo}>
            {loading
              ? 'Wird geladen …'
              : partial
                ? `Lädt … ${formatNumber(records.length)}${serverTotal && serverTotal.total !== undefined ? ` von ${formatCount(Math.min(serverTotal.total, maxRecords), serverTotal.totalAtLeast && serverTotal.total < maxRecords)}` : ''}`
                : `${formatNumber(filtered.length)} von ${formatNumber(records.length)}${state && state.truncated ? '+' : ''}`}
          </span>
          <TooltipHost content={compact ? 'Normale Zeilenhöhe' : 'Kompakte Zeilenhöhe'}>
            <IconButton iconProps={{ iconName: compact ? 'DoubleChevronDown' : 'DoubleChevronUp' }} ariaLabel="Zeilenhöhe" onClick={() => setCompact(!compact)} />
          </TooltipHost>
          <TooltipHost content="Spalten ein-/ausblenden">
            <IconButton iconProps={{ iconName: 'ColumnOptions' }} ariaLabel="Spalten" menuProps={{ items: columnMenu, shouldFocusOnMount: true }} onRenderMenuIcon={() => null} />
          </TooltipHost>
          <TooltipHost content="Als CSV exportieren (Excel)">
            <IconButton iconProps={{ iconName: 'ExcelDocument' }} ariaLabel="Export" onClick={exportCsv} disabled={loading || filtered.length === 0} />
          </TooltipHost>
        </div>

        {state && (state.warnings.length > 0 || state.truncated) && (
          <div className={styles.notices}>
            {state.truncated && (
              <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
                Geladen sind die ersten {formatNumber(records.length)}
                {serverTotal && serverTotal.total !== undefined ? ` von ${formatCount(serverTotal.total, serverTotal.totalAtLeast)} Einträgen` : ' Einträge'}
                {activeTable.defaultSort ? ' (nach der Standardsortierung)' : ''}. Suche und Spaltenfilter berücksichtigen nur diese Einträge.
                Das Limit lässt sich in den Webpart-Eigenschaften erhöhen (bis {formatNumber(50000)}).
              </MessageBar>
            )}
            {state.warnings.length > 0 && (
              <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
                {state.warnings.map(w => <div key={w}>{w}</div>)}
              </MessageBar>
            )}
          </div>
        )}

        <DataTable
          columns={visibleColumns}
          records={pageRows}
          idField={state && state.meta ? state.meta.primaryId : ''}
          loading={loading}
          compact={compact}
          sort={sort}
          filters={filters}
          themeVars={themeVars}
          getFilterSource={field => applyFilters(records, field)}
          onSort={onSort}
          onSetSort={(field, desc) => { setSort({ field, desc }); setPage(0); }}
          onFilterChange={setColumnFilter}
          onOpen={r => activeTable && setPanelStack([{ table: activeTable, record: r, meta: state ? state.meta : undefined }])}
        />

        {!loading && filtered.length === 0 && (
          <div className={styles.state}>
            <div className={styles.stateIcon}><Icon iconName={records.length ? 'SearchIssue' : 'Inbox'} /></div>
            <p className={styles.stateTitle}>{records.length ? 'Keine Treffer' : 'Keine Einträge'}</p>
            <p className={styles.stateText}>
              {records.length ? 'Passe die Suche oder die Filter an.' : 'Diese Tabelle enthält (mit dem aktuellen Filter) keine Datensätze.'}
            </p>
            {records.length > 0 && activeFilterCount > 0 && (
              <div className={styles.stateActions}>
                <DefaultButton text="Alle Filter zurücksetzen" iconProps={{ iconName: 'ClearFilter' }} onClick={() => { setFilters({}); setSearch(''); setPage(0); }} />
              </div>
            )}
          </div>
        )}

        {!loading && filtered.length > 0 && (
          <div className={styles.footer}>
            <span>
              {formatNumber(safePage * pageSize + 1)}–{formatNumber(Math.min(filtered.length, (safePage + 1) * pageSize))} von {formatNumber(filtered.length)}
            </span>
            <span style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <Dropdown
                selectedKey={pageSize}
                options={PAGE_SIZES.map(n => ({ key: n, text: `${n} pro Seite` }))}
                onChange={(_e, o) => { setPageSize(o ? Number(o.key) : 25); setPage(0); }}
                styles={{ root: { width: 130 }, title: { borderRadius: 6, height: 28, lineHeight: 26 }, caretDownWrapper: { height: 28, lineHeight: 28 } }}
              />
              {renderPager()}
            </span>
          </div>
        )}
      </>
    );
  };

  // ---------- Navigation (Reiter bzw. gruppierte Seitenleiste wie in der CoE Admin View) ----------
  interface INavItem { key: string; title: string; icon: string; group: string; count?: string; active: boolean; danger?: boolean; locked?: boolean; onClick: () => void }
  const navItems: INavItem[] = (props.showOverview
    ? [{ key: OVERVIEW_KEY, title: 'Übersicht', icon: 'ViewDashboard', group: START_GROUP, active: !!inOverview, onClick: () => setActiveKey(OVERVIEW_KEY) }]
    : [] as INavItem[]).concat(tables.map(t => ({
    key: t.key, title: t.title, icon: t.icon || 'Table', group: t.group || DEFAULT_GROUP,
    count: countFor(t), active: t.key === activeKey, onClick: () => setActiveKey(t.key)
  })));
  if (props.showDefaultEnvironment) {
    navItems.push({ key: '_default', title: 'Default Environment', icon: 'Home', group: ANALYSIS_GROUP, active: !!inDefaultView, onClick: () => setActiveKey(lastDefaultKey) });
  }
  devTables.forEach(t => navItems.push({
    key: t.key, title: t.title, icon: 'Warning', group: ANALYSIS_GROUP, count: countFor(t), active: !!inDevView, danger: true, onClick: () => setActiveKey(t.key)
  }));
  if (props.showSecurity) {
    navItems.push({ key: SECURITY_KEY, title: 'Sicherheit & Freigaben', icon: 'Shield', group: ANALYSIS_GROUP, active: !!inSecurity, onClick: () => setActiveKey(SECURITY_KEY) });
  }
  pipelineTables.forEach(t => navItems.push({
    key: t.key, title: t.title, icon: isLocked(t.key) ? 'Lock' : t.icon, group: GROUP_ALM, count: countFor(t),
    active: t.key === activeKey, locked: isLocked(t.key), onClick: () => setActiveKey(t.key)
  }));
  const navGroups: { name: string; items: INavItem[] }[] = [];
  navItems.forEach(n => {
    let g = navGroups.filter(x => x.name === n.group)[0];
    if (!g) {
      g = { name: n.group, items: [] };
      navGroups.push(g);
    }
    g.items.push(n);
  });
  // Seitenleiste, sobald es mehr Einträge gibt, als in eine Reiterzeile passen (oder Gruppen konfiguriert sind)
  const useSideNav = navItems.length > MAX_TABS || pipelineTables.length > 0 || tables.some(t => !!t.group && t.group !== DEFAULT_GROUP);

  const renderNavItem = (n: INavItem, side: boolean): JSX.Element => (
    <button
      key={n.key}
      type="button"
      role="tab"
      aria-selected={n.active}
      className={side ? `${styles.sideItem} ${n.active ? styles.sideItemActive : ''}` : `${styles.tab} ${n.active ? styles.tabActive : ''}`}
      onClick={n.onClick}
      title={n.locked ? `${n.title} – kein Zugriff (Adminrolle aktivieren)` : n.title}
    >
      <Icon iconName={n.icon} className={`${side ? styles.sideItemIcon : styles.tabIcon} ${n.danger ? styles.toneDanger : ''}`} />
      <span className={side ? styles.sideItemText : undefined}>{n.title}</span>
      {n.count !== undefined && <span className={`${styles.badge} ${n.danger ? styles.badgeDanger : ''}`}>{n.count}</span>}
    </button>
  );

  const lastLoaded = state && state.loadedAt;

  return (
    <ThemeProvider theme={theme as unknown as PartialTheme} applyTo="none">
      <section ref={rootRef} className={`${styles.root} ${narrow ? styles.narrow : ''}`} style={themeVars}>
        <header className={styles.header}>
          <div className={styles.brand}>
            <div className={styles.brandMark}><Icon iconName="Org" /></div>
            <div style={{ minWidth: 0 }}>
              <h2 className={styles.title}>{props.title || 'CoE Cockpit'}</h2>
              <div className={styles.subtitle}>
                {props.subtitle && <span>{props.subtitle}</span>}
                <span className={styles.envChip} title={envValid ? props.environmentUrl : 'Nicht verbunden'}>
                  <span className={`${styles.envDot} ${envValid ? '' : styles.envDotOff}`} />
                  {envValid ? hostOf(props.environmentUrl) : 'Nicht verbunden'}
                </span>
              </div>
            </div>
          </div>
          {service && (
            <GlobalSearch
              tables={searchTables}
              serviceFor={serviceFor}
              themeVars={themeVars}
              onOpenRecord={(t, record, meta) => openRecord(t, record, meta.primaryId, meta)}
              onShowAll={(t, q) => navigateSearch(t.key, q)}
            />
          )}
          <div className={styles.headerActions}>
            {lastLoaded && <span className={styles.metaText}>Aktualisiert {formatTime(lastLoaded)}</span>}
            {service && (
              <TooltipHost content={linkCopied ? 'Link kopiert' : 'Link zu dieser Ansicht kopieren (inkl. Filter)'}>
                <IconButton
                  iconProps={{ iconName: linkCopied ? 'CheckMark' : 'Link' }}
                  ariaLabel="Link kopieren"
                  onClick={() => copyCurrentLink().then(ok => { if (ok) { setLinkCopied(true); window.setTimeout(() => setLinkCopied(false), 2000); } })}
                />
              </TooltipHost>
            )}
            {service && (
              <TooltipHost content="Daten neu laden">
                <IconButton iconProps={{ iconName: 'Refresh' }} ariaLabel="Aktualisieren" onClick={() => setReloadTick(t => t + 1)} disabled={loading} />
              </TooltipHost>
            )}
            {isEditMode && (
              <DefaultButton text="Tabellen verwalten" iconProps={{ iconName: 'Settings' }} onClick={() => setEditorOpen(true)} styles={{ root: { borderRadius: 6 } }} />
            )}
          </div>
        </header>

        {!service && renderOnboarding()}

        {service && (
          <>
            {renderAccessNotice()}

            {showKpis && !inOverview && !inSecurity && <KpiStrip tables={tables} kpis={kpis} activeKey={activeKey} onSelect={setActiveKey} />}

            {navItems.length > 0 && !useSideNav && (
              <nav className={styles.tabs} role="tablist" aria-label="Tabellen">
                {navItems.map((n, i) => (
                  <React.Fragment key={n.key}>
                    {i > 0 && n.group !== navItems[i - 1].group && (n.group === ANALYSIS_GROUP || n.group === GROUP_ALM) && <span className={styles.tabDivider} aria-hidden="true" />}
                    {renderNavItem(n, false)}
                  </React.Fragment>
                ))}
              </nav>
            )}

            {useSideNav
              ? (
                <div className={styles.layout}>
                  <nav className={styles.sideNav} role="tablist" aria-label="Tabellen" aria-orientation="vertical">
                    {navGroups.map(g => {
                      const isCollapsed = !!collapsedGroups[g.name] && !g.items.some(n => n.active);
                      return (
                        <div key={g.name} className={styles.sideGroup}>
                          <button
                            type="button"
                            className={styles.sideGroupHead}
                            aria-expanded={!isCollapsed}
                            onClick={() => setCollapsedGroups(prev => ({ ...prev, [g.name]: !isCollapsed }))}
                          >
                            <span>{g.name}</span>
                            <Icon iconName={isCollapsed ? 'ChevronDown' : 'ChevronUp'} />
                          </button>
                          {!isCollapsed && g.items.map(n => renderNavItem(n, true))}
                        </div>
                      );
                    })}
                  </nav>
                  <div className={styles.layoutMain}>
                    {renderMain()}
                  </div>
                </div>
              )
              : renderMain()}
          </>
        )}

        <RecordPanel
          record={panel ? panel.record : undefined}
          table={panel ? panel.table : undefined}
          meta={panel ? panel.meta : undefined}
          service={serviceFor(panel ? panel.table : undefined)}
          themeVars={themeVars}
          onDismiss={() => setPanelStack([])}
          canOpen={logical => !!tableForLogical(logical)}
          onOpenRelated={openRelated}
          onBack={panelStack.length > 1 ? () => setPanelStack(prev => prev.slice(0, prev.length - 1)) : undefined}
          backLabel={panelStack.length > 1 ? panelStack[panelStack.length - 2].table.title : undefined}
        />

        <TableEditor
          isOpen={editorOpen}
          tables={tables}
          service={service}
          themeVars={themeVars}
          onDismiss={() => setEditorOpen(false)}
          onSave={t => { props.onTablesChanged(t); setEditorOpen(false); }}
        />
      </section>
    </ThemeProvider>
  );
};

export default CoeCockpit;
