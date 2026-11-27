import * as React from 'react';
import { ThemeProvider, PartialTheme } from '@fluentui/react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { IconButton, DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { Icon } from '@fluentui/react/lib/Icon';
import { AadHttpClientFactory, AadTokenProviderFactory, MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './CoeCockpit.module.scss';
import { IColumnConfig, IRecord, ISortConfig, ITableConfig } from '../models/ITableConfig';
import {
  DataverseError, DataverseService, environmentLinkId, findDefaultEnvironment, ICountResult, IEntityMeta, IEnvironmentInfo, normalizeDataverseUrl
} from '../services/DataverseService';
import { buildDefaultEnvironmentTables, DEFAULT_ENV_PREFIX, isDefaultEnvKey } from '../models/defaultEnvironment';
import { buildDevEnvTable, DEV_ENV_DEADLINES, DEV_ENV_KEY, DEV_FIELDS, loadAffectedEnvironments, NOTIFIED, RECOMMENDATION } from '../models/devEnvironments';
import { OVERVIEW_KEY } from '../models/overview';
import { buildPipelineTables, GROUP_ALM, PIPELINE_KEY_PREFIX } from '../models/pipelines';
import { OverviewView } from './OverviewView';
import { SECURITY_KEY, SecurityView } from './SecurityView';
import { GlobalSearch } from './GlobalSearch';
import { copyText, NotifyDialog } from './NotifyDialog';
import { BulkNotifyPanel } from './BulkNotifyPanel';
import { MailService } from '../services/MailService';
import { describeEntry, INotifyLogMap, NotifyLogService, NotifyMethod } from '../services/NotifyLogService';
import { ActionService, IActionSetupState } from '../services/ActionService';
import { ACTION_ICON, ACTION_LABEL, ACTION_LOG_KEY, actionsFor, IActionItem, resourceOf } from '../models/actions';
import { ActionDialog, IActionTarget } from './ActionDialog';
import { ActionLogView } from './ActionLogView';
import { CopilotCreditsView } from './CopilotCreditsView';
import { SolutionContents } from './SolutionContents';
import { SettingsPanel } from './SettingsPanel';
import { RoleBadge } from './RoleBadge';
import { PremiumView } from './PremiumView';
import { PREMIUM_KEY } from '../models/premiumLicensing';
import { IUsageTarget, UsageNotify, usageLogKey } from './UsageNotify';
import { usageSupported } from '../models/usageMail';
import { IWebPartSettings } from '../models/settings';
import { BILLING_KEY, CREDITS_KEY, GROUP_LICENSING } from '../models/copilotCredits';
import { USER_LICENSES_KEY } from '../models/userLicenses';
import { UserLicensesView } from './UserLicensesView';
import { ORPHANS_KEY } from '../models/orphans';
import { OrphansView } from './OrphansView';
import { LICENSE_OPT_KEY } from '../models/licenseOptimization';
import { LicenseOptimizationView } from './LicenseOptimizationView';
import { STORAGE_KEY } from '../models/storage';
import { StorageView } from './StorageView';
import { MESSAGES_KEY } from '../models/messageCenter';
import { MessageCenterView } from './MessageCenterView';
import { PowerPlatformApiService, PPAPI_RESOURCE } from '../services/PowerPlatformApiService';
import { DirectActionService } from '../services/DirectActionService';
import { checkAdminRole, IRoleStatus } from '../services/RoleCheckService';
import { GROUP_GOVERNANCE } from '../models/defaultTables';
import { INotifySettings } from '../models/notification';
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
  /** Für die Prüfung der aktiven Admin-Rolle (PIM) aus dem Anmelde-Token. */
  aadTokenProviderFactory?: AadTokenProviderFactory;
  /** Für den Mailversand über Microsoft Graph (MSGraphClientV3). */
  msGraphClientFactory: MSGraphClientFactory;
  /** Für das gemeinsame Versandprotokoll (SharePoint-Liste auf dieser Website). */
  spHttpClient: SPHttpClient;
  webUrl: string;
  currentUser: { displayName: string; email: string };
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
  /** Ansichten „Copilot-Credits“ und „Abrechnungspläne“ (live über die Power Platform API) anzeigen. */
  showCopilotCredits: boolean;
  /** Auswertung „Betroffen ab Februar 2027“ (Premium-Lizenz in Managed Environments) anzeigen. */
  showPremium2027?: boolean;
  /** Ansicht „Lizenzen je Benutzer“ (live aus Microsoft Graph, Standard: an) */
  showUserLicenses?: boolean;
  /** SharePoint-App-Leiste links auf dieser Seite ausblenden (Standard: an) */
  hideAppBar?: boolean;
  /** „Verwaiste Ressourcen“ (Standard: an) */
  showOrphans?: boolean;
  /** „Lizenz-Optimierung“ (Standard: an) */
  showLicenseOptimization?: boolean;
  /** „Dataverse-Speicher“ (Standard: an) */
  showStorage?: boolean;
  /** „Was ändert Microsoft?“ (Standard: an) */
  showMessageCenter?: boolean;
  /** Dataverse-URL des Pipelines-Host-Environments (leer = ALM-Pipelines ausblenden). */
  pipelinesHostUrl: string;
  /** Vorlage für die E-Mail an die Maker betroffener Developer-Environments. */
  notify: INotifySettings;
  /** Automatischen Versand über Microsoft Graph (Mail.Send) anbieten. */
  notifyAutoSend: boolean;
  /** Aktionen (Quarantäne, Löschen, Besitzer, Freigaben, Flows ein/aus) anbieten. */
  showActions: boolean;
  /** SharePoint-Gruppe, deren Mitglieder Aktionen auslösen dürfen. */
  actionsGroup: string;
  /** Löschen erst, wenn die App in Quarantäne bzw. der Flow ausgeschaltet ist. */
  deleteRequiresStop: boolean;
  /** Ändert sich nach dem Einrichten von Gruppe/Liste in den Eigenschaften (Stand neu prüfen). */
  actionsVersion?: number;
  onTablesChanged: (tables: ITableConfig[]) => void;
  onConfigure: () => void;
  /** Werte und Rückrufe für den Bereich „Webpart-Einstellungen“ (nur im Bearbeitungsmodus angeboten). */
  settings?: IWebPartSettings;
  /** Zähler: erhöht sich, wenn der Bereich „Webpart-Einstellungen“ geöffnet werden soll (aus dem Eigenschaftenbereich). */
  settingsRequest?: number;
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
/** Felder für die Aktionen je Tabelle (werden für den geöffneten Datensatz direkt nachgeladen). */
const ACTION_FIELDS: { [logicalName: string]: string[] } = {
  admin_app: ['admin_appid', 'admin_appenvironmentid', 'admin_issolutionapp', 'admin_appisquarantined', 'admin_quarantineapp', 'admin_appdeleted'],
  admin_flow: ['admin_flowid', 'admin_flowidstring', 'admin_flowenvironmentid', 'admin_issolutionflow', 'admin_flowstate', 'admin_flowdeleted'],
  admin_solution: ['admin_solutionid', 'admin_solutionenvtobjectguidstring', 'admin_solutionismanaged', 'admin_solutiondeleted', 'admin_name',
    'admin_solutiondisplayname', 'admin_solutionpublisher', '_admin_solutionenvironment_value']
};
/** Bereich für die Übersicht (Startseite). */
const START_GROUP = 'Start';
/** Bereich für berechnete Ansichten (Default Environment, Betroffen ab 15.12.). */
const ANALYSIS_GROUP = 'Auswertungen';
/** Ab so vielen Einträgen wird statt der Reiterzeile die Seitenleiste verwendet. */
const MAX_TABS = 7;
/** Unter dieser Breite (px) wird die schmale Darstellung verwendet. */
const NARROW_WIDTH = 860;
const COLLATOR = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });
const NO_RECORDS: IRecord[] = [];

/** Spalten „Maker verständigt / am / von“ aus dem gemeinsamen Versandprotokoll ergänzen. */
function withNotifyInfo(r: IRecord, log: INotifyLogMap): IRecord {
  const e = log[String(r.admin_environmentid || '').toLowerCase()];
  return { ...r, [DEV_FIELDS.notified]: e ? NOTIFIED.yes : NOTIFIED.no, [DEV_FIELDS.notifiedOn]: e ? e.at : null, [DEV_FIELDS.notifiedBy]: e ? e.by : '' };
}

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
  const ppApi = React.useMemo(() => new PowerPlatformApiService(props.aadHttpClientFactory), [props.aadHttpClientFactory]);
  const [search, setSearch] = React.useState<string>('');
  const [filters, setFilters] = React.useState<{ [field: string]: IColumnFilter }>({});
  const [sort, setSort] = React.useState<ISortConfig | undefined>(undefined);
  const [page, setPage] = React.useState<number>(0);
  const [pageSize, setPageSize] = React.useState<number>(props.pageSize || 25);
  const [compact, setCompact] = React.useState<boolean>(props.density === 'compact');
  /** Geöffnete Datensätze (Detailansicht); mehrere = durchgeklickt, „Zurück“ führt zum vorherigen. */
  const [panelStack, setPanelStack] = React.useState<{ table: ITableConfig; record: IRecord; meta?: IEntityMeta }[]>([]);
  const panel = panelStack.length ? panelStack[panelStack.length - 1] : undefined;
  /**
   * Felder, die die Aktionen brauchen (Solution-Kennzeichen, Quarantäne, Flow-Status, IDs). Sie werden für den geöffneten
   * Datensatz immer direkt geladen – unabhängig davon, ob sie in den (evtl. älteren) gespeicherten Tabelleneinstellungen stehen.
   */
  const [actionExtra, setActionExtra] = React.useState<{ [key: string]: IRecord }>({});
  const [actionExtraTick, setActionExtraTick] = React.useState<number>(0);
  const actionExtraKey = panel && ACTION_FIELDS[panel.table.logicalName] && !panel.table.environmentUrl && !panel.table.virtual
    ? `${panel.table.logicalName}:${String(panel.record[(panel.meta && panel.meta.primaryId) || `${panel.table.logicalName}id`] || '')}`
    : '';
  const actionExtraPending = !!actionExtraKey && !actionExtra[actionExtraKey];
  /** Datensatz für die Aktionen: Listenzeile + direkt geladene Aktionsfelder. */
  const actionRecord = (): IRecord | undefined => (panel ? { ...panel.record, ...(actionExtraKey ? actionExtra[actionExtraKey] || {} : {}) } : undefined);
  const [editorOpen, setEditorOpen] = React.useState<boolean>(false);
  const [settingsOpen, setSettingsOpen] = React.useState<boolean>(false);
  /** „Eigenschaften öffnen“: übersichtliche Webpart-Einstellungen, sonst der SharePoint-Eigenschaftenbereich. */
  const openSettings = (): void => (props.settings ? setSettingsOpen(true) : props.onConfigure());
  /** Environment, für das die E-Mail an die Maker vorbereitet wird. */
  const [notifyRecord, setNotifyRecord] = React.useState<IRecord | undefined>(undefined);
  /** Komponente, deren Besitzer nach der Verwendung gefragt wird („Nachfrage Verwendung“). */
  const [usageTarget, setUsageTarget] = React.useState<IUsageTarget | undefined>(undefined);
  const [addressesCopied, setAddressesCopied] = React.useState<string>('');
  /** Environments für den Serienversand (Panel offen, solange gesetzt). */
  const [bulkRecords, setBulkRecords] = React.useState<IRecord[] | undefined>(undefined);
  const mailService = React.useMemo(
    () => (props.notifyAutoSend ? new MailService(props.msGraphClientFactory) : undefined),
    [props.notifyAutoSend, props.msGraphClientFactory]
  );
  const accentColor = String((themeVars as { [k: string]: unknown })['--coe-accent'] || '#0f6cbd');
  // Cockpit-Aktionen (Aufträge in der SharePoint-Liste, ausgeführt von einem Flow)
  const actionService = React.useMemo(
    // Auch bei ausgeschalteten Aktionen nötig: Mitgliedschaft in der Gruppe entscheidet über alle Aktions-Schaltflächen
    () => (props.spHttpClient && props.webUrl ? new ActionService(props.spHttpClient, props.webUrl) : undefined),
    [props.spHttpClient, props.webUrl]
  );
  const [actionSetup, setActionSetup] = React.useState<IActionSetupState | undefined>(undefined);
  const [actionItems, setActionItems] = React.useState<IActionItem[] | undefined>(undefined);
  const [actionLoading, setActionLoading] = React.useState<boolean>(false);
  const [actionProblem, setActionProblem] = React.useState<string>('');
  const [actionTarget, setActionTarget] = React.useState<IActionTarget | undefined>(undefined);
  const refreshActions = React.useCallback((): void => {
    if (!actionService) {
      return;
    }
    setActionLoading(true);
    actionService.list().then(
      items => { setActionItems(items || []); setActionProblem(items ? '' : 'Die Liste „CoE Cockpit Aktionen“ gibt es noch nicht – in den Webpart-Eigenschaften einrichten.'); setActionLoading(false); },
      (e: unknown) => { setActionProblem(`Aktionsprotokoll nicht lesbar: ${e instanceof Error ? e.message : String(e)}`); setActionLoading(false); }
    );
  }, [actionService]);
  /** Mitglied der SharePoint-Gruppe (Standard „CoE Cockpit Admins“)? Nur dann gibt es Aktions-Schaltflächen – auch für E-Mails. */
  const isGroupMember = !!actionSetup && !!actionSetup.group && actionSetup.isMember;
  // Admin-Rolle (PIM) der angemeldeten Person – Anzeige rechts oben und Voraussetzung für alle Admin-Aktionen
  const [roleStatus, setRoleStatus] = React.useState<IRoleStatus | undefined>(undefined);
  const [roleChecking, setRoleChecking] = React.useState<boolean>(!!props.aadTokenProviderFactory);
  const roleResources = [service ? service.baseUrl : '', 'https://graph.microsoft.com', PPAPI_RESOURCE];
  const checkRole = React.useCallback((fresh: boolean): void => {
    if (!props.aadTokenProviderFactory) {
      return;
    }
    setRoleChecking(true);
    checkAdminRole(props.aadTokenProviderFactory, roleResources, fresh).then(
      st => { setRoleStatus(st); setRoleChecking(false); },
      (e: Error) => { setRoleStatus({ state: 'unknown', roles: [], checkedAt: new Date(), problem: e.message }); setRoleChecking(false); }
    );
  }, [props.aadTokenProviderFactory, roleResources.join('|')]);
  React.useEffect(() => { checkRole(false); }, [checkRole]);
  /** Rolle „Power Platform Administrator“ bzw. „Globaler Administrator“ gerade aktiv? */
  const adminRoleActive = !!roleStatus && roleStatus.state === 'active';
  /**
   * Admin-Aktionen (Quarantäne, Löschen …) nur für Gruppenmitglieder mit AKTIVER Admin-Rolle, eingerichteter Liste
   * (Aktionsprotokoll) und eingeschalteten Aktionen. Ohne aktive Rolle sind die Schaltflächen ausgeblendet.
   */
  const canAct = !!props.showActions && !!actionService && !!actionSetup && actionSetup.listExists && isGroupMember && adminRoleActive;
  /** Direkte Ausführung im Namen der angemeldeten Person (alles außer „Solution löschen“). */
  const directActions = React.useMemo(() => new DirectActionService(
    props.aadHttpClientFactory,
    ppApi,
    service,
    async (mail: string) => {
      const graph = await props.msGraphClientFactory.getClient('3');
      const u = await graph.api(`/users/${encodeURIComponent(mail)}`).version('v1.0').select('id').get() as { id?: string };
      return u && u.id ? u.id : undefined;
    }
  ), [props.aadHttpClientFactory, ppApi, service, props.msGraphClientFactory]);
  // Gemeinsames Versandprotokoll (wer hat welches Environment wann verständigt)
  const notifyLogService = React.useMemo(
    () => (props.spHttpClient && props.webUrl ? new NotifyLogService(props.spHttpClient, props.webUrl) : undefined),
    [props.spHttpClient, props.webUrl]
  );
  const [notifyLog, setNotifyLog] = React.useState<INotifyLogMap>({});
  const [notifyLogProblem, setNotifyLogProblem] = React.useState<string>('');
  const refreshNotifyLog = React.useCallback((): Promise<INotifyLogMap> => {
    if (!notifyLogService) {
      return Promise.resolve({});
    }
    return notifyLogService.load().then(
      res => { setNotifyLog(res.map); setNotifyLogProblem(''); return res.map; },
      (e: unknown) => { setNotifyLogProblem(e instanceof Error ? e.message : String(e)); throw e; }
    );
  }, [notifyLogService]);
  const logNotification = (record: IRecord, method: NotifyMethod, subject: string, to: string[]): Promise<void> => {
    if (!notifyLogService) {
      return Promise.reject(new Error('Kein Versandprotokoll verfügbar.'));
    }
    const entry = {
      envId: String(record.admin_environmentid || '').toLowerCase(),
      envName: displayValue(record, 'admin_displayname'),
      at: new Date().toISOString(),
      by: props.currentUser.displayName,
      byEmail: props.currentUser.email,
      method, subject, recipients: to.join('; ')
    };
    return notifyLogService.add(entry).then(() => {
      setNotifyLog(prev => ({ ...prev, [entry.envId]: entry }));
      setNotifyLogProblem('');
    });
  };
  /** „Nachfrage Verwendung“ im gemeinsamen Versandprotokoll vermerken (Schlüssel usage:<Tabelle>:<ID>). */
  const logUsage = (key: string, title: string, method: NotifyMethod, subject: string, to: string[]): Promise<void> => {
    if (!notifyLogService) {
      return Promise.reject(new Error('Kein Versandprotokoll verfügbar.'));
    }
    const entry = {
      envId: key, envName: title, at: new Date().toISOString(),
      by: props.currentUser.displayName, byEmail: props.currentUser.email,
      method, subject, recipients: to.join('; ')
    };
    return notifyLogService.add(entry).then(() => {
      setNotifyLog(prev => ({ ...prev, [key]: entry }));
      setNotifyLogProblem('');
    });
  };
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

  /** Dataverse-Service für eine beliebige Umgebungs-URL (z. B. das Environment einer Solution). */
  const serviceForUrl = (url: string): DataverseService => {
    const u = normalizeDataverseUrl(url) || url;
    if (service && u.toLowerCase() === service.baseUrl.toLowerCase()) {
      return service;
    }
    if (!servicesRef.current[u]) {
      servicesRef.current[u] = new DataverseService(props.aadHttpClientFactory, u);
    }
    return servicesRef.current[u];
  };

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
  // SharePoint-App-Leiste (links: Discover/Publish/Build/OneDrive bzw. Startseite/Meine Websites …) auf dieser Seite ausblenden.
  // Microsoft bietet dafür keinen Schalter mehr – daher per CSS, nur solange das Cockpit auf der Seite ist.
  React.useEffect(() => {
    if (props.hideAppBar === false) {
      return;
    }
    const id = 'coe-cockpit-hide-appbar';
    let style = document.getElementById(id) as HTMLStyleElement | null;
    if (!style) {
      style = document.createElement('style');
      style.id = id;
      style.textContent = '#sp-appBar, #sp-appBar-container, [data-automation-id="sp-appBar"], [data-automationid="sp-appBar"] { display: none !important; }';
      document.head.appendChild(style);
    }
    return () => {
      const el = document.getElementById(id);
      if (el && el.parentNode) {
        el.parentNode.removeChild(el);
      }
    };
  }, [props.hideAppBar]);
  React.useEffect(() => { setCompact(props.density === 'compact'); }, [props.density]);
  React.useEffect(() => { if (props.editorRequest > 0) { setEditorOpen(true); } }, [props.editorRequest]);
  React.useEffect(() => { if ((props.settingsRequest || 0) > 0 && props.settings) { setSettingsOpen(true); } }, [props.settingsRequest]);

  React.useEffect(() => {
    const waitingForDefault = (props.showDefaultEnvironment && isDefaultEnvKey(activeKey)) || (props.showOverview && activeKey === OVERVIEW_KEY)
      || (props.showSecurity && activeKey === SECURITY_KEY) || (props.showActions && activeKey === ACTION_LOG_KEY)
      || (props.showCopilotCredits && (activeKey === CREDITS_KEY || activeKey === BILLING_KEY))
      || (props.showPremium2027 !== false && activeKey === PREMIUM_KEY) || (props.showUserLicenses !== false && activeKey === USER_LICENSES_KEY)
      || (props.showOrphans !== false && activeKey === ORPHANS_KEY) || (props.showLicenseOptimization !== false && activeKey === LICENSE_OPT_KEY)
      || (props.showStorage !== false && activeKey === STORAGE_KEY) || (props.showMessageCenter !== false && activeKey === MESSAGES_KEY);
    if (!activeTable && tables.length && !waitingForDefault) {
      setActiveKey(tables[0].key);
    }
    // activeTable/defaultSig gehören dazu: z. B. wird beim Ausschalten von „Default Environment“
    // der aktive Reiter erst im nächsten Render ungültig.
    if (((!props.showOverview && activeKey === OVERVIEW_KEY) || (!props.showSecurity && activeKey === SECURITY_KEY)
      || (!props.showCopilotCredits && (activeKey === CREDITS_KEY || activeKey === BILLING_KEY))) && tables.length) {
      setActiveKey(props.showOverview ? OVERVIEW_KEY : tables[0].key);
    }
  }, [tablesSig, props.showDefaultEnvironment, props.showDevEnvironments, props.showOverview, props.showSecurity, props.showCopilotCredits, props.pipelinesHostUrl, defaultSig, !!activeTable]);

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
  const rawRecords = state ? state.records : NO_RECORDS;
  const isDevTable = !!activeTable && activeTable.key === DEV_ENV_KEY;
  const records = React.useMemo(
    () => (isDevTable ? rawRecords.map(r => withNotifyInfo(r, notifyLog)) : rawRecords),
    [rawRecords, isDevTable, notifyLog]
  );
  // Aktionen: Einrichtung (Gruppe, Liste, Mitgliedschaft) prüfen und Protokoll laden
  React.useEffect(() => {
    if (!actionService) {
      setActionSetup(undefined);
      return;
    }
    let cancelled = false;
    actionService.getSetupState(props.actionsGroup).then(st => {
      if (cancelled) {
        return;
      }
      setActionSetup(st);
      if (st.listExists) {
        refreshActions();
      }
    }, () => { if (!cancelled) { setActionSetup({ listExists: false, isMember: false }); } });
    return () => { cancelled = true; };
  }, [actionService, props.actionsGroup, props.actionsVersion, reloadTick]);

  // Versandprotokoll laden (beim Start, beim Aktualisieren und beim Öffnen des Reiters)
  React.useEffect(() => {
    if (service && props.showDevEnvironments && isDevTable) {
      refreshNotifyLog().catch(() => undefined);
    }
  }, [service, props.showDevEnvironments, isDevTable, reloadTick, refreshNotifyLog]);
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

  /**
   * Anzahl der angezeigten Einträge – gut sichtbar neben dem Suchfeld.
   * Gefiltert/gesucht: „9 von 571 Einträgen“ (hervorgehoben), sonst „571 Einträge“.
   */
  const renderResultBadge = (): JSX.Element => {
    const narrowed = activeFilterCount > 0;
    const more = state && state.truncated ? '+' : '';
    let content: React.ReactNode;
    if (loading) {
      content = 'Wird geladen …';
    } else if (partial) {
      content = `Lädt … ${formatNumber(records.length)}${serverTotal && serverTotal.total !== undefined ? ` von ${formatCount(Math.min(serverTotal.total, maxRecords), serverTotal.totalAtLeast && serverTotal.total < maxRecords)}` : ''}`;
    } else if (narrowed) {
      content = <><strong>{formatNumber(filtered.length)}</strong> von {formatNumber(records.length)}{more} {records.length === 1 ? 'Eintrag' : 'Einträgen'}</>;
    } else {
      content = <><strong>{formatNumber(records.length)}{more}</strong> {records.length === 1 ? 'Eintrag' : 'Einträge'}</>;
    }
    return (
      <span
        className={`${styles.resultBadge} ${narrowed && !loading && !partial ? styles.resultBadgeActive : ''}`}
        role="status"
        aria-live="polite"
        title={narrowed ? 'Angezeigte Einträge (nach Suche und Filtern) von allen geladenen Einträgen' : 'Geladene Einträge'}
      >
        {narrowed && !loading && !partial && <Icon iconName="Filter" />}
        {content}
      </span>
    );
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
          <PrimaryButton text="Webpart-Einstellungen" iconProps={{ iconName: 'Settings' }} onClick={openSettings} />
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
              <PrimaryButton text="Webpart-Einstellungen" iconProps={{ iconName: 'Settings' }} onClick={openSettings} />
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
  const inActionLog = props.showActions && activeKey === ACTION_LOG_KEY;
  const inCredits = props.showCopilotCredits && (activeKey === CREDITS_KEY || activeKey === BILLING_KEY);
  const inPremium = props.showPremium2027 !== false && activeKey === PREMIUM_KEY;
  const inUserLicenses = props.showUserLicenses !== false && activeKey === USER_LICENSES_KEY;
  const inOrphans = props.showOrphans !== false && activeKey === ORPHANS_KEY;
  const inLicenseOpt = props.showLicenseOptimization !== false && activeKey === LICENSE_OPT_KEY;
  const inStorage = props.showStorage !== false && activeKey === STORAGE_KEY;
  const inMessages = props.showMessageCenter !== false && activeKey === MESSAGES_KEY;
  /** Warum „Verwaiste Ressourcen“ nicht neu zuweisen darf (leer = darf). */
  const orphanActHint = canAct ? '' : !props.showActions ? 'Aktionen sind in den Webpart-Einstellungen ausgeschaltet.'
    : !isGroupMember ? `Neu zuweisen dürfen Mitglieder der SharePoint-Gruppe „${props.actionsGroup}“ mit aktiver Rolle „Power Platform Administrator“.`
      : !adminRoleActive ? 'Neu zuweisen ist ausgeblendet, weil die Rolle „Power Platform Administrator“ gerade nicht aktiv ist (PIM).'
        : 'Die Aktionsliste ist noch nicht eingerichtet (Webpart-Einstellungen → Aktionen).';
  const devState = data[DEV_ENV_KEY];
  const devAffected = devState && devState.status === 'ready'
    ? {
      count: devState.records.length,
      empty: devState.records.filter(r => r[DEV_FIELDS.apps] === 0 && r[DEV_FIELDS.flows] === 0).length,
      disabled: devState.records.filter(r => r[DEV_FIELDS.state] === 'Deaktiviert').length
    }
    : undefined;

  const renderMain = (): JSX.Element => (inMessages
    ? <MessageCenterView graphFactory={props.msGraphClientFactory} reloadKey={String(reloadTick)} />
    : inStorage && service
    ? <StorageView service={service} ppApi={ppApi} reloadKey={String(reloadTick)} onOpenEnvironment={id => openRelated('admin_environment', id)} />
    : inLicenseOpt && service
    ? <LicenseOptimizationView service={service} graphFactory={props.msGraphClientFactory} reloadKey={String(reloadTick)} />
    : inOrphans && service
    ? (
      <OrphansView
        service={service}
        maxRecords={maxRecords}
        reloadKey={String(reloadTick)}
        canAct={canAct}
        actHint={orphanActHint}
        actions={actionService}
        direct={directActions}
        requestedBy={props.currentUser.displayName}
        onOpen={(logical, id) => openRelated(logical, id)}
        onActionsChanged={refreshActions}
      />
    )
    : inUserLicenses
    ? <UserLicensesView graphFactory={props.msGraphClientFactory} reloadKey={String(reloadTick)} />
    : inPremium && service
    ? (
      <PremiumView
        service={service}
        graphFactory={props.msGraphClientFactory}
        ppApi={props.showCopilotCredits ? ppApi : undefined}
        managedField={props.devEnvManagedField}
        reloadKey={String(reloadTick)}
        onOpenApp={id => openRelated('admin_app', id)}
      />
    )
    : inCredits
    ? (
      <CopilotCreditsView
        api={ppApi}
        mode={activeKey === BILLING_KEY ? 'billing' : 'credits'}
        reloadKey={String(reloadTick)}
        onSwitch={m => setActiveKey(m === 'billing' ? BILLING_KEY : CREDITS_KEY)}
      />
    )
    : inOverview && service
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
      : inActionLog
        ? <ActionLogView items={actionItems} loading={actionLoading} problem={actionProblem || (actionSetup && !actionSetup.listExists ? 'Die Liste „CoE Cockpit Aktionen“ ist noch nicht eingerichtet (Webpart-Eigenschaften → Aktionen).' : '')} onRefresh={refreshActions} />
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
    // Adressen der Maker aller aktuell gefilterten Environments (ohne Doppelte)
    const mailTargets: string[] = [];
    (ready ? filtered : []).forEach(r => String(r[DEV_FIELDS.makerEmails] || '').split(',').map(x => x.trim()).filter(x => !!x)
      .forEach(a => { if (mailTargets.indexOf(a) < 0) { mailTargets.push(a); } }));
    const tiles = [
      { key: 'all', label: 'Betroffen', value: rows.length, icon: 'Warning', active: allActive,
        onClick: () => { setFilters({}); setSearch(''); setPage(0); } },
      byRec(RECOMMENDATION.inUse, 'In Nutzung', 'Rocket'),
      byRec(RECOMMENDATION.inactive, 'Inaktiv > 90 Tage', 'Clock'),
      byRec(RECOMMENDATION.empty, 'Leer', 'Inbox'),
      byRec(RECOMMENDATION.disabled, 'Deaktiviert', 'Blocked2'),
      {
        key: 'notnotified', label: 'Noch nicht verständigt', icon: 'Mail',
        value: rows.filter(r => !notifyLog[String(r.admin_environmentid || '').toLowerCase()]).length,
        active: isFilter(DEV_FIELDS.notified, NOTIFIED.no) && activeFilterFields.length === 1,
        onClick: () => { setFilters({ [DEV_FIELDS.notified]: { values: [NOTIFIED.no] } }); setSearch(''); setPage(0); }
      }
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
          {isGroupMember && (
          <div className={styles.devActions}>
            <span className={styles.devActionsText}>
              <Icon iconName="Mail" /> Maker verständigen: einzeln über das Brief-Symbol in der Zeile oder alle gefilterten auf einmal – je Environment eine eigene E-Mail.
            </span>
            <span className={styles.grow} />
            <PrimaryButton
              text={`Alle gefilterten verständigen (${formatNumber(ready ? filtered.length : 0)})`}
              iconProps={{ iconName: 'Send' }}
              disabled={!ready || filtered.length === 0}
              onClick={() => {
                const list = filtered.slice();
                refreshNotifyLog().catch(() => undefined).then(() => setBulkRecords(list), () => setBulkRecords(list));
              }}
            />
            <DefaultButton
              text={addressesCopied || `E-Mail-Adressen kopieren (${formatNumber(mailTargets.length)})`}
              iconProps={{ iconName: addressesCopied ? 'CheckMark' : 'Copy' }}
              disabled={!ready || mailTargets.length === 0}
              title="Adressen aller Maker der gefilterten Environments kopieren – z. B. für eine Sammelmail in BCC"
              onClick={() => copyText(mailTargets.join('; ')).then(ok => {
                setAddressesCopied(ok ? `${formatNumber(mailTargets.length)} Adressen kopiert` : 'Kopieren nicht möglich');
                window.setTimeout(() => setAddressesCopied(''), 2500);
              })}
            />
          </div>
          )}
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
              <DefaultButton text="Webpart-Einstellungen" iconProps={{ iconName: 'Settings' }} onClick={openSettings} />
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
          {renderResultBadge()}
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
          rowAction={inDevView && isGroupMember ? {
            icon: 'Mail',
            label: 'Maker per E-Mail verständigen',
            isDisabled: r => !r[DEV_FIELDS.makerEmails],
            onClick: r => { refreshNotifyLog().catch(() => undefined); setNotifyRecord(r); }
          } : undefined}
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

  /** Schaltflächen im Bereich „Aktionen“: Mail + Environment (Betroffen ab 15.12.) bzw. Apps und Flows. */
  const panelActions = (): { key: string; text: string; icon: string; disabled?: boolean; title?: string; onClick: () => void }[] | undefined => {
    if (!panel) {
      return undefined;
    }
    if (!isGroupMember) {
      return undefined;
    }
    const out: { key: string; text: string; icon: string; disabled?: boolean; title?: string; onClick: () => void }[] = [];
    const isDev = panel.table.key === DEV_ENV_KEY;
    // „Nachfrage Verwendung“: Mail an den Besitzer (für alle Inventar-Bereiche, unabhängig von der Aktionsliste)
    if (!panel.table.virtual && usageSupported(panel.table.logicalName)) {
      const pid = (panel.meta && panel.meta.primaryId) || `${panel.table.logicalName}id`;
      const asked = notifyLog[usageLogKey(panel.table.logicalName, String(panel.record[pid] || ''))];
      out.push({
        key: 'usage', text: 'Nachfrage Verwendung', icon: 'Mail',
        title: asked ? `Bereits nachgefragt – ${describeEntry(asked)}` : 'E-Mail an den Besitzer: Wird die Komponente noch benötigt?',
        onClick: () => { refreshNotifyLog().catch(() => undefined); setUsageTarget({ table: panel.table, record: panel.record, primaryId: pid }); }
      });
    }
    if (isDev) {
      out.push({
        key: 'mail', text: 'Maker per E-Mail verständigen', icon: 'Mail',
        disabled: !panel.record[DEV_FIELDS.makerEmails],
        title: panel.record[DEV_FIELDS.makerEmails] ? undefined : 'Im CoE ist kein Maker mit E-Mail-Adresse erfasst.',
        onClick: () => { refreshNotifyLog().catch(() => undefined); setNotifyRecord(panel.record); }
      });
    }
    // Environment-Aktionen nur in der Auswertung „Betroffen ab 15.12.“, App/Flow-Aktionen in Apps und Flows
    const rec = actionRecord() || panel.record;
    const res = canAct && (isDev || panel.table.logicalName !== 'admin_environment') ? resourceOf(panel.table.logicalName, rec) : undefined;
    if (res && rec.admin_appdeleted !== true && rec.admin_flowdeleted !== true && rec.admin_solutiondeleted !== true) {
      const pending = (actionItems || []).some(i => i.resourceId === res.id && (i.status === 'Angefordert' || i.status === 'In Arbeit'));
      // Environment-Aktionen setzt Microsoft im Hintergrund um, das CoE zeigt sie erst nach dem nächsten Sync:
      // dieselbe Aktion deshalb 60 Minuten nach „Erledigt“ nicht noch einmal anbieten.
      const recentEnv = (code: string): IActionItem | undefined => res.type !== 'environment' ? undefined : (actionItems || []).filter(i =>
        i.resourceId === res.id && i.action === code && i.status === 'Erledigt' && Date.now() - new Date(i.completedAt || i.created).getTime() < 3600000)[0];
      const fmtTime = (iso: string): string => new Intl.DateTimeFormat('de-AT', { hour: '2-digit', minute: '2-digit' }).format(new Date(iso));
      actionsFor(res.type, rec, props.deleteRequiresStop).forEach(a => {
        const recent = recentEnv(a.code);
        out.push({
          key: a.code,
          text: ACTION_LABEL[a.code].replace(' (Benutzer)', '').replace(' (Mitbesitzer)', ''),
          icon: ACTION_ICON[a.code],
          disabled: !!a.disabled || pending || actionExtraPending || !!recent,
          title: actionExtraPending ? 'Daten für die Aktionen werden geladen …' : pending ? 'Für diese Ressource läuft bereits ein Auftrag.'
            : recent ? `Bereits um ${fmtTime(recent.completedAt || recent.created)} ausgeführt (${recent.requestedBy}) – Microsoft setzt die Änderung um, das CoE-Inventar zeigt sie nach dem nächsten Sync.`
              : a.disabled,
          onClick: () => setActionTarget({ logicalName: panel.table.logicalName, record: rec, code: a.code })
        });
      });
    }
    return out.length ? out : undefined;
  };

  /** Warum im Bereich „Aktionen“ nichts angeboten wird. */
  const panelActionsNote = (): string => {
    if (!isGroupMember) {
      return isEditMode
        ? `Aktionen sind nur für Mitglieder der SharePoint-Gruppe „${props.actionsGroup}“ sichtbar (Webpart-Einstellungen → Aktionen). Andere sehen diesen Abschnitt nicht.`
        : '';
    }
    if (!panel || !(panel.table.logicalName === 'admin_app' || panel.table.logicalName === 'admin_flow' || panel.table.logicalName === 'admin_solution' || panel.table.key === DEV_ENV_KEY)) {
      return 'Für diesen Bereich gibt es keine Aktionen.';
    }
    if (!props.showActions) {
      return 'Aktionen sind in den Webpart-Eigenschaften ausgeschaltet.';
    }
    if (!actionSetup) {
      return 'Berechtigung wird geprüft …';
    }
    if (!actionSetup.listExists) {
      return 'Aktionen sind noch nicht eingerichtet: in den Webpart-Eigenschaften unter „Aktionen für Apps und Flows“ Gruppe anlegen und Aktionsliste einrichten.';
    }
    if (!actionSetup.isMember) {
      return `Aktionen dürfen nur Mitglieder der SharePoint-Gruppe „${props.actionsGroup}“ auslösen.`;
    }
    if (!adminRoleActive) {
      return roleNote();
    }
    return 'Für diesen Datensatz gibt es keine Aktionen.';
  };

  /** Hinweis für Gruppenmitglieder, wenn die Admin-Aktionen wegen der Rolle ausgeblendet sind. */
  const roleNote = (): string => roleChecking && !roleStatus
    ? 'Die Admin-Rolle wird geprüft …'
    : 'Die Aktionen (Quarantäne, Löschen, Besitzer, Freigaben …) sind ausgeblendet, weil die Rolle „Power Platform Administrator“ gerade nicht aktiv ist. '
      + 'In PIM aktivieren und danach rechts oben beim Rollen-Abzeichen „Erneut prüfen“.';

  /** Zusätzlicher Hinweis unter den (Mail-)Schaltflächen, wenn Admin-Aktionen nur wegen der Rolle fehlen. */
  const panelRoleInfo = (): string | undefined => {
    if (!panel || adminRoleActive || !isGroupMember || !props.showActions || !actionSetup || !actionSetup.listExists) {
      return undefined;
    }
    const actionable = panel.table.logicalName === 'admin_app' || panel.table.logicalName === 'admin_flow' || panel.table.logicalName === 'admin_solution' || panel.table.key === DEV_ENV_KEY;
    return actionable ? roleNote() : undefined;
  };

  // Aktionsfelder des geöffneten Datensatzes direkt laden (alle Spalten, damit ältere CoE-Versionen ohne einzelne Felder nicht scheitern)
  React.useEffect(() => {
    if (!actionExtraKey || !service || !props.showActions) {
      return;
    }
    let cancelled = false;
    const sep = actionExtraKey.indexOf(':');
    const logical = actionExtraKey.substring(0, sep);
    const id = actionExtraKey.substring(sep + 1);
    const done = (rec: IRecord): void => {
      if (!cancelled) {
        setActionExtra(prev => ({ ...prev, [actionExtraKey]: rec }));
      }
    };
    if (!id) {
      done({});
      return;
    }
    service.getRecord(logical, id).then(async full => {
      const picked: IRecord = {};
      ACTION_FIELDS[logical].forEach(f => {
        if (f in full) {
          picked[f] = full[f];
        }
      });
      // Solution: Dataverse-URL des Environments (für „Solution löschen“)
      const envId = logical === 'admin_solution' ? String(full._admin_solutionenvironment_value || '') : '';
      if (envId) {
        try {
          const env = await service.getRecord('admin_environment', envId);
          picked.__envurl = String(env.admin_environmentcdsinstanceurl || '').replace(/\/+$/, '');
        } catch {
          // ohne URL bleibt „Solution löschen“ gesperrt
        }
      }
      done(picked);
    }, () => done({}));
    return () => { cancelled = true; };
  }, [actionExtraKey, service, props.showActions, actionExtraTick]);

  /** Geöffneten Datensatz neu laden (nach einer Aktion); gelöschte schließen das Panel. */
  const refreshPanelRecord = (finished?: IActionItem): void => {
    const top = panel;
    if (actionExtraKey) {
      setActionExtra(prev => {
        const next = { ...prev };
        delete next[actionExtraKey];
        return next;
      });
      setActionExtraTick(t => t + 1);
    }
    if (top && top.table.virtual) {
      // Berechnete Tabelle (Betroffen ab 15.12.): Zustand lokal nachführen, Tabelle neu laden
      if (finished && finished.status === 'Erledigt') {
        const patchFields: IRecord = finished.action === 'envmanaged' ? { [DEV_FIELDS.managed]: 'Ja' }
          : finished.action === 'envdisable' ? { [DEV_FIELDS.state]: 'Deaktiviert' }
            : finished.action === 'envenable' ? { [DEV_FIELDS.state]: 'Aktiv' } : {};
        setPanelStack(prev => prev.map(e => (e.record === top.record ? { ...e, record: { ...e.record, ...patchFields } } : e)));
        setReloadTick(t => t + 1);
      }
      return;
    }
    const svc = top ? serviceFor(top.table) : undefined;
    if (!top || !svc || !top.meta || !top.meta.primaryId) {
      return;
    }
    const id = String(top.record[top.meta.primaryId] || '');
    svc.getRecords({ ...top.table, key: `_reopen_${top.table.key}`, strictFilter: true, filter: `${top.meta.primaryId} eq ${id}` }, 1).then(r => {
      setPanelStack(prev => {
        if (!prev.length || prev[prev.length - 1].record !== top.record) {
          return prev;
        }
        return r.records.length
          ? prev.slice(0, prev.length - 1).concat([{ ...top, record: r.records[0] }])
          : prev.slice(0, prev.length - 1);
      });
      setReloadTick(t => t + 1);
    }, () => undefined);
  };

  /** Letzte Aufträge zur geöffneten App bzw. zum geöffneten Flow. */
  const panelActionHistory = (): JSX.Element | undefined => {
    const rec = actionRecord();
    const res = panel && rec && canAct ? resourceOf(panel.table.logicalName, rec) : undefined;
    const list = res ? (actionItems || []).filter(i => i.resourceId === res.id).slice(0, 3) : [];
    if (!list.length) {
      return undefined;
    }
    const dt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
    return (
      <div className={styles.actionHistory} style={themeVars}>
        {list.map(i => (
          <div key={i.id} className={styles.actionHistoryRow}>
            <Icon iconName={ACTION_ICON[i.action] || 'Settings'} />
            <span>{ACTION_LABEL[i.action]}{i.targetName ? ` – ${i.targetName}` : ''}</span>
            <span className={styles.metaText}>{dt.format(new Date(i.created))} · {i.requestedBy}</span>
            <span className={`${styles.actionHistoryStatus} ${i.status === 'Erledigt' ? styles.bulkOk : i.status === 'Fehler' || i.status === 'Abgelehnt' ? styles.bulkErr : styles.bulkInfo}`}>{i.status}</span>
          </div>
        ))}
      </div>
    );
  };

  // ---------- Navigation (Reiter bzw. gruppierte Seitenleiste wie in der CoE Admin View) ----------
  interface INavItem { key: string; title: string; icon: string; group: string; count?: string; active: boolean; danger?: boolean; locked?: boolean; onClick: () => void }
  const navItems: INavItem[] = (props.showOverview
    ? [{ key: OVERVIEW_KEY, title: 'Übersicht', icon: 'ViewDashboard', group: START_GROUP, active: !!inOverview, onClick: () => setActiveKey(OVERVIEW_KEY) }]
    : [] as INavItem[]).concat(props.showMessageCenter !== false
    ? [{ key: MESSAGES_KEY, title: 'Was ändert Microsoft?', icon: 'Megaphone', group: START_GROUP, active: !!inMessages, onClick: () => setActiveKey(MESSAGES_KEY) }]
    : []).concat(tables.map(t => ({
    key: t.key, title: t.title, icon: t.icon || 'Table', group: t.group || DEFAULT_GROUP,
    count: countFor(t), active: t.key === activeKey, onClick: () => setActiveKey(t.key)
  })));
  if (props.showDefaultEnvironment) {
    navItems.push({ key: '_default', title: 'Default Environment', icon: 'Home', group: ANALYSIS_GROUP, active: !!inDefaultView, onClick: () => setActiveKey(lastDefaultKey) });
  }
  devTables.forEach(t => navItems.push({
    key: t.key, title: t.title, icon: 'Warning', group: ANALYSIS_GROUP, count: countFor(t), active: !!inDevView, danger: true, onClick: () => setActiveKey(t.key)
  }));
  if (props.showPremium2027 !== false) {
    navItems.push({ key: PREMIUM_KEY, title: 'Betroffen ab Februar 2027', icon: 'Permissions', group: ANALYSIS_GROUP, active: !!inPremium, danger: true, onClick: () => setActiveKey(PREMIUM_KEY) });
  }
  if (props.showOrphans !== false) {
    navItems.push({ key: ORPHANS_KEY, title: 'Verwaiste Ressourcen', icon: 'UserRemove', group: ANALYSIS_GROUP, active: !!inOrphans, onClick: () => setActiveKey(ORPHANS_KEY) });
  }
  if (props.showSecurity) {
    navItems.push({ key: SECURITY_KEY, title: 'Sicherheit & Freigaben', icon: 'Shield', group: ANALYSIS_GROUP, active: !!inSecurity, onClick: () => setActiveKey(SECURITY_KEY) });
  }
  if (props.showCopilotCredits) {
    navItems.push({ key: CREDITS_KEY, title: 'Copilot-Credits', icon: 'AllCurrency', group: GROUP_LICENSING, active: activeKey === CREDITS_KEY, onClick: () => setActiveKey(CREDITS_KEY) });
    navItems.push({ key: BILLING_KEY, title: 'Abrechnungspläne', icon: 'Bank', group: GROUP_LICENSING, active: activeKey === BILLING_KEY, onClick: () => setActiveKey(BILLING_KEY) });
  }
  if (props.showUserLicenses !== false) {
    navItems.push({ key: USER_LICENSES_KEY, title: 'Lizenzen je Benutzer', icon: 'ContactCard', group: GROUP_LICENSING, active: !!inUserLicenses, onClick: () => setActiveKey(USER_LICENSES_KEY) });
  }
  if (props.showLicenseOptimization !== false) {
    navItems.push({ key: LICENSE_OPT_KEY, title: 'Lizenz-Optimierung', icon: 'Savings', group: GROUP_LICENSING, active: !!inLicenseOpt, onClick: () => setActiveKey(LICENSE_OPT_KEY) });
  }
  if (props.showStorage !== false) {
    navItems.push({ key: STORAGE_KEY, title: 'Dataverse-Speicher', icon: 'Database', group: GROUP_LICENSING, active: !!inStorage, onClick: () => setActiveKey(STORAGE_KEY) });
  }
  if (props.showActions && actionSetup && actionSetup.listExists && isGroupMember) {
    navItems.push({ key: ACTION_LOG_KEY, title: 'Aktionsprotokoll', icon: 'History', group: GROUP_GOVERNANCE, active: !!inActionLog, onClick: () => setActiveKey(ACTION_LOG_KEY) });
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
            {props.aadTokenProviderFactory && (
              <RoleBadge
                status={roleStatus}
                checking={roleChecking}
                onRecheck={() => checkRole(true)}
                actionsHint={isGroupMember && !!props.showActions}
                themeVars={themeVars}
              />
            )}
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
              <DefaultButton text="Tabellen verwalten" iconProps={{ iconName: 'Table' }} onClick={() => setEditorOpen(true)} styles={{ root: { borderRadius: 6 } }} />
            )}
            {isEditMode && (
              <DefaultButton text="Webpart-Einstellungen" iconProps={{ iconName: 'Settings' }} onClick={openSettings} styles={{ root: { borderRadius: 6 } }} />
            )}
          </div>
        </header>

        {!service && renderOnboarding()}

        {service && (
          <>
            {renderAccessNotice()}

            {showKpis && !inOverview && !inSecurity && !inCredits && !inPremium && !inUserLicenses && !inOrphans && !inLicenseOpt && !inStorage && !inMessages && <KpiStrip tables={tables} kpis={kpis} activeKey={activeKey} onSelect={setActiveKey} />}

            {navItems.length > 0 && !useSideNav && (
              <nav className={styles.tabs} role="tablist" aria-label="Tabellen">
                {navItems.map((n, i) => (
                  <React.Fragment key={n.key}>
                    {i > 0 && n.group !== navItems[i - 1].group && (n.group === ANALYSIS_GROUP || n.group === GROUP_ALM || n.group === GROUP_LICENSING) && <span className={styles.tabDivider} aria-hidden="true" />}
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
          onDismiss={ev => {
            // Ein Klick in einen Dialog über dem Panel (Aktion, E-Mail) schließt das Panel nicht
            if (actionTarget || notifyRecord || usageTarget) {
              if (ev && 'preventDefault' in ev) {
                ev.preventDefault();
              }
              return;
            }
            setPanelStack([]);
          }}
          canOpen={logical => !!tableForLogical(logical)}
          onOpenRelated={openRelated}
          onBack={panelStack.length > 1 ? () => setPanelStack(prev => prev.slice(0, prev.length - 1)) : undefined}
          backLabel={panelStack.length > 1 ? panelStack[panelStack.length - 2].table.title : undefined}
          actions={panelActions()}
          headerExtra={panelActionHistory()}
          actionsNote={panelActionsNote()}
          actionsInfo={panelRoleInfo()}
          hideActions={!isGroupMember && !isEditMode}
          extraSection={panel && service && panel.table.logicalName === 'admin_solution' && !panel.table.environmentUrl
            ? {
              title: 'Inhalt der Solution',
              content: (
                <SolutionContents
                  key={String(panel.record[(panel.meta && panel.meta.primaryId) || 'admin_solutionid'])}
                  coe={service}
                  serviceFor={serviceForUrl}
                  solutionId={String(panel.record[(panel.meta && panel.meta.primaryId) || 'admin_solutionid'])}
                  record={panel.record}
                  canOpen={logical => !!tableForLogical(logical)}
                  onOpenRelated={openRelated}
                />
              )
            }
            : undefined}
        />

        <ActionDialog
          target={actionTarget}
          actions={actionService as ActionService}
          direct={directActions}
          tokenFactory={props.aadTokenProviderFactory}
          dataverse={service}
          requestedBy={props.currentUser.displayName}
          themeVars={themeVars}
          onFinished={it => { refreshActions(); refreshPanelRecord(it); }}
          onDismiss={() => { setActionTarget(undefined); refreshActions(); }}
        />

        <NotifyDialog
          record={notifyRecord}
          settings={props.notify}
          themeVars={themeVars}
          accent={accentColor}
          mailService={mailService}
          notified={notifyRecord ? notifyLog[String(notifyRecord.admin_environmentid || '').toLowerCase()] : undefined}
          logProblem={notifyLogProblem}
          onLogged={logNotification}
          onDismiss={() => setNotifyRecord(undefined)}
        />

        <UsageNotify
          target={usageTarget}
          service={usageTarget ? serviceFor(usageTarget.table) : undefined}
          settings={props.notify}
          themeVars={themeVars}
          accent={accentColor}
          mailService={mailService}
          log={notifyLog}
          logProblem={notifyLogProblem}
          onLogged={logUsage}
          onDismiss={() => setUsageTarget(undefined)}
        />

        <BulkNotifyPanel
          records={bulkRecords}
          settings={props.notify}
          accent={accentColor}
          themeVars={themeVars}
          mailService={mailService}
          log={notifyLog}
          logProblem={notifyLogProblem}
          refreshLog={refreshNotifyLog}
          onLogged={logNotification}
          onDismiss={() => setBulkRecords(undefined)}
        />

        <TableEditor
          isOpen={editorOpen}
          tables={tables}
          service={service}
          themeVars={themeVars}
          onDismiss={() => setEditorOpen(false)}
          onSave={t => { props.onTablesChanged(t); setEditorOpen(false); }}
        />

        {props.settings && (
          <SettingsPanel
            isOpen={settingsOpen}
            settings={{
              ...props.settings,
              onOpenTableEditor: () => { setSettingsOpen(false); setEditorOpen(true); },
              onOpenPropertyPane: () => { setSettingsOpen(false); props.settings!.onOpenPropertyPane(); }
            }}
            themeVars={themeVars}
            accent={accentColor}
            onDismiss={() => setSettingsOpen(false)}
          />
        )}
      </section>
    </ThemeProvider>
  );
};

export default CoeCockpit;
