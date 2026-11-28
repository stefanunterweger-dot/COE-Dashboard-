import { IRecord, ITableConfig } from './ITableConfig';
import { getDefaultTables } from './defaultTables';
import { DataverseError, DataverseService, environmentType, IAttributeMeta, IRecordResult } from '../services/DataverseService';
import { displayValue, isMailAddress, LIST_SEP, rawValue } from '../utils/format';

/**
 * Reiter „Betroffen ab 15.12.“: nicht verwaltete (unmanaged) Developer-Environments.
 * Microsoft sperrt sie ab 15.12.2026 und löscht sie ab 15.01.2027. Wird ein Environment auf
 * Managed umgestellt, bleibt es erhalten – App-User brauchen dann ab Februar 2027 eine Premium-Lizenz.
 */
export const DEV_ENV_KEY = 'devenv:affected';

export const DEV_ENV_DEADLINES = {
  block: new Date(2026, 11, 15),   // 15.12.2026 – Sperre
  delete: new Date(2027, 0, 15),   // 15.01.2027 – Löschung
  license: new Date(2027, 1, 1)    // Februar 2027 – Lizenzpflicht in Managed Environments
};

/** Berechnete Spalten (nicht in Dataverse). */
export const DEV_FIELDS = {
  makers: 'coe_makers',
  makerCount: 'coe_makercount',
  /** E-Mail-Adressen der Maker (für die Benachrichtigung). */
  makerEmails: 'coe_makeremails',
  /** Maker ohne E-Mail-Adresse im CoE. */
  makersNoMail: 'coe_makersnomail',
  /** Aus dem gemeinsamen Versandprotokoll (SharePoint-Liste): „Verständigt“ / „Offen“, wann, von wem. */
  notified: 'coe_notified',
  notifiedOn: 'coe_notifiedon',
  notifiedBy: 'coe_notifiedby',
  admins: 'coe_sysadmins',
  state: 'coe_state',
  managed: 'coe_managed',
  apps: 'coe_apps',
  flows: 'coe_flows',
  lastActivity: 'coe_lastactivity',
  recommendation: 'coe_recommendation'
};

/** Empfehlungen je Environment (Anzeige + Schnellfilter). */
export const RECOMMENDATION = {
  inUse: 'In Nutzung – sichern oder auf Managed umstellen',
  inactive: 'Inaktiv (> 90 Tage) – prüfen',
  empty: 'Leer – kann auslaufen',
  disabled: 'Deaktiviert',
  /** Aktivität bzw. Inhalt konnte nicht (vollständig) ermittelt werden – keine Aussage „Inaktiv“ oder „Leer“. */
  unknown: 'Unbekannt – Aktivität nicht ermittelt'
};
const INACTIVE_DAYS = 90;

/** Werte der Spalte „Maker verständigt“. */
export const NOTIFIED = { yes: 'Verständigt', no: 'Offen' };

const ENV_TABLE = 'admin_environment';
const SYSTEM_FIELDS = ['createdby', 'createdonbehalfby', 'modifiedby', 'modifiedonbehalfby', 'ownerid', 'owninguser', 'owningteam', 'owningbusinessunit', 'statecode', 'statuscode'];

/** Tabellenkonfiguration des Reiters (Spalten wie in den anderen Reitern: Suche, Filter, Sortierung, Export). */
export function buildDevEnvTable(): ITableConfig {
  return {
    key: DEV_ENV_KEY,
    title: 'Betroffen ab 15.12.',
    logicalName: ENV_TABLE,
    icon: 'Warning',
    virtual: true,
    showInKpis: false,
    defaultSort: { field: 'admin_displayname', desc: false },
    columns: [
      { field: 'admin_displayname', label: 'Environment', type: 'title' },
      { field: DEV_FIELDS.makers, label: 'Maker', type: 'list' },
      { field: DEV_FIELDS.makerCount, label: 'Anzahl Maker', type: 'number', hidden: true },
      { field: DEV_FIELDS.makerEmails, label: 'Maker (E-Mail)', type: 'list', hidden: true },
      { field: DEV_FIELDS.admins, label: 'System-Administratoren', type: 'list' },
      {
        field: DEV_FIELDS.recommendation, label: 'Empfehlung', type: 'status',
        statusMap: { [RECOMMENDATION.inUse]: 'danger', [RECOMMENDATION.inactive]: 'warning', [RECOMMENDATION.empty]: 'neutral', [RECOMMENDATION.disabled]: 'neutral', [RECOMMENDATION.unknown]: 'warning' }
      },
      { field: DEV_FIELDS.notified, label: 'Maker verständigt', type: 'status', statusMap: { [NOTIFIED.yes]: 'success', [NOTIFIED.no]: 'warning' } },
      { field: DEV_FIELDS.notifiedOn, label: 'Verständigt am', type: 'datetime' },
      { field: DEV_FIELDS.notifiedBy, label: 'Verständigt von', type: 'text' },
      { field: DEV_FIELDS.state, label: 'Zustand', type: 'status', statusMap: { 'Aktiv': 'success', 'Deaktiviert': 'danger', 'Unbekannt': 'neutral' } },
      { field: DEV_FIELDS.apps, label: 'Apps', type: 'number' },
      { field: DEV_FIELDS.flows, label: 'Flows', type: 'number' },
      { field: DEV_FIELDS.lastActivity, label: 'Letzte Aktivität', type: 'date' },
      { field: 'admin_region', label: 'Region', type: 'choice', hidden: true },
      { field: 'admin_environmentcreatedon', label: 'Erstellt', type: 'date' },
      { field: DEV_FIELDS.managed, label: 'Managed', type: 'status', hidden: true, statusMap: { 'Nein': 'warning', 'Ja': 'success', 'Unbekannt': 'neutral' } },
      { field: 'admin_environmentcdsinstanceurl', label: 'Dataverse-URL', type: 'url', hidden: true },
      { field: 'admin_environmentid', label: 'Umgebungs-ID', type: 'text', hidden: true },
      { field: 'admin_environmentname', label: 'Environment-ID (Power Platform)', type: 'text', hidden: true }
    ],
    links: [
      { label: 'Im Power Platform Admin Center öffnen', url: 'https://admin.powerplatform.microsoft.com/environments/{admin_environmentid}/hub' },
      { label: 'In Power Apps öffnen', url: 'https://make.powerapps.com/environments/{admin_environmentid}/home' }
    ],
    // Reiter der Detailansicht wie in der Admin View (Apps, Flows, Bots, … des Environments)
    related: (getDefaultTables().filter(t => t.logicalName === ENV_TABLE)[0] || { related: [] }).related
  };
}

/** Erste passende Spalte; die Muster werden der Reihe nach geprüft (erst logischer Name, dann Anzeigename). */
function findAttribute(attrs: IAttributeMeta[], patterns: RegExp[], types?: string[]): IAttributeMeta | undefined {
  const usable = attrs.filter(a => SYSTEM_FIELDS.indexOf(a.logicalName) < 0 && (!types || types.indexOf(a.type) >= 0));
  for (const p of patterns) {
    const hit = usable.filter(a => p.test(a.logicalName))[0] || usable.filter(a => p.test(a.displayName))[0];
    if (hit) {
      return hit;
    }
  }
  return undefined;
}

/** Spalte mit dem Managed-Status im CoE-Inventar (Name je nach CoE-Version unterschiedlich). */
export function detectManagedAttribute(attrs: IAttributeMeta[]): IAttributeMeta | undefined {
  // CoE: „Governance Protection Level“ (Standard = managed, Basic = nicht managed)
  return findAttribute(attrs, [/protection ?level/i, /ismanaged|managed ?environment/i, /managed|verwaltet/i, /governance/i], ['Boolean', 'Picklist', 'String']);
}

export function detectStateAttribute(attrs: IAttributeMeta[]): IAttributeMeta | undefined {
  // CoE: „Environment Runtime State“ (Enabled/Disabled), sonst „Environment State“ (Ready …)
  return findAttribute(attrs, [/runtime ?state/i, /environment ?state/i, /zustand/i, /disabled|enabled/i], ['Boolean', 'Picklist', 'String']);
}


/**
 * Spalten der Environment-Tabelle. Leer nur, wenn Dataverse die Tabelle nicht kennt (404) – Drosselung, Server-
 * oder Netzwerkfehler werden weitergegeben, sonst gälte ein managed Environment fälschlich als „nicht managed“.
 */
export async function environmentAttributes(service: DataverseService): Promise<IAttributeMeta[]> {
  try {
    return await service.getAttributes(ENV_TABLE);
  } catch (e) {
    if (e instanceof DataverseError && e.status === 404) {
      return [];
    }
    throw e;
  }
}

/** Gewählte Spalte für den Managed-Status (sonst automatisch erkannt). Sind keine Metadaten da, gilt die Wahl direkt. */
export function chooseManagedAttribute(attrs: IAttributeMeta[], managedField: string): IAttributeMeta | undefined {
  const chosen = managedField ? attrs.filter(a => a.selectName === managedField || a.logicalName === managedField)[0] : undefined;
  if (chosen) {
    return chosen;
  }
  if (managedField && attrs.length === 0) {
    return { selectName: managedField, logicalName: managedField.replace(/^_(.+)_value$/, '$1'), displayName: managedField, type: 'String', isPrimaryName: false };
  }
  return detectManagedAttribute(attrs);
}

/** Spalten, die für den Managed-Status in Frage kommen (Auswahl in den Webpart-Eigenschaften). */
export function managedCandidates(attrs: IAttributeMeta[]): IAttributeMeta[] {
  return attrs.filter(a => ['Boolean', 'Picklist', 'String'].indexOf(a.type) >= 0 && SYSTEM_FIELDS.indexOf(a.logicalName) < 0);
}

function isDeveloper(r: IRecord): boolean {
  return /developer|entwickler/i.test(environmentType(r));
}

/** true / false / undefined (unbekannt). */
export function managedValue(r: IRecord, field: string): boolean | undefined {
  if (/protectionlevel/i.test(field)) {
    // CoE „Governance Protection Level“: „Basic“ = nicht managed, leer = unbekannt, sonst (z. B. „Standard“) managed
    const level = rawValue(r, field).trim();
    return !level ? undefined : !/^basic$/i.test(level);
  }
  const raw = r[field];
  if (raw === true || raw === false) {
    return raw;
  }
  const v = displayValue(r, field).trim();
  if (!v) {
    return undefined;
  }
  if (/^(true|yes|ja|managed|verwaltet|standard|enabled|aktiviert)$/i.test(v)) {
    return true;
  }
  if (/^(false|no|nein|unmanaged|nicht verwaltet|basic|none|disabled|deaktiviert)$/i.test(v)) {
    return false;
  }
  return undefined;
}

function stateLabel(r: IRecord, field: string | undefined): string {
  if (!field) {
    return 'Unbekannt';
  }
  const raw = r[field];
  const v = displayValue(r, field);
  if (/disabled|deaktiv|suspend|gesperrt|inactive|inaktiv|stopped/i.test(v) || (raw === true && /disabled/i.test(field)) || (raw === false && /enabled/i.test(field))) {
    return 'Deaktiviert';
  }
  if (/ready|enabled|aktiv|active|bereit|running/i.test(v) || (raw === false && /disabled/i.test(field)) || (raw === true && /enabled/i.test(field))) {
    return 'Aktiv';
  }
  return v || 'Unbekannt';
}

const BATCH = 25;

/**
 * Jüngstes Datum je Environment aus mehreren Spalten (App-Start, App-/Flow-Änderung, Bots, Desktop Flows).
 * incomplete: Environments, bei denen mindestens eine Quelle nicht lesbar war (dann ist „kein Datum“ keine Aussage).
 */
async function lastActivityPerEnvironment(service: DataverseService, envIds: string[]): Promise<{ values: { [envId: string]: string }; incomplete: { [envId: string]: boolean }; errors: string[] }> {
  const out: { [envId: string]: string } = {};
  const incomplete: { [envId: string]: boolean } = {};
  const errors: string[] = [];
  const sources: [string, string, string, string][] = [
    ['admin_app', '_admin_appenvironment_value', 'admin_appdeleted', 'admin_applastlaunchedon'],
    ['admin_app', '_admin_appenvironment_value', 'admin_appdeleted', 'admin_appmodifiedon'],
    ['admin_flow', '_admin_flowenvironment_value', 'admin_flowdeleted', 'admin_flowmodifiedon'],
    ['admin_pva', '_admin_pvaenvironment_value', 'admin_pvadeleted', 'admin_pvalastlaunchedon'],
    ['admin_pva', '_admin_pvaenvironment_value', 'admin_pvadeleted', 'admin_pvamodifiedon'],
    ['admin_rpa', '_admin_desktopflowenvironment_value', 'admin_rpadeleted', 'admin_rpamodifiedon']
  ];
  for (let i = 0; i < envIds.length; i += BATCH) {
    const ids = envIds.slice(i, i + BATCH);
    const results = await Promise.all(sources.map(([table, envField, deletedField, dateField]) => service
      .getGroupedMax(table, envField, `(${deletedField} ne true) and (${ids.map(id => `${envField} eq ${id}`).join(' or ')})`, dateField)
      .catch((e: Error) => {
        ids.forEach(id => { incomplete[id.toLowerCase()] = true; });
        if (errors.length < 2) {
          errors.push(`${table}: ${e.message}`);
        }
        return {} as { [k: string]: string };
      })));
    results.forEach(r => Object.keys(r).forEach(k => {
      if (!out[k] || r[k] > out[k]) {
        out[k] = r[k];
      }
    }));
  }
  return { values: out, incomplete, errors };
}

/** Ressourcen-Tabellen mit Environment- und Besitzer-Spalte (wie die Admin View). */
const MAKER_SOURCES: { table: string; envField: string; ownerField: string; baseFilter: string }[] = [
  { table: 'admin_app', envField: '_admin_appenvironment_value', ownerField: '_admin_appowner_value', baseFilter: 'admin_appdeleted ne true' },
  { table: 'admin_flow', envField: '_admin_flowenvironment_value', ownerField: '_admin_derivedowner_value', baseFilter: 'admin_flowdeleted ne true' },
  { table: 'admin_pva', envField: '_admin_pvaenvironment_value', ownerField: '_admin_pvaowner_value', baseFilter: 'admin_pvadeleted ne true' },
  { table: 'admin_rpa', envField: '_admin_desktopflowenvironment_value', ownerField: '_admin_rpaowner_value', baseFilter: 'statecode eq 0' }
];

/**
 * Maker je Environment: alle Besitzer von Apps, Flows, Bots und Desktop Flows,
 * sortiert nach Anzahl ihrer Ressourcen. „SYSTEM“ wird ausgeblendet.
 */
/** Gültige, zustellbare Adresse (Gäste-UPNs mit #EXT# sind keine Postfächer). */
export function mailAddress(r: IRecord): string {
  const candidates = [displayValue(r, 'admin_useremail'), displayValue(r, 'admin_userprincipalname')].map(x => x.trim());
  return candidates.filter(isMailAddress)[0] || '';
}

/** Kein Mensch: „SYSTEM“ und Anwendungsbenutzer (in Dataverse mit „#“ vor dem Namen, z. B. „# CatalogServiceEur“). */
export function isServiceAccount(name: string): boolean {
  const n = name.trim();
  return /^#/.test(n) || /^system$/i.test(n);
}

interface IMakerResult {
  makers: { [envId: string]: { name: string; email: string; count: number }[] };
  /** Anzahl Bots und Desktop Flows je Environment */
  others: { [envId: string]: number };
  /** Environments, bei denen Bots oder Desktop Flows nicht gezählt werden konnten (dann nie „Leer“) */
  othersUnknown: { [envId: string]: boolean };
}

async function makersPerEnvironment(service: DataverseService, envIds: string[], creators: { [envId: string]: string }, warnings: string[]): Promise<IMakerResult> {
  const perEnv: { [envId: string]: { [makerId: string]: number } } = {};
  const others: { [envId: string]: number } = {};
  const othersUnknown: { [envId: string]: boolean } = {};
  const failed: string[] = [];
  for (let i = 0; i < envIds.length; i += BATCH) {
    const ids = envIds.slice(i, i + BATCH);
    const results = await Promise.all(MAKER_SOURCES.map(src => service
      .getGroupedPairs(src.table, src.envField, src.ownerField, `(${src.baseFilter}) and (${ids.map(id => `${src.envField} eq ${id}`).join(' or ')})`)
      .catch((e: Error) => {
        failed.push(`${src.table}: ${e.message}`);
        if (src.table === 'admin_pva' || src.table === 'admin_rpa') {
          ids.forEach(id => { othersUnknown[id.toLowerCase()] = true; });
        }
        return [] as { a: string; b: string; n: number }[];
      })));
    results.forEach((pairs, k) => pairs.forEach(p => {
      if (!p.a) {
        return;
      }
      const env = p.a.toLowerCase();
      if (MAKER_SOURCES[k].table === 'admin_pva' || MAKER_SOURCES[k].table === 'admin_rpa') {
        others[env] = (others[env] || 0) + p.n;
      }
      if (!p.b) {
        return;
      }
      perEnv[env] = perEnv[env] || {};
      perEnv[env][p.b.toLowerCase()] = (perEnv[env][p.b.toLowerCase()] || 0) + p.n;
    }));
  }
  // Wer das Environment erstellt hat, gehört immer dazu (auch ohne eigene Apps oder Flows)
  Object.keys(creators).forEach(env => {
    const c = creators[env];
    if (c) {
      perEnv[env] = perEnv[env] || {};
      perEnv[env][c] = (perEnv[env][c] || 0) + 1;
    }
  });
  // Namen der Maker nachladen
  const makerIds: string[] = [];
  Object.keys(perEnv).forEach(e => Object.keys(perEnv[e]).forEach(m => { if (makerIds.indexOf(m) < 0) { makerIds.push(m); } }));
  const names: { [makerId: string]: string } = {};
  const emails: { [makerId: string]: string } = {};
  for (let i = 0; i < makerIds.length; i += BATCH) {
    const part = makerIds.slice(i, i + BATCH);
    const res = await service.getRecords({
      key: '_makers', title: 'Maker', logicalName: 'admin_maker', icon: 'People', strictFilter: true,
      filter: part.map(id => `admin_makerid eq ${id}`).join(' or '),
      columns: [
        { field: 'admin_displayname', label: 'Name', type: 'text' },
        { field: 'admin_useremail', label: 'E-Mail', type: 'text' },
        { field: 'admin_userprincipalname', label: 'UPN', type: 'text' }
      ]
    }, part.length).catch((e: Error) => { failed.push(`Maker: ${e.message}`); return undefined; });
    (res ? res.records : []).forEach(r => {
      const id = String(r.admin_makerid || '').toLowerCase();
      names[id] = displayValue(r, 'admin_displayname');
      emails[id] = mailAddress(r);
    });
  }
  const out: { [envId: string]: { name: string; email: string; count: number }[] } = {};
  Object.keys(perEnv).forEach(e => {
    out[e] = Object.keys(perEnv[e])
      .map(m => ({ name: names[m] || '', email: emails[m] || '', count: perEnv[e][m] }))
      .filter(m => !!m.name && !isServiceAccount(m.name))
      .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name, 'de'));
  });
  if (failed.length) {
    warnings.push(`Maker konnten nicht vollständig ermittelt werden – einzelne Environments zeigen deshalb zu wenige Maker (${failed.slice(0, 2).join('; ')}). Bitte „Aktualisieren“.`);
  }
  return { makers: out, others, othersUnknown };
}

/** Rollenvorlage „System Administrator“ (in allen Dataverse-Umgebungen gleich). */
const SYSADMIN_TEMPLATE = '627090ff-40a3-4053-8790-584edc5be201';
const APP_USER_TYPE = 597910002;

/**
 * System-Administratoren je Environment aus der CoE-Tabelle „Environment Security Role Permission“.
 * Ohne „SYSTEM“ und App-Benutzer (Typ „App User“ oder Name mit „#“); doppelte Namen (z. B. direkt + über Gruppe) nur einmal.
 */
async function sysAdminsPerEnvironment(service: DataverseService, envIds: string[]): Promise<{ [envId: string]: string[] }> {
  const out: { [envId: string]: string[] } = {};
  for (let i = 0; i < envIds.length; i += BATCH) {
    const ids = envIds.slice(i, i + BATCH);
    const res = await service.getRecords({
      key: '_sysadmins', title: 'Sicherheitsrollen', logicalName: 'admin_environmentsecurityrolepermission', icon: 'Shield', strictFilter: true,
      filter: `(admin_securityroletemplateid eq '${SYSADMIN_TEMPLATE}' or admin_securityrolename eq 'System Administrator') and (${ids.map(id => `_admin_environment_value eq ${id}`).join(' or ')})`,
      columns: [
        { field: '_admin_environment_value', label: 'Environment', type: 'text' },
        { field: 'admin_name', label: 'Benutzer', type: 'text' },
        { field: 'admin_usertype', label: 'Typ', type: 'text' }
      ]
    }, 50000);
    res.records.forEach(r => {
      const env = String(r._admin_environment_value || '').toLowerCase();
      const name = String(r.admin_name || '').trim();
      if (!env || !name || isServiceAccount(name) || Number(r.admin_usertype) === APP_USER_TYPE) {
        return;
      }
      out[env] = out[env] || [];
      if (out[env].indexOf(name) < 0) {
        out[env].push(name);
      }
    });
  }
  Object.keys(out).forEach(e => out[e].sort((a, b) => a.localeCompare(b, 'de')));
  return out;
}

async function countPerEnvironment(service: DataverseService, table: string, envField: string, deletedField: string, envIds: string[]): Promise<{ [envId: string]: number }> {
  const out: { [envId: string]: number } = {};
  for (let i = 0; i < envIds.length; i += BATCH) {
    const ids = envIds.slice(i, i + BATCH);
    const filter = `(${deletedField} ne true) and (${ids.map(id => `${envField} eq ${id}`).join(' or ')})`;
    const counts = await service.getGroupedCounts(table, envField, filter);
    Object.keys(counts).forEach(k => { out[k.toLowerCase()] = (out[k.toLowerCase()] || 0) + counts[k]; });
  }
  return out;
}

/**
 * Lädt die betroffenen Environments: Typ „Developer“, nicht managed, nicht gelöscht –
 * ergänzt um Ersteller, Zustand sowie Anzahl Apps und Flows.
 * managedField: in den Webpart-Eigenschaften gewählte Spalte (leer = automatisch erkennen).
 */
export async function loadAffectedEnvironments(service: DataverseService, managedField: string, maxRecords: number): Promise<IRecordResult> {
  const attrs = await environmentAttributes(service);
  const managedAttr = chooseManagedAttribute(attrs, managedField);
  const stateAttr = detectStateAttribute(attrs);
  // Typ-Spalte: „Environment Sku“ (aktuell) bzw. „Environment Type“ (ältere CoE-Versionen)
  const typeFields = ['admin_environmentsku', 'admin_environmenttype'].filter(f => attrs.length === 0 || attrs.some(a => a.logicalName === f));

  const extra = [managedAttr, stateAttr].filter(a => !!a).map(a => ({ field: a!.selectName, label: a!.displayName, type: 'text' as const }))
    // Ersteller des Environments (CoE-Lookup „Maker“) – wird ebenfalls angeschrieben
    .concat([{ field: '_admin_maker_value', label: 'Erstellt von', type: 'text' as const }]);
  const query: ITableConfig = {
    key: '_devenv', title: 'Environments', logicalName: ENV_TABLE, icon: 'Globe',
    filter: 'admin_environmentdeleted ne true',
    columns: buildDevEnvTable().columns.filter(c => c.field.indexOf('coe_') !== 0)
      .concat(typeFields.map(f => ({ field: f, label: 'Typ', type: 'text' as const })))
      .concat(extra)
  };
  const res = await service.getRecords(query, maxRecords);
  const warnings = res.warnings.slice();

  const developer = res.records.filter(isDeveloper);
  const affected = developer.filter(r => !managedAttr || managedValue(r, managedAttr.selectName) !== true);
  if (!managedAttr) {
    warnings.push('Im CoE-Inventar gibt es keine Spalte für den Managed-Status. Angezeigt werden deshalb alle Developer-Environments – ' +
      'ob eines bereits managed ist, bitte im Power Platform Admin Center prüfen oder die Spalte in den Webpart-Eigenschaften wählen.');
  }

  const ids = affected.map(r => String(r[res.meta.primaryId] || '')).filter(id => !!id);
  const [apps, flows, activity] = await Promise.all([
    countPerEnvironment(service, 'admin_app', '_admin_appenvironment_value', 'admin_appdeleted', ids).catch((e: Error) => {
      warnings.push(`Anzahl Apps konnte nicht ermittelt werden: ${e.message}`);
      return undefined;
    }),
    countPerEnvironment(service, 'admin_flow', '_admin_flowenvironment_value', 'admin_flowdeleted', ids).catch((e: Error) => {
      warnings.push(`Anzahl Flows konnte nicht ermittelt werden: ${e.message}`);
      return undefined;
    }),
    lastActivityPerEnvironment(service, ids).catch((e: Error) => {
      const all: { [envId: string]: boolean } = {};
      ids.forEach(id => { all[id.toLowerCase()] = true; });
      return { values: {} as { [envId: string]: string }, incomplete: all, errors: [e.message] };
    })
  ]);
  const creators: { [envId: string]: string } = {};
  affected.forEach(r => {
    const env = String(r[res.meta.primaryId] || '').toLowerCase();
    const maker = String(r._admin_maker_value || '').toLowerCase();
    if (env && maker) {
      creators[env] = maker;
    }
  });
  const [makerResult, admins] = await Promise.all([
    makersPerEnvironment(service, ids, creators, warnings).catch((e: Error) => {
      warnings.push(`Maker konnten nicht ermittelt werden: ${e.message}`);
      const unknown: { [envId: string]: boolean } = {};
      ids.forEach(id => { unknown[id.toLowerCase()] = true; });
      return { makers: {}, others: {}, othersUnknown: unknown } as IMakerResult;
    }),
    sysAdminsPerEnvironment(service, ids).catch((e: Error) => {
      warnings.push(`System-Administratoren konnten nicht ermittelt werden: ${e.message}`);
      return undefined;
    })
  ]);
  if (admins && ids.length && Object.keys(admins).length === 0) {
    warnings.push('Das CoE liefert keine System-Administratoren. Dafür in der CoE-Tabelle „Tenant Security Roles“ bei „System Administrator“ ' +
      '„Track SR“ aktivieren und den Flow „Admin | Sync Template v4 (Security Roles)“ einschalten.');
  }
  const inactiveBefore = new Date(Date.now() - INACTIVE_DAYS * 86400000).toISOString();

  const makers = makerResult.makers;
  const records = affected.map(r => {
    const id = String(r[res.meta.primaryId] || '').toLowerCase();
    const managed = managedAttr ? managedValue(r, managedAttr.selectName) : undefined;
    const state = stateLabel(r, stateAttr ? stateAttr.selectName : undefined);
    const appCount = apps ? apps[id] || 0 : null;
    const flowCount = flows ? flows[id] || 0 : null;
    const last = activity.values[id] || '';
    const recent = !!last && last >= inactiveBefore;
    const recommendation = state === 'Deaktiviert'
      ? RECOMMENDATION.disabled
      : recent
        ? RECOMMENDATION.inUse
        : appCount === 0 && flowCount === 0 && !makerResult.others[id]
          // „Leer“ nur, wenn auch Bots und Desktop Flows gezählt werden konnten
          ? (makerResult.othersUnknown[id] ? RECOMMENDATION.unknown : RECOMMENDATION.empty)
          // „Inaktiv“ nur, wenn alle Aktivitätsquellen lesbar waren
          : activity.incomplete[id] ? RECOMMENDATION.unknown : RECOMMENDATION.inactive;
    return {
      ...r,
      [DEV_FIELDS.recommendation]: recommendation,
      [DEV_FIELDS.lastActivity]: last || null,
      [DEV_FIELDS.makers]: (makers[id] || []).map(m => m.name).join(LIST_SEP),
      [DEV_FIELDS.makerCount]: (makers[id] || []).length,
      [DEV_FIELDS.makerEmails]: (makers[id] || []).map(m => m.email).filter(x => !!x).join(', '),
      [DEV_FIELDS.makersNoMail]: (makers[id] || []).filter(m => !m.email).map(m => m.name).join(LIST_SEP),
      [DEV_FIELDS.admins]: admins ? (admins[id] || []).join(LIST_SEP) : '',
      [DEV_FIELDS.state]: state,
      [DEV_FIELDS.managed]: managed === undefined ? 'Unbekannt' : managed ? 'Ja' : 'Nein',
      [DEV_FIELDS.apps]: appCount,
      [DEV_FIELDS.flows]: flowCount
    } as IRecord;
  });
  // nur die Environments zählen, bei denen die fehlende Aktivität wirklich die Empfehlung bestimmt
  const unknownCount = records.filter(r => r[DEV_FIELDS.recommendation] === RECOMMENDATION.unknown).length;
  if (unknownCount) {
    const why = activity.errors.length ? ` (${activity.errors.join('; ')})` : '';
    warnings.push(`Aktivität bzw. Inhalt konnte für ${unknownCount} Environment(s) nicht vollständig ermittelt werden – dort steht die Empfehlung „${RECOMMENDATION.unknown}“ statt „Inaktiv“ bzw. „Leer“${why}. Bitte „Aktualisieren“.`);
  }
  return { records, meta: res.meta, warnings, truncated: res.truncated };
}
