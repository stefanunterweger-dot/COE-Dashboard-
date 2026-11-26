import { IRecord, ITableConfig } from './ITableConfig';
import { DataverseService, environmentType, IAttributeMeta, IRecordResult } from '../services/DataverseService';
import { displayValue, rawValue } from '../utils/format';

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
  creator: 'coe_creator',
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
  disabled: 'Deaktiviert'
};
const INACTIVE_DAYS = 90;

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
      { field: DEV_FIELDS.creator, label: 'Ersteller', type: 'person' },
      {
        field: DEV_FIELDS.recommendation, label: 'Empfehlung', type: 'status',
        statusMap: { [RECOMMENDATION.inUse]: 'danger', [RECOMMENDATION.inactive]: 'warning', [RECOMMENDATION.empty]: 'neutral', [RECOMMENDATION.disabled]: 'neutral' }
      },
      { field: DEV_FIELDS.state, label: 'Zustand', type: 'status', statusMap: { 'Aktiv': 'success', 'Deaktiviert': 'danger', 'Unbekannt': 'neutral' } },
      { field: DEV_FIELDS.apps, label: 'Apps', type: 'number' },
      { field: DEV_FIELDS.flows, label: 'Flows', type: 'number' },
      { field: DEV_FIELDS.lastActivity, label: 'Letzte Aktivität', type: 'date' },
      { field: 'admin_region', label: 'Region', type: 'choice', hidden: true },
      { field: 'admin_environmentcreatedon', label: 'Erstellt', type: 'date' },
      { field: DEV_FIELDS.managed, label: 'Managed', type: 'status', hidden: true, statusMap: { 'Nein': 'warning', 'Ja': 'success', 'Unbekannt': 'neutral' } },
      { field: 'admin_environmentcdsinstanceurl', label: 'Dataverse-URL', type: 'url', hidden: true },
      { field: 'admin_environmentid', label: 'Umgebungs-ID', type: 'text', hidden: true }
    ],
    links: [
      { label: 'Im Power Platform Admin Center öffnen', url: 'https://admin.powerplatform.microsoft.com/environments/{admin_environmentid}/hub' }
    ]
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

export function detectCreatorAttribute(attrs: IAttributeMeta[]): IAttributeMeta | undefined {
  // CoE: „Environment Maker“ (Lookup auf Maker) bevorzugen
  return findAttribute(attrs, [/maker|creator|ersteller|erstellt von|created ?by/i], ['Lookup', 'Owner', 'Customer'])
    || findAttribute(attrs, [/maker|creator|ersteller|erstellt von|created ?by/i], ['String']);
}

/** Spalten, die für den Managed-Status in Frage kommen (Auswahl in den Webpart-Eigenschaften). */
export function managedCandidates(attrs: IAttributeMeta[]): IAttributeMeta[] {
  return attrs.filter(a => ['Boolean', 'Picklist', 'String'].indexOf(a.type) >= 0 && SYSTEM_FIELDS.indexOf(a.logicalName) < 0);
}

function isDeveloper(r: IRecord): boolean {
  return /developer|entwickler/i.test(environmentType(r));
}

/** true / false / undefined (unbekannt). */
function managedValue(r: IRecord, field: string): boolean | undefined {
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

/** Jüngstes Datum je Environment aus mehreren Spalten (App-Start, App-/Flow-Änderung). */
async function lastActivityPerEnvironment(service: DataverseService, envIds: string[]): Promise<{ [envId: string]: string }> {
  const out: { [envId: string]: string } = {};
  const sources: [string, string, string, string][] = [
    ['admin_app', '_admin_appenvironment_value', 'admin_appdeleted', 'admin_applastlaunchedon'],
    ['admin_app', '_admin_appenvironment_value', 'admin_appdeleted', 'admin_appmodifiedon'],
    ['admin_flow', '_admin_flowenvironment_value', 'admin_flowdeleted', 'admin_flowmodifiedon']
  ];
  for (let i = 0; i < envIds.length; i += BATCH) {
    const ids = envIds.slice(i, i + BATCH);
    const results = await Promise.all(sources.map(([table, envField, deletedField, dateField]) => service
      .getGroupedMax(table, envField, `(${deletedField} ne true) and (${ids.map(id => `${envField} eq ${id}`).join(' or ')})`, dateField)
      .catch(() => ({} as { [k: string]: string }))));
    results.forEach(r => Object.keys(r).forEach(k => {
      if (!out[k] || r[k] > out[k]) {
        out[k] = r[k];
      }
    }));
  }
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
  const attrs = await service.getAttributes(ENV_TABLE).catch(() => [] as IAttributeMeta[]);
  const chosen = managedField ? attrs.filter(a => a.selectName === managedField || a.logicalName === managedField)[0] : undefined;
  const managedAttr = chosen || detectManagedAttribute(attrs);
  const stateAttr = detectStateAttribute(attrs);
  const creatorAttr = detectCreatorAttribute(attrs);
  // Typ-Spalte: „Environment Sku“ (aktuell) bzw. „Environment Type“ (ältere CoE-Versionen)
  const typeFields = ['admin_environmentsku', 'admin_environmenttype'].filter(f => attrs.length === 0 || attrs.some(a => a.logicalName === f));

  const extra = [managedAttr, stateAttr, creatorAttr].filter(a => !!a).map(a => ({ field: a!.selectName, label: a!.displayName, type: 'text' as const }));
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
    lastActivityPerEnvironment(service, ids).catch(() => ({} as { [envId: string]: string }))
  ]);
  const inactiveBefore = new Date(Date.now() - INACTIVE_DAYS * 86400000).toISOString();

  const records = affected.map(r => {
    const id = String(r[res.meta.primaryId] || '').toLowerCase();
    const managed = managedAttr ? managedValue(r, managedAttr.selectName) : undefined;
    const state = stateLabel(r, stateAttr ? stateAttr.selectName : undefined);
    const appCount = apps ? apps[id] || 0 : null;
    const flowCount = flows ? flows[id] || 0 : null;
    const last = activity[id] || '';
    const recommendation = state === 'Deaktiviert'
      ? RECOMMENDATION.disabled
      : appCount === 0 && flowCount === 0
        ? RECOMMENDATION.empty
        : !last || last < inactiveBefore ? RECOMMENDATION.inactive : RECOMMENDATION.inUse;
    return {
      ...r,
      [DEV_FIELDS.recommendation]: recommendation,
      [DEV_FIELDS.lastActivity]: last || null,
      [DEV_FIELDS.creator]: creatorAttr ? displayValue(r, creatorAttr.selectName) : '',
      [DEV_FIELDS.state]: state,
      [DEV_FIELDS.managed]: managed === undefined ? 'Unbekannt' : managed ? 'Ja' : 'Nein',
      [DEV_FIELDS.apps]: appCount,
      [DEV_FIELDS.flows]: flowCount
    } as IRecord;
  });
  return { records, meta: res.meta, warnings, truncated: res.truncated };
}
