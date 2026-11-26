import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import { IRecord, IRelatedConfig, ITableConfig } from '../models/ITableConfig';
import { combineFilters, displayValue, templateFields } from '../utils/format';

const API_PATH = '/api/data/v9.2/';
const MAX_PAGE = 5000;
/** $count=true zählt in Dataverse höchstens bis hierher. */
const ODATA_COUNT_LIMIT = 5000;
/** Aggregat-Abfragen ($apply) zählen höchstens bis hierher (AggregateQueryRecordLimit). */
const AGGREGATE_LIMIT = 50000;

type CountMode = 'count' | 'distinct' | 'odata';
const FORMATTED_VALUE = '@OData.Community.Display.V1.FormattedValue';

export interface ICountResult {
  value: number;
  /** true = es sind mindestens so viele (Zählgrenze von Dataverse erreicht). */
  atLeast: boolean;
}

function isAggregateLimit(e: unknown): boolean {
  const msg = e instanceof Error ? e.message : String(e);
  return /AggregateQueryRecordLimit|maximum record limit|0x8004E023/i.test(msg);
}

export interface IEntityMeta {
  logicalName: string;
  entitySetName: string;
  primaryId: string;
  primaryName: string;
  displayName: string;
}

export interface IAttributeMeta {
  logicalName: string;
  /** Name für $select (Lookups: _feld_value). */
  selectName: string;
  type: string;
  displayName: string;
  isPrimaryName: boolean;
}

export interface IRecordResult {
  records: IRecord[];
  meta: IEntityMeta;
  warnings: string[];
  truncated: boolean;
}

interface IODataPage {
  value: IRecord[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

interface ILabel {
  UserLocalizedLabel?: { Label?: string } | null;
}

interface IEntityDefinition {
  LogicalName: string;
  EntitySetName: string;
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string;
  DisplayName?: ILabel;
}

interface IAttributeDefinition {
  LogicalName: string;
  AttributeType: string;
  DisplayName?: ILabel;
  IsValidForRead?: boolean;
  AttributeOf?: string | null;
}

export interface ITableInfo {
  logicalName: string;
  displayName: string;
  isCustom: boolean;
}

export interface IEnvironmentInfo {
  /** Primärschlüssel des CoE-Datensatzes (admin_environmentid). */
  id: string;
  name: string;
  type: string;
  /** Dataverse-URL der Umgebung (falls vorhanden). */
  dataverseUrl?: string;
}

/** Typ eines CoE-Environments (Production, Sandbox, Developer, Default …) – „Environment Sku“ bzw. älter „Environment Type“. */
export function environmentType(r: IRecord): string {
  return displayValue(r, 'admin_environmentsku') || displayValue(r, 'admin_environmenttype');
}

/** Ergänzt https:// und entfernt den abschließenden Schrägstrich. */
export function normalizeDataverseUrl(url: string | undefined): string | undefined {
  const v = (url || '').trim();
  if (!v) {
    return undefined;
  }
  return (/^https?:\/\//i.test(v) ? v : `https://${v}`).replace(/^http:/i, 'https:').replace(/\/+$/, '');
}

/** Sucht das Default Environment: konfigurierte ID, sonst Typ „Default“, sonst Name „(default)“. */
export function findDefaultEnvironment(envs: IEnvironmentInfo[], configuredId?: string): IEnvironmentInfo | undefined {
  const wanted = (configuredId || '').trim().toLowerCase().replace(/^default-/, '');
  if (wanted) {
    return envs.filter(e => e.id.toLowerCase().replace(/^default-/, '') === wanted)[0];
  }
  return envs.filter(e => e.type.toLowerCase() === 'default')[0] ||
    envs.filter(e => /\(default\)|^default\b/i.test(e.name))[0];
}

/** Umgebungs-ID für Links ins Maker Portal / Admin Center (Default-Umgebung: „Default-<Tenant-ID>“). */
export function environmentLinkId(env: IEnvironmentInfo): string {
  if (/^default-/i.test(env.id)) {
    return env.id;
  }
  return env.type.toLowerCase() === 'default' ? `Default-${env.id}` : env.id;
}

export class DataverseError extends Error {
  public status: number;
  constructor(message: string, status: number) {
    super(message);
    this.status = status;
    Object.setPrototypeOf(this, DataverseError.prototype);
  }
}

const LOOKUP_TYPES = ['Lookup', 'Owner', 'Customer'];
const SKIP_TYPES = ['Virtual', 'EntityName', 'CalendarRules', 'PartyList', 'ManagedProperty'];

export class DataverseService {
  public readonly baseUrl: string;
  private _factory: AadHttpClientFactory;
  private _client: Promise<AadHttpClient> | undefined;
  private _meta: { [logicalName: string]: Promise<IEntityMeta> } = {};
  private _attributes: { [logicalName: string]: Promise<IAttributeMeta[]> } = {};
  private _environments: Promise<IEnvironmentInfo[]> | undefined;
  private _countMode: CountMode | undefined;
  private _tables: Promise<ITableInfo[]> | undefined;

  constructor(factory: AadHttpClientFactory, environmentUrl: string) {
    this._factory = factory;
    this.baseUrl = environmentUrl.trim().replace(/\/+$/, '');
  }

  /** Metadaten einer Tabelle (Entity-Set, Primärschlüssel, Primärname). */
  public getEntityMeta(logicalName: string): Promise<IEntityMeta> {
    const key = logicalName.trim().toLowerCase();
    if (!this._meta[key]) {
      const p = this._get<IEntityDefinition>(
        `EntityDefinitions(LogicalName='${key}')?$select=LogicalName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,DisplayName`
      ).then(d => ({
        logicalName: d.LogicalName,
        entitySetName: d.EntitySetName,
        primaryId: d.PrimaryIdAttribute,
        primaryName: d.PrimaryNameAttribute,
        displayName: labelOf(d.DisplayName) || d.LogicalName
      }));
      this._meta[key] = p;
      p.catch(() => { delete this._meta[key]; });
    }
    return this._meta[key];
  }

  /** Alle Tabellen der Umgebung, die sich anzeigen lassen (für die Auswahl im Tabellen-Editor). */
  public getTables(): Promise<ITableInfo[]> {
    if (!this._tables) {
      const p = this._get<{ value: (IEntityDefinition & { IsCustomEntity?: boolean; IsValidForAdvancedFind?: boolean; IsIntersect?: boolean; IsPrivate?: boolean })[] }>(
        'EntityDefinitions?$select=LogicalName,EntitySetName,DisplayName,IsCustomEntity,IsValidForAdvancedFind,IsIntersect,IsPrivate'
      ).then(res => res.value
        .filter(d => !!d.EntitySetName && d.IsValidForAdvancedFind !== false && !d.IsIntersect && !d.IsPrivate)
        .map(d => ({
          logicalName: d.LogicalName,
          displayName: labelOf(d.DisplayName) || d.LogicalName,
          isCustom: !!d.IsCustomEntity
        }))
        .sort((x, y) => x.displayName.localeCompare(y.displayName, 'de')));
      this._tables = p;
      p.catch(() => { this._tables = undefined; });
    }
    return this._tables;
  }

  /** Lesbare Spalten einer Tabelle (für den Tabellen-Editor und die Detailansicht). */
  public getAttributes(logicalName: string): Promise<IAttributeMeta[]> {
    const key = logicalName.trim().toLowerCase();
    if (!this._attributes[key]) {
      const p = Promise.all([
        this.getEntityMeta(key),
        this._get<{ value: IAttributeDefinition[] }>(
          `EntityDefinitions(LogicalName='${key}')/Attributes?$select=LogicalName,AttributeType,DisplayName,IsValidForRead,AttributeOf`
        )
      ]).then(([meta, res]) => res.value
        .filter(a => a.IsValidForRead !== false && !a.AttributeOf && SKIP_TYPES.indexOf(a.AttributeType) < 0)
        .map(a => ({
          logicalName: a.LogicalName,
          selectName: LOOKUP_TYPES.indexOf(a.AttributeType) >= 0 ? `_${a.LogicalName}_value` : a.LogicalName,
          type: a.AttributeType,
          displayName: labelOf(a.DisplayName) || a.LogicalName,
          isPrimaryName: a.LogicalName === meta.primaryName
        }))
        .sort((x, y) => x.displayName.localeCompare(y.displayName, 'de')));
      this._attributes[key] = p;
      p.catch(() => { delete this._attributes[key]; });
    }
    return this._attributes[key];
  }

  /**
   * Lädt Datensätze einer konfigurierten Tabelle.
   * Nicht vorhandene Spalten werden automatisch entfernt und als Warnung gemeldet.
   */
  public async getRecords(
    table: ITableConfig,
    maxRecords: number,
    onProgress?: (records: IRecord[], meta: IEntityMeta) => void
  ): Promise<IRecordResult> {
    const meta = await this.getEntityMeta(table.logicalName);
    let fields = collectFields(table, meta);
    let filter = table.filter && table.filter.trim() ? table.filter.trim() : undefined;
    const warnings: string[] = [];

    for (let attempt = 0; attempt < 20; attempt++) {
      try {
        const sortField = table.defaultSort && fields.indexOf(table.defaultSort.field) >= 0 ? table.defaultSort : undefined;
        const page = await this._fetchAll(meta.entitySetName, fields, filter, sortField, maxRecords,
          onProgress ? (records: IRecord[]) => onProgress(records, meta) : undefined);
        return { records: page.records, meta, warnings, truncated: page.truncated };
      } catch (e) {
        const missing = missingProperty(e);
        if (missing && fields.indexOf(missing) >= 0 && missing !== meta.primaryId) {
          fields = fields.filter(f => f !== missing);
          warnings.push(`Spalte „${missing}“ gibt es in dieser Tabelle nicht und wurde übersprungen.`);
          continue;
        }
        if (missing && filter && filter.indexOf(missing) >= 0 && table.strictFilter) {
          // Filter begrenzt den Bereich (z. B. auf das Default Environment) – nie stillschweigend weglassen.
          throw new DataverseError(`Der Filter verwendet die unbekannte Spalte „${missing}“: ${filter}`, 400);
        }
        if (missing && filter && filter.indexOf(missing) >= 0) {
          warnings.push(`Der Filter verwendet die unbekannte Spalte „${missing}“ und wurde ignoriert.`);
          filter = undefined;
          continue;
        }
        throw e;
      }
    }
    throw new DataverseError('Die Daten konnten nicht geladen werden.', 400);
  }

  /** Alle Umgebungen aus dem CoE-Inventar (ohne gelöschte). */
  public getEnvironments(): Promise<IEnvironmentInfo[]> {
    if (!this._environments) {
      const table: ITableConfig = {
        key: '_environments',
        title: 'Environments',
        logicalName: 'admin_environment',
        icon: 'Globe',
        defaultSort: { field: 'admin_displayname', desc: false },
        columns: [
          { field: 'admin_displayname', label: 'Name', type: 'title' },
          // Typ: in aktuellen CoE-Versionen „Environment Sku“, in älteren „Environment Type“
          { field: 'admin_environmentsku', label: 'Typ', type: 'choice' },
          { field: 'admin_environmenttype', label: 'Typ (alt)', type: 'choice' },
          { field: 'admin_environmentcdsinstanceurl', label: 'Dataverse-URL', type: 'url' },
          { field: 'admin_environmentdeleted', label: 'Gelöscht', type: 'boolean' }
        ]
      };
      const p = this.getRecords(table, 5000).then(res => res.records
        .filter(r => r.admin_environmentdeleted !== true)
        .map(r => ({
          id: String(r[res.meta.primaryId] || ''),
          name: displayValue(r, 'admin_displayname') || String(r[res.meta.primaryId] || ''),
          type: environmentType(r),
          dataverseUrl: normalizeDataverseUrl(displayValue(r, 'admin_environmentcdsinstanceurl'))
        })));
      this._environments = p;
      p.catch(() => { this._environments = undefined; });
    }
    return this._environments;
  }

  /**
   * Exakte Anzahl der Datensätze.
   * 1. OData-Aggregation ($apply=…/aggregate) – zählt exakt bis 50.000 (Dataverse-Grenze für Aggregate).
   * 2. Rückfall auf $count=true – Dataverse zählt damit nur bis 5.000.
   * Welche Variante die Umgebung unterstützt, wird beim ersten Aufruf ermittelt und gemerkt.
   */
  public async getCount(logicalName: string, filter?: string): Promise<ICountResult> {
    const meta = await this.getEntityMeta(logicalName);
    const prefix = filter ? `filter(${filter})/` : '';
    const variants: { mode: CountMode; apply: string }[] = [
      { mode: 'count', apply: `${prefix}aggregate($count as total)` },
      { mode: 'distinct', apply: `${prefix}aggregate(${meta.primaryId} with countdistinct as total)` }
    ];
    let firstError: unknown;
    for (const v of variants) {
      if (this._countMode && this._countMode !== v.mode) {
        continue;
      }
      try {
        const res = await this._get<{ value: { total?: number | string }[] }>(
          `${meta.entitySetName}?$apply=${encodeURIComponent(v.apply)}`
        );
        const total = res.value && res.value.length ? Number(res.value[0].total) : 0;
        if (!isNaN(total)) {
          this._countMode = v.mode;
          return { value: total, atLeast: false };
        }
      } catch (e) {
        if (isAggregateLimit(e)) {
          return { value: AGGREGATE_LIMIT, atLeast: true };
        }
        if (this._countMode) {
          throw e; // Variante funktioniert grundsätzlich – Fehler liegt an der Abfrage (z. B. Filter)
        }
        firstError = firstError || e;
      }
    }
    let q = `${meta.entitySetName}?$select=${meta.primaryId}&$count=true`;
    if (filter) {
      q += `&$filter=${encodeURIComponent(filter)}`;
    }
    try {
      const page = await this._get<IODataPage>(q, 1);
      const n = typeof page['@odata.count'] === 'number' ? page['@odata.count'] : page.value.length;
      if (!this._countMode && firstError !== undefined) {
        this._countMode = 'odata'; // Aggregation wird nicht unterstützt
      }
      return { value: n, atLeast: n >= ODATA_COUNT_LIMIT };
    } catch (e) {
      throw firstError !== undefined && !missingProperty(e) ? firstError : e;
    }
  }

  /**
   * Anzahl (oder Summe) je Wert einer Spalte, z. B. Flows je Environment.
   * Nutzt OData-Aggregation; wird sie nicht unterstützt oder ist die Datenmenge zu groß,
   * werden die Datensätze geladen und im Browser gezählt.
   */
  public async getGroupedCounts(logicalName: string, groupField: string, filter?: string, sumField?: string): Promise<{ [value: string]: number }> {
    const meta = await this.getEntityMeta(logicalName);
    const aggregate = sumField ? `${sumField} with sum as n` : '$count as n';
    const apply = `${filter ? `filter(${filter})/` : ''}groupby((${groupField}),aggregate(${aggregate}))`;
    // Lookups nach ID gruppieren (der formatierte Wert wäre der Name), sonst nach Anzeigewert
    const useRaw = /^_.+_value$/.test(groupField);
    const keyOf = (r: IRecord): string => String(!useRaw && r[groupField + FORMATTED_VALUE] !== undefined
      ? r[groupField + FORMATTED_VALUE]
      : (r[groupField] === null || r[groupField] === undefined ? '' : r[groupField]));
    const out: { [value: string]: number } = {};
    try {
      const res = await this._get<{ value: IRecord[] }>(`${meta.entitySetName}?$apply=${encodeURIComponent(apply)}`);
      res.value.forEach(r => { out[keyOf(r)] = (out[keyOf(r)] || 0) + Number(r.n || 0); });
      return out;
    } catch (e) {
      if (missingProperty(e)) {
        throw e;
      }
    }
    const table: ITableConfig = {
      key: '_grouped', title: logicalName, logicalName, icon: 'Table', filter,
      columns: [{ field: groupField, label: groupField, type: 'text' }].concat(sumField ? [{ field: sumField, label: sumField, type: 'number' }] : []) as ITableConfig['columns']
    };
    const rec = await this.getRecords(table, AGGREGATE_LIMIT);
    rec.records.forEach(r => { out[keyOf(r)] = (out[keyOf(r)] || 0) + (sumField ? Number(r[sumField] || 0) : 1); });
    return out;
  }

  /**
   * Verknüpfte Datensätze eines Datensatzes (Reiter in der Detailansicht).
   * 1:N über eine Lookup-Spalte der verknüpften Tabelle, N:N über die Navigationseigenschaft der Beziehung.
   */
  public async getRelated(parentLogicalName: string, parentId: string, rel: IRelatedConfig, maxRecords: number): Promise<IRecordResult> {
    if (rel.lookupField) {
      return this.getRecords({
        key: `_rel_${rel.key}`, title: rel.title, logicalName: rel.logicalName, icon: 'Table', strictFilter: true,
        filter: combineFilters(rel.filter, `${rel.lookupField} eq ${parentId}`),
        defaultSort: rel.defaultSort,
        columns: rel.columns
      }, maxRecords);
    }
    if (!rel.navigationProperty) {
      throw new DataverseError(`Für „${rel.title}“ ist keine Verknüpfung konfiguriert.`, 400);
    }
    const [parent, meta] = await Promise.all([this.getEntityMeta(parentLogicalName), this.getEntityMeta(rel.logicalName)]);
    let fields = collectFields({ key: '_rel', title: rel.title, logicalName: rel.logicalName, icon: 'Table', columns: rel.columns }, meta);
    const warnings: string[] = [];
    for (let attempt = 0; attempt < 20; attempt++) {
      try {
        const base = `${parent.entitySetName}(${parentId})/${rel.navigationProperty}`;
        const page = await this._fetchAll(base, fields, rel.filter, undefined, maxRecords);
        return { records: page.records, meta, warnings, truncated: page.truncated };
      } catch (e) {
        const missing = missingProperty(e);
        if (missing && fields.indexOf(missing) >= 0 && missing !== meta.primaryId) {
          fields = fields.filter(f => f !== missing);
          warnings.push(`Spalte „${missing}“ gibt es in dieser Tabelle nicht und wurde übersprungen.`);
          continue;
        }
        throw e;
      }
    }
    throw new DataverseError('Die verknüpften Datensätze konnten nicht geladen werden.', 400);
  }

  public getCountWithBase(table: ITableConfig, extraFilter?: string): Promise<ICountResult> {
    return this.getCount(table.logicalName, combineFilters(table.filter, extraFilter));
  }

  /** Vollständiger Datensatz mit allen Spalten. */
  public async getRecord(logicalName: string, id: string): Promise<IRecord> {
    const meta = await this.getEntityMeta(logicalName);
    return this._get<IRecord>(`${meta.entitySetName}(${id})`);
  }

  public recordUrl(logicalName: string, id: string): string {
    return `${this.baseUrl}/main.aspx?pagetype=entityrecord&etn=${encodeURIComponent(logicalName)}&id=${encodeURIComponent(id)}`;
  }

  private async _fetchAll(
    entitySet: string,
    fields: string[],
    filter: string | undefined,
    sort: { field: string; desc: boolean } | undefined,
    maxRecords: number,
    onProgress?: (records: IRecord[]) => void
  ): Promise<{ records: IRecord[]; truncated: boolean }> {
    let url: string | undefined = `${entitySet}?$select=${fields.join(',')}`;
    if (filter) {
      url += `&$filter=${encodeURIComponent(filter)}`;
    }
    if (sort) {
      url += `&$orderby=${sort.field}${sort.desc ? ' desc' : ' asc'}`;
    }
    const pageSize = Math.max(1, Math.min(MAX_PAGE, maxRecords));
    let records: IRecord[] = [];
    while (url && records.length < maxRecords) {
      const page: IODataPage = await this._get<IODataPage>(url, pageSize);
      records = records.concat(page.value);
      url = page['@odata.nextLink'];
      if (onProgress && url && records.length < maxRecords) {
        onProgress(records); // Zwischenstand anzeigen, während weitere Seiten geladen werden
      }
    }
    const truncated = !!url || records.length > maxRecords;
    return { records: records.slice(0, maxRecords), truncated };
  }

  private _getClient(): Promise<AadHttpClient> {
    if (!this._client) {
      const p = this._factory.getClient(this.baseUrl);
      this._client = p;
      p.catch(() => { this._client = undefined; });
    }
    return this._client;
  }

  private async _get<T>(pathOrUrl: string, maxPageSize?: number): Promise<T> {
    const client = await this._getClient();
    const url = /^https:\/\//i.test(pathOrUrl) ? pathOrUrl : this.baseUrl + API_PATH + pathOrUrl;
    // Nur formatierte Werte anfordern – bei großen Tabellen deutlich kleinere Antworten als mit "*".
    const prefer = ['odata.include-annotations="OData.Community.Display.V1.FormattedValue"'];
    if (maxPageSize) {
      prefer.push(`odata.maxpagesize=${maxPageSize}`);
    }
    const res: HttpClientResponse = await client.get(url, AadHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'Prefer': prefer.join(',')
      }
    });
    if (!res.ok) {
      throw await toError(res);
    }
    return (await res.json()) as T;
  }
}

function labelOf(label?: ILabel): string {
  return (label && label.UserLocalizedLabel && label.UserLocalizedLabel.Label) || '';
}

function collectFields(table: ITableConfig, meta: { primaryId: string; primaryName: string }): string[] {
  const set: { [f: string]: boolean } = {};
  const out: string[] = [];
  const add = (f: string | undefined): void => {
    const v = (f || '').trim();
    if (v && !set[v]) {
      set[v] = true;
      out.push(v);
    }
  };
  add(meta.primaryId);
  add(meta.primaryName);
  table.columns.forEach(c => add(c.field));
  (table.links || []).forEach(l => templateFields(l.url).forEach(add));
  if (table.defaultSort) {
    add(table.defaultSort.field);
  }
  return out;
}

function missingProperty(e: unknown): string | undefined {
  const msg = e instanceof Error ? e.message : String(e);
  const m = /property named '([^']+)'/i.exec(msg) || /Could not find a property named ([A-Za-z0-9_]+)/i.exec(msg);
  return m ? m[1] : undefined;
}

async function toError(res: HttpClientResponse): Promise<DataverseError> {
  let message = `HTTP ${res.status}`;
  try {
    const text = await res.text();
    try {
      const json = JSON.parse(text) as { error?: { message?: string } };
      message = (json.error && json.error.message) || text || message;
    } catch {
      message = text || message;
    }
  } catch {
    // ignore
  }
  return new DataverseError(message, res.status);
}
