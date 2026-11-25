import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import { ITableConfig, IRecord } from '../models/ITableConfig';
import { combineFilters, displayValue, templateFields } from '../utils/format';

const API_PATH = '/api/data/v9.2/';
const MAX_PAGE = 5000;

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

export interface IEnvironmentInfo {
  /** Primärschlüssel des CoE-Datensatzes (admin_environmentid). */
  id: string;
  name: string;
  type: string;
  /** Dataverse-URL der Umgebung (falls vorhanden). */
  dataverseUrl?: string;
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
  public async getRecords(table: ITableConfig, maxRecords: number): Promise<IRecordResult> {
    const meta = await this.getEntityMeta(table.logicalName);
    let fields = collectFields(table, meta);
    let filter = table.filter && table.filter.trim() ? table.filter.trim() : undefined;
    const warnings: string[] = [];

    for (let attempt = 0; attempt < 20; attempt++) {
      try {
        const sortField = table.defaultSort && fields.indexOf(table.defaultSort.field) >= 0 ? table.defaultSort : undefined;
        const page = await this._fetchAll(meta.entitySetName, fields, filter, sortField, maxRecords);
        return { records: page.records, meta, warnings, truncated: page.truncated };
      } catch (e) {
        const missing = missingProperty(e);
        if (missing && fields.indexOf(missing) >= 0 && missing !== meta.primaryId) {
          fields = fields.filter(f => f !== missing);
          warnings.push(`Spalte „${missing}“ gibt es in dieser Tabelle nicht und wurde übersprungen.`);
          continue;
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
          { field: 'admin_environmenttype', label: 'Typ', type: 'choice' },
          { field: 'admin_environmentcdsinstanceurl', label: 'Dataverse-URL', type: 'url' },
          { field: 'admin_environmentdeleted', label: 'Gelöscht', type: 'boolean' }
        ]
      };
      const p = this.getRecords(table, 5000).then(res => res.records
        .filter(r => r.admin_environmentdeleted !== true)
        .map(r => ({
          id: String(r[res.meta.primaryId] || ''),
          name: displayValue(r, 'admin_displayname') || String(r[res.meta.primaryId] || ''),
          type: displayValue(r, 'admin_environmenttype'),
          dataverseUrl: normalizeDataverseUrl(displayValue(r, 'admin_environmentcdsinstanceurl'))
        })));
      this._environments = p;
      p.catch(() => { this._environments = undefined; });
    }
    return this._environments;
  }

  /** Anzahl der Datensätze (Dataverse zählt bis maximal 5.000). */
  public async getCount(logicalName: string, filter?: string): Promise<number> {
    const meta = await this.getEntityMeta(logicalName);
    let q = `${meta.entitySetName}?$select=${meta.primaryId}&$count=true`;
    if (filter) {
      q += `&$filter=${encodeURIComponent(filter)}`;
    }
    const page = await this._get<IODataPage>(q, 1);
    return typeof page['@odata.count'] === 'number' ? page['@odata.count'] : page.value.length;
  }

  public getCountWithBase(table: ITableConfig, extraFilter?: string): Promise<number> {
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
    maxRecords: number
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
    const prefer = ['odata.include-annotations="*"'];
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
