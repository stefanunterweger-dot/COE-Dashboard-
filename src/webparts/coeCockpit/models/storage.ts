import { DataverseService } from '../services/DataverseService';
import { PowerPlatformApiService, PpApiError, PPAPI_VERSIONS } from '../services/PowerPlatformApiService';

/**
 * Dataverse-Speicher: Kontingent des Tenants live aus der Power Platform API („licensing/tenantCapacity“),
 * Verbrauch je Environment aus dem CoE („Environment Capacity“). Alle Werte in MB.
 */
export const STORAGE_KEY = 'dataverse-storage';

export type StorageType = 'Database' | 'File' | 'Log';
export const STORAGE_TYPES: StorageType[] = ['Database', 'File', 'Log'];
export const STORAGE_LABEL: { [t in StorageType]: string } = { Database: 'Datenbank', File: 'Datei', Log: 'Log' };

export interface ITenantStorage {
  type: StorageType;
  /** Kontingent (MB) */
  total?: number;
  /** genutzt (MB) */
  used?: number;
  /** über ein anderes Kontingent gedeckt (MB) */
  overflow?: number;
}

export interface IEnvStorage {
  id: string;
  name: string;
  type: string;
  db: number;
  file: number;
  log: number;
}

export interface IStorageData {
  tenant: ITenantStorage[];
  tenantStatus: string;
  tenantProblem?: string;
  tenantProblemKind?: string;
  envs: IEnvStorage[];
  envProblem?: string;
  loadedAt: Date;
}

interface ITenantCapacity {
  capacitySummary?: { status?: string; statusMessage?: string; statusMessageCode?: string };
  tenantCapacities?: {
    capacityType?: string; capacityUnits?: string; totalCapacity?: number; maxCapacity?: number;
    consumption?: { actual?: number; rated?: number };
    overflowCapacity?: { capacityType?: string; value?: number }[];
  }[];
}

const STATUS_TEXT: { [code: string]: string } = {
  AllCapacityAvailable: 'Genug Kapazität vorhanden.',
  DBCapacityOver: 'Datenbank-Kapazität überschritten.',
  LogCapacityOver: 'Log-Kapazität überschritten.',
  FileCapacityOver: 'Datei-Kapazität überschritten.',
  LogOverDBCover: 'Log überschritten, durch freie Datenbank-Kapazität gedeckt.',
  FileOverDBCover: 'Datei überschritten, durch freie Datenbank-Kapazität gedeckt.',
  FileOverLogCover: 'Datei überschritten, durch freie Log-Kapazität gedeckt.',
  FileOverDBAndLogCover: 'Datei überschritten, durch Datenbank und Log gedeckt.',
  DBAndLogOver: 'Datenbank und Log überschritten.',
  DBAndFileOver: 'Datenbank und Datei überschritten.',
  LogAndFileOver: 'Log und Datei überschritten.',
  AllCapacityOver: 'Alle Kapazitäten überschritten.',
  DBCapacityMoreThanEightyFive: 'Datenbank zu über 85 % belegt.',
  FileCapacityMoreThanEightyFive: 'Datei zu über 85 % belegt.',
  LogCapacityMoreThanEightyFive: 'Log zu über 85 % belegt.',
  AllCapacityMoreThanEightyFive: 'Alle Kapazitäten zu über 85 % belegt.'
};

export async function loadStorage(service: DataverseService, ppApi: PowerPlatformApiService | undefined): Promise<IStorageData> {
  const out: IStorageData = { tenant: [], tenantStatus: '', envs: [], loadedAt: new Date() };

  // Tenant-Kontingent (live)
  if (ppApi) {
    try {
      const cap = await ppApi.get<ITenantCapacity>('/licensing/tenantCapacity', PPAPI_VERSIONS.licensing);
      STORAGE_TYPES.forEach(t => {
        const c = (cap.tenantCapacities || []).filter(x => String(x.capacityType || '').toLowerCase() === t.toLowerCase())[0];
        out.tenant.push({
          type: t,
          total: c && typeof c.totalCapacity === 'number' ? c.totalCapacity : undefined,
          used: c && c.consumption && typeof c.consumption.actual === 'number' ? c.consumption.actual : undefined,
          overflow: c && c.overflowCapacity ? c.overflowCapacity.reduce((s, o) => s + (o.value || 0), 0) : undefined
        });
      });
      const sum = cap.capacitySummary || {};
      out.tenantStatus = STATUS_TEXT[String(sum.statusMessageCode || '')] || String(sum.statusMessage || sum.status || '');
    } catch (e) {
      out.tenantProblem = e instanceof PpApiError ? e.message : e instanceof Error ? e.message : String(e);
      out.tenantProblemKind = e instanceof PpApiError ? e.kind : 'failed';
    }
  } else {
    out.tenantProblem = 'Die Power Platform API ist ausgeschaltet (Webpart-Einstellungen → Ansichten → Copilot-Credits).';
  }

  // Verbrauch je Environment (CoE)
  try {
    const [capRes, envRes] = await Promise.all([
      service.getRecords({
        key: '_storage', title: 'Kapazität', logicalName: 'admin_environmentcapacity', icon: 'Database', filter: 'statecode eq 0',
        columns: [
          { field: '_admin_environment_value', label: 'Environment', type: 'text' },
          { field: 'admin_capacitytype', label: 'Typ', type: 'text' },
          { field: 'admin_actualconsumption', label: 'Verbrauch', type: 'number' },
          { field: 'admin_capacityunit', label: 'Einheit', type: 'text' }
        ]
      }, 100000),
      service.getRecords({
        key: '_storage_envs', title: 'Environments', logicalName: 'admin_environment', icon: 'Globe', filter: 'statecode eq 0',
        columns: [
          { field: 'admin_displayname', label: 'Name', type: 'text' },
          { field: 'admin_environmentsku', label: 'Typ', type: 'text' },
          { field: 'admin_environmentdeleted', label: 'Gelöscht', type: 'boolean' }
        ]
      }, 100000)
    ]);
    const rows = capRes.records;
    const envs = envRes.records;
    const byId: { [id: string]: IEnvStorage } = {};
    envs.filter(e => e.admin_environmentdeleted !== true).forEach(e => {
      const id = String(e.admin_environmentid || '').toLowerCase();
      byId[id] = {
        id, name: String(e.admin_displayname || id),
        type: String(e['admin_environmentsku@OData.Community.Display.V1.FormattedValue'] || e.admin_environmentsku || ''), db: 0, file: 0, log: 0
      };
    });
    rows.forEach(r => {
      const env = byId[String(r._admin_environment_value || '').toLowerCase()];
      if (!env) {
        return;
      }
      const type = String(r['admin_capacitytype@OData.Community.Display.V1.FormattedValue'] || r.admin_capacitytype || '').toLowerCase();
      const unit = String(r.admin_capacityunit || 'MB').toUpperCase();
      const value = Number(r.admin_actualconsumption) || 0;
      const mb = unit === 'GB' ? value * 1024 : value;
      if (/database|datenbank/.test(type)) { env.db += mb; } else if (/file|datei/.test(type)) { env.file += mb; } else if (/log/.test(type)) { env.log += mb; }
    });
    out.envs = Object.keys(byId).map(k => byId[k]);
  } catch (e) {
    out.envProblem = `Verbrauch je Environment aus dem CoE nicht lesbar: ${e instanceof Error ? e.message : String(e)}`;
  }
  return out;
}

export function gb(mb: number | undefined, digits: number = 1): string {
  if (mb === undefined) {
    return '–';
  }
  return (mb / 1024).toLocaleString('de-AT', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}
