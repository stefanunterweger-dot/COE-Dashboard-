import { IColumnConfig, ITableConfig } from './ITableConfig';
import { getDefaultTables } from './defaultTables';
import { combineFilters } from '../utils/format';
import { environmentLinkId, IEnvironmentInfo } from '../services/DataverseService';

export const DEFAULT_ENV_PREFIX = 'default:';

export function isDefaultEnvKey(key: string | undefined): boolean {
  return !!key && key.indexOf(DEFAULT_ENV_PREFIX) === 0;
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function clone<T>(v: T): T {
  return JSON.parse(JSON.stringify(v)) as T;
}

/** Spalte „Umgebung“ ist in der Default-Ansicht überflüssig (ausgeblendete ID-Spalten bleiben für Links). */
function withoutEnvironmentColumn(columns: IColumnConfig[]): IColumnConfig[] {
  return columns.filter(c => c.hidden || !/environment/i.test(c.field));
}

function baseTable(tables: ITableConfig[], logicalName: string): ITableConfig {
  return clone(
    tables.filter(t => t.logicalName === logicalName)[0] ||
    getDefaultTables().filter(t => t.logicalName === logicalName)[0]
  );
}

/**
 * Erzeugt die drei Tabellen der Ansicht „Default Environment“:
 * Apps und Flows aus dem CoE-Inventar (gefiltert auf die Umgebung),
 * Solutions live aus dem Dataverse des Default Environments.
 */
export function buildDefaultEnvironmentTables(
  tables: ITableConfig[],
  env: IEnvironmentInfo,
  dataverseUrlOverride?: string
): ITableConfig[] {
  const envId = env.id.replace(/^default-/i, '');
  const validId = GUID.test(envId);
  const notResolvable = validId ? undefined : `Die Umgebungs-ID „${env.id}“ ist keine gültige GUID.`;

  const apps = baseTable(tables, 'admin_app');
  const flows = baseTable(tables, 'admin_flow');

  const appsTable: ITableConfig = {
    ...apps,
    key: `${DEFAULT_ENV_PREFIX}apps`,
    title: 'Apps',
    showInKpis: false,
    filter: validId ? combineFilters(apps.filter, `_admin_appenvironment_value eq ${envId}`) : apps.filter,
    columns: withoutEnvironmentColumn(apps.columns),
    unavailable: notResolvable
  };

  const flowsTable: ITableConfig = {
    ...flows,
    key: `${DEFAULT_ENV_PREFIX}flows`,
    title: 'Flows',
    showInKpis: false,
    filter: validId ? combineFilters(flows.filter, `_admin_flowenvironment_value eq ${envId}`) : flows.filter,
    columns: withoutEnvironmentColumn(flows.columns),
    unavailable: notResolvable
  };

  const dataverseUrl = (dataverseUrlOverride || '').trim() || env.dataverseUrl;
  const solutionsTable: ITableConfig = {
    key: `${DEFAULT_ENV_PREFIX}solutions`,
    title: 'Solutions',
    logicalName: 'solution',
    icon: 'Package',
    showInKpis: false,
    environmentUrl: dataverseUrl,
    filter: 'isvisible eq true',
    defaultSort: { field: 'installedon', desc: true },
    columns: [
      { field: 'friendlyname', label: 'Solution', type: 'title' },
      { field: 'uniquename', label: 'Name (eindeutig)', type: 'text' },
      { field: 'version', label: 'Version', type: 'text' },
      { field: 'ismanaged', label: 'Verwaltet', type: 'boolean' },
      { field: '_publisherid_value', label: 'Herausgeber', type: 'text' },
      { field: 'installedon', label: 'Installiert', type: 'date' },
      { field: 'modifiedon', label: 'Geändert', type: 'date' },
      { field: 'description', label: 'Beschreibung', type: 'text', hidden: true }
    ],
    links: [
      {
        label: 'In Power Apps öffnen',
        url: `https://make.powerapps.com/environments/${encodeURIComponent(environmentLinkId(env))}/solutions/{solutionid}`
      }
    ],
    unavailable: dataverseUrl
      ? undefined
      : 'Für das Default Environment ist keine Dataverse-URL bekannt. Solutions gibt es nur in Umgebungen mit Dataverse-Datenbank. ' +
        'Falls vorhanden, trage die URL in den Webpart-Eigenschaften unter „Default Environment“ ein.'
  };

  return [appsTable, flowsTable, solutionsTable];
}
