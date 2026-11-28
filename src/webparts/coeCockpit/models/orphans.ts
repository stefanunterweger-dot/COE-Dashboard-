import { IRecord, ITableConfig } from './ITableConfig';
import { DataverseService } from '../services/DataverseService';
import { displayValue, rawValue } from '../utils/format';

/**
 * Verwaiste Ressourcen aus dem CoE: Apps und Flows, deren Besitzer laut CoE nicht mehr im Unternehmen ist
 * („App is Orphaned“ / „Flow is Orphaned“ = Yes). Grundlage für die Neuzuweisung im Cockpit.
 */
export const ORPHANS_KEY = 'orphans';

export interface IOrphan {
  /** eindeutiger Schlüssel (Typ + CoE-ID) */
  key: string;
  type: 'app' | 'flow';
  logicalName: 'admin_app' | 'admin_flow';
  record: IRecord;
  name: string;
  environment: string;
  /** ehemaliger Besitzer */
  ownerId: string;
  ownerName: string;
  /** App: zuletzt gestartet; Flow: zuletzt geändert */
  last: string;
  inSolution: boolean | undefined;
  /** Flow: Status (Aktiv, Angehalten, Gesperrt) */
  state: string;
  /** Grund, warum sich der Eintrag nicht neu zuweisen lässt (z. B. modellgesteuerte App) */
  blocked?: string;
}

/** Modellgesteuerte Apps: im CoE ohne Power-Apps-App-ID – Besitzerwechsel über die Admin-API geht nicht. */
const MODEL_DRIVEN = '597910001';

const APP_TABLE: ITableConfig = {
  key: '_orphan_apps', title: 'Verwaiste Apps', logicalName: 'admin_app', icon: 'AppIconDefault', strictFilter: true,
  filter: "statecode eq 0 and admin_appdeleted ne true and cr5d5_appisorphaned eq 'Yes'",
  columns: [
    { field: 'admin_displayname', label: 'App', type: 'title' },
    { field: '_admin_appowner_value', label: 'Besitzer', type: 'person' },
    { field: '_admin_appenvironment_value', label: 'Environment', type: 'text' },
    { field: 'admin_appenvironmentid', label: 'Environment-ID', type: 'text' },
    { field: 'admin_applastlaunchedon', label: 'Zuletzt gestartet', type: 'date' },
    { field: 'admin_appmodifiedon', label: 'Geändert', type: 'date' },
    { field: 'admin_issolutionapp', label: 'In Solution', type: 'boolean' },
    { field: 'admin_appsharedusers', label: 'Freigegeben (User)', type: 'number' },
    { field: 'admin_powerappstype', label: 'App-Typ', type: 'choice' }
  ]
};

const FLOW_TABLE: ITableConfig = {
  key: '_orphan_flows', title: 'Verwaiste Flows', logicalName: 'admin_flow', icon: 'Flow', strictFilter: true,
  filter: "statecode eq 0 and admin_flowdeleted ne true and cr5d5_flowisorphaned eq 'Yes'",
  columns: [
    { field: 'admin_displayname', label: 'Flow', type: 'title' },
    { field: '_admin_derivedowner_value', label: 'Besitzer', type: 'person' },
    { field: '_admin_flowcreator_value', label: 'Ersteller', type: 'person' },
    { field: '_admin_flowenvironment_value', label: 'Environment', type: 'text' },
    { field: 'admin_flowenvironmentid', label: 'Environment-ID', type: 'text' },
    { field: 'admin_flowidstring', label: 'Flow-ID', type: 'text' },
    { field: 'admin_flowstate', label: 'Status', type: 'status' },
    { field: 'admin_flowmodifiedon', label: 'Geändert', type: 'date' },
    { field: 'admin_issolutionflow', label: 'In Solution', type: 'boolean' }
  ]
};

export interface IOrphanData { items: IOrphan[]; warnings: string[]; truncated: boolean; loadedAt: Date }

const bool = (v: unknown): boolean | undefined => (v === true ? true : v === false ? false : undefined);

export async function loadOrphans(service: DataverseService, maxRecords: number): Promise<IOrphanData> {
  const [apps, flows] = await Promise.all([service.getRecords(APP_TABLE, maxRecords), service.getRecords(FLOW_TABLE, maxRecords)]);
  const items: IOrphan[] = [];
  apps.records.forEach(r => {
    const owner = rawValue(r, '_admin_appowner_value');
    items.push({
      key: `app:${rawValue(r, apps.meta.primaryId)}`, type: 'app', logicalName: 'admin_app', record: r,
      name: displayValue(r, 'admin_displayname'), environment: displayValue(r, '_admin_appenvironment_value'),
      ownerId: owner.toLowerCase(), ownerName: displayValue(r, '_admin_appowner_value') || (owner ? `Objekt-ID ${owner}` : 'Unbekannt'),
      last: rawValue(r, 'admin_applastlaunchedon'), inSolution: bool(r.admin_issolutionapp), state: '',
      blocked: rawValue(r, 'admin_powerappstype') === MODEL_DRIVEN
        ? 'Modellgesteuerte App – Zugriff über Sicherheitsrollen im Environment, nicht über einen Besitzer.' : undefined
    });
  });
  flows.records.forEach(r => {
    const owner = rawValue(r, '_admin_derivedowner_value') || rawValue(r, '_admin_flowcreator_value');
    const ownerName = displayValue(r, '_admin_derivedowner_value') || displayValue(r, '_admin_flowcreator_value');
    items.push({
      key: `flow:${rawValue(r, flows.meta.primaryId)}`, type: 'flow', logicalName: 'admin_flow', record: r,
      name: displayValue(r, 'admin_displayname'), environment: displayValue(r, '_admin_flowenvironment_value'),
      ownerId: owner.toLowerCase(), ownerName: ownerName || (owner ? `Objekt-ID ${owner}` : 'Unbekannt'),
      last: rawValue(r, 'admin_flowmodifiedon'), inSolution: bool(r.admin_issolutionflow), state: displayValue(r, 'admin_flowstate')
    });
  });
  return { items, warnings: apps.warnings.concat(flows.warnings), truncated: apps.truncated || flows.truncated, loadedAt: new Date() };
}
