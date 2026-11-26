import { IColumnConfig, IRecord } from './ITableConfig';
import { FORMATTED } from '../utils/format';
import { IInvItem, ILiveEnvironment } from '../services/LiveInventoryService';

/**
 * Live-Abgleich: Das CoE bleibt die Grundlage (alle Spalten, Nutzung, Verlauf, Beziehungen). Die Live-Daten
 *  - überschreiben den aktuellen Zustand (Name, Geändert, Quarantäne, Besitzer, Managed, Zustand),
 *  - ergänzen Ressourcen, die das CoE noch nicht kennt („Neu (nur live)“),
 *  - markieren CoE-Einträge, die es live nicht mehr gibt („Nicht im Live-Inventar“) – sie bleiben sichtbar.
 * Neu/fehlend nur, wenn beide Seiten vollständig geladen sind und die Tabelle das ganze Inventar zeigt.
 */
export const LIVE_FIELD = 'coe_live';
export const LIVE_STATUS = {
  same: 'Aktuell',
  changed: 'Geändert (live)',
  added: 'Neu (nur live)',
  missing: 'Nicht im Live-Inventar'
};
export const LIVE_COLUMN: IColumnConfig = {
  field: LIVE_FIELD, label: 'Live-Abgleich', type: 'status',
  statusMap: { [LIVE_STATUS.same]: 'success', [LIVE_STATUS.changed]: 'accent', [LIVE_STATUS.added]: 'info', [LIVE_STATUS.missing]: 'warning' }
};
/** Live: Connectors bzw. Änderungen gegenüber dem CoE (für den Tooltip/Export) */
export const LIVE_NOTE_FIELD = 'coe_livenote';

export type LiveArea = 'apps' | 'flows' | 'environments';

/** Felder, die der Live-Abgleich setzt – sie haben Vorrang vor nachgeladenen CoE-Werten (z. B. für die Aktionen). */
export const LIVE_OVERRIDE_FIELDS = ['admin_displayname', 'admin_appisquarantined', 'admin_quarantineapp', 'admin_appmodifiedon', 'admin_flowmodifiedon',
  'admin_governanceprotectionlevel', 'admin_environmentruntimestate', 'coe_managed', 'coe_state'];

export interface ILiveStats { same: number; changed: number; added: number; missing: number }

export interface ILiveAreaState {
  status: 'loading' | 'ready' | 'error' | 'off';
  /** Grund, warum nicht live (Rolle, Freigabe, Fehler) */
  problem?: string;
  items?: Map<string, IInvItem>;
  envs?: Map<string, ILiveEnvironment>;
  truncated?: boolean;
  loadedAt?: Date;
}

/** Nur Filter der Standard-Inventartabellen erlauben „neu“/„fehlend“ (sonst würden gefilterte Tabellen verfälscht). */
export function isInventoryFilter(filter: string | undefined): boolean {
  const rest = (filter || '')
    .replace(/statecode eq 0/gi, '')
    .replace(/admin_(app|flow|environment)deleted ne true/gi, '')
    .replace(/admin_excusefrominventory ne true/gi, '')
    .replace(/\band\b/gi, '')
    .replace(/[()\s]/g, '');
  return rest === '';
}

export function areaOf(logicalName: string): LiveArea | undefined {
  return logicalName === 'admin_app' ? 'apps' : logicalName === 'admin_flow' ? 'flows' : logicalName === 'admin_environment' ? 'environments' : undefined;
}

const lower = (v: unknown): string => (v === undefined || v === null ? '' : String(v)).toLowerCase();

/** Wert setzen; formatierte Anzeige entfernen, damit der neue Rohwert angezeigt wird. */
function setValue(r: IRecord, field: string, value: unknown, formatted?: string): void {
  r[field] = value;
  if (formatted !== undefined) {
    r[field + FORMATTED] = formatted;
  } else {
    delete r[field + FORMATTED];
  }
}

const sameDay = (a: unknown, b: unknown): boolean => {
  const da = Date.parse(String(a || ''));
  const db = Date.parse(String(b || ''));
  return isNaN(da) || isNaN(db) ? !a === !b : Math.abs(da - db) < 60000;
};

export function recordKey(logicalName: string, r: IRecord): string {
  if (logicalName === 'admin_app') {
    return lower(r.admin_appid);
  }
  if (logicalName === 'admin_flow') {
    return lower(r.admin_flowidstring || r.admin_flowid);
  }
  return lower(r.admin_environmentname || r.admin_environmentid);
}

/** Namen von Personen und Environments aus den CoE-Daten (für Live-Einträge ohne Namen). */
export interface ILiveNames { people: { [id: string]: string }; environments: { [id: string]: string } }

export function namesFrom(records: IRecord[], extraEnvs?: Map<string, ILiveEnvironment>): ILiveNames {
  const people: { [id: string]: string } = {};
  const environments: { [id: string]: string } = {};
  records.forEach(r => {
    ['_admin_appowner_value', '_admin_derivedowner_value', '_admin_flowcreator_value', '_admin_maker_value'].forEach(f => {
      const id = lower(r[f]);
      const name = r[f + FORMATTED];
      if (id && name) {
        people[id] = String(name);
      }
    });
    const envId = lower(r.admin_appenvironmentid || r.admin_flowenvironmentid);
    const envName = r['_admin_appenvironment_value' + FORMATTED] || r['_admin_flowenvironment_value' + FORMATTED];
    if (envId && envName) {
      environments[envId] = String(envName);
    }
  });
  if (extraEnvs) {
    extraEnvs.forEach((e, id) => { if (e.displayName) { environments[id] = e.displayName; } });
  }
  return { people, environments };
}

function appOverlay(r: IRecord, it: IInvItem, names: ILiveNames): string[] {
  const changes: string[] = [];
  if (it.displayName && it.displayName !== r.admin_displayname) {
    changes.push('Name');
    setValue(r, 'admin_displayname', it.displayName);
  }
  if (it.isQuarantined !== undefined && it.isQuarantined !== (r.admin_appisquarantined === true)) {
    changes.push(it.isQuarantined ? 'in Quarantäne' : 'nicht mehr in Quarantäne');
    setValue(r, 'admin_appisquarantined', it.isQuarantined);
    setValue(r, 'admin_quarantineapp', it.isQuarantined);
  }
  if (it.lastModifiedAt && !sameDay(it.lastModifiedAt, r.admin_appmodifiedon)) {
    changes.push('geändert');
    setValue(r, 'admin_appmodifiedon', it.lastModifiedAt);
  }
  if (it.ownerId && lower(r._admin_appowner_value) && it.ownerId !== lower(r._admin_appowner_value)) {
    changes.push('Besitzer');
    setValue(r, '_admin_appowner_value', it.ownerId, names.people[it.ownerId] || `Objekt-ID ${it.ownerId}`);
  }
  return changes;
}

function flowOverlay(r: IRecord, it: IInvItem): string[] {
  const changes: string[] = [];
  if (it.displayName && it.displayName !== r.admin_displayname) {
    changes.push('Name');
    setValue(r, 'admin_displayname', it.displayName);
  }
  if (it.lastModifiedAt && !sameDay(it.lastModifiedAt, r.admin_flowmodifiedon)) {
    changes.push('geändert');
    setValue(r, 'admin_flowmodifiedon', it.lastModifiedAt);
  }
  return changes;
}

/** Environment: Managed und Zustand – im Reiter „Environments“ und in „Betroffen ab 15.12.“ (Felder coe_managed/coe_state). */
function envOverlay(r: IRecord, e: ILiveEnvironment | undefined, it: IInvItem | undefined): string[] {
  const changes: string[] = [];
  const name = (e && e.displayName) || (it && it.displayName);
  if (name && name !== r.admin_displayname) {
    changes.push('Name');
    setValue(r, 'admin_displayname', name);
  }
  const managed = e && e.protectionLevel ? !/^basic$/i.test(e.protectionLevel) : it ? it.isManaged : undefined;
  if (managed !== undefined) {
    if ('coe_managed' in r) {
      if ((r.coe_managed === 'Ja') !== managed) {
        changes.push(managed ? 'jetzt Managed' : 'nicht mehr Managed');
        r.coe_managed = managed ? 'Ja' : 'Nein';
      }
    } else {
      const level = String(r.admin_governanceprotectionlevel || '');
      if (!!level && !/^basic$/i.test(level) !== managed) {
        changes.push(managed ? 'jetzt Managed' : 'nicht mehr Managed');
        setValue(r, 'admin_governanceprotectionlevel', managed ? (e && e.protectionLevel) || 'Standard' : 'Basic');
      }
    }
  }
  const state = e && e.state ? e.state : '';
  if (state) {
    const disabled = /disabled|deactivated|suspended/i.test(state);
    if ('coe_state' in r) {
      const was = r.coe_state === 'Deaktiviert';
      if (was !== disabled) {
        changes.push(disabled ? 'deaktiviert' : 'aktiviert');
        r.coe_state = disabled ? 'Deaktiviert' : 'Aktiv';
      }
    } else if (r.admin_environmentruntimestate !== undefined) {
      const was = /^disabled$/i.test(String(r.admin_environmentruntimestate || ''));
      if (was !== disabled) {
        changes.push(disabled ? 'deaktiviert' : 'aktiviert');
        setValue(r, 'admin_environmentruntimestate', disabled ? 'Disabled' : 'Enabled');
      }
    }
  }
  return changes;
}

/** Neue Datensätze (nur live) in CoE-Form bringen. */
function newRecord(logicalName: string, it: IInvItem, names: ILiveNames, env?: ILiveEnvironment): IRecord {
  const r: IRecord = { __liveOnly: true };
  const envName = names.environments[it.environmentId] || it.environmentId;
  const owner = it.ownerId || it.createdBy;
  if (logicalName === 'admin_app') {
    r.admin_appid = it.id;
    setValue(r, 'admin_displayname', it.displayName);
    r.admin_appenvironmentid = it.environmentId;
    setValue(r, '_admin_appenvironment_value', it.environmentId, envName);
    if (owner) {
      setValue(r, '_admin_appowner_value', owner, names.people[owner] || `Objekt-ID ${owner}`);
    }
    r.admin_appcreatedon = it.createdAt;
    r.admin_appmodifiedon = it.lastModifiedAt || it.createdAt;
    if (it.isQuarantined !== undefined) {
      r.admin_appisquarantined = it.isQuarantined;
    }
    setValue(r, 'admin_powerappstype', it.type, /modeldriven/.test(it.type) ? 'Model-driven' : /codeapps/.test(it.type) ? 'Code-App' : /canvasapps/.test(it.type) ? 'Canvas' : 'App');
  } else if (logicalName === 'admin_flow') {
    r.admin_flowid = it.id;
    r.admin_flowidstring = it.id;
    setValue(r, 'admin_displayname', it.displayName);
    r.admin_flowenvironmentid = it.environmentId;
    setValue(r, '_admin_flowenvironment_value', it.environmentId, envName);
    if (owner) {
      setValue(r, '_admin_derivedowner_value', owner, names.people[owner] || `Objekt-ID ${owner}`);
      setValue(r, '_admin_flowcreator_value', it.createdBy || owner, names.people[it.createdBy || owner] || `Objekt-ID ${it.createdBy || owner}`);
    }
    r.admin_flowmodifiedon = it.lastModifiedAt || it.createdAt;
    if (it.trigger) {
      r.cr5d5_flowtrigger = it.triggerOperation ? `${it.trigger} – ${it.triggerOperation}` : it.trigger;
    }
  } else {
    r.admin_environmentid = it.id;
    r.admin_environmentname = it.id;
    setValue(r, 'admin_displayname', (env && env.displayName) || it.displayName);
    setValue(r, 'admin_environmentsku', (env && env.type) || it.environmentType);
    r.admin_environmentcreatedon = (env && env.createdDateTime) || it.createdAt;
    if (env && env.url) {
      r.admin_environmentcdsinstanceurl = env.url;
    }
    if (env && env.state) {
      setValue(r, 'admin_environmentruntimestate', /disabled/i.test(env.state) ? 'Disabled' : 'Enabled');
    }
    const managed = env && env.protectionLevel ? !/^basic$/i.test(env.protectionLevel) : it.isManaged;
    if (managed !== undefined) {
      setValue(r, 'admin_governanceprotectionlevel', managed ? 'Standard' : 'Basic');
    }
  }
  r[LIVE_FIELD] = LIVE_STATUS.added;
  r[LIVE_NOTE_FIELD] = 'Nur live – das CoE kennt diese Ressource noch nicht (nächster CoE-Sync).';
  return r;
}

export interface IOverlayInput {
  logicalName: string;
  records: IRecord[];
  live: ILiveAreaState;
  /** Liste der Environments (für Namen und Zustand) */
  envs?: Map<string, ILiveEnvironment>;
  /** Darf ergänzt bzw. als fehlend markiert werden? */
  complete: boolean;
  /** Bekannte IDs im CoE (auch ausgefilterte) – wird nicht als „neu“ ergänzt */
  knownIds?: Set<string>;
  /** Ressourcen mit eben ausgeführter Aktion: CoE-Stand ist aktueller als das Inventar (bis 15 Min. Verzögerung) */
  skipIds?: Set<string>;
}

/** CoE-Datensätze mit dem Live-Stand abgleichen (gibt neue Objekte zurück, die CoE-Daten bleiben unverändert). */
export function overlay(input: IOverlayInput): { records: IRecord[]; stats: ILiveStats } {
  const { logicalName, live } = input;
  const stats: ILiveStats = { same: 0, changed: 0, added: 0, missing: 0 };
  const items = live.items || new Map<string, IInvItem>();
  const names = namesFrom(input.records, input.envs);
  const seen = new Set<string>();
  const out: IRecord[] = input.records.map(src => {
    const r: IRecord = { ...src };
    const key = recordKey(logicalName, r);
    if (key && input.skipIds && input.skipIds.has(key)) {
      seen.add(key);
      r[LIVE_FIELD] = LIVE_STATUS.same;
      r[LIVE_NOTE_FIELD] = 'Eben per Aktion geändert – das Live-Inventar zieht in bis zu 15 Minuten nach; angezeigt wird der Stand nach der Aktion.';
      stats.same++;
      return r;
    }
    const it = items.get(key);
    const env = logicalName === 'admin_environment' && input.envs ? input.envs.get(key) : undefined;
    if (!it && !env) {
      if (input.complete && key) {
        r[LIVE_FIELD] = LIVE_STATUS.missing;
        r[LIVE_NOTE_FIELD] = 'Im Live-Inventar nicht gefunden – vermutlich gelöscht (das CoE zieht das beim nächsten Sync nach).';
        stats.missing++;
      }
      return r;
    }
    seen.add(key);
    const changes = logicalName === 'admin_app' ? appOverlay(r, it!, names)
      : logicalName === 'admin_flow' ? flowOverlay(r, it!)
        : envOverlay(r, env, it);
    r[LIVE_FIELD] = changes.length ? LIVE_STATUS.changed : LIVE_STATUS.same;
    r[LIVE_NOTE_FIELD] = changes.length ? `Live: ${changes.join(', ')}` : 'Live bestätigt';
    if (changes.length) {
      stats.changed++;
    } else {
      stats.same++;
    }
    return r;
  });
  if (input.complete) {
    const addFrom = (id: string, it: IInvItem | undefined, env?: ILiveEnvironment): void => {
      if (seen.has(id) || (input.knownIds && input.knownIds.has(id))) {
        return;
      }
      seen.add(id);
      const base: IInvItem = it || { id, type: 'microsoft.powerplatform/environments', displayName: env ? env.displayName : id, environmentId: id, ownerId: '', createdBy: '', createdAt: env ? env.createdDateTime : '', lastModifiedAt: '', raw: {} };
      out.push(newRecord(logicalName, base, names, env));
      stats.added++;
    };
    if (logicalName === 'admin_environment' && input.envs && input.envs.size) {
      input.envs.forEach((env, id) => addFrom(id, items.get(id), env));
    } else {
      items.forEach((it, id) => addFrom(id, it));
    }
  }
  return { records: out, stats };
}
