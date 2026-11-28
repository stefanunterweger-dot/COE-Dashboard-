import { IRecord } from './ITableConfig';
import { displayValue, rawValue } from '../utils/format';

/**
 * Cockpit-Aktionen: Das Cockpit führt sie direkt mit den Rechten der angemeldeten Person aus (Rolle
 * „Power Platform Administrator“, siehe services/DirectActionService.ts) und protokolliert jede in der SharePoint-Liste
 * „CoE Cockpit Aktionen“. Nur „Solution löschen“ braucht Rechte im Environment und geht als Auftrag an den Flow im
 * CoE-Environment (Dienstkonto). Sichtbar sind die Aktionen nur für Mitglieder der SharePoint-Gruppe (Standard
 * „CoE Cockpit Admins“) mit aktiver Admin-Rolle.
 */
export const DEFAULT_ACTION_GROUP = 'CoE Cockpit Admins';
export const ACTION_LIST_TITLE = 'CoE Cockpit Aktionen';
export const ACTION_LOG_KEY = 'actions:log';

/** Technische Codes (stehen im Aktionsprotokoll und werden vom Flow ausgewertet – nicht übersetzen). */
export type ActionCode = 'quarantine' | 'unquarantine' | 'delete' | 'setowner' | 'addviewer' | 'addeditor' | 'removeuser' | 'flowoff' | 'flowon'
  | 'envmanaged' | 'envdisable' | 'envenable' | 'soldelete' | 'dlpupdate' | 'dlpcreate' | 'dlpdelete';
export type ResourceType = 'app' | 'flow' | 'environment' | 'solution' | 'dlppolicy';

export const ACTION_LABEL: { [c in ActionCode]: string } = {
  quarantine: 'In Quarantäne stellen',
  unquarantine: 'Aus Quarantäne holen',
  delete: 'Löschen',
  setowner: 'Besitzer ändern',
  addviewer: 'Freigabe hinzufügen (Benutzer)',
  addeditor: 'Freigabe hinzufügen (Mitbesitzer)',
  removeuser: 'Freigabe entfernen',
  flowoff: 'Flow ausschalten',
  flowon: 'Flow einschalten',
  envmanaged: 'Auf Managed umstellen',
  envdisable: 'Environment deaktivieren',
  envenable: 'Environment aktivieren',
  soldelete: 'Solution löschen',
  dlpupdate: 'DLP-Richtlinie ändern',
  dlpcreate: 'DLP-Richtlinie anlegen',
  dlpdelete: 'DLP-Richtlinie löschen'
};

export const ACTION_ICON: { [c in ActionCode]: string } = {
  quarantine: 'Blocked2', unquarantine: 'Unlock', delete: 'Delete', setowner: 'Contact',
  addviewer: 'AddFriend', addeditor: 'AddFriend', removeuser: 'UserRemove', flowoff: 'Pause', flowon: 'Play',
  envmanaged: 'Shield', envdisable: 'Blocked2', envenable: 'Play', soldelete: 'Delete',
  dlpupdate: 'Shield', dlpcreate: 'Copy', dlpdelete: 'Delete'
};

/** Hinweis im Dialog – was die Aktion bewirkt. */
export const ACTION_HINT: { [c in ActionCode]: string } = {
  quarantine: 'Benutzer, für die die App freigegeben ist, können sie nicht mehr starten. Maker können sie weiter bearbeiten. Jederzeit umkehrbar.',
  unquarantine: 'Benutzer können die App wieder starten.',
  delete: 'Die Ressource wird im Tenant gelöscht – nicht nur aus dem Inventar. Das lässt sich nicht rückgängig machen.',
  setowner: 'Die gewählte Person wird Besitzer. Der bisherige Besitzer wird Mitbesitzer.',
  addviewer: 'Die gewählte Person oder Gruppe darf die Ressource verwenden.',
  addeditor: 'Die gewählte Person oder Gruppe darf die Ressource bearbeiten und freigeben.',
  removeuser: 'Die gewählte Freigabe wird entfernt. Den Besitzer kann man nicht entfernen.',
  flowoff: 'Der Flow wird angehalten und startet nicht mehr, bis er wieder eingeschaltet wird.',
  flowon: 'Der Flow startet wieder bei seinem Auslöser bzw. nach Zeitplan.',
  envmanaged: 'Das Environment wird zum Managed Environment und bleibt nach dem 15.12.2026 erhalten. Wer dort Apps oder Flows nutzt, braucht dann eine Premium-Lizenz (ab Februar 2027 werden Apps für Benutzer ohne Lizenz blockiert).',
  envdisable: 'Das Environment wird deaktiviert: Apps, Flows und Bots darin laufen nicht mehr, niemand kann es verwenden. Die Daten bleiben erhalten – ein Admin kann es wieder aktivieren.',
  envenable: 'Das Environment wird wieder aktiviert; Apps, Flows und Bots laufen wieder.',
  soldelete: 'Gelöscht wird nur die Solution selbst (der „Container“). Apps, Flows, Tabellen und alle anderen Komponenten bleiben im Environment erhalten '
    + 'und gehören danach nur noch zur Default Solution. Das lässt sich nicht rückgängig machen.',
  dlpupdate: 'Die DLP-Richtlinie wird im Tenant geändert. Apps und Flows, die danach gegen sie verstoßen, werden blockiert bzw. ausgesetzt.',
  dlpcreate: 'Eine neue DLP-Richtlinie wird im Tenant angelegt und gilt sofort für die gewählten Environments.',
  dlpdelete: 'Die DLP-Richtlinie wird im Tenant gelöscht. Ihre Regeln gelten danach nicht mehr. Das lässt sich nicht rückgängig machen.'
};

export type ActionStatus = 'Angefordert' | 'In Arbeit' | 'Erledigt' | 'Fehler' | 'Abgelehnt';
export const FINAL_STATUS: ActionStatus[] = ['Erledigt', 'Fehler', 'Abgelehnt'];

export interface IActionRequest {
  action: ActionCode;
  resourceType: ResourceType;
  /** Power-Platform-ID der App bzw. des Flows. */
  resourceId: string;
  /** Primärschlüssel im CoE (für das Aktualisieren des Inventars). */
  coeRecordId: string;
  resourceName: string;
  /** Power-Platform-ID des Environments. */
  environmentId: string;
  /** Zielperson/-gruppe: Entra-Objekt-ID oder E-Mail-Adresse (nur Besitzer/Freigaben). */
  targetUser?: string;
  targetName?: string;
  targetIsGroup?: boolean;
  reason: string;
}

export interface IActionItem extends IActionRequest {
  id: number;
  status: ActionStatus;
  result: string;
  created: string;
  completedAt: string;
  requestedBy: string;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * ID einer Solution im Environment (Tabelle solutions). Das CoE vergibt für admin_solution eine eigene,
 * zufällige ID; die echte Solution-ID steht in admin_solutionenvtguid bzw. hinten in
 * admin_solutionenvtobjectguidstring („<Environment-ID>_<Solution-ID>“). Leer, wenn nicht erfasst.
 */
export function envSolutionId(r: IRecord): string {
  const direct = String(r.admin_solutionenvtguid || '').trim();
  if (GUID_RE.test(direct)) {
    return direct;
  }
  const obj = String(r.admin_solutionenvtobjectguidstring || '').trim();
  const tail = obj.slice(obj.lastIndexOf('_') + 1);
  return GUID_RE.test(tail) ? tail : '';
}

/** Ressource einer Tabellenzeile (Apps/Flows aus dem CoE). */
export function resourceOf(logicalName: string, r: IRecord): { type: ResourceType; id: string; coeId: string; name: string; env: string } | undefined {
  if (logicalName === 'admin_app') {
    const id = String(r.admin_appid || '');
    return { type: 'app', id, coeId: id, name: displayValue(r, 'admin_displayname'), env: rawValue(r, 'admin_appenvironmentid') };
  }
  if (logicalName === 'admin_environment') {
    const coeId = String(r.admin_environmentid || '');
    const pid = rawValue(r, 'admin_environmentname');
    const id = /^(default-)?[0-9a-f]{8}-/i.test(pid) ? pid : coeId;
    return { type: 'environment', id, coeId, name: displayValue(r, 'admin_displayname'), env: id };
  }
  if (logicalName === 'admin_solution') {
    const coeId = String(r.admin_solutionid || '');
    // ID der Solution im Environment (nicht die CoE-ID – die vergibt das CoE zufällig)
    const id = envSolutionId(r);
    // Environment = Dataverse-URL des Environments (wird beim Öffnen nachgeladen: __envurl)
    return { type: 'solution', id, coeId, name: displayValue(r, 'admin_solutiondisplayname') || displayValue(r, 'admin_name'), env: String(r.__envurl || '') };
  }
  if (logicalName === 'admin_flow') {
    const coeId = String(r.admin_flowid || '');
    return { type: 'flow', id: rawValue(r, 'admin_flowidstring') || coeId, coeId, name: displayValue(r, 'admin_displayname'), env: rawValue(r, 'admin_flowenvironmentid') };
  }
  return undefined;
}

/** Environment bereits Managed? (Reiter „Betroffen ab 15.12.“ bzw. CoE „Governance Protection Level“) */
export function isManagedEnv(r: IRecord): boolean {
  const dev = String(r.coe_managed || '');
  if (dev) {
    return dev === 'Ja';
  }
  const level = rawValue(r, 'admin_governanceprotectionlevel');
  return !!level && !/^basic$/i.test(level);
}

/** Environment deaktiviert? */
export function isEnvDisabled(r: IRecord): boolean {
  return String(r.coe_state || '') === 'Deaktiviert' || /^disabled$/i.test(rawValue(r, 'admin_environmentruntimestate'));
}

/** Ist die App in Quarantäne? (CoE kennt zwei Felder: „App Is Quarantined“ und „Quarantine App“.) */
export function isQuarantined(r: IRecord): boolean {
  return r.admin_appisquarantined === true || r.admin_quarantineapp === true;
}

/** Ist der Flow eingeschaltet? */
export function isFlowOn(r: IRecord): boolean {
  return /^started$/i.test(rawValue(r, 'admin_flowstate'));
}

/**
 * Teil einer Solution? CoE: „Is Solution App“ (admin_issolutionapp) bzw. „Is Solution Flow“ (admin_issolutionflow).
 * true = in einer Solution, false = nicht, undefined = unbekannt (Spalte fehlt) – dann wird nicht gelöscht.
 */
export function inSolution(type: ResourceType, r: IRecord): boolean | undefined {
  const v = type === 'app' ? r.admin_issolutionapp : r.admin_issolutionflow;
  return v === true ? true : v === false ? false : undefined;
}

/** Grund, warum nicht gelöscht werden darf (leer = darf). Regel: nie, wenn Teil einer Solution. */
export function deleteBlocked(type: ResourceType, r: IRecord, deleteRequiresStop: boolean): string {
  const sol = inSolution(type, r);
  if (sol === true) {
    return `${type === 'app' ? 'Die App' : 'Der Flow'} ist Teil einer Solution – Löschen nur über die Solution (ALM), nicht hier.`;
  }
  if (sol === undefined) {
    return 'Ob die Ressource Teil einer Solution ist, ist im CoE nicht erfasst – Löschen deshalb gesperrt.';
  }
  if (deleteRequiresStop && type === 'app' && !isQuarantined(r)) {
    return 'Löschen ist erst möglich, wenn die App in Quarantäne ist (Einstellung „Löschen erst nach Quarantäne“).';
  }
  if (deleteRequiresStop && type === 'flow' && isFlowOn(r)) {
    return 'Löschen ist erst möglich, wenn der Flow ausgeschaltet ist (Einstellung „Löschen erst nach Quarantäne/Ausschalten“).';
  }
  return '';
}

/** System- und Microsoft-Solutions (nie löschen). */
const SYSTEM_SOLUTIONS = ['default', 'active', 'basic', 'system', 'crmservicesolution'];

/** Grund, warum eine Solution nicht gelöscht werden darf (leer = darf). Nur nicht verwaltete Solutions. */
export function solutionDeleteBlocked(r: IRecord): string {
  const managed = r.admin_solutionismanaged;
  const name = String(r.admin_name || '').toLowerCase();
  const publisher = String(r.admin_solutionpublisher || '');
  if (SYSTEM_SOLUTIONS.indexOf(name) >= 0 || /^microsoft/i.test(publisher) || /^msdyn|^msft/i.test(name)) {
    return 'System- bzw. Microsoft-Solution – wird nicht gelöscht.';
  }
  if (managed === true) {
    return 'Verwaltete Solution: Löschen wäre eine Deinstallation mit allen Komponenten und Daten – bitte über das Admin Center bzw. die ALM-Pipeline.';
  }
  if (managed !== false) {
    return 'Ob die Solution verwaltet ist, ist im CoE nicht erfasst – Löschen deshalb gesperrt.';
  }
  if (!r.__envurl) {
    return 'Für das Environment der Solution ist im CoE keine Dataverse-URL erfasst.';
  }
  if (!envSolutionId(r)) {
    return 'Die ID der Solution im Environment ist im CoE nicht erfasst (Spalte „Solution Envt Guid“) – Löschen deshalb gesperrt.';
  }
  return '';
}

/** Verfügbare Aktionen je Ressource und Zustand. */
export function actionsFor(type: ResourceType, r: IRecord, deleteRequiresStop: boolean): { code: ActionCode; disabled?: string }[] {
  if (type === 'solution') {
    return [{ code: 'soldelete', disabled: solutionDeleteBlocked(r) || undefined }];
  }
  if (type === 'environment') {
    const managed = isManagedEnv(r);
    return [
      { code: 'envmanaged', disabled: managed ? 'Das Environment ist bereits Managed.' : undefined },
      { code: isEnvDisabled(r) ? 'envenable' : 'envdisable' }
    ];
  }
  const del = deleteBlocked(type, r, deleteRequiresStop) || undefined;
  if (type === 'app' && rawValue(r, 'admin_powerappstype') === '597910001') {
    // Modellgesteuerte App: im CoE ohne Power-Apps-App-ID – Quarantäne, Besitzer und Freigaben gibt es dafür nicht
    const why = 'Modellgesteuerte App – Zugriff über Sicherheitsrollen im Environment, Änderungen über die Solution.';
    return (['quarantine', 'setowner', 'addviewer', 'removeuser', 'delete'] as ActionCode[]).map(code => ({ code, disabled: why }));
  }
  if (type === 'app') {
    return [
      { code: isQuarantined(r) ? 'unquarantine' : 'quarantine' },
      { code: 'setowner' },
      { code: 'addviewer' },
      { code: 'removeuser' },
      { code: 'delete', disabled: del }
    ];
  }
  return [
    { code: isFlowOn(r) ? 'flowoff' : 'flowon' },
    { code: 'setowner' },
    { code: 'addeditor' },
    { code: 'removeuser' },
    { code: 'delete', disabled: del }
  ];
}
