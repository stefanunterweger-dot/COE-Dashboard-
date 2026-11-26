import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import { ActionCode, IActionRequest } from '../models/actions';
import { DataverseService } from './DataverseService';
import { PowerPlatformApiService, PpApiError } from './PowerPlatformApiService';

/**
 * Cockpit-Aktionen direkt im Webpart – im Namen der angemeldeten Person, ohne Flow und ohne Dienstkonto.
 * Microsoft prüft bei jedem Aufruf selbst, ob die Person die Rolle „Power Platform Administrator“ aktiv hat (PIM).
 *
 *  - Apps:  Power-Apps-Admin-Schnittstelle (api.powerapps.com, dieselbe wie „Power Apps for Admins“ und das
 *           PowerShell-Modul Microsoft.PowerApps.Administration.PowerShell) – Token für „PowerApps Service“
 *  - Flows: Power-Automate-Admin-Schnittstelle (api.flow.microsoft.com, wie „Power Automate for Admins“)
 *           – Token für „Microsoft Flow Service“
 *  - Environments: Power Platform API (api.powerplatform.com, offiziell dokumentiert)
 *
 * Nur „Solution löschen“ braucht zusätzlich Rechte im jeweiligen Environment (System Administrator) und läuft
 * deshalb weiter über den Flow „CoE Cockpit | Aktion ausführen“.
 */
export const POWERAPPS_RESOURCE = 'https://service.powerapps.com/';
export const FLOW_RESOURCE = 'https://service.flow.microsoft.com/';
const POWERAPPS_ADMIN = 'https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments';
const FLOW_ADMIN = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments';

/** Aktionen, die weiter der Flow ausführt (brauchen Rechte im Environment, die die Admin-Rolle nicht hat). */
export const FLOW_ACTIONS: ActionCode[] = ['soldelete'];

export function runsDirect(code: ActionCode): boolean {
  return FLOW_ACTIONS.indexOf(code) < 0;
}

export type DirectErrorKind = 'consent' | 'forbidden' | 'notFound' | 'rejected' | 'failed';

export class DirectActionError extends Error {
  constructor(public kind: DirectErrorKind, message: string, public detail?: string) {
    super(message);
    Object.setPrototypeOf(this, DirectActionError.prototype);
  }
}

export interface IDirectResult {
  /** Ergebnis für das Aktionsprotokoll */
  text: string;
  /** Hinweis, z. B. wenn das CoE-Inventar nicht aktualisiert werden konnte */
  note: string;
}

const FORBIDDEN = 'Zugriff verweigert. Für diese Aktion brauchst du die aktive Rolle „Power Platform Administrator“ (bei PIM vorher aktivieren; '
  + 'nach dem Aktivieren kann es einige Minuten dauern – notfalls einmal ab- und wieder anmelden).';

const CONSENT: { [resource: string]: string } = {
  [POWERAPPS_RESOURCE]: 'Das Webpart darf die Power-Apps-Verwaltung noch nicht aufrufen. Ein Administrator muss die Berechtigung „User“ der Microsoft-App '
    + '„PowerApps Service“ für SharePoint freigeben (scripts/Grant-CoeCockpitPermissions.ps1 -Actions).',
  [FLOW_RESOURCE]: 'Das Webpart darf die Power-Automate-Verwaltung noch nicht aufrufen. Ein Administrator muss die Berechtigungen „User“, „Flows.Read.All“ und '
    + '„Flows.Manage.All“ der Microsoft-App „Microsoft Flow Service“ für SharePoint freigeben (scripts/Grant-CoeCockpitPermissions.ps1 -Actions).'
};

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function wait(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function errorText(e: unknown): string {
  const m = typeof e === 'object' && e !== null ? (e as { message?: unknown }).message : undefined;
  return typeof m === 'string' && m ? m : String(e);
}

function isConsentError(msg: string): boolean {
  return /AADSTS(65001|650057|500011|50105|90094|65004)|consent|resource principal|not been granted|has not consented/i.test(msg);
}

function messageOf(text: string): string {
  try {
    const b = JSON.parse(text) as { error?: { message?: string; code?: string }; message?: string };
    return (b.error && (b.error.message || b.error.code)) || b.message || text;
  } catch {
    return text;
  }
}

/** Fehler der Power Platform API in die Sprache dieses Dienstes übersetzen. */
function fromPpApi(e: unknown): DirectActionError {
  if (e instanceof PpApiError) {
    const kind: DirectErrorKind = e.kind === 'consent' ? 'consent' : e.kind === 'forbidden' ? 'forbidden' : e.kind === 'notFound' ? 'notFound' : 'failed';
    let msg = e.message;
    if (kind === 'forbidden') {
      msg = FORBIDDEN;
    } else if (kind === 'notFound') {
      msg = 'Das Environment wurde nicht gefunden (bereits gelöscht?).';
    } else if (e.status === 400 && /lifecycle|operation.*(progress|running|active)/i.test(e.detail || '')) {
      msg = 'Für dieses Environment läuft bei Microsoft gerade ein anderer Vorgang. Bitte in ein paar Minuten erneut versuchen.';
    } else if ((e.status === 400 || e.status === 409) && e.detail) {
      msg = `Power Platform API: ${e.detail}`;
    }
    return new DirectActionError(kind, msg, e.detail);
  }
  return new DirectActionError('failed', errorText(e));
}

export class DirectActionService {
  private clients: { [resource: string]: Promise<AadHttpClient> } = {};

  constructor(
    private factory: AadHttpClientFactory,
    private ppApi: PowerPlatformApiService,
    /** CoE-Environment (für das Nachziehen des Inventars und die Solution-Prüfung) */
    private coe: DataverseService | undefined,
    /** E-Mail-Adresse → Entra-Objekt-ID (Microsoft Graph) */
    private resolveUser?: (mail: string) => Promise<string | undefined>
  ) {}

  private client(resource: string): Promise<AadHttpClient> {
    if (!this.clients[resource]) {
      const p = this.factory.getClient(resource);
      this.clients[resource] = p;
      p.catch(() => { delete this.clients[resource]; });
    }
    return this.clients[resource];
  }

  /** Aufruf einer Admin-Schnittstelle mit Wiederholung bei Drosselung. */
  private async call(resource: string, method: 'POST' | 'DELETE', url: string, body?: unknown): Promise<void> {
    for (let attempt = 0; ; attempt++) {
      let res: HttpClientResponse;
      try {
        const client = await this.client(resource);
        res = await client.fetch(url, AadHttpClient.configurations.v1, {
          method,
          headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' },
          body: body === undefined ? undefined : JSON.stringify(body)
        });
      } catch (e) {
        const msg = errorText(e);
        if (isConsentError(msg)) {
          throw new DirectActionError('consent', CONSENT[resource] || msg, msg);
        }
        throw new DirectActionError('failed', `Die Verwaltungsschnittstelle ist nicht erreichbar: ${msg}`, msg);
      }
      if (res.ok) {
        return;
      }
      const text = await res.text().catch(() => '');
      if ((res.status === 429 || res.status === 503) && attempt < 3) {
        const after = Number(res.headers && res.headers.get('Retry-After')) || 0;
        await wait(Math.min(after > 0 ? after * 1000 : 2000 * (attempt + 1), 15000));
        continue;
      }
      const msg = messageOf(text) || `HTTP ${res.status}`;
      if (res.status === 401 || res.status === 403) {
        throw new DirectActionError(isConsentError(msg) ? 'consent' : 'forbidden', isConsentError(msg) ? (CONSENT[resource] || msg) : `${FORBIDDEN} (${msg})`, msg);
      }
      if (res.status === 404) {
        throw new DirectActionError('notFound', `Nicht gefunden – vermutlich bereits gelöscht oder in einem anderen Environment (${msg}).`, msg);
      }
      throw new DirectActionError('failed', `${res.status}: ${msg}`, msg);
    }
  }

  /** Zielperson/-gruppe als Entra-Objekt-ID. */
  private async target(req: IActionRequest): Promise<string> {
    const t = (req.targetUser || '').trim();
    if (GUID.test(t)) {
      return t;
    }
    if (!t) {
      throw new DirectActionError('rejected', 'Keine Person bzw. Gruppe angegeben.');
    }
    const id = this.resolveUser ? await this.resolveUser(t).catch(() => undefined) : undefined;
    if (!id) {
      throw new DirectActionError('rejected', `Die Person „${t}“ wurde nicht gefunden. Gruppen bitte aus den Vorschlägen wählen.`);
    }
    return id;
  }

  /** Löschen nur, wenn die Ressource laut CoE NICHT Teil einer Solution ist (unbekannt = ablehnen). */
  private async ensureNotInSolution(logicalName: string, field: string, coeId: string, label: string): Promise<void> {
    let value: unknown;
    try {
      value = this.coe ? (await this.coe.getRecord(logicalName, coeId))[field] : undefined;
    } catch {
      value = undefined;
    }
    if (value !== false) {
      throw new DirectActionError('rejected', `${label} ist Teil einer Solution (oder das ist im CoE nicht erfasst) – Löschen nur über die Solution (ALM).`);
    }
  }

  /** CoE-Inventar nachziehen – Fehler dabei sind nur ein Hinweis (der nächste CoE-Sync korrigiert es ohnehin). */
  private async coeUpdate(entitySet: string, id: string, values: { [field: string]: unknown }): Promise<string> {
    if (!this.coe || !id) {
      return '';
    }
    try {
      await this.coe.update(entitySet, id, values);
      return '';
    } catch (e) {
      return ` (CoE-Inventar nicht aktualisiert: ${errorText(e)} – der nächste CoE-Sync zieht es nach.)`;
    }
  }

  public async run(req: IActionRequest): Promise<IDirectResult> {
    const env = encodeURIComponent(req.environmentId);
    const id = encodeURIComponent(req.resourceId);
    const now = new Date().toISOString();
    const app = `${POWERAPPS_ADMIN}/${env}/apps/${id}`;
    const flow = `${FLOW_ADMIN}/${env}/flows/${id}`;
    const v2016 = 'api-version=2016-11-01';
    const isApp = req.resourceType === 'app';
    const isFlow = req.resourceType === 'flow';
    const ptype = req.targetIsGroup ? 'group' : 'user';
    if (!req.environmentId || (!req.resourceId && req.resourceType !== 'environment')) {
      throw new DirectActionError('rejected', 'Environment oder ID der Ressource ist im CoE nicht erfasst.');
    }

    switch (req.action) {
      case 'quarantine':
      case 'unquarantine': {
        if (!isApp) {
          break;
        }
        const on = req.action === 'quarantine';
        await this.call(POWERAPPS_RESOURCE, 'POST', `${app}/setPowerAppQuarantineState?api-version=2021-02-01`, { quarantineStatus: on ? 'Quarantined' : 'Unquarantined' });
        const note = await this.coeUpdate('admin_apps', req.coeRecordId, on ? { admin_quarantineapp: true, admin_quarantineappdate: now } : { admin_quarantineapp: false });
        if (!note) {
          await this.coeUpdate('admin_apps', req.coeRecordId, { admin_appisquarantined: on });
        }
        return { text: on ? 'Die App ist jetzt in Quarantäne.' : 'Die App ist nicht mehr in Quarantäne.', note };
      }

      case 'delete': {
        if (isApp) {
          await this.ensureNotInSolution('admin_app', 'admin_issolutionapp', req.coeRecordId, 'Die App');
          await this.call(POWERAPPS_RESOURCE, 'DELETE', `${app}?${v2016}`);
          return { text: 'Die App wurde im Tenant gelöscht.', note: await this.coeUpdate('admin_apps', req.coeRecordId, { admin_appdeleted: true, admin_appdeletedon: now }) };
        }
        if (isFlow) {
          await this.ensureNotInSolution('admin_flow', 'admin_issolutionflow', req.coeRecordId, 'Der Flow');
          await this.call(FLOW_RESOURCE, 'DELETE', `${flow}?${v2016}`);
          return { text: 'Der Flow wurde im Tenant gelöscht.', note: await this.coeUpdate('admin_flows', req.coeRecordId, { admin_flowdeleted: true, admin_flowdeletedon: now }) };
        }
        break;
      }

      case 'setowner': {
        const who = await this.target(req);
        if (isApp) {
          await this.call(POWERAPPS_RESOURCE, 'POST', `${app}/modifyAppOwner?${v2016}`, { roleForOldAppOwner: 'CanEdit', newAppOwner: who });
          return {
            text: `${req.targetName || who} ist jetzt Besitzer der App; der bisherige Besitzer ist Mitbesitzer.`,
            note: await this.coeUpdate('admin_apps', req.coeRecordId, { 'admin_AppOwner@odata.bind': `/admin_makers(${who})` })
          };
        }
        if (isFlow) {
          await this.call(FLOW_RESOURCE, 'POST', `${flow}/modifyOwners?${v2016}`, { put: [{ properties: { principal: { id: who, type: ptype }, roleName: 'CanEdit' } }] });
          return {
            text: `${req.targetName || who} ist jetzt Mitbesitzer des Flows. Den Ersteller eines Flows außerhalb einer Solution kann Power Automate nicht ändern.`,
            note: await this.coeUpdate('admin_flows', req.coeRecordId, { 'admin_DerivedOwner@odata.bind': `/admin_makers(${who})` })
          };
        }
        break;
      }

      case 'addviewer':
      case 'addeditor': {
        const who = await this.target(req);
        const role = req.action === 'addeditor' ? 'CanEdit' : 'CanView';
        if (isApp) {
          await this.call(POWERAPPS_RESOURCE, 'POST', `${app}/modifyPermissions?${v2016}`, {
            put: [{ properties: { roleName: role, principal: { id: who, type: ptype, tenantId: null } } }]
          });
          return { text: `Freigabe für ${req.targetName || who} hinzugefügt (${role === 'CanEdit' ? 'Mitbesitzer' : 'Benutzer'}).`, note: '' };
        }
        if (isFlow) {
          await this.call(FLOW_RESOURCE, 'POST', `${flow}/modifyOwners?${v2016}`, { put: [{ properties: { principal: { id: who, type: ptype }, roleName: 'CanEdit' } }] });
          return { text: `${req.targetName || who} ist jetzt Mitbesitzer des Flows.`, note: '' };
        }
        break;
      }

      case 'removeuser': {
        const who = await this.target(req);
        if (isApp) {
          await this.call(POWERAPPS_RESOURCE, 'POST', `${app}/modifyPermissions?${v2016}`, { delete: [{ id: who }] });
          return { text: `Freigabe für ${req.targetName || who} entfernt.`, note: '' };
        }
        if (isFlow) {
          await this.call(FLOW_RESOURCE, 'POST', `${flow}/modifyOwners?${v2016}`, { delete: [{ id: who }] });
          return { text: `Freigabe für ${req.targetName || who} entfernt.`, note: '' };
        }
        break;
      }

      case 'flowoff':
      case 'flowon': {
        if (!isFlow) {
          break;
        }
        const on = req.action === 'flowon';
        await this.call(FLOW_RESOURCE, 'POST', `${flow}/${on ? 'start' : 'stop'}?${v2016}`);
        return { text: on ? 'Der Flow ist eingeschaltet.' : 'Der Flow ist ausgeschaltet.', note: await this.coeUpdate('admin_flows', req.coeRecordId, { admin_flowstate: on ? 'Started' : 'Stopped' }) };
      }

      case 'envmanaged':
      case 'envdisable':
      case 'envenable': {
        if (req.resourceType !== 'environment') {
          break;
        }
        const path = `/environmentmanagement/environments/${env}`;
        try {
          if (req.action === 'envmanaged') {
            await this.ppApi.post(`${path}/governanceSetting/enableManaged`, {});
          } else if (req.action === 'envdisable') {
            await this.ppApi.post(`${path}/Disable`, { reason: 'UserRequested' });
          } else {
            await this.ppApi.post(`${path}/Enable`, {});
          }
        } catch (e) {
          throw fromPpApi(e);
        }
        const what = req.action === 'envmanaged' ? 'wird zum Managed Environment' : req.action === 'envdisable' ? 'wird deaktiviert' : 'wird wieder aktiviert';
        return { text: `Das Environment ${what} (Vorgang läuft bei Microsoft weiter, Dauer wenige Minuten).`, note: '' };
      }

      default:
        throw new DirectActionError('rejected', 'Diese Aktion führt der Flow „CoE Cockpit | Aktion ausführen“ aus.');
    }
    throw new DirectActionError('rejected', 'Diese Aktion gibt es für diese Art von Ressource nicht.');
  }
}
