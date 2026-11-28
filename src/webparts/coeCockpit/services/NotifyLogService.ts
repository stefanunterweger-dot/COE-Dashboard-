import { SPHttpClient } from '@microsoft/sp-http';

/**
 * Gemeinsames Versandprotokoll in einer SharePoint-Liste auf der Website des Cockpits.
 * Alle Admins sehen damit, welches Environment wann von wem verständigt wurde – und niemand
 * schreibt dieselben Maker versehentlich doppelt an. Zugriff mit den Rechten des angemeldeten
 * Benutzers auf die Website (keine zusätzliche API-Berechtigung).
 */
export const NOTIFY_LIST_TITLE = 'CoE Cockpit Benachrichtigungen';

export type NotifyMethod = 'graph' | 'outlook';

export interface INotifyEntry {
  envId: string;
  envName: string;
  /** Zeitpunkt (ISO). */
  at: string;
  /** Anzeigename der Person, die verständigt hat. */
  by: string;
  byEmail: string;
  method: NotifyMethod;
  recipients: string;
  subject: string;
}

/** Jüngster Eintrag je Environment (Schlüssel: Environment-ID in Kleinbuchstaben). */
export interface INotifyLogMap { [envId: string]: INotifyEntry }

const METHOD_LABEL: { [m: string]: string } = {
  graph: 'Automatisch gesendet (Microsoft Graph)',
  outlook: 'In Outlook gesendet (bestätigt)'
};

const dtFmt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });

/** „am 26.11.2026, 14:03 von Max Muster (automatisch)“ */
export function describeEntry(e: INotifyEntry | undefined): string {
  if (!e) {
    return '';
  }
  const d = new Date(e.at);
  const when = isNaN(d.getTime()) ? '' : dtFmt.format(d);
  return `am ${when} von ${e.by || e.byEmail || 'unbekannt'} (${e.method === 'graph' ? 'automatisch' : 'über Outlook'})`;
}

const JSON_HEADERS = {
  'Accept': 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

/** Spalten der Liste (interner Name = Anzeigename). */
const FIELDS: string[] = [
  '<Field Type="Text" DisplayName="EnvironmentId" Name="EnvironmentId" StaticName="EnvironmentId" MaxLength="80" />',
  '<Field Type="DateTime" DisplayName="SentAt" Name="SentAt" StaticName="SentAt" Format="DateTime" Indexed="TRUE" />',
  '<Field Type="Text" DisplayName="SentBy" Name="SentBy" StaticName="SentBy" MaxLength="255" />',
  '<Field Type="Text" DisplayName="SentByEmail" Name="SentByEmail" StaticName="SentByEmail" MaxLength="255" />',
  '<Field Type="Text" DisplayName="Method" Name="Method" StaticName="Method" MaxLength="80" />',
  '<Field Type="Note" DisplayName="Recipients" Name="Recipients" StaticName="Recipients" NumLines="4" RichText="FALSE" />',
  '<Field Type="Text" DisplayName="Subject" Name="Subject" StaticName="Subject" MaxLength="255" />'
];

interface IListItem {
  Title?: string; EnvironmentId?: string; SentAt?: string; SentBy?: string; SentByEmail?: string; Method?: string; Recipients?: string; Subject?: string;
}

export class NotifyLogError extends Error {
  constructor(message: string, public status: number) {
    super(message);
    Object.setPrototypeOf(this, NotifyLogError.prototype);
  }
}

export class NotifyLogService {
  private _ensured: Promise<void> | undefined;

  constructor(private _client: SPHttpClient, private _webUrl: string) {}

  private get _list(): string {
    return `${this._webUrl.replace(/\/$/, '')}/_api/web/lists/GetByTitle('${NOTIFY_LIST_TITLE.replace(/'/g, "''")}')`;
  }

  /**
   * Alle Einträge lesen; exists = false, wenn die Liste (noch) nicht angelegt ist.
   * Ohne $orderby (Standardreihenfolge nach der ID ist immer indiziert) und seitenweise – so bricht das Lesen
   * auch ab 5.000 Einträgen nicht am Schwellenwert für Listenansichten ab, und kein älterer Eintrag fehlt.
   */
  public async load(): Promise<{ map: INotifyLogMap; exists: boolean }> {
    let url = `${this._list}/items?$select=Title,EnvironmentId,SentAt,SentBy,SentByEmail,Method,Recipients,Subject&$top=5000`;
    const items: IListItem[] = [];
    for (let page = 0; url && page < 200; page++) {
      const res = await this._client.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
      if (res.status === 404 && page === 0) {
        return { map: {}, exists: false };
      }
      if (!res.ok) {
        throw new NotifyLogError(await this._message(res), res.status);
      }
      const body = await res.json() as { value: IListItem[]; 'odata.nextLink'?: string };
      items.push(...(body.value || []));
      url = body['odata.nextLink'] || '';
    }
    const map: INotifyLogMap = {};
    items.forEach(it => {
      const id = String(it.EnvironmentId || '').toLowerCase();
      if (!id) {
        return;
      }
      const entry: INotifyEntry = {
        envId: id, envName: it.Title || '', at: it.SentAt || '', by: it.SentBy || '', byEmail: it.SentByEmail || '',
        method: /graph|automatisch/i.test(it.Method || '') ? 'graph' : 'outlook', recipients: it.Recipients || '', subject: it.Subject || ''
      };
      if (!map[id] || entry.at > map[id].at) {
        map[id] = entry;
      }
    });
    return { map, exists: true };
  }

  /** Eintrag anlegen (legt die Liste beim ersten Mal an). */
  public async add(e: INotifyEntry): Promise<void> {
    await this._ensureList();
    const res = await this._client.post(`${this._list}/items`, SPHttpClient.configurations.v1, {
      headers: JSON_HEADERS,
      body: JSON.stringify({
        Title: (e.envName || e.envId).slice(0, 255),
        EnvironmentId: e.envId,
        SentAt: e.at,
        SentBy: e.by.slice(0, 255),
        SentByEmail: e.byEmail.slice(0, 255),
        Method: METHOD_LABEL[e.method] || e.method,
        Recipients: e.recipients,
        Subject: e.subject.slice(0, 255)
      })
    });
    if (!res.ok) {
      throw new NotifyLogError(await this._message(res), res.status);
    }
  }

  private _ensureList(): Promise<void> {
    if (!this._ensured) {
      const p = this._createIfMissing();
      this._ensured = p;
      p.catch(() => { this._ensured = undefined; });
    }
    return this._ensured;
  }

  private async _createIfMissing(): Promise<void> {
    const check = await this._client.get(`${this._list}?$select=Id`, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    if (check.ok) {
      return;
    }
    if (check.status !== 404) {
      throw new NotifyLogError(await this._message(check), check.status);
    }
    const web = `${this._webUrl.replace(/\/$/, '')}/_api/web`;
    const created = await this._client.post(`${web}/lists`, SPHttpClient.configurations.v1, {
      headers: JSON_HEADERS,
      body: JSON.stringify({
        Title: NOTIFY_LIST_TITLE,
        Description: 'Protokoll der Benachrichtigungen aus dem CoE Cockpit (welches Environment wann von wem verständigt wurde).',
        BaseTemplate: 100,
        AllowContentTypes: false,
        ContentTypesEnabled: false
      })
    });
    if (!created.ok) {
      const msg = await this._message(created);
      throw new NotifyLogError(created.status === 403
        ? `Die Liste „${NOTIFY_LIST_TITLE}“ konnte nicht angelegt werden – dafür braucht es Besitzer-Rechte auf dieser Website. ${msg}`
        : msg, created.status);
    }
    for (const xml of FIELDS) {
      const f = await this._client.post(`${this._list}/fields/CreateFieldAsXml`, SPHttpClient.configurations.v1, {
        headers: JSON_HEADERS,
        // 8 = interner Name wie angegeben, 16 = in die Standardansicht aufnehmen
        body: JSON.stringify({ parameters: { SchemaXml: xml, Options: 24 } })
      });
      if (!f.ok) {
        throw new NotifyLogError(await this._message(f), f.status);
      }
    }
  }

  private async _message(res: { status: number; text(): Promise<string> }): Promise<string> {
    const text = await res.text().catch(() => '');
    try {
      const b = JSON.parse(text) as { 'odata.error'?: { message?: { value?: string } }; error?: { message?: string | { value?: string } } };
      const m = b['odata.error'] && b['odata.error'].message ? b['odata.error'].message.value
        : b.error ? (typeof b.error.message === 'string' ? b.error.message : b.error.message && b.error.message.value) : '';
      return m || `HTTP ${res.status}`;
    } catch {
      return text || `HTTP ${res.status}`;
    }
  }
}
