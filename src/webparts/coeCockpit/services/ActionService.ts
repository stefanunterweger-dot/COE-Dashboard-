import { SPHttpClient } from '@microsoft/sp-http';
import { ACTION_LIST_TITLE, ActionCode, ActionStatus, IActionItem, IActionRequest, ResourceType } from '../models/actions';

/**
 * SharePoint-Seite der Cockpit-Aktionen: Admin-Gruppe, Aktionsliste mit eigenen Berechtigungen,
 * Aufträge anlegen und Status lesen. Alles mit den SharePoint-Rechten des angemeldeten Benutzers.
 * Die Liste ist zugleich das Aktionsprotokoll – auch für Aktionen, die das Webpart direkt ausführt.
 */
const JSON_HEADERS = {
  'Accept': 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

/** Berechtigungsstufen (IDs der Standard-Rollendefinitionen in SharePoint). */
const ROLE_CONTRIBUTE = 1073741827;
const ROLE_FULL_CONTROL = 1073741829;

const FIELDS: string[] = [
  '<Field Type="Text" DisplayName="Action" Name="Action" StaticName="Action" MaxLength="40" Indexed="TRUE" />',
  '<Field Type="Text" DisplayName="ResourceType" Name="ResourceType" StaticName="ResourceType" MaxLength="20" />',
  '<Field Type="Text" DisplayName="ResourceId" Name="ResourceId" StaticName="ResourceId" MaxLength="80" />',
  '<Field Type="Text" DisplayName="CoeRecordId" Name="CoeRecordId" StaticName="CoeRecordId" MaxLength="80" />',
  '<Field Type="Text" DisplayName="ResourceName" Name="ResourceName" StaticName="ResourceName" MaxLength="255" />',
  '<Field Type="Text" DisplayName="EnvironmentId" Name="EnvironmentId" StaticName="EnvironmentId" MaxLength="120" />',
  '<Field Type="Text" DisplayName="TargetUser" Name="TargetUser" StaticName="TargetUser" MaxLength="255" />',
  '<Field Type="Text" DisplayName="TargetName" Name="TargetName" StaticName="TargetName" MaxLength="255" />',
  '<Field Type="Boolean" DisplayName="TargetIsGroup" Name="TargetIsGroup" StaticName="TargetIsGroup"><Default>0</Default></Field>',
  '<Field Type="Note" DisplayName="Reason" Name="Reason" StaticName="Reason" NumLines="4" RichText="FALSE" />',
  '<Field Type="Text" DisplayName="Status" Name="Status" StaticName="Status" MaxLength="40" Indexed="TRUE" />',
  '<Field Type="Note" DisplayName="Result" Name="Result" StaticName="Result" NumLines="4" RichText="FALSE" />',
  '<Field Type="DateTime" DisplayName="CompletedAt" Name="CompletedAt" StaticName="CompletedAt" Format="DateTime" />',
  '<Field Type="Text" DisplayName="RequestedBy" Name="RequestedBy" StaticName="RequestedBy" MaxLength="255" />'
];
const SELECT = 'Id,Title,Action,ResourceType,ResourceId,CoeRecordId,ResourceName,EnvironmentId,TargetUser,TargetName,TargetIsGroup,Reason,Status,Result,CompletedAt,RequestedBy,Created';

export interface ISiteGroup { id: number; title: string }

export interface IActionSetupState {
  /** Gruppe vorhanden? */
  group?: ISiteGroup;
  /** Liste vorhanden? */
  listExists: boolean;
  /** Liste hat eigene Berechtigungen (nicht von der Website geerbt)? */
  listUnique?: boolean;
  /** Ist der angemeldete Benutzer Mitglied der Gruppe? */
  isMember: boolean;
}

export class ActionError extends Error {
  constructor(message: string, public status: number) {
    super(message);
    Object.setPrototypeOf(this, ActionError.prototype);
  }
}

interface IListItem {
  Id: number; Title?: string; Action?: string; ResourceType?: string; ResourceId?: string; CoeRecordId?: string; ResourceName?: string;
  EnvironmentId?: string; TargetUser?: string; TargetName?: string; TargetIsGroup?: boolean; Reason?: string; Status?: string;
  Result?: string; CompletedAt?: string; RequestedBy?: string; Created?: string;
}

function toItem(it: IListItem): IActionItem {
  return {
    id: it.Id,
    action: (it.Action || '') as ActionCode,
    resourceType: (it.ResourceType || 'app') as ResourceType,
    resourceId: it.ResourceId || '',
    coeRecordId: it.CoeRecordId || '',
    resourceName: it.ResourceName || it.Title || '',
    environmentId: it.EnvironmentId || '',
    targetUser: it.TargetUser || '',
    targetName: it.TargetName || '',
    targetIsGroup: !!it.TargetIsGroup,
    reason: it.Reason || '',
    status: (it.Status || 'Angefordert') as ActionStatus,
    result: it.Result || '',
    created: it.Created || '',
    completedAt: it.CompletedAt || '',
    requestedBy: it.RequestedBy || ''
  };
}

export class ActionService {
  constructor(private _client: SPHttpClient, private _webUrl: string) {}

  private get _web(): string {
    return `${this._webUrl.replace(/\/$/, '')}/_api/web`;
  }

  private get _list(): string {
    return `${this._web}/lists/GetByTitle('${ACTION_LIST_TITLE.replace(/'/g, "''")}')`;
  }

  private async _get<T>(url: string): Promise<T | undefined> {
    const res = await this._client.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    if (res.status === 404) {
      return undefined;
    }
    if (!res.ok) {
      throw new ActionError(await this._message(res), res.status);
    }
    return res.json() as Promise<T>;
  }

  private async _post<T>(url: string, body?: unknown): Promise<T> {
    const res = await this._client.post(url, SPHttpClient.configurations.v1, {
      headers: JSON_HEADERS,
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (!res.ok) {
      throw new ActionError(await this._message(res), res.status);
    }
    const text = await res.text();
    return (text ? JSON.parse(text) : {}) as T;
  }

  /** SharePoint-Gruppen der Website (für die Auswahl in den Eigenschaften). */
  public async listGroups(): Promise<ISiteGroup[]> {
    const res = await this._get<{ value: { Id: number; Title: string }[] }>(`${this._web}/sitegroups?$select=Id,Title&$orderby=Title`);
    return ((res && res.value) || []).map(g => ({ id: g.Id, title: g.Title }));
  }

  public async getGroup(title: string): Promise<ISiteGroup | undefined> {
    const res = await this._get<{ Id: number; Title: string }>(`${this._web}/sitegroups/getbyname('${encodeURIComponent(title.replace(/'/g, "''"))}')?$select=Id,Title`);
    return res ? { id: res.Id, title: res.Title } : undefined;
  }

  /**
   * Gruppe anlegen: Besitzer = Besitzer-Gruppe der Website, Mitgliedschaft für alle sichtbar
   * (der Flow prüft sie), die anlegende Person wird erstes Mitglied.
   */
  public async createGroup(title: string): Promise<ISiteGroup> {
    const existing = await this.getGroup(title);
    if (existing) {
      return existing;
    }
    const g = await this._post<{ Id: number; Title: string }>(`${this._web}/sitegroups`, {
      Title: title,
      Description: 'Power-Platform-Admins, die im CoE Cockpit Aktionen auslösen dürfen (Quarantäne, Löschen, Besitzer und Freigaben, Flows ein/aus, Environments, Solutions).',
      OnlyAllowMembersViewMembership: false,
      AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false
    });
    // Besitzer der Gruppe = Besitzer der Website (optional – Fehler hier sind nicht schlimm)
    try {
      const owners = await this._get<{ Id: number }>(`${this._web}/associatedownergroup?$select=Id`);
      if (owners) {
        await this._post(`${this._web}/sitegroups(${g.Id})/SetUserAsOwner(${owners.Id})`);
      }
    } catch {
      // bleibt dann bei der anlegenden Person
    }
    // anlegende Person als erstes Mitglied
    try {
      const me = await this._get<{ LoginName: string }>(`${this._web}/currentuser?$select=LoginName`);
      if (me) {
        await this._post(`${this._web}/sitegroups(${g.Id})/users`, { LoginName: me.LoginName });
      }
    } catch {
      // Mitglied kann auch später hinzugefügt werden
    }
    return { id: g.Id, title: g.Title };
  }

  /** Ist der angemeldete Benutzer Mitglied dieser Gruppe? */
  public async isMember(groupTitle: string): Promise<boolean> {
    const res = await this._get<{ value: { Title: string }[] }>(`${this._web}/currentuser/groups?$select=Title`);
    return ((res && res.value) || []).some(g => g.Title.toLowerCase() === groupTitle.toLowerCase());
  }

  /** Zustand der Einrichtung (Gruppe, Liste, Berechtigungen, Mitgliedschaft). */
  public async getSetupState(groupTitle: string): Promise<IActionSetupState> {
    const [group, list, isMember] = await Promise.all([
      this.getGroup(groupTitle).catch(() => undefined),
      this._get<{ HasUniqueRoleAssignments: boolean }>(`${this._list}?$select=HasUniqueRoleAssignments`).catch(() => undefined),
      this.isMember(groupTitle).catch(() => false)
    ]);
    return { group, listExists: !!list, listUnique: list ? list.HasUniqueRoleAssignments : undefined, isMember };
  }

  /**
   * Aktionsliste anlegen (falls nötig) und absichern: eigene Berechtigungen, die Gruppe darf
   * bearbeiten, die Besitzer der Website haben Vollzugriff, sonst niemand.
   */
  public async setupList(groupTitle: string): Promise<void> {
    const group = await this.getGroup(groupTitle);
    if (!group) {
      throw new ActionError(`Die Gruppe „${groupTitle}“ gibt es auf dieser Website nicht – zuerst anlegen.`, 404);
    }
    const exists = await this._get<{ Id: string }>(`${this._list}?$select=Id`);
    if (!exists) {
      await this._post(`${this._web}/lists`, {
        Title: ACTION_LIST_TITLE,
        Description: 'Aktionsprotokoll des CoE Cockpits (Quarantäne, Löschen, Besitzer, Freigaben, Flows ein/aus, Environments, Solutions). Das Cockpit führt die meisten Aktionen direkt aus; „Solution löschen“ übernimmt der Flow im CoE-Environment.',
        BaseTemplate: 100,
        AllowContentTypes: false,
        ContentTypesEnabled: false
      });
      for (const xml of FIELDS) {
        // 8 = interner Name wie angegeben, 16 = in die Standardansicht
        await this._post(`${this._list}/fields/CreateFieldAsXml`, { parameters: { SchemaXml: xml, Options: 24 } });
      }
      // Versionsverlauf: Änderungen an Protokolleinträgen bleiben nachvollziehbar (Fehler hier stoppt die Einrichtung nicht)
      await this._client.post(this._list, SPHttpClient.configurations.v1, {
        headers: { ...JSON_HEADERS, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
        body: JSON.stringify({ EnableVersioning: true })
      }).catch(() => undefined);
    }
    // Berechtigungen: Vererbung aufheben, nur Gruppe (Bearbeiten) und Websitebesitzer (Vollzugriff)
    await this._post(`${this._list}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
    await this._post(`${this._list}/roleassignments/addroleassignment(principalid=${group.id},roledefid=${ROLE_CONTRIBUTE})`);
    const owners = await this._get<{ Id: number }>(`${this._web}/associatedownergroup?$select=Id`);
    if (owners) {
      await this._post(`${this._list}/roleassignments/addroleassignment(principalid=${owners.Id},roledefid=${ROLE_FULL_CONTROL})`);
    }
  }

  /**
   * Auftrag anlegen; liefert die ID des Listenelements.
   * Status „Angefordert“ = der Flow führt ihn aus. Direkt im Webpart ausgeführte Aktionen werden mit
   * „In Arbeit“ angelegt (der Flow übernimmt nur „Angefordert“) und danach mit complete() abgeschlossen.
   */
  public async add(req: IActionRequest, requestedBy: string, status: ActionStatus = 'Angefordert', result: string = ''): Promise<number> {
    const item = await this._post<{ Id: number }>(`${this._list}/items`, {
      Title: `${req.action}: ${req.resourceName}`.slice(0, 255),
      Action: req.action,
      ResourceType: req.resourceType,
      ResourceId: req.resourceId,
      CoeRecordId: req.coeRecordId,
      ResourceName: req.resourceName.slice(0, 255),
      EnvironmentId: req.environmentId,
      TargetUser: req.targetUser || '',
      TargetName: (req.targetName || '').slice(0, 255),
      TargetIsGroup: !!req.targetIsGroup,
      Reason: req.reason,
      Status: status,
      Result: result,
      RequestedBy: requestedBy.slice(0, 255)
    });
    return item.Id;
  }

  /** Ergebnis einer direkt ausgeführten Aktion zurückschreiben (zweiter Versuch bei Fehlern). */
  public async complete(id: number, status: ActionStatus, result: string): Promise<void> {
    const body = { Status: status, Result: result, CompletedAt: new Date().toISOString() };
    const send = async (): Promise<void> => {
      const res = await this._client.post(`${this._list}/items(${id})`, SPHttpClient.configurations.v1, {
        headers: { ...JSON_HEADERS, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
        body: JSON.stringify(body)
      });
      if (!res.ok) {
        throw new ActionError(await this._message(res), res.status);
      }
    };
    try {
      await send();
    } catch {
      await new Promise(resolve => setTimeout(resolve, 1500));
      await send();
    }
  }

  public async get(id: number): Promise<IActionItem | undefined> {
    const it = await this._get<IListItem>(`${this._list}/items(${id})?$select=${SELECT}`);
    return it ? toItem(it) : undefined;
  }

  /**
   * Die letzten Aufträge (neueste zuerst). Sortiert nach der ID (immer indiziert) – eine Sortierung nach
   * „Created“ bricht ab 5.000 Einträgen am Schwellenwert für Listenansichten ab.
   */
  public async list(top: number = 500): Promise<IActionItem[] | undefined> {
    const res = await this._get<{ value: IListItem[] }>(`${this._list}/items?$select=${SELECT}&$orderby=ID desc&$top=${top}`);
    if (!res) {
      return undefined;
    }
    const items = res.value || [];
    // Offene Aufträge zusätzlich gezielt lesen – nach einem großen Sammellauf lägen sie sonst außerhalb der neuesten Einträge
    // und die Sperre gegen Doppelaufträge griffe nicht (Status ist indiziert; Fehler hier = nur die neuesten Einträge).
    const open = await this._get<{ value: IListItem[] }>(`${this._list}/items?$select=${SELECT}&$filter=${encodeURIComponent("(Status eq 'Angefordert') or (Status eq 'In Arbeit')")}&$orderby=ID desc&$top=500`)
      .catch(() => undefined);
    const seen: { [id: number]: boolean } = {};
    items.forEach(i => { seen[i.Id] = true; });
    ((open && open.value) || []).forEach(i => { if (!seen[i.Id]) { seen[i.Id] = true; items.push(i); } });
    return items.map(toItem);
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
