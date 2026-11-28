/* Simulierte SharePoint-REST-API: Versandprotokoll, Aktionsliste, Gruppen – und ein simulierter Flow. */
import { IMockResponse } from './odata';
import { DATABASES } from './data';

interface IList { exists: boolean; fields: string[]; items: Record<string, unknown>[]; unique: boolean; nextId: number }
interface IGroup { Id: number; Title: string; members: string[] }
interface IStore { lists: { [title: string]: IList }; groups: IGroup[]; nextGroupId: number }
interface ISpWindow { __spStore?: IList; __spAll?: IStore; __spReset?: () => void; __spFlowDelayMs?: number; __spFlowOff?: boolean; __spNotMember?: boolean }

const NOTIFY = 'CoE Cockpit Benachrichtigungen';
const ACTIONS = 'CoE Cockpit Aktionen';
const ME = 'i:0#.f|membership|alex.admin@contoso.com';
const w = window as unknown as ISpWindow;

const emptyList = (): IList => ({ exists: false, fields: [], items: [], unique: false, nextId: 1 });
function all(): IStore {
  if (!w.__spAll) {
    w.__spAll = {
      lists: { [NOTIFY]: emptyList(), [ACTIONS]: emptyList() },
      groups: [{ Id: 3, Title: 'Power Platform Besitzer', members: [ME] }, { Id: 4, Title: 'Power Platform Mitglieder', members: [ME] }, { Id: 5, Title: 'Power Platform Besucher', members: [] }],
      nextGroupId: 10
    };
  }
  // Kurzzugriff für Tests: Versandprotokoll
  w.__spStore = w.__spAll.lists[NOTIFY];
  return w.__spAll;
}
w.__spReset = () => { all().lists[NOTIFY] = emptyList(); w.__spStore = all().lists[NOTIFY]; };

/** Vorbelegung: ein Environment, das ein anderer Admin schon verständigt hat. */
export function seedNotifyEntry(item: Record<string, unknown>): void {
  const l = all().lists[NOTIFY];
  l.exists = true;
  l.fields = ['EnvironmentId', 'SentAt', 'SentBy', 'SentByEmail', 'Method', 'Recipients', 'Subject'];
  l.items.push(item);
}

/** Vorbelegung der Aktionen: Gruppe mit dem angemeldeten Admin, eingerichtete Liste. */
export function seedActions(): void {
  const s = all();
  if (!s.groups.some(g => g.Title === 'CoE Cockpit Admins')) {
    s.groups.push({ Id: 9, Title: 'CoE Cockpit Admins', members: [ME] });
  }
  const l = s.lists[ACTIONS];
  l.exists = true;
  l.unique = true;
  l.fields = ['Action', 'ResourceType', 'ResourceId', 'CoeRecordId', 'ResourceName', 'EnvironmentId', 'TargetUser', 'TargetName', 'TargetIsGroup', 'Reason', 'Status', 'Result', 'CompletedAt', 'RequestedBy'];
}

function respond(status: number, body?: unknown): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return { ok: status >= 200 && status < 300, status, json: () => Promise.resolve(text ? JSON.parse(text) : {}), text: () => Promise.resolve(text) };
}
const notFound = (what: string): IMockResponse => respond(404, { 'odata.error': { code: '-1, System.ArgumentException', message: { lang: 'de-DE', value: `${what} ist nicht vorhanden.` } } });

/** Simulierter Flow „CoE Cockpit | Aktion ausführen“ (ab 1.2.0.0 nur noch „Solution löschen“): In Arbeit → Erledigt, CoE-Daten anpassen. */
function simulateFlow(item: Record<string, unknown>): void {
  if (w.__spFlowOff) {
    return;
  }
  const delay = w.__spFlowDelayMs === undefined ? 1200 : w.__spFlowDelayMs;
  setTimeout(() => {
    item.Status = 'In Arbeit';
    setTimeout(() => {
      const coe = DATABASES[0].tables;
      // Wie der echte Flow: Solution im Environment über ResourceId (= Solution-ID im Environment, nicht die CoE-ID)
      const id = String(item.ResourceId || '').toLowerCase();
      const sol = coe.admin_solution.rows.filter(r => String(r.admin_solutionenvtguid || '').toLowerCase() === id)[0];
      if (item.Action === 'soldelete' && !sol) {
        item.Status = 'Fehler';
        item.Result = 'Solution im Environment nicht gefunden (404) – Vorschau – simulierter Flow.';
        item.CompletedAt = new Date().toISOString();
        return;
      }
      let result = 'Ausgeführt (Vorschau – simulierter Flow).';
      switch (item.Action) {
        case 'soldelete':
          if (sol && sol.admin_solutionismanaged === false) {
            sol.admin_solutiondeleted = true;
            result = 'Die Solution wurde gelöscht; ihre Komponenten bleiben im Environment (Vorschau – simulierter Flow).';
          } else {
            item.Status = 'Abgelehnt';
            item.Result = 'Verwaltete Solution – nicht gelöscht.';
            item.CompletedAt = new Date().toISOString();
            return;
          }
          break;
        default:
          item.Status = 'Abgelehnt';
          item.Result = `Die Aktion „${item.Action}“ führt das CoE Cockpit (ab Version 1.14) direkt aus, nicht dieser Flow. Bitte das Webpart aktualisieren.`;
          item.CompletedAt = new Date().toISOString();
          return;
      }
      item.Status = 'Erledigt';
      item.Result = result;
      item.CompletedAt = new Date().toISOString();
    }, delay);
  }, delay);
}

export function mockSharePoint(method: 'GET' | 'POST', url: string, body?: string, headers?: { [k: string]: string }): Promise<IMockResponse> {
  return new Promise(resolve => setTimeout(() => resolve(handle(method, url, body, headers || {})), 80));
}

function handle(method: 'GET' | 'POST', url: string, body: string | undefined, headers: { [k: string]: string }): IMockResponse {
  const s = all();
  const u = decodeURIComponent(url);
  const json = (): Record<string, unknown> => (body ? JSON.parse(body) as Record<string, unknown> : {});

  // ---- Gruppen und Benutzer ----
  let m = /\/_api\/web\/sitegroups\/getbyname\('(.+?)'\)/.exec(u);
  if (m) {
    const g = s.groups.filter(x => x.Title.toLowerCase() === m![1].replace(/''/g, "'").toLowerCase())[0];
    return g ? respond(200, { Id: g.Id, Title: g.Title }) : notFound(`Die Gruppe „${m[1]}“`);
  }
  if (method === 'GET' && /\/_api\/web\/sitegroups\?/.test(u)) {
    return respond(200, { value: s.groups.map(g => ({ Id: g.Id, Title: g.Title })) });
  }
  if (method === 'POST' && /\/_api\/web\/sitegroups$/.test(u)) {
    const b = json();
    const g = { Id: s.nextGroupId++, Title: String(b.Title), members: [] as string[] };
    s.groups.push(g);
    return respond(201, { Id: g.Id, Title: g.Title });
  }
  m = /\/_api\/web\/sitegroups\((\d+)\)\/users$/.exec(u);
  if (m && method === 'POST') {
    const g = s.groups.filter(x => x.Id === Number(m![1]))[0];
    if (g) { g.members.push(String(json().LoginName)); }
    return respond(201, {});
  }
  if (/\/_api\/web\/sitegroups\(\d+\)\/SetUserAsOwner/.test(u)) {
    return respond(200, {});
  }
  if (/\/_api\/web\/associatedownergroup/.test(u)) {
    return respond(200, { Id: 3 });
  }
  if (/\/_api\/web\/currentuser\/groups/.test(u)) {
    return respond(200, { value: w.__spNotMember ? [] : s.groups.filter(g => g.members.indexOf(ME) >= 0).map(g => ({ Title: g.Title })) });
  }
  if (/\/_api\/web\/currentuser/.test(u)) {
    return respond(200, { LoginName: ME, Id: 7 });
  }

  // ---- Listen ----
  if (method === 'POST' && /\/_api\/web\/lists$/.test(u)) {
    const title = String(json().Title);
    s.lists[title] = s.lists[title] || emptyList();
    s.lists[title].exists = true;
    w.__spStore = s.lists[NOTIFY];
    return respond(201, { Id: 'f0000000-0000-4000-8000-000000000001', Title: title });
  }
  m = /GetByTitle\('(.+?)'\)/.exec(u);
  if (!m || !s.lists[m[1]]) {
    return respond(400, { 'odata.error': { message: { value: `Unsupported ${u}` } } });
  }
  const list = s.lists[m[1]];
  if (!list.exists) {
    return notFound('Die Liste');
  }
  if (method === 'POST' && /\/fields\/CreateFieldAsXml$/.test(u)) {
    const xml = (json() as { parameters: { SchemaXml: string } }).parameters.SchemaXml;
    list.fields.push((/Name="([^"]+)"/.exec(xml) || [])[1] || '');
    return respond(201, {});
  }
  if (method === 'POST' && /\/breakroleinheritance/.test(u)) {
    list.unique = true;
    return respond(200, {});
  }
  if (method === 'POST' && /\/roleassignments\/addroleassignment/.test(u)) {
    return respond(200, {});
  }
  if (method === 'POST' && /\/items$/.test(u)) {
    const item = json();
    const unknown = Object.keys(item).filter(k => k !== 'Title' && list.fields.indexOf(k) < 0);
    if (unknown.length) {
      return respond(400, { 'odata.error': { message: { value: `Die Eigenschaft '${unknown[0]}' ist nicht vorhanden.` } } });
    }
    const stored = { ...item, Id: list.nextId++, Created: new Date().toISOString() };
    list.items.push(stored);
    // wie die Auslösebedingung des Flows: nur Aufträge mit Status „Angefordert“
    if (m[1] === ACTIONS && (!stored.Status || stored.Status === 'Angefordert')) {
      simulateFlow(stored);
    }
    return respond(201, stored);
  }
  const byId = /\/items\((\d+)\)/.exec(u);
  if (method === 'POST' && byId && /merge/i.test(headers['X-HTTP-Method'] || headers['x-http-method'] || '')) {
    const it = list.items.filter(x => x.Id === Number(byId[1]))[0];
    if (!it) {
      return notFound('Das Element');
    }
    Object.assign(it, json());
    return respond(204);
  }
  if (method === 'GET' && byId) {
    const it = list.items.filter(x => x.Id === Number(byId[1]))[0];
    return it ? respond(200, it) : notFound('Das Element');
  }
  if (method === 'GET' && /\/items\?/.test(u)) {
    // wie SharePoint: ab 5.000 Einträgen nur noch nach indizierten Spalten sortieren (ID immer) – sonst Schwellenwert-Fehler
    const orderby = (/\$orderby=([^&]+)/.exec(u) || [])[1] || '';
    if (orderby && !/^ID(%20|\s|\+)/.test(orderby) && list.items.length > 5000) {
      return respond(500, { 'odata.error': { message: { value: 'The attempted operation is prohibited because it exceeds the list view threshold.' } } });
    }
    const desc = /desc/i.test(decodeURIComponent(orderby));
    const top = Number((/\$top=(\d+)/.exec(u) || [])[1] || 100);
    const after = Number((/p_ID=(\d+)/.exec(u) || [])[1] || 0);
    const sorted = list.items.slice().sort((a, b) => desc ? Number(b.Id) - Number(a.Id) : Number(a.Id) - Number(b.Id))
      .filter(x => desc || Number(x.Id) > after);
    const pageItems = sorted.slice(0, top);
    const body: Record<string, unknown> = { value: pageItems };
    if (!desc && sorted.length > top) {
      const base = u.replace(/&\$skiptoken=Paged(%3D|=)TRUE(%26|&)p_ID=\d+/, '');
      body['odata.nextLink'] = `${base}&$skiptoken=Paged%3DTRUE%26p_ID=${pageItems[pageItems.length - 1].Id}`;
    }
    return respond(200, body);
  }
  if (method === 'GET') {
    return respond(200, { Id: 'f0000000-0000-4000-8000-000000000001', HasUniqueRoleAssignments: list.unique });
  }
  return respond(400, {});
}
