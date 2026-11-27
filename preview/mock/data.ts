/* Deterministische Beispieldaten im Schema des CoE Starter Kits (Core Components 4.50, „Power Platform Admin View“). */

export interface IAttr {
  name: string;          // logischer Name (bei Lookups ohne _..._value)
  type: string;          // Dataverse AttributeType
  label: string;
  target?: string;       // Lookup-Ziel
  options?: { [value: number]: string }; // Picklist
}

export interface ITableDef {
  logicalName: string;
  entitySet: string;
  displayName: string;
  primaryId: string;
  primaryName: string;
  attrs: IAttr[];
  rows: Record<string, unknown>[];
  navs?: { [prop: string]: string };
}

export interface IDatabase {
  host: string;
  tables: { [logicalName: string]: ITableDef };
}

// ---------- Zufall (deterministisch) ----------
let seed = 20261125;
function rnd(): number {
  seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
  let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
  t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
}
const pick = <T>(a: T[]): T => a[Math.floor(rnd() * a.length)];
const chance = (p: number): boolean => rnd() < p;
const int = (min: number, max: number): number => min + Math.floor(rnd() * (max - min + 1));
function guid(): string {
  const h = (n: number): string => Array.from({ length: n }, () => Math.floor(rnd() * 16).toString(16)).join('');
  return `${h(8)}-${h(4)}-4${h(3)}-${pick(['8', '9', 'a', 'b'])}${h(3)}-${h(12)}`;
}
const DAY = 86400000;
const NOW = Date.now();
function daysAgo(d: number): string {
  return new Date(NOW - d * DAY - int(0, 86000) * 1000).toISOString().replace(/\.\d{3}Z$/, 'Z');
}
function recentish(maxDays: number): number {
  const r = rnd();
  if (r < 0.12) { return int(0, 1); }
  if (r < 0.35) { return int(2, 7); }
  if (r < 0.6) { return int(8, 30); }
  if (r < 0.8) { return int(31, 90); }
  return int(91, maxDays);
}
/** Spalten, die jede CoE-Tabelle hat. */
const COMMON = (): Record<string, unknown> => ({ statecode: 0, createdon: daysAgo(int(30, 900)), modifiedon: daysAgo(recentish(300)) });

// ---------- Basisdaten ----------
const TENANT_ID = '8f3c2a1d-6b4e-4c7a-9d2f-1e5b7a9c3d40';
export const COE_HOST = 'contoso-coe.crm4.dynamics.com';
export const DEFAULT_HOST = 'org8f3c2a1d.crm4.dynamics.com';

const FIRST = ['Anna', 'Lukas', 'Sophie', 'Maximilian', 'Laura', 'Tobias', 'Julia', 'Florian', 'Katharina', 'David',
  'Lena', 'Stefan', 'Marie', 'Michael', 'Sarah', 'Thomas', 'Elena', 'Andreas', 'Theresa', 'Markus', 'Hannah',
  'Christoph', 'Lisa', 'Daniel', 'Johanna', 'Patrick', 'Verena', 'Bernhard', 'Nina', 'Alexander'];
const LAST = ['Berger', 'Huber', 'Gruber', 'Wagner', 'Pichler', 'Moser', 'Steiner', 'Hofer', 'Leitner', 'Bauer',
  'Fuchs', 'Eder', 'Schmid', 'Winkler', 'Weber', 'Mayer', 'Lang', 'Brunner', 'Wimmer', 'Egger', 'Reiter', 'Haas',
  'Schwarz', 'Koller', 'Maier', 'Lechner', 'Wolf', 'Aigner', 'Baumgartner', 'Hackl'];
const DEPARTMENTS = ['IT', 'Finanzen', 'Vertrieb', 'Personal', 'Marketing', 'Einkauf', 'Produktion', 'Logistik', 'Recht', 'Qualität'];
const COUNTRIES = ['Österreich', 'Österreich', 'Österreich', 'Deutschland', 'Deutschland', 'Schweiz'];
const umlaut = (s: string): string => s.toLowerCase().replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss');

// ---------- Maker (ID = Entra-Objekt-ID) ----------
const makers: Record<string, unknown>[] = [];
const usedNames: { [n: string]: boolean } = {};
const addMaker = (f: string, l: string, guest?: boolean): Record<string, unknown> => {
  usedNames[f + l] = true;
  const country = pick(COUNTRIES);
  const upn = guest ? `${umlaut(f)}.${umlaut(l)}_partner.at#EXT#@contoso.onmicrosoft.com` : `${umlaut(f)}.${umlaut(l)}@contoso.com`;
  const m = {
    ...COMMON(),
    admin_makerid: guid(),
    admin_displayname: `${f} ${l}`,
    admin_userprincipalname: upn,
    admin_useremail: `${umlaut(f)}.${umlaut(l)}@${guest ? 'partner.at' : 'contoso.com'}`,
    admin_department: pick(DEPARTMENTS),
    admin_company: country === 'Schweiz' ? 'Contoso AG' : 'Contoso GmbH',
    admin_country: country,
    admin_city: country === 'Österreich' ? pick(['Wien', 'Graz', 'Linz', 'Salzburg', 'Innsbruck']) : country === 'Deutschland' ? pick(['München', 'Berlin', 'Hamburg']) : 'Zürich',
    admin_jobtitle: pick(['Sachbearbeiter:in', 'Teamleiter:in', 'Analyst:in', 'Projektleiter:in', 'Entwickler:in', 'Assistenz']),
    admin_userisguest: !!guest,
    admin_accountenabled: true,
    admin_numberofapps: 0, admin_numberofflows: 0, admin_numberofenvironments: 0
  };
  makers.push(m);
  return m;
};
while (makers.length < 64) {
  const f = pick(FIRST);
  const l = pick(LAST);
  if (!usedNames[f + l]) { addMaker(f, l, makers.length % 21 === 20); }
}
const DEV_OWNERS = ['Lena Brunner', 'Michael Fuchs', 'Sarah Gruber', 'David Leitner', 'Julia Moser', 'Florian Pichler', 'Theresa Eder',
  'Markus Wagner', 'Nina Bauer', 'Patrick Haas', 'Johanna Hofer', 'Christoph Lang', 'Verena Winkler', 'Tobias Steiner', 'Elena Koller',
  'Bernhard Aigner', 'Hannah Reiter', 'Andreas Mayer', 'Lisa Wimmer', 'Daniel Egger', 'Katharina Huber', 'Stefan Lechner'];
DEV_OWNERS.forEach(n => { const [f, l] = n.split(' '); if (!usedNames[f + l]) { addMaker(f, l); } });
const SYSTEM_MAKER = { ...COMMON(), admin_makerid: guid(), admin_displayname: 'SYSTEM', admin_userprincipalname: '', admin_department: '', admin_userisguest: false, admin_accountenabled: true, admin_numberofapps: 0, admin_numberofflows: 0 };
makers.push(SYSTEM_MAKER);
const departed = makers.slice(0, 6);
departed.forEach(m => { m.admin_accountenabled = false; });
const humans = makers.filter(m => m !== SYSTEM_MAKER);
const byName = (n: string): Record<string, unknown> => humans.filter(m => m.admin_displayname === n)[0] || pick(humans);

// ---------- Environments ----------
interface IEnvSeed { name: string; sku: string; region: string; host?: string; weight: number; deleted?: boolean; id?: string; level?: string; disabled?: boolean; owner?: string }
const envSeeds: IEnvSeed[] = [
  { name: 'Contoso (default)', sku: 'Default', region: 'europe', host: DEFAULT_HOST, weight: 55, id: TENANT_ID, level: 'Basic' },
  { name: 'CoE', sku: 'Production', region: 'europe', host: COE_HOST, weight: 2, level: 'Standard' },
  { name: 'HR Produktion', sku: 'Production', region: 'europe', host: 'contoso-hr.crm4.dynamics.com', weight: 6, level: 'Standard' },
  { name: 'HR Test', sku: 'Sandbox', region: 'europe', host: 'contoso-hr-test.crm4.dynamics.com', weight: 3, level: 'Standard' },
  { name: 'Finanzen Produktion', sku: 'Production', region: 'europe', host: 'contoso-fin.crm4.dynamics.com', weight: 7, level: 'Standard' },
  { name: 'Finanzen Entwicklung', sku: 'Sandbox', region: 'europe', host: 'contoso-fin-dev.crm4.dynamics.com', weight: 4, level: 'Basic' },
  { name: 'Vertrieb Produktion', sku: 'Production', region: 'europe', host: 'contoso-sales.crm4.dynamics.com', weight: 6, level: 'Standard' },
  { name: 'Einkauf Sandbox', sku: 'Sandbox', region: 'europe', host: 'contoso-proc.crm4.dynamics.com', weight: 3, level: 'Basic' },
  { name: 'Projektraum Nord', sku: 'Teams', region: 'europe', host: 'org5b1e77c2.crm4.dynamics.com', weight: 3, level: 'Basic' },
  { name: 'Trial – Lukas Huber', sku: 'Trial', region: 'unitedstates', host: 'org77d1e9a3.crm.dynamics.com', weight: 2, level: 'Basic' },
  { name: 'Marketing Kampagnen', sku: 'Production', region: 'europe', host: 'contoso-mkt.crm4.dynamics.com', weight: 4, level: 'Standard' },
  { name: 'Altsystem Migration', sku: 'Sandbox', region: 'europe', host: 'contoso-legacy.crm4.dynamics.com', weight: 0, deleted: true }
];
DEV_OWNERS.forEach((owner, i) => {
  const r = rnd();
  envSeeds.push({
    name: `Dev – ${owner}`, sku: 'Developer', region: 'europe', owner,
    host: `orgdev${(i + 11).toString(16)}${Math.floor(rnd() * 1e6).toString(16)}.crm4.dynamics.com`,
    weight: r < 0.2 ? 0 : r < 0.6 ? 1 : 3,
    level: i % 3 === 0 ? 'Standard' : i % 5 === 1 ? '' : 'Basic',
    disabled: i % 7 === 4
  });
});
const environments = envSeeds.map(e => {
  const maker = e.owner ? byName(e.owner) : pick(humans);
  const id = e.id || guid();
  return {
    ...COMMON(),
    admin_environmentid: id,
    admin_environmentname: e.sku === 'Default' ? `Default-${id}` : id,
    admin_displayname: e.name,
    admin_environmentsku: e.sku,
    _admin_maker_value: maker.admin_makerid,
    admin_environmentmakerdisplayname: maker.admin_displayname,
    admin_governanceprotectionlevel: e.level || null,
    admin_environmentruntimestate: e.disabled ? 'Disabled' : 'Enabled',
    admin_environmentstate: e.deleted ? 'Deleted' : 'Ready',
    admin_region: e.region,
    admin_environmentcdsinstanceurl: e.host ? `https://${e.host}/` : null,
    admin_environmentcreatedon: daysAgo(e.sku === 'Default' ? 1650 : int(40, 1200)),
    admin_environmentdeleted: !!e.deleted,
    admin_excusefrominventory: false,
    admin_securitygroupname: e.sku === 'Production' ? `SG-PP-${e.name.split(' ')[0]}` : 'Not assigned',
    admin_numberofapps: 0, admin_numberofflows: 0
  };
});
const envWeighted: number[] = [];
envSeeds.forEach((e, i) => { for (let k = 0; k < e.weight; k++) { envWeighted.push(i); } });

// ---------- Connectors ----------
const CONNECTORS = ['SharePoint', 'Office 365 Outlook', 'Microsoft Teams', 'OneDrive for Business', 'Excel Online (Business)',
  'Approvals', 'Office 365 Users', 'Microsoft Dataverse', 'SQL Server', 'HTTP', 'Azure Blob Storage', 'Planner',
  'Microsoft Forms', 'Notifications', 'Power BI', 'Salesforce', 'SAP ERP', 'Adobe PDF Services', 'DocuSign', '1DocStop',
  'Contoso Personal-API', 'Contoso Lager-API'];
const connectors = CONNECTORS.map(n => ({
  ...COMMON(),
  admin_connectorid: guid(),
  admin_displayname: n,
  admin_name: `shared_${n.toLowerCase().replace(/[^a-z0-9]/g, '')}`,
  admin_tier: /Contoso/.test(n) ? 'Custom' : /SQL|HTTP|Azure|Salesforce|SAP|Adobe|DocuSign|Dataverse|1DocStop/.test(n) ? 'Premium' : 'Standard',
  admin_publisher: /Contoso/.test(n) ? 'Contoso' : n === '1DocStop' ? '1DocStop' : 'Microsoft',
  admin_iscustomapi: /Contoso/.test(n),
  admin_connectordeleted: false,
  admin_numapps: 0, admin_numflows: 0
}));
const commonConnectors = connectors.slice(0, 8);

// ---------- Apps ----------
const APP_WORDS_A = ['Urlaubs', 'Reise', 'Spesen', 'Inventar', 'Besucher', 'Schulungs', 'Projekt', 'Raum', 'Fuhrpark', 'Angebots',
  'Lieferanten', 'Qualitäts', 'Onboarding', 'Wartungs', 'Kunden', 'Rechnungs', 'Bestell', 'Zeit', 'Ideen', 'Störungs'];
const APP_WORDS_B = ['antrag', 'buchung', 'verwaltung', 'erfassung', 'tracker', 'portal', 'übersicht', 'checkliste', 'planer', 'meldung'];
const APP_TYPES: { [v: number]: string } = { 597910000: 'Canvas', 597910001: 'Model Driven', 597910002: 'SharePointFormApp', 597910003: 'Custom Page', 597910004: 'Component Library' };
const apps: Record<string, unknown>[] = [];
for (let i = 0; i < 6840; i++) {
  const envIdx = pick(envWeighted);
  const env = environments[envIdx];
  const orphan = chance(0.11);
  const owner = orphan ? pick(departed) : pick(humans);
  const type = envSeeds[envIdx].sku === 'Default' ? (chance(0.9) ? 597910000 : 597910002) : pick([597910000, 597910000, 597910001, 597910003]);
  const created = int(20, 1400);
  const modified = Math.min(created, recentish(700));
  const launched = chance(0.85) ? Math.min(recentish(400), created) : null;
  const name = `${pick(APP_WORDS_A)}${pick(APP_WORDS_B)}${chance(0.2) ? ' ' + pick(['v2', '(Test)', 'Neu', '2025', 'Mobil']) : ''}`;
  apps.push({
    ...COMMON(),
    admin_appid: guid(),
    admin_displayname: name.charAt(0).toUpperCase() + name.slice(1),
    _admin_appowner_value: owner.admin_makerid,
    _admin_appenvironment_value: env.admin_environmentid,
    admin_appenvironmentid: env.admin_environmentname,
    admin_powerappstype: type,
    cr5d5_appisorphaned: orphan ? 'Yes' : 'No',
    admin_applastlaunchedon: launched === null ? null : daysAgo(Math.max(0, launched)),
    admin_appmodifiedon: daysAgo(modified),
    admin_appcreatedon: daysAgo(created),
    admin_appdeleted: chance(0.04),
    admin_appplanclassification: chance(0.28) ? 'Premium' : 'Standard',
    admin_appsharedusers: int(0, 140),
    admin_appsharedgroups: int(0, 6),
    // wie in eurem CoE: Spalte bleibt leer – die Übersicht zählt dann die Starts im Audit-Log
    admin_launchesinthepast30days: null
  });
}

// ---------- Flows ----------
const FLOW_TEMPLATES = ['Genehmigung {x} an Vorgesetzte', 'Täglicher Bericht {x}', 'Benachrichtigung bei neuem {x}',
  '{x} aus E-Mail in SharePoint speichern', 'Erinnerung {x} fällig', 'Sync {x} → Excel', '{x}-Eingang an Teams posten',
  'Wöchentliche Zusammenfassung {x}', '{x} archivieren', 'Anhänge {x} ablegen'];
const FLOW_X = ['Urlaubsantrag', 'Rechnung', 'Bestellung', 'Lieferschein', 'Reisekosten', 'Störungsmeldung', 'Vertrag', 'Angebot',
  'Bewerbung', 'Schulung', 'Wartung', 'Kundenanfrage'];
const TRIGGERS = ['OpenApiConnection : SharePoint – Wenn ein Element erstellt wird', 'Recurrence', 'Request : manual',
  'OpenApiConnectionWebhook : Office 365 Outlook – Bei neuer E-Mail', 'OpenApiConnection : Dataverse – Wenn eine Zeile hinzugefügt wird'];
const flows: Record<string, unknown>[] = [];
for (let i = 0; i < 12460; i++) {
  const envIdx = pick(envWeighted);
  const env = environments[envIdx];
  const orphan = chance(0.09);
  const creator = orphan ? pick(departed) : pick(humans);
  const r = rnd();
  const state = r < 0.74 ? 'Started' : r < 0.93 ? 'Stopped' : 'Suspended';
  const created = int(10, 1300);
  const modified = Math.min(created, recentish(600));
  const id = guid();
  flows.push({
    ...COMMON(),
    admin_flowid: id,
    admin_flowidstring: id,
    admin_displayname: pick(FLOW_TEMPLATES).replace('{x}', pick(FLOW_X)),
    _admin_flowcreator_value: creator.admin_makerid,
    _admin_derivedowner_value: creator.admin_makerid,
    _admin_flowenvironment_value: env.admin_environmentid,
    admin_flowenvironmentid: env.admin_environmentname,
    admin_flowstate: state,
    cr5d5_flowisorphaned: orphan ? 'Yes' : 'No',
    cr5d5_flowtrigger: pick(TRIGGERS),
    admin_flowmodifiedon: daysAgo(modified),
    admin_flowcreatedon: daysAgo(created),
    admin_flowdeleted: chance(0.04),
    admin_flowispublished: true
  });
}

// Zähler wie im CoE (Environment, Maker)
const envById: { [id: string]: Record<string, unknown> } = {};
environments.forEach(e => { envById[e.admin_environmentid] = e; });
const makerById: { [id: string]: Record<string, unknown> } = {};
makers.forEach(m => { makerById[String(m.admin_makerid)] = m; });
apps.forEach(a => {
  if (a.admin_appdeleted) { return; }
  (envById[String(a._admin_appenvironment_value)].admin_numberofapps as number)++;
  (makerById[String(a._admin_appowner_value)].admin_numberofapps as number)++;
});
flows.forEach(f => {
  if (f.admin_flowdeleted) { return; }
  (envById[String(f._admin_flowenvironment_value)].admin_numberofflows as number)++;
  (makerById[String(f._admin_flowcreator_value)].admin_numberofflows as number)++;
});

// ---------- Power Platform User (User / Gruppe / Tenant) ----------
const ppUsers: Record<string, unknown>[] = humans.map(m => ({
  ...COMMON(),
  admin_powerplatformuserid: m.admin_makerid,
  admin_displayname: m.admin_displayname,
  admin_type: 'User',
  admin_userprincipalname: m.admin_userprincipalname,
  admin_useremail: m.admin_useremail,
  admin_userisguest: m.admin_userisguest,
  admin_groupsize: null
}));
['SG Vertrieb Österreich', 'SG Alle Mitarbeitenden HR', 'SG Power Platform Maker', 'M365 Projekt Nord', 'SG Einkauf'].forEach((g, i) => ppUsers.push({
  ...COMMON(), admin_powerplatformuserid: guid(), admin_displayname: g, admin_type: 'Group', admin_userprincipalname: null,
  admin_useremail: null, admin_userisguest: false, admin_groupsize: [48, 312, 95, 17, 36][i]
}));
ppUsers.push({ ...COMMON(), admin_powerplatformuserid: guid(), admin_displayname: 'Contoso (gesamter Tenant)', admin_type: 'Tenant', admin_groupsize: null, admin_userisguest: false });
const ppPrincipals = ppUsers.filter(u => u.admin_type !== 'User');

// ---------- Lazy erzeugte, große Verknüpfungstabellen ----------
function lazyRows(make: () => Record<string, unknown>[]): () => Record<string, unknown>[] {
  let cache: Record<string, unknown>[] | undefined;
  return () => (cache = cache || make());
}
const userRoles = lazyRows(() => {
  const out: Record<string, unknown>[] = [];
  const add = (res: Record<string, unknown>, kind: 'App' | 'Flow', ownerId: unknown): void => {
    const envId = kind === 'App' ? res._admin_appenvironment_value : res._admin_flowenvironment_value;
    const link = kind === 'App' ? { _admin_app_value: res.admin_appid, _admin_flow_value: null } : { _admin_app_value: null, _admin_flow_value: res.admin_flowid };
    out.push({ ...COMMON(), admin_powerplatformuserroleid: guid(), admin_name: 'Owner', admin_resourcetype: kind, admin_rolename: 'Owner', _admin_powerplatformuser_value: ownerId, _admin_environment_value: envId, ...link });
    const n = kind === 'App' ? int(0, 4) : int(0, 2);
    for (let i = 0; i < n; i++) {
      const principal = chance(0.15) ? pick(ppPrincipals) : pick(ppUsers);
      const role = kind === 'App' ? (chance(0.8) ? 'CanView' : 'CanEdit') : 'CanEdit';
      out.push({ ...COMMON(), admin_powerplatformuserroleid: guid(), admin_name: role, admin_resourcetype: kind, admin_rolename: role, _admin_powerplatformuser_value: principal.admin_powerplatformuserid, _admin_environment_value: envId, ...link });
    }
  };
  apps.forEach(a => add(a, 'App', a._admin_appowner_value));
  flows.forEach(f => add(f, 'Flow', f._admin_flowcreator_value));
  return out;
});
const connectionRefs = lazyRows(() => {
  const out: Record<string, unknown>[] = [];
  const add = (appId: unknown, flowId: unknown): void => {
    const n = int(1, 3);
    const used: { [k: string]: boolean } = {};
    for (let i = 0; i < n; i++) {
      const c = chance(0.75) ? pick(commonConnectors) : pick(connectors);
      if (used[String(c.admin_connectorid)]) { continue; }
      used[String(c.admin_connectorid)] = true;
      out.push({ ...COMMON(), admin_connectionreferenceid: guid(), admin_name: String(c.admin_name), _admin_app_value: appId, _admin_flow_value: flowId, _admin_connector_value: c.admin_connectorid, admin_flow_connectionreferencesource: flowId ? pick(['Trigger', 'Aktion', 'Trigger, Aktion']) : null });
    }
  };
  apps.forEach(a => add(a.admin_appid, null));
  flows.forEach(f => add(null, f.admin_flowid));
  return out;
});
// Connector-Zähler aus den Verbindungen (beim ersten Zugriff auf Connectors berechnet)
let connectorCountsDone = false;
const connectorRows = (): Record<string, unknown>[] => {
  if (!connectorCountsDone) {
    connectorCountsDone = true;
    const byId: { [id: string]: Record<string, unknown> } = {};
    connectors.forEach(c => { byId[String(c.admin_connectorid)] = c; });
    connectionRefs().forEach(r => {
      const c = byId[String(r._admin_connector_value)];
      if (r._admin_app_value) { (c.admin_numapps as number)++; } else { (c.admin_numflows as number)++; }
    });
  }
  return connectors;
};
const ACTIONS = ['OpenApiConnection', 'Condition', 'Compose', 'ApplyToEach', 'InitializeVariable', 'Http'];
const flowActions = lazyRows(() => {
  const out: Record<string, unknown>[] = [];
  flows.forEach(f => {
    const n = int(2, 4);
    for (let i = 0; i < n; i++) {
      out.push({ ...COMMON(), admin_flowactiondetailid: guid(), admin_name: `Aktion ${i + 1}`, _admin_flow_value: f.admin_flowid, admin_actiontype: i === 0 ? 'Trigger' : pick(ACTIONS), admin_operation: pick(['GetItems', 'SendEmailV2', 'PostMessageToConversation', 'CreateFile', 'ListRecords', '']), cr5d5_istypetrigger: i === 0 });
    }
  });
  return out;
});
const auditLog = lazyRows(() => {
  const out: Record<string, unknown>[] = [];
  apps.slice(0, 3000).forEach(a => {
    if (!a.admin_applastlaunchedon) { return; }
    const n = Math.round(60 / Math.pow(int(1, 400), 0.8));
    for (let i = 0; i < n; i++) {
      const u = pick(humans);
      out.push({ ...COMMON(), admin_auditlogid: guid(), admin_title: `LaunchPowerApp – ${a.admin_displayname}`, _admin_applookup_value: a.admin_appid, admin_appenvironment: envById[String(a._admin_appenvironment_value)].admin_displayname, admin_operation: 'LaunchPowerApp', admin_workload: 'PowerApps', admin_userupn: u.admin_userprincipalname, admin_creationtime: daysAgo(int(0, 60)) });
    }
  });
  return out;
});

// ---------- weitere Inventar-Tabellen ----------
const envPick = (): Record<string, unknown> => environments[pick(envWeighted)];
const bots = Array.from({ length: 64 }, (_x, i) => {
  const env = envPick(); const owner = pick(humans);
  return { ...COMMON(), admin_pvaid: guid(), admin_name: `${pick(['HR', 'IT', 'Vertrieb', 'Einkauf', 'Kundenservice'])}-${pick(['Assistent', 'Helpdesk', 'FAQ-Bot', 'Copilot'])} ${i + 1}`, _admin_pvaowner_value: owner.admin_makerid, _admin_pvaenvironment_value: env.admin_environmentid, admin_pvastate: chance(0.7) ? 'Published' : 'Provisioned', admin_pvanumberlaunches: int(0, 900), admin_pvalastlaunchedon: daysAgo(recentish(200)), admin_bottemplate: pick(['default-2.1.0', 'default-2.2.0', 'Copilot']), admin_pvadeleted: false };
});
const rpas = Array.from({ length: 85 }, () => {
  const env = envPick(); const owner = pick(humans);
  return { ...COMMON(), admin_rpaid: guid(), admin_name: guid(), admin_displayname: `${pick(['SAP-Export', 'Rechnungserfassung', 'Excel-Makro', 'Legacy-Login', 'PDF-Ablage'])} ${pick(['täglich', 'Monatsabschluss', 'manuell'])}`, _admin_rpaowner_value: owner.admin_makerid, _admin_desktopflowenvironment_value: env.admin_environmentid, admin_rpastate: pick(['Published', 'Published', 'Unpublished']), admin_rpamodifiedon: daysAgo(recentish(400)) };
});
const portals = Array.from({ length: 18 }, (_x, i) => {
  const env = envPick(); const owner = pick(humans);
  const name = `${pick(['Lieferanten', 'Bewerber', 'Partner', 'Kunden', 'Veranstaltungs'])}portal${i > 9 ? ' ' + i : ''}`;
  return { ...COMMON(), admin_portalid: guid(), admin_portaldisplayname: name, _admin_portalowner_value: owner.admin_makerid, _admin_portalenvironment_value: env.admin_environmentid, admin_portalwebsitename: name, admin_portalwebsiterecordstatus: pick(['Active', 'Active', 'Inactive']), admin_authtablepermissionsenabled: chance(0.6), admin_portalmodifiedon: daysAgo(recentish(300)) };
});
const bpfs = Array.from({ length: 36 }, () => {
  const env = envPick(); const owner = pick(humans); const ent = pick(['account', 'opportunity', 'incident', 'lead', 'cr123_antrag']);
  return { ...COMMON(), admin_businessprocessflowsid: guid(), admin_name: `${ent}process`, admin_businessprocessflowdisplayname: `${pick(['Vertriebs', 'Service', 'Antrags', 'Lead'])}prozess`, _admin_businessprocessflowowner_value: owner.admin_makerid, _admin_businessprocessflowenvironment_value: env.admin_environmentid, admin_businessprocessflowprimaryentity: ent, admin_businessprocessflowstatus: chance(0.8) ? 'Activated' : 'Draft', admin_businessprocessflowmodifiedon: daysAgo(recentish(500)), admin_businessprocessflowdeleted: false };
});
const coeSolutions = Array.from({ length: 160 }, () => {
  const env = envPick(); const owner = pick(humans); const managed = chance(0.45);
  return { ...COMMON(), admin_solutionid: guid(), admin_name: managed ? pick(['msdyn_AIBuilder', 'msdyn_FlowApprovals', 'PowerAppsChecker', 'msdyn_TeamsApprovals']) : `Contoso${pick(['HR', 'Finanzen', 'Vertrieb', 'Einkauf'])}${pick(['Apps', 'Automatisierung', 'Basis'])}`, _admin_solutionenvironment_value: env.admin_environmentid, _admin_solutioncreator_value: owner.admin_makerid, admin_solutionpublisher: managed ? 'Microsoft Corporation' : 'Contoso IT', admin_solutionismanaged: managed, admin_solutiondeleted: false };
});
const aiModels = Array.from({ length: 28 }, () => {
  const env = envPick(); const owner = pick(humans);
  return { ...COMMON(), admin_aibuildermodelid: guid(), admin_name: `${pick(['Rechnungen', 'Lieferscheine', 'Bewerbungen', 'Formulare'])} – ${pick(['Dokumentverarbeitung', 'Klassifizierung', 'Texterkennung'])}`, _admin_aicreator_value: owner.admin_makerid, _admin_aienvironment_value: env.admin_environmentid, admin_aitemplate: pick(['DocumentScanning', 'TextClassification', 'ObjectDetection', 'Prediction']), admin_aistatus: pick(['Published', 'Trained', 'Draft']), admin_ailastused: daysAgo(recentish(300)), admin_aideleted: false };
});
const ROLE_TEMPLATES: { [role: string]: string } = {
  'System Administrator': '627090ff-40a3-4053-8790-584edc5be201', 'System Customizer': '119f245c-3cc8-4b62-b31c-d1a046ced15d',
  'Environment Maker': 'd5e2ae5a-1e4e-4ad8-a9c3-9c9a1c5a0f00', 'Basic User': 'b9b5a2f3-7a1e-4ad5-9c43-1ab3ed4f0bb2', 'Power Platform Admin SR': '0f3cbd2c-5a9e-4c2b-8f1e-3c8d6e2a9b11'
};
const securityRoles = Array.from({ length: 420 }, () => {
  const env = environments[int(0, environments.length - 1)]; const u = pick(humans); const viaGroup = chance(0.3);
  const role = pick(['System Administrator', 'System Customizer', 'Environment Maker', 'Basic User', 'Power Platform Admin SR']);
  return { ...COMMON(), admin_environmentsecurityrolepermissionid: guid(), admin_name: u.admin_displayname, admin_securityrolename: role, admin_securityroletemplateid: ROLE_TEMPLATES[role], _admin_environment_value: env.admin_environmentid, admin_assignedvia: viaGroup ? 'Team' : 'Direkt', admin_groupname: viaGroup ? pick(ppPrincipals).admin_displayname : null, admin_businessunitname: `org-${String(env.admin_displayname).toLowerCase().replace(/[^a-z]/g, '').slice(0, 10)}`, admin_usertype: u.admin_userisguest ? 597910001 : 597910000, admin_useridentityprovider: 'Entra ID' };
});
// Jedes Environment: SYSTEM und ein App-Benutzer als System Administrator (werden im Cockpit ausgeblendet)
environments.forEach(env => {
  [['SYSTEM', 597910000], ['# CoE Sync App', 597910002]].forEach(([name, type]) => securityRoles.push({
    ...COMMON(), admin_environmentsecurityrolepermissionid: guid(), admin_name: name, admin_securityrolename: 'System Administrator',
    admin_securityroletemplateid: ROLE_TEMPLATES['System Administrator'], _admin_environment_value: env.admin_environmentid,
    admin_assignedvia: 'Direkt', admin_groupname: null, admin_businessunitname: 'org', admin_usertype: type, admin_useridentityprovider: 'Entra ID'
  }));
});
const connIdentities = Array.from({ length: 260 }, () => {
  const env = envPick(); const c = pick(connectors); const u = pick(ppUsers.filter(p => p.admin_type === 'User')); const cross = chance(0.05);
  return { ...COMMON(), admin_connectionreferenceidentityid: guid(), admin_name: `${c.admin_name}-${String(u.admin_displayname).split(' ')[1]}`, admin_accountname: cross ? 'extern@fabrikam.com' : u.admin_userprincipalname, _admin_environment_value: env.admin_environmentid, _admin_connector_value: c.admin_connectorid, _admin_connectionreferencecreator_value: u.admin_powerplatformuserid, admin_noneorcrosstenantidentity: cross };
});
const capacity: Record<string, unknown>[] = [];
environments.filter(e => !e.admin_environmentdeleted && e.admin_environmentcdsinstanceurl).forEach(e => {
  [['Database', 1.0], ['File', 2.5], ['Log', 0.3]].forEach(([type, scale]) => {
    const used = Math.round(rnd() * 4000 * (scale as number));
    capacity.push({ ...COMMON(), admin_environmentcapacityid: guid(), admin_capacitytype: type, _admin_environment_value: e.admin_environmentid, admin_actualconsumption: used, admin_ratedconsumption: Math.round(used * 1.05 * 100) / 100, admin_approvedcapacity: null, admin_percentusedofapproved: null, admin_capacityunit: 'MB', admin_environmentdeleted: false });
  });
});

// ---------- Solutions/Publisher im Default Environment (live aus dessen Dataverse) ----------
const publishers = [
  { publisherid: guid(), friendlyname: 'Default Publisher for org8f3c2a1d', uniquename: 'DefaultPublisherorg8f3c2a1d' },
  { publisherid: guid(), friendlyname: 'CDS Default Publisher', uniquename: 'Cr6b2e1' },
  { publisherid: guid(), friendlyname: 'Microsoft Corporation', uniquename: 'microsoftcorporation' },
  { publisherid: guid(), friendlyname: 'Contoso IT', uniquename: 'contosoit' },
  { publisherid: guid(), friendlyname: 'Dynamics 365', uniquename: 'dynamics365' }
];
const solutionSeeds: [string, string, number, boolean, boolean][] = [
  ['Default Solution', 'Default', 0, false, true], ['Common Data Services Default Solution', 'Crc6b2e1', 1, false, true],
  ['Contoso Basis-Apps', 'ContosoBasisApps', 3, false, true], ['Contoso Urlaubsverwaltung', 'ContosoUrlaub', 3, false, true],
  ['Contoso Spesen', 'ContosoSpesen', 3, false, true], ['Contoso Besucheranmeldung', 'ContosoBesucher', 3, true, true],
  ['Teams Approvals', 'msdyn_TeamsApprovals', 2, true, true], ['Power Apps Checker', 'msdyn_PowerAppsChecker', 2, true, true],
  ['Power Automate Desktop', 'msdyn_PowerAutomateDesktop', 2, true, true], ['AI Builder', 'msdyn_AIBuilder', 2, true, true],
  ['Approvals Core', 'msdyn_FlowApprovals', 2, true, true], ['Connection References', 'msdyn_ConnectionReferences', 2, true, true],
  ['Dynamics 365 Base', 'msdynce_Base', 4, true, true], ['Copilot Studio Basis', 'msdyn_CopilotStudio', 2, true, true],
  ['Active', 'Active', 2, false, false], ['Basic', 'Basic', 2, true, false], ['System', 'System', 2, true, false]
];
const solutions = solutionSeeds.map(([name, unique, pub, managed, visible]) => {
  const installed = managed ? int(5, 900) : int(100, 1600);
  return {
    solutionid: guid(), friendlyname: name, uniquename: unique,
    version: managed ? `${int(1, 9)}.${int(0, 4)}.${int(0, 999)}.${int(0, 99)}` : `1.0.${int(0, 40)}.${int(0, 9)}`,
    ismanaged: managed, isvisible: visible, _publisherid_value: publishers[pub].publisherid,
    installedon: daysAgo(installed), modifiedon: daysAgo(Math.min(installed, recentish(400))), createdon: daysAgo(installed),
    description: managed ? 'Von Microsoft bereitgestellte verwaltete Solution.' : 'Nicht verwaltete Contoso-Solution.'
  };
});


// ---------- Sicherheit & Governance (nachträglich, damit die übrigen Beispieldaten gleich bleiben) ----------
export const PIPELINES_HOST = 'contoso-pipelines.crm4.dynamics.com';
const pipelinesEnv: Record<string, unknown> = {
  ...COMMON(), admin_environmentid: guid(), admin_displayname: 'Pipelines Host', admin_environmentsku: 'Production',
  _admin_maker_value: humans[0].admin_makerid, admin_environmentmakerdisplayname: humans[0].admin_displayname,
  admin_governanceprotectionlevel: 'Standard', admin_environmentruntimestate: 'Enabled', admin_environmentstate: 'Ready',
  admin_region: 'europe', admin_environmentcdsinstanceurl: `https://${PIPELINES_HOST}/`, admin_environmentcreatedon: daysAgo(380),
  admin_environmentdeleted: false, admin_excusefrominventory: false, admin_securitygroupname: 'SG-PP-ALM', admin_numberofapps: 0, admin_numberofflows: 0
};
environments.push(pipelinesEnv as typeof environments[0]);
envById[String(pipelinesEnv.admin_environmentid)] = pipelinesEnv;
const envNamed = (n: string): Record<string, unknown> => environments.filter(e => e.admin_displayname === n)[0];
const liveEnvs = environments.filter(e => !e.admin_environmentdeleted);

// DLP-Richtlinien (N:N zu Environments und Connectors wie im CoE)
const DLP_SCOPES: { [v: number]: string } = { 597910000: 'All environments', 597910001: 'Include certain environments', 597910002: 'Exclude certain environments' };
const conn = (n: string): unknown => connectors.filter(c => c.admin_displayname === n)[0].admin_connectorid;
const excludedFromTenant = liveEnvs.filter(e => e.admin_environmentsku === 'Trial' || e.admin_environmentsku === 'Teams'
  || (e.admin_environmentsku === 'Developer' && chance(0.35)));
const dlpPolicies: Record<string, unknown>[] = [
  { name: 'Contoso Standard-DLP', scope: 597910002, desc: 'Basisrichtlinie für alle Environments außer den ausgenommenen.', envs: excludedFromTenant,
    blocked: ['HTTP', '1DocStop', 'Salesforce'], business: ['SharePoint', 'Office 365 Outlook', 'Microsoft Teams', 'OneDrive for Business', 'Approvals', 'Microsoft Dataverse'] },
  { name: 'Default Environment – streng', scope: 597910001, desc: 'Nur Microsoft-365-Connectors im Default Environment.', envs: [envNamed('Contoso (default)')],
    blocked: ['HTTP', 'SQL Server', 'Azure Blob Storage', 'Salesforce', 'SAP ERP', 'DocuSign', '1DocStop'], business: ['SharePoint', 'Office 365 Outlook', 'Microsoft Teams', 'Approvals'] },
  { name: 'Produktion – Business-Daten', scope: 597910001, desc: 'Produktive Environments mit SAP- und SQL-Anbindung.', envs: liveEnvs.filter(e => e.admin_environmentsku === 'Production'),
    blocked: ['1DocStop'], business: ['SharePoint', 'Microsoft Dataverse', 'SQL Server', 'SAP ERP', 'Contoso Personal-API', 'Contoso Lager-API'] },
  { name: 'Finanzen – Sonderregel', scope: 597910001, desc: '', envs: [envNamed('Finanzen Produktion'), envNamed('Finanzen Entwicklung')],
    blocked: ['HTTP', 'DocuSign'], business: ['SharePoint', 'Excel Online (Business)', 'SQL Server'] }
].map(d => ({
  ...COMMON(), admin_dlppolicyid: guid(), admin_name: d.name, admin_scope: d.scope, admin_description: d.desc || null, admin_isshared: false, admin_policyid: guid(),
  __nav_admin_policy_environments: (d.envs as Record<string, unknown>[]).map(e => e.admin_environmentid),
  __nav_admin_blocked_policy_connectors: (d.blocked as string[]).map(conn),
  __nav_admin_confidential_policy_connectors: (d.business as string[]).map(conn)
}));
liveEnvs.forEach(e => {
  const names = dlpPolicies.filter(d => {
    const listed = (d.__nav_admin_policy_environments as unknown[]).indexOf(e.admin_environmentid) >= 0;
    return d.admin_scope === 597910002 ? !listed : d.admin_scope === 597910000 || listed;
  }).map(d => String(d.admin_name));
  (e as Record<string, unknown>).admin_hasdlppolicy = names.length > 0;
  (e as Record<string, unknown>).admin_dlppolicies = names.length ? names.join(', ') : null;
});

// ---------- DLP: Connectors je App/Flow („App Connections“ / „Flow Connections“) ----------
// deterministisch aus der ID – verbraucht den Zufallsgenerator nicht (alle übrigen Testdaten bleiben gleich)
function hashOf(text: string): number {
  let h = 2166136261;
  for (let i = 0; i < text.length; i++) { h ^= text.charCodeAt(i); h = Math.imul(h, 16777619); }
  return h >>> 0;
}
function connsFor(id: string, app: boolean): string {
  let h = hashOf(id) || 1;
  const next = (): number => { h ^= h << 13; h ^= h >>> 17; h ^= h << 5; h >>>= 0; return h / 4294967296; };
  const n = 1 + Math.floor(next() * 3);
  const out: string[] = [];
  for (let i = 0; i < n; i++) {
    const list = next() < 0.86 ? commonConnectors : connectors;
    const c = String(list[Math.floor(next() * list.length)].admin_name);
    if (out.indexOf(c) < 0) { out.push(c); }
  }
  if (app && next() < 0.08) { out.push('shared_logicflows'); }
  return out.join(',');
}
apps.forEach(a => { a.admin_appconnections = connsFor(String(a.admin_appid), true); });
flows.forEach(f => { f.admin_flowconnections = connsFor(String(f.admin_flowid), false); });

// Entwürfe (CoE-Tabelle „Data Policy Drafts“) und Aufgabenliste („DLP Impact Analysis“) – feste IDs/Daten, ebenfalls ohne Zufallsgenerator
const fixedId = (n: number): string => `d1a90000-0000-4000-8000-${String(n).padStart(12, '0')}`;
const fixedDate = (days: number): string => new Date(NOW - days * DAY).toISOString().replace(/\.\d{3}Z$/, 'Z');
const FIXED = (created: number, modified: number): Record<string, unknown> => ({ statecode: 0, createdon: fixedDate(created), modifiedon: fixedDate(modified) });
const conRef = (n: string): { id: string; name: string; type: string } => {
  const name = String(connectors.filter(c => c.admin_displayname === n)[0].admin_name);
  return { id: `/providers/Microsoft.PowerApps/apis/${name}`, name, type: 'Microsoft.PowerApps/apis' };
};
const draftEnvs = liveEnvs.filter(e => e.admin_environmentsku === 'Production').slice(0, 3).map(e => ({
  name: String(e.admin_environmentname), type: 'Microsoft.BusinessAppPlatform/scopes/environments',
  id: `/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${e.admin_environmentname}`
}));
const dlpDrafts: Record<string, unknown>[] = [{
  ...FIXED(12, 3), admin_datapolicydraftsid: fixedId(1), admin_name: 'Produktion ohne SQL', admin_displayname: 'Produktion – Business-Daten', admin_apiversion: '2020-10-01',
  admin_defaultconnectorclassification: 'General', admin_scope: 'OnlyEnvironments', admin_environments: JSON.stringify(draftEnvs),
  admin_connectorgroups: JSON.stringify([
    { classification: 'Confidential', connectors: ['SharePoint', 'Microsoft Dataverse', 'SAP ERP'].map(conRef) },
    { classification: 'General', connectors: ['Office 365 Outlook', 'Microsoft Teams'].map(conRef) },
    { classification: 'Blocked', connectors: ['1DocStop', 'SQL Server'].map(conRef) }
  ]),
  _createdby_value: 'Anna Berger', _modifiedby_value: 'Anna Berger'
}];
const dlpTasks: Record<string, unknown>[] = apps.filter(a => !a.admin_appdeleted).slice(0, 3).map((a, i) => ({
  ...FIXED(9 - i, 2), admin_dlpimpactanalysisid: fixedId(100 + i), admin_name: a.admin_displayname, admin_dlppolicyname: 'Contoso Standard-DLP', admin_datapolicyid: '',
  _admin_impactedapp_value: a.admin_appid, _admin_impactedflow_value: null, _admin_impactedenvironment_value: a._admin_appenvironment_value,
  _admin_impactowner_value: a._admin_appowner_value, admin_conflictingconnectorbusiness: 'SharePoint', admin_conflictingconnectornonbusiness: i === 2 ? '' : 'Excel Online (Business)',
  admin_conflictingconnectorblocked: i === 2 ? 'HTTP' : '', admin_decision: [597910000, 597910004, null][i], admin_decisionstatus: [597910000, 597910000, null][i],
  admin_actionrequiredon: i === 0 ? fixedDate(-14) : null
}));

// Freigaben, Gäste, DLP-Verstöße, unverschlüsselte Custom Connectors
apps.forEach(a => {
  a.admin_appsharedwithtenant = chance(0.025);
  a.admin_appsharedguestuser = chance(0.035);
  a.admin_dlppoliciesviolated = chance(0.012) ? pick(['Contoso Standard-DLP', 'Default Environment – streng']) : null;
  a.admin_appisquarantined = !!a.admin_dlppoliciesviolated && chance(0.5);
  a.admin_quarantineapp = a.admin_appisquarantined;
  a.admin_issolutionapp = a.admin_powerappstype !== 597910000 || chance(0.25);
});
flows.forEach(f => { f.admin_flowsharedguestuser = chance(0.015); f.admin_issolutionflow = chance(0.3); });
connectors.forEach(c => {
  c.admin_customconnectorhost = c.admin_iscustomapi ? (c.admin_displayname === 'Contoso Lager-API' ? 'lager.intern.contoso.com' : 'api.contoso.com') : null;
  c.admin_customconnectorhostishttp = c.admin_displayname === 'Contoso Lager-API';
});

// Anwendungsbenutzer wie bei euch: „# …“ als System Administrator, aber nicht als App-Benutzer markiert
environments.forEach(env => {
  ['# CatalogServiceEur', '# CCADataAnalyticsML'].forEach(name => securityRoles.push({
    ...COMMON(), admin_environmentsecurityrolepermissionid: guid(), admin_name: name, admin_securityrolename: 'System Administrator',
    admin_securityroletemplateid: ROLE_TEMPLATES['System Administrator'], _admin_environment_value: env.admin_environmentid,
    admin_assignedvia: 'Direkt', admin_groupname: null, admin_businessunitname: 'org', admin_usertype: 597910000, admin_useridentityprovider: 'Entra ID'
  }));
});
// Gäste ohne Postfach-Adresse im CoE (nur UPN mit #EXT#) – werden in der E-Mail als „ohne Adresse“ gemeldet
makers.filter(m => m.admin_userisguest).forEach(m => { m.admin_useremail = null; });

// ---------- Power Platform Pipelines (eigenes Host-Environment) ----------
const sysUsers = humans.slice(6, 14).map(m => ({ systemuserid: guid(), fullname: m.admin_displayname, ...COMMON() }));
const ENV_ROLE: { [v: number]: string } = { 200000000: 'Entwicklungsumgebung', 200000001: 'Zielumgebung' };
const VALIDATION: { [v: number]: string } = { 200000000: 'Ausstehend', 200000001: 'Erfolgreich', 200000002: 'Fehlgeschlagen' };
const RUN_STATUS: { [v: number]: string } = { 200000000: 'Nicht gestartet', 200000002: 'Wird ausgeführt', 200000003: 'Erfolgreich', 200000004: 'Fehlgeschlagen', 200000005: 'Abgebrochen', 200000006: 'Genehmigung ausstehend' };
const depEnvSeeds: [string, number][] = [['HR Test', 200000000], ['HR Produktion', 200000001], ['Finanzen Entwicklung', 200000000],
  ['Finanzen Produktion', 200000001], ['Einkauf Sandbox', 200000000], ['Vertrieb Produktion', 200000001], ['Marketing Kampagnen', 200000001]];
const depEnvs = depEnvSeeds.map(([n, role]) => ({
  ...COMMON(), deploymentenvironmentid: guid(), name: n, environmenttype: role, validationstatus: n === 'Marketing Kampagnen' ? 200000002 : 200000001,
  environmentid: envNamed(n).admin_environmentid
}));
const depEnv = (n: string): unknown => depEnvs.filter(e => e.name === n)[0].deploymentenvironmentid;
const pipelineSeeds: [string, string, string[], string][] = [
  ['HR-Apps', 'HR Test', ['HR Produktion'], 'Urlaubs- und Onboarding-Apps'],
  ['Finanzen', 'Finanzen Entwicklung', ['Finanzen Produktion'], 'Spesen, Rechnungsfreigabe'],
  ['Einkauf & Vertrieb', 'Einkauf Sandbox', ['Vertrieb Produktion', 'Marketing Kampagnen'], 'Gemeinsame Lösungen für Einkauf und Vertrieb']
];
const pipelines: Record<string, unknown>[] = [];
const stages: Record<string, unknown>[] = [];
pipelineSeeds.forEach(([name, , targets, desc]) => {
  const pl = { ...COMMON(), deploymentpipelineid: guid(), name, description: desc, _ownerid_value: pick(sysUsers).systemuserid };
  pipelines.push(pl);
  let prev: unknown = null;
  targets.forEach((t, i) => {
    const st = { ...COMMON(), deploymentstageid: guid(), name: i === 0 && targets.length > 1 ? 'Test' : 'Produktion', _deploymentpipelineid_value: pl.deploymentpipelineid,
      _targetdeploymentenvironmentid_value: depEnv(t), _previousdeploymentstageid_value: prev };
    stages.push(st);
    prev = st.deploymentstageid;
  });
});
const SOLUTION_NAMES: { [p: string]: string[] } = { 'HR-Apps': ['ContosoUrlaub', 'ContosoOnboarding'], 'Finanzen': ['ContosoSpesen', 'ContosoRechnungsfreigabe'], 'Einkauf & Vertrieb': ['ContosoBestellung', 'ContosoAngebote', 'ContosoBasis'] };
const stageRuns: Record<string, unknown>[] = [];
for (let i = 0; i < 140; i++) {
  const pl = pick(pipelines);
  const st = pick(stages.filter(x => x._deploymentpipelineid_value === pl.deploymentpipelineid));
  const r = rnd();
  const status = i < 2 ? 200000002 : i < 4 ? 200000006 : r < 0.8 ? 200000003 : r < 0.93 ? 200000004 : 200000005;
  const startDays = i < 4 ? 0 : int(0, 160);
  const start = daysAgo(startDays);
  stageRuns.push({
    ...COMMON(), deploymentstagerunid: guid(), artifactname: pick(SOLUTION_NAMES[String(pl.name)]),
    artifactversion: `1.${int(0, 9)}.${int(0, 40)}.${int(0, 9)}`, _deploymentpipelineid_value: pl.deploymentpipelineid, _deploymentstageid_value: st.deploymentstageid,
    stagerunstatus: status, _createdby_value: pick(sysUsers).systemuserid, starttime: start,
    endtime: status === 200000002 || status === 200000006 ? null : new Date(Date.parse(start) + int(2, 25) * 60000).toISOString().replace(/\.\d{3}Z$/, 'Z'),
    createdon: start
  });
}

// Entwicklungs-Environments je Pipeline (N:N „deploymentpipeline_deploymentenvironment“) und je Bereitstellung
pipelineSeeds.forEach(([name, dev]) => {
  const pl = pipelines.filter(x => x.name === name)[0];
  pl.__nav_deploymentpipeline_deploymentenvironment = [depEnv(dev)];
});
stageRuns.forEach(r => {
  const pl = pipelines.filter(x => x.deploymentpipelineid === r._deploymentpipelineid_value)[0];
  r._devdeploymentenvironment_value = (pl.__nav_deploymentpipeline_deploymentenvironment as unknown[])[0];
});
depEnvs.forEach(e => { (e as Record<string, unknown>).errormessage = e.validationstatus === 200000002 ? 'Das Environment ist bereits mit einem anderen Pipelines-Host verknüpft (Platform host).' : null; });

// Zweiter Pipelines-Host: das CoE-Environment (feste IDs/Daten, ohne Zufallsgenerator)
const h2Id = (n: number): string => `b2b20000-0000-4000-8000-${String(n).padStart(12, '0')}`;
const h2Date = (days: number): string => new Date(Date.now() - days * 86400000).toISOString().replace(/\.\d{3}Z$/, 'Z');
const h2Common = (days: number): Record<string, unknown> => ({ statecode: 0, createdon: h2Date(days + 30), modifiedon: h2Date(days) });
const devEnvs2 = environments.filter(e => e.admin_environmentsku === 'Developer' && !e.admin_environmentdeleted).slice(0, 2);
const h2EnvSeeds: [Record<string, unknown>, number][] = [[devEnvs2[0], 200000000], [devEnvs2[1], 200000000], [envNamed('CoE'), 200000001], [envNamed('Trial – Lukas Huber'), 200000001]];
const depEnvs2 = h2EnvSeeds.map(([e, role], i) => ({
  ...h2Common(40 - i), deploymentenvironmentid: h2Id(10 + i), name: String(e.admin_displayname), environmenttype: role, validationstatus: 200000001,
  environmentid: e.admin_environmentid, errormessage: null
}));
const pipelines2: Record<string, unknown>[] = [
  { ...h2Common(12), deploymentpipelineid: h2Id(1), name: 'COE Prod', description: null, _ownerid_value: sysUsers[0].systemuserid, __nav_deploymentpipeline_deploymentenvironment: [h2Id(10)] },
  { ...h2Common(30), deploymentpipelineid: h2Id(2), name: 'Schulung', description: 'Schulungs-Apps für Citizen Developer', _ownerid_value: sysUsers[1].systemuserid, __nav_deploymentpipeline_deploymentenvironment: [h2Id(11)] }
];
const stages2: Record<string, unknown>[] = [
  { ...h2Common(12), deploymentstageid: h2Id(20), name: 'Produktion', _deploymentpipelineid_value: h2Id(1), _targetdeploymentenvironmentid_value: h2Id(12), _previousdeploymentstageid_value: null },
  { ...h2Common(30), deploymentstageid: h2Id(21), name: 'Test', _deploymentpipelineid_value: h2Id(2), _targetdeploymentenvironmentid_value: h2Id(13), _previousdeploymentstageid_value: null }
];
const stageRuns2: Record<string, unknown>[] = [
  [30, 1, 20, 'CoEStarterKitCore', '4.50.6.0', 200000003, 1], [31, 1, 20, 'CoEStarterKitCore', '4.50.5.0', 200000004, 5], [32, 2, 21, 'SchulungsApps', '1.2.0.0', 200000003, 9]
].map(([n, pl, st, art, ver, status, days]) => ({
  ...h2Common(Number(days)), deploymentstagerunid: h2Id(Number(n)), artifactname: art, artifactversion: ver, _deploymentpipelineid_value: h2Id(Number(pl)),
  _deploymentstageid_value: h2Id(Number(st)), stagerunstatus: status, _createdby_value: sysUsers[2].systemuserid, starttime: h2Date(Number(days)),
  endtime: h2Date(Number(days) - 0.01), _devdeploymentenvironment_value: Number(pl) === 1 ? h2Id(10) : h2Id(11)
}));

// ---------- Tabellen-Definitionen ----------
const A = (name: string, type: string, label: string, target?: string, options?: { [v: number]: string }): IAttr => ({ name, type, label, target, options });
const STD = [A('statecode', 'State', 'Status'), A('createdon', 'DateTime', 'Erstellt am'), A('modifiedon', 'DateTime', 'Geändert am')];

function table(logicalName: string, entitySet: string, displayName: string, primaryId: string, primaryName: string,
  attrs: IAttr[], rows: Record<string, unknown>[] | (() => Record<string, unknown>[])): ITableDef {
  const def = { logicalName, entitySet, displayName, primaryId, primaryName, attrs: [A(primaryId, 'Uniqueidentifier', displayName)].concat(attrs, STD) } as ITableDef;
  Object.defineProperty(def, 'rows', { get: typeof rows === 'function' ? rows : () => rows, enumerable: true });
  return def;
}
/** N:N-Beziehungen: Navigationseigenschaft → Zieltabelle (IDs stehen im Datensatz unter __nav_<name>). */
function withNav(def: ITableDef, navs: { [prop: string]: string }): ITableDef {
  def.navs = navs;
  return def;
}
function db(host: string, defs: ITableDef[]): IDatabase {
  const tables: { [n: string]: ITableDef } = {};
  defs.forEach(d => { tables[d.logicalName] = d; });
  return { host, tables };
}
const LK = (name: string, label: string, target: string): IAttr => A(name, 'Lookup', label, target);
/** Tabellen der Power Platform Pipelines (je Host-Environment) */
function pipelineTables(pl: Record<string, unknown>[], st: Record<string, unknown>[], runs: Record<string, unknown>[], envs: Record<string, unknown>[]): ITableDef[] {
  return [
    withNav(table('deploymentpipeline', 'deploymentpipelines', 'Bereitstellungspipeline', 'deploymentpipelineid', 'name', [
      A('name', 'String', 'Name'), A('description', 'Memo', 'Beschreibung'), LK('ownerid', 'Besitzer', 'systemuser')
    ], pl), { deploymentpipeline_deploymentenvironment: 'deploymentenvironment' }),
    table('deploymentstage', 'deploymentstages', 'Bereitstellungsphase', 'deploymentstageid', 'name', [
      A('name', 'String', 'Name'), LK('deploymentpipelineid', 'Pipeline', 'deploymentpipeline'),
      LK('targetdeploymentenvironmentid', 'Zielumgebung', 'deploymentenvironment'), LK('previousdeploymentstageid', 'Vorherige Phase', 'deploymentstage')
    ], st),
    table('deploymentstagerun', 'deploymentstageruns', 'Ausführung der Bereitstellungsphase', 'deploymentstagerunid', 'artifactname', [
      A('artifactname', 'String', 'Artefaktname'), A('artifactversion', 'String', 'Artefaktversion'), LK('deploymentpipelineid', 'Pipeline', 'deploymentpipeline'),
      LK('deploymentstageid', 'Phase', 'deploymentstage'), A('stagerunstatus', 'Picklist', 'Status', undefined, RUN_STATUS), LK('createdby', 'Erstellt von', 'systemuser'),
      A('starttime', 'DateTime', 'Startzeit'), A('endtime', 'DateTime', 'Endzeit'), LK('devdeploymentenvironment', 'Entwicklungsumgebung', 'deploymentenvironment')
    ], runs),
    table('deploymentenvironment', 'deploymentenvironments', 'Bereitstellungsumgebung', 'deploymentenvironmentid', 'name', [
      A('name', 'String', 'Name'), A('environmenttype', 'Picklist', 'Umgebungstyp', undefined, ENV_ROLE),
      A('validationstatus', 'Picklist', 'Validierungsstatus', undefined, VALIDATION), A('environmentid', 'String', 'Umgebungs-ID'), A('errormessage', 'String', 'Fehlermeldung')
    ], envs)
  ];
}
const ENV_LK = (name: string): IAttr => LK(name, 'Environment', 'admin_environment');
const MAKER_LK = (name: string, label: string): IAttr => LK(name, label, 'admin_maker');

export const DATABASES: IDatabase[] = [
  db(COE_HOST, [
    table('admin_app', 'admin_apps', 'PowerApps App', 'admin_appid', 'admin_displayname', [
      A('admin_displayname', 'String', 'App Display Name'), MAKER_LK('admin_appowner', 'App Owner'), ENV_LK('admin_appenvironment'),
      A('admin_appenvironmentid', 'String', 'App Environment Id'), A('admin_powerappstype', 'Picklist', 'App Type', undefined, APP_TYPES),
      A('cr5d5_appisorphaned', 'String', 'App is Orphaned'), A('admin_applastlaunchedon', 'DateTime', 'App Last Launched On'),
      A('admin_appmodifiedon', 'DateTime', 'App Modified On'), A('admin_appcreatedon', 'DateTime', 'App Created On'),
      A('admin_appdeleted', 'Boolean', 'App Deleted'), A('admin_appplanclassification', 'String', 'App Plan Classification'),
      A('admin_appsharedusers', 'Integer', 'App Shared Users'), A('admin_appsharedgroups', 'Integer', 'App Shared Groups'),
      A('admin_launchesinthepast30days', 'Integer', 'Launches in the past 30 days'),
      A('admin_appsharedwithtenant', 'Boolean', 'App Shared With Tenant'), A('admin_appsharedguestuser', 'Boolean', 'App Shared With Guest User'),
      A('admin_dlppoliciesviolated', 'String', 'DLP Policies Violated'),
      A('admin_appisquarantined', 'Boolean', 'App Is Quarantined'), A('admin_quarantineapp', 'Boolean', 'Quarantine App'),
      A('admin_issolutionapp', 'Boolean', 'Is Solution App'), A('admin_appconnections', 'Memo', 'App Connections')
    ], apps),
    table('admin_flow', 'admin_flows', 'Flow', 'admin_flowid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Display Name'), MAKER_LK('admin_flowcreator', 'Maker'), MAKER_LK('admin_derivedowner', 'Derived Owner'),
      ENV_LK('admin_flowenvironment'), A('admin_flowenvironmentid', 'String', 'Flow Environment Id'), A('admin_flowidstring', 'String', 'Flow ID String'),
      A('admin_flowstate', 'String', 'Flow State'), A('cr5d5_flowisorphaned', 'String', 'Flow is Orphaned'), A('cr5d5_flowtrigger', 'String', 'Trigger'),
      A('admin_flowmodifiedon', 'DateTime', 'Flow Modified On'), A('admin_flowcreatedon', 'DateTime', 'Flow Created On'), A('admin_flowdeleted', 'Boolean', 'Flow Deleted'), A('admin_flowispublished', 'Boolean', 'Flow Is Published'),
      A('admin_flowsharedguestuser', 'Boolean', 'Flow Shared With Guest User'),
      A('admin_issolutionflow', 'Boolean', 'Is Solution Flow'), A('admin_flowconnections', 'Memo', 'Flow Connections')
    ], flows),
    table('admin_environment', 'admin_environments', 'Environment', 'admin_environmentid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Display Name'), A('admin_environmentname', 'String', 'Name'), A('admin_environmentsku', 'String', 'Environment Sku'),
      MAKER_LK('admin_maker', 'Environment Maker'), A('admin_environmentmakerdisplayname', 'String', 'Environment Maker'),
      A('admin_governanceprotectionlevel', 'String', 'Governance Protection Level'), A('admin_environmentruntimestate', 'String', 'Environment Runtime State'),
      A('admin_environmentstate', 'String', 'Environment State'), A('admin_region', 'String', 'Environment Region'),
      A('admin_environmentcdsinstanceurl', 'String', 'Environment CDS Instance URL'), A('admin_environmentcreatedon', 'DateTime', 'Environment Created On'),
      A('admin_environmentdeleted', 'Boolean', 'Has the environment been deleted?'), A('admin_excusefrominventory', 'Boolean', 'Excuse From Inventory'),
      A('admin_securitygroupname', 'String', 'Security Group Name'), A('admin_numberofapps', 'Integer', 'Number of Apps'), A('admin_numberofflows', 'Integer', 'Number of Flows'),
      A('admin_hasdlppolicy', 'Boolean', 'Has DLP Policy'), A('admin_dlppolicies', 'String', 'DLP Policies')
    ], environments),
    table('admin_connector', 'admin_connectors', 'PowerApps Connector', 'admin_connectorid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Connector Display Name'), A('admin_name', 'String', 'Name'), A('admin_tier', 'String', 'Tier'),
      A('admin_publisher', 'String', 'Publisher'), A('admin_iscustomapi', 'Boolean', 'Is Custom API'), A('admin_connectordeleted', 'Boolean', 'Connector Deleted'),
      A('admin_numapps', 'Integer', 'Number of Apps'), A('admin_numflows', 'Integer', 'Number of Flows'),
      A('admin_customconnectorhost', 'String', 'Custom Connector Host'), A('admin_customconnectorhostishttp', 'Boolean', 'Custom Connector Host is HTTP')
    ], connectorRows),
    table('admin_connectionreference', 'admin_connectionreferences', 'Connection Reference', 'admin_connectionreferenceid', 'admin_name', [
      A('admin_name', 'String', 'Name'), LK('admin_app', 'App', 'admin_app'), LK('admin_flow', 'Flow', 'admin_flow'), LK('admin_connector', 'Connector', 'admin_connector'),
      A('admin_flow_connectionreferencesource', 'String', 'Source')
    ], connectionRefs),
    table('admin_flowactiondetail', 'admin_flowactiondetails', 'Flow Action Detail', 'admin_flowactiondetailid', 'admin_name', [
      A('admin_name', 'String', 'Name'), LK('admin_flow', 'Flow', 'admin_flow'), A('admin_actiontype', 'String', 'Action Type'),
      A('admin_operation', 'String', 'Operation'), A('cr5d5_istypetrigger', 'Boolean', 'Is Type Trigger')
    ], flowActions),
    table('admin_auditlog', 'admin_auditlogs', 'Audit Log', 'admin_auditlogid', 'admin_title', [
      A('admin_title', 'String', 'Title'), LK('admin_applookup', 'App', 'admin_app'), A('admin_appenvironment', 'String', 'App Environment'),
      A('admin_operation', 'String', 'Operation'), A('admin_workload', 'String', 'Workload'), A('admin_userupn', 'String', 'User UPN'),
      A('admin_creationtime', 'DateTime', 'Creation Time')
    ], auditLog),
    table('admin_pva', 'admin_pvas', 'PVA', 'admin_pvaid', 'admin_name', [
      A('admin_name', 'String', 'Name'), MAKER_LK('admin_pvaowner', 'Bot Owner'), ENV_LK('admin_pvaenvironment'), A('admin_pvastate', 'String', 'State'),
      A('admin_pvanumberlaunches', 'Integer', 'Conversations (Last Month)'), A('admin_pvalastlaunchedon', 'DateTime', 'Last Launched On'),
      A('admin_bottemplate', 'String', 'Bot Template'), A('admin_pvadeleted', 'Boolean', 'Deleted')
    ], bots),
    table('admin_rpa', 'admin_rpas', 'RPA', 'admin_rpaid', 'admin_name', [
      A('admin_name', 'String', 'Name'), A('admin_displayname', 'String', 'Display Name'), MAKER_LK('admin_rpaowner', 'Owner'),
      ENV_LK('admin_desktopflowenvironment'), A('admin_rpastate', 'String', 'State'), A('admin_rpamodifiedon', 'DateTime', 'Modified On')
    ], rpas),
    table('admin_portal', 'admin_portals', 'Power Pages Site', 'admin_portalid', 'admin_portaldisplayname', [
      A('admin_portaldisplayname', 'String', 'Display Name'), MAKER_LK('admin_portalowner', 'Owner'), ENV_LK('admin_portalenvironment'),
      A('admin_portalwebsitename', 'String', 'Website Name'), A('admin_portalwebsiterecordstatus', 'String', 'Website Record Status'),
      A('admin_authtablepermissionsenabled', 'Boolean', 'Table Permissions Enabled'), A('admin_portalmodifiedon', 'DateTime', 'Modified On')
    ], portals),
    table('admin_businessprocessflows', 'admin_businessprocessflowses', 'Business Process Flow', 'admin_businessprocessflowsid', 'admin_name', [
      A('admin_name', 'String', 'Name'), A('admin_businessprocessflowdisplayname', 'String', 'Display Name'), MAKER_LK('admin_businessprocessflowowner', 'Owner'),
      ENV_LK('admin_businessprocessflowenvironment'), A('admin_businessprocessflowprimaryentity', 'String', 'Primary Entity'),
      A('admin_businessprocessflowstatus', 'String', 'Status'), A('admin_businessprocessflowmodifiedon', 'DateTime', 'Modified On'),
      A('admin_businessprocessflowdeleted', 'Boolean', 'Deleted')
    ], bpfs),
    withNav(table('admin_solution', 'admin_solutions', 'Solution', 'admin_solutionid', 'admin_name', [
      A('admin_name', 'String', 'Solution Unique Name'), ENV_LK('admin_solutionenvironment'), MAKER_LK('admin_solutioncreator', 'Creator'),
      A('admin_solutionpublisher', 'String', 'Publisher'), A('admin_solutionismanaged', 'Boolean', 'Is Managed'), A('admin_solutiondeleted', 'Boolean', 'Deleted'),
      A('admin_solutionenvtobjectguidstring', 'String', 'Solution Envt Object Guid String'), A('admin_solutionenvtguid', 'String', 'Solution Envt Guid')
    ], coeSolutions), { admin_PPSolution_Apps: 'admin_app', admin_PPSolution_Flow: 'admin_flow' }),
    table('admin_aibuildermodel', 'admin_aibuildermodels', 'AI Model', 'admin_aibuildermodelid', 'admin_name', [
      A('admin_name', 'String', 'Name'), MAKER_LK('admin_aicreator', 'Creator'), ENV_LK('admin_aienvironment'), A('admin_aitemplate', 'String', 'Template'),
      A('admin_aistatus', 'String', 'Status'), A('admin_ailastused', 'DateTime', 'Last Used'), A('admin_aideleted', 'Boolean', 'Deleted')
    ], aiModels),
    table('admin_maker', 'admin_makers', 'Maker', 'admin_makerid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Display Name'), A('admin_userprincipalname', 'String', 'User Principal Name'), A('admin_useremail', 'String', 'User Email'),
      A('admin_department', 'String', 'Department'), A('admin_company', 'String', 'Company'), A('admin_country', 'String', 'Country'), A('admin_city', 'String', 'City'),
      A('admin_jobtitle', 'String', 'Job Title'), A('admin_userisguest', 'Boolean', 'User Is Guest'), A('admin_accountenabled', 'Boolean', 'Account Enabled'),
      A('admin_numberofapps', 'Integer', 'Number of Apps'), A('admin_numberofflows', 'Integer', 'Number of Flows'), A('admin_numberofenvironments', 'Integer', 'Number of Environments')
    ], makers),
    table('admin_powerplatformuser', 'admin_powerplatformusers', 'Power Platform User', 'admin_powerplatformuserid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Display Name'), A('admin_type', 'String', 'Type'), A('admin_userprincipalname', 'String', 'User Principal Name'),
      A('admin_useremail', 'String', 'User Email'), A('admin_userisguest', 'Boolean', 'User Is Guest'), A('admin_groupsize', 'Integer', 'Group Size')
    ], ppUsers),
    table('admin_powerplatformuserrole', 'admin_powerplatformuserroles', 'Power Platform User Role', 'admin_powerplatformuserroleid', 'admin_name', [
      A('admin_name', 'String', 'Name'), LK('admin_powerplatformuser', 'Power Platform User', 'admin_powerplatformuser'), LK('admin_app', 'App', 'admin_app'),
      LK('admin_flow', 'Flow', 'admin_flow'), ENV_LK('admin_environment'), A('admin_resourcetype', 'String', 'Resource Type'), A('admin_rolename', 'String', 'Role Name')
    ], userRoles),
    table('admin_environmentsecurityrolepermission', 'admin_environmentsecurityrolepermissions', 'Security Role Permission', 'admin_environmentsecurityrolepermissionid', 'admin_name', [
      A('admin_name', 'String', 'User Name'), A('admin_securityrolename', 'String', 'Security Role Name'), A('admin_securityroletemplateid', 'String', 'Security Role Template ID'), ENV_LK('admin_environment'),
      A('admin_assignedvia', 'String', 'Assigned Via'), A('admin_groupname', 'String', 'Group Name'), A('admin_businessunitname', 'String', 'Business Unit Name'),
      A('admin_usertype', 'Picklist', 'User Type', undefined, { 597910000: 'Member', 597910001: 'Guest', 597910002: 'App User', 597910003: 'Unknown', 597910004: 'Microsoft Account' }),
      A('admin_useridentityprovider', 'String', 'User Identity Provider')
    ], securityRoles),
    table('admin_connectionreferenceidentity', 'admin_connectionreferenceidentities', 'Connection Identity', 'admin_connectionreferenceidentityid', 'admin_name', [
      A('admin_name', 'String', 'Name'), A('admin_accountname', 'String', 'Account Name'), ENV_LK('admin_environment'), LK('admin_connector', 'Connector', 'admin_connector'),
      LK('admin_connectionreferencecreator', 'Creator', 'admin_powerplatformuser'), A('admin_noneorcrosstenantidentity', 'Boolean', 'None Or Cross Tenant Identity')
    ], connIdentities),
    table('admin_environmentcapacity', 'admin_environmentcapacities', 'Environment Capacity', 'admin_environmentcapacityid', 'admin_capacitytype', [
      A('admin_capacitytype', 'String', 'Capacity Type'), ENV_LK('admin_environment'), A('admin_actualconsumption', 'Decimal', 'Actual Consumption'),
      A('admin_ratedconsumption', 'Decimal', 'Rated Consumption'), A('admin_approvedcapacity', 'Decimal', 'Approved Capacity'),
      A('admin_percentusedofapproved', 'Decimal', 'Percent Used'), A('admin_capacityunit', 'String', 'Capacity Unit'), A('admin_environmentdeleted', 'Boolean', 'Environment Deleted')
    ], capacity),
    withNav(table('admin_dlppolicy', 'admin_dlppolicies', 'DLP Policy', 'admin_dlppolicyid', 'admin_name', [
      A('admin_name', 'String', 'Display Name'), A('admin_scope', 'Picklist', 'Environment Scope', undefined, DLP_SCOPES), A('admin_description', 'String', 'Description'),
      A('admin_isshared', 'Boolean', 'Is Shared'), A('admin_policyid', 'String', 'Policy ID')
    ], dlpPolicies), { admin_policy_environments: 'admin_environment', admin_blocked_policy_connectors: 'admin_connector', admin_confidential_policy_connectors: 'admin_connector' }),
    table('admin_datapolicydrafts', 'admin_datapolicydraftses', 'Data Policy Drafts', 'admin_datapolicydraftsid', 'admin_name', [
      A('admin_name', 'String', 'Name'), A('admin_displayname', 'String', 'DisplayName'), A('admin_apiversion', 'String', 'API Version'),
      A('admin_defaultconnectorclassification', 'String', 'Default Connector Classification'), A('admin_scope', 'String', 'Scope'),
      A('admin_environments', 'Memo', 'Environments'), A('admin_connectorgroups', 'Memo', 'Connector Groups'),
      LK('createdby', 'Erstellt von', 'systemuser'), LK('modifiedby', 'Geändert von', 'systemuser')
    ], dlpDrafts),
    table('admin_dlpimpactanalysis', 'admin_dlpimpactanalysises', 'DLP Impact Analysis', 'admin_dlpimpactanalysisid', 'admin_name', [
      A('admin_name', 'String', 'Name'), A('admin_dlppolicyname', 'String', 'Data Policy Name'), A('admin_datapolicyid', 'String', 'Data Policy ID'),
      LK('admin_impactedapp', 'Impacted App', 'admin_app'), LK('admin_impactedflow', 'Impacted Flow', 'admin_flow'), ENV_LK('admin_impactedenvironment'),
      MAKER_LK('admin_impactowner', 'Impact Owner'), A('admin_conflictingconnectorbusiness', 'Memo', 'Conflicting connector (business)'),
      A('admin_conflictingconnectornonbusiness', 'Memo', 'Conflicting connector (non-business)'), A('admin_conflictingconnectorblocked', 'Memo', 'Conflicting connector (blocked)'),
      A('admin_decision', 'Picklist', 'Decision', undefined, { 597910000: 'In progress', 597910001: 'Not required - Maker will delete', 597910002: 'Not required - Can be deleted by IT', 597910003: 'Required - Needs to be migrated', 597910004: 'Required - Maker will update app' }),
      A('admin_decisionstatus', 'Picklist', 'Decision status', undefined, { 597910000: 'In progress', 597910001: 'Completed' }),
      A('admin_actionrequiredon', 'DateTime', 'Action Required On')
    ], dlpTasks),
    // das CoE-Environment ist zugleich ein zweiter Pipelines-Host
    ...pipelineTables(pipelines2, stages2, stageRuns2, depEnvs2)
  ]),
  db(DEFAULT_HOST, [
    table('solution', 'solutions', 'Lösung', 'solutionid', 'friendlyname', [
      A('friendlyname', 'String', 'Anzeigename'), A('uniquename', 'String', 'Name'), A('version', 'String', 'Version'), A('ismanaged', 'Boolean', 'Paket-Typ'),
      A('isvisible', 'Boolean', 'Ist sichtbar'), LK('publisherid', 'Herausgeber', 'publisher'), A('installedon', 'DateTime', 'Installiert am'), A('description', 'Memo', 'Beschreibung')
    ], solutions),
    table('publisher', 'publishers', 'Herausgeber', 'publisherid', 'friendlyname', [
      A('friendlyname', 'String', 'Anzeigename'), A('uniquename', 'String', 'Name')
    ], publishers)
  ]),
  db(PIPELINES_HOST, [
    ...pipelineTables(pipelines, stages, stageRuns, depEnvs),
    table('systemuser', 'systemusers', 'Benutzer', 'systemuserid', 'fullname', [A('fullname', 'String', 'Vollständiger Name')], sysUsers)
  ])
];

// ---------- Inhalt der Solutions (CoE-N:N + Systemtabellen im jeweiligen Environment) ----------
// Eigener Zufallsgenerator, damit die übrigen Beispieldaten unverändert bleiben.
(function seedSolutionContents(): void {
  let s2 = 4711;
  const r2 = (): number => { s2 = (s2 * 1103515245 + 12345) & 0x7fffffff; return s2 / 0x7fffffff; };
  const g2 = (): string => {
    const h = (n: number): string => Array.from({ length: n }, () => Math.floor(r2() * 16).toString(16)).join('');
    return `${h(8)}-${h(4)}-4${h(3)}-a${h(3)}-${h(12)}`;
  };
  const pick2 = <T>(a: T[]): T => a[Math.floor(r2() * a.length)];
  const envById: { [id: string]: Record<string, unknown> } = {};
  environments.forEach(e => { envById[String(e.admin_environmentid)] = e; });
  const BOT_TYPE = 10187;
  const SITE_TYPE = 10254;
  const envDb = (host: string): IDatabase => {
    let d = DATABASES.filter(x => x.host.toLowerCase() === host.toLowerCase())[0];
    if (!d) {
      d = { host, tables: {} };
      DATABASES.push(d);
    }
    if (!d.tables.solutioncomponent) {
      const add = (t: ITableDef): void => { d!.tables[t.logicalName] = t; };
      add(table('solutioncomponent', 'solutioncomponents', 'Lösungskomponente', 'solutioncomponentid', 'objectid', [
        LK('solutionid', 'Lösung', 'solution'), A('componenttype', 'Integer', 'Objekttyp'), A('objectid', 'Uniqueidentifier', 'Objekt-ID')
      ], []));
      add(table('solutioncomponentdefinition', 'solutioncomponentdefinitions', 'Lösungskomponentendefinition', 'solutioncomponentdefinitionid', 'primaryentityname', [
        A('solutioncomponenttype', 'Integer', 'Komponententyp'), A('primaryentityname', 'String', 'Primäre Entität')
      ], [
        { solutioncomponentdefinitionid: g2(), solutioncomponenttype: BOT_TYPE, primaryentityname: 'bot' },
        { solutioncomponentdefinitionid: g2(), solutioncomponenttype: BOT_TYPE + 1, primaryentityname: 'botcomponent' },
        { solutioncomponentdefinitionid: g2(), solutioncomponenttype: SITE_TYPE, primaryentityname: 'powerpagesite' }
      ]));
      add(table('workflow', 'workflows', 'Prozess', 'workflowid', 'name', [
        A('name', 'String', 'Name'), A('category', 'Integer', 'Kategorie'), A('statecode', 'Integer', 'Status')
      ], []));
      add(table('canvasapp', 'canvasapps', 'Canvas-App', 'canvasappid', 'displayname', [
        A('displayname', 'String', 'Anzeigename'), A('name', 'String', 'Name')
      ], []));
      add(table('appmodule', 'appmodules', 'Modellgesteuerte App', 'appmoduleid', 'name', [A('name', 'String', 'Name')], []));
      add(table('bot', 'bots', 'Copilot', 'botid', 'name', [A('name', 'String', 'Name'), A('statecode', 'Integer', 'Status')], []));
      add(table('powerpagesite', 'powerpagesites', 'Power Pages-Site', 'powerpagesiteid', 'name', [A('name', 'String', 'Name')], []));
    }
    return d;
  };
  const appsByEnv: { [id: string]: Record<string, unknown>[] } = {};
  apps.forEach(a => { (appsByEnv[String(a._admin_appenvironment_value)] = appsByEnv[String(a._admin_appenvironment_value)] || []).push(a); });
  const flowsByEnv: { [id: string]: Record<string, unknown>[] } = {};
  flows.forEach(f => { (flowsByEnv[String(f._admin_flowenvironment_value)] = flowsByEnv[String(f._admin_flowenvironment_value)] || []).push(f); });

  coeSolutions.forEach((sol, i) => {
    const env = envById[String(sol._admin_solutionenvironment_value)];
    sol.admin_solutionenvtobjectguidstring = sol.admin_solutionid;
    sol.admin_solutionenvtguid = env ? env.admin_environmentname : null;
    const host = env && env.admin_environmentcdsinstanceurl ? String(env.admin_environmentcdsinstanceurl).replace(/^https:\/\//, '').replace(/\/$/, '') : '';
    const managed = !!sol.admin_solutionismanaged;
    const ea = appsByEnv[String(sol._admin_solutionenvironment_value)] || [];
    const ef = flowsByEnv[String(sol._admin_solutionenvironment_value)] || [];
    // Jede vierte Solution ohne CoE-Zuordnung (dann zeigt das Cockpit Apps/Flows live aus dem Environment)
    const coeLinked = i % 4 !== 1;
    const nApps = managed ? 0 : Math.min(ea.length, Math.floor(r2() * 4));
    const nFlows = managed ? Math.min(ef.length, Math.floor(r2() * 2)) : Math.min(ef.length, 1 + Math.floor(r2() * 5));
    const solApps = Array.from({ length: nApps }, () => pick2(ea)).filter((x, k, arr) => arr.indexOf(x) === k);
    const solFlows = Array.from({ length: nFlows }, () => pick2(ef)).filter((x, k, arr) => arr.indexOf(x) === k);
    sol.__nav_admin_PPSolution_Apps = coeLinked ? solApps.map(a => a.admin_appid) : [];
    sol.__nav_admin_PPSolution_Flow = coeLinked ? solFlows.map(f => f.admin_flowid) : [];
    if (!host) {
      return;
    }
    const d = envDb(host);
    const comp = (type: number, objectid: unknown): void => {
      d.tables.solutioncomponent.rows.push({ solutioncomponentid: g2(), _solutionid_value: sol.admin_solutionid, componenttype: type, objectid });
    };
    solApps.forEach(a => {
      d.tables.canvasapp.rows.push({ canvasappid: a.admin_appid, displayname: a.admin_displayname, name: `cr5d5_app_${String(a.admin_appid).slice(0, 5)}` });
      comp(300, a.admin_appid);
    });
    solFlows.forEach(f => {
      const wid = g2();
      d.tables.workflow.rows.push({ workflowid: wid, name: f.admin_displayname, category: 5, statecode: f.admin_flowstate === 'Started' ? 1 : 0 });
      comp(29, wid);
    });
    if (!managed) {
      if (r2() < 0.35) {
        const id = g2();
        d.tables.appmodule.rows.push({ appmoduleid: id, name: `${pick2(['Antrags', 'Service', 'Vertriebs'])}verwaltung` });
        comp(80, id);
      }
      for (let k = 0; k < Math.floor(r2() * 3); k++) {
        const id = g2();
        d.tables.workflow.rows.push({ workflowid: id, name: `${pick2(['SAP-Export', 'Rechnungsimport', 'Excel-Abgleich', 'Belegablage'])} (Desktop)`, category: 6, statecode: 1 });
        comp(29, id);
      }
      for (let k = 0; k < Math.floor(r2() * 3); k++) {
        const id = g2();
        d.tables.bot.rows.push({ botid: id, name: `${pick2(['HR', 'IT', 'Einkauf', 'Service'])}-Copilot`, statecode: 0 });
        comp(BOT_TYPE, id);
        // Themen (botcomponent) sind keine eigenen Einträge
        comp(BOT_TYPE + 1, g2());
      }
      if (r2() < 0.3) {
        const id = g2();
        d.tables.powerpagesite.rows.push({ powerpagesiteid: id, name: `${pick2(['Lieferanten', 'Bewerber', 'Kunden'])}portal` });
        comp(SITE_TYPE, id);
      }
      // Tabellen, Spalten usw. – ignoriert das Cockpit
      comp(1, g2());
      comp(2, g2());
    }
  });
})();

// ---------- Für die simulierte Graph-API (Lizenzen, Gruppenmitglieder) ----------
export const GRAPH_PEOPLE = humans.map(h => ({ id: String(h.admin_makerid), upn: String(h.admin_userprincipalname || ''), name: String(h.admin_displayname || '') }));
export const GRAPH_GROUPS = ppUsers.filter(u => u.admin_type === 'Group').map(g => ({ id: String(g.admin_powerplatformuserid), name: String(g.admin_displayname) }));

/** Grundlage der simulierten DLP-Schnittstelle (api.bap.microsoft.com). */
export const DLP_MOCK = {
  policies: dlpPolicies,
  connectors: connectors.map(c => ({ coeId: String(c.admin_connectorid), name: String(c.admin_name), displayName: String(c.admin_displayname), custom: c.admin_iscustomapi === true })),
  environments: environments.map(e => ({ coeId: String(e.admin_environmentid), name: String(e.admin_environmentname) }))
};
