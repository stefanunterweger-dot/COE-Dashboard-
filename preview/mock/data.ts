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
    admin_appsharedgroups: int(0, 6)
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
    const n = int(0, 6);
    for (let i = 0; i < n; i++) {
      const u = pick(humans);
      out.push({ ...COMMON(), admin_auditlogid: guid(), admin_title: `LaunchPowerApp – ${a.admin_displayname}`, _admin_applookup_value: a.admin_appid, admin_appenvironment: envById[String(a._admin_appenvironment_value)].admin_displayname, admin_operation: 'LaunchPowerApp', admin_workload: 'PowerApps', admin_userupn: u.admin_userprincipalname, admin_creationtime: daysAgo(recentish(120)) });
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
const securityRoles = Array.from({ length: 420 }, () => {
  const env = environments[int(0, environments.length - 1)]; const u = pick(humans); const viaGroup = chance(0.3);
  return { ...COMMON(), admin_environmentsecurityrolepermissionid: guid(), admin_name: u.admin_displayname, admin_securityrolename: pick(['System Administrator', 'System Customizer', 'Environment Maker', 'Basic User', 'Power Platform Admin SR']), _admin_environment_value: env.admin_environmentid, admin_assignedvia: viaGroup ? 'Team' : 'Direkt', admin_groupname: viaGroup ? pick(ppPrincipals).admin_displayname : null, admin_businessunitname: `org-${String(env.admin_displayname).toLowerCase().replace(/[^a-z]/g, '').slice(0, 10)}`, admin_usertype: u.admin_userisguest ? 597910001 : 597910000, admin_useridentityprovider: 'Entra ID' };
});
const connIdentities = Array.from({ length: 260 }, () => {
  const env = envPick(); const c = pick(connectors); const u = pick(ppUsers.filter(p => p.admin_type === 'User')); const cross = chance(0.05);
  return { ...COMMON(), admin_connectionreferenceidentityid: guid(), admin_name: `${c.admin_name}-${String(u.admin_displayname).split(' ')[1]}`, admin_accountname: cross ? 'extern@fabrikam.com' : u.admin_userprincipalname, _admin_environment_value: env.admin_environmentid, _admin_connector_value: c.admin_connectorid, _admin_connectionreferencecreator_value: u.admin_powerplatformuserid, admin_noneorcrosstenantidentity: cross };
});
const capacity: Record<string, unknown>[] = [];
environments.filter(e => !e.admin_environmentdeleted && e.admin_environmentcdsinstanceurl).forEach(e => {
  [['Database', 1.0], ['File', 2.5], ['Log', 0.3]].forEach(([type, scale]) => {
    const used = Math.round(rnd() * 4000 * (scale as number)) / 100;
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

// ---------- Tabellen-Definitionen ----------
const A = (name: string, type: string, label: string, target?: string, options?: { [v: number]: string }): IAttr => ({ name, type, label, target, options });
const STD = [A('statecode', 'State', 'Status'), A('createdon', 'DateTime', 'Erstellt am'), A('modifiedon', 'DateTime', 'Geändert am')];

function table(logicalName: string, entitySet: string, displayName: string, primaryId: string, primaryName: string,
  attrs: IAttr[], rows: Record<string, unknown>[] | (() => Record<string, unknown>[])): ITableDef {
  const def = { logicalName, entitySet, displayName, primaryId, primaryName, attrs: [A(primaryId, 'Uniqueidentifier', displayName)].concat(attrs, STD) } as ITableDef;
  Object.defineProperty(def, 'rows', { get: typeof rows === 'function' ? rows : () => rows, enumerable: true });
  return def;
}
function db(host: string, defs: ITableDef[]): IDatabase {
  const tables: { [n: string]: ITableDef } = {};
  defs.forEach(d => { tables[d.logicalName] = d; });
  return { host, tables };
}
const LK = (name: string, label: string, target: string): IAttr => A(name, 'Lookup', label, target);
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
      A('admin_appsharedusers', 'Integer', 'App Shared Users'), A('admin_appsharedgroups', 'Integer', 'App Shared Groups')
    ], apps),
    table('admin_flow', 'admin_flows', 'Flow', 'admin_flowid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Display Name'), MAKER_LK('admin_flowcreator', 'Maker'), MAKER_LK('admin_derivedowner', 'Derived Owner'),
      ENV_LK('admin_flowenvironment'), A('admin_flowenvironmentid', 'String', 'Flow Environment Id'), A('admin_flowidstring', 'String', 'Flow ID String'),
      A('admin_flowstate', 'String', 'Flow State'), A('cr5d5_flowisorphaned', 'String', 'Flow is Orphaned'), A('cr5d5_flowtrigger', 'String', 'Trigger'),
      A('admin_flowmodifiedon', 'DateTime', 'Flow Modified On'), A('admin_flowdeleted', 'Boolean', 'Flow Deleted'), A('admin_flowispublished', 'Boolean', 'Flow Is Published')
    ], flows),
    table('admin_environment', 'admin_environments', 'Environment', 'admin_environmentid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Display Name'), A('admin_environmentname', 'String', 'Name'), A('admin_environmentsku', 'String', 'Environment Sku'),
      MAKER_LK('admin_maker', 'Environment Maker'), A('admin_environmentmakerdisplayname', 'String', 'Environment Maker'),
      A('admin_governanceprotectionlevel', 'String', 'Governance Protection Level'), A('admin_environmentruntimestate', 'String', 'Environment Runtime State'),
      A('admin_environmentstate', 'String', 'Environment State'), A('admin_region', 'String', 'Environment Region'),
      A('admin_environmentcdsinstanceurl', 'String', 'Environment CDS Instance URL'), A('admin_environmentcreatedon', 'DateTime', 'Environment Created On'),
      A('admin_environmentdeleted', 'Boolean', 'Has the environment been deleted?'), A('admin_excusefrominventory', 'Boolean', 'Excuse From Inventory'),
      A('admin_securitygroupname', 'String', 'Security Group Name'), A('admin_numberofapps', 'Integer', 'Number of Apps'), A('admin_numberofflows', 'Integer', 'Number of Flows')
    ], environments),
    table('admin_connector', 'admin_connectors', 'PowerApps Connector', 'admin_connectorid', 'admin_displayname', [
      A('admin_displayname', 'String', 'Connector Display Name'), A('admin_name', 'String', 'Name'), A('admin_tier', 'String', 'Tier'),
      A('admin_publisher', 'String', 'Publisher'), A('admin_iscustomapi', 'Boolean', 'Is Custom API'), A('admin_connectordeleted', 'Boolean', 'Connector Deleted'),
      A('admin_numapps', 'Integer', 'Number of Apps'), A('admin_numflows', 'Integer', 'Number of Flows')
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
    table('admin_solution', 'admin_solutions', 'Solution', 'admin_solutionid', 'admin_name', [
      A('admin_name', 'String', 'Solution Unique Name'), ENV_LK('admin_solutionenvironment'), MAKER_LK('admin_solutioncreator', 'Creator'),
      A('admin_solutionpublisher', 'String', 'Publisher'), A('admin_solutionismanaged', 'Boolean', 'Is Managed'), A('admin_solutiondeleted', 'Boolean', 'Deleted')
    ], coeSolutions),
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
      A('admin_name', 'String', 'User Name'), A('admin_securityrolename', 'String', 'Security Role Name'), ENV_LK('admin_environment'),
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
    ], capacity)
  ]),
  db(DEFAULT_HOST, [
    table('solution', 'solutions', 'Lösung', 'solutionid', 'friendlyname', [
      A('friendlyname', 'String', 'Anzeigename'), A('uniquename', 'String', 'Name'), A('version', 'String', 'Version'), A('ismanaged', 'Boolean', 'Paket-Typ'),
      A('isvisible', 'Boolean', 'Ist sichtbar'), LK('publisherid', 'Herausgeber', 'publisher'), A('installedon', 'DateTime', 'Installiert am'), A('description', 'Memo', 'Beschreibung')
    ], solutions),
    table('publisher', 'publishers', 'Herausgeber', 'publisherid', 'friendlyname', [
      A('friendlyname', 'String', 'Anzeigename'), A('uniquename', 'String', 'Name')
    ], publishers)
  ])
];
