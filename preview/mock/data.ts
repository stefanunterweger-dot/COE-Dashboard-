/* Deterministische Beispieldaten im Format des CoE Starter Kits. */

export interface IAttr {
  name: string;          // logischer Name (bei Lookups ohne _..._value)
  type: string;          // Dataverse AttributeType
  label: string;
  target?: string;       // Lookup-Ziel (logischer Tabellenname)
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
/** Datum vor x Tagen (mit Uhrzeit), als ISO-String wie von Dataverse. */
function daysAgo(d: number): string {
  return new Date(NOW - d * DAY - int(0, 86000) * 1000).toISOString().replace(/\.\d{3}Z$/, 'Z');
}
/** Viele aktuelle, einige alte Werte. */
function recentish(maxDays: number): number {
  const r = rnd();
  if (r < 0.12) { return int(0, 1); }
  if (r < 0.35) { return int(2, 7); }
  if (r < 0.6) { return int(8, 30); }
  if (r < 0.8) { return int(31, 90); }
  return int(91, maxDays);
}

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

// ---------- Maker ----------
const makers: Record<string, unknown>[] = [];
const usedNames: { [n: string]: boolean } = {};
while (makers.length < 64) {
  const f = pick(FIRST);
  const l = pick(LAST);
  if (usedNames[f + l]) { continue; }
  usedNames[f + l] = true;
  const country = pick(COUNTRIES);
  makers.push({
    admin_makerid: guid(),
    admin_displayname: `${f} ${l}`,
    admin_userprincipalname: `${umlaut(f)}.${umlaut(l)}@contoso.com`,
    admin_department: pick(DEPARTMENTS),
    admin_company: country === 'Schweiz' ? 'Contoso AG' : 'Contoso GmbH',
    admin_country: country,
    admin_city: country === 'Österreich' ? pick(['Wien', 'Graz', 'Linz', 'Salzburg', 'Innsbruck']) : country === 'Deutschland' ? pick(['München', 'Berlin', 'Hamburg']) : 'Zürich',
    admin_jobtitle: pick(['Sachbearbeiter:in', 'Teamleiter:in', 'Analyst:in', 'Projektleiter:in', 'Entwickler:in', 'Assistenz']),
    admin_userisservicepricipal: false,
    createdon: daysAgo(int(30, 900)),
    modifiedon: daysAgo(recentish(300))
  });
}
// Ehemalige Mitarbeitende (für verwaiste Apps/Flows)
const departed = makers.slice(0, 6);

// ---------- Environments ----------
interface IEnvSeed { name: string; type: string; region: string; host?: string; weight: number; deleted?: boolean; id?: string }
const envSeeds: IEnvSeed[] = [
  { name: 'Contoso (default)', type: 'Default', region: 'europe', host: DEFAULT_HOST, weight: 55, id: TENANT_ID },
  { name: 'CoE', type: 'Production', region: 'europe', host: COE_HOST, weight: 2 },
  { name: 'HR Produktion', type: 'Production', region: 'europe', host: 'contoso-hr.crm4.dynamics.com', weight: 6 },
  { name: 'HR Test', type: 'Sandbox', region: 'europe', host: 'contoso-hr-test.crm4.dynamics.com', weight: 3 },
  { name: 'Finanzen Produktion', type: 'Production', region: 'europe', host: 'contoso-fin.crm4.dynamics.com', weight: 7 },
  { name: 'Finanzen Entwicklung', type: 'Sandbox', region: 'europe', host: 'contoso-fin-dev.crm4.dynamics.com', weight: 4 },
  { name: 'Vertrieb Produktion', type: 'Production', region: 'europe', host: 'contoso-sales.crm4.dynamics.com', weight: 6 },
  { name: 'Einkauf Sandbox', type: 'Sandbox', region: 'europe', host: 'contoso-proc.crm4.dynamics.com', weight: 3 },
  { name: 'Projektraum Nord', type: 'Teams', region: 'europe', host: 'org5b1e77c2.crm4.dynamics.com', weight: 3 },
  { name: 'Dev – Anna Berger', type: 'Developer', region: 'europe', host: 'org9a2c41f0.crm4.dynamics.com', weight: 2 },
  { name: 'Trial – Lukas Huber', type: 'Trial', region: 'unitedstates', host: 'org77d1e9a3.crm.dynamics.com', weight: 2 },
  { name: 'Marketing Kampagnen', type: 'Production', region: 'europe', host: 'contoso-mkt.crm4.dynamics.com', weight: 4 },
  { name: 'Altsystem Migration', type: 'Sandbox', region: 'europe', host: 'contoso-legacy.crm4.dynamics.com', weight: 0, deleted: true }
];
const REGION_LABEL: { [k: string]: string } = { europe: 'Europe', unitedstates: 'United States' };
const environments = envSeeds.map(e => ({
  admin_environmentid: e.id || guid(),
  admin_displayname: e.name,
  admin_environmenttype: e.type,
  admin_environmentregion: REGION_LABEL[e.region],
  admin_environmentcdsinstanceurl: e.host ? `https://${e.host}/` : null,
  admin_environmentcreatedon: daysAgo(e.type === 'Default' ? 1650 : int(40, 1200)),
  admin_environmentdeleted: !!e.deleted,
  admin_environmentsku: e.type === 'Default' ? 'Default' : e.type,
  admin_environmentstate: e.deleted ? 'Deleted' : 'Ready',
  createdon: daysAgo(int(200, 600)),
  modifiedon: daysAgo(recentish(60))
}));
const envWeighted: number[] = [];
envSeeds.forEach((e, i) => { for (let k = 0; k < e.weight; k++) { envWeighted.push(i); } });
const envLinkId = (i: number): string =>
  envSeeds[i].type === 'Default' ? `Default-${environments[i].admin_environmentid}` : String(environments[i].admin_environmentid);

// ---------- Apps ----------
const APP_WORDS_A = ['Urlaubs', 'Reise', 'Spesen', 'Inventar', 'Besucher', 'Schulungs', 'Projekt', 'Raum', 'Fuhrpark', 'Angebots',
  'Lieferanten', 'Qualitäts', 'Onboarding', 'Wartungs', 'Kunden', 'Rechnungs', 'Bestell', 'Zeit', 'Ideen', 'Störungs'];
const APP_WORDS_B = ['antrag', 'buchung', 'verwaltung', 'erfassung', 'tracker', 'portal', 'übersicht', 'checkliste', 'planer', 'meldung'];
const apps: Record<string, unknown>[] = [];
for (let i = 0; i < 6840; i++) {
  const envIdx = pick(envWeighted);
  const env = environments[envIdx];
  const orphan = chance(0.11);
  const owner = orphan ? pick(departed) : pick(makers);
  const type = envSeeds[envIdx].type === 'Default' ? (chance(0.95) ? 'Canvas' : 'Custom Page') : pick(['Canvas', 'Canvas', 'Model Driven', 'Custom Page']);
  const created = int(20, 1400);
  const modified = Math.min(created, recentish(700));
  const launched = chance(0.85) ? Math.min(recentish(400), created) : null;
  const name = `${pick(APP_WORDS_A)}${pick(APP_WORDS_B)}${chance(0.2) ? ' ' + pick(['v2', '(Test)', 'Neu', '2025', 'Mobil']) : ''}`;
  apps.push({
    admin_appid: guid(),
    admin_displayname: name.charAt(0).toUpperCase() + name.slice(1),
    _admin_appowner_value: owner.admin_makerid,
    _admin_appenvironment_value: env.admin_environmentid,
    admin_appenvironmentid: envLinkId(envIdx),
    admin_apptype: type,
    admin_appisorphaned: orphan ? 'Orphaned' : 'Not Orphaned',
    admin_applastlaunchedon: launched === null ? null : daysAgo(Math.max(0, launched)),
    admin_appmodifiedon: daysAgo(modified),
    admin_appcreatedon: daysAgo(created),
    admin_appdeleted: chance(0.04),
    admin_appplanclassification: chance(0.28) ? 'Premium' : 'Standard',
    admin_appshareduserscount: int(0, 140),
    admin_appsharedgroupscount: int(0, 6),
    admin_appdescription: chance(0.6) ? `Erstellt von der Abteilung ${String(owner.admin_department)}.` : null,
    admin_appisquarantined: false,
    createdon: daysAgo(created),
    modifiedon: daysAgo(modified)
  });
}

// ---------- Flows ----------
const FLOW_TEMPLATES = ['Genehmigung {x} an Vorgesetzte', 'Täglicher Bericht {x}', 'Benachrichtigung bei neuem {x}',
  '{x} aus E-Mail in SharePoint speichern', 'Erinnerung {x} fällig', 'Sync {x} → Excel', '{x}-Eingang an Teams posten',
  'Wöchentliche Zusammenfassung {x}', '{x} archivieren', 'Anhänge {x} ablegen'];
const FLOW_X = ['Urlaubsantrag', 'Rechnung', 'Bestellung', 'Lieferschein', 'Reisekosten', 'Störungsmeldung', 'Vertrag', 'Angebot',
  'Bewerbung', 'Schulung', 'Wartung', 'Kundenanfrage'];
const flows: Record<string, unknown>[] = [];
for (let i = 0; i < 12460; i++) {
  const envIdx = pick(envWeighted);
  const env = environments[envIdx];
  const orphan = chance(0.09);
  const creator = orphan ? pick(departed) : pick(makers);
  const r = rnd();
  const state = r < 0.74 ? 'Started' : r < 0.93 ? 'Stopped' : 'Suspended';
  const created = int(10, 1300);
  const modified = Math.min(created, recentish(600));
  flows.push({
    admin_flowid: guid(),
    admin_displayname: pick(FLOW_TEMPLATES).replace('{x}', pick(FLOW_X)),
    _admin_flowcreator_value: creator.admin_makerid,
    _admin_flowenvironment_value: env.admin_environmentid,
    admin_flowenvironmentid: envLinkId(envIdx),
    admin_flowstate: state,
    admin_flowisorphaned: orphan ? 'Orphaned' : 'Not Orphaned',
    admin_flowmodifiedon: daysAgo(modified),
    admin_flowcreatedon: daysAgo(created),
    admin_flowdeleted: chance(0.04),
    admin_flowtriggertype: pick(['Automatisiert', 'Automatisiert', 'Geplant', 'Sofort']),
    admin_flowsuspensionreason: state === 'Suspended' ? pick(['DLP-Richtlinie verletzt', 'Verbindung ungültig', 'Besitzer deaktiviert']) : null,
    admin_flowisquarantined: false,
    createdon: daysAgo(created),
    modifiedon: daysAgo(modified)
  });
}

// ---------- Connectors / Solutions (für den Tabellen-Editor) ----------
const CONNECTORS = ['SharePoint', 'Office 365 Outlook', 'Microsoft Teams', 'OneDrive for Business', 'Excel Online (Business)',
  'Approvals', 'Office 365 Users', 'Microsoft Dataverse', 'SQL Server', 'HTTP', 'Azure Blob Storage', 'Planner',
  'Microsoft Forms', 'Notifications', 'Power BI', 'Salesforce', 'SAP ERP', 'Adobe PDF Services', 'DocuSign', 'Twilio',
  'Contoso Personal-API', 'Contoso Lager-API'];
const connectors = CONNECTORS.map(n => ({
  admin_connectorid: guid(),
  admin_displayname: n,
  admin_tier: /SQL|HTTP|Azure|Salesforce|SAP|Adobe|DocuSign|Twilio|Dataverse|Contoso/.test(n) ? 'Premium' : 'Standard',
  admin_publisher: /Contoso/.test(n) ? 'Contoso' : 'Microsoft',
  admin_iscustomapi: /Contoso/.test(n),
  createdon: daysAgo(int(100, 900)),
  modifiedon: daysAgo(recentish(400))
}));
const coeSolutions = Array.from({ length: 30 }, () => {
  const envIdx = pick(envWeighted.filter(i => envSeeds[i].type !== 'Default'));
  return {
    admin_solutionid: guid(),
    admin_displayname: `${pick(['HR', 'Finanzen', 'Vertrieb', 'Einkauf', 'Marketing', 'Basis'])} ${pick(['Kernlösung', 'Apps', 'Automatisierung', 'Berichte', 'Integration'])}`,
    _admin_environment_value: environments[envIdx].admin_environmentid,
    admin_version: `1.${int(0, 9)}.${int(0, 30)}.${int(0, 9)}`,
    admin_ismanaged: chance(0.5),
    createdon: daysAgo(int(30, 700)),
    modifiedon: daysAgo(recentish(300))
  };
});

// ---------- Solutions im Default Environment ----------
const publishers = [
  { publisherid: guid(), friendlyname: 'Default Publisher for org8f3c2a1d', uniquename: 'DefaultPublisherorg8f3c2a1d' },
  { publisherid: guid(), friendlyname: 'CDS Default Publisher', uniquename: 'Cr6b2e1' },
  { publisherid: guid(), friendlyname: 'Microsoft Corporation', uniquename: 'microsoftcorporation' },
  { publisherid: guid(), friendlyname: 'Contoso IT', uniquename: 'contosoit' },
  { publisherid: guid(), friendlyname: 'Dynamics 365', uniquename: 'dynamics365' }
];
const solutionSeeds: [string, string, number, boolean, boolean][] = [
  // name, unique, publisher, managed, visible
  ['Default Solution', 'Default', 0, false, true],
  ['Common Data Services Default Solution', 'Crc6b2e1', 1, false, true],
  ['Contoso Basis-Apps', 'ContosoBasisApps', 3, false, true],
  ['Contoso Urlaubsverwaltung', 'ContosoUrlaub', 3, false, true],
  ['Contoso Spesen', 'ContosoSpesen', 3, false, true],
  ['Contoso Besucheranmeldung', 'ContosoBesucher', 3, true, true],
  ['Teams Approvals', 'msdyn_TeamsApprovals', 2, true, true],
  ['Power Apps Checker', 'msdyn_PowerAppsChecker', 2, true, true],
  ['Power Automate Desktop', 'msdyn_PowerAutomateDesktop', 2, true, true],
  ['AI Builder', 'msdyn_AIBuilder', 2, true, true],
  ['Approvals Core', 'msdyn_FlowApprovals', 2, true, true],
  ['Connection References', 'msdyn_ConnectionReferences', 2, true, true],
  ['Microsoft Dataverse Accelerator', 'msdyn_DataverseAccelerator', 2, true, true],
  ['Power Pages Core', 'mspp_PowerPagesCore', 2, true, true],
  ['Dynamics 365 Base', 'msdynce_Base', 4, true, true],
  ['Copilot Studio Basis', 'msdyn_CopilotStudio', 2, true, true],
  ['Prozessmining Vorlagen', 'msdyn_ProcessMining', 2, true, true],
  ['Mobile Offline', 'msdyn_MobileOffline', 2, true, true],
  ['Active', 'Active', 2, false, false],
  ['Basic', 'Basic', 2, true, false],
  ['System', 'System', 2, true, false]
];
const solutions = solutionSeeds.map(([name, unique, pub, managed, visible]) => {
  const installed = managed ? int(5, 900) : int(100, 1600);
  return {
    solutionid: guid(),
    friendlyname: name,
    uniquename: unique,
    version: managed ? `${int(1, 9)}.${int(0, 4)}.${int(0, 999)}.${int(0, 99)}` : `1.0.${int(0, 40)}.${int(0, 9)}`,
    ismanaged: managed,
    isvisible: visible,
    _publisherid_value: publishers[pub].publisherid,
    installedon: daysAgo(installed),
    modifiedon: daysAgo(Math.min(installed, recentish(400))),
    createdon: daysAgo(installed),
    description: managed ? 'Von Microsoft bereitgestellte verwaltete Solution.' : (unique === 'Default' ? 'Standard-Solution der Umgebung.' : 'Nicht verwaltete Contoso-Solution.')
  };
});

// ---------- Tabellen-Definitionen ----------
const A = (name: string, type: string, label: string, target?: string): IAttr => ({ name, type, label, target });
const common = [A('createdon', 'DateTime', 'Erstellt am'), A('modifiedon', 'DateTime', 'Geändert am')];

function db(host: string, defs: ITableDef[]): IDatabase {
  const tables: { [n: string]: ITableDef } = {};
  defs.forEach(d => { tables[d.logicalName] = d; });
  return { host, tables };
}

export const DATABASES: IDatabase[] = [
  db(COE_HOST, [
    {
      logicalName: 'admin_app', entitySet: 'admin_apps', displayName: 'PowerApps App', primaryId: 'admin_appid', primaryName: 'admin_displayname',
      attrs: [
        A('admin_appid', 'Uniqueidentifier', 'App'), A('admin_displayname', 'String', 'Anzeigename'),
        A('admin_appowner', 'Lookup', 'App-Besitzer', 'admin_maker'), A('admin_appenvironment', 'Lookup', 'App-Umgebung', 'admin_environment'),
        A('admin_appenvironmentid', 'String', 'App-Umgebungs-ID'), A('admin_apptype', 'String', 'App-Typ'),
        A('admin_appisorphaned', 'String', 'App ist verwaist'), A('admin_applastlaunchedon', 'DateTime', 'App zuletzt gestartet am'),
        A('admin_appmodifiedon', 'DateTime', 'App geändert am'), A('admin_appcreatedon', 'DateTime', 'App erstellt am'),
        A('admin_appdeleted', 'Boolean', 'App gelöscht'), A('admin_appplanclassification', 'String', 'Lizenzklassifizierung'),
        A('admin_appshareduserscount', 'Integer', 'Freigegeben für Benutzer'), A('admin_appsharedgroupscount', 'Integer', 'Freigegeben für Gruppen'),
        A('admin_appdescription', 'Memo', 'Beschreibung'), A('admin_appisquarantined', 'Boolean', 'Unter Quarantäne'), ...common
      ],
      rows: apps
    },
    {
      logicalName: 'admin_flow', entitySet: 'admin_flows', displayName: 'Flow', primaryId: 'admin_flowid', primaryName: 'admin_displayname',
      attrs: [
        A('admin_flowid', 'Uniqueidentifier', 'Flow'), A('admin_displayname', 'String', 'Anzeigename'),
        A('admin_flowcreator', 'Lookup', 'Flow-Ersteller', 'admin_maker'), A('admin_flowenvironment', 'Lookup', 'Flow-Umgebung', 'admin_environment'),
        A('admin_flowenvironmentid', 'String', 'Flow-Umgebungs-ID'), A('admin_flowstate', 'String', 'Flow-Status'),
        A('admin_flowisorphaned', 'String', 'Flow ist verwaist'), A('admin_flowmodifiedon', 'DateTime', 'Flow geändert am'),
        A('admin_flowcreatedon', 'DateTime', 'Flow erstellt am'), A('admin_flowdeleted', 'Boolean', 'Flow gelöscht'),
        A('admin_flowtriggertype', 'String', 'Auslöser-Typ'), A('admin_flowsuspensionreason', 'String', 'Grund der Sperre'),
        A('admin_flowisquarantined', 'Boolean', 'Unter Quarantäne'), ...common
      ],
      rows: flows
    },
    {
      logicalName: 'admin_environment', entitySet: 'admin_environments', displayName: 'Umgebung', primaryId: 'admin_environmentid', primaryName: 'admin_displayname',
      attrs: [
        A('admin_environmentid', 'Uniqueidentifier', 'Umgebung'), A('admin_displayname', 'String', 'Anzeigename'),
        A('admin_environmenttype', 'String', 'Umgebungstyp'), A('admin_environmentregion', 'String', 'Region'),
        A('admin_environmentcdsinstanceurl', 'String', 'Dataverse-Instanz-URL'), A('admin_environmentcreatedon', 'DateTime', 'Umgebung erstellt am'),
        A('admin_environmentdeleted', 'Boolean', 'Umgebung gelöscht'), A('admin_environmentsku', 'String', 'SKU'),
        A('admin_environmentstate', 'String', 'Status'), ...common
      ],
      rows: environments
    },
    {
      logicalName: 'admin_maker', entitySet: 'admin_makers', displayName: 'Maker', primaryId: 'admin_makerid', primaryName: 'admin_displayname',
      attrs: [
        A('admin_makerid', 'Uniqueidentifier', 'Maker'), A('admin_displayname', 'String', 'Anzeigename'),
        A('admin_userprincipalname', 'String', 'Benutzerprinzipalname'), A('admin_department', 'String', 'Abteilung'),
        A('admin_company', 'String', 'Firma'), A('admin_country', 'String', 'Land'), A('admin_city', 'String', 'Ort'),
        A('admin_jobtitle', 'String', 'Position'), A('admin_userisservicepricipal', 'Boolean', 'Ist Dienstprinzipal'), ...common
      ],
      rows: makers
    },
    {
      logicalName: 'admin_connector', entitySet: 'admin_connectors', displayName: 'Connector', primaryId: 'admin_connectorid', primaryName: 'admin_displayname',
      attrs: [
        A('admin_connectorid', 'Uniqueidentifier', 'Connector'), A('admin_displayname', 'String', 'Anzeigename'),
        A('admin_tier', 'String', 'Stufe'), A('admin_publisher', 'String', 'Herausgeber'), A('admin_iscustomapi', 'Boolean', 'Benutzerdefinierte API'), ...common
      ],
      rows: connectors
    },
    {
      logicalName: 'admin_solution', entitySet: 'admin_solutions', displayName: 'Solution', primaryId: 'admin_solutionid', primaryName: 'admin_displayname',
      attrs: [
        A('admin_solutionid', 'Uniqueidentifier', 'Solution'), A('admin_displayname', 'String', 'Anzeigename'),
        A('admin_environment', 'Lookup', 'Umgebung', 'admin_environment'), A('admin_version', 'String', 'Version'),
        A('admin_ismanaged', 'Boolean', 'Verwaltet'), ...common
      ],
      rows: coeSolutions
    }
  ]),
  db(DEFAULT_HOST, [
    {
      logicalName: 'solution', entitySet: 'solutions', displayName: 'Lösung', primaryId: 'solutionid', primaryName: 'friendlyname',
      attrs: [
        A('solutionid', 'Uniqueidentifier', 'Lösung'), A('friendlyname', 'String', 'Anzeigename'), A('uniquename', 'String', 'Name'),
        A('version', 'String', 'Version'), A('ismanaged', 'Boolean', 'Paket-Typ'), A('isvisible', 'Boolean', 'Ist sichtbar'),
        A('publisherid', 'Lookup', 'Herausgeber', 'publisher'), A('installedon', 'DateTime', 'Installiert am'),
        A('description', 'Memo', 'Beschreibung'), ...common
      ],
      rows: solutions
    },
    {
      logicalName: 'publisher', entitySet: 'publishers', displayName: 'Herausgeber', primaryId: 'publisherid', primaryName: 'friendlyname',
      attrs: [A('publisherid', 'Uniqueidentifier', 'Herausgeber'), A('friendlyname', 'String', 'Anzeigename'), A('uniquename', 'String', 'Name')],
      rows: publishers
    }
  ])
];
