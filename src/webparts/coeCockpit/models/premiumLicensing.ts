import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { IRecord } from './ITableConfig';
import { DataverseService, IAttributeMeta } from '../services/DataverseService';
import { detectManagedAttribute, managedValue } from './devEnvironments';
import { displayValue, rawValue } from '../utils/format';

/**
 * „Betroffen ab Februar 2027“: Ab Februar 2027 öffnet eine App in einem Managed Environment nur noch für Personen mit
 * passender Premium-Lizenz (Power Apps Premium, Power Automate Premium oder Dynamics 365 Enterprise mit Power-Apps-Rechten).
 * Quelle: Microsoft Learn „Licensing requirements for managed environments“.
 *
 * Ermittelt werden ALLE Berechtigten der Apps in Managed Environments (Besitzer, Mitbesitzer, Benutzer, Mitglieder
 * freigegebener Gruppen) und – live über Microsoft Graph – ihre Lizenzen. Dazu, wer die App in den letzten 30 Tagen
 * geöffnet hat (CoE-Audit-Log).
 */
export const PREMIUM_KEY = 'premium-2027';
export const PREMIUM_INFO_URL = 'https://learn.microsoft.com/en-us/power-platform/admin/managed-environment-licensing';
/** Microsoft nennt „ab Februar 2027“ (kein genauer Tag) – für den Countdown der 1. Februar. */
export const PREMIUM_DEADLINE = new Date(2027, 1, 1);
export const USAGE_DAYS = 30;
const CHUNK = 15;
const GRAPH_BATCH = 20;

/** Service-Pläne, die laut Microsoft für Apps in Managed Environments genügen (Graph: servicePlanName). */
export const QUALIFYING_PLANS: { [plan: string]: 'premium' | 'trial' } = {
  POWERAPPS_PER_USER: 'premium', POWERAPPS_PER_USER_GCC: 'premium', POWERAPPS_P2: 'premium',
  FLOW_PER_USER: 'premium', FLOW_PER_USER_GCC: 'premium', FLOW_P2: 'premium', POWER_AUTOMATE_ATTENDED_RPA: 'premium',
  POWERAPPS_DYN_APPS: 'premium', POWERAPPS_DYN_P2: 'premium', FLOW_DYN_APPS: 'premium', FLOW_DYN_P2: 'premium',
  POWERAPPS_P2_VIRAL: 'trial', FLOW_P2_VIRAL: 'trial'
};
/** Pro-App-Lizenzen (benutzerbezogen, deckt nur einzelne Apps). */
const PER_APP_PLANS = ['POWERAPPS_PER_APP', 'POWERAPPS_PER_APP_NEW', 'POWERAPPS_PER_APP_IW'];

const SKU_NAMES: { [sku: string]: string } = {
  POWERAPPS_PER_USER: 'Power Apps Premium', POWERAUTOMATE_ATTENDED_RPA: 'Power Automate Premium', FLOW_PER_USER: 'Power Automate pro Benutzer',
  POWERAPPS_VIRAL: 'Power Apps Plan 2 (Testversion)', POWERAPPS_DEV: 'Power Apps Developer Plan', FLOW_FREE: 'Power Automate Free',
  POWERAPPS_PER_APP: 'Power Apps pro App', POWERAPPS_PER_APP_NEW: 'Power Apps pro App',
  SPE_E3: 'Microsoft 365 E3', SPE_E5: 'Microsoft 365 E5', SPE_F1: 'Microsoft 365 F3', ENTERPRISEPACK: 'Office 365 E3', ENTERPRISEPREMIUM: 'Office 365 E5',
  STANDARDPACK: 'Office 365 E1', DESKLESSPACK: 'Office 365 F3', SPB: 'Microsoft 365 Business Premium', O365_BUSINESS_PREMIUM: 'Microsoft 365 Business Standard',
  POWER_BI_STANDARD: 'Power BI (kostenlos)', POWER_BI_PRO: 'Power BI Pro', DYN365_TEAM_MEMBERS: 'Dynamics 365 Team Members',
  DYN365_ENTERPRISE_SALES: 'Dynamics 365 Sales Enterprise', DYN365_ENTERPRISE_CUSTOMER_SERVICE: 'Dynamics 365 Customer Service Enterprise',
  DYN365_ENTERPRISE_PLAN1: 'Dynamics 365 Customer Engagement Plan', Microsoft_Teams_Premium: 'Teams Premium', TEAMS_EXPLORATORY: 'Teams Exploratory'
};

export function skuName(sku: string): string {
  return SKU_NAMES[sku] || sku.replace(/_/g, ' ');
}

export type LicenseStatus = 'premium' | 'trial' | 'perApp' | 'none' | 'notfound' | 'unknown';

export const LICENSE_LABEL: { [k in LicenseStatus]: string } = {
  premium: 'Premium vorhanden',
  trial: 'Nur Testversion',
  perApp: 'Nur Pro-App-Lizenz',
  none: 'Ohne Premium-Lizenz',
  notfound: 'Konto nicht gefunden',
  unknown: 'Nicht geprüft'
};

/** Wie ist die Person berechtigt? */
export interface IAccess { appId: string; role: string; via: string }

export interface IPremiumUser {
  /** Entra-Objekt-ID (falls bekannt) */
  id: string;
  upn: string;
  name: string;
  guest: boolean;
  access: IAccess[];
  /** App-Starts in den letzten 30 Tagen (alle Apps in Managed Environments) */
  launches: number;
  launchedApps: string[];
  license: LicenseStatus;
  licenses: string[];
}

export interface IPremiumApp {
  id: string;
  name: string;
  envId: string;
  envName: string;
  owner: string;
  tenantShared: boolean;
  users: string[];
  unlicensed: number;
  launches: number;
}

export interface IPremiumEnv {
  id: string;
  name: string;
  type: string;
  apps: number;
  users: string[];
  unlicensed: number;
  /** über Pay-as-you-go (Power Apps) abgedeckt – dann braucht niemand eine eigene Lizenz */
  payGo?: string;
}

export interface IPremiumData {
  envs: IPremiumEnv[];
  apps: IPremiumApp[];
  users: IPremiumUser[];
  /** nicht aufgelöste Gruppen (Mitglieder nicht lesbar) */
  unresolvedGroups: { id: string; name: string; size?: number }[];
  licenseState: 'ok' | 'consent' | 'forbidden' | 'error';
  licenseProblem?: string;
  groupState: 'ok' | 'forbidden' | 'none';
  usageAvailable: boolean;
  managedField: string;
  warnings: string[];
  loadedAt: Date;
}

export interface IPremiumOptions {
  managedField: string;
  onProgress?: (text: string) => void;
  /** Umgebungen mit Pay-as-you-go für Power Apps (envId → Planname), optional */
  payGo?: () => Promise<{ [envId: string]: string }>;
}

function lc(v: unknown): string {
  return String(v === null || v === undefined ? '' : v).toLowerCase();
}

function chunks<T>(a: T[], n: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < a.length; i += n) {
    out.push(a.slice(i, i + n));
  }
  return out;
}

async function pool<T>(items: T[], size: number, fn: (item: T) => Promise<void>): Promise<void> {
  let i = 0;
  const worker = async (): Promise<void> => {
    while (i < items.length) {
      await fn(items[i++]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(size, items.length) }, () => worker()));
}

interface IGraphErr { statusCode?: number; code?: string; message?: string }

function graphStatus(e: unknown): number {
  const code = typeof e === 'object' && e !== null ? (e as IGraphErr).statusCode : undefined;
  return typeof code === 'number' ? code : 0;
}

/** Lizenzstatus aus den licenseDetails einer Person. */
export function classifyLicenses(details: { skuPartNumber?: string; servicePlans?: { servicePlanName?: string; provisioningStatus?: string }[] }[]): { status: LicenseStatus; names: string[] } {
  let best: LicenseStatus = 'none';
  const names: string[] = [];
  details.forEach(d => {
    if (d.skuPartNumber) {
      names.push(skuName(d.skuPartNumber));
    }
    (d.servicePlans || []).forEach(p => {
      const name = String(p.servicePlanName || '').toUpperCase();
      const active = !p.provisioningStatus || /^(success|pendingactivation|pendingprovisioning)$/i.test(p.provisioningStatus);
      if (!active) {
        return;
      }
      const q = QUALIFYING_PLANS[name];
      if (q === 'premium') {
        best = 'premium';
      } else if (q === 'trial' && best !== 'premium') {
        best = 'trial';
      } else if (PER_APP_PLANS.indexOf(name) >= 0 && best === 'none') {
        best = 'perApp';
      }
    });
  });
  return { status: best, names: names.filter((n, i) => names.indexOf(n) === i) };
}

export async function loadPremiumImpact(service: DataverseService, graphFactory: MSGraphClientFactory | undefined, opts: IPremiumOptions): Promise<IPremiumData> {
  const warnings: string[] = [];
  const progress = (t: string): void => { if (opts.onProgress) { opts.onProgress(t); } };

  // 1. Managed Environments (gleiche Spalte wie „Betroffen ab 15.12.“)
  progress('Managed Environments werden ermittelt …');
  const attrs = await service.getAttributes('admin_environment').catch(() => [] as IAttributeMeta[]);
  const chosen = opts.managedField ? attrs.filter(a => a.selectName === opts.managedField || a.logicalName === opts.managedField)[0] : undefined;
  const managedAttr = chosen || detectManagedAttribute(attrs);
  if (!managedAttr) {
    throw new Error('Im CoE-Inventar gibt es keine Spalte für den Managed-Status (Webpart-Einstellungen → Betroffen ab 15.12. → Spalte mit dem Managed-Status).');
  }
  const typeField = attrs.some(a => a.logicalName === 'admin_environmentsku') ? 'admin_environmentsku' : 'admin_environmenttype';
  const envRows = await service.query('admin_environments', ['admin_environmentid', 'admin_displayname', typeField, managedAttr.selectName], 'admin_environmentdeleted ne true', 10000);
  const managed = envRows.filter(r => managedValue(r, managedAttr.selectName) === true);
  const envById: { [id: string]: IRecord } = {};
  managed.forEach(r => { envById[lc(r.admin_environmentid)] = r; });
  const envIds = Object.keys(envById);

  // 2. Apps in diesen Environments
  progress('Apps werden geladen …');
  const appRows: IRecord[] = [];
  for (const part of chunks(envIds, CHUNK)) {
    appRows.push(...await service.query('admin_apps', ['admin_appid', 'admin_displayname', '_admin_appenvironment_value', '_admin_appowner_value', 'admin_appsharedwithtenant'],
      `admin_appdeleted ne true and (${part.map(id => `_admin_appenvironment_value eq ${id}`).join(' or ')})`, 50000)
      .catch(async () => service.query('admin_apps', ['admin_appid', 'admin_displayname', '_admin_appenvironment_value', '_admin_appowner_value'],
        `admin_appdeleted ne true and (${part.map(id => `_admin_appenvironment_value eq ${id}`).join(' or ')})`, 50000)));
  }
  const apps: { [id: string]: IPremiumApp } = {};
  appRows.forEach(r => {
    const id = lc(r.admin_appid);
    const envId = lc(r._admin_appenvironment_value);
    apps[id] = {
      id, name: displayValue(r, 'admin_displayname') || id, envId,
      envName: envById[envId] ? displayValue(envById[envId], 'admin_displayname') : displayValue(r, '_admin_appenvironment_value'),
      owner: displayValue(r, '_admin_appowner_value'), tenantShared: r.admin_appsharedwithtenant === true,
      users: [], unlicensed: 0, launches: 0
    };
  });
  const appIds = Object.keys(apps);

  // 3. Berechtigungen (Besitzer, Mitbesitzer, Benutzer, Gruppen)
  progress('Berechtigungen werden geladen …');
  const roles: IRecord[] = [];
  for (const part of chunks(envIds, CHUNK)) {
    try {
      roles.push(...await service.query('admin_powerplatformuserroles', ['_admin_app_value', '_admin_powerplatformuser_value', 'admin_rolename'],
        `statecode eq 0 and _admin_app_value ne null and (${part.map(id => `_admin_environment_value eq ${id}`).join(' or ')})`, 100000));
    } catch (e) {
      warnings.push(`Berechtigungen konnten nicht vollständig geladen werden: ${(e as Error).message}`);
      break;
    }
  }
  const principalIds = roles.map(r => lc(r._admin_powerplatformuser_value)).filter((x, i, a) => !!x && a.indexOf(x) === i);
  const principals: { [id: string]: IRecord } = {};
  for (const part of chunks(principalIds, 25)) {
    const rows = await service.query('admin_powerplatformusers', ['admin_powerplatformuserid', 'admin_displayname', 'admin_type', 'admin_userprincipalname', 'admin_useremail', 'admin_userisguest', 'admin_groupsize'],
      part.map(id => `admin_powerplatformuserid eq ${id}`).join(' or '), 100).catch(() => [] as IRecord[]);
    rows.forEach(p => { principals[lc(p.admin_powerplatformuserid)] = p; });
  }

  const users: { [key: string]: IPremiumUser } = {};
  const userFor = (id: string, upn: string, name: string, guest: boolean): IPremiumUser => {
    const key = (upn || id).toLowerCase();
    if (!users[key]) {
      users[key] = { id, upn, name: name || upn || id, guest, access: [], launches: 0, launchedApps: [], license: 'unknown', licenses: [] };
    } else if (!users[key].id && id) {
      users[key].id = id;
    }
    return users[key];
  };
  const groups: { [id: string]: { name: string; size?: number; apps: { appId: string; role: string }[] } } = {};
  const ROLE: { [k: string]: string } = { owner: 'Besitzer', canedit: 'Mitbesitzer', canview: 'Benutzer', canviewwithshare: 'Benutzer' };
  roles.forEach(r => {
    const appId = lc(r._admin_app_value);
    if (!apps[appId]) {
      return;
    }
    const pid = lc(r._admin_powerplatformuser_value);
    const p = principals[pid];
    const role = ROLE[lc(r.admin_rolename)] || String(r.admin_rolename || 'Benutzer');
    const type = lc(p ? p.admin_type : '');
    if (/group/.test(type)) {
      const g = groups[pid] = groups[pid] || { name: p ? displayValue(p, 'admin_displayname') : pid, size: p && p.admin_groupsize ? Number(p.admin_groupsize) : undefined, apps: [] };
      g.apps.push({ appId, role });
      return;
    }
    if (/tenant/.test(type)) {
      apps[appId].tenantShared = true;
      return;
    }
    const upn = p ? (rawValue(p, 'admin_userprincipalname') || rawValue(p, 'admin_useremail')) : '';
    const u = userFor(pid, upn, p ? displayValue(p, 'admin_displayname') : displayValue(r, '_admin_powerplatformuser_value'), !!p && p.admin_userisguest === true);
    u.access.push({ appId, role, via: 'direkt' });
  });

  // 4. Gruppen auflösen (Graph, GroupMember.Read.All)
  let graph: MSGraphClientV3 | undefined;
  let licenseState: IPremiumData['licenseState'] = 'ok';
  let licenseProblem: string | undefined;
  if (graphFactory) {
    try {
      graph = await graphFactory.getClient('3');
    } catch (e) {
      licenseState = 'consent';
      licenseProblem = (e as Error).message;
    }
  } else {
    licenseState = 'error';
    licenseProblem = 'Microsoft Graph ist nicht verfügbar.';
  }
  const unresolvedGroups: IPremiumData['unresolvedGroups'] = [];
  let groupState: IPremiumData['groupState'] = Object.keys(groups).length ? 'ok' : 'none';
  const groupIds = Object.keys(groups);
  if (groupIds.length) {
    progress(`Mitglieder von ${groupIds.length} Gruppe(n) werden gelesen …`);
    await pool(groupIds, 3, async gid => {
      const g = groups[gid];
      if (!graph || groupState === 'forbidden') {
        unresolvedGroups.push({ id: gid, name: g.name, size: g.size });
        return;
      }
      try {
        let url: string | undefined = `/groups/${gid}/transitiveMembers/microsoft.graph.user?$select=id,userPrincipalName,displayName,mail&$top=999`;
        for (let n = 0; url && n < 30; n++) {
          const page = await graph.api(url).version('v1.0').get() as { value?: { id: string; userPrincipalName?: string; displayName?: string; mail?: string }[]; '@odata.nextLink'?: string };
          (page.value || []).forEach(m => {
            const u = userFor(lc(m.id), m.userPrincipalName || m.mail || '', m.displayName || '', /#ext#/i.test(m.userPrincipalName || ''));
            g.apps.forEach(a => u.access.push({ appId: a.appId, role: a.role, via: `Gruppe „${g.name}“` }));
          });
          url = page['@odata.nextLink'];
        }
      } catch (e) {
        if (graphStatus(e) === 403 || graphStatus(e) === 401) {
          groupState = 'forbidden';
        }
        unresolvedGroups.push({ id: gid, name: g.name, size: g.size });
      }
    });
  }

  // 5. Wer hat die Apps in den letzten 30 Tagen geöffnet? (CoE-Audit-Log)
  progress('App-Starts der letzten 30 Tage werden ausgewertet …');
  let usageAvailable = false;
  try {
    const since = new Date(Date.now() - USAGE_DAYS * 86400000).toISOString();
    const pairs = await service.getGroupedPairs('admin_auditlog', 'admin_userupn', '_admin_applookup_value', `admin_operation eq 'LaunchPowerApp' and admin_creationtime ge ${since}`);
    usageAvailable = true;
    pairs.forEach(p => {
      const app = apps[p.b];
      if (!app || !p.a) {
        return;
      }
      app.launches += p.n;
      const u = userFor('', p.a, '', false);
      u.launches += p.n;
      if (u.launchedApps.indexOf(p.b) < 0) {
        u.launchedApps.push(p.b);
      }
      // Bei „für alle freigegeben“ sind Starts die einzige Quelle der Berechtigten
      if (!u.access.some(a => a.appId === p.b)) {
        u.access.push({ appId: p.b, role: 'Benutzer', via: app.tenantShared ? 'Organisation' : 'geöffnet' });
      }
    });
  } catch (e) {
    warnings.push(`App-Starts aus dem CoE-Audit-Log nicht verfügbar: ${(e as Error).message}`);
  }

  // 6. Lizenzen (Graph, User.Read.All) – je 20 Personen in einem $batch
  const list = Object.keys(users).map(k => users[k]).filter(u => u.access.length > 0);
  if (graph && licenseState === 'ok') {
    progress(`Lizenzen von ${list.length} Personen werden geprüft …`);
    const targets = list.filter(u => !!(u.id || u.upn));
    let done = 0;
    await pool(chunks(targets, GRAPH_BATCH), 3, async part => {
      if (licenseState !== 'ok') {
        return;
      }
      try {
        const res = await graph!.api('/$batch').version('v1.0').post({
          requests: part.map((u, i) => ({ id: String(i), method: 'GET', url: `/users/${encodeURIComponent(u.id || u.upn)}/licenseDetails?$select=skuPartNumber,servicePlans` }))
        }) as { responses?: { id: string; status: number; body?: { value?: { skuPartNumber?: string; servicePlans?: { servicePlanName?: string; provisioningStatus?: string }[] }[]; error?: { message?: string } } }[] };
        (res.responses || []).forEach(r => {
          const u = part[Number(r.id)];
          if (!u) {
            return;
          }
          if (r.status === 200) {
            const c = classifyLicenses((r.body && r.body.value) || []);
            u.license = c.status;
            u.licenses = c.names;
          } else if (r.status === 404) {
            u.license = 'notfound';
          } else if (r.status === 403 || r.status === 401) {
            licenseState = 'forbidden';
            licenseProblem = (r.body && r.body.error && r.body.error.message) || 'Zugriff verweigert';
          }
        });
      } catch (e) {
        if (graphStatus(e) === 403 || graphStatus(e) === 401) {
          licenseState = 'forbidden';
        } else {
          licenseState = 'error';
        }
        licenseProblem = (e as IGraphErr).message || String(e);
      }
      done += part.length;
      progress(`Lizenzen werden geprüft … ${done} von ${targets.length}`);
    });
  }

  // 7. Pay-as-you-go (Power Apps) je Environment – optional über die Power Platform API
  let payGo: { [envId: string]: string } = {};
  if (opts.payGo) {
    payGo = await opts.payGo().catch(() => ({}));
  }

  // 8. Zusammenfassen
  const blocked = (u: IPremiumUser): boolean => u.license === 'none' || u.license === 'perApp';
  list.forEach(u => {
    u.access.forEach(a => {
      const app = apps[a.appId];
      if (app && app.users.indexOf(u.upn || u.id) < 0) {
        app.users.push(u.upn || u.id);
        if (blocked(u) && !payGo[app.envId]) {
          app.unlicensed++;
        }
      }
    });
  });
  const envs: IPremiumEnv[] = envIds.map(id => {
    const r = envById[id];
    const envApps = appIds.filter(a => apps[a].envId === id);
    const envUsers: string[] = [];
    let unlicensed = 0;
    list.forEach(u => {
      if (u.access.some(a => apps[a.appId] && apps[a.appId].envId === id)) {
        envUsers.push(u.upn || u.id);
        if (blocked(u) && !payGo[id]) {
          unlicensed++;
        }
      }
    });
    return { id, name: displayValue(r, 'admin_displayname'), type: displayValue(r, typeField), apps: envApps.length, users: envUsers, unlicensed, payGo: payGo[id] };
  });

  return {
    envs,
    apps: appIds.map(a => apps[a]),
    users: list,
    unresolvedGroups,
    licenseState,
    licenseProblem,
    groupState,
    usageAvailable,
    managedField: managedAttr.displayName,
    warnings,
    loadedAt: new Date()
  };
}
