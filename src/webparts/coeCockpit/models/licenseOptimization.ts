import { MSGraphClientFactory } from '@microsoft/sp-http';
import { DataverseService } from '../services/DataverseService';
import { daysAgoIso } from './overview';
import { ILicUser, loadUserLicenses } from './userLicenses';

/**
 * Lizenz-Optimierung: Wer hat eine Power-Platform-Premium-Lizenz (live aus Microsoft Graph), nutzt sie aber nicht?
 *  - Power Apps Premium: genutzt, wenn die Person im Zeitraum eine Premium-App gestartet hat (CoE-Audit-Log „LaunchPowerApp“,
 *    Lizenzklasse der App laut CoE).
 *  - Power Automate Premium: genutzt, wenn die Person Besitzer eines eingeschalteten Flows ist (CoE).
 * Das ist eine Näherung – sie zeigt Kandidaten zum Nachfragen, keine sichere Aussage.
 */
export const LICENSE_OPT_KEY = 'license-optimization';

export interface IPremiumSkuInfo { kind: 'apps' | 'flows'; price: number }

/** Premium-Lizenzen (Teilenummer) mit Art und Listenpreis in Euro pro Monat (anpassbar). */
export const PREMIUM_SKUS: { [part: string]: IPremiumSkuInfo } = {
  POWERAPPS_PER_USER: { kind: 'apps', price: 20 },
  POWERAPPS_P2: { kind: 'apps', price: 40 },
  POWERAPPS_P1: { kind: 'apps', price: 7 },
  POWERAUTOMATE_ATTENDED_RPA: { kind: 'flows', price: 15 },
  FLOW_PER_USER: { kind: 'flows', price: 15 },
  FLOW_P2: { kind: 'flows', price: 15 }
};

export interface IOptLicense { id: string; part: string; name: string; kind: 'apps' | 'flows'; used: boolean | undefined }

export interface IOptRow {
  user: ILicUser;
  licenses: IOptLicense[];
  /** Starts von Premium-Apps im Zeitraum */
  premiumLaunches: number;
  /** verschiedene Premium-Apps im Zeitraum */
  premiumApps: number;
  /** Starts aller Apps im Zeitraum */
  anyLaunches: number;
  /** letzter App-Start überhaupt (laut CoE) */
  lastLaunch: string;
  /** eingeschaltete Flows, deren Besitzer die Person ist */
  activeFlows: number;
  status: 'unused' | 'partial' | 'used' | 'unknown';
}

export interface IOptData {
  rows: IOptRow[];
  /** Premium-Lizenzen im Tenant (id → Name, Teilenummer) */
  skus: { id: string; part: string; name: string; kind: 'apps' | 'flows' }[];
  days: number;
  problems: string[];
  loadedAt: Date;
}

export async function loadLicenseOptimization(service: DataverseService, graph: MSGraphClientFactory, days: number): Promise<IOptData> {
  const lic = await loadUserLicenses(graph);
  const problems: string[] = [];
  const skus = Object.keys(lic.skus).map(id => lic.skus[id]).filter(s => !!PREMIUM_SKUS[s.part])
    .map(s => ({ id: s.id, part: s.part, name: s.name, kind: PREMIUM_SKUS[s.part].kind }));
  if (lic.skuSource === 'ids') {
    problems.push('Namen der Lizenzen nicht lesbar – Premium-Lizenzen können nicht erkannt werden.');
  }
  const premiumIds = skus.map(s => s.id);
  const candidates = lic.users.filter(u => u.skus.some(id => premiumIds.indexOf(id) >= 0));

  const since = daysAgoIso(days);
  const [pairs, last, apps, flows] = await Promise.all([
    service.getGroupedPairs('admin_auditlog', 'admin_userupn', '_admin_applookup_value', `admin_operation eq 'LaunchPowerApp' and admin_creationtime ge ${since}`)
      .catch(e => { problems.push(`App-Starts aus dem CoE-Audit-Log nicht lesbar: ${e instanceof Error ? e.message : String(e)}`); return undefined; }),
    service.getGroupedMax('admin_auditlog', 'admin_userupn', "admin_operation eq 'LaunchPowerApp'", 'admin_creationtime')
      .catch(() => ({} as { [upn: string]: string })),
    service.query('admin_apps', ['admin_appid', 'admin_appplanclassification'], 'statecode eq 0', 200000)
      .catch(e => { problems.push(`Lizenzklasse der Apps nicht lesbar: ${e instanceof Error ? e.message : String(e)}`); return undefined; }),
    service.query('admin_flows', ['_admin_derivedowner_value'], "statecode eq 0 and admin_flowdeleted ne true and admin_flowstate eq 'Started'", 200000)
      .catch(e => { problems.push(`Flows aus dem CoE nicht lesbar: ${e instanceof Error ? e.message : String(e)}`); return undefined; })
  ]);
  if (pairs && pairs.length === 0) {
    problems.push(`Im CoE-Audit-Log gibt es in den letzten ${days} Tagen keine App-Starts – ist die Audit-Log-Synchronisierung des CoE eingerichtet? Ohne sie gelten alle Power-Apps-Lizenzen als ungenutzt.`);
  }

  const premiumApps = new Set<string>();
  (apps || []).forEach(a => {
    const cls = String(a['admin_appplanclassification@OData.Community.Display.V1.FormattedValue'] || a.admin_appplanclassification || '');
    if (/premium/i.test(cls)) {
      premiumApps.add(String(a.admin_appid || '').toLowerCase());
    }
  });
  const usage: { [upn: string]: { premium: number; any: number; apps: Set<string> } } = {};
  (pairs || []).forEach(p => {
    const u = usage[p.a] = usage[p.a] || { premium: 0, any: 0, apps: new Set<string>() };
    u.any += p.n;
    if (premiumApps.has(p.b)) {
      u.premium += p.n;
      u.apps.add(p.b);
    }
  });
  const lastBy: { [upn: string]: string } = {};
  Object.keys(last || {}).forEach(k => { lastBy[k.toLowerCase()] = (last as { [k: string]: string })[k]; });
  const flowCount: { [owner: string]: number } = {};
  (flows || []).forEach(f => {
    const o = String(f._admin_derivedowner_value || '').toLowerCase();
    if (o) {
      flowCount[o] = (flowCount[o] || 0) + 1;
    }
  });

  const rows: IOptRow[] = candidates.map(user => {
    const upn = user.upn.toLowerCase();
    const use = usage[upn];
    const activeFlows = flowCount[user.id.toLowerCase()] || 0;
    const licenses: IOptLicense[] = skus.filter(s => user.skus.indexOf(s.id) >= 0).map(s => ({
      ...s,
      used: s.kind === 'apps' ? (pairs ? !!use && use.premium > 0 : undefined) : (flows ? activeFlows > 0 : undefined)
    }));
    const known = licenses.filter(l => l.used !== undefined);
    const status: IOptRow['status'] = !known.length ? 'unknown'
      : known.every(l => l.used) ? 'used' : known.every(l => !l.used) ? 'unused' : 'partial';
    return {
      user, licenses,
      premiumLaunches: use ? use.premium : 0,
      premiumApps: use ? use.apps.size : 0,
      anyLaunches: use ? use.any : 0,
      lastLaunch: lastBy[upn] || '',
      activeFlows,
      status
    };
  });
  return { rows, skus, days, problems, loadedAt: new Date() };
}
