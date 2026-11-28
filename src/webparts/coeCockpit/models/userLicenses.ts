import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { skuName } from './premiumLicensing';

/**
 * „Lizenzen je Benutzer“: live aus Microsoft Graph (nicht im CoE).
 *  - Benutzer mit zugewiesenen Lizenzen: GET /users?$select=…,assignedLicenses  (Berechtigung „User.Read.All“)
 *  - Namen und Mengen der Lizenzen: GET /subscribedSkus (gekauft/verbraucht; braucht je nach Tenant „Organization.Read.All“
 *    bzw. „LicenseAssignment.Read.All“). Geht das nicht, holt das Cockpit die Namen über „licenseDetails“ einzelner Benutzer
 *    (reicht mit „User.Read.All“) – dann ohne gekauft/verbraucht.
 */
export const USER_LICENSES_KEY = 'user-licenses';

export interface ISku {
  /** skuId (klein geschrieben) */
  id: string;
  /** Teilenummer, z. B. SPE_E3 */
  part: string;
  /** Anzeigename, z. B. Microsoft 365 E3 */
  name: string;
  /** gekauft (aktiv) laut subscribedSkus */
  enabled?: number;
  /** verbraucht laut subscribedSkus */
  consumed?: number;
  /** zugewiesen in den geladenen Benutzern */
  assigned: number;
}

export interface ILicUser {
  id: string;
  name: string;
  upn: string;
  department: string;
  jobTitle: string;
  enabled: boolean;
  guest: boolean;
  /** skuIds (klein) */
  skus: string[];
}

export interface IUserLicenseData {
  users: ILicUser[];
  skus: { [id: string]: ISku };
  /** Woher die Namen der Lizenzen stammen */
  skuSource: 'subscribedSkus' | 'licenseDetails' | 'ids';
  skuProblem?: string;
  truncated: boolean;
  loadedAt: Date;
}

export class LicenseLoadError extends Error {
  constructor(message: string, public kind: 'consent' | 'forbidden' | 'failed', public detail?: string) {
    super(message);
    Object.setPrototypeOf(this, LicenseLoadError.prototype);
  }
}

const USER_FIELDS = 'id,displayName,userPrincipalName,mail,department,jobTitle,accountEnabled,userType,assignedLicenses';

interface IGraphError { statusCode?: number; code?: string; message?: string }

function graphError(e: unknown, what: string): LicenseLoadError {
  const err = (e || {}) as IGraphError;
  const msg = err.message || (e instanceof Error ? e.message : String(e));
  if (/AADSTS65001|consent/i.test(msg)) {
    return new LicenseLoadError(`${what}: Microsoft Graph ist für SharePoint noch nicht freigegeben (Berechtigung „User.Read.All“, Skript -Licenses).`, 'consent', msg);
  }
  if (err.statusCode === 401 || err.statusCode === 403 || /Authorization_RequestDenied|Insufficient privileges/i.test(msg)) {
    return new LicenseLoadError(`${what}: Zugriff verweigert. Nötig ist die Berechtigung „User.Read.All“ für SharePoint (Skript -Licenses).`, 'forbidden', msg);
  }
  return new LicenseLoadError(`${what}: ${msg}`, 'failed', msg);
}

const store: { at: number; key: string | undefined; promise: Promise<IUserLicenseData> | undefined } = { at: 0, key: undefined, promise: undefined };

export function resetUserLicenseCache(): void {
  store.at = 0;
  store.key = undefined;
  store.promise = undefined;
}

/**
 * Laden (10 Minuten zwischengespeichert). reloadKey: Zähler von „Aktualisieren“ im Cockpit – ändert er sich,
 * wird neu geladen (sonst kämen nach „Aktualisieren“ bis zu 10 Minuten alte Lizenzen).
 */
export function loadUserLicenses(factory: MSGraphClientFactory, onProgress?: (users: number) => void, fresh: boolean = false, reloadKey?: string): Promise<IUserLicenseData> {
  const keyChanged = reloadKey !== undefined && store.key !== undefined && reloadKey !== store.key;
  if (!fresh && !keyChanged && store.promise && Date.now() - store.at < 10 * 60000) {
    return store.promise;
  }
  const p = load(factory, onProgress);
  store.at = Date.now();
  store.key = reloadKey !== undefined ? reloadKey : store.key;
  store.promise = p;
  p.catch(() => resetUserLicenseCache());
  return p;
}

async function load(factory: MSGraphClientFactory, onProgress?: (users: number) => void): Promise<IUserLicenseData> {
  const graph: MSGraphClientV3 = await factory.getClient('3');
  const users: ILicUser[] = [];
  let url: string | undefined = `/users?$select=${USER_FIELDS}&$top=999`;
  let pages = 0;
  while (url && pages < 200) {
    let page: { value?: { [k: string]: unknown }[]; '@odata.nextLink'?: string };
    try {
      page = await graph.api(url).version('v1.0').get() as typeof page;
    } catch (e) {
      throw graphError(e, 'Benutzer konnten nicht gelesen werden');
    }
    (page.value || []).forEach(u => {
      const lic = (u.assignedLicenses || []) as { skuId?: string }[];
      users.push({
        id: String(u.id || ''),
        name: String(u.displayName || u.userPrincipalName || ''),
        upn: String(u.userPrincipalName || u.mail || ''),
        department: String(u.department || ''),
        jobTitle: String(u.jobTitle || ''),
        enabled: u.accountEnabled !== false,
        guest: /guest/i.test(String(u.userType || '')),
        skus: lic.map(l => String(l.skuId || '').toLowerCase()).filter(x => !!x)
      });
    });
    pages++;
    if (onProgress) {
      onProgress(users.length);
    }
    url = page['@odata.nextLink'];
  }
  const truncated = !!url;

  // Zugewiesen je Lizenz
  const skus: { [id: string]: ISku } = {};
  users.forEach(u => u.skus.forEach(id => {
    skus[id] = skus[id] || { id, part: '', name: '', assigned: 0 };
    skus[id].assigned++;
  }));

  // Namen und Mengen: subscribedSkus, sonst licenseDetails einzelner Benutzer
  let skuSource: IUserLicenseData['skuSource'] = 'ids';
  let skuProblem: string | undefined;
  try {
    const res = await graph.api('/subscribedSkus').version('v1.0').get() as { value?: { skuId?: string; skuPartNumber?: string; consumedUnits?: number; prepaidUnits?: { enabled?: number } }[] };
    (res.value || []).forEach(s => {
      const id = String(s.skuId || '').toLowerCase();
      if (!id) {
        return;
      }
      skus[id] = skus[id] || { id, part: '', name: '', assigned: 0 };
      skus[id].part = String(s.skuPartNumber || '');
      skus[id].enabled = s.prepaidUnits && typeof s.prepaidUnits.enabled === 'number' ? s.prepaidUnits.enabled : undefined;
      skus[id].consumed = typeof s.consumedUnits === 'number' ? s.consumedUnits : undefined;
    });
    skuSource = 'subscribedSkus';
  } catch (e) {
    skuProblem = 'Gekaufte und freie Lizenzen nicht lesbar: Microsoft Graph „subscribedSkus“ braucht zusätzlich „Organization.Read.All“ (Skript -Licenses). Die Namen der Lizenzen kommen aus den Lizenzdetails der Benutzer.';
    // je Lizenz einen Benutzer nehmen und dessen licenseDetails lesen
    const missing = Object.keys(skus).filter(id => !skus[id].part);
    const probe: { [userId: string]: true } = {};
    missing.forEach(id => {
      const u = users.filter(x => x.skus.indexOf(id) >= 0)[0];
      if (u) {
        probe[u.id] = true;
      }
    });
    const ids = Object.keys(probe);
    try {
      for (let i = 0; i < ids.length; i += 20) {
        const part = ids.slice(i, i + 20);
        const batch = await graph.api('/$batch').version('v1.0').post({
          requests: part.map((id, n) => ({ id: String(n), method: 'GET', url: `/users/${id}/licenseDetails?$select=skuId,skuPartNumber` }))
        }) as { responses?: { status: number; body?: { value?: { skuId?: string; skuPartNumber?: string }[] } }[] };
        (batch.responses || []).forEach(r => {
          if (r.status === 200 && r.body && r.body.value) {
            r.body.value.forEach(l => {
              const id = String(l.skuId || '').toLowerCase();
              if (skus[id] && !skus[id].part) {
                skus[id].part = String(l.skuPartNumber || '');
              }
            });
          }
        });
      }
      skuSource = 'licenseDetails';
    } catch {
      skuProblem = 'Namen der Lizenzen nicht lesbar – angezeigt werden die Lizenz-IDs.';
    }
  }
  Object.keys(skus).forEach(id => {
    const s = skus[id];
    s.name = s.part ? skuName(s.part) : `Lizenz ${id.slice(0, 8)}`;
  });
  return { users, skus, skuSource, skuProblem, truncated, loadedAt: new Date() };
}
