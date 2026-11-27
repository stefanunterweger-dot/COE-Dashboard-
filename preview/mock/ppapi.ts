/* Simulierte Power Platform API (https://api.powerplatform.com) für die Vorschau: Copilot-Credits und Abrechnungspläne. */
import { COE_HOST, DATABASES } from './data';
import { IMockResponse } from './odata';

interface IWin { __coePpApiDeny?: '' | 'forbidden' | 'consent'; __coePpApiCalls?: string[] }
const win = window as unknown as IWin;

function respond(status: number, body: unknown, headers: { [k: string]: string } = {}): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (n: string) => headers[n] || headers[n.toLowerCase()] || null },
    json: () => Promise.resolve(text ? JSON.parse(text) : undefined),
    text: () => Promise.resolve(text)
  };
}

// ---------- Daten (aus den CoE-Environments der Vorschau abgeleitet) ----------
interface IEnv { id: string; name: string; type: string; managed: boolean }
let cache: {
  envs: IEnv[];
  alloc: { [id: string]: { allocated: number; pool: boolean } };
  consumed: { [id: string]: number };
  payGo: { [id: string]: number };
  plans: { id: string; name: string; status: string; location: string; sub: string; rg: string; created: string; products: string[]; envs: string[] }[];
} | undefined;

function data(): NonNullable<typeof cache> {
  if (cache) {
    return cache;
  }
  const coe = DATABASES.filter(d => d.host === COE_HOST)[0];
  const rows = coe.tables.admin_environment.rows.filter(r => !r.admin_environmentdeleted);
  const envs: IEnv[] = rows.map(r => ({
    id: String(r.admin_environmentname).toLowerCase(),
    name: String(r.admin_displayname),
    type: String(r.admin_environmentsku),
    managed: !!r.admin_governanceprotectionlevel && r.admin_governanceprotectionlevel !== 'Basic'
  }));
  const byName = (n: string): IEnv | undefined => envs.filter(e => e.name === n)[0];
  const alloc: NonNullable<typeof cache>['alloc'] = {};
  const consumed: { [id: string]: number } = {};
  const payGo: { [id: string]: number } = {};
  const set = (name: string, allocated: number, used: number, pool: boolean, pay?: number): void => {
    const e = byName(name);
    if (!e) {
      return;
    }
    if (allocated) {
      alloc[e.id] = { allocated, pool };
    }
    consumed[e.id] = used;
    if (pay !== undefined) {
      payGo[e.id] = pay;
    }
  };
  set('Vertrieb Produktion', 25000, 18430.5, true, 0);
  set('Finanzen Produktion', 20000, 20000, false, 274517);
  set('HR Test', 12500, 2449.91, true);
  set('Marketing Kampagnen', 10000, 9120, true);
  set('Finanzen Entwicklung', 4000, 1466.38, false);
  set('Einkauf Sandbox', 1000, 1000, false, 5720.4);
  // Verbrauch ohne Zuweisung (zieht aus der Tenant-Kapazität)
  const dev = envs.filter(e => e.type === 'Developer').slice(0, 2);
  dev.forEach((e, i) => { consumed[e.id] = i === 0 ? 312.4 : 0; });
  const plans = [
    { id: 'b7d1c2e0-1111-4a6e-9a01-000000000001', name: 'CopilotCreditsProd', status: 'Enabled', location: 'europe', sub: '4f1c2b7e-9d3a-4c11-8e2f-5a6b7c8d9e01', rg: 'rg-pp-billing-prd', created: '2025-10-23T09:12:00Z', products: ['PowerVirtualAgent'], envs: ['Finanzen Produktion', 'Einkauf Sandbox'] },
    { id: 'b7d1c2e0-1111-4a6e-9a01-000000000002', name: 'SelfServiceBilling', status: 'Enabled', location: 'europe', sub: '4f1c2b7e-9d3a-4c11-8e2f-5a6b7c8d9e02', rg: 'rg-pp-selfservice', created: '2026-03-29T14:40:00Z', products: ['Dataverse', 'PowerApps', 'PowerAutomate'], envs: ['Vertrieb Produktion'] },
    { id: 'b7d1c2e0-1111-4a6e-9a01-000000000003', name: 'SubscriptionCreation', status: 'DisabledBySubscription', location: 'europe', sub: '', rg: '', created: '2023-06-12T08:00:00Z', products: ['PowerAutomate', 'PowerApps'], envs: [] },
    { id: 'b7d1c2e0-1111-4a6e-9a01-000000000004', name: 'HackathonBilling', status: 'Enabled', location: 'europe', sub: '4f1c2b7e-9d3a-4c11-8e2f-5a6b7c8d9e04', rg: 'rg-hackathon', created: '2025-03-27T10:00:00Z', products: ['Dataverse', 'PowerApps', 'PowerAutomate', 'PowerVirtualAgent'], envs: [] }
  ].map(p => ({ ...p, envs: p.envs.map(n => (byName(n) || { id: '' }).id).filter(x => !!x) }));
  cache = { envs, alloc, consumed, payGo, plans };
  return cache;
}

const MCS = 'MCSMessages';

function envEntitlement(e: IEnv): unknown {
  const d = data();
  const a = d.alloc[e.id];
  const used = d.consumed[e.id] || 0;
  const plan = d.plans.filter(p => p.envs.indexOf(e.id) >= 0)[0];
  const status = !a ? 'NotSpecified' : used >= a.allocated ? (plan || a.pool ? 'CoveredOverage' : 'Overage') : 'WithinCapacity';
  return {
    entitlementId: MCS,
    environmentId: e.id,
    environmentName: e.name,
    environmentType: e.type,
    isManagedEnvironment: e.managed,
    productCategories: ['PowerVirtualAgent'],
    entitlement: {
      unit: 'Count',
      capacity: {
        allocated: { value: a ? a.allocated : 0, autoAllocated: 0 },
        availableQuantity: a ? Math.max(0, a.allocated - used) : 0,
        consumed: { consumptionType: 'Actual', value: used, lastUpdatedOn: '2026-11-26T06:00:00Z' },
        status,
        enforcementRules: [{ ruleType: 'TenantPool', enabled: a ? a.pool : true }]
      },
      payGo: plan ? { consumed: { value: d.payGo[e.id] || 0 }, entitled: { value: 1.7976931348623157e308 } } : undefined
    }
  };
}

function route(url: URL): IMockResponse {
  const path = url.pathname.replace(/\/+$/, '');
  const version = url.searchParams.get('api-version') || '';
  const d = data();
  let m: RegExpMatchArray | null;

  if (path === '/licensing/tenantCapacity/currencyReports') {
    // Testet den Fallback auf eine andere api-version
    if (version !== '2022-03-01-preview') {
      return respond(400, { error: { code: 'InvalidApiVersion', message: `The api-version '${version}' is not supported for this resource.` } });
    }
    const allocated = Object.keys(d.alloc).reduce((s, k) => s + d.alloc[k].allocated, 0);
    return respond(200, [{ currencyType: MCS, purchased: 100000, allocated }, { currencyType: 'AI', purchased: 0, allocated: 0 }]);
  }
  if (path === '/licensing/tenantCapacity') {
    // Verbrauch aus den CoE-Kapazitätsdaten der Vorschau, Kontingent fest
    const coe = DATABASES.filter(x => x.host === COE_HOST)[0];
    const rows = coe.tables.admin_environmentcapacity ? coe.tables.admin_environmentcapacity.rows : [];
    const sum = (t: string): number => rows.filter(r => String(r.admin_capacitytype) === t).reduce((s2, r) => s2 + (Number(r.admin_actualconsumption) || 0), 0);
    const cap = (t: string, total: number) => ({ capacityType: t, capacityUnits: 'MB', totalCapacity: total, maxCapacity: total, consumption: { actual: Math.round(sum(t)), rated: Math.round(sum(t) * 1.02) }, status: 'Available', overflowCapacity: [] });
    const db = sum('Database');
    return respond(200, {
      tenantId: '8f3c2a1d-6b4e-4c7a-9d2f-1e5b7a9c3d40', licenseModelType: 'StorageDriven',
      tenantCapacities: [cap('Database', Math.round(db / 0.86)), cap('File', Math.round(sum('File') / 0.42 + 1)), cap('Log', Math.round(sum('Log') / 0.95 + 1))],
      capacitySummary: { status: 'Available', statusMessage: 'Database capacity more than 85 % used', statusMessageCode: 'DBCapacityMoreThanEightyFive' }
    });
  }
  if (path === `/licensing/entitlements/${MCS}`) {
    const allocated = Object.keys(d.alloc).reduce((s, k) => s + d.alloc[k].allocated, 0);
    const used = Object.keys(d.consumed).reduce((s, k) => s + d.consumed[k], 0);
    const pay = Object.keys(d.payGo).reduce((s, k) => s + d.payGo[k], 0);
    return respond(200, {
      entitlementId: MCS, productCategories: ['PowerVirtualAgent'],
      entitlement: { unit: 'Count', capacity: { entitled: { value: 100000 }, allocated: { value: allocated }, consumed: { value: used }, availableQuantity: 100000 - used, status: 'WithinCapacity' }, payGo: { consumed: { value: pay } } }
    });
  }
  if (path === `/licensing/entitlements/${MCS}/resources`) {
    if (!url.searchParams.get('fromDate') || !url.searchParams.get('toDate')) {
      return respond(400, { error: { code: 'BadRequest', message: 'fromDate and toDate are required.' } });
    }
    const value = Object.keys(d.consumed).filter(k => d.consumed[k] > 0).map((k, i) => ({ environmentId: k, resourceId: `agent-${i + 1}`, consumed: d.consumed[k], unit: 'Count' }));
    // zweite Seite über continuationToken
    if (url.searchParams.get('continuationToken') === 'page2') {
      return respond(200, { value: value.slice(3) });
    }
    return respond(200, { value: value.slice(0, 3), continuationToken: value.length > 3 ? 'page2' : undefined });
  }
  if (path === '/licensing/allocationsByEnvironment') {
    return respond(200, Object.keys(d.alloc).map(k => ({
      environmentId: k,
      currencyAllocations: [{ currencyType: MCS, allocated: d.alloc[k].allocated, enforcementRules: [{ ruleType: 'TenantPool', enabled: d.alloc[k].pool }] }]
    })));
  }
  if (path === '/licensing/billingPolicies') {
    return respond(200, {
      value: d.plans.map(p => ({
        id: p.id, name: p.name, status: p.status, location: p.location, type: 'TenantOwned', createdOn: p.created,
        billingInstrument: p.sub ? { id: `/subscriptions/${p.sub}/resourceGroups/${p.rg}`, subscriptionId: p.sub, resourceGroup: p.rg } : undefined,
        payGoEntitlements: ['PowerApps', 'PowerAutomate', 'PowerVirtualAgent', 'Dataverse', 'PowerPages'].map(c => ({
          entitlementId: c === 'PowerVirtualAgent' ? MCS : c, productCategory: c, payAsYouGoState: p.products.indexOf(c) >= 0, value: 1.7976931348623157e308
        }))
      }))
    });
  }
  if ((m = path.match(/^\/licensing\/billingPolicies\/([^/]+)\/environments$/))) {
    const plan = d.plans.filter(p => p.id === decodeURIComponent(m![1]))[0];
    return plan ? respond(200, { value: plan.envs.map(e => ({ billingPolicyId: plan.id, environmentId: e })) }) : respond(404, { error: { code: 'NotFound', message: 'Billing policy not found' } });
  }
  if (path === '/environmentmanagement/environments') {
    const all = d.envs.map(e => ({ id: e.id, displayName: e.name, type: e.type }));
    if (url.searchParams.get('$skiptoken') === '2') {
      return respond(200, { value: all.slice(20) });
    }
    return respond(200, { value: all.slice(0, 20), '@odata.nextLink': all.length > 20 ? `https://api.powerplatform.com/environmentmanagement/environments?api-version=${version}&$skiptoken=2` : undefined });
  }
  if ((m = path.match(/^\/licensing\/environments\/([^/]+)\/entitlements$/))) {
    const e = d.envs.filter(x => x.id === decodeURIComponent(m![1]).toLowerCase())[0];
    if (!e) {
      return respond(404, { error: { code: 'EnvironmentNotFound', message: 'Environment not found' } });
    }
    const hasMcs = !!d.alloc[e.id] || d.consumed[e.id] !== undefined || d.plans.some(p => p.envs.indexOf(e.id) >= 0);
    const storage = { entitlementId: 'Database', environmentId: e.id, environmentName: e.name, environmentType: e.type, entitlement: { unit: 'MB', capacity: { consumed: { value: 512 } } } };
    return respond(200, hasMcs ? [storage, envEntitlement(e)] : [storage]);
  }
  return respond(404, { error: { code: 'NotFound', message: `Keine simulierte Route für ${path}` } });
}

export function mockPpApiGet(url: string): Promise<IMockResponse> {
  const delay = 120 + Math.random() * 260;
  return new Promise(resolve => {
    setTimeout(() => {
      (win.__coePpApiCalls = win.__coePpApiCalls || []).push(url);
      if (win.__coePpApiDeny === 'forbidden') {
        resolve(respond(403, { error: { code: 'Forbidden', message: 'The caller does not have the Power Platform Administrator role.' } }));
        return;
      }
      resolve(route(new URL(url)));
    }, delay);
  });
}

/** Simuliert die fehlende Freigabe (AADSTS65001) beim Token-Abruf. */
export function ppApiConsentError(): Error | undefined {
  return win.__coePpApiDeny === 'consent'
    ? new Error("AADSTS65001: The user or administrator has not consented to use the application with ID '08e18876-6177-487e-b8b5-cf950c1e598c' named 'SharePoint Online Web Client Extensibility'.")
    : undefined;
}
