import { PowerPlatformApiService, PpApiError, PPAPI_VERSIONS } from '../services/PowerPlatformApiService';

/**
 * Copilot-Credits (Copilot Studio) und Abrechnungspläne – live aus der Power Platform API, ohne CoE/Dataverse.
 * Copilot-Credits heißen in der API „MCSMessages“ (Entitlement-ID).
 */
export const CREDITS_KEY = 'copilot-credits';
export const BILLING_KEY = 'billing-plans';
export const GROUP_LICENSING = 'Lizenzen & Kapazität';
export const MCS = 'MCSMessages';
/** Wie lange geladene Daten beim Wechsel zwischen den Ansichten wiederverwendet werden. */
export const CREDITS_CACHE_MS = 15 * 60 * 1000;
/** Parallele Abfragen je Environment. */
export const CREDITS_CONCURRENCY = 6;

export type CreditStatus = 'within' | 'overage' | 'covered' | '';

export interface ICreditRow {
  environmentId: string;
  name: string;
  /** Environment-Typ (deutsch) */
  type: string;
  managed?: boolean;
  /** Zugewiesene Copilot-Credits; undefined = keine Zuweisung */
  allocated?: number;
  /** Verbraucht aus Prepaid-Kapazität */
  consumed?: number;
  /** Laut API noch verfügbar (zugewiesen − verbraucht) */
  available?: number;
  status: CreditStatus;
  /** „Aus der verfügbaren Tenant-Kapazität ziehen“ (Regel TenantPool) */
  tenantPool?: boolean;
  /** Verrechnet über einen Abrechnungsplan mit Copilot Studio */
  payGo: boolean;
  /** Name des verknüpften Abrechnungsplans (auch wenn er Copilot Studio nicht verrechnet) */
  payGoPlan?: string;
  payGoConsumed?: number;
  updatedOn?: string;
  /** Details je Environment geladen? */
  detail: 'ok' | 'missing' | 'error';
}

export interface ICreditSummary {
  purchased?: number;
  allocated?: number;
  consumed?: number;
  /** Noch zuweisbar (gekauft − zugewiesen) */
  assignable?: number;
  payGoConsumed?: number;
  status?: CreditStatus;
}

export interface IBillingPlan {
  id: string;
  name: string;
  status: string;
  location: string;
  subscriptionId: string;
  resourceGroup: string;
  createdOn: string;
  products: string[];
  /** Verrechnet Copilot Studio? (Fehlen die Produktangaben, wird „ja“ angenommen.) */
  copilot: boolean;
  environments: { id: string; name: string }[];
}

export interface ICreditsData {
  rows: ICreditRow[];
  summary: ICreditSummary;
  plans: IBillingPlan[];
  plansError?: string;
  loadedAt: Date;
  /** true = alle Environments des Tenants geprüft, false = nur solche mit Zuweisung, Verbrauch oder Abrechnungsplan */
  allChecked: boolean;
  envTotal?: number;
  warnings: string[];
}

// ---------- Rohformate der API (nur die genutzten Felder) ----------
interface IValue { value?: number; autoAllocated?: number; lastUpdatedOn?: string }
interface IRule { ruleType?: string; enabled?: boolean }
interface IEnvEntitlement {
  entitlementId?: string;
  environmentId?: string;
  environmentName?: string;
  environmentType?: string;
  isManagedEnvironment?: boolean;
  entitlement?: {
    capacity?: { allocated?: IValue; consumed?: IValue; availableQuantity?: number; status?: string; enforcementRules?: IRule[] };
    payGo?: { consumed?: IValue; entitled?: IValue };
  };
}
interface ITenantEntitlement {
  entitlementId?: string;
  entitlement?: {
    capacity?: { entitled?: IValue; allocated?: IValue; consumed?: IValue; availableQuantity?: number; status?: string };
    payGo?: { consumed?: IValue };
  };
}
interface ICurrencyReport { currencyType?: string; purchased?: number; allocated?: number }
interface IAllocation { environmentId?: string; currencyAllocations?: { currencyType?: string; allocated?: number; autoAllocated?: number; enforcementRules?: IRule[] }[] }
interface IPolicy {
  id?: string; name?: string; status?: string; location?: string; createdOn?: string;
  billingInstrument?: { subscriptionId?: string; resourceGroup?: string };
  payGoEntitlements?: { entitlementId?: string; productCategory?: string; payAsYouGoState?: boolean }[];
}
interface IPolicyEnv { environmentId?: string; billingPolicyId?: string }
interface IEnvironment { id?: string; environmentId?: string; name?: string; displayName?: string; type?: string; environmentSku?: string; properties?: { displayName?: string; environmentSku?: string } }
interface IResource { environmentId?: string; consumed?: number }

const TYPE_LABEL: { [k: string]: string } = {
  production: 'Produktion', sandbox: 'Sandbox', developer: 'Developer', default: 'Default', trial: 'Testversion',
  subscriptionbasedtrial: 'Testversion (Abo)', teams: 'Teams', support: 'Support', preview: 'Vorschau', platform: 'Plattform'
};

export function typeLabel(t: string | undefined): string {
  if (!t) {
    return '';
  }
  return TYPE_LABEL[t.toLowerCase()] || t;
}

export const STATUS_LABEL: { [k in CreditStatus]: string } = {
  within: 'Im Rahmen',
  overage: 'Überschritten',
  covered: 'Überschritten (gedeckt)',
  '': '–'
};

export const STATUS_HINT: { [k in CreditStatus]: string } = {
  within: 'Verbrauch liegt innerhalb der zugewiesenen Copilot-Credits.',
  overage: 'Die Zuweisung ist aufgebraucht und nichts fängt den Mehrverbrauch auf – Agents können stoppen.',
  covered: 'Die Zuweisung ist aufgebraucht; der Mehrverbrauch wird aus der Tenant-Kapazität oder über Pay-as-you-go gedeckt.',
  '': ''
};

function statusOf(s: string | undefined): CreditStatus {
  switch ((s || '').toLowerCase()) {
    case 'withincapacity': return 'within';
    case 'overage': return 'overage';
    case 'coveredoverage': return 'covered';
    default: return '';
  }
}

const PRODUCT_LABEL: { [k: string]: string } = {
  powerapps: 'Power Apps', powerautomate: 'Power Automate', powervirtualagent: 'Copilot Studio', copilotstudio: 'Copilot Studio',
  dataverse: 'Dataverse', powerpages: 'Power Pages', d365apps: 'Dynamics 365', aibuilder: 'AI Builder', windows365: 'Windows 365'
};

export function productLabel(p: string): string {
  return PRODUCT_LABEL[p.toLowerCase()] || p;
}

export const PLAN_STATUS_LABEL: { [k: string]: string } = {
  enabled: 'Aktiv', disabled: 'Deaktiviert', disabledbysubscription: 'Durch Abo deaktiviert', suspended: 'Ausgesetzt', deleted: 'Gelöscht'
};

export function planStatusLabel(s: string): string {
  return PLAN_STATUS_LABEL[(s || '').toLowerCase()] || s.replace(/([a-z])([A-Z])/g, '$1 $2') || '–';
}

function num(v: IValue | number | undefined): number | undefined {
  if (typeof v === 'number') {
    return v;
  }
  return v && typeof v.value === 'number' ? v.value : undefined;
}

/** Riesige Platzhalterwerte der API (z. B. 1.79e308 = „unbegrenzt“) nicht als Zahl anzeigen. */
function sane(n: number | undefined): number | undefined {
  return n === undefined || !isFinite(n) || Math.abs(n) > 1e12 ? undefined : n;
}

function arrayOf<T>(x: unknown): T[] {
  if (Array.isArray(x)) {
    return x as T[];
  }
  const v = x && (x as { value?: T[] }).value;
  return Array.isArray(v) ? v : [];
}

function lower(id: string | undefined): string {
  return String(id || '').toLowerCase();
}

function settle<T>(p: Promise<T>): Promise<{ ok: true; value: T } | { ok: false; error: PpApiError }> {
  return p.then(value => ({ ok: true as const, value }), (error: PpApiError) => ({ ok: false as const, error }));
}

function isoDate(d: Date): string {
  const two = (n: number): string => (n < 10 ? '0' : '') + n;
  return `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())}`;
}

/** Läuft fn für alle Einträge mit begrenzter Parallelität. */
async function pool<T>(items: T[], size: number, fn: (item: T) => Promise<void>, stop: () => boolean): Promise<void> {
  let i = 0;
  const worker = async (): Promise<void> => {
    while (i < items.length && !stop()) {
      const item = items[i++];
      await fn(item);
    }
  };
  await Promise.all(Array.from({ length: Math.min(size, items.length) }, () => worker()));
}

export interface ILoadOptions {
  /** Alle Environments des Tenants prüfen (langsamer) */
  all: boolean;
  onProgress?: (done: number, total: number) => void;
  cancelled?: () => boolean;
}

/**
 * Lädt Copilot-Credits und Abrechnungspläne.
 * 1. Tenant-Summen, Zuweisungen je Environment, Abrechnungspläne, Environment-Namen und Verbrauch (laufender Monat) parallel
 * 2. Environments der Abrechnungspläne
 * 3. Details je relevantem Environment (Zuweisung, Verbrauch, Status, Pay-as-you-go)
 */
export async function loadCopilotCredits(api: PowerPlatformApiService, opts: ILoadOptions): Promise<ICreditsData> {
  const warnings: string[] = [];
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  const [tenantR, reportR, allocR, policyR, envR, resR] = await Promise.all([
    settle(api.get<ITenantEntitlement>(`/licensing/entitlements/${MCS}`)),
    settle(api.get<ICurrencyReport[] | { value?: ICurrencyReport[] }>('/licensing/tenantCapacity/currencyReports')),
    settle(api.getAll<IAllocation>('/licensing/allocationsByEnvironment')),
    settle(api.getAll<IPolicy>('/licensing/billingPolicies', PPAPI_VERSIONS.billing)),
    settle(api.getAll<IEnvironment>('/environmentmanagement/environments', PPAPI_VERSIONS.environments)),
    settle(api.getAll<IResource>(`/licensing/entitlements/${MCS}/resources?fromDate=${isoDate(monthStart)}&toDate=${isoDate(now)}`))
  ]);

  // Ohne Berechtigung bringt keine Abfrage etwas → Fehler nach oben (Hinweis auf Rolle/Freigabe).
  const results = [tenantR, reportR, allocR, policyR, envR, resR];
  if (results.every(r => !r.ok)) {
    const errs = results.map(r => (r as { error: PpApiError }).error);
    throw errs.filter(e => e.kind === 'consent')[0] || errs.filter(e => e.kind === 'forbidden')[0] || errs[0];
  }

  // Namen und Typen aus der Environment-Liste
  const names: { [id: string]: { name: string; type: string } } = {};
  if (envR.ok) {
    envR.value.forEach(e => {
      const id = lower(e.environmentId || e.name || e.id);
      if (id) {
        names[id] = { name: e.displayName || (e.properties && e.properties.displayName) || id, type: e.type || e.environmentSku || (e.properties && e.properties.environmentSku) || '' };
      }
    });
  } else if (envR.error.kind !== 'notFound') {
    warnings.push('Environment-Liste nicht verfügbar – Namen kommen aus den Lizenzdaten.');
  }

  // Zuweisungen
  const alloc: { [id: string]: { allocated?: number; tenantPool?: boolean } } = {};
  if (allocR.ok) {
    allocR.value.forEach(a => {
      const id = lower(a.environmentId);
      const c = (a.currencyAllocations || []).filter(x => lower(x.currencyType) === lower(MCS))[0];
      if (id && c) {
        const rule = (c.enforcementRules || []).filter(r => lower(r.ruleType) === 'tenantpool')[0];
        alloc[id] = { allocated: (c.allocated || 0) + (c.autoAllocated || 0), tenantPool: rule ? rule.enabled !== false : undefined };
      }
    });
  } else if (allocR.error.kind !== 'notFound') {
    warnings.push(`Zuweisungen je Environment nicht abrufbar (${allocR.error.message}).`);
  }

  // Verbrauch laufender Monat (je Agent/Ressource) → Environments mit Verbrauch
  const usedEnvs: { [id: string]: number } = {};
  if (resR.ok) {
    resR.value.forEach(r => {
      const id = lower(r.environmentId);
      if (id) {
        usedEnvs[id] = (usedEnvs[id] || 0) + (r.consumed || 0);
      }
    });
  }

  // Abrechnungspläne + ihre Environments
  const plans: IBillingPlan[] = [];
  const planOf: { [envId: string]: IBillingPlan } = {};
  let plansError: string | undefined;
  if (policyR.ok) {
    await Promise.all(policyR.value.map(async p => {
      const products = (p.payGoEntitlements || []).filter(x => x.payAsYouGoState).map(x => productLabel(x.productCategory || x.entitlementId || ''));
      const unique = products.filter((x, i) => x && products.indexOf(x) === i);
      const plan: IBillingPlan = {
        id: String(p.id || ''),
        name: p.name || String(p.id || ''),
        status: p.status || '',
        location: p.location || '',
        subscriptionId: (p.billingInstrument && p.billingInstrument.subscriptionId) || '',
        resourceGroup: (p.billingInstrument && p.billingInstrument.resourceGroup) || '',
        createdOn: p.createdOn || '',
        products: unique,
        copilot: !p.payGoEntitlements || !p.payGoEntitlements.length
          || p.payGoEntitlements.some(x => !!x.payAsYouGoState && (lower(x.entitlementId) === lower(MCS) || lower(x.productCategory) === 'powervirtualagent')),
        environments: []
      };
      plans.push(plan);
      if (!plan.id) {
        return;
      }
      try {
        const envs = await api.getAll<IPolicyEnv>(`/licensing/billingPolicies/${encodeURIComponent(plan.id)}/environments`, PPAPI_VERSIONS.billing);
        plan.environments = envs.map(e => ({ id: lower(e.environmentId), name: '' })).filter(e => !!e.id);
        plan.environments.forEach(e => { planOf[e.id] = plan; });
      } catch (e) {
        if ((e as PpApiError).kind !== 'notFound') {
          warnings.push(`Environments des Abrechnungsplans „${plan.name}“ nicht abrufbar.`);
        }
      }
    }));
    plans.sort((a, b) => a.name.localeCompare(b.name));
  } else {
    plansError = policyR.error.kind === 'forbidden' ? 'Für Abrechnungspläne fehlt die Berechtigung (Rolle „Power Platform Administrator“).' : policyR.error.message;
  }

  // Relevante Environments
  const ids: string[] = [];
  const add = (id: string): void => { if (id && ids.indexOf(id) < 0) { ids.push(id); } };
  Object.keys(alloc).forEach(add);
  Object.keys(planOf).forEach(add);
  Object.keys(usedEnvs).forEach(add);
  if (opts.all) {
    Object.keys(names).forEach(add);
  }

  // Details je Environment
  const details: { [id: string]: IEnvEntitlement | null | undefined } = {};
  let done = 0;
  if (opts.onProgress) {
    opts.onProgress(0, ids.length);
  }
  await pool(ids, CREDITS_CONCURRENCY, async id => {
    try {
      const list = arrayOf<IEnvEntitlement>(await api.get<unknown>(`/licensing/environments/${encodeURIComponent(id)}/entitlements`));
      details[id] = list.filter(e => lower(e.entitlementId) === lower(MCS))[0] || null;
    } catch (e) {
      details[id] = (e as PpApiError).kind === 'notFound' ? null : undefined;
    }
    done++;
    if (opts.onProgress) {
      opts.onProgress(done, ids.length);
    }
  }, () => !!opts.cancelled && opts.cancelled());
  const failed = ids.filter(id => details[id] === undefined).length;
  if (failed) {
    warnings.push(`Für ${failed} Environment(s) konnten keine Details geladen werden – angezeigt werden nur Zuweisung und Abrechnungsplan.`);
  }

  const rows: ICreditRow[] = ids.map(id => {
    const d = details[id];
    const cap = d && d.entitlement ? d.entitlement.capacity : undefined;
    const pg = d && d.entitlement ? d.entitlement.payGo : undefined;
    const a = alloc[id];
    const plan = planOf[id];
    const rule = cap && (cap.enforcementRules || []).filter(r => lower(r.ruleType) === 'tenantpool')[0];
    const allocated = sane(cap && cap.allocated ? (num(cap.allocated) || 0) + (cap.allocated.autoAllocated || 0) : a && a.allocated);
    const nm = names[id];
    const capConsumed = cap ? sane(num(cap.consumed)) : undefined;
    return {
      environmentId: id,
      name: (d && d.environmentName) || (nm && nm.name) || id,
      type: typeLabel((d ? d.environmentType : undefined) || (nm ? nm.type : undefined)),
      managed: d ? d.isManagedEnvironment : undefined,
      allocated: allocated !== undefined && allocated > 0 ? allocated : undefined,
      consumed: capConsumed !== undefined ? capConsumed : usedEnvs[id],
      available: sane(cap ? cap.availableQuantity : undefined),
      status: statusOf(cap && cap.status),
      tenantPool: rule ? rule.enabled !== false : a ? a.tenantPool : undefined,
      payGo: !!plan && plan.copilot,
      payGoPlan: plan ? plan.name : undefined,
      payGoConsumed: sane(num(pg ? pg.consumed : undefined)),
      updatedOn: (cap && cap.consumed && cap.consumed.lastUpdatedOn) || undefined,
      detail: d ? 'ok' : d === null ? 'missing' : 'error'
    };
  });

  // Namen auch in den Abrechnungsplänen
  const rowName: { [id: string]: string } = {};
  rows.forEach(r => { rowName[r.environmentId] = r.name; });
  plans.forEach(p => p.environments.forEach(e => { e.name = rowName[e.id] || (names[e.id] && names[e.id].name) || e.id; }));

  // Summen: bevorzugt die Tenant-Werte der API, sonst Währungsbericht bzw. Summe der Zeilen
  const summary: ICreditSummary = {};
  const sum = (f: (r: ICreditRow) => number | undefined): number => rows.reduce((s, r) => s + (f(r) || 0), 0);
  if (tenantR.ok && tenantR.value && tenantR.value.entitlement) {
    const cap = tenantR.value.entitlement.capacity;
    summary.purchased = sane(num(cap && cap.entitled));
    summary.allocated = sane(cap && cap.allocated ? (num(cap.allocated) || 0) + (cap.allocated.autoAllocated || 0) : undefined);
    summary.consumed = sane(num(cap && cap.consumed));
    summary.payGoConsumed = sane(num(tenantR.value.entitlement.payGo && tenantR.value.entitlement.payGo.consumed));
    summary.status = statusOf(cap && cap.status);
  }
  if (reportR.ok) {
    const rep = arrayOf<ICurrencyReport>(reportR.value).filter(r => lower(r.currencyType) === lower(MCS))[0];
    if (rep) {
      if (summary.purchased === undefined) {
        summary.purchased = sane(rep.purchased);
      }
      if (summary.allocated === undefined) {
        summary.allocated = sane(rep.allocated);
      }
    }
  }
  if (summary.allocated === undefined && rows.length) {
    summary.allocated = sum(r => r.allocated);
  }
  if (summary.consumed === undefined && rows.length) {
    summary.consumed = sum(r => r.consumed);
  }
  if (summary.payGoConsumed === undefined && rows.some(r => r.payGoConsumed !== undefined)) {
    summary.payGoConsumed = sum(r => r.payGoConsumed);
  }
  if (summary.purchased !== undefined && summary.allocated !== undefined) {
    summary.assignable = Math.max(0, summary.purchased - summary.allocated);
  }

  return {
    rows,
    summary,
    plans,
    plansError,
    loadedAt: new Date(),
    allChecked: opts.all,
    envTotal: envR.ok ? envR.value.length : undefined,
    warnings
  };
}
