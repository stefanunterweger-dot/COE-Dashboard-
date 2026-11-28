import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';

/**
 * Live-Zugriff auf die Power Platform API (https://api.powerplatform.com) im Kontext des angemeldeten Benutzers.
 * Kein Dataverse, kein CoE: Die Daten kommen direkt vom Dienst. Voraussetzungen:
 *  1. Das SPFx-Prinzipal darf die „Power Platform API“ aufrufen (Freigabe der Berechtigungen, siehe README).
 *  2. Der angemeldete Benutzer hat die Rolle „Power Platform Administrator“ (oder „Globaler Administrator“) aktiv –
 *     bei PIM also vorher aktiviert.
 */
export const PPAPI_RESOURCE = 'https://api.powerplatform.com';
export const PPAPI_BASE = 'https://api.powerplatform.com';

/** API-Versionen je Bereich: die erste wird versucht, bei „unbekannter api-version“ die nächste. */
export const PPAPI_VERSIONS = {
  licensing: ['2024-10-01', '2022-03-01-preview'],
  billing: ['2022-03-01-preview', '2024-10-01'],
  environments: ['2022-03-01-preview', '2024-10-01'],
  /** Environment-Aktionen (Managed, Deaktivieren, Aktivieren) */
  environmentActions: ['2024-10-01', '2022-03-01-preview']
};

export const PIM_URL = 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles';

export type PpApiErrorKind = 'consent' | 'forbidden' | 'notFound' | 'throttled' | 'failed';

export class PpApiError extends Error {
  constructor(public kind: PpApiErrorKind, message: string, public status?: number, public detail?: string) {
    super(message);
  }
}

export const PPAPI_FORBIDDEN_HINT = 'Die Power Platform API hat den Zugriff verweigert. Für diese Ansicht brauchst du die aktive Rolle „Power Platform Administrator“ (oder „Globaler Administrator“). '
  + 'Aktiviere die Rolle in Privileged Identity Management (PIM) und lade die Seite danach neu. Wurde sie gerade erst aktiviert, kann es einige Minuten dauern – notfalls einmal ab- und wieder anmelden.';

export const PPAPI_CONSENT_HINT = 'Das Webpart darf die Power Platform API noch nicht aufrufen. Ein Administrator muss die Berechtigungen „Licensing.Allocations.Read“, „Licensing.BillingPolicies.Read“ und '
  + '„EnvironmentManagement.Environments.Read“ der „Power Platform API“ für SharePoint freigeben (Skript im README, Abschnitt „Copilot-Credits“).';

export const PPAPI_WRITE_CONSENT_HINT = 'Das Webpart darf Environments über die Power Platform API noch nicht ändern. Ein Administrator muss die Berechtigungen „EnvironmentManagement.Write“ und '
  + '„EnvironmentManagement.Settings.ReadWrite“ der „Power Platform API“ für SharePoint freigeben (scripts/Grant-CoeCockpitPermissions.ps1 -Actions).';

/** Fehlt dem Token eine Berechtigung (Scope) – im Unterschied zu einer fehlenden Rolle? */
function isScopeProblem(msg: string): boolean {
  return /\bscopes?\b|\bscp\b|required claims|insufficient claims|not authorized for (the )?(scope|permission)/i.test(msg);
}

interface IPaged<T> { value?: T[]; '@odata.nextLink'?: string; nextLink?: string; continuationToken?: string; continuationtoken?: string }

function classifyTokenError(e: unknown): PpApiError {
  const m = typeof e === 'object' && e !== null ? (e as { message?: unknown }).message : undefined;
  const msg: string = typeof m === 'string' && m ? m : String(e);
  if (/AADSTS(65001|650057|500011|50105|90094|65004)|consent|resource principal|not been granted|has not consented/i.test(msg)) {
    return new PpApiError('consent', PPAPI_CONSENT_HINT, undefined, msg);
  }
  return new PpApiError('failed', `Anmeldung an der Power Platform API fehlgeschlagen: ${msg}`, undefined, msg);
}

function messageOf(text: string): string {
  try {
    const body = JSON.parse(text) as { error?: { message?: string; code?: string }; message?: string };
    return (body.error && (body.error.message || body.error.code)) || body.message || text;
  } catch {
    return text;
  }
}

function wait(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export class PowerPlatformApiService {
  private clientPromise: Promise<AadHttpClient> | undefined;

  constructor(private factory: AadHttpClientFactory) {}

  private client(): Promise<AadHttpClient> {
    if (!this.clientPromise) {
      this.clientPromise = this.factory.getClient(PPAPI_RESOURCE);
      this.clientPromise.catch(() => { this.clientPromise = undefined; });
    }
    return this.clientPromise;
  }

  private async raw(url: string, body?: unknown): Promise<HttpClientResponse> {
    let client: AadHttpClient;
    try {
      client = await this.client();
    } catch (e) {
      throw classifyTokenError(e);
    }
    try {
      if (body !== undefined) {
        return await client.post(url, AadHttpClient.configurations.v1, {
          headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
          body: JSON.stringify(body)
        });
      }
      return await client.get(url, AadHttpClient.configurations.v1, { headers: { Accept: 'application/json' } });
    } catch (e) {
      throw classifyTokenError(e);
    }
  }

  /** GET (bzw. POST mit body) mit Wiederholung bei 429/503. */
  private async fetchJson<T>(url: string, body?: unknown): Promise<T> {
    for (let attempt = 0; ; attempt++) {
      const res = await this.raw(url, body);
      if (res.ok) {
        if (res.status === 204) {
          return undefined as unknown as T;
        }
        const text = await res.text();
        return (text ? JSON.parse(text) : undefined) as T;
      }
      const text = await res.text();
      if ((res.status === 429 || res.status === 503) && attempt < 3) {
        const after = Number(res.headers && res.headers.get('Retry-After')) || 0;
        await wait(Math.min(after > 0 ? after * 1000 : 1500 * (attempt + 1), 10000));
        continue;
      }
      const msg = messageOf(text);
      if (res.status === 401 || res.status === 403) {
        if (body !== undefined && isScopeProblem(msg)) {
          throw new PpApiError('consent', PPAPI_WRITE_CONSENT_HINT, res.status, msg);
        }
        throw new PpApiError('forbidden', PPAPI_FORBIDDEN_HINT, res.status, msg);
      }
      if (res.status === 404) {
        throw new PpApiError('notFound', msg || 'Nicht gefunden', 404, msg);
      }
      if (res.status === 429) {
        throw new PpApiError('throttled', 'Die Power Platform API drosselt gerade die Anfragen – bitte in einer Minute erneut versuchen.', 429, msg);
      }
      throw new PpApiError('failed', `Power Platform API: ${res.status} ${msg}`.trim(), res.status, msg);
    }
  }

  private url(path: string, version: string): string {
    return `${PPAPI_BASE}${path}${path.indexOf('?') >= 0 ? '&' : '?'}api-version=${encodeURIComponent(version)}`;
  }

  /** Einzelne Ressource. versions: siehe PPAPI_VERSIONS. */
  public get<T>(path: string, versions: string[] = PPAPI_VERSIONS.licensing): Promise<T> {
    return this.call<T>(path, versions);
  }

  /**
   * Schreibender Aufruf (POST) im Namen des angemeldeten Benutzers. Microsoft prüft dabei selbst,
   * ob die Person die nötige Rolle hat (Power Platform Administrator).
   */
  public post<T>(path: string, body: unknown, versions: string[] = PPAPI_VERSIONS.environmentActions): Promise<T> {
    return this.call<T>(path, versions, body === undefined ? {} : body);
  }

  /** Aufruf mit Fallback auf die nächste api-version. */
  private async call<T>(path: string, versions: string[], body?: unknown): Promise<T> {
    let last: PpApiError | undefined;
    for (let i = 0; i < versions.length; i++) {
      try {
        return await this.fetchJson<T>(this.url(path, versions[i]), body);
      } catch (e) {
        const err = e as PpApiError;
        const versionProblem = err.kind === 'failed' && err.status === 400 && /api.?version/i.test(err.detail || '');
        if (!versionProblem || i === versions.length - 1) {
          throw err;
        }
        last = err;
      }
    }
    throw last || new PpApiError('failed', 'Power Platform API: keine passende api-version');
  }

  /** Liste – folgt @odata.nextLink / nextLink bzw. continuationToken (max. 50 Seiten). Akzeptiert auch reine Arrays. */
  public async getAll<T>(path: string, versions: string[] = PPAPI_VERSIONS.licensing): Promise<T[]> {
    const first = await this.get<IPaged<T> | T[]>(path, versions);
    if (Array.isArray(first)) {
      return first;
    }
    const out: T[] = (first && first.value) || [];
    let page: IPaged<T> | undefined = first;
    for (let n = 0; n < 50 && page; n++) {
      // Folgelink in jeder Schreibweise („@odata.nextLink“, „@odata.nextlink“, „nextLink“)
      const p: IPaged<T> = page;
      const nextKey = Object.keys(p).filter(k => /^(@odata\.)?nextlink$/i.test(k))[0];
      const next = nextKey ? String((p as { [k: string]: unknown })[nextKey] || '') : '';
      const token = page.continuationToken || page.continuationtoken;
      let url = '';
      if (next) {
        url = /^https:\/\//i.test(next) ? next : `${PPAPI_BASE}${next}`;
      } else if (token) {
        // Pfad + bereits funktionierende api-version + Token
        url = `${this.url(path, versions[0])}&continuationToken=${encodeURIComponent(token)}`;
      } else {
        break;
      }
      if (!/^https:\/\/api\.powerplatform\.com\//i.test(url)) {
        break;
      }
      page = await this.fetchJson<IPaged<T>>(url);
      if (page && page.value) {
        out.push(...page.value);
      }
    }
    return out;
  }
}
