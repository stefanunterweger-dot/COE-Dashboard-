/* Ersatz für @microsoft/sp-http in der Vorschau: leitet Aufrufe an die simulierte Dataverse-API. */
import { IMockResponse, mockGet, mockPatch } from '../mock/odata';
import { adminConsentError, mockAdminCall } from '../mock/adminapi';
import { mockGraphGet, mockGraphPost } from '../mock/graph';
import { mockSharePoint } from '../mock/sharepoint';
import { mockPpApiGet, ppApiConsentError } from '../mock/ppapi';

export type HttpClientResponse = IMockResponse;

export class AadHttpClient {
  public static configurations = { v1: {} };
  constructor(private resource: string) {}
  public get(url: string, _config: unknown, options?: { headers?: { [k: string]: string } }): Promise<HttpClientResponse> {
    if (/^https:\/\/api\.powerplatform\.com\//i.test(url)) {
      const consent = ppApiConsentError();
      return consent ? Promise.reject(consent) : mockPpApiGet(url);
    }
    return mockGet(url, (options && options.headers) || {});
  }
  public post(url: string, _config: unknown, options?: { headers?: { [k: string]: string }; body?: string }): Promise<HttpClientResponse> {
    if (/^https:\/\/api\.powerplatform\.com\//i.test(url)) {
      const consent = adminConsentError(this.resource) || ppApiConsentError();
      return consent ? Promise.reject(consent) : mockAdminCall('POST', url, options && options.body);
    }
    return mockGraphPost(url, (options && options.body) || '');
  }
  /** Allgemeiner Aufruf (Methode in options) – Admin-Schnittstellen und PATCH ins CoE. */
  public fetch(url: string, _config: unknown, options?: { method?: string; headers?: { [k: string]: string }; body?: string }): Promise<HttpClientResponse> {
    const method = ((options && options.method) || 'GET').toUpperCase();
    if (/^https:\/\/api\.(powerapps\.com|flow\.microsoft\.com)\//i.test(url)) {
      const consent = adminConsentError(this.resource);
      return consent ? Promise.reject(consent) : mockAdminCall(method, url, options && options.body);
    }
    if (method === 'PATCH') {
      return mockPatch(url, (options && options.body) || '{}');
    }
    return mockGet(url, (options && options.headers) || {});
  }
}

export class AadHttpClientFactory {
  public getClient(resource: string): Promise<AadHttpClient> {
    return Promise.resolve(new AadHttpClient(resource));
  }
}

/** Ersatz für MSGraphClientV3: leitet Aufrufe an die simulierte Graph-API (wirft wie GraphError). */
export class MSGraphClientV3 {
  public api(path: string): { version(v: string): { post(body: unknown): Promise<unknown>; get(): Promise<unknown>; select(s: string): { get(): Promise<unknown> } }; post(body: unknown): Promise<unknown>; get(): Promise<unknown> } {
    const full = (version: string): string => (/^https:/.test(path) ? path : `https://graph.microsoft.com/${version}${path}`);
    const handle = async (res: IMockResponse): Promise<unknown> => {
      const text = await res.text();
      if (res.ok) {
        return text ? JSON.parse(text) : {};
      }
      const err = JSON.parse(text || '{}') as { error?: { code?: string; message?: string } };
      throw { statusCode: res.status, code: err.error && err.error.code, message: err.error && err.error.message, headers: res.headers };
    };
    const send = (version: string, body: unknown): Promise<unknown> => mockGraphPost(full(version), JSON.stringify(body)).then(handle);
    const get = (version: string): Promise<unknown> => mockGraphGet(full(version)).then(handle);
    return {
      version: (v: string) => ({ post: (b: unknown) => send(v, b), get: () => get(v), select: () => ({ get: () => get(v) }) }),
      post: (b: unknown) => send('v1.0', b),
      get: () => get('v1.0')
    };
  }
}

export class MSGraphClientFactory {
  public getClient(_version: string): Promise<MSGraphClientV3> {
    return Promise.resolve(new MSGraphClientV3());
  }
}

/** Ersatz für SPHttpClient (SharePoint-REST der eigenen Website). */
export class SPHttpClient {
  public static configurations = { v1: {} };
  public get(url: string, _config: unknown, _options?: unknown): Promise<HttpClientResponse> {
    return mockSharePoint('GET', url);
  }
  public post(url: string, _config: unknown, options?: { body?: string; headers?: { [k: string]: string } }): Promise<HttpClientResponse> {
    return mockSharePoint('POST', url, options && options.body, options && options.headers);
  }
}

/** Ersatz für AadTokenProviderFactory: liefert ein Beispiel-Token mit den aktiven Rollen (Claim „wids“). */
export class AadTokenProviderFactory {
  public getTokenProvider(): Promise<{ getToken(resource: string, useCachedToken?: boolean): Promise<string> }> {
    return Promise.resolve({
      getToken: (_resource: string) => {
        const inactive = !!(window as unknown as { __coeRoleInactive?: boolean }).__coeRoleInactive;
        const wids = ['b79fbf4d-3ef9-4689-8143-76b194e85509'].concat(inactive ? [] : ['11648597-926c-4cf3-9c36-bcebb0ba8dcc']);
        const b64 = (o: unknown): string => btoa(JSON.stringify(o)).replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
        return new Promise<string>(resolve => setTimeout(() => resolve(`${b64({ alg: 'none' })}.${b64({ wids, exp: Math.floor(Date.now() / 1000) + 3600 })}.x`), 300));
      }
    });
  }
}
