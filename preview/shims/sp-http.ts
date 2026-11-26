/* Ersatz für @microsoft/sp-http in der Vorschau: leitet Aufrufe an die simulierte Dataverse-API. */
import { IMockResponse, mockGet } from '../mock/odata';
import { mockGraphPost } from '../mock/graph';

export type HttpClientResponse = IMockResponse;

export class AadHttpClient {
  public static configurations = { v1: {} };
  constructor(private resource: string) {}
  public get(url: string, _config: unknown, options?: { headers?: { [k: string]: string } }): Promise<HttpClientResponse> {
    return mockGet(url, (options && options.headers) || {});
  }
  public post(url: string, _config: unknown, options?: { headers?: { [k: string]: string }; body?: string }): Promise<HttpClientResponse> {
    return mockGraphPost(url, (options && options.body) || '');
  }
}

export class AadHttpClientFactory {
  public getClient(resource: string): Promise<AadHttpClient> {
    return Promise.resolve(new AadHttpClient(resource));
  }
}

/** Ersatz für MSGraphClientV3: leitet POST /me/sendMail an die simulierte Graph-API (wirft wie GraphError). */
export class MSGraphClientV3 {
  public api(path: string): { version(v: string): { post(body: unknown): Promise<unknown> }; post(body: unknown): Promise<unknown> } {
    const send = (version: string, body: unknown): Promise<unknown> =>
      mockGraphPost(`https://graph.microsoft.com/${version}${path}`, JSON.stringify(body)).then(async res => {
        if (res.ok) {
          return {};
        }
        const err = JSON.parse((await res.text()) || '{}') as { error?: { code?: string; message?: string } };
        throw { statusCode: res.status, code: err.error && err.error.code, message: err.error && err.error.message, headers: res.headers };
      });
    const req = { version: (v: string) => ({ post: (b: unknown) => send(v, b) }), post: (b: unknown) => send('v1.0', b) };
    return req;
  }
}

export class MSGraphClientFactory {
  public getClient(_version: string): Promise<MSGraphClientV3> {
    return Promise.resolve(new MSGraphClientV3());
  }
}
