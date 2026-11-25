/* Ersatz für @microsoft/sp-http in der Vorschau: leitet Aufrufe an die simulierte Dataverse-API. */
import { IMockResponse, mockGet } from '../mock/odata';

export type HttpClientResponse = IMockResponse;

export class AadHttpClient {
  public static configurations = { v1: {} };
  constructor(private resource: string) {}
  public get(url: string, _config: unknown, options?: { headers?: { [k: string]: string } }): Promise<HttpClientResponse> {
    return mockGet(url, (options && options.headers) || {});
  }
}

export class AadHttpClientFactory {
  public getClient(resource: string): Promise<AadHttpClient> {
    return Promise.resolve(new AadHttpClient(resource));
  }
}
