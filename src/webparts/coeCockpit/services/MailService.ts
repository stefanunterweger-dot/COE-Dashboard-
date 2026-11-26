import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';

/**
 * Versand über Microsoft Graph im Namen des angemeldeten Benutzers (POST /me/sendMail).
 * Benötigt die delegierte Berechtigung „Microsoft Graph – Mail.Send“ (SharePoint Admin Center → API-Zugriff).
 * Die Mails liegen danach in „Gesendete Elemente“ des Benutzers; Antworten gehen an ihn.
 */
const GRAPH = 'https://graph.microsoft.com';

/** Abstand zwischen zwei Mails – Exchange Online erlaubt etwa 30 Nachrichten pro Minute und Postfach. */
export const MAIL_TIMING = { intervalMs: 2100, maxRetries: 3 };

export type MailErrorKind = 'notApproved' | 'throttled' | 'failed';

export class MailError extends Error {
  constructor(message: string, public kind: MailErrorKind, public status: number, public retryAfterMs?: number) {
    super(message);
    Object.setPrototypeOf(this, MailError.prototype);
  }
}

export interface IOutgoingMail {
  to: string[];
  cc: string[];
  subject: string;
  html: string;
}

export const NOT_APPROVED_HINT = 'Die Berechtigung „Microsoft Graph – Mail.Send“ ist noch nicht genehmigt. Ein SharePoint-Admin gibt sie im ' +
  'SharePoint Admin Center unter Erweitert → API-Zugriff frei. Bis dahin „Nacheinander in Outlook“ verwenden.';

const wait = (ms: number): Promise<void> => new Promise(resolve => window.setTimeout(resolve, ms));

export class MailService {
  private _client: Promise<AadHttpClient> | undefined;

  constructor(private _factory: AadHttpClientFactory) {}

  /** Eine Mail senden; bei Drosselung (429/503) wird nach „Retry-After“ erneut versucht. */
  public async send(mail: IOutgoingMail): Promise<void> {
    for (let attempt = 0; ; attempt++) {
      try {
        await this._sendOnce(mail);
        return;
      } catch (e) {
        if (e instanceof MailError && e.kind === 'throttled' && attempt < MAIL_TIMING.maxRetries) {
          await wait(e.retryAfterMs || 10000);
          continue;
        }
        throw e;
      }
    }
  }

  private _getClient(): Promise<AadHttpClient> {
    if (!this._client) {
      const p = this._factory.getClient(GRAPH);
      this._client = p;
      p.catch(() => { this._client = undefined; });
    }
    return this._client;
  }

  private async _sendOnce(mail: IOutgoingMail): Promise<void> {
    let client: AadHttpClient;
    try {
      client = await this._getClient();
    } catch (e) {
      throw this._toError(0, e instanceof Error ? e.message : String(e));
    }
    const payload = {
      message: {
        subject: mail.subject,
        body: { contentType: 'HTML', content: `<html><head><meta charset="utf-8"></head><body>${mail.html}</body></html>` },
        toRecipients: mail.to.map(address => ({ emailAddress: { address } })),
        ccRecipients: mail.cc.map(address => ({ emailAddress: { address } }))
      },
      saveToSentItems: true
    };
    let res: { ok: boolean; status: number; headers?: { get(name: string): string | null }; text(): Promise<string> };
    try {
      res = await client.post(`${GRAPH}/v1.0/me/sendMail`, AadHttpClient.configurations.v1, {
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    } catch (e) {
      throw this._toError(0, e instanceof Error ? e.message : String(e));
    }
    if (res.ok || res.status === 202) {
      return;
    }
    const text = await res.text().catch(() => '');
    let message = text;
    try {
      const body = JSON.parse(text) as { error?: { code?: string; message?: string } };
      message = body.error ? `${body.error.code || ''} ${body.error.message || ''}`.trim() : text;
    } catch {
      // kein JSON
    }
    const retry = res.headers ? Number(res.headers.get('Retry-After') || 0) : 0;
    throw this._toError(res.status, message || `HTTP ${res.status}`, retry > 0 ? retry * 1000 : undefined);
  }

  private _toError(status: number, message: string, retryAfterMs?: number): MailError {
    if (/AADSTS65001|consent|AADSTS500011|AADSTS50105|ErrorAccessDenied|Access is denied|Authorization_RequestDenied/i.test(message) || status === 401 || status === 403) {
      return new MailError(message, 'notApproved', status);
    }
    if (status === 429 || status === 503 || status === 504) {
      return new MailError(message, 'throttled', status, retryAfterMs);
    }
    return new MailError(message, 'failed', status);
  }
}
