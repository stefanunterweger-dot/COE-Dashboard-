import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Versand über Microsoft Graph im Namen des angemeldeten Benutzers (POST /me/sendMail) mit dem von
 * Microsoft für Graph empfohlenen MSGraphClientV3.
 * Benötigt die delegierte Berechtigung „Microsoft Graph – Mail.Send“ (SharePoint Admin Center → API-Zugriff).
 * Die Mails liegen danach in „Gesendete Elemente“ des Benutzers; Antworten gehen an ihn.
 */

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
  private _client: Promise<MSGraphClientV3> | undefined;

  constructor(private _factory: MSGraphClientFactory) {}

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

  private _getClient(): Promise<MSGraphClientV3> {
    if (!this._client) {
      const p = this._factory.getClient('3');
      this._client = p;
      p.catch(() => { this._client = undefined; });
    }
    return this._client;
  }

  private async _sendOnce(mail: IOutgoingMail): Promise<void> {
    let client: MSGraphClientV3;
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
        ccRecipients: mail.cc.map(address => ({ emailAddress: { address } })),
        // Antworten an das CC-Postfach (Team), nicht nur an die Person, die gerade sendet
        replyTo: mail.cc.map(address => ({ emailAddress: { address } }))
      },
      saveToSentItems: true
    };
    try {
      await client.api('/me/sendMail').version('v1.0').post(payload);
    } catch (e) {
      // GraphError: statusCode, code, message, headers
      const g = e as { statusCode?: number; code?: string; message?: string; headers?: { get?(name: string): string | null } };
      const retry = g.headers && g.headers.get ? Number(g.headers.get('Retry-After') || 0) : 0;
      const text = [g.code, g.message].filter(x => !!x && x !== 'null').join(' ') || String(e);
      throw this._toError(Number(g.statusCode || 0), text, retry > 0 ? retry * 1000 : undefined);
    }
  }

  private _toError(status: number, message: string, retryAfterMs?: number): MailError {
    if (/AADSTS65001|consent|AADSTS500011|AADSTS50105|ErrorAccessDenied|Access is denied|Authorization_RequestDenied/i.test(message) || status === 401 || status === 403) {
      return new MailError(message, 'notApproved', status);
    }
    // Nur echte Drosselung erneut versuchen. Bei 503 ohne „Retry-After“ und 504 hat Exchange die Mail oft schon angenommen –
    // ein automatischer Neuversuch würde sie doppelt zustellen (Microsoft Graph versucht ohnehin selbst schon erneut).
    if (status === 429 || (status === 503 && !!retryAfterMs)) {
      return new MailError(message, 'throttled', status, retryAfterMs);
    }
    if (status === 503 || status === 504) {
      return new MailError(`${message} – Zeitüberschreitung bei Exchange: Die Mail wurde möglicherweise trotzdem gesendet. Bitte zuerst unter „Gesendete Elemente“ nachsehen, bevor sie erneut gesendet wird.`, 'failed', status);
    }
    return new MailError(message, 'failed', status);
  }
}
