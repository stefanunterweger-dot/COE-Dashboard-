/* Simulierter Versand über Microsoft Graph (POST /me/sendMail) für die Vorschau. */
import { IMockResponse } from './odata';

export interface ISentMail { at: string; subject: string; to: string[]; cc: string[]; html: string }

interface IGraphWindow {
  /** Versendete Mails (zum Prüfen in Tests). */
  __sentMails?: ISentMail[];
  /** true = Berechtigung Mail.Send nicht genehmigt. */
  __coeDenyMailSend?: boolean;
  __graphCalls?: number;
}

function respond(status: number, body: unknown, headers: { [k: string]: string } = {}): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (name: string) => headers[name] || headers[name.toLowerCase()] || null },
    json: () => Promise.resolve(text ? JSON.parse(text) : {}),
    text: () => Promise.resolve(text)
  };
}

export function mockGraphPost(url: string, body: string): Promise<IMockResponse> {
  const w = window as unknown as IGraphWindow;
  return new Promise(resolve => setTimeout(() => {
    if (!/\/v1\.0\/me\/sendMail$/.test(url)) {
      resolve(respond(404, { error: { code: 'ResourceNotFound', message: `Unsupported ${url}` } }));
      return;
    }
    if (w.__coeDenyMailSend) {
      resolve(respond(403, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } }));
      return;
    }
    w.__graphCalls = (w.__graphCalls || 0) + 1;
    // Einmal Drosselung simulieren (3. Aufruf), damit der erneute Versuch mitgetestet wird
    if (w.__graphCalls === 3) {
      resolve(respond(429, { error: { code: 'ApplicationThrottled', message: 'Too many requests' } }, { 'Retry-After': '1' }));
      return;
    }
    const msg = (JSON.parse(body) as { message: { subject: string; body: { content: string }; toRecipients: { emailAddress: { address: string } }[]; ccRecipients: { emailAddress: { address: string } }[] } }).message;
    w.__sentMails = (w.__sentMails || []).concat([{
      at: new Date().toISOString(),
      subject: msg.subject,
      to: msg.toRecipients.map(r => r.emailAddress.address),
      cc: msg.ccRecipients.map(r => r.emailAddress.address),
      html: msg.body.content
    }]);
    resolve(respond(202, undefined));
  }, 120));
}
