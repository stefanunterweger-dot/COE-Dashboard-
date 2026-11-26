/* Simulierter Versand über Microsoft Graph (POST /me/sendMail) für die Vorschau. */
import { IMockResponse } from './odata';
import { GRAPH_GROUPS, GRAPH_PEOPLE } from './data';

export interface ISentMail { at: string; subject: string; to: string[]; cc: string[]; html: string }

interface IGraphWindow {
  /** Versendete Mails (zum Prüfen in Tests). */
  __sentMails?: ISentMail[];
  /** true = Berechtigung Mail.Send nicht genehmigt. */
  __coeDenyMailSend?: boolean;
  __graphCalls?: number;
  /** true = Berechtigung User.Read.All / GroupMember.Read.All nicht genehmigt. */
  __coeDenyUserRead?: boolean;
  __graphBatchCalls?: number;
}

function hash(s: string): number {
  let h = 7;
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) >>> 0;
  }
  return h;
}

/** Beispiel-Lizenzen je Person (deterministisch). */
function licensesFor(user: string): { status: number; body: unknown } {
  const h = hash(user.toLowerCase()) % 16;
  const sku = (part: string, plans: string[]): unknown => ({ skuPartNumber: part, servicePlans: plans.map(p => ({ servicePlanName: p, provisioningStatus: 'Success' })) });
  const e3 = sku('SPE_E3', ['POWERAPPS_O365_P2', 'FLOW_O365_P2', 'EXCHANGE_S_ENTERPRISE', 'SHAREPOINTENTERPRISE']);
  if (h === 15) {
    return { status: 404, body: { error: { code: 'Request_ResourceNotFound', message: `Resource '${user}' does not exist.` } } };
  }
  if (h < 7) {
    return { status: 200, body: { value: [e3, sku('POWERAPPS_PER_USER', ['POWERAPPS_PER_USER', 'DYN365_CDS_P2'])] } };
  }
  if (h === 7) {
    return { status: 200, body: { value: [e3, sku('POWERAUTOMATE_ATTENDED_RPA', ['POWER_AUTOMATE_ATTENDED_RPA', 'FLOW_PER_USER'])] } };
  }
  if (h === 8) {
    return { status: 200, body: { value: [e3, sku('DYN365_ENTERPRISE_SALES', ['POWERAPPS_DYN_APPS', 'FLOW_DYN_APPS'])] } };
  }
  if (h === 9) {
    return { status: 200, body: { value: [e3, sku('POWERAPPS_VIRAL', ['POWERAPPS_P2_VIRAL', 'FLOW_P2_VIRAL'])] } };
  }
  return { status: 200, body: { value: [e3] } };
}

// Für Tests: einige Personen, die Graph kennt (E-Mail → Objekt-ID)
(window as unknown as { __graphPeople?: unknown }).__graphPeople = GRAPH_PEOPLE.slice(0, 5);

/** GET (z. B. transitive Gruppenmitglieder) */
export function mockGraphGet(url: string): Promise<IMockResponse> {
  const w = window as unknown as IGraphWindow;
  return new Promise(resolve => setTimeout(() => {
    const user = /\/v1\.0\/users\/([^/?]+)$/.exec(url.split('?')[0]);
    if (user) {
      const key = decodeURIComponent(user[1]).toLowerCase();
      const p = GRAPH_PEOPLE.filter(x => x.upn.toLowerCase() === key || x.id.toLowerCase() === key)[0];
      resolve(p ? respond(200, { id: p.id }) : respond(404, { error: { code: 'Request_ResourceNotFound', message: `Resource '${key}' does not exist.` } }));
      return;
    }
    const m = /\/groups\/([^/]+)\/transitiveMembers/.exec(url);
    if (!m) {
      resolve(respond(404, { error: { code: 'ResourceNotFound', message: `Unsupported ${url}` } }));
      return;
    }
    if (w.__coeDenyUserRead) {
      resolve(respond(403, { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } }));
      return;
    }
    const gid = decodeURIComponent(m[1]).toLowerCase();
    // inkl. verschachtelter Gruppen (transitiv): rund ein Viertel der Personen
    const value = GRAPH_PEOPLE.filter(p => hash(gid + p.upn) % 4 === 0).map(p => ({ id: p.id, userPrincipalName: p.upn, displayName: p.name, mail: p.upn }));
    resolve(respond(200, { value }));
  }, 150));
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
    if (/\/v1\.0\/\$batch$/.test(url)) {
      w.__graphBatchCalls = (w.__graphBatchCalls || 0) + 1;
      const req = JSON.parse(body) as { requests: { id: string; url: string }[] };
      const responses = req.requests.map(r => {
        if (w.__coeDenyUserRead) {
          return { id: r.id, status: 403, body: { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } } };
        }
        const u = /\/users\/([^/]+)\/licenseDetails/.exec(r.url);
        const id = u ? decodeURIComponent(u[1]) : '';
        const person = GRAPH_PEOPLE.filter(p => p.id.toLowerCase() === id.toLowerCase() || p.upn.toLowerCase() === id.toLowerCase())[0];
        const res = licensesFor(person ? person.upn : id);
        return { id: r.id, status: res.status, body: res.body };
      });
      resolve(respond(200, { responses }));
      return;
    }
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
