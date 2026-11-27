/* Simulierter Versand über Microsoft Graph (POST /me/sendMail) für die Vorschau. */
import { IMockResponse } from './odata';
import { GRAPH_GROUPS, GRAPH_PEOPLE } from './data';

export interface ISentMail { at: string; subject: string; to: string[]; cc: string[]; html: string }

interface IGraphWindow {
  /** E-Mail-Adresse → UPN (Personen, deren E-Mail-Adresse nicht der UPN ist) */
  __coeMailAliases?: { [mail: string]: string };
  /** Versendete Mails (zum Prüfen in Tests). */
  __sentMails?: ISentMail[];
  /** true = Berechtigung Mail.Send nicht genehmigt. */
  __coeDenyMailSend?: boolean;
  __graphCalls?: number;
  /** true = Berechtigung User.Read.All / GroupMember.Read.All nicht genehmigt. */
  __coeDenyUserRead?: boolean;
  __graphBatchCalls?: number;
  /** true = /subscribedSkus nicht erlaubt (Rückfall über licenseDetails) */
  __coeDenySkus?: boolean;
  /** true = Message Center nicht lesbar (ServiceMessage.Read.All fehlt) */
  __coeDenyMessages?: boolean;
}

/** Feste SKU-IDs der Beispiel-Lizenzen (wie in Entra: GUID je Produkt). */
const SKU_IDS: { [part: string]: string } = {
  SPE_E3: '05e9a617-0261-4cee-bb44-138d3ef5d965',
  POWERAPPS_PER_USER: 'b30411f5-fea1-4a59-9ad9-3db7c7ead579',
  POWERAUTOMATE_ATTENDED_RPA: 'eda1941c-3c4f-4995-b5eb-e85a42175ab9',
  DYN365_ENTERPRISE_SALES: '1e1a282c-9c54-43a2-9310-98ef728faace',
  POWERAPPS_VIRAL: 'dcb1a3ae-b33f-4487-846a-a640262fadf4',
  POWER_BI_PRO: 'f8a1db68-be16-40ed-86d5-cb42ce701560',
  Microsoft_365_Copilot: '639dec6b-bb19-468b-871c-c5c441c4b0cb'
};
const DEPARTMENTS = ['Vertrieb', 'Finanzen', 'IT', 'Personal', 'Einkauf', 'Marketing'];

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
  const sku = (part: string, plans: string[]): unknown => ({ skuId: SKU_IDS[part], skuPartNumber: part, servicePlans: plans.map(p => ({ servicePlanName: p, provisioningStatus: 'Success' })) });
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

/** Beispiel-Meldungen des Message Centers (Power Platform und andere Dienste). */
function messages(): { [k: string]: unknown }[] {
  const m = (id: string, title: string, services: string[], category: string, opts: { major?: boolean; start: string; modified?: string; action?: string; html?: string; roadmap?: string }) => ({
    id, title, services, category, severity: opts.major ? 'high' : 'normal', isMajorChange: !!opts.major,
    startDateTime: opts.start, endDateTime: '2027-06-30T00:00:00Z', lastModifiedDateTime: opts.modified || opts.start,
    actionRequiredByDateTime: opts.action || null, tags: opts.major ? ['Admin impact', 'Feature update'] : ['Updated message'],
    body: { contentType: 'html', content: opts.html || `<p><b>${title}</b></p><p>Beispieltext der Meldung. <a href="https://learn.microsoft.com/power-platform/">Mehr erfahren</a></p><script>alert(1)</script>` },
    details: opts.roadmap ? [{ name: 'RoadmapIds', value: opts.roadmap }, { name: 'ExternalLink', value: 'https://learn.microsoft.com/power-platform/admin/' }] : []
  });
  return [
    m('MC1180001', 'Premium-Lizenz für Apps in Managed Environments ab Februar 2027', ['Power Apps', 'Microsoft Power Platform'], 'planForChange',
      { major: true, start: '2026-03-12T10:00:00Z', modified: '2026-11-20T08:00:00Z', action: '2027-02-01T00:00:00Z', roadmap: '498765' }),
    m('MC1180002', 'Nicht verwaltete Developer-Environments werden ab 15. Dezember 2026 gesperrt', ['Microsoft Power Platform'], 'planForChange',
      { major: true, start: '2026-06-02T10:00:00Z', modified: '2026-11-25T08:00:00Z', action: '2026-12-15T00:00:00Z' }),
    m('MC1180003', 'Copilot Studio: neue Abrechnung für Copilot-Credits', ['Microsoft Copilot Studio'], 'stayInformed', { start: '2026-11-24T10:00:00Z' }),
    m('MC1180004', 'Power Automate: Ende der Unterstützung für Flows mit veralteten SharePoint-Triggern', ['Power Automate'], 'planForChange',
      { start: '2026-10-01T10:00:00Z', action: '2027-04-30T00:00:00Z' }),
    m('MC1180005', 'Dataverse: Änderungen bei der Speicherberechnung für Log-Kapazität', ['Microsoft Dataverse'], 'stayInformed', { start: '2026-09-15T10:00:00Z' }),
    m('MC1180006', 'Power Pages: Authentifizierung mit lokalen Konten wird entfernt', ['Power Pages'], 'preventOrFixIssue',
      { major: true, start: '2026-08-01T10:00:00Z', action: '2026-11-01T00:00:00Z' }),
    m('MC1180007', 'Power BI: neue Ansicht für Arbeitsbereiche', ['Power BI'], 'stayInformed', { start: '2026-11-22T10:00:00Z' }),
    m('MC1180008', 'Exchange Online: Änderungen an Postfachrichtlinien', ['Exchange Online'], 'stayInformed', { start: '2026-11-23T10:00:00Z' }),
    m('MC1180009', 'Microsoft Teams: neues Besprechungslayout', ['Microsoft Teams'], 'stayInformed', { start: '2026-11-21T10:00:00Z' }),
    m('MC1180010', 'SharePoint Online: neue App-Leiste', ['SharePoint Online'], 'planForChange', { start: '2026-10-10T10:00:00Z', action: '2027-01-15T00:00:00Z' })
  ];
}

/** Alle Benutzer mit zugewiesenen Lizenzen (wie /users?$select=…,assignedLicenses). */
function allUsers(): { [k: string]: unknown }[] {
  // Personen mit Premium-Lizenz ohne Nutzung (kommen in den CoE-Daten nicht vor) – für die Lizenz-Optimierung
  const idle = [
    ['Paula Premium', 'POWERAPPS_PER_USER', 'Finanzen'], ['Peter Planlos', 'POWERAPPS_PER_USER', 'Einkauf'], ['Petra Pause', 'POWERAPPS_PER_USER', 'Marketing'],
    ['Pius Probe', 'POWERAPPS_PER_USER', 'IT'], ['Frieda Flowfrei', 'POWERAUTOMATE_ATTENDED_RPA', 'Personal'], ['Felix Flowlos', 'POWERAUTOMATE_ATTENDED_RPA', 'Vertrieb']
  ].map(([name, sku, dept], i) => ({
    id: `7d1e0000-0000-4000-8000-00000000000${i + 1}`, displayName: name, userPrincipalName: `${name.toLowerCase().replace(' ', '.')}@contoso.com`, mail: `${name.toLowerCase().replace(' ', '.')}@contoso.com`,
    department: dept, jobTitle: '', accountEnabled: true, userType: 'Member', assignedLicenses: [{ skuId: SKU_IDS.SPE_E3, disabledPlans: [] }, { skuId: SKU_IDS[sku], disabledPlans: [] }]
  }));
  return idle.concat(GRAPH_PEOPLE.map(p => {
    const h = hash(p.upn.toLowerCase());
    const lic = licensesFor(p.upn);
    const skus = lic.status === 200 ? ((lic.body as { value: { skuId: string }[] }).value || []).map(l => ({ skuId: l.skuId, disabledPlans: [] as string[] })) : [];
    if (lic.status === 200 && h % 5 === 0) { skus.push({ skuId: SKU_IDS.POWER_BI_PRO, disabledPlans: [] }); }
    if (lic.status === 200 && h % 9 === 0) { skus.push({ skuId: SKU_IDS.Microsoft_365_Copilot, disabledPlans: [] }); }
    return {
      id: p.id, displayName: p.name, userPrincipalName: p.upn, mail: p.upn, department: DEPARTMENTS[h % DEPARTMENTS.length],
      jobTitle: '', accountEnabled: h % 11 !== 0, userType: h % 13 === 0 ? 'Guest' : 'Member', assignedLicenses: skus
    };
  }) as typeof idle);
}

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
    // Person über die E-Mail-Adresse suchen ($filter=mail eq '…' or userPrincipalName eq '…')
    const byMail = /[?&]\$filter=([^&]+)/.exec(url);
    if (/\/v1\.0\/users\?/.test(url) && byMail && /mail eq/.test(decodeURIComponent(byMail[1]))) {
      const wanted = (/mail eq '((?:[^']|'')*)'/.exec(decodeURIComponent(byMail[1])) || [])[1] || '';
      const key = wanted.replace(/''/g, "'").toLowerCase();
      resolve(respond(200, { value: GRAPH_PEOPLE.filter(x => x.upn.toLowerCase() === key || (w.__coeMailAliases || {})[key] === x.upn.toLowerCase()).map(x => ({ id: x.id })) }));
      return;
    }
    // Benutzer mit zugewiesenen Lizenzen (Seiten zu 40, damit das Weiterblättern mitgetestet wird)
    if (/\/v1\.0\/users\?/.test(url) || /\/v1\.0\/users$/.test(url)) {
      if (w.__coeDenyUserRead) {
        resolve(respond(403, { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } }));
        return;
      }
      const skip = Number((/[?&]\$skiptoken=(\d+)/.exec(url) || [])[1] || 0);
      const all = allUsers();
      const page = all.slice(skip, skip + 40);
      resolve(respond(200, { value: page, '@odata.nextLink': skip + 40 < all.length ? `https://graph.microsoft.com/v1.0/users?$select=id&$top=999&$skiptoken=${skip + 40}` : undefined }));
      return;
    }
    if (/\/v1\.0\/admin\/serviceAnnouncement\/messages/.test(url)) {
      if (w.__coeDenyMessages) {
        resolve(respond(403, { error: { code: 'Forbidden', message: 'Insufficient privileges to complete the operation.' } }));
        return;
      }
      resolve(respond(200, { value: messages() }));
      return;
    }
    if (/\/v1\.0\/subscribedSkus/.test(url)) {
      if (w.__coeDenySkus) {
        resolve(respond(403, { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } }));
        return;
      }
      const counts: { [id: string]: number } = {};
      allUsers().forEach(u => (u.assignedLicenses as { skuId: string }[]).forEach(l => { counts[l.skuId] = (counts[l.skuId] || 0) + 1; }));
      resolve(respond(200, { value: Object.keys(SKU_IDS).map(part => ({
        skuId: SKU_IDS[part], skuPartNumber: part, capabilityStatus: 'Enabled', appliesTo: 'User',
        consumedUnits: counts[SKU_IDS[part]] || 0, prepaidUnits: { enabled: (counts[SKU_IDS[part]] || 0) + (part === 'SPE_E3' ? 25 : 5), suspended: 0, warning: 0 }
      })) }));
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
