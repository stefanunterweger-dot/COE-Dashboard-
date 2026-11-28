import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * „Was ändert Microsoft?“: Meldungen aus dem Microsoft 365 Message Center (Microsoft Graph
 * /admin/serviceAnnouncement/messages, Berechtigung „ServiceMessage.Read.All“), gefiltert auf die Power Platform.
 */
export const MESSAGES_KEY = 'microsoft-changes';

/** Dienste, die zur Power Platform gehören (wie sie in den Meldungen heißen). */
export const PP_SERVICES = [
  'Power Apps', 'Microsoft Power Apps', 'Power Automate', 'Microsoft Power Automate', 'Power Platform', 'Microsoft Power Platform',
  'Microsoft Copilot Studio', 'Copilot Studio', 'Microsoft Dataverse', 'Dataverse', 'Dynamics 365 Apps', 'Power Pages', 'Microsoft Power Pages',
  'Power BI', 'Microsoft Power BI'
];
// „DLP“ allein nicht: sonst kämen die (häufigen) Meldungen zu Microsoft Purview DLP mit
const PP_TITLE = /power apps|power automate|power platform|copilot studio|dataverse|power pages|power bi|managed environment|\bdlp\b.*(connector|konnektor)/i;

export interface IMcMessage {
  id: string;
  title: string;
  services: string[];
  category: string;
  severity: string;
  major: boolean;
  start: string;
  end: string;
  modified: string;
  /** Handlung nötig bis */
  actionBy: string;
  tags: string[];
  html: string;
  links: { label: string; url: string }[];
}

export interface IMcData { messages: IMcMessage[]; totalAll: number; loadedAt: Date }

export class MessageLoadError extends Error {
  constructor(message: string, public kind: 'consent' | 'forbidden' | 'failed', public detail?: string) {
    super(message);
    Object.setPrototypeOf(this, MessageLoadError.prototype);
  }
}

export const CATEGORY_LABEL: { [c: string]: string } = { planForChange: 'Änderung planen', stayInformed: 'Zur Information', preventOrFixIssue: 'Problem vermeiden/beheben' };

let cache: { at: number; promise: Promise<IMcData> } | undefined;
export function resetMessageCache(): void {
  cache = undefined;
}

export function loadMessages(factory: MSGraphClientFactory, fresh: boolean = false): Promise<IMcData> {
  if (!fresh && cache && Date.now() - cache.at < 10 * 60000) {
    return cache.promise;
  }
  const p = load(factory);
  cache = { at: Date.now(), promise: p };
  p.catch(() => { cache = undefined; });
  return p;
}

async function load(factory: MSGraphClientFactory): Promise<IMcData> {
  const graph: MSGraphClientV3 = await factory.getClient('3');
  const all: { [k: string]: unknown }[] = [];
  let url: string | undefined = '/admin/serviceAnnouncement/messages?$top=100';
  for (let page = 0; url && page < 30; page++) {
    let res: { value?: { [k: string]: unknown }[]; '@odata.nextLink'?: string };
    try {
      res = await graph.api(url).version('v1.0').get() as typeof res;
    } catch (e) {
      const err = (e || {}) as { statusCode?: number; message?: string };
      const msg = err.message || String(e);
      if (/AADSTS65001|consent/i.test(msg)) {
        throw new MessageLoadError('Microsoft Graph „ServiceMessage.Read.All“ ist für SharePoint noch nicht freigegeben (Skript -MessageCenter).', 'consent', msg);
      }
      if (err.statusCode === 401 || err.statusCode === 403 || /Authorization_RequestDenied|Forbidden|UnknownError/i.test(msg)) {
        throw new MessageLoadError('Zugriff verweigert: Es fehlt die Berechtigung „ServiceMessage.Read.All“ (Skript -MessageCenter) oder deine Rolle darf das Message Center nicht lesen (z. B. Power Platform Administrator, Globaler Leser, Leseberechtigter für Nachrichtencenter).', 'forbidden', msg);
      }
      throw new MessageLoadError(`Message Center nicht lesbar: ${msg}`, 'failed', msg);
    }
    (res.value || []).forEach(v => all.push(v));
    url = res['@odata.nextLink'];
  }
  const messages = all.map(toMessage).filter(m => m.services.some(s => PP_SERVICES.indexOf(s) >= 0) || PP_TITLE.test(m.title));
  return { messages, totalAll: all.length, loadedAt: new Date() };
}

const str = (v: unknown): string => (v === undefined || v === null ? '' : String(v));

function toMessage(v: { [k: string]: unknown }): IMcMessage {
  const body = (v.body || {}) as { content?: string };
  const details = (Array.isArray(v.details) ? v.details : []) as { name?: string; value?: string }[];
  const links: { label: string; url: string }[] = [];
  details.forEach(d => {
    const value = str(d.value);
    if (/^https:\/\//i.test(value)) {
      links.push({ label: /blog/i.test(str(d.name)) ? 'Blog' : /roadmap/i.test(str(d.name)) ? 'Roadmap' : 'Weitere Informationen', url: value });
    } else if (/roadmapids/i.test(str(d.name)) && value) {
      value.split(/[,;\s]+/).filter(x => /^\d+$/.test(x)).slice(0, 3).forEach(id => links.push({ label: `Roadmap ${id}`, url: `https://www.microsoft.com/microsoft-365/roadmap?featureid=${id}` }));
    }
  });
  return {
    id: str(v.id),
    title: str(v.title),
    services: (Array.isArray(v.services) ? v.services : []).map(str),
    category: str(v.category),
    severity: str(v.severity),
    major: v.isMajorChange === true,
    start: str(v.startDateTime),
    end: str(v.endDateTime),
    modified: str(v.lastModifiedDateTime),
    actionBy: str(v.actionRequiredByDateTime),
    tags: (Array.isArray(v.tags) ? v.tags : []).map(str),
    html: str(body.content),
    links
  };
}
