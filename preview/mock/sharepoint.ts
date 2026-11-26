/* Simulierte SharePoint-REST-API für das Versandprotokoll (Liste „CoE Cockpit Benachrichtigungen“). */
import { IMockResponse } from './odata';

interface IStore { exists: boolean; fields: string[]; items: Record<string, unknown>[] }
interface ISpWindow { __spStore?: IStore; __spReset?: () => void }

const w = window as unknown as ISpWindow;
const store = (): IStore => (w.__spStore = w.__spStore || { exists: false, fields: [], items: [] });
w.__spReset = () => { w.__spStore = { exists: false, fields: [], items: [] }; };

/** Vorbelegung: ein Environment, das ein anderer Admin schon verständigt hat. */
export function seedNotifyEntry(item: Record<string, unknown>): void {
  const s = store();
  s.exists = true;
  s.fields = ['EnvironmentId', 'SentAt', 'SentBy', 'SentByEmail', 'Method', 'Recipients', 'Subject'];
  s.items.push(item);
}

function respond(status: number, body?: unknown): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return { ok: status >= 200 && status < 300, status, json: () => Promise.resolve(text ? JSON.parse(text) : {}), text: () => Promise.resolve(text) };
}

const notFound = (): IMockResponse => respond(404, { 'odata.error': { code: '-1, System.ArgumentException', message: { lang: 'de-DE', value: 'Die Liste ist auf der Website nicht vorhanden.' } } });

export function mockSharePoint(method: 'GET' | 'POST', url: string, body?: string): Promise<IMockResponse> {
  return new Promise(resolve => setTimeout(() => {
    const s = store();
    const u = decodeURIComponent(url);
    if (method === 'POST' && /\/_api\/web\/lists$/.test(u)) {
      s.exists = true;
      resolve(respond(201, { Id: 'f0000000-0000-4000-8000-000000000001', Title: JSON.parse(body || '{}').Title }));
      return;
    }
    if (!/GetByTitle\('CoE Cockpit Benachrichtigungen'\)/.test(u)) {
      resolve(respond(400, { 'odata.error': { message: { value: `Unsupported ${u}` } } }));
      return;
    }
    if (!s.exists) {
      resolve(notFound());
      return;
    }
    if (method === 'POST' && /\/fields\/CreateFieldAsXml$/.test(u)) {
      const xml = (JSON.parse(body || '{}') as { parameters: { SchemaXml: string } }).parameters.SchemaXml;
      s.fields.push((/Name="([^"]+)"/.exec(xml) || [])[1] || '');
      resolve(respond(201, {}));
      return;
    }
    if (method === 'POST' && /\/items$/.test(u)) {
      const item = JSON.parse(body || '{}') as Record<string, unknown>;
      const unknown = Object.keys(item).filter(k => k !== 'Title' && s.fields.indexOf(k) < 0);
      if (unknown.length) {
        resolve(respond(400, { 'odata.error': { message: { value: `Die Eigenschaft '${unknown[0]}' ist nicht vorhanden.` } } }));
        return;
      }
      s.items.push(item);
      resolve(respond(201, item));
      return;
    }
    if (method === 'GET' && /\/items\?/.test(u)) {
      const items = s.items.slice().sort((a, b) => String(b.SentAt).localeCompare(String(a.SentAt)));
      resolve(respond(200, { value: items }));
      return;
    }
    if (method === 'GET') {
      resolve(respond(200, { Id: 'f0000000-0000-4000-8000-000000000001' }));
      return;
    }
    resolve(respond(400, {}));
  }, 90));
}
