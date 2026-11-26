/* Simuliertes Power-Platform-Inventar (resourcequery/resources/query) und weitere Live-Schnittstellen für die Vorschau.
   Die Daten kommen aus den simulierten CoE-Tabellen – ohne Abweichungen ist der Live-Abgleich also „alles aktuell“.
   window.__coeLiveDrift = true erzeugt Abweichungen: 2 Apps fehlen live, 3 neue Apps, 1 App live in Quarantäne,
   1 App umbenannt, 1 Developer-Environment live Managed. window.__coeInvDeny = 'consent' | 'forbidden' simuliert Fehler. */
import { COE_HOST, DATABASES } from './data';
import { IMockResponse } from './odata';

interface IWin { __coeLiveDrift?: boolean; __coeInvDeny?: '' | 'consent' | 'forbidden'; __coeInvCalls?: number }
const w = window as unknown as IWin;

function respond(status: number, body: unknown): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: () => null },
    json: () => Promise.resolve(text ? JSON.parse(text) : undefined),
    text: () => Promise.resolve(text)
  };
}

const coe = (): { [t: string]: { rows: Record<string, unknown>[] } } => DATABASES.filter(d => d.host === COE_HOST)[0].tables as unknown as { [t: string]: { rows: Record<string, unknown>[] } };
const lower = (v: unknown): string => (v === undefined || v === null ? '' : String(v)).toLowerCase();
const alive = (r: Record<string, unknown>, del: string): boolean => r.statecode !== 1 && r[del] !== true;

interface IItem { name: string; type: string; location: string; properties: { [k: string]: unknown } }

const CONNECTORS = [
  { connectorId: 'shared_sharepointonline', operations: [{ operationId: 'GetItems' }, { operationId: 'PostItem' }] },
  { connectorId: 'shared_office365users', operations: [{ operationId: 'MyProfile_V2' }] },
  { connectorId: 'shared_commondataserviceforapps', operations: [] }
];

function hash(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) {
    h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  }
  return Math.abs(h);
}

/** Zusatzfelder eines Environments für „environmentmanagement/environments“ (Zustand, Schutzstufe, URL). */
export function envLiveFields(id: string): { [k: string]: unknown } {
  const r = coe().admin_environment.rows.filter(x => lower(x.admin_environmentname) === id)[0];
  if (!r) {
    return {};
  }
  let level = String(r.admin_governanceprotectionlevel || '');
  if (w.__coeLiveDrift && driftEnvId() === id) {
    level = 'Standard';
  }
  return {
    state: /disabled/i.test(String(r.admin_environmentruntimestate || '')) ? 'Disabled' : 'Ready',
    protectionLevel: level,
    url: r.admin_environmentcdsinstanceurl || '',
    azureRegion: r.admin_region || 'europe',
    createdDateTime: r.admin_environmentcreatedon || ''
  };
}

/** Developer-Environment, das bei Abweichungen live schon Managed ist. */
function driftEnvId(): string {
  const r = coe().admin_environment.rows.filter(x => alive(x, 'admin_environmentdeleted') && /developer/i.test(String(x.admin_environmentsku || ''))
    && (!x.admin_governanceprotectionlevel || /basic/i.test(String(x.admin_governanceprotectionlevel))))[0];
  return r ? lower(r.admin_environmentname) : '';
}

function items(): IItem[] {
  const t = coe();
  const out: IItem[] = [];
  const apps = t.admin_app.rows.filter(r => alive(r, 'admin_appdeleted'));
  apps.forEach((r, i) => {
    if (w.__coeLiveDrift && (i === 0 || i === 1)) {
      return; // live gelöscht
    }
    let name = String(r.admin_displayname || '');
    let quarantined = r.admin_appisquarantined === true;
    if (w.__coeLiveDrift && i === 2) { name += ' (umbenannt)'; }
    if (w.__coeLiveDrift && i === 3) { quarantined = !quarantined; }
    out.push({
      name: lower(r.admin_appid), type: /model/i.test(String(r.admin_powerappstype || '')) ? 'microsoft.powerapps/modeldrivenapps' : 'microsoft.powerapps/canvasapps', location: 'europe',
      properties: {
        displayName: name, environmentId: lower(r.admin_appenvironmentid), ownerId: lower(r._admin_appowner_value), createdBy: lower(r._admin_appowner_value),
        createdAt: r.admin_appcreatedon, lastModifiedAt: r.admin_appmodifiedon, lastModifiedBy: lower(r._admin_appowner_value), isQuarantined: quarantined
      }
    });
  });
  if (w.__coeLiveDrift && apps.length) {
    const tmpl = apps[5] || apps[0];
    [1, 2, 3].forEach(n => out.push({
      name: `0fa1e000-0000-4000-8000-00000000000${n}`, type: 'microsoft.powerapps/canvasapps', location: 'europe',
      properties: {
        displayName: `Neue Live-App ${n}`, environmentId: lower(tmpl.admin_appenvironmentid), ownerId: lower(tmpl._admin_appowner_value), createdBy: lower(tmpl._admin_appowner_value),
        createdAt: '2026-11-26T08:0' + n + ':00Z', lastModifiedAt: '2026-11-26T08:1' + n + ':00Z', isQuarantined: false
      }
    }));
  }
  t.admin_flow.rows.filter(r => alive(r, 'admin_flowdeleted')).forEach(r => {
    const id = lower(r.admin_flowidstring || r.admin_flowid);
    out.push({
      name: id, type: 'microsoft.powerautomate/cloudflows', location: 'europe',
      properties: {
        displayName: r.admin_displayname, environmentId: lower(r.admin_flowenvironmentid), ownerId: lower(r._admin_derivedowner_value || r._admin_flowcreator_value),
        createdBy: lower(r._admin_flowcreator_value), createdAt: r.admin_flowcreatedon, lastModifiedAt: r.admin_flowmodifiedon,
        workflowEntityId: `9${id.slice(1)}`, trigger: hash(id) % 3 === 0 ? 'sharepointonline' : '', triggerOperation: hash(id) % 3 === 0 ? 'GetOnNewItems' : ''
      }
    });
  });
  t.admin_environment.rows.filter(r => alive(r, 'admin_environmentdeleted')).forEach(r => {
    const id = lower(r.admin_environmentname);
    const level = String(envLiveFields(id).protectionLevel || '');
    out.push({
      name: id, type: 'microsoft.powerplatform/environments', location: 'europe',
      properties: { displayName: r.admin_displayname, environmentType: r.admin_environmentsku, isManaged: level ? !/basic/i.test(level) : undefined, createdAt: r.admin_environmentcreatedon }
    });
  });
  (t.admin_pva ? t.admin_pva.rows : []).filter(r => r.admin_pvadeleted !== true).forEach(r => out.push({
    name: lower(r.admin_pvaid), type: 'microsoft.copilotstudio/agents', location: 'europe',
    properties: { displayName: r.admin_name, environmentId: lower(r.admin_pvaenvironmentid || r._admin_pvaenvironment_value), createdAt: r.admin_pvacreatedon }
  }));
  (t.admin_portal ? t.admin_portal.rows : []).forEach(r => out.push({
    name: lower(r.admin_portalid), type: 'microsoft.powerpages/sites', location: 'europe',
    properties: { displayName: r.admin_portaldisplayname, status: 'Online', siteVisibility: 'private' }
  }));
  return out;
}

const unquote = (v: string): string => v.replace(/^'(.*)'$/, '$1').toLowerCase();

function fieldOf(it: IItem, field: string): string {
  if (field === 'type') { return it.type; }
  if (field === 'name') { return it.name; }
  const m = /^properties\.(.+)$/.exec(field);
  return m ? lower(it.properties[m[1]]) : '';
}

interface IClause { $type: string; FieldName?: string; Operator?: string; Values?: string[]; TakeCount?: number; SummarizeClauseExpression?: { FieldList?: string[]; OperatorFieldName?: string } }

export function mockInventoryQuery(body: string): Promise<IMockResponse> {
  return new Promise(resolve => setTimeout(() => {
    w.__coeInvCalls = (w.__coeInvCalls || 0) + 1;
    if (w.__coeInvDeny === 'forbidden') {
      resolve(respond(403, { error: { code: 'Forbidden', message: 'The user does not have a supported role to query inventory.' } }));
      return;
    }
    if (w.__coeInvDeny === 'consent') {
      resolve(respond(403, { error: { code: 'InsufficientScopes', message: "The access token is missing the required scope 'ResourceQuery.Resources.Read'." } }));
      return;
    }
    const req = JSON.parse(body || '{}') as { TableName?: string; Clauses?: IClause[]; Options?: { Top?: number; SkipToken?: string } };
    if (req.TableName !== 'PowerPlatformResources') {
      resolve(respond(400, { error: { code: 'BadRequest', message: 'Unknown table' } }));
      return;
    }
    let rows = items();
    let summarize: IClause | undefined;
    let take: number | undefined;
    (req.Clauses || []).forEach(c => {
      if (c.$type === 'where' && c.FieldName) {
        const vals = (c.Values || []).map(unquote);
        rows = rows.filter(it => vals.indexOf(fieldOf(it, c.FieldName!)) >= 0);
      } else if (c.$type === 'summarize') {
        summarize = c;
      } else if (c.$type === 'take') {
        take = c.TakeCount;
      }
    });
    if (summarize) {
      const counts: { [t: string]: number } = {};
      rows.forEach(r => { counts[r.type] = (counts[r.type] || 0) + 1; });
      const data = Object.keys(counts).map(t => ({ type: t, resourceCount: counts[t] }));
      resolve(respond(200, { totalRecords: data.length, count: data.length, resultTruncated: 'false', data }));
      return;
    }
    if (take !== undefined) {
      rows = rows.slice(0, take);
    }
    // Einzelabfrage (where name) mit Connectors wie im echten Inventar
    if ((req.Clauses || []).some(c => c.$type === 'where' && c.FieldName === 'name')) {
      rows = rows.map(r => ({ ...r, properties: { ...r.properties, powerPlatformConnectors: /powerapps|powerautomate/.test(r.type) ? CONNECTORS.slice(0, 1 + (hash(r.name) % 3)) : undefined } }));
    }
    const top = Math.min(1000, (req.Options && req.Options.Top) || 1000);
    const skip = Number((req.Options && req.Options.SkipToken) || 0) || 0;
    const page = rows.slice(skip, skip + top);
    const next = skip + top < rows.length ? String(skip + top) : undefined;
    resolve(respond(200, { totalRecords: rows.length, count: page.length, resultTruncated: next ? 'true' : 'false', skipToken: next, data: page }));
  }, 180 + Math.random() * 200));
}

/** Letzte Flow-Ausführungen (powerautomate/environments/{env}/flowRuns). */
export function mockFlowRuns(workflowId: string): IMockResponse {
  const h = hash(workflowId);
  const value = [0, 1, 2, 3, 4, 5].map(i => {
    const start = new Date(Date.UTC(2026, 10, 26, 8, 0, 0) - i * 3600000 * (1 + (h % 5)));
    const status = (h + i) % 7 === 0 ? 'Failed' : 'Succeeded';
    return { flowRunId: `run-${h}-${i}`, name: `08584${h}${i}`, status, startTime: start.toISOString(), endTime: new Date(start.getTime() + 4200).toISOString(), durationMs: 4200 + i * 300, triggerType: 'Automated', workflowId };
  });
  return respond(200, { value });
}

/** Freigaben einer App (Admin-Schnittstelle) bzw. Flow-Zustand. */
export function mockAdminGet(url: string): IMockResponse | undefined {
  const u = new URL(url);
  let m = /\/apps\/([^/]+)\/permissions$/.exec(u.pathname);
  if (m) {
    const id = decodeURIComponent(m[1]).toLowerCase();
    const app = coe().admin_app.rows.filter(r => lower(r.admin_appid) === id)[0];
    const owner = app ? lower(app._admin_appowner_value) : '';
    const maker = (coe().admin_maker ? coe().admin_maker.rows : []).filter(r => lower(r.admin_makerid) === owner)[0];
    return respond(200, {
      value: [
        { name: owner, properties: { roleName: 'Owner', principal: { id: owner, displayName: String((maker && maker.admin_displayname) || 'Besitzer'), email: String((maker && maker.admin_userprincipalname) || ''), type: 'User' } } },
        { name: 'g1', properties: { roleName: 'CanView', principal: { id: '1b2c3d4e-0000-4000-8000-0000000000a1', displayName: 'Vertrieb Innendienst', email: 'vertrieb-innendienst@contoso.com', type: 'Group' } } },
        { name: 'u1', properties: { roleName: 'CanEdit', principal: { id: '1b2c3d4e-0000-4000-8000-0000000000a2', displayName: 'Lisa Wimmer', email: 'lisa.wimmer@contoso.com', type: 'User' } } }
      ]
    });
  }
  m = /\/flows\/([^/]+)$/.exec(u.pathname);
  if (m) {
    const id = decodeURIComponent(m[1]).toLowerCase();
    const flow = coe().admin_flow.rows.filter(r => lower(r.admin_flowidstring || r.admin_flowid) === id)[0];
    return flow ? respond(200, { name: id, properties: { displayName: flow.admin_displayname, state: flow.admin_flowstate || 'Started' } }) : respond(404, { error: { code: 'FlowNotFound', message: 'Flow not found' } });
  }
  return undefined;
}
