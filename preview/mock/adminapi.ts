/* Simulierte Admin-Schnittstellen für direkte Aktionen: Power Apps (api.powerapps.com), Power Automate (api.flow.microsoft.com)
   und die schreibenden Environment-Aufrufe der Power Platform API. Jeder Aufruf wird in window.__coeAdminCalls festgehalten. */
import { IMockResponse } from './odata';

interface IAdminWindow {
  /** 'forbidden' = Rolle nicht aktiv, 'consent' = Berechtigung für SharePoint fehlt, 'busy' = anderer Vorgang läuft (PPAPI) */
  __coeAdminDeny?: '' | 'forbidden' | 'consent' | 'busy';
  __coeAdminCalls?: { method: string; url: string; body: unknown }[];
}
const w = window as unknown as IAdminWindow;

function respond(status: number, body?: unknown): IMockResponse {
  const text = body === undefined ? '' : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: () => null },
    json: () => Promise.resolve(text ? JSON.parse(text) : undefined),
    text: () => Promise.resolve(text)
  };
}

/** Fehlende Freigabe beim Token-Abruf (wie AadHttpClient bei AADSTS65001). */
export function adminConsentError(resource: string): Error | undefined {
  return w.__coeAdminDeny === 'consent'
    ? new Error(`AADSTS65001: The user or administrator has not consented to use the application with ID '08e18876-6177-487e-b8b5-cf950c1e598c' named 'SharePoint Online Web Client Extensibility'. Send an interactive authorization request for this user and resource. Resource: ${resource}`)
    : undefined;
}

const ROUTES: RegExp[] = [
  /^\/providers\/Microsoft\.PowerApps\/scopes\/admin\/environments\/[^/]+\/apps\/[^/]+(\/(setPowerAppQuarantineState|modifyAppOwner|modifyPermissions))?$/i,
  /^\/providers\/Microsoft\.ProcessSimple\/scopes\/admin\/environments\/[^/]+\/flows\/[^/]+(\/(start|stop|modifyOwners))?$/i,
  /^\/environmentmanagement\/environments\/[^/]+\/(governanceSetting\/enableManaged|Disable|Enable)$/i
];

export function mockAdminCall(method: string, url: string, body?: string): Promise<IMockResponse> {
  return new Promise(resolve => setTimeout(() => {
    const u = new URL(url);
    const parsed = body ? JSON.parse(body) as unknown : undefined;
    (w.__coeAdminCalls = w.__coeAdminCalls || []).push({ method, url, body: parsed });
    if (w.__coeAdminDeny === 'forbidden') {
      resolve(respond(403, { error: { code: 'Forbidden', message: 'The caller does not have permission to perform this action (Power Platform Administrator role required).' } }));
      return;
    }
    if (w.__coeAdminDeny === 'busy' && /powerplatform/.test(u.hostname)) {
      resolve(respond(400, { error: { code: 'EnvironmentLifecycleOperationInProgress', message: 'There is an active lifecycle operation on this environment.' } }));
      return;
    }
    if (!ROUTES.some(r => r.test(u.pathname))) {
      resolve(respond(404, { error: { code: 'NotFound', message: `Keine simulierte Route für ${u.pathname}` } }));
      return;
    }
    if (/powerplatform/.test(u.hostname)) {
      resolve(respond(202));
      return;
    }
    resolve(respond(method === 'DELETE' ? 204 : 200, method === 'DELETE' ? undefined : { put: [] }));
  }, 260));
}
