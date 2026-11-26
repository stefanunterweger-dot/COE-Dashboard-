import { AadTokenProviderFactory } from '@microsoft/sp-http';

/**
 * Prüft beim Start, ob die angemeldete Person die Rolle „Power Platform Administrator“ (bzw. „Globaler Administrator“)
 * gerade AKTIV hat – bei PIM also aktiviert. Grundlage ist der Claim „wids“ (aktive Entra-Rollen) im Zugriffstoken,
 * das SharePoint für das Webpart ohnehin ausstellt. Es wird nichts an einen Dienst gesendet, keine zusätzliche Berechtigung.
 * Wichtig: Tokens werden bis zu einer Stunde zwischengespeichert – „Erneut prüfen“ holt ein frisches Token.
 */
export const ROLE_PP_ADMIN = '11648597-926c-4cf3-9c36-bcebb0ba8dcc';
export const ROLE_GLOBAL_ADMIN = '62e90394-69f5-4237-9190-012177145e10';

/** Rollen-Vorlagen-IDs → Anzeigename (nur die für das Cockpit relevanten). */
export const ROLE_NAMES: { [templateId: string]: string } = {
  [ROLE_PP_ADMIN]: 'Power Platform Administrator',
  [ROLE_GLOBAL_ADMIN]: 'Globaler Administrator',
  '44367163-eba1-44c3-98af-f5787879f96a': 'Dynamics 365 Administrator',
  'f28a1f50-f6e7-4571-818b-6a12f2af6b6c': 'SharePoint Administrator',
  'f2ef992c-3afb-46b9-b7cf-a126ee74c451': 'Globaler Leser'
};

export type RoleState = 'active' | 'inactive' | 'unknown';

export interface IRoleStatus {
  state: RoleState;
  /** Aktive, für das Cockpit relevante Rollen (Anzeigenamen) */
  roles: string[];
  checkedAt: Date;
  /** Für welche Ressource das geprüfte Token ausgestellt war */
  source?: string;
  /** Grund, wenn nicht prüfbar */
  problem?: string;
}

interface IJwtPayload { wids?: string[]; exp?: number }

/** Mittelteil eines JWT lesen (nur zur Anzeige, keine Signaturprüfung nötig). */
export function decodeJwt(token: string): IJwtPayload | undefined {
  const part = token.split('.')[1];
  if (!part) {
    return undefined;
  }
  try {
    const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64 + '==='.slice((b64.length + 3) % 4);
    const bin = atob(padded);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) {
      bytes[i] = bin.charCodeAt(i);
    }
    return JSON.parse(new TextDecoder().decode(bytes)) as IJwtPayload;
  } catch {
    return undefined;
  }
}

/**
 * Rolle prüfen. resources: Ressourcen, für die das Webpart Tokens bekommt (z. B. die CoE-Dataverse-URL);
 * das erste Token mit „wids“ entscheidet. fresh = Token neu anfordern statt aus dem Zwischenspeicher.
 */
export async function checkAdminRole(factory: AadTokenProviderFactory, resources: string[], fresh: boolean): Promise<IRoleStatus> {
  const provider = await factory.getTokenProvider();
  let lastProblem = '';
  for (const resource of resources.filter((r, i) => !!r && resources.indexOf(r) === i)) {
    try {
      const token = await provider.getToken(resource, !fresh);
      const payload = decodeJwt(token);
      if (!payload || !Array.isArray(payload.wids)) {
        lastProblem = 'Das Token enthält keine Rolleninformation.';
        continue;
      }
      const wids = payload.wids.map(w => String(w).toLowerCase());
      const roles = Object.keys(ROLE_NAMES).filter(id => wids.indexOf(id) >= 0).map(id => ROLE_NAMES[id]);
      const active = wids.indexOf(ROLE_PP_ADMIN) >= 0 || wids.indexOf(ROLE_GLOBAL_ADMIN) >= 0;
      return { state: active ? 'active' : 'inactive', roles, checkedAt: new Date(), source: resource };
    } catch (e) {
      lastProblem = e instanceof Error ? e.message : String(e);
    }
  }
  return { state: 'unknown', roles: [], checkedAt: new Date(), problem: lastProblem || 'Kein Token verfügbar.' };
}
