import { IColumnFilter } from './filters';

/**
 * Teilbare Links: Reiter, Spaltenfilter, Suche und geöffneter Datensatz stehen im Adress-Anker (#coe=…).
 * Beispiel: „verwaiste Apps in HR Produktion“ als Link an Kolleg:innen schicken.
 * Andere Anker-Parameter der Seite bleiben erhalten.
 */
export interface IViewState {
  /** Reiter */
  t?: string;
  /** Spaltenfilter */
  f?: { [field: string]: IColumnFilter };
  /** Suche */
  q?: string;
  /** Geöffneter Datensatz: logischer Tabellenname + ID */
  r?: { l: string; id: string };
}

const PARAM = 'coe';

function hashParts(): string[] {
  return (window.location.hash || '').replace(/^#/, '').split('&').filter(p => !!p);
}

export function readViewState(): IViewState | undefined {
  try {
    const part = hashParts().filter(p => p.indexOf(`${PARAM}=`) === 0)[0];
    if (!part) {
      return undefined;
    }
    const parsed = JSON.parse(decodeURIComponent(part.substring(PARAM.length + 1))) as IViewState;
    return parsed && typeof parsed === 'object' ? parsed : undefined;
  } catch {
    return undefined;
  }
}

/** Schreibt den Zustand in die Adresse, ohne einen Verlaufseintrag anzulegen. */
export function writeViewState(state: IViewState): void {
  try {
    const clean: IViewState = {};
    if (state.t) { clean.t = state.t; }
    if (state.f && Object.keys(state.f).length) { clean.f = state.f; }
    if (state.q) { clean.q = state.q; }
    if (state.r) { clean.r = state.r; }
    const others = hashParts().filter(p => p.indexOf(`${PARAM}=`) !== 0);
    const own = Object.keys(clean).length ? [`${PARAM}=${encodeURIComponent(JSON.stringify(clean))}`] : [];
    const hash = others.concat(own).join('&');
    const url = `${window.location.pathname}${window.location.search}${hash ? `#${hash}` : ''}`;
    if (url !== `${window.location.pathname}${window.location.search}${window.location.hash}`) {
      window.history.replaceState(window.history.state, '', url);
    }
  } catch {
    // Adresse lässt sich nicht ändern (z. B. eingebettet) – Links sind dann nicht teilbar.
  }
}

/** Kopiert den aktuellen Link in die Zwischenablage. */
export function copyCurrentLink(): Promise<boolean> {
  const nav = navigator as Navigator & { clipboard?: { writeText: (t: string) => Promise<void> } };
  if (nav.clipboard) {
    return nav.clipboard.writeText(window.location.href).then(() => true, () => false);
  }
  return Promise.resolve(false);
}
