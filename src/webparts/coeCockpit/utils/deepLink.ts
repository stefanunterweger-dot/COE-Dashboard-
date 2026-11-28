import { DATE_PRESETS, IColumnFilter } from './filters';

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
    return sanitize(JSON.parse(decodeURIComponent(part.substring(PARAM.length + 1))));
  } catch {
    return undefined;
  }
}

/** Nur gültige Teile übernehmen – ein verstümmelter Link darf das Webpart nicht zum Absturz bringen. */
function sanitize(v: unknown): IViewState | undefined {
  if (!v || typeof v !== 'object' || Array.isArray(v)) {
    return undefined;
  }
  const raw = v as { t?: unknown; f?: unknown; q?: unknown; r?: unknown };
  const out: IViewState = {};
  if (typeof raw.t === 'string') {
    out.t = raw.t;
  }
  if (typeof raw.q === 'string') {
    out.q = raw.q;
  }
  if (raw.f && typeof raw.f === 'object' && !Array.isArray(raw.f)) {
    const f: { [field: string]: IColumnFilter } = {};
    const src = raw.f as { [k: string]: unknown };
    Object.keys(src).forEach(k => {
      const c = src[k] as { values?: unknown; date?: unknown } | undefined;
      if (!c || typeof c !== 'object') {
        return;
      }
      if (Array.isArray(c.values) && c.values.every(x => typeof x === 'string')) {
        f[k] = { values: c.values as string[] };
      } else if (typeof c.date === 'string' && DATE_PRESETS.some(p => p.key === c.date)) {
        f[k] = { date: c.date as IColumnFilter['date'] };
      }
    });
    if (Object.keys(f).length) {
      out.f = f;
    }
  }
  const r = raw.r as { l?: unknown; id?: unknown } | undefined;
  if (r && typeof r === 'object' && typeof r.l === 'string' && /^[a-z0-9_]+$/i.test(r.l)
    && typeof r.id === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(r.id)) {
    out.r = { l: r.l, id: r.id };
  }
  return out;
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
