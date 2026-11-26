import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { withAlpha } from '../utils/format';

type Dict = { [key: string]: string | undefined };

/** Mischt eine Hex-Farbe mit Schwarz (amount 0…1) – für dunkle Verlaufsfarben unabhängig vom Theme. */
function darken(color: string, amount: number): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(color.trim());
  if (!m) {
    return color;
  }
  const n = parseInt(m[1], 16);
  const ch = (shift: number): string => {
    const v = Math.round(((n >> shift) & 255) * (1 - amount));
    return (v < 16 ? '0' : '') + v.toString(16);
  };
  return `#${ch(16)}${ch(8)}${ch(0)}`;
}

/** Leitet die CSS-Variablen des Webparts aus dem SharePoint-Theme ab. */
export function buildThemeVars(theme: IReadonlyTheme | undefined): React.CSSProperties {
  const p = ((theme && theme.palette) || {}) as Dict;
  const s = ((theme && theme.semanticColors) || {}) as Dict;
  const dark = !!(theme && theme.isInverted);

  const accent = p.themePrimary || '#0f6cbd';
  const success = dark ? '#54b054' : '#0e700e';
  const warning = dark ? '#f2c661' : '#8a5300';
  const danger = dark ? '#f1707b' : '#b10e1c';
  const info = dark ? '#62abf5' : '#0f6cbd';

  const vars: { [key: string]: string } = {
    '--coe-accent': accent,
    '--coe-accent-dark': p.themeDarker || p.themeDark || accent,
    '--coe-accent-soft': withAlpha(accent, dark ? 0.22 : 0.09),
    '--coe-accent-border': withAlpha(accent, 0.45),
    '--coe-text': s.bodyText || (dark ? '#f5f5f5' : '#242424'),
    '--coe-text-2': s.bodySubtext || (dark ? '#c7c7c7' : '#616161'),
    '--coe-text-3': p.neutralTertiary || (dark ? '#8a8a8a' : '#a19f9d'),
    '--coe-bg': s.bodyBackground || (dark ? '#1f1f1f' : '#ffffff'),
    '--coe-surface': p.neutralLighterAlt || (dark ? '#262626' : '#faf9f8'),
    '--coe-hover': p.neutralLighter || (dark ? '#2e2e2e' : '#f3f2f1'),
    '--coe-border': p.neutralLight || (dark ? '#3b3b3b' : '#e1dfdd'),
    '--coe-border-soft': p.neutralLighter || (dark ? '#2e2e2e' : '#f0efee'),
    '--coe-success': success,
    '--coe-success-bg': withAlpha(success, dark ? 0.2 : 0.1),
    '--coe-warning': warning,
    '--coe-warning-bg': withAlpha(dark ? warning : '#f7b500', dark ? 0.2 : 0.16),
    '--coe-danger': danger,
    '--coe-danger-bg': withAlpha(danger, dark ? 0.2 : 0.09),
    '--coe-info': info,
    '--coe-info-bg': withAlpha(info, dark ? 0.2 : 0.09),
    // Diagramme: Kategorien in fester Reihenfolge (für Hell und Dunkel getrennt gewählt, farbfehlsichtig geprüft)
    '--coe-series-1': dark ? '#338ced' : '#257de0',
    '--coe-series-2': dark ? '#e76624' : '#f47230',
    '--coe-series-3': dark ? '#00a888' : '#00b894',
    '--coe-series-4': dark ? '#d28e05' : '#f2a600',
    '--coe-series-5': dark ? '#e24d8b' : '#ee7fac',
    '--coe-series-6': dark ? '#0aac23' : '#008a17',
    '--coe-series-7': dark ? '#9887ef' : '#6a4ce2',
    '--coe-series-8': dark ? '#eb6263' : '#ec4647',
    // Status: nur wo die Farbe gut/schlecht bedeutet – immer mit Symbol und Text
    '--coe-status-good': '#0aac23',
    '--coe-status-warning': '#fbb315',
    '--coe-status-serious': '#f18a59',
    '--coe-status-critical': '#dc3c3d',
    '--coe-status-neutral': dark ? '#6b6b6b' : '#a3a3a3',
    '--coe-grid': dark ? '#303030' : '#e7e7e7',
    '--coe-axis': dark ? '#474747' : '#c2c2c2',
    '--coe-delta-up': dark ? '#53c556' : '#006c10',
    '--coe-card-shadow': dark ? 'rgba(0, 0, 0, 0.45)' : 'rgba(15, 23, 42, 0.08)',
    '--coe-hero-shadow': withAlpha(accent, dark ? 0.25 : 0.28),
    // Kopfbereich der Übersicht: immer dunkel genug für weiße Schrift (auch in dunklen Abschnitten)
    '--coe-hero-start': darken(accent, dark ? 0.62 : 0.35),
    '--coe-hero-mid': darken(accent, dark ? 0.38 : 0),
    '--coe-hero-end': dark ? '#2c2170' : '#5b4bd6'
  };
  return vars as unknown as React.CSSProperties;
}
