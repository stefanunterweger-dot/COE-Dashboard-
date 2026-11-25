import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { withAlpha } from '../utils/format';

type Dict = { [key: string]: string | undefined };

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
    '--coe-info-bg': withAlpha(info, dark ? 0.2 : 0.09)
  };
  return vars as unknown as React.CSSProperties;
}
