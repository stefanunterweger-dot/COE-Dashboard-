import { IColumnConfig, IRecord } from '../models/ITableConfig';
import { translateValue } from './valueLabels';

export const FORMATTED = '@OData.Community.Display.V1.FormattedValue';

const dateFmt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric' });
const dateTimeFmt = new Intl.DateTimeFormat('de-AT', {
  day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
});
const timeFmt = new Intl.DateTimeFormat('de-AT', { hour: '2-digit', minute: '2-digit' });
const numberFmt = new Intl.NumberFormat('de-AT');

const AVATAR_COLORS = [
  '#0f6cbd', '#8764b8', '#038387', '#ca5010', '#4f6bed',
  '#c239b3', '#498205', '#986f0b', '#e3008c', '#005b70'
];

/**
 * Anzeigewert eines Feldes (bevorzugt den formatierten Wert von Dataverse).
 * Englische CoE-Werte werden für die Anzeige übersetzt (z. B. „Suspended“ → „Gesperrt“).
 */
export function displayValue(record: IRecord, field: string): string {
  const raw = record[field];
  if (typeof raw === 'boolean') {
    return raw ? 'Ja' : 'Nein';
  }
  const formatted = record[field + FORMATTED];
  if (formatted !== undefined && formatted !== null && formatted !== '') {
    return translateValue(field, String(formatted));
  }
  if (raw === undefined || raw === null) {
    return '';
  }
  return translateValue(field, String(raw));
}

/** Originalwert eines Feldes ohne Übersetzung (für Logik und Abfragen). */
export function rawValue(record: IRecord, field: string): string {
  const raw = record[field];
  return raw === undefined || raw === null ? '' : String(raw);
}

export function toDate(value: unknown): Date | undefined {
  if (value === undefined || value === null || value === '') {
    return undefined;
  }
  const d = new Date(String(value));
  return isNaN(d.getTime()) ? undefined : d;
}

export function formatDate(d: Date, withTime?: boolean): string {
  return withTime ? dateTimeFmt.format(d) : dateFmt.format(d);
}

export function formatTime(d: Date): string {
  return timeFmt.format(d);
}

export function formatNumber(n: number): string {
  return numberFmt.format(n);
}

/** Anzahl; mit atLeast = true als „mindestens“ (z. B. „50.000+“). */
export function formatCount(n: number | undefined, atLeast?: boolean): string {
  if (n === undefined) {
    return '–';
  }
  return formatNumber(n) + (atLeast ? '+' : '');
}

/** Nur sichere Links (https) zulassen – verhindert z. B. javascript:-Links aus Link-Vorlagen. */
export function isSafeLink(url: string | undefined): boolean {
  return !!url && /^https:\/\//i.test(url.trim());
}

export function relativeTime(d: Date): string {
  const diffMs = Date.now() - d.getTime();
  const day = 24 * 60 * 60 * 1000;
  const days = Math.floor(diffMs / day);
  if (days < 0) {
    return 'in der Zukunft';
  }
  if (days === 0) {
    return 'heute';
  }
  if (days === 1) {
    return 'gestern';
  }
  if (days < 30) {
    return `vor ${days} Tagen`;
  }
  const months = Math.floor(days / 30);
  if (months < 12) {
    return months === 1 ? 'vor 1 Monat' : `vor ${months} Monaten`;
  }
  const years = Math.floor(days / 365);
  return years <= 1 ? 'vor 1 Jahr' : `vor ${years} Jahren`;
}

export function initials(name: string): string {
  const clean = name.replace(/[(<].*$/, '').trim();
  if (!clean) {
    return '?';
  }
  const parts = clean.split(/[\s.,_-]+/).filter(p => p.length > 0);
  if (parts.length === 1) {
    return parts[0].substring(0, 2).toUpperCase();
  }
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
}

export function hashColor(text: string): string {
  let h = 0;
  for (let i = 0; i < text.length; i++) {
    h = (h * 31 + text.charCodeAt(i)) | 0;
  }
  return AVATAR_COLORS[Math.abs(h) % AVATAR_COLORS.length];
}

/** Hex-Farbe mit Transparenz als rgba(). */
export function withAlpha(color: string, alpha: number): string {
  let hex = color.trim();
  if (hex.charAt(0) !== '#') {
    return color;
  }
  hex = hex.substring(1);
  if (hex.length === 3) {
    hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
  }
  if (hex.length !== 6) {
    return color;
  }
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

/** Vergleichswert für die Sortierung. */
export function sortValue(record: IRecord, column: IColumnConfig): string | number {
  if (column.type === 'date' || column.type === 'datetime') {
    const d = toDate(record[column.field]);
    return d ? d.getTime() : -Infinity;
  }
  if (column.type === 'number') {
    const n = Number(record[column.field]);
    return isNaN(n) ? -Infinity : n;
  }
  return displayValue(record, column.field).toLowerCase();
}

/** Ersetzt {feld}-Platzhalter. Gibt undefined zurück, wenn ein Wert fehlt. */
export function fillTemplate(template: string, record: IRecord): string | undefined {
  let missing = false;
  const result = template.replace(/\{([A-Za-z0-9_]+)\}/g, (_m: string, field: string) => {
    const v = record[field];
    if (v === undefined || v === null || v === '') {
      missing = true;
      return '';
    }
    return encodeURIComponent(String(v));
  });
  return missing ? undefined : result;
}

export function templateFields(template: string): string[] {
  const out: string[] = [];
  const re = /\{([A-Za-z0-9_]+)\}/g;
  let m: RegExpExecArray | null = re.exec(template);
  while (m) {
    out.push(m[1]);
    m = re.exec(template);
  }
  return out;
}

export function combineFilters(a?: string, b?: string): string | undefined {
  const parts = [a, b].filter(p => !!p && p.trim().length > 0) as string[];
  if (parts.length === 0) {
    return undefined;
  }
  if (parts.length === 1) {
    return parts[0];
  }
  return parts.map(p => `(${p})`).join(' and ');
}

/** CSV mit Semikolon (Excel, deutschsprachig) und BOM. */
export function toCsv(records: IRecord[], columns: IColumnConfig[]): string {
  const esc = (raw: string): string => {
    // Werte, die Excel als Formel ausführen würde (=, +, -, @), mit ' entschärfen
    const s = /^[=+\-@\t\r]/.test(raw) && isNaN(Number(raw.replace(',', '.'))) ? `'${raw}` : raw;
    return /[";\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines: string[] = [columns.map(c => esc(c.label)).join(';')];
  records.forEach(r => {
    lines.push(columns.map(c => {
      if (c.type === 'date' || c.type === 'datetime') {
        const d = toDate(r[c.field]);
        return d ? formatDate(d, c.type === 'datetime') : '';
      }
      return esc(displayValue(r, c.field));
    }).join(';'));
  });
  return '﻿' + lines.join('\r\n');
}

export function downloadText(fileName: string, content: string, mime: string): void {
  const blob = new Blob([content], { type: mime });
  const nav = window.navigator as unknown as { msSaveOrOpenBlob?: (b: Blob, n: string) => void };
  if (nav.msSaveOrOpenBlob) {
    nav.msSaveOrOpenBlob(blob, fileName);
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function isValidEnvironmentUrl(url: string | undefined): boolean {
  if (!url) {
    return false;
  }
  return /^https:\/\/[a-z0-9-]+(\.[a-z0-9-]+)*\.(dynamics\.com|microsoftdynamics\.de|microsoftdynamics\.us|appsplatform\.us|dynamics\.cn)\/?$/i.test(url.trim());
}

export function hostOf(url: string): string {
  return url.replace(/^https?:\/\//i, '').replace(/\/.*$/, '');
}

/**
 * Einfache Prüfung einer E-Mail-Adresse. Erlaubt nur Zeichen, die in einem mailto-Link nicht maskiert
 * werden müssen (Gäste-UPNs mit „#EXT#“ sind keine Postfächer und fallen damit heraus).
 */
export function isMailAddress(x: string): boolean {
  return /^[A-Za-z0-9._+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(x);
}
