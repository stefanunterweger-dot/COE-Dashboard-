import { IRecord } from './ITableConfig';
import { DEV_ENV_DEADLINES, DEV_FIELDS, RECOMMENDATION } from './devEnvironments';
import { displayValue, isMailAddress } from '../utils/format';

/**
 * E-Mail an die Maker eines betroffenen Developer-Environments.
 * Die Mail wird im lokalen Outlook des Admins geöffnet (mailto) – das Webpart versendet selbst nichts.
 */

/** Offizielle Microsoft-Seite zu den Änderungen (Developer-Environments ohne Managed Environments). */
export const DEFAULT_INFO_URL = 'https://learn.microsoft.com/de-de/power-platform/important-changes-coming';

/**
 * Längere mailto-Links schneiden Browser bzw. Outlook teilweise ab (ab ca. 2.000 Zeichen).
 * Darüber wird der Text in die Zwischenablage kopiert und in Outlook eingefügt.
 */
export const MAILTO_LIMIT = 1900;

export interface INotifySettings {
  /** Anrede: Sie (true) oder Du (false). */
  formal: boolean;
  /** Frist für die Rückmeldung in Tagen ab heute. */
  replyDays: number;
  /** Link zur offiziellen Microsoft-Information. */
  infoUrl: string;
  /** Optional: Formular für die Rückmeldung (Platzhalter {environment} und {id} werden ersetzt). */
  feedbackUrl: string;
  /** Optional: Adresse(n) in CC, z. B. das Postfach des Power-Platform-Teams. */
  cc: string;
  /** Grußzeile, z. B. „Ihr Power-Platform-Team“. */
  signature: string;
}

export interface IMailDraft {
  to: string[];
  cc: string[];
  subject: string;
  body: string;
  /** Maker ohne E-Mail-Adresse im CoE (werden nicht angeschrieben). */
  missing: string[];
  /** Antwortfrist (für die Anzeige). */
  replyBy: Date;
}

const DAY = 86400000;
const dateFmt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric' });
const fmt = (d: Date): string => dateFmt.format(d);

/** Adressen aus einem Freitext (getrennt durch Komma, Semikolon oder Leerzeichen). */
export function parseAddresses(text: string): string[] {
  return (text || '').split(/[;,\s]+/).map(x => x.trim()).filter(isMailAddress);
}

/** Frist: heute + Tage, aber spätestens 3 Tage vor dem nächsten Stichtag (Sperre bzw. Löschung). */
export function replyDeadline(days: number, today: Date = new Date()): Date {
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const wanted = new Date(start.getTime() + Math.max(1, days) * DAY);
  const next = start < DEV_ENV_DEADLINES.block ? DEV_ENV_DEADLINES.block : DEV_ENV_DEADLINES.delete;
  const latest = new Date(next.getTime() - 3 * DAY);
  return latest > start && wanted > latest ? latest : wanted;
}

function assessment(rec: string): string {
  switch (rec) {
    case RECOMMENDATION.inUse: return 'wird aktiv genutzt – bitte sichern oder auf Managed umstellen';
    case RECOMMENDATION.inactive: return 'seit über 90 Tagen keine Aktivität';
    case RECOMMENDATION.empty: return 'enthält keine Apps und Flows';
    case RECOMMENDATION.disabled: return 'ist bereits deaktiviert';
    default: return rec || '–';
  }
}

function listOf(r: IRecord, field: string): string[] {
  return String(r[field] || '').split(',').map(x => x.trim()).filter(x => !!x);
}

/** Mail für ein Environment aus dem Reiter „Betroffen ab 15.12.“. */
export function buildDevEnvMail(r: IRecord, s: INotifySettings, today: Date = new Date()): IMailDraft {
  const formal = s.formal;
  const env = displayValue(r, 'admin_displayname') || String(r.admin_environmentid || '');
  const to = listOf(r, DEV_FIELDS.makerEmails);
  const missing = listOf(r, DEV_FIELDS.makersNoMail);
  const replyBy = replyDeadline(s.replyDays, today);
  const suspended = today >= DEV_ENV_DEADLINES.block;
  const lastRaw = r[DEV_FIELDS.lastActivity];
  const lastDate = lastRaw ? new Date(String(lastRaw)) : undefined;
  const last = lastDate && !isNaN(lastDate.getTime()) ? fmt(lastDate) : 'keine erfasst';
  const apps = r[DEV_FIELDS.apps] === null || r[DEV_FIELDS.apps] === undefined ? '?' : String(r[DEV_FIELDS.apps]);
  const flows = r[DEV_FIELDS.flows] === null || r[DEV_FIELDS.flows] === undefined ? '?' : String(r[DEV_FIELDS.flows]);
  const feedback = (s.feedbackUrl || '').trim()
    .replace(/\{environment\}/g, encodeURIComponent(env))
    .replace(/\{id\}/g, encodeURIComponent(String(r.admin_environmentid || '')));
  const infoUrl = (s.infoUrl || DEFAULT_INFO_URL).trim();
  const block = fmt(DEV_ENV_DEADLINES.block);
  const del = fmt(DEV_ENV_DEADLINES.delete);

  const subject = `Rückmeldung bis ${fmt(replyBy)}: Developer-Environment „${env}“ wird ${suspended ? 'gelöscht' : 'gesperrt'}`;
  const lines: string[] = formal
    ? [
      'Guten Tag,',
      '',
      `Sie sind als Maker im Power-Platform-Developer-Environment „${env}“ eingetragen.`,
      '',
      suspended
        ? `Microsoft hat nicht verwaltete (unmanaged) Developer-Environments am ${block} gesperrt und löscht sie ab ${del} – samt Apps, Flows und Daten.`
        : `Microsoft sperrt ab ${block} alle nicht verwalteten (unmanaged) Developer-Environments und löscht sie ab ${del} – unabhängig von der Nutzung, samt Apps, Flows und Daten.`,
      '',
      'Ihr Environment:',
      `- Apps: ${apps}, Flows: ${flows}`,
      `- Letzte Aktivität: ${last}`,
      `- Einschätzung: ${assessment(String(r[DEV_FIELDS.recommendation] || ''))}`,
      '',
      feedback
        ? `Bitte geben Sie uns bis ${fmt(replyBy)} Rückmeldung (Formular: ${feedback}) oder antworten Sie auf diese E-Mail mit:`
        : `Bitte antworten Sie bis ${fmt(replyBy)} auf diese E-Mail mit:`,
      '1 = Wird weiter benötigt: Umstellung auf Managed. Wer dort Apps oder Flows nutzt, braucht dann eine Premium-Lizenz.',
      '2 = Kann gelöscht werden.',
      '3 = Inhalte kommen in ein anderes Environment (bitte nennen).',
      'Nicht zuständig? Bitte nennen Sie uns die richtige Person.',
      '',
      'Ohne Rückmeldung unternehmen wir nichts – Microsoft sperrt und löscht das Environment dann automatisch.',
      '',
      `Information von Microsoft: ${infoUrl}`,
      '',
      'Freundliche Grüße',
      s.signature || 'Ihr Power-Platform-Team'
    ]
    : [
      'Hallo,',
      '',
      `du bist als Maker im Power-Platform-Developer-Environment „${env}“ eingetragen.`,
      '',
      suspended
        ? `Microsoft hat nicht verwaltete (unmanaged) Developer-Environments am ${block} gesperrt und löscht sie ab ${del} – samt Apps, Flows und Daten.`
        : `Microsoft sperrt ab ${block} alle nicht verwalteten (unmanaged) Developer-Environments und löscht sie ab ${del} – unabhängig von der Nutzung, samt Apps, Flows und Daten.`,
      '',
      'Dein Environment:',
      `- Apps: ${apps}, Flows: ${flows}`,
      `- Letzte Aktivität: ${last}`,
      `- Einschätzung: ${assessment(String(r[DEV_FIELDS.recommendation] || ''))}`,
      '',
      feedback
        ? `Bitte gib uns bis ${fmt(replyBy)} Rückmeldung (Formular: ${feedback}) oder antworte auf diese E-Mail mit:`
        : `Bitte antworte bis ${fmt(replyBy)} auf diese E-Mail mit:`,
      '1 = Wird weiter benötigt: Umstellung auf Managed. Wer dort Apps oder Flows nutzt, braucht dann eine Premium-Lizenz.',
      '2 = Kann gelöscht werden.',
      '3 = Inhalte kommen in ein anderes Environment (bitte nennen).',
      'Nicht zuständig? Bitte nenne uns die richtige Person.',
      '',
      'Ohne Rückmeldung unternehmen wir nichts – Microsoft sperrt und löscht das Environment dann automatisch.',
      '',
      `Information von Microsoft: ${infoUrl}`,
      '',
      'Viele Grüße',
      s.signature || 'Dein Power-Platform-Team'
    ];
  return { to, cc: parseAddresses(s.cc), subject, body: lines.join('\r\n'), missing, replyBy };
}

/** mailto-Link (Outlook trennt Empfänger standardmäßig mit Semikolon). */
export function mailtoUrl(to: string[], cc: string[], subject: string, body: string): string {
  // Adressen sind geprüft (nur Buchstaben, Ziffern, @ . - _ +) und bleiben im Link lesbar
  const params: string[] = [];
  if (cc.length) {
    params.push(`cc=${cc.join(';')}`);
  }
  params.push(`subject=${encodeURIComponent(subject)}`);
  if (body) {
    params.push(`body=${encodeURIComponent(body.replace(/\r?\n/g, '\r\n'))}`);
  }
  return `mailto:${to.join(';')}?${params.join('&')}`;
}
