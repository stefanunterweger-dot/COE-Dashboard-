import { IRecord } from './ITableConfig';
import { DEV_ENV_DEADLINES, DEV_FIELDS, RECOMMENDATION } from './devEnvironments';
import { displayValue, isMailAddress } from '../utils/format';
import { fillText, IMailTemplate, IMailTemplateOverrides, resolveTemplate } from './mailTemplates';
import { escHtml, mailKit } from './mailHtml';

/**
 * E-Mail an die Maker eines betroffenen Developer-Environments.
 * Die Mail wird im lokalen Outlook des Admins geöffnet – das Webpart versendet selbst nichts.
 * Es gibt sie als formatierte HTML-Mail (Zwischenablage bzw. .eml-Entwurf) und als reinen Text (mailto).
 */

/**
 * Offizielle Microsoft-Ankündigung: Abschnitt „Deprecation of unmanaged developer environments and the
 * Power Apps Developer Plan“ (Microsoft Learn). Englische Seite, weil die Übersetzung nachhinkt und
 * der Abschnittsanker nur dort sicher gleich bleibt.
 */
export const DEFAULT_INFO_URL = 'https://learn.microsoft.com/en-us/power-platform/important-changes-coming#deprecation-of-unmanaged-developer-environments-and-the-power-apps-developer-plan';
/** Frühere Standardwerte (werden automatisch durch den aktuellen ersetzt). */
export const OLD_INFO_URLS = ['https://learn.microsoft.com/de-de/power-platform/important-changes-coming'];
/** Microsoft-Anleitung zum Sichern (Solutions exportieren) – Sprache folgt dem Browser. */
export const EXPORT_GUIDE_URL = 'https://learn.microsoft.com/power-apps/maker/data-platform/export-solutions';

/** Längere mailto-Links schneiden Browser bzw. Outlook teilweise ab (ab ca. 2.000 Zeichen). */
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
  /** Gepflegte Mail-Texte (nur Abweichungen vom Standard). */
  templates?: IMailTemplateOverrides;
}

export interface IMailDraft {
  to: string[];
  cc: string[];
  subject: string;
  /** Reiner Text (mailto, Text-Teil des Entwurfs). */
  body: string;
  /** Formatierte Fassung (Outlook-taugliches HTML mit Inline-Styles). */
  html: string;
  /** Maker ohne E-Mail-Adresse im CoE (werden nicht angeschrieben). */
  missing: string[];
  /** Antwortfrist (für die Anzeige). */
  replyBy: Date;
  /** Link zum Environment in Power Apps (make.powerapps.com). */
  makerUrl: string;
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

/** Power-Platform-ID des Environments (in der CoE-Tabelle „Environment Name“ bzw. Primärschlüssel). */
export function environmentPlatformId(r: IRecord): string {
  const name = String(r.admin_environmentname || '').trim();
  return /^(default-)?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(name) ? name : String(r.admin_environmentid || '').trim();
}

/** Link zum Environment in Power Apps. */
export function makerPortalUrl(r: IRecord): string {
  const id = environmentPlatformId(r);
  return id ? `https://make.powerapps.com/environments/${encodeURIComponent(id)}/home` : 'https://make.powerapps.com/';
}

function assessment(rec: string): { text: string; tone: 'danger' | 'warning' | 'neutral' } {
  switch (rec) {
    case RECOMMENDATION.inUse: return { text: 'Wird aktiv genutzt – bitte sichern oder auf Managed umstellen', tone: 'danger' };
    case RECOMMENDATION.inactive: return { text: 'Seit über 90 Tagen keine Aktivität', tone: 'warning' };
    case RECOMMENDATION.empty: return { text: 'Enthält keine Apps und Flows', tone: 'neutral' };
    case RECOMMENDATION.disabled: return { text: 'Bereits deaktiviert', tone: 'neutral' };
    default: return { text: rec || '–', tone: 'neutral' };
  }
}

function listOf(r: IRecord, field: string): string[] {
  return String(r[field] || '').split(',').map(x => x.trim()).filter(x => !!x);
}

/** Alle Inhalte der Mail (einmal bestimmt, dann als Text und als HTML gesetzt). */
interface IContent {
  formal: boolean;
  env: string;
  makerUrl: string;
  apps: string;
  flows: string;
  last: string;
  assess: { text: string; tone: 'danger' | 'warning' | 'neutral' };
  replyBy: string;
  block: string;
  del: string;
  suspended: boolean;
  feedback: string;
  infoUrl: string;
  signature: string;
  /** Gepflegte Texte (Vorlage „devenv“, Platzhalter bereits ersetzt) */
  tpl: IMailTemplate;
}

/** Feste Texte (nicht über die Vorlage pflegbar). */
function fixedTexts(c: IContent): { yourEnv: string; replyTitle: string; notResponsible: string; backup: string; open: string; formButton: string } {
  const S = c.formal;
  return {
    yourEnv: S ? 'Ihr Environment' : 'Dein Environment',
    replyTitle: S ? `Ihre Rückmeldung bis ${c.replyBy}` : `Deine Rückmeldung bis ${c.replyBy}`,
    notResponsible: S ? 'Nicht zuständig? Bitte nennen Sie uns die richtige Person.' : 'Nicht zuständig? Bitte nenne uns die richtige Person.',
    backup: S ? 'Tipp: Sichern Sie wichtige Arbeit vorab, indem Sie Ihre Solutions exportieren:' : 'Tipp: Sichere wichtige Arbeit vorab, indem du deine Solutions exportierst:',
    open: 'Environment in Power Apps öffnen',
    formButton: 'Rückmeldung im Formular geben'
  };
}

/** Satz zur Änderung durch Microsoft (Platzhalter {Änderung}). */
function changeSentence(suspended: boolean, block: string, del: string): string {
  return suspended
    ? `Microsoft hat nicht verwaltete (unmanaged) Developer-Environments am ${block} gesperrt und löscht sie ab ${del} – samt Apps, Flows und Daten.`
    : `Microsoft sperrt ab ${block} alle nicht verwalteten (unmanaged) Developer-Environments und löscht sie ab ${del} – unabhängig von der Nutzung, samt Apps, Flows und Daten.`;
}

function plainText(c: IContent): string {
  const t = fixedTexts(c);
  const lines = [
    c.tpl.greeting,
    '',
    c.tpl.intro,
    '',
    `${t.yourEnv.toUpperCase()}`,
    `• Name: ${c.env}`,
    `• In Power Apps öffnen: ${c.makerUrl}`,
    `• Apps: ${c.apps} · Flows: ${c.flows}`,
    `• Letzte Aktivität: ${c.last}`,
    `• Einschätzung: ${c.assess.text}`,
    '',
    'TERMINE',
    `• Rückmeldung bis ${c.replyBy}`,
    `• Sperre ab ${c.block}`,
    `• Löschung ab ${c.del}`,
    '',
    t.replyTitle.toUpperCase(),
    c.tpl.question
  ].concat(c.tpl.options.map((o, i) => `${i + 1} – ${o}`)).concat(
    c.feedback ? [`Formular: ${c.feedback}`] : [],
    [
      t.notResponsible,
      '',
      c.tpl.note,
      '',
      `${t.backup} ${EXPORT_GUIDE_URL}`,
      '',
      `Offizielle Ankündigung von Microsoft: ${c.infoUrl}`,
      '',
      c.tpl.closing,
      c.signature
    ]
  );
  return lines.join('\r\n');
}

/** Formatierte, „Outlook-feste“ Mail (Bausteine siehe mailHtml.ts). */
function htmlMail(c: IContent, accent: string): string {
  const t = fixedTexts(c);
  const k = mailKit(accent);
  const toneColor = { danger: '#c50f1f', warning: '#8a4b00', neutral: '#616161' }[c.assess.tone];
  const intro = escHtml(c.tpl.intro).replace(escHtml(c.block), `<b>${escHtml(c.block)}</b>`).replace(escHtml(c.del), `<b>${escHtml(c.del)}</b>`);
  return k.frame([
    k.header('POWER PLATFORM &middot; HANDLUNGSBEDARF', c.tpl.headline),
    k.space(24),
    k.block(k.tx(escHtml(c.tpl.greeting))),
    k.space(12),
    k.block(k.tx(intro)),
    k.space(20),
    // Fristen
    k.block(k.table(`<tr>${k.dateCell('Rückmeldung bis', c.replyBy, k.ink, '#f5f5f5')}${k.gap}${k.dateCell('Sperre ab', c.block, '#8a4b00', '#fff8f0')}${k.gap}${k.dateCell('Löschung ab', c.del, '#c50f1f', '#fdf3f4')}</tr>`)),
    k.space(26),
    // Environment
    k.heading(t.yourEnv),
    k.space(10),
    k.block(k.table([
      k.fact('Name', k.link(c.makerUrl, c.env, k.ink, true)),
      k.fact('Apps / Flows', k.tx(`${escHtml(c.apps)} Apps &middot; ${escHtml(c.flows)} Flows`)),
      k.fact('Letzte Aktivität', k.tx(escHtml(c.last))),
      k.fact('Einschätzung', k.tx(escHtml(c.assess.text), { color: toneColor, bold: true }), true)
    ].join(''), `border:1px solid ${k.BORDER};`)),
    k.space(28),
    // Rückmeldung
    k.heading(t.replyTitle),
    k.space(8),
    k.block(k.tx(escHtml(c.tpl.question))),
    k.space(12),
    k.block(k.table(c.tpl.options.map((o, i) => k.option(String(i + 1), o)).join(''))),
    c.feedback ? k.space(4) + k.block(k.button(c.feedback, t.formButton)) + k.space(14) : '',
    k.block(k.tx(escHtml(t.notResponsible), { color: k.MUTED })),
    k.space(16),
    c.tpl.note ? k.note(c.tpl.note) : '',
    k.space(18),
    // Schaltfläche zum Environment – unter dem Hinweis „Ohne Rückmeldung …“
    k.block(k.button(c.makerUrl, t.open)),
    k.space(22),
    k.block(`${k.tx(escHtml(t.backup))} ${k.link(EXPORT_GUIDE_URL, 'Anleitung von Microsoft', k.ink)}`),
    k.space(8),
    k.block(`${k.tx('Offizielle Ankündigung von Microsoft:')} ${k.link(c.infoUrl, 'Deprecation of unmanaged developer environments (Microsoft Learn)', k.ink)}`),
    k.space(24),
    k.block(`${k.tx(escHtml(c.tpl.closing))}<br>${k.tx(escHtml(c.signature), { bold: true })}`),
    k.space(28)
  ]);
}

/** Mail für ein Environment aus dem Reiter „Betroffen ab 15.12.“. accent = Akzentfarbe des Webparts. */
export function buildDevEnvMail(r: IRecord, s: INotifySettings, today: Date = new Date(), accent: string = '#0f6cbd'): IMailDraft {
  const env = displayValue(r, 'admin_displayname') || String(r.admin_environmentid || '');
  const lastRaw = r[DEV_FIELDS.lastActivity];
  const lastDate = lastRaw ? new Date(String(lastRaw)) : undefined;
  const replyBy = replyDeadline(s.replyDays, today);
  const num = (v: unknown): string => (v === null || v === undefined ? '?' : String(v));
  const makerUrl = makerPortalUrl(r);
  const suspended = today >= DEV_ENV_DEADLINES.block;
  const block = fmt(DEV_ENV_DEADLINES.block);
  const del = fmt(DEV_ENV_DEADLINES.delete);
  const feedback = (s.feedbackUrl || '').trim()
    .replace(/\{environment\}/g, encodeURIComponent(env))
    .replace(/\{id\}/g, encodeURIComponent(environmentPlatformId(r)));
  const raw = resolveTemplate('devenv', s.formal, !!feedback, s.templates);
  const vars: { [name: string]: string } = {
    Name: env,
    Frist: fmt(replyBy),
    Sperre: block,
    Löschung: del,
    Änderung: changeSentence(suspended, block, del),
    Zustand: suspended ? 'gelöscht' : 'gesperrt',
    ZustandMitDatum: suspended ? 'gelöscht' : `am ${block} gesperrt`
  };
  const tpl: IMailTemplate = {
    subject: fillText(raw.subject, vars),
    headline: fillText(raw.headline, vars),
    greeting: fillText(raw.greeting, vars),
    intro: fillText(raw.intro, vars),
    question: fillText(raw.question, vars),
    options: raw.options.map(o => fillText(o, vars)),
    note: fillText(raw.note, vars),
    closing: fillText(raw.closing, vars)
  };
  const c: IContent = {
    formal: s.formal,
    env,
    makerUrl,
    apps: num(r[DEV_FIELDS.apps]),
    flows: num(r[DEV_FIELDS.flows]),
    last: lastDate && !isNaN(lastDate.getTime()) ? fmt(lastDate) : 'keine erfasst',
    assess: assessment(String(r[DEV_FIELDS.recommendation] || '')),
    replyBy: fmt(replyBy),
    block,
    del,
    suspended,
    feedback,
    infoUrl: (s.infoUrl || DEFAULT_INFO_URL).trim(),
    signature: s.signature || (s.formal ? 'Ihr Power-Platform-Team' : 'Dein Power-Platform-Team'),
    tpl
  };
  return {
    to: listOf(r, DEV_FIELDS.makerEmails),
    cc: parseAddresses(s.cc),
    subject: tpl.subject,
    body: plainText(c),
    html: htmlMail(c, accent),
    missing: listOf(r, DEV_FIELDS.makersNoMail),
    replyBy,
    makerUrl
  };
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

/** UTF-8 → Base64 in Zeilen zu 76 Zeichen (für den .eml-Entwurf). */
function utf8Base64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, i + 0x8000) as number[]);
  }
  return btoa(bin);
}

function base64Lines(text: string): string {
  return (utf8Base64(text).match(/.{1,76}/g) || []).join('\r\n');
}

/**
 * Outlook-Entwurf als .eml: „X-Unsent: 1“ öffnet die Datei in Outlook (klassisch) als neue, noch
 * nicht gesendete Mail – mit Empfängern, Betreff und formatiertem Text.
 */
export function emlDraft(d: { to: string[]; cc: string[]; subject: string; body: string; html: string }): string {
  const boundary = `----=_CoeCockpit_${Date.now().toString(36)}`;
  const headers = [
    // Lange Empfängerlisten auf mehrere Zeilen falten (Kopfzeilen höchstens 998 Zeichen)
    `To: ${d.to.join(',\r\n ')}`,
    d.cc.length ? `Cc: ${d.cc.join(',\r\n ')}` : '',
    `Subject: =?UTF-8?B?${utf8Base64(d.subject)}?=`,
    'X-Unsent: 1',
    'MIME-Version: 1.0',
    `Content-Type: multipart/alternative; boundary="${boundary}"`
  ].filter(x => !!x);
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:16px;background:#ffffff;">${d.html}</body></html>`;
  return headers.join('\r\n') + '\r\n\r\n' + [
    `--${boundary}`,
    'Content-Type: text/plain; charset="utf-8"',
    'Content-Transfer-Encoding: base64',
    '',
    base64Lines(d.body),
    `--${boundary}`,
    'Content-Type: text/html; charset="utf-8"',
    'Content-Transfer-Encoding: base64',
    '',
    base64Lines(html),
    `--${boundary}--`,
    ''
  ].join('\r\n');
}
