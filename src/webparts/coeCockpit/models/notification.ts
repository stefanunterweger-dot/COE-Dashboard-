import { IRecord } from './ITableConfig';
import { DEV_ENV_DEADLINES, DEV_FIELDS, RECOMMENDATION } from './devEnvironments';
import { displayValue, isMailAddress } from '../utils/format';
import { IMailTemplateOverrides } from './mailTemplates';

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

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Kontrastfarbe für Text auf der Akzentfarbe. */
function onColor(hex: string): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) {
    return '#ffffff';
  }
  const n = parseInt(m[1], 16);
  const lin = (c: number): number => { const v = c / 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
  const L = 0.2126 * lin((n >> 16) & 255) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
  // Weiß nur, wenn der Kontrast mindestens 4.5:1 erreicht
  return L > 0.183 ? '#1b1b1b' : '#ffffff';
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
}

function texts(c: IContent): {
  greeting: string; intro: string; change: string; yourEnv: string; replyTitle: string; replyAsk: string;
  options: [string, string][]; notResponsible: string; noReply: string; backup: string; open: string; formButton: string; closing: string;
} {
  const S = c.formal;
  return {
    greeting: S ? 'Guten Tag,' : 'Hallo,',
    intro: S
      ? `Sie sind als Maker im Power-Platform-Developer-Environment „${c.env}“ eingetragen.`
      : `du bist als Maker im Power-Platform-Developer-Environment „${c.env}“ eingetragen.`,
    change: c.suspended
      ? `Microsoft hat nicht verwaltete (unmanaged) Developer-Environments am ${c.block} gesperrt und löscht sie ab ${c.del} – samt Apps, Flows und Daten.`
      : `Microsoft sperrt ab ${c.block} alle nicht verwalteten (unmanaged) Developer-Environments und löscht sie ab ${c.del} – unabhängig von der Nutzung, samt Apps, Flows und Daten.`,
    yourEnv: S ? 'Ihr Environment' : 'Dein Environment',
    replyTitle: S ? `Ihre Rückmeldung bis ${c.replyBy}` : `Deine Rückmeldung bis ${c.replyBy}`,
    replyAsk: c.feedback
      ? (S ? 'Bitte nutzen Sie das Formular oder antworten Sie auf diese E-Mail mit der passenden Zahl:' : 'Bitte nutze das Formular oder antworte auf diese E-Mail mit der passenden Zahl:')
      : (S ? 'Bitte antworten Sie einfach auf diese E-Mail mit der passenden Zahl:' : 'Bitte antworte einfach auf diese E-Mail mit der passenden Zahl:'),
    options: [
      ['1', 'Wird weiter benötigt – das Environment wird auf Managed umgestellt. Wer dort Apps oder Flows nutzt, braucht dann eine Premium-Lizenz.'],
      ['2', 'Kann gelöscht werden.'],
      ['3', S ? 'Inhalte kommen in ein anderes Environment – bitte nennen Sie das Ziel.' : 'Inhalte kommen in ein anderes Environment – bitte nenne das Ziel.']
    ],
    notResponsible: S ? 'Nicht zuständig? Bitte nennen Sie uns die richtige Person.' : 'Nicht zuständig? Bitte nenne uns die richtige Person.',
    noReply: 'Ohne Rückmeldung unternehmen wir nichts – Microsoft sperrt und löscht das Environment dann automatisch.',
    backup: S ? 'Tipp: Sichern Sie wichtige Arbeit vorab, indem Sie Ihre Solutions exportieren:' : 'Tipp: Sichere wichtige Arbeit vorab, indem du deine Solutions exportierst:',
    open: 'Environment in Power Apps öffnen',
    formButton: 'Rückmeldung im Formular geben',
    closing: S ? 'Freundliche Grüße' : 'Viele Grüße'
  };
}

function plainText(c: IContent): string {
  const t = texts(c);
  const lines = [
    t.greeting,
    '',
    t.intro,
    t.change,
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
    t.replyAsk
  ].concat(t.options.map(o => `${o[0]} – ${o[1]}`)).concat(
    c.feedback ? [`Formular: ${c.feedback}`] : [],
    [
      t.notResponsible,
      '',
      t.noReply,
      '',
      `${texts(c).backup} ${EXPORT_GUIDE_URL}`,
    '',
    `Offizielle Ankündigung von Microsoft: ${c.infoUrl}`,
      '',
      t.closing,
      c.signature
    ]
  );
  return lines.join('\r\n');
}

/**
 * Formatierte Mail, „Outlook-fest“: Outlook (klassisch) stellt HTML mit der Word-Engine dar und ignoriert
 * dabei vieles (div-Formate, Abstände von Absätzen, runde Ecken, Linkfarben). Deshalb nur Tabellen,
 * Zellen-Innenabstände, eigene Abstandszeilen, Farben direkt am Text (span + font) und keine runden Elemente.
 */
function htmlMail(c: IContent, accent: string): string {
  const t = texts(c);
  const on = onColor(accent);
  /** Farbe für Links und Akzente auf weißem Grund (helle Akzentfarben → Standardblau). */
  const ink = on === '#ffffff' ? accent : '#0f6cbd';
  const FONT = "'Segoe UI', Calibri, Arial, sans-serif";
  const TEXT = '#242424';
  const MUTED = '#616161';
  const BORDER = '#e0e0e0';
  const toneColor = { danger: '#c50f1f', warning: '#8a4b00', neutral: '#616161' }[c.assess.tone];

  /** Text mit fester Schrift, Größe und Farbe – span (CSS) und font (für Word) zugleich. */
  const tx = (html: string, o: { size?: number; color?: string; bold?: boolean; spacing?: number } = {}): string => {
    const size = o.size || 14;
    const color = o.color || TEXT;
    const inner = o.bold ? `<b>${html}</b>` : html;
    return `<span style="font-family:${FONT};font-size:${size}px;line-height:${Math.round(size * 1.45)}px;color:${color};` +
      `${o.bold ? 'font-weight:bold;' : ''}${o.spacing ? `letter-spacing:${o.spacing}px;` : ''}"><font face="Segoe UI, Calibri, Arial, sans-serif" color="${color}">${inner}</font></span>`;
  };
  const link = (href: string, label: string, color: string, bold?: boolean): string =>
    `<a href="${esc(href)}" style="color:${color};text-decoration:underline;">${tx(esc(label), { color, bold })}</a>`;
  const space = (h: number): string =>
    `<tr><td height="${h}" style="height:${h}px;font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>`;
  /** Inhaltszeile mit seitlichem Abstand. */
  const block = (html: string): string => `<tr><td style="padding:0 28px;">${html}</td></tr>`;
  const table = (rows: string, extra?: string): string =>
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;${extra || ''}">${rows}</table>`;
  const button = (href: string, label: string): string =>
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr>` +
    `<td bgcolor="${accent}" style="background-color:${accent};padding:11px 22px;">` +
    `<a href="${esc(href)}" style="color:${on};text-decoration:none;">${tx(`${esc(label)}&nbsp;&#8594;`, { color: on, bold: true })}</a>` +
    `</td></tr></table>`;
  const dateCell = (label: string, value: string, color: string, bg: string): string =>
    `<td width="32%" valign="top" bgcolor="${bg}" style="background-color:${bg};border-left:4px solid ${color};padding:10px 14px;">` +
    `${tx(esc(label), { size: 12, color: MUTED })}<br>${tx(esc(value), { size: 18, color, bold: true })}</td>`;
  const gap = '<td width="2%" style="font-size:1px;line-height:1px;">&nbsp;</td>';
  const fact = (label: string, valueHtml: string, last?: boolean): string => {
    const line = last ? '' : `border-bottom:1px solid ${BORDER};`;
    return `<tr><td width="34%" valign="top" style="padding:9px 14px;${line}">${tx(esc(label), { size: 13, color: MUTED })}</td>` +
      `<td valign="top" style="padding:9px 14px;${line}">${valueHtml}</td></tr>`;
  };
  const option = (n: string, text: string): string =>
    `<tr><td width="30" valign="top" style="padding:0 0 10px 0;">${tx(`${n}.`, { size: 16, color: ink, bold: true })}</td>` +
    `<td valign="top" style="padding:1px 0 10px 0;">${tx(esc(text))}</td></tr>`;
  const heading = (text: string): string => block(tx(esc(text), { size: 16, bold: true }));
  const intro = `${esc(t.intro)} ${esc(t.change)}`
    .replace(esc(c.block), `<b>${esc(c.block)}</b>`).replace(esc(c.del), `<b>${esc(c.del)}</b>`);

  return [
    `<table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" style="width:640px;border-collapse:collapse;border:1px solid ${BORDER};background-color:#ffffff;">`,
    // Kopfbalken
    `<tr><td bgcolor="${accent}" style="background-color:${accent};padding:20px 28px 22px 28px;">`,
    tx('POWER PLATFORM &middot; HANDLUNGSBEDARF', { size: 12, color: on, spacing: 1 }),
    '<br>',
    tx(esc(c.suspended ? `Developer-Environment „${c.env}“ wird gelöscht` : `Developer-Environment „${c.env}“ wird am ${c.block} gesperrt`), { size: 20, color: on, bold: true }),
    `</td></tr>`,
    space(24),
    block(tx(esc(t.greeting))),
    space(12),
    block(tx(intro)),
    space(20),
    // Fristen
    block(table(`<tr>${dateCell('Rückmeldung bis', c.replyBy, ink, '#f5f5f5')}${gap}${dateCell('Sperre ab', c.block, '#8a4b00', '#fff8f0')}${gap}${dateCell('Löschung ab', c.del, '#c50f1f', '#fdf3f4')}</tr>`)),
    space(26),
    // Environment
    heading(t.yourEnv),
    space(10),
    block(table([
      fact('Name', link(c.makerUrl, c.env, ink, true)),
      fact('Apps / Flows', tx(`${esc(c.apps)} Apps &middot; ${esc(c.flows)} Flows`)),
      fact('Letzte Aktivität', tx(esc(c.last))),
      fact('Einschätzung', tx(esc(c.assess.text), { color: toneColor, bold: true }), true)
    ].join(''), `border:1px solid ${BORDER};`)),
    space(28),
    // Rückmeldung
    heading(t.replyTitle),
    space(8),
    block(tx(esc(t.replyAsk))),
    space(12),
    block(table(t.options.map(o => option(o[0], o[1])).join(''))),
    c.feedback ? space(4) + block(button(c.feedback, t.formButton)) + space(14) : '',
    block(tx(esc(t.notResponsible), { color: MUTED })),
    space(16),
    block(table(`<tr><td bgcolor="#f5f5f5" style="background-color:#f5f5f5;border-left:4px solid #8a8886;padding:12px 16px;">${tx(esc(t.noReply), { size: 13, color: '#424242' })}</td></tr>`)),
    space(18),
    // Schaltfläche zum Environment – unter dem Hinweis „Ohne Rückmeldung …“
    block(button(c.makerUrl, t.open)),
    space(22),
    block(`${tx(esc(t.backup))} ${link(EXPORT_GUIDE_URL, 'Anleitung von Microsoft', ink)}`),
    space(8),
    block(`${tx('Offizielle Ankündigung von Microsoft:')} ${link(c.infoUrl, 'Deprecation of unmanaged developer environments (Microsoft Learn)', ink)}`),
    space(24),
    block(`${tx(esc(t.closing))}<br>${tx(esc(c.signature), { bold: true })}`),
    space(28),
    `</table>`
  ].join('');
}

/** Mail für ein Environment aus dem Reiter „Betroffen ab 15.12.“. accent = Akzentfarbe des Webparts. */
export function buildDevEnvMail(r: IRecord, s: INotifySettings, today: Date = new Date(), accent: string = '#0f6cbd'): IMailDraft {
  const env = displayValue(r, 'admin_displayname') || String(r.admin_environmentid || '');
  const lastRaw = r[DEV_FIELDS.lastActivity];
  const lastDate = lastRaw ? new Date(String(lastRaw)) : undefined;
  const replyBy = replyDeadline(s.replyDays, today);
  const num = (v: unknown): string => (v === null || v === undefined ? '?' : String(v));
  const makerUrl = makerPortalUrl(r);
  const c: IContent = {
    formal: s.formal,
    env,
    makerUrl,
    apps: num(r[DEV_FIELDS.apps]),
    flows: num(r[DEV_FIELDS.flows]),
    last: lastDate && !isNaN(lastDate.getTime()) ? fmt(lastDate) : 'keine erfasst',
    assess: assessment(String(r[DEV_FIELDS.recommendation] || '')),
    replyBy: fmt(replyBy),
    block: fmt(DEV_ENV_DEADLINES.block),
    del: fmt(DEV_ENV_DEADLINES.delete),
    suspended: today >= DEV_ENV_DEADLINES.block,
    feedback: (s.feedbackUrl || '').trim()
      .replace(/\{environment\}/g, encodeURIComponent(env))
      .replace(/\{id\}/g, encodeURIComponent(environmentPlatformId(r))),
    infoUrl: (s.infoUrl || DEFAULT_INFO_URL).trim(),
    signature: s.signature || (s.formal ? 'Ihr Power-Platform-Team' : 'Dein Power-Platform-Team')
  };
  const color = /^#[0-9a-f]{6}$/i.test(accent) ? accent : '#0f6cbd';
  return {
    to: listOf(r, DEV_FIELDS.makerEmails),
    cc: parseAddresses(s.cc),
    subject: `Rückmeldung bis ${c.replyBy}: Developer-Environment „${env}“ wird ${c.suspended ? 'gelöscht' : 'gesperrt'}`,
    body: plainText(c),
    html: htmlMail(c, color),
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
