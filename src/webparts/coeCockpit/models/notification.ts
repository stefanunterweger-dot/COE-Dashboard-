import { IRecord } from './ITableConfig';
import { DEV_ENV_DEADLINES, DEV_FIELDS, RECOMMENDATION } from './devEnvironments';
import { displayValue, isMailAddress } from '../utils/format';

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

/** Formatierte Mail – Tabellen und Inline-Styles, damit Outlook (Word-Engine) sie sauber darstellt. */
function htmlMail(c: IContent, accent: string): string {
  const t = texts(c);
  const on = onColor(accent);
  /** Farbe für Links und Akzente auf weißem Grund (helle Akzentfarben → Standardblau). */
  const ink = on === '#ffffff' ? accent : '#0f6cbd';
  const font = "font-family:'Segoe UI',Calibri,Arial,sans-serif;";
  const muted = '#605e5c';
  const border = '#e1dfdd';
  const toneColor = { danger: '#c50f1f', warning: '#b35c00', neutral: '#605e5c' }[c.assess.tone];
  const toneBg = { danger: '#fdf3f4', warning: '#fff8f0', neutral: '#f3f2f1' }[c.assess.tone];
  const p = (html: string, extra?: string): string => `<p style="margin:0 0 12px;${font}font-size:14px;line-height:1.5;color:#1f1f1f;${extra || ''}">${html}</p>`;
  const h = (text: string): string => `<p style="margin:22px 0 8px;${font}font-size:15px;font-weight:600;color:#1f1f1f;">${esc(text)}</p>`;
  const button = (href: string, label: string): string =>
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:12px 0 4px;border-collapse:collapse;"><tr>` +
    `<td bgcolor="${accent}" style="background:${accent};padding:10px 18px;border-radius:6px;">` +
    `<a href="${esc(href)}" style="${font}font-size:14px;font-weight:600;color:${on};text-decoration:none;">${esc(label)} &#8594;</a></td></tr></table>`;
  const dateCell = (label: string, value: string, color: string, bg: string): string =>
    `<td width="33%" valign="top" bgcolor="${bg}" style="background:${bg};border-left:4px solid ${color};padding:10px 12px;">` +
    `<div style="${font}font-size:12px;color:${muted};">${esc(label)}</div>` +
    `<div style="${font}font-size:18px;font-weight:700;color:${color};">${esc(value)}</div></td>`;
  const row = (label: string, valueHtml: string, last?: boolean): string =>
    `<tr><td width="34%" style="${font}font-size:13px;color:${muted};padding:8px 12px;${last ? '' : `border-bottom:1px solid ${border};`}">${esc(label)}</td>` +
    `<td style="${font}font-size:14px;color:#1f1f1f;padding:8px 12px;${last ? '' : `border-bottom:1px solid ${border};`}">${valueHtml}</td></tr>`;
  const option = (n: string, text: string): string =>
    `<tr><td width="34" valign="top" style="padding:6px 0;">` +
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td width="26" height="26" align="center" bgcolor="${accent}" style="background:${accent};border-radius:13px;${font}font-size:13px;font-weight:700;color:${on};">${n}</td></tr></table></td>` +
    `<td valign="top" style="padding:8px 0 6px 8px;${font}font-size:14px;line-height:1.45;color:#1f1f1f;">${esc(text)}</td></tr>`;

  return [
    `<div style="${font}font-size:14px;color:#1f1f1f;">`,
    `<table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" style="width:640px;border-collapse:collapse;">`,
    // Kopf
    `<tr><td bgcolor="${accent}" style="background:${accent};padding:18px 24px;">`,
    `<div style="${font}font-size:12px;letter-spacing:1px;text-transform:uppercase;color:${on};">Power Platform · Handlungsbedarf</div>`,
    `<div style="${font}font-size:21px;font-weight:700;color:${on};margin-top:4px;">${esc(c.suspended ? `Developer-Environment „${c.env}“ wird gelöscht` : `Developer-Environment „${c.env}“ wird am ${c.block} gesperrt`)}</div>`,
    `</td></tr>`,
    // Inhalt
    `<tr><td style="padding:22px 24px 24px;border:1px solid ${border};border-top:0;">`,
    p(esc(t.greeting)),
    p(`${esc(t.intro)} ${esc(t.change).replace(esc(c.block), `<b>${esc(c.block)}</b>`).replace(esc(c.del), `<b>${esc(c.del)}</b>`)}`),
    // Fristen
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;border-spacing:6px 0;margin:6px -6px 4px;"><tr>`,
    dateCell('Rückmeldung bis', c.replyBy, ink, '#f5f5f5'),
    dateCell('Sperre ab', c.block, '#b35c00', '#fff8f0'),
    dateCell('Löschung ab', c.del, '#c50f1f', '#fdf3f4'),
    `</tr></table>`,
    // Environment
    h(t.yourEnv),
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:1px solid ${border};">`,
    row('Name', `<a href="${esc(c.makerUrl)}" style="color:${ink};font-weight:600;text-decoration:none;">${esc(c.env)}</a>`),
    row('Apps / Flows', `${esc(c.apps)} Apps · ${esc(c.flows)} Flows`),
    row('Letzte Aktivität', esc(c.last)),
    row('Einschätzung', `<span style="background:${toneBg};color:${toneColor};font-weight:600;padding:2px 8px;">${esc(c.assess.text)}</span>`, true),
    `</table>`,
    button(c.makerUrl, t.open),
    // Rückmeldung
    h(t.replyTitle),
    p(esc(t.replyAsk), 'margin-bottom:4px;'),
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">`,
    t.options.map(o => option(o[0], o[1])).join(''),
    `</table>`,
    c.feedback ? button(c.feedback, t.formButton) : '',
    p(esc(t.notResponsible), `color:${muted};margin-top:8px;`),
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:8px 0 16px;"><tr>`,
    `<td bgcolor="#f3f2f1" style="background:#f3f2f1;padding:10px 14px;${font}font-size:13px;color:#323130;">${esc(t.noReply)}</td>`,
    `</tr></table>`,
    p(`${esc(t.backup)} <a href="${EXPORT_GUIDE_URL}" style="color:${ink};">Anleitung von Microsoft</a>`),
    p(`Offizielle Ankündigung von Microsoft: <a href="${esc(c.infoUrl)}" style="color:${ink};">Deprecation of unmanaged developer environments (Microsoft Learn)</a>`),
    p(`${esc(t.closing)}<br>${esc(c.signature)}`, 'margin:18px 0 0;'),
    `</td></tr></table></div>`
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
