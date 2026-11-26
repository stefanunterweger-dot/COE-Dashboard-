import { IRecord, ITableConfig } from './ITableConfig';
import { IMailDraft, INotifySettings, parseAddresses } from './notification';
import { fillText, resolveTemplate } from './mailTemplates';
import { escHtml, mailKit } from './mailHtml';
import { displayValue, fillTemplate, isMailAddress, isSafeLink } from '../utils/format';

/**
 * „Nachfrage Verwendung“: Mail an den Besitzer einer Komponente (App, Flow, Bot, Environment …),
 * ob sie noch benötigt wird. Texte aus der Vorlage „usage“ (Webpart-Einstellungen → E-Mail & Vorlagen).
 */
export interface IUsageInfo {
  /** Tabelle (für die typspezifische Vorlage), z. B. admin_flow */
  logicalName?: string;
  /** Art, z. B. „App“ */
  typeLabel: string;
  name: string;
  environment: string;
  ownerName: string;
  /** E-Mail des Besitzers (leer = unbekannt) */
  ownerEmail: string;
  /** z. B. „Zuletzt verwendet“ */
  lastLabel?: string;
  lastValue?: string;
  openUrl?: string;
  openLabel?: string;
}

/** Tabellen, für die „Nachfrage Verwendung“ angeboten wird, mit Bezeichnung und Besitzer-Spalte. */
export const USAGE_TYPES: { [logicalName: string]: { label: string; owner: string[]; last: string[] } } = {
  admin_app: { label: 'App', owner: ['_admin_appowner_value'], last: ['admin_applastlaunchedon', 'admin_appmodifiedon'] },
  admin_flow: { label: 'Flow', owner: ['_admin_derivedowner_value', '_admin_flowcreator_value'], last: ['admin_flowlastrunon', 'admin_flowmodifiedon'] },
  admin_environment: { label: 'Environment', owner: ['_admin_maker_value'], last: ['admin_environmentlastactivityon', 'admin_environmentcreatedon'] },
  admin_connector: { label: 'Connector', owner: ['_admin_connectorcreatedby_value', '_admin_connectorcreator_value'], last: ['admin_connectormodifiedon'] },
  admin_pva: { label: 'Bot', owner: ['_admin_pvaowner_value'], last: ['admin_pvalastlaunchedon', 'admin_pvamodifiedon'] },
  admin_rpa: { label: 'Desktop Flow', owner: ['_admin_rpaowner_value', '_admin_rpacreator_value'], last: ['admin_rpamodifiedon'] },
  admin_portal: { label: 'Power Pages-Site', owner: ['_admin_portalowner_value', '_admin_portalcreatedby_value'], last: ['admin_portalmodifiedon'] },
  admin_businessprocessflows: { label: 'Business Process Flow', owner: ['_admin_businessprocessflowowner_value'], last: ['admin_businessprocessflowmodifiedon'] },
  admin_solution: { label: 'Solution', owner: ['_admin_solutioncreator_value', '_admin_solutionowner_value'], last: ['admin_solutionmodifiedon', 'admin_solutioninstalledon'] },
  admin_aibuildermodel: { label: 'AI-Modell', owner: ['_admin_aicreator_value'], last: ['admin_ailastused', 'admin_aimodifiedon'] }
};

export function usageSupported(logicalName: string): boolean {
  return Object.prototype.hasOwnProperty.call(USAGE_TYPES, logicalName);
}

/** Besitzer-Lookup: bekannte Spalte, sonst die erste Personen-Spalte der Tabelle. */
export function ownerField(table: ITableConfig, r: IRecord): string {
  const known = (USAGE_TYPES[table.logicalName] || { owner: [] as string[] }).owner.filter(f => !!r[f])[0];
  if (known) {
    return known;
  }
  const person = table.columns.filter(c => c.type === 'person' && /^_.+_value$/.test(c.field) && !!r[c.field])[0];
  return person ? person.field : '';
}

/** Environment-Name aus dem Datensatz (erste Environment-Lookup-Spalte). */
export function environmentName(logicalName: string, r: IRecord): string {
  if (logicalName === 'admin_environment') {
    return displayValue(r, 'admin_displayname');
  }
  const key = Object.keys(r).filter(k => /^_admin_[a-z]*environment_value$/.test(k) && !!r[k])[0];
  return key ? displayValue(r, key) : '';
}

/** Letzte Verwendung bzw. Änderung (erste vorhandene Spalte). */
export function lastActivity(logicalName: string, r: IRecord): { label: string; value: string } | undefined {
  const f = (USAGE_TYPES[logicalName] || { last: [] as string[] }).last.filter(x => !!r[x])[0];
  if (!f) {
    return undefined;
  }
  const d = new Date(String(r[f]));
  const value = isNaN(d.getTime()) ? displayValue(r, f) : new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(d);
  return { label: /launched|lastrun|lastactivity|lastused/i.test(f) ? 'Zuletzt verwendet' : /created|installed/i.test(f) ? 'Erstellt' : 'Zuletzt geändert', value };
}

/** Link zum Öffnen: bevorzugt make.powerapps/powerautomate aus den Links der Tabelle. */
export function openLink(table: ITableConfig, r: IRecord): { url: string; label: string } | undefined {
  const links = (table.links || []).map(l => ({ label: l.label, url: fillTemplate(l.url, r) || '' })).filter(l => !!l.url && isSafeLink(l.url));
  const pick = links.filter(l => /make\.power(apps|automate)\.com|copilotstudio|powerpages/i.test(l.url))[0] || links[0];
  return pick ? { url: pick.url, label: pick.label } : undefined;
}

const DAY = 86400000;
const dateFmt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric' });

export function buildUsageMail(info: IUsageInfo, s: INotifySettings, today: Date = new Date(), accent: string = '#0f6cbd'): IMailDraft {
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const replyBy = new Date(start.getTime() + Math.max(1, s.replyDays || 14) * DAY);
  const frist = dateFmt.format(replyBy);
  const feedback = (s.feedbackUrl || '').trim()
    .replace(/\{environment\}/g, encodeURIComponent(info.name))
    .replace(/\{name\}/g, encodeURIComponent(info.name))
    .replace(/\{id\}/g, '');
  const t = resolveTemplate('usage', s.formal, !!feedback, s.templates, info.logicalName);
  const vars = { Name: info.name, Typ: info.typeLabel, Environment: info.environment || '–', Besitzer: info.ownerName || '–', Frist: frist };
  const f = (x: string): string => fillText(x, vars);
  const signature = s.signature || (s.formal ? 'Ihr Power-Platform-Team' : 'Dein Power-Platform-Team');
  const replyTitle = s.formal ? `Ihre Rückmeldung bis ${frist}` : `Deine Rückmeldung bis ${frist}`;

  const k = mailKit(accent);
  const facts = [
    k.fact('Typ', k.tx(escHtml(info.typeLabel))),
    k.fact('Name', info.openUrl ? k.link(info.openUrl, info.name, k.ink, true) : k.tx(escHtml(info.name), { bold: true })),
    k.fact('Environment', k.tx(escHtml(info.environment || '–'))),
    k.fact('Besitzer', k.tx(escHtml(info.ownerName || '–')), !info.lastValue)
  ].concat(info.lastValue ? [k.fact(info.lastLabel || 'Zuletzt verwendet', k.tx(escHtml(info.lastValue)), true)] : []);
  const html = k.frame([
    k.header('POWER PLATFORM &middot; NACHFRAGE', f(t.headline)),
    k.space(24),
    k.block(k.tx(escHtml(f(t.greeting)))),
    k.space(12),
    k.block(k.tx(escHtml(f(t.intro)))),
    k.space(20),
    k.block(k.table(`<tr>${k.dateCell('Rückmeldung bis', frist, k.ink, '#f5f5f5', '40%')}<td width="60%" style="font-size:1px;line-height:1px;">&nbsp;</td></tr>`)),
    k.space(24),
    k.heading(info.typeLabel),
    k.space(10),
    k.block(k.table(facts.join(''), `border:1px solid ${k.BORDER};`)),
    k.space(26),
    k.heading(replyTitle),
    k.space(8),
    k.block(k.tx(escHtml(f(t.question)))),
    k.space(12),
    k.block(k.table(t.options.map((o, i) => k.option(String(i + 1), f(o))).join(''))),
    feedback ? k.space(4) + k.block(k.button(feedback, 'Rückmeldung im Formular geben')) + k.space(14) : '',
    t.note ? k.space(6) + k.note(f(t.note)) : '',
    info.openUrl ? k.space(18) + k.block(k.button(info.openUrl, info.openLabel || `${info.typeLabel} öffnen`)) : '',
    k.space(24),
    k.block(`${k.tx(escHtml(f(t.closing)))}<br>${k.tx(escHtml(signature), { bold: true })}`),
    k.space(28)
  ]);

  const body = [
    f(t.greeting), '', f(t.intro), '',
    info.typeLabel.toUpperCase(),
    `• Name: ${info.name}`,
    `• Environment: ${info.environment || '–'}`,
    `• Besitzer: ${info.ownerName || '–'}`
  ].concat(info.lastValue ? [`• ${info.lastLabel || 'Zuletzt verwendet'}: ${info.lastValue}`] : [])
    .concat(info.openUrl ? [`• Öffnen: ${info.openUrl}`] : [])
    .concat(['', replyTitle.toUpperCase(), f(t.question)])
    .concat(t.options.map((o, i) => `${i + 1} – ${f(o)}`))
    .concat(feedback ? [`Formular: ${feedback}`] : [])
    .concat(t.note ? ['', f(t.note)] : [])
    .concat(['', f(t.closing), signature])
    .join('\r\n');

  const email = (info.ownerEmail || '').trim();
  return {
    to: isMailAddress(email) ? [email] : [],
    cc: parseAddresses(s.cc),
    subject: f(t.subject),
    body,
    html,
    missing: isMailAddress(email) ? [] : [info.ownerName || 'Besitzer unbekannt'],
    replyBy,
    makerUrl: info.openUrl || ''
  };
}
