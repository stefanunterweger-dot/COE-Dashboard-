/**
 * Bausteine für „Outlook-feste“ HTML-Mails: Outlook (klassisch) stellt HTML mit der Word-Engine dar und ignoriert
 * dabei vieles (div-Formate, Abstände von Absätzen, runde Ecken, Linkfarben). Deshalb nur Tabellen,
 * Zellen-Innenabstände, eigene Abstandszeilen, Farben direkt am Text (span + font) und keine runden Elemente.
 */
export const escHtml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Kontrastfarbe für Text auf der Akzentfarbe (Weiß nur bei Kontrast ≥ 4.5:1). */
export function onColor(hex: string): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) {
    return '#ffffff';
  }
  const n = parseInt(m[1], 16);
  const lin = (c: number): number => { const v = c / 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
  const L = 0.2126 * lin((n >> 16) & 255) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
  return L > 0.183 ? '#1b1b1b' : '#ffffff';
}

export interface IMailKit {
  accent: string;
  /** Textfarbe auf der Akzentfarbe */
  on: string;
  /** Farbe für Links und Akzente auf weißem Grund */
  ink: string;
  TEXT: string;
  MUTED: string;
  BORDER: string;
  tx: (html: string, o?: { size?: number; color?: string; bold?: boolean; spacing?: number }) => string;
  link: (href: string, label: string, color: string, bold?: boolean) => string;
  space: (h: number) => string;
  block: (html: string) => string;
  table: (rows: string, extra?: string) => string;
  button: (href: string, label: string) => string;
  dateCell: (label: string, value: string, color: string, bg: string, width?: string) => string;
  gap: string;
  fact: (label: string, valueHtml: string, last?: boolean) => string;
  option: (n: string, text: string) => string;
  heading: (text: string) => string;
  note: (text: string) => string;
  header: (eyebrow: string, title: string) => string;
  frame: (rows: string[]) => string;
}

export function mailKit(accentIn: string): IMailKit {
  const accent = /^#[0-9a-f]{6}$/i.test(accentIn) ? accentIn : '#0f6cbd';
  const on = onColor(accent);
  const ink = on === '#ffffff' ? accent : '#0f6cbd';
  const FONT = "'Segoe UI', Calibri, Arial, sans-serif";
  const TEXT = '#242424';
  const MUTED = '#616161';
  const BORDER = '#e0e0e0';
  const tx = (html: string, o: { size?: number; color?: string; bold?: boolean; spacing?: number } = {}): string => {
    const size = o.size || 14;
    const color = o.color || TEXT;
    const inner = o.bold ? `<b>${html}</b>` : html;
    return `<span style="font-family:${FONT};font-size:${size}px;line-height:${Math.round(size * 1.45)}px;color:${color};` +
      `${o.bold ? 'font-weight:bold;' : ''}${o.spacing ? `letter-spacing:${o.spacing}px;` : ''}"><font face="Segoe UI, Calibri, Arial, sans-serif" color="${color}">${inner}</font></span>`;
  };
  const link = (href: string, label: string, color: string, bold?: boolean): string =>
    `<a href="${escHtml(href)}" style="color:${color};text-decoration:underline;">${tx(escHtml(label), { color, bold })}</a>`;
  const space = (h: number): string =>
    `<tr><td height="${h}" style="height:${h}px;font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</td></tr>`;
  const block = (html: string): string => `<tr><td style="padding:0 28px;">${html}</td></tr>`;
  const table = (rows: string, extra?: string): string =>
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;${extra || ''}">${rows}</table>`;
  const button = (href: string, label: string): string =>
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr>` +
    `<td bgcolor="${accent}" style="background-color:${accent};padding:11px 22px;">` +
    `<a href="${escHtml(href)}" style="color:${on};text-decoration:none;">${tx(`${escHtml(label)}&nbsp;&#8594;`, { color: on, bold: true })}</a>` +
    `</td></tr></table>`;
  const dateCell = (label: string, value: string, color: string, bg: string, width: string = '32%'): string =>
    `<td width="${width}" valign="top" bgcolor="${bg}" style="background-color:${bg};border-left:4px solid ${color};padding:10px 14px;">` +
    `${tx(escHtml(label), { size: 12, color: MUTED })}<br>${tx(escHtml(value), { size: 18, color, bold: true })}</td>`;
  const gap = '<td width="2%" style="font-size:1px;line-height:1px;">&nbsp;</td>';
  const fact = (label: string, valueHtml: string, last?: boolean): string => {
    const line = last ? '' : `border-bottom:1px solid ${BORDER};`;
    return `<tr><td width="34%" valign="top" style="padding:9px 14px;${line}">${tx(escHtml(label), { size: 13, color: MUTED })}</td>` +
      `<td valign="top" style="padding:9px 14px;${line}">${valueHtml}</td></tr>`;
  };
  const option = (n: string, text: string): string =>
    `<tr><td width="30" valign="top" style="padding:0 0 10px 0;">${tx(`${n}.`, { size: 16, color: ink, bold: true })}</td>` +
    `<td valign="top" style="padding:1px 0 10px 0;">${tx(escHtml(text))}</td></tr>`;
  const heading = (text: string): string => block(tx(escHtml(text), { size: 16, bold: true }));
  const note = (text: string): string =>
    block(table(`<tr><td bgcolor="#f5f5f5" style="background-color:#f5f5f5;border-left:4px solid #8a8886;padding:12px 16px;">${tx(escHtml(text), { size: 13, color: '#424242' })}</td></tr>`));
  const header = (eyebrow: string, title: string): string =>
    `<tr><td bgcolor="${accent}" style="background-color:${accent};padding:20px 28px 22px 28px;">` +
    `${tx(eyebrow, { size: 12, color: on, spacing: 1 })}<br>${tx(escHtml(title), { size: 20, color: on, bold: true })}</td></tr>`;
  const frame = (rows: string[]): string =>
    `<table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" style="width:640px;border-collapse:collapse;border:1px solid ${BORDER};background-color:#ffffff;">` +
    rows.join('') + '</table>';
  return { accent, on, ink, TEXT, MUTED, BORDER, tx, link, space, block, table, button, dateCell, gap, fact, option, heading, note, header, frame };
}
