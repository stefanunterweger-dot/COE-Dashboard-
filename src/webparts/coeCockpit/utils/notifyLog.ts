/**
 * Merkt sich im Browser, welche Environments der Admin schon verständigt hat (nur für ihn sichtbar),
 * damit niemand versehentlich doppelt angeschrieben wird.
 */
export interface INotifyLogEntry {
  /** Zeitpunkt (ISO). */
  at: string;
  /** „sent“ = automatisch gesendet, „outlook“ = in Outlook geöffnet. */
  how: 'sent' | 'outlook';
}

const KEY = 'coe-cockpit:notified';

export function readNotifyLog(): { [envId: string]: INotifyLogEntry } {
  try {
    const raw = window.localStorage.getItem(KEY);
    return raw ? JSON.parse(raw) as { [envId: string]: INotifyLogEntry } : {};
  } catch {
    return {};
  }
}

export function writeNotifyLog(envId: string, how: INotifyLogEntry['how']): { [envId: string]: INotifyLogEntry } {
  const log = readNotifyLog();
  if (envId) {
    log[envId.toLowerCase()] = { at: new Date().toISOString(), how };
    try {
      window.localStorage.setItem(KEY, JSON.stringify(log));
    } catch {
      // Browser-Speicher nicht verfügbar – dann nur für diese Sitzung
    }
  }
  return log;
}

export function describeLogEntry(e: INotifyLogEntry | undefined): string {
  if (!e) {
    return '';
  }
  const d = new Date(e.at);
  const when = isNaN(d.getTime()) ? '' : new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).format(d);
  return e.how === 'sent' ? `von dir gesendet am ${when}` : `von dir in Outlook geöffnet am ${when}`;
}
