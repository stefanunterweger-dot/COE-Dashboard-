import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord } from '../models/ITableConfig';
import { buildDevEnvMail, emlDraft, INotifySettings, MAILTO_LIMIT, mailtoUrl } from '../models/notification';
import { displayValue } from '../utils/format';
import { MailError, MailService, NOT_APPROVED_HINT } from '../services/MailService';
import { describeLogEntry, readNotifyLog, writeNotifyLog } from '../utils/notifyLog';

/**
 * Vorschau der E-Mail an die Maker eines Environments.
 * „In Outlook öffnen“: Empfänger und Betreff kommen per mailto ins lokale Outlook, die formatierte Mail
 * liegt in der Zwischenablage (Strg+V). „Als Outlook-Entwurf“: fertige .eml-Datei mit allem.
 * Gesendet wird immer erst in Outlook vom Admin selbst.
 */
export interface INotifyDialogProps {
  record: IRecord | undefined;
  settings: INotifySettings;
  themeVars: React.CSSProperties;
  /** Akzentfarbe des Webparts (Kopfbalken und Schaltflächen in der Mail). */
  accent: string;
  /** Automatischer Versand über Microsoft Graph (fehlt, wenn ausgeschaltet). */
  mailService?: MailService;
  onDismiss: () => void;
}

/** So viele Empfänger werden im Dialog einzeln gezeigt, der Rest als „+N“. */
const SHOW_RECIPIENTS = 12;

/** Text in die Zwischenablage (mit Rückfall für ältere Browser). */
export function copyText(text: string): Promise<boolean> {
  const fallback = (): boolean => {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try {
      ok = document.execCommand('copy');
    } catch {
      ok = false;
    }
    document.body.removeChild(ta);
    return ok;
  };
  const nav = navigator as Navigator & { clipboard?: { writeText: (t: string) => Promise<void> } };
  if (nav.clipboard && nav.clipboard.writeText) {
    return nav.clipboard.writeText(text).then(() => true, () => fallback());
  }
  return Promise.resolve(fallback());
}

/** Formatierten Text (HTML + Text) in die Zwischenablage – wird in Outlook mit Strg+V formatiert eingefügt. */
export function copyRich(html: string, text: string): Promise<boolean> {
  const fallback = (): boolean => {
    // Markierten, formatierten Inhalt kopieren (älterer Weg, von allen Browsern unterstützt)
    const div = document.createElement('div');
    div.setAttribute('contenteditable', 'true');
    div.style.position = 'fixed';
    div.style.left = '-10000px';
    div.style.top = '0';
    div.innerHTML = html; // selbst erzeugtes HTML, alle Werte maskiert
    document.body.appendChild(div);
    const range = document.createRange();
    range.selectNodeContents(div);
    const sel = window.getSelection();
    let ok = false;
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(range);
      try {
        ok = document.execCommand('copy');
      } catch {
        ok = false;
      }
      sel.removeAllRanges();
    }
    document.body.removeChild(div);
    return ok;
  };
  const w = window as unknown as { ClipboardItem?: new (items: { [type: string]: Blob }) => unknown };
  const nav = navigator as Navigator & { clipboard?: { write?: (items: unknown[]) => Promise<void> } };
  if (w.ClipboardItem && nav.clipboard && nav.clipboard.write) {
    const item = new w.ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([text], { type: 'text/plain' })
    });
    return nav.clipboard.write([item]).then(() => true, () => fallback());
  }
  return Promise.resolve(fallback());
}

/** Externen Link (mailto) öffnen, ohne dass SharePoint die Navigation abfängt. */
export function launch(url: string): void {
  const a = document.createElement('a');
  a.href = url;
  a.setAttribute('data-interception', 'off');
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

/** Datei herunterladen (der Browser bietet „Öffnen“ an → Outlook). */
function download(name: string, content: string, type: string): void {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.setAttribute('data-interception', 'off');
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/** Dateiname nur aus einfachen Zeichen (Umlaute umschreiben) – sonst nennen manche Browser die Datei „download“. */
export function safeFileName(name: string): string {
  const map: { [c: string]: string } = { 'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'Ä': 'Ae', 'Ö': 'Oe', 'Ü': 'Ue', 'ß': 'ss', '–': '-', '—': '-' };
  return name.replace(/[äöüÄÖÜß–—]/g, c => map[c])
    .replace(/[^A-Za-z0-9 ._-]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Benachrichtigung';
}

type Done = '' | 'pasted' | 'plain' | 'eml' | 'addresses' | 'copyFailed' | 'sent';

export const NotifyDialog: React.FC<INotifyDialogProps> = props => {
  const { record, settings } = props;
  const draft = React.useMemo(() => (record ? buildDevEnvMail(record, settings, new Date(), props.accent) : undefined), [record, settings, props.accent]);
  const [subject, setSubject] = React.useState<string>('');
  const [done, setDone] = React.useState<Done>('');
  const [allRecipients, setAllRecipients] = React.useState<boolean>(false);
  const [confirmSend, setConfirmSend] = React.useState<boolean>(false);
  const [sending, setSending] = React.useState<boolean>(false);
  const [sendError, setSendError] = React.useState<string>('');
  const [logText, setLogText] = React.useState<string>('');
  const envId = record ? String(record.admin_environmentid || '').toLowerCase() : '';

  React.useEffect(() => {
    setSubject(draft ? draft.subject : '');
    setDone('');
    setAllRecipients(false);
    setConfirmSend(false);
    setSending(false);
    setSendError('');
    setLogText(envId ? describeLogEntry(readNotifyLog()[envId]) : '');
  }, [draft]);

  if (!record || !draft) {
    return null;
  }
  const env = displayValue(record, 'admin_displayname');
  const noRecipients = draft.to.length === 0;
  const headerUrl = mailtoUrl(draft.to, draft.cc, subject, '');
  /** Schon Empfänger + Betreff sprengen die Länge – dann Adressen über die Zwischenablage. */
  const tooManyRecipients = !noRecipients && headerUrl.length > MAILTO_LIMIT;
  const shownTo = allRecipients ? draft.to : draft.to.slice(0, SHOW_RECIPIENTS);
  const fileName = safeFileName(`Benachrichtigung ${env}`) + '.eml';

  const openOutlook = (): void => {
    copyRich(draft.html, draft.body).then(ok => {
      writeNotifyLog(envId, 'outlook');
      if (ok) {
        launch(headerUrl);
        setDone('pasted');
        return;
      }
      // Zwischenablage gesperrt: reinen Text direkt übergeben (sofern kurz genug)
      const plain = mailtoUrl(draft.to, draft.cc, subject, draft.body);
      launch(plain.length <= MAILTO_LIMIT ? plain : headerUrl);
      setDone(plain.length <= MAILTO_LIMIT ? 'plain' : 'copyFailed');
    }, () => setDone('copyFailed'));
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={props.onDismiss}
      minWidth={720}
      maxWidth={820}
      dialogContentProps={{
        type: DialogType.normal,
        title: 'Maker per E-Mail verständigen',
        subText: `Environment „${env}“ – die E-Mail wird in deinem Outlook geöffnet und erst dort von dir gesendet.`
      }}
      modalProps={{ isBlocking: false, styles: { main: { maxWidth: 'calc(100vw - 32px) !important' } } }}
    >
      <div className={styles.notify} style={props.themeVars}>
        <div className={styles.notifyRow}>
          <span className={styles.notifyLabel}>An</span>
          <span className={styles.notifyValue}>
            {noRecipients
              ? <span className={styles.notifyWarn}><Icon iconName="Warning" /> Für dieses Environment ist im CoE kein Maker mit E-Mail-Adresse erfasst.</span>
              : shownTo.map(a => <span key={a} className={styles.notifyChip}>{a}</span>)}
            {!allRecipients && draft.to.length > SHOW_RECIPIENTS && (
              <button type="button" className={styles.textButton} onClick={() => setAllRecipients(true)}>
                +{draft.to.length - SHOW_RECIPIENTS} weitere
              </button>
            )}
          </span>
        </div>
        {draft.cc.length > 0 && (
          <div className={styles.notifyRow}>
            <span className={styles.notifyLabel}>CC</span>
            <span className={styles.notifyValue}>{draft.cc.map(a => <span key={a} className={styles.notifyChip}>{a}</span>)}</span>
          </div>
        )}
        {logText && <MessageBar messageBarType={MessageBarType.info}>Dieses Environment wurde bereits verständigt ({logText}).</MessageBar>}
        {draft.missing.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Ohne E-Mail-Adresse im CoE (nicht angeschrieben): {draft.missing.join(', ')}
          </MessageBar>
        )}
        <TextField label="Betreff" value={subject} onChange={(_e, v) => setSubject(v || '')} />
        <div>
          <div className={styles.notifyPreviewLabel}>Vorschau – so kommt die Mail an (in Outlook noch änderbar)</div>
          <iframe
            className={styles.notifyPreview}
            title="Vorschau der E-Mail"
            sandbox=""
            srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:12px;background:#ffffff;">${draft.html}</body></html>`}
          />
        </div>
        <div className={styles.notifySteps}>
          <span><b>1</b> „In Outlook öffnen“ – Empfänger und Betreff sind schon eingetragen.</span>
          <span><b>2</b> In den Mailtext klicken und <b>Strg+V</b> – die formatierte Mail wird eingefügt.</span>
          <span><b>3</b> Prüfen und senden.</span>
        </div>
        {tooManyRecipients && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Sehr viele Empfänger – Outlook übernimmt sie über den Link eventuell nicht vollständig. Am sichersten:
            „Adressen kopieren“ und in Outlook in das Feld „An“ einfügen, oder den Outlook-Entwurf verwenden.
          </MessageBar>
        )}
        {done === 'pasted' && <MessageBar messageBarType={MessageBarType.success}>Outlook wird geöffnet – die formatierte Mail liegt in der Zwischenablage: in den Text klicken und Strg+V.</MessageBar>}
        {done === 'plain' && <MessageBar messageBarType={MessageBarType.info}>Die Zwischenablage ist gesperrt – Outlook wurde mit der Textfassung geöffnet. Für die formatierte Fassung den Outlook-Entwurf verwenden.</MessageBar>}
        {done === 'eml' && <MessageBar messageBarType={MessageBarType.success}>Entwurf heruntergeladen – Datei öffnen, sie erscheint in Outlook als neue Mail mit allem.</MessageBar>}
        {done === 'addresses' && <MessageBar messageBarType={MessageBarType.success}>{draft.to.length} Adressen kopiert.</MessageBar>}
        {done === 'copyFailed' && <MessageBar messageBarType={MessageBarType.error}>Kopieren nicht möglich – bitte den Outlook-Entwurf (.eml) verwenden.</MessageBar>}
        {done === 'sent' && <MessageBar messageBarType={MessageBarType.success}>Gesendet – die Mail liegt in deinen „Gesendeten Elementen“.</MessageBar>}
        {sendError && <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{sendError}</MessageBar>}
        {confirmSend && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Die Mail wird jetzt aus deinem Postfach an {draft.to.length} Empfänger{draft.cc.length ? ` (CC: ${draft.cc.join(', ')})` : ''} gesendet.
          </MessageBar>
        )}
      </div>
      {confirmSend ? (
        <DialogFooter>
          <PrimaryButton
            text={sending ? 'Wird gesendet …' : 'Ja, jetzt senden'}
            iconProps={{ iconName: 'Send' }}
            disabled={sending}
            onClick={() => {
              if (!props.mailService) {
                return;
              }
              setSending(true);
              setSendError('');
              props.mailService.send({ to: draft.to, cc: draft.cc, subject, html: draft.html }).then(
                () => { writeNotifyLog(envId, 'sent'); setDone('sent'); setSending(false); setConfirmSend(false); },
                (e: unknown) => {
                  const err = e instanceof MailError ? e : undefined;
                  const msg = e instanceof Error ? e.message : String(e);
                  setSendError(err && err.kind === 'notApproved' ? `${NOT_APPROVED_HINT} Technische Meldung: ${msg}` : `Senden fehlgeschlagen: ${msg}`);
                  setSending(false);
                  setConfirmSend(false);
                }
              );
            }}
          />
          <DefaultButton text="Abbrechen" disabled={sending} onClick={() => setConfirmSend(false)} />
        </DialogFooter>
      ) : (
      <DialogFooter>
        <PrimaryButton text="In Outlook öffnen" iconProps={{ iconName: 'Mail' }} disabled={noRecipients} onClick={openOutlook} />
        {props.mailService && (
          <DefaultButton text="Direkt senden" iconProps={{ iconName: 'Send' }} disabled={noRecipients || done === 'sent'} onClick={() => { setConfirmSend(true); setSendError(''); }} />
        )}
        <DefaultButton
          text="Als Outlook-Entwurf (.eml)"
          iconProps={{ iconName: 'Download' }}
          disabled={noRecipients}
          title="Fertige Mail mit Empfängern, Betreff und formatiertem Text – öffnet sich in Outlook als Entwurf"
          onClick={() => {
            download(fileName, emlDraft({ to: draft.to, cc: draft.cc, subject, body: draft.body, html: draft.html }), 'message/rfc822');
            setDone('eml');
          }}
        />
        {tooManyRecipients && (
          <DefaultButton
            text="Adressen kopieren"
            iconProps={{ iconName: 'People' }}
            onClick={() => copyText(draft.to.join('; ')).then(ok => setDone(ok ? 'addresses' : 'copyFailed'), () => setDone('copyFailed'))}
          />
        )}
        <DefaultButton text="Schließen" onClick={props.onDismiss} />
      </DialogFooter>
      )}
    </Dialog>
  );
};
