import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord } from '../models/ITableConfig';
import { buildDevEnvMail, INotifySettings, MAILTO_LIMIT, mailtoUrl } from '../models/notification';
import { displayValue } from '../utils/format';

/**
 * Vorschau der E-Mail an die Maker eines Environments. „In Outlook öffnen“ übergibt An, CC, Betreff
 * und Text per mailto an das lokale Outlook; gesendet wird dort vom Admin selbst.
 */
export interface INotifyDialogProps {
  record: IRecord | undefined;
  settings: INotifySettings;
  themeVars: React.CSSProperties;
  onDismiss: () => void;
}

const PASTE_HINT = '[Text wurde in die Zwischenablage kopiert – hier mit Strg+V einfügen]';
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

export const NotifyDialog: React.FC<INotifyDialogProps> = props => {
  const { record, settings } = props;
  const draft = React.useMemo(() => (record ? buildDevEnvMail(record, settings) : undefined), [record, settings]);
  const [subject, setSubject] = React.useState<string>('');
  const [body, setBody] = React.useState<string>('');
  const [done, setDone] = React.useState<'' | 'opened' | 'pasted' | 'copied' | 'addresses' | 'copyFailed'>('');
  const [allRecipients, setAllRecipients] = React.useState<boolean>(false);

  React.useEffect(() => {
    setSubject(draft ? draft.subject : '');
    setBody(draft ? draft.body : '');
    setDone('');
    setAllRecipients(false);
  }, [draft]);

  if (!record || !draft) {
    return null;
  }
  const env = displayValue(record, 'admin_displayname');
  const fullUrl = mailtoUrl(draft.to, draft.cc, subject, body);
  const tooLong = fullUrl.length > MAILTO_LIMIT;
  const href = tooLong ? mailtoUrl(draft.to, draft.cc, subject, PASTE_HINT) : fullUrl;
  const noRecipients = draft.to.length === 0;
  /** Schon die Empfänger allein sprengen die Länge – dann Adressen über die Zwischenablage. */
  const tooManyRecipients = !noRecipients && href.length > MAILTO_LIMIT;
  const shownTo = allRecipients ? draft.to : draft.to.slice(0, SHOW_RECIPIENTS);

  return (
    <Dialog
      hidden={false}
      onDismiss={props.onDismiss}
      minWidth={640}
      maxWidth={760}
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
        {draft.missing.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Ohne E-Mail-Adresse im CoE (nicht angeschrieben): {draft.missing.join(', ')}
          </MessageBar>
        )}
        <TextField label="Betreff" value={subject} onChange={(_e, v) => setSubject(v || '')} />
        <TextField
          label={`Text (Rückmeldung bis ${new Intl.DateTimeFormat('de-AT').format(draft.replyBy)})`}
          multiline={true}
          rows={14}
          value={body}
          onChange={(_e, v) => setBody(v || '')}
          styles={{ field: { fontFamily: 'inherit', fontSize: 13, lineHeight: '1.45' } }}
        />
        {tooManyRecipients && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Sehr viele Empfänger – Outlook übernimmt sie über den Link eventuell nicht vollständig. Am sichersten:
            „Adressen kopieren“ und in Outlook in das Feld „An“ einfügen.
          </MessageBar>
        )}
        {tooLong && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline={true}>
            Der Text ist für die direkte Übergabe an Outlook zu lang. Beim Öffnen wird er in die Zwischenablage kopiert –
            in Outlook einfach mit Strg+V einfügen.
          </MessageBar>
        )}
        {done === 'opened' && <MessageBar messageBarType={MessageBarType.success}>Outlook wurde geöffnet – bitte prüfen und senden.</MessageBar>}
        {done === 'pasted' && <MessageBar messageBarType={MessageBarType.success}>Outlook wurde geöffnet und der Text kopiert – in der Mail mit Strg+V einfügen.</MessageBar>}
        {done === 'copied' && <MessageBar messageBarType={MessageBarType.success}>Text kopiert.</MessageBar>}
        {done === 'addresses' && <MessageBar messageBarType={MessageBarType.success}>{draft.to.length} Adressen kopiert.</MessageBar>}
        {done === 'copyFailed' && <MessageBar messageBarType={MessageBarType.error}>Kopieren nicht möglich – bitte den Text oben markieren und kopieren.</MessageBar>}
      </div>
      <DialogFooter>
        <PrimaryButton
          text="In Outlook öffnen"
          iconProps={{ iconName: 'Mail' }}
          href={noRecipients ? undefined : href}
          disabled={noRecipients}
          data-interception="off"
          onClick={() => {
            if (tooLong) {
              copyText(body).then(ok => setDone(ok ? 'pasted' : 'copyFailed'), () => setDone('copyFailed'));
            } else {
              setDone('opened');
            }
          }}
        />
        {tooManyRecipients && (
          <DefaultButton
            text="Adressen kopieren"
            iconProps={{ iconName: 'People' }}
            onClick={() => copyText(draft.to.join('; ')).then(ok => setDone(ok ? 'addresses' : 'copyFailed'), () => setDone('copyFailed'))}
          />
        )}
        <DefaultButton
          text="Text kopieren"
          iconProps={{ iconName: 'Copy' }}
          onClick={() => copyText(`${subject}\r\n\r\n${body}`).then(ok => setDone(ok ? 'copied' : 'copyFailed'), () => setDone('copyFailed'))}
        />
        <DefaultButton text="Schließen" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
};
