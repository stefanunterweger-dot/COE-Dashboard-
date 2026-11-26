import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord } from '../models/ITableConfig';
import { buildDevEnvMail, IMailDraft, INotifySettings, mailtoUrl } from '../models/notification';
import { MAIL_TIMING, MailError, MailService, NOT_APPROVED_HINT } from '../services/MailService';
import { describeLogEntry, INotifyLogEntry, readNotifyLog, writeNotifyLog } from '../utils/notifyLog';
import { displayValue, formatNumber } from '../utils/format';
import { copyRich, launch } from './NotifyDialog';

/**
 * Serienversand: jedes ausgewählte Environment bekommt eine eigene E-Mail an seine Maker.
 * Automatisch über Microsoft Graph aus dem Postfach des angemeldeten Benutzers – oder nacheinander in Outlook.
 */
export interface IBulkNotifyPanelProps {
  /** Die aktuell gefilterten Environments (Panel offen, solange gesetzt). */
  records: IRecord[] | undefined;
  settings: INotifySettings;
  accent: string;
  themeVars: React.CSSProperties;
  /** Fehlt, wenn der automatische Versand in den Eigenschaften ausgeschaltet ist. */
  mailService: MailService | undefined;
  onDismiss: () => void;
}

type Status = 'sending' | 'sent' | 'failed' | 'opened';
type Mode = 'select' | 'confirm' | 'sending' | 'done' | 'steps';

interface IItem { id: string; record: IRecord; draft: IMailDraft; name: string }

const wait = (ms: number): Promise<void> => new Promise(resolve => window.setTimeout(resolve, ms));
const idOf = (r: IRecord): string => String(r.admin_environmentid || '').toLowerCase();

export const BulkNotifyPanel: React.FC<IBulkNotifyPanelProps> = props => {
  const { records, settings, accent, mailService } = props;
  const items = React.useMemo<IItem[]>(() => (records || []).map(r => ({
    id: idOf(r),
    record: r,
    draft: buildDevEnvMail(r, settings, new Date(), accent),
    name: displayValue(r, 'admin_displayname')
  })), [records, settings, accent]);

  const [log, setLog] = React.useState<{ [id: string]: INotifyLogEntry }>({});
  const [selected, setSelected] = React.useState<{ [id: string]: boolean }>({});
  const [status, setStatus] = React.useState<{ [id: string]: { s: Status; error?: string } }>({});
  const [previewId, setPreviewId] = React.useState<string>('');
  const [mode, setMode] = React.useState<Mode>('select');
  const [step, setStep] = React.useState<number>(0);
  /** Technische Meldung, wenn Mail.Send fehlt (leer = kein Problem). */
  const [notApproved, setNotApproved] = React.useState<string>('');
  const [stepNote, setStepNote] = React.useState<string>('');
  const [run, setRun] = React.useState<{ total: number; done: number }>({ total: 0, done: 0 });
  const stopRef = React.useRef<boolean>(false);

  // Beim Öffnen: alle mit Empfängern auswählen, die noch nicht verständigt wurden
  React.useEffect(() => {
    if (!records) {
      return;
    }
    const l = readNotifyLog();
    setLog(l);
    const sel: { [id: string]: boolean } = {};
    items.forEach(i => { sel[i.id] = i.draft.to.length > 0 && !l[i.id]; });
    setSelected(sel);
    setStatus({});
    setMode('select');
    setStep(0);
    setNotApproved('');
    setStepNote('');
    setPreviewId((items.filter(i => sel[i.id])[0] || items[0] || { id: '' }).id);
  }, [records]);

  if (!records) {
    return null;
  }

  const chosen = items.filter(i => selected[i.id] && i.draft.to.length > 0);
  const withoutRecipients = items.filter(i => i.draft.to.length === 0).length;
  const recipientCount = chosen.reduce((set, i) => { i.draft.to.forEach(a => { set[a] = true; }); return set; }, {} as { [a: string]: boolean });
  const preview = items.filter(i => i.id === previewId)[0];
  const current = mode === 'steps' ? chosen[step] : undefined;
  const shown = current || preview;
  const sentCount = items.filter(i => status[i.id] && status[i.id].s === 'sent').length;
  const failed = items.filter(i => status[i.id] && status[i.id].s === 'failed');
  const busy = mode === 'sending';

  const setAll = (fn: (i: IItem) => boolean): void => {
    const sel: { [id: string]: boolean } = {};
    items.forEach(i => { sel[i.id] = i.draft.to.length > 0 && fn(i); });
    setSelected(sel);
  };

  const sendAll = async (list: IItem[]): Promise<void> => {
    if (!mailService) {
      return;
    }
    stopRef.current = false;
    setNotApproved('');
    setRun({ total: list.length, done: 0 });
    setMode('sending');
    for (let n = 0; n < list.length; n++) {
      if (stopRef.current) {
        break;
      }
      const it = list[n];
      setPreviewId(it.id);
      setStatus(prev => ({ ...prev, [it.id]: { s: 'sending' } }));
      try {
        await mailService.send({ to: it.draft.to, cc: it.draft.cc, subject: it.draft.subject, html: it.draft.html });
        setStatus(prev => ({ ...prev, [it.id]: { s: 'sent' } }));
        setLog(writeNotifyLog(it.id, 'sent'));
        setSelected(prev => ({ ...prev, [it.id]: false }));
      } catch (e) {
        const err = e instanceof MailError ? e : new MailError(e instanceof Error ? e.message : String(e), 'failed', 0);
        setStatus(prev => ({ ...prev, [it.id]: { s: 'failed', error: err.message } }));
        if (err.kind === 'notApproved') {
          setNotApproved(err.message || `HTTP ${err.status}`);
          break; // ohne Berechtigung scheitern alle weiteren ebenso
        }
      }
      setRun(prev => ({ ...prev, done: prev.done + 1 }));
      if (n < list.length - 1 && !stopRef.current) {
        await wait(MAIL_TIMING.intervalMs);
      }
    }
    setMode('done');
  };

  const openInOutlook = (it: IItem): void => {
    copyRich(it.draft.html, it.draft.body).then(ok => {
      launch(mailtoUrl(it.draft.to, it.draft.cc, it.draft.subject, ok ? '' : it.draft.body));
      setStatus(prev => ({ ...prev, [it.id]: { s: 'opened' } }));
      setLog(writeNotifyLog(it.id, 'outlook'));
      setStepNote(ok ? 'Outlook wird geöffnet – in den Mailtext klicken, Strg+V, senden. Danach „Weiter“.' : 'Outlook wird mit der Textfassung geöffnet (Zwischenablage gesperrt).');
    }, () => setStepNote('Kopieren nicht möglich.'));
  };

  const statusCell = (it: IItem): JSX.Element => {
    const st = status[it.id];
    if (it.draft.to.length === 0) {
      return <span className={styles.bulkMuted}><Icon iconName="Blocked2" /> keine E-Mail-Adresse</span>;
    }
    if (st && st.s === 'sending') {
      return <span className={styles.bulkInfo}><Icon iconName="Send" /> wird gesendet …</span>;
    }
    if (st && st.s === 'sent') {
      return <span className={styles.bulkOk}><Icon iconName="CheckMark" /> gesendet</span>;
    }
    if (st && st.s === 'opened') {
      return <span className={styles.bulkOk}><Icon iconName="Mail" /> in Outlook geöffnet</span>;
    }
    if (st && st.s === 'failed') {
      return <span className={styles.bulkErr} title={st.error}><Icon iconName="ErrorBadge" /> Fehler</span>;
    }
    return log[it.id] ? <span className={styles.bulkMuted}>{describeLogEntry(log[it.id])}</span> : <span className={styles.bulkMuted}>offen</span>;
  };

  const footer = (): JSX.Element => {
    if (mode === 'steps') {
      return (
        <div className={styles.bulkFooter}>
          <span className={styles.bulkStepInfo}>Environment {formatNumber(Math.min(step + 1, chosen.length))} von {formatNumber(chosen.length)}{current ? `: ${current.name}` : ''}</span>
          <PrimaryButton text="In Outlook öffnen" iconProps={{ iconName: 'Mail' }} disabled={!current} onClick={() => current && openInOutlook(current)} />
          <DefaultButton text="Zurück" iconProps={{ iconName: 'ChevronLeft' }} disabled={step === 0} onClick={() => { setStep(step - 1); setStepNote(''); }} />
          <DefaultButton text={step + 1 >= chosen.length ? 'Fertig' : 'Weiter'} iconProps={{ iconName: 'ChevronRight' }}
            onClick={() => { if (step + 1 >= chosen.length) { setMode('done'); } else { setStep(step + 1); setStepNote(''); } }} />
        </div>
      );
    }
    if (mode === 'confirm') {
      return (
        <div className={styles.bulkFooter}>
          <PrimaryButton text={`Ja, ${formatNumber(chosen.length)} E-Mails jetzt senden`} iconProps={{ iconName: 'Send' }} onClick={() => { sendAll(chosen).catch(() => setMode('done')); }} />
          <DefaultButton text="Abbrechen" onClick={() => setMode('select')} />
        </div>
      );
    }
    if (mode === 'sending') {
      return (
        <div className={styles.bulkFooter}>
          <DefaultButton text="Anhalten" iconProps={{ iconName: 'CircleStop' }} onClick={() => { stopRef.current = true; }} />
        </div>
      );
    }
    return (
      <div className={styles.bulkFooter}>
        {mailService && (
          <PrimaryButton text={`Automatisch senden (${formatNumber(chosen.length)})`} iconProps={{ iconName: 'Send' }} disabled={chosen.length === 0} onClick={() => setMode('confirm')} />
        )}
        {mailService && mode === 'done' && failed.length > 0 && !notApproved && (
          <DefaultButton text={`Fehlgeschlagene erneut senden (${formatNumber(failed.length)})`} iconProps={{ iconName: 'Refresh' }} onClick={() => { sendAll(failed).catch(() => setMode('done')); }} />
        )}
        <DefaultButton text="Nacheinander in Outlook" iconProps={{ iconName: 'OutlookLogo' }} disabled={chosen.length === 0}
          onClick={() => { setStep(0); setStepNote(''); setMode('steps'); }} />
        <DefaultButton text="Schließen" onClick={props.onDismiss} />
      </div>
    );
  };

  return (
    <Panel
      isOpen={true}
      type={PanelType.custom}
      customWidth="1040px"
      headerText="Alle gefilterten Environments verständigen"
      isBlocking={true}
      isLightDismiss={false}
      onDismiss={(ev?: React.SyntheticEvent<HTMLElement>) => { if (busy) { ev && ev.preventDefault && ev.preventDefault(); return; } props.onDismiss(); }}
      onRenderFooterContent={footer}
      isFooterAtBottom={true}
      closeButtonAriaLabel="Schließen"
    >
      <div className={styles.bulk} style={props.themeVars}>
        <p className={styles.bulkIntro}>
          Jedes Environment bekommt eine <b>eigene E-Mail</b> an seine Maker – mit denselben Inhalten wie die Einzelmail.
          {mailService ? ' Automatisch gesendet wird aus deinem Postfach (Gesendete Elemente); Antworten kommen zu dir.' : ''}
          {settings.cc ? ` CC: ${settings.cc}.` : ''}
        </p>
        <div className={styles.bulkStats}>
          <span className={styles.notifyChip}><b>{formatNumber(chosen.length)}</b> ausgewählt</span>
          <span className={styles.notifyChip}><b>{formatNumber(Object.keys(recipientCount).length)}</b> Empfänger</span>
          {withoutRecipients > 0 && <span className={styles.notifyChip}><b>{formatNumber(withoutRecipients)}</b> ohne E-Mail-Adresse (übersprungen)</span>}
          {sentCount > 0 && <span className={styles.notifyChip}><b>{formatNumber(sentCount)}</b> gesendet</span>}
        </div>

        {notApproved && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
            {NOT_APPROVED_HINT}
            <div className={styles.techMessage}>Technische Meldung: {notApproved}</div>
          </MessageBar>
        )}
        {mode === 'confirm' && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Es werden jetzt <b>{formatNumber(chosen.length)} E-Mails</b> an {formatNumber(Object.keys(recipientCount).length)} Empfänger aus deinem Postfach gesendet –
            eine pro Environment, im Abstand von etwa {Math.round(MAIL_TIMING.intervalMs / 1000)} Sekunden (Exchange-Limit).
            Dauer ca. {formatNumber(Math.max(1, Math.ceil(chosen.length * MAIL_TIMING.intervalMs / 60000)))} Min.
          </MessageBar>
        )}
        {(mode === 'sending' || (mode === 'done' && run.total > 0)) && (
          <ProgressIndicator
            label={mode === 'sending'
              ? `Versand läuft … ${formatNumber(run.done)} von ${formatNumber(run.total)}`
              : `Fertig: ${formatNumber(sentCount)} gesendet${failed.length ? `, ${formatNumber(failed.length)} fehlgeschlagen` : ''}${run.done < run.total ? ' (angehalten)' : ''}`}
            percentComplete={run.total ? run.done / run.total : 1}
          />
        )}
        {mode === 'steps' && (
          <MessageBar messageBarType={stepNote ? MessageBarType.success : MessageBarType.info} isMultiline={true}>
            {stepNote || '„In Outlook öffnen“ trägt Empfänger und Betreff ein, die formatierte Mail liegt in der Zwischenablage: in den Mailtext klicken, Strg+V, senden – dann „Weiter“.'}
          </MessageBar>
        )}

        <div className={styles.bulkBody}>
          <div className={styles.bulkList}>
            {mode === 'select' && (
              <div className={styles.bulkListTools}>
                <button type="button" className={styles.textButton} onClick={() => setAll(() => true)}>Alle</button>
                <button type="button" className={styles.textButton} onClick={() => setAll(i => !log[i.id])}>Nur noch nicht verständigte</button>
                <button type="button" className={styles.textButton} onClick={() => setAll(() => false)}>Keine</button>
              </div>
            )}
            <table className={styles.bulkTable}>
              <thead>
                <tr><th aria-label="Auswahl" /><th>Environment</th><th>Empfänger</th><th>Status</th></tr>
              </thead>
              <tbody>
                {items.map(it => (
                  <tr
                    key={it.id || it.name}
                    className={`${shown && shown.id === it.id ? styles.bulkRowActive : ''}`}
                    onClick={() => { if (mode !== 'steps') { setPreviewId(it.id); } }}
                  >
                    <td onClick={e => e.stopPropagation()}>
                      <Checkbox
                        ariaLabel={`${it.name} auswählen`}
                        checked={!!selected[it.id] && it.draft.to.length > 0}
                        disabled={it.draft.to.length === 0 || mode !== 'select'}
                        onChange={(_e, v) => setSelected(prev => ({ ...prev, [it.id]: !!v }))}
                      />
                    </td>
                    <td className={styles.bulkName}>{it.name}</td>
                    <td title={it.draft.to.join('; ')}>{formatNumber(it.draft.to.length)}</td>
                    <td title={status[it.id] && status[it.id].error ? status[it.id].error : undefined}>
                      {statusCell(it)}
                      {status[it.id] && status[it.id].s === 'failed' && status[it.id].error && <div className={styles.techMessage}>{status[it.id].error}</div>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className={styles.bulkPreview}>
            {shown ? (
              <>
                <div className={styles.notifyPreviewLabel}>An: {shown.draft.to.length ? shown.draft.to.slice(0, 4).join('; ') + (shown.draft.to.length > 4 ? ` … (+${shown.draft.to.length - 4})` : '') : '—'}</div>
                <div className={styles.bulkSubject}>{shown.draft.subject}</div>
                <iframe
                  className={styles.notifyPreview}
                  title="Vorschau der E-Mail"
                  sandbox=""
                  srcDoc={`<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:12px;background:#ffffff;">${shown.draft.html}</body></html>`}
                />
              </>
            ) : <div className={styles.bulkMuted}>Kein Environment gewählt.</div>}
          </div>
        </div>
      </div>
    </Panel>
  );
};
