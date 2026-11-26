import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IRecord, ITableConfig } from '../models/ITableConfig';
import { INotifySettings } from '../models/notification';
import { buildUsageMail, environmentName, IUsageInfo, lastActivity, openLink, ownerField, USAGE_TYPES } from '../models/usageMail';
import { mailAddress } from '../models/devEnvironments';
import { DataverseService } from '../services/DataverseService';
import { MailService } from '../services/MailService';
import { INotifyEntry, NotifyMethod } from '../services/NotifyLogService';
import { NotifyDialog } from './NotifyDialog';
import { displayValue } from '../utils/format';

/**
 * „Nachfrage Verwendung“: lädt den vollständigen Datensatz und den Besitzer (CoE-Maker mit E-Mail),
 * baut die Mail aus der Vorlage und öffnet den bekannten E-Mail-Dialog (Outlook-Entwurf, Zwischenablage, direkt senden).
 */
export interface IUsageTarget { table: ITableConfig; record: IRecord; primaryId: string }

/** Schlüssel im Versandprotokoll (Spalte EnvironmentId), getrennt von den Environment-Benachrichtigungen. */
export function usageLogKey(logicalName: string, id: string): string {
  return `usage:${logicalName}:${id}`.toLowerCase();
}

export interface IUsageNotifyProps {
  target: IUsageTarget | undefined;
  service: DataverseService | undefined;
  settings: INotifySettings;
  themeVars: React.CSSProperties;
  accent: string;
  mailService?: MailService;
  /** Letzte Einträge im Versandprotokoll (Schlüssel siehe usageLogKey). */
  log: { [key: string]: INotifyEntry };
  logProblem?: string;
  onLogged: (key: string, title: string, method: NotifyMethod, subject: string, to: string[]) => Promise<void>;
  onDismiss: () => void;
}

export const UsageNotify: React.FC<IUsageNotifyProps> = props => {
  const { target, service } = props;
  const [state, setState] = React.useState<{ status: 'loading' | 'ready' | 'error'; info?: IUsageInfo; record?: IRecord; error?: string }>({ status: 'loading' });

  React.useEffect(() => {
    if (!target || !service) {
      return;
    }
    let cancelled = false;
    setState({ status: 'loading' });
    const id = String(target.record[target.primaryId] || '');
    (async () => {
      const full = id ? await service.getRecord(target.table.logicalName, id).catch(() => target.record) : target.record;
      const rec: IRecord = { ...target.record, ...full };
      const field = ownerField(target.table, rec);
      const ownerId = field ? String(rec[field] || '') : '';
      let ownerName = field ? displayValue(rec, field) : '';
      let ownerEmail = '';
      if (ownerId) {
        try {
          const maker = await service.getRecord('admin_maker', ownerId);
          ownerEmail = mailAddress(maker);
          ownerName = ownerName || displayValue(maker, 'admin_displayname');
        } catch {
          // Besitzer nicht im CoE-Maker-Inventar → ohne Adresse
        }
      }
      const type = USAGE_TYPES[target.table.logicalName];
      const last = lastActivity(target.table.logicalName, rec);
      const open = openLink(target.table, rec);
      const name = displayValue(rec, 'admin_displayname') || displayValue(rec, 'admin_name')
        || (target.table.columns[0] ? displayValue(rec, target.table.columns[0].field) : '') || id;
      const info: IUsageInfo = {
        typeLabel: type ? type.label : target.table.title,
        name,
        environment: environmentName(target.table.logicalName, rec),
        ownerName,
        ownerEmail,
        lastLabel: last ? last.label : undefined,
        lastValue: last ? last.value : undefined,
        openUrl: open ? open.url : undefined,
        openLabel: open ? open.label : undefined
      };
      if (!cancelled) {
        setState({ status: 'ready', info, record: rec });
      }
    })().catch((e: Error) => { if (!cancelled) { setState({ status: 'error', error: e.message }); } });
    return () => { cancelled = true; };
  }, [target, service]);

  const draft = React.useMemo(
    () => (state.info ? buildUsageMail(state.info, props.settings, new Date(), props.accent) : undefined),
    [state.info, props.settings, props.accent]
  );

  if (!target) {
    return null;
  }
  if (state.status !== 'ready' || !state.info || !state.record || !draft) {
    return (
      <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ type: DialogType.normal, title: 'Nachfrage Verwendung' }} minWidth={420}>
        {state.status === 'error'
          ? <span>Die Daten konnten nicht geladen werden: {state.error}</span>
          : <Spinner size={SpinnerSize.medium} label="Besitzer und Details werden ermittelt …" />}
        <DialogFooter><DefaultButton text="Schließen" onClick={props.onDismiss} /></DialogFooter>
      </Dialog>
    );
  }
  const id = String(target.record[target.primaryId] || '');
  const key = usageLogKey(target.table.logicalName, id);
  const title = `${state.info.typeLabel}: ${state.info.name}`;
  return (
    <NotifyDialog
      record={state.record}
      draft={draft}
      settings={props.settings}
      themeVars={props.themeVars}
      accent={props.accent}
      mailService={props.mailService}
      notified={props.log[key]}
      logProblem={props.logProblem}
      title="Nachfrage Verwendung"
      subText={`${state.info.typeLabel} „${state.info.name}“ – Frage an den Besitzer, ob die Komponente noch benötigt wird. Die E-Mail wird in deinem Outlook geöffnet oder direkt gesendet.`}
      fileLabel={`Nachfrage ${state.info.name}`}
      noRecipientText={`Für ${state.info.ownerName ? `„${state.info.ownerName}“` : 'diese Komponente'} ist im CoE keine E-Mail-Adresse des Besitzers erfasst.`}
      alreadyLabel="Bereits nachgefragt"
      onLogged={(_r, method, subject, to) => props.onLogged(key, title, method, subject, to)}
      onDismiss={props.onDismiss}
    />
  );
};
