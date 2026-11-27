import * as React from 'react';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { Pill } from './Cells';
import { activateRole, deactivateRole, IMyRole, IMyRoles, loadMyRoles, PimError } from '../models/pim';
import { PIM_URL } from '../services/PowerPlatformApiService';
import { formatDate, formatTime } from '../utils/format';

/**
 * „Meine Rollen“: eigene, in PIM berechtigte Entra-Rollen live (Microsoft Graph) – aktivieren und vorzeitig beenden.
 * Nur Selbst-Aktionen der angemeldeten Person; Microsoft prüft dabei die PIM-Richtlinie der Rolle und protokolliert im Entra-Audit-Log.
 */
export interface IMyRolesPanelProps {
  graphFactory: MSGraphClientFactory;
  themeVars: React.CSSProperties;
  onDismiss: () => void;
  /** nach Aktivieren/Deaktivieren: Rollenstatus (Token) neu prüfen lassen */
  onChanged: () => void;
}

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));
const hoursText = (min: number): string => (min % 60 === 0 ? `${min / 60} Std.` : min < 60 ? `${min} Min.` : `${Math.floor(min / 60)} Std. ${min % 60} Min.`);

export const MyRolesPanel: React.FC<IMyRolesPanelProps> = props => {
  const [data, setData] = React.useState<IMyRoles | undefined>(undefined);
  const [error, setError] = React.useState<PimError | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [form, setForm] = React.useState<string>('');
  const [minutes, setMinutes] = React.useState<number>(60);
  const [reason, setReason] = React.useState<string>('');
  const [ticketNo, setTicketNo] = React.useState<string>('');
  const [ticketSys, setTicketSys] = React.useState<string>('');
  const [busy, setBusy] = React.useState<string>('');
  const [notice, setNotice] = React.useState<{ type: MessageBarType; text: string; key?: string } | undefined>(undefined);
  const [search, setSearch] = React.useState<string>('');

  const load = (): void => {
    setLoading(true);
    setError(undefined);
    loadMyRoles(props.graphFactory).then(d => { setData(d); setLoading(false); }, e => {
      setError(e instanceof PimError ? e : new PimError(errText(e), 'failed'));
      setLoading(false);
    });
  };
  React.useEffect(() => { load(); }, [props.graphFactory]);

  const openForm = (r: IMyRole): void => {
    setForm(r.key);
    const max = r.policy ? r.policy.maxMinutes : 8 * 60;
    setMinutes(Math.min(max, 60));
    setReason('');
    setTicketNo('');
    setTicketSys('');
    setNotice(undefined);
  };

  const activate = async (r: IMyRole): Promise<void> => {
    if (!data) {
      return;
    }
    setBusy(r.key);
    setNotice(undefined);
    try {
      const status = await activateRole(props.graphFactory, { principalId: data.principalId, role: r, minutes, justification: reason, ticketNumber: ticketNo, ticketSystem: ticketSys });
      const pending = /pending.*approval|approval/i.test(status);
      setNotice({
        type: pending ? MessageBarType.warning : MessageBarType.success,
        key: r.key,
        text: pending
          ? `„${r.name}“: Anfrage gestellt – die Rolle wird aktiv, sobald sie genehmigt ist.`
          : `„${r.name}“ ist für ${hoursText(minutes)} aktiviert. Bis die Rolle in der Anmeldung steht, dauert es oft 1–2 Minuten – das Cockpit prüft automatisch nach.`
      });
      setForm('');
      props.onChanged();
      load();
    } catch (e) {
      setNotice({ type: MessageBarType.error, key: r.key, text: errText(e) });
    }
    setBusy('');
  };

  const deactivate = async (r: IMyRole): Promise<void> => {
    if (!data) {
      return;
    }
    setBusy(r.key);
    setNotice(undefined);
    try {
      await deactivateRole(props.graphFactory, data.principalId, r);
      setNotice({ type: MessageBarType.success, key: r.key, text: `„${r.name}“ wurde beendet.` });
      props.onChanged();
      load();
    } catch (e) {
      setNotice({ type: MessageBarType.error, key: r.key, text: errText(e) });
    }
    setBusy('');
  };

  const durationOptions = (r: IMyRole): { key: number; text: string }[] => {
    const max = r.policy ? r.policy.maxMinutes : 8 * 60;
    const steps = [30, 60, 120, 180, 240, 360, 480, 600, 720, 1440].filter(m => m <= max);
    if (steps.indexOf(max) < 0) {
      steps.push(max);
    }
    return steps.map(m => ({ key: m, text: `${hoursText(m)}${m === max ? ' (Maximum)' : ''}` }));
  };

  const q = search.trim().toLowerCase();
  const roles = (data ? data.roles : []).filter(r => !q || r.name.toLowerCase().indexOf(q) >= 0);
  const eligibleCount = data ? data.roles.filter(r => r.eligible).length : 0;

  const card = (r: IMyRole): JSX.Element => {
    const p = r.policy;
    const needReason = !p || p.justification;
    const canActivate = r.eligible && !r.viaGroup && !r.active && !r.pending && !(p && p.authContext);
    const reasonOk = !needReason || reason.trim().length >= 5;
    const ticketOk = !p || !p.ticket || (ticketNo.trim().length > 0 && ticketSys.trim().length > 0);
    return (
      <div key={r.key} className={styles.mrCard} data-role={r.name}>
        <div className={styles.mrHead}>
          <div className={styles.mrTitle}>
            <b>{r.name}</b>
            {r.directoryScopeId !== '/' && <span className={styles.muted}> · eingeschränkt auf eine Verwaltungseinheit</span>}
          </div>
          <span className={styles.dlpChips}>
            {r.active && r.permanent && <Pill tone="success">dauerhaft aktiv</Pill>}
            {r.active && !r.permanent && <Pill tone="success">aktiv{r.activeUntil ? ` bis ${formatDate(new Date(r.activeUntil), true)}` : ''}</Pill>}
            {!r.active && r.pending && <Pill tone="warning">Genehmigung ausstehend</Pill>}
            {!r.active && !r.pending && r.eligible && <Pill tone="neutral">nicht aktiv</Pill>}
          </span>
        </div>
        {r.description && <div className={styles.mrDesc}>{r.description}</div>}
        {r.eligible && (
          <div className={styles.dlpChips}>
            {r.viaGroup && <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`}>berechtigt über eine Gruppe</span>}
            {p && <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`}>max. {hoursText(p.maxMinutes)}</span>}
            {p && p.justification && <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`}>Begründung</span>}
            {p && p.ticket && <span className={`${styles.dlpChip} ${styles.dlpChipGeneral}`}>Ticket</span>}
            {p && p.mfa && <span className={`${styles.dlpChip} ${styles.dlpChipBusiness}`}>MFA</span>}
            {p && p.approval && <span className={`${styles.dlpChip} ${styles.dlpChipNew}`}>Genehmigung</span>}
            {p && p.authContext && <span className={`${styles.dlpChip} ${styles.dlpChipBlocked}`}>Authentifizierungskontext</span>}
            {r.eligibleUntil && <span className={styles.muted}>berechtigt bis {formatDate(new Date(r.eligibleUntil))}</span>}
          </div>
        )}
        {notice && notice.key === r.key && <MessageBar messageBarType={notice.type} isMultiline={true} onDismiss={() => setNotice(undefined)}>{notice.text}</MessageBar>}
        {form === r.key && canActivate && (
          <div className={styles.mrForm}>
            <Dropdown label="Dauer" selectedKey={minutes} options={durationOptions(r)} onChange={(_e, o) => o && setMinutes(Number(o.key))} styles={{ root: { maxWidth: 220 } }} />
            <TextField label="Begründung" required={needReason} multiline={true} rows={2} value={reason} onChange={(_e, v) => setReason(v || '')}
              placeholder="z. B. Freigabe von Apps im CoE Cockpit" />
            {p && p.ticket && (
              <div className={styles.setGrid2}>
                <TextField label="Ticketnummer" required={true} value={ticketNo} onChange={(_e, v) => setTicketNo(v || '')} />
                <TextField label="Ticketsystem" required={true} value={ticketSys} onChange={(_e, v) => setTicketSys(v || '')} />
              </div>
            )}
            {p && p.approval && <div className={styles.muted}><Icon iconName="Info" /> Diese Rolle braucht eine Genehmigung – sie wird erst danach aktiv.</div>}
            {p && p.mfa && <div className={styles.muted}><Icon iconName="Info" /> PIM verlangt MFA: klappt, wenn du dich in dieser Sitzung mit MFA angemeldet hast – sonst im PIM-Portal aktivieren.</div>}
            <div className={styles.roleCalloutButtons}>
              <PrimaryButton text={p && p.approval ? 'Anfrage stellen' : 'Aktivieren'} iconProps={{ iconName: 'Unlock' }} disabled={!!busy || !reasonOk || !ticketOk}
                onClick={() => { activate(r).catch(() => setBusy('')); }} />
              <DefaultButton text="Abbrechen" onClick={() => setForm('')} disabled={!!busy} />
            </div>
          </div>
        )}
        {form !== r.key && (canActivate || (r.active && r.selfActivated) || ((r.viaGroup || (p && p.authContext)) && !r.active)) && (
          <div className={styles.roleCalloutButtons}>
            {canActivate && <PrimaryButton text="Aktivieren" iconProps={{ iconName: 'Unlock' }} disabled={!!busy} onClick={() => openForm(r)} />}
            {r.active && r.selfActivated && (
              <DefaultButton text={busy === r.key ? 'Wird beendet …' : 'Vorzeitig beenden'} iconProps={{ iconName: 'Lock' }} disabled={!!busy} onClick={() => { deactivate(r).catch(() => setBusy('')); }} />
            )}
            {(r.viaGroup || (p && p.authContext)) && !r.active && (
              <DefaultButton text="Im PIM-Portal aktivieren" iconProps={{ iconName: 'OpenInNewWindow' }} href={PIM_URL} target="_blank" rel="noopener noreferrer" />
            )}
          </div>
        )}
      </div>
    );
  };

  return (
    <Panel
      isOpen={true}
      onDismiss={props.onDismiss}
      isLightDismiss={true}
      type={PanelType.custom}
      customWidth="560px"
      closeButtonAriaLabel="Schließen"
      headerText="Meine Rollen (PIM)"
    >
      <div style={props.themeVars} className={`${styles.scope} ${styles.mrBody}`}>
        <div className={styles.muted}>
          Die Entra-Rollen, für die <b>du</b> in Privileged Identity Management berechtigt bist (live) – nur diese lassen sich hier aktivieren.
          Microsoft prüft die Regeln der Rolle (Dauer, Begründung, Ticket, MFA, Genehmigung) und protokolliert im Entra-Audit-Log.
        </div>
        <div className={styles.roleCalloutButtons}>
          <DefaultButton text={loading ? 'Lädt …' : 'Aktualisieren'} iconProps={{ iconName: 'Refresh' }} disabled={loading} onClick={load} />
          <DefaultButton text="PIM-Portal" iconProps={{ iconName: 'OpenInNewWindow' }} href={PIM_URL} target="_blank" rel="noopener noreferrer" />
          {data && <span className={styles.muted}>Stand {formatTime(data.loadedAt)} · {eligibleCount} Rolle{eligibleCount === 1 ? '' : 'n'}</span>}
        </div>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
            {error.message}
            {error.kind !== 'consent' && <> <a className={styles.link} href={PIM_URL} target="_blank" rel="noopener noreferrer">Im PIM-Portal aktivieren</a></>}
          </MessageBar>
        )}
        {data && data.policyProblem && <div className={styles.muted}><Icon iconName="Info" /> Regeln der Rollen nicht lesbar ({data.policyProblem}) – Dauer bis 8 Std., Begründung wird verlangt.</div>}
        {data && data.roles.length > 6 && (
          <TextField placeholder="Rolle suchen" value={search} onChange={(_e, v) => setSearch(v || '')} iconProps={{ iconName: 'Search' }} />
        )}
        {loading && !data && <div className={styles.ccLoading}><span className={styles.skeleton} style={{ width: '50%', height: 14 }} /><span className={styles.muted}>Rollen werden gelesen …</span></div>}
        {data && !data.roles.length && <div className={styles.muted}>Du bist in PIM für keine Entra-Rolle berechtigt.</div>}
        {roles.map(card)}
      </div>
    </Panel>
  );
};
