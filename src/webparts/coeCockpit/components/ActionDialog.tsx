import * as React from 'react';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { ChoiceGroup } from '@fluentui/react/lib/ChoiceGroup';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord } from '../models/ITableConfig';
import { ACTION_HINT, ACTION_LABEL, ActionCode, FINAL_STATUS, IActionItem, resourceOf } from '../models/actions';
import { ActionService } from '../services/ActionService';
import { DataverseService } from '../services/DataverseService';
import { displayValue, isMailAddress } from '../utils/format';

/**
 * Dialog für eine Cockpit-Aktion: Zusammenfassung, Pflicht-Begründung, bei Besitzer/Freigaben die
 * Person bzw. Gruppe. Legt den Auftrag in der SharePoint-Liste an und verfolgt den Status, bis der
 * Flow ihn erledigt hat.
 */
export interface IActionTarget {
  logicalName: string;
  record: IRecord;
  code: ActionCode;
}

export interface IActionDialogProps {
  target: IActionTarget | undefined;
  actions: ActionService;
  dataverse: DataverseService | undefined;
  requestedBy: string;
  themeVars: React.CSSProperties;
  onFinished: (item: IActionItem) => void;
  onDismiss: () => void;
}

/** Abfrage-Takt für den Status (Vorschau setzt kürzere Werte). */
export const ACTION_TIMING = { pollMs: 5000, maxMs: 5 * 60000 };

interface IPrincipal { id: string; name: string; sub: string; isGroup: boolean }

const DANGER: ActionCode[] = ['delete', 'quarantine', 'flowoff', 'envdisable', 'envmanaged'];

export const ActionDialog: React.FC<IActionDialogProps> = props => {
  const { target } = props;
  const res = target ? resourceOf(target.logicalName, target.record) : undefined;
  const [code, setCode] = React.useState<ActionCode>('quarantine');
  const [reason, setReason] = React.useState<string>('');
  const [query, setQuery] = React.useState<string>('');
  const [suggestions, setSuggestions] = React.useState<IPrincipal[]>([]);
  const [picked, setPicked] = React.useState<IPrincipal | undefined>(undefined);
  const [shares, setShares] = React.useState<IPrincipal[] | undefined>(undefined);
  const [phase, setPhase] = React.useState<'edit' | 'sending' | 'waiting' | 'done'>('edit');
  const [item, setItem] = React.useState<IActionItem | undefined>(undefined);
  const [error, setError] = React.useState<string>('');
  const pollRef = React.useRef<number>(0);

  React.useEffect(() => {
    if (!target) {
      return;
    }
    setCode(target.code);
    setReason('');
    setQuery('');
    setSuggestions([]);
    setPicked(undefined);
    setShares(undefined);
    setPhase('edit');
    setItem(undefined);
    setError('');
    pollRef.current++;
  }, [target]);

  const needsPerson = code === 'setowner' || code === 'addviewer' || code === 'addeditor';

  // Vorschläge für Personen/Gruppen aus dem CoE (Power Platform User)
  React.useEffect(() => {
    const q = query.trim();
    if (!needsPerson || !props.dataverse || q.length < 2 || (picked && picked.name === q)) {
      setSuggestions([]);
      return;
    }
    const lit = q.replace(/'/g, "''");
    const timer = window.setTimeout(() => {
      props.dataverse!.getRecords({
        key: '_principals', title: 'Power Platform User', logicalName: 'admin_powerplatformuser', icon: 'People', strictFilter: true,
        filter: `(contains(admin_displayname,'${lit}') or contains(admin_userprincipalname,'${lit}')) and (admin_type eq 'User' or admin_type eq 'Group')`,
        defaultSort: { field: 'admin_displayname', desc: false },
        columns: [
          { field: 'admin_displayname', label: 'Name', type: 'text' },
          { field: 'admin_userprincipalname', label: 'UPN', type: 'text' },
          { field: 'admin_type', label: 'Typ', type: 'text' }
        ]
      }, 8).then(r => setSuggestions(r.records.map(x => ({
        id: String(x[r.meta.primaryId] || ''),
        name: displayValue(x, 'admin_displayname'),
        sub: displayValue(x, 'admin_userprincipalname') || (String(x.admin_type) === 'Group' ? 'Gruppe' : ''),
        isGroup: String(x.admin_type) === 'Group'
      }))), () => setSuggestions([]));
    }, 300);
    return () => window.clearTimeout(timer);
  }, [query, needsPerson]);

  // Bestehende Freigaben (für „Freigabe entfernen“) aus dem CoE
  React.useEffect(() => {
    if (!target || !res || code !== 'removeuser' || !props.dataverse || shares) {
      return;
    }
    const lookup = res.type === 'app' ? '_admin_app_value' : '_admin_flow_value';
    props.dataverse.getRecords({
      key: '_shares', title: 'Freigaben', logicalName: 'admin_powerplatformuserrole', icon: 'People', strictFilter: true,
      filter: `${lookup} eq ${res.coeId} and statecode eq 0`,
      columns: [
        { field: '_admin_powerplatformuser_value', label: 'Person', type: 'text' },
        { field: 'admin_rolename', label: 'Rolle', type: 'text' }
      ]
    }, 500).then(async r => {
      const list = r.records
        .filter(x => !/^owner$/i.test(String(x.admin_rolename || '')) && !!x._admin_powerplatformuser_value)
        .map(x => ({
          id: String(x._admin_powerplatformuser_value).toLowerCase(),
          name: displayValue(x, '_admin_powerplatformuser_value'),
          sub: displayValue(x, 'admin_rolename'),
          isGroup: false
        }));
      // Gruppe oder Person? (Typ steht beim Power Platform User)
      const ids = list.map(x => x.id).filter((v, i, a) => a.indexOf(v) === i).slice(0, 100);
      for (let i = 0; i < ids.length; i += 25) {
        const part = ids.slice(i, i + 25);
        const t = await props.dataverse!.getRecords({
          key: '_principaltypes', title: 'Typ', logicalName: 'admin_powerplatformuser', icon: 'People', strictFilter: true,
          filter: part.map(id => `admin_powerplatformuserid eq ${id}`).join(' or '),
          columns: [{ field: 'admin_type', label: 'Typ', type: 'text' }]
        }, part.length).catch(() => undefined);
        (t ? t.records : []).forEach(x => {
          const pid = String(x.admin_powerplatformuserid || '').toLowerCase();
          list.forEach(s => { if (s.id === pid) { s.isGroup = String(x.admin_type) === 'Group'; } });
        });
      }
      setShares(list);
    }, () => setShares([]));
  }, [target, code]);

  if (!target || !res) {
    return null;
  }

  const typedMail = !picked && isMailAddress(query.trim()) ? query.trim() : '';
  const targetOk = !needsPerson ? (code !== 'removeuser' || !!picked) : (!!picked || !!typedMail);
  const reasonOk = reason.trim().length >= 5;
  const canSend = phase === 'edit' && targetOk && reasonOk;
  const final = !!item && FINAL_STATUS.indexOf(item.status) >= 0;

  const poll = (id: number, run: number, started: number): void => {
    window.setTimeout(() => {
      if (pollRef.current !== run) {
        return;
      }
      props.actions.get(id).then(it => {
        if (pollRef.current !== run || !it) {
          return;
        }
        setItem(it);
        if (FINAL_STATUS.indexOf(it.status) >= 0) {
          setPhase('done');
          props.onFinished(it);
        } else if (Date.now() - started < ACTION_TIMING.maxMs) {
          poll(id, run, started);
        } else {
          setPhase('done');
        }
      }, () => {
        if (pollRef.current === run && Date.now() - started < ACTION_TIMING.maxMs) {
          poll(id, run, started);
        }
      });
    }, ACTION_TIMING.pollMs);
  };

  const send = (): void => {
    setPhase('sending');
    setError('');
    const person = picked || (typedMail ? { id: typedMail, name: typedMail, sub: '', isGroup: false } : undefined);
    props.actions.add({
      action: code, resourceType: res.type, resourceId: res.id, coeRecordId: res.coeId, resourceName: res.name,
      environmentId: res.env, targetUser: person ? person.id : '', targetName: person ? person.name : '', targetIsGroup: person ? person.isGroup : false,
      reason: reason.trim()
    }, props.requestedBy).then(
      id => {
        const run = ++pollRef.current;
        setItem({ id, action: code, resourceType: res.type, resourceId: res.id, coeRecordId: res.coeId, resourceName: res.name, environmentId: res.env, reason, status: 'Angefordert', result: '', created: new Date().toISOString(), completedAt: '', requestedBy: props.requestedBy });
        setPhase('waiting');
        poll(id, run, Date.now());
      },
      (e: unknown) => { setPhase('edit'); setError(e instanceof Error ? e.message : String(e)); }
    );
  };

  const statusText = item ? ({
    'Angefordert': 'Auftrag angelegt – wartet auf den Flow (meist unter einer Minute) …',
    'In Arbeit': 'Der Flow führt den Auftrag gerade aus …',
    'Erledigt': 'Erledigt.',
    'Fehler': 'Fehler beim Ausführen.',
    'Abgelehnt': 'Abgelehnt.'
  } as { [s: string]: string })[item.status] : '';

  const shareOptions: IDropdownOption[] = (shares || []).map(s => ({ key: s.id, text: `${s.name} – ${s.sub}` }));

  return (
    <Dialog
      hidden={false}
      onDismiss={props.onDismiss}
      minWidth={560}
      maxWidth={640}
      dialogContentProps={{ type: DialogType.normal, title: `${ACTION_LABEL[code]}`, subText: `${res.type === 'app' ? 'App' : res.type === 'flow' ? 'Flow' : 'Environment'} „${res.name}“` }}
      modalProps={{ isBlocking: phase === 'sending' || phase === 'waiting' }}
    >
      <div className={styles.notify} style={props.themeVars}>
        <MessageBar messageBarType={DANGER.indexOf(code) >= 0 ? MessageBarType.severeWarning : MessageBarType.info} isMultiline={true}>
          {ACTION_HINT[code]}
        </MessageBar>

        {phase === 'edit' && (code === 'addviewer' || code === 'addeditor') && res.type === 'app' && (
          <ChoiceGroup
            label="Rolle"
            selectedKey={code}
            options={[{ key: 'addviewer', text: 'Benutzer (darf die App verwenden)' }, { key: 'addeditor', text: 'Mitbesitzer (darf bearbeiten und freigeben)' }]}
            onChange={(_e, o) => o && setCode(o.key as ActionCode)}
          />
        )}

        {phase === 'edit' && needsPerson && (
          <div>
            <TextField
              label={code === 'setowner' ? 'Neuer Besitzer' : 'Person oder Gruppe'}
              placeholder="Name oder E-Mail-Adresse"
              value={picked ? picked.name : query}
              onChange={(_e, v) => { setPicked(undefined); setQuery(v || ''); }}
              description={typedMail ? 'Nicht im CoE – der Flow sucht die Person über die E-Mail-Adresse.' : 'Vorschläge aus dem CoE (Power Platform User). Personen außerhalb des CoE per E-Mail-Adresse.'}
            />
            {suggestions.length > 0 && !picked && (
              <div className={styles.actionSuggest} role="listbox" aria-label="Vorschläge">
                {suggestions.filter(s => code !== 'setowner' || !s.isGroup).map(s => (
                  <button key={s.id} type="button" role="option" aria-selected={false} className={styles.searchItem} onClick={() => { setPicked(s); setSuggestions([]); }}>
                    <span className={styles.searchItemTitle}><Icon iconName={s.isGroup ? 'Group' : 'Contact'} /> {s.name}</span>
                    {s.sub && <span className={styles.searchItemSub}>{s.sub}</span>}
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

        {phase === 'edit' && code === 'removeuser' && (
          <Dropdown
            label="Freigabe"
            placeholder={shares === undefined ? 'Freigaben werden geladen …' : shareOptions.length ? 'Person oder Gruppe wählen' : 'Keine entfernbaren Freigaben im CoE'}
            options={shareOptions}
            selectedKey={picked ? picked.id : undefined}
            disabled={!shareOptions.length}
            onChange={(_e, o) => setPicked((shares || []).filter(s => s.id === (o && o.key))[0])}
          />
        )}

        {phase === 'edit' && (
          <TextField
            label="Begründung"
            required={true}
            multiline={true}
            rows={3}
            value={reason}
            onChange={(_e, v) => setReason(v || '')}
            placeholder="Warum? (wird im Aktionsprotokoll gespeichert)"
          />
        )}

        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>Auftrag konnte nicht angelegt werden: {error}</MessageBar>}

        {item && (
          <div className={styles.actionStatus}>
            {!final && <ProgressIndicator label={`Auftrag #${item.id}`} description={statusText} />}
            {final && (
              <MessageBar messageBarType={item.status === 'Erledigt' ? MessageBarType.success : MessageBarType.error} isMultiline={true}>
                <b>Auftrag #{item.id}: {item.status}</b>{item.result ? ` – ${item.result}` : ''}
              </MessageBar>
            )}
            {phase === 'done' && !final && (
              <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
                Der Flow hat den Auftrag noch nicht übernommen. Er bleibt in der Liste und wird ausgeführt, sobald der Flow läuft – den Stand siehst du im Aktionsprotokoll.
              </MessageBar>
            )}
          </div>
        )}
      </div>
      <DialogFooter>
        {phase === 'edit' && (
          <PrimaryButton text="Auftrag senden" iconProps={{ iconName: 'Send' }} disabled={!canSend} onClick={send}
            title={!reasonOk ? 'Bitte eine Begründung angeben (mind. 5 Zeichen).' : !targetOk ? 'Bitte Person bzw. Freigabe wählen.' : undefined} />
        )}
        <DefaultButton text={phase === 'edit' ? 'Abbrechen' : 'Schließen'} disabled={phase === 'sending'} onClick={() => { pollRef.current++; props.onDismiss(); }} />
      </DialogFooter>
    </Dialog>
  );
};
