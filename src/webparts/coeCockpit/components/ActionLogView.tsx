import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { ACTION_ICON, ACTION_LABEL, IActionItem } from '../models/actions';
import { Pill } from './Cells';
import { Tone } from '../models/ITableConfig';

/** Aktionsprotokoll: alle Aufträge aus der Liste „CoE Cockpit Aktionen“, neueste zuerst. */
export interface IActionLogViewProps {
  items: IActionItem[] | undefined;
  loading: boolean;
  problem: string;
  onRefresh: () => void;
}

const TONE: { [s: string]: Tone } = { 'Angefordert': 'info', 'In Arbeit': 'info', 'Erledigt': 'success', 'Fehler': 'danger', 'Abgelehnt': 'warning' };
const dt = new Intl.DateTimeFormat('de-AT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
const fmt = (s: string): string => { const d = new Date(s); return s && !isNaN(d.getTime()) ? dt.format(d) : '—'; };

type Filter = 'all' | 'open' | 'done' | 'failed';

export const ActionLogView: React.FC<IActionLogViewProps> = props => {
  const [filter, setFilter] = React.useState<Filter>('all');
  const items = props.items || [];
  const count = (f: Filter): number => items.filter(i => match(i, f)).length;
  const shown = items.filter(i => match(i, filter));
  return (
    <div className={styles.actionLog}>
      <div className={styles.toolbar}>
        {(['all', 'open', 'done', 'failed'] as Filter[]).map(f => (
          <button key={f} type="button" aria-pressed={filter === f} className={`${styles.filterChip} ${filter === f ? styles.actionFilterActive : ''}`} onClick={() => setFilter(f)}>
            {{ all: 'Alle', open: 'Offen', done: 'Erledigt', failed: 'Fehler / abgelehnt' }[f]} ({count(f)})
          </button>
        ))}
        <span className={styles.grow} />
        <DefaultButton text="Aktualisieren" iconProps={{ iconName: 'Refresh' }} onClick={props.onRefresh} disabled={props.loading} />
      </div>
      {props.problem && <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>{props.problem}</MessageBar>}
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              {['Zeitpunkt', 'Aktion', 'Ressource', 'Person / Gruppe', 'Status', 'Ergebnis', 'Angefordert von', 'Begründung'].map(h => <th key={h} className={styles.th}><div className={styles.thInner}><span className={styles.thButton}>{h}</span></div></th>)}
            </tr>
          </thead>
          <tbody>
            {props.loading && !items.length && <tr><td className={styles.td} colSpan={8}><span className={styles.skeleton} style={{ width: '60%' }} /></td></tr>}
            {shown.map(i => (
              <tr key={i.id} className={styles.row}>
                <td className={styles.td}>{fmt(i.created)}</td>
                <td className={styles.td}><Icon iconName={ACTION_ICON[i.action] || 'Settings'} /> {ACTION_LABEL[i.action] || i.action}</td>
                <td className={styles.td}><b>{i.resourceName}</b><div className={styles.metaText}>{i.resourceType === 'flow' ? 'Flow' : i.resourceType === 'environment' ? 'Environment' : 'App'}</div></td>
                <td className={styles.td}>{i.targetName || i.targetUser || '—'}</td>
                <td className={styles.td}><Pill tone={TONE[i.status] || 'neutral'}>{i.status}</Pill></td>
                <td className={styles.td}>{i.result || (i.completedAt ? fmt(i.completedAt) : '—')}</td>
                <td className={styles.td}>{i.requestedBy || '—'}</td>
                <td className={styles.td}>{i.reason || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!props.loading && shown.length === 0 && (
        <div className={styles.state}>
          <div className={styles.stateIcon}><Icon iconName="Inbox" /></div>
          <p className={styles.stateTitle}>Keine Aufträge</p>
          <p className={styles.stateText}>Aufträge entstehen über die Aktionen in der Detailansicht einer App oder eines Flows.</p>
        </div>
      )}
    </div>
  );
};

function match(i: IActionItem, f: Filter): boolean {
  switch (f) {
    case 'open': return i.status === 'Angefordert' || i.status === 'In Arbeit';
    case 'done': return i.status === 'Erledigt';
    case 'failed': return i.status === 'Fehler' || i.status === 'Abgelehnt';
    default: return true;
  }
}
