import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IColumnConfig, IRecord, Tone } from '../models/ITableConfig';
import { displayValue, formatDate, hashColor, hostOf, initials, relativeTime, toDate } from '../utils/format';

export function toneClass(tone: Tone | undefined): string {
  switch (tone) {
    case 'success': return styles.toneSuccess;
    case 'warning': return styles.toneWarning;
    case 'danger': return styles.toneDanger;
    case 'info': return styles.toneInfo;
    case 'accent': return styles.toneAccent;
    default: return styles.toneNeutral;
  }
}

export function resolveTone(column: IColumnConfig, record: IRecord): Tone {
  const map = column.statusMap || {};
  const shown = displayValue(record, column.field).toLowerCase();
  const raw = record[column.field];
  const rawText = raw === undefined || raw === null ? '' : String(raw).toLowerCase();
  const keys = Object.keys(map);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i].toLowerCase();
    if (k === shown || k === rawText) {
      return map[keys[i]];
    }
  }
  return 'neutral';
}

export const Pill: React.FC<{ tone?: Tone; title?: string }> = ({ tone, title, children }) => (
  <span className={`${styles.pill} ${toneClass(tone)}`} title={title}>
    <span className={styles.dot} />
    {children}
  </span>
);

export const EmptyValue: React.FC = () => <span className={styles.empty}>—</span>;

export const Avatar: React.FC<{ name: string }> = ({ name }) => (
  <span className={styles.avatar} style={{ background: hashColor(name) }} aria-hidden="true">
    {initials(name)}
  </span>
);

export interface ICellValueProps {
  record: IRecord;
  column: IColumnConfig;
  /** In der Detailansicht: mehr Platz, relative Datumsangabe zusätzlich. */
  expanded?: boolean;
}

export const CellValue: React.FC<ICellValueProps> = ({ record, column, expanded }) => {
  const text = displayValue(record, column.field);
  const raw = record[column.field];

  switch (column.type) {
    case 'title': {
      if (!text) {
        return <EmptyValue />;
      }
      return (
        <span className={styles.titleCell} title={text}>
          <span className={styles.titleAvatar} style={{ background: hashColor(text) }} aria-hidden="true">
            {text.charAt(0).toUpperCase()}
          </span>
          <span className={styles.ellipsis}>{text}</span>
        </span>
      );
    }
    case 'person': {
      if (!text) {
        return <EmptyValue />;
      }
      return (
        <span className={styles.person} title={text}>
          <Avatar name={text} />
          <span className={styles.ellipsis}>{text}</span>
        </span>
      );
    }
    case 'date':
    case 'datetime': {
      const d = toDate(raw);
      if (!d) {
        return <EmptyValue />;
      }
      const withTime = column.type === 'datetime' || !!expanded;
      return (
        <span title={`${formatDate(d, true)} · ${relativeTime(d)}`}>
          {formatDate(d, withTime)}
          {expanded && <span className={styles.muted}> · {relativeTime(d)}</span>}
        </span>
      );
    }
    case 'boolean': {
      if (raw === undefined || raw === null) {
        return <EmptyValue />;
      }
      const yes = raw === true || String(raw).toLowerCase() === 'true';
      return <Pill tone={yes ? 'success' : 'neutral'}>{text || (yes ? 'Ja' : 'Nein')}</Pill>;
    }
    case 'status':
      return text ? <Pill tone={resolveTone(column, record)}>{text}</Pill> : <EmptyValue />;
    case 'choice':
      return text ? <span className={styles.chip} title={text}>{text}</span> : <EmptyValue />;
    case 'url': {
      if (!text) {
        return <EmptyValue />;
      }
      const href = /^https?:\/\//i.test(text) ? text : `https://${text}`;
      return (
        <a className={styles.link} href={href} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()} title={text}>
          {expanded ? text : hostOf(text)}
          <Icon iconName="OpenInNewWindow" style={{ fontSize: 11 }} />
        </a>
      );
    }
    case 'number':
      return text ? <span className={styles.numeric}>{text}</span> : <EmptyValue />;
    case 'list': {
      const items = text ? text.split(', ').filter(s => !!s) : [];
      if (!items.length) {
        return <EmptyValue />;
      }
      if (expanded) {
        return <span>{items.join(', ')}</span>;
      }
      const shown = items.slice(0, 2);
      return (
        <span className={styles.listCell} title={items.join('\n')}>
          <span className={styles.ellipsis}>{shown.join(', ')}</span>
          {items.length > shown.length && <span className={styles.listMore}>+{items.length - shown.length}</span>}
        </span>
      );
    }
    default:
      return text ? <span title={text}>{text}</span> : <EmptyValue />;
  }
};
