import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { ITableConfig } from '../models/ITableConfig';
import { formatCount } from '../utils/format';
import { toneClass } from './Cells';

export interface IKpiState {
  status: 'loading' | 'ready' | 'error';
  loadKey: string;
  total?: number;
  highlight?: number;
}

export interface IKpiStripProps {
  tables: ITableConfig[];
  kpis: { [key: string]: IKpiState };
  activeKey: string | undefined;
  onSelect: (key: string) => void;
}

export const KpiStrip: React.FC<IKpiStripProps> = ({ tables, kpis, activeKey, onSelect }) => {
  const visible = tables.filter(t => t.showInKpis !== false && !!t.logicalName);
  if (visible.length === 0) {
    return null;
  }
  return (
    <div className={styles.kpis}>
      {visible.map(t => {
        const k = kpis[t.key];
        const loading = !k || k.status === 'loading';
        const hasHighlight = !!(t.kpi && t.kpi.highlightFilter);
        return (
          <button
            key={t.key}
            type="button"
            className={`${styles.kpi} ${activeKey === t.key ? styles.kpiActive : ''}`}
            onClick={() => onSelect(t.key)}
            aria-pressed={activeKey === t.key}
          >
            <div className={styles.kpiTop}>
              <span className={styles.kpiLabel}>{t.title}</span>
              <span className={styles.kpiIcon}><Icon iconName={t.icon || 'Table'} /></span>
            </div>
            {loading
              ? <span className={`${styles.skeleton} ${styles.skeletonValue}`} />
              : <span className={styles.kpiValue}>{k.status === 'error' ? '–' : formatCount(k.total)}</span>}
            <div className={`${styles.kpiFoot} ${toneClass(t.kpi && t.kpi.tone)}`}>
              {!loading && hasHighlight && k.highlight !== undefined && (
                <>
                  <span className={styles.dot} />
                  <span><span className={styles.kpiFootValue}>{formatCount(k.highlight)}</span> {t.kpi && t.kpi.highlightLabel}</span>
                </>
              )}
              {!loading && (!hasHighlight || k.highlight === undefined) && (
                <span>{k.status === 'error' ? 'Nicht verfügbar' : 'Einträge gesamt'}</span>
              )}
            </div>
          </button>
        );
      })}
    </div>
  );
};
