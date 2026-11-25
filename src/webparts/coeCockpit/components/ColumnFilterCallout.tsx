import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IColumnConfig, IRecord, ISortConfig } from '../models/ITableConfig';
import { formatNumber } from '../utils/format';
import { buildDateCounts, buildOptions, DATE_PRESETS, IColumnFilter, isDateColumn } from '../utils/filters';
import { Avatar, Pill, resolveTone } from './Cells';

export interface IColumnFilterCalloutProps {
  column: IColumnConfig;
  target: HTMLElement;
  filter: IColumnFilter | undefined;
  /** Datensätze nach Suche und allen anderen Spaltenfiltern (für Werte und Anzahl). */
  records: IRecord[];
  sort: ISortConfig | undefined;
  themeVars: React.CSSProperties;
  onApply: (filter: IColumnFilter | undefined) => void;
  onSort: (desc: boolean) => void;
  onDismiss: () => void;
}

const MAX_OPTIONS = 300;

export const ColumnFilterCallout: React.FC<IColumnFilterCalloutProps> = props => {
  const { column, target, filter, records, sort, themeVars, onApply, onSort, onDismiss } = props;
  const isDate = isDateColumn(column);

  const options = React.useMemo(() => (isDate ? [] : buildOptions(records, column)), [records, column.field, isDate]);
  const dateCounts = React.useMemo(() => (isDate ? buildDateCounts(records, column) : {}), [records, column.field, isDate]);

  const [query, setQuery] = React.useState<string>('');
  const [checked, setChecked] = React.useState<{ [value: string]: boolean }>(() => {
    const init: { [value: string]: boolean } = {};
    if (filter && filter.values && filter.values.length > 0) {
      filter.values.forEach(v => { init[v] = true; });
      // Gefilterte Werte, die aktuell nicht mehr vorkommen, bleiben erhalten.
    } else {
      options.forEach(o => { init[o.value] = true; });
    }
    return init;
  });

  const q = query.trim().toLowerCase();
  const visible = q ? options.filter(o => o.label.toLowerCase().indexOf(q) >= 0) : options;
  const shown = visible.slice(0, MAX_OPTIONS);
  const allVisibleChecked = visible.length > 0 && visible.every(o => !!checked[o.value]);
  const someVisibleChecked = visible.some(o => !!checked[o.value]);
  const checkedCount = options.filter(o => !!checked[o.value]).length;

  const toggleAllVisible = (on: boolean): void => {
    setChecked(prev => {
      const next = { ...prev };
      visible.forEach(o => { next[o.value] = on; });
      return next;
    });
  };

  const apply = (): void => {
    // Bei aktiver Suche: nur die Treffer übernehmen (wie in Excel).
    const selected = (q ? visible : options).filter(o => !!checked[o.value]).map(o => o.value);
    if (!q && selected.length === options.length) {
      onApply(undefined);
    } else {
      onApply({ values: selected });
    }
  };

  const sortedAsc = !!sort && sort.field === column.field && !sort.desc;
  const sortedDesc = !!sort && sort.field === column.field && sort.desc;
  const ascLabel = isDate ? 'Älteste zuerst' : column.type === 'number' ? 'Kleinste zuerst' : 'A bis Z';
  const descLabel = isDate ? 'Neueste zuerst' : column.type === 'number' ? 'Größte zuerst' : 'Z bis A';

  const renderOptionLabel = (value: string, label: string): JSX.Element => {
    if (value !== '' && column.type === 'status') {
      return <Pill tone={resolveTone(column, { [column.field]: value })}>{label}</Pill>;
    }
    if (value !== '' && (column.type === 'person' || column.type === 'title')) {
      return <span className={styles.person}><Avatar name={label} /><span className={styles.ellipsis}>{label}</span></span>;
    }
    return <span className={`${styles.ellipsis} ${value === '' ? styles.muted : ''}`}>{label}</span>;
  };

  return (
    <Callout
      target={target}
      onDismiss={onDismiss}
      isBeakVisible={false}
      gapSpace={4}
      directionalHint={DirectionalHint.bottomLeftEdge}
      setInitialFocus={true}
      role="dialog"
      ariaLabel={`Filter ${column.label}`}
      styles={{ calloutMain: { borderRadius: 10 }, root: { borderRadius: 10, boxShadow: '0 8px 28px rgba(0,0,0,.16)' } }}
    >
      <div style={themeVars} className={`${styles.scope} ${styles.filterCallout}`}>
        <div className={styles.filterHead}>
          <span className={styles.filterTitle}>{column.label}</span>
        </div>

        <div className={styles.filterSort}>
          <button type="button" className={`${styles.filterSortButton} ${sortedAsc ? styles.filterSortActive : ''}`} onClick={() => onSort(false)}>
            <Icon iconName="SortUp" /> {ascLabel}
          </button>
          <button type="button" className={`${styles.filterSortButton} ${sortedDesc ? styles.filterSortActive : ''}`} onClick={() => onSort(true)}>
            <Icon iconName="SortDown" /> {descLabel}
          </button>
        </div>

        {isDate && (
          <div className={styles.filterList}>
            <button type="button" className={`${styles.presetItem} ${!filter || !filter.date ? styles.presetActive : ''}`} onClick={() => onApply(undefined)}>
              <span>Alle</span>
              <span className={styles.filterCount}>{formatNumber(records.length)}</span>
            </button>
            {DATE_PRESETS.map(p => (
              <button
                key={p.key}
                type="button"
                className={`${styles.presetItem} ${filter && filter.date === p.key ? styles.presetActive : ''}`}
                onClick={() => onApply({ date: p.key })}
                disabled={!dateCounts[p.key] && !(filter && filter.date === p.key)}
              >
                <span>{p.label}</span>
                <span className={styles.filterCount}>{formatNumber(dateCounts[p.key] || 0)}</span>
              </button>
            ))}
          </div>
        )}

        {!isDate && (
          <>
            <div className={styles.filterSearch}>
              <SearchBox placeholder="Werte suchen" value={query} onChange={(_e, v) => setQuery(v || '')} underlined={false} styles={{ root: { borderRadius: 6 } }} />
            </div>
            <div className={styles.filterList}>
              {visible.length > 0 && (
                <Checkbox
                  className={styles.filterOption}
                  label={q ? 'Alle Treffer auswählen' : 'Alle auswählen'}
                  checked={allVisibleChecked}
                  indeterminate={!allVisibleChecked && someVisibleChecked}
                  onChange={() => toggleAllVisible(!allVisibleChecked)}
                  styles={{ text: { fontWeight: 600 } }}
                />
              )}
              {shown.map(o => (
                <Checkbox
                  key={o.value || '__empty'}
                  className={styles.filterOption}
                  checked={!!checked[o.value]}
                  onChange={(_e, v) => setChecked(prev => ({ ...prev, [o.value]: !!v }))}
                  ariaLabel={o.label}
                  onRenderLabel={() => (
                    <span className={styles.filterOptionLabel}>
                      {renderOptionLabel(o.value, o.label)}
                      <span className={styles.filterCount}>{formatNumber(o.count)}</span>
                    </span>
                  )}
                  styles={{ label: { width: '100%', alignItems: 'center' }, root: { width: '100%' } }}
                />
              ))}
              {visible.length === 0 && <div className={styles.filterEmpty}>Keine passenden Werte</div>}
              {visible.length > MAX_OPTIONS && (
                <div className={styles.filterEmpty}>
                  {formatNumber(visible.length - MAX_OPTIONS)} weitere Werte – bitte die Suche verwenden.
                </div>
              )}
            </div>
            <div className={styles.filterFooter}>
              <span className={styles.filterSummary}>
                {formatNumber(checkedCount)} von {formatNumber(options.length)} ausgewählt
              </span>
              <span className={styles.grow} />
              <DefaultButton text="Zurücksetzen" onClick={() => onApply(undefined)} styles={{ root: { borderRadius: 6, minWidth: 0 } }} />
              <PrimaryButton
                text="Anwenden"
                onClick={apply}
                disabled={(q ? visible : options).filter(o => !!checked[o.value]).length === 0}
                styles={{ root: { borderRadius: 6, minWidth: 0 } }}
              />
            </div>
          </>
        )}
      </div>
    </Callout>
  );
};
