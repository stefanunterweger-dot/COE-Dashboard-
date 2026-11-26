import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IColumnConfig, IRecord, ISortConfig } from '../models/ITableConfig';
import { IColumnFilter, isFilterActive } from '../utils/filters';
import { CellValue } from './Cells';
import { ColumnFilterCallout } from './ColumnFilterCallout';

export interface IDataTableProps {
  columns: IColumnConfig[];
  records: IRecord[];
  idField: string;
  loading: boolean;
  compact: boolean;
  sort: ISortConfig | undefined;
  filters: { [field: string]: IColumnFilter };
  themeVars: React.CSSProperties;
  /** Datensätze für die Filterwerte einer Spalte (Suche + alle anderen Filter). */
  getFilterSource: (field: string) => IRecord[];
  onSort: (field: string) => void;
  onSetSort: (field: string, desc: boolean) => void;
  onFilterChange: (field: string, filter: IColumnFilter | undefined) => void;
  onOpen: (record: IRecord) => void;
  /** Optionale Aktion je Zeile (Symbol-Schaltfläche vor dem Pfeil), z. B. „Maker verständigen“. */
  rowAction?: { icon: string; label: string; isDisabled?: (record: IRecord) => boolean; onClick: (record: IRecord) => void };
}

const SKELETON_WIDTHS = [72, 55, 64, 40, 80, 48, 60, 52];

export const DataTable: React.FC<IDataTableProps> = props => {
  const { columns, records, idField, loading, compact, sort, filters, themeVars, getFilterSource, onSort, onSetSort, onFilterChange, onOpen } = props;
  const [openFilter, setOpenFilter] = React.useState<{ field: string; target: HTMLElement } | undefined>(undefined);

  const onRowKey = (e: React.KeyboardEvent<HTMLTableRowElement>, r: IRecord): void => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      onOpen(r);
    }
  };

  const openColumn = openFilter ? columns.filter(c => c.field === openFilter.field)[0] : undefined;

  return (
    <div className={styles.tableWrap}>
      <table className={`${styles.table} ${compact ? styles.compact : ''}`}>
        <thead>
          <tr>
            {props.rowAction && <th className={`${styles.th} ${styles.thAction}`} aria-label={props.rowAction.label} />}
            {columns.map(c => {
              const sorted = !!sort && sort.field === c.field;
              const ariaSort = sorted ? (sort && sort.desc ? 'descending' : 'ascending') : 'none';
              const filtered = isFilterActive(filters[c.field]);
              const isOpen = !!openFilter && openFilter.field === c.field;
              return (
                <th key={c.field} className={styles.th} aria-sort={ariaSort} scope="col">
                  <div className={styles.thInner}>
                    <button type="button" className={`${styles.thButton} ${sorted ? styles.thSorted : ''}`} onClick={() => onSort(c.field)}>
                      <span className={styles.ellipsis}>{c.label}</span>
                      <Icon className={styles.sortIcon} iconName={sorted && sort && sort.desc ? 'SortDown' : 'SortUp'} />
                    </button>
                    <button
                      type="button"
                      className={`${styles.thFilter} ${filtered ? styles.thFilterActive : ''} ${isOpen ? styles.thFilterOpen : ''}`}
                      title={filtered ? `Filter „${c.label}“ aktiv` : `Nach „${c.label}“ filtern`}
                      aria-label={`Filter ${c.label}`}
                      aria-haspopup="dialog"
                      aria-expanded={isOpen}
                      onClick={e => {
                        const el = e.currentTarget as HTMLElement;
                        setOpenFilter(isOpen ? undefined : { field: c.field, target: el });
                      }}
                    >
                      <Icon iconName={filtered ? 'FilterSolid' : 'Filter'} />
                    </button>
                  </div>
                </th>
              );
            })}
            <th className={styles.th} aria-label="Details" />
          </tr>
        </thead>
        <tbody>
          {loading && SKELETON_WIDTHS.map((w, i) => (
            <tr key={`sk${i}`}>
              {props.rowAction && <td className={styles.td} />}
              {columns.map((c, ci) => (
                <td key={c.field} className={styles.td}>
                  <span className={styles.skeleton} style={{ width: `${Math.max(30, (w + ci * 13) % 90)}%` }} />
                </td>
              ))}
              <td className={`${styles.td} ${styles.tdChevron}`} />
            </tr>
          ))}
          {!loading && records.map((r, i) => (
            <tr
              key={String(r[idField] || i)}
              className={styles.row}
              tabIndex={0}
              onClick={() => onOpen(r)}
              onKeyDown={e => onRowKey(e, r)}
            >
              {props.rowAction && (() => {
                const action = props.rowAction;
                const disabled = !!action.isDisabled && action.isDisabled(r);
                return (
                  <td className={`${styles.td} ${styles.tdAction}`}>
                    <button
                      type="button"
                      className={styles.rowActionButton}
                      title={disabled ? `${action.label} – keine E-Mail-Adresse vorhanden` : action.label}
                      aria-label={action.label}
                      disabled={disabled}
                      onClick={e => { e.stopPropagation(); action.onClick(r); }}
                      onKeyDown={e => e.stopPropagation()}
                    >
                      <Icon iconName="Mail" />
                    </button>
                  </td>
                );
              })()}
              {columns.map(c => (
                <td key={c.field} className={`${styles.td} ${c.type === 'number' ? styles.numeric : ''}`}>
                  <CellValue record={r} column={c} />
                </td>
              ))}
              <td className={`${styles.td} ${styles.tdChevron}`}>
                <Icon iconName="ChevronRight" className={styles.chevron} />
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {openFilter && openColumn && (
        <ColumnFilterCallout
          key={openFilter.field}
          column={openColumn}
          target={openFilter.target}
          filter={filters[openFilter.field]}
          records={getFilterSource(openFilter.field)}
          sort={sort}
          themeVars={themeVars}
          onApply={f => { onFilterChange(openFilter.field, f); setOpenFilter(undefined); }}
          onSort={desc => { onSetSort(openFilter.field, desc); setOpenFilter(undefined); }}
          onDismiss={() => setOpenFilter(undefined)}
        />
      )}
    </div>
  );
};
