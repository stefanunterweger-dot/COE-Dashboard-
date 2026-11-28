import * as React from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IColumnConfig, IRecord, ITableConfig } from '../models/ITableConfig';
import { DataverseService, IEntityMeta } from '../services/DataverseService';
import { displayValue, odataLikeLiteral } from '../utils/format';

/**
 * Suche über alle Bereiche (Apps, Flows, Environments, Maker, Connectors …) direkt in Dataverse.
 * Je Tabelle werden die ersten Treffer im Namen gezeigt; ein Klick öffnet die Detailansicht.
 */
export interface IGlobalSearchProps {
  tables: ITableConfig[];
  serviceFor: (t: ITableConfig) => DataverseService | undefined;
  themeVars: React.CSSProperties;
  onOpenRecord: (t: ITableConfig, record: IRecord, meta: IEntityMeta) => void;
  onShowAll: (t: ITableConfig, query: string) => void;
}

/** Große Verknüpfungs- und Protokolltabellen – dort lohnt keine Namenssuche. */
const SKIP = ['admin_auditlog', 'admin_flowactiondetail', 'admin_powerplatformuserrole', 'admin_environmentsecurityrolepermission',
  'admin_connectionreference', 'admin_environmentcapacity', 'admin_connectionreferenceidentity', 'deploymentenvironment'];
const PER_TABLE = 5;
const MAX_TABLES = 16;
/** Weitere Spalten, die (falls konfiguriert) mit durchsucht werden. */
const EXTRA_FIELDS = ['admin_userprincipalname', 'uniquename', 'artifactname'];

interface IHit { table: ITableConfig; records: IRecord[]; meta?: IEntityMeta; status: 'loading' | 'ready' | 'error' }

function titleColumn(t: ITableConfig): IColumnConfig | undefined {
  const plain = t.columns.filter(c => !!c.field && !/^_.+_value$/.test(c.field));
  return plain.filter(c => c.type === 'title')[0] || plain.filter(c => c.type === 'person')[0] || plain[0];
}

function subColumn(t: ITableConfig, title: IColumnConfig): IColumnConfig | undefined {
  return t.columns.filter(c => c.field !== title.field && !c.hidden && c.type !== 'status' && c.type !== 'boolean' && c.type !== 'number')[0];
}

export function searchableTables(tables: ITableConfig[]): ITableConfig[] {
  return tables.filter(t => !!t.logicalName && !t.virtual && !t.unavailable && SKIP.indexOf(t.logicalName) < 0 && !!titleColumn(t)).slice(0, MAX_TABLES);
}

export const GlobalSearch: React.FC<IGlobalSearchProps> = props => {
  const [query, setQuery] = React.useState<string>('');
  const [open, setOpen] = React.useState<boolean>(false);
  const [hits, setHits] = React.useState<IHit[]>([]);
  const boxRef = React.useRef<HTMLDivElement>(null);
  const runRef = React.useRef<number>(0);
  const searchTables = searchableTables(props.tables);
  const tablesSig = searchTables.map(t => `${t.key}|${t.environmentUrl || ''}|${t.filter || ''}`).join(',');

  React.useEffect(() => {
    const q = query.trim();
    if (q.length < 2) {
      setHits([]);
      return;
    }
    const run = ++runRef.current;
    const timer = window.setTimeout(() => {
      const literal = odataLikeLiteral(q);
      setHits(searchTables.map(t => ({ table: t, records: [], status: 'loading' as const })));
      searchTables.forEach(t => {
        const svc = props.serviceFor(t);
        const title = titleColumn(t)!;
        if (!svc) {
          return;
        }
        const fields = [title.field].concat(EXTRA_FIELDS.filter(f => f !== title.field && t.columns.some(c => c.field === f)));
        const match = fields.map(f => `contains(${f},'${literal}')`).join(' or ');
        svc.getRecords({
          ...t, key: `_search_${t.key}`, strictFilter: true,
          filter: t.filter ? `(${t.filter}) and (${match})` : match,
          defaultSort: { field: title.field, desc: false }
        }, PER_TABLE).then(
          res => {
            if (runRef.current === run) {
              setHits(prev => prev.map(h => (h.table.key === t.key ? { ...h, records: res.records, meta: res.meta, status: 'ready' } : h)));
            }
          },
          () => {
            if (runRef.current === run) {
              setHits(prev => prev.map(h => (h.table.key === t.key ? { ...h, status: 'error' } : h)));
            }
          }
        );
      });
    }, 350);
    return () => window.clearTimeout(timer);
  }, [query, tablesSig]);

  const visible = hits.filter(h => h.status === 'loading' || h.records.length > 0);
  const pending = hits.some(h => h.status === 'loading');
  const total = hits.reduce((n, h) => n + h.records.length, 0);
  const q = query.trim();

  return (
    <div ref={boxRef} className={styles.globalSearch}>
      <SearchBox
        placeholder="Alles durchsuchen (Apps, Flows, Maker …)"
        value={query}
        onChange={(_e, v) => { setQuery(v || ''); setOpen(true); }}
        onFocus={() => setOpen(true)}
        onEscape={() => setOpen(false)}
        onClear={() => { setQuery(''); setOpen(false); }}
        styles={{ root: { width: '100%', borderRadius: 6 } }}
        ariaLabel="Alle Bereiche durchsuchen"
      />
      {open && q.length >= 2 && boxRef.current && (
        <Callout
          target={boxRef.current}
          isBeakVisible={false}
          gapSpace={4}
          directionalHint={DirectionalHint.bottomRightEdge}
          onDismiss={() => setOpen(false)}
          setInitialFocus={false}
          calloutMaxHeight={480}
          styles={{ calloutMain: { borderRadius: 8 } }}
        >
          <div className={styles.searchResults} style={props.themeVars} role="listbox" aria-label="Suchergebnisse">
            {visible.map(h => {
              const title = titleColumn(h.table)!;
              const sub = subColumn(h.table, title);
              return (
                <div key={h.table.key} className={styles.searchGroup}>
                  <div className={styles.searchGroupHead}>
                    <Icon iconName={h.table.icon || 'Table'} />
                    <span>{h.table.title}</span>
                    {h.status === 'ready' && h.records.length >= PER_TABLE && (
                      <button type="button" className={styles.textButton} onClick={() => { setOpen(false); props.onShowAll(h.table, q); }}>
                        Alle anzeigen
                      </button>
                    )}
                  </div>
                  {h.status === 'loading' && <span className={styles.skeleton} style={{ display: 'block', height: 14, width: '60%', margin: '6px 12px' }} />}
                  {h.records.map((r, i) => (
                    <button
                      key={String((h.meta && r[h.meta.primaryId]) || i)}
                      type="button"
                      role="option"
                      aria-selected={false}
                      className={styles.searchItem}
                      onClick={() => { if (h.meta) { setOpen(false); props.onOpenRecord(h.table, r, h.meta); } }}
                    >
                      <span className={styles.searchItemTitle}>{displayValue(r, title.field) || '(ohne Namen)'}</span>
                      {sub && <span className={styles.searchItemSub}>{displayValue(r, sub.field)}</span>}
                    </button>
                  ))}
                </div>
              );
            })}
            {!pending && total === 0 && (
              <div className={styles.searchEmpty}><Icon iconName="SearchIssue" /> Keine Treffer für „{q}“</div>
            )}
          </div>
        </Callout>
      )}
    </div>
  );
};
