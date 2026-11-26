import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { DefaultButton } from '@fluentui/react/lib/Button';
import styles from './CoeCockpit.module.scss';
import { IRecord, IRelatedConfig, ITableConfig } from '../models/ITableConfig';
import { DataverseService, IAttributeMeta, IEntityMeta } from '../services/DataverseService';
import { displayValue, fillTemplate, FORMATTED, formatNumber, hashColor, hostOf, isSafeLink } from '../utils/format';
import { CellValue, Pill, resolveTone } from './Cells';

export interface IRecordPanelProps {
  record: IRecord | undefined;
  table: ITableConfig | undefined;
  meta: IEntityMeta | undefined;
  service: DataverseService | undefined;
  themeVars: React.CSSProperties;
  /** ev.preventDefault() hält das Panel offen (z. B. während ein Dialog darüber geöffnet ist). */
  onDismiss: (ev?: React.SyntheticEvent<HTMLElement> | KeyboardEvent) => void;
  /** Gibt es für diese Tabelle eine Konfiguration (dann sind Zeilen in den Reitern anklickbar)? */
  canOpen?: (logicalName: string) => boolean;
  /** Verknüpften Datensatz öffnen (Durchklicken). */
  onOpenRelated?: (logicalName: string, id: string) => void;
  /** Zurück zum vorherigen Datensatz. */
  onBack?: () => void;
  backLabel?: string;
  /** Zusätzliche Aktionen im Kopf der Detailansicht (z. B. „Maker per E-Mail verständigen“, Quarantäne …). */
  actions?: { key: string; text: string; icon: string; disabled?: boolean; title?: string; primary?: boolean; onClick: () => void }[];
  /** Zusätzlicher Inhalt unter den Aktionen (z. B. letzte Aufträge). */
  headerExtra?: React.ReactNode;
  /** Hinweis im Bereich „Aktionen“, wenn keine Aktionen angeboten werden (z. B. fehlende Gruppe). */
  actionsNote?: string;
  /** Abschnitt „Aktionen“ ganz ausblenden (z. B. für Personen außerhalb der berechtigten SharePoint-Gruppe). */
  hideActions?: boolean;
  /** Zusätzlicher Abschnitt in der Übersicht unter „Eigenschaften“ (z. B. „Inhalt der Solution“). */
  extraSection?: { title: string; content: React.ReactNode };
}

interface IRelatedState {
  status: 'loading' | 'ready' | 'error';
  records: IRecord[];
  primaryId?: string;
  warnings: string[];
  truncated?: boolean;
  error?: string;
}

const RELATED_LIMIT = 5000;
/** So viele Zeilen werden in einem Reiter der Detailansicht angezeigt (Rest: Hinweis). */
const RELATED_RENDER = 200;

interface IFullState {
  status: 'idle' | 'loading' | 'ready' | 'error';
  record?: IRecord;
  attributes?: IAttributeMeta[];
  error?: string;
}

export const RecordPanel: React.FC<IRecordPanelProps> = props => {
  const { record, table, meta, service, themeVars, onDismiss } = props;
  const [tab, setTab] = React.useState<string>('overview');
  const [full, setFull] = React.useState<IFullState>({ status: 'idle' });
  const [query, setQuery] = React.useState<string>('');
  const [hideEmpty, setHideEmpty] = React.useState<boolean>(true);
  const [related, setRelated] = React.useState<{ [key: string]: IRelatedState }>({});

  const id = record && meta ? String(record[meta.primaryId] || '') : '';
  /** Datensatz, dessen Felder geladen sind bzw. gerade geladen werden. */
  const fullFor = React.useRef<string>('');

  React.useEffect(() => {
    setTab('overview');
    setFull({ status: 'idle' });
    setQuery('');
    setRelated({});
    fullFor.current = '';
  }, [id]);

  // Verknüpfte Datensätze des gewählten Reiters laden
  const activeRelated: IRelatedConfig | undefined = table && tab.indexOf('rel:') === 0
    ? (table.related || []).filter(r => `rel:${r.key}` === tab)[0]
    : undefined;
  React.useEffect(() => {
    if (!activeRelated || !service || !table || !id || related[activeRelated.key]) {
      return;
    }
    const relKey = activeRelated.key;
    const forId = id;
    setRelated(prev => ({ ...prev, [relKey]: { status: 'loading', records: [], warnings: [] } }));
    service.getRelated(table.logicalName, id, activeRelated, RELATED_LIMIT, record).then(
      res => setRelated(prev => (forId !== idRef.current ? prev : { ...prev, [relKey]: { status: 'ready', records: res.records, primaryId: res.meta.primaryId, warnings: res.warnings, truncated: res.truncated } })),
      (e: Error) => setRelated(prev => (forId !== idRef.current ? prev : { ...prev, [relKey]: { status: 'error', records: [], warnings: [], error: e.message } }))
    );
  }, [tab, id, service, table]);
  const idRef = React.useRef<string>(id);
  idRef.current = id;

  // „Alle Felder“ laden. Der Ladestatus darf hier keine Abhängigkeit sein:
  // sonst bricht das Setzen auf „loading“ die eigene Anfrage sofort wieder ab.
  React.useEffect(() => {
    if (tab !== 'all' || !service || !table || !id) {
      return;
    }
    const key = `${table.logicalName}|${id}`;
    if (fullFor.current === key) {
      return;
    }
    fullFor.current = key;
    let cancelled = false;
    let done = false;
    setFull({ status: 'loading' });
    Promise.all([
      service.getRecord(table.logicalName, id),
      service.getAttributes(table.logicalName).catch(() => [] as IAttributeMeta[])
    ]).then(
      ([rec, attrs]) => {
        done = true;
        if (!cancelled) { setFull({ status: 'ready', record: rec, attributes: attrs }); }
      },
      (e: Error) => {
        done = true;
        if (!cancelled) {
          fullFor.current = ''; // beim nächsten Öffnen von „Alle Felder“ erneut versuchen
          setFull({ status: 'error', error: e.message });
        }
      }
    );
    return () => {
      cancelled = true;
      if (!done) {
        fullFor.current = '';
      }
    };
  }, [tab, service, table, id]);

  if (!record || !table) {
    return null;
  }

  const title = (meta && displayValue(record, meta.primaryName)) ||
    (table.columns.length > 0 ? displayValue(record, table.columns[0].field) : '') || 'Ohne Namen';
  const statusColumns = table.columns.filter(c => c.type === 'status' || c.type === 'boolean');

  const links: { label: string; href: string }[] = [];
  (table.links || []).forEach(l => {
    const href = fillTemplate(l.url, record);
    if (href && isSafeLink(href)) {
      links.push({ label: l.label, href });
    }
  });
  if (service && meta && id) {
    links.push({ label: 'Datensatz in Dataverse öffnen', href: service.recordUrl(meta.logicalName, id) });
  }

  const header = (): JSX.Element => (
    <div style={themeVars} className={styles.scope}>
      <div className={styles.panelHead}>
        <span className={styles.panelAvatar} style={{ background: hashColor(title) }}>{title.charAt(0).toUpperCase()}</span>
        <div style={{ minWidth: 0 }}>
          <div className={styles.panelEyebrow}>
            <Icon iconName={table.icon || 'Table'} />
            <span>{table.title}</span>
          </div>
          <h2 className={styles.panelTitle}>{title}</h2>

          {statusColumns.length > 0 && (
            <div className={styles.panelBadges}>
              {statusColumns.map(c => {
                const v = displayValue(record, c.field);
                if (!v) {
                  return null;
                }
                const tone = c.type === 'status' ? resolveTone(c, record) : (record[c.field] === true ? 'success' : 'neutral');
                return <Pill key={c.field} tone={tone} title={c.label}>{c.type === 'boolean' ? `${c.label}: ${v}` : v}</Pill>;
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );

  const renderOverview = (): JSX.Element => (
    <>
      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Eigenschaften</h3>
        <div className={styles.fieldGrid}>
          {table.columns.map((c, i) => {
            const last = i === table.columns.length - 1;
            return (
              <React.Fragment key={c.field}>
                <div className={`${styles.fieldLabel} ${last ? styles.lastRow : ''}`}>{c.label}</div>
                <div className={`${styles.fieldValue} ${last ? styles.lastRow : ''}`}>
                  <CellValue record={record} column={c} expanded={true} />
                </div>
              </React.Fragment>
            );
          })}
        </div>
      </div>

      {props.extraSection && (
        <div className={styles.section}>
          <h3 className={styles.sectionTitle}>{props.extraSection.title}</h3>
          {props.extraSection.content}
        </div>
      )}

      {links.length > 0 && (
        <div className={styles.section}>
          <h3 className={styles.sectionTitle}>Öffnen in</h3>
          <div className={styles.linkList}>
            {links.map(l => (
              <a key={l.href} className={styles.linkCard} href={l.href} target="_blank" rel="noopener noreferrer">
                <span className={styles.linkIcon}><Icon iconName="OpenInNewWindow" /></span>
                <span className={styles.linkText}>
                  {l.label}
                  <span className={styles.linkSub}>{hostOf(l.href)}</span>
                </span>
                <Icon iconName="ChevronRight" />
              </a>
            ))}
          </div>
        </div>
      )}

      {!(props.hideActions && !(props.actions && props.actions.length)) && <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Aktionen</h3>
        {props.actions && props.actions.length > 0 ? (
          <>
            <div className={styles.panelAction}>
              {props.actions.map(a => (
                <DefaultButton
                  key={a.key}
                  text={a.text}
                  iconProps={{ iconName: a.icon }}
                  disabled={a.disabled}
                  title={a.title}
                  primary={a.primary}
                  onClick={a.onClick}
                />
              ))}
            </div>
            {props.actions.some(a => a.disabled && !!a.title) && (
              <ul className={styles.actionReasons}>
                {props.actions.filter(a => a.disabled && !!a.title).map(a => <li key={a.key}><b>{a.text}:</b> {a.title}</li>)}
              </ul>
            )}
            {props.headerExtra}
          </>
        ) : (
          <div className={styles.infoCard}>
            <Icon iconName="Info" className={styles.infoCardIcon} />
            <div>{props.actionsNote || 'Für diesen Bereich gibt es keine Aktionen.'}</div>
          </div>
        )}
      </div>}
    </>
  );

  const renderRelated = (rel: IRelatedConfig): JSX.Element => {
    const st = related[rel.key];
    if (!st || st.status === 'loading') {
      return <div className={styles.state}><Spinner size={SpinnerSize.medium} label={`Lade ${rel.title} …`} /></div>;
    }
    if (st.status === 'error') {
      return <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{st.error}</MessageBar>;
    }
    const cols = rel.columns.filter(c => !c.hidden);
    return (
      <div className={styles.section}>
        {st.warnings.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true} styles={{ root: { marginBottom: 8 } }}>
            {st.warnings.map(w => <div key={w}>{w}</div>)}
          </MessageBar>
        )}
        <p className={styles.hint}>
          {formatNumber(st.records.length)}{st.truncated ? '+' : ''} {st.records.length === 1 ? 'Eintrag' : 'Einträge'}
          {st.records.length > RELATED_RENDER ? ` – angezeigt werden die ersten ${formatNumber(RELATED_RENDER)}` : ''}
        </p>
        {st.records.length === 0
          ? (
            <div className={styles.state}>
              <div className={styles.stateIcon}><Icon iconName="Inbox" /></div>
              <p className={styles.stateTitle}>Keine Einträge</p>
            </div>
          )
          : (
            <div className={styles.tableWrap}>
              <table className={`${styles.table} ${styles.compact}`}>
                <thead>
                  <tr>{cols.map(c => <th key={c.field} className={styles.th} scope="col"><span className={styles.relatedTh}>{c.label}</span></th>)}</tr>
                </thead>
                <tbody>
                  {st.records.slice(0, RELATED_RENDER).map((r, i) => {
                    const relId = st.primaryId ? String(r[st.primaryId] || '') : '';
                    const clickable = !!relId && !!props.onOpenRelated && !!props.canOpen && props.canOpen(rel.logicalName);
                    const openIt = (): void => { if (clickable) { props.onOpenRelated!(rel.logicalName, relId); } };
                    return (
                      <tr
                        key={i}
                        className={clickable ? styles.relatedRow : undefined}
                        onClick={openIt}
                        onKeyDown={e => { if (clickable && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openIt(); } }}
                        tabIndex={clickable ? 0 : undefined}
                        title={clickable ? 'Datensatz öffnen' : undefined}
                      >
                        {cols.map(c => <td key={c.field}><CellValue record={r} column={c} /></td>)}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
      </div>
    );
  };

  const renderAllFields = (): JSX.Element => {
    if (full.status === 'loading' || full.status === 'idle') {
      return <div className={styles.state}><Spinner size={SpinnerSize.medium} label="Lade alle Felder …" /></div>;
    }
    if (full.status === 'error' || !full.record) {
      return <MessageBar messageBarType={MessageBarType.error}>{full.error || 'Datensatz konnte nicht geladen werden.'}</MessageBar>;
    }
    const rec = full.record;
    const labels: { [key: string]: string } = {};
    (full.attributes || []).forEach(a => { labels[a.selectName] = a.displayName; });
    const q = query.trim().toLowerCase();
    const rows = Object.keys(rec)
      .filter(k => k.indexOf('@') < 0 && k.indexOf('odata') !== 0)
      .map(k => ({ key: k, label: labels[k] || k, value: displayValue(rec, k) }))
      .filter(r => !hideEmpty || r.value !== '')
      .filter(r => !q || r.label.toLowerCase().indexOf(q) >= 0 || r.key.toLowerCase().indexOf(q) >= 0 || r.value.toLowerCase().indexOf(q) >= 0)
      .sort((a, b) => a.label.localeCompare(b.label, 'de'));

    return (
      <div className={styles.section}>
        <div className={styles.panelToolbar}>
          <SearchBox placeholder="Felder durchsuchen" value={query} onChange={(_e, v) => setQuery(v || '')} styles={{ root: { flex: '1 1 auto' } }} />
          <Toggle inlineLabel label="Leere ausblenden" checked={hideEmpty} onChange={(_e, v) => setHideEmpty(!!v)} styles={{ root: { margin: 0 } }} />
        </div>
        <div className={styles.fieldGrid}>
          {rows.map((r, i) => {
            const last = i === rows.length - 1;
            const isFormatted = rec[r.key + FORMATTED] !== undefined;
            return (
              <React.Fragment key={r.key}>
                <div className={`${styles.fieldLabel} ${last ? styles.lastRow : ''}`}>
                  {r.label}
                  <span className={styles.fieldKey}>{r.key}</span>
                </div>
                <div className={`${styles.fieldValue} ${last ? styles.lastRow : ''} ${!isFormatted && r.value.length > 60 ? styles.mono : ''}`}>
                  {r.value || <span className={styles.empty}>—</span>}
                </div>
              </React.Fragment>
            );
          })}
        </div>
      </div>
    );
  };

  return (
    <Panel
      isOpen={!!record}
      onDismiss={onDismiss}
      isLightDismiss={true}
      type={PanelType.custom}
      customWidth={(table.related || []).length > 0 ? '760px' : '600px'}
      closeButtonAriaLabel="Schließen"
      styles={{
        // Kopfbereich steht im Inhalt (über den Reitern) – sonst überdeckt er bei vielen Reitern deren Beschriftung
        content: { paddingTop: 20, paddingLeft: 24, paddingRight: 24, paddingBottom: 32 },
        navigation: { position: 'absolute', right: 8, top: 8, zIndex: 2 }
      }}
    >
      <div style={themeVars} className={`${styles.scope} ${styles.panelBody}`}>
        {props.onBack && (
          <button type="button" className={styles.panelBack} onClick={props.onBack}>
            <Icon iconName="Back" /> Zurück{props.backLabel ? ` zu ${props.backLabel}` : ''}
          </button>
        )}
        <div className={styles.panelHeadWrap}>{header()}</div>
        <Pivot selectedKey={tab} onLinkClick={item => setTab((item && item.props.itemKey) || 'overview')} headersOnly={true} overflowBehavior="menu">
          <PivotItem itemKey="overview" headerText="Übersicht" />
          {(table.related || []).map(r => {
            const st = related[r.key];
            return (
              <PivotItem
                key={r.key}
                itemKey={`rel:${r.key}`}
                headerText={r.title}
                itemCount={st && st.status === 'ready' ? `${st.records.length}${st.truncated ? '+' : ''}` : undefined}
              />
            );
          })}
          <PivotItem itemKey="all" headerText="Alle Felder" />
        </Pivot>
        {tab === 'overview' ? renderOverview() : activeRelated ? renderRelated(activeRelated) : renderAllFields()}
      </div>
    </Panel>
  );
};
