import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { IRecord, ITableConfig } from '../models/ITableConfig';
import { DataverseService, IAttributeMeta, IEntityMeta } from '../services/DataverseService';
import { displayValue, fillTemplate, FORMATTED, hashColor, hostOf } from '../utils/format';
import { CellValue, Pill, resolveTone } from './Cells';

export interface IRecordPanelProps {
  record: IRecord | undefined;
  table: ITableConfig | undefined;
  meta: IEntityMeta | undefined;
  service: DataverseService | undefined;
  themeVars: React.CSSProperties;
  onDismiss: () => void;
}

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

  const id = record && meta ? String(record[meta.primaryId] || '') : '';

  React.useEffect(() => {
    setTab('overview');
    setFull({ status: 'idle' });
    setQuery('');
  }, [id]);

  React.useEffect(() => {
    if (tab !== 'all' || full.status !== 'idle' || !service || !table || !id) {
      return;
    }
    let cancelled = false;
    setFull({ status: 'loading' });
    Promise.all([
      service.getRecord(table.logicalName, id),
      service.getAttributes(table.logicalName).catch(() => [] as IAttributeMeta[])
    ]).then(
      ([rec, attrs]) => { if (!cancelled) { setFull({ status: 'ready', record: rec, attributes: attrs }); } },
      (e: Error) => { if (!cancelled) { setFull({ status: 'error', error: e.message }); } }
    );
    return () => { cancelled = true; };
  }, [tab, full.status, service, table, id]);

  if (!record || !table) {
    return null;
  }

  const title = (meta && displayValue(record, meta.primaryName)) ||
    (table.columns.length > 0 ? displayValue(record, table.columns[0].field) : '') || 'Ohne Namen';
  const statusColumns = table.columns.filter(c => c.type === 'status' || c.type === 'boolean');

  const links: { label: string; href: string }[] = [];
  (table.links || []).forEach(l => {
    const href = fillTemplate(l.url, record);
    if (href) {
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

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Aktionen</h3>
        <div className={styles.infoCard}>
          <Icon iconName="Rocket" className={styles.infoCardIcon} />
          <div>
            Admin-Aktionen wie Besitzer ändern, Flows aktivieren oder deaktivieren, Freigaben verwalten
            und Umgebungseinstellungen anpassen folgen in der nächsten Ausbaustufe direkt an dieser Stelle.
          </div>
        </div>
      </div>
    </>
  );

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
      customWidth="600px"
      closeButtonAriaLabel="Schließen"
      onRenderHeader={header}
      styles={{
        header: { paddingTop: 20, paddingLeft: 24, paddingRight: 48 },
        content: { paddingLeft: 24, paddingRight: 24, paddingBottom: 32 },
        navigation: { position: 'absolute', right: 8, top: 8, zIndex: 2 }
      }}
    >
      <div style={themeVars} className={`${styles.scope} ${styles.panelBody}`}>
        <Pivot selectedKey={tab} onLinkClick={item => setTab((item && item.props.itemKey) || 'overview')} headersOnly={true}>
          <PivotItem itemKey="overview" headerText="Übersicht" />
          <PivotItem itemKey="all" headerText="Alle Felder" />
        </Pivot>
        {tab === 'overview' ? renderOverview() : renderAllFields()}
      </div>
    </Panel>
  );
};
