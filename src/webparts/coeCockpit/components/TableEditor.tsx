import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { PrimaryButton, DefaultButton, IconButton, ActionButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CoeCockpit.module.scss';
import { ColumnType, IColumnConfig, ITableConfig, Tone } from '../models/ITableConfig';
import { getDefaultTables } from '../models/defaultTables';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { DataverseService, IAttributeMeta, IEntityMeta, ITableInfo } from '../services/DataverseService';
import { getTablePresets, newTableConfig } from '../models/tablePresets';

export interface ITableEditorProps {
  isOpen: boolean;
  tables: ITableConfig[];
  service: DataverseService | undefined;
  themeVars: React.CSSProperties;
  onSave: (tables: ITableConfig[]) => void;
  onDismiss: () => void;
}

interface ICatalogState {
  status: 'idle' | 'loading' | 'ready' | 'error';
  tables?: ITableInfo[];
  error?: string;
}

interface IPickerItem {
  logicalName: string;
  title: string;
  subtitle: string;
  icon: string;
  preset?: ITableConfig;
  used: boolean;
}

interface IResolveState {
  status: 'idle' | 'loading' | 'ready' | 'error';
  meta?: IEntityMeta;
  attributes?: IAttributeMeta[];
  error?: string;
}

const TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'title', text: 'Titel (mit Symbol)' },
  { key: 'text', text: 'Text' },
  { key: 'person', text: 'Person' },
  { key: 'status', text: 'Status (farbig)' },
  { key: 'choice', text: 'Auswahl' },
  { key: 'boolean', text: 'Ja/Nein' },
  { key: 'date', text: 'Datum' },
  { key: 'datetime', text: 'Datum & Uhrzeit' },
  { key: 'number', text: 'Zahl' },
  { key: 'url', text: 'Link' },
  { key: 'list', text: 'Liste (mehrere Werte, „, “ getrennt)' }
];

const TONE_OPTIONS: IDropdownOption[] = [
  { key: 'accent', text: 'Akzent' },
  { key: 'success', text: 'Grün' },
  { key: 'warning', text: 'Gelb' },
  { key: 'danger', text: 'Rot' },
  { key: 'info', text: 'Blau' },
  { key: 'neutral', text: 'Grau' }
];

const TONES = ['accent', 'success', 'warning', 'danger', 'info', 'neutral'];

function clone<T>(v: T): T {
  return JSON.parse(JSON.stringify(v)) as T;
}

function mapAttributeType(a: IAttributeMeta): ColumnType {
  if (a.isPrimaryName) {
    return 'title';
  }
  switch (a.type) {
    case 'DateTime': return 'date';
    case 'Boolean': return 'boolean';
    case 'Picklist':
    case 'State':
    case 'Status': return 'choice';
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money': return 'number';
    case 'Owner': return 'person';
    default: return 'text';
  }
}

function statusMapToText(map?: { [k: string]: Tone }): string {
  return map ? Object.keys(map).map(k => `${k}=${map[k]}`).join('; ') : '';
}

function textToStatusMap(text: string): { [k: string]: Tone } {
  const out: { [k: string]: Tone } = {};
  text.split(/[;\n]/).forEach(part => {
    const idx = part.lastIndexOf('=');
    if (idx > 0) {
      const k = part.substring(0, idx).trim();
      const v = part.substring(idx + 1).trim().toLowerCase();
      if (k && TONES.indexOf(v) >= 0) {
        out[k] = v as Tone;
      }
    }
  });
  return out;
}

function linksToText(t: ITableConfig): string {
  return (t.links || []).map(l => `${l.label} | ${l.url}`).join('\n');
}

function textToLinks(text: string): { label: string; url: string }[] {
  return text.split('\n').map(line => {
    const idx = line.indexOf('|');
    if (idx < 0) {
      return undefined;
    }
    const label = line.substring(0, idx).trim();
    const url = line.substring(idx + 1).trim();
    return label && url ? { label, url } : undefined;
  }).filter(l => !!l) as { label: string; url: string }[];
}

export const TableEditor: React.FC<ITableEditorProps> = props => {
  const { isOpen, tables, service, themeVars, onSave, onDismiss } = props;
  const [draft, setDraft] = React.useState<ITableConfig[]>([]);
  const [index, setIndex] = React.useState<number>(0);
  const [mode, setMode] = React.useState<string>('form');
  const [jsonText, setJsonText] = React.useState<string>('');
  const [jsonError, setJsonError] = React.useState<string>('');
  const [resolve, setResolve] = React.useState<{ [logical: string]: IResolveState }>({});
  const [linkText, setLinkText] = React.useState<string>('');
  const [catalog, setCatalog] = React.useState<ICatalogState>({ status: 'idle' });
  /** Tabellenauswahl: 'add' = neue Tabelle, 'change' = Tabelle des aktuellen Eintrags ändern. */
  const [picker, setPicker] = React.useState<'add' | 'change' | undefined>(undefined);
  const [pickerQuery, setPickerQuery] = React.useState<string>('');

  React.useEffect(() => {
    if (isOpen) {
      setDraft(clone(tables));
      setIndex(0);
      setMode('form');
      setJsonError('');
      setPicker(tables.length === 0 ? 'add' : undefined);
      setPickerQuery('');
    }
  }, [isOpen]);

  // Tabellenliste der Umgebung für die Auswahl laden
  React.useEffect(() => {
    if (!isOpen || !service || catalog.status === 'loading' || catalog.status === 'ready') {
      return;
    }
    setCatalog({ status: 'loading' });
    service.getTables().then(
      list => setCatalog({ status: 'ready', tables: list }),
      (e: Error) => setCatalog({ status: 'error', error: e.message })
    );
  }, [isOpen, service]);

  const current: ITableConfig | undefined = draft[index];

  React.useEffect(() => {
    setLinkText(current ? linksToText(current) : '');
  }, [index, current ? current.key + JSON.stringify(current.links || []) : '']);

  const update = (patch: Partial<ITableConfig>): void => {
    setDraft(prev => prev.map((t, i) => (i === index ? { ...t, ...patch } : t)));
  };

  const updateColumn = (ci: number, patch: Partial<IColumnConfig>): void => {
    if (!current) {
      return;
    }
    const columns = current.columns.map((c, i) => (i === ci ? { ...c, ...patch } : c));
    update({ columns });
  };

  const moveColumn = (ci: number, dir: number): void => {
    if (!current) {
      return;
    }
    const target = ci + dir;
    if (target < 0 || target >= current.columns.length) {
      return;
    }
    const columns = current.columns.slice();
    const tmp = columns[ci];
    columns[ci] = columns[target];
    columns[target] = tmp;
    update({ columns });
  };

  const removeColumn = (ci: number): void => {
    if (current) {
      update({ columns: current.columns.filter((_c, i) => i !== ci) });
    }
  };

  const addColumn = (): void => {
    if (current) {
      update({ columns: current.columns.concat([{ field: '', label: '', type: 'text' }]) });
    }
  };

  /** Auswahl einer Tabelle (Vorlage oder beliebige Dataverse-Tabelle) übernehmen. */
  const pickTable = (item: IPickerItem): void => {
    const base: ITableConfig = item.preset ? clone(item.preset) : newTableConfig(item.logicalName, item.title);
    if (picker === 'change' && current) {
      const changed: ITableConfig = {
        ...base,
        key: current.key,
        showInKpis: current.showInKpis
      };
      setDraft(prev => prev.map((t, i) => (i === index ? changed : t)));
      resolveFor(index, item.logicalName, true);
    } else {
      const usedKeys = draft.map(t => t.key);
      let key = base.key;
      while (usedKeys.indexOf(key) >= 0) {
        key = `${key}_`;
      }
      const target = draft.length;
      setDraft(prev => prev.concat([{ ...base, key }]));
      setIndex(target);
      resolveFor(target, item.logicalName, true);
    }
    setPicker(undefined);
    setPickerQuery('');
  };

  const removeTable = (): void => {
    setDraft(prev => prev.filter((_t, i) => i !== index));
    setIndex(Math.max(0, index - 1));
  };

  const moveTable = (dir: number): void => {
    const target = index + dir;
    if (target < 0 || target >= draft.length) {
      return;
    }
    setDraft(prev => {
      const next = prev.slice();
      const tmp = next[index];
      next[index] = next[target];
      next[target] = tmp;
      return next;
    });
    setIndex(target);
  };

  /**
   * Lädt die Metadaten einer Tabelle (für die Spaltenauswahl).
   * prune = Spalten entfernen, die es in dieser Umgebung nicht gibt (beim Hinzufügen einer Vorlage).
   */
  const resolveFor = (targetIndex: number, logicalName: string, prune: boolean): void => {
    if (!service || !logicalName.trim()) {
      return;
    }
    const logical = logicalName.trim().toLowerCase();
    setResolve(prev => ({ ...prev, [logical]: { status: 'loading' } }));
    Promise.all([service.getEntityMeta(logical), service.getAttributes(logical)]).then(
      ([meta, attributes]) => {
        setResolve(prev => ({ ...prev, [logical]: { status: 'ready', meta, attributes } }));
        setDraft(prev => prev.map((t, i) => {
          if (i !== targetIndex || t.logicalName.trim().toLowerCase() !== logical) {
            return t;
          }
          const patch: Partial<ITableConfig> = { logicalName: logical };
          if (prune && t.columns.length > 0) {
            const known: { [f: string]: boolean } = {};
            attributes.forEach(a => { known[a.selectName] = true; });
            patch.columns = t.columns.filter(c => known[c.field]);
            if (t.defaultSort && !known[t.defaultSort.field]) {
              patch.defaultSort = undefined;
            }
          }
          if (!t.title || t.title === 'Neue Tabelle') {
            patch.title = meta.displayName;
          }
          if ((patch.columns || t.columns).length === 0) {
            const primary = attributes.filter(a => a.isPrimaryName)[0];
            const created = attributes.filter(a => a.logicalName === 'createdon')[0];
            const modified = attributes.filter(a => a.logicalName === 'modifiedon')[0];
            const cols: IColumnConfig[] = [];
            if (primary) { cols.push({ field: primary.selectName, label: primary.displayName, type: 'title' }); }
            if (modified) { cols.push({ field: modified.selectName, label: 'Geändert', type: 'date' }); }
            if (created) { cols.push({ field: created.selectName, label: 'Erstellt', type: 'date' }); }
            patch.columns = cols;
          }
          return { ...t, ...patch };
        }));
      },
      (e: Error) => setResolve(prev => ({ ...prev, [logical]: { status: 'error', error: e.message } }))
    );
  };

  const save = (): void => {
    const used: { [k: string]: boolean } = {};
    const cleaned = draft
      .filter(t => t.logicalName.trim().length > 0)
      .map(t => {
        let key = (t.key || t.logicalName).replace(/[^A-Za-z0-9_-]/g, '') || 't';
        while (used[key]) {
          key = `${key}_`;
        }
        used[key] = true;
        return {
          ...t,
          key,
          logicalName: t.logicalName.trim().toLowerCase(),
          columns: t.columns.filter(c => c.field.trim().length > 0).map(c => ({ ...c, field: c.field.trim(), label: c.label || c.field }))
        };
      });
    onSave(cleaned);
  };

  const applyJson = (): void => {
    try {
      const parsed = JSON.parse(jsonText) as ITableConfig[];
      if (!Array.isArray(parsed) || parsed.some(t => !t || typeof t.logicalName !== 'string' || !Array.isArray(t.columns))) {
        throw new Error('Erwartet wird ein Array von Tabellen mit "logicalName" und "columns".');
      }
      setDraft(parsed);
      setIndex(0);
      setJsonError('');
      setMode('form');
    } catch (e) {
      setJsonError((e as Error).message);
    }
  };

  const resolveState: IResolveState | undefined = current ? resolve[current.logicalName.trim().toLowerCase()] : undefined;
  const attributeOptions: IComboBoxOption[] = resolveState && resolveState.attributes
    ? resolveState.attributes.map(a => ({ key: a.selectName, text: `${a.displayName} (${a.selectName})` }))
    : [];

  const onPickAttribute = (ci: number, key: string | undefined, freeText: string | undefined): void => {
    const field = (key || freeText || '').trim();
    const attr = resolveState && resolveState.attributes ? resolveState.attributes.filter(a => a.selectName === field)[0] : undefined;
    const col = current ? current.columns[ci] : undefined;
    const patch: Partial<IColumnConfig> = { field };
    if (attr) {
      if (!col || !col.label) {
        patch.label = attr.displayName;
      }
      patch.type = mapAttributeType(attr);
    }
    updateColumn(ci, patch);
  };

  // Metadaten der gewählten Tabelle automatisch laden (für die Spaltenauswahl)
  React.useEffect(() => {
    if (!isOpen || !current || !service) {
      return;
    }
    const logical = current.logicalName.trim().toLowerCase();
    if (logical && !resolve[logical]) {
      resolveFor(index, logical, false);
    }
  }, [isOpen, index, current ? current.logicalName : '', service]);

  const buildPickerItems = (): { presets: IPickerItem[]; others: IPickerItem[]; total: number } => {
    const used: { [l: string]: boolean } = {};
    draft.forEach((t, i) => {
      if (!(picker === 'change' && i === index)) {
        used[t.logicalName.trim().toLowerCase()] = true;
      }
    });
    const available: { [l: string]: boolean } | undefined = catalog.tables ? {} : undefined;
    (catalog.tables || []).forEach(t => { available![t.logicalName] = true; });
    const q = pickerQuery.trim().toLowerCase();
    const matches = (i: IPickerItem): boolean =>
      !q || i.title.toLowerCase().indexOf(q) >= 0 || i.logicalName.toLowerCase().indexOf(q) >= 0;

    const presetNames: { [l: string]: boolean } = {};
    const presets: IPickerItem[] = [];
    getTablePresets().forEach(p => {
      if (presetNames[p.logicalName] || (available && !available[p.logicalName])) {
        return;
      }
      presetNames[p.logicalName] = true;
      presets.push({ logicalName: p.logicalName, title: p.title, subtitle: `${p.logicalName} · vorkonfiguriert`, icon: p.icon, preset: p, used: !!used[p.logicalName] });
    });
    const rank = (t: ITableInfo): number => (/^admin_/.test(t.logicalName) ? 0 : t.isCustom ? 1 : 2);
    const others = (catalog.tables || [])
      .filter(t => !presetNames[t.logicalName])
      .sort((a, b) => rank(a) - rank(b) || a.displayName.localeCompare(b.displayName, 'de'))
      .map(t => ({
        logicalName: t.logicalName,
        title: t.displayName,
        subtitle: `${t.logicalName}${/^admin_/.test(t.logicalName) ? ' · CoE' : t.isCustom ? ' · benutzerdefiniert' : ''}`,
        icon: /^admin_/.test(t.logicalName) ? 'Org' : 'Table',
        used: !!used[t.logicalName]
      }));
    return { presets: presets.filter(matches), others: others.filter(matches), total: others.length };
  };

  const renderPickerItem = (item: IPickerItem): JSX.Element => (
    <button
      key={item.logicalName}
      type="button"
      className={styles.pickerItem}
      disabled={item.used}
      onClick={() => pickTable(item)}
      title={item.used ? 'Wird bereits angezeigt' : `${item.title} hinzufügen`}
    >
      <span className={styles.pickerItemIcon}><Icon iconName={item.icon || 'Table'} /></span>
      <span className={styles.pickerItemText}>
        <span className={styles.pickerItemTitle}>{item.title}</span>
        <span className={styles.pickerItemSub}>{item.subtitle}</span>
      </span>
      {item.used ? <span className={styles.pickerUsed}>bereits angezeigt</span> : <Icon iconName="Add" className={styles.pickerAdd} />}
    </button>
  );

  const renderPicker = (): JSX.Element => {
    const { presets, others, total } = buildPickerItems();
    const LIMIT = 200;
    return (
      <div className={styles.picker}>
        <div className={styles.pickerHead}>
          <h3 className={styles.sectionTitle} style={{ margin: 0 }}>
            {picker === 'change' ? 'Andere Tabelle wählen' : 'Tabelle hinzufügen'}
          </h3>
          <span className={styles.grow} />
          {draft.length > 0 && <DefaultButton text="Abbrechen" onClick={() => { setPicker(undefined); setPickerQuery(''); }} />}
        </div>
        <SearchBox
          placeholder="Tabelle suchen (Name oder logischer Name)"
          value={pickerQuery}
          onChange={(_e, v) => setPickerQuery(v || '')}
          styles={{ root: { margin: '12px 0 8px' } }}
        />
        {!service && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Zuerst die CoE-Umgebung in den Webpart-Eigenschaften auswählen – danach erscheinen hier alle Tabellen.
          </MessageBar>
        )}
        {catalog.status === 'error' && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
            Die Tabellenliste konnte nicht geladen werden ({catalog.error}). Die Vorlagen stehen trotzdem zur Auswahl.
          </MessageBar>
        )}

        <div className={styles.pickerGroupTitle}>Vorlagen aus dem CoE Starter Kit</div>
        {presets.length === 0
          ? <p className={styles.hint}>Keine passende Vorlage.</p>
          : <div className={styles.pickerList}>{presets.map(renderPickerItem)}</div>}

        <div className={styles.pickerGroupTitle}>
          Alle Tabellen der Umgebung{catalog.status === 'ready' ? ` (${total})` : ''}
        </div>
        {catalog.status === 'loading' && (
          <span className={styles.resolveInfo}><Spinner size={SpinnerSize.xSmall} /> Tabellen werden geladen …</span>
        )}
        {catalog.status === 'ready' && others.length === 0 && <p className={styles.hint}>Keine Tabelle gefunden.</p>}
        {others.length > 0 && (
          <div className={`${styles.pickerList} ${styles.pickerScroll}`}>
            {others.slice(0, LIMIT).map(renderPickerItem)}
          </div>
        )}
        {others.length > LIMIT && (
          <p className={styles.hint}>{others.length - LIMIT} weitere – Suche verfeinern, um sie zu sehen.</p>
        )}
      </div>
    );
  };

  const footer = (): JSX.Element => (
    <div className={styles.editorFooter}>
      <PrimaryButton text="Speichern" iconProps={{ iconName: 'Save' }} onClick={save} />
      <DefaultButton text="Abbrechen" onClick={onDismiss} />
      <span className={styles.grow} />
      <ActionButton iconProps={{ iconName: 'Refresh' }} text="Standard wiederherstellen" onClick={() => { setDraft(getDefaultTables()); setIndex(0); }} />
    </div>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.custom}
      customWidth="960px"
      headerText="Tabellen verwalten"
      closeButtonAriaLabel="Schließen"
      isFooterAtBottom={true}
      onRenderFooterContent={footer}
      styles={{ content: { paddingBottom: 24 } }}
    >
      <div style={themeVars} className={styles.scope}>
        <p className={styles.hint}>
          Lege fest, welche Dataverse-Tabellen das Cockpit anzeigt: Tabelle aus der Liste wählen (Vorlagen aus dem
          CoE Starter Kit oder jede andere Tabelle der Umgebung), danach Spalten aus der Liste auswählen.
          Änderungen werden mit der Seite gespeichert.
        </p>
        <Pivot selectedKey={mode} onLinkClick={item => {
          const k = (item && item.props.itemKey) || 'form';
          if (k === 'json') {
            setJsonText(JSON.stringify(draft, null, 2));
            setJsonError('');
          }
          setMode(k);
        }} headersOnly={true} styles={{ root: { margin: '12px 0 16px' } }}>
          <PivotItem itemKey="form" headerText="Editor" itemIcon="Edit" />
          <PivotItem itemKey="json" headerText="JSON Import/Export" itemIcon="Code" />
        </Pivot>

        {mode === 'json' && (
          <div>
            <TextField
              multiline={true}
              rows={22}
              value={jsonText}
              onChange={(_e, v) => setJsonText(v || '')}
              styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
            />
            {jsonError && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginTop: 8 } }}>{jsonError}</MessageBar>}
            <div className={styles.editorFooter} style={{ marginTop: 12 }}>
              <PrimaryButton text="In Editor übernehmen" onClick={applyJson} />
              <DefaultButton text="In Zwischenablage kopieren" iconProps={{ iconName: 'Copy' }} onClick={() => {
                const nav = navigator as Navigator & { clipboard?: { writeText: (t: string) => Promise<void> } };
                if (nav.clipboard) { nav.clipboard.writeText(jsonText).catch(() => undefined); }
              }} />
            </div>
          </div>
        )}

        {mode === 'form' && (
          <div className={styles.editor}>
            <nav className={styles.editorNav}>
              {draft.map((t, i) => (
                <button key={t.key + i} type="button" className={`${styles.editorNavItem} ${i === index ? styles.editorNavItemActive : ''}`} onClick={() => setIndex(i)}>
                  <Icon iconName={t.icon || 'Table'} />
                  <span className={styles.ellipsis}>{t.title || t.logicalName || 'Ohne Namen'}</span>
                  <span className={styles.editorNavMeta}>{t.columns.length}</span>
                </button>
              ))}
              <div className={styles.editorNavFooter}>
                <ActionButton iconProps={{ iconName: 'Add' }} text="Tabelle hinzufügen" onClick={() => { setPicker('add'); setPickerQuery(''); }} />
              </div>
            </nav>

            {picker && renderPicker()}

            {!picker && !current && (
              <div className={styles.state}>
                <div className={styles.stateIcon}><Icon iconName="Table" /></div>
                <p className={styles.stateTitle}>Keine Tabellen konfiguriert</p>
                <p className={styles.stateText}>Füge eine Tabelle hinzu oder stelle die Standardkonfiguration wieder her.</p>
              </div>
            )}

            {!picker && current && (
              <div>
                <h3 className={styles.sectionTitle}>Allgemein</h3>
                <div className={styles.formGrid}>
                  <TextField label="Anzeigename" value={current.title} onChange={(_e, v) => update({ title: v || '' })} />
                  <div>
                    <div className={styles.inline}>
                      <TextField
                        label="Dataverse-Tabelle"
                        value={current.logicalName}
                        readOnly={true}
                        styles={{ root: { flex: '1 1 auto' } }}
                      />
                      <DefaultButton text="Ändern" iconProps={{ iconName: 'Switch' }} onClick={() => { setPicker('change'); setPickerQuery(''); }} />
                    </div>
                    {resolveState && resolveState.status === 'loading' && (
                      <span className={styles.resolveInfo}><Spinner size={SpinnerSize.xSmall} /> Prüfe Tabelle …</span>
                    )}
                    {resolveState && resolveState.status === 'ready' && resolveState.meta && (
                      <span className={`${styles.resolveInfo} ${styles.toneSuccess}`}>
                        <span className={styles.dot} />
                        {resolveState.meta.displayName} · {resolveState.attributes ? resolveState.attributes.length : 0} Spalten
                      </span>
                    )}
                    {resolveState && resolveState.status === 'error' && (
                      <span className={`${styles.resolveInfo} ${styles.toneDanger}`}><span className={styles.dot} />{resolveState.error}</span>
                    )}
                    {!service && <p className={styles.hint}>Zum Prüfen zuerst die Umgebungs-URL in den Webpart-Eigenschaften eintragen.</p>}
                  </div>
                  <TextField
                    label="Symbol (Fluent UI Icon)"
                    value={current.icon}
                    onChange={(_e, v) => update({ icon: v || '' })}
                    onRenderSuffix={() => <Icon iconName={current.icon || 'Table'} />}
                    description="z. B. AppIconDefault, Flow, Globe, People, Table"
                  />
                  <TextField
                    label="Filter (OData)"
                    placeholder="z. B. statecode eq 0"
                    value={current.filter || ''}
                    onChange={(_e, v) => update({ filter: v || '' })}
                  />
                  <Dropdown
                    label="Standard-Sortierung"
                    selectedKey={current.defaultSort ? current.defaultSort.field : ''}
                    options={[{ key: '', text: '(keine)' } as IDropdownOption].concat(current.columns.filter(c => !!c.field).map(c => ({ key: c.field, text: c.label || c.field })))}
                    onChange={(_e, o) => update({ defaultSort: o && o.key ? { field: String(o.key), desc: current.defaultSort ? current.defaultSort.desc : false } : undefined })}
                  />
                  <Dropdown
                    label="Richtung"
                    selectedKey={current.defaultSort && current.defaultSort.desc ? 'desc' : 'asc'}
                    disabled={!current.defaultSort}
                    options={[{ key: 'asc', text: 'Aufsteigend' }, { key: 'desc', text: 'Absteigend' }]}
                    onChange={(_e, o) => current.defaultSort && update({ defaultSort: { field: current.defaultSort.field, desc: !!o && o.key === 'desc' } })}
                  />
                </div>

                <h3 className={styles.sectionTitle} style={{ marginTop: 24 }}>Kennzahl</h3>
                <div className={styles.formGrid}>
                  <Toggle
                    className={styles.span2}
                    label="In der Kennzahlen-Leiste anzeigen"
                    inlineLabel={true}
                    checked={current.showInKpis !== false}
                    onChange={(_e, v) => update({ showInKpis: !!v })}
                  />
                  <TextField
                    label="Hervorhebung – Text"
                    placeholder="z. B. verwaist"
                    value={(current.kpi && current.kpi.highlightLabel) || ''}
                    onChange={(_e, v) => update({ kpi: { ...(current.kpi || {}), highlightLabel: v || '' } })}
                  />
                  <Dropdown
                    label="Hervorhebung – Farbe"
                    selectedKey={(current.kpi && current.kpi.tone) || 'warning'}
                    options={TONE_OPTIONS}
                    onChange={(_e, o) => update({ kpi: { ...(current.kpi || {}), tone: (o ? o.key : 'warning') as Tone } })}
                  />
                  <TextField
                    className={styles.span2}
                    label="Hervorhebung – Filter (OData)"
                    placeholder="z. B. admin_appisorphaned eq 'Orphaned'"
                    value={(current.kpi && current.kpi.highlightFilter) || ''}
                    onChange={(_e, v) => update({ kpi: { ...(current.kpi || {}), highlightFilter: v || '' } })}
                  />
                </div>

                <h3 className={styles.sectionTitle} style={{ marginTop: 24 }}>Spalten</h3>
                <table className={styles.colTable}>
                  <thead>
                    <tr>
                      <th style={{ width: '36%' }}>Feld</th>
                      <th style={{ width: '26%' }}>Bezeichnung</th>
                      <th style={{ width: '20%' }}>Darstellung</th>
                      <th title="Standardmäßig ausgeblendet">Aus</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {current.columns.map((c, ci) => (
                      <React.Fragment key={ci}>
                        <tr>
                          <td>
                            <ComboBox
                              allowFreeform={true}
                              autoComplete="on"
                              text={c.field}
                              selectedKey={c.field || null}
                              options={attributeOptions}
                              placeholder="feldname"
                              onChange={(_e, o, _i, v) => onPickAttribute(ci, o ? String(o.key) : undefined, v)}
                              useComboBoxAsMenuWidth={false}
                              styles={{ optionsContainerWrapper: { maxHeight: 320 } }}
                            />
                          </td>
                          <td><TextField value={c.label} onChange={(_e, v) => updateColumn(ci, { label: v || '' })} /></td>
                          <td>
                            <Dropdown selectedKey={c.type} options={TYPE_OPTIONS} onChange={(_e, o) => updateColumn(ci, { type: (o ? o.key : 'text') as ColumnType })} />
                          </td>
                          <td><Checkbox checked={!!c.hidden} onChange={(_e, v) => updateColumn(ci, { hidden: !!v })} ariaLabel="Standardmäßig ausgeblendet" /></td>
                          <td>
                            <div className={styles.colActions}>
                              <IconButton iconProps={{ iconName: 'ChevronUp' }} title="Nach oben" disabled={ci === 0} onClick={() => moveColumn(ci, -1)} />
                              <IconButton iconProps={{ iconName: 'ChevronDown' }} title="Nach unten" disabled={ci === current.columns.length - 1} onClick={() => moveColumn(ci, 1)} />
                              <IconButton iconProps={{ iconName: 'Delete' }} title="Entfernen" onClick={() => removeColumn(ci)} />
                            </div>
                          </td>
                        </tr>
                        {c.type === 'status' && (
                          <tr>
                            <td colSpan={5}>
                              <TextField
                                key={`${current.key}|${c.field}|sm`}
                                prefix="Farben"
                                placeholder="Started=success; Stopped=neutral; Suspended=danger"
                                defaultValue={statusMapToText(c.statusMap)}
                                onBlur={e => updateColumn(ci, { statusMap: textToStatusMap((e.target as HTMLInputElement).value) })}
                                description="Wert=Farbe, getrennt mit Strichpunkt. Farben: success, warning, danger, info, accent, neutral"
                              />
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
                <ActionButton iconProps={{ iconName: 'Add' }} text="Spalte hinzufügen" onClick={addColumn} />

                <h3 className={styles.sectionTitle} style={{ marginTop: 24 }}>Links in der Detailansicht</h3>
                <TextField
                  multiline={true}
                  rows={3}
                  value={linkText}
                  onChange={(_e, v) => setLinkText(v || '')}
                  onBlur={() => update({ links: textToLinks(linkText) })}
                  placeholder="In Power Apps öffnen | https://make.powerapps.com/environments/{admin_appenvironmentid}/apps/{admin_appid}/details"
                  description="Eine Zeile pro Link: Bezeichnung | URL. Feldwerte mit {feldname} einsetzen."
                  styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
                />

                <div className={styles.editorFooter} style={{ marginTop: 24 }}>
                  <DefaultButton iconProps={{ iconName: 'Up' }} text="Nach oben" disabled={index === 0} onClick={() => moveTable(-1)} />
                  <DefaultButton iconProps={{ iconName: 'Down' }} text="Nach unten" disabled={index === draft.length - 1} onClick={() => moveTable(1)} />
                  <span className={styles.grow} />
                  <DefaultButton iconProps={{ iconName: 'Delete' }} text="Tabelle entfernen" onClick={removeTable} />
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </Panel>
  );
};
