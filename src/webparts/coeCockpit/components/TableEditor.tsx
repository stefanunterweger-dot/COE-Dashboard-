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
import { DataverseService, IAttributeMeta, IEntityMeta } from '../services/DataverseService';

export interface ITableEditorProps {
  isOpen: boolean;
  tables: ITableConfig[];
  service: DataverseService | undefined;
  themeVars: React.CSSProperties;
  onSave: (tables: ITableConfig[]) => void;
  onDismiss: () => void;
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
  { key: 'url', text: 'Link' }
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

  React.useEffect(() => {
    if (isOpen) {
      setDraft(clone(tables));
      setIndex(0);
      setMode('form');
      setJsonError('');
    }
  }, [isOpen]);

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

  const addTable = (): void => {
    const t: ITableConfig = {
      key: `t${Date.now().toString(36)}`,
      title: 'Neue Tabelle',
      logicalName: '',
      icon: 'Table',
      showInKpis: true,
      columns: [],
      links: []
    };
    setDraft(prev => prev.concat([t]));
    setIndex(draft.length);
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

  const resolveTable = (): void => {
    if (!service || !current || !current.logicalName.trim()) {
      return;
    }
    const logical = current.logicalName.trim().toLowerCase();
    setResolve(prev => ({ ...prev, [logical]: { status: 'loading' } }));
    Promise.all([service.getEntityMeta(logical), service.getAttributes(logical)]).then(
      ([meta, attributes]) => {
        setResolve(prev => ({ ...prev, [logical]: { status: 'ready', meta, attributes } }));
        setDraft(prev => prev.map((t, i) => {
          if (i !== index) {
            return t;
          }
          const patch: Partial<ITableConfig> = { logicalName: logical };
          if (!t.title || t.title === 'Neue Tabelle') {
            patch.title = meta.displayName;
          }
          if (t.columns.length === 0) {
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
          Lege fest, welche Dataverse-Tabellen das Cockpit anzeigt. Gib den logischen Tabellennamen ein und klicke auf
          „Prüfen“ – danach kannst du Spalten bequem aus der Liste wählen. Änderungen werden mit der Seite gespeichert.
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
                <ActionButton iconProps={{ iconName: 'Add' }} text="Tabelle hinzufügen" onClick={addTable} />
              </div>
            </nav>

            {!current && (
              <div className={styles.state}>
                <div className={styles.stateIcon}><Icon iconName="Table" /></div>
                <p className={styles.stateTitle}>Keine Tabellen konfiguriert</p>
                <p className={styles.stateText}>Füge eine Tabelle hinzu oder stelle die Standardkonfiguration wieder her.</p>
              </div>
            )}

            {current && (
              <div>
                <h3 className={styles.sectionTitle}>Allgemein</h3>
                <div className={styles.formGrid}>
                  <TextField label="Anzeigename" value={current.title} onChange={(_e, v) => update({ title: v || '' })} />
                  <div>
                    <div className={styles.inline}>
                      <TextField
                        label="Logischer Tabellenname"
                        placeholder="z. B. admin_app"
                        value={current.logicalName}
                        onChange={(_e, v) => update({ logicalName: v || '' })}
                        styles={{ root: { flex: '1 1 auto' } }}
                      />
                      <DefaultButton text="Prüfen" iconProps={{ iconName: 'CheckMark' }} onClick={resolveTable} disabled={!service || !current.logicalName.trim()} />
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
