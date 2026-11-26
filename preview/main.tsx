import * as React from 'react';
import * as ReactDom from 'react-dom';
import { createTheme, ITheme, Panel, PanelType, TextField, Toggle, Dropdown, PrimaryButton, Label, initializeIcons } from '@fluentui/react';
import CoeCockpit from '../src/webparts/coeCockpit/components/CoeCockpit';
import { getDefaultTables } from '../src/webparts/coeCockpit/models/defaultTables';
import { ITableConfig } from '../src/webparts/coeCockpit/models/ITableConfig';
import { AadHttpClientFactory } from './shims/sp-http';
import { DATABASES, COE_HOST, DEFAULT_HOST, PIPELINES_HOST } from './mock/data';

initializeIcons();

const lightTheme: ITheme = createTheme({
  palette: {
    themePrimary: '#0f6cbd', themeLighterAlt: '#f3f8fc', themeLighter: '#d0e3f4', themeLight: '#a9cbeb', themeTertiary: '#5d9cd7',
    themeSecondary: '#2278c6', themeDarkAlt: '#0e61aa', themeDark: '#0c528f', themeDarker: '#093c6a',
    neutralLighterAlt: '#faf9f8', neutralLighter: '#f3f2f1', neutralLight: '#edebe9', neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0', neutralTertiaryAlt: '#c8c6c4', neutralTertiary: '#a19f9d', neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39', neutralPrimary: '#323130', neutralDark: '#201f1e', black: '#000000', white: '#ffffff'
  }
});

/** Rotes SharePoint-Theme (wie bei euch im Intranet) */
const redTheme: ITheme = createTheme({
  palette: {
    themePrimary: '#c8102e', themeLighterAlt: '#fdf4f5', themeLighter: '#f6d3d9', themeLight: '#eeaeb9', themeTertiary: '#de6578',
    themeSecondary: '#ce2641', themeDarkAlt: '#b40f29', themeDark: '#980d23', themeDarker: '#700919',
    neutralLighterAlt: '#faf9f8', neutralLighter: '#f3f2f1', neutralLight: '#edebe9', neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0', neutralTertiaryAlt: '#c8c6c4', neutralTertiary: '#a19f9d', neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39', neutralPrimary: '#323130', neutralDark: '#201f1e', black: '#000000', white: '#ffffff'
  }
});

const darkTheme: ITheme = createTheme({
  isInverted: true,
  palette: {
    themePrimary: '#479ef5', themeLighterAlt: '#03060a', themeLighter: '#0b1927', themeLight: '#152f4a', themeTertiary: '#2b5f93',
    themeSecondary: '#3f8bd8', themeDarkAlt: '#59a8f6', themeDark: '#72b5f7', themeDarker: '#97c9f9',
    neutralLighterAlt: '#262626', neutralLighter: '#2e2e2e', neutralLight: '#3b3b3b', neutralQuaternaryAlt: '#434343',
    neutralQuaternary: '#4a4a4a', neutralTertiaryAlt: '#686868', neutralTertiary: '#8a8a8a', neutralSecondary: '#c7c7c7',
    neutralPrimaryAlt: '#dadada', neutralPrimary: '#f5f5f5', neutralDark: '#f8f8f8', black: '#ffffff', white: '#1f1f1f'
  },
  semanticColors: { bodyBackground: '#1f1f1f', bodyText: '#f5f5f5', bodySubtext: '#c7c7c7' }
});

const environments = DATABASES[0].tables.admin_environment.rows.filter(r => !r.admin_environmentdeleted);

const App: React.FC = () => {
  const [tables, setTables] = React.useState<ITableConfig[]>(getDefaultTables());
  const [editMode, setEditMode] = React.useState<boolean>(true);
  const [dark, setDark] = React.useState<boolean>(false);
  const [red, setRed] = React.useState<boolean>(false);
  const [paneOpen, setPaneOpen] = React.useState<boolean>(false);
  const [editorRequest, setEditorRequest] = React.useState<number>(0);
  const [p, setP] = React.useState({
    title: 'CoE Cockpit',
    subtitle: 'Power Platform Center of Excellence',
    environmentUrl: `https://${COE_HOST}`,
    showKpis: true,
    density: 'comfortable' as 'compact' | 'comfortable',
    pageSize: 25,
    maxRecords: 20000,
    showDefaultEnvironment: true,
    defaultEnvironmentId: '',
    defaultEnvironmentDataverseUrl: '',
    showDevEnvironments: true,
    devEnvManagedField: '',
    showSecurity: true,
    pipelinesHostUrl: `https://${PIPELINES_HOST}`
  });
  const [noAdmin, setNoAdmin] = React.useState<boolean>(false);
  (window as unknown as { __coeDenyHosts?: string[] }).__coeDenyHosts = noAdmin ? [DEFAULT_HOST, PIPELINES_HOST] : [];
  const factory = React.useMemo(() => new AadHttpClientFactory(), []);
  const set = (patch: Partial<typeof p>): void => setP(prev => ({ ...prev, ...patch }));

  return (
    <div className={`page ${dark ? 'pageDark' : ''}`}>
      <div className="previewBar">
        <span className="previewBadge">Vorschau</span>
        <span className="previewText">Beispieldaten · simulierte Dataverse-API · echter Webpart-Code</span>
        <span className="grow" />
        <Toggle inlineLabel label="Bearbeitungsmodus" checked={editMode} onChange={(_e, v) => setEditMode(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Dunkler Abschnitt" checked={dark} onChange={(_e, v) => setDark(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Rotes Theme" checked={red} onChange={(_e, v) => setRed(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Ohne Adminrolle" checked={noAdmin} onChange={(_e, v) => setNoAdmin(!!v)} styles={{ root: { margin: 0 } }} />
      </div>
      <div className={`section ${dark ? 'sectionDark' : ''}`}>
        <CoeCockpit
          key={noAdmin ? 'noadmin' : 'admin'}
          instanceId="preview"
          title={p.title}
          subtitle={p.subtitle}
          environmentUrl={p.environmentUrl}
          tables={tables}
          pageSize={p.pageSize}
          maxRecords={p.maxRecords}
          showKpis={p.showKpis}
          density={p.density}
          theme={(dark ? darkTheme : red ? redTheme : lightTheme) as never}
          isEditMode={editMode}
          editorRequest={editorRequest}
          aadHttpClientFactory={factory as never}
          showDefaultEnvironment={p.showDefaultEnvironment}
          defaultEnvironmentId={p.defaultEnvironmentId}
          defaultEnvironmentDataverseUrl={p.defaultEnvironmentDataverseUrl}
          showDevEnvironments={p.showDevEnvironments}
          showOverview={true}
          devEnvManagedField={p.devEnvManagedField}
          showSecurity={p.showSecurity}
          pipelinesHostUrl={p.pipelinesHostUrl}
          onTablesChanged={t => setTables(t)}
          onConfigure={() => setPaneOpen(true)}
        />
      </div>

      <Panel
        isOpen={paneOpen}
        onDismiss={() => setPaneOpen(false)}
        type={PanelType.custom}
        customWidth="340px"
        headerText="Webpart-Eigenschaften"
        isLightDismiss
      >
        <p className="paneNote">In SharePoint ist das der Eigenschaftenbereich des Webparts.</p>
        <TextField label="URL der CoE-Umgebung" value={p.environmentUrl} onChange={(_e, v) => set({ environmentUrl: v || '' })} />
        <Label styles={{ root: { marginTop: 16 } }}>Default Environment</Label>
        <Toggle label="Bereich „Default Environment“ anzeigen" checked={p.showDefaultEnvironment} onChange={(_e, v) => set({ showDefaultEnvironment: !!v })} />
        <Dropdown
          label="Default Environment"
          selectedKey={p.defaultEnvironmentId}
          disabled={!p.showDefaultEnvironment}
          options={[{ key: '', text: 'Automatisch erkennen (Typ „Default“)' }].concat(
            environments.map(e => ({ key: String(e.admin_environmentid), text: `${e.admin_displayname} (${e.admin_environmentsku})` })))}
          onChange={(_e, o) => set({ defaultEnvironmentId: o ? String(o.key) : '' })}
        />
        <Label styles={{ root: { marginTop: 16 } }}>Betroffen ab 15.12. (Developer-Environments)</Label>
        <Toggle label="Reiter „Betroffen ab 15.12.“ anzeigen" checked={p.showDevEnvironments} onChange={(_e, v) => set({ showDevEnvironments: !!v })} />
        <Dropdown
          label="Spalte mit dem Managed-Status"
          selectedKey={p.devEnvManagedField}
          disabled={!p.showDevEnvironments}
          options={[{ key: '', text: 'Automatisch: Governance Protection Level (admin_governanceprotectionlevel)' }, { key: 'admin_governanceprotectionlevel', text: 'Governance Protection Level (admin_governanceprotectionlevel)' }]}
          onChange={(_e, o) => set({ devEnvManagedField: o ? String(o.key) : '' })}
        />
        <Label styles={{ root: { marginTop: 16 } }}>Sicherheit &amp; ALM-Pipelines</Label>
        <Toggle label="Ansicht „Sicherheit & Freigaben“ anzeigen" checked={p.showSecurity} onChange={(_e, v) => set({ showSecurity: !!v })} />
        <Dropdown
          label="Pipelines-Host-Environment"
          selectedKey={p.pipelinesHostUrl}
          options={[{ key: '', text: 'Keine – ALM-Pipelines ausblenden' }].concat(
            environments.filter(e => !!e.admin_environmentcdsinstanceurl).map(e => ({ key: String(e.admin_environmentcdsinstanceurl).replace(/\/$/, ''), text: `${e.admin_displayname} (${e.admin_environmentsku})` })))}
          onChange={(_e, o) => set({ pipelinesHostUrl: o ? String(o.key) : '' })}
        />
        <Label styles={{ root: { marginTop: 16 } }}>Tabellen</Label>
        <PrimaryButton text="Tabellen-Editor öffnen" iconProps={{ iconName: 'Settings' }} onClick={() => { setPaneOpen(false); setEditorRequest(n => n + 1); }} />
        <Label styles={{ root: { marginTop: 16 } }}>Darstellung</Label>
        <TextField label="Titel" value={p.title} onChange={(_e, v) => set({ title: v || '' })} />
        <TextField label="Untertitel" value={p.subtitle} onChange={(_e, v) => set({ subtitle: v || '' })} />
        <Toggle label="Kennzahlen anzeigen" checked={p.showKpis} onChange={(_e, v) => set({ showKpis: !!v })} styles={{ root: { marginTop: 12 } }} />
        <Dropdown label="Zeilenhöhe" selectedKey={p.density}
          options={[{ key: 'comfortable', text: 'Normal' }, { key: 'compact', text: 'Kompakt' }]}
          onChange={(_e, o) => set({ density: (o ? o.key : 'comfortable') as 'compact' | 'comfortable' })} />
        <Dropdown label="Einträge pro Seite" selectedKey={p.pageSize}
          options={[10, 25, 50, 100].map(n => ({ key: n, text: String(n) }))}
          onChange={(_e, o) => set({ pageSize: o ? Number(o.key) : 25 })} />
        <Dropdown label="Maximale Anzahl geladener Einträge pro Tabelle" selectedKey={p.maxRecords}
          options={[5000, 10000, 20000, 50000].map(n => ({ key: n, text: n.toLocaleString('de-AT') }))}
          onChange={(_e, o) => set({ maxRecords: o ? Number(o.key) : 20000 })} />
      </Panel>
    </div>
  );
};

ReactDom.render(<App />, document.getElementById('root'));
