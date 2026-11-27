import * as React from 'react';
import * as ReactDom from 'react-dom';
import { createTheme, ITheme, Panel, PanelType, TextField, Toggle, Dropdown, PrimaryButton, Label, initializeIcons } from '@fluentui/react';
import CoeCockpit from '../src/webparts/coeCockpit/components/CoeCockpit';
import { getDefaultTables } from '../src/webparts/coeCockpit/models/defaultTables';
import { ITableConfig } from '../src/webparts/coeCockpit/models/ITableConfig';
import { AadHttpClientFactory, AadTokenProviderFactory, MSGraphClientFactory, SPHttpClient } from './shims/sp-http';
import { seedActions, seedNotifyEntry } from './mock/sharepoint';
import { ACTION_TIMING } from '../src/webparts/coeCockpit/components/ActionDialog';
import { resetCreditsCache } from '../src/webparts/coeCockpit/components/CopilotCreditsView';
import { resetPremiumCache } from '../src/webparts/coeCockpit/components/PremiumView';
import { resetUserLicenseCache } from '../src/webparts/coeCockpit/models/userLicenses';
import { resetMessageCache } from '../src/webparts/coeCockpit/models/messageCenter';
import { resetDlpCache } from '../src/webparts/coeCockpit/models/dlp';
import { resetPipelineCache } from '../src/webparts/coeCockpit/models/pipelineHosts';
import { IWebPartSettings } from '../src/webparts/coeCockpit/models/settings';
import { IMailTemplateOverrides } from '../src/webparts/coeCockpit/models/mailTemplates';
import { ActionService, IActionSetupState } from '../src/webparts/coeCockpit/services/ActionService';

// Vorschau: Aktionen sind eingerichtet (Gruppe „CoE Cockpit Admins“ mit dem angemeldeten Admin, Liste vorhanden)
seedActions();
ACTION_TIMING.pollMs = 600;
import { DATABASES, COE_HOST, DEFAULT_HOST, PIPELINES_HOST } from './mock/data';
import { DEFAULT_INFO_URL } from '../src/webparts/coeCockpit/models/notification';
import { MAIL_TIMING } from '../src/webparts/coeCockpit/services/MailService';

// Vorschau: Versand ohne Exchange-Wartezeit
MAIL_TIMING.intervalMs = 250;

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

// Versandprotokoll: ein Developer-Environment hat ein anderer Admin schon vor zwei Tagen verständigt
const seededEnv = environments.filter(r => r.admin_environmentsku === 'Developer' && r.admin_governanceprotectionlevel === 'Basic' && r.admin_environmentruntimestate === 'Enabled' && Number(r.admin_numberofapps) > 0)[0];
// nur für Tests: Name → Environment-ID
(window as unknown as { __coeEnvIds?: { [name: string]: string } }).__coeEnvIds = environments.reduce((m, r) => { m[String(r.admin_displayname)] = String(r.admin_environmentid).toLowerCase(); return m; }, {} as { [name: string]: string });
if (seededEnv) {
  seedNotifyEntry({
    Title: seededEnv.admin_displayname, EnvironmentId: String(seededEnv.admin_environmentid).toLowerCase(),
    SentAt: new Date(Date.now() - 2 * 86400000).toISOString(), SentBy: 'Andrea Admin', SentByEmail: 'andrea.admin@contoso.com',
    Method: 'Automatisch gesendet (Microsoft Graph)', Recipients: '', Subject: 'Rückmeldung bis …'
  });
}

const App: React.FC = () => {
  // ?alteTabellen=1: gespeicherte Tabelleneinstellungen eines älteren Webparts nachstellen (ohne die unsichtbaren Aktionsspalten)
  const [tables, setTables] = React.useState<ITableConfig[]>(() => {
    const t = getDefaultTables();
    if (!/[?&]alteTabellen=1/.test(window.location.search)) {
      return t;
    }
    const drop = ['admin_issolutionapp', 'admin_issolutionflow', 'admin_appisquarantined', 'admin_quarantineapp'];
    return t.map(x => ({ ...x, columns: x.columns.filter(c => drop.indexOf(c.field) < 0) }));
  });
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
    pipelinesHostUrl: `https://${PIPELINES_HOST}`,
    notifyTone: 'sie' as 'sie' | 'du',
    notifyReplyDays: 14,
    notifyFeedbackUrl: '',
    notifyCc: 'powerplatform@contoso.com',
    notifySignature: '',
    notifyInfoUrl: '',
    notifyAutoSend: true,
    showOverview: true,
    showCopilotCredits: true,
    showActions: true,
    actionsGroup: 'CoE Cockpit Admins',
    deleteRequiresStop: true,
    showUserLicenses: true,
    hideAppBar: true,
    showOrphans: true,
    showLicenseOptimization: true,
    showStorage: true,
    showMessageCenter: true,
    showDlp: true,
    showPipelineHosts: true,
    mailTemplates: {} as IMailTemplateOverrides
  });
  const notify = React.useMemo(() => ({
    formal: p.notifyTone !== 'du', replyDays: p.notifyReplyDays, infoUrl: DEFAULT_INFO_URL,
    feedbackUrl: p.notifyFeedbackUrl, cc: p.notifyCc, signature: p.notifySignature, templates: p.mailTemplates
  }), [p.notifyTone, p.notifyReplyDays, p.notifyFeedbackUrl, p.notifyCc, p.notifySignature, p.mailTemplates]);
  const [noAdmin, setNoAdmin] = React.useState<boolean>(false);
  const [noMailSend, setNoMailSend] = React.useState<boolean>(false);
  const [noPpApi, setNoPpApi] = React.useState<boolean>(false);
  const [notMember, setNotMember] = React.useState<boolean>(false);
  const [noUserRead, setNoUserRead] = React.useState<boolean>(false);
  (window as unknown as { __coeDenyUserRead?: boolean }).__coeDenyUserRead = noUserRead;
  React.useMemo(() => { resetPremiumCache(); resetUserLicenseCache(); resetMessageCache(); resetDlpCache(); resetPipelineCache(true); }, [noUserRead, noAdmin]);
  (window as unknown as { __spNotMember?: boolean }).__spNotMember = notMember;
  // ohne Admin-Rolle liefert die DLP-Schnittstelle 403 – das Cockpit zeigt dann die Richtlinien aus dem CoE
  (window as unknown as { __coeBapDown?: boolean }).__coeBapDown = noAdmin;
  (window as unknown as { __coePpApiDeny?: string }).__coePpApiDeny = noPpApi ? 'consent' : noAdmin ? 'forbidden' : '';
  React.useMemo(() => resetCreditsCache(), [noAdmin, noPpApi]);
  (window as unknown as { __coeDenyMailSend?: boolean }).__coeDenyMailSend = noMailSend;
  (window as unknown as { __coeDenyHosts?: string[] }).__coeDenyHosts = noAdmin ? [DEFAULT_HOST, PIPELINES_HOST] : [];
  const factory = React.useMemo(() => new AadHttpClientFactory(), []);
  const tokenFactory = React.useMemo(() => new AadTokenProviderFactory(), []);
  (window as unknown as { __coeRoleInactive?: boolean }).__coeRoleInactive = noAdmin;
  const graphFactory = React.useMemo(() => new MSGraphClientFactory(), []);
  const spClient = React.useMemo(() => new SPHttpClient(), []);
  const set = (patch: Partial<typeof p>): void => setP(prev => ({ ...prev, ...patch }));

  // Webpart-Einstellungen (in SharePoint liefert das Webpart diese Werte und Rückrufe)
  const actionSvc = React.useMemo(() => new ActionService(spClient as never, 'https://contoso.sharepoint.com/sites/powerplatform'), [spClient]);
  const [setup, setSetup] = React.useState<{ groups: string[]; state?: IActionSetupState; busy: string; message: string }>({ groups: [], busy: '', message: '' });
  const loadSetup = (group: string): void => {
    Promise.all([actionSvc.listGroups(), actionSvc.getSetupState(group)]).then(([g, st]) => setSetup(prev => ({ ...prev, groups: g.map(x => x.title), state: st })));
  };
  const settings: IWebPartSettings = {
    values: { ...p, notifyInfoUrl: p.notifyInfoUrl },
    envOptions: [{ key: '', text: 'Automatisch erkennen (Umgebungstyp „Default“)' }].concat(
      environments.map(e => ({ key: String(e.admin_environmentid), text: `${e.admin_displayname} (${e.admin_environmentsku})` }))),
    envStatus: `${environments.length} Umgebungen aus dem CoE-Inventar verfügbar.`,
    managedOptions: [{ key: '', text: 'Automatisch: Governance Protection Level (admin_governanceprotectionlevel)' }, { key: 'admin_governanceprotectionlevel', text: 'Governance Protection Level (admin_governanceprotectionlevel)' }],
    managedInfo: 'Angezeigt werden Developer-Environments, die laut dieser Spalte nicht managed sind.',
    hostOptions: [{ key: '', text: 'Keine – ALM-Pipelines ausblenden' }].concat(
      environments.filter(e => !!e.admin_environmentcdsinstanceurl).map(e => ({ key: String(e.admin_environmentcdsinstanceurl).replace(/\/$/, ''), text: `${e.admin_displayname} (${e.admin_environmentsku})` }))),
    groupOptions: (setup.groups.indexOf(p.actionsGroup) < 0 ? [p.actionsGroup] : []).concat(setup.groups).map(g => ({ key: g, text: g })),
    actionSetup: setup.state,
    actionBusy: setup.busy,
    actionMessage: setup.message,
    tableCount: tables.length,
    defaultInfoUrl: DEFAULT_INFO_URL,
    onChange: patch => {
      set(patch as Partial<typeof p>);
      if (patch.actionsGroup) {
        loadSetup(patch.actionsGroup);
      }
    },
    onLoad: () => loadSetup(p.actionsGroup),
    onActionSetup: kind => {
      setSetup(prev => ({ ...prev, busy: kind === 'group' ? 'Gruppe wird angelegt …' : 'Liste wird eingerichtet …', message: '' }));
      (kind === 'group' ? actionSvc.createGroup(p.actionsGroup) : actionSvc.setupList(p.actionsGroup)).then(
        () => setSetup(prev => ({ ...prev, busy: '', message: kind === 'group' ? 'Gruppe ist angelegt.' : 'Liste ist eingerichtet.' })),
        (e: Error) => setSetup(prev => ({ ...prev, busy: '', message: `Fehler: ${e.message}` }))
      ).then(() => loadSetup(p.actionsGroup));
    },
    onOpenTableEditor: () => setEditorRequest(n => n + 1),
    onOpenPropertyPane: () => setPaneOpen(true)
  };

  return (
    <div className={`page ${dark ? 'pageDark' : ''}`}>
      {/* Nachbildung der SharePoint-App-Leiste (wird vom Cockpit ausgeblendet, wenn „hideAppBar“ an ist) */}
      <div id="sp-appBar" className="fakeAppBar">Discover · Publish · Build · OneDrive</div>
      <div className="previewBar">
        <span className="previewBadge">Vorschau</span>
        <span className="previewText">Beispieldaten · simulierte Dataverse-API · echter Webpart-Code</span>
        <span className="grow" />
        <Toggle inlineLabel label="Bearbeitungsmodus" checked={editMode} onChange={(_e, v) => setEditMode(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Dunkler Abschnitt" checked={dark} onChange={(_e, v) => setDark(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Rotes Theme" checked={red} onChange={(_e, v) => setRed(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Ohne Adminrolle" checked={noAdmin} onChange={(_e, v) => setNoAdmin(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Mail.Send fehlt" checked={noMailSend} onChange={(_e, v) => setNoMailSend(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="PP-API nicht freigegeben" checked={noPpApi} onChange={(_e, v) => setNoPpApi(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Nicht in der Admin-Gruppe" checked={notMember} onChange={(_e, v) => setNotMember(!!v)} styles={{ root: { margin: 0 } }} />
        <Toggle inlineLabel label="Lizenzen nicht lesbar" checked={noUserRead} onChange={(_e, v) => setNoUserRead(!!v)} styles={{ root: { margin: 0 } }} />
      </div>
      <div className={`section ${dark ? 'sectionDark' : ''}`}>
        <CoeCockpit
          key={`${noAdmin ? 'noadmin' : 'admin'}-${noPpApi ? 'noppapi' : 'ppapi'}-${notMember ? 'guest' : 'member'}-${noUserRead ? 'nolic' : 'lic'}`}
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
          aadTokenProviderFactory={tokenFactory as never}
          msGraphClientFactory={graphFactory as never}
          spHttpClient={spClient as never}
          webUrl="https://contoso.sharepoint.com/sites/powerplatform"
          currentUser={{ displayName: 'Alex Admin', email: 'alex.admin@contoso.com' }}
          showDefaultEnvironment={p.showDefaultEnvironment}
          defaultEnvironmentId={p.defaultEnvironmentId}
          defaultEnvironmentDataverseUrl={p.defaultEnvironmentDataverseUrl}
          showDevEnvironments={p.showDevEnvironments}
          showOverview={p.showOverview}
          devEnvManagedField={p.devEnvManagedField}
          showSecurity={p.showSecurity}
          showCopilotCredits={p.showCopilotCredits}
          showUserLicenses={p.showUserLicenses}
          hideAppBar={p.hideAppBar}
          showOrphans={p.showOrphans}
          showLicenseOptimization={p.showLicenseOptimization}
          showStorage={p.showStorage}
          showMessageCenter={p.showMessageCenter}
          showDlp={p.showDlp}
          showPipelineHosts={p.showPipelineHosts}
          pipelinesHostUrl={p.pipelinesHostUrl}
          notify={notify}
          notifyAutoSend={p.notifyAutoSend}
          showActions={p.showActions}
          actionsGroup={p.actionsGroup}
          deleteRequiresStop={p.deleteRequiresStop}
          settings={editMode ? settings : undefined}
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
        <Label styles={{ root: { marginTop: 16 } }}>E-Mail an Maker (Betroffen ab 15.12.)</Label>
        <Dropdown label="Anrede" selectedKey={p.notifyTone} options={[{ key: 'sie', text: 'Sie' }, { key: 'du', text: 'Du' }]}
          onChange={(_e, o) => set({ notifyTone: (o ? o.key : 'sie') as 'sie' | 'du' })} />
        <Dropdown label="Frist für die Rückmeldung" selectedKey={p.notifyReplyDays} options={[7, 10, 14, 21].map(n => ({ key: n, text: `${n} Tage` }))}
          onChange={(_e, o) => set({ notifyReplyDays: o ? Number(o.key) : 14 })} />
        <TextField label="Formular für die Rückmeldung (optional)" value={p.notifyFeedbackUrl} onChange={(_e, v) => set({ notifyFeedbackUrl: v || '' })} />
        <TextField label="CC (optional)" value={p.notifyCc} onChange={(_e, v) => set({ notifyCc: v || '' })} />
        <TextField label="Grußzeile" placeholder="Ihr Power-Platform-Team" value={p.notifySignature} onChange={(_e, v) => set({ notifySignature: v || '' })} />
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
