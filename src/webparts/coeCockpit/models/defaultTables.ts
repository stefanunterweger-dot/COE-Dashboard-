import { IColumnConfig, IRelatedConfig, ITableConfig } from './ITableConfig';

/**
 * Vorkonfigurierte Tabellen – Nachbau der CoE-App „Power Platform Admin View“ (CoE Starter Kit, Core Components).
 * Spalten folgen den Standardansichten der Admin View, die Reiter der Detailansicht den Formularen
 * („Shared with“, „Connections“, „Usage“, „Resources“, „Used in“ …). Nur lesend.
 * Spalten, die es in der eigenen CoE-Version nicht gibt, werden beim Laden automatisch übersprungen.
 */

export const GROUP_INVENTORY = 'Inventar';
export const GROUP_USERS = 'Benutzer';
export const GROUP_USAGE = 'Nutzung';

const YES_NO_ORPHANED = { 'Yes': 'warning', 'yes': 'warning', 'Ja': 'warning', 'No': 'success', 'no': 'success', 'Nein': 'success' } as const;
const ROLE_MAP = { 'Owner': 'accent', 'CanEdit': 'info', 'CanView': 'neutral', 'CanViewWithComponentImport': 'neutral' } as const;

// ---------- wiederverwendete Reiter der Detailansicht ----------
function sharedWith(lookupField: string): IRelatedConfig {
  return {
    key: 'shared',
    title: 'Freigegeben für',
    logicalName: 'admin_powerplatformuserrole',
    lookupField,
    filter: 'statecode eq 0',
    columns: [
      { field: '_admin_powerplatformuser_value', label: 'Benutzer / Gruppe', type: 'person' },
      { field: 'admin_rolename', label: 'Rolle', type: 'status', statusMap: { ...ROLE_MAP } }
    ]
  };
}

function connections(lookupField: string, withSource: boolean): IRelatedConfig {
  const columns: IColumnConfig[] = [{ field: '_admin_connector_value', label: 'Connector', type: 'title' }];
  if (withSource) {
    columns.push({ field: 'admin_flow_connectionreferencesource', label: 'Verwendung', type: 'text' });
  }
  return { key: 'connections', title: 'Verbindungen', logicalName: 'admin_connectionreference', lookupField, columns };
}

const RESOURCE_APPS = (lookupField: string, title: string = 'Apps'): IRelatedConfig => ({
  key: `apps_${lookupField}`, title, logicalName: 'admin_app', lookupField,
  filter: 'admin_appdeleted ne true', defaultSort: { field: 'admin_appmodifiedon', desc: true },
  columns: [
    { field: 'admin_displayname', label: 'App', type: 'title' },
    { field: '_admin_appowner_value', label: 'Besitzer', type: 'person' },
    { field: '_admin_appenvironment_value', label: 'Environment', type: 'text' },
    { field: 'admin_powerappstype', label: 'Typ', type: 'choice' },
    { field: 'admin_appmodifiedon', label: 'Geändert', type: 'date' }
  ]
});

const RESOURCE_FLOWS = (lookupField: string, title: string = 'Flows'): IRelatedConfig => ({
  key: `flows_${lookupField}`, title, logicalName: 'admin_flow', lookupField,
  filter: 'admin_flowdeleted ne true', defaultSort: { field: 'admin_flowmodifiedon', desc: true },
  columns: [
    { field: 'admin_displayname', label: 'Flow', type: 'title' },
    { field: '_admin_derivedowner_value', label: 'Besitzer', type: 'person' },
    { field: '_admin_flowenvironment_value', label: 'Environment', type: 'text' },
    { field: 'admin_flowstate', label: 'Status', type: 'status', statusMap: { 'Started': 'success', 'Stopped': 'neutral', 'Suspended': 'danger' } },
    { field: 'admin_flowmodifiedon', label: 'Geändert', type: 'date' }
  ]
});

export function getDefaultTables(): ITableConfig[] {
  return [
    // ===================== Inventar =====================
    {
      key: 'apps',
      title: 'Apps',
      logicalName: 'admin_app',
      icon: 'AppIconDefault',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0 and admin_appdeleted ne true',
      showInKpis: true,
      defaultSort: { field: 'admin_appmodifiedon', desc: true },
      // „App is Orphaned“ ist im CoE ein Textfeld mit „Yes“/„No“ (Vergleich in Dataverse ohne Groß-/Kleinschreibung)
      kpi: { highlightLabel: 'verwaist', highlightFilter: "cr5d5_appisorphaned eq 'Yes'", tone: 'warning' },
      columns: [
        { field: 'admin_displayname', label: 'App', type: 'title' },
        { field: '_admin_appowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_appenvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_powerappstype', label: 'Typ', type: 'choice' },
        { field: 'cr5d5_appisorphaned', label: 'Verwaist', type: 'status', statusMap: { ...YES_NO_ORPHANED } },
        { field: 'admin_applastlaunchedon', label: 'Zuletzt gestartet', type: 'date' },
        { field: 'admin_appmodifiedon', label: 'Geändert', type: 'date' },
        { field: 'admin_appplanclassification', label: 'Lizenzklasse', type: 'choice', hidden: true },
        { field: 'admin_appsharedusers', label: 'Freigegeben (User)', type: 'number', hidden: true },
        { field: 'admin_appsharedgroups', label: 'Freigegeben (Gruppen)', type: 'number', hidden: true },
        { field: 'admin_appcreatedon', label: 'Erstellt', type: 'date', hidden: true },
        { field: 'admin_appenvironmentid', label: 'Environment-ID', type: 'text', hidden: true }
      ],
      links: [
        { label: 'In Power Apps öffnen', url: 'https://make.powerapps.com/environments/{admin_appenvironmentid}/apps/{admin_appid}/details' }
      ],
      related: [
        sharedWith('_admin_app_value'),
        connections('_admin_app_value', false),
        {
          key: 'usage', title: 'Nutzung', logicalName: 'admin_auditlog', lookupField: '_admin_applookup_value',
          filter: "admin_operation eq 'LaunchPowerApp'", defaultSort: { field: 'admin_creationtime', desc: true },
          columns: [
            { field: 'admin_userupn', label: 'Benutzer', type: 'text' },
            { field: 'admin_creationtime', label: 'Gestartet', type: 'datetime' }
          ]
        }
      ]
    },
    {
      key: 'flows',
      title: 'Flows',
      logicalName: 'admin_flow',
      icon: 'Flow',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0 and admin_flowdeleted ne true',
      showInKpis: true,
      defaultSort: { field: 'admin_flowmodifiedon', desc: true },
      kpi: { highlightLabel: 'gesperrt', highlightFilter: "admin_flowstate eq 'Suspended'", tone: 'danger' },
      columns: [
        { field: 'admin_displayname', label: 'Flow', type: 'title' },
        { field: '_admin_derivedowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_flowenvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_flowstate', label: 'Status', type: 'status', statusMap: { 'Started': 'success', 'Stopped': 'neutral', 'Suspended': 'danger' } },
        { field: 'cr5d5_flowisorphaned', label: 'Verwaist', type: 'status', statusMap: { ...YES_NO_ORPHANED } },
        { field: 'admin_flowmodifiedon', label: 'Geändert', type: 'date' },
        { field: '_admin_flowcreator_value', label: 'Ersteller', type: 'person', hidden: true },
        { field: 'cr5d5_flowtrigger', label: 'Trigger', type: 'text', hidden: true },
        { field: 'admin_flowidstring', label: 'Flow-ID', type: 'text', hidden: true },
        { field: 'admin_flowenvironmentid', label: 'Environment-ID', type: 'text', hidden: true }
      ],
      links: [
        { label: 'In Power Automate öffnen', url: 'https://make.powerautomate.com/environments/{admin_flowenvironmentid}/flows/{admin_flowidstring}/details' }
      ],
      related: [
        sharedWith('_admin_flow_value'),
        connections('_admin_flow_value', true),
        {
          key: 'actions', title: 'Aktionen', logicalName: 'admin_flowactiondetail', lookupField: '_admin_flow_value',
          columns: [
            { field: 'admin_actiontype', label: 'Typ', type: 'choice' },
            { field: 'admin_operation', label: 'Vorgang', type: 'text' },
            { field: 'cr5d5_istypetrigger', label: 'Trigger', type: 'boolean' }
          ]
        }
      ]
    },
    {
      key: 'environments',
      title: 'Environments',
      logicalName: 'admin_environment',
      icon: 'Globe',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0 and admin_environmentdeleted ne true and admin_excusefrominventory ne true',
      showInKpis: true,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Environment', type: 'title' },
        {
          field: 'admin_environmentsku', label: 'Typ', type: 'status',
          statusMap: { 'Production': 'accent', 'Default': 'info', 'Sandbox': 'neutral', 'Developer': 'warning', 'Trial': 'warning', 'Teams': 'info' }
        },
        { field: '_admin_maker_value', label: 'Ersteller', type: 'person' },
        { field: 'admin_governanceprotectionlevel', label: 'Managed', type: 'status', statusMap: { 'Standard': 'success', 'Basic': 'neutral' } },
        { field: 'admin_environmentruntimestate', label: 'Zustand', type: 'status', statusMap: { 'Enabled': 'success', 'Disabled': 'danger' } },
        { field: 'admin_numberofapps', label: 'Apps', type: 'number' },
        { field: 'admin_numberofflows', label: 'Flows', type: 'number' },
        { field: 'admin_region', label: 'Region', type: 'choice' },
        { field: 'admin_environmentcreatedon', label: 'Erstellt', type: 'date' },
        { field: 'admin_securitygroupname', label: 'Sicherheitsgruppe', type: 'text', hidden: true },
        { field: 'admin_environmentstate', label: 'Status', type: 'choice', hidden: true },
        { field: 'admin_environmentcdsinstanceurl', label: 'Dataverse-URL', type: 'url', hidden: true },
        { field: 'admin_environmentname', label: 'Environment-ID', type: 'text', hidden: true }
      ],
      links: [
        { label: 'Im Power Platform Admin Center öffnen', url: 'https://admin.powerplatform.microsoft.com/environments/{admin_environmentname}/hub' },
        { label: 'In Power Apps öffnen', url: 'https://make.powerapps.com/environments/{admin_environmentname}/home' }
      ],
      related: [
        RESOURCE_APPS('_admin_appenvironment_value'),
        RESOURCE_FLOWS('_admin_flowenvironment_value'),
        {
          key: 'bots', title: 'Bots', logicalName: 'admin_pva', lookupField: '_admin_pvaenvironment_value', filter: 'admin_pvadeleted ne true',
          columns: [
            { field: 'admin_name', label: 'Bot', type: 'title' },
            { field: '_admin_pvaowner_value', label: 'Besitzer', type: 'person' },
            { field: 'admin_pvastate', label: 'Status', type: 'choice' }
          ]
        },
        {
          key: 'desktopflows', title: 'Desktop Flows', logicalName: 'admin_rpa', lookupField: '_admin_desktopflowenvironment_value',
          columns: [
            { field: 'admin_displayname', label: 'Desktop Flow', type: 'title' },
            { field: '_admin_rpaowner_value', label: 'Besitzer', type: 'person' },
            { field: 'admin_rpastate', label: 'Status', type: 'choice' }
          ]
        },
        {
          key: 'pages', title: 'Power Pages', logicalName: 'admin_portal', lookupField: '_admin_portalenvironment_value',
          columns: [
            { field: 'admin_portaldisplayname', label: 'Site', type: 'title' },
            { field: '_admin_portalowner_value', label: 'Besitzer', type: 'person' }
          ]
        },
        {
          key: 'solutions', title: 'Solutions', logicalName: 'admin_solution', lookupField: '_admin_solutionenvironment_value', filter: 'admin_solutiondeleted ne true',
          columns: [
            { field: 'admin_name', label: 'Solution', type: 'title' },
            { field: '_admin_solutioncreator_value', label: 'Ersteller', type: 'person' },
            { field: 'admin_solutionismanaged', label: 'Verwaltet', type: 'boolean' }
          ]
        },
        {
          key: 'capacity', title: 'Kapazität', logicalName: 'admin_environmentcapacity', lookupField: '_admin_environment_value',
          columns: [
            { field: 'admin_capacitytype', label: 'Typ', type: 'title' },
            { field: 'admin_actualconsumption', label: 'Verbrauch', type: 'number' },
            { field: 'admin_approvedcapacity', label: 'Genehmigt', type: 'number' },
            { field: 'admin_capacityunit', label: 'Einheit', type: 'text' }
          ]
        },
        {
          key: 'roles', title: 'Sicherheitsrollen', logicalName: 'admin_environmentsecurityrolepermission', lookupField: '_admin_environment_value',
          columns: [
            { field: 'admin_name', label: 'Benutzer', type: 'person' },
            { field: 'admin_securityrolename', label: 'Rolle', type: 'choice' },
            { field: 'admin_assignedvia', label: 'Zugewiesen über', type: 'text' }
          ]
        }
      ]
    },
    {
      key: 'connectors',
      title: 'Connectors',
      logicalName: 'admin_connector',
      icon: 'PlugConnected',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0 and admin_connectordeleted ne true',
      showInKpis: false,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Connector', type: 'title' },
        { field: 'admin_tier', label: 'Stufe', type: 'status', statusMap: { 'Premium': 'warning', 'Standard': 'neutral', 'Custom': 'info' } },
        { field: 'admin_publisher', label: 'Herausgeber', type: 'choice' },
        { field: 'admin_iscustomapi', label: 'Benutzerdefiniert', type: 'boolean' },
        { field: 'admin_numapps', label: 'Apps', type: 'number' },
        { field: 'admin_numflows', label: 'Flows', type: 'number' },
        { field: 'admin_name', label: 'Name', type: 'text', hidden: true }
      ],
      links: [],
      related: [
        {
          key: 'usedapps', title: 'Verwendet in Apps', logicalName: 'admin_connectionreference', lookupField: '_admin_connector_value',
          filter: '_admin_app_value ne null',
          columns: [{ field: '_admin_app_value', label: 'App', type: 'title' }]
        },
        {
          key: 'usedflows', title: 'Verwendet in Flows', logicalName: 'admin_connectionreference', lookupField: '_admin_connector_value',
          filter: '_admin_flow_value ne null',
          columns: [{ field: '_admin_flow_value', label: 'Flow', type: 'title' }]
        }
      ]
    },
    {
      key: 'bots',
      title: 'Bots',
      logicalName: 'admin_pva',
      icon: 'Robot',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0 and admin_pvadeleted ne true',
      showInKpis: false,
      defaultSort: { field: 'admin_name', desc: false },
      columns: [
        { field: 'admin_name', label: 'Bot', type: 'title' },
        { field: '_admin_pvaowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_pvaenvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_pvastate', label: 'Status', type: 'status', statusMap: { 'Published': 'success', 'Provisioned': 'neutral' } },
        { field: 'admin_pvanumberlaunches', label: 'Unterhaltungen (letzter Monat)', type: 'number' },
        { field: 'admin_pvalastlaunchedon', label: 'Zuletzt verwendet', type: 'date' },
        { field: 'admin_bottemplate', label: 'Vorlage', type: 'text', hidden: true }
      ],
      links: []
    },
    {
      key: 'desktopflows',
      title: 'Desktop Flows',
      logicalName: 'admin_rpa',
      icon: 'Processing',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0',
      showInKpis: false,
      defaultSort: { field: 'admin_rpamodifiedon', desc: true },
      columns: [
        { field: 'admin_displayname', label: 'Desktop Flow', type: 'title' },
        { field: '_admin_rpaowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_desktopflowenvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_rpastate', label: 'Status', type: 'status', statusMap: { 'Published': 'success', 'Unpublished': 'neutral', 'Deleted': 'danger' } },
        { field: 'admin_rpamodifiedon', label: 'Geändert', type: 'date' }
      ],
      links: []
    },
    {
      key: 'pages',
      title: 'Power Pages Sites',
      logicalName: 'admin_portal',
      icon: 'Website',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0',
      showInKpis: false,
      defaultSort: { field: 'admin_portalmodifiedon', desc: true },
      columns: [
        { field: 'admin_portaldisplayname', label: 'Site', type: 'title' },
        { field: '_admin_portalowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_portalenvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_portalwebsitename', label: 'Website', type: 'text' },
        { field: 'admin_portalwebsiterecordstatus', label: 'Status', type: 'choice' },
        { field: 'admin_authtablepermissionsenabled', label: 'Tabellenberechtigungen', type: 'boolean' },
        { field: 'admin_portalmodifiedon', label: 'Geändert', type: 'date' }
      ],
      links: []
    },
    {
      key: 'bpf',
      title: 'Business Process Flows',
      logicalName: 'admin_businessprocessflows',
      icon: 'Flowchart',
      group: GROUP_INVENTORY,
      filter: 'admin_businessprocessflowdeleted ne true',
      showInKpis: false,
      defaultSort: { field: 'admin_businessprocessflowmodifiedon', desc: true },
      columns: [
        { field: 'admin_businessprocessflowdisplayname', label: 'Business Process Flow', type: 'title' },
        { field: '_admin_businessprocessflowowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_businessprocessflowenvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_businessprocessflowprimaryentity', label: 'Tabelle', type: 'choice' },
        { field: 'admin_businessprocessflowstatus', label: 'Status', type: 'status', statusMap: { 'Activated': 'success', 'Draft': 'neutral' } },
        { field: 'admin_businessprocessflowmodifiedon', label: 'Geändert', type: 'date' },
        { field: 'admin_name', label: 'Name', type: 'text', hidden: true }
      ],
      links: []
    },
    {
      key: 'solutions',
      title: 'Solutions',
      logicalName: 'admin_solution',
      icon: 'Package',
      group: GROUP_INVENTORY,
      filter: 'statecode eq 0 and admin_solutiondeleted ne true',
      showInKpis: false,
      defaultSort: { field: 'admin_name', desc: false },
      columns: [
        { field: 'admin_name', label: 'Solution', type: 'title' },
        { field: '_admin_solutionenvironment_value', label: 'Environment', type: 'text' },
        { field: '_admin_solutioncreator_value', label: 'Ersteller', type: 'person' },
        { field: 'admin_solutionpublisher', label: 'Herausgeber', type: 'choice' },
        { field: 'admin_solutionismanaged', label: 'Verwaltet', type: 'boolean' }
      ],
      links: []
    },
    {
      key: 'aimodels',
      title: 'AI-Modelle',
      logicalName: 'admin_aibuildermodel',
      icon: 'Lightbulb',
      group: GROUP_INVENTORY,
      filter: 'admin_aideleted ne true',
      showInKpis: false,
      defaultSort: { field: 'admin_ailastused', desc: true },
      columns: [
        { field: 'admin_name', label: 'AI-Modell', type: 'title' },
        { field: '_admin_aicreator_value', label: 'Ersteller', type: 'person' },
        { field: '_admin_aienvironment_value', label: 'Environment', type: 'text' },
        { field: 'admin_aitemplate', label: 'Vorlage', type: 'choice' },
        { field: 'admin_aistatus', label: 'Status', type: 'choice' },
        { field: 'admin_ailastused', label: 'Zuletzt verwendet', type: 'date' }
      ],
      links: []
    },

    // ===================== Benutzer =====================
    {
      key: 'makers',
      title: 'Maker',
      logicalName: 'admin_maker',
      icon: 'People',
      group: GROUP_USERS,
      filter: 'statecode eq 0',
      showInKpis: true,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Name', type: 'person' },
        { field: 'admin_userprincipalname', label: 'E-Mail / UPN', type: 'text' },
        { field: 'admin_department', label: 'Abteilung', type: 'choice' },
        { field: 'admin_company', label: 'Firma', type: 'choice' },
        { field: 'admin_country', label: 'Land', type: 'choice' },
        { field: 'admin_numberofapps', label: 'Apps', type: 'number' },
        { field: 'admin_numberofflows', label: 'Flows', type: 'number' },
        { field: 'admin_userisguest', label: 'Gast', type: 'boolean' },
        { field: 'admin_accountenabled', label: 'Konto aktiv', type: 'boolean', hidden: true },
        { field: 'admin_numberofenvironments', label: 'Environments', type: 'number', hidden: true }
      ],
      links: [],
      related: [
        RESOURCE_APPS('_admin_appowner_value', 'Apps'),
        RESOURCE_FLOWS('_admin_flowcreator_value', 'Flows'),
        {
          key: 'envs', title: 'Environments', logicalName: 'admin_environment', lookupField: '_admin_maker_value', filter: 'admin_environmentdeleted ne true',
          columns: [
            { field: 'admin_displayname', label: 'Environment', type: 'title' },
            { field: 'admin_environmentsku', label: 'Typ', type: 'choice' },
            { field: 'admin_environmentcreatedon', label: 'Erstellt', type: 'date' }
          ]
        }
      ]
    },
    {
      key: 'ppusers',
      title: 'Power Platform User',
      logicalName: 'admin_powerplatformuser',
      icon: 'Contact',
      group: GROUP_USERS,
      filter: 'statecode eq 0',
      showInKpis: false,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Name', type: 'person' },
        { field: 'admin_type', label: 'Typ', type: 'status', statusMap: { 'User': 'neutral', 'Group': 'info', 'Tenant': 'warning', 'ServicePrincipal': 'accent' } },
        { field: 'admin_userprincipalname', label: 'UPN', type: 'text' },
        { field: 'admin_userisguest', label: 'Gast', type: 'boolean' },
        { field: 'admin_groupsize', label: 'Gruppengröße', type: 'number' },
        { field: 'admin_useremail', label: 'E-Mail', type: 'text', hidden: true }
      ],
      links: [],
      related: [
        {
          key: 'roles', title: 'Rollen', logicalName: 'admin_powerplatformuserrole', lookupField: '_admin_powerplatformuser_value', filter: 'statecode eq 0',
          columns: [
            { field: 'admin_resourcetype', label: 'Art', type: 'choice' },
            { field: '_admin_app_value', label: 'App', type: 'text' },
            { field: '_admin_flow_value', label: 'Flow', type: 'text' },
            { field: 'admin_rolename', label: 'Rolle', type: 'status', statusMap: { ...ROLE_MAP } },
            { field: '_admin_environment_value', label: 'Environment', type: 'text' }
          ]
        }
      ]
    },
    {
      key: 'userroles',
      title: 'Benutzerrollen',
      logicalName: 'admin_powerplatformuserrole',
      icon: 'Permissions',
      group: GROUP_USERS,
      filter: 'statecode eq 0',
      showInKpis: false,
      columns: [
        { field: '_admin_powerplatformuser_value', label: 'Benutzer / Gruppe', type: 'person' },
        { field: 'admin_resourcetype', label: 'Art', type: 'status', statusMap: { 'App': 'info', 'Flow': 'accent', 'Environment': 'neutral' } },
        { field: '_admin_app_value', label: 'App', type: 'text' },
        { field: '_admin_flow_value', label: 'Flow', type: 'text' },
        { field: 'admin_rolename', label: 'Rolle', type: 'status', statusMap: { ...ROLE_MAP } },
        { field: '_admin_environment_value', label: 'Environment', type: 'text' }
      ],
      links: []
    },
    {
      key: 'securityroles',
      title: 'Sicherheitsrollen',
      logicalName: 'admin_environmentsecurityrolepermission',
      icon: 'Shield',
      group: GROUP_USERS,
      showInKpis: false,
      defaultSort: { field: 'admin_name', desc: false },
      columns: [
        { field: 'admin_name', label: 'Benutzer', type: 'person' },
        { field: 'admin_securityrolename', label: 'Sicherheitsrolle', type: 'choice' },
        { field: '_admin_environment_value', label: 'Environment', type: 'text' },
        { field: 'admin_assignedvia', label: 'Zugewiesen über', type: 'choice' },
        { field: 'admin_groupname', label: 'Gruppe', type: 'text' },
        { field: 'admin_usertype', label: 'Benutzertyp', type: 'choice' },
        { field: 'admin_businessunitname', label: 'Unternehmenseinheit', type: 'text', hidden: true },
        { field: 'admin_useridentityprovider', label: 'Identitätsanbieter', type: 'text', hidden: true }
      ],
      links: []
    },
    {
      key: 'connidentities',
      title: 'Verbindungsidentitäten',
      logicalName: 'admin_connectionreferenceidentity',
      icon: 'ContactCard',
      group: GROUP_USERS,
      showInKpis: false,
      defaultSort: { field: 'admin_accountname', desc: false },
      columns: [
        { field: 'admin_accountname', label: 'Konto', type: 'title' },
        { field: '_admin_connector_value', label: 'Connector', type: 'choice' },
        { field: '_admin_environment_value', label: 'Environment', type: 'text' },
        { field: '_admin_connectionreferencecreator_value', label: 'Ersteller', type: 'person' },
        { field: 'admin_noneorcrosstenantidentity', label: 'Fremder/kein Tenant', type: 'boolean' },
        { field: 'admin_name', label: 'Name', type: 'text', hidden: true }
      ],
      links: []
    },

    // ===================== Nutzung =====================
    {
      key: 'capacity',
      title: 'Environment-Kapazität',
      logicalName: 'admin_environmentcapacity',
      icon: 'Database',
      group: GROUP_USAGE,
      filter: 'statecode eq 0',
      showInKpis: false,
      defaultSort: { field: 'admin_actualconsumption', desc: true },
      columns: [
        { field: '_admin_environment_value', label: 'Environment', type: 'title' },
        { field: 'admin_capacitytype', label: 'Kapazitätstyp', type: 'choice' },
        { field: 'admin_actualconsumption', label: 'Verbrauch', type: 'number' },
        { field: 'admin_ratedconsumption', label: 'Bewerteter Verbrauch', type: 'number' },
        { field: 'admin_approvedcapacity', label: 'Genehmigt', type: 'number' },
        { field: 'admin_percentusedofapproved', label: '% genutzt', type: 'number' },
        { field: 'admin_capacityunit', label: 'Einheit', type: 'text' }
      ],
      links: []
    },
    {
      key: 'auditlog',
      title: 'Audit-Log',
      logicalName: 'admin_auditlog',
      icon: 'History',
      group: GROUP_USAGE,
      showInKpis: false,
      defaultSort: { field: 'admin_creationtime', desc: true },
      columns: [
        { field: 'admin_creationtime', label: 'Zeitpunkt', type: 'datetime' },
        { field: 'admin_userupn', label: 'Benutzer', type: 'text' },
        { field: 'admin_operation', label: 'Vorgang', type: 'choice' },
        { field: '_admin_applookup_value', label: 'App', type: 'text' },
        { field: 'admin_appenvironment', label: 'Environment', type: 'text' },
        { field: 'admin_workload', label: 'Workload', type: 'choice', hidden: true }
      ],
      links: []
    }
  ];
}
