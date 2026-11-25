import { ITableConfig } from './ITableConfig';

/**
 * Vorkonfigurierte Tabellen des CoE Starter Kits (Core Components).
 * Spalten, die es in deiner CoE-Version nicht gibt, werden beim Laden
 * automatisch übersprungen und als Hinweis angezeigt.
 */
export function getDefaultTables(): ITableConfig[] {
  return [
    {
      key: 'apps',
      title: 'Apps',
      logicalName: 'admin_app',
      icon: 'AppIconDefault',
      filter: 'admin_appdeleted ne true',
      showInKpis: true,
      defaultSort: { field: 'admin_appmodifiedon', desc: true },
      kpi: { highlightLabel: 'verwaist', highlightFilter: "admin_appisorphaned eq 'Orphaned'", tone: 'warning' },
      columns: [
        { field: 'admin_displayname', label: 'App', type: 'title' },
        { field: '_admin_appowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_appenvironment_value', label: 'Umgebung', type: 'text' },
        { field: 'admin_apptype', label: 'Typ', type: 'choice' },
        {
          field: 'admin_appisorphaned',
          label: 'Besitzerstatus',
          type: 'status',
          statusMap: { 'Orphaned': 'warning', 'Not Orphaned': 'success' }
        },
        { field: 'admin_applastlaunchedon', label: 'Zuletzt gestartet', type: 'date' },
        { field: 'admin_appmodifiedon', label: 'Geändert', type: 'date' },
        { field: 'admin_appid', label: 'App-ID', type: 'text', hidden: true },
        { field: 'admin_appenvironmentid', label: 'Umgebungs-ID', type: 'text', hidden: true }
      ],
      links: [
        {
          label: 'In Power Apps öffnen',
          url: 'https://make.powerapps.com/environments/{admin_appenvironmentid}/apps/{admin_appid}/details'
        }
      ]
    },
    {
      key: 'flows',
      title: 'Flows',
      logicalName: 'admin_flow',
      icon: 'Flow',
      filter: 'admin_flowdeleted ne true',
      showInKpis: true,
      defaultSort: { field: 'admin_flowmodifiedon', desc: true },
      kpi: { highlightLabel: 'gesperrt', highlightFilter: "admin_flowstate eq 'Suspended'", tone: 'danger' },
      columns: [
        { field: 'admin_displayname', label: 'Flow', type: 'title' },
        { field: '_admin_flowcreator_value', label: 'Ersteller', type: 'person' },
        { field: '_admin_flowenvironment_value', label: 'Umgebung', type: 'text' },
        {
          field: 'admin_flowstate',
          label: 'Status',
          type: 'status',
          statusMap: { 'Started': 'success', 'Stopped': 'neutral', 'Suspended': 'danger' }
        },
        {
          field: 'admin_flowisorphaned',
          label: 'Besitzerstatus',
          type: 'status',
          statusMap: { 'Orphaned': 'warning', 'Not Orphaned': 'success' }
        },
        { field: 'admin_flowmodifiedon', label: 'Geändert', type: 'date' },
        { field: 'admin_flowid', label: 'Flow-ID', type: 'text', hidden: true },
        { field: 'admin_flowenvironmentid', label: 'Umgebungs-ID', type: 'text', hidden: true }
      ],
      links: [
        {
          label: 'In Power Automate öffnen',
          url: 'https://make.powerautomate.com/environments/{admin_flowenvironmentid}/flows/{admin_flowid}/details'
        }
      ]
    },
    {
      key: 'environments',
      title: 'Environments',
      logicalName: 'admin_environment',
      icon: 'Globe',
      filter: 'admin_environmentdeleted ne true',
      showInKpis: true,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Umgebung', type: 'title' },
        {
          field: 'admin_environmenttype',
          label: 'Typ',
          type: 'status',
          statusMap: {
            'Production': 'accent',
            'Default': 'info',
            'Sandbox': 'neutral',
            'Developer': 'neutral',
            'Trial': 'warning',
            'Teams': 'info'
          }
        },
        { field: 'admin_environmentregion', label: 'Region', type: 'choice' },
        { field: 'admin_environmentcdsinstanceurl', label: 'Dataverse-URL', type: 'url' },
        { field: 'admin_environmentcreatedon', label: 'Erstellt', type: 'date' },
        { field: 'admin_environmentid', label: 'Umgebungs-ID', type: 'text', hidden: true }
      ],
      links: [
        {
          label: 'Im Power Platform Admin Center öffnen',
          url: 'https://admin.powerplatform.microsoft.com/environments/{admin_environmentid}/hub'
        }
      ]
    },
    {
      key: 'makers',
      title: 'Maker',
      logicalName: 'admin_maker',
      icon: 'People',
      showInKpis: true,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Name', type: 'person' },
        { field: 'admin_userprincipalname', label: 'E-Mail / UPN', type: 'text' },
        { field: 'admin_department', label: 'Abteilung', type: 'choice' },
        { field: 'admin_company', label: 'Firma', type: 'text' },
        { field: 'admin_country', label: 'Land', type: 'choice' },
        { field: 'createdon', label: 'Erfasst', type: 'date' }
      ],
      links: []
    }
  ];
}
