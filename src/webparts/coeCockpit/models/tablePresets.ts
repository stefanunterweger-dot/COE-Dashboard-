import { ITableConfig } from './ITableConfig';
import { getDefaultTables } from './defaultTables';

/**
 * Vorkonfigurierte Tabellen des CoE Starter Kits zur Auswahl im Tabellen-Editor.
 * Enthält die Standardtabellen (Apps, Flows, Environments, Maker) und weitere CoE-Tabellen.
 * Angeboten wird eine Vorlage nur, wenn es die Tabelle in der Umgebung gibt; Spalten, die es in
 * der eigenen CoE-Version nicht gibt, werden beim Hinzufügen automatisch entfernt.
 */
export function getTablePresets(): ITableConfig[] {
  const extra: ITableConfig[] = [
    {
      key: 'connectors',
      title: 'Connectors',
      logicalName: 'admin_connector',
      icon: 'PlugConnected',
      showInKpis: false,
      defaultSort: { field: 'admin_displayname', desc: false },
      columns: [
        { field: 'admin_displayname', label: 'Connector', type: 'title' },
        { field: 'admin_tier', label: 'Stufe', type: 'status', statusMap: { 'Premium': 'warning', 'Standard': 'neutral' } },
        { field: 'admin_publisher', label: 'Herausgeber', type: 'choice' },
        { field: 'admin_iscustomapi', label: 'Benutzerdefiniert', type: 'boolean' },
        { field: 'modifiedon', label: 'Geändert', type: 'date' }
      ],
      links: []
    },
    {
      key: 'copilots',
      title: 'Copilots',
      logicalName: 'admin_pva',
      icon: 'Robot',
      showInKpis: false,
      defaultSort: { field: 'modifiedon', desc: true },
      columns: [
        { field: 'admin_displayname', label: 'Copilot', type: 'title' },
        { field: '_admin_pvaowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_pvaenvironment_value', label: 'Umgebung', type: 'text' },
        { field: 'admin_pvastate', label: 'Status', type: 'status', statusMap: { 'Published': 'success', 'Provisioned': 'neutral' } },
        { field: 'modifiedon', label: 'Geändert', type: 'date' }
      ],
      links: []
    },
    {
      key: 'desktopflows',
      title: 'Desktop Flows',
      logicalName: 'admin_rpa',
      icon: 'ProcessMetaTask',
      showInKpis: false,
      defaultSort: { field: 'modifiedon', desc: true },
      columns: [
        { field: 'admin_displayname', label: 'Desktop Flow', type: 'title' },
        { field: '_admin_rpaowner_value', label: 'Besitzer', type: 'person' },
        { field: '_admin_rpaenvironment_value', label: 'Umgebung', type: 'text' },
        { field: 'modifiedon', label: 'Geändert', type: 'date' }
      ],
      links: []
    },
    {
      key: 'powerpages',
      title: 'Power Pages',
      logicalName: 'admin_powerappsportal',
      icon: 'Website',
      showInKpis: false,
      defaultSort: { field: 'modifiedon', desc: true },
      columns: [
        { field: 'admin_displayname', label: 'Website', type: 'title' },
        { field: 'admin_websiteurl', label: 'URL', type: 'url' },
        { field: '_admin_portalenvironment_value', label: 'Umgebung', type: 'text' },
        { field: 'modifiedon', label: 'Geändert', type: 'date' }
      ],
      links: []
    }
  ];
  return getDefaultTables().concat(extra);
}

/** Leere Konfiguration für eine beliebige Dataverse-Tabelle (Spalten werden nach dem Laden der Metadaten ergänzt). */
export function newTableConfig(logicalName: string, displayName: string): ITableConfig {
  return {
    key: logicalName.replace(/[^A-Za-z0-9_-]/g, '') || `t${Date.now().toString(36)}`,
    title: displayName || logicalName,
    logicalName,
    icon: 'Table',
    showInKpis: true,
    columns: [],
    links: []
  };
}
