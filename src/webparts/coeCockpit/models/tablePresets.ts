import { ITableConfig } from './ITableConfig';
import { getDefaultTables } from './defaultTables';

/**
 * Vorkonfigurierte Tabellen des CoE Starter Kits zur Auswahl im Tabellen-Editor.
 * Enthält alle Tabellen der CoE-App „Power Platform Admin View“.
 * Angeboten wird eine Vorlage nur, wenn es die Tabelle in der Umgebung gibt; Spalten, die es in
 * der eigenen CoE-Version nicht gibt, werden beim Hinzufügen automatisch entfernt.
 */
export function getTablePresets(): ITableConfig[] {
  // Alle Tabellen der CoE Admin View (siehe defaultTables.ts)
  return getDefaultTables();
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
