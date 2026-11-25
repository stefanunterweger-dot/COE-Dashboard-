/** Darstellungsart einer Spalte. */
export type ColumnType =
  | 'title'
  | 'text'
  | 'person'
  | 'date'
  | 'datetime'
  | 'number'
  | 'boolean'
  | 'choice'
  | 'status'
  | 'url';

/** Farbton für Status-Pillen und Kennzahlen. */
export type Tone = 'success' | 'warning' | 'danger' | 'info' | 'neutral' | 'accent';

export interface IColumnConfig {
  /** Feldname für $select, z. B. admin_displayname oder _admin_appowner_value (Lookup). */
  field: string;
  /** Spaltenüberschrift. */
  label: string;
  type: ColumnType;
  /** Standardmäßig ausgeblendet (über "Spalten" einblendbar). */
  hidden?: boolean;
  /** Nur für type = status: Wert → Farbton, z. B. { "Started": "success" }. */
  statusMap?: { [value: string]: Tone };
}

export interface ILinkConfig {
  label: string;
  /** URL-Vorlage, Platzhalter in geschweiften Klammern: {admin_appid} */
  url: string;
}

export interface IKpiConfig {
  /** Text hinter der Zahl, z. B. "verwaist". */
  highlightLabel?: string;
  /** OData-Filter für die hervorgehobene Zahl. */
  highlightFilter?: string;
  tone?: Tone;
}

export interface ISortConfig {
  field: string;
  desc: boolean;
}

export interface ITableConfig {
  /** Eindeutiger Schlüssel innerhalb des Webparts. */
  key: string;
  /** Anzeigename (Reiter). */
  title: string;
  /** Logischer Tabellenname in Dataverse, z. B. admin_app. */
  logicalName: string;
  /** Fluent-UI-Iconname, z. B. AppIconDefault. */
  icon: string;
  /** Optionaler OData-Basisfilter, z. B. admin_appdeleted ne true. */
  filter?: string;
  showInKpis?: boolean;
  defaultSort?: ISortConfig;
  kpi?: IKpiConfig;
  columns: IColumnConfig[];
  links?: ILinkConfig[];
  /** Optional: Tabelle aus einer anderen Dataverse-Umgebung lesen (Standard: CoE-Umgebung). */
  environmentUrl?: string;
  /** Intern: Daten werden berechnet statt direkt aus einer Tabelle geladen (z. B. „Betroffen ab 15.12.“). */
  virtual?: boolean;
  /** Intern: Filter darf nie weggelassen werden (begrenzt den Bereich, z. B. auf das Default Environment). */
  strictFilter?: boolean;
  /** Intern: Grund, warum die Tabelle nicht geladen werden kann (wird statt der Daten angezeigt). */
  unavailable?: string;
}

export interface IRecord {
  [field: string]: unknown;
}
