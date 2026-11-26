import { IActionSetupState } from '../services/ActionService';

/** Alle Webpart-Eigenschaften, die im Bereich „Webpart-Einstellungen“ bearbeitet werden. */
export interface ISettingsValues {
  title: string;
  subtitle: string;
  environmentUrl: string;
  showOverview: boolean;
  showKpis: boolean;
  density: 'comfortable' | 'compact';
  pageSize: number;
  maxRecords: number;
  showDefaultEnvironment: boolean;
  defaultEnvironmentId: string;
  showDevEnvironments: boolean;
  devEnvManagedField: string;
  showSecurity: boolean;
  showCopilotCredits: boolean;
  pipelinesHostUrl: string;
  notifyAutoSend: boolean;
  notifyTone: 'sie' | 'du';
  notifyReplyDays: number;
  notifyInfoUrl: string;
  notifyFeedbackUrl: string;
  notifyCc: string;
  notifySignature: string;
  showActions: boolean;
  actionsGroup: string;
  deleteRequiresStop: boolean;
}

export interface ISettingsOption { key: string | number; text: string }

/** Was das Webpart dem Einstellungsbereich mitgibt (Werte, Auswahllisten, Stand, Rückrufe). */
export interface IWebPartSettings {
  values: ISettingsValues;
  /** Auswahl „Default Environment“ (inkl. „automatisch“) */
  envOptions: ISettingsOption[];
  envStatus: string;
  /** Auswahl der Managed-Spalte (inkl. „automatisch“) */
  managedOptions: ISettingsOption[];
  managedInfo: string;
  /** Auswahl Pipelines-Host (inkl. „keine“) */
  hostOptions: ISettingsOption[];
  /** SharePoint-Gruppen der Website */
  groupOptions: ISettingsOption[];
  actionSetup?: IActionSetupState;
  actionBusy: string;
  actionMessage: string;
  tableCount: number;
  defaultInfoUrl: string;
  /** Werte übernehmen (wie eine Änderung im Eigenschaftenbereich). */
  onChange: (patch: Partial<ISettingsValues>) => void;
  /** Auswahllisten und Einrichtungsstand laden (beim Öffnen). */
  onLoad: () => void;
  onActionSetup: (kind: 'group' | 'list') => void;
  onOpenTableEditor: () => void;
  /** Klassischen SharePoint-Eigenschaftenbereich öffnen. */
  onOpenPropertyPane: () => void;
}
