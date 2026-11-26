import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneGroup,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';

import * as strings from 'CoeCockpitWebPartStrings';
import CoeCockpit, { ICoeCockpitProps } from './components/CoeCockpit';
import { ITableConfig } from './models/ITableConfig';
import { getDefaultTables } from './models/defaultTables';
import { ISettingsValues, IWebPartSettings } from './models/settings';
import { IMailTemplateOverrides } from './models/mailTemplates';
import { isValidEnvironmentUrl } from './utils/format';
import { DataverseService, normalizeDataverseUrl } from './services/DataverseService';
import { detectManagedAttribute, managedCandidates } from './models/devEnvironments';
import { DEFAULT_INFO_URL, OLD_INFO_URLS, parseAddresses } from './models/notification';
import { ACTION_LIST_TITLE, DEFAULT_ACTION_GROUP } from './models/actions';
import { ActionService, IActionSetupState, ISiteGroup } from './services/ActionService';

export interface ICoeCockpitWebPartProps {
  title: string;
  subtitle: string;
  environmentUrl: string;
  tables: ITableConfig[];
  pageSize: number;
  maxRecords: number;
  showKpis: boolean;
  density: 'compact' | 'comfortable';
  showDefaultEnvironment: boolean;
  /** Leer = automatisch erkennen (Umgebungstyp „Default“). */
  defaultEnvironmentId: string;
  /** Optional, sonst aus dem CoE-Inventar. */
  defaultEnvironmentDataverseUrl: string;
  /** Reiter „Betroffen ab 15.12.“ anzeigen. */
  showDevEnvironments: boolean;
  /** Übersicht (Dashboard) als Startseite. */
  showOverview: boolean;
  /** Spalte mit dem Managed-Status (leer = automatisch erkennen). */
  devEnvManagedField: string;
  /** Ansicht „Sicherheit & Freigaben“ anzeigen. */
  showSecurity: boolean;
  showCopilotCredits: boolean;
  /** Auswertung „Betroffen ab Februar 2027“ (Premium-Lizenz in Managed Environments). */
  showPremium2027: boolean;
  /** Dataverse-URL des Pipelines-Host-Environments (leer = ALM-Pipelines ausblenden). */
  pipelinesHostUrl: string;
  /** E-Mail an Maker: Anrede „sie“ oder „du“. */
  notifyTone: 'sie' | 'du';
  /** Frist für die Rückmeldung in Tagen. */
  notifyReplyDays: number;
  /** Link zur offiziellen Microsoft-Information. */
  notifyInfoUrl: string;
  /** Optional: Formular für die Rückmeldung. */
  notifyFeedbackUrl: string;
  /** Optional: CC-Adresse(n). */
  notifyCc: string;
  /** Grußzeile. */
  notifySignature: string;
  /** Automatischen Versand über Microsoft Graph anbieten. */
  notifyAutoSend: boolean;
  /** Gepflegte Mail-Texte (nur Abweichungen vom Standard; Webpart-Einstellungen → E-Mail & Vorlagen). */
  mailTemplates: IMailTemplateOverrides;
  /** Aktionen für Apps und Flows anbieten. */
  showActions: boolean;
  /** SharePoint-Gruppe, deren Mitglieder Aktionen auslösen dürfen. */
  actionsGroup: string;
  /** Löschen erst nach Quarantäne (Apps) bzw. Ausschalten (Flows). */
  deleteRequiresStop: boolean;
}

export default class CoeCockpitWebPart extends BaseClientSideWebPart<ICoeCockpitWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _theme: IReadonlyTheme | undefined;
  private _editorRequest: number = 0;
  private _settingsRequest: number = 0;
  private _envOptions: IPropertyPaneDropdownOption[] | undefined;
  private _hostOptions: IPropertyPaneDropdownOption[] | undefined;
  private _groups: ISiteGroup[] | undefined;
  private _actionSetup: IActionSetupState | undefined;
  private _actionBusy: string = '';
  private _actionMessage: string = '';
  private _actionsVersion: number = 0;
  private _envLoadError: string | undefined;
  private _envLoadedFor: string | undefined;
  private _managedOptions: IPropertyPaneDropdownOption[] | undefined;
  private _managedDetected: string | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    if (!this.properties.tables || this.properties.tables.length === 0) {
      this.properties.tables = getDefaultTables();
    }
    // Früheren Standard-Link (allgemeine deutsche Seite) auf den direkten Abschnitt umstellen
    if (OLD_INFO_URLS.indexOf((this.properties.notifyInfoUrl || '').trim()) >= 0) {
      this.properties.notifyInfoUrl = '';
    }

    const themeProvider: ThemeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeProvider = themeProvider;
    this._theme = themeProvider.tryGetTheme();
    themeProvider.themeChangedEvent.add(this, this._onThemeChanged);
  }

  public render(): void {
    const element: React.ReactElement<ICoeCockpitProps> = React.createElement(CoeCockpit, {
      instanceId: this.context.instanceId,
      title: this.properties.title,
      subtitle: this.properties.subtitle,
      environmentUrl: (this.properties.environmentUrl || '').trim(),
      tables: this.properties.tables || [],
      pageSize: this.properties.pageSize || 25,
      maxRecords: this.properties.maxRecords || 20000,
      showKpis: this.properties.showKpis !== false,
      density: this.properties.density || 'comfortable',
      showDefaultEnvironment: this.properties.showDefaultEnvironment !== false,
      defaultEnvironmentId: (this.properties.defaultEnvironmentId || '').trim(),
      defaultEnvironmentDataverseUrl: normalizeDataverseUrl(this.properties.defaultEnvironmentDataverseUrl) || '',
      showDevEnvironments: this.properties.showDevEnvironments !== false,
      showOverview: this.properties.showOverview !== false,
      devEnvManagedField: (this.properties.devEnvManagedField || '').trim(),
      showSecurity: this.properties.showSecurity !== false,
      showCopilotCredits: this.properties.showCopilotCredits !== false,
      showPremium2027: this.properties.showPremium2027 !== false,
      pipelinesHostUrl: normalizeDataverseUrl(this.properties.pipelinesHostUrl) || '',
      notify: {
        formal: this.properties.notifyTone !== 'du',
        replyDays: this.properties.notifyReplyDays || 14,
        infoUrl: this._infoUrl(),
        feedbackUrl: /^https:\/\//i.test((this.properties.notifyFeedbackUrl || '').trim()) ? this.properties.notifyFeedbackUrl.trim() : '',
        cc: (this.properties.notifyCc || '').trim(),
        signature: (this.properties.notifySignature || '').trim(),
        templates: this.properties.mailTemplates || {}
      },
      notifyAutoSend: this.properties.notifyAutoSend !== false,
      showActions: this.properties.showActions !== false,
      actionsGroup: (this.properties.actionsGroup || DEFAULT_ACTION_GROUP).trim(),
      deleteRequiresStop: this.properties.deleteRequiresStop !== false,
      actionsVersion: this._actionsVersion,
      theme: this._theme,
      isEditMode: this.displayMode === DisplayMode.Edit,
      editorRequest: this._editorRequest,
      aadHttpClientFactory: this.context.aadHttpClientFactory,
      aadTokenProviderFactory: this.context.aadTokenProviderFactory,
      msGraphClientFactory: this.context.msGraphClientFactory,
      spHttpClient: this.context.spHttpClient,
      webUrl: this.context.pageContext.web.absoluteUrl,
      currentUser: { displayName: this.context.pageContext.user.displayName, email: this.context.pageContext.user.email },
      onTablesChanged: (tables: ITableConfig[]) => {
        this.properties.tables = tables;
        this.render();
        this.context.propertyPane.refresh();
      },
      onConfigure: () => this.context.propertyPane.open(),
      settings: this.displayMode === DisplayMode.Edit ? this._settings() : undefined,
      settingsRequest: this._settingsRequest
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    if (this._themeProvider) {
      this._themeProvider.themeChangedEvent.remove(this, this._onThemeChanged);
    }
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _onThemeChanged(args: ThemeChangedEventArgs): void {
    this._theme = args.theme;
    this.render();
  }

  /** Werte und Rückrufe für den Bereich „Webpart-Einstellungen“. */
  private _settings(): IWebPartSettings {
    const p = this.properties;
    return {
      values: {
        title: p.title || '',
        subtitle: p.subtitle || '',
        environmentUrl: (p.environmentUrl || '').trim(),
        showOverview: p.showOverview !== false,
        showKpis: p.showKpis !== false,
        density: p.density === 'compact' ? 'compact' : 'comfortable',
        pageSize: p.pageSize || 25,
        maxRecords: p.maxRecords || 20000,
        showDefaultEnvironment: p.showDefaultEnvironment !== false,
        defaultEnvironmentId: (p.defaultEnvironmentId || '').trim(),
        showDevEnvironments: p.showDevEnvironments !== false,
        devEnvManagedField: (p.devEnvManagedField || '').trim(),
        showSecurity: p.showSecurity !== false,
        showCopilotCredits: p.showCopilotCredits !== false,
        showPremium2027: p.showPremium2027 !== false,
        pipelinesHostUrl: normalizeDataverseUrl(p.pipelinesHostUrl) || '',
        notifyAutoSend: p.notifyAutoSend !== false,
        notifyTone: p.notifyTone === 'du' ? 'du' : 'sie',
        notifyReplyDays: p.notifyReplyDays || 14,
        notifyInfoUrl: p.notifyInfoUrl || '',
        notifyFeedbackUrl: p.notifyFeedbackUrl || '',
        notifyCc: p.notifyCc || '',
        notifySignature: p.notifySignature || '',
        mailTemplates: p.mailTemplates || {},
        showActions: p.showActions !== false,
        actionsGroup: this._groupName,
        deleteRequiresStop: p.deleteRequiresStop !== false
      },
      envOptions: this._environmentDropdownOptions().map(o => ({ key: o.key, text: o.text })),
      envStatus: this._environmentStatusText(),
      managedOptions: this._managedFieldOptions().map(o => ({ key: o.key, text: o.text })),
      managedInfo: this._managedInfoText(),
      hostOptions: this._pipelineHostOptions().map(o => ({ key: o.key, text: o.text })),
      groupOptions: this._actionGroupOptions().map(o => ({ key: o.key, text: o.text })),
      actionSetup: this._actionSetup,
      actionBusy: this._actionBusy,
      actionMessage: this._actionMessage,
      tableCount: (p.tables || []).length,
      defaultInfoUrl: DEFAULT_INFO_URL,
      onChange: patch => this._applySettings(patch),
      onLoad: () => { this._loadEnvironmentOptions(); this._loadActionSetup(); },
      onActionSetup: kind => this._runActionSetup(kind),
      onOpenTableEditor: () => this._openEditor(),
      onOpenPropertyPane: () => this.context.propertyPane.open()
    };
  }

  /** Änderungen aus den Webpart-Einstellungen übernehmen – wie eine Änderung im Eigenschaftenbereich. */
  private _applySettings(patch: Partial<ISettingsValues>): void {
    const props = this.properties as unknown as { [key: string]: unknown };
    const values = patch as unknown as { [key: string]: unknown };
    Object.keys(values).forEach(key => {
      const oldValue = props[key];
      const newValue = values[key];
      if (oldValue === newValue) {
        return;
      }
      props[key] = newValue;
      this.onPropertyPaneFieldChanged(key, oldValue, newValue);
    });
    this.render();
    this.context.propertyPane.refresh();
  }

  private _managedInfoText(): string {
    return !isValidEnvironmentUrl(this.properties.environmentUrl)
      ? 'Zuerst die URL der CoE-Umgebung eintragen.'
      : !this._managedOptions
        ? 'Spalten werden aus dem CoE-Inventar geladen …'
        : this._managedDetected
          ? 'Angezeigt werden Developer-Environments, die laut dieser Spalte nicht managed sind.'
          : 'Keine Managed-Spalte erkannt – ohne Auswahl werden alle Developer-Environments angezeigt.';
  }

  /** Übersichtliche Webpart-Einstellungen öffnen (aus dem SharePoint-Eigenschaftenbereich heraus). */
  private _openSettings(): void {
    this._settingsRequest++;
    this.context.propertyPane.close();
    this.render();
  }

  private _openEditor(): void {
    this._editorRequest++;
    this.context.propertyPane.close();
    this.render();
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadEnvironmentOptions();
    this._loadActionSetup();
  }

  private get _actionService(): ActionService {
    return new ActionService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
  }

  private get _groupName(): string {
    return (this.properties.actionsGroup || DEFAULT_ACTION_GROUP).trim();
  }

  /** Gruppen der Website und Stand der Einrichtung laden (für die Eigenschaften). */
  private _loadActionSetup(): void {
    const svc = this._actionService;
    Promise.all([svc.listGroups().catch(() => [] as ISiteGroup[]), svc.getSetupState(this._groupName).catch(() => undefined)]).then(([groups, st]) => {
      this._groups = groups;
      this._actionSetup = st;
      this.context.propertyPane.refresh();
      this.render();
    });
  }

  /** Gruppe anlegen bzw. Liste einrichten – mit Rückmeldung im Eigenschaftenbereich. */
  private _runActionSetup(kind: 'group' | 'list'): void {
    if (this._actionBusy) {
      return;
    }
    const svc = this._actionService;
    const name = this._groupName;
    this._actionBusy = kind === 'group' ? `Gruppe „${name}“ wird angelegt …` : `Liste „${ACTION_LIST_TITLE}“ wird eingerichtet …`;
    this._actionMessage = '';
    this.context.propertyPane.refresh();
    this.render();
    const job = kind === 'group'
      ? svc.createGroup(name).then(() => { this.properties.actionsGroup = name; return `Gruppe „${name}“ ist angelegt. Du bist erstes Mitglied – weitere Admins über die Websiteberechtigungen hinzufügen.`; })
      : svc.setupList(name).then(() => `Liste „${ACTION_LIST_TITLE}“ ist eingerichtet: nur „${name}“ (Bearbeiten) und die Websitebesitzer (Vollzugriff) haben Zugriff.`);
    job.then(
      msg => { this._actionMessage = msg; },
      (e: unknown) => { this._actionMessage = `Fehler: ${e instanceof Error ? e.message : String(e)}`; }
    ).then(() => {
      this._actionBusy = '';
      this._actionsVersion++;
      this._loadActionSetup();
      this.render();
    });
  }

  private _actionGroupOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = (this._groups || []).map(g => ({ key: g.title, text: g.title }));
    const current = this._groupName;
    if (!options.some(o => String(o.key).toLowerCase() === current.toLowerCase())) {
      options.unshift({ key: current, text: `${current} (noch nicht angelegt)` });
    }
    return options;
  }

  private _actionStatusText(): string {
    if (this._actionBusy) {
      return this._actionBusy;
    }
    const st = this._actionSetup;
    const lines: string[] = [];
    if (!st) {
      lines.push('Stand wird geprüft …');
    } else {
      lines.push(st.group ? `✓ Gruppe „${st.group.title}“ vorhanden` : `✗ Gruppe „${this._groupName}“ fehlt`);
      lines.push(st.listExists
        ? (st.listUnique ? `✓ Liste „${ACTION_LIST_TITLE}“ eingerichtet (eigene Berechtigungen)` : `⚠ Liste vorhanden, erbt aber die Berechtigungen der Website – „Aktionsliste einrichten“ klicken`)
        : `✗ Liste „${ACTION_LIST_TITLE}“ fehlt`);
      lines.push(st.isMember ? '✓ Du bist Mitglied der Gruppe' : '✗ Du bist nicht Mitglied der Gruppe – Aktionen sind für dich ausgeblendet');
    }
    if (this._actionMessage) {
      lines.push(this._actionMessage);
    }
    return lines.join(' · ');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'actionsGroup') {
      this._actionMessage = '';
      this._actionSetup = undefined;
      this._loadActionSetup();
    }
    if (propertyPath === 'environmentUrl') {
      this._envLoadedFor = undefined;
      this._loadEnvironmentOptions();
    }
  }

  /** Lädt die Umgebungen aus dem CoE-Inventar für die Auswahl „Default Environment“. */
  private _loadEnvironmentOptions(): void {
    const url = (this.properties.environmentUrl || '').trim();
    if (!isValidEnvironmentUrl(url) || this._envLoadedFor === url) {
      return;
    }
    this._envLoadedFor = url;
    this._envOptions = undefined;
    this._envLoadError = undefined;
    this._managedOptions = undefined;
    const service = new DataverseService(this.context.aadHttpClientFactory, url);
    service.getAttributes('admin_environment').then(
      attrs => {
        if (this._envLoadedFor !== url) {
          return;
        }
        const detected = detectManagedAttribute(attrs);
        this._managedDetected = detected ? `${detected.displayName} (${detected.selectName})` : undefined;
        this._managedOptions = managedCandidates(attrs).map(a => ({ key: a.selectName, text: `${a.displayName} (${a.selectName})` }));
        this.context.propertyPane.refresh();
        this.render();
      },
      () => { this._managedOptions = []; this.context.propertyPane.refresh(); }
    );
    service.getEnvironments().then(
      envs => {
        if (this._envLoadedFor !== url) {
          return;
        }
        this._envOptions = envs.map(e => ({
          key: e.id,
          text: `${e.name}${e.type ? ` (${e.type})` : ''}`
        }));
        // Pipelines-Host: nur Environments mit Dataverse
        this._hostOptions = envs.filter(e => !!e.dataverseUrl).map(e => ({
          key: e.dataverseUrl as string,
          text: `${e.name}${e.type ? ` (${e.type})` : ''}`
        }));
        this.context.propertyPane.refresh();
        this.render();
      },
      (e: Error) => {
        if (this._envLoadedFor !== url) {
          return;
        }
        this._envLoadError = e.message;
        this._envLoadedFor = undefined;
        this.context.propertyPane.refresh();
      }
    );
  }

  private _environmentDropdownOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [{ key: '', text: strings.DefaultEnvAuto }];
    const current = (this.properties.defaultEnvironmentId || '').trim();
    (this._envOptions || []).forEach(o => options.push(o));
    if (current && !(this._envOptions || []).some(o => String(o.key).toLowerCase() === current.toLowerCase())) {
      options.push({ key: current, text: current });
    }
    return options;
  }

  /** Link zur Microsoft-Ankündigung; leere, ungültige und frühere Standardwerte → aktueller Standard. */
  private _infoUrl(): string {
    const v = (this.properties.notifyInfoUrl || '').trim();
    return /^https:\/\/\S+$/i.test(v) && OLD_INFO_URLS.indexOf(v) < 0 ? v : DEFAULT_INFO_URL;
  }

  /** Auswahl des Pipelines-Host-Environments (Schlüssel = Dataverse-URL). */
  private _pipelineHostOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [{ key: '', text: 'Keine – ALM-Pipelines ausblenden' }];
    const current = normalizeDataverseUrl(this.properties.pipelinesHostUrl) || '';
    (this._hostOptions || []).forEach(o => options.push(o));
    if (current && !(this._hostOptions || []).some(o => String(o.key).toLowerCase() === current.toLowerCase())) {
      options.push({ key: current, text: current });
    }
    return options;
  }

  /** Auswahl der Managed-Spalte: automatisch erkennen oder eine Spalte der CoE-Environment-Tabelle. */
  private _managedFieldOptions(): IPropertyPaneDropdownOption[] {
    const options: IPropertyPaneDropdownOption[] = [
      { key: '', text: this._managedDetected ? `Automatisch: ${this._managedDetected}` : 'Automatisch erkennen' }
    ];
    (this._managedOptions || []).forEach(o => options.push(o));
    const current = (this.properties.devEnvManagedField || '').trim();
    if (current && !(this._managedOptions || []).some(o => o.key === current)) {
      options.push({ key: current, text: current });
    }
    return options;
  }

  private _environmentStatusText(): string {
    if (!isValidEnvironmentUrl(this.properties.environmentUrl)) {
      return strings.DefaultEnvNeedsUrl;
    }
    if (this._envLoadError) {
      return `${strings.DefaultEnvLoadError} ${this._envLoadError}`;
    }
    if (!this._envOptions) {
      return strings.DefaultEnvLoading;
    }
    return strings.DefaultEnvCount.replace('{0}', String(this._envOptions.length));
  }

  private _validateUrl(value: string): string {
    if (!value || !value.trim()) {
      return '';
    }
    return isValidEnvironmentUrl(value) ? '' : strings.EnvironmentUrlError;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const config = this._propertyPaneConfiguration();
    // Übersichtlicher: oben die Schaltfläche zu den Webpart-Einstellungen, alle Abschnitte außer „Verbindung“ zugeklappt
    const page = config.pages[0];
    page.groups = [{
      groupName: 'Webpart-Einstellungen',
      groupFields: [
        PropertyPaneLabel('settingsHint', { text: 'Alle Einstellungen übersichtlich nach Bereichen – mit Status der Einrichtung.' }),
        PropertyPaneButton('openSettings', {
          text: 'Webpart-Einstellungen öffnen',
          buttonType: PropertyPaneButtonType.Primary,
          icon: 'Settings',
          onClick: () => { this._openSettings(); return undefined; }
        })
      ]
    } as IPropertyPaneGroup].concat((page.groups as IPropertyPaneGroup[]).map((g, i) => (i === 0 ? g : { ...g, isCollapsed: true })));
    return config;
  }

  private _propertyPaneConfiguration(): IPropertyPaneConfiguration {
    const tableCount = (this.properties.tables || []).length;
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ConnectionGroupName,
              groupFields: [
                PropertyPaneTextField('environmentUrl', {
                  label: strings.EnvironmentUrlLabel,
                  description: strings.EnvironmentUrlDescription,
                  placeholder: 'https://contoso-coe.crm4.dynamics.com',
                  onGetErrorMessage: this._validateUrl.bind(this),
                  deferredValidationTime: 600
                })
              ]
            },
            {
              groupName: strings.DefaultEnvGroupName,
              groupFields: [
                PropertyPaneToggle('showDefaultEnvironment', {
                  label: strings.DefaultEnvShowLabel,
                  checked: this.properties.showDefaultEnvironment !== false
                }),
                PropertyPaneDropdown('defaultEnvironmentId', {
                  label: strings.DefaultEnvLabel,
                  options: this._environmentDropdownOptions(),
                  selectedKey: (this.properties.defaultEnvironmentId || '').trim(),
                  disabled: this.properties.showDefaultEnvironment === false
                }),
                PropertyPaneLabel('defaultEnvironmentStatus', {
                  text: this._environmentStatusText()
                })
              ]
            },
            {
              groupName: 'Betroffen ab 15.12. (Developer-Environments)',
              groupFields: [
                PropertyPaneToggle('showDevEnvironments', {
                  label: 'Reiter „Betroffen ab 15.12.“ anzeigen',
                  checked: this.properties.showDevEnvironments !== false
                }),
                PropertyPaneDropdown('devEnvManagedField', {
                  label: 'Spalte mit dem Managed-Status',
                  options: this._managedFieldOptions(),
                  selectedKey: this.properties.devEnvManagedField || '',
                  disabled: this.properties.showDevEnvironments === false
                }),
                PropertyPaneLabel('devEnvInfo', {
                  text: this._managedInfoText()
                })
              ]
            },
            {
              groupName: 'E-Mail an Maker (Betroffen ab 15.12.)',
              groupFields: [
                PropertyPaneToggle('notifyAutoSend', {
                  label: 'Automatischen Versand anbieten (Microsoft Graph)',
                  checked: this.properties.notifyAutoSend !== false
                }),
                PropertyPaneLabel('notifyAutoSendInfo', {
                  text: 'Sendet aus dem Postfach des angemeldeten Benutzers. Voraussetzung: Die API-Berechtigung „Microsoft Graph – Mail.Send“ ist im SharePoint Admin Center genehmigt.'
                }),
                PropertyPaneDropdown('notifyTone', {
                  label: 'Anrede',
                  options: [{ key: 'sie', text: 'Sie' }, { key: 'du', text: 'Du' }],
                  selectedKey: this.properties.notifyTone === 'du' ? 'du' : 'sie'
                }),
                PropertyPaneDropdown('notifyReplyDays', {
                  label: 'Frist für die Rückmeldung',
                  options: [7, 10, 14, 21].map(n => ({ key: n, text: `${n} Tage` })),
                  selectedKey: this.properties.notifyReplyDays || 14
                }),
                PropertyPaneTextField('notifyInfoUrl', {
                  label: 'Link zur Microsoft-Information',
                  description: 'Offizielle Microsoft-Seite zur Änderung (wird in der E-Mail verlinkt).',
                  placeholder: DEFAULT_INFO_URL,
                  onGetErrorMessage: (v: string) => (!v || /^https:\/\/\S+$/i.test(v.trim()) ? '' : 'Bitte einen https-Link eintragen.')
                }),
                PropertyPaneTextField('notifyFeedbackUrl', {
                  label: 'Formular für die Rückmeldung (optional)',
                  description: 'z. B. Microsoft Forms. {environment} im Link wird durch den Namen des Environments ersetzt. Ohne Formular antworten die Maker per E-Mail.',
                  onGetErrorMessage: (v: string) => (!v || /^https:\/\/\S+$/i.test(v.trim()) ? '' : 'Bitte einen https-Link eintragen.')
                }),
                PropertyPaneTextField('notifyCc', {
                  label: 'CC (optional)',
                  description: 'z. B. das Postfach des Power-Platform-Teams, damit Rückmeldungen dort ankommen.',
                  onGetErrorMessage: (v: string) => (!v || parseAddresses(v).length === v.split(/[;,\s]+/).filter(x => !!x).length ? '' : 'Bitte gültige E-Mail-Adressen eintragen (getrennt durch ;).')
                }),
                PropertyPaneTextField('notifySignature', {
                  label: 'Grußzeile',
                  placeholder: 'Ihr Power-Platform-Team'
                })
              ]
            },
            {
              groupName: 'Aktionen für Apps und Flows',
              groupFields: [
                PropertyPaneToggle('showActions', {
                  label: 'Aktionen anbieten (Quarantäne, Löschen, Besitzer, Freigaben, Flows ein/aus)',
                  checked: this.properties.showActions !== false
                }),
                PropertyPaneDropdown('actionsGroup', {
                  label: 'SharePoint-Gruppe mit Berechtigung',
                  options: this._actionGroupOptions(),
                  selectedKey: this._groupName,
                  disabled: this.properties.showActions === false
                }),
                PropertyPaneButton('createActionGroup', {
                  text: `Gruppe „${this._groupName}“ anlegen`,
                  buttonType: PropertyPaneButtonType.Normal,
                  icon: 'AddGroup',
                  disabled: this.properties.showActions === false || !!this._actionBusy || !!(this._actionSetup && this._actionSetup.group),
                  onClick: () => { this._runActionSetup('group'); return undefined; }
                }),
                PropertyPaneButton('setupActionList', {
                  text: 'Aktionsliste einrichten',
                  buttonType: PropertyPaneButtonType.Normal,
                  icon: 'Lock',
                  disabled: this.properties.showActions === false || !!this._actionBusy || !(this._actionSetup && this._actionSetup.group),
                  onClick: () => { this._runActionSetup('list'); return undefined; }
                }),
                PropertyPaneLabel('actionStatus', { text: this._actionStatusText() }),
                PropertyPaneToggle('deleteRequiresStop', {
                  label: 'Löschen erst nach Quarantäne (Apps) bzw. Ausschalten (Flows)',
                  checked: this.properties.deleteRequiresStop !== false,
                  disabled: this.properties.showActions === false
                }),
                PropertyPaneLabel('actionInfo', {
                  text: 'Sichtbar nur für Mitglieder der Gruppe mit aktiver Rolle „Power Platform Administrator“. Das Cockpit führt die Aktionen direkt aus und protokolliert sie in der Liste; nur „Solution löschen“ übernimmt der Flow „CoE Cockpit | Aktion ausführen“ (dort denselben Gruppennamen eintragen).'
                })
              ]
            },
            {
              groupName: 'Lizenzen & Kapazität (live)',
              groupFields: [
                PropertyPaneToggle('showCopilotCredits', {
                  label: 'Ansichten „Copilot-Credits“ und „Abrechnungspläne“ anzeigen',
                  checked: this.properties.showCopilotCredits !== false
                }),
                PropertyPaneLabel('copilotCreditsInfo', {
                  text: 'Die Daten kommen live aus der Power Platform API – ohne CoE und Dataverse. Voraussetzungen: Ein Admin hat die Power Platform API für SharePoint freigegeben '
                    + '(Skript im README), und wer die Ansicht öffnet, hat die Rolle „Power Platform Administrator“ aktiv (bei PIM vorher aktivieren).'
                })
              ]
            },
            {
              groupName: 'Sicherheit & ALM-Pipelines',
              groupFields: [
                PropertyPaneToggle('showSecurity', {
                  label: 'Ansicht „Sicherheit & Freigaben“ anzeigen',
                  checked: this.properties.showSecurity !== false
                }),
                PropertyPaneToggle('showPremium2027', {
                  label: 'Auswertung „Betroffen ab Februar 2027“ anzeigen',
                  checked: this.properties.showPremium2027 !== false
                }),
                PropertyPaneDropdown('pipelinesHostUrl', {
                  label: 'Pipelines-Host-Environment',
                  options: this._pipelineHostOptions(),
                  selectedKey: normalizeDataverseUrl(this.properties.pipelinesHostUrl) || ''
                }),
                PropertyPaneLabel('pipelinesInfo', {
                  text: 'Das Environment, in dem die App „Deployment Pipeline Configuration“ läuft. Die Bereitstellungen werden dort mit deinen eigenen Rechten gelesen – ohne Zugriff weist das Cockpit darauf hin.'
                })
              ]
            },
            {
              groupName: strings.TablesGroupName,
              groupFields: [
                PropertyPaneLabel('tablesSummary', {
                  text: strings.TablesSummary.replace('{0}', String(tableCount))
                }),
                PropertyPaneButton('openTableEditor', {
                  text: strings.OpenEditorLabel,
                  buttonType: PropertyPaneButtonType.Primary,
                  icon: 'Settings',
                  onClick: () => {
                    this._openEditor();
                    return undefined;
                  }
                })
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleLabel }),
                PropertyPaneTextField('subtitle', { label: strings.SubtitleLabel }),
                PropertyPaneToggle('showOverview', { label: 'Übersicht (Dashboard) als Startseite anzeigen', checked: this.properties.showOverview !== false }),
                PropertyPaneToggle('showKpis', { label: strings.ShowKpisLabel }),
                PropertyPaneDropdown('density', {
                  label: strings.DensityLabel,
                  options: [
                    { key: 'comfortable', text: strings.DensityComfortable },
                    { key: 'compact', text: strings.DensityCompact }
                  ]
                }),
                PropertyPaneDropdown('pageSize', {
                  label: strings.PageSizeLabel,
                  options: [10, 25, 50, 100].map(n => ({ key: n, text: String(n) }))
                }),
                PropertyPaneSlider('maxRecords', {
                  label: strings.MaxRecordsLabel,
                  min: 1000,
                  max: 50000,
                  step: 1000,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
