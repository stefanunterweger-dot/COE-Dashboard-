import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
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
import { isValidEnvironmentUrl } from './utils/format';
import { DataverseService, normalizeDataverseUrl } from './services/DataverseService';
import { detectManagedAttribute, managedCandidates } from './models/devEnvironments';

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
  /** Spalte mit dem Managed-Status (leer = automatisch erkennen). */
  devEnvManagedField: string;
}

export default class CoeCockpitWebPart extends BaseClientSideWebPart<ICoeCockpitWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _theme: IReadonlyTheme | undefined;
  private _editorRequest: number = 0;
  private _envOptions: IPropertyPaneDropdownOption[] | undefined;
  private _envLoadError: string | undefined;
  private _envLoadedFor: string | undefined;
  private _managedOptions: IPropertyPaneDropdownOption[] | undefined;
  private _managedDetected: string | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    if (!this.properties.tables || this.properties.tables.length === 0) {
      this.properties.tables = getDefaultTables();
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
      devEnvManagedField: (this.properties.devEnvManagedField || '').trim(),
      theme: this._theme,
      isEditMode: this.displayMode === DisplayMode.Edit,
      editorRequest: this._editorRequest,
      aadHttpClientFactory: this.context.aadHttpClientFactory,
      onTablesChanged: (tables: ITableConfig[]) => {
        this.properties.tables = tables;
        this.render();
        this.context.propertyPane.refresh();
      },
      onConfigure: () => this.context.propertyPane.open()
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

  private _openEditor(): void {
    this._editorRequest++;
    this.context.propertyPane.close();
    this.render();
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadEnvironmentOptions();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
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
        this.context.propertyPane.refresh();
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
                  text: !isValidEnvironmentUrl(this.properties.environmentUrl)
                    ? 'Zuerst die URL der CoE-Umgebung eintragen.'
                    : !this._managedOptions
                      ? 'Spalten werden aus dem CoE-Inventar geladen …'
                      : this._managedDetected
                        ? 'Angezeigt werden Developer-Environments, die laut dieser Spalte nicht managed sind.'
                        : 'Keine Managed-Spalte erkannt – ohne Auswahl werden alle Developer-Environments angezeigt.'
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
