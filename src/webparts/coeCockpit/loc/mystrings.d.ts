declare interface ICoeCockpitWebPartStrings {
  PropertyPaneDescription: string;
  ConnectionGroupName: string;
  EnvironmentUrlLabel: string;
  EnvironmentUrlDescription: string;
  EnvironmentUrlError: string;
  DisplayGroupName: string;
  TitleLabel: string;
  SubtitleLabel: string;
  ShowKpisLabel: string;
  DensityLabel: string;
  DensityComfortable: string;
  DensityCompact: string;
  PageSizeLabel: string;
  MaxRecordsLabel: string;
  TablesGroupName: string;
  OpenEditorLabel: string;
  TablesSummary: string;
  DefaultEnvGroupName: string;
  DefaultEnvShowLabel: string;
  DefaultEnvLabel: string;
  DefaultEnvAuto: string;
  DefaultEnvNeedsUrl: string;
  DefaultEnvLoading: string;
  DefaultEnvLoadError: string;
  DefaultEnvCount: string;
  DefaultEnvDataverseLabel: string;
  DefaultEnvDataverseDescription: string;
}

declare module 'CoeCockpitWebPartStrings' {
  const strings: ICoeCockpitWebPartStrings;
  export = strings;
}
