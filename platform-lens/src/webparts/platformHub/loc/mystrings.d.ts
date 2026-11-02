declare interface IPlatformHubWebPartStrings {
  PropertyPaneDescription: string;
  SettingsGroupName: string;
  DataGroupName: string;
  TitleFieldLabel: string;
  DataSourceFieldLabel: string;
  DataSourceLive: string;
  DataSourceSample: string;
  IncludeDataverseLabel: string;
  MaxRowsLabel: string;
  DefaultPageLabel: string;
  PermissionsNote: string;
}

declare module 'PlatformHubWebPartStrings' {
  const strings: IPlatformHubWebPartStrings;
  export = strings;
}
