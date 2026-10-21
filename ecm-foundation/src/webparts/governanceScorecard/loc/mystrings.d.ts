declare interface IGovernanceScorecardWebPartStrings {
  PropertyPaneDescription: string;
  DisplayGroup: string;
  MeasureGroup: string;
  SourcesGroup: string;
  TitleLabel: string;
  RequiredFieldsLabel: string;
  RequiredFieldsDescription: string;
  StaleYearsLabel: string;
  MaxDocsLabel: string;
  ConfigSiteLabel: string;
  ConfigSiteDescription: string;
  NavListLabel: string;
  SnapshotListLabel: string;
}

declare module 'GovernanceScorecardWebPartStrings' {
  const strings: IGovernanceScorecardWebPartStrings;
  export = strings;
}
