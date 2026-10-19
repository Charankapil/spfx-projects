declare interface IMetadataExplorerWebPartStrings {
  PropertyPaneDescription: string;
  GeneralGroup: string;
  SourceGroup: string;
  LibrariesGroup: string;
  SearchGroup: string;
  TitleLabel: string;
  SourceLabel: string;
  SourceLibraries: string;
  SourceSearch: string;
  LibraryTitlesLabel: string;
  LibraryTitlesDescription: string;
  ExcludedFieldsLabel: string;
  ScopeLabel: string;
  ScopeSite: string;
  ScopeHub: string;
  ScopeTenant: string;
  FieldMappingsLabel: string;
  FieldMappingsDescription: string;
  ExtraQueryLabel: string;
  DefaultGroupByLabel: string;
  DefaultThenByLabel: string;
  MaxItemsLabel: string;
}

declare module 'MetadataExplorerWebPartStrings' {
  const strings: IMetadataExplorerWebPartStrings;
  export = strings;
}
