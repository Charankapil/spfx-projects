define([], function () {
  return {
    PropertyPaneDescription: 'Shows the documents in this site grouped by their metadata, with links to every file.',
    GeneralGroup: 'Display',
    SourceGroup: 'Where documents come from',
    LibrariesGroup: 'Document libraries',
    SearchGroup: 'Search',
    TitleLabel: 'Title',
    SourceLabel: 'Source',
    SourceLibraries: 'Document libraries in this site (no setup needed)',
    SourceSearch: 'Search index (site, hub or whole organisation)',
    LibraryTitlesLabel: 'Libraries to include',
    LibraryTitlesDescription: 'One library title per line. Leave empty for every document library.',
    ExcludedFieldsLabel: 'Columns to hide (one per line)',
    ScopeLabel: 'Scope',
    ScopeSite: 'This site',
    ScopeHub: 'This hub (all associated sites)',
    ScopeTenant: 'Whole organisation',
    FieldMappingsLabel: 'Metadata categories',
    FieldMappingsDescription: 'One per line as Title|ManagedProperty, e.g. Document Type|RefinableString01. The property must be retrievable in the search schema.',
    ExtraQueryLabel: 'Extra KQL filter (optional)',
    DefaultGroupByLabel: 'Group by (column title)',
    DefaultThenByLabel: 'Then by (column title, optional)',
    MaxItemsLabel: 'Maximum documents'
  };
});
