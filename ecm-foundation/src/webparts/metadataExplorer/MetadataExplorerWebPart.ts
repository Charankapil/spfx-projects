import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MetadataExplorerWebPartStrings';

import { IMetadataExplorerProps, MetadataExplorer } from './components/MetadataExplorer';
import { IMetadataSource } from './services/IMetadataSource';
import { LibrariesSource } from './services/LibrariesSource';
import { SearchScope, SearchSource } from './services/SearchSource';

export interface IMetadataExplorerWebPartProps {
  title: string;
  source: 'libraries' | 'search';
  libraryTitles: string;
  excludedFields: string;
  searchScope: SearchScope;
  fieldMappings: string;
  extraQuery: string;
  defaultGroupBy: string;
  defaultThenBy: string;
  maxItems: number;
}

function lines(text: string): string[] {
  return (text || '')
    .split(/\r?\n|,/)
    .map((l) => l.trim())
    .filter(Boolean);
}

export default class MetadataExplorerWebPart extends BaseClientSideWebPart<IMetadataExplorerWebPartProps> {
  private source: IMetadataSource | undefined;

  /** One source per configuration, so re-rendering does not refetch every library. */
  private createSource(): IMetadataSource {
    const p = this.properties;
    const page = this.context.pageContext;
    const maxItems = Math.max(100, p.maxItems || 5000);
    if (p.source === 'search') {
      return new SearchSource(this.context.spHttpClient, {
        siteUrl: page.site.absoluteUrl,
        siteTitle: page.web.title,
        scope: p.searchScope || 'site',
        hubSiteId: (page.legacyPageContext as { hubSiteId?: string })?.hubSiteId,
        fieldMappings: p.fieldMappings,
        extraQuery: p.extraQuery,
        maxItems
      });
    }
    return new LibrariesSource(this.context.spHttpClient, {
      webUrl: page.web.absoluteUrl,
      webTitle: page.web.title,
      libraryTitles: lines(p.libraryTitles),
      excludedFields: lines(p.excludedFields),
      maxItems
    });
  }

  public render(): void {
    if (!this.source) {
      this.source = this.createSource();
    }
    const element: React.ReactElement<IMetadataExplorerProps> = React.createElement(MetadataExplorer, {
      title: this.properties.title,
      source: this.source,
      defaultGroupBy: this.properties.defaultGroupBy,
      defaultThenBy: this.properties.defaultThenBy
    });
    ReactDom.render(element, this.domElement);
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this.source = undefined;
    this.render();
  }

  /** Changes apply when the pane's Apply button is pressed, not on every keystroke. */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const search = this.properties.source === 'search';
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.GeneralGroup,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleLabel }),
                PropertyPaneTextField('defaultGroupBy', { label: strings.DefaultGroupByLabel }),
                PropertyPaneTextField('defaultThenBy', { label: strings.DefaultThenByLabel }),
                PropertyPaneSlider('maxItems', { label: strings.MaxItemsLabel, min: 500, max: 20000, step: 500 })
              ]
            },
            {
              groupName: strings.SourceGroup,
              groupFields: [
                PropertyPaneChoiceGroup('source', {
                  label: strings.SourceLabel,
                  options: [
                    { key: 'libraries', text: strings.SourceLibraries },
                    { key: 'search', text: strings.SourceSearch }
                  ]
                })
              ]
            },
            search
              ? {
                  groupName: strings.SearchGroup,
                  groupFields: [
                    PropertyPaneChoiceGroup('searchScope', {
                      label: strings.ScopeLabel,
                      options: [
                        { key: 'site', text: strings.ScopeSite },
                        { key: 'hub', text: strings.ScopeHub },
                        { key: 'tenant', text: strings.ScopeTenant }
                      ]
                    }),
                    PropertyPaneTextField('fieldMappings', {
                      label: strings.FieldMappingsLabel,
                      description: strings.FieldMappingsDescription,
                      multiline: true,
                      rows: 8
                    }),
                    PropertyPaneTextField('extraQuery', { label: strings.ExtraQueryLabel })
                  ]
                }
              : {
                  groupName: strings.LibrariesGroup,
                  groupFields: [
                    PropertyPaneTextField('libraryTitles', {
                      label: strings.LibraryTitlesLabel,
                      description: strings.LibraryTitlesDescription,
                      multiline: true,
                      rows: 4
                    }),
                    PropertyPaneTextField('excludedFields', {
                      label: strings.ExcludedFieldsLabel,
                      multiline: true,
                      rows: 4
                    })
                  ]
                }
          ]
        }
      ]
    };
  }
}
