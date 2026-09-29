import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'StorageActivityAnalyserWebPartStrings';
import { StorageActivityAnalyser } from './components/StorageActivityAnalyser';
import { IStorageActivityAnalyserProps } from './components/IStorageActivityAnalyserProps';
import { ScanScope } from './models/IScanResult';
import { DEFAULT_THRESHOLD_MONTHS, THRESHOLD_OPTIONS, thresholdLabel } from './services/activity';

export interface IStorageActivityAnalyserWebPartProps {
  title: string;
  scope: ScanScope;
  thresholdMonths: number;
}

export default class StorageActivityAnalyserWebPart extends BaseClientSideWebPart<IStorageActivityAnalyserWebPartProps> {
  public render(): void {
    const threshold = Number(this.properties.thresholdMonths);
    const element: React.ReactElement<IStorageActivityAnalyserProps> = React.createElement(StorageActivityAnalyser, {
      title: this.properties.title,
      scope: this.properties.scope === 'currentWeb' ? 'currentWeb' : 'siteCollection',
      thresholdMonths: THRESHOLD_OPTIONS.indexOf(threshold) >= 0 ? threshold : DEFAULT_THRESHOLD_MONTHS,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneDropdown('scope', {
                  label: strings.ScopeFieldLabel,
                  options: [
                    { key: 'siteCollection', text: strings.ScopeSiteCollection },
                    { key: 'currentWeb', text: strings.ScopeCurrentWeb }
                  ]
                }),
                PropertyPaneDropdown('thresholdMonths', {
                  label: strings.ThresholdFieldLabel,
                  options: THRESHOLD_OPTIONS.map((m) => ({ key: m, text: thresholdLabel(m) }))
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
