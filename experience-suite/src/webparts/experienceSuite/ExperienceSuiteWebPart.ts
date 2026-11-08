import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ExperienceSuiteWebPartStrings';
import { ExperienceSuite } from './components/ExperienceSuite';

export interface IExperienceSuiteWebPartProps { title: string; enablePeopleSearch: boolean; }

export default class ExperienceSuiteWebPart extends BaseClientSideWebPart<IExperienceSuiteWebPartProps> {
  private _root: Root | undefined;

  public render(): void {
    const element = React.createElement(ExperienceSuite, {
      title: this.properties.title || 'SharePoint Experience Suite',
      context: this.context,
      enablePeopleSearch: this.properties.enablePeopleSearch !== false
    });
    // React 18: render through a root (ReactDOM.render runs React 18 in legacy React 17 mode and is removed in React 19).
    if (!this._root) this._root = createRoot(this.domElement);
    this._root.render(element);
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
            PropertyPaneToggle('enablePeopleSearch', { label: strings.PeopleSearchFieldLabel, onText: 'On', offText: 'Off' })
          ]
        }]
      }]
    };
  }
}
