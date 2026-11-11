import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, IPropertyPaneField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'SiteToolkitStrings';

export interface IToolkitComponentProps<P> {
  context: WebPartContext;
  properties: P;
}

/**
 * Shared plumbing for every Site Toolkit web part: render the React
 * component, clean it up, and build a one-group property pane.
 */
export abstract class ToolkitWebPart<P> extends BaseClientSideWebPart<P> {
  protected abstract readonly component: React.ComponentType<IToolkitComponentProps<P>>;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected get fields(): IPropertyPaneField<any>[] {
    return [];
  }

  public render(): void {
    ReactDom.render(
      React.createElement(this.component, { context: this.context, properties: this.properties }),
      this.domElement
    );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    // Settings apply on "Apply", so typing in the pane never fires a request per keystroke.
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [{ groupName: strings.BasicGroupName, groupFields: this.fields }]
        }
      ]
    };
  }
}
