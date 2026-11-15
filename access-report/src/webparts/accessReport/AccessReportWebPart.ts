import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AccessReportWebPartStrings';
import { AccessReport, IAccessReportProps } from './components/AccessReport';
import { SPClient } from './core/SPClient';

export interface IAccessReportWebPartProps {
  heading: string;
}

export default class AccessReportWebPart extends BaseClientSideWebPart<IAccessReportWebPartProps> {
  private client?: SPClient;

  protected onInit(): Promise<void> {
    // One throttle-aware client per web part instance, bound to this tenant only. Everything
    // goes through SPHttpClient as the signed-in user: no Graph, no API permission requests.
    const origin = (/^(https:\/\/[^/]+)/i.exec(this.context.pageContext.web.absoluteUrl) || [''])[1];
    this.client = new SPClient(this.context.spHttpClient, origin);
    return Promise.resolve();
  }

  public render(): void {
    const aad = this.context.pageContext.aadInfo;
    const element: React.ReactElement<IAccessReportProps> = React.createElement(AccessReport, {
      client: this.client as SPClient,
      homeWebUrl: this.context.pageContext.web.absoluteUrl,
      heading: this.properties.heading || strings.DefaultHeading,
      tenantId: aad && aad.tenantId ? String(aad.tenantId) : '',
      version: this.manifest.version
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    if (this.client) {
      this.client.cancelPending();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [{ groupName: strings.BasicGroupName, groupFields: [PropertyPaneTextField('heading', { label: strings.HeadingFieldLabel })] }]
        }
      ]
    };
  }
}
