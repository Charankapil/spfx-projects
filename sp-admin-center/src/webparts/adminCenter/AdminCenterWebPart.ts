import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AdminCenterWebPartStrings';
import { AdminCenter, IAdminCenterProps } from './components/AdminCenter';
import { SPClient } from './core/SPClient';

export interface IAdminCenterWebPartProps {
  heading: string;
}

export default class AdminCenterWebPart extends BaseClientSideWebPart<IAdminCenterWebPartProps> {
  private client?: SPClient;

  protected onInit(): Promise<void> {
    // One throttle-aware client per web part instance, bound to this tenant only.
    const origin = (/^(https:\/\/[^/]+)/i.exec(this.context.pageContext.web.absoluteUrl) || [''])[1];
    this.client = new SPClient(this.context.spHttpClient, origin);
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IAdminCenterProps> = React.createElement(AdminCenter, {
      client: this.client as SPClient,
      homeWebUrl: this.context.pageContext.web.absoluteUrl,
      homeSiteUrl: this.context.pageContext.site.absoluteUrl,
      homeTitle: this.context.pageContext.web.title,
      heading: this.properties.heading || strings.DefaultHeading,
      currentUser: this.context.pageContext.user.email || this.context.pageContext.user.loginName || ''
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [PropertyPaneTextField('heading', { label: strings.HeadingFieldLabel })]
            }
          ]
        }
      ]
    };
  }
}
