import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ESignatureWebPartStrings';
import { ESignApp } from './components/ESignApp';
import { ESignService } from './services/ESignService';

export interface IESignatureWebPartProps {
  title: string;
  pageUrl: string;
}

export default class ESignatureWebPart extends BaseClientSideWebPart<IESignatureWebPartProps> {
  public render(): void {
    const pageUrl = (this.properties.pageUrl || '').trim() || `${window.location.origin}${window.location.pathname}`;
    const service = new ESignService(this.context, pageUrl);
    ReactDom.render(
      React.createElement(ESignApp, {
        title: this.properties.title || 'eSignature',
        service,
        canManage:
          this.context.pageContext.web.permissions.hasPermission(SPPermission.manageLists) ||
          this.context.pageContext.legacyPageContext?.isSiteAdmin === true
      }),
      this.domElement
    );
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                PropertyPaneTextField('pageUrl', {
                  label: strings.PageUrlFieldLabel,
                  description: strings.PageUrlFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
