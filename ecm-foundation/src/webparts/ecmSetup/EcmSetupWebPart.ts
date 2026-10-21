import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'EcmSetupWebPartStrings';

import { EcmSetup, IEcmSetupProps } from './components/EcmSetup';
import { SetupService } from './services/SetupService';

export interface IEcmSetupWebPartProps {
  listTitle: string;
}

export default class EcmSetupWebPart extends BaseClientSideWebPart<IEcmSetupWebPartProps> {
  /**
   * Kept across renders: a new service would make the component reload the
   * saved brand and overwrite what the person has typed (renders happen on
   * edit/view switches and theme changes, not only on settings changes).
   */
  private service: SetupService | undefined;

  public render(): void {
    const page = this.context.pageContext;
    if (!this.service) {
      this.service = new SetupService(this.context.spHttpClient, page.web.absoluteUrl, this.properties.listTitle || 'ECM Navigation');
    }
    const element: React.ReactElement<IEcmSetupProps> = React.createElement(EcmSetup, {
      service: this.service,
      tenantId: page.aadInfo?.tenantId?.toString() || '',
      siteTitle: page.web.title,
      isSiteOwner: !!page.web.permissions?.hasPermission(SPPermission.manageWeb)
    });
    ReactDom.render(element, this.domElement);
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this.service = undefined;
    this.render();
  }

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
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [{ groupName: '', groupFields: [PropertyPaneTextField('listTitle', { label: strings.ListTitleLabel })] }]
        }
      ]
    };
  }
}
