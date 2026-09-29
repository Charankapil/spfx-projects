import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneToggle, PropertyPaneLabel } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WorkBoardsWebPartStrings';
import { App } from '../../app/App';
import { SPHttpClient } from '@microsoft/sp-http';
import { initializeIcons } from '@fluentui/react';
import { SpClient } from '../../services/SpClient';

export interface IWorkBoardsWebPartProps {
  fillPage: boolean;
}

/**
 * Work Boards: monday.com-style boards stored in SharePoint lists on this site.
 * All requests use SPFx's SPHttpClient as the signed-in user. No app registration,
 * client id or Microsoft Graph permissions.
 */
export default class WorkBoardsWebPart extends BaseClientSideWebPart<IWorkBoardsWebPartProps> {
  private sp!: SpClient;
  private themeVars: { [name: string]: string } = {};

  protected async onInit(): Promise<void> {
    await super.onInit();
    // SharePoint pages usually register these already; Teams tabs and some pages do not.
    initializeIcons(undefined, { disableWarnings: true });
    this.sp = new SpClient(this.context.spHttpClient, SPHttpClient.configurations.v1, this.context.pageContext.web.absoluteUrl);
  }

  public render(): void {
    const element = React.createElement(App, {
      sp: this.sp,
      siteTitle: this.context.pageContext.web.title,
      fullPage: this.properties.fillPage !== false && this.displayMode === DisplayMode.Read,
      themeVars: this.themeVars,
      version: this.context.manifest.version
    });
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    if (!theme || !theme.semanticColors) {
      return;
    }
    const s = theme.semanticColors;
    const p = theme.palette || {};
    const vars: { [name: string]: string } = {};
    const set = (name: string, value: string | undefined): void => {
      if (value) {
        vars[name] = value;
      }
    };
    set('--wb-primary', p.themePrimary);
    set('--wb-bg', s.bodyBackground);
    set('--wb-surface', s.bodyStandoutBackground);
    set('--wb-text', s.bodyText);
    set('--wb-subtle', s.bodySubtext);
    set('--wb-border', s.inputBorder);
    set('--wb-border-light', s.bodyDivider);
    set('--wb-hover', s.listItemBackgroundHovered);
    set('--wb-selected', s.listItemBackgroundChecked);
    set('--wb-primary-text', s.primaryButtonText);
    this.themeVars = vars;
    if (this.renderedOnce) {
      this.render();
    }
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
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneToggle('fillPage', { label: strings.FillPageLabel, onText: strings.On, offText: strings.Off, checked: this.properties.fillPage !== false }),
                PropertyPaneLabel('fillPageHint', { text: strings.FillPageHint })
              ]
            }
          ]
        }
      ]
    };
  }
}
