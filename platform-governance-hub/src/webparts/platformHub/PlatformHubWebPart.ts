import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneTextField, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { createTheme, getIcon, initializeIcons, ITheme } from '@fluentui/react';
import * as strings from 'PlatformHubWebPartStrings';

import { App } from './app/ui/App';
import { PageId } from './app/ui/AppContext';
import { PAGES } from './app/ui/sections';
import { IHttpRequestInit, IHttpResponseLike, IHttpTransport } from './app/http/ApiClient';
import { IDataProvider } from './app/providers/IDataProvider';
import { LiveDataProvider } from './app/providers/LiveDataProvider';
import { SampleDataProvider } from './app/providers/SampleDataProvider';

export interface IPlatformHubWebPartProps {
  title: string;
  dataSource: 'live' | 'sample';
  includeDataverse: boolean;
  maxRows: number;
  defaultPage: PageId;
}

/** Acquires tokens through SharePoint (AadHttpClient), one client per Entra ID resource. */
class SpfxTransport implements IHttpTransport {
  private readonly _factory: AadHttpClientFactory;
  private _clients: Record<string, Promise<AadHttpClient>> = {};

  constructor(factory: AadHttpClientFactory) {
    this._factory = factory;
  }

  public async fetch(resource: string, url: string, init: IHttpRequestInit): Promise<IHttpResponseLike> {
    let pending = this._clients[resource];
    if (!pending) {
      pending = this._clients[resource] = this._factory.getClient(resource);
      pending.catch(() => {
        delete this._clients[resource];
      });
    }
    const client = await pending;
    return client.fetch(url, AadHttpClient.configurations.v1, { method: init.method, headers: init.headers, body: init.body });
  }
}

export default class PlatformHubWebPart extends BaseClientSideWebPart<IPlatformHubWebPartProps> {
  private _provider: IDataProvider | undefined;
  private _providerKey: string = '';
  private _themeProvider: ThemeProvider | undefined;
  private _theme: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    // SharePoint pages normally register Fluent icons already; only load them if missing.
    if (!getIcon('ViewDashboard')) initializeIcons(undefined, { disableWarnings: true });
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._onThemeChanged);
  }

  private _onThemeChanged(args: ThemeChangedEventArgs): void {
    this._theme = args.theme;
    this.render();
  }

  private _getProvider(): IDataProvider {
    const mode = this.properties.dataSource === 'sample' ? 'sample' : 'live';
    const maxRows = Math.max(1000, Math.min(500000, Number(this.properties.maxRows) || 50000));
    const key = `${mode}|${this.properties.includeDataverse !== false}|${maxRows}`;
    if (!this._provider || key !== this._providerKey) {
      this._providerKey = key;
      if (mode === 'sample') {
        this._provider = new SampleDataProvider();
      } else {
        let host = '';
        try {
          host = new URL(this.context.pageContext.site.absoluteUrl).hostname.replace(/\.sharepoint\.com$/i, '');
        } catch {
          host = 'this tenant';
        }
        this._provider = new LiveDataProvider(new SpfxTransport(this.context.aadHttpClientFactory), {
          tenantLabel: host,
          includeDataverse: this.properties.includeDataverse !== false,
          maxRows
        });
      }
    }
    return this._provider;
  }

  private _fluentTheme(): { theme: ITheme; isDark: boolean } {
    const t = this._theme;
    const isDark = !!(t && t.isInverted);
    if (!t || !t.palette) return { theme: createTheme({}), isDark: false };
    const theme = createTheme({
      palette: t.palette as ITheme['palette'],
      semanticColors: t.semanticColors as ITheme['semanticColors'],
      isInverted: isDark
    });
    return { theme, isDark };
  }

  public render(): void {
    const { theme, isDark } = this._fluentTheme();
    const element = React.createElement(App, {
      provider: this._getProvider(),
      theme,
      isDark,
      title: this.properties.title || 'Platform Governance Hub',
      initialPage: this.properties.defaultPage || 'overview',
      // Re-mount when the data source changes so no stale state survives.
      key: this._providerKey
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    if (this._themeProvider) this._themeProvider.themeChangedEvent.remove(this, this._onThemeChanged);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    // Apply settings when the pane's Apply button is pressed, not on every keystroke.
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.SettingsGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                PropertyPaneDropdown('defaultPage', {
                  label: strings.DefaultPageLabel,
                  options: PAGES.map((p) => ({ key: p.id, text: p.title }))
                })
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneDropdown('dataSource', {
                  label: strings.DataSourceFieldLabel,
                  options: [
                    { key: 'live', text: strings.DataSourceLive },
                    { key: 'sample', text: strings.DataSourceSample }
                  ]
                }),
                PropertyPaneToggle('includeDataverse', { label: strings.IncludeDataverseLabel }),
                PropertyPaneTextField('maxRows', {
                  label: strings.MaxRowsLabel,
                  onGetErrorMessage: (v: string) => (/^\d+$/.test(String(v)) && Number(v) >= 1000 ? '' : 'Enter a whole number of at least 1000.')
                }),
                PropertyPaneLabel('permissionsNote', { text: strings.PermissionsNote })
              ]
            }
          ]
        }
      ]
    };
  }
}
