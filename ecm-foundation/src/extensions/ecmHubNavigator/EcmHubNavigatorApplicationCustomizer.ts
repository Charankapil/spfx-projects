import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import * as strings from 'EcmHubNavigatorStrings';

import { applyEcmCssVariables } from '../../common/theme/ecmTheme';
import { HubNavigator, IHubNavigatorProps } from './components/HubNavigator';
import { INavigationLocation, INavigationModel } from './models/INavigation';
import { NavigationService } from './services/NavigationService';

export interface IEcmHubNavigatorProperties {
  /** Site holding the navigation list: server-relative ("/sites/ecm-hub") or absolute. */
  configSiteUrl: string;
  /** Defaults to "ECM Navigation". */
  listTitle?: string;
  /** How long the menu is cached in the browser. Defaults to 30. */
  cacheMinutes?: number;
  /** Title shown in the brand area. Defaults to "Corporate ECM Hub". */
  homeTitle?: string;
  /** Text in the brand badge, e.g. "AT". Defaults to the home title's initials. */
  brandBadge?: string;
  /** Search page the box submits to; "{q}" is the query. Defaults to the tenant search page. */
  searchUrl?: string;
}

interface ILegacyPageContext {
  hubSiteId?: string;
  siteClassification?: string;
}

export default class EcmHubNavigatorApplicationCustomizer extends BaseApplicationCustomizer<IEcmHubNavigatorProperties> {
  private placeholder: PlaceholderContent | undefined;
  private service!: NavigationService;
  private model: INavigationModel | undefined;
  private error: string | undefined;
  private location: INavigationLocation = {};
  private hubSiteUrl: string | undefined;

  private get origin(): string {
    return /^(https?:\/\/[^/]+)/i.exec(this.context.pageContext.site.absoluteUrl)?.[1] || '';
  }

  private get configSiteUrl(): string {
    const configured = (this.properties.configSiteUrl || '').trim().replace(/\/+$/, '');
    if (!configured) {
      return this.context.pageContext.site.absoluteUrl;
    }
    return configured.indexOf('/') === 0 ? this.origin + configured : configured;
  }

  @override
  public async onInit(): Promise<void> {
    applyEcmCssVariables();

    this.service = new NavigationService(this.context.spHttpClient, {
      configSiteUrl: this.configSiteUrl,
      listTitle: this.properties.listTitle || 'ECM Navigation',
      cacheMinutes: typeof this.properties.cacheMinutes === 'number' ? this.properties.cacheMinutes : 30,
      sharedServicesTitle: strings.SharedServicesTitle,
      userKey: this.context.pageContext.user.loginName || 'anonymous'
    });

    // Paint straight from cache when we can, then refresh in the background.
    this.model = this.service.getCached();
    this.context.placeholderProvider.changedEvent.add(this, this.render);
    this.context.application.navigatedEvent.add(this, this.onNavigated);
    this.render();

    if (!this.model) {
      await this.refresh();
    } else {
      this.refresh().catch(() => undefined);
    }
  }

  private async refresh(): Promise<void> {
    try {
      this.model = await this.service.load();
      this.error = undefined;
    } catch (err) {
      // Keep showing a cached menu if we have one; only surface the error otherwise.
      this.error = this.model ? undefined : (err as Error).message;
    }
    await this.resolveLocation();
    this.render();
  }

  private onNavigated(): void {
    this.hubSiteUrl = undefined;
    this.resolveLocation()
      .then(() => this.render())
      .catch(() => this.render());
  }

  /**
   * Matches the current site against the list first; a site that is not
   * listed but is associated with a hub still gets that hub's breadcrumb.
   */
  private async resolveLocation(): Promise<void> {
    if (!this.model) {
      this.location = {};
      return;
    }
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    this.location = this.service.locate(this.model, siteUrl);
    const legacy = this.context.pageContext.legacyPageContext as ILegacyPageContext;
    const hubSiteId = legacy?.hubSiteId;
    if (this.location.hubId === undefined && hubSiteId && hubSiteId !== '00000000-0000-0000-0000-000000000000') {
      this.hubSiteUrl = this.hubSiteUrl || (await this.service.getHubSiteUrl(this.context.pageContext.web.absoluteUrl, hubSiteId));
      this.location = this.service.locate(this.model, siteUrl, this.hubSiteUrl);
    }
  }

  private render(): void {
    if (!this.placeholder) {
      this.placeholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: () => {
          if (this.placeholder) {
            ReactDom.unmountComponentAtNode(this.placeholder.domElement);
          }
        }
      });
    }
    if (!this.placeholder) {
      return;
    }

    const legacy = this.context.pageContext.legacyPageContext as ILegacyPageContext;
    const props: IHubNavigatorProps = {
      model: this.model,
      error: this.error,
      homeTitle: this.properties.homeTitle || strings.HomeTitle,
      homeUrl: this.configSiteUrl,
      brandBadge: this.properties.brandBadge,
      searchUrlTemplate: this.properties.searchUrl || `${this.origin}/_layouts/15/search.aspx/?q={q}`,
      location: this.location,
      currentSiteTitle: this.context.pageContext.web.title,
      classification: legacy?.siteClassification || undefined
    };
    ReactDom.render(React.createElement(HubNavigator, props), this.placeholder.domElement);
  }

  @override
  public onDispose(): void {
    if (this.placeholder) {
      ReactDom.unmountComponentAtNode(this.placeholder.domElement);
    }
  }
}
