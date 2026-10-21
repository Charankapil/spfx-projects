import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import * as strings from 'EcmHubNavigatorStrings';

import { applyEcmCssVariables, IEcmBrand, resolveBrand } from '../../common/theme/ecmTheme';
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
  /** Logo image (absolute or server-relative URL). ECM Setup uploads it to the hub's Site Assets. */
  logoUrl?: string;
  /** Brand colour for buttons, links and the accent line, e.g. "#C91F3C". */
  brandPrimary?: string;
  /** Dark brand colour for the hub band, e.g. "#1E2227". */
  brandDark?: string;
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
  /** Bumped on every in-site page navigation so open menus close. */
  private navigationKey = 0;

  private get origin(): string {
    return /^(https?:\/\/[^/]+)/i.exec(this.context.pageContext.site.absoluteUrl)?.[1] || '';
  }

  private get configSiteUrl(): string {
    const raw = (this.properties.configSiteUrl || '').trim();
    if (!raw) {
      return this.context.pageContext.site.absoluteUrl;
    }
    // "/" is the tenant root site, so trimming slashes must not turn it into "not configured".
    const configured = raw.replace(/\/+$/, '');
    if (!configured) {
      return this.origin;
    }
    return configured.indexOf('/') === 0 ? this.origin + configured : configured;
  }

  private get brand(): IEcmBrand {
    return resolveBrand({ primary: this.properties.brandPrimary, dark: this.properties.brandDark, logoUrl: this.properties.logoUrl });
  }

  @override
  public async onInit(): Promise<void> {
    applyEcmCssVariables(this.brand);

    this.service = new NavigationService(this.context.spHttpClient, {
      configSiteUrl: this.configSiteUrl,
      listTitle: this.properties.listTitle || 'ECM Navigation',
      cacheMinutes: typeof this.properties.cacheMinutes === 'number' ? this.properties.cacheMinutes : 30,
      sharedServicesTitle: strings.SharedServicesTitle,
      userKey: this.context.pageContext.user.loginName || 'anonymous'
    });

    // A fresh cache means no request at all for this page view; the menu is
    // reloaded once the cache is older than cacheMinutes.
    this.model = this.service.getCached();
    this.context.placeholderProvider.changedEvent.add(this, this.render);
    this.context.application.navigatedEvent.add(this, this.onNavigated);

    if (this.model) {
      await this.resolveLocation();
      this.render();
    } else if (this.service.isTemporarilyUnavailable()) {
      // Failed moments ago (e.g. ECM Setup has not run yet): stay hidden without retrying on every page.
      this.error = 'The navigation list was not available a few minutes ago.';
      this.render();
    } else {
      this.render();
      await this.refresh();
    }
  }

  private async refresh(): Promise<void> {
    try {
      this.model = await this.service.load();
      this.error = undefined;
    } catch (err) {
      // Keep showing a cached menu if we have one; otherwise the navigator
      // hides itself (see render), so a missing or unreadable list never
      // shows a broken band on every site.
      this.error = this.model ? undefined : (err as Error).message;
      if (!this.model) {
        this.service.markUnavailable();
      }
    }
    await this.resolveLocation();
    this.render();
  }

  private onNavigated(): void {
    this.hubSiteUrl = undefined;
    this.navigationKey++;
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
        onDispose: (content) => {
          ReactDom.unmountComponentAtNode(content.domElement);
          // SharePoint can recreate the placeholder; the next render must use the new one.
          this.placeholder = undefined;
        }
      });
      if (this.placeholder) {
        // Lets the mega menu overlay the page content below the header.
        this.placeholder.domElement.style.position = 'relative';
        this.placeholder.domElement.style.zIndex = '1000';
      }
    }
    if (!this.placeholder) {
      return;
    }
    if (this.error && !this.model) {
      console.warn(`[ECM Hub Navigator] Hidden: the navigation list could not be loaded from ${this.configSiteUrl}. ${this.error}`);
      ReactDom.unmountComponentAtNode(this.placeholder.domElement);
      return;
    }

    const legacy = this.context.pageContext.legacyPageContext as ILegacyPageContext;
    const props: IHubNavigatorProps = {
      model: this.model,
      error: this.error,
      brand: this.brand,
      homeTitle: this.properties.homeTitle || strings.HomeTitle,
      homeUrl: this.configSiteUrl,
      brandBadge: this.properties.brandBadge,
      searchUrlTemplate: this.properties.searchUrl || `${this.origin}/_layouts/15/search.aspx/?q={q}`,
      location: this.location,
      navigationKey: this.navigationKey,
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
