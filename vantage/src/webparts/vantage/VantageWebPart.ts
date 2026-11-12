import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPPermission } from '@microsoft/sp-page-context';

import * as strings from 'VantageWebPartStrings';
import { VantageApp, type IVantageAppProps } from './components/VantageApp';
import { SharePointService } from './services/SharePointService';
import type { ICurrentUser } from './models/Models';

export default class VantageWebPart extends BaseClientSideWebPart<Record<string, never>> {
  private _root: Root | undefined;
  private _theme: IReadonlyTheme | undefined;
  private _sp: SharePointService | undefined;
  private _user: ICurrentUser | undefined;

  protected onInit(): Promise<void> {
    const ctx = this.context.pageContext;
    this._sp = new SharePointService(this.context.spHttpClient, ctx.web.absoluteUrl);
    this._user = {
      displayName: ctx.user.displayName,
      loginName: SharePointService.claims(ctx.user.loginName),
      email: ctx.user.email
    };
    return super.onInit();
  }

  public render(): void {
    const permissions = this.context.pageContext.web.permissions;
    const element = React.createElement<IVantageAppProps>(VantageApp, {
      sp: this._sp as SharePointService,
      user: this._user as ICurrentUser,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      siteOwner: permissions.hasPermission(SPPermission.managePermissions),
      canManageLists: permissions.hasAllPermissions(SPPermission.manageLists, SPPermission.managePermissions),
      siteCanEdit: permissions.hasAllPermissions(SPPermission.addListItems, SPPermission.editListItems),
      theme: this._theme,
      title: strings.AppTitle,
      subtitle: strings.AppSubtitle
    });

    if (!this._root) this._root = createRoot(this.domElement);
    this._root.render(element);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    this._theme = currentTheme;
    if (this._root) this.render();
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
