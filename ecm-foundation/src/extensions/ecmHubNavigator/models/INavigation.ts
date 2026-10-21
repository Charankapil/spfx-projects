/** Values of the EcmNavType choice column in the navigation list. */
export type NavItemType = 'Hub' | 'Department site' | 'Shared service' | 'Hub resource' | 'Header link';

export interface INavLink {
  id: number;
  title: string;
  url: string;
  description: string;
  restricted: boolean;
  order: number;
}

export interface IHub extends INavLink {
  color: string;
  /** Department sites, or the enterprise services for the shared-services pseudo hub. */
  sites: INavLink[];
  resources: INavLink[];
  isSharedServices: boolean;
}

export interface INavigationModel {
  hubs: IHub[];
  sharedServices: INavLink[];
  headerLinks: INavLink[];
}

export interface INavigationLocation {
  hubId?: number;
  siteId?: number;
  /** The site is not in the menu but is associated with the hub in SharePoint. */
  associated?: boolean;
}
