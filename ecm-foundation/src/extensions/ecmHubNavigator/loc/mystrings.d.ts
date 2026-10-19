declare interface IEcmHubNavigatorStrings {
  HomeTitle: string;
  HomeSubtitle: string;
  SharedServicesTitle: string;
  SearchPlaceholder: string;
  HubsNavLabel: string;
  DepartmentSites: string;
  EnterpriseServices: string;
  HubResources: string;
  SharedServicesColumn: string;
  Governance: string;
  GoToHub: string;
  YouAreHere: string;
  Restricted: string;
  TrimmedNotice: string;
  CloseMenu: string;
  OpenMenu: string;
  Loading: string;
  LoadFailed: string;
  NotConfigured: string;
}

declare module 'EcmHubNavigatorStrings' {
  const strings: IEcmHubNavigatorStrings;
  export = strings;
}
