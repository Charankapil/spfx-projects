declare interface IVantageWebPartStrings {
  AppTitle: string;
  AppSubtitle: string;
}

declare module 'VantageWebPartStrings' {
  const strings: IVantageWebPartStrings;
  export = strings;
}
