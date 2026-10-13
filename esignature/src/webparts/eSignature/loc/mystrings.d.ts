declare interface IESignatureWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  PageUrlFieldLabel: string;
  PageUrlFieldDescription: string;
}

declare module 'ESignatureWebPartStrings' {
  const strings: IESignatureWebPartStrings;
  export = strings;
}
