declare interface IMyRequestWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MyRequestWebPartStrings' {
  const strings: IMyRequestWebPartStrings;
  export = strings;
}
