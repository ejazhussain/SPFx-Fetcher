declare interface ISpFxFetchStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'spFxFetchStrings' {
  const strings: ISpFxFetchStrings;
  export = strings;
}
