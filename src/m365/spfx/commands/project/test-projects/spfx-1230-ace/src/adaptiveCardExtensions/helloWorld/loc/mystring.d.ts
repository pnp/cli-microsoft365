declare interface IHelloWorldAdaptiveCardExtensionStrings {
  PropertyPaneDescription: string;
  TitleFieldLabel: string;
  Title: string;
  SubTitle: string;
  PrimaryText: string;
  Description: string;
  QuickViewButton: string;
}

declare module 'HelloWorldAdaptiveCardExtensionStrings' {
  const strings: IHelloWorldAdaptiveCardExtensionStrings;
  export = strings;
}
