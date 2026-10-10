declare interface IHelloWorldApplicationCustomizerStrings {
  Title: string;
}

declare module 'HelloWorldApplicationCustomizerStrings' {
  const strings: IHelloWorldApplicationCustomizerStrings;
  export = strings;
}
