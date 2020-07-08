declare interface IHelloWorldStrings {
  Title: string;
}

declare module 'helloWorldStrings' {
  const strings: IHelloWorldStrings;
  export = strings;
}
