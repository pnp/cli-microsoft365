declare interface IHelloWorldStrings {
  Command1: string;
  Command2: string;
}

declare module 'helloWorldStrings' {
  const strings: IHelloWorldStrings;
  export = strings;
}
