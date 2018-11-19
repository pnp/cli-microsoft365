declare interface IHelloWorldCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'HelloWorldCommandSetStrings' {
  const strings: IHelloWorldCommandSetStrings;
  export = strings;
}
